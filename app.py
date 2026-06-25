"""
AGM Generator — Flask Application (Web-Native)
Auth: Clerk Organizations (JWT via PyJWT + JWKS)
DB:   Supabase (PostgreSQL via supabase-py, service role key)
"""

import base64
import io
import json
import os
import re
import zipfile
from datetime import datetime, timedelta
from functools import wraps

import flask
from dotenv import load_dotenv
from flask import Flask, g, jsonify, render_template, request, send_file
from supabase import create_client, Client as SupabaseClient
import jwt as pyjwt
from jwt import PyJWKClient

from generator import AGMGenerator
from history import DirectorHistory

load_dotenv()

app = Flask(__name__)

# ── Supabase client (service role — never exposed to browser) ──────────────── #
_supabase: SupabaseClient = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SERVICE_KEY"],
)

# ── Clerk JWT verification ─────────────────────────────────────────────────── #
_jwks_client: PyJWKClient | None = None
_clerk_frontend_api: str = ""


def _get_clerk_frontend_api() -> str:
    global _clerk_frontend_api
    if not _clerk_frontend_api:
        pk = os.environ.get("CLERK_PUBLISHABLE_KEY", "")
        try:
            # pk format: pk_test_<base64url> or pk_live_<base64url>
            # base64url decodes to "{frontend_api_domain}$"
            b64 = pk.split("_", 2)[2]
            padding = 4 - len(b64) % 4
            if padding != 4:
                b64 += "=" * padding
            _clerk_frontend_api = base64.b64decode(b64).decode("utf-8").rstrip("$")
        except Exception:
            _clerk_frontend_api = os.environ.get("CLERK_FRONTEND_API", "")
    return _clerk_frontend_api


def _get_jwks_client() -> PyJWKClient:
    global _jwks_client
    if _jwks_client is None:
        frontend_api = _get_clerk_frontend_api()
        _jwks_client = PyJWKClient(
            f"https://{frontend_api}/.well-known/jwks.json",
            cache_keys=True,
        )
    return _jwks_client


def _decode_token(token: str) -> dict:
    client = _get_jwks_client()
    signing_key = client.get_signing_key_from_jwt(token)
    issuer = f"https://{_get_clerk_frontend_api()}"
    return pyjwt.decode(
        token,
        signing_key.key,
        algorithms=["RS256"],
        options={"verify_aud": False},
        issuer=issuer,
        leeway=timedelta(seconds=10),
    )


def _require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization", "").removeprefix("Bearer ").strip()
        if not token:
            return jsonify({"error": "Unauthorized"}), 401
        try:
            payload = _decode_token(token)
            g.clerk_payload = payload
        except Exception as exc:
            return jsonify({"error": "Unauthorized", "detail": str(exc)}), 401
        return f(*args, **kwargs)
    return decorated


def _user_id() -> str:
    return g.clerk_payload.get("sub", "unknown")


def _org_id() -> str:
    # Prefer org_id from token; fall back to user sub for solo/dev use
    return g.clerk_payload.get("org_id") or g.clerk_payload.get("sub", "unknown")


# ── Routes ─────────────────────────────────────────────────────────────────── #

@app.route("/")
def index():
    return render_template(
        "index.html",
        clerk_publishable_key=os.environ.get("CLERK_PUBLISHABLE_KEY", ""),
    )


@app.route("/api/settings", methods=["GET", "POST"])
@_require_auth
def manage_settings():
    uid = _user_id()
    if request.method == "GET":
        result = _supabase.table("user_settings") \
            .select("fy_year") \
            .eq("user_id", uid) \
            .maybe_single() \
            .execute()
        row = result.data if result is not None else None
        return jsonify({
            "agm_financial_year": row["fy_year"] if row else str(datetime.now().year - 1),
        })
    else:
        payload = request.json or {}
        fy_year = payload.get("agm_financial_year", "")
        _supabase.table("user_settings") \
            .upsert({"user_id": uid, "fy_year": fy_year, "updated_at": datetime.utcnow().isoformat()}, on_conflict="user_id") \
            .execute()
        return jsonify({"success": True})


@app.route("/api/mappings", methods=["GET", "POST"])
@_require_auth
def field_mappings():
    uid = _user_id()

    if request.method == "GET":
        filename = request.args.get("filename", "").strip()
        if not filename:
            return jsonify({"error": "filename is required"}), 400
        result = _supabase.table("field_mappings") \
            .select("field, cell") \
            .eq("user_id", uid) \
            .eq("filename", filename) \
            .execute()
        mappings = {r["field"]: {"cell": r["cell"]} for r in (result.data or [])}
        return jsonify({"mappings": mappings})

    else:
        payload = request.json or {}
        filename = payload.get("filename", "").strip()
        mappings = payload.get("mappings", {})
        if not filename:
            return jsonify({"error": "filename is required"}), 400

        now = datetime.utcnow().isoformat()
        rows = [
            {"user_id": uid, "filename": filename, "field": field, "cell": cell, "updated_at": now}
            for field, cell in mappings.items()
        ]
        if rows:
            _supabase.table("field_mappings") \
                .upsert(rows, on_conflict="user_id,filename,field") \
                .execute()
        return jsonify({"success": True})


@app.route("/api/history", methods=["GET"])
@_require_auth
def get_director_history():
    uid = _user_id()
    reg_no = request.args.get("reg_no", "").strip()
    fy_year = request.args.get("fy_year", "").strip()
    if not reg_no:
        return jsonify({"error": "Registration number is required."}), 400

    try:
        current_year = int(fy_year) if fy_year else datetime.now().year - 1
    except ValueError:
        current_year = datetime.now().year - 1

    try:
        history = DirectorHistory(_supabase, uid)
        summary = history.rotation_summary(reg_no, current_year)
        return jsonify({"history": summary})
    except Exception as e:
        return jsonify({"error": f"Error loading history: {e}"}), 500


@app.route("/api/generate", methods=["POST"])
@_require_auth
def generate_resolutions():
    uid = _user_id()
    payload = request.json or {}
    companies = payload.get("companies", [])
    global_fy_year = payload.get("fy_year", "").strip()

    if not companies:
        return jsonify({"error": "No companies provided."}), 400

    generator = AGMGenerator()
    history = DirectorHistory(_supabase, uid)

    zip_buffer = io.BytesIO()
    results = []
    success_count = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in companies:
            data = item.get("data", {})

            fye = data.get("financial_year_end", "").strip()
            match = re.search(r'\b(19|20)\d{2}\b', fye)
            fy_year = match.group(0) if match else (global_fy_year or str(datetime.now().year))

            display_name = data.get("company_name", "").strip() or item.get("filename", "unknown")
            safe_name = "".join(c for c in display_name if c.isalnum() or c in " _-").strip()[:60]
            docx_filename = f"{safe_name}_AGM_{fy_year}.docx"

            ok, doc_bytes, msg = generator.generate_to_bytes(data)
            results.append({
                "filename": item.get("filename", ""),
                "company_name": display_name,
                "success": ok,
                "message": msg or f"Generated {docx_filename}",
            })

            if ok:
                success_count += 1
                zf.writestr(docx_filename, doc_bytes)
                reg_no = data.get("reg_no", "").strip()
                directors = data.get("directors", [])
                if reg_no and fy_year:
                    try:
                        history.record(reg_no, int(fy_year), directors)
                    except Exception:
                        pass

    zip_buffer.seek(0)

    if success_count == 0:
        return jsonify({
            "success": False,
            "results": results,
            "success_count": 0,
            "total_count": len(companies),
        }), 500

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"AGM_Documents_{timestamp}.zip"

    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=zip_name,
    )


# ── Sync endpoints ─────────────────────────────────────────────────────────── #

_COMPARE_FIELDS = [
    "company_name", "reg_no", "address",
    "financial_year_end", "agm_number", "agm_date",
]


@app.route("/api/sync", methods=["POST"])
@_require_auth
def sync_preview():
    """Phase 1: compute diffs vs stored snapshots. No writes."""
    org = _org_id()
    payload = request.json or {}
    companies = payload.get("companies", [])
    if not companies:
        return jsonify({"diffs": [], "total": 0})

    filenames = [c["filename"] for c in companies]
    existing_result = _supabase.table("company_snapshots") \
        .select("filename, company_name, reg_no, address, financial_year_end, agm_number, agm_date, directors") \
        .eq("org_id", org) \
        .in_("filename", filenames) \
        .execute()
    existing = {r["filename"]: r for r in (existing_result.data or [])}

    diffs = []
    for company in companies:
        fn = company["filename"]
        old = existing.get(fn)
        if not old:
            continue  # New company — no diff to show

        changes = []
        for field in _COMPARE_FIELDS:
            old_val = (old.get(field) or "").strip()
            new_val = (company.get(field) or "").strip()
            if old_val != new_val:
                changes.append({"field": field, "old": old_val, "new": new_val})

        old_dirs = sorted(d.strip() for d in (old.get("directors") or []) if d.strip())
        new_dirs = sorted(d.strip() for d in (company.get("all_directors") or []) if d.strip())
        if old_dirs != new_dirs:
            changes.append({
                "field": "directors",
                "old": ", ".join(old_dirs),
                "new": ", ".join(new_dirs),
            })

        if changes:
            diffs.append({
                "filename": fn,
                "company_name": company.get("company_name") or fn,
                "changes": changes,
            })

    return jsonify({"diffs": diffs, "total": len(companies)})


@app.route("/api/sync/confirm", methods=["POST"])
@_require_auth
def sync_confirm():
    """Phase 2: upsert all company snapshots."""
    org = _org_id()
    uid = _user_id()
    payload = request.json or {}
    companies = payload.get("companies", [])
    user_name = g.clerk_payload.get("email", "") or g.clerk_payload.get("sub", "")

    rows = []
    for c in companies:
        rows.append({
            "org_id": org,
            "filename": c["filename"],
            "reg_no": c.get("reg_no", ""),
            "company_name": c.get("company_name", ""),
            "address": c.get("address", ""),
            "financial_year_end": c.get("financial_year_end", ""),
            "agm_number": c.get("agm_number", ""),
            "agm_date": c.get("agm_date", ""),
            "directors": c.get("all_directors") or c.get("directors") or [],
            "selected_directors": c.get("selected_directors") or [],
            "last_scanned_at": datetime.utcnow().isoformat(),
            "last_scanned_by_name": user_name,
            "last_scanned_by_id": uid,
        })

    if rows:
        _supabase.table("company_snapshots") \
            .upsert(rows, on_conflict="org_id,filename") \
            .execute()

    return jsonify({"synced": len(rows)})


@app.route("/api/companies", methods=["GET"])
@_require_auth
def get_companies():
    """Mobile read: all company snapshots for the active org."""
    org = _org_id()
    result = _supabase.table("company_snapshots") \
        .select(
            "filename, company_name, reg_no, address, financial_year_end, "
            "agm_number, agm_date, directors, last_scanned_at, last_scanned_by_name"
        ) \
        .eq("org_id", org) \
        .order("company_name") \
        .execute()
    return jsonify({"companies": result.data or []})


if __name__ == "__main__":
    app.run(debug=True)

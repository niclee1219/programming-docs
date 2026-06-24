"""
AGM Generator — Flask Application (Web-Native)
Serves the UI and handles server-side concerns:
  - Per-user settings and field mapping persistence (SQLAlchemy)
  - Director history storage and rotation queries
  - In-memory DOCX generation with zip download response
File reading and .xlsm extraction happen entirely in the browser via SheetJS.
"""

import io
import json
import os
import re
import zipfile
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file, session

from models import db, UserSettings, FieldMapping, DirectorHistoryEntry
from generator import AGMGenerator
from history import DirectorHistory

load_dotenv()

app = Flask(__name__, instance_path="/tmp")
_db_url = os.environ.get("DATABASE_URL", "sqlite:////tmp/agm.db")
# Vercel Postgres URLs use postgres:// — SQLAlchemy needs postgresql://
if _db_url.startswith("postgres://"):
    _db_url = _db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = _db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

db.init_app(app)

with app.app_context():
    db.create_all()


def _require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("authenticated"):
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return decorated


def _user_id() -> str:
    # Stub: returns 'default' until Clerk JWT verification is wired in.
    # Replace with: return verify_clerk_token(request.headers.get('Authorization'))
    return request.headers.get("X-User-Id", "default")


# ------------------------------------------------------------------ #
# Routes
# ------------------------------------------------------------------ #

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/auth", methods=["POST"])
def login():
    login_code = os.environ.get("LOGIN_CODE", "")
    if not login_code:
        return jsonify({"error": "LOGIN_CODE not configured on server"}), 500
    submitted = (request.json or {}).get("code", "").strip()
    if submitted == login_code:
        session.permanent = True
        session["authenticated"] = True
        return jsonify({"success": True})
    return jsonify({"error": "Invalid code"}), 401


@app.route("/api/auth/check", methods=["GET"])
def auth_check():
    return jsonify({"authenticated": bool(session.get("authenticated"))})


@app.route("/api/auth/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/api/settings", methods=["GET", "POST"])
@_require_auth
def manage_settings():
    uid = _user_id()
    if request.method == "GET":
        row = UserSettings.query.filter_by(user_id=uid).first()
        return jsonify({
            "agm_financial_year": row.fy_year if row else str(datetime.now().year - 1),
        })
    else:
        payload = request.json or {}
        row = UserSettings.query.filter_by(user_id=uid).first()
        if not row:
            row = UserSettings(user_id=uid)
            db.session.add(row)
        row.fy_year = payload.get("agm_financial_year", row.fy_year)
        db.session.commit()
        return jsonify({"success": True})


@app.route("/api/mappings", methods=["GET", "POST"])
@_require_auth
def field_mappings():
    uid = _user_id()

    if request.method == "GET":
        filename = request.args.get("filename", "").strip()
        if not filename:
            return jsonify({"error": "filename is required"}), 400
        rows = FieldMapping.query.filter_by(user_id=uid, filename=filename).all()
        mappings = {r.field: {"cell": r.cell} for r in rows}
        return jsonify({"mappings": mappings})

    else:
        payload = request.json or {}
        filename = payload.get("filename", "").strip()
        mappings = payload.get("mappings", {})   # {field: cell_ref}
        if not filename:
            return jsonify({"error": "filename is required"}), 400

        for field, cell in mappings.items():
            row = FieldMapping.query.filter_by(
                user_id=uid, filename=filename, field=field
            ).first()
            if row:
                row.cell = cell
            else:
                db.session.add(FieldMapping(
                    user_id=uid, filename=filename, field=field, cell=cell
                ))
        db.session.commit()
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
        history = DirectorHistory(db, uid)
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
    history = DirectorHistory(db, uid)

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


if __name__ == "__main__":
    app.run(debug=True)

# AGM Web — Fixes & Security Hardening

**Date:** 2026-06-25  
**Status:** Approved

## Overview

Seven targeted fixes covering two UI bugs, one UX improvement, and four security / data-quality hardening changes. No architectural shifts — all changes are additive or in-place replacements.

---

## 1. Toast X Button (CSS)

**Problem:** `.toast-close` has no explicit styles; inherits browser default white button background.

**Fix:** Add to `static/style.css`:
```css
.toast-close {
    background: transparent;
    border: none;
    color: var(--text-muted);
    cursor: pointer;
    font-size: 1.1rem;
    padding: 0 2px;
    line-height: 1;
    flex-shrink: 0;
}
.toast-close:hover {
    color: var(--text-primary);
}
```

**Files:** `static/style.css`

---

## 2. Manage Account — In-App Popup

**Problem:** `clerk.openUserProfile()` throws "Clerk was not loaded with Ui components". Fallback `redirectToUserProfile()` navigates away from the app.

**Fix:** Mount the Clerk `<UserProfile>` component into a custom overlay:

- **HTML:** Add to `index.html` (before closing `</body>`):
  ```html
  <div id="clerk-profile-overlay" style="display:none;">
    <div id="clerk-profile-backdrop"></div>
    <div id="clerk-profile-container"></div>
  </div>
  ```
- **JS:** Replace `_openUserProfile()` to call `window.__clerk.mountUserProfile(containerEl)` on open and `window.__clerk.unmountUserProfile(containerEl)` on close (X button or backdrop click). Works for both desktop and mobile avatar dropdowns.
- **CSS:** Full-screen fixed overlay with backdrop blur + centered white/dark card matching app theme.

**Files:** `templates/index.html`, `static/app.js`, `static/style.css`

---

## 3. RLS — Implicit Deny

**Problem:** All four tables have `DISABLE ROW LEVEL SECURITY`. Direct database access (psql, Supabase dashboard with anon/authenticated role) has unrestricted read/write.

**Fix:** Enable RLS on all tables with zero permissive policies → Postgres implicit deny for all non-service-role connections. The Flask server's service role key bypasses RLS — app behaviour unchanged.

**Migration `migrations/002_enable_rls.sql`:**
```sql
ALTER TABLE user_settings       ENABLE ROW LEVEL SECURITY;
ALTER TABLE field_mappings      ENABLE ROW LEVEL SECURITY;
ALTER TABLE director_history    ENABLE ROW LEVEL SECURITY;
ALTER TABLE company_snapshots   ENABLE ROW LEVEL SECURITY;
```

**Files:** `migrations/002_enable_rls.sql` (new)

---

## 4. Name Capture — Email from JWT

**Problem:** `last_scanned_by_name` in `company_snapshots` is stored from a client-supplied `user_name` field which is blank when the Clerk user has no `fullName` set. Result: Supabase shows `user_3FcO...` (ID) with an empty name column.

**Fix:** In `sync_confirm` in `app.py`, read email server-side from the verified JWT payload:
```python
user_name = g.clerk_payload.get("email", "") or g.clerk_payload.get("sub", "")
```
Remove reliance on client-supplied `user_name`. The frontend no longer needs to include it in the payload (harmless if it still does).

**Files:** `app.py`

---

## 5. JWT Issuer Verification

**Problem:** `pyjwt.decode()` is called with `options={"verify_aud": False}` — a JWT from any Clerk instance with a valid signature would be accepted.

**Fix:** Pass `issuer=f"https://{frontend_api}"` to `pyjwt.decode()`. The `frontend_api` value is already derived from the publishable key in `_get_jwks_client()` — extract it to a module-level variable so `_decode_token` can reference it. Clerk does not set `aud` by default so audience verification stays off; `iss` is the reliable cross-instance discriminator.

**Files:** `app.py`

---

## 6. `updated_at` Audit Columns

**Problem:** `field_mappings`, `user_settings`, and `director_history` have no timestamp — impossible to tell when a row was last written.

**Fix:**

**Migration `migrations/003_add_updated_at.sql`:**
```sql
ALTER TABLE user_settings    ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE field_mappings   ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE director_history ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
```

Server sets `updated_at` explicitly on every upsert (not relying on DB default) so it reflects the actual write time even when the row already exists.

**Files:** `migrations/003_add_updated_at.sql` (new), `app.py`

---

## 7. Mobile Sync Banner — Flex Truncation

**Problem:** `.mobile-sync-banner` has `max-width: 220px` which clips on narrow phones now that the avatar button occupies the right edge of the mobile header.

**Fix:** Replace fixed max-width with flex grow + text truncation:
```css
.mobile-sync-banner {
    flex: 1;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
```

**Files:** `static/style.css`

---

## Change Surface Summary

| File | Changes |
|------|---------|
| `static/style.css` | Toast X styles (#1), clerk profile overlay styles (#2), mobile sync banner (#7) |
| `templates/index.html` | Clerk profile overlay div (#2) |
| `static/app.js` | `_openUserProfile` mount/unmount (#2) |
| `app.py` | Email from JWT (#4), issuer verification (#5), `updated_at` on upserts (#6) |
| `migrations/002_enable_rls.sql` | Enable RLS (#3) |
| `migrations/003_add_updated_at.sql` | Add `updated_at` columns (#6) |

## Out of Scope

- Rate limiting (deferred — internal team, Clerk-gated)
- Supabase audience (`aud`) claim verification (Clerk doesn't set it by default)
- Switching to anon key + per-request JWT RLS (Option B, deferred)

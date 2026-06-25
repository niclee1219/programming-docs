# Design: Mobile People Tab, Sync Banner, Avatar Profile Button

**Date:** 2026-06-25

---

## Scope

Four changes to the AGM web app:

1. People tab on mobile view
2. Sync banner — correct name + add HH:mm time
3. Avatar profile button replacing "Manage Team" + "Sign Out"
4. Remove all org-gating (already partially done)

---

## 1. Mobile People Tab

### Data source
`/api/companies` already returns a `directors` array per company snapshot. No backend changes required.

### UI
- Add a Companies / People tab toggle above the mobile company list, matching the desktop sidebar toggle.
- Switching to People builds an inverted index client-side: director name → list of companies they appear in. Built from `_mobileCompaniesCache` (already loaded).
- Each director renders as an accordion row: tap header to expand/collapse a list of their companies. Company list is display-only (no tap-through navigation).
- The search input filters by director name.
- The index rebuilds whenever `_mobileCompaniesCache` is refreshed (i.e. after `loadMobileView` completes).

### State
- Add `_mobilePeopleCache` (object: `{ [directorName]: string[] }`) alongside existing `_mobileCompaniesCache`.
- Add `_mobileViewMode` (`'companies' | 'people'`) to track active tab.

### Empty state
If no companies have been synced yet, show a placeholder consistent with the existing empty states.

---

## 2. Sync Banner — Name + Time

### Name fix
In both sync confirm payloads in `app.js` (lines ~936 and ~997), `user_name` is set to `window.__clerk.user.fullName`. Keep this but ensure it sends `''` rather than `null` when `fullName` is not set. Users who appear as blank should set their name via Manage Account.

No fallback to email — if fullName is empty, the banner shows "Last synced by  on …" which is acceptable and prompts the user to set their name.

### Time format
Change `updateMobileSyncBanner` to format `last_scanned_at` as:

```
Last synced by {name} on {date} at {HH:mm}
```

Example: `Last synced by Nicholas on 25 Jun 2026 at 14:32`

Use the existing `en-SG` locale for the date part. Extract hours/minutes from the same `Date` object, zero-padded.

---

## 3. Avatar Profile Button

### Remove
- `#btn-manage-team` button from toolbar HTML
- `#btn-logout` button from toolbar HTML
- Their event listeners in `registerEvents()`

### Add
A single avatar button in the toolbar's `scan-column` div:

**Button appearance**
- If `window.__clerk.user.imageUrl` is set: show a circular profile image.
- Otherwise: show a circle with the user's initials. Derive initials from `fullName` (first letter of first word + first letter of last word). Fall back to first letter of `fullName` if only one word. Fall back to "?" if fullName is empty.
- Styled to match existing toolbar icon buttons.

**Dropdown**
- Toggled by clicking the avatar button.
- Dismissed by clicking anywhere outside (document click listener).
- Two items:
  - "Manage Account" → `window.__clerk.openUserProfile()`
  - "Sign Out" → `window.__clerk.signOut().then(() => location.reload())`
- Positioned below the avatar button, right-aligned.
- Styled consistently with the existing modal/card patterns in `style.css`.

### No orgs
`window.__clerk.openOrganizationProfile()` is removed entirely. `_org_id()` on the backend already falls back to `user.sub`. No further org changes needed.

---

## Files changed

| File | Change |
|------|--------|
| `static/app.js` | Mobile people tab logic, rebuilt sync banner, avatar button + dropdown |
| `templates/index.html` | Replace two toolbar buttons with avatar button HTML |
| `static/style.css` | Avatar button + dropdown styles |

No backend changes required.

---

## Out of scope

- Sign-out on mobile: the toolbar (and avatar button) is desktop-only. The mobile header has no sign-out UI currently. This is not addressed in this spec.

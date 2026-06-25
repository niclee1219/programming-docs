# Mobile People Tab, Sync Banner & Avatar Profile Button — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a People tab to mobile view, fix the sync banner name/time display, and replace the Manage Team + Sign Out toolbar buttons with a single avatar profile dropdown.

**Architecture:** All changes are client-side only — no backend changes required. The mobile People index is built client-side from the `directors` array already returned by `/api/companies`. The avatar dropdown is a plain HTML/CSS/JS component that calls Clerk's `openUserProfile()` and `signOut()`.

**Tech Stack:** Vanilla JS (ES2020), Flask/Jinja2 HTML templates, plain CSS custom properties

## Global Constraints

- No new npm dependencies or CDN scripts
- Follow existing CSS custom property naming (`var(--bg-primary)`, `var(--accent-primary)`, etc.)
- All Clerk calls via `window.__clerk`
- No TypeScript — plain `.js` only
- Flask Jinja2 template syntax in `templates/index.html`

---

### Task 1: Fix sync banner — user name and HH:mm time

**Files:**
- Modify: `static/app.js` (two `user_name` payloads ~line 936 and ~997, plus `updateMobileSyncBanner` ~line 1039)

**Interfaces:**
- Produces: `updateMobileSyncBanner(companies)` displays `"Last synced by {fullName} on {date} at {HH:mm}"`

- [ ] **Step 1: Fix the two `user_name` payload lines**

Find and replace both occurrences of:
```js
user_name: window.__clerk.user ? window.__clerk.user.fullName : 'Unknown',
```
with:
```js
user_name: window.__clerk.user ? (window.__clerk.user.fullName || '') : '',
```
There are exactly two — one around line 936 (auto-sync path) and one around line 997 (manual confirm button).

- [ ] **Step 2: Update `updateMobileSyncBanner` to include HH:mm**

Replace the entire `updateMobileSyncBanner` function:
```js
function updateMobileSyncBanner(companies) {
    const banner = document.getElementById('mobile-sync-banner');
    if (!banner || !companies.length) return;

    const latest = companies.reduce((a, b) =>
        new Date(a.last_scanned_at) > new Date(b.last_scanned_at) ? a : b
    );

    if (!latest.last_scanned_at) {
        banner.textContent = 'No sync data yet';
        return;
    }

    const d = new Date(latest.last_scanned_at);
    const date = d.toLocaleDateString('en-SG', {
        day: 'numeric', month: 'short', year: 'numeric',
    });
    const hh = String(d.getHours()).padStart(2, '0');
    const mm = String(d.getMinutes()).padStart(2, '0');
    const name = latest.last_scanned_by_name || '';
    banner.textContent = name
        ? `Last synced by ${name} on ${date} at ${hh}:${mm}`
        : `Last synced on ${date} at ${hh}:${mm}`;
}
```

- [ ] **Step 3: Verify manually**

Run `flask run --port 5001`, open the app on desktop, scan a folder, trigger a sync. Check the mobile banner shows e.g. `"Last synced by Nicholas on 25 Jun 2026 at 14:32"` with no "Unknown".

- [ ] **Step 4: Commit**

```bash
git add static/app.js
git commit -m "fix: sync banner shows fullName and HH:mm timestamp"
```

---

### Task 2: Avatar profile button

**Files:**
- Modify: `templates/index.html` (replace two toolbar buttons with avatar button + dropdown)
- Modify: `static/app.js` (replace two event listeners with avatar/dropdown logic)
- Modify: `static/style.css` (add avatar + dropdown styles)

**Interfaces:**
- Produces: `#btn-avatar` button in toolbar; `#avatar-dropdown` div with "Manage Account" and "Sign Out"

- [ ] **Step 1: Replace toolbar HTML**

In `templates/index.html`, find and replace the "Team & account controls" `div`:
```html
<!-- Team & account controls -->
<div class="scan-column" style="gap:6px;">
  <button id="btn-manage-team" class="btn-theme-subtle" title="Manage team members">
    <i data-lucide="users"></i>
    <span class="theme-label">Manage Team</span>
  </button>
  <button id="btn-logout" class="btn-theme-subtle" title="Sign out">
    <i data-lucide="log-out"></i>
    <span class="theme-label">Sign Out</span>
  </button>
</div>
```
with:
```html
<!-- Avatar profile button -->
<div class="avatar-wrapper" id="avatar-wrapper">
  <button id="btn-avatar" class="btn-avatar" title="Account"></button>
  <div id="avatar-dropdown" class="avatar-dropdown" style="display:none;">
    <button id="btn-manage-account" class="avatar-dropdown-item">
      <i data-lucide="user"></i> Manage Account
    </button>
    <button id="btn-signout" class="avatar-dropdown-item avatar-dropdown-item--danger">
      <i data-lucide="log-out"></i> Sign Out
    </button>
  </div>
</div>
```

- [ ] **Step 2: Add avatar + dropdown CSS**

Append to the end of `static/style.css` (before the final closing line if any, otherwise just append):
```css
/* ── Avatar profile button ───────────────────────────────────────────── */
.avatar-wrapper {
    position: relative;
    display: flex;
    align-items: center;
}

.btn-avatar {
    width: 32px;
    height: 32px;
    border-radius: 50%;
    border: 2px solid var(--border-color);
    background: var(--accent-primary);
    color: #fff;
    font-size: 0.75rem;
    font-weight: 700;
    font-family: inherit;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 0;
    overflow: hidden;
    transition: border-color 0.2s, opacity 0.2s;
    line-height: 1;
}

.btn-avatar:hover {
    border-color: var(--accent-primary);
    opacity: 0.85;
}

.btn-avatar img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    border-radius: 50%;
}

.avatar-dropdown {
    position: absolute;
    top: calc(100% + 8px);
    right: 0;
    min-width: 180px;
    background: var(--bg-secondary);
    border: 1px solid var(--border-color);
    border-radius: 10px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.35);
    z-index: 500;
    overflow: hidden;
    display: flex;
    flex-direction: column;
}

body.light-theme .avatar-dropdown {
    box-shadow: 0 8px 24px rgba(0,0,0,0.12);
}

.avatar-dropdown-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 11px 16px;
    font-size: 0.82rem;
    font-weight: 500;
    font-family: inherit;
    color: var(--text-primary);
    background: transparent;
    border: none;
    cursor: pointer;
    text-align: left;
    transition: background 0.15s;
}

.avatar-dropdown-item i,
.avatar-dropdown-item svg {
    width: 14px;
    height: 14px;
    flex-shrink: 0;
    color: var(--text-muted);
}

.avatar-dropdown-item:hover {
    background: var(--bg-tertiary);
}

.avatar-dropdown-item--danger {
    color: var(--color-error, #e05252);
    border-top: 1px solid var(--border-color);
}

.avatar-dropdown-item--danger i,
.avatar-dropdown-item--danger svg {
    color: var(--color-error, #e05252);
}
```

- [ ] **Step 3: Replace event listeners in `app.js`**

Find the block in `registerEvents()` that handles `btnManageTeam` and `btnLogout` (~lines 485–502):
```js
if (el.btnManageTeam) {
    el.btnManageTeam.addEventListener('click', () => {
        if (!window.__clerk || !window.__clerk.organization) {
            showToast('No active organization. Please sign out and sign in again.', 'warn');
            return;
        }
        window.__clerk.openOrganizationProfile();
    });
}

if (el.btnLogout) {
    el.btnLogout.addEventListener('click', async () => {
        if (window.__clerk) {
            await window.__clerk.signOut();
        }
        location.reload();
    });
}
```
Replace with:
```js
const avatarBtn = document.getElementById('btn-avatar');
const avatarDropdown = document.getElementById('avatar-dropdown');
const avatarWrapper = document.getElementById('avatar-wrapper');

if (avatarBtn) {
    // Populate avatar: image or initials
    const user = window.__clerk && window.__clerk.user;
    if (user) {
        if (user.imageUrl) {
            avatarBtn.innerHTML = `<img src="${user.imageUrl}" alt="avatar" />`;
        } else {
            const full = user.fullName || '';
            const parts = full.trim().split(/\s+/);
            const initials = parts.length >= 2
                ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
                : (parts[0] ? parts[0][0].toUpperCase() : '?');
            avatarBtn.textContent = initials;
        }
    } else {
        avatarBtn.textContent = '?';
    }

    avatarBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const isOpen = avatarDropdown.style.display !== 'none';
        avatarDropdown.style.display = isOpen ? 'none' : 'flex';
        lucide.createIcons();
    });

    document.addEventListener('click', (e) => {
        if (!avatarWrapper.contains(e.target)) {
            avatarDropdown.style.display = 'none';
        }
    });
}

const btnManageAccount = document.getElementById('btn-manage-account');
if (btnManageAccount) {
    btnManageAccount.addEventListener('click', () => {
        avatarDropdown.style.display = 'none';
        window.__clerk.openUserProfile();
    });
}

const btnSignout = document.getElementById('btn-signout');
if (btnSignout) {
    btnSignout.addEventListener('click', async () => {
        avatarDropdown.style.display = 'none';
        if (window.__clerk) await window.__clerk.signOut();
        location.reload();
    });
}
```

- [ ] **Step 4: Remove stale `el` references**

In the `el` object at the top of `app.js`, find and remove these two lines:
```js
btnManageTeam: document.getElementById('btn-manage-team'),
btnLogout: document.getElementById('btn-logout'),
```

- [ ] **Step 5: Verify manually**

Reload the app. Confirm:
- Avatar button shows initials or profile image in the toolbar
- Clicking it opens the dropdown
- "Manage Account" opens Clerk's profile modal
- "Sign Out" signs out and reloads
- Clicking outside closes the dropdown

- [ ] **Step 6: Commit**

```bash
git add templates/index.html static/app.js static/style.css
git commit -m "feat: replace Manage Team + Sign Out with avatar profile dropdown"
```

---

### Task 3: Mobile People tab

**Files:**
- Modify: `templates/index.html` (add tab toggle + People list container to mobile view)
- Modify: `static/app.js` (add `_mobilePeopleCache`, `_mobileViewMode`, `buildMobilePeopleIndex`, `renderMobilePeople`, tab wiring)
- Modify: `static/style.css` (add mobile tab toggle + mobile person item styles)

**Interfaces:**
- Consumes: `_mobileCompaniesCache` (array of company objects with `directors: string[]`)
- Produces: `buildMobilePeopleIndex(companies)` → `{ [name: string]: string[] }` (director → company_name list); `renderMobilePeople(query)` renders the accordion list into `#mobile-people-list`

- [ ] **Step 1: Add tab toggle and People list to mobile HTML**

In `templates/index.html`, replace the entire mobile view block:
```html
<div id="mobile-view">
  <div class="mobile-header">
    <img src="/static/logo_white.png" id="mobile-logo" class="logo-img" alt="AGM Logo" />
    <div id="mobile-sync-banner" class="mobile-sync-banner">Loading...</div>
  </div>
  <div class="mobile-search-bar">
    <input type="text" id="mobile-search" placeholder="Search companies or UEN..." />
  </div>
  <div id="mobile-company-list" class="mobile-company-list">
    <div class="mobile-loading">
      <div class="spinner" style="margin:0 auto 12px;"></div>
      <p>Loading company data...</p>
    </div>
  </div>
</div>
```
with:
```html
<div id="mobile-view">
  <div class="mobile-header">
    <img src="/static/logo_white.png" id="mobile-logo" class="logo-img" alt="AGM Logo" />
    <div id="mobile-sync-banner" class="mobile-sync-banner">Loading...</div>
  </div>
  <div class="mobile-tab-toggle">
    <button id="mobile-tab-companies" class="mobile-tab-btn active">
      <i data-lucide="building"></i> Companies
    </button>
    <button id="mobile-tab-people" class="mobile-tab-btn">
      <i data-lucide="users"></i> People
    </button>
  </div>
  <div class="mobile-search-bar">
    <input type="text" id="mobile-search" placeholder="Search companies or UEN..." />
  </div>
  <div id="mobile-company-list" class="mobile-company-list">
    <div class="mobile-loading">
      <div class="spinner" style="margin:0 auto 12px;"></div>
      <p>Loading company data...</p>
    </div>
  </div>
  <div id="mobile-people-list" class="mobile-company-list" style="display:none;"></div>
</div>
```

- [ ] **Step 2: Add mobile tab toggle CSS**

Append to `static/style.css`:
```css
/* ── Mobile tab toggle ───────────────────────────────────────────────── */
.mobile-tab-toggle {
    display: flex;
    background: var(--bg-secondary);
    border-bottom: 1px solid var(--border-color);
    flex-shrink: 0;
}

.mobile-tab-btn {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    padding: 10px;
    font-size: 0.78rem;
    font-weight: 600;
    font-family: inherit;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    color: var(--text-muted);
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    cursor: pointer;
    transition: color 0.15s, border-color 0.15s;
}

.mobile-tab-btn i,
.mobile-tab-btn svg {
    width: 13px;
    height: 13px;
}

.mobile-tab-btn.active {
    color: var(--accent-primary);
    border-bottom-color: var(--accent-primary);
}

/* ── Mobile person accordion ─────────────────────────────────────────── */
.mobile-person-item {
    background: var(--bg-secondary);
    border: 1px solid var(--border-color);
    border-radius: 12px;
    margin-bottom: 10px;
    overflow: hidden;
}

.mobile-person-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 16px;
    cursor: pointer;
    gap: 10px;
}

.mobile-person-header:active {
    background: var(--bg-tertiary);
}

.mobile-person-name {
    font-size: 0.9rem;
    font-weight: 600;
    color: var(--text-primary);
    flex-grow: 1;
}

.mobile-person-count {
    font-size: 0.72rem;
    color: var(--text-muted);
    white-space: nowrap;
}

.mobile-person-chevron {
    width: 16px;
    height: 16px;
    color: var(--text-muted);
    transition: transform 0.2s ease;
    flex-shrink: 0;
}

.mobile-person-companies {
    display: none;
    padding: 0 16px 12px 16px;
    border-top: 1px dashed var(--border-color);
    background: var(--bg-tertiary);
}

.mobile-person-company {
    padding: 8px 4px;
    font-size: 0.82rem;
    color: var(--text-secondary);
    border-bottom: 1px solid var(--border-color);
}

.mobile-person-company:last-child {
    border-bottom: none;
}
```

- [ ] **Step 3: Add state variables and `buildMobilePeopleIndex`**

In `static/app.js`, find the existing state variables at the top of the mobile section:
```js
let _mobileCompaniesCache = [];
```
Replace with:
```js
let _mobileCompaniesCache = [];
let _mobilePeopleCache = {};   // { [directorName]: companyName[] }
let _mobileViewMode = 'companies'; // 'companies' | 'people'
```

Then add the `buildMobilePeopleIndex` function immediately after `loadMobileView`:
```js
function buildMobilePeopleIndex(companies) {
    const index = {};
    companies.forEach(c => {
        const cName = c.company_name || c.filename || 'Unnamed';
        (c.directors || []).forEach(dir => {
            const name = (dir || '').trim();
            if (!name) return;
            if (!index[name]) index[name] = [];
            if (!index[name].includes(cName)) index[name].push(cName);
        });
    });
    return index;
}
```

- [ ] **Step 4: Add `renderMobilePeople`**

Add the following function after `buildMobilePeopleIndex`:
```js
function renderMobilePeople(query) {
    const listEl = document.getElementById('mobile-people-list');
    if (!listEl) return;

    const q = (query || '').toLowerCase().trim();
    const sortedNames = Object.keys(_mobilePeopleCache).sort((a, b) => a.localeCompare(b));

    const filtered = q
        ? sortedNames.filter(name =>
            name.toLowerCase().includes(q) ||
            (_mobilePeopleCache[name] || []).some(c => c.toLowerCase().includes(q)))
        : sortedNames;

    if (filtered.length === 0) {
        listEl.innerHTML = `
            <div class="mobile-loading">
                <p>${q ? `No directors matching "${query}"` : 'No directors found. Sync companies first.'}</p>
            </div>`;
        return;
    }

    listEl.innerHTML = filtered.map((name, idx) => {
        const companies = _mobilePeopleCache[name] || [];
        const count = companies.length === 1 ? '1 company' : `${companies.length} companies`;
        return `
            <div class="mobile-person-item" id="mperson-${idx}">
                <div class="mobile-person-header" onclick="toggleMobilePerson(${idx})">
                    <span class="mobile-person-name">${name}</span>
                    <span class="mobile-person-count">${count}</span>
                    <i data-lucide="chevron-down" class="mobile-person-chevron" id="mpchevron-${idx}"></i>
                </div>
                <div class="mobile-person-companies" id="mpcompanies-${idx}">
                    ${companies.map(c => `<div class="mobile-person-company">${c}</div>`).join('')}
                </div>
            </div>`;
    }).join('');
    lucide.createIcons();
}
```

- [ ] **Step 5: Add `toggleMobilePerson`**

Add immediately after `renderMobilePeople`:
```js
function toggleMobilePerson(idx) {
    const body = document.getElementById(`mpcompanies-${idx}`);
    const chevron = document.getElementById(`mpchevron-${idx}`);
    if (!body) return;
    const isOpen = body.style.display === 'block';
    body.style.display = isOpen ? 'none' : 'block';
    if (chevron) chevron.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
}
```

- [ ] **Step 6: Update `loadMobileView` to build the people index after loading**

Replace the existing `loadMobileView` function:
```js
async function loadMobileView() {
    const listEl = document.getElementById('mobile-company-list');
    try {
        const res = await authFetch('/api/companies');
        if (!res.ok) throw new Error(`Server returned ${res.status}`);
        const { companies } = await res.json();
        _mobileCompaniesCache = companies || [];
        _mobilePeopleCache = buildMobilePeopleIndex(_mobileCompaniesCache);
        updateMobileSyncBanner(_mobileCompaniesCache);
        if (_mobileViewMode === 'people') {
            renderMobilePeople(document.getElementById('mobile-search')?.value || '');
        } else {
            renderMobileCards(_mobileCompaniesCache, '');
        }
    } catch (err) {
        console.error('loadMobileView error:', err);
        if (listEl) {
            listEl.innerHTML = `<p style="padding:20px;color:var(--color-error);">Failed to load: ${err.message}</p>`;
        }
    }
}
```

- [ ] **Step 7: Wire up tab buttons and search in `initMobileOrDesktop`**

Find `initMobileOrDesktop` in `app.js`. After the splash dismissal and before `loadMobileView()`, add tab and search wiring:
```js
// Mobile tab toggle
const tabCompanies = document.getElementById('mobile-tab-companies');
const tabPeople = document.getElementById('mobile-tab-people');
const mobileCompanyList = document.getElementById('mobile-company-list');
const mobilePeopleList = document.getElementById('mobile-people-list');
const mobileSearch = document.getElementById('mobile-search');

if (tabCompanies && tabPeople) {
    tabCompanies.addEventListener('click', () => {
        _mobileViewMode = 'companies';
        tabCompanies.classList.add('active');
        tabPeople.classList.remove('active');
        mobileCompanyList.style.display = '';
        mobilePeopleList.style.display = 'none';
        if (mobileSearch) mobileSearch.placeholder = 'Search companies or UEN...';
    });

    tabPeople.addEventListener('click', () => {
        _mobileViewMode = 'people';
        tabPeople.classList.add('active');
        tabCompanies.classList.remove('active');
        mobileCompanyList.style.display = 'none';
        mobilePeopleList.style.display = '';
        if (mobileSearch) mobileSearch.placeholder = 'Search directors...';
        renderMobilePeople(mobileSearch ? mobileSearch.value : '');
    });
}

// Wire mobile search to active tab
if (mobileSearch) {
    mobileSearch.addEventListener('input', (e) => {
        if (_mobileViewMode === 'people') {
            renderMobilePeople(e.target.value);
        } else {
            renderMobileCards(_mobileCompaniesCache, e.target.value);
        }
    });
}
```

Also remove the old search wiring that appears further down in `app.js` (the block that used to wire `mobile-search` to `renderMobileCards` only):
```js
// Wire up mobile search
const mobileSearch = document.getElementById('mobile-search');
if (mobileSearch) {
    mobileSearch.addEventListener('input', (e) => {
        renderMobileCards(_mobileCompaniesCache, e.target.value);
    });
}
```

- [ ] **Step 8: Verify manually**

On a narrow viewport (< 768px) or mobile device:
1. Load the app — Companies tab is active, company cards load
2. Tap "People" tab — list switches to director accordions
3. Tap a director row — expands to show their companies
4. Tap again — collapses
5. Type a director name in the search box — list filters correctly
6. Switch back to Companies tab — company list reappears, search placeholder resets

- [ ] **Step 9: Commit**

```bash
git add templates/index.html static/app.js static/style.css
git commit -m "feat: add People tab to mobile view with director accordion"
```

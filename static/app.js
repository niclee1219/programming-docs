/**
 * AGM Generator — Web-Native Application JavaScript
 * File reading moved to browser via File System Access API + SheetJS.
 * Files never leave the local disk — only extracted JSON reaches the server.
 * Generated documents are returned as a zip download.
 */

// ------------------------------------------------------------------ #
// SheetJS Extraction Constants (ported from config.py)
// ------------------------------------------------------------------ #

const SHEET_CANDIDATES = [
    'Export', 'EXPORT', 'export',
    'Data', 'DATA', 'data',
    'Info', 'INFO', 'Company', 'COMPANY',
    'Summary', 'SUMMARY', 'Details', 'DETAILS',
];

const KEY_INFO_CELLS = {
    company_name:       'B2',
    reg_no:             'B5',
    address:            'B7',
    financial_year_end: 'B11',
    agm_date:           'B14',
};

const FIELD_KEYWORDS = {
    company_name:       ['company name', 'name of company', 'company:'],
    reg_no:             ['reg no', 'registration no', 'uen', 'reg. no', 'company reg'],
    address:            ['registered address', 'office address', 'address'],
    financial_year_end: ['financial year end', 'fye', 'year end', 'year ended', 'fy end'],
    agm_number:         ['agm no', 'agm number', 'agm #', 'meeting no', 'annual general meeting no'],
    agm_date:           ['agm date', 'meeting date', 'date of agm', 'date of meeting'],
};

const DEFAULT_CELLS = {
    company_name:       ['B2', 'B1', 'C2', 'A2'],
    reg_no:             ['B3', 'B4', 'C3', 'D3'],
    address:            ['B5', 'B6', 'C5', 'D5'],
    financial_year_end: ['B7', 'B8', 'C7'],
    agm_number:         ['B9', 'B10', 'C9'],
    agm_date:           ['B11', 'B12', 'C11'],
};

const SCALAR_FIELDS = ['company_name', 'reg_no', 'address', 'financial_year_end', 'agm_number', 'agm_date'];

// ------------------------------------------------------------------ #
// SheetJS Extraction Helpers
// ------------------------------------------------------------------ #

function getCellValue(ws, ref) {
    const cell = ws[ref];
    if (!cell || cell.v === undefined || cell.v === null) return '';
    if (cell.v instanceof Date) return cell.v;
    return String(cell.v).trim();
}

function formatExcelDate(val) {
    if (!val) return '';
    if (val instanceof Date) {
        const months = ['January','February','March','April','May','June',
                        'July','August','September','October','November','December'];
        return `${val.getDate()} ${months[val.getMonth()]} ${val.getFullYear()}`;
    }
    if (typeof val === 'string') {
        const s = val.trim();
        if (!s) return '';
        const fmts = [
            [/^(\d{4})-(\d{2})-(\d{2})/, (m) => new Date(+m[1], +m[2]-1, +m[3])],
            [/^(\d{2})\/(\d{2})\/(\d{4})/, (m) => new Date(+m[3], +m[2]-1, +m[1])],
            [/^(\d{2})-(\d{2})-(\d{4})/, (m) => new Date(+m[3], +m[2]-1, +m[1])],
        ];
        for (const [rx, fn] of fmts) {
            const m = s.match(rx);
            if (m) { const d = fn(m); return formatExcelDate(d); }
        }
        return s;
    }
    return String(val).trim();
}

function findSheet(wb, candidates) {
    for (const cand of candidates) {
        const found = wb.SheetNames.find(n => n.trim().toLowerCase() === cand.toLowerCase());
        if (found) return wb.Sheets[found];
    }
    return wb.SheetNames.length ? wb.Sheets[wb.SheetNames[0]] : null;
}

function keywordSearchSheet(ws) {
    const results = {};
    if (!ws || !ws['!ref']) return results;

    const range = XLSX.utils.decode_range(ws['!ref']);
    const maxRow = Math.min(range.e.r, 79);

    for (let r = range.s.r; r <= maxRow; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({r, c});
            const cell = ws[addr];
            if (!cell || !cell.v) continue;
            const label = String(cell.v).trim().toLowerCase();

            for (const [field, keywords] of Object.entries(FIELD_KEYWORDS)) {
                if (results[field]) continue;
                for (const kw of keywords) {
                    if (label.includes(kw)) {
                        for (const offset of [1, 2]) {
                            const adjAddr = XLSX.utils.encode_cell({r, c: c + offset});
                            const adj = ws[adjAddr];
                            if (adj && adj.v !== undefined && adj.v !== null && String(adj.v).trim()) {
                                results[field] = {value: adj.v, cell: adjAddr};
                                break;
                            }
                        }
                        if (!results[field]) {
                            const belowAddr = XLSX.utils.encode_cell({r: r+1, c});
                            const below = ws[belowAddr];
                            if (below && below.v !== undefined && below.v !== null && String(below.v).trim()) {
                                results[field] = {value: below.v, cell: belowAddr};
                            }
                        }
                        break;
                    }
                }
            }
        }
    }
    return results;
}

function readDirectorsFromSheet(wb) {
    const exportName = wb.SheetNames.find(n => n.trim().toLowerCase() === 'export');
    if (exportName) {
        const ws = wb.Sheets[exportName];
        const directors = [];
        let row = 3;
        while (row < 200) {
            const addr = XLSX.utils.encode_cell({r: row, c: 0});
            const cell = ws[addr];
            if (!cell || cell.v === undefined || cell.v === null) break;
            const val = String(cell.v).trim();
            if (!val || val.toLowerCase() === 'n/a') break;
            directors.push(val);
            row++;
        }
        if (directors.length > 0) return directors;
    }

    const ws = findSheet(wb, SHEET_CANDIDATES);
    if (!ws || !ws['!ref']) return [];

    const range = XLSX.utils.decode_range(ws['!ref']);
    const dirKeywords = ['directors', 'board of directors', 'director names', 'name of directors'];

    for (let r = range.s.r; r <= Math.min(range.e.r, 79); r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({r, c});
            const cell = ws[addr];
            if (!cell || !cell.v) continue;
            const label = String(cell.v).trim().toLowerCase();
            if (dirKeywords.some(kw => label.includes(kw))) {
                const directors = [];
                const skip = new Set(['directors','name','director name','director names','s/n','no.','no']);
                for (let dr = r + 1; dr <= Math.min(r + 30, range.e.r); dr++) {
                    const da = XLSX.utils.encode_cell({r: dr, c});
                    const dc = ws[da];
                    if (!dc || !dc.v) continue;
                    const name = String(dc.v).trim();
                    if (!name || skip.has(name.toLowerCase())) continue;
                    if (directors.length > 0 && /^(Note|Remark|Total)/.test(name)) break;
                    directors.push(name);
                }
                if (directors.length > 0) return directors;
            }
        }
    }
    return [];
}

async function extractCompanyData(fileHandle, filename, savedMappings) {
    const file = await fileHandle.getFile();
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, {type: 'array', cellDates: true});

    const result = {
        company_name: '', reg_no: '', address: '',
        financial_year_end: '', agm_number: '', agm_date: '',
        directors: [], errors: [],
    };

    const keyInfoName = wb.SheetNames.find(n => n.trim().toLowerCase() === 'key information');
    let tier1ok = false;
    if (keyInfoName) {
        const ws = wb.Sheets[keyInfoName];
        result.company_name       = getCellValue(ws, 'B2');
        result.reg_no             = getCellValue(ws, 'B5');
        result.address            = getCellValue(ws, 'B7');
        result.financial_year_end = formatExcelDate(getCellValue(ws, 'B11'));
        result.agm_date           = formatExcelDate(getCellValue(ws, 'B14'));
        tier1ok = true;
    }

    const missing = SCALAR_FIELDS.filter(f => f !== 'agm_number' && !result[f]);

    if (!tier1ok || missing.length > 0) {
        const ws = findSheet(wb, SHEET_CANDIDATES);
        const learnedMappings = {};

        if (ws) {
            const fieldsToTry = tier1ok ? missing : SCALAR_FIELDS;
            for (const field of fieldsToTry) {
                if (result[field]) continue;
                const saved = savedMappings[field];
                if (saved && saved.cell) {
                    const raw = getCellValue(ws, saved.cell);
                    if (raw) {
                        result[field] = (field === 'financial_year_end' || field === 'agm_date')
                            ? formatExcelDate(raw) : String(raw).trim();
                    }
                }
            }

            const kwResults = keywordSearchSheet(ws);
            for (const field of fieldsToTry) {
                if (result[field]) continue;
                if (kwResults[field]) {
                    const raw = kwResults[field].value;
                    result[field] = (field === 'financial_year_end' || field === 'agm_date')
                        ? formatExcelDate(raw) : String(raw).trim();
                    learnedMappings[field] = kwResults[field].cell;
                }
            }

            for (const field of fieldsToTry) {
                if (result[field]) continue;
                for (const ref of (DEFAULT_CELLS[field] || [])) {
                    const raw = getCellValue(ws, ref);
                    if (raw) {
                        result[field] = (field === 'financial_year_end' || field === 'agm_date')
                            ? formatExcelDate(raw) : String(raw).trim();
                        learnedMappings[field] = ref;
                        break;
                    }
                }
                if (!result[field] && field !== 'agm_number') {
                    result.errors.push(`Could not locate: ${field}`);
                }
            }
        }

        if (Object.keys(learnedMappings).length > 0) {
            authFetch('/api/mappings', {
                method: 'POST',
                body: JSON.stringify({filename, mappings: learnedMappings}),
            }).catch(() => {});
        }
    }

    result.directors = readDirectorsFromSheet(wb);
    if (result.directors.length === 0) result.errors.push('No directors found');

    return result;
}

// ------------------------------------------------------------------ #
// Authenticated Fetch Helper
// ------------------------------------------------------------------ #

async function authFetch(url, options = {}) {
    const session = window.__clerk?.session;
    if (!session) throw new Error('No active Clerk session');

    const tokenPromise = session.getToken();
    const timeout = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Token fetch timed out')), 8000)
    );
    const token = await Promise.race([tokenPromise, timeout]);

    return fetch(url, {
        ...options,
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`,
            ...(options.headers || {}),
        },
    });
}

// ------------------------------------------------------------------ #
// Application State
// ------------------------------------------------------------------ #

const state = {
    settings: { agm_financial_year: String(new Date().getFullYear() - 1) },
    companies: {},
    selectedPath: null,
    activeTab: 'overview-tab',
    searchMode: 'companies',
    people: {},
    expandedPeople: [],
    queue: [],
    activeWorkers: 0,
    MAX_WORKERS: 2,
    dirHandle: null,
    fileHandles: {},
    syncFired: false,   // prevents double-fire of auto-sync per scan
};

// DOM Elements (desktop app)
const el = {
    xlsmFolder: document.getElementById('xlsm-folder'),
    btnBrowseXlsm: document.getElementById('btn-browse-xlsm'),
    btnScan: document.getElementById('btn-scan'),

    searchFilter: document.getElementById('search-filter'),
    companyList: document.getElementById('company-list'),
    companyCount: document.getElementById('company-count'),
    btnSelectAll: document.getElementById('btn-select-all'),
    btnDeselectAll: document.getElementById('btn-deselect-all'),

    tabBtns: document.querySelectorAll('.tab-btn'),
    tabPanes: document.querySelectorAll('.tab-pane'),

    btnModeCompanies: document.getElementById('btn-mode-companies'),
    btnModePeople: document.getElementById('btn-mode-people'),

    btnToggleSidebar: document.getElementById('btn-toggle-sidebar'),
    btnTogglePreview: document.getElementById('btn-toggle-preview'),
    btnClosePreview: document.getElementById('btn-close-preview'),
    btnThemeToggle: document.getElementById('btn-theme-toggle'),
    appLogo: document.getElementById('app-logo'),
    mainWorkspace: document.getElementById('main-workspace'),
    previewSidebar: document.getElementById('preview-sidebar'),
    previewDocument: document.getElementById('preview-document'),

    detailsView: document.getElementById('details-view'),
    detailsEmptyView: document.getElementById('details-empty-view'),
    ordinalBadge: document.getElementById('ordinal-badge'),
    ordinalText: document.getElementById('ordinal-text'),

    companyName: document.getElementById('company-name'),
    regNo: document.getElementById('reg-no'),
    address: document.getElementById('address'),
    fyEnd: document.getElementById('financial-year-end'),

    agmDate: document.getElementById('agm-date'),
    agmDatePicker: document.getElementById('agm-date-picker'),

    agmNumber: document.getElementById('agm-number'),
    directorsList: document.getElementById('directors-list'),
    btnHistoryRotation: document.getElementById('btn-history-rotation'),

    warningsBox: document.getElementById('warnings-box'),
    warningsList: document.getElementById('warnings-list'),

    btnSave: document.getElementById('btn-save'),
    btnReread: document.getElementById('btn-reread'),

    companyTooltip: document.getElementById('company-tooltip'),

    statusText: document.getElementById('status-text'),
    progressContainer: document.getElementById('generation-progress-container'),
    progressBar: document.getElementById('generation-progress'),
    progressText: document.getElementById('progress-text'),
    btnGenerate: document.getElementById('btn-generate'),

    historyModal: document.getElementById('history-modal'),
    historyCompanyName: document.getElementById('history-company-name'),
    historyCompanyUen: document.getElementById('history-company-uen'),
    historyPlanningYear: document.getElementById('history-planning-year'),
    historyLoading: document.getElementById('history-loading'),
    historyTableContainer: document.getElementById('history-table-container'),
    historyTableBody: document.getElementById('history-table-body'),
    btnCloseModal: document.getElementById('btn-close-modal'),
    btnConfirmClose: document.getElementById('btn-confirm-close'),
};

const ORDINALS = {
    1: "First", 2: "Second", 3: "Third", 4: "Fourth", 5: "Fifth",
    6: "Sixth", 7: "Seventh", 8: "Eighth", 9: "Ninth", 10: "Tenth",
    11: "Eleventh", 12: "Twelfth", 13: "Thirteenth", 14: "Fourteenth",
    15: "Fifteenth", 16: "Sixteenth", 17: "Seventeenth", 18: "Eighteenth",
    19: "Nineteenth", 20: "Twentieth", 21: "Twenty-First", 22: "Twenty-Second",
    23: "Twenty-Third", 24: "Twenty-Fourth", 25: "Twenty-Fifth",
    26: "Twenty-Sixth", 27: "Twenty-Seventh", 28: "Twenty-Eighth",
    29: "Twenty-Ninth", 30: "Thirtieth",
};

// ------------------------------------------------------------------ #
// Init and Setup
// ------------------------------------------------------------------ #

document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    // loadSettings and registerEvents are called from initMobileOrDesktop after Clerk auth resolves
});

// Called from the Clerk auth gate in index.html once the user is authenticated
function initMobileOrDesktop() {
    if (window.innerWidth < 768) {
        // Mobile: hide desktop app, show read-only mobile view
        document.querySelector('.app-container').style.display = 'none';
        const mobileView = document.getElementById('mobile-view');
        if (mobileView) mobileView.style.display = 'flex';

        // Apply theme logo to mobile header
        const mobileLogo = document.getElementById('mobile-logo');
        if (mobileLogo) {
            const isLight = document.body.classList.contains('light-theme');
            mobileLogo.src = isLight ? '/static/logo_black.png' : '/static/logo_white.png';
        }

        // Dismiss the welcome splash before loading data
        const splash = document.getElementById('welcome-splash');
        if (splash) {
            splash.classList.add('splash-exit');
            setTimeout(() => splash.remove(), 600);
        }

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
                renderMobileCards(_mobileCompaniesCache, mobileSearch ? mobileSearch.value : '');
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

        loadMobileView();
    } else {
        // Desktop: boot normally
        loadSettings();
        registerEvents();
        setTimeout(() => {
            const splash = document.getElementById('welcome-splash');
            if (splash) {
                splash.classList.add('splash-exit');
                setTimeout(() => splash.remove(), 600);
            }
        }, 800);
    }
}

function initTheme() {
    const savedTheme = localStorage.getItem('theme');
    const toggleBtn = document.getElementById('btn-theme-toggle');
    if (!toggleBtn) return;

    if (savedTheme === 'light') {
        document.body.classList.add('light-theme');
        if (el.appLogo) el.appLogo.src = '/static/logo_black.png';
        toggleBtn.innerHTML = '<i data-lucide="moon"></i> <span class="theme-label">Dark Mode</span>';
    } else {
        document.body.classList.remove('light-theme');
        if (el.appLogo) el.appLogo.src = '/static/logo_white.png';
        toggleBtn.innerHTML = '<i data-lucide="sun"></i> <span class="theme-label">Light Mode</span>';
    }
    lucide.createIcons();
}

function toggleTheme() {
    const isLight = document.body.classList.toggle('light-theme');
    localStorage.setItem('theme', isLight ? 'light' : 'dark');
    const toggleBtn = document.getElementById('btn-theme-toggle');
    if (!toggleBtn) return;
    if (isLight) {
        if (el.appLogo) el.appLogo.src = '/static/logo_black.png';
        toggleBtn.innerHTML = '<i data-lucide="moon"></i> <span class="theme-label">Dark Mode</span>';
    } else {
        if (el.appLogo) el.appLogo.src = '/static/logo_white.png';
        toggleBtn.innerHTML = '<i data-lucide="sun"></i> <span class="theme-label">Light Mode</span>';
    }
    lucide.createIcons();
}

function registerEvents() {
    el.btnBrowseXlsm.addEventListener('click', pickFolder);
    el.btnScan.addEventListener('click', scanFolder);
    el.btnThemeToggle.addEventListener('click', toggleTheme);

    el.btnToggleSidebar.addEventListener('click', () => {
        el.mainWorkspace.classList.toggle('sidebar-collapsed');
    });
    el.btnTogglePreview.addEventListener('click', () => {
        el.mainWorkspace.classList.toggle('preview-expanded');
    });
    el.btnClosePreview.addEventListener('click', () => {
        el.mainWorkspace.classList.remove('preview-expanded');
    });

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

        const btnManageAccount = document.getElementById('btn-manage-account');
        if (btnManageAccount) {
            btnManageAccount.addEventListener('click', () => {
                avatarDropdown.style.display = 'none';
                if (window.__clerk) window.__clerk.openUserProfile();
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
    }

    el.searchFilter.addEventListener('input', filterCompanyList);
    el.btnSelectAll.addEventListener('click', () => toggleAllSelections(true));
    el.btnDeselectAll.addEventListener('click', () => toggleAllSelections(false));

    if (el.btnModeCompanies && el.btnModePeople) {
        el.btnModeCompanies.addEventListener('click', () => {
            if (state.searchMode === 'companies') return;
            state.searchMode = 'companies';
            el.btnModeCompanies.classList.add('active');
            el.btnModePeople.classList.remove('active');
            el.searchFilter.placeholder = 'Filter companies by name...';
            el.searchFilter.value = '';
            renderCompanyList();
        });

        el.btnModePeople.addEventListener('click', () => {
            if (el.btnModePeople.classList.contains('disabled')) {
                showToast('People Mode is compiled only after all files are fully parsed.', 'warn');
                return;
            }
            if (state.searchMode === 'people') return;
            state.searchMode = 'people';
            el.btnModePeople.classList.add('active');
            el.btnModeCompanies.classList.remove('active');
            el.searchFilter.placeholder = 'Search by name or company...';
            el.searchFilter.value = '';
            renderCompanyList();
        });
    }

    const formInputs = [el.companyName, el.regNo, el.address, el.fyEnd, el.agmNumber];
    formInputs.forEach(input => {
        input.addEventListener('input', () => {
            syncFormToCache();
            updateLivePreview();
        });
    });
    el.agmNumber.addEventListener('input', updateOrdinalPreview);

    document.querySelector('.datepicker-trigger').addEventListener('click', () => {
        el.agmDatePicker.showPicker();
    });
    el.agmDatePicker.addEventListener('change', (e) => {
        const val = e.target.value;
        if (!val) return;
        const formatted = formatPickerDate(val);
        el.agmDate.value = formatted;
        el.agmDate.classList.remove('is-default-date');
        if (state.selectedPath && state.companies[state.selectedPath]) {
            const c = state.companies[state.selectedPath];
            if (c.data) {
                c.data.is_custom_agm = true;
                c.data.is_default_agm = false;
            }
        }
        syncFormToCache();
        updateLivePreview();
    });

    el.btnHistoryRotation.addEventListener('click', openRotationHistory);
    el.btnSave.addEventListener('click', () => {
        saveCompanyEdits();
        showStatus('Edits saved to local memory.', 'success');
    });
    el.btnReread.addEventListener('click', rereadSelectedCompany);

    el.tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            el.tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const tabName = btn.getAttribute('data-tab');
            state.activeTab = tabName;
            el.tabPanes.forEach(pane => pane.classList.remove('active'));
            document.getElementById(tabName).classList.add('active');
        });
    });

    el.btnGenerate.addEventListener('click', runGeneration);

    const closeModal = () => el.historyModal.style.display = 'none';
    el.btnCloseModal.addEventListener('click', closeModal);
    el.btnConfirmClose.addEventListener('click', closeModal);
}

// ------------------------------------------------------------------ #
// Date Conversion Utilities
// ------------------------------------------------------------------ #

function formatPickerDate(dateStr) {
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const year = parts[0];
    const monthIdx = parseInt(parts[1]) - 1;
    const day = parseInt(parts[2]);
    const months = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    return `${day} ${months[monthIdx]} ${year}`;
}

function parseFormattedDate(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.trim().split(/\s+/);
    if (parts.length !== 3) return '';
    const day = parts[0].padStart(2, '0');
    const monthStr = parts[1];
    const year = parts[2];
    const months = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    const monthIdx = months.findIndex(m => m.toLowerCase() === monthStr.toLowerCase());
    if (monthIdx === -1) return '';
    const month = String(monthIdx + 1).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// ------------------------------------------------------------------ #
// Toast Notification Engine
// ------------------------------------------------------------------ #

function showToast(message, type = 'info', duration = 3500) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'info';
    if (type === 'success') icon = 'check-circle';
    if (type === 'error') icon = 'alert-octagon';
    if (type === 'warn') icon = 'alert-triangle';

    toast.innerHTML = `
        <i data-lucide="${icon}" class="toast-icon"></i>
        <span class="toast-message">${message}</span>
        <button class="toast-close">&times;</button>
    `;

    container.appendChild(toast);
    lucide.createIcons();

    requestAnimationFrame(() => {
        requestAnimationFrame(() => { toast.classList.add('toast-visible'); });
    });

    const dismiss = () => {
        toast.classList.remove('toast-visible');
        toast.classList.add('toast-exit');
        setTimeout(() => toast.remove(), 400);
    };

    toast.querySelector('.toast-close').addEventListener('click', dismiss);
    setTimeout(dismiss, duration);
}

// ------------------------------------------------------------------ #
// Settings
// ------------------------------------------------------------------ #

async function loadSettings() {
    try {
        const response = await authFetch('/api/settings');
        const s = await response.json();
        state.settings = s;
    } catch (err) {
        showStatus('Failed to load settings from server.', 'error');
    }
}

async function saveSettings() {
    try {
        await authFetch('/api/settings', {
            method: 'POST',
            body: JSON.stringify(state.settings),
        });
    } catch (err) {
        showStatus('Failed to save settings to server.', 'error');
    }
}

// ------------------------------------------------------------------ #
// Folder Picker — File System Access API (Chrome/Edge)
// ------------------------------------------------------------------ #

async function pickFolder() {
    if (!window.showDirectoryPicker) {
        showStatus('Folder picker requires Chrome or Edge. Please switch browsers.', 'error');
        return;
    }
    try {
        state.dirHandle = await window.showDirectoryPicker({ mode: 'read' });
        el.xlsmFolder.value = state.dirHandle.name;
        showStatus(`Folder "${state.dirHandle.name}" selected.`, 'success');
        await scanFolder();
    } catch (err) {
        if (err.name !== 'AbortError') {
            showStatus('Could not open folder picker.', 'error');
        }
    }
}

// ------------------------------------------------------------------ #
// Folder Scanner
// ------------------------------------------------------------------ #

async function scanFolder() {
    if (!state.dirHandle) {
        await pickFolder();
        return;
    }

    showStatus('Scanning folder...');
    el.btnScan.disabled = true;
    state.syncFired = false;  // reset sync guard for new scan

    try {
        const prevCompanies = state.companies;
        state.companies = {};
        state.queue = [];
        state.people = {};
        state.expandedPeople = [];
        state.fileHandles = {};

        for await (const [name, handle] of state.dirHandle.entries()) {
            if (!name.endsWith('.xlsm')) continue;
            state.fileHandles[name] = handle;
            state.companies[name] = {
                path: name,
                filename: name,
                status: 'pending',
                checked: prevCompanies[name] ? prevCompanies[name].checked : false,
                data: prevCompanies[name] ? prevCompanies[name].data : null,
            };
            if (!state.companies[name].data) {
                state.queue.push(name);
            } else {
                state.companies[name].status = prevCompanies[name].status;
            }
        }

        const count = Object.keys(state.companies).length;
        el.btnModePeople.classList.add('disabled');
        updateParsingProgress();
        rebuildKnowledgeGraph();
        renderCompanyList();
        updateSelectedCounter();

        showStatus(`Scan complete. Found ${count} client files.`, 'success');
        processBackgroundQueue();

    } catch (err) {
        showStatus(`Scan error: ${err.message}`, 'error');
    } finally {
        el.btnScan.disabled = false;
    }
}

// ------------------------------------------------------------------ #
// Background Loaders
// ------------------------------------------------------------------ #

function processBackgroundQueue() {
    if (state.activeWorkers === 0 && state.queue.length === 0) {
        rebuildKnowledgeGraph();
        renderCompanyList();
        updateParsingProgress();
        return;
    }

    while (state.activeWorkers < state.MAX_WORKERS && state.queue.length > 0) {
        const filename = state.queue.shift();
        state.activeWorkers++;
        updateRowStatusUI(filename, 'pending-pulse');
        readCompanyFile(filename).then(() => {
            state.activeWorkers--;
            updateParsingProgress();
            processBackgroundQueue();
        });
    }
}

async function readCompanyFile(filename, force = false) {
    const handle = state.fileHandles[filename];
    if (!handle) {
        state.companies[filename].status = 'error';
        state.companies[filename].raw_errors = ['File handle lost — re-scan the folder.'];
        state.companies[filename].data = {directors: []};
        applyDateDefaultRules(state.companies[filename].data);
        initDirectorsChecklist(state.companies[filename].data);
        rebuildKnowledgeGraph();
        updateRowUI(filename);
        return;
    }

    try {
        let savedMappings = {};
        try {
            const r = await authFetch(`/api/mappings?filename=${encodeURIComponent(filename)}`);
            const j = await r.json();
            savedMappings = j.mappings || {};
        } catch (_) {}

        const data = await extractCompanyData(handle, filename, savedMappings);

        state.companies[filename].data = {
            company_name:       data.company_name || '',
            reg_no:             data.reg_no || '',
            address:            data.address || '',
            financial_year_end: data.financial_year_end || '',
            agm_number:         data.agm_number || '',
            agm_date:           data.agm_date || '',
            directors:          data.directors || [],
        };

        applyDateDefaultRules(state.companies[filename].data);
        initDirectorsChecklist(state.companies[filename].data);

        let status = 'ok';
        if (data.errors && data.errors.length > 0) {
            status = data.company_name ? 'partial' : 'error';
        }
        state.companies[filename].status = status;
        state.companies[filename].raw_errors = data.errors || [];

        rebuildKnowledgeGraph();
        updateRowUI(filename);

        if (state.selectedPath === filename) {
            loadCompanyDetails(filename);
        }

    } catch (err) {
        state.companies[filename].status = 'error';
        state.companies[filename].raw_errors = ['Cannot open or parse this spreadsheet.'];
        state.companies[filename].data = {directors: []};
        applyDateDefaultRules(state.companies[filename].data);
        initDirectorsChecklist(state.companies[filename].data);
        rebuildKnowledgeGraph();
        updateRowUI(filename);
        if (state.selectedPath === filename) loadCompanyDetails(filename);
    }
}

function applyDateDefaultRules(dataObj) {
    const currentYear = new Date().getFullYear();
    const nextYear = String(currentYear + 1);
    dataObj.financial_year_end = `31 December ${currentYear}`;
    if (!dataObj.is_custom_agm) {
        dataObj.agm_date = `1 June ${nextYear}`;
        dataObj.is_default_agm = true;
    }
}

function initDirectorsChecklist(dataObj) {
    if (!dataObj.all_directors) {
        dataObj.all_directors = [...dataObj.directors];
        dataObj.selected_directors = [...dataObj.directors];
    }
}

// ------------------------------------------------------------------ #
// Parsing Progress Indicator
// ------------------------------------------------------------------ #

function updateParsingProgress() {
    const total = Object.keys(state.companies).length;
    if (total === 0) return;

    const loaded = Object.values(state.companies).filter(c => c.data).length;
    const container = document.getElementById('parsing-progress-container');
    const bar = document.getElementById('parsing-progress');
    const text = document.getElementById('parsing-progress-text');

    if (loaded < total) {
        if (container) container.style.display = 'flex';
        const pct = (loaded / total) * 100;
        if (bar) bar.style.width = `${pct}%`;
        if (text) text.textContent = `Loading... ${loaded}/${total}`;
        el.btnModePeople.classList.add('disabled');
    } else {
        if (container) {
            if (bar) bar.style.width = '100%';
            if (text) text.textContent = `Completed ${loaded}/${total}`;
            setTimeout(() => { container.style.display = 'none'; }, 800);
        }
        el.btnModePeople.classList.remove('disabled');
        rebuildKnowledgeGraph();

        // Auto-sync to cloud once all files are loaded (once per scan)
        if (!state.syncFired) {
            state.syncFired = true;
            runAutoSync();
        }
    }
}

// ------------------------------------------------------------------ #
// Auto-Sync to Supabase
// ------------------------------------------------------------------ #

async function runAutoSync() {
    if (!window.__clerk || !window.__clerk.session) return;

    const companies = Object.values(state.companies)
        .filter(c => c.data)
        .map(c => ({
            filename:            c.filename,
            company_name:        c.data.company_name || '',
            reg_no:              c.data.reg_no || '',
            address:             c.data.address || '',
            financial_year_end:  c.data.financial_year_end || '',
            agm_number:          c.data.agm_number || '',
            agm_date:            c.data.agm_date || '',
            all_directors:       c.data.all_directors || [],
            selected_directors:  c.data.selected_directors || [],
        }));

    if (companies.length === 0) return;

    try {
        // Phase 1: preview diffs (no DB write yet)
        const previewRes = await authFetch('/api/sync', {
            method: 'POST',
            body: JSON.stringify({ companies }),
        });
        if (!previewRes.ok) return;
        const { diffs, total } = await previewRes.json();

        if (diffs.length === 0) {
            // No changes — confirm immediately without user interaction
            await authFetch('/api/sync/confirm', {
                method: 'POST',
                body: JSON.stringify({
                    companies,
                    user_name: window.__clerk.user ? (window.__clerk.user.fullName || '') : '',
                }),
            });
            showToast(`Synced ${total} companies — no changes`, 'success', 2500);
        } else {
            showDiffModal(diffs, companies);
        }
    } catch (err) {
        // Sync is best-effort — never block the user
        console.warn('Auto-sync failed:', err);
    }
}

const _DIFF_FIELD_LABELS = {
    company_name:       'Company Name',
    reg_no:             'UEN / Reg No',
    address:            'Address',
    financial_year_end: 'Financial Year End',
    agm_number:         'AGM Number',
    agm_date:           'AGM Date',
    directors:          'Directors',
};

function showDiffModal(diffs, allCompanies) {
    const modal = document.getElementById('sync-diff-modal');
    const body = document.getElementById('sync-diff-body');
    if (!modal || !body) return;

    const plural = diffs.length === 1 ? 'company record has' : 'company records have';
    body.innerHTML = `
        <p style="margin-bottom:16px;color:var(--text-secondary);">
            ${diffs.length} ${plural} changed since last sync.
        </p>
        ${diffs.map(diff => `
            <div class="diff-company-block">
                <h4 class="diff-company-name">${diff.company_name}</h4>
                <table class="diff-table">
                    <thead>
                        <tr><th>Field</th><th class="diff-old-col">Before</th><th class="diff-new-col">After</th></tr>
                    </thead>
                    <tbody>
                        ${diff.changes.map(ch => `
                            <tr>
                                <td>${_DIFF_FIELD_LABELS[ch.field] || ch.field}</td>
                                <td class="diff-old">${ch.old || '<em>empty</em>'}</td>
                                <td class="diff-new">${ch.new || '<em>empty</em>'}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>
            </div>`).join('')}`;

    modal.style.display = 'flex';
    lucide.createIcons();

    document.getElementById('btn-sync-confirm').onclick = async () => {
        modal.style.display = 'none';
        try {
            await authFetch('/api/sync/confirm', {
                method: 'POST',
                body: JSON.stringify({
                    companies: allCompanies,
                    user_name: window.__clerk.user ? (window.__clerk.user.fullName || '') : '',
                }),
            });
            showToast(`Cloud data updated for ${allCompanies.length} companies`, 'success');
        } catch (err) {
            showToast('Failed to update cloud data', 'error');
        }
    };

    document.getElementById('btn-sync-cancel').onclick = () => {
        modal.style.display = 'none';
        showToast('Cloud update skipped — local data unchanged', 'info');
    };

    document.getElementById('btn-close-diff-modal').onclick = () => {
        modal.style.display = 'none';
    };
}

// ------------------------------------------------------------------ #
// Mobile View
// ------------------------------------------------------------------ #

let _mobileCompaniesCache = [];
let _mobilePeopleCache = {};   // { [directorName]: companyName[] }
let _mobileViewMode = 'companies'; // 'companies' | 'people'

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
        const activeList = _mobileViewMode === 'people'
            ? document.getElementById('mobile-people-list')
            : listEl;
        if (activeList) {
            activeList.innerHTML = `<p style="padding:20px;color:var(--color-error);">Failed to load: ${err.message}</p>`;
        }
    }
}

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

function renderMobilePeople(query) {
    const listEl = document.getElementById('mobile-people-list');
    if (!listEl) return;

    const esc = s => (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

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
                <p>${q ? `No directors matching "${esc(query)}"` : 'No directors found. Sync companies first.'}</p>
            </div>`;
        return;
    }

    listEl.innerHTML = filtered.map((name, idx) => {
        const companies = _mobilePeopleCache[name] || [];
        const count = companies.length === 1 ? '1 company' : `${companies.length} companies`;
        return `
            <div class="mobile-person-item" id="mperson-${idx}">
                <div class="mobile-person-header" onclick="toggleMobilePerson(${idx})">
                    <span class="mobile-person-name">${esc(name)}</span>
                    <span class="mobile-person-count">${count}</span>
                    <i data-lucide="chevron-down" class="mobile-person-chevron" id="mpchevron-${idx}"></i>
                </div>
                <div class="mobile-person-companies" id="mpcompanies-${idx}">
                    ${companies.map(c => `<div class="mobile-person-company">${esc(c)}</div>`).join('')}
                </div>
            </div>`;
    }).join('');
    lucide.createIcons({ elements: Array.from(listEl.querySelectorAll('[data-lucide]')) });
}

function toggleMobilePerson(idx) {
    const body = document.getElementById(`mpcompanies-${idx}`);
    const chevron = document.getElementById(`mpchevron-${idx}`);
    if (!body) return;
    const isOpen = body.style.display === 'block';
    body.style.display = isOpen ? 'none' : 'block';
    if (chevron) chevron.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
}

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

function renderMobileCards(companies, query) {
    const listEl = document.getElementById('mobile-company-list');
    if (!listEl) return;

    const q = (query || '').toLowerCase().trim();
    const filtered = q
        ? companies.filter(c =>
            (c.company_name || '').toLowerCase().includes(q) ||
            (c.reg_no || '').toLowerCase().includes(q))
        : companies;

    if (filtered.length === 0) {
        listEl.innerHTML = q
            ? `<p style="padding:20px;color:var(--text-muted);text-align:center;">No companies match "${query}".</p>`
            : '<p style="padding:20px;color:var(--text-muted);text-align:center;">No company data synced yet. Scan a folder from the desktop app.</p>';
        return;
    }

    listEl.innerHTML = filtered.map((c, idx) => `
        <div class="mobile-card" id="mobile-card-${idx}">
            <div class="mobile-card-header" onclick="toggleMobileCard(${idx})">
                <div>
                    <div class="mobile-card-title">${c.company_name || c.filename || 'Unnamed'}</div>
                    <div class="mobile-card-uen">${c.reg_no || '—'}</div>
                </div>
                <i data-lucide="chevron-down" class="mobile-card-chevron" id="chevron-${idx}"></i>
            </div>
            <div class="mobile-card-body" id="mobile-card-body-${idx}" style="display:none">
                <div class="mobile-field"><strong>Address:</strong><span>${c.address || '—'}</span></div>
                <div class="mobile-field"><strong>Financial Year End:</strong><span>${c.financial_year_end || '—'}</span></div>
                <div class="mobile-field"><strong>AGM:</strong><span>${_agmLabel(c)}</span></div>
                <div class="mobile-field mobile-field-directors">
                    <strong>Directors:</strong>
                    <ul class="mobile-directors-list">
                        ${(c.directors || []).map(d => `<li>• ${d}</li>`).join('') || '<li style="color:var(--text-muted)">None recorded</li>'}
                    </ul>
                </div>
            </div>
        </div>`).join('');

    lucide.createIcons();
}

function toggleMobileCard(idx) {
    const body = document.getElementById(`mobile-card-body-${idx}`);
    const chevron = document.getElementById(`chevron-${idx}`);
    if (!body) return;
    const isOpen = body.style.display !== 'none';
    body.style.display = isOpen ? 'none' : 'block';
    if (chevron) {
        chevron.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
    }
}

function _agmLabel(c) {
    const n = parseInt(c.agm_number || '0');
    const suffix = n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th';
    const num = n ? `${n}${suffix} AGM` : '';
    const date = c.agm_date || '';
    if (num && date) return `${num} — ${date}`;
    return num || date || '—';
}


// ------------------------------------------------------------------ #
// Knowledge Graph Engine
// ------------------------------------------------------------------ #

function rebuildKnowledgeGraph() {
    state.people = {};
    Object.values(state.companies).forEach(company => {
        if (company.data && company.data.all_directors) {
            company.data.all_directors.forEach(directorName => {
                const name = directorName.trim();
                if (!name) return;
                if (!state.people[name]) state.people[name] = [];
                if (!state.people[name].includes(company.path)) state.people[name].push(company.path);
            });
        }
    });
}

// ------------------------------------------------------------------ #
// UI Rendering - Sidebar
// ------------------------------------------------------------------ #

function renderCompanyList() {
    if (state.searchMode === 'people') { renderPeopleList(); return; }

    el.companyList.innerHTML = '';
    const statusWeight = {ok:1, pending:1, 'pending-pulse':1, partial:2, error:3};
    const items = Object.values(state.companies).sort((a, b) => {
        const wA = statusWeight[a.status] || 1;
        const wB = statusWeight[b.status] || 1;
        if (wA !== wB) return wA - wB;
        const aName = (a.data && a.data.company_name) ? a.data.company_name.toLowerCase() : a.filename.toLowerCase();
        const bName = (b.data && b.data.company_name) ? b.data.company_name.toLowerCase() : b.filename.toLowerCase();
        return aName.localeCompare(bName);
    });

    if (items.length === 0) {
        el.companyList.innerHTML = `
            <div class="empty-list-placeholder">
                <i data-lucide="folder-search" class="placeholder-icon"></i>
                <p>No .xlsm client files found in the selected folder.</p>
            </div>`;
        lucide.createIcons();
        return;
    }

    const successfulPendingItems = items.filter(c => c.status !== 'error');
    const errorItems = items.filter(c => c.status === 'error');

    function createItemElement(c) {
        const li = document.createElement('li');
        const statusClass = c.status;
        li.className = `company-item ${state.selectedPath === c.path ? 'selected' : ''} ${statusClass}-status`;
        li.setAttribute('data-path', c.path);

        const isChecked = c.checked ? 'checked' : '';
        const displayName = (c.data && c.data.company_name) ? c.data.company_name : c.filename;

        let subName = c.filename;
        if (c.status === 'pending') {
            subName = 'Waiting to read...';
        } else if (c.status === 'error') {
            subName = `Error: ${(c.raw_errors && c.raw_errors.length > 0) ? c.raw_errors[0] : 'Failed to parse file.'}`;
        } else if (c.status === 'partial') {
            const missing = [];
            if (!c.data.company_name) missing.push('Company Name');
            if (!c.data.reg_no) missing.push('UEN');
            if (!c.data.address) missing.push('AGM Address');
            if (!c.data.financial_year_end) missing.push('FYE');
            if (!c.data.agm_number) missing.push('AGM Number');
            if (!c.data.agm_date) missing.push('AGM Date');
            subName = missing.length > 0 ? `Missing: ${missing.join(', ')}` :
                      (c.raw_errors && c.raw_errors.length > 0) ? `Warning: ${c.raw_errors[0]}` : 'Warning: Missing information';
        }

        li.innerHTML = `
            <div class="custom-checkbox">
                <input type="checkbox" ${isChecked}>
                <div class="checkbox-visual"><i data-lucide="check"></i></div>
            </div>
            <div class="status-badge ${statusClass}"></div>
            <div class="company-info">
                <span class="company-title" title="${displayName}">${displayName}</span>
                <span class="company-filename" title="${subName}">${subName}</span>
            </div>`;

        const checkbox = li.querySelector('input[type="checkbox"]');
        checkbox.addEventListener('change', (e) => {
            state.companies[c.path].checked = e.target.checked;
            updateSelectedCounter();
        });
        li.addEventListener('click', (e) => {
            if (e.target.closest('.custom-checkbox')) return;
            selectCompany(c.path);
        });
        return li;
    }

    successfulPendingItems.forEach(c => el.companyList.appendChild(createItemElement(c)));

    if (errorItems.length > 0) {
        const divider = document.createElement('li');
        divider.className = 'sidebar-divider';
        divider.innerHTML = '<span>Errors / Unloaded Files</span>';
        el.companyList.appendChild(divider);
        errorItems.forEach(c => el.companyList.appendChild(createItemElement(c)));
    }

    lucide.createIcons();
}

function renderPeopleList() {
    el.companyList.innerHTML = '';
    const query = el.searchFilter.value.toLowerCase().trim();
    const sortedPeople = Object.keys(state.people).sort((a, b) => a.localeCompare(b));

    const filteredPeople = sortedPeople.filter(name => {
        if (name.toLowerCase().includes(query)) return true;
        return (state.people[name] || []).some(path => {
            const c = state.companies[path];
            if (!c) return false;
            if (c.filename.toLowerCase().includes(query)) return true;
            return c.data && c.data.company_name && c.data.company_name.toLowerCase().includes(query);
        });
    });

    if (filteredPeople.length === 0) {
        el.companyList.innerHTML = `
            <div class="empty-list-placeholder">
                <i data-lucide="users" class="placeholder-icon"></i>
                <p>No directors or companies found matching "${query}".</p>
            </div>`;
        lucide.createIcons();
        return;
    }

    filteredPeople.forEach(personName => {
        const companyPaths = state.people[personName] || [];
        const isExpanded = state.expandedPeople.includes(personName);
        const li = document.createElement('li');
        li.className = 'person-item';
        const countLabel = companyPaths.length === 1 ? '1 company' : `${companyPaths.length} companies`;

        li.innerHTML = `
            <div class="person-header">
                <i data-lucide="user" class="person-icon"></i>
                <div class="person-info">
                    <span class="person-name">${personName}</span>
                    <span class="person-subtitle">${countLabel}</span>
                </div>
                <i data-lucide="chevron-right" class="person-arrow" style="transform: ${isExpanded ? 'rotate(90deg)' : 'rotate(0deg)'};"></i>
            </div>
            <ul class="person-companies-list" style="display: ${isExpanded ? 'block' : 'none'};">
                ${companyPaths.map(path => {
                    const c = state.companies[path];
                    const cName = (c && c.data && c.data.company_name) ? c.data.company_name : (c ? c.filename : path);
                    return `<li class="person-company-subitem ${state.selectedPath === path ? 'selected' : ''}" data-path="${path}">
                        <i data-lucide="building" class="subitem-icon"></i>
                        <span class="subitem-name" title="${cName}">${cName}</span>
                    </li>`;
                }).join('')}
            </ul>`;

        const header = li.querySelector('.person-header');
        const list = li.querySelector('.person-companies-list');
        const arrow = li.querySelector('.person-arrow');

        header.addEventListener('click', () => {
            const isVisible = list.style.display !== 'none';
            list.style.display = isVisible ? 'none' : 'block';
            arrow.style.transform = isVisible ? 'rotate(0deg)' : 'rotate(90deg)';
            if (isVisible) {
                state.expandedPeople = state.expandedPeople.filter(p => p !== personName);
            } else {
                if (!state.expandedPeople.includes(personName)) state.expandedPeople.push(personName);
            }
        });

        list.querySelectorAll('.person-company-subitem').forEach(subitem => {
            subitem.addEventListener('click', (e) => {
                e.stopPropagation();
                const path = subitem.getAttribute('data-path');
                el.companyList.querySelectorAll('.person-company-subitem').forEach(s => s.classList.remove('selected'));
                subitem.classList.add('selected');
                selectCompany(path);
                showCompanyTooltip(e, path);
            });
        });

        el.companyList.appendChild(li);
    });

    lucide.createIcons();
}

// ------------------------------------------------------------------ #
// Floating Company Tooltip
// ------------------------------------------------------------------ #

function showCompanyTooltip(event, path) {
    const c = state.companies[path];
    if (!c || !c.data) return;
    const tooltip = el.companyTooltip;
    if (!tooltip) return;

    const activeDirs = c.data.directors || [];
    tooltip.innerHTML = `
        <div class="tooltip-header">
            <h4>${c.data.company_name || c.filename}</h4>
            <button class="btn-tooltip-close">&times;</button>
        </div>
        <div class="tooltip-body">
            <div class="tooltip-row"><strong>UEN:</strong> <span>${c.data.reg_no || 'N/A'}</span></div>
            <div class="tooltip-row"><strong>FYE:</strong> <span>${c.data.financial_year_end || 'N/A'}</span></div>
            <div class="tooltip-row"><strong>AGM Date:</strong> <span>${c.data.agm_date || 'N/A'}</span></div>
            <div class="tooltip-row"><strong>Address:</strong> <span>${c.data.address || 'N/A'}</span></div>
            <div class="tooltip-row"><strong>Active Signatories:</strong>
                <ul class="tooltip-dirs">${activeDirs.map(d => `<li>• ${d}</li>`).join('')}</ul>
            </div>
        </div>`;

    tooltip.style.display = 'block';
    const tooltipWidth = 290, tooltipHeight = 220, offset = 12;
    let left = event.clientX + offset;
    if (left + tooltipWidth > window.innerWidth) left = event.clientX - tooltipWidth - offset;
    if (left < offset) left = offset;
    let top = event.clientY + offset;
    if (top + tooltipHeight > window.innerHeight) top = event.clientY - tooltipHeight - offset;
    if (top < offset) top = offset;
    tooltip.style.left = `${left}px`;
    tooltip.style.top = `${top}px`;

    tooltip.querySelector('.btn-tooltip-close').addEventListener('click', (e) => {
        e.stopPropagation();
        tooltip.style.display = 'none';
    });
    const handleDocumentClick = (e) => {
        if (!tooltip.contains(e.target) && !e.target.closest('.person-company-subitem')) {
            tooltip.style.display = 'none';
            document.removeEventListener('click', handleDocumentClick);
        }
    };
    setTimeout(() => document.addEventListener('click', handleDocumentClick), 100);
}

function updateRowUI(path) {
    let li = null;
    el.companyList.querySelectorAll('.company-item').forEach(row => {
        if (row.getAttribute('data-path') === path) li = row;
    });
    if (!li) return;

    const c = state.companies[path];
    const badge = li.querySelector('.status-badge');
    const title = li.querySelector('.company-title');
    const filename = li.querySelector('.company-filename');

    badge.className = `status-badge ${c.status}`;
    li.classList.remove('pending-status','ok-status','partial-status','error-status');
    li.classList.add(`${c.status}-status`);

    const displayName = (c.data && c.data.company_name) ? c.data.company_name : c.filename;
    title.textContent = displayName;
    title.setAttribute('title', displayName);

    let subName = c.filename;
    if (c.status === 'pending') subName = 'Waiting to read...';
    else if (c.status === 'error') subName = `Error: ${(c.raw_errors && c.raw_errors.length > 0) ? c.raw_errors[0] : 'Failed to parse file.'}`;
    else if (c.status === 'partial') {
        const missing = [];
        if (!c.data.company_name) missing.push('Company Name');
        if (!c.data.reg_no) missing.push('UEN');
        if (!c.data.address) missing.push('AGM Address');
        subName = missing.length > 0 ? `Missing: ${missing.join(', ')}` : `Warning: ${(c.raw_errors||[])[0] || 'Missing information'}`;
    }
    filename.textContent = subName;
    filename.setAttribute('title', subName);
}

function updateRowStatusUI(path, statusClass) {
    el.companyList.querySelectorAll('.company-item').forEach(row => {
        if (row.getAttribute('data-path') === path) {
            const badge = row.querySelector('.status-badge');
            if (badge) { badge.className = `status-badge ${statusClass}`; badge.textContent = ''; }
        }
    });
}

function filterCompanyList() {
    if (state.searchMode === 'people') { renderPeopleList(); return; }
    const query = el.searchFilter.value.toLowerCase().trim();
    el.companyList.querySelectorAll('.company-item').forEach(row => {
        const path = row.getAttribute('data-path');
        const c = state.companies[path];
        const displayName = (c && c.data && c.data.company_name) ? c.data.company_name.toLowerCase() : (c ? c.filename.toLowerCase() : '');
        row.style.display = (displayName.includes(query) || (c && c.filename.toLowerCase().includes(query))) ? 'flex' : 'none';
    });
}

function toggleAllSelections(checked) {
    Object.keys(state.companies).forEach(path => {
        const c = state.companies[path];
        if (checked) {
            c.checked = (c.status === 'ok' || c.status === 'partial');
        } else {
            c.checked = false;
        }
    });
    if (state.searchMode !== 'people') {
        el.companyList.querySelectorAll('.company-item').forEach(row => {
            const path = row.getAttribute('data-path');
            const cb = row.querySelector('input[type="checkbox"]');
            if (cb) cb.checked = state.companies[path].checked;
        });
    }
    updateSelectedCounter();
}

function updateSelectedCounter() {
    const successfulItems = Object.values(state.companies).filter(c => c.status === 'ok' || c.status === 'partial');
    const total = successfulItems.length;
    const selected = successfulItems.filter(c => c.checked).length;
    el.companyCount.textContent = `${selected} of ${total} organisations selected`;
    el.btnGenerate.disabled = Object.values(state.companies).filter(c => c.checked).length === 0;
}

// ------------------------------------------------------------------ #
// Company Details
// ------------------------------------------------------------------ #

function selectCompany(path) {
    if (state.selectedPath) saveCompanyEdits();
    state.selectedPath = path;
    el.companyList.querySelectorAll('.company-item, .person-company-subitem').forEach(item => item.classList.remove('selected'));
    if (state.searchMode !== 'people') {
        el.companyList.querySelectorAll('.company-item').forEach(row => {
            if (row.getAttribute('data-path') === path) row.classList.add('selected');
        });
    } else {
        el.companyList.querySelectorAll('.person-company-subitem').forEach(sub => {
            if (sub.getAttribute('data-path') === path) sub.classList.add('selected');
        });
    }
    loadCompanyDetails(path);
}

function loadCompanyDetails(path) {
    const c = state.companies[path];
    if (!c) { showStatus('Error: Selected company data not found in cache.', 'error'); return; }

    el.detailsEmptyView.style.display = 'none';
    el.detailsView.style.display = 'flex';
    const oView = document.getElementById('overview-view');
    const oEmpty = document.getElementById('overview-empty-view');
    if (oView) oView.style.display = 'flex';
    if (oEmpty) oEmpty.style.display = 'none';

    if (!c.data) {
        showStatus(`Reading detailed client info for ${c.filename}...`);
        updateRowStatusUI(path, 'pending-pulse');
        readCompanyFile(path, true);
        return;
    }

    applyDateDefaultRules(c.data);
    initDirectorsChecklist(c.data);

    el.companyName.value = c.data.company_name || '';
    el.regNo.value = c.data.reg_no || '';
    el.address.value = c.data.address || '';
    el.fyEnd.value = c.data.financial_year_end || '';
    el.agmNumber.value = c.data.agm_number || '';

    const oName = document.getElementById('overview-company-name');
    const oReg = document.getElementById('overview-reg-no');
    const oFy = document.getElementById('overview-fy-end');
    const oAddr = document.getElementById('overview-address');
    const oAgmNum = document.getElementById('overview-agm-number');
    const oAgmDate = document.getElementById('overview-agm-date');

    if (oName) oName.textContent = c.data.company_name || '-';
    if (oReg) oReg.textContent = c.data.reg_no || '-';
    if (oFy) oFy.textContent = c.data.financial_year_end || '-';
    if (oAddr) oAddr.textContent = c.data.address || '-';
    const agmNum = parseInt(c.data.agm_number || '1') || 1;
    const ordinalWord = ORDINALS[agmNum] || `${agmNum}th`;
    if (oAgmNum) oAgmNum.textContent = c.data.agm_number ? `${c.data.agm_number} (${ordinalWord} AGM)` : '-';
    if (oAgmDate) oAgmDate.textContent = c.data.agm_date || '-';

    el.agmDate.value = c.data.agm_date || '';
    if (c.data.is_default_agm) {
        el.agmDate.classList.add('is-default-date');
    } else {
        el.agmDate.classList.remove('is-default-date');
    }
    el.agmDatePicker.value = parseFormattedDate(c.data.agm_date);

    updateOrdinalPreview();
    renderDirectorsList(c.data.directors);
    renderOverviewDirectorsList(c.data.directors);
    renderWarnings(c);
    updateLivePreview();
}

function updateOrdinalPreview() {
    const val = parseInt(el.agmNumber.value);
    if (!isNaN(val) && ORDINALS[val]) {
        el.ordinalText.textContent = `${ORDINALS[val]} AGM`;
        el.ordinalBadge.style.display = 'inline-block';
    } else if (!isNaN(val)) {
        el.ordinalText.textContent = `${val}th AGM`;
        el.ordinalBadge.style.display = 'inline-block';
    } else {
        el.ordinalBadge.style.display = 'none';
    }
}

function renderDirectorsList(directors) {
    el.directorsList.innerHTML = '';
    const c = state.companies[state.selectedPath];
    const allDirs = c.data.all_directors || [];

    if (allDirs.length === 0) {
        el.directorsList.innerHTML = `<div class="empty-list-placeholder" style="padding: 12px; border: 1px dashed var(--border-color); border-radius: 8px;"><p>No directors found in this client file.</p></div>`;
        return;
    }

    allDirs.forEach(name => {
        const row = document.createElement('div');
        row.className = 'director-row';
        const isChecked = c.data.selected_directors.includes(name) ? 'checked' : '';
        row.innerHTML = `
            <div class="custom-checkbox" style="margin-right: 12px;">
                <input type="checkbox" class="director-checkbox" data-name="${name}" ${isChecked}>
                <div class="checkbox-visual"><i data-lucide="check"></i></div>
            </div>
            <span class="director-name-span">${name}</span>`;

        row.querySelector('.director-checkbox').addEventListener('change', (e) => {
            const dirName = e.target.getAttribute('data-name');
            if (e.target.checked) {
                if (!c.data.selected_directors.includes(dirName)) c.data.selected_directors.push(dirName);
            } else {
                c.data.selected_directors = c.data.selected_directors.filter(d => d !== dirName);
            }
            syncFormToCache();
            updateLivePreview();
        });

        el.directorsList.appendChild(row);
    });
    lucide.createIcons();
}

function renderOverviewDirectorsList(activeDirectors) {
    const listEl = document.getElementById('overview-directors-list');
    if (!listEl) return;
    listEl.innerHTML = '';

    if (activeDirectors.length === 0) {
        listEl.innerHTML = `<div class="empty-list-placeholder" style="padding: 12px; border: 1px dashed var(--border-color); border-radius: 8px;"><p>No active signatories selected. Open the Form Editor tab to select.</p></div>`;
        return;
    }

    activeDirectors.forEach(name => {
        const div = document.createElement('div');
        div.className = 'overview-director-badge';
        div.innerHTML = `<i data-lucide="user-check"></i><span>${name}</span>`;
        listEl.appendChild(div);
    });
    lucide.createIcons();
}

function renderWarnings(c) {
    el.warningsList.innerHTML = '';
    const warnings = [];
    if (c.raw_errors && c.raw_errors.length > 0) c.raw_errors.forEach(err => warnings.push(err));
    if (!el.companyName.value.trim()) warnings.push('Company Name is empty.');
    if (!el.regNo.value.trim()) warnings.push('Registration UEN number is missing.');
    if (!el.address.value.trim()) warnings.push('AGM Address is missing.');
    if (!el.fyEnd.value.trim()) warnings.push('Financial Year End is empty.');
    if (!el.agmNumber.value.trim()) warnings.push('AGM Number is empty.');
    if (!el.agmDate.value.trim()) warnings.push('AGM Meeting Date is missing.');
    if (c.data.directors.filter(d => d.trim()).length === 0) warnings.push('No directors selected (at least one check signatory required).');

    if (warnings.length === 0) {
        el.warningsBox.style.display = 'none';
    } else {
        el.warningsBox.style.display = 'block';
        warnings.forEach(warn => {
            const li = document.createElement('li');
            li.textContent = warn;
            el.warningsList.appendChild(li);
        });
    }
}

function syncFormToCache() {
    if (!state.selectedPath) return;
    const c = state.companies[state.selectedPath];
    c.data.company_name = el.companyName.value.trim();
    c.data.reg_no = el.regNo.value.trim();
    c.data.address = el.address.value.trim();
    c.data.financial_year_end = el.fyEnd.value.trim();
    c.data.agm_number = el.agmNumber.value.trim();
    c.data.agm_date = el.agmDate.value.trim();

    const currentYear = new Date().getFullYear();
    const defaultAgmVal = `1 June ${currentYear + 1}`;
    c.data.is_default_agm = (c.data.agm_date.trim() === defaultAgmVal);
    c.data.is_custom_agm = !c.data.is_default_agm;
    c.data.directors = (c.data.all_directors || []).filter(name => c.data.selected_directors.includes(name));

    const oName = document.getElementById('overview-company-name');
    const oReg = document.getElementById('overview-reg-no');
    const oFy = document.getElementById('overview-fy-end');
    const oAddr = document.getElementById('overview-address');
    const oAgmNum = document.getElementById('overview-agm-number');
    const oAgmDate = document.getElementById('overview-agm-date');

    if (oName) oName.textContent = c.data.company_name || '-';
    if (oReg) oReg.textContent = c.data.reg_no || '-';
    if (oFy) oFy.textContent = c.data.financial_year_end || '-';
    if (oAddr) oAddr.textContent = c.data.address || '-';
    const agmNum = parseInt(c.data.agm_number || '1') || 1;
    const ordinalWord = ORDINALS[agmNum] || `${agmNum}th`;
    if (oAgmNum) oAgmNum.textContent = c.data.agm_number ? `${c.data.agm_number} (${ordinalWord} AGM)` : '-';
    if (oAgmDate) oAgmDate.textContent = c.data.agm_date || '-';

    renderOverviewDirectorsList(c.data.directors);
    rebuildKnowledgeGraph();
    updateRowUI(state.selectedPath);
    renderWarnings(c);
}

function saveCompanyEdits() { syncFormToCache(); }

async function rereadSelectedCompany() {
    if (!state.selectedPath) return;
    const filename = state.selectedPath;
    showStatus(`Re-reading data from ${state.companies[filename].filename}...`);
    updateRowStatusUI(filename, 'pending-pulse');
    if (state.companies[filename].data) {
        state.companies[filename].data.all_directors = null;
        state.companies[filename].data.selected_directors = null;
    }
    await readCompanyFile(filename, true);
    showStatus('Reloaded client data from sheet. Overrides discarded.', 'success');
}

// ------------------------------------------------------------------ #
// Live Preview
// ------------------------------------------------------------------ #

function updateLivePreview() {
    if (!state.selectedPath) {
        el.previewDocument.innerHTML = `<div class="preview-empty-state"><i data-lucide="file-text"></i><p>Select a company to load resolution document preview.</p></div>`;
        lucide.createIcons();
        return;
    }

    const c = state.companies[state.selectedPath];
    if (!c.data) return;

    const companyName = (c.data.company_name || 'COMPANY NAME').toUpperCase();
    const regNo = c.data.reg_no || 'UEN REGISTRATION NO';
    const address = c.data.address || 'AGM MEETING ADDRESS';
    const yearEnd = c.data.financial_year_end || 'FINANCIAL YEAR END';
    const agmDate = c.data.agm_date || 'AGM MEETING DATE';
    const agmNum = parseInt(c.data.agm_number || '1') || 1;
    const ordinalCap = ORDINALS[agmNum] || `${agmNum}th`;
    const ordinalUpper = ordinalCap.toUpperCase();

    let sigsHtml = '';
    const activeDirs = c.data.directors.filter(d => d.trim());

    if (activeDirs.length === 0) {
        sigsHtml = `<div class="preview-doc-signatures-grid"><div class="preview-doc-signature-block"><div class="preview-doc-signature-line"></div><div class="preview-doc-signature-name">[NO SIGNATORIES SELECTED]</div></div></div>`;
    } else {
        for (let i = 0; i < activeDirs.length; i += 2) {
            const dir1 = activeDirs[i] || '';
            const dir2 = activeDirs[i + 1] || '';
            sigsHtml += `
                <div class="preview-doc-signatures-grid">
                    <div class="preview-doc-signature-block">
                        <div class="preview-doc-signature-line"></div>
                        <div class="preview-doc-signature-name">${dir1}</div>
                    </div>
                    ${dir2 ? `<div class="preview-doc-signature-block"><div class="preview-doc-signature-line"></div><div class="preview-doc-signature-name">${dir2}</div></div>` : '<div></div>'}
                </div>`;
        }
    }

    el.previewDocument.innerHTML = `
        <div class="preview-doc-title">${companyName}</div>
        <div class="preview-doc-divider">...........................................................................................</div>
        <div class="preview-doc-meta">Company Regn No. &nbsp; ${regNo}</div>
        <div class="preview-doc-meta">(Incorporated in the Republic of Singapore)</div>
        <div class="preview-doc-banner">DIRECTORS' RESOLUTION</div>
        <div class="preview-doc-subtitle">In writing pursuant to the authority given by Article 93<br>of the Company's Articles of Association, hereby RESOLVED:-</div>
        <div class="preview-doc-section-title">AUDITED ACCOUNTS/FINANCIAL STATEMENTS</div>
        <div class="preview-doc-text">That the audited accounts/financial statements of the Company for the year ended ${yearEnd} having been examined by the Directors be and are hereby approved and authorised for issue.</div>
        <div class="preview-doc-section-title">DIRECTORS' STATEMENT & AUDITORS' REPORT</div>
        <div class="preview-doc-text">That the Directors' Statement and Report of Auditors of the Company for the year ended ${yearEnd} be and are hereby adopted and approved and that any two Directors be authorised to sign on behalf of the Board the Directors' Statement.</div>
        <div class="preview-doc-section-title">${ordinalUpper} ANNUAL GENERAL MEETING</div>
        <div class="preview-doc-text">That the ${ordinalCap} Annual General Meeting of members of the Company will be held at ${address} on &nbsp; <b>${agmDate}</b>.</div>
        <div class="preview-doc-section-title">AGENDA OF ${ordinalUpper} ANNUAL GENERAL MEETING</div>
        <div class="preview-doc-text">That the Agenda of the ${ordinalCap} Annual General Meeting of the members of the Company be respectively as follows :-</div>
        <div class="preview-doc-agenda-list">
            <div class="preview-doc-agenda-item"><span>(a)</span> <span>To receive and approve the Company's Audited Accounts/financial statements together with the Directors' Statement and Report of Auditors for the year ended ${yearEnd}.</span></div>
            <div class="preview-doc-agenda-item"><span>(b)</span> <span>To approve Auditors' Remuneration.</span></div>
            <div class="preview-doc-agenda-item"><span>(c)</span> <span>To elect Directors.</span></div>
            <div class="preview-doc-agenda-item"><span>(d)</span> <span>To appoint Auditors.</span></div>
            <div class="preview-doc-agenda-item"><span>(e)</span> <span>To transact any other business, if any.</span></div>
        </div>
        <div class="preview-doc-signatures-header">DIRECTORS</div>
        ${sigsHtml}
        <div class="preview-doc-date">Dated this &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; day of</div>`;
}

// ------------------------------------------------------------------ #
// Director History Modal
// ------------------------------------------------------------------ #

async function openRotationHistory() {
    if (!state.selectedPath) return;
    const c = state.companies[state.selectedPath];
    const reg_no = c.data.reg_no;
    if (!reg_no) { showStatus('Director history requires UEN Registration number configured.', 'warn'); return; }

    let fy_year = String(new Date().getFullYear());
    if (c.data.financial_year_end) {
        const match = c.data.financial_year_end.match(/\b(19|20)\d{2}\b/);
        if (match) fy_year = match[0];
    }

    el.historyCompanyName.textContent = c.data.company_name || c.filename;
    el.historyCompanyUen.textContent = reg_no;
    el.historyPlanningYear.textContent = fy_year;
    el.historyModal.style.display = 'flex';
    el.historyLoading.style.display = 'flex';
    el.historyTableContainer.style.display = 'none';
    el.historyTableBody.innerHTML = '';

    try {
        const response = await authFetch(`/api/history?reg_no=${encodeURIComponent(reg_no)}&fy_year=${encodeURIComponent(fy_year)}`);
        const data = await response.json();
        if (data.error) throw new Error(data.error);
        renderHistoryRows(data.history);
    } catch (err) {
        el.historyTableBody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--color-error);padding:20px;">Failed to fetch history: ${err.message}</td></tr>`;
        el.historyTableContainer.style.display = 'block';
    } finally {
        el.historyLoading.style.display = 'none';
    }
}

function renderHistoryRows(records) {
    el.historyTableBody.innerHTML = '';
    if (records.length === 0) {
        el.historyTableBody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--text-muted);padding:30px;">No service history recorded for this UEN yet.<br><span style="font-size:0.8rem;">History populates automatically during resolution generation runs.</span></td></tr>`;
        el.historyTableContainer.style.display = 'block';
        return;
    }
    records.forEach(r => {
        const tr = document.createElement('tr');
        const yearsStr = r.years_served && r.years_served.length > 0 ? r.years_served.join(', ') : 'None';
        const badgeClass = r.rotation_flag ? 'rotate' : 'ok';
        const badgeText = r.rotation_flag ? '⚠ Rotate?' : 'OK';
        tr.innerHTML = `
            <td style="font-weight:600;">${r.name}</td>
            <td>${yearsStr}</td>
            <td style="font-weight:700;">${r.consecutive} year${r.consecutive === 1 ? '' : 's'}</td>
            <td><span class="rotation-flag ${badgeClass}">${badgeText}</span></td>`;
        el.historyTableBody.appendChild(tr);
    });
    el.historyTableContainer.style.display = 'block';
}

// ------------------------------------------------------------------ #
// Generation Pipeline — returns zip download
// ------------------------------------------------------------------ #

async function runGeneration() {
    if (state.selectedPath) saveCompanyEdits();

    const selectedCompanies = Object.values(state.companies).filter(c => c.checked);
    if (selectedCompanies.length === 0) {
        showStatus('No companies selected for document generation.', 'warn');
        return;
    }

    const unloaded = selectedCompanies.filter(c => !c.data);
    if (unloaded.length > 0) {
        showStatus(`Please wait: ${unloaded.length} selected files are still being parsed...`, 'warn');
        return;
    }

    if (selectedCompanies.length > 3) {
        const proceed = confirm(`You are about to generate ${selectedCompanies.length} documents. Do you wish to proceed?`);
        if (!proceed) { showStatus('Document generation cancelled.', 'info'); return; }
    }

    el.btnGenerate.disabled = true;
    el.progressContainer.style.display = 'flex';
    el.progressBar.style.width = '0%';
    el.progressText.textContent = `Preparing ${selectedCompanies.length} documents...`;
    showStatus('Bulk document generation starting...');

    const currentYear = new Date().getFullYear();
    const sampleFye = selectedCompanies[0].data.financial_year_end || '';
    const match = sampleFye.match(/\b(19|20)\d{2}\b/);
    const global_fy_year = match ? match[0] : String(currentYear);

    const payload = {
        companies: selectedCompanies.map(c => ({
            filename: c.filename,
            data: c.data,
        })),
        fy_year: global_fy_year,
    };

    try {
        const token = await window.__clerk.session.getToken();
        const response = await fetch('/api/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`,
            },
            body: JSON.stringify(payload),
        });

        if (!response.ok) {
            const errData = await response.json().catch(() => ({}));
            throw new Error(errData.error || 'Pipeline execution failed.');
        }

        const blob = await response.blob();
        const disposition = response.headers.get('Content-Disposition') || '';
        const nameMatch = disposition.match(/filename="?([^"]+)"?/);
        const zipName = nameMatch ? nameMatch[1] : `AGM_Documents_${global_fy_year}.zip`;

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = zipName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        el.progressBar.style.width = '100%';
        el.progressText.textContent = `Downloaded ${selectedCompanies.length} documents`;
        showStatus(`Success! Generated all ${selectedCompanies.length} documents. Check your Downloads folder.`, 'success');

    } catch (err) {
        showStatus(err.message || 'Failed execution pipeline.', 'error');
    } finally {
        el.btnGenerate.disabled = false;
        setTimeout(() => { el.progressContainer.style.display = 'none'; }, 3000);
    }
}

// ------------------------------------------------------------------ #
// Utilities
// ------------------------------------------------------------------ #

function showStatus(text, type = 'info') {
    if (text.includes('Scanning') || text.includes('Reading') || text.includes('Loading')) {
        el.statusText.textContent = text;
    } else {
        el.statusText.textContent = 'Ready.';
    }

    const statusIcon = document.querySelector('.status-icon');
    if (statusIcon) {
        if (type === 'error') statusIcon.style.color = 'var(--color-error)';
        else if (type === 'warn') statusIcon.style.color = 'var(--color-warn)';
        else statusIcon.style.color = 'var(--accent-primary)';
    }

    if (text !== 'Ready.' && text !== 'Ready') showToast(text, type);
}

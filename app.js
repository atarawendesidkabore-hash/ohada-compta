// OHADA Compta — Application principale
// SYSCOHADA Revise & SYCEBNL

let currentTab = "dashboard";
let currentRef = "syscohada";
let sycebnlType = "associations"; // 'associations' or 'projets'
let journalEntries = [];
let searchTerm = "";
let filterClass = null;
let currentCompanyDetails = {};
const COMPANY_PROFILE_FIELDS = [
  "raisonSociale",
  "accountingSystem",
  "formeJuridique",
  "sigleUsuel",
  "rccm",
  "nif",
  "nes",
  "siegeSocial",
  "activitePrincipale",
  "capitalSocial",
  "regimeFiscal",
  "pays",
  "exerciceDu",
  "exerciceAu",
  "tel",
  "emailCompta",
  "expertComptable",
  "commissaireAuxComptes"
];
const MONTE_CARLO_FIELD_IDS = [
  "baseRevenue",
  "baseFixedCosts",
  "baseDepreciation",
  "baseCapex",
  "taxRate",
  "iterations",
  "revenueGrowthMin",
  "revenueGrowthMode",
  "revenueGrowthMax",
  "grossMarginMin",
  "grossMarginMode",
  "grossMarginMax",
  "fixedCostMin",
  "fixedCostMode",
  "fixedCostMax",
  "workingCapitalMin",
  "workingCapitalMode",
  "workingCapitalMax",
  "capexMin",
  "capexMode",
  "capexMax"
];
const MONTE_CARLO_PERCENT_FIELDS = new Set([
  "taxRate",
  "revenueGrowthMin",
  "revenueGrowthMode",
  "revenueGrowthMax",
  "grossMarginMin",
  "grossMarginMode",
  "grossMarginMax",
  "workingCapitalMin",
  "workingCapitalMode",
  "workingCapitalMax"
]);
const COST_REDUCTION_STATUS_OPTIONS = [
  "A etudier",
  "Priorite 30 jours",
  "En cours",
  "Validee",
  "Appliquee"
];
const COST_REDUCTION_PILLARS = [
  {
    key: "combinaison",
    title: "Combinaison",
    description: "Mutualiser les achats, regrouper les besoins et profiter des effets de volume."
  },
  {
    key: "adaptation",
    title: "Adaptation",
    description: "Repenser les specifications, le niveau de service et la facon d'executer."
  },
  {
    key: "elimination",
    title: "Elimination",
    description: "Supprimer les depenses sans impact direct sur la proposition de valeur."
  },
  {
    key: "substitution",
    title: "Substitution",
    description: "Remplacer une ressource, un fournisseur ou un processus par une option plus efficiente."
  },
  {
    key: "reaffectation",
    title: "Reaffectation",
    description: "Reutiliser les equipes, actifs et budgets sous-employes avant tout nouvel achat."
  },
  {
    key: "optimisation",
    title: "Optimisation",
    description: "Mesurer les couts, automatiser, securiser la marge et piloter les actions jusqu'au resultat."
  }
];
const COST_REDUCTION_ERRORS = [
  "Eviter les coupes generales sans analyse detaillee des couts.",
  "Ne pas ralentir l'exploitation en supprimant des depenses critiques.",
  "Ne pas negliger le suivi du cash-flow, des prix et des marges.",
  "Verifier les contrats fournisseurs avant renegociation ou resiliation.",
  "Maintenir les actions dans le temps avec responsables et echeances claires.",
  "Profiter du numerique sans ajouter des abonnements inutiles ou redondants."
];

function hasCompanyProfileData() {
  return COMPANY_PROFILE_FIELDS.some((key) => String(currentCompanyDetails[key] || "").trim() !== "");
}

function isCompanyProfileComplete() {
  return ["raisonSociale", "accountingSystem", "formeJuridique", "nif", "siegeSocial", "pays", "exerciceDu", "exerciceAu"]
    .every((key) => String(currentCompanyDetails[key] || "").trim() !== "");
}

// ═══════════════════════════════════════════════════════════
// ACCOUNT MANAGEMENT (multi-company, localStorage)
// ═══════════════════════════════════════════════════════════
const ACCOUNTS_KEY = 'ohada_accounts';
const SESSION_KEY  = 'ohada_session';
const DATA_PREFIX  = 'ohada_data_';
const BF_LIASSE_PACKET_NAME = "BFA Liasse Fiscale Sys Normal SYSCOHADA Révisé DGI-BF[254]";
const BF_LIASSE_DOWNLOAD_NAME = `${BF_LIASSE_PACKET_NAME}.xlsx`;
const EXACT_LIASSE_TEMPLATE_PATH = "LIASSE.xlsx";
const EXACT_LIASSE_DOWNLOAD_NAME = "LIASSE.xlsx";
const SYCEBNL_ASSOCIATIONS_PACKET_NAME = "ETATS FINANCIERS DES ASSOCIATIONS, ORDRES PROFESSIONNELS, FONDATIONS ET ASSIMILES (SYCEBNL)";
const SYCEBNL_ASSOCIATIONS_DOWNLOAD_NAME = `${SYCEBNL_ASSOCIATIONS_PACKET_NAME}.xlsx`;
const SYCEBNL_ASSOCIATIONS_TEMPLATE_PATH = "SYCEBNL_ASSOCIATIONS.xlsx";
const SYCEBNL_PROJETS_PACKET_NAME = "ETATS FINANCIERS DES PROJETS DE DEVELOPPEMENT ET ASSIMILES (SYCEBNL)";
const SYCEBNL_PROJETS_DOWNLOAD_NAME = `${SYCEBNL_PROJETS_PACKET_NAME}.xlsx`;
const SYCEBNL_PROJETS_TEMPLATE_PATH = "SYCEBNL_PROJETS.xlsx";
const SYCEBNL_BALANCE_START_ROW = 3;
const SYCEBNL_BALANCE_END_ROW = 2201;
const EXACT_FORECAST_TEMPLATE_PATH = "PLAN_FINANCIER_PREVISIONNEL.xlsx";
const EXACT_FORECAST_DOWNLOAD_NAME = "Modele-Excel-plan-financier-previsionnel-entreprise.xlsx";
const EXACT_FORECAST_INPUT_SHEET_NAME = "Donn\u00e9es \u00e0 saisir";
const EXACT_FORECAST_PRINT_SHEET_NAME = "Plan financier \u00e0 imprimer";
const OHADA_SITE_HOME_URL = "https://www.ohada.com/";
const OHADA_SITE_NEWS_URL = "https://www.ohada.com/actualite.html";
const OHADA_WATCH_STATE_KEY = "ohada_watch_state";
let currentCompanyId = null;
let ohadaWatchTickStarted = false;

function loadOhadaWatchState() {
  try {
    const raw = localStorage.getItem(OHADA_WATCH_STATE_KEY);
    if (!raw) return { lastRefreshDay: "", lastRefreshAt: "", frameUrl: "" };
    const parsed = JSON.parse(raw);
    return {
      lastRefreshDay: String(parsed.lastRefreshDay || ""),
      lastRefreshAt: String(parsed.lastRefreshAt || ""),
      frameUrl: String(parsed.frameUrl || "")
    };
  } catch (e) {
    return { lastRefreshDay: "", lastRefreshAt: "", frameUrl: "" };
  }
}

let ohadaWatchState = loadOhadaWatchState();

function simpleHash(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);
  return (h >>> 0).toString(16);
}

function getAccounts() {
  try { return JSON.parse(localStorage.getItem(ACCOUNTS_KEY) || '[]'); }
  catch(e) { return []; }
}
function saveAccounts(arr) { localStorage.setItem(ACCOUNTS_KEY, JSON.stringify(arr)); }

function normalizeReferential(ref) {
  return ref === "sycebnl" ? "sycebnl" : "syscohada";
}

function normalizeSycebnlEntityType(type) {
  return type === "projets" ? "projets" : "associations";
}

function getReferentialLabel(ref = currentRef) {
  return normalizeReferential(ref) === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise";
}

function syncReferentielSelect() {
  const sel = document.getElementById("referentiel-select");
  const lockNote = document.getElementById("referentiel-lock-note");
  if (sel) {
    sel.value = normalizeReferential(currentRef);
    const locked = !!currentCompanyId;
    sel.disabled = locked;
    sel.title = locked
      ? "Le systeme comptable est gere par le setup entreprise et les parametres."
      : "Choisissez le systeme comptable.";
  }
  if (lockNote) lockNote.style.display = currentCompanyId ? "inline-block" : "none";
}

function updateCurrentAccountPreferences(patch = {}) {
  if (!currentCompanyId) return;
  const accounts = getAccounts();
  const index = accounts.findIndex((account) => account.id === currentCompanyId);
  if (index === -1) return;
  accounts[index] = Object.assign({}, accounts[index], patch);
  saveAccounts(accounts);
}

function applyAccountingSystemSelection(ref, options = {}) {
  currentRef = normalizeReferential(ref);
  const nextSycebnlType = normalizeSycebnlEntityType(options.sycebnlType || currentCompanyDetails.sycebnlEntityType || sycebnlType);
  sycebnlType = nextSycebnlType;
  currentCompanyDetails.accountingSystem = currentRef;
  currentCompanyDetails.sycebnlEntityType = nextSycebnlType;
  syncReferentielSelect();
  updateCurrentAccountPreferences({
    preferredRef: currentRef,
    preferredSycebnlType: nextSycebnlType
  });
  if (options.save !== false) saveCompanyData();
}

function syncRegisterAccountingSystemUI() {
  const select = document.getElementById("reg-accountingSystem");
  const wrap = document.getElementById("reg-sycebnl-wrap");
  if (!select || !wrap) return;
  wrap.style.display = normalizeReferential(select.value) === "sycebnl" ? "" : "none";
}

function syncParamAccountingSystemUI() {
  const select = document.getElementById("p-accountingSystem");
  const wrap = document.getElementById("p-sycebnl-wrap");
  if (!select || !wrap) return;
  wrap.style.display = normalizeReferential(select.value) === "sycebnl" ? "" : "none";
}

function bindStaticAuthFormEvents() {
  const regSystem = document.getElementById("reg-accountingSystem");
  if (regSystem && !regSystem.dataset.bound) {
    regSystem.addEventListener("change", syncRegisterAccountingSystemUI);
    regSystem.dataset.bound = "1";
  }
  syncRegisterAccountingSystemUI();
}

function saveCompanyData() {
  if (!currentCompanyId) return;
  currentCompanyDetails.accountingSystem = normalizeReferential(currentCompanyDetails.accountingSystem || currentRef);
  currentCompanyDetails.sycebnlEntityType = normalizeSycebnlEntityType(currentCompanyDetails.sycebnlEntityType || sycebnlType);
  const data = {
    journalEntries,
    openingBalances: Object.assign({}, OPENING_BALANCES),
    currentRef,
    sycebnlType,
    companyDetails: currentCompanyDetails
  };
  localStorage.setItem(DATA_PREFIX + currentCompanyId, JSON.stringify(data));
}

function loadCompanyData(id) {
  const raw = localStorage.getItem(DATA_PREFIX + id);
  if (!raw) return; // fresh company, keep defaults
  try {
    const data = JSON.parse(raw);
    if (data.journalEntries) journalEntries = data.journalEntries;
    if (data.openingBalances) {
      // Clear and repopulate OPENING_BALANCES (it is const but mutable)
      Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);
      Object.assign(OPENING_BALANCES, data.openingBalances);
    }
    if (data.currentRef) currentRef = data.currentRef;
    if (data.sycebnlType) sycebnlType = data.sycebnlType;
    if (data.companyDetails) currentCompanyDetails = data.companyDetails;
    if (!currentCompanyDetails.accountingSystem) currentCompanyDetails.accountingSystem = normalizeReferential(data.currentRef || currentRef);
    if (!currentCompanyDetails.sycebnlEntityType) currentCompanyDetails.sycebnlEntityType = normalizeSycebnlEntityType(data.sycebnlType || sycebnlType);
    currentRef = normalizeReferential(currentCompanyDetails.accountingSystem || currentRef);
    sycebnlType = normalizeSycebnlEntityType(currentCompanyDetails.sycebnlEntityType || sycebnlType);
    syncReferentielSelect();
  } catch(e) { console.warn('loadCompanyData error', e); }
}

function loginCompany(id) {
  currentCompanyId = id;
  localStorage.setItem(SESSION_KEY, id);
  const accounts = getAccounts();
  const acct = accounts.find(a => a.id === id);
  // Reset to blank state — new companies start empty, returning ones get their data loaded
  journalEntries = [];
  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);
  currentRef = normalizeReferential(acct && acct.preferredRef ? acct.preferredRef : 'syscohada');
  sycebnlType = normalizeSycebnlEntityType(acct && acct.preferredSycebnlType ? acct.preferredSycebnlType : 'associations');
  currentCompanyDetails = {
    accountingSystem: currentRef,
    sycebnlEntityType: sycebnlType
  };
  syncReferentielSelect();
  loadCompanyData(id);
  applyAccountingSystemSelection(currentCompanyDetails.accountingSystem || currentRef, {
    sycebnlType: currentCompanyDetails.sycebnlEntityType || sycebnlType,
    save: false
  });
  // Update topbar
  const badge = document.getElementById('company-name-display');
  const logoutBtn = document.getElementById('btn-logout');
  if (badge) { badge.textContent = acct ? acct.company : 'Mon compte'; badge.style.display = 'inline-block'; }
  if (logoutBtn) logoutBtn.style.display = 'inline-block';
  // Hide auth overlay
  const overlay = document.getElementById('auth-overlay');
  if (overlay) overlay.classList.remove('active');
  render();
}

function logoutCompany() {
  saveCompanyData();
  currentCompanyId = null;
  localStorage.removeItem(SESSION_KEY);
  // Reset app state to defaults
  journalEntries = [];
  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);
  currentRef = 'syscohada';
  sycebnlType = 'associations';
  currentCompanyDetails = {};
  currentTab = 'dashboard';
  searchTerm = "";
  filterClass = null;
  syncReferentielSelect();
  // Hide topbar elements
  const badge = document.getElementById('company-name-display');
  const logoutBtn = document.getElementById('btn-logout');
  if (badge) badge.style.display = 'none';
  if (logoutBtn) logoutBtn.style.display = 'none';
  // Show auth overlay
  showAuthOverlay('login');
}

function showAuthOverlay(tab) {
  const overlay = document.getElementById('auth-overlay');
  if (overlay) overlay.classList.add('active');
  showAuthTab(tab || 'login');
}

function showAuthTab(tab) {
  document.getElementById('form-login').style.display = tab === 'login' ? '' : 'none';
  document.getElementById('form-register').style.display = tab === 'register' ? '' : 'none';
  document.getElementById('tab-login').classList.toggle('active', tab === 'login');
  document.getElementById('tab-register').classList.toggle('active', tab === 'register');
  bindStaticAuthFormEvents();
}

function handleLogin() {
  const email = (document.getElementById('login-email').value || '').trim().toLowerCase();
  const pass  = document.getElementById('login-pass').value || '';
  const msg   = document.getElementById('login-msg');
  msg.className = 'auth-msg';
  if (!email || !pass) { msg.className = 'auth-msg error'; msg.textContent = 'Veuillez remplir tous les champs.'; return; }
  const accounts = getAccounts();
  const acct = accounts.find(a => a.email === email && a.passHash === simpleHash(pass));
  if (!acct) { msg.className = 'auth-msg error'; msg.textContent = 'Email ou mot de passe incorrect.'; return; }
  loginCompany(acct.id);
}

function handleRegister() {
  const company = (document.getElementById('reg-company').value || '').trim();
  const accountingSystem = normalizeReferential((document.getElementById('reg-accountingSystem').value || 'syscohada').trim());
  const regSycebnlType = normalizeSycebnlEntityType((document.getElementById('reg-sycebnlType').value || 'associations').trim());
  const email   = (document.getElementById('reg-email').value || '').trim().toLowerCase();
  const pass    = document.getElementById('reg-pass').value || '';
  const pass2   = document.getElementById('reg-pass2').value || '';
  const msg     = document.getElementById('reg-msg');
  msg.className = 'auth-msg';
  if (!company || !email || !pass) { msg.className = 'auth-msg error'; msg.textContent = 'Tous les champs sont obligatoires.'; return; }
  if (pass !== pass2) { msg.className = 'auth-msg error'; msg.textContent = 'Les mots de passe ne correspondent pas.'; return; }
  if (pass.length < 6) { msg.className = 'auth-msg error'; msg.textContent = 'Le mot de passe doit contenir au moins 6 caracteres.'; return; }
  const accounts = getAccounts();
  if (accounts.find(a => a.email === email)) { msg.className = 'auth-msg error'; msg.textContent = 'Cet email est deja utilise.'; return; }
  const id = 'c_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  accounts.push({
    id,
    company,
    email,
    passHash: simpleHash(pass),
    preferredRef: accountingSystem,
    preferredSycebnlType: regSycebnlType,
    createdAt: new Date().toISOString()
  });
  saveAccounts(accounts);
  msg.className = 'auth-msg success';
  msg.textContent = 'Compte cree ! Connexion en cours...';
  setTimeout(() => loginCompany(id), 600);
}

function checkAuth() {
  const sessionId = localStorage.getItem(SESSION_KEY);
  if (sessionId) {
    const accounts = getAccounts();
    if (accounts.find(a => a.id === sessionId)) {
      loginCompany(sessionId);
      return;
    }
  }
  // No valid session — show auth overlay
  showAuthOverlay('login');
}

// Clock
setInterval(() => {
  const el = document.getElementById("clock");
  if (el) el.textContent = new Date().toLocaleTimeString("fr-FR", { hour12: false });
}, 1000);

// Referentiel switch
document.getElementById("referentiel-select").addEventListener("change", (e) => {
  applyAccountingSystemSelection(e.target.value);
  render();
});

// SYCEBNL entity type switch (injected dynamically when SYCEBNL is active)
function setSycebnlType(type) {
  sycebnlType = normalizeSycebnlEntityType(type);
  currentCompanyDetails.sycebnlEntityType = sycebnlType;
  updateCurrentAccountPreferences({ preferredSycebnlType: sycebnlType });
  saveCompanyData();
  render();
}

// Mobile menu
const mobileBtn = document.getElementById("mobile-menu-btn");
const mobileOverlay = document.getElementById("mobile-overlay");
const sidebar = document.getElementById("sidebar");

function openMobileMenu() {
  sidebar.classList.add("mobile-open");
  mobileOverlay.classList.add("open");
}
function closeMobileMenu() {
  sidebar.classList.remove("mobile-open");
  mobileOverlay.classList.remove("open");
}
if (mobileBtn) mobileBtn.addEventListener("click", openMobileMenu);
if (mobileOverlay) mobileOverlay.addEventListener("click", closeMobileMenu);

function getTabFromLocationHash() {
  if (typeof window === "undefined") return "";
  const hash = String(window.location.hash || "").replace(/^#/, "").trim();
  if (!hash) return "";
  return document.querySelector(`.nav-btn[data-tab="${hash}"]`) ? hash : "";
}

function syncTabFromLocationHash() {
  const tabFromHash = getTabFromLocationHash();
  if (tabFromHash) currentTab = tabFromHash;
}

function updateLocationHashForTab(tab) {
  if (typeof window === "undefined" || typeof history === "undefined" || typeof history.replaceState !== "function") return;
  const nextUrl = tab === "dashboard"
    ? `${window.location.pathname}${window.location.search}`
    : `${window.location.pathname}${window.location.search}#${tab}`;
  history.replaceState(null, "", nextUrl);
}

function navigateToTab(tab) {
  const target = document.querySelector(`.nav-btn[data-tab="${tab}"]`);
  if (!target) return;
  document.querySelectorAll(".nav-btn").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tab));
  currentTab = tab;
  updateLocationHashForTab(tab);
  closeMobileMenu();
  render();
  window.scrollTo(0, 0);
}

// Navigation
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => navigateToTab(btn.dataset.tab));
});

if (typeof window !== "undefined") {
  window.addEventListener("hashchange", () => {
    const tabFromHash = getTabFromLocationHash();
    const nextTab = tabFromHash || "dashboard";
    if (nextTab === currentTab) return;
    currentTab = nextTab;
    document.querySelectorAll(".nav-btn").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === nextTab));
    render();
    window.scrollTo(0, 0);
  });
}

function getPlan() {
  if (currentRef === "sycebnl") {
    const base = PLAN_COMPTABLE_SYSCOHADA.filter(a => a.numero !== "13" && a.numero !== "131" && a.numero !== "139");
    return [...base, ...PLAN_COMPTABLE_SYCEBNL_ADDITIONS, ...PLAN_COMPTABLE_SYCEBNL_CLASSE9].sort((a, b) => a.numero.localeCompare(b.numero));
  }
  return PLAN_COMPTABLE_SYSCOHADA;
}

function fmt(n) { return n.toLocaleString("fr-FR"); }

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function fmtPercent(value, digits = 1) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "—";
  return `${(num * 100).toFixed(digits)}%`;
}

function percentile(sortedValues, ratio) {
  if (!Array.isArray(sortedValues) || !sortedValues.length) return 0;
  const bounded = clampNumber(ratio, 0, 1, 0.5);
  const position = (sortedValues.length - 1) * bounded;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sortedValues[lower];
  const weight = position - lower;
  return sortedValues[lower] + ((sortedValues[upper] - sortedValues[lower]) * weight);
}

function average(values) {
  if (!Array.isArray(values) || !values.length) return 0;
  return values.reduce((sum, value) => sum + (Number(value) || 0), 0) / values.length;
}

function triangularSample(minValue, modeValue, maxValue) {
  let min = Number(minValue);
  let mode = Number(modeValue);
  let max = Number(maxValue);
  if (!Number.isFinite(min) || !Number.isFinite(mode) || !Number.isFinite(max)) return 0;
  if (min > max) [min, max] = [max, min];
  mode = Math.min(max, Math.max(min, mode));
  if (Math.abs(max - min) < 0.000001) return min;
  const pivot = (mode - min) / (max - min);
  const u = Math.random();
  if (u <= pivot) return min + Math.sqrt(u * (max - min) * (mode - min));
  return max - Math.sqrt((1 - u) * (max - min) * (max - mode));
}

function normalizeTriangularInputs(minValue, modeValue, maxValue, fallbackValues) {
  const defaults = fallbackValues || { min: 0, mode: 0, max: 0 };
  let min = Number.isFinite(Number(minValue)) ? Number(minValue) : Number(defaults.min);
  let mode = Number.isFinite(Number(modeValue)) ? Number(modeValue) : Number(defaults.mode);
  let max = Number.isFinite(Number(maxValue)) ? Number(maxValue) : Number(defaults.max);
  if (min > max) [min, max] = [max, min];
  mode = Math.min(max, Math.max(min, mode));
  return { min, mode, max };
}

function buildHistogram(values, binCount = 8) {
  if (!Array.isArray(values) || !values.length) return [];
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (!Number.isFinite(min) || !Number.isFinite(max)) return [];
  if (Math.abs(max - min) < 0.000001) {
    return [{
      label: `${fmt(Math.round(min))} XOF`,
      count: values.length,
      widthPct: 100
    }];
  }

  const safeBinCount = Math.max(4, Math.min(12, Math.round(binCount)));
  const step = (max - min) / safeBinCount;
  const bins = Array.from({ length: safeBinCount }, (_, index) => ({
    start: min + (index * step),
    end: index === safeBinCount - 1 ? max : min + ((index + 1) * step),
    count: 0
  }));

  values.forEach((value) => {
    const rawIndex = step > 0 ? Math.floor((value - min) / step) : 0;
    const targetIndex = Math.max(0, Math.min(safeBinCount - 1, rawIndex));
    bins[targetIndex].count += 1;
  });

  const peak = Math.max(...bins.map((bin) => bin.count), 1);
  return bins.map((bin) => ({
    label: `${fmt(Math.round(bin.start))} - ${fmt(Math.round(bin.end))}`,
    count: bin.count,
    widthPct: (bin.count / peak) * 100
  }));
}

function formatDateValue(value) {
  if (!value) return "";
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? String(value) : parsed.toLocaleDateString("fr-FR");
}

function formatDateTimeValue(value) {
  if (!value) return "";
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? String(value) : parsed.toLocaleString("fr-FR");
}

function parseIsoDate(value) {
  if (!value) return null;
  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatIsoDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function addYearsToIsoDate(value, years = 1) {
  const parsed = parseIsoDate(value);
  if (!parsed) return "";
  const shifted = new Date(parsed);
  shifted.setFullYear(shifted.getFullYear() + years);
  return formatIsoDate(shifted);
}

function saveOhadaWatchState() {
  localStorage.setItem(OHADA_WATCH_STATE_KEY, JSON.stringify(ohadaWatchState));
}

function buildOhadaWatchFrameUrl() {
  const separator = OHADA_SITE_HOME_URL.includes("?") ? "&" : "?";
  return `${OHADA_SITE_HOME_URL}${separator}source=ohada-compta&t=${Date.now()}`;
}

function ensureOhadaWatchDailyRefresh(force = false) {
  const todayKey = new Date().toISOString().slice(0, 10);
  if (force || ohadaWatchState.lastRefreshDay !== todayKey || !ohadaWatchState.frameUrl) {
    ohadaWatchState = {
      lastRefreshDay: todayKey,
      lastRefreshAt: new Date().toISOString(),
      frameUrl: buildOhadaWatchFrameUrl()
    };
    saveOhadaWatchState();
  }
  return ohadaWatchState;
}

function startOhadaWatchTicker() {
  if (ohadaWatchTickStarted) return;
  ohadaWatchTickStarted = true;
  window.setInterval(() => {
    if (currentTab !== "veille") return;
    const previousDay = ohadaWatchState.lastRefreshDay;
    ensureOhadaWatchDailyRefresh(false);
    if (ohadaWatchState.lastRefreshDay !== previousDay) render();
  }, 60000);
}

function refreshOhadaWatch() {
  ensureOhadaWatchDailyRefresh(true);
  if (currentTab === "veille") render();
}

function openOhadaWebsite(url = OHADA_SITE_HOME_URL) {
  window.open(url, "_blank", "noopener,noreferrer");
}

function getCurrentAccountMeta() {
  const accounts = getAccounts();
  return currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
}

function getCompanyDisplayName() {
  const acct = getCurrentAccountMeta();
  return currentCompanyDetails.raisonSociale || (acct ? acct.company : "");
}

function sensSpan(s) {
  if (s === "Debit") return '<span class="sens-debit">Debit</span>';
  if (s === "Credit") return '<span class="sens-credit">Credit</span>';
  return '<span class="sens-variable">Variable</span>';
}

function getClassColor(c) {
  const cl = SYSCOHADA_CLASSES.find(x => x.num === c);
  return cl ? cl.color : "#8A9BB5";
}

// Compute balances from journal (opening balances from OPENING_BALANCES seeded first)
function computeBalances() {
  const balances = {};
  // Seed opening balances (N-1)
  Object.keys(OPENING_BALANCES).forEach(code => {
    const ob = OPENING_BALANCES[code];
    balances[code] = { n1d: ob.n1d, n1c: ob.n1c, debit: ob.n1d, credit: ob.n1c };
  });
  // Add current-year movements
  journalEntries.forEach((e) => {
    if (!balances[e.compte]) balances[e.compte] = { n1d: 0, n1c: 0, debit: 0, credit: 0 };
    balances[e.compte].debit += e.debit;
    balances[e.compte].credit += e.credit;
  });
  return balances;
}

function isBalancedAmount(totalDebit, totalCredit, tolerance = 0.005) {
  return Math.abs((totalDebit || 0) - (totalCredit || 0)) < tolerance;
}

function getOpeningBalanceSummary(balanceMap = OPENING_BALANCES) {
  let totalDebit = 0;
  let totalCredit = 0;
  let count = 0;

  Object.keys(balanceMap).forEach((code) => {
    const row = balanceMap[code] || {};
    const n1d = Number(row.n1d) || 0;
    const n1c = Number(row.n1c) || 0;
    if (n1d !== 0 || n1c !== 0) count++;
    totalDebit += n1d;
    totalCredit += n1c;
  });

  return {
    count,
    totalDebit,
    totalCredit,
    gap: totalDebit - totalCredit,
    isBalanced: isBalancedAmount(totalDebit, totalCredit)
  };
}

function getComputedBalanceSummary(bal = computeBalances()) {
  let totalDebit = 0;
  let totalCredit = 0;

  Object.keys(bal).forEach((code) => {
    const row = bal[code] || {};
    totalDebit += Number(row.debit) || 0;
    totalCredit += Number(row.credit) || 0;
  });

  return {
    totalDebit,
    totalCredit,
    gap: totalDebit - totalCredit,
    isBalanced: isBalancedAmount(totalDebit, totalCredit)
  };
}

function replaceOpeningBalances(nextBalances) {
  Object.keys(OPENING_BALANCES).forEach((code) => delete OPENING_BALANCES[code]);
  Object.keys(nextBalances).forEach((code) => {
    const row = nextBalances[code] || {};
    OPENING_BALANCES[code] = {
      n1d: Number(row.n1d) || 0,
      n1c: Number(row.n1c) || 0
    };
  });
}

function mergeOpeningNet(target, code, signedNet) {
  if (!code || Math.abs(Number(signedNet) || 0) < 0.005) return;
  const current = ((target[code] && target[code].n1d) || 0) - ((target[code] && target[code].n1c) || 0);
  const next = current + signedNet;
  target[code] = {
    n1d: next > 0 ? next : 0,
    n1c: next < 0 ? Math.abs(next) : 0
  };
}

function findAccountByCode(code, plan = getPlan()) {
  const normalized = String(code || "").trim();
  if (!normalized) return null;

  let fallback = null;
  for (const account of plan) {
    if (account.numero === normalized) return account;
    if (normalized.startsWith(account.numero) && (!fallback || account.numero.length > fallback.numero.length)) {
      fallback = account;
    }
  }
  return fallback;
}

function computeResultatExercice(bal = computeBalances()) {
  let produits = 0;
  let charges = 0;

  Object.keys(bal).forEach((code) => {
    const row = bal[code] || {};
    const mvtD = (row.debit || 0) - (row.n1d || 0);
    const mvtC = (row.credit || 0) - (row.n1c || 0);
    const classe = parseInt(String(code)[0], 10);
    if (classe === 7 || code.startsWith("82") || code.startsWith("84")) produits += mvtC - mvtD;
    if (classe === 6 || code.startsWith("81") || code.startsWith("83") || code.startsWith("85") || code.startsWith("87") || code.startsWith("89")) {
      charges += mvtD - mvtC;
    }
  });

  return produits - charges;
}

function getClotureSnapshot() {
  const bal = computeBalances();
  const plan = getPlan();
  const profileComplete = isCompanyProfileComplete();
  const exerciceDu = currentCompanyDetails.exerciceDu || "";
  const exerciceAu = currentCompanyDetails.exerciceAu || "";
  const exerciseStartDate = parseIsoDate(exerciceDu);
  const exerciseEndDate = parseIsoDate(exerciceAu);
  const validDates = !!(exerciseStartDate && exerciseEndDate && exerciseEndDate >= exerciseStartDate);
  const balanceSummary = getComputedBalanceSummary(bal);
  const hasData = Object.keys(bal).length > 0;
  const carryforward = {};

  Object.keys(bal).forEach((code) => {
    const account = findAccountByCode(code, plan);
    const classe = account && account.classe ? account.classe : parseInt(String(code)[0], 10);
    if (![1, 2, 3, 4, 5].includes(classe)) return;
    const net = (bal[code].debit || 0) - (bal[code].credit || 0);
    mergeOpeningNet(carryforward, code, net);
  });

  const resultatExercice = computeResultatExercice(bal);
  mergeOpeningNet(carryforward, resultatExercice >= 0 ? "121" : "129", -resultatExercice);

  const carryforwardSummary = getOpeningBalanceSummary(carryforward);
  const nextExerciceDu = validDates ? addYearsToIsoDate(exerciceDu, 1) : "";
  const nextExerciceAu = validDates ? addYearsToIsoDate(exerciceAu, 1) : "";

  return {
    bal,
    profileComplete,
    exerciceDu,
    exerciceAu,
    validDates,
    hasData,
    balanceSummary,
    resultatExercice,
    carryforward,
    carryforwardSummary,
    nextExerciceDu,
    nextExerciceAu,
    canClose: profileComplete && validDates && hasData && balanceSummary.isBalanced && carryforwardSummary.isBalanced
  };
}

function getBilanSnapshot() {
  const bal = computeBalances();
  const plan = getPlan();
  const actifRows = BILAN_STRUCTURE.actif.map(section => ({ section: section.section, total: 0 }));
  const passifRows = BILAN_STRUCTURE.passif.map(section => ({ section: section.section, total: 0 }));
  const actifImmobilise = BILAN_STRUCTURE.actif[0].section;
  const actifCirculant = BILAN_STRUCTURE.actif[1].section;
  const tresorerieActif = BILAN_STRUCTURE.actif[2].section;
  const capitauxPropres = BILAN_STRUCTURE.passif[0].section;
  const dettesFinancieres = BILAN_STRUCTURE.passif[1].section;
  const passifCirculant = BILAN_STRUCTURE.passif[2].section;
  const tresoreriePassif = BILAN_STRUCTURE.passif[3].section;

  function addSectionTotal(rows, sectionName, amount) {
    const target = rows.find((row) => row.section === sectionName);
    if (target) target.total += amount;
  }

  Object.keys(bal).forEach((code) => {
    const row = bal[code] || {};
    const net = (row.debit || 0) - (row.credit || 0);
    if (Math.abs(net) < 0.005) return;

    const account = findAccountByCode(code, plan);
    const classe = account && account.classe ? account.classe : parseInt(String(code)[0], 10);
    if (![1, 2, 3, 4, 5].includes(classe)) return;

    if (classe === 1) {
      addSectionTotal(passifRows, String(code).startsWith("16") ? dettesFinancieres : capitauxPropres, -net);
      return;
    }

    if (classe === 2) {
      addSectionTotal(actifRows, actifImmobilise, net);
      return;
    }

    if (classe === 3) {
      addSectionTotal(actifRows, actifCirculant, net);
      return;
    }

    if (classe === 4) {
      let goesToActif = net >= 0;
      if (account && account.sens === "Credit") goesToActif = net > 0;
      if (account && account.sens === "Debit") goesToActif = net >= 0;

      addSectionTotal(
        goesToActif ? actifRows : passifRows,
        goesToActif ? actifCirculant : passifCirculant,
        goesToActif ? net : -net
      );
      return;
    }

    const naturalPassiveTreasury = String(code).startsWith("56") || (account && account.sens === "Credit");
    const goesToActif = naturalPassiveTreasury ? net > 0 : net >= 0;
    addSectionTotal(
      goesToActif ? actifRows : passifRows,
      goesToActif ? tresorerieActif : tresoreriePassif,
      goesToActif ? net : -net
    );
  });

  const resultatExercice = computeResultatExercice(bal);
  addSectionTotal(passifRows, capitauxPropres, resultatExercice);

  const totalActif = actifRows.reduce((sum, row) => sum + row.total, 0);
  const totalPassif = passifRows.reduce((sum, row) => sum + row.total, 0);
  const balanceSummary = getComputedBalanceSummary(bal);

  return {
    bal,
    actifRows,
    passifRows,
    resultatExercice,
    totalActif,
    totalPassif,
    balanceSummary,
    isBalanced: balanceSummary.isBalanced && isBalancedAmount(totalActif, totalPassif)
  };
}

function getExactFiscalTemplateMeta() {
  if (currentRef === "sycebnl") {
    if (sycebnlType === "projets") {
      return {
        referential: "sycebnl",
        entityType: "projets",
        packetName: SYCEBNL_PROJETS_PACKET_NAME,
        downloadName: SYCEBNL_PROJETS_DOWNLOAD_NAME,
        templatePath: SYCEBNL_PROJETS_TEMPLATE_PATH,
        buttonLabel: "liasse projets SYCEBNL"
      };
    }
    return {
      referential: "sycebnl",
      entityType: "associations",
      packetName: SYCEBNL_ASSOCIATIONS_PACKET_NAME,
      downloadName: SYCEBNL_ASSOCIATIONS_DOWNLOAD_NAME,
      templatePath: SYCEBNL_ASSOCIATIONS_TEMPLATE_PATH,
      buttonLabel: "liasse ONG / associations SYCEBNL"
    };
  }

  return {
    referential: "syscohada",
    entityType: "syscohada",
    packetName: BF_LIASSE_PACKET_NAME,
    downloadName: EXACT_LIASSE_DOWNLOAD_NAME,
    templatePath: EXACT_LIASSE_TEMPLATE_PATH,
    buttonLabel: "LIASSE.xlsx"
  };
}

function getDsfPacketName() {
  return getExactFiscalTemplateMeta().packetName;
}

function getDsfStatusRows() {
  const bal = computeBalances();
  const hasBalances = Object.keys(bal).length > 0;
  const hasJournal = journalEntries.length > 0;
  const hasImmos = Object.keys(bal).some((code) => code.startsWith("2") || code.startsWith("28"));
  const hasTiers = Object.keys(bal).some((code) => code.startsWith("4"));
  const profileComplete = isCompanyProfileComplete();

  if (currentRef === "sycebnl") {
    const isProject = sycebnlType === "projets";
    return [
      { code: "SYC-01", label: "Informations generales de l'entite", status: profileComplete ? "Pret" : "En attente", hint: profileComplete ? "L'identification peut etre injectee dans le modele officiel." : "Completez la fiche entreprise avant export." },
      { code: "SYC-02", label: "Balance N et N-1", status: hasBalances ? "Pret" : "En attente", hint: hasBalances ? "Des soldes sont disponibles pour le classeur exact." : "Importez une balance ou des ecritures." },
      { code: "SYC-03", label: "Bilan Actif / Passif", status: hasBalances ? "Pret" : "En attente", hint: hasBalances ? "Les etats de situation peuvent etre consolides." : "Aucune balance disponible." },
      { code: "SYC-04", label: "Compte d'exploitation", status: hasJournal ? "Pret" : hasBalances ? "A completer" : "En attente", hint: hasJournal ? "Les mouvements de gestion sont presents." : "Ajoutez les ecritures de l'exercice." },
      { code: "SYC-05", label: isProject ? "TER / TRC / TEB" : "TFT", status: hasJournal ? "A completer" : "En attente", hint: isProject ? "Les tableaux ressources / emplois restent a valider." : "Le tableau de flux doit etre revu avant depot." },
      { code: "SYC-06", label: "Notes annexes", status: hasBalances ? "A completer" : "En attente", hint: "Le modele officiel contient deja les onglets de notes a documenter." },
      { code: "SYC-07", label: "Informations fiscales et teledeclaration", status: profileComplete ? "A completer" : "En attente", hint: profileComplete ? "Le NES, le regime et le pays peuvent etre injectes." : "Renseignez NIF, pays, regime fiscal et NES." },
      { code: "SYC-08", label: isProject ? "Etats financiers des projets" : "Etats financiers des associations / ONG", status: hasBalances ? "Pret" : "En attente", hint: isProject ? "Le modele projets SYCEBNL sera utilise." : "Le modele ONG / associations SYCEBNL sera utilise." }
    ];
  }

  return [
    { code: "DSF-01", label: "Bilan — Systeme normal", status: hasBalances ? "Pret" : "En attente", hint: hasBalances ? "Disponible a partir des soldes charges." : "Chargez une balance ou des ecritures." },
    { code: "DSF-02", label: "Compte de resultat — Systeme normal", status: hasJournal ? "Pret" : hasBalances ? "A completer" : "En attente", hint: hasJournal ? "Les mouvements de l'exercice sont disponibles." : "Ajoutez les ecritures de l'exercice." },
    { code: "DSF-03", label: "Tableau de flux de tresorerie", status: hasJournal ? "A completer" : "En attente", hint: "Necessite les flux de l'exercice et la revue de cloture." },
    { code: "DSF-04", label: "Tableau des immobilisations", status: hasImmos ? "A completer" : "En attente", hint: hasImmos ? "Des immobilisations ont ete detectees." : "Aucune immobilisation detectee." },
    { code: "DSF-05", label: "Tableau des amortissements", status: hasImmos ? "A completer" : "En attente", hint: hasImmos ? "Prevoir les dotations et cumuls." : "Aucune base amortissable detectee." },
    { code: "DSF-06", label: "Tableau des provisions", status: hasJournal ? "A completer" : "En attente", hint: "A completer selon vos ecritures d'inventaire." },
    { code: "DSF-07", label: "Etat des creances et dettes", status: hasTiers ? "A completer" : "En attente", hint: hasTiers ? "Des comptes de tiers sont presents." : "Aucun compte de tiers detecte." },
    { code: "DSF-08", label: "Tableau des resultat et soldes intermediaires", status: hasJournal ? "A completer" : "En attente", hint: "Genere a partir des mouvements de gestion." },
    { code: "DSF-09", label: "Notes annexes", status: hasBalances ? "A completer" : "En attente", hint: "Completez les notes en fonction des etats produits." },
    { code: "DSF-10", label: "Informations complementaires DGI", status: profileComplete ? "A completer" : "En attente", hint: profileComplete ? "La fiche entreprise est prete pour la liasse." : "Renseignez les informations legales et fiscales." },
  ];
}

function getBalanceExportRows() {
  const bal = computeBalances();
  const plan = getPlan();
  const comptes = Object.keys(bal).sort();
  let totN1D = 0, totN1C = 0, totMvtD = 0, totMvtC = 0, totSD = 0, totSC = 0;

  const rows = comptes.map(code => {
    const account = plan.find(a => a.numero === code);
    const n1d = bal[code].n1d || 0;
    const n1c = bal[code].n1c || 0;
    const mvtD = bal[code].debit - n1d;
    const mvtC = bal[code].credit - n1c;
    const solde = bal[code].debit - bal[code].credit;
    const soldeD = solde > 0 ? solde : 0;
    const soldeC = solde < 0 ? Math.abs(solde) : 0;
    totN1D += n1d;
    totN1C += n1c;
    totMvtD += mvtD;
    totMvtC += mvtC;
    totSD += soldeD;
    totSC += soldeC;
    return [code, account ? account.libelle : "Inconnu", account ? account.classe : "", n1d, n1c, mvtD, mvtC, soldeD, soldeC];
  });

  rows.unshift(["Compte", "Libelle", "Classe", "S.O. Debit", "S.O. Credit", "Mvt Debit", "Mvt Credit", "Solde Debit", "Solde Credit"]);
  rows.push(["TOTAUX", "", "", totN1D, totN1C, totMvtD, totMvtC, totSD, totSC]);
  return rows;
}

function getBilanExportRows() {
  const { actifRows, passifRows, resultatExercice, totalActif, totalPassif, isBalanced, balanceSummary } = getBilanSnapshot();
  const rows = [["ACTIF", "Montant (XOF)", "", "PASSIF", "Montant (XOF)"]];
  const rowCount = Math.max(actifRows.length, passifRows.length);
  for (let i = 0; i < rowCount; i++) {
    rows.push([
      actifRows[i] ? actifRows[i].section : "",
      actifRows[i] ? Math.abs(actifRows[i].total) : "",
      "",
      passifRows[i] ? passifRows[i].section : "",
      passifRows[i] ? Math.abs(passifRows[i].total) : "",
    ]);
  }

  rows.push(["TOTAL ACTIF", Math.abs(totalActif), "", "TOTAL PASSIF", Math.abs(totalPassif)]);
  rows.push(["RESULTAT EXERCICE", resultatExercice, "", "EQUILIBRE", isBalanced ? "OK" : "ECART"]);
  if (!balanceSummary.isBalanced) {
    rows.push(["BALANCE GENERALE", "", "", "ECART SOURCE", Math.abs(balanceSummary.gap)]);
  }
  return rows;
}

function getResultatExportRows() {
  const bal = computeBalances();
  let totalProduits = 0;
  let totalCharges = 0;

  const rows = RESULTAT_STRUCTURE.map(section => {
    let total = 0;
    Object.keys(bal).forEach(code => {
      if (section.comptes.some(prefix => code.startsWith(prefix))) {
        const mvtD = (bal[code].debit || 0) - (bal[code].n1d || 0);
        const mvtC = (bal[code].credit || 0) - (bal[code].n1c || 0);
        total += section.sens === "credit" ? mvtC - mvtD : mvtD - mvtC;
      }
    });
    if (section.sens === "credit") totalProduits += total;
    else totalCharges += total;
    return [section.section, section.sens === "credit" ? "Produit" : "Charge", Math.abs(total)];
  });

  rows.unshift(["Section", "Type", "Montant (XOF)"]);
  rows.push(["TOTAL PRODUITS", "", totalProduits]);
  rows.push(["TOTAL CHARGES", "", totalCharges]);
  rows.push(["RESULTAT NET", totalProduits - totalCharges >= 0 ? "Benefice" : "Perte", Math.abs(totalProduits - totalCharges)]);
  return rows;
}

function getFluxTresorerieExportRows() {
  const bal = computeBalances();
  let tresorerieOuverture = 0;
  let tresorerieCloture = 0;
  let dotations = 0;
  let resultatNet = 0;

  Object.keys(bal).forEach(code => {
    const openingNet = (bal[code].n1d || 0) - (bal[code].n1c || 0);
    const closingNet = (bal[code].debit || 0) - (bal[code].credit || 0);
    const mvtD = (bal[code].debit || 0) - (bal[code].n1d || 0);
    const mvtC = (bal[code].credit || 0) - (bal[code].n1c || 0);

    if (code.startsWith("5")) {
      tresorerieOuverture += openingNet;
      tresorerieCloture += closingNet;
    }

    if (code.startsWith("68")) dotations += mvtD - mvtC;
    if (code.startsWith("7") || code.startsWith("82") || code.startsWith("84")) resultatNet += mvtC - mvtD;
    if (code.startsWith("6") || code.startsWith("81") || code.startsWith("83") || code.startsWith("85") || code.startsWith("87") || code.startsWith("89")) resultatNet -= (mvtD - mvtC);
  });

  return [
    ["TABLEAU DE FLUX DE TRESORERIE (TFT) — Synthese preparatoire"],
    [""],
    ["Ligne", "Montant (XOF)", "Observation"],
    ["Tresorerie a l'ouverture", tresorerieOuverture, "Solde net des comptes de tresorerie a l'ouverture."],
    ["Tresorerie a la cloture", tresorerieCloture, "Solde net des comptes de tresorerie a la cloture."],
    ["Variation nette de tresorerie", tresorerieCloture - tresorerieOuverture, "Variation brute observee entre l'ouverture et la cloture."],
    ["Resultat net de l'exercice", resultatNet, "Base de travail pour les flux operationnels."],
    ["Dotations aux amortissements", dotations, "Retraitees dans la construction detaillee du TFT."],
    [""],
    ["Note", "", "Le SYSCOHADA revise remplace le TAFIRE par le Tableau de flux de tresorerie. Cette feuille constitue une base preparatoire et doit etre completee avant depot officiel."],
  ];
}

function getAmortissementExportRows() {
  const bal = computeBalances();
  const plan = getPlan();
  const immos = plan.filter(a => a.classe === 2 && a.sens === "Debit" && bal[a.numero]);
  const durees = {"21":5,"211":5,"212":5,"213":10,"214":10,"22":0,"23":20,"231":20,"232":20,"234":10,"235":7,"24":5,"241":5,"242":5,"244":3,"245":5,"246":7,"248":5};

  function amortCode(code) {
    if (code.startsWith("21")) return "281";
    if (code.startsWith("22")) return "282";
    if (code.startsWith("23")) return "283";
    return "284";
  }

  const rows = [["Compte", "Libelle", "VBO", "Duree", "Taux %", "Cumul N-1", "Dotation N", "Cumul N", "VNC"]];
  immos.forEach(account => {
    const vbo = bal[account.numero].debit;
    if (!vbo) return;
    const ac = amortCode(account.numero);
    const cumulN1 = ac && bal[ac] ? (bal[ac].n1c || 0) : 0;
    const duree = durees[account.numero] || 5;
    const taux = duree > 0 ? Math.round(10000 / duree) / 100 : 0;
    const dotation = duree > 0 ? Math.round(vbo / duree) : 0;
    const cumulN = cumulN1 + dotation;
    const vnc = Math.max(0, vbo - cumulN);
    rows.push([account.numero, account.libelle, vbo, duree || "", taux || "", cumulN1, dotation, cumulN, vnc]);
  });
  return rows;
}

function appendWorkbookSheet(workbook, name, rows, widths, mergeAcross) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if (Array.isArray(widths)) ws["!cols"] = widths.map(width => ({ wch: width }));
  if (mergeAcross) ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: mergeAcross } }];
  XLSX.utils.book_append_sheet(workbook, ws, name);
}

function normalizeTemplateKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getCountryCodeForLiasse(country) {
  const key = normalizeTemplateKey(country);
  const map = {
    "benin": "01",
    "burkina": "02",
    "burkina faso": "02",
    "cote d ivoire": "03",
    "cote divoire": "03",
    "guinee bissau": "04",
    "mali": "05",
    "niger": "06",
    "senegal": "07",
    "togo": "08",
    "cameroun": "09",
    "centrafrique": "10",
    "congo": "11",
    "gabon": "12",
    "guinee equatoriale": "13",
    "tchad": "14",
    "comores": "15",
    "guinee": "16",
    "guinee conakry": "16",
  };
  return map[key] || "99";
}

function getLegalFormCodeForLiasse(formeJuridique) {
  const key = normalizeTemplateKey(formeJuridique);
  if (!key) return "09";
  if (key.includes("participation publique") && key.includes("sa")) return "00";
  if (key === "sa" || key.includes("societe anonyme")) return "01";
  if (key.includes("sas")) return "02";
  if (key.includes("sarl") || key.includes("eurl")) return "03";
  if (key.includes("scs")) return "04";
  if (key.includes("snc")) return "05";
  if (key.includes("participation")) return "06";
  if (key.includes("gie")) return "07";
  if (key.includes("association") || key.includes("ong") || key.includes("fondation")) return "08";
  return "09";
}

function getFiscalRegimeCodeForLiasse(regimeFiscal) {
  const key = normalizeTemplateKey(regimeFiscal);
  if (key.includes("normal")) return "1";
  if (key.includes("simplifie")) return "2";
  if (key.includes("micro") || key.includes("cme")) return "3";
  return "4";
}

function toExcelSerial(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return Math.round((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000) + 25569);
}

function getWritableCell(existing) {
  const cell = Object.assign({}, existing || {});
  delete cell.f;
  delete cell.F;
  delete cell.w;
  delete cell.r;
  delete cell.h;
  return cell;
}

function setWorksheetText(sheet, ref, value) {
  if (!sheet || value === undefined || value === null) return;
  const cell = getWritableCell(sheet[ref]);
  sheet[ref] = Object.assign(cell, { t: "s", v: String(value) });
}

function setWorksheetNumber(sheet, ref, value) {
  if (!sheet || value === undefined || value === null || value === "") return;
  const num = Number(value);
  if (Number.isNaN(num)) return;
  const cell = getWritableCell(sheet[ref]);
  sheet[ref] = Object.assign(cell, { t: "n", v: num });
}

function setWorksheetDate(sheet, ref, value) {
  const serial = toExcelSerial(value);
  if (!sheet || serial === null) return;
  const cell = getWritableCell(sheet[ref]);
  sheet[ref] = Object.assign(cell, { t: "n", v: serial, z: cell.z || "dd/mm/yyyy" });
}

function setWorksheetFormulaNumber(sheet, ref, formula, cachedValue) {
  if (!sheet || !formula) return;
  const cell = getWritableCell(sheet[ref]);
  const nextCell = Object.assign(cell, { t: "n", f: formula });
  if (cachedValue !== undefined && cachedValue !== null && cachedValue !== "") {
    const num = Number(cachedValue);
    if (Number.isFinite(num)) nextCell.v = num;
  } else {
    delete nextCell.v;
  }
  sheet[ref] = nextCell;
}

function setDigitSequence(sheet, startRef, rawValue, length) {
  if (!sheet) return;
  const match = /^([A-Z]+)(\d+)$/.exec(startRef);
  if (!match) return;
  const start = XLSX.utils.decode_cell(startRef);
  const digits = String(rawValue || "").replace(/\D/g, "").padStart(length, "0").slice(-length);
  for (let i = 0; i < length; i++) {
    const ref = XLSX.utils.encode_cell({ c: start.c + i, r: start.r });
    setWorksheetNumber(sheet, ref, Number(digits[i]));
  }
}

function guessCityFromAddress(address, country) {
  const parts = String(address || "").split(",").map(part => part.trim()).filter(Boolean);
  if (!parts.length) return "";
  const normalizedCountry = normalizeTemplateKey(country);
  const filtered = parts.filter(part => normalizeTemplateKey(part) !== normalizedCountry);
  return filtered.length > 1 ? filtered[filtered.length - 1] : (filtered[0] || "");
}

function getPhoneCountryCode(phone) {
  const match = String(phone || "").match(/\+?(\d{1,3})/);
  return match ? match[1] : "";
}

function getExerciseDurationMonths(start, end) {
  if (!start || !end) return 12;
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return 12;
  return Math.max(1, ((endDate.getFullYear() - startDate.getFullYear()) * 12) + (endDate.getMonth() - startDate.getMonth()) + 1);
}

function getPreviousExerciseEnd(endDate) {
  if (!endDate) return "";
  const date = new Date(endDate);
  if (Number.isNaN(date.getTime())) return "";
  return new Date(date.getFullYear() - 1, date.getMonth(), date.getDate()).toISOString().slice(0, 10);
}

function populateExactLiasseTemplate(workbook) {
  const acct = getCurrentAccountMeta();
  const companyName = getCompanyDisplayName();
  const countryLabel = (currentCompanyDetails.pays || "Burkina Faso").trim();
  const countryCode = getCountryCodeForLiasse(countryLabel);
  const legalFormCode = getLegalFormCodeForLiasse(currentCompanyDetails.formeJuridique || "");
  const regimeCode = getFiscalRegimeCodeForLiasse(currentCompanyDetails.regimeFiscal || "");
  const exerciseStart = currentCompanyDetails.exerciceDu || "";
  const exerciseEnd = currentCompanyDetails.exerciceAu || "";
  const previousExerciseEnd = getPreviousExerciseEnd(exerciseEnd);
  const durationMonths = getExerciseDurationMonths(exerciseStart, exerciseEnd);
  const city = guessCityFromAddress(currentCompanyDetails.siegeSocial || "", countryLabel);
  const phoneCode = getPhoneCountryCode(currentCompanyDetails.tel || "");
  const companyCountryUpper = countryLabel.toUpperCase();
  const contactLine = [
    currentCompanyDetails.expertComptable || companyName,
    currentCompanyDetails.tel || "",
    currentCompanyDetails.emailCompta || (acct ? acct.email : "")
  ].filter(Boolean).join(" | ");
  const isPublicEntity = normalizeTemplateKey(currentCompanyDetails.formeJuridique || "").includes("public");
  const isForeignControlled = countryCode !== "02";

  const couverture = workbook.Sheets["COUVERTURE"];
  const garde = workbook.Sheets["GARDE"];
  const r1 = workbook.Sheets["FICHE R1"];
  const r2 = workbook.Sheets["FICHE R2"];

  setWorksheetText(couverture, "F4", companyCountryUpper);
  setWorksheetText(garde, "B2", companyCountryUpper);
  setWorksheetText(garde, "D22", companyName);
  setWorksheetText(garde, "C26", currentCompanyDetails.sigleUsuel || "");
  setWorksheetText(garde, "C28", currentCompanyDetails.siegeSocial || "");
  setWorksheetText(garde, "D30", currentCompanyDetails.nif || "");
  setWorksheetText(garde, "D31", currentCompanyDetails.nes || "");
  setWorksheetDate(garde, "E17", exerciseEnd);

  setWorksheetDate(r1, "N8", exerciseStart);
  setWorksheetDate(r1, "S8", exerciseEnd);
  setWorksheetDate(r1, "H10", exerciseEnd);
  setWorksheetDate(r1, "H12", previousExerciseEnd);
  setWorksheetNumber(r1, "U12", durationMonths);
  setWorksheetText(r1, "D14", city);
  setWorksheetText(r1, "K14", currentCompanyDetails.rccm || "");
  setWorksheetText(r1, "C20", currentCompanyDetails.formeJuridique || "Entreprise");
  setWorksheetText(r1, "Q20", currentCompanyDetails.sigleUsuel || "");
  setWorksheetText(r1, "C23", currentCompanyDetails.tel || "");
  setWorksheetText(r1, "E23", currentCompanyDetails.emailCompta || (acct ? acct.email : ""));
  setWorksheetNumber(r1, "K23", phoneCode);
  setWorksheetText(r1, "Q23", city);
  setWorksheetText(r1, "C29", currentCompanyDetails.activitePrincipale || "");
  setWorksheetText(r1, "C32", contactLine);
  setWorksheetText(r1, "C35", contactLine);
  setWorksheetText(r1, "C43", currentCompanyDetails.expertComptable || "");

  setDigitSequence(r2, "Q9", legalFormCode, 2);
  setDigitSequence(r2, "Q11", regimeCode, 1);
  setDigitSequence(r2, "Q13", countryCode, 2);
  setDigitSequence(r2, "Q15", 1, 2);
  setDigitSequence(r2, "Q17", 0, 2);
  setDigitSequence(r2, "Q20", new Date(exerciseStart || exerciseEnd || Date.now()).getFullYear(), 4);
  setWorksheetNumber(r2, "AK9", isPublicEntity ? 1 : 0);
  setWorksheetNumber(r2, "AK11", !isPublicEntity && !isForeignControlled ? 1 : 0);
  setWorksheetNumber(r2, "AK13", isForeignControlled ? 1 : 0);
}

function parsePercentFieldValue(value, mode = "standard") {
  const raw = String(value ?? "").replace(",", ".").trim();
  if (!raw) return null;
  const num = Number(raw);
  if (!Number.isFinite(num)) return null;
  if (mode === "small-percent") {
    if (num > 0.2) return num / 100;
    return num;
  }
  return num > 1 ? num / 100 : num;
}

function normalizeSycebnlTemplateAccountCode(rawCode) {
  return String(rawCode || "").replace(/[^\d]/g, "");
}

function buildSycebnlBalanceSourceRows(balanceMap = computeBalances()) {
  const aggregated = new Map();

  Object.keys(balanceMap).forEach((rawCode) => {
    const code = normalizeSycebnlTemplateAccountCode(rawCode);
    if (!code) return;
    const row = balanceMap[rawCode] || {};
    const target = aggregated.get(code) || { code, n1d: 0, n1c: 0, debit: 0, credit: 0 };
    target.n1d += Number(row.n1d) || 0;
    target.n1c += Number(row.n1c) || 0;
    target.debit += Number(row.debit) || 0;
    target.credit += Number(row.credit) || 0;
    aggregated.set(code, target);
  });

  return Array.from(aggregated.values())
    .filter((row) => Math.abs(row.n1d) + Math.abs(row.n1c) + Math.abs(row.debit) + Math.abs(row.credit) >= 0.005)
    .sort((a, b) => a.code.localeCompare(b.code, "fr", { numeric: true }));
}

function getSycebnlBalanceRowLabel(code, plan = getPlan()) {
  const account = findAccountByCode(code, plan);
  if (account && account.libelle) return account.libelle;

  const journalLabel = journalEntries.find(
    (entry) => normalizeSycebnlTemplateAccountCode(entry.compte) === code && String(entry.libelle || "").trim()
  );
  return journalLabel ? String(journalLabel.libelle).trim() : `Compte ${code}`;
}

function getSycebnlBalanceSheetValues(row, sheetMode) {
  const openingDebit = Math.max(0, Number(row.n1d) || 0);
  const openingCredit = Math.max(0, Number(row.n1c) || 0);
  const movementDebit = Math.max(0, (Number(row.debit) || 0) - openingDebit);
  const movementCredit = Math.max(0, (Number(row.credit) || 0) - openingCredit);

  if (sheetMode === "n-1") {
    const priorClosingNet = openingDebit - openingCredit;
    return {
      previousDebit: 0,
      previousCredit: 0,
      movementDebit: openingDebit,
      movementCredit: openingCredit,
      closingDebit: priorClosingNet > 0 ? priorClosingNet : 0,
      closingCredit: priorClosingNet < 0 ? Math.abs(priorClosingNet) : 0
    };
  }

  const closingNet = (Number(row.debit) || 0) - (Number(row.credit) || 0);
  return {
    previousDebit: openingDebit,
    previousCredit: openingCredit,
    movementDebit,
    movementCredit,
    closingDebit: closingNet > 0 ? closingNet : 0,
    closingCredit: closingNet < 0 ? Math.abs(closingNet) : 0
  };
}

function populateSycebnlBalanceSheet(sheet, rows, sheetMode, plan = getPlan()) {
  if (!sheet) return;
  const capacity = SYCEBNL_BALANCE_END_ROW - SYCEBNL_BALANCE_START_ROW + 1;
  if (rows.length > capacity) {
    throw new Error(`Le modele SYCEBNL accepte au maximum ${capacity} comptes detailles. Veuillez reduire les comptes actifs avant export.`);
  }

  rows.forEach((row, index) => {
    const excelRow = SYCEBNL_BALANCE_START_ROW + index;
    const refs = getSycebnlBalanceSheetValues(row, sheetMode);
    const closingDebitFormula = `MAX((K${excelRow}+M${excelRow})-(L${excelRow}+N${excelRow}),0)`;
    const closingCreditFormula = `MAX((L${excelRow}+N${excelRow})-(K${excelRow}+M${excelRow}),0)`;

    setWorksheetText(sheet, `I${excelRow}`, row.code);
    setWorksheetText(sheet, `J${excelRow}`, getSycebnlBalanceRowLabel(row.code, plan));
    if (refs.previousDebit) setWorksheetNumber(sheet, `K${excelRow}`, refs.previousDebit);
    if (refs.previousCredit) setWorksheetNumber(sheet, `L${excelRow}`, refs.previousCredit);
    if (refs.movementDebit) setWorksheetNumber(sheet, `M${excelRow}`, refs.movementDebit);
    if (refs.movementCredit) setWorksheetNumber(sheet, `N${excelRow}`, refs.movementCredit);
    setWorksheetFormulaNumber(sheet, `O${excelRow}`, closingDebitFormula, refs.closingDebit);
    setWorksheetFormulaNumber(sheet, `P${excelRow}`, closingCreditFormula, refs.closingCredit);
  });
}

function populateSycebnlBalanceWorksheets(workbook) {
  const rows = buildSycebnlBalanceSourceRows();
  const plan = getPlan();
  populateSycebnlBalanceSheet(workbook.Sheets["Balance N"], rows, "n", plan);
  populateSycebnlBalanceSheet(workbook.Sheets["Balance N-1"], rows, "n-1", plan);
}

function populateExactSycebnlTemplate(workbook) {
  const acct = getCurrentAccountMeta();
  const meta = getExactFiscalTemplateMeta();
  const companyName = getCompanyDisplayName();
  const countryLabel = (currentCompanyDetails.pays || "Burkina Faso").trim();
  const countryCode = getCountryCodeForLiasse(countryLabel);
  const legalFormCode = getLegalFormCodeForLiasse(currentCompanyDetails.formeJuridique || "");
  const regimeCode = getFiscalRegimeCodeForLiasse(currentCompanyDetails.regimeFiscal || "");
  const exerciseStart = currentCompanyDetails.exerciceDu || "";
  const exerciseEnd = currentCompanyDetails.exerciceAu || "";
  const previousExerciseEnd = getPreviousExerciseEnd(exerciseEnd);
  const city = guessCityFromAddress(currentCompanyDetails.siegeSocial || "", countryLabel);
  const phoneCode = getPhoneCountryCode(currentCompanyDetails.tel || "");
  const companyCountryUpper = countryLabel.toUpperCase();
  const contactName = currentCompanyDetails.expertComptable || companyName;
  const contactAddress = currentCompanyDetails.siegeSocial || "";
  const legalFormLabel = currentCompanyDetails.formeJuridique || (meta.entityType === "projets" ? "Projet de developpement" : "Association");
  const fiscalRegimeLabel = currentCompanyDetails.regimeFiscal || (meta.entityType === "projets" ? "Exonere" : "Exonere");

  const info = workbook.Sheets["INFORMATIONS GENERALES"];
  const couverture = workbook.Sheets["COUVERTURE"];
  const garde = workbook.Sheets["GARDE"];
  const r1 = workbook.Sheets["FICHE R1"];
  const r2 = workbook.Sheets["FICHE R2"];

  setWorksheetText(couverture, "F4", companyCountryUpper);
  setWorksheetText(garde, "B2", companyCountryUpper);
  setWorksheetText(garde, "D22", companyName);
  setWorksheetText(garde, "C26", currentCompanyDetails.sigleUsuel || "");
  setWorksheetText(garde, "C28", currentCompanyDetails.siegeSocial || "");
  setWorksheetText(garde, "D30", currentCompanyDetails.nif || "");
  setWorksheetText(garde, "D31", currentCompanyDetails.nes || "");
  setWorksheetDate(garde, "E17", exerciseEnd);

  setWorksheetDate(r1, "N8", exerciseStart);
  setWorksheetDate(r1, "S8", exerciseEnd);
  setWorksheetDate(r1, "H10", exerciseEnd);
  setWorksheetDate(r1, "H12", previousExerciseEnd);
  setWorksheetText(r1, "C20", legalFormLabel);
  setWorksheetText(r1, "Q20", currentCompanyDetails.sigleUsuel || "");
  setWorksheetText(r1, "C23", currentCompanyDetails.tel || "");
  setWorksheetText(r1, "E23", currentCompanyDetails.emailCompta || (acct ? acct.email : ""));
  setWorksheetNumber(r1, "K23", phoneCode);
  setWorksheetText(r1, "Q23", city);
  setWorksheetText(r1, "C29", currentCompanyDetails.activitePrincipale || "");
  setWorksheetText(r1, "C32", [contactName, currentCompanyDetails.tel || "", currentCompanyDetails.emailCompta || (acct ? acct.email : "")].filter(Boolean).join(" | "));
  setWorksheetText(r1, "C35", contactName);
  setWorksheetText(r1, "C43", currentCompanyDetails.expertComptable || "");

  setDigitSequence(r2, "Q9", legalFormCode, 2);
  setDigitSequence(r2, "Q11", regimeCode, 1);
  setDigitSequence(r2, "Q13", countryCode, 2);
  setDigitSequence(r2, "Q15", 1, 2);
  setDigitSequence(r2, "Q17", 0, 2);
  setDigitSequence(r2, "Q20", new Date(exerciseStart || exerciseEnd || Date.now()).getFullYear(), 4);

  setWorksheetText(info, "E3", currentCompanyDetails.nes || "");
  setWorksheetText(info, "P53", currentCompanyDetails.nes || "");
  setWorksheetDate(info, "P9", exerciseStart);
  setWorksheetDate(info, "U9", exerciseEnd);
  setWorksheetDate(info, "J11", exerciseEnd);
  setWorksheetDate(info, "J13", previousExerciseEnd);
  setWorksheetText(info, "F15", city);
  setWorksheetText(info, "M15", currentCompanyDetails.rccm || currentCompanyDetails.nif || "");
  setWorksheetText(info, "E21", companyName);
  setWorksheetText(info, "L21", currentCompanyDetails.sigleUsuel || companyName);
  setWorksheetText(info, "E24", currentCompanyDetails.tel || "");
  setWorksheetText(info, "G24", currentCompanyDetails.emailCompta || (acct ? acct.email : ""));
  setWorksheetText(info, "K24", countryCode);
  setWorksheetText(info, "L24", "");
  setWorksheetText(info, "N24", city);
  setWorksheetText(info, "E28", currentCompanyDetails.siegeSocial || "");
  setWorksheetText(info, "E31", currentCompanyDetails.activitePrincipale || "");
  setWorksheetText(info, "E34", contactName);
  setWorksheetText(info, "G34", contactAddress);
  setWorksheetText(info, "K34", currentCompanyDetails.tel || "");
  setWorksheetText(info, "M34", currentCompanyDetails.emailCompta || (acct ? acct.email : ""));
  setWorksheetText(info, "U34", currentCompanyDetails.expertComptable ? "EXPERT-COMPTABLE" : "COMPTABILITE");
  setWorksheetText(info, "E56", "Personne morale");
  setWorksheetText(info, "G56", legalFormLabel);
  setWorksheetText(info, "I56", fiscalRegimeLabel);
  setWorksheetText(info, "J56", countryLabel);
  setWorksheetText(info, "P56", currentCompanyDetails.nes || "Neant");

  const tauxIS = parsePercentFieldValue(currentCompanyDetails.tauxIS, "standard");
  const tauxIMF = parsePercentFieldValue(currentCompanyDetails.tauxIMF, "small-percent");
  const tauxTVA = parsePercentFieldValue(currentCompanyDetails.tauxTVA, "standard");
  if (tauxIS !== null) setWorksheetNumber(info, "G59", tauxIS);
  if (tauxIMF !== null) setWorksheetNumber(info, "J59", tauxIMF);
  if (tauxTVA !== null) setWorksheetNumber(info, "E62", tauxTVA);

  populateSycebnlBalanceWorksheets(workbook);
}

async function loadTemplateWorkbook(templatePath, cacheVersion = "20260406b") {
  const response = await fetch(`${templatePath}?v=${cacheVersion}`);
  if (!response.ok) throw new Error(`Template ${templatePath} introuvable (${response.status})`);
  const buffer = await response.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellStyles: true, cellFormula: true });
}

async function loadExactLiasseTemplateWorkbook() {
  return loadTemplateWorkbook(EXACT_LIASSE_TEMPLATE_PATH, "20260406b");
}

async function loadExactSycebnlTemplateWorkbook(templatePath) {
  return loadTemplateWorkbook(templatePath, "20260406b");
}

async function loadExactForecastTemplateWorkbook() {
  return loadTemplateWorkbook(EXACT_FORECAST_TEMPLATE_PATH, "20260406b");
}

function clampNumber(value, minValue, maxValue, fallbackValue = 0) {
  const num = Number(value);
  if (Number.isNaN(num)) return fallbackValue;
  return Math.min(maxValue, Math.max(minValue, num));
}

function roundPositiveAmount(value) {
  const num = Number(value) || 0;
  return num > 0 ? Math.round(num) : 0;
}

function spreadAnnualAmountOverMonths(total) {
  const safeTotal = roundPositiveAmount(total);
  const base = Math.floor(safeTotal / 12);
  let remainder = safeTotal - (base * 12);
  return Array.from({ length: 12 }, () => {
    const amount = base + (remainder > 0 ? 1 : 0);
    if (remainder > 0) remainder -= 1;
    return amount;
  });
}

function sumCurrentMovementsByPrefixes(bal, prefixes, normalSide) {
  let total = 0;
  Object.keys(bal).forEach((code) => {
    if (!prefixes.some((prefix) => String(code).startsWith(String(prefix)))) return;
    const row = bal[code] || {};
    const mvtD = (row.debit || 0) - (row.n1d || 0);
    const mvtC = (row.credit || 0) - (row.n1c || 0);
    total += normalSide === "credit" ? (mvtC - mvtD) : (mvtD - mvtC);
  });
  return total > 0 ? total : 0;
}

function sumClosingNetByPrefixes(bal, prefixes) {
  let total = 0;
  Object.keys(bal).forEach((code) => {
    if (!prefixes.some((prefix) => String(code).startsWith(String(prefix)))) return;
    const row = bal[code] || {};
    total += (row.debit || 0) - (row.credit || 0);
  });
  return total > 0 ? total : 0;
}

function sumClosingCreditNetByPrefixes(bal, prefixes) {
  let total = 0;
  Object.keys(bal).forEach((code) => {
    if (!prefixes.some((prefix) => String(code).startsWith(String(prefix)))) return;
    const row = bal[code] || {};
    total += (row.credit || 0) - (row.debit || 0);
  });
  return total > 0 ? total : 0;
}

function getForecastLegalStatus(formeJuridique) {
  const key = normalizeTemplateKey(formeJuridique);
  if (!key) return "SARL (IS)";
  if (key.includes("micro")) return "Micro-entreprise";
  if (key.includes("entreprise individuelle") || key === "ei") return "Entreprise individuelle au r\u00e9el (IR)";
  if (key.includes("eurl")) return "EURL (IS)";
  if (key.includes("sasu")) return "SASU (IS)";
  if (key.includes("sas") || key === "sa" || key.includes("societe anonyme")) return "SAS (IS)";
  if (key.includes("sarl")) return "SARL (IS)";
  return "SARL (IS)";
}

function getForecastSalesType(merchandiseRevenue, serviceRevenue, activityLabel) {
  if (merchandiseRevenue > 0 && serviceRevenue > 0) return "Mixte";
  if (serviceRevenue > 0) return "Services";
  if (merchandiseRevenue > 0) return "Marchandises (y compris h\u00e9bergement et restauration)";

  const activityKey = normalizeTemplateKey(activityLabel);
  if (activityKey.includes("service") || activityKey.includes("consult") || activityKey.includes("prestation")) return "Services";
  if (activityKey.includes("commerce") || activityKey.includes("vente") || activityKey.includes("boutique") || activityKey.includes("restaurant")) {
    return "Marchandises (y compris h\u00e9bergement et restauration)";
  }
  return "Mixte";
}

function getForecastTemplateSnapshot() {
  const bal = computeBalances();
  const companyName = getCompanyDisplayName();
  const durationMonths = getExerciseDurationMonths(currentCompanyDetails.exerciceDu || "", currentCompanyDetails.exerciceAu || "");
  const annualizationFactor = durationMonths > 0 ? (12 / durationMonths) : 1;

  const merchandiseRevenue = roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["701", "702", "703", "704"], "credit") * annualizationFactor);
  const serviceRevenue = roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["706", "707"], "credit") * annualizationFactor);
  const annualPurchases = roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["601", "602", "604", "605"], "debit") * annualizationFactor);
  const employeeCompensation = roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["661", "662", "663"], "debit") * annualizationFactor);

  const intangibleSetup = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["21"]));
  const realEstateSetup = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["22", "231", "232"]));
  const worksSetup = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["233", "234", "235", "238"]));
  const equipmentSetup = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["241", "242", "243", "245", "246", "248"]));
  const officeEquipmentSetup = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["244"]));
  const openingStock = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["31", "32", "33", "35"]));
  const startingCash = roundPositiveAmount(sumClosingNetByPrefixes(bal, ["50", "52", "53", "54", "57", "58"]));

  const growthYear2 = (merchandiseRevenue + serviceRevenue) > 0 ? 0.1 : 0.15;
  const growthYear3 = (merchandiseRevenue + serviceRevenue) > 0 ? 0.08 : 0.1;
  const purchaseRatio = clampNumber(merchandiseRevenue > 0 ? (annualPurchases / merchandiseRevenue) : 0.5, 0, 0.95, 0.5);
  const cityOrCommune = guessCityFromAddress(currentCompanyDetails.siegeSocial || "", currentCompanyDetails.pays || "") || currentCompanyDetails.siegeSocial || "";
  const ownerName = companyName || "";
  const projectName = [companyName, currentCompanyDetails.activitePrincipale].filter(Boolean).join(" - ") || companyName || "Projet OHADA Compta";
  const legalStatus = getForecastLegalStatus(currentCompanyDetails.formeJuridique || "");
  const salesType = getForecastSalesType(merchandiseRevenue, serviceRevenue, currentCompanyDetails.activitePrincipale || "");
  const merchandiseMonthly = spreadAnnualAmountOverMonths(merchandiseRevenue);
  const serviceMonthly = spreadAnnualAmountOverMonths(serviceRevenue);

  const employeeYear1 = employeeCompensation;
  const employeeYear2 = roundPositiveAmount(employeeYear1 * (1 + growthYear2));
  const employeeYear3 = roundPositiveAmount(employeeYear2 * (1 + growthYear3));

  return {
    ownerName,
    projectName,
    legalStatus,
    cityOrCommune,
    salesType,
    merchandiseRevenue,
    serviceRevenue,
    annualPurchases,
    purchaseRatio,
    intangibleSetup,
    realEstateSetup,
    worksSetup,
    equipmentSetup,
    officeEquipmentSetup,
    openingStock,
    startingCash,
    growthYear2,
    growthYear3,
    employeeYear1,
    employeeYear2,
    employeeYear3,
    managerYear1: 0,
    managerYear2: 0,
    managerYear3: 0,
    merchandiseMonthly,
    serviceMonthly,
    amortizationYears: 5,
    acreEligible: false,
    assumptionsCount: [
      currentCompanyDetails.raisonSociale,
      currentCompanyDetails.formeJuridique,
      currentCompanyDetails.tel,
      currentCompanyDetails.emailCompta,
      merchandiseRevenue + serviceRevenue > 0 ? "ca" : "",
      startingCash > 0 ? "cash" : ""
    ].filter(Boolean).length
  };
}

function getMonteCarloBaseScenario() {
  const bal = computeBalances();
  const forecast = getForecastTemplateSnapshot();
  let totalCharges = 0;
  let depreciationCharge = 0;

  Object.keys(bal).forEach((code) => {
    const row = bal[code] || {};
    const movementDebit = (row.debit || 0) - (row.n1d || 0);
    const movementCredit = (row.credit || 0) - (row.n1c || 0);
    const chargeAmount = movementDebit - movementCredit;
    if (chargeAmount <= 0) return;

    if (
      String(code).startsWith("6")
      || String(code).startsWith("81")
      || String(code).startsWith("83")
      || String(code).startsWith("85")
      || String(code).startsWith("87")
      || String(code).startsWith("89")
    ) {
      totalCharges += chargeAmount;
    }
    if (String(code).startsWith("681") || String(code).startsWith("68")) {
      depreciationCharge += chargeAmount;
    }
  });

  const totalRevenue = roundPositiveAmount(forecast.merchandiseRevenue + forecast.serviceRevenue);
  const totalSetup = roundPositiveAmount(
    forecast.intangibleSetup
    + forecast.realEstateSetup
    + forecast.worksSetup
    + forecast.equipmentSetup
    + forecast.officeEquipmentSetup
  );
  const directCosts = roundPositiveAmount(forecast.annualPurchases);
  const fixedCosts = roundPositiveAmount(Math.max(totalCharges - directCosts, 0));
  const workingCapitalBase = roundPositiveAmount(
    forecast.openingStock
    + sumClosingNetByPrefixes(bal, ["41"])
    - sumClosingCreditNetByPrefixes(bal, ["40", "42", "43", "44"])
  );
  const workingCapitalPct = totalRevenue > 0
    ? clampNumber(workingCapitalBase / totalRevenue, 0.02, 0.45, 0.08)
    : 0.08;
  const grossMargin = totalRevenue > 0
    ? clampNumber((totalRevenue - directCosts) / totalRevenue, 0.08, 0.92, 0.35)
    : clampNumber(1 - forecast.purchaseRatio, 0.08, 0.92, 0.35);
  const depreciation = roundPositiveAmount(
    depreciationCharge || (totalSetup / Math.max(1, forecast.amortizationYears || 5))
  );
  const capex = roundPositiveAmount(
    Math.max(totalSetup * 0.08, depreciation, totalRevenue * 0.04)
  );
  const fallbackRevenue = roundPositiveAmount(Math.max((forecast.startingCash || 0) * 2, 25000000));
  const revenue = totalRevenue || fallbackRevenue;

  return {
    revenue,
    directCosts,
    totalCharges: roundPositiveAmount(totalCharges),
    fixedCosts: fixedCosts || roundPositiveAmount(revenue * 0.24),
    depreciation,
    capex,
    workingCapitalBase,
    workingCapitalPct,
    grossMargin,
    taxRate: clampNumber((Number(currentCompanyDetails.tauxIS) || 25) / 100, 0, 0.45, 0.25),
    openingCash: forecast.startingCash,
    totalSetup,
    forecast
  };
}

function buildMonteCarloDefaults() {
  const base = getMonteCarloBaseScenario();
  return {
    baseRevenue: base.revenue,
    baseFixedCosts: roundPositiveAmount(base.fixedCosts),
    baseDepreciation: roundPositiveAmount(base.depreciation),
    baseCapex: roundPositiveAmount(base.capex),
    taxRate: base.taxRate,
    iterations: 2500,
    revenueGrowthMin: -0.08,
    revenueGrowthMode: base.revenue > 0 ? 0.06 : 0.12,
    revenueGrowthMax: 0.18,
    grossMarginMin: clampNumber(base.grossMargin - 0.08, 0.05, 0.9, 0.25),
    grossMarginMode: clampNumber(base.grossMargin, 0.08, 0.92, 0.35),
    grossMarginMax: clampNumber(base.grossMargin + 0.08, 0.1, 0.98, 0.45),
    fixedCostMin: 0.92,
    fixedCostMode: 1,
    fixedCostMax: 1.16,
    workingCapitalMin: clampNumber(base.workingCapitalPct * 0.75, 0.01, 0.35, 0.05),
    workingCapitalMode: clampNumber(base.workingCapitalPct, 0.02, 0.4, 0.08),
    workingCapitalMax: clampNumber((base.workingCapitalPct * 1.25) + 0.02, 0.03, 0.45, 0.12),
    capexMin: 0.85,
    capexMode: 1,
    capexMax: 1.3
  };
}

function sanitizeMonteCarloConfig(config = {}) {
  const defaults = buildMonteCarloDefaults();
  const revenueGrowth = normalizeTriangularInputs(
    config.revenueGrowthMin,
    config.revenueGrowthMode,
    config.revenueGrowthMax,
    {
      min: defaults.revenueGrowthMin,
      mode: defaults.revenueGrowthMode,
      max: defaults.revenueGrowthMax
    }
  );
  const grossMargin = normalizeTriangularInputs(
    config.grossMarginMin,
    config.grossMarginMode,
    config.grossMarginMax,
    {
      min: defaults.grossMarginMin,
      mode: defaults.grossMarginMode,
      max: defaults.grossMarginMax
    }
  );
  const fixedCost = normalizeTriangularInputs(
    config.fixedCostMin,
    config.fixedCostMode,
    config.fixedCostMax,
    {
      min: defaults.fixedCostMin,
      mode: defaults.fixedCostMode,
      max: defaults.fixedCostMax
    }
  );
  const workingCapital = normalizeTriangularInputs(
    config.workingCapitalMin,
    config.workingCapitalMode,
    config.workingCapitalMax,
    {
      min: defaults.workingCapitalMin,
      mode: defaults.workingCapitalMode,
      max: defaults.workingCapitalMax
    }
  );
  const capex = normalizeTriangularInputs(
    config.capexMin,
    config.capexMode,
    config.capexMax,
    {
      min: defaults.capexMin,
      mode: defaults.capexMode,
      max: defaults.capexMax
    }
  );

  return {
    baseRevenue: roundPositiveAmount(config.baseRevenue ?? defaults.baseRevenue),
    baseFixedCosts: roundPositiveAmount(config.baseFixedCosts ?? defaults.baseFixedCosts),
    baseDepreciation: roundPositiveAmount(config.baseDepreciation ?? defaults.baseDepreciation),
    baseCapex: roundPositiveAmount(config.baseCapex ?? defaults.baseCapex),
    taxRate: clampNumber(config.taxRate ?? defaults.taxRate, 0, 0.5, defaults.taxRate),
    iterations: Math.round(clampNumber(config.iterations ?? defaults.iterations, 250, 10000, defaults.iterations)),
    revenueGrowthMin: clampNumber(revenueGrowth.min, -0.8, 2, defaults.revenueGrowthMin),
    revenueGrowthMode: clampNumber(revenueGrowth.mode, -0.8, 2, defaults.revenueGrowthMode),
    revenueGrowthMax: clampNumber(revenueGrowth.max, -0.8, 2, defaults.revenueGrowthMax),
    grossMarginMin: clampNumber(grossMargin.min, 0.01, 0.98, defaults.grossMarginMin),
    grossMarginMode: clampNumber(grossMargin.mode, 0.01, 0.98, defaults.grossMarginMode),
    grossMarginMax: clampNumber(grossMargin.max, 0.01, 0.98, defaults.grossMarginMax),
    fixedCostMin: clampNumber(fixedCost.min, 0.2, 3, defaults.fixedCostMin),
    fixedCostMode: clampNumber(fixedCost.mode, 0.2, 3, defaults.fixedCostMode),
    fixedCostMax: clampNumber(fixedCost.max, 0.2, 3, defaults.fixedCostMax),
    workingCapitalMin: clampNumber(workingCapital.min, 0, 0.8, defaults.workingCapitalMin),
    workingCapitalMode: clampNumber(workingCapital.mode, 0, 0.8, defaults.workingCapitalMode),
    workingCapitalMax: clampNumber(workingCapital.max, 0, 0.8, defaults.workingCapitalMax),
    capexMin: clampNumber(capex.min, 0.1, 4, defaults.capexMin),
    capexMode: clampNumber(capex.mode, 0.1, 4, defaults.capexMode),
    capexMax: clampNumber(capex.max, 0.1, 4, defaults.capexMax)
  };
}

function getMonteCarloConfig() {
  return sanitizeMonteCarloConfig(currentCompanyDetails.monteCarloConfig || {});
}

function saveMonteCarloConfig(config = {}) {
  currentCompanyDetails.monteCarloConfig = sanitizeMonteCarloConfig(config);
  saveCompanyData();
  return currentCompanyDetails.monteCarloConfig;
}

function getMonteCarloInputDisplayValue(field, config) {
  const value = config[field];
  if (!Number.isFinite(Number(value))) return "";
  if (MONTE_CARLO_PERCENT_FIELDS.has(field)) return (Number(value) * 100).toFixed(2);
  if (field === "iterations") return String(Math.round(Number(value)));
  return String(Math.round(Number(value)));
}

function collectMonteCarloConfigFromForm() {
  const patch = {};
  MONTE_CARLO_FIELD_IDS.forEach((field) => {
    const el = document.getElementById(`mc-${field}`);
    if (!el) return;
    const rawValue = el.value;
    patch[field] = MONTE_CARLO_PERCENT_FIELDS.has(field)
      ? (Number(rawValue) || 0) / 100
      : (Number(rawValue) || 0);
  });
  return saveMonteCarloConfig({
    ...getMonteCarloConfig(),
    ...patch
  });
}

function runMonteCarloSimulation() {
  const config = collectMonteCarloConfigFromForm();
  if (!config.baseRevenue || config.iterations < 250) {
    showToast("Renseignez un chiffre d'affaires de base et au moins 250 iterations.", "error");
    return;
  }

  const revenues = [];
  const ebitdas = [];
  const netIncomes = [];
  const freeCashFlows = [];
  const operatingMargins = [];
  let lossCount = 0;
  let negativeCashCount = 0;

  for (let i = 0; i < config.iterations; i++) {
    const growthRate = triangularSample(config.revenueGrowthMin, config.revenueGrowthMode, config.revenueGrowthMax);
    const grossMargin = triangularSample(config.grossMarginMin, config.grossMarginMode, config.grossMarginMax);
    const fixedCostFactor = triangularSample(config.fixedCostMin, config.fixedCostMode, config.fixedCostMax);
    const workingCapitalPct = triangularSample(config.workingCapitalMin, config.workingCapitalMode, config.workingCapitalMax);
    const capexFactor = triangularSample(config.capexMin, config.capexMode, config.capexMax);

    const revenue = Math.max(0, config.baseRevenue * (1 + growthRate));
    const grossProfit = revenue * grossMargin;
    const fixedCosts = config.baseFixedCosts * fixedCostFactor;
    const ebitda = grossProfit - fixedCosts;
    const ebit = ebitda - config.baseDepreciation;
    const tax = ebit > 0 ? ebit * config.taxRate : 0;
    const netIncome = ebit - tax;
    const freeCashFlow = netIncome + config.baseDepreciation - (config.baseCapex * capexFactor) - (revenue * workingCapitalPct);
    const operatingMargin = revenue > 0 ? ebitda / revenue : 0;

    revenues.push(revenue);
    ebitdas.push(ebitda);
    netIncomes.push(netIncome);
    freeCashFlows.push(freeCashFlow);
    operatingMargins.push(operatingMargin);
    if (netIncome < 0) lossCount += 1;
    if (freeCashFlow < 0) negativeCashCount += 1;
  }

  const sortAsc = (values) => [...values].sort((a, b) => a - b);
  const sortedRevenue = sortAsc(revenues);
  const sortedEbitda = sortAsc(ebitdas);
  const sortedNet = sortAsc(netIncomes);
  const sortedFcf = sortAsc(freeCashFlows);
  const sortedMargin = sortAsc(operatingMargins);

  currentCompanyDetails.monteCarloLastRun = {
    generatedAt: new Date().toISOString(),
    iterations: config.iterations,
    baseScenario: getMonteCarloBaseScenario(),
    metrics: {
      revenue: {
        p10: percentile(sortedRevenue, 0.10),
        p50: percentile(sortedRevenue, 0.50),
        p90: percentile(sortedRevenue, 0.90),
        average: average(revenues)
      },
      ebitda: {
        p10: percentile(sortedEbitda, 0.10),
        p50: percentile(sortedEbitda, 0.50),
        p90: percentile(sortedEbitda, 0.90),
        average: average(ebitdas)
      },
      netIncome: {
        p10: percentile(sortedNet, 0.10),
        p50: percentile(sortedNet, 0.50),
        p90: percentile(sortedNet, 0.90),
        average: average(netIncomes)
      },
      freeCashFlow: {
        p10: percentile(sortedFcf, 0.10),
        p50: percentile(sortedFcf, 0.50),
        p90: percentile(sortedFcf, 0.90),
        average: average(freeCashFlows)
      },
      operatingMargin: {
        p10: percentile(sortedMargin, 0.10),
        p50: percentile(sortedMargin, 0.50),
        p90: percentile(sortedMargin, 0.90),
        average: average(operatingMargins)
      }
    },
    risk: {
      lossProbability: lossCount / config.iterations,
      negativeCashProbability: negativeCashCount / config.iterations,
      stressedMarginProbability: operatingMargins.filter((margin) => margin < 0.12).length / config.iterations
    },
    tails: {
      worstNetIncome: sortedNet[0],
      bestNetIncome: sortedNet[sortedNet.length - 1],
      worstFreeCashFlow: sortedFcf[0],
      bestFreeCashFlow: sortedFcf[sortedFcf.length - 1]
    },
    histogram: buildHistogram(freeCashFlows, 9)
  };
  saveCompanyData();
  render();
  showToast(`Simulation Monte Carlo terminee (${fmt(config.iterations)} iterations).`, "success");
}

function applyMonteCarloAccountingBase() {
  const currentConfig = getMonteCarloConfig();
  const refreshed = buildMonteCarloDefaults();
  saveMonteCarloConfig({
    ...refreshed,
    iterations: currentConfig.iterations
  });
  render();
  showToast("Hypotheses Monte Carlo recalculees a partir de la comptabilite courante.", "success");
}

function resetMonteCarloModule() {
  currentCompanyDetails.monteCarloConfig = sanitizeMonteCarloConfig({});
  delete currentCompanyDetails.monteCarloLastRun;
  saveCompanyData();
  render();
  showToast("Module Monte Carlo reinitialise.", "info");
}

function normalizeCostReductionRow(row = {}, index = 0) {
  const pillar = COST_REDUCTION_PILLARS.find((item) => item.key === row.pillar) ? row.pillar : "optimisation";
  const status = COST_REDUCTION_STATUS_OPTIONS.includes(row.status) ? row.status : "A etudier";
  return {
    id: String(row.id || `cost-${index + 1}`),
    technique: String(row.technique || "").trim(),
    pillar,
    classes: String(row.classes || "").trim(),
    estimatedGain: roundPositiveAmount(row.estimatedGain),
    status,
    owner: String(row.owner || "").trim(),
    action: String(row.action || "").trim(),
    targetRate: clampNumber(row.targetRate, 0.01, 0.3, 0.03),
    recommendedPrefixes: Array.isArray(row.recommendedPrefixes)
      ? row.recommendedPrefixes.map((prefix) => String(prefix))
      : []
  };
}

function createDefaultCostReductionPlan() {
  return [
    {
      id: "cost-1",
      technique: "Regrouper les achats et renegocier les volumes",
      pillar: "combinaison",
      classes: "60 / 61",
      estimatedGain: 0,
      status: "Priorite 30 jours",
      owner: "Direction achats",
      action: "Consolider les fournisseurs et lancer une renegociation cadre.",
      targetRate: 0.05,
      recommendedPrefixes: ["60", "61"]
    },
    {
      id: "cost-2",
      technique: "Mettre en concurrence les fournisseurs critiques",
      pillar: "adaptation",
      classes: "60 / 62",
      estimatedGain: 0,
      status: "A etudier",
      owner: "Achats / DAF",
      action: "Organiser un appel d'offres et standardiser les cahiers des charges.",
      targetRate: 0.04,
      recommendedPrefixes: ["60", "62"]
    },
    {
      id: "cost-3",
      technique: "Analyser les couts par activite (ABC)",
      pillar: "optimisation",
      classes: "61 / 62 / 65",
      estimatedGain: 0,
      status: "A etudier",
      owner: "Controle de gestion",
      action: "Identifier les activites a faible marge et supprimer les taches non utiles.",
      targetRate: 0.03,
      recommendedPrefixes: ["61", "62", "65"]
    },
    {
      id: "cost-4",
      technique: "Automatiser les taches repetitives (low-code / RPA)",
      pillar: "optimisation",
      classes: "62 / 65",
      estimatedGain: 0,
      status: "A etudier",
      owner: "Finance / SI",
      action: "Cibler la saisie, les rapprochements et les validations manuelles.",
      targetRate: 0.04,
      recommendedPrefixes: ["62", "65"]
    },
    {
      id: "cost-5",
      technique: "Externaliser selectivement les activites non coeur",
      pillar: "substitution",
      classes: "62 / 64 / 65",
      estimatedGain: 0,
      status: "A etudier",
      owner: "Direction generale",
      action: "Comparer le cout interne complet avec une prestation encadree.",
      targetRate: 0.05,
      recommendedPrefixes: ["62", "64", "65"]
    },
    {
      id: "cost-6",
      technique: "Revoir voyages, carburant et missions",
      pillar: "elimination",
      classes: "61 / 62",
      estimatedGain: 0,
      status: "Priorite 30 jours",
      owner: "Moyens generaux",
      action: "Fixer des plafonds, mutualiser les deplacements et suivre les ecarts.",
      targetRate: 0.06,
      recommendedPrefixes: ["61", "62"]
    },
    {
      id: "cost-7",
      technique: "Rationaliser les systemes herites et abonnements",
      pillar: "elimination",
      classes: "62 / 65",
      estimatedGain: 0,
      status: "A etudier",
      owner: "SI",
      action: "Supprimer les licences inactives et fusionner les outils redondants.",
      targetRate: 0.08,
      recommendedPrefixes: ["62", "65"]
    },
    {
      id: "cost-8",
      technique: "Reaffecter les equipes et actifs sous-utilises",
      pillar: "reaffectation",
      classes: "64 / 65",
      estimatedGain: 0,
      status: "A etudier",
      owner: "RH / Operations",
      action: "Rediriger les capacites libres avant tout nouvel engagement de depense.",
      targetRate: 0.04,
      recommendedPrefixes: ["64", "65"]
    },
    {
      id: "cost-9",
      technique: "Conception a cout et reduction des rebuts",
      pillar: "adaptation",
      classes: "60 / 61 / 62",
      estimatedGain: 0,
      status: "A etudier",
      owner: "Production",
      action: "Repenser les standards de production pour reduire la non-qualite.",
      targetRate: 0.05,
      recommendedPrefixes: ["60", "61", "62"]
    }
  ].map((row, index) => normalizeCostReductionRow(row, index));
}

function estimateCostReductionGain(prefixes, targetRate, bal = computeBalances()) {
  const baseAmount = roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, prefixes, "debit"));
  return roundPositiveAmount(baseAmount * (Number(targetRate) || 0));
}

function buildCostReductionPlanFromAccounting() {
  const bal = computeBalances();
  return createDefaultCostReductionPlan().map((row, index) => normalizeCostReductionRow({
    ...row,
    estimatedGain: estimateCostReductionGain(row.recommendedPrefixes, row.targetRate, bal)
  }, index));
}

function getCostReductionPlan() {
  const rawPlan = Array.isArray(currentCompanyDetails.costReductionPlan) && currentCompanyDetails.costReductionPlan.length
    ? currentCompanyDetails.costReductionPlan
    : buildCostReductionPlanFromAccounting();
  return rawPlan.map((row, index) => normalizeCostReductionRow(row, index));
}

function saveCostReductionPlan(plan) {
  currentCompanyDetails.costReductionPlan = plan.map((row, index) => normalizeCostReductionRow(row, index));
  saveCompanyData();
  return currentCompanyDetails.costReductionPlan;
}

function getCostFamilyOpportunities() {
  const bal = computeBalances();
  return [
    {
      code: "60",
      label: "Achats et consommables",
      note: "Marchandises, matieres et consommations directes.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["60"], "debit"))
    },
    {
      code: "61-62",
      label: "Transport et services exterieurs",
      note: "Prestations, loyers, honoraires, maintenance.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["61", "62"], "debit"))
    },
    {
      code: "63",
      label: "Impots et taxes",
      note: "Taxes non recuperables et versements assimiles.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["63"], "debit"))
    },
    {
      code: "661-663",
      label: "Personnel",
      note: "Salaires, indemnites et charges liees.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["661", "662", "663"], "debit"))
    },
    {
      code: "64-65",
      label: "Autres charges de structure",
      note: "Charges d'exploitation et frais divers.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["64", "65"], "debit"))
    },
    {
      code: "67",
      label: "Charges exceptionnelles",
      note: "Depenses non recurrentes a surveiller.",
      amount: roundPositiveAmount(sumCurrentMovementsByPrefixes(bal, ["67"], "debit"))
    }
  ].sort((a, b) => b.amount - a.amount);
}

function getCostReductionSnapshot() {
  const plan = getCostReductionPlan();
  const opportunities = getCostFamilyOpportunities();
  const totalPotential = plan.reduce((sum, row) => sum + (Number(row.estimatedGain) || 0), 0);
  const totalCharges = opportunities.reduce((sum, row) => sum + (Number(row.amount) || 0), 0);

  return {
    plan,
    opportunities,
    totalPotential,
    totalCharges,
    coverageRate: totalCharges > 0 ? totalPotential / totalCharges : 0,
    implementedCount: plan.filter((row) => row.status === "Appliquee").length,
    inProgressCount: plan.filter((row) => row.status === "En cours" || row.status === "Validee").length,
    priorityCount: plan.filter((row) => row.status === "Priorite 30 jours").length,
    topOpportunity: opportunities.find((row) => row.amount > 0) || null
  };
}

function collectCostReductionPlanFromForm() {
  const rows = Array.from(document.querySelectorAll("[data-cost-row]"));
  if (!rows.length) return getCostReductionPlan();

  return rows.map((row, index) => normalizeCostReductionRow({
    id: row.dataset.costRow || `cost-${index + 1}`,
    technique: (row.querySelector('[data-cost-field="technique"]') || {}).value || "",
    pillar: (row.querySelector('[data-cost-field="pillar"]') || {}).value || "optimisation",
    classes: (row.querySelector('[data-cost-field="classes"]') || {}).value || "",
    estimatedGain: (row.querySelector('[data-cost-field="estimatedGain"]') || {}).value || 0,
    status: (row.querySelector('[data-cost-field="status"]') || {}).value || "A etudier",
    owner: (row.querySelector('[data-cost-field="owner"]') || {}).value || "",
    action: (row.querySelector('[data-cost-field="action"]') || {}).value || "",
    targetRate: Number(row.dataset.targetRate) || 0.03,
    recommendedPrefixes: JSON.parse(row.dataset.recommendedPrefixes || "[]")
  }, index));
}

function saveCostReductionPlanFromForm() {
  const nextPlan = collectCostReductionPlanFromForm();
  saveCostReductionPlan(nextPlan);
  render();
  showToast("Plan de reduction des couts enregistre.", "success");
}

function seedCostReductionPlanFromAccounting() {
  saveCostReductionPlan(buildCostReductionPlanFromAccounting());
  render();
  showToast("Plan de reduction des couts recalcule a partir des charges comptables.", "success");
}

function addCostReductionAction() {
  const currentPlan = collectCostReductionPlanFromForm();
  currentPlan.push(normalizeCostReductionRow({
    id: `cost-${Date.now()}`,
    technique: "Nouvelle action",
    pillar: "optimisation",
    classes: "",
    estimatedGain: 0,
    status: "A etudier",
    owner: "",
    action: ""
  }, currentPlan.length));
  saveCostReductionPlan(currentPlan);
  render();
}

function resetCostReductionPlan() {
  if (!window.confirm("Reinitialiser le plan de reduction des couts avec les suggestions par defaut ?")) return;
  delete currentCompanyDetails.costReductionPlan;
  saveCostReductionPlan(buildCostReductionPlanFromAccounting());
  render();
  showToast("Plan de reduction des couts reinitialise.", "info");
}

function populateExactForecastTemplate(workbook) {
  const acct = getCurrentAccountMeta();
  const snapshot = getForecastTemplateSnapshot();
  const inputSheet = workbook.Sheets[EXACT_FORECAST_INPUT_SHEET_NAME];
  if (!inputSheet) throw new Error(`Feuille ${EXACT_FORECAST_INPUT_SHEET_NAME} introuvable dans le modele previsionnel.`);

  setWorksheetText(inputSheet, "B6", snapshot.ownerName || (acct ? acct.company : ""));
  setWorksheetText(inputSheet, "B7", snapshot.projectName);
  setWorksheetText(inputSheet, "B8", snapshot.legalStatus);
  setWorksheetText(inputSheet, "B9", currentCompanyDetails.tel || "");
  setWorksheetText(inputSheet, "B10", currentCompanyDetails.emailCompta || (acct ? acct.email : ""));
  setWorksheetText(inputSheet, "B11", snapshot.cityOrCommune);
  setWorksheetText(inputSheet, "B13", snapshot.salesType);

  setWorksheetNumber(inputSheet, "B19", snapshot.intangibleSetup);
  setWorksheetNumber(inputSheet, "B28", snapshot.realEstateSetup);
  setWorksheetNumber(inputSheet, "B29", snapshot.worksSetup);
  setWorksheetNumber(inputSheet, "B30", snapshot.equipmentSetup);
  setWorksheetNumber(inputSheet, "B31", snapshot.officeEquipmentSetup);
  setWorksheetNumber(inputSheet, "B32", snapshot.openingStock);
  setWorksheetNumber(inputSheet, "B33", snapshot.startingCash);
  setWorksheetNumber(inputSheet, "C36", snapshot.amortizationYears);

  snapshot.merchandiseMonthly.forEach((amount, index) => {
    const row = 103 + index;
    setWorksheetNumber(inputSheet, `B${row}`, amount > 0 ? 1 : 0);
    setWorksheetNumber(inputSheet, `C${row}`, amount);
  });

  snapshot.serviceMonthly.forEach((amount, index) => {
    const row = 103 + index;
    setWorksheetNumber(inputSheet, `G${row}`, amount > 0 ? 1 : 0);
    setWorksheetNumber(inputSheet, `H${row}`, amount);
  });

  setWorksheetNumber(inputSheet, "D117", snapshot.growthYear2);
  setWorksheetNumber(inputSheet, "I117", snapshot.growthYear2);
  setWorksheetNumber(inputSheet, "D118", snapshot.growthYear3);
  setWorksheetNumber(inputSheet, "I118", snapshot.growthYear3);
  setWorksheetNumber(inputSheet, "D123", snapshot.purchaseRatio);

  setWorksheetNumber(inputSheet, "B133", snapshot.employeeYear1);
  setWorksheetNumber(inputSheet, "C133", snapshot.employeeYear2);
  setWorksheetNumber(inputSheet, "D133", snapshot.employeeYear3);
  setWorksheetNumber(inputSheet, "B134", snapshot.managerYear1);
  setWorksheetNumber(inputSheet, "C134", snapshot.managerYear2);
  setWorksheetNumber(inputSheet, "D134", snapshot.managerYear3);
  setWorksheetText(inputSheet, "C136", snapshot.acreEligible ? "Oui" : "Non");
}

function getMimeTypeForFilename(filename) {
  const lower = String(filename || "").toLowerCase();
  if (lower.endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (lower.endsWith(".csv")) return "text/csv;charset=utf-8";
  if (lower.endsWith(".json")) return "application/json;charset=utf-8";
  return "application/octet-stream";
}

function downloadBlobFile(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

function workbookToBlob(workbook, filename, bookType = "xlsx") {
  const buffer = XLSX.write(workbook, { bookType, type: "array", compression: true });
  return new Blob([buffer], { type: getMimeTypeForFilename(filename) });
}

async function tryShareBlobFile(blob, filename, title, text) {
  if (typeof navigator === "undefined" || typeof navigator.share !== "function" || typeof File === "undefined") {
    return { shared: false, unsupported: true };
  }

  const file = new File([blob], filename, {
    type: blob.type || getMimeTypeForFilename(filename),
    lastModified: Date.now()
  });

  if (typeof navigator.canShare === "function" && !navigator.canShare({ files: [file] })) {
    return { shared: false, unsupported: true };
  }

  try {
    await navigator.share({
      title: title || filename,
      text: text || `Fichier exporte depuis OHADA COMPTA: ${filename}`,
      files: [file]
    });
    return { shared: true };
  } catch (error) {
    if (error && error.name === "AbortError") return { shared: false, cancelled: true };
    console.warn("File share failed", error);
    return { shared: false, error };
  }
}

async function shareOrDownloadBlobFile(blob, filename, options = {}) {
  const result = await tryShareBlobFile(blob, filename, options.title, options.text);
  if (result.shared) {
    showToast(`${filename} partage avec succes.`, "success");
    return "shared";
  }
  if (result.cancelled) return "cancelled";
  downloadBlobFile(blob, filename);
  showToast(`Le partage direct n'est pas disponible ici. ${filename} a ete telecharge pour partage manuel.`, "info");
  return "downloaded";
}

function buildGeneratedBfaLiasseWorkbook() {
  if (currentRef !== "syscohada") {
    throw new Error("Le fichier BF[254] est reserve au referentiel SYSCOHADA. Basculez sur SYSCOHADA avant l'export.");
  }

  if (typeof XLSX === "undefined") {
    throw new Error("Librairie Excel non chargee. Verifiez votre connexion.");
  }

  const acct = getCurrentAccountMeta();
  const companyName = getCompanyDisplayName();
  const dsfStatuses = getDsfStatusRows();
  const now = new Date();
  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: BF_LIASSE_PACKET_NAME,
    Subject: "Liasse fiscale Burkina Faso SYSCOHADA",
    Author: "OHADA COMPTA",
    Company: companyName || "OHADA COMPTA",
    CreatedDate: now
  };

  const coverRows = [
    [BF_LIASSE_PACKET_NAME],
    [""],
    ["Genere depuis OHADA COMPTA", now.toLocaleString("fr-FR")],
    [""],
    ["Raison sociale", currentCompanyDetails.raisonSociale || (acct ? acct.company : "")],
    ["Sigle usuel", currentCompanyDetails.sigleUsuel || ""],
    ["Forme juridique", currentCompanyDetails.formeJuridique || ""],
    ["RCCM", currentCompanyDetails.rccm || ""],
    ["NIF / IFU", currentCompanyDetails.nif || ""],
    ["Siege social", currentCompanyDetails.siegeSocial || ""],
    ["Activite principale", currentCompanyDetails.activitePrincipale || ""],
    ["Capital social (XOF)", currentCompanyDetails.capitalSocial || ""],
    ["Regime fiscal", currentCompanyDetails.regimeFiscal || ""],
    ["Pays", currentCompanyDetails.pays || "Burkina Faso"],
    ["Exercice du", formatDateValue(currentCompanyDetails.exerciceDu || "")],
    ["Exercice au", formatDateValue(currentCompanyDetails.exerciceAu || "")],
    ["Telephone", currentCompanyDetails.tel || ""],
    ["Email", currentCompanyDetails.emailCompta || (acct ? acct.email : "")],
    ["Expert-comptable", currentCompanyDetails.expertComptable || ""],
    ["Commissaire aux comptes", currentCompanyDetails.commissaire || ""],
    ["Numero de teledeclarant (NES)", currentCompanyDetails.nes || ""],
    [""],
    ["Etat du dossier", `${dsfStatuses.filter(item => item.status === "Pret").length}/${dsfStatuses.length} rubriques pretes`],
    ["Observations", "Classeur genere automatiquement a partir des donnees saisies dans OHADA COMPTA."]
  ];
  appendWorkbookSheet(workbook, "Couverture", coverRows, [36, 55], 3);

  const checklistRows = [["Code", "Rubrique", "Statut", "Observation"]]
    .concat(dsfStatuses.map(item => [item.code, item.label, item.status, item.hint]));
  appendWorkbookSheet(workbook, "DSF_Checklist", checklistRows, [12, 42, 16, 60]);

  const journalRows = [["Date", "Journal", "Piece", "Compte", "Libelle", "Debit", "Credit", "Reference"]]
    .concat(journalEntries.map(entry => [entry.date, entry.journal, entry.piece, entry.compte, entry.libelle, entry.debit, entry.credit, entry.ref]));
  appendWorkbookSheet(workbook, "Journal", journalRows, [14, 10, 14, 12, 44, 16, 16, 18]);

  appendWorkbookSheet(workbook, "Balance", getBalanceExportRows(), [14, 36, 10, 16, 16, 16, 16, 16, 16]);
  appendWorkbookSheet(workbook, "Bilan", getBilanExportRows(), [30, 18, 6, 30, 18]);
  appendWorkbookSheet(workbook, "Resultat", getResultatExportRows(), [42, 16, 20]);
  appendWorkbookSheet(workbook, "FluxTresorerie", getFluxTresorerieExportRows(), [38, 18, 70], 2);
  appendWorkbookSheet(workbook, "Amortissements", getAmortissementExportRows(), [12, 36, 14, 10, 10, 14, 14, 14, 14]);
  appendWorkbookSheet(workbook, "Annexes", [
    ["Notes annexes"],
    [""],
    ["Rappel OHADA", "Le Systeme normal du SYSCOHADA revise comprend quarante-six (46) tableaux en notes annexes."],
    ["Portee du classeur", "Cette feuille recense un noyau de notes preparatoires a completer avant depot."],
    [""],
    ["Note 1", "Regles et methodes comptables"],
    ["Note 2", "Immobilisations incorporelles et corporelles"],
    ["Note 3", "Tableau des amortissements"],
    ["Note 4", "Immobilisations financieres"],
    ["Note 5", "Stocks et en-cours"],
    ["Note 6", "Creances et dettes"],
    ["Note 7", "Tresorerie"],
    ["Note 8", "Capitaux propres"],
    ["Note 9", "Emprunts et dettes financieres"],
    ["Note 10", "Charges de personnel"],
    ["Note 11", "Engagements hors bilan"],
    ["Note 12", "Evenements posterieurs a la cloture"],
    ["Note 13", "Parties liees"],
    ["Note 14", "Informations fiscales (DSF)"],
  ], [12, 48], 2);

  return workbook;
}

function downloadGeneratedBfaLiasseFiscale() {
  try {
    const workbook = buildGeneratedBfaLiasseWorkbook();
    XLSX.writeFile(workbook, BF_LIASSE_DOWNLOAD_NAME, { compression: true });
    showToast(`Classeur ${BF_LIASSE_DOWNLOAD_NAME} genere avec succes.`, "success");
  } catch (error) {
    showToast(error.message || "Impossible de generer la liasse interne.", "error");
  }
}

async function shareGeneratedBfaLiasseFiscale() {
  try {
    const workbook = buildGeneratedBfaLiasseWorkbook();
    const blob = workbookToBlob(workbook, BF_LIASSE_DOWNLOAD_NAME);
    await shareOrDownloadBlobFile(blob, BF_LIASSE_DOWNLOAD_NAME, {
      title: BF_LIASSE_PACKET_NAME,
      text: `Liasse fiscale exportee depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
    });
  } catch (error) {
    showToast(error.message || "Impossible de partager la liasse interne.", "error");
  }
}

async function buildExactLiasseWorkbook() {
  if (typeof XLSX === "undefined") {
    throw new Error("Librairie Excel non chargee. Verifiez votre connexion.");
  }

  const templateMeta = getExactFiscalTemplateMeta();
  let workbook;

  if (templateMeta.referential === "syscohada") {
    workbook = await loadExactLiasseTemplateWorkbook();
    populateExactLiasseTemplate(workbook);
  } else {
    workbook = await loadExactSycebnlTemplateWorkbook(templateMeta.templatePath);
    populateExactSycebnlTemplate(workbook);
  }

  return { workbook, templateMeta };
}

async function downloadExactLiasseFiscale() {
  const { workbook, templateMeta } = await buildExactLiasseWorkbook();
  XLSX.writeFile(workbook, templateMeta.downloadName, { bookType: "xlsx", compression: true });
  showToast(`Le modele exact ${templateMeta.downloadName} a ete rempli avec succes.`, "success");
}

async function shareExactLiasseFiscale() {
  const { workbook, templateMeta } = await buildExactLiasseWorkbook();
  const blob = workbookToBlob(workbook, templateMeta.downloadName);
  await shareOrDownloadBlobFile(blob, templateMeta.downloadName, {
    title: templateMeta.downloadName,
    text: `${templateMeta.downloadName} preparee depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
  });
}

async function downloadBfaLiasseFiscale() {
  try {
    await downloadExactLiasseFiscale();
  } catch (error) {
    console.warn("Exact template export failed", error);
    if (currentRef === "syscohada") {
      showToast("Modele LIASSE.xlsx indisponible. Generation du classeur interne en secours.", "info");
      downloadGeneratedBfaLiasseFiscale();
      return;
    }
    showToast(error.message || "Le modele SYCEBNL exact est indisponible.", "error");
  }
}

async function shareBfaLiasseFiscale() {
  try {
    await shareExactLiasseFiscale();
  } catch (error) {
    console.warn("Exact template share failed", error);
    if (currentRef === "syscohada") {
      showToast("Modele LIASSE.xlsx indisponible. Partage du classeur interne en secours.", "info");
      await shareGeneratedBfaLiasseFiscale();
      return;
    }
    showToast(error.message || "Le modele SYCEBNL exact est indisponible.", "error");
  }
}

async function buildExactForecastWorkbook() {
  if (typeof XLSX === "undefined") {
    throw new Error("Librairie Excel non chargee. Verifiez votre connexion.");
  }

  const workbook = await loadExactForecastTemplateWorkbook();
  populateExactForecastTemplate(workbook);
  return workbook;
}

async function downloadForecastWorkbook() {
  try {
    const workbook = await buildExactForecastWorkbook();
    XLSX.writeFile(workbook, EXACT_FORECAST_DOWNLOAD_NAME, { bookType: "xlsx", compression: true });
    showToast(`Le modele ${EXACT_FORECAST_DOWNLOAD_NAME} a ete genere avec succes.`, "success");
  } catch (error) {
    showToast(error.message || "Impossible de generer le plan financier previsionnel.", "error");
  }
}

async function shareForecastWorkbook() {
  try {
    const workbook = await buildExactForecastWorkbook();
    const blob = workbookToBlob(workbook, EXACT_FORECAST_DOWNLOAD_NAME);
    await shareOrDownloadBlobFile(blob, EXACT_FORECAST_DOWNLOAD_NAME, {
      title: "Plan financier previsionnel",
      text: `Plan financier previsionnel exporte depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
    });
  } catch (error) {
    showToast(error.message || "Impossible de partager le plan financier previsionnel.", "error");
  }
}

function buildBalanceTemplateBlob() {
  const plan = getPlan();
  let csv = "Compte,Libelle,S.O. Debit,S.O. Credit\n";
  plan.forEach(a => {
    const ob = OPENING_BALANCES[a.numero] || { n1d: 0, n1c: 0 };
    csv += `"${a.numero}","${a.libelle}",${ob.n1d},${ob.n1c}\n`;
  });
  return new Blob([csv], { type: getMimeTypeForFilename("balance_ouverture_template.csv") });
}

async function shareBalanceTemplate() {
  const blob = buildBalanceTemplateBlob();
  await shareOrDownloadBlobFile(blob, "balance_ouverture_template.csv", {
    title: "Modele de balance d'ouverture",
    text: `Modele CSV de balance d'ouverture exporte depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
  });
}

function buildAmortCsvBlob() {
  const bal = computeBalances();
  const plan = getPlan();
  const durees = {"21":5,"211":5,"212":5,"213":10,"214":10,"22":0,"23":20,"231":20,"232":20,"234":10,"235":7,"24":5,"241":5,"242":5,"244":3,"245":5,"246":7,"248":5};
  function ac(c){if(c.startsWith("21"))return"281";if(c.startsWith("22"))return"282";if(c.startsWith("23"))return"283";return"284";}
  let csv = "Compte,Libelle,VBO,Duree,Taux%,Cumul N-1,Dotation N,Cumul N,VNC\n";
  plan.filter(a=>a.classe===2&&a.sens==="Debit"&&bal[a.numero]).forEach(a=>{
    const vbo=bal[a.numero].debit; if(!vbo) return;
    const cn1=bal[ac(a.numero)]?(bal[ac(a.numero)].n1c||0):0;
    const d=durees[a.numero]||5;
    const dot=d>0?Math.round(vbo/d):0;
    csv += `"${a.numero}","${a.libelle}",${vbo},${d},${d>0?Math.round(10000/d)/100:0},${cn1},${dot},${cn1+dot},${Math.max(0,vbo-cn1-dot)}\n`;
  });
  return new Blob([csv], { type: getMimeTypeForFilename("amortissements.csv") });
}

async function shareAmortCsv() {
  const blob = buildAmortCsvBlob();
  await shareOrDownloadBlobFile(blob, "amortissements.csv", {
    title: "Tableau des amortissements",
    text: `Export CSV des amortissements genere depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
  });
}

// ═══════════════════════════════════════════════════════════
// RENDER
// ═══════════════════════════════════════════════════════════
function render() {
  const main = document.getElementById("main-content");
  document.querySelectorAll(".nav-btn").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === currentTab));
  switch (currentTab) {
    case "dashboard": main.innerHTML = renderDashboard(); break;
    case "plan": main.innerHTML = renderPlan(); break;
    case "journal": main.innerHTML = renderJournal(); break;
    case "grandlivre": main.innerHTML = renderGrandLivre(); break;
    case "balance": main.innerHTML = renderBalance(); break;
    case "bilan": main.innerHTML = renderBilan(); break;
    case "resultat": main.innerHTML = renderResultat(); break;
    case "tafire": main.innerHTML = renderTafire(); break;
    case "amortissements": main.innerHTML = renderAmortissements(); break;
    case "annexes": main.innerHTML = renderAnnexes(); break;
    case "saisie": main.innerHTML = renderSaisie(); break;
    case "cloture": main.innerHTML = renderCloture(); break;
    case "previsionnel": main.innerHTML = renderPrevisionnel(); break;
    case "montecarlo": main.innerHTML = renderMonteCarlo(); break;
    case "couts": main.innerHTML = renderReductionCouts(); break;
    case "dsf": main.innerHTML = renderDSF(); break;
    case "guide": main.innerHTML = renderGuide(); break;
    case "veille": main.innerHTML = renderVeilleOhada(); break;
    case "comparaison": main.innerHTML = renderComparaison(); break;
    case "parametres": main.innerHTML = renderParametres(); break;
    default: main.innerHTML = renderDashboard();
  }
  attachEvents();
  if (currentCompanyId) saveCompanyData();
}

// ═══════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════
function renderDashboard() {
  const plan = getPlan();
  const bal = computeBalances();
  const balanceSummary = getComputedBalanceSummary(bal);
  const totalDebit = balanceSummary.totalDebit;
  const totalCredit = balanceSummary.totalCredit;
  const refLabel = getReferentialLabel();

  // Company info
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
  const compName = currentCompanyDetails.raisonSociale || (acct ? acct.company : '');
  const capital = parseFloat(currentCompanyDetails.capitalSocial) || 0;
  const hasOpeningBalances = Object.keys(OPENING_BALANCES).length > 0;
  const hasProfile = hasCompanyProfileData();
  const profileComplete = isCompanyProfileComplete();
  const needsSetup = !profileComplete || !hasOpeningBalances || journalEntries.length === 0;
  const clotureSnapshot = getClotureSnapshot();

  // Financial ratios from balance
  let totalActifNet = 0, totalPassif = 0, totalProduits = 0, totalCharges = 0;
  let totalActifCirc = 0, totalPassifCirc = 0;
  Object.keys(bal).forEach(code => {
    const net = (bal[code].debit||0) - (bal[code].credit||0);
    const c1 = parseInt(code[0]);
    if ([1,2,3,4,5].includes(c1)) {
      if (net > 0) totalActifNet += net;
      if (net < 0) totalPassif += Math.abs(net);
    }
    if (c1 === 3 || c1 === 4 || c1 === 5) { if (net > 0) totalActifCirc += net; else totalPassifCirc += Math.abs(net); }
    if (c1 === 7 || code.startsWith('82') || code.startsWith('84')) totalProduits += (bal[code].credit||0) - (bal[code].debit||0);
    if (c1 === 6 || code.startsWith('81') || code.startsWith('83')) totalCharges += (bal[code].debit||0) - (bal[code].credit||0);
  });
  const resultat = totalProduits - totalCharges;
  const margePct = totalProduits > 0 ? (resultat / totalProduits * 100).toFixed(1) : null;
  const liquidite = totalPassifCirc > 0 ? (totalActifCirc / totalPassifCirc).toFixed(2) : null;
  const endettement = capital > 0 ? ((totalPassif / capital) * 100).toFixed(1) : null;

  return `
    ${compName ? `<div class="company-header"><div class="company-header-name">${compName}</div><div class="company-header-meta">${currentCompanyDetails.formeJuridique ? currentCompanyDetails.formeJuridique+' &bull; ' : ''}${currentCompanyDetails.siegeSocial || ''}${currentCompanyDetails.nif ? ' &bull; NIF: '+currentCompanyDetails.nif : ''}${currentCompanyDetails.rccm ? ' &bull; RCCM: '+currentCompanyDetails.rccm : ''}${currentCompanyDetails.exerciceDu ? ' &bull; Exercice: '+currentCompanyDetails.exerciceDu.slice(0,4) : ''}</div></div>` : ''}
    ${needsSetup ? `
    <div class="card" style="margin-bottom:20px;border-color:rgba(200,146,42,0.35);">
      <div class="card-header">
        <div>
          <div class="card-title">Espace de production initialise a vide</div>
          <div class="card-subtitle">Aucune donnee prechargee n'est presente. Terminez la configuration pour preparer vos etats OHADA.</div>
        </div>
      </div>
      <div class="grid-3">
        <div class="card" style="background:var(--surface2);border-color:rgba(68,138,255,0.22);">
          <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:8px;">
            <strong>Fiche entreprise</strong>
            <span style="font-size:0.74rem;font-weight:700;color:${profileComplete ? 'var(--green)' : hasProfile ? 'var(--orange)' : 'var(--muted)'};">${profileComplete ? 'Complete' : hasProfile ? 'A completer' : 'A renseigner'}</span>
          </div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Renseignez la raison sociale, choisissez le systeme comptable, le NIF, l'exercice et les informations legales.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('parametres')">Ouvrir les parametres</button>
        </div>
        <div class="card" style="background:var(--surface2);border-color:rgba(0,229,118,0.2);">
          <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:8px;">
            <strong>Balance d'ouverture</strong>
            <span style="font-size:0.74rem;font-weight:700;color:${hasOpeningBalances ? 'var(--green)' : 'var(--muted)'};">${hasOpeningBalances ? 'Chargee' : 'Vide'}</span>
          </div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Importez une balance CSV ou saisissez vos soldes avant d'editer le bilan.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('balance')">Importer ou verifier</button>
        </div>
        <div class="card" style="background:var(--surface2);border-color:rgba(255,145,0,0.24);">
          <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:8px;">
            <strong>Journal general</strong>
            <span style="font-size:0.74rem;font-weight:700;color:${journalEntries.length > 0 ? 'var(--green)' : 'var(--muted)'};">${journalEntries.length > 0 ? `${journalEntries.length} ecritures` : 'Aucune ecriture'}</span>
          </div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Ajoutez vos ecritures ou importez vos mouvements pour produire les etats financiers.</div>
          <button class="btn btn-gold" style="width:100%;" onclick="navigateToTab('saisie')">Saisir une ecriture</button>
        </div>
      </div>
    </div>` : ''}
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Referentiel</div><div class="kpi-value" style="font-size:1.2rem;color:var(--gold);">${refLabel}</div><div class="kpi-note">Norme en vigueur</div></div>
      <div class="kpi"><div class="kpi-label">Comptes</div><div class="kpi-value">${plan.length}</div><div class="kpi-note">Plan comptable actif</div></div>
      <div class="kpi"><div class="kpi-label">Classes</div><div class="kpi-value">${currentRef==="sycebnl"?9:8}</div><div class="kpi-note">Classes 1 a ${currentRef==="sycebnl"?9:8}</div></div>
      <div class="kpi"><div class="kpi-label">Ecritures</div><div class="kpi-value">${journalEntries.length}</div><div class="kpi-note">Journal general</div></div>
      <div class="kpi"><div class="kpi-label">Total debit</div><div class="kpi-value" style="color:var(--green);font-size:1.1rem;">${fmt(totalDebit)}</div><div class="kpi-note">XOF</div></div>
      <div class="kpi"><div class="kpi-label">Total credit</div><div class="kpi-value" style="color:var(--red);font-size:1.1rem;">${fmt(totalCredit)}</div><div class="kpi-note">XOF</div></div>
      <div class="kpi"><div class="kpi-label">Equilibre</div><div class="kpi-value" style="color:${balanceSummary.isBalanced ? 'var(--green)' : 'var(--red)'};">${balanceSummary.isBalanced ? 'OK' : 'ERREUR'}</div><div class="kpi-note">${balanceSummary.isBalanced ? 'Bilan pret a etre equilibre' : 'Ecart cumule: ' + fmt(Math.abs(balanceSummary.gap))}</div></div>
      <div class="kpi"><div class="kpi-label">Etats OHADA</div><div class="kpi-value">17</div><div class="kpi-note">Pays membres</div></div>
    </div>

    <div class="card" style="margin-top:16px;">
      <div class="card-header">
        <div>
          <div class="card-title">Actions rapides</div>
          <div class="card-subtitle">Acces direct a la cloture, aux simulations et aux exports.</div>
        </div>
      </div>
      <div class="grid-3">
        <div class="card" style="background:var(--surface2);">
          <div style="font-weight:700;margin-bottom:8px;">Cloture de l'exercice</div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Statut: <strong style="color:${clotureSnapshot.canClose ? 'var(--green)' : 'var(--orange)'};">${clotureSnapshot.canClose ? 'Pret a cloturer' : 'A verifier'}</strong></div>
          <button class="btn btn-gold" style="width:100%;" onclick="navigateToTab('cloture')">Ouvrir la cloture</button>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div style="font-weight:700;margin-bottom:8px;">Plan financier</div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Exportez le modele previsionnel exact rempli avec vos donnees actuelles.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('previsionnel')">Ouvrir le plan financier</button>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div style="font-weight:700;margin-bottom:8px;">Simulation Monte Carlo</div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Tester la sensibilite du resultat et de la tresorerie sur plusieurs scenarios.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('montecarlo')">Ouvrir la simulation</button>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div style="font-weight:700;margin-bottom:8px;">Reduction des couts</div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Transformer les charges OHADA en plan d'actions concret et mesurable.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('couts')">Ouvrir le plan d'actions</button>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div style="font-weight:700;margin-bottom:8px;">Declaration DSF</div>
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Generez LIASSE.xlsx et partagez-la directement depuis l'application.</div>
          <button class="btn btn-outline" style="width:100%;" onclick="navigateToTab('dsf')">Ouvrir la DSF</button>
        </div>
      </div>
    </div>

    ${(margePct !== null || liquidite !== null || endettement !== null) ? `
    <div class="kpi-grid" style="margin-top:0;">
      ${margePct !== null ? `<div class="kpi"><div class="kpi-label">Marge nette</div><div class="kpi-value" style="color:${parseFloat(margePct) >= 0 ? 'var(--green)' : 'var(--red)'}">${margePct}%</div><div class="kpi-note">${resultat >= 0 ? 'Benefice' : 'Perte'} ${fmt(Math.abs(resultat))} XOF</div></div>` : ''}
      ${liquidite !== null ? `<div class="kpi"><div class="kpi-label">Liquidite generale</div><div class="kpi-value" style="color:${parseFloat(liquidite) >= 1 ? 'var(--green)' : 'var(--red)'}">${liquidite}</div><div class="kpi-note">${parseFloat(liquidite) >= 1 ? 'Solvable' : 'Risque liquidite'}</div></div>` : ''}
      ${endettement !== null ? `<div class="kpi"><div class="kpi-label">Taux endettement</div><div class="kpi-value" style="color:${parseFloat(endettement) <= 100 ? 'var(--green)' : 'var(--red)'}">${endettement}%</div><div class="kpi-note">Capital social: ${fmt(capital)} XOF</div></div>` : ''}
      ${capital > 0 ? `<div class="kpi"><div class="kpi-label">Rentabilite CP</div><div class="kpi-value" style="color:${resultat >= 0 ? 'var(--green)' : 'var(--red)'}">${capital > 0 ? (resultat / capital * 100).toFixed(1) + '%' : '—'}</div><div class="kpi-note">Resultat / Capital</div></div>` : ''}
    </div>` : ''}

    <div class="grid-2">
      <div class="card">
        <div class="card-header"><div class="card-title">Classes comptables</div></div>
        ${SYSCOHADA_CLASSES.map(c => {
          const count = plan.filter(a => a.classe === c.num).length;
          return `<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid rgba(30,58,95,0.3);">
            <div style="display:flex;align-items:center;gap:10px;">
              <span class="class-badge class-${c.num}">${c.num}</span>
              <span style="font-size:0.84rem;">${c.label}</span>
            </div>
            <span style="font-family:var(--mono);font-size:0.82rem;color:var(--muted);">${count} comptes</span>
          </div>`;
        }).join("")}
      </div>
      <div class="card">
        <div class="card-header"><div class="card-title">Journaux</div></div>
        ${JOURNAL_CODES.map(j => {
          const count = journalEntries.filter(e => e.journal === j.code).length;
          return `<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid rgba(30,58,95,0.3);">
            <div style="display:flex;align-items:center;gap:10px;">
              <span style="font-family:var(--mono);font-weight:700;color:var(--gold);width:30px;">${j.code}</span>
              <span style="font-size:0.84rem;">${j.label}</span>
            </div>
            <span style="font-family:var(--mono);font-size:0.82rem;color:var(--muted);">${count}</span>
          </div>`;
        }).join("")}
      </div>
    </div>

    <div class="card" style="margin-top:16px;">
      <div class="card-header"><div class="card-title">Conformite OHADA</div></div>
      <div class="info-box">
        <strong>Referentiel actif: ${refLabel}</strong><br><br>
        ${currentRef === "sycebnl"
          ? "Le SYCEBNL s'applique aux entites a but non lucratif: associations, ONG, fondations, syndicats. Il inclut des comptes specifiques pour les fonds dedies (classe 1), les donateurs et bailleurs (classe 4), et les dons/contributions (classe 7)."
          : "Le SYSCOHADA revise s'applique a toutes les entites a but lucratif des 17 Etats membres de l'OHADA. Il est conforme a l'Acte Uniforme AUDCIF et structure en 8 classes de comptes."
        }<br><br>
        <strong>Etats financiers:</strong> Bilan | Compte de resultat | Tableau de flux de tresorerie (TFT) | Notes annexes<br>
        <strong>Pays membres:</strong> ${OHADA_MEMBER_STATES.join(", ")}<br>
        <strong>Choix entreprise:</strong> ce referentiel est maintenant memorise dans la fiche entreprise et dans le compte.
        <div style="margin-top:14px;">
          <button class="btn btn-outline" onclick="navigateToTab('veille')">Ouvrir la veille OHADA quotidienne</button>
        </div>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// PLAN COMPTABLE
// ═══════════════════════════════════════════════════════════
function renderPlan() {
  const plan = getPlan();
  const filtered = plan.filter(a => {
    if (filterClass && a.classe !== filterClass) return false;
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      return a.numero.includes(q) || a.libelle.toLowerCase().includes(q);
    }
    return true;
  });

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Plan comptable ${currentRef === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise"}</div>
          <div class="card-subtitle">${filtered.length} comptes affiches sur ${plan.length}</div>
        </div>
      </div>
      <input class="search-box" id="plan-search" placeholder="Rechercher par numero ou libelle..." value="${searchTerm}" style="margin-bottom:16px;">
      <div class="filter-row">
        <button class="filter-pill ${!filterClass ? 'active' : ''}" data-class="0">Toutes (${plan.length})</button>
        ${SYSCOHADA_CLASSES.map(c => {
          const count = plan.filter(a => a.classe === c.num).length;
          return `<button class="filter-pill ${filterClass === c.num ? 'active' : ''}" data-class="${c.num}">Cl. ${c.num} (${count})</button>`;
        }).join("")}
      </div>
      <div style="overflow-x:auto;">
        <table class="data-table">
          <thead><tr><th>Numero</th><th>Libelle</th><th>Classe</th><th>Type</th><th>Sens normal</th>${currentRef === "sycebnl" ? "<th>Spec.</th>" : ""}</tr></thead>
          <tbody>
            ${filtered.map(a => `
              <tr>
                <td class="code">${a.numero}</td>
                <td>${a.libelle}</td>
                <td><span class="class-badge class-${a.classe}">${a.classe}</span></td>
                <td style="font-size:0.78rem;color:var(--muted);">${a.type}</td>
                <td>${sensSpan(a.sens)}</td>
                ${currentRef === "sycebnl" ? `<td>${a.sycebnl ? '<span style="color:var(--cyan);font-weight:700;font-size:0.72rem;">SYCEBNL</span>' : ''}</td>` : ""}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// JOURNAL
// ═══════════════════════════════════════════════════════════
function renderJournal() {
  const totalD = journalEntries.reduce((s, e) => s + e.debit, 0);
  const totalC = journalEntries.reduce((s, e) => s + e.credit, 0);

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Journal general</div>
          <div class="card-subtitle">${journalEntries.length} ecritures | Equilibre: ${isBalancedAmount(totalD, totalC) ? 'OK' : 'ERREUR'}</div>
        </div>
        <button class="btn btn-gold" onclick="navigateToTab('saisie')">+ Nouvelle ecriture</button>
      </div>
      ${journalEntries.length === 0 ? `
        <div class="info-box">
          Le journal est vide. Saisissez votre premiere ecriture ou importez une balance d'ouverture pour demarrer sur un dossier de production.
          <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px;">
            <button class="btn btn-gold" onclick="navigateToTab('saisie')">Saisir une ecriture</button>
            <button class="btn btn-outline" onclick="navigateToTab('balance')">Importer une balance</button>
          </div>
        </div>
      ` : `
      <div style="overflow-x:auto;">
        <table class="data-table">
          <thead><tr><th>Date</th><th>Journal</th><th>Piece</th><th>Compte</th><th>Libelle</th><th style="text-align:right;">Debit</th><th style="text-align:right;">Credit</th><th>Ref</th></tr></thead>
          <tbody>
            ${journalEntries.map(e => `
              <tr>
                <td style="font-family:var(--mono);font-size:0.78rem;">${e.date}</td>
                <td><span style="font-family:var(--mono);font-weight:700;color:var(--gold);">${e.journal}</span></td>
                <td style="font-size:0.78rem;">${e.piece}</td>
                <td class="code">${e.compte}</td>
                <td>${e.libelle}</td>
                <td style="text-align:right;" class="${e.debit > 0 ? 'debit' : ''}">${e.debit > 0 ? fmt(e.debit) : ''}</td>
                <td style="text-align:right;" class="${e.credit > 0 ? 'credit' : ''}">${e.credit > 0 ? fmt(e.credit) : ''}</td>
                <td style="font-size:0.78rem;color:var(--dim);">${e.ref}</td>
              </tr>
            `).join("")}
            <tr style="border-top:2px solid var(--gold);font-weight:700;">
              <td colspan="5" style="text-align:right;color:var(--gold);">TOTAUX</td>
              <td style="text-align:right;" class="debit">${fmt(totalD)}</td>
              <td style="text-align:right;" class="credit">${fmt(totalC)}</td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
      `}
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// GRAND LIVRE
// ═══════════════════════════════════════════════════════════
function renderGrandLivre() {
  const bal = computeBalances();
  const plan = getPlan();
  const comptes = Object.keys(bal).sort();

  if (comptes.length === 0) {
    return `
      <div class="card">
        <div class="card-header">
          <div class="card-title">Grand livre</div>
        </div>
        <div class="info-box">
          Aucun compte n'a encore de solde ou de mouvement. Le grand livre se generera automatiquement apres l'import de la balance d'ouverture ou la saisie des ecritures.
        </div>
      </div>
    `;
  }

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Grand livre</div>
      </div>
      ${comptes.map(code => {
        const account = plan.find(a => a.numero === code);
        const entries = journalEntries.filter(e => e.compte === code);
        const d = bal[code].debit;
        const c = bal[code].credit;
        const solde = d - c;
        return `
          <div style="margin-bottom:20px;">
            <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:var(--surface2);border-radius:var(--radius);margin-bottom:4px;">
              <div><span class="code" style="margin-right:8px;">${code}</span> ${account ? account.libelle : 'Compte inconnu'}</div>
              <div style="font-family:var(--mono);font-size:0.82rem;">Solde: <span style="color:${solde >= 0 ? 'var(--green)' : 'var(--red)'};">${fmt(Math.abs(solde))} ${solde >= 0 ? 'D' : 'C'}</span></div>
            </div>
            <table class="data-table" style="font-size:0.8rem;">
              <thead><tr><th>Date</th><th>Libelle</th><th>Ref</th><th style="text-align:right;">Debit</th><th style="text-align:right;">Credit</th></tr></thead>
              <tbody>
                ${entries.map(e => `<tr>
                  <td style="font-family:var(--mono);">${e.date}</td>
                  <td>${e.libelle}</td>
                  <td style="color:var(--dim);">${e.ref}</td>
                  <td style="text-align:right;" class="${e.debit > 0 ? 'debit' : ''}">${e.debit > 0 ? fmt(e.debit) : ''}</td>
                  <td style="text-align:right;" class="${e.credit > 0 ? 'credit' : ''}">${e.credit > 0 ? fmt(e.credit) : ''}</td>
                </tr>`).join("")}
                <tr style="font-weight:700;border-top:1px solid var(--border);">
                  <td colspan="3" style="text-align:right;">Total</td>
                  <td style="text-align:right;" class="debit">${fmt(d)}</td>
                  <td style="text-align:right;" class="credit">${fmt(c)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// BALANCE
// ═══════════════════════════════════════════════════════════
function renderBalance() {
  const bal = computeBalances();
  const plan = getPlan();
  const comptes = Object.keys(bal).sort();
  let totN1D=0,totN1C=0,totMvtD=0,totMvtC=0,totSD=0,totSC=0;

  const rows = comptes.map(code => {
    const account = plan.find(a => a.numero === code);
    const n1d = bal[code].n1d||0, n1c = bal[code].n1c||0;
    const mvtD = bal[code].debit - n1d, mvtC = bal[code].credit - n1c;
    const solde = bal[code].debit - bal[code].credit;
    totN1D+=n1d; totN1C+=n1c; totMvtD+=mvtD; totMvtC+=mvtC;
    if (solde>0) totSD+=solde; else totSC+=Math.abs(solde);
    return {code, label:account?account.libelle:'Inconnu', classe:account?account.classe:0, n1d,n1c,mvtD,mvtC,solde};
  });

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Balance generale — ${currentRef==="sycebnl"?"SYCEBNL":"SYSCOHADA Revise"}</div>
          <div class="card-subtitle">${comptes.length} comptes | S.O.: ${isBalancedAmount(totN1D, totN1C)?'OK':'ERR'} | Mvt: ${isBalancedAmount(totMvtD, totMvtC)?'OK':'ERREUR'} | Solde: ${isBalancedAmount(totSD, totSC)?'OK':'ERREUR'}</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <label class="btn btn-outline" style="cursor:pointer;font-size:0.78rem;">
            Importer CSV <input type="file" accept=".csv" style="display:none;" onchange="importBalanceCsv(this)">
          </label>
          <button class="btn btn-outline" style="font-size:0.78rem;" onclick="shareBalanceTemplate()">Partager modele CSV</button>
          <button class="btn btn-outline" style="font-size:0.78rem;" onclick="downloadBalanceTemplate()">Modele CSV</button>
        </div>
      </div>
      ${comptes.length === 0 ? `
        <div class="info-box">
          La balance est vide. Importez un fichier CSV de soldes d'ouverture ou commencez par passer des ecritures pour alimenter automatiquement cette vue.
        </div>
      ` : `
      <div style="overflow-x:auto;">
        <table class="data-table">
          <thead>
            <tr>
              <th>Compte</th><th>Libelle</th><th>Cl.</th>
              <th style="text-align:right;color:var(--cyan);">S.O. Debit</th>
              <th style="text-align:right;color:var(--cyan);">S.O. Credit</th>
              <th style="text-align:right;">Mvt Debit</th>
              <th style="text-align:right;">Mvt Credit</th>
              <th style="text-align:right;color:var(--gold);">Solde D</th>
              <th style="text-align:right;color:var(--gold);">Solde C</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td class="code">${r.code}</td>
                <td>${r.label}</td>
                <td><span class="class-badge class-${r.classe}">${r.classe}</span></td>
                <td style="text-align:right;color:var(--cyan);font-size:0.82rem;">${r.n1d>0?fmt(r.n1d):''}</td>
                <td style="text-align:right;color:var(--cyan);font-size:0.82rem;">${r.n1c>0?fmt(r.n1c):''}</td>
                <td style="text-align:right;">${r.mvtD>0?fmt(r.mvtD):''}</td>
                <td style="text-align:right;">${r.mvtC>0?fmt(r.mvtC):''}</td>
                <td style="text-align:right;" class="debit">${r.solde>0?fmt(r.solde):''}</td>
                <td style="text-align:right;" class="credit">${r.solde<0?fmt(Math.abs(r.solde)):''}</td>
              </tr>
            `).join("")}
            <tr style="border-top:2px solid var(--gold);font-weight:700;font-size:0.82rem;">
              <td colspan="3" style="text-align:right;color:var(--gold);">TOTAUX</td>
              <td style="text-align:right;color:var(--cyan);">${fmt(totN1D)}</td>
              <td style="text-align:right;color:var(--cyan);">${fmt(totN1C)}</td>
              <td style="text-align:right;">${fmt(totMvtD)}</td>
              <td style="text-align:right;">${fmt(totMvtC)}</td>
              <td style="text-align:right;" class="debit">${fmt(totSD)}</td>
              <td style="text-align:right;" class="credit">${fmt(totSC)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      `}
    </div>
  `;
}

function importBalanceCsv(input) {
  const file = input.files[0];
  if (!file) return;
  handleDroppedFile(file);
}

function downloadBalanceTemplate() {
  const blob = buildBalanceTemplateBlob();
  downloadBlobFile(blob, "balance_ouverture_template.csv");
  showToast("Le modele balance_ouverture_template.csv a ete genere avec succes.", "success");
}

// ═══════════════════════════════════════════════════════════
// BILAN
// ═══════════════════════════════════════════════════════════
function renderBilan() {
  const { bal, actifRows, passifRows, resultatExercice, totalActif, totalPassif, balanceSummary, isBalanced } = getBilanSnapshot();
  const hasBalances = Object.keys(bal).length > 0;
  const bilanGap = Math.abs(totalActif - totalPassif);
  const sourceGap = Math.abs(balanceSummary.gap);
  const equilibriumLabel = isBalanced
    ? "OK ✓"
    : balanceSummary.isBalanced
      ? "ERREUR — ecart " + fmt(bilanGap)
      : "ERREUR — balance generale desequilibree de " + fmt(sourceGap);

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Bilan — ${currentRef === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise"}</div>
        <div class="card-subtitle" style="color:${isBalanced?'var(--green)':'var(--red)'};">Equilibre: ${equilibriumLabel}</div>
      </div>
      ${!hasBalances ? `<div class="info-box" style="margin-bottom:16px;">Aucune donnee comptable n'est encore disponible. Le bilan s'affichera automatiquement apres import des soldes d'ouverture ou saisie des ecritures.</div>` : ''}
      ${hasBalances && !balanceSummary.isBalanced ? `<div class="info-box" style="margin-bottom:16px;border-color:rgba(227,98,98,0.35);color:var(--red);">La balance generale n'est pas equilibree. Verifiez les soldes d'ouverture ou les imports d'ecritures avant de valider la liasse.</div>` : ''}
      <div class="grid-2">
        <div>
          <div class="section-title">ACTIF (valeurs nettes)</div>
          <table class="data-table">
            <thead><tr><th>Section</th><th style="text-align:right;">Net (XOF)</th></tr></thead>
            <tbody>
              ${actifRows.map(r => `<tr><td>${r.section}</td><td style="text-align:right;font-family:var(--mono);" class="debit">${fmt(Math.abs(r.total))}</td></tr>`).join("")}
              <tr style="font-weight:700;border-top:2px solid var(--gold);"><td style="color:var(--gold);">TOTAL ACTIF</td><td style="text-align:right;font-family:var(--mono);color:var(--gold);">${fmt(Math.abs(totalActif))}</td></tr>
            </tbody>
          </table>
        </div>
        <div>
          <div class="section-title">PASSIF</div>
          <table class="data-table">
            <thead><tr><th>Section</th><th style="text-align:right;">Montant (XOF)</th></tr></thead>
            <tbody>
              ${passifRows.map((r,i) => `<tr><td>${r.section}${i===0?' <span style="font-size:0.72rem;color:var(--muted);">(dont Resultat: '+fmt(Math.abs(resultatExercice))+' '+(resultatExercice>=0?'Excedent':'Deficit')+')</span>':''}</td><td style="text-align:right;font-family:var(--mono);" class="credit">${fmt(Math.abs(r.total))}</td></tr>`).join("")}
              <tr style="font-weight:700;border-top:2px solid var(--gold);"><td style="color:var(--gold);">TOTAL PASSIF</td><td style="text-align:right;font-family:var(--mono);color:var(--gold);">${fmt(Math.abs(totalPassif))}</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// COMPTE DE RESULTAT
// ═══════════════════════════════════════════════════════════
function renderResultat() {
  const bal = computeBalances();
  const hasBalances = Object.keys(bal).length > 0;

  if (currentRef === "sycebnl") {
    // SYCEBNL — structure officielle AUDCIF
    // Associations/ONG/Fondations: RA-RH (Revenus) / TA-TL (Charges) / XA/XB/XC/XD
    // Projets de developpement: RA-RE (Revenus) / TA-TN (Charges)
    const struct = sycebnlType === "projets" ? RESULTAT_SYCEBNL_PROJETS_STRUCTURE : RESULTAT_SYCEBNL_STRUCTURE;
    const entityLabel = sycebnlType === "projets" ? "Projets de developpement" : "Associations / ONG / Fondations";

    function sumRows(rows) {
      return rows.map(s => {
        let total = 0;
        Object.keys(bal).forEach(code => {
          if (s.comptes.some(p => code.startsWith(p))) {
            const mvtD = (bal[code].debit||0) - (bal[code].n1d||0);
            const mvtC = (bal[code].credit||0) - (bal[code].n1c||0);
            total += s.sens === "credit" ? mvtC - mvtD : mvtD - mvtC;
          }
        });
        return { ...s, total };
      });
    }

    const revRows = sumRows(struct.revenus);
    const chgRows = sumRows(struct.charges);
    const xA = revRows.reduce((s,r)=>s+r.total, 0);
    const xB = chgRows.filter(r=>r.sens==="debit").reduce((s,r)=>s+r.total, 0);
    const haoP = chgRows.filter(r=>r.sens==="credit").reduce((s,r)=>s+r.total, 0);
    const haoC = (struct.hao||[]).reduce((acc,s)=>{
      let t=0; Object.keys(bal).forEach(code=>{if(s.comptes.some(p=>code.startsWith(p))){const mvtD=(bal[code].debit||0)-(bal[code].n1d||0);const mvtC=(bal[code].credit||0)-(bal[code].n1c||0);t+=s.sens==="credit"?mvtC-mvtD:mvtD-mvtC;}});return{...s,total:t};
    }, []);
    const haoRows = struct.hao ? sumRows(struct.hao) : [];
    const xC = xA - xB;
    const xD = haoRows.filter(r=>r.sens==="credit").reduce((s,r)=>s+r.total,0) -
               haoRows.filter(r=>r.sens==="debit").reduce((s,r)=>s+r.total,0);
    const resultatNet = xC + xD;

    // Classe 9 contributions volontaires
    let cv9E=0, cv9R=0;
    Object.keys(bal).forEach(code => {
      const d=(bal[code].debit||0)-(bal[code].n1d||0), c=(bal[code].credit||0)-(bal[code].n1c||0);
      if (code.startsWith("91")) cv9E += d-c;
      if (code.startsWith("92")) cv9R += c-d;
    });

    return `
      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Compte de resultat — SYCEBNL</div>
            <div class="card-subtitle">${entityLabel}</div>
          </div>
          <div style="display:flex;gap:8px;">
            <button class="btn ${sycebnlType==="associations"?"btn-gold":"btn-outline"}" onclick="setSycebnlType('associations')">Associations</button>
            <button class="btn ${sycebnlType==="projets"?"btn-gold":"btn-outline"}" onclick="setSycebnlType('projets')">Projets</button>
          </div>
        </div>
        ${!hasBalances ? `<div class="info-box" style="margin-bottom:16px;">Aucun mouvement n'est encore disponible pour produire le compte de resultat. Saisissez vos ecritures de l'exercice pour lancer le calcul.</div>` : ''}

        <div class="section-title" style="color:var(--green);margin-top:8px;">Revenus des Activites Ordinaires</div>
        <table class="data-table" style="margin-bottom:8px;">
          <thead><tr><th style="width:50px;">REF</th><th>LIBELLES</th><th style="text-align:right;">N (XOF)</th></tr></thead>
          <tbody>
            ${revRows.map(r=>`<tr>
              <td class="code" style="color:var(--cyan);font-weight:700;">${r.ref}</td>
              <td>${r.label}</td>
              <td style="text-align:right;" class="credit">${r.total>0?fmt(r.total):'-'}</td>
            </tr>`).join("")}
            <tr style="font-weight:700;border-top:2px solid var(--green);background:rgba(0,230,118,0.06);">
              <td style="color:var(--green);">XA</td>
              <td style="color:var(--green);">REVENUS DES ACTIVITES ORDINAIRES</td>
              <td style="text-align:right;" class="credit">${fmt(xA)}</td>
            </tr>
          </tbody>
        </table>

        <div class="section-title" style="color:var(--red);">Charges des Activites Ordinaires</div>
        <table class="data-table" style="margin-bottom:8px;">
          <thead><tr><th style="width:50px;">REF</th><th>LIBELLES</th><th style="text-align:right;">N (XOF)</th></tr></thead>
          <tbody>
            ${chgRows.filter(r=>r.sens==="debit").map(r=>`<tr>
              <td class="code" style="color:var(--orange);font-weight:700;">${r.ref}</td>
              <td>${r.label}</td>
              <td style="text-align:right;" class="debit">${r.total>0?fmt(r.total):'-'}</td>
            </tr>`).join("")}
            <tr style="font-weight:700;border-top:2px solid var(--red);background:rgba(255,82,82,0.06);">
              <td style="color:var(--red);">XB</td>
              <td style="color:var(--red);">CHARGES DES ACTIVITES ORDINAIRES</td>
              <td style="text-align:right;" class="debit">${fmt(xB)}</td>
            </tr>
          </tbody>
        </table>

        <table class="data-table" style="margin-bottom:8px;">
          <tbody>
            <tr style="font-weight:700;border-top:2px solid var(--gold);font-size:1rem;">
              <td style="width:50px;color:var(--gold);">XC</td>
              <td style="color:var(--gold);">RESULTAT DES ACTIVITES ORDINAIRES (XA - XB)</td>
              <td style="text-align:right;color:${xC>=0?'var(--green)':'var(--red)'};font-family:var(--mono);font-weight:700;">${fmt(xC)}</td>
            </tr>
            ${haoRows.length>0?haoRows.map(r=>`<tr>
              <td class="code" style="color:var(--muted);">${r.ref}</td><td>${r.label}</td>
              <td style="text-align:right;" class="${r.sens==="credit"?"credit":"debit"}">${r.total>0?fmt(r.total):'-'}</td>
            </tr>`).join(""):""}
            ${haoRows.length>0?`<tr style="font-weight:700;"><td style="color:var(--muted);">XD</td><td style="color:var(--muted);">RESULTAT HAO (TM - TN)</td><td style="text-align:right;">${fmt(xD)}</td></tr>`:""}
            <tr style="font-weight:700;border-top:3px solid var(--gold);font-size:1.05rem;background:rgba(200,146,42,0.08);">
              <td style="color:var(--gold);">XC</td>
              <td style="color:var(--gold);">SOLDE DES OPERATIONS — ${resultatNet>=0?'EXCEDENT':'DEFICIT'}</td>
              <td style="text-align:right;font-family:var(--mono);color:${resultatNet>=0?'var(--green)':'var(--red)'};">${fmt(Math.abs(resultatNet))} XOF</td>
            </tr>
          </tbody>
        </table>

        ${cv9E>0||cv9R>0?`
        <div class="section-title" style="color:var(--cyan);">Classe 9 — Contributions volontaires en nature</div>
        <table class="data-table">
          <tbody>
            <tr><td class="code" style="color:var(--cyan);">91x</td><td>Emplois valorises</td><td style="text-align:right;" class="debit">${fmt(cv9E)}</td></tr>
            <tr><td class="code" style="color:var(--cyan);">92x</td><td>Ressources valorisees</td><td style="text-align:right;" class="credit">${fmt(cv9R)}</td></tr>
          </tbody>
        </table>`:""}
      </div>
    `;
  }

  // SYSCOHADA standard
  let totalProduits = 0, totalCharges = 0;
  const rows = RESULTAT_STRUCTURE.map(s => {
    let total = 0;
    Object.keys(bal).forEach(code => {
      if (s.comptes.some(p => code.startsWith(p))) {
        const mvtD=(bal[code].debit||0)-(bal[code].n1d||0);
        const mvtC=(bal[code].credit||0)-(bal[code].n1c||0);
        if (s.sens === "credit") total += mvtC - mvtD;
        else total += mvtD - mvtC;
      }
    });
    if (s.sens === "credit") totalProduits += total; else totalCharges += total;
    return { section: s.section, total, sens: s.sens };
  });
  const resultat = totalProduits - totalCharges;

  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Compte de resultat — SYSCOHADA Revise</div></div>
      ${!hasBalances ? `<div class="info-box" style="margin-bottom:16px;">Aucun mouvement n'est encore disponible pour produire le compte de resultat. Saisissez vos ecritures de l'exercice pour lancer le calcul.</div>` : ''}
      <table class="data-table">
        <thead><tr><th>Section</th><th>Type</th><th style="text-align:right;">Montant (XOF)</th></tr></thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              <td>${r.section}</td>
              <td style="font-size:0.78rem;"><span class="${r.sens==='credit'?'sens-credit':'sens-debit'}">${r.sens==='credit'?'Produit':'Charge'}</span></td>
              <td style="text-align:right;font-family:var(--mono);" class="${r.sens==='credit'?'credit':'debit'}">${fmt(Math.abs(r.total))}</td>
            </tr>
          `).join("")}
          <tr style="font-weight:700;border-top:2px solid var(--gold);">
            <td>TOTAL PRODUITS</td><td></td><td style="text-align:right;" class="credit">${fmt(totalProduits)}</td>
          </tr>
          <tr style="font-weight:700;">
            <td>TOTAL CHARGES</td><td></td><td style="text-align:right;" class="debit">${fmt(totalCharges)}</td>
          </tr>
          <tr style="font-weight:700;border-top:2px solid var(--gold);font-size:1.1rem;">
            <td style="color:var(--gold);">RESULTAT NET</td><td></td>
            <td style="text-align:right;color:${resultat>=0?'var(--green)':'var(--red)'};">${
resultat>=0?'Benefice':'Perte'}: ${fmt(Math.abs(resultat))} XOF</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}
// ═══════════════════════════════════════════════════════════
// FLUX DE TRESORERIE (legacy tab key: tafire)
// ═══════════════════════════════════════════════════════════
function renderTafire() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Tableau de flux de tresorerie (TFT)</div></div>
      <div class="info-box">
        Dans le SYSCOHADA revise, le <strong>Tableau de flux de tresorerie (TFT)</strong> remplace le TAFIRE dans le jeu d'etats financiers du systeme normal.<br><br>
        Cette page sert de base preparatoire pour suivre les flux de tresorerie de l'exercice avant production du tableau detaille conforme.<br><br>
        <em>Le calcul automatique complet du TFT sera affine apres cloture de l'exercice avec les donnees completes.</em>
      </div>
      <div class="grid-2" style="margin-top:16px;">
        <div class="card" style="border-color:var(--green);">
          <div style="color:var(--green);font-weight:700;margin-bottom:8px;">FLUX POSITIFS</div>
          <div style="font-size:0.84rem;color:var(--muted);line-height:1.8;">
            Flux de tresorerie lies aux activites operationnelles<br>
            Encaissements sur cessions d'immobilisations<br>
            Flux de financement recus<br>
            Augmentation nette de tresorerie<br>
          </div>
        </div>
        <div class="card" style="border-color:var(--red);">
          <div style="color:var(--red);font-weight:700;margin-bottom:8px;">FLUX NEGATIFS</div>
          <div style="font-size:0.84rem;color:var(--muted);line-height:1.8;">
            Decaissements d'exploitation<br>
            Investissements (acquisitions d'immobilisations)<br>
            Remboursements de dettes financieres<br>
            Diminution nette de tresorerie<br>
          </div>
        </div>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// TABLEAU DES AMORTISSEMENTS
// ═══════════════════════════════════════════════════════════
function renderAmortissements() {
  const bal = computeBalances();
  const plan = getPlan();
  const immos = plan.filter(a => a.classe===2 && a.sens==="Debit" && bal[a.numero]);
  const durees = {"21":5,"211":5,"212":5,"213":10,"214":10,"22":0,"23":20,"231":20,"232":20,
    "234":10,"235":7,"24":5,"241":5,"242":5,"244":3,"245":5,"246":7,"248":5};
  function amortCode(c){if(c.startsWith("21"))return"281";if(c.startsWith("22"))return"282";if(c.startsWith("23"))return"283";return"284";}

  const rows = immos.map(a => {
    const vbo = bal[a.numero].debit;
    if (!vbo) return null;
    const ac = amortCode(a.numero);
    const cumulN1 = ac && bal[ac] ? (bal[ac].n1c||0) : 0;
    const duree = durees[a.numero]||5;
    const taux = duree>0 ? Math.round(10000/duree)/100 : 0;
    const dotation = duree>0 ? Math.round(vbo/duree) : 0;
    const cumulN = cumulN1 + dotation;
    const vnc = Math.max(0, vbo - cumulN);
    return {code:a.numero,label:a.libelle,vbo,duree,taux,cumulN1,dotation,cumulN,vnc};
  }).filter(r=>r&&r.vbo>0);

  const totVBO=rows.reduce((s,r)=>s+r.vbo,0);
  const totDot=rows.reduce((s,r)=>s+r.dotation,0);
  const totCN1=rows.reduce((s,r)=>s+r.cumulN1,0);
  const totCN=rows.reduce((s,r)=>s+r.cumulN,0);
  const totVNC=rows.reduce((s,r)=>s+r.vnc,0);

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Tableau des amortissements des immobilisations</div>
          <div class="card-subtitle">${rows.length} immobilisations</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn btn-outline" style="font-size:0.78rem;" onclick="shareAmortCsv()">Partager CSV</button>
          <button class="btn btn-outline" style="font-size:0.78rem;" onclick="exportAmortCsv()">Export CSV</button>
        </div>
      </div>
      ${rows.length===0?'<div class="info-box">Aucune immobilisation mouvementee dans le journal. Saisissez des ecritures sur les comptes de classe 2.</div>':''}
      ${rows.length>0?`
      <div style="overflow-x:auto;">
        <table class="data-table">
          <thead><tr>
            <th>Compte</th><th>Libelle</th>
            <th style="text-align:right;">V.B.O.</th>
            <th style="text-align:center;">Duree</th>
            <th style="text-align:center;">Taux</th>
            <th style="text-align:right;color:var(--cyan);">Cumul N-1</th>
            <th style="text-align:right;color:var(--gold);">Dotation N</th>
            <th style="text-align:right;color:var(--orange);">Cumul N</th>
            <th style="text-align:right;color:var(--green);">V.N.C.</th>
          </tr></thead>
          <tbody>
            ${rows.map(r=>`<tr>
              <td class="code">${r.code}</td><td>${r.label}</td>
              <td style="text-align:right;font-family:var(--mono);">${fmt(r.vbo)}</td>
              <td style="text-align:center;">${r.duree>0?r.duree+' ans':'N/A'}</td>
              <td style="text-align:center;font-family:var(--mono);">${r.taux>0?r.taux+'%':'—'}</td>
              <td style="text-align:right;color:var(--cyan);font-family:var(--mono);">${fmt(r.cumulN1)}</td>
              <td style="text-align:right;color:var(--gold);font-family:var(--mono);font-weight:700;">${fmt(r.dotation)}</td>
              <td style="text-align:right;color:var(--orange);font-family:var(--mono);">${fmt(r.cumulN)}</td>
              <td style="text-align:right;font-family:var(--mono);" class="debit">${fmt(r.vnc)}</td>
            </tr>`).join("")}
            <tr style="border-top:2px solid var(--gold);font-weight:700;">
              <td colspan="2" style="color:var(--gold);">TOTAUX</td>
              <td style="text-align:right;">${fmt(totVBO)}</td>
              <td colspan="2"></td>
              <td style="text-align:right;color:var(--cyan);">${fmt(totCN1)}</td>
              <td style="text-align:right;color:var(--gold);">${fmt(totDot)}</td>
              <td style="text-align:right;color:var(--orange);">${fmt(totCN)}</td>
              <td style="text-align:right;">${fmt(totVNC)}</td>
            </tr>
          </tbody>
        </table>
      </div>`:''}
    </div>
  `;
}

function exportAmortCsv() {
  const blob = buildAmortCsvBlob();
  downloadBlobFile(blob, "amortissements.csv");
  showToast("Le fichier amortissements.csv a ete genere avec succes.", "success");
}

// ═══════════════════════════════════════════════════════════
// NOTES ANNEXES
// ═══════════════════════════════════════════════════════════
function renderAnnexes() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Notes annexes aux etats financiers</div></div>
      <div class="info-box" style="margin-bottom:16px;">
        Les notes annexes font partie integrante des etats financiers OHADA. Pour le systeme normal du SYSCOHADA revise, l'elaboration complete comprend quarante-six (46) tableaux en notes annexes. La liste ci-dessous constitue une base preparatoire a enrichir pour atteindre le dossier complet.
      </div>
      <div class="stack">
        ${[
          "Note 1 — Regles et methodes comptables",
          "Note 2 — Immobilisations incorporelles et corporelles",
          "Note 3 — Tableau des amortissements",
          "Note 4 — Immobilisations financieres",
          "Note 5 — Stocks et en-cours",
          "Note 6 — Creances et dettes",
          "Note 7 — Tresorerie",
          "Note 8 — Capitaux propres",
          "Note 9 — Emprunts et dettes financieres",
          "Note 10 — Charges de personnel",
          "Note 11 — Engagements hors bilan",
          "Note 12 — Evenements posterieurs a la cloture",
          "Note 13 — Parties liees",
          "Note 14 — Informations fiscales (DSF)",
        ].map(n => `<div class="card" style="padding:12px 16px;"><span style="color:var(--gold);font-weight:600;">${n}</span></div>`).join("")}
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// SAISIE D'ECRITURES
// ═══════════════════════════════════════════════════════════
function renderSaisie() {
  const plan = getPlan();
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Saisie d'ecritures comptables</div></div>
      <div class="grid-2">
        <div class="form-group"><div class="form-label">Date</div><input class="form-input" type="date" id="saisie-date" value="${new Date().toISOString().split('T')[0]}"></div>
        <div class="form-group"><div class="form-label">Journal</div><select class="form-select" id="saisie-journal">${JOURNAL_CODES.map(j => `<option value="${j.code}">${j.code} — ${j.label}</option>`).join("")}</select></div>
      </div>
      <div style="margin-top:12px;" class="form-group"><div class="form-label">Libelle</div><input class="form-input" id="saisie-libelle" placeholder="Description de l'operation"></div>
      <div style="margin-top:12px;" class="form-group"><div class="form-label">Reference piece</div><input class="form-input" id="saisie-ref" placeholder="Ex: FA-201, REC-005"></div>

      <div class="section-title" style="margin-top:20px;">Lignes d'ecritures</div>
      <div id="saisie-lignes">
        <div class="grid-3" style="margin-bottom:8px;">
          <div class="form-group"><div class="form-label">Compte</div><select class="form-select saisie-compte">${plan.map(a => `<option value="${a.numero}">${a.numero} — ${a.libelle}</option>`).join("")}</select></div>
          <div class="form-group"><div class="form-label">Debit (XOF)</div><input class="form-input saisie-debit" type="number" placeholder="0"></div>
          <div class="form-group"><div class="form-label">Credit (XOF)</div><input class="form-input saisie-credit" type="number" placeholder="0"></div>
        </div>
        <div class="grid-3" style="margin-bottom:8px;">
          <div class="form-group"><select class="form-select saisie-compte">${plan.map(a => `<option value="${a.numero}">${a.numero} — ${a.libelle}</option>`).join("")}</select></div>
          <div class="form-group"><input class="form-input saisie-debit" type="number" placeholder="0"></div>
          <div class="form-group"><input class="form-input saisie-credit" type="number" placeholder="0"></div>
        </div>
      </div>
      <div style="margin-top:16px;display:flex;gap:10px;">
        <button class="btn btn-gold" id="btn-valider-ecriture">Valider l'ecriture</button>
        <button class="btn btn-outline" id="btn-ajouter-ligne">+ Ajouter une ligne</button>
      </div>
      <div id="saisie-msg" style="margin-top:12px;"></div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// CLOTURE
// ═══════════════════════════════════════════════════════════
function renderCloture() {
  const snapshot = getClotureSnapshot();
  const history = Array.isArray(currentCompanyDetails.closureHistory) ? currentCompanyDetails.closureHistory : [];
  const statusRows = [
    {
      label: "Fiche entreprise",
      ok: snapshot.profileComplete,
      detail: snapshot.profileComplete ? "Informations minimales renseignees." : "Completez la raison sociale, le NIF, le pays et les dates d'exercice."
    },
    {
      label: "Periode d'exercice",
      ok: snapshot.validDates,
      detail: snapshot.validDates
        ? `${formatDateValue(snapshot.exerciceDu)} au ${formatDateValue(snapshot.exerciceAu)}`
        : "Dates d'exercice invalides ou absentes."
    },
    {
      label: "Balance generale",
      ok: snapshot.balanceSummary.isBalanced,
      detail: snapshot.balanceSummary.isBalanced
        ? `Debit = Credit (${fmt(snapshot.balanceSummary.totalDebit)} XOF)`
        : `Ecart cumule: ${fmt(Math.abs(snapshot.balanceSummary.gap))} XOF`
    },
    {
      label: "A-nouveaux de l'exercice suivant",
      ok: snapshot.carryforwardSummary.isBalanced,
      detail: `${snapshot.carryforwardSummary.count} comptes reportes | Resultat: ${snapshot.resultatExercice >= 0 ? "benefice" : "perte"} ${fmt(Math.abs(snapshot.resultatExercice))} XOF`
    }
  ];

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Cloture de l'exercice</div>
          <div class="card-subtitle">${snapshot.exerciceDu && snapshot.exerciceAu ? `Exercice courant: ${formatDateValue(snapshot.exerciceDu)} au ${formatDateValue(snapshot.exerciceAu)}` : "Periode d'exercice a definir dans les parametres."}</div>
        </div>
        <button
          class="btn btn-gold"
          onclick="cloturerExercice()"
          ${snapshot.canClose ? "" : "disabled"}
          style="${snapshot.canClose ? "" : "opacity:0.6;cursor:not-allowed;"}">Cloturer l'exercice</button>
      </div>
      <div class="info-box" style="margin-bottom:16px;">
        La cloture de l'exercice comptable OHADA comprend les etapes suivantes:<br><br>
        1. <strong>Inventaire physique</strong> — Verification des stocks, immobilisations, tresorerie<br>
        2. <strong>Ecritures de regularisation</strong> — Amortissements, provisions, charges constatees d'avance<br>
        3. <strong>Balance apres inventaire</strong> — Verification de l'equilibre<br>
        4. <strong>Determination du resultat</strong> — Solde des comptes de gestion<br>
        5. <strong>Etablissement des etats financiers</strong> — Bilan, Resultat, Tableau de flux de tresorerie, Annexes<br>
        6. <strong>Ecritures de cloture</strong> — A-nouveaux pour l'exercice suivant<br>
        7. <strong>Declaration DSF/DGI</strong> — Liasse fiscale obligatoire
      </div>
      <div class="stack">
        ${statusRows.map((row) => `
          <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);gap:12px;">
            <div>
              <div style="font-weight:700;color:${row.ok ? "var(--green)" : "var(--orange)"};">${row.label}</div>
              <div style="font-size:0.8rem;color:var(--muted);margin-top:4px;">${row.detail}</div>
            </div>
            <div style="font-size:0.76rem;font-weight:700;color:${row.ok ? "var(--green)" : "var(--orange)"};">${row.ok ? "PRET" : "A VERIFIER"}</div>
          </div>
        `).join("")}
      </div>
      <div class="grid-2" style="margin-top:16px;">
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Impact de la cloture</div>
          <div style="font-size:0.86rem;color:var(--muted);line-height:1.7;">
            Les comptes des classes 1 a 5 seront reportes en a-nouveaux.<br>
            Les ecritures du journal courant seront archivees en historique de cloture puis remises a zero.<br>
            Le resultat de l'exercice sera reporte en <strong>${snapshot.resultatExercice >= 0 ? "121 - Report a nouveau crediteur" : "129 - Report a nouveau debiteur"}</strong>.<br>
            ${snapshot.nextExerciceDu && snapshot.nextExerciceAu ? `Le prochain exercice sera positionne du <strong>${formatDateValue(snapshot.nextExerciceDu)}</strong> au <strong>${formatDateValue(snapshot.nextExerciceAu)}</strong>.` : "Le prochain exercice ne peut pas encore etre calcule sans dates valides."}
          </div>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Historique recent</div>
          ${history.length === 0 ? `<div style="font-size:0.84rem;color:var(--muted);">Aucune cloture n'a encore ete enregistree.</div>` : `
            <div class="stack">
              ${history.slice(0, 3).map((item) => `
                <div style="padding:10px 12px;border:1px solid rgba(30,58,95,0.45);border-radius:var(--radius);">
                  <div style="font-weight:700;color:var(--gold);">${item.exerciseLabel || "Exercice cloture"}</div>
                  <div style="font-size:0.8rem;color:var(--muted);margin-top:4px;">Cloture le ${formatDateTimeValue(item.closedAt)} | Resultat ${fmt(Math.abs(item.resultatExercice || 0))} XOF ${item.resultatExercice >= 0 ? "benefice" : "perte"}</div>
                </div>
              `).join("")}
            </div>
          `}
        </div>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// DSF / DGI
// ═══════════════════════════════════════════════════════════
function renderPrevisionnel() {
  const snapshot = getForecastTemplateSnapshot();
  const projectName = snapshot.projectName || "Projet OHADA Compta";
  const totalRevenue = snapshot.merchandiseRevenue + snapshot.serviceRevenue;
  const totalSetup = snapshot.intangibleSetup + snapshot.realEstateSetup + snapshot.worksSetup + snapshot.equipmentSetup + snapshot.officeEquipmentSetup;
  const projectedYear2Revenue = Math.round(totalRevenue * (1 + snapshot.growthYear2));
  const projectedYear3Revenue = Math.round(projectedYear2Revenue * (1 + snapshot.growthYear3));
  const profileReady = isCompanyProfileComplete();
  const accountingReady = totalRevenue > 0 || totalSetup > 0 || snapshot.startingCash > 0 || snapshot.openingStock > 0;
  const checklist = [
    {
      label: "Fiche entreprise",
      ready: profileReady,
      hint: profileReady ? "Identite et exercice disponibles." : "Completez raison sociale, pays, NIF, siege et exercice."
    },
    {
      label: "Base comptable",
      ready: accountingReady,
      hint: accountingReady ? "Des soldes ou mouvements existent deja." : "Ajoutez une balance d'ouverture ou des ecritures."
    },
    {
      label: "Hypotheses exportables",
      ready: snapshot.assumptionsCount >= 4,
      hint: snapshot.assumptionsCount >= 4 ? "Le modele peut etre pre-rempli avec une base utile." : "Renseignez plus de donnees pour un export plus riche."
    }
  ];
  const totals = [
    { label: "Chiffre d'affaires annualise", value: totalRevenue, note: "Annee 1" },
    { label: "Investissements de depart", value: totalSetup, note: "Immobilisations" },
    { label: "Tresorerie et stock de depart", value: snapshot.startingCash + snapshot.openingStock, note: "Lancement" },
    { label: "Projection de CA annee 3", value: projectedYear3Revenue, note: "Scenario calcule" }
  ];
  const sourceRows = [
    ["Identite du projet", projectName, "Fiche entreprise"],
    ["Forme juridique exportee", snapshot.legalStatus, "Parametres entreprise"],
    ["Activite retenue", snapshot.salesType, "Activite principale / comptes 70"],
    ["CA marchandises", fmt(snapshot.merchandiseRevenue), "Comptes 701 a 704"],
    ["CA services", fmt(snapshot.serviceRevenue), "Comptes 706 a 707"],
    ["Salaires annee 1", fmt(snapshot.employeeYear1), "Comptes 661 a 663"]
  ];
  const monthlyRows = snapshot.merchandiseMonthly.map((amount, index) => ({
    month: index + 1,
    merchandise: amount,
    service: snapshot.serviceMonthly[index] || 0,
    total: amount + (snapshot.serviceMonthly[index] || 0)
  })).slice(0, 6);

  return `
    <div class="stack">
      <div class="card" style="border-color:rgba(200,146,42,0.32);background:linear-gradient(180deg, rgba(200,146,42,0.08), rgba(10,22,40,0.96));">
        <div class="card-header">
          <div>
            <div class="card-title">Plan financier previsionnel</div>
            <div class="card-subtitle">Module distinct du tableau de bord pour preparer puis exporter ${EXACT_FORECAST_DOWNLOAD_NAME}.</div>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button class="btn btn-outline" onclick="navigateToTab('parametres')">Completer la fiche</button>
            <button class="btn btn-outline" onclick="navigateToTab('balance')">Verifier la balance</button>
            <button class="btn btn-gold" onclick="downloadForecastWorkbook()">Telecharger le modele rempli</button>
          </div>
        </div>
        <div class="grid-3">
          ${checklist.map((item) => `
            <div class="card" style="background:var(--surface2);padding:16px;">
              <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;">
                <strong>${item.label}</strong>
                <span style="font-size:0.74rem;font-weight:700;color:${item.ready ? "var(--green)" : "var(--orange)"};">${item.ready ? "Pret" : "A revoir"}</span>
              </div>
              <div style="margin-top:10px;font-size:0.84rem;color:var(--muted);line-height:1.6;">${item.hint}</div>
            </div>
          `).join("")}
        </div>
      </div>

      <div class="kpi-grid" style="margin-bottom:0;">
        ${totals.map((item) => `
          <div class="kpi">
            <div class="kpi-label">${item.label}</div>
            <div class="kpi-value" style="font-size:1rem;">${fmt(item.value)}</div>
            <div class="kpi-note">${item.note} · XOF</div>
          </div>
        `).join("")}
      </div>

      <div class="grid-2">
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Vue de synthese</div>
          <div style="font-size:0.85rem;color:var(--muted);line-height:1.85;">
            Projet: <strong>${projectName}</strong><br>
            Statut juridique exporte: <strong>${snapshot.legalStatus}</strong><br>
            Activite retenue: <strong>${snapshot.salesType}</strong><br>
            Croissance annee 2: <strong>${(snapshot.growthYear2 * 100).toFixed(0)}%</strong><br>
            Croissance annee 3: <strong>${(snapshot.growthYear3 * 100).toFixed(0)}%</strong><br>
            Cout d'achat des marchandises: <strong>${(snapshot.purchaseRatio * 100).toFixed(0)}%</strong><br>
            Salaires employes annee 1 a 3: <strong>${fmt(snapshot.employeeYear1)}</strong> / <strong>${fmt(snapshot.employeeYear2)}</strong> / <strong>${fmt(snapshot.employeeYear3)}</strong>
          </div>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Actions du module</div>
          <div style="font-size:0.85rem;color:var(--muted);line-height:1.8;">
            1. Ce module collecte les hypotheses a partir de vos soldes et de la fiche entreprise.<br>
            2. Il prepare le fichier exact <strong>${EXACT_FORECAST_DOWNLOAD_NAME}</strong>.<br>
            3. Le classeur exporte reste modifiable dans Excel apres telechargement.<br><br>
            <button class="btn btn-outline" onclick="shareForecastWorkbook()">Partager le modele</button>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Donnees injectees dans le modele</div>
            <div class="card-subtitle">Apercu des informations reprises avant export Excel.</div>
          </div>
        </div>
        <div style="overflow-x:auto;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Rubrique</th>
                <th>Valeur</th>
                <th>Source</th>
              </tr>
            </thead>
            <tbody>
              ${sourceRows.map((row) => `
                <tr>
                  <td>${row[0]}</td>
                  <td>${row[1]}</td>
                  <td>${row[2]}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Projection mensuelle N+1</div>
              <div class="card-subtitle">Apercu des 6 premiers mois utilises pour le pre-remplissage.</div>
            </div>
          </div>
          <div style="overflow-x:auto;">
            <table class="data-table">
              <thead>
                <tr>
                  <th>Mois</th>
                  <th>Marchandises</th>
                  <th>Services</th>
                  <th>Total</th>
                </tr>
              </thead>
              <tbody>
                ${monthlyRows.map((row) => `
                  <tr>
                    <td>M${row.month}</td>
                    <td>${fmt(row.merchandise)}</td>
                    <td>${fmt(row.service)}</td>
                    <td>${fmt(row.total)}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Controle avant export</div>
              <div class="card-subtitle">Ce qu'il faut verifier avant de partager le fichier.</div>
            </div>
          </div>
          <div class="stack">
            <div class="info-box">
              <strong>Modele cible:</strong> ${EXACT_FORECAST_DOWNLOAD_NAME}<br><br>
              Verifiez les investissements de demarrage ligne par ligne, la remuneration du dirigeant, l'ACRE, les delais clients/fournisseurs et les hypotheses mensuelles de vente.
            </div>
            <div class="card" style="background:var(--surface2);">
              <div style="font-size:0.84rem;color:var(--muted);line-height:1.8;">
                CA annee 1: <strong>${fmt(totalRevenue)}</strong> XOF<br>
                CA projete annee 2: <strong>${fmt(projectedYear2Revenue)}</strong> XOF<br>
                CA projete annee 3: <strong>${fmt(projectedYear3Revenue)}</strong> XOF<br>
                Besoin de lancement estime: <strong>${fmt(totalSetup + snapshot.startingCash + snapshot.openingStock)}</strong> XOF
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderMonteCarlo() {
  const config = getMonteCarloConfig();
  const base = getMonteCarloBaseScenario();
  const lastRun = currentCompanyDetails.monteCarloLastRun || null;
  const topCards = lastRun ? [
    {
      label: "Resultat net median",
      value: `${fmt(Math.round(lastRun.metrics.netIncome.p50))} XOF`,
      note: `P10 ${fmt(Math.round(lastRun.metrics.netIncome.p10))} | P90 ${fmt(Math.round(lastRun.metrics.netIncome.p90))}`,
      color: lastRun.metrics.netIncome.p50 >= 0 ? "var(--green)" : "var(--red)"
    },
    {
      label: "Free cash-flow median",
      value: `${fmt(Math.round(lastRun.metrics.freeCashFlow.p50))} XOF`,
      note: `P10 ${fmt(Math.round(lastRun.metrics.freeCashFlow.p10))} | P90 ${fmt(Math.round(lastRun.metrics.freeCashFlow.p90))}`,
      color: lastRun.metrics.freeCashFlow.p50 >= 0 ? "var(--green)" : "var(--red)"
    },
    {
      label: "Probabilite de perte",
      value: fmtPercent(lastRun.risk.lossProbability, 1),
      note: `Iterations ${fmt(lastRun.iterations)} | lancee le ${formatDateTimeValue(lastRun.generatedAt)}`,
      color: lastRun.risk.lossProbability <= 0.25 ? "var(--green)" : "var(--orange)"
    },
    {
      label: "Probabilite de cash negatif",
      value: fmtPercent(lastRun.risk.negativeCashProbability, 1),
      note: `Marge EBITDA mediane ${fmtPercent(lastRun.metrics.operatingMargin.p50, 1)}`,
      color: lastRun.risk.negativeCashProbability <= 0.35 ? "var(--green)" : "var(--orange)"
    }
  ] : [
    {
      label: "CA de base",
      value: `${fmt(config.baseRevenue)} XOF`,
      note: "Reference courante utilisee pour la simulation",
      color: "var(--gold)"
    },
    {
      label: "Charges fixes",
      value: `${fmt(config.baseFixedCosts)} XOF`,
      note: "Structure de couts hors cout variable",
      color: "var(--text)"
    },
    {
      label: "Marge brute de reference",
      value: fmtPercent(base.grossMargin, 1),
      note: "Base calculee a partir des ventes et achats",
      color: "var(--green)"
    },
    {
      label: "BFR cible",
      value: fmtPercent(base.workingCapitalPct, 1),
      note: `Tresorerie de depart ${fmt(base.openingCash || 0)} XOF`,
      color: "var(--cyan)"
    }
  ];

  const sourceRows = [
    ["Chiffre d'affaires de base", fmt(base.revenue), "Comptes 70 / plan financier"],
    ["Achats directs retenus", fmt(base.directCosts), "Comptes 60"],
    ["Charges fixes retenues", fmt(base.fixedCosts), "Charges totales hors achats directs"],
    ["Amortissement annuel", fmt(base.depreciation), "Comptes 68 ou stock d'immobilisations"],
    ["Capex de reference", fmt(base.capex), "Stock d'actifs + niveau d'investissement"],
    ["BFR de reference", `${fmt(base.workingCapitalBase)} XOF`, "Stocks + clients - dettes court terme"]
  ];
  const percentileRows = lastRun ? [
    ["Chiffre d'affaires", lastRun.metrics.revenue],
    ["EBITDA", lastRun.metrics.ebitda],
    ["Resultat net", lastRun.metrics.netIncome],
    ["Free cash-flow", lastRun.metrics.freeCashFlow]
  ] : [];

  function renderInput(field, label, suffix = "") {
    return `
      <div class="form-group">
        <div class="form-label">${label}</div>
        <input class="form-input" id="mc-${field}" type="number" step="0.01" value="${escapeHtml(getMonteCarloInputDisplayValue(field, config))}" placeholder="0">
        ${suffix ? `<div class="field-help">${suffix}</div>` : ""}
      </div>
    `;
  }

  return `
    <div class="stack">
      <div class="card feature-gradient-card">
        <div class="card-header">
          <div>
            <div class="card-title">Simulation Monte Carlo</div>
            <div class="card-subtitle">Module inspire du raisonnement Crystal Ball pour tester le resultat, la marge et la tresorerie avec plusieurs milliers de scenarios.</div>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button class="btn btn-outline" onclick="applyMonteCarloAccountingBase()">Recharger depuis la compta</button>
            <button class="btn btn-outline" onclick="resetMonteCarloModule()">Reinitialiser</button>
            <button class="btn btn-gold" onclick="runMonteCarloSimulation()">Lancer la simulation</button>
          </div>
        </div>
        <div class="grid-3">
          <div class="card feature-highlight-card">
            <div class="section-title">Scenario de base</div>
            <div style="font-size:0.85rem;color:var(--muted);line-height:1.7;">
              CA de base: <strong>${fmt(base.revenue)}</strong> XOF<br>
              Charges fixes: <strong>${fmt(base.fixedCosts)}</strong> XOF<br>
              Capex de reference: <strong>${fmt(base.capex)}</strong> XOF
            </div>
          </div>
          <div class="card feature-highlight-card">
            <div class="section-title">Structure economique</div>
            <div style="font-size:0.85rem;color:var(--muted);line-height:1.7;">
              Marge brute repere: <strong>${fmtPercent(base.grossMargin, 1)}</strong><br>
              BFR repere: <strong>${fmtPercent(base.workingCapitalPct, 1)}</strong><br>
              IS retenu: <strong>${fmtPercent(base.taxRate, 1)}</strong>
            </div>
          </div>
          <div class="card feature-highlight-card">
            <div class="section-title">Usage</div>
            <div style="font-size:0.85rem;color:var(--muted);line-height:1.7;">
              1. Ajuster les hypotheses ci-dessous.<br>
              2. Lancer plusieurs milliers d'iterations.<br>
              3. Lire les percentiles P10 / P50 / P90 avant de decider.
            </div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Hypotheses de simulation</div>
            <div class="card-subtitle">Toutes les valeurs sont memorisees par entreprise.</div>
          </div>
        </div>
        <div class="grid-2">
          <div class="stack">
            <div class="section-title">Base du scenario</div>
            <div class="grid-2">
              ${renderInput("baseRevenue", "CA de base (XOF)", "Reference annuelle sur laquelle la simulation se construit.")}
              ${renderInput("baseFixedCosts", "Charges fixes (XOF)", "Hors cout variable et achats directement lies aux ventes.")}
              ${renderInput("baseDepreciation", "Amortissements (XOF)", "Impact annuel non cash sur le resultat.")}
              ${renderInput("baseCapex", "Capex de base (XOF)", "Investissements annuels de maintien ou de croissance.")}
              ${renderInput("taxRate", "Taux IS (%)", "Le module applique l'impot uniquement en cas de resultat positif.")}
              ${renderInput("iterations", "Nombre d'iterations", "250 a 10 000 iterations. 2 500 est un bon point de depart.")}
            </div>

            <div class="section-title">Croissance du chiffre d'affaires (%)</div>
            <div class="grid-3">
              ${renderInput("revenueGrowthMin", "Min")}
              ${renderInput("revenueGrowthMode", "Central")}
              ${renderInput("revenueGrowthMax", "Max")}
            </div>

            <div class="section-title">Marge brute (%)</div>
            <div class="grid-3">
              ${renderInput("grossMarginMin", "Min")}
              ${renderInput("grossMarginMode", "Central")}
              ${renderInput("grossMarginMax", "Max")}
            </div>
          </div>

          <div class="stack">
            <div class="section-title">Couts fixes (multiplicateur)</div>
            <div class="grid-3">
              ${renderInput("fixedCostMin", "Min", "Exemple 0.92 = -8%")}
              ${renderInput("fixedCostMode", "Central")}
              ${renderInput("fixedCostMax", "Max", "Exemple 1.16 = +16%")}
            </div>

            <div class="section-title">BFR (% du CA)</div>
            <div class="grid-3">
              ${renderInput("workingCapitalMin", "Min")}
              ${renderInput("workingCapitalMode", "Central")}
              ${renderInput("workingCapitalMax", "Max")}
            </div>

            <div class="section-title">Capex (multiplicateur)</div>
            <div class="grid-3">
              ${renderInput("capexMin", "Min", "Exemple 0.85 = capex allege")}
              ${renderInput("capexMode", "Central")}
              ${renderInput("capexMax", "Max", "Exemple 1.30 = programme d'investissement charge")}
            </div>

            <div class="card feature-highlight-card">
              <div class="section-title">Sources comptables reprises</div>
              <div style="overflow-x:auto;">
                <table class="data-table compact-table">
                  <thead>
                    <tr>
                      <th>Rubrique</th>
                      <th>Valeur</th>
                      <th>Source</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${sourceRows.map((row) => `
                      <tr>
                        <td>${row[0]}</td>
                        <td>${row[1]}</td>
                        <td>${row[2]}</td>
                      </tr>
                    `).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div class="kpi-grid" style="margin-bottom:0;">
        ${topCards.map((card) => `
          <div class="kpi">
            <div class="kpi-label">${card.label}</div>
            <div class="kpi-value" style="font-size:1rem;color:${card.color};">${card.value}</div>
            <div class="kpi-note">${card.note}</div>
          </div>
        `).join("")}
      </div>

      <div class="grid-2">
        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Distribution du free cash-flow</div>
              <div class="card-subtitle">${lastRun ? "Histogramme des iterations Monte Carlo." : "Lancez une simulation pour obtenir la distribution."}</div>
            </div>
          </div>
          ${lastRun ? `
            <div class="mc-histogram">
              ${lastRun.histogram.map((bin) => `
                <div class="mc-bar-row">
                  <div class="mc-bar-label">${bin.label}</div>
                  <div class="mc-bar-track"><span style="width:${bin.widthPct}%;"></span></div>
                  <div class="mc-bar-value">${fmt(bin.count)}</div>
                </div>
              `).join("")}
            </div>
            <div class="info-box" style="margin-top:16px;">
              <strong>Lecture rapide:</strong> plus la masse des iterations se deplace vers la droite, plus la probabilite d'un cash positif augmente. Les queues basses donnent le niveau de stress financier a surveiller.
            </div>
          ` : `
            <div class="info-box">
              <strong>Aucune simulation lancee pour le moment.</strong><br><br>
              Chargez les hypotheses, cliquez sur <strong>Lancer la simulation</strong> puis utilisez les percentiles pour tester vos decisions d'investissement, de prix ou de reduction de couts.
            </div>
          `}
        </div>

        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Percentiles de sortie</div>
              <div class="card-subtitle">P10 = prudent, P50 = median, P90 = upside.</div>
            </div>
          </div>
          ${lastRun ? `
            <div style="overflow-x:auto;">
              <table class="data-table compact-table">
                <thead>
                  <tr>
                    <th>Indicateur</th>
                    <th>P10</th>
                    <th>P50</th>
                    <th>P90</th>
                    <th>Moyenne</th>
                  </tr>
                </thead>
                <tbody>
                  ${percentileRows.map(([label, metric]) => `
                    <tr>
                      <td>${label}</td>
                      <td>${fmt(Math.round(metric.p10))}</td>
                      <td>${fmt(Math.round(metric.p50))}</td>
                      <td>${fmt(Math.round(metric.p90))}</td>
                      <td>${fmt(Math.round(metric.average))}</td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
            <div class="stack" style="margin-top:16px;">
              <div class="card feature-highlight-card">
                <div class="section-title">Risques observes</div>
                <div style="font-size:0.85rem;color:var(--muted);line-height:1.8;">
                  Probabilite de perte: <strong>${fmtPercent(lastRun.risk.lossProbability, 1)}</strong><br>
                  Probabilite de cash negatif: <strong>${fmtPercent(lastRun.risk.negativeCashProbability, 1)}</strong><br>
                  Probabilite de marge EBITDA sous 12%: <strong>${fmtPercent(lastRun.risk.stressedMarginProbability, 1)}</strong>
                </div>
              </div>
              <div class="card feature-highlight-card">
                <div class="section-title">Bornes extraites</div>
                <div style="font-size:0.85rem;color:var(--muted);line-height:1.8;">
                  Pire resultat net: <strong>${fmt(Math.round(lastRun.tails.worstNetIncome))}</strong> XOF<br>
                  Meilleur resultat net: <strong>${fmt(Math.round(lastRun.tails.bestNetIncome))}</strong> XOF<br>
                  Pire free cash-flow: <strong>${fmt(Math.round(lastRun.tails.worstFreeCashFlow))}</strong> XOF
                </div>
              </div>
            </div>
          ` : `
            <div class="info-box">
              <strong>Sorties attendues:</strong><br><br>
              Le module produit des percentiles sur le chiffre d'affaires, l'EBITDA, le resultat net et le free cash-flow. Cela permet de discuter en comite de direction d'un cas prudent, median et offensif avec un seul dossier comptable.
            </div>
          `}
        </div>
      </div>
    </div>
  `;
}

function renderReductionCouts() {
  const snapshot = getCostReductionSnapshot();
  const plan = snapshot.plan;
  const opportunities = snapshot.opportunities.filter((row) => row.amount > 0).slice(0, 5);

  return `
    <div class="stack">
      <div class="card feature-gradient-card">
        <div class="card-header">
          <div>
            <div class="card-title">Reduction des couts</div>
            <div class="card-subtitle">Module de pilotage inspire des techniques de reduction des couts, relie aux familles de charges OHADA et transformable en plan d'actions entreprise.</div>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button class="btn btn-outline" onclick="seedCostReductionPlanFromAccounting()">Generer depuis les charges</button>
            <button class="btn btn-outline" onclick="addCostReductionAction()">Ajouter une action</button>
            <button class="btn btn-outline" onclick="resetCostReductionPlan()">Reinitialiser</button>
            <button class="btn btn-gold" onclick="saveCostReductionPlanFromForm()">Enregistrer et recalculer</button>
          </div>
        </div>
        <div class="grid-3">
          ${COST_REDUCTION_PILLARS.map((pillar) => `
            <div class="card feature-highlight-card">
              <div class="section-title">${pillar.title}</div>
              <div style="font-size:0.84rem;color:var(--muted);line-height:1.7;">${pillar.description}</div>
            </div>
          `).join("")}
        </div>
      </div>

      <div class="kpi-grid" style="margin-bottom:0;">
        <div class="kpi">
          <div class="kpi-label">Potentiel cumule</div>
          <div class="kpi-value" style="font-size:1rem;color:var(--green);">${fmt(snapshot.totalPotential)}</div>
          <div class="kpi-note">XOF / an identifies dans le plan d'actions</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Couverture des charges</div>
          <div class="kpi-value" style="font-size:1rem;color:${snapshot.coverageRate <= 0.2 ? 'var(--green)' : 'var(--orange)'};">${fmtPercent(snapshot.coverageRate, 1)}</div>
          <div class="kpi-note">${fmt(snapshot.totalCharges)} XOF de charges suivies</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Actions avancees</div>
          <div class="kpi-value" style="font-size:1rem;">${snapshot.inProgressCount + snapshot.implementedCount}</div>
          <div class="kpi-note">${snapshot.implementedCount} appliquees | ${snapshot.priorityCount} priorites 30 jours</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Famille dominante</div>
          <div class="kpi-value" style="font-size:1rem;">${snapshot.topOpportunity ? snapshot.topOpportunity.code : "—"}</div>
          <div class="kpi-note">${snapshot.topOpportunity ? `${snapshot.topOpportunity.label} · ${fmt(snapshot.topOpportunity.amount)} XOF` : "Aucune charge detectee"}</div>
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Lecture OHADA des charges</div>
              <div class="card-subtitle">Les plus gros postes actuels donnent vos premiers axes d'action.</div>
            </div>
          </div>
          ${opportunities.length ? `
            <div class="stack">
              ${opportunities.map((row) => `
                <div class="cost-family-row">
                  <div>
                    <div class="cost-family-label">${row.code} — ${row.label}</div>
                    <div class="cost-family-note">${row.note}</div>
                  </div>
                  <div class="cost-family-amount">${fmt(row.amount)} XOF</div>
                </div>
              `).join("")}
            </div>
          ` : `
            <div class="info-box">
              Aucune charge significative n'a encore ete lue dans la balance ou le journal. Importez vos soldes puis relancez la generation du plan pour obtenir des gains cibles.
            </div>
          `}
        </div>

        <div class="card">
          <div class="card-header">
            <div>
              <div class="card-title">Erreurs a eviter</div>
              <div class="card-subtitle">A garder visibles pendant le programme d'economies.</div>
            </div>
          </div>
          <div class="stack">
            ${COST_REDUCTION_ERRORS.map((item) => `
              <div class="cost-warning-row">${item}</div>
            `).join("")}
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Plan d'actions reduction des couts</div>
            <div class="card-subtitle">Editable par entreprise. Les gains sont exprimes en XOF par an.</div>
          </div>
        </div>
        <div style="overflow-x:auto;">
          <table class="data-table cost-plan-table">
            <thead>
              <tr>
                <th>Technique</th>
                <th>Axe</th>
                <th>Classes OHADA</th>
                <th>Gain potentiel</th>
                <th>Statut</th>
                <th>Responsable</th>
                <th>Action prioritaire</th>
              </tr>
            </thead>
            <tbody>
              ${plan.map((row) => `
                <tr
                  data-cost-row="${escapeHtml(row.id)}"
                  data-target-rate="${row.targetRate}"
                  data-recommended-prefixes='${escapeHtml(JSON.stringify(row.recommendedPrefixes))}'>
                  <td><input class="form-input" data-cost-field="technique" value="${escapeHtml(row.technique)}" placeholder="Technique"></td>
                  <td>
                    <select class="form-select" data-cost-field="pillar">
                      ${COST_REDUCTION_PILLARS.map((pillar) => `
                        <option value="${pillar.key}" ${row.pillar === pillar.key ? "selected" : ""}>${pillar.title}</option>
                      `).join("")}
                    </select>
                  </td>
                  <td><input class="form-input" data-cost-field="classes" value="${escapeHtml(row.classes)}" placeholder="60 / 62 / 65"></td>
                  <td><input class="form-input" type="number" data-cost-field="estimatedGain" value="${escapeHtml(String(row.estimatedGain || 0))}" placeholder="0"></td>
                  <td>
                    <select class="form-select" data-cost-field="status">
                      ${COST_REDUCTION_STATUS_OPTIONS.map((status) => `
                        <option value="${status}" ${row.status === status ? "selected" : ""}>${status}</option>
                      `).join("")}
                    </select>
                  </td>
                  <td><input class="form-input" data-cost-field="owner" value="${escapeHtml(row.owner)}" placeholder="Responsable"></td>
                  <td><input class="form-input" data-cost-field="action" value="${escapeHtml(row.action)}" placeholder="Action prioritaire"></td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function renderDSF() {
  const packetName = getDsfPacketName();
  const templateMeta = getExactFiscalTemplateMeta();
  const statuses = getDsfStatusRows();
  const readyCount = statuses.filter((item) => item.status === "Pret").length;

  function getStatusColor(status) {
    if (status === "Pret") return "var(--green)";
    if (status === "A completer") return "var(--orange)";
    return "var(--muted)";
  }

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Declaration Statistique et Fiscale (DSF)</div>
          <div class="card-subtitle">Remplissage du modele exact ${templateMeta.downloadName} a partir des donnees de l'application.</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn btn-outline" onclick="shareBfaLiasseFiscale()">Partager ${templateMeta.buttonLabel}</button>
          <button class="btn btn-gold" onclick="downloadBfaLiasseFiscale()">Telecharger ${templateMeta.buttonLabel}</button>
        </div>
      </div>
      <div class="info-box" style="margin-bottom:16px;">
        <strong>Packet cible:</strong> ${packetName}<br><br>
        Le telechargement repose desormais sur le modele exact <strong>${templateMeta.downloadName}</strong> integre au projet. ${
          currentRef === "sycebnl"
            ? `Le choix du modele se fait automatiquement selon le type d'entite SYCEBNL (${sycebnlType === "projets" ? "projets de developpement" : "associations / ONG / fondations"}).`
            : "Les champs d'identification et de codification pays/forme/regime sont pre-remplis a partir de vos donnees."
        } Etat d'avancement actuel: <strong>${readyCount}/${statuses.length}</strong> rubriques pretes.
      </div>
      <div class="stack">
        ${statuses.map(d => `
          <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);">
            <div>
              <span style="font-family:var(--mono);font-weight:700;color:var(--gold);margin-right:10px;">${d.code}</span>
              <span>${d.label}</span>
              <div style="font-size:0.78rem;color:var(--muted);margin-top:4px;">${d.hint}</div>
            </div>
            <span style="font-size:0.78rem;font-weight:700;color:${getStatusColor(d.status)};">${d.status}</span>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// GUIDE OHADA
// ═══════════════════════════════════════════════════════════
function renderVeilleOhada() {
  startOhadaWatchTicker();
  const state = ensureOhadaWatchDailyRefresh(false);
  const lastRefreshDate = state.lastRefreshAt ? new Date(state.lastRefreshAt) : null;
  const hasValidRefreshDate = !!(lastRefreshDate && !Number.isNaN(lastRefreshDate.getTime()));
  const lastRefreshText = hasValidRefreshDate ? formatDateTimeValue(state.lastRefreshAt) : "Jamais";
  const nextRefreshText = hasValidRefreshDate
    ? formatDateTimeValue(new Date(lastRefreshDate.getTime() + 24 * 60 * 60 * 1000).toISOString())
    : "Au prochain affichage";

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Veille OHADA quotidienne</div>
          <div class="card-subtitle">Recharge automatique une fois par jour du site officiel OHADA.com pendant la consultation.</div>
        </div>
      </div>
      <div class="watch-meta-grid">
        <div class="watch-meta-card">
          <div class="watch-meta-label">Source</div>
          <div class="watch-meta-value">OHADA.com</div>
          <div class="watch-meta-note">Affichage en direct du site officiel ${OHADA_SITE_HOME_URL}</div>
        </div>
        <div class="watch-meta-card">
          <div class="watch-meta-label">Derniere recharge</div>
          <div class="watch-meta-value">${lastRefreshText}</div>
          <div class="watch-meta-note">L'iframe est regeneree avec un cache-buster journalier.</div>
        </div>
        <div class="watch-meta-card">
          <div class="watch-meta-label">Prochaine recharge</div>
          <div class="watch-meta-value">${nextRefreshText}</div>
          <div class="watch-meta-note">Si cette page reste ouverte, OHADA Compta verifiera automatiquement le changement de jour.</div>
        </div>
      </div>
      <div class="info-box" style="margin-bottom:16px;">
        <strong>Mode de mise a jour active</strong><br><br>
        Cette veille recharge le site officiel OHADA.com une fois par jour pendant l'affichage. Sur GitHub Pages, le navigateur ne permet pas de lire automatiquement le contenu editorial de ce site en arriere-plan a cause des restrictions CORS, mais l'affichage direct du site officiel reste possible et est maintenant integre dans OHADA Compta.
      </div>
      <div class="watch-toolbar">
        <button class="btn btn-gold" onclick="refreshOhadaWatch()">Actualiser maintenant</button>
        <button class="btn btn-outline" onclick="openOhadaWebsite()">Ouvrir OHADA.com</button>
        <button class="btn btn-outline" onclick="openOhadaWebsite('https://www.ohada.com/actualite.html')">Ouvrir les actualites</button>
      </div>
      <div class="watch-frame-shell">
        <iframe
          class="watch-frame"
          src="${state.frameUrl}"
          title="Veille quotidienne OHADA.com"
          loading="lazy"
          referrerpolicy="strict-origin-when-cross-origin"></iframe>
      </div>
    </div>
  `;
}

// GUIDE OHADA
// ═══════════════════════════════════════════════════════════
function renderGuide() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">${OHADA_GUIDE.title}</div></div>
      <div class="stack">
        ${OHADA_GUIDE.sections.map(s => `
          <div class="card" style="background:var(--surface2);">
            <div style="font-weight:700;color:var(--gold);margin-bottom:8px;">${s.title}</div>
            <div style="font-size:0.88rem;color:var(--muted);line-height:1.7;">${s.content}</div>
          </div>
        `).join("")}
      </div>
      <div style="margin-top:20px;">
        <div class="section-title">Pays membres OHADA (17)</div>
        <div class="filter-row">
          ${OHADA_MEMBER_STATES.map(p => `<span class="filter-pill">${p}</span>`).join("")}
        </div>
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// COMPARAISON SYSCOHADA vs SYCEBNL
// ═══════════════════════════════════════════════════════════
function renderComparaison() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">SYSCOHADA Revise vs SYCEBNL</div></div>
      <table class="data-table">
        <thead><tr><th>Critere</th><th style="color:var(--gold);">SYSCOHADA Revise</th><th style="color:var(--cyan);">SYCEBNL</th></tr></thead>
        <tbody>
          ${[
            ["Entites visees", "Entreprises a but lucratif", "Associations, ONG, fondations, syndicats"],
            ["Base juridique", "Acte Uniforme AUDCIF (2017)", "Acte Uniforme AUDCIF — volet EBNL"],
            ["Classe 1", "Capital, reserves, resultat", "Fonds propres, fonds dedies par projet"],
            ["Classe 4 — spec.", "Clients, fournisseurs, Etat", "Donateurs, bailleurs de fonds (45x)"],
            ["Classe 7 — spec.", "Ventes, prestations", "Dons, contributions, cotisations (74x)"],
            ["Classe 6 — spec.", "Charges standard", "Charges liees aux projets (65x)"],
            ["Resultat", "Benefice / Perte", "Excedent / Deficit"],
            ["Tableau de flux / Emplois-Ressources", "Tableau de flux de tresorerie (TFT)", "TFT pour associations; tableaux emplois/ressources pour projets"],
            ["DSF / DGI", "Liasse fiscale standard", "Liasse adaptee EBNL"],
            ["Etats financiers", "Bilan, Resultat, TFT, Annexes", "Jeu adapte selon la categorie EBNL"],
            ["Nombre de comptes", `${PLAN_COMPTABLE_SYSCOHADA.length}`, `${PLAN_COMPTABLE_SYSCOHADA.length + PLAN_COMPTABLE_SYCEBNL_ADDITIONS.length} (avec ajouts)`],
            ["Pays applicables", "17 Etats membres OHADA", "17 Etats membres OHADA"],
          ].map(([crit, sys, bnl]) => `
            <tr>
              <td style="font-weight:600;">${crit}</td>
              <td style="font-size:0.84rem;">${sys}</td>
              <td style="font-size:0.84rem;">${bnl}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
      <div class="info-box" style="margin-top:16px;">
        <strong>Note importante:</strong> Le SYCEBNL partage la quasi-totalite du plan comptable SYSCOHADA. Les differences portent principalement sur les comptes de fonds dedies (classe 1), les donateurs/bailleurs (classe 4), les dons et contributions (classe 7), et les charges de projets (classe 6). Le basculement entre les deux referentiels se fait via le selecteur en haut de l'application.
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// EVENT HANDLERS
// ═══════════════════════════════════════════════════════════
function attachEvents() {
  const paramAccountingSystem = document.getElementById("p-accountingSystem");
  if (paramAccountingSystem) {
    paramAccountingSystem.addEventListener("change", syncParamAccountingSystemUI);
  }
  syncParamAccountingSystemUI();

  // Plan search
  const searchEl = document.getElementById("plan-search");
  if (searchEl) {
    searchEl.addEventListener("input", (e) => { searchTerm = e.target.value; render(); });
  }

  // Class filter
  document.querySelectorAll(".filter-pill[data-class]").forEach(btn => {
    btn.addEventListener("click", () => {
      const val = parseInt(btn.dataset.class);
      filterClass = val === 0 ? null : val;
      render();
    });
  });

  // Saisie — valider
  const btnValider = document.getElementById("btn-valider-ecriture");
  if (btnValider) {
    btnValider.addEventListener("click", () => {
      const date = document.getElementById("saisie-date").value;
      const journal = document.getElementById("saisie-journal").value;
      const libelle = document.getElementById("saisie-libelle").value;
      const ref = document.getElementById("saisie-ref").value;
      const comptes = document.querySelectorAll(".saisie-compte");
      const debits = document.querySelectorAll(".saisie-debit");
      const credits = document.querySelectorAll(".saisie-credit");
      let totalD = 0, totalC = 0;
      const lignes = [];
      comptes.forEach((c, i) => {
        const d = parseFloat(debits[i].value) || 0;
        const cr = parseFloat(credits[i].value) || 0;
        if (d > 0 || cr > 0) {
          lignes.push({ compte: c.value, debit: d, credit: cr });
          totalD += d; totalC += cr;
        }
      });
      const msg = document.getElementById("saisie-msg");
      if (lignes.length < 2) { msg.innerHTML = '<span style="color:var(--red);">Minimum 2 lignes requises.</span>'; return; }
      if (!isBalancedAmount(totalD, totalC)) { msg.innerHTML = `<span style="color:var(--red);">Desequilibre: Debit ${fmt(totalD)} != Credit ${fmt(totalC)}</span>`; return; }
      if (!libelle) { msg.innerHTML = '<span style="color:var(--red);">Libelle requis.</span>'; return; }

      const piece = journal + "-" + String(journalEntries.length + 1).padStart(3, "0");
      lignes.forEach(l => {
        journalEntries.push({
          id: journalEntries.length + 1,
          date, journal, piece, compte: l.compte, libelle, debit: l.debit, credit: l.credit, ref
        });
      });
      msg.innerHTML = `<span style="color:var(--green);">Ecriture ${piece} enregistree (${lignes.length} lignes, ${fmt(totalD)} XOF).</span>`;
    });
  }

  // Saisie — ajouter ligne
  const btnAjout = document.getElementById("btn-ajouter-ligne");
  if (btnAjout) {
    btnAjout.addEventListener("click", () => {
      const plan = getPlan();
      const container = document.getElementById("saisie-lignes");
      const div = document.createElement("div");
      div.className = "grid-3";
      div.style.marginBottom = "8px";
      div.innerHTML = `
        <div class="form-group"><select class="form-select saisie-compte">${plan.map(a => `<option value="${a.numero}">${a.numero} — ${a.libelle}</option>`).join("")}</select></div>
        <div class="form-group"><input class="form-input saisie-debit" type="number" placeholder="0"></div>
        <div class="form-group"><input class="form-input saisie-credit" type="number" placeholder="0"></div>
      `;
      container.appendChild(div);
    });
  }
}


// ═══════════════════════════════════════════════════════════
// DRAG & DROP FILE IMPORT
// ═══════════════════════════════════════════════════════════
(function () {
  const overlay = document.getElementById('drop-overlay');
  let dragCounter = 0;

  document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    overlay.classList.add('active');
  });
  document.addEventListener('dragleave', () => {
    dragCounter--;
    if (dragCounter <= 0) { dragCounter = 0; overlay.classList.remove('active'); }
  });
  document.addEventListener('dragover', (e) => { e.preventDefault(); });
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    overlay.classList.remove('active');
    const file = e.dataTransfer.files[0];
    if (file) handleDroppedFile(file);
  });
})();

function handleDroppedFile(file) {
  const name = file.name.toLowerCase();

  if (name.endsWith('.csv') || name.endsWith('.txt')) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const result = parseBalanceCsv(e.target.result);
      if (result.count > 0 && result.isBalanced) {
        replaceOpeningBalances(result.balances);
        showToast(result.count + ' comptes importes depuis ' + file.name + ' - balance d\'ouverture equilibree.', 'success');
        render();
      } else if (result.count > 0) {
        showToast('Import refuse: la balance d\'ouverture n\'est pas equilibree (ecart ' + fmt(Math.abs(result.gap)) + ').', 'error');
      } else {
        showToast('Aucun compte detecte. Verifiez le format CSV.', 'error');
      }
    };
    reader.readAsText(file);

  } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    if (typeof XLSX === 'undefined') {
      showToast('Librairie Excel non chargee. Verifiez votre connexion.', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        // Try to find a balance sheet (first sheet with account data)
        const importedBalances = {};
        let count = 0;
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        // Detect header row: look for a row with 'Compte' or numeric-looking first cell
        let startRow = 1;
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
          const first = String(rows[i][0]||'').trim();
          if (first.toLowerCase().includes('compte') || /^\d{1,4}$/.test(first)) {
            startRow = /^\d{1,4}$/.test(first) ? i : i + 1;
            break;
          }
        }
        rows.slice(startRow).forEach(row => {
          const code = String(row[0]||'').replace(/\..*/, '').trim(); // strip .0
          const n1d = parseFloat(row[2]) || 0;
          const n1c = parseFloat(row[3]) || 0;
          if (code && /^\d+$/.test(code)) {
            importedBalances[code] = { n1d, n1c };
            count++;
          }
        });
        const summary = getOpeningBalanceSummary(importedBalances);
        if (count > 0 && summary.isBalanced) {
          replaceOpeningBalances(importedBalances);
          showToast(count + ' comptes importes depuis ' + file.name + ' - balance d\'ouverture equilibree.', 'success');
          render();
        } else if (count > 0) {
          showToast('Import refuse: la balance d\'ouverture n\'est pas equilibree (ecart ' + fmt(Math.abs(summary.gap)) + ').', 'error');
        } else {
          showToast('Aucun compte trouve dans ' + file.name + '. Format attendu: Compte | Libelle | S.O.D | S.O.C', 'error');
        }
      } catch (err) {
        showToast('Erreur lecture Excel: ' + err.message, 'error');
      }
    };
    reader.readAsBinaryString(file);

  } else if (name.endsWith('.json')) {
    // JSON journal entries import
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        const entries = Array.isArray(data) ? data : (data.journal || data.ecritures || []);
        const importedEntries = [];
        const totalsByPiece = {};
        const fallbackPiece = 'IMP-' + String(journalEntries.length + 1).padStart(3, '0');
        let count = 0;
        entries.forEach(entry => {
          const d = parseFloat(entry.debit) || 0;
          const c = parseFloat(entry.credit) || 0;
          if (!entry.compte || (!d && !c)) return;
          const piece = entry.piece || fallbackPiece;
          if (!totalsByPiece[piece]) totalsByPiece[piece] = { debit: 0, credit: 0 };
          totalsByPiece[piece].debit += d;
          totalsByPiece[piece].credit += c;
          importedEntries.push({
            date: entry.date || new Date().toISOString().split('T')[0],
            journal: entry.journal || 'OD',
            piece,
            compte: String(entry.compte),
            libelle: entry.libelle || 'Import',
            debit: d, credit: c,
            ref: entry.ref || 'IMPORT'
          });
          count++;
        });
        const invalidPieces = Object.keys(totalsByPiece).filter((piece) => !isBalancedAmount(totalsByPiece[piece].debit, totalsByPiece[piece].credit));
        if (invalidPieces.length > 0) {
          showToast('Import refuse: piece(s) desequilibree(s) ' + invalidPieces.slice(0, 3).join(', ') + '.', 'error');
          return;
        }
        if (count > 0) {
          const startingId = journalEntries.length;
          importedEntries.forEach((entry, index) => {
            journalEntries.push({
              id: startingId + index + 1,
              ...entry
            });
          });
          showToast(count + ' ecritures importees depuis ' + file.name, 'success');
          render();
        } else {
          showToast('Aucune ecriture valide dans le fichier JSON.', 'error');
        }
      } catch (err) {
        showToast('Erreur lecture JSON: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);

  } else {
    showToast('Format non supporte: ' + file.name + ' — Utiliser CSV, Excel (.xlsx) ou JSON', 'error');
  }
}

function parseBalanceCsv(text) {
  const lines = text.split('\n').slice(1);
  const balances = {};
  let count = 0;
  lines.forEach(line => {
    const cols = line.split(',');
    if (cols.length < 4) return;
    const code = cols[0].trim().replace(/"/g, '');
    const n1d = parseFloat(cols[2]) || 0;
    const n1c = parseFloat(cols[3]) || 0;
    if (code && /^\d+$/.test(code)) {
      balances[code] = { n1d, n1c };
      count++;
    }
  });
  return {
    count,
    balances,
    ...getOpeningBalanceSummary(balances)
  };
}

function showToast(msg, type) {
  const container = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast toast-' + (type || 'info');
  t.textContent = msg;
  container.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transition = 'opacity .4s'; setTimeout(() => t.remove(), 400); }, 3500);
}


// ═══════════════════════════════════════════════════════════
// PARAMETRES ENTREPRISE
// ═══════════════════════════════════════════════════════════
function renderParametres() {
  const d = currentCompanyDetails;
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
  const selectedAccountingSystem = normalizeReferential(d.accountingSystem || currentRef);
  const selectedSycebnlType = normalizeSycebnlEntityType(d.sycebnlEntityType || sycebnlType);

  const formesJuridiques = ["SA","SARL","SAS","SASU","EURL","SNC","SCS","GIE","Cooperative","Association","ONG","Fondation","Projet de developpement","Etablissement public","Autre"];
  const regimesFiscaux = ["Reel Normal d'Imposition (RNI)","Reel Simplifie d'Imposition (RSI)","Contribution des Micro-Entreprises (CME)","Forfait d'Imposition","Exonere"];

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Identification de l'entreprise</div>
        <div class="card-subtitle">Le referentiel comptable choisi ici devient celui du dossier et s'applique immediatement.</div>
      </div>

      <div class="grid-2">
        <div class="form-group">
          <div class="form-label">Systeme comptable</div>
          <select class="form-select" id="p-accountingSystem">
            <option value="syscohada" ${selectedAccountingSystem === "syscohada" ? "selected" : ""}>SYSCOHADA Revise</option>
            <option value="sycebnl" ${selectedAccountingSystem === "sycebnl" ? "selected" : ""}>SYCEBNL</option>
          </select>
        </div>
        <div class="form-group" id="p-sycebnl-wrap" style="${selectedAccountingSystem === "sycebnl" ? "" : "display:none;"}">
          <div class="form-label">Type d'entite SYCEBNL</div>
          <select class="form-select" id="p-sycebnlType">
            <option value="associations" ${selectedSycebnlType === "associations" ? "selected" : ""}>Association / ONG / Fondation</option>
            <option value="projets" ${selectedSycebnlType === "projets" ? "selected" : ""}>Projet de developpement</option>
          </select>
        </div>
        <div class="form-group">
          <div class="form-label">Raison sociale</div>
          <input class="form-input" id="p-raisonSociale" value="${d.raisonSociale || (acct ? acct.company : '')}" placeholder="Nom officiel de l'entreprise">
        </div>
        <div class="form-group">
          <div class="form-label">Forme juridique</div>
          <select class="form-select" id="p-formeJuridique">
            <option value="">-- Choisir --</option>
            ${formesJuridiques.map(f => `<option value="${f}" ${(d.formeJuridique||'')===f?'selected':''}>${f}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <div class="form-label">Sigle usuel</div>
          <input class="form-input" id="p-sigleUsuel" value="${d.sigleUsuel || ''}" placeholder="Sigle ou abreviation">
        </div>
        <div class="form-group">
          <div class="form-label">RCCM</div>
          <input class="form-input" id="p-rccm" value="${d.rccm || ''}" placeholder="Numero RCCM">
        </div>
        <div class="form-group">
          <div class="form-label">NIF (Numero d'Identification Fiscale)</div>
          <input class="form-input" id="p-nif" value="${d.nif || ''}" placeholder="Numero d'identification fiscale">
        </div>
        <div class="form-group">
          <div class="form-label">N° de teledeclarant (NES)</div>
          <input class="form-input" id="p-nes" value="${d.nes || ''}" placeholder="Numero NES">
        </div>
        <div class="form-group" style="grid-column:1/-1;">
          <div class="form-label">Siege social</div>
          <input class="form-input" id="p-siegeSocial" value="${d.siegeSocial || ''}" placeholder="Adresse du siege social">
        </div>
        <div class="form-group" style="grid-column:1/-1;">
          <div class="form-label">Activite principale</div>
          <input class="form-input" id="p-activitePrincipale" value="${d.activitePrincipale || ''}" placeholder="Description de l'activite principale">
        </div>
        <div class="form-group">
          <div class="form-label">Capital social (XOF)</div>
          <input class="form-input" type="number" id="p-capitalSocial" value="${d.capitalSocial || ''}" placeholder="Montant du capital social">
        </div>
        <div class="form-group">
          <div class="form-label">Regime fiscal</div>
          <select class="form-select" id="p-regimeFiscal">
            <option value="">-- Choisir --</option>
            ${regimesFiscaux.map(r => `<option value="${r}" ${(d.regimeFiscal||'')===r?'selected':''}>${r}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <div class="form-label">Pays</div>
          <input class="form-input" id="p-pays" value="${d.pays || ''}" placeholder="Pays d'immatriculation">
        </div>
        <div class="form-group" style="display:flex;gap:12px;align-items:flex-end;">
          <div style="flex:1;">
            <div class="form-label">Exercice du</div>
            <input class="form-input" type="date" id="p-exerciceDu" value="${d.exerciceDu || (new Date().getFullYear()+'-01-01')}">
          </div>
          <div style="flex:1;">
            <div class="form-label">au</div>
            <input class="form-input" type="date" id="p-exerciceAu" value="${d.exerciceAu || (new Date().getFullYear()+'-12-31')}">
          </div>
        </div>
        <div class="form-group">
          <div class="form-label">Telephone</div>
          <input class="form-input" id="p-tel" value="${d.tel || ''}" placeholder="Numero de telephone">
        </div>
        <div class="form-group">
          <div class="form-label">Email</div>
          <input class="form-input" id="p-emailCompta" value="${d.emailCompta || (acct ? acct.email : '')}" placeholder="Adresse email comptable">
        </div>
        <div class="form-group">
          <div class="form-label">Expert-comptable</div>
          <input class="form-input" id="p-expertComptable" value="${d.expertComptable || ''}" placeholder="Nom et cabinet">
        </div>
        <div class="form-group">
          <div class="form-label">Commissaire aux comptes</div>
          <input class="form-input" id="p-commissaire" value="${d.commissaire || ''}" placeholder="Nom et cabinet">
        </div>
      </div>

      <div class="section-title" style="margin-top:20px;">Taux fiscaux applicables</div>
      <div class="grid-2">
        <div class="form-group">
          <div class="form-label">Taux IS — Impot sur les Societes (%)</div>
          <input class="form-input" type="number" step="0.01" id="p-tauxIS" value="${d.tauxIS || ''}" placeholder="Ex: 25">
        </div>
        <div class="form-group">
          <div class="form-label">Taux IMF — Impot Minimum Forfaitaire (%)</div>
          <input class="form-input" type="number" step="0.01" id="p-tauxIMF" value="${d.tauxIMF || ''}" placeholder="Ex: 0.5">
        </div>
        <div class="form-group">
          <div class="form-label">Taux TVA (%)</div>
          <input class="form-input" type="number" step="0.01" id="p-tauxTVA" value="${d.tauxTVA || ''}" placeholder="Ex: 18">
        </div>
      </div>

      <div id="p-msg" style="margin-top:16px;min-height:20px;font-size:0.86rem;text-align:center;"></div>
      <button class="btn btn-gold" style="margin-top:8px;width:100%;" onclick="saveParametres()">Enregistrer la fiche entreprise</button>
    </div>

    <div class="card" style="margin-top:16px;">
      <div class="card-header"><div class="card-title">Compte utilisateur</div></div>
      <div style="display:flex;flex-direction:column;gap:8px;font-size:0.9rem;">
        <div><span style="color:var(--muted);">Entreprise enregistree:</span> <strong>${acct ? acct.company : '—'}</strong></div>
        <div><span style="color:var(--muted);">Email:</span> <strong>${acct ? acct.email : '—'}</strong></div>
        <div><span style="color:var(--muted);">Compte cree le:</span> <strong>${acct ? new Date(acct.createdAt).toLocaleDateString('fr-FR') : '—'}</strong></div>
        <div><span style="color:var(--muted);">Ecritures sauvegardees:</span> <strong>${journalEntries.length}</strong></div>
      </div>
    </div>
  `;
}

function saveParametres() {
  const accountingSystemEl = document.getElementById('p-accountingSystem');
  const sycebnlTypeEl = document.getElementById('p-sycebnlType');
  const accountingSystem = normalizeReferential(accountingSystemEl ? accountingSystemEl.value : currentRef);
  const nextSycebnlType = normalizeSycebnlEntityType(sycebnlTypeEl ? sycebnlTypeEl.value : sycebnlType);
  const fields = ['raisonSociale','formeJuridique','sigleUsuel','rccm','nif','nes','siegeSocial','activitePrincipale',
                  'capitalSocial','regimeFiscal','pays','exerciceDu','exerciceAu',
                  'tel','emailCompta','expertComptable','commissaire',
                  'tauxIS','tauxIMF','tauxTVA'];
  fields.forEach(f => {
    const el = document.getElementById('p-' + f);
    if (el) currentCompanyDetails[f] = el.value.trim();
  });
  applyAccountingSystemSelection(accountingSystem, { sycebnlType: nextSycebnlType, save: false });
  saveCompanyData();
  // Refresh topbar badge with updated raison sociale
  const badge = document.getElementById('company-name-display');
  if (badge && currentCompanyDetails.raisonSociale) badge.textContent = currentCompanyDetails.raisonSociale;
  const msg = document.getElementById('p-msg');
  if (msg) {
    msg.style.color = 'var(--green)';
    msg.textContent = `Fiche enregistree avec succes. Referentiel actif: ${getReferentialLabel()}.`;
    setTimeout(() => { msg.textContent = ''; }, 2500);
  }
  render();
}

function cloturerExercice() {
  const snapshot = getClotureSnapshot();

  if (!snapshot.profileComplete) {
    showToast("Impossible de cloturer: completez d'abord la fiche entreprise.", "error");
    navigateToTab("parametres");
    return;
  }
  if (!snapshot.validDates) {
    showToast("Impossible de cloturer: verifiez les dates d'exercice.", "error");
    navigateToTab("parametres");
    return;
  }
  if (!snapshot.hasData) {
    showToast("Impossible de cloturer: aucune balance ni ecriture n'est disponible.", "error");
    return;
  }
  if (!snapshot.balanceSummary.isBalanced) {
    showToast(`Impossible de cloturer: la balance generale presente un ecart de ${fmt(Math.abs(snapshot.balanceSummary.gap))} XOF.`, "error");
    navigateToTab("balance");
    return;
  }
  if (!snapshot.carryforwardSummary.isBalanced) {
    showToast("Impossible de cloturer: le report des a-nouveaux n'est pas equilibre.", "error");
    return;
  }

  const exerciseLabel = `${formatDateValue(snapshot.exerciceDu)} au ${formatDateValue(snapshot.exerciceAu)}`;
  const confirmed = window.confirm(
    `Cloturer l'exercice ${exerciseLabel} ?\n\n` +
    `Cette action va:\n` +
    `- reporter les comptes de bilan en a-nouveaux,\n` +
    `- remettre le journal courant a zero,\n` +
    `- ouvrir l'exercice suivant.\n\n` +
    `Resultat reporte: ${fmt(Math.abs(snapshot.resultatExercice))} XOF ${snapshot.resultatExercice >= 0 ? "benefice" : "perte"}.`
  );
  if (!confirmed) return;

  const history = Array.isArray(currentCompanyDetails.closureHistory) ? currentCompanyDetails.closureHistory : [];
  currentCompanyDetails.closureHistory = [{
    closedAt: new Date().toISOString(),
    exerciseLabel,
    exerciceDu: snapshot.exerciceDu,
    exerciceAu: snapshot.exerciceAu,
    resultatExercice: snapshot.resultatExercice,
    totalDebit: snapshot.balanceSummary.totalDebit,
    totalCredit: snapshot.balanceSummary.totalCredit,
    openingAccounts: snapshot.carryforwardSummary.count
  }, ...history].slice(0, 12);

  replaceOpeningBalances(snapshot.carryforward);
  journalEntries = [];

  if (snapshot.nextExerciceDu) currentCompanyDetails.exerciceDu = snapshot.nextExerciceDu;
  if (snapshot.nextExerciceAu) currentCompanyDetails.exerciceAu = snapshot.nextExerciceAu;

  saveCompanyData();
  render();
  showToast(
    `Exercice cloture avec succes. Nouvel exercice: ${formatDateValue(currentCompanyDetails.exerciceDu)} au ${formatDateValue(currentCompanyDetails.exerciceAu)}.`,
    "success"
  );
}

// Initial auth check (shows login or loads company data)
syncTabFromLocationHash();
bindStaticAuthFormEvents();
checkAuth();

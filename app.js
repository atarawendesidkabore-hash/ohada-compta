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

function hasCompanyProfileData() {
  return COMPANY_PROFILE_FIELDS.some((key) => String(currentCompanyDetails[key] || "").trim() !== "");
}

function isCompanyProfileComplete() {
  return ["raisonSociale", "formeJuridique", "nif", "siegeSocial", "pays", "exerciceDu", "exerciceAu"]
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

function saveCompanyData() {
  if (!currentCompanyId) return;
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
    // Sync referentiel select
    const sel = document.getElementById('referentiel-select');
    if (sel) sel.value = currentRef;
  } catch(e) { console.warn('loadCompanyData error', e); }
}

function loginCompany(id) {
  currentCompanyId = id;
  localStorage.setItem(SESSION_KEY, id);
  // Reset to blank state — new companies start empty, returning ones get their data loaded
  journalEntries = [];
  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);
  currentRef = 'syscohada';
  sycebnlType = 'associations';
  currentCompanyDetails = {};
  const selReset = document.getElementById('referentiel-select');
  if (selReset) selReset.value = 'syscohada';
  loadCompanyData(id);
  // Update topbar
  const accounts = getAccounts();
  const acct = accounts.find(a => a.id === id);
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
  const sel = document.getElementById('referentiel-select');
  if (sel) sel.value = 'syscohada';
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
  accounts.push({ id, company, email, passHash: simpleHash(pass), createdAt: new Date().toISOString() });
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
  currentRef = e.target.value;
  render();
});

// SYCEBNL entity type switch (injected dynamically when SYCEBNL is active)
function setSycebnlType(type) {
  sycebnlType = type;
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

function navigateToTab(tab) {
  const target = document.querySelector(`.nav-btn[data-tab="${tab}"]`);
  if (!target) return;
  document.querySelectorAll(".nav-btn").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tab));
  currentTab = tab;
  closeMobileMenu();
  render();
  window.scrollTo(0, 0);
}

// Navigation
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => navigateToTab(btn.dataset.tab));
});

function getPlan() {
  if (currentRef === "sycebnl") {
    const base = PLAN_COMPTABLE_SYSCOHADA.filter(a => a.numero !== "13" && a.numero !== "131" && a.numero !== "139");
    return [...base, ...PLAN_COMPTABLE_SYCEBNL_ADDITIONS, ...PLAN_COMPTABLE_SYCEBNL_CLASSE9].sort((a, b) => a.numero.localeCompare(b.numero));
  }
  return PLAN_COMPTABLE_SYSCOHADA;
}

function fmt(n) { return n.toLocaleString("fr-FR"); }

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

function getDsfPacketName() {
  return currentRef === "sycebnl" ? "Liasse fiscale adaptee EBNL" : BF_LIASSE_PACKET_NAME;
}

function getDsfStatusRows() {
  const bal = computeBalances();
  const hasBalances = Object.keys(bal).length > 0;
  const hasJournal = journalEntries.length > 0;
  const hasImmos = Object.keys(bal).some((code) => code.startsWith("2") || code.startsWith("28"));
  const hasTiers = Object.keys(bal).some((code) => code.startsWith("4"));
  const profileComplete = isCompanyProfileComplete();

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

async function loadExactLiasseTemplateWorkbook() {
  const response = await fetch(`${EXACT_LIASSE_TEMPLATE_PATH}?v=20260326`);
  if (!response.ok) throw new Error(`Template ${EXACT_LIASSE_TEMPLATE_PATH} introuvable (${response.status})`);
  const buffer = await response.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellStyles: true, cellFormula: true });
}

async function loadExactForecastTemplateWorkbook() {
  const response = await fetch(`${EXACT_FORECAST_TEMPLATE_PATH}?v=20260327`);
  if (!response.ok) throw new Error(`Template ${EXACT_FORECAST_TEMPLATE_PATH} introuvable (${response.status})`);
  const buffer = await response.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellStyles: true, cellFormula: true });
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
  if (currentRef !== "syscohada") {
    throw new Error("Le fichier LIASSE.xlsx est reserve au referentiel SYSCOHADA. Basculez sur SYSCOHADA avant l'export.");
  }

  if (typeof XLSX === "undefined") {
    throw new Error("Librairie Excel non chargee. Verifiez votre connexion.");
  }

  const workbook = await loadExactLiasseTemplateWorkbook();
  populateExactLiasseTemplate(workbook);
  return workbook;
}

async function downloadExactLiasseFiscale() {
  const workbook = await buildExactLiasseWorkbook();
  XLSX.writeFile(workbook, EXACT_LIASSE_DOWNLOAD_NAME, { bookType: "xlsx", compression: true });
  showToast(`Le modele exact ${EXACT_LIASSE_DOWNLOAD_NAME} a ete rempli avec succes.`, "success");
}

async function shareExactLiasseFiscale() {
  const workbook = await buildExactLiasseWorkbook();
  const blob = workbookToBlob(workbook, EXACT_LIASSE_DOWNLOAD_NAME);
  await shareOrDownloadBlobFile(blob, EXACT_LIASSE_DOWNLOAD_NAME, {
    title: EXACT_LIASSE_DOWNLOAD_NAME,
    text: `LIASSE.xlsx preparee depuis OHADA COMPTA pour ${getCompanyDisplayName() || "l'entreprise"}.`
  });
}

async function downloadBfaLiasseFiscale() {
  try {
    await downloadExactLiasseFiscale();
  } catch (error) {
    console.warn("Exact template export failed", error);
    showToast("Modele LIASSE.xlsx indisponible. Generation du classeur interne en secours.", "info");
    downloadGeneratedBfaLiasseFiscale();
  }
}

async function shareBfaLiasseFiscale() {
  try {
    await shareExactLiasseFiscale();
  } catch (error) {
    console.warn("Exact template share failed", error);
    showToast("Modele LIASSE.xlsx indisponible. Partage du classeur interne en secours.", "info");
    await shareGeneratedBfaLiasseFiscale();
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
  const refLabel = currentRef === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise";

  // Company info
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
  const compName = currentCompanyDetails.raisonSociale || (acct ? acct.company : '');
  const capital = parseFloat(currentCompanyDetails.capitalSocial) || 0;
  const hasOpeningBalances = Object.keys(OPENING_BALANCES).length > 0;
  const hasProfile = hasCompanyProfileData();
  const profileComplete = isCompanyProfileComplete();
  const needsSetup = !profileComplete || !hasOpeningBalances || journalEntries.length === 0;

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
          <div style="font-size:0.84rem;color:var(--muted);margin-bottom:14px;">Renseignez la raison sociale, le NIF, l'exercice et les informations legales.</div>
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
        <strong>Pays membres:</strong> ${OHADA_MEMBER_STATES.join(", ")}
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
  const totals = [
    { label: "CA marchandises annualise", value: snapshot.merchandiseRevenue },
    { label: "CA services annualise", value: snapshot.serviceRevenue },
    { label: "Tresorerie de depart", value: snapshot.startingCash },
    { label: "Stock de depart", value: snapshot.openingStock },
    { label: "Investissements de depart", value: snapshot.intangibleSetup + snapshot.realEstateSetup + snapshot.worksSetup + snapshot.equipmentSetup + snapshot.officeEquipmentSetup },
    { label: "Salaires employes annee 1", value: snapshot.employeeYear1 }
  ];

  return `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Plan financier previsionnel</div>
          <div class="card-subtitle">Remplissage du modele exact ${EXACT_FORECAST_DOWNLOAD_NAME} a partir des donnees d'OHADA Compta.</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn btn-outline" onclick="shareForecastWorkbook()">Partager le modele</button>
          <button class="btn btn-gold" onclick="downloadForecastWorkbook()">Telecharger le modele rempli</button>
        </div>
      </div>
      <div class="info-box" style="margin-bottom:16px;">
        <strong>Modele cible:</strong> ${EXACT_FORECAST_DOWNLOAD_NAME}<br><br>
        OHADA Compta pre-remplit l'identite du projet, le statut juridique, les contacts, quelques besoins de demarrage, le chiffre d'affaires annualise, un ratio d'achats, les salaires employes et les hypotheses de croissance. Le classeur reste entierement modifiable dans Excel apres export.
      </div>
      <div class="grid-3">
        ${totals.map((item) => `
          <div class="kpi">
            <div class="kpi-label">${item.label}</div>
            <div class="kpi-value" style="font-size:1rem;">${fmt(item.value)}</div>
            <div class="kpi-note">XOF</div>
          </div>
        `).join("")}
      </div>
      <div class="grid-2" style="margin-top:16px;">
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Hypotheses injectees</div>
          <div style="font-size:0.85rem;color:var(--muted);line-height:1.8;">
            Projet: <strong>${projectName}</strong><br>
            Statut juridique: <strong>${snapshot.legalStatus}</strong><br>
            Activite retenue: <strong>${snapshot.salesType}</strong><br>
            Croissance annee 2: <strong>${(snapshot.growthYear2 * 100).toFixed(0)}%</strong><br>
            Croissance annee 3: <strong>${(snapshot.growthYear3 * 100).toFixed(0)}%</strong><br>
            Cout d'achat des marchandises: <strong>${(snapshot.purchaseRatio * 100).toFixed(0)}%</strong>
          </div>
        </div>
        <div class="card" style="background:var(--surface2);">
          <div class="section-title">Points a revoir dans Excel</div>
          <div style="font-size:0.85rem;color:var(--muted);line-height:1.8;">
            Verifiez les investissements de demarrage ligne par ligne, les hypotheses mensuelles de vente, la remuneration du dirigeant, l'ACRE et les delais clients/fournisseurs. Les formes juridiques OHADA qui n'existent pas dans ce modele sont rapprochees du statut le plus proche. Ce modele est pre-rempli automatiquement, mais il doit rester un document de travail ajuste par l'entreprise.
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderDSF() {
  const packetName = getDsfPacketName();
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
          <div class="card-subtitle">Remplissage du modele exact LIASSE.xlsx a partir des donnees de l'application.</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn btn-outline" onclick="shareBfaLiasseFiscale()">Partager LIASSE.xlsx</button>
          <button class="btn btn-gold" onclick="downloadBfaLiasseFiscale()">Telecharger LIASSE.xlsx rempli</button>
        </div>
      </div>
      <div class="info-box" style="margin-bottom:16px;">
        <strong>Packet cible:</strong> ${packetName}<br><br>
        Le telechargement repose desormais sur le modele exact <strong>LIASSE.xlsx</strong> integre au projet. Les champs d'identification et de codification pays/forme/regime sont pre-remplis a partir de vos donnees. Etat d'avancement actuel: <strong>${readyCount}/${statuses.length}</strong> rubriques pretes.
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

  const formesJuridiques = ["SA","SARL","SAS","SASU","EURL","SNC","SCS","GIE","Cooperative","Association","ONG","Fondation","Projet de developpement","Etablissement public","Autre"];
  const regimesFiscaux = ["Reel Normal d'Imposition (RNI)","Reel Simplifie d'Imposition (RSI)","Contribution des Micro-Entreprises (CME)","Forfait d'Imposition","Exonere"];

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Identification de l'entreprise</div>
        <div class="card-subtitle">Ces informations figurent sur tous vos etats financiers OHADA.</div>
      </div>

      <div class="grid-2">
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
  const fields = ['raisonSociale','formeJuridique','sigleUsuel','rccm','nif','nes','siegeSocial','activitePrincipale',
                  'capitalSocial','regimeFiscal','pays','exerciceDu','exerciceAu',
                  'tel','emailCompta','expertComptable','commissaire',
                  'tauxIS','tauxIMF','tauxTVA'];
  fields.forEach(f => {
    const el = document.getElementById('p-' + f);
    if (el) currentCompanyDetails[f] = el.value.trim();
  });
  saveCompanyData();
  // Refresh topbar badge with updated raison sociale
  const badge = document.getElementById('company-name-display');
  if (badge && currentCompanyDetails.raisonSociale) badge.textContent = currentCompanyDetails.raisonSociale;
  const msg = document.getElementById('p-msg');
  if (msg) { msg.style.color = 'var(--green)'; msg.textContent = 'Fiche enregistree avec succes.'; setTimeout(() => { msg.textContent = ''; }, 2500); }
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
checkAuth();

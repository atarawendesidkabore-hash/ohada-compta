// OHADA Compta — Application principale
// SYSCOHADA Revise & SYCEBNL

let currentTab = "dashboard";
let currentRef = "syscohada";
let sycebnlType = "associations"; // 'associations' or 'projets'
let journalEntries = [...SAMPLE_JOURNAL];
let searchTerm = "";
let filterClass = null;
let currentCompanyDetails = {};

// ═══════════════════════════════════════════════════════════
// ACCOUNT MANAGEMENT (multi-company, localStorage)
// ═══════════════════════════════════════════════════════════
const ACCOUNTS_KEY = 'ohada_accounts';
const SESSION_KEY  = 'ohada_session';
const DATA_PREFIX  = 'ohada_data_';
let currentCompanyId = null;

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
  journalEntries = [...SAMPLE_JOURNAL];
  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);
  Object.assign(OPENING_BALANCES, DEFAULT_OPENING_BALANCES);
  currentRef = 'syscohada';
  sycebnlType = 'associations';
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

// Navigation
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    currentTab = btn.dataset.tab;
    closeMobileMenu();
    render();
    window.scrollTo(0, 0);
  });
});

function getPlan() {
  if (currentRef === "sycebnl") {
    const base = PLAN_COMPTABLE_SYSCOHADA.filter(a => a.numero !== "13" && a.numero !== "131" && a.numero !== "139");
    return [...base, ...PLAN_COMPTABLE_SYCEBNL_ADDITIONS, ...PLAN_COMPTABLE_SYCEBNL_CLASSE9].sort((a, b) => a.numero.localeCompare(b.numero));
  }
  return PLAN_COMPTABLE_SYSCOHADA;
}

function fmt(n) { return n.toLocaleString("fr-FR"); }

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
    case "dsf": main.innerHTML = renderDSF(); break;
    case "guide": main.innerHTML = renderGuide(); break;
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
  const totalDebit = journalEntries.reduce((s, e) => s + e.debit, 0);
  const totalCredit = journalEntries.reduce((s, e) => s + e.credit, 0);
  const refLabel = currentRef === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise";

  // Company info
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
  const compName = currentCompanyDetails.raisonSociale || (acct ? acct.company : '');
  const capital = parseFloat(currentCompanyDetails.capitalSocial) || 0;

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
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Referentiel</div><div class="kpi-value" style="font-size:1.2rem;color:var(--gold);">${refLabel}</div><div class="kpi-note">Norme en vigueur</div></div>
      <div class="kpi"><div class="kpi-label">Comptes</div><div class="kpi-value">${plan.length}</div><div class="kpi-note">Plan comptable actif</div></div>
      <div class="kpi"><div class="kpi-label">Classes</div><div class="kpi-value">${currentRef==="sycebnl"?9:8}</div><div class="kpi-note">Classes 1 a ${currentRef==="sycebnl"?9:8}</div></div>
      <div class="kpi"><div class="kpi-label">Ecritures</div><div class="kpi-value">${journalEntries.length}</div><div class="kpi-note">Journal general</div></div>
      <div class="kpi"><div class="kpi-label">Total debit</div><div class="kpi-value" style="color:var(--green);font-size:1.1rem;">${fmt(totalDebit)}</div><div class="kpi-note">XOF</div></div>
      <div class="kpi"><div class="kpi-label">Total credit</div><div class="kpi-value" style="color:var(--red);font-size:1.1rem;">${fmt(totalCredit)}</div><div class="kpi-note">XOF</div></div>
      <div class="kpi"><div class="kpi-label">Equilibre</div><div class="kpi-value" style="color:${totalDebit === totalCredit ? 'var(--green)' : 'var(--red)'};">${totalDebit === totalCredit ? 'OK' : 'ERREUR'}</div><div class="kpi-note">${totalDebit === totalCredit ? 'Debit = Credit' : 'Ecart: ' + fmt(Math.abs(totalDebit - totalCredit))}</div></div>
      <div class="kpi"><div class="kpi-label">Etats OHADA</div><div class="kpi-value">17</div><div class="kpi-note">Pays membres</div></div>
    </div>

    ${(margePct !== null || liquidite !== null || endettement !== null) ? `
    <div class="kpi-grid" style="margin-top:0;">
      ${margePct !== null ? `<div class="kpi"><div class="kpi-label">Marge nette</div><div class="kpi-value" style="color:${parseFloat(margePct)>=0?'var(--green)":'var(--red)'}">${margePct}%</div><div class="kpi-note">${resultat>=0?'Benefice':'Perte'} ${fmt(Math.abs(resultat))} XOF</div></div>` : ''}
      ${liquidite !== null ? `<div class="kpi"><div class="kpi-label">Liquidite generale</div><div class="kpi-value" style="color:${parseFloat(liquidite)>=1?'var(--green)":'var(--red)'}">${liquidite}</div><div class="kpi-note">${parseFloat(liquidite)>=1?'Solvable':'Risque liquidite'}</div></div>` : ''}
      ${endettement !== null ? `<div class="kpi"><div class="kpi-label">Taux endettement</div><div class="kpi-value" style="color:${parseFloat(endettement)<=100?'var(--green)":'var(--red)'}">${endettement}%</div><div class="kpi-note">Capital social: ${fmt(capital)} XOF</div></div>` : ''}
      ${capital > 0 ? `<div class="kpi"><div class="kpi-label">Rentabilite CP</div><div class="kpi-value" style="color:${resultat>=0?'var(--green)":'var(--red)'}">${capital > 0 ? (resultat/capital*100).toFixed(1)+'%' : '—'}</div><div class="kpi-note">Resultat / Capital</div></div>` : ''}
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
        <strong>Etats financiers:</strong> Bilan | Compte de resultat | TAFIRE | Notes annexes<br>
        <strong>Pays membres:</strong> ${OHADA_MEMBER_STATES.join(", ")}
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
          <div class="card-subtitle">${journalEntries.length} ecritures | Equilibre: ${totalD === totalC ? 'OK' : 'ERREUR'}</div>
        </div>
        <button class="btn btn-gold" onclick="document.querySelectorAll('.nav-btn').forEach(b=>{if(b.dataset.tab==='saisie'){b.click();}});">+ Nouvelle ecriture</button>
      </div>
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
          <div class="card-subtitle">${comptes.length} comptes | S.O.: ${totN1D===totN1C?'OK':'ERR'} | Mvt: ${totMvtD===totMvtC?'OK':'ERREUR'}</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <label class="btn btn-outline" style="cursor:pointer;font-size:0.78rem;">
            Importer CSV <input type="file" accept=".csv" style="display:none;" onchange="importBalanceCsv(this)">
          </label>
          <button class="btn btn-outline" style="font-size:0.78rem;" onclick="downloadBalanceTemplate()">Modele CSV</button>
        </div>
      </div>
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
    </div>
  `;
}

function importBalanceCsv(input) {
  const file = input.files[0];
  if (!file) return;
  handleDroppedFile(file);
}

function downloadBalanceTemplate() {
  const plan = getPlan();
  let csv = 'Compte,Libelle,S.O. Debit,S.O. Credit\n';
  plan.forEach(a => {
    const ob = OPENING_BALANCES[a.numero]||{n1d:0,n1c:0};
    csv += '"' + a.numero + '","' + a.libelle + '",' + ob.n1d + ',' + ob.n1c + '\n';
  });
  const blob = new Blob([csv], {type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'balance_ouverture_template.csv'; a.click();
  URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════
// BILAN
// ═══════════════════════════════════════════════════════════
function renderBilan() {
  const bal = computeBalances();

  // sum debit-credit for a list of account prefixes (net value)
  function sumByPrefix(prefixes) {
    let total = 0;
    Object.keys(bal).forEach(code => {
      if (prefixes.some(p => code.startsWith(String(p)))) {
        total += bal[code].debit - bal[code].credit;
      }
    });
    return total;
  }

  // Actif: net value = accounts + amortissements/depreciations (debit-credit auto-nets)
  const actifRows = BILAN_STRUCTURE.actif.map(s => ({
    section: s.section,
    total: sumByPrefix([...s.accounts, ...(s.amort||[])])
  }));

  // Passif: negate (credit-normal accounts)
  const passifRows = BILAN_STRUCTURE.passif.map(s => ({
    section: s.section,
    total: -sumByPrefix(s.accounts)
  }));

  // Compute exercise result from gestion account MOVEMENTS (not opening balances)
  let produits = 0, charges = 0;
  Object.keys(bal).forEach(code => {
    const mvtD = (bal[code].debit||0) - (bal[code].n1d||0);
    const mvtC = (bal[code].credit||0) - (bal[code].n1c||0);
    const c = parseInt(code[0]);
    if (c === 7 || code.startsWith('82') || code.startsWith('84')) produits += mvtC - mvtD;
    if (c === 6 || code.startsWith('81') || code.startsWith('83') ||
        code.startsWith('85') || code.startsWith('87') || code.startsWith('89')) charges += mvtD - mvtC;
  });
  const resultatExercice = produits - charges;

  // Add result to CAPITAUX PROPRES (first passif row)
  if (passifRows.length > 0) passifRows[0].total += resultatExercice;

  const totalActif = actifRows.reduce((s, r) => s + r.total, 0);
  const totalPassif = passifRows.reduce((s, r) => s + r.total, 0);
  const isBalanced = Math.abs(totalActif - totalPassif) < 1;

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Bilan — ${currentRef === "sycebnl" ? "SYCEBNL" : "SYSCOHADA Revise"}</div>
        <div class="card-subtitle" style="color:${isBalanced?'var(--green)':'var(--red)'};">Equilibre: ${isBalanced?'OK ✓':'ERREUR — ecart '+fmt(Math.abs(totalActif-totalPassif))}</div>
      </div>
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
// TAFIRE
// ═══════════════════════════════════════════════════════════
function renderTafire() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">TAFIRE — Tableau Financier des Ressources et Emplois</div></div>
      <div class="info-box">
        Le TAFIRE est l'etat financier OHADA equivalent au tableau des flux de tresorerie. Il presente les flux de ressources et d'emplois en deux parties:<br><br>
        <strong>Partie I:</strong> Determination des soldes financiers de l'exercice (CAF, variation BFR, ETE)<br>
        <strong>Partie II:</strong> Tableau des emplois et ressources (investissements, financements, variation de tresorerie)<br><br>
        <em>Le calcul automatique sera disponible apres cloture de l'exercice avec les donnees completes.</em>
      </div>
      <div class="grid-2" style="margin-top:16px;">
        <div class="card" style="border-color:var(--green);">
          <div style="color:var(--green);font-weight:700;margin-bottom:8px;">RESSOURCES</div>
          <div style="font-size:0.84rem;color:var(--muted);line-height:1.8;">
            Capacite d'Autofinancement Globale (CAFG)<br>
            Cessions et reductions d'immobilisations<br>
            Augmentations de capitaux propres<br>
            Augmentations de dettes financieres<br>
            Diminution du BFR<br>
          </div>
        </div>
        <div class="card" style="border-color:var(--red);">
          <div style="color:var(--red);font-weight:700;margin-bottom:8px;">EMPLOIS</div>
          <div style="font-size:0.84rem;color:var(--muted);line-height:1.8;">
            Investissements (acquisitions d'immobilisations)<br>
            Remboursements de dettes financieres<br>
            Dividendes distribues<br>
            Augmentation du BFR<br>
            Variation de tresorerie<br>
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
        <button class="btn btn-outline" style="font-size:0.78rem;" onclick="exportAmortCsv()">Export CSV</button>
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
  const bal = computeBalances();
  const plan = getPlan();
  const durees={"21":5,"211":5,"212":5,"213":10,"214":10,"22":0,"23":20,"231":20,"232":20,"234":10,"235":7,"24":5,"241":5,"242":5,"244":3,"245":5,"246":7,"248":5};
  function ac(c){if(c.startsWith("21"))return"281";if(c.startsWith("22"))return"282";if(c.startsWith("23"))return"283";return"284";}
  let csv='Compte,Libelle,VBO,Duree,Taux%,Cumul N-1,Dotation N,Cumul N,VNC\n';
  plan.filter(a=>a.classe===2&&a.sens==="Debit"&&bal[a.numero]).forEach(a=>{
    const vbo=bal[a.numero].debit; if(!vbo) return;
    const cn1=bal[ac(a.numero)]?(bal[ac(a.numero)].n1c||0):0;
    const d=durees[a.numero]||5;
    const dot=d>0?Math.round(vbo/d):0;
    csv+='"'+a.numero+'","'+a.libelle+'",'+vbo+','+d+','+(d>0?Math.round(10000/d)/100:0)+','+cn1+','+dot+','+(cn1+dot)+','+Math.max(0,vbo-cn1-dot)+'\n';
  });
  const blob=new Blob([csv],{type:'text/csv'});
  const url=URL.createObjectURL(blob);
  const el=document.createElement('a');el.href=url;el.download='amortissements.csv';el.click();
  URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════
// NOTES ANNEXES
// ═══════════════════════════════════════════════════════════
function renderAnnexes() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Notes annexes aux etats financiers</div></div>
      <div class="info-box" style="margin-bottom:16px;">
        Les notes annexes font partie integrante des etats financiers OHADA. Elles completent et commentent les informations des autres etats financiers.
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
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Cloture de l'exercice</div></div>
      <div class="info-box">
        La cloture de l'exercice comptable OHADA comprend les etapes suivantes:<br><br>
        1. <strong>Inventaire physique</strong> — Verification des stocks, immobilisations, tresorerie<br>
        2. <strong>Ecritures de regularisation</strong> — Amortissements, provisions, charges constatees d'avance<br>
        3. <strong>Balance apres inventaire</strong> — Verification de l'equilibre<br>
        4. <strong>Determination du resultat</strong> — Solde des comptes de gestion<br>
        5. <strong>Etablissement des etats financiers</strong> — Bilan, Resultat, TAFIRE, Annexes<br>
        6. <strong>Ecritures de cloture</strong> — A-nouveaux pour l'exercice suivant<br>
        7. <strong>Declaration DSF/DGI</strong> — Liasse fiscale obligatoire
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
// DSF / DGI
// ═══════════════════════════════════════════════════════════
function renderDSF() {
  return `
    <div class="card">
      <div class="card-header"><div class="card-title">Declaration Statistique et Fiscale (DSF)</div></div>
      <div class="info-box" style="margin-bottom:16px;">
        La DSF est la declaration fiscale annuelle obligatoire dans les pays OHADA. Elle comprend la liasse comptable normalisee conforme au SYSCOHADA, transmise a la Direction Generale des Impots (DGI).
      </div>
      <div class="stack">
        ${[
          { code: "DSF-01", label: "Bilan — Systeme normal", status: "Pret" },
          { code: "DSF-02", label: "Compte de resultat — Systeme normal", status: "Pret" },
          { code: "DSF-03", label: "TAFIRE", status: "En attente" },
          { code: "DSF-04", label: "Tableau des immobilisations", status: "En attente" },
          { code: "DSF-05", label: "Tableau des amortissements", status: "En attente" },
          { code: "DSF-06", label: "Tableau des provisions", status: "En attente" },
          { code: "DSF-07", label: "Etat des creances et dettes", status: "En attente" },
          { code: "DSF-08", label: "Tableau des resultat et soldes intermediaires", status: "Pret" },
          { code: "DSF-09", label: "Notes annexes", status: "En attente" },
          { code: "DSF-10", label: "Informations complementaires DGI", status: "En attente" },
        ].map(d => `
          <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);">
            <div>
              <span style="font-family:var(--mono);font-weight:700;color:var(--gold);margin-right:10px;">${d.code}</span>
              <span>${d.label}</span>
            </div>
            <span style="font-size:0.78rem;font-weight:700;color:${d.status === 'Pret' ? 'var(--green)' : 'var(--orange)'};">${d.status}</span>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
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
            ["TAFIRE", "Obligatoire (systeme normal)", "Adapte aux flux de projets"],
            ["DSF / DGI", "Liasse fiscale standard", "Liasse adaptee EBNL"],
            ["Etats financiers", "Bilan, Resultat, TAFIRE, Annexes", "Idem avec adaptations EBNL"],
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
      if (totalD !== totalC) { msg.innerHTML = `<span style="color:var(--red);">Desequilibre: Debit ${fmt(totalD)} != Credit ${fmt(totalC)}</span>`; return; }
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
      const count = parseBalanceCsv(e.target.result);
      if (count > 0) {
        showToast(count + ' comptes importes depuis ' + file.name, 'success');
        render();
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
            OPENING_BALANCES[code] = { n1d, n1c };
            count++;
          }
        });
        if (count > 0) {
          showToast(count + ' comptes importes depuis ' + file.name, 'success');
          render();
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
        let count = 0;
        entries.forEach(entry => {
          const d = parseFloat(entry.debit) || 0;
          const c = parseFloat(entry.credit) || 0;
          if (!entry.compte || (!d && !c)) return;
          journalEntries.push({
            id: journalEntries.length + 1,
            date: entry.date || new Date().toISOString().split('T')[0],
            journal: entry.journal || 'OD',
            piece: entry.piece || ('IMP-' + String(journalEntries.length + 1).padStart(3, '0')),
            compte: String(entry.compte),
            libelle: entry.libelle || 'Import',
            debit: d, credit: c,
            ref: entry.ref || 'IMPORT'
          });
          count++;
        });
        if (count > 0) {
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
  let count = 0;
  lines.forEach(line => {
    const cols = line.split(',');
    if (cols.length < 4) return;
    const code = cols[0].trim().replace(/"/g, '');
    const n1d = parseFloat(cols[2]) || 0;
    const n1c = parseFloat(cols[3]) || 0;
    if (code && /^\d+$/.test(code)) {
      OPENING_BALANCES[code] = { n1d, n1c };
      count++;
    }
  });
  return count;
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

  const formesJuridiques = ['SA','SARL','SAS','SASU','EURL','SNC','SCS','GIE','Cooperative','Association','ONG','Fondation','Projet de developpement','Etablissement public','Autre'];
  const regimesFiscaux = ['Reel Normal d'Imposition (RNI)','Reel Simplifie d'Imposition (RSI)','Contribution des Micro-Entreprises (CME)','Forfait d'Imposition','Exonere'];

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Identification de l'entreprise</div>
        <div class="card-subtitle">Ces informations figurent sur tous vos etats financiers OHADA.</div>
      </div>

      <div class="grid-2">
        <div class="form-group">
          <div class="form-label">Raison sociale</div>
          <input class="form-input" id="p-raisonSociale" value="${d.raisonSociale || (acct ? acct.company : '')}" placeholder="Ex: WASI Ecosystem SAS">
        </div>
        <div class="form-group">
          <div class="form-label">Forme juridique</div>
          <select class="form-select" id="p-formeJuridique">
            <option value="">-- Choisir --</option>
            ${formesJuridiques.map(f => `<option value="${f}" ${(d.formeJuridique||'')===f?'selected':''}>${f}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <div class="form-label">RCCM</div>
          <input class="form-input" id="p-rccm" value="${d.rccm || ''}" placeholder="Ex: BF-OUA-2020-B-12345">
        </div>
        <div class="form-group">
          <div class="form-label">NIF (Numero d'Identification Fiscale)</div>
          <input class="form-input" id="p-nif" value="${d.nif || ''}" placeholder="Ex: 00123456A">
        </div>
        <div class="form-group" style="grid-column:1/-1;">
          <div class="form-label">Siege social</div>
          <input class="form-input" id="p-siegeSocial" value="${d.siegeSocial || ''}" placeholder="Ex: 12 Avenue Kwame Nkrumah, Ouagadougou, Burkina Faso">
        </div>
        <div class="form-group" style="grid-column:1/-1;">
          <div class="form-label">Activite principale</div>
          <input class="form-input" id="p-activitePrincipale" value="${d.activitePrincipale || ''}" placeholder="Ex: Commerce general de produits alimentaires">
        </div>
        <div class="form-group">
          <div class="form-label">Capital social (XOF)</div>
          <input class="form-input" type="number" id="p-capitalSocial" value="${d.capitalSocial || ''}" placeholder="Ex: 10000000">
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
          <input class="form-input" id="p-pays" value="${d.pays || ''}" placeholder="Ex: Burkina Faso">
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
          <input class="form-input" id="p-tel" value="${d.tel || ''}" placeholder="+226 XX XX XX XX">
        </div>
        <div class="form-group">
          <div class="form-label">Email</div>
          <input class="form-input" id="p-emailCompta" value="${d.emailCompta || (acct ? acct.email : '')}" placeholder="compta@entreprise.com">
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
  const fields = ['raisonSociale','formeJuridique','rccm','nif','siegeSocial','activitePrincipale',
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

// Initial auth check (shows login or loads company data)
checkAuth();

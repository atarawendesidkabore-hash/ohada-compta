#!/usr/bin/env python3
# fix_account_issues.py — Fix 3 issues: empty state, company name, company details + ratios

BASE = r'C:\Users\eu\ohada-compta'

# ─── app.js ──────────────────────────────────────────────────────────────────
with open(BASE + r'\app.js', 'r', encoding='utf-8') as f:
    app = f.read()
print(f"app.js: {len(app)} chars")

# 1. Add currentCompanyDetails global
OLD_GLOBALS = 'let filterClass = null;\n'
NEW_GLOBALS = 'let filterClass = null;\nlet currentCompanyDetails = {};\n'
assert OLD_GLOBALS in app, "filterClass global not found"
app = app.replace(OLD_GLOBALS, NEW_GLOBALS, 1)
print("currentCompanyDetails global added.")

# 2. Fix loginCompany: reset to EMPTY state before loading data (new companies start blank)
OLD_LOGIN = (
    'function loginCompany(id) {\n'
    '  currentCompanyId = id;\n'
    '  localStorage.setItem(SESSION_KEY, id);\n'
    '  loadCompanyData(id);\n'
)
NEW_LOGIN = (
    'function loginCompany(id) {\n'
    '  currentCompanyId = id;\n'
    '  localStorage.setItem(SESSION_KEY, id);\n'
    '  // Reset to blank state — new companies start empty, returning ones get their data loaded\n'
    '  journalEntries = [];\n'
    '  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);\n'
    '  currentRef = \'syscohada\';\n'
    '  sycebnlType = \'associations\';\n'
    '  currentCompanyDetails = {};\n'
    '  const selReset = document.getElementById(\'referentiel-select\');\n'
    '  if (selReset) selReset.value = \'syscohada\';\n'
    '  loadCompanyData(id);\n'
)
assert OLD_LOGIN in app, "loginCompany() header not found"
app = app.replace(OLD_LOGIN, NEW_LOGIN, 1)
print("loginCompany() now resets to blank state before loading data.")

# 3. Fix saveCompanyData to include companyDetails
OLD_SAVE = (
    'function saveCompanyData() {\n'
    '  if (!currentCompanyId) return;\n'
    '  const data = {\n'
    '    journalEntries,\n'
    '    openingBalances: Object.assign({}, OPENING_BALANCES),\n'
    '    currentRef,\n'
    '    sycebnlType\n'
    '  };\n'
)
NEW_SAVE = (
    'function saveCompanyData() {\n'
    '  if (!currentCompanyId) return;\n'
    '  const data = {\n'
    '    journalEntries,\n'
    '    openingBalances: Object.assign({}, OPENING_BALANCES),\n'
    '    currentRef,\n'
    '    sycebnlType,\n'
    '    companyDetails: currentCompanyDetails\n'
    '  };\n'
)
assert OLD_SAVE in app, "saveCompanyData() data object not found"
app = app.replace(OLD_SAVE, NEW_SAVE, 1)
print("saveCompanyData() updated to include companyDetails.")

# 4. Fix loadCompanyData to restore companyDetails
OLD_LOAD_REF = (
    '    if (data.currentRef) currentRef = data.currentRef;\n'
    '    if (data.sycebnlType) sycebnlType = data.sycebnlType;\n'
    '    // Sync referentiel select\n'
    '    const sel = document.getElementById(\'referentiel-select\');\n'
    '    if (sel) sel.value = currentRef;\n'
)
NEW_LOAD_REF = (
    '    if (data.currentRef) currentRef = data.currentRef;\n'
    '    if (data.sycebnlType) sycebnlType = data.sycebnlType;\n'
    '    if (data.companyDetails) currentCompanyDetails = data.companyDetails;\n'
    '    // Sync referentiel select\n'
    '    const sel = document.getElementById(\'referentiel-select\');\n'
    '    if (sel) sel.value = currentRef;\n'
)
assert OLD_LOAD_REF in app, "loadCompanyData() ref section not found"
app = app.replace(OLD_LOAD_REF, NEW_LOAD_REF, 1)
print("loadCompanyData() updated to restore companyDetails.")

# 5. Add "parametres" case to render() switch
OLD_RENDER_SWITCH = '    case "comparaison": main.innerHTML = renderComparaison(); break;\n    default: main.innerHTML = renderDashboard();'
NEW_RENDER_SWITCH = '    case "comparaison": main.innerHTML = renderComparaison(); break;\n    case "parametres": main.innerHTML = renderParametres(); break;\n    default: main.innerHTML = renderDashboard();'
assert OLD_RENDER_SWITCH in app, "render() switch comparaison case not found"
app = app.replace(OLD_RENDER_SWITCH, NEW_RENDER_SWITCH, 1)
print("renderParametres() case added to render() switch.")

# 6. Update renderDashboard() to show company name + financial ratios
# Find the start of the KPI grid in renderDashboard and insert company banner above it
OLD_DASHBOARD_START = (
    '  return `\n'
    '    <div class="kpi-grid">\n'
    '      <div class="kpi"><div class="kpi-label">Referentiel</div>'
)
NEW_DASHBOARD_START = (
    '  // Company info\n'
    '  const accounts = getAccounts();\n'
    '  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;\n'
    '  const compName = currentCompanyDetails.raisonSociale || (acct ? acct.company : \'\');\n'
    '  const capital = parseFloat(currentCompanyDetails.capitalSocial) || 0;\n'
    '\n'
    '  // Financial ratios from balance\n'
    '  let totalActifNet = 0, totalPassif = 0, totalProduits = 0, totalCharges = 0;\n'
    '  let totalActifCirc = 0, totalPassifCirc = 0;\n'
    '  Object.keys(bal).forEach(code => {\n'
    '    const net = (bal[code].debit||0) - (bal[code].credit||0);\n'
    '    const c1 = parseInt(code[0]);\n'
    '    if ([1,2,3,4,5].includes(c1)) {\n'
    '      if (net > 0) totalActifNet += net;\n'
    '      if (net < 0) totalPassif += Math.abs(net);\n'
    '    }\n'
    '    if (c1 === 3 || c1 === 4 || c1 === 5) { if (net > 0) totalActifCirc += net; else totalPassifCirc += Math.abs(net); }\n'
    '    if (c1 === 7 || code.startsWith(\'82\') || code.startsWith(\'84\')) totalProduits += (bal[code].credit||0) - (bal[code].debit||0);\n'
    '    if (c1 === 6 || code.startsWith(\'81\') || code.startsWith(\'83\')) totalCharges += (bal[code].debit||0) - (bal[code].credit||0);\n'
    '  });\n'
    '  const resultat = totalProduits - totalCharges;\n'
    '  const margePct = totalProduits > 0 ? (resultat / totalProduits * 100).toFixed(1) : null;\n'
    '  const liquidite = totalPassifCirc > 0 ? (totalActifCirc / totalPassifCirc).toFixed(2) : null;\n'
    '  const endettement = capital > 0 ? ((totalPassif / capital) * 100).toFixed(1) : null;\n'
    '\n'
    '  return `\n'
    '    ${compName ? `<div class="company-header"><div class="company-header-name">${compName}</div><div class="company-header-meta">${currentCompanyDetails.formeJuridique||\'\'} ${currentCompanyDetails.ville ? \'&bull; \'+currentCompanyDetails.ville : \'\'} ${currentCompanyDetails.nif ? \'&bull; NIF: \'+currentCompanyDetails.nif : \'\'}</div></div>` : \'\'}\n'
    '    <div class="kpi-grid">\n'
    '      <div class="kpi"><div class="kpi-label">Referentiel</div>'
)
assert OLD_DASHBOARD_START in app, "renderDashboard() return start not found"
app = app.replace(OLD_DASHBOARD_START, NEW_DASHBOARD_START, 1)
print("renderDashboard() updated with company banner.")

# 7. Add ratios KPIs to dashboard after existing KPIs, before Classes card
OLD_AFTER_KPI_GRID = (
    '      <div class="kpi"><div class="kpi-label">Etats OHADA</div><div class="kpi-value">17</div><div class="kpi-note">Pays membres</div></div>\n'
    '    </div>\n'
    '\n'
    '    <div class="grid-2">'
)
NEW_AFTER_KPI_GRID = (
    '      <div class="kpi"><div class="kpi-label">Etats OHADA</div><div class="kpi-value">17</div><div class="kpi-note">Pays membres</div></div>\n'
    '    </div>\n'
    '\n'
    '    ${(margePct !== null || liquidite !== null || endettement !== null) ? `\n'
    '    <div class="kpi-grid" style="margin-top:0;">\n'
    '      ${margePct !== null ? `<div class="kpi"><div class="kpi-label">Marge nette</div><div class="kpi-value" style="color:${parseFloat(margePct)>=0?\'var(--green)\":\'var(--red)\'}">${margePct}%</div><div class="kpi-note">${resultat>=0?\'Benefice\':\'Perte\'} ${fmt(Math.abs(resultat))} XOF</div></div>` : \'\'}\n'
    '      ${liquidite !== null ? `<div class="kpi"><div class="kpi-label">Liquidite generale</div><div class="kpi-value" style="color:${parseFloat(liquidite)>=1?\'var(--green)\":\'var(--red)\'}">${liquidite}</div><div class="kpi-note">${parseFloat(liquidite)>=1?\'Solvable\':\'Risque liquidite\'}</div></div>` : \'\'}\n'
    '      ${endettement !== null ? `<div class="kpi"><div class="kpi-label">Taux endettement</div><div class="kpi-value" style="color:${parseFloat(endettement)<=100?\'var(--green)\":\'var(--red)\'}">${endettement}%</div><div class="kpi-note">Capital social: ${fmt(capital)} XOF</div></div>` : \'\'}\n'
    '      ${capital > 0 ? `<div class="kpi"><div class="kpi-label">Rentabilite CP</div><div class="kpi-value" style="color:${resultat>=0?\'var(--green)\":\'var(--red)\'}">${capital > 0 ? (resultat/capital*100).toFixed(1)+\'%\' : \'—\'}</div><div class="kpi-note">Resultat / Capital</div></div>` : \'\'}\n'
    '    </div>` : \'\'}\n'
    '\n'
    '    <div class="grid-2">'
)
assert OLD_AFTER_KPI_GRID in app, "kpi-grid end not found"
app = app.replace(OLD_AFTER_KPI_GRID, NEW_AFTER_KPI_GRID, 1)
print("Financial ratios KPIs added to dashboard.")

# 8. Add renderParametres() function before the final // Initial auth check comment
PARAMETRES_FN = '''
// ═══════════════════════════════════════════════════════════
// PARAMETRES ENTREPRISE
// ═══════════════════════════════════════════════════════════
function renderParametres() {
  const d = currentCompanyDetails;
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;
  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Fiche entreprise / organisation</div>
        <div class="card-subtitle">Ces informations apparaissent sur vos etats financiers et calculent vos ratios.</div>
      </div>
      <div class="grid-2">
        <div class="form-group">
          <div class="form-label">Raison sociale</div>
          <input class="form-input" id="p-raisonSociale" value="${d.raisonSociale || (acct ? acct.company : '')}" placeholder="Ex: WASI Ecosystem SAS">
        </div>
        <div class="form-group">
          <div class="form-label">Forme juridique</div>
          <select class="form-select" id="p-formeJuridique">
            ${['SA','SARL','SAS','EURL','SNC','GIE','Association','ONG','Fondation','Projet','Autre'].map(f =>
              `<option value="${f}" ${(d.formeJuridique||'') === f ? 'selected' : ''}>${f}</option>`
            ).join('')}
          </select>
        </div>
        <div class="form-group">
          <div class="form-label">Secteur d\'activite</div>
          <input class="form-input" id="p-secteur" value="${d.secteur || ''}" placeholder="Ex: Commerce, Agriculture, BTP...">
        </div>
        <div class="form-group">
          <div class="form-label">Capital social (XOF)</div>
          <input class="form-input" type="number" id="p-capitalSocial" value="${d.capitalSocial || ''}" placeholder="Ex: 10000000">
        </div>
        <div class="form-group">
          <div class="form-label">NIF / TIN</div>
          <input class="form-input" id="p-nif" value="${d.nif || ''}" placeholder="Numero d\'identification fiscale">
        </div>
        <div class="form-group">
          <div class="form-label">RCCM</div>
          <input class="form-input" id="p-rccm" value="${d.rccm || ''}" placeholder="Registre du commerce">
        </div>
        <div class="form-group">
          <div class="form-label">Annee d\'exercice</div>
          <input class="form-input" type="number" id="p-exercice" value="${d.exercice || new Date().getFullYear()}" placeholder="${new Date().getFullYear()}">
        </div>
        <div class="form-group">
          <div class="form-label">Pays</div>
          <input class="form-input" id="p-pays" value="${d.pays || ''}" placeholder="Ex: Burkina Faso, Cote d\'Ivoire...">
        </div>
        <div class="form-group">
          <div class="form-label">Ville</div>
          <input class="form-input" id="p-ville" value="${d.ville || ''}" placeholder="Ex: Ouagadougou, Abidjan...">
        </div>
        <div class="form-group">
          <div class="form-label">Adresse</div>
          <input class="form-input" id="p-adresse" value="${d.adresse || ''}" placeholder="Ex: 12 Avenue Kwame Nkrumah">
        </div>
        <div class="form-group">
          <div class="form-label">Telephone</div>
          <input class="form-input" id="p-tel" value="${d.tel || ''}" placeholder="+226 ...">
        </div>
        <div class="form-group">
          <div class="form-label">Email comptabilite</div>
          <input class="form-input" id="p-emailCompta" value="${d.emailCompta || (acct ? acct.email : '')}" placeholder="compta@entreprise.com">
        </div>
      </div>
      <div id="p-msg" style="margin-top:12px;min-height:20px;font-size:0.82rem;"></div>
      <button class="btn btn-gold" style="margin-top:8px;" onclick="saveParametres()">Enregistrer la fiche</button>
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
  const fields = ['raisonSociale','formeJuridique','secteur','capitalSocial','nif','rccm','exercice','pays','ville','adresse','tel','emailCompta'];
  fields.forEach(f => {
    const el = document.getElementById('p-' + f);
    if (el) currentCompanyDetails[f] = el.value.trim();
  });
  saveCompanyData();
  const msg = document.getElementById('p-msg');
  if (msg) { msg.style.color = 'var(--green)'; msg.textContent = 'Fiche enregistree avec succes.'; setTimeout(() => { msg.textContent = ''; }, 2500); }
}

'''

OLD_INIT_COMMENT = '// Initial auth check (shows login or loads company data)\ncheckAuth();'
assert OLD_INIT_COMMENT in app, "checkAuth() call not found"
app = app.replace(OLD_INIT_COMMENT, PARAMETRES_FN + OLD_INIT_COMMENT, 1)
print("renderParametres() + saveParametres() added.")

with open(BASE + r'\app.js', 'w', encoding='utf-8') as f:
    f.write(app)
print(f"app.js written: {len(app)} chars\n")


# ─── index.html: add Parametres nav button ───────────────────────────────────
with open(BASE + r'\index.html', 'r', encoding='utf-8') as f:
    html = f.read()
print(f"index.html: {len(html)} chars")

OLD_NAV_GUIDE = '        <button class="nav-btn" data-tab="guide">Guide OHADA</button>'
NEW_NAV_GUIDE = (
    '        <button class="nav-btn" data-tab="guide">Guide OHADA</button>\n'
    '        <button class="nav-btn" data-tab="parametres">Parametres entreprise</button>'
)
assert OLD_NAV_GUIDE in html, "Guide OHADA nav button not found"
html = html.replace(OLD_NAV_GUIDE, NEW_NAV_GUIDE, 1)
print("Parametres nav button added.")

with open(BASE + r'\index.html', 'w', encoding='utf-8') as f:
    f.write(html)
print(f"index.html written: {len(html)} chars\n")


# ─── styles.css: add company header styles ───────────────────────────────────
with open(BASE + r'\styles.css', 'r', encoding='utf-8') as f:
    css = f.read()
print(f"styles.css: {len(css)} chars")

COMPANY_HEADER_CSS = """
/* ========== COMPANY HEADER (dashboard) ========== */
.company-header {
  background: linear-gradient(135deg, rgba(200,146,42,0.08) 0%, rgba(200,146,42,0.04) 100%);
  border: 1px solid rgba(200,146,42,0.25);
  border-radius: var(--radius);
  padding: 16px 24px;
  margin-bottom: 20px;
}
.company-header-name {
  font-size: 1.3rem;
  font-weight: 800;
  color: var(--gold);
  font-family: var(--mono);
  letter-spacing: 0.04em;
}
.company-header-meta {
  font-size: 0.8rem;
  color: var(--muted);
  margin-top: 4px;
  letter-spacing: 0.03em;
}
"""

css = css + COMPANY_HEADER_CSS
with open(BASE + r'\styles.css', 'w', encoding='utf-8') as f:
    f.write(css)
print(f"styles.css written: {len(css)} chars\n")

print("=" * 50)
print("Fixes applied:")
print("  1. New accounts start EMPTY (no sample journal entries)")
print("  2. Company name shown in dashboard header + topbar badge")
print("  3. Parametres form: NIF, RCCM, capital, forme juridique, ville...")
print("  4. Financial ratios on dashboard: marge nette, liquidite, endettement, rentabilite")
print("  5. Ratios calculated automatically from journal + capital social")

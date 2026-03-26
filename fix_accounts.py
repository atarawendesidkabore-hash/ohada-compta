#!/usr/bin/env python3
# fix_accounts.py — Multi-company account system (localStorage) for ohada-compta

BASE = r'C:\Users\eu\ohada-compta'

# ─── index.html ──────────────────────────────────────────────────────────────
with open(BASE + r'\index.html', 'r', encoding='utf-8') as f:
    html = f.read()
print(f"index.html: {len(html)} chars")

# 1. Add company display + logout to topbar-right
OLD_TOPBAR_RIGHT = (
    '    <div class="topbar-right">\n'
    '      <select id="referentiel-select" class="ref-select">\n'
    '        <option value="syscohada">SYSCOHADA Revise</option>\n'
    '        <option value="sycebnl">SYCEBNL</option>\n'
    '      </select>\n'
    '      <span class="topbar-meta" id="clock"></span>\n'
    '    </div>'
)
NEW_TOPBAR_RIGHT = (
    '    <div class="topbar-right">\n'
    '      <span id="company-name-display" class="company-badge" style="display:none;"></span>\n'
    '      <select id="referentiel-select" class="ref-select">\n'
    '        <option value="syscohada">SYSCOHADA Revise</option>\n'
    '        <option value="sycebnl">SYCEBNL</option>\n'
    '      </select>\n'
    '      <span class="topbar-meta" id="clock"></span>\n'
    '      <button id="btn-logout" class="btn-logout" style="display:none;" onclick="logoutCompany()">Deconnexion</button>\n'
    '    </div>'
)
assert OLD_TOPBAR_RIGHT in html, "topbar-right not found"
html = html.replace(OLD_TOPBAR_RIGHT, NEW_TOPBAR_RIGHT, 1)
print("Company badge + logout button added to topbar.")

# 2. Add auth overlay before </body>
OLD_BODY_END = '  <!-- Toast notifications -->\n  <div id="toast-container"></div>\n</body>'
NEW_BODY_END = (
    '  <!-- Toast notifications -->\n'
    '  <div id="toast-container"></div>\n'
    '  <!-- Auth overlay -->\n'
    '  <div id="auth-overlay" class="auth-overlay">\n'
    '    <div class="auth-card">\n'
    '      <div class="auth-logo">OHADA<span class="accent">COMPTA</span></div>\n'
    '      <div class="auth-subtitle">Comptabilite SYSCOHADA &amp; SYCEBNL</div>\n'
    '      <div class="auth-tabs">\n'
    '        <button class="auth-tab active" id="tab-login" onclick="showAuthTab(\'login\')">Se connecter</button>\n'
    '        <button class="auth-tab" id="tab-register" onclick="showAuthTab(\'register\')">Creer un compte</button>\n'
    '      </div>\n'
    '      <!-- Login form -->\n'
    '      <div id="form-login">\n'
    '        <div class="form-group"><div class="form-label">Email</div><input class="form-input" type="email" id="login-email" placeholder="contact@entreprise.com" autocomplete="email"></div>\n'
    '        <div class="form-group" style="margin-top:12px;"><div class="form-label">Mot de passe</div><input class="form-input" type="password" id="login-pass" placeholder="Mot de passe" autocomplete="current-password"></div>\n'
    '        <div id="login-msg" class="auth-msg"></div>\n'
    '        <button class="btn btn-gold" style="width:100%;margin-top:16px;" onclick="handleLogin()">Se connecter</button>\n'
    '      </div>\n'
    '      <!-- Register form -->\n'
    '      <div id="form-register" style="display:none;">\n'
    '        <div class="form-group"><div class="form-label">Nom de l\'entreprise / organisation</div><input class="form-input" type="text" id="reg-company" placeholder="Ex: WASI SAS, ONG Espoir, ..."></div>\n'
    '        <div class="form-group" style="margin-top:12px;"><div class="form-label">Email</div><input class="form-input" type="email" id="reg-email" placeholder="contact@entreprise.com" autocomplete="email"></div>\n'
    '        <div class="form-group" style="margin-top:12px;"><div class="form-label">Mot de passe</div><input class="form-input" type="password" id="reg-pass" placeholder="Choisir un mot de passe" autocomplete="new-password"></div>\n'
    '        <div class="form-group" style="margin-top:12px;"><div class="form-label">Confirmer le mot de passe</div><input class="form-input" type="password" id="reg-pass2" placeholder="Confirmer le mot de passe" autocomplete="new-password"></div>\n'
    '        <div id="reg-msg" class="auth-msg"></div>\n'
    '        <button class="btn btn-gold" style="width:100%;margin-top:16px;" onclick="handleRegister()">Creer le compte</button>\n'
    '      </div>\n'
    '      <div class="auth-footer">Vos donnees sont stockees localement sur votre appareil.</div>\n'
    '    </div>\n'
    '  </div>\n'
    '</body>'
)
assert OLD_BODY_END in html, "</body> sequence not found"
html = html.replace(OLD_BODY_END, NEW_BODY_END, 1)
print("Auth overlay added to HTML.")

with open(BASE + r'\index.html', 'w', encoding='utf-8') as f:
    f.write(html)
print(f"index.html written: {len(html)} chars\n")


# ─── styles.css ──────────────────────────────────────────────────────────────
with open(BASE + r'\styles.css', 'r', encoding='utf-8') as f:
    css = f.read()
print(f"styles.css: {len(css)} chars")

AUTH_CSS = """
/* ========== AUTH OVERLAY ========== */
.auth-overlay {
  display: none;
  position: fixed; inset: 0; z-index: 99999;
  background: var(--bg);
  align-items: center; justify-content: center;
}
.auth-overlay.active { display: flex; }
.auth-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 20px;
  padding: 40px 48px;
  width: 100%;
  max-width: 420px;
  box-shadow: 0 24px 64px rgba(0,0,0,0.5);
}
.auth-logo {
  font-size: 1.9rem;
  font-weight: 900;
  letter-spacing: 0.06em;
  color: var(--text);
  text-align: center;
  font-family: var(--mono);
  margin-bottom: 4px;
}
.auth-subtitle {
  text-align: center;
  color: var(--muted);
  font-size: 0.82rem;
  margin-bottom: 28px;
  letter-spacing: 0.04em;
}
.auth-tabs {
  display: flex;
  border-bottom: 1px solid var(--border);
  margin-bottom: 24px;
}
.auth-tab {
  flex: 1;
  background: none;
  border: none;
  color: var(--muted);
  font-size: 0.9rem;
  font-weight: 600;
  padding: 10px;
  cursor: pointer;
  transition: color .2s;
  border-bottom: 2px solid transparent;
  margin-bottom: -1px;
}
.auth-tab.active { color: var(--gold); border-bottom-color: var(--gold); }
.auth-tab:hover { color: var(--text); }
.auth-msg {
  min-height: 20px;
  font-size: 0.82rem;
  margin-top: 8px;
  text-align: center;
}
.auth-msg.error { color: var(--red); }
.auth-msg.success { color: var(--green); }
.auth-footer {
  margin-top: 20px;
  text-align: center;
  font-size: 0.75rem;
  color: var(--muted);
  opacity: 0.7;
}

/* ========== COMPANY BADGE ========== */
.company-badge {
  background: rgba(200,146,42,0.14);
  border: 1px solid var(--gold);
  color: var(--gold);
  font-size: 0.78rem;
  font-weight: 700;
  padding: 4px 12px;
  border-radius: 20px;
  max-width: 180px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-family: var(--mono);
}
.btn-logout {
  background: none;
  border: 1px solid var(--border);
  color: var(--muted);
  font-size: 0.78rem;
  padding: 4px 12px;
  border-radius: 8px;
  cursor: pointer;
  transition: color .2s, border-color .2s;
}
.btn-logout:hover { color: var(--red); border-color: var(--red); }

"""

css = css + AUTH_CSS
print("Auth CSS added.")

with open(BASE + r'\styles.css', 'w', encoding='utf-8') as f:
    f.write(css)
print(f"styles.css written: {len(css)} chars\n")


# ─── app.js ──────────────────────────────────────────────────────────────────
with open(BASE + r'\app.js', 'r', encoding='utf-8') as f:
    app = f.read()
print(f"app.js: {len(app)} chars")

# 1. After global vars, inject account management code
OLD_VARS_END = (
    'let journalEntries = [...SAMPLE_JOURNAL];\n'
    'let searchTerm = "";\n'
    'let filterClass = null;'
)
NEW_VARS_END = (
    'let journalEntries = [...SAMPLE_JOURNAL];\n'
    'let searchTerm = "";\n'
    'let filterClass = null;\n'
    '\n'
    '// ═══════════════════════════════════════════════════════════\n'
    '// ACCOUNT MANAGEMENT (multi-company, localStorage)\n'
    '// ═══════════════════════════════════════════════════════════\n'
    'const ACCOUNTS_KEY = \'ohada_accounts\';\n'
    'const SESSION_KEY  = \'ohada_session\';\n'
    'const DATA_PREFIX  = \'ohada_data_\';\n'
    'let currentCompanyId = null;\n'
    '\n'
    'function simpleHash(str) {\n'
    '  let h = 5381;\n'
    '  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);\n'
    '  return (h >>> 0).toString(16);\n'
    '}\n'
    '\n'
    'function getAccounts() {\n'
    '  try { return JSON.parse(localStorage.getItem(ACCOUNTS_KEY) || \'[]\'); }\n'
    '  catch(e) { return []; }\n'
    '}\n'
    'function saveAccounts(arr) { localStorage.setItem(ACCOUNTS_KEY, JSON.stringify(arr)); }\n'
    '\n'
    'function saveCompanyData() {\n'
    '  if (!currentCompanyId) return;\n'
    '  const data = {\n'
    '    journalEntries,\n'
    '    openingBalances: Object.assign({}, OPENING_BALANCES),\n'
    '    currentRef,\n'
    '    sycebnlType\n'
    '  };\n'
    '  localStorage.setItem(DATA_PREFIX + currentCompanyId, JSON.stringify(data));\n'
    '}\n'
    '\n'
    'function loadCompanyData(id) {\n'
    '  const raw = localStorage.getItem(DATA_PREFIX + id);\n'
    '  if (!raw) return; // fresh company, keep defaults\n'
    '  try {\n'
    '    const data = JSON.parse(raw);\n'
    '    if (data.journalEntries) journalEntries = data.journalEntries;\n'
    '    if (data.openingBalances) {\n'
    '      // Clear and repopulate OPENING_BALANCES (it is const but mutable)\n'
    '      Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);\n'
    '      Object.assign(OPENING_BALANCES, data.openingBalances);\n'
    '    }\n'
    '    if (data.currentRef) currentRef = data.currentRef;\n'
    '    if (data.sycebnlType) sycebnlType = data.sycebnlType;\n'
    '    // Sync referentiel select\n'
    '    const sel = document.getElementById(\'referentiel-select\');\n'
    '    if (sel) sel.value = currentRef;\n'
    '  } catch(e) { console.warn(\'loadCompanyData error\', e); }\n'
    '}\n'
    '\n'
    'function loginCompany(id) {\n'
    '  currentCompanyId = id;\n'
    '  localStorage.setItem(SESSION_KEY, id);\n'
    '  loadCompanyData(id);\n'
    '  // Update topbar\n'
    '  const accounts = getAccounts();\n'
    '  const acct = accounts.find(a => a.id === id);\n'
    '  const badge = document.getElementById(\'company-name-display\');\n'
    '  const logoutBtn = document.getElementById(\'btn-logout\');\n'
    '  if (badge) { badge.textContent = acct ? acct.company : \'Mon compte\'; badge.style.display = \'inline-block\'; }\n'
    '  if (logoutBtn) logoutBtn.style.display = \'inline-block\';\n'
    '  // Hide auth overlay\n'
    '  const overlay = document.getElementById(\'auth-overlay\');\n'
    '  if (overlay) overlay.classList.remove(\'active\');\n'
    '  render();\n'
    '}\n'
    '\n'
    'function logoutCompany() {\n'
    '  saveCompanyData();\n'
    '  currentCompanyId = null;\n'
    '  localStorage.removeItem(SESSION_KEY);\n'
    '  // Reset app state to defaults\n'
    '  journalEntries = [...SAMPLE_JOURNAL];\n'
    '  Object.keys(OPENING_BALANCES).forEach(k => delete OPENING_BALANCES[k]);\n'
    '  Object.assign(OPENING_BALANCES, DEFAULT_OPENING_BALANCES);\n'
    '  currentRef = \'syscohada\';\n'
    '  sycebnlType = \'associations\';\n'
    '  const sel = document.getElementById(\'referentiel-select\');\n'
    '  if (sel) sel.value = \'syscohada\';\n'
    '  // Hide topbar elements\n'
    '  const badge = document.getElementById(\'company-name-display\');\n'
    '  const logoutBtn = document.getElementById(\'btn-logout\');\n'
    '  if (badge) badge.style.display = \'none\';\n'
    '  if (logoutBtn) logoutBtn.style.display = \'none\';\n'
    '  // Show auth overlay\n'
    '  showAuthOverlay(\'login\');\n'
    '}\n'
    '\n'
    'function showAuthOverlay(tab) {\n'
    '  const overlay = document.getElementById(\'auth-overlay\');\n'
    '  if (overlay) overlay.classList.add(\'active\');\n'
    '  showAuthTab(tab || \'login\');\n'
    '}\n'
    '\n'
    'function showAuthTab(tab) {\n'
    '  document.getElementById(\'form-login\').style.display = tab === \'login\' ? \'\' : \'none\';\n'
    '  document.getElementById(\'form-register\').style.display = tab === \'register\' ? \'\' : \'none\';\n'
    '  document.getElementById(\'tab-login\').classList.toggle(\'active\', tab === \'login\');\n'
    '  document.getElementById(\'tab-register\').classList.toggle(\'active\', tab === \'register\');\n'
    '}\n'
    '\n'
    'function handleLogin() {\n'
    '  const email = (document.getElementById(\'login-email\').value || \'\').trim().toLowerCase();\n'
    '  const pass  = document.getElementById(\'login-pass\').value || \'\';\n'
    '  const msg   = document.getElementById(\'login-msg\');\n'
    '  msg.className = \'auth-msg\';\n'
    '  if (!email || !pass) { msg.className = \'auth-msg error\'; msg.textContent = \'Veuillez remplir tous les champs.\'; return; }\n'
    '  const accounts = getAccounts();\n'
    '  const acct = accounts.find(a => a.email === email && a.passHash === simpleHash(pass));\n'
    '  if (!acct) { msg.className = \'auth-msg error\'; msg.textContent = \'Email ou mot de passe incorrect.\'; return; }\n'
    '  loginCompany(acct.id);\n'
    '}\n'
    '\n'
    'function handleRegister() {\n'
    '  const company = (document.getElementById(\'reg-company\').value || \'\').trim();\n'
    '  const email   = (document.getElementById(\'reg-email\').value || \'\').trim().toLowerCase();\n'
    '  const pass    = document.getElementById(\'reg-pass\').value || \'\';\n'
    '  const pass2   = document.getElementById(\'reg-pass2\').value || \'\';\n'
    '  const msg     = document.getElementById(\'reg-msg\');\n'
    '  msg.className = \'auth-msg\';\n'
    '  if (!company || !email || !pass) { msg.className = \'auth-msg error\'; msg.textContent = \'Tous les champs sont obligatoires.\'; return; }\n'
    '  if (pass !== pass2) { msg.className = \'auth-msg error\'; msg.textContent = \'Les mots de passe ne correspondent pas.\'; return; }\n'
    '  if (pass.length < 6) { msg.className = \'auth-msg error\'; msg.textContent = \'Le mot de passe doit contenir au moins 6 caracteres.\'; return; }\n'
    '  const accounts = getAccounts();\n'
    '  if (accounts.find(a => a.email === email)) { msg.className = \'auth-msg error\'; msg.textContent = \'Cet email est deja utilise.\'; return; }\n'
    '  const id = \'c_\' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);\n'
    '  accounts.push({ id, company, email, passHash: simpleHash(pass), createdAt: new Date().toISOString() });\n'
    '  saveAccounts(accounts);\n'
    '  msg.className = \'auth-msg success\';\n'
    '  msg.textContent = \'Compte cree ! Connexion en cours...\';\n'
    '  setTimeout(() => loginCompany(id), 600);\n'
    '}\n'
    '\n'
    'function checkAuth() {\n'
    '  const sessionId = localStorage.getItem(SESSION_KEY);\n'
    '  if (sessionId) {\n'
    '    const accounts = getAccounts();\n'
    '    if (accounts.find(a => a.id === sessionId)) {\n'
    '      loginCompany(sessionId);\n'
    '      return;\n'
    '    }\n'
    '  }\n'
    '  // No valid session — show auth overlay\n'
    '  showAuthOverlay(\'login\');\n'
    '}'
)
assert OLD_VARS_END in app, "global vars end not found"
app = app.replace(OLD_VARS_END, NEW_VARS_END, 1)
print("Account management module injected.")

# 2. Add saveCompanyData() at end of render() (before closing brace)
OLD_RENDER_END = '  attachEvents();\n}'
NEW_RENDER_END = '  attachEvents();\n  if (currentCompanyId) saveCompanyData();\n}'
assert OLD_RENDER_END in app, "render() closing not found"
app = app.replace(OLD_RENDER_END, NEW_RENDER_END, 1)
print("saveCompanyData() hooked into render().")

# 3. Replace "// Initial render\nrender();" with checkAuth()
OLD_INIT = '// Initial render\nrender();'
NEW_INIT = '// Initial auth check (shows login or loads company data)\ncheckAuth();'
assert OLD_INIT in app, "Initial render not found"
app = app.replace(OLD_INIT, NEW_INIT, 1)
print("Initial render replaced with checkAuth().")

with open(BASE + r'\app.js', 'w', encoding='utf-8') as f:
    f.write(app)
print(f"app.js written: {len(app)} chars\n")


# ─── data.js: add DEFAULT_OPENING_BALANCES snapshot for logout reset ──────────
with open(BASE + r'\data.js', 'r', encoding='utf-8') as f:
    data = f.read()
print(f"data.js: {len(data)} chars")

# Find end of OPENING_BALANCES object and add DEFAULT_OPENING_BALANCES after it
OLD_OB_END = '  "910":  {n1d:1200000, n1c:0},\n};\n// Journal codes'
NEW_OB_END = (
    '  "910":  {n1d:1200000, n1c:0},\n'
    '};\n'
    '// Snapshot for logout reset (do not modify)\n'
    'const DEFAULT_OPENING_BALANCES = Object.assign({}, OPENING_BALANCES);\n'
    '// Journal codes'
)
assert OLD_OB_END in data, "end of OPENING_BALANCES not found"
data = data.replace(OLD_OB_END, NEW_OB_END, 1)
print("DEFAULT_OPENING_BALANCES snapshot added to data.js.")

with open(BASE + r'\data.js', 'w', encoding='utf-8') as f:
    f.write(data)
print(f"data.js written: {len(data)} chars\n")


print("=" * 50)
print("Multi-company account system implemented:")
print("  - Register: company name + email + password")
print("  - Login: email + password (simple hash)")
print("  - Per-company data: journalEntries + openingBalances + ref in localStorage")
print("  - Auto-save on every render()")
print("  - Company badge + logout button in topbar")
print("  - Auth overlay shown on page load if no session")

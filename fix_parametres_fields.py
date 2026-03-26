#!/usr/bin/env python3
# fix_parametres_fields.py — Update Parametres form to official SYSCOHADA fields

BASE = r'C:\Users\eu\ohada-compta'

with open(BASE + r'\app.js', 'r', encoding='utf-8') as f:
    app = f.read()
print(f"app.js: {len(app)} chars")

OLD_PARAMETRES = '''// ═══════════════════════════════════════════════════════════
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

NEW_PARAMETRES = '''// ═══════════════════════════════════════════════════════════
// PARAMETRES ENTREPRISE
// ═══════════════════════════════════════════════════════════
function renderParametres() {
  const d = currentCompanyDetails;
  const accounts = getAccounts();
  const acct = currentCompanyId ? accounts.find(a => a.id === currentCompanyId) : null;

  const formesJuridiques = ['SA','SARL','SAS','SASU','EURL','SNC','SCS','GIE','Cooperative','Association','ONG','Fondation','Projet de developpement','Etablissement public','Autre'];
  const regimesFiscaux = ['Reel Normal d\'Imposition (RNI)','Reel Simplifie d\'Imposition (RSI)','Contribution des Micro-Entreprises (CME)','Forfait d\'Imposition','Exonere'];

  return `
    <div class="card">
      <div class="card-header">
        <div class="card-title">Identification de l\'entreprise</div>
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
          <div class="form-label">NIF (Numero d\'Identification Fiscale)</div>
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

'''

assert OLD_PARAMETRES in app, "renderParametres() block not found"
app = app.replace(OLD_PARAMETRES, NEW_PARAMETRES, 1)
print("renderParametres() updated with official SYSCOHADA fields.")

print("Dashboard company banner already updated directly.")

with open(BASE + r'\app.js', 'w', encoding='utf-8') as f:
    f.write(app)
print(f"app.js written: {len(app)} chars\n")

print("=" * 50)
print("Parametres form updated to official SYSCOHADA fields:")
print("  Raison sociale, Forme juridique, RCCM, NIF")
print("  Siege social, Activite principale")
print("  Capital social, Regime fiscal")
print("  Pays, Exercice du / au")
print("  Telephone, Email, Expert-comptable, Commissaire aux comptes")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OHADA-COMPTA - Interface Web Streamlit
Application de comptabilite SYSCOHADA/SYCEBNL
"""

import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
from pathlib import Path
import sys

# Ajouter le dossier src au path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

try:
    from database import Database, get_database
    from data.referentiels import PAYS_UEMOA, FORMES_JURIDIQUES, REGIMES_FISCAUX
    from data.plan_comptable import PLAN_COMPTABLE_SYSCOHADA, PLAN_COMPTABLE_SYCEBNL
except ImportError as e:
    st.error(f"Erreur d'importation: {e}")
    st.info("Assurez-vous que les fichiers source sont dans le dossier 'src'")
    st.stop()

# ============================================================================
# CONFIG STREAMLIT
# ============================================================================
st.set_page_config(
    page_title="OHADA-COMPTA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { padding: 0rem 1rem; }
    .metric-card { background-color: #f0f4f8; padding: 20px; border-radius: 10px; margin: 10px 0; }
    .success-box { background-color: #d4edda; padding: 15px; border-radius: 5px; border-left: 4px solid #28a745; }
    .error-box { background-color: #f8d7da; padding: 15px; border-radius: 5px; border-left: 4px solid #dc3545; }
    </style>
    """, unsafe_allow_html=True)

# ============================================================================
# INITIALISATION SESSION
# ============================================================================
if 'db' not in st.session_state:
    st.session_state.db = get_database()

db = st.session_state.db

# ============================================================================
# SIDEBAR - NAVIGATION
# ============================================================================
st.sidebar.title("OHADA-COMPTA")
st.sidebar.write("*Comptabilite SYSCOHADA/SYCEBNL*")
st.sidebar.divider()

app_mode = st.sidebar.radio(
    "Navigation",
    [
        "Accueil",
        "Configuration",
        "Journal",
        "Plan Comptable",
        "Balance Generale",
        "Grand Livre",
        "Etats Financiers",
        "Tableau de Bord",
        "Demo SCI",
        "Outils"
    ]
)

st.sidebar.divider()
entreprise = db.get_entreprise_active()
if entreprise:
    st.sidebar.success(f"Entreprise: {entreprise[1]}")
    st.sidebar.info(f"Ecritures: {db.count_ecritures()}")
else:
    st.sidebar.warning("Aucune entreprise configuree")

# Cross-links to other WASI Ecosystem modules
st.sidebar.divider()
st.sidebar.caption("Ecosysteme WASI")
st.sidebar.markdown("[WASI Intelligence](https://wasi-backend-api.onrender.com/docs)")
st.sidebar.markdown("[AfriTax - Module Fiscal](https://fiscal-liberal-api.onrender.com)")

# ============================================================================
# PAGE: ACCUEIL
# ============================================================================
if app_mode == "Accueil":
    st.title("Bienvenue dans OHADA-COMPTA")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        ### A propos

        OHADA-COMPTA est une application de comptabilite conforme
        aux normes **OHADA** pour les entreprises de l'espace **UEMOA**.

        #### Fonctionnalites
        - Plans comptables pre-charges (SYSCOHADA + SYCEBNL)
        - Saisie des ecritures avec controle d'equilibre
        - Balance generale automatique
        - Bilan et Compte de Resultat
        - Module immobilier SIB-SCI integre
        """)

    with col2:
        st.markdown("""
        ### Pays supportes

        | Pays | Code |
        |------|------|
        | Benin | BJ |
        | Burkina Faso | BF |
        | Cote d'Ivoire | CI |
        | Guinee-Bissau | GW |
        | Mali | ML |
        | Niger | NE |
        | Senegal | SN |
        | Togo | TG |
        """)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Ecritures", db.count_ecritures())
    with col2:
        st.metric("Comptes", len(db.get_plan_comptable()))
    with col3:
        st.metric("Journaux", len(db.get_journaux()))

# ============================================================================
# PAGE: CONFIGURATION
# ============================================================================
elif app_mode == "Configuration":
    st.title("Configuration de l'Entreprise")

    tabs = st.tabs(["Entreprise", "Fiscal", "Contact", "Exercice", "Plan Comptable"])

    with tabs[0]:
        st.subheader("Identification")
        col1, col2 = st.columns(2)
        with col1:
            raison_sociale = st.text_input("Raison sociale *", placeholder="Nom complet")
            sigle = st.text_input("Sigle (optionnel)", placeholder="Acronyme")
        with col2:
            pays_list = [p['nom'] for p in PAYS_UEMOA]
            pays_selected = st.selectbox("Pays *", pays_list)
        col1, col2 = st.columns(2)
        with col1:
            formes_list = sorted(set([f['libelle'] for f in FORMES_JURIDIQUES]))
            forme_selected = st.selectbox("Forme juridique *", formes_list)
        with col2:
            capital = st.number_input("Capital social (FCFA)", min_value=0, step=100000)

        systeme = st.selectbox("Systeme comptable", ["SYSCOHADA", "SYCEBNL"])

        if st.button("Sauvegarder Configuration", use_container_width=True):
            if not raison_sociale:
                st.error("La raison sociale est obligatoire")
            else:
                pays_code = next((p['code'] for p in PAYS_UEMOA if p['nom'] == pays_selected), "BF")
                db.create_entreprise(
                    raison_sociale=raison_sociale,
                    pays=pays_code,
                    forme_juridique=forme_selected,
                    systeme_comptable=systeme,
                    capital=capital,
                    sigle=sigle if sigle else None,
                )
                st.success(f"Entreprise '{raison_sociale}' creee avec succes!")
                st.rerun()

    with tabs[1]:
        st.subheader("Regime Fiscal")
        regimes = [r['libelle'] for r in REGIMES_FISCAUX]
        regime = st.selectbox("Regime fiscal", regimes)
        col1, col2, col3 = st.columns(3)
        with col1:
            rccm = st.text_input("RCCM", placeholder="CI-ABJ-2024-B-12345")
        with col2:
            nif = st.text_input("NIF", placeholder="Numero d'Identification Fiscale")
        with col3:
            cnps = st.text_input("N Securite Sociale", placeholder="CNPS")

        if st.button("Sauvegarder Fiscal", use_container_width=True):
            if entreprise:
                db.update_entreprise(entreprise[0], regime_fiscal=regime, rccm=rccm, nif=nif, cnps=cnps)
                st.success("Informations fiscales sauvegardees!")
                st.rerun()
            else:
                st.warning("Creez d'abord une entreprise dans l'onglet Entreprise")

    with tabs[2]:
        st.subheader("Coordonnees")
        col1, col2 = st.columns(2)
        with col1:
            adresse = st.text_input("Adresse")
            ville = st.text_input("Ville")
        with col2:
            telephone = st.text_input("Telephone", placeholder="+225 XX XX XX XX XX")
            email = st.text_input("Email", placeholder="contact@entreprise.com")

        if st.button("Sauvegarder Contact", use_container_width=True):
            if entreprise:
                db.update_entreprise(entreprise[0], adresse=adresse, ville=ville,
                                     telephone=telephone, email=email)
                st.success("Coordonnees sauvegardees!")
                st.rerun()
            else:
                st.warning("Creez d'abord une entreprise")

    with tabs[3]:
        st.subheader("Exercice Comptable")
        col1, col2 = st.columns(2)
        with col1:
            debut_exercice = st.date_input("Debut exercice", value=date(date.today().year, 1, 1))
        with col2:
            fin_exercice = st.date_input("Fin exercice", value=date(date.today().year, 12, 31))

        if st.button("Sauvegarder Exercice", use_container_width=True):
            if entreprise:
                db.update_entreprise(entreprise[0],
                                     debut_exercice=str(debut_exercice),
                                     fin_exercice=str(fin_exercice))
                st.success("Exercice comptable sauvegarde!")
                st.rerun()
            else:
                st.warning("Creez d'abord une entreprise")

    with tabs[4]:
        st.subheader("Initialiser le Plan Comptable")
        existing = db.get_plan_comptable()
        st.info(f"{len(existing)} comptes actuellement dans la base")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Charger SYSCOHADA", use_container_width=True):
                db.init_plan_comptable(PLAN_COMPTABLE_SYSCOHADA)
                st.success(f"{len(PLAN_COMPTABLE_SYSCOHADA)} comptes SYSCOHADA charges!")
                st.rerun()
        with col2:
            if st.button("Charger SYCEBNL", use_container_width=True):
                db.init_plan_comptable(PLAN_COMPTABLE_SYCEBNL)
                st.success(f"{len(PLAN_COMPTABLE_SYCEBNL)} comptes SYCEBNL charges!")
                st.rerun()

# ============================================================================
# PAGE: JOURNAL
# ============================================================================
elif app_mode == "Journal":
    st.title("Journal Comptable")

    st.subheader("Saisie d'une nouvelle ecriture")

    journaux = db.get_journaux()
    journal_options = [f"{j[1]} - {j[2]}" for j in journaux] if journaux else ["VT - Ventes"]
    comptes = db.get_plan_comptable()
    compte_options = [f"{c[1]} - {c[2]}" for c in comptes] if comptes else ["101 - Capital"]

    col1, col2, col3 = st.columns(3)
    with col1:
        journal_sel = st.selectbox("Journal", journal_options)
    with col2:
        date_ecriture = st.date_input("Date")
    with col3:
        libelle = st.text_input("Libelle de l'ecriture")

    st.markdown("#### Lignes d'ecriture")

    col1, col2, col3 = st.columns(3)
    with col1:
        compte1 = st.selectbox("Compte 1", compte_options, key="c1")
    with col2:
        debit1 = st.number_input("Debit", min_value=0.0, step=1000.0, key="d1")
    with col3:
        credit1 = st.number_input("Credit", min_value=0.0, step=1000.0, key="cr1")

    col1, col2, col3 = st.columns(3)
    with col1:
        compte2 = st.selectbox("Compte 2", compte_options, index=min(1, len(compte_options)-1), key="c2")
    with col2:
        debit2 = st.number_input("Debit", min_value=0.0, step=1000.0, key="d2")
    with col3:
        credit2 = st.number_input("Credit", min_value=0.0, step=1000.0, key="cr2")

    total_debit = debit1 + debit2
    total_credit = credit1 + credit2

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Debit", f"{total_debit:,.0f} FCFA")
    with col2:
        st.metric("Total Credit", f"{total_credit:,.0f} FCFA")
    with col3:
        if total_debit == total_credit and total_debit > 0:
            st.success("Equilibree")
        else:
            st.error("Non equilibree")

    if st.button("Enregistrer l'ecriture", use_container_width=True):
        if total_debit != total_credit or total_debit == 0:
            st.error("L'ecriture doit etre equilibree (debit = credit) et non nulle")
        elif not libelle:
            st.error("Le libelle est obligatoire")
        else:
            journal_code = journal_sel.split(" - ")[0]
            journal_row = next((j for j in journaux if j[1] == journal_code), None)
            if journal_row:
                num_c1 = compte1.split(" - ")[0]
                num_c2 = compte2.split(" - ")[0]
                lignes = [
                    {"numero_compte": num_c1, "libelle_compte": compte1.split(" - ",1)[1] if " - " in compte1 else "", "debit": debit1, "credit": credit1},
                    {"numero_compte": num_c2, "libelle_compte": compte2.split(" - ",1)[1] if " - " in compte2 else "", "debit": debit2, "credit": credit2},
                ]
                lignes = [l for l in lignes if l["debit"] > 0 or l["credit"] > 0]
                try:
                    eid = db.create_ecriture(
                        date_ecriture=str(date_ecriture),
                        journal_id=journal_row[0],
                        libelle=libelle,
                        lignes=lignes,
                    )
                    st.success(f"Ecriture #{eid} enregistree!")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

    st.divider()
    st.subheader("Ecritures recentes")
    ecritures = db.get_ecritures(limit=20)
    if ecritures:
        for ec in ecritures:
            with st.expander(f"{ec[1]} | {ec[4]} (#{ec[0]})"):
                lignes = db.get_lignes_ecriture(ec[0])
                if lignes:
                    df = pd.DataFrame(lignes, columns=["id", "ecriture_id", "compte", "libelle", "debit", "credit", "ref"])
                    st.dataframe(df[["compte", "libelle", "debit", "credit"]], hide_index=True, use_container_width=True)
    else:
        st.info("Aucune ecriture enregistree")

# ============================================================================
# PAGE: PLAN COMPTABLE
# ============================================================================
elif app_mode == "Plan Comptable":
    st.title("Plan Comptable")

    comptes = db.get_plan_comptable()
    if comptes:
        df_plan = pd.DataFrame(comptes, columns=["id", "numero", "libelle", "classe", "type", "sens"])
        col1, col2 = st.columns(2)
        with col1:
            classe = st.multiselect("Filtrer par classe", sorted(df_plan['classe'].unique()))
        with col2:
            search = st.text_input("Rechercher", placeholder="Capital, Banque, Ventes...")
        result = df_plan.copy()
        if classe:
            result = result[result['classe'].isin(classe)]
        if search:
            result = result[
                result['libelle'].str.contains(search, case=False, na=False) |
                result['numero'].astype(str).str.contains(search, na=False)
            ]
        st.dataframe(result[['numero', 'libelle', 'classe', 'type', 'sens']],
                      use_container_width=True, hide_index=True)
        st.info(f"Total: {len(result)} comptes affiches")
    else:
        st.warning("Plan comptable vide. Allez dans Configuration > Plan Comptable pour le charger.")

# ============================================================================
# PAGE: BALANCE GENERALE
# ============================================================================
elif app_mode == "Balance Generale":
    st.title("Balance Generale")

    balance = db.get_balance()
    if balance:
        df = pd.DataFrame(balance, columns=["Numero", "Libelle", "Debit", "Credit", "Solde"])
        st.dataframe(df, use_container_width=True, hide_index=True)

        total_d = df["Debit"].sum()
        total_c = df["Credit"].sum()
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Debit", f"{total_d:,.0f} FCFA")
        with col2:
            st.metric("Total Credit", f"{total_c:,.0f} FCFA")
        with col3:
            if abs(total_d - total_c) < 1:
                st.success("Balance equilibree!")
            else:
                st.warning(f"Ecart: {abs(total_d - total_c):,.0f} FCFA")
    else:
        st.info("Aucune donnee. Saisissez des ecritures dans le Journal.")

# ============================================================================
# PAGE: GRAND LIVRE
# ============================================================================
elif app_mode == "Grand Livre":
    st.title("Grand Livre")

    comptes = db.get_plan_comptable()
    if comptes:
        compte_options = [f"{c[1]} - {c[2]}" for c in comptes]
        compte_sel = st.selectbox("Selectionner un compte", compte_options)
        numero = compte_sel.split(" - ")[0]

        data = db.get_grand_livre(numero)
        if data:
            df = pd.DataFrame(data, columns=["Date", "Libelle", "Debit", "Credit", "Compte"])
            # Compute running balance
            df["Solde"] = (df["Debit"].fillna(0) - df["Credit"].fillna(0)).cumsum()
            st.dataframe(df[["Date", "Libelle", "Debit", "Credit", "Solde"]],
                          use_container_width=True, hide_index=True)
        else:
            st.info("Aucun mouvement sur ce compte")
    else:
        st.warning("Chargez d'abord le plan comptable")

# ============================================================================
# PAGE: ETATS FINANCIERS
# ============================================================================
elif app_mode == "Etats Financiers":
    st.title("Etats Financiers")

    etat_type = st.radio("Selectionner un etat", ["Bilan", "Compte de Resultat"])

    if etat_type == "Bilan":
        st.subheader("Bilan")
        bilan = db.get_bilan()

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### ACTIF")
            if bilan["actif"]:
                total_actif = 0
                for row in bilan["actif"]:
                    st.write(f"**{row[0]}** {row[1]}: {row[2]:,.0f} FCFA")
                    total_actif += row[2] or 0
                st.markdown(f"**TOTAL ACTIF: {total_actif:,.0f} FCFA**")
            else:
                st.info("Pas de donnees")

        with col2:
            st.markdown("### PASSIF")
            if bilan["passif"]:
                total_passif = 0
                for row in bilan["passif"]:
                    st.write(f"**{row[0]}** {row[1]}: {row[2]:,.0f} FCFA")
                    total_passif += row[2] or 0
                st.markdown(f"**TOTAL PASSIF: {total_passif:,.0f} FCFA**")
            else:
                st.info("Pas de donnees")

    else:
        st.subheader("Compte de Resultat")
        cr = db.get_compte_resultat()

        st.markdown("### PRODUITS")
        total_produits = 0
        if cr["produits"]:
            for row in cr["produits"]:
                if row[2] and row[2] > 0:
                    st.write(f"**{row[0]}** {row[1]}: {row[2]:,.0f} FCFA")
                    total_produits += row[2]
        st.markdown(f"**TOTAL PRODUITS: {total_produits:,.0f} FCFA**")

        st.divider()

        st.markdown("### CHARGES")
        total_charges = 0
        if cr["charges"]:
            for row in cr["charges"]:
                if row[2] and row[2] > 0:
                    st.write(f"**{row[0]}** {row[1]}: {row[2]:,.0f} FCFA")
                    total_charges += row[2]
        st.markdown(f"**TOTAL CHARGES: {total_charges:,.0f} FCFA**")

        st.divider()
        resultat = total_produits - total_charges
        if resultat >= 0:
            st.success(f"RESULTAT NET: {resultat:,.0f} FCFA (Benefice)")
        else:
            st.error(f"RESULTAT NET: {resultat:,.0f} FCFA (Deficit)")

# ============================================================================
# PAGE: TABLEAU DE BORD
# ============================================================================
elif app_mode == "Tableau de Bord":
    st.title("Tableau de Bord")

    cr = db.get_compte_resultat()
    total_produits = sum(r[2] or 0 for r in cr["produits"])
    total_charges = sum(r[2] or 0 for r in cr["charges"])
    resultat = total_produits - total_charges

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Produits", f"{total_produits:,.0f} FCFA")
    with col2:
        st.metric("Charges", f"{total_charges:,.0f} FCFA")
    with col3:
        st.metric("Resultat Net", f"{resultat:,.0f} FCFA")
    with col4:
        st.metric("Ecritures", db.count_ecritures())

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Repartition des Charges")
        if cr["charges"]:
            charges_data = {
                'Categorie': [r[1] for r in cr["charges"] if r[2] and r[2] > 0],
                'Montant': [r[2] for r in cr["charges"] if r[2] and r[2] > 0]
            }
            if charges_data['Categorie']:
                import plotly.express as px
                fig = px.pie(charges_data, names='Categorie', values='Montant', hole=0.3)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Pas de charges enregistrees")
        else:
            st.info("Pas de donnees")

    with col2:
        st.markdown("### Balance par classe")
        balance = db.get_balance()
        if balance:
            df = pd.DataFrame(balance, columns=["Numero", "Libelle", "Debit", "Credit", "Solde"])
            df["Classe"] = df["Numero"].str[0].astype(int)
            classe_totals = df.groupby("Classe")["Solde"].sum().reset_index()
            import plotly.express as px
            fig = px.bar(classe_totals, x="Classe", y="Solde")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Pas de donnees")

# ============================================================================
# PAGE: DEMO SCI
# ============================================================================
elif app_mode == "Demo SCI":
    st.title("Demo: Simulation SIB-SCI")
    st.markdown("""
    Cette simulation charge les donnees comptables d'une **Societe Civile Immobiliere**
    (SIB-SCI) avec 132 locaux commerciaux a Abidjan.

    - Capital: 300 000 000 FCFA
    - Loyers mensuels: 25 000 000 FCFA (300M/an)
    - 12 mois d'ecritures + charges + amortissement
    """)

    if st.button("Lancer la simulation SIB-SCI", use_container_width=True, type="primary"):
        try:
            sys.path.insert(0, str(Path(__file__).parent))
            from simulation_sib_sci import init_database, insert_entreprise, insert_journaux
            from simulation_sib_sci import insert_comptes, insert_simulation_data
            from simulation_sib_sci import create_connection

            with st.spinner("Generation des ecritures..."):
                conn, cursor = create_connection()
                init_database(cursor, conn)
                insert_entreprise(cursor, conn)
                insert_journaux(cursor, conn)
                insert_comptes(cursor, conn)
                insert_simulation_data(cursor, conn)
                conn.close()

            st.success("Simulation SIB-SCI chargee! Consultez Balance, Etats Financiers, Grand Livre.")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur: {e}")

    st.divider()
    st.subheader("Donnees actuelles")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Ecritures", db.count_ecritures())
    with col2:
        st.metric("Comptes", len(db.get_plan_comptable()))

# ============================================================================
# PAGE: OUTILS
# ============================================================================
elif app_mode == "Outils":
    st.title("Outils & Utilitaires")

    tab1, tab2, tab3 = st.tabs(["Exports", "Sauvegarde", "Aide"])

    with tab1:
        st.subheader("Exporter les donnees")
        balance = db.get_balance()
        if balance:
            df = pd.DataFrame(balance, columns=["Numero", "Libelle", "Debit", "Credit", "Solde"])
            csv = df.to_csv(index=False)
            st.download_button("Telecharger Balance (CSV)", csv, "balance.csv", "text/csv",
                               use_container_width=True)
        else:
            st.info("Aucune donnee a exporter")

    with tab2:
        st.subheader("Base de donnees")
        st.info(f"Emplacement: {db.db_path}")
        st.info("Sauvegarde automatique a chaque operation.")

    with tab3:
        st.subheader("Aide")
        st.markdown("""
        ### Guide rapide

        1. **Configuration**: Declarez votre entreprise et chargez le plan comptable
        2. **Journal**: Saisissez vos ecritures (debit = credit obligatoire)
        3. **Balance**: Consultez la balance generale
        4. **Etats Financiers**: Bilan et Compte de Resultat
        5. **Demo SCI**: Chargez une simulation complete pour tester

        ### Liens
        - [OHADA](https://www.ohada.org)
        - [Guide SYSCOHADA](https://www.ohada.org)
        """)

# ============================================================================
# FOOTER
# ============================================================================
st.divider()
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9em;'>
    <p>OHADA-COMPTA - Comptabilite SYSCOHADA/SYCEBNL</p>
</div>
""", unsafe_allow_html=True)

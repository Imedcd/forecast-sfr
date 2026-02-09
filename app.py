import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
from pathlib import Path
from pulp import LpProblem, LpVariable, LpMaximize, LpMinimize, lpSum, value, PULP_CBC_CMD, LpStatus, LpInteger
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill


# =============================================================================
# AUTHENTIFICATION
# =============================================================================
def check_password():
    """Retourne True si le mot de passe est correct."""
    def password_entered():
        if st.session_state["password"] == "Batail-Log":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("# 🔐 FORECAST SFR - Authentification")
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        st.info("Entrez le mot de passe pour accéder à l'application")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("# 🔐 FORECAST SFR - Authentification")
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        st.error("❌ Mot de passe incorrect")
        return False
    else:
        return True

if not check_password():
    st.stop()


# =============================================================================
# CONFIG PAGE
# =============================================================================
st.set_page_config(page_title="FORECAST SFR", page_icon="📦", layout="wide")
st.title("📦 FORECAST SFR - Optimisation RO")
st.caption("Allocation optimale de composants telecom | Batail-log")


# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================
def load_bom_sheet(biblio_bytes, sheet_name):
    """Charge une feuille BOM (NOKIA ou huawei)."""
    if sheet_name == "NOKIA":
        df = pd.read_excel(biblio_bytes, sheet_name=sheet_name, header=None, skiprows=1)
        df = df.iloc[:, [2, 3, 5]].copy()
        df.columns = ["Conf|version", "Référence", "Quantité"]
    elif sheet_name == "huawei":
        df = pd.read_excel(biblio_bytes, sheet_name=sheet_name, header=0)
        df = df[['CONF', 'Qty', 'Radical text']].copy()
        df.columns = ["Conf|version", "Quantité", "Référence"]
    else:
        raise ValueError(f"Sheet inconnu : {sheet_name}")

    df["Référence"] = (
        df["Référence"].astype(str).str.strip().str.upper()
        .str.replace(r'_R$', '', regex=True)
        .str.replace(r'\s+', ' ', regex=True)
    )
    df["Quantité"] = pd.to_numeric(df["Quantité"], errors="coerce").fillna(0)
    df = df[df["Quantité"] > 0].copy()

    part_before = df["Conf|version"].str.split("|").str[0].str.strip()
    part_after = df["Conf|version"].str.extract(r'\|(.+)$', expand=False).fillna("")
    df["Conf"] = part_after
    df["Version"] = part_before
    return df[["Conf|version", "Conf", "Version", "Référence", "Quantité"]]


def prepare_common_data(biblio_file, prix_file, stock_file, acc_file):
    """Prépare les données communes (BOM + Stock + Prix)."""

    # --- BOM ---
    bom_nokia = load_bom_sheet(BytesIO(biblio_file.getvalue()), "NOKIA")
    bom_huawei = load_bom_sheet(BytesIO(biblio_file.getvalue()), "huawei")
    bom_df = pd.concat([bom_nokia, bom_huawei], ignore_index=True)

    # Normalisation BOM via acc
    acc_bom = pd.read_excel(BytesIO(acc_file.getvalue()), sheet_name="Feuil1")
    acc_bom.columns = acc_bom.columns.str.strip().str.lower().str.replace(r'[\s,]+', '', regex=True)
    code_col = next((c for c in acc_bom.columns if 'code' in c), None)
    ref_col = next((c for c in acc_bom.columns if any(k in c for k in ['ref', 'reference', 'article'])), None)

    if code_col and ref_col:
        acc_bom[code_col] = acc_bom[code_col].astype(str).str.strip().str.upper()
        acc_bom[ref_col] = acc_bom[ref_col].astype(str).str.strip().str.upper()
        mapping = acc_bom[[code_col, ref_col]].drop_duplicates(subset=code_col, keep='first')
        code_to_ref = dict(zip(mapping[code_col], mapping[ref_col]))
        for idx, ref in bom_df["Référence"].items():
            if ref in code_to_ref:
                bom_df.at[idx, "Référence"] = code_to_ref[ref]

    # --- Stock ---
    df_stock = pd.read_excel(stock_file, sheet_name="Stock", usecols=["Code", "Stock Dispo", "Prévisionnel"])
    df_stock = df_stock.rename(columns={"Code": "Référence"})
    df_stock["Stock"] = (
        pd.to_numeric(df_stock["Stock Dispo"], errors='coerce').fillna(0) +
        pd.to_numeric(df_stock["Prévisionnel"], errors='coerce').fillna(0)
    )
    df_stock = df_stock[["Référence", "Stock"]].copy()
    df_stock["Référence"] = df_stock["Référence"].astype(str).str.strip().str.upper()

    # Conversion stock via acc
    acc = pd.read_excel(BytesIO(acc_file.getvalue()), sheet_name="Feuil1")
    acc.columns = acc.columns.str.strip().str.lower().str.replace(r'[\s,]+', '', regex=True)

    column_map = {}
    for col in acc.columns:
        if 'code' in col: column_map[col] = 'code'
        elif any(k in col for k in ['ref', 'reference', 'article']): column_map[col] = 'ref'
        elif any(k in col for k in ['bom', 'multi', 'coeff']): column_map[col] = 'bom'
    acc = acc.rename(columns=column_map)

    acc["code"] = acc["code"].astype(str).str.strip().str.upper()
    acc["ref"] = acc["ref"].astype(str).str.strip().str.upper()
    acc["bom"] = pd.to_numeric(acc["bom"], errors='coerce').fillna(1)

    acc_mapping = acc[["code", "ref", "bom"]].drop_duplicates(subset="code", keep="first")
    code_to_ref_dict = dict(zip(acc_mapping["code"], acc_mapping["ref"]))
    code_to_bom_dict = dict(zip(acc_mapping["code"], acc_mapping["bom"]))

    df_stock["Referentiel"] = df_stock["Référence"].map(code_to_ref_dict).fillna(df_stock["Référence"])
    df_stock["multiplicateur"] = df_stock["Référence"].map(code_to_bom_dict).fillna(1)

    ref_to_bom_fallback = acc.groupby("ref")["bom"].first().to_dict()
    mask = (df_stock["multiplicateur"] == 1) & (df_stock["Referentiel"] != df_stock["Référence"])
    df_stock.loc[mask, "multiplicateur"] = df_stock.loc[mask, "Referentiel"].map(ref_to_bom_fallback).fillna(1)

    df_stock["Valeur_NVX"] = df_stock["Stock"] * df_stock["multiplicateur"]

    stock_df = df_stock.groupby("Referentiel", as_index=False).agg({"Valeur_NVX": "sum", "Stock": "sum"})
    stock_df = stock_df.rename(columns={"Valeur_NVX": "NVX STOCK", "Stock": "Stock Physique"})

    # --- Prix ---
    df_prix_ref = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="References", usecols=["Référence", "Prix (pj)"])
    df_prix_ref["Référence"] = df_prix_ref["Référence"].astype(str).str.strip().str.upper()
    df_prix_ref["Prix (pj)"] = pd.to_numeric(df_prix_ref["Prix (pj)"], errors='coerce').fillna(0)

    prix_unique = df_prix_ref.groupby("Référence", as_index=False)["Prix (pj)"].first()
    ref_to_prix_dict = dict(zip(prix_unique["Référence"], prix_unique["Prix (pj)"]))
    stock_df["Prix (pj)"] = stock_df["Referentiel"].map(ref_to_prix_dict).fillna(0)
    stock_df = stock_df.rename(columns={"Stock Physique": "Stock"})

    return bom_df, stock_df, ref_to_prix_dict


def generate_excel_bytes(wb):
    """Convertit un workbook openpyxl en bytes."""
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# =============================================================================
# ONGLETS
# =============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📁 Upload Fichiers",
    "🎯 Jalon 1 - Sans Priorité",
    "⭐ Jalon 1 - Avec Priorité",
    "🚀 Jalon 2 - Complet (avec achats)"
])


# =============================================================================
# TAB 1 : UPLOAD FICHIERS
# =============================================================================
with tab1:
    st.header("📁 Upload des fichiers sources")
    st.markdown("""
    Uploadez les fichiers Excel nécessaires pour les optimisations.
    Ces fichiers seront utilisés dans les onglets suivants.
    """)

    col1, col2 = st.columns(2)
    with col1:
        biblio_file = st.file_uploader("📚 biblio.xlsx (BOM Nokia + Huawei)", type="xlsx", key="biblio")
        prix_file = st.file_uploader("💰 Prix.xlsx", type="xlsx", key="prix")
        stock_file = st.file_uploader("📦 Stock.xlsx", type="xlsx", key="stock")
    with col2:
        acc_file = st.file_uploader("🔄 acc.xlsx (Conversion)", type="xlsx", key="acc")
        forecast_file = st.file_uploader("📊 Forecast.xlsx (pour Jalon 1 sans prio)", type="xlsx", key="forecast")
        prio_file = st.file_uploader("⭐ Prio.xlsx (pour Jalon 1 avec prio + Jalon 2)", type="xlsx", key="prio")

    files_common = all([biblio_file, prix_file, stock_file, acc_file])

    if files_common:
        st.success("✅ Fichiers communs chargés (BOM, Prix, Stock, Conversion)")
        if st.button("🔄 Préparer les données communes", type="primary"):
            with st.spinner("Préparation en cours..."):
                try:
                    bom_df, stock_df, ref_to_prix = prepare_common_data(biblio_file, prix_file, stock_file, acc_file)
                    st.session_state['bom_df'] = bom_df
                    st.session_state['stock_df'] = stock_df
                    st.session_state['ref_to_prix'] = ref_to_prix
                    st.session_state['data_prepared'] = True
                    st.success("✅ Données préparées avec succès !")

                    c1, c2 = st.columns(2)
                    c1.metric("Lignes BOM", len(bom_df))
                    c2.metric("Références Stock", len(stock_df))
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")

    if forecast_file:
        st.info("✅ Forecast.xlsx chargé → Jalon 1 sans priorité disponible")
    if prio_file:
        st.info("✅ Prio.xlsx chargé → Jalon 1 avec priorité + Jalon 2 disponibles")


# =============================================================================
# TAB 2 : JALON 1 SANS PRIORITE
# =============================================================================
with tab2:
    st.header("🎯 Jalon 1 - Optimisation SANS Priorité")
    st.markdown("""
    **Objectif** : Maximiser le nombre total de configurations produites avec le stock disponible.

    **Méthode** :
    1. Maximiser le nombre de configs
    2. Maximiser le coût du stock consommé (à nombre fixe)
    """)

    if not st.session_state.get('data_prepared'):
        st.warning("⚠️ Préparez d'abord les données dans l'onglet 'Upload Fichiers'")
    elif not forecast_file:
        st.warning("⚠️ Uploadez le fichier Forecast.xlsx dans l'onglet 'Upload Fichiers'")
    else:
        if st.button("▶️ Lancer Jalon 1 Sans Priorité", key="run_j1_sans", type="primary"):
            with st.spinner("Optimisation en cours..."):
                try:
                    # Charger la demande
                    df_forecast = pd.read_excel(forecast_file, sheet_name="Feuil1",
                        usecols=["Constructeur", "Conf", "Version", "Conf|version", "Demande"])
                    df_forecast["Demande"] = pd.to_numeric(df_forecast["Demande"], errors="coerce").fillna(0)
                    df_forecast = df_forecast[df_forecast["Demande"] > 0].copy()

                    df_prix_conf = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="Conf", usecols=["Conf|version", "Prix"])
                    df_prix_conf["Prix"] = pd.to_numeric(df_prix_conf["Prix"], errors="coerce")
                    demand_df = df_forecast.merge(df_prix_conf, on="Conf|version", how="left")

                    # Récupérer BOM et Stock
                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()

                    # Optimisation
                    stock_optim = stock_df.set_index("Referentiel")
                    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)

                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)
                    configs = [cv for cv in demand_df['Conf|version'].unique() if cv in q.index]
                    demand_map = dict(zip(demand_df['Conf|version'], demand_df['Demande']))

                    # Etape 1
                    prob1 = LpProblem("Max_Nb_Configs", LpMaximize)
                    X1 = {}
                    for cv in configs:
                        dem_max = demand_map.get(cv, 0)
                        if dem_max > 0:
                            X1[cv] = LpVariable(f"X1_{cv.replace('|','_').replace(' ','_').replace('-','_')}", 0, dem_max, LpInteger)

                    prob1 += lpSum(X1.values())
                    for ref in q.columns:
                        coeffs = [(q.at[cv, ref], X1[cv]) for cv in X1 if q.at[cv, ref] > 0]
                        if coeffs:
                            lim = stock_optim.at[ref, 'NVX STOCK'] if ref in stock_optim.index else 0
                            prob1 += lpSum(coef * var for coef, var in coeffs) <= lim

                    prob1.solve(PULP_CBC_CMD(msg=0, timeLimit=180, gapRel=0.0))
                    max_configs = value(prob1.objective) or 0

                    # Etape 2
                    prob2 = LpProblem("Max_Cout_Stock", LpMaximize)
                    X = {}
                    for cv in configs:
                        dem_max = demand_map.get(cv, 0)
                        if dem_max > 0:
                            X[cv] = LpVariable(f"X_{cv.replace('|','_').replace(' ','_').replace('-','_')}", 0, dem_max, LpInteger)

                    cout_stock = lpSum(
                        X[cv] * sum(q.at[cv, ref] * stock_optim.at[ref, 'Prix (pj)']
                                    for ref in q.columns if ref in stock_optim.index and q.at[cv, ref] > 0)
                        for cv in X
                    )
                    prob2 += cout_stock
                    prob2 += lpSum(X.values()) == max_configs

                    for ref in q.columns:
                        coeffs = [(q.at[cv, ref], X[cv]) for cv in X if q.at[cv, ref] > 0]
                        if coeffs:
                            lim = stock_optim.at[ref, 'NVX STOCK'] if ref in stock_optim.index else 0
                            prob2 += lpSum(coef * var for coef, var in coeffs) <= lim

                    prob2.solve(PULP_CBC_CMD(msg=0, timeLimit=300, gapRel=0.01))

                    produced = {cv: max(0, int(round(value(X[cv].varValue) or 0))) for cv in X}
                    consumed = {}
                    for ref in q.columns:
                        qty = sum(q.at[cv, ref] * produced.get(cv, 0) for cv in configs if cv in produced)
                        if qty > 0:
                            consumed[ref] = round(qty)

                    cout_total = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed.items() if ref in stock_optim.index)

                    # Résultats
                    demand_df['Qté produite'] = demand_df['Conf|version'].map(produced).fillna(0).astype(int)
                    demand_df['Restant'] = demand_df['Demande'] - demand_df['Qté produite']

                    total_demande = demand_df['Demande'].sum()
                    total_produit = demand_df['Qté produite'].sum()
                    pct = (total_produit / total_demande * 100) if total_demande > 0 else 0

                    st.success("✅ Optimisation terminée !")

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale", f"{int(total_demande)}")
                    c2.metric("Produit", f"{int(total_produit)}")
                    c3.metric("Taux", f"{pct:.1f}%")
                    c4.metric("Coût stock", f"{cout_total:,.0f} €")

                    st.subheader("📊 Détail par configuration")
                    detail_df = demand_df[['Constructeur', 'Conf|version', 'Demande', 'Prix', 'Qté produite', 'Restant']].copy()
                    detail_df['% Réalisé'] = (detail_df['Qté produite'] / detail_df['Demande'] * 100).round(1)
                    st.dataframe(detail_df, use_container_width=True)

                    st.session_state['j1_sans_result'] = (demand_df, bom_df, stock_df, produced, consumed, cout_total)

                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())


# =============================================================================
# TAB 3 : JALON 1 AVEC PRIORITE
# =============================================================================
with tab3:
    st.header("⭐ Jalon 1 - Optimisation AVEC Priorité")
    st.markdown("""
    **Objectif** : Maximiser les configs en respectant les priorités (séquentiel).

    **Méthode** :
    - Traiter priorité par priorité
    - Consommer le stock progressivement
    """)

    if not st.session_state.get('data_prepared'):
        st.warning("⚠️ Préparez d'abord les données dans l'onglet 'Upload Fichiers'")
    elif not prio_file:
        st.warning("⚠️ Uploadez le fichier Prio.xlsx dans l'onglet 'Upload Fichiers'")
    else:
        if st.button("▶️ Lancer Jalon 1 Avec Priorité", key="run_j1_avec", type="primary"):
            with st.spinner("Optimisation en cours..."):
                try:
                    # Charger Prio
                    df_prio = pd.read_excel(prio_file, sheet_name="Feuil1", engine="openpyxl")
                    cols_utiles = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Priorité', 'Demande']
                    df_prio = df_prio[[c for c in cols_utiles if c in df_prio.columns]].copy()
                    df_prio['Demande'] = pd.to_numeric(df_prio['Demande'], errors='coerce').fillna(0).astype(int)
                    df_prio['Priorité'] = pd.to_numeric(df_prio['Priorité'], errors='coerce').fillna(999).astype(int)
                    df_prio = df_prio[df_prio['Demande'] > 0].copy()

                    df_prix_conf = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="Conf", usecols=["Conf|version", "Prix"])
                    df_prix_conf["Prix"] = pd.to_numeric(df_prix_conf["Prix"], errors="coerce")
                    all_demands = df_prio.merge(df_prix_conf, on="Conf|version", how="left")

                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()

                    # Optimisation par priorité
                    stock_optim = stock_df.set_index("Referentiel")
                    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)

                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)
                    all_demands['Demande'] = all_demands['Demande'].clip(lower=0).astype(int)

                    stock_current = stock_optim['NVX STOCK'].copy()
                    produced_j1 = {}
                    priorities = sorted(all_demands['Priorité'].unique())

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for phase, prio_level in enumerate(priorities, 1):
                        status_text.text(f"Traitement priorité {prio_level}...")
                        progress_bar.progress(phase / len(priorities))

                        phase_rows = all_demands[all_demands['Priorité'] == prio_level]
                        if phase_rows.empty:
                            continue

                        phase_configs = [cv for cv in phase_rows['Conf|version'].unique() if cv in q.index]
                        if not phase_configs:
                            continue

                        demand_map = {}
                        for _, row in phase_rows.iterrows():
                            cv = row['Conf|version']
                            if cv in phase_configs:
                                demand_map[cv] = demand_map.get(cv, 0) + row['Demande']

                        # Etape 1
                        prob1 = LpProblem(f"Prio_{prio_level}_MaxQty", LpMaximize)
                        X1 = {}
                        for cv in phase_configs:
                            dem_max = demand_map.get(cv, 0)
                            if dem_max > 0:
                                X1[cv] = LpVariable(f"X1_{cv.replace('|','_').replace(' ','_')}", 0, dem_max, LpInteger)

                        if not X1:
                            continue

                        prob1 += lpSum(X1.values())
                        for ref in q.columns:
                            coeffs = [(q.at[cv, ref], X1[cv]) for cv in X1 if q.at[cv, ref] > 0]
                            if coeffs:
                                lim = stock_current[ref] if ref in stock_current.index else 0
                                prob1 += lpSum(coef * var for coef, var in coeffs) <= lim

                        prob1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
                        max_phase = value(prob1.objective) or 0

                        if max_phase == 0:
                            for cv in X1:
                                produced_j1[(cv, prio_level)] = 0
                            continue

                        # Etape 2
                        prob2 = LpProblem(f"Prio_{prio_level}_MaxCost", LpMaximize)
                        X = {}
                        for cv in phase_configs:
                            dem_max = demand_map.get(cv, 0)
                            if dem_max > 0:
                                X[cv] = LpVariable(f"X_{cv.replace('|','_').replace(' ','_')}", 0, dem_max, LpInteger)

                        cout_phase = lpSum(
                            X[cv] * sum(q.at[cv, ref] * stock_optim.at[ref, 'Prix (pj)']
                                        for ref in q.columns if ref in stock_optim.index and q.at[cv, ref] > 0)
                            for cv in X
                        )
                        prob2 += cout_phase
                        prob2 += lpSum(X.values()) == max_phase

                        for ref in q.columns:
                            coeffs = [(q.at[cv, ref], X[cv]) for cv in X if q.at[cv, ref] > 0]
                            if coeffs:
                                lim = stock_current[ref] if ref in stock_current.index else 0
                                prob2 += lpSum(coef * var for coef, var in coeffs) <= lim

                        prob2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))

                        for cv in X:
                            qty = max(0, int(round(value(X[cv].varValue) or 0)))
                            produced_j1[(cv, prio_level)] = qty
                            for ref in q.columns:
                                cons = q.at[cv, ref] * qty
                                if cons > 0 and ref in stock_current.index:
                                    stock_current[ref] = max(0, stock_current[ref] - cons)

                    progress_bar.progress(1.0)
                    status_text.text("✅ Optimisation terminée !")

                    # Résultats
                    all_demands['J1 produit'] = all_demands.apply(
                        lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                    )
                    all_demands['Restant'] = all_demands['Demande'] - all_demands['J1 produit']

                    total_demande = all_demands['Demande'].sum()
                    total_j1 = all_demands['J1 produit'].sum()
                    pct_j1 = (total_j1 / total_demande * 100) if total_demande > 0 else 0

                    consumed_j1 = {}
                    for (cv, prio), qty in produced_j1.items():
                        if cv in q.index and qty > 0:
                            for ref in q.columns:
                                cons = q.at[cv, ref] * qty
                                if cons > 0:
                                    consumed_j1[ref] = consumed_j1.get(ref, 0) + cons

                    cout_j1 = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed_j1.items() if ref in stock_optim.index)

                    st.success("✅ Optimisation terminée !")

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale", f"{int(total_demande)}")
                    c2.metric("Produit", f"{int(total_j1)}")
                    c3.metric("Taux", f"{pct_j1:.1f}%")
                    c4.metric("Coût stock", f"{cout_j1:,.0f} €")

                    st.subheader("📊 Par priorité")
                    prio_summary = []
                    for prio in sorted(all_demands['Priorité'].unique()):
                        prio_data = all_demands[all_demands['Priorité'] == prio]
                        dem_p = int(prio_data['Demande'].sum())
                        prod_p = int(prio_data['J1 produit'].sum())
                        pct_p = (prod_p / dem_p * 100) if dem_p > 0 else 0
                        prio_summary.append({
                            'Priorité': prio,
                            'Demande': dem_p,
                            'Produit': prod_p,
                            'Taux %': f"{pct_p:.1f}%"
                        })
                    st.dataframe(pd.DataFrame(prio_summary), use_container_width=True)

                    st.session_state['j1_avec_result'] = (all_demands, bom_df, stock_df, produced_j1, consumed_j1, cout_j1)

                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())


# =============================================================================
# TAB 4 : JALON 2 COMPLET
# =============================================================================
with tab4:
    st.header("🚀 Jalon 2 - Optimisation Complète (Jalon 1 + Achats)")
    st.markdown("""
    **Objectif** : Après le Jalon 1, acheter des références pour produire plus de configs.

    **Méthode** :
    1. Jalon 1 : Maximiser avec le stock disponible
    2. Jalon 2 : Acheter des références (seuils 20%, 40%, 60%, 80% du prix config)
    """)

    if not st.session_state.get('data_prepared'):
        st.warning("⚠️ Préparez d'abord les données dans l'onglet 'Upload Fichiers'")
    elif not prio_file:
        st.warning("⚠️ Uploadez le fichier Prio.xlsx dans l'onglet 'Upload Fichiers'")
    else:
        st.info("⚠️ **Attention** : L'exécution complète du Jalon 2 peut prendre 5-10 minutes.")

        if st.button("▶️ Lancer Jalon 2 Complet", key="run_j2", type="primary"):
            st.warning("🚧 **Jalon 2 en cours de développement** - Fonctionnalité disponible prochainement.")
            st.markdown("""
            Le Jalon 2 complet nécessite :
            - Optimisation séquentielle par priorité (Jalon 1)
            - Puis optimisation d'achat de références par seuils (20%, 40%, 60%, 80%)
            - Génération d'un Excel détaillé avec toutes les sections

            Pour l'instant, utilisez le script Python `RO_Jalon2.py` directement.
            """)


# =============================================================================
# FOOTER
# =============================================================================
st.divider()
st.caption("FORECAST SFR - Batail-log | Python + PuLP (CBC Solver) | Version 1.0")

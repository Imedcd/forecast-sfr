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

SEUILS = [0.20, 0.40, 0.60, 0.80]

# =============================================================================
# CONFIG PAGE
# =============================================================================
st.set_page_config(
    page_title="FORECAST SFR",
    page_icon="■",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =============================================================================
# CSS GLOBAL
# =============================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #0A0D14; }
    #MainMenu, footer, header {visibility: hidden;}
    .block-container { padding: 2rem 3rem; max-width: 1400px; }

    .main-header {
        display: flex; align-items: center; justify-content: space-between;
        padding: 0 0 1.5rem 0; border-bottom: 1px solid #1E2A3A; margin-bottom: 2rem;
    }
    .main-title { font-size: 1.6rem; font-weight: 700; color: #FFFFFF; margin: 0; }
    .main-title span { color: #00D9FF; }
    .main-badge {
        font-size: 0.7rem; background: rgba(0,217,255,0.1); color: #00D9FF;
        border: 1px solid rgba(0,217,255,0.3); padding: 0.3rem 0.8rem;
        border-radius: 100px; font-weight: 500; letter-spacing: 1px; text-transform: uppercase;
    }
    .stTabs [data-baseweb="tab-list"] {
        background: #131720; border-radius: 10px; padding: 0.4rem;
        gap: 4px; border: 1px solid #1E2A3A;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent; border-radius: 8px; color: #6B7689;
        font-weight: 500; font-size: 0.9rem; padding: 0.6rem 1.4rem; border: none;
    }
    .stTabs [aria-selected="true"] { background: #00D9FF !important; color: #0A0D14 !important; font-weight: 600; }
    .info-block {
        background: #131720; border: 1px solid #1E2A3A; border-left: 3px solid #00D9FF;
        border-radius: 10px; padding: 1.25rem 1.5rem; margin: 1rem 0 1.5rem 0;
        font-size: 0.9rem; color: #B0B8C4; line-height: 1.8;
    }
    .warning-block {
        background: #131720; border: 1px solid #1E2A3A; border-left: 3px solid #FFB300;
        border-radius: 10px; padding: 1.25rem 1.5rem; margin: 1rem 0 1.5rem 0;
        font-size: 0.9rem; color: #B0B8C4;
    }
    .stButton > button[kind="primary"] {
        background: #00D9FF; color: #0A0D14; font-weight: 600; border: none;
        border-radius: 8px; padding: 0.65rem 2rem; font-size: 0.9rem; transition: all 0.2s;
    }
    .stButton > button[kind="primary"]:hover {
        background: #33E3FF; box-shadow: 0 0 20px rgba(0,217,255,0.35); transform: translateY(-1px);
    }
    .stButton > button:not([kind="primary"]) {
        background: #131720; color: #B0B8C4; border: 1px solid #1E2A3A; border-radius: 8px;
    }
    [data-testid="metric-container"] {
        background: #131720; border: 1px solid #1E2A3A; border-radius: 10px; padding: 1rem 1.25rem;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.8rem; color: #6B7689; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px;
    }
    [data-testid="stMetricValue"] { font-size: 1.8rem; font-weight: 700; color: #00D9FF; }
    [data-testid="stFileUploader"] {
        background: #131720; border: 1.5px dashed #1E2A3A; border-radius: 10px; padding: 0.5rem;
    }
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; border: 1px solid #1E2A3A; }
    hr { border-color: #1E2A3A; }
    .footer-text { font-size: 0.78rem; color: #3A4455; text-align: center; padding-top: 1rem; }
    .section-title { font-size: 1.4rem; font-weight: 600; color: #FFFFFF; margin: 0 0 0.5rem 0; }
    .section-sub { font-size: 0.85rem; color: #6B7689; margin-bottom: 1.5rem; }
    .label-cat { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #6B7689; margin-bottom: 0.75rem; }
    .kpi-label { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #6B7689; margin: 1.5rem 0 0.75rem 0; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# AUTHENTIFICATION
# =============================================================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "Batail-Log":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    wrong = st.session_state.get("password_correct") == False
    if "password_correct" not in st.session_state or wrong:
        _, center, _ = st.columns([1.2, 1, 1.2])
        with center:
            st.markdown("""
            <div style="text-align:center; padding: 4rem 0 2rem 0;">
                <div style="font-size:3rem; font-weight:700; color:#FFFFFF; letter-spacing:-2px;">
                    FORECAST<span style="color:#00D9FF;">SFR</span>
                </div>
                <div style="font-size:0.8rem; color:#6B7689; text-transform:uppercase; letter-spacing:3px; margin-top:0.5rem;">
                    Optimisation RO &nbsp;·&nbsp; Batail-log
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.text_input("Mot de passe", type="password", on_change=password_entered, key="password", placeholder="••••••••••")
            if wrong:
                st.markdown("""
                <div style="background:rgba(255,82,82,0.08); border:1px solid rgba(255,82,82,0.25);
                            border-radius:8px; padding:0.7rem 1rem; color:#FF5252;
                            font-size:0.83rem; text-align:center; margin-top:0.5rem;">
                    Mot de passe incorrect
                </div>
                """, unsafe_allow_html=True)
        return False
    return True

if not check_password():
    st.stop()


# =============================================================================
# HEADER
# =============================================================================
st.markdown("""
<div class="main-header">
    <div>
        <div class="main-title">FORECAST<span>SFR</span> &nbsp;—&nbsp; Optimisation RO</div>
        <div style="font-size:0.82rem; color:#6B7689; margin-top:0.3rem;">Allocation optimale de composants telecom</div>
    </div>
    <div class="main-badge">Batail-log v1.0</div>
</div>
""", unsafe_allow_html=True)


# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================
def load_bom_sheet(biblio_bytes, sheet_name):
    if sheet_name == "NOKIA":
        df = pd.read_excel(biblio_bytes, sheet_name=sheet_name, header=None, skiprows=1)
        df = df.iloc[:, [2, 3, 5]].copy()
        df.columns = ["Conf|version", "Reference", "Quantite"]
    elif sheet_name == "huawei":
        df = pd.read_excel(biblio_bytes, sheet_name=sheet_name, header=0)
        df = df[['CONF', 'Qty', 'Radical text']].copy()
        df.columns = ["Conf|version", "Quantite", "Reference"]
    else:
        raise ValueError(f"Sheet inconnu : {sheet_name}")
    df["Reference"] = (df["Reference"].astype(str).str.strip().str.upper()
                       .str.replace(r'_R$', '', regex=True).str.replace(r'\s+', ' ', regex=True))
    df["Quantite"] = pd.to_numeric(df["Quantite"], errors="coerce").fillna(0)
    df = df[df["Quantite"] > 0].copy()
    part_before = df["Conf|version"].str.split("|").str[0].str.strip()
    part_after = df["Conf|version"].str.extract(r'\|(.+)$', expand=False).fillna("")
    df["Conf"] = part_after
    df["Version"] = part_before
    return df[["Conf|version", "Conf", "Version", "Reference", "Quantite"]]


def prepare_common_data(biblio_file, prix_file, stock_file, acc_file):
    bom_nokia = load_bom_sheet(BytesIO(biblio_file.getvalue()), "NOKIA")
    bom_huawei = load_bom_sheet(BytesIO(biblio_file.getvalue()), "huawei")
    bom_df = pd.concat([bom_nokia, bom_huawei], ignore_index=True)
    bom_df = bom_df.rename(columns={"Reference": "Référence", "Quantite": "Quantité"})

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

    df_stock = pd.read_excel(stock_file, sheet_name="Stock", usecols=["Code", "Stock Dispo", "Prévisionnel"])
    df_stock = df_stock.rename(columns={"Code": "Référence"})
    df_stock["Stock"] = (pd.to_numeric(df_stock["Stock Dispo"], errors='coerce').fillna(0) +
                         pd.to_numeric(df_stock["Prévisionnel"], errors='coerce').fillna(0))
    df_stock = df_stock[["Référence", "Stock"]].copy()
    df_stock["Référence"] = df_stock["Référence"].astype(str).str.strip().str.upper()

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

    df_prix_ref = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="References", usecols=["Référence", "Prix (pj)"])
    df_prix_ref["Référence"] = df_prix_ref["Référence"].astype(str).str.strip().str.upper()
    df_prix_ref["Prix (pj)"] = pd.to_numeric(df_prix_ref["Prix (pj)"], errors='coerce').fillna(0)
    prix_unique = df_prix_ref.groupby("Référence", as_index=False)["Prix (pj)"].first()
    ref_to_prix_dict = dict(zip(prix_unique["Référence"], prix_unique["Prix (pj)"]))
    stock_df["Prix (pj)"] = stock_df["Referentiel"].map(ref_to_prix_dict).fillna(0)
    stock_df = stock_df.rename(columns={"Stock Physique": "Stock"})
    return bom_df, stock_df, ref_to_prix_dict


def sname(s):
    return re.sub(r'[^a-zA-Z0-9_]', '_', str(s))


def lp_val(v):
    x = value(v)
    return max(0, x) if x is not None else 0


def run_jalon1(all_demands, bom_df, stock_df, q, priorities, prix_file_bytes):
    stock_optim = stock_df.set_index("Referentiel")
    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)
    stock_current = stock_optim['NVX STOCK'].copy()
    produced_j1 = {}

    progress_bar = st.progress(0)
    status_text = st.empty()

    for phase, prio_level in enumerate(priorities, 1):
        status_text.text(f"Jalon 1 — Priorité {prio_level} / {len(priorities)}...")
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
    status_text.text("Jalon 1 termine")

    consumed_j1 = {}
    for (cv, prio), qty in produced_j1.items():
        if cv in q.index and qty > 0:
            for ref in q.columns:
                cons = q.at[cv, ref] * qty
                if cons > 0:
                    consumed_j1[ref] = consumed_j1.get(ref, 0) + cons
    cout_j1 = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed_j1.items() if ref in stock_optim.index)
    return produced_j1, consumed_j1, cout_j1, stock_current.copy(), stock_optim


# =============================================================================
# ONGLETS
# =============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "  Upload Fichiers  ",
    "  Jalon 1 — Sans Priorité  ",
    "  Jalon 1 — Avec Priorité  ",
    "  Jalon 2 — Complet  "
])


# =============================================================================
# TAB 1 : UPLOAD
# =============================================================================
with tab1:
    st.markdown('<div class="section-title">Upload des fichiers sources</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Chargez les fichiers Excel avant de lancer une optimisation.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown('<div class="label-cat">Fichiers obligatoires</div>', unsafe_allow_html=True)
        biblio_file = st.file_uploader("biblio.xlsx — BOM Nokia + Huawei", type="xlsx", key="biblio")
        prix_file = st.file_uploader("Prix.xlsx", type="xlsx", key="prix")
        stock_file = st.file_uploader("Stock.xlsx", type="xlsx", key="stock")
    with col2:
        st.markdown('<div class="label-cat">Fichiers de configuration</div>', unsafe_allow_html=True)
        acc_file = st.file_uploader("acc.xlsx — Table de conversion", type="xlsx", key="acc")
        forecast_file = st.file_uploader("Forecast.xlsx — Jalon 1 sans priorité", type="xlsx", key="forecast")
        prio_file = st.file_uploader("Prio.xlsx — Jalon 1 avec priorité + Jalon 2", type="xlsx", key="prio")

    st.markdown("<br>", unsafe_allow_html=True)
    files_common = all([biblio_file, prix_file, stock_file, acc_file])
    if files_common:
        st.success("Fichiers communs chargés")
        if st.button("Préparer les données communes", type="primary"):
            with st.spinner("Chargement et normalisation..."):
                try:
                    bom_df, stock_df, ref_to_prix = prepare_common_data(biblio_file, prix_file, stock_file, acc_file)
                    st.session_state['bom_df'] = bom_df
                    st.session_state['stock_df'] = stock_df
                    st.session_state['ref_to_prix'] = ref_to_prix
                    st.session_state['data_prepared'] = True
                    st.success("Données prêtes")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Lignes BOM", len(bom_df))
                    c2.metric("Références stock", len(stock_df))
                    val_stock = (stock_df['NVX STOCK'] * stock_df['Prix (pj)']).sum()
                    c3.metric("Valeur stock total", f"{val_stock:,.0f} €")
                except Exception as e:
                    st.error(f"Erreur : {e}")
    if forecast_file:
        st.info("Forecast.xlsx chargé — Jalon 1 sans priorité disponible")
    if prio_file:
        st.info("Prio.xlsx chargé — Jalon 1 avec priorité et Jalon 2 disponibles")


# =============================================================================
# TAB 2 : JALON 1 SANS PRIORITE
# =============================================================================
with tab2:
    st.markdown('<div class="section-title">Jalon 1 — Sans Priorité</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Maximise le nombre total de configurations avec le stock disponible.</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode en 2 étapes</strong><br>
        1. Maximiser le nombre de configurations réalisables<br>
        2. A nombre fixé, maximiser la valeur du stock consommé
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Upload Fichiers")
    elif not forecast_file:
        st.warning("Chargez le fichier Forecast.xlsx dans l'onglet Upload Fichiers")
    else:
        if st.button("Lancer l'optimisation", key="run_j1_sans", type="primary"):
            with st.spinner("Résolution en cours..."):
                try:
                    xl_fc = pd.ExcelFile(BytesIO(forecast_file.getvalue()), engine="openpyxl")
                    sheet_fc = xl_fc.sheet_names[0]
                    df_forecast = pd.read_excel(BytesIO(forecast_file.getvalue()), sheet_name=sheet_fc,
                        usecols=["Constructeur", "Conf", "Version", "Conf|version", "Demande"])
                    df_forecast["Demande"] = pd.to_numeric(df_forecast["Demande"], errors="coerce").fillna(0)
                    df_forecast = df_forecast[df_forecast["Demande"] > 0].copy()

                    df_prix_conf = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="Conf", usecols=["Conf|version", "Prix"])
                    df_prix_conf["Prix"] = pd.to_numeric(df_prix_conf["Prix"], errors="coerce")
                    demand_df = df_forecast.merge(df_prix_conf, on="Conf|version", how="left")

                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()
                    stock_optim = stock_df.set_index("Referentiel")
                    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)

                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)
                    configs = [cv for cv in demand_df['Conf|version'].unique() if cv in q.index]
                    demand_map = dict(zip(demand_df['Conf|version'], demand_df['Demande']))

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

                    demand_df['Qte produite'] = demand_df['Conf|version'].map(produced).fillna(0).astype(int)
                    demand_df['Restant'] = demand_df['Demande'] - demand_df['Qte produite']
                    total_demande = demand_df['Demande'].sum()
                    total_produit = demand_df['Qte produite'].sum()
                    pct = (total_produit / total_demande * 100) if total_demande > 0 else 0
                    nb_rupture = (demand_df['Restant'] > 0).sum()
                    val_manquante = (demand_df['Restant'] * demand_df['Prix'].fillna(0)).sum()
                    val_stock_init = (stock_optim['NVX STOCK'] * stock_optim['Prix (pj)']).sum()
                    val_stock_restant = val_stock_init - cout_total

                    st.success("Optimisation terminée")
                    st.markdown("<br>", unsafe_allow_html=True)

                    st.markdown('<div class="kpi-label">Résultats de production</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale", f"{int(total_demande)}")
                    c2.metric("Configurations produites", f"{int(total_produit)}")
                    c3.metric("Taux de réalisation", f"{pct:.1f}%")
                    c4.metric("Valeur stock consommé", f"{cout_total:,.0f} €")

                    st.markdown("<br>", unsafe_allow_html=True)
                    col_a, col_b = st.columns([3, 2])
                    with col_a:
                        st.markdown('<div class="kpi-label">Détail par configuration</div>', unsafe_allow_html=True)
                        detail_df = demand_df[['Constructeur', 'Conf|version', 'Demande', 'Prix', 'Qte produite', 'Restant']].copy()
                        detail_df['% Realise'] = (detail_df['Qte produite'] / detail_df['Demande'] * 100).round(1)
                        st.dataframe(detail_df, use_container_width=True, height=350)
                    with col_b:
                        st.markdown('<div class="kpi-label">Top 10 références consommées</div>', unsafe_allow_html=True)
                        if consumed:
                            top_refs = sorted(consumed.items(),
                                              key=lambda x: x[1] * (stock_optim.at[x[0], 'Prix (pj)'] if x[0] in stock_optim.index else 0),
                                              reverse=True)[:10]
                            top_df = pd.DataFrame([{
                                'Référence': r,
                                'Quantité': int(q_),
                                'Valeur (€)': f"{round(q_ * (stock_optim.at[r, 'Prix (pj)'] if r in stock_optim.index else 0)):,}"
                            } for r, q_ in top_refs])
                            st.dataframe(top_df, use_container_width=True, height=350)

                    if 'Constructeur' in demand_df.columns:
                        st.markdown('<div class="kpi-label">Répartition par constructeur</div>', unsafe_allow_html=True)
                        constr = demand_df.groupby('Constructeur').agg(
                            Demande=('Demande', 'sum'), Produit=('Qte produite', 'sum')
                        ).reset_index()
                        constr['Restant'] = constr['Demande'] - constr['Produit']
                        constr['Taux %'] = (constr['Produit'] / constr['Demande'] * 100).round(1).astype(str) + '%'
                        st.dataframe(constr, use_container_width=True)

                    # Export Excel - 4 onglets identiques au script standalone
                    _BLEU     = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    _BLEU_C   = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
                    _VERT     = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    _ORANGE   = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    _ROUGE    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    _GRIS     = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    _BLEU_F   = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
                    _BB       = Font(bold=True, color="FFFFFF", size=11)
                    _BB14     = Font(bold=True, color="FFFFFF", size=14)
                    _BN       = Font(bold=True, size=11)
                    _EUR      = '#,##0 €'
                    _NUM      = '#,##0'

                    def _fmt_h(ws, c=None):
                        c = c or _BLEU
                        for cell in ws[1]:
                            cell.font = _BB; cell.fill = c
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 30
                        from openpyxl.utils import get_column_letter
                        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

                    def _adj_w(ws):
                        for col in ws.columns:
                            ml = max((len(str(c.value or "")) for c in col), default=8)
                            ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 45)

                    def _ligne_tot(ws, row_num, nb_cols):
                        for ci in range(1, nb_cols+1):
                            ws.cell(row=row_num, column=ci).font = _BN
                            ws.cell(row=row_num, column=ci).fill = _GRIS

                    wb_sp = Workbook()

                    # ---- Onglet 1 : Demandes ----
                    ws_dem = wb_sp.active; ws_dem.title = "Demandes"
                    cols_exp = [c for c in ['Constructeur','Conf','Version','Conf|version','Demande','Prix','Qte produite','Restant'] if c in demand_df.columns]
                    df_exp = demand_df[cols_exp].sort_values('Constructeur') if 'Constructeur' in demand_df.columns else demand_df[cols_exp]
                    for rr in dataframe_to_rows(df_exp, index=False, header=True): ws_dem.append(rr)
                    tot_row = ws_dem.max_row + 1
                    ws_dem.cell(tot_row, 1, "TOTAL")
                    cidx = {cols_exp[i]: i+1 for i in range(len(cols_exp))}
                    for cn in ['Demande','Qte produite','Restant']:
                        if cn in cidx: ws_dem.cell(tot_row, cidx[cn], df_exp[cn].sum())
                    _ligne_tot(ws_dem, tot_row, len(cols_exp))
                    _fmt_h(ws_dem); _adj_w(ws_dem)
                    for row in ws_dem.iter_rows(min_row=2, max_row=ws_dem.max_row):
                        for cell in row:
                            h = ws_dem.cell(1, cell.column).value
                            if h == 'Prix' and isinstance(cell.value, (int, float)): cell.number_format = _EUR
                            elif h in ['Demande','Qte produite','Restant'] and isinstance(cell.value, (int, float)): cell.number_format = _NUM
                    if 'Restant' in cidx:
                        for rn in range(2, ws_dem.max_row):
                            c = ws_dem.cell(rn, cidx['Restant'])
                            if isinstance(c.value, (int, float)):
                                c.fill = _VERT if c.value == 0 else _ORANGE if c.value > 0 else c.fill

                    # ---- Onglet 2 : BOM ----
                    ws_bom = wb_sp.create_sheet("BOM")
                    bom_exp = bom_df[['Conf|version','Référence','Quantité']].copy()
                    bom_exp['Stock initial'] = bom_exp['Référence'].map(lambda r: stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0)
                    bom_exp['Stock a consommer'] = bom_exp.apply(lambda row: produced.get(row['Conf|version'], 0) * row['Quantité'], axis=1)
                    bom_exp['Stock final'] = bom_exp['Référence'].map(lambda r: (stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0) - consumed.get(r, 0))
                    bom_exp['Bloquant'] = bom_exp.apply(lambda row: 'Oui' if row['Stock final'] < row['Quantité'] else 'Non', axis=1)
                    bom_exp = bom_exp.sort_values(['Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_exp, index=False, header=True): ws_bom.append(rr)
                    _fmt_h(ws_bom); _adj_w(ws_bom)
                    col_bl = list(bom_exp.columns).index('Bloquant') + 1
                    col_sf = list(bom_exp.columns).index('Stock final') + 1
                    for rn in range(2, ws_bom.max_row + 1):
                        cb = ws_bom.cell(rn, col_bl); csf = ws_bom.cell(rn, col_sf)
                        if cb.value == 'Oui': cb.fill = _ROUGE; csf.fill = _ROUGE
                        else: cb.fill = _VERT

                    # ---- Onglet 3 : Stock ----
                    ws_stk = wb_sp.create_sheet("Stock")
                    stk_out = stock_df.copy().rename(columns={'NVX STOCK':'Stock initial','Prix (pj)':'Prix unitaire'})
                    stk_out['Consomme'] = stk_out['Referentiel'].map(consumed).fillna(0).round(2)
                    stk_out['Stock final'] = (stk_out['Stock initial'] - stk_out['Consomme']).round(2)
                    cols_stk = ['Referentiel','Stock initial','Prix unitaire','Consomme','Stock final']
                    stk_out = stk_out[cols_stk].sort_values('Referentiel')
                    for rr in dataframe_to_rows(stk_out, index=False, header=True): ws_stk.append(rr)
                    _fmt_h(ws_stk); _adj_w(ws_stk)
                    for row in ws_stk.iter_rows(min_row=2, max_row=ws_stk.max_row):
                        for cell in row:
                            if ws_stk.cell(1, cell.column).value == 'Prix unitaire' and isinstance(cell.value, (int, float)): cell.number_format = _EUR
                    col_sf2 = cols_stk.index('Stock final') + 1
                    for rn in range(2, ws_stk.max_row + 1):
                        c = ws_stk.cell(rn, col_sf2)
                        if isinstance(c.value, (int, float)):
                            if c.value < 0: c.fill = _ROUGE
                            elif c.value == 0: c.fill = _ORANGE

                    # ---- Onglet 4 : Résultats ----
                    ws_res = wb_sp.create_sheet("Résultats")
                    rr = 1
                    ws_res.cell(rr, 1, "RÉCAPITULATIF GLOBAL").font = _BB14
                    for ci in range(1, 3): ws_res.cell(rr, ci).fill = _BLEU_F
                    rr = 3
                    cout_stk_init = float((stock_optim['NVX STOCK'] * stock_optim['Prix (pj)']).sum())
                    cout_stk_fin = cout_stk_init - cout_total
                    for lbl, val in [
                        ("Total demande", int(total_demande)),
                        ("Total produit", int(total_produit)),
                        ("% Réalisation", f"{pct:.1f}%"),
                        ("Restant", int(total_demande - total_produit)),
                        ("Stock initial (EUR)", round(cout_stk_init, 2)),
                        ("Stock consommé (EUR)", round(cout_total, 2)),
                        ("Stock final (EUR)", round(cout_stk_fin, 2)),
                    ]:
                        ws_res.cell(rr, 1, lbl); ws_res.cell(rr, 2, val)
                        if isinstance(val, (int, float)) and "EUR" in lbl: ws_res.cell(rr, 2).number_format = _EUR
                        elif isinstance(val, int): ws_res.cell(rr, 2).number_format = _NUM
                        rr += 1
                    rr += 1
                    ws_res.cell(rr, 1, "DÉTAIL PAR CONF").font = _BB14
                    for ci in range(1, 6): ws_res.cell(rr, ci).fill = _BLEU_F
                    rr += 1
                    for i, h in enumerate(["Constructeur","Conf|version","Demande","Qte produite","Réalisation"], 1):
                        ws_res.cell(rr, i, h).font = _BN; ws_res.cell(rr, i).fill = _BLEU_C
                        ws_res.cell(rr, i).alignment = Alignment(horizontal="center")
                    rr += 1
                    recap = demand_df.groupby(['Constructeur','Conf|version'] if 'Constructeur' in demand_df.columns else ['Conf|version']).agg({'Demande':'sum','Qte produite':'sum'}).reset_index()
                    for _, row_data in recap.iterrows():
                        ws_res.cell(rr, 1, row_data.get('Constructeur',''))
                        ws_res.cell(rr, 2, row_data['Conf|version'])
                        ws_res.cell(rr, 3, int(row_data['Demande'])).number_format = _NUM
                        ws_res.cell(rr, 4, int(row_data['Qte produite'])).number_format = _NUM
                        real_pct = row_data['Qte produite'] / row_data['Demande'] * 100 if row_data['Demande'] > 0 else 0
                        ws_res.cell(rr, 5, f"{real_pct:.1f}%").alignment = Alignment(horizontal="center")
                        if row_data['Qte produite'] == row_data['Demande']: ws_res.cell(rr, 5).fill = _VERT
                        elif row_data['Qte produite'] == 0: ws_res.cell(rr, 5).fill = _ROUGE
                        else: ws_res.cell(rr, 5).fill = _ORANGE
                        rr += 1
                    for cl, w in [('A',22),('B',22),('C',14),('D',16),('E',14)]: ws_res.column_dimensions[cl].width = w

                    out_sp = BytesIO(); wb_sp.save(out_sp)
                    ts_sp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state['j1_sp_excel'] = out_sp.getvalue()
                    st.session_state['j1_sp_excel_name'] = f"J1_SansPrio_{ts_sp}.xlsx"

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        if 'j1_sp_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="Télécharger les résultats Jalon 1 Sans Priorité",
                data=st.session_state['j1_sp_excel'],
                file_name=st.session_state['j1_sp_excel_name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_j1_sp"
            )


# =============================================================================
# TAB 3 : JALON 1 AVEC PRIORITE
# =============================================================================
with tab3:
    st.markdown('<div class="section-title">Jalon 1 — Avec Priorité</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Optimisation séquentielle respectant l\'ordre des priorités.</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode séquentielle</strong><br>
        Chaque niveau de priorité est traité indépendamment.<br>
        Le stock est consommé progressivement de la priorité la plus haute à la plus basse.
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Upload Fichiers")
    elif not prio_file:
        st.warning("Chargez le fichier Prio.xlsx dans l'onglet Upload Fichiers")
    else:
        if st.button("Lancer l'optimisation", key="run_j1_avec", type="primary"):
            with st.spinner("Résolution en cours..."):
                try:
                    xl_prio = pd.ExcelFile(BytesIO(prio_file.getvalue()), engine="openpyxl")
                    sheet_prio = xl_prio.sheet_names[0]
                    df_prio = pd.read_excel(BytesIO(prio_file.getvalue()), sheet_name=sheet_prio, engine="openpyxl")
                    df_prio.columns = df_prio.columns.str.strip()
                    # Détection flexible de la colonne Demande
                    dem_col = next((c for c in df_prio.columns if c.lower() in ['demande', 'qte', 'quantite', 'quantité', 'qty', 'volume']), None)
                    if dem_col is None:
                        raise ValueError(f"Colonne 'Demande' introuvable. Colonnes disponibles : {list(df_prio.columns)}")
                    if dem_col != 'Demande':
                        df_prio = df_prio.rename(columns={dem_col: 'Demande'})
                    # Détection flexible de la colonne Priorité
                    prio_col = next((c for c in df_prio.columns if c.lower() in ['priorité', 'priorite', 'prio', 'priority']), None)
                    if prio_col and prio_col != 'Priorité':
                        df_prio = df_prio.rename(columns={prio_col: 'Priorité'})
                    if 'Priorité' not in df_prio.columns:
                        df_prio['Priorité'] = 1
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
                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)
                    all_demands['Demande'] = all_demands['Demande'].clip(lower=0).astype(int)
                    priorities = sorted(all_demands['Priorité'].unique())

                    produced_j1, consumed_j1, cout_j1, stock_after_j1, stock_optim = run_jalon1(
                        all_demands, bom_df, stock_df, q, priorities, prix_file.getvalue()
                    )

                    all_demands['J1 produit'] = all_demands.apply(
                        lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                    )
                    all_demands['Restant J1'] = all_demands['Demande'] - all_demands['J1 produit']
                    total_demande = all_demands['Demande'].sum()
                    total_j1 = all_demands['J1 produit'].sum()
                    pct_j1 = (total_j1 / total_demande * 100) if total_demande > 0 else 0
                    nb_en_rupture = (all_demands['Restant J1'] > 0).sum()
                    val_rupture = (all_demands['Restant J1'] * all_demands['Prix'].fillna(0)).sum()

                    st.success("Optimisation terminée")
                    st.markdown("<br>", unsafe_allow_html=True)

                    st.markdown('<div class="kpi-label">Résultats de production</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale", f"{int(total_demande)}")
                    c2.metric("Configurations produites", f"{int(total_j1)}")
                    c3.metric("Taux de réalisation", f"{pct_j1:.1f}%")
                    c4.metric("Valeur stock consommé", f"{cout_j1:,.0f} €")

                    st.markdown("<br>", unsafe_allow_html=True)
                    if 'Constructeur' in all_demands.columns:
                        st.markdown('<div class="kpi-label">Détail par constructeur</div>', unsafe_allow_html=True)
                        constr_detail = all_demands[['Constructeur', 'Conf|version', 'Priorité', 'Demande', 'J1 produit', 'Restant J1']].copy()
                        constr_detail = constr_detail.rename(columns={'J1 produit': 'J1 Produit', 'Restant J1': 'Restant'})
                        constr_detail['Taux %'] = (constr_detail['J1 Produit'] / constr_detail['Demande'].replace(0, 1) * 100).round(1).astype(str) + '%'
                        constr_detail = constr_detail.sort_values(['Constructeur', 'Priorité'])
                        st.dataframe(constr_detail, use_container_width=True)

                    # Export Excel - 4 onglets identiques au script standalone
                    _BLEU2   = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    _BLEU_C2 = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
                    _VERT2   = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    _ORANGE2 = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    _ROUGE2  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    _GRIS2   = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    _BLEU_F2 = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
                    _BB2     = Font(bold=True, color="FFFFFF", size=11)
                    _BB142   = Font(bold=True, color="FFFFFF", size=14)
                    _BN2     = Font(bold=True, size=11)
                    _EUR2    = '#,##0 €'
                    _NUM2    = '#,##0'

                    def _fmt_h2(ws, c=None):
                        c = c or _BLEU2
                        for cell in ws[1]:
                            cell.font = _BB2; cell.fill = c
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 30
                        from openpyxl.utils import get_column_letter
                        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

                    def _adj_w2(ws):
                        for col in ws.columns:
                            ml = max((len(str(c.value or "")) for c in col), default=8)
                            ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 45)

                    def _ligne_tot2(ws, row_num, nb_cols):
                        for ci in range(1, nb_cols+1):
                            ws.cell(row=row_num, column=ci).font = _BN2
                            ws.cell(row=row_num, column=ci).fill = _GRIS2

                    wb_ap = Workbook()

                    # Préparer un df export avec colonnes renommées
                    df_ap_exp = all_demands.rename(columns={'J1 produit': 'Qté produite', 'Restant J1': 'Restant'})

                    # ---- Onglet 1 : Demandes ----
                    ws_dem2 = wb_ap.active; ws_dem2.title = "Demandes"
                    cols_exp2 = [c for c in ['Constructeur','Conf','Version','Conf|version','Priorité','Demande','Prix','Qté produite','Restant'] if c in df_ap_exp.columns]
                    df_exp2 = df_ap_exp[cols_exp2].sort_values(['Constructeur','Priorité'] if 'Constructeur' in df_ap_exp.columns else ['Priorité'])
                    for rr in dataframe_to_rows(df_exp2, index=False, header=True): ws_dem2.append(rr)
                    tot_row2 = ws_dem2.max_row + 1
                    ws_dem2.cell(tot_row2, 1, "TOTAL")
                    cidx2 = {cols_exp2[i]: i+1 for i in range(len(cols_exp2))}
                    for cn in ['Demande','Qté produite','Restant']:
                        if cn in cidx2: ws_dem2.cell(tot_row2, cidx2[cn], df_exp2[cn].sum())
                    _ligne_tot2(ws_dem2, tot_row2, len(cols_exp2))
                    _fmt_h2(ws_dem2); _adj_w2(ws_dem2)
                    for row in ws_dem2.iter_rows(min_row=2, max_row=ws_dem2.max_row):
                        for cell in row:
                            h = ws_dem2.cell(1, cell.column).value
                            if h == 'Prix' and isinstance(cell.value, (int, float)): cell.number_format = _EUR2
                            elif h in ['Demande','Qté produite','Restant'] and isinstance(cell.value, (int, float)): cell.number_format = _NUM2
                    if 'Restant' in cidx2:
                        for rn in range(2, ws_dem2.max_row):
                            c = ws_dem2.cell(rn, cidx2['Restant'])
                            if isinstance(c.value, (int, float)):
                                c.fill = _VERT2 if c.value == 0 else _ORANGE2 if c.value > 0 else c.fill

                    # ---- Onglet 2 : BOM ----
                    ws_bom2 = wb_ap.create_sheet("BOM")
                    produced_total_j1 = {}
                    for (cv, p_), qty in produced_j1.items():
                        produced_total_j1[cv] = produced_total_j1.get(cv, 0) + qty
                    bom_exp2 = bom_df[['Conf|version','Référence','Quantité']].copy()
                    bom_exp2['Stock initial'] = bom_exp2['Référence'].map(lambda r: stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0)
                    bom_exp2['Stock a consommer'] = bom_exp2.apply(lambda row: produced_total_j1.get(row['Conf|version'], 0) * row['Quantité'], axis=1)
                    bom_exp2['Stock final'] = bom_exp2['Référence'].map(lambda r: (stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0) - consumed_j1.get(r, 0))
                    bom_exp2['Bloquant'] = bom_exp2.apply(lambda row: 'Oui' if row['Stock final'] < row['Quantité'] else 'Non', axis=1)
                    bom_exp2 = bom_exp2.sort_values(['Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_exp2, index=False, header=True): ws_bom2.append(rr)
                    _fmt_h2(ws_bom2); _adj_w2(ws_bom2)
                    col_bl2 = list(bom_exp2.columns).index('Bloquant') + 1
                    col_sf2b = list(bom_exp2.columns).index('Stock final') + 1
                    for rn in range(2, ws_bom2.max_row + 1):
                        cb = ws_bom2.cell(rn, col_bl2); csf = ws_bom2.cell(rn, col_sf2b)
                        if cb.value == 'Oui': cb.fill = _ROUGE2; csf.fill = _ROUGE2
                        else: cb.fill = _VERT2

                    # ---- Onglet 3 : Stock ----
                    ws_stk2 = wb_ap.create_sheet("Stock")
                    stk_out2 = stock_df.copy().rename(columns={'NVX STOCK':'Stock initial','Prix (pj)':'Prix unitaire'})
                    stk_out2['Consomme'] = stk_out2['Referentiel'].map(consumed_j1).fillna(0).round(2)
                    stk_out2['Stock final'] = (stk_out2['Stock initial'] - stk_out2['Consomme']).round(2)
                    cols_stk2 = ['Referentiel','Stock initial','Prix unitaire','Consomme','Stock final']
                    stk_out2 = stk_out2[cols_stk2].sort_values('Referentiel')
                    for rr in dataframe_to_rows(stk_out2, index=False, header=True): ws_stk2.append(rr)
                    _fmt_h2(ws_stk2); _adj_w2(ws_stk2)
                    for row in ws_stk2.iter_rows(min_row=2, max_row=ws_stk2.max_row):
                        for cell in row:
                            if ws_stk2.cell(1, cell.column).value == 'Prix unitaire' and isinstance(cell.value, (int, float)): cell.number_format = _EUR2
                    col_sf2c = cols_stk2.index('Stock final') + 1
                    for rn in range(2, ws_stk2.max_row + 1):
                        c = ws_stk2.cell(rn, col_sf2c)
                        if isinstance(c.value, (int, float)):
                            if c.value < 0: c.fill = _ROUGE2
                            elif c.value == 0: c.fill = _ORANGE2

                    # ---- Onglet 4 : Résultats ----
                    ws_res2 = wb_ap.create_sheet("Résultats")
                    rr2 = 1
                    ws_res2.cell(rr2, 1, "RÉCAPITULATIF GLOBAL").font = _BB142
                    for ci in range(1, 3): ws_res2.cell(rr2, ci).fill = _BLEU_F2
                    rr2 = 3
                    cout_stk_init2 = float((stock_optim['NVX STOCK'] * stock_optim['Prix (pj)']).sum())
                    cout_stk_fin2 = cout_stk_init2 - cout_j1
                    for lbl, val in [
                        ("Total demande", int(total_demande)),
                        ("Total produit", int(total_j1)),
                        ("% Réalisation", f"{pct_j1:.1f}%"),
                        ("Restant", int(total_demande - total_j1)),
                        ("Stock initial (EUR)", round(cout_stk_init2, 2)),
                        ("Stock consommé (EUR)", round(cout_j1, 2)),
                        ("Stock final (EUR)", round(cout_stk_fin2, 2)),
                    ]:
                        ws_res2.cell(rr2, 1, lbl); ws_res2.cell(rr2, 2, val)
                        if isinstance(val, (int, float)) and "EUR" in lbl: ws_res2.cell(rr2, 2).number_format = _EUR2
                        elif isinstance(val, int): ws_res2.cell(rr2, 2).number_format = _NUM2
                        rr2 += 1
                    rr2 += 1
                    ws_res2.cell(rr2, 1, "DÉTAIL PAR CONSTRUCTEUR / CONF").font = _BB142
                    for ci in range(1, 6): ws_res2.cell(rr2, ci).fill = _BLEU_F2
                    rr2 += 1
                    for i, h in enumerate(["Constructeur","Conf","Demande","Qté produite","Réalisation"], 1):
                        ws_res2.cell(rr2, i, h).font = _BN2; ws_res2.cell(rr2, i).fill = _BLEU_C2
                        ws_res2.cell(rr2, i).alignment = Alignment(horizontal="center")
                    rr2 += 1
                    recap2 = df_ap_exp.groupby(['Constructeur','Conf'] if 'Constructeur' in df_ap_exp.columns else ['Conf']).agg({'Demande':'sum','Qté produite':'sum'}).reset_index()
                    recap2 = recap2.sort_values(['Constructeur','Conf'] if 'Constructeur' in recap2.columns else ['Conf'])
                    for _, row_data in recap2.iterrows():
                        ws_res2.cell(rr2, 1, row_data.get('Constructeur',''))
                        ws_res2.cell(rr2, 2, row_data.get('Conf',''))
                        ws_res2.cell(rr2, 3, int(row_data['Demande'])).number_format = _NUM2
                        ws_res2.cell(rr2, 4, int(row_data['Qté produite'])).number_format = _NUM2
                        real_pct2 = row_data['Qté produite'] / row_data['Demande'] * 100 if row_data['Demande'] > 0 else 0
                        ws_res2.cell(rr2, 5, f"{real_pct2:.1f}%").alignment = Alignment(horizontal="center")
                        if row_data['Qté produite'] == row_data['Demande']: ws_res2.cell(rr2, 5).fill = _VERT2
                        elif row_data['Qté produite'] == 0: ws_res2.cell(rr2, 5).fill = _ROUGE2
                        else: ws_res2.cell(rr2, 5).fill = _ORANGE2
                        rr2 += 1
                    for cl, w in [('A',18),('B',14),('C',14),('D',16),('E',14)]: ws_res2.column_dimensions[cl].width = w

                    out_ap = BytesIO(); wb_ap.save(out_ap)
                    ts_ap = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state['j1_ap_excel'] = out_ap.getvalue()
                    st.session_state['j1_ap_excel_name'] = f"J1_AvecPrio_{ts_ap}.xlsx"

                    st.session_state['j1_avec_result'] = {
                        'all_demands': all_demands,
                        'stock_optim': stock_optim,
                        'produced_j1': produced_j1,
                        'consumed_j1': consumed_j1,
                        'cout_j1': cout_j1,
                        'q': q,
                        'priorities': priorities,
                        'stock_after_j1': stock_after_j1,
                    }

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        if 'j1_ap_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="Télécharger les résultats Jalon 1 Avec Priorité",
                data=st.session_state['j1_ap_excel'],
                file_name=st.session_state['j1_ap_excel_name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_j1_ap"
            )


# =============================================================================
# TAB 4 : JALON 2 COMPLET
# =============================================================================
with tab4:
    st.markdown('<div class="section-title">Jalon 2 — Optimisation Complète</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Jalon 1 + achats de références pour maximiser la production.</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode en 2 phases</strong><br>
        1. Jalon 1 : maximisation avec le stock disponible (par priorité)<br>
        2. Jalon 2 : achat de références manquantes par seuils (20% → 40% → 60% → 80% du prix config)
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Upload Fichiers")
    elif not prio_file:
        st.warning("Chargez le fichier Prio.xlsx dans l'onglet Upload Fichiers")
    else:
        has_j1 = 'j1_avec_result' in st.session_state
        if has_j1:
            st.success("Résultats Jalon 1 disponibles — le Jalon 2 repartira de ces résultats")
        else:
            st.markdown("""<div class="warning-block">
                Jalon 1 non calculé — il sera recalculé automatiquement au lancement du Jalon 2.
            </div>""", unsafe_allow_html=True)

        if st.button("Lancer Jalon 2 complet", key="run_j2", type="primary"):
            with st.spinner("Optimisation Jalon 2 en cours..."):
                try:
                    xl_prio = pd.ExcelFile(BytesIO(prio_file.getvalue()), engine="openpyxl")
                    sheet_prio = xl_prio.sheet_names[0]
                    df_prio = pd.read_excel(BytesIO(prio_file.getvalue()), sheet_name=sheet_prio, engine="openpyxl")
                    df_prio.columns = df_prio.columns.str.strip()
                    dem_col = next((c for c in df_prio.columns if c.lower() in ['demande', 'qte', 'quantite', 'quantité', 'qty', 'volume']), None)
                    if dem_col is None:
                        raise ValueError(f"Colonne 'Demande' introuvable. Colonnes disponibles : {list(df_prio.columns)}")
                    if dem_col != 'Demande':
                        df_prio = df_prio.rename(columns={dem_col: 'Demande'})
                    prio_col = next((c for c in df_prio.columns if c.lower() in ['priorité', 'priorite', 'prio', 'priority']), None)
                    if prio_col and prio_col != 'Priorité':
                        df_prio = df_prio.rename(columns={prio_col: 'Priorité'})
                    if 'Priorité' not in df_prio.columns:
                        df_prio['Priorité'] = 1
                    cols_utiles = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Priorité', 'Demande']
                    df_prio = df_prio[[c for c in cols_utiles if c in df_prio.columns]].copy()
                    df_prio['Demande'] = pd.to_numeric(df_prio['Demande'], errors='coerce').fillna(0).astype(int)
                    df_prio['Priorité'] = pd.to_numeric(df_prio['Priorité'], errors='coerce').fillna(999).astype(int)
                    df_prio = df_prio[df_prio['Demande'] > 0].copy()

                    df_prix_conf = pd.read_excel(BytesIO(prix_file.getvalue()), sheet_name="Conf", usecols=["Conf|version", "Prix"])
                    df_prix_conf["Prix"] = pd.to_numeric(df_prix_conf["Prix"], errors="coerce")
                    all_demands = df_prio.merge(df_prix_conf, on="Conf|version", how="left")
                    all_demands['Demande'] = all_demands['Demande'].clip(lower=0).astype(int)

                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()
                    ref_to_prix = st.session_state['ref_to_prix']
                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)
                    priorities = sorted(all_demands['Priorité'].unique())

                    # ---- JALON 1 ----
                    if has_j1:
                        j1_data = st.session_state['j1_avec_result']
                        produced_j1 = j1_data['produced_j1']
                        consumed_j1 = j1_data['consumed_j1']
                        cout_j1 = j1_data['cout_j1']
                        stock_j2 = j1_data['stock_after_j1'].copy()
                        stock_optim = j1_data['stock_optim']
                        all_demands['J1 produit'] = all_demands.apply(
                            lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )
                        st.info("Jalon 1 repris depuis les résultats existants")
                    else:
                        produced_j1, consumed_j1, cout_j1, stock_j2, stock_optim = run_jalon1(
                            all_demands, bom_df, stock_df, q, priorities, prix_file.getvalue()
                        )
                        all_demands['J1 produit'] = all_demands.apply(
                            lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )

                    all_demands['Restant J1'] = all_demands['Demande'] - all_demands['J1 produit']
                    total_demande = all_demands['Demande'].sum()
                    total_j1 = all_demands['J1 produit'].sum()
                    pct_j1 = (total_j1 / total_demande * 100) if total_demande > 0 else 0

                    # ---- JALON 2 ----
                    bom_par_config = {cv: {ref: q.at[cv, ref] for ref in q.columns if q.at[cv, ref] > 0} for cv in q.index}
                    prix_configs = {}
                    for _, row in all_demands.drop_duplicates('Conf|version').iterrows():
                        cv = row['Conf|version']
                        if pd.notna(row.get('Prix')) and row['Prix'] > 0:
                            prix_configs[cv] = row['Prix']

                    restant = {}
                    for _, row in all_demands.iterrows():
                        rest = int(row['Restant J1'])
                        if rest > 0:
                            restant[(row['Conf|version'], row['Priorité'])] = rest

                    detail_j2 = []
                    achats_j2 = {}
                    consumed_j2 = {}
                    j2_prod_par_seuil = {}
                    j2_cout_par_seuil = {}
                    achats_detail_par_config = {}

                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_steps = len(SEUILS) * len(priorities)
                    step = 0

                    for seuil in SEUILS:
                        seuil_pct = int(seuil * 100)
                        total_seuil = 0
                        cout_seuil = 0

                        for prio in priorities:
                            step += 1
                            status_text.text(f"Jalon 2 — Seuil {seuil_pct}% / Priorité {prio}...")
                            progress_bar.progress(min(0.99, step / total_steps))

                            configs_prio = sorted(set(cv for (cv, p) in restant if p == prio and restant[(cv, p)] > 0 and cv in q.index and cv in prix_configs))
                            sans_prix_prio = sorted(set(cv for (cv, p) in restant if p == prio and restant[(cv, p)] > 0 and cv in q.index and cv not in prix_configs))

                            if not configs_prio:
                                for cv in sans_prix_prio:
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': restant.get((cv, prio), 0),
                                                      'Prix conf': None, 'Cout achat': None, '% du prix': None, 'Statut': 'Pas de prix'})
                                continue

                            demand_prio = {cv: restant[(cv, prio)] for cv in configs_prio}
                            all_refs_prio = sorted(set(ref for cv in configs_prio for ref in bom_par_config.get(cv, {}).keys()))

                            prob1 = LpProblem(f"J2_S{seuil_pct}_P{prio}_Max", LpMaximize)
                            Xv, Sv, Bv = {}, {}, {}
                            for cv in configs_prio:
                                Xv[cv] = LpVariable(f"X_{sname(cv)}", 0, demand_prio[cv], LpInteger)
                                for ref, bom_qty in bom_par_config[cv].items():
                                    Sv[(cv, ref)] = LpVariable(f"S_{sname(cv)}_{sname(ref)}", 0)
                                    Bv[(cv, ref)] = LpVariable(f"B_{sname(cv)}_{sname(ref)}", 0)
                                    prob1 += Sv[(cv, ref)] + Bv[(cv, ref)] == bom_qty * Xv[cv]
                            prob1 += lpSum(Xv.values())
                            for ref in all_refs_prio:
                                s_vars = [Sv[(cv, ref)] for cv in configs_prio if (cv, ref) in Sv]
                                if s_vars:
                                    dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                    prob1 += lpSum(s_vars) <= dispo
                            for cv in configs_prio:
                                budget_cv = seuil * prix_configs[cv]
                                cout_cv = lpSum(Bv[(cv, ref)] * ref_to_prix.get(ref, 0) for ref in bom_par_config[cv] if (cv, ref) in Bv)
                                prob1 += cout_cv <= budget_cv * Xv[cv]
                            prob1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
                            max_n = int(round(value(prob1.objective) or 0))

                            if max_n == 0:
                                for cv in configs_prio:
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': demand_prio[cv],
                                                      'Prix conf': prix_configs[cv], 'Cout achat': None, '% du prix': None, 'Statut': 'Bloque'})
                                continue

                            prob2 = LpProblem(f"J2_S{seuil_pct}_P{prio}_Min", LpMinimize)
                            X2, S2, B2 = {}, {}, {}
                            for cv in configs_prio:
                                X2[cv] = LpVariable(f"X2_{sname(cv)}", 0, demand_prio[cv], LpInteger)
                                for ref, bom_qty in bom_par_config[cv].items():
                                    S2[(cv, ref)] = LpVariable(f"S2_{sname(cv)}_{sname(ref)}", 0)
                                    B2[(cv, ref)] = LpVariable(f"B2_{sname(cv)}_{sname(ref)}", 0)
                                    prob2 += S2[(cv, ref)] + B2[(cv, ref)] == bom_qty * X2[cv]
                            prob2 += lpSum(B2[(cv, ref)] * ref_to_prix.get(ref, 0)
                                          for cv in configs_prio for ref in bom_par_config[cv] if (cv, ref) in B2)
                            prob2 += lpSum(X2.values()) == max_n
                            for ref in all_refs_prio:
                                s_vars = [S2[(cv, ref)] for cv in configs_prio if (cv, ref) in S2]
                                if s_vars:
                                    dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                    prob2 += lpSum(s_vars) <= dispo
                            for cv in configs_prio:
                                budget_cv = seuil * prix_configs[cv]
                                cout_cv = lpSum(B2[(cv, ref)] * ref_to_prix.get(ref, 0) for ref in bom_par_config[cv] if (cv, ref) in B2)
                                prob2 += cout_cv <= budget_cv * X2[cv]
                            prob2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))

                            for cv in configs_prio:
                                qty = int(round(lp_val(X2[cv])))
                                cout_achat_cv = sum(lp_val(B2[(cv, ref)]) * ref_to_prix.get(ref, 0)
                                                   for ref in bom_par_config[cv] if (cv, ref) in B2)
                                if qty > 0:
                                    pct_moy = (cout_achat_cv / (prix_configs[cv] * qty) * 100) if prix_configs[cv] > 0 else 0
                                    restant[(cv, prio)] -= qty
                                    total_seuil += qty
                                    cout_seuil += cout_achat_cv
                                    refs_list = []
                                    for ref in bom_par_config[cv]:
                                        if (cv, ref) in S2:
                                            used = round(lp_val(S2[(cv, ref)]), 4)
                                            if used > 0 and ref in stock_j2.index:
                                                stock_j2[ref] -= used
                                                consumed_j2[ref] = consumed_j2.get(ref, 0) + used
                                        if (cv, ref) in B2:
                                            bought = round(lp_val(B2[(cv, ref)]), 4)
                                            if bought > 0:
                                                achats_j2[ref] = achats_j2.get(ref, 0) + bought
                                                prix_ref = ref_to_prix.get(ref, 0)
                                                refs_list.append({'Reference': ref, 'Qte achetee': bought,
                                                                  'Prix unitaire': prix_ref, 'Cout total': round(bought * prix_ref, 2)})
                                    if refs_list:
                                        achats_detail_par_config[(cv, prio, seuil_pct)] = {'nb_configs': qty, 'references': refs_list}
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': qty, 'Demande restante': demand_prio[cv],
                                                      'Prix conf': prix_configs[cv], 'Cout achat': round(cout_achat_cv, 2),
                                                      '% du prix': round(pct_moy, 1), 'Statut': 'Produit'})
                                else:
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': demand_prio[cv],
                                                      'Prix conf': prix_configs[cv], 'Cout achat': None,
                                                      '% du prix': None, 'Statut': 'Bloque'})
                            for cv in sans_prix_prio:
                                detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                  'Nb produit': 0, 'Demande restante': restant.get((cv, prio), 0),
                                                  'Prix conf': None, 'Cout achat': None, '% du prix': None, 'Statut': 'Pas de prix'})

                        j2_prod_par_seuil[seuil_pct] = total_seuil
                        j2_cout_par_seuil[seuil_pct] = cout_seuil

                    progress_bar.progress(1.0)
                    status_text.text("Terminé")

                    total_j2 = sum(j2_prod_par_seuil.values())
                    total_produit = total_j1 + total_j2
                    pct_final = (total_produit / total_demande * 100) if total_demande > 0 else 0
                    restant_final = total_demande - total_produit
                    cout_achats_j2 = sum(d['Cout achat'] for d in detail_j2 if d['Cout achat'] is not None)
                    cout_conf_restantes = sum(restant.get((cv, p), 0) * prix_configs.get(cv, 0)
                                             for (cv, p) in restant if restant[(cv, p)] > 0)
                    cout_total_estime = cout_j1 + cout_achats_j2 + cout_conf_restantes

                    st.success("Optimisation Jalon 2 terminée")
                    st.markdown("<br>", unsafe_allow_html=True)

                    st.markdown('<div class="kpi-label">Bilan global J1 + J2</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale", f"{int(total_demande)}")
                    c2.metric("Total produit J1+J2", f"{int(total_produit)}")
                    c3.metric("Taux de réalisation final", f"{pct_final:.1f}%")
                    c4.metric("Configs en rupture finale", f"{int(restant_final)}")

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div class="kpi-label">Décomposition des coûts</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Stock consommé J1", f"{cout_j1:,.0f} €")
                    c2.metric("Achats références J2", f"{cout_achats_j2:,.0f} €")
                    c3.metric("Configs restantes (estimé)", f"{cout_conf_restantes:,.0f} €")
                    c4.metric("Coût total estimé", f"{cout_total_estime:,.0f} €")

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div class="kpi-label">Production cumulée par seuil</div>', unsafe_allow_html=True)
                    cumul = total_j1
                    seuil_rows = [{'Etape': 'Jalon 1 — stock existant', 'Nouvelles configs': int(total_j1),
                                   'Cout (EUR)': f"{cout_j1:,.0f}", 'Cumul': int(total_j1), 'Taux': f"{pct_j1:.1f}%"}]
                    for s_pct in [20, 40, 60, 80]:
                        prod = j2_prod_par_seuil.get(s_pct, 0)
                        cout = j2_cout_par_seuil.get(s_pct, 0)
                        cumul += prod
                        seuil_rows.append({'Etape': f'Seuil {s_pct}%', 'Nouvelles configs': int(prod),
                                           'Cout (EUR)': f"{cout:,.0f}", 'Cumul': int(cumul),
                                           'Taux': f"{(cumul/total_demande*100):.1f}%"})
                    st.dataframe(pd.DataFrame(seuil_rows), use_container_width=True)

                    # ---- KPI PAR CONSTRUCTEUR ----
                    if detail_j2 and 'Constructeur' in all_demands.columns:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown('<div class="kpi-label">Résultats par constructeur</div>', unsafe_allow_html=True)
                        df_j2_all = pd.DataFrame(detail_j2)
                        df_j2_all = df_j2_all.merge(
                            all_demands[['Conf|version','Priorité','Constructeur']].drop_duplicates(),
                            on=['Conf|version','Priorité'], how='left'
                        )
                        constr_j2 = df_j2_all.groupby('Constructeur').agg(
                            Demande_initiale=('Demande restante', 'sum'),
                            Produit_J2=('Nb produit', 'sum'),
                            Cout_achat=('Cout achat', 'sum')
                        ).reset_index()
                        # Ajouter J1 par constructeur
                        j1_constr = all_demands.groupby('Constructeur').agg(
                            Demande_totale=('Demande', 'sum'),
                            J1_produit=('J1 produit', 'sum')
                        ).reset_index()
                        constr_j2 = constr_j2.merge(j1_constr, on='Constructeur', how='left')
                        constr_j2['Total produit'] = constr_j2['J1_produit'].fillna(0) + constr_j2['Produit_J2']
                        constr_j2['Restant'] = constr_j2['Demande_totale'].fillna(0) - constr_j2['Total produit']
                        constr_j2['Taux %'] = (constr_j2['Total produit'] / constr_j2['Demande_totale'].replace(0,1) * 100).round(1).astype(str) + '%'
                        constr_j2['Cout achat J2 (€)'] = constr_j2['Cout_achat'].fillna(0).round(0).astype(int)
                        display_constr = constr_j2[['Constructeur','Demande_totale','J1_produit','Produit_J2','Total produit','Restant','Taux %','Cout achat J2 (€)']].rename(columns={
                            'Demande_totale': 'Demande', 'J1_produit': 'J1 Produit', 'Produit_J2': 'J2 Produit'
                        })
                        st.dataframe(display_constr, use_container_width=True)

                    # ---- KPI PAR SEUIL DÉTAILLÉ ----
                    if detail_j2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown('<div class="kpi-label">Détail par seuil d\'achat</div>', unsafe_allow_html=True)
                        df_j2_all2 = pd.DataFrame(detail_j2)
                        seuil_detail = df_j2_all2.groupby('Seuil').agg(
                            Configs_produites=('Nb produit', 'sum'),
                            Cout_achat=('Cout achat', 'sum'),
                            Nb_bloquees=('Statut', lambda x: (x == 'Bloque').sum()),
                            Nb_sans_prix=('Statut', lambda x: (x == 'Pas de prix').sum())
                        ).reset_index()
                        seuil_detail['Cout achat (€)'] = seuil_detail['Cout_achat'].fillna(0).round(0).astype(int)
                        seuil_detail = seuil_detail[['Seuil','Configs_produites','Cout achat (€)','Nb_bloquees','Nb_sans_prix']].rename(columns={
                            'Configs_produites': 'Configs produites', 'Nb_bloquees': 'Bloquées', 'Nb_sans_prix': 'Sans prix'
                        })
                        st.dataframe(seuil_detail, use_container_width=True)

                    # ---- KPI PAR PRIORITÉ ----
                    if detail_j2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown('<div class="kpi-label">Détail par niveau de priorité</div>', unsafe_allow_html=True)
                        df_j2_prio = pd.DataFrame(detail_j2)
                        prio_j2 = df_j2_prio.groupby('Priorité').agg(
                            Produit_J2=('Nb produit', 'sum'),
                            Cout_achat=('Cout achat', 'sum')
                        ).reset_index()
                        prio_j1 = all_demands.groupby('Priorité').agg(
                            Demande=('Demande', 'sum'),
                            J1_produit=('J1 produit', 'sum')
                        ).reset_index()
                        prio_summary = prio_j1.merge(prio_j2, on='Priorité', how='left')
                        prio_summary['Total produit'] = prio_summary['J1_produit'].fillna(0) + prio_summary['Produit_J2'].fillna(0)
                        prio_summary['Restant'] = prio_summary['Demande'] - prio_summary['Total produit']
                        prio_summary['Taux %'] = (prio_summary['Total produit'] / prio_summary['Demande'].replace(0,1) * 100).round(1).astype(str) + '%'
                        prio_summary['Cout achat J2 (€)'] = prio_summary['Cout_achat'].fillna(0).round(0).astype(int)
                        display_prio = prio_summary[['Priorité','Demande','J1_produit','Produit_J2','Total produit','Restant','Taux %','Cout achat J2 (€)']].rename(columns={
                            'J1_produit': 'J1 Produit', 'Produit_J2': 'J2 Produit'
                        })
                        st.dataframe(display_prio, use_container_width=True)

                    # ---- DÉTAIL CONFIGURATIONS PRODUITES EN J2 ----
                    if detail_j2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown('<div class="kpi-label">Configurations produites en Jalon 2</div>', unsafe_allow_html=True)
                        df_j2_prod = pd.DataFrame([d for d in detail_j2 if d['Statut'] == 'Produit'])
                        if not df_j2_prod.empty:
                            st.dataframe(df_j2_prod[['Seuil','Priorité','Conf|version','Nb produit','Prix conf','Cout achat','% du prix']], use_container_width=True, height=300)

                    # ---- EXPORT EXCEL - 5 onglets identiques au script standalone ----
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div class="kpi-label">Export des résultats</div>', unsafe_allow_html=True)

                    BLEU      = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    BLEU_CLAIR= PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
                    VERT      = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    ORANGE    = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    ROUGE     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    GRIS      = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    BLEU_FONCE= PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
                    JAUNE     = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                    BOLD_BLANC   = Font(bold=True, color="FFFFFF", size=11)
                    BOLD_BLANC_14= Font(bold=True, color="FFFFFF", size=14)
                    BOLD_BLANC_12= Font(bold=True, color="FFFFFF", size=12)
                    BOLD_NOIR    = Font(bold=True, size=11)
                    EURO_FMT = '#,##0 €'
                    NUM_FMT  = '#,##0'

                    def fmt_h(ws, c=None):
                        c = c or BLEU
                        for cell in ws[1]:
                            cell.font = BOLD_BLANC; cell.fill = c
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 30
                        from openpyxl.utils import get_column_letter
                        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

                    def adj_w(ws):
                        for col in ws.columns:
                            ml = max((len(str(cell.value or "")) for cell in col), default=8)
                            ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 50)

                    def ligne_total_j2(ws, row_num, nb_cols):
                        for ci in range(1, nb_cols+1):
                            ws.cell(row=row_num, column=ci).font = BOLD_NOIR
                            ws.cell(row=row_num, column=ci).fill = GRIS

                    # Préparer all_demands pour export
                    for s_pct in [20, 40, 60, 80]:
                        all_demands[f'{s_pct}%'] = 0
                        all_demands[f'Cout {s_pct}%'] = 0.0
                    for d in detail_j2:
                        if d['Statut'] == 'Produit':
                            mask = (all_demands['Conf|version'] == d['Conf|version']) & (all_demands['Priorité'] == d['Priorité'])
                            s_col = d['Seuil']
                            all_demands.loc[mask, s_col] = all_demands.loc[mask, s_col] + d['Nb produit']
                            all_demands.loc[mask, f"Cout {d['Seuil']}"] = all_demands.loc[mask, f"Cout {d['Seuil']}"] + (d['Cout achat'] or 0)

                    all_demands_xls = all_demands.rename(columns={'J1 produit': 'Jalon 1'})
                    all_demands_xls['Total'] = (all_demands_xls['Jalon 1'] + all_demands_xls['20%'] + all_demands_xls['40%'] + all_demands_xls['60%'] + all_demands_xls['80%'])
                    all_demands_xls['Restant'] = all_demands_xls['Demande'] - all_demands_xls['Total']
                    all_demands_xls['RUPTURE'] = all_demands_xls['Restant'].apply(lambda x: 'OUI' if x > 0 else '')
                    all_demands_xls['Cout conf restantes'] = all_demands_xls['Restant'] * all_demands_xls['Prix'].fillna(0)

                    def _compute_seuil_rupture(row):
                        if row['Restant'] == 0: return ''
                        last = None
                        for s in [20, 40, 60, 80]:
                            if row[f'{s}%'] > 0: last = s
                        if last is None:
                            return 'apres Jalon 1' if row['Jalon 1'] > 0 else 'des le depart'
                        seuils_list = [20, 40, 60, 80]
                        idx = seuils_list.index(last)
                        return f'a {seuils_list[idx+1]}%' if idx < len(seuils_list)-1 else 'apres 80%'
                    all_demands_xls['Seuil rupture'] = all_demands_xls.apply(_compute_seuil_rupture, axis=1)

                    wb = Workbook()

                    # ---- Onglet 1 : Demande ----
                    ws_d = wb.active; ws_d.title = "Demande"
                    cols_d = [c for c in ['Constructeur','Conf','Version','Conf|version','Priorité','Demande','Prix'] if c in all_demands_xls.columns]
                    df_d = all_demands_xls[cols_d].drop_duplicates().sort_values(['Constructeur','Priorité','Conf|version'] if 'Constructeur' in all_demands_xls.columns else ['Priorité','Conf|version'])
                    tot_row_d = ws_d.max_row + 1
                    for rr in dataframe_to_rows(df_d, index=False, header=True): ws_d.append(rr)
                    tot_row_d = ws_d.max_row + 1
                    ws_d.cell(tot_row_d, 1, "TOTAL")
                    cidx_d = {cols_d[i]: i+1 for i in range(len(cols_d))}
                    if 'Demande' in cidx_d: ws_d.cell(tot_row_d, cidx_d['Demande'], df_d['Demande'].sum())
                    ligne_total_j2(ws_d, tot_row_d, len(cols_d))
                    fmt_h(ws_d); adj_w(ws_d)
                    for row in ws_d.iter_rows(min_row=2, max_row=ws_d.max_row):
                        for cell in row:
                            h = ws_d.cell(1, cell.column).value
                            if h == 'Prix' and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                            elif h == 'Demande' and isinstance(cell.value, (int, float)): cell.number_format = NUM_FMT

                    # ---- Onglet 2 : BOM ----
                    ws_bom_j2 = wb.create_sheet("BOM")
                    bom_j2_exp = bom_df[['Conf|version','Référence','Quantité']].copy().sort_values(['Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_j2_exp, index=False, header=True): ws_bom_j2.append(rr)
                    fmt_h(ws_bom_j2); adj_w(ws_bom_j2)

                    # ---- Onglet 3 : Stock ----
                    ws_s = wb.create_sheet("Stock")
                    stock_out = stock_df.copy().rename(columns={'NVX STOCK': 'Stock initial', 'Prix (pj)': 'Prix unitaire'})
                    stock_out['Consomme J1'] = stock_out['Referentiel'].map(consumed_j1).fillna(0).round(2)
                    stock_out['Consomme J2'] = stock_out['Referentiel'].map(consumed_j2).fillna(0).round(2)
                    stock_out['Achete J2']   = stock_out['Referentiel'].map(achats_j2).fillna(0).round(2)
                    stock_out['Cout achats J2'] = (stock_out['Achete J2'] * stock_out['Prix unitaire']).round(2)
                    stock_out['Stock final'] = (stock_out['Stock initial'] - stock_out['Consomme J1'] - stock_out['Consomme J2']).round(2)
                    cols_s = ['Referentiel','Stock initial','Prix unitaire','Consomme J1','Consomme J2','Achete J2','Cout achats J2','Stock final']
                    stock_out = stock_out[cols_s].sort_values('Referentiel')
                    for rr in dataframe_to_rows(stock_out, index=False, header=True): ws_s.append(rr)
                    tot_s = ws_s.max_row + 1
                    ws_s.cell(tot_s, 1, "TOTAL")
                    cidx_s = {cols_s[i]: i+1 for i in range(len(cols_s))}
                    for cn in ['Consomme J1','Consomme J2','Achete J2','Cout achats J2']:
                        if cn in cidx_s: ws_s.cell(tot_s, cidx_s[cn], stock_out[cn].sum())
                    ligne_total_j2(ws_s, tot_s, len(cols_s))
                    fmt_h(ws_s); adj_w(ws_s)
                    for row in ws_s.iter_rows(min_row=2, max_row=ws_s.max_row):
                        for cell in row:
                            h = ws_s.cell(1, cell.column).value
                            if h in ['Prix unitaire','Cout achats J2'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                    col_sf_s = cidx_s.get('Stock final')
                    if col_sf_s:
                        for rn in range(2, ws_s.max_row):
                            c = ws_s.cell(rn, col_sf_s)
                            if isinstance(c.value, (int, float)):
                                if c.value < 0: c.fill = ROUGE
                                elif c.value == 0: c.fill = ORANGE

                    # ---- Onglet 4 : Résultat (format complet) ----
                    ws_r = wb.create_sheet("Résultat")
                    r_row = 1

                    # -- Section A : Bilan Global 5 colonnes --
                    ws_r.cell(r_row, 1, "BILAN GLOBAL").font = BOLD_BLANC_14
                    for ci in range(1, 6): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 2
                    for i, h in enumerate(["Description","Nb configs","% demande","Cout reel","Cout si conf completes"], 1):
                        ws_r.cell(r_row, i, h).font = BOLD_NOIR; ws_r.cell(r_row, i).fill = BLEU_CLAIR
                    r_row += 1

                    # Calculs économies J2
                    cout_j2_si_complet_par_seuil = {20: 0.0, 40: 0.0, 60: 0.0, 80: 0.0}
                    for d in detail_j2:
                        if d['Statut'] == 'Produit':
                            sn = int(d['Seuil'].replace('%',''))
                            cout_j2_si_complet_par_seuil[sn] += d['Nb produit'] * prix_configs.get(d['Conf|version'], 0)
                    cout_j2_si_complet = sum(cout_j2_si_complet_par_seuil.values())
                    economie_j2 = cout_j2_si_complet - cout_achats_j2

                    # Total demande
                    ws_r.cell(r_row, 1, "Total demande"); ws_r.cell(r_row, 2, total_demande).number_format = NUM_FMT
                    r_row += 2

                    # Jalon 1
                    ws_r.cell(r_row, 1, "JALON 1 - Production avec stock existant")
                    for ci in range(1, 6): ws_r.cell(r_row, ci).font = BOLD_NOIR; ws_r.cell(r_row, ci).fill = BLEU_CLAIR
                    r_row += 1
                    ws_r.cell(r_row, 1, "  Configs produites"); ws_r.cell(r_row, 2, total_j1).number_format = NUM_FMT
                    ws_r.cell(r_row, 3, f"{pct_j1:.1f}%"); ws_r.cell(r_row, 4, cout_j1).number_format = EURO_FMT
                    r_row += 2

                    # Jalon 2
                    ws_r.cell(r_row, 1, "JALON 2 - Production avec achats de references unitaires")
                    for ci in range(1, 6): ws_r.cell(r_row, ci).font = BOLD_NOIR; ws_r.cell(r_row, ci).fill = BLEU_CLAIR
                    r_row += 1
                    for sp in [20, 40, 60, 80]:
                        ws_r.cell(r_row, 1, f"  Seuil {sp}%")
                        ws_r.cell(r_row, 2, j2_prod_par_seuil.get(sp, 0)).number_format = NUM_FMT
                        ws_r.cell(r_row, 4, j2_cout_par_seuil.get(sp, 0)).number_format = EURO_FMT
                        ws_r.cell(r_row, 5, cout_j2_si_complet_par_seuil.get(sp, 0)).number_format = EURO_FMT
                        ws_r.cell(r_row, 5).font = Font(italic=True, size=10)
                        r_row += 1
                    pct_j2_s = f"{total_j2/total_demande*100:.1f}%" if total_demande > 0 else ""
                    ws_r.cell(r_row, 1, "  Sous-total Jalon 2"); ws_r.cell(r_row, 2, total_j2).number_format = NUM_FMT
                    ws_r.cell(r_row, 3, pct_j2_s); ws_r.cell(r_row, 4, cout_achats_j2).number_format = EURO_FMT
                    ws_r.cell(r_row, 5, cout_j2_si_complet).number_format = EURO_FMT
                    for ci in range(1, 6): ws_r.cell(r_row, ci).font = BOLD_NOIR
                    r_row += 1
                    ws_r.cell(r_row, 1, "  \u2192 Cout economise (refs unitaires vs achat conf completes)")
                    ws_r.cell(r_row, 4, economie_j2).number_format = EURO_FMT
                    ws_r.cell(r_row, 4).font = BOLD_NOIR
                    for ci in range(1, 6): ws_r.cell(r_row, ci).fill = VERT
                    r_row += 2

                    # Restant en rupture
                    ws_r.cell(r_row, 1, "RESTANT EN RUPTURE")
                    for ci in range(1, 6): ws_r.cell(r_row, ci).font = BOLD_NOIR; ws_r.cell(r_row, ci).fill = ORANGE
                    r_row += 1
                    ws_r.cell(r_row, 1, "  (necessitent achat config complete)")
                    ws_r.cell(r_row, 2, restant_final).number_format = NUM_FMT
                    ws_r.cell(r_row, 3, f"{restant_final/total_demande*100:.1f}%" if total_demande > 0 else "")
                    r_row += 3

                    # -- Section B : Detail par priorité --
                    ws_r.cell(r_row, 1, "DETAIL PAR PRIORITE").font = BOLD_BLANC_14
                    for ci in range(1, 9): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 1
                    for i, h in enumerate(["Priorite","Demande","Jalon 1","J2 Total","Cout J2","Total","Restant","% couvert"], 1):
                        ws_r.cell(r_row, i, h).font = BOLD_NOIR; ws_r.cell(r_row, i).fill = BLEU_CLAIR
                        ws_r.cell(r_row, i).alignment = Alignment(horizontal="center")
                    r_row += 1
                    for prio in priorities:
                        prio_data = all_demands_xls[all_demands_xls['Priorité'] == prio]
                        dem_p = int(prio_data['Demande'].sum())
                        j1_p  = int(prio_data['Jalon 1'].sum())
                        j2_p  = int(prio_data[['20%','40%','60%','80%']].sum().sum())
                        cout_j2_p = prio_data[['Cout 20%','Cout 40%','Cout 60%','Cout 80%']].sum().sum()
                        tot_p = j1_p + j2_p; rest_p = dem_p - tot_p
                        pct_p = f"{tot_p/dem_p*100:.0f}%" if dem_p > 0 else "0%"
                        ws_r.cell(r_row, 1, prio).alignment = Alignment(horizontal="center")
                        ws_r.cell(r_row, 2, dem_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 3, j1_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 4, j2_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 5, round(cout_j2_p, 2)).number_format = EURO_FMT
                        ws_r.cell(r_row, 6, tot_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 7, rest_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 8, pct_p).alignment = Alignment(horizontal="center")
                        if tot_p == dem_p: ws_r.cell(r_row, 8).fill = VERT
                        elif tot_p == 0: ws_r.cell(r_row, 8).fill = ROUGE
                        else: ws_r.cell(r_row, 8).fill = ORANGE
                        r_row += 1

                    # -- Section C : Detail par configuration --
                    r_row += 3
                    ws_r.cell(r_row, 1, "DETAIL PAR CONFIGURATION").font = BOLD_BLANC_14
                    for ci in range(1, 22): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 1
                    cols_cfg = [c for c in ['Constructeur','Conf','Version','Conf|version','Priorité','Demande','Prix',
                                            'Jalon 1','20%','Cout 20%','40%','Cout 40%','60%','Cout 60%','80%','Cout 80%',
                                            'Total','Restant','RUPTURE','Seuil rupture','Cout conf restantes'] if c in all_demands_xls.columns]
                    for i, h in enumerate(cols_cfg, 1):
                        ws_r.cell(r_row, i, h).font = BOLD_NOIR; ws_r.cell(r_row, i).fill = BLEU_CLAIR
                        ws_r.cell(r_row, i).alignment = Alignment(horizontal="center", wrap_text=True)
                    r_row += 1
                    df_cfg = all_demands_xls[cols_cfg].sort_values(['Constructeur','Priorité','Conf|version'] if 'Constructeur' in all_demands_xls.columns else ['Priorité','Conf|version'])
                    start_cfg = r_row
                    for row_data in dataframe_to_rows(df_cfg, index=False, header=False):
                        for ci, val in enumerate(row_data, 1): ws_r.cell(r_row, ci, val)
                        r_row += 1
                    # Ligne TOTAL
                    ws_r.cell(r_row, 1, "TOTAL")
                    cidx_cfg = {cols_cfg[i]: i+1 for i in range(len(cols_cfg))}
                    for cn in ['Demande','Jalon 1','20%','Cout 20%','40%','Cout 40%','60%','Cout 60%','80%','Cout 80%','Total','Restant','Cout conf restantes']:
                        if cn in cidx_cfg: ws_r.cell(r_row, cidx_cfg[cn], df_cfg[cn].sum())
                    ligne_total_j2(ws_r, r_row, len(cols_cfg))
                    # Formatage Section C
                    for rn in range(start_cfg, r_row + 1):
                        for ci in range(1, len(cols_cfg)+1):
                            cell = ws_r.cell(rn, ci); h = cols_cfg[ci-1]
                            if h in ['Prix','Cout 20%','Cout 40%','Cout 60%','Cout 80%','Cout conf restantes'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                            elif h in ['Demande','Jalon 1','20%','40%','60%','80%','Total','Restant'] and isinstance(cell.value, (int, float)): cell.number_format = NUM_FMT
                    if 'RUPTURE' in cidx_cfg:
                        for rn in range(start_cfg, r_row):
                            c = ws_r.cell(rn, cidx_cfg['RUPTURE'])
                            if c.value == 'OUI': c.fill = ROUGE; c.font = BOLD_NOIR
                            elif c.value == '': c.fill = VERT
                    if 'Restant' in cidx_cfg:
                        for rn in range(start_cfg, r_row):
                            c = ws_r.cell(rn, cidx_cfg['Restant'])
                            if isinstance(c.value, (int, float)):
                                if c.value == 0: c.fill = VERT
                                elif c.value > 0: c.fill = ORANGE
                    if 'Seuil rupture' in cidx_cfg:
                        for rn in range(start_cfg, r_row):
                            c = ws_r.cell(rn, cidx_cfg['Seuil rupture'])
                            if c.value and str(c.value).strip(): c.fill = ORANGE

                    # -- Section D : Configs en rupture --
                    r_row += 3
                    ws_r.cell(r_row, 1, "CONFIGURATIONS EN RUPTURE - ACHAT COMPLET NECESSAIRE").font = BOLD_BLANC_14
                    for ci in range(1, 7): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 1
                    for i, h in enumerate(["Priorite","Conf|version","Qte restante","Dernier seuil tente","Prix unitaire","Cout total"], 1):
                        ws_r.cell(r_row, i, h).font = BOLD_NOIR; ws_r.cell(r_row, i).fill = BLEU_CLAIR
                        ws_r.cell(r_row, i).alignment = Alignment(horizontal="center")
                    r_row += 1
                    total_cout_rupt = 0
                    for prio in priorities:
                        for (cv, p) in sorted(restant):
                            if p == prio and restant[(cv, p)] > 0:
                                qty_r = restant[(cv, p)]; prix_cv = prix_configs.get(cv, 0); cout_r = qty_r * prix_cv
                                total_cout_rupt += cout_r
                                row_m = all_demands_xls[(all_demands_xls['Conf|version'] == cv) & (all_demands_xls['Priorité'] == p)]
                                dern_seuil = row_m.iloc[0].get('Seuil rupture', '') if not row_m.empty else ''
                                ws_r.cell(r_row, 1, prio).alignment = Alignment(horizontal="center")
                                ws_r.cell(r_row, 2, cv)
                                ws_r.cell(r_row, 3, qty_r).number_format = NUM_FMT
                                ws_r.cell(r_row, 4, dern_seuil)
                                ws_r.cell(r_row, 5, prix_cv).number_format = EURO_FMT
                                ws_r.cell(r_row, 6, cout_r).number_format = EURO_FMT
                                r_row += 1
                    if restant_final > 0:
                        ws_r.cell(r_row, 1, "TOTAL"); ws_r.cell(r_row, 3, restant_final).number_format = NUM_FMT
                        ws_r.cell(r_row, 6, total_cout_rupt).number_format = EURO_FMT
                        ligne_total_j2(ws_r, r_row, 6)

                    # Largeurs colonnes Résultat
                    ws_r.column_dimensions['A'].width = 50; ws_r.column_dimensions['B'].width = 22
                    ws_r.column_dimensions['C'].width = 16; ws_r.column_dimensions['D'].width = 22
                    ws_r.column_dimensions['E'].width = 22; ws_r.column_dimensions['F'].width = 16
                    for cl in ['G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U']: ws_r.column_dimensions[cl].width = 14

                    # ---- Onglet 5 : Detail J2 ----
                    if detail_j2:
                        ws_j2 = wb.create_sheet("Detail J2")
                        df_dj2 = pd.DataFrame(detail_j2).copy()
                        df_dj2['Budget max'] = df_dj2.apply(
                            lambda row: (int(row['Seuil'].replace('%','')) / 100) * row['Prix conf'] if pd.notna(row.get('Prix conf')) else None, axis=1)
                        df_dj2 = df_dj2.rename(columns={'Nb produit': 'Produit', 'Prix conf': 'Prix config'})
                        cols_dj2 = [c for c in ['Seuil','Priorité','Conf|version','Demande restante','Produit','Budget max','Cout achat','% du prix','Statut'] if c in df_dj2.columns]
                        df_dj2 = df_dj2[cols_dj2].sort_values(['Seuil','Priorité','Conf|version'])
                        for rr in dataframe_to_rows(df_dj2, index=False, header=True): ws_j2.append(rr)
                        tot_dj2 = ws_j2.max_row + 1
                        ws_j2.cell(tot_dj2, 1, "TOTAL")
                        cidx_dj2 = {cols_dj2[i]: i+1 for i in range(len(cols_dj2))}
                        if 'Produit' in cidx_dj2: ws_j2.cell(tot_dj2, cidx_dj2['Produit'], df_dj2['Produit'].sum())
                        if 'Cout achat' in cidx_dj2: ws_j2.cell(tot_dj2, cidx_dj2['Cout achat'], df_dj2['Cout achat'].sum())
                        ligne_total_j2(ws_j2, tot_dj2, len(cols_dj2))
                        fmt_h(ws_j2); adj_w(ws_j2)
                        for row in ws_j2.iter_rows(min_row=2, max_row=ws_j2.max_row):
                            for cell in row:
                                h = ws_j2.cell(1, cell.column).value
                                if h in ['Prix config','Budget max','Cout achat'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                        col_stat = cidx_dj2.get('Statut')
                        if col_stat:
                            for rn in range(2, ws_j2.max_row):
                                c = ws_j2.cell(rn, col_stat)
                                if c.value == 'Produit': c.fill = VERT
                                elif c.value == 'Bloque': c.fill = ORANGE
                                elif c.value == 'Pas de prix': c.fill = GRIS
                        # Bonus : detail refs achetées par config
                        if achats_detail_par_config:
                            r_bon = ws_j2.max_row + 3
                            ws_j2.cell(r_bon, 1, "DETAIL DES REFERENCES ACHETEES PAR CONFIG").font = BOLD_BLANC_14
                            for ci in range(1, 7): ws_j2.cell(r_bon, ci).fill = BLEU_FONCE
                            r_bon += 2
                            for (cv, prio_b, sp_b), details in sorted(achats_detail_par_config.items()):
                                ws_j2.cell(r_bon, 1, f"Config: {cv}").font = BOLD_NOIR
                                ws_j2.cell(r_bon, 2, f"Priorité: {prio_b}")
                                ws_j2.cell(r_bon, 3, f"Seuil: {sp_b}%")
                                ws_j2.cell(r_bon, 4, f"Nb configs: {details['nb_configs']}")
                                for ci in range(1, 7): ws_j2.cell(r_bon, ci).fill = JAUNE
                                r_bon += 1
                                for i, h in enumerate(["Référence","Qté achetée","Prix unitaire","Coût total"], 1):
                                    ws_j2.cell(r_bon, i, h).font = BOLD_NOIR; ws_j2.cell(r_bon, i).fill = GRIS
                                r_bon += 1
                                for ref_d in details['references']:
                                    ws_j2.cell(r_bon, 1, ref_d['Reference'])
                                    ws_j2.cell(r_bon, 2, ref_d['Qte achetee']).number_format = '#,##0.00'
                                    ws_j2.cell(r_bon, 3, ref_d['Prix unitaire']).number_format = EURO_FMT
                                    ws_j2.cell(r_bon, 4, ref_d['Cout total']).number_format = EURO_FMT
                                    r_bon += 1
                                r_bon += 1

                    output = BytesIO()
                    wb.save(output)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state['j2_excel'] = output.getvalue()
                    st.session_state['j2_excel_name'] = f"RO_Production_{ts}.xlsx"

                except Exception as e:
                    st.error(f"Erreur Jalon 2 : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        if 'j2_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="Télécharger le fichier Excel complet",
                data=st.session_state['j2_excel'],
                file_name=st.session_state['j2_excel_name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )


# =============================================================================
# FOOTER
# =============================================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.divider()
st.markdown('<div class="footer-text">FORECAST SFR — Batail-log &nbsp;·&nbsp; PuLP CBC Solver &nbsp;·&nbsp; v1.0</div>', unsafe_allow_html=True)

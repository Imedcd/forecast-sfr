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
    page_title="FORECAST",
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
    .main-title span { color: #457DD9; }
    .main-badge {
        font-size: 0.7rem; background: rgba(69,125,217,0.1); color: #457DD9;
        border: 1px solid rgba(69,125,217,0.3); padding: 0.3rem 0.8rem;
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
    .stTabs [aria-selected="true"] { background: #457DD9 !important; color: #FFFFFF !important; font-weight: 600; }
    .info-block {
        background: #131720; border: 1px solid #1E2A3A; border-left: 3px solid #457DD9;
        border-radius: 10px; padding: 1.25rem 1.5rem; margin: 1rem 0 1.5rem 0;
        font-size: 0.9rem; color: #B0B8C4; line-height: 1.8;
    }
    .warning-block {
        background: #131720; border: 1px solid #1E2A3A; border-left: 3px solid #FFB300;
        border-radius: 10px; padding: 1.25rem 1.5rem; margin: 1rem 0 1.5rem 0;
        font-size: 0.9rem; color: #B0B8C4;
    }
    .stButton > button[kind="primary"] {
        background: #457DD9; color: #FFFFFF; font-weight: 600; border: none;
        border-radius: 8px; padding: 0.65rem 2rem; font-size: 0.9rem; transition: all 0.2s;
    }
    .stButton > button[kind="primary"]:hover {
        background: #5B8FE0; box-shadow: 0 0 20px rgba(69,125,217,0.35); transform: translateY(-1px);
    }
    .stDownloadButton > button {
        background: #457DD9 !important; color: #FFFFFF !important; font-weight: 600 !important;
        border: none !important; border-radius: 8px !important; padding: 0.65rem 2rem !important;
        font-size: 0.9rem !important; transition: all 0.2s !important;
    }
    .stDownloadButton > button:hover {
        background: #5B8FE0 !important; box-shadow: 0 0 20px rgba(69,125,217,0.35) !important;
        transform: translateY(-1px) !important;
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
    [data-testid="stMetricValue"] { font-size: 1.8rem; font-weight: 700; color: #457DD9; }
    [data-testid="stFileUploader"] {
        background: #131720; border: 1.5px dashed #457DD9; border-radius: 10px; padding: 0.5rem;
    }
    [data-testid="stFileUploader"] label { color: #B0B8C4 !important; font-weight: 500; }
    [data-testid="stFileUploaderDropzone"] {
        background: rgba(69,125,217,0.05) !important; border: 1.5px dashed #457DD9 !important; border-radius: 8px !important;
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
                    FORECAST
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
        <div class="main-title">FORECAST &nbsp;—&nbsp; Optimisation RO</div>
        <div style="font-size:0.82rem; color:#6B7689; margin-top:0.3rem;">Allocation optimale de composants télécom</div>
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
        df['Constructeur'] = 'NOKIA'
    elif sheet_name == "huawei":
        df = pd.read_excel(biblio_bytes, sheet_name=sheet_name, header=0)
        df = df[['CONF', 'Qty', 'Radical text']].copy()
        df.columns = ["Conf|version", "Quantite", "Reference"]
        df['Constructeur'] = 'Huawei'
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
    return df[["Constructeur", "Conf|version", "Conf", "Version", "Reference", "Quantite"]]


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

    df_stock = pd.read_excel(stock_file, sheet_name="Stock", usecols=["Code", "Designation", "Stock Dispo", "Prévisionnel", "Stock shields"])
    df_stock = df_stock.rename(columns={"Code": "Référence"})
    df_stock["Designation"] = df_stock["Designation"].fillna("")
    df_stock["Stock"] = (pd.to_numeric(df_stock["Stock Dispo"], errors='coerce').fillna(0) +
                         pd.to_numeric(df_stock["Prévisionnel"], errors='coerce').fillna(0) +
                         pd.to_numeric(df_stock["Stock shields"], errors='coerce').fillna(0))
    df_stock = df_stock[["Référence", "Designation", "Stock"]].copy()
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
    stock_df = df_stock.groupby("Referentiel", as_index=False).agg({"Valeur_NVX": "sum", "Stock": "sum", "Designation": "first"})
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


def run_jalon1(all_demands, bom_df, stock_df, q, priorities, prix_file_bytes, next_map=None):
    if next_map is None:
        next_map = {}
    stock_optim = stock_df.set_index("Referentiel")
    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)
    stock_current = stock_optim['NVX STOCK'].copy()
    produced_j1 = {}
    produced_next_j1 = {}

    progress_bar = st.progress(0)
    status_text = st.empty()

    for phase, prio_level in enumerate(priorities, 1):
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

        # --- Passe de substitution (Next) pour cette priorité ---
        if next_map:
            next_candidates = {}
            for cv in X:
                qty_orig = produced_j1.get((cv, prio_level), 0)
                remaining = demand_map.get(cv, 0) - qty_orig
                if remaining > 0 and cv in next_map:
                    next_cv = next_map[cv]
                    if next_cv in q.index:
                        next_candidates[cv] = (next_cv, remaining)
            if next_candidates:
                prob_n1 = LpProblem(f"Next_P{prio_level}_MaxQty", LpMaximize)
                XN1 = {}
                for cv, (next_cv, rem) in next_candidates.items():
                    safe = cv.replace('|','_').replace(' ','_')
                    XN1[cv] = LpVariable(f"XN1_{safe}_p{prio_level}", 0, rem, LpInteger)
                prob_n1 += lpSum(XN1.values())
                for ref in q.columns:
                    coeffs = []
                    for cv, (next_cv, rem) in next_candidates.items():
                        coef = q.at[next_cv, ref]
                        if coef > 0:
                            coeffs.append((coef, XN1[cv]))
                    if coeffs:
                        lim = stock_current[ref] if ref in stock_current.index else 0
                        prob_n1 += lpSum(c * v for c, v in coeffs) <= lim
                prob_n1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
                max_next = value(prob_n1.objective) or 0
                if max_next > 0:
                    prob_n2 = LpProblem(f"Next_P{prio_level}_MaxCost", LpMaximize)
                    XN = {}
                    for cv, (next_cv, rem) in next_candidates.items():
                        safe = cv.replace('|','_').replace(' ','_')
                        XN[cv] = LpVariable(f"XN_{safe}_p{prio_level}", 0, rem, LpInteger)
                    cout_next = lpSum(
                        XN[cv] * sum(q.at[next_candidates[cv][0], ref] * stock_optim.at[ref, 'Prix (pj)']
                                     for ref in q.columns if ref in stock_optim.index
                                     and q.at[next_candidates[cv][0], ref] > 0)
                        for cv in XN
                    )
                    prob_n2 += cout_next
                    prob_n2 += lpSum(XN.values()) == max_next
                    for ref in q.columns:
                        coeffs = []
                        for cv, (next_cv, rem) in next_candidates.items():
                            coef = q.at[next_cv, ref]
                            if coef > 0:
                                coeffs.append((coef, XN[cv]))
                        if coeffs:
                            lim = stock_current[ref] if ref in stock_current.index else 0
                            prob_n2 += lpSum(c * v for c, v in coeffs) <= lim
                    prob_n2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))
                    for cv in XN:
                        qty_n = max(0, int(round(value(XN[cv].varValue) or 0)))
                        if qty_n > 0:
                            produced_next_j1[(cv, prio_level)] = qty_n
                            next_cv = next_candidates[cv][0]
                            for ref in q.columns:
                                cons = q.at[next_cv, ref] * qty_n
                                if cons > 0 and ref in stock_current.index:
                                    stock_current[ref] = max(0, stock_current[ref] - cons)

    progress_bar.progress(1.0)

    consumed_j1 = {}
    for (cv, prio), qty in produced_j1.items():
        if cv in q.index and qty > 0:
            for ref in q.columns:
                cons = q.at[cv, ref] * qty
                if cons > 0:
                    consumed_j1[ref] = consumed_j1.get(ref, 0) + cons
    for (cv, prio), qty in produced_next_j1.items():
        next_cv = next_map.get(cv, '')
        if next_cv and next_cv in q.index and qty > 0:
            for ref in q.columns:
                cons = q.at[next_cv, ref] * qty
                if cons > 0:
                    consumed_j1[ref] = consumed_j1.get(ref, 0) + cons
    cout_j1 = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed_j1.items() if ref in stock_optim.index)
    return produced_j1, consumed_j1, cout_j1, stock_current.copy(), stock_optim, produced_next_j1


# =============================================================================
# ONGLETS
# =============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "  Chargement des données  ",
    "  Optimisation du réemploi  ",
    "  Optimisation du réemploi avec priorité  ",
    "  Optimisation du réemploi avec priorité et achat de références unitaires  "
])


# =============================================================================
# TAB 1 : UPLOAD
# =============================================================================
with tab1:
    st.markdown('<div class="section-title">Préparation des données</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Chargez les 6 fichiers Excel ci-dessous, puis cliquez sur <strong style="color:#FFFFFF;">Préparer les données</strong>.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown('<div class="label-cat">Fichiers requis</div>', unsafe_allow_html=True)
        biblio_file = st.file_uploader("📂 biblio.xlsx — BOM Nokia + Huawei", type="xlsx", key="biblio")
        prix_file = st.file_uploader("📂 Prix.xlsx — Prix des configurations et références", type="xlsx", key="prix")
        stock_file = st.file_uploader("📂 Stock.xlsx — Stock disponible", type="xlsx", key="stock")
    with col2:
        acc_file = st.file_uploader("📂 acc.xlsx — Table de conversion des références", type="xlsx", key="acc")
        prio_file = st.file_uploader("📂 Prio.xlsx — Demandes et priorités", type="xlsx", key="prio")

    st.markdown("<br>", unsafe_allow_html=True)
    files_common = all([biblio_file, prix_file, stock_file, acc_file])
    if files_common:
        if st.button("Préparer les données", type="primary"):
            with st.spinner("Chargement et normalisation..."):
                try:
                    bom_df, stock_df, ref_to_prix = prepare_common_data(biblio_file, prix_file, stock_file, acc_file)
                    st.session_state['bom_df'] = bom_df
                    st.session_state['stock_df'] = stock_df
                    st.session_state['ref_to_prix'] = ref_to_prix
                    st.session_state['data_prepared'] = True
                    st.success("Données prêtes")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Nombre de configurations", len(bom_df))
                    c2.metric("Références stock", len(stock_df))
                    val_stock = (stock_df['NVX STOCK'] * stock_df['Prix (pj)']).sum()
                    c3.metric("Valeur stock total", f"{val_stock:,.0f} €".replace(',', ' '))
                except Exception as e:
                    st.error(f"Erreur : {e}")


# =============================================================================
# TAB 2 : JALON 1 SANS PRIORITE
# =============================================================================
with tab2:
    st.markdown('<div class="section-title">Optimisation du réemploi</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Maximise le nombre total de configurations avec le stock disponible.</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode en 2 étapes</strong><br>
        1. Maximiser le nombre de configurations réalisables<br>
        2. À nombre fixé, maximiser la valeur du stock consommé
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Chargement des données")
    elif not prio_file:
        st.warning("Chargez le fichier Prio.xlsx dans l'onglet Chargement des données")
    else:
        if st.button("Lancer l'optimisation", key="run_j1_sans", type="primary"):
            with st.spinner("Résolution en cours..."):
                try:
                    # Lire Prio.xlsx et agréger par Conf (le solveur choisit les versions)
                    xl_fc = pd.ExcelFile(BytesIO(prio_file.getvalue()), engine="openpyxl")
                    sheet_fc = xl_fc.sheet_names[0]
                    df_prio_raw = pd.read_excel(BytesIO(prio_file.getvalue()), sheet_name=sheet_fc, engine="openpyxl")
                    df_prio_raw.columns = df_prio_raw.columns.str.strip()
                    cols_utiles = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Demande']
                    df_prio_raw = df_prio_raw[[c for c in cols_utiles if c in df_prio_raw.columns]].copy()
                    df_prio_raw['Demande'] = pd.to_numeric(df_prio_raw['Demande'], errors='coerce').fillna(0).astype(int)
                    df_prio_raw = df_prio_raw[df_prio_raw['Demande'] > 0].copy()

                    # Agréger la demande par Conf
                    demand_by_conf = df_prio_raw.groupby('Conf', as_index=False)['Demande'].sum()
                    demand_conf_map = dict(zip(demand_by_conf['Conf'], demand_by_conf['Demande']))

                    # Versions disponibles par Conf
                    versions_by_conf = {}
                    for conf in demand_conf_map:
                        cvs = df_prio_raw[df_prio_raw['Conf'] == conf]['Conf|version'].unique().tolist()
                        versions_by_conf[conf] = cvs

                    # Détail par version (pour l'affichage)
                    df_versions = df_prio_raw[['Constructeur', 'Conf', 'Version', 'Conf|version']].drop_duplicates('Conf|version')

                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()
                    stock_optim = stock_df.set_index("Referentiel")
                    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)

                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)

                    # Toutes les Conf|versions avec BOM disponible
                    all_cvs = []
                    for conf, cvs in versions_by_conf.items():
                        for cv in cvs:
                            if cv in q.index:
                                all_cvs.append(cv)

                    cv_to_conf = {}
                    for conf, cvs in versions_by_conf.items():
                        for cv in cvs:
                            cv_to_conf[cv] = conf

                    # Étape 1 : max nombre de configs
                    prob1 = LpProblem("Max_Nb_Configs", LpMaximize)
                    X1 = {}
                    for cv in all_cvs:
                        safe = cv.replace('|','_').replace(' ','_').replace('-','_')
                        X1[cv] = LpVariable(f"X1_{safe}", 0, cat=LpInteger)
                    prob1 += lpSum(X1.values())
                    for conf, dem in demand_conf_map.items():
                        cvs_conf = [cv for cv in versions_by_conf[conf] if cv in X1]
                        if cvs_conf:
                            prob1 += lpSum(X1[cv] for cv in cvs_conf) <= dem
                    for ref in q.columns:
                        coeffs = [(q.at[cv, ref], X1[cv]) for cv in X1 if q.at[cv, ref] > 0]
                        if coeffs:
                            lim = stock_optim.at[ref, 'NVX STOCK'] if ref in stock_optim.index else 0
                            prob1 += lpSum(coef * var for coef, var in coeffs) <= lim
                    prob1.solve(PULP_CBC_CMD(msg=0, timeLimit=900, gapRel=0.0))
                    max_configs = value(prob1.objective) or 0

                    # Étape 2 : max coût stock consommé
                    prob2 = LpProblem("Max_Cout_Stock", LpMaximize)
                    X = {}
                    for cv in all_cvs:
                        safe = cv.replace('|','_').replace(' ','_').replace('-','_')
                        X[cv] = LpVariable(f"X_{safe}", 0, cat=LpInteger)
                    cout_stock = lpSum(
                        X[cv] * sum(q.at[cv, ref] * stock_optim.at[ref, 'Prix (pj)']
                                    for ref in q.columns if ref in stock_optim.index and q.at[cv, ref] > 0)
                        for cv in X
                    )
                    prob2 += cout_stock
                    prob2 += lpSum(X.values()) == max_configs
                    for conf, dem in demand_conf_map.items():
                        cvs_conf = [cv for cv in versions_by_conf[conf] if cv in X]
                        if cvs_conf:
                            prob2 += lpSum(X[cv] for cv in cvs_conf) <= dem
                    for ref in q.columns:
                        coeffs = [(q.at[cv, ref], X[cv]) for cv in X if q.at[cv, ref] > 0]
                        if coeffs:
                            lim = stock_optim.at[ref, 'NVX STOCK'] if ref in stock_optim.index else 0
                            prob2 += lpSum(coef * var for coef, var in coeffs) <= lim
                    prob2.solve(PULP_CBC_CMD(msg=0, timeLimit=1800, gapRel=0.01))

                    produced = {cv: max(0, int(round(value(X[cv].varValue) or 0))) for cv in X}
                    consumed = {}
                    for ref in q.columns:
                        qty = sum(q.at[cv, ref] * produced.get(cv, 0) for cv in all_cvs)
                        if qty > 0:
                            consumed[ref] = round(qty)

                    cout_total = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed.items() if ref in stock_optim.index)

                    total_demande = sum(demand_conf_map.values())
                    total_produit = sum(produced.values())
                    pct = (total_produit / total_demande * 100) if total_demande > 0 else 0
                    val_stock_init = (stock_optim['NVX STOCK'] * stock_optim['Prix (pj)']).sum()
                    val_stock_restant = val_stock_init - cout_total

                    # Stocker les résultats pour affichage persistant
                    st.session_state['j1_sp_result'] = {
                        'demand_conf_map': demand_conf_map,
                        'versions_by_conf': versions_by_conf,
                        'produced': produced,
                        'consumed': consumed,
                        'df_versions': df_versions,
                        'total_demande': total_demande,
                        'total_produit': total_produit,
                        'pct': pct,
                        'cout_total': cout_total,
                        'val_stock_init': float(val_stock_init),
                        'stock_optim': stock_optim,
                    }

                    # Export Excel - 4 onglets
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

                    # ---- Onglet 1 : Demandes (cellules Conf fusionnées) ----
                    ws_dem = wb_sp.active; ws_dem.title = "Demandes"
                    hdrs_dem = ["Constructeur","Conf","Version","Conf|version","Demande (Conf)","Produit (version)","Produit (Conf)","% Réalisation","Restant (Conf)"]
                    for i, h in enumerate(hdrs_dem, 1):
                        ws_dem.cell(1, i, h).font = _BB; ws_dem.cell(1, i).fill = _BLEU
                        ws_dem.cell(1, i).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    rw = 2
                    _merge_dem = [2, 5, 7, 8, 9]  # colonnes Conf-level
                    for conf in sorted(demand_conf_map.keys()):
                        dem = demand_conf_map[conf]
                        cvs = versions_by_conf[conf]
                        prod_c = sum(produced.get(cv, 0) for cv in cvs)
                        rest = dem - prod_c
                        pct_c = round(prod_c / dem * 100, 1) if dem > 0 else 0
                        rs = rw
                        for cv in cvs:
                            qty = produced.get(cv, 0)
                            vm = df_versions[df_versions['Conf|version'] == cv]
                            ver = vm['Version'].iloc[0] if not vm.empty else cv.split('|')[0]
                            cv_cstr = vm['Constructeur'].iloc[0] if not vm.empty else ''
                            ws_dem.cell(rw, 1, cv_cstr); ws_dem.cell(rw, 2, conf)
                            ws_dem.cell(rw, 3, ver); ws_dem.cell(rw, 4, cv)
                            ws_dem.cell(rw, 5, dem).number_format = _NUM
                            ws_dem.cell(rw, 6, qty).number_format = _NUM
                            ws_dem.cell(rw, 7, prod_c).number_format = _NUM
                            ws_dem.cell(rw, 8, f"{pct_c:.1f}%").alignment = Alignment(horizontal="center")
                            ws_dem.cell(rw, 9, rest).number_format = _NUM
                            rw += 1
                        re = rw - 1
                        if re > rs:
                            for col in _merge_dem:
                                ws_dem.merge_cells(start_row=rs, start_column=col, end_row=re, end_column=col)
                                ws_dem.cell(rs, col).alignment = Alignment(horizontal="center", vertical="center")
                        if rest == 0: ws_dem.cell(rs, 9).fill = _VERT
                        elif rest > 0: ws_dem.cell(rs, 9).fill = _ORANGE
                    # Ligne TOTAL
                    gt_dem = int(sum(demand_conf_map.values())); gt_prod = int(sum(produced.values()))
                    ws_dem.cell(rw, 1, "TOTAL")
                    ws_dem.cell(rw, 5, gt_dem); ws_dem.cell(rw, 6, gt_prod); ws_dem.cell(rw, 7, gt_prod)
                    pct_t = round(gt_prod / gt_dem * 100, 1) if gt_dem > 0 else 0
                    ws_dem.cell(rw, 8, f"{pct_t:.1f}%"); ws_dem.cell(rw, 9, gt_dem - gt_prod)
                    _ligne_tot(ws_dem, rw, 9)
                    ws_dem.freeze_panes = "A2"; ws_dem.row_dimensions[1].height = 30
                    ws_dem.auto_filter.ref = f"A1:I{rw}"; _adj_w(ws_dem)

                    # ---- Onglet 2 : BOM ----
                    ws_bom = wb_sp.create_sheet("BOM")
                    bom_exp = bom_df[['Constructeur','Conf|version','Conf','Version','Référence','Quantité']].copy()
                    ref_to_designation = dict(zip(stock_df['Referentiel'], stock_df['Designation']))
                    bom_exp['Designation'] = bom_exp['Référence'].map(ref_to_designation).fillna("")
                    bom_exp['Stock initial'] = bom_exp['Référence'].map(lambda r: stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0)
                    bom_exp['Stock consommable'] = bom_exp.apply(lambda row: produced.get(row['Conf|version'], 0) * row['Quantité'], axis=1)
                    bom_exp['Stock final'] = bom_exp['Référence'].map(lambda r: (stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0) - consumed.get(r, 0))
                    bom_exp = bom_exp.sort_values(['Constructeur','Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_exp, index=False, header=True): ws_bom.append(rr)
                    _fmt_h(ws_bom); _adj_w(ws_bom)
                    col_sf = list(bom_exp.columns).index('Stock final') + 1
                    for rn in range(2, ws_bom.max_row + 1):
                        csf = ws_bom.cell(rn, col_sf)
                        if isinstance(csf.value, (int, float)):
                            if csf.value < 0: csf.fill = _ROUGE
                            elif csf.value == 0: csf.fill = _ORANGE

                    # ---- Onglet 3 : Stock ----
                    ws_stk = wb_sp.create_sheet("Stock")
                    stk_out = stock_df.copy().rename(columns={'Referentiel':'Référence','NVX STOCK':'Stock initial','Prix (pj)':'Prix unitaire'})
                    stk_out['Stock consommable'] = stk_out['Référence'].map(consumed).fillna(0).round(2)
                    stk_out['Stock final'] = (stk_out['Stock initial'] - stk_out['Stock consommable']).round(2)
                    cols_stk = ['Référence','Designation','Stock initial','Prix unitaire','Stock consommable','Stock final']
                    stk_out = stk_out[cols_stk].sort_values('Référence')
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
                    for ci in range(1, 4): ws_res.cell(rr, ci).fill = _BLEU_F
                    rr = 3
                    cout_stk_init = float(val_stock_init)
                    cout_stk_fin = cout_stk_init - cout_total
                    recap_lines = [
                        ("Total de conf demandé", int(total_demande)),
                        ("Total de conf réalisable", int(total_produit)),
                        ("% Réalisation", f"{pct:.1f}%"),
                        ("Restant", int(total_demande - total_produit)),
                        ("Stock initial (EUR)", round(cout_stk_init, 2)),
                        ("Stock consommable (EUR)", round(cout_total, 2)),
                        ("Stock final (EUR)", round(cout_stk_fin, 2)),
                    ]
                    for lbl, val in recap_lines:
                        ws_res.cell(rr, 1, lbl); ws_res.cell(rr, 2, val)
                        if isinstance(val, (int, float)) and "EUR" in lbl: ws_res.cell(rr, 2).number_format = _EUR
                        elif isinstance(val, int): ws_res.cell(rr, 2).number_format = _NUM
                        rr += 1
                    rr += 1
                    ws_res.cell(rr, 1, "DÉTAIL PAR CONF / VERSION").font = _BB14
                    for ci in range(1, 10): ws_res.cell(rr, ci).fill = _BLEU_F
                    rr += 1
                    hdrs_res = ["Constructeur","Conf","Version","Conf|version","Demande (Conf)","Produit (version)","Produit (Conf)","% Réalisation","Restant (Conf)"]
                    for i, h in enumerate(hdrs_res, 1):
                        ws_res.cell(rr, i, h).font = _BN; ws_res.cell(rr, i).fill = _BLEU_C
                        ws_res.cell(rr, i).alignment = Alignment(horizontal="center")
                    rr += 1
                    _merge_res = [2, 5, 7, 8, 9]
                    for conf in sorted(demand_conf_map.keys()):
                        dem = demand_conf_map[conf]
                        cvs = versions_by_conf[conf]
                        prod_c = sum(produced.get(cv, 0) for cv in cvs)
                        rest = dem - prod_c
                        pct_c = (prod_c / dem * 100) if dem > 0 else 0
                        rs = rr
                        for cv in cvs:
                            qty = produced.get(cv, 0)
                            vm = df_versions[df_versions['Conf|version'] == cv]
                            ver = vm['Version'].iloc[0] if not vm.empty else cv.split('|')[0]
                            cv_cstr = vm['Constructeur'].iloc[0] if not vm.empty else ''
                            ws_res.cell(rr, 1, cv_cstr)
                            ws_res.cell(rr, 2, conf)
                            ws_res.cell(rr, 3, ver)
                            ws_res.cell(rr, 4, cv)
                            ws_res.cell(rr, 5, dem).number_format = _NUM
                            ws_res.cell(rr, 6, qty).number_format = _NUM
                            ws_res.cell(rr, 7, prod_c).number_format = _NUM
                            ws_res.cell(rr, 8, f"{pct_c:.1f}%").alignment = Alignment(horizontal="center")
                            ws_res.cell(rr, 9, rest).number_format = _NUM
                            rr += 1
                        re = rr - 1
                        if re > rs:
                            for col in _merge_res:
                                ws_res.merge_cells(start_row=rs, start_column=col, end_row=re, end_column=col)
                                ws_res.cell(rs, col).alignment = Alignment(horizontal="center", vertical="center")
                        if rest == 0: ws_res.cell(rs, 9).fill = _VERT
                        elif rest > 0: ws_res.cell(rs, 9).fill = _ORANGE
                    for cl, w in [('A',16),('B',14),('C',12),('D',22),('E',14),('F',16),('G',14),('H',14),('I',14)]: ws_res.column_dimensions[cl].width = w

                    out_sp = BytesIO(); wb_sp.save(out_sp)
                    ts_sp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state['j1_sp_excel'] = out_sp.getvalue()
                    st.session_state['j1_sp_excel_name'] = f"Optimisation_Reemploi_{ts_sp}.xlsx"

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        # -------- Affichage persistant des résultats --------
        if 'j1_sp_result' in st.session_state:
            _r = st.session_state['j1_sp_result']
            _dcm = _r['demand_conf_map']
            _vbc = _r['versions_by_conf']
            _prod = _r['produced']
            _cons = _r['consumed']
            _dfv = _r['df_versions']
            _td = _r['total_demande']
            _tp = _r['total_produit']
            _pct = _r['pct']
            _ct = _r['cout_total']
            _sopt = _r['stock_optim']

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="kpi-label">Résultats de l\'optimisation</div>', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Demande totale de configuration", f"{int(_td)}")
            k2.metric("Configurations réalisables", f"{int(_tp)}")
            k3.metric("Taux de réalisation", f"{_pct:.1f}%")
            k4.metric("Valeur stock consommable", f"{_ct:,.0f} €".replace(',', ' '))

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="kpi-label">Détail par Conf</div>', unsafe_allow_html=True)

            rows_disp = []
            conf_summ = []
            for conf in sorted(_dcm.keys()):
                dem = _dcm[conf]; cvs = _vbc[conf]
                prod_c = sum(_prod.get(cv, 0) for cv in cvs)
                rest = dem - prod_c
                pct_c = round(prod_c / dem * 100, 1) if dem > 0 else 0
                m = _dfv[_dfv['Conf'] == conf]
                cstr = m['Constructeur'].iloc[0] if not m.empty else '?'
                conf_summ.append({'Constructeur': cstr, 'Demande': dem, 'Produit': prod_c})
                for cv in cvs:
                    qty = _prod.get(cv, 0)
                    vm = _dfv[_dfv['Conf|version'] == cv]
                    ver = vm['Version'].iloc[0] if not vm.empty else cv.split('|')[0]
                    cv_c = vm['Constructeur'].iloc[0] if not vm.empty else ''
                    rows_disp.append({
                        'Constructeur': cv_c, 'Conf': conf, 'Version': ver,
                        'Conf|version': cv, 'Demande (Conf)': dem, 'Produit': qty,
                        '% Réalisé (Conf)': pct_c, 'Restant (Conf)': rest,
                    })
            detail_df = pd.DataFrame(rows_disp)
            st.dataframe(detail_df, use_container_width=True, height=450, hide_index=True)

            st.markdown('<div class="kpi-label">Répartition par constructeur</div>', unsafe_allow_html=True)
            sdf = pd.DataFrame(conf_summ)
            cdf = sdf.groupby('Constructeur').agg(Demande=('Demande', 'sum'), Produit=('Produit', 'sum')).reset_index()
            cdf['Restant'] = cdf['Demande'] - cdf['Produit']
            cdf['Taux %'] = (cdf['Produit'] / cdf['Demande'] * 100).round(1).astype(str) + '%'
            cdf = cdf.rename(columns={'Demande': 'Configurations demandées', 'Produit': 'Configuration réalisable'})
            st.dataframe(cdf, use_container_width=True, hide_index=True)

            # -------- Dashboard : détail des références par configuration --------
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="kpi-label">Détail des références par configuration</div>', unsafe_allow_html=True)

            conf_opts = sorted(_dcm.keys())
            sel_conf = st.selectbox("Configuration", conf_opts, key="sel_conf_sp")

            if sel_conf:
                dem_s = _dcm[sel_conf]; cvs_s = _vbc[sel_conf]
                prod_s = sum(_prod.get(cv, 0) for cv in cvs_s)
                rest_s = dem_s - prod_s
                pct_s = round(prod_s / dem_s * 100, 1) if dem_s > 0 else 0

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Demande", f"{dem_s}")
                m2.metric("Produit", f"{prod_s}")
                m3.metric("% Réalisation", f"{pct_s}%")
                m4.metric("Restant", f"{rest_s}")

                # Versions produites
                st.markdown('<div class="kpi-label">Versions produites</div>', unsafe_allow_html=True)
                vr = []
                for cv in cvs_s:
                    qty = _prod.get(cv, 0)
                    vm = _dfv[_dfv['Conf|version'] == cv]
                    ver = vm['Version'].iloc[0] if not vm.empty else cv.split('|')[0]
                    cv_c = vm['Constructeur'].iloc[0] if not vm.empty else ''
                    vr.append({'Constructeur': cv_c, 'Version': ver, 'Conf|version': cv, 'Produit': qty})
                st.dataframe(pd.DataFrame(vr), use_container_width=True, hide_index=True)

                # Références consommées
                _bom = st.session_state['bom_df']
                ref_rows = []
                for cv in cvs_s:
                    qty_cv = _prod.get(cv, 0)
                    if qty_cv > 0:
                        bom_cv = _bom[_bom['Conf|version'] == cv]
                        for _, br in bom_cv.iterrows():
                            ref = br['Référence']; bom_qty = br['Quantité']
                            cons_qty = bom_qty * qty_cv
                            prix = float(_sopt.at[ref, 'Prix (pj)']) if ref in _sopt.index else 0
                            desig = _sopt.at[ref, 'Designation'] if ref in _sopt.index and 'Designation' in _sopt.columns else ''
                            stk = float(_sopt.at[ref, 'NVX STOCK']) if ref in _sopt.index else 0
                            ref_rows.append({
                                'Référence': ref, 'Designation': desig,
                                'BOM/unité': int(bom_qty), 'Qté consommée': int(cons_qty),
                                'Stock initial': int(stk), 'Prix unitaire': round(prix, 2),
                                'Coût': round(cons_qty * prix, 2),
                            })
                if ref_rows:
                    st.markdown('<div class="kpi-label">Références consommées</div>', unsafe_allow_html=True)
                    rdf = pd.DataFrame(ref_rows)
                    ragg = rdf.groupby(['Référence', 'Designation']).agg({
                        'Qté consommée': 'sum', 'Stock initial': 'first', 'Prix unitaire': 'first',
                    }).reset_index()
                    ragg['Coût'] = ragg['Qté consommée'] * ragg['Prix unitaire']
                    ragg = ragg.sort_values('Coût', ascending=False)
                    st.dataframe(ragg, use_container_width=True, hide_index=True)
                    tot_cout_ref = ragg['Coût'].sum()
                    st.markdown(f"**Coût total des références consommées : {tot_cout_ref:,.0f} EUR**".replace(',', ' '))
                else:
                    st.info("Aucune référence consommée pour cette configuration (0 produit)")

        if 'j1_sp_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="TÉLÉCHARGER LES RÉSULTATS",
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
    st.markdown('<div class="section-title">Optimisation du réemploi avec priorité</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Maximise le nombre total de configurations avec le stock disponible, tout en respectant l\'ordre des priorités.</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode séquentielle en 2 étapes</strong><br>
        1. Chaque niveau de priorité est traité indépendamment.<br>
        2. Le stock est consommé progressivement de la priorité la plus haute à la plus basse.
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Chargement des données")
    elif not prio_file:
        st.warning("Chargez le fichier Prio.xlsx dans l'onglet Chargement des données")
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
                    cols_utiles = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Priorité', 'Demande', 'Next']
                    df_prio = df_prio[[c for c in cols_utiles if c in df_prio.columns]].copy()

                    # Construire next_map avant filtrage
                    next_map = {}
                    if 'Next' in df_prio.columns:
                        for _, row in df_prio.dropna(subset=['Next']).iterrows():
                            nxt = str(row['Next']).strip()
                            if nxt and nxt != 'nan' and nxt != row['Conf|version']:
                                next_map[row['Conf|version']] = nxt

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

                    produced_j1, consumed_j1, cout_j1, stock_after_j1, stock_optim, produced_next_j1 = run_jalon1(
                        all_demands, bom_df, stock_df, q, priorities, prix_file.getvalue(), next_map
                    )

                    all_demands['J1 produit'] = all_demands.apply(
                        lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                    )
                    all_demands['J1 Next'] = all_demands.apply(
                        lambda row: produced_next_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                    )
                    all_demands['J1 total'] = all_demands['J1 produit'] + all_demands['J1 Next']
                    all_demands['BOM Next'] = all_demands.apply(
                        lambda row: next_map.get(row['Conf|version'], '') if produced_next_j1.get((row['Conf|version'], row['Priorité']), 0) > 0 else '', axis=1
                    )
                    all_demands['Restant J1'] = all_demands['Demande'] - all_demands['J1 total']
                    total_demande = all_demands['Demande'].sum()
                    total_j1 = all_demands['J1 total'].sum()
                    total_j1_next = all_demands['J1 Next'].sum()
                    pct_j1 = (total_j1 / total_demande * 100) if total_demande > 0 else 0
                    nb_en_rupture = (all_demands['Restant J1'] > 0).sum()
                    val_rupture = (all_demands['Restant J1'] * all_demands['Prix'].fillna(0)).sum()

                    st.markdown("<br>", unsafe_allow_html=True)

                    st.markdown('<div class="kpi-label">Résultats de l\'optimisation</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale de configuration", f"{int(total_demande)}")
                    _label_real = f"Configurations réalisables (dont {int(total_j1_next)} via substitution)" if total_j1_next > 0 else "Configurations réalisables"
                    c2.metric(_label_real, f"{int(total_j1)}")
                    c3.metric("Taux de réalisation", f"{pct_j1:.1f}%")
                    c4.metric("Valeur stock consommable", f"{cout_j1:,.0f} €".replace(',', ' '))

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div class="kpi-label">Détail par configuration</div>', unsafe_allow_html=True)
                    det_cols3 = ['Constructeur', 'Conf|version', 'Priorité', 'Demande', 'J1 produit']
                    if total_j1_next > 0:
                        det_cols3 += ['J1 Next', 'BOM Next', 'J1 total']
                    det_cols3.append('Restant J1')
                    detail_df = all_demands[det_cols3].copy()
                    rename3 = {'Demande': 'Configuration demandée', 'J1 produit': 'Réalisable (BOM originale)', 'Restant J1': 'Restant'}
                    if total_j1_next > 0:
                        rename3['J1 Next'] = 'Réalisable (substitution)'
                        rename3['BOM Next'] = 'BOM substitution'
                        rename3['J1 total'] = 'Configuration réalisable'
                        detail_df = detail_df.rename(columns=rename3)
                        detail_df['% Réalisé'] = (detail_df['Configuration réalisable'] / detail_df['Configuration demandée'] * 100).round(1)
                    else:
                        rename3['J1 produit'] = 'Configuration réalisable'
                        detail_df = detail_df.rename(columns=rename3)
                        detail_df['% Réalisé'] = (detail_df['Configuration réalisable'] / detail_df['Configuration demandée'] * 100).round(1)
                    st.dataframe(detail_df, use_container_width=True, height=350, hide_index=True)

                    if 'Constructeur' in all_demands.columns:
                        st.markdown('<div class="kpi-label">Répartition par constructeur</div>', unsafe_allow_html=True)
                        constr = all_demands.groupby('Constructeur').agg(
                            Demande=('Demande', 'sum'), Produit=('J1 total', 'sum')
                        ).reset_index()
                        constr['Restant'] = constr['Demande'] - constr['Produit']
                        constr['Taux %'] = (constr['Produit'] / constr['Demande'] * 100).round(1).astype(str) + '%'
                        constr = constr.rename(columns={
                            'Demande': 'Configurations demandées',
                            'Produit': 'Configuration réalisable'
                        })
                        st.dataframe(constr, use_container_width=True, hide_index=True)

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
                    df_ap_exp = all_demands.copy()
                    df_ap_exp['Qté produite'] = df_ap_exp['J1 total']
                    df_ap_exp['Restant'] = df_ap_exp['Restant J1']

                    # ---- Onglet 1 : Demandes ----
                    ws_dem2 = wb_ap.active; ws_dem2.title = "Demandes"
                    cols_exp2 = ['Constructeur','Conf','Version','Priorité','Demande','J1 produit']
                    if total_j1_next > 0:
                        cols_exp2 += ['J1 Next', 'BOM Next', 'J1 total']
                    cols_exp2.append('Restant J1')
                    cols_exp2 = [c for c in cols_exp2 if c in df_ap_exp.columns]
                    df_exp2 = df_ap_exp[cols_exp2].copy()
                    rename_dem2 = {'Demande': 'Demande de configuration', 'J1 produit': 'Nb réalisable (BOM originale)', 'Restant J1': 'Restant'}
                    if total_j1_next > 0:
                        rename_dem2['J1 Next'] = 'Nb réalisable (substitution)'
                        rename_dem2['BOM Next'] = 'BOM substitution'
                        rename_dem2['J1 total'] = 'Nombre de configuration réalisable'
                    else:
                        rename_dem2['J1 produit'] = 'Nombre de configuration réalisable'
                    df_exp2 = df_exp2.rename(columns=rename_dem2)
                    df_exp2 = df_exp2.sort_values(['Constructeur','Priorité'] if 'Constructeur' in df_exp2.columns else ['Priorité'])
                    for rr in dataframe_to_rows(df_exp2, index=False, header=True): ws_dem2.append(rr)
                    tot_row2 = ws_dem2.max_row + 1
                    ws_dem2.cell(tot_row2, 1, "TOTAL")
                    cidx2 = {list(df_exp2.columns)[i]: i+1 for i in range(len(df_exp2.columns))}
                    for cn in ['Demande de configuration','Nombre de configuration réalisable','Restant']:
                        if cn in cidx2: ws_dem2.cell(tot_row2, cidx2[cn], df_exp2[cn].sum())
                    _ligne_tot2(ws_dem2, tot_row2, len(df_exp2.columns))
                    _fmt_h2(ws_dem2); _adj_w2(ws_dem2)
                    for row in ws_dem2.iter_rows(min_row=2, max_row=ws_dem2.max_row):
                        for cell in row:
                            h = ws_dem2.cell(1, cell.column).value
                            if h in ['Demande de configuration','Nombre de configuration réalisable','Restant'] and isinstance(cell.value, (int, float)): cell.number_format = _NUM2
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
                    bom_exp2 = bom_df[['Constructeur','Conf|version','Conf','Version','Référence','Quantité']].copy()
                    ref_to_designation = dict(zip(stock_df['Referentiel'], stock_df['Designation']))
                    bom_exp2['Designation'] = bom_exp2['Référence'].map(ref_to_designation).fillna("")
                    bom_exp2['Stock initial'] = bom_exp2['Référence'].map(lambda r: stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0)
                    bom_exp2['Stock consommable'] = bom_exp2.apply(lambda row: produced_total_j1.get(row['Conf|version'], 0) * row['Quantité'], axis=1)
                    bom_exp2['Stock final'] = bom_exp2['Référence'].map(lambda r: (stock_optim.at[r,'NVX STOCK'] if r in stock_optim.index else 0) - consumed_j1.get(r, 0))
                    bom_exp2 = bom_exp2.sort_values(['Constructeur','Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_exp2, index=False, header=True): ws_bom2.append(rr)
                    _fmt_h2(ws_bom2); _adj_w2(ws_bom2)
                    col_sf2b = list(bom_exp2.columns).index('Stock final') + 1
                    for rn in range(2, ws_bom2.max_row + 1):
                        csf = ws_bom2.cell(rn, col_sf2b)
                        if isinstance(csf.value, (int, float)):
                            if csf.value < 0: csf.fill = _ROUGE2
                            elif csf.value == 0: csf.fill = _ORANGE2

                    # ---- Onglet 3 : Stock ----
                    ws_stk2 = wb_ap.create_sheet("Stock")
                    stk_out2 = stock_df.copy().rename(columns={'Referentiel':'Référence','NVX STOCK':'Stock initial','Prix (pj)':'Prix unitaire'})
                    stk_out2['Stock consommable'] = stk_out2['Référence'].map(consumed_j1).fillna(0).round(2)
                    stk_out2['Stock final'] = (stk_out2['Stock initial'] - stk_out2['Stock consommable']).round(2)
                    cols_stk2 = ['Référence','Designation','Stock initial','Prix unitaire','Stock consommable','Stock final']
                    stk_out2 = stk_out2[cols_stk2].sort_values('Référence')
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
                    for ci in range(1, 4): ws_res2.cell(rr2, ci).fill = _BLEU_F2
                    rr2 = 3
                    cout_stk_init2 = float((stock_optim['NVX STOCK'] * stock_optim['Prix (pj)']).sum())
                    cout_stk_fin2 = cout_stk_init2 - cout_j1
                    recap_lines2 = [
                        ("Total de conf demandé", int(total_demande)),
                        ("Total de conf réalisable", int(total_j1)),
                    ]
                    if total_j1_next > 0:
                        recap_lines2.append(("  dont via substitution (Next)", int(total_j1_next)))
                    recap_lines2 += [
                        ("% Réalisation", f"{pct_j1:.1f}%"),
                        ("Restant", int(total_demande - total_j1)),
                        ("Stock initial (EUR)", round(cout_stk_init2, 2)),
                        ("Stock consommable", round(cout_j1, 2)),
                        ("Stock final (EUR)", round(cout_stk_fin2, 2)),
                    ]
                    for lbl, val in recap_lines2:
                        ws_res2.cell(rr2, 1, lbl); ws_res2.cell(rr2, 2, val)
                        if isinstance(val, (int, float)) and "EUR" in lbl: ws_res2.cell(rr2, 2).number_format = _EUR2
                        elif isinstance(val, int): ws_res2.cell(rr2, 2).number_format = _NUM2
                        rr2 += 1
                    rr2 += 1
                    ws_res2.cell(rr2, 1, "DÉTAIL PAR CONSTRUCTEUR / CONF").font = _BB142
                    for ci in range(1, 6): ws_res2.cell(rr2, ci).fill = _BLEU_F2
                    rr2 += 1
                    for i, h in enumerate(["Constructeur","Conf","Demande de configuration","Nombre de configuration réalisable","Réalisation"], 1):
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
                    st.session_state['j1_ap_excel_name'] = f"Optimisation_Reemploi_Prio_{ts_ap}.xlsx"

                    st.session_state['j1_avec_result'] = {
                        'all_demands': all_demands,
                        'stock_optim': stock_optim,
                        'produced_j1': produced_j1,
                        'produced_next_j1': produced_next_j1,
                        'consumed_j1': consumed_j1,
                        'cout_j1': cout_j1,
                        'q': q,
                        'priorities': priorities,
                        'stock_after_j1': stock_after_j1,
                        'next_map': next_map,
                    }

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        if 'j1_ap_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="TÉLÉCHARGER LES RÉSULTATS",
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
    st.markdown('<div class="section-title">Optimisation du réemploi avec priorité et achat de références unitaires</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Stock existant et achat de références manquantes</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-block">
        <strong style="color:#FFFFFF;">Méthode en 2 étapes</strong><br>
        1. Maximisation avec le stock disponible<br>
        2. Achat de références par seuils de coût de configuration neuve
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get('data_prepared'):
        st.warning("Préparez d'abord les données dans l'onglet Chargement des données")
    elif not prio_file:
        st.warning("Chargez le fichier Prio.xlsx dans l'onglet Chargement des données")
    else:
        has_j1 = 'j1_avec_result' in st.session_state

        if st.button("Lancer l'optimisation", key="run_j2", type="primary"):
            with st.spinner("Optimisation en cours..."):
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
                    cols_utiles = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Priorité', 'Demande', 'Next']
                    df_prio = df_prio[[c for c in cols_utiles if c in df_prio.columns]].copy()

                    # Construire next_map avant filtrage
                    next_map = {}
                    if 'Next' in df_prio.columns:
                        for _, row in df_prio.dropna(subset=['Next']).iterrows():
                            nxt = str(row['Next']).strip()
                            if nxt and nxt != 'nan' and nxt != row['Conf|version']:
                                next_map[row['Conf|version']] = nxt

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
                        produced_next_j1 = j1_data.get('produced_next_j1', {})
                        consumed_j1 = j1_data['consumed_j1']
                        cout_j1 = j1_data['cout_j1']
                        stock_j2 = j1_data['stock_after_j1'].copy()
                        stock_optim = j1_data['stock_optim']
                        next_map_j1 = j1_data.get('next_map', next_map)
                        all_demands['J1 produit'] = all_demands.apply(
                            lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )
                        all_demands['J1 Next'] = all_demands.apply(
                            lambda row: produced_next_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )
                    else:
                        produced_j1, consumed_j1, cout_j1, stock_j2, stock_optim, produced_next_j1 = run_jalon1(
                            all_demands, bom_df, stock_df, q, priorities, prix_file.getvalue(), next_map
                        )
                        all_demands['J1 produit'] = all_demands.apply(
                            lambda row: produced_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )
                        all_demands['J1 Next'] = all_demands.apply(
                            lambda row: produced_next_j1.get((row['Conf|version'], row['Priorité']), 0), axis=1
                        )

                    all_demands['J1 total'] = all_demands['J1 produit'] + all_demands['J1 Next']
                    all_demands['BOM Next J1'] = all_demands.apply(
                        lambda row: next_map.get(row['Conf|version'], '') if produced_next_j1.get((row['Conf|version'], row['Priorité']), 0) > 0 else '', axis=1
                    )
                    all_demands['Restant J1'] = all_demands['Demande'] - all_demands['J1 total']
                    total_demande = all_demands['Demande'].sum()
                    total_j1 = all_demands['J1 total'].sum()
                    total_j1_next = all_demands['J1 Next'].sum()
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
                            progress_bar.progress(min(0.99, step / total_steps))

                            configs_prio = sorted(set(cv for (cv, p) in restant if p == prio and restant[(cv, p)] > 0 and cv in q.index and cv in prix_configs))
                            sans_prix_prio = sorted(set(cv for (cv, p) in restant if p == prio and restant[(cv, p)] > 0 and cv in q.index and cv not in prix_configs))

                            if not configs_prio:
                                for cv in sans_prix_prio:
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': restant.get((cv, prio), 0),
                                                      'Prix conf': None, 'Cout achat': None, 'Cout moyen/conf': None,
                                                      '% du prix': None, 'Statut': 'Pas de prix'})
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
                                    cout_refs_necessaire = 0
                                    for ref, bom_qty in bom_par_config[cv].items():
                                        stock_dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                        qty_a_acheter = max(0, bom_qty - stock_dispo)
                                        cout_refs_necessaire += qty_a_acheter * ref_to_prix.get(ref, 0)
                                    pct_cout = (cout_refs_necessaire / prix_configs[cv] * 100) if prix_configs[cv] > 0 else 0
                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': demand_prio[cv],
                                                      'Prix conf': prix_configs[cv], 'Cout achat': round(cout_refs_necessaire, 2),
                                                      'Cout moyen/conf': round(cout_refs_necessaire, 2),
                                                      '% du prix': round(pct_cout, 1), 'Statut': 'Bloque'})
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
                                                      'Cout moyen/conf': round(cout_achat_cv / qty, 2),
                                                      '% du prix': round(pct_moy, 1), 'Statut': 'Produit'})
                                else:
                                    # Calculer le coût minimum des refs pour faire 1 config (même si bloqué)
                                    cout_refs_necessaire = 0
                                    for ref, bom_qty in bom_par_config[cv].items():
                                        stock_dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                        prix_ref = ref_to_prix.get(ref, 0)
                                        qty_a_acheter = max(0, bom_qty - stock_dispo)
                                        cout_refs_necessaire += qty_a_acheter * prix_ref
                                    pct_cout = (cout_refs_necessaire / prix_configs[cv] * 100) if prix_configs[cv] > 0 else 0

                                    detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                      'Nb produit': 0, 'Demande restante': demand_prio[cv],
                                                      'Prix conf': prix_configs[cv], 'Cout achat': round(cout_refs_necessaire, 2),
                                                      'Cout moyen/conf': round(cout_refs_necessaire, 2),
                                                      '% du prix': round(pct_cout, 1), 'Statut': 'Bloque'})
                            # --- Substitution Next pour J2 (seuil/prio) ---
                            if next_map:
                                next_j2_cand = {}
                                for cv in configs_prio:
                                    rem = restant.get((cv, prio), 0)
                                    if rem > 0 and cv in next_map:
                                        ncv = next_map[cv]
                                        if ncv in q.index and cv in prix_configs:
                                            next_j2_cand[cv] = (ncv, rem)
                                if next_j2_cand:
                                    next_refs_j2 = sorted(set(ref for cv, (ncv, _) in next_j2_cand.items() for ref in q.columns if q.at[ncv, ref] > 0))
                                    next_bom_j2 = {cv: {ref: q.at[ncv, ref] for ref in q.columns if q.at[ncv, ref] > 0} for cv, (ncv, _) in next_j2_cand.items()}
                                    prob_nj1 = LpProblem(f"NextJ2_S{seuil_pct}_P{prio}_Max", LpMaximize)
                                    XNv, SNv, BNv = {}, {}, {}
                                    for cv, (ncv, rem) in next_j2_cand.items():
                                        XNv[cv] = LpVariable(f"XN_{sname(cv)}_s{seuil_pct}", 0, rem, LpInteger)
                                        for ref, bom_qty in next_bom_j2[cv].items():
                                            SNv[(cv, ref)] = LpVariable(f"SN_{sname(cv)}_{sname(ref)}_s{seuil_pct}", 0)
                                            BNv[(cv, ref)] = LpVariable(f"BN_{sname(cv)}_{sname(ref)}_s{seuil_pct}", 0)
                                            prob_nj1 += SNv[(cv, ref)] + BNv[(cv, ref)] == bom_qty * XNv[cv]
                                    prob_nj1 += lpSum(XNv.values())
                                    for ref in next_refs_j2:
                                        s_vars = [SNv[(cv, ref)] for cv in next_j2_cand if (cv, ref) in SNv]
                                        if s_vars:
                                            dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                            prob_nj1 += lpSum(s_vars) <= dispo
                                    for cv in next_j2_cand:
                                        budget_cv = seuil * prix_configs[cv]
                                        cout_cv = lpSum(BNv[(cv, ref)] * ref_to_prix.get(ref, 0) for ref in next_bom_j2[cv] if (cv, ref) in BNv)
                                        prob_nj1 += cout_cv <= budget_cv * XNv[cv]
                                    prob_nj1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
                                    max_n_next = int(round(value(prob_nj1.objective) or 0))
                                    if max_n_next > 0:
                                        prob_nj2 = LpProblem(f"NextJ2_S{seuil_pct}_P{prio}_Min", LpMinimize)
                                        XN2, SN2, BN2 = {}, {}, {}
                                        for cv, (ncv, rem) in next_j2_cand.items():
                                            XN2[cv] = LpVariable(f"XN2_{sname(cv)}_s{seuil_pct}", 0, rem, LpInteger)
                                            for ref, bom_qty in next_bom_j2[cv].items():
                                                SN2[(cv, ref)] = LpVariable(f"SN2_{sname(cv)}_{sname(ref)}_s{seuil_pct}", 0)
                                                BN2[(cv, ref)] = LpVariable(f"BN2_{sname(cv)}_{sname(ref)}_s{seuil_pct}", 0)
                                                prob_nj2 += SN2[(cv, ref)] + BN2[(cv, ref)] == bom_qty * XN2[cv]
                                        prob_nj2 += lpSum(BN2[(cv, ref)] * ref_to_prix.get(ref, 0)
                                                         for cv in next_j2_cand for ref in next_bom_j2[cv] if (cv, ref) in BN2)
                                        prob_nj2 += lpSum(XN2.values()) == max_n_next
                                        for ref in next_refs_j2:
                                            s_vars = [SN2[(cv, ref)] for cv in next_j2_cand if (cv, ref) in SN2]
                                            if s_vars:
                                                dispo = max(0, float(stock_j2[ref])) if ref in stock_j2.index else 0
                                                prob_nj2 += lpSum(s_vars) <= dispo
                                        for cv in next_j2_cand:
                                            budget_cv = seuil * prix_configs[cv]
                                            cout_cv = lpSum(BN2[(cv, ref)] * ref_to_prix.get(ref, 0) for ref in next_bom_j2[cv] if (cv, ref) in BN2)
                                            prob_nj2 += cout_cv <= budget_cv * XN2[cv]
                                        prob_nj2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))
                                        for cv in next_j2_cand:
                                            qty_n = int(round(lp_val(XN2[cv])))
                                            if qty_n > 0:
                                                cout_achat_n = sum(lp_val(BN2[(cv, ref)]) * ref_to_prix.get(ref, 0) for ref in next_bom_j2[cv] if (cv, ref) in BN2)
                                                pct_moy_n = (cout_achat_n / (prix_configs[cv] * qty_n) * 100) if prix_configs[cv] > 0 else 0
                                                restant[(cv, prio)] -= qty_n
                                                total_seuil += qty_n
                                                cout_seuil += cout_achat_n
                                                for ref in next_bom_j2[cv]:
                                                    if (cv, ref) in SN2:
                                                        used = round(lp_val(SN2[(cv, ref)]), 4)
                                                        if used > 0 and ref in stock_j2.index:
                                                            stock_j2[ref] -= used
                                                            consumed_j2[ref] = consumed_j2.get(ref, 0) + used
                                                    if (cv, ref) in BN2:
                                                        bought = round(lp_val(BN2[(cv, ref)]), 4)
                                                        if bought > 0:
                                                            achats_j2[ref] = achats_j2.get(ref, 0) + bought
                                                ncv = next_j2_cand[cv][0]
                                                detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                                  'Nb produit': qty_n, 'Demande restante': next_j2_cand[cv][1],
                                                                  'Prix conf': prix_configs[cv], 'Cout achat': round(cout_achat_n, 2),
                                                                  'Cout moyen/conf': round(cout_achat_n / qty_n, 2),
                                                                  '% du prix': round(pct_moy_n, 1), 'Statut': f'Produit (Next: {ncv})'})

                            for cv in sans_prix_prio:
                                detail_j2.append({'Seuil': f"{seuil_pct}%", 'Priorité': prio, 'Conf|version': cv,
                                                  'Nb produit': 0, 'Demande restante': restant.get((cv, prio), 0),
                                                  'Prix conf': None, 'Cout achat': None, 'Cout moyen/conf': None,
                                                  '% du prix': None, 'Statut': 'Pas de prix'})

                        j2_prod_par_seuil[seuil_pct] = total_seuil
                        j2_cout_par_seuil[seuil_pct] = cout_seuil

                    progress_bar.progress(1.0)

                    total_j2 = sum(j2_prod_par_seuil.values())
                    total_produit = total_j1 + total_j2
                    pct_final = (total_produit / total_demande * 100) if total_demande > 0 else 0
                    restant_final = total_demande - total_produit
                    cout_achats_j2 = sum(d['Cout achat'] for d in detail_j2 if d['Cout achat'] is not None)
                    cout_conf_restantes = sum(restant.get((cv, p), 0) * prix_configs.get(cv, 0)
                                             for (cv, p) in restant if restant[(cv, p)] > 0)
                    cout_total_estime = cout_j1 + cout_achats_j2 + cout_conf_restantes

                    st.markdown("<br>", unsafe_allow_html=True)

                    st.markdown('<div class="kpi-label">Bilan global</div>', unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Demande totale de configuration", f"{int(total_demande)}")
                    c2.metric("Total de configurations réalisables", f"{int(total_produit)}")
                    c3.metric("Taux de réalisation final", f"{pct_final:.1f}%")
                    c4.metric("Configurations non réalisables", f"{int(restant_final)}")

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown('<div class="kpi-label">Configuration réalisable par seuil</div>', unsafe_allow_html=True)

                    # Compter les substitutions Next dans J2 par seuil
                    next_j2_par_seuil = {}
                    next_j2_details = []
                    for d in detail_j2:
                        if d['Statut'].startswith('Produit (Next:'):
                            s_val = d['Seuil'].replace('%', '')
                            s_int = int(s_val)
                            next_j2_par_seuil[s_int] = next_j2_par_seuil.get(s_int, 0) + d['Nb produit']
                            # Extraire le BOM Next du statut
                            bom_next = d['Statut'].replace('Produit (Next: ', '').rstrip(')')
                            next_j2_details.append({
                                'Conf|version': d['Conf|version'],
                                'BOM substitution': bom_next,
                                'Seuil': d['Seuil'],
                                'Priorité': d['Priorité'],
                                'Nb produit': d['Nb produit'],
                                'Coût achat': d['Cout achat'],
                            })
                    total_next_j2 = sum(next_j2_par_seuil.values())

                    cumul = total_j1
                    seuil_rows = [{'Etape': 'Opti du réemploi', 'Config réalisable': int(total_j1),
                                   'Coût stock consommable': f"{cout_j1:,.0f}".replace(',', ' '),
                                   'Coût achat réf. unitaire': '0',
                                   'Cumul des conf réalisable': int(total_j1),
                                   'Taux réalisable': f"{pct_j1:.1f}%"}]
                    if total_j1_next > 0:
                        seuil_rows[0]['Etape'] = f"Opti du réemploi (dont {int(total_j1_next)} via Next)"
                    for s_pct in [20, 40, 60, 80]:
                        prod = j2_prod_par_seuil.get(s_pct, 0)
                        cout = j2_cout_par_seuil.get(s_pct, 0)
                        cumul += prod
                        etape_label = f'Seuil {s_pct}%'
                        next_count = next_j2_par_seuil.get(s_pct, 0)
                        if next_count > 0:
                            etape_label = f'Seuil {s_pct}% (dont {int(next_count)} via Next)'
                        seuil_rows.append({'Etape': etape_label, 'Config réalisable': int(prod),
                                           'Coût stock consommable': '0',
                                           'Coût achat réf. unitaire': f"{cout:,.0f}".replace(',', ' '),
                                           'Cumul des conf réalisable': int(cumul),
                                           'Taux réalisable': f"{(cumul/total_demande*100):.1f}%"})
                    st.dataframe(pd.DataFrame(seuil_rows), use_container_width=True, hide_index=True)

                    # ---- DETAIL DES SUBSTITUTIONS (NEXT) ----
                    total_next_all = int(total_j1_next) + total_next_j2
                    if total_next_all > 0:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(f'<div class="kpi-label">Configurations produites via substitution (Next) — {int(total_next_all)} au total</div>', unsafe_allow_html=True)

                        # J1 Next
                        next_j1_rows = []
                        if total_j1_next > 0:
                            for _, row in all_demands.iterrows():
                                if row['J1 Next'] > 0:
                                    next_j1_rows.append({
                                        'Conf|version': row['Conf|version'],
                                        'BOM substitution': row['BOM Next J1'],
                                        'Etape': 'Réemploi (J1)',
                                        'Priorité': row['Priorité'],
                                        'Nb produit': int(row['J1 Next']),
                                        'Coût achat': 0,
                                    })

                        all_next_rows = next_j1_rows + next_j2_details
                        df_next = pd.DataFrame(all_next_rows)
                        df_next = df_next.rename(columns={'Nb produit': 'Qté produite', 'Coût achat': 'Coût achat (€)'})
                        if 'Seuil' in df_next.columns:
                            df_next['Etape'] = df_next['Etape'].fillna('')
                            mask_j2 = df_next['Etape'] == ''
                            df_next.loc[mask_j2, 'Etape'] = 'Seuil ' + df_next.loc[mask_j2, 'Seuil'].astype(str)
                            df_next = df_next.drop(columns=['Seuil'])
                        display_cols = ['Conf|version', 'BOM substitution', 'Etape', 'Priorité', 'Qté produite', 'Coût achat (€)']
                        display_cols = [c for c in display_cols if c in df_next.columns]
                        st.dataframe(df_next[display_cols], use_container_width=True, hide_index=True)

                    # ---- RÉPARTITION PAR CONSTRUCTEUR ----
                    if 'Constructeur' in all_demands.columns:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown('<div class="kpi-label">Répartition par constructeur</div>', unsafe_allow_html=True)

                        # Calculer J1 par constructeur
                        j1_constr = all_demands.groupby('Constructeur').agg(
                            Demande_totale=('Demande', 'sum'),
                            J1_produit=('J1 total', 'sum')
                        ).reset_index()

                        # Calculer J2 par constructeur depuis detail_j2
                        if detail_j2:
                            df_j2_all = pd.DataFrame(detail_j2)
                            df_j2_all = df_j2_all.merge(
                                all_demands[['Conf|version','Priorité','Constructeur']].drop_duplicates(),
                                on=['Conf|version','Priorité'], how='left'
                            )
                            constr_j2 = df_j2_all.groupby('Constructeur').agg(
                                Produit_J2=('Nb produit', 'sum')
                            ).reset_index()

                            # Fusionner J1 et J2
                            constr_summary = j1_constr.merge(constr_j2, on='Constructeur', how='left')
                        else:
                            constr_summary = j1_constr
                            constr_summary['Produit_J2'] = 0

                        # Calculer les totaux
                        constr_summary['Produit_J2'] = constr_summary['Produit_J2'].fillna(0)
                        constr_summary['Total_realisable'] = constr_summary['J1_produit'] + constr_summary['Produit_J2']
                        constr_summary['Pourcentage'] = (constr_summary['Total_realisable'] / constr_summary['Demande_totale'].replace(0, 1) * 100).round(1).astype(str) + '%'

                        # Renommer les colonnes
                        constr_summary = constr_summary.rename(columns={
                            'Demande_totale': 'Conf demandé',
                            'J1_produit': 'Conf réalisable avec réemploi du stock',
                            'Produit_J2': 'Conf réalisable avec achat ref',
                            'Total_realisable': 'Total réalisable',
                            'Pourcentage': 'Pourcentage'
                        })

                        st.dataframe(constr_summary, use_container_width=True, hide_index=True)

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
                        if d['Statut'].startswith('Produit'):
                            mask = (all_demands['Conf|version'] == d['Conf|version']) & (all_demands['Priorité'] == d['Priorité'])
                            s_col = d['Seuil']
                            all_demands.loc[mask, s_col] = all_demands.loc[mask, s_col] + d['Nb produit']
                            all_demands.loc[mask, f"Cout {d['Seuil']}"] = all_demands.loc[mask, f"Cout {d['Seuil']}"] + (d['Cout achat'] or 0)

                    all_demands_xls = all_demands.rename(columns={'J1 produit': 'Opti réemploi', 'J1 Next': 'Opti Next'})
                    all_demands_xls['BOM Next'] = all_demands_xls.apply(
                        lambda row: next_map.get(row['Conf|version'], '') if row.get('Opti Next', 0) > 0 else '', axis=1
                    )
                    _opti_next = all_demands_xls['Opti Next'] if 'Opti Next' in all_demands_xls.columns else 0
                    all_demands_xls['Total'] = (all_demands_xls['Opti réemploi'] + _opti_next
                                                + all_demands_xls['20%'] + all_demands_xls['40%']
                                                + all_demands_xls['60%'] + all_demands_xls['80%'])
                    all_demands_xls['Restant'] = all_demands_xls['Demande'] - all_demands_xls['Total']
                    all_demands_xls['RUPTURE'] = all_demands_xls['Restant'].apply(lambda x: 'OUI' if x > 0 else '')
                    all_demands_xls['Cout conf restantes'] = all_demands_xls['Restant'] * all_demands_xls['Prix'].fillna(0)

                    def _compute_seuil_rupture(row):
                        if row['Restant'] == 0: return ''
                        last = None
                        for s in [20, 40, 60, 80]:
                            if row[f'{s}%'] > 0: last = s
                        if last is None:
                            if row.get('Opti Next', 0) > 0:
                                return 'apres substitution Next'
                            return 'apres Opti réemploi' if row['Opti réemploi'] > 0 else 'des le depart'
                        seuils_list = [20, 40, 60, 80]
                        idx = seuils_list.index(last)
                        return f'a {seuils_list[idx+1]}%' if idx < len(seuils_list)-1 else 'apres 80%'
                    all_demands_xls['Seuil rupture'] = all_demands_xls.apply(_compute_seuil_rupture, axis=1)

                    wb = Workbook()

                    # ---- Onglet 1 : Demande ----
                    ws_d = wb.active; ws_d.title = "Demande"
                    cols_d = [c for c in ['Constructeur','Conf','Version','Priorité','Demande'] if c in all_demands_xls.columns]
                    df_d = all_demands_xls[cols_d].copy()
                    df_d = df_d.rename(columns={'Demande': 'Demande de configuration'})
                    df_d = df_d.drop_duplicates().sort_values(['Constructeur','Priorité'] if 'Constructeur' in df_d.columns else ['Priorité'])
                    tot_row_d = ws_d.max_row + 1
                    for rr in dataframe_to_rows(df_d, index=False, header=True): ws_d.append(rr)
                    tot_row_d = ws_d.max_row + 1
                    ws_d.cell(tot_row_d, 1, "TOTAL")
                    cidx_d = {list(df_d.columns)[i]: i+1 for i in range(len(df_d.columns))}
                    if 'Demande de configuration' in cidx_d: ws_d.cell(tot_row_d, cidx_d['Demande de configuration'], df_d['Demande de configuration'].sum())
                    ligne_total_j2(ws_d, tot_row_d, len(df_d.columns))
                    fmt_h(ws_d); adj_w(ws_d)
                    for row in ws_d.iter_rows(min_row=2, max_row=ws_d.max_row):
                        for cell in row:
                            h = ws_d.cell(1, cell.column).value
                            if h == 'Demande de configuration' and isinstance(cell.value, (int, float)): cell.number_format = NUM_FMT

                    # ---- Onglet 2 : BOM ----
                    ws_bom_j2 = wb.create_sheet("BOM")
                    bom_j2_exp = bom_df[['Constructeur','Conf|version','Conf','Version','Référence','Quantité']].copy()
                    ref_to_designation = dict(zip(stock_df['Referentiel'], stock_df['Designation']))
                    bom_j2_exp['Designation'] = bom_j2_exp['Référence'].map(ref_to_designation).fillna("")
                    bom_j2_exp = bom_j2_exp.sort_values(['Constructeur','Conf|version','Référence'])
                    for rr in dataframe_to_rows(bom_j2_exp, index=False, header=True): ws_bom_j2.append(rr)
                    fmt_h(ws_bom_j2); adj_w(ws_bom_j2)

                    # ---- Onglet 3 : Stock ----
                    ws_s = wb.create_sheet("Stock")
                    stock_out = stock_df.copy().rename(columns={'Referentiel':'Référence','NVX STOCK': 'Stock initial', 'Prix (pj)': 'Prix unitaire'})
                    _consomme_j1 = stock_out['Référence'].map(consumed_j1).fillna(0).round(2)
                    _consomme_j2 = stock_out['Référence'].map(consumed_j2).fillna(0).round(2)
                    stock_out['Stock consommable'] = (_consomme_j1 + _consomme_j2).round(2)
                    stock_out['Stock final'] = (stock_out['Stock initial'] - stock_out['Stock consommable']).round(2)
                    cols_s = ['Référence','Designation','Stock initial','Prix unitaire','Stock consommable','Stock final']
                    stock_out = stock_out[cols_s].sort_values('Référence')
                    for rr in dataframe_to_rows(stock_out, index=False, header=True): ws_s.append(rr)
                    tot_s = ws_s.max_row + 1
                    ws_s.cell(tot_s, 1, "TOTAL")
                    cidx_s = {cols_s[i]: i+1 for i in range(len(cols_s))}
                    if 'Stock consommable' in cidx_s: ws_s.cell(tot_s, cidx_s['Stock consommable'], stock_out['Stock consommable'].sum())
                    ligne_total_j2(ws_s, tot_s, len(cols_s))
                    fmt_h(ws_s); adj_w(ws_s)
                    for row in ws_s.iter_rows(min_row=2, max_row=ws_s.max_row):
                        for cell in row:
                            h = ws_s.cell(1, cell.column).value
                            if h in ['Prix unitaire'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
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
                        if d['Statut'].startswith('Produit'):
                            sn = int(d['Seuil'].replace('%',''))
                            cout_j2_si_complet_par_seuil[sn] += d['Nb produit'] * prix_configs.get(d['Conf|version'], 0)
                    cout_j2_si_complet = sum(cout_j2_si_complet_par_seuil.values())
                    economie_j2 = cout_j2_si_complet - cout_achats_j2

                    # Total de conf demandé
                    ws_r.cell(r_row, 1, "Total de conf demandé"); ws_r.cell(r_row, 2, total_demande).number_format = NUM_FMT
                    r_row += 2

                    # Optimisation du réemploi
                    ws_r.cell(r_row, 1, "OPTIMISATION DU RÉEMPLOI - Production avec stock existant")
                    for ci in range(1, 6): ws_r.cell(r_row, ci).font = BOLD_NOIR; ws_r.cell(r_row, ci).fill = BLEU_CLAIR
                    r_row += 1
                    _opti_next_si = all_demands_xls['Opti Next'] if 'Opti Next' in all_demands_xls.columns else 0
                    cout_j1_si_complet = float(((all_demands_xls['Opti réemploi'] + _opti_next_si) * all_demands_xls['Prix'].fillna(0)).sum())
                    ws_r.cell(r_row, 1, "  Configs réalisables"); ws_r.cell(r_row, 2, total_j1).number_format = NUM_FMT
                    ws_r.cell(r_row, 3, f"{pct_j1:.1f}%"); ws_r.cell(r_row, 4, cout_j1).number_format = EURO_FMT
                    r_row += 1
                    if total_j1_next > 0:
                        ws_r.cell(r_row, 1, "    dont via substitution (Next)")
                        ws_r.cell(r_row, 2, int(total_j1_next)).number_format = NUM_FMT
                        r_row += 1
                    r_row += 1

                    # Achat de références
                    ws_r.cell(r_row, 1, "ACHAT DE RÉFÉRENCES - Production avec achats de references unitaires")
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
                    ws_r.cell(r_row, 1, "  Sous-total Achat références"); ws_r.cell(r_row, 2, total_j2).number_format = NUM_FMT
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
                    for ci in range(1, 10): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 1
                    for i, h in enumerate(["Priorite","Demande","Opti réemploi","Opti Next","J2 Total","Cout J2","Total","Restant","% couvert"], 1):
                        ws_r.cell(r_row, i, h).font = BOLD_NOIR; ws_r.cell(r_row, i).fill = BLEU_CLAIR
                        ws_r.cell(r_row, i).alignment = Alignment(horizontal="center")
                    r_row += 1
                    for prio in priorities:
                        prio_data = all_demands_xls[all_demands_xls['Priorité'] == prio]
                        dem_p = int(prio_data['Demande'].sum())
                        j1_p  = int(prio_data['Opti réemploi'].sum())
                        j1_next_p = int(prio_data['Opti Next'].sum()) if 'Opti Next' in prio_data.columns else 0
                        j2_p  = int(prio_data[['20%','40%','60%','80%']].sum().sum())
                        cout_j2_p = prio_data[['Cout 20%','Cout 40%','Cout 60%','Cout 80%']].sum().sum()
                        tot_p = j1_p + j1_next_p + j2_p; rest_p = dem_p - tot_p
                        pct_p = f"{tot_p/dem_p*100:.0f}%" if dem_p > 0 else "0%"
                        ws_r.cell(r_row, 1, prio).alignment = Alignment(horizontal="center")
                        ws_r.cell(r_row, 2, dem_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 3, j1_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 4, j1_next_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 5, j2_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 6, round(cout_j2_p, 2)).number_format = EURO_FMT
                        ws_r.cell(r_row, 7, tot_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 8, rest_p).number_format = NUM_FMT
                        ws_r.cell(r_row, 9, pct_p).alignment = Alignment(horizontal="center")
                        if tot_p == dem_p: ws_r.cell(r_row, 9).fill = VERT
                        elif tot_p == 0: ws_r.cell(r_row, 9).fill = ROUGE
                        else: ws_r.cell(r_row, 9).fill = ORANGE
                        r_row += 1

                    # -- Section C : Detail par configuration --
                    r_row += 3
                    ws_r.cell(r_row, 1, "DETAIL PAR CONFIGURATION").font = BOLD_BLANC_14
                    for ci in range(1, 22): ws_r.cell(r_row, ci).fill = BLEU_FONCE
                    r_row += 1
                    cols_cfg = [c for c in ['Constructeur','Conf','Version','Conf|version','Priorité','Demande','Prix',
                                            'Opti réemploi','Opti Next','BOM Next',
                                            '20%','Cout 20%','40%','Cout 40%','60%','Cout 60%','80%','Cout 80%',
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
                    for cn in ['Demande','Opti réemploi','Opti Next','20%','Cout 20%','40%','Cout 40%','60%','Cout 60%','80%','Cout 80%','Total','Restant','Cout conf restantes']:
                        if cn in cidx_cfg: ws_r.cell(r_row, cidx_cfg[cn], df_cfg[cn].sum())
                    ligne_total_j2(ws_r, r_row, len(cols_cfg))
                    # Formatage Section C
                    for rn in range(start_cfg, r_row + 1):
                        for ci in range(1, len(cols_cfg)+1):
                            cell = ws_r.cell(rn, ci); h = cols_cfg[ci-1]
                            if h in ['Prix','Cout 20%','Cout 40%','Cout 60%','Cout 80%','Cout conf restantes'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                            elif h in ['Demande','Opti réemploi','Opti Next','20%','40%','60%','80%','Total','Restant'] and isinstance(cell.value, (int, float)): cell.number_format = NUM_FMT
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
                        df_dj2 = df_dj2.rename(columns={'Nb produit': 'Réalisable', 'Prix conf': 'Prix config', 'Cout moyen/conf': 'Cout moy/conf'})
                        df_dj2['Restant apres'] = (df_dj2['Demande restante'].fillna(0) - df_dj2['Réalisable'].fillna(0)).astype(int)
                        cols_dj2 = [c for c in ['Seuil','Priorité','Conf|version','Demande restante','Réalisable','Restant apres','Prix config','Budget max','Cout achat','Cout moy/conf','% du prix','Statut'] if c in df_dj2.columns]
                        df_dj2 = df_dj2[cols_dj2].sort_values(['Seuil','Priorité','Conf|version'])
                        for rr in dataframe_to_rows(df_dj2, index=False, header=True): ws_j2.append(rr)
                        tot_dj2 = ws_j2.max_row + 1
                        ws_j2.cell(tot_dj2, 1, "TOTAL")
                        cidx_dj2 = {cols_dj2[i]: i+1 for i in range(len(cols_dj2))}
                        if 'Réalisable' in cidx_dj2: ws_j2.cell(tot_dj2, cidx_dj2['Réalisable'], df_dj2['Réalisable'].sum())
                        if 'Cout achat' in cidx_dj2: ws_j2.cell(tot_dj2, cidx_dj2['Cout achat'], df_dj2['Cout achat'].sum())
                        ligne_total_j2(ws_j2, tot_dj2, len(cols_dj2))
                        fmt_h(ws_j2); adj_w(ws_j2)
                        for row in ws_j2.iter_rows(min_row=2, max_row=ws_j2.max_row):
                            for cell in row:
                                h = ws_j2.cell(1, cell.column).value
                                if h in ['Prix config','Budget max','Cout achat','Cout moy/conf'] and isinstance(cell.value, (int, float)): cell.number_format = EURO_FMT
                        col_stat = cidx_dj2.get('Statut')
                        if col_stat:
                            for rn in range(2, ws_j2.max_row):
                                c = ws_j2.cell(rn, col_stat)
                                if c.value and str(c.value).startswith('Produit'): c.fill = VERT
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
                    st.session_state['j2_excel_name'] = f"Optimisation_Reemploi_Avec_Achat_{ts}.xlsx"

                except Exception as e:
                    st.error(f"Erreur Jalon 2 : {e}")
                    import traceback
                    st.code(traceback.format_exc())

        if 'j2_excel' in st.session_state:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="TÉLÉCHARGER LES RÉSULTATS",
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
st.markdown('<div class="footer-text">FORECAST — Batail-log &nbsp;·&nbsp; PuLP CBC Solver &nbsp;·&nbsp; v1.0</div>', unsafe_allow_html=True)

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


def run_jalon1(all_demands, bom_df, stock_df, q, priorities, prix_file_bytes, next_map=None, confc_map=None):
    if next_map is None:
        next_map = {}
    if confc_map is None:
        confc_map = {}
    stock_optim = stock_df.set_index("Referentiel")
    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)
    stock_current = stock_optim['NVX STOCK'].copy()
    produced_j1 = {}
    produced_next_j1 = {}
    produced_confc_j1 = {}

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

        # --- Passe de substitution (CONF C) pour cette priorité ---
        if confc_map:
            confc_candidates = {}
            for cv in X:
                qty_orig = produced_j1.get((cv, prio_level), 0)
                qty_next = produced_next_j1.get((cv, prio_level), 0)
                remaining = demand_map.get(cv, 0) - qty_orig - qty_next
                if remaining > 0 and cv in confc_map:
                    confc_cv = confc_map[cv]
                    if confc_cv in q.index:
                        confc_candidates[cv] = (confc_cv, remaining)
            if confc_candidates:
                prob_c1 = LpProblem(f"ConfC_P{prio_level}_MaxQty", LpMaximize)
                XC1 = {}
                for cv, (confc_cv, rem) in confc_candidates.items():
                    safe = cv.replace('|','_').replace(' ','_')
                    XC1[cv] = LpVariable(f"XC1_{safe}_p{prio_level}", 0, rem, LpInteger)
                prob_c1 += lpSum(XC1.values())
                for ref in q.columns:
                    coeffs = []
                    for cv, (confc_cv, rem) in confc_candidates.items():
                        coef = q.at[confc_cv, ref]
                        if coef > 0:
                            coeffs.append((coef, XC1[cv]))
                    if coeffs:
                        lim = stock_current[ref] if ref in stock_current.index else 0
                        prob_c1 += lpSum(c * v for c, v in coeffs) <= lim
                prob_c1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
                max_confc = value(prob_c1.objective) or 0
                if max_confc > 0:
                    prob_c2 = LpProblem(f"ConfC_P{prio_level}_MaxCost", LpMaximize)
                    XC = {}
                    for cv, (confc_cv, rem) in confc_candidates.items():
                        safe = cv.replace('|','_').replace(' ','_')
                        XC[cv] = LpVariable(f"XC_{safe}_p{prio_level}", 0, rem, LpInteger)
                    cout_confc = lpSum(
                        XC[cv] * sum(q.at[confc_candidates[cv][0], ref] * stock_optim.at[ref, 'Prix (pj)']
                                     for ref in q.columns if ref in stock_optim.index
                                     and q.at[confc_candidates[cv][0], ref] > 0)
                        for cv in XC
                    )
                    prob_c2 += cout_confc
                    prob_c2 += lpSum(XC.values()) == max_confc
                    for ref in q.columns:
                        coeffs = []
                        for cv, (confc_cv, rem) in confc_candidates.items():
                            coef = q.at[confc_cv, ref]
                            if coef > 0:
                                coeffs.append((coef, XC[cv]))
                        if coeffs:
                            lim = stock_current[ref] if ref in stock_current.index else 0
                            prob_c2 += lpSum(c * v for c, v in coeffs) <= lim
                    prob_c2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))
                    for cv in XC:
                        qty_c = max(0, int(round(value(XC[cv].varValue) or 0)))
                        if qty_c > 0:
                            produced_confc_j1[(cv, prio_level)] = qty_c
                            confc_cv = confc_candidates[cv][0]
                            for ref in q.columns:
                                cons = q.at[confc_cv, ref] * qty_c
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
    for (cv, prio), qty in produced_confc_j1.items():
        confc_cv = confc_map.get(cv, '')
        if confc_cv and confc_cv in q.index and qty > 0:
            for ref in q.columns:
                cons = q.at[confc_cv, ref] * qty
                if cons > 0:
                    consumed_j1[ref] = consumed_j1.get(ref, 0) + cons
    cout_j1 = sum(qty * stock_optim.at[ref, 'Prix (pj)'] for ref, qty in consumed_j1.items() if ref in stock_optim.index)
    return produced_j1, consumed_j1, cout_j1, stock_current.copy(), stock_optim, produced_next_j1, produced_confc_j1


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

                    # Agréger la demande par (Constructeur, Conf)
                    demand_by_cconf = df_prio_raw.groupby(['Constructeur', 'Conf'], as_index=False)['Demande'].sum()
                    demand_conf_map = {}
                    for _, row in demand_by_cconf.iterrows():
                        demand_conf_map[(row['Constructeur'], row['Conf'])] = row['Demande']

                    # Versions disponibles par (Constructeur, Conf)
                    versions_by_conf = {}
                    for (cstr, conf) in demand_conf_map:
                        mask = (df_prio_raw['Constructeur'] == cstr) & (df_prio_raw['Conf'] == conf)
                        cvs = df_prio_raw[mask]['Conf|version'].unique().tolist()
                        versions_by_conf[(cstr, conf)] = cvs

                    # Détail par version (pour l'affichage)
                    df_versions = df_prio_raw[['Constructeur', 'Conf', 'Version', 'Conf|version']].drop_duplicates('Conf|version')

                    bom_df = st.session_state['bom_df']
                    stock_df = st.session_state['stock_df'].copy()
                    stock_optim = stock_df.set_index("Referentiel")
                    stock_optim['Prix (pj)'] = pd.to_numeric(stock_optim['Prix (pj)'], errors='coerce').fillna(0.0)

                    q = bom_df.pivot_table(index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum').fillna(0)

                    # Toutes les Conf|versions avec BOM disponible
                    all_cvs = []
                    for key, cvs in versions_by_conf.items():
                        for cv in cvs:
                            if cv in q.index:
                                all_cvs.append(cv)

                    cv_to_conf = {}
                    for key, cvs in versions_by_conf.items():
                        for cv in cvs:
                            cv_to_conf[cv] = key

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
                            prob2 += lpSum(X[cv] for cv in cvs_conf) <= d

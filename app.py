import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pulp import LpProblem, LpVariable, LpMaximize, lpSum, value, PULP_CBC_CMD, LpStatus, LpInteger
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill


# =============================================================================
# CONFIG PAGE
# =============================================================================
st.set_page_config(page_title="FORECAST SFR", page_icon="📦", layout="wide")
st.title("📦 FORECAST SFR")
st.caption("Allocation optimale de composants telecom")


# =============================================================================
# FONCTIONS : PREPARATION DONNEES (logique de fichier entrée.py)
# =============================================================================
def load_bom_sheet(biblio_bytes, sheet_name):
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


def prepare_data(biblio_file, prix_file, forecast_file, stock_file, acc_file):
    """Reproduit la logique de fichier entrée.py avec des fichiers uploadés."""

    # --- Demandes + Prix ---
    df_forecast = pd.read_excel(
        forecast_file, sheet_name="Feuil1",
        usecols=["Constructeur", "Conf", "Version", "Conf|version", "Demande"]
    )
    df_forecast["Demande"] = pd.to_numeric(df_forecast["Demande"], errors="coerce").fillna(0)
    df_forecast = df_forecast[df_forecast["Demande"] > 0].copy()

    df_prix_conf = pd.read_excel(prix_file, sheet_name="Conf", usecols=["Conf|version", "Prix"])
    df_prix_conf["Prix"] = pd.to_numeric(df_prix_conf["Prix"], errors="coerce")

    demand_df = df_forecast.merge(df_prix_conf, on="Conf|version", how="left")

    # --- BOM ---
    biblio_bytes_nokia = BytesIO(biblio_file.getvalue())
    biblio_bytes_huawei = BytesIO(biblio_file.getvalue())
    bom_nokia = load_bom_sheet(biblio_bytes_nokia, "NOKIA")
    bom_huawei = load_bom_sheet(biblio_bytes_huawei, "huawei")
    bom_df = pd.concat([bom_nokia, bom_huawei], ignore_index=True)

    # --- Normalisation BOM via acc.xlsx ---
    acc_bytes_bom = BytesIO(acc_file.getvalue())
    acc_bom = pd.read_excel(acc_bytes_bom, sheet_name="Feuil1")
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

    # --- acc mapping stock ---
    acc_bytes_stock = BytesIO(acc_file.getvalue())
    acc = pd.read_excel(acc_bytes_stock, sheet_name="Feuil1")
    acc.columns = acc.columns.str.strip().str.lower().str.replace(r'[\s,]+', '', regex=True)

    column_map = {}
    for col in acc.columns:
        if 'code' in col:
            column_map[col] = 'code'
        elif any(k in col for k in ['ref', 'reference', 'article']):
            column_map[col] = 'ref'
        elif any(k in col for k in ['bom', 'multi', 'coeff']):
            column_map[col] = 'bom'
    acc = acc.rename(columns=column_map)

    required = {'code', 'ref', 'bom'}
    if not required.issubset(acc.columns):
        raise ValueError(f"acc.xlsx : Colonnes manquantes -> {required - set(acc.columns)}")

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

    stock_df = df_stock.groupby("Referentiel", as_index=False).agg({
        "Valeur_NVX": "sum", "Stock": "sum"
    }).rename(columns={"Valeur_NVX": "NVX STOCK", "Stock": "Stock Physique"})
    stock_df = stock_df.sort_values("NVX STOCK", ascending=False).reset_index(drop=True)

    # --- Prix (pj) ---
    prix_bytes = BytesIO(prix_file.getvalue())
    df_prix_ref = pd.read_excel(prix_bytes, sheet_name="References", usecols=["Référence", "Prix (pj)"])
    df_prix_ref["Référence"] = df_prix_ref["Référence"].astype(str).str.strip().str.upper()
    df_prix_ref["Prix (pj)"] = pd.to_numeric(df_prix_ref["Prix (pj)"], errors='coerce').fillna(0)

    prix_unique = df_prix_ref.groupby("Référence", as_index=False)["Prix (pj)"].first()
    ref_to_prix_dict = dict(zip(prix_unique["Référence"], prix_unique["Prix (pj)"]))
    stock_df["Prix (pj)"] = stock_df["Referentiel"].map(ref_to_prix_dict).fillna(0)

    stock_df = stock_df[["Referentiel", "Stock Physique", "NVX STOCK", "Prix (pj)"]].rename(
        columns={"Stock Physique": "Stock"}
    )

    return demand_df, bom_df, stock_df


def data_to_excel_bytes(demand_df, bom_df, stock_df):
    """Génère le fichier consolidé en mémoire."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        demand_df.to_excel(writer, sheet_name="Demandes", index=False)
        bom_df.to_excel(writer, sheet_name="BOM", index=False)
        stock_df.to_excel(writer, sheet_name="Stock_Clean", index=False)
    return output.getvalue()


# =============================================================================
# FONCTIONS : OPTIMISATION SANS PRIORITE (logique de RO_J1_sans_prio.py)
# =============================================================================
def run_optim_sans_prio(demand_df, bom_df, stock_df_raw):
    stock_df = stock_df_raw.copy().set_index("Referentiel")
    stock_df['Prix (pj)'] = pd.to_numeric(stock_df['Prix (pj)'], errors='coerce').fillna(0.0)

    q = bom_df.pivot_table(
        index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum'
    ).reindex(demand_df['Conf|version'].unique()).fillna(0)

    configs = demand_df['Conf|version'].unique()
    refs_bloquantes = [ref for ref in q.columns if ref not in stock_df.index or stock_df.at[ref, 'NVX STOCK'] <= 0]

    # ETAPE 1 : Max configs
    prob1 = LpProblem("Max_Nb_Configs", LpMaximize)
    X1 = {}
    for cv in configs:
        dem_max = demand_df[demand_df['Conf|version'] == cv]['Demande'].iloc[0]
        X1[cv] = LpVariable(f"X1_{cv.replace('|','_').replace('-','_')}", 0, dem_max, cat='Integer')

    prob1 += lpSum(X1.values())
    for ref in q.columns:
        if ref in stock_df.index:
            prob1 += lpSum(q.at[cv, ref] * X1[cv] for cv in configs if q.at[cv, ref] > 0) <= stock_df.at[ref, 'NVX STOCK']
        else:
            prob1 += lpSum(q.at[cv, ref] * X1[cv] for cv in configs if q.at[cv, ref] > 0) <= 0

    prob1.solve(PULP_CBC_CMD(msg=0, timeLimit=900, gapRel=0.0))
    max_configs = value(prob1.objective) or 0

    # ETAPE 2 : Max cout stock
    prob2 = LpProblem("Max_Cout_Stock", LpMaximize)
    X = {}
    for cv in configs:
        dem_max = demand_df[demand_df['Conf|version'] == cv]['Demande'].iloc[0]
        X[cv] = LpVariable(f"X_{cv.replace('|','_').replace('-','_')}", 0, dem_max, cat='Integer')

    cout_stock = lpSum(
        X[cv] * lpSum(q.at[cv, ref] * stock_df.at[ref, 'Prix (pj)'] for ref in q.columns if ref in stock_df.index)
        for cv in configs
    )
    prob2 += cout_stock
    prob2 += lpSum(X.values()) == max_configs

    for ref in q.columns:
        if ref in stock_df.index:
            prob2 += lpSum(q.at[cv, ref] * X[cv] for cv in configs if q.at[cv, ref] > 0) <= stock_df.at[ref, 'NVX STOCK']
        else:
            prob2 += lpSum(q.at[cv, ref] * X[cv] for cv in configs if q.at[cv, ref] > 0) <= 0

    prob2.solve(PULP_CBC_CMD(msg=0, timeLimit=1800, gapRel=0.01))

    produced = {cv: max(0, int(round(value(X[cv]) or 0))) for cv in X}

    consumed = {}
    for ref in q.columns:
        qty = sum(q.at[cv, ref] * produced.get(cv, 0) for cv in configs)
        if qty > 0:
            consumed[ref] = round(qty)

    cout_total = sum(qty * stock_df.at[ref, 'Prix (pj)'] for ref, qty in consumed.items() if ref in stock_df.index)

    return {
        'produced': produced,
        'consumed': consumed,
        'cout_total': cout_total,
        'max_configs': max_configs,
        'refs_bloquantes': refs_bloquantes,
        'status': LpStatus[prob2.status],
        'q': q,
        'stock_df': stock_df,
        'configs': configs,
    }


# =============================================================================
# FONCTIONS : OPTIMISATION AVEC PRIORITE (logique de RO_J2_avec_prio.py)
# =============================================================================
def run_optim_avec_prio(demand_df, bom_df, stock_df_raw, prio_df):
    stock_df = stock_df_raw.copy().set_index("Referentiel")
    stock_df['Prix (pj)'] = pd.to_numeric(stock_df['Prix (pj)'], errors='coerce').fillna(0.0)

    demand_df = demand_df.copy()
    demand_df['Demande'] = pd.to_numeric(demand_df['Demande'], errors='coerce').fillna(0).clip(lower=0).astype(int)

    q = bom_df.pivot_table(
        index='Conf|version', columns='Référence', values='Quantité', aggfunc='sum'
    ).reindex(demand_df['Conf|version'].unique()).fillna(0)

    configs = demand_df['Conf|version'].unique()
    refs_bloquantes = [ref for ref in q.columns if ref not in stock_df.index or stock_df.at[ref, 'NVX STOCK'] <= 0]

    # Ajout priorité
    demand_df['Priorité'] = 999
    prio_clean = prio_df[['Conf|version', 'Priorité']].copy()
    prio_clean['Priorité'] = pd.to_numeric(prio_clean['Priorité'], errors='coerce').fillna(999).astype(int)
    demand_df = demand_df.merge(prio_clean, on='Conf|version', how='left', suffixes=('', '_p'))
    demand_df['Priorité'] = demand_df['Priorité_p'].combine_first(demand_df['Priorité'])
    demand_df.drop(columns=['Priorité_p'], errors='ignore', inplace=True)

    stock_current = stock_df['NVX STOCK'].copy()
    produced = {cv: 0 for cv in configs}
    log_lines = []

    priorities = sorted(demand_df['Priorité'].unique())

    for phase, prio_level in enumerate(priorities, 1):
        phase_configs = demand_df[demand_df['Priorité'] == prio_level]['Conf|version'].tolist()
        if not phase_configs:
            continue

        # Etape 1 : Max configs
        prob1 = LpProblem(f"Force_Prio_{prio_level}_MaxQty", LpMaximize)
        X1 = {}
        for cv in phase_configs:
            dem_max = demand_df[demand_df['Conf|version'] == cv]['Demande'].iloc[0]
            X1[cv] = LpVariable(f"X1_{cv.replace('|','_')}", 0, dem_max, LpInteger)

        prob1 += lpSum(X1.values())
        for ref in q.columns:
            if ref in stock_current.index:
                prob1 += lpSum(q.at[cv, ref] * X1[cv] for cv in phase_configs if q.at[cv, ref] > 0) <= stock_current[ref]
            else:
                prob1 += lpSum(q.at[cv, ref] * X1[cv] for cv in phase_configs if q.at[cv, ref] > 0) <= 0

        prob1.solve(PULP_CBC_CMD(msg=0, gapRel=0.0))
        max_phase = value(prob1.objective) or 0

        # Etape 2 : Max cout
        prob2 = LpProblem(f"Force_Prio_{prio_level}_MaxCost", LpMaximize)
        X = {}
        for cv in phase_configs:
            dem_max = demand_df[demand_df['Conf|version'] == cv]['Demande'].iloc[0]
            X[cv] = LpVariable(f"X_{cv.replace('|','_')}", 0, dem_max, LpInteger)

        cout_phase = lpSum(
            X[cv] * lpSum(q.at[cv, ref] * stock_df.at[ref, 'Prix (pj)'] for ref in q.columns if ref in stock_df.index)
            for cv in phase_configs
        )
        prob2 += cout_phase
        prob2 += lpSum(X.values()) == max_phase

        for ref in q.columns:
            if ref in stock_current.index:
                prob2 += lpSum(q.at[cv, ref] * X[cv] for cv in phase_configs if q.at[cv, ref] > 0) <= stock_current[ref]
            else:
                prob2 += lpSum(q.at[cv, ref] * X[cv] for cv in phase_configs if q.at[cv, ref] > 0) <= 0

        prob2.solve(PULP_CBC_CMD(msg=0, gapRel=0.01))

        phase_prod = 0
        for cv in phase_configs:
            qty = max(0, int(round(value(X[cv].varValue) or 0)))
            produced[cv] = qty
            phase_prod += qty
            for ref in q.columns:
                cons = q.at[cv, ref] * qty
                if cons > 0 and ref in stock_current.index:
                    stock_current[ref] = max(0, stock_current[ref] - cons)

        log_lines.append(f"Phase {phase} - Priorite {prio_level} : {phase_prod} configs produites")

    consumed = {}
    for ref in q.columns:
        qty = sum(q.at[cv, ref] * produced.get(cv, 0) for cv in configs)
        if qty > 0:
            consumed[ref] = round(qty)

    cout_total = sum(qty * stock_df.at[ref, 'Prix (pj)'] for ref, qty in consumed.items() if ref in stock_df.index)

    return {
        'produced': produced,
        'consumed': consumed,
        'cout_total': cout_total,
        'refs_bloquantes': refs_bloquantes,
        'demand_df': demand_df,
        'log_lines': log_lines,
        'q': q,
        'stock_df': stock_df,
        'configs': configs,
    }


# =============================================================================
# FONCTIONS : EXPORT EXCEL EN MEMOIRE
# =============================================================================
def generate_excel_sans_prio(demand_df, bom_df, stock_df_raw, res):
    produced = res['produced']
    consumed = res['consumed']
    cout_total = res['cout_total']
    refs_bloquantes = res['refs_bloquantes']
    stock_df = res['stock_df']
    q = res['q']
    configs = res['configs']

    wb = Workbook()

    # --- BOM ---
    ws_bom = wb.active
    ws_bom.title = "BOM"

    bom_enrichi = bom_df.copy()
    bom_enrichi = bom_enrichi.merge(
        stock_df[['NVX STOCK']].rename(columns={'NVX STOCK': 'Stock initial'}),
        left_on='Référence', right_index=True, how='left'
    ).fillna({'Stock initial': 0})

    produced_series = pd.Series(produced, name='Qté produite config').rename_axis('Conf|version')
    bom_enrichi = bom_enrichi.merge(produced_series.reset_index(), on='Conf|version', how='left').fillna({'Qté produite config': 0})
    bom_enrichi['Stock consommé (cette config)'] = bom_enrichi['Quantité'] * bom_enrichi['Qté produite config']

    consumed_series = pd.Series(consumed, name='A consomé total').rename_axis('Référence')
    bom_enrichi = bom_enrichi.merge(consumed_series.reset_index(), on='Référence', how='left').fillna({'A consomé total': 0})
    bom_enrichi['Stock final (global ref)'] = bom_enrichi['Stock initial'] - bom_enrichi['A consomé total']

    cols_bom = ['Conf|version', 'Conf', 'Version', 'Référence', 'Quantité',
                'Stock initial', 'Stock consommé (cette config)', 'Stock final (global ref)', 'Qté produite config']
    bom_enrichi = bom_enrichi[[c for c in cols_bom if c in bom_enrichi.columns]]

    for r in dataframe_to_rows(bom_enrichi, index=False, header=True):
        ws_bom.append(r)

    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    for row in ws_bom.iter_rows(min_row=2):
        for cell in row:
            if cell.column_letter in ['F', 'G', 'H']:
                cell.fill = yellow_fill

    # --- Demandes ---
    ws_demand = wb.create_sheet("Demandes")
    demand_out = demand_df.copy()
    if 'Constructeur' not in demand_out.columns:
        demand_out[['Constructeur', 'Conf', 'Version']] = demand_out['Conf|version'].str.split('_', n=2, expand=True)
    demand_out['Qté produite'] = demand_out['Conf|version'].map(produced).fillna(0).astype(int)
    cols_demand = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Demande', 'Prix', 'Qté produite']
    cols_demand = [c for c in cols_demand if c in demand_out.columns]
    for r in dataframe_to_rows(demand_out[cols_demand], index=False, header=True):
        ws_demand.append(r)

    # --- Stock ---
    ws_stock = wb.create_sheet("Stock_Clean")
    stock_out = stock_df.reset_index()
    stock_out['A consomé'] = stock_out['Referentiel'].map(consumed).fillna(0).astype(int)
    cols_stock = ['Referentiel', 'Stock', 'NVX STOCK', 'Prix (pj)', 'A consomé']
    cols_stock = [c for c in cols_stock if c in stock_out.columns]
    for r in dataframe_to_rows(stock_out[cols_stock], index=False, header=True):
        ws_stock.append(r)

    # --- Résultats ---
    ws_res = wb.create_sheet("Résultats")
    total_demand = demand_df['Demande'].sum()
    configs_produites = sum(produced.values())
    pct = (configs_produites / total_demand * 100) if total_demand > 0 else 0

    metriques = [
        ["Métrique", "Valeur"],
        ["MAX configurations demandées", int(total_demand)],
        ["Configurations produites", configs_produites],
        ["% Réalisation", f"{pct:.1f} %"],
        ["Coût total stock consommé (euros)", f"{cout_total:,.2f}"],
        ["Nombre références utilisées", len(consumed)],
        ["Nombre références bloquantes", len(refs_bloquantes)],
    ]
    for ligne in metriques:
        ws_res.append(ligne)

    ws_res.column_dimensions['A'].width = 40
    ws_res.column_dimensions['B'].width = 20

    # Auto-width
    for ws in [ws_bom, ws_demand, ws_stock]:
        for col in ws.columns:
            max_length = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = max_length + 4

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def generate_excel_avec_prio(demand_df_orig, bom_df, stock_df_raw, res):
    produced = res['produced']
    consumed = res['consumed']
    cout_total = res['cout_total']
    refs_bloquantes = res['refs_bloquantes']
    demand_df = res['demand_df']
    stock_df = res['stock_df']

    wb = Workbook()

    # --- BOM ---
    ws_bom = wb.active
    ws_bom.title = "BOM"

    bom_enrichi = bom_df.copy()
    bom_enrichi = bom_enrichi.merge(
        stock_df[['NVX STOCK']].rename(columns={'NVX STOCK': 'Stock initial'}),
        left_on='Référence', right_index=True, how='left'
    ).fillna({'Stock initial': 0})

    produced_series = pd.Series(produced, name='Qté produite config').rename_axis('Conf|version')
    bom_enrichi = bom_enrichi.merge(produced_series.reset_index(), on='Conf|version', how='left').fillna({'Qté produite config': 0})
    bom_enrichi['Stock consommé (cette config)'] = bom_enrichi['Quantité'] * bom_enrichi['Qté produite config']

    consumed_series = pd.Series(consumed, name='A consomé total').rename_axis('Référence')
    bom_enrichi = bom_enrichi.merge(consumed_series.reset_index(), on='Référence', how='left').fillna({'A consomé total': 0})
    bom_enrichi['Stock final (global ref)'] = bom_enrichi['Stock initial'] - bom_enrichi['A consomé total']

    cols_bom = ['Conf|version', 'Conf', 'Version', 'Référence', 'Quantité',
                'Stock initial', 'Stock consommé (cette config)', 'Stock final (global ref)', 'Qté produite config']
    bom_enrichi = bom_enrichi[[c for c in cols_bom if c in bom_enrichi.columns]]

    for r in dataframe_to_rows(bom_enrichi, index=False, header=True):
        ws_bom.append(r)

    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    for row in ws_bom.iter_rows(min_row=2):
        for cell in row:
            if cell.column_letter in ['F', 'G', 'H']:
                cell.fill = yellow_fill

    # --- Demandes avec Priorité ---
    ws_demand = wb.create_sheet("Demandes")
    demand_out = demand_df.copy()
    if 'Constructeur' not in demand_out.columns:
        demand_out[['Constructeur', 'Conf', 'Version']] = demand_out['Conf|version'].str.split('_', n=2, expand=True)
    demand_out['Qté produite'] = demand_out['Conf|version'].map(produced).fillna(0).astype(int)
    cols_demand = ['Constructeur', 'Conf', 'Version', 'Conf|version', 'Priorité', 'Demande', 'Prix', 'Qté produite']
    cols_demand = [c for c in cols_demand if c in demand_out.columns]
    for r in dataframe_to_rows(demand_out[cols_demand], index=False, header=True):
        ws_demand.append(r)

    # --- Stock ---
    ws_stock = wb.create_sheet("Stock_Clean")
    stock_out = stock_df.reset_index()
    stock_out['A consomé'] = stock_out['Referentiel'].map(consumed).fillna(0).astype(int)
    cols_stock = ['Referentiel', 'Stock', 'NVX STOCK', 'Prix (pj)', 'A consomé']
    cols_stock = [c for c in cols_stock if c in stock_out.columns]
    for r in dataframe_to_rows(stock_out[cols_stock], index=False, header=True):
        ws_stock.append(r)

    # --- Résultats ---
    ws_res = wb.create_sheet("Résultats")
    total_demand = demand_df['Demande'].sum()
    configs_produites = sum(produced.values())
    pct = (configs_produites / total_demand * 100) if total_demand > 0 else 0

    metriques = [
        ["Métrique", "Valeur"],
        ["MAX configurations demandées", int(total_demand)],
        ["Configurations produites", configs_produites],
        ["% Réalisation", f"{pct:.1f} %"],
        ["Coût total stock consommé (euros)", f"{cout_total:,.2f}"],
        ["Nombre références utilisées", len(consumed)],
        ["Nombre références bloquantes", len(refs_bloquantes)],
    ]
    for ligne in metriques:
        ws_res.append(ligne)

    # Détails par priorité
    ws_res.append([])
    ws_res.append(["DÉTAILS PAR PRIORITÉ", ""])
    ws_res.append(["Priorité", "Nb Configs", "Demande", "Produit", "Taux %"])

    for prio in sorted(demand_df['Priorité'].unique()):
        prio_rows = demand_df[demand_df['Priorité'] == prio]
        prio_dem = prio_rows['Demande'].sum()
        prio_prod = sum(produced.get(cv, 0) for cv in prio_rows['Conf|version'])
        prio_taux = (prio_prod / prio_dem * 100) if prio_dem > 0 else 0
        ws_res.append([f"Priorité {int(prio)}", len(prio_rows), int(prio_dem), int(prio_prod), f"{prio_taux:.1f} %"])

    ws_res.column_dimensions['A'].width = 50
    ws_res.column_dimensions['B'].width = 20
    ws_res.column_dimensions['C'].width = 20
    ws_res.column_dimensions['D'].width = 15

    for row in ws_res.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.font = Font(bold=True)

    # Auto-width
    for ws in [ws_bom, ws_demand, ws_stock]:
        for col in ws.columns:
            max_length = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = max_length + 4

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# =============================================================================
# ONGLETS
# =============================================================================
tab1, tab2, tab3 = st.tabs([
    "1 - Preparation Donnees",
    "2 - Optimisation SANS Priorite",
    "3 - Optimisation AVEC Priorite"
])


# =============================================================================
# TAB 1 : PREPARATION DONNEES
# =============================================================================
with tab1:
    st.header("Etape 1 : Preparation des donnees")
    st.markdown("""
    Uploadez les **5 fichiers sources** pour generer le fichier consolide.
    """)

    col1, col2 = st.columns(2)
    with col1:
        biblio_file = st.file_uploader("biblio.xlsx (BOM Nokia + Huawei)", type="xlsx", key="biblio")
        prix_file = st.file_uploader("Prix.xlsx (Prix configs + references)", type="xlsx", key="prix")
        forecast_file = st.file_uploader("Forecast.xlsx (Demande par config)", type="xlsx", key="forecast")
    with col2:
        stock_file = st.file_uploader("Stock.xlsx (Stock dispo + previsionnel)", type="xlsx", key="stock")
        acc_file = st.file_uploader("acc.xlsx (Table conversion code → ref)", type="xlsx", key="acc")

    all_uploaded = all([biblio_file, prix_file, forecast_file, stock_file, acc_file])

    if all_uploaded:
        st.success("5/5 fichiers uploades")
        if st.button("Lancer la preparation", key="btn_s1", type="primary"):
            with st.spinner("Preparation en cours..."):
                try:
                    demand_df, bom_df, stock_df = prepare_data(
                        biblio_file, prix_file, forecast_file, stock_file, acc_file
                    )
                    st.session_state['demand_df'] = demand_df
                    st.session_state['bom_df'] = bom_df
                    st.session_state['stock_df'] = stock_df
                    st.session_state['prep_done'] = True

                    excel_bytes = data_to_excel_bytes(demand_df, bom_df, stock_df)
                    st.session_state['input_excel'] = excel_bytes

                    st.success("Donnees preparees avec succes !")

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    st.session_state['prep_done'] = False

    if st.session_state.get('prep_done'):
        demand_df = st.session_state['demand_df']
        bom_df = st.session_state['bom_df']
        stock_df = st.session_state['stock_df']

        c1, c2, c3 = st.columns(3)
        c1.metric("Configurations", len(demand_df))
        c2.metric("Lignes BOM", len(bom_df))
        c3.metric("References stock", len(stock_df))

        with st.expander("Apercu Demandes"):
            st.dataframe(demand_df, width=None, hide_index=True)
        with st.expander("Apercu Stock"):
            st.dataframe(stock_df.head(20), width=None, hide_index=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            "Telecharger le fichier consolide",
            data=st.session_state['input_excel'],
            file_name=f"Entree_RO_Reel_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =============================================================================
# TAB 2 : OPTIMISATION SANS PRIORITE
# =============================================================================
with tab2:
    st.header("Etape 2 : Optimisation SANS priorite (max configs)")
    st.markdown("""
    **Etape 1** : Maximiser le nombre total de configurations.
    **Etape 2** : A nombre de configs fixe, maximiser le cout du stock consomme.
    """)

    # Source de données
    data_ready = False
    if st.session_state.get('prep_done'):
        st.info("Donnees de l'etape 1 detectees. Pret a optimiser.")
        demand_df = st.session_state['demand_df']
        bom_df = st.session_state['bom_df']
        stock_df = st.session_state['stock_df']
        data_ready = True
    else:
        st.warning("Uploadez le fichier consolide (genere a l'etape 1) :")
        uploaded_input = st.file_uploader("Fichier Entree_RO_Reel_*.xlsx", type="xlsx", key="input_s2")
        if uploaded_input:
            try:
                demand_df = pd.read_excel(uploaded_input, sheet_name="Demandes")
                bom_df = pd.read_excel(uploaded_input, sheet_name="BOM")
                stock_df = pd.read_excel(uploaded_input, sheet_name="Stock_Clean")
                data_ready = True
                st.success("Fichier charge !")
            except Exception as e:
                st.error(f"Erreur lecture : {e}")

    if data_ready:
        if st.button("Lancer l'optimisation SANS priorite", key="btn_s2", type="primary"):
            with st.spinner("Optimisation en cours..."):
                try:
                    res = run_optim_sans_prio(demand_df, bom_df, stock_df)
                    st.session_state['res_sans_prio'] = res
                    st.session_state['sans_prio_done'] = True
                except Exception as e:
                    st.error(f"Erreur optimisation : {e}")
                    st.session_state['sans_prio_done'] = False

    if st.session_state.get('sans_prio_done'):
        res = st.session_state['res_sans_prio']
        produced = res['produced']
        total_demand = demand_df['Demande'].sum()
        configs_produites = sum(produced.values())
        pct = (configs_produites / total_demand * 100) if total_demand > 0 else 0

        st.subheader("Resultats")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Configs produites", configs_produites)
        c2.metric("Configs demandees", int(total_demand))
        c3.metric("% Realisation", f"{pct:.1f} %")
        c4.metric("Cout stock consomme", f"{res['cout_total']:,.2f} EUR")

        st.subheader("Detail par configuration")
        detail = demand_df[['Conf|version', 'Demande']].copy()
        detail['Qté produite'] = detail['Conf|version'].map(produced).fillna(0).astype(int)
        detail['Ecart'] = detail['Demande'] - detail['Qté produite']
        detail['% Realise'] = (detail['Qté produite'] / detail['Demande'] * 100).round(1)
        st.dataframe(detail, width=None, hide_index=True)

        if len(res['refs_bloquantes']) > 0:
            with st.expander(f"{len(res['refs_bloquantes'])} references bloquantes"):
                st.write(", ".join(res['refs_bloquantes']))

        excel_bytes = generate_excel_sans_prio(demand_df, bom_df, stock_df, res)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Telecharger le fichier Excel",
            data=excel_bytes,
            file_name=f"Optim_Max_Configs_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =============================================================================
# TAB 3 : OPTIMISATION AVEC PRIORITE
# =============================================================================
with tab3:
    st.header("Etape 3 : Optimisation AVEC priorite (sequentielle)")
    st.markdown("""
    Optimisation par priorite stricte :
    - **Phase 1** : Maximiser prio 1 (avec tout le stock)
    - **Phase 2** : Maximiser prio 2 (avec le stock restant)
    - etc.
    """)

    data_ready_p = False
    prio_ready = False

    if st.session_state.get('prep_done'):
        st.info("Donnees de l'etape 1 detectees.")
        demand_df_p = st.session_state['demand_df']
        bom_df_p = st.session_state['bom_df']
        stock_df_p = st.session_state['stock_df']
        data_ready_p = True
    else:
        st.warning("Uploadez le fichier consolide :")
        uploaded_input_p = st.file_uploader("Fichier Entree_RO_Reel_*.xlsx", type="xlsx", key="input_s3")
        if uploaded_input_p:
            try:
                demand_df_p = pd.read_excel(uploaded_input_p, sheet_name="Demandes")
                bom_df_p = pd.read_excel(uploaded_input_p, sheet_name="BOM")
                stock_df_p = pd.read_excel(uploaded_input_p, sheet_name="Stock_Clean")
                data_ready_p = True
                st.success("Fichier charge !")
            except Exception as e:
                st.error(f"Erreur lecture : {e}")

    prio_file = st.file_uploader("Prio.xlsx (fichier priorites)", type="xlsx", key="prio")
    if prio_file:
        try:
            prio_df = pd.read_excel(prio_file, sheet_name="Feuil1")
            prio_ready = True
            st.success("Fichier priorites charge !")
        except Exception as e:
            st.error(f"Erreur lecture Prio.xlsx : {e}")

    if data_ready_p and prio_ready:
        if st.button("Lancer l'optimisation AVEC priorite", key="btn_s3", type="primary"):
            with st.spinner("Optimisation en cours..."):
                try:
                    res_p = run_optim_avec_prio(demand_df_p, bom_df_p, stock_df_p, prio_df)
                    st.session_state['res_avec_prio'] = res_p
                    st.session_state['avec_prio_done'] = True
                except Exception as e:
                    st.error(f"Erreur optimisation : {e}")
                    st.session_state['avec_prio_done'] = False

    if st.session_state.get('avec_prio_done'):
        res_p = st.session_state['res_avec_prio']
        produced_p = res_p['produced']
        demand_df_prio = res_p['demand_df']
        total_demand_p = demand_df_prio['Demande'].sum()
        configs_produites_p = sum(produced_p.values())
        pct_p = (configs_produites_p / total_demand_p * 100) if total_demand_p > 0 else 0

        st.subheader("Resultats globaux")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Configs produites", configs_produites_p)
        c2.metric("Configs demandees", int(total_demand_p))
        c3.metric("% Realisation", f"{pct_p:.1f} %")
        c4.metric("Cout stock consomme", f"{res_p['cout_total']:,.2f} EUR")

        # Log phases
        with st.expander("Log d'execution"):
            for line in res_p['log_lines']:
                st.text(line)

        # Detail par priorité
        st.subheader("Detail par priorite")
        if 'Priorité' in demand_df_prio.columns:
            prio_summary = demand_df_prio.groupby('Priorité').agg(
                Nb_configs=('Conf|version', 'count'),
                Demande_totale=('Demande', 'sum'),
            ).reset_index()
            prio_summary['Produit_total'] = prio_summary.apply(
                lambda row: sum(produced_p.get(cv, 0) for cv in demand_df_prio[demand_df_prio['Priorité'] == row['Priorité']]['Conf|version']),
                axis=1
            )
            prio_summary['Taux %'] = (prio_summary['Produit_total'] / prio_summary['Demande_totale'] * 100).round(1)
            st.dataframe(prio_summary, width=None, hide_index=True)

        # Detail par config
        st.subheader("Detail par configuration")
        detail_p = demand_df_prio[['Conf|version', 'Priorité', 'Demande']].copy()
        detail_p['Qté produite'] = detail_p['Conf|version'].map(produced_p).fillna(0).astype(int)
        detail_p['Ecart'] = detail_p['Demande'] - detail_p['Qté produite']
        detail_p['% Realise'] = (detail_p['Qté produite'] / detail_p['Demande'] * 100).round(1)
        st.dataframe(detail_p, width=None, hide_index=True)

        excel_bytes_p = generate_excel_avec_prio(demand_df_p, bom_df_p, stock_df_p, res_p)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Telecharger le fichier Excel",
            data=excel_bytes_p,
            file_name=f"Optim_Force_Prio_Max_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =============================================================================
# FOOTER
# =============================================================================
st.divider()
st.caption("FORECAST SFR - Batail-log | Python + PuLP (CBC Solver)")

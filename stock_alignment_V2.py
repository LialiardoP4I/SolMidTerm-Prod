# -*- coding: utf-8 -*-
"""
================================================================================
STOCK ALIGNMENT & DEMAND SCENARIO COMPARISON V2
================================================================================

Confronta la domanda per SKU sotto due scenari di Take Rate (actual vs forecast)
per quantificare l'eccesso o la carenza di scorte causata dalla differenza
tra TR previsti e TR effettivi (Mix-Driven Excess).

DIFFERENZE RISPETTO A main_demand_comparison_V2.py:
  - Nessuna dipendenza dal file residual (quantity e residual.xlsx)
  - n_months è un parametro esplicito (non derivato dal max lead time)
  - Tutti gli SKU BOM entrano (nessun filtro per lead time)
  - L'aggregazione domanda usa n_months uguale per tutti gli SKU
  - Output arricchito con metadati BOM (Livello, Tipo approvv., Versione)
    per filtraggio interattivo in Excel

OUTPUT Excel — un unico file multi-foglio:
  - "Risultati":    tutti gli SKU con excess/shortage e valore
  - "Actual":       domanda media raw per TR_actual
  - "Forecast":     domanda media raw per TR_forecast

================================================================================
@author: AntonioLiardo-SMOPS
================================================================================
"""

import time
import gc
import psutil
import os
import numpy as np
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional

# ============================================================================
# IMPORTAZIONE MODULI (read-only — non modificati)
# ============================================================================
from tr_parser_V2 import parse_tr_file, load_monthly_tr
from mc_simulator_V2 import (
    SimulationConfig,
    load_monthly_forecast,
    parse_exclusions,
    run_monthly_simulations_optimized,
    verify_demand_conservation,
    build_wide_dataframe,
)
from sku_matching_V2 import (
    prenormalize_simulation_output_fast,
    sku_matches_combination_numpy,
    parse_month_to_sortable,
    get_memory_usage_mb,
    build_column_mapping_auto,
    _reinit_norm_cache,
)


# ============================================================================
# UTILITY
# ============================================================================

def _get_memory_mb() -> float:
    process = psutil.Process(os.getpid())
    return process.memory_info().rss / 1024 / 1024


def _load_prezzi(filepath: str = ".\\Input\\Prezzi.xlsx") -> pd.DataFrame:
    """
    Carica e normalizza prezzi unitari in EUR indicizzati per Materiale.

    Logica:
    - Legge Sheet1: Materiale, Prezzo netto, Unità di prezzo, Divisa, Data documento
    - Calcola prezzo unitario: Prezzo netto / Unità di prezzo
    - Converte in EUR con 4 livelli di fallback per il cambio valuta:
        1. File 'Cambio Valuta.xlsx' separato
        2. Foglio con "cambio"/"valuta"/"eur" nel nome dentro Prezzi.xlsx
        3. Stima da Sheet1 (assume tasso=1.0 per tutte le valute)
        4. Default EUR=1.0 per valute mancanti
    - Per ogni Materiale prende il prezzo più recente (Data documento più alta)

    Returns:
        DataFrame con colonne [Materiale, Prezzo_Unitario], indicizzato su Materiale.
    """
    if not os.path.exists(filepath):
        print(f"  [WARN] File prezzi non trovato: {filepath}")
        return pd.DataFrame(columns=["Materiale", "Prezzo_Unitario"]).set_index("Materiale")

    all_data = pd.read_excel(filepath, sheet_name='Sheet1')

    material_col = "Materiale"
    prezzo_col = "Prezzo netto"
    data_doc_col = "Data documento"
    unit_prezzo_col = None
    divisa_col = None

    # Identifica colonne con caratteri speciali (come in main_demand_comparison_V2)
    for col in all_data.columns:
        col_lower = str(col).lower().replace('\xa0', ' ')
        if 'unit' in col_lower and 'prezzo' in col_lower:
            unit_prezzo_col = col
        if 'divisa' in col_lower or 'currency' in col_lower:
            divisa_col = col

    if unit_prezzo_col is None or divisa_col is None:
        print(f"  [WARN] Colonne non trovate (unit_prezzo={unit_prezzo_col}, divisa={divisa_col}) → prezzo omesso")
        return pd.DataFrame(columns=["Materiale", "Prezzo_Unitario"]).set_index("Materiale")

    df = all_data[[material_col, prezzo_col, unit_prezzo_col, divisa_col, data_doc_col]].drop_duplicates()
    df["Materiale"] = df[material_col].astype(str).str.strip()
    df["Prezzo_Unitario"] = pd.to_numeric(df[prezzo_col], errors='coerce') / \
                            pd.to_numeric(df[unit_prezzo_col], errors='coerce').replace(0, np.nan)

    # ── Cambio valuta con 4 livelli di fallback ──
    cambio = None

    # Fallback 1: file Cambio Valuta.xlsx separato
    cambio_path = os.path.join(os.path.dirname(filepath), "Cambio Valuta.xlsx")
    try:
        if os.path.exists(cambio_path):
            cambio_temp = pd.read_excel(cambio_path)
            if len(cambio_temp) > 0 and len(cambio_temp.columns) > 0:
                code_col = rate_col = None
                for col in cambio_temp.columns:
                    cl = str(col).lower()
                    if "codice" in cl or "currency" in cl or "divisa" in cl:
                        code_col = col
                    if "eur" in cl or "tasso" in cl or "rate" in cl:
                        rate_col = col
                if code_col and rate_col:
                    cambio = cambio_temp[[code_col, rate_col]].rename(
                        columns={code_col: "Codice", rate_col: "tasso_eur"}
                    )
                    print(f"  [INFO] Tassi di cambio da Cambio Valuta.xlsx")
    except Exception as e:
        print(f"  [WARN] Errore lettura Cambio Valuta.xlsx: {e}")

    # Fallback 2: foglio con "cambio"/"valuta"/"eur" nel nome dentro Prezzi.xlsx
    if cambio is None or len(cambio) == 0:
        try:
            for sheet_name in pd.ExcelFile(filepath).sheet_names:
                if any(k in sheet_name.lower() for k in ("cambio", "valuta", "eur")):
                    tmp = pd.read_excel(filepath, sheet_name=sheet_name)
                    code_col = rate_col = None
                    for col in tmp.columns:
                        cl = str(col).lower()
                        if "codice" in cl or "currency" in cl or "divisa" in cl:
                            code_col = col
                        if "eur" in cl or "tasso" in cl or "rate" in cl:
                            rate_col = col
                    if code_col and rate_col:
                        cambio = tmp[[code_col, rate_col]].rename(
                            columns={code_col: "Codice", rate_col: "tasso_eur"}
                        )
                        print(f"  [INFO] Tassi di cambio dal foglio '{sheet_name}' in Prezzi.xlsx")
                        break
        except Exception:
            pass

    # Fallback 3: stima da Sheet1 (assume tasso=1.0 per tutte le valute)
    if cambio is None or len(cambio) == 0:
        valute = df[divisa_col].dropna().unique()
        cambio = pd.DataFrame({"Codice": valute, "tasso_eur": 1.0})
        print(f"  [INFO] Tassi di cambio non trovati → tasso=1.0 per tutte le valute ({list(valute)})")

    # Applica il cambio
    df = df.merge(cambio, left_on=divisa_col, right_on="Codice", how="left")
    n_missing = df["tasso_eur"].isna().sum()
    if n_missing > 0:
        print(f"  [WARN] {n_missing} righe con valuta non trovata → tasso=1.0 (EUR)")
    df["tasso_eur"] = df["tasso_eur"].fillna(1.0)
    df["Prezzo_Unitario"] = df["Prezzo_Unitario"] * df["tasso_eur"]

    # Prezzo più recente per Materiale
    df = (
        df[["Materiale", data_doc_col, "Prezzo_Unitario"]]
        .dropna(subset=["Prezzo_Unitario"])
        .sort_values(data_doc_col)
        .groupby("Materiale", as_index=True)[["Prezzo_Unitario"]]
        .last()
    )
    print(f"  Prezzi caricati: {len(df)} materiali")
    return df


# ============================================================================
# METADATI BOM PER FILTRAGGIO INTERATTIVO
# ============================================================================

def _load_bom_metadata() -> pd.DataFrame:
    """
    Carica metadati BOM (Livello, Tipo approvv., Descrizione, Check Generale)
    dalle tre BOM arricchite, per arricchire l'output finale.
    """
    base = Path(".\\Input")
    frames = []

    for ver in ["V21E", "V22E", "V24T"]:
        bom_path = base / ver / f"BOM_150{ver}.xlsx"
        if not bom_path.exists():
            continue

        df = pd.read_excel(bom_path, usecols=lambda c: c in [
            "Numero componenti", "Livello numerico", "Tipo approvv.",
            "Testo breve oggetto", "Check Generale", "Car. MRP",
        ])
        df["BOM_Versione"] = ver
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    all_bom = pd.concat(frames, ignore_index=True)
    all_bom["Numero componenti"] = all_bom["Numero componenti"].astype(str).str.strip()

    # Per ogni componente, prendi il livello minimo (più vicino alla radice)
    # e il tipo approvv. più comune
    meta = all_bom.groupby("Numero componenti").agg(
        Livello_Min=("Livello numerico", "min"),
        Livello_Max=("Livello numerico", "max"),
        Tipo_Approvv=("Tipo approvv.", "first"),
        Descrizione_BOM=("Testo breve oggetto", "first"),
        Car_MRP=("Car. MRP", "first"),
        BOM_Versioni=("BOM_Versione", lambda x: ",".join(sorted(set(x)))),
    ).reset_index()

    print(f"  Metadati BOM: {len(meta)} componenti da {len(frames)} BOM")
    return meta




# ============================================================================
# SIMULAZIONE MC CON TR PARAMETRIZZABILE
# ============================================================================

def _run_simulation(
    tr_path: str,
    n_months: int,
    n_runs: int,
    start_month_offset: int = 0,
    label: str = "",
) -> pd.DataFrame:
    """
    Esegue la simulazione Monte Carlo con il file TR specificato.
    n_months è esplicito (non derivato dal residual).
    """
    sep = "=" * 60
    print(f"\n{sep}")
    print(f"SIMULAZIONE MC - TR {label}")
    print(f"  File TR: {tr_path}")
    print(f"  Mesi: {n_months}")
    print(f"{sep}")

    config = SimulationConfig(
        n_runs=n_runs,
        random_seed=42,
        start_month_offset=start_month_offset,
        use_int16=True,
    )

    # Forecast domanda
    forecast_values, month_names = load_monthly_forecast(
        ".\\Input\\Total_demand.xlsx",
        start_offset=config.start_month_offset,
    )

    # Limita ai mesi richiesti
    actual_months = min(n_months, len(forecast_values))
    if actual_months < n_months:
        print(f"  [WARN] Richiesti {n_months} mesi ma disponibili solo {actual_months}")

    # Carica TR direttamente (come Safety Stock — senza normalizzazione)
    # Nota: il MC simulator ha già logica per gestire ALTRO (righe 727-743)
    model_mix, characteristics = parse_tr_file(tr_path)
    print(f"  Modelli: {len(model_mix.models)}")
    print(f"  Caratteristiche: {len(characteristics)}")

    monthly_tr = None
    try:
        monthly_tr = load_monthly_tr(tr_path)
        print(f"  TR mensili: {len(monthly_tr)} mesi")
    except Exception:
        print("  Nessun TR mensile → uso TR globali")

    # Esclusioni
    exclusions = parse_exclusions(".\\Input\\esclusioni.xlsx")

    # Simulazione
    print(f"\n  Simulazione ({n_runs} run, {actual_months} mesi)...")
    try:
        results_by_month = run_monthly_simulations_optimized(
            forecast_values=forecast_values,
            month_names=month_names,
            n_months=actual_months,
            model_mix=model_mix,
            characteristics=characteristics,
            exclusions=exclusions,
            config=config,
            monthly_tr=monthly_tr,
        )
    except ValueError as e:
        if "probabilities do not sum to 1" in str(e):
            print(f"\n[DIAG] CRASH: {e}")
            print(f"[DIAG] Model mix models: {model_mix.models}")
            print(f"[DIAG] Model mix take_rates: {model_mix.take_rates}")
            print(f"[DIAG] Caratteristiche ({len(characteristics)}):")
            for cname, cg in list(characteristics.items())[:5]:
                print(f"  - '{cname}': {len(cg.values)} opzioni, TR keys={list(cg.tr.keys())[:3]}")
                for model_key in list(cg.tr.keys())[:1]:
                    tr_arr = cg.tr[model_key]
                    print(f"    {model_key}: TR={tr_arr}, sum={tr_arr.sum():.4f}, any_nan={np.isnan(tr_arr).any()}")
            if monthly_tr:
                first_month = next(iter(monthly_tr.keys()))
                mm, chars = monthly_tr[first_month]
                print(f"[DIAG] Monthly TR ({first_month}): models={mm.models}")
        raise

    verify_demand_conservation(
        results_by_month=results_by_month,
        n_runs=config.n_runs,
        forecast_values=forecast_values[:actual_months],
    )

    df_wide = build_wide_dataframe(
        results_by_month=results_by_month,
        month_names=month_names,
        n_runs=config.n_runs,
        save_to_file=False,
    )

    gc.collect()
    print(f"\n[OK] Simulazione TR {label} completata — {len(df_wide)} combinazioni")
    return df_wide


# ============================================================================
# CALCOLO DOMANDA MEDIA PER SKU (SENZA FILTRO RESIDUAL)
# ============================================================================

def compute_mean_demands_all(
    simulation_output: pd.DataFrame,
    sku_catalog: pd.DataFrame,
    column_mapping: Dict[str, str],
    n_runs: int,
    n_months: int,
    use_supermodel_matching: bool = True,
) -> pd.DataFrame:
    """
    Calcola la domanda media per ogni SKU su n_months fissi.

    DIFFERENZE rispetto a compute_mean_demands in main_demand_comparison_V2:
    - Nessun join/filtro con lead_times
    - Tutti gli SKU del catalogo entrano
    - n_months uguale per tutti gli SKU (non lead_time individuale)
    - Matching con chiave composta (SKU, Supermodel) se use_supermodel_matching=True

    Args:
        use_supermodel_matching: Se True, usa matching con supermodel
    """
    sep = "-" * 60
    print(f"\n{sep}")
    print("CALCOLO DOMANDA MEDIA PER SKU (Stock Alignment)")
    print(f"  n_months = {n_months}")
    print(f"{sep}")
    print(f"Memoria iniziale: {get_memory_usage_mb():.1f} MB")

    # Normalizzazione
    _reinit_norm_cache(column_mapping)
    normalized_df, norm_arrays_raw = prenormalize_simulation_output_fast(
        simulation_output, column_mapping
    )

    from sku_matching_V2 import _norm_cache
    category_maps = {}
    norm_arrays = {}
    for col_name in norm_arrays_raw.keys():
        if col_name in normalized_df.columns:
            cat_col = normalized_df[col_name]
            if hasattr(cat_col, "cat"):
                category_maps[col_name] = {
                    cat: code for code, cat in enumerate(cat_col.cat.categories)
                }
                norm_arrays[col_name] = cat_col.cat.codes.values

    n_rows = len(normalized_df)

    # Matrice numpy delle colonne run
    month_run_columns = [col for col in simulation_output.columns if "_Run_" in col]
    month_prefixes = []
    seen = set()
    for col in month_run_columns:
        prefix = col.split("_Run_")[0]
        if prefix not in seen:
            month_prefixes.append(prefix)
            seen.add(prefix)
    month_prefixes.sort(key=parse_month_to_sortable)

    # Limita ai primi n_months
    month_prefixes = month_prefixes[:n_months]

    run_columns = []
    for prefix in month_prefixes:
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in simulation_output.columns:
                run_columns.append(col_name)

    data_matrix = simulation_output[run_columns].values.astype(np.int16)

    # Pre-calcola indici colonne per mese
    run_col_to_idx = {col: idx for idx, col in enumerate(run_columns)}
    month_col_indices = {}
    for month_idx, prefix in enumerate(month_prefixes):
        indices = []
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in run_col_to_idx:
                indices.append(run_col_to_idx[col_name])
        month_col_indices[month_idx] = indices

    # Prepara SKU — SENZA filtro lead time
    sku_col = "Numero componenti"
    sku_catalog = sku_catalog.copy()
    sku_catalog[sku_col] = sku_catalog[sku_col].astype(str).str.strip()

    sku_groups = sku_catalog.groupby(sku_col)
    n_sku_groups = len(sku_groups)
    print(f"SKU da processare: {n_sku_groups}")

    # Colonna Model per multi-model expansion
    model_sku_col = next(
        (c for c in column_mapping.keys() if c.lower().strip() == "model"),
        None,
    )

    # Pre-estrai dati per ogni SKU
    sku_ids = []
    sku_data_list = []
    for sku_id, sku_rows_group in sku_groups:
        sku_ids.append(sku_id)
        rows_values = []
        for _, row in sku_rows_group.iterrows():
            row_values = {
                sc: row[sc]
                for sc in column_mapping.keys()
                if sc in row.index
            }
            # Multi-model expansion: "MSV4+MSV4S" → due entry separate
            if model_sku_col and model_sku_col in row_values:
                model_val = str(row_values[model_sku_col]).strip()
                if "+" in model_val:
                    for single_model in model_val.split("+"):
                        expanded = dict(row_values)
                        expanded[model_sku_col] = single_model.strip()
                        rows_values.append(expanded)
                else:
                    rows_values.append(row_values)
            else:
                rows_values.append(row_values)
        sku_data_list.append(rows_values)

    # Loop principale: matching + media
    print(f"Matching SKU... (ogni '.' = 50 SKU)")
    start_time = time.time()
    results = []

    for sku_idx, (sku_id, sku_rows) in enumerate(zip(sku_ids, sku_data_list)):
        if sku_idx % 50 == 0:
            print(".", end="", flush=True)
        if sku_idx % 500 == 0 and sku_idx > 0:
            elapsed = time.time() - start_time
            rate = sku_idx / elapsed
            remaining = (n_sku_groups - sku_idx) / rate if rate > 0 else 0
            print(f" [{sku_idx}/{n_sku_groups}] {rate:.1f} SKU/s, ETA: {remaining/60:.1f}min")

        # Matching: OR tra le righe dello SKU
        combined_mask = np.zeros(n_rows, dtype=bool)
        for row_values in sku_rows:
            row_mask = sku_matches_combination_numpy(
                row_values, norm_arrays, category_maps, column_mapping, n_rows
            )
            combined_mask |= row_mask

        matching_indices = np.where(combined_mask)[0]
        if len(matching_indices) == 0:
            continue

        # Aggregazione domanda su TUTTI i mesi (uguale per ogni SKU)
        aggregated = np.zeros(n_runs, dtype=np.float32)
        for month_idx in range(len(month_prefixes)):
            col_indices = month_col_indices.get(month_idx, [])
            if not col_indices:
                continue
            month_data = data_matrix[np.ix_(matching_indices, col_indices)]
            aggregated += month_data.sum(axis=0).astype(np.float32)

        mean_demand = float(np.mean(aggregated))

        results.append({
            "SKU": sku_id,
            "N_Matching_Combinations": int(len(matching_indices)),
            "mean_demand": round(mean_demand, 2),
        })

        if sku_idx % 1000 == 0:
            gc.collect()

    elapsed = time.time() - start_time
    print(f"\n\n[OK] Completato in {elapsed:.1f}s — {len(results)} SKU con match")
    return pd.DataFrame(results)


# ============================================================================
# PIPELINE PRINCIPALE
# ============================================================================

def main(
    tr_actual_path: str = ".\\Input\\TR TOTALV21E - actual - Copia.xlsx",
    tr_forecast_path: str = ".\\Input\\TR TOTALV21E - forecast - Copia.xlsx",
    n_months: int = 3,
    n_runs: int = 200,
    start_month_offset: int = 0,
    output_path: str = ".\\Output\\stock_alignment_V2.xlsx",
    catalog_path: str = ".\\Input\\tutte_righe_univoche_SA.xlsx",
):
    """
    Pipeline Stock Alignment: due simulazioni (actual vs forecast) + confronto.

    Args:
        tr_actual_path:     File TR effettivi
        tr_forecast_path:   File TR previsionali
        n_months:           Mesi su cui aggregare la domanda (parametro esplicito)
        n_runs:             Numero run Monte Carlo per ogni simulazione
        start_month_offset: Offset mese nel file forecast
        output_path:        Path file output Excel
        catalog_path:       Catalogo SKU SA (tutte_righe_univoche_SA.xlsx)
                            Se non trovato, fallback su tutte_righe_univoche_V2.xlsx
    """
    total_start = time.time()
    print(f"\n{'='*70}")
    print("STOCK ALIGNMENT & DEMAND SCENARIO COMPARISON V2")
    print(f"{'='*70}")
    print(f"  TR Actual:   {tr_actual_path}")
    print(f"  TR Forecast: {tr_forecast_path}")
    print(f"  N. mesi:     {n_months}")
    print(f"  N. run MC:   {n_runs}")
    print(f"  Memoria:     {_get_memory_mb():.1f} MB")

    # =========================================================================
    # STEP 1: Simulazione con TR_actual
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 1: SIMULAZIONE MC — TR ACTUAL")
    print(f"{'='*70}")

    t1 = time.time()
    df_actual = _run_simulation(
        tr_path=tr_actual_path,
        n_months=n_months,
        n_runs=n_runs,
        start_month_offset=start_month_offset,
        label="ACTUAL",
    )
    print(f"Step 1 completato in {time.time() - t1:.1f}s")
    gc.collect()

    # =========================================================================
    # STEP 2: Simulazione con TR_forecast
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 2: SIMULAZIONE MC — TR FORECAST")
    print(f"{'='*70}")

    t2 = time.time()
    df_forecast = _run_simulation(
        tr_path=tr_forecast_path,
        n_months=n_months,
        n_runs=n_runs,
        start_month_offset=start_month_offset,
        label="FORECAST",
    )
    print(f"Step 2 completato in {time.time() - t2:.1f}s")
    gc.collect()

    # =========================================================================
    # STEP 3: Caricamento catalogo SKU + metadati BOM
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 3: CARICAMENTO DATI SKU + METADATI BOM")
    print(f"{'='*70}")

    # Catalogo SA dedicato (con livelli BOM); fallback su catalogo Safety Stock
    _catalog_file = catalog_path
    if not os.path.exists(_catalog_file):
        _fallback = ".\\Input\\tutte_righe_univoche_V2.xlsx"
        print(f"  [WARN] Catalogo SA non trovato ({_catalog_file})")
        print(f"  [WARN] Fallback su: {_fallback}")
        print(f"  [WARN] Esegui 'crea_catalogo_sa.py' per generare il catalogo dedicato.")
        _catalog_file = _fallback
    print(f"  Catalogo: {_catalog_file}")
    sku_catalog_raw = pd.read_excel(_catalog_file)
    if "ID" in sku_catalog_raw.columns:
        sku_catalog_raw = sku_catalog_raw.drop(columns=["ID"])
    sku_catalog = sku_catalog_raw.drop_duplicates(keep="first")
    print(f"  SKU catalogo: {len(sku_catalog)}")
    print(f"  SKU univoci:  {sku_catalog['Numero componenti'].nunique()}")

    # Metadati BOM per arricchimento output
    bom_meta = _load_bom_metadata()

    # Auto-rileva COLUMN_MAPPING
    sim_cols = [c for c in df_actual.columns if "_Run_" not in c]
    column_mapping = build_column_mapping_auto(
        sku_cols=list(sku_catalog.columns),
        sim_cols=sim_cols,
    )
    print(f"  COLUMN_MAPPING auto: {len(column_mapping)} colonne")

    # Identifica n_runs dal DataFrame
    month_run_cols = [c for c in df_actual.columns if "_Run_" in c]
    if month_run_cols:
        first_prefix = month_run_cols[0].split("_Run_")[0]
        n_runs_actual = len([c for c in month_run_cols if c.startswith(first_prefix + "_Run_")])
    else:
        n_runs_actual = n_runs

    month_run_cols_f = [c for c in df_forecast.columns if "_Run_" in c]
    if month_run_cols_f:
        first_prefix_f = month_run_cols_f[0].split("_Run_")[0]
        n_runs_forecast = len([c for c in month_run_cols_f if c.startswith(first_prefix_f + "_Run_")])
    else:
        n_runs_forecast = n_runs

    # =========================================================================
    # STEP 4: Calcolo domanda media — ACTUAL
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 4: DOMANDA MEDIA SKU — TR ACTUAL")
    print(f"{'='*70}")

    t4 = time.time()

    # Usa matching con supermodel per Stock Alignment
    try:
        from sku_matching_SA import main_sku_sa
        print("\n  [SA] Usando matching con supermodel (sku_matching_SA)...")
        df_mean_actual_sa = main_sku_sa(
            simulation_output=df_actual,
            sku_catalog_path=_catalog_file,
            n_months=n_months,
        )
        # Rinomina per compatibilità
        df_mean_actual_sa = df_mean_actual_sa.rename(
            columns={"mean_demand": "mean_demand_actual"}
        )
        df_mean_actual = df_mean_actual_sa
    except Exception as e:
        print(f"  [WARN] Matching con supermodel fallito ({e}), uso versione standard...")
        df_mean_actual = compute_mean_demands_all(
            simulation_output=df_actual,
            sku_catalog=sku_catalog,
            column_mapping=column_mapping,
            n_runs=n_runs_actual,
            n_months=n_months,
        )
    print(f"Step 4 completato in {time.time() - t4:.1f}s")
    del df_actual
    gc.collect()

    # =========================================================================
    # STEP 5: Calcolo domanda media — FORECAST
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 5: DOMANDA MEDIA SKU — TR FORECAST")
    print(f"{'='*70}")

    t5 = time.time()

    # Usa matching con supermodel per Forecast
    try:
        from sku_matching_SA import main_sku_sa
        df_mean_forecast_sa = main_sku_sa(
            simulation_output=df_forecast,
            sku_catalog_path=_catalog_file,
            n_months=n_months,
        )
        df_mean_forecast_sa = df_mean_forecast_sa.rename(
            columns={"mean_demand": "mean_demand_forecast"}
        )
        df_mean_forecast = df_mean_forecast_sa
    except Exception as e:
        print(f"  [WARN] Matching forecast fallito, uso versione standard...")
        df_mean_forecast = compute_mean_demands_all(
            simulation_output=df_forecast,
            sku_catalog=sku_catalog,
            column_mapping=column_mapping,
            n_runs=n_runs_forecast,
            n_months=n_months,
        )
    print(f"Step 5 completato in {time.time() - t5:.1f}s")
    del df_forecast
    gc.collect()

    # =========================================================================
    # STEP 6: Merge, metadati, prezzi, calcolo excess/shortage
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 6: CALCOLO EXCESS / SHORTAGE + ARRICCHIMENTO")
    print(f"{'='*70}")

    # Merge actual e forecast — supporta chiave composta (SKU, Supermodel)
    actual_cols = ["SKU", "mean_demand_actual"]
    forecast_cols = ["SKU", "mean_demand_forecast"]
    merge_keys = ["SKU"]

    if "Supermodel" in df_mean_actual.columns and "Supermodel" in df_mean_forecast.columns:
        actual_cols.insert(1, "Supermodel")
        forecast_cols.insert(1, "Supermodel")
        merge_keys.append("Supermodel")
        print(f"  [SA] Merge con chiave composta: {merge_keys}")
    elif "N_Matching_Combinations" in df_mean_actual.columns:
        actual_cols.insert(1, "N_Matching_Combinations")

    df_merged = df_mean_actual[actual_cols].merge(
        df_mean_forecast[forecast_cols],
        on=merge_keys,
        how="outer",
        suffixes=("_actual", "_forecast"),
    )

    df_merged["mean_demand_actual"] = df_merged["mean_demand_actual"].fillna(0.0).round(2)
    df_merged["mean_demand_forecast"] = df_merged["mean_demand_forecast"].fillna(0.0).round(2)
    if "N_Matching_Combinations" in df_merged.columns:
        df_merged["N_Matching_Combinations"] = df_merged["N_Matching_Combinations"].fillna(0).astype(int)
    else:
        df_merged["N_Matching_Combinations"] = 0

    # Delta, Excess, Shortage
    df_merged["delta_demand"] = (
        df_merged["mean_demand_forecast"] - df_merged["mean_demand_actual"]
    ).round(2)
    df_merged["excess"] = df_merged["delta_demand"].clip(lower=0).round(2)
    df_merged["shortage"] = (-df_merged["delta_demand"]).clip(lower=0).round(2)

    # Arricchisci con metadati BOM
    if len(bom_meta) > 0:
        df_merged = df_merged.merge(
            bom_meta,
            left_on="SKU",
            right_on="Numero componenti",
            how="left",
        )
        df_merged = df_merged.drop(columns=["Numero componenti"], errors="ignore")

    # Fallback Livello_Min/Max: dal catalogo SA (già calcolati da crea_catalogo_sa.py)
    # Usato quando _load_bom_metadata() non trova i BOM o "Livello numerico" manca nei file
    lev_cols_missing = [c for c in ["Livello_Min", "Livello_Max"]
                        if c not in df_merged.columns or df_merged[c].isna().all()]
    if lev_cols_missing:
        lev_avail = [c for c in lev_cols_missing if c in sku_catalog_raw.columns]
        if lev_avail:
            lev_map = (
                sku_catalog_raw[["Numero componenti"] + lev_avail]
                .drop_duplicates("Numero componenti")
                .rename(columns={"Numero componenti": "SKU"})
            )
            # Drop eventuali colonne già presenti (vuote) prima del merge
            df_merged = df_merged.drop(columns=[c for c in lev_avail if c in df_merged.columns],
                                       errors="ignore")
            df_merged = df_merged.merge(lev_map, on="SKU", how="left")
            print(f"  [INFO] Livello da catalogo SA: {lev_avail}")

    # Arricchisci con Versione da catalogo SKU
    versione_map = (
        sku_catalog[["Numero componenti", "Versione"]]
        .drop_duplicates("Numero componenti")
        .rename(columns={"Numero componenti": "SKU"})
    )
    # Un componente può apparire in più versioni
    versione_multi = (
        sku_catalog.groupby("Numero componenti")["Versione"]
        .apply(lambda x: ",".join(sorted(set(x))))
        .reset_index()
        .rename(columns={"Numero componenti": "SKU", "Versione": "Versioni_SKU"})
    )
    df_merged = df_merged.merge(versione_multi, on="SKU", how="left")

    # Arricchisci con Supermodel (Model) da catalogo SKU
    # Un componente può apparire in più supermodel (MSV4, MSV4S, MSV4PP, Rally, RS, etc.)
    if "Model" in sku_catalog.columns:
        supermodel_multi = (
            sku_catalog.groupby("Numero componenti")["Model"]
            .apply(lambda x: ",".join(sorted(set(str(v) for v in x if pd.notna(v)))))
            .reset_index()
            .rename(columns={"Numero componenti": "SKU", "Model": "Supermodel"})
        )
        df_merged = df_merged.merge(supermodel_multi, on="SKU", how="left")

    # Prezzi
    prezzi = _load_prezzi(".\\Input\\Prezzi.xlsx")

    # DIAGNOSTICA PREZZI
    print(f"\n[DIAG] Caricamento prezzi:")
    print(f"  Righe prezzi caricate: {len(prezzi)}")
    if len(prezzi) > 0:
        print(f"  Prezzo min: EUR {prezzi['Prezzo_Unitario'].min():,.2f}")
        print(f"  Prezzo max: EUR {prezzi['Prezzo_Unitario'].max():,.2f}")
        print(f"  Prezzo media: EUR {prezzi['Prezzo_Unitario'].mean():,.2f}")
        print(f"  Prezzo mediana: EUR {prezzi['Prezzo_Unitario'].median():,.2f}")
        # Top 10 prezzi anomali (> 1000€)
        anomali = prezzi[prezzi['Prezzo_Unitario'] > 1000].sort_values('Prezzo_Unitario', ascending=False)
        if len(anomali) > 0:
            print(f"\n  ⚠️  ANOMALIE RILEVATE: {len(anomali)} SKU con prezzo > EUR 1.000:")
            for idx, (mat, row) in enumerate(anomali.head(5).iterrows()):
                print(f"    - {mat}: EUR {row['Prezzo_Unitario']:,.2f}")
            if len(anomali) > 5:
                print(f"    ... e altri {len(anomali) - 5}")

    df_merged = df_merged.merge(
        prezzi,
        left_on="SKU",
        right_index=True,
        how="left",
    )

    # Valori economici
    df_merged["excess_value"] = (df_merged["excess"] * df_merged["Prezzo_Unitario"]).round(2)
    df_merged["shortage_value"] = (df_merged["shortage"] * df_merged["Prezzo_Unitario"]).round(2)

    # DIAGNOSTICA VALORI ECONOMICI
    print(f"\n[DIAG] Valori economici:")
    print(f"  SKU con prezzo: {df_merged['Prezzo_Unitario'].notna().sum()}")
    print(f"  SKU senza prezzo: {df_merged['Prezzo_Unitario'].isna().sum()}")
    if df_merged["excess_value"].notna().any():
        print(f"  Excess value max: EUR {df_merged['excess_value'].max():,.2f}")
        # SKU con excess > 100k€
        big_excess = df_merged[df_merged["excess_value"] > 100000].sort_values("excess_value", ascending=False)
        if len(big_excess) > 0:
            print(f"\n  ⚠️  SKU con EXCESS > EUR 100.000 ({len(big_excess)}):")
            for _, row in big_excess.head(5).iterrows():
                print(f"    - {row['SKU']}: {row['excess']:.0f} pz × EUR {row['Prezzo_Unitario']:,.2f} = EUR {row['excess_value']:,.2f}")

    # Ordina per |delta_demand| decrescente
    df_merged = (
        df_merged
        .sort_values("delta_demand", key=abs, ascending=False)
        .reset_index(drop=True)
    )

    # =========================================================================
    # STEP 7: Salvataggio output
    # =========================================================================
    print(f"\n{'='*70}")
    print("STEP 7: SALVATAGGIO OUTPUT")
    print(f"{'='*70}")

    os.makedirs(os.path.dirname(output_path) or ".\\Output", exist_ok=True)

    # Colonne output ordinate
    cols_risultati = [
        "SKU", "Descrizione_BOM", "Versioni_SKU", "BOM_Versioni",
        "Livello_Min", "Livello_Max",
        "Tipo_Approvv", "Car_MRP",
        "N_Matching_Combinations",
        "mean_demand_actual", "mean_demand_forecast",
        "delta_demand", "excess", "shortage",
        "Prezzo_Unitario", "excess_value", "shortage_value",
    ]
    # Tieni solo le colonne che esistono
    cols_risultati = [c for c in cols_risultati if c in df_merged.columns]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_merged[cols_risultati].to_excel(
            writer, sheet_name="Risultati", index=False
        )
        actual_sort_col = "mean_demand_actual" if "mean_demand_actual" in df_mean_actual.columns else "mean_demand"
        forecast_sort_col = "mean_demand_forecast" if "mean_demand_forecast" in df_mean_forecast.columns else "mean_demand"
        df_mean_actual.sort_values(actual_sort_col, ascending=False).to_excel(
            writer, sheet_name="Actual", index=False
        )
        df_mean_forecast.sort_values(forecast_sort_col, ascending=False).to_excel(
            writer, sheet_name="Forecast", index=False
        )

    # Stats
    n_excess = (df_merged["excess"] > 0).sum()
    n_shortage = (df_merged["shortage"] > 0).sum()
    n_aligned = ((df_merged["delta_demand"] == 0) | df_merged["delta_demand"].isna()).sum()

    print(f"\n  [OK] {output_path}")
    print(f"  SKU totali:              {len(df_merged)}")
    print(f"  SKU con excess (>0):     {n_excess}")
    print(f"  SKU con shortage (>0):   {n_shortage}")
    print(f"  SKU allineati (delta=0):     {n_aligned}")

    if df_merged["Prezzo_Unitario"].notna().any():
        tot_excess_val = df_merged["excess_value"].sum()
        tot_shortage_val = df_merged["shortage_value"].sum()
        print(f"\n  Valore excess totale:    EUR {tot_excess_val:,.2f}")
        print(f"  Valore shortage totale:  EUR {tot_shortage_val:,.2f}")
        print(f"  Delta netto:             EUR {tot_excess_val - tot_shortage_val:+,.2f}")

    # =========================================================================
    # RIEPILOGO
    # =========================================================================
    total_elapsed = time.time() - total_start
    print(f"\n{'='*70}")
    print("PIPELINE COMPLETATA")
    print(f"{'='*70}")
    print(f"  Tempo totale: {total_elapsed:.1f}s ({total_elapsed/60:.1f} min)")
    print(f"  Memoria finale: {_get_memory_mb():.1f} MB")
    print(f"  Output: {output_path}")
    print(f"\n  Fogli Excel:")
    print(f"    'Risultati' — excess, shortage, valore, metadati BOM (filtrabile)")
    print(f"    'Actual'    — domanda media raw per TR actual")
    print(f"    'Forecast'  — domanda media raw per TR forecast")

    return df_merged


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    # -------------------------------------------------------------------------
    # CONFIGURAZIONE
    # -------------------------------------------------------------------------
    TR_ACTUAL_PATH   = ".\\Input\\TR TOTALV21E - actual - Copia.xlsx"
    TR_FORECAST_PATH = ".\\Input\\TR TOTALV21E - forecast - Copia.xlsx"
    N_MONTHS         = 3       # Mesi su cui aggregare la domanda
    N_RUNS           = 200     # Run Monte Carlo per scenario
    START_OFFSET     = 0       # Offset mese iniziale

    main(
        tr_actual_path=TR_ACTUAL_PATH,
        tr_forecast_path=TR_FORECAST_PATH,
        n_months=N_MONTHS,
        n_runs=N_RUNS,
        start_month_offset=START_OFFSET,
    )

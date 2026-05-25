# -*- coding: utf-8 -*-
"""
================================================================================
SKU MATCHING PER STOCK ALIGNMENT — Chiave Composta (SKU, Supermodel)
================================================================================

DIFFERENZE rispetto a sku_matching_V2.py:
  1. Nessun calcolo percentili/safety stock
  2. Chiave composta: (Numero componenti, Supermodel)
  3. Aggregazione domanda media su n_months fissi (non per lead time)
  4. Output semplice: domanda media actual/forecast per ogni (SKU, Supermodel)

LOGICA:
  - Ogni riga del catalogo SA ha (SKU, Supermodel)
  - Ogni configurazione simulata ha un Model (MSV4, MSV4S, Rally, RS, etc.)
  - Match: SKU + tutte le caratteristiche corrispondono + Model corrispondente
  - Aggregazione: media domanda per (SKU, Supermodel)

@author: AntonioLiardo-SMOPS
================================================================================
"""

import numpy as np
import pandas as pd
from typing import Dict, List, Tuple
import sys
import os
import time
from datetime import datetime


# ============================================================================
# UTILITY CONDIVISE (precedentemente in sku_matching_V2 / sku_matching_SS)
# ============================================================================
# Duplicate qui per rendere il modulo SA autonomo. Vedi sku_matching_SS.py
# per la documentazione completa di queste funzioni.
# ============================================================================

def build_column_mapping_auto(sku_cols: List[str],
                               sim_cols: List[str]) -> Dict[str, str]:
    """Mapping automatico colonne SKU -> simulazione (case-insensitive)."""
    EXCLUDE = {'id', 'numero componenti', 'numero componenti.1',
               'numero componenti.2', 'numero componenti.3'}
    sim_lower_map = {c.lower().strip(): c for c in sim_cols}
    mapping = {}
    for sku_col in sku_cols:
        norm = sku_col.lower().strip()
        if norm in EXCLUDE:
            continue
        sim_match = sim_lower_map.get(norm)
        if sim_match:
            mapping[sku_col] = sim_match
    return mapping


def get_memory_usage_mb():
    """Memoria RAM processo corrente in MB. 0.0 se psutil non disponibile."""
    try:
        import psutil
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024
    except ImportError:
        return 0.0


class NormalizationCache:
    """Cache normalizzazione stringhe con lookup O(1)."""

    def __init__(self, prefixes: List[str] = None):
        self.cache = {}
        self.prefixes = prefixes if prefixes is not None else []

    def normalize(self, value, column_name: str = None) -> str:
        cache_key = (value, column_name)
        if cache_key in self.cache:
            return self.cache[cache_key]
        if pd.isna(value):
            result = "no"
        else:
            value_str = str(value).strip().lower()
            if value_str == "nessuno":
                result = "no"
            else:
                for prefix in self.prefixes:
                    if value_str.startswith(prefix):
                        value_str = value_str[len(prefix):]
                        break
                if column_name:
                    prefix = column_name.lower() + " - "
                    if value_str.startswith(prefix):
                        value_str = value_str[len(prefix):]
                value_str = value_str.replace(" - ", "-")
                value_str = value_str.replace(" (", "(").replace("( ", "(")
                value_str = value_str.replace(") ", ")").replace(" )", ")")
                value_str = value_str.replace("spoked-black", "spoked black")
                result = value_str
        self.cache[cache_key] = result
        return result

    def preload_column(self, series: pd.Series, column_name: str = None):
        unique_values = series.dropna().unique()
        for val in unique_values:
            self.normalize(val, column_name)


_norm_cache = NormalizationCache()


def normalize_value_cached(value, column_name: str = None) -> str:
    """Wrapper pubblico per normalizzazione con cache (istanza globale)."""
    return _norm_cache.normalize(value, column_name)


def _reinit_norm_cache(column_mapping: Dict[str, str]) -> None:
    """Re-inizializza _norm_cache con i prefissi derivati dal column_mapping."""
    global _norm_cache
    prefixes = ['bundle - '] + [v.lower().strip() + ' - '
                                 for v in column_mapping.values()]
    _norm_cache = NormalizationCache(prefixes=prefixes)
    print(f"[NormCache SA] Re-inizializzata con {len(prefixes)} prefissi dinamici")


def prenormalize_simulation_output_fast(
    simulation_output: pd.DataFrame,
    column_mapping: Dict[str, str]
) -> Tuple[pd.DataFrame, Dict[str, np.ndarray]]:
    """Pre-normalizza colonne simulazione, ritorna (df, dict array numpy)."""
    print(f"\n[PRE-NORM] Inizio pre-normalizzazione veloce...")
    sys.stdout.flush()
    start = time.time()
    normalized_df = simulation_output
    norm_arrays = {}
    for i, (sku_col, combo_col) in enumerate(column_mapping.items(), 1):
        if combo_col in normalized_df.columns:
            _norm_cache.preload_column(normalized_df[combo_col], combo_col)
            norm_col_name = f"{combo_col}_NORM"
            unique_vals = normalized_df[combo_col].dropna().unique()
            norm_map = {v: normalize_value_cached(v, combo_col) for v in unique_vals}
            norm_map[np.nan] = "no"
            normalized_df[norm_col_name] = normalized_df[combo_col].map(
                lambda x: norm_map.get(x, "no") if pd.notna(x) else "no"
            )
            normalized_df[norm_col_name] = normalized_df[norm_col_name].astype('category')
            norm_arrays[norm_col_name] = normalized_df[norm_col_name].cat.codes.values
    elapsed = time.time() - start
    print(f"[PRE-NORM] Completato in {elapsed:.2f}s")
    sys.stdout.flush()
    return normalized_df, norm_arrays


def sku_matches_combination_numpy(
    sku_row_values: Dict[str, str],
    norm_arrays: Dict[str, np.ndarray],
    category_maps: Dict[str, Dict[str, int]],
    column_mapping: Dict[str, str],
    n_rows: int,
) -> np.ndarray:
    """Maschera bool: True per configurazioni che matchano lo SKU (numpy vettoriale)."""
    final_mask = np.ones(n_rows, dtype=bool)
    for sku_col, combo_col in column_mapping.items():
        if sku_col not in sku_row_values:
            continue
        norm_col = f"{combo_col}_NORM"
        if norm_col not in norm_arrays:
            continue
        sku_value = normalize_value_cached(sku_row_values[sku_col], sku_col)
        if sku_value == "no":
            continue
        cat_map = category_maps.get(norm_col, {})
        if sku_value == "yes":
            no_code = cat_map.get("no", None)
            not_code = cat_map.get("not", None)
            if no_code is not None:
                final_mask &= norm_arrays[norm_col] != no_code
            if not_code is not None:
                final_mask &= norm_arrays[norm_col] != not_code
            if not final_mask.any():
                return final_mask
            continue
        if sku_value not in cat_map:
            final_mask[:] = False
            return final_mask
        target_code = cat_map[sku_value]
        char_mask = norm_arrays[norm_col] == target_code
        final_mask &= char_mask
        if not final_mask.any():
            return final_mask
    return final_mask


def parse_month_to_sortable(month_str: str) -> tuple:
    """Converte 'Jan 2026' in (2026, 1) per ordinamento cronologico."""
    try:
        parts = month_str.strip().split()
        if len(parts) == 2:
            month_name, year = parts
            month_num = datetime.strptime(month_name, '%b').month
            return (int(year), month_num)
        else:
            return (9999, 99)
    except Exception:
        return (9999, 99)


# ============================================================================
# FUNZIONI SPECIFICHE STOCK ALIGNMENT
# ============================================================================

def match_sku_with_supermodel(
    sku_catalog: pd.DataFrame,
    simulation_output: pd.DataFrame,
    column_mapping: Dict[str, str],
    n_months: int = 3,
) -> pd.DataFrame:
    """
    Matcha SKU alle configurazioni usando chiave composta (SKU, Supermodel).

    Args:
        sku_catalog: Catalogo con colonne [Numero componenti, Supermodel, ...]
        simulation_output: Simulazione MC con colonne configurazione + runs
        column_mapping: Mapping colonne SKU -> simulazione
        n_months: Numero mesi su cui aggregare

    Returns:
        DataFrame con (SKU, Supermodel, mean_demand)
    """
    print("\n" + "="*70)
    print("MATCHING SKU-CONFIGURAZIONI — Chiave Composta (SKU, Supermodel)")
    print("="*70)

    # Identifica colonne caratteristiche (non run)
    config_cols = [c for c in simulation_output.columns if "_Run_" not in c]
    run_cols = [c for c in simulation_output.columns if "_Run_" in c]

    n_configs = len(simulation_output)
    n_runs = len(run_cols)
    n_months_actual = min(n_months, len([c for c in run_cols if c.split("_Run_")[0].startswith("Jan") or
                                         c.split("_Run_")[0].startswith("Feb") or
                                         c.split("_Run_")[0].startswith("Mar")]))

    print(f"  Configurazioni: {n_configs}")
    print(f"  Runs per mese: {n_runs}")
    print(f"  Mesi aggregati: {n_months}")
    print(f"  Colonne caratteristiche: {len(config_cols)}")

    # Estrai solo i run dei primi n_months
    month_prefixes = []
    seen = set()
    for col in run_cols:
        prefix = col.split("_Run_")[0]
        if prefix not in seen:
            month_prefixes.append(prefix)
            seen.add(prefix)

    month_prefixes = month_prefixes[:n_months]
    n_months_actual = len(month_prefixes)

    run_cols_subset = []
    for prefix in month_prefixes:
        for i in range(n_runs):
            col = f"{prefix}_Run_{i}"
            if col in run_cols:
                run_cols_subset.append(col)

    # Inferisce n_runs_per_month (numero reale di run per mese)
    n_runs_per_month = (
        len(run_cols_subset) // n_months_actual if n_months_actual > 0 else len(run_cols_subset)
    )

    print(f"  Run considerati: {len(run_cols_subset)}")
    print(f"  Mesi effettivi: {n_months_actual},  Run/mese: {n_runs_per_month}")

    # Matrice domanda shape: (n_configs, n_months_actual * n_runs_per_month)
    demand_matrix = simulation_output[run_cols_subset].values.astype(np.float32)

    # Reshape a (n_configs, n_months_actual, n_runs_per_month) e somma i mesi
    # Risultato: demand_by_config_run[c, r] = domanda totale configurazione c al run r (su tutti i mesi)
    demand_by_config_run = (
        demand_matrix
        .reshape(n_configs, n_months_actual, n_runs_per_month)
        .sum(axis=1)  # somma sui mesi → shape (n_configs, n_runs_per_month)
    )

    # Normalizza SKU
    sku_catalog = sku_catalog.copy()
    sku_catalog["Numero componenti"] = sku_catalog["Numero componenti"].astype(str).str.strip()
    sku_catalog["Supermodel"] = sku_catalog["Supermodel"].fillna("Unknown").astype(str).str.strip()

    # Risultati per (SKU, Supermodel)
    results = []
    n_sku = len(sku_catalog)

    print(f"\n  Matching {n_sku} SKU...")

    # Perf: to_dict('records') ~10x più veloce di iterrows() per N alti
    for idx, sku_row in enumerate(sku_catalog.to_dict('records')):
        if (idx + 1) % max(1, n_sku // 10) == 0:
            print(f"    {idx + 1}/{n_sku}...", end="\r")
            sys.stdout.flush()

        sku = sku_row["Numero componenti"]
        supermodel = sku_row["Supermodel"]

        # Estrai caratteristiche rilevanti per questo SKU
        sku_config = {}
        for sku_col, sim_col in column_mapping.items():
            if sku_col in sku_row:
                val = str(sku_row[sku_col]).strip().lower()
                if val and val != "no":
                    sku_config[sim_col] = val

        # Match: tutte le caratteristiche + supermodel (Model colonna)
        match_mask = np.ones(n_configs, dtype=bool)

        for sim_col, sku_val in sku_config.items():
            if sim_col in simulation_output.columns:
                sim_vals = simulation_output[sim_col].astype(str).str.strip().str.lower()
                match_mask &= (sim_vals == sku_val)

        # Match supermodel (se presente)
        if "Model" in simulation_output.columns:
            model_vals = simulation_output["Model"].astype(str).str.strip().str.lower()
            supermodel_lower = supermodel.lower()
            # Match esatto o contenimento
            # Bugfix: regex=False evita interpretazione di metachar regex (es. '(', '+') in supermodel
            match_mask &= (model_vals == supermodel_lower) | (model_vals.str.contains(supermodel_lower, na=False, regex=False))

        # Aggrega domanda per questo (SKU, Supermodel)
        # Logica corretta:
        #   1. Somma sulle configurazioni matching (OR-logic)  → shape (n_runs_per_month,)
        #   2. Media sui run                                   → scalare
        if match_mask.any():
            demand_per_run = demand_by_config_run[match_mask].sum(axis=0)  # (n_runs_per_month,)
            mean_demand = float(demand_per_run.mean())
            n_matching = match_mask.sum()

            results.append({
                "SKU": sku,
                "Supermodel": supermodel,
                "mean_demand": mean_demand,
                "n_matching_configs": n_matching,
            })

    print(f"\n  Matching completato: {len(results)} (SKU, Supermodel) trovate")

    df_results = pd.DataFrame(results)
    return df_results


def aggregate_demand_by_supermodel(
    matched_demand: pd.DataFrame,
    n_runs: int,
    n_months: int,
) -> pd.DataFrame:
    """
    Aggrega domanda per (SKU, Supermodel) con statistiche di base.
    """
    print(f"\n  Aggregazione per n_months={n_months}, n_runs={n_runs}...")

    # Domanda media è già calcolata nel matching
    # Qui possiamo aggiungere statistiche se necessario

    return matched_demand


def main_sku_sa(
    simulation_output: pd.DataFrame,
    sku_catalog_path: str = '.\\Input\\tutte_righe_univoche_SA.xlsx',
    n_months: int = 3,
) -> pd.DataFrame:
    """
    Main entry point per Stock Alignment matching con supermodel.

    Args:
        simulation_output: DataFrame da Monte Carlo
        sku_catalog_path: Catalogo SKU (tutte_righe_univoche_SA.xlsx)
        n_months: Mesi da aggregare

    Returns:
        DataFrame (SKU, Supermodel, mean_demand)
    """
    print(f"\n{'='*70}")
    print("STOCK ALIGNMENT — SKU MATCHING CON SUPERMODEL")
    print(f"{'='*70}")

    # Carica catalogo
    try:
        sku_catalog = pd.read_excel(sku_catalog_path)
        print(f"  Catalogo caricato: {len(sku_catalog)} righe")

        # Verifica colonne essenziali
        if "Numero componenti" not in sku_catalog.columns:
            raise ValueError("Colonna 'Numero componenti' non trovata")
        if "Supermodel" not in sku_catalog.columns:
            print("  [WARN] Colonna 'Supermodel' non trovata — uso 'Model' come fallback")
            if "Model" in sku_catalog.columns:
                sku_catalog["Supermodel"] = sku_catalog["Model"]
            else:
                raise ValueError("Nessuna colonna di supermodel disponibile")
    except Exception as e:
        print(f"  [ERR] Errore caricamento catalogo: {e}")
        raise

    # Auto-rileva column mapping
    sim_cols = [c for c in simulation_output.columns if "_Run_" not in c]
    column_mapping = {}
    for sku_col in sku_catalog.columns:
        for sim_col in sim_cols:
            if sku_col.strip().lower() == sim_col.strip().lower():
                column_mapping[sku_col] = sim_col
                break

    print(f"  Column mapping: {len(column_mapping)} colonne")

    # Matching con supermodel
    # Conta mesi presenti (una colonna `*_Run_0` per ciascun mese simulato)
    n_months_detected = len([c for c in simulation_output.columns if "_Run_0" in c])
    df_matched = match_sku_with_supermodel(
        sku_catalog, simulation_output, column_mapping, n_months=n_months
    )

    # Aggrega
    df_results = aggregate_demand_by_supermodel(df_matched, n_months_detected, n_months)

    print(f"\n  Risultati finali: {len(df_results)} (SKU, Supermodel)")

    return df_results

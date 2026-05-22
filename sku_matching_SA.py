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

    for idx, (_, sku_row) in enumerate(sku_catalog.iterrows()):
        if (idx + 1) % max(1, n_sku // 10) == 0:
            print(f"    {idx + 1}/{n_sku}...", end="\r")
            sys.stdout.flush()

        sku = sku_row["Numero componenti"]
        supermodel = sku_row["Supermodel"]

        # Estrai caratteristiche rilevanti per questo SKU
        sku_config = {}
        for sku_col, sim_col in column_mapping.items():
            if sku_col in sku_row.index:
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
            match_mask &= (model_vals == supermodel_lower) | (model_vals.str.contains(supermodel_lower, na=False))

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
    n_runs = len([c for c in simulation_output.columns if "_Run_0" in c])
    df_matched = match_sku_with_supermodel(
        sku_catalog, simulation_output, column_mapping, n_months=n_months
    )

    # Aggrega
    df_results = aggregate_demand_by_supermodel(df_matched, n_runs, n_months)

    print(f"\n  Risultati finali: {len(df_results)} (SKU, Supermodel)")

    return df_results

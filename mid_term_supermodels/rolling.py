# -*- coding: utf-8 -*-
"""Pipeline ROLLING: safety stock al variare del mese di lancio.

In parole semplici: invece di calcolare la safety stock per un solo mese di
partenza (run singolo), la ricalcola per OGNI mese di lancio successivo e
combina i risultati in un'unica tabella con la colonna `mese_lancio`. Serve a
vedere come cambia la scorta di sicurezza a seconda di quando si parte.

E' un'ALTERNATIVA al run singolo (run_pipeline.py), NON lo sostituisce. Riusa
la stessa SafetyStockPipeline fixata/ottimizzata.

Ottimizzazione: il BOM/catalogo e i dati TR (mix/caratteristiche/esclusioni) NON
dipendono dal mese di lancio -> vengono costruiti UNA volta sola; nel loop per
ogni mese si rifanno solo Monte Carlo (Step 2) + matching (Step 3) + safety
(Step 4). Risultati identici al run singolo, molto piu' veloce.

Mesi di lancio di fine orizzonte (con meno mesi forecast residui di quelli
richiesti dal lead time) vengono comunque calcolati ma marcati con la colonna
`finestra_troncata = True`: i loro numeri sono sottostimati e vanno letti con
cautela.

Output (in config.output_dir):
  sku_safety_stock_rolling.xlsx          aggregato per SKU, tutte le iterazioni
  sku_safety_stock_rolling_valore.xlsx   riepilogo per mese di lancio (n_sku, unita', euro)
  sku_safety_stock_PER_MESE_rolling.xlsx vista mensile per ogni iterazione
"""
import gc
import os
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pandas as pd

from mid_term_supermodels._logging import get_logger
from mid_term_supermodels.config import PipelineConfig, SimulationConfig
from mid_term_supermodels.results import TRData
from mid_term_supermodels.pipeline import SafetyStockPipeline
from mid_term_supermodels.tr import load_monthly_tr, parse_tr_file
from mid_term_supermodels.montecarlo import (
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
    rescale_alphas,
)
from mid_term_supermodels.matching import (
    save_results_monthly,
    load_lead_times,
    build_column_mapping_auto,
    build_solo_modello_flags,
)

log = get_logger("rolling")

ROLLING_AGG_FILENAME    = "sku_safety_stock_rolling.xlsx"
ROLLING_VALORE_FILENAME = "sku_safety_stock_rolling_valore.xlsx"
ROLLING_MESE_FILENAME   = "sku_safety_stock_PER_MESE_rolling.xlsx"
ROLLING_DEF_FILENAME    = "output_rolling_def.xlsx"


def run_rolling(config: PipelineConfig,
                sim_config: SimulationConfig,
                max_iterations: int = None) -> "tuple[pd.DataFrame, pd.DataFrame]":
    """Esegue la pipeline safety stock in modalita' rolling sui mesi di lancio.

    Args:
        config:        percorsi I/O (PipelineConfig).
        sim_config:    parametri simulazione (n_runs, percentile, affidabilita').
                       Il campo start_month_offset viene IGNORATO: il rolling
                       itera da offset 0 in avanti.
        max_iterations: se valorizzato, limita il numero di mesi di lancio.

    Returns:
        (df_combined, df_valore) e scrive i 3 file rolling in config.output_dir.
    """
    t0 = time.time()
    percentile = max(1, min(99, int(sim_config.percentile_safety_stock)))
    log.info("ROLLING START | percentile=p%d n_runs=%d aff_model=%s aff_optional=%s",
             percentile, sim_config.n_runs,
             sim_config.affidabilita_model, sim_config.affidabilita_optional)

    # Non scrivere gli output della SS classica: combiniamo noi i risultati rolling.
    config_norw = replace(config, write_final_output=False)

    # ---- (una volta) STEP 1: BOM/catalogo, indipendente dal mese di lancio ----
    pipe_base = SafetyStockPipeline(config_norw, sim_config, validate_config=False)
    log.info("ROLLING: costruzione BOM (una volta sola)...")
    bom = pipe_base.run_step_1_build_bom()

    # ---- (una volta) TR base: mix/caratteristiche/alpha/esclusioni ----
    tr_path         = str(config.resolve_input(config.tr_file))
    forecast_path   = str(config.resolve_input(config.forecast_file))
    exclusions_path = str(config.resolve_input(config.exclusions_file))

    log.info("ROLLING: caricamento TR base (una volta sola)...")
    monthly_tr = load_monthly_tr(tr_path)
    model_mix, characteristics = parse_tr_file(tr_path)
    model_mix, characteristics, monthly_tr = rescale_alphas(
        model_mix, characteristics, monthly_tr,
        aff_model=sim_config.affidabilita_model,
        aff_optional=sim_config.affidabilita_optional,
    )
    exclusions = parse_exclusions(exclusions_path)
    n_months_needed = calculate_n_months_needed()

    # Lead times: indipendenti dal mese di lancio -> costruiti UNA volta sola e
    # riusati in ogni iterazione (evita 15 riletture dei file logistica/dashboard).
    # Sono costruiti in modo lazy alla prima iterazione, quando e' disponibile il
    # primo output di simulazione (serve per il column_mapping -> SOLO_MODELLO ->
    # gate parametrico). lead_times=None finche' non costruiti.
    lead_times = None

    # ---- Mesi di lancio disponibili dal forecast completo ----
    all_values, all_months = load_monthly_forecast(forecast_path, start_offset=0)
    total_months = len(all_values)
    if total_months == 0:
        raise ValueError("Forecast vuoto: impossibile eseguire il rolling.")
    n_iter = total_months
    if max_iterations is not None and max_iterations > 0:
        n_iter = min(n_iter, max_iterations)
    n_full = max(0, total_months - n_months_needed + 1)  # mesi con finestra piena
    log.info("ROLLING: %d mesi forecast, n_months_needed=%d -> %d iterazioni "
             "(%d a finestra piena, le altre troncate)",
             total_months, n_months_needed, n_iter, min(n_full, n_iter))

    tmp_monthly = str(config.resolve_output("_rolling_tmp_monthly.xlsx"))
    all_sku = []
    all_mese = []

    # --- output_rolling_def: formato lungo (Run x Item x Plan_Month) ---
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")  # identificativo univoco esecuzione
    run_month = all_months[0]                           # "oggi" = primo mese dell'esecuzione (fisso)
    all_def = []                                        # righe del nuovo output
    freq_map = {}      # SKU -> Freq_mesi (intervallo riordino, mesi)
    supply_map = {}    # SKU -> Supply_cycle (= FROZEN + Freq, Total Lead Time, mesi)
    ss_total_col = f"safety_stock_p{percentile}_total"

    for offset in range(n_iter):
        remaining = total_months - offset
        if remaining <= 0:
            break
        mese_lancio = all_months[offset]
        n_sim = min(n_months_needed, remaining)
        troncata = remaining < n_months_needed
        log.info("ROLLING [%d/%d] mese_lancio=%s n_sim=%d%s",
                 offset + 1, n_iter, mese_lancio, n_sim,
                 " (finestra troncata)" if troncata else "")

        fc_values, fc_months = load_monthly_forecast(forecast_path, start_offset=offset)
        tr_off = TRData(
            model_mix=model_mix,
            characteristics=characteristics,
            monthly_tr=monthly_tr,
            forecast_values=list(fc_values),
            forecast_months=list(fc_months),
            exclusions=exclusions,
            n_months_needed=n_sim,
        )
        sim_cfg_off = replace(sim_config, start_month_offset=offset)
        pipe = SafetyStockPipeline(config_norw, sim_cfg_off, validate_config=False)

        sim = pipe.run_step_2_montecarlo(tr_off)

        # Costruzione lazy dei lead_times (una sola volta): col primo sim calcolo
        # column_mapping -> SOLO_MODELLO -> gate parametrico, poi riuso in tutte
        # le iterazioni successive.
        if lead_times is None:
            sku_catalog_raw = bom.catalog.copy()
            if "ID" in sku_catalog_raw.columns:
                sku_catalog_raw = sku_catalog_raw.drop(columns=["ID"])
            _feat = [c for c in sku_catalog_raw.columns
                     if c.lower().strip() not in
                     {"id", "numero componenti",
                      "numero componenti.1", "versione"}]
            _simcols = [c for c in sim.df_wide.columns if "_Run_" not in c]
            _cmap = build_column_mapping_auto(_feat, _simcols)
            _flags = build_solo_modello_flags(sku_catalog_raw, _cmap)
            _solo = (set(_flags.loc[_flags["SOLO_MODELLO"], "SKU"].astype(str))
                     if len(_flags) else set())
            lead_times = load_lead_times(
                solo_modello_skus=_solo,
                gate_max_mesi=sim_config.gate_max_mesi,
            )
            log.info("ROLLING: lead_times costruiti una volta (%d SKU, %d SOLO_MODELLO, gate_max=%s mesi)",
                     len(lead_times), len(_solo), sim_config.gate_max_mesi)
            # Mappe per output_rolling_def (Freq_mesi e Supply_cycle per SKU)
            if "Freq_mesi" in lead_times.columns:
                freq_map = dict(zip(lead_times["SKU"].astype(str), lead_times["Freq_mesi"]))
            if "Supply_cycle" in lead_times.columns:
                supply_map = dict(zip(lead_times["SKU"].astype(str), lead_times["Supply_cycle"]))

        mat = pipe.run_step_3_matching(sim, bom, lead_times=lead_times)
        ss  = pipe.run_step_4_safety_stock(mat, bom)

        per_sku = ss.per_sku
        if per_sku is None or len(per_sku) == 0:
            log.warning("ROLLING [%s]: nessun risultato, salto", mese_lancio)
            continue

        # --- Raccolta righe output_rolling_def per questa iterazione ---
        # Plan_Month = mese_lancio; Safety_Stock_Qty = SS windowed su residual LT
        # calcolata DA QUESTA iterazione (dinamica per Plan_Month); Demand_Qty =
        # domanda del SINGOLO mese di lancio (monthly_stats) x qty; Cycle = Demand x Freq.
        for _rec in per_sku.to_dict("records"):
            _sku = str(_rec.get("SKU"))
            _qty = float(_rec.get("Quantity_Per_Bike", 1.0) or 0.0)
            _ms = _rec.get("monthly_stats") or {}
            _sm = _ms.get(mese_lancio) or {}
            _demand = round(float(_sm.get("mean_demand", 0.0)) * _qty, 2)
            _freq = freq_map.get(_sku)
            _cycle = (round(_demand * float(_freq), 2)
                      if _freq is not None and pd.notna(_freq) else 0.0)
            _ss = round(float(_rec.get(ss_total_col, 0.0) or 0.0), 2)
            _supply = supply_map.get(_sku)
            _flag = (1 if _supply is not None and pd.notna(_supply)
                     and int(round(float(_supply))) == offset else 0)
            all_def.append({
                "RunID": run_id,
                "Run_Month": run_month,
                "Item": _sku,
                "Plan_Month": mese_lancio,
                "Demand_Qty": _demand,
                "Safety_Stock_Qty": _ss,
                "Cycle_Stock_Qty": _cycle,
                "Flag_Static_SS": _flag,
            })

        per_sku = per_sku.copy()
        per_sku.insert(0, "mese_lancio", mese_lancio)
        per_sku.insert(1, "offset", offset)
        per_sku.insert(2, "finestra_troncata", troncata)
        per_sku.insert(3, "mesi_simulati", n_sim)
        all_sku.append(per_sku)

        # Vista mensile: la produce save_results_monthly (espansione per mese).
        try:
            save_results_monthly(ss.per_mese, tmp_monthly, percentile=percentile)
            df_pm = pd.read_excel(tmp_monthly)
            df_pm.insert(0, "mese_lancio", mese_lancio)
            df_pm.insert(1, "offset", offset)
            df_pm.insert(2, "finestra_troncata", troncata)
            all_mese.append(df_pm)
        except Exception as e:
            log.warning("ROLLING [%s]: vista mensile non disponibile: %s", mese_lancio, e)

        gc.collect()

    if os.path.exists(tmp_monthly):
        try:
            os.remove(tmp_monthly)
        except Exception:
            pass

    if not all_sku:
        raise RuntimeError("ROLLING: nessuna iterazione ha prodotto risultati.")

    # ---- Combina + riepilogo valore per mese di lancio ----
    df_combined = pd.concat(all_sku, ignore_index=True)

    ss_total_col = f"safety_stock_p{percentile}_total"
    agg = {"SKU": "count"}
    if ss_total_col in df_combined.columns:
        agg[ss_total_col] = "sum"
    if "prezzo_safety" in df_combined.columns:
        agg["prezzo_safety"] = "sum"
    df_valore = (
        df_combined.groupby("mese_lancio", sort=False)
        .agg(agg)
        .rename(columns={"SKU": "n_sku",
                         ss_total_col: "ss_unita_totale",
                         "prezzo_safety": "ss_valore_eur"})
        .reset_index()
    )
    # ordine cronologico + flag finestra_troncata per mese
    meta = (df_combined[["mese_lancio", "offset", "finestra_troncata"]]
            .drop_duplicates("mese_lancio")
            .sort_values("offset"))
    df_valore = meta.merge(df_valore, on="mese_lancio", how="left")

    # ---- Salvataggio ----
    out_agg    = config.resolve_output(ROLLING_AGG_FILENAME)
    out_valore = config.resolve_output(ROLLING_VALORE_FILENAME)
    out_mese   = config.resolve_output(ROLLING_MESE_FILENAME)
    df_combined.to_excel(str(out_agg), index=False)
    df_valore.to_excel(str(out_valore), index=False)
    if all_mese:
        pd.concat(all_mese, ignore_index=True).to_excel(str(out_mese), index=False)

    # ---- output_rolling_def: formato lungo richiesto (Run x Item x Plan_Month) ----
    out_def = config.resolve_output(ROLLING_DEF_FILENAME)
    _def_cols = ["RunID", "Run_Month", "Item", "Plan_Month",
                 "Demand_Qty", "Safety_Stock_Qty", "Cycle_Stock_Qty", "Flag_Static_SS"]
    df_def = (pd.DataFrame(all_def, columns=_def_cols) if all_def
              else pd.DataFrame(columns=_def_cols))
    df_def.to_excel(str(out_def), index=False)
    log.info("ROLLING: output_rolling_def scritto: %d righe, %d Flag_Static_SS=1 -> %s",
             len(df_def), int(df_def["Flag_Static_SS"].sum()) if len(df_def) else 0, out_def)

    elapsed = time.time() - t0
    log.info("ROLLING COMPLETATO in %.1fs | %d mesi di lancio, %d righe totali",
             elapsed, df_combined["mese_lancio"].nunique(), len(df_combined))
    log.info("ROLLING output: %s | %s | %s | %s", out_agg, out_valore, out_mese, out_def)
    return df_combined, df_valore

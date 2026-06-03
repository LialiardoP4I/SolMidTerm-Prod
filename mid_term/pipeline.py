# -*- coding: utf-8 -*-
"""Orchestratore end-to-end pipeline SS Classica (step 0..4)."""
import time
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from mid_term._logging import get_logger
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.results import (TRData, BOMData, SimulationResult,
                               MatchResult, SafetyStockResult)
from mid_term.tr import load_monthly_tr, parse_tr_file
from mid_term.bom import (
    process_combination,
    consolidate_results,
    run_quality_checks,
    build_unified_mapping,
)
from mid_term.matching import (
    match_skus_ultra_optimized,
    build_column_mapping_auto,
    load_lead_times,
    _reinit_norm_cache,
    _load_cambio_eur,
    canonical_model_name,
    save_results,
    save_results_monthly,
    CambioValutaError,
)
from mid_term.montecarlo import (
    SimulationConfig as MCSimulationConfig,
    run_monthly_simulations_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
    build_zcol_pack_color_mapping,
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
    rescale_alphas,
)

log = get_logger("pipeline")


class SafetyStockPipeline:
    """Orchestratore end-to-end della pipeline Safety Stock Classica.

    Espone:
      - run() esegue tutti gli step 0..4 in sequenza
      - run_step_N() invocabile singolarmente (utile per debug, test parziali,
        integrazione con orchestratori esterni tipo Airflow)

    Args:
        config:     PipelineConfig con path file + flag I/O
        sim_config: SimulationConfig con parametri MC + percentili
        validate_config: se True (default) valida la config in __init__
    """

    def __init__(
        self,
        config: PipelineConfig,
        sim_config: Optional[SimulationConfig] = None,
        validate_config: bool = True,
    ):
        self.config = config
        self.sim_config = sim_config or SimulationConfig()
        if validate_config:
            self.config.validate(strict=True)

    def run(self) -> SafetyStockResult:
        tr  = self.run_step_0_load_tr()
        bom = self.run_step_1_build_bom()
        sim = self.run_step_2_montecarlo(tr)
        mat = self.run_step_3_matching(sim, bom)
        ss  = self.run_step_4_safety_stock(mat, bom)
        return ss

    # ---- STEP 0 ------------------------------------------------------------
    def run_step_0_load_tr(self) -> TRData:
        """Carica TR + caratteristiche + forecast + esclusioni."""
        t0 = time.time()
        tr_path         = str(self.config.resolve_input(self.config.tr_file))
        forecast_path   = str(self.config.resolve_input(self.config.forecast_file))
        exclusions_path = str(self.config.resolve_input(self.config.exclusions_file))

        log.info("Step 0: caricamento TR %s", tr_path)
        monthly_tr = load_monthly_tr(tr_path)
        model_mix, characteristics = parse_tr_file(tr_path)

        log.info("Step 0: caricamento forecast %s", forecast_path)
        forecast_values, month_names = load_monthly_forecast(
            forecast_path, start_offset=self.sim_config.start_month_offset
        )

        # NOTA: calculate_n_months_needed() legge hardcoded Input/residual_lead_time.xlsx.
        # Limitazione nota: refactoring previsto in Fase 8.
        n_months_needed = calculate_n_months_needed()

        log.info("Step 0: caricamento esclusioni %s", exclusions_path)
        exclusions = parse_exclusions(exclusions_path)

        elapsed = time.time() - t0
        log.info("Step 0 completato in %.1fs: %d caratteristiche, %d mesi",
                 elapsed, len(characteristics), len(month_names))

        return TRData(
            model_mix=model_mix,
            characteristics=characteristics,
            monthly_tr=monthly_tr,
            forecast_values=list(forecast_values),
            forecast_months=list(month_names),
            exclusions=exclusions,
            n_months_needed=n_months_needed,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 1 ------------------------------------------------------------
    def run_step_1_build_bom(self) -> BOMData:
        """Costruisce matrice BOM + catalogo configurazioni multi-supermodel.

        Replica la logica di process_all_bom_combinations.main():
          1. build_unified_mapping (genera Mappatura_Unificata.xlsx)
          2. run_quality_checks (TR char options, prezzi, states)
          3. process_combination per ogni supermodel -> matrici per BOM
          4. consolidate_results -> catalogo finale deduplicato

        NOTE LIMITAZIONI:
          - process_combination scrive sempre matrix_output_{nome}.xlsx
            e tutte_righe_univoche_{nome}.xlsx su OUTPUT_DIR hardcoded
            (mid_term.bom.OUTPUT_DIR = Input/). write_intermediate qui
            controlla SOLO il salvataggio del catalogo finale.
          - run_quality_checks legge da path hardcoded INPUT_DIR/.
          - build_unified_mapping legge da MAPPATURE_ORIGINALI hardcoded.
        """
        t0 = time.time()
        supermodels = list(self.config.bom_files.keys())
        log.info("Step 1: build BOM per supermodel %s", supermodels)

        # 1. Mappatura unificata (path hardcoded in mid_term.bom)
        build_unified_mapping()

        # 2. Quality checks (legge file hardcoded in mid_term.bom)
        residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set = \
            run_quality_checks()

        # 3. Per ogni supermodel costruisci la matrice via process_combination.
        # process_combination si aspetta un dict con chiavi:
        #   nome, in_dir, mappatura, tr
        # Usiamo i path della config quando possibile.
        matrices = {}
        tr_path = self.config.resolve_input(self.config.tr_file)

        for supermodel in supermodels:
            # in_dir e mappatura non sono parametrizzabili direttamente nella
            # PipelineConfig (vivono sotto Input/MODEL/<supermodel>/).
            # Costruiamo i path coerentemente con la struttura legacy.
            in_dir = self.config.input_dir / "MODEL" / supermodel
            mappatura_path = in_dir / f"Mappatura_{supermodel}.xlsx"

            log.info("Step 1: processing %s (in_dir=%s)", supermodel, in_dir)
            combo = {
                "nome": supermodel,
                "in_dir": in_dir,
                "mappatura": mappatura_path,
                "tr": tr_path,
            }
            matrix_df = process_combination(
                combo,
                tr_char_options=tr_char_options,
                filter_untraced=self.config.filter_untraced,
            )
            matrices[supermodel] = matrix_df

            if self.config.write_intermediate:
                # process_combination ha gia' scritto su OUTPUT_DIR hardcoded;
                # copiamo/risalviamo nella output_dir richiesta dalla pipeline
                # solo se diverso da quel default.
                out_path = self.config.resolve_output(
                    self.config.matrix_filename_tpl.format(supermodel=supermodel)
                )
                try:
                    if out_path.parent.resolve() != \
                            (self.config.input_dir).resolve():
                        matrix_df.to_excel(out_path, index=False)
                        log.info("Step 1: matrice copiata in %s", out_path)
                except Exception as e:
                    log.warning("Step 1: impossibile salvare matrice in output_dir: %s", e)

        # 4. Consolida tutte le matrici in catalogo unico
        df_consolidated = consolidate_results(matrices)

        # 5. Quality report come DataFrame (riepilogo issues)
        quality_counts = len(residual_issues) + len(price_issues) + len(ct_cv_issues)
        quality_report = pd.DataFrame({
            "Categoria": ["Residual Issues", "Price Issues", "CT/CV Issues", "TOTALE"],
            "Conteggio": [len(residual_issues), len(price_issues),
                          len(ct_cv_issues), quality_counts],
        })

        # 6. Salva catalogo consolidato se richiesto
        if self.config.write_intermediate:
            catalog_path = self.config.resolve_output(self.config.catalog_filename)
            try:
                with pd.ExcelWriter(catalog_path, engine="openpyxl") as writer:
                    df_consolidated.to_excel(
                        writer, sheet_name="Tutte_Righe_Univoche", index=False
                    )
                    for nome, df in matrices.items():
                        df.to_excel(writer, sheet_name=f"Righe_{nome}", index=False)
                    quality_report.to_excel(
                        writer, sheet_name="Quality_Summary", index=False
                    )
                log.info("Step 1: catalogo consolidato salvato %s", catalog_path)
            except Exception as e:
                log.warning("Step 1: impossibile salvare catalogo: %s", e)

        elapsed = time.time() - t0
        log.info("Step 1 completato in %.1fs: %d supermodel, %d righe catalogo",
                 elapsed, len(matrices), len(df_consolidated))

        return BOMData(
            catalog=df_consolidated,
            matrices=matrices,
            quality_report=quality_report,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 2 ------------------------------------------------------------
    def run_step_2_montecarlo(self, tr: TRData) -> SimulationResult:
        """Esegue Monte Carlo + Dirichlet usando TR + esclusioni dello step 0."""
        t0 = time.time()
        # Mapping campo: max_conflict_iterations (config pubblico)
        #             -> max_conflict_resolution_iterations (interno)
        mc_cfg = MCSimulationConfig(
            n_runs=self.sim_config.n_runs,
            random_seed=self.sim_config.random_seed,
            start_month_offset=self.sim_config.start_month_offset,
            use_int16=self.sim_config.use_int16,
            max_conflict_resolution_iterations=self.sim_config.max_conflict_iterations,
        )
        log.info("Step 2: MC n_runs=%d seed=%d", mc_cfg.n_runs, mc_cfg.random_seed)

        results_by_month = run_monthly_simulations_optimized(
            forecast_values=tr.forecast_values,
            month_names=tr.forecast_months,
            n_months=tr.n_months_needed,
            model_mix=tr.model_mix,
            characteristics=tr.characteristics,
            exclusions=tr.exclusions,
            config=mc_cfg,
            monthly_tr=tr.monthly_tr,
        )

        # build_wide_dataframe firma reale: (results_by_month, month_names, n_runs, ...)
        df_wide = build_wide_dataframe(
            results_by_month,
            tr.forecast_months,
            mc_cfg.n_runs,
        )

        # Pack colors mapping (post-processing).
        # build_zcol_pack_color_mapping firma reale:
        #   (mappatura_path, rally_mappatura_path=None) -- NON accetta `characteristics`.
        mapping_path = str(self.config.resolve_input(self.config.mapping_file))
        try:
            pack_color_mapping = build_zcol_pack_color_mapping(mapping_path)
            df_wide = apply_pack_colors_to_wide_df(df_wide, pack_color_mapping)
        except Exception as e:
            log.warning("Pack colors mapping non applicabile: %s", e)

        elapsed = time.time() - t0
        log.info("Step 2 completato in %.1fs: %d configurazioni",
                 elapsed, len(df_wide))

        return SimulationResult(
            df_wide=df_wide,
            n_runs=mc_cfg.n_runs,
            n_months=tr.n_months_needed,
            month_names=tr.forecast_months,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 3 ------------------------------------------------------------
    def run_step_3_matching(self, sim: SimulationResult, bom: BOMData) -> MatchResult:
        """SKU matching: confronta catalogo BOM con output simulazione MC.

        Replica la prima meta' di sku_matching_SS.main_sku_v2():
          1. Prepara catalogo (rimuove ID, dedup)
          2. Carica lead times (path hardcoded in mid_term.matching)
          3. Costruisce column_mapping auto-detect
          4. Calcola flag SOLO_MODELLO
          5. Re-inizializza NormalizationCache con prefissi dinamici
          6. Esegue match_skus_ultra_optimized -> DataFrame matched

        NOTE LIMITAZIONI:
          - load_lead_times() legge sempre Input/residual_lead_time.xlsx (hardcoded).
        """
        t0 = time.time()
        percentile = max(1, min(99, int(self.sim_config.percentile_safety_stock)))
        log.info("Step 3: SKU matching (percentile=%d)", percentile)

        simulation_output = sim.df_wide

        # 1. Prepara catalogo SKU dal BOM step
        sku_catalog_raw = bom.catalog.copy()
        if "ID" in sku_catalog_raw.columns:
            sku_catalog_raw = sku_catalog_raw.drop(columns=["ID"])
        sku_catalog = sku_catalog_raw.drop_duplicates(keep="first")
        log.info("Step 3: catalogo %d righe (raw=%d)",
                 len(sku_catalog), len(sku_catalog_raw))

        # 2. Carica lead times
        lead_times = load_lead_times()
        log.info("Step 3: lead_times %d SKU", len(lead_times))

        # 3. Identifica n_runs dalle colonne _Run_
        month_run_cols = [c for c in simulation_output.columns if "_Run_" in c]
        if month_run_cols:
            first_prefix = month_run_cols[0].split("_Run_")[0]
            first_month_cols = [c for c in month_run_cols
                                if c.startswith(first_prefix)]
            n_runs = len(first_month_cols)
        else:
            n_runs = sim.n_runs
        log.info("Step 3: n_runs rilevati=%d", n_runs)

        # 4. Costruisci column_mapping auto-detect
        # 'versione' aggiunta da process_all_bom_combinations (catalogo multi-BOM)
        # e non e' una caratteristica di configurazione: va esclusa.
        sku_feature_cols = [c for c in sku_catalog.columns
                            if c.lower().strip() not in
                            {"id", "numero componenti",
                             "numero componenti.1", "versione"}]
        sim_config_cols = [c for c in simulation_output.columns
                           if "_Run_" not in c]
        column_mapping = build_column_mapping_auto(sku_feature_cols, sim_config_cols)
        log.info("Step 3: column_mapping (%d colonne)", len(column_mapping))
        for sku_col, sim_col in sorted(column_mapping.items()):
            log.debug("  '%s' -> '%s'", sku_col, sim_col)

        # 5. Calcola flag SOLO_MODELLO (identico a main_sku_v2)
        _model_col = next(
            (c for c in sku_catalog_raw.columns if c.lower().strip() == "model"),
            None,
        )
        _opt_flag_cols = [k for k in column_mapping.keys()
                          if _model_col is None or k.lower().strip() != "model"]

        def _is_no_val(v):
            return str(v).strip().lower() in ("no", "nan", "")

        _all_models_catalog = set()
        if _model_col:
            for _mv in sku_catalog_raw[_model_col].dropna().unique():
                for _m in str(_mv).split("+"):
                    _canon = canonical_model_name(_m)
                    if _canon:
                        _all_models_catalog.add(_canon)

        _flag_rows = []
        for _sku, _grp in sku_catalog_raw.groupby("Numero componenti"):
            _all_no = all(
                _is_no_val(v) for _c in _opt_flag_cols
                if _c in _grp.columns
                for v in _grp[_c]
            )
            _sku_models = set()
            if _model_col:
                for _mv in _grp[_model_col].dropna():
                    for _m in str(_mv).split("+"):
                        _canon = canonical_model_name(_m)
                        if _canon:
                            _sku_models.add(_canon)
            _not_all = bool(_all_models_catalog) and (
                _sku_models != _all_models_catalog
            )
            _flag_rows.append({"SKU": _sku, "SOLO_MODELLO": _all_no and _not_all})

        df_flag_solo_modello = pd.DataFrame(_flag_rows)
        _n_solo = int(df_flag_solo_modello["SOLO_MODELLO"].sum()) \
            if len(df_flag_solo_modello) > 0 else 0
        log.info("Step 3: SOLO_MODELLO %d/%d SKU",
                 _n_solo, len(df_flag_solo_modello))

        # 6. Re-inizializza NormalizationCache con prefissi dinamici
        _reinit_norm_cache(column_mapping)

        # 7. Matching (include espansione multi-modello)
        df_matched = match_skus_ultra_optimized(
            sku_catalog=sku_catalog,
            lead_times=lead_times,
            simulation_output=simulation_output,
            column_mapping=column_mapping,
            n_runs=n_runs,
            percentile=percentile,
        )

        elapsed = time.time() - t0
        log.info("Step 3 completato in %.1fs: %d righe matched",
                 elapsed, len(df_matched) if df_matched is not None else 0)

        return MatchResult(
            df_matched=df_matched,
            column_mapping=column_mapping,
            solo_modello=df_flag_solo_modello,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 4 ------------------------------------------------------------
    def run_step_4_safety_stock(self, mat: MatchResult, bom: BOMData) -> SafetyStockResult:
        """Applica prezzi + merge SOLO_MODELLO + scrive output Excel.

        Replica la seconda meta' di sku_matching_SS.main_sku_v2():
          1. Merge prezzi EUR (se config.prices_file disponibile)
          2. Merge flag SOLO_MODELLO
          3. Scrive output aggregato per SKU
          4. Scrive output mensile per SKU
        """
        t0 = time.time()
        percentile = max(1, min(99, int(self.sim_config.percentile_safety_stock)))
        results = mat.df_matched
        df_flag_solo_modello = mat.solo_modello

        # 1. Merge prezzi con conversione valuta -> EUR
        if self.config.prices_file:
            try:
                prezzi_path = self.config.resolve_input(self.config.prices_file)
                all_data = pd.read_excel(prezzi_path, sheet_name="Sheet1")

                material_col = "Materiale"
                prezzo_netto_col = "Prezzo netto"
                data_doc_col = "Data documento"
                unit_prezzo_col = None
                divisa_col = None
                for col in all_data.columns:
                    col_lower = str(col).lower().replace("\xa0", " ").replace("\x92", "'")
                    if "unit" in col_lower and "prezzo" in col_lower:
                        unit_prezzo_col = col
                        break
                for col in all_data.columns:
                    col_lower = str(col).lower().replace("\xa0", " ").replace("\x92", "'")
                    if "divisa" in col_lower or "currency" in col_lower:
                        divisa_col = col
                        break

                if unit_prezzo_col is None or divisa_col is None:
                    raise ValueError(
                        f"Required columns not found. unit_prezzo_col={unit_prezzo_col}, "
                        f"divisa_col={divisa_col}"
                    )

                prezzi = all_data[[material_col, prezzo_netto_col, unit_prezzo_col,
                                   divisa_col, data_doc_col]].drop_duplicates()
                _prezzo_num = pd.to_numeric(prezzi[prezzo_netto_col], errors="coerce")
                _unit_num = pd.to_numeric(prezzi[unit_prezzo_col],
                                          errors="coerce").replace(0, np.nan)
                prezzi = prezzi.copy()
                prezzi["Prezzo netto"] = _prezzo_num / _unit_num

                # Tassi di cambio EUR (path hardcoded in matching._load_cambio_eur)
                cambio = _load_cambio_eur()
                log.info("Step 4: tassi di cambio caricati: %d valute", len(cambio))

                prezzi[divisa_col] = (prezzi[divisa_col].astype(str)
                                      .str.strip().str.upper())
                prezzi = prezzi.merge(cambio, left_on=divisa_col,
                                      right_on="Codice", how="left")

                _miss_mask = prezzi["1 unità in EUR ≈"].isna()
                if _miss_mask.any():
                    _missing_curr = sorted(prezzi.loc[_miss_mask, divisa_col]
                                           .dropna().astype(str).unique())
                    raise CambioValutaError(
                        f"Valute presenti in Prezzi.xlsx ma non mappate in "
                        f"Cambio Valuta.xlsx: {_missing_curr}\n"
                        f"Aggiungere queste valute al file Cambio Valuta prima "
                        f"di rilanciare."
                    )

                prezzi["Prezzo netto EUR"] = (prezzi["Prezzo netto"]
                                              * prezzi["1 unità in EUR ≈"])

                prezzi = (prezzi[[material_col, data_doc_col, "Prezzo netto EUR"]]
                          .sort_values(data_doc_col)
                          .groupby(material_col)[["Prezzo netto EUR"]]
                          .last()
                          .rename(columns={"Prezzo netto EUR": "Prezzo netto"}))

                ss_key = f"safety_stock_p{percentile}_total"
                results = results.merge(prezzi, left_on="SKU",
                                        right_index=True, how="left")
                results["prezzo_safety"] = results["Prezzo netto"] * results[ss_key]
            except CambioValutaError:
                raise
            except Exception as e:
                log.warning("Step 4: prezzi non disponibili: %s", e)

        # 2. Merge SOLO_MODELLO
        if results is None or len(results) == 0:
            log.warning("Step 4: nessun risultato dal matching - risultati vuoti")
            results = pd.DataFrame(columns=["SKU", "SOLO_MODELLO"])
        elif "SKU" not in results.columns:
            log.warning("Step 4: colonna 'SKU' assente in results")
            results["SOLO_MODELLO"] = False
        else:
            results = results.merge(df_flag_solo_modello, on="SKU", how="left")
            results["SOLO_MODELLO"] = results["SOLO_MODELLO"].fillna(False)

        # 3. Scrittura output (se richiesto)
        output_paths = {}
        if self.config.write_final_output:
            per_sku_path = self.config.resolve_output(
                self.config.safety_stock_per_sku_filename
            )
            per_mese_path = self.config.resolve_output(
                self.config.safety_stock_per_mese_filename
            )
            log.info("Step 4: salvataggio %s e %s", per_sku_path, per_mese_path)
            save_results(results, str(per_sku_path), percentile=percentile)
            save_results_monthly(results, str(per_mese_path), percentile=percentile)
            output_paths["per_sku"] = str(per_sku_path)
            output_paths["per_mese"] = str(per_mese_path)

        # 4. Costruisci DataFrame finale per SafetyStockResult
        # save_results scrive il file aggregato; in memoria, per_sku = results,
        # per_mese = vista per-mese derivata da results (stessa logica di save_results_monthly).
        per_sku = results
        per_mese = results  # save_results_monthly espande lo stesso df

        elapsed = time.time() - t0
        log.info("Step 4 completato in %.1fs: %d righe finali",
                 elapsed, len(per_sku))

        return SafetyStockResult(
            per_sku=per_sku,
            per_mese=per_mese,
            percentile=percentile,
            output_paths=output_paths,
            elapsed_seconds=elapsed,
        )

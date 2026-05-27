# -*- coding: utf-8 -*-
"""Orchestratore end-to-end pipeline SS Classica (step 0..4)."""
import time
from typing import Optional

import pandas as pd

from mid_term._logging import get_logger
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.results import (TRData, BOMData, SimulationResult,
                               MatchResult, SafetyStockResult)
from mid_term.tr import load_monthly_tr, parse_tr_file
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
        """Costruisce matrice BOM + catalogo configurazioni multi-supermodel."""
        raise NotImplementedError("TODO Fase 6 parte 2")

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
        """TODO Fase 6 parte 2: SKU matching."""
        raise NotImplementedError("TODO Fase 6 parte 2")

    # ---- STEP 4 ------------------------------------------------------------
    def run_step_4_safety_stock(self, mat: MatchResult, bom: BOMData) -> SafetyStockResult:
        """TODO Fase 6 parte 2: calcolo percentili + scrittura output."""
        raise NotImplementedError("TODO Fase 6 parte 2")

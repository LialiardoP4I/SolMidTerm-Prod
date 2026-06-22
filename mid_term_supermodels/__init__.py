# -*- coding: utf-8 -*-
"""Mid-Term Safety Stock Pipeline.

In parole semplici: pacchetto che calcola la scorta di sicurezza (safety stock)
dei componenti, simulando la domanda con il metodo Monte Carlo e abbinando i
componenti agli scenari sulle distinte base dei vari modelli.

API pubblica:
    PipelineConfig, SimulationConfig — configurazione (percorsi file e parametri)
    Eccezioni: MidTermError + sotto-classi specifiche per ogni step
"""
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.exceptions import (
    MidTermError,
    ConfigurationError,
    InputFileError,
    TRParseError,
    BOMParseError,
    SimulationError,
    MatchingError,
)
from mid_term.tr import (
    canonical_model_name,
    ModelMix,
    CharacteristicGroup,
    parse_tr_file,
    parse_model_mix,
    load_monthly_tr,
    get_tr_for_month,
)
from mid_term.bom import (
    process_excel_to_matrix_v2,
    crea_tutte_righe_univoche_v2,
    process_combination,
    run_quality_checks,
    build_unified_mapping,
    consolidate_results,
)
from mid_term.montecarlo import (
    run_monthly_simulations_optimized,
    run_single_month_simulation_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
    build_zcol_pack_color_mapping,
    rescale_alphas,
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
)
from mid_term.matching import (
    match_skus_ultra_optimized,
    NormalizationCache,
    normalize_value_cached,
    build_column_mapping_auto,
    load_lead_times,
)
from mid_term.pipeline import SafetyStockPipeline
from mid_term.results import (
    TRData,
    BOMData,
    SimulationResult,
    MatchResult,
    SafetyStockResult,
)

__all__ = [
    "PipelineConfig",
    "SimulationConfig",
    "MidTermError",
    "ConfigurationError",
    "InputFileError",
    "TRParseError",
    "BOMParseError",
    "SimulationError",
    "MatchingError",
    "canonical_model_name",
    "ModelMix",
    "CharacteristicGroup",
    "parse_tr_file",
    "parse_model_mix",
    "load_monthly_tr",
    "get_tr_for_month",
    "process_excel_to_matrix_v2",
    "crea_tutte_righe_univoche_v2",
    "process_combination",
    "run_quality_checks",
    "build_unified_mapping",
    "consolidate_results",
    "run_monthly_simulations_optimized",
    "run_single_month_simulation_optimized",
    "build_wide_dataframe",
    "apply_pack_colors_to_wide_df",
    "build_zcol_pack_color_mapping",
    "rescale_alphas",
    "load_monthly_forecast",
    "calculate_n_months_needed",
    "parse_exclusions",
    "match_skus_ultra_optimized",
    "NormalizationCache",
    "normalize_value_cached",
    "build_column_mapping_auto",
    "load_lead_times",
    "SafetyStockPipeline",
    "TRData",
    "BOMData",
    "SimulationResult",
    "MatchResult",
    "SafetyStockResult",
]

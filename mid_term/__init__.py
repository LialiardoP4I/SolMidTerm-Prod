# -*- coding: utf-8 -*-
"""Mid-Term Safety Stock Pipeline.

Package per il calcolo del Safety Stock di componenti tramite simulazione
Monte Carlo + matching SKU su distinte base multi-supermodel.

API pubblica (in espansione mano a mano che i moduli vengono migrati):
    PipelineConfig, SimulationConfig — configurazione dichiarativa
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
]

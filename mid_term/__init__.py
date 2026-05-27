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
]

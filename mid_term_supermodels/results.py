# -*- coding: utf-8 -*-
"""Strutture dati che raccolgono i risultati di ogni passo della pipeline.

In parole semplici: ogni passo (step 0..4) restituisce un "contenitore" con i
suoi risultati: TRData (dati TR), BOMData (distinta base), SimulationResult
(scenari simulati), MatchResult (abbinamento componenti) e SafetyStockResult
(safety stock finale). Servono a passare i dati in modo ordinato da un passo al
successivo.

ModelMix e CharacteristicGroup vivono in mid_term.tr; qui usati come
forward reference per evitare import circolari.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Tuple, TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from mid_term.tr import ModelMix, CharacteristicGroup


@dataclass
class TRData:
    """Output step 0: TR + caratteristiche + forecast + esclusioni."""
    model_mix:        "ModelMix"
    characteristics:  Dict[str, "CharacteristicGroup"]
    monthly_tr:       Dict[str, Tuple["ModelMix", Dict[str, "CharacteristicGroup"]]]
    forecast_values:  List[float]
    forecast_months:  List[str]
    exclusions:       Dict[str, List[str]]
    n_months_needed:  int
    elapsed_seconds:  float = 0.0


@dataclass
class BOMData:
    """Output step 1: catalogo configurazioni + matrici per supermodel."""
    catalog:         pd.DataFrame
    matrices:        Dict[str, pd.DataFrame]
    quality_report:  pd.DataFrame
    elapsed_seconds: float = 0.0


@dataclass
class SimulationResult:
    """Output step 2: DataFrame wide delle configurazioni simulate."""
    df_wide:         pd.DataFrame
    n_runs:          int
    n_months:        int
    month_names:     List[str]
    elapsed_seconds: float = 0.0


@dataclass
class MatchResult:
    """Output step 3: SKU matched + column mapping + flag SOLO_MODELLO."""
    df_matched:      pd.DataFrame
    column_mapping:  Dict[str, str]
    solo_modello:    pd.DataFrame
    no_lt:           pd.DataFrame = field(default_factory=pd.DataFrame)
    elapsed_seconds: float = 0.0


@dataclass
class SafetyStockResult:
    """Output step 4: safety stock per SKU e per mese."""
    per_sku:         pd.DataFrame
    per_mese:        pd.DataFrame
    percentile:      int
    output_paths:    Dict[str, str] = field(default_factory=dict)
    elapsed_seconds: float = 0.0

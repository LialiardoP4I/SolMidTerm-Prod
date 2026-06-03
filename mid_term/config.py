# -*- coding: utf-8 -*-
"""Configurazione dichiarativa della pipeline SS Classica."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from mid_term.exceptions import ConfigurationError


@dataclass(frozen=True)
class PipelineConfig:
    """Configurazione path file + opzioni I/O. Frozen, immutabile, hashable."""

    input_dir:  Path
    output_dir: Path

    tr_file:         str = "TR TOTALV21E - Copia.xlsx"
    mapping_file:    str = "Mappatura_Unificata.xlsx"
    forecast_file:   str = "Total_demand.xlsx"
    exclusions_file: str = "esclusioni.xlsx"
    states_file:     str = "States.xlsx"

    bom_files: Dict[str, str] = field(default_factory=lambda: {
        "V21E": "BOM150_V21E.xlsx",
        "V22E": "BOM150_V22E.xlsx",
        "V24T": "BOM150_V24T.xlsx",
    })

    catalog_filename:    str = "tutte_righe_univoche_V2.xlsx"
    matrix_filename_tpl: str = "matrix_output_{supermodel}.xlsx"

    safety_stock_per_sku_filename:  str = "sku_safety_stock_leadtime_PER_SKU_V2.xlsx"
    safety_stock_per_mese_filename: str = "sku_safety_stock_PER_MESE_V2.xlsx"

    lead_times_file: str = "residual_lead_time.xlsx"
    prices_file:     Optional[str] = "Prezzi.xlsx"

    write_intermediate: bool = True
    write_final_output: bool = True
    filter_untraced:    bool = True

    def resolve_input(self, name: str) -> Path:
        return self.input_dir / name

    def resolve_output(self, name: str) -> Path:
        return self.output_dir / name

    def validate(self, strict: bool = True) -> List[str]:
        """Verifica esistenza file di input + directory output.

        Args:
            strict: se True solleva ConfigurationError al primo problema;
                    se False raccoglie tutti gli errori e li restituisce.

        Returns:
            Lista descrittiva degli errori (vuota se OK).
        """
        errors: List[str] = []
        required = [
            self.tr_file, self.mapping_file, self.forecast_file,
            self.exclusions_file, self.states_file,
            *self.bom_files.values(),
        ]
        for fname in required:
            p = self.resolve_input(fname)
            if not p.exists():
                errors.append(f"Input mancante: {p}")
            elif not p.is_file():
                errors.append(f"Input non e' un file: {p}")
        if not self.output_dir.exists():
            errors.append(f"Output dir non esiste: {self.output_dir}")
        if strict and errors:
            raise ConfigurationError("; ".join(errors))
        return errors


@dataclass(frozen=True)
class SimulationConfig:
    """Parametri simulazione Monte Carlo + safety stock."""
    n_runs:                  int  = 200
    random_seed:             int  = 42
    start_month_offset:      int  = 0
    use_int16:               bool = True
    max_conflict_iterations: int  = 10
    affidabilita:            str  = "MEDIA"     # NULLA | BASSA | MEDIA | ALTA
    percentile_safety_stock: int  = 99          # 1-99

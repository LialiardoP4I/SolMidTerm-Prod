# -*- coding: utf-8 -*-
"""Configurazione della pipeline: dove stanno i file e quali parametri usare.

In parole semplici: raccoglie in un unico posto i percorsi dei file di
input/output (PipelineConfig) e i parametri della simulazione (SimulationConfig:
numero di scenari, percentile, affidabilita', ecc.), con controlli che bloccano
subito eventuali valori mancanti o errati.
"""
import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from mid_term_supermodels.exceptions import ConfigurationError


def shared_logistica_dir():
    """Cartella 'Dati Logistica' condivisa tra supermodel.

    Se la variabile d'ambiente MIDTERM_SHARED_LOGISTICA e' impostata, usa quel
    percorso; altrimenti il default storico <cwd>/Input/Dati Logistica.
    Fallback = comportamento invariato per il single-supermodel.
    """
    env = os.environ.get('MIDTERM_SHARED_LOGISTICA')
    if env:
        return Path(env)
    return Path('.') / 'Input' / 'Dati Logistica'


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
    sku_senza_lt_filename:          str = "sku_senza_lead_time.xlsx"

    lead_times_file: str = "residual_lead_time.xlsx"
    prices_file:     Optional[str] = "Dati Logistica/PREZZI.XLSX"

    write_intermediate: bool = True
    write_final_output: bool = True
    filter_untraced:    bool = True

    def resolve_input(self, name: str) -> Path:
        """Path assoluto di un file di input (input_dir / name)."""
        return self.input_dir / name

    def resolve_output(self, name: str) -> Path:
        """Path assoluto di un file di output (output_dir / name)."""
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
    affidabilita_model:      str  = "MEDIA"     # L0 Model Mix — NULLA|BASSA|MEDIA|ALTA
    affidabilita_optional:   str  = "MEDIA"     # L1 Optional  — NULLA|BASSA|MEDIA|ALTA
    percentile_safety_stock: int  = 99          # 1-99
    gate_max_mesi:           int  = 4           # GATE max (mesi): cap globale su tutti gli SKU + default per SOLO_MODELLO senza gate


# Chiavi simulazione obbligatorie nel JSON di run
_RUN_REQUIRED_KEYS = (
    "n_runs", "random_seed", "start_month",
    "percentile_safety_stock", "affidabilita_model", "affidabilita_optional",
)


def load_simulation_config(json_path, forecast_month_names) -> "SimulationConfig":
    """Costruisce SimulationConfig da un JSON di run obbligatorio e completo.

    Risolve start_month (nome, es. 'MAY 2026') in start_month_offset usando i
    nomi-mese del forecast. Valida presenza di tutte le chiavi, percentile 1-99
    e livelli affidabilità contro la mappa VSS (Input/affidabilita_config.json).

    Solleva ConfigurationError con messaggio chiaro su qualunque problema.
    """
    from mid_term_supermodels.tr import _AFFIDABILITA_MAP   # lazy: evita import cycle

    p = Path(json_path)
    if not p.is_file():
        raise ConfigurationError(f"JSON di run non trovato: {p}")
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ConfigurationError(f"JSON di run malformato ({p}): {e}")
    if not isinstance(data, dict):
        raise ConfigurationError(
            f"JSON di run deve essere un oggetto, non {type(data).__name__}"
        )

    missing = [k for k in _RUN_REQUIRED_KEYS if k not in data]
    if missing:
        raise ConfigurationError(
            f"JSON di run incompleto ({p}): chiavi mancanti {missing}. "
            f"Richieste tutte: {list(_RUN_REQUIRED_KEYS)}"
        )

    # --- tipi/range ---
    if not isinstance(data["n_runs"], int) or data["n_runs"] <= 0:
        raise ConfigurationError(f"n_runs deve essere int > 0, ricevuto {data['n_runs']!r}")
    if not isinstance(data["random_seed"], int):
        raise ConfigurationError(f"random_seed deve essere int, ricevuto {data['random_seed']!r}")
    pct = data["percentile_safety_stock"]
    if not isinstance(pct, int) or not (1 <= pct <= 99):
        raise ConfigurationError(f"percentile_safety_stock deve essere int 1-99, ricevuto {pct!r}")

    valid_levels = set(_AFFIDABILITA_MAP.keys())
    for key in ("affidabilita_model", "affidabilita_optional"):
        if data[key] not in valid_levels:
            raise ConfigurationError(
                f"{key}='{data[key]}' non valido. Ammessi: {sorted(valid_levels)}"
            )

    # --- gate max (opzionale nel JSON, default 4 mesi) ---
    # Nuova chiave: gate_max_mesi. Fallback alle vecchie chiavi
    # (gate_solo_modello_mesi, residual_lead_time_solo_modello) per retro-compat.
    gate_mx = data.get("gate_max_mesi",
                       data.get("gate_solo_modello_mesi",
                                data.get("residual_lead_time_solo_modello", 4)))
    if not isinstance(gate_mx, int) or gate_mx <= 0:
        raise ConfigurationError(
            f"gate_max_mesi deve essere int > 0, ricevuto {gate_mx!r}"
        )

    # --- risoluzione mese -> offset ---
    start_month = data["start_month"]
    names = list(forecast_month_names)
    if start_month not in names:
        raise ConfigurationError(
            f"start_month='{start_month}' non presente nel forecast. "
            f"Disponibili ({len(names)}): {names}"
        )
    offset = names.index(start_month)

    return SimulationConfig(
        n_runs=data["n_runs"],
        random_seed=data["random_seed"],
        start_month_offset=offset,
        percentile_safety_stock=pct,
        affidabilita_model=data["affidabilita_model"],
        affidabilita_optional=data["affidabilita_optional"],
        gate_max_mesi=gate_mx,
        # use_int16 / max_conflict_iterations: fuori scope JSON -> default dataclass
    )

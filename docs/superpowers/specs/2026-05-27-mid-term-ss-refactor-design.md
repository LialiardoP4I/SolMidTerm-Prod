# Refactoring Pipeline SS Classica — Design

**Data:** 2026-05-27
**Progetto:** SolMidTerm - Prod (Ducati Multistrada V4)
**Perimetro:** pipeline Safety Stock classica, step 0 → step 4
**Esclusioni:** rolling simulation, stock alignment, demand comparison, frontend Streamlit

---

## 1. Perimetro e architettura

### File in scope

Pipeline core "SS Classica" su 6 moduli sorgente:

| Step | File sorgente | Funzione nella pipeline |
|------|---------------|-------------------------|
| 0 | `tr_parser_V2.py` | Parsing TR + Take Rate + Mappatura |
| 1 | `gestione_dipendenze_V2.py` | Parser BOM, espansione OR/IN, gestione bundle |
| 1 | `process_all_bom_combinations.py` | Orchestratore multi-BOM (V21E/V22E/V24T) |
| 1 | `skubundle.py` + `skubundle_OPTIMIZED_V2.py` | Deduplicazione configurazioni |
| 2 | `mc_simulator_V2.py` (senza funzioni rolling) | Monte Carlo + Dirichlet + esclusioni |
| 3-4 | `sku_matching_SS.py` | Matching SKU + percentili + safety stock |

### File esclusi dal refactoring

Restano invariati nella loro forma attuale (potranno essere rifattorizzati in seguito):
`mc_simulator_rolling.py`, `sku_matching_SA.py`, `main_demand_comparison_V2.py`, `stock_alignment_V2.py`, `crea_catalogo_sa.py`, `tr_implied.py`, `optional_cost_allocation.py`, `sensitivity_dirichlet.py`, `genera_report_bugfix.py`, `app_V3.py`.

### Nuova struttura package

```
mid_term/
    __init__.py        API pubblica
    config.py          PipelineConfig + SimulationConfig (dataclass frozen)
    exceptions.py      Gerarchia eccezioni dedicata
    _logging.py        Helper logger + configure_default_logging()
    results.py         Dataclass risultato per ogni step
    tr.py              Ex tr_parser_V2: parser TR + canonical_model_name + ModelMix + CharacteristicGroup
    bom.py             Ex gestione_dipendenze_V2 + skubundle (consolidato) + process_all_bom_combinations
    montecarlo.py      Ex mc_simulator_V2 senza funzioni rolling
    matching.py        Ex sku_matching_SS: NormalizationCache + matching + safety stock
    pipeline.py        Orchestratore SafetyStockPipeline (step 0..4)
```

**Totale: 10 file Python nel package** (6 moduli funzionali + 4 di supporto config/exceptions/logging/results).

### Shim retro-compatibili nella root

Per non rompere `app_V3.py` e altri eventuali consumer non rifattorizzati, restano nella root 6 file Python che riesportano i nomi dal package:

```python
# tr_parser_V2.py (shim)
from mid_term.tr import *
from mid_term.tr import (parse_tr_file, load_monthly_tr, get_tr_for_month,
                          ModelMix, CharacteristicGroup, canonical_model_name)
```

Stesso pattern per `mc_simulator_V2.py`, `sku_matching_SS.py`, `gestione_dipendenze_V2.py`, `skubundle.py`, `process_all_bom_combinations.py`.

Nota su `skubundle_OPTIMIZED_V2.py`: durante il refactoring si verifica se è un duplicato di `skubundle.py` o se ne differisce per logica. Se duplicato, lo shim punta al modulo consolidato; se diverge, le due versioni vengono fuse.

---

## 2. API pubblica + configurazione

### `mid_term/__init__.py`

```python
from mid_term.config     import PipelineConfig, SimulationConfig
from mid_term.pipeline   import SafetyStockPipeline
from mid_term.tr         import canonical_model_name
from mid_term.results    import (TRData, BOMData, SimulationResult,
                                  MatchResult, SafetyStockResult)
from mid_term.exceptions import (MidTermError, ConfigurationError, InputFileError,
                                  TRParseError, BOMParseError, SimulationError,
                                  MatchingError)

__all__ = [
    'PipelineConfig', 'SimulationConfig', 'SafetyStockPipeline',
    'canonical_model_name',
    'TRData', 'BOMData', 'SimulationResult', 'MatchResult', 'SafetyStockResult',
    'MidTermError', 'ConfigurationError', 'InputFileError',
    'TRParseError', 'BOMParseError', 'SimulationError', 'MatchingError',
]
```

### `mid_term/config.py`

```python
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, List

@dataclass(frozen=True)
class PipelineConfig:
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
        """Verifica esistenza input. Lista errori; solleva ConfigurationError se strict."""
        errors = []
        required = [self.tr_file, self.mapping_file, self.forecast_file,
                    self.exclusions_file, self.states_file, *self.bom_files.values()]
        for fname in required:
            p = self.resolve_input(fname)
            if not p.exists():
                errors.append(f"Input mancante: {p}")
            elif not p.is_file():
                errors.append(f"Input non e' un file: {p}")
        if not self.output_dir.exists():
            errors.append(f"Output dir non esiste: {self.output_dir}")
        if strict and errors:
            from mid_term.exceptions import ConfigurationError
            raise ConfigurationError("; ".join(errors))
        return errors


@dataclass(frozen=True)
class SimulationConfig:
    n_runs:                  int  = 200
    random_seed:             int  = 42
    start_month_offset:      int  = 0
    use_int16:               bool = True
    max_conflict_iterations: int  = 10
    affidabilita:            str  = "MEDIA"     # NULLA | BASSA | MEDIA | ALTA
    percentile_safety_stock: int  = 99          # 1-99
```

### `mid_term/results.py`

```python
from dataclasses import dataclass, field
from typing import Dict, List, Tuple
import pandas as pd
import numpy as np

@dataclass
class TRData:
    model_mix:        'ModelMix'
    characteristics:  Dict[str, 'CharacteristicGroup']
    monthly_tr:       Dict[str, Tuple['ModelMix', Dict[str, 'CharacteristicGroup']]]
    forecast_values:  List[float]
    forecast_months:  List[str]
    exclusions:       Dict[str, List[str]]
    n_months_needed:  int
    elapsed_seconds:  float = 0.0

@dataclass
class BOMData:
    catalog:         pd.DataFrame
    matrices:        Dict[str, pd.DataFrame]
    quality_report:  pd.DataFrame
    elapsed_seconds: float = 0.0

@dataclass
class SimulationResult:
    df_wide:         pd.DataFrame
    n_runs:          int
    n_months:        int
    month_names:     List[str]
    elapsed_seconds: float = 0.0

@dataclass
class MatchResult:
    df_matched:      pd.DataFrame
    column_mapping:  Dict[str, str]
    solo_modello:    pd.DataFrame
    elapsed_seconds: float = 0.0

@dataclass
class SafetyStockResult:
    per_sku:         pd.DataFrame
    per_mese:        pd.DataFrame
    percentile:      int
    output_paths:    Dict[str, str] = field(default_factory=dict)
    elapsed_seconds: float = 0.0
```

### `mid_term/pipeline.py`

```python
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.results import (TRData, BOMData, SimulationResult,
                               MatchResult, SafetyStockResult)
from mid_term._logging import get_logger

log = get_logger("pipeline")


class SafetyStockPipeline:
    """Orchestratore end-to-end della pipeline SS Classica.

    Espone sia un metodo run() che esegue tutti gli step in sequenza,
    sia metodi singoli per ogni step (utili per debugging o batch scheduler).
    """

    def __init__(self, config: PipelineConfig,
                 sim_config: SimulationConfig = None,
                 validate_config: bool = True):
        self.config     = config
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

    def run_step_0_load_tr(self)                          -> TRData: ...
    def run_step_1_build_bom(self)                        -> BOMData: ...
    def run_step_2_montecarlo(self, tr: TRData)           -> SimulationResult: ...
    def run_step_3_matching(self, sim: SimulationResult,
                            bom: BOMData)                 -> MatchResult: ...
    def run_step_4_safety_stock(self, mat: MatchResult,
                                bom: BOMData)             -> SafetyStockResult: ...
```

### Esempio d'uso

```python
from pathlib import Path
from mid_term import SafetyStockPipeline, PipelineConfig, SimulationConfig

config = PipelineConfig(
    input_dir  = Path("./Input"),
    output_dir = Path("./Output"),
)
sim = SimulationConfig(n_runs=500, random_seed=42, affidabilita="MEDIA")
result = SafetyStockPipeline(config, sim).run()
```

---

## 3. Logging, errori, validazione, test

### Logging

`mid_term/_logging.py` espone:

```python
def get_logger(name: str) -> logging.Logger:
    """Ritorna logger 'mid_term.<name>' (un logger per modulo)."""

def configure_default_logging(level: int = logging.INFO,
                              format_str: str = "%(asctime)s [%(levelname)s] %(name)s: %(message)s") -> None:
    """Configura stdout handler. Consumer puo' override."""
```

Nessun `print()` nei moduli core. Tutti i messaggi attuali (warnings TR=0, completamenti step, statistiche) mappati su `INFO`/`WARNING`/`ERROR`. Il consumer (CLI, FastAPI, Airflow) eredita o sovrascrive la configurazione tramite root logger.

### Eccezioni

`mid_term/exceptions.py`:

```python
class MidTermError(Exception): ...
class ConfigurationError(MidTermError): ...
class InputFileError(MidTermError): ...
class TRParseError(MidTermError): ...
class BOMParseError(MidTermError): ...
class SimulationError(MidTermError): ...
class MatchingError(MidTermError): ...
```

I moduli sollevano eccezioni specifiche con messaggi descrittivi. Niente più `try/except Exception` silenziosi, niente `sys.exit()`, niente `raise ValueError` generici.

### Validazione

`PipelineConfig.validate()` è chiamato automaticamente da `SafetyStockPipeline.__init__` (disattivabile con `validate_config=False`). Verifica esistenza file e directory output, raccoglie tutti gli errori in batch o solleva al primo (modo strict).

### Rimozione tracce "Claude"

Grep su tutto il package per qualsiasi menzione di `Claude` (case-insensitive, commenti, docstring, log, variabili). Eliminata ogni occorrenza.

### Test di regressione

`tests/test_regression.py` esegue la pipeline su `TR TOTALV21E - Copia.xlsx` con `n_runs=20, seed=42` e confronta:

1. **Test logica MC (step 2):** hash dei nomi modello + somma colonne Monte Carlo. Garantisce sequenze RNG e Dirichlet identici.
2. **Test canonicalizzazione:** ogni nome modello in output è ASCII pulito, uppercase.
3. **Test idempotenza `canonical_model_name`:** suite di casi (NBSP, ZWS, ZWNJ, BOM, casing).
4. **Test end-to-end:** hash output `sku_safety_stock_PER_MESE` deve matchare baseline pre-refactoring.

I valori di hash attesi vengono calcolati al primo run del refactoring e fissati come costanti nel test.

### Tracciamento profiling

Ogni risultato (`TRData`, `BOMData`, ecc.) porta un campo `elapsed_seconds`. Ogni step logga su `INFO` durata e statistiche chiave.

---

## 4. Piano di esecuzione

### Ordine refactoring (bottom-up)

| Fase | Modulo | Dipendenze | Verifica |
|------|--------|------------|----------|
| 1 | `config.py`, `exceptions.py`, `_logging.py`, `results.py` | — | import OK |
| 2 | `tr.py` | config, results | `canonical_model_name` + parser TR identici a baseline |
| 3 | `bom.py` | tr | hash `tutte_righe_univoche_V2.xlsx` identico |
| 4 | `montecarlo.py` | tr | fingerprint df_wide identico (test già verificato) |
| 5 | `matching.py` | tr, results | hash output safety stock identico |
| 6 | `pipeline.py` | tutti | end-to-end identico |
| 7 | Shim retro-compat in root | tutti | `app_V3.py` parte e gira |

### Cosa viene rimosso

- Tutti i `print()` → logging strutturato.
- Docstring multipagina che descrivono "differenze rispetto a V1".
- Variabili globali (`_norm_cache`, `_MODULE_DIR`, `_AFFIDABILITA_MAP` cache) → attributi di classe o parametri.
- Path hardcoded (`.\\Input\\...`, `.\\Output\\...`) → `PipelineConfig`.
- Funzioni rolling (`rescale_alphas`, qualsiasi cosa con `mese_lancio`).
- `main()`, `main_simulator()`, `main_sku_v2()`, blocchi `if __name__ == "__main__"` → orchestrazione in `pipeline.py`.
- Codice morto / commentato.
- Riferimenti "Claude" ovunque presenti.
- Duplicazione `skubundle.py` ↔ `skubundle_OPTIMIZED_V2.py`: dopo confronto, una versione viene consolidata.

### Cosa NON viene toccato (logica preservata)

- Algoritmo Monte Carlo (Dirichlet, esclusioni, conflict resolution).
- Parsing BOM (espansione OR/IN, gestione bundle, RIMS_Substitutions).
- Matching SKU (column_mapping auto, NormalizationCache, percentili 95/99).
- Set hardcoded `{'msv4','msv4s','msv4pp','msv4ri'}` in `apply_pack_colors_to_wide_df`.
- Alias MSV4-family in `CharacteristicGroup._lookup` (legacy, mantenuto per sicurezza).
- Take rate, alphas, percentili, sequenze RNG, ordini di iterazione.

### Criterio di stop

Refactoring considerato completo solo quando:

1. Tutti i moduli `mid_term/*.py` esistono e sono importabili.
2. Tutti i 6 shim retro-compatibili in root esistono e riesportano l'API pubblica.
3. `app_V3.py` parte e gira senza modifiche.
4. Test di regressione end-to-end passa: hash output safety stock identico a baseline pre-refactoring.
5. Nessuna citazione "Claude" residua nel codice.
6. `mypy mid_term/ --ignore-missing-imports` non riporta errori bloccanti.

---

## Appendice A — Mappatura simbolo vecchio → nuovo

| Vecchio nome | Modulo origine | Nuovo nome | Modulo destinazione |
|--------------|----------------|------------|---------------------|
| `parse_tr_file` | `tr_parser_V2` | `parse_tr_file` | `mid_term.tr` |
| `load_monthly_tr` | `tr_parser_V2` | `load_monthly_tr` | `mid_term.tr` |
| `get_tr_for_month` | `tr_parser_V2` | `get_tr_for_month` | `mid_term.tr` |
| `parse_model_mix` | `tr_parser_V2` | `parse_model_mix` | `mid_term.tr` |
| `ModelMix` | `tr_parser_V2` | `ModelMix` | `mid_term.tr` |
| `CharacteristicGroup` | `tr_parser_V2` | `CharacteristicGroup` | `mid_term.tr` |
| `canonical_model_name` | `tr_parser_V2` | `canonical_model_name` | `mid_term.tr` |
| `process_excel_to_matrix_v2` | `gestione_dipendenze_V2` | `process_excel_to_matrix` | `mid_term.bom` |
| `process_combination` | `process_all_bom_combinations` | `process_combination` | `mid_term.bom` |
| `main` (BOM) | `process_all_bom_combinations` | (rimosso, orchestrato da `pipeline`) | — |
| `skubundle_*` | `skubundle*` | `dedup_configurations` | `mid_term.bom` |
| `SimulationConfig` | `mc_simulator_V2` | `SimulationConfig` | `mid_term.config` |
| `run_monthly_simulations_optimized` | `mc_simulator_V2` | `run_monthly_simulations` | `mid_term.montecarlo` |
| `run_single_month_simulation_optimized` | `mc_simulator_V2` | `run_single_month` | `mid_term.montecarlo` |
| `build_wide_dataframe` | `mc_simulator_V2` | `build_wide_dataframe` | `mid_term.montecarlo` |
| `apply_pack_colors_to_wide_df` | `mc_simulator_V2` | `apply_pack_colors` | `mid_term.montecarlo` |
| `main_simulator` | `mc_simulator_V2` | (rimosso, orchestrato da `pipeline`) | — |
| `rescale_alphas` | `mc_simulator_V2` | (escluso: usato solo dal rolling) | — |
| `match_skus_ultra_optimized` | `sku_matching_SS` | `match_skus` | `mid_term.matching` |
| `main_sku_v2` | `sku_matching_SS` | (rimosso, orchestrato da `pipeline`) | — |
| `NormalizationCache` | `sku_matching_SS` | `NormalizationCache` | `mid_term.matching` |
| `normalize_value_cached` | `sku_matching_SS` | `normalize_value_cached` | `mid_term.matching` |

## Appendice B — Contenuto dei file shim

Ogni shim ha solo riesportazioni, nessuna logica:

```python
# tr_parser_V2.py
from mid_term.tr import *  # noqa: F401,F403
from mid_term.tr import (parse_tr_file, load_monthly_tr, get_tr_for_month,
                          parse_model_mix, ModelMix, CharacteristicGroup,
                          canonical_model_name)

# mc_simulator_V2.py
from mid_term.montecarlo import *  # noqa: F401,F403
from mid_term.montecarlo import (SimulationConfig, run_monthly_simulations,
                                  run_single_month, build_wide_dataframe,
                                  apply_pack_colors)
# Alias retrocompat
from mid_term.montecarlo import run_monthly_simulations as run_monthly_simulations_optimized
from mid_term.montecarlo import run_single_month as run_single_month_simulation_optimized
from mid_term.montecarlo import apply_pack_colors as apply_pack_colors_to_wide_df

# sku_matching_SS.py
from mid_term.matching import *  # noqa: F401,F403
from mid_term.matching import (match_skus, NormalizationCache, normalize_value_cached)
from mid_term.matching import match_skus as match_skus_ultra_optimized

# gestione_dipendenze_V2.py
from mid_term.bom import *  # noqa: F401,F403
from mid_term.bom import process_excel_to_matrix as process_excel_to_matrix_v2

# skubundle.py + skubundle_OPTIMIZED_V2.py
from mid_term.bom import dedup_configurations

# process_all_bom_combinations.py
from mid_term.bom import process_combination
```

## Appendice C — Test di regressione (snippet)

```python
import hashlib
from pathlib import Path
import pytest
from mid_term import SafetyStockPipeline, PipelineConfig, SimulationConfig, canonical_model_name

TR_FIXTURE = Path("Input/TR TOTALV21E - Copia.xlsx")
EXPECTED_MC_FINGERPRINT  = "<da popolare al primo run post-refactoring>"
EXPECTED_SS_PER_SKU_HASH = "<da popolare al primo run post-refactoring>"

@pytest.fixture
def pipeline(tmp_path):
    if not TR_FIXTURE.exists():
        pytest.skip("Fixture TR non disponibile")
    cfg = PipelineConfig(input_dir=TR_FIXTURE.parent, output_dir=tmp_path,
                         write_intermediate=False, write_final_output=False)
    sim = SimulationConfig(n_runs=20, random_seed=42, percentile_safety_stock=99)
    return SafetyStockPipeline(cfg, sim, validate_config=False)

def test_canonical_idempotent():
    assert canonical_model_name("MSV4​") == "MSV4"
    assert canonical_model_name("PAN7GSP")    == "PAN7GSP"
    assert canonical_model_name(None)         == ""
    assert canonical_model_name(float("nan")) == ""
    assert canonical_model_name("  msv4s  ")  == "MSV4S"

def test_step_2_fingerprint(pipeline):
    tr  = pipeline.run_step_0_load_tr()
    sim = pipeline.run_step_2_montecarlo(tr)
    parts = []
    for col in sorted(sim.df_wide.columns):
        parts.append(f"{col}:{sim.df_wide[col].sum() if sim.df_wide[col].dtype != object else hashlib.sha256('|'.join(map(str, sim.df_wide[col])).encode()).hexdigest()}")
    fp = hashlib.sha256("\n".join(parts).encode()).hexdigest()
    assert fp == EXPECTED_MC_FINGERPRINT

def test_canonical_names_in_output(pipeline):
    tr = pipeline.run_step_0_load_tr()
    for m in tr.model_mix.models:
        assert "​" not in m
        assert " " not in m
        assert m == m.strip().upper()
```

---

## Storia revisioni

| Data | Autore | Modifica |
|------|--------|----------|
| 2026-05-27 | AntonioLiardo-SMOPS | Versione iniziale del design |

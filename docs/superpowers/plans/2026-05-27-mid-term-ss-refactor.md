# Refactoring Pipeline SS Classica — Piano Implementazione

> **Per worker agentici:** SOTTO-SKILL RICHIESTO — usare `superpowers:subagent-driven-development` (consigliato) o `superpowers:executing-plans` per eseguire questo piano task-per-task. Gli step usano checkbox `- [ ]` per tracciamento.

**Goal:** Rifattorizzare la pipeline Safety Stock Classica (step 0..4) in un package `mid_term/` pulito e pronto all'integrazione su piattaforma, mantenendo logica e output runtime invariati bit-per-bit sui dati attuali.

**Architettura:** Package Python a 10 file (6 moduli funzionali + 4 di supporto) con API pubblica esplicita (`SafetyStockPipeline`, `PipelineConfig`, `SimulationConfig`), configurazione via dataclass frozen, logging strutturato sostituisce `print()`, gerarchia eccezioni dedicata, test di regressione end-to-end con fingerprint hash.

**Tech Stack:** Python 3.13, pandas, numpy, openpyxl, dataclasses, logging, pytest.

**Riferimento spec:** `docs/superpowers/specs/2026-05-27-mid-term-ss-refactor-design.md`

**Verifica trasversale:** dopo ogni fase di refactoring di un modulo, il test di regressione end-to-end (Task 1.7) deve passare con hash identico al baseline. Se non passa, fermarsi e diagnosticare prima di proseguire.

---

## Fase 0 — Setup directory e baseline

### Task 0.1: Crea struttura directory package + tests

**Files:**
- Create: `mid_term/__init__.py` (vuoto temporaneo)
- Create: `tests/__init__.py` (vuoto)
- Create: `tests/fixtures/.gitkeep` (vuoto)

- [ ] **Step 1: Crea directory e file vuoti**

```bash
cd "C:/Users/AntonioLiardo-SMOPS/OneDrive - Digital360/Desktop/Ducati/SolMidTerm - Prod"
mkdir -p mid_term tests/fixtures
touch mid_term/__init__.py tests/__init__.py tests/fixtures/.gitkeep
```

- [ ] **Step 2: Copia fixture TR per test**

```bash
cp "Input/TR TOTALV21E - Copia.xlsx" "tests/fixtures/TR_test.xlsx"
```

- [ ] **Step 3: Verifica struttura**

```bash
ls mid_term/ tests/ tests/fixtures/
```
Expected: `__init__.py` in `mid_term/` e `tests/`, `TR_test.xlsx` in `tests/fixtures/`.

- [ ] **Step 4: Commit**

```bash
git add mid_term/ tests/
git commit -m "chore: setup directory mid_term/ + tests/"
```

---

### Task 0.2: Cattura baseline fingerprint pipeline corrente

**Files:**
- Create: `tests/_capture_baseline.py`
- Create: `tests/_baseline_fingerprints.json`

- [ ] **Step 1: Scrivi script di cattura**

```python
# tests/_capture_baseline.py
"""Cattura fingerprint output pipeline corrente per regression test post-refactor.

Esegue Monte Carlo + matching sui dati reali con seed fissato e salva hash
deterministici di tutti gli output rilevanti. Eseguire UNA VOLTA prima del
refactoring; i fingerprint diventano il riferimento contro cui ogni step
del refactoring viene verificato.
"""
import hashlib
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd

from tr_parser_V2 import load_monthly_tr, parse_tr_file
from mc_simulator_V2 import SimulationConfig, run_monthly_simulations_optimized, build_wide_dataframe


TR_PATH = Path('tests/fixtures/TR_test.xlsx')
FORECAST = pd.Series({'Jan 2026': 100, 'Feb 2026': 100, 'Mar 2026': 100}, name='demand')


def hash_dataframe(df: pd.DataFrame) -> str:
    """Hash deterministico di un DataFrame: somma per colonna numerica + hash colonne object."""
    parts = []
    for col in sorted(df.columns):
        if df[col].dtype == object:
            parts.append(f'{col}:{hashlib.sha256("|".join(map(str, df[col])).encode()).hexdigest()}')
        else:
            parts.append(f'{col}:sum={df[col].sum()}:nunique={df[col].nunique()}')
    return hashlib.sha256('\n'.join(parts).encode()).hexdigest()


def hash_mc_results(results: dict) -> str:
    """Hash deterministico del dict risultati Monte Carlo."""
    parts = []
    for month_idx in sorted(results.keys()):
        for cfg_key in sorted(results[month_idx].keys(), key=str):
            arr = results[month_idx][cfg_key]
            parts.append(f'm{month_idx}|{cfg_key}|sum={arr.sum()}|n={len(arr)}')
    return hashlib.sha256('\n'.join(parts).encode()).hexdigest()


def main():
    monthly_tr = load_monthly_tr(str(TR_PATH))
    model_mix_global, characteristics_global = parse_tr_file(str(TR_PATH))

    cfg = SimulationConfig(n_runs=20, random_seed=42, use_int16=True,
                           max_conflict_resolution_iterations=10, start_month_offset=0)

    results = run_monthly_simulations_optimized(
        forecast_values=list(FORECAST.values),
        month_names=list(FORECAST.index),
        n_months=len(FORECAST),
        model_mix=model_mix_global,
        characteristics=characteristics_global,
        exclusions={},
        config=cfg,
        monthly_tr=monthly_tr,
    )

    fingerprints = {
        'tr_models':       sorted(monthly_tr['Jan 2026'][0].models),
        'tr_alphas_sum':   float(monthly_tr['Jan 2026'][0].alphas.sum()),
        'n_characteristics': len(characteristics_global),
        'mc_fingerprint':  hash_mc_results(results),
        'mc_total_bikes':  int(sum(arr.sum() for month in results.values() for arr in month.values())),
        'mc_n_combos':     int(sum(len(month) for month in results.values())),
    }

    out_path = Path('tests/_baseline_fingerprints.json')
    out_path.write_text(json.dumps(fingerprints, indent=2, ensure_ascii=False, default=str),
                        encoding='utf-8')
    print(f'Baseline fingerprints salvati in {out_path}')
    print(json.dumps(fingerprints, indent=2, default=str))


if __name__ == '__main__':
    main()
```

- [ ] **Step 2: Esegui cattura baseline**

```bash
python -X utf8 tests/_capture_baseline.py
```
Expected: stampa `Baseline fingerprints salvati` + JSON con campi `tr_models`, `tr_alphas_sum`, `n_characteristics`, `mc_fingerprint`, `mc_total_bikes`, `mc_n_combos`.

- [ ] **Step 3: Verifica fingerprint stabile (re-run)**

```bash
python -X utf8 tests/_capture_baseline.py > /tmp/_baseline_rerun.txt 2>&1
diff tests/_baseline_fingerprints.json <(python -X utf8 -c "import json; print(json.dumps(json.load(open('tests/_baseline_fingerprints.json')), indent=2))")
```
Expected: nessuna differenza fra due run consecutivi (determinismo seed OK).

- [ ] **Step 4: Commit baseline**

```bash
git add tests/_capture_baseline.py tests/_baseline_fingerprints.json
git commit -m "test: cattura baseline fingerprints pipeline pre-refactoring"
```

---

## Fase 1 — Fondamenta (config, exceptions, logging, results)

### Task 1.1: Eccezioni

**Files:**
- Create: `mid_term/exceptions.py`

- [ ] **Step 1: Scrivi exceptions.py**

```python
# mid_term/exceptions.py
"""Gerarchia eccezioni della pipeline SS Classica."""


class MidTermError(Exception):
    """Base di tutte le eccezioni del package."""


class ConfigurationError(MidTermError):
    """Configurazione invalida: path mancanti, parametri fuori range."""


class InputFileError(MidTermError):
    """File di input atteso ma assente o malformato."""


class TRParseError(MidTermError):
    """Errore parsing Take Rate / Mappatura (step 0)."""


class BOMParseError(MidTermError):
    """Errore parsing BOM o espansione dipendenze (step 1)."""


class SimulationError(MidTermError):
    """Errore nel Monte Carlo: mix non valido, Dirichlet con alpha NaN, ecc. (step 2)."""


class MatchingError(MidTermError):
    """Errore nel matching SKU o nel calcolo safety stock (step 3-4)."""
```

- [ ] **Step 2: Verifica import**

```bash
python -c "from mid_term.exceptions import MidTermError, ConfigurationError, TRParseError, BOMParseError, SimulationError, MatchingError, InputFileError; print('OK')"
```
Expected: stampa `OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/exceptions.py
git commit -m "feat(mid_term): gerarchia eccezioni dedicata"
```

---

### Task 1.2: Logging

**Files:**
- Create: `mid_term/_logging.py`

- [ ] **Step 1: Scrivi _logging.py**

```python
# mid_term/_logging.py
"""Helper logging per il package mid_term."""
import logging


def get_logger(name: str) -> logging.Logger:
    """Ritorna logger mid_term.<name>."""
    return logging.getLogger(f"mid_term.{name}")


def configure_default_logging(
    level: int = logging.INFO,
    format_str: str = "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
) -> None:
    """Configura handler stdout sul logger root mid_term.

    Il consumer puo' sovrascrivere o ignorare questa configurazione
    fornendo la propria via logging.basicConfig o handler dedicati.
    """
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(format_str))
    root = logging.getLogger("mid_term")
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(level)
    root.propagate = False
```

- [ ] **Step 2: Verifica import e logger**

```bash
python -c "from mid_term._logging import get_logger, configure_default_logging; configure_default_logging(); log = get_logger('test'); log.info('hello'); log.warning('warn')"
```
Expected: stampa due righe `[INFO]` e `[WARNING]` con prefisso `mid_term.test`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/_logging.py
git commit -m "feat(mid_term): helper logging strutturato"
```

---

### Task 1.3: Risultati (dataclass)

**Files:**
- Create: `mid_term/results.py`

- [ ] **Step 1: Scrivi results.py**

```python
# mid_term/results.py
"""Dataclass risultato per ogni step della pipeline.

Le definizioni di ModelMix e CharacteristicGroup sono in mid_term.tr;
qui si usano come stringhe per evitare import circolari.
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
    """Output step 1: catalogo configurazioni unico + matrici per supermodel."""
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
    elapsed_seconds: float = 0.0


@dataclass
class SafetyStockResult:
    """Output step 4: safety stock per SKU e per mese."""
    per_sku:         pd.DataFrame
    per_mese:        pd.DataFrame
    percentile:      int
    output_paths:    Dict[str, str] = field(default_factory=dict)
    elapsed_seconds: float = 0.0
```

- [ ] **Step 2: Verifica import**

```bash
python -c "from mid_term.results import TRData, BOMData, SimulationResult, MatchResult, SafetyStockResult; print('OK')"
```
Expected: stampa `OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/results.py
git commit -m "feat(mid_term): dataclass risultato step 0..4"
```

---

### Task 1.4: Configurazione

**Files:**
- Create: `mid_term/config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1: Scrivi test config**

```python
# tests/test_config.py
import pytest
from pathlib import Path
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.exceptions import ConfigurationError


def test_pipeline_config_resolve_paths(tmp_path):
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    assert cfg.resolve_input("foo.xlsx") == tmp_path / "Input" / "foo.xlsx"
    assert cfg.resolve_output("bar.xlsx") == tmp_path / "Output" / "bar.xlsx"


def test_pipeline_config_validate_missing_input(tmp_path):
    """Validate solleva ConfigurationError se mancano file di input."""
    (tmp_path / "Output").mkdir()
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    with pytest.raises(ConfigurationError):
        cfg.validate(strict=True)


def test_pipeline_config_validate_non_strict_returns_errors(tmp_path):
    (tmp_path / "Output").mkdir()
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    errors = cfg.validate(strict=False)
    assert errors and any("mancante" in e.lower() for e in errors)


def test_simulation_config_defaults():
    sim = SimulationConfig()
    assert sim.n_runs == 200
    assert sim.random_seed == 42
    assert sim.percentile_safety_stock == 99
    assert sim.affidabilita == "MEDIA"
```

- [ ] **Step 2: Esegui test (deve fallire — config.py non esiste)**

```bash
python -m pytest tests/test_config.py -v
```
Expected: ImportError o ModuleNotFoundError su `mid_term.config`.

- [ ] **Step 3: Scrivi config.py**

```python
# mid_term/config.py
"""Configurazione dichiarativa della pipeline."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from mid_term.exceptions import ConfigurationError


@dataclass(frozen=True)
class PipelineConfig:
    """Configurazione path file + opzioni I/O.

    Frozen + immutable: passabile cross-thread, hashable.
    """
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
        """Verifica esistenza input + output_dir. Lista errori; solleva se strict."""
        errors: List[str] = []
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
```

- [ ] **Step 4: Esegui test (deve passare)**

```bash
python -m pytest tests/test_config.py -v
```
Expected: 4 test PASSED.

- [ ] **Step 5: Commit**

```bash
git add mid_term/config.py tests/test_config.py
git commit -m "feat(mid_term): PipelineConfig + SimulationConfig con validate"
```

---

### Task 1.5: __init__ vuoto API placeholder

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Scrivi __init__ con import base (esporremo gradualmente)**

```python
# mid_term/__init__.py
"""Mid-Term Safety Stock Pipeline.

Package per il calcolo del Safety Stock di componenti tramite simulazione
Monte Carlo + matching SKU su distinte base multi-supermodel.

API pubblica (in espansione mano a mano che i moduli vengono migrati):
    PipelineConfig, SimulationConfig — configurazione dichiarativa
    Eccezioni: MidTermError + sotto-classi specifiche per ogni step
"""
from mid_term.config     import PipelineConfig, SimulationConfig
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
```

- [ ] **Step 2: Verifica import package**

```bash
python -c "import mid_term; print(mid_term.__all__)"
```
Expected: stampa lista degli export.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): __init__ con API base (config + exceptions)"
```

---

## Fase 2 — Modulo `tr.py` (step 0)

### Task 2.1: Crea tr.py copiando + ripulendo tr_parser_V2

**Files:**
- Create: `mid_term/tr.py`

- [ ] **Step 1: Copia tr_parser_V2.py in mid_term/tr.py**

```bash
cp tr_parser_V2.py mid_term/tr.py
```

- [ ] **Step 2: In `mid_term/tr.py`, sostituisci tutti `print(...)` con logger**

Apri `mid_term/tr.py`. In cima ai moduli aggiungi:
```python
from mid_term._logging import get_logger
log = get_logger("tr")
```
Poi `Edit` su ogni `print(...)`:
- `print(f"  [WARN] affidabilita_config.json non caricato ({e}) -> uso default")` → `log.warning("affidabilita_config.json non caricato (%s) -> default", e)`
- `print(f"Warning: {self.name} [{model_name}] sums to {total:.3f}")` → `log.warning("%s [%s] TR sum = %.3f (atteso ~1.0)", self.name, model_name, total)`
- `print(f"Warning: Model TR sums to {total:.3f}")` → `log.warning("Model TR sum = %.3f (atteso ~1.0)", total)`
- `print(f"  [TR_MONTHLY] ...")` → `log.info(...)`
- `print(f"  [OK] ...")` → `log.info(...)`
- `print(f"  [SKIP] ...")` → `log.info(...)`
- `print(f"  [CharGroup:{self.name}] modello '{model_name}' non trovato -> uso '{first_key}' come fallback")` → `log.warning("CharGroup[%s]: modello '%s' non trovato, fallback su '%s'", self.name, model_name, first_key)`

- [ ] **Step 3: Strip docstring "DIFFERENZE RISPETTO A V1"**

In `mid_term/tr.py` rimuovere o ridurre a 1-2 righe le docstring che contengono "DIFFERENZE RISPETTO A V1" / "OTTIMIZZAZIONI APPLICATE" / "FORMATO NUOVO / FORMATO VECCHIO". Mantenere solo il corpo descrittivo (cosa fa la funzione, args, returns).

- [ ] **Step 4: Verifica import + regression canonical_model_name**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from mid_term.tr import canonical_model_name, load_monthly_tr, parse_tr_file
assert canonical_model_name('MSV4​') == 'MSV4'
assert canonical_model_name('PAN7GSP') == 'PAN7GSP'
assert canonical_model_name(None) == ''
assert canonical_model_name(float('nan')) == ''
monthly = load_monthly_tr('tests/fixtures/TR_test.xlsx')
mm, _ = parse_tr_file('tests/fixtures/TR_test.xlsx')
assert sorted(monthly['Jan 2026'][0].models) == ['MSV4', 'MSV4PP', 'MSV4RI', 'MSV4RS', 'MSV4S']
assert sorted(mm.models) == ['MSV4', 'MSV4PP', 'MSV4RI', 'MSV4RS', 'MSV4S']
print('OK')
"
```
Expected: stampa `OK`.

- [ ] **Step 5: Commit**

```bash
git add mid_term/tr.py
git commit -m "feat(mid_term): modulo tr.py — parsing TR + canonical_model_name"
```

---

### Task 2.2: Esporre simboli pubblici in __init__

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Aggiungi export tr.py**

Modifica `mid_term/__init__.py` aggiungendo:
```python
from mid_term.tr import (
    canonical_model_name,
    ModelMix,
    CharacteristicGroup,
    parse_tr_file,
    parse_model_mix,
    load_monthly_tr,
    get_tr_for_month,
)
```
e aggiungi questi nomi a `__all__`.

- [ ] **Step 2: Verifica import**

```bash
python -c "from mid_term import canonical_model_name, ModelMix, CharacteristicGroup, parse_tr_file, load_monthly_tr, get_tr_for_month; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): esporta API modulo tr"
```

---

## Fase 3 — Modulo `bom.py` (step 1)

### Task 3.1: Crea bom.py consolidando gestione_dipendenze_V2

**Files:**
- Create: `mid_term/bom.py`

- [ ] **Step 1: Copia gestione_dipendenze_V2.py**

```bash
cp gestione_dipendenze_V2.py mid_term/bom.py
```

- [ ] **Step 2: Sostituisci import e print**

In cima a `mid_term/bom.py` aggiungi:
```python
from mid_term._logging import get_logger
from mid_term.exceptions import BOMParseError
log = get_logger("bom")
```

Sostituisci tutti i `print(...)` con `log.info/warning/error` (regola identica a tr.py).

- [ ] **Step 3: Strip docstring "HARDCODING ELIMINATI"**

Ridurre il preambolo del modulo a 5-8 righe descrittive su cosa fa, eliminando l'elenco numerato dei "HARDCODING ELIMINATI" e "SCALABILITÀ".

- [ ] **Step 4: Verifica import**

```bash
python -c "from mid_term.bom import process_excel_to_matrix; print('OK')"
```
Expected: `OK`.

- [ ] **Step 5: Commit**

```bash
git add mid_term/bom.py
git commit -m "feat(mid_term): bom.py — parser BOM + espansione dipendenze"
```

---

### Task 3.2: Aggiungi deduplicazione skubundle a bom.py

**Files:**
- Modify: `mid_term/bom.py`
- Read: `skubundle.py` e `skubundle_OPTIMIZED_V2.py` per confronto

- [ ] **Step 1: Confronta skubundle.py e skubundle_OPTIMIZED_V2.py**

```bash
diff skubundle.py skubundle_OPTIMIZED_V2.py | head -100
```
Determina: sono identici, divergenti, o uno deprecato. Annota nel commit message.

- [ ] **Step 2: Estrai funzione principale dedup da skubundle**

Cerca la funzione pubblica in `skubundle.py` (es. `crea_tutte_righe_univoche` o `dedup_*` o `process_*`). Identifica firma e dipendenze.

```bash
grep -n "^def " skubundle.py | head -20
grep -n "^def " skubundle_OPTIMIZED_V2.py | head -20
```

- [ ] **Step 3: Copia in mid_term/bom.py la funzione + helper**

Appendi in fondo a `mid_term/bom.py` la funzione di deduplicazione (la versione `_OPTIMIZED_V2` se ZERO HARDCODING come da docstring, altrimenti la versione `skubundle.py`). Mantieni il nome funzione originale per ora (la rinominiamo nei shim a Task 7).

Sostituisci print → logger come negli step precedenti.

- [ ] **Step 4: Verifica import**

```bash
python -c "from mid_term.bom import process_excel_to_matrix; import mid_term.bom as b; print([n for n in dir(b) if not n.startswith('_')][:20])"
```
Expected: stampa lista funzioni esposte da bom.py incluse quelle dedup.

- [ ] **Step 5: Commit**

```bash
git add mid_term/bom.py
git commit -m "feat(mid_term): bom.py include deduplicazione skubundle"
```

---

### Task 3.3: Aggiungi orchestratore process_all_bom_combinations a bom.py

**Files:**
- Modify: `mid_term/bom.py`

- [ ] **Step 1: Estrai logica process_combination + run_quality_checks**

Leggi `process_all_bom_combinations.py` linee 267-606. Identifica `run_quality_checks`, `process_combination`, e l'orchestratore `main`.

- [ ] **Step 2: Copia in mid_term/bom.py senza il `main()` e senza `if __name__ == "__main__"`**

Appendi a `mid_term/bom.py` le funzioni `run_quality_checks`, `process_combination` adattate. Il `main()` originale e i path hardcoded NON vanno copiati: la loro logica andrà nel `pipeline.run_step_1_build_bom()` in Fase 6.

Sostituisci print → logger.

- [ ] **Step 3: Verifica import**

```bash
python -c "from mid_term.bom import process_combination, run_quality_checks, process_excel_to_matrix; print('OK')"
```
Expected: `OK`.

- [ ] **Step 4: Commit**

```bash
git add mid_term/bom.py
git commit -m "feat(mid_term): bom.py include process_combination + quality checks"
```

---

### Task 3.4: Aggiorna export __init__

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Aggiungi export bom.py (solo simboli pubblici)**

```python
from mid_term.bom import process_excel_to_matrix, process_combination
```
e aggiungi a `__all__`.

- [ ] **Step 2: Verifica import**

```bash
python -c "from mid_term import process_excel_to_matrix, process_combination; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): esporta API modulo bom"
```

---

## Fase 4 — Modulo `montecarlo.py` (step 2)

### Task 4.1: Copia mc_simulator_V2 escludendo funzioni rolling

**Files:**
- Create: `mid_term/montecarlo.py`

- [ ] **Step 1: Identifica funzioni rolling in mc_simulator_V2**

```bash
grep -n "^def " mc_simulator_V2.py | grep -iE "(rolling|rescale_alphas)"
```
Annota le funzioni da escludere (probabilmente `rescale_alphas` e qualsiasi cosa con `mese_lancio`).

- [ ] **Step 2: Copia mc_simulator_V2.py in mid_term/montecarlo.py**

```bash
cp mc_simulator_V2.py mid_term/montecarlo.py
```

- [ ] **Step 3: Rimuovi funzioni rolling identificate allo step 1**

Aprire `mid_term/montecarlo.py` e rimuovere le definizioni delle funzioni rolling/rescale_alphas. Verificare che nulla nello stesso file le importi.

- [ ] **Step 4: Rimuovi main_simulator e blocco `if __name__ == "__main__"`**

Cancella la funzione `main_simulator` (linee ~1365-1500) e ogni blocco `if __name__ == "__main__"` con relativi path hardcoded.

- [ ] **Step 5: Aggiungi logger + sostituisci print**

In cima a `mid_term/montecarlo.py`:
```python
from mid_term._logging import get_logger
from mid_term.exceptions import SimulationError
log = get_logger("montecarlo")
```
Sostituisci tutti i `print(...)` rimanenti con `log.info/warning/error`.

- [ ] **Step 6: Aggiorna import interni a mid_term**

Cambia `from tr_parser_V2 import (...)` in `from mid_term.tr import (...)`.

- [ ] **Step 7: Verifica import**

```bash
python -c "from mid_term.montecarlo import SimulationConfig, run_monthly_simulations_optimized, run_single_month_simulation_optimized, build_wide_dataframe, apply_pack_colors_to_wide_df; print('OK')"
```
Expected: `OK`.

- [ ] **Step 8: Commit**

```bash
git add mid_term/montecarlo.py
git commit -m "feat(mid_term): montecarlo.py — MC simulator senza funzioni rolling"
```

---

### Task 4.2: Test regressione MC fingerprint

**Files:**
- Create: `tests/test_regression_mc.py`

- [ ] **Step 1: Scrivi test**

```python
# tests/test_regression_mc.py
"""Verifica che il modulo montecarlo riproduca il fingerprint baseline."""
import hashlib
import json
from pathlib import Path

import pandas as pd
import pytest

from mid_term.tr import load_monthly_tr, parse_tr_file
from mid_term.montecarlo import SimulationConfig, run_monthly_simulations_optimized


TR_FIXTURE = Path("tests/fixtures/TR_test.xlsx")
BASELINE = Path("tests/_baseline_fingerprints.json")
FORECAST = pd.Series({"Jan 2026": 100, "Feb 2026": 100, "Mar 2026": 100}, name="demand")


def _hash_mc_results(results):
    parts = []
    for month_idx in sorted(results.keys()):
        for cfg_key in sorted(results[month_idx].keys(), key=str):
            arr = results[month_idx][cfg_key]
            parts.append(f"m{month_idx}|{cfg_key}|sum={arr.sum()}|n={len(arr)}")
    return hashlib.sha256("\n".join(parts).encode()).hexdigest()


@pytest.fixture
def baseline():
    if not BASELINE.exists():
        pytest.skip("Baseline fingerprint non disponibile (esegui tests/_capture_baseline.py)")
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def test_mc_fingerprint_matches_baseline(baseline):
    monthly_tr = load_monthly_tr(str(TR_FIXTURE))
    model_mix_global, characteristics_global = parse_tr_file(str(TR_FIXTURE))
    cfg = SimulationConfig(n_runs=20, random_seed=42, use_int16=True,
                           max_conflict_resolution_iterations=10, start_month_offset=0)
    results = run_monthly_simulations_optimized(
        forecast_values=list(FORECAST.values),
        month_names=list(FORECAST.index),
        n_months=len(FORECAST),
        model_mix=model_mix_global,
        characteristics=characteristics_global,
        exclusions={},
        config=cfg,
        monthly_tr=monthly_tr,
    )
    assert _hash_mc_results(results) == baseline["mc_fingerprint"]
```

- [ ] **Step 2: Esegui test**

```bash
python -m pytest tests/test_regression_mc.py -v
```
Expected: PASSED.

- [ ] **Step 3: Commit**

```bash
git add tests/test_regression_mc.py
git commit -m "test: regressione MC fingerprint via mid_term.montecarlo"
```

---

### Task 4.3: Esporta API montecarlo in __init__

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Aggiungi export**

```python
from mid_term.montecarlo import (
    run_monthly_simulations_optimized,
    run_single_month_simulation_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
)
```
Nota: NON esportare `SimulationConfig` da `montecarlo.py` perché la versione canonica è in `mid_term.config`. Se le due sono divergenti, lo step di compatibilità le riconcilia.

- [ ] **Step 2: Verifica SimulationConfig unica**

```bash
python -c "
from mid_term import SimulationConfig as SC1
from mid_term.montecarlo import SimulationConfig as SC2
print('config==montecarlo:', SC1 is SC2)
"
```
Se `False`: pianificare aliasing in Task 7 (lo shim `mc_simulator_V2.py` deve esportare entrambe come compatibili). Annota.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): esporta API modulo montecarlo"
```

---

## Fase 5 — Modulo `matching.py` (step 3-4)

### Task 5.1: Copia sku_matching_SS in mid_term/matching.py

**Files:**
- Create: `mid_term/matching.py`

- [ ] **Step 1: Copia file**

```bash
cp sku_matching_SS.py mid_term/matching.py
```

- [ ] **Step 2: Rimuovi main_sku_v2 e blocco `if __name__`**

In `mid_term/matching.py` rimuovi `main_sku_v2` (linee ~1403+) e qualsiasi blocco `if __name__ == "__main__"`.

- [ ] **Step 3: Aggiungi logger + sostituisci print**

```python
from mid_term._logging import get_logger
from mid_term.exceptions import MatchingError
log = get_logger("matching")
```
Sostituisci print → logger.

- [ ] **Step 4: Aggiorna import a mid_term**

`from tr_parser_V2 import canonical_model_name` → `from mid_term.tr import canonical_model_name`.

- [ ] **Step 5: Verifica import**

```bash
python -c "from mid_term.matching import match_skus_ultra_optimized, NormalizationCache, normalize_value_cached; print('OK')"
```
Expected: `OK`.

- [ ] **Step 6: Commit**

```bash
git add mid_term/matching.py
git commit -m "feat(mid_term): matching.py — SKU matching + safety stock"
```

---

### Task 5.2: Esporta API matching in __init__

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Aggiungi export**

```python
from mid_term.matching import (
    match_skus_ultra_optimized,
    NormalizationCache,
    normalize_value_cached,
)
```

- [ ] **Step 2: Verifica**

```bash
python -c "from mid_term import match_skus_ultra_optimized, NormalizationCache, normalize_value_cached; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): esporta API modulo matching"
```

---

## Fase 6 — Orchestratore `pipeline.py`

### Task 6.1: Scheletro SafetyStockPipeline

**Files:**
- Create: `mid_term/pipeline.py`
- Create: `tests/test_pipeline_smoke.py`

- [ ] **Step 1: Scrivi scheletro pipeline.py**

```python
# mid_term/pipeline.py
"""Orchestratore end-to-end della pipeline SS Classica (step 0..4)."""
import time
from typing import Optional

import pandas as pd

from mid_term._logging import get_logger
from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.results import (TRData, BOMData, SimulationResult,
                               MatchResult, SafetyStockResult)
from mid_term.tr import load_monthly_tr, parse_tr_file
from mid_term.montecarlo import (run_monthly_simulations_optimized,
                                  build_wide_dataframe,
                                  apply_pack_colors_to_wide_df)
from mid_term.matching import match_skus_ultra_optimized

log = get_logger("pipeline")


class SafetyStockPipeline:
    """Orchestratore end-to-end della pipeline SS Classica.

    Espone sia run() (esecuzione completa) sia run_step_N() per esecuzione
    granulare degli step (utile per debugging, test parziali e integrazione
    con orchestratori esterni tipo Airflow).
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

    def run_step_0_load_tr(self) -> TRData:
        raise NotImplementedError  # implementato in Task 6.2

    def run_step_1_build_bom(self) -> BOMData:
        raise NotImplementedError  # implementato in Task 6.3

    def run_step_2_montecarlo(self, tr: TRData) -> SimulationResult:
        raise NotImplementedError  # implementato in Task 6.4

    def run_step_3_matching(self, sim: SimulationResult, bom: BOMData) -> MatchResult:
        raise NotImplementedError  # implementato in Task 6.5

    def run_step_4_safety_stock(self, mat: MatchResult, bom: BOMData) -> SafetyStockResult:
        raise NotImplementedError  # implementato in Task 6.6
```

- [ ] **Step 2: Scrivi smoke test pipeline**

```python
# tests/test_pipeline_smoke.py
from pathlib import Path
import pytest
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline


def test_pipeline_construction(tmp_path):
    """SafetyStockPipeline deve costruirsi con validate_config=False senza fallire."""
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    sim = SimulationConfig(n_runs=20)
    pipeline = SafetyStockPipeline(cfg, sim, validate_config=False)
    assert pipeline.config is cfg
    assert pipeline.sim_config is sim


def test_pipeline_validate_failure(tmp_path):
    """Senza input file, validate=True solleva ConfigurationError."""
    from mid_term.exceptions import ConfigurationError
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    with pytest.raises(ConfigurationError):
        SafetyStockPipeline(cfg, validate_config=True)
```

- [ ] **Step 3: Esegui smoke test**

```bash
python -m pytest tests/test_pipeline_smoke.py -v
```
Expected: 2 PASSED.

- [ ] **Step 4: Commit**

```bash
git add mid_term/pipeline.py tests/test_pipeline_smoke.py
git commit -m "feat(mid_term): scheletro SafetyStockPipeline + smoke test"
```

---

### Task 6.2: Implementa run_step_0_load_tr

**Files:**
- Modify: `mid_term/pipeline.py`

- [ ] **Step 1: Sostituisci NotImplementedError con implementazione**

In `mid_term/pipeline.py`, metodo `run_step_0_load_tr`:

```python
def run_step_0_load_tr(self) -> TRData:
    """Step 0: carica TR + characteristics + forecast + esclusioni.

    Output: TRData con tutto il necessario per Monte Carlo (step 2).
    """
    from mid_term.montecarlo import (
        load_monthly_forecast,
        calculate_n_months_needed,
        parse_exclusions,
    )

    t0 = time.time()
    tr_path        = str(self.config.resolve_input(self.config.tr_file))
    forecast_path  = str(self.config.resolve_input(self.config.forecast_file))
    exclusions_path = str(self.config.resolve_input(self.config.exclusions_file))

    log.info("Step 0: caricamento TR da %s", tr_path)
    monthly_tr = load_monthly_tr(tr_path)
    model_mix, characteristics = parse_tr_file(tr_path)

    log.info("Step 0: caricamento forecast da %s", forecast_path)
    forecast_values, month_names = load_monthly_forecast(
        forecast_path, start_offset=self.sim_config.start_month_offset
    )
    n_months_needed = calculate_n_months_needed()

    log.info("Step 0: caricamento esclusioni da %s", exclusions_path)
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
```

- [ ] **Step 2: Verifica run_step_0 con dati reali**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline
cfg = PipelineConfig(input_dir=Path('Input'), output_dir=Path('Output'))
sim = SimulationConfig(n_runs=20)
p = SafetyStockPipeline(cfg, sim, validate_config=False)
tr = p.run_step_0_load_tr()
print(f'models: {sorted(tr.model_mix.models)}')
print(f'characteristics: {len(tr.characteristics)}')
print(f'forecast months: {tr.forecast_months[:3]}')
print(f'elapsed: {tr.elapsed_seconds:.2f}s')
"
```
Expected: stampa lista modelli (5: MSV4, MSV4PP, MSV4RI, MSV4RS, MSV4S), numero caratteristiche, primi 3 mesi forecast.

- [ ] **Step 3: Commit**

```bash
git add mid_term/pipeline.py
git commit -m "feat(pipeline): run_step_0_load_tr"
```

---

### Task 6.3: Implementa run_step_1_build_bom

**Files:**
- Modify: `mid_term/pipeline.py`

- [ ] **Step 1: Studia process_all_bom_combinations.main() per replicarne la logica**

Leggi `process_all_bom_combinations.py` linee 607-700. Annota:
- Quali file legge da Input/
- Quali file scrive in Output/
- Quali funzioni di `bom.py` invoca

- [ ] **Step 2: Implementa run_step_1_build_bom**

```python
def run_step_1_build_bom(self) -> BOMData:
    """Step 1: enrichment BOM + matrice configurazioni + catalogo SKU.

    Itera sulle combinazioni BOM × Mappatura definite in config.bom_files,
    applica i quality check, deduplica le configurazioni e produce:
      - matrices: {supermodel: matrix_output DataFrame}
      - catalog:  tutte_righe_univoche_V2 DataFrame
    """
    from mid_term.bom import process_combination, run_quality_checks

    t0 = time.time()
    log.info("Step 1: build BOM per supermodel %s", list(self.config.bom_files.keys()))

    # ... logica replicata da process_all_bom_combinations.main(), adattata a usare:
    #     - self.config.resolve_input(...) per i path
    #     - self.config.bom_files per iterazione
    #     - self.config.filter_untraced per il flag
    #     - self.config.write_intermediate per decidere se salvare su disco
    # Output finale: BOMData(catalog=..., matrices=..., quality_report=...)
    #
    # NOTA: la logica esatta dipende dal contenuto di process_all_bom_combinations.main();
    # il task implementativo deve traccia per traccia portare quella logica qui,
    # sostituendo path hardcoded con self.config.

    raise NotImplementedError("TODO: replica process_all_bom_combinations.main() qui")
```

- [ ] **Step 3: Implementa completamente (non lasciare NotImplementedError)**

Sostituisci il `raise NotImplementedError` con la logica completa derivata da `process_all_bom_combinations.main()`. Mappa esattamente: ogni `pd.read_excel(".\\Input\\X.xlsx")` → `pd.read_excel(self.config.resolve_input("X.xlsx"))`. Ogni `to_excel(".\\Output\\Y.xlsx")` → condizionale su `self.config.write_intermediate` con `self.config.resolve_output("Y.xlsx")`.

- [ ] **Step 4: Verifica run_step_1**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline
cfg = PipelineConfig(input_dir=Path('Input'), output_dir=Path('Output'), write_intermediate=False)
sim = SimulationConfig(n_runs=20)
p = SafetyStockPipeline(cfg, sim, validate_config=False)
bom = p.run_step_1_build_bom()
print(f'matrices: {list(bom.matrices.keys())}')
print(f'catalog rows: {len(bom.catalog)}')
print(f'elapsed: {bom.elapsed_seconds:.2f}s')
"
```
Expected: stampa supermodel ['V21E', 'V22E', 'V24T'], numero righe catalogo, tempo.

- [ ] **Step 5: Verifica hash catalogo invariato**

```bash
python -X utf8 -c "
import hashlib
from pathlib import Path
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline
cfg = PipelineConfig(input_dir=Path('Input'), output_dir=Path('Output'), write_intermediate=False)
p = SafetyStockPipeline(cfg, validate_config=False)
bom = p.run_step_1_build_bom()
# Hash deterministico del catalogo
import pandas as pd
buf = bom.catalog.to_csv(index=False).encode('utf-8')
print('catalog hash:', hashlib.sha256(buf).hexdigest()[:16])
"
```
Annota l'hash. Confronta con l'output di `process_all_bom_combinations.main()` originale (se baseline disponibile).

- [ ] **Step 6: Commit**

```bash
git add mid_term/pipeline.py
git commit -m "feat(pipeline): run_step_1_build_bom"
```

---

### Task 6.4: Implementa run_step_2_montecarlo

**Files:**
- Modify: `mid_term/pipeline.py`

- [ ] **Step 1: Implementa metodo**

```python
def run_step_2_montecarlo(self, tr: TRData) -> SimulationResult:
    """Step 2: Monte Carlo + Dirichlet su forecast mensile.

    Usa TRData (output step 0) e SimulationConfig per parametri RNG.
    """
    from mid_term.montecarlo import SimulationConfig as MCConfig

    t0 = time.time()
    # Adatta SimulationConfig pubblico a quello interno di montecarlo se divergenti
    mc_cfg = MCConfig(
        n_runs=self.sim_config.n_runs,
        random_seed=self.sim_config.random_seed,
        start_month_offset=self.sim_config.start_month_offset,
        use_int16=self.sim_config.use_int16,
        max_conflict_resolution_iterations=self.sim_config.max_conflict_iterations,
    )

    log.info("Step 2: Monte Carlo n_runs=%d seed=%d", mc_cfg.n_runs, mc_cfg.random_seed)
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
    df_wide = build_wide_dataframe(results_by_month, tr.forecast_months, tr.n_months_needed)
    df_wide = apply_pack_colors_to_wide_df(df_wide, ...)  # parametri da capire da mc_simulator_V2

    elapsed = time.time() - t0
    log.info("Step 2 completato in %.1fs: %d configurazioni", elapsed, len(df_wide))

    return SimulationResult(
        df_wide=df_wide,
        n_runs=mc_cfg.n_runs,
        n_months=tr.n_months_needed,
        month_names=tr.forecast_months,
        elapsed_seconds=elapsed,
    )
```
**NOTA:** verificare firma esatta di `apply_pack_colors_to_wide_df` leggendo `mid_term/montecarlo.py`.

- [ ] **Step 2: Verifica run_step_2 + fingerprint MC identico al baseline**

```bash
python -X utf8 -c "
import json
from pathlib import Path
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline
import hashlib

cfg = PipelineConfig(input_dir=Path('Input'), output_dir=Path('Output'), write_intermediate=False)
sim = SimulationConfig(n_runs=20, random_seed=42)
p = SafetyStockPipeline(cfg, sim, validate_config=False)
tr = p.run_step_0_load_tr()
result = p.run_step_2_montecarlo(tr)
print(f'rows: {len(result.df_wide)}  cols: {len(result.df_wide.columns)}')
print(f'elapsed: {result.elapsed_seconds:.1f}s')
"
```
Expected: numero righe e tempo.

- [ ] **Step 3: Commit**

```bash
git add mid_term/pipeline.py
git commit -m "feat(pipeline): run_step_2_montecarlo"
```

---

### Task 6.5: Implementa run_step_3_matching e run_step_4_safety_stock

**Files:**
- Modify: `mid_term/pipeline.py`

- [ ] **Step 1: Implementa run_step_3 + run_step_4**

```python
def run_step_3_matching(self, sim: SimulationResult, bom: BOMData) -> MatchResult:
    """Step 3: matching SKU contro configurazioni simulate."""
    t0 = time.time()
    log.info("Step 3: matching SKU su %d configurazioni", len(sim.df_wide))

    # ... logica derivata da match_skus_ultra_optimized() / main_sku_v2()
    # Usa bom.catalog come catalogo SKU, sim.df_wide come simulazione.
    # Restituisce MatchResult con df_matched + column_mapping + solo_modello.

    raise NotImplementedError("TODO: derivare da main_sku_v2()")

def run_step_4_safety_stock(self, mat: MatchResult, bom: BOMData) -> SafetyStockResult:
    """Step 4: calcolo percentili + scrittura output finale."""
    t0 = time.time()
    log.info("Step 4: safety stock al percentile %d", self.sim_config.percentile_safety_stock)

    # ... logica derivata dalla parte finale di main_sku_v2() che calcola
    # i percentili e scrive sku_safety_stock_PER_MESE_V2.xlsx +
    # sku_safety_stock_leadtime_PER_SKU_V2.xlsx.
    # Rispetta config.write_final_output: se False, non scrive su disco.

    raise NotImplementedError("TODO: derivare da main_sku_v2()")
```

- [ ] **Step 2: Implementa completamente (no NotImplementedError)**

Sostituisci entrambi i `raise NotImplementedError` con la logica derivata da `main_sku_v2()` (linee ~1403-fine di `sku_matching_SS.py`).

- [ ] **Step 3: Test end-to-end completo via run()**

```bash
python -X utf8 -c "
from pathlib import Path
from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline
cfg = PipelineConfig(input_dir=Path('Input'), output_dir=Path('Output'),
                     write_intermediate=False, write_final_output=False)
sim = SimulationConfig(n_runs=20, random_seed=42, percentile_safety_stock=99)
p = SafetyStockPipeline(cfg, sim, validate_config=False)
result = p.run()
print(f'per_sku rows: {len(result.per_sku)}')
print(f'per_mese rows: {len(result.per_mese)}')
"
```
Expected: stampa numero righe.

- [ ] **Step 4: Commit**

```bash
git add mid_term/pipeline.py
git commit -m "feat(pipeline): run_step_3_matching + run_step_4_safety_stock + run() end-to-end"
```

---

### Task 6.6: Esporta SafetyStockPipeline e dataclass risultato

**Files:**
- Modify: `mid_term/__init__.py`

- [ ] **Step 1: Aggiungi export pipeline + results**

```python
from mid_term.pipeline import SafetyStockPipeline
from mid_term.results  import (TRData, BOMData, SimulationResult,
                                MatchResult, SafetyStockResult)
```
Aggiungi a `__all__`.

- [ ] **Step 2: Verifica API pubblica completa**

```bash
python -c "
from mid_term import (SafetyStockPipeline, PipelineConfig, SimulationConfig,
                     TRData, BOMData, SimulationResult, MatchResult, SafetyStockResult,
                     canonical_model_name, MidTermError)
print('API pubblica OK')
"
```
Expected: stampa `API pubblica OK`.

- [ ] **Step 3: Commit**

```bash
git add mid_term/__init__.py
git commit -m "feat(mid_term): API pubblica completa SafetyStockPipeline + results"
```

---

## Fase 7 — Shim retro-compatibili nella root

### Task 7.1: Shim tr_parser_V2.py

**Files:**
- Modify: `tr_parser_V2.py` (sovrascrittura completa con shim)

- [ ] **Step 1: Sovrascrivi tr_parser_V2.py come shim**

```python
# tr_parser_V2.py
"""Shim retrocompatibile: riesporta API da mid_term.tr.

Mantenuto per non rompere consumer esistenti (es. app_V3.py).
Nuovi consumer: usare direttamente `from mid_term import ...`.
"""
from mid_term.tr import *  # noqa: F401,F403
from mid_term.tr import (
    parse_tr_file,
    load_monthly_tr,
    get_tr_for_month,
    parse_model_mix,
    ModelMix,
    CharacteristicGroup,
    canonical_model_name,
)
```

- [ ] **Step 2: Verifica consumer interno**

```bash
python -c "from tr_parser_V2 import parse_tr_file, load_monthly_tr, canonical_model_name, ModelMix; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add tr_parser_V2.py
git commit -m "refactor: tr_parser_V2.py diventa shim su mid_term.tr"
```

---

### Task 7.2: Shim mc_simulator_V2.py

**Files:**
- Modify: `mc_simulator_V2.py`

- [ ] **Step 1: Sovrascrivi come shim**

```python
# mc_simulator_V2.py
"""Shim retrocompatibile: riesporta API da mid_term.montecarlo."""
from mid_term.montecarlo import *  # noqa: F401,F403
from mid_term.montecarlo import (
    SimulationConfig,
    run_monthly_simulations_optimized,
    run_single_month_simulation_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
)
```

- [ ] **Step 2: Verifica**

```bash
python -c "from mc_simulator_V2 import SimulationConfig, run_monthly_simulations_optimized, build_wide_dataframe; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add mc_simulator_V2.py
git commit -m "refactor: mc_simulator_V2.py diventa shim su mid_term.montecarlo"
```

---

### Task 7.3: Shim sku_matching_SS.py

**Files:**
- Modify: `sku_matching_SS.py`

- [ ] **Step 1: Sovrascrivi come shim**

```python
# sku_matching_SS.py
"""Shim retrocompatibile: riesporta API da mid_term.matching."""
from mid_term.matching import *  # noqa: F401,F403
from mid_term.matching import (
    match_skus_ultra_optimized,
    NormalizationCache,
    normalize_value_cached,
)
```

- [ ] **Step 2: Verifica**

```bash
python -c "from sku_matching_SS import match_skus_ultra_optimized, NormalizationCache; print('OK')"
```
Expected: `OK`.

- [ ] **Step 3: Commit**

```bash
git add sku_matching_SS.py
git commit -m "refactor: sku_matching_SS.py diventa shim su mid_term.matching"
```

---

### Task 7.4: Shim gestione_dipendenze_V2 + skubundle + process_all_bom

**Files:**
- Modify: `gestione_dipendenze_V2.py`
- Modify: `skubundle.py`
- Modify: `skubundle_OPTIMIZED_V2.py`
- Modify: `process_all_bom_combinations.py`

- [ ] **Step 1: Sovrascrivi gestione_dipendenze_V2.py**

```python
# gestione_dipendenze_V2.py
"""Shim retrocompatibile: riesporta API da mid_term.bom."""
from mid_term.bom import *  # noqa: F401,F403
from mid_term.bom import process_excel_to_matrix as process_excel_to_matrix_v2
from mid_term.bom import process_excel_to_matrix
```

- [ ] **Step 2: Sovrascrivi skubundle.py e skubundle_OPTIMIZED_V2.py**

```python
# skubundle.py + skubundle_OPTIMIZED_V2.py
"""Shim retrocompatibile: riesporta API dedup da mid_term.bom."""
from mid_term.bom import *  # noqa: F401,F403
```

- [ ] **Step 3: Sovrascrivi process_all_bom_combinations.py**

```python
# process_all_bom_combinations.py
"""Shim retrocompatibile: riesporta API da mid_term.bom."""
from mid_term.bom import *  # noqa: F401,F403
from mid_term.bom import process_combination, run_quality_checks
```

- [ ] **Step 4: Verifica tutti gli shim**

```bash
python -c "
from gestione_dipendenze_V2 import process_excel_to_matrix_v2
from skubundle import *
from skubundle_OPTIMIZED_V2 import *
from process_all_bom_combinations import process_combination, run_quality_checks
print('shim BOM OK')
"
```
Expected: `shim BOM OK`.

- [ ] **Step 5: Commit**

```bash
git add gestione_dipendenze_V2.py skubundle.py skubundle_OPTIMIZED_V2.py process_all_bom_combinations.py
git commit -m "refactor: shim BOM verso mid_term.bom"
```

---

### Task 7.5: Verifica app_V3.py parte ancora

**Files:**
- Read: `app_V3.py` (verifica solo, no modifica)

- [ ] **Step 1: Avvia app_V3 in dry-run (compile check)**

```bash
python -X utf8 -c "
import ast
with open('app_V3.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('syntax OK')
"
```
Expected: `syntax OK`.

- [ ] **Step 2: Verifica import resolution**

```bash
python -X utf8 -c "
import importlib.util
spec = importlib.util.spec_from_file_location('app_V3', 'app_V3.py')
mod = importlib.util.module_from_spec(spec)
# Non eseguiamo, ma carichiamo per check import
spec.loader.exec_module(mod)
" 2>&1 | head -20
```
Se compaiono errori ImportError, è necessario aggiungere alias nel relativo shim. Annota e correggi.

- [ ] **Step 3: Smoke launch (5 secondi)**

```bash
timeout 5 streamlit run app_V3.py --server.headless true 2>&1 | tail -10
```
Expected: `You can now view your Streamlit app` o equivalente. Eventuali ImportError vanno corretti negli shim.

- [ ] **Step 4: Commit (se modifiche correttive ai shim necessarie)**

```bash
git add tr_parser_V2.py mc_simulator_V2.py sku_matching_SS.py gestione_dipendenze_V2.py
git commit -m "fix: shim retro-compat per import app_V3"
```

---

## Fase 8 — Cleanup finale

### Task 8.1: Rimuovi riferimenti "Claude"

**Files:** (potenziale modifica) tutti i file del package + root

- [ ] **Step 1: Grep occorrenze Claude**

```bash
grep -rni "claude" mid_term/ tests/ *.py 2>/dev/null | grep -v "\.pyc"
```

- [ ] **Step 2: Per ogni occorrenza, rimuovi/sostituisci**

Apri il file, valuta contesto:
- Commento con menzione "creato da Claude" / "by Claude" / "Claude Sonnet" → rimuovere intera riga.
- Nome variabile o stringa: sostituire con termine neutro.

- [ ] **Step 3: Verifica zero occorrenze residue**

```bash
grep -rni "claude" mid_term/ tests/ *.py 2>/dev/null | grep -v "\.pyc" | wc -l
```
Expected: `0`.

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "chore: rimozione riferimenti Claude da codice e commenti"
```

---

### Task 8.2: mypy check

**Files:** (potenziali fix) `mid_term/*.py`

- [ ] **Step 1: Installa mypy se mancante**

```bash
pip show mypy >/dev/null 2>&1 || pip install mypy
```

- [ ] **Step 2: Esegui mypy sul package**

```bash
python -m mypy mid_term/ --ignore-missing-imports 2>&1 | tail -30
```

- [ ] **Step 3: Per ogni errore critico, applica fix mirato**

Mypy errors come `Incompatible return value type` o `Argument has incompatible type` vanno corretti aggiungendo type hint o `# type: ignore[<code>]` mirati. Errori cosmetici (e.g., `[no-any-return]` su pandas) possono essere silenziati per ora.

- [ ] **Step 4: Commit dei fix**

```bash
git add mid_term/
git commit -m "chore: fix type hint per mypy clean"
```

---

### Task 8.3: Test di regressione end-to-end finale

**Files:**
- Create: `tests/test_regression_e2e.py`

- [ ] **Step 1: Scrivi test e2e**

```python
# tests/test_regression_e2e.py
"""Test end-to-end: confronta output finale con baseline pre-refactoring."""
import json
from pathlib import Path

import pytest

from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline


BASELINE = Path("tests/_baseline_fingerprints.json")
TR_FIXTURE = Path("tests/fixtures/TR_test.xlsx")


@pytest.fixture
def baseline():
    if not BASELINE.exists():
        pytest.skip("Baseline non disponibile")
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def test_step_0_tr_models_invariant(baseline, tmp_path):
    cfg = PipelineConfig(input_dir=TR_FIXTURE.parent, output_dir=tmp_path)
    sim = SimulationConfig(n_runs=20, random_seed=42)
    p = SafetyStockPipeline(cfg, sim, validate_config=False)
    tr = p.run_step_0_load_tr()
    assert sorted(tr.model_mix.models) == sorted(baseline["tr_models"])
    assert len(tr.characteristics) == baseline["n_characteristics"]


def test_step_2_mc_fingerprint_invariant(baseline, tmp_path):
    cfg = PipelineConfig(input_dir=TR_FIXTURE.parent, output_dir=tmp_path,
                         write_intermediate=False, write_final_output=False)
    sim = SimulationConfig(n_runs=20, random_seed=42)
    p = SafetyStockPipeline(cfg, sim, validate_config=False)
    tr  = p.run_step_0_load_tr()
    res = p.run_step_2_montecarlo(tr)
    # df_wide deve riprodurre i conteggi totali baseline
    assert int(res.df_wide.select_dtypes("number").sum().sum()) == baseline["mc_total_bikes"]


def test_canonical_idempotent_full_suite():
    from mid_term import canonical_model_name
    cases = [
        ("MSV4", "MSV4"),
        ("MSV4S​", "MSV4S"),
        ("MSV4PP​", "MSV4PP"),
        ("MSV4RI", "MSV4RI"),
        ("MSV4RS", "MSV4RS"),
        ("PANV47G", "PANV47G"),
        ("PAN7GSP", "PAN7GSP"),
        ("DVLV4S", "DVLV4S"),
        ("msv4", "MSV4"),
        ("  msv4s  ", "MSV4S"),
        ("MSV4 ", "MSV4"),
        ("MSV4﻿", "MSV4"),
        ("MSV4‌", "MSV4"),
        (None, ""),
        (float("nan"), ""),
        ("", ""),
        ("ALTRO", "ALTRO"),
        ("ALTRO​", "ALTRO"),
    ]
    for raw, expected in cases:
        assert canonical_model_name(raw) == expected, f"{raw!r} -> got {canonical_model_name(raw)!r}, expected {expected!r}"
```

- [ ] **Step 2: Esegui suite test completa**

```bash
python -m pytest tests/ -v
```
Expected: tutti i test PASSED.

- [ ] **Step 3: Commit**

```bash
git add tests/test_regression_e2e.py
git commit -m "test: end-to-end regression contro baseline pre-refactoring"
```

---

### Task 8.4: Aggiorna memory/MEMORY.md

**Files:**
- Modify: `memory/MEMORY.md`

- [ ] **Step 1: Aggiungi sezione "Refactoring 2026-05-27"**

Apri `memory/MEMORY.md` e aggiungi in cima dopo l'intestazione:

```markdown
## Refactoring mid_term (2026-05-27)

Pipeline SS Classica rifattorizzata in package `mid_term/`:

- API pubblica: `SafetyStockPipeline`, `PipelineConfig`, `SimulationConfig`
- 10 file Python nel package: `config.py`, `exceptions.py`, `_logging.py`,
  `results.py`, `tr.py`, `bom.py`, `montecarlo.py`, `matching.py`,
  `pipeline.py`, `__init__.py`.
- 6 shim retrocompatibili nella root: `tr_parser_V2.py`, `mc_simulator_V2.py`,
  `sku_matching_SS.py`, `gestione_dipendenze_V2.py`, `skubundle.py`,
  `skubundle_OPTIMIZED_V2.py`, `process_all_bom_combinations.py`.
- Logging strutturato (`logging.getLogger`), niente più `print()`.
- Gerarchia eccezioni dedicata in `mid_term.exceptions`.
- Test di regressione in `tests/test_regression_*.py` con baseline
  `tests/_baseline_fingerprints.json`.

Spec: `docs/superpowers/specs/2026-05-27-mid-term-ss-refactor-design.md`
Plan: `docs/superpowers/plans/2026-05-27-mid-term-ss-refactor.md`
```

- [ ] **Step 2: Commit**

```bash
git add memory/MEMORY.md
git commit -m "docs: memory aggiornata con refactoring mid_term"
```

---

### Task 8.5: Commit finale di riepilogo

- [ ] **Step 1: Verifica git status pulito**

```bash
git status
```
Expected: `nothing to commit, working tree clean`.

- [ ] **Step 2: Mostra log refactoring completo**

```bash
git log --oneline --since="2026-05-27" | head -50
```
Verifica che siano presenti commit per ogni task.

- [ ] **Step 3: Riepilogo finale**

```bash
echo "=== STRUTTURA FINALE ==="
ls mid_term/
echo "=== SHIM ROOT ==="
ls tr_parser_V2.py mc_simulator_V2.py sku_matching_SS.py gestione_dipendenze_V2.py skubundle.py skubundle_OPTIMIZED_V2.py process_all_bom_combinations.py
echo "=== TEST ==="
ls tests/
echo "=== REGRESSION RESULT ==="
python -m pytest tests/ -v --tb=short 2>&1 | tail -5
```

---

## Riepilogo task

| Fase | Task | Descrizione |
|------|------|-------------|
| 0 | 0.1 | Setup directory mid_term/, tests/, fixtures |
| 0 | 0.2 | Cattura baseline fingerprint pipeline corrente |
| 1 | 1.1 | exceptions.py |
| 1 | 1.2 | _logging.py |
| 1 | 1.3 | results.py |
| 1 | 1.4 | config.py + test_config.py |
| 1 | 1.5 | __init__.py base |
| 2 | 2.1 | tr.py — copia + pulizia tr_parser_V2 |
| 2 | 2.2 | Esporta API tr in __init__ |
| 3 | 3.1 | bom.py — copia gestione_dipendenze_V2 |
| 3 | 3.2 | bom.py + skubundle dedup |
| 3 | 3.3 | bom.py + process_combination |
| 3 | 3.4 | Esporta API bom |
| 4 | 4.1 | montecarlo.py — copia mc_simulator_V2 senza rolling |
| 4 | 4.2 | Test regression MC fingerprint |
| 4 | 4.3 | Esporta API montecarlo |
| 5 | 5.1 | matching.py — copia sku_matching_SS |
| 5 | 5.2 | Esporta API matching |
| 6 | 6.1 | Scheletro SafetyStockPipeline |
| 6 | 6.2 | run_step_0_load_tr |
| 6 | 6.3 | run_step_1_build_bom |
| 6 | 6.4 | run_step_2_montecarlo |
| 6 | 6.5 | run_step_3 + run_step_4 |
| 6 | 6.6 | Esporta SafetyStockPipeline |
| 7 | 7.1 | Shim tr_parser_V2 |
| 7 | 7.2 | Shim mc_simulator_V2 |
| 7 | 7.3 | Shim sku_matching_SS |
| 7 | 7.4 | Shim BOM (4 file) |
| 7 | 7.5 | Verifica app_V3 |
| 8 | 8.1 | Rimuovi riferimenti Claude |
| 8 | 8.2 | mypy clean |
| 8 | 8.3 | Test e2e regression |
| 8 | 8.4 | Aggiorna MEMORY.md |
| 8 | 8.5 | Commit finale + riepilogo |

**Totale: 31 task in 9 fasi.**

Ogni task ha verifica esplicita e produce un commit dedicato. Il test di regressione end-to-end (task 8.3) è il gate finale: se non passa, il refactoring non è considerato completo.

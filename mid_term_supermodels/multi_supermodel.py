# -*- coding: utf-8 -*-
"""Orchestratore Safety Stock multi-supermodel.

Esegue la pipeline SS per ogni cartella supermodel (domanda + TR propri,
seed diverso) e fa il pooling dei vettori per-run a livello componente.
"""
import os
import shutil
import time
from collections import namedtuple
from dataclasses import replace
from pathlib import Path
from typing import List

from mid_term_supermodels._logging import get_logger

log = get_logger(__name__)

# File condivisi (non per-supermodel) copiati dentro lo staged Input/.
_SHARED_FILES = (
    "States.xlsx",
    "esclusioni.xlsx",
    "Cambio Valuta.xlsx",
    "affidabilita_config.json",
)
_SHARED_DIRS = ("Dati Logistica",)

# MultiSupermodelResult e' introdotto nel Task 7 (results.py). Finche' non
# esiste, usiamo un namedtuple locale con gli stessi campi: l'orchestratore
# resta utilizzabile e il passaggio al dataclass vero sara' trasparente.
_FALLBACK_RESULT_FIELDS = (
    "per_sku_pooled", "per_supermodel_breakdown", "output_paths", "elapsed_seconds",
)
_MultiSupermodelResultFallback = namedtuple(
    "MultiSupermodelResult", _FALLBACK_RESULT_FIELDS
)


def discover_supermodels(input_dir: str) -> List[str]:
    """Trova le cartelle ROOT supermodel sotto input_dir (layout A-staging, flat).

    Una sottocartella `sub` di input_dir e' un supermodel se contiene
    DIRETTAMENTE (non sotto `sub/Input`):
      - una sottocartella MODEL/,
      - un file Total_demand.xlsx,
      - almeno un file il cui nome inizia con 'TR' (.xlsx/.xls).

    Restituisce il path ROOT `sub` (la cartella per-supermodel da cui lo staging
    copiera' tutti i dati). Questo esclude naturalmente le cartelle condivise
    ('Dati Logistica') e le cartelle sorgente grezze ('DVL', 'PANI'): non hanno
    una sottocartella MODEL/ ne' Total_demand/TR diretti.
    """
    base = Path(input_dir)
    result = []
    for sub in sorted(base.iterdir()):
        if not sub.is_dir():
            continue
        has_model = (sub / 'MODEL').is_dir()
        has_demand = (sub / 'Total_demand.xlsx').exists()
        has_tr = any(f.name.upper().startswith('TR') and f.suffix.lower() in ('.xlsx', '.xls')
                     for f in sub.iterdir() if f.is_file())
        if has_model and has_demand and has_tr:
            result.append(str(sub))
    log.info("Supermodel trovati: %d", len(result))
    return result


def stage_supermodel_input(sm_dir, shared_dir, work_root):
    """Crea <work_root>/<SM>/Input/ con TUTTO il per-supermodel (copiato da sm_dir)
    + i file condivisi (copiati da shared_dir). Ritorna il path <work_root>/<SM>
    (la dir in cui fare chdir; il package vedra' ./Input/...).

    Copia (shutil): l'intero contenuto di sm_dir in <work>/Input/, poi i condivisi
    noti da shared_dir: 'States.xlsx', 'esclusioni.xlsx', 'Cambio Valuta.xlsx',
    'affidabilita_config.json' (file) e 'Dati Logistica' (dir).
    """
    sm_dir = Path(sm_dir)
    shared_dir = Path(shared_dir)
    work_sm = Path(work_root) / sm_dir.name
    staged_input = work_sm / 'Input'

    # Clean slate: rimuovi un eventuale staging precedente (sempre sotto work_root,
    # MAI il vero Input/ dei dati).
    if work_sm.exists():
        shutil.rmtree(work_sm)

    # Copia l'intero contenuto per-supermodel: MODEL/, Total_demand, TR,
    # Mappatura_Unificata, matrix_output_*, tutte_righe_* (tutto cio' che esiste).
    shutil.copytree(sm_dir, staged_input)

    # Copia i file condivisi dentro lo stesso Input/ staged.
    for fname in _SHARED_FILES:
        src = shared_dir / fname
        if src.exists():
            shutil.copy2(src, staged_input / fname)
        else:
            log.warning("File condiviso assente, skip: %s", src)

    # Copia le dir condivise (es. 'Dati Logistica').
    for dname in _SHARED_DIRS:
        src = shared_dir / dname
        if src.is_dir():
            shutil.copytree(src, staged_input / dname)
        else:
            log.warning("Cartella condivisa assente, skip: %s", src)

    return work_sm


def _make_result(per_sku_pooled, per_supermodel_breakdown, output_paths,
                 elapsed_seconds):
    """Costruisce il risultato finale.

    Preferisce results.MultiSupermodelResult (Task 7) se disponibile; altrimenti
    degrada a un namedtuple locale con gli stessi campi.
    """
    try:
        from mid_term_supermodels.results import MultiSupermodelResult  # package copiato (Task 7)
        return MultiSupermodelResult(
            per_sku_pooled=per_sku_pooled,
            per_supermodel_breakdown=per_supermodel_breakdown,
            output_paths=output_paths,
            elapsed_seconds=elapsed_seconds,
        )
    except (ImportError, AttributeError):
        log.warning("results.MultiSupermodelResult non disponibile (Task 7 non "
                    "ancora applicato): uso fallback namedtuple locale.")
        return _MultiSupermodelResultFallback(
            per_sku_pooled=per_sku_pooled,
            per_supermodel_breakdown=per_supermodel_breakdown,
            output_paths=output_paths,
            elapsed_seconds=elapsed_seconds,
        )


def run_multi_supermodel(input_dir: str, output_dir: str, json_path: str,
                         seed_base: int = 42):
    """Esegue la pipeline SS per ogni supermodel e ne fa il pooling.

    Per ogni ROOT supermodel scoperta sotto input_dir esegue la pipeline
    SafetyStockPipeline.run_collect_run_vectors() con seed = seed_base + k
    (indipendenza statistica, decisione D8), raccoglie i vettori di domanda
    per-run per componente, poi chiama matching.pool_safety_stock().

    LAYOUT (A-staging, flat):
        <input_dir>/Dati Logistica/                 <- CONDIVISA (fissa)
        <input_dir>/States.xlsx, esclusioni.xlsx,
                    Cambio Valuta.xlsx, affidabilita_config.json   <- CONDIVISI
        <input_dir>/<SuperModel>/{MODEL/, Total_demand.xlsx, TR*.xlsx, derivati}

    PATH/CWD (punto critico): il package usa path RELATIVI ALLA CWD, hardcoded
    (es. matching.calcola_residual_lead_time(input_dir='.\\Input'),
    matching._load_cambio_eur('Input/Cambio Valuta.xlsx'),
    montecarlo.load_monthly_forecast('.\\Input\\...'),
    bom.OUTPUT_DIR/INPUT_DIR = Path('.')/'Input'). bom.py inoltre SCRIVE i file
    derivati (matrix_output_*, tutte_righe_*, Mappatura_Unificata) dentro
    ./Input. Per questo, invece di entrare nella cartella supermodel, si fa lo
    STAGING: per ogni supermodel si assembla <work>/<SM>/Input/ COPIANDO tutto il
    per-supermodel (da sm_dir) + i condivisi (da input_dir), si fa
    os.chdir(<work>/<SM>) e si esegue. Cosi' './Input/...' punta a una copia
    isolata, riscrivibile, che a fine run viene rimossa (rmtree su copie, mai sui
    dati reali). La 'Dati Logistica' viene copiata in ./Input/Dati Logistica, per
    cui config.shared_logistica_dir() con env NON impostata risolve correttamente.

    Returns:
        MultiSupermodelResult (o fallback namedtuple) con campi
        per_sku_pooled, per_supermodel_breakdown, output_paths, elapsed_seconds.
    """
    from mid_term_supermodels import matching  # package copiato: pool_safety_stock + preserve_run_vectors
    from mid_term_supermodels.config import PipelineConfig, load_simulation_config
    from mid_term_supermodels.montecarlo import load_monthly_forecast
    from mid_term_supermodels.pipeline import SafetyStockPipeline

    t0 = time.time()

    supermodel_dirs = discover_supermodels(input_dir)
    if not supermodel_dirs:
        raise RuntimeError(f"Nessun supermodel trovato in {input_dir}")

    # I condivisi (Dati Logistica + i 4 file) stanno direttamente in input_dir.
    shared_dir = Path(input_dir)
    # Area temporanea per lo staging, ACCANTO a Input/ (mai dentro), cleanabile.
    work_root = Path(input_dir).parent / '_work_supermodels'

    sm_results = {}
    percentile = None
    cwd0 = os.getcwd()
    try:
        for k, sm_dir in enumerate(supermodel_dirs):
            sm_name = Path(sm_dir).name
            seed = seed_base + k
            log.info("=== Supermodel %d/%d: %s (seed=%d) ===",
                     k + 1, len(supermodel_dirs), sm_name, seed)

            # Staging: copia tutto il per-supermodel + i condivisi in <work>/<SM>/Input.
            work_sm = stage_supermodel_input(sm_dir, shared_dir, work_root)

            try:
                # Il package risolve i path hardcoded rispetto alla CWD: entra
                # nella dir staged cosi' './Input/...' punta alla copia isolata.
                os.chdir(work_sm)

                # input_dir relativo ('Input') coerente con la CWD: le funzioni
                # che rispettano config.input_dir e quelle che leggono da './Input'
                # puntano allo stesso posto (l'Input staged del supermodel corrente).
                config = PipelineConfig(
                    input_dir=Path('Input'),
                    output_dir=Path(output_dir) / sm_name,
                )

                # Nomi-mese dal forecast del supermodel per risolvere start_month.
                forecast_path = config.resolve_input(config.forecast_file)
                _, month_names = load_monthly_forecast(str(forecast_path))

                sim_config = load_simulation_config(json_path, month_names)
                # Indipendenza statistica tra supermodel (D8): seed diverso per ciascuno.
                sim_config = replace(sim_config, random_seed=seed_base + k)
                if percentile is None:
                    percentile = sim_config.percentile_safety_stock

                pipeline = SafetyStockPipeline(config, sim_config, validate_config=False)
                vectors, _df_no_lt = pipeline.run_collect_run_vectors()
                sm_results[sm_name] = vectors
            finally:
                os.chdir(cwd0)
                # Rimuove SOLO la copia staged (sotto _work_supermodels), mai i dati reali.
                shutil.rmtree(work_sm, ignore_errors=True)
    finally:
        os.chdir(cwd0)
        # Best-effort: rimuove l'intera area temporanea di staging.
        shutil.rmtree(work_root, ignore_errors=True)

    df_pool, breakdown = matching.pool_safety_stock(
        sm_results, percentile=percentile,
    )

    elapsed = time.time() - t0
    log.info("Pooling completato: %d componenti, %.1fs",
             len(df_pool) if df_pool is not None else 0, elapsed)

    # Salva il risultato pooled su Excel (Task 7)
    out_path = os.path.join(output_dir, 'sku_safety_stock_POOLED_supermodels.xlsx')
    os.makedirs(output_dir, exist_ok=True)
    matching.save_results_pooled(df_pool, out_path, percentile=percentile)

    return _make_result(
        per_sku_pooled=df_pool,
        per_supermodel_breakdown=breakdown,
        output_paths={'pooled': out_path},
        elapsed_seconds=elapsed,
    )

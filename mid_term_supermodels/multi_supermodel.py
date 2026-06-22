# -*- coding: utf-8 -*-
"""Orchestratore Safety Stock multi-supermodel.

Esegue la pipeline SS per ogni cartella supermodel (domanda + TR propri,
seed diverso) e fa il pooling dei vettori per-run a livello componente.
"""
import os
import time
from collections import namedtuple
from dataclasses import replace
from pathlib import Path
from typing import List

from mid_term_supermodels._logging import get_logger

log = get_logger(__name__)

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
    """Trova le cartelle supermodel sotto input_dir.

    Una cartella è un supermodel se contiene: sottocartella MODEL/,
    un Total_demand.xlsx e almeno un file il cui nome inizia con 'TR'.
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

    Per ogni cartella supermodel scoperta sotto input_dir esegue la pipeline
    SafetyStockPipeline.run_collect_run_vectors() con seed = seed_base + k
    (indipendenza statistica, decisione D8), raccoglie i vettori di domanda
    per-run per componente, poi chiama matching.pool_safety_stock().

    PATH/CWD (punto critico): il package usa path RELATIVI ALLA CWD, hardcoded
    (es. matching.calcola_residual_lead_time(input_dir='.\\Input'),
    matching._load_cambio_eur('Input/Cambio Valuta.xlsx'),
    montecarlo.load_monthly_forecast('.\\Input\\...'),
    bom.OUTPUT_DIR/INPUT_DIR = Path('.')/'Input', calculate_n_months_needed()
    che legge 'Input/residual_lead_time.xlsx'). Passare solo input_dir a
    PipelineConfig NON basta: molte funzioni ignorano config.input_dir e leggono
    da CWD/Input. Per questo l'orchestratore fa os.chdir(sm_dir) per ciascun
    supermodel (con ripristino in finally), replicando cio' che il launcher
    run_pipeline.py fa con os.chdir(PKG). Si assume che dentro sm_dir esista un
    layout 'Input/' compatibile. config.input_dir e' comunque impostato a
    sm_dir/'Input' per le funzioni che invece lo rispettano.

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

    sm_results = {}
    percentile = None
    cwd0 = os.getcwd()
    try:
        for k, sm_dir in enumerate(supermodel_dirs):
            sm_name = Path(sm_dir).name
            seed = seed_base + k
            log.info("=== Supermodel %d/%d: %s (seed=%d) ===",
                     k + 1, len(supermodel_dirs), sm_name, seed)

            # Il package risolve i path hardcoded rispetto alla CWD: entra
            # nella cartella del supermodel cosi' '.\Input\...' punta al suo Input.
            os.chdir(sm_dir)

            sm_input = Path(sm_dir) / "Input"
            config = PipelineConfig(
                input_dir=sm_input,
                output_dir=Path(output_dir) / sm_name,
            )

            # Nomi-mese dal forecast del supermodel per risolvere start_month.
            forecast_path = config.resolve_input(config.forecast_file)
            _, month_names = load_monthly_forecast(str(forecast_path))

            sim_config = load_simulation_config(json_path, month_names)
            # Indipendenza statistica tra supermodel (D8): seed diverso per ciascuno.
            sim_config = replace(sim_config, random_seed=seed)
            if percentile is None:
                percentile = sim_config.percentile_safety_stock

            pipeline = SafetyStockPipeline(config, sim_config, validate_config=False)
            vectors, _df_no_lt = pipeline.run_collect_run_vectors()
            sm_results[sm_name] = vectors
    finally:
        os.chdir(cwd0)

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

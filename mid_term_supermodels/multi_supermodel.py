# -*- coding: utf-8 -*-
"""Orchestratore Safety Stock multi-supermodel.

Esegue la pipeline SS per ogni cartella supermodel (domanda + TR propri,
seed diverso) e fa il pooling dei vettori per-run a livello componente.
"""
import dataclasses
import json
import os
import shutil
import stat
import tempfile
import time
from collections import namedtuple
from dataclasses import replace
from pathlib import Path
from typing import List

from mid_term_supermodels._logging import get_logger
# Importati a livello modulo cosi' sono monkeypatchabili nei test
# (multi_supermodel.load_monthly_forecast / .collect_run_vectors_rolling).
from mid_term_supermodels.montecarlo import load_monthly_forecast
from mid_term_supermodels.rolling import collect_run_vectors_rolling

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


def _robust_rmtree(path):
    """Rimuove una cartella in modo robusto su Windows (best-effort, non solleva).

    I file copiati da sorgenti OneDrive possono essere read-only: rmtree
    fallirebbe con PermissionError. L'handler forza il permesso di scrittura
    e ritenta. Compatibile con onexc (Py 3.12+) e onerror (precedenti).
    """
    if not os.path.exists(path):
        return

    def _onerr(func, p, _exc):
        try:
            os.chmod(p, stat.S_IWRITE)
            func(p)
        except Exception:
            pass

    try:
        shutil.rmtree(path, onexc=_onerr)
    except TypeError:
        shutil.rmtree(path, onerror=_onerr)
    except Exception:
        pass


def _new_work_root():
    """Crea una dir di staging unica in TEMP (fuori da OneDrive: niente lock/sync)."""
    return Path(tempfile.mkdtemp(prefix='midterm_sm_'))


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
    # MAI il vero Input/ dei dati). rmtree robusto per i read-only Windows/OneDrive.
    if work_sm.exists():
        _robust_rmtree(work_sm)

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
    # Area temporanea per lo staging in TEMP (fuori OneDrive: niente lock/sync che
    # romperebbero rmtree su Windows). Unica per run, cleanabile.
    work_root = _new_work_root()

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
                # Crea Output/<SM>/ prima del run: write_intermediate salva il
                # catalogo + matrici per supermodel in modo persistente.
                os.makedirs(config.output_dir, exist_ok=True)

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
                # Rimuove SOLO la copia staged (in TEMP), mai i dati reali.
                _robust_rmtree(work_sm)
    finally:
        os.chdir(cwd0)
        # Best-effort: rimuove l'intera area temporanea di staging.
        _robust_rmtree(work_root)

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


def _build_sim_config_for_supermodel(json_path, month_names, k, seed_base):
    """Costruisce la SimulationConfig per UN supermodel in modalita' rolling.

    Problema: load_simulation_config valida start_month (dal JSON) contro i
    nomi-mese del forecast. In rolling ogni supermodel ha un orizzonte (e quindi
    nomi-mese) diversi, per cui lo start_month del JSON puo' NON esistere in un
    certo supermodel -> ConfigurationError. Ma il rolling IGNORA start_month e
    itera da offset 0 (vedi docstring di collect_run_vectors_rolling), quindi
    l'unica cosa che serve davvero qui e' ottenere una SimulationConfig valida
    con n_runs/percentile/affidabilita/gate_max_mesi corretti e offset = 0.

    Strategia (senza MAI mutare il file JSON):
      1) prova load_simulation_config(json_path, month_names) normalmente;
      2) se solleva ConfigurationError perche' lo start_month non e' nei
         month_names del supermodel, rilegge lo start_month dal JSON grezzo e
         ritenta passando [start_month] + month_names cosi' la validazione passa
         (l'offset risultante e' irrilevante: viene azzerato sotto).
    In ogni caso forza start_month_offset=0 e random_seed = seed_base + k.
    """
    from mid_term_supermodels.config import load_simulation_config
    from mid_term_supermodels.exceptions import ConfigurationError

    try:
        sim_config = load_simulation_config(json_path, month_names)
    except ConfigurationError as e:
        # Fallback robusto: start_month del JSON assente in questo supermodel.
        try:
            raw = json.loads(Path(json_path).read_text(encoding="utf-8"))
            start_month = raw.get("start_month")
        except Exception:
            start_month = None
        if not start_month:
            raise
        log.warning(
            "load_simulation_config fallita per il supermodel #%d (%s): "
            "start_month '%s' non presente nei mesi del forecast. "
            "Rolling itera da offset 0: uso fallback con start_month_offset=0.",
            k, e, start_month,
        )
        # month_names augmentati: garantisce che start_month sia presente.
        sim_config = load_simulation_config(json_path, [start_month] + list(month_names))

    # Rolling: offset ignorato, si itera 0..N. Seed diverso per supermodel (D8).
    return dataclasses.replace(sim_config, start_month_offset=0,
                               random_seed=seed_base + k)


def run_multi_supermodel_rolling(input_dir: str, output_dir: str, json_path: str,
                                 seed_base: int = 42, only_supermodels=None):
    """Orchestratore multi-supermodel in modalita' ROLLING + pooling per mese.

    Per ogni supermodel esegue collect_run_vectors_rolling (rolling sui mesi di
    lancio che PRESERVA i vettori per-run per componente), poi fa il pooling
    ALLINEATO PER MESE DI LANCIO: per ciascun mese di lancio M (unione di tutti i
    mesi visti tra i supermodel) raccoglie i vettori dei supermodel che hanno M e
    chiama matching.pool_safety_stock(). Il risultato e' un'unica tabella con la
    colonna mese_lancio.

    Mirroring di run_multi_supermodel per staging/chdir/cleanup: stesso schema
    (stage -> chdir -> run -> finally chdir+rmtree della copia), ma il motore
    per-supermodel e' il rolling al posto del run singolo.

    Args:
        input_dir:  cartella con i supermodel + i file condivisi.
        output_dir: cartella di output (l'Excel pooled rolling vi viene scritto).
        json_path:  JSON di run (n_runs, percentile, affidabilita', gate, ...).
        seed_base:  seed base; il supermodel k usa seed_base + k.

    Returns:
        MultiSupermodelResult (o fallback namedtuple) con per_sku_pooled che
        include la colonna mese_lancio.
    """
    from mid_term_supermodels import matching
    from mid_term_supermodels.config import PipelineConfig

    import pandas as pd

    t0 = time.time()

    supermodel_dirs = discover_supermodels(input_dir)
    if not supermodel_dirs:
        raise RuntimeError(f"Nessun supermodel trovato in {input_dir}")
    # Filtro opzionale: usa solo i supermodel selezionati (per nome cartella).
    if only_supermodels:
        _wanted = {str(s).strip() for s in only_supermodels}
        supermodel_dirs = [d for d in supermodel_dirs if Path(d).name in _wanted]
        if not supermodel_dirs:
            raise RuntimeError(f"Nessuno dei supermodel selezionati trovato: {only_supermodels}")

    shared_dir = Path(input_dir)
    # Staging in TEMP (fuori OneDrive: evita lock/sync che romperebbero rmtree).
    work_root = _new_work_root()

    sm_rolling = {}      # { sm_name : { mese_lancio : {'vectors', 'troncata', ...} } }
    percentile = None
    cwd0 = os.getcwd()
    try:
        for k, sm_dir in enumerate(supermodel_dirs):
            sm_name = Path(sm_dir).name
            seed = seed_base + k
            log.info("=== [ROLLING] Supermodel %d/%d: %s (seed=%d) ===",
                     k + 1, len(supermodel_dirs), sm_name, seed)

            work_sm = stage_supermodel_input(sm_dir, shared_dir, work_root)
            try:
                os.chdir(work_sm)

                config_sm = PipelineConfig(
                    input_dir=Path('Input'),
                    output_dir=Path(output_dir) / sm_name,
                )
                # Crea Output/<SM>/ PRIMA del run cosi' write_intermediate salva
                # tutte_righe_univoche_V2.xlsx + matrix_output_* per supermodel in
                # modo persistente (altrimenti finirebbero solo nello staged ./Input
                # temporaneo, cancellato dal rmtree).
                os.makedirs(config_sm.output_dir, exist_ok=True)

                # Nomi-mese dal forecast per-supermodel (per costruire sim_config).
                forecast_path = config_sm.resolve_input(config_sm.forecast_file)
                _, month_names = load_monthly_forecast(str(forecast_path))

                sim_config = _build_sim_config_for_supermodel(
                    json_path, month_names, k, seed_base,
                )
                if percentile is None:
                    percentile = sim_config.percentile_safety_stock

                rolling_res = collect_run_vectors_rolling(config_sm, sim_config)
                sm_rolling[sm_name] = rolling_res
            finally:
                os.chdir(cwd0)
                _robust_rmtree(work_sm)
    finally:
        os.chdir(cwd0)
        _robust_rmtree(work_root)

    if percentile is None:
        percentile = 99

    # ---- Pooling allineato per mese di lancio ----
    df_pool_rolling, per_sm_breakdown = _pool_rolling_by_month(
        sm_rolling, percentile, matching, pd,
    )

    elapsed = time.time() - t0
    log.info("[ROLLING] Pooling per mese completato: %d righe, %.1fs",
             len(df_pool_rolling) if df_pool_rolling is not None else 0, elapsed)

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(
        output_dir, 'sku_safety_stock_POOLED_rolling_supermodels.xlsx')
    matching.save_results_pooled(df_pool_rolling, out_path, percentile=percentile)

    return _make_result(
        per_sku_pooled=df_pool_rolling,
        per_supermodel_breakdown=per_sm_breakdown,
        output_paths={'pooled_rolling': out_path},
        elapsed_seconds=elapsed,
    )


def _pool_rolling_by_month(sm_rolling, percentile, matching, pd):
    """Pooling cross-supermodel ALLINEATO per mese di lancio.

    Args:
        sm_rolling: { sm_name : { mese_lancio : {'vectors':..., 'troncata':bool, ...} } }
        percentile: percentile per la safety stock.
        matching:   modulo matching (pool_safety_stock + parse_month_to_sortable).
        pd:         modulo pandas.

    Returns:
        (df_pool_rolling, breakdown_by_month)
        df_pool_rolling: concat dei pooling per mese, con colonne mese_lancio
            (posizione 0) e finestra_troncata; vuoto se nessun risultato.
        breakdown_by_month: { mese_lancio : breakdown_dict } da pool_safety_stock.
    """
    supermodels = list(sm_rolling.keys())

    # Unione di tutti i mesi di lancio visti tra i supermodel, ordine cronologico.
    all_months = set()
    for sm in supermodels:
        all_months.update(sm_rolling[sm].keys())
    months_sorted = sorted(all_months, key=matching.parse_month_to_sortable)

    all_dfs = []
    breakdown_by_month = {}
    for M in months_sorted:
        present = [sm for sm in supermodels if M in sm_rolling[sm]]
        if not present:
            continue
        per_month_results = {sm: sm_rolling[sm][M]['vectors'] for sm in present}
        df_m, bd = matching.pool_safety_stock(
            per_month_results, percentile=percentile, n_runs=None,
        )
        breakdown_by_month[M] = bd
        if df_m is None or len(df_m) == 0:
            continue
        troncata = any(sm_rolling[sm][M].get('troncata', False) for sm in present)
        df_m = df_m.copy()
        df_m.insert(0, 'mese_lancio', M)
        df_m.insert(1, 'finestra_troncata', troncata)
        all_dfs.append(df_m)

    if not all_dfs:
        return pd.DataFrame(), breakdown_by_month
    return pd.concat(all_dfs, ignore_index=True), breakdown_by_month

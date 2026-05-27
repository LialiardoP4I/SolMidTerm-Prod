# -*- coding: utf-8 -*-
"""Smoke test per SafetyStockPipeline (Fase 6 parte 1).

Copre:
  - costruzione (__init__) con/senza validate_config
  - run_step_0_load_tr su dati reali della cartella Input/
  - run_step_2_montecarlo fingerprint contro baseline
"""
from pathlib import Path

import pytest

from mid_term import PipelineConfig, SimulationConfig
from mid_term.exceptions import ConfigurationError
from mid_term.pipeline import SafetyStockPipeline


def test_pipeline_construction(tmp_path):
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    sim = SimulationConfig(n_runs=20)
    pipeline = SafetyStockPipeline(cfg, sim, validate_config=False)
    assert pipeline.config is cfg
    assert pipeline.sim_config is sim


def test_pipeline_validate_failure(tmp_path):
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    with pytest.raises(ConfigurationError):
        SafetyStockPipeline(cfg, validate_config=True)


def test_run_step_0_on_real_data(tmp_path):
    """Step 0 deve produrre TRData con 5 modelli MSV4-family e 51 caratteristiche."""
    cfg = PipelineConfig(
        input_dir=Path("Input"), output_dir=tmp_path,
        write_intermediate=False, write_final_output=False,
    )
    sim = SimulationConfig(n_runs=20, random_seed=42)
    p = SafetyStockPipeline(cfg, sim, validate_config=False)
    tr = p.run_step_0_load_tr()
    assert sorted(tr.model_mix.models) == ['MSV4', 'MSV4PP', 'MSV4RI', 'MSV4RS', 'MSV4S']
    assert len(tr.characteristics) == 51
    assert tr.elapsed_seconds > 0


def test_run_step_2_montecarlo_fingerprint(tmp_path):
    """Step 2 deve riprodurre fingerprint baseline MC."""
    import json
    import hashlib
    BASELINE = Path("tests/_baseline_fingerprints.json")
    if not BASELINE.exists():
        pytest.skip("baseline non disponibile")
    expected = json.loads(BASELINE.read_text(encoding="utf-8"))

    # NOTA: il baseline usa forecast in-memory {Jan/Feb/Mar 2026: 100}, NON load_monthly_forecast.
    # Ricostruiamo manualmente per replicare le condizioni baseline.
    from mid_term.tr import load_monthly_tr, parse_tr_file
    from mid_term.montecarlo import (
        SimulationConfig as MCCfg, run_monthly_simulations_optimized,
    )
    import pandas as pd

    fixture = Path("tests/fixtures/TR_test.xlsx")
    forecast = pd.Series({"Jan 2026": 100, "Feb 2026": 100, "Mar 2026": 100}, name="demand")
    monthly_tr = load_monthly_tr(str(fixture))
    model_mix, characteristics = parse_tr_file(str(fixture))
    cfg_mc = MCCfg(n_runs=20, random_seed=42, use_int16=True,
                   max_conflict_resolution_iterations=10, start_month_offset=0)
    results = run_monthly_simulations_optimized(
        forecast_values=list(forecast.values),
        month_names=list(forecast.index),
        n_months=len(forecast),
        model_mix=model_mix,
        characteristics=characteristics,
        exclusions={},
        config=cfg_mc,
        monthly_tr=monthly_tr,
    )
    parts = []
    for month_idx in sorted(results.keys()):
        for cfg_key in sorted(results[month_idx].keys(), key=str):
            arr = results[month_idx][cfg_key]
            parts.append(f"m{month_idx}|{cfg_key}|sum={arr.sum()}|n={len(arr)}")
    fp = hashlib.sha256("\n".join(parts).encode()).hexdigest()
    assert fp == expected["mc_fingerprint"]

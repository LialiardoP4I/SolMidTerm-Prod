# -*- coding: utf-8 -*-
"""Regression test Monte Carlo: verifica che il refactor di mid_term/montecarlo.py
preservi il fingerprint deterministico catturato in baseline (n_runs=20, seed=42)."""
import hashlib
import json
from pathlib import Path

import pandas as pd
import pytest

from mid_term.tr import load_monthly_tr, parse_tr_file
from mid_term.montecarlo import SimulationConfig, run_monthly_simulations_optimized

TR_FIXTURE = Path(__file__).parent / "fixtures" / "TR_test.xlsx"
BASELINE = Path(__file__).parent / "_baseline_fingerprints.json"
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
        pytest.skip("Baseline non disponibile")
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def test_mc_fingerprint_matches_baseline(baseline):
    monthly_tr = load_monthly_tr(str(TR_FIXTURE))
    model_mix, characteristics = parse_tr_file(str(TR_FIXTURE))
    cfg = SimulationConfig(n_runs=20, random_seed=42, use_int16=True,
                           max_conflict_resolution_iterations=10, start_month_offset=0)
    results = run_monthly_simulations_optimized(
        forecast_values=list(FORECAST.values),
        month_names=list(FORECAST.index),
        n_months=len(FORECAST),
        model_mix=model_mix,
        characteristics=characteristics,
        exclusions={},
        config=cfg,
        monthly_tr=monthly_tr,
    )
    assert _hash_mc_results(results) == baseline["mc_fingerprint"]

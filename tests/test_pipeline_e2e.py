# -*- coding: utf-8 -*-
"""End-to-end test della pipeline su dati reali (Fase 6 parte 2).

Copre:
  - run_step_1_build_bom su catalogo multi-supermodel
  - run() completo (step 0..4) -> SafetyStockResult finale

NOTA: il run() completo legge dati reali da Input/ ed esegue Monte Carlo
+ SKU matching: tempi nell'ordine di alcuni minuti anche con n_runs ridotto.
"""
from pathlib import Path

import pytest

from mid_term import PipelineConfig, SimulationConfig
from mid_term.pipeline import SafetyStockPipeline


@pytest.fixture
def real_pipeline(tmp_path):
    cfg = PipelineConfig(
        input_dir=Path("Input"),
        output_dir=tmp_path,
        write_intermediate=False,
        write_final_output=False,
    )
    sim = SimulationConfig(n_runs=5, random_seed=42, percentile_safety_stock=99)
    return SafetyStockPipeline(cfg, sim, validate_config=False)


def test_step_1_build_bom(real_pipeline):
    """Step 1 deve produrre BOMData con 3 supermodel (V21E, V22E, V24T)."""
    bom = real_pipeline.run_step_1_build_bom()
    assert len(bom.matrices) == 3
    assert set(bom.matrices.keys()) == {"V21E", "V22E", "V24T"}
    assert len(bom.catalog) > 0
    assert bom.elapsed_seconds > 0


def test_full_run_end_to_end(real_pipeline):
    """run() completo deve produrre SafetyStockResult non vuoto."""
    result = real_pipeline.run()
    assert result.per_sku is not None and len(result.per_sku) > 0
    assert result.per_mese is not None and len(result.per_mese) > 0
    assert result.percentile == 99
    assert result.elapsed_seconds > 0

# -*- coding: utf-8 -*-
"""Test PipelineConfig + SimulationConfig."""
import pytest

from mid_term.config import PipelineConfig, SimulationConfig
from mid_term.exceptions import ConfigurationError


def test_pipeline_config_resolve_paths(tmp_path):
    cfg = PipelineConfig(input_dir=tmp_path / "Input", output_dir=tmp_path / "Output")
    assert cfg.resolve_input("foo.xlsx") == tmp_path / "Input" / "foo.xlsx"
    assert cfg.resolve_output("bar.xlsx") == tmp_path / "Output" / "bar.xlsx"


def test_pipeline_config_validate_missing_input(tmp_path):
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

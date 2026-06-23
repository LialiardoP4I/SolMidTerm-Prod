"""Test di orchestrazione per collect_run_vectors_rolling.

Un run reale e' troppo pesante (BOM/MC/matching su file Excel veri), quindi qui
si MONKEYPATCHA tutto il setup nel namespace del modulo `rolling` con stand-in
sintetici minimi. L'obiettivo NON e' verificare i numeri (coperti altrove) ma la
FORMA dell'orchestrazione: chiavi = mesi di lancio, flag troncata corretto,
offset progressivo, e che i vettori del matching finiscano nella struttura.

Determinismo: collect_run_vectors_rolling NON usa datetime.now (a differenza di
run_rolling), quindi i test sono ripetibili.
"""
from pathlib import Path

import numpy as np
import pandas as pd

from mid_term_supermodels import rolling
from mid_term_supermodels.config import PipelineConfig, SimulationConfig


# --- Stand-in del forecast: 3 mesi a partire da offset 0, code per offset>0 ---
_ALL_MONTHS = ["JAN 2027", "FEB 2027", "MAR 2027"]
_ALL_VALUES = [10.0, 20.0, 30.0]


def _fake_load_monthly_forecast(forecast_path, start_offset=0):
    return _ALL_VALUES[start_offset:], _ALL_MONTHS[start_offset:]


class _FakeBOM:
    def __init__(self):
        # Catalogo minimo: una colonna feature ('MODEL') + 'SKU'.
        self.catalog = pd.DataFrame({"SKU": ["SKU1"], "MODEL": ["M1"]})


class _FakeSim:
    def __init__(self):
        # df_wide minimo con una colonna config e una colonna _Run_.
        self.df_wide = pd.DataFrame({"MODEL": ["M1"], "JAN 2027_Run_1": [10]})
        self.n_runs = 3


class _FakePipeline:
    def __init__(self, config, sim_config, validate_config=True):
        self.config = config
        self.sim_config = sim_config

    def run_step_1_build_bom(self):
        return _FakeBOM()

    def run_step_2_montecarlo(self, tr):
        return _FakeSim()


def _patch_common(monkeypatch):
    monkeypatch.setattr(rolling, "load_monthly_forecast", _fake_load_monthly_forecast)
    monkeypatch.setattr(rolling, "calculate_n_months_needed", lambda: 2)
    monkeypatch.setattr(rolling, "load_monthly_tr", lambda path: {})
    monkeypatch.setattr(rolling, "parse_tr_file", lambda path: (None, {}))
    monkeypatch.setattr(rolling, "rescale_alphas",
                        lambda mm, ch, mtr, **kw: (mm, ch, mtr))
    monkeypatch.setattr(rolling, "parse_exclusions", lambda path: {})
    monkeypatch.setattr(rolling, "SafetyStockPipeline", _FakePipeline)
    monkeypatch.setattr(rolling, "build_column_mapping_auto",
                        lambda feat, simcols: {"MODEL": "MODEL"})
    monkeypatch.setattr(rolling, "build_solo_modello_flags",
                        lambda raw, cmap: pd.DataFrame(columns=["SKU", "SOLO_MODELLO"]))
    monkeypatch.setattr(rolling, "load_lead_times",
                        lambda solo_modello_skus, gate_max_mesi: pd.DataFrame(columns=["SKU"]))
    monkeypatch.setattr(rolling, "_reinit_norm_cache", lambda cmap: None)
    monkeypatch.setattr(
        rolling, "match_skus_preserve_run_vectors",
        lambda **kw: ({"SKU1": (np.array([1, 2, 3]), 1.0, 1.0, "desc")}, None),
    )


def _cfg():
    cfg = PipelineConfig(input_dir=Path("."), output_dir=Path("."))
    sim = SimulationConfig(n_runs=3)
    return cfg, sim


def test_function_exists_and_importable():
    assert hasattr(rolling, "collect_run_vectors_rolling")


def test_collect_run_vectors_rolling_shape(monkeypatch):
    _patch_common(monkeypatch)
    cfg, sim = _cfg()

    out = rolling.collect_run_vectors_rolling(cfg, sim)

    # Una entry per ciascuno dei 3 mesi di lancio, nell'ordine cronologico.
    assert list(out.keys()) == _ALL_MONTHS

    for mese in _ALL_MONTHS:
        entry = out[mese]
        assert set(entry.keys()) == {"vectors", "troncata", "offset", "mesi_simulati"}
        assert "SKU1" in entry["vectors"]
        rd, qty, lt, desc = entry["vectors"]["SKU1"]
        assert rd.tolist() == [1, 2, 3]
        assert qty == 1.0 and lt == 1.0 and desc == "desc"

    # offset progressivo 0,1,2
    assert [out[m]["offset"] for m in _ALL_MONTHS] == [0, 1, 2]


def test_troncata_flags(monkeypatch):
    # n_months_needed=2, total=3:
    #   offset0 remaining=3 -> piena;  offset1 remaining=2 -> piena;
    #   offset2 remaining=1 < 2 -> troncata. mesi_simulati = min(2, remaining).
    _patch_common(monkeypatch)
    cfg, sim = _cfg()

    out = rolling.collect_run_vectors_rolling(cfg, sim)

    assert out["JAN 2027"]["troncata"] is False
    assert out["FEB 2027"]["troncata"] is False
    assert out["MAR 2027"]["troncata"] is True

    assert out["JAN 2027"]["mesi_simulati"] == 2
    assert out["FEB 2027"]["mesi_simulati"] == 2
    assert out["MAR 2027"]["mesi_simulati"] == 1


def test_max_iterations_limits_months(monkeypatch):
    _patch_common(monkeypatch)
    cfg, sim = _cfg()

    out = rolling.collect_run_vectors_rolling(cfg, sim, max_iterations=1)
    assert list(out.keys()) == ["JAN 2027"]

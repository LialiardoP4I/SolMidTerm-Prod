"""Test path-level: TR.xlsx in config + PREZZI letto da Dati Logistica condivisa.

Questi test sono volutamente leggeri: verificano il VALORE di config
(tr_file) e l'ORDINE di ricerca del file PREZZI (Dati Logistica condivisa
prima della input_path del supermodel), senza richiedere fixture Excel pesanti.
"""
import inspect
from pathlib import Path

from mid_term_supermodels import config, matching, pipeline


def test_default_tr_file_is_tr_xlsx():
    # Il nome di default del file TR deve essere "TR.xlsx".
    cfg = config.PipelineConfig(input_dir=Path("."), output_dir=Path("."))
    assert cfg.tr_file == "TR.xlsx"


def test_residual_lt_cerca_prezzi_in_logistica_prima():
    # PREZZI condiviso: la ricerca deve usare logistica_path (Dati Logistica)
    # PRIMA della input_path del supermodel (fallback). Verifica sull'ordine.
    src = inspect.getsource(matching.calcola_residual_lead_time)
    i_logist = src.find('logistica_path.glob("PREZZI')
    i_input = src.find('input_path.glob("PREZZI')
    assert i_logist != -1, "lookup PREZZI su logistica_path non trovato"
    assert i_input != -1, "fallback PREZZI su input_path non trovato"
    assert i_logist < i_input, "logistica_path deve essere cercato prima del fallback"


def test_step4_cerca_prezzi_in_logistica_prima():
    # PREZZI condiviso in step 4: da shared_logistica_dir(), con fallback
    # a self.config.input_dir.
    src = inspect.getsource(pipeline.SafetyStockPipeline.run_step_4_safety_stock)
    i_shared = src.find("shared_logistica_dir()")
    i_input = src.find("self.config.input_dir.glob")
    assert i_shared != -1, "lookup PREZZI su shared_logistica_dir non trovato"
    assert i_input != -1, "fallback PREZZI su input_dir non trovato"
    assert i_shared < i_input, "shared_logistica_dir deve essere cercato prima del fallback"

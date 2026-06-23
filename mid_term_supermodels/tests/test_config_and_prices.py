"""Test path-level: TR.xlsx in config + PREZZI letto per-supermodel.

Questi test sono volutamente leggeri: verificano il VALORE di config
(tr_file) e l'ORDINE di ricerca del file PREZZI (per-supermodel input_path
prima della Dati Logistica condivisa), senza richiedere fixture Excel pesanti.
"""
import inspect
from pathlib import Path

from mid_term_supermodels import config, matching, pipeline


def test_default_tr_file_is_tr_xlsx():
    # CHANGE 1: il nome di default del file TR deve essere "TR.xlsx".
    cfg = config.PipelineConfig(input_dir=Path("."), output_dir=Path("."))
    assert cfg.tr_file == "TR.xlsx"


def test_residual_lt_cerca_prezzi_in_input_path_prima():
    # CHANGE 2 (matching): la ricerca PREZZI deve usare input_path PRIMA
    # del logistica_path (fallback). Verifica sull'ordine nel sorgente.
    src = inspect.getsource(matching.calcola_residual_lead_time)
    i_input = src.find('input_path.glob("PREZZI')
    i_logist = src.find('logistica_path.glob("PREZZI')
    assert i_input != -1, "lookup PREZZI su input_path non trovato"
    assert i_logist != -1, "fallback PREZZI su logistica_path non trovato"
    assert i_input < i_logist, "input_path deve essere cercato prima del fallback"


def test_step4_cerca_prezzi_in_input_dir_prima():
    # CHANGE 2 (pipeline step 4): PREZZI da self.config.input_dir, con
    # fallback a shared_logistica_dir().
    src = inspect.getsource(pipeline.SafetyStockPipeline.run_step_4_safety_stock)
    i_input = src.find("self.config.input_dir.glob")
    i_shared = src.find("shared_logistica_dir()")
    assert i_input != -1, "lookup PREZZI su input_dir non trovato"
    assert i_shared != -1, "fallback PREZZI su shared_logistica_dir non trovato"
    assert i_input < i_shared, "input_dir deve essere cercato prima del fallback"

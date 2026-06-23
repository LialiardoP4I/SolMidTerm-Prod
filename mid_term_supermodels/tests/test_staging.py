# -*- coding: utf-8 -*-
"""Test dello staging (copia) per A-staging.

Verifica che stage_supermodel_input assembli un Input/ completo COPIANDO il
per-supermodel + i condivisi, e che rmtree sulla copia non tocchi gli originali.
"""
import shutil
from pathlib import Path

from mid_term_supermodels import multi_supermodel


def _build_fixture(tmp_path):
    inp = tmp_path / 'Input'
    sm_dir = inp / 'Multistrada'
    (sm_dir / 'MODEL').mkdir(parents=True)
    (sm_dir / 'MODEL' / 'dummy.txt').write_text('m')
    (sm_dir / 'Total_demand.xlsx').write_text('demand')
    (sm_dir / 'TR X.xlsx').write_text('tr')

    (inp / 'Dati Logistica').mkdir()
    (inp / 'Dati Logistica' / 'PREZZI.XLSX').write_text('prezzi')
    (inp / 'States.xlsx').write_text('states')
    (inp / 'esclusioni.xlsx').write_text('escl')
    (inp / 'Cambio Valuta.xlsx').write_text('cambio')
    (inp / 'affidabilita_config.json').write_text('{}')
    return inp, sm_dir


def test_stage_supermodel_input(tmp_path):
    inp, sm_dir = _build_fixture(tmp_path)
    work_root = tmp_path / '_work_supermodels'

    work_sm = multi_supermodel.stage_supermodel_input(sm_dir, inp, work_root)
    work_sm = Path(work_sm)
    staged = work_sm / 'Input'

    # Per-supermodel copiato.
    assert (staged / 'MODEL').is_dir()
    assert (staged / 'MODEL' / 'dummy.txt').exists()
    assert (staged / 'Total_demand.xlsx').exists()
    assert (staged / 'TR X.xlsx').exists()

    # Condivisi copiati.
    assert (staged / 'States.xlsx').exists()
    assert (staged / 'esclusioni.xlsx').exists()
    assert (staged / 'Cambio Valuta.xlsx').exists()
    assert (staged / 'affidabilita_config.json').exists()
    assert (staged / 'Dati Logistica').is_dir()
    assert (staged / 'Dati Logistica' / 'PREZZI.XLSX').exists()

    # Il path staged e' sotto _work_supermodels (mai dentro Input/ reale).
    assert '_work_supermodels' in str(work_sm)

    # rmtree sulla copia non tocca gli originali.
    shutil.rmtree(work_sm)
    assert not work_sm.exists()
    assert sm_dir.exists()
    assert (sm_dir / 'Total_demand.xlsx').exists()
    assert (inp / 'States.xlsx').exists()
    assert (inp / 'Dati Logistica' / 'PREZZI.XLSX').exists()


def test_stage_overwrites_existing(tmp_path):
    """Se work_sm esiste gia', viene ripulito (clean slate) prima di ricopiare."""
    inp, sm_dir = _build_fixture(tmp_path)
    work_root = tmp_path / '_work_supermodels'

    work_sm = multi_supermodel.stage_supermodel_input(sm_dir, inp, work_root)
    # Sporca lo staging precedente con un file estraneo.
    (Path(work_sm) / 'Input' / 'STALE.txt').write_text('stale')

    work_sm2 = multi_supermodel.stage_supermodel_input(sm_dir, inp, work_root)
    assert not (Path(work_sm2) / 'Input' / 'STALE.txt').exists()
    assert (Path(work_sm2) / 'Input' / 'Total_demand.xlsx').exists()

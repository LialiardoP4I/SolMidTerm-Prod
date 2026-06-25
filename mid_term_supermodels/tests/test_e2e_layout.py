# -*- coding: utf-8 -*-
"""Smoke end-to-end del CONTRATTO DI LAYOUT (A-staging, flat).

Non esegue la pipeline Monte Carlo (troppo pesante / richiede veri Excel):
verifica che discovery + staging + os.chdir(<work>/<SM>) rispettino il
contratto, e che con env NON impostata config.shared_logistica_dir() risolva
al default ./Input/Dati Logistica (la copia staged).
"""
import os
from pathlib import Path

from mid_term_supermodels import multi_supermodel, config


def test_layout_contract(tmp_path):
    base = tmp_path / 'Input'
    (base / 'Dati Logistica').mkdir(parents=True)
    (base / 'States.xlsx').write_text('x')
    (base / 'esclusioni.xlsx').write_text('x')
    (base / 'Cambio Valuta.xlsx').write_text('x')
    (base / 'affidabilita_config.json').write_text('{}')
    for sm in ['SuperA', 'SuperB']:
        sm_dir = base / sm
        (sm_dir / 'MODEL').mkdir(parents=True)
        (sm_dir / 'Total_demand.xlsx').write_text('x')
        (sm_dir / 'TR FOO.xlsx').write_text('x')

    found = multi_supermodel.discover_supermodels(str(base))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']

    # Staging + chdir: ./Input/MODEL, ./Input/Total_demand.xlsx e i condivisi risolvono.
    work_root = tmp_path / '_work_supermodels'
    cwd0 = os.getcwd()
    try:
        work_sm = multi_supermodel.stage_supermodel_input(found[0], base, work_root)
        os.chdir(work_sm)
        assert Path('Input/MODEL').is_dir()
        assert Path('Input/Total_demand.xlsx').exists()
        assert Path('Input/States.xlsx').exists()
        assert Path('Input/Dati Logistica').is_dir()
        # env NON impostata -> default risolve alla copia staged.
        _prev = os.environ.pop('MIDTERM_SHARED_LOGISTICA', None)
        try:
            assert config.shared_logistica_dir() == Path('.') / 'Input' / 'Dati Logistica'
            assert config.shared_logistica_dir().is_dir()
        finally:
            if _prev is not None:
                os.environ['MIDTERM_SHARED_LOGISTICA'] = _prev
    finally:
        os.chdir(cwd0)


def test_shared_logistica_default_when_env_unset():
    """Fallback storico: env non impostata -> <cwd>/Input/Dati Logistica."""
    _prev = os.environ.pop('MIDTERM_SHARED_LOGISTICA', None)
    try:
        assert config.shared_logistica_dir() == Path('.') / 'Input' / 'Dati Logistica'
    finally:
        if _prev is not None:
            os.environ['MIDTERM_SHARED_LOGISTICA'] = _prev

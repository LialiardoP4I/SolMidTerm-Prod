# -*- coding: utf-8 -*-
"""Smoke end-to-end del CONTRATTO DI LAYOUT (opzione B).

Non esegue la pipeline Monte Carlo (troppo pesante / richiede veri Excel):
verifica che discovery + risoluzione path + os.chdir(sm_root) rispettino il
contratto, e che l'override Dati Logistica condivisa sia raccolto da
config.shared_logistica_dir() quando MIDTERM_SHARED_LOGISTICA e' impostata.
"""
import os
from pathlib import Path

from mid_term_supermodels import multi_supermodel, config


def test_layout_contract(tmp_path):
    base = tmp_path / 'Input'
    (base / 'Dati Logistica').mkdir(parents=True)
    for sm in ['SuperA', 'SuperB']:
        inp = base / sm / 'Input'
        (inp / 'MODEL').mkdir(parents=True)
        (inp / 'Total_demand.xlsx').write_text('x')
        (inp / 'TR FOO.xlsx').write_text('x')

    found = multi_supermodel.discover_supermodels(str(base))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']

    # chdir nella ROOT del supermodel -> ./Input/MODEL e ./Input/Total_demand.xlsx risolvono
    cwd0 = os.getcwd()
    try:
        os.chdir(found[0])
        assert Path('Input/MODEL').is_dir()
        assert Path('Input/Total_demand.xlsx').exists()
    finally:
        os.chdir(cwd0)

    # override Dati Logistica condivisa
    shared = str(base / 'Dati Logistica')
    os.environ['MIDTERM_SHARED_LOGISTICA'] = shared
    try:
        assert str(config.shared_logistica_dir()) == shared
    finally:
        del os.environ['MIDTERM_SHARED_LOGISTICA']


def test_shared_logistica_default_when_env_unset():
    """Fallback storico: env non impostata -> <cwd>/Input/Dati Logistica."""
    _prev = os.environ.pop('MIDTERM_SHARED_LOGISTICA', None)
    try:
        assert config.shared_logistica_dir() == Path('.') / 'Input' / 'Dati Logistica'
    finally:
        if _prev is not None:
            os.environ['MIDTERM_SHARED_LOGISTICA'] = _prev

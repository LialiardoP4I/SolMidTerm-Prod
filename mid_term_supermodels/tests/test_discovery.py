import os
from pathlib import Path
from mid_term_supermodels import multi_supermodel


def test_discover_supermodels(tmp_path):
    # Layout A-staging (flat): ogni supermodel ha MODEL/, Total_demand.xlsx e
    # un file TR* DIRETTAMENTE nella propria cartella (nessun Input/ annidato).
    inp = tmp_path / 'Input'
    inp.mkdir()
    for sm in ['SuperA', 'SuperB']:
        sm_dir = inp / sm
        (sm_dir / 'MODEL').mkdir(parents=True)
        (sm_dir / 'Total_demand.xlsx').write_text('x')
        (sm_dir / 'TR FOO.xlsx').write_text('x')

    # Cartelle/file condivisi: NON devono essere scoperti come supermodel.
    (inp / 'Dati Logistica').mkdir()
    (inp / 'States.xlsx').write_text('x')
    (inp / 'Cambio Valuta.xlsx').write_text('x')

    # Cartelle sorgente grezze (no MODEL/, no Total_demand/TR): IGNORATE.
    (inp / 'DVL').mkdir()
    (inp / 'PANI').mkdir()

    found = multi_supermodel.discover_supermodels(str(inp))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']

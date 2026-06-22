import os
from pathlib import Path
from mid_term_supermodels import multi_supermodel


def test_discover_supermodels(tmp_path):
    # Layout opzione B: ogni supermodel ha il proprio sotto-Input/.
    # Crea 2 cartelle supermodel valide + 1 cartella condivisa senza marker.
    for sm in ['SuperA', 'SuperB']:
        inp = tmp_path / sm / 'Input'
        (inp / 'MODEL').mkdir(parents=True)
        (inp / 'Total_demand.xlsx').write_text('x')
        (inp / 'TR FOO.xlsx').write_text('x')
    (tmp_path / 'Dati Logistica').mkdir()  # condivisa: NON deve essere un supermodel

    found = multi_supermodel.discover_supermodels(str(tmp_path))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']

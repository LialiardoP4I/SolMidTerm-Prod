import os
from pathlib import Path
import multi_supermodel


def test_discover_supermodels(tmp_path):
    # Crea 2 cartelle supermodel valide + 1 cartella senza i marker
    for sm in ['SuperA', 'SuperB']:
        d = tmp_path / sm
        (d / 'MODEL').mkdir(parents=True)
        (d / 'Total_demand.xlsx').write_text('x')
        (d / 'TR FOO.xlsx').write_text('x')
    (tmp_path / 'Dati Logistica').mkdir()  # NON deve essere considerata supermodel

    found = multi_supermodel.discover_supermodels(str(tmp_path))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']

# -*- coding: utf-8 -*-
"""Orchestratore Safety Stock multi-supermodel.

Esegue la pipeline SS per ogni cartella supermodel (domanda + TR propri,
seed diverso) e fa il pooling dei vettori per-run a livello componente.
"""
import os
from pathlib import Path
from typing import List

from _logging import get_logger

log = get_logger(__name__)


def discover_supermodels(input_dir: str) -> List[str]:
    """Trova le cartelle supermodel sotto input_dir.

    Una cartella è un supermodel se contiene: sottocartella MODEL/,
    un Total_demand.xlsx e almeno un file il cui nome inizia con 'TR'.
    """
    base = Path(input_dir)
    result = []
    for sub in sorted(base.iterdir()):
        if not sub.is_dir():
            continue
        has_model = (sub / 'MODEL').is_dir()
        has_demand = (sub / 'Total_demand.xlsx').exists()
        has_tr = any(f.name.upper().startswith('TR') and f.suffix.lower() in ('.xlsx', '.xls')
                     for f in sub.iterdir() if f.is_file())
        if has_model and has_demand and has_tr:
            result.append(str(sub))
    log.info("Supermodel trovati: %d", len(result))
    return result

# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta dedup configurazioni da mid_term.bom.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.bom import ...`. Questo modulo esiste per non rompere
consumer legacy (process_all_bom_combinations.py shim).
"""
from mid_term.bom import (  # noqa: F401
    crea_tutte_righe_univoche_v2,
    _strip_and_consolidate,
    _collapse_pack_color_columns,
    _rileva_colonne,
)

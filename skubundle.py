# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta dedup configurazioni da mid_term.bom.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.bom import ...`. Questo modulo esiste per non rompere
consumer legacy (crea_catalogo_sa.py).

NOTA: le funzioni helper `_get_available_matrix_files` e
`_cleanup_safety_stock_files` dello skubundle originale erano solo usate
nel main __main__ block dello script (CLI), MAI importate da consumer
esterni — verificato con grep sull'intero repository. Non vengono
riesportate.
"""
from mid_term.bom import (  # noqa: F401
    crea_tutte_righe_univoche_v2,
    _strip_and_consolidate,
    _collapse_pack_color_columns,
    _rileva_colonne,
)

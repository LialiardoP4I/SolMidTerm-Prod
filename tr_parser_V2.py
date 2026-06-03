# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta API da mid_term.tr.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.tr import ...`. Questo modulo esiste per non rompere
consumer legacy (app_V3.py, sku_matching_SA, mc_simulator_rolling,
main_demand_comparison_V2, stock_alignment_V2, sku_matching_SS shim, ecc.).

Riesporta tutti i simboli pubblici di mid_term.tr piu' alcuni privati
effettivamente importati dai consumer (es. _AFFIDABILITA_MAP).
"""
from mid_term.tr import *  # noqa: F401,F403

# Simboli effettivamente referenziati dai consumer legacy:
#   - parse_tr_file, load_monthly_tr, get_tr_for_month, canonical_model_name
#     (app_V3 via importlib + main_demand_comparison_V2, stock_alignment_V2,
#      mc_simulator_rolling, sku_matching_SA, sku_matching_SS)
#   - ModelMix, CharacteristicGroup (tipi pubblici usati nei type hint)
#   - parse_model_mix (entry alternativa per parser model mix)
#   - _AFFIDABILITA_MAP, _load_affidabilita_config (app_V3 + mc_simulator_rolling)
from mid_term.tr import (  # noqa: F401
    parse_tr_file,
    load_monthly_tr,
    get_tr_for_month,
    parse_model_mix,
    ModelMix,
    CharacteristicGroup,
    canonical_model_name,
    _AFFIDABILITA_MAP,
    _load_affidabilita_config,
    # Helper interni referenziati dinamicamente da mid_term.bom (linee 63, 145, 770)
    # tramite `from tr_parser_V2 import _parse_month_str` — il wildcard non li espone
    # perche' iniziano con underscore.
    _parse_month_str,
)

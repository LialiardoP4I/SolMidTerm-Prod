# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta API da mid_term.bom.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.bom import ...`. Questo modulo esiste per non rompere
consumer legacy (app_V3.py, crea_catalogo_sa.py, process_all_bom_combinations
shim).
"""
from mid_term.bom import *  # noqa: F401,F403

# Simboli effettivamente referenziati dai consumer legacy
#   - process_excel_to_matrix_v2:        crea_catalogo_sa, process_all_bom_combinations
#   - enrich_bom_version:                process_all_bom_combinations
#   - _check_residual_quality, _check_price_quality, _check_ct_cv_traceability,
#     _write_quality_report:             app_V3 (QC) + process_all_bom_combinations
#   - _build_tr_char_options,_align_to_tr_option, _build_active_cv_sets,
#     _load_states_set, _is_state_specific: process_all_bom_combinations
#   - load_mapping_from_file:            app_V3, crea_catalogo_sa, process_all_bom_combinations
#   - ConditionParser:                   app_V3 (QC inline)
#   - _find_raw_file, _bom_to_str, _DEP_COL_ENRICH: crea_catalogo_sa
from mid_term.bom import (  # noqa: F401
    process_excel_to_matrix_v2,
    enrich_bom_version,
    _check_residual_quality,
    _check_price_quality,
    _check_ct_cv_traceability,
    _write_quality_report,
    _build_tr_char_options,
    _align_to_tr_option,
    _build_active_cv_sets,
    _load_states_set,
    _is_state_specific,
    load_mapping_from_file,
    ConditionParser,
    _find_raw_file,
    _bom_to_str,
    _DEP_COL_ENRICH,
    _classify_rule,
    _apply_mostra_nascondi,
)

# Alias di retro-compatibilita' (alcuni vecchi script usano nome senza suffisso)
process_excel_to_matrix = process_excel_to_matrix_v2  # noqa: F401

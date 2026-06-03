# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta API da mid_term.montecarlo.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.montecarlo import ...`. Questo modulo esiste per non rompere
consumer legacy (app_V3.py, mc_simulator_rolling, main_demand_comparison_V2,
stock_alignment_V2, tests/_capture_baseline.py).

Riesporta tutti i simboli pubblici di mid_term.montecarlo. La funzione
top-level `main_simulator` originale (entry point CLI) NON e' migrata in
mid_term/ ed e' usata solo internamente in `if __name__ == "__main__"`:
nessun consumer la importa, quindi non viene riesportata.
"""
from mid_term.montecarlo import *  # noqa: F401,F403

# Simboli effettivamente referenziati dai consumer legacy:
#   - SimulationConfig: app_V3, mc_simulator_rolling, main_demand_comparison_V2,
#                       stock_alignment_V2, tests/_capture_baseline.py
#   - run_monthly_simulations_optimized: idem
#   - run_single_month_simulation_optimized: API pubblica
#   - build_wide_dataframe: app_V3, mc_simulator_rolling, *_demand_comparison
#   - apply_pack_colors_to_wide_df, build_zcol_pack_color_mapping: app_V3 + rolling
#   - rescale_alphas: app_V3 wrapper + mc_simulator_rolling
#   - load_monthly_forecast, calculate_n_months_needed: tutti i caller
#   - parse_exclusions: tutti i caller
#   - verify_demand_conservation: main_demand_comparison_V2, stock_alignment_V2
#   - find_conflicts, precompute_char_weights, generate_configuration_fast,
#     generate_valid_configuration_fast, get_alpha_for_option,
#     force_option_to_not, resolve_conflict_weighted,
#     get_model_tr_mapping, get_model_alpha_mapping, PrecomputedCharWeights:
#     simboli pubblici (riesposti via wildcard ma elencati per chiarezza)
from mid_term.montecarlo import (  # noqa: F401
    SimulationConfig,
    PrecomputedCharWeights,
    run_monthly_simulations_optimized,
    run_single_month_simulation_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
    build_zcol_pack_color_mapping,
    rescale_alphas,
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
    verify_demand_conservation,
    find_conflicts,
    precompute_char_weights,
    generate_configuration_fast,
    generate_valid_configuration_fast,
    get_alpha_for_option,
    get_model_tr_mapping,
    get_model_alpha_mapping,
    force_option_to_not,
    resolve_conflict_weighted,
)

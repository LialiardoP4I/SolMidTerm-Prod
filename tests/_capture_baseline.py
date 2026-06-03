# -*- coding: utf-8 -*-
"""Cattura fingerprint output pipeline corrente per regression test post-refactor.

Esegue Monte Carlo sui dati reali con seed fissato e salva hash deterministici
di tutti gli output rilevanti. Eseguire UNA VOLTA prima del refactoring; i
fingerprint diventano riferimento contro cui ogni step viene verificato.
"""
import hashlib
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd

# Aggiunge la root del progetto al path (i moduli legacy sono lì)
sys.path.insert(0, str(Path(__file__).parent.parent))

from tr_parser_V2 import load_monthly_tr, parse_tr_file
from mc_simulator_V2 import SimulationConfig, run_monthly_simulations_optimized

TR_PATH = Path(__file__).parent / 'fixtures' / 'TR_test.xlsx'
FORECAST = pd.Series({'Jan 2026': 100, 'Feb 2026': 100, 'Mar 2026': 100}, name='demand')


def hash_mc_results(results: dict) -> str:
    parts = []
    for month_idx in sorted(results.keys()):
        for cfg_key in sorted(results[month_idx].keys(), key=str):
            arr = results[month_idx][cfg_key]
            parts.append(f'm{month_idx}|{cfg_key}|sum={arr.sum()}|n={len(arr)}')
    return hashlib.sha256('\n'.join(parts).encode()).hexdigest()


def main():
    monthly_tr = load_monthly_tr(str(TR_PATH))
    model_mix_global, characteristics_global = parse_tr_file(str(TR_PATH))

    cfg = SimulationConfig(
        n_runs=20, random_seed=42, use_int16=True,
        max_conflict_resolution_iterations=10, start_month_offset=0,
    )

    results = run_monthly_simulations_optimized(
        forecast_values=list(FORECAST.values),
        month_names=list(FORECAST.index),
        n_months=len(FORECAST),
        model_mix=model_mix_global,
        characteristics=characteristics_global,
        exclusions={},
        config=cfg,
        monthly_tr=monthly_tr,
    )

    fingerprints = {
        'tr_models':         sorted(monthly_tr['Jan 2026'][0].models),
        'tr_alphas_sum':     float(monthly_tr['Jan 2026'][0].alphas.sum()),
        'n_characteristics': len(characteristics_global),
        'mc_fingerprint':    hash_mc_results(results),
        'mc_total_bikes':    int(sum(arr.sum() for month in results.values() for arr in month.values())),
        'mc_n_combos':       int(sum(len(month) for month in results.values())),
    }

    out_path = Path(__file__).parent / '_baseline_fingerprints.json'
    out_path.write_text(json.dumps(fingerprints, indent=2, ensure_ascii=False, default=str),
                        encoding='utf-8')
    print(f'Baseline fingerprints salvati in {out_path}')
    print(json.dumps(fingerprints, indent=2, default=str))


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""Launcher ROLLING — alternativa a run_pipeline.py.

Esegue la safety stock in modalita' rolling (un calcolo per ogni mese di lancio)
invece del run singolo puntuale. Stessi dati e parametri di run_pipeline.py:
    mid_term/Input/          dataset di input
    mid_term/Output/         file di output (i 3 file rolling)
    mid_term/run_config.json parametri (n_runs, percentile, affidabilita';
                             opzionale: "max_iterations" per limitare i mesi)

Uso:
    cd "...\\SolMidTerm - Prod\\mid_term" ; python run_rolling.py
"""
import json
import os
import sys
from pathlib import Path

PKG = Path(__file__).resolve().parent
REPO = PKG.parent

sys.path.insert(0, str(REPO))
sys.path.insert(0, str(PKG))
os.chdir(PKG)

from mid_term_supermodels import PipelineConfig
from mid_term_supermodels.config import load_simulation_config
from mid_term_supermodels.montecarlo import load_monthly_forecast
from mid_term_supermodels.rolling import run_rolling

JSON_PATH = PKG / "run_config.json"


def main():
    """Entry point rolling: prepara config + parametri da run_config.json ed
    esegue run_rolling, stampando il riepilogo per mese di lancio."""
    config = PipelineConfig(input_dir=PKG / "Input", output_dir=PKG / "Output")

    forecast_path = config.resolve_input(config.forecast_file)
    _, month_names = load_monthly_forecast(str(forecast_path))
    sim_config = load_simulation_config(JSON_PATH, month_names)

    # Parametro opzionale max_iterations dal JSON (limita i mesi di lancio).
    max_iterations = None
    try:
        max_iterations = json.loads(JSON_PATH.read_text(encoding="utf-8")).get("max_iterations")
    except Exception:
        pass

    df_combined, df_valore = run_rolling(config, sim_config, max_iterations=max_iterations)

    print("\n=== ROLLING COMPLETATO ===")
    print(f"Mesi di lancio: {df_combined['mese_lancio'].nunique()}")
    print(f"Righe totali  : {len(df_combined)}")
    print(f"Percentile    : p{sim_config.percentile_safety_stock} | n_runs={sim_config.n_runs}")
    print("\nRiepilogo per mese di lancio:")
    print(df_valore.to_string(index=False))
    return df_combined, df_valore


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Launcher pipeline Safety Stock Classica — tutto sotto mid_term/.

Posizione: mid_term/run_pipeline.py

Uso (entrambi funzionano):
    cd "...\\SolMidTerm - Prod\\mid_term" ; python run_pipeline.py
    python "...\\SolMidTerm - Prod\\mid_term\\run_pipeline.py"

Tutti i riferimenti stanno sotto mid_term/:
    mid_term/Input/           dataset di input
    mid_term/Output/          file di output
    mid_term/run_config.json  parametri di simulazione (n_runs, start_month, ...)

NOTA: il package usa path hardcoded relativi (.\\Input\\...). Il launcher fa
os.chdir(PKG) così risolvono tutti sotto mid_term/.
"""
import os
import sys
from pathlib import Path

PKG = Path(__file__).resolve().parent       # = .../mid_term  (codice + dati + output)
REPO = PKG.parent                            # root repo (parent del package)

sys.path.insert(0, str(REPO))               # per "import mid_term"
sys.path.insert(0, str(PKG))                # per import interni del package (from config import ...)
os.chdir(PKG)                               # CWD = mid_term -> path hardcoded .\Input risolvono qui

from mid_term import SafetyStockPipeline, PipelineConfig
from mid_term.config import load_simulation_config
from mid_term.montecarlo import load_monthly_forecast

JSON_PATH = PKG / "run_config.json"


def main():
    """Entry point: prepara la configurazione (path Input/Output + parametri da
    run_config.json), istanzia ed esegue SafetyStockPipeline e stampa il
    riepilogo finale (righe per SKU, percentile, affidabilita', file prodotti)."""
    # Tutti i riferimenti sotto mid_term/
    config = PipelineConfig(input_dir=PKG / "Input", output_dir=PKG / "Output")

    # Nomi-mese dal forecast per risolvere start_month (nome -> offset)
    forecast_path = config.resolve_input(config.forecast_file)
    _, month_names = load_monthly_forecast(str(forecast_path))

    # Parametri simulazione dal JSON (obbligatorio e completo, fail-fast)
    sim_config = load_simulation_config(JSON_PATH, month_names)

    pipeline = SafetyStockPipeline(config, sim_config, validate_config=False)
    result = pipeline.run()

    print("\n=== PIPELINE COMPLETATA ===")
    print(f"Righe per SKU : {len(result.per_sku)}")
    print(f"Percentile    : p{result.percentile}")
    print(f"Affidabilità  : model={sim_config.affidabilita_model} "
          f"optional={sim_config.affidabilita_optional}")
    print(f"Output scritti: {result.output_paths}")
    return result


if __name__ == "__main__":
    main()

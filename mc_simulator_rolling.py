# -*- coding: utf-8 -*-
"""
================================================================================
ROLLING SIMULATION PIPELINE
================================================================================

Esegue la pipeline Monte Carlo + SKU Matching in modo rolling:
  - Iteration 0: mese lancio = mese 1
  - Iteration 1: mese lancio = mese 2
  - ...
  - Continua finché rimangono abbastanza mesi nel forecast (>= n_months_needed)

Output: singola tabella con colonna `mese_lancio` che indica il mese di partenza
        di ogni iterazione.

Files prodotti:
  Output/sku_safety_stock_rolling.xlsx       — aggregato per SKU, tutte le iterazioni
  Output/sku_safety_stock_PER_MESE_rolling.xlsx  — vista mensile rolling

@author: AntonioLiardo-SMOPS
================================================================================
"""

import gc
import sys
import time
from pathlib import Path

import pandas as pd

# Fix encoding per terminal Windows CP1252
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# ============================================================================
# IMPORT MODULI PIPELINE
# ============================================================================

# mc_simulator_V2: simulazione MC + rescale_alphas + pack color mapping
from mc_simulator_V2 import (
    SimulationConfig,
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
    run_monthly_simulations_optimized,
    build_wide_dataframe,
    rescale_alphas,
    build_zcol_pack_color_mapping,
    apply_pack_colors_to_wide_df,
)

# parse_tr_file e load_monthly_tr vengono da tr_parser_V2
from tr_parser_V2 import parse_tr_file, load_monthly_tr

from sku_matching_SS import main_sku_v2

try:
    from tr_parser_V2 import _AFFIDABILITA_MAP as _AFF_MAP
    _AFF_MAP = {k: float(v) for k, v in _AFF_MAP.items()}
except Exception:
    _AFF_MAP = {"BASSA": 10.0, "MEDIA": 100.0, "ALTA": 1000.0}


# ============================================================================
# FUNZIONE PRINCIPALE ROLLING
# ============================================================================

def run_rolling_simulation(
    n_runs: int = 200,
    percentile: int = 99,
    input_dir: str = ".\\Input",
    output_dir: str = ".\\Output",
    demand_file: str = ".\\Input\\Total_demand.xlsx",
    tr_file: str = ".\\Input\\TR TOTALV21E - Copia.xlsx",
    exclusions_file: str = ".\\Input\\esclusioni.xlsx",
    sku_catalog_path: str = ".\\Input\\tutte_righe_univoche_V2.xlsx",
    prezzi_path: str = ".\\Input\\Dati Logistica\\PREZZI.XLSX",
    keep_per_offset_files: bool = False,
    max_iterations: int = None,
    aff_model: str = "BASSA",
    aff_optional: str = "BASSA",
) -> pd.DataFrame:
    """
    Esegue simulazione rolling: ogni iterazione parte da un mese successivo.

    Args:
        n_runs:               Numero run Monte Carlo per iterazione
        percentile:           Percentile safety stock (default 99)
        input_dir:            Directory input
        output_dir:           Directory output
        demand_file:          File forecast domanda (Total_demand.xlsx)
        tr_file:              File Take Rate
        exclusions_file:      File esclusioni incompatibilità
        sku_catalog_path:     tutte_righe_univoche_V2.xlsx
        prezzi_path:          Prezzi.xlsx
        keep_per_offset_files: Se True mantiene file per singola iterazione
        max_iterations:       Numero massimo iterazioni (default None = tutte).
                              Se 3, esegue solo le prime 3 iterazioni (offset 0, 1, 2)

    Returns:
        DataFrame combinato con tutte le iterazioni rolling + colonna `mese_lancio`
    """

    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    print("=" * 60)
    print("ROLLING SIMULATION PIPELINE")
    print("=" * 60)

    # =========================================================================
    # STEP 0: Pre-carica dati statici (una sola volta per tutte le iterazioni)
    # =========================================================================
    print("\n[ROLLING] Caricamento dati statici...")

    # Numero mesi necessari (basato su max lead time — fisso per tutte le iterazioni)
    n_months_needed = calculate_n_months_needed()
    print(f"  Mesi necessari per simulazione: {n_months_needed}")

    # Forecast completo (tutti i mesi disponibili) per determinare range rolling
    all_forecast_values, all_month_names = load_monthly_forecast(demand_file, start_offset=0)
    total_months = len(all_forecast_values)
    print(f"  Mesi totali nel forecast: {total_months}")

    # Numero massimo di iterazioni rolling
    max_offset = total_months - n_months_needed
    if max_offset < 0:
        raise ValueError(
            f"Forecast contiene solo {total_months} mesi, "
            f"ma servono almeno {n_months_needed}. "
            f"Impossibile eseguire rolling simulation."
        )
    n_iterations = max_offset + 1

    # Limita a max_iterations se specificato
    if max_iterations is not None and max_iterations > 0:
        n_iterations = min(n_iterations, max_iterations)
        print(f"  Iterazioni rolling: {n_iterations} (offset 0..{n_iterations-1}) [LIMITATO a {max_iterations}]")
    else:
        print(f"  Iterazioni rolling: {n_iterations} (offset 0..{max_offset})")

    # TR globali (fissi — stessa struttura per tutti gli offset)
    print("\n[ROLLING] Caricamento TR globali...")
    model_mix_raw, characteristics_raw = parse_tr_file(tr_file)
    print(f"  Modelli: {len(model_mix_raw.models)}  |  Caratteristiche: {len(characteristics_raw)}")

    # TR mensili (opzionale)
    monthly_tr_raw = None
    try:
        monthly_tr_raw = load_monthly_tr(tr_file)
        print(f"  TR mensili: {len(monthly_tr_raw)} mesi")
    except Exception as _e:
        print(f"  TR mensili non disponibili ({_e}) -> uso globali")

    # Applica rescale_alphas (affidabilita TR) — identico a SS classica
    # rescale_alphas importata da mc_simulator_V2 (no dipendenza da Streamlit/app_V3)
    print(f"\n[ROLLING] Affidabilita TR -> Model: {aff_model} (VSS={_AFF_MAP.get(aff_model,0):.0f})"
          f"  |  Optional: {aff_optional} (VSS={_AFF_MAP.get(aff_optional,0):.0f})")
    model_mix, characteristics, monthly_tr = rescale_alphas(
        model_mix_raw, characteristics_raw, monthly_tr_raw,
        aff_model=aff_model, aff_optional=aff_optional,
        aff_map=_AFF_MAP,
    )
    print(f"  Alpha riscalati (Model VSS={_AFF_MAP.get(aff_model,0):.0f},"
          f" Optional VSS={_AFF_MAP.get(aff_optional,0):.0f})")

    # Esclusioni (fisse)
    exclusions = parse_exclusions(exclusions_file)
    print(f"  Esclusioni caricate")

    # =========================================================================
    # Pack color mapping (caricato una volta — identico a SS classica Step 2)
    # =========================================================================
    input_path = Path(input_dir)
    _pack_color_mapping = {}
    try:
        _map_v21e = input_path / "MODEL" / "V21E" / "Mappatura_V21E.xlsx"
        _map_v22e = input_path / "MODEL" / "V22E" / "Mappatura_V22E.xlsx"
        _map_v24t = input_path / "MODEL" / "V24T" / "Mappatura_V24T.xlsx"
        if _map_v21e.exists():
            _pack_color_mapping = build_zcol_pack_color_mapping(
                str(_map_v21e),
                rally_mappatura_path=str(_map_v22e) if _map_v22e.exists() else None,
            )
            if _map_v24t.exists():
                _vm_v24t = build_zcol_pack_color_mapping(str(_map_v24t))
                for _k, _v in _vm_v24t.items():
                    _pack_color_mapping.setdefault(_k, {}).update(_v)
            print(f"  Pack color mapping caricato ({len(_pack_color_mapping)} colonne)")
        else:
            print(f"  [WARN] Mappatura_V21E.xlsx non trovata in {_map_v21e} -> pack colors saltati")
    except Exception as _pe:
        print(f"  [WARN] Pack color mapping fallito ({_pe}) -> pack colors saltati")
        _pack_color_mapping = {}

    # =========================================================================
    # STEP 1: Loop rolling
    # =========================================================================
    all_results: list[pd.DataFrame] = []
    all_monthly_results: list[pd.DataFrame] = []

    for offset in range(n_iterations):
        mese_lancio = all_month_names[offset]
        print(f"\n{'=' * 60}")
        print(f"[ROLLING] Iterazione {offset + 1}/{n_iterations} — Mese lancio: {mese_lancio}")
        print(f"{'=' * 60}")

        t_start = time.time()

        # ── Configurazione simulazione ─────────────────────────────────────
        config = SimulationConfig(
            n_runs=n_runs,
            random_seed=42,
            start_month_offset=offset,
            use_int16=True,
        )

        # ── Carica forecast per questo offset ─────────────────────────────
        forecast_values, month_names = load_monthly_forecast(demand_file, start_offset=offset)

        # Limita ai mesi necessari
        n_sim = min(n_months_needed, len(forecast_values))

        # ── Simulazione Monte Carlo ────────────────────────────────────────
        print(f"\n  [SIM] Simulazione {n_sim} mesi × {n_runs} run...")
        results_by_month = run_monthly_simulations_optimized(
            forecast_values=forecast_values,
            month_names=month_names,
            n_months=n_sim,
            model_mix=model_mix,
            characteristics=characteristics,
            exclusions=exclusions,
            config=config,
            monthly_tr=monthly_tr,
        )

        # ── Build wide DataFrame ───────────────────────────────────────────
        df_sim = build_wide_dataframe(results_by_month, month_names, n_runs=n_runs, save_to_file=False)
        print(f"  [SIM] DataFrame simulazione: {df_sim.shape}")

        # ── Applica pack colors (identico a SS classica Step 2) ───────────
        if _pack_color_mapping:
            try:
                df_sim = apply_pack_colors_to_wide_df(df_sim, _pack_color_mapping)
                print(f"  [SIM] Pack colors applicati")
            except Exception as _pe2:
                print(f"  [WARN] apply_pack_colors fallito ({_pe2})")

        gc.collect()

        # ── Output paths per questa iterazione ────────────────────────────
        offset_tag = mese_lancio.replace(" ", "_").replace("/", "-")
        if keep_per_offset_files:
            out_agg_i     = str(output_path / f"sku_safety_stock_rolling_{offset_tag}.xlsx")
            out_monthly_i = str(output_path / f"sku_safety_stock_PER_MESE_rolling_{offset_tag}.xlsx")
        else:
            # File temporanei — verranno sovrascritti ogni iterazione, non importa
            out_agg_i     = str(output_path / "_rolling_temp_agg.xlsx")
            out_monthly_i = str(output_path / "_rolling_temp_monthly.xlsx")

        # ── SKU Matching + Safety Stock ───────────────────────────────────
        print(f"\n  [SKU] Matching e calcolo safety stock (P{percentile})...")
        iter_results = main_sku_v2(
            simulation_output=df_sim,
            percentile=percentile,
            sku_catalog_path=sku_catalog_path,
            prezzi_path=prezzi_path,
            output_agg=out_agg_i,
            output_monthly=out_monthly_i,
        )

        # ── Aggiungi colonna mese_lancio ──────────────────────────────────
        if iter_results is not None and len(iter_results) > 0:
            iter_results = iter_results.copy()
            iter_results.insert(0, "mese_lancio", mese_lancio)
            iter_results.insert(1, "offset", offset)
            all_results.append(iter_results)

            # Carica anche il file mensile se esiste
            try:
                df_monthly_i = pd.read_excel(out_monthly_i)
                df_monthly_i.insert(0, "mese_lancio", mese_lancio)
                df_monthly_i.insert(1, "offset", offset)
                all_monthly_results.append(df_monthly_i)
            except Exception:
                pass

        elapsed = time.time() - t_start
        print(f"\n  [ROLLING] Iterazione {offset + 1} completata in {elapsed:.1f}s")

        gc.collect()

    # =========================================================================
    # STEP 2: Combina tutte le iterazioni
    # =========================================================================
    print(f"\n{'=' * 60}")
    print("[ROLLING] Combinazione risultati...")
    print(f"{'=' * 60}")

    if not all_results:
        print("[WARN] Nessun risultato prodotto.")
        return pd.DataFrame()

    df_combined = pd.concat(all_results, ignore_index=True)
    print(f"  Righe totali (aggregato): {len(df_combined)}")
    print(f"  Iterazioni: {df_combined['mese_lancio'].nunique()}")

    # ── Vista aggregata per valore (totale per mese lancio) ───────────────
    ss_total_col    = next((c for c in df_combined.columns if 'safety_stock' in c.lower() and 'total' in c.lower()), None)
    prezzo_col      = 'prezzo_safety' if 'prezzo_safety' in df_combined.columns else None

    agg_dict: dict = {'SKU': 'count'}
    if ss_total_col:
        agg_dict[ss_total_col] = 'sum'
    if prezzo_col:
        agg_dict[prezzo_col] = 'sum'

    df_valore = (
        df_combined.groupby('mese_lancio', sort=False)
        .agg(agg_dict)
        .rename(columns={
            'SKU': 'n_sku',
            ss_total_col: 'ss_unita_totale' if ss_total_col else None,
            prezzo_col:   'ss_valore_eur'   if prezzo_col   else None,
        })
        .reset_index()
    )
    # Mantieni ordine cronologico offset
    _ordine = df_combined[['mese_lancio', 'offset']].drop_duplicates().sort_values('offset')
    df_valore = _ordine[['mese_lancio']].merge(df_valore, on='mese_lancio', how='left')

    print(f"\n  Vista valore per mese lancio:")
    print(df_valore.to_string(index=False))

    # Salva output rolling combinato
    out_rolling_agg     = output_path / "sku_safety_stock_rolling.xlsx"
    out_rolling_monthly = output_path / "sku_safety_stock_PER_MESE_rolling.xlsx"
    out_rolling_valore  = output_path / "sku_safety_stock_rolling_valore.xlsx"

    print(f"\n[ROLLING] Salvataggio output combinato...")
    df_combined.to_excel(str(out_rolling_agg), index=False)
    print(f"  → {out_rolling_agg}")

    df_valore.to_excel(str(out_rolling_valore), index=False)
    print(f"  → {out_rolling_valore}")

    if all_monthly_results:
        df_monthly_combined = pd.concat(all_monthly_results, ignore_index=True)
        df_monthly_combined.to_excel(str(out_rolling_monthly), index=False)
        print(f"  → {out_rolling_monthly}")

    # Pulizia file temporanei
    if not keep_per_offset_files:
        for tmp in [output_path / "_rolling_temp_agg.xlsx",
                    output_path / "_rolling_temp_monthly.xlsx"]:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass

    print(f"\n{'=' * 60}")
    print("ROLLING SIMULATION COMPLETATA!")
    print(f"  Iterazioni:  {n_iterations}")
    print(f"  SKU unici:   {df_combined['SKU'].nunique() if 'SKU' in df_combined.columns else 'N/A'}")
    print(f"  Output SKU:  {out_rolling_agg}")
    print(f"  Output €:    {out_rolling_valore}")
    print(f"{'=' * 60}")

    return df_combined, df_valore


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """
    Esecuzione diretta: python mc_simulator_rolling.py [n_runs] [percentile]

    Esempi:
      python mc_simulator_rolling.py            → 200 run, P99
      python mc_simulator_rolling.py 100        → 100 run, P99
      python mc_simulator_rolling.py 100 95     → 100 run, P95
    """
    _n_runs    = int(sys.argv[1]) if len(sys.argv) > 1 else 200
    _percentile = int(sys.argv[2]) if len(sys.argv) > 2 else 99

    print(f"Avvio rolling simulation: {_n_runs} run, P{_percentile}")
    df_combined, df_valore = run_rolling_simulation(n_runs=_n_runs, percentile=_percentile)
    print(f"\nFine. {len(df_combined)} righe totali (quantita').")
    print(f"      {len(df_valore)} righe totali (valore).")

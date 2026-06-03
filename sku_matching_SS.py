# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta API da mid_term.matching.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.matching import ...`. Questo modulo esiste per non rompere
consumer legacy (app_V3.py, mc_simulator_rolling.py).

CONCERN/BLOCKER PARZIALE — funzioni NON migrate in mid_term/matching.py:
- `main_sku_v2`  : usata da app_V3.py (line 741) e mc_simulator_rolling (line 58)
- `_match_skus_v2`: helper interno di main_sku_v2
La logica di orchestrazione e' replicata in `mid_term/pipeline.py`, ma i due
consumer importano direttamente le funzioni. Per non rompere la
retro-compatibilita' le due funzioni sono MANTENUTE inline in questo shim
(logica IDENTICA all'originale). Le building block (matching, normalizzazione,
statistiche, save_results) sono importate da `mid_term.matching` — single
source of truth.
"""
import os
import tempfile
import time
import gc
from typing import Dict

import numpy as np
import pandas as pd

# Re-export wildcard di tutti i simboli pubblici migrati
from mid_term.matching import *  # noqa: F401,F403

# Simboli espliciti effettivamente importati dai consumer legacy
from mid_term.matching import (  # noqa: F401
    CambioValutaError,
    NormalizationCache,
    match_skus_ultra_optimized,
    normalize_value_cached,
    build_column_mapping_auto,
    load_lead_times,
    calcola_residual_lead_time,
    prenormalize_simulation_output_fast,
    sku_matches_combination_numpy,
    aggregate_monthly_demands_numpy,
    calculate_statistics,
    parse_month_to_sortable,
    get_memory_usage_mb,
    save_results,
    save_results_monthly,
    _reinit_norm_cache,
    _load_cambio_eur,
)

# Dipendenza esterna usata da main_sku_v2 per la canonicalizzazione del modello
from mid_term.tr import canonical_model_name  # noqa: F401


# ============================================================================
# ENTRY POINTS V2 (orchestrazione) — non migrati in mid_term/matching.py
# Mantenuti inline qui per retro-compatibilita' con app_V3.py e
# mc_simulator_rolling.py. Logica IDENTICA all'originale sku_matching_SS.py.
# ============================================================================


def main_sku_v2(simulation_output: pd.DataFrame,
                percentile: int = 99,
                sku_catalog_path: str = '.\\Input\\tutte_righe_univoche_V2.xlsx',
                prezzi_path:      str = '.\\Input\\Prezzi.xlsx',
                output_agg:       str = '.\\Output\\sku_safety_stock_leadtime_PER_SKU_V2.xlsx',
                output_monthly:   str = '.\\Output\\sku_safety_stock_PER_MESE_V2.xlsx',
                **_kwargs) -> pd.DataFrame:
    """
    Funzione principale V2: orchestra il processo di SKU matching senza hardcoding.

    DIFFERENZE RISPETTO A V1 (main_sku_ultra_optimized):
    ----------------------------------------------------
    1. Nessun COLUMN_MAPPING hardcoded: costruito automaticamente con
       build_column_mapping_auto() dal confronto colonne SKU vs simulazione.
    2. Prefissi normalizzazione derivati dinamicamente dal column_mapping.
    3. Gestione multi-modello: righe con Model='MSV4+MSV4S' vengono espanse
       in entries separate prima del matching.
    4. Legge da tutte_righe_univoche_V2.xlsx (pipeline V2).

    Args:
        simulation_output: DataFrame dal Monte Carlo V2 simulator
        percentile:        Percentile safety stock (default 99). Range: 1-99.
        sku_catalog_path:  Path catalogo SKU V2 (default: tutte_righe_univoche_V2.xlsx)
        prezzi_path:       Path file prezzi
        output_agg:        Path output aggregato per SKU
        output_monthly:    Path output mensile

    Returns:
        DataFrame con statistiche e safety stock per ogni SKU
    """
    percentile = max(1, min(99, int(percentile)))  # clamp 1-99
    print("="*60)
    print("SKU MATCHING - VERSIONE V2 (zero hardcoding)")
    print("="*60)
    print(f"Memoria iniziale: {get_memory_usage_mb():.1f} MB")
    print("\nOttimizzazioni attive:")
    print("  - Auto-rilevamento column mapping (zero hardcoding)")
    print("  - Cache normalizzazioni (O(1) lookup) con prefissi dinamici")
    print("  - Numpy arrays per matching vettoriale")
    print("  - Matrice int16 (risparmio 75% memoria)")
    print("  - Gestione multi-modello (MSV4+MSV4S -> entries separate)")

    # =========================================================================
    # 1. Carica catalogo SKU V2
    # =========================================================================
    print("\n1. Caricamento dati...")

    sku_catalog_raw = pd.read_excel(sku_catalog_path)
    print(f"   [OK] {len(sku_catalog_raw)} righe catalogo V2")
    print(f"   Colonne: {list(sku_catalog_raw.columns)}")

    # Rimuovi colonna ID se presente (non è una caratteristica di configurazione)
    if 'ID' in sku_catalog_raw.columns:
        sku_catalog_raw = sku_catalog_raw.drop(columns=['ID'])

    # Rimuovi duplicati
    sku_catalog = sku_catalog_raw.drop_duplicates(keep='first')

    # =========================================================================
    # 2. Carica lead times
    # =========================================================================
    lead_times = load_lead_times()

    print(f"   [OK] {len(simulation_output)} combinazioni simulate")

    # =========================================================================
    # 3. Identifica numero run
    # =========================================================================
    month_run_cols = [col for col in simulation_output.columns if '_Run_' in col]
    if month_run_cols:
        first_month_cols = [col for col in month_run_cols
                           if col.startswith(month_run_cols[0].split('_Run_')[0])]
        n_runs = len(first_month_cols)
    else:
        n_runs = 1000
    print(f"   [OK] {n_runs} run per mese")

    # =========================================================================
    # 4. Costruisci column_mapping automaticamente (ZERO HARDCODING)
    # =========================================================================
    print("\n2. Auto-rilevamento column mapping...")

    # Colonne caratteristiche del catalogo SKU (escluse metadati)
    # 'versione' è aggiunta da process_all_bom_combinations.py nel catalogo multi-BOM
    # e non è una caratteristica di configurazione -> va esclusa dal matching
    sku_feature_cols = [c for c in sku_catalog.columns
                        if c.lower().strip() not in
                        {'id', 'numero componenti', 'numero componenti.1', 'versione'}]

    # Colonne caratteristiche della simulazione (escluse colonne run)
    sim_config_cols = [c for c in simulation_output.columns if '_Run_' not in c]

    column_mapping = build_column_mapping_auto(sku_feature_cols, sim_config_cols)

    print(f"   Colonne SKU disponibili: {len(sku_feature_cols)}")
    print(f"   Colonne simulazione (no run): {len(sim_config_cols)}")
    print(f"   Mapping trovato ({len(column_mapping)} colonne):")
    for sku_col, sim_col in sorted(column_mapping.items()):
        print(f"     '{sku_col}' -> '{sim_col}'")

    unmapped = [c for c in sku_feature_cols if c not in column_mapping]
    if unmapped:
        print(f"   [WARN] Colonne SKU senza corrispondenza in simulazione: {unmapped}")

    # =========================================================================
    # 4b. Calcola flag SOLO_MODELLO
    # Usa SOLO le colonne in column_mapping (caratteristiche reali di config):
    # evita di includere metadati del V2 (ID_Globale, Versioni_BOM, ecc.)
    # Condizioni:
    #   1. Tutte le colonne optional/bundle = No/Not per TUTTE le righe dello SKU
    #   2. Lo SKU NON è presente in tutti i modelli del catalogo
    # =========================================================================
    print("\n   [FLAG] Calcolo SOLO_MODELLO...")
    _model_col = next((c for c in sku_catalog_raw.columns if c.lower().strip() == 'model'), None)
    # Solo colonne effettivamente mappate a caratteristiche di configurazione, escluso Model
    _opt_flag_cols = [k for k in column_mapping.keys()
                      if _model_col is None or k.lower().strip() != 'model']

    def _is_no_val(v):
        return str(v).strip().lower() in ('no', 'nan', '')

    # Tutti i modelli distinti nel catalogo (gestisce "MSV4+MSV4S" split).
    # Canonicalizza per coerenza con la colonna Model dell'output simulazione.
    _all_models_catalog = set()
    if _model_col:
        for _mv in sku_catalog_raw[_model_col].dropna().unique():
            for _m in str(_mv).split('+'):
                _canon = canonical_model_name(_m)
                if _canon:
                    _all_models_catalog.add(_canon)

    _flag_rows = []
    for _sku, _grp in sku_catalog_raw.groupby('Numero componenti'):
        # Cond 1: tutte le celle optional/bundle = No/Not per TUTTE le righe
        _all_no = all(_is_no_val(v) for _c in _opt_flag_cols if _c in _grp.columns
                      for v in _grp[_c])
        # Cond 2: lo SKU NON copre tutti i modelli
        _sku_models = set()
        if _model_col:
            for _mv in _grp[_model_col].dropna():
                for _m in str(_mv).split('+'):
                    _canon = canonical_model_name(_m)
                    if _canon:
                        _sku_models.add(_canon)
        _not_all = bool(_all_models_catalog) and (_sku_models != _all_models_catalog)
        _flag_rows.append({'SKU': _sku, 'SOLO_MODELLO': _all_no and _not_all})

    df_flag_solo_modello = pd.DataFrame(_flag_rows)
    _n_solo = int(df_flag_solo_modello['SOLO_MODELLO'].sum())
    print(f"   [FLAG] SOLO_MODELLO: {_n_solo} SKU su {len(df_flag_solo_modello)} "
          f"(invarianti caratteristiche, non in tutti i modelli)")

    # =========================================================================
    # 5. Re-inizializza NormalizationCache con prefissi dinamici
    # =========================================================================
    _reinit_norm_cache(column_mapping)

    # =========================================================================
    # 6. Esegue matching V2
    # =========================================================================
    print("\n3. Matching SKU V2...")
    results = _match_skus_v2(
        sku_catalog=sku_catalog,
        lead_times=lead_times,
        simulation_output=simulation_output,
        column_mapping=column_mapping,
        n_runs=n_runs,
        percentile=percentile
    )

    # Aggiungi prezzi con conversione valuta -> EUR
    try:
        # Foglio 1: prezzi (include Data documento e Divisa)
        # Note: The column names have special characters like "Unitá di prezzo" with á
        # We need to read all columns first, then select the ones we need
        # IMPORTANT: Read from Sheet1, not the default first sheet
        all_data = pd.read_excel(prezzi_path, sheet_name='Sheet1')
        
        # Find the correct column names (they may have special characters)
        material_col = "Materiale"
        prezzo_netto_col = "Prezzo netto"
        unit_prezzo_col = None
        divisa_col = None
        data_doc_col = "Data documento"
        
        # Find unit di prezzo column (may have special characters)
        for col in all_data.columns:
            col_lower = str(col).lower().replace('\xa0', ' ').replace('\x92', "'")
            if 'unit' in col_lower and 'prezzo' in col_lower:
                unit_prezzo_col = col
                break
        
        # Find divisa column
        for col in all_data.columns:
            col_lower = str(col).lower().replace('\xa0', ' ').replace('\x92', "'")
            if 'divisa' in col_lower or 'currency' in col_lower:
                divisa_col = col
                break
        
        if unit_prezzo_col is None or divisa_col is None:
            raise ValueError(f"Required columns not found. unit_prezzo_col={unit_prezzo_col}, divisa_col={divisa_col}")
        
        prezzi = all_data[[material_col, prezzo_netto_col, unit_prezzo_col, divisa_col, data_doc_col]].drop_duplicates()
        # Converti a numerico: valori non parsabili (simboli valuta, "N/A", stringhe) -> NaN
        _prezzo_num = pd.to_numeric(prezzi[prezzo_netto_col], errors='coerce')
        _unit_num   = pd.to_numeric(prezzi[unit_prezzo_col],  errors='coerce').replace(0, np.nan)
        _n_bad_prezzo = int(_prezzo_num.isna().sum() - prezzi[prezzo_netto_col].isna().sum())
        _n_bad_unit   = int(_unit_num.isna().sum()   - prezzi[unit_prezzo_col].isna().sum())
        if _n_bad_prezzo > 0:
            print(f"   [INFO] Prezzi.xlsx: {_n_bad_prezzo} valori non numerici in '{prezzo_netto_col}' -> NaN")
        if _n_bad_unit > 0:
            print(f"   [INFO] Prezzi.xlsx: {_n_bad_unit} valori non numerici in '{unit_prezzo_col}' -> NaN")
        prezzi = prezzi.copy()
        prezzi["Prezzo netto"] = _prezzo_num / _unit_num

        # ── Tassi di cambio EUR — SOLO da Input/Cambio Valuta.xlsx ──────────
        # Nessun fallback patologico. Se cambio mancante/incompleto, errore esplicito
        # (CambioValutaError propaga fuori dal try outer grazie al re-raise).
        cambio = _load_cambio_eur()
        print(f"   [INFO] Tassi di cambio caricati: {len(cambio)} valute "
              f"({sorted(cambio['Codice'].unique())})")

        # Normalizza codici valuta nei prezzi prima del join
        prezzi[divisa_col] = prezzi[divisa_col].astype(str).str.strip().str.upper()

        # Join sul codice valuta
        prezzi = prezzi.merge(cambio, left_on=divisa_col, right_on='Codice', how='left')

        # Errore esplicito se ci sono valute non mappate (no più fillna 1.0 silenzioso)
        _miss_mask = prezzi['1 unità in EUR ≈'].isna()
        if _miss_mask.any():
            _missing_curr = sorted(prezzi.loc[_miss_mask, divisa_col]
                                   .dropna().astype(str).unique())
            raise CambioValutaError(
                f"Valute presenti in Prezzi.xlsx ma non mappate in Cambio Valuta.xlsx: "
                f"{_missing_curr}\n"
                f"Aggiungere queste valute al file Cambio Valuta prima di rilanciare."
            )

        # Converte il prezzo normalizzato in EUR
        prezzi["Prezzo netto EUR"] = prezzi["Prezzo netto"] * prezzi["1 unità in EUR ≈"]

        # Prende il prezzo più recente per SKU (ordina per data, last() per gruppo)
        prezzi = (prezzi[[material_col, data_doc_col, "Prezzo netto EUR"]]
                  .sort_values(data_doc_col)
                  .groupby(material_col)[["Prezzo netto EUR"]]
                  .last()
                  .rename(columns={"Prezzo netto EUR": "Prezzo netto"}))

        ss_key = f'safety_stock_p{percentile}_total'
        results = results.merge(prezzi, left_on="SKU", right_index=True, how="left")
        results["prezzo_safety"] = results["Prezzo netto"] * results[ss_key]
    except CambioValutaError:
        # Errore fatale sul cambio valuta: propaga al chiamante per stop pipeline.
        raise
    except Exception as e:
        print(f"   [WARN] Prezzi non disponibili: {e}")
        import traceback
        traceback.print_exc()

    # Merge flag SOLO_MODELLO nei risultati
    if results is None or len(results) == 0:
        print("   [WARN] Nessun risultato dal matching — risultati vuoti.")
        results = pd.DataFrame(columns=['SKU', 'SOLO_MODELLO'])
    elif 'SKU' not in results.columns:
        print(f"   [WARN] Colonna 'SKU' assente in results (colonne: {list(results.columns)[:10]})")
        results['SOLO_MODELLO'] = False
    else:
        results = results.merge(df_flag_solo_modello, on='SKU', how='left')
        results['SOLO_MODELLO'] = results['SOLO_MODELLO'].fillna(False)

    # =========================================================================
    # 7. Salva risultati
    # =========================================================================
    print("\n4. Salvataggio...")
    save_results(results, output_agg, percentile=percentile)
    save_results_monthly(results, output_monthly, percentile=percentile)

    print(f"\n{'='*60}")
    print("COMPLETATO V2!")
    print(f"  - Output aggregato: {output_agg}")
    print(f"  - Output mensile:   {output_monthly}")
    print(f"{'='*60}")

    return results


def _match_skus_v2(sku_catalog: pd.DataFrame,
                   lead_times: pd.DataFrame,
                   simulation_output: pd.DataFrame,
                   column_mapping: Dict[str, str],
                   n_runs: int,
                   percentile: int = 99) -> pd.DataFrame:
    """
    Wrapper interno che chiama match_skus_ultra_optimized con gestione
    multi-modello aggiuntiva (espansione 'MSV4+MSV4S' -> entries separate).

    La logica di matching/aggregazione/statistiche è IDENTICA a V1.
    L'unica differenza è nel modo in cui vengono preparati i dati SKU
    (STEP 3 in match_skus_ultra_optimized), qui replicato con l'espansione.
    """
    print(f"\n{'='*60}")
    print(f"MATCHING SKU V2 (ULTRA-OPTIMIZED)")
    print(f"{'='*60}")
    print(f"Memoria iniziale: {get_memory_usage_mb():.1f} MB")

    # STEP 1: Pre-normalizza colonne simulazione
    normalized_df, norm_arrays_raw = prenormalize_simulation_output_fast(
        simulation_output, column_mapping
    )
    print(f"Memoria dopo pre-norm: {get_memory_usage_mb():.1f} MB")

    category_maps = {}
    norm_arrays = {}
    for col_name in norm_arrays_raw.keys():
        if col_name in normalized_df.columns:
            cat_col = normalized_df[col_name]
            if hasattr(cat_col, 'cat'):
                category_maps[col_name] = {
                    cat: code for code, cat in enumerate(cat_col.cat.categories)
                }
                norm_arrays[col_name] = cat_col.cat.codes.values

    n_rows = len(normalized_df)

    # STEP 2: Prepara matrice numpy
    print("\n[SETUP] Preparazione matrice numpy...")
    month_run_columns = [col for col in simulation_output.columns if '_Run_' in col]
    month_prefixes = []
    seen_prefixes = set()
    for col in month_run_columns:
        if "_Run_" in col:
            prefix = col.split("_Run_")[0]
            if prefix not in seen_prefixes:
                month_prefixes.append(prefix)
                seen_prefixes.add(prefix)
    month_prefixes.sort(key=parse_month_to_sortable)

    run_columns = []
    for prefix in month_prefixes:
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in simulation_output.columns:
                run_columns.append(col_name)

    print(f"  Colonne run: {len(run_columns)}")
    print(f"  Mesi unici  : {len(month_prefixes)}")

    # Stima preventiva memoria prima di allocare
    _est_gb = n_rows * len(run_columns) * 2 / (1024 ** 3)
    print(f"  RAM stimata per data_matrix: ~{_est_gb:.1f} GB  "
          f"(shape {n_rows} × {len(run_columns)} × int16)")

    # Soglia per usare memmap (disco) invece della RAM: evita OOM su matrici grandi
    _MEMMAP_THRESHOLD_GB = 4.0
    _memmap_path = None

    if _est_gb > _MEMMAP_THRESHOLD_GB:
        # Matrice troppo grande per la RAM -> scrivi su file temporaneo (memmap)
        # numpy accede al memmap esattamente come un normale array: nessun cambio logico
        print(f"  -> Matrice grande ({_est_gb:.1f} GB > {_MEMMAP_THRESHOLD_GB} GB): "
              f"uso memmap su disco anziché RAM")
        _memmap_path = os.path.join(tempfile.gettempdir(), '_sku_data_matrix.dat')
        data_matrix = np.memmap(_memmap_path, dtype='int16', mode='w+',
                                shape=(n_rows, len(run_columns)))
        _CHUNK = 5000
        for _s in range(0, n_rows, _CHUNK):
            _e = min(_s + _CHUNK, n_rows)
            data_matrix[_s:_e] = (simulation_output.iloc[_s:_e][run_columns]
                                  .to_numpy(dtype=np.int16))
            if (_s // _CHUNK) % 20 == 0 and _s > 0:
                print(f"    {_e:,}/{n_rows:,} righe scritte su disco...")
        data_matrix.flush()
        print(f"  -> Memmap pronto: {_memmap_path}")
    else:
        data_matrix = simulation_output[run_columns].values.astype(np.int16, copy=False)

    print(f"  Matrice shape: {data_matrix.shape}")

    run_col_to_idx = {col: idx for idx, col in enumerate(run_columns)}
    month_col_indices = {}
    for month_idx, prefix in enumerate(month_prefixes):
        indices = []
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in run_col_to_idx:
                indices.append(run_col_to_idx[col_name])
        month_col_indices[month_idx] = indices

    # STEP 3: Prepara dati SKU con espansione multi-modello
    print("\n[SETUP] Preparazione dati SKU V2 (con espansione multi-modello)...")

    sku_with_lt = sku_catalog.merge(
        lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                    'Quantity_Per_Bike', 'Description']],
        left_on='Numero componenti',
        right_on='SKU',
        how='left'
    )
    sku_with_lt = sku_with_lt[sku_with_lt['Lead_Time_Months_Ceil'].notna()]
    sku_with_lt = sku_with_lt[sku_with_lt['Lead_Time_Months_Ceil'] > 0]
    # Mantieni come float per supportare interpolazione lineare (0.25, 0.5, etc.)
    # sku_with_lt['Lead_Time_Months_Ceil'] = sku_with_lt['Lead_Time_Months_Ceil'].astype(int)
    sku_with_lt['Quantity_Per_Bike'] = sku_with_lt['Quantity_Per_Bike'].fillna(1.0)

    sku_groups = sku_with_lt.groupby('SKU')
    n_sku_groups = len(sku_groups)
    print(f"  SKU da processare: {n_sku_groups}")

    # Identifica la colonna Model nel mapping (per espansione multi-modello)
    model_sku_col = next((c for c in column_mapping if c.lower().strip() == 'model'), None)

    sku_ids = []
    sku_data_list = []

    for sku_id, sku_rows_group in sku_groups:
        sku_ids.append(sku_id)
        rows_values = []

        for _, row in sku_rows_group.iterrows():
            row_values = {}
            for sku_col in column_mapping.keys():
                if sku_col in row.index:
                    row_values[sku_col] = row[sku_col]

            # --- ESPANSIONE MULTI-MODELLO (V2) ---
            # Se Model='MSV4+MSV4S', crea due entries: una per MSV4, una per MSV4S.
            # La logica OR esistente le gestisce: un componente con Model='MSV4+MSV4S'
            # matcha qualsiasi configurazione con MODEL='MSV4' O MODEL='MSV4S'.
            if model_sku_col and model_sku_col in row_values:
                model_val = str(row_values[model_sku_col]).strip()
                if '+' in model_val:
                    for single_model in model_val.split('+'):
                        expanded = dict(row_values)
                        expanded[model_sku_col] = single_model.strip()
                        rows_values.append(expanded)
                else:
                    rows_values.append(row_values)
            else:
                rows_values.append(row_values)

        sku_data_list.append({
            'values': rows_values,
            'lead_time': float(sku_rows_group['Lead_Time_Months_Ceil'].max()),
            'quantity': float(sku_rows_group['Quantity_Per_Bike'].mean()),
            'description': sku_rows_group['Description'].iloc[0]
                           if 'Description' in sku_rows_group.columns else '',
            'lead_time_raw': sku_rows_group['Lead_Time_Months'].iloc[0]
                             if 'Lead_Time_Months' in sku_rows_group.columns else np.nan
        })

    # STEP 4: Processing loop (identico a V1)
    print(f"\n[MATCHING] Elaborazione {n_sku_groups} SKU...")
    print(f"  Progress: ogni '.' = 50 SKU")

    start_time = time.time()
    all_results = []

    for sku_idx, (sku_id, sku_data) in enumerate(zip(sku_ids, sku_data_list)):
        if sku_idx % 50 == 0:
            print('.', end='', flush=True)
        if sku_idx % 500 == 0 and sku_idx > 0:
            elapsed = time.time() - start_time
            rate = sku_idx / elapsed
            remaining = (n_sku_groups - sku_idx) / rate if rate > 0 else 0
            print(f' [{sku_idx}/{n_sku_groups}] {rate:.1f} SKU/s, ETA: {remaining/60:.1f}min')

        combined_mask = np.zeros(n_rows, dtype=bool)
        for row_values in sku_data['values']:
            row_mask = sku_matches_combination_numpy(
                row_values, norm_arrays, category_maps, column_mapping, n_rows
            )
            combined_mask |= row_mask

        matching_indices = np.where(combined_mask)[0]
        if len(matching_indices) == 0:
            continue

        run_demands = aggregate_monthly_demands_numpy(
            matching_indices, data_matrix, month_col_indices,
            sku_data['lead_time'], n_runs
        )
        stats = calculate_statistics(run_demands, percentile=percentile)

        monthly_stats = {}
        for month_idx, prefix in enumerate(month_prefixes):
            col_indices_month = month_col_indices.get(month_idx, [])
            if not col_indices_month:
                continue
            month_data = data_matrix[np.ix_(matching_indices, col_indices_month)]
            run_demands_month = month_data.sum(axis=0).astype(np.int32)
            monthly_stats[prefix] = calculate_statistics(run_demands_month,
                                                          percentile=percentile)

        ss_key = f'safety_stock_p{percentile}'
        result = {
            'SKU': sku_id,
            'Description': sku_data['description'],
            'Quantity_Per_Bike': sku_data['quantity'],
            'Lead_Time_Months': sku_data['lead_time_raw'],
            'Lead_Time_Months_Ceil': sku_data['lead_time'],
            'N_Matching_Combinations': len(matching_indices),
            'N_SKU_Rows_Original': len(sku_data['values']),
            'monthly_stats': monthly_stats,
        }
        result.update(stats)
        result[f'{ss_key}_total'] = result[ss_key] * sku_data['quantity']
        all_results.append(result)

        if sku_idx % 1000 == 0:
            gc.collect()

    elapsed = time.time() - start_time
    rate = len(all_results) / elapsed if elapsed > 0 else 0
    print(f"\n\n[OK] Matching V2 completato!")
    print(f"  Tempo: {elapsed:.1f}s ({elapsed/60:.1f}min)")
    print(f"  Velocità: {rate:.1f} SKU/s")
    print(f"  SKU con match: {len(all_results)}")

    # Cleanup file memmap temporaneo (se usato)
    if _memmap_path is not None:
        del data_matrix
        try:
            os.remove(_memmap_path)
            print(f"  -> File memmap temporaneo rimosso: {_memmap_path}")
        except Exception:
            pass

    return pd.DataFrame(all_results)



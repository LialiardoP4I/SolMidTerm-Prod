# -*- coding: utf-8 -*-
"""
================================================================================
SKU MATCHING - VERSIONE V2 (compatibile con pipeline V2)
================================================================================

DIFFERENZE RISPETTO ALLA VERSIONE ORIGINALE (sku_matching_ULTRA_OPTIMIZED.py):
-------------------------------------------------------------------------------
1. NESSUN COLUMN_MAPPING HARDCODED:
   Il mapping colonne SKU -> simulazione viene costruito automaticamente
   con build_column_mapping_auto() tramite matching case-insensitive.
   Aggiungere nuovi bundle/caratteristiche al TR non richiede modifiche al codice.

2. PREFISSI NORMALIZZAZIONE DINAMICI:
   I prefissi da rimuovere (es. "fairing color - ", "subframe - ") vengono
   derivati automaticamente dai nomi delle colonne mappa, non hardcoded.

3. GESTIONE MULTI-MODELLO:
   Righe con Model='MSV4+MSV4S' vengono espanse in due entries separate
   prima del matching. La logica OR esistente le gestisce correttamente.

4. FILE V2:
   Legge da tutte_righe_univoche_V2.xlsx invece di tutte_righe_univoche.xlsx.

LOGICA INVARIATA: matching, aggregazione, statistiche, salvataggio identici a V1.
================================================================================

SCOPO:
Questo modulo esegue il matching tra i componenti (SKU) e le combinazioni di
configurazioni moto generate dalla simulazione Monte Carlo. Per ogni SKU calcola:
- Domanda aggregata nel periodo di lead time
- Safety stock necessario (percentili 95° e 99°)
- Statistiche sulla frequenza di utilizzo

FLUSSO PRINCIPALE:
1. Carica il catalogo SKU (tutte_righe_univoche.xlsx) con le caratteristiche
2. Carica i lead time dei componenti (quantity e residual.xlsx)
3. Riceve il DataFrame delle simulazioni (dal Monte Carlo)
4. Per ogni SKU:
   a. Trova tutte le configurazioni moto che richiedono quello SKU
   b. Somma la domanda di quelle configurazioni per i mesi del lead time
   c. Calcola statistiche sui run (media, std, percentili)
5. Salva i risultati con safety stock calcolati

CONCETTO CHIAVE - MATCHING:
Uno SKU è richiesto da una configurazione se TUTTE le caratteristiche dello SKU
corrispondono a quelle della configurazione. Esempio:
- SKU "Parafango Rosso MSV4S" matcha configurazioni con:
  - MODEL = MSV4S
  - FAIRING COLOR = Rosso
- Le caratteristiche "No" nello SKU sono ignorate (wildcard)

CONCETTO CHIAVE - LEAD TIME:
Se un componente ha lead time di 3 mesi, dobbiamo ordinarlo 3 mesi prima.
Quindi aggreghiamo la domanda dei prossimi 3 mesi per calcolare il safety stock.

OTTIMIZZAZIONI APPLICATE:
1. Cache normalizzazioni: lookup O(1) invece di ricalcolo stringa
2. Numpy arrays: matching vettoriale invece di loop Python
3. Categorical dtype: risparmio memoria per stringhe ripetute
4. Matrice int16: risparmio 75% memoria per dati numerici
5. Pre-calcolo indici: evita ricerca colonne ripetuta
6. Early stopping: interrompe matching se maschera è tutta False

LOGICA INVARIATA: L'output è identico alla versione originale

@author: AntonioLiardo-SMOPS
================================================================================
"""

import numpy as np
import pandas as pd
from typing import Dict, List, Tuple, Optional
import math
import time
from datetime import datetime
import sys
import gc
# import psutil  # optional, commented for portability
import os
import tempfile
from multiprocessing import Pool, cpu_count
from functools import partial
import warnings

warnings.filterwarnings('ignore')


# ============================================================================
# CALCOLO RESIDUAL LEAD TIME (NUOVO V3)
# ============================================================================
# V3: Sostituisce load_lead_times() con processo ricavato da Gates MTSV4
# e Dashboard Supplier. Zero dipendenza da "quantity e residual.xlsx"
# ============================================================================

def calcola_residual_lead_time(input_dir: str = '.\\Input',
                               logistica_dir: str = '.\\Input\\Dati Logistica') -> pd.DataFrame:
    """
    Calcola Residual Lead Time utilizzando:
    1. tutte_righe_univoche_V2.xlsx (configurazioni)
    2. Gates MTSV4.xlsx (lead time per gates)
    3. Dashboard Supplier (Frozen + Frequenza)
    4. Prezzi.xlsx (prezzi componenti)

    Processo:
    - Mapping colonne tutte_righe_univoche con Gates.Optional
    - Estrazione gate weeks per ogni configurazione
    - Calcolo MAX(gate_weeks) per riga
    - Merge con Dashboard + Prezzi per Frozen + Frequenza
    - Formula: Residual_LT = (FROZEN + 1/FREQ_mesi) - MAX_gate/4

    Returns:
        DataFrame con: Numero componenti (SKU), Residual_LT in mesi
    """
    from pathlib import Path
    import re

    input_path = Path(input_dir)
    logistica_path = Path(logistica_dir)

    print("\n[RESIDUAL_LT V3] Caricamento dati...")

    # 1. Carica tutte_righe_univoche_V2
    df_univoche = pd.read_excel(input_path / "tutte_righe_univoche_V2.xlsx")
    print(f"   tutte_righe_univoche: {df_univoche.shape}")

    # 2. Carica Gates
    df_gates = pd.read_excel(logistica_path / "Gates MTSV4.xlsx", skiprows=1)
    print(f"   Gates: {df_gates.shape}")

    GATE_COL = 'Gates (Weeks) rev'

    # 3a. Auto-detect colonne caratteristiche: Optional TRANSCOD (escluso BUNDLE) == colonna univoche
    gate_optionals = [c for c in df_gates['Optional TRANSCOD'].dropna().unique() if c != 'BUNDLE']
    config_cols = [c for c in gate_optionals if c in df_univoche.columns]
    print(f"   Colonne caratteristiche matchate: {config_cols}")

    # 3b. Auto-detect colonne bundle: Value TRANSCOD dei BUNDLE == colonna univoche
    # Con Supermodel filtering
    bundle_rows = df_gates[df_gates['Optional TRANSCOD'] == 'BUNDLE'].copy()
    bundle_gate_map = {}  # col_name -> gate_weeks (fallback)
    bundle_gate_map_sm = {}  # (col_name, supermodel) -> gate_weeks
    for _, row in bundle_rows.iterrows():
        val = str(row['Value TRANSCOD']).strip() if pd.notna(row['Value TRANSCOD']) else None
        supermodel = row.get('Supermodel', None) if 'Supermodel' in df_gates.columns else None
        gw  = row[GATE_COL]
        if val and val in df_univoche.columns and pd.notna(gw):
            # Fallback (senza supermodel)
            if val not in bundle_gate_map:
                bundle_gate_map[val] = gw
            # Con supermodel
            if pd.notna(supermodel):
                supermodel_str = str(supermodel).strip().lower()
                key = (val, supermodel_str)
                bundle_gate_map_sm[key] = gw
    bundle_cols = list(bundle_gate_map.keys())
    print(f"   Colonne bundle matchate: {bundle_cols}")

    all_cols = config_cols + bundle_cols
    df_output = df_univoche[['Numero componenti'] + all_cols].copy()

    # 3c. Aggiungi colonna Versione per matching Supermodel
    versione_col = 'Versione' if 'Versione' in df_univoche.columns else None
    if versione_col:
        df_output['_versione'] = df_univoche[versione_col]
        print(f"   Colonna Versione rilevata per Supermodel matching")

    # 4. Converti a object type per mixed values
    for col in all_cols:
        df_output[col] = df_output[col].astype('object')

    # 5. Crea gate mapping per colonne caratteristiche
    # Mapping con Supermodel: (value, supermodel) -> gate_weeks
    # Fallback senza Supermodel: value -> gate_weeks
    gates_by_col = {}
    gates_by_col_with_supermodel = {}
    for col in config_cols:
        sub_df = df_gates[df_gates['Optional TRANSCOD'] == col]
        mapping = {}  # fallback: valore senza supermodel
        mapping_with_sm = {}  # con supermodel: (valore, supermodel) -> gate
        for _, row in sub_df.iterrows():
            value = row['Value TRANSCOD']
            supermodel = row.get('Supermodel', None) if 'Supermodel' in df_gates.columns else None
            gw = row[GATE_COL]
            if pd.notna(value) and pd.notna(gw):
                value_lower = str(value).strip().lower()
                # Mapping fallback (senza supermodel)
                if value_lower not in mapping:
                    mapping[value_lower] = gw
                # Mapping con supermodel
                if pd.notna(supermodel):
                    supermodel_str = str(supermodel).strip().lower()
                    key = (value_lower, supermodel_str)
                    mapping_with_sm[key] = gw
        gates_by_col[col] = mapping
        gates_by_col_with_supermodel[col] = mapping_with_sm
    print(f"   Gate mapping creato con Supermodel filtering")

    # 6. Processa celle caratteristiche - sostituisci valori con gate weeks
    # Matching con Supermodel + fallback senza Supermodel
    print("[RESIDUAL_LT V3] Processing celle caratteristiche (con Supermodel matching)...")
    replaced = 0
    replaced_with_sm = 0
    for col in config_cols:
        col_gate_map = gates_by_col[col]
        col_gate_map_sm = gates_by_col_with_supermodel[col]
        for idx in df_output.index:
            cell_value = df_output.loc[idx, col]
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip().lower()
                if cell_str not in ('no', ''):
                    # Try matching con Supermodel first
                    versione = df_output.loc[idx, '_versione'] if versione_col else None
                    gw_found = None
                    if versione and pd.notna(versione):
                        versione_lower = str(versione).strip().lower()
                        key = (cell_str, versione_lower)
                        if key in col_gate_map_sm:
                            gw_found = col_gate_map_sm[key]
                            replaced_with_sm += 1
                    # Fallback: matching senza Supermodel
                    if gw_found is None and cell_str in col_gate_map:
                        gw_found = col_gate_map[cell_str]
                    # Sostituisci se trovato
                    if gw_found is not None:
                        df_output.loc[idx, col] = gw_found
                        replaced += 1
    print(f"   Celle caratteristiche replicate: {replaced} (con Supermodel: {replaced_with_sm})")

    # 6b. Processa bundle - qualsiasi valore != "Not"/"No" → metti gate weeks
    # Con Supermodel matching + fallback
    _NO_VALUES = {'no', 'nan', '', 'none', 'n/a', 'na'}
    replaced_bundle = 0
    replaced_bundle_sm = 0
    for col in bundle_cols:
        gw_fallback = bundle_gate_map[col]
        for idx in df_output.index:
            cell_value = df_output.loc[idx, col]
            if pd.notna(cell_value) and str(cell_value).strip().lower() not in _NO_VALUES:
                # Try matching con Supermodel first
                versione = df_output.loc[idx, '_versione'] if versione_col else None
                gw_found = None
                if versione and pd.notna(versione):
                    versione_lower = str(versione).strip().lower()
                    key = (col, versione_lower)
                    if key in bundle_gate_map_sm:
                        gw_found = bundle_gate_map_sm[key]
                        replaced_bundle_sm += 1
                # Fallback: usa gate senza supermodel
                if gw_found is None:
                    gw_found = gw_fallback
                df_output.loc[idx, col] = gw_found
                replaced_bundle += 1
            else:
                df_output.loc[idx, col] = None
    print(f"   Celle bundle replicate: {replaced_bundle} (con Supermodel: {replaced_bundle_sm})")

    # 6c. Rimuovi colonna temporanea versione
    if '_versione' in df_output.columns:
        df_output = df_output.drop(columns=['_versione'])

    # 7. Calcola MAX_gate per ogni riga (su tutte le colonne)
    print("[RESIDUAL_LT V3] Calcola MAX_gate...")
    max_gate_values = []
    for idx in df_output.index:
        row_values = []
        for col in all_cols:
            val = df_output.loc[idx, col]
            if pd.notna(val) and str(val).strip().lower() not in ('no', ''):
                try:
                    row_values.append(float(val))
                except (ValueError, TypeError):
                    pass
        max_gate_values.append(max(row_values) if row_values else None)

    df_output['MAX_gate'] = max_gate_values

    # 7b. Flag: almeno un valore non-No nei config/bundle (per diagnostica NO_GATE)
    _NO_CHECK = {'no', 'nan', '', 'none', 'n/a', 'na'}
    def _has_non_no_value(row):
        for col in all_cols:
            v = row[col]
            if pd.notna(v) and str(v).strip().lower() not in _NO_CHECK:
                return True
        return False
    df_output['_has_non_no'] = df_output.apply(_has_non_no_value, axis=1)

    # Aggrega: True se almeno una riga del SKU ha valore non-No
    has_non_no_per_sku = (df_output.groupby('Numero componenti')['_has_non_no']
                          .any()
                          .reset_index()
                          .rename(columns={'_has_non_no': 'has_non_no'}))

    # 8. Carica Dashboard + Prezzi
    print("[RESIDUAL_LT V3] Carica Dashboard + Prezzi...")
    prezzi_files = sorted(logistica_path.glob("PREZZI*.XLS*"))
    if not prezzi_files:
        prezzi_files = sorted(input_path.glob("PREZZI*.XLS*"))
    if not prezzi_files:
        raise FileNotFoundError(
            f"Nessun file PREZZI trovato in: {logistica_path} o {input_path}"
        )
    prezzi_path_found = prezzi_files[-1]
    print(f"   Prezzi: {prezzi_path_found.name}")
    df_prezzi = pd.read_excel(prezzi_path_found, sheet_name='Sheet1', skiprows=0)

    dash_files = sorted(logistica_path.glob("*Dashboard Supplier*.xlsx"))
    if not dash_files:
        raise FileNotFoundError(
            f"Nessun file 'Dashboard Supplier' trovato in: {logistica_path}\n"
            f"Atteso pattern: *Dashboard Supplier*.xlsx"
        )
    dash_path = dash_files[-1]
    print(f"   Dashboard: {dash_path.name}")
    df_dash = pd.read_excel(dash_path, sheet_name='Foglio1')
    df_dash_filt = df_dash[df_dash['V/M'].isin(['V', 'M/V'])].copy()

    # 9. Estrai ID supplier — prende primo token (prima dello spazio)
    # Gestisce sia format numerico ('12345 SUPPLIER') che alfanumerico ('FF15352 SUPPLIER')
    # Auto-detect colonna fornitore (nomi diversi tra Sheet1 e altri fogli)
    vendor_col_candidates = [
        'Fornitore/divisione fornitrice',
        "Vendor/supplying plant COD PIU' RECENTE",
        'Fornitore',
    ]
    vendor_col = next((c for c in vendor_col_candidates if c in df_prezzi.columns), None)
    if vendor_col is None:
        raise ValueError(f"Colonna fornitore non trovata in PREZZI. Colonne: {df_prezzi.columns.tolist()}")
    print(f"   Colonna fornitore PREZZI: '{vendor_col}'")

    def extract_id(text):
        if pd.isna(text):
            return None
        token = str(text).strip().split()[0] if str(text).strip() else None
        if not token:
            return None
        return token

    df_prezzi['ID_Supplier_extracted'] = df_prezzi[vendor_col].apply(extract_id)

    # Normalizza ID Supplier Dashboard a stringa per confronto uniforme
    df_dash_filt = df_dash_filt.copy()
    df_dash_filt['ID Supplier'] = df_dash_filt['ID Supplier'].apply(
        lambda x: str(int(x)) if pd.notna(x) and str(x).strip().lstrip('-').isdigit() else str(x).strip() if pd.notna(x) else None
    )

    # 10. Merge Prezzi + Dashboard
    # Colonna prezzo — nomi diversi tra fogli
    price_col_candidates = ['Prezzo netto', "Net price for unit (€) PIU' RECENTE", 'Prezzo', 'Price']
    price_col = next((c for c in price_col_candidates if c in df_prezzi.columns), None)

    prezzi_cols = ['Materiale', 'ID_Supplier_extracted']
    if 'Testo breve' in df_prezzi.columns:
        prezzi_cols.append('Testo breve')
    if price_col:
        prezzi_cols.append(price_col)
        print(f"   Colonna prezzo: '{price_col}'")
    df_prezzi_dash = pd.merge(
        df_prezzi[prezzi_cols],
        df_dash_filt[['ID Supplier', 'Frequency[WK]', 'FROZEN[Months] ']],
        left_on='ID_Supplier_extracted',
        right_on='ID Supplier',
        how='left'
    )

    # 11. Calcola Freq_mesi
    df_prezzi_dash['Freq_mesi'] = df_prezzi_dash['Frequency[WK]'] / 4.33

    # 12. Aggrega MAX_gate per SKU (MAX tra tutte le configurazioni del componente)
    max_gate_per_sku = (df_output[['Numero componenti', 'MAX_gate']]
                        .groupby('Numero componenti', as_index=False)['MAX_gate']
                        .max())
    # Merge flag has_non_no per diagnostica
    max_gate_per_sku = max_gate_per_sku.merge(has_non_no_per_sku, on='Numero componenti', how='left')

    # 12b. Merge con prezzi/dashboard
    dash_cols = ['Materiale', 'ID Supplier', 'FROZEN[Months] ', 'Freq_mesi']
    if 'Testo breve' in df_prezzi_dash.columns:
        dash_cols.append('Testo breve')
    if price_col and price_col in df_prezzi_dash.columns:
        dash_cols.append(price_col)
    df_final = pd.merge(
        max_gate_per_sku,
        df_prezzi_dash[dash_cols],
        left_on='Numero componenti',
        right_on='Materiale',
        how='left'
    )

    # 12c. Default gate per NO_GATE_ALL_NO: assegna 4 mesi (= 4 * 4.33 settimane)
    DEFAULT_GATE_WK = 4 * 4.33
    mask_default = df_final['MAX_gate'].isna() & ~df_final['has_non_no'].fillna(False)
    df_final['default_gate_applied'] = mask_default
    df_final.loc[mask_default, 'MAX_gate'] = DEFAULT_GATE_WK
    n_default = int(mask_default.sum())
    print(f"   [RESIDUAL_LT] Default gate 4 mesi applicato a {n_default} SKU (NO_GATE_ALL_NO)")

    # 13. Calcola Residual_LT (clampato a 0 se negativo)
    df_final['MAX_gate_mesi'] = df_final['MAX_gate'] / 4.33
    df_final['Supply_cycle'] = df_final['FROZEN[Months] '] + df_final['Freq_mesi']
    df_final['Residual_LT'] = (df_final['Supply_cycle'] - df_final['MAX_gate_mesi']).clip(lower=0)

    # 14. Aggiungi colonna diagnostica match_status
    def _match_status(row):
        if pd.isna(row.get('MAX_gate')):
            # MAX_gate NaN qui significa NO_GATE_MISSING (tutti-No già gestiti sopra con default)
            if row.get('has_non_no', False):
                return 'NO_GATE_MISSING'
            else:
                return 'NO_GATE_ALL_NO'   # fallback safety (non dovrebbe accadere)
        if pd.isna(row.get('ID Supplier')):
            return 'NO_DASHBOARD_MATCH'
        if pd.isna(row.get('FROZEN[Months] ')):
            return 'NO_FROZEN'
        if pd.isna(row.get('Freq_mesi')):
            return 'NO_FREQUENCY'
        if row.get('default_gate_applied', False):
            return 'OK_DEFAULT_GATE'      # gate di default 4 mesi applicato, residual calcolato
        if row.get('Residual_LT', 0) == 0:
            return 'OK_ZERO'
        return 'OK'
    df_final['match_status'] = df_final.apply(_match_status, axis=1)

    # 14b. Salva debug completo in Dati Logistica
    try:
        debug_cols = ['Numero componenti', 'MAX_gate', 'MAX_gate_mesi',
                      'default_gate_applied',
                      'ID Supplier', 'FROZEN[Months] ', 'Frequency[WK]',
                      'Freq_mesi', 'Supply_cycle', 'Residual_LT', 'match_status']
        if 'Testo breve' in df_final.columns:
            debug_cols.insert(1, 'Testo breve')
        if price_col and price_col in df_final.columns:
            debug_cols.insert(2, price_col)
        # Aggiungi Frequency[WK] da df_prezzi_dash se non già in df_final
        if 'Frequency[WK]' not in df_final.columns and 'Frequency[WK]' in df_prezzi_dash.columns:
            df_final = df_final.merge(
                df_prezzi_dash[['Materiale', 'Frequency[WK]']].drop_duplicates('Materiale'),
                left_on='Numero componenti', right_on='Materiale', how='left', suffixes=('', '_dup')
            )
        debug_cols_exist = [c for c in debug_cols if c in df_final.columns]
        df_debug = df_final[debug_cols_exist].copy()
        df_debug.columns = [c.strip() for c in df_debug.columns]
        debug_path = logistica_path / 'residual_lead_time_debug.xlsx'
        with pd.ExcelWriter(debug_path, engine='openpyxl') as writer:
            df_debug.to_excel(writer, sheet_name='debug_completo', index=False)
            # Foglio codici problematici
            df_problematici = df_debug[
                df_debug['match_status'].isin(['NO_DASHBOARD_MATCH', 'NO_GATE_MISSING'])
            ].copy()
            df_problematici.to_excel(writer, sheet_name='codici_problematici', index=False)
        print(f"   [DEBUG] File completo salvato: {debug_path}")
        print(f"   [DEBUG] match_status breakdown:")
        for status, cnt in df_debug.get('match_status', pd.Series(dtype=str)).value_counts().items():
            print(f"      {status}: {cnt}")
        print(f"   [DEBUG] Codici problematici: {len(df_problematici)}")
    except Exception as _e:
        print(f"   [WARN] Impossibile salvare debug: {_e}")
        import traceback; traceback.print_exc()

    # 15. Restituisci colonne essenziali
    out_cols = ['Numero componenti', 'Residual_LT', 'ID Supplier']
    out_rename = ['SKU', 'Lead_Time_Months', 'ID_Supplier_ref']
    if 'Testo breve' in df_final.columns:
        out_cols.append('Testo breve')
        out_rename.append('Description')
    result = df_final[out_cols].copy()
    result.columns = out_rename

    print(f"   [OK] Residual_LT calcolato: {result['Lead_Time_Months'].notna().sum()} SKU con valore")
    return result


# ============================================================================
# AUTO-RILEVAMENTO MAPPING COLONNE (ZERO HARDCODING)
# ============================================================================
# V2 non usa un dizionario COLUMN_MAPPING fisso. Il mapping viene costruito
# dinamicamente da build_column_mapping_auto() dopo aver caricato i dati.
# ============================================================================

def build_column_mapping_auto(sku_cols: List[str],
                               sim_cols: List[str]) -> Dict[str, str]:
    """
    Costruisce automaticamente il mapping colonne SKU -> simulazione.

    LOGICA:
    -------
    Per ogni colonna del catalogo SKU (tutte_righe_univoche_V2), cerca la
    corrispondente colonna nella simulazione tramite matching case-insensitive.

    COLONNE ESCLUSE:
    - 'id'               -> identificatore progressivo, non è una caratteristica
    - 'numero componenti' -> codice SKU, non è una caratteristica di configurazione

    NESSUNA ECCEZIONE HARDCODED:
    Il nome 'Model' in SKU matcha 'MODEL' in simulazione perché la ricerca è
    case-insensitive. Funziona per qualsiasi variazione maiuscolo/minuscolo.

    Args:
        sku_cols: Colonne di tutte_righe_univoche_V2 (catalogo SKU)
        sim_cols: Colonne del DataFrame simulazione (output MC)

    Returns:
        Dict {sku_col: sim_col} per il matching
    """
    # Colonne da escludere (metadati, non caratteristiche di configurazione)
    EXCLUDE = {'id', 'numero componenti', 'numero componenti.1',
               'numero componenti.2', 'numero componenti.3'}

    # Lookup case-insensitive: nome_lower -> nome_originale per sim_cols
    sim_lower_map = {c.lower().strip(): c for c in sim_cols}

    mapping = {}
    for sku_col in sku_cols:
        norm = sku_col.lower().strip()
        if norm in EXCLUDE:
            continue
        sim_match = sim_lower_map.get(norm)
        if sim_match:
            mapping[sku_col] = sim_match

    return mapping


# ============================================================================
# UTILITY - MONITORAGGIO MEMORIA
# ============================================================================

def get_memory_usage_mb():
    """
    Restituisce la memoria RAM usata dal processo corrente in MB.
    Utile per monitorare l'impatto delle ottimizzazioni sulla memoria.
    """
    try:
        import psutil
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024
    except ImportError:
        return None


# ============================================================================
# NORMALIZZAZIONE STRINGHE - SISTEMA DI CACHE
# ============================================================================
# La normalizzazione converte stringhe in un formato standard per il confronto.
# Esempio: "Bundle - Fairing Color - Rosso Ducati" -> "rosso ducati"
#
# Il problema: la stessa stringa viene normalizzata migliaia di volte.
# La soluzione: cache con lookup O(1) invece di ricalcolo O(n).
# ============================================================================

class NormalizationCache:
    """
    Cache per la normalizzazione delle stringhe.

    Problema originale:
    - normalize_value() chiamata ~1 milione di volte
    - Ogni chiamata fa operazioni su stringhe (lowercase, replace, etc.)

    Soluzione:
    - Calcola la normalizzazione UNA volta per valore unico
    - Le chiamate successive fanno solo lookup in dizionario O(1)

    Risparmio stimato: ~90% del tempo di normalizzazione
    """

    def __init__(self, prefixes: List[str] = None):
        # Dizionario cache: (valore_originale, nome_colonna) -> valore_normalizzato
        self.cache = {}

        # Prefissi da rimuovere durante la normalizzazione.
        # V2: derivati dinamicamente dai nomi delle colonne mappate (zero hardcoding).
        # Es: colonna 'FAIRING COLOR' -> prefisso 'fairing color - '
        #     colonna 'SUBFRAME'      -> prefisso 'subframe - '
        # Il TR genera valori come "FAIRING COLOR - Ducati Red gloss";
        # rimuovendo il prefisso si ottiene "Ducati Red gloss" per il matching.
        self.prefixes = prefixes if prefixes is not None else []

    def normalize(self, value, column_name: str = None) -> str:
        """
        Normalizza un valore con caching.

        Il processo di normalizzazione:
        1. Converte in minuscolo
        2. Rimuove prefissi noti (es. "bundle - ")
        3. Normalizza spazi e trattini
        4. Gestisce casi speciali (es. "spoked-black" -> "spoked black")

        Args:
            value: Valore da normalizzare (può essere qualsiasi tipo)
            column_name: Nome colonna (per casi specifici di normalizzazione)

        Returns:
            Stringa normalizzata. "no" se il valore è None/NaN.
        """
        # Chiave cache: include column_name perché la stessa stringa
        # potrebbe normalizzarsi diversamente in colonne diverse
        cache_key = (value, column_name)

        # Cache hit: restituisci risultato già calcolato
        if cache_key in self.cache:
            return self.cache[cache_key]

        # Cache miss: calcola normalizzazione
        if pd.isna(value):
            result = "no"
        else:
            value_str = str(value).strip().lower()

            # "Nessuno" è equivalente a "no" (opzione non selezionata)
            if value_str == "nessuno":
                result = "no"
            else:
                # Rimuovi prefissi noti
                for prefix in self.prefixes:
                    if value_str.startswith(prefix):
                        value_str = value_str[len(prefix):]
                        break

                # Rimuovi prefisso basato sul nome colonna
                # Es: "FAIRING COLOR - Rosso" -> "Rosso"
                if column_name:
                    prefix = column_name.lower() + " - "
                    if value_str.startswith(prefix):
                        value_str = value_str[len(prefix):]

                # Normalizza spazi e trattini per consistenza
                value_str = value_str.replace(" - ", "-")
                value_str = value_str.replace(" (", "(").replace("( ", "(")
                value_str = value_str.replace(") ", ")").replace(" )", ")")

                # Caso speciale: "spoked-black" vs "spoked black"
                value_str = value_str.replace("spoked-black", "spoked black")

                result = value_str

        # Salva in cache per usi futuri
        self.cache[cache_key] = result
        return result

    def preload_column(self, series: pd.Series, column_name: str = None):
        """
        Pre-carica tutti i valori unici di una colonna nella cache.

        Chiamato all'inizio per "scaldare" la cache con tutti i valori
        che incontreremo durante il matching.

        Args:
            series: Colonna pandas da pre-caricare
            column_name: Nome colonna per normalizzazione
        """
        unique_values = series.dropna().unique()
        for val in unique_values:
            self.normalize(val, column_name)


# Istanza globale della cache (condivisa da tutto il modulo).
# V2: viene re-inizializzata in main_sku_v2() con i prefissi derivati
# dinamicamente dal column_mapping (vedere _reinit_norm_cache()).
_norm_cache = NormalizationCache()


def normalize_value_cached(value, column_name: str = None) -> str:
    """
    Wrapper pubblico per la normalizzazione con cache.
    Usa l'istanza globale _norm_cache.
    """
    return _norm_cache.normalize(value, column_name)


def _reinit_norm_cache(column_mapping: Dict[str, str]) -> None:
    """
    Re-inizializza _norm_cache con i prefissi derivati dal column_mapping.

    V2: i prefissi NON sono hardcoded ma derivati dai nomi delle colonne
    simulazione presenti nel mapping. Questo garantisce che qualsiasi
    nuova colonna aggiunta al TR venga normalizzata correttamente
    senza modifiche al codice.

    Esempio: se column_mapping ha valori ['FAIRING COLOR', 'SUBFRAME', ...]
    -> prefissi = ['fairing color - ', 'subframe - ', ...]
    -> 'FAIRING COLOR - Ducati Red gloss' -> 'Ducati Red gloss'

    Chiamare DOPO aver costruito column_mapping e PRIMA di pre-normalizzare
    il simulation output.
    """
    global _norm_cache
    # Deriva prefissi da tutti i nomi colonna-simulazione nel mapping
    # Aggiunge anche "bundle - " per retrocompatibilità con eventuali valori TR
    prefixes = ['bundle - '] + [v.lower().strip() + ' - '
                                 for v in column_mapping.values()]
    _norm_cache = NormalizationCache(prefixes=prefixes)
    print(f"[NormCache V2] Re-inizializzata con {len(prefixes)} prefissi dinamici")


# ============================================================================
# CARICAMENTO DATI - LEAD TIME
# ============================================================================

def load_lead_times(filepath: str = '.\\Input\\quantity e residual.xlsx') -> pd.DataFrame:
    """
    Carica i lead time dei componenti - VERSIONE V3 (NEW - Residual LT da Gates)

    NEW (V3): Sostituisce il caricamento da "quantity e residual.xlsx" con il processo
    calcola_residual_lead_time() che usa Gates MTSV4 + Dashboard Supplier + Prezzi.

    Per backward-compatibility, il parametro filepath è mantenuto ma ignorato.

    FLUSSO:
    1. Calcola Residual Lead Time via calcola_residual_lead_time()
    2. Carica BOM per le quantità per componente (Qtà comp. UMC)
    3. Merge su SKU
    4. Restituisce formato identico a V1/V2 per compatibilità

    Returns:
        DataFrame con colonne: SKU, Lead_Time_Months, Lead_Time_Months_Ceil,
        Quantity_Per_Bike, Description
    """
    from pathlib import Path

    print(f"\n{'='*60}")
    print(f"CARICAMENTO LEAD TIME - V3 (Residual LT da Gates)")
    print(f"{'='*60}")

    input_dir = Path('.') / 'Input'
    logistica_dir = input_dir / 'Dati Logistica'

    # Trova BOM_150 in qualsiasi sottocartella MODEL/*/
    def _find_bom150(base: Path) -> Path:
        model_root = base / 'MODEL'
        if model_root.exists():
            for sub in sorted(model_root.iterdir()):
                if sub.is_dir():
                    for f in sub.iterdir():
                        if f.name.lower().startswith('bom_150') and f.suffix.lower() in ('.xlsx', '.xls'):
                            return f
        raise FileNotFoundError(
            f"Nessun file BOM_150*.xlsx trovato in {model_root}. "
            f"Verifica che esista almeno una sottocartella MODEL/<versione>/BOM_150*.xlsx"
        )

    # 1. Calcola Residual_LT via nuova funzione
    lead_times_residual = calcola_residual_lead_time(str(input_dir), str(logistica_dir))
    print(f"   Residual LT calcolato: {len(lead_times_residual)} SKU")

    # (debug salvato dentro calcola_residual_lead_time)

    # 2. Carica BOM per le quantità
    print(f"\n   Caricamento BOM per quantità...")
    bom_path = _find_bom150(input_dir)
    print(f"   BOM trovato: {bom_path}")
    df_bom = pd.read_excel(bom_path, sheet_name=0)

    # Estrai SKU + Qty da BOM
    bom_qty = df_bom[['Numero componenti', 'Qtà comp. (UMC)']].copy()
    bom_qty.columns = ['SKU', 'Quantity_Per_Bike']

    # Aggrega per SKU (stesso SKU può apparire più volte con qty diverse)
    # Prendi max qty
    bom_qty_unique = bom_qty.groupby('SKU', as_index=False).agg({
        'Quantity_Per_Bike': 'max'
    })
    print(f"   BOM SKU univoci: {len(bom_qty_unique)}")

    # 3. Merge Residual_LT + BOM Qty
    lead_times = pd.merge(
        lead_times_residual,
        bom_qty_unique,
        on='SKU',
        how='left'
    )

    # 4. Default: Quantity_Per_Bike = 1.0 se non in BOM
    lead_times['Quantity_Per_Bike'] = lead_times['Quantity_Per_Bike'].fillna(1.0)

    # 5. Crea Lead_Time_Months_Ceil (stesso valore di Lead_Time_Months per decimali)
    lead_times['Lead_Time_Months_Ceil'] = lead_times['Lead_Time_Months'].copy()

    # 6. Aggiungi Description da Testo breve (Prezzi), fallback a 'Residual LT (da Gates)'
    if 'Description' in lead_times_residual.columns:
        desc_map = (lead_times_residual[['SKU', 'Description']]
                    .drop_duplicates(subset='SKU', keep='first')
                    .set_index('SKU')['Description'])
        lead_times['Description'] = lead_times['SKU'].map(desc_map).fillna('Residual LT (da Gates)')

    # 7. Filtra: solo lead time non-null
    lead_times = lead_times.dropna(subset=['Lead_Time_Months'])

    # 8. Riordina colonne
    lead_times = lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                             'Quantity_Per_Bike', 'Description']]

    print(f"\n   [OK] Lead times caricati: {len(lead_times)} SKU")
    print(f"   Non-null Lead_Time_Months: {lead_times['Lead_Time_Months'].notna().sum()}")

    # 9. Aggiungi foglio 'input_simulatore' al debug file
    try:
        debug_path = logistica_dir / 'residual_lead_time_debug.xlsx'
        if debug_path.exists():
            with pd.ExcelWriter(debug_path, engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                lead_times.to_excel(writer, sheet_name='input_simulatore', index=False)
            print(f"   [DEBUG] Foglio 'input_simulatore' aggiunto: {debug_path}")
    except Exception as _e:
        print(f"   [WARN] Impossibile aggiungere foglio simulatore: {_e}")

    return lead_times


# ============================================================================
# PRE-NORMALIZZAZIONE - PREPARAZIONE DATI SIMULAZIONE
# ============================================================================

def prenormalize_simulation_output_fast(simulation_output: pd.DataFrame,
                                        column_mapping: Dict[str, str]) -> Tuple[pd.DataFrame, Dict[str, np.ndarray]]:
    """
    Pre-normalizza tutte le colonne rilevanti del DataFrame simulazione.

    Questa funzione prepara i dati per il matching veloce:
    1. Normalizza tutte le stringhe nelle colonne di configurazione
    2. Converte in categorical per risparmiare memoria
    3. Estrae array numpy per matching vettoriale

    OTTIMIZZAZIONI:
    - Usa map con dizionario invece di apply (più veloce)
    - Converte in categorical (risparmio memoria ~50%)
    - Pre-calcola array numpy di codici categoria

    Args:
        simulation_output: DataFrame dal Monte Carlo
        column_mapping: Mapping colonne SKU -> simulazione

    Returns:
        Tupla (DataFrame modificato, dizionario di array numpy per matching)
    """
    print(f"\n[PRE-NORM] Inizio pre-normalizzazione veloce...")
    sys.stdout.flush()
    start = time.time()

    # Modifica in-place per risparmiare memoria (non copiare)
    normalized_df = simulation_output
    norm_arrays = {}

    # Per ogni colonna nel mapping
    for i, (sku_col, combo_col) in enumerate(column_mapping.items(), 1):
        if combo_col in normalized_df.columns:

            # Pre-carica cache con tutti i valori unici
            _norm_cache.preload_column(normalized_df[combo_col], combo_col)

            # Nome colonna normalizzata (aggiunge suffisso _NORM)
            norm_col_name = f"{combo_col}_NORM"

            # OTTIMIZZAZIONE: crea dizionario di mapping e usa .map()
            # Molto più veloce di .apply() perché evita chiamate Python
            unique_vals = normalized_df[combo_col].dropna().unique()
            norm_map = {v: normalize_value_cached(v, combo_col) for v in unique_vals}
            norm_map[np.nan] = "no"

            # Applica normalizzazione con map
            normalized_df[norm_col_name] = normalized_df[combo_col].map(
                lambda x: norm_map.get(x, "no") if pd.notna(x) else "no"
            )

            # OTTIMIZZAZIONE: converti in categorical
            # Risparmia memoria perché le stringhe ripetute sono memorizzate una volta
            normalized_df[norm_col_name] = normalized_df[norm_col_name].astype('category')

            # OTTIMIZZAZIONE: estrai array di codici categoria per matching numpy
            # I codici sono interi (0, 1, 2...) molto più veloci da confrontare
            norm_arrays[norm_col_name] = normalized_df[norm_col_name].cat.codes.values

    elapsed = time.time() - start
    print(f"[PRE-NORM] Completato in {elapsed:.2f}s")
    sys.stdout.flush()

    return normalized_df, norm_arrays


# ============================================================================
# MATCHING SKU-COMBINAZIONI - VERSIONE NUMPY VETTORIALE
# ============================================================================

def sku_matches_combination_numpy(sku_row_values: Dict[str, str],
                                  norm_arrays: Dict[str, np.ndarray],
                                  category_maps: Dict[str, Dict[str, int]],
                                  column_mapping: Dict[str, str],
                                  n_rows: int) -> np.ndarray:
    """
    Determina quali configurazioni matchano uno SKU usando numpy vettoriale.

    LOGICA DI MATCHING:
    Uno SKU matcha una configurazione se TUTTE le caratteristiche rilevanti
    dello SKU corrispondono. Le caratteristiche con valore "no" sono ignorate
    (wildcard - lo SKU può essere usato con qualsiasi valore).

    Esempio:
    SKU con Model="MSV4S", Fairing_colors="Rosso", Rims="No"
    -> Matcha tutte le configurazioni MSV4S rosse, indipendentemente dai cerchi

    OTTIMIZZAZIONE:
    - Usa array numpy di codici categoria invece di confronti stringa
    - Early stopping: se la maschera diventa tutta False, esce subito
    - Operazioni vettoriali invece di loop Python

    Args:
        sku_row_values: Dizionario {nome_colonna_sku: valore}
        norm_arrays: Dizionario {nome_colonna_norm: array numpy codici}
        category_maps: Dizionario {colonna: {valore: codice}}
        column_mapping: Mapping colonne SKU -> simulazione
        n_rows: Numero righe nel DataFrame simulazione

    Returns:
        Array booleano: True dove la configurazione matcha lo SKU
    """
    # Inizia con maschera tutta True (tutte le configurazioni matchano)
    final_mask = np.ones(n_rows, dtype=bool)

    # Per ogni colonna nel mapping
    for sku_col, combo_col in column_mapping.items():
        # Salta se lo SKU non ha questa colonna
        if sku_col not in sku_row_values:
            continue

        # Nome colonna normalizzata
        norm_col = f"{combo_col}_NORM"
        if norm_col not in norm_arrays:
            continue

        # Normalizza il valore dello SKU
        sku_value = normalize_value_cached(sku_row_values[sku_col], sku_col)

        # WILDCARD: "no" significa "qualsiasi valore va bene"
        # Non aggiunge vincoli, passa alla prossima colonna
        if sku_value == "no":
            continue

        # Trova il codice categoria corrispondente al valore SKU
        cat_map = category_maps.get(norm_col, {})

        # WILDCARD "Yes": pack presente, matcha qualsiasi valore != "no"/"not"
        # Usato per righe pack-invariant in tutte_righe_univoche (es. colore panniers
        # = "Yes" significa "qualsiasi colore pack attivo va bene").
        if sku_value == "yes":
            no_code = cat_map.get("no", None)
            not_code = cat_map.get("not", None)
            if no_code is not None:
                final_mask &= norm_arrays[norm_col] != no_code
            if not_code is not None:
                final_mask &= norm_arrays[norm_col] != not_code
            if not final_mask.any():
                return final_mask
            continue

        if sku_value not in cat_map:
            # Valore non trovato -> nessuna configurazione matcha
            final_mask[:] = False
            return final_mask

        target_code = cat_map[sku_value]

        # MATCHING VETTORIALE: confronta array numpy (molto veloce)
        char_mask = norm_arrays[norm_col] == target_code

        # Combina con maschera precedente (AND logico)
        final_mask &= char_mask

        # EARLY STOPPING: se tutto False, inutile continuare
        if not final_mask.any():
            return final_mask

    return final_mask


# ============================================================================
# AGGREGAZIONE DOMANDE - SOMMA PER LEAD TIME
# ============================================================================

def precalculate_month_columns_fast(month_prefixes: List[str],
                                    n_runs: int,
                                    columns: List[str]) -> Dict[int, List[int]]:
    """
    Pre-calcola gli INDICI delle colonne per ogni mese.

    Le colonne run hanno nomi come "Jan 2026_Run_0", "Jan 2026_Run_1", etc.
    Invece di cercare le colonne per nome ogni volta, pre-calcoliamo gli indici.

    OTTIMIZZAZIONE: Accesso per indice O(1) invece di ricerca per nome O(n)

    Args:
        month_prefixes: Lista prefissi mese (es. ["Jan 2026", "Feb 2026"])
        n_runs: Numero run per mese
        columns: Lista tutte le colonne del DataFrame

    Returns:
        Dizionario {indice_mese: [lista indici colonne]}
    """
    columns_set = set(columns)
    columns_list = list(columns)

    # Mappa nome colonna -> indice
    col_to_idx = {col: idx for idx, col in enumerate(columns_list)}

    month_columns_cache = {}

    for month_idx, month_prefix in enumerate(month_prefixes):
        indices = []
        for i in range(n_runs):
            col_name = f"{month_prefix}_Run_{i}"
            if col_name in columns_set:
                indices.append(col_to_idx[col_name])
        month_columns_cache[month_idx] = indices

    return month_columns_cache


def aggregate_monthly_demands_numpy(matching_indices: np.ndarray,
                                    data_matrix: np.ndarray,
                                    month_col_indices: Dict[int, List[int]],
                                    lead_time_months: float,
                                    n_runs: int) -> np.ndarray:
    """
    Aggrega la domanda delle configurazioni matching per i mesi del lead time.
    Supporta INTERPOLAZIONE LINEARE per lead time decimali.

    LOGICA:
    Se lead_time = 2.3 mesi:
    - Somma domanda Mese1 + Mese2 (mesi interi)
    - Aggiungi 0.3 × domanda Mese3 (frazione del mese parziale)

    Questo evita l'arrotondamento per eccesso (ceil) e mantiene accuratezza.

    OTTIMIZZAZIONE:
    - Usa fancy indexing numpy (np.ix_) per selezionare sottomatrice
    - Somma vettoriale invece di loop
    - Lavora su matrice int16 pre-estratta

    Args:
        matching_indices: Indici righe che matchano (configurazioni che usano SKU)
        data_matrix: Matrice numpy dati (righe=config, colonne=run per mese)
        month_col_indices: Indici colonne pre-calcolati per mese
        lead_time_months: Quanti mesi aggregare (DECIMALE, es. 2.3)
        n_runs: Numero run

    Returns:
        Array con domanda aggregata per ogni run (float32 per supportare frazioni)
    """
    # Nessun match -> domanda zero
    if len(matching_indices) == 0:
        return np.zeros(n_runs, dtype=np.float32)

    # Accumula domanda per tutti i run (usa float32 per frazioni)
    aggregated = np.zeros(n_runs, dtype=np.float32)

    # Estrai parte intera e frazione del lead time
    lead_time_full = int(lead_time_months)          # es. 2
    lead_time_frac = lead_time_months - lead_time_full  # es. 0.3

    # MESI INTERI: somma domanda per tutti i mesi pieni
    for month_idx in range(min(lead_time_full, len(month_col_indices))):
        col_indices = month_col_indices.get(month_idx, [])

        if not col_indices:
            continue

        # NUMPY FANCY INDEXING: estrai sottomatrice [righe matching × colonne mese]
        month_data = data_matrix[np.ix_(matching_indices, col_indices)]

        # Somma per colonna (ogni colonna = un run)
        aggregated[:len(col_indices)] += month_data.sum(axis=0).astype(np.float32)

    # MESE PARZIALE: aggiungi frazione del mese successivo (interpolazione lineare)
    if lead_time_frac > 0 and lead_time_full < len(month_col_indices):
        col_indices_partial = month_col_indices.get(lead_time_full, [])

        if col_indices_partial:
            # Estrai domanda del mese parziale
            partial_month_data = data_matrix[np.ix_(matching_indices, col_indices_partial)]

            # Somma per colonna e moltiplica per la frazione
            partial_demand = partial_month_data.sum(axis=0).astype(np.float32) * lead_time_frac
            aggregated[:len(col_indices_partial)] += partial_demand

    return aggregated


# ============================================================================
# CALCOLO STATISTICHE
# ============================================================================

def calculate_statistics(run_demands: np.ndarray, percentile: int = 99) -> Dict:
    """
    Calcola statistiche sulla domanda aggregata dai run Monte Carlo.

    STATISTICHE CALCOLATE:
    - mean_demand: Media della domanda attraverso i run
    - std_demand: Deviazione standard (misura variabilità)
    - p50_demand: Mediana (percentile 50°)
    - p<N>_demand: Percentile N° scelto dall'utente (default 99°)
    - safety_stock_p<N>: Buffer = percentile_N - media
    - frequency: Frazione di run dove lo SKU è usato

    CONCETTO SAFETY STOCK:
    safety_stock = percentile - media
    Rappresenta quanto stock extra tenere oltre la media attesa.

    Args:
        run_demands: Array con domanda per ogni run
        percentile:  Percentile su cui calcolare la safety stock (default 99).
                     Valori ammessi: 1-99 (intero).

    Returns:
        Dizionario con tutte le statistiche. La chiave della safety stock
        usa il percentile scelto (es. 'safety_stock_p99' o 'safety_stock_p95').
    """
    pct_key = f'p{percentile}'  # es. "p99", "p95"
    ss_key  = f'safety_stock_{pct_key}'  # es. "safety_stock_p99"

    # Caso speciale: nessuna domanda
    if len(run_demands) == 0 or np.all(run_demands == 0):
        return {
            'mean_demand': 0.0,
            'std_demand': 0.0,
            f'{pct_key}_demand': 0.0,
            ss_key: 0.0,
            'min_demand': 0.0,
            'max_demand': 0.0,
            'n_runs_with_demand': 0,
            'frequency': 0.0
        }

    # Calcola statistiche base
    mean_demand = float(np.mean(run_demands))
    std_demand  = float(np.std(run_demands))
    pct_demand  = float(np.percentile(run_demands, percentile))

    return {
        'mean_demand': mean_demand,
        'std_demand': std_demand,
        f'{pct_key}_demand': pct_demand,
        ss_key: max(0.0, pct_demand - mean_demand),  # Safety stock (mai negativa)
        'min_demand': float(np.min(run_demands)),
        'max_demand': float(np.max(run_demands)),
        'n_runs_with_demand': int(np.sum(run_demands > 0)),
        'frequency': float(np.sum(run_demands > 0) / len(run_demands))
    }


# ============================================================================
# UTILITY - PARSING MESI
# ============================================================================

def parse_month_to_sortable(month_str: str) -> tuple:
    """
    Converte stringa mese in tupla ordinabile (anno, mese_numero).

    Usato per ordinare i mesi cronologicamente.
    Es: "Jan 2026" -> (2026, 1), "Dec 2025" -> (2025, 12)

    Args:
        month_str: Stringa mese (es. "Jan 2026")

    Returns:
        Tupla (anno, mese) per ordinamento
    """
    try:
        parts = month_str.strip().split()
        if len(parts) == 2:
            month_name, year = parts
            month_num = datetime.strptime(month_name, '%b').month
            return (int(year), month_num)
        else:
            return (9999, 99)  # Valore alto per mettere in fondo
    except:
        return (9999, 99)


# ============================================================================
# FUNZIONE PRINCIPALE MATCHING
# ============================================================================

def match_skus_ultra_optimized(sku_catalog: pd.DataFrame,
                               lead_times: pd.DataFrame,
                               simulation_output: pd.DataFrame,
                               column_mapping: Dict[str, str],
                               n_runs: int,
                               n_processes: int = 0,
                               percentile: int = 99) -> pd.DataFrame:
    """
    Esegue il matching SKU-configurazioni con tutte le ottimizzazioni.

    FLUSSO:
    1. Pre-normalizza le colonne del simulation output
    2. Prepara matrice numpy per accesso veloce
    3. Per ogni SKU:
       - Trova configurazioni matching
       - Aggrega domanda per lead time
       - Calcola statistiche

    Args:
        sku_catalog: DataFrame catalogo SKU
        lead_times: DataFrame lead time componenti
        simulation_output: DataFrame dal Monte Carlo
        column_mapping: Mapping colonne
        n_runs: Numero run Monte Carlo
        n_processes: Numero processi paralleli (non usato, per compatibilità)

    Returns:
        DataFrame con statistiche per ogni SKU
    """
    print(f"\n{'='*60}")
    print(f"MATCHING SKU (ULTRA-OPTIMIZED)")
    print(f"{'='*60}")
    print(f"Memoria iniziale: {get_memory_usage_mb():.1f} MB")

    # =========================================================================
    # STEP 1: Pre-normalizza colonne simulazione
    # =========================================================================
    # Converte tutte le stringhe in formato normalizzato per matching
    normalized_df, norm_arrays_raw = prenormalize_simulation_output_fast(
        simulation_output, column_mapping
    )
    print(f"Memoria dopo pre-norm: {get_memory_usage_mb():.1f} MB")

    # Costruisci mappe categoria -> codice per matching numpy
    # Esempio: {"rosso": 0, "nero": 1, "bianco": 2}
    category_maps = {}
    norm_arrays = {}
    for col_name in norm_arrays_raw.keys():
        if col_name in normalized_df.columns:
            cat_col = normalized_df[col_name]
            if hasattr(cat_col, 'cat'):
                # Mappa valore -> codice numerico
                category_maps[col_name] = {
                    cat: code for code, cat in enumerate(cat_col.cat.categories)
                }
                norm_arrays[col_name] = cat_col.cat.codes.values

    n_rows = len(normalized_df)

    # =========================================================================
    # STEP 2: Prepara matrice numpy (solo colonne run)
    # =========================================================================
    print("\n[SETUP] Preparazione matrice numpy...")

    # Identifica colonne run (contengono "_Run_" nel nome)
    month_run_columns = [col for col in simulation_output.columns if '_Run_' in col]

    # Estrai prefissi mese unici e ordinali cronologicamente
    month_prefixes = []
    seen_prefixes = set()
    for col in month_run_columns:
        if "_Run_" in col:
            prefix = col.split("_Run_")[0]  # Es: "Jan 2026"
            if prefix not in seen_prefixes:
                month_prefixes.append(prefix)
                seen_prefixes.add(prefix)
    month_prefixes.sort(key=parse_month_to_sortable)

    # Lista ordinata di tutte le colonne run
    run_columns = []
    for prefix in month_prefixes:
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in simulation_output.columns:
                run_columns.append(col_name)

    print(f"  Colonne run: {len(run_columns)}")

    # OTTIMIZZAZIONE: Estrai matrice numpy int16
    # int16 usa 2 byte invece di 8 (int64) -> risparmio 75%
    data_matrix = simulation_output[run_columns].values.astype(np.int16)
    print(f"  Matrice shape: {data_matrix.shape}")
    print(f"  Memoria matrice: {data_matrix.nbytes / 1024 / 1024:.1f} MB")

    # Pre-calcola indici colonne per ogni mese
    run_col_to_idx = {col: idx for idx, col in enumerate(run_columns)}
    month_col_indices = {}
    for month_idx, prefix in enumerate(month_prefixes):
        indices = []
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in run_col_to_idx:
                indices.append(run_col_to_idx[col_name])
        month_col_indices[month_idx] = indices

    # =========================================================================
    # STEP 3: Prepara dati SKU
    # =========================================================================
    print("\n[SETUP] Preparazione dati SKU...")

    # Join catalogo SKU con lead times (LEFT: tutti i componenti BOM, anche senza LT)
    sku_with_lt = sku_catalog.merge(
        lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                    'Quantity_Per_Bike', 'Description']],
        left_on='Numero componenti',
        right_on='SKU',
        how='left'
    )

    # Componenti BOM senza lead time nel file residual: inclusi con LT=0 e flag
    # (la fonte primaria è la BOM, non il file residual)
    _no_lt_mask = sku_with_lt['Lead_Time_Months_Ceil'].isna() | (sku_with_lt['Lead_Time_Months_Ceil'] <= 0)
    _n_no_lt = int(_no_lt_mask.sum())
    if _n_no_lt > 0:
        _missing_skus = sku_with_lt.loc[_no_lt_mask, 'Numero componenti'].unique()
        print(f"  [INFO] {_n_no_lt} righe ({len(_missing_skus)} SKU unici) senza lead time "
              f"in quantity e residual — inclusi con LT=0, SS=0")
        sku_with_lt.loc[_no_lt_mask, 'Lead_Time_Months']      = 0.0
        sku_with_lt.loc[_no_lt_mask, 'Lead_Time_Months_Ceil'] = 0.0
        sku_with_lt.loc[_no_lt_mask, 'Lead_Time_Missing']     = True
    sku_with_lt['Lead_Time_Missing'] = sku_with_lt.get('Lead_Time_Missing', False).fillna(False)

    # Quantity_Per_Bike default 1 se mancante
    sku_with_lt['Quantity_Per_Bike'] = sku_with_lt['Quantity_Per_Bike'].fillna(1.0)

    # Raggruppa per SKU (stesso SKU può avere più righe con caratteristiche diverse)
    sku_groups = sku_with_lt.groupby('SKU')
    n_sku_groups = len(sku_groups)

    print(f"  SKU da processare: {n_sku_groups} "
          f"(di cui {len(_missing_skus) if _n_no_lt > 0 else 0} senza LT)")

    # Pre-estrai dati per ogni SKU (evita accesso ripetuto al DataFrame)
    sku_ids = []
    sku_data_list = []

    for sku_id, sku_rows_group in sku_groups:
        sku_ids.append(sku_id)

        # Estrai valori caratteristiche per ogni riga dello SKU
        rows_values = []
        for _, row in sku_rows_group.iterrows():
            row_values = {}
            for sku_col in column_mapping.keys():
                if sku_col in row.index:
                    row_values[sku_col] = row[sku_col]
            rows_values.append(row_values)

        sku_data_list.append({
            'values': rows_values,
            'lead_time': float(sku_rows_group['Lead_Time_Months_Ceil'].max()),  # Mantieni decimali per interpolazione lineare
            'quantity': float(sku_rows_group['Quantity_Per_Bike'].mean()),
            'description': sku_rows_group['Description'].iloc[0] if 'Description' in sku_rows_group.columns else '',
            'lead_time_raw': sku_rows_group['Lead_Time_Months'].iloc[0] if 'Lead_Time_Months' in sku_rows_group.columns else np.nan,
            'lead_time_missing': bool(sku_rows_group['Lead_Time_Missing'].any()) if 'Lead_Time_Missing' in sku_rows_group.columns else False
        })

    # =========================================================================
    # STEP 4: Processing loop principale
    # =========================================================================
    print(f"\n[MATCHING] Elaborazione {n_sku_groups} SKU...")
    print(f"  Progress: ogni '.' = 50 SKU")

    start_time = time.time()
    all_results = []

    for sku_idx, (sku_id, sku_data) in enumerate(zip(sku_ids, sku_data_list)):
        # Progress indicator
        if sku_idx % 50 == 0:
            print('.', end='', flush=True)
        if sku_idx % 500 == 0 and sku_idx > 0:
            elapsed = time.time() - start_time
            rate = sku_idx / elapsed
            remaining = (n_sku_groups - sku_idx) / rate if rate > 0 else 0
            print(f' [{sku_idx}/{n_sku_groups}] {rate:.1f} SKU/s, ETA: {remaining/60:.1f}min')

        # -----------------------------------------------------------------
        # MATCHING: trova configurazioni che usano questo SKU
        # -----------------------------------------------------------------
        # Uno SKU può avere più righe (es. stesso componente per colori diversi)
        # Una configurazione matcha se corrisponde ad ALMENO UNA riga
        combined_mask = np.zeros(n_rows, dtype=bool)

        for row_values in sku_data['values']:
            row_mask = sku_matches_combination_numpy(
                row_values, norm_arrays, category_maps,
                column_mapping, n_rows
            )
            combined_mask |= row_mask  # OR: matcha se almeno una riga corrisponde

        # Indici delle configurazioni che usano questo SKU
        matching_indices = np.where(combined_mask)[0]

        # Nessun match -> salta questo SKU
        if len(matching_indices) == 0:
            continue

        # -----------------------------------------------------------------
        # AGGREGAZIONE: somma domanda per i mesi del lead time (aggregato)
        # -----------------------------------------------------------------
        run_demands = aggregate_monthly_demands_numpy(
            matching_indices, data_matrix, month_col_indices,
            sku_data['lead_time'], n_runs
        )

        # -----------------------------------------------------------------
        # STATISTICHE AGGREGATE: calcola media, percentili, safety stock
        # -----------------------------------------------------------------
        stats = calculate_statistics(run_demands, percentile=percentile)

        # -----------------------------------------------------------------
        # STATISTICHE MENSILI: calcola statistiche mese per mese
        # -----------------------------------------------------------------
        # Per ogni mese disponibile, aggreghiamo solo quel mese (lead_time=1)
        monthly_stats = {}
        for month_idx, prefix in enumerate(month_prefixes):
            col_indices_month = month_col_indices.get(month_idx, [])
            if not col_indices_month:
                continue

            # Seleziona solo i dati delle configurazioni matching per questo mese
            month_data = data_matrix[np.ix_(matching_indices, col_indices_month)]
            # Somma per run (ogni colonna = un run)
            run_demands_month = month_data.sum(axis=0).astype(np.int32)

            # Calcola statistiche per questo mese
            monthly_stats[prefix] = calculate_statistics(run_demands_month, percentile=percentile)

        # -----------------------------------------------------------------
        # RISULTATO: componi record output
        # -----------------------------------------------------------------
        ss_key = f'safety_stock_p{percentile}'   # es. 'safety_stock_p99'

        result = {
            'SKU': sku_id,
            'Description': sku_data['description'],
            'Quantity_Per_Bike': sku_data['quantity'],
            'Lead_Time_Months': sku_data['lead_time_raw'],
            'Lead_Time_Months_Ceil': sku_data['lead_time'],
            'Lead_Time_Missing': sku_data.get('lead_time_missing', False),
            'N_Matching_Combinations': len(matching_indices),
            'N_SKU_Rows_Original': len(sku_data['values']),
            'monthly_stats': monthly_stats,   # dict mensile per save_results_monthly()
        }

        # Aggiungi statistiche aggregate calcolate
        result.update(stats)

        # Safety stock totale = safety stock unitario × quantità per moto
        result[f'{ss_key}_total'] = result[ss_key] * sku_data['quantity']

        all_results.append(result)

        # Garbage collection periodico per evitare accumulo memoria
        if sku_idx % 1000 == 0:
            gc.collect()

    # Report finale
    elapsed = time.time() - start_time
    rate = len(all_results) / elapsed if elapsed > 0 else 0

    print(f"\n\n[OK] Matching completato!")
    print(f"  Tempo: {elapsed:.1f}s ({elapsed/60:.1f}min)")
    print(f"  Velocità: {rate:.1f} SKU/s")
    print(f"  SKU con match: {len(all_results)}")
    print(f"  Memoria finale: {get_memory_usage_mb():.1f} MB")

    return pd.DataFrame(all_results)


# ============================================================================
# SALVATAGGIO RISULTATI
# ============================================================================

def save_results(results_df: pd.DataFrame,
                 filename: str = ".\\Output\\sku_safety_stock_leadtime_PER_SKU.xlsx",
                 percentile: int = 99):
    """
    Salva i risultati del matching in formato Excel - OUTPUT AGGREGATO.

    Colonne output (ordine fisso):
        SKU, Description, mean_demand, safety_stock_p<N>, Quantity_Per_Bike,
        safety_stock_p<N>_total, Lead_Time_Months_Ceil,
        N_Matching_Combinations, N_SKU_Rows_Original

    I risultati sono ordinati per safety_stock_p<N>_total decrescente
    (componenti più critici in cima).

    Args:
        results_df:  DataFrame con risultati matching
        filename:    Nome file output
        percentile:  Percentile usato per la safety stock (default 99)

    Returns:
        DataFrame ordinato (lo stesso salvato su file)
    """
    if len(results_df) == 0:
        print("\nNESSUN RISULTATO DA SALVARE!")
        return results_df

    ss_key       = f'safety_stock_p{percentile}'        # es. 'safety_stock_p99'
    ss_total_key = f'safety_stock_p{percentile}_total'  # es. 'safety_stock_p99_total'

    # Ordina per criticità (safety stock più alto prima)
    sort_col = ss_total_key if ss_total_key in results_df.columns else results_df.columns[-1]
    results_sorted = results_df.sort_values(sort_col, ascending=False).copy()

    # Arrotonda colonne numeriche per leggibilità
    numeric_cols = ['mean_demand', ss_key, ss_total_key, 'Quantity_Per_Bike']
    for col in numeric_cols:
        if col in results_sorted.columns:
            results_sorted[col] = results_sorted[col].round(2)

    # Colonna safety stock totale arrotondata all'intero
    ss_total_int_key = f'{ss_total_key}_int'
    if ss_total_key in results_sorted.columns:
        results_sorted[ss_total_int_key] = results_sorted[ss_total_key].round(0).astype('Int64')

    # Ordine colonne output (solo quelle richieste)
    colonne = [
        'SKU',
        'Description',
        'mean_demand',
        ss_key,
        'Quantity_Per_Bike',
        ss_total_key,
        ss_total_int_key,
        'Lead_Time_Months_Ceil',
        'Lead_Time_Missing',
        'N_Matching_Combinations',
        'N_SKU_Rows_Original',
        'prezzo_safety',
        'SOLO_MODELLO',
    ]

    # Filtra solo colonne esistenti nel DataFrame
    colonne_esistenti = [col for col in colonne if col in results_sorted.columns]
    results_output = results_sorted[colonne_esistenti]

    # Salva
    results_output.to_excel(filename, index=False)

    print(f"\n[OK] File aggregato salvato: {filename}")
    print(f"   Percentile safety stock: p{percentile}")
    print(f"   Righe: {len(results_output)}")
    print(f"   Colonne: {list(results_output.columns)}")

    return results_sorted


def save_results_monthly(results_df: pd.DataFrame,
                         filename: str = ".\\Output\\sku_safety_stock_PER_MESE.xlsx",
                         percentile: int = 99):
    """
    Salva le statistiche mensili per ogni SKU - OUTPUT PER MESE.

    Per ogni SKU produce una riga con le colonne mensili ripetute:
        SKU, Description, Quantity_Per_Bike, Lead_Time_Months_Ceil,
        N_Matching_Combinations, N_SKU_Rows_Original,
        mean_demand_<mese>, safety_stock_p<N>_<mese>, safety_stock_p<N>_total_<mese>,
        ... (ripetuto per ogni mese)

    I mesi sono ordinati cronologicamente.

    Args:
        results_df:  DataFrame con risultati matching (deve contenere
                     la colonna 'monthly_stats' prodotta dal loop principale)
        filename:    Nome file output
        percentile:  Percentile usato per la safety stock (default 99)

    Returns:
        DataFrame mensile salvato
    """
    if len(results_df) == 0:
        print("\nNESSUN RISULTATO MENSILE DA SALVARE!")
        return pd.DataFrame()

    # Verifica presenza colonna monthly_stats
    if 'monthly_stats' not in results_df.columns:
        print("\n[WARN] Colonna 'monthly_stats' non presente: output mensile non disponibile")
        return pd.DataFrame()

    # ------------------------------------------------------------------
    # Costruisci l'elenco GLOBALE dei mesi in ordine cronologico,
    # raccogliendo tutti i mesi presenti in qualsiasi SKU.
    # Serve per garantire che ogni riga abbia le stesse colonne.
    # ------------------------------------------------------------------
    all_months_set = set()
    for _, row in results_df.iterrows():
        ms = row.get('monthly_stats', {})
        if ms:
            all_months_set.update(ms.keys())

    all_months_sorted = sorted(all_months_set, key=parse_month_to_sortable)

    monthly_rows = []

    for _, row in results_df.iterrows():
        monthly_stats = row['monthly_stats']  # dict: {mese: {mean, safety_p99, ...}}

        if not monthly_stats:
            continue

        lead_time = float(row['Lead_Time_Months_Ceil'])  # numero mesi da coprire (float per LT < 1)
        qty       = float(row['Quantity_Per_Bike'])

        # Colonne fisse identificative
        base = {
            'SKU': row['SKU'],
            'Description': row['Description'],
            'Quantity_Per_Bike': round(qty, 2),
            'Lead_Time_Months_Ceil': lead_time,
            'N_Matching_Combinations': row['N_Matching_Combinations'],
            'N_SKU_Rows_Original': row['N_SKU_Rows_Original'],
        }

        ss_key = f'safety_stock_p{percentile}'  # es. 'safety_stock_p99'

        # Itera su TUTTI i mesi globali in ordine cronologico.
        # Per i primi `lead_time` mesi -> usa i dati reali.
        # Per i mesi oltre il lead time -> forza zero.
        # Arrotonda lead_time per eccesso per includere il primo mese anche con LT < 1
        lead_time_ceil = math.ceil(lead_time) if lead_time > 0 else 0
        for month_rank, mese in enumerate(all_months_sorted, start=1):
            mese_key = mese.replace(' ', '')  # "Jan 2026" -> "Jan2026"

            if month_rank <= lead_time_ceil and mese in monthly_stats:
                # Mese coperto dal lead time: popola con i dati calcolati
                stats_m  = monthly_stats[mese]
                mean_val = round(stats_m.get('mean_demand', 0.0), 2)
                ss_val   = round(stats_m.get(ss_key, 0.0), 2)
                ss_tot   = round(ss_val * qty, 2)
            else:
                # Mese oltre il lead time (o mese non simulato): zero
                mean_val = 0.0
                ss_val   = 0.0
                ss_tot   = 0.0

            base[f'mean_demand_{mese_key}']              = mean_val
            base[f'{ss_key}_{mese_key}']                 = ss_val
            base[f'{ss_key}_total_{mese_key}']           = ss_tot

        monthly_rows.append(base)

    if not monthly_rows:
        print("\n[WARN] Nessuna riga mensile generata")
        return pd.DataFrame()

    df_monthly = pd.DataFrame(monthly_rows)

    # Ordina per SKU
    df_monthly = df_monthly.sort_values('SKU').reset_index(drop=True)

    df_monthly.to_excel(filename, index=False)

    print(f"\n[OK] File mensile salvato: {filename}")
    print(f"   Righe: {len(df_monthly)}")
    # Stampa i mesi trovati
    mesi_cols = [c for c in df_monthly.columns if c.startswith('mean_demand_')]
    print(f"   Mesi: {[c.replace('mean_demand_','') for c in mesi_cols]}")

    return df_monthly


# ============================================================================
# FUNZIONE MAIN - ENTRY POINT
# ============================================================================

def main_sku_ultra_optimized(simulation_output: pd.DataFrame,
                             percentile: int = 99) -> pd.DataFrame:
    """Alias per retrocompatibilità -> chiama main_sku_v2."""
    return main_sku_v2(simulation_output, percentile=percentile)


def main_sku_v2(simulation_output: pd.DataFrame,
                percentile: int = 99,
                sku_catalog_path: str = '.\\Input\\tutte_righe_univoche_V2.xlsx',
                lead_times_path:  str = '.\\Input\\quantity e residual.xlsx',
                prezzi_path:      str = '.\\Input\\Prezzi.xlsx',
                output_agg:       str = '.\\Output\\sku_safety_stock_leadtime_PER_SKU_V2.xlsx',
                output_monthly:   str = '.\\Output\\sku_safety_stock_PER_MESE_V2.xlsx') -> pd.DataFrame:
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
        lead_times_path:   Path file lead times
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
    lead_times = load_lead_times(lead_times_path)

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

    # Tutti i modelli distinti nel catalogo (gestisce "MSV4+MSV4S" split)
    _all_models_catalog = set()
    if _model_col:
        for _mv in sku_catalog_raw[_model_col].dropna().unique():
            for _m in str(_mv).split('+'):
                _all_models_catalog.add(_m.strip())

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
                    _sku_models.add(_m.strip())
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

        # Foglio 2: tassi di cambio verso EUR
        # Prima cerchiamo nel file Cambio Valuta.xlsx (se esiste e ha dati)
        cambio = None
        cambio_path = 'Input/Cambio Valuta.xlsx'
        
        try:
            if os.path.exists(cambio_path):
                cambio_temp = pd.read_excel(cambio_path)
                if len(cambio_temp) > 0 and len(cambio_temp.columns) > 0:
                    print(f"   [INFO] Trovato file Cambio Valuta.xlsx con {len(cambio_temp)} righe")
                    # Try to find currency code and exchange rate columns
                    code_col = None
                    rate_col = None
                    
                    for col in cambio_temp.columns:
                        col_lower = str(col).lower()
                        if "codice" in col_lower or "currency" in col_lower or "divisa" in col_lower:
                            code_col = col
                        if "eur" in col_lower or "tasso" in col_lower or "rate" in col_lower:
                            rate_col = col
                    
                    if code_col and rate_col:
                        cambio = cambio_temp[[code_col, rate_col]].rename(columns={
                            code_col: "Codice",
                            rate_col: "1 unità in EUR ≈"
                        })
                        print(f"   [INFO] Tassi di cambio caricati da Cambio Valuta.xlsx")
                    else:
                        print(f"   [WARN] File Cambio Valuta.xlsx trovato ma colonne non riconosciute")
        except Exception as e:
            print(f"   [WARN] Errore nella lettura di Cambio Valuta.xlsx: {e}")
        
        # Se Cambio Valuta.xlsx è vuoto o non esiste, cerchiamo all'interno di PREZZI.xlsx
        if cambio is None or len(cambio) == 0:
            sheet_names = pd.ExcelFile(prezzi_path).sheet_names
            
            # Try common sheet names for currency exchange
            for sheet_name in sheet_names:
                if "cambio" in sheet_name.lower() or "valuta" in sheet_name.lower() or "eur" in sheet_name.lower():
                    try:
                        cambio = pd.read_excel(prezzi_path, sheet_name=sheet_name)
                        # Try to find currency code and exchange rate columns
                        code_col = None
                        rate_col = None
                        
                        for col in cambio.columns:
                            col_lower = str(col).lower()
                            if "codice" in col_lower or "currency" in col_lower or "divisa" in col_lower:
                                code_col = col
                            if "eur" in col_lower or "tasso" in col_lower or "rate" in col_lower:
                                rate_col = col
                        
                        if code_col and rate_col:
                            cambio = cambio[[code_col, rate_col]].rename(columns={
                                code_col: "Codice",
                                rate_col: "1 unità in EUR ≈"
                            })
                            print(f"   [INFO] Tassi di cambio caricati da foglio {sheet_name} in PREZZI.xlsx")
                            break
                    except Exception:
                        continue
        
        # If no currency sheet found, try to extract from Sheet1 if it has exchange rate info
        if cambio is None or len(cambio) == 0:
            try:
                # Look for exchange rate information in Sheet1
                sheet1_data = pd.read_excel(prezzi_path, sheet_name="Sheet1")
                # Check if there are multiple currencies and exchange rates
                if "Divisa" in sheet1_data.columns and "Prezzo netto" in sheet1_data.columns:
                    # Group by currency and get average exchange rate (approximation)
                    cambio_data = sheet1_data.groupby("Divisa")["Prezzo netto"].mean().reset_index()
                    cambio_data["1 unità in EUR ≈"] = 1.0  # Default to 1.0 (EUR)
                    cambio = cambio_data[["Divisa", "1 unità in EUR ≈"]].rename(columns={"Divisa": "Codice"})
                    print(f"   [INFO] Tassi di cambio stimati da Sheet1 in PREZZI.xlsx")
            except Exception:
                pass

        # If still no currency data, create a default with EUR=1.0
        if cambio is None or len(cambio) == 0:
            print("   [INFO] Nessun tasso di cambio trovato. Usando EUR=1.0 come predefinito")
            # Create a minimal currency table with just EUR
            cambio = pd.DataFrame({"Codice": ["EUR"], "1 unità in EUR ≈": [1.0]})

        # Join tasso di cambio sul codice valuta
        prezzi = prezzi.merge(cambio, left_on=divisa_col, right_on="Codice", how="left")

        # Fallback: se la valuta non è in tabella (es. voce mancante) assume già EUR
        n_missing = prezzi["1 unità in EUR ≈"].isna().sum()
        if n_missing > 0:
            print(f"   [WARN] {n_missing} righe con valuta non trovata -> tasso=1.0 (assumendo EUR)")
        prezzi["1 unità in EUR ≈"] = prezzi["1 unità in EUR ≈"].fillna(1.0)

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


# ============================================================================
# ENTRY POINT PER ESECUZIONE DIRETTA
# ============================================================================

if __name__ == "__main__":
    # Test standalone V2: carica simulazione da file
    simulation_output = pd.read_excel('.\\Input\\monthly_output_wide.xlsx')
    results = main_sku_v2(simulation_output)

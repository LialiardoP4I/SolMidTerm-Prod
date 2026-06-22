# -*- coding: utf-8 -*-
"""Abbinamento componenti-scenari e calcolo della safety stock.

In parole semplici: confronta ogni componente (SKU) con gli scenari di vendita
simulati, calcola quanta domanda riceve durante il tempo di approvvigionamento e
quanta scorta di sicurezza serve (al percentile scelto), poi la valorizza in euro.

(dettaglio tecnico) SKU matching e calcolo safety stock per la pipeline mid-term.

Per ogni SKU questo modulo:
- carica catalogo, lead time residui (da Gates + Dashboard + BOM) e prezzi;
- abbina lo SKU alle configurazioni del Monte Carlo via matching vettoriale
  numpy con cache di normalizzazione e column mapping auto-rilevato (zero
  hardcoding);
- aggrega la domanda sui mesi di lead time (interpolazione lineare per LT
  decimali) e calcola statistiche / safety stock (percentile parametrico).

La logica è invariata rispetto a sku_matching_SS.py; questo modulo espone le
funzioni puramente algoritmiche (catalogo, lead-time, matching) chiamate
dall'orchestratore mid_term.pipeline.
"""

import numpy as np
import pandas as pd
from typing import Dict, List, Tuple
import math
import time
from datetime import datetime
import gc
# import psutil  # optional, commented for portability
import os
import tempfile
import warnings

from mid_term.tr import canonical_model_name
from mid_term._logging import get_logger
from mid_term.exceptions import MatchingError

log = get_logger("matching")

# Restringi soppressione warning a categorie note (FutureWarning/Deprecation)
# per non nascondere bug futuri (es. DeprecationWarning su .values, applymap, ecc.)
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=DeprecationWarning)


# ============================================================================
# ECCEZIONI CUSTOM
# ============================================================================

class CambioValutaError(Exception):
    """Errore fatale sul caricamento dei tassi di cambio EUR.
    Indica che Input/Cambio Valuta.xlsx è mancante, vuoto, malformato
    o non copre tutte le valute presenti in Prezzi.xlsx.
    """
    pass


def _load_cambio_eur(cambio_path: str = 'Input/Cambio Valuta.xlsx') -> pd.DataFrame:
    """Carica i tassi di cambio EUR da Cambio Valuta.xlsx.

    Ritorna DataFrame con colonne ['Codice', '1 unità in EUR ≈'].
    Solleva CambioValutaError in caso di problema (file mancante, vuoto,
    colonne non riconosciute). Nessun fallback patologico.
    """
    if not os.path.exists(cambio_path):
        raise CambioValutaError(
            f"File obbligatorio mancante: {cambio_path}\n"
            f"Necessario per conversione prezzi multi-valuta in EUR."
        )
    cambio_temp = pd.read_excel(cambio_path)
    if cambio_temp.empty:
        raise CambioValutaError(f"File {cambio_path} è vuoto.")

    code_col = next((c for c in cambio_temp.columns
                     if any(kw in str(c).lower() for kw in ('codice', 'currency', 'divisa'))), None)
    rate_col = next((c for c in cambio_temp.columns
                     if any(kw in str(c).lower() for kw in ('eur', 'tasso', 'rate'))), None)
    if code_col is None or rate_col is None:
        raise CambioValutaError(
            f"File {cambio_path}: colonne attese non trovate.\n"
            f"Servono colonna 'Codice'/'Divisa'/'Currency' e 'EUR'/'Tasso'/'Rate'.\n"
            f"Colonne trovate: {list(cambio_temp.columns)}"
        )

    cambio = (cambio_temp[[code_col, rate_col]]
              .rename(columns={code_col: 'Codice', rate_col: '1 unità in EUR ≈'})
              .dropna(subset=['Codice', '1 unità in EUR ≈']))
    cambio['Codice'] = cambio['Codice'].astype(str).str.strip().str.upper()
    if cambio.empty:
        raise CambioValutaError(
            f"File {cambio_path}: nessuna riga valida (codice+tasso non-null) trovata."
        )
    return cambio


# ============================================================================
# CALCOLO RESIDUAL LEAD TIME (NUOVO V3)
# ============================================================================
# V3: Sostituisce load_lead_times() con processo ricavato da Gates MTSV4
# e Dashboard Supplier. Zero dipendenza da "quantity e residual.xlsx"
# ============================================================================

def calcola_residual_lead_time(input_dir: str = '.\\Input',
                               logistica_dir: str = '.\\Input\\Dati Logistica',
                               solo_modello_skus=None,
                               gate_max_mesi: float = 4) -> pd.DataFrame:
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

    log.info("[RESIDUAL_LT V3] Caricamento dati...")

    # 1. Carica tutte_righe_univoche_V2
    df_univoche = pd.read_excel(input_path / "tutte_righe_univoche_V2.xlsx")
    log.info("tutte_righe_univoche: %s", df_univoche.shape)

    # 2. Carica Gates
    df_gates = pd.read_excel(logistica_path / "Gates MTSV4.xlsx", skiprows=1)
    log.info("Gates: %s", df_gates.shape)

    GATE_COL = 'Gates (Weeks) rev'

    # 3a. Auto-detect colonne caratteristiche: Optional TRANSCOD (escluso BUNDLE) == colonna univoche
    gate_optionals = [c for c in df_gates['Optional TRANSCOD'].dropna().unique() if c != 'BUNDLE']
    config_cols = [c for c in gate_optionals if c in df_univoche.columns]
    log.info("Colonne caratteristiche matchate: %s", config_cols)

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
    log.info("Colonne bundle matchate: %s", bundle_cols)

    all_cols = config_cols + bundle_cols
    df_output = df_univoche[['Numero componenti'] + all_cols].copy()

    # 3c. Aggiungi colonna Versione per matching Supermodel
    versione_col = 'Versione' if 'Versione' in df_univoche.columns else None
    if versione_col:
        df_output['_versione'] = df_univoche[versione_col]
        log.info("Colonna Versione rilevata per Supermodel matching")

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
    log.info("Gate mapping creato con Supermodel filtering")

    # 6. Processa celle caratteristiche - sostituisci valori con gate weeks
    # Matching con Supermodel + fallback senza Supermodel
    log.info("[RESIDUAL_LT V3] Processing celle caratteristiche (con Supermodel matching)...")
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
    log.info("Celle caratteristiche replicate: %d (con Supermodel: %d)", replaced, replaced_with_sm)

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
    log.info("Celle bundle replicate: %d (con Supermodel: %d)", replaced_bundle, replaced_bundle_sm)

    # 6c. Rimuovi colonna temporanea versione
    if '_versione' in df_output.columns:
        df_output = df_output.drop(columns=['_versione'])

    # 7. Calcola MAX_gate per ogni riga (su tutte le colonne)
    log.info("[RESIDUAL_LT V3] Calcola MAX_gate...")
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
        """True se la riga ha almeno un valore "attivo" (diverso da No/NaN/vuoto)
        nelle colonne config/bundle."""
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
    log.info("[RESIDUAL_LT V3] Carica Dashboard + Prezzi...")
    prezzi_files = sorted(logistica_path.glob("PREZZI*.XLS*"))
    if not prezzi_files:
        prezzi_files = sorted(input_path.glob("PREZZI*.XLS*"))
    if not prezzi_files:
        raise FileNotFoundError(
            f"Nessun file PREZZI trovato in: {logistica_path} o {input_path}"
        )
    prezzi_path_found = prezzi_files[-1]
    log.info("Prezzi: %s", prezzi_path_found.name)
    df_prezzi = pd.read_excel(prezzi_path_found, sheet_name='Sheet1', skiprows=0)

    dash_files = sorted(logistica_path.glob("*Dashboard Supplier*.xlsx"))
    if not dash_files:
        raise FileNotFoundError(
            f"Nessun file 'Dashboard Supplier' trovato in: {logistica_path}\n"
            f"Atteso pattern: *Dashboard Supplier*.xlsx"
        )
    dash_path = dash_files[-1]
    log.info("Dashboard: %s", dash_path.name)
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
    log.info("Colonna fornitore PREZZI: '%s'", vendor_col)

    def extract_id(text):
        """Estrae il primo token (ID fornitore) da una cella; None se vuota."""
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
        log.info("Colonna prezzo: '%s'", price_col)
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

    # 12c. (RIMOSSO) Il vecchio default gate fisso 4 mesi per NO_GATE_ALL_NO e'
    # stato eliminato. Le SKU all-No NON-SOLO_MODELLO senza gate optional ora
    # restano senza gate (MAX_gate NaN -> NO_GATE_MISSING/ALL_NO -> scartate).

    # 13. Calcola Residual_LT (clampato a 0 se negativo)
    df_final['MAX_gate_mesi'] = df_final['MAX_gate'] / 4.33

    # 13b. GATE parametrico = gate_max_mesi. Due ruoli:
    #   (1) CAP GLOBALE: ogni SKU con MAX_gate_mesi > gate_max_mesi viene riportato
    #       a gate_max_mesi (i gate maggiori sostituiti dal parametrico in input).
    #   (2) DEFAULT SOLO_MODELLO: le SOLO_MODELLO senza gate optional (MAX_gate NaN)
    #       ricevono gate_max_mesi. Le NON-SOLO senza gate restano NaN -> scartate.
    # Poi formula standard; senza FROZEN/Freq (NO_DASHBOARD_MATCH) -> NaN -> scartata.
    _gmax = float(gate_max_mesi)
    df_final['gate_capped'] = False
    df_final['gate_solo_applied'] = False
    # (1) cap globale
    mask_cap = df_final['MAX_gate_mesi'].notna() & (df_final['MAX_gate_mesi'] > _gmax)
    df_final.loc[mask_cap, 'MAX_gate_mesi'] = _gmax
    df_final.loc[mask_cap, 'gate_capped'] = True
    # (2) default per SOLO_MODELLO senza gate
    if solo_modello_skus:
        _solo_set = {str(s) for s in solo_modello_skus}
        mask_solo_nogate = (df_final['Numero componenti'].astype(str).isin(_solo_set)
                            & df_final['MAX_gate_mesi'].isna())
        df_final.loc[mask_solo_nogate, 'MAX_gate_mesi'] = _gmax
        df_final.loc[mask_solo_nogate, 'gate_solo_applied'] = True
    log.info("[RESIDUAL_LT] gate cap=%s mesi su %d SKU; default SOLO_MODELLO su %d SKU",
             _gmax, int(mask_cap.sum()), int(df_final['gate_solo_applied'].sum()))

    df_final['Supply_cycle'] = df_final['FROZEN[Months] '] + df_final['Freq_mesi']
    df_final['Residual_LT'] = (df_final['Supply_cycle'] - df_final['MAX_gate_mesi']).clip(lower=0)

    # 14. Aggiungi colonna diagnostica match_status
    def _match_status(row):
        """Classifica l'esito del calcolo residual lead-time della riga
        (OK / NO_GATE_MISSING / NO_DASHBOARD_MATCH / NO_FROZEN / ...), per diagnostica."""
        solo = row.get('gate_solo_applied', False)
        if not solo and pd.isna(row.get('MAX_gate')):
            # MAX_gate NaN e non-SOLO -> nessun gate optional (default fisso rimosso).
            if row.get('has_non_no', False):
                return 'NO_GATE_MISSING'
            else:
                return 'NO_GATE_ALL_NO'   # all-No non-SOLO senza gate -> scartata
        if pd.isna(row.get('ID Supplier')):
            return 'NO_DASHBOARD_MATCH'
        if pd.isna(row.get('FROZEN[Months] ')):
            return 'NO_FROZEN'
        if pd.isna(row.get('Freq_mesi')):
            return 'NO_FREQUENCY'
        if solo:
            return 'OK_GATE_SOLO_MODELLO'  # gate parametrico SOLO_MODELLO, residual calcolato
        if row.get('Residual_LT', 0) == 0:
            return 'OK_ZERO'
        return 'OK'
    df_final['match_status'] = df_final.apply(_match_status, axis=1)

    # 14b. Salva debug completo in Dati Logistica
    try:
        debug_cols = ['Numero componenti', 'MAX_gate', 'MAX_gate_mesi',
                      'gate_capped', 'gate_solo_applied',
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
        log.info("[DEBUG] File completo salvato: %s", debug_path)
        log.info("[DEBUG] match_status breakdown:")
        for status, cnt in df_debug.get('match_status', pd.Series(dtype=str)).value_counts().items():
            log.info("   %s: %s", status, cnt)
        log.info("[DEBUG] Codici problematici: %d", len(df_problematici))
    except Exception as _e:
        log.warning("Impossibile salvare debug: %s", _e)
        log.debug("traceback", exc_info=True)

    # 15. Restituisci colonne essenziali
    # Freq_mesi e Supply_cycle (=FROZEN+Freq) servono al rolling per Cycle_Stock
    # e per il Flag_Static_SS (mese a distanza Total Lead Time).
    out_cols = ['Numero componenti', 'Residual_LT', 'ID Supplier']
    out_rename = ['SKU', 'Lead_Time_Months', 'ID_Supplier_ref']
    for _c in ('Freq_mesi', 'Supply_cycle'):
        if _c in df_final.columns:
            out_cols.append(_c)
            out_rename.append(_c)
    if 'Testo breve' in df_final.columns:
        out_cols.append('Testo breve')
        out_rename.append('Description')
    result = df_final[out_cols].copy()
    result.columns = out_rename

    log.info("[OK] Residual_LT calcolato: %d SKU con valore", result['Lead_Time_Months'].notna().sum())
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


def build_solo_modello_flags(sku_catalog_raw: pd.DataFrame,
                             column_mapping: Dict[str, str]) -> pd.DataFrame:
    """Calcola il flag SOLO_MODELLO per ogni SKU del catalogo.

    SOLO_MODELLO = (tutti gli optional == No)  AND  (set modelli != TUTTI i modelli).
    Cioe' la SKU non dipende da nessun optional (dipende solo dal modello) e non e'
    presente su tutti i modelli (altrimenti sarebbe un componente comune).

    Logica identica a sku_matching_SS.main_sku_v2 (ALL). Estratta per essere
    condivisa tra run singolo (pipeline) e rolling.

    Args:
        sku_catalog_raw: catalogo SKU con colonna 'Numero componenti', 'Model' e
                         le colonne caratteristiche.
        column_mapping:  mapping {col_catalogo -> col_simulazione}; le sue chiavi
                         (meno 'model') sono le colonne optional da controllare.

    Returns:
        DataFrame con colonne ['SKU', 'SOLO_MODELLO'].
    """
    _model_col = next(
        (c for c in sku_catalog_raw.columns if c.lower().strip() == "model"),
        None,
    )
    _opt_flag_cols = [k for k in column_mapping.keys()
                      if _model_col is None or k.lower().strip() != "model"]

    def _is_no_val(v):
        return str(v).strip().lower() in ("no", "nan", "")

    _all_models_catalog = set()
    if _model_col:
        for _mv in sku_catalog_raw[_model_col].dropna().unique():
            for _m in str(_mv).split("+"):
                _canon = canonical_model_name(_m)
                if _canon:
                    _all_models_catalog.add(_canon)

    _flag_rows = []
    for _sku, _grp in sku_catalog_raw.groupby("Numero componenti"):
        _all_no = all(
            _is_no_val(v) for _c in _opt_flag_cols
            if _c in _grp.columns
            for v in _grp[_c]
        )
        _sku_models = set()
        if _model_col:
            for _mv in _grp[_model_col].dropna():
                for _m in str(_mv).split("+"):
                    _canon = canonical_model_name(_m)
                    if _canon:
                        _sku_models.add(_canon)
        _not_all = bool(_all_models_catalog) and (
            _sku_models != _all_models_catalog
        )
        _flag_rows.append({"SKU": _sku, "SOLO_MODELLO": _all_no and _not_all})

    return pd.DataFrame(_flag_rows)


# ============================================================================
# UTILITY - MONITORAGGIO MEMORIA
# ============================================================================

def get_memory_usage_mb():
    """
    Restituisce la memoria RAM usata dal processo corrente in MB.
    Utile per monitorare l'impatto delle ottimizzazioni sulla memoria.
    Ritorna 0.0 se psutil non disponibile (evita TypeError su f-string `:.1f`).
    """
    try:
        import psutil
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024
    except ImportError:
        return 0.0


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
        """Inizializza la cache di normalizzazione. `prefixes` sono i prefissi
        (derivati dai nomi colonna, es. 'fairing color - ') rimossi dai valori TR
        prima del confronto."""
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
            # Rimuove caratteri Unicode invisibili (ZWS/ZWNJ/ZWJ/BOM/NBSP/...)
            # presenti in alcuni export Excel da PDF/Word/SAP: prima del normale
            # strip+lower garantisce match simmetrico fra catalogo SKU e
            # output simulazione canonicalizzati a monte.
            value_str = str(value)
            for _ch in ('​', '‌', '‍', '﻿',
                        ' ', ' ', '⁠', '᠎'):
                if _ch in value_str:
                    value_str = value_str.replace(_ch, '')
            value_str = value_str.strip().lower()

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
    log.info("[NormCache V2] Re-inizializzata con %d prefissi dinamici", len(prefixes))


# ============================================================================
# CARICAMENTO DATI - LEAD TIME
# ============================================================================

def load_lead_times(solo_modello_skus=None,
                    gate_max_mesi: float = 4) -> pd.DataFrame:
    """
    Carica i lead time dei componenti - VERSIONE V3 (Residual LT da Gates)

    gate_max_mesi: cap massimo (mesi) sul gate di OGNI SKU + default per le
    SOLO_MODELLO senza gate. solo_modello_skus: insieme SKU SOLO_MODELLO
    (vedi calcola_residual_lead_time).

    V3: deriva tutto da Gates MTSV4 + Dashboard Supplier + Prezzi + BOM.
    Il file 'quantity e residual.xlsx' è stato eliminato dal progetto.

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

    log.info("CARICAMENTO LEAD TIME - V3 (Residual LT da Gates)")

    input_dir = Path('.') / 'Input'
    logistica_dir = input_dir / 'Dati Logistica'

    # Trova BOM_150 in qualsiasi sottocartella MODEL/*/
    def _find_bom150(base: Path) -> Path:
        """Cerca il primo file BOM_150*.xlsx sotto base/MODEL/<versione>/.
        Solleva FileNotFoundError se nessuno e' presente."""
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
    lead_times_residual = calcola_residual_lead_time(
        str(input_dir), str(logistica_dir),
        solo_modello_skus=solo_modello_skus,
        gate_max_mesi=gate_max_mesi)
    log.info("Residual LT calcolato: %d SKU", len(lead_times_residual))

    # (debug salvato dentro calcola_residual_lead_time)

    # 2. Carica BOM per le quantità
    log.info("Caricamento BOM per quantità...")
    bom_path = _find_bom150(input_dir)
    log.info("BOM trovato: %s", bom_path)
    df_bom = pd.read_excel(bom_path, sheet_name=0)

    # Estrai SKU + Qty da BOM
    bom_qty = df_bom[['Numero componenti', 'Qtà comp. (UMC)']].copy()
    bom_qty.columns = ['SKU', 'Quantity_Per_Bike']

    # Aggrega per SKU (stesso SKU può apparire più volte con qty diverse)
    # Prendi max qty
    bom_qty_unique = bom_qty.groupby('SKU', as_index=False).agg({
        'Quantity_Per_Bike': 'max'
    })
    log.info("BOM SKU univoci: %d", len(bom_qty_unique))

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

    # 8. Riordina colonne (Freq_mesi/Supply_cycle: presenti se calcolati, servono al rolling)
    _base_cols = ['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                  'Quantity_Per_Bike', 'Description']
    _extra_cols = [c for c in ('Freq_mesi', 'Supply_cycle') if c in lead_times.columns]
    lead_times = lead_times[_base_cols + _extra_cols]

    log.info("[OK] Lead times caricati: %d SKU", len(lead_times))
    log.info("Non-null Lead_Time_Months: %d", lead_times['Lead_Time_Months'].notna().sum())

    # 9. Aggiungi foglio 'input_simulatore' al debug file
    try:
        debug_path = logistica_dir / 'residual_lead_time_debug.xlsx'
        if debug_path.exists():
            with pd.ExcelWriter(debug_path, engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                lead_times.to_excel(writer, sheet_name='input_simulatore', index=False)
            log.info("[DEBUG] Foglio 'input_simulatore' aggiunto: %s", debug_path)
    except Exception as _e:
        log.warning("Impossibile aggiungere foglio simulatore: %s", _e)

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
    log.info("[PRE-NORM] Inizio pre-normalizzazione veloce...")
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
    log.info("[PRE-NORM] Completato in %.2fs", elapsed)

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
    except Exception:
        return (9999, 99)




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
        log.warning("NESSUN RISULTATO DA SALVARE!")
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

    log.info("[OK] File aggregato salvato: %s", filename)
    log.info("Percentile safety stock: p%d", percentile)
    log.info("Righe: %d", len(results_output))
    log.info("Colonne: %s", list(results_output.columns))

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
        log.warning("NESSUN RISULTATO MENSILE DA SALVARE!")
        return pd.DataFrame()

    # Verifica presenza colonna monthly_stats
    if 'monthly_stats' not in results_df.columns:
        log.warning("Colonna 'monthly_stats' non presente: output mensile non disponibile")
        return pd.DataFrame()

    # ------------------------------------------------------------------
    # Costruisci l'elenco GLOBALE dei mesi in ordine cronologico,
    # raccogliendo tutti i mesi presenti in qualsiasi SKU.
    # Serve per garantire che ogni riga abbia le stesse colonne.
    # ------------------------------------------------------------------
    # Perf: itera solo la colonna 'monthly_stats' invece di iterrows()
    all_months_set = set()
    for ms in results_df['monthly_stats']:
        if ms:
            all_months_set.update(ms.keys())

    all_months_sorted = sorted(all_months_set, key=parse_month_to_sortable)

    monthly_rows = []

    # Perf: to_dict('records') ~10x più veloce di iterrows() per molti SKU
    for row in results_df.to_dict('records'):
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
        log.warning("Nessuna riga mensile generata")
        return pd.DataFrame()

    df_monthly = pd.DataFrame(monthly_rows)

    # Ordina per SKU
    df_monthly = df_monthly.sort_values('SKU').reset_index(drop=True)

    df_monthly.to_excel(filename, index=False)

    log.info("[OK] File mensile salvato: %s", filename)
    log.info("Righe: %d", len(df_monthly))
    # Stampa i mesi trovati
    mesi_cols = [c for c in df_monthly.columns if c.startswith('mean_demand_')]
    log.info("Mesi: %s", [c.replace('mean_demand_', '') for c in mesi_cols])

    return df_monthly


def match_skus_ultra_optimized(sku_catalog: pd.DataFrame,
                               lead_times: pd.DataFrame,
                               simulation_output: pd.DataFrame,
                               column_mapping: Dict[str, str],
                               n_runs: int,
                               percentile: int = 99) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Esegue il matching SKU vs configurazioni MC e calcola statistiche.

    Pre-normalizza l'output simulazione, costruisce la matrice numpy int16
    (con fallback memmap per dataset > 4 GB), prepara le righe SKU espandendo
    eventuali modelli combinati (es. 'MSV4+MSV4S' -> entries separate) e per
    ogni SKU aggrega la domanda dei mesi di lead time e produce statistiche
    aggregate e mensili (mean, std, percentile, safety stock).

    Returns:
        (df_results, df_no_lt) — df_results = SKU con statistiche/safety stock;
        df_no_lt = SKU esclusi perche' senza lead time utile (LT mancante o <= 0),
        con colonna 'Motivo' (replica main_sku_v2 di SolMidTerm-ALL).
    """
    log.info("MATCHING SKU V2 (ULTRA-OPTIMIZED)")
    log.info("Memoria iniziale: %.1f MB", get_memory_usage_mb())

    # STEP 1: Pre-normalizza colonne simulazione
    normalized_df, norm_arrays_raw = prenormalize_simulation_output_fast(
        simulation_output, column_mapping
    )
    log.info("Memoria dopo pre-norm: %.1f MB", get_memory_usage_mb())

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
    log.info("[SETUP] Preparazione matrice numpy...")
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

    log.info("Colonne run: %d", len(run_columns))
    log.info("Mesi unici: %d", len(month_prefixes))

    # Stima preventiva memoria prima di allocare
    _est_gb = n_rows * len(run_columns) * 2 / (1024 ** 3)
    log.info("RAM stimata per data_matrix: ~%.1f GB (shape %d x %d x int16)",
             _est_gb, n_rows, len(run_columns))

    # Soglia per usare memmap (disco) invece della RAM: evita OOM su matrici grandi
    _MEMMAP_THRESHOLD_GB = 4.0
    _memmap_path = None

    if _est_gb > _MEMMAP_THRESHOLD_GB:
        # Matrice troppo grande per la RAM -> scrivi su file temporaneo (memmap)
        # numpy accede al memmap esattamente come un normale array: nessun cambio logico
        log.info("Matrice grande (%.1f GB > %.1f GB): uso memmap su disco anziche' RAM",
                 _est_gb, _MEMMAP_THRESHOLD_GB)
        _memmap_path = os.path.join(tempfile.gettempdir(), '_sku_data_matrix.dat')
        data_matrix = np.memmap(_memmap_path, dtype='int16', mode='w+',
                                shape=(n_rows, len(run_columns)))
        _CHUNK = 5000
        for _s in range(0, n_rows, _CHUNK):
            _e = min(_s + _CHUNK, n_rows)
            data_matrix[_s:_e] = (simulation_output.iloc[_s:_e][run_columns]
                                  .to_numpy(dtype=np.int16))
            if (_s // _CHUNK) % 20 == 0 and _s > 0:
                log.debug("%d/%d righe scritte su disco...", _e, n_rows)
        data_matrix.flush()
        log.info("Memmap pronto: %s", _memmap_path)
    else:
        data_matrix = simulation_output[run_columns].values.astype(np.int16, copy=False)

    log.info("Matrice shape: %s", data_matrix.shape)

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
    log.info("[SETUP] Preparazione dati SKU V2 (con espansione multi-modello)...")

    _sku_merged_all = sku_catalog.merge(
        lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                    'Quantity_Per_Bike', 'Description']],
        left_on='Numero componenti',
        right_on='SKU',
        how='left'
    )

    # Traccia SKU esclusi (LT mancante o zero) per reporting (replica main_sku_v2 ALL)
    _mask_no_lt = (
        _sku_merged_all['Lead_Time_Months_Ceil'].isna() |
        (_sku_merged_all['Lead_Time_Months_Ceil'] <= 0)
    )
    df_no_lt = (
        _sku_merged_all[_mask_no_lt][['Numero componenti', 'Lead_Time_Months', 'Description']]
        .rename(columns={'Numero componenti': 'SKU'})
        .drop_duplicates(subset=['SKU'])
        .reset_index(drop=True)
    )
    df_no_lt['Motivo'] = df_no_lt['Lead_Time_Months'].apply(
        lambda x: 'Lead time = 0'
        if (pd.notna(x) and float(x) == 0)
        else 'Non trovato nei dati residual'
    )
    log.info("SKU senza lead time utile: %d", len(df_no_lt))

    sku_with_lt = _sku_merged_all[~_mask_no_lt].copy()
    # Mantieni come float per supportare interpolazione lineare (0.25, 0.5, etc.)
    # sku_with_lt['Lead_Time_Months_Ceil'] = sku_with_lt['Lead_Time_Months_Ceil'].astype(int)
    sku_with_lt['Quantity_Per_Bike'] = sku_with_lt['Quantity_Per_Bike'].fillna(1.0)

    sku_groups = sku_with_lt.groupby('SKU')
    n_sku_groups = len(sku_groups)
    log.info("SKU da processare: %d", n_sku_groups)

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
    log.info("[MATCHING] Elaborazione %d SKU...", n_sku_groups)

    start_time = time.time()
    all_results = []

    for sku_idx, (sku_id, sku_data) in enumerate(zip(sku_ids, sku_data_list)):
        if sku_idx % 500 == 0 and sku_idx > 0:
            elapsed = time.time() - start_time
            rate = sku_idx / elapsed
            remaining = (n_sku_groups - sku_idx) / rate if rate > 0 else 0
            log.info("[%d/%d] %.1f SKU/s, ETA: %.1fmin",
                     sku_idx, n_sku_groups, rate, remaining / 60)

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
    log.info("[OK] Matching V2 completato!")
    log.info("Tempo: %.1fs (%.1fmin)", elapsed, elapsed / 60)
    log.info("Velocita': %.1f SKU/s", rate)
    log.info("SKU con match: %d", len(all_results))

    # Cleanup file memmap temporaneo (se usato)
    if _memmap_path is not None:
        del data_matrix
        try:
            os.remove(_memmap_path)
            log.info("File memmap temporaneo rimosso: %s", _memmap_path)
        except Exception:
            pass

    return pd.DataFrame(all_results), df_no_lt

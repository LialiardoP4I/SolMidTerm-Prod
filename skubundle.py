# -*- coding: utf-8 -*-
"""
================================================================================
SKUBUNDLE - VERSIONE 2 (ZERO HARDCODING)
================================================================================

Versione V2 di skubundle_OPTIMIZED.py.
Elimina tutti gli hardcoding presenti nella V1:

HARDCODING ELIMINATI:
---------------------
1. df.columns[4:7]   -> posizione hardcoded delle colonne modello
2. df.columns[148:]  -> posizione hardcoded delle colonne calcolate
3. ['MSV4', 'MSV4S', 'MSV4PP'] -> nomi modello hardcoded

COME IDENTIFICA LE COLONNE:
----------------------------
Il file Excel matrix_output_V2.xlsx ha 3 righe header:
  Riga 1 (xlsx) = Bundle del mapping Schema ('' per colonne calcolate e metadati)
  Riga 2 (xlsx) = Descrizione del mapping Schema ('' per colonne calcolate e metadati)
  Riga 3 (xlsx) = Nome leggibile della colonna (usato come header del DataFrame)

Tipi di colonne:
  • Metadati (prime 4 col): Numero componenti, Condizione_Originale,
    Condizione_Espansa, colonna condizione BOM -> escluse dall'output
  • Mapping (dalla Schema): bundle e/o desc non vuoti -> includono le colonne modello
  • Calcolate (dal TR): bundle='' E desc='' E posizione >= 4
    -> Fairing Color, Radar Pack, Touring Pack, ecc.
  • Modello: sottoinsieme delle colonne mapping il cui nome readable
    corrisponde a un codice modello (es: MSV4, MSV4S, MSV4PP)

SCALABILITÀ:
------------
- Aggiungere un nuovo modello: compare automaticamente se presente nel mapping
  (non è necessario modificare questo script)
- Aggiungere colonne calcolate: il TR è la fonte di verità, vengono rilevate
  automaticamente tramite il pattern bundle=''+desc=''
- Cambiare il prefisso modello: modifica il parametro model_prefix o
  model_col_names nella chiamata alla funzione

INPUT:
------
- matrix_output_V2.xlsx con foglio "Matrix" (header su riga 3)
  Generato da gestione_dipendenze_V2.py

OUTPUT:
-------
- tutte_righe_univoche_V2.xlsx con:
  - "Tutte_Righe_Univoche": configurazioni uniche con ID
  - "Riepilogo": statistiche di riduzione
  - "Colonne_Usate": dettaglio delle colonne selezionate (debug/tracciabilità)

@author: AntonioLiardo-SMOPS
"""

import re
import pandas as pd
import numpy as np
from pathlib import Path
import shutil


# ============================================================================
# HELPER: RIMOZIONE SUFFISSI PANDAS (.1 .2) E CONSOLIDAMENTO DUPLICATI
# ============================================================================

def _strip_and_consolidate(df_in: pd.DataFrame,
                            col_names_all: list,
                            raw_name_row: list) -> pd.DataFrame:
    """
    Rimuove i suffissi .1 .2 aggiunti da pandas per colonne con nome duplicato
    nel file Excel sorgente, poi consolida le coppie con lo stesso nome.

    Logica di consolidamento per colonne Yes/No:
      - 'Yes' prevale su 'No' / NaN (OR logico)
      - Garantisce che un componente associato a una caratteristica non venga
        perso per via di nomi duplicati nella matrice.

    Args:
        df_in:          DataFrame con possibili colonne .1/.2
        col_names_all:  Lista COMPLETA dei nomi pandas di df originale
        raw_name_row:   Lista COMPLETA dei nomi raw (riga Excel) di df originale

    Returns:
        DataFrame con nomi colonna puliti e senza duplicati.
    """
    # --- Costruisci mappa: pandas_name -> raw_name per le colonne presenti ---
    pname_to_raw = {}
    for pname, rname in zip(col_names_all, raw_name_row):
        pname_to_raw[pname] = rname

    # --- Rinomina le colonne di df_in usando i raw names ---
    # re.sub rimuove suffissi .1/.2/... che potrebbero essere rimasti letterali
    # nel file Excel sorgente (non generati da pandas ma scritti tali-quali)
    new_names = [re.sub(r'\.\d+$', '', pname_to_raw.get(c, c)) for c in df_in.columns]

    dup_before = {c for c in new_names if new_names.count(c) > 1}
    if not dup_before:
        # Nessun duplicato: rinomina e restituisci
        df_out = df_in.copy()
        df_out.columns = new_names
        return df_out

    print(f"\n  [FIX .1/.2] Colonne con nome duplicato trovate: {dup_before}")

    # --- Consolida colonne duplicate ---
    # Iteriamo per posizione per gestire correttamente più occorrenze
    unique_ordered = list(dict.fromkeys(new_names))   # ordine preservato
    consolidated = {}

    for col in unique_ordered:
        positions = [i for i, n in enumerate(new_names) if n == col]
        if len(positions) == 1:
            consolidated[col] = df_in.iloc[:, positions[0]].reset_index(drop=True)
        else:
            # OR logico generalizzato:
            #   - Qualsiasi valore "reale" (diverso da No/nan/vuoto) prevale
            #   - Tra più valori reali, 'Yes' prevale (bundle attivo)
            #   - Se nessun valore reale, resta 'No'
            _inactive = {'no', 'nan', '', 'not'}
            cols_series = [df_in.iloc[:, p].astype(str).str.strip()
                          for p in positions]
            n = len(df_in)
            merged = pd.Series(['No'] * n)
            for s in cols_series:
                for idx in range(n):
                    val = s.iloc[idx]
                    val_low = val.lower()
                    cur_low = merged.iloc[idx].lower()
                    if val_low in _inactive:
                        continue  # valore inattivo, non sovrascrive
                    if cur_low in _inactive:
                        merged.iloc[idx] = val  # primo valore reale trovato
                    elif val_low == 'yes':
                        merged.iloc[idx] = 'Yes'  # Yes vince su altri valori
            consolidated[col] = merged.reset_index(drop=True)
            print(f"    '{col}': {len(positions)} occorrenze -> consolidate (OR)")

    return pd.DataFrame(consolidated)


# ============================================================================
# COLLAPSE PACK COLOR COLUMNS
# ============================================================================

_PACK_BUNDLES_TARGET = {'touring pack', 'top case pack', 'touring pack rally'}

def _collapse_pack_color_columns(df: pd.DataFrame,
                                  bundle_row: list,
                                  desc_row: list,
                                  raw_name_row: list) -> pd.DataFrame:
    """
    Crea colonne valore collassate per i componenti di Touring Pack e Top Case Pack.

    Per ogni gruppo (Bundle, CT_desc) nei pack target:
      - Color variant (>1 opzione OPPURE "color variant" in desc):
          valore = descrizione colore attivo (raw_name_row)
      - Pack invariant (1 opzione):
          valore = "Yes" se presente

    Regola "Yes" fallback: se nessuna opzione del CT e' attiva per questa riga
    MA il pack e' comunque attivo (almeno una opzione QUALSIASI del pack e' attiva),
    il valore diventa "Yes" anziche' "No". Questo garantisce che le righe SKU
    pack-invariant (es. Central Stand) matchino correttamente le configurazioni
    simulate con touring pack attivo, indipendentemente dal colore panniers.

    Regola "vince color variant": se un CT ha un valore colore specifico, prevale su "Yes".

    Args:
        df:            DataFrame matrice (header=2, TUTTE le colonne incluse)
        bundle_row:    Lista nomi bundle per ogni colonna (riga 0 matrice)
        desc_row:      Lista descrizioni CT per ogni colonna (riga 1 matrice)
        raw_name_row:  Lista nomi readable per ogni colonna (riga 2 matrice,
                       senza suffissi .1/.2 di pandas)

    Returns:
        DataFrame con le colonne collassate (stesso indice di df)
    """
    from collections import defaultdict

    _inactive_vals = {'', 'nan', 'none', 'no', 'not'}

    # Raccoglie colonne per (bundle_norm, ct_desc) E per bundle (per pack-activity check)
    groups = defaultdict(list)          # (bundle_norm, ct_desc) -> [(col_idx, raw_name), ...]
    bundle_all_col_indices = defaultdict(list)  # bundle_norm -> [col_idx, ...]

    for i, (b, d, rn) in enumerate(zip(bundle_row, desc_row, raw_name_row)):
        b_str = str(b).strip()
        b_norm = b_str.lower()
        d_str = str(d).strip()

        if b_norm not in _PACK_BUNDLES_TARGET:
            continue
        if not b_str or b_str in ('nan', 'None', 'none'):
            continue
        if not d_str or d_str in ('nan', 'None', 'none'):
            continue

        groups[(b_norm, d_str)].append((i, str(rn).strip()))
        bundle_all_col_indices[b_norm].append(i)

    if not groups:
        return pd.DataFrame(index=df.index)

    n_rows = len(df)

    # Identifica colonne calcolate per ogni pack: bundle='' E desc='' nel header.
    # Queste sono le colonne finali Yes/No/Not prodotte da gestione_dipendenze_V2
    # e rappresentano la fonte autorevole per stabilire se un pack e' attivo.
    # Usarle al posto delle colonne raw previene falsi positivi quando un CT/CV
    # e' condiviso tra pack incompatibili (es: CT000051/CV000025 in V22E mappa
    # sia su 'Alu Top case pack' che su 'Top Case Pack': la colonna raw e' attiva
    # per entrambi ma la colonna calcolata riflette il discriminante corretto).
    calc_pack_col_indices = {}  # bundle_norm -> col_index_in_df
    for i, (b, d, rn) in enumerate(zip(bundle_row, desc_row, raw_name_row)):
        b_str = str(b).strip()
        d_str = str(d).strip()
        rn_lower = str(rn).strip().lower()
        is_empty_bundle = not b_str or b_str in ('nan', 'None', 'none')
        is_empty_desc   = not d_str or d_str in ('nan', 'None', 'none')
        if is_empty_bundle and is_empty_desc and rn_lower in _PACK_BUNDLES_TARGET:
            calc_pack_col_indices[rn_lower] = i

    # Pre-calcola per ogni (bundle_norm, row_idx) se il pack e' attivo.
    # Usa la colonna calcolata (autorevole) se disponibile, altrimenti
    # deriva l'attivita' dalle colonne raw del mapping (comportamento precedente).
    pack_active = {}  # (bundle_norm, row_idx) -> bool
    for b_norm, all_cols in bundle_all_col_indices.items():
        if b_norm in calc_pack_col_indices:
            # Colonna calcolata disponibile: usa Yes/No/Not gia' risolti
            calc_idx = calc_pack_col_indices[b_norm]
            for row_idx in range(n_rows):
                cell = df.iloc[row_idx, calc_idx]
                cell_str = str(cell).strip().lower() if pd.notna(cell) else ''
                pack_active[(b_norm, row_idx)] = cell_str not in _inactive_vals
        else:
            # Fallback: deriva attivita' dalle colonne raw del mapping
            valid_cols = [ci for ci in all_cols if ci < df.shape[1]]
            for row_idx in range(n_rows):
                active = False
                for ci in valid_cols:
                    cell = df.iloc[row_idx, ci]
                    if pd.notna(cell) and str(cell).strip().lower() not in _inactive_vals:
                        active = True
                        break
                pack_active[(b_norm, row_idx)] = active

    new_cols = {}  # col_name -> list of values

    for (bundle_norm, ct_desc), col_list in groups.items():
        n_options = len(col_list)
        is_cv = ('color variant' in ct_desc.lower()) or (n_options > 1)

        col_indices   = [ci for ci, _ in col_list]
        col_raw_names = [rn for _, rn in col_list]

        values = ['No'] * n_rows

        for row_idx in range(n_rows):
            specific_found = None
            for ci, rn in zip(col_indices, col_raw_names):
                if ci >= df.shape[1]:
                    continue
                cell = df.iloc[row_idx, ci]
                if pd.notna(cell) and str(cell).strip().lower() not in _inactive_vals:
                    specific_found = rn if is_cv else 'Yes'
                    break

            if specific_found is not None:
                values[row_idx] = specific_found
            elif pack_active.get((bundle_norm, row_idx), False):
                # Pack attivo ma nessuna opzione specifica di questo CT -> "Yes"
                values[row_idx] = 'Yes'
            # else: pack non attivo -> resta 'No'

        col_name = ct_desc  # nome colonna = descrizione CT

        if col_name in new_cols:
            # Merge: color variant prevale su Yes/No
            existing = new_cols[col_name]
            for i in range(n_rows):
                if existing[i] == 'No' and values[i] != 'No':
                    existing[i] = values[i]
                elif values[i] not in ('No', 'Yes') and existing[i] in ('No', 'Yes'):
                    existing[i] = values[i]  # color prevale su Yes
        else:
            new_cols[col_name] = values

    return pd.DataFrame(new_cols, index=df.index)


# ============================================================================
# FUNZIONE DI RILEVAMENTO COLONNE (cuore del zero-hardcoding)
# ============================================================================

def _rileva_colonne(file_path: str, model_col_names=None, model_prefix: str = 'MSV4'):
    """
    Legge le 3 righe header della matrice V2 e classifica ogni colonna.

    LOGICA DI CLASSIFICAZIONE:
    --------------------------
    1. Prime 4 colonne -> metadati (escluse)
    2. Colonne con bundle='' E desc='' E posizione >= 4 -> colonne calcolate (dal TR)
    3. Colonne rimanenti (posizione >= 4, non calcolate) -> colonne mapping (dal Schema)
    4. Tra le mapping: quelle il cui nome readable corrisponde a model_col_names
       oppure al prefisso model_prefix -> colonne modello

    PERCHÉ bundle='' E desc='' identifica le calcolate:
    In gestione_dipendenze_V2.py, load_calculated_cols_from_tr produce tuple
    ('', '', calc_name, search_key). Il salvataggio Excel scrive '' in riga 1 e
    riga 2 per queste colonne. Le colonne mapping del Schema hanno sempre almeno
    una descrizione caratteristica non vuota in riga 2.

    Args:
        file_path:      Path al file matrix_output_V2.xlsx
        model_col_names: Lista esplicita di nomi colonne modello (es: ['MSV4','MSV4S'])
                         Se None, usa model_prefix per rilevamento automatico
        model_prefix:   Prefisso per rilevamento automatico modelli (default 'MSV4')

    Returns:
        dict con chiavi:
          'col_num_componenti': nome colonna "Numero componenti"
          'col_modello':        lista nomi colonne modello
          'col_calcolate':      lista nomi colonne calcolate
          'col_mapping_altre':  lista nomi altre colonne mapping (non modello, non metadati)
    """
    N_METADATA = 4  # Prime 4 colonne sempre escluse

    # Leggi le 3 righe header senza applicare nessun header
    df_raw = pd.read_excel(file_path, sheet_name='Matrix', header=None, nrows=3)

    bundle_row = df_raw.iloc[0].fillna('').astype(str).str.strip()
    desc_row   = df_raw.iloc[1].fillna('').astype(str).str.strip()
    name_row   = df_raw.iloc[2].fillna('').astype(str).str.strip()

    n_cols = len(name_row)

    # -------------------------------------------------------------------------
    # Classifica ogni colonna
    # -------------------------------------------------------------------------
    is_calcolata = []
    for i in range(n_cols):
        if i < N_METADATA:
            is_calcolata.append(False)
        else:
            # Calcolata = bundle vuoto E desc vuoto (invariante di gestione_dipendenze_V2)
            is_calcolata.append(bundle_row.iloc[i] == '' and desc_row.iloc[i] == '')

    # Leggi il DataFrame con header=2 (riga 3 Excel = nomi colonne)
    df = pd.read_excel(file_path, sheet_name='Matrix', header=2)

    col_names = list(df.columns)          # Nomi effettivi dopo pandas (gestisce duplicati con .1 .2)
    raw_names = list(name_row)            # Nomi raw dalla riga 3 Excel (senza suffisso pandas)

    # -------------------------------------------------------------------------
    # Colonna "Numero componenti" = prima colonna
    # -------------------------------------------------------------------------
    col_num_componenti = col_names[0]

    # -------------------------------------------------------------------------
    # Colonne calcolate
    # -------------------------------------------------------------------------
    col_calcolate = [col_names[i] for i, calc in enumerate(is_calcolata) if calc]

    # -------------------------------------------------------------------------
    # Colonne mapping (non metadata, non calcolate)
    # -------------------------------------------------------------------------
    mapping_indices = [i for i in range(N_METADATA, n_cols) if not is_calcolata[i]]
    col_mapping_tutte = [col_names[i] for i in mapping_indices]
    raw_mapping_names = [raw_names[i] for i in mapping_indices]

    # -------------------------------------------------------------------------
    # Colonne modello: sottoinsieme delle mapping il cui nome readable
    # corrisponde a un codice modello
    # -------------------------------------------------------------------------
    if model_col_names is not None:
        # Lista esplicita fornita dall'utente
        model_names_lower = {m.lower() for m in model_col_names}
        col_modello = [col_names[i]
                       for i in mapping_indices
                       if raw_names[i].lower() in model_names_lower
                       or col_names[i].lower() in model_names_lower]
    else:
        # Rilevamento automatico tramite prefisso
        prefix_lower = model_prefix.lower()
        col_modello = [col_names[i]
                       for i in mapping_indices
                       if raw_names[i].lower().startswith(prefix_lower)
                       or col_names[i].lower().startswith(prefix_lower)]

    col_mapping_altre = [c for c in col_mapping_tutte if c not in col_modello]

    return {
        'df':                  df,
        'col_num_componenti':  col_num_componenti,
        'col_modello':         col_modello,
        'col_calcolate':       col_calcolate,
        'col_mapping_altre':   col_mapping_altre,
        'raw_name_row':        raw_names,
        'bundle_row':          list(bundle_row),
        'desc_row':            list(desc_row),
    }


# ============================================================================
# FUNZIONE PRINCIPALE
# ============================================================================

def crea_tutte_righe_univoche_v2(file_path: str,
                                  output_path: str = None,
                                  model_col_names=None,
                                  model_prefix: str = 'MSV4',
                                  exclude_always_no: bool = True,
                                  write_output: bool = True):
    """
    Crea il file tutte_righe_univoche_V2.xlsx dalla matrice V2.

    DIFFERENZE RISPETTO A V1 (skubundle_OPTIMIZED.py):
    --------------------------------------------------
    - Nessun indice colonna hardcoded (no df.columns[4:7] o df.columns[148:])
    - Nessun nome modello hardcoded (no ['MSV4','MSV4S','MSV4PP'])
    - Rilevamento automatico colonne calcolate dalle righe header Excel
    - Rilevamento automatico colonne modello dal prefisso (default: 'MSV4')
    - Opzione per escludere colonne calcolate sempre a 'No' (inutili per dedup)
    - Foglio di debug "Colonne_Usate" nell'output per tracciabilità

    Args:
        file_path:         Path al file matrix_output_V2.xlsx
        output_path:       Path per tutte_righe_univoche_V2.xlsx
        model_col_names:   Lista esplicita nomi colonne modello (es: ['MSV4','MSV4S'])
                           Se None, usa model_prefix per rilevamento automatico
        model_prefix:      Prefisso per rilevamento automatico modelli (default 'MSV4')
        exclude_always_no: Se True (default), esclude le colonne calcolate che sono
                           sempre 'No' (es: bundle Rally non ancora in Schema)
                           Riduce la dimensione dell'output senza perdere informazioni

    Returns:
        DataFrame con le righe univoche
    """
    print(f"\n=== SKUBUNDLE V2 — CREAZIONE RIGHE UNIVOCHE ===")
    print(f"Input:  {file_path}")
    print(f"Output: {output_path}")

    # -------------------------------------------------------------------------
    # FASE 1: Rilevamento dinamico delle colonne
    # -------------------------------------------------------------------------
    print(f"\n[1/5] Analisi struttura colonne...")
    info = _rileva_colonne(file_path, model_col_names=model_col_names,
                           model_prefix=model_prefix)

    df                   = info['df']
    col_num_componenti   = info['col_num_componenti']
    col_modello          = info['col_modello']
    col_calcolate        = info['col_calcolate']

    righe_originali = df.shape[0]
    print(f"  Righe originali:        {righe_originali}")
    print(f"  Colonne totali:         {df.shape[1]}")
    print(f"  Colonne modello ({len(col_modello)}):   {col_modello}")
    print(f"  Colonne calcolate ({len(col_calcolate)}): {col_calcolate}")

    # -------------------------------------------------------------------------
    # FASE 2: Filtra colonne calcolate sempre 'No' (opzionale)
    # -------------------------------------------------------------------------
    col_calcolate_usate = []
    col_calcolate_escluse = []

    for col in col_calcolate:
        vals = df[col].fillna('').astype(str)
        unique_vals = set(vals.unique()) - {'', 'nan'}
        if exclude_always_no and unique_vals == {'No'}:
            col_calcolate_escluse.append(col)
        elif exclude_always_no and not unique_vals:
            col_calcolate_escluse.append(col)
        else:
            col_calcolate_usate.append(col)

    if col_calcolate_escluse:
        print(f"\n  Colonne calcolate escluse (sempre 'No', {len(col_calcolate_escluse)}):")
        for c in col_calcolate_escluse:
            print(f"    - '{c}'")

    print(f"\n  Colonne calcolate usate per deduplicazione ({len(col_calcolate_usate)}):")
    for c in col_calcolate_usate:
        print(f"    - '{c}'")

    # -------------------------------------------------------------------------
    # FASE 3: Selezione colonne rilevanti
    # -------------------------------------------------------------------------
    print(f"\n[2/5] Selezione colonne...")

    # Seleziona: Numero componenti + colonne modello + colonne calcolate (usate)
    colonne_da_esplodere = col_modello + col_calcolate_usate
    colonne_selezionate  = [col_num_componenti] + colonne_da_esplodere
    df_result = df[colonne_selezionate].copy()

    # -------------------------------------------------------------------------
    # FASE 3.5: Collapse colonne pack color/invariant
    # Per le colonne color-variant (panniers, top case): sovrascrive le colonne
    # pack esistenti ("Touring pack", "Top Case pack") con il valore colore.
    # Per le colonne pack-invariant (Central Stand, Grips): aggiunge come nuove.
    # Deve avvenire PRIMA della dedup per separare configurazioni per colore.
    # -------------------------------------------------------------------------
    print(f"\n[2b/5] Collapse colonne pack color (Touring Pack, Top Case Pack)...")
    _pack_color_df = _collapse_pack_color_columns(
        df,
        bundle_row=info['bundle_row'],
        desc_row=info['desc_row'],
        raw_name_row=info['raw_name_row'],
    )
    if not _pack_color_df.empty:
        _pack_color_df = _pack_color_df.reset_index(drop=True)
        df_result = df_result.reset_index(drop=True)

        # Mappa CT description -> (lista candidati target_lower, colonna_finale)
        # target: nomi possibili nel matrix_output (gestisce sia old che new matrix).
        # colonna_finale: nome desiderato nell'output (deve coincidere con df_wide/TR).
        #
        # VECCHIO matrix: "Touring pack (color variant)" come colonna calcolata
        # NUOVO  matrix:  "Touring pack"  (dopo rigenerazione BOM con TR aggiornato)
        # Il codice prova entrambi in ordine.
        _COLOR_CT_TO_PACK_COL = {
            # ct_desc_lower              -> (candidati target lower,                       finale TR)
            'panniers covers (color variant)': (('touring pack', 'touring pack (color variant)'), 'Touring pack'),
            'top case cover (color variant)':  (('top case pack',),                               'Top Case Pack'),
            'panniers covers':                 (('touring pack rally',),                          'Touring pack Rally'),
            'valigie laterali':                (('touring pack', 'touring pack (color variant)'), 'Touring pack'),
        }
        _df_result_cols_lower = {c.lower(): c for c in df_result.columns}

        # Le CT color-variant sovrascrivono la colonna pack con il valore colore.
        # "Valigie Laterali" (RS) e' pack-invariant: sovrascrive con "Yes"
        # (nessuna scelta colore -> il matching usera' l'esatto "Yes" == "Yes").
        # Tutti gli altri CT del bundle (Central Stand, Grips, ecc.) vengono
        # ignorati: la loro presenza e' gia' catturata dal valore non-"No" nella
        # colonna pack.
        _ct_cols_ignored = []
        _renames = {}   # { colonna_attuale: colonna_finale }

        for _ct_col in _pack_color_df.columns:
            _ct_lower = _ct_col.lower()
            if _ct_lower in _COLOR_CT_TO_PACK_COL:
                _target_candidates, _final_name = _COLOR_CT_TO_PACK_COL[_ct_lower]
                # Prova i candidati in ordine finche' uno viene trovato
                _target_actual = None
                for _cand in _target_candidates:
                    _target_actual = _df_result_cols_lower.get(_cand, None)
                    if _target_actual is not None:
                        break
                if _target_actual is not None:
                    import numpy as np
                    _new_vals = _pack_color_df[_ct_col].values
                    _old_vals = df_result[_target_actual].astype(str).str.strip().values
                    # Preserva "Not" originale da gestione_dipendenze:
                    # se il colore-collapse darebbe "No" ma la colonna calcolata
                    # aveva "Not" (riga "Not X Pack"), mantieni "Not".
                    _merged = np.where(
                        (_new_vals == 'No') & (_old_vals == 'Not'),
                        'Not',
                        _new_vals
                    )
                    _n_not_preserved = int(((_new_vals == 'No') & (_old_vals == 'Not')).sum())
                    df_result[_target_actual] = _merged
                    print(f"  '{_ct_col}' -> sovrascritto '{_target_actual}' "
                          f"(Not preservati: {_n_not_preserved})")
                    # Segna rinomina se il nome attuale != nome finale desiderato
                    if _target_actual.lower() != _final_name.lower():
                        _renames[_target_actual] = _final_name
                else:
                    print(f"  [WARN] Colonna pack per '{_ct_col}' non trovata "
                          f"(candidati: {_target_candidates}) -> ignorato")
            else:
                _ct_cols_ignored.append(_ct_col)

        # Rinomina colonne (es. 'Touring pack (color variant)' -> 'Touring pack')
        # per allineare i nomi a df_wide (simulazione)
        if _renames:
            df_result = df_result.rename(columns=_renames)
            for old, new in _renames.items():
                print(f"  Rinominata colonna: '{old}' -> '{new}'")

        if _ct_cols_ignored:
            print(f"  CT bundle ignorati (non usati per matching): {_ct_cols_ignored}")
    else:
        print(f"  Nessuna colonna pack trovata nei pack target.")

    # -------------------------------------------------------------------------
    # FASE 4: Creazione colonna Model combinata
    # -------------------------------------------------------------------------
    print(f"[3/5] Creazione colonna Model...")

    # Una riga è "attiva" per un modello se il valore è non-vuoto e non 'No'
    # (stesso criterio della V1)
    masks = {}
    for model_col in col_modello:
        masks[model_col] = (
            df_result[model_col].notna() &
            (df_result[model_col].astype(str) != '') &
            (df_result[model_col].astype(str) != 'No')
        )

    # Costruzione vettorizzata della stringa Model
    # Usa numpy arrays per performance (identico a V1)
    n_rows = len(df_result)
    mask_arrays = {col: masks[col].values for col in col_modello}

    model_values = []
    for i in range(n_rows):
        parts = [col for col in col_modello if mask_arrays[col][i]]
        model_values.append('+'.join(parts) if parts else 'Nessuno')

    df_result['Model'] = model_values

    # Rimuovi le colonne modello originali (ora ridondanti)
    cols_to_drop = [col for col in col_modello if col in df_result.columns]
    df_result = df_result.drop(columns=cols_to_drop)

    # Sposta 'Model' come seconda colonna (dopo Numero componenti)
    cols_ordered = [col_num_componenti, 'Model'] + \
                   [c for c in df_result.columns if c not in [col_num_componenti, 'Model']]
    df_result = df_result[cols_ordered]

    # -------------------------------------------------------------------------
    # FIX V2: Strip suffissi .1/.2 e consolida colonne duplicate
    # -------------------------------------------------------------------------
    # Deve avvenire DOPO la rimozione delle colonne modello (che potrebbero avere
    # suffissi propri) e PRIMA della deduplicazione (altrimenti righe identiche
    # che differivano solo per la colonna duplicata non vengono unite).
    df_result = _strip_and_consolidate(
        df_result,
        col_names_all=list(df.columns),
        raw_name_row=info['raw_name_row']
    )
    # Aggiorna col_calcolate_usate ai nuovi nomi (senza .1/.2) e rimuovi
    # eventuali duplicati che ora puntano allo stesso nome raw.
    _p2r = dict(zip(list(df.columns), info['raw_name_row']))
    _seen_cal: set = set()
    _cal_new: list = []
    for _c in col_calcolate_usate:
        _new_c = _p2r.get(_c, _c)
        if _new_c not in _seen_cal:
            _seen_cal.add(_new_c)
            _cal_new.append(_new_c)
    col_calcolate_usate = _cal_new
    # Aggiorna anche col_calcolate_escluse (usata nei fogli debug/riepilogo)
    _seen_excl: set = set()
    _excl_new: list = []
    for _c in col_calcolate_escluse:
        _new_c = _p2r.get(_c, _c)
        if _new_c not in _seen_excl:
            _seen_excl.add(_new_c)
            _excl_new.append(_new_c)
    col_calcolate_escluse = _excl_new
    col_num_componenti  = df_result.columns[0]   # sempre prima colonna (invariante)

    # -------------------------------------------------------------------------
    # FASE 5: Deduplicazione
    # -------------------------------------------------------------------------
    print(f"[4/5] Deduplicazione...")
    print(f"  Righe prima  di drop_duplicates: {len(df_result)}")

    df_result = df_result.drop_duplicates()

    n_univoche = len(df_result)
    n_rimossi  = righe_originali - n_univoche
    pct        = (1 - n_univoche / righe_originali) * 100 if righe_originali > 0 else 0

    print(f"  Righe dopo   drop_duplicates:    {n_univoche}")
    print(f"  Duplicati rimossi:               {n_rimossi} ({pct:.2f}%)")

    # -------------------------------------------------------------------------
    # FASE 5.5: Espansione righe "(Ducati Red+PP)" nei pack columns
    # La voce combinata "(Ducati Red+PP)" nella matrice rappresenta due ZCOL
    # distinte ("Ducati Red" e "Pikes Peak Livery"). Ogni riga con questo valore
    # viene espansa in due righe separate per garantire il matching 1:1 con la
    # simulazione (che produce "Ducati Red" o "Pikes Peak Livery" come valori).
    # -------------------------------------------------------------------------
    _PACK_COMBINED_EXPANSIONS = [
        ('touring pack',   '(Ducati Red+PP)', ['Ducati Red', 'Pikes Peak Livery']),
        ('top case pack',  '(Ducati Red+PP)', ['Ducati Red', 'Pikes Peak Livery']),
    ]
    _df_res_cols_lower = {c.lower(): c for c in df_result.columns}
    _expansion_done = False
    for _col_lower, _combined_val, _split_vals in _PACK_COMBINED_EXPANSIONS:
        _col_actual = _df_res_cols_lower.get(_col_lower, None)
        if _col_actual is None:
            continue
        _mask = df_result[_col_actual] == _combined_val
        _n_combined = int(_mask.sum())
        if _n_combined == 0:
            continue
        _rows_combined = df_result[_mask].copy()
        df_result = df_result[~_mask].copy()
        _expanded = pd.concat(
            [_rows_combined.assign(**{_col_actual: sv}) for sv in _split_vals],
            ignore_index=True,
        )
        df_result = pd.concat([df_result, _expanded], ignore_index=True)
        print(f"  [{_col_actual}] {_n_combined} righe '{_combined_val}' "
              f"-> {_n_combined * len(_split_vals)} righe (split: {_split_vals})")
        _df_res_cols_lower = {c.lower(): c for c in df_result.columns}  # aggiorna dopo concat
        _expansion_done = True
    if not _expansion_done:
        print("  Nessuna riga '(Ducati Red+PP)' da espandere.")

    # Dedup post-espansione (l'espansione potrebbe creare righe gia' presenti)
    if _expansion_done:
        _pre = len(df_result)
        df_result = df_result.drop_duplicates()
        _post = len(df_result)
        if _pre != _post:
            print(f"  Dedup post-espansione: {_pre} -> {_post} ({_pre - _post} duplicati rimossi)")

    # Se write_output=False, ritorna il df grezzo senza ID ne' salvataggio.
    # Usato quando si processano piu' file separatamente prima di concatenare.
    if not write_output:
        return df_result

    n_univoche = len(df_result)
    n_rimossi  = righe_originali - n_univoche
    pct        = (1 - n_univoche / righe_originali) * 100 if righe_originali > 0 else 0

    # Aggiungi ID progressivo come prima colonna
    df_result.insert(0, 'ID', range(1, n_univoche + 1))

    # -------------------------------------------------------------------------
    # FASE 6: Salvataggio
    # -------------------------------------------------------------------------
    print(f"[5/5] Salvataggio...")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # Foglio principale: righe univoche
        df_result.to_excel(writer, sheet_name='Tutte_Righe_Univoche', index=False)

        # Foglio riepilogo statistiche
        riepilogo = pd.DataFrame({
            'Metrica': [
                'Righe totali originali',
                'Righe univoche',
                'Duplicati rimossi',
                'Riduzione %',
                'Colonne modello',
                'Colonne calcolate usate',
                'Colonne calcolate escluse (sempre No)',
            ],
            'Valore': [
                righe_originali,
                n_univoche,
                n_rimossi,
                f"{pct:.2f}%",
                ', '.join(col_modello),
                ', '.join(col_calcolate_usate),
                ', '.join(col_calcolate_escluse) if col_calcolate_escluse else 'Nessuna',
            ]
        })
        riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)

        # Foglio debug: dettaglio colonne usate (tracciabilità zero-hardcoding)
        colonne_debug = []
        for col in df_result.columns:
            if col == 'ID':
                tipo = 'ID progressivo'
            elif col == col_num_componenti:
                tipo = 'Metadato'
            elif col == 'Model':
                tipo = f'Modello combinato (da: {", ".join(col_modello)})'
            elif col in col_calcolate_usate:
                tipo = 'Calcolata (dal TR)'
            else:
                tipo = 'Sconosciuto'
            colonne_debug.append({'Colonna': col, 'Tipo': tipo})

        for col in col_calcolate_escluse:
            colonne_debug.append({
                'Colonna': col,
                'Tipo': 'Calcolata — ESCLUSA (sempre No)'
            })

        pd.DataFrame(colonne_debug).to_excel(writer, sheet_name='Colonne_Usate', index=False)

    print(f"\n=== RISULTATO SKUBUNDLE V2 ===")
    print(f"Salvato: {output_path}")
    print(f"Dimensioni: {n_univoche} righe × {len(df_result.columns)} colonne")
    print(f"\nPrime 3 righe:")
    print(df_result.head(3).to_string())

    return df_result


# ============================================================================
# DISCOVERY: RILEVA AUTOMATICAMENTE MATRIX_OUTPUT FILES
# ============================================================================

def _get_available_matrix_files(input_dir: str = ".\\Input") -> list:
    """
    Scopre automaticamente tutti i file matrix_output_*.xlsx nella directory Input.
    Esclude file con '_SA_' nel nome (Stock Alignment — trattati separatamente).
    Ordina per nome per coerenza.

    Returns:
        Lista di Path ordinati alfabeticamente
    """
    input_path = Path(input_dir)
    matrix_files = sorted(input_path.glob("matrix_output_*.xlsx"))
    # Esclude SA (Safety Stock Alignment — gestiti da crea_catalogo_sa.py)
    return [str(f) for f in matrix_files if "_SA_" not in f.name]


def _cleanup_safety_stock_files(input_dir: str = ".\\Input") -> None:
    """
    Elimina tutti i file matrix_output_*.xlsx e tutte_righe_univoche_*.xlsx
    (esclude SA che sono trattati separatamente).
    Eseguito all'inizio di __main__ per evitare file stali.
    """
    input_path = Path(input_dir)

    # Elimina matrix_output (esclude SA)
    print("\n[CLEANUP] Rimozione matrix_output precedenti (non-SA)...")
    for f in input_path.glob("matrix_output_*.xlsx"):
        if "_SA_" not in f.name:
            try:
                f.unlink()
                print(f"  Rimosso: {f.name}")
            except Exception as e:
                print(f"  [WARN] Impossibile rimuovere {f.name}: {e}")

    # Elimina tutte_righe_univoche (esclude SA)
    print("[CLEANUP] Rimozione tutte_righe_univoche precedenti (non-SA)...")
    for f in input_path.glob("tutte_righe_univoche_*.xlsx"):
        if "_SA_" not in f.name:
            try:
                f.unlink()
                print(f"  Rimosso: {f.name}")
            except Exception as e:
                print(f"  [WARN] Impossibile rimuovere {f.name}: {e}")


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    # =========================================================================
    # CLEANUP: Elimina file stali
    # =========================================================================
    _cleanup_safety_stock_files(input_dir=".\\Input")

    # =========================================================================
    # DISCOVERY: Rileva automaticamente tutti i matrix_output_*.xlsx
    # =========================================================================
    # Prima erano hardcoded solo V21E, V22E, V24T — ora include PANI e qualsiasi
    # nuovo modello aggiunto in futuro.
    _input_files = _get_available_matrix_files(input_dir=".\\Input")
    _output_path = ".\\Input\\tutte_righe_univoche_V2.xlsx"

    if not _input_files:
        print("\n[ERROR] Nessun file matrix_output_*.xlsx trovato in ./Input/")
        exit(1)

    print(f"\n[INFO] File matrix_output rilevati ({len(_input_files)}):")
    for _fp in _input_files:
        print(f"  - {Path(_fp).name}")

    _dfs = []
    _righe_originali_totale = 0
    for _fp in _input_files:
        print(f"\n{'='*60}")
        print(f"Processo: {_fp}")
        _df_partial = crea_tutte_righe_univoche_v2(
            file_path=_fp,
            write_output=False,
        )
        _righe_originali_totale += len(_df_partial)
        _dfs.append(_df_partial)
        print(f"  -> {len(_df_partial)} righe univoche per questo file")

    print(f"\n{'='*60}")
    print(f"Concatenazione di {len(_dfs)} file...")
    _df_all = pd.concat(_dfs, ignore_index=True)

    # Colonne mancanti (es. Touring pack Rally assente in V21E/V22E) -> 'No'
    _df_all = _df_all.fillna('No')

    # Dedup globale tra tutti i file
    _pre_dedup = len(_df_all)
    _df_all = _df_all.drop_duplicates()
    _post_dedup = len(_df_all)
    if _pre_dedup != _post_dedup:
        print(f"  Dedup globale: {_pre_dedup} -> {_post_dedup} "
              f"({_pre_dedup - _post_dedup} duplicati rimossi)")

    # ID progressivo
    _df_all.insert(0, 'ID', range(1, len(_df_all) + 1))

    # Salvataggio
    print(f"\nSalvataggio in: {_output_path}")
    with pd.ExcelWriter(_output_path, engine='openpyxl') as writer:
        _df_all.to_excel(writer, sheet_name='Tutte_Righe_Univoche', index=False)
        _riepilogo = pd.DataFrame({
            'Metrica': [
                'File processati',
                'Righe totali (pre-dedup globale)',
                'Righe univoche finali',
                'Duplicati rimossi (dedup globale)',
                'Colonne totali',
            ],
            'Valore': [
                ', '.join(_input_files),
                _pre_dedup,
                _post_dedup,
                _pre_dedup - _post_dedup,
                len(_df_all.columns),
            ]
        })
        _riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)

    print(f"\n=== RISULTATO FINALE ===")
    print(f"Salvato: {_output_path}")
    print(f"Dimensioni: {len(_df_all)} righe x {len(_df_all.columns)} colonne")

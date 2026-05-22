# -*- coding: utf-8 -*-
"""
================================================================================
PROCESSOR: TUTTE LE COMBINAZIONI BOM × MAPPATURA (AGGIORNATO)
================================================================================

Versione AGGIORNATA con QUALITY CONTROL (One Supermodel Edition)

Processa tutte le 3 combinazioni con validazione input:
  1. BOM150_V21E + Mappatura_Unificata
  2. BOM150_V22E + Mappatura_Unificata
  3. BOM150_V24T + Mappatura_Unificata

QUALITY CHECKS INTEGRATI (nel processo_excel_to_matrix_v2):
  ✓ Residual/Quantity validation (passato a gestione_dipendenze_V2)
  ✓ Price data validation (passato a gestione_dipendenze_V2)
  ✓ CT/CV traceability check (passato a gestione_dipendenze_V2)
  ✓ TR characteristic options alignment
  ✓ Market-specific condition filtering (States.xlsx)
  ✓ Active CV sets (TR > 0) con filter_untraced=True
  ✓ Pre-process validation (globale: residual, price, TR options)
  ✓ Per-combination quality checks (CT/CV traceability)

Output finale: UN UNICO file tutte_righe_univoche_V2.xlsx consolidato
             + Quality summary per ogni combinazione

Dipendenze: gestione_dipendenze_V2.py (versione aggiornata One Supermodel)

@author: AntonioLiardo-SMOPS (Adattato alle best practices One Supermodel)
"""

import os
import glob
import time
import shutil
import numpy as np
import pandas as pd
from pathlib import Path

# Import dei moduli che contengono le funzioni
from gestione_dipendenze_V2 import (
    process_excel_to_matrix_v2,
    enrich_bom_version,           # Step 1 integrato: BOM grezzo → DataFrame arricchito
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
)
from skubundle_OPTIMIZED_V2 import crea_tutte_righe_univoche_v2


# ============================================================================
# CONFIGURAZIONE
# ============================================================================

INPUT_DIR = Path(".") / "Input"
OUTPUT_DIR = Path(".") / "Input"

# File per quality checks
PREZZI_FILE = INPUT_DIR / "Prezzi.xlsx"
STATES_FILE = INPUT_DIR / "States.xlsx"

_MAPPATURA_UNIFICATA = INPUT_DIR / "Mappatura_Unificata.xlsx"

# Leggi dinamicamente le cartelle dei modelli
MODEL_DIR = INPUT_DIR / "MODEL"

def get_available_models():
    """
    Restituisce un dizionario con i modelli disponibili e i loro percorsi.
    """
    models = {}
    if MODEL_DIR.exists():
        for model_folder in MODEL_DIR.iterdir():
            if model_folder.is_dir():
                model_name = model_folder.name
                mappatura_path = model_folder / f"Mappatura_{model_name}.xlsx"
                if mappatura_path.exists():
                    models[model_name] = {
                        "in_dir": model_folder,
                        "mappatura": mappatura_path
                    }
    return models

# Ottieni i modelli disponibili
AVAILABLE_MODELS = get_available_models()

# Costruisci le combinazioni dinamicamente
COMBINAZIONI = [
    {
        "nome": model_name,
        "in_dir": model_info["in_dir"],
        "mappatura": model_info["mappatura"],
        "tr": INPUT_DIR / "TR TOTALV21E - Copia.xlsx",
    }
    for model_name, model_info in AVAILABLE_MODELS.items()
]

# Costruisci la lista delle mappature originali dinamicamente
MAPPATURE_ORIGINALI = [model_info["mappatura"] for model_info in AVAILABLE_MODELS.values()]


def _normalize_bundle(name) -> str:
    """
    Normalizza il nome di un Bundle: rimuove spazi ridondanti e uniforma le
    maiuscole/minuscole in Title Case.

    Convenzione:
      - Strip leading/trailing whitespace + collasso spazi interni
      - Title Case su ogni parola (Python str.title gestisce anche separatori -/_) 
      - Valori vuoti/NaN -> stringa vuota
    """
    if name is None or (isinstance(name, float) and np.isnan(name)):
        return ''
    s = ' '.join(str(name).split())  # strip + collasso spazi multipli
    if not s or s.lower() in ('nan', 'none'):
        return ''
    return s.title()


def build_unified_mapping(source_files=None, output_path=None) -> pd.DataFrame:
    """
    Legge le mappature originali (source_files), le unisce con dedup/merge
    case-insensitive sui bundle name e salva il risultato in output_path.

    Regole di merge per coppie (Bundle_norm, Codice) presenti in più BOM:
      - Opzioni identiche -> una sola riga (deduplicata)
      - Opzioni diverse, unione ≤ 19 -> riga unica con tutte le opzioni unite
      - Opzioni diverse, unione > 19 -> righe separate

    Args:
        source_files: lista di Path delle mappature originali (default: MAPPATURE_ORIGINALI)
        output_path:  Path dove salvare il file unificato (default: _MAPPATURA_UNIFICATA)
    """
    if source_files is None:
        source_files = MAPPATURE_ORIGINALI
    if output_path is None:
        output_path = _MAPPATURA_UNIFICATA

    files = [str(f) for f in source_files if Path(f).exists()]
    if not files:
        raise FileNotFoundError(f"Nessuna mappatura originale trovata: {source_files}")

    print(f"\n[MAPPATURA UNIFICATA] Leggo {len(files)} file: {[Path(f).name for f in files]}")

    opz_cols  = [f'Opz{i}'  for i in range(1, 20)]
    desc_cols = [f'Desc{i}' for i in range(1, 20)]
    # IMPORTANTE: le colonne devono essere interleaved (Opz1,Desc1,Opz2,Desc2,...)
    # perché load_mapping_from_file legge coppie consecutive col_idx / col_idx+1
    interleaved_cols = [col for i in range(1, 20) for col in (f'Opz{i}', f'Desc{i}')]
    schema_cols = ['Bundle', 'Codice caratteristica', 'Descrizione caratteristica'] + interleaved_cols

    # -----------------------------------------------------------------
    # 1. Leggi e normalizza
    # -----------------------------------------------------------------
    dfs = []
    for f in files:
        df = pd.read_excel(f, sheet_name='Schema')
        # Assicura che le colonne esistano
        for c in schema_cols:
            if c not in df.columns:
                df[c] = np.nan
        df = df[schema_cols].copy()
        df['Bundle'] = df['Bundle'].apply(_normalize_bundle)
        df['_src'] = Path(f).name
        dfs.append(df)

    all_df = pd.concat(dfs, ignore_index=True)

    # Rimuovi righe prive di Codice caratteristica
    all_df = all_df[all_df['Codice caratteristica'].notna()].copy()
    all_df['Codice caratteristica'] = all_df['Codice caratteristica'].astype(str).str.strip()

    # -----------------------------------------------------------------
    # 2. Dedup / merge per (Bundle_norm_lower, Codice)
    # -----------------------------------------------------------------
    all_df['_bkey'] = all_df['Bundle'].str.lower() + '|||' + all_df['Codice caratteristica'].str.lower()

    result_rows = []
    seen_bkeys = {}  # bkey -> index nella result_rows

    for _, row in all_df.iterrows():
        bkey = row['_bkey']
        row_opz  = [row[c] for c in opz_cols]
        row_desc = [row[c] for c in desc_cols]

        def _clean_vals(vals):
            return [v for v in vals if pd.notna(v) and str(v).strip() not in ('', 'nan')]

        if bkey not in seen_bkeys:
            seen_bkeys[bkey] = len(result_rows)
            result_rows.append({
                'Bundle': row['Bundle'],
                'Codice caratteristica': row['Codice caratteristica'],
                'Descrizione caratteristica': row['Descrizione caratteristica'],
                **{f'Opz{i+1}':  row_opz[i]  for i in range(19)},
                **{f'Desc{i+1}': row_desc[i] for i in range(19)},
            })
        else:
            # Controlla se si può fare il merge
            idx = seen_bkeys[bkey]
            existing = result_rows[idx]
            ex_opz  = _clean_vals([existing[f'Opz{i+1}']  for i in range(19)])
            ex_desc = _clean_vals([existing[f'Desc{i+1}'] for i in range(19)])
            new_opz  = _clean_vals(row_opz)
            new_desc = _clean_vals(row_desc)

            # Opzioni da aggiungere (non già presenti)
            ex_opz_set = set(str(v) for v in ex_opz)
            to_add_opz  = [v for v in new_opz  if str(v) not in ex_opz_set]
            to_add_desc = [new_desc[new_opz.index(v)] if v in new_opz and new_opz.index(v) < len(new_desc) else ''
                             for v in to_add_opz]

            if not to_add_opz:
                # Identiche -> già presente, salta
                continue

            merged_opz  = ex_opz  + to_add_opz
            merged_desc = ex_desc + to_add_desc

            if len(merged_opz) <= 19:
                # Merge in place
                for i in range(19):
                    existing[f'Opz{i+1}']  = merged_opz[i]  if i < len(merged_opz)  else np.nan
                    existing[f'Desc{i+1}'] = merged_desc[i] if i < len(merged_desc) else np.nan
            else:
                # Non rientra -> riga separata (con nuovo bkey disambiguato)
                new_bkey = bkey + f'__{len(result_rows)}'
                seen_bkeys[new_bkey] = len(result_rows)
                result_rows.append({
                    'Bundle': row['Bundle'],
                    'Codice caratteristica': row['Codice caratteristica'],
                    'Descrizione caratteristica': row['Descrizione caratteristica'],
                    **{f'Opz{i+1}':  row_opz[i]  for i in range(19)},
                    **{f'Desc{i+1}': row_desc[i] for i in range(19)},
                })

    final_df = pd.DataFrame(result_rows, columns=schema_cols)

    print(f"[MAPPATURA UNIFICATA] Righe totali input: {len(all_df)} | "
          f"Righe unificate: {len(final_df)}")

    # -----------------------------------------------------------------
    # 3. Salva
    # -----------------------------------------------------------------
    if output_path is not None:
        try:
            with pd.ExcelWriter(str(output_path), engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='Schema', index=False)
            print(f"[MAPPATURA UNIFICATA] Salvato: {Path(output_path).name}")
        except PermissionError:
            raise PermissionError(
                f"Impossibile scrivere '{output_path}': il file e' aperto in Excel. "
                f"Chiudilo e rilancia lo Step 1."
            )

    return final_df


def run_quality_checks() -> tuple:
    """
    Esegue quality checks su input file comuni a tutte le combinazioni.

    Returns:
        (residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set)
    """
    print(f"\n{'='*80}")
    print("QUALITY CHECKS - VALIDAZIONE INPUT FILE")
    print(f"{'='*80}")

    residual_issues = []
    price_issues = []
    ct_cv_issues = []
    tr_char_options = {}
    states_set = set()

    # Quality check residual/quantity: rimosso (file quantity e residual.xlsx eliminato)

    # Quality check prezzi
    if PREZZI_FILE.exists():
        try:
            prezzi_df = pd.read_excel(PREZZI_FILE)
            # price_issues = _check_price_quality(prezzi_df, 'Prezzo netto')
            print(f"  [INFO] Price check: {len(price_issues)} issues trovati")
        except Exception as e:
            print(f"  [WARN] Price check skipped: {e}")

    # Load states (market-specific filtering)
    if STATES_FILE.exists():
        states_set = _load_states_set(str(STATES_FILE))

    # Build TR char options (per allineamento nomi)
    try:
        tr_char_options = _build_tr_char_options(str(INPUT_DIR / "TR TOTALV21E - Copia.xlsx"))
    except Exception as e:
        print(f"  [WARN] TR char options: {e}")

    return residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set


def process_combination(combo: dict, tr_char_options: dict = None, filter_untraced: bool = True) -> pd.DataFrame:
    """
    Processa una singola combinazione BOM+Mappatura in un unico step integrato:
      1. Enrichment BOM grezzo → DataFrame arricchito (enrich_bom_version)
      2. Elaborazione DataFrame → matrice (process_excel_to_matrix_v2 con df_input)

    Args:
        combo: dict con chiavi 'nome', 'in_dir', 'mappatura', 'tr'

    Returns:
        DataFrame con le righe univoche della combinazione
    """
    nome          = combo["nome"]
    in_dir        = combo["in_dir"]
    mappatura_path = combo["mappatura"]
    tr_path       = combo["tr"]

    print(f"\n{'='*80}")
    print(f"PROCESSAMENTO COMBINAZIONE: {nome}")
    print(f"{'='*80}")
    print(f"  In_dir:    {in_dir}")
    print(f"  Mappatura: {mappatura_path.name}")
    print(f"  TR:        {tr_path.name}")

    # -------------------------------------------------------------------------
    # STEP 0: Enrichment BOM grezzo (ex bom_enrich.py)
    # -------------------------------------------------------------------------
    print(f"\n[0/2] Enrichment BOM grezzo per {nome}...")
    df_enriched = enrich_bom_version(nome, in_dir, save=True)

    # Quality check CT/CV (usa il df già in memoria, nessuna lettura file extra)
    try:
        mapping_df = load_mapping_from_file(str(mappatura_path))
        ct_cv_issues = []   # decommentare _check_ct_cv_traceability se necessario
    except Exception as e:
        print(f"  [WARN] CT/CV check skipped per {nome}: {e}")
        ct_cv_issues = []

    # -------------------------------------------------------------------------
    # STEP 1: Elaborazione DataFrame arricchito → matrix_output
    # -------------------------------------------------------------------------
    matrix_output_path = OUTPUT_DIR / f"matrix_output_{nome}.xlsx"

    print(f"\n[1/2] Generazione matrix_output_{nome}.xlsx...")
    process_excel_to_matrix_v2(
        df_input=df_enriched,           # DataFrame in memoria — nessuna lettura file
        mapping_file=str(mappatura_path),
        tr_file=str(tr_path),
        output_file=str(matrix_output_path),
        states_file=str(STATES_FILE) if STATES_FILE.exists() else None,
        prezzi_file=str(PREZZI_FILE) if PREZZI_FILE.exists() else None,
        filter_untraced=filter_untraced,
        bom_file=f"BOM_150{nome}.xlsx"
    )

    # Verifica che il file sia stato creato
    if not matrix_output_path.exists():
        raise FileNotFoundError(f"Matrix output non creato: {matrix_output_path}")

    file_size = matrix_output_path.stat().st_size
    print(f"[OK] Salvato: {matrix_output_path.name} ({file_size:,} bytes)")

    # Attendi che il file sia completamente scritto
    time.sleep(2)

    # -------------------------------------------------------------------------
    # STEP 2: Crea tutte_righe_univoche per questa combinazione
    # -------------------------------------------------------------------------
    righe_univoche_path = OUTPUT_DIR / f"tutte_righe_univoche_{nome}.xlsx"

    print(f"\n[2/2] Generazione tutte_righe_univoche_{nome}.xlsx...")

    # Rileva modelli dalla mappatura per pasarli a crea_tutte_righe_univoche_v2
    bom_models_for_univoche = []
    try:
        _mapp_df = pd.read_excel(str(mappatura_path), sheet_name='Schema')
        _model_row = _mapp_df[_mapp_df['Codice caratteristica'].astype(str).str.upper() == 'ZPRODGRP1']
        if _model_row.empty:
            _model_row = _mapp_df.head(1)
        _opz_cols = [f'Opz{i}' for i in range(1, 20)]
        for _c in _opz_cols:
            if _c in _model_row.columns:
                for _v in _model_row[_c].dropna():
                    _s = str(_v).strip()
                    if _s and _s.upper() != 'FINE':  # esclude 'FINE' (marcatore di fine)
                        bom_models_for_univoche.append(_s)
        bom_models_for_univoche = list(dict.fromkeys(bom_models_for_univoche))  # dedup preserva ordine
        print(f"  Modelli rilevati per univoche: {bom_models_for_univoche}")
    except Exception as e:
        print(f"  [WARN] Rilevamento modelli per univoche fallito: {e}")

    df_righe = crea_tutte_righe_univoche_v2(
        file_path=str(matrix_output_path),
        output_path=str(righe_univoche_path),
        model_col_names=bom_models_for_univoche if bom_models_for_univoche else None
    )

    # -------------------------------------------------------------------------
    # POST-PROCESSING: gestione righe Model = "Nessuno"
    #   - BOM multi-modello (es. V21E con MSV4+MSV4S+MSV4PP):
    #       ogni riga Nessuno viene espansa in N righe, una per ogni modello
    #   - BOM mono-modello (es. V22E -> MSV4RI, V24T -> MSV4RS):
    #       "Nessuno" viene sostituito direttamente con l'unico modello BOM
    #
    # I modelli vengono rilevati dinamicamente dalla riga 3 del foglio Matrix,
    # senza hardcoding: funziona anche se in futuro cambiano nomi/numero modelli.
    # -------------------------------------------------------------------------
    if 'Model' in df_righe.columns:
        # Rileva modelli dalla matrice già generata
        bom_models = []
        try:
            df_raw = pd.read_excel(str(matrix_output_path), sheet_name='Matrix',
                                   header=None, nrows=3)
            name_row = df_raw.iloc[2].fillna('').astype(str).str.strip()
            # Legge i valori modello dalla mappatura (riga ZPRODGRP1 o prima riga Bundle=No)
            try:
                _mapp_df = pd.read_excel(str(mappatura_path), sheet_name='Schema')
                _model_row = _mapp_df[_mapp_df['Codice caratteristica'].astype(str).str.upper() == 'ZPRODGRP1']
                if _model_row.empty:
                    _model_row = _mapp_df.head(1)
                _opz_cols = [f'Opz{i}' for i in range(1, 20)]
                _model_vals = set()
                for _c in _opz_cols:
                    if _c in _model_row.columns:
                        for _v in _model_row[_c].dropna():
                            _s = str(_v).strip()
                            if _s:
                                _model_vals.add(_s.lower())
                bom_models = [v for v in name_row if v.lower() in _model_vals and v]
            except Exception:
                # fallback: qualsiasi token non vuoto e non numerico nella riga 3
                bom_models = [v for v in name_row if v and not v.replace('.','').isdigit()]
            print(f"  [Model] Modelli rilevati per {nome}: {bom_models}")
        except Exception as e:
            print(f"  [WARN] Rilevamento modelli BOM fallito: {e}")

        mask_nessuno = df_righe['Model'] == 'Nessuno'
        n_nessuno = int(mask_nessuno.sum())

        if n_nessuno > 0 and bom_models:
            if len(bom_models) > 1:
                # Multi-modello: split ogni riga Nessuno in len(bom_models) righe
                righe_nessuno = df_righe[mask_nessuno].copy()
                righe_altri   = df_righe[~mask_nessuno].copy()

                esplose = []
                for model in bom_models:
                    blocco = righe_nessuno.copy()
                    blocco['Model'] = model
                    esplose.append(blocco)

                df_esplose = pd.concat(esplose, ignore_index=True)
                df_righe   = pd.concat([righe_altri, df_esplose], ignore_index=True)

                # Rinumera ID dopo l'esplosione
                if 'ID' in df_righe.columns:
                    df_righe['ID'] = range(1, len(df_righe) + 1)

                print(f"  [Model] {n_nessuno} righe 'Nessuno' esplose in "
                      f"{len(df_esplose)} ({len(bom_models)} modelli: "
                      f"{', '.join(bom_models)})")
            else:
                # Mono-modello: sostituzione diretta
                df_righe.loc[mask_nessuno, 'Model'] = bom_models[0]
                print(f"  [Model] {n_nessuno} righe 'Nessuno' sostituiti con "
                      f"'{bom_models[0]}'")
        elif n_nessuno > 0:
            print(f"  [WARN] {n_nessuno} righe 'Nessuno' rimaste: "
                  f"nessun modello rilevato dalla matrice")

    # Risalva il file con la colonna Model corretta (sovrascrive la versione
    # prodotta da crea_tutte_righe_univoche_v2 che aveva ancora i Nessuno)
    with pd.ExcelWriter(str(righe_univoche_path), engine='openpyxl') as writer:
        df_righe.to_excel(writer, sheet_name='Tutte_Righe_Univoche', index=False)
    print(f"[OK] Salvato (con Model corretti): {righe_univoche_path}")

    # Aggiungi colonna di versione per tracciabilità nel DataFrame in memoria
    # (non viene risalvata nel file singolo; compare solo nel consolidato V2)
    df_righe.insert(1, 'Versione', nome)

    return df_righe


def consolidate_results(results: dict) -> pd.DataFrame:
    """
    Consolida tutti i risultati in un unico DataFrame con deduplicazione e tracciabilità.

    Args:
        results: dict {nome_versione: DataFrame}

    Returns:
        DataFrame consolidato con colonna 'Versione_BOM' tracciata
    """
    print(f"\n{'='*80}")
    print("CONSOLIDAMENTO RISULTATI CON DEDUPLICAZIONE")
    print(f"{'='*80}")

    dfs = []
    for nome, df in results.items():
        print(f"  {nome}: {len(df)} righe univoche")
        dfs.append(df)

    # Concatena mantenendo tutte le righe
    df_all = pd.concat(dfs, ignore_index=True, sort=False)

    # ─────────────────────────────────────────────────────────────────────────────
    # FILL NaN → 'No' sulle colonne caratteristiche (optional/pack/seat/etc.)
    # Serve a rendere invarianti le righe rispetto alle caratteristiche assenti
    # nella propria BOM: una riga V24T senza FAIRING COLOR diventa equivalente
    # a una riga che ha esplicitamente 'No' per quella caratteristica.
    # Questo garantisce che il matching con l'output della simulazione funzioni
    # correttamente anche tra versioni con colonne optional diverse.
    # ─────────────────────────────────────────────────────────────────────────────
    _structural = {'ID', 'ID_Globale', 'Versione', 'Versioni_BOM',
                   'Model', 'Numero componenti'}
    _char_cols = [c for c in df_all.columns if c not in _structural]
    df_all[_char_cols] = df_all[_char_cols].fillna('No')

    righe_pre_dedup = len(df_all)
    print(f"\n  Righe PRIMA deduplicazione: {righe_pre_dedup}")

    # ─────────────────────────────────────────────────────────────────────────────
    # DEDUPLICAZIONE INTRA-VERSIONE: rimuove solo duplicati all'interno della
    # stessa BOM (es. 13 righe V21E identiche tra loro).
    # NON deduplica mai tra versioni diverse: una riga V24T con Model=MSV4RS e
    # stessa configurazione di una riga V21E con Model=MSV4 è considerata
    # semanticamente distinta e viene mantenuta in entrambe.
    # ─────────────────────────────────────────────────────────────────────────────

    # Colonne di identità INCLUSA 'Versione' → dedup solo intra-BOM
    id_cols_intra = [col for col in df_all.columns if col not in ['ID', 'ID_Globale']]

    # Prima eseguiamo la dedup intra-versione
    df_all['Versioni_BOM'] = df_all['Versione']   # ogni riga appartiene a una sola versione
    df_consolidated = df_all.drop_duplicates(subset=id_cols_intra, keep='first').copy()

    righe_post_dedup = len(df_consolidated)
    duplicati_rimossi = righe_pre_dedup - righe_post_dedup

    print(f"  Righe DOPO deduplicazione:  {righe_post_dedup}")
    print(f"  Duplicati intra-BOM rimossi: {duplicati_rimossi} ({(duplicati_rimossi/righe_pre_dedup*100):.1f}%)")
    for ver in df_consolidated['Versione'].dropna().unique():
        n = (df_consolidated['Versione'] == ver).sum()
        print(f"    {ver}: {n} righe")

    # Renumera gli ID
    df_consolidated.insert(0, 'ID_Globale', range(1, len(df_consolidated) + 1))

    # Riordina colonne: ID_Globale, Versioni_BOM (tracciabilità), Versione, Model, resto
    cols_order = ['ID_Globale', 'Versioni_BOM', 'Versione', 'Model']
    cols_order = [c for c in cols_order if c in df_consolidated.columns]
    cols_remaining = [c for c in df_consolidated.columns if c not in cols_order]
    df_consolidated = df_consolidated[cols_order + cols_remaining]

    # Statistiche per tracciabilità
    print(f"\n  TOTALE CONSOLIDATO: {righe_post_dedup} righe univoche")

    # Mostra quante righe sono comuni/uniche per versione
    print(f"\n  Distribuzione Versioni BOM:")
    versioni_dist = df_consolidated['Versioni_BOM'].value_counts()
    for versioni, count in versioni_dist.items():
        print(f"    {versioni}: {count} righe")

    return df_consolidated


# ============================================================================
# CLEANUP: ELIMINA FILE STALI
# ============================================================================

def _cleanup_safety_stock_files(input_dir: str = ".\\Input") -> None:
    """
    Elimina tutti i file matrix_output_*.xlsx e tutte_righe_univoche_*.xlsx
    (esclude SA che sono trattati da crea_catalogo_sa.py).
    Eseguito all'inizio di main() per evitare file stali.
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


def main(filter_untraced: bool = True):
    """Esegui il processamento completo."""

    # =========================================================================
    # CLEANUP: Elimina file stali prima di rigenerare
    # =========================================================================
    _cleanup_safety_stock_files(input_dir=".\\Input")

    print("\n" + "="*80)
    print("PROCESSAMENTO MULTI-BOM: V21E + V22E + V24T")
    print("="*80)

    # =========================================================================
    # FASE -1: Costruisce la mappatura unificata a runtime dalle originali
    # =========================================================================
    build_unified_mapping(
        source_files=MAPPATURE_ORIGINALI,
        output_path=_MAPPATURA_UNIFICATA,
    )

    # =========================================================================
    # FASE 0: Quality checks globali
    # =========================================================================
    residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set = run_quality_checks()

    # =========================================================================
    # FASE 1: Processa tutte le combinazioni
    # =========================================================================
    results = {}
    for combo in COMBINAZIONI:
        try:
            df = process_combination(combo, tr_char_options=tr_char_options, filter_untraced=filter_untraced)
            results[combo["nome"]] = df
        except Exception as e:
            print(f"\n[ERRORE] nel processamento {combo['nome']}: {e}")
            raise

    # =========================================================================
    # FASE 2: Consolida in un unico file
    # =========================================================================
    df_final = consolidate_results(results)

    # =========================================================================
    # FASE 3: Salva il file finale consolidato con Quality Summary
    # =========================================================================
    output_path = OUTPUT_DIR / "tutte_righe_univoche_V2.xlsx"

    print(f"\n{'='*80}")
    print("SALVATAGGIO FILE FINALE + QUALITY SUMMARY")
    print(f"{'='*80}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # Foglio principale: tutte le righe consolidate
        df_final.to_excel(writer, sheet_name='Tutte_Righe_Univoche', index=False)

        # Foglio per ogni versione (per tracciabilità)
        for nome, df in results.items():
            sheet_name = f"Righe_{nome}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Foglio riepilogo con deduplicazione
        riepilogo_data = {
            'Versione': list(results.keys()),
            'Righe Univoche': [len(df) for df in results.values()],
        }
        df_riepilogo = pd.DataFrame(riepilogo_data)
        df_riepilogo.loc[len(df_riepilogo)] = {
            'Versione': 'TOTALE CONSOLIDATO (dopo dedup)',
            'Righe Univoche': len(df_final)
        }
        df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)

        # Foglio deduplicazione: tracciabilità delle versioni BOM per ogni riga
        dedup_summary = df_final[['ID_Globale', 'Versioni_BOM']].copy()
        dedup_summary['Num_Versioni'] = dedup_summary['Versioni_BOM'].str.count(r'\+') + 1
        dedup_summary = dedup_summary.sort_values('Num_Versioni', ascending=False)
        dedup_summary.to_excel(writer, sheet_name='Dedup_Tracciabilita', index=False)

        # Foglio quality summary (riepilogo issues)
        quality_counts = len(residual_issues) + len(price_issues) + len(ct_cv_issues)
        quality_summary = {
            'Categoria': ['Residual Issues', 'Price Issues', 'CT/CV Issues', 'TOTALE'],
            'Conteggio': [len(residual_issues), len(price_issues), len(ct_cv_issues), quality_counts]
        }
        df_quality = pd.DataFrame(quality_summary)
        df_quality.to_excel(writer, sheet_name='Quality_Summary', index=False)

        if quality_counts > 0:
            print(f"\n[INFO] Quality issues rilevati: {quality_counts}")
            print(f"       - Residual: {len(residual_issues)}")
            print(f"       - Price: {len(price_issues)}")
            print(f"       - CT/CV: {len(ct_cv_issues)}")
        else:
            print(f"\n[OK] Nessun quality issue rilevato")

    print(f"\n[OK] File consolidato salvato:")
    print(f"  {output_path}")
    print(f"  Dimensioni: {len(df_final)} righe × {len(df_final.columns)} colonne")

    # Analisi righe univoche vs condivise
    df_final_for_analysis = df_final.copy()
    righe_univoche = (df_final_for_analysis['Versioni_BOM'].str.count(r'\+') == 0).sum()
    righe_condivise = (df_final_for_analysis['Versioni_BOM'].str.count(r'\+') > 0).sum()

    print(f"\nAnalisi Deduplicazione:")
    print(f"  - Righe UNIVOCHE (single version): {righe_univoche}")
    print(f"  - Righe CONDIVISE (multi version): {righe_condivise}")

    print(f"\nComposizione per Versione BOM Originale:")
    for nome in results.keys():
        count = len(results[nome])
        pct = (count / sum(len(df) for df in results.values()) * 100)
        print(f"  - {nome}: {count} righe ({pct:.1f}%)")

    return df_final


if __name__ == "__main__":
    df_result = main()
    print("\n" + "="*80)
    print("[OK] PROCESSAMENTO COMPLETATO")
    print("="*80)

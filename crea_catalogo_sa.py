# -*- coding: utf-8 -*-
"""
================================================================================
PREPARAZIONE CATALOGO STOCK ALIGNMENT — tutti i componenti BOM
================================================================================

Genera tutte_righe_univoche_SA.xlsx includendo TUTTI i componenti BOM,
senza il filtro "Check Generale = Mostra" usato da Safety Stock.

DIFFERENZE rispetto a process_all_bom_combinations.py:
  - L'enrichment BOM NON filtra su Check Generale = Mostra
  - Tutti i componenti (Mostra E Nascondi) vengono processati
  - I componenti con dipendenza → righe Bundle/CT/CV correttamente parsate
  - I componenti senza dipendenza → matchano tutte le configurazioni (OK)
  - Aggiunge Livello_Min / Livello_Max per filtraggio interattivo in dashboard

Script Safety Stock INTOCCATI: questo script non modifica nessun file esistente.

================================================================================
"""

import re
import time
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict

# ============================================================================
# IMPORT dai moduli esistenti (sola lettura — nessuna modifica)
# ============================================================================
from gestione_dipendenze_V2 import (
    process_excel_to_matrix_v2,
    _find_raw_file,
    _bom_to_str,
    _DEP_COL_ENRICH,
    _build_active_cv_sets,
    _load_states_set,
    load_mapping_from_file,
)
from skubundle import crea_tutte_righe_univoche_v2


# ============================================================================
# CONFIGURAZIONE
# ============================================================================

INPUT_DIR = Path(".") / "Input"
MODEL_DIR = INPUT_DIR / "MODEL"

STATES_FILE   = INPUT_DIR / "States.xlsx"
PREZZI_FILE   = INPUT_DIR / "Prezzi.xlsx"
TR_FILE       = INPUT_DIR / "TR TOTALV21E - Copia.xlsx"

# Dinamicamente trova modelli disponibili in INPUT_DIR/MODEL/
def _get_available_models():
    """Trova modelli disponibili in INPUT_DIR/MODEL/"""
    models = []
    if MODEL_DIR.exists():
        for model_folder in sorted(MODEL_DIR.iterdir()):
            if model_folder.is_dir():
                model_name = model_folder.name
                mappatura = model_folder / f"Mappatura_{model_name}.xlsx"
                if mappatura.exists():
                    models.append({
                        "nome": model_name,
                        "in_dir": model_folder,
                        "mappatura": mappatura,
                        "tr": TR_FILE,
                    })
    return models

COMBINAZIONI = _get_available_models()
if not COMBINAZIONI:
    # Fallback: try hardcoded paths
    COMBINAZIONI = [
        {
            "nome": "V21E",
            "in_dir": MODEL_DIR / "V21E",
            "mappatura": MODEL_DIR / "V21E" / "Mappatura_V21E.xlsx",
            "tr": TR_FILE,
        },
    ]


# ============================================================================
# ENRICHMENT BOM SENZA FILTRO CHECK GENERALE
# Replica di enrich_bom_version da gestione_dipendenze_V2.py
# con l'unica differenza che NON applica il filtro "Check Generale = Mostra"
# ============================================================================

def _enrich_bom_full(ver: str, in_dir) -> pd.DataFrame:
    """
    Arricchisce il BOM grezzo per la versione VER includendo TUTTI i componenti
    (senza il filtro Check Generale = Mostra che usa Safety Stock).

    Logica identica a enrich_bom_version in gestione_dipendenze_V2.py,
    tranne che la riga:
        bom = bom[bom["Check Generale"] == "Mostra"].copy()
    viene OMESSA.

    Aggiunge la colonna "Check_Generale_SA" per mantenere tracciabilità
    (Mostra / Nascondi) senza alterare il flusso di process_excel_to_matrix_v2.
    """
    in_dir = Path(in_dir)
    print(f"\n  [ENRICH-SA] Versione: {ver} (TUTTI i livelli, no filtro Mostra)")

    bom_file  = _find_raw_file(in_dir, f"BOM150_D{ver}*.XLSX")
    dag_file  = _find_raw_file(in_dir, f"DATI_AGGIUNTI*_D{ver}*.XLSX")
    b125_file = _find_raw_file(in_dir, f"ZIBP_BOM_125*_D{ver}*.XLSX")
    print(f"  [ENRICH-SA]   BOM150 : {bom_file.name}")
    print(f"  [ENRICH-SA]   DAG    : {dag_file.name}")
    print(f"  [ENRICH-SA]   BOM125 : {b125_file.name}")

    bom = pd.read_excel(bom_file)
    dag = pd.read_excel(dag_file)

    # Join Dati Aggiuntivi
    dag_sel = (dag[["Materiale", "Car. MRP", "Tipo approvv.",
                    "Approvv. speciale", "Merce sfusa"]]
               .rename(columns={"Merce sfusa": "Merce sfusa (DA)"}))
    bom = (bom.merge(dag_sel, left_on="Numero componenti",
                     right_on="Materiale", how="left")
              .drop(columns=["Materiale"]))

    sfusa_bom = [_bom_to_str(v) for v in bom["Merce sfusa"]]
    bom = bom.drop(columns=["Merce sfusa"])
    bom.insert(0, "ID", range(1, len(bom) + 1))
    bom["Livello numerico"] = bom["Liv. esplosione"].apply(
        lambda v: str(v).count(".") if pd.notna(v) else 0
    )

    # Formule Check Generale (calcolate ma NON usate per filtrare)
    n      = len(bom)
    H      = [0] * n;  I_chk = [""] * n
    J      = [0] * n;  K_chk = [""] * n;  L_chk = [""] * n
    lev    = bom["Livello numerico"].tolist()
    tipo   = [_bom_to_str(v) for v in bom["Tipo approvv."]]
    approv = [_bom_to_str(v) for v in bom["Approvv. speciale"]]
    numpos = ([_bom_to_str(v) for v in bom["Numero posizione"]]
              if "Numero posizione" in bom.columns else [""] * n)

    for i in range(n):
        G_i = lev[i];  C_i = tipo[i];  D_i = approv[i]
        O_i = sfusa_bom[i];  R_i = numpos[i]
        Hp  = H[i - 1] if i > 0 else 0
        Jp  = J[i - 1] if i > 0 else 0

        if C_i == "F" and D_i != "30" and O_i != "X" and (Hp >= G_i or Hp == 0):
            H[i] = G_i
        elif G_i == Hp and C_i == "F" and D_i != "30" and R_i != "X":
            H[i] = 0
        elif G_i <= Hp:
            H[i] = 0
        else:
            H[i] = Hp

        if C_i == "" and G_i == 2 and O_i == "":
            I_chk[i] = "Mostra"
        elif C_i in ("E", "") and D_i != "30" and O_i != "X":
            I_chk[i] = "Nascondi"
        elif C_i == "F" and D_i != "30" and O_i != "X" and (H[i] == 0 or H[i] == G_i):
            I_chk[i] = "Mostra"
        elif Hp > 0 and G_i > Hp:
            I_chk[i] = "Nascondi"
        else:
            I_chk[i] = "Mostra"

        J[i] = G_i if O_i == "X" else (Jp if Jp > 0 and G_i > Jp else 0)
        K_chk[i] = "Nascondi" if J[i] > 0 else "Mostra"
        if G_i == 1:
            L_chk[i] = "Mostra"
        elif K_chk[i] == "Nascondi" or I_chk[i] == "Nascondi":
            L_chk[i] = "Nascondi"
        else:
            L_chk[i] = "Mostra"

    bom["Livello Segnalibro Merce Non Sfusa"] = H
    bom["Check Merce non sfusa"]              = I_chk
    bom["Livello Segnalibro Merce Sfusa"]     = J
    bom["Check Merce sfusa"]                  = K_chk
    bom["Check Generale"]                     = L_chk
    # Colonna tracciabilità SA (non interferisce con process_excel_to_matrix_v2)
    bom["Check_Generale_SA"]                  = L_chk

    n_mostra   = sum(1 for v in L_chk if v == "Mostra")
    n_nascondi = sum(1 for v in L_chk if v == "Nascondi")
    print(f"  [ENRICH-SA]   Mostra:   {n_mostra}")
    print(f"  [ENRICH-SA]   Nascondi: {n_nascondi}  (inclusi in SA)")

    # BOM125: aggrega colonna dipendenze per componente
    bom125 = pd.read_excel(b125_file, header=0)
    bom125_dep: Dict[str, str] = {}
    for _, row in bom125.iterrows():
        comp    = _bom_to_str(row["Componente critico"])
        dep_raw = _bom_to_str(row.get(_DEP_COL_ENRICH, ""))
        dep_val = re.sub(r'\s*\|\s*', ' OR ', dep_raw).strip() if dep_raw else ""
        if not comp:
            continue
        if comp not in bom125_dep:
            bom125_dep[comp] = dep_val
        elif dep_val and not bom125_dep[comp]:
            bom125_dep[comp] = dep_val
        elif dep_val:
            bom125_dep[comp] += " OR " + dep_val

    bom125_df = pd.DataFrame(
        list(bom125_dep.items()),
        columns=["Componente critico", _DEP_COL_ENRICH]
    )

    # Join dipendenze su tutte le righe (PRIMA del filtro — che qui non applichiamo)
    bom = bom.merge(bom125_df, left_on="Numero componenti",
                    right_on="Componente critico", how="left")
    bom = bom.drop(columns=["Componente critico"])
    dep_vals = bom.pop(_DEP_COL_ENRICH)
    bom.insert(3, _DEP_COL_ENRICH, dep_vals)
    bom[_DEP_COL_ENRICH] = bom[_DEP_COL_ENRICH].fillna("")

    # Propagazione dipendenze
    dep_arr    = bom[_DEP_COL_ENRICH].tolist()
    levels     = bom["Livello numerico"].tolist()
    anchor_dep = "";  anchor_set = False;  propagated = 0
    for i in range(len(bom)):
        lvl = levels[i];  row_dep = dep_arr[i]
        if lvl == 2:
            anchor_dep = row_dep if row_dep else ""; anchor_set = True
        elif lvl == 1:
            pass
        elif anchor_set and not row_dep and anchor_dep:
            dep_arr[i] = anchor_dep; propagated += 1
    bom[_DEP_COL_ENRICH] = dep_arr
    print(f"  [ENRICH-SA]   Dipendenze propagate: {propagated}")

    # *** NESSUN FILTRO Check Generale ***
    # Safety Stock fa: bom = bom[bom["Check Generale"] == "Mostra"].copy()
    # Noi NON lo facciamo.

    bom["ID"] = range(1, len(bom) + 1)
    print(f"  [ENRICH-SA]   Righe totali (tutti livelli): {len(bom)}")

    return bom


# ============================================================================
# PROCESS COMBINATION SA
# ============================================================================

def _process_combination_sa(combo: dict, filter_untraced: bool = True) -> pd.DataFrame:
    """
    Processa una combinazione BOM+Mappatura per Stock Alignment:
      1. Enrichment SENZA filtro Mostra (_enrich_bom_full)
      2. process_excel_to_matrix_v2 con df_input
      3. crea_tutte_righe_univoche_v2
      4. Gestione righe Model=Nessuno (identica a process_all_bom_combinations)
    """
    nome           = combo["nome"]
    in_dir         = combo["in_dir"]
    mappatura_path = combo["mappatura"]
    tr_path        = combo["tr"]

    print(f"\n{'='*70}")
    print(f"STOCK ALIGNMENT — COMBINAZIONE: {nome}")
    print(f"{'='*70}")

    # Step 0: Enrichment SENZA filtro
    df_enriched = _enrich_bom_full(nome, in_dir)

    # Step 1: matrix output
    matrix_output_path = INPUT_DIR / f"matrix_output_SA_{nome}.xlsx"
    print(f"\n[1/2] Generazione matrix_output_SA_{nome}.xlsx...")
    process_excel_to_matrix_v2(
        df_input=df_enriched,
        mapping_file=str(mappatura_path),
        tr_file=str(tr_path),
        output_file=str(matrix_output_path),
        states_file=str(STATES_FILE) if STATES_FILE.exists() else None,
        prezzi_file=str(PREZZI_FILE) if PREZZI_FILE.exists() else None,
        filter_untraced=filter_untraced,
    )

    if not matrix_output_path.exists():
        raise FileNotFoundError(f"Matrix output non creato: {matrix_output_path}")
    print(f"  [OK] {matrix_output_path.name} ({matrix_output_path.stat().st_size:,} bytes)")
    time.sleep(1)

    # Step 2: righe univoche
    righe_path = INPUT_DIR / f"tutte_righe_univoche_SA_{nome}.xlsx"
    print(f"\n[2/2] Generazione tutte_righe_univoche_SA_{nome}.xlsx...")
    df_righe = crea_tutte_righe_univoche_v2(
        file_path=str(matrix_output_path),
        output_path=str(righe_path),
    )

    # Gestione Model=Nessuno (identica a process_all_bom_combinations)
    if "Model" in df_righe.columns:
        bom_models = []
        try:
            df_raw = pd.read_excel(str(matrix_output_path), sheet_name="Matrix",
                                   header=None, nrows=3)
            name_row = df_raw.iloc[2].fillna("").astype(str).str.strip()
            bom_models = [v for v in name_row if v.lower().startswith("msv4")]
            print(f"  [Model] Modelli rilevati per {nome}: {bom_models}")
        except Exception as e:
            print(f"  [WARN] Rilevamento modelli fallito: {e}")

        mask_nessuno = df_righe["Model"] == "Nessuno"
        n_nessuno = int(mask_nessuno.sum())

        if n_nessuno > 0 and bom_models:
            if len(bom_models) > 1:
                righe_altri = df_righe[~mask_nessuno].copy()
                esplose = []
                for model in bom_models:
                    blocco = df_righe[mask_nessuno].copy()
                    blocco["Model"] = model
                    esplose.append(blocco)
                df_righe = pd.concat([righe_altri, pd.concat(esplose, ignore_index=True)],
                                     ignore_index=True)
                if "ID" in df_righe.columns:
                    df_righe["ID"] = range(1, len(df_righe) + 1)
                print(f"  [Model] {n_nessuno} righe Nessuno esplose in "
                      f"{len(bom_models)} modelli")
            else:
                df_righe.loc[mask_nessuno, "Model"] = bom_models[0]
                print(f"  [Model] {n_nessuno} righe Nessuno -> '{bom_models[0]}'")

    # Risalva con Model corretti
    with pd.ExcelWriter(str(righe_path), engine="openpyxl") as writer:
        df_righe.to_excel(writer, sheet_name="Tutte_Righe_Univoche", index=False)
    print(f"  [OK] Salvato: {righe_path.name}")

    df_righe.insert(1, "Versione", nome)
    return df_righe


# ============================================================================
# CONSOLIDAMENTO
# ============================================================================

def _consolidate_sa(results: dict) -> pd.DataFrame:
    """
    Consolida i risultati SA con dedup intra-versione e tracciabilità Versioni_BOM.
    Logica identica a consolidate_results in process_all_bom_combinations.py.
    """
    print(f"\n{'='*70}")
    print("CONSOLIDAMENTO SA — dedup intra-versione")
    print(f"{'='*70}")

    dfs = [df for df in results.values()]
    df_all = pd.concat(dfs, ignore_index=True, sort=False)

    _structural = {"ID", "ID_Globale", "Versione", "Versioni_BOM",
                   "Model", "Numero componenti", "Check_Generale_SA"}
    _char_cols = [c for c in df_all.columns if c not in _structural]
    df_all[_char_cols] = df_all[_char_cols].fillna("No")

    print(f"  Righe prima dedup: {len(df_all)}")
    df_all["Versioni_BOM"] = df_all["Versione"]
    id_cols_intra = [col for col in df_all.columns if col not in ["ID", "ID_Globale"]]
    df_cons = df_all.drop_duplicates(subset=id_cols_intra, keep="first").copy()
    print(f"  Righe dopo dedup:  {len(df_cons)}")

    df_cons.insert(0, "ID_Globale", range(1, len(df_cons) + 1))

    cols_order = ["ID_Globale", "Versioni_BOM", "Versione", "Model"]
    cols_order = [c for c in cols_order if c in df_cons.columns]
    cols_rest  = [c for c in df_cons.columns if c not in cols_order]
    df_cons = df_cons[cols_order + cols_rest]

    return df_cons


# ============================================================================
# AGGIUNTA LIVELLI BOM
# ============================================================================

def _add_bom_levels(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggiunge Livello_Min / Livello_Max leggendo i file BOM150 di tutte le versioni.
    Un componente può apparire a livelli diversi nelle varie versioni BOM.
    """
    print("\n  Calcolo Livello_Min / Livello_Max da BOM150...")
    records = []
    # Estrai versioni da COMBINAZIONI
    for combo in COMBINAZIONI:
        ver = combo['nome']
        ver_dir = combo['in_dir']
        if not ver_dir.exists():
            print(f"    {ver}: cartella non trovata, skip")
            continue
        try:
            bom_file = _find_raw_file(ver_dir, f"BOM150_D{ver}*.XLSX")
            bom = pd.read_excel(bom_file, usecols=["Numero componenti", "Liv. esplosione"])
            bom["_lev"] = bom["Liv. esplosione"].apply(
                lambda v: str(v).count(".") if pd.notna(v) else 0
            )
            records.append(bom[["Numero componenti", "_lev"]].copy())
        except Exception as e:
            print(f"    [WARN] {ver}: {e}")

    if not records:
        print("    [WARN] Nessun livello BOM disponibile")
        df["Livello_Min"] = None
        df["Livello_Max"] = None
        return df

    all_lev = pd.concat(records, ignore_index=True)
    all_lev["Numero componenti"] = all_lev["Numero componenti"].astype(str).str.strip()
    all_lev["_lev"] = pd.to_numeric(all_lev["_lev"], errors="coerce").fillna(0).astype(int)

    level_map = (
        all_lev.groupby("Numero componenti")["_lev"]
        .agg(Livello_Min="min", Livello_Max="max")
        .reset_index()
    )

    df["Numero componenti"] = df["Numero componenti"].astype(str).str.strip()
    for col in ["Livello_Min", "Livello_Max"]:
        if col in df.columns:
            df = df.drop(columns=[col])

    df = df.merge(level_map, on="Numero componenti", how="left")
    covered = df["Livello_Min"].notna().sum()
    print(f"    Righe con livello: {covered} / {len(df)}")

    # Distribuzione
    dist = (
        df.drop_duplicates("Numero componenti")["Livello_Min"]
        .dropna().astype(int).value_counts().sort_index()
    )
    for lev, cnt in dist.items():
        print(f"    Livello {lev}: {cnt} SKU")

    return df


# ============================================================================
# MAIN
# ============================================================================

def main(
    output_catalog: str = ".\\Input\\tutte_righe_univoche_SA.xlsx",
    filter_untraced: bool = True,
) -> pd.DataFrame:
    """
    Genera il catalogo SKU per Stock Alignment includendo TUTTI i componenti BOM.

    Args:
        output_catalog:  Path file output
        filter_untraced: Applica logica AND/OR/IN per CT/CV sconosciuti (default True)

    Returns:
        DataFrame del catalogo SA.
    """
    sep = "=" * 70
    print(f"\n{sep}")
    print("PREPARAZIONE CATALOGO STOCK ALIGNMENT")
    print(f"  Tutti i componenti BOM — senza filtro Mostra/Nascondi")
    print(f"  filter_untraced={filter_untraced}")
    print(f"{sep}")

    # Cleanup file precedenti
    print("\n[CLEANUP] Rimozione file SA precedenti...")
    input_path = Path(".") / "Input"
    removed = 0
    for f in input_path.glob("tutte_righe_univoche_SA*.xlsx"):
        try:
            f.unlink()
            print(f"  Rimosso: {f.name}")
            removed += 1
        except Exception as e:
            print(f"  [WARN] Non posso rimuovere {f.name}: {e}")
    if removed == 0:
        print("  (nessun file precedente trovato)")

    # Verifica prerequisiti
    for combo in COMBINAZIONI:
        if not combo["in_dir"].exists():
            raise FileNotFoundError(f"Cartella non trovata: {combo['in_dir']}")
        if not combo["mappatura"].exists():
            raise FileNotFoundError(f"Mappatura non trovata: {combo['mappatura']}")
        if not combo["tr"].exists():
            raise FileNotFoundError(f"File TR non trovato: {combo['tr']}")

    # Processa ogni combinazione
    results = {}
    for combo in COMBINAZIONI:
        try:
            df = _process_combination_sa(combo, filter_untraced=filter_untraced)
            results[combo["nome"]] = df
        except Exception as e:
            print(f"\n[ERRORE] Combinazione {combo['nome']}: {e}")
            raise

    # Consolida
    df_sa = _consolidate_sa(results)

    # Aggiungi livelli BOM
    df_sa = _add_bom_levels(df_sa)

    # Salva
    print(f"\n  Salvataggio: {output_catalog}")
    df_sa.to_excel(output_catalog, sheet_name="Tutte_Righe_Univoche", index=False)

    n_sku = df_sa["Numero componenti"].nunique()
    print(f"\n{sep}")
    print("CATALOGO SA PRONTO")
    print(f"  Righe totali:  {len(df_sa)}")
    print(f"  SKU univoci:   {n_sku}")
    print(f"  Output:        {output_catalog}")
    print(f"{sep}")

    return df_sa


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()

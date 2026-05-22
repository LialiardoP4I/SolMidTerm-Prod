"""
optional_cost_allocation.py
===========================
Ripartisce il costo Safety Stock di ogni componente tra gli optional/configurazioni
che ne determinano la domanda.

Algoritmo
---------
1. Per ogni riga del catalogo SKU (tutte_righe_univoche_V2.xlsx):
     - Calcola il peso TR = TR_model * prod(TR_attr) per ogni attributo
       non-wildcard, usando i valori di parsed_data_complete.json
2. Per ogni componente (dal file SS output):
     - Peso_riga_r = TR_r / somma(TR_r')    =>  somma = 1
     - Per ogni riga r, identifica le colonne attive (valore != "no"/wildcard)
     - Suddividi il costo SS della riga tra le colonne attive in parti uguali
3. Aggrega per colonna optional: somma(costo_allocato)

Verifica: per ogni componente  somma(allocato) == prezzo_safety  (entro 1 cent)

Output (Excel multi-sheet):
  - Dettaglio       : una riga per (SKU, optional, valore)
  - Pivot_Pack      : righe = optional pack, colonne = top-N SKU
  - Pivot_Config    : righe = attributi configurazione, colonne = top-N SKU
  - Riepilogo_Pack  : totale SS per optional pack (ordinato)
  - Riepilogo_Config: totale SS per attributo config (ordinato)
"""

import json
import warnings
import numpy as np
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore")

# ── Percorsi ──────────────────────────────────────────────────────────────────
BASE   = Path(r"C:\Users\AntonioLiardo-SMOPS\OneDrive - Digital360\Desktop\Ducati\SolMidTerm")
INPUT  = BASE / "Input"
OUTPUT = BASE / "Output"

SS_FILE      = OUTPUT / "sku_safety_stock_leadtime_PER_SKU_OPTIMIZED.xlsx"
SKU_CATALOG  = INPUT  / "tutte_righe_univoche_V2.xlsx"
SKU_CATALOG_FALLBACK = INPUT / "_sku_tmp.xlsx"
PARSED_JSON  = OUTPUT / "parsed_data_complete.json"
OUT_FILE     = OUTPUT / "optional_cost_allocation.xlsx"

# Colonne che sono "optional pack" (scelta del cliente)
PACK_COLS = [
    "Radar Pack", "Touring pack", "Sport Optional Pack", "Adventure Pack",
    "Carbon fiber rear mud guard", "Auxiliary LED Lights", "Tech Pack",
    "Top Case Pack", "Alu top case pack", "Rear luggage rack",
    "Smoked windshield pack", "Smoked Windshield Pack Low",
    "Enduro pack - Silver", "Enduro pack - Black",
]

# Colonne che sono attributi di configurazione (non pack discrezionali)
CONFIG_COLS = [
    "FAIRING COLOR", "FRONT BRAKE", "PILLION SEAT",
    "RIDER SEAT", "RIMS", "SUBFRAME", "SUSPENSION",
]

# Valori considerati wildcard (componente richiesto indipendentemente dal valore)
WILDCARD_VALS = {"no", "", "nan", "nessuno", "none"}


# ── Funzioni di supporto ──────────────────────────────────────────────────────

def is_wildcard(val) -> bool:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return True
    return str(val).strip().lower() in WILDCARD_VALS


def build_tr_lookup(parsed: dict):
    """
    Ritorna:
      model_tr   : {model_name: float}   es. {"MSV4S": 0.91, ...}
      char_lookup: {char_name: {model: {value: float}}}
    """
    mm = parsed["model_mix"]
    model_tr = dict(zip(mm["models"], mm["take_rates"]))

    char_lookup = {}
    for char_name, char_data in parsed["characteristics"].items():
        values   = char_data["values"]
        char_lookup[char_name] = {}
        for model, tr_arr in char_data["tr"].items():
            char_lookup[char_name][model] = dict(zip(values, tr_arr))

    return model_tr, char_lookup


def lookup_tr(col: str, val: str, model: str, char_lookup: dict):
    """
    Cerca il TR per (col, val, model). Ritorna None se non trovato.
    Gestisce piccole discrepanze di nome con matching parziale.
    """
    if col not in char_lookup:
        return None
    by_model = char_lookup[col].get(model, {})
    if not by_model:
        return None

    # Match esatto
    if val in by_model:
        return by_model[val]

    # Match case-insensitive
    val_lo = val.lower()
    for k, v in by_model.items():
        if k.lower() == val_lo:
            return v

    # Match parziale (substring)
    for k, v in by_model.items():
        if val_lo in k.lower() or k.lower() in val_lo:
            return v

    return None


def compute_row_tr(row: pd.Series,
                   model_tr: dict,
                   char_lookup: dict,
                   attr_cols: list) -> float:
    """
    TR combinato = TR_model * prod(TR_attr per ogni colonna attiva).
    """
    model = str(row.get("Model", "")).strip()
    m_tr  = model_tr.get(model)
    if m_tr is None:
        # Fallback: matching parziale
        for k, v in model_tr.items():
            if k.lower() == model.lower():
                m_tr = v
                break
    if m_tr is None or m_tr == 0:
        return 0.0

    combined = m_tr
    for col in attr_cols:
        val = row.get(col)
        if is_wildcard(val):
            continue  # non restringe: non riduce la probabilità
        val = str(val).strip()
        tr = lookup_tr(col, val, model, char_lookup)
        if tr is not None:
            combined *= tr
        # Se non trovato non moltiplico (evito azzerare per bug di naming)

    return combined


# ── Core allocation ───────────────────────────────────────────────────────────

def allocate(ss_df: pd.DataFrame,
             sku_df: pd.DataFrame,
             model_tr: dict,
             char_lookup: dict) -> pd.DataFrame:
    """
    Ritorna DataFrame con colonne:
      SKU, Descrizione, SS_Cost_Totale, SS_Unita, Domanda_Media,
      Tipo_Optional, Colonna_Optional, Valore_Optional,
      Costo_SS_Allocato, Percentuale
    """
    _EXCLUDE_EXACT = {"ID", "Numero componenti", "Model"}
    _EXCLUDE_LOWER = {
        "versione", "versione_bom", "versioni_bom",
        "id_globale", "id globale",
    }
    ATTR_COLS = [
        c for c in sku_df.columns
        if c not in _EXCLUDE_EXACT
        and str(c).lower().strip() not in _EXCLUDE_LOWER
        and not str(c).lower().startswith("unnamed")
    ]

    # Pre-calcola TR per tutte le righe del catalogo
    sku_df = sku_df.copy()
    sku_df["_tr"] = sku_df.apply(
        lambda r: compute_row_tr(r, model_tr, char_lookup, ATTR_COLS), axis=1
    )

    records = []

    for _, ss_row in ss_df.iterrows():
        sku      = ss_row["SKU"]
        ss_cost  = ss_row.get("prezzo_safety", np.nan)
        ss_units = ss_row.get("safety_stock_p99_total", np.nan)
        mean_d   = ss_row.get("mean_demand", np.nan)
        desc     = ss_row.get("Description", "")

        if pd.isna(ss_cost) or ss_cost <= 0:
            continue

        comp = sku_df[sku_df["Numero componenti"] == sku].copy()
        if comp.empty:
            # Componente non nel catalogo SKU: attribuito a Base
            records.append({
                "SKU":               sku,
                "Descrizione":       desc,
                "SS_Cost_Totale":    ss_cost,
                "SS_Unita":          ss_units,
                "Domanda_Media":     mean_d,
                "Tipo_Optional":     "Base",
                "Colonna_Optional":  "Base (non in catalogo)",
                "Valore_Optional":   "—",
                "Costo_SS_Allocato": ss_cost,
                "Percentuale":       100.0,
            })
            continue

        total_tr = comp["_tr"].sum()
        if total_tr <= 0:
            # Fallback: pesi uguali se TR non calcolabile
            comp["_w"] = 1.0 / len(comp)
        else:
            comp["_w"] = comp["_tr"] / total_tr

        # Alloca riga per riga
        row_alloc: dict[tuple, float] = {}   # (tipo, col, val) -> costo

        for _, cat in comp.iterrows():
            row_cost = ss_cost * cat["_w"]
            model    = str(cat.get("Model", "")).strip()

            # Colonne attive per questa riga
            active: list[tuple[str, str, str]] = []  # (tipo, col, val)

            for col in ATTR_COLS:
                val = cat.get(col)
                if is_wildcard(val):
                    continue
                val_str = str(val).strip()
                tipo = "Pack" if col in PACK_COLS else "Configurazione"
                active.append((tipo, col, val_str))

            if not active:
                # Componente base (non restritto da nessun optional)
                key = ("Base", "Base", model if model else "—")
                row_alloc[key] = row_alloc.get(key, 0.0) + row_cost
            else:
                share = row_cost / len(active)
                for tipo, col, val_str in active:
                    key = (tipo, col, val_str)
                    row_alloc[key] = row_alloc.get(key, 0.0) + share

        # Scrivi record
        for (tipo, col, val_str), cost in row_alloc.items():
            records.append({
                "SKU":               sku,
                "Descrizione":       desc,
                "SS_Cost_Totale":    ss_cost,
                "SS_Unita":          ss_units,
                "Domanda_Media":     mean_d,
                "Tipo_Optional":     tipo,
                "Colonna_Optional":  col,
                "Valore_Optional":   val_str,
                "Costo_SS_Allocato": cost,
                "Percentuale":       cost / ss_cost * 100,
            })

    return pd.DataFrame(records)


# ── Verifica ──────────────────────────────────────────────────────────────────

def verify(detail_df: pd.DataFrame, ss_df: pd.DataFrame, tol: float = 0.02):
    """
    Controlla che per ogni SKU la somma allocata = prezzo_safety.
    Ritorna DataFrame con gli SKU che non tornano.
    """
    agg = (detail_df.groupby("SKU")["Costo_SS_Allocato"]
           .sum().reset_index()
           .rename(columns={"Costo_SS_Allocato": "Allocato_Totale"}))

    merged = ss_df[["SKU", "prezzo_safety"]].merge(agg, on="SKU", how="left")
    merged["Differenza"] = (merged["Allocato_Totale"] - merged["prezzo_safety"]).abs()
    merged["OK"] = merged["Differenza"] < tol

    bad = merged[~merged["OK"] & merged["prezzo_safety"].notna()
                 & (merged["prezzo_safety"] > 0)]
    return bad, merged


# ── Pivot helpers ─────────────────────────────────────────────────────────────

def make_pivot(detail_df: pd.DataFrame, tipo: str, top_n: int = 20) -> pd.DataFrame:
    """
    Pivot: righe = Colonna_Optional, colonne = SKU (top_n per costo).
    """
    sub = detail_df[detail_df["Tipo_Optional"] == tipo].copy()
    if sub.empty:
        return pd.DataFrame()

    # Top-N SKU per costo totale allocato
    top_skus = (sub.groupby("SKU")["Costo_SS_Allocato"]
                .sum().nlargest(top_n).index.tolist())
    sub_top = sub[sub["SKU"].isin(top_skus)]

    pivot = (sub_top.groupby(["Colonna_Optional", "SKU"])["Costo_SS_Allocato"]
             .sum().unstack(fill_value=0).round(2))
    pivot["TOTALE"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("TOTALE", ascending=False)
    return pivot


def make_summary(detail_df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    """
    Riepilogo: Colonna_Optional → totale SS cost allocato.
    """
    sub = detail_df[detail_df["Tipo_Optional"] == tipo]
    if sub.empty:
        return pd.DataFrame()

    summary = (sub.groupby("Colonna_Optional")
               .agg(
                   Costo_SS_Allocato=("Costo_SS_Allocato", "sum"),
                   N_SKU=("SKU", "nunique"),
               )
               .sort_values("Costo_SS_Allocato", ascending=False)
               .reset_index())
    summary["Percentuale_sul_Totale"] = (
        summary["Costo_SS_Allocato"] / summary["Costo_SS_Allocato"].sum() * 100
    ).round(1)
    return summary


# ── Main ──────────────────────────────────────────────────────────────────────

def run(ss_path=SS_FILE, sku_path=SKU_CATALOG, json_path=PARSED_JSON, out_path=OUT_FILE):
    print("=" * 60)
    print("Optional Cost Allocation — Ducati Safety Stock")
    print("=" * 60)

    # ── Carica dati ──────────────────────────────────────────────
    print(f"\n[1/5] Caricamento parsed_data_complete.json...")
    with open(json_path, encoding="utf-8") as f:
        parsed = json.load(f)
    model_tr, char_lookup = build_tr_lookup(parsed)
    print(f"      Modelli: {list(model_tr.keys())}")
    print(f"      Caratteristiche: {len(char_lookup)}")

    print(f"\n[2/5] Caricamento SS output: {ss_path.name}...")
    ss_df = pd.read_excel(ss_path)
    ss_df = ss_df[ss_df["prezzo_safety"].notna() & (ss_df["prezzo_safety"] > 0)].copy()
    print(f"      {len(ss_df)} componenti con costo SS > 0")

    print(f"\n[3/5] Caricamento SKU catalog: {sku_path.name}...")
    try:
        sku_df = pd.read_excel(sku_path)
    except PermissionError:
        fb = sku_path.parent / "_sku_tmp.xlsx"
        print(f"      File bloccato, uso fallback: {fb.name}")
        sku_df = pd.read_excel(fb)
    print(f"      {len(sku_df)} righe, {sku_df['Numero componenti'].nunique()} componenti unici")

    # ── Allocation ────────────────────────────────────────────────
    print(f"\n[4/5] Calcolo allocazione...")
    detail_df = allocate(ss_df, sku_df, model_tr, char_lookup)
    print(f"      {len(detail_df)} righe di allocazione generate")
    print(f"      Copertura: {detail_df['SKU'].nunique()} / {len(ss_df)} componenti allocati")

    # ── Verifica ──────────────────────────────────────────────────
    bad, check_df = verify(detail_df, ss_df)
    if bad.empty:
        print(f"      Verifica OK: tutti i totali tornano (tol=0.02 EUR)")
    else:
        print(f"      ATTENZIONE: {len(bad)} SKU con discrepanza > 0.02 EUR:")
        for _, r in bad.head(5).iterrows():
            print(f"        {r['SKU']}: atteso={r['prezzo_safety']:.2f}, "
                  f"allocato={r['Allocato_Totale']:.2f}, "
                  f"diff={r['Differenza']:.4f}")

    # ── Output Excel ──────────────────────────────────────────────
    print(f"\n[5/5] Scrittura output: {out_path.name}...")

    summ_pack   = make_summary(detail_df, "Pack")
    summ_config = make_summary(detail_df, "Configurazione")
    summ_base   = make_summary(detail_df, "Base")
    pivot_pack  = make_pivot(detail_df, "Pack",   top_n=25)
    pivot_cfg   = make_pivot(detail_df, "Configurazione", top_n=15)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Sheet 1: Riepilogo Pack
        if not summ_pack.empty:
            summ_pack.to_excel(writer, sheet_name="Riepilogo_Pack", index=False)
        # Sheet 2: Riepilogo Config
        if not summ_config.empty:
            summ_config.to_excel(writer, sheet_name="Riepilogo_Config", index=False)
        # Sheet 3: Riepilogo Base
        if not summ_base.empty:
            summ_base.to_excel(writer, sheet_name="Riepilogo_Base", index=False)
        # Sheet 4: Pivot Pack
        if not pivot_pack.empty:
            pivot_pack.to_excel(writer, sheet_name="Pivot_Pack")
        # Sheet 5: Pivot Config
        if not pivot_cfg.empty:
            pivot_cfg.to_excel(writer, sheet_name="Pivot_Config")
        # Sheet 6: Dettaglio completo
        detail_df.sort_values(
            ["Tipo_Optional", "Colonna_Optional", "Costo_SS_Allocato"],
            ascending=[True, True, False],
        ).to_excel(writer, sheet_name="Dettaglio", index=False)
        # Sheet 7: Verifica
        check_df.to_excel(writer, sheet_name="Verifica", index=False)

    print(f"\nFile salvato: {out_path}")

    # ── Stampa riepilogo ──────────────────────────────────────────
    print("\n--- Riepilogo Pack (top 10) ---")
    if not summ_pack.empty:
        print(summ_pack.head(10).to_string(index=False))
    print("\n--- Riepilogo Config (top 10) ---")
    if not summ_config.empty:
        print(summ_config.head(10).to_string(index=False))

    return detail_df, summ_pack, summ_config


if __name__ == "__main__":
    detail, summ_p, summ_c = run()

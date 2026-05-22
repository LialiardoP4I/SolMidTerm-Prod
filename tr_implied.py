"""
tr_implied.py
Post-processing script: calcola per ogni SKU il TR nominale e il TR massimo
impliciti nella simulazione Monte Carlo.

Nota: mean_demand e safety_stock_p99 sono espressi in MOTO (non in pezzi),
perché la moltiplicazione per qty_per_bike avviene solo per le colonne _total.
La qty non entra nel TR — è un moltiplicatore per i pezzi fisici, non per la
probabilità di montaggio.

TR nominale = mean_demand / N_moto_in_LT
           = take rate medio stimato dal modello

TR massimo  = (mean_demand + safety_stock_p99) / N_moto_in_LT
           = take rate coperto dalla safety (percentile 99)

N_moto_in_LT = somma del forecast sui mesi coperti dal lead time del componente
               (con interpolazione lineare per LT decimali, es. 2.3 mesi)
"""

import pandas as pd
import numpy as np
import math
import os

# ── Percorsi ─────────────────────────────────────────────────────────────────
INPUT_DIR  = ".\\Input"
OUTPUT_DIR = ".\\Output"

FORECAST_FILE  = os.path.join(INPUT_DIR,  "Total_demand.xlsx")
SKU_FILE       = os.path.join(OUTPUT_DIR, "sku_safety_stock_leadtime_PER_SKU_V2.xlsx")
MONTHLY_FILE   = os.path.join(OUTPUT_DIR, "sku_safety_stock_PER_MESE_V2.xlsx")
OUTPUT_FILE    = os.path.join(OUTPUT_DIR, "tr_implied.xlsx")

# Fallback se il V2 non esiste ancora
_SKU_FALLBACKS = [
    os.path.join(OUTPUT_DIR, "sku_safety_stock_leadtime_PER_SKU_OPTIMIZED.xlsx"),
    os.path.join(OUTPUT_DIR, "sku_safety_stock_leadtime_PER_SKU_V3.xlsx"),
]


# ── Helper: carica forecast ───────────────────────────────────────────────────

def load_forecast(filepath: str = FORECAST_FILE,
                  start_month_offset: int = 2) -> tuple[list[float], list[str]]:
    """
    Carica il forecast mensile dal file Excel.
    Restituisce (valori, nomi_mesi) a partire da start_month_offset.

    IMPORTANTE: parser allineato a mc_simulator_V2.load_forecast_demand()
    per garantire che il denominatore del TR implicito usi gli stessi
    valori usati dalla simulazione come target_demand.
    """
    # Stesso parser di mc_simulator_V2.py (header=None, iloc[0]=mesi, iloc[1]=valori)
    df = pd.read_excel(filepath, header=None)

    # Riga 0: nomi mesi (esclude colonne TOTALE, come il simulatore)
    month_row = df.iloc[0, 0:]
    # Bugfix: usare startswith per evitare falsi positivi (es. "OTTOBRE" contiene "TOT")
    month_names = [str(m).strip() for m in month_row
                   if pd.notna(m) and not str(m).strip().upper().startswith('TOT')]

    # Riga 1: valori forecast
    forecast_row = df.iloc[1, 0:]
    forecast_values = []
    for v in forecast_row[:len(month_names)]:
        try:
            fv = float(v)
            if not math.isnan(fv):
                forecast_values.append(fv)
            else:
                break  # stop al primo NaN
        except (ValueError, TypeError):
            break

    # Tronca month_names alla stessa lunghezza dei valori validi
    month_names = month_names[:len(forecast_values)]

    # Applica offset (es. salta JAN e FEB se la simulazione parte da MAR)
    forecast_values = forecast_values[start_month_offset:]
    month_names     = month_names[start_month_offset:]

    return forecast_values, month_names


# ── Helper: N moto nel lead time ─────────────────────────────────────────────

def n_moto_in_lt(lead_time: float, forecast_values: list[float]) -> float:
    """
    Somma il forecast per i mesi coperti dal lead time.
    Usa interpolazione lineare per la parte decimale.

    Esempi:
      LT=3.0 → forecast[0] + forecast[1] + forecast[2]
      LT=2.3 → forecast[0] + forecast[1] + 0.3×forecast[2]
      LT=0.25 → 0.25×forecast[0]
    """
    lt_full = int(lead_time)
    lt_frac = lead_time - lt_full

    total = 0.0

    # Mesi interi
    for i in range(min(lt_full, len(forecast_values))):
        total += forecast_values[i]

    # Mese parziale
    if lt_frac > 0 and lt_full < len(forecast_values):
        total += lt_frac * forecast_values[lt_full]

    return total


# ── Calcolo TR impliciti ──────────────────────────────────────────────────────

def compute_implied_tr(sku_file: str = SKU_FILE,
                       forecast_file: str = FORECAST_FILE,
                       start_month_offset: int = 2,
                       output_file: str = OUTPUT_FILE) -> pd.DataFrame:
    """
    Calcola TR nominale e TR massimo per ogni SKU.
    """
    # Carica output SKU (con fallback)
    if not os.path.exists(sku_file):
        for fb in _SKU_FALLBACKS:
            if os.path.exists(fb):
                print(f"[WARN] {sku_file} non trovato, uso fallback: {fb}")
                sku_file = fb
                break
        else:
            raise FileNotFoundError(f"Nessun file SKU trovato in {OUTPUT_DIR}")

    df = pd.read_excel(sku_file)
    print(f"SKU caricati: {len(df)}")

    # Carica forecast
    forecast_values, month_names = load_forecast(forecast_file, start_month_offset)
    print(f"Forecast caricato: {len(month_names)} mesi da {month_names[0] if month_names else '?'}")
    print(f"  Volumi: {[round(v) for v in forecast_values]}")

    # Colonne necessarie
    required = ['SKU', 'mean_demand', 'safety_stock_p99',
                'Quantity_Per_Bike', 'Lead_Time_Months_Ceil']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel file SKU: {missing}")

    results = []
    for _, row in df.iterrows():
        sku_id  = row['SKU']
        mean_d  = float(row['mean_demand'])
        ss_p99  = float(row['safety_stock_p99'])
        qty     = float(row['Quantity_Per_Bike'])
        lt      = float(row['Lead_Time_Months_Ceil'])

        # N moto nel periodo di lead time
        n_moto = n_moto_in_lt(lt, forecast_values)

        # Denominatore = solo moto attese (mean_demand è già in moto, non in pezzi)
        # La qty_per_bike NON entra nel TR: è un moltiplicatore per i pezzi fisici
        denom = n_moto if n_moto > 0 else None

        # TR nominale e massimo
        if denom and denom > 0:
            tr_nom = mean_d / denom
            tr_max = (mean_d + ss_p99) / denom
        else:
            tr_nom = None
            tr_max = None

        rec = {
            'SKU':                row.get('SKU'),
            'Description':        row.get('Description', ''),
            'Lead_Time_Months':   lt,
            'Quantity_Per_Bike':  qty,
            'Mean_Demand':        round(mean_d, 2),
            'Safety_Stock_p99':   round(ss_p99, 2),
            'N_Moto_in_LT':       round(n_moto, 1),
            'TR_Medio_%':         round(tr_nom * 100, 2) if tr_nom is not None else None,
            'TR_Massimo_%':       round(tr_max * 100, 2) if tr_max is not None else None,
            'Delta_TR_pp':        round((tr_max - tr_nom) * 100, 2)
                                  if (tr_nom is not None and tr_max is not None) else None,
            'Delta_TR_%':         round((tr_max - tr_nom) / tr_nom * 100, 2)
                                  if (tr_nom is not None and tr_max is not None and tr_nom > 0)
                                  else None,
        }

        # Aggiungi prezzo_safety se presente
        if 'prezzo_safety' in row:
            rec['Prezzo_Safety'] = row['prezzo_safety']

        results.append(rec)

    df_out = pd.DataFrame(results)

    # Ordina per TR medio decrescente
    df_out = df_out.sort_values('TR_Medio_%', ascending=False).reset_index(drop=True)

    # Salva
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='TR_Impliciti', index=False)

        # Sheet riepilogo per range
        bins   = [0, 10, 30, 50, 70, 90, 100]
        labels = ['0-10%', '10-30%', '30-50%', '50-70%', '70-90%', '90-100%']
        df_out['TR_Band'] = pd.cut(
            df_out['TR_Medio_%'].fillna(0),
            bins=bins, labels=labels, include_lowest=True
        )
        summary = df_out.groupby('TR_Band', observed=True).agg(
            N_SKU=('SKU', 'count'),
            TR_Medio_Medio=('TR_Medio_%', 'mean'),
            TR_Max_Medio=('TR_Massimo_%', 'mean'),
            SS_Medio=('Safety_Stock_p99', 'mean'),
        ).round(2)
        summary.to_excel(writer, sheet_name='Riepilogo_Band')

    print(f"\nSalvato: {output_file}")
    print(f"  SKU elaborati:      {len(df_out)}")
    print(f"  TR medio simulaz.:  {df_out['TR_Medio_%'].mean():.1f}%")
    print(f"  TR massimo medio:   {df_out['TR_Massimo_%'].mean():.1f}%")
    print(f"  Delta medio:        {df_out['Delta_TR_%'].mean():.1f}%")

    return df_out


# ============================================================================
# CALCOLO MENSILE (CORRETTO) — TR PER OGNI MESE
# ============================================================================

def compute_implied_tr_monthly(sku_monthly_file: str = MONTHLY_FILE,
                               forecast_file: str = FORECAST_FILE,
                               start_month_offset: int = 2,
                               output_file: str = None) -> pd.DataFrame:
    """
    Calcola TR nominale e TR massimo PER OGNI MESE per ogni SKU.

    Formula corretta:
      TR_mese = mean_demand_mese / forecast_mese
      dove:
        - mean_demand_mese = moto medie (dalla simulazione) che usano lo SKU IN QUEL MESE
        - forecast_mese    = moto totali previste IN QUEL MESE (denominatore corretto)

      TR_max_mese = (mean_demand_mese + safety_stock_p99_mese) / forecast_mese

    Nota: mean_demand e safety_stock sono in MOTO (non pezzi).
    La qty_per_bike NON entra nel TR.
    """
    if output_file is None:
        output_file = os.path.join(OUTPUT_DIR, "tr_implied_monthly.xlsx")

    # Carica file mensile
    if not os.path.exists(sku_monthly_file):
        raise FileNotFoundError(f"File mensile non trovato: {sku_monthly_file}")

    df = pd.read_excel(sku_monthly_file)
    print(f"[MONTHLY] SKU caricati: {len(df)}")

    # Carica forecast COMPLETO (offset=0) per allineamento mesi
    # I mesi del file PER_MESE usano tutti i mesi simulati indipendentemente dall'offset
    all_forecast_values, all_month_names = load_forecast(forecast_file, start_month_offset=0)
    print(f"[MONTHLY] Forecast completo: {len(all_month_names)} mesi")
    print(f"  Mesi: {all_month_names}")
    print(f"  Volumi: {[round(v) for v in all_forecast_values]}")

    # Mappa nome_mese_normalizzato → forecast_valore
    # Il file PER_MESE usa formato senza spazi: "Jan 2026" → "Jan2026"
    forecast_map = {}
    for name, val in zip(all_month_names, all_forecast_values):
        key = name.replace(' ', '')   # "Dec 2026" → "Dec2026"
        forecast_map[key] = val

    print(f"[MONTHLY] Forecast map: {forecast_map}")

    # Rileva colonne mean_demand_MESE
    mean_cols = [c for c in df.columns if c.startswith('mean_demand_')]
    if not mean_cols:
        raise ValueError("Nessuna colonna 'mean_demand_*' trovata nel file mensile")

    # Estrai e ordina mesi CRONOLOGICAMENTE usando la posizione nel forecast_map
    def _month_sort_key(col_name: str) -> int:
        month_key = col_name.replace('mean_demand_', '')
        keys = list(forecast_map.keys())
        return keys.index(month_key) if month_key in keys else 9999

    mean_cols_sorted = sorted(mean_cols, key=_month_sort_key)
    months_in_file = [c.replace('mean_demand_', '') for c in mean_cols_sorted]
    print(f"[MONTHLY] Mesi nel file (ordine cronologico): {months_in_file}")

    # Mesi senza corrispondenza nel forecast
    missing_months = [m for m in months_in_file if m not in forecast_map]
    if missing_months:
        print(f"[WARN] Mesi nel file senza forecast: {missing_months} (TR = None)")

    results = []

    for _, row in df.iterrows():
        sku_id = row['SKU']
        qty    = float(row['Quantity_Per_Bike'])
        lt     = float(row['Lead_Time_Months_Ceil'])

        rec = {
            'SKU':              sku_id,
            'Description':      row.get('Description', ''),
            'Lead_Time_Months': lt,
            'Quantity_Per_Bike': qty,
        }

        for month_name in months_in_file:
            mean_col = f'mean_demand_{month_name}'
            ss_col   = f'safety_stock_p99_{month_name}'

            if mean_col not in row.index or ss_col not in row.index:
                rec[f'TR_Medio_%_{month_name}']   = None
                rec[f'TR_Massimo_%_{month_name}'] = None
                continue

            mean_d = float(row[mean_col]) if pd.notna(row[mean_col]) else 0.0
            ss_p99 = float(row[ss_col])   if pd.notna(row[ss_col])   else 0.0

            # Denominatore: forecast DI QUEL MESE (non somma lead time)
            # TR = moto_con_SKU / moto_totali → take rate implicito
            forecast_mese = forecast_map.get(month_name)

            if forecast_mese and forecast_mese > 0:
                tr_nom = mean_d / forecast_mese
                tr_max = (mean_d + ss_p99) / forecast_mese
            else:
                tr_nom = None
                tr_max = None

            rec[f'TR_Medio_%_{month_name}']   = round(tr_nom * 100, 2) if tr_nom is not None else None
            rec[f'TR_Massimo_%_{month_name}'] = round(tr_max * 100, 2) if tr_max is not None else None

        results.append(rec)

    df_out = pd.DataFrame(results)

    # Calcola media TR su tutti i mesi (per ogni SKU)
    tr_cols_nom = [c for c in df_out.columns if c.startswith('TR_Medio_%_')
                   and not c.endswith('_Media')]
    tr_cols_max = [c for c in df_out.columns if c.startswith('TR_Massimo_%_')
                   and not c.endswith('_Media')]

    if tr_cols_nom:
        df_out['TR_Medio_%_Media'] = df_out[tr_cols_nom].mean(axis=1).round(2)
    if tr_cols_max:
        df_out['TR_Massimo_%_Media'] = df_out[tr_cols_max].mean(axis=1).round(2)

    # Ordina per TR medio decrescente
    if 'TR_Medio_%_Media' in df_out.columns:
        df_out = df_out.sort_values('TR_Medio_%_Media', ascending=False,
                                     na_position='last').reset_index(drop=True)

    # Salva
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='TR_Mensile', index=False)

        # Riepilogo
        summary_data = {
            'Metrica': [
                'SKU elaborati',
                'TR Medio medio (su tutti i mesi)',
                'TR Massimo medio',
            ],
            'Valore': [
                len(df_out),
                f"{df_out['TR_Medio_%_Media'].mean():.1f}%",
                f"{df_out['TR_Massimo_%_Media'].mean():.1f}%",
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Riepilogo', index=False)

    print(f"\n[MONTHLY] Salvato: {output_file}")
    print(f"  SKU elaborati: {len(df_out)}")
    print(f"  TR medio medio: {df_out['TR_Medio_%_Media'].mean():.1f}%")
    print(f"  TR massimo medio: {df_out['TR_Massimo_%_Media'].mean():.1f}%")

    return df_out


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    df = compute_implied_tr()
    print("\nPrime 10 righe:")
    print(df[['SKU', 'Description', 'Lead_Time_Months',
              'TR_Medio_%', 'TR_Massimo_%', 'Delta_TR_%']].head(10).to_string())

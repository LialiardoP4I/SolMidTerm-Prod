# Quantità per-configurazione (qty per-config) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Calcolare il fabbisogno di ogni componente con la **quantità corretta per configurazione** (somma delle qty a parità di dipendenza), invece dell'attuale scalare `max` per componente applicato dopo il conteggio.

**Architecture:** Portare `Qtà comp. (UMC)` attraverso la matrice BOM (nuova 5ª colonna metadata), sostituire il `drop_duplicates()` del catalogo con un `groupby(caratteristiche).sum(qty)` (somma a parità di dipendenza), e nel matching pesare ogni configurazione per la qty della riga-dipendenza che soddisfa (anziché moltiplicare per uno scalare). Il vettore domanda diventa già in **pezzi**.

**Tech Stack:** Python 3.13, pandas, numpy, openpyxl, pytest. Package `mid_term_supermodels`.

**Razionale e validazioni a monte (già fatte):**
- Dipendenze diverse dello stesso componente sono **mutuamente esclusive**: 0 coppie sovrapponibili su 1177 componenti (Multistrada 461, DVL 135, PANI 581). → nel matching ogni config attiva al più una riga-dipendenza → niente doppio conteggio.
- Caso patologico (testo dipendenza riordinato → stessa caratteristica, qty diverse): 3 casi in V21E, 0 in V22E/V24T. Risolto sommando per **caratteristiche** (canonico), non per testo grezzo.

> **VINCOLO (CLAUDE.md di SolMidTerm - Prod):** `bom.py` contiene logica di business protetta. Questo piano LO MODIFICA su autorizzazione esplicita dell'utente. Ogni modifica va fatta con diff PRIMA/DOPO e test; non alterare la logica oltre quanto specificato.

> **Tutti i path sono relativi a** `SolMidTerm - Prod/mid_term_supermodels/`. Comandi `python`/`pytest` da dentro quella cartella.

---

## Punti di codice coinvolti (riferimento, file:line attuali)

`bom.py`:
- `col_structure` (~1900-1917): 4 colonne metadata fisse + mapping + calcolate.
- `_populate_row_from_condition_v2` (604): mapping via indici positivi da `enumerate(col_structure)`; calcolate via indici **negativi** dalla fine.
- hardcoded `mapping_cols_count = len(col_structure) - len(calculated_cols) - 4` (~1923) e `range(4, mapping_cols_count + 4)` (~1934).
- loop righe BOM (~2185-2310): 4 punti `row = ['' for _ in col_structure]`; legge `condition` e `numero_componenti`, **non** la qty.
- save matrice (~2398-2410): header 3 righe da `col_structure`, dati da `matrix_rows`.
- `_rileva_colonne` (2677-2783): `N_METADATA = 4` (2708); classifica colonne ≥4 come calcolate se bundle+desc vuoti.
- `crea_tutte_righe_univoche_v2` (2790+): `colonne_selezionate = [col_num_componenti] + col_modello + col_calcolate_usate` (2875); `df_result.drop_duplicates()` (3041); FASE 5.5 espansione (3050+); save (3092+).

`matching.py`:
- `load_lead_times` (~879-896): estrae qty da BOM con `groupby('SKU').max()`.
- `_build_matching_context` (~1623-1696): merge lead_times (qty scalare), `sku_data['values']` + `'quantity'` (mean).
- `_compute_run_demands_for_sku` (1708-1727): OR delle maschere + `aggregate` × scalare.
- `aggregate_monthly_demands_numpy` (1112-1179): somma config con peso 1.
- `match_skus_ultra_optimized` (~1786): `_total = ss * quantity`.
- `match_skus_preserve_run_vectors` (1799-1828): ritorna (bici, qty, lt, desc).
- `pool_safety_stock` (1831-1907): `pezzi_sm = run_demands * qty` (1882).

---

## Strategia di test

`bom.py` non è unit-testabile in isolamento (richiede mapping/TR/BOM reali). Per le sue modifiche si usa **validazione su dati reali**: eseguire la pipeline su un supermodel e verificare il catalogo prodotto (`Output/<SM>/tutte_righe_univoche_V2.xlsx`) — con qty attesa nota (`15621191E` → 5; i 3 patologici sommati). Per `matching.py` si usano **test unit con dati sintetici** (catalogo+qty per-riga costruiti a mano).

---

## Task 1: bom.py — portare la qty nella matrice (5ª colonna metadata)

**Files:** Modify `bom.py` (`col_structure`, i due hardcoded `4`, i 4 `row` del loop, `N_METADATA`)

- [ ] **Step 1: Leggere e confermare gli offset**

Run:
```bash
grep -n "col_structure = \[" bom.py
grep -n "N_METADATA = 4" bom.py
grep -n "len(calculated_cols) - 4\|range(4, " bom.py
```
Annota le righe reali. Confermare che gli hardcoded `4` siano solo i 2 segnalati (`mapping_cols_count` e il `range`). Se ne trovi altri (`grep -n "- 4\b\|, 4)" bom.py` nelle vicinanze di col_structure/calc), includili nello Step 3.

- [ ] **Step 2: Aggiungere la 5ª colonna metadata in `col_structure`**

PRIMA (`bom.py` ~1900):
```python
    col_structure = [
        ('', '', 'Numero componenti'),
        ('', '', 'Condizione_Originale'),
        ('', '', 'Condizione_Espansa'),
        (start_col, '', '')
    ]
```
DOPO:
```python
    col_structure = [
        ('', '', 'Numero componenti'),
        ('', '', 'Condizione_Originale'),
        ('', '', 'Condizione_Espansa'),
        (start_col, '', ''),
        ('', '', 'Qta_comp_UMC'),   # 5a colonna metadata: quantità per riga BOM
    ]
```
(Usiamo `Qta_comp_UMC` come header leggibile per evitare problemi di encoding con l'accentata; il valore numerico è ciò che conta.)

- [ ] **Step 3: Bumpare i metadata count da 4 a 5**

In `_rileva_colonne` (~2708): `N_METADATA = 4` → `N_METADATA = 5`.
In `process_excel_to_matrix_v2` (~1923): `mapping_cols_count = len(col_structure) - len(calculated_cols) - 4` → `... - 5`.
In (~1934): `range(4, mapping_cols_count + 4)` → `range(5, mapping_cols_count + 5)`.
(Applicare anche ad eventuali altri `4` trovati allo Step 1 che si riferiscono al numero di colonne metadata.)

- [ ] **Step 4: Leggere la qty da ogni riga BOM e assegnarla a `row[4]`**

Nel loop (~2187), subito dopo aver ricavato `numero_componenti` (~2193-2195), aggiungere la lettura qty:
```python
        qta_umc = bom_data.loc[bom_index, 'Qtà comp. (UMC)'] if 'Qtà comp. (UMC)' in bom_data.columns else 0
        try:
            qta_umc = float(qta_umc)
        except (TypeError, ValueError):
            qta_umc = 0.0
```
Poi, in **ognuno** dei 4 punti `row = ['' for _ in col_structure]` (i blocchi: no-condition ~2198, normale ~2224, split ~2248, discriminant ~2285), aggiungere subito dopo gli assegnamenti `row[0..3]`:
```python
            row[4] = qta_umc
```
(L'indice 4 è la nuova 5ª colonna. I write su mapping/calcolate restano corretti: i mapping usano indici live da `enumerate(col_structure)`, le calcolate usano indici negativi dalla fine — entrambi non rotti da una colonna inserita in posizione 4.)

- [ ] **Step 5: Validazione su dati reali (matrice)**

Run (dalla cartella package; usa un supermodel mono-modello per velocità, es. DVL non ha la qty multipla famosa; usa Multistrada V21E per `15621191E`). Esegui solo lo step 1 della pipeline tramite uno script temporaneo, oppure il run completo multi-supermodel breve. Poi:
```bash
python -c "
import pandas as pd
df = pd.read_excel('Output/Multistrada/matrix_output_V21E.xlsx', sheet_name='Matrix', header=2)
print('colonna qty presente:', 'Qta_comp_UMC' in df.columns)
sub = df[df.iloc[:,0]=='15621191E']
print('valori qty per 15621191E:', sorted(sub['Qta_comp_UMC'].dropna().unique().tolist()))
"
```
Expected: `colonna qty presente: True`; valori qty includono `1.0` e `4.0` (pre-somma, ancora per-riga). Se la colonna non c'è o le caratteristiche risultano spostate → offset rotto, rivedere Step 3.

- [ ] **Step 6: Commit**

```bash
git add bom.py
git commit -m "feat(bom): porta Qtà comp. (UMC) nella matrice come 5a colonna metadata"
```

---

## Task 2: bom.py — somma qty per DIPENDENZA (a monte) + qty nel catalogo

> **CORREZIONE (validazione 2026-06-25):** l'approccio "groupby caratteristiche-catalogo + sum" è SBAGLIATO: somma su dimensioni NON presenti nel catalogo (mercato `ZCOUNTRY`, plant `ZPLT`). Validato: `15621191E` risultava 40 (= 5 × ~8 mercati) invece di 5, perché ha caratteristiche-config vuote e dipendenza solo su mercato → tutte le righe-mercato collassano nel catalogo e venivano sommate. **Soluzione corretta:** sommare la qty per `(componente, dipendenza COMPLETA canonicalizzata)` **a monte** (sul BOM elaborato, PRIMA dell'espansione mercati), poi `drop_duplicates` NORMALE al catalogo (i mercati sono alternative: stessa qty su righe diverse → tenuta una volta).

**Files:** Modify `bom.py` (`process_excel_to_matrix_v2`: pre-somma su `bom_data`; `crea_tutte_righe_univoche_v2`: include qty + `drop_duplicates` invariato)

- [ ] **Step 0: Resettare le modifiche errate di T2**

bom.py contiene le modifiche errate (groupby+sum) NON committate. Task 1 è committato (`a9ae1e8`). Resettare bom.py a Task 1:
```bash
git checkout -- mid_term_supermodels/bom.py   # riporta bom.py al commit a9ae1e8 (solo Task 1)
```
Verificare: `git diff bom.py` vuoto; `grep -n "Qta_comp_UMC" bom.py` mostra ancora le modifiche Task 1 (col_structure + row[4]).

- [ ] **Step 1: Pre-somma qty per (componente, dipendenza canonica) a monte**

In `process_excel_to_matrix_v2`, SUBITO DOPO `bom_data = df.copy()` (~bom.py:1833) e PRIMA del loop `for bom_row in bom_data.itertuples()` (~2188), inserire:
```python
    # PRE-SOMMA qty a parità di (componente, dipendenza): somma le righe con la
    # STESSA condizione (posizioni diverse della distinta) PRIMA dell'espansione
    # mercati. La chiave è la condizione CANONICALIZZATA (predicati ordinati,
    # spazi normalizzati) per unire dipendenze logicamente identiche scritte in
    # modo diverso (es. predicati riordinati). NON somma su mercati: la condizione
    # completa (incl. ZCOUNTRY/ZPLT) resta la chiave.
    import re as _re
    _qcol_pre = next((c for c in bom_data.columns if str(c).startswith('Qt') and 'UMC' in str(c)), None)
    if _qcol_pre is not None and start_col in bom_data.columns and 'Numero componenti' in bom_data.columns:
        def _canon_cond(_c):
            if pd.isna(_c):
                return ''
            _s = _re.sub(r'\s+', ' ', str(_c)).strip()
            if ' OR ' not in _s.upper():          # solo congiunzioni AND -> ordina i predicati
                _parts = [p.strip() for p in _re.split(r'\bAND\b', _s)]
                _s = ' AND '.join(sorted(_parts))
            return _s
        _orig_cols = list(bom_data.columns)
        bom_data = bom_data.copy()
        bom_data['_canon_cond'] = bom_data[start_col].map(_canon_cond)
        bom_data[_qcol_pre] = pd.to_numeric(bom_data[_qcol_pre], errors='coerce').fillna(0.0)
        _agg_map = {_qcol_pre: 'sum'}
        for _c in _orig_cols:
            if _c not in ('Numero componenti', _qcol_pre):
                _agg_map[_c] = 'first'
        bom_data = (bom_data.groupby(['Numero componenti', '_canon_cond'],
                                     dropna=False, as_index=False)
                    .agg(_agg_map))
        bom_data = bom_data[_orig_cols]   # ripristina colonne/ordine originali per il loop
```
Effetto: per `15621191E` (pos15 qty1 + pos20 qty4, stessa condizione) → una riga con qty **5**; i 3 patologici (testo riordinato) → canonica identica → sommati (15, 152, 3). I mercati restano righe separate (condizione diversa) con la loro qty 5 ciascuno.

- [ ] **Step 2: Includere la qty nel catalogo (crea_tutte_righe_univoche_v2)**

In `crea_tutte_righe_univoche_v2` (~2875):
PRIMA:
```python
    colonne_da_esplodere = col_modello + col_calcolate_usate
    colonne_selezionate  = [col_num_componenti] + colonne_da_esplodere
    df_result = df[colonne_selezionate].copy()
```
DOPO:
```python
    colonne_da_esplodere = col_modello + col_calcolate_usate
    _qty_col = 'Qta_comp_UMC' if 'Qta_comp_UMC' in df.columns else None
    colonne_selezionate  = [col_num_componenti] + colonne_da_esplodere + ([_qty_col] if _qty_col else [])
    df_result = df[colonne_selezionate].copy()
    if _qty_col:
        df_result[_qty_col] = pd.to_numeric(df_result[_qty_col], errors='coerce').fillna(0.0)
```
Nota: `_rileva_colonne` con `N_METADATA=5` esclude la qty (posizione 4) dalle calcolate/mapping; qui la ri-includiamo per nome.

- [ ] **Step 3: `drop_duplicates` INVARIATO (NON groupby+sum)**

Il `drop_duplicates()` (~3041) resta **standard**. Con la pre-somma a monte, le righe di mercati diversi hanno la STESSA qty (5) e stesse caratteristiche-config → `drop_duplicates` le fonde in una (qty 5). NON modificare questa riga. Aggiungere SOLO un check di sicurezza subito dopo:
```python
    if 'Qta_comp_UMC' in df_result.columns:
        _ch = [c for c in df_result.columns if c != 'Qta_comp_UMC']
        _dups = df_result.duplicated(subset=_ch, keep=False)
        if _dups.any():
            log.warning("[QTY] %d righe con stesse caratteristiche ma qty diversa "
                        "(qty dipende da mercato/plant non nel catalogo): verificare",
                        int(_dups.sum()))
```
(Se scatta: esistono componenti la cui qty dipende da mercato/plant — dimensione non simulata dal MC; va deciso a parte. Per i dati attuali non dovrebbe scattare su `15621191E`.)

- [ ] **Step 4: FASE 5.5 invariata**

FASE 5.5 (~3050) usa `.assign(...)` (preserva la qty) + `drop_duplicates()` post-espansione. Lasciare invariata.

- [ ] **Step 5: Validazione su dati reali (catalogo) — via step 1 isolato (veloce)**

Il run single-month crasha sugli orizzonti diversi; valida SOLO lo step 1 di Multistrada con questo script (da dentro la cartella package):
```bash
python - <<'PY'
import sys, os
from pathlib import Path
sys.path.insert(0, str(Path('..').resolve()))
from mid_term_supermodels import multi_supermodel as ms
from mid_term_supermodels.config import PipelineConfig, load_simulation_config
from mid_term_supermodels.montecarlo import load_monthly_forecast
from mid_term_supermodels.pipeline import SafetyStockPipeline
input_dir=Path('Input').resolve(); work=ms._new_work_root()
work_sm=ms.stage_supermodel_input(str(input_dir/'Multistrada'), str(input_dir), str(work))
cwd=os.getcwd(); out=Path(cwd)/'Output'/'Multistrada'
try:
    os.chdir(work_sm); cfg=PipelineConfig(input_dir=Path('Input'), output_dir=out)
    os.makedirs(out, exist_ok=True)
    _,m=load_monthly_forecast(str(cfg.resolve_input(cfg.forecast_file)))
    sim=load_simulation_config(str(Path(cwd)/'run_config.json'), m)
    SafetyStockPipeline(cfg, sim, validate_config=False).run_step_1_build_bom()
finally:
    os.chdir(cwd); ms._robust_rmtree(work)
import pandas as pd
cat=pd.read_excel(out/'tutte_righe_univoche_V2.xlsx')
print('15621191E qty:', sorted(pd.to_numeric(cat[cat['Numero componenti']=='15621191E']['Qta_comp_UMC'],errors='coerce').dropna().unique().tolist()))
for code in ['72010513C','81510411AA','83340031A']:
    print(code, sorted(pd.to_numeric(cat[cat['Numero componenti']==code]['Qta_comp_UMC'],errors='coerce').dropna().unique().tolist()))
PY
```
Expected: `15621191E qty: [5.0]` (NON 40); `72010513C: [15.0]`, `81510411AA: [152.0]`, `83340031A: [3.0]`. Se compaiono valori multipli o 40 → la pre-somma non ha agito o la chiave canonica è errata.

- [ ] **Step 6: Commit**

```bash
git add bom.py
git commit -m "feat(bom): somma qty per dipendenza canonica a monte + qty nel catalogo (drop_duplicates invariato)"
```

---

## Task 3: matching.py — portare la qty per-riga nel contesto di matching

**Files:** Modify `matching.py` (`_build_matching_context`, merge lead_times), Test `tests/test_qty_context.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_qty_context.py`:
```python
import numpy as np, pandas as pd
from mid_term_supermodels import matching


def _sim(n_runs=2):
    rows=[{'MODEL':'MSV4','JAN 2026_Run_0':10,'JAN 2026_Run_1':10},
          {'MODEL':'MSV4S','JAN 2026_Run_0':4,'JAN 2026_Run_1':4}]
    return pd.DataFrame(rows)

def test_context_attaches_per_row_qty():
    sim=_sim()
    # catalogo con qty per riga (colonna Qta_comp_UMC)
    cat=pd.DataFrame({'Numero componenti':['X','X'],'MODEL':['MSV4','MSV4S'],
                      'Qta_comp_UMC':[5.0,3.0]})
    lt=pd.DataFrame({'SKU':['X'],'Lead_Time_Months':[1.0],'Lead_Time_Months_Ceil':[1.0],
                     'Quantity_Per_Bike':[1.0],'Description':['c']})
    mapping=matching.build_column_mapping_auto(['MODEL'],list(sim.columns))
    ctx=matching._build_matching_context(cat,lt,sim,mapping,n_runs=2)
    try:
        d=next(dd for sid,dd in zip(ctx.sku_ids,ctx.sku_data_list) if sid=='X')
        # ogni entry in values deve avere una qty allineata in qtys
        assert 'qtys' in d and len(d['qtys'])==len(d['values'])
        assert sorted(d['qtys'])==[3.0,5.0]
    finally:
        ctx.cleanup()
```

- [ ] **Step 2: Eseguire (fallisce: niente 'qtys')**

Run: `python -m pytest tests/test_qty_context.py -v` → FAIL.

- [ ] **Step 3: Far arrivare la qty dal catalogo (non più da lead_times)**

In `_build_matching_context`, il merge lead_times NON deve più portare `Quantity_Per_Bike` (ora la qty è nel catalogo come `Qta_comp_UMC`). Modificare il merge (~1623):
PRIMA:
```python
    _sku_merged_all = sku_catalog.merge(
        lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                    'Quantity_Per_Bike', 'Description']],
        left_on='Numero componenti', right_on='SKU', how='left')
```
DOPO:
```python
    _lt_cols = ['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil', 'Description']
    _sku_merged_all = sku_catalog.merge(
        lead_times[_lt_cols],
        left_on='Numero componenti', right_on='SKU', how='left')
    # qty per-riga dal catalogo (Qta_comp_UMC); fallback 1.0 se assente
    if 'Qta_comp_UMC' in _sku_merged_all.columns:
        _sku_merged_all['Qta_comp_UMC'] = pd.to_numeric(
            _sku_merged_all['Qta_comp_UMC'], errors='coerce').fillna(1.0)
    else:
        _sku_merged_all['Qta_comp_UMC'] = 1.0
```

- [ ] **Step 4: Costruire `qtys` allineato a `values`**

Nel loop di `_build_matching_context` (~1662), modificare la costruzione per accumulare la qty per riga, replicandola sull'espansione multi-modello:
PRIMA (estratto):
```python
        for _, row in sku_rows_group.iterrows():
            row_values = {}
            for sku_col in column_mapping.keys():
                if sku_col in row.index:
                    row_values[sku_col] = row[sku_col]
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
```
DOPO (aggiungere `rows_qtys` parallelo):
```python
        rows_qtys = []
        for _, row in sku_rows_group.iterrows():
            row_values = {}
            for sku_col in column_mapping.keys():
                if sku_col in row.index:
                    row_values[sku_col] = row[sku_col]
            _q = float(row['Qta_comp_UMC']) if 'Qta_comp_UMC' in row.index else 1.0
            if model_sku_col and model_sku_col in row_values:
                model_val = str(row_values[model_sku_col]).strip()
                if '+' in model_val:
                    for single_model in model_val.split('+'):
                        expanded = dict(row_values)
                        expanded[model_sku_col] = single_model.strip()
                        rows_values.append(expanded)
                        rows_qtys.append(_q)
                else:
                    rows_values.append(row_values)
                    rows_qtys.append(_q)
            else:
                rows_values.append(row_values)
                rows_qtys.append(_q)
```
E nel dict `sku_data_list.append({...})` aggiungere `'qtys': rows_qtys` e cambiare `'quantity'` in valore rappresentativo per il report:
```python
        sku_data_list.append({
            'values': rows_values,
            'qtys': rows_qtys,
            'lead_time': float(sku_rows_group['Lead_Time_Months_Ceil'].max()),
            'quantity': float(max(rows_qtys)) if rows_qtys else 1.0,  # solo per report
            'description': sku_rows_group['Description'].iloc[0]
                           if 'Description' in sku_rows_group.columns else '',
            'lead_time_raw': sku_rows_group['Lead_Time_Months'].iloc[0]
                             if 'Lead_Time_Months' in sku_rows_group.columns else np.nan
        })
```
Rimuovere il `sku_with_lt['Quantity_Per_Bike'] = ...fillna(1.0)` (non più usato) se presente nel blocco.

- [ ] **Step 5: Eseguire il test**

Run: `python -m pytest tests/test_qty_context.py -v` → PASS. Poi `python -m pytest tests/ -q` → tutti PASS (i test esistenti del matching non usano 'qtys'; verificare non rotti).

- [ ] **Step 6: Commit**

```bash
git add matching.py tests/test_qty_context.py
git commit -m "feat(matching): qty per-riga dal catalogo nel matching context"
```

---

## Task 4: matching.py — pesare ogni configurazione per la sua qty

**Files:** Modify `matching.py` (`_compute_run_demands_for_sku`), Test `tests/test_qty_weighting.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_qty_weighting.py`:
```python
import numpy as np, pandas as pd
from mid_term_supermodels import matching

def _sim():
    # 2 config: MSV4 (10 bici), MSV4S (4 bici), 2 run
    return pd.DataFrame([
        {'MODEL':'MSV4','JAN 2026_Run_0':10,'JAN 2026_Run_1':10},
        {'MODEL':'MSV4S','JAN 2026_Run_0':4,'JAN 2026_Run_1':4}])

def test_per_config_qty_weighting():
    sim=_sim()
    # componente X: su MSV4 qty 5, su MSV4S qty 3 (dipendenze esclusive via MODEL)
    cat=pd.DataFrame({'Numero componenti':['X','X'],'MODEL':['MSV4','MSV4S'],
                      'Qta_comp_UMC':[5.0,3.0]})
    lt=pd.DataFrame({'SKU':['X'],'Lead_Time_Months':[1.0],'Lead_Time_Months_Ceil':[1.0],
                     'Description':['c']})
    mapping=matching.build_column_mapping_auto(['MODEL'],list(sim.columns))
    ctx=matching._build_matching_context(cat,lt,sim,mapping,n_runs=2)
    try:
        d=next(dd for sid,dd in zip(ctx.sku_ids,ctx.sku_data_list) if sid=='X')
        rd,idx=matching._compute_run_demands_for_sku(d,ctx)
        # pezzi attesi per run = 10*5 (MSV4) + 4*3 (MSV4S) = 62
        assert rd.tolist()==[62.0,62.0]
    finally:
        ctx.cleanup()
```

- [ ] **Step 2: Eseguire (fallisce: oggi pesa 1 e moltiplica scalare)**

Run: `python -m pytest tests/test_qty_weighting.py -v` → FAIL (oggi ritorna bici non pesate).

- [ ] **Step 3: Riscrivere `_compute_run_demands_for_sku` pesato per qty**

PRIMA (1708-1727): vedi riferimento. DOPO:
```python
def _compute_run_demands_for_sku(sku_data, ctx):
    """Calcola (run_demands_PEZZI[n_runs], matching_indices) per uno SKU.

    Per ogni riga-dipendenza del componente pesa le configurazioni che la
    soddisfano per la sua qty (Qta_comp_UMC), e somma. Dato che le dipendenze
    diverse sono mutuamente esclusive (verificato), ogni config è attivata da
    una sola riga -> nessun doppio conteggio. run_demands è già in PEZZI.
    """
    total = np.zeros(ctx.n_runs, dtype=np.float32)
    assigned = np.zeros(ctx.n_rows, dtype=bool)
    qtys = sku_data.get('qtys') or [1.0] * len(sku_data['values'])
    overlap_cfgs = 0
    for row_values, qty_r in zip(sku_data['values'], qtys):
        row_mask = sku_matches_combination_numpy(
            row_values, ctx.norm_arrays, ctx.category_maps, ctx.column_mapping, ctx.n_rows
        )
        idx = np.where(row_mask)[0]
        if len(idx) == 0:
            continue
        # check sovrapposizione: una config già assegnata da un'altra dipendenza
        ov = int(assigned[idx].sum())
        if ov:
            overlap_cfgs += ov
        assigned[idx] = True
        agg = aggregate_monthly_demands_numpy(
            idx, ctx.data_matrix, ctx.month_col_indices, sku_data['lead_time'], ctx.n_runs
        )
        total += np.asarray(agg, dtype=np.float32) * float(qty_r)
    matching_indices = np.where(assigned)[0]
    if len(matching_indices) == 0:
        return np.zeros(ctx.n_runs, dtype=np.float32), matching_indices
    if overlap_cfgs:
        log.warning("[QTY] SKU %s: %d configurazioni attivate da >1 dipendenza "
                    "(sovrapposizione) -> qty potenzialmente sommata due volte",
                    sku_data.get('description', '?'), overlap_cfgs)
    return total, matching_indices
```
Nota: nel caso di sovrapposizione (oggi 0) la qty verrebbe sommata; il warning lo segnala. Per i dati attuali è sicuro.

- [ ] **Step 4: Eseguire il test**

Run: `python -m pytest tests/test_qty_weighting.py -v` → PASS.

- [ ] **Step 5: Commit**

```bash
git add matching.py tests/test_qty_weighting.py
git commit -m "feat(matching): pesa ogni configurazione per la qty della dipendenza (pezzi)"
```

---

## Task 5: matching.py — rimuovere la moltiplicazione scalare (output già in pezzi)

**Files:** Modify `matching.py` (`match_skus_ultra_optimized`)

- [ ] **Step 1: Rimuovere `* sku_data['quantity']`**

PRIMA (~1786):
```python
            result.update(stats)
            result[f'{ss_key}_total'] = result[ss_key] * sku_data['quantity']
```
DOPO:
```python
            result.update(stats)
            # run_demands è già in PEZZI (qty per-config applicata nel matching):
            # _total coincide con la safety stock in pezzi.
            result[f'{ss_key}_total'] = result[ss_key]
```
La riga `'Quantity_Per_Bike': sku_data['quantity']` resta (ora = qty rappresentativa/max, solo informativa).

- [ ] **Step 2: Verificare la suite**

Run: `python -m pytest tests/ -q` → tutti PASS (i test esistenti del matching non dipendono dal `*quantity`; se un test asseriva `_total = ss*qty` con qty>1, aggiornarlo a `_total == ss`).

- [ ] **Step 3: Commit**

```bash
git add matching.py
git commit -m "feat(matching): _total in pezzi (rimossa moltiplicazione scalare qty)"
```

---

## Task 6: matching.py — preserve_run_vectors e pooling in pezzi

**Files:** Modify `matching.py` (`match_skus_preserve_run_vectors`, `pool_safety_stock`), Test: aggiornare `tests/test_pooling.py`, `tests/test_preserve_vectors.py`, `tests/test_validation.py`, `tests/test_save_pooled.py`

- [ ] **Step 1: `match_skus_preserve_run_vectors` ritorna pezzi, qty=1.0**

In `match_skus_preserve_run_vectors` (~1820), il vettore è ora già in pezzi:
PRIMA:
```python
            vectors[sku_id] = (np.array(run_demands, dtype=np.float32, copy=True),
                               float(sku_data['quantity']),
                               float(sku_data['lead_time']),
                               sku_data['description'])
```
DOPO:
```python
            # run_demands è già in PEZZI: qty=1.0 (la moltiplicazione nel pool è neutra).
            vectors[sku_id] = (np.array(run_demands, dtype=np.float32, copy=True),
                               1.0,
                               float(sku_data['lead_time']),
                               sku_data['description'])
```

- [ ] **Step 2: `pool_safety_stock` non moltiplica più per qty**

PRIMA (~1882):
```python
            pezzi_sm = np.asarray(run_demands, dtype=np.float64) * qty
```
DOPO:
```python
            # run_demands è già in PEZZI (qty per-config applicata a monte).
            pezzi_sm = np.asarray(run_demands, dtype=np.float64)
```
(`qty` resta nella tupla solo per il breakdown; con qty=1.0 anche un eventuale uso resta neutro.)

- [ ] **Step 3: Aggiornare i test sintetici esistenti**

I test `test_pooling.py`, `test_validation.py`, `test_save_pooled.py` costruiscono `sm_results` con tuple `(vettore, qty, lt, desc)` e si aspettavano `pezzi = vettore*qty`. Ora il vettore è già in pezzi e il pool non moltiplica. Aggiornare le aspettative: i vettori di test vanno interpretati come **pezzi** e qty messa a `1.0`. Esempio per `test_pooling.py::test_pooling_sumproduct_and_risk_pooling`:
```python
    sm_results = {
        'SuperA': {'X': (np.array([10,10,10,10],dtype=np.float32), 1.0, 1.0, 'x')},
        'SuperB': {'X': (np.array([0,20,40,80],dtype=np.float32), 1.0, 1.0, 'x')},  # già pezzi (era bici*qty)
    }
```
e aggiornare i valori attesi di conseguenza (pooled = somma diretta dei vettori). Applicare lo stesso ragionamento agli altri test (qty→1.0, vettori già in pezzi).

- [ ] **Step 4: Eseguire la suite**

Run: `python -m pytest tests/ -q` → tutti PASS.

- [ ] **Step 5: Commit**

```bash
git add matching.py tests/
git commit -m "feat(matching): preserve_vectors e pool in pezzi (qty già applicata per-config)"
```

---

## Task 7: matching.py — load_lead_times non estrae più la qty

**Files:** Modify `matching.py` (`load_lead_times`)

- [ ] **Step 1: Rimuovere l'estrazione/merge qty da BOM**

In `load_lead_times` (~879-896) la qty veniva estratta dal BOM con `groupby max` e mergeata. Ora la qty viaggia nel catalogo, quindi questo non serve più per il matching. Rimuovere il blocco `bom_qty`/`bom_qty_unique` e il merge della qty, lasciando il resto (residual LT, Description) invariato. Se altre parti (es. output report) si aspettano `Quantity_Per_Bike` in `lead_times`, lasciare una colonna costante `1.0` per compatibilità:
```python
    # qty non più gestita qui: è per-configurazione nel catalogo (Qta_comp_UMC).
    lead_times['Quantity_Per_Bike'] = 1.0
```
Verificare che `_build_matching_context` non richieda più `Quantity_Per_Bike` da lead_times (Task 3 l'ha tolto dal merge).

- [ ] **Step 2: Verificare la suite + import**

Run: `python -m pytest tests/ -q` → PASS. `cd .. && python -c "from mid_term_supermodels import matching, bom, pipeline; print('ok')"`.

- [ ] **Step 3: Commit**

```bash
git add matching.py
git commit -m "refactor(matching): load_lead_times non gestisce più la qty (ora per-config nel catalogo)"
```

---

## Task 8: Validazione end-to-end su dati reali

**Files:** Test `tests/test_qty_e2e_doc.md` (note) — esecuzione manuale

- [ ] **Step 1: Run completo multi-supermodel**

Con `run_config.json` `multi_supermodel: true`, `rolling: true`, `n_runs: 10`:
```bash
python run_pipeline.py > _run_qty.log 2>&1
```
Expected: exit 0, output prodotti in `Output/`.

- [ ] **Step 2: Verificare la qty nel catalogo e nei risultati**

```bash
python -c "
import pandas as pd
cat=pd.read_excel('Output/Multistrada/tutte_righe_univoche_V2.xlsx')
s=cat[cat['Numero componenti']=='15621191E']
print('15621191E qty nel catalogo:', sorted(s['Qta_comp_UMC'].dropna().unique().tolist()))  # atteso include 5.0
pool=pd.read_excel('Output/sku_safety_stock_POOLED_rolling_supermodels.xlsx')
print('15621191E nel pooled:', (pool['SKU']=='15621191E').sum(), 'righe')
"
```
Expected: catalogo `15621191E` → riga con qty 5.0; presente nel pooled.

- [ ] **Step 3: Confronto SS prima/dopo (sanity)**

Confrontare la SS pooled di alcuni componenti a qty variabile (es. `15621191E`, `069191030`) prima (max) e dopo (per-config). La SS dei componenti cumulativi dovrebbe **aumentare** dove `max` sottostimava (1+4=5 > 4) e **diminuire** dove `max` sovrastimava (alternative). Documentare 3-4 esempi nel log/commit.

- [ ] **Step 4: Commit della validazione**

```bash
git add _run_qty.log docs/plans/2026-06-25-qty-per-configurazione.md
git commit -m "test: validazione end-to-end qty per-configurazione"
```

---

## Note finali / rischi

- **bom.py offset (Task 1)**: il rischio maggiore. Se dopo Step 5 di Task 1 le caratteristiche risultano spostate, c'è un hardcoded `4` non aggiornato — rifare `grep -n "\b4\b" bom.py` nelle funzioni `process_excel_to_matrix_v2` e `_rileva_colonne`.
- **Performance matching (Task 4)**: ora `aggregate` è chiamato una volta per riga-dipendenza invece di una sola con OR. Per SKU con molte dipendenze è più lento; accettabile (il MC resta il collo di bottiglia), ma monitorare il tempo dello step 3.
- **Semantica colonne output**: `safety_stock_p99` e `_total` ora sono entrambi in **pezzi**; `Quantity_Per_Bike` nell'output diventa solo informativo (max delle qty per-config). Documentarlo per chi legge gli Excel.
- **Sovrapposizioni**: oggi 0; il warning `[QTY] ... sovrapposizione` in `_compute_run_demands_for_sku` segnala se dati futuri le introducono.
- **DVL/PANI**: i BOM grezzi vengono processati al run; verificare che la colonna `Qtà comp. (UMC)` sia presente anche nei loro BOM_150 generati (lo è, è la stessa colonna SAP).

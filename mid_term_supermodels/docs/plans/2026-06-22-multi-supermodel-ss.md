# Safety Stock multi-supermodel — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Calcolare la safety stock condivisa (pooled) tra più supermodel con domanda indipendente, sfruttando il risk pooling sui componenti condivisi.

**Architecture:** Nuova cartella `mid_term_supermodels/` (copia di `mid_term/`). L'attuale `SafetyStockPipeline.run()` diventa la run di UN supermodel; un orchestratore la esegue per ogni cartella supermodel (seed diverso), preserva i vettori di domanda per-run per componente, e fa il pooling somma-prodotto per-run a livello componente. Output: SS pooled + breakdown per supermodel + risk_pooling_saving.

**Tech Stack:** Python 3.13, numpy, pandas, openpyxl, pytest 9.

**Spec di riferimento:** `docs/specs/2026-06-22-multi-supermodel-ss-design.md`

> **VINCOLO (CLAUDE.md di `SolMidTerm - Prod`):** non modificare la logica del package `mid_term/` originale. Tutto il lavoro avviene nella copia `mid_term_supermodels/`. Le modifiche a file copiati che cambiano comportamento richiedono diff PRIMA/DOPO e conferma esplicita prima della scrittura.

> **Path:** tutti i percorsi `Create/Modify/Test` sono relativi a `SolMidTerm - Prod/mid_term_supermodels/` salvo diversa indicazione. I comandi `python`/`pytest` si lanciano da dentro `mid_term_supermodels/`.

---

## File structure (nuovi / modificati nella copia)

| File | Responsabilità | Stato |
|------|----------------|-------|
| `matching.py` | + `_build_matching_context()`, `_compute_run_demands_for_sku()` (helper condivisi); + `match_skus_preserve_run_vectors()`; + port Description da BOM; + `pool_safety_stock()`; + `save_results_pooled()` | Modificato |
| `multi_supermodel.py` | Orchestratore: discovery supermodel, esecuzione per-supermodel, raccolta vettori, chiamata pooling | **Nuovo** |
| `results.py` | + dataclass `MultiSupermodelResult` | Modificato |
| `run_pipeline.py` | Selettore modalità via `run_config.json` (`"multi_supermodel": true`) | Modificato |
| `tests/` | Test pytest dei nuovi componenti | **Nuovo** |

---

## Task 0: Setup cartella e branch

**Files:**
- Create: cartella `mid_term_supermodels/` (copia di `mid_term/`)
- Create: `tests/__init__.py`, `tests/conftest.py`

- [ ] **Step 1: Creare il branch git**

Run (dalla root della repo che contiene `SolMidTerm - Prod`):
```bash
git checkout -b feat/multi-supermodel-ss
```
Expected: `Switched to a new branch 'feat/multi-supermodel-ss'`

- [ ] **Step 2: Copiare il package (escludendo cache e output)**

Run (da `SolMidTerm - Prod/`):
```bash
cp -r mid_term mid_term_supermodels
rm -rf mid_term_supermodels/__pycache__ mid_term_supermodels/.spyproject
rm -rf mid_term_supermodels/Output/*
```
Expected: nessun errore; `ls mid_term_supermodels` mostra `matching.py montecarlo.py pipeline.py ...`

- [ ] **Step 3: Smoke test import del package copiato**

Run (da `mid_term_supermodels/`):
```bash
python -c "import matching, montecarlo, pipeline, tr, bom, config, results; print('import OK')"
```
Expected: `import OK`

- [ ] **Step 4: Creare l'infrastruttura test**

Create `tests/__init__.py` (file vuoto).

Create `tests/conftest.py`:
```python
import sys
from pathlib import Path

# Permette ai test di importare i moduli del package (matching, montecarlo, ...)
PKG = Path(__file__).resolve().parent.parent
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))
```

- [ ] **Step 5: Verificare che pytest scopra la suite (ancora vuota)**

Run (da `mid_term_supermodels/`):
```bash
python -m pytest tests/ -q
```
Expected: `no tests ran` (exit 5) — l'infrastruttura è a posto.

- [ ] **Step 6: Commit**

```bash
git add mid_term_supermodels
git commit -m "chore: copia mid_term in mid_term_supermodels + infra test"
```

---

## Task 1: Port correzione Description (BOM → fallback 'Non disponibile')

Porta nel package la correzione già applicata nel monolite `sku_matching_SS.py`: la descrizione viene dal `Testo breve oggetto` della BOM (già caricata come `df_bom`), con fallback `Non disponibile`.

**Files:**
- Modify: `matching.py:899-904` (step 6 dentro `load_lead_times`)
- Test: `tests/test_description_bom.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_description_bom.py`:
```python
import pandas as pd
import numpy as np


def _desc_from_bom(df_bom: pd.DataFrame, skus: list) -> pd.Series:
    """Replica isolata della nuova logica Description (per testarla senza l'intera pipeline)."""
    bom_desc_map = (df_bom[['Numero componenti', 'Testo breve oggetto']]
                    .dropna(subset=['Testo breve oggetto'])
                    .assign(_k=df_bom['Numero componenti'].astype(str))
                    .drop_duplicates(subset='_k', keep='first')
                    .set_index('_k')['Testo breve oggetto'])
    return pd.Series(skus, index=skus).astype(str).map(bom_desc_map).fillna('Non disponibile')


def test_description_taken_from_bom_when_present():
    df_bom = pd.DataFrame({
        'Numero componenti': ['586P3481AB', '586P3481AB', '99999XX'],
        'Testo breve oggetto': ['STGR SERB CARBURANTE GIALLO V24A', None, 'VITE M4'],
    })
    out = _desc_from_bom(df_bom, ['586P3481AB', '99999XX'])
    assert out['586P3481AB'] == 'STGR SERB CARBURANTE GIALLO V24A'
    assert out['99999XX'] == 'VITE M4'


def test_description_fallback_when_missing_in_bom():
    df_bom = pd.DataFrame({
        'Numero componenti': ['111AA'],
        'Testo breve oggetto': ['DADO'],
    })
    out = _desc_from_bom(df_bom, ['222BB'])  # SKU assente in BOM
    assert out['222BB'] == 'Non disponibile'
```

- [ ] **Step 2: Eseguire il test (verifica logica attesa)**

Run: `python -m pytest tests/test_description_bom.py -v`
Expected: PASS (il test valida la logica di riferimento che applicheremo).

- [ ] **Step 3: Applicare la modifica in `matching.py`**

In `load_lead_times`, sostituire il blocco "6. Aggiungi Description" (righe ~899-904).

PRIMA:
```python
    # 6. Aggiungi Description da Testo breve (Prezzi), fallback a 'Residual LT (da Gates)'
    if 'Description' in lead_times_residual.columns:
        desc_map = (lead_times_residual[['SKU', 'Description']]
                    .drop_duplicates(subset='SKU', keep='first')
                    .set_index('SKU')['Description'])
        lead_times['Description'] = lead_times['SKU'].map(desc_map).fillna('Residual LT (da Gates)')
```

DOPO:
```python
    # 6. Description: fonte primaria = BOM ('Testo breve oggetto'); fallback 'Non disponibile'.
    #    df_bom è già caricato sopra (punto 2). La condivisione tra supermodel è implicita
    #    sul codice 'Numero componenti'.
    if 'Testo breve oggetto' in df_bom.columns:
        bom_desc_map = (df_bom[['Numero componenti', 'Testo breve oggetto']]
                        .dropna(subset=['Testo breve oggetto'])
                        .assign(_k=df_bom['Numero componenti'].astype(str))
                        .drop_duplicates(subset='_k', keep='first')
                        .set_index('_k')['Testo breve oggetto'])
        lead_times['Description'] = (lead_times['SKU'].astype(str)
                                     .map(bom_desc_map)
                                     .fillna('Non disponibile'))
    else:
        lead_times['Description'] = 'Non disponibile'
```

- [ ] **Step 4: Verificare che la suite resti verde**

Run: `python -m pytest tests/ -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add matching.py tests/test_description_bom.py
git commit -m "fix: Description da BOM (Testo breve oggetto), fallback 'Non disponibile'"
```

---

## Task 2: Estrarre l'helper di matching condiviso (refactor, decisione D12)

Estrae il setup di matching e il calcolo del vettore per-SKU in helper riusabili, senza cambiare l'output di `match_skus_ultra_optimized`. Sarà la base condivisa con la variante del Task 3.

**Files:**
- Modify: `matching.py` (`match_skus_ultra_optimized` + nuovi `MatchingContext`, `_build_matching_context`, `_compute_run_demands_for_sku`)
- Test: `tests/test_matching_helpers.py`, `tests/_fixtures.py`

- [ ] **Step 1: Creare il fixture sintetico condiviso**

Create `tests/_fixtures.py`:
```python
import numpy as np
import pandas as pd


def tiny_simulation_output(n_runs=4):
    """2 configurazioni, 1 mese, n_runs run. Colonne: MODEL + <Mese>_Run_<i>."""
    rows = []
    # config A: MODEL=MSV4  -> domanda costante 10 per run
    rowA = {'MODEL': 'MSV4'}
    # config B: MODEL=MSV4S -> domanda variabile
    rowB = {'MODEL': 'MSV4S'}
    demA = [10, 10, 10, 10][:n_runs]
    demB = [0, 5, 10, 20][:n_runs]
    for i in range(n_runs):
        rowA[f'JAN 2026_Run_{i}'] = demA[i]
        rowB[f'JAN 2026_Run_{i}'] = demB[i]
    return pd.DataFrame([rowA, rowB])


def tiny_sku_catalog():
    """SKU SKU1 usato da MSV4, SKU2 usato da MSV4S."""
    return pd.DataFrame({
        'Numero componenti': ['SKU1', 'SKU2'],
        'MODEL': ['MSV4', 'MSV4S'],
    })


def tiny_lead_times():
    return pd.DataFrame({
        'SKU': ['SKU1', 'SKU2'],
        'Lead_Time_Months': [1.0, 1.0],
        'Lead_Time_Months_Ceil': [1.0, 1.0],
        'Quantity_Per_Bike': [1.0, 2.0],
        'Description': ['Comp 1', 'Comp 2'],
    })
```

- [ ] **Step 2: Scrivere il test di parità (deve fallire: helper non esistono)**

Create `tests/test_matching_helpers.py`:
```python
import numpy as np
from tests._fixtures import tiny_simulation_output, tiny_sku_catalog, tiny_lead_times
import matching


def test_build_context_and_run_demands():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    ctx = matching._build_matching_context(cat, lt, sim, mapping, n_runs=4)
    try:
        # SKU1 (MODEL=MSV4) deve aggregare la domanda della config A: 10 per ogni run
        sku_data = next(d for sid, d in zip(ctx.sku_ids, ctx.sku_data_list) if sid == 'SKU1')
        run_demands, idxs = matching._compute_run_demands_for_sku(sku_data, ctx)
        assert run_demands.tolist() == [10.0, 10.0, 10.0, 10.0]
    finally:
        ctx.cleanup()


def test_match_skus_ultra_optimized_still_works():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    df, df_no_lt = matching.match_skus_ultra_optimized(cat, lt, sim, mapping, n_runs=4, percentile=99)
    skus = set(df['SKU'])
    assert {'SKU1', 'SKU2'} <= skus
    # SKU1 domanda costante 10 -> safety_stock_p99 == 0 (nessuna variabilità)
    row1 = df[df['SKU'] == 'SKU1'].iloc[0]
    assert abs(row1['safety_stock_p99']) < 1e-6
```

- [ ] **Step 3: Eseguire il test (verifica fallimento)**

Run: `python -m pytest tests/test_matching_helpers.py -v`
Expected: FAIL — `module 'matching' has no attribute '_build_matching_context'`.

- [ ] **Step 4: Aggiungere `MatchingContext` + helper in `matching.py`**

Inserire PRIMA di `match_skus_ultra_optimized` (riga ~1457):
```python
from dataclasses import dataclass, field


@dataclass
class MatchingContext:
    """Stato condiviso del matching: matrice dati, indici mesi, dati SKU pre-estratti.

    Prodotto da _build_matching_context() e consumato da _compute_run_demands_for_sku().
    Possiede l'eventuale memmap su disco: chiamare cleanup() a fine uso.
    """
    sku_ids: list
    sku_data_list: list
    data_matrix: object
    month_col_indices: dict
    month_prefixes: list
    norm_arrays: dict
    category_maps: dict
    column_mapping: dict
    n_rows: int
    n_runs: int
    df_no_lt: "pd.DataFrame"
    _memmap_path: object = None

    def cleanup(self):
        if self._memmap_path is not None:
            try:
                del self.data_matrix
                os.remove(self._memmap_path)
                log.info("File memmap temporaneo rimosso: %s", self._memmap_path)
            except Exception:
                pass
            self._memmap_path = None


def _build_matching_context(sku_catalog, lead_times, simulation_output,
                            column_mapping, n_runs):
    """Costruisce lo stato di matching (STEP 1-3 di match_skus_ultra_optimized).

    Sposta qui, INVARIATO, il corpo delle righe 1479-1640 dell'originale:
    pre-normalizzazione, category_maps/norm_arrays, costruzione data_matrix
    (con fallback memmap), month_col_indices, df_no_lt, sku_ids/sku_data_list
    (con espansione multi-modello). Ritorna un MatchingContext.
    """
    # --- INIZIO blocco spostato da match_skus_ultra_optimized (righe 1479-1640) ---
    normalized_df, norm_arrays_raw = prenormalize_simulation_output_fast(
        simulation_output, column_mapping
    )
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

    _est_gb = n_rows * len(run_columns) * 2 / (1024 ** 3)
    _MEMMAP_THRESHOLD_GB = 4.0
    _memmap_path = None
    if _est_gb > _MEMMAP_THRESHOLD_GB:
        _memmap_path = os.path.join(tempfile.gettempdir(), '_sku_data_matrix.dat')
        data_matrix = np.memmap(_memmap_path, dtype='int16', mode='w+',
                                shape=(n_rows, len(run_columns)))
        _CHUNK = 5000
        for _s in range(0, n_rows, _CHUNK):
            _e = min(_s + _CHUNK, n_rows)
            data_matrix[_s:_e] = (simulation_output.iloc[_s:_e][run_columns]
                                  .to_numpy(dtype=np.int16))
        data_matrix.flush()
    else:
        data_matrix = simulation_output[run_columns].values.astype(np.int16, copy=False)

    run_col_to_idx = {col: idx for idx, col in enumerate(run_columns)}
    month_col_indices = {}
    for month_idx, prefix in enumerate(month_prefixes):
        indices = []
        for i in range(n_runs):
            col_name = f"{prefix}_Run_{i}"
            if col_name in run_col_to_idx:
                indices.append(run_col_to_idx[col_name])
        month_col_indices[month_idx] = indices

    _sku_merged_all = sku_catalog.merge(
        lead_times[['SKU', 'Lead_Time_Months', 'Lead_Time_Months_Ceil',
                    'Quantity_Per_Bike', 'Description']],
        left_on='Numero componenti', right_on='SKU', how='left'
    )
    _mask_no_lt = (
        _sku_merged_all['Lead_Time_Months_Ceil'].isna() |
        (_sku_merged_all['Lead_Time_Months_Ceil'] <= 0)
    )
    df_no_lt = (
        _sku_merged_all[_mask_no_lt][['Numero componenti', 'Lead_Time_Months', 'Description']]
        .rename(columns={'Numero componenti': 'SKU'})
        .drop_duplicates(subset=['SKU']).reset_index(drop=True)
    )
    df_no_lt['Motivo'] = df_no_lt['Lead_Time_Months'].apply(
        lambda x: 'Lead time = 0' if (pd.notna(x) and float(x) == 0)
        else 'Non trovato nei dati residual'
    )

    sku_with_lt = _sku_merged_all[~_mask_no_lt].copy()
    sku_with_lt['Quantity_Per_Bike'] = sku_with_lt['Quantity_Per_Bike'].fillna(1.0)
    sku_groups = sku_with_lt.groupby('SKU')
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
    # --- FINE blocco spostato ---

    return MatchingContext(
        sku_ids=sku_ids, sku_data_list=sku_data_list, data_matrix=data_matrix,
        month_col_indices=month_col_indices, month_prefixes=month_prefixes,
        norm_arrays=norm_arrays, category_maps=category_maps,
        column_mapping=column_mapping, n_rows=n_rows, n_runs=n_runs,
        df_no_lt=df_no_lt, _memmap_path=_memmap_path,
    )


def _compute_run_demands_for_sku(sku_data, ctx):
    """Calcola (run_demands[n_runs], matching_indices) per uno SKU.

    Sposta qui, INVARIATO, il core del loop (righe 1656-1670 dell'originale).
    run_demands è in BICI (non moltiplicato per qty).
    """
    combined_mask = np.zeros(ctx.n_rows, dtype=bool)
    for row_values in sku_data['values']:
        row_mask = sku_matches_combination_numpy(
            row_values, ctx.norm_arrays, ctx.category_maps, ctx.column_mapping, ctx.n_rows
        )
        combined_mask |= row_mask
    matching_indices = np.where(combined_mask)[0]
    if len(matching_indices) == 0:
        return np.zeros(ctx.n_runs, dtype=np.float32), matching_indices
    run_demands = aggregate_monthly_demands_numpy(
        matching_indices, ctx.data_matrix, ctx.month_col_indices,
        sku_data['lead_time'], ctx.n_runs
    )
    return run_demands, matching_indices
```

- [ ] **Step 5: Riscrivere `match_skus_ultra_optimized` usando gli helper**

Sostituire il corpo (dopo i due log iniziali) con:
```python
    log.info("MATCHING SKU V2 (ULTRA-OPTIMIZED)")
    log.info("Memoria iniziale: %.1f MB", get_memory_usage_mb())

    ctx = _build_matching_context(sku_catalog, lead_times, simulation_output,
                                  column_mapping, n_runs)
    log.info("SKU da processare: %d", len(ctx.sku_ids))

    start_time = time.time()
    all_results = []
    for sku_idx, (sku_id, sku_data) in enumerate(zip(ctx.sku_ids, ctx.sku_data_list)):
        run_demands, matching_indices = _compute_run_demands_for_sku(sku_data, ctx)
        if len(matching_indices) == 0:
            continue
        stats = calculate_statistics(run_demands, percentile=percentile)

        monthly_stats = {}
        for month_idx, prefix in enumerate(ctx.month_prefixes):
            col_indices_month = ctx.month_col_indices.get(month_idx, [])
            if not col_indices_month:
                continue
            month_data = ctx.data_matrix[np.ix_(matching_indices, col_indices_month)]
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

    log.info("SKU con match: %d", len(all_results))
    df_no_lt = ctx.df_no_lt
    ctx.cleanup()
    return pd.DataFrame(all_results), df_no_lt
```

- [ ] **Step 6: Eseguire i test**

Run: `python -m pytest tests/test_matching_helpers.py -v`
Expected: PASS (entrambi i test).

- [ ] **Step 7: Verificare l'intera suite**

Run: `python -m pytest tests/ -q`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add matching.py tests/test_matching_helpers.py tests/_fixtures.py
git commit -m "refactor: estrai helper condivisi di matching (MatchingContext)"
```

---

## Task 3: `match_skus_preserve_run_vectors`

Variante che, invece delle statistiche, restituisce i vettori per-run grezzi (in bici) + qty + lead time, per il pooling.

**Files:**
- Modify: `matching.py` (nuova funzione dopo `match_skus_ultra_optimized`)
- Test: `tests/test_preserve_vectors.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_preserve_vectors.py`:
```python
from tests._fixtures import tiny_simulation_output, tiny_sku_catalog, tiny_lead_times
import matching


def test_preserve_run_vectors_returns_bike_demand():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    vectors, df_no_lt = matching.match_skus_preserve_run_vectors(
        cat, lt, sim, mapping, n_runs=4, percentile=99
    )
    # SKU1 -> vettore bici [10,10,10,10], qty 1.0, lead_time 1.0
    rd1, qty1, ltm1 = vectors['SKU1']
    assert rd1.tolist() == [10.0, 10.0, 10.0, 10.0]
    assert qty1 == 1.0
    assert ltm1 == 1.0
    # SKU2 -> vettore bici [0,5,10,20], qty 2.0 (NON ancora moltiplicato)
    rd2, qty2, ltm2 = vectors['SKU2']
    assert rd2.tolist() == [0.0, 5.0, 10.0, 20.0]
    assert qty2 == 2.0
```

- [ ] **Step 2: Eseguire (verifica fallimento)**

Run: `python -m pytest tests/test_preserve_vectors.py -v`
Expected: FAIL — `has no attribute 'match_skus_preserve_run_vectors'`.

- [ ] **Step 3: Implementare la funzione in `matching.py`**

```python
def match_skus_preserve_run_vectors(sku_catalog, lead_times, simulation_output,
                                    column_mapping, n_runs, percentile=99):
    """Come match_skus_ultra_optimized, ma PRESERVA i vettori per-run.

    Returns:
        ( { sku : (run_demands_bici[n_runs] float32, qty_per_bike float, lead_time float) },
          df_no_lt )
    run_demands è in BICI: la moltiplicazione per qty avviene nel pooling.
    """
    log.info("MATCHING SKU (PRESERVE RUN VECTORS)")
    ctx = _build_matching_context(sku_catalog, lead_times, simulation_output,
                                  column_mapping, n_runs)
    vectors = {}
    for sku_id, sku_data in zip(ctx.sku_ids, ctx.sku_data_list):
        run_demands, matching_indices = _compute_run_demands_for_sku(sku_data, ctx)
        if len(matching_indices) == 0:
            continue
        # Copia esplicita: data_matrix/memmap viene liberata da cleanup()
        vectors[sku_id] = (np.array(run_demands, dtype=np.float32, copy=True),
                           float(sku_data['quantity']),
                           float(sku_data['lead_time']))
    df_no_lt = ctx.df_no_lt
    ctx.cleanup()
    log.info("Vettori per-run raccolti: %d SKU", len(vectors))
    return vectors, df_no_lt
```

- [ ] **Step 4: Eseguire il test**

Run: `python -m pytest tests/test_preserve_vectors.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add matching.py tests/test_preserve_vectors.py
git commit -m "feat: match_skus_preserve_run_vectors (vettori per-run per pooling)"
```

---

## Task 4: `pool_safety_stock` (cuore statistico)

Pooling somma-prodotto per-run a livello componente; SS pooled + breakdown per supermodel + risk_pooling_saving.

**Files:**
- Modify: `matching.py` (nuova funzione)
- Test: `tests/test_pooling.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_pooling.py`:
```python
import numpy as np
import matching


def test_pooling_sumproduct_and_risk_pooling():
    # 2 supermodel, n_runs=4. Componente 'X' condiviso.
    # SuperA: bici [10,10,10,10] qty 1 -> pezzi [10,10,10,10] (varianza 0)
    # SuperB: bici [0,5,10,20]   qty 2 -> pezzi [0,10,20,40]
    sm_results = {
        'SuperA': {'X': (np.array([10, 10, 10, 10], dtype=np.float32), 1.0, 1.0)},
        'SuperB': {'X': (np.array([0, 5, 10, 20], dtype=np.float32), 2.0, 1.0)},
    }
    df_pool, breakdown = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]

    # pezzi pooled per run = [10, 20, 30, 50]
    pooled = np.array([10, 20, 30, 50], dtype=np.float64)
    assert abs(row['mean_demand_pooled'] - pooled.mean()) < 1e-6
    expected_ss = max(0.0, np.percentile(pooled, 99) - pooled.mean())
    assert abs(row['safety_stock_p99_pooled'] - expected_ss) < 1e-6

    # breakdown per supermodel presente
    assert 'safety_stock_standalone_SuperA' in row
    assert 'safety_stock_standalone_SuperB' in row
    # SuperA pezzi costanti -> SS standalone 0
    assert abs(row['safety_stock_standalone_SuperA']) < 1e-6

    # risk pooling: saving >= 0
    assert row['risk_pooling_saving'] >= -1e-9


def test_pooling_component_in_single_supermodel():
    # Componente 'Y' solo in SuperA
    sm_results = {
        'SuperA': {'Y': (np.array([1, 2, 3, 4], dtype=np.float32), 1.0, 1.0)},
        'SuperB': {},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'Y'].iloc[0]
    pooled = np.array([1, 2, 3, 4], dtype=np.float64)
    assert abs(row['mean_demand_pooled'] - pooled.mean()) < 1e-6
    # niente contributo da SuperB
    assert row['mean_demand_SuperB'] == 0.0
```

- [ ] **Step 2: Eseguire (verifica fallimento)**

Run: `python -m pytest tests/test_pooling.py -v`
Expected: FAIL — `has no attribute 'pool_safety_stock'`.

- [ ] **Step 3: Implementare in `matching.py`**

```python
def pool_safety_stock(supermodel_results, percentile=99, n_runs=None):
    """Pool per-run delle domande dei componenti tra supermodel (somma-prodotto).

    Args:
        supermodel_results: { supermodel : { sku : (run_demands_bici[n_runs], qty, lead_time) } }
        percentile: percentile per la safety stock (default 99).
        n_runs: lunghezza attesa dei vettori (se None, dedotta dal primo vettore).

    Returns:
        (df_per_sku_pooled, breakdown_dict)
        df_per_sku_pooled: una riga per componente con SS pooled + colonne breakdown.
        breakdown_dict: { sku : { supermodel : {mean, ss_standalone, qty, lead_time} } }
    """
    ss_key = f'safety_stock_p{percentile}'
    supermodels = list(supermodel_results.keys())

    # Deduci n_runs
    if n_runs is None:
        for sm in supermodels:
            for _sku, (rd, _q, _lt) in supermodel_results[sm].items():
                n_runs = len(rd)
                break
            if n_runs:
                break
    if not n_runs:
        return pd.DataFrame(), {}

    # Unione di tutti i codici componente
    all_skus = set()
    for sm in supermodels:
        all_skus.update(supermodel_results[sm].keys())

    rows = []
    breakdown = {}
    for sku in sorted(all_skus):
        pooled = np.zeros(n_runs, dtype=np.float64)
        row = {'SKU': sku}
        breakdown[sku] = {}
        sum_standalone_ss = 0.0
        lead_times_seen = []
        for sm in supermodels:
            entry = supermodel_results[sm].get(sku)
            if entry is None:
                row[f'mean_demand_{sm}'] = 0.0
                row[f'{ss_key}_standalone_{sm}'] = 0.0
                continue
            run_demands, qty, lead_time = entry
            pezzi_sm = np.asarray(run_demands, dtype=np.float64) * qty
            pooled += pezzi_sm
            lead_times_seen.append(lead_time)

            mean_sm = float(pezzi_sm.mean())
            p_sm = float(np.percentile(pezzi_sm, percentile))
            ss_sm = max(0.0, p_sm - mean_sm)
            sum_standalone_ss += ss_sm

            row[f'mean_demand_{sm}'] = round(mean_sm, 4)
            row[f'{ss_key}_standalone_{sm}'] = round(ss_sm, 4)
            breakdown[sku][sm] = {'mean': mean_sm, 'ss_standalone': ss_sm,
                                  'qty': qty, 'lead_time': lead_time}

        mean_pooled = float(pooled.mean())
        p_pooled = float(np.percentile(pooled, percentile))
        ss_pooled = max(0.0, p_pooled - mean_pooled)

        row['mean_demand_pooled'] = round(mean_pooled, 4)
        row[f'{ss_key}_pooled'] = round(ss_pooled, 4)
        row['risk_pooling_saving'] = round(sum_standalone_ss - ss_pooled, 4)
        row['Lead_Time_max'] = round(max(lead_times_seen), 4) if lead_times_seen else 0.0
        rows.append(row)

    return pd.DataFrame(rows), breakdown
```

- [ ] **Step 4: Eseguire i test**

Run: `python -m pytest tests/test_pooling.py -v`
Expected: PASS (entrambi).

- [ ] **Step 5: Commit**

```bash
git add matching.py tests/test_pooling.py
git commit -m "feat: pool_safety_stock (somma-prodotto per-run + risk pooling)"
```

---

## Task 5: Discovery supermodel + config per-supermodel

**Files:**
- Create: `multi_supermodel.py`
- Test: `tests/test_discovery.py`

- [ ] **Step 1: Scrivere il test che fallisce**

Create `tests/test_discovery.py`:
```python
import os
from pathlib import Path
import multi_supermodel


def test_discover_supermodels(tmp_path):
    # Crea 2 cartelle supermodel valide + 1 cartella senza i marker
    for sm in ['SuperA', 'SuperB']:
        d = tmp_path / sm
        (d / 'MODEL').mkdir(parents=True)
        (d / 'Total_demand.xlsx').write_text('x')
        (d / 'TR FOO.xlsx').write_text('x')
    (tmp_path / 'Dati Logistica').mkdir()  # NON deve essere considerata supermodel

    found = multi_supermodel.discover_supermodels(str(tmp_path))
    names = sorted(Path(p).name for p in found)
    assert names == ['SuperA', 'SuperB']
```

- [ ] **Step 2: Eseguire (verifica fallimento)**

Run: `python -m pytest tests/test_discovery.py -v`
Expected: FAIL — `No module named 'multi_supermodel'`.

- [ ] **Step 3: Creare `multi_supermodel.py` con la discovery**

```python
# -*- coding: utf-8 -*-
"""Orchestratore Safety Stock multi-supermodel.

Esegue la pipeline SS per ogni cartella supermodel (domanda + TR propri,
seed diverso) e fa il pooling dei vettori per-run a livello componente.
"""
import os
from pathlib import Path
from typing import List

from _logging import get_logger

log = get_logger(__name__)


def discover_supermodels(input_dir: str) -> List[str]:
    """Trova le cartelle supermodel sotto input_dir.

    Una cartella è un supermodel se contiene: sottocartella MODEL/,
    un Total_demand.xlsx e almeno un file il cui nome inizia con 'TR'.
    """
    base = Path(input_dir)
    result = []
    for sub in sorted(base.iterdir()):
        if not sub.is_dir():
            continue
        has_model = (sub / 'MODEL').is_dir()
        has_demand = (sub / 'Total_demand.xlsx').exists()
        has_tr = any(f.name.upper().startswith('TR') and f.suffix.lower() in ('.xlsx', '.xls')
                     for f in sub.iterdir() if f.is_file())
        if has_model and has_demand and has_tr:
            result.append(str(sub))
    log.info("Supermodel trovati: %d", len(result))
    return result
```

> Nota: verificare il nome reale dell'helper di logging in `_logging.py` (es. `get_logger`). Se diverso, adeguare l'import. Controllare con: `grep -n "def " _logging.py`.

- [ ] **Step 4: Eseguire il test**

Run: `python -m pytest tests/test_discovery.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add multi_supermodel.py tests/test_discovery.py
git commit -m "feat: discovery cartelle supermodel"
```

---

## Task 6: Orchestratore `run_multi_supermodel`

Esegue la pipeline per ogni supermodel (con il suo input_dir e seed) raccogliendo i vettori per-run, poi chiama il pooling. Gestisce il vincolo dei path hardcoded via `os.chdir` per-supermodel.

**Files:**
- Modify: `multi_supermodel.py`
- Modify: `pipeline.py` (esporre un hook che ritorna i vettori; vedi Step 1)
- Test: `tests/test_orchestrator_smoke.py`

- [ ] **Step 1: Verificare come `pipeline.py` espone gli step**

Run: `grep -n "def run_step_2_montecarlo\|def run_step_3_matching\|results_by_month\|build_wide_dataframe\|match_skus_ultra_optimized" pipeline.py`
Expected: localizza le righe di `run_step_2_montecarlo` (~343) e `run_step_3_matching` (~425).

Obiettivo dello step: confermare che in `run_step_3_matching` la chiamata a `match_skus_ultra_optimized` riceve `simulation_output` (wide DF) e `column_mapping`. La variante userà gli STESSI input ma chiamerà `match_skus_preserve_run_vectors`.

- [ ] **Step 2: Aggiungere a `pipeline.py` un metodo che ritorna i vettori per-run**

Dopo `run_step_3_matching`, aggiungere (riusa step 0,1,2 esistenti, sostituisce solo il matching finale):
```python
def run_collect_run_vectors(self):
    """Esegue step 0-2 standard e poi il matching che PRESERVA i vettori per-run.

    Returns:
        ( { sku : (run_demands_bici[n_runs], qty, lead_time) }, df_no_lt )
    """
    import matching
    tr = self.run_step_0_load_tr()
    bom = self.run_step_1_build_bom()
    sim = self.run_step_2_montecarlo(tr)

    # Replica la preparazione input di run_step_3_matching FINO alla chiamata di matching.
    # NOTA per l'implementatore: copiare da run_step_3_matching le righe che costruiscono
    # sku_catalog, lead_times, column_mapping, n_runs, percentile (pipeline.py:425-518),
    # SENZA la chiamata finale match_skus_ultra_optimized. Poi:
    vectors, df_no_lt = matching.match_skus_preserve_run_vectors(
        sku_catalog, lead_times, sim.df_wide, column_mapping,
        n_runs=sim.n_runs, percentile=self.sim_config.percentile_safety_stock
    )
    return vectors, df_no_lt
```

> L'implementatore deve aprire `pipeline.py:425-518` e replicare la preparazione di `sku_catalog`, `lead_times`, `column_mapping` esattamente come in `run_step_3_matching`, fermandosi prima della chiamata `match_skus_ultra_optimized`. Questo blocco NON va inventato: va copiato dal metodo esistente.

- [ ] **Step 3: Implementare l'orchestratore in `multi_supermodel.py`**

```python
def run_multi_supermodel(input_dir: str, output_dir: str, json_path: str,
                         seed_base: int = 42) -> "MultiSupermodelResult":
    """Esegue la pipeline per ogni supermodel e fa il pooling."""
    import matching
    from config import PipelineConfig, load_simulation_config
    from montecarlo import load_monthly_forecast
    from pipeline import SafetyStockPipeline
    from results import MultiSupermodelResult

    supermodel_dirs = discover_supermodels(input_dir)
    if not supermodel_dirs:
        raise RuntimeError(f"Nessun supermodel trovato in {input_dir}")

    sm_results = {}
    cwd0 = os.getcwd()
    for k, sm_dir in enumerate(supermodel_dirs):
        sm_name = Path(sm_dir).name
        log.info("=== Supermodel %d/%d: %s (seed=%d) ===",
                 k + 1, len(supermodel_dirs), sm_name, seed_base + k)
        # Il package usa path hardcoded .\Input\...: entra nella cartella supermodel.
        # Si assume che dentro sm_dir ci sia un layout compatibile (Input/ del supermodel).
        config = PipelineConfig(input_dir=Path(sm_dir), output_dir=Path(output_dir) / sm_name)
        forecast_path = config.resolve_input(config.forecast_file)
        _, month_names = load_monthly_forecast(str(forecast_path))
        sim_config = load_simulation_config(json_path, month_names)
        sim_config.random_seed = seed_base + k  # indipendenza statistica (D8)

        pipeline = SafetyStockPipeline(config, sim_config, validate_config=False)
        vectors, _df_no_lt = pipeline.run_collect_run_vectors()
        sm_results[sm_name] = vectors
    os.chdir(cwd0)

    df_pool, breakdown = matching.pool_safety_stock(
        sm_results, percentile=load_simulation_config(json_path, ['x']).percentile_safety_stock
    )
    return MultiSupermodelResult(
        per_sku_pooled=df_pool, per_supermodel_breakdown=breakdown,
        output_paths={}, elapsed_seconds=0.0
    )
```

> **Punto di attenzione (path):** il package risolve `.\Input\...` rispetto alla CWD impostata dal launcher. L'implementatore deve verificare in `config.py:resolve_input` e in `montecarlo.py`/`matching.py` se i path sono relativi alla CWD o a `config.input_dir`. Se sono CWD-based, l'orchestratore deve fare `os.chdir(sm_dir)` (con relativo ripristino) invece di passare `input_dir`. Decidere in base a `grep -n "resolve_input\|\.\\\\Input\|Path('.')" config.py matching.py montecarlo.py`.

- [ ] **Step 4: Smoke test dell'orchestratore (solo discovery + errore su input vuoto)**

Create `tests/test_orchestrator_smoke.py`:
```python
import pytest
import multi_supermodel


def test_run_multi_supermodel_raises_on_empty(tmp_path):
    with pytest.raises(RuntimeError, match="Nessun supermodel"):
        multi_supermodel.run_multi_supermodel(
            str(tmp_path), str(tmp_path / 'out'), 'run_config.json'
        )
```

- [ ] **Step 5: Eseguire il test**

Run: `python -m pytest tests/test_orchestrator_smoke.py -v`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add multi_supermodel.py pipeline.py tests/test_orchestrator_smoke.py
git commit -m "feat: orchestratore run_multi_supermodel + hook vettori per-run"
```

---

## Task 7: `MultiSupermodelResult` + output Excel

**Files:**
- Modify: `results.py`
- Modify: `matching.py` (nuova `save_results_pooled`)
- Test: `tests/test_save_pooled.py`

- [ ] **Step 1: Aggiungere la dataclass in `results.py`**

```python
@dataclass
class MultiSupermodelResult:
    per_sku_pooled: "pd.DataFrame"
    per_supermodel_breakdown: dict
    output_paths: dict
    elapsed_seconds: float
```
> Verificare gli import esistenti in `results.py` (pandas, dataclass). Allinearsi allo stile del file.

- [ ] **Step 2: Scrivere il test che fallisce**

Create `tests/test_save_pooled.py`:
```python
import numpy as np
import pandas as pd
import openpyxl
import matching


def test_save_results_pooled_writes_columns(tmp_path):
    sm_results = {
        'SuperA': {'X': (np.array([10, 10, 10, 10], dtype=np.float32), 1.0, 1.0)},
        'SuperB': {'X': (np.array([0, 5, 10, 20], dtype=np.float32), 2.0, 1.0)},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    out = tmp_path / 'pooled.xlsx'
    matching.save_results_pooled(df_pool, str(out), percentile=99)

    wb = openpyxl.load_workbook(out)
    hdr = [c.value for c in wb.active[1]]
    assert 'SKU' in hdr
    assert 'safety_stock_p99_pooled' in hdr
    assert 'risk_pooling_saving' in hdr
    assert 'mean_demand_SuperA' in hdr
```

- [ ] **Step 3: Eseguire (verifica fallimento)**

Run: `python -m pytest tests/test_save_pooled.py -v`
Expected: FAIL — `has no attribute 'save_results_pooled'`.

- [ ] **Step 4: Implementare `save_results_pooled` in `matching.py`**

```python
def save_results_pooled(df_pool, filename, percentile=99):
    """Scrive l'output pooled multi-supermodel su Excel.

    Ordina per safety_stock pooled decrescente (componenti più critici in cima).
    """
    ss_pooled = f'safety_stock_p{percentile}_pooled'
    df_out = df_pool.copy()
    if ss_pooled in df_out.columns:
        df_out = df_out.sort_values(ss_pooled, ascending=False)
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name='SS_pooled')
    log.info("Output pooled salvato: %s (%d componenti)", filename, len(df_out))
    return df_out
```

- [ ] **Step 5: Eseguire il test**

Run: `python -m pytest tests/test_save_pooled.py -v`
Expected: PASS.

- [ ] **Step 6: Collegare il salvataggio nell'orchestratore**

In `multi_supermodel.run_multi_supermodel`, prima del `return`, aggiungere:
```python
    import matching
    out_path = os.path.join(output_dir, 'sku_safety_stock_POOLED_supermodels.xlsx')
    os.makedirs(output_dir, exist_ok=True)
    matching.save_results_pooled(df_pool, out_path,
                                 percentile=load_simulation_config(json_path, ['x']).percentile_safety_stock)
```
e impostare `output_paths={'pooled': out_path}` nel risultato.

- [ ] **Step 7: Verificare la suite**

Run: `python -m pytest tests/ -q`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add results.py matching.py multi_supermodel.py tests/test_save_pooled.py
git commit -m "feat: MultiSupermodelResult + save_results_pooled (output Excel)"
```

---

## Task 8: Entry point — flag modalità in `run_pipeline.py`

**Files:**
- Modify: `run_pipeline.py`
- Modify: `run_config.json` (aggiungere chiave)
- Test: manuale (smoke)

- [ ] **Step 1: Aggiungere la chiave in `run_config.json`**

Aggiungere `"multi_supermodel": false` al JSON (default = comportamento attuale).

- [ ] **Step 2: Aggiungere il ramo in `run_pipeline.py:main()`**

Dopo il caricamento config, prima di costruire `SafetyStockPipeline`:
```python
    import json
    with open(JSON_PATH, encoding='utf-8') as _f:
        _cfg_raw = json.load(_f)
    if _cfg_raw.get('multi_supermodel', False):
        from multi_supermodel import run_multi_supermodel
        result = run_multi_supermodel(
            input_dir=str(config.input_dir),
            output_dir=str(config.output_dir),
            json_path=str(JSON_PATH),
            seed_base=_cfg_raw.get('random_seed', 42),
        )
        print(f"[multi-supermodel] pooled output: {result.output_paths}")
        return
    # ... ramo single-supermodel esistente invariato ...
```

- [ ] **Step 3: Smoke test single-supermodel (non-regressione)**

Con `"multi_supermodel": false`, verificare che il launcher parta come prima (anche solo fino al primo step, su dati ridotti se disponibili).
Run: `python run_pipeline.py` (interrompere dopo l'avvio se i dati completi sono troppo pesanti).
Expected: nessun errore di import/branch; parte il ramo classico.

- [ ] **Step 4: Commit**

```bash
git add run_pipeline.py run_config.json
git commit -m "feat: flag multi_supermodel in run_config + ramo entry point"
```

---

## Task 9: Validazione end-to-end (caso degenere 1 supermodel + risk pooling)

**Files:**
- Test: `tests/test_validation.py`

- [ ] **Step 1: Test — caso degenere: 1 supermodel ≈ pipeline classica per componente**

Create `tests/test_validation.py`:
```python
import numpy as np
import matching


def test_single_supermodel_equals_standalone():
    # Con UN solo supermodel, SS pooled == SS standalone di quel supermodel.
    rd = np.array([0, 5, 10, 20], dtype=np.float32)
    sm_results = {'SuperA': {'X': (rd, 1.0, 1.0)}}
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]
    assert abs(row['safety_stock_p99_pooled'] - row['safety_stock_p99_standalone_SuperA']) < 1e-6
    assert abs(row['risk_pooling_saving']) < 1e-6


def test_risk_pooling_reduces_or_equals_sum():
    # SS pooled <= somma SS standalone, per due flussi.
    sm_results = {
        'SuperA': {'X': (np.array([0, 5, 10, 20], dtype=np.float32), 1.0, 1.0)},
        'SuperB': {'X': (np.array([20, 10, 5, 0], dtype=np.float32), 1.0, 1.0)},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]
    sum_standalone = (row['safety_stock_p99_standalone_SuperA']
                      + row['safety_stock_p99_standalone_SuperB'])
    assert row['safety_stock_p99_pooled'] <= sum_standalone + 1e-9
    assert row['risk_pooling_saving'] >= -1e-9
```

- [ ] **Step 2: Eseguire i test**

Run: `python -m pytest tests/test_validation.py -v`
Expected: PASS (entrambi).

- [ ] **Step 3: Eseguire l'intera suite**

Run: `python -m pytest tests/ -q`
Expected: tutti PASS.

- [ ] **Step 4: Commit**

```bash
git add tests/test_validation.py
git commit -m "test: validazione caso degenere + risk pooling"
```

---

## Note finali per l'implementatore

- **Dati reali multi-supermodel**: questo piano è validato con fixture sintetici. Per un test end-to-end con dati reali servono cartelle supermodel popolate sotto `Input/`. Eseguire `python run_pipeline.py` con `"multi_supermodel": true` solo quando i dati sono pronti.
- **Path hardcoded** (Task 6, Step 3): è il rischio tecnico maggiore. Risolvere PRIMA di scrivere l'orchestratore, ispezionando `config.py`/`montecarlo.py`/`matching.py`. Se i path sono CWD-based, usare `os.chdir(sm_dir)` con ripristino in `finally`.
- **Memoria** (rischio §10 spec): i vettori per-run di tutti i componenti × supermodel restano in RAM durante il pooling. Con molti componenti e `n_runs` grande, valutare `float32` (già usato) e, se necessario, processare i componenti a blocchi.
- **CLAUDE.md**: ogni modifica a file copiati che cambia comportamento va mostrata PRIMA/DOPO e confermata prima della scrittura.

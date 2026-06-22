# Safety Stock multi-supermodel con pooling per-run — Design

**Data:** 2026-06-22
**Package:** `SolMidTerm - Prod/mid_term`
**Ambito:** Safety Stock (pipeline Monte Carlo + SKU matching). NON tocca lo Stock Alignment.

---

## 1. Obiettivo

Oggi la pipeline SS calcola la safety stock su **un'unica domanda globale** (`Total_demand.xlsx`, riga unica "linea Multistrada"). Le versioni BOM (V21E, V22E, V24T) vengono consolidate e il pooling per componente avviene già, ma su quella domanda unica.

Si vuole passare a una struttura a **supermodel**: più linee di prodotto, ciascuna con la **propria domanda indipendente** e i **propri take rate**, che però **condividono molti componenti**. Poiché i componenti sono condivisi, la safety stock va calcolata in modo **condiviso (pooled)** sfruttando il risk pooling: la varianza della domanda sommata di flussi indipendenti è inferiore alla somma delle deviazioni standard, quindi

```
SS_pooled  ≤  Σ SS_standalone_per_supermodel
```

Calcolare le SS per supermodel e sommarle sovradimensiona le scorte. Il calcolo condiviso è statisticamente corretto e più economico.

---

## 2. Decisioni di design (concordate)

| # | Tema | Decisione |
|---|------|-----------|
| D1 | Natura della domanda | Flussi **indipendenti**: un forecast per supermodel, non quote di un totale globale. |
| D2 | Gerarchia | **Supermodel SOPRA il model/versione** (Scenario B). Ogni supermodel contiene al suo interno l'attuale meccanica (più versioni consolidate + ModelMix). |
| D3 | Struttura input | Una **cartella per supermodel**, ciascuna con `MODEL/`, file TR, `Total_demand.xlsx` propri. |
| D4 | Forecast | Un `Total_demand.xlsx` **dentro ogni cartella supermodel** (una riga, prima colonna = nome supermodel). NON un file centralizzato. |
| D5 | Dati logistici | **Condivisi**, una sola volta (`Input/Dati Logistica/`: Gates, Dashboard Supplier, Prezzi). Il residual LT è già supermodel/Versione-aware via la colonna `Supermodel` di `Gates MTSV4.xlsx`. |
| D6 | Lead time | **Finestra per supermodel**: ogni supermodel aggrega la sua domanda sulla **propria** finestra LT, poi si sommano i vettori per-run. |
| D7 | Quantità | **Somma-prodotto per-run**: dentro ogni supermodel, `bici_per_run × qty_per_bike_sm`, POI somma tra supermodel. La qty va applicata **prima** del pooling (le qty possono differire per supermodel). |
| D8 | Indipendenza statistica | **Seed diverso per supermodel** (`seed_base + indice_supermodel`), stesso `n_runs`. Necessario perché i flussi sono indipendenti; seed uguale falserebbe la correlazione e gonfierebbe la varianza. |
| D9 | Output | SS **pooled + breakdown per supermodel** + colonna `risk_pooling_saving`. |
| D10 | Codebase | Package `SolMidTerm - Prod/mid_term` (il più recente). `Desktop/mid_term` ignorato. |
| D11 | Description | Da BOM (`Testo breve oggetto`), fallback `Non disponibile` (modifica già applicata in `sku_matching_SS.py` monolite; va portata nel package — vedi §7). |

---

## 3. Struttura input target

```
Input/
  <SuperModelA>/
    MODEL/
      V21E/   (Mappatura_*.xlsx, BOM_150*.xlsx, ...)
      V22E/
    TR ...xlsx
    Total_demand.xlsx        # una riga, prima colonna = "SuperModelA"
  <SuperModelB>/
    MODEL/
      V24T/
      ...
    TR ...xlsx
    Total_demand.xlsx        # una riga, prima colonna = "SuperModelB"
  Dati Logistica/            # CONDIVISO
    Gates MTSV4.xlsx
    20260520 Dashboard Supplier.xlsx
    PREZZI.xlsx
    ...
```

Auto-discovery dei supermodel: ogni sottocartella di `Input/` che contiene `MODEL/` + un file TR + `Total_demand.xlsx`.

---

## 4. Pipeline attuale (riferimento, file:line)

Entry: `run_pipeline.py:36-63` → `SafetyStockPipeline.run()` (`pipeline.py:112-124`):

```
run_step_0_load_tr()        # TR + forecast + esclusioni      pipeline.py:127-179
run_step_1_build_bom()      # discovery versioni + consolida   pipeline.py:182-340
run_step_2_montecarlo(tr)   # MC                               pipeline.py:343-422
run_step_3_matching(...)    # matching + lead time             pipeline.py:425-540
run_step_4_safety_stock()   # SS + prezzi + output Excel       pipeline.py:543-683
```

Punti chiave:
- **Forecast**: `load_monthly_forecast()` `montecarlo.py:491-535`. Riga 0 = mesi, riga 1 = valori.
- **Seed**: `montecarlo.py:652` → `np.random.seed(config.random_seed + start_month_offset*1000 + month_index)`.
- **MC per-run**: `run_monthly_simulations_optimized()` `montecarlo.py:745-824` restituisce `{month_idx: {config_key: np.ndarray[n_runs]}}` — **i vettori per-run esistono già**.
- **Wide DF**: `build_wide_dataframe()` `montecarlo.py:875-978` appiattisce in colonne `<Mese>_Run_<i>`.
- **Matching**: `match_skus_ultra_optimized()` `matching.py:1457-1717`. Per ogni SKU costruisce `run_demands[n_runs]` via `aggregate_monthly_demands_numpy()` (`matching.py:1100-1166`, `:1667-1670`), poi **chiama subito** `calculate_statistics()` (`matching.py:1173-1228`, `:1671`) **scartando il vettore**.
- **Qty applicata DOPO**: `matching.py:1695` → `result['safety_stock_p99_total'] = result['safety_stock_p99'] * sku_data['quantity']`.
- **Residual LT supermodel-aware**: `calcola_residual_lead_time()` `matching.py:104-300+`, usa colonna `Versione`/`Supermodel` + `Gates MTSV4.xlsx`.
- **Output**: `save_results()` `matching.py:1266-1340`, `save_results_monthly()` `matching.py:1342-1440`.

---

## 5. Architettura della soluzione

L'attuale `SafetyStockPipeline.run()` diventa concettualmente **la run di UN supermodel**. Si aggiunge un **orchestratore** sopra che:

1. scopre le cartelle supermodel;
2. esegue la pipeline per ciascuno (con il suo input e il suo seed);
3. raccoglie i **vettori per-run per componente**;
4. fa il **pooling** e produce l'output.

### 5.1 Flusso

```
run_multi_supermodel()
│
├─ discovery supermodel  (sottocartelle Input/ con MODEL/+TR+Total_demand)
│
├─ PER OGNI supermodel  sm  (indice k):
│   ├─ config_sm = config con input_dir = Input/<sm>, random_seed = seed_base + k
│   ├─ step 0: load TR/forecast/esclusioni  (del supermodel)
│   ├─ step 1: build BOM  (versioni interne → catalogo consolidato con Versione)
│   ├─ step 2: MC  → results_by_month_sm = {month: {config: array[n_runs]}}
│   ├─ step 3 (variante): match preservando i vettori
│   │     → sku_vectors_sm = { sku : (run_demands_bici[n_runs], qty_sm, lead_time_sm) }
│   └─ raccogli sku_vectors_sm + lead_times_sm + descrizioni
│
└─ POOLING (per ogni codice componente, unione su tutti i supermodel):
    pezzi_pooled[i] = Σ_sm ( run_demands_bici_sm[i] × qty_sm )      # somma-prodotto per-run
    mean_pooled     = mean(pezzi_pooled)
    p99_pooled      = percentile(pezzi_pooled, 99)
    SS_pooled       = max(0, p99_pooled − mean_pooled)
    # breakdown per supermodel
    per ogni sm:  mean_sm = mean(run_demands_bici_sm × qty_sm)
                  SS_standalone_sm = max(0, p99(run_demands_bici_sm × qty_sm) − mean_sm)
    risk_pooling_saving = Σ_sm SS_standalone_sm − SS_pooled
    → riga output
```

### 5.2 Allineamento dei vettori per-run

Tutti i supermodel usano lo **stesso `n_runs`**. Il pooling somma per **indice di run** (`run i` di A con `run i` di B). Con seed diversi per supermodel, ogni coppia `(A_i, B_i)` è una realizzazione di flussi indipendenti → la distribuzione della somma è corretta.

### 5.3 Finestra lead time per supermodel

Ogni supermodel calcola la propria finestra LT (gate del suo supermodel via `Gates MTSV4.xlsx`) dentro `aggregate_monthly_demands_numpy`, sul **proprio** MC. Il vettore `run_demands_bici_sm` è quindi già aggregato sulla finestra LT corretta di quel supermodel. Il pooling somma vettori già "finestrati".

---

## 6. Componenti software (nuovi / modificati)

### 6.1 NUOVO — variante matching che preserva i vettori
`matching.py`, nuova funzione accanto a `match_skus_ultra_optimized` (dopo `:1717`):

```
match_skus_preserve_run_vectors(sku_catalog, lead_times, simulation_output,
                                column_mapping, n_runs, percentile)
  -> ( { sku : (run_demands_bici[n_runs], qty_per_bike, lead_time) }, df_no_lt )
```

Identica a `match_skus_ultra_optimized` fino a `run_demands` (`:1667-1670`), ma invece di chiamare `calculate_statistics` e scartare, **restituisce il vettore** `run_demands` (in **bici**, NON moltiplicato per qty) insieme a `qty_per_bike` e `lead_time`. La moltiplicazione per qty avviene nel pooling (D7).

Rationale: i vettori per-run esistono già internamente; serve solo non scartarli.

**Decisione (D12): NO duplicazione.** Il corpo comune di matching (normalizzazione, costruzione `data_matrix`, merge lead times, loop per-SKU fino a `run_demands`) viene **estratto in un helper condiviso**, usato sia da `match_skus_ultra_optimized` sia dalla variante `match_skus_preserve_run_vectors`. Le due funzioni differiscono solo nello step finale: una chiama `calculate_statistics` e restituisce il DataFrame stats, l'altra restituisce il vettore grezzo. Evita due copie divergenti della logica.

### 6.2 NUOVO — orchestratore multi-supermodel
`pipeline.py`, nuovo metodo (o nuovo modulo `multi_supermodel.py`):

```
SafetyStockPipeline.run_multi_supermodel(supermodel_dirs=None) -> MultiSupermodelResult
```

- Se `supermodel_dirs` è `None`, auto-discovery.
- Per ogni supermodel istanzia una `PipelineConfig` con `input_dir` sulla cartella del supermodel e `random_seed = seed_base + k`.
- Esegue step 0→3 (variante), raccoglie `sku_vectors_sm`.
- Invoca il pooling e scrive l'output.

### 6.3 NUOVO — funzione di pooling
`matching.py` o `multi_supermodel.py`:

```
pool_safety_stock(supermodel_results, percentile, n_runs) -> (df_per_sku_pooled, df_breakdown)
```

Implementa la formula §5.1. Gestisce componenti presenti in un sottoinsieme di supermodel (somma solo gli addendi presenti).

### 6.4 NUOVO — dataclass risultato
`results.py`:

```
@dataclass
class MultiSupermodelResult:
    per_sku_pooled: pd.DataFrame
    per_supermodel_breakdown: dict
    output_paths: dict
    elapsed_seconds: float
```

### 6.5 MODIFICATO — output
Estendere `save_results` (o nuova `save_results_pooled`) per scrivere le colonne:
`SKU, Description, mean_demand_pooled, safety_stock_p99_pooled, Lead_Time, Quantity_*,`
`mean_demand_<SM>, safety_stock_standalone_<SM>` (per ogni supermodel), `risk_pooling_saving`.

### 6.6 MODIFICATO — entry point
`run_pipeline.py`: aggiungere selezione modalità (flag CLI o chiave in `run_config.json`, es. `"multi_supermodel": true`) per invocare `run_multi_supermodel()`.

---

## 7. Allineamento Description (da §2 D11)

La modifica "Description da BOM (`Testo breve oggetto`), fallback `Non disponibile`" è stata applicata al monolite `sku_matching_SS.py` (`SolMidTerm - ALL`). Nel package `mid_term`, `calcola_residual_lead_time`/`load_lead_times` (`matching.py:~900`) usa ancora `Prezzi` con fallback `Residual LT (da Gates)`. Portare la stessa modifica nel package per coerenza dell'output.

---

## 8. Validazione

1. **Caso degenere (1 supermodel)**: l'output multi-supermodel deve coincidere con la pipeline attuale (a meno del breakdown). Sanity check di non-regressione.
2. **Risk pooling**: verificare `SS_pooled ≤ Σ SS_standalone` e `risk_pooling_saving ≥ 0` su tutti i componenti condivisi.
3. **Somma-prodotto qty**: componente con qty diverse tra supermodel — verificare che il pooled rifletta i pezzi (non le bici).
4. **Indipendenza seed**: confronto varianza pooled con seed uguali vs diversi — i diversi devono dare varianza (e SS) inferiore.
5. **Finestra LT**: componente condiviso con LT diverso tra supermodel — verificare che ogni addendo usi la sua finestra.

---

## 9. Fuori ambito

- Stock Alignment (`app_V3.py`) e il suo concetto di "Supermodel" comma-separated: non toccati.
- Modifiche alla logica Monte Carlo, alla formula del residual LT, al parsing TR.
- Ottimizzazioni di performance oltre a quanto necessario per non scartare i vettori per-run.

---

## 10. Rischi / punti aperti per il piano

- **Memoria**: conservare i vettori per-run di tutti i componenti × supermodel. Stimare `n_componenti × n_runs × n_supermodel × 4 byte`; con `n_runs` grande valutare `float32` e rilascio incrementale.
- **Matching duplicato**: decidere se estrarre il corpo comune tra `match_skus_ultra_optimized` e la variante (preferibile per non divergere).
- **Mapping codice→supermodel**: la condivisione è implicita (stesso `Numero componenti` in più supermodel). Confermato: nessuna colonna dedicata necessaria.
- **Coerenza mesi tra supermodel**: se i `Total_demand` hanno orizzonti diversi, allineare per nome mese (union, zero dove assente).

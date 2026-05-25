# -*- coding: utf-8 -*-
"""
================================================================================
MONTE CARLO SIMULATOR - VERSIONE V2 (compatibile con pipeline V2)
================================================================================

DIFFERENZE RISPETTO ALLA VERSIONE ORIGINALE (mc_simulator_OPTIMIZED.py):
-------------------------------------------------------------------------
- Usa TR TOTALV21E - Copia.xlsx invece di TR TOTALV21E.xlsx
  (coerente con gestione_dipendenze_V2.py e skubundle_OPTIMIZED_V2.py)
- Nessun'altra differenza logica: la simulazione è già zero-hardcoding

SCOPO:
Questo modulo esegue simulazioni Monte Carlo per generare combinazioni di
configurazioni moto (Ducati Multistrada V4) basate su distribuzioni Dirichlet.
L'output viene utilizzato per calcolare i safety stock dei componenti.

FLUSSO PRINCIPALE:
1. Carica il forecast di domanda mensile (Total_demand.xlsx)
2. Carica i Take Rate (TR) e parametri Dirichlet (TR TOTALV21E - Copia.xlsx)
3. Carica le regole di esclusione tra opzioni incompatibili (esclusioni.xlsx)
4. Per ogni mese e per ogni run:
   - Campiona le probabilità delle opzioni dalla distribuzione Dirichlet
   - Genera N configurazioni moto (dove N = domanda mensile)
   - Risolve eventuali conflitti tra opzioni incompatibili
5. Aggrega i risultati in un DataFrame wide per il successivo SKU matching

OTTIMIZZAZIONI APPLICATE:
1. Pre-calcolo pesi Dirichlet: campionati UNA volta per run invece che per ogni moto
2. Batch model sampling: selezione modelli vettorizzata con numpy
3. Int16 per array: risparmio 75% memoria rispetto a int64
4. Single Dirichlet per run: i pesi rimangono fissi dentro ogni run
5. Garbage collection strategico: libera memoria ogni 2 mesi simulati

LOGICA INVARIATA: L'output è identico al simulatore originale

@author: AntonioLiardo-SMOPS
================================================================================
"""

import numpy as np
import pandas as pd
from tr_parser_V2 import parse_tr_file, load_monthly_tr, get_tr_for_month, ModelMix, CharacteristicGroup
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
from pathlib import Path
import re
import math
import gc
import time
import sys

_MODULE_DIR = Path(__file__).parent


# ============================================================================
# CONFIGURAZIONE SIMULAZIONE
# ============================================================================

@dataclass
class SimulationConfig:
    """
    Parametri di configurazione per la simulazione Monte Carlo.

    Attributes:
        n_runs: Numero di run Monte Carlo (più run = statistiche più stabili)
        random_seed: Seed per riproducibilità dei risultati
        max_conflict_resolution_iterations: Max tentativi per risolvere conflitti
        start_month_offset: Quanti mesi saltare nel file forecast (0 = primo mese)
        use_int16: Se True usa int16 per risparmiare memoria (75% in meno)
    """
    n_runs: int = 1000
    random_seed: int = 42
    max_conflict_resolution_iterations: int = 10
    start_month_offset: int = 0
    use_int16: bool = True


# ============================================================================
# FUNZIONI DIRICHLET - CAMPIONAMENTO PROBABILITÀ
# ============================================================================
# La distribuzione Dirichlet genera vettori di probabilità che sommano a 1.
# Viene usata per modellare l'incertezza sui Take Rate delle opzioni.
# ============================================================================




# ============================================================================
# MAPPING MODELLO -> PARAMETRI
# ============================================================================
# Ogni modello (MSV4, MSV4S, MSV4PP) ha i propri Take Rate e parametri Dirichlet.
# Queste funzioni restituiscono i parametri corretti per ogni modello.
# ============================================================================

def get_model_tr_mapping(model: str, char_group: CharacteristicGroup) -> np.ndarray:
    """
    Restituisce i Take Rate corretti per il modello specificato.

    VERSIONE SCALABILE: delega il lookup a CharacteristicGroup.get_tr(),
    che cerca il modello nel dizionario tr costruito dinamicamente dall'Excel.
    Nessun if/elif hardcoded: aggiungere un modello in Excel è sufficiente.

    Args:
        model: Nome del modello (es. 'MSV4', 'MSV4S', 'MSV4PP', ...)
        char_group: Gruppo caratteristica

    Returns:
        Array numpy con i Take Rate per ogni opzione
    """
    return char_group.get_tr(model)


def get_model_alpha_mapping(model: str, char_group: CharacteristicGroup) -> np.ndarray:
    """
    Restituisce i parametri Dirichlet (alpha) corretti per il modello.

    VERSIONE SCALABILE: delega il lookup a CharacteristicGroup.get_alpha(),
    che cerca il modello nel dizionario alpha costruito dinamicamente dall'Excel.

    Args:
        model: Nome del modello
        char_group: Gruppo caratteristica

    Returns:
        Array numpy con i parametri alpha Dirichlet
    """
    return char_group.get_alpha(model)


# ============================================================================
# GESTIONE ESCLUSIONI E CONFLITTI
# ============================================================================
# Alcune opzioni sono incompatibili tra loro (es. non puoi avere due colori).
# Queste funzioni gestiscono il caricamento delle regole e la risoluzione.
# ============================================================================

def parse_exclusions(filepath: str = '.\\Input\\esclusioni.xlsx') -> Dict[str, List[str]]:
    """
    Carica le regole di esclusione dal file Excel.

    Il file contiene coppie di opzioni incompatibili. Esempio:
    "Radar Pack" è incompatibile con "Tech Pack base"

    Formato file:
    - Colonna A: Nome opzione
    - Colonna B: Lista opzioni incompatibili (tra virgolette)

    Args:
        filepath: Percorso al file Excel delle esclusioni

    Returns:
        Dizionario {opzione: [lista opzioni incompatibili]}
    """
    try:
        df = pd.read_excel(filepath, sheet_name='Foglio1')
        exclusions = {}

        # Perf: zip su colonne ~10x più veloce di iterrows() (file piccolo, ok lo stesso)
        pattern = r'"([^"]+)"'
        for option_raw, incomp_raw in zip(df.iloc[:, 0], df.iloc[:, 1]):
            option = str(option_raw).strip()
            incompatible_options = re.findall(pattern, str(incomp_raw))
            if incompatible_options:
                exclusions[option] = incompatible_options

        print(f"[OK] Esclusioni caricate: {len(exclusions)} regole")
        return exclusions

    except Exception as e:
        print(f"WARN: Impossibile caricare esclusioni: {e}")
        return {}


def find_conflicts(config: Dict, exclusions: Dict[str, List[str]]) -> List[Tuple]:
    """
    Identifica i conflitti in una configurazione moto.

    Scorre tutte le opzioni selezionate (non-NOT) e verifica se ci sono
    coppie di opzioni che violano le regole di esclusione.

    Args:
        config: Dizionario {caratteristica: valore_selezionato}
        exclusions: Regole di esclusione caricate da parse_exclusions()

    Returns:
        Lista di conflitti, ogni conflitto è ((char1,val1), (char2,val2))
    """
    conflicts = []
    selected_chars = []

    # Raccogli tutte le opzioni "attive" (non NOT, non vuote)
    for char_name, char_value in config.items():
        if char_name == 'MODEL':
            continue

        value_str = str(char_value).strip()

        # YES significa opzione selezionata (per bundle binari)
        if value_str.upper() == 'YES':
            selected_chars.append((char_name, char_value))
        # Qualsiasi valore che non sia NOT è considerato selezionato
        elif value_str.upper() != 'NOT' and 'NOT' not in value_str.upper():
            selected_chars.append((char_name, char_value))

    # Cerca conflitti tra tutte le coppie di opzioni selezionate
    for char1, val1 in selected_chars:
        if char1 in exclusions:
            for incompatible_char_name in exclusions[char1]:
                for char2, val2 in selected_chars:
                    if char2 == incompatible_char_name:
                        conflicts.append(((char1, val1), (char2, val2)))

    # Rimuovi duplicati (A-B e B-A sono lo stesso conflitto)
    unique_conflicts = []
    seen = set()
    for c in conflicts:
        key = tuple(sorted([c[0][0], c[1][0]]))
        if key not in seen:
            unique_conflicts.append(c)
            seen.add(key)

    return unique_conflicts


def get_alpha_for_option(char_name: str, char_value: str, model: str,
                         characteristics: Dict[str, CharacteristicGroup]) -> float:
    """
    Ottiene il parametro alpha Dirichlet per una specifica opzione.

    Usato nella risoluzione conflitti: l'opzione con alpha più alto
    ha più probabilità di essere mantenuta.

    VERSIONE SCALABILE: usa CharacteristicGroup.get_alpha() che risolve
    il modello dinamicamente dal dizionario, senza if/elif hardcoded.

    Args:
        char_name: Nome caratteristica (es. 'FAIRING COLOR')
        char_value: Valore selezionato (es. 'Rosso')
        model: Modello moto
        characteristics: Dizionario di tutte le caratteristiche

    Returns:
        Valore alpha (float). Default 1.0 se non trovato.
    """
    if char_name not in characteristics:
        return 1.0

    char_group  = characteristics[char_name]
    alpha_array = char_group.get_alpha(model)

    # Trova indice del valore e restituisci alpha corrispondente
    try:
        value_idx = char_group.values.index(char_value)
        return float(alpha_array[value_idx])
    except (ValueError, IndexError):
        return 1.0


def force_option_to_not(config: Dict, char_to_change: str,
                        characteristics: Dict[str, CharacteristicGroup]):
    """
    Forza una caratteristica al valore "NOT" (non selezionata).

    Usato per risolvere conflitti: una delle due opzioni in conflitto
    viene forzata a NOT.

    Args:
        config: Configurazione da modificare (modificata in-place)
        char_to_change: Nome caratteristica da forzare a NOT
        characteristics: Dizionario caratteristiche per trovare il valore NOT
    """
    if char_to_change in characteristics:
        char_group = characteristics[char_to_change]

        # Cerca un valore che contenga "NOT"
        not_values = [v for v in char_group.values if 'NOT' in v.upper() or v.upper() == 'NOT']

        if not_values:
            config[char_to_change] = not_values[0]
        else:
            # Fallback: usa ultimo valore (spesso è il default/NOT)
            config[char_to_change] = char_group.values[-1]


def resolve_conflict_weighted(conflict: Tuple, model: str,
                               characteristics: Dict[str, CharacteristicGroup]) -> Tuple[str, str]:
    """
    Risolve un conflitto decidendo quale opzione forzare a NOT.

    La decisione è probabilistica, pesata sugli alpha:
    - L'opzione con alpha più alto ha più probabilità di essere MANTENUTA
    - L'opzione con alpha più basso ha più probabilità di essere RIMOSSA

    Args:
        conflict: Tupla ((char1,val1), (char2,val2)) delle opzioni in conflitto
        model: Modello moto (per determinare gli alpha corretti)
        characteristics: Dizionario caratteristiche

    Returns:
        Tupla (char_da_rimuovere, valore_da_rimuovere)
    """
    (char1, val1), (char2, val2) = conflict

    # Ottieni alpha per entrambe le opzioni
    alpha1 = get_alpha_for_option(char1, val1, model, characteristics)
    alpha2 = get_alpha_for_option(char2, val2, model, characteristics)

    # Calcola probabilità di mantenere opzione 1
    total = alpha1 + alpha2
    if total <= 0:
        p1 = 0.5  # Se entrambi zero, 50/50
    else:
        p1 = alpha1 / total

    # Decisione casuale pesata
    if np.random.random() < p1:
        # Mantieni opzione 1, rimuovi opzione 2
        return (char2, val2)
    else:
        # Mantieni opzione 2, rimuovi opzione 1
        return (char1, val1)


# ============================================================================
# GENERAZIONE CONFIGURAZIONI - VERSIONE OTTIMIZZATA
# ============================================================================
# Queste funzioni generano configurazioni moto campionando dalla Dirichlet.
# L'ottimizzazione chiave è il pre-calcolo dei pesi per evitare di
# ricampionare la Dirichlet per ogni singola moto.
# ============================================================================

@dataclass
class PrecomputedCharWeights:
    """
    Cache dei pesi Dirichlet pre-computati per una caratteristica.

    Invece di campionare dalla Dirichlet per ogni moto, campioniamo UNA volta
    per run e riusiamo questi pesi per tutte le moto della run.

    Attributes:
        probs: Vettore probabilità (somma a 1)
        values: Lista valori validi (opzioni con TR > 0)
        valid_indices: Indici originali delle opzioni valide
    """
    probs: np.ndarray
    values: List[str]
    valid_indices: np.ndarray


def precompute_char_weights(model: str,
                            characteristics: Dict[str, CharacteristicGroup]) -> Dict[str, PrecomputedCharWeights]:
    """
    Pre-computa i pesi Dirichlet per TUTTE le caratteristiche di un modello.

    Questa è l'ottimizzazione chiave "Single Dirichlet per run":
    - Invece di campionare N volte (una per moto), campioniamo UNA volta
    - I pesi rimangono fissi per tutte le moto della run
    - Simula uno scenario "in questa run il mercato ha queste preferenze"

    Args:
        model: Modello moto
        characteristics: Dizionario di tutte le caratteristiche

    Returns:
        Dizionario {nome_caratteristica: PrecomputedCharWeights}
    """
    weights = {}

    for char_name, char_group in characteristics.items():
        # Ottieni parametri per questo modello
        alpha = get_model_alpha_mapping(model, char_group)
        tr = get_model_tr_mapping(model, char_group)

        # Filtra opzioni con TR > 0 (opzioni impossibili escluse)
        valid_mask = tr > 0

        if not np.any(valid_mask):
            # Fallback: se tutto è TR=0, distribuzione uniforme
            probs = np.ones(len(char_group.values)) / len(char_group.values)
            valid_indices = np.arange(len(char_group.values))
            valid_values = char_group.values
        else:
            # Filtra e campiona solo dalle opzioni valide
            filtered_alpha = alpha[valid_mask]
            valid_indices = np.where(valid_mask)[0]
            valid_values = [char_group.values[i] for i in valid_indices]

            if filtered_alpha.sum() <= 0:
                probs = np.ones(len(valid_values)) / len(valid_values)
            else:
                safe_alpha = np.where(filtered_alpha <= 0, 0.001, filtered_alpha)
                # CAMPIONAMENTO DIRICHLET: questo è il punto chiave!
                probs = np.random.dirichlet(safe_alpha)

        weights[char_name] = PrecomputedCharWeights(
            probs=probs,
            values=valid_values,
            valid_indices=valid_indices
        )

    return weights


def generate_configuration_fast(model: str,
                                char_names: List[str],
                                char_weights: Dict[str, PrecomputedCharWeights]) -> Dict:
    """
    Genera una singola configurazione moto usando pesi PRE-COMPUTATI.

    NON ricampiona dalla Dirichlet - usa i pesi già estratti.
    Molto più veloce perché fa solo scelte casuali con probabilità fisse.

    Args:
        model: Modello moto da inserire nella configurazione
        char_names: Lista nomi caratteristiche da configurare
        char_weights: Pesi pre-computati da precompute_char_weights()

    Returns:
        Dizionario configurazione {caratteristica: valore_scelto}
    """
    config = {'MODEL': model}

    for char_name in char_names:
        if char_name in char_weights:
            w = char_weights[char_name]
            # Scelta casuale pesata (molto veloce con numpy)
            chosen_idx = np.random.choice(len(w.values), p=w.probs)
            config[char_name] = w.values[chosen_idx]

    return config


def generate_valid_configuration_fast(model: str,
                                      characteristics: Dict[str, CharacteristicGroup],
                                      char_names: List[str],
                                      exclusions: Dict[str, List[str]],
                                      char_weights: Dict[str, PrecomputedCharWeights],
                                      max_iterations: int = 10) -> Tuple[Dict, int]:
    """
    Genera una configurazione VALIDA risolvendo eventuali conflitti.

    Processo:
    1. Genera configurazione iniziale con generate_configuration_fast()
    2. Cerca conflitti con find_conflicts()
    3. Se ci sono conflitti, risolvi con resolve_conflict_weighted()
    4. Ripeti fino a nessun conflitto o max iterazioni

    Args:
        model: Modello moto
        characteristics: Dizionario caratteristiche
        char_names: Lista nomi caratteristiche
        exclusions: Regole esclusione
        char_weights: Pesi pre-computati
        max_iterations: Max tentativi risoluzione conflitti

    Returns:
        Tupla (configurazione_valida, numero_conflitti_risolti)
    """
    total_conflicts_resolved = 0

    # Genera configurazione iniziale
    config = generate_configuration_fast(model, char_names, char_weights)

    # Ciclo di risoluzione conflitti
    last_conflicts_signature = None
    for iteration in range(max_iterations):
        conflicts = find_conflicts(config, exclusions)

        if not conflicts:
            break  # Nessun conflitto, configurazione valida

        # Detect stagnazione: se i conflitti non cambiano tra iterazioni,
        # force_option_to_not non sta risolvendo nulla (es. char non in
        # characteristics). Esci subito per evitare loop infinito silenzioso.
        sig = tuple(sorted(str(c) for c in conflicts))
        if sig == last_conflicts_signature:
            break
        last_conflicts_signature = sig

        total_conflicts_resolved += len(conflicts)

        # Risolvi ogni conflitto
        for conflict in conflicts:
            char_to_change, _ = resolve_conflict_weighted(conflict, model, characteristics)
            force_option_to_not(config, char_to_change, characteristics)

    return config, total_conflicts_resolved


# ============================================================================
# CARICAMENTO DATI INPUT
# ============================================================================

def load_monthly_forecast(filepath: str = '.\\Input\\Total_demand.xlsx',
                          start_offset: int = 0) -> Tuple[List[float], List[str]]:
    """
    Carica il forecast di domanda mensile dal file Excel.

    Il file ha una struttura specifica:
    - Riga 2 (indice 1): nomi dei mesi
    - Riga 26 (indice 25): valori forecast per la linea Multistrada

    Args:
        filepath: Percorso al file Total_demand.xlsx
        start_offset: Quanti mesi saltare dall'inizio (0 = primo mese)

    Returns:
        Tupla (lista_valori_forecast, lista_nomi_mesi)
    """
    print(f"\n{'='*60}")
    print(f"CARICAMENTO FORECAST MENSILE")
    print(f"{'='*60}")

    df = pd.read_excel(filepath, header=None)

    # Riga 1: nomi mesi
    month_row = df.iloc[0, 0:]
    month_names = [str(m) for m in month_row if pd.notna(m) and 'TOT' not in str(m)]

    # Riga 26: valori forecast
    forecast_row = df.iloc[1, 0:]
    forecast_values = [float(v) for v in forecast_row[:len(month_names)] if pd.notna(v)]

    # Applica offset per partire dal mese desiderato
    forecast_values = forecast_values[start_offset:]
    month_names = month_names[start_offset:]

    # Avvisa se il numero di mesi è insolitamente alto (possibile errore struttura file)
    if len(month_names) > 60:
        print(f"\n  *** ATTENZIONE ***")
        print(f"  Trovati {len(month_names)} mesi nel file previsioni (attesi tipicamente <=60).")
        print(f"  Questo può causare un consumo di RAM molto elevato durante la simulazione.")
        print(f"  Verificare la struttura del file '{filepath}':")
        print(f"  - Riga 1 deve contenere i nomi dei mesi (es. 'Gen 2026', 'Feb 2026', ...)")
        print(f"  - Riga 2 deve contenere i valori forecast numerici")
        print(f"  Primi 5 nomi mesi trovati: {month_names[:5]}")

    print(f"Mesi disponibili: {len(forecast_values)}")
    print(f"Mese iniziale: {month_names[0] if month_names else 'N/A'}")

    return forecast_values, month_names


def calculate_n_months_needed() -> int:
    """
    Calcola quanti mesi simulare basandosi sul lead time massimo dei componenti.
    Legge da Gates MTSV4.xlsx (Dati Logistica) — col. 'Gates (Weeks)'.
    """
    from pathlib import Path

    gates_file = Path('.\\Input\\Dati Logistica\\Gates MTSV4.xlsx')

    if not gates_file.exists():
        raise FileNotFoundError(
            f"File obbligatorio mancante: {gates_file}\n"
            f"Necessario per calcolare max lead time durante simulazione Monte Carlo."
        )

    df_gates = pd.read_excel(str(gates_file), sheet_name=0, skiprows=1)
    col_name = 'Gates (Weeks)'

    if col_name not in df_gates.columns:
        raise ValueError(
            f"Colonna '{col_name}' non trovata in Gates MTSV4.xlsx\n"
            f"Colonne disponibili: {list(df_gates.columns)}"
        )

    all_values = df_gates[col_name].dropna()
    all_values = pd.to_numeric(all_values, errors='coerce').dropna()

    # Converti settimane -> mesi
    all_values_months = all_values / 4.33

    MAX_PLAUSIBLE_MONTHS = 60
    valid_values  = all_values_months[(all_values_months > 0) & (all_values_months <= MAX_PLAUSIBLE_MONTHS)]
    anomalous_vals = all_values_months[(all_values_months <= 0) | (all_values_months > MAX_PLAUSIBLE_MONTHS)]

    if len(anomalous_vals) > 0:
        print(f"\n  *** ATTENZIONE ***")
        print(f"  {len(anomalous_vals)} riga/righe con gate weeks anomalo "
              f"(<=0 o >{MAX_PLAUSIBLE_MONTHS} mesi equivalenti) — IGNORATI nel calcolo del max:")
        print(f"  Valori anomali: {list(anomalous_vals.values[:10])}")

    if len(valid_values) == 0:
        raise ValueError(
            "Nessun valore valido in Gates (Weeks) di Gates MTSV4.xlsx.\n"
            "Verificare che il file contenga dati numerici validi."
        )

    max_lead_time = valid_values.max()
    n_months = math.ceil(max_lead_time)

    print(f"Max lead time (valori validi <={MAX_PLAUSIBLE_MONTHS}m): "
          f"{max_lead_time:.1f} mesi -> {n_months} mesi da simulare")

    return n_months


# ============================================================================
# SIMULATORE MONTE CARLO - CORE
# ============================================================================

def run_single_month_simulation_optimized(
    month_index: int,
    month_name: str,
    total_demand: float,
    model_mix: ModelMix,
    characteristics: Dict[str, CharacteristicGroup],
    exclusions: Dict[str, List[str]],
    config: SimulationConfig,
    monthly_tr: Dict = None
) -> Dict[Tuple, np.ndarray]:
    """
    Esegue la simulazione Monte Carlo per UN singolo mese.

    Per ogni run:
    1. Campiona i pesi Dirichlet (UNA volta per run - ottimizzazione chiave)
    2. Campiona i modelli per tutte le moto del mese
    3. Per ogni moto: genera configurazione usando i pesi pre-campionati
    4. Conta quante volte appare ogni combinazione unica

    TR MENSILI:
    Se monthly_tr è fornito, usa i Take Rate specifici del mese (con fallback
    al mese disponibile più recente tramite get_tr_for_month()).
    Se monthly_tr è None, usa i parametri globali passati (model_mix, characteristics).

    Args:
        month_index: Indice del mese (0-based)
        month_name: Nome mese per logging (es. "Jan 2026")
        total_demand: Domanda totale del mese (numero moto da generare)
        model_mix: Mix modelli globale (usato come fallback se monthly_tr è None)
        characteristics: Caratteristiche globali (usate come fallback se monthly_tr è None)
        exclusions: Regole esclusione
        config: Parametri simulazione
        monthly_tr: Dict {mese: (ModelMix, characteristics)} con TR mensili.
                    Se None, usa i parametri globali model_mix e characteristics.

    Returns:
        Dizionario {chiave_configurazione: array_conteggi_per_run}
        dove chiave_configurazione è una tupla ordinata di (caratteristica, valore)
    """
    # ------------------------------------------------------------------
    # Risolvi TR per questo mese specifico
    # ------------------------------------------------------------------
    if monthly_tr is not None:
        month_model_mix, month_characteristics = get_tr_for_month(month_name, monthly_tr)
    else:
        # Fallback: usa i parametri globali passati alla funzione
        month_model_mix   = model_mix
        month_characteristics = characteristics
    print(f"\n  Mese {month_index + 1}: {month_name} (domanda: {total_demand:.0f})")
    sys.stdout.flush()

    # Imposta seed per riproducibilità (diverso per ogni mese e per ogni offset rolling).
    # Moltiplica offset per 1000 per evitare sovrapposizioni tra iterazioni rolling.
    np.random.seed(config.random_seed + config.start_month_offset * 1000 + month_index)

    char_names = list(month_characteristics.keys())

    # Filtra modelli: esclude "ALTRO" che non è un modello reale
    # Normalizza nomi modello: 'MSV RS' -> 'MSV4RS', 'MSV RI' -> 'MSV4RI', ecc.
    # Serve quando il foglio TR_Model ha nomi come 'MSV RS' invece di 'MSV4RS',
    # che altrimenti non farebbero match con le colonne BOM (es. MSV4RS).
    def _normalize_model_name(m: str) -> str:
        """Normalizza nome modello MSV4-family: 'MSV RS' / 'MSVRS' -> 'MSV4RS'."""
        norm = m.strip().upper().replace(' ', '').replace('_', '')
        # Se inizia con 'MSV' ma non ha '4' in quarta posizione -> inserisci '4'
        if norm.startswith('MSV') and len(norm) > 3 and norm[3] != '4':
            return 'MSV4' + norm[3:]
        return m.strip()

    valid_models = [
        _normalize_model_name(m)
        for m in month_model_mix.models
        if m.strip().upper() != "ALTRO"
    ]

    # Rimuovi alpha del modello ALTRO
    altro_idx = None
    for i, m in enumerate(month_model_mix.models):
        if m.strip().upper() == "ALTRO":
            altro_idx = i
            break

    if altro_idx is not None:
        valid_alphas = np.delete(month_model_mix.alphas, altro_idx)
    else:
        valid_alphas = month_model_mix.alphas

    # Pre-computa pesi caratteristiche (verranno aggiornati per ogni run)
    model_char_weights = {}
    for model in valid_models:
        model_char_weights[model] = precompute_char_weights(model, month_characteristics)

    # Struttura risultati: {config_key: array[n_runs] di conteggi}
    # Usa int16 per risparmiare memoria (max 32767 per combinazione per run)
    combination_demands_by_run = defaultdict(
        lambda: np.zeros(config.n_runs, dtype=np.int16 if config.use_int16 else np.int32)
    )

    stats_total_configs = 0
    stats_configs_with_conflicts = 0

    target_demand = int(total_demand)

    # =====================================================================
    # LOOP PRINCIPALE: per ogni run Monte Carlo
    # =====================================================================
    for run in range(config.n_runs):
        # Progress indicator
        if run % 200 == 0 and run > 0:
            print(f"    Run {run}/{config.n_runs}", end='\r')
            sys.stdout.flush()

        # -----------------------------------------------------------------
        # OTTIMIZZAZIONE 1: Campiona probabilità modelli UNA volta per run
        # -----------------------------------------------------------------
        model_probs = np.random.dirichlet(valid_alphas)

        # -----------------------------------------------------------------
        # OTTIMIZZAZIONE 2: Pre-campiona tutti i modelli in batch
        # -----------------------------------------------------------------
        model_choices = np.random.choice(len(valid_models), size=target_demand, p=model_probs)

        # -----------------------------------------------------------------
        # OTTIMIZZAZIONE 3: Ri-calcola pesi caratteristiche UNA volta per run
        # Questa è la logica "Single Dirichlet per run"
        # -----------------------------------------------------------------
        for model in valid_models:
            model_char_weights[model] = precompute_char_weights(model, month_characteristics)

        # Contatore combinazioni per questa run
        run_combination_counts = defaultdict(int)

        # Genera ogni moto del mese
        for bike_idx in range(target_demand):
            model_idx = model_choices[bike_idx]
            model = valid_models[model_idx]

            stats_total_configs += 1

            # Genera configurazione usando pesi pre-computati
            config_obj, num_conflicts = generate_valid_configuration_fast(
                model, month_characteristics, char_names, exclusions,
                model_char_weights[model],
                config.max_conflict_resolution_iterations
            )

            if num_conflicts > 0:
                stats_configs_with_conflicts += 1

            # Crea chiave univoca per questa configurazione
            config_key = tuple(sorted(config_obj.items()))
            run_combination_counts[config_key] += 1

        # Salva conteggi di questa run
        for config_key, count in run_combination_counts.items():
            combination_demands_by_run[config_key][run] = count

    print(f"    Run {config.n_runs}/{config.n_runs} - {len(combination_demands_by_run)} combinazioni uniche")

    return dict(combination_demands_by_run)


def run_monthly_simulations_optimized(
    forecast_values: List[float],
    month_names: List[str],
    n_months: int,
    model_mix: ModelMix,
    characteristics: Dict[str, CharacteristicGroup],
    exclusions: Dict[str, List[str]],
    config: SimulationConfig,
    monthly_tr: Dict = None
) -> Dict[int, Dict[Tuple, np.ndarray]]:
    """
    Esegue le simulazioni Monte Carlo per TUTTI i mesi richiesti.

    Itera sui mesi chiamando run_single_month_simulation_optimized()
    e raccoglie i risultati in un dizionario.

    TR MENSILI:
    Se monthly_tr è fornito, ogni mese usa i propri TR (con fallback al mese
    più recente disponibile tramite get_tr_for_month()).
    Se monthly_tr è None, tutti i mesi usano i parametri globali (model_mix,
    characteristics), comportamento identico alla versione precedente.

    Args:
        forecast_values: Lista domanda per ogni mese
        month_names: Lista nomi mesi
        n_months: Quanti mesi simulare
        model_mix: Mix modelli globale (fallback se monthly_tr è None)
        characteristics: Caratteristiche globali (fallback se monthly_tr è None)
        exclusions: Esclusioni
        config: Parametri simulazione
        monthly_tr: Dict {mese: (ModelMix, characteristics)} con TR mensili.
                    Se None usa i parametri globali per tutti i mesi.

    Returns:
        Dizionario {indice_mese: {config_key: array_conteggi}}
    """
    print(f"\n{'='*60}")
    print(f"SIMULAZIONI MULTI-MESE (OTTIMIZZATO)")
    print(f"{'='*60}")
    print(f"Mesi da simulare: {n_months}")
    print(f"Run per mese: {config.n_runs}")
    print(f"Dtype: {'int16' if config.use_int16 else 'int32'}")

    # Verifica disponibilità mesi
    if n_months > len(forecast_values):
        print(f"WARNING: Richiesti {n_months} mesi ma disponibili solo {len(forecast_values)}")
        n_months = len(forecast_values)

    results_by_month = {}
    start_time = time.time()

    # Simula ogni mese
    for month_idx in range(n_months):
        month_start = time.time()

        results_by_month[month_idx] = run_single_month_simulation_optimized(
            month_index=month_idx,
            month_name=month_names[month_idx],
            total_demand=forecast_values[month_idx],
            model_mix=model_mix,
            characteristics=characteristics,
            exclusions=exclusions,
            config=config,
            monthly_tr=monthly_tr
        )

        month_elapsed = time.time() - month_start
        print(f"    Tempo: {month_elapsed:.1f}s")

        # Garbage collection periodico per liberare memoria
        if month_idx % 2 == 0:
            gc.collect()

    total_elapsed = time.time() - start_time
    print(f"\n[OK] Simulazione completata in {total_elapsed:.1f}s ({total_elapsed/60:.1f}min)")

    return results_by_month


# ============================================================================
# VERIFICA RISULTATI
# ============================================================================

def verify_demand_conservation(results_by_month: Dict[int, Dict[Tuple, np.ndarray]],
                               n_runs: int, forecast_values: List[float]):
    """
    Verifica che la domanda totale sia conservata in ogni simulazione.

    La somma delle configurazioni generate deve essere uguale alla domanda
    del mese. Se non lo è, c'è un bug nella simulazione.

    Args:
        results_by_month: Risultati simulazione
        n_runs: Numero run
        forecast_values: Domanda attesa per mese

    Returns:
        True se tutti i mesi sono corretti, False altrimenti
    """
    print(f"\n{'='*60}")
    print(f"VERIFICA CONSERVAZIONE DOMANDA")
    print(f"{'='*60}")

    all_correct = True
    for month_idx, month_results in results_by_month.items():
        if month_idx >= len(forecast_values):
            continue

        expected_demand = forecast_values[month_idx]

        # Verifica sul primo run (rappresentativo)
        run_total = sum(int(demands[0]) for demands in month_results.values())

        if abs(run_total - expected_demand) > 1:  # Tolleranza di 1 unità
            all_correct = False
            print(f"Mese {month_idx+1}: {run_total:.0f} / {expected_demand:.0f} X")
        else:
            print(f"Mese {month_idx+1}: [OK] {expected_demand:.0f}")

    return all_correct


# ============================================================================
# COSTRUZIONE OUTPUT DATAFRAME
# ============================================================================

def build_wide_dataframe(
    results_by_month: Dict[int, Dict[Tuple, np.ndarray]],
    month_names: List[str],
    n_runs: int,
    save_to_file: bool = False,
    filename: str = "monthly_output_wide.xlsx"
) -> pd.DataFrame:
    """
    Costruisce il DataFrame in formato "wide" per il successivo SKU matching.

    Il formato wide ha:
    - Una riga per ogni combinazione unica di configurazione
    - Colonne caratteristiche (MODEL, FAIRING COLOR, etc.)
    - Colonne per ogni run di ogni mese (Jan2026_Run_0, Jan2026_Run_1, ...)

    OTTIMIZZAZIONE: NON salva su file di default (passa direttamente in memoria)

    Args:
        results_by_month: Risultati simulazione
        month_names: Nomi mesi
        n_runs: Numero run
        save_to_file: Se True, salva anche su Excel
        filename: Nome file output

    Returns:
        DataFrame pandas in formato wide
    """
    print(f"\n{'='*60}")
    print(f"COSTRUZIONE DATAFRAME WIDE (in memoria)")
    print(f"{'='*60}")

    # Raccogli tutte le configurazioni uniche da tutti i mesi
    all_configs = set()
    for month_results in results_by_month.values():
        all_configs.update(month_results.keys())

    all_configs = sorted(all_configs)
    print(f"Combinazioni totali: {len(all_configs)}")

    if not all_configs:
        return pd.DataFrame()

    # Estrai nomi caratteristiche dalla prima configurazione
    char_names = [k for k, v in all_configs[0]]

    # -----------------------------------------------------------------
    # COSTRUZIONE COLONNE CARATTERISTICHE
    # -----------------------------------------------------------------
    char_data = {char_names[i]: [] for i in range(len(char_names))}

    for config_key in all_configs:
        config_dict = dict(config_key)
        for i, char_name in enumerate(char_names):
            char_data[char_names[i]].append(config_dict.get(char_name, ''))

    # -----------------------------------------------------------------
    # COSTRUZIONE COLONNE RUN (una per ogni run di ogni mese)
    # -----------------------------------------------------------------
    print("Aggiungendo colonne run...")

    # Pre-alloca un unico array numpy per TUTTI i mesi × run
    n_months_to_build = len(results_by_month)
    all_col_names = []
    all_run_data = np.zeros((len(all_configs), n_months_to_build * n_runs), dtype=np.int16)

    for list_idx, month_idx in enumerate(sorted(results_by_month.keys())):
        month_name = month_names[month_idx] if month_idx < len(month_names) else f"Month_{month_idx}"
        month_results = results_by_month[month_idx]

        col_start = list_idx * n_runs
        for row_idx, config_key in enumerate(all_configs):
            if config_key in month_results:
                demands = month_results[config_key]
                n = min(len(demands), n_runs)
                all_run_data[row_idx, col_start:col_start + n] = demands[:n]

        all_col_names.extend([f"{month_name}_Run_{run_idx}" for run_idx in range(n_runs)])

        if list_idx % 2 == 0:
            print(f"  Mese {list_idx + 1}/{n_months_to_build}")

    # Costruisce il DataFrame finale SENZA pd.concat (che copia internamente il blocco int16)
    # copy=False -> DataFrame condivide la memoria di all_run_data senza duplicarla
    # insert()   -> aggiunge le colonne char in-place senza toccare il blocco int16
    if n_months_to_build > 0:
        df = pd.DataFrame(all_run_data, columns=all_col_names, copy=False)
        del all_run_data
        for i, col_name in enumerate(char_names):
            df.insert(i, col_name, char_data[col_name])
    else:
        df = pd.DataFrame(char_data)

    print(f"  Righe: {len(df)}")
    print(f"  Colonne: {len(df.columns)}")

    # Stima memoria
    mem_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
    print(f"  Memoria DataFrame: {mem_mb:.1f} MB")

    # Salva su file SOLO se esplicitamente richiesto
    if save_to_file:
        print(f"\n  Salvataggio su {filename}...")
        df.to_excel(filename, index=False)
        print(f"  [OK] File salvato: {filename}")

    return df


# ============================================================================
# DERIVAZIONE COLORI PACK (Touring Pack / Top Case Pack)
# ============================================================================

def _fix_color_typos(s: str) -> str:
    """Corregge i typo noti nei nomi colore Ducati (input: stringa lowercase)."""
    s = s.replace('strom ', 'storm ')         # MV: "Strom Green" -> "Storm Green"
    s = s.replace('sparking ', 'sparkling ')  # DSP: "Sparking Blue" -> "Sparkling Blue"
    s = s.replace(' bast ', ' blast ')        # L01: "Viola Bast" -> "Viola Blast"
    return s


def _normalize_color_desc_for_match(s: str) -> str:
    """Normalizza descrizione colore per matching fuzzy.
    Lowercase + correggi typo + rimuovi suffissi decorativi (gloss/matt/matte/satin).
    """
    s = str(s).lower().strip()
    s = _fix_color_typos(s)
    for suf in (' gloss', ' matt', ' matte', ' satin'):
        if s.endswith(suf):
            s = s[:-len(suf)].strip()
    return s


def build_zcol_pack_color_mapping(mappatura_path: str,
                                   rally_mappatura_path: str = None) -> dict:
    """
    Legge la mappatura standard (Schema sheet) e opzionalmente la mappatura Rally,
    costruendo il lookup:
      FAIRING COLOR (desc lowercase, typo-corrected)
        -> colore Panniers covers (color variant)   [Touring Pack V21E — Mappatura - Copia.xlsx]
        -> colore Top Case cover (color variant)    [Top Case Pack V21E — Mappatura - Copia.xlsx]
        -> colore Panniers Covers                   [Touring pack Rally V22E — Mappatura Rally.xlsx]

    Returns:
        {
          'Panniers covers (color variant)': { zcol_desc_canonical: ct422_desc, ... },
          'Top Case cover (color variant)':  { zcol_desc_canonical: ct478_desc, ... },
          'Panniers Covers':                 { zcol_desc_canonical: ct422_rally_desc, ... },
        }
    Dove zcol_desc_canonical = lowercase + typo-corrected (senza rimozione suffissi).
    """
    try:
        xl = pd.ExcelFile(mappatura_path)
        if 'Schema_Long' in xl.sheet_names:
            df_map_raw = xl.parse('Schema_Long', header=0)
            # Normalizza nomi colonne per compatibilità con logica sotto
            df_map = df_map_raw.rename(columns={
                'CT': 'Codice caratteristica',
                'CT_Desc': 'Descrizione caratteristica',
                'CV': '_CV', 'CV_Desc': '_CV_Desc',
            })
            _use_long = True
        else:
            df_map = xl.parse('Schema', header=0)
            _use_long = False
    except Exception as e:
        print(f"  [WARN] Impossibile leggere mappatura '{mappatura_path}': {e}")
        return {}

    def _get_descs_long(ct_code, bundle_lower):
        """Estrae CV_Desc per CT+Bundle dal formato long."""
        mask = (
            (df_map['Codice caratteristica'] == ct_code) &
            (df_map['Bundle'].astype(str).str.strip().str.lower() == bundle_lower)
        )
        rows = df_map[mask]
        return [str(r).strip() for r in rows['_CV_Desc'] if pd.notna(r) and str(r).strip() not in ('', 'nan')]

    def _get_descs_wide(row):
        """Estrae Desc1..Desc19 dal formato wide."""
        descs = []
        for i in range(1, 20):
            d = row.get(f'Desc{i}', None)
            if pd.notna(d) and str(d).strip() not in ('', 'nan'):
                descs.append(str(d).strip())
        return descs

    # ZCOL
    zcol_rows = df_map[df_map['Codice caratteristica'] == 'ZCOL']
    if zcol_rows.empty:
        print("  [WARN] ZCOL non trovato in mappatura -> colori pack non derivati")
        return {}
    if _use_long:
        zcol_descs = [str(r).strip() for r in zcol_rows['_CV_Desc']
                      if pd.notna(r) and str(r).strip() not in ('', 'nan')]
    else:
        zcol_descs = _get_descs_wide(zcol_rows.iloc[0])

    # CT000422 Touring Pack (Panniers covers)
    if _use_long:
        # Cerca sia "Touring Pack" che "Touring Pack Rally"
        ct422_descs = _get_descs_long('CT000422', 'touring pack')
        if not ct422_descs:
            ct422_descs = _get_descs_long('CT000422', 'touring pack rally')
    else:
        ct422_rows = df_map[
            (df_map['Codice caratteristica'] == 'CT000422') &
            (df_map['Bundle'].astype(str).str.strip().str.lower() == 'touring pack')
        ]
        ct422_descs = _get_descs_wide(ct422_rows.iloc[0]) if not ct422_rows.empty else []

    # CT000478 Top Case Pack (Top Case cover)
    if _use_long:
        ct478_descs = _get_descs_long('CT000478', 'top case pack')
    else:
        ct478_rows = df_map[
            (df_map['Codice caratteristica'] == 'CT000478') &
            (df_map['Bundle'].astype(str).str.strip().str.lower() == 'top case pack')
        ]
        ct478_descs = _get_descs_wide(ct478_rows.iloc[0]) if not ct478_rows.empty else []

    # Build normalized lookup tables for CT422 and CT478
    # Key: normalized (no suffix, typo-fixed) -> original readable description
    ct422_norm = {_normalize_color_desc_for_match(d): d for d in ct422_descs}
    ct478_norm = {_normalize_color_desc_for_match(d): d for d in ct478_descs}

    mapping_422 = {}  # { zcol_canonical_key: ct422_readable_desc }
    mapping_478 = {}  # { zcol_canonical_key: ct478_readable_desc }

    for zcol_desc in zcol_descs:
        # Canonical key: lowercase + typo-fixed (keeps suffixes like "gloss"/"matt")
        zcol_key = _fix_color_typos(zcol_desc.lower())
        # Normalized key: also removes suffixes -> for matching CT descs
        zcol_norm = _normalize_color_desc_for_match(zcol_desc)

        if zcol_norm in ct422_norm:
            mapping_422[zcol_key] = ct422_norm[zcol_norm]
        else:
            print(f"  [INFO] ZCOL '{zcol_desc}' -> CT000422: nessun match")

        if zcol_norm in ct478_norm:
            mapping_478[zcol_key] = ct478_norm[zcol_norm]
        else:
            print(f"  [INFO] ZCOL '{zcol_desc}' -> CT000478: nessun match")

    # Aggiungi mapping espliciti per casi combinati che non matchano per normalizzazione.
    # "Ducati Red" e "Pikes Peak Livery" vengono mappati come valori SEPARATI (non al
    # combined "(Ducati Red+PP)"), in modo che la simulazione produca righe distinte
    # per ciascun colore e il matching SKU funzioni correttamente.
    for ct_mapping in [mapping_422, mapping_478]:
        ct_mapping['ducati red']          = 'Ducati Red'
        ct_mapping['pikes peak livery']   = 'Pikes Peak Livery'

    print(f"  Mapping CT000422 (Panniers V21E):       {len(mapping_422)} ZCOL -> colore")
    print(f"  Mapping CT000478 (Top Case):            {len(mapping_478)} ZCOL -> colore")

    result = {
        'Panniers covers (color variant)': mapping_422,
        'Top Case cover (color variant)':  mapping_478,
    }

    # -------------------------------------------------------------------------
    # Touring Pack Rally (Mappatura Rally.xlsx, CT000422 bundle "Touring pack Rally")
    # -------------------------------------------------------------------------
    if rally_mappatura_path:
        try:
            xl_r = pd.ExcelFile(rally_mappatura_path)
            if 'Schema_Long' in xl_r.sheet_names:
                df_rally_raw = xl_r.parse('Schema_Long', header=0)
                df_rally = df_rally_raw.rename(columns={
                    'CT': 'Codice caratteristica',
                    'CT_Desc': 'Descrizione caratteristica',
                    'CV': '_CV', 'CV_Desc': '_CV_Desc',
                })
                _rally_long = True
            else:
                df_rally = xl_r.parse('Schema', header=0)
                _rally_long = False
        except Exception as e:
            print(f"  [WARN] Impossibile leggere mappatura Rally '{rally_mappatura_path}': {e}")
            df_rally = None
            _rally_long = False

        if df_rally is not None:
            ct422_rally_rows = df_rally[
                (df_rally['Codice caratteristica'] == 'CT000422') &
                (df_rally['Bundle'].astype(str).str.strip().str.lower() == 'touring pack rally')
            ]
            if _rally_long:
                ct422_rally_descs = [str(r).strip() for r in ct422_rally_rows['_CV_Desc']
                                     if pd.notna(r) and str(r).strip() not in ('', 'nan')]
            else:
                ct422_rally_descs = (
                    _get_descs_wide(ct422_rally_rows.iloc[0]) if not ct422_rally_rows.empty else []
                )

            if ct422_rally_descs:
                ct422_rally_norm = {
                    _normalize_color_desc_for_match(d): d for d in ct422_rally_descs
                }

                # ZCOL Rally: cerca nel file Rally, altrimenti usa CT422 Rally come proxy
                zcol_rally_rows = df_rally[df_rally['Codice caratteristica'] == 'ZCOL']
                if not zcol_rally_rows.empty:
                    if _rally_long:
                        zcol_rally_descs = [str(r).strip() for r in zcol_rally_rows['_CV_Desc']
                                            if pd.notna(r) and str(r).strip() not in ('', 'nan')]
                    else:
                        zcol_rally_descs = _get_descs_wide(zcol_rally_rows.iloc[0])
                    print(f"  ZCOL Rally trovato: {len(zcol_rally_descs)} colori")
                else:
                    # Fallback: ogni CT422 Rally desc mappa su se stessa
                    zcol_rally_descs = ct422_rally_descs
                    print(f"  ZCOL Rally non trovato -> uso CT422 Rally come proxy "
                          f"({len(zcol_rally_descs)} colori)")

                # Nel Mappatura Rally (V22E), ZCOL usa formato "Brushed Alu / Ducati Red".
                # In df_wide, FAIRING COLOR per V22E puo' arrivare in DUE formati:
                #   - "brushed alu / ducati red"  (V22E Rally, ZCOL con prefisso manubrio)
                #   - "ducati red"                (V22E plain, senza prefisso)
                # Registriamo ENTRAMBE le chiavi -> stesso valore CT422 Rally.
                _BRUSHED_ALU_PREFIX = 'brushed alu / '
                mapping_422_rally = {}
                for zcol_desc in zcol_rally_descs:
                    zcol_key_full  = _fix_color_typos(zcol_desc.lower())
                    # Chiave plain: strip "brushed alu / " se presente
                    zcol_for_match = zcol_desc
                    if zcol_for_match.lower().startswith(_BRUSHED_ALU_PREFIX):
                        zcol_for_match = zcol_for_match[len(_BRUSHED_ALU_PREFIX):]
                    zcol_key_plain = _fix_color_typos(zcol_for_match.lower())
                    zcol_norm      = _normalize_color_desc_for_match(zcol_for_match)
                    if zcol_norm in ct422_rally_norm:
                        ct_val = ct422_rally_norm[zcol_norm]
                        mapping_422_rally[zcol_key_full]  = ct_val  # "brushed alu / x"
                        mapping_422_rally[zcol_key_plain] = ct_val  # "x" (plain)
                    else:
                        print(f"  [INFO] ZCOL Rally '{zcol_desc}' -> CT000422 Rally: nessun match")

                print(f"  Mapping CT000422 Rally (Panniers V22E):  "
                      f"{len(mapping_422_rally)} ZCOL -> colore")
                result['Panniers Covers'] = mapping_422_rally
            else:
                print("  [WARN] CT000422 'Touring pack Rally' non trovato in Mappatura Rally")

            # ------------------------------------------------------------------
            # Top Case Pack V22E (Rally): stessi colori V21E (CT000478 identico),
            # ma ZCOL usa formato "Brushed Alu / X". Aggiungo le chiavi Rally
            # al mapping_478 gia' costruito da V21E, usando lo stesso strip-prefix.
            # ------------------------------------------------------------------
            if zcol_rally_descs and ct478_norm:
                _n_ct478_rally = 0
                for zcol_desc in zcol_rally_descs:
                    zcol_key_full = _fix_color_typos(zcol_desc.lower())
                    zcol_for_match = zcol_desc
                    if zcol_for_match.lower().startswith(_BRUSHED_ALU_PREFIX):
                        zcol_for_match = zcol_for_match[len(_BRUSHED_ALU_PREFIX):]
                    zcol_key_plain = _fix_color_typos(zcol_for_match.lower())
                    zcol_norm_478  = _normalize_color_desc_for_match(zcol_for_match)
                    if zcol_norm_478 in ct478_norm:
                        ct_val = ct478_norm[zcol_norm_478]
                        mapping_478[zcol_key_full]  = ct_val  # "brushed alu / x"
                        mapping_478[zcol_key_plain] = ct_val  # "x" (plain)
                        _n_ct478_rally += 1
                # Aggiorna il result con il mapping_478 esteso
                result['Top Case cover (color variant)'] = mapping_478
                print(f"  Mapping CT000478 Rally (Top Case V22E):  "
                      f"{_n_ct478_rally} ZCOL Rally aggiunti -> CT478")

    return result


def apply_pack_colors_to_wide_df(df_wide: pd.DataFrame, pack_color_mapping: dict) -> pd.DataFrame:
    """
    Sovrascrive i valori delle colonne pack esistenti con il colore derivato da FAIRING COLOR.

    Logica:
    - Se pack == "Yes" E FAIRING COLOR ha un match -> pack = colore specifico
    - Se pack == "Yes" ma FAIRING COLOR non mappa -> ANOMALIA, pack resta "Yes"
      (la mappatura ZCOL->pack color deve coprire tutti gli ZCOL attivi con il pack)
    - "No" rimane invariato

    INVARIANTE: dopo questa funzione, ogni pack attivo deve avere un colore specifico.
    Eventuali "Yes" residui indicano una mappatura incompleta e vengono segnalati.

    Il lookup usa FAIRING COLOR normalizzato (lowercase + typo-corrected) come chiave,
    coerente con come sono costruite le chiavi in build_zcol_pack_color_mapping.

    Args:
        df_wide: DataFrame wide output di build_wide_dataframe
        pack_color_mapping: output di build_zcol_pack_color_mapping

    Returns:
        df_wide con colonne pack sostituite in-place con i valori colore
    """
    if not pack_color_mapping:
        return df_wide

    fairing_col = 'FAIRING COLOR'
    if fairing_col not in df_wide.columns:
        print(f"  [WARN] Colonna '{fairing_col}' non trovata in df_wide -> skip pack colors")
        return df_wide

    # Canonical fairing values: lowercase + typo-corrected (matching lookup keys)
    # I valori in df_wide hanno il formato "FAIRING COLOR - Arctic White gloss"
    # (prefisso = nome caratteristica + " - "), serve solo la parte dopo " - ".
    _prefix = fairing_col.lower() + ' - '  # "fairing color - "
    fairing_canonical = (
        df_wide[fairing_col]
        .astype(str)
        .str.lower()
        .str.replace(_prefix, '', n=1)
        .apply(_fix_color_typos)
    )

    # Case-insensitive column lookup
    cols_lower = {c.lower(): c for c in df_wide.columns}

    # Mappa: (ct_key, pack_col_lower, model_ids_include)
    # model_ids_include: set di Model ID (lowercase) su cui applicare la derivazione.
    #   V21E (MSV4/MSV4S/MSV4PP): color variant per Touring Pack e Top Case
    #   V22E (MSV4RI Rally):       color variant per Touring Pack Rally e Top Case
    #   V24T (MSV4RS):             pack fissi (RS Livery) senza derivazione da ZCOL -> escluso
    pack_col_map = [
        ('Panniers covers (color variant)', 'touring pack',       {'msv4', 'msv4s', 'msv4pp'}),
        ('Top Case cover (color variant)',  'top case pack',      {'msv4', 'msv4s', 'msv4pp', 'msv4ri'}),
        ('Panniers Covers',                 'touring pack rally', {'msv4ri'}),
    ]

    model_id_col = cols_lower.get('model', None)

    modified = []
    for ct_col_name, pack_col_lower, model_ids_include in pack_col_map:
        if ct_col_name not in pack_color_mapping:
            continue
        lookup = pack_color_mapping[ct_col_name]
        if not lookup:
            continue

        # Trova colonna pack effettiva (case-insensitive)
        pack_col_actual = cols_lower.get(pack_col_lower, None)
        if pack_col_actual is None:
            print(f"  [WARN] Colonna '{pack_col_lower}' non trovata in df_wide -> skip")
            continue

        pack_active = df_wide[pack_col_actual].astype(str).str.lower() == 'yes'

        # Filtra per Model ID: RS (MSV4RS) ha pack fissi, nessuna derivazione da ZCOL
        if model_id_col is not None:
            model_mask = df_wide[model_id_col].astype(str).str.lower().isin(model_ids_include)
            pack_active = pack_active & model_mask
        else:
            print(f"  [WARN] Colonna 'Model ID' non trovata -> filtro modello non applicato "
                  f"per '{pack_col_lower}'")

        # Vectorized color lookup: NaN dove FAIRING COLOR non ha match
        colors = fairing_canonical.map(lookup)

        # Sovrascrittura in-place: pack attivo + colore trovato -> colore specifico
        mask_color = pack_active & colors.notna()
        df_wide.loc[mask_color, pack_col_actual] = colors[mask_color]

        n_yes    = pack_active.sum()
        n_color  = mask_color.sum()
        n_nomap  = n_yes - n_color
        if n_nomap > 0:
            _unmapped_mask = pack_active & colors.isna()
            _unmapped_vals = fairing_canonical[_unmapped_mask].unique().tolist()
            print(f"  [WARN] [{pack_col_actual}]: {n_nomap}/{n_yes} pack attivi "
                  f"rimasti 'Yes' - ZCOL non mappati: {_unmapped_vals}")
        else:
            print(f"  [{pack_col_actual}]: {n_yes} pack attivi -> tutti con colore specifico")
        modified.append(pack_col_actual)

    if modified:
        print(f"  Colonne pack aggiornate: {modified}")
    else:
        print("  Nessuna colonna pack modificata.")

    return df_wide


# ============================================================================
# FUNZIONE MAIN - ENTRY POINT
# ============================================================================

def main_simulator(n_runs: int, start_month_offset: int = 0) -> pd.DataFrame:
    """
    Funzione principale che orchestra l'intera simulazione Monte Carlo.

    FLUSSO:
    1. Carica forecast domanda mensile
    2. Calcola quanti mesi servono (basato su max lead time)
    3. Carica Take Rate e parametri Dirichlet
    4. Carica regole esclusione
    5. Esegue simulazione per ogni mese
    6. Verifica conservazione domanda
    7. Costruisce DataFrame output

    Args:
        n_runs: Numero run Monte Carlo (più run = statistiche più stabili)
                Consigliato: 200-1000 per produzione, 50-100 per test
        start_month_offset: Quanti mesi saltare nel file forecast
                           0 = primo mese, 10 = undicesimo mese, etc.

    Returns:
        DataFrame in formato wide pronto per SKU matching
    """
    print("="*60)
    print("MONTE CARLO SIMULATOR - VERSIONE V2")
    print("="*60)
    print("Ottimizzazioni attive:")
    print("  - Pre-calcolo pesi Dirichlet")
    print("  - Batch model sampling")
    print("  - Int16 per risparmio memoria 75%")
    print("  - Single Dirichlet per run")

    # Crea configurazione
    config = SimulationConfig(
        n_runs=n_runs,
        random_seed=42,
        start_month_offset=start_month_offset,
        use_int16=True
    )

    print(f"\nConfigurazione:")
    print(f"  - Run: {config.n_runs}")
    print(f"  - Seed: {config.random_seed}")
    print(f"  - Offset mese: {config.start_month_offset}")

    # =====================================================================
    # STEP 1: Carica forecast domanda
    # =====================================================================
    print("\n1. Caricamento forecast...")
    forecast_values, month_names = load_monthly_forecast(
        '.\\Input\\Total_demand.xlsx',
        start_offset=config.start_month_offset
    )

    # =====================================================================
    # STEP 2: Calcola mesi necessari (basato su lead time max)
    # =====================================================================
    print("\n2. Calcolo mesi necessari...")
    n_months_needed = calculate_n_months_needed()

    if n_months_needed > len(forecast_values):
        n_months_needed = len(forecast_values)

    # Stima preventiva consumo memoria per il successivo SKU matching
    # (shape: n_configs × n_months × n_runs × 2 byte int16)
    # Usa stima conservativa: n_configs ≈ 20.000, poi aggiorna dopo simulazione
    estimated_cols = n_months_needed * config.n_runs
    estimated_gb   = 20_000 * estimated_cols * 2 / (1024 ** 3)
    print(f"\n  [Diagnostica] Mesi da simulare : {n_months_needed}")
    print(f"  [Diagnostica] Colonne run stimat: {estimated_cols:,}")
    print(f"  [Diagnostica] RAM stimata (worst): ~{estimated_gb:.1f} GB")
    if estimated_gb > 4.0:
        print(f"\n  *** ATTENZIONE ***")
        print(f"  La simulazione di {n_months_needed} mesi × {config.n_runs} run")
        print(f"  richiederà circa {estimated_gb:.1f} GB di RAM (stima conservativa).")
        print(f"  Se si verifica un errore OOM, verificare:")
        print(f"  1. Valori anomali nella colonna lead time di 'quantity e residual.xlsx'")
        print(f"  2. Struttura del file 'Total_demand.xlsx' (quanti mesi contiene)")
        print(f"  3. Ridurre il numero di run (attuale: {config.n_runs})")

    # =====================================================================
    # STEP 3: Carica Take Rate e parametri Dirichlet
    # =====================================================================
    print("\n3. Caricamento TR...")

    # TR globali (usati come fallback se non ci sono TR mensili)
    model_mix, characteristics = parse_tr_file('.\\Input\\TR TOTALV21E - Copia.xlsx')
    print(f"  [Globale] Modelli: {len(model_mix.models)}")
    print(f"  [Globale] Caratteristiche: {len(characteristics)}")

    # TR mensili (opzionale): se il file esiste li carica, altrimenti usa globali
    monthly_tr = None
    monthly_tr_path = '.\\Input\\TR TOTALV21E - Copia.xlsx'
    try:
        monthly_tr = load_monthly_tr(monthly_tr_path)
        print(f"  [Mensile] TR mensili caricati: {len(monthly_tr)} mesi -> {list(monthly_tr.keys())}")
    except FileNotFoundError:
        print(f"  [Mensile] File '{monthly_tr_path}' non trovato -> uso TR globali per tutti i mesi")
    except ValueError as e:
        # Nessun foglio TR_Optional trovato nel file -> il file non ha ancora TR mensili
        print(f"  [Mensile] Nessun foglio TR_Optional nel file -> uso TR globali per tutti i mesi")
    except Exception as e:
        print(f"  [Mensile] WARN: impossibile caricare TR mensili ({e}) -> uso TR globali")

    # =====================================================================
    # STEP 4: Carica regole esclusione
    # =====================================================================
    print("\n4. Caricamento esclusioni...")
    exclusions = parse_exclusions(str(_MODULE_DIR / 'Input' / 'esclusioni.xlsx'))
    if not exclusions:
        print("  [WARN] Nessuna regola di esclusione caricata — conflitti tra optional NON risolti")

    # =====================================================================
    # STEP 5: Esegui simulazioni mensili
    # =====================================================================
    print("\n5. Esecuzione simulazioni...")
    results_by_month = run_monthly_simulations_optimized(
        forecast_values=forecast_values,
        month_names=month_names,
        n_months=n_months_needed,
        model_mix=model_mix,
        characteristics=characteristics,
        exclusions=exclusions,
        config=config,
        monthly_tr=monthly_tr     # None -> usa TR globali (backward compatible)
    )

    # =====================================================================
    # STEP 6: Verifica conservazione domanda
    # =====================================================================
    print("\n6. Verifica...")
    verify_demand_conservation(
        results_by_month=results_by_month,
        n_runs=config.n_runs,
        forecast_values=forecast_values[:n_months_needed]
    )

    # =====================================================================
    # STEP 7: Costruisci DataFrame output (senza salvare su file)
    # =====================================================================
    print("\n7. Costruzione DataFrame...")
    df_wide = build_wide_dataframe(
        results_by_month=results_by_month,
        month_names=month_names,
        n_runs=config.n_runs,
        save_to_file=False  # Non salva su disco: passa direttamente in memoria
    )

    # =====================================================================
    # STEP 7b: Derivazione colori pack da FAIRING COLOR (Touring/Top Case)
    # =====================================================================
    print("\n7b. Derivazione colori pack (Touring Pack / Top Case Pack / Touring Pack Rally)...")
    try:
        _pack_color_mapping = build_zcol_pack_color_mapping(
            str(_MODULE_DIR / 'Input' / 'Mappatura - Copia.xlsx'),
            rally_mappatura_path=str(_MODULE_DIR / 'Input' / 'Mappatura Rally.xlsx'),
        )
        if _pack_color_mapping:
            df_wide = apply_pack_colors_to_wide_df(df_wide, _pack_color_mapping)
        else:
            print("  [WARN] Mapping ZCOL->pack color vuoto, colori pack non derivati")
    except Exception as e:
        print(f"  [WARN] Errore derivazione colori pack: {e} -> colori pack non aggiunti")

    # Pulizia memoria
    gc.collect()

    print(f"\n{'='*60}")
    print("SIMULAZIONE COMPLETATA!")
    print(f"{'='*60}")

    n_total = len(set().union(*[r.keys() for r in results_by_month.values()]))
    print(f"  Mesi: {n_months_needed}")
    print(f"  Run: {config.n_runs}")
    print(f"  Combinazioni: {n_total}")

    return df_wide


# ============================================================================
# RESCALE ALPHAS — Affidabilita TR (condiviso con mc_simulator_rolling)
# ============================================================================

def rescale_alphas(model_mix, characteristics, monthly_tr,
                   aff_model: str, aff_optional: str,
                   aff_map: dict = None):
    """
    Applica moltiplicatori di confidenza TR (VSS) agli alpha di model_mix e
    characteristics. Funzione condivisa da SS classica e rolling.

    Args:
        model_mix:        oggetto ModelMix (con .alphas numpy array)
        characteristics:  dict {nome: CharacteristicGroup}
        monthly_tr:       dict mese -> (ModelMix, characteristics) o None
        aff_model:        chiave affidabilita model mix (es. 'BASSA')
        aff_optional:     chiave affidabilita optional (es. 'BASSA')
        aff_map:          dizionario {chiave: VSS float}; se None carica da
                          tr_parser_V2._AFFIDABILITA_MAP con fallback hardcoded.

    Returns:
        (model_mix_scaled, characteristics_scaled, monthly_tr_scaled)
    """
    import copy
    if aff_map is None:
        try:
            from tr_parser_V2 import _AFFIDABILITA_MAP
            aff_map = {k: float(v) for k, v in _AFFIDABILITA_MAP.items()}
        except Exception:
            aff_map = {"NULLA": 50.0, "BASSA": 200.0, "MEDIA": 1000.0, "ALTA": 5000.0}

    valid_keys = list(aff_map.keys())
    if aff_model not in aff_map:
        raise ValueError(
            f"Livello affidabilità modello '{aff_model}' non valido. "
            f"Valori ammessi: {valid_keys}"
        )
    if aff_optional not in aff_map:
        raise ValueError(
            f"Livello affidabilità optional '{aff_optional}' non valido. "
            f"Valori ammessi: {valid_keys}"
        )
    scale_m = aff_map[aff_model]
    scale_o = aff_map[aff_optional]

    mm   = copy.deepcopy(model_mix)
    char = copy.deepcopy(characteristics)

    mm.alphas = mm.alphas * scale_m
    for cg in char.values():
        for k in list(cg.alpha.keys()):
            cg.alpha[k] = cg.alpha[k] * scale_o

    mtr = None
    if monthly_tr is not None:
        mtr = {}
        for month_label, (mm_m, ch_m) in monthly_tr.items():
            mm_m2 = copy.deepcopy(mm_m)
            ch_m2 = copy.deepcopy(ch_m)
            mm_m2.alphas = mm_m2.alphas * scale_m
            for cg in ch_m2.values():
                for k in list(cg.alpha.keys()):
                    cg.alpha[k] = cg.alpha[k] * scale_o
            mtr[month_label] = (mm_m2, ch_m2)

    return mm, char, mtr


# ============================================================================
# ENTRY POINT PER ESECUZIONE DIRETTA
# ============================================================================

if __name__ == "__main__":
    # Esecuzione standalone con 1000 run
    output = main_simulator(n_runs=25)
    print(f"\n[OK] Output pronto")

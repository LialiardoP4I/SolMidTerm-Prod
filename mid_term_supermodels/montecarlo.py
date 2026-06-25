# -*- coding: utf-8 -*-
"""Simulatore Monte Carlo della domanda.

In parole semplici: genera molti scenari di vendita realistici (migliaia) per
stimare l'incertezza della domanda di ogni configurazione di moto. Per ogni mese
estrae le configurazioni rispettando i Take Rate, risolve gli optional
incompatibili e produce la tabella usata poi per abbinare i componenti.

(dettaglio tecnico) Monte Carlo Simulator per la pipeline mid-term Safety Stock.

Esegue simulazioni Monte Carlo per generare combinazioni di configurazioni
moto (Ducati Multistrada V4) basate su distribuzioni Dirichlet. L'output
viene utilizzato a valle per il calcolo dei safety stock dei componenti.

Per ogni mese e per ogni run:
- campiona le probabilità delle opzioni dalla distribuzione Dirichlet
- genera N configurazioni moto (N = domanda mensile)
- risolve eventuali conflitti tra opzioni incompatibili

I risultati vengono aggregati in un DataFrame wide pronto per lo SKU matching.
"""

import numpy as np
import pandas as pd
from mid_term_supermodels.tr import (
    parse_tr_file, load_monthly_tr, get_tr_for_month,
    ModelMix, CharacteristicGroup, canonical_model_name,
)
from mid_term_supermodels._logging import get_logger
from mid_term_supermodels.exceptions import SimulationError
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
from pathlib import Path
import re
import math
import gc
import time

log = get_logger("montecarlo")

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
                # Accumula: chiavi duplicate su piu righe NON si sovrascrivono.
                exclusions.setdefault(option, []).extend(incompatible_options)

        log.info("[OK] Esclusioni caricate: %d regole", len(exclusions))
        return exclusions

    except Exception as e:
        log.warning("Impossibile caricare esclusioni: %s", e)
        return {}


def _parse_selector(token):
    """Token -> selettore (gruppo, valore_or_None).
    'GRUPPO:VALORE' -> valore-livello; 'GRUPPO' -> gruppo-livello (None)."""
    token = str(token).strip()
    if ':' in token:
        group, value = token.split(':', 1)
        return (group.strip(), value.strip())
    return (token, None)


def _selector_matches(selector, config):
    """True se il selettore matcha il config."""
    group, value = selector
    if group not in config:
        return False
    cfg_val = str(config[group]).strip()
    if value is None:
        # gruppo-livello (bundle): match se selezionato (!= NOT)
        return 'NOT' not in cfg_val.upper()
    # valore-livello: match su valore esatto
    return cfg_val == value


def find_conflicts(config: Dict, exclusions: Dict[str, List[str]]) -> List[Tuple]:
    """
    Identifica i conflitti in una configurazione moto.

    Gestisce entrambe le forme di esclusione:
    - gruppo-livello (bundle binario Yes/Not): chiave = nome gruppo;
    - valore-livello: chiave/incompatibile = 'GRUPPO:VALORE'.

    Un conflitto scatta quando ENTRAMBI i selettori di una coppia vietata
    matchano il config.

    Args:
        config: Dizionario {caratteristica: valore_selezionato}
        exclusions: Regole di esclusione caricate da parse_exclusions()

    Returns:
        Lista di conflitti, ogni conflitto è ((char1,val1), (char2,val2))
    """
    conflicts = []
    seen = set()

    for key_token, incompat_list in exclusions.items():
        sel_a = _parse_selector(key_token)
        if not _selector_matches(sel_a, config):
            continue
        for inc_token in incompat_list:
            sel_b = _parse_selector(inc_token)
            if sel_a[0] == sel_b[0]:
                continue  # stesso gruppo: salta
            if not _selector_matches(sel_b, config):
                continue
            pair = tuple(sorted([sel_a[0], sel_b[0]]))
            if pair in seen:
                continue
            seen.add(pair)
            conflicts.append(((sel_a[0], config[sel_a[0]]),
                              (sel_b[0], config[sel_b[0]])))

    return conflicts


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
                                      max_iterations=None) -> Tuple[Dict, int]:
    """
    Genera una configurazione VALIDA tramite rejection sampling.

    Processo: genera una configurazione; se viola un'esclusione la scarta e ne
    rigenera una nuova da zero (stessi pesi Dirichlet della run); ripete finche
    la configurazione e valida. Nessun repair.

    Args:
        model: Modello moto
        characteristics: Dizionario caratteristiche (non piu usato internamente,
            mantenuto per compatibilita di chiamata)
        char_names: Lista nomi caratteristiche
        exclusions: Regole esclusione
        char_weights: Pesi pre-computati
        max_iterations: None (default) = nessun cap funzionale: rigenera fino a
            config valida, protetto da una valvola SAFETY=100000 che emette un
            warning (anti-hang se i dati fossero infattibili, P=0). Un int finito
            cappa a quel numero di estrazioni (usato dai test).

    Returns:
        Tupla (configurazione_valida, numero_configurazioni_scartate).
        Con cap finito l'ultima config puo restare invalida.
    """
    # Rejection sampling: genera e rigenera finche la config e valida.
    # Verificato (su dati ALL) che esiste sempre una combinazione valida
    # (P(accettazione) > 0); SAFETY e' solo una valvola anti-hang.
    SAFETY = 100_000
    limit = max_iterations if max_iterations is not None else SAFETY

    rejected = 0
    config = generate_configuration_fast(model, char_names, char_weights)

    for _ in range(limit):
        if not find_conflicts(config, exclusions):
            return config, rejected
        rejected += 1
        config = generate_configuration_fast(model, char_names, char_weights)

    if max_iterations is None:
        log.warning("generate_valid_configuration_fast: SAFETY %d raggiunto per "
                    "modello %s -> possibile infattibilita totale (dati?)", SAFETY, model)
    return config, rejected


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
    log.info("Caricamento forecast mensile da %s", filepath)

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
        log.warning(
            "Trovati %d mesi nel file previsioni (attesi tipicamente <=60). "
            "Possibile elevato consumo di RAM. Verificare struttura file '%s' "
            "(riga 1=nomi mesi, riga 2=valori). Primi 5 mesi: %s",
            len(month_names), filepath, month_names[:5]
        )

    log.info("Mesi disponibili: %d", len(forecast_values))
    log.info("Mese iniziale: %s", month_names[0] if month_names else 'N/A')

    return forecast_values, month_names


def calculate_n_months_needed() -> int:
    """
    Calcola quanti mesi simulare basandosi sul lead time massimo dei componenti.
    Legge da Gates MTSV4.xlsx (Dati Logistica) — col. 'Gates (Weeks)'.
    """
    from pathlib import Path
    from mid_term_supermodels.config import shared_logistica_dir

    gates_file = shared_logistica_dir() / 'Gates MTSV4.xlsx'

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
        log.warning(
            "%d riga/righe con gate weeks anomalo (<=0 o >%d mesi) — IGNORATI. "
            "Valori anomali: %s",
            len(anomalous_vals), MAX_PLAUSIBLE_MONTHS, list(anomalous_vals.values[:10])
        )

    if len(valid_values) == 0:
        raise ValueError(
            "Nessun valore valido in Gates (Weeks) di Gates MTSV4.xlsx.\n"
            "Verificare che il file contenga dati numerici validi."
        )

    max_lead_time = valid_values.max()
    n_months = math.ceil(max_lead_time)

    log.info(
        "Max lead time (valori validi <=%dm): %.1f mesi -> %d mesi da simulare",
        MAX_PLAUSIBLE_MONTHS, max_lead_time, n_months
    )

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
    log.info("Mese %d: %s (domanda: %.0f)", month_index + 1, month_name, total_demand)

    # Imposta seed per riproducibilità (diverso per ogni mese e per ogni offset rolling).
    # Moltiplica offset per 1000 per evitare sovrapposizioni tra iterazioni rolling.
    np.random.seed(config.random_seed + config.start_month_offset * 1000 + month_index)

    char_names = list(month_characteristics.keys())

    # Canonicalizza i nomi modello (ZWS/invisibili/casing) ed escludi "ALTRO"
    # e celle vuote. Il filtraggio in un unico passaggio zip-pair garantisce
    # che len(valid_models) == len(valid_alphas) anche se ci sono piu' righe
    # ALTRO o nomi vuoti, eliminando il rischio di mismatch in np.delete.
    _pairs = [
        (canonical_model_name(m), a)
        for m, a in zip(month_model_mix.models, month_model_mix.alphas)
        if canonical_model_name(m) not in ('', 'ALTRO')
    ]
    valid_models = [p[0] for p in _pairs]
    valid_alphas = np.array([p[1] for p in _pairs], dtype=float)

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
        # Progress indicator (debug-only, log.debug evita di sporcare l'output di default)
        if run % 200 == 0 and run > 0:
            log.debug("Run %d/%d", run, config.n_runs)

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
            config_obj, num_rejected = generate_valid_configuration_fast(
                model, month_characteristics, char_names, exclusions,
                model_char_weights[model],
                None  # cap rimosso: rejection sampling fino a config valida
            )

            if num_rejected > 0:
                stats_configs_with_conflicts += 1

            # Crea chiave univoca per questa configurazione
            config_key = tuple(sorted(config_obj.items()))
            run_combination_counts[config_key] += 1

        # Salva conteggi di questa run
        for config_key, count in run_combination_counts.items():
            combination_demands_by_run[config_key][run] = count

    log.info(
        "Run %d/%d completate - %d combinazioni uniche",
        config.n_runs, config.n_runs, len(combination_demands_by_run)
    )

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
    log.info(
        "Simulazioni multi-mese: %d mesi, %d run/mese, dtype=%s",
        n_months, config.n_runs, 'int16' if config.use_int16 else 'int32'
    )

    # Verifica disponibilità mesi
    if n_months > len(forecast_values):
        log.warning(
            "Richiesti %d mesi ma disponibili solo %d", n_months, len(forecast_values)
        )
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
        log.info("Tempo mese: %.1fs", month_elapsed)

        # Garbage collection periodico per liberare memoria
        if month_idx % 2 == 0:
            gc.collect()

    total_elapsed = time.time() - start_time
    log.info(
        "[OK] Simulazione completata in %.1fs (%.1fmin)",
        total_elapsed, total_elapsed / 60
    )

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
    log.info("Verifica conservazione domanda")

    all_correct = True
    for month_idx, month_results in results_by_month.items():
        if month_idx >= len(forecast_values):
            continue

        expected_demand = forecast_values[month_idx]

        # Verifica sul primo run (rappresentativo)
        run_total = sum(int(demands[0]) for demands in month_results.values())

        if abs(run_total - expected_demand) > 1:  # Tolleranza di 1 unità
            all_correct = False
            log.warning(
                "Mese %d: %.0f / %.0f MISMATCH",
                month_idx + 1, run_total, expected_demand
            )
        else:
            log.info("Mese %d: [OK] %.0f", month_idx + 1, expected_demand)

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
    log.info("Costruzione DataFrame wide (in memoria)")

    # Raccogli tutte le configurazioni uniche da tutti i mesi
    all_configs = set()
    for month_results in results_by_month.values():
        all_configs.update(month_results.keys())

    all_configs = sorted(all_configs)
    log.info("Combinazioni totali: %d", len(all_configs))

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
    log.info("Aggiungendo colonne run...")

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
            log.debug("Mese %d/%d", list_idx + 1, n_months_to_build)

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

    log.info("Righe: %d", len(df))
    log.info("Colonne: %d", len(df.columns))

    # Stima memoria
    mem_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
    log.info("Memoria DataFrame: %.1f MB", mem_mb)

    # Salva su file SOLO se esplicitamente richiesto
    if save_to_file:
        log.info("Salvataggio su %s...", filename)
        df.to_excel(filename, index=False)
        log.info("[OK] File salvato: %s", filename)

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
    In parole semplici: costruisce la tabella "colore della carena -> colore del
    pack" (es. il colore delle borse segue il colore della moto). Serve perche'
    alcuni pack esistono in piu' varianti di colore.

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
        log.warning("Impossibile leggere mappatura '%s': %s", mappatura_path, e)
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
        log.warning("ZCOL non trovato in mappatura -> colori pack non derivati")
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
            log.info("ZCOL '%s' -> CT000422: nessun match", zcol_desc)

        if zcol_norm in ct478_norm:
            mapping_478[zcol_key] = ct478_norm[zcol_norm]
        else:
            log.info("ZCOL '%s' -> CT000478: nessun match", zcol_desc)

    # Aggiungi mapping espliciti per casi combinati che non matchano per normalizzazione.
    # "Ducati Red" e "Pikes Peak Livery" vengono mappati come valori SEPARATI (non al
    # combined "(Ducati Red+PP)"), in modo che la simulazione produca righe distinte
    # per ciascun colore e il matching SKU funzioni correttamente.
    for ct_mapping in [mapping_422, mapping_478]:
        ct_mapping['ducati red']          = 'Ducati Red'
        ct_mapping['pikes peak livery']   = 'Pikes Peak Livery'

    log.info("Mapping CT000422 (Panniers V21E): %d ZCOL -> colore", len(mapping_422))
    log.info("Mapping CT000478 (Top Case):      %d ZCOL -> colore", len(mapping_478))

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
            log.warning(
                "Impossibile leggere mappatura Rally '%s': %s", rally_mappatura_path, e
            )
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

            zcol_rally_descs = []  # init: evita UnboundLocalError nel blocco Top Case Rally se manca ZCOL/bundle Rally
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
                    log.info("ZCOL Rally trovato: %d colori", len(zcol_rally_descs))
                else:
                    # Fallback: ogni CT422 Rally desc mappa su se stessa
                    zcol_rally_descs = ct422_rally_descs
                    log.info(
                        "ZCOL Rally non trovato -> uso CT422 Rally come proxy (%d colori)",
                        len(zcol_rally_descs)
                    )

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
                        log.info(
                            "ZCOL Rally '%s' -> CT000422 Rally: nessun match", zcol_desc
                        )

                log.info(
                    "Mapping CT000422 Rally (Panniers V22E): %d ZCOL -> colore",
                    len(mapping_422_rally)
                )
                result['Panniers Covers'] = mapping_422_rally
            else:
                log.warning(
                    "CT000422 'Touring pack Rally' non trovato in Mappatura Rally"
                )

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
                log.info(
                    "Mapping CT000478 Rally (Top Case V22E): %d ZCOL Rally aggiunti -> CT478",
                    _n_ct478_rally
                )

    return result


def apply_pack_colors_to_wide_df(df_wide: pd.DataFrame, pack_color_mapping: dict) -> pd.DataFrame:
    """
    In parole semplici: nelle configurazioni simulate sostituisce il generico
    "Yes" di un pack con il suo colore specifico, ricavato dal colore della
    carena, cosi' il pack potra' essere abbinato allo SKU del colore giusto.

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
        log.warning(
            "Colonna '%s' non trovata in df_wide -> skip pack colors", fairing_col
        )
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
            log.warning(
                "Colonna '%s' non trovata in df_wide -> skip", pack_col_lower
            )
            continue

        pack_active = df_wide[pack_col_actual].astype(str).str.lower() == 'yes'

        # Filtra per Model ID: RS (MSV4RS) ha pack fissi, nessuna derivazione da ZCOL.
        # Canonicalizza il valore prima del confronto per essere resilienti a
        # eventuali invisibili / casing residui nella colonna Model.
        if model_id_col is not None:
            model_canon = df_wide[model_id_col].map(canonical_model_name).str.lower()
            model_mask = model_canon.isin(model_ids_include)
            pack_active = pack_active & model_mask
        else:
            log.warning(
                "Colonna 'Model ID' non trovata -> filtro modello non applicato per '%s'",
                pack_col_lower
            )

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
            log.warning(
                "[%s]: %d/%d pack attivi rimasti 'Yes' - ZCOL non mappati: %s",
                pack_col_actual, n_nomap, n_yes, _unmapped_vals
            )
        else:
            log.info(
                "[%s]: %d pack attivi -> tutti con colore specifico",
                pack_col_actual, n_yes
            )
        modified.append(pack_col_actual)

    if modified:
        log.info("Colonne pack aggiornate: %s", modified)
    else:
        log.info("Nessuna colonna pack modificata.")

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
                          mid_term.tr._AFFIDABILITA_MAP con fallback hardcoded.

    Returns:
        (model_mix_scaled, characteristics_scaled, monthly_tr_scaled)
    """
    import copy
    if aff_map is None:
        try:
            from mid_term_supermodels.tr import _AFFIDABILITA_MAP
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

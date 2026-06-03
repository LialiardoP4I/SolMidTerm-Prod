"""TR File Parser: caratteristiche e mix modelli da Excel TR mensili."""
import json
import os
import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from datetime import datetime

from mid_term._logging import get_logger

log = get_logger("tr")

# ---------------------------------------------------------------------------
# Configurazione affidabilità -> moltiplicatore Dirichlet
# ---------------------------------------------------------------------------
_AFFIDABILITA_CONFIG_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'Input', 'affidabilita_config.json',
)
_AFFIDABILITA_DEFAULT = {'BASSA': 10, 'MEDIA': 100, 'ALTA': 1000}

def _load_affidabilita_config(path: str = _AFFIDABILITA_CONFIG_PATH) -> Dict[str, float]:
    """
    Carica la mappa Affidabilità TR -> moltiplicatore Dirichlet dal JSON esterno.
    Fallback ai valori di default (BASSA=10, MEDIA=100, ALTA=1000) se il file
    non esiste o non è leggibile.
    """
    try:
        with open(path, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        # Filtra la chiave di commento e normalizza in uppercase
        return {k.upper(): float(v) for k, v in raw.items() if not k.startswith('_')}
    except Exception as e:
        log.warning("affidabilita_config.json non caricato (%s) -> uso default", e)
        return dict(_AFFIDABILITA_DEFAULT)

# Cache della config (caricata una volta sola)
_AFFIDABILITA_MAP: Dict[str, float] = _load_affidabilita_config()


# ---------------------------------------------------------------------------
# Canonicalizzazione nomi modello
# ---------------------------------------------------------------------------
# Caratteri Unicode invisibili che Excel puo' iniettare nei nomi (copy/paste
# da PDF, Word, browser, export SAP). Rimossi prima di qualsiasi confronto.
#   U+200B ZERO WIDTH SPACE       U+200C ZERO WIDTH NON-JOINER
#   U+200D ZERO WIDTH JOINER      U+FEFF BYTE ORDER MARK / ZWNBSP
#   U+00A0 NO-BREAK SPACE         U+202F NARROW NO-BREAK SPACE
#   U+2060 WORD JOINER            U+180E MONGOLIAN VOWEL SEPARATOR
_INVISIBLE_CHARS = ('​', '‌', '‍', '﻿',
                    ' ', ' ', '⁠', '᠎')


def canonical_model_name(raw) -> str:
    """
    Canonicalizza un nome modello in forma confrontabile end-to-end.

    Operazioni (in ordine):
      1. Guard NaN/None/cella vuota -> stringa vuota ''.
      2. str() per gestire input non-stringa senza creare la stringa 'nan'.
      3. Rimozione di tutti i caratteri Unicode invisibili noti.
      4. strip() + upper().

    NON modifica spazi interni, underscore o lunghezza del nome: i nomi
    reali (MSV4, MSV4S, MSV4PP, MSV4RS, MSV4RI, PANV47G, PANV4S7G,
    PAN7GSP, DVLV4, DVLV4S, ...) vengono preservati esattamente, solo
    ripuliti da invisibili e casing inconsistente.

    Casi limite:
      - raw is None     -> ''
      - raw is NaN      -> ''
      - raw == 'MSV4​' -> 'MSV4'
      - raw == ' msv4s '    -> 'MSV4S'

    Il chiamante deve filtrare le stringhe vuote prima di usarle come
    chiave modello.
    """
    if raw is None:
        return ''
    # Pandas NaN: float('nan') != float('nan'), test con isinstance + math.isnan
    if isinstance(raw, float):
        try:
            import math
            if math.isnan(raw):
                return ''
        except Exception:
            pass
    s = str(raw)
    for ch in _INVISIBLE_CHARS:
        if ch in s:
            s = s.replace(ch, '')
    return s.strip().upper()


@dataclass
class ModelMix:
    """Model selection distribution (Level 0 Dirichlet)"""
    models: List[str]
    take_rates: np.ndarray
    alphas: np.ndarray
    
    def validate(self):
        total = np.sum(self.take_rates)
        if not (0.99 <= total <= 1.01):
            log.warning("Model TR sums to %.3f", total)

@dataclass
class CharacteristicGroup:
    """Caratteristica Dirichlet con N valori (TR/alpha indicizzati per modello).

    Attributes:
        name:   Nome della caratteristica.
        values: Lista valori possibili.
        tr:     Dict {model_name: np.ndarray} con i Take Rate per modello.
        alpha:  Dict {model_name: np.ndarray} con i parametri Dirichlet per modello.
    """
    name: str
    values: List[str]
    tr: Dict[str, np.ndarray]      # { model_name: take_rate_array }
    alpha: Dict[str, np.ndarray]   # { model_name: alpha_array }

    def validate(self):
        """Check that TR sums to ~1.0 for each model"""
        for model_name, tr_arr in self.tr.items():
            total = np.sum(tr_arr)
            if not (0.99 <= total <= 1.01):
                log.warning("%s [%s] sums to %.3f", self.name, model_name, total)

    @staticmethod
    def _norm(s: str) -> str:
        """
        Normalizza nome modello per il lookup interno.

        Delega a canonical_model_name + rimozione spazi/underscore residui:
        questa funzione conserva la retrocompatibilita' del matching fuzzy
        in _lookup (es. legacy 'MSV RS' -> 'MSVRS') senza alterare la forma
        canonica usata dal resto della pipeline.
        """
        return canonical_model_name(s).replace(' ', '').replace('_', '')

    def get_tr(self, model_name: str) -> np.ndarray:
        """
        Restituisce i Take Rate per il modello richiesto.

        Logica di ricerca (in ordine):
        1. Hit esatto (case insensitive, spazi/underscore normalizzati)
        2. Hit parziale: il nome modello contiene o è contenuto nella chiave
        3. Fallback: prima chiave disponibile

        Garantisce retrocompatibilità e robustezza a piccole variazioni
        nei nomi modello tra Excel e forecast.
        """
        return self._lookup(self.tr, model_name)

    def get_alpha(self, model_name: str) -> np.ndarray:
        """
        Restituisce i parametri alpha Dirichlet per il modello richiesto.
        Stessa logica di fallback di get_tr().
        """
        return self._lookup(self.alpha, model_name)

    def _lookup(self, d: Dict[str, np.ndarray], model_name: str) -> np.ndarray:
        """Ricerca fuzzy: esatto -> MSV4 family alias -> parziale -> primo disponibile."""
        if not d:
            return np.array([1.0])
        norm_target = self._norm(model_name)
        # 1. Hit esatto
        for k, v in d.items():
            if self._norm(k) == norm_target:
                return v
        # 2. MSV4-family normalization: 'MSVRS' -> prova anche 'MSV4RS', ecc.
        #    Gestisce il caso in cui TR_Model usa 'MSV RS' mentre TR_Optional ha 'MSV4RS'.
        if norm_target.startswith('MSV') and len(norm_target) > 3 and norm_target[3] != '4':
            alt_target = 'MSV4' + norm_target[3:]
            for k, v in d.items():
                if self._norm(k) == alt_target:
                    return v
        # 3. Hit parziale (uno contiene l'altro)
        for k, v in d.items():
            nk = self._norm(k)
            if norm_target in nk or nk in norm_target:
                return v
        # 4. Fallback: primo disponibile
        first_key = next(iter(d))
        log.warning(
            "CharGroup[%s]: modello '%s' non trovato, fallback su '%s'",
            self.name, model_name, first_key,
        )
        return d[first_key]
                
def _detect_header_row(filepath: str, sheet_name: str,
                       probe_col: str = 'Nome Opzione/Bundle',
                       max_check: int = 4) -> int:
    """
    Rileva automaticamente in quale riga del foglio si trova l'header.

    Legge le prime `max_check` righe con header=None e cerca `probe_col`
    (case-insensitive) come valore nella prima cella non-NaN di ciascuna riga.
    Ritorna l'indice della riga trovata (0-based), oppure 0 come default.

    Serve a gestire file TR che hanno una riga informativa extra in cima
    (es. "Valori presi dal file"), spostando l'header reale in riga 1.
    """
    try:
        df_raw = pd.read_excel(filepath, sheet_name=sheet_name,
                               header=None, nrows=max_check)
        probe_lower = probe_col.strip().lower()
        for i, row in df_raw.iterrows():
            # Controlla le prime 3 celle della riga (non solo la prima)
            for cell_val in row.iloc[:3]:
                if pd.isna(cell_val):
                    continue
                if probe_lower in str(cell_val).strip().lower():
                    return int(i)
    except Exception:
        pass
    return 0


def parse_tr_file(filepath: str) -> tuple:
    """
    Parse TR file into structured characteristics.

    Strategia: usa il foglio TR_Optional_* più vecchio disponibile come
    sorgente globale (fallback per i mesi non coperti dai TR mensili).
    Se esistono ancora i vecchi fogli statici 'TR-TOTV21E' / 'TR-Model'
    li usa direttamente per retrocompatibilità.

    Returns:
        model_mix: ModelMix
        characteristics: Dict[str, CharacteristicGroup]
    """
    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names

    # --- Retrocompatibilità: vecchi fogli statici ancora presenti ---
    if 'TR-TOTV21E' in sheets and 'TR-Model' in sheets:
        model_mix = parse_model_mix(filepath, sheet='TR-Model')
        _hdr = _detect_header_row(filepath, 'TR-TOTV21E')
        df = pd.read_excel(filepath, sheet_name='TR-TOTV21E', header=_hdr)
    else:
        # Nuovo formato: usa il primo mese disponibile (ordine cronologico)
        import re as _re
        opt_sheets = sorted(
            [s for s in sheets if s.lower().startswith('tr_optional_')],
            key=lambda s: _parse_month_str(s[len('tr_optional_'):])
        )
        mod_sheets = sorted(
            [s for s in sheets if s.lower().startswith('tr_model_')],
            key=lambda s: _parse_month_str(s[len('tr_model_'):])
        )

        if not opt_sheets:
            raise ValueError(
                f"Nessun foglio 'TR_Optional_*' o 'TR-TOTV21E' trovato in '{filepath}'."
            )

        first_opt = opt_sheets[0]
        first_mod = mod_sheets[0] if mod_sheets else None

        log.info("Uso '%s' come TR globale di riferimento", first_opt)
        if first_mod is None:
            raise ValueError(
                f"Nessun foglio 'TR_Model_*' trovato in '{filepath}'. "
                "Impossibile derivare model_mix."
            )
        model_mix = parse_model_mix(filepath, sheet=first_mod)
        _hdr = _detect_header_row(filepath, first_opt)
        df = pd.read_excel(filepath, sheet_name=first_opt, header=_hdr)

    characteristics = {}
    for char_name, group in df.groupby('Nome Opzione/Bundle'):
        if pd.isna(char_name):
            continue
        if char_name == 'BUNDLE':
            characteristics.update(_parse_bundles(group))
        else:
            try:
                mc = _extract_model_columns(group.columns)
            except ValueError:
                mc = {}
            if mc and _all_tr_zero(group, mc):
                log.info("[SKIP] Caratteristica '%s': TR=0 su tutti i modelli, ignorata", char_name)
                continue
            characteristics[char_name] = _parse_characteristic(char_name, group)

    return model_mix, characteristics


def parse_model_mix(filepath: str, sheet: str = 'TR-Model') -> ModelMix:
    """
    Parse Model mix dal foglio TR-Model (o TR_Model_*).

    Args:
        filepath: Percorso al file Excel
        sheet: Nome foglio (default 'TR-Model' per retrocompatibilità,
               può essere 'TR_Model_Gen26' ecc.)
    """
    _hdr = _detect_header_row(filepath, sheet,
                              probe_col='CONCAT Nome Opzione/Bundle')
    df = pd.read_excel(filepath, sheet_name=sheet, header=_hdr)

    raw_models = df['CONCAT Nome Opzione/Bundle - Caratteristica'].tolist()
    raw_trs    = df['Take Rate Model'].tolist()

    # Canonicalizza nomi modello e scarta righe vuote/NaN mantenendo
    # l'allineamento posizionale fra models e take_rates.
    pairs = [(canonical_model_name(m), tr)
             for m, tr in zip(raw_models, raw_trs)
             if canonical_model_name(m) != '']
    models     = [p[0] for p in pairs]
    take_rates = np.array([p[1] for p in pairs], dtype=float)

    # Alpha = TR direttamente (nessuna pre-moltiplicazione).
    # rescale_alphas in app_v3 applica alpha = TR × AFF_MAP[ui_choice]
    # senza alcuna normalizzazione intermedia.
    multipliers = np.ones(len(take_rates))

    alphas = take_rates * multipliers

    mix = ModelMix(models=models, take_rates=take_rates, alphas=alphas)
    mix.validate()
    return mix

def _extract_model_columns(df_columns) -> Dict[str, Dict[str, str]]:
    """Rileva dinamicamente le colonne TR per modello nel DataFrame.

    Supporta sia il formato nuovo ('MSV4', 'MSV4S', ...) sia il vecchio
    ('Take Rate BASE\\nCluster3', ...). Il nome della colonna in uppercase
    diventa la chiave del dizionario tr/alpha in CharacteristicGroup.
    Ritorna {model_suffix: {'tr_col': original_col_name}}.
    """
    import re as _re

    cols = [str(c) for c in df_columns]

    # Colonne strutturali da escludere (non sono colonne modello)
    EXCLUDED = {
        'NOME OPZIONE/BUNDLE',
        'NOME OPZIONE/BUNDE',                              # retrocompatibilità typo
        'CONCAT NOME OPZIONE/BUNDLE - CARATTERISTICA',
        'CONCAT NOME OPZIONE/BUNDE - CARATTERISTICA',      # retrocompatibilità typo
        'AFFIDABILITÀ TR', 'AFFIDABILITA TR', "AFFIDABILITA' TR",
    }
    alpha_pattern = _re.compile(r'^Parametro Dirichelet', _re.IGNORECASE)

    tr_map = {}  # { suffisso_upper: colonna_originale }

    # --- Formato vecchio: "Take Rate BASE\nCluster3" ---
    tr_pattern_old = _re.compile(r'^Take Rate (.+?)(?:\\n|\n)Cluster\d*$', _re.IGNORECASE)
    for col in cols:
        m = tr_pattern_old.match(col)
        if m:
            suffix = m.group(1).strip().upper()
            tr_map[suffix] = col

    if not tr_map:
        # --- Formato nuovo: ogni colonna non esclusa e non alpha è una colonna TR ---
        for col in cols:
            col_upper = col.strip().upper()
            if col_upper in EXCLUDED:
                continue
            if alpha_pattern.match(col):
                continue
            tr_map[col_upper] = col

    if not tr_map:
        raise ValueError(
            f"Nessuna colonna TR trovata nelle colonne: {cols}. "
            "Verifica i nomi colonna nel foglio TR_Optional o TR-TOTV21E."
        )

    return {suffix: {'tr_col': original} for suffix, original in tr_map.items()}


def _all_tr_zero(group: pd.DataFrame, model_cols: Dict[str, Dict[str, str]]) -> bool:
    """
    Restituisce True se la somma di tutti i TR su tutti i modelli è zero.
    Usato per scartare caratteristiche/bundle che non esistono sul prodotto.
    """
    for cols in model_cols.values():
        if group[cols['tr_col']].fillna(0).sum() > 0:
            return False
    return True


def _parse_characteristic(name: str, group: pd.DataFrame) -> CharacteristicGroup:
    """Parsea una caratteristica mutualmente esclusiva (rilevamento modelli dinamico).

    Le opzioni con TR=0 su tutti i modelli vengono rimosse da values e dagli
    array TR/alpha (non sarebbero mai generate dalla simulazione).
    """
    values = group['CONCAT Nome Opzione/Bundle - Caratteristica'].tolist()

    # Colonna affidabilità (una per riga, indipendente dal modello)
    affidabilita_col = next(
        (c for c in group.columns if c.strip().upper().startswith('AFFIDABILIT')),
        None
    )

    # Rileva modelli disponibili nelle colonne
    model_cols = _extract_model_columns(group.columns)

    tr_dict    = {}
    alpha_dict = {}

    for model_suffix, cols in model_cols.items():
        tr_arr = group[cols['tr_col']].fillna(0).values

        # Alpha = TR direttamente: rescale_alphas applica AFF_MAP[ui_choice] diretto.
        multipliers = np.ones(len(tr_arr))

        tr_dict[model_suffix]    = tr_arr
        alpha_dict[model_suffix] = tr_arr * multipliers

    # ── Filtro opzioni TR=0 ───────────────────────────────────────────────────
    # Costruisce maschera booleana: True se l'opzione ha TR > 0 in almeno un modello.
    n = len(values)
    active_mask = np.zeros(n, dtype=bool)
    for tr_arr in tr_dict.values():
        active_mask |= (tr_arr > 0)

    n_zero = int(np.sum(~active_mask))
    if n_zero > 0:
        log.info("[SKIP] '%s': %d opzioni con TR=0 su tutti i modelli, rimosse", name, n_zero)
        values     = [v for v, keep in zip(values, active_mask) if keep]
        tr_dict    = {m: arr[active_mask] for m, arr in tr_dict.items()}
        alpha_dict = {m: arr[active_mask] for m, arr in alpha_dict.items()}
    # ─────────────────────────────────────────────────────────────────────────

    char = CharacteristicGroup(
        name=name, values=values,
        tr=tr_dict, alpha=alpha_dict
    )
    char.validate()
    return char

def _parse_bundles(group: pd.DataFrame) -> Dict[str, CharacteristicGroup]:
    """
    Parse BUNDLE rows as binary characteristics — VERSIONE SCALABILE.

    Per ogni bundle "BUNDLE - <Nome>" crea un CharacteristicGroup binario
    con values=['Yes', 'Not'] e dizionari tr/alpha con un array per modello.
    I modelli vengono rilevati dinamicamente dai nomi colonna.
    """
    bundles = {}
    processed = set()

    # Rileva modelli disponibili nelle colonne del gruppo
    model_cols = _extract_model_columns(group.columns)

    for idx, row in group.iterrows():
        full_name = row['CONCAT Nome Opzione/Bundle - Caratteristica']

        if pd.isna(full_name) or full_name in processed:
            continue

        if 'NOT' in full_name.upper():
            processed.add(full_name)
            continue

        bundle_name = full_name.replace('BUNDLE - ', '')
        not_name    = f"BUNDLE - NOT {bundle_name}"
        not_rows    = group[group['CONCAT Nome Opzione/Bundle - Caratteristica'] == not_name]

        tr_dict    = {}
        alpha_dict = {}

        # Controlla se il bundle YES ha TR=0 su TUTTI i modelli -> scarta
        all_yes_zero = all(
            (row[cols['tr_col']] == 0 or pd.isna(row[cols['tr_col']]))
            for cols in model_cols.values()
        )
        if all_yes_zero:
            log.info("[SKIP] Bundle '%s': TR=0 su tutti i modelli, ignorato", bundle_name)
            processed.add(full_name)
            if not not_rows.empty:
                processed.add(not_name)
            continue

        # Alpha = TR direttamente: rescale_alphas applica AFF_MAP[ui_choice] diretto.
        mult_yes = 1.0
        mult_no  = 1.0

        for model_suffix, cols in model_cols.items():
            tr_col = cols['tr_col']

            # Valori YES (riga corrente)
            tr_yes = row[tr_col] if not pd.isna(row[tr_col]) else 0.0

            # Valori NOT (riga "BUNDLE - NOT <Nome>")
            if not not_rows.empty:
                tr_no = not_rows[tr_col].values[0] if not pd.isna(not_rows[tr_col].values[0]) else 0.0
            else:
                # Complementare: NOT = 1 - YES
                tr_no = 1.0 - tr_yes if tr_yes <= 1.0 else 0.0

            # Alpha calcolato: TR × moltiplicatore(Affidabilità TR)
            alpha_yes = tr_yes * mult_yes
            alpha_no  = tr_no  * mult_no

            tr_dict[model_suffix]    = np.array([tr_yes, tr_no])
            alpha_dict[model_suffix] = np.array([alpha_yes, alpha_no])

        char = CharacteristicGroup(
            name=bundle_name, values=['Yes', 'Not'],
            tr=tr_dict, alpha=alpha_dict
        )
        char.validate()
        bundles[bundle_name] = char

        if not not_rows.empty:
            processed.add(not_name)
        processed.add(full_name)

    return bundles

def _parse_month_str(month_str: str) -> Tuple[int, int]:
    """
    Converte una stringa mese in (anno, mese_numero) per ordinamento.

    Accetta formati inglesi e italiani abbreviati, con anno a 2 o 4 cifre:
        "Jan 2026", "Jan2026", "january 2026"   -> EN
        "Gen26", "Gen 26", "Feb26", "Mar26"     -> IT abbreviato a 2 cifre anno

    Returns:
        Tupla (anno, mese) oppure (9999, 99) se non parsabile.
    """
    # Mapping abbreviazioni italiane -> inglesi
    IT_TO_EN = {
        'gen': 'jan', 'feb': 'feb', 'mar': 'mar', 'apr': 'apr',
        'mag': 'may', 'giu': 'jun', 'lug': 'jul', 'ago': 'aug',
        'set': 'sep', 'ott': 'oct', 'nov': 'nov', 'dic': 'dec'
    }

    s = month_str.strip()

    # Normalizza: sostituisci abbreviazioni italiane con inglesi (case insensitive)
    # Usa lookahead (?=\d|$|\s) invece di \b finale: \b non matcha tra lettera
    # e cifra (es. "Gen26" -> \b non c'e' tra 'n' e '2' perche' entrambi sono \w)
    import re as _re
    for it, en in IT_TO_EN.items():
        s = _re.sub(r'(?i)(?<![a-zA-Z])' + it + r'(?=\d|\s|$)', en, s)

    s = s.replace('_', ' ')
    s = ' '.join(s.split())

    # Espandi anno a 2 cifre -> 4 cifre (es. "26" -> "2026")
    # Pattern: "<mese><spazio?><2cifre>" dove le 2 cifre sono l'anno
    year_2digit = _re.match(r'^([a-zA-Z]+)\s*(\d{2})$', s)
    if year_2digit:
        s = f"{year_2digit.group(1)} 20{year_2digit.group(2)}"

    for fmt in ('%b %Y', '%B %Y', '%b%Y', '%B%Y'):
        try:
            dt = datetime.strptime(s, fmt)
            return (dt.year, dt.month)
        except ValueError:
            continue
    return (9999, 99)


def load_monthly_tr(
    filepath: str = '.\\Input\\TR TOTALV21E - Copia.xlsx'
) -> Dict[str, Tuple['ModelMix', Dict[str, 'CharacteristicGroup']]]:
    """Carica i Take Rate mensili dal file Excel con fogli TR_Optional_<mese> e TR_Model_<mese>.

    Mesi accettati in formato italiano (Gen/Feb/.../Dic) o inglese, anno a 2 o 4 cifre.
    Se manca il foglio TR_Model per un mese, viene usato il mese precedente come fallback.

    Args:
        filepath: percorso al file Excel con i TR mensili.

    Returns:
        Dizionario { 'Jan 2026': (ModelMix, characteristics), ... } ordinato cronologicamente.
    """
    import re as _re

    xl = pd.ExcelFile(filepath)
    all_sheets = xl.sheet_names

    log.info("File TR: %s", filepath)
    log.info("Fogli trovati: %s", all_sheets)

    # -----------------------------------------------------------------------
    # Classifica i fogli: TR_Optional_<mese> e TR_Model_<mese>
    # Prefissi attesi (case insensitive):
    #   "TR_Optional_"  -> foglio L1 caratteristiche
    #   "TR_Model_"     -> foglio L0 mix modelli
    # -----------------------------------------------------------------------
    optional_sheets = {}   # { (anno, mese): sheet_name }
    model_sheets    = {}   # { (anno, mese): sheet_name }

    for sheet in all_sheets:
        clean = sheet.strip()
        low   = clean.lower()

        if low.startswith('tr_optional_'):
            date_part = clean[len('tr_optional_'):]
            key = _parse_month_str(date_part)
            if key != (9999, 99):
                optional_sheets[key] = sheet
            else:
                log.warning("foglio '%s' non parsabile come mese, ignorato", sheet)

        elif low.startswith('tr_model_'):
            date_part = clean[len('tr_model_'):]
            key = _parse_month_str(date_part)
            if key != (9999, 99):
                model_sheets[key] = sheet
            else:
                log.warning("foglio '%s' non parsabile come mese, ignorato", sheet)

        else:
            log.info("Foglio ignorato (non TR_Optional/TR_Model): '%s'", sheet)

    if not optional_sheets:
        raise ValueError(
            f"Nessun foglio 'TR_Optional_<mese>' trovato in '{filepath}'. "
            f"Usa nomi come 'TR_Optional_Gen26', 'TR_Optional_Jan2026'."
        )

    # -----------------------------------------------------------------------
    # Parsing di ogni mese disponibile per gli opzionali
    # -----------------------------------------------------------------------
    monthly_tr: Dict[str, Tuple['ModelMix', Dict]] = {}

    # Tieni traccia dell'ultimo ModelMix caricato per il fallback modelli
    last_model_mix: Optional['ModelMix'] = None

    for month_key in sorted(optional_sheets.keys()):
        year, month_num = month_key
        opt_sheet  = optional_sheets[month_key]
        month_label = datetime(year, month_num, 1).strftime('%b %Y')  # "Jan 2026"

        log.info("Parsing mese: %s (optional: '%s')", month_label, opt_sheet)

        try:
            # ---- L1: Caratteristiche e opzionali ----
            _hdr_opt = _detect_header_row(filepath, opt_sheet)
            df = pd.read_excel(filepath, sheet_name=opt_sheet, header=_hdr_opt)
            characteristics: Dict[str, 'CharacteristicGroup'] = {}

            for char_name, group in df.groupby('Nome Opzione/Bundle'):
                if pd.isna(char_name):
                    continue
                if char_name == 'BUNDLE':
                    characteristics.update(_parse_bundles(group))
                else:
                    try:
                        mc = _extract_model_columns(group.columns)
                    except ValueError:
                        mc = {}
                    if mc and _all_tr_zero(group, mc):
                        log.info("[SKIP] Caratteristica '%s': TR=0 su tutti i modelli, ignorata", char_name)
                        continue
                    characteristics[char_name] = _parse_characteristic(char_name, group)

            # ---- L0: Mix modelli ----
            # Priorità: foglio TR_Model_<stessomese> > ultimo caricato > errore
            if month_key in model_sheets:
                mod_sheet = model_sheets[month_key]
                _hdr_mod = _detect_header_row(filepath, mod_sheet,
                                              probe_col='CONCAT Nome Opzione/Bundle')
                df_model  = pd.read_excel(filepath, sheet_name=mod_sheet, header=_hdr_mod)
                raw_models = df_model['CONCAT Nome Opzione/Bundle - Caratteristica'].tolist()
                raw_trs    = df_model['Take Rate Model'].tolist()
                # Canonicalizza nomi modello e scarta righe vuote/NaN
                # mantenendo l'allineamento posizionale fra models e trs.
                pairs = [(canonical_model_name(m), tr)
                         for m, tr in zip(raw_models, raw_trs)
                         if canonical_model_name(m) != '']
                models = [p[0] for p in pairs]
                trs    = np.array([p[1] for p in pairs], dtype=float)
                # Alpha calcolato: Take Rate × moltiplicatore(Affidabilità TR)
                aff_col_m = next(
                    (c for c in df_model.columns if c.strip().upper().startswith('AFFIDABILIT')),
                    None
                )
                # Alpha = TR direttamente: affidabilità file ignorata.
                # rescale_alphas applica AFF_MAP[ui_choice] direttamente.
                alphas_m  = trs * 1.0
                month_model_mix = ModelMix(models=models, take_rates=trs, alphas=alphas_m)
                month_model_mix.validate()
                last_model_mix = month_model_mix
                log.info("[OK] Mix modelli da foglio '%s'", mod_sheet)
            elif last_model_mix is not None:
                month_model_mix = last_model_mix
                log.info("[OK] Mix modelli: fallback mese precedente")
            else:
                raise ValueError(
                    f"Nessun foglio TR_Model disponibile per {month_label} "
                    f"e nessun mese precedente come fallback. "
                    f"Aggiungere foglio 'TR_Model_{opt_sheet[len('TR_Optional_'):]}' al file."
                )

            monthly_tr[month_label] = (month_model_mix, characteristics)
            log.info("[OK] %d caratteristiche caricate", len(characteristics))

        except Exception as e:
            log.warning("ERRORE parsing mese %s: %s", month_label, e)
            raise

    log.info("Totale mesi caricati: %d -> %s", len(monthly_tr), list(monthly_tr.keys()))
    return monthly_tr


def get_tr_for_month(
    month_name: str,
    monthly_tr: Dict[str, Tuple['ModelMix', Dict[str, 'CharacteristicGroup']]]
) -> Tuple['ModelMix', Dict[str, 'CharacteristicGroup']]:
    """
    Restituisce i TR per il mese richiesto, con fallback al più recente disponibile.

    LOGICA FALLBACK:
    Se il mese esatto non è presente nel dizionario mensile, viene cercato
    il mese disponibile più recente che sia <= al mese richiesto.
    Se non esiste nemmeno uno precedente, si usa il più vecchio disponibile.

    Questo permette di gestire scenari dove i TR non cambiano ogni mese:
    basta inserire il primo mese con i nuovi TR e tutti i mesi successivi
    fino al prossimo foglio useranno automaticamente quei valori.

    Args:
        month_name: Nome mese da cercare (es. "Mar 2026")
        monthly_tr: Dizionario { mese: (ModelMix, characteristics) }

    Returns:
        Tupla (ModelMix, characteristics) per il mese o il fallback.

    Esempio:
        monthly_tr ha: "Jan 2026", "Apr 2026"
        get_tr_for_month("Feb 2026", ...) -> restituisce "Jan 2026"  (più recente <=)
        get_tr_for_month("May 2026", ...) -> restituisce "Apr 2026"  (più recente <=)
        get_tr_for_month("Dec 2025", ...) -> restituisce "Jan 2026"  (il più vecchio)
    """
    # Normalizza il nome mese in ingresso in (anno, mese_num) per confronto
    # robusto indipendente da case, spazi, formato italiano/inglese.
    # Poi ricostruiamo la label canonica "Jan 2026" per il lookup esatto.
    target_key = _parse_month_str(month_name)

    if target_key != (9999, 99):
        # Ricostruisci label canonica e prova hit esatto normalizzato
        from datetime import datetime as _dt
        canonical = _dt(target_key[0], target_key[1], 1).strftime('%b %Y')  # "Jan 2026"
        if canonical in monthly_tr:
            return monthly_tr[canonical]

    # Converti tutti i mesi disponibili in (anno, mese) per confronto
    available = sorted(monthly_tr.keys(), key=_parse_month_str)

    # Cerca il mese più recente <= target
    fallback = None
    for m in available:
        if _parse_month_str(m) <= target_key:
            fallback = m
        else:
            break  # lista è ordinata, inutile continuare

    if fallback is None:
        # Tutti i mesi disponibili sono più recenti del target -> usa il più vecchio
        fallback = available[0]
        log.warning(
            "'%s' precedente a tutti i TR disponibili -> uso '%s' (il più vecchio)",
            month_name, fallback,
        )
    else:
        log.warning(
            "'%s' non trovato -> fallback a '%s' (più recente disponibile)",
            month_name, fallback,
        )

    return monthly_tr[fallback]

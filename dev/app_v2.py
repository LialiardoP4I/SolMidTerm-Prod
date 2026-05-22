# -*- coding: utf-8 -*-
"""
================================================================================
DUCATI MULTISTRADA - SAFETY STOCK SIMULATOR
Frontend Streamlit per la pipeline Monte Carlo
================================================================================
Struttura:
  Tab 1 — ▶️  Esegui Analisi   : lancio Safety Stock + Cycle Stock, verifica input
  Tab 2 — 📊  Visualizza Risultati : analisi integrata, risultati, grafici, dati
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import sys
import time
import gc
import psutil
import io
import shutil
import copy
from pathlib import Path

# ============================================================================
# CONFIGURAZIONE PAGINA
# ============================================================================

st.set_page_config(
    page_title="Ducati Safety Stock Simulator",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================================
# CSS PERSONALIZZATO
# ============================================================================

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #E3000F 0%, #a80000 100%);
        padding: 1.8rem 2rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 1.8rem;
        box-shadow: 0 4px 20px rgba(227,0,15,0.3);
    }
    .main-header h1 { margin: 0; font-size: 2rem; font-weight: 800; }
    .main-header p  { margin: 0.4rem 0 0; opacity: 0.85; font-size: 0.95rem; }

    .step-card {
        border-left: 4px solid #E3000F;
        background: #1a1a2e;
        padding: 0.9rem 1.3rem;
        margin: 0.4rem 0;
        border-radius: 0 8px 8px 0;
    }
    .step-card h4 { margin: 0 0 0.25rem; color: #E3000F; }
    .step-card p  { margin: 0; color: #ccc; font-size: 0.88rem; }

    .status-ok   { background:#1a4a1a; color:#4caf50; padding:3px 10px; border-radius:20px; font-size:0.8rem; display:inline-block; }
    .status-err  { background:#4a1a1a; color:#f44336; padding:3px 10px; border-radius:20px; font-size:0.8rem; display:inline-block; }
    .status-warn { background:#4a3a1a; color:#ff9800; padding:3px 10px; border-radius:20px; font-size:0.8rem; display:inline-block; }

    .log-console {
        background: #0d1117;
        border: 1px solid #333;
        border-radius: 8px;
        padding: 1rem;
        font-family: monospace;
        font-size: 0.78rem;
        color: #39ff14;
        max-height: 280px;
        overflow-y: auto;
        white-space: pre-wrap;
    }
    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #E3000F;
        margin: 1rem 0 0.5rem;
    }

    /* ── Titoli sezione uniformi rosso Ducati ── */
    [data-testid="stMarkdownContainer"] h2 {
        color: #E3000F !important;
        border-bottom: 2px solid #E3000F;
        padding-bottom: 0.25rem;
        margin-top: 1.2rem;
    }
    [data-testid="stMarkdownContainer"] h3 {
        color: #E3000F !important;
    }

    /* ── Tab principali più grandi e prominenti ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #0e1117;
        border-radius: 10px;
        padding: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        padding: 0.75rem 2.5rem !important;
        border-radius: 8px !important;
        letter-spacing: 0.04em;
        transition: background 0.2s;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #E3000F 0%, #a80000 100%) !important;
        color: white !important;
    }
    .stTabs [aria-selected="false"] {
        color: #aaa !important;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none !important;
    }
    .stTabs [data-baseweb="tab-border"] {
        display: none !important;
    }

    /* ── Card parametri pipeline ── */
    .param-card {
        background: #1a1a2e;
        border: 1px solid #333;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.8rem;
    }
    .param-card h4 {
        color: #E3000F;
        margin: 0 0 0.5rem;
        font-size: 0.95rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER
# ============================================================================

_hc_logo, _hc_title = st.columns([1, 4])
with _hc_logo:
    _logo_path = Path(__file__).parent / "Immagine1.png"
    if _logo_path.exists():
        st.image(str(_logo_path), use_container_width=True)
with _hc_title:
    st.markdown("""
<div class="main-header">
    <h1>Ducati Safety Stock Simulator</h1>
    <p>Monte Carlo Pipeline · Multistrada V4 · Supply Chain Planning</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# UTILITIES & COSTANTI
# ============================================================================

BASE_DIR   = Path(__file__).parent
INPUT_DIR  = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"

ABC_COLORS = {"A": "#E3000F", "B": "#ff8c00", "C": "#1a6fa8", "N/D": "#666"}

INPUT_FILES = {
    "BOM150":      INPUT_DIR / "BOM150_V21E.xlsm",
    "Mappatura":   INPUT_DIR / "Mappatura - Copia.xlsx",       # [V2]
    "Esclusioni":  INPUT_DIR / "esclusioni.xlsx",
    "Domanda":     INPUT_DIR / "Total_demand.xlsx",
    "Take Rate":   INPUT_DIR / "TR TOTALV21E - Copia.xlsx",    # [V2]
    "Lead Times":  INPUT_DIR / "quantity e residual.xlsx",
    "Matrix":      INPUT_DIR / "matrix_output_V2.xlsx",        # SUB-TAB D — TOTAL DEMAND
    "Unique Rows": INPUT_DIR / "tutte_righe_univoche_V2.xlsx", # [V2]
    "Prezzi":      INPUT_DIR / "PREZZI.xlsx",   # opzionale
}

OUTPUT_FILES = {
    "Safety Stock":         OUTPUT_DIR / "sku_safety_stock_leadtime_PER_SKU_V2.xlsx",
    "Safety Stock Mensile": OUTPUT_DIR / "sku_safety_stock_PER_MESE_V2.xlsx",
    "Demand Comparison":    OUTPUT_DIR / "demand_comparison_actual_vs_forecast_V2.xlsx",
}

# ── Helpers ─────────────────────────────────────────────────────────────────

def get_memory_mb() -> float:
    return psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024

def file_exists(path: Path) -> bool:
    return path.exists()

def format_size(path: Path) -> str:
    if not path.exists():
        return "N/A"
    s = path.stat().st_size
    if s < 1024:      return f"{s} B"
    if s < 1024**2:   return f"{s/1024:.1f} KB"
    return f"{s/1024**2:.1f} MB"

def safe_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].astype(str)
    return out

# ── Affidabilità TR ──────────────────────────────────────────────────────────
# Legge la mappa direttamente da tr_parser_V2, che la carica da
# Input/affidabilita_config.json — si aggiorna automaticamente a ogni
# modifica del JSON (senza toccare app.py).
try:
    from tr_parser_V2 import _AFFIDABILITA_MAP as _AFF_MAP
    _AFF_MAP = {k: float(v) for k, v in _AFF_MAP.items()}
except Exception:
    _AFF_MAP = {"BASSA": 10.0, "MEDIA": 100.0, "ALTA": 1000.0}

# Livelli ordinati dal più basso al più alto moltiplicatore
_AFF_LEVELS = sorted(_AFF_MAP, key=_AFF_MAP.__getitem__)
# Moltiplicatore "originale" usato dal parser come fallback (il massimo nella mappa)
_AFF_ORIG_MULT = max(_AFF_MAP.values())

def rescale_alphas(model_mix, characteristics, monthly_tr,
                   aff_model: str, aff_optional: str):
    """
    Riscala gli alpha di model_mix e characteristics in base ai livelli
    di affidabilità scelti nella UI.

    Il parser calcola alpha = TR × moltiplicatore_originale (letto dall'Excel;
    quando Affidabilità non è presente usa il massimo della mappa come fallback).
    Qui rimuoviamo quel moltiplicatore e applichiamo quello scelto dall'utente.
    La mappa viene letta da Input/affidabilita_config.json via tr_parser,
    quindi si adatta automaticamente a nuovi livelli aggiunti nel JSON.
    """
    orig    = _AFF_ORIG_MULT                    # massimo nella mappa corrente
    scale_m = _AFF_MAP[aff_model]    / orig
    scale_o = _AFF_MAP[aff_optional] / orig

    mm   = copy.deepcopy(model_mix)
    char = copy.deepcopy(characteristics)

    # Riscala Model Mix (L0)
    mm.alphas = mm.alphas * scale_m

    # Riscala Caratteristiche (L1)
    for cg in char.values():
        for k in list(cg.alpha.keys()):
            cg.alpha[k] = cg.alpha[k] * scale_o

    # Riscala monthly_tr se presente
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

# ── Archiviazione storico ────────────────────────────────────────────────────

def archive_run(aff_model: str, aff_optional: str, n_runs: int = 0,
                percentile: int = 99, month: str = ""):
    """
    Copia il file Safety Stock PER_SKU in Output/history/ con nome
    GG-MM-AAAA_<mese>_ModelLV-OptLV_NNNrun_P{percentile}.xlsx
    (es. 19-02-2026_Gen26_ALTA-MEDIA_500run_P99.xlsx).
    Ritorna il Path del file archiviato, o None se la sorgente non esiste.
    Se esiste già un file con lo stesso nome lo sovrascrive.
    """
    src = OUTPUT_FILES["Safety Stock"]
    if not src.exists():
        return None
    history_dir = OUTPUT_DIR / "history"
    history_dir.mkdir(exist_ok=True)
    date_str  = time.strftime("%d-%m-%Y")
    month_str = f"_{month}" if month else ""
    fname = f"{date_str}{month_str}_{aff_model}-{aff_optional}_{n_runs}run_P{percentile}.xlsx"
    dst = history_dir / fname
    shutil.copy2(src, dst)
    return dst

# ── Cached loaders ──────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def load_safety_stock(path: str) -> pd.DataFrame:
    return pd.read_excel(path)

@st.cache_data(ttl=300, show_spinner=False)
def load_safety_stock_monthly(path: str) -> pd.DataFrame:
    return pd.read_excel(path)

@st.cache_data(ttl=300, show_spinner=False)
def load_demand(path: str) -> pd.DataFrame:
    return pd.read_excel(path, header=None)

@st.cache_data(ttl=300, show_spinner=False)
def load_tr_sheets_list(path: str):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    return sheets

@st.cache_data(ttl=300, show_spinner=False)
def load_tr_sheet(path: str, sheet_name: str, header_row: int = 0) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, header=header_row)

@st.cache_data(ttl=300, show_spinner=False)
def load_comparison(path: str) -> dict:
    xl = pd.ExcelFile(path)
    return {sh: xl.parse(sh) for sh in xl.sheet_names}

@st.cache_data(ttl=300, show_spinner=False)
def get_demand_months(path: str):
    try:
        df_raw = pd.read_excel(path, header=None)
        if len(df_raw) >= 1:
            months_row = df_raw.iloc[0].dropna().tolist()
            return [str(m) for m in months_row if "tot" not in str(m).lower()]
    except Exception:
        pass
    return []

# ============================================================================
# PARAMETRI PIPELINE — definiti a livello globale per SS e CS
# ============================================================================

# Leggo mesi domanda (necessario per le selectbox)
demand_path          = INPUT_FILES["Domanda"]
demand_months        = []
start_month_offset   = 0
selected_start_month = ""

if file_exists(demand_path):
    demand_months = get_demand_months(str(demand_path))

# Default affidabilità
_aff_def_model    = _AFF_LEVELS[-1]
_aff_def_optional = _AFF_LEVELS[len(_AFF_LEVELS) // 2]


# ============================================================================
# NAVIGAZIONE PRINCIPALE
# ============================================================================

if "active_module" not in st.session_state:
    st.session_state["active_module"] = "ss"

_active_module = st.session_state["active_module"]

_col_ss, _col_cs = st.columns(2)
with _col_ss:
    if st.button(
        "Safety Stock",
        use_container_width=True,
        type="primary" if _active_module == "ss" else "secondary",
        key="nav_ss",
    ):
        st.session_state["active_module"] = "ss"
        st.rerun()
with _col_cs:
    if st.button(
        "Cycle Stock",
        use_container_width=True,
        type="primary" if _active_module == "cs" else "secondary",
        key="nav_cs",
    ):
        st.session_state["active_module"] = "cs"
        st.rerun()

# ── Valori parametri letti da session_state (sovrascritta dai widget nel sub-tab Parametri) ──
n_runs               = st.session_state.get("ss_n_runs",        200)
percentile           = st.session_state.get("ss_percentile",    99)
start_month_offset   = st.session_state.get("ss_start_offset",  0)
selected_start_month = st.session_state.get("ss_start_month",   demand_months[0] if demand_months else "")
run_step0            = st.session_state.get("ss_step0",         False)
run_step2            = st.session_state.get("ss_step2",         True)
run_step3            = st.session_state.get("ss_step3",         True)
aff_model            = st.session_state.get("ss_aff_model",     _aff_def_model)
aff_optional         = st.session_state.get("ss_aff_optional",  _aff_def_optional)
filter_untraced      = st.session_state.get("ss_filter_untraced", True)


# ============================================================================
# MODULO — SAFETY STOCK
# ============================================================================

if _active_module == "ss":

    tab_run, tab_view = st.tabs(["Esegui Analisi", "Visualizza Risultati"])

    # ########################################################################
    # SS — TAB 1: ESEGUI ANALISI
    # ########################################################################

    with tab_run:

        # ── Leggo valori parametri da session_state (scritti dai widget sub-tab Parametri) ──
        n_runs               = st.session_state.get("ss_n_runs",        200)
        percentile           = st.session_state.get("ss_percentile",    99)
        start_month_offset   = st.session_state.get("ss_start_offset",  0)
        selected_start_month = st.session_state.get("ss_start_month",   demand_months[0] if demand_months else "")
        run_step0            = st.session_state.get("ss_step0",         False)
        run_step2            = st.session_state.get("ss_step2",         True)
        run_step3            = st.session_state.get("ss_step3",         True)
        aff_model            = st.session_state.get("ss_aff_model",     _aff_def_model)
        aff_optional         = st.session_state.get("ss_aff_optional",  _aff_def_optional)
        _ram_ss              = st.session_state.get("last_ram",         get_memory_mb())

        # ── Stato file di input ──────────────────────────────────────────────────

        SS_REQUIRED  = ["BOM150", "Mappatura", "Esclusioni", "Domanda", "Take Rate",
                        "Lead Times", "Matrix", "Unique Rows"]
        CMP_REQUIRED = ["Esclusioni", "Domanda", "Lead Times", "Unique Rows"]

        def _file_status_row(names: list, label: str):
            """Mostra una riga di badge per un gruppo di file."""
            cols = st.columns(len(names))
            for col, name in zip(cols, names):
                path = INPUT_FILES[name]
                ok   = file_exists(path)
                badge_cls = "status-ok" if ok else "status-err"
                icon      = "✅" if ok else "❌"
                with col:
                    st.markdown(
                        f'{icon} **{name}**<br>'
                        f'<span class="{badge_cls}">{format_size(path) if ok else "mancante"}</span>',
                        unsafe_allow_html=True,
                    )

        with st.expander("Safety Stock — file richiesti", expanded=True):
            _file_status_row(SS_REQUIRED, "ss")
            _miss_ss = [n for n in SS_REQUIRED
                        if not file_exists(INPUT_FILES[n]) and n not in ("Matrix", "Unique Rows")]
            if _miss_ss:
                st.error(f"File mancanti: {', '.join(_miss_ss)}")
            else:
                st.success("Tutti i file richiesti per Safety Stock sono presenti.")

        with st.expander("Cycle Stock — file richiesti", expanded=False):
            _file_status_row(CMP_REQUIRED, "cmp")
            prezzi_ok = file_exists(INPUT_FILES["Prezzi"])
            pz_icon   = "✅" if prezzi_ok else ""
            st.markdown(
                f"{pz_icon} **Prezzi** *(opzionale)*  — "
                + (f"`{format_size(INPUT_FILES['Prezzi'])}` — calcolo Δ € abilitato" if prezzi_ok
                   else "assente · confronto solo su volumi, nessun delta €"),
            )

        st.markdown("---")

        # ════════════════════════════════════════════════════════════════════════
        # SEZIONE A — SAFETY STOCK
        # ════════════════════════════════════════════════════════════════════════

        st.markdown("## Safety Stock")

        # Pipeline SS — eseguita quando l'utente preme "AVVIA" in fondo alla pagina
        if st.session_state.pop("_do_run_ss", False):
            if st.session_state.get("_ss_running", False):
                st.warning("Pipeline Safety Stock già in esecuzione — attendi il completamento.")
                st.stop()
            st.session_state["_ss_running"] = True

            log_container = st.empty()
            progress_bar  = st.progress(0)
            status_text   = st.empty()
            logs = []

            def log(msg: str):
                logs.append(msg)
                log_container.markdown(
                    f'<div class="log-console">{"<br>".join(logs[-30:])}</div>',
                    unsafe_allow_html=True,
                )

            pipeline_start = time.time()

            try:
                if str(BASE_DIR) not in sys.path:
                    sys.path.insert(0, str(BASE_DIR))

                log("=" * 60)
                log("PIPELINE DUCATI SAFETY STOCK SIMULATOR")
                log(f"Avvio: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                log(f"Parametri: n_runs={n_runs}, offset={start_month_offset}, P{percentile}")
                if demand_months:
                    log(f"Mese inizio: {selected_start_month}")
                log("=" * 60)

                # ── Step 1 ──────────────────────────────────────────────────────
                # Verifica se è già stato eseguito process_all_bom_combinations.py
                _consolidated_catalog = INPUT_DIR / "tutte_righe_univoche_V2.xlsx"
                _has_consolidated = False

                if _consolidated_catalog.exists():
                    try:
                        # Verifica se il file consolidato contiene dati da V21E, V22E, V24T
                        _test_df = pd.read_excel(_consolidated_catalog, sheet_name='Tutte_Righe_Univoche', nrows=1)
                        if 'Versione' in _test_df.columns:
                            _has_consolidated = True
                            log("[STEP 1] Rilevato file consolidato da process_all_bom_combinations.py")
                            log("  → File contiene dati da V21E, V22E, V24T")
                    except:
                        pass

                if run_step0 and not _has_consolidated:
                    status_text.text("Step 1: Elaborazione BOM...")
                    progress_bar.progress(5)
                    log("\n[STEP 1] Preparazione Dati BOM (V21E solo)...")
                    t0 = time.time()
                    try:
                        from gestione_dipendenze_V2 import process_excel_to_matrix_v2    # [V2]
                        from skubundle_OPTIMIZED_V2 import crea_tutte_righe_univoche_v2  # [V2]
                        log("  → Elaborazione BOM150 → Matrice (V2)...")
                        process_excel_to_matrix_v2(
                            excel_path=str(INPUT_DIR / "BOM150_V21E.xlsm"),
                            mapping_file=str(INPUT_DIR / "Mappatura - Copia.xlsx"),      # [V2]
                            tr_file=str(INPUT_DIR / "TR TOTALV21E - Copia.xlsx"),        # [V2]
                            data_sheet="BOM150",
                            output_file=str(INPUT_DIR / "matrix_output_V2.xlsx"),        # [V2]
                            prezzi_file=str(INPUT_FILES["Prezzi"]) if file_exists(INPUT_FILES["Prezzi"]) else None,
                            filter_untraced=filter_untraced,
                        )
                        log("  → Creazione righe univoche (V2)...")
                        crea_tutte_righe_univoche_v2(
                            str(INPUT_DIR / "matrix_output_V2.xlsx"),                    # [V2]
                            str(INPUT_DIR / "tutte_righe_univoche_V2.xlsx"),             # [V2]
                        )
                        log(f"  ✓ Step 1 completato in {time.time()-t0:.1f}s")
                        progress_bar.progress(25)
                    except Exception as e:
                        log(f"  ✗ ERRORE Step 1: {e}")
                        st.error(f"Errore Step 1: {e}")
                    gc.collect()
                else:
                    if _has_consolidated:
                        log("[STEP 1] Saltato — usando file consolidato (V21E + V22E + V24T)")
                    else:
                        log("[STEP 1] Saltato (usa file esistenti)")
                    progress_bar.progress(20)

                df_simulation = None

                # ── Step 2 ──────────────────────────────────────────────────────
                if run_step2:
                    status_text.text("Step 2: Simulazione Monte Carlo...")
                    progress_bar.progress(30)
                    log("\n[STEP 2] Simulazione Monte Carlo...")
                    log(f"  Config: {n_runs} run, offset={start_month_offset}")
                    log(f"  Affidabilità → Model Mix: {aff_model}  |  Optional: {aff_optional}")
                    t2 = time.time()
                    try:
                        from mc_simulator_V2 import (
                            SimulationConfig, load_monthly_forecast,
                            calculate_n_months_needed,
                            run_monthly_simulations_optimized,
                            build_wide_dataframe,
                        )
                        from tr_parser_V2 import parse_tr_file
                        from tr_parser_V2 import load_monthly_tr as _load_mtr
                        from mc_simulator_V2 import parse_exclusions as _parse_excl

                        # Carica TR grezzi
                        log("  → Caricamento TR (V2 — Copia)...")
                        _model_mix_raw, _char_raw = parse_tr_file(
                            str(INPUT_DIR / "TR TOTALV21E - Copia.xlsx")
                        )
                        _monthly_tr_raw = None
                        try:
                            _monthly_tr_raw = _load_mtr(str(INPUT_DIR / "TR TOTALV21E - Copia.xlsx"))
                            log(f"  → TR mensili: {len(_monthly_tr_raw)} mesi")
                        except Exception:
                            log("  → Nessun TR mensile, uso TR globale")

                        # Riscala alpha secondo la scelta UI
                        _model_mix_sc, _char_sc, _monthly_tr_sc = rescale_alphas(
                            _model_mix_raw, _char_raw, _monthly_tr_raw,
                            aff_model=aff_model, aff_optional=aff_optional,
                        )
                        _sc_m = _AFF_MAP[aff_model]    / _AFF_ORIG_MULT
                        _sc_o = _AFF_MAP[aff_optional] / _AFF_ORIG_MULT
                        log(f"  → Alpha riscalati (Model ×{_sc_m:.3f}, Optional ×{_sc_o:.3f})")
                        _n_var = sum(
                            1 for cg in _char_sc.values()
                            if sum(1 for arr in cg.alpha.values() if (arr > 0).sum() > 1) > 0
                        )
                        _alpha_mm = float(_model_mix_sc.alphas.mean())
                        log(f"  → Caratteristiche con ≥2 opzioni attive: {_n_var}  "
                            f"| alpha Model Mix medio: {_alpha_mm:.1f}")

                        # Esegui la simulazione con oggetti riscalati
                        _cfg = SimulationConfig(
                            n_runs=n_runs, random_seed=42,
                            start_month_offset=start_month_offset, use_int16=True,
                        )
                        _forecast_values, _month_names = load_monthly_forecast(
                            str(INPUT_DIR / "Total_demand.xlsx"),
                            start_offset=start_month_offset,
                        )
                        _n_months = min(
                            calculate_n_months_needed(str(INPUT_DIR / "quantity e residual.xlsx")),
                            len(_forecast_values),
                        )
                        _exclusions = _parse_excl(str(INPUT_DIR / "esclusioni.xlsx"))

                        _results_by_month = run_monthly_simulations_optimized(
                            forecast_values=_forecast_values,
                            month_names=_month_names,
                            n_months=_n_months,
                            model_mix=_model_mix_sc,
                            characteristics=_char_sc,
                            exclusions=_exclusions,
                            config=_cfg,
                            monthly_tr=_monthly_tr_sc,
                        )
                        df_simulation = build_wide_dataframe(
                            _results_by_month, _month_names, n_runs=_cfg.n_runs, save_to_file=False  # Non salva su disco: passa direttamente in memoria
                                                )

                        nc = len(df_simulation.columns) if df_simulation is not None else 0
                        log(f"  ✓ Simulazione completata in {time.time()-t2:.1f}s")
                        log(f"  ✓ {nc} colonne simulate generate")
                        st.session_state["df_simulation"] = df_simulation
                        progress_bar.progress(65)
                    except Exception as e:
                        log(f"  ✗ ERRORE Step 2: {e}")
                        st.error(f"Errore Step 2: {e}")
                        raise
                    gc.collect()
                elif "df_simulation" in st.session_state:
                    df_simulation = st.session_state["df_simulation"]
                    log("[STEP 2] Saltato (usa simulazione in sessione)")
                    progress_bar.progress(65)
                else:
                    log("[STEP 2] Saltato — nessuna simulazione in sessione")
                    progress_bar.progress(65)

                # ── Step 3 ──────────────────────────────────────────────────────
                if run_step3:
                    status_text.text("Step 3: Matching SKU & Safety Stock...")
                    log("\n[STEP 3] Matching SKU e calcolo Safety Stock...")
                    log(f"  Percentile: P{percentile}")
                    t3 = time.time()
                    try:
                        from sku_matching_V2 import main_sku_v2
                        if df_simulation is None:
                            log("  ✗ Nessun risultato simulazione disponibile")
                            st.error("Step 3 richiede Step 2. Abilita Step 2 dalla sidebar.")
                        else:
                            results = main_sku_v2(
                                df_simulation,
                                percentile=percentile,
                                sku_catalog_path=str(INPUT_DIR / "tutte_righe_univoche_V2.xlsx"),
                                lead_times_path=str(INPUT_DIR / "quantity e residual.xlsx"),
                                prezzi_path=str(INPUT_DIR / "PREZZI.xlsx"),
                                output_agg=str(OUTPUT_FILES["Safety Stock"]),
                                output_monthly=str(OUTPUT_FILES["Safety Stock Mensile"]),
                            )
                            nr = len(results) if results is not None else 0
                            log(f"  ✓ Matching completato in {time.time()-t3:.1f}s")
                            log(f"  ✓ {nr} componenti SKU elaborati")
                            # Diagnostica prezzi — visibile in UI, non solo terminale
                            if results is not None and "prezzo_safety" in results.columns:
                                _n_con_prezzo = results["prezzo_safety"].notna().sum()
                                _tot_eur = results["prezzo_safety"].sum(skipna=True)
                                log(f"  ✓ Prezzi: {_n_con_prezzo}/{nr} SKU con prezzo  "
                                    f"| Totale SS economica: € {_tot_eur:,.0f}")
                            else:
                                log("  ⚠ prezzo_safety assente — "
                                    "verifica che Prezzi.xlsx esista e contenga i fogli "
                                    "'Sheet1' (colonne: Materiale, Prezzo netto, Unità di prezzo, "
                                    "Divisa, Data documento) e 'Cambio Valuta' "
                                    "(colonne: Codice, '1 unità in EUR ≈')")
                            st.session_state["results"] = results
                            st.session_state["percentile_used"] = percentile
                            progress_bar.progress(95)
                    except Exception as e:
                        log(f"  ✗ ERRORE Step 3: {e}")
                        st.error(f"Errore Step 3: {e}")
                        raise
                else:
                    log("[STEP 3] Saltato")
                    progress_bar.progress(95)

                # ── Fine ────────────────────────────────────────────────────────
                total_elapsed = time.time() - pipeline_start
                progress_bar.progress(100)
                status_text.text("Pipeline completata!")
                log("\n" + "=" * 60)
                log(f"PIPELINE COMPLETATA in {total_elapsed:.1f}s ({total_elapsed/60:.1f} min)")
                log(f"RAM finale: {get_memory_mb():.0f} MB")
                log("=" * 60)

                st.success(
                    f"✅ Pipeline completata in {total_elapsed:.1f}s — "
                    "vai alla tab **Visualizza Risultati** per l'analisi."
                )
                st.session_state["last_ram"] = get_memory_mb()
                load_safety_stock.clear()
                load_safety_stock_monthly.clear()

                # Archivia output appena generato
                _arch = archive_run(aff_model, aff_optional, n_runs=n_runs,
                                   percentile=percentile, month=selected_start_month)
                if _arch:
                    log(f"Archiviato → history/{_arch.name}")
                    st.toast(f"Salvato in storico: {_arch.name}")

            except Exception as e:
                progress_bar.progress(100)
                status_text.text("Pipeline interrotta per errore.")
                log(f"\n✗ ERRORE CRITICO: {e}")
                st.error(f"Errore durante l'esecuzione: {e}")

            finally:
                st.session_state["_ss_running"] = False


        # ════════════════════════════════════════════════════════════════════════
        # SUB-NAVIGAZIONE — stessa logica della nav principale
        # ════════════════════════════════════════════════════════════════════════

        if "_ss_sub" not in st.session_state:
            st.session_state["_ss_sub"] = "Avvia Pipeline"
        # se sessione precedente aveva "Parametri" (rimosso), porta su Avvia Pipeline
        if st.session_state["_ss_sub"] == "Parametri":
            st.session_state["_ss_sub"] = "Avvia Pipeline"
        _ss_sub = st.session_state["_ss_sub"]

        _sub_labels = ["File di Input", "Qualità Dati", "Affidabilità", "Avvia Pipeline"]
        _sub_cols   = st.columns(len(_sub_labels))
        for _si, (_sc, _sl) in enumerate(zip(_sub_cols, _sub_labels)):
            with _sc:
                if st.button(
                    _sl,
                    use_container_width=True,
                    type="primary" if _ss_sub == _sl else "secondary",
                    key=f"_ss_sub_{_si}",
                ):
                    st.session_state["_ss_sub"] = _sl
                    st.rerun()

        st.markdown("---")

        # ════════════════════════════════════════════════════════════════════════
        # SUB-TAB: PARAMETRI
        # ════════════════════════════════════════════════════════════════════════

        # ════════════════════════════════════════════════════════════════════════
        # SUB-TAB: FILE DI INPUT
        # ════════════════════════════════════════════════════════════════════════

        if _ss_sub == "File di Input":

            st.markdown("## File di Input")

            with st.expander("Safety Stock — file richiesti", expanded=True):
                _file_status_row(SS_REQUIRED, "ss")
                _miss_ss = [n for n in SS_REQUIRED
                            if not file_exists(INPUT_FILES[n]) and n not in ("Matrix", "Unique Rows")]
                if _miss_ss:
                    st.error(f"File mancanti: {', '.join(_miss_ss)}")
                else:
                    st.success("Tutti i file richiesti per Safety Stock sono presenti.")

            with st.expander("Cycle Stock — file richiesti", expanded=False):
                _file_status_row(CMP_REQUIRED, "cmp")
                prezzi_ok = file_exists(INPUT_FILES["Prezzi"])
                pz_icon   = "✅" if prezzi_ok else ""
                st.markdown(
                    f"{pz_icon} **Prezzi** *(opzionale)*  — "
                    + (f"`{format_size(INPUT_FILES['Prezzi'])}` — calcolo Δ € abilitato" if prezzi_ok
                       else "assente · confronto solo su volumi, nessun delta €"),
                )

        # ════════════════════════════════════════════════════════════════════════
        # SUB-TAB: QUALITÀ DATI
        # ════════════════════════════════════════════════════════════════════════

        elif _ss_sub == "Qualità Dati":

            st.markdown("## Verifica Qualità Input")
            st.caption(
            "Analisi automatica della qualità dei file di input: "
            "componenti con residual/quantity anomali e CT/CV non tracciati nella BOM."
            )

            _qc_path = OUTPUT_DIR / "quality_check_report.xlsx"

            # ── Stato del report ─────────────────────────────────────────────────
            if not _qc_path.exists():
                st.warning(
                    "Report non ancora generato.  \n"
                    "Usa il pulsante **Esegui Quality Check** qui sotto, oppure "
                    "esegui la pipeline con Step 1 abilitato."
                )
            else:
                _qc_mtime = time.strftime(
                    "%d/%m/%Y %H:%M", time.localtime(_qc_path.stat().st_mtime)
                )
                st.caption(f"Ultimo aggiornamento: {_qc_mtime}")

                # ── Lettura e pannello di sintesi ────────────────────────────
                def _clean_qc_df(df):
                    """Rimuove colonne Unnamed e normalizza tipi per Arrow/Streamlit."""
                    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")], errors="ignore")
                    for col in df.columns:
                        df[col] = df[col].astype(str).str.strip().replace("nan", "")
                    return df

                try:
                    _xl_qc = pd.ExcelFile(str(_qc_path))

                    # Nomi leggibili per i fogli
                    _qc_sheet_labels = {
                        "Residual & Quantity": "Residual & Quantità",
                        "CT non tracciati":    "CT non tracciati in Mappatura",
                        "CV non tracciati":    "CV non tracciati in Mappatura",
                        "CT di sistema":       "CT di sistema (ignorabili)",
                    }

                    # Rinomina colonne per renderle leggibili
                    _col_rename = {
                        "Problema":            "Tipo problema",
                        "Codice":              "Codice componente",
                        "Descrizione":         "Descrizione",
                        "Residual (mesi)":     "Residual (mesi)",
                        "Quantity":            "Quantità",
                        "CT":                  "CT",
                        "N condizioni":        "N. occorrenze in BOM",
                        "Esempio condizione 1": "Esempio condizione BOM",
                        "CV":                  "CV",
                        "Nota":                "Nota",
                    }
                    # Colonne da nascondere perché ridondanti o troppo verbose
                    _cols_to_drop = {"Esempio condizione 2", "Esempio condizione 3"}

                    _qc_summary = {}
                    for _sn in _xl_qc.sheet_names:
                        # header=2: riga 0=titolo, riga 1=descrizione/legenda, riga 2=intestazioni reali
                        _df_q = _xl_qc.parse(_sn, header=2)
                        _df_q = _df_q.drop(columns=[c for c in _df_q.columns if c in _cols_to_drop], errors="ignore")
                        _df_q = _clean_qc_df(_df_q)
                        # Scarta righe vuote o senza severità valida
                        if "Severita" in _df_q.columns:
                            _df_q = _df_q[_df_q["Severita"].isin(["ERRORE", "WARN", "INFO"])].copy()
                            _df_q.rename(columns=_col_rename, inplace=True)
                            _qc_summary[_sn] = (
                                int((_df_q["Severita"] == "ERRORE").sum()),
                                int((_df_q["Severita"] == "WARN").sum()),
                                int((_df_q["Severita"] == "INFO").sum()),
                                _df_q,
                            )
                        else:
                            _qc_summary[_sn] = (0, 0, 0, _df_q)

                    _tot_err  = sum(v[0] for v in _qc_summary.values())
                    _tot_warn = sum(v[1] for v in _qc_summary.values())
                    _tot_info = sum(v[2] for v in _qc_summary.values())

                    # Banner globale
                    if _tot_err > 0:
                        st.error(
                            f"❌ **{_tot_err} ERRORI** nei dati di input — "
                            "correggi prima di avviare la pipeline."
                        )
                    elif _tot_warn > 0:
                        st.warning(
                            f"⚠️ **{_tot_warn} AVVISI** — la pipeline può girare "
                            "ma verifica i dati segnalati."
                        )
                    else:
                        st.success("✅ Qualità dati: nessun errore o avviso rilevato.")

                    # KPI
                    _qkpi1, _qkpi2, _qkpi3, _qkpi4 = st.columns(4)
                    _qkpi1.metric("🔴 Errori da correggere",  _tot_err)
                    _qkpi2.metric("⚠️ Avvisi da verificare",  _tot_warn)
                    _qkpi3.metric("ℹ️ Info (attesi)",         _tot_info)
                    _qkpi4.metric("📋 Categorie analizzate",  len(_qc_summary))
                    st.markdown("---")

                    # Dettaglio per foglio
                    _sev_cfg = [
                        ("ERRORE", "🔴", "Errori — da correggere urgentemente"),
                        ("WARN",   "⚠️", "Avvisi — verificare prima di procedere"),
                        ("INFO",   "ℹ️", "Info — attesi, nessuna azione richiesta"),
                    ]
                    for _sn, (_ne, _nw, _ni, _df_q) in _qc_summary.items():
                        _sheet_label = _qc_sheet_labels.get(_sn, _sn)
                        if _ne > 0:
                            _col_label = f"🔴 {_sheet_label} — {_ne} errori"
                        elif _nw > 0:
                            _col_label = f"⚠️ {_sheet_label} — {_nw} avvisi"
                        elif _ni > 0:
                            _col_label = f"ℹ️ {_sheet_label} — {_ni} info"
                        else:
                            _col_label = f"✅ {_sheet_label} — nessun problema"

                        with st.expander(_col_label, expanded=(_ne > 0)):
                            if _df_q.empty:
                                st.success("Nessun problema rilevato in questa categoria.")
                            elif "Severita" in _df_q.columns:
                                for _sev, _sev_icon, _sev_label in _sev_cfg:
                                    _sub = _df_q[_df_q["Severita"] == _sev].drop(
                                        columns=["Severita"], errors="ignore"
                                    )
                                    if not _sub.empty:
                                        st.markdown(f"**{_sev_icon} {_sev_label}** ({len(_sub)} voci)")
                                        st.dataframe(
                                            _sub.reset_index(drop=True),
                                            use_container_width=True,
                                            height=min(300, 38 + 35 * len(_sub)),
                                        )
                            else:
                                st.dataframe(
                                    _df_q.reset_index(drop=True),
                                    use_container_width=True,
                                    height=min(300, 38 + 35 * len(_df_q)),
                                )

                    with open(_qc_path, "rb") as _f:
                        st.download_button(
                            "📥 Scarica Report Excel completo",
                            data=_f.read(),
                            file_name="quality_check_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_qc",
                        )

                except Exception as _e:
                    st.error(f"Errore lettura report: {_e}")

            # ── Pulsante standalone ──────────────────────────────────────────────
            _run_qc = st.button(
                "Esegui solo Quality Check",
                help=(
                    "Analizza residual/quantity e dipendenze BOM senza rigenerare "
                    "la matrice completa — utile dopo correzioni ai file di input."
                ),
                key="btn_qc_standalone",
            )
            if _run_qc:
                _qc_files_ok = (
                    file_exists(INPUT_FILES["Lead Times"])
                    and file_exists(INPUT_FILES["Mappatura"])
                    and file_exists(INPUT_FILES["Take Rate"])
                    and file_exists(INPUT_FILES["BOM150"])
                )
                if not _qc_files_ok:
                    st.error("File mancanti: assicurati che Lead Times, Mappatura, Take Rate e BOM150 siano presenti.")
                else:
                    with st.spinner("Analisi qualità in corso..."):
                        try:
                            from gestione_dipendenze_V2 import (
                                load_mapping_from_file    as _gd_lmap,
                                _build_active_cv_sets     as _gd_acv,
                                ConditionParser           as _GD_CP,
                                _check_residual_quality   as _gd_rq,
                                _check_price_quality      as _gd_pq,
                                _check_ct_cv_traceability as _gd_tr,
                                _write_quality_report     as _gd_wr,
                            )

                            _RESIDUAL_COL_QC = (
                                'Residual Purchasing Time (months), '
                                'quanto rimane di mesi rispetto al frozen period'
                            )
                            _COMP_COL_QC = 'Numero componenti'

                            _res_df_qc = pd.read_excel(str(INPUT_FILES["Lead Times"]))
                            if 'Codice componenti' in _res_df_qc.columns:
                                _res_df_qc = _res_df_qc.rename(
                                    columns={'Codice componenti': _COMP_COL_QC}
                                )
                            _res_df_qc[_RESIDUAL_COL_QC] = pd.to_numeric(
                                _res_df_qc[_RESIDUAL_COL_QC], errors='coerce'
                            )
                            _riss = _gd_rq(_res_df_qc, _RESIDUAL_COL_QC, _COMP_COL_QC)

                            _piss = []
                            if file_exists(INPUT_FILES["Prezzi"]):
                                try:
                                    _piss = _gd_pq(
                                        _res_df_qc, _RESIDUAL_COL_QC, _COMP_COL_QC,
                                        str(INPUT_FILES["Prezzi"]),
                                    )
                                except Exception as _pe:
                                    st.warning(f"Verifica prezzi non completata: {_pe}")

                            _map_df_qc = _gd_lmap(str(INPUT_FILES["Mappatura"]))
                            _acv_qc    = _gd_acv(str(INPUT_FILES["Take Rate"]), _map_df_qc)
                            _bom_qc    = pd.read_excel(
                                str(INPUT_FILES["BOM150"]), sheet_name="BOM150", header=2
                            )
                            _bom_qc = _bom_qc[
                                _bom_qc['Blocco dati per archivio dipendenze'].notnull()
                            ]
                            _parser_qc = _GD_CP(_map_df_qc, _acv_qc)
                            _trace_qc  = _gd_tr(
                                _bom_qc, 'Blocco dati per archivio dipendenze', _parser_qc
                            )
                            _gd_wr(_riss, _trace_qc, _piss,
                                   str(OUTPUT_DIR / "quality_check_report.xlsx"))
                            st.success("Quality check completato — scarica il report o apri l'anteprima.")
                            st.rerun()
                        except Exception as _e:
                            st.error(f"Errore durante quality check: {_e}")


            # ════════════════════════════════════════════════════════════════════════
            # SUB-TAB: AFFIDABILITÀ
            # ════════════════════════════════════════════════════════════════════════

        elif _ss_sub == "Affidabilità":

            st.markdown("## Guida alla Scelta dell'Affidabilità")

            # ── Spiegazione plain-language ────────────────────────────────────────────────────
            st.info(
                "**Cos'è il livello di affidabilità?**  \n"
                "Nella simulazione Monte Carlo, ogni Take Rate (TR) viene campionato "
                "da una distribuzione Dirichlet. Il **moltiplicatore** controlla quanto "
                "le estrazioni sono fedeli al valore storico:  \n"
                "- **Moltiplicatore basso** → alta variabilità → la simulazione esplora scenari distanti dal TR storico  \n"
                "- **Moltiplicatore alto** → bassa variabilità → la simulazione rimane vicina ai valori reali  \n\n"
                "Scegli un livello alto se ti fidi del dato storico. "
                "Scegli un livello basso se vuoi testare la robustezza del sistema a scenari alternativi."
            )

            # ── Configurazione corrente ─────────────────────────────────────────────────────
            _sd_aff_labels = {
                "NULLA":  ("🔴", "Alta incertezza — scenari molto distanti dal TR storico. "
                           "Usa solo per stress test."),
                "BASSA":  ("🟠", "Incertezza moderata — buona esplorazione degli scenari."),
                "MEDIA":  ("🟡", "Equilibrio tra fedeltà storica e varianza. Consigliato per analisi ordinarie."),
                "ALTA":   ("🟢", "Massima fedeltà al TR storico — minima varianza. "
                           "Consigliato per pianificazione operativa."),
            }
            _ci_m = _sd_aff_labels.get(aff_model,    ("", aff_model))
            _ci_o = _sd_aff_labels.get(aff_optional, ("", aff_optional))

            st.markdown("### Configurazione Attuale")
            _cfg_c1, _cfg_c2 = st.columns(2)
            with _cfg_c1:
                st.markdown(
                    f"**Model Mix (L0):** {_ci_m[0]} **{aff_model}**  \n"
                    f"*{_ci_m[1]}*"
                )
            with _cfg_c2:
                st.markdown(
                    f"**Caratteristiche / Optional (L1):** {_ci_o[0]} **{aff_optional}**  \n"
                    f"*{_ci_o[1]}*"
                )

            # ── Confronto visivo tra i 4 livelli ─────────────────────────────────────────────────────
            st.markdown("### Confronto tra i Livelli di Affidabilità")
            st.caption(
                "Impatto calcolato su un TR del 10% con domanda mensile configurabile. "
                "Valori più bassi di CV e Safety Stock indicano minor incertezza."
            )

            _sens_path = OUTPUT_DIR / "sensitivity_dirichlet.xlsx"

            _sd_demand_col, _sd_pct_col = st.columns([2, 1])
            _monthly_demand_sens = _sd_demand_col.number_input(
                "Domanda mensile di riferimento (moto/mese)",
                min_value=50, max_value=5000, value=500, step=50,
                key="sens_demand",
                help="Usata per calcolare le unità fisiche di safety stock",
            )
            _percentile_sens = _sd_pct_col.selectbox(
                "Percentile",
                options=[90, 95, 99],
                index=2,
                key="sens_pct",
            )

            try:
                from sensitivity_dirichlet import (
                    compute_metrics as _sd_cm,
                    MULTIPLIERS     as _SD_MULT,
                    TR_VALUES       as _SD_TR,
                )

                _z_map_sens = {90: 1.282, 95: 1.645, 99: 2.326}
                _z_sens = _z_map_sens.get(_percentile_sens, 2.326)

                # Schede riepilogative per i 4 livelli (TR fisso al 10%)
                _tr_ref = 0.10
                _card_cols = st.columns(4)
                _lv_order = ["NULLA", "BASSA", "MEDIA", "ALTA"]
                _lv_icons = {"NULLA": "🔴", "BASSA": "🟠", "MEDIA": "🟡", "ALTA": "🟢"}
                _lv_bg    = {
                    "NULLA": "#4a1a1a", "BASSA": "#4a3a1a",
                    "MEDIA": "#3a3a1a", "ALTA": "#1a4a1a",
                }
                for _ci, _lv in enumerate(_lv_order):
                    _M = _SD_MULT[_lv]
                    _m = _sd_cm(_tr_ref, _M, monthly_demand=_monthly_demand_sens)
                    _ss_u = round(_z_sens * _monthly_demand_sens * ((_tr_ref * (1 - _tr_ref) / _M) ** 0.5), 0)
                    _is_current = (aff_model == _lv or aff_optional == _lv)
                    _border = "border: 2px solid #E3000F;" if _is_current else ""
                    with _card_cols[_ci]:
                        st.markdown(
                            f'<div style="background:{_lv_bg[_lv]};border-radius:10px;'
                            f'padding:0.8rem 1rem;{_border}">'
                            f'<div style="font-size:1.3rem;font-weight:800;color:#fff">'
                            f'{_lv_icons[_lv]} {_lv}</div>'
                            f'<div style="color:#aaa;font-size:0.78rem;margin:0.3rem 0">'
                            f'M = {_M:,}</div>'
                            f'<hr style="border-color:#444;margin:0.4rem 0">'
                            f'<div style="color:#fff;font-size:0.88rem">'
                            f'CV: <b>{_m["cv"]:.1f}%</b></div>'
                            f'<div style="color:#fff;font-size:0.88rem">'
                            f'TR range (P{_percentile_sens}): '
                            f'<b>{max(0.0, _tr_ref*100 - _z_sens*_m["sigma_tr"]):.1f}%'
                            f' – {min(100.0, _tr_ref*100 + _z_sens*_m["sigma_tr"]):.1f}%</b></div>'
                            f'<div style="color:#fff;font-size:0.88rem">'
                            f'SS P{_percentile_sens}: <b>{int(_ss_u)} u.</b></div>'
                            f'<div style="color:#aaa;font-size:0.75rem;margin-top:0.3rem">'
                            f'{_m["label"]}</div>'
                            + ('<div style="color:#E3000F;font-size:0.75rem;font-weight:700">'
                               'IN USO</div>' if _is_current else '')
                            + '</div>',
                            unsafe_allow_html=True,
                        )

                st.caption(
                    f"Valori calcolati per TR = {_tr_ref*100:.0f}%, "
                    f"domanda mensile = {_monthly_demand_sens} moto/mese, "
                    f"percentile P{_percentile_sens}."
                )

                # ── Tabelle pivot per TR × Livello ───────────────────────────────────────────────
                with st.expander("📊 Tabelle di dettaglio (tutti i TR × tutti i livelli)", expanded=False):
                    # Raccolta dati grezzi
                    _sens_rows = []
                    for _tr in _SD_TR:
                        for _lv, _M in _SD_MULT.items():
                            _m = _sd_cm(_tr, _M, monthly_demand=_monthly_demand_sens)
                            _ss_pct = round(
                                _z_sens * _monthly_demand_sens
                                * ((_tr * (1.0 - _tr) / _M) ** 0.5),
                                1,
                            )
                            _sens_rows.append({
                                "TR storico":  f"{_tr*100:.0f}%",
                                "Livello":     _lv,
                                "SS":          int(round(_ss_pct)),
                                "CV":          round(_m['cv'], 1),
                                "sigma_tr":    _m['sigma_tr'],
                                "TR_pct":      _tr * 100,
                                "Incertezza":  _m['label'],
                            })
                    _df_raw = pd.DataFrame(_sens_rows)
                    _lv_order_cols = ["NULLA", "BASSA", "MEDIA", "ALTA"]
                    _lv_headers    = {
                        "NULLA": "🔴 NULLA",
                        "BASSA": "🟠 BASSA",
                        "MEDIA": "🟡 MEDIA",
                        "ALTA":  "🟢 ALTA",
                    }

                    # ── PIVOT 1: Safety Stock (unità buffer) ───────────────────────
                    st.markdown(f"#### Safety Stock necessario — P{_percentile_sens} (unità per {_monthly_demand_sens} moto/mese)")
                    st.caption("Quante unità extra devi tenere in magazzino per coprire l'incertezza della simulazione.")

                    _piv_ss = _df_raw.pivot(index="TR storico", columns="Livello", values="SS")
                    _piv_ss = _piv_ss.reindex(columns=_lv_order_cols)
                    _piv_ss.columns = [_lv_headers[c] for c in _piv_ss.columns]
                    _piv_ss.index.name = "TR storico"

                    # Colora le celle per valore assoluto (basso=verde, alto=rosso)
                    def _style_ss(val):
                        try:
                            v = int(val)
                            if v <= 2:   return "background-color:#C6EFCE; color:#1a4a1a"
                            elif v <= 8: return "background-color:#FFEB9C; color:#7F6000"
                            elif v <= 20: return "background-color:#FFCC99; color:#7F3F00"
                            else:        return "background-color:#FFC7CE; color:#C00000"
                        except Exception:
                            return ""

                    st.dataframe(
                        _piv_ss.style.applymap(_style_ss),
                        use_container_width=True,
                        height=38 + 35 * len(_piv_ss),
                    )

                    st.markdown("---")

                    # ── PIVOT 2: Range TR simulato ─────────────────────────────────
                    st.markdown(f"#### Range del Take Rate nella simulazione — P{_percentile_sens}")
                    st.caption(f"In che intervallo varia il TR durante i run Monte Carlo (P{_percentile_sens} — stesso livello di servizio scelto).")

                    _piv_rng = _df_raw.copy()
                    # Range calcolato con lo stesso z del percentile selezionato
                    _piv_rng["Range"] = _piv_rng.apply(
                        lambda r: (
                            f"{max(0.0, r['TR_pct'] - _z_sens * r['sigma_tr']):.1f}%"
                            f" – {min(100.0, r['TR_pct'] + _z_sens * r['sigma_tr']):.1f}%"
                        ), axis=1
                    )
                    _piv_range = _piv_rng.pivot(index="TR storico", columns="Livello", values="Range")
                    _piv_range = _piv_range.reindex(columns=_lv_order_cols)
                    _piv_range.columns = [_lv_headers[c] for c in _piv_range.columns]
                    _piv_range.index.name = "TR storico"

                    st.dataframe(
                        _piv_range,
                        use_container_width=True,
                        height=38 + 35 * len(_piv_range),
                    )

                    st.markdown("---")

                    # ── Legenda ────────────────────────────────────────────────────
                    st.markdown("**Legenda colori Safety Stock:**")
                    _tsl1, _tsl2, _tsl3, _tsl4 = st.columns(4)
                    _tsl1.markdown('<span style="background:#C6EFCE;padding:3px 10px;border-radius:4px">'
                                   '🟢 <b>≤ 2 u.</b> — Stabile</span>', unsafe_allow_html=True)
                    _tsl2.markdown('<span style="background:#FFEB9C;padding:3px 10px;border-radius:4px">'
                                   '🟡 <b>3–8 u.</b> — Moderato</span>', unsafe_allow_html=True)
                    _tsl3.markdown('<span style="background:#FFCC99;padding:3px 10px;border-radius:4px">'
                                   '🟠 <b>9–20 u.</b> — Alto</span>', unsafe_allow_html=True)
                    _tsl4.markdown('<span style="background:#FFC7CE;padding:3px 10px;border-radius:4px">'
                                   '🔴 <b>&gt; 20 u.</b> — Critico</span>', unsafe_allow_html=True)

            except Exception as _e:
                st.error(f"Errore generazione analisi sensibilità: {_e}")

            # ── Pulsanti rigenera / scarica ────────────────────────────────────────────────
            _sens_c1, _sens_c2 = st.columns(2)
            with _sens_c1:
                _run_sens = st.button(
                    "Rigenera Excel Sensibilità",
                    key="btn_sens",
                    help="Ricalcola e salva Output/sensitivity_dirichlet.xlsx con tutti i fogli esplicativi",
                )
                if _run_sens:
                    with st.spinner("Generazione Excel in corso..."):
                        try:
                            from sensitivity_dirichlet import main as _sd_main
                            _sd_main()
                            st.success("File salvato: Output/sensitivity_dirichlet.xlsx")
                            st.rerun()
                        except Exception as _e:
                            st.error(f"Errore: {_e}")
            with _sens_c2:
                if _sens_path.exists():
                    _smtime = time.strftime(
                        "%d/%m/%Y %H:%M", time.localtime(_sens_path.stat().st_mtime)
                    )
                    st.caption(f"Ultimo aggiornamento: {_smtime}")
                    with open(_sens_path, "rb") as _f:
                        st.download_button(
                            "Scarica Excel Sensibilità",
                            data=_f.read(),
                            file_name="sensitivity_dirichlet.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="dl_sens",
                        )

            # ════════════════════════════════════════════════════════════════════════
            # SUB-TAB: AVVIA PIPELINE
            # ════════════════════════════════════════════════════════════════════════

        elif _ss_sub == "Avvia Pipeline":

            st.markdown("## Avvia Pipeline")

            # ── Parametri ────────────────────────────────────────────────────────
            st.markdown("### Parametri")
            _pr1, _pr2, _pr3 = st.columns(3)

            with _pr1:
                st.markdown("**Simulazione**")
                st.number_input(
                    "Numero di Run MC",
                    min_value=10, max_value=5000,
                    value=st.session_state.get("ss_n_runs", 200), step=1,
                    help="Più run = statistiche più stabili, ma più tempo di calcolo",
                    key="ss_n_runs",
                )
                _pct_av = st.slider(
                    "Livello di servizio (%)",
                    min_value=50, max_value=99,
                    value=st.session_state.get("ss_percentile", 99), step=1,
                    help="Percentile per il calcolo della Safety Stock",
                    key="ss_percentile",
                )
                st.caption(f"Safety Stock = quantile {_pct_av}° − media domanda")

            with _pr2:
                st.markdown("**Mese & Steps**")
                if demand_months:
                    _def_month_av = st.session_state.get("ss_start_month", demand_months[0])
                    _def_month_av = _def_month_av if _def_month_av in demand_months else demand_months[0]
                    _sm_av = st.selectbox(
                        "Mese di inizio elaborazione",
                        options=demand_months,
                        index=demand_months.index(_def_month_av),
                        help="Primo mese da includere nel calcolo",
                        key="ss_start_month",
                    )
                    st.session_state["ss_start_offset"] = demand_months.index(_sm_av)
                else:
                    st.warning("Carica `Total_demand.xlsx` per selezionare il mese.")
                    st.number_input(
                        "Start Month Offset",
                        min_value=0, max_value=24,
                        value=st.session_state.get("ss_start_offset", 0), step=1,
                        key="ss_start_offset",
                    )
                st.markdown("**Step Pipeline**")
                st.checkbox(
                    "Step 1 · Rigenera Matrice BOM",
                    value=st.session_state.get("ss_step0", False),
                    help="Rielabora BOM150 → matrix_output_V2.xlsx (solo se BOM è cambiato)",
                    key="ss_step0",
                )
                st.checkbox("Step 2 · Simulazione Monte Carlo",
                    value=st.session_state.get("ss_step2", True), key="ss_step2")
                st.checkbox("Step 3 · Matching SKU & Safety Stock",
                    value=st.session_state.get("ss_step3", True), key="ss_step3")

            with _pr3:
                st.markdown("**Affidabilità Take Rate**")
                st.caption("Moltiplicatore Dirichlet: più alto = più fedele al TR storico")
                st.selectbox(
                    "Model Mix (L0)",
                    options=_AFF_LEVELS,
                    index=_AFF_LEVELS.index(st.session_state.get("ss_aff_model", _aff_def_model)),
                    help=f"Livelli: {', '.join(f'{k}={int(v)}' for k,v in sorted(_AFF_MAP.items(), key=lambda x: x[1]))}",
                    key="ss_aff_model",
                )
                st.selectbox(
                    "Caratteristiche / Optional (L1)",
                    options=_AFF_LEVELS,
                    index=_AFF_LEVELS.index(st.session_state.get("ss_aff_optional", _aff_def_optional)),
                    help=f"Livelli: {', '.join(f'{k}={int(v)}' for k,v in sorted(_AFF_MAP.items(), key=lambda x: x[1]))}",
                    key="ss_aff_optional",
                )
                if st.button("Aggiorna RAM", key="ram_btn_av"):
                    st.session_state["last_ram"] = get_memory_mb()
                st.caption(f"RAM processo: {st.session_state.get('last_ram', get_memory_mb()):.0f} MB")

            st.markdown("---")
            st.markdown("**Qualità Dati BOM**")
            st.checkbox(
                "Escludi condizioni CT/CV non tracciati (Step 1)",
                value=st.session_state.get("ss_filter_untraced", True),
                key="ss_filter_untraced",
                help=(
                    "Se attivo, durante la costruzione della matrice BOM i predicati "
                    "che referenziano CT o CV non presenti nella Mappatura vengono rimossi. "
                    "Se rimane condizione vuota, il componente è escluso dall'analisi. "
                    "Disattiva solo se vuoi il comportamento originale (predicati non tracciati "
                    "vengono ignorati silenziosamente)."
                ),
            )

            st.markdown("---")
            st.markdown("### Avvia")

            _miss_ss_run = [n for n in SS_REQUIRED
                            if not file_exists(INPUT_FILES[n]) and n not in ("Matrix", "Unique Rows")]

            _av_c1, _av_c2 = st.columns([1, 2])
            with _av_c1:
                if st.button(
                    "AVVIA SAFETY STOCK",
                    type="primary",
                    disabled=bool(_miss_ss_run),
                    key="btn_ss",
                    use_container_width=True,
                ):
                    st.session_state["_do_run_ss"] = True
                    st.rerun()
            with _av_c2:
                if _miss_ss_run:
                    st.warning(f"Impossibile avviare: mancano {', '.join(_miss_ss_run)}.")
                else:
                    _sv_pc1, _sv_pc2, _sv_pc3, _sv_pc4 = st.columns(4)
                    _sv_pc1.metric("Run MC",      st.session_state.get("ss_n_runs", 200))
                    _sv_pc2.metric("Mese",        st.session_state.get("ss_start_month", demand_months[0] if demand_months else f"Offset {start_month_offset}"))
                    _sv_pc3.metric("Percentile",  f"P{st.session_state.get('ss_percentile', 99)}")
                    _sv_pc4.metric("Affidabilità", f"{st.session_state.get('ss_aff_model', aff_model)}/{st.session_state.get('ss_aff_optional', aff_optional)}")

    # SS — TAB 2: VISUALIZZA RISULTATI
    # ########################################################################

    with tab_view:

        # Sub-tab interni
        st_analytics, st_results, st_monthly, st_tr = st.tabs([
            "Analisi Integrata",
            "Risultati Safety Stock",
            "Safety Mensile",
            "Take Rate",
        ])

        # ========================================================================
        # SUB-TAB A — ANALISI INTEGRATA
        # ========================================================================

        with st_analytics:
            st.subheader("Analisi Integrata Safety Stock")
            st.caption(
                "Dashboard che unisce Safety Stock e Cycle Stock per visione "
                "Strategica, Tattica e Operativa."
            )

            # ── Selettore versione da analizzare ────────────────────────────────
            _ss_path  = OUTPUT_FILES["Safety Stock"]
            _cmp_path = OUTPUT_FILES["Demand Comparison"]
            _hist_dir_an = OUTPUT_DIR / "history"
            _hist_files_an = sorted(
                [f for f in (_hist_dir_an.glob("*.xlsx") if _hist_dir_an.exists() else [])],
                key=lambda p: p.stat().st_mtime, reverse=True,
            )
            _ver_opts_an = []
            if file_exists(_ss_path):
                _ver_opts_an.append("Versione attuale")
            _ver_opts_an += [f.name for f in _hist_files_an]

            _selected_ver_an = "Versione attuale"
            if _ver_opts_an:
                _selected_ver_an = st.selectbox(
                    "Versione da analizzare",
                    options=_ver_opts_an,
                    index=0,
                    key="an_version_sel",
                    help="Scegli quale versione di Safety Stock vuoi visualizzare nei KPI e grafici",
                )
            else:
                st.info(
                    "Nessun dato Safety Stock disponibile.  \n"
                    "Esegui prima la pipeline **Safety Stock** dalla tab ▶️."
                )

            # Carica dati
            _df_ss  = None
            _df_cmp = None

            if _selected_ver_an == "Versione attuale":
                if file_exists(_ss_path):
                    try:
                        _df_ss = load_safety_stock(str(_ss_path))
                    except Exception as _e:
                        st.error(f"Errore lettura Safety Stock: {_e}")
            elif _selected_ver_an:
                _hist_path_an = _hist_dir_an / _selected_ver_an
                if _hist_path_an.exists():
                    try:
                        _df_ss = pd.read_excel(str(_hist_path_an))
                        st.caption(f"Stai visualizzando la versione storica: **{_selected_ver_an}**")
                    except Exception as _e:
                        st.error(f"Errore lettura versione storica: {_e}")

            if file_exists(_cmp_path):
                try:
                    _raw = load_comparison(str(_cmp_path))
                    if "Con Prezzo" in _raw:
                        _df_cmp = pd.concat(
                            [_raw["Con Prezzo"], _raw.get("Senza Prezzo", pd.DataFrame())],
                            ignore_index=True,
                        )
                    else:
                        _df_cmp = next(iter(_raw.values()))
                except Exception as _e:
                    st.error(f"Errore lettura Cycle Stock: {_e}")

            # Check sessione per Cycle Stock
            if _df_cmp is None and "df_comparison" in st.session_state:
                _df_cmp = st.session_state["df_comparison"]

            if _df_ss is None:
                st.info(
                    "Dati insufficienti.  \n"
                    "Esegui prima la pipeline **Safety Stock** dalla tab ▶️.  \n"
                    "Per scatter e Δ economico esegui anche **Cycle Stock**."
                )
            else:
                # ── Join ────────────────────────────────────────────────────────
                if _df_cmp is not None:
                    _cmp_slim = _df_cmp[[c for c in (
                        "SKU", "delta_demand", "mean_demand_actual",
                        "mean_demand_forecast", "delta_prezzo", "Prezzo netto",
                    ) if c in _df_cmp.columns]].copy()
                    df_m = _df_ss.merge(_cmp_slim, on="SKU", how="left")
                else:
                    df_m = _df_ss.copy()
                    for _c in ("delta_demand", "mean_demand_actual",
                               "mean_demand_forecast", "delta_prezzo", "Prezzo netto"):
                        df_m[_c] = np.nan

                # ── ABC (vettorializzata) ────────────────────────────────────────
                _has_prezzo = "prezzo_safety" in df_m.columns
                if _has_prezzo:
                    _sorted = (df_m[["SKU", "prezzo_safety"]].dropna()
                               .sort_values("prezzo_safety", ascending=False).copy())
                    _tot = _sorted["prezzo_safety"].sum()
                    if _tot > 0:
                        _sorted["_cp"] = _sorted["prezzo_safety"].cumsum() / _tot
                        _sorted["ABC"] = np.select(
                            [_sorted["_cp"] <= 0.80, _sorted["_cp"] <= 0.95],
                            ["A", "B"], default="C",
                        )
                        df_m = df_m.merge(_sorted[["SKU", "ABC"]], on="SKU", how="left")
                        df_m["ABC"] = df_m["ABC"].fillna("C")
                    else:
                        df_m["ABC"] = "C"
                else:
                    df_m["ABC"] = "N/D"

                # Prezzo unitario stimato
                if _has_prezzo and "safety_stock_p99_total" in df_m.columns:
                    df_m["Prezzo_Unitario_Stimato"] = np.where(
                        df_m["safety_stock_p99_total"] > 0,
                        df_m["prezzo_safety"] / df_m["safety_stock_p99_total"],
                        np.nan,
                    )

                # ── Priorità (vettorializzata) ───────────────────────────────────
                _abc_sc = df_m["ABC"].map({"A": 3, "B": 2, "C": 1}).fillna(1)
                _lt_col = df_m.get("Lead_Time_Months_Ceil", pd.Series(0, index=df_m.index))
                _lt_sc  = np.where(_lt_col > 4, 2, 1)
                _nc_col = df_m.get("N_Matching_Combinations",
                                   pd.Series(0, index=df_m.index)).fillna(0)
                _nc_sc  = np.select([_nc_col > 1000, _nc_col >= 100], [3, 2], default=1)
                _score  = _abc_sc.values + _lt_sc + _nc_sc
                df_m["Priorità"] = np.select(
                    [_score >= 7, _score >= 4], ["Alta", "Media"], default="Bassa"
                )

                # ── Filtri globali ───────────────────────────────────────────────
                with st.expander("Filtri", expanded=True):
                    fa, fb, fc = st.columns(3)
                    with fa:
                        g_search = st.text_input("Cerca SKU / Descrizione",
                                                 key="an_search", placeholder="es: RAGGIO, 4961...")
                    with fb:
                        _abc_vals = sorted(df_m["ABC"].dropna().unique().tolist())
                        g_abc = st.multiselect("Classe ABC", _abc_vals, default=_abc_vals, key="an_abc")
                    with fc:
                        if "Lead_Time_Months_Ceil" in df_m.columns:
                            _lt_min = int(df_m["Lead_Time_Months_Ceil"].min())
                            _lt_max = int(df_m["Lead_Time_Months_Ceil"].max())
                            g_lt = st.slider("Lead Time (mesi)", _lt_min, _lt_max,
                                             (_lt_min, _lt_max), key="an_lt")
                        else:
                            g_lt = (0, 99)

                df_f = df_m.copy()
                if g_search:
                    _tok = g_search.strip().lower()
                    df_f = df_f[
                        df_f["SKU"].astype(str).str.lower().str.contains(_tok, regex=False, na=False)
                        | df_f["Description"].astype(str).str.lower().str.contains(_tok, regex=False, na=False)
                    ]
                if g_abc:
                    df_f = df_f[df_f["ABC"].isin(g_abc)]
                if "Lead_Time_Months_Ceil" in df_f.columns:
                    df_f = df_f[df_f["Lead_Time_Months_Ceil"].between(g_lt[0], g_lt[1])]

                st.caption(f"**{len(df_f)} SKU** nel dataset filtrato (totale: {len(df_m)})")

                # ════════════════════════════════════════════════════════════════
                # SEZIONE 1 — STRATEGICA
                # ════════════════════════════════════════════════════════════════
                st.markdown("---")
                st.markdown("## Visione Strategica")

                kp2, = st.columns(1)

                with kp2:
                    if _has_prezzo:
                        _cap = df_f["prezzo_safety"].sum()
                        st.metric("Capitale Immobilizzato", f"€ {_cap:,.0f}",
                                  help="Valore totale del buffer di sicurezza (SKU filtrati)")
                    else:
                        st.metric("Capitale Immobilizzato", "N/D")

                st.markdown("---")
                str_c1, str_c2 = st.columns([3, 2])

                with str_c1:
                    st.markdown("#### Top 10 SKU per Costo Safety Stock")
                    if _has_prezzo:
                        _top10 = (df_f.nlargest(10, "prezzo_safety")
                                  [["SKU", "Description", "prezzo_safety", "ABC",
                                    "Lead_Time_Months_Ceil"]]
                                  .copy()
                                  .sort_values("prezzo_safety"))
                        _top10["Label"] = (
                            _top10["SKU"].astype(str) + "  "
                            + _top10["Description"].astype(str).str[:28]
                        )
                        fig_top10 = go.Figure(go.Bar(
                            x=_top10["prezzo_safety"], y=_top10["Label"],
                            orientation="h",
                            marker_color=[ABC_COLORS.get(str(a), "#666") for a in _top10["ABC"]],
                            text=[f"€ {v:,.0f}  [{a}]"
                                  for v, a in zip(_top10["prezzo_safety"], _top10["ABC"])],
                            textposition="outside",
                            hovertemplate="<b>%{y}</b><br>€ %{x:,.0f}<extra></extra>",
                        ))
                        fig_top10.update_layout(
                            height=380, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                            font_color="#eee", xaxis_title="€",
                            margin=dict(l=10, r=100, t=10, b=10),
                        )
                        st.plotly_chart(fig_top10, use_container_width=True)
                        st.caption(
                            "I 10 componenti con il maggiore valore economico di safety stock. "
                            "Il colore indica la classe ABC: rosso = A (impatto alto), arancione = B, blu = C (impatto basso). "
                            "Su questi SKU conviene concentrare l'attenzione per ridurre il capitale immobilizzato."
                        )
                    else:
                        st.info("Colonna prezzo_safety non disponibile.")

                with str_c2:
                    st.markdown("#### Riepilogo per Classe ABC")
                    if "ABC" in df_f.columns and _has_prezzo:
                        _agg_dict = {
                            "N_SKU":    ("SKU", "count"),
                            "Costo_SS": ("prezzo_safety", "sum"),
                        }
                        if "mean_demand" in df_f.columns:
                            _agg_dict["Dom_Media"] = ("mean_demand", "mean")
                        _abc_summary = df_f.groupby("ABC").agg(**_agg_dict).reset_index()
                        _abc_summary["Incidenza %"] = (
                            _abc_summary["Costo_SS"] / _abc_summary["Costo_SS"].sum() * 100
                        ).round(1).astype(str) + " %"
                        _abc_summary["Costo_SS"] = _abc_summary["Costo_SS"].map("€ {:,.0f}".format)
                        st.dataframe(safe_df(_abc_summary), use_container_width=True,
                                     hide_index=True, height=200)

                        _abc_pie = df_f.groupby("ABC")["prezzo_safety"].sum().reset_index()
                        fig_pie = px.pie(
                            _abc_pie, names="ABC", values="prezzo_safety",
                            color="ABC", color_discrete_map=ABC_COLORS, hole=0.45,
                        )
                        fig_pie.update_traces(textinfo="label+percent")
                        fig_pie.update_layout(
                            paper_bgcolor="#0d1117", font_color="#eee",
                            margin=dict(t=10, b=0, l=0, r=0), height=240, showlegend=False,
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)
                        st.caption(
                            "Ripartizione del capitale totale immobilizzato in safety stock per classe ABC. "
                            "La classe A comprende il 20% degli SKU che genera l'80% del valore: "
                            "una fetta grande indica che pochi componenti costosi dominano il buffer."
                        )

                # ── Curva di Pareto ──────────────────────────────────────────────
                st.markdown("---")
                st.markdown("#### Curva di Pareto — Concentrazione del Capitale")
                if _has_prezzo:
                    _par = (df_f[["SKU", "Description", "prezzo_safety", "ABC"]]
                            .dropna(subset=["prezzo_safety"])
                            .sort_values("prezzo_safety", ascending=False)
                            .reset_index(drop=True)
                            .copy())
                    if not _par.empty:
                        _par_tot = _par["prezzo_safety"].sum()
                        _par["_cum"]     = _par["prezzo_safety"].cumsum() / _par_tot * 100
                        _par["_idx"]     = range(1, len(_par) + 1)
                        _par["_tooltip"] = (
                            _par["SKU"].astype(str) + " – "
                            + _par["Description"].astype(str).str[:35]
                        )
                        _n80 = int((_par["_cum"] <= 80).sum()) + 1

                        fig_par = go.Figure()
                        fig_par.add_trace(go.Bar(
                            x=_par["_idx"],
                            y=_par["prezzo_safety"],
                            marker_color=[ABC_COLORS.get(str(a), "#666") for a in _par["ABC"]],
                            name="Costo SS (€)",
                            customdata=_par[["_tooltip", "ABC"]].values,
                            hovertemplate=(
                                "<b>%{customdata[0]}</b><br>"
                                "€ %{y:,.0f} · Classe %{customdata[1]}<extra></extra>"
                            ),
                            yaxis="y1",
                        ))
                        fig_par.add_trace(go.Scatter(
                            x=_par["_idx"],
                            y=_par["_cum"],
                            mode="lines",
                            name="Cumulata %",
                            line=dict(color="#ff9800", width=2.5),
                            yaxis="y2",
                            hovertemplate="SKU #%{x} → %{y:.1f}% cumulato<extra></extra>",
                        ))
                        fig_par.add_shape(
                            type="line", xref="paper",
                            x0=0, x1=1, y0=80, y1=80, yref="y2",
                            line=dict(color="#aaa", dash="dash", width=1.5),
                        )
                        fig_par.add_annotation(
                            xref="paper", x=0.01, y=83, yref="y2",
                            text="80 %", showarrow=False,
                            font=dict(color="#aaa", size=11),
                        )
                        fig_par.update_layout(
                            height=360,
                            plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                            font_color="#eee",
                            xaxis=dict(
                                title="SKU (ordinati per costo decrescente)",
                                showticklabels=False, gridcolor="#222",
                            ),
                            yaxis=dict(title="Costo Safety Stock (€)", gridcolor="#222"),
                            yaxis2=dict(
                                title=dict(text="Cumulata %", font=dict(color="#ff9800")),
                                overlaying="y", side="right",
                                range=[0, 105],
                                ticksuffix="%",
                                tickfont=dict(color="#ff9800"),
                                showgrid=False,
                            ),
                            legend=dict(orientation="h", y=1.06),
                            margin=dict(t=30, b=30),
                            bargap=0.02,
                        )
                        st.plotly_chart(fig_par, use_container_width=True)
                        st.caption(
                            f"**{_n80} SKU** ({_n80 / len(_par) * 100:.0f}%) concentrano l'80% del capitale "
                            f"immobilizzato (€ {_par_tot * 0.8:,.0f} su € {_par_tot:,.0f} totali). "
                            "La curva arancione mostra la percentuale cumulata; "
                            "le barre sono colorate per classe ABC (rosso = A, arancione = B, blu = C). "
                            "Concentrare l'attenzione sui primi SKU permette il massimo impatto con il minimo sforzo."
                        )
                    else:
                        st.info("Nessun dato disponibile per il Pareto.")
                else:
                    st.info("Pareto non disponibile: esegui Step 3 con prezzi per ottenere `prezzo_safety`.")

                # ════════════════════════════════════════════════════════════════
                # SEZIONE 2 — TATTICA
                # ════════════════════════════════════════════════════════════════
                st.markdown("---")
                st.markdown("## Visione Tattica")

                st.markdown("#### Costo SS per Lead Time")
                if "Lead_Time_Months_Ceil" in df_f.columns and _has_prezzo:
                    _lt_grp = (df_f.groupby("Lead_Time_Months_Ceil")["prezzo_safety"]
                               .agg(Totale="sum", N="count").reset_index())
                    fig_lt = go.Figure(go.Bar(
                        x=_lt_grp["Lead_Time_Months_Ceil"].astype(str),
                        y=_lt_grp["Totale"],
                        marker_color="#E3000F",
                        text=_lt_grp["N"].apply(lambda n: f"{n} SKU"),
                        textposition="outside",
                        hovertemplate="LT %{x} mesi<br>€ %{y:,.0f}<extra></extra>",
                    ))
                    fig_lt.update_layout(
                        height=260, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                        font_color="#eee", xaxis_title="Lead Time (mesi)", yaxis_title="€",
                        margin=dict(t=10, b=10),
                    )
                    st.plotly_chart(fig_lt, use_container_width=True)
                    st.caption(
                        "Capitale totale di safety stock raggruppato per lead time del fornitore (in mesi). "
                        "Lead time più lunghi richiedono buffer più grandi: le barre alte a destra "
                        "indicano categorie di fornitori su cui si concentra il rischio di immobilizzo."
                    )

                st.markdown("#### ABC x Lead Time")
                if "ABC" in df_f.columns and "Lead_Time_Months_Ceil" in df_f.columns and _has_prezzo:
                    _pivot = (df_f.groupby(["ABC", "Lead_Time_Months_Ceil"])
                              .agg(N=("SKU", "count"), Costo_SS=("prezzo_safety", "sum"))
                              .reset_index())
                    _pivot["Costo_SS"] = _pivot["Costo_SS"].map("€ {:,.0f}".format)
                    _pivot.columns = ["Classe", "LT (mesi)", "N SKU", "Costo SS"]
                    st.dataframe(safe_df(_pivot), use_container_width=True,
                                 hide_index=True, height=280)


                # ── Scatter Efficienza: Quota Domanda vs Quota Buffer ────────────
                st.markdown("---")
                st.markdown("#### Efficienza Buffer: Quota Domanda vs Quota Safety Stock")
                _dem_col_eff = next(
                    (c for c in ("mean_demand_actual", "mean_demand_forecast")
                     if c in df_f.columns and df_f[c].notna().any()),
                    None,
                )
                if _dem_col_eff and _has_prezzo:
                    _eff = (df_f[["SKU", "Description", "prezzo_safety", "ABC", _dem_col_eff]]
                            .dropna().copy())
                    _tot_ss_e  = _eff["prezzo_safety"].sum()
                    _tot_dem_e = _eff[_dem_col_eff].sum()
                    if not _eff.empty and _tot_ss_e > 0 and _tot_dem_e > 0:
                        _eff["_pct_ss"]  = _eff["prezzo_safety"]     / _tot_ss_e  * 100
                        _eff["_pct_dem"] = _eff[_dem_col_eff]         / _tot_dem_e * 100
                        _eff["_tooltip"] = (
                            _eff["SKU"].astype(str) + " – "
                            + _eff["Description"].astype(str).str[:40]
                        )
                        _ax_max = max(
                            float(_eff["_pct_ss"].quantile(0.97)),
                            float(_eff["_pct_dem"].quantile(0.97)),
                        ) * 1.20

                        fig_eff = px.scatter(
                            _eff,
                            x="_pct_dem", y="_pct_ss",
                            color="ABC", color_discrete_map=ABC_COLORS,
                            custom_data=["_tooltip", "ABC"],
                            labels={
                                "_pct_dem": "% Domanda totale",
                                "_pct_ss":  "% Costo Safety Stock totale",
                            },
                        )
                        fig_eff.update_traces(
                            marker=dict(size=8, opacity=0.80,
                                        line=dict(width=0.5, color="#fff")),
                            hovertemplate=(
                                "<b>%{customdata[0]}</b><br>"
                                "Quota domanda: %{x:.2f}%<br>"
                                "Quota SS: %{y:.2f}%<br>"
                                "Classe: %{customdata[1]}<extra></extra>"
                            ),
                        )
                        # Bisettrice y = x
                        fig_eff.add_shape(
                            type="line", x0=0, y0=0, x1=_ax_max, y1=_ax_max,
                            line=dict(color="#aaa", dash="dash", width=1.5),
                        )
                        # Annotazioni quadranti
                        _q = _ax_max * 0.65
                        for _lbl, _qx, _qy in [
                            ("Over-buffered", _q * 0.22, _q * 0.78),
                            ("Under-buffered", _q * 0.78, _q * 0.22),
                        ]:
                            fig_eff.add_annotation(
                                x=_qx, y=_qy, text=_lbl, showarrow=False,
                                font=dict(size=11, color="#ccc"),
                                bgcolor="rgba(30,30,60,0.75)",
                                bordercolor="#555", borderpad=4,
                            )
                        fig_eff.update_layout(
                            height=480,
                            plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                            font_color="#eee",
                            xaxis=dict(gridcolor="#222", range=[0, _ax_max]),
                            yaxis=dict(gridcolor="#222", range=[0, _ax_max]),
                            legend_title="Classe ABC",
                            margin=dict(t=30, b=40),
                        )
                        st.plotly_chart(fig_eff, use_container_width=True)
                        st.caption(
                            "Ogni punto è uno SKU. "
                            "**Asse X** = % della domanda totale · **Asse Y** = % del costo totale di safety stock. "
                            "La linea tratteggiata è la bisettrice: un punto su di essa ha un buffer perfettamente "
                            "proporzionale alla sua domanda. "
                            "**Sopra** la bisettrice = capitale immobilizzato sproporzionato rispetto alla domanda "
                            "(candidato a riduzione affidabilità TR o lead time). "
                            "**Sotto** = buffer potenzialmente insufficiente rispetto al peso sulla domanda."
                        )
                    else:
                        st.info("Dati insufficienti per lo scatter efficienza.")
                else:
                    st.info(
                        "Scatter efficienza non disponibile: "
                        "esegui **Cycle Stock** per ottenere i dati di domanda per SKU."
                    )

                # ════════════════════════════════════════════════════════════════
                # SEZIONE 4 — CONFRONTO VERSIONI STORICHE
                # ════════════════════════════════════════════════════════════════
                st.markdown("---")
                st.markdown("## Confronto Versioni")
                st.caption(
                    "Seleziona due versioni archiviate e confronta la Safety Stock SKU per SKU. "
                    "Utile per valutare l'impatto di diversi livelli di affidabilità TR o di run differenti."
                )

                _hist_dir_a = OUTPUT_DIR / "history"
                _hist_files_a = sorted(_hist_dir_a.glob("*.xlsx"), reverse=True) \
                    if _hist_dir_a.exists() else []

                # Aggiungi la versione corrente come opzione
                _curr_ss = OUTPUT_FILES["Safety Stock"]
                _ver_options = []
                if _curr_ss.exists():
                    _ver_options.append("Versione attuale")
                _ver_options += [f.name for f in _hist_files_a]

                if len(_ver_options) < 2:
                    st.info(
                        "Servono almeno due versioni per il confronto.  \n"
                        "Esegui la pipeline Safety Stock almeno due volte (anche con parametri diversi)."
                    )
                else:
                    _cmp_c1, _cmp_c2 = st.columns(2)
                    with _cmp_c1:
                        _ver_a = st.selectbox("Versione A (base)", _ver_options,
                                              index=0, key="ver_cmp_a")
                    with _cmp_c2:
                        _ver_b = st.selectbox("Versione B (confronto)",
                                              _ver_options,
                                              index=min(1, len(_ver_options) - 1),
                                              key="ver_cmp_b")

                    if st.button("Esegui confronto", key="btn_ver_compare"):
                        def _load_ver(label):
                            if label == "Versione attuale":
                                return pd.read_excel(str(_curr_ss))
                            else:
                                return pd.read_excel(str(_hist_dir_a / label))

                        try:
                            _dfa = _load_ver(_ver_a)
                            _dfb = _load_ver(_ver_b)

                            _ss_col_a = next(
                                (c for c in _dfa.columns
                                 if "safety_stock_p" in c.lower() and "total" in c.lower()), None
                            )
                            _ss_col_b = next(
                                (c for c in _dfb.columns
                                 if "safety_stock_p" in c.lower() and "total" in c.lower()), None
                            )

                            if not _ss_col_a or not _ss_col_b or "SKU" not in _dfa.columns:
                                st.error("Colonne richieste (SKU, safety_stock_p*_total) non trovate.")
                            else:
                                _cols_a = ["SKU", "Description", _ss_col_a]
                                if "prezzo_safety" in _dfa.columns:
                                    _cols_a.append("prezzo_safety")
                                _cols_b = ["SKU", _ss_col_b]
                                if "prezzo_safety" in _dfb.columns:
                                    _cols_b.append("prezzo_safety")

                                _ver_merged = _dfa[_cols_a].merge(
                                    _dfb[_cols_b].rename(
                                        columns={_ss_col_b: "SS_B",
                                                 "prezzo_safety": "prezzo_B"}
                                    ),
                                    on="SKU", how="outer"
                                ).fillna(0)

                                # Δ Safety Stock
                                _ver_merged["Δ SS assoluto"] = (
                                    _ver_merged[_ss_col_a] - _ver_merged["SS_B"]
                                ).round(2)
                                _ver_merged["Δ SS %"] = np.where(
                                    _ver_merged["SS_B"] != 0,
                                    (_ver_merged["Δ SS assoluto"] / _ver_merged["SS_B"] * 100).round(1),
                                    np.nan,
                                )

                                # Δ Economico
                                if "prezzo_safety" in _ver_merged.columns and "prezzo_B" in _ver_merged.columns:
                                    _ver_merged["Δ € assoluto"] = (
                                        _ver_merged["prezzo_safety"] - _ver_merged["prezzo_B"]
                                    ).round(2)
                                    _ver_merged["Δ € %"] = np.where(
                                        _ver_merged["prezzo_B"] != 0,
                                        (_ver_merged["Δ € assoluto"] / _ver_merged["prezzo_B"] * 100).round(1),
                                        np.nan,
                                    )

                                _ver_merged = _ver_merged.rename(columns={
                                    _ss_col_a:       f"SS — {_ver_a[:25]}",
                                    "SS_B":          f"SS — {_ver_b[:25]}",
                                    "prezzo_safety": f"€ — {_ver_a[:25]}",
                                    "prezzo_B":      f"€ — {_ver_b[:25]}",
                                }).sort_values("Δ SS assoluto", ascending=False, key=abs)

                                st.session_state["_ver_compare_df"]  = _ver_merged
                                st.session_state["_ver_compare_lbla"] = _ver_a
                                st.session_state["_ver_compare_lblb"] = _ver_b
                        except Exception as _vce:
                            st.error(f"Errore nel confronto: {_vce}")

                    # Mostra risultato confronto
                    if "_ver_compare_df" in st.session_state:
                        _vdf  = st.session_state["_ver_compare_df"]
                        _lbl_a = st.session_state.get("_ver_compare_lbla", "A")
                        _lbl_b = st.session_state.get("_ver_compare_lblb", "B")

                        st.markdown(f"**Confronto:** `{_lbl_a}` vs `{_lbl_b}`")

                        # Totali economici delle singole versioni
                        _col_prezzo_a = f"€ — {_lbl_a[:25]}"
                        _col_prezzo_b = f"€ — {_lbl_b[:25]}"
                        if _col_prezzo_a in _vdf.columns and _col_prezzo_b in _vdf.columns:
                            _tot_a = _vdf[_col_prezzo_a].sum()
                            _tot_b = _vdf[_col_prezzo_b].sum()
                            _vt1, _vt2, _vt3 = st.columns(3)
                            _vt1.metric(f"Totale € Versione A", f"{_tot_a:,.0f} €",
                                        help=f"Valore economico totale Safety Stock: {_lbl_a}")
                            _vt2.metric(f"Totale € Versione B", f"{_tot_b:,.0f} €",
                                        help=f"Valore economico totale Safety Stock: {_lbl_b}")
                            _vt3.metric(f"Δ Totale (A − B)", f"{_tot_a - _tot_b:+,.0f} €",
                                        delta_color="inverse" if (_tot_a - _tot_b) > 0 else "normal",
                                        help="Differenza tra i totali economici delle due versioni")

                        _vc1, _vc2, _vc3, _vc4 = st.columns(4)
                        _vc1.metric("SKU totali", len(_vdf))
                        _vc2.metric("A > B (SS aumentata)", int((_vdf["Δ SS assoluto"] > 0).sum()))
                        _vc3.metric("A < B (SS ridotta)",   int((_vdf["Δ SS assoluto"] < 0).sum()))
                        _vc4.metric("Invariati",             int((_vdf["Δ SS assoluto"] == 0).sum()))

                        # Metriche economiche delta se disponibili
                        if "Δ € assoluto" in _vdf.columns:
                            _ve1, _ve2, _ve3 = st.columns(3)
                            _delta_eur_tot = _vdf["Δ € assoluto"].sum()
                            _delta_eur_pos = _vdf.loc[_vdf["Δ € assoluto"] > 0, "Δ € assoluto"].sum()
                            _delta_eur_neg = _vdf.loc[_vdf["Δ € assoluto"] < 0, "Δ € assoluto"].sum()
                            _ve1.metric("Δ € totale (A − B)",
                                        f"{_delta_eur_tot:+,.0f} €")
                            _ve2.metric("Δ € positivi (aumento)",
                                        f"{_delta_eur_pos:+,.0f} €")
                            _ve3.metric("Δ € negativi (riduzione)",
                                        f"{_delta_eur_neg:+,.0f} €")

                        # Grafico top delta SS
                        _top_delta = _vdf.reindex(
                            _vdf["Δ SS assoluto"].abs().nlargest(20).index
                        ).copy()
                        _top_delta["Label"] = (
                            _top_delta["SKU"].astype(str) + " "
                            + _top_delta["Description"].astype(str).str[:25]
                        )
                        fig_vdelta = go.Figure(go.Bar(
                            x=_top_delta["Δ SS assoluto"],
                            y=_top_delta["Label"],
                            orientation="h",
                            marker_color=[
                                "#E3000F" if v > 0 else "#1a6fa8"
                                for v in _top_delta["Δ SS assoluto"]
                            ],
                            text=_top_delta["Δ SS assoluto"].round(1),
                            textposition="outside",
                            hovertemplate="<b>%{y}</b><br>Δ SS = %{x:,.1f}<extra></extra>",
                        ))
                        fig_vdelta.update_layout(
                            title="Top 20 SKU per |Δ Safety Stock| (A − B)",
                            height=520,
                            plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                            font_color="#eee",
                            xaxis_title="Δ Safety Stock",
                            margin=dict(l=10, r=80, t=40, b=10),
                        )
                        st.plotly_chart(fig_vdelta, use_container_width=True)

                        # Grafico top delta economico
                        if "Δ € assoluto" in _vdf.columns:
                            _top_eur = _vdf.reindex(
                                _vdf["Δ € assoluto"].abs().nlargest(20).index
                            ).copy()
                            _top_eur["Label"] = (
                                _top_eur["SKU"].astype(str) + " "
                                + _top_eur["Description"].astype(str).str[:25]
                            )
                            fig_veur = go.Figure(go.Bar(
                                x=_top_eur["Δ € assoluto"],
                                y=_top_eur["Label"],
                                orientation="h",
                                marker_color=[
                                    "#e67e00" if v > 0 else "#2ecc71"
                                    for v in _top_eur["Δ € assoluto"]
                                ],
                                text=_top_eur["Δ € assoluto"].apply(lambda v: f"{v:+,.0f} €"),
                                textposition="outside",
                                hovertemplate="<b>%{y}</b><br>Δ € = %{x:,.0f} €<extra></extra>",
                            ))
                            fig_veur.update_layout(
                                title="Top 20 SKU per |Δ Valore Economico Safety| (A − B)",
                                height=520,
                                plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                                font_color="#eee",
                                xaxis_title="Δ € Safety Stock",
                                margin=dict(l=10, r=100, t=40, b=10),
                            )
                            st.plotly_chart(fig_veur, use_container_width=True)

                        # Tabella completa
                        with st.expander("Tabella completa SKU per SKU"):
                            st.dataframe(
                                safe_df(_vdf.reset_index(drop=True)),
                                use_container_width=True, height=420, hide_index=True,
                            )

                        # Download
                        _vbuf = io.BytesIO()
                        with pd.ExcelWriter(_vbuf, engine="openpyxl") as _vwr:
                            _vdf.to_excel(_vwr, index=False, sheet_name="Confronto Versioni")
                        _vbuf.seek(0)
                        st.download_button(
                            "Scarica confronto (Excel)",
                            data=_vbuf,
                            file_name=f"confronto_{_lbl_a[:20]}_vs_{_lbl_b[:20]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_ver_compare",
                        )

        # ========================================================================
        # SUB-TAB B — RISULTATI SAFETY STOCK
        # ========================================================================

        with st_results:
            st.subheader("Risultati Safety Stock per SKU")

            df_results = None

            src1, src2 = st.columns(2)
            with src1:
                if "results" in st.session_state and st.session_state["results"] is not None:
                    if st.button("Usa risultati sessione corrente"):
                        df_results = st.session_state["results"]
                        st.success("Risultati caricati dalla sessione corrente.")
            with src2:
                _ss_p = OUTPUT_FILES["Safety Stock"]
                if file_exists(_ss_p):
                    if st.button("Ricarica da file Excel"):
                        load_safety_stock.clear()
                        df_results = load_safety_stock(str(_ss_p))
                        st.success(f"Risultati caricati da `{_ss_p.name}`.")

            # Auto-load
            if df_results is None and "results" in st.session_state:
                df_results = st.session_state["results"]
            if df_results is None and file_exists(OUTPUT_FILES["Safety Stock"]):
                try:
                    df_results = load_safety_stock(str(OUTPUT_FILES["Safety Stock"]))
                except Exception:
                    pass

            if df_results is not None and len(df_results) > 0:
                _pct_used   = st.session_state.get("percentile_used", percentile)
                col_ss_total = next(
                    (c for c in df_results.columns
                     if f"p{_pct_used}" in c.lower() and "total" in c.lower()),
                    next((c for c in df_results.columns
                          if "safety_stock_p" in c.lower() and "total" in c.lower()), None),
                )

                def _prepare_results(df: pd.DataFrame):
                    _hide   = ["p50", "dirich", "alpha"]
                    vis_cols = [c for c in df.columns
                                if not any(p in c.lower() for p in _hide)]
                    df_vis   = df[vis_cols].copy()
                    df_vis["_search"] = df_vis.astype(str).apply(
                        lambda row: " ".join(row.values), axis=1
                    ).str.lower()
                    return df_vis, vis_cols

                df_vis, vis_cols = _prepare_results(df_results)

                with st.expander("Filtri e Ordinamento", expanded=True):
                    f1, f2 = st.columns(2)
                    with f1:
                        search_sku = st.text_input("Cerca SKU / Codice",
                                                   placeholder="es: Z_COLOR, RED, ...")
                    with f2:
                        sort_col_opts = [c for c in vis_cols
                                         if df_results[c].dtype in [np.float64, np.int64,
                                                                     float, int]]
                        sort_by = st.selectbox(
                            "Ordina per", sort_col_opts,
                            index=(sort_col_opts.index(col_ss_total)
                                   if col_ss_total and col_ss_total in sort_col_opts else 0),
                        )
                    sort_asc = st.checkbox("Ordine crescente", value=False)

                df_filtered = df_vis.copy()
                if search_sku:
                    df_filtered = df_filtered[df_filtered["_search"].str.contains(
                        search_sku.strip().lower(), regex=False, na=False,
                    )]
                if sort_by and sort_by in df_filtered.columns:
                    df_filtered = df_filtered.sort_values(sort_by, ascending=sort_asc)
                df_filtered = df_filtered.drop(columns=["_search"], errors="ignore")

                st.markdown(f"**{len(df_filtered)} componenti** (di {len(df_results)} totali)")
                st.dataframe(safe_df(df_filtered.reset_index(drop=True)),
                             use_container_width=True, height=450)

                st.markdown("---")
                dl1, dl2 = st.columns(2)
                with dl1:
                    _buf = io.BytesIO()
                    with pd.ExcelWriter(_buf, engine="openpyxl") as wr:
                        df_filtered.to_excel(wr, index=False, sheet_name="Safety Stock")
                    _buf.seek(0)
                    st.download_button("Scarica Excel (filtrato)", data=_buf,
                                       file_name="safety_stock_results.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with dl2:
                    st.download_button("Scarica CSV (filtrato)",
                                       data=df_filtered.to_csv(index=False).encode("utf-8"),
                                       file_name="safety_stock_results.csv", mime="text/csv")

                # ── Visualizzazioni rapide ────────────────────────────────────────
                st.markdown("---")
                col_ss_viz = next(
                    (c for c in df_results.columns
                     if "safety_stock_p" in c.lower() and "total" in c.lower()), None,
                )
                pct_label = col_ss_viz.split("_")[2].upper() if col_ss_viz else "P99"
                col_lt_r  = next((c for c in df_results.columns if "lead_time" in c.lower()), None)
                col_sku_r = next((c for c in df_results.columns
                                  if c.lower() in ["sku", "codice", "id"]), None)

                if col_ss_viz:
                    st.markdown(f"#### Top SKU per Safety Stock {pct_label}")
                    n_top = st.slider("Numero SKU da mostrare", 5, 50, 20)
                    df_top = df_results.nlargest(n_top, col_ss_viz)
                    x_vals = df_top[col_sku_r].astype(str) if col_sku_r else df_top.index.astype(str)
                    fig_top = px.bar(
                        df_top, x=x_vals, y=col_ss_viz,
                        color=col_lt_r if col_lt_r else None,
                        color_continuous_scale="Reds",
                        title=f"Top {n_top} SKU per Safety Stock {pct_label}",
                        labels={col_ss_viz: f"Safety Stock {pct_label}", "x": "SKU"},
                    )
                    fig_top.update_layout(
                        plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                        font_color="#eee", xaxis_tickangle=-45,
                    )
                    st.plotly_chart(fig_top, use_container_width=True)

                    st.markdown(f"#### Distribuzione Safety Stock {pct_label}")
                    fig_dist = go.Figure(go.Histogram(
                        x=df_results[col_ss_viz].dropna(),
                        marker_color="#E3000F", opacity=0.85, nbinsx=40,
                    ))
                    fig_dist.update_layout(
                        xaxis_title=f"Safety Stock {pct_label}", yaxis_title="N° SKU",
                        plot_bgcolor="#0d1117", paper_bgcolor="#0d1117", font_color="#eee",
                    )
                    st.plotly_chart(fig_dist, use_container_width=True)

            else:
                st.info("Nessun risultato disponibile. Esegui la pipeline dalla tab **Esegui Analisi**.")

        # ========================================================================
        # SUB-TAB C — SAFETY MENSILE
        # ========================================================================

        with st_monthly:
            st.subheader("Safety Stock Mensile per SKU")

            monthly_path = OUTPUT_FILES["Safety Stock Mensile"]

            if file_exists(monthly_path):
                _btn_reload = st.button("Ricarica da file", key="btn_reload_monthly")
                if _btn_reload:
                    load_safety_stock_monthly.clear()
                    st.rerun()
                try:
                    df_monthly = load_safety_stock_monthly(str(monthly_path))
                    st.markdown(f"**File:** `{monthly_path.name}` · {format_size(monthly_path)}")
                except Exception as e:
                    st.error(f"Errore lettura file mensile: {e}")
                    df_monthly = None
            else:
                st.info(
                    f"File `{monthly_path.name}` non trovato.  \n"
                    "Esegui la pipeline con **Step 3** abilitato."
                )
                df_monthly = None

            if df_monthly is not None and len(df_monthly) > 0:
                ss_month_cols = [c for c in df_monthly.columns
                                 if "safety_stock_p" in c.lower()
                                 and not c.lower().endswith("_total")
                                 and "_" in c]
                month_labels = []
                for c in ss_month_cols:
                    parts = c.split("_")
                    month_labels.append(parts[-1] if len(parts) >= 4 else c)

                col_sku_m    = next((c for c in df_monthly.columns
                                     if c.lower() in ["sku", "codice", "id"]), None)
                col_ss_total_m = next((c for c in df_monthly.columns
                                       if "safety_stock_p" in c.lower() and "total" in c.lower()), None)

                kpm1, kpm2, kpm3, kpm4 = st.columns(4)
                kpm1.metric("SKU presenti",      len(df_monthly))
                kpm2.metric("Mesi disponibili",  len(ss_month_cols))
                if col_ss_total_m:
                    kpm3.metric("Max SS Totale",   f"{df_monthly[col_ss_total_m].max():,.0f}")
                    kpm4.metric("Media SS Totale", f"{df_monthly[col_ss_total_m].mean():,.1f}")

                st.markdown("---")

                if ss_month_cols:
                    st.markdown("#### Dettaglio per SKU")
                    sku_options = df_monthly[col_sku_m].astype(str).tolist() if col_sku_m else []
                    if sku_options:
                        selected_sku_m = st.selectbox("Seleziona SKU", sku_options, key="sel_sku_monthly")
                        df_sku_row = df_monthly[df_monthly[col_sku_m].astype(str) == selected_sku_m]
                        if not df_sku_row.empty:
                            ss_values = df_sku_row[ss_month_cols].iloc[0].values.astype(float)
                            fig_sku = go.Figure(go.Bar(
                                x=month_labels, y=ss_values,
                                marker_color="#E3000F", name="Safety Stock",
                                text=[f"{v:,.0f}" for v in ss_values], textposition="outside",
                            ))
                            if col_ss_total_m:
                                _tot_v = float(df_sku_row[col_ss_total_m].iloc[0])
                                fig_sku.add_hline(
                                    y=_tot_v / len(ss_month_cols) if ss_month_cols else 0,
                                    line_dash="dot", line_color="#ff9800",
                                    annotation_text=f"Media: {_tot_v/len(ss_month_cols):,.1f}",
                                )
                            fig_sku.update_layout(
                                title=f"Safety Stock Mensile — SKU: {selected_sku_m}",
                                xaxis_title="Mese", yaxis_title="Safety Stock",
                                plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                                font_color="#eee", xaxis_tickangle=-35,
                            )
                            st.plotly_chart(fig_sku, use_container_width=True)

                    st.markdown("---")
                    st.markdown("#### Andamento Aggregato nel Tempo")
                    agg_by_month = df_monthly[ss_month_cols].sum()
                    df_agg = pd.DataFrame({
                        "Mese": month_labels,
                        "Safety Stock Totale": agg_by_month.values,
                    })
                    fig_agg = px.line(
                        df_agg, x="Mese", y="Safety Stock Totale",
                        title="Somma Safety Stock per Mese (tutti gli SKU)", markers=True,
                    )
                    fig_agg.update_traces(line_color="#E3000F", marker_color="#ff6b6b", marker_size=8)
                    fig_agg.update_layout(
                        plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                        font_color="#eee", xaxis_tickangle=-35,
                    )
                    st.plotly_chart(fig_agg, use_container_width=True)

                with st.expander("Tabella dati completa"):
                    search_m = st.text_input("Filtra SKU", placeholder="Cerca...", key="search_monthly")
                    df_m_show = df_monthly.copy()
                    if search_m and col_sku_m:
                        df_m_show = df_m_show[
                            df_m_show[col_sku_m].astype(str).str.lower().str.contains(
                                search_m.strip().lower(), regex=False, na=False,
                            )
                        ]
                    st.markdown(f"**{len(df_m_show)} righe**")
                    st.dataframe(safe_df(df_m_show.reset_index(drop=True)),
                                 use_container_width=True, height=400)

                st.markdown("---")
                dl_m1, dl_m2 = st.columns(2)
                with dl_m1:
                    buf_m = io.BytesIO()
                    with pd.ExcelWriter(buf_m, engine="openpyxl") as wr:
                        df_monthly.to_excel(wr, index=False, sheet_name="Safety Stock Mensile")
                    buf_m.seek(0)
                    st.download_button("Scarica Excel Mensile", data=buf_m,
                                       file_name="safety_stock_mensile.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with dl_m2:
                    st.download_button("Scarica CSV Mensile",
                                       data=df_monthly.to_csv(index=False).encode("utf-8"),
                                       file_name="safety_stock_mensile.csv", mime="text/csv")

        # ========================================================================
        # SUB-TAB E — TAKE RATE
        # ========================================================================

        with st_tr:
            st.subheader("Anteprima Take Rate")

            tr_path = INPUT_FILES["Take Rate"]
            if file_exists(tr_path):
                try:
                    st.markdown(f"**File:** `{tr_path.name}` · {format_size(tr_path)}")

                    all_sheets = load_tr_sheets_list(str(tr_path))

                    model_sheets    = [s for s in all_sheets if s.lower().startswith("tr_model_")]
                    if not model_sheets:
                        model_sheets = [s for s in all_sheets if s.lower() == "tr-model"]

                    optional_sheets = [s for s in all_sheets if s.lower().startswith("tr_optional_")]
                    if not optional_sheets:
                        optional_sheets = [s for s in all_sheets if s.lower() == "tr-totv21e"]

                    st.markdown("---")
                    tr_c1, tr_c2 = st.columns(2)

                    with tr_c1:
                        if model_sheets:
                            sel_model_sheet = st.selectbox(
                                "Periodo Model Mix", options=model_sheets,
                                index=len(model_sheets) - 1,
                            )
                        else:
                            sel_model_sheet = None
                            st.warning("Nessun foglio Model Mix trovato.")

                    with tr_c2:
                        if optional_sheets:
                            sel_optional_sheet = st.selectbox(
                                "Periodo Caratteristiche", options=optional_sheets,
                                index=len(optional_sheets) - 1,
                            )
                        else:
                            sel_optional_sheet = None
                            st.warning("Nessun foglio Optional trovato.")

                    # ── Model Mix ───────────────────────────────────────────────
                    st.markdown("---")
                    st.markdown(f"### Model Mix  —  `{sel_model_sheet}`")

                    if sel_model_sheet:
                        df_model = load_tr_sheet(str(tr_path), sel_model_sheet, header_row=0)

                        col_model_name = next(
                            (c for c in df_model.columns
                             if "opzione" in c.lower() or "concat" in c.lower() or "model" in c.lower()),
                            df_model.columns[0],
                        )
                        col_tr_model = next(
                            (c for c in df_model.columns if "take rate" in c.lower()), None,
                        )
                        if col_tr_model is None:
                            num_cols = [c for c in df_model.columns
                                        if pd.api.types.is_numeric_dtype(df_model[c])
                                        and c != col_model_name]
                            if num_cols:
                                col_tr_model = num_cols[0]

                        df_model_valid = df_model.dropna(subset=[col_model_name])

                        mm1, mm2, mm3 = st.columns(3)
                        mm1.metric("Modelli configurati", len(df_model_valid))
                        if col_tr_model:
                            tr_sum = pd.to_numeric(df_model_valid[col_tr_model], errors="coerce").sum()
                            mm2.metric("Somma Take Rate",
                                       f"{tr_sum:.2%}" if tr_sum <= 1 else f"{tr_sum:.1f}%")
                        mm3.metric("Foglio", sel_model_sheet)

                        if col_tr_model:
                            df_mc = df_model_valid[[col_model_name, col_tr_model]].copy()
                            df_mc.columns = ["Modello", "Take Rate"]
                            df_mc["Take Rate"] = pd.to_numeric(df_mc["Take Rate"], errors="coerce")
                            df_mc = df_mc.dropna(subset=["Take Rate"])
                            df_mc["Take Rate %"] = df_mc["Take Rate"] * 100

                            _mm_c1, _mm_c2 = st.columns([2, 3])
                            with _mm_c1:
                                st.markdown("**Dati Model Mix**")
                                st.dataframe(safe_df(df_mc), use_container_width=True, hide_index=True)
                            with _mm_c2:
                                fig_mm = px.pie(
                                    df_mc, names="Modello", values="Take Rate",
                                    title=f"Model Mix — {sel_model_sheet}",
                                    hole=0.35,
                                )
                                fig_mm.update_traces(textinfo="label+percent")
                                fig_mm.update_layout(
                                    paper_bgcolor="#0d1117", font_color="#eee",
                                    height=340, margin=dict(t=40, b=0, l=0, r=0),
                                )
                                st.plotly_chart(fig_mm, use_container_width=True)

                    # ── Optional / Caratteristiche ───────────────────────────────
                    if sel_optional_sheet:
                        st.markdown("---")
                        st.markdown(f"### Caratteristiche TR  —  `{sel_optional_sheet}`")

                        df_opt = load_tr_sheet(str(tr_path), sel_optional_sheet, header_row=1)

                        col_char = next(
                            (c for c in df_opt.columns
                             if "caract" in c.lower() or "char" in c.lower()
                             or "caratt" in c.lower() or "opzione" in c.lower()),
                            df_opt.columns[0] if len(df_opt.columns) > 0 else None,
                        )
                        col_tr_opt = next(
                            (c for c in df_opt.columns if "take rate" in c.lower()), None,
                        )
                        if col_tr_opt is None:
                            _num_opt = [c for c in df_opt.columns
                                        if pd.api.types.is_numeric_dtype(df_opt[c])
                                        and c != col_char]
                            if _num_opt:
                                col_tr_opt = _num_opt[0]

                        if col_char and col_tr_opt:
                            df_opt_valid = df_opt.dropna(subset=[col_char, col_tr_opt]).copy()
                            df_opt_valid["_tr_num"] = pd.to_numeric(
                                df_opt_valid[col_tr_opt], errors="coerce"
                            )

                            ot1, ot2, ot3 = st.columns(3)
                            ot1.metric("Caratteristiche", len(df_opt_valid))
                            _sum_opt = df_opt_valid["_tr_num"].sum()
                            ot2.metric("Somma TR",
                                       f"{_sum_opt:.2%}" if _sum_opt <= 1 else f"{_sum_opt:.1f}%")
                            ot3.metric("Foglio", sel_optional_sheet)

                            # Raggruppa per caratteristica se colonna gruppo presente
                            col_group = next(
                                (c for c in df_opt.columns
                                 if "group" in c.lower() or "gruppo" in c.lower()
                                 or "cat" in c.lower()),
                                None,
                            )
                            if col_group:
                                _grp_list = sorted(df_opt_valid[col_group].dropna().unique().tolist())
                                sel_grp = st.selectbox("Filtra per gruppo/categoria",
                                                       ["Tutti"] + _grp_list, key="tr_grp")
                                if sel_grp != "Tutti":
                                    df_opt_valid = df_opt_valid[df_opt_valid[col_group] == sel_grp]

                            _opt_c1, _opt_c2 = st.columns([2, 3])
                            with _opt_c1:
                                st.dataframe(
                                    safe_df(df_opt_valid[[col_char, col_tr_opt]].rename(columns={
                                        col_char: "Caratteristica", col_tr_opt: "Take Rate",
                                    }).reset_index(drop=True)),
                                    use_container_width=True, hide_index=True,
                                )
                            with _opt_c2:
                                fig_opt = px.bar(
                                    df_opt_valid, x=col_char, y="_tr_num",
                                    title=f"Take Rate Caratteristiche — {sel_optional_sheet}",
                                    color="_tr_num",
                                    color_continuous_scale="Reds",
                                    labels={col_char: "Caratteristica", "_tr_num": "Take Rate"},
                                )
                                fig_opt.update_layout(
                                    plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                                    font_color="#eee", xaxis_tickangle=-45, height=360,
                                )
                                st.plotly_chart(fig_opt, use_container_width=True)
                        else:
                            st.dataframe(safe_df(df_opt), use_container_width=True)

                    with st.expander("Tutti i fogli disponibili"):
                        st.write(all_sheets)

                except Exception as e:
                    st.error(f"Errore nella lettura del file TR: {e}")
            else:
                st.warning(f"File non trovato: `{tr_path}`. Caricalo nella cartella `Input/`.")

# ============================================================================
# MODULO — CYCLE STOCK
# ============================================================================

elif _active_module == "cs":

    tab_run, tab_view = st.tabs(["Esegui Analisi", "Visualizza Risultati"])

    # ########################################################################
    # CS — TAB 1: ESEGUI ANALISI
    # ########################################################################

    with tab_run:

        st.markdown("## Cycle Stock — Actual vs Forecast TR")
        st.caption(
            "Esegue due simulazioni MC con file TR distinti e confronta la domanda media per SKU. "
            "Con `Prezzi.xlsx` calcola anche il delta economico (€)."
        )

        # File TR disponibili
        available_xlsx = sorted([
            f.name for f in INPUT_DIR.glob("*.xlsx")
            if f.name.lower() != "prezzi.xlsx"
        ]) if INPUT_DIR.exists() else []

        cs_c1, cs_c2 = st.columns(2)

        with cs_c1:
            st.markdown("**TR Actual** — situazione corrente")
            tr_actual_choice = st.selectbox(
                "File TR Actual",
                options=available_xlsx,
                index=next((i for i, f in enumerate(available_xlsx)
                            if "actual" in f.lower()), 0) if available_xlsx else 0,
                key="cmp_tr_actual",
            )
            tr_actual_path = INPUT_DIR / tr_actual_choice if tr_actual_choice else None
            if tr_actual_path and tr_actual_path.exists():
                st.caption(f"✅ {format_size(tr_actual_path)}")
            elif tr_actual_choice:
                st.caption("❌ File non trovato")

        with cs_c2:
            st.markdown("**TR Forecast** — scenario futuro")
            tr_forecast_choice = st.selectbox(
                "File TR Forecast",
                options=available_xlsx,
                index=next((i for i, f in enumerate(available_xlsx)
                            if "forecast" in f.lower()),
                           min(1, len(available_xlsx) - 1)) if len(available_xlsx) > 1 else 0,
                key="cmp_tr_forecast",
            )
            tr_forecast_path = INPUT_DIR / tr_forecast_choice if tr_forecast_choice else None
            if tr_forecast_path and tr_forecast_path.exists():
                st.caption(f"✅ {format_size(tr_forecast_path)}")
            elif tr_forecast_choice:
                st.caption("❌ File non trovato")

        # Prezzi opzionale
        prezzi_path = INPUT_FILES["Prezzi"]
        if prezzi_path.exists():
            st.success(f"`Prezzi.xlsx` trovato ({format_size(prezzi_path)}) — calcolo Δ € abilitato")
        else:
            st.info("`Prezzi.xlsx` non trovato — confronto solo su volumi (no delta €)")

        if (tr_actual_choice and tr_forecast_choice
                and tr_actual_choice == tr_forecast_choice):
            st.warning("⚠️ I due file TR selezionati sono identici — il confronto non avrà senso.")

        # Prerequisiti Cycle Stock
        CMP_REQUIRED = ["Esclusioni", "Domanda", "Lead Times", "Unique Rows"]
        _cmp_missing = [n for n in CMP_REQUIRED if not file_exists(INPUT_FILES[n])]
        if not (tr_actual_path and tr_actual_path.exists()):
            _cmp_missing.append("TR Actual")
        if not (tr_forecast_path and tr_forecast_path.exists()):
            _cmp_missing.append("TR Forecast")

        if _cmp_missing:
            st.error(f"⚠️ File mancanti: {', '.join(_cmp_missing)}")
        else:
            st.success("✅ Tutti i file necessari per Cycle Stock sono presenti.")

        cmp_btn = st.button(
            "AVVIA CYCLE STOCK",
            type="primary",
            disabled=bool(_cmp_missing),
            key="btn_cmp",
        )

        if cmp_btn:
            if st.session_state.get("_cs_running", False):
                st.warning("Cycle Stock già in esecuzione — attendi il completamento.")
                st.stop()
            st.session_state["_cs_running"] = True

            cmp_log_container = st.empty()
            cmp_progress      = st.progress(0)
            cmp_status        = st.empty()
            cmp_logs          = []

            def cmp_log(msg: str):
                cmp_logs.append(msg)
                cmp_log_container.markdown(
                    f'<div class="log-console">{"<br>".join(cmp_logs[-30:])}</div>',
                    unsafe_allow_html=True,
                )

            cmp_start = time.time()
            try:
                if str(BASE_DIR) not in sys.path:
                    sys.path.insert(0, str(BASE_DIR))

                from main_demand_comparison_V2 import main as run_comparison  # [V2]

                cmp_log("=" * 60)
                cmp_log("CYCLE STOCK — ACTUAL vs FORECAST")
                cmp_log(f"Avvio: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                cmp_log(f"TR Actual:   {tr_actual_choice}")
                cmp_log(f"TR Forecast: {tr_forecast_choice}")
                cmp_log(f"Run MC: {n_runs}  |  Offset: {start_month_offset}")
                cmp_log("=" * 60)

                cmp_status.text("Simulazione in corso (può richiedere diversi minuti)...")
                cmp_progress.progress(10)

                df_cmp = run_comparison(
                    tr_actual_path=str(tr_actual_path),
                    tr_forecast_path=str(tr_forecast_path),
                    n_runs=n_runs,
                    start_month_offset=start_month_offset,
                )

                cmp_progress.progress(95)
                cmp_elapsed = time.time() - cmp_start
                cmp_log("\n" + "=" * 60)
                cmp_log(f"COMPLETATO in {cmp_elapsed:.1f}s ({cmp_elapsed/60:.1f} min)")
                if df_cmp is not None and "delta_demand" in df_cmp.columns:
                    n_pos = int((df_cmp["delta_demand"] > 0).sum())
                    n_neg = int((df_cmp["delta_demand"] < 0).sum())
                    cmp_log(f"SKU delta > 0: {n_pos}  |  delta < 0: {n_neg}")
                cmp_log("=" * 60)
                cmp_progress.progress(100)
                cmp_status.text("Completato!")

                st.session_state["df_comparison"]      = df_cmp
                st.session_state["cmp_actual_label"]   = tr_actual_choice
                st.session_state["cmp_forecast_label"] = tr_forecast_choice
                load_comparison.clear()

                st.success(
                    f"✅ Cycle Stock completato in {cmp_elapsed:.1f}s — "
                    "vai a **Visualizza Risultati** per l'analisi."
                )
                st.session_state["last_ram"] = get_memory_mb()

            except Exception as e:
                import traceback
                cmp_progress.progress(100)
                cmp_status.text("Errore.")
                cmp_log(f"\n✗ ERRORE: {e}")
                st.error(f"Errore durante il confronto: {e}")
                st.code(traceback.format_exc(), language="python")

            finally:
                st.session_state["_cs_running"] = False


    # ########################################################################
    # CS — TAB 2: VISUALIZZA RISULTATI
    # ########################################################################

    with tab_view:

        st.subheader("Risultati — Cycle Stock")

        _df_cmp = st.session_state.get("df_comparison")

        if _df_cmp is None:
            _cmp_path = OUTPUT_FILES["Demand Comparison"]
            if _cmp_path.exists():
                _df_cmp = pd.read_excel(_cmp_path)

        if _df_cmp is None:
            st.info(
                "Nessun risultato Cycle Stock disponibile. "
                "Esegui prima l'analisi nella scheda **Esegui Analisi**."
            )
        else:
            _actual_lbl   = st.session_state.get("cmp_actual_label",   "Actual")
            _forecast_lbl = st.session_state.get("cmp_forecast_label", "Forecast")

            st.caption(
                f"Confronto domanda media per SKU: **{_actual_lbl}** vs "
                f"**{_forecast_lbl}**. Δ = Forecast − Actual."
            )

            # ── Filtro testo ─────────────────────────────────────────────
            _search_cs = st.text_input(
                "Filtra SKU / configurazione", key="cs_view_search"
            )
            _df_show = _df_cmp.copy()
            if _search_cs:
                _mask = _df_show.apply(
                    lambda col: col.astype(str).str.contains(
                        _search_cs, case=False, na=False
                    )
                ).any(axis=1)
                _df_show = _df_show[_mask]

            st.dataframe(_df_show, use_container_width=True, hide_index=True)

            # ── Metriche rapide ──────────────────────────────────────────
            if "delta_demand" in _df_show.columns:
                _m1, _m2, _m3 = st.columns(3)
                _m1.metric(
                    "SKU con Δ > 0 (↑ Forecast)",
                    int((_df_show["delta_demand"] > 0).sum()),
                )
                _m2.metric(
                    "SKU con Δ < 0 (↓ Forecast)",
                    int((_df_show["delta_demand"] < 0).sum()),
                )
                _m3.metric(
                    "Δ Medio",
                    f"{_df_show['delta_demand'].mean():.2f}",
                )

                # ── Grafico top-30 per |Δ| ───────────────────────────────
                _sku_col = next(
                    (
                        c for c in _df_show.columns
                        if c.lower() in ("sku", "codice", "sku_code", "config_code")
                    ),
                    _df_show.columns[0],
                )
                _top30 = (
                    _df_show[[_sku_col, "delta_demand"]]
                    .assign(_abs=lambda d: d["delta_demand"].abs())
                    .sort_values("_abs", ascending=False)
                    .drop(columns="_abs")
                    .head(30)
                )
                _fig_cs = px.bar(
                    _top30,
                    x=_sku_col,
                    y="delta_demand",
                    color="delta_demand",
                    color_continuous_scale="RdYlGn",
                    title="Top-30 SKU per |Δ Domanda| (Forecast − Actual)",
                    labels={"delta_demand": "Δ Domanda", _sku_col: "SKU"},
                )
                _fig_cs.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(_fig_cs, use_container_width=True)

            # ── Delta economico (se disponibile) ─────────────────────────
            _euro_col = next(
                (
                    c for c in _df_show.columns
                    if "€" in c or "euro" in c.lower() or "delta_value" in c.lower()
                ),
                None,
            )
            if _euro_col:
                _tot_eur = _df_show[_euro_col].sum()
                st.metric("Δ Economico Totale", f"€ {_tot_eur:,.0f}")

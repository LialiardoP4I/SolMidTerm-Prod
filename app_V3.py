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
import re
import sys
import time
import gc
import psutil
import io
import shutil
import copy
import importlib
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

def _find_model_file(pattern: str) -> Path:
    """Cerca il primo file che matcha `pattern` in qualsiasi sottocartella MODEL/."""
    _model_root = INPUT_DIR / "MODEL"
    if _model_root.exists():
        for _sub in sorted(_model_root.iterdir()):
            if _sub.is_dir():
                for _f in _sub.iterdir():
                    if _f.name.lower().startswith(pattern.lower()):
                        return _f
    # fallback path (non esistente, ma mostra path leggibile nel UI)
    return _model_root / "MODEL" / pattern

_bom150_path   = _find_model_file("BOM_150")
_mappatura_path = _find_model_file("Mappatura_")

INPUT_FILES = {
    "BOM150":      _bom150_path,
    "Mappatura":   _mappatura_path,
    "Esclusioni":  INPUT_DIR / "esclusioni.xlsx",
    "Domanda":     INPUT_DIR / "Total_demand.xlsx",
    "Take Rate":   INPUT_DIR / "TR TOTALV21E - Copia.xlsx",    # [V2]
    "Matrix":      INPUT_DIR / "matrix_output_V2.xlsx",        # SUB-TAB D — TOTAL DEMAND
    "Unique Rows": INPUT_DIR / "tutte_righe_univoche_V2.xlsx", # [V2]
    "Prezzi":      INPUT_DIR / "Dati Logistica" / "PREZZI.XLSX",   # opzionale
}

OUTPUT_FILES = {
    "Safety Stock":         OUTPUT_DIR / "sku_safety_stock_leadtime_PER_SKU_V2.xlsx",
    "Safety Stock Mensile": OUTPUT_DIR / "sku_safety_stock_PER_MESE_V2.xlsx",
    "Demand Comparison":    OUTPUT_DIR / "demand_comparison_actual_vs_forecast_V2.xlsx",
    "Stock Alignment":      OUTPUT_DIR / "stock_alignment_V2.xlsx",
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

# ── Confidenza TR ────────────────────────────────────────────────────────────
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

def rescale_alphas(model_mix, characteristics, monthly_tr,
                   aff_model: str, aff_optional: str):
    """
    Wrapper: delega a mc_simulator_V2.rescale_alphas passando _AFF_MAP corrente.
    Fonte di verita' unica in mc_simulator_V2 (usato anche da mc_simulator_rolling).
    """
    try:
        from mc_simulator_V2 import rescale_alphas as _rescale_impl
        return _rescale_impl(
            model_mix, characteristics, monthly_tr,
            aff_model=aff_model, aff_optional=aff_optional,
            aff_map=_AFF_MAP,
        )
    except ImportError:
        # Fallback inline se mc_simulator_V2 non disponibile
        import copy
        scale_m = _AFF_MAP[aff_model]
        scale_o = _AFF_MAP[aff_optional]
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

def archive_run_sa(actual_label: str, forecast_label: str,
                   n_months: int = 0, n_runs: int = 0) -> "Path | None":
    """
    Copia stock_alignment_V2.xlsx in Output/history_sa/ con nome
    GG-MM-AAAA_<actual>-vs-<forecast>_Xmesi_NNNrun.xlsx
    Ritorna il Path del file archiviato, o None se la sorgente non esiste.
    """
    src = OUTPUT_FILES["Stock Alignment"]
    if not src.exists():
        return None
    history_dir = OUTPUT_DIR / "history_sa"
    history_dir.mkdir(exist_ok=True)
    date_str = time.strftime("%d-%m-%Y")
    # Sanitizza etichette (rimuove slash, spazi e caratteri problematici)
    def _san(s):
        return re.sub(r'[\\/:*?"<>|\s]+', "_", str(s))[:20]
    fname = f"{date_str}_{_san(actual_label)}-vs-{_san(forecast_label)}_{n_months}mesi_{n_runs}run.xlsx"
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

# Default confidenza TR
_aff_def_model    = _AFF_LEVELS[-1]
_aff_def_optional = _AFF_LEVELS[len(_AFF_LEVELS) // 2]


# ============================================================================
# NAVIGAZIONE PRINCIPALE
# ============================================================================

if "active_module" not in st.session_state:
    st.session_state["active_module"] = "ss"

_active_module = st.session_state["active_module"]

_col_ss, _col_cs, _col_sa = st.columns(3)
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
with _col_sa:
    if st.button(
        "Stock Alignment",
        use_container_width=True,
        type="primary" if _active_module == "sa" else "secondary",
        key="nav_sa",
    ):
        st.session_state["active_module"] = "sa"
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

    tab_run, tab_view, tab_rolling = st.tabs(["Esegui Analisi", "Visualizza Risultati", "🔄 Rolling"])

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

        # BOM150 e Mappatura: necessari SOLO per Step 1 — non bloccano Step 2/3
        SS_REQUIRED  = ["Esclusioni", "Domanda", "Take Rate",
                        "Matrix", "Unique Rows"]
        CMP_REQUIRED = ["Esclusioni", "Domanda", "Unique Rows"]

        _logistica_dir_ss = INPUT_DIR / "Dati Logistica"
        _gates_path_ss    = _logistica_dir_ss / "Gates MTSV4.xlsx"
        _dash_files_ss    = list(_logistica_dir_ss.glob("*Dashboard Supplier*.xlsx")) if _logistica_dir_ss.exists() else []
        _dash_path_ss     = _dash_files_ss[-1] if _dash_files_ss else None

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

        def _logistica_status_badge(label: str, path, found: bool):
            badge_cls = "status-ok" if found else "status-err"
            icon      = "✅" if found else "❌"
            size_str  = format_size(path) if (found and path) else "mancante"
            st.markdown(
                f'{icon} **{label}**<br>'
                f'<span class="{badge_cls}">{size_str}</span>',
                unsafe_allow_html=True,
            )

        with st.expander("Safety Stock — file richiesti", expanded=True):
            _file_status_row(SS_REQUIRED, "ss")
            # Lead Times — calcolati on-demand da file Dati Logistica
            _lt_cols = st.columns(2)
            with _lt_cols[0]:
                _logistica_status_badge("Gates MTSV4", _gates_path_ss, _gates_path_ss.exists())
            with _lt_cols[1]:
                _logistica_status_badge(
                    "Dashboard Supplier",
                    _dash_path_ss,
                    _dash_path_ss is not None,
                )
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
                if run_step0:
                    # Spunta attiva → rigenera SEMPRE, sovrascrive file esistenti
                    status_text.text("Step 1: Elaborazione Multi-BOM...")
                    progress_bar.progress(5)
                    log("\n[STEP 1] Rigenera Multi-BOM (V21E + V22E + V24T) — sovrascrittura file esistenti...")
                    t0 = time.time()
                    try:
                        import importlib
                        import process_all_bom_combinations as _pabc
                        importlib.reload(_pabc)          # forza reload se già importato
                        _pabc.main(filter_untraced=filter_untraced)
                        log(f"  ✓ Step 1 completato in {time.time()-t0:.1f}s")
                        progress_bar.progress(25)
                    except Exception as e:
                        log(f"  ✗ ERRORE Step 1: {e}")
                        st.error(f"Errore Step 1: {e}")
                    gc.collect()
                else:
                    # Spunta inattiva → usa file consolidato esistente (se presente)
                    _consolidated_catalog = INPUT_DIR / "tutte_righe_univoche_V2.xlsx"
                    if _consolidated_catalog.exists():
                        log("[STEP 1] Saltato — usando file consolidato esistente (V21E + V22E + V24T)")
                    else:
                        log("[STEP 1] Saltato — ATTENZIONE: file consolidato non trovato, abilita Step 1 per generarlo")
                        st.warning("⚠️ File tutte_righe_univoche_V2.xlsx non trovato. Abilita 'Step 1 · Rigenera Multi-BOM' e riesegui.")
                    progress_bar.progress(20)

                df_simulation = None

                # ── Step 2 ──────────────────────────────────────────────────────
                if run_step2:
                    status_text.text("Step 2: Simulazione Monte Carlo...")
                    progress_bar.progress(30)
                    log("\n[STEP 2] Simulazione Monte Carlo...")
                    log(f"  Config: {n_runs} run, offset={start_month_offset}")
                    log(f"  Confidenza TR -> Model Mix: {aff_model} (VSS={_AFF_MAP.get(aff_model,0):.0f})  |  Optional: {aff_optional} (VSS={_AFF_MAP.get(aff_optional,0):.0f})")
                    t2 = time.time()
                    try:
                        import importlib
                        import mc_simulator_V2 as _mc_mod
                        importlib.reload(_mc_mod)
                        SimulationConfig                  = _mc_mod.SimulationConfig
                        load_monthly_forecast             = _mc_mod.load_monthly_forecast
                        calculate_n_months_needed         = _mc_mod.calculate_n_months_needed
                        run_monthly_simulations_optimized = _mc_mod.run_monthly_simulations_optimized
                        build_wide_dataframe              = _mc_mod.build_wide_dataframe

                        import tr_parser_V2 as _tr_parser_mod
                        importlib.reload(_tr_parser_mod)
                        parse_tr_file  = _tr_parser_mod.parse_tr_file
                        _load_mtr      = _tr_parser_mod.load_monthly_tr
                        # Aggiorna _AFF_MAP e _AFF_LEVELS usati da rescale_alphas
                        # (uso globals() per evitare SyntaxError da 'global' in scope annidato)
                        _new_aff_map = {k: float(v) for k, v in _tr_parser_mod._AFFIDABILITA_MAP.items()}
                        globals()['_AFF_MAP']    = _new_aff_map
                        globals()['_AFF_LEVELS'] = sorted(_new_aff_map, key=_new_aff_map.__getitem__)
                        log(f"  → affidabilita_config.json ricaricato: {_new_aff_map}")

                        _parse_excl = _mc_mod.parse_exclusions

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
                        _sc_m = _AFF_MAP[aff_model]
                        _sc_o = _AFF_MAP[aff_optional]
                        log(f"  → Alpha riscalati (Model VSS={_sc_m:.0f}, Optional VSS={_sc_o:.0f})")
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
                            calculate_n_months_needed(),
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
                            _results_by_month, _month_names, n_runs=_cfg.n_runs, save_to_file=False
                        )

                        # Step 7b: deriva colori pack da FAIRING COLOR
                        try:
                            build_zcol_pack_color_mapping = _mc_mod.build_zcol_pack_color_mapping
                            apply_pack_colors_to_wide_df  = _mc_mod.apply_pack_colors_to_wide_df
                            # V21E standard + V22E Rally (sostituisce il vecchio Mappatura Rally.xlsx)
                            _pack_color_mapping = build_zcol_pack_color_mapping(
                                str(INPUT_DIR / "MODEL" / "V21E" / "Mappatura_V21E.xlsx"),
                                rally_mappatura_path=str(INPUT_DIR / "MODEL" / "V22E" / "Mappatura_V22E.xlsx"),
                            )
                            # V24T: aggiunge eventuali colori pack specifici
                            _v24t_mp = INPUT_DIR / "MODEL" / "V24T" / "Mappatura_V24T.xlsx"
                            if _v24t_mp.exists():
                                _vm_v24t = build_zcol_pack_color_mapping(str(_v24t_mp))
                                for _k, _v in _vm_v24t.items():
                                    _pack_color_mapping.setdefault(_k, {}).update(_v)
                            if _pack_color_mapping:
                                df_simulation = apply_pack_colors_to_wide_df(
                                    df_simulation, _pack_color_mapping
                                )
                                log("  -> Colori pack derivati da FAIRING COLOR (Touring/Top Case/Rally)")
                            else:
                                log("  [WARN] Mapping ZCOL->pack color vuoto, colori pack non derivati")
                        except Exception as _e:
                            log(f"  [WARN] Derivazione colori pack fallita: {_e}")

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
                        import sku_matching_SS as _sku_mod
                        importlib.reload(_sku_mod)
                        main_sku_v2 = _sku_mod.main_sku_v2
                        if df_simulation is None:
                            log("  ✗ Nessun risultato simulazione disponibile")
                            st.error("Step 3 richiede Step 2. Abilita Step 2 dalla sidebar.")
                        else:
                            results = main_sku_v2(
                                df_simulation,
                                percentile=percentile,
                                sku_catalog_path=str(INPUT_DIR / "tutte_righe_univoche_V2.xlsx"),
                                prezzi_path=str(INPUT_FILES["Prezzi"]),
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
        # Migrazione nomi vecchi tab -> nuovi
        if st.session_state["_ss_sub"] in ("Parametri", "Affidabilità"):
            st.session_state["_ss_sub"] = "Confidenza Take Rate"
        _ss_sub = st.session_state["_ss_sub"]

        _sub_labels = ["File di Input", "Qualità Dati", "Confidenza Take Rate", "Avvia Pipeline"]
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
                            if not file_exists(INPUT_FILES[n]) and n not in ("Matrix", "Unique Rows", "Lead Times")]
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
                    file_exists(INPUT_FILES["Mappatura"])
                    and file_exists(INPUT_FILES["Take Rate"])
                    and file_exists(INPUT_FILES["BOM150"])
                )
                if not _qc_files_ok:
                    st.error("File mancanti: assicurati che Mappatura, Take Rate e BOM150 siano presenti.")
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

                            _riss = []
                            _piss = []
                            st.info("Verifica Lead Times: calcolati on-demand durante STEP 2 (SKU Matching)")

                            # QC multi-versione: V21E + V22E + V24T
                            _DEP_COL_QC = 'Blocco dati per archivio dipendenze'
                            _tr_path_qc = str(INPUT_FILES["Take Rate"])
                            _trace_qc = {
                                'ct_missing': [], 'cv_missing': [], 'ct_system': [],
                                'unknown_ct_set': set(), 'unknown_cv_set': set(),
                            }
                            _QC_VERSIONS = [
                                ('V21E', INPUT_DIR / "MODEL" / "V21E" / "Mappatura_V21E.xlsx",
                                         INPUT_DIR / "MODEL" / "V21E" / "BOM_150V21E.xlsx"),
                                ('V22E', INPUT_DIR / "MODEL" / "V22E" / "Mappatura_V22E.xlsx",
                                         INPUT_DIR / "MODEL" / "V22E" / "BOM_150V22E.xlsx"),
                                ('V24T', INPUT_DIR / "MODEL" / "V24T" / "Mappatura_V24T.xlsx",
                                         INPUT_DIR / "MODEL" / "V24T" / "BOM_150V24T.xlsx"),
                                ('V24T', INPUT_DIR / "PANIV4" / "Mappatura_PANIV4.xlsx",
                                         INPUT_DIR / "PANIV4" / "BOM_150PANIV4.xlsx"),        
                            ]
                            for _ver_qc, _map_path_qc, _bom_path_qc in _QC_VERSIONS:
                                if not _map_path_qc.exists() or not _bom_path_qc.exists():
                                    continue
                                _map_df_qc  = _gd_lmap(str(_map_path_qc))
                                _acv_qc     = _gd_acv(_tr_path_qc, _map_df_qc)
                                _bom_qc     = pd.read_excel(str(_bom_path_qc), header=0)
                                _bom_qc     = _bom_qc[_bom_qc[_DEP_COL_QC].notnull()]
                                _parser_qc  = _GD_CP(_map_df_qc, _acv_qc)
                                _trace_ver  = _gd_tr(_bom_qc, _DEP_COL_QC, _parser_qc)
                                for _row in _trace_ver.get('ct_missing', []):
                                    _trace_qc['ct_missing'].append({'Versione': _ver_qc, **_row})
                                for _row in _trace_ver.get('cv_missing', []):
                                    _trace_qc['cv_missing'].append({'Versione': _ver_qc, **_row})
                                for _row in _trace_ver.get('ct_system', []):
                                    _trace_qc['ct_system'].append({'Versione': _ver_qc, **_row})
                                _trace_qc['unknown_ct_set'] |= _trace_ver.get('unknown_ct_set', set())
                                _trace_qc['unknown_cv_set'] |= _trace_ver.get('unknown_cv_set', set())
                            _gd_wr(_riss, _trace_qc, _piss,
                                   str(OUTPUT_DIR / "quality_check_report.xlsx"))
                            st.success("Quality check completato — scarica il report o apri l'anteprima.")
                            st.rerun()
                        except Exception as _e:
                            st.error(f"Errore durante quality check: {_e}")


            # ════════════════════════════════════════════════════════════════════════
            # SUB-TAB: AFFIDABILITÀ
            # ════════════════════════════════════════════════════════════════════════

        elif _ss_sub == "Confidenza Take Rate":

            st.markdown("## Confidenza Take Rate")

            # ── Spiegazione plain-language ────────────────────────────────────────────────────
            st.info(
                "**Cos'e la Confidenza TR?**  \n"
                "Ogni livello corrisponde a un **Virtual Sample Size (VSS)**: il numero equivalente "
                "di ordini storici che sostengono la stima del Take Rate.  \n"
                "- **VSS basso** (pochi ordini virtuali) → la simulazione esplora scenari molto diversi dal TR stimato  \n"
                "- **VSS alto** (molti ordini virtuali) → la simulazione resta vicina alla stima  \n\n"
                "I valori sono calibrati per volumi Ducati (2.000-5.000 unita/anno per variante), "
                "seguendo la regola del 10%: il VSS non deve superare il 10% degli ordini attesi "
                "(Morita, Thall, Muller 2008)."
            )

            # ── Legenda livelli ────────────────────────────────────────────────────────────────
            _sd_aff_labels = {
                "NULLA":  ("🔴", 50,
                           "Stima senza precedenti",
                           "TR inventato o per opzione mai vista su nessun modello. "
                           "Usa per stress test o opzioni completamente nuove."),
                "BASSA":  ("🟠", 200,
                           "Expert judgment sales/marketing",
                           "TR stimato a priori dal team vendite e product marketing, "
                           "senza dati di vendita reali. Consigliato per la fase pre-lancio."),
                "MEDIA":  ("🟡", 1000,
                           "Analogia con modello precedente",
                           "TR basato su dati di un modello analogo (es. Panigale V4 -> MSV4). "
                           "Consigliato dopo i primi mesi di vendita."),
                "ALTA":   ("🟢", 5000,
                           "Storico consolidato o mix modelli",
                           "TR supportato da dati di vendita reali o decisione commerciale definita. "
                           "Consigliato per il model mix e dopo 1+ anno di vendite."),
            }
            _ci_m = _sd_aff_labels.get(aff_model,    ("", 0, aff_model, ""))
            _ci_o = _sd_aff_labels.get(aff_optional, ("", 0, aff_optional, ""))

            st.markdown("### Configurazione Attuale")
            _cfg_c1, _cfg_c2 = st.columns(2)
            with _cfg_c1:
                st.markdown(
                    f"**Model Mix (L0):** {_ci_m[0]} **{aff_model}** (VSS={_ci_m[1]})  \n"
                    f"*{_ci_m[2]}*  \n{_ci_m[3]}"
                )
            with _cfg_c2:
                st.markdown(
                    f"**Optional (L1):** {_ci_o[0]} **{aff_optional}** (VSS={_ci_o[1]})  \n"
                    f"*{_ci_o[2]}*  \n{_ci_o[3]}"
                )

            # ── Legenda completa con tabella ──────────────────────────────────────────────────
            with st.expander("📖 Legenda completa dei livelli", expanded=True):
                st.markdown(
                    "| Livello | VSS | Significato | Fonte del TR | σ (TR=30%) | Quando usarlo |\n"
                    "|---------|-----|-------------|-------------|------------|---------------|\n"
                    "| 🔴 **NULLA** | 50 | 50 ordini virtuali | Nessun precedente | ±6.5% | Stress test, opzioni mai viste |\n"
                    "| 🟠 **BASSA** | 200 | 200 ordini virtuali | Expert judgment | ±3.2% | **Pre-lancio** (fase attuale) |\n"
                    "| 🟡 **MEDIA** | 1.000 | 1.000 ordini virtuali | Modello precedente | ±1.4% | Primi 3-6 mesi vendita |\n"
                    "| 🟢 **ALTA** | 5.000 | 5.000 ordini virtuali | Dati reali | ±0.6% | Model mix / dopo 1 anno vendita |"
                )
                st.caption(
                    "**Regola pratica**: man mano che arrivano ordini reali, "
                    "alza il livello di confidenza. Pre-lancio → BASSA, "
                    "primi mesi → MEDIA, dopo 1 anno → ALTA."
                )

            # ── Confronto visivo tra i 4 livelli ─────────────────────────────────────────────────────
            st.markdown("### Impatto sulla Simulazione")
            st.caption(
                "Impatto calcolato su un TR del 10% con domanda mensile configurabile. "
                "Valori piu bassi di CV e Safety Stock indicano minor incertezza."
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
                _lv_desc = {
                    "NULLA": "Nessun precedente",
                    "BASSA": "Expert judgment",
                    "MEDIA": "Modello precedente",
                    "ALTA":  "Storico consolidato",
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
                            f'VSS = {_M:,} ordini equiv.</div>'
                            f'<div style="color:#ccc;font-size:0.72rem;font-style:italic">'
                            f'{_lv_desc[_lv]}</div>'
                            f'<hr style="border-color:#444;margin:0.4rem 0">'
                            f'<div style="color:#fff;font-size:0.88rem">'
                            f'CV: <b>{_m["cv"]:.1f}%</b></div>'
                            f'<div style="color:#fff;font-size:0.88rem">'
                            f'TR range (P{_percentile_sens}): '
                            f'<b>{max(0.0, _tr_ref*100 - _z_sens*_m["sigma_tr"]):.1f}%'
                            f' - {min(100.0, _tr_ref*100 + _z_sens*_m["sigma_tr"]):.1f}%</b></div>'
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
                    # Ordine crescente per TR numerico
                    _tr_order = sorted(_df_raw["TR_pct"].unique())
                    _tr_label_order = [f"{int(t)}%" if t == int(t) else f"{t:.0f}%" for t in _tr_order]

                    # ── PIVOT 1: Safety Stock (unita buffer) ───────────────────────
                    st.markdown(f"#### Safety Stock necessario — P{_percentile_sens} (unita per {_monthly_demand_sens} moto/mese)")
                    st.caption("Quante unita extra devi tenere in magazzino per coprire l'incertezza della simulazione.")

                    _piv_ss = _df_raw.pivot(index="TR storico", columns="Livello", values="SS")
                    _piv_ss = _piv_ss.reindex(index=_tr_label_order, columns=_lv_order_cols)
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
                    _piv_range = _piv_range.reindex(index=_tr_label_order, columns=_lv_order_cols)
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
                st.markdown("**Confidenza Take Rate**")
                st.caption("VSS = ordini equivalenti. Piu alto = simulazione piu stabile.")
                _aff_format = {
                    "NULLA": "NULLA (VSS=50) — nessun precedente",
                    "BASSA": "BASSA (VSS=200) — expert judgment",
                    "MEDIA": "MEDIA (VSS=1000) — modello precedente",
                    "ALTA":  "ALTA (VSS=5000) — storico consolidato",
                }
                st.selectbox(
                    "Model Mix (L0)",
                    options=_AFF_LEVELS,
                    format_func=lambda x: _aff_format.get(x, x),
                    index=_AFF_LEVELS.index(st.session_state.get("ss_aff_model", _aff_def_model)),
                    help="Confidenza sulla ripartizione tra varianti (MSV4/MSV4S/MSV4PP). Consigliato: ALTA.",
                    key="ss_aff_model",
                )
                st.selectbox(
                    "Optional (L1)",
                    options=_AFF_LEVELS,
                    format_func=lambda x: _aff_format.get(x, x),
                    index=_AFF_LEVELS.index(st.session_state.get("ss_aff_optional", _aff_def_optional)),
                    help="Confidenza sui TR delle opzioni/accessori. Pre-lancio: BASSA. Con dati vendita: MEDIA/ALTA.",
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

            # File rigenerabili da Step 1: se il toggle Step 1 è attivo,
            # BOM150 e Mappatura vengono prodotti da enrich_bom_version
            # e non bloccano il lancio.
            _step1_generates = {"Matrix", "Unique Rows"}
            if run_step0:
                _step1_generates |= {"BOM150", "Mappatura"}

            _logistica_dir = INPUT_DIR / "Dati Logistica"
            _dash_files = list(_logistica_dir.glob("*Dashboard Supplier*.xlsx")) if _logistica_dir.exists() else []
            _logistica_required = {
                "Gates": _logistica_dir / "Gates MTSV4.xlsx",
            }
            _missing_logistica = [k for k, p in _logistica_required.items() if not file_exists(p)]
            if not _dash_files:
                _missing_logistica.append("Dashboard Supplier")

            _miss_ss_run = [n for n in SS_REQUIRED
                            if not file_exists(INPUT_FILES[n])
                            and n not in _step1_generates
                            and n not in ("Lead Times",)]

            if "Lead Times" in SS_REQUIRED and _missing_logistica:
                _miss_ss_run.append("Lead Times (dati Logistica: " + ", ".join(_missing_logistica) + ")")

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
                    st.warning(
                        f"Impossibile avviare: mancano {', '.join(_miss_ss_run)}. "
                        f"Abilita **Step 1 · Rigenera Multi-BOM** per generare i BOM arricchiti dai file grezzi."
                    )
                else:
                    _sv_pc1, _sv_pc2, _sv_pc3, _sv_pc4 = st.columns(4)
                    _sv_pc1.metric("Run MC",      st.session_state.get("ss_n_runs", 200))
                    _sv_pc2.metric("Mese",        st.session_state.get("ss_start_month", demand_months[0] if demand_months else f"Offset {start_month_offset}"))
                    _sv_pc3.metric("Percentile",  f"P{st.session_state.get('ss_percentile', 99)}")
                    _sv_pc4.metric("Confidenza TR", f"{st.session_state.get('ss_aff_model', aff_model)}/{st.session_state.get('ss_aff_optional', aff_optional)}")

    # SS — TAB 2: VISUALIZZA RISULTATI
    # ########################################################################

    with tab_view:

        # Sub-tab interni
        st_analytics, st_results, st_monthly, st_tr, st_alloc, st_tr_impl = st.tabs([
            "Analisi Integrata",
            "Risultati Safety Stock",
            "Safety Mensile",
            "Take Rate",
            "Allocazione Costi Optional",
            "TR Impliciti",
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
                    fa, fb, fc, fd = st.columns(4)
                    with fa:
                        g_search = st.text_input("Cerca SKU / Descrizione",
                                                 key="an_search", placeholder="es: RAGGIO, 4961...")
                    with fb:
                        _abc_vals = sorted(df_m["ABC"].dropna().unique().tolist())
                        g_abc = st.multiselect("Classe ABC", _abc_vals, default=_abc_vals, key="an_abc")
                    with fc:
                        if "Lead_Time_Months_Ceil" in df_m.columns:
                            import math
                            _lt_min = int(df_m["Lead_Time_Months_Ceil"].min())
                            _lt_max = math.ceil(df_m["Lead_Time_Months_Ceil"].max())
                            g_lt = st.slider("Lead Time (mesi)", _lt_min, _lt_max,
                                             (_lt_min, _lt_max), key="an_lt")
                        else:
                            g_lt = (0, 99)
                    with fd:
                        _has_solo_modello = "SOLO_MODELLO" in df_m.columns
                        g_solo_modello = st.selectbox(
                            "Solo Modello",
                            ["Tutti", "Solo invarianti modello", "Escludi invarianti modello"],
                            index=0,
                            key="an_solo_modello",
                            help=(
                                "Invarianti modello = SKU con tutte le caratteristiche optional/bundle = No "
                                "ma non presenti in tutti i modelli. Dipendono solo dal modello."
                            ),
                            disabled=not _has_solo_modello,
                        )

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
                if _has_solo_modello and g_solo_modello == "Solo invarianti modello":
                    df_f = df_f[df_f["SOLO_MODELLO"] == True]
                elif _has_solo_modello and g_solo_modello == "Escludi invarianti modello":
                    df_f = df_f[df_f["SOLO_MODELLO"] != True]

                _n_solo_mod = int(df_m["SOLO_MODELLO"].sum()) if "SOLO_MODELLO" in df_m.columns else 0
                st.caption(
                    f"**{len(df_f)} SKU** nel dataset filtrato (totale: {len(df_m)}"
                    + (f" · {_n_solo_mod} invarianti modello" if _n_solo_mod else "")
                    + ")"
                )

                # ════════════════════════════════════════════════════════════════
                # SEZIONE 1 — STRATEGICA
                # ════════════════════════════════════════════════════════════════
                st.markdown("---")
                st.markdown("## Visione Strategica")

                kp1, kp2, kp3 = st.columns(3)

                with kp1:
                    if _has_prezzo:
                        _cap = df_f["prezzo_safety"].sum()
                        st.metric("Capitale Immobilizzato", f"€ {_cap:,.0f}",
                                  help="Valore totale del buffer di sicurezza (SKU filtrati)")
                    else:
                        st.metric("Capitale Immobilizzato", "N/D")

                with kp2:
                    _ss_qty_col = next(
                        (c for c in df_f.columns
                         if "safety_stock_p" in c.lower() and "total" in c.lower()), None
                    )
                    if _ss_qty_col:
                        _qty_tot = df_f[_ss_qty_col].sum()
                        st.metric("Quantità Totale Scorte",
                                  f"{_qty_tot:,.0f} pz",
                                  help="Somma delle safety stock in pezzi fisici (SKU filtrati)")
                    else:
                        st.metric("Quantità Totale Scorte", "N/D")

                with kp3:
                    if "SOLO_MODELLO" in df_f.columns:
                        _n_sm = int(df_f["SOLO_MODELLO"].sum())
                        _pct_sm = _n_sm / len(df_f) * 100 if len(df_f) > 0 else 0
                        st.metric(
                            "Invarianti Modello",
                            f"{_n_sm} SKU",
                            delta=f"{_pct_sm:.1f}% del totale filtrato",
                            delta_color="off",
                            help=(
                                "SKU con tutte le caratteristiche optional/bundle = No "
                                "ma non presenti in tutti i modelli. "
                                "Dipendono solo dalla variante modello, non dalle opzioni."
                            ),
                        )
                    else:
                        st.metric("Invarianti Modello", "N/D")

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
                            "(candidato a riduzione confidenza TR o lead time). "
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
                    "Utile per valutare l'impatto di diversi livelli di confidenza TR o di run differenti."
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
                                 and "total" not in c.lower()
                                 and "_" in c]
                month_labels = []
                for c in ss_month_cols:
                    parts = c.split("_")
                    if len(parts) >= 5:
                        month_labels.append(f"{parts[-2]}-{parts[-1]}")
                    elif len(parts) >= 4:
                        month_labels.append(parts[-1])
                    else:
                        month_labels.append(c)

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

        # ====================================================================
        # SUB-TAB E — ALLOCAZIONE COSTI OPTIONAL
        # ====================================================================
        with st_alloc:
            st.subheader("Allocazione Costi Safety Stock per Optional")
            st.caption(
                "Ripartisce il costo Safety Stock di ogni componente tra gli optional "
                "e le configurazioni che ne determinano la domanda, usando pesi TR. "
                "Per righe con N optional attivi il costo viene suddiviso in parti uguali."
            )

            # ── Import funzioni allocation ───────────────────────────────────
            try:
                import sys as _sys
                import importlib as _imp
                _base_path = str(BASE_DIR)
                if _base_path not in _sys.path:
                    _sys.path.insert(0, _base_path)
                _alloc_mod = _imp.import_module("optional_cost_allocation")
                _build_tr_lookup  = _alloc_mod.build_tr_lookup
                _allocate         = _alloc_mod.allocate
                _make_summary     = _alloc_mod.make_summary
                _make_pivot       = _alloc_mod.make_pivot
                _alloc_available  = True
            except Exception as _e:
                st.error(f"Modulo optional_cost_allocation non trovato: {_e}")
                _alloc_available = False

            if _alloc_available:
                # ── Carica dati SS ───────────────────────────────────────────
                _al_ss_path = OUTPUT_FILES["Safety Stock"]
                _al_json    = OUTPUT_DIR / "parsed_data_complete.json"
                _al_sku_paths = [
                    INPUT_DIR / "tutte_righe_univoche_V2.xlsx",
                    INPUT_DIR / "_sku_tmp.xlsx",
                ]

                _al_ready = (
                    file_exists(_al_ss_path)
                    and _al_json.exists()
                    and any(p.exists() for p in _al_sku_paths)
                )

                if not _al_ready:
                    st.info(
                        "File necessari non trovati.  \n"
                        "Esegui prima la pipeline Safety Stock e assicurati che "
                        "`parsed_data_complete.json` sia presente in Output/."
                    )
                else:
                    # Bottone ricalcola
                    _run_alloc = st.button(
                        "Calcola / Aggiorna Allocazione",
                        key="btn_alloc",
                        help="Ricalcola la distribuzione dei costi SS per optional",
                    )

                    if _run_alloc or "alloc_detail" not in st.session_state:
                        with st.spinner("Calcolo allocazione in corso..."):
                            try:
                                import json as _json
                                with open(_al_json, encoding="utf-8") as _f:
                                    _parsed = _json.load(_f)
                                _model_tr, _char_lookup = _build_tr_lookup(_parsed)

                                _al_ss_df = pd.read_excel(str(_al_ss_path))
                                _al_ss_df = _al_ss_df[
                                    _al_ss_df["prezzo_safety"].notna()
                                    & (_al_ss_df["prezzo_safety"] > 0)
                                ].copy()

                                # SKU catalog con fallback
                                _al_sku_df = None
                                for _p in _al_sku_paths:
                                    try:
                                        _al_sku_df = pd.read_excel(str(_p))
                                        break
                                    except Exception:
                                        continue
                                if _al_sku_df is None:
                                    st.error("Impossibile leggere il catalogo SKU.")
                                    st.stop()

                                _detail_df = _allocate(
                                    _al_ss_df, _al_sku_df, _model_tr, _char_lookup
                                )
                                st.session_state["alloc_detail"] = _detail_df
                                st.session_state["alloc_ss_df"]  = _al_ss_df
                                st.toast("Allocazione completata.")
                            except Exception as _e:
                                st.error(f"Errore allocazione: {_e}")
                                st.stop()

                    _detail_df = st.session_state.get("alloc_detail")
                    _al_ss_df  = st.session_state.get("alloc_ss_df")

                    if _detail_df is not None and not _detail_df.empty:

                        # ── KPI ─────────────────────────────────────────────
                        _tot_ss    = _al_ss_df["prezzo_safety"].sum()
                        _tot_alloc = _detail_df["Costo_SS_Allocato"].sum()
                        _n_comp    = _detail_df["SKU"].nunique()
                        _n_opts    = _detail_df[
                            _detail_df["Tipo_Optional"].isin(["Pack", "Configurazione"])
                        ]["Colonna_Optional"].nunique()

                        _kc1, _kc2, _kc3, _kc4 = st.columns(4)
                        _kc1.metric("Costo SS Totale", f"€ {_tot_ss:,.0f}")
                        _kc2.metric("Costo Allocato", f"€ {_tot_alloc:,.0f}")
                        _kc3.metric("Componenti Coperti", f"{_n_comp}")
                        _kc4.metric("Optional / Config Attivi", f"{_n_opts}")

                        st.markdown("---")

                        # ── Filtro Tipo ──────────────────────────────────────
                        _tipo_sel = st.radio(
                            "Visualizza:",
                            ["Pack Optional", "Configurazione", "Entrambi"],
                            horizontal=True,
                            key="alloc_tipo_radio",
                        )
                        _tipo_map = {
                            "Pack Optional":  ["Pack"],
                            "Configurazione": ["Configurazione"],
                            "Entrambi":       ["Pack", "Configurazione"],
                        }
                        _tipi = _tipo_map[_tipo_sel]
                        _sub  = _detail_df[_detail_df["Tipo_Optional"].isin(_tipi)].copy()

                        # ── Riepilogo per Optional ───────────────────────────
                        st.markdown("### Costo Safety Stock per Optional")
                        _summ = (
                            _sub.groupby(["Tipo_Optional", "Colonna_Optional"])
                            .agg(
                                Costo_SS=("Costo_SS_Allocato", "sum"),
                                N_SKU=("SKU", "nunique"),
                            )
                            .reset_index()
                            .sort_values("Costo_SS", ascending=False)
                        )
                        _summ["% sul Totale"] = (
                            _summ["Costo_SS"] / _tot_alloc * 100
                        ).round(1)
                        _summ["Costo_SS"] = _summ["Costo_SS"].round(2)

                        _col_bar, _col_tbl = st.columns([3, 2])
                        with _col_bar:
                            _fig_alloc = px.bar(
                                _summ,
                                x="Costo_SS",
                                y="Colonna_Optional",
                                color="Tipo_Optional",
                                orientation="h",
                                title="Costo SS allocato per optional / configurazione",
                                labels={
                                    "Costo_SS": "EUR",
                                    "Colonna_Optional": "",
                                    "Tipo_Optional": "Tipo",
                                },
                                color_discrete_map={
                                    "Pack":           "#E3000F",
                                    "Configurazione": "#2E4F8A",
                                },
                                text_auto=".0f",
                            )
                            _fig_alloc.update_layout(
                                yaxis={"categoryorder": "total ascending"},
                                height=max(350, len(_summ) * 28 + 80),
                                margin={"l": 10, "r": 10, "t": 40, "b": 10},
                            )
                            st.plotly_chart(_fig_alloc, use_container_width=True)

                        with _col_tbl:
                            st.dataframe(
                                _summ.rename(columns={
                                    "Tipo_Optional":     "Tipo",
                                    "Colonna_Optional":  "Optional",
                                    "Costo_SS":          "Costo SS (EUR)",
                                    "N_SKU":             "N. SKU",
                                }),
                                use_container_width=True,
                                hide_index=True,
                            )

                        st.markdown("---")

                        # ── Drill-down per Optional selezionato ──────────────
                        st.markdown("### Dettaglio per Optional")
                        _all_opts = sorted(_sub["Colonna_Optional"].unique())
                        _sel_opt  = st.selectbox(
                            "Seleziona Optional:",
                            _all_opts,
                            key="alloc_opt_sel",
                        )
                        _drill = (
                            _sub[_sub["Colonna_Optional"] == _sel_opt]
                            .groupby(["SKU", "Descrizione", "Valore_Optional"])
                            .agg(
                                Costo_Allocato=("Costo_SS_Allocato", "sum"),
                                SS_Cost_Totale=("SS_Cost_Totale", "first"),
                                SS_Unita=("SS_Unita", "first"),
                            )
                            .reset_index()
                            .sort_values("Costo_Allocato", ascending=False)
                        )
                        _drill["% su componente"] = (
                            _drill["Costo_Allocato"] / _drill["SS_Cost_Totale"] * 100
                        ).round(1)

                        _dc1, _dc2 = st.columns([1, 2])
                        _dc1.metric(
                            f"Totale SS attribuito a '{_sel_opt}'",
                            f"€ {_drill['Costo_Allocato'].sum():,.0f}",
                        )
                        _dc2.metric(
                            "Componenti coinvolti",
                            f"{_drill['SKU'].nunique()}",
                        )

                        # Grafico top componenti per questo optional
                        _top_drill = _drill.head(20)
                        _fig_drill = px.bar(
                            _top_drill,
                            x="Costo_Allocato",
                            y="SKU",
                            orientation="h",
                            color="Valore_Optional",
                            title=f"Top-20 componenti per '{_sel_opt}'",
                            labels={"Costo_Allocato": "EUR", "SKU": ""},
                            hover_data=["Descrizione", "% su componente"],
                            text_auto=".0f",
                        )
                        _fig_drill.update_layout(
                            yaxis={"categoryorder": "total ascending"},
                            height=max(300, len(_top_drill) * 26 + 80),
                        )
                        st.plotly_chart(_fig_drill, use_container_width=True)

                        st.dataframe(
                            _drill[[
                                "SKU", "Descrizione", "Valore_Optional",
                                "Costo_Allocato", "SS_Cost_Totale",
                                "% su componente", "SS_Unita",
                            ]].rename(columns={
                                "Costo_Allocato":  "Costo Allocato (EUR)",
                                "SS_Cost_Totale":  "Costo SS Totale (EUR)",
                                "SS_Unita":        "SS Unita",
                                "Valore_Optional": "Valore",
                            }),
                            use_container_width=True,
                            hide_index=True,
                        )

                        st.markdown("---")

                        # ── Download ─────────────────────────────────────────
                        st.markdown("### Export")
                        _al_out = OUTPUT_DIR / "optional_cost_allocation.xlsx"
                        if _al_out.exists():
                            with open(str(_al_out), "rb") as _fh:
                                st.download_button(
                                    label="Scarica Excel completo (tutti i fogli)",
                                    data=_fh.read(),
                                    file_name="optional_cost_allocation.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument"
                                         ".spreadsheetml.sheet",
                                    key="dl_alloc",
                                )
                        else:
                            # Genera al volo
                            import io as _io
                            _buf = _io.BytesIO()
                            with pd.ExcelWriter(_buf, engine="openpyxl") as _wr:
                                _detail_df.to_excel(
                                    _wr, sheet_name="Dettaglio", index=False
                                )
                                for _t in ["Pack", "Configurazione", "Base"]:
                                    _s = _make_summary(_detail_df, _t)
                                    if not _s.empty:
                                        _s.to_excel(
                                            _wr, sheet_name=f"Riepilogo_{_t}", index=False
                                        )
                            st.download_button(
                                label="Scarica Excel (dettaglio)",
                                data=_buf.getvalue(),
                                file_name="optional_cost_allocation.xlsx",
                                mime="application/vnd.openxmlformats-officedocument"
                                     ".spreadsheetml.sheet",
                                key="dl_alloc_inline",
                            )

        # ====================================================================
        # SUB-TAB F — TR IMPLICITI
        # ====================================================================

        with st_tr_impl:
            st.subheader("Take Rate Impliciti nella Safety Stock")
            st.caption(
                "Calcola per ogni SKU il TR nominale (domanda media stimata) e "
                "il TR massimo (coperto dalla safety al 99°percentile) impliciti "
                "nella simulazione Monte Carlo. "
                "Formula: TR = domanda / (N_moto nel lead time × qty per moto)."
            )

            # ── Import modulo tr_implied ─────────────────────────────────────
            try:
                import importlib as _imp_tri
                import sys as _sys_tri
                _base_tri = str(BASE_DIR)
                if _base_tri not in _sys_tri.path:
                    _sys_tri.path.insert(0, _base_tri)
                _tri_mod = _imp_tri.import_module("tr_implied")
                if hasattr(_imp_tri, 'reload'):
                    try:
                        _imp_tri.reload(_tri_mod)
                    except Exception:
                        pass
                # Usa versione MENSILE (corretta) invece della globale
                _compute_implied_tr = _tri_mod.compute_implied_tr_monthly
                _tri_available      = True
            except Exception as _e_tri:
                st.error(f"Modulo tr_implied non trovato: {_e_tri}")
                _tri_available = False

            if _tri_available:
                # ── Percorsi file ────────────────────────────────────────────
                # Usa file MENSILE (sku_safety_stock_PER_MESE_V2.xlsx) per calcolo mese-per-mese
                _tri_sku_path      = OUTPUT_DIR / "sku_safety_stock_PER_MESE_V2.xlsx"
                _tri_forecast_path = INPUT_DIR / "Total_demand.xlsx"
                _tri_output_path   = OUTPUT_DIR / "tr_implied_monthly.xlsx"

                # Fallback file mensile
                _tri_sku_fallbacks = [
                    OUTPUT_DIR / "sku_safety_stock_PER_MESE_V3.xlsx",
                    OUTPUT_DIR / "sku_safety_stock_PER_MESE.xlsx",
                ]

                _tri_sku_ok      = file_exists(_tri_sku_path) or any(
                    p.exists() for p in _tri_sku_fallbacks
                )
                _tri_forecast_ok = _tri_forecast_path.exists()

                if not _tri_sku_ok:
                    st.info(
                        "File Safety Stock non trovato.  \n"
                        "Esegui prima la pipeline Safety Stock."
                    )
                elif not _tri_forecast_ok:
                    st.warning(
                        f"File forecast `{_tri_forecast_path.name}` non trovato in Input/.  \n"
                        "Assicurati che `Total_demand.xlsx` sia presente."
                    )
                else:
                    # ── Mese di inizio (allineato alla pipeline Safety Stock) ──
                    if demand_months:
                        _tri_def_month = st.session_state.get(
                            "ss_start_month", demand_months[0]
                        )
                        _tri_def_month = (
                            _tri_def_month
                            if _tri_def_month in demand_months
                            else demand_months[0]
                        )
                        _tri_month = st.selectbox(
                            "Mese di inizio (stesso della pipeline SS)",
                            options=demand_months,
                            index=demand_months.index(_tri_def_month),
                            help="Primo mese del forecast da includere nel calcolo. "
                                 "Deve corrispondere al mese di inizio usato nella pipeline Safety Stock.",
                            key="tri_start_month",
                        )
                        _tri_offset = demand_months.index(_tri_month)
                    else:
                        _tri_offset = st.number_input(
                            "Start Month Offset",
                            min_value=0, max_value=24,
                            value=st.session_state.get("ss_start_offset", 0),
                            step=1,
                            help="Fallback: offset numerico se Total_demand.xlsx non e' caricato.",
                            key="tri_offset",
                        )

                    # ── Bottone ricalcola ────────────────────────────────────
                    _run_tri = st.button(
                        "Calcola / Aggiorna TR Impliciti",
                        key="btn_tri",
                        help="Ricalcola TR nominale e TR massimo per tutti gli SKU",
                    )

                    # Invalida cache se le colonne sono vecchie (non-mensili)
                    if "tri_df" in st.session_state:
                        _cached_cols = st.session_state["tri_df"].columns.tolist()
                        # Verifica se sono colonne mensili (nuove) o globali (vecchie)
                        has_monthly = any("TR_Medio_%_" in c for c in _cached_cols)
                        if not has_monthly or "TR_Medio_%_Media" not in _cached_cols:
                            del st.session_state["tri_df"]

                    if _run_tri or "tri_df" not in st.session_state:
                        with st.spinner("Calcolo TR impliciti mensili in corso..."):
                            try:
                                _tri_df = _compute_implied_tr(
                                    sku_monthly_file=str(_tri_sku_path),
                                    forecast_file=str(_tri_forecast_path),
                                    start_month_offset=int(_tri_offset),
                                    output_file=str(_tri_output_path),
                                )
                                st.session_state["tri_df"] = _tri_df
                                st.toast("TR Impliciti mensili calcolati e salvati.")
                            except Exception as _e_tri_run:
                                st.error(f"Errore calcolo TR: {_e_tri_run}")
                                st.stop()

                    _tri_df = st.session_state.get("tri_df")

                    if _tri_df is not None and not _tri_df.empty:

                        # ── KPI ─────────────────────────────────────────────
                        _tri_k1, _tri_k2, _tri_k3, _tri_k4 = st.columns(4)
                        _tri_k1.metric(
                            "SKU elaborati",
                            f"{len(_tri_df)}"
                        )
                        _tri_k2.metric(
                            "TR Medio (media mesi)",
                            f"{_tri_df['TR_Medio_%_Media'].mean():.1f}%"
                        )
                        _tri_k3.metric(
                            "TR Massimo (media mesi)",
                            f"{_tri_df['TR_Massimo_%_Media'].mean():.1f}%"
                        )
                        # Calcola delta medio da colonne mensili
                        _tri_delta_cols = [c for c in _tri_df.columns if c.startswith('TR_Massimo_%_')
                                          and not c.endswith('_Media')]
                        _tri_deltas = _tri_df[_tri_delta_cols].mean(axis=1) - \
                                      _tri_df[[c.replace('Massimo', 'Medio') for c in _tri_delta_cols]].mean(axis=1)
                        _tri_k4.metric(
                            "Delta medio (safety mesi)",
                            f"{_tri_deltas.mean():.1f} pp"
                        )

                        st.markdown("---")

                        # ── Filtri ───────────────────────────────────────────
                        _tri_c1, _tri_c2 = st.columns(2)
                        with _tri_c1:
                            _tri_tr_min, _tri_tr_max = st.slider(
                                "Filtro TR Nominale Media (%)",
                                min_value=0.0,
                                max_value=float(
                                    _tri_df["TR_Medio_%_Media"].fillna(0).max()
                                ),
                                value=(
                                    0.0,
                                    float(_tri_df["TR_Medio_%_Media"].fillna(0).max()),
                                ),
                                step=0.5,
                                key="tri_slider",
                            )
                        with _tri_c2:
                            _tri_search = st.text_input(
                                "Cerca SKU / Descrizione",
                                value="",
                                key="tri_search",
                            )

                        _tri_filt = _tri_df[
                            (_tri_df["TR_Medio_%_Media"].fillna(0) >= _tri_tr_min)
                            & (_tri_df["TR_Medio_%_Media"].fillna(0) <= _tri_tr_max)
                        ].copy()

                        if _tri_search.strip():
                            _q = _tri_search.strip().lower()
                            _tri_filt = _tri_filt[
                                _tri_filt["SKU"].astype(str).str.lower().str.contains(_q)
                                | _tri_filt["Description"].astype(str).str.lower().str.contains(_q)
                            ]

                        # ── Scatter TR nominale vs TR massimo (media su mesi) ────────────────
                        st.markdown(f"### Distribuzione TR (media mesi) — {len(_tri_filt)} SKU selezionati")

                        # Calcola delta per il grafico
                        _tri_filt_plot = _tri_filt.copy()
                        _tri_filt_plot['Delta_TR_%_Media'] = (
                            (_tri_filt_plot['TR_Massimo_%_Media'] - _tri_filt_plot['TR_Medio_%_Media'])
                            / (_tri_filt_plot['TR_Medio_%_Media'] + 0.001) * 100
                        ).round(2)

                        _tri_fig = px.scatter(
                            _tri_filt_plot,
                            x="TR_Medio_%_Media",
                            y="TR_Massimo_%_Media",
                            hover_data=["SKU", "Description", "Lead_Time_Months"],
                            color="Delta_TR_%_Media",
                            color_continuous_scale="RdYlGn_r",
                            title="TR Nominale vs TR Massimo (media su mesi)",
                            labels={
                                "TR_Medio_%_Media": "TR Medio (%) — media mesi",
                                "TR_Massimo_%_Media":  "TR Massimo (%) — media mesi",
                                "Delta_TR_%_Media":   "Delta TR (%)",
                            },
                        )
                        # Linea diagonale di riferimento
                        _tri_ax_max = max(
                            _tri_filt_plot["TR_Medio_%_Media"].fillna(0).max(),
                            _tri_filt_plot["TR_Massimo_%_Media"].fillna(0).max(),
                        ) * 1.05
                        _tri_fig.add_shape(
                            type="line",
                            x0=0, y0=0, x1=_tri_ax_max, y1=_tri_ax_max,
                            line={"color": "grey", "dash": "dot", "width": 1},
                        )
                        _tri_fig.update_layout(
                            height=480,
                            margin={"l": 10, "r": 10, "t": 40, "b": 10},
                        )
                        st.plotly_chart(_tri_fig, use_container_width=True)

                        # ── Scatter Delta TR vs TR nominale ──────────────────
                        st.markdown("### Delta TR (%) in funzione del TR Nominale — media mesi")
                        st.caption(
                            "Ogni punto è uno SKU. L'asse X mostra il TR nominale medio su tutti i mesi, "
                            "l'asse Y il delta (%) aggiunto dalla safety. "
                            "SKU in alto a sinistra hanno safety elevata rispetto alla domanda."
                        )
                        _tri_fig2 = px.scatter(
                            _tri_filt_plot,
                            x="TR_Medio_%_Media",
                            y="Delta_TR_%_Media",
                            hover_data=["SKU", "Description",
                                        "TR_Massimo_%_Media",
                                        "Lead_Time_Months"],
                            color="Lead_Time_Months",
                            color_continuous_scale="Blues",
                            title="Delta TR % (safety relativa) vs TR Nominale — media mesi",
                            labels={
                                "TR_Medio_%_Media":    "TR Nominale Media (%)",
                                "Delta_TR_%_Media":       "Delta TR — safety / nominale (%)",
                                "Lead_Time_Months": "Lead Time (mesi)",
                            },
                        )
                        _tri_fig2.update_layout(
                            height=420,
                            margin={"l": 10, "r": 10, "t": 40, "b": 10},
                        )
                        st.plotly_chart(_tri_fig2, use_container_width=True)

                        # ── Tabella dettaglio ────────────────────────────────
                        st.markdown("### Dettaglio TR per SKU (media mensile)")

                        # Colonne base sempre presenti
                        _tri_show_cols = [
                            "SKU", "Description", "Lead_Time_Months",
                            "Quantity_Per_Bike",
                            "TR_Medio_%_Media", "TR_Massimo_%_Media",
                        ]

                        _tri_tab_data = _tri_filt_plot[_tri_show_cols].copy()
                        _tri_tab_data = _tri_tab_data.sort_values(
                            "TR_Medio_%_Media", ascending=False
                        ).reset_index(drop=True)

                        st.dataframe(
                            _tri_tab_data,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "TR_Medio_%_Media": st.column_config.NumberColumn(
                                    "TR Medio (%)", format="%.2f %%"
                                ),
                                "TR_Massimo_%_Media": st.column_config.NumberColumn(
                                    "TR Max (%)", format="%.2f %%"
                                ),
                                "Lead_Time_Months": st.column_config.NumberColumn(
                                    "Lead Time (mesi)", format="%.2f"
                                ),
                            },
                        )

                        # Espansione: mostra TR per ogni mese
                        st.markdown("#### TR Mensili per SKU")
                        _tri_month_cols = [c for c in _tri_filt_plot.columns
                                          if c.startswith('TR_Medio_%_') and not c.endswith('_Media')]
                        if _tri_month_cols:
                            _tri_months_sorted = sorted(_tri_month_cols)
                            _tri_monthly_view = _tri_filt_plot[["SKU", "Description"] + _tri_months_sorted].copy()
                            _tri_monthly_view = _tri_monthly_view.sort_values(
                                "SKU", ascending=True
                            ).reset_index(drop=True)
                            st.dataframe(
                                _tri_monthly_view,
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    col: st.column_config.NumberColumn(
                                        col.replace("TR_Medio_%_", ""), format="%.1f %%"
                                    )
                                    for col in _tri_months_sorted
                                },
                            )
                        else:
                            st.info("Nessun dato mensile disponibile")

                        st.markdown("---")

                        # ── Download ─────────────────────────────────────────
                        st.markdown("### Export")
                        if _tri_output_path.exists():
                            with open(str(_tri_output_path), "rb") as _fh_tri:
                                st.download_button(
                                    label="Scarica Excel TR Impliciti",
                                    data=_fh_tri.read(),
                                    file_name="tr_implied.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument"
                                         ".spreadsheetml.sheet",
                                    key="dl_tri",
                                )
                        else:
                            import io as _io_tri
                            _buf_tri = _io_tri.BytesIO()
                            with pd.ExcelWriter(_buf_tri, engine="openpyxl") as _wr_tri:
                                _tri_df.to_excel(
                                    _wr_tri, sheet_name="TR_Impliciti", index=False
                                )
                            st.download_button(
                                label="Scarica Excel TR Impliciti",
                                data=_buf_tri.getvalue(),
                                file_name="tr_implied.xlsx",
                                mime="application/vnd.openxmlformats-officedocument"
                                     ".spreadsheetml.sheet",
                                key="dl_tri_inline",
                            )

    # ########################################################################
    # SS — TAB 3: ROLLING SIMULATION
    # ########################################################################

    with tab_rolling:

        st.markdown("## 🔄 Rolling Simulation")
        st.caption(
            "Esegui la simulazione Monte Carlo + SKU Matching in modo rolling: "
            "ogni iterazione parte da un mese successivo, finché ci sono abbastanza dati di domanda. "
            "Output: singola tabella con colonna **mese_lancio** per confrontare i safety stock nel tempo."
        )

        # ── Parametri rolling ─────────────────────────────────────────────
        st.markdown("### Parametri")
        _rc1, _rc2, _rc3, _rc4 = st.columns(4)
        with _rc1:
            _roll_runs = st.number_input(
                "Run Monte Carlo",
                min_value=10, max_value=2000, value=100, step=10,
                help="Più run = più preciso, ma più lento",
                key="roll_n_runs",
            )
        with _rc2:
            _roll_pct = st.selectbox(
                "Percentile Safety Stock",
                options=[90, 95, 99],
                index=2,
                key="roll_percentile",
            )
        with _rc3:
            _roll_max_iter = st.number_input(
                "Max iterazioni rolling",
                min_value=1, max_value=36, value=None,
                help="Numero massimo mesi da iterare. None = tutte le possibili",
                key="roll_max_iterations",
            )
        with _rc4:
            _roll_keep_files = st.checkbox(
                "Mantieni file per ogni iterazione",
                value=False,
                help="Se attivo, salva anche un file Excel per ogni mese di lancio",
                key="roll_keep_files",
            )

        # ── Affidabilità TR ───────────────────────────────────────────────
        st.markdown("### Affidabilità Take Rate")
        _raff1, _raff2 = st.columns(2)
        with _raff1:
            _roll_aff_model = st.selectbox(
                "Affidabilità Model Mix (L0)",
                options=_AFF_LEVELS,
                index=_AFF_LEVELS.index(aff_model) if aff_model in _AFF_LEVELS else 1,
                help="VSS per il model mix — uguale alla pipeline SS per confronto coerente",
                key="roll_aff_model",
            )
        with _raff2:
            _roll_aff_optional = st.selectbox(
                "Affidabilità Optional (L1)",
                options=_AFF_LEVELS,
                index=_AFF_LEVELS.index(aff_optional) if aff_optional in _AFF_LEVELS else 1,
                help="VSS per i take rate optional — uguale alla pipeline SS per confronto coerente",
                key="roll_aff_optional",
            )

        st.markdown("---")

        # ── Avvia rolling ─────────────────────────────────────────────────
        _col_btn, _col_status = st.columns([1, 3])
        with _col_btn:
            _run_rolling = st.button(
                "▶ Avvia Rolling",
                type="primary",
                use_container_width=True,
                disabled=st.session_state.get("_roll_running", False),
                key="btn_run_rolling",
            )

        _roll_log_placeholder = st.empty()
        _roll_progress = st.progress(0)
        _roll_status   = st.empty()

        if _run_rolling and not st.session_state.get("_roll_running", False):
            st.session_state["_roll_running"] = True
            _roll_log_lines: list[str] = []

            def _rlog(msg: str):
                _roll_log_lines.append(msg)
                _roll_log_placeholder.code("\n".join(_roll_log_lines[-60:]), language="")

            try:
                import importlib
                import mc_simulator_rolling as _rolling_mod
                importlib.reload(_rolling_mod)

                _rlog("Avvio rolling simulation...")
                _roll_status.text("Rolling in corso...")
                _roll_progress.progress(5)

                _t_roll = time.time()
                _rlog(f"  Affidabilità → Model: {_roll_aff_model} (VSS={_AFF_MAP.get(_roll_aff_model,0):.0f})"
                      f"  |  Optional: {_roll_aff_optional} (VSS={_AFF_MAP.get(_roll_aff_optional,0):.0f})")
                _roll_result = _rolling_mod.run_rolling_simulation(
                    n_runs=int(_roll_runs),
                    percentile=int(_roll_pct),
                    input_dir=str(INPUT_DIR),
                    output_dir=str(OUTPUT_DIR),
                    demand_file=str(INPUT_DIR / "Total_demand.xlsx"),
                    tr_file=str(INPUT_DIR / "TR TOTALV21E - Copia.xlsx"),
                    exclusions_file=str(INPUT_DIR / "esclusioni.xlsx"),
                    sku_catalog_path=str(INPUT_DIR / "tutte_righe_univoche_V2.xlsx"),
                    prezzi_path=str(INPUT_FILES["Prezzi"]),
                    keep_per_offset_files=bool(_roll_keep_files),
                    max_iterations=int(_roll_max_iter) if _roll_max_iter else None,
                    aff_model=_roll_aff_model,
                    aff_optional=_roll_aff_optional,
                )
                # run_rolling_simulation restituisce (df_combined, df_valore)
                if isinstance(_roll_result, tuple):
                    df_rolling, df_rolling_valore = _roll_result
                else:
                    df_rolling, df_rolling_valore = _roll_result, pd.DataFrame()
                _roll_progress.progress(90)
                st.session_state["df_rolling"] = df_rolling
                st.session_state["df_rolling_valore"] = df_rolling_valore
                _rlog(f"Completato in {time.time()-_t_roll:.1f}s")
                _rlog(f"Righe totali: {len(df_rolling)}")
                _rlog(f"Iterazioni: {df_rolling['mese_lancio'].nunique() if 'mese_lancio' in df_rolling.columns else 'N/A'}")
                _roll_progress.progress(100)
                _roll_status.text("Rolling completata!")
                st.success("✅ Rolling simulation completata — risultati sotto.")
            except Exception as _re:
                _rlog(f"ERRORE: {_re}")
                import traceback as _tb
                _rlog(_tb.format_exc())
                st.error(f"Errore rolling: {_re}")
            finally:
                st.session_state["_roll_running"] = False

        # ── Visualizza risultati rolling ───────────────────────────────────
        df_rolling_view  = st.session_state.get("df_rolling", None)
        df_rolling_valore_view = st.session_state.get("df_rolling_valore", None)

        # Carica da file se non in sessione
        _roll_out_file   = OUTPUT_DIR / "sku_safety_stock_rolling.xlsx"
        _roll_val_file   = OUTPUT_DIR / "sku_safety_stock_rolling_valore.xlsx"
        if df_rolling_view is None and _roll_out_file.exists():
            try:
                df_rolling_view = pd.read_excel(str(_roll_out_file))
                st.info(f"Risultati caricati da file: `{_roll_out_file.name}`")
            except Exception:
                pass
        if df_rolling_valore_view is None and _roll_val_file.exists():
            try:
                df_rolling_valore_view = pd.read_excel(str(_roll_val_file))
            except Exception:
                pass

        if df_rolling_view is not None and len(df_rolling_view) > 0:
            st.markdown("### Risultati Rolling")

            _roll_ss_col = next(
                (c for c in df_rolling_view.columns if "safety_stock" in c.lower() and "total" in c.lower()),
                None,
            )

            # ── VISTA VALORE — grafico linea per mese lancio ──────────────
            if df_rolling_valore_view is not None and len(df_rolling_valore_view) > 0:
                st.markdown("#### 💶 Safety Stock Totale per Mese Lancio")

                _val_cols = [c for c in df_rolling_valore_view.columns if c != "mese_lancio" and c != "offset"]
                _rv1, _rv2 = st.columns(2)

                # Grafico valore €
                if "ss_valore_eur" in df_rolling_valore_view.columns:
                    _fig_val = px.line(
                        df_rolling_valore_view,
                        x="mese_lancio",
                        y="ss_valore_eur",
                        markers=True,
                        title="Safety Stock totale (€) per mese di lancio",
                        labels={"mese_lancio": "Mese Lancio", "ss_valore_eur": "Safety Stock (€)"},
                    )
                    _fig_val.update_traces(line_color="#E3000F", marker_size=8)
                    st.plotly_chart(_fig_val, use_container_width=True)

                # Grafico unità
                _ss_u_col = next((c for c in df_rolling_valore_view.columns if "ss_unita" in c.lower()), None)
                _fig_cols = st.columns(2)
                if _ss_u_col:
                    with _fig_cols[0]:
                        _fig_u = px.bar(
                            df_rolling_valore_view,
                            x="mese_lancio",
                            y=_ss_u_col,
                            title="Safety Stock totale (unità) per mese lancio",
                            labels={"mese_lancio": "Mese Lancio", _ss_u_col: "Unità SS"},
                            color_discrete_sequence=["#1a6fa8"],
                        )
                        st.plotly_chart(_fig_u, use_container_width=True)
                if "n_sku" in df_rolling_valore_view.columns:
                    with _fig_cols[1]:
                        _fig_n = px.bar(
                            df_rolling_valore_view,
                            x="mese_lancio",
                            y="n_sku",
                            title="N° SKU per mese lancio",
                            labels={"mese_lancio": "Mese Lancio", "n_sku": "SKU con match"},
                            color_discrete_sequence=["#666"],
                        )
                        st.plotly_chart(_fig_n, use_container_width=True)

                # Tabella valore sintetica
                with st.expander("📋 Tabella riepilogo valore", expanded=True):
                    st.dataframe(df_rolling_valore_view, use_container_width=True)
                    _buf_val = io.BytesIO()
                    df_rolling_valore_view.to_excel(_buf_val, index=False)
                    st.download_button(
                        "⬇ Scarica riepilogo valore",
                        data=_buf_val.getvalue(),
                        file_name="rolling_riepilogo_valore.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_rolling_val",
                    )

            st.markdown("---")

            # ── VISTA PER SKU ─────────────────────────────────────────────
            st.markdown("#### 📦 Dettaglio per SKU")

            _mesi_disponibili = (
                list(df_rolling_view["mese_lancio"].unique())
                if "mese_lancio" in df_rolling_view.columns else []
            )
            _sel_mese = st.multiselect(
                "Filtra per mese lancio",
                options=_mesi_disponibili,
                default=_mesi_disponibili,
                key="roll_filter_mese",
            )
            _df_rv_f = df_rolling_view[df_rolling_view["mese_lancio"].isin(_sel_mese)] if _sel_mese else df_rolling_view

            # Metriche
            _rm1, _rm2, _rm3 = st.columns(3)
            _rm1.metric("Iterazioni", len(_mesi_disponibili))
            _rm2.metric("SKU univoci", df_rolling_view["SKU"].nunique() if "SKU" in df_rolling_view.columns else "N/A")
            if "prezzo_safety" in _df_rv_f.columns:
                _rm3.metric("SS economica totale (filtro)", f"€ {_df_rv_f['prezzo_safety'].sum(skipna=True):,.0f}")

            # Grafico top SKU
            if _roll_ss_col and "SKU" in _df_rv_f.columns:
                _top_sku = (
                    _df_rv_f.groupby("SKU")[_roll_ss_col].mean()
                    .nlargest(20).index.tolist()
                )
                _df_chart = _df_rv_f[_df_rv_f["SKU"].isin(_top_sku)]
                if len(_df_chart) > 0:
                    _fig_roll = px.bar(
                        _df_chart,
                        x="mese_lancio",
                        y=_roll_ss_col,
                        color="SKU",
                        title="Top 20 SKU — Safety Stock per Mese Lancio",
                        labels={"mese_lancio": "Mese Lancio", _roll_ss_col: "Safety Stock (u.)"},
                        barmode="group",
                    )
                    st.plotly_chart(_fig_roll, use_container_width=True)

            with st.expander("📋 Dati tabellari rolling (per SKU)", expanded=False):
                st.dataframe(_df_rv_f, use_container_width=True)
                _buf_roll = io.BytesIO()
                _df_rv_f.to_excel(_buf_roll, index=False)
                st.download_button(
                    "⬇ Scarica Excel (filtrato)",
                    data=_buf_roll.getvalue(),
                    file_name="sku_safety_stock_rolling_filtrato.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_rolling",
                )
        else:
            st.info("Nessun risultato rolling disponibile. Avvia la simulazione sopra.")


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


# ============================================================================
# MODULO — STOCK ALIGNMENT
# ============================================================================

elif _active_module == "sa":

    tab_run, tab_view = st.tabs(["Esegui Analisi", "Visualizza Risultati"])

    # ########################################################################
    # SA — TAB 1: ESEGUI ANALISI
    # ########################################################################

    with tab_run:

        st.markdown("## Stock Alignment — Mix-Driven Excess")
        st.caption(
            "Confronta la domanda media per SKU sotto due scenari TR (actual vs forecast) "
            "su tutti gli SKU BOM, senza filtro residual. "
            "Quantifica l'eccesso/carenza di scorte causata da errori nel Take Rate previsto."
        )

        # ── Preparazione Catalogo SA ─────────────────────────────────────────
        _sa_catalog_path = INPUT_DIR / "tutte_righe_univoche_SA.xlsx"

        with st.expander("Passo 0 — Preparazione Catalogo SA", expanded=not _sa_catalog_path.exists()):
            if _sa_catalog_path.exists():
                st.success(
                    f"Catalogo SA presente ({format_size(_sa_catalog_path)}). "
                    "Rigeneralo se hai aggiornato i file BOM."
                )
            else:
                st.warning(
                    "Catalogo SA non trovato. Generalo prima di avviare Stock Alignment. "
                    "Richiede i file BOM150, DATI_AGGIUNTI e ZIBP_BOM_125 nelle cartelle V21E/V22E/V24T."
                )

            _prep_btn = st.button(
                "GENERA CATALOGO SA",
                type="primary" if not _sa_catalog_path.exists() else "secondary",
                key="btn_sa_prep",
            )
            if _prep_btn:
                _prep_log = st.empty()
                _prep_logs = []

                def _plog(msg):
                    _prep_logs.append(msg)
                    _prep_log.markdown(
                        f'<div class="log-console">{"<br>".join(_prep_logs[-20:])}</div>',
                        unsafe_allow_html=True,
                    )

                try:
                    if str(BASE_DIR) not in sys.path:
                        sys.path.insert(0, str(BASE_DIR))
                    import importlib
                    import crea_catalogo_sa as _csa
                    importlib.reload(_csa)

                    _plog("Avvio preparazione catalogo SA...")
                    _df_cat = _csa.main(
                        output_catalog=str(_sa_catalog_path),
                        filter_untraced=True,
                    )
                    _plog(f"Catalogo SA generato: {len(_df_cat)} righe, "
                          f"{_df_cat['Numero componenti'].nunique()} SKU univoci")
                    st.success(f"Catalogo SA pronto — {len(_df_cat)} righe")
                    st.rerun()
                except Exception as _e:
                    import traceback
                    _plog(f"ERRORE: {_e}")
                    st.error(f"Errore preparazione catalogo: {_e}")
                    st.code(traceback.format_exc(), language="python")

        # File TR disponibili
        _sa_xlsx = sorted([
            f.name for f in INPUT_DIR.glob("*.xlsx")
            if f.name.lower() != "prezzi.xlsx"
        ]) if INPUT_DIR.exists() else []

        sa_c1, sa_c2 = st.columns(2)

        with sa_c1:
            st.markdown("**TR Actual** — situazione corrente")
            sa_tr_actual_choice = st.selectbox(
                "File TR Actual",
                options=_sa_xlsx,
                index=next((i for i, f in enumerate(_sa_xlsx)
                            if "actual" in f.lower()), 0) if _sa_xlsx else 0,
                key="sa_tr_actual",
            )
            sa_tr_actual_path = INPUT_DIR / sa_tr_actual_choice if sa_tr_actual_choice else None
            if sa_tr_actual_path and sa_tr_actual_path.exists():
                st.caption(f"✅ {format_size(sa_tr_actual_path)}")
            elif sa_tr_actual_choice:
                st.caption("❌ File non trovato")

        with sa_c2:
            st.markdown("**TR Forecast** — scenario futuro")
            sa_tr_forecast_choice = st.selectbox(
                "File TR Forecast",
                options=_sa_xlsx,
                index=next((i for i, f in enumerate(_sa_xlsx)
                            if "forecast" in f.lower()),
                           min(1, len(_sa_xlsx) - 1)) if len(_sa_xlsx) > 1 else 0,
                key="sa_tr_forecast",
            )
            sa_tr_forecast_path = INPUT_DIR / sa_tr_forecast_choice if sa_tr_forecast_choice else None
            if sa_tr_forecast_path and sa_tr_forecast_path.exists():
                st.caption(f"✅ {format_size(sa_tr_forecast_path)}")
            elif sa_tr_forecast_choice:
                st.caption("❌ File non trovato")

        if (sa_tr_actual_choice and sa_tr_forecast_choice
                and sa_tr_actual_choice == sa_tr_forecast_choice):
            st.warning("⚠️ I due file TR selezionati sono identici — il confronto non avrà senso.")

        # Parametri
        sa_p1, sa_p2, sa_p3 = st.columns(3)
        with sa_p1:
            sa_n_months = st.number_input(
                "N. mesi domanda",
                min_value=1, max_value=24, value=3, step=1,
                key="sa_n_months",
                help="Numero di mesi su cui aggregare la domanda simulata.",
            )
        with sa_p2:
            sa_n_runs = st.number_input(
                "Run Monte Carlo",
                min_value=5, max_value=5000, value=200, step=5,
                key="sa_n_runs",
                help="Numero di run MC per scenario.",
            )
        with sa_p3:
            sa_offset = st.number_input(
                "Offset mese iniziale",
                min_value=0, max_value=11, value=0, step=1,
                key="sa_offset",
                help="Quanti mesi saltare nel forecast (0 = primo mese disponibile).",
            )

        # Prezzi opzionale
        _sa_prezzi_path = INPUT_FILES["Prezzi"]
        if _sa_prezzi_path.exists():
            st.success(f"`Prezzi.xlsx` trovato ({format_size(_sa_prezzi_path)}) — calcolo valore excess/shortage abilitato")
        else:
            st.info("`Prezzi.xlsx` non trovato — confronto solo su volumi (no valore €)")

        # Prerequisiti
        SA_REQUIRED = ["Unique Rows"]
        _sa_missing = [n for n in SA_REQUIRED if not file_exists(INPUT_FILES[n])]
        if not (sa_tr_actual_path and sa_tr_actual_path.exists()):
            _sa_missing.append("TR Actual")
        if not (sa_tr_forecast_path and sa_tr_forecast_path.exists()):
            _sa_missing.append("TR Forecast")

        if _sa_missing:
            st.error(f"⚠️ File mancanti: {', '.join(_sa_missing)}")
        else:
            st.success("✅ Tutti i file necessari per Stock Alignment sono presenti.")

        sa_btn = st.button(
            "AVVIA STOCK ALIGNMENT",
            type="primary",
            disabled=bool(_sa_missing),
            key="btn_sa",
        )

        if sa_btn:
            if st.session_state.get("_sa_running", False):
                st.warning("Stock Alignment già in esecuzione — attendi il completamento.")
                st.stop()
            st.session_state["_sa_running"] = True

            sa_log_container = st.empty()
            sa_progress      = st.progress(0)
            sa_status        = st.empty()
            sa_logs          = []

            def sa_log(msg: str):
                sa_logs.append(msg)
                sa_log_container.markdown(
                    f'<div class="log-console">{"<br>".join(sa_logs[-30:])}</div>',
                    unsafe_allow_html=True,
                )

            sa_start = time.time()
            try:
                if str(BASE_DIR) not in sys.path:
                    sys.path.insert(0, str(BASE_DIR))

                from stock_alignment_V2 import main as run_stock_alignment

                sa_log("=" * 60)
                sa_log("STOCK ALIGNMENT — MIX-DRIVEN EXCESS")
                sa_log(f"Avvio: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                sa_log(f"TR Actual:   {sa_tr_actual_choice}")
                sa_log(f"TR Forecast: {sa_tr_forecast_choice}")
                sa_log(f"N. mesi: {sa_n_months}  |  Run MC: {sa_n_runs}  |  Offset: {sa_offset}")
                sa_log("=" * 60)

                sa_status.text("Simulazione in corso (può richiedere diversi minuti)...")
                sa_progress.progress(10)

                _sa_output_path = str(OUTPUT_FILES["Stock Alignment"])
                df_alignment = run_stock_alignment(
                    tr_actual_path=str(sa_tr_actual_path),
                    tr_forecast_path=str(sa_tr_forecast_path),
                    n_months=int(sa_n_months),
                    n_runs=int(sa_n_runs),
                    start_month_offset=int(sa_offset),
                    output_path=_sa_output_path,
                )

                sa_progress.progress(95)
                sa_elapsed = time.time() - sa_start
                sa_log("\n" + "=" * 60)
                sa_log(f"COMPLETATO in {sa_elapsed:.1f}s ({sa_elapsed/60:.1f} min)")

                if df_alignment is not None:
                    _n_excess   = int((df_alignment["excess"] > 0).sum()) if "excess" in df_alignment.columns else 0
                    _n_shortage = int((df_alignment["shortage"] > 0).sum()) if "shortage" in df_alignment.columns else 0
                    sa_log(f"SKU totali: {len(df_alignment)}  |  Excess: {_n_excess}  |  Shortage: {_n_shortage}")
                    if "excess_value" in df_alignment.columns:
                        _tot_exc = df_alignment["excess_value"].fillna(0).sum()
                        _tot_sho = df_alignment["shortage_value"].fillna(0).sum() if "shortage_value" in df_alignment.columns else 0
                        sa_log(f"Valore excess: EUR {_tot_exc:,.0f}  |  Shortage: EUR {_tot_sho:,.0f}")

                sa_log("=" * 60)
                sa_progress.progress(100)
                sa_status.text("Completato!")

                st.session_state["df_alignment"]      = df_alignment
                st.session_state["sa_actual_label"]   = sa_tr_actual_choice
                st.session_state["sa_forecast_label"] = sa_tr_forecast_choice

                st.success(
                    f"✅ Stock Alignment completato in {sa_elapsed:.1f}s — "
                    "vai a **Visualizza Risultati** per l'analisi."
                )
                st.session_state["last_ram"] = get_memory_mb()

                # Archivia output appena generato
                _arch_sa = archive_run_sa(
                    actual_label=sa_tr_actual_choice,
                    forecast_label=sa_tr_forecast_choice,
                    n_months=int(sa_n_months),
                    n_runs=int(sa_n_runs),
                )
                if _arch_sa:
                    sa_log(f"Archiviato → history_sa/{_arch_sa.name}")
                    st.toast(f"Salvato in storico SA: {_arch_sa.name}")

            except Exception as e:
                import traceback
                sa_progress.progress(100)
                sa_status.text("Errore.")
                sa_log(f"\n✗ ERRORE: {e}")
                st.error(f"Errore durante Stock Alignment: {e}")
                st.code(traceback.format_exc(), language="python")

            finally:
                st.session_state["_sa_running"] = False


    # ########################################################################
    # SA — TAB 2: VISUALIZZA RISULTATI
    # ########################################################################

    with tab_view:

        st.subheader("Risultati — Stock Alignment")

        _df_sa = st.session_state.get("df_alignment")

        if _df_sa is None:
            _sa_path = OUTPUT_FILES["Stock Alignment"]
            if _sa_path.exists():
                try:
                    _df_sa = pd.read_excel(_sa_path, sheet_name="Risultati")
                except Exception as _e:
                    st.error(f"Errore lettura Stock Alignment: {_e}")

        if _df_sa is None:
            st.info(
                "Nessun risultato Stock Alignment disponibile. "
                "Esegui prima l'analisi nella scheda **Esegui Analisi**."
            )
        else:
            _sa_act_lbl  = st.session_state.get("sa_actual_label",   "Actual")
            _sa_fore_lbl = st.session_state.get("sa_forecast_label", "Forecast")

            st.caption(
                f"Mix-Driven Excess: **{_sa_act_lbl}** vs **{_sa_fore_lbl}**. "
                "Δ = Forecast − Actual. Excess = domanda forecast > actual (scorte in eccesso)."
            )

            # ── Debug: mostra composizione versioni ──────────────────────────
            with st.expander("🔍 Debug — Composizione dati", expanded=False):
                if "Versioni_SKU" in _df_sa.columns:
                    _all_v_counts = _df_sa["Versioni_SKU"].value_counts().head(10)
                    st.write("**Top 10 valori Versioni_SKU:**")
                    st.write(_all_v_counts)

                    # Mostra quanti SKU contengono V21E
                    _v21e_count = _df_sa["Versioni_SKU"].astype(str).str.contains("V21E", case=False, na=False).sum()
                    _v24t_count = _df_sa["Versioni_SKU"].astype(str).str.contains("V24T", case=False, na=False).sum()
                    _v22e_count = _df_sa["Versioni_SKU"].astype(str).str.contains("V22E", case=False, na=False).sum()
                    st.metric("Contengono V21E", _v21e_count)
                    st.metric("Contengono V24T", _v24t_count)
                    st.metric("Contengono V22E", _v22e_count)

            # ── Filtri interattivi ──────────────────────────────────────────
            sa_f1, sa_f2, sa_f3, sa_f4, sa_f5 = st.columns(5)
            _df_sa_show = _df_sa.copy()

            with sa_f1:
                _sa_search = st.text_input("Filtra SKU / Descrizione", key="sa_search")
                if _sa_search:
                    _mask = _df_sa_show.apply(
                        lambda col: col.astype(str).str.contains(_sa_search, case=False, na=False)
                    ).any(axis=1)
                    _df_sa_show = _df_sa_show[_mask]

            with sa_f2:
                if "Supermodel" in _df_sa_show.columns:
                    _all_super = sorted(set(
                        v.strip() for vs in _df_sa_show["Supermodel"].dropna()
                        for v in str(vs).split(",") if v.strip()
                    ))
                    _sel_super = st.multiselect(
                        "Supermodel", _all_super, key="sa_filter_super",
                        help="MSV4, MSV4S, MSV4PP, Rally, RS, etc."
                    )
                    if _sel_super:
                        _df_sa_show = _df_sa_show[
                            _df_sa_show["Supermodel"].apply(
                                lambda x: any(s in str(x) for s in _sel_super) if pd.notna(x) else False
                            )
                        ]

            with sa_f3:
                if "Versioni_SKU" in _df_sa_show.columns:
                    _all_vers = sorted(set(
                        v for vs in _df_sa_show["Versioni_SKU"].dropna()
                        for v in str(vs).split(",")
                    ))
                    _sel_vers = st.multiselect("Versione BOM", _all_vers, key="sa_filter_vers")
                    if _sel_vers:
                        _df_sa_show = _df_sa_show[
                            _df_sa_show["Versioni_SKU"].apply(
                                lambda x: any(v in str(x) for v in _sel_vers)
                            )
                        ]

            with sa_f4:
                if "Livello_Min" in _df_sa_show.columns and _df_sa_show["Livello_Min"].notna().any():
                    _all_lev = sorted(_df_sa_show["Livello_Min"].dropna().astype(int).unique())
                    _sel_lev = st.multiselect("Livello BOM", _all_lev, key="sa_filter_livello")
                    if _sel_lev:
                        _df_sa_show = _df_sa_show[
                            _df_sa_show["Livello_Min"].isna() |
                            _df_sa_show["Livello_Min"].astype(int).isin(_sel_lev)
                        ]

            with sa_f5:
                if "Tipo_Approvv" in _df_sa_show.columns:
                    _all_tipo = sorted(_df_sa_show["Tipo_Approvv"].dropna().unique())
                    _sel_tipo = st.multiselect("Tipo Approvv.", _all_tipo, key="sa_filter_tipo")
                    if _sel_tipo:
                        _df_sa_show = _df_sa_show[_df_sa_show["Tipo_Approvv"].isin(_sel_tipo)]

            # Filtro delta
            _sa_delta_filter = st.radio(
                "Mostra",
                ["Tutti", "Solo Excess (Delta > 0)", "Solo Shortage (Delta < 0)", "Solo Allineati (Delta = 0)"],
                horizontal=True,
                key="sa_delta_filter",
            )
            if _sa_delta_filter == "Solo Excess (Delta > 0)" and "delta_demand" in _df_sa_show.columns:
                _df_sa_show = _df_sa_show[_df_sa_show["delta_demand"] > 0]
            elif _sa_delta_filter == "Solo Shortage (Delta < 0)" and "delta_demand" in _df_sa_show.columns:
                _df_sa_show = _df_sa_show[_df_sa_show["delta_demand"] < 0]
            elif _sa_delta_filter == "Solo Allineati (Delta = 0)" and "delta_demand" in _df_sa_show.columns:
                _df_sa_show = _df_sa_show[_df_sa_show["delta_demand"] == 0]

            st.dataframe(_df_sa_show, use_container_width=True, hide_index=True)

            # ── KPI ─────────────────────────────────────────────────────────
            # IMPORTANTE: usare _df_sa_show (filtrato) non _df_sa (intero)
            if "delta_demand" in _df_sa_show.columns:
                _sa_m1, _sa_m2, _sa_m3, _sa_m4 = st.columns(4)
                _sa_m1.metric("SKU totali", len(_df_sa_show))
                _sa_m2.metric("Excess (Delta > 0)", int((_df_sa_show["delta_demand"] > 0).sum()))
                _sa_m3.metric("Shortage (Delta < 0)", int((_df_sa_show["delta_demand"] < 0).sum()))
                _sa_m4.metric("Allineati (Delta = 0)", int((_df_sa_show["delta_demand"] == 0).sum()))

            if "excess_value" in _df_sa_show.columns and _df_sa_show["excess_value"].notna().any():
                _sa_v1, _sa_v2, _sa_v3 = st.columns(3)
                _tot_exc_val = _df_sa_show["excess_value"].fillna(0).sum()
                _tot_sho_val = _df_sa_show["shortage_value"].fillna(0).sum() if "shortage_value" in _df_sa_show.columns else 0
                _sa_v1.metric("Valore Excess totale", f"EUR {_tot_exc_val:,.0f}")
                _sa_v2.metric("Valore Shortage totale", f"EUR {_tot_sho_val:,.0f}")
                _sa_v3.metric("Delta netto", f"EUR {_tot_exc_val - _tot_sho_val:+,.0f}")

                # ── Suddivisione per Supermodel ──────────────────────────────
                st.markdown("### Suddivisione per Supermodel")
                if "Supermodel" in _df_sa_show.columns:
                    # Esplodi Supermodel separando per virgola
                    _df_expl = _df_sa_show.copy()
                    _df_expl["Supermodel"] = _df_expl["Supermodel"].fillna("Unknown")
                    _df_expl = _df_expl.assign(
                        Supermodel=_df_expl["Supermodel"].str.split(",")
                    ).explode("Supermodel")
                    _df_expl["Supermodel"] = _df_expl["Supermodel"].str.strip()

                    _super_agg = (
                        _df_expl.groupby("Supermodel").agg({
                            "SKU": "count",
                            "excess_value": "sum",
                            "shortage_value": "sum" if "shortage_value" in _df_expl.columns else lambda x: 0,
                        }).rename(columns={"SKU": "SKU_count"})
                    )
                    _super_agg["Delta"] = _super_agg["excess_value"] - _super_agg["shortage_value"]

                    st.dataframe(
                        _super_agg.style.format({
                            "excess_value": "EUR {:,.0f}",
                            "shortage_value": "EUR {:,.0f}",
                            "Delta": "EUR {:+,.0f}",
                        }),
                        use_container_width=True
                    )
                else:
                    st.info("Colonna Supermodel non disponibile nel catalogo")

            # ── Grafico top-30 per |Delta| ───────────────────────────────────
            if "delta_demand" in _df_sa_show.columns and len(_df_sa_show) > 0:
                _sa_sku_col = "SKU" if "SKU" in _df_sa_show.columns else _df_sa_show.columns[0]
                _sa_top30 = (
                    _df_sa_show[[_sa_sku_col, "delta_demand"]]
                    .assign(_abs=lambda d: d["delta_demand"].abs())
                    .sort_values("_abs", ascending=False)
                    .drop(columns="_abs")
                    .head(30)
                )
                _fig_sa = px.bar(
                    _sa_top30,
                    x=_sa_sku_col,
                    y="delta_demand",
                    color="delta_demand",
                    color_continuous_scale="RdYlGn_r",
                    title="Top-30 SKU per |Delta Domanda| (Forecast - Actual)",
                    labels={"delta_demand": "Delta Domanda (pz)", _sa_sku_col: "SKU"},
                )
                _fig_sa.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(_fig_sa, use_container_width=True)

            # ── Download ────────────────────────────────────────────────────
            _sa_out_path = OUTPUT_FILES["Stock Alignment"]
            if _sa_out_path.exists():
                with open(_sa_out_path, "rb") as _fh:
                    st.download_button(
                        label="Scarica stock_alignment_V2.xlsx",
                        data=_fh.read(),
                        file_name="stock_alignment_V2.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="sa_download",
                    )

            # ════════════════════════════════════════════════════════════════
            # SEZIONE — CONFRONTO VERSIONI STORICHE SA
            # ════════════════════════════════════════════════════════════════
            st.markdown("---")
            st.markdown("## Confronto Versioni Storiche")
            st.caption(
                "Seleziona due versioni archiviate e confronta excess, shortage "
                "e delta domanda SKU per SKU. Utile per valutare l'impatto di "
                "diversi scenari TR (actual/forecast) o periodi di analisi."
            )

            _sa_hist_dir = OUTPUT_DIR / "history_sa"
            _sa_hist_files = sorted(_sa_hist_dir.glob("*.xlsx"), reverse=True) \
                if _sa_hist_dir.exists() else []

            _sa_curr = OUTPUT_FILES["Stock Alignment"]
            _sa_ver_options = []
            if _sa_curr.exists():
                _sa_ver_options.append("Versione attuale")
            _sa_ver_options += [f.name for f in _sa_hist_files]

            if len(_sa_ver_options) < 2:
                st.info(
                    "Servono almeno due versioni per il confronto.  \n"
                    "Esegui Stock Alignment almeno due volte (anche con parametri diversi)."
                )
            else:
                _sa_cmp_c1, _sa_cmp_c2 = st.columns(2)
                with _sa_cmp_c1:
                    _sa_ver_a = st.selectbox("Versione A (base)", _sa_ver_options,
                                             index=0, key="sa_ver_cmp_a")
                with _sa_cmp_c2:
                    _sa_ver_b = st.selectbox("Versione B (confronto)", _sa_ver_options,
                                             index=min(1, len(_sa_ver_options) - 1),
                                             key="sa_ver_cmp_b")

                if st.button("Esegui confronto", key="btn_sa_ver_compare"):
                    def _load_sa_ver(label):
                        if label == "Versione attuale":
                            return pd.read_excel(str(_sa_curr), sheet_name="Risultati")
                        else:
                            return pd.read_excel(str(_sa_hist_dir / label),
                                                 sheet_name="Risultati")
                    try:
                        _sa_dfa = _load_sa_ver(_sa_ver_a)
                        _sa_dfb = _load_sa_ver(_sa_ver_b)

                        if "SKU" not in _sa_dfa.columns or "SKU" not in _sa_dfb.columns:
                            st.error("Colonna 'SKU' non trovata in uno dei file.")
                        else:
                            # Colonne da portare nel merge
                            _sa_cols_a = ["SKU"]
                            _sa_cols_b = ["SKU"]
                            for _c in ["Descrizione_BOM", "Versioni_SKU"]:
                                if _c in _sa_dfa.columns:
                                    _sa_cols_a.append(_c)
                            for _c in ["excess_value", "shortage_value", "delta_demand"]:
                                if _c in _sa_dfa.columns:
                                    _sa_cols_a.append(_c)
                                if _c in _sa_dfb.columns:
                                    _sa_cols_b.append(_c)

                            _sa_merged = (
                                _sa_dfa[_sa_cols_a]
                                .merge(
                                    _sa_dfb[_sa_cols_b],
                                    on="SKU", how="outer",
                                    suffixes=("_A", "_B"),
                                )
                                .fillna(0)
                            )

                            # Δ delta_demand
                            if "delta_demand_A" in _sa_merged.columns and "delta_demand_B" in _sa_merged.columns:
                                _sa_merged["Δ delta_demand"] = (
                                    _sa_merged["delta_demand_A"] - _sa_merged["delta_demand_B"]
                                ).round(2)

                            # Δ excess_value
                            if "excess_value_A" in _sa_merged.columns and "excess_value_B" in _sa_merged.columns:
                                _sa_merged["Δ excess_value (€)"] = (
                                    _sa_merged["excess_value_A"] - _sa_merged["excess_value_B"]
                                ).round(2)

                            # Δ shortage_value
                            if "shortage_value_A" in _sa_merged.columns and "shortage_value_B" in _sa_merged.columns:
                                _sa_merged["Δ shortage_value (€)"] = (
                                    _sa_merged["shortage_value_A"] - _sa_merged["shortage_value_B"]
                                ).round(2)

                            # Rinomina colonne per chiarezza
                            _sa_merged = _sa_merged.rename(columns={
                                "excess_value_A":   f"Excess € — {_sa_ver_a[:20]}",
                                "excess_value_B":   f"Excess € — {_sa_ver_b[:20]}",
                                "shortage_value_A": f"Shortage € — {_sa_ver_a[:20]}",
                                "shortage_value_B": f"Shortage € — {_sa_ver_b[:20]}",
                                "delta_demand_A":   f"Delta pz — {_sa_ver_a[:20]}",
                                "delta_demand_B":   f"Delta pz — {_sa_ver_b[:20]}",
                            })

                            # Ordina per |Δ excess_value| decrescente
                            if "Δ excess_value (€)" in _sa_merged.columns:
                                _sa_merged = _sa_merged.sort_values(
                                    "Δ excess_value (€)", key=abs, ascending=False
                                )

                            st.session_state["_sa_ver_compare_df"]  = _sa_merged
                            st.session_state["_sa_ver_compare_lbla"] = _sa_ver_a
                            st.session_state["_sa_ver_compare_lblb"] = _sa_ver_b
                    except Exception as _sa_vce:
                        st.error(f"Errore nel confronto: {_sa_vce}")

                # Mostra risultato confronto
                if "_sa_ver_compare_df" in st.session_state:
                    _sa_vdf  = st.session_state["_sa_ver_compare_df"]
                    _sa_lbl_a = st.session_state.get("_sa_ver_compare_lbla", "A")
                    _sa_lbl_b = st.session_state.get("_sa_ver_compare_lblb", "B")

                    st.markdown(f"**Confronto:** `{_sa_lbl_a}` vs `{_sa_lbl_b}`")

                    # KPI economici di sintesi
                    _sa_exc_a_col = f"Excess € — {_sa_lbl_a[:20]}"
                    _sa_exc_b_col = f"Excess € — {_sa_lbl_b[:20]}"
                    if _sa_exc_a_col in _sa_vdf.columns and _sa_exc_b_col in _sa_vdf.columns:
                        _sa_tot_a = _sa_vdf[_sa_exc_a_col].sum()
                        _sa_tot_b = _sa_vdf[_sa_exc_b_col].sum()
                        _sa_kt1, _sa_kt2, _sa_kt3 = st.columns(3)
                        _sa_kt1.metric("Excess totale — A", f"EUR {_sa_tot_a:,.0f}")
                        _sa_kt2.metric("Excess totale — B", f"EUR {_sa_tot_b:,.0f}")
                        _sa_kt3.metric("Δ Excess (A − B)", f"EUR {_sa_tot_a - _sa_tot_b:+,.0f}",
                                       delta_color="inverse" if (_sa_tot_a - _sa_tot_b) > 0 else "normal")

                    _sa_kc1, _sa_kc2, _sa_kc3 = st.columns(3)
                    _sa_kc1.metric("SKU totali", len(_sa_vdf))
                    if "Δ excess_value (€)" in _sa_vdf.columns:
                        _sa_kc2.metric("A > B (excess aumentato)",
                                       int((_sa_vdf["Δ excess_value (€)"] > 0).sum()))
                        _sa_kc3.metric("A < B (excess ridotto)",
                                       int((_sa_vdf["Δ excess_value (€)"] < 0).sum()))

                    st.dataframe(_sa_vdf, use_container_width=True, hide_index=True)

                    # Grafico top-20 per |Δ excess_value|
                    if "Δ excess_value (€)" in _sa_vdf.columns:
                        _sa_top20 = (
                            _sa_vdf[["SKU", "Δ excess_value (€)"]]
                            .assign(_abs=lambda d: d["Δ excess_value (€)"].abs())
                            .sort_values("_abs", ascending=False)
                            .drop(columns="_abs")
                            .head(20)
                        )
                        _fig_sa_cmp = go.Figure(go.Bar(
                            x=_sa_top20["SKU"].astype(str),
                            y=_sa_top20["Δ excess_value (€)"],
                            marker_color=[
                                "#E3000F" if v > 0 else "#1a6fa8"
                                for v in _sa_top20["Δ excess_value (€)"]
                            ],
                            text=_sa_top20["Δ excess_value (€)"].round(0),
                            textposition="outside",
                        ))
                        _fig_sa_cmp.update_layout(
                            title=f"Top-20 SKU per |Δ Excess €| — {_sa_lbl_a[:25]} vs {_sa_lbl_b[:25]}",
                            xaxis_title="SKU",
                            yaxis_title="Δ Excess Value (€)",
                            xaxis_tickangle=-45,
                            height=420,
                        )
                        st.plotly_chart(_fig_sa_cmp, use_container_width=True)


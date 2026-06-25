# -*- coding: utf-8 -*-
"""Frontend Streamlit — Safety Stock (run singolo).

Adattato da SolMidTerm-ALL/app_V3.py, ridotto al SOLO calcolo Safety Stock con
run singolo, sopra la pipeline modulare PROD (mid_term.SafetyStockPipeline).
Niente rolling / stock alignment / cost allocation / sensitivity / TR implied.

Avvio:
    cd "...\\SolMidTerm - Prod\\mid_term"
    streamlit run app.py
"""
from __future__ import annotations

import io
import json
import logging
import os
import re
import shutil
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Palette classi ABC (stessa di SolMidTerm-ALL/app_V3.py)
ABC_COLORS = {"A": "#E3000F", "B": "#ff8c00", "C": "#1a6fa8", "N/D": "#666"}

# --- Path: il package usa path hardcoded .\Input -> CWD deve essere mid_term/ ---
PKG = Path(__file__).resolve().parent
REPO = PKG.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))
os.chdir(PKG)

st.set_page_config(
    page_title="Ducati Safety Stock Simulator",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Import del package DOPO il chdir (la discovery modelli legge .\Input\MODEL a import-time)
from mid_term_supermodels import SafetyStockPipeline, PipelineConfig, SimulationConfig
from mid_term_supermodels.results import BOMData
from mid_term_supermodels.montecarlo import load_monthly_forecast
try:
    from mid_term_supermodels.tr import _AFFIDABILITA_MAP
    AFF_LEVELS = sorted(_AFFIDABILITA_MAP, key=_AFFIDABILITA_MAP.get)
except Exception:
    AFF_LEVELS = ["NULLA", "BASSA", "MEDIA", "ALTA"]

CONFIG = PipelineConfig(input_dir=PKG / "Input", output_dir=PKG / "Output")
RUN_CONFIG_PATH = PKG / "run_config.json"


# ============================================================================
# Helper
# ============================================================================
def _load_run_config_defaults() -> dict:
    """Default parametri da run_config.json (fallback a valori sensati)."""
    base = {
        "n_runs": 100, "random_seed": 42, "start_month": None,
        "percentile_safety_stock": 99, "affidabilita_model": "MEDIA",
        "affidabilita_optional": "MEDIA", "gate_max_mesi": 4,
    }
    try:
        base.update(json.loads(RUN_CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        pass
    return base


@st.cache_data(show_spinner=False)
def _forecast_months() -> list:
    """Nomi-mese dal forecast (per la scelta del mese di partenza)."""
    try:
        _, months = load_monthly_forecast(str(CONFIG.resolve_input(CONFIG.forecast_file)))
        return list(months)
    except Exception as e:
        st.warning(f"Forecast non leggibile: {e}")
        return []


def _ss_total_col(df: pd.DataFrame, pct: int) -> str | None:
    c = f"safety_stock_p{pct}_total"
    return c if c in df.columns else next(
        (x for x in df.columns if "safety_stock" in x.lower() and "total" in x.lower()
         and "int" not in x.lower()), None)


def _run_pipeline(sim: SimulationConfig, rigenera_bom: bool):
    """Esegue gli step 0..4 (run singolo). Se rigenera_bom=False riusa il
    catalogo tutte_righe_univoche_V2.xlsx esistente. Cattura i log mid_term."""
    buf = io.StringIO()
    handler = logging.StreamHandler(buf)
    handler.setFormatter(logging.Formatter("%(name)s: %(message)s"))
    root = logging.getLogger("mid_term")
    prev_level = root.level
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    try:
        pipe = SafetyStockPipeline(CONFIG, sim, validate_config=False)
        tr = pipe.run_step_0_load_tr()
        if rigenera_bom:
            bom = pipe.run_step_1_build_bom()
        else:
            cat_path = CONFIG.resolve_input(CONFIG.catalog_filename)
            if not cat_path.exists():
                raise FileNotFoundError(
                    f"Catalogo non trovato ({cat_path}). Attiva 'Rigenera BOM' al primo run.")
            catalog = pd.read_excel(cat_path)
            bom = BOMData(catalog=catalog, matrices={}, quality_report=pd.DataFrame())
        sim_res = pipe.run_step_2_montecarlo(tr)
        mat = pipe.run_step_3_matching(sim_res, bom)
        ss = pipe.run_step_4_safety_stock(mat, bom)
        return ss, mat, buf.getvalue()
    finally:
        root.removeHandler(handler)
        root.setLevel(prev_level)


def _read_bytes(path) -> bytes | None:
    try:
        return Path(path).read_bytes()
    except Exception:
        return None


# Path dei file di output (scritti dalla pipeline, letti dal tab Risultati come app_V3)
OUT_PER_SKU  = CONFIG.resolve_output(CONFIG.safety_stock_per_sku_filename)
OUT_PER_MESE = CONFIG.resolve_output(CONFIG.safety_stock_per_mese_filename)
OUT_NO_LT    = CONFIG.resolve_output(CONFIG.sku_senza_lt_filename)
HISTORY_DIR  = CONFIG.output_dir / "history"


def archive_run(aff_model: str, aff_optional: str, n_runs: int = 0,
                percentile: int = 99, month: str = "") -> "Path | None":
    """Archivia il PER_SKU corrente in Output/history/ con nome descrittivo
    GG-MM-AAAA_<mese>_Model-Opt_NNNrun_PXX.xlsx (come app_V3). Sovrascrive se esiste."""
    if not OUT_PER_SKU.exists():
        return None
    HISTORY_DIR.mkdir(exist_ok=True)
    date_str = time.strftime("%d-%m-%Y")
    month_str = f"_{re.sub(r'[^A-Za-z0-9]+', '', str(month))}" if month else ""
    fname = f"{date_str}{month_str}_{aff_model}-{aff_optional}_{int(n_runs)}run_P{int(percentile)}.xlsx"
    dst = HISTORY_DIR / fname
    shutil.copy2(OUT_PER_SKU, dst)
    return dst


def _list_history() -> list:
    """File storici in Output/history/ ordinati per data modifica (recente prima)."""
    if not HISTORY_DIR.exists():
        return []
    return sorted(HISTORY_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)


def _compute_abc(df: pd.DataFrame, value_col: str) -> pd.Series:
    """Classe ABC via Pareto cumulato su value_col: <=80%->A, <=95%->B, resto C.
    Stessa logica di app_V3. Ritorna Series allineata a df.index (default 'C')."""
    abc = pd.Series("C", index=df.index)
    s = pd.to_numeric(df[value_col], errors="coerce")
    valid = s.dropna()
    tot = valid.sum()
    if tot > 0:
        order = valid.sort_values(ascending=False)
        cum = order.cumsum() / tot
        cls = np.select([cum <= 0.80, cum <= 0.95], ["A", "B"], default="C")
        abc.loc[order.index] = cls
    return abc


@st.cache_data(show_spinner=False)
def _read_excel_cached(path: str, mtime: float) -> pd.DataFrame:
    """Legge un Excel; la cache si invalida quando cambia mtime (= nuovo run)."""
    return pd.read_excel(path)


def _load_output(path):
    """Legge un file di output da disco se esiste, altrimenti None."""
    p = Path(path)
    if not p.exists():
        return None
    try:
        return _read_excel_cached(str(p), p.stat().st_mtime)
    except Exception as e:
        st.warning(f"Impossibile leggere {p.name}: {e}")
        return None


def _detect_percentile(cols) -> int:
    """Ricava il percentile dal nome colonna safety_stock_pXX."""
    for c in cols:
        m = re.match(r"safety_stock_p(\d+)$", str(c).strip())
        if m:
            return int(m.group(1))
    return 99


# ============================================================================
# CSS PERSONALIZZATO (stesso tema Ducati di SolMidTerm-ALL/app_V3.py)
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

    /* Titoli sezione uniformi rosso Ducati */
    [data-testid="stMarkdownContainer"] h2 {
        color: #E3000F !important;
        border-bottom: 2px solid #E3000F;
        padding-bottom: 0.25rem;
        margin-top: 1.2rem;
    }
    [data-testid="stMarkdownContainer"] h3 { color: #E3000F !important; }

    /* Tab principali piu' grandi e prominenti */
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
    .stTabs [aria-selected="false"] { color: #aaa !important; }
    .stTabs [data-baseweb="tab-highlight"] { display: none !important; }
    .stTabs [data-baseweb="tab-border"] { display: none !important; }

    /* Card parametri pipeline */
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
# HEADER (logo + titolo gradient, come app_V3)
# ============================================================================
_hc_logo, _hc_title = st.columns([1, 4])
with _hc_logo:
    _logo_path = PKG / "Immagine1.png"
    if _logo_path.exists():
        st.image(str(_logo_path), use_container_width=True)
with _hc_title:
    st.markdown("""
<div class="main-header">
    <h1>Ducati Safety Stock Simulator</h1>
    <p>Monte Carlo Pipeline · Multistrada V4 · Run Singolo</p>
</div>
""", unsafe_allow_html=True)

defaults = _load_run_config_defaults()
months = _forecast_months()

tab_run, tab_res = st.tabs(["▶️  Esegui Analisi", "📊  Risultati"])

# ============================================================================
# Tab Esegui
# ============================================================================
with tab_run:
    st.markdown("## Parametri")
    _pc1, _pc2, _pc3 = st.columns(3)
    with _pc1:
        n_runs = st.number_input("Numero scenari (n_runs)", min_value=1, max_value=5000,
                                 value=int(defaults["n_runs"]), step=50)
        percentile = st.slider("Percentile safety stock", 1, 99,
                               int(defaults["percentile_safety_stock"]))
    with _pc2:
        if months:
            _def_m = defaults.get("start_month") if defaults.get("start_month") in months else months[0]
            start_month = st.selectbox("Mese di partenza", months, index=months.index(_def_m))
        else:
            start_month = None
            st.error("Nessun mese disponibile nel forecast.")
        gate_max = st.number_input("Gate max (mesi)", min_value=1, max_value=60,
                                   value=int(defaults["gate_max_mesi"]),
                                   help="Cap massimo sul gate di OGNI SKU (gate maggiori sostituiti da questo valore) "
                                        "+ default per le SKU SOLO_MODELLO senza gate. "
                                        "Residual = (FROZEN+Freq) - gate; senza dato supplier la SKU non viene calcolata.")
    with _pc3:
        _amm = defaults["affidabilita_model"] if defaults["affidabilita_model"] in AFF_LEVELS else AFF_LEVELS[-1]
        _amo = defaults["affidabilita_optional"] if defaults["affidabilita_optional"] in AFF_LEVELS else AFF_LEVELS[0]
        aff_model = st.selectbox("Affidabilita' Model", AFF_LEVELS, index=AFF_LEVELS.index(_amm))
        aff_optional = st.selectbox("Affidabilita' Optional", AFF_LEVELS, index=AFF_LEVELS.index(_amo))

    rigenera_bom = st.checkbox("Rigenera BOM (Step 1)", value=True,
                               help="Se disattivo riusa tutte_righe_univoche_V2.xlsx esistente (run piu' veloce).")

    st.markdown("## Esecuzione")
    if st.button("Esegui calcolo Safety Stock", type="primary", disabled=(start_month is None)):
        sim = SimulationConfig(
            n_runs=int(n_runs),
            random_seed=int(defaults.get("random_seed", 42)),
            start_month_offset=months.index(start_month),
            percentile_safety_stock=int(percentile),
            affidabilita_model=aff_model,
            affidabilita_optional=aff_optional,
            gate_max_mesi=int(gate_max),
        )
        try:
            with st.spinner("Esecuzione pipeline (Monte Carlo + matching + safety)..."):
                ss, mat, logtext = _run_pipeline(sim, rigenera_bom)
            st.session_state["ss_per_sku"] = ss.per_sku
            st.session_state["no_lt"] = mat.no_lt
            st.session_state["percentile"] = int(percentile)
            st.session_state["out_paths"] = ss.output_paths
            st.session_state["log"] = logtext
            # Auto-archiviazione in Output/history/ (come app_V3)
            try:
                _arch = archive_run(aff_model, aff_optional, int(n_runs),
                                    int(percentile), start_month)
                if _arch is not None:
                    st.toast(f"Salvato in storico: {_arch.name}")
            except Exception as _ae:
                st.warning(f"Archiviazione storico non riuscita: {_ae}")
            st.success(f"Completato: {len(ss.per_sku)} SKU. Vai al tab Risultati.")
        except Exception as e:
            st.error(f"Errore durante il calcolo: {e}")
            st.exception(e)

    if "log" in st.session_state:
        with st.expander("Log esecuzione", expanded=False):
            _logtxt = (st.session_state["log"] or "(nessun log)")
            _logtxt = _logtxt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            st.markdown(f'<div class="log-console">{_logtxt}</div>', unsafe_allow_html=True)

# ============================================================================
# Tab Risultati
# ============================================================================
with tab_res:
    # ── Selettore versione (attuale o storicizzata in Output/history/) ──────────
    _ver_opts = (["Versione attuale"] if OUT_PER_SKU.exists() else [])
    _hist_files = _list_history()
    _ver_opts += [f.name for f in _hist_files]

    if not _ver_opts:
        st.info("Nessun risultato in Output/. Esegui un calcolo dal tab 'Esegui Analisi'.")
    else:
        _sel = st.selectbox("Versione da analizzare", _ver_opts, index=0, key="res_ver",
                            help="Versione attuale o una storicizzata (Output/history/). "
                                 "Ogni run viene archiviato automaticamente.")
        _is_current = (_sel == "Versione attuale")
        df = _load_output(OUT_PER_SKU if _is_current else HISTORY_DIR / _sel)

        if df is None:
            st.info("Versione non leggibile.")
        else:
            pct = _detect_percentile(df.columns)
            ss_tot_col = _ss_total_col(df, pct)
            _has_prezzo = "prezzo_safety" in df.columns
            no_lt = _load_output(OUT_NO_LT) if _is_current else None
            st.caption(f"Versione: **{_sel}**"
                       + ("" if _is_current else "  ·  (storico)"))

            # Classe ABC (Pareto cumulato su prezzo_safety)
            df = df.copy()
            df["ABC"] = _compute_abc(df, "prezzo_safety") if _has_prezzo else "N/D"

            # ── KPI ─────────────────────────────────────────────────────────────
            st.markdown("## Riepilogo")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("SKU", f"{len(df):,}")
            units = pd.to_numeric(df[ss_tot_col], errors="coerce").sum() if ss_tot_col else 0
            c2.metric("Unita' safety stock", f"{units:,.0f}")
            valore = pd.to_numeric(df["prezzo_safety"], errors="coerce").sum() if _has_prezzo else 0
            c3.metric("Valore safety (EUR)", f"{valore:,.0f}")
            n_solo = int(df["SOLO_MODELLO"].sum()) if "SOLO_MODELLO" in df else 0
            c4.metric("SKU SOLO_MODELLO", f"{n_solo:,}")
            c5.metric("SKU senza LT", f"{len(no_lt):,}" if isinstance(no_lt, pd.DataFrame) else "—")

            # ── Visione Strategica: Top 10 + Torta ABC ──────────────────────────
            st.markdown("## Visione Strategica")
            _has_desc = "Description" in df.columns
            sc1, sc2 = st.columns(2)

            with sc1:
                st.markdown("#### Top 10 SKU per Costo Safety Stock")
                if _has_prezzo:
                    _top = (df.dropna(subset=["prezzo_safety"])
                            .sort_values("prezzo_safety", ascending=False).head(10).copy())
                    if not _top.empty:
                        _top["Label"] = (_top["SKU"].astype(str)
                                         + (("  " + _top["Description"].astype(str).str[:28]) if _has_desc else ""))
                        fig_top = go.Figure(go.Bar(
                            x=_top["prezzo_safety"], y=_top["Label"], orientation="h",
                            marker_color=[ABC_COLORS.get(str(a), "#666") for a in _top["ABC"]],
                            text=[f"€ {v:,.0f}" for v in _top["prezzo_safety"]],
                            textposition="auto",
                        ))
                        fig_top.update_layout(
                            height=360, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                            font_color="#eee", margin=dict(t=10, b=10, l=10, r=10),
                            yaxis=dict(autorange="reversed"), xaxis_title="Costo SS (€)",
                        )
                        st.plotly_chart(fig_top, use_container_width=True)
                        st.caption("Colore = classe ABC (rosso=A, arancione=B, blu=C).")
                    else:
                        st.info("Nessun dato prezzo disponibile.")
                else:
                    st.info("Top 10 per costo non disponibile: manca `prezzo_safety` (prezzi).")

            with sc2:
                st.markdown("#### Torta di Pareto — Classe ABC")
                if _has_prezzo:
                    _pie = df.groupby("ABC")["prezzo_safety"].sum().reset_index()
                    fig_pie = px.pie(_pie, names="ABC", values="prezzo_safety",
                                     color="ABC", color_discrete_map=ABC_COLORS, hole=0.45)
                    fig_pie.update_traces(textinfo="label+percent")
                    fig_pie.update_layout(paper_bgcolor="#0d1117", font_color="#eee",
                                          margin=dict(t=10, b=0, l=0, r=0), height=300,
                                          showlegend=False)
                    st.plotly_chart(fig_pie, use_container_width=True)
                    _abc_sum = (df.groupby("ABC")
                                .agg(N_SKU=("SKU", "count"), Costo_SS=("prezzo_safety", "sum"))
                                .reset_index())
                    _abc_sum["Incidenza %"] = (_abc_sum["Costo_SS"] / _abc_sum["Costo_SS"].sum() * 100).round(1)
                    _abc_sum["Costo_SS"] = _abc_sum["Costo_SS"].map("€ {:,.0f}".format)
                    st.dataframe(_abc_sum, use_container_width=True, hide_index=True, height=160)
                else:
                    st.info("Torta ABC non disponibile: manca `prezzo_safety`.")

            # ── Curva di Pareto ─────────────────────────────────────────────────
            st.markdown("#### Curva di Pareto — Concentrazione del Capitale")
            if _has_prezzo:
                _par = (df[["SKU", "prezzo_safety", "ABC"] + (["Description"] if _has_desc else [])]
                        .dropna(subset=["prezzo_safety"])
                        .sort_values("prezzo_safety", ascending=False).reset_index(drop=True).copy())
                if not _par.empty and _par["prezzo_safety"].sum() > 0:
                    _tot = _par["prezzo_safety"].sum()
                    _par["_cum"] = _par["prezzo_safety"].cumsum() / _tot * 100
                    _par["_idx"] = range(1, len(_par) + 1)
                    _n80 = int((_par["_cum"] <= 80).sum()) + 1
                    fig_par = go.Figure()
                    fig_par.add_trace(go.Bar(
                        x=_par["_idx"], y=_par["prezzo_safety"],
                        marker_color=[ABC_COLORS.get(str(a), "#666") for a in _par["ABC"]],
                        name="Costo SS (€)", yaxis="y1",
                    ))
                    fig_par.add_trace(go.Scatter(
                        x=_par["_idx"], y=_par["_cum"], mode="lines",
                        name="Cumulata %", line=dict(color="#ff9800", width=2.5), yaxis="y2",
                    ))
                    fig_par.add_shape(type="line", xref="paper", x0=0, x1=1,
                                      y0=80, y1=80, yref="y2",
                                      line=dict(color="#aaa", dash="dash", width=1.5))
                    fig_par.update_layout(
                        height=360, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                        font_color="#eee",
                        xaxis=dict(title="SKU (ordinati per costo decrescente)",
                                   showticklabels=False, gridcolor="#222"),
                        yaxis=dict(title="Costo Safety Stock (€)", gridcolor="#222"),
                        yaxis2=dict(title=dict(text="Cumulata %", font=dict(color="#ff9800")),
                                    overlaying="y", side="right", range=[0, 105],
                                    ticksuffix="%", tickfont=dict(color="#ff9800"), showgrid=False),
                        legend=dict(orientation="h", y=1.06), margin=dict(t=30, b=30), bargap=0.02,
                    )
                    st.plotly_chart(fig_par, use_container_width=True)
                    st.caption(f"**{_n80} SKU** ({_n80/len(_par)*100:.0f}%) concentrano l'80% del capitale "
                               f"(€ {_tot*0.8:,.0f} su € {_tot:,.0f}).")
                else:
                    st.info("Pareto non disponibile: nessun dato prezzo.")
            else:
                st.info("Pareto non disponibile: manca `prezzo_safety` (esegui con prezzi).")

            # ── Elenco completo SKU ─────────────────────────────────────────────
            st.markdown("## Elenco completo SKU")
            _df_show = df.drop(
                columns=[c for c in df.columns
                         if df[c].map(lambda v: isinstance(v, (dict, list))).any()],
                errors="ignore",
            )
            st.dataframe(_df_show, use_container_width=True, height=400)

            if isinstance(no_lt, pd.DataFrame) and len(no_lt):
                with st.expander(f"SKU senza lead time ({len(no_lt)})"):
                    st.dataframe(no_lt, use_container_width=True, height=240)

            # ── Confronto Versioni (diff a 2 versioni storiche) ─────────────────
            st.markdown("## Confronto Versioni")
            st.caption("Seleziona due versioni archiviate e confronta la Safety Stock SKU per SKU "
                       "(Δ quantità e Δ valore €). Utile per valutare l'impatto di parametri/run diversi.")
            _cmp_opts = (["Versione attuale"] if OUT_PER_SKU.exists() else []) + [f.name for f in _hist_files]
            if len(_cmp_opts) < 2:
                st.info("Servono almeno due versioni per il confronto. "
                        "Esegui la pipeline almeno due volte (anche con parametri diversi).")
            else:
                _ca, _cb = st.columns(2)
                with _ca:
                    _va = st.selectbox("Versione A (base)", _cmp_opts, index=0, key="cmp_a")
                with _cb:
                    _vb = st.selectbox("Versione B (confronto)", _cmp_opts,
                                       index=min(1, len(_cmp_opts) - 1), key="cmp_b")
                if st.button("Esegui confronto", key="btn_cmp"):
                    def _ld_ver(lbl):
                        return _load_output(OUT_PER_SKU if lbl == "Versione attuale" else HISTORY_DIR / lbl)
                    _da, _db = _ld_ver(_va), _ld_ver(_vb)
                    if _da is None or _db is None or "SKU" not in _da.columns:
                        st.error("Versioni non leggibili o senza colonna SKU.")
                    else:
                        _sa = _ss_total_col(_da, _detect_percentile(_da.columns))
                        _sb = _ss_total_col(_db, _detect_percentile(_db.columns))
                        if not _sa or not _sb:
                            st.error("Colonna safety_stock_p*_total non trovata in una delle versioni.")
                        else:
                            _cols_a = ["SKU", _sa] + (["Description"] if "Description" in _da.columns else []) \
                                      + (["prezzo_safety"] if "prezzo_safety" in _da.columns else [])
                            _cols_b = ["SKU", _sb] + (["prezzo_safety"] if "prezzo_safety" in _db.columns else [])
                            _m = (_da[_cols_a].merge(
                                    _db[_cols_b].rename(columns={_sb: "SS_B", "prezzo_safety": "prezzo_B"}),
                                    on="SKU", how="outer").fillna(0))
                            _m["Δ SS"] = (_m[_sa] - _m["SS_B"]).round(2)
                            _m["Δ SS %"] = np.where(_m["SS_B"] != 0,
                                                    (_m["Δ SS"] / _m["SS_B"] * 100).round(1), np.nan)
                            if "prezzo_safety" in _m.columns and "prezzo_B" in _m.columns:
                                _m["Δ €"] = (_m["prezzo_safety"] - _m["prezzo_B"]).round(2)
                            _m = _m.sort_values("Δ SS", ascending=False, key=abs)
                            st.session_state["_cmp_df"] = _m
                            st.session_state["_cmp_la"] = _va
                            st.session_state["_cmp_lb"] = _vb

                if "_cmp_df" in st.session_state:
                    _m = st.session_state["_cmp_df"]
                    _la = st.session_state.get("_cmp_la", "A")
                    _lb = st.session_state.get("_cmp_lb", "B")
                    _has_desc_c = "Description" in _m.columns
                    st.markdown(f"**Confronto:** `{_la}` vs `{_lb}`")
                    if "prezzo_safety" in _m.columns and "prezzo_B" in _m.columns:
                        _ta, _tb = _m["prezzo_safety"].sum(), _m["prezzo_B"].sum()
                        e1, e2, e3 = st.columns(3)
                        e1.metric("Totale € A", f"{_ta:,.0f}")
                        e2.metric("Totale € B", f"{_tb:,.0f}")
                        e3.metric("Δ € (A−B)", f"{_ta - _tb:+,.0f}")
                    q1, q2, q3, q4 = st.columns(4)
                    q1.metric("SKU totali", len(_m))
                    q2.metric("A > B (SS aumentata)", int((_m["Δ SS"] > 0).sum()))
                    q3.metric("A < B (SS ridotta)", int((_m["Δ SS"] < 0).sum()))
                    q4.metric("Invariati", int((_m["Δ SS"] == 0).sum()))
                    _td = _m.reindex(_m["Δ SS"].abs().nlargest(20).index).copy()
                    _td["Label"] = (_td["SKU"].astype(str)
                                    + ((" " + _td["Description"].astype(str).str[:25]) if _has_desc_c else ""))
                    fig_d = go.Figure(go.Bar(
                        x=_td["Δ SS"], y=_td["Label"], orientation="h",
                        marker_color=["#E3000F" if v > 0 else "#1a6fa8" for v in _td["Δ SS"]],
                        text=_td["Δ SS"].round(1), textposition="outside",
                    ))
                    fig_d.update_layout(title="Top 20 SKU per |Δ Safety Stock| (A−B)",
                                        height=500, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                                        font_color="#eee", xaxis_title="Δ Safety Stock",
                                        margin=dict(l=10, r=80, t=40, b=10))
                    st.plotly_chart(fig_d, use_container_width=True)
                    if "Δ €" in _m.columns:
                        _te = _m.reindex(_m["Δ €"].abs().nlargest(20).index).copy()
                        _te["Label"] = (_te["SKU"].astype(str)
                                        + ((" " + _te["Description"].astype(str).str[:25]) if _has_desc_c else ""))
                        fig_e = go.Figure(go.Bar(
                            x=_te["Δ €"], y=_te["Label"], orientation="h",
                            marker_color=["#e67e00" if v > 0 else "#2ecc71" for v in _te["Δ €"]],
                            text=_te["Δ €"].apply(lambda v: f"{v:+,.0f} €"), textposition="outside",
                        ))
                        fig_e.update_layout(title="Top 20 SKU per |Δ Valore €| (A−B)",
                                            height=500, plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                                            font_color="#eee", xaxis_title="Δ € Safety Stock",
                                            margin=dict(l=10, r=100, t=40, b=10))
                        st.plotly_chart(fig_e, use_container_width=True)
                    with st.expander("Tabella completa SKU per SKU"):
                        st.dataframe(_m.reset_index(drop=True), use_container_width=True,
                                     height=420, hide_index=True)
                    _cbuf = io.BytesIO()
                    with pd.ExcelWriter(_cbuf, engine="openpyxl") as _cw:
                        _m.to_excel(_cw, index=False, sheet_name="Confronto Versioni")
                    _cbuf.seek(0)
                    st.download_button("Scarica confronto (Excel)", data=_cbuf,
                                       file_name=f"confronto_{_la[:18]}_vs_{_lb[:18]}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="dl_cmp")

            # ── Download ────────────────────────────────────────────────────────
            st.markdown("## Download")
            d1, d2, d3 = st.columns(3)
            for col, path, label in [
                (d1, OUT_PER_SKU, "PER_SKU.xlsx"),
                (d2, OUT_PER_MESE, "PER_MESE.xlsx"),
                (d3, OUT_NO_LT, "SKU_senza_LT.xlsx"),
            ]:
                b = _read_bytes(path) if Path(path).exists() else None
                if b:
                    col.download_button(f"Scarica {label}", data=b, file_name=Path(path).name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    col.caption(f"{label}: n/d")

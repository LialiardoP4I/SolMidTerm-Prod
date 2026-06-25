# -*- coding: utf-8 -*-
"""Frontend Streamlit — Safety Stock MULTI-SUPERMODEL (ROLLING + pooling).

Dedicato alla modalità multi-supermodel rolling con risk pooling sui componenti
condivisi e quantità per-configurazione. Distinto da app.py (run singolo).

Avvio:
    cd "...\\SolMidTerm - Prod\\mid_term_supermodels"
    streamlit run app_supermodels.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd
import streamlit as st

# --- Path: il package usa path hardcoded .\Input -> CWD deve essere il package ---
PKG = Path(__file__).resolve().parent
REPO = PKG.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))
os.chdir(PKG)

st.set_page_config(
    page_title="Ducati Safety Stock — Multi-Supermodel (Rolling)",
    page_icon="🏍️",
    layout="wide",
)

# Import del package DOPO il chdir
from mid_term_supermodels import multi_supermodel as ms
from mid_term_supermodels.matching import parse_month_to_sortable
try:
    from mid_term_supermodels.tr import _AFFIDABILITA_MAP
    AFF_LEVELS = list(_AFFIDABILITA_MAP.keys())
except Exception:
    AFF_LEVELS = ["ALTA", "MEDIA", "BASSA"]

INPUT_DIR = PKG / "Input"
OUTPUT_DIR = PKG / "Output"
RUN_CONFIG_PATH = PKG / "run_config.json"
POOLED_NAME = "sku_safety_stock_POOLED_rolling_supermodels.xlsx"


# ----------------------------------------------------------------------------
# Helper
# ----------------------------------------------------------------------------
def _default_config() -> dict:
    base = {"n_runs": 100, "random_seed": 42, "start_month": "MAY 2026",
            "percentile_safety_stock": 99, "affidabilita_model": "ALTA",
            "affidabilita_optional": "BASSA", "gate_max_mesi": 4,
            "multi_supermodel": True, "rolling": True}
    try:
        base.update(json.loads(RUN_CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        pass
    return base


@st.cache_data(show_spinner=False)
def _discover() -> list:
    try:
        dirs = ms.discover_supermodels(str(INPUT_DIR))
        return [Path(d).name for d in dirs]
    except Exception:
        return []


def _pooled_cols(df):
    """Trova i nomi colonna SS pooled/standalone (il percentile è dinamico)."""
    pooled = next((c for c in df.columns if c.startswith("safety_stock_p") and c.endswith("_pooled")), None)
    standalone = [c for c in df.columns if "_standalone_" in str(c) and c.startswith("safety_stock_p")]
    sm_mean = [c for c in df.columns if c.startswith("mean_demand_") and c != "mean_demand_pooled"]
    return pooled, standalone, sm_mean


def _load_pooled(path: Path):
    if not path.exists():
        return None
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"Impossibile leggere il pooled: {e}")
        return None


def _month_order(df):
    if "mese_lancio" not in df.columns:
        return []
    return sorted(df["mese_lancio"].dropna().unique().tolist(), key=parse_month_to_sortable)


# ----------------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------------
_logo = next((p for p in [PKG / "logo.png", REPO / "logo.png"] if p.exists()), None)
_hc = st.columns([1, 5]) if _logo else None
if _logo:
    with _hc[0]:
        st.image(str(_logo), use_container_width=True)
    _title_ctx = _hc[1]
else:
    _title_ctx = st.container()
with _title_ctx:
    st.title("🏍️ Safety Stock — Multi-Supermodel (Rolling)")
    st.caption("Pooling con risk pooling sui componenti condivisi · quantità per-configurazione · "
               "safety stock per mese di lancio (rolling).")

tab_run, tab_res = st.tabs(["▶️  Esegui Analisi", "📊  Risultati"])

# ----------------------------------------------------------------------------
# TAB ESEGUI
# ----------------------------------------------------------------------------
with tab_run:
    cfg = _default_config()
    supermodels = _discover()

    st.markdown("### 1 · Supermodel da includere")
    if not supermodels:
        st.warning(f"Nessun supermodel trovato in {INPUT_DIR}. "
                   "Serve `Input/<SuperModel>/MODEL/` + `Total_demand.xlsx` + `TR*.xlsx`.")
        selected = []
    else:
        cols = st.columns(min(4, len(supermodels)))
        selected = []
        for i, sm in enumerate(supermodels):
            with cols[i % len(cols)]:
                if st.checkbox(sm, value=True, key=f"sm_{sm}"):
                    selected.append(sm)
        st.caption(f"Selezionati: {len(selected)}/{len(supermodels)} · "
                   "i componenti presenti in più supermodel vengono poolati (risk pooling).")

    st.markdown("### 2 · Parametri")
    p1, p2, p3 = st.columns(3)
    with p1:
        n_runs = st.number_input("Scenari Monte Carlo (n_runs)", min_value=1, max_value=5000,
                                 value=int(cfg.get("n_runs", 100)), step=10,
                                 help="Più scenari = stima più stabile ma più lento.")
        gate_max = st.number_input("Gate max (mesi)", min_value=1, max_value=24,
                                   value=int(cfg.get("gate_max_mesi", 4)), step=1)
    with p2:
        percentile = st.slider("Percentile safety stock", 1, 99,
                               int(cfg.get("percentile_safety_stock", 99)))
        aff_model = st.selectbox("Affidabilità modello (L0)", AFF_LEVELS,
                                 index=AFF_LEVELS.index(cfg.get("affidabilita_model", AFF_LEVELS[0]))
                                 if cfg.get("affidabilita_model") in AFF_LEVELS else 0)
    with p3:
        aff_optional = st.selectbox("Affidabilità optional (L1)", AFF_LEVELS,
                                    index=AFF_LEVELS.index(cfg.get("affidabilita_optional", AFF_LEVELS[-1]))
                                    if cfg.get("affidabilita_optional") in AFF_LEVELS else len(AFF_LEVELS) - 1)

    st.info("⏱️ Il rolling multi-supermodel è intensivo (MC su tutti i mesi di lancio × supermodel). "
            "Per una prova rapida usa n_runs basso (es. 10).")

    if st.button("▶️  Esegui rolling multi-supermodel", type="primary",
                 disabled=not selected):
        run_cfg = dict(cfg)
        run_cfg.update({"n_runs": int(n_runs), "percentile_safety_stock": int(percentile),
                        "affidabilita_model": aff_model, "affidabilita_optional": aff_optional,
                        "gate_max_mesi": int(gate_max), "multi_supermodel": True, "rolling": True,
                        "only_supermodels": selected})
        # JSON temporaneo con i parametri scelti (non sovrascrive run_config.json).
        tmp = Path(tempfile.gettempdir()) / "_app_run_config_supermodels.json"
        tmp.write_text(json.dumps(run_cfg, indent=2), encoding="utf-8")
        # IMPORTANTE: il run gira in SUBPROCESS. L'orchestratore fa os.chdir nelle
        # cartelle di staging temporanee: eseguirlo in-process cambierebbe la CWD
        # del processo Streamlit e lo romperebbe (non ritroverebbe lo script).
        env = dict(os.environ)
        env["MIDTERM_RUN_CONFIG"] = str(tmp)
        t0 = time.time()
        try:
            with st.spinner(f"Esecuzione su {len(selected)} supermodel — può richiedere parecchi minuti…"):
                proc = subprocess.run(
                    [sys.executable, str(PKG / "run_pipeline.py")],
                    cwd=str(PKG), env=env, capture_output=True, text=True, timeout=7200,
                )
            if proc.returncode == 0:
                st.success(f"✅ Completato in {time.time() - t0:.0f}s. "
                           f"Output: {OUTPUT_DIR / POOLED_NAME}")
                st.cache_data.clear()
                with st.expander("Log esecuzione"):
                    st.code((proc.stdout or proc.stderr or "")[-6000:])
            else:
                st.error("❌ Errore durante l'esecuzione (vedi log).")
                with st.expander("Log errore", expanded=True):
                    st.code((proc.stderr or proc.stdout or "")[-6000:])
        except subprocess.TimeoutExpired:
            st.error("❌ Timeout (>2h). Riduci n_runs o il numero di supermodel.")
        except Exception as e:
            st.error(f"❌ Errore avvio run: {e}")

# ----------------------------------------------------------------------------
# TAB RISULTATI
# ----------------------------------------------------------------------------
with tab_res:
    pooled_path = OUTPUT_DIR / POOLED_NAME
    df = _load_pooled(pooled_path)
    if df is None or len(df) == 0:
        st.info("Nessun risultato disponibile. Esegui un'analisi nel tab «Esegui» "
                f"o assicurati che esista `{pooled_path}`.")
    else:
        ss_pooled_col, standalone_cols, sm_mean_cols = _pooled_cols(df)
        st.caption(f"Sorgente: {pooled_path.name} · ultimo aggiornamento "
                   f"{time.strftime('%Y-%m-%d %H:%M', time.localtime(pooled_path.stat().st_mtime))}")

        # --- KPI ---
        n_sku = df["SKU"].nunique() if "SKU" in df else 0
        n_mesi = df["mese_lancio"].nunique() if "mese_lancio" in df else 0
        n_shared = 0
        if sm_mean_cols:
            shared_mask = (df[sm_mean_cols] > 0).sum(axis=1) > 1
            n_shared = int(df.loc[shared_mask, "SKU"].nunique())
        saving_tot = float(df["risk_pooling_saving"].clip(lower=0).sum()) if "risk_pooling_saving" in df else 0.0
        ss_tot = float(df[ss_pooled_col].sum()) if ss_pooled_col else 0.0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Componenti (SKU)", f"{n_sku:,}")
        k2.metric("Mesi di lancio", n_mesi)
        k3.metric("Componenti condivisi", f"{n_shared:,}")
        k4.metric("Risparmio risk-pooling", f"{saving_tot:,.0f}")
        k5.metric("SS pooled totale", f"{ss_tot:,.0f}")

        st.divider()

        # --- Filtri ---
        f1, f2 = st.columns([1, 2])
        months = _month_order(df)
        with f1:
            mese_sel = st.selectbox("Mese di lancio", ["(tutti)"] + months)
        with f2:
            query = st.text_input("Cerca SKU o descrizione", "")

        view = df.copy()
        if mese_sel != "(tutti)":
            view = view[view["mese_lancio"] == mese_sel]
        if query.strip():
            q = query.strip().lower()
            mask = df["SKU"].astype(str).str.lower().str.contains(q)
            if "Description" in df:
                mask = mask | df["Description"].astype(str).str.lower().str.contains(q)
            view = view[mask.reindex(view.index, fill_value=False)]

        # flag finestra troncata evidenziato
        st.markdown(f"### Tabella ({len(view):,} righe)")
        if "finestra_troncata" in view.columns:
            n_tronc = int(view["finestra_troncata"].astype(bool).sum())
            if n_tronc:
                st.caption(f"⚠️ {n_tronc} righe con finestra troncata (mesi di fine orizzonte, "
                           "safety stock sottostimata).")
        st.dataframe(view, use_container_width=True, height=420)

        st.divider()

        # --- Risk pooling ---
        st.markdown("### 🔗 Risk pooling (componenti condivisi)")
        if "risk_pooling_saving" in df and ss_pooled_col:
            shared = view[view["risk_pooling_saving"] > 0].copy()
            if len(shared):
                cols_show = (["mese_lancio", "SKU", "Description", "mean_demand_pooled",
                              ss_pooled_col, "risk_pooling_saving"] + standalone_cols)
                cols_show = [c for c in cols_show if c in shared.columns]
                shared = shared.sort_values("risk_pooling_saving", ascending=False)
                st.caption("SS pooled < Σ SS standalone: il risparmio nasce dal poolare la domanda "
                           "indipendente dei supermodel sui componenti condivisi.")
                st.dataframe(shared[cols_show].head(200), use_container_width=True, height=300)
            else:
                st.info("Nessun componente condiviso nella vista corrente (risk_pooling_saving = 0).")

        st.divider()

        # --- Grafici ---
        st.markdown("### 📈 Grafici")
        g1, g2 = st.columns(2)
        with g1:
            if ss_pooled_col and "mese_lancio" in df:
                by_month = (df.groupby("mese_lancio")[ss_pooled_col].sum()
                            .reindex(months).rename("SS pooled totale"))
                st.caption("Safety stock pooled totale per mese di lancio")
                st.line_chart(by_month)
        with g2:
            if ss_pooled_col:
                top = (view.groupby("SKU")[ss_pooled_col].max()
                       .sort_values(ascending=False).head(15).rename("SS pooled (max)"))
                st.caption("Top 15 componenti per SS pooled")
                st.bar_chart(top)
        if "risk_pooling_saving" in df:
            topsav = (view.groupby("SKU")["risk_pooling_saving"].max()
                      .sort_values(ascending=False).head(15).rename("Risparmio pooling"))
            topsav = topsav[topsav > 0]
            if len(topsav):
                st.caption("Top 15 componenti per risparmio da risk pooling")
                st.bar_chart(topsav)

        st.divider()

        # --- Download ---
        with open(pooled_path, "rb") as fh:
            st.download_button("⬇️  Scarica pooled (Excel)", fh.read(),
                               file_name=POOLED_NAME,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

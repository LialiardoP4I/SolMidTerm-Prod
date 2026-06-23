# -*- coding: utf-8 -*-
"""Orchestratore della pipeline Safety Stock: esegue in ordine i 5 passi.

In parole semplici: prende i dati (domanda prevista, distinte base, regole),
simula molte vendite possibili e calcola quanta scorta di sicurezza tenere per
ogni componente.

I 5 passi (step):
  0) Carica i dati TR, il forecast di domanda e le esclusioni.
  1) Costruisce la distinta base (BOM) consolidata.
  2) Simula la domanda mese per mese (Monte Carlo).
  3) Abbina ogni componente agli scenari simulati.
  4) Calcola la safety stock per componente/mese e la valorizza in euro.

La classe SafetyStockPipeline espone run() (tutto) e run_step_N() (singolo passo).
"""
import time
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from mid_term_supermodels._logging import get_logger
from mid_term_supermodels.config import PipelineConfig, SimulationConfig, shared_logistica_dir
from mid_term_supermodels.results import (TRData, BOMData, SimulationResult,
                               MatchResult, SafetyStockResult)
from mid_term_supermodels.tr import load_monthly_tr, parse_tr_file
from mid_term_supermodels.bom import (
    process_combination,
    consolidate_results,
    run_quality_checks,
    build_unified_mapping,
    _cleanup_bom_files,
    get_available_models,
    OUTPUT_DIR as BOM_OUTPUT_DIR,
)
from mid_term_supermodels.matching import (
    match_skus_ultra_optimized,
    match_skus_preserve_run_vectors,
    build_column_mapping_auto,
    build_solo_modello_flags,
    load_lead_times,
    _reinit_norm_cache,
    _load_cambio_eur,
    canonical_model_name,
    save_results,
    save_results_monthly,
    CambioValutaError,
)
from mid_term_supermodels.montecarlo import (
    SimulationConfig as MCSimulationConfig,
    run_monthly_simulations_optimized,
    build_wide_dataframe,
    apply_pack_colors_to_wide_df,
    build_zcol_pack_color_mapping,
    load_monthly_forecast,
    calculate_n_months_needed,
    parse_exclusions,
    rescale_alphas,
)

log = get_logger("pipeline")


def _write_codici_varianti_modello(catalog_path, solo_skus) -> None:
    """Appende il foglio 'Codici_Varianti_Modello' (lista SKU SOLO_MODELLO,
    cioe' i codici dipendenti solo dal modello) al consolidato
    tutte_righe_univoche_V2.xlsx. Sovrascrive il foglio se gia' presente."""
    from pathlib import Path
    p = Path(catalog_path)
    if not p.exists():
        log.warning("Codici_Varianti_Modello: consolidato non trovato (%s) - foglio non scritto", p)
        return
    try:
        df = pd.DataFrame({"SKU": list(solo_skus)})
        with pd.ExcelWriter(p, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="Codici_Varianti_Modello", index=False)
        log.info("Foglio 'Codici_Varianti_Modello' scritto: %d SKU -> %s", len(df), p)
    except Exception as e:
        log.warning("Codici_Varianti_Modello: impossibile scrivere il foglio (%s): %s", p, e)


class SafetyStockPipeline:
    """Orchestratore end-to-end della pipeline Safety Stock Classica.

    Espone:
      - run() esegue tutti gli step 0..4 in sequenza
      - run_step_N() invocabile singolarmente (utile per debug, test parziali,
        integrazione con orchestratori esterni tipo Airflow)

    Args:
        config:     PipelineConfig con path file + flag I/O
        sim_config: SimulationConfig con parametri MC + percentili
        validate_config: se True (default) valida la config in __init__
    """

    def __init__(
        self,
        config: PipelineConfig,
        sim_config: Optional[SimulationConfig] = None,
        validate_config: bool = True,
    ):
        """Inizializza la pipeline con la configurazione path/IO e i parametri di
        simulazione. Se validate_config=True verifica subito l'esistenza dei file
        di input richiesti (fail-fast)."""
        self.config = config
        self.sim_config = sim_config or SimulationConfig()
        if validate_config:
            self.config.validate(strict=True)

    def run(self) -> SafetyStockResult:
        """Esegue l'intera pipeline Safety Stock in sequenza (step 0->4) e
        restituisce il risultato finale:
          0) carica TR/forecast/esclusioni; 1) costruisce la BOM consolidata;
          2) simula la domanda (Monte Carlo); 3) abbina gli SKU agli scenari;
          4) calcola la safety stock per SKU/mese + valorizzazione economica.
        """
        tr  = self.run_step_0_load_tr()
        bom = self.run_step_1_build_bom()
        sim = self.run_step_2_montecarlo(tr)
        mat = self.run_step_3_matching(sim, bom)
        ss  = self.run_step_4_safety_stock(mat, bom)
        return ss

    # ---- STEP 0 ------------------------------------------------------------
    def run_step_0_load_tr(self) -> TRData:
        """Passo 0 - In parole semplici: carica i dati di partenza.

        Legge il TR (mix modelli + take rate), il forecast di domanda e le regole
        di esclusione; prepara gli alpha Dirichlet in base all'affidabilita'.
        """
        t0 = time.time()
        tr_path         = str(self.config.resolve_input(self.config.tr_file))
        forecast_path   = str(self.config.resolve_input(self.config.forecast_file))
        exclusions_path = str(self.config.resolve_input(self.config.exclusions_file))

        log.info("Step 0: caricamento TR %s", tr_path)
        monthly_tr = load_monthly_tr(tr_path)
        model_mix, characteristics = parse_tr_file(tr_path)

        # Affidabilità TR: scala gli alpha Dirichlet (Model L0 / Optional L1).
        # parse_tr_file lascia alpha = TR grezzo; rescale_alphas applica i VSS.
        # Stesso ordine di mc_simulator_rolling (parse -> rescale -> simulate).
        log.info("Step 0: affidabilità TR model=%s optional=%s",
                 self.sim_config.affidabilita_model,
                 self.sim_config.affidabilita_optional)
        model_mix, characteristics, monthly_tr = rescale_alphas(
            model_mix, characteristics, monthly_tr,
            aff_model=self.sim_config.affidabilita_model,
            aff_optional=self.sim_config.affidabilita_optional,
        )

        log.info("Step 0: caricamento forecast %s", forecast_path)
        forecast_values, month_names = load_monthly_forecast(
            forecast_path, start_offset=self.sim_config.start_month_offset
        )

        # NOTA: calculate_n_months_needed() legge hardcoded Input/residual_lead_time.xlsx.
        # Limitazione nota: refactoring previsto in Fase 8.
        n_months_needed = calculate_n_months_needed()

        log.info("Step 0: caricamento esclusioni %s", exclusions_path)
        exclusions = parse_exclusions(exclusions_path)

        elapsed = time.time() - t0
        log.info("Step 0 completato in %.1fs: %d caratteristiche, %d mesi",
                 elapsed, len(characteristics), len(month_names))

        return TRData(
            model_mix=model_mix,
            characteristics=characteristics,
            monthly_tr=monthly_tr,
            forecast_values=list(forecast_values),
            forecast_months=list(month_names),
            exclusions=exclusions,
            n_months_needed=n_months_needed,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 1 ------------------------------------------------------------
    def run_step_1_build_bom(self) -> BOMData:
        """Passo 1 - In parole semplici: costruisce la distinta base consolidata.

        Costruisce matrice BOM + catalogo configurazioni multi-supermodel.

        Replica la logica di process_all_bom_combinations.main():
          1. build_unified_mapping (genera Mappatura_Unificata.xlsx)
          2. run_quality_checks (TR char options, prezzi, states)
          3. process_combination per ogni supermodel -> matrici per BOM
          4. consolidate_results -> catalogo finale deduplicato

        NOTE LIMITAZIONI:
          - process_combination scrive sempre matrix_output_{nome}.xlsx
            e tutte_righe_univoche_{nome}.xlsx su OUTPUT_DIR hardcoded
            (mid_term.bom.OUTPUT_DIR = Input/). write_intermediate qui
            controlla SOLO il salvataggio del catalogo finale.
          - run_quality_checks legge da path hardcoded INPUT_DIR/.
          - build_unified_mapping legge da MAPPATURE_ORIGINALI hardcoded.
        """
        t0 = time.time()

        # Discovery dinamica dei modelli da filesystem (come SolMidTerm-ALL):
        # scansiona input_dir/MODEL/* cercando le cartelle con Mappatura_<nome>.xlsx,
        # invece di affidarsi alla lista statica config.bom_files. Cosi' un modello
        # aggiunto/rimosso a filesystem viene preso/saltato in automatico.
        # Fallback alla config se la discovery non trova nulla (layout diverso).
        _discovered = get_available_models(self.config.input_dir / "MODEL")
        if _discovered:
            supermodels = list(_discovered.keys())
            log.info("Step 1: modelli scoperti da filesystem (%s): %s",
                     self.config.input_dir / "MODEL", supermodels)
        else:
            supermodels = list(self.config.bom_files.keys())
            log.warning("Step 1: nessun modello trovato in %s -> fallback a config.bom_files: %s",
                        self.config.input_dir / "MODEL", supermodels)
        log.info("Step 1: build BOM per supermodel %s", supermodels)

        # 0. Cleanup file BOM intermedi stali da run precedenti (replica
        # _cleanup_safety_stock_files di SolMidTerm-ALL, FASE -1 di main()).
        # Pulisce sia la dir hardcoded dove process_combination scrive
        # (mid_term.bom.OUTPUT_DIR == .\Input) sia le dir di config, deduplicate.
        _cleanup_dirs = {Path(BOM_OUTPUT_DIR).resolve(),
                         Path(self.config.input_dir).resolve(),
                         Path(self.config.output_dir).resolve()}
        for _d in _cleanup_dirs:
            _cleanup_bom_files(_d)

        # 1. Mappatura unificata: passa le mappature dei modelli scoperti per il
        #    supermodel corrente (config.input_dir/MODEL/<modello>/Mappatura_*),
        #    invece della costante MAPPATURE_ORIGINALI di bom.py (calcolata
        #    all'import su ./Input/MODEL generico, vuota nel layout multi-supermodel).
        _mappature_src = [info["mappatura"] for info in _discovered.values()
                          if info.get("mappatura")]
        build_unified_mapping(source_files=_mappature_src if _mappature_src else None)

        # 2. Quality checks (legge file hardcoded in mid_term.bom)
        residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set = \
            run_quality_checks()

        # 2b. Esclusioni optional incompatibili: usate per espandere le righe BOM
        # con coppie incompatibili (split deterministico, come SolMidTerm-ALL).
        # Caricate qui in modo autonomo (mirror di process_all_bom_combinations.main).
        exclusions_path = str(self.config.resolve_input(self.config.exclusions_file))
        log.info("Step 1: caricamento esclusioni %s", exclusions_path)
        exclusions = parse_exclusions(exclusions_path)

        # 3. Per ogni supermodel costruisci la matrice via process_combination.
        # process_combination si aspetta un dict con chiavi:
        #   nome, in_dir, mappatura, tr
        # Usiamo i path della config quando possibile.
        matrices = {}
        tr_path = self.config.resolve_input(self.config.tr_file)

        for supermodel in supermodels:
            # Path da discovery quando disponibile (filesystem), altrimenti
            # ricostruiti dalla struttura legacy Input/MODEL/<supermodel>/.
            if supermodel in _discovered:
                in_dir = _discovered[supermodel]["in_dir"]
                mappatura_path = _discovered[supermodel]["mappatura"]
            else:
                in_dir = self.config.input_dir / "MODEL" / supermodel
                mappatura_path = in_dir / f"Mappatura_{supermodel}.xlsx"

            log.info("Step 1: processing %s (in_dir=%s)", supermodel, in_dir)
            combo = {
                "nome": supermodel,
                "in_dir": in_dir,
                "mappatura": mappatura_path,
                "tr": tr_path,
            }
            matrix_df = process_combination(
                combo,
                tr_char_options=tr_char_options,
                filter_untraced=self.config.filter_untraced,
                exclusions=exclusions,
            )
            matrices[supermodel] = matrix_df

            if self.config.write_intermediate:
                # process_combination ha gia' scritto su OUTPUT_DIR hardcoded;
                # copiamo/risalviamo nella output_dir richiesta dalla pipeline
                # solo se diverso da quel default.
                out_path = self.config.resolve_output(
                    self.config.matrix_filename_tpl.format(supermodel=supermodel)
                )
                try:
                    if out_path.parent.resolve() != \
                            (self.config.input_dir).resolve():
                        matrix_df.to_excel(out_path, index=False)
                        log.info("Step 1: matrice copiata in %s", out_path)
                except Exception as e:
                    log.warning("Step 1: impossibile salvare matrice in output_dir: %s", e)

        # 4. Consolida tutte le matrici in catalogo unico
        df_consolidated = consolidate_results(matrices)

        # 5. Quality report come DataFrame (riepilogo issues)
        quality_counts = len(residual_issues) + len(price_issues) + len(ct_cv_issues)
        quality_report = pd.DataFrame({
            "Categoria": ["Residual Issues", "Price Issues", "CT/CV Issues", "TOTALE"],
            "Conteggio": [len(residual_issues), len(price_issues),
                          len(ct_cv_issues), quality_counts],
        })

        # 6. Salva catalogo consolidato se richiesto
        if self.config.write_intermediate:
            catalog_path = self.config.resolve_output(self.config.catalog_filename)
            try:
                with pd.ExcelWriter(catalog_path, engine="openpyxl") as writer:
                    df_consolidated.to_excel(
                        writer, sheet_name="Tutte_Righe_Univoche", index=False
                    )
                    for nome, df in matrices.items():
                        df.to_excel(writer, sheet_name=f"Righe_{nome}", index=False)
                    quality_report.to_excel(
                        writer, sheet_name="Quality_Summary", index=False
                    )
                log.info("Step 1: catalogo consolidato salvato %s", catalog_path)
            except Exception as e:
                log.warning("Step 1: impossibile salvare catalogo: %s", e)

        # 7. Rigenera il catalogo anche in Input/ (NON gated da write_intermediate):
        # lo step 3 (calcola_residual_lead_time) lo rilegge da .\Input via read_excel.
        # Replica il legacy process_all_bom_combinations.main() che salvava in
        # OUTPUT_DIR == Input/. Senza questo, step 3 -> FileNotFoundError.
        input_catalog_path = self.config.resolve_input(self.config.catalog_filename)
        try:
            df_consolidated.to_excel(
                input_catalog_path, sheet_name="Tutte_Righe_Univoche", index=False
            )
            log.info("Step 1: catalogo rigenerato in Input %s", input_catalog_path)
        except Exception as e:
            log.warning("Step 1: impossibile rigenerare catalogo in Input: %s", e)

        elapsed = time.time() - t0
        log.info("Step 1 completato in %.1fs: %d supermodel, %d righe catalogo",
                 elapsed, len(matrices), len(df_consolidated))

        return BOMData(
            catalog=df_consolidated,
            matrices=matrices,
            quality_report=quality_report,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 2 ------------------------------------------------------------
    def run_step_2_montecarlo(self, tr: TRData) -> SimulationResult:
        """Passo 2 - In parole semplici: simula la domanda con il metodo Monte Carlo.

        Genera molti scenari di vendita usando TR ed esclusioni dello step 0, poi
        deriva i colori dei pack e produce la tabella degli scenari.
        """
        t0 = time.time()
        # Mapping campo: max_conflict_iterations (config pubblico)
        #             -> max_conflict_resolution_iterations (interno)
        mc_cfg = MCSimulationConfig(
            n_runs=self.sim_config.n_runs,
            random_seed=self.sim_config.random_seed,
            start_month_offset=self.sim_config.start_month_offset,
            use_int16=self.sim_config.use_int16,
            max_conflict_resolution_iterations=self.sim_config.max_conflict_iterations,
        )
        log.info("Step 2: MC n_runs=%d seed=%d", mc_cfg.n_runs, mc_cfg.random_seed)

        results_by_month = run_monthly_simulations_optimized(
            forecast_values=tr.forecast_values,
            month_names=tr.forecast_months,
            n_months=tr.n_months_needed,
            model_mix=tr.model_mix,
            characteristics=tr.characteristics,
            exclusions=tr.exclusions,
            config=mc_cfg,
            monthly_tr=tr.monthly_tr,
        )

        # build_wide_dataframe firma reale: (results_by_month, month_names, n_runs, ...)
        df_wide = build_wide_dataframe(
            results_by_month,
            tr.forecast_months,
            mc_cfg.n_runs,
        )

        # Pack colors mapping (post-processing): costruito PER-MODELLO e fuso.
        # Per ogni modello scoperto si legge la sua Mappatura_<ver>.xlsx, usata sia
        # come mappatura standard sia come Rally (il bundle "touring pack rally"
        # viene riconosciuto solo dove presente). Le mappe colore dei modelli vengono
        # fuse in un unico dizionario "colore carena -> colore pack".
        # Vantaggi vs vecchio approccio (unificata + V22E hardcoded): nessuna
        # dipendenza dalla Mappatura_Unificata, nessun modello Rally cablato,
        # coerente con la discovery dinamica dello Step 1. Verificato equivalente
        # (0 collisioni, copertura >=) al vecchio metodo.
        models = get_available_models(self.config.input_dir / "MODEL")
        pack_color_mapping = {}
        for ver, info in models.items():
            mp = str(info["mappatura"])
            try:
                m = build_zcol_pack_color_mapping(mp, rally_mappatura_path=mp)
            except Exception as e:
                log.warning("Pack colors: mappatura %s non elaborabile: %s", ver, e)
                continue
            for k, sub in m.items():
                pack_color_mapping.setdefault(k, {}).update(sub)
        if not pack_color_mapping:
            # Fallback: mappatura unificata di config se la discovery non produce nulla
            try:
                pack_color_mapping = build_zcol_pack_color_mapping(
                    str(self.config.resolve_input(self.config.mapping_file))
                )
            except Exception as e:
                log.warning("Pack colors: fallback unificata fallito: %s", e)
        try:
            df_wide = apply_pack_colors_to_wide_df(df_wide, pack_color_mapping)
        except Exception as e:
            log.warning("Pack colors: applicazione fallita: %s", e)

        elapsed = time.time() - t0
        log.info("Step 2 completato in %.1fs: %d configurazioni",
                 elapsed, len(df_wide))

        return SimulationResult(
            df_wide=df_wide,
            n_runs=mc_cfg.n_runs,
            n_months=tr.n_months_needed,
            month_names=tr.forecast_months,
            elapsed_seconds=elapsed,
        )

    # ---- STEP 3 ------------------------------------------------------------
    def run_step_3_matching(self, sim: SimulationResult, bom: BOMData, lead_times=None) -> MatchResult:
        """Passo 3 - In parole semplici: abbina ogni componente agli scenari.

        SKU matching: confronta catalogo BOM con output simulazione MC.

        Replica la prima meta' di sku_matching_SS.main_sku_v2():
          1. Prepara catalogo (rimuove ID, dedup)
          2. Carica lead times (path hardcoded in mid_term.matching)
          3. Costruisce column_mapping auto-detect
          4. Calcola flag SOLO_MODELLO
          5. Re-inizializza NormalizationCache con prefissi dinamici
          6. Esegue match_skus_ultra_optimized -> DataFrame matched

        NOTE LIMITAZIONI:
          - load_lead_times() legge sempre Input/residual_lead_time.xlsx (hardcoded).
        """
        t0 = time.time()
        percentile = max(1, min(99, int(self.sim_config.percentile_safety_stock)))
        log.info("Step 3: SKU matching (percentile=%d)", percentile)

        simulation_output = sim.df_wide

        # 1. Prepara catalogo SKU dal BOM step
        sku_catalog_raw = bom.catalog.copy()
        if "ID" in sku_catalog_raw.columns:
            sku_catalog_raw = sku_catalog_raw.drop(columns=["ID"])
        sku_catalog = sku_catalog_raw.drop_duplicates(keep="first")
        log.info("Step 3: catalogo %d righe (raw=%d)",
                 len(sku_catalog), len(sku_catalog_raw))

        # 2. (lead times) Caricati piu' sotto, DOPO il calcolo di SOLO_MODELLO:
        # il set SOLO_MODELLO serve a calcola_residual_lead_time per imporre il
        # gate parametrico a quelle SKU. Se lead_times e' passato (rolling) il
        # gate SOLO e' gia' incluso e si riusa cosi' com'e'.

        # 3. Identifica n_runs dalle colonne _Run_
        month_run_cols = [c for c in simulation_output.columns if "_Run_" in c]
        if month_run_cols:
            first_prefix = month_run_cols[0].split("_Run_")[0]
            first_month_cols = [c for c in month_run_cols
                                if c.startswith(first_prefix)]
            n_runs = len(first_month_cols)
        else:
            n_runs = sim.n_runs
        log.info("Step 3: n_runs rilevati=%d", n_runs)

        # 4. Costruisci column_mapping auto-detect
        # 'versione' aggiunta da process_all_bom_combinations (catalogo multi-BOM)
        # e non e' una caratteristica di configurazione: va esclusa.
        sku_feature_cols = [c for c in sku_catalog.columns
                            if c.lower().strip() not in
                            {"id", "numero componenti",
                             "numero componenti.1", "versione"}]
        sim_config_cols = [c for c in simulation_output.columns
                           if "_Run_" not in c]
        column_mapping = build_column_mapping_auto(sku_feature_cols, sim_config_cols)
        log.info("Step 3: column_mapping (%d colonne)", len(column_mapping))
        for sku_col, sim_col in sorted(column_mapping.items()):
            log.debug("  '%s' -> '%s'", sku_col, sim_col)

        # 5. Calcola flag SOLO_MODELLO (helper condiviso con il rolling)
        df_flag_solo_modello = build_solo_modello_flags(sku_catalog_raw, column_mapping)
        _n_solo = int(df_flag_solo_modello["SOLO_MODELLO"].sum()) \
            if len(df_flag_solo_modello) > 0 else 0
        log.info("Step 3: SOLO_MODELLO %d/%d SKU",
                 _n_solo, len(df_flag_solo_modello))
        _solo_skus = (set(df_flag_solo_modello.loc[df_flag_solo_modello["SOLO_MODELLO"], "SKU"]
                          .astype(str))
                      if len(df_flag_solo_modello) else set())

        # 5b. Carica lead times (se non passati dal chiamante, es. rolling).
        # gate_max_mesi: cap globale sul gate di ogni SKU + default per le
        # SOLO_MODELLO senza gate. Residual = formula standard (FROZEN+Freq - gate);
        # senza dato supplier -> scartata.
        if lead_times is None:
            lead_times = load_lead_times(
                solo_modello_skus=_solo_skus,
                gate_max_mesi=self.sim_config.gate_max_mesi,
            )
        log.info("Step 3: lead_times %d SKU", len(lead_times))

        # 5c. Scrive il foglio 'Codici_Varianti_Modello' (lista SKU SOLO_MODELLO)
        # nel consolidato tutte_righe_univoche_V2.xlsx. Solo nel run finale
        # (write_final_output): nel rolling write_final_output=False -> saltato.
        if self.config.write_final_output and _solo_skus:
            _write_codici_varianti_modello(
                self.config.resolve_output(self.config.catalog_filename),
                sorted(_solo_skus),
            )

        # 6. Re-inizializza NormalizationCache con prefissi dinamici
        _reinit_norm_cache(column_mapping)

        # 7. Matching (include espansione multi-modello)
        df_matched, df_no_lt = match_skus_ultra_optimized(
            sku_catalog=sku_catalog,
            lead_times=lead_times,
            simulation_output=simulation_output,
            column_mapping=column_mapping,
            n_runs=n_runs,
            percentile=percentile,
        )

        elapsed = time.time() - t0
        log.info("Step 3 completato in %.1fs: %d righe matched, %d SKU senza LT",
                 elapsed,
                 len(df_matched) if df_matched is not None else 0,
                 len(df_no_lt) if df_no_lt is not None else 0)

        return MatchResult(
            df_matched=df_matched,
            column_mapping=column_mapping,
            solo_modello=df_flag_solo_modello,
            no_lt=df_no_lt,
            elapsed_seconds=elapsed,
        )

    # ---- HOOK MULTI-SUPERMODEL --------------------------------------------
    def run_collect_run_vectors(self, lead_times=None):
        """Esegue step 0-2 standard e poi il matching che PRESERVA i vettori per-run.

        Variante di run_step_3_matching pensata per l'orchestratore multi-supermodel:
        prepara gli STESSI input (sku_catalog, lead_times, column_mapping, n_runs,
        percentile) ESATTAMENTE come run_step_3_matching, ma sostituisce la chiamata
        finale match_skus_ultra_optimized con match_skus_preserve_run_vectors, che
        restituisce i vettori di domanda per-run (in bici) anziche' le statistiche.

        NOTA import: match_skus_preserve_run_vectors arriva dall'import a livello
        modulo `from mid_term_supermodels.matching import ...` (package COPIATO,
        l'unico che la espone). Idem gli altri helper (load_lead_times,
        build_solo_modello_flags, _reinit_norm_cache, build_column_mapping_auto):
        tutti risolti sul package self-contained mid_term_supermodels.

        Returns:
            ( { sku : (run_demands_bici[n_runs], qty, lead_time) }, df_no_lt )
        """
        tr  = self.run_step_0_load_tr()
        _bom = self.run_step_1_build_bom()
        sim = self.run_step_2_montecarlo(tr)

        # --- Replica della preparazione input di run_step_3_matching ---
        # (copia 1:1 della logica righe 441-516; ci si ferma PRIMA del matching)
        percentile = max(1, min(99, int(self.sim_config.percentile_safety_stock)))
        log.info("Collect run vectors: SKU matching (percentile=%d)", percentile)

        simulation_output = sim.df_wide

        # 1. Prepara catalogo SKU dal BOM step
        sku_catalog_raw = _bom.catalog.copy()
        if "ID" in sku_catalog_raw.columns:
            sku_catalog_raw = sku_catalog_raw.drop(columns=["ID"])
        sku_catalog = sku_catalog_raw.drop_duplicates(keep="first")
        log.info("Collect run vectors: catalogo %d righe (raw=%d)",
                 len(sku_catalog), len(sku_catalog_raw))

        # 3. Identifica n_runs dalle colonne _Run_
        month_run_cols = [c for c in simulation_output.columns if "_Run_" in c]
        if month_run_cols:
            first_prefix = month_run_cols[0].split("_Run_")[0]
            first_month_cols = [c for c in month_run_cols
                                if c.startswith(first_prefix)]
            n_runs = len(first_month_cols)
        else:
            n_runs = sim.n_runs
        log.info("Collect run vectors: n_runs rilevati=%d", n_runs)

        # 4. Costruisci column_mapping auto-detect
        sku_feature_cols = [c for c in sku_catalog.columns
                            if c.lower().strip() not in
                            {"id", "numero componenti",
                             "numero componenti.1", "versione"}]
        sim_config_cols = [c for c in simulation_output.columns
                           if "_Run_" not in c]
        column_mapping = build_column_mapping_auto(sku_feature_cols, sim_config_cols)
        log.info("Collect run vectors: column_mapping (%d colonne)", len(column_mapping))

        # 5. Calcola flag SOLO_MODELLO (helper condiviso con il rolling)
        df_flag_solo_modello = build_solo_modello_flags(sku_catalog_raw, column_mapping)
        _solo_skus = (set(df_flag_solo_modello.loc[df_flag_solo_modello["SOLO_MODELLO"], "SKU"]
                          .astype(str))
                      if len(df_flag_solo_modello) else set())

        # 5b. Carica lead times (se non passati dal chiamante)
        if lead_times is None:
            lead_times = load_lead_times(
                solo_modello_skus=_solo_skus,
                gate_max_mesi=self.sim_config.gate_max_mesi,
            )
        log.info("Collect run vectors: lead_times %d SKU", len(lead_times))

        # 6. Re-inizializza NormalizationCache con prefissi dinamici
        _reinit_norm_cache(column_mapping)

        # 7. Matching che PRESERVA i vettori per-run (al posto di
        #    match_skus_ultra_optimized): stessi input di run_step_3_matching.
        vectors, df_no_lt = match_skus_preserve_run_vectors(
            sku_catalog=sku_catalog,
            lead_times=lead_times,
            simulation_output=simulation_output,
            column_mapping=column_mapping,
            n_runs=n_runs,
            percentile=percentile,
        )
        return vectors, df_no_lt

    # ---- STEP 4 ------------------------------------------------------------
    def run_step_4_safety_stock(self, mat: MatchResult, bom: BOMData) -> SafetyStockResult:
        """Passo 4 - In parole semplici: calcola la safety stock e la valorizza.

        Applica prezzi + merge SOLO_MODELLO + scrive output Excel.

        Replica la seconda meta' di sku_matching_SS.main_sku_v2():
          1. Merge prezzi EUR (se config.prices_file disponibile)
          2. Merge flag SOLO_MODELLO
          3. Scrive output aggregato per SKU
          4. Scrive output mensile per SKU
        """
        t0 = time.time()
        percentile = max(1, min(99, int(self.sim_config.percentile_safety_stock)))
        results = mat.df_matched
        df_flag_solo_modello = mat.solo_modello

        # 1. Merge prezzi con conversione valuta -> EUR
        if self.config.prices_file:
            try:
                # PREZZI e' CONDIVISO in Dati Logistica (override via env
                # MIDTERM_SHARED_LOGISTICA, default <cwd>/Input/Dati Logistica);
                # fallback alla input_dir del supermodel solo se assente.
                _logistica = shared_logistica_dir()
                _shared_matches = sorted(_logistica.glob("PREZZI*.XLS*"))
                if _shared_matches:
                    prezzi_path = _shared_matches[0]
                else:
                    _prezzi_matches = sorted(self.config.input_dir.glob("PREZZI*.XLS*"))
                    prezzi_path = (_prezzi_matches[0] if _prezzi_matches
                                   else _logistica / "PREZZI.XLSX")
                all_data = pd.read_excel(prezzi_path, sheet_name="Sheet1")

                material_col = "Materiale"
                prezzo_netto_col = "Prezzo netto"
                data_doc_col = "Data documento"
                unit_prezzo_col = None
                divisa_col = None
                for col in all_data.columns:
                    col_lower = str(col).lower().replace("\xa0", " ").replace("\x92", "'")
                    if "unit" in col_lower and "prezzo" in col_lower:
                        unit_prezzo_col = col
                        break
                for col in all_data.columns:
                    col_lower = str(col).lower().replace("\xa0", " ").replace("\x92", "'")
                    if "divisa" in col_lower or "currency" in col_lower:
                        divisa_col = col
                        break

                if unit_prezzo_col is None or divisa_col is None:
                    raise ValueError(
                        f"Required columns not found. unit_prezzo_col={unit_prezzo_col}, "
                        f"divisa_col={divisa_col}"
                    )

                prezzi = all_data[[material_col, prezzo_netto_col, unit_prezzo_col,
                                   divisa_col, data_doc_col]].drop_duplicates()
                _prezzo_num = pd.to_numeric(prezzi[prezzo_netto_col], errors="coerce")
                _unit_num = pd.to_numeric(prezzi[unit_prezzo_col],
                                          errors="coerce").replace(0, np.nan)
                prezzi = prezzi.copy()
                prezzi["Prezzo netto"] = _prezzo_num / _unit_num

                # Tassi di cambio EUR (path hardcoded in matching._load_cambio_eur)
                cambio = _load_cambio_eur()
                log.info("Step 4: tassi di cambio caricati: %d valute", len(cambio))

                prezzi[divisa_col] = (prezzi[divisa_col].astype(str)
                                      .str.strip().str.upper())
                prezzi = prezzi.merge(cambio, left_on=divisa_col,
                                      right_on="Codice", how="left")

                _miss_mask = prezzi["1 unità in EUR ≈"].isna()
                if _miss_mask.any():
                    _missing_curr = sorted(prezzi.loc[_miss_mask, divisa_col]
                                           .dropna().astype(str).unique())
                    raise CambioValutaError(
                        f"Valute presenti in Prezzi.xlsx ma non mappate in "
                        f"Cambio Valuta.xlsx: {_missing_curr}\n"
                        f"Aggiungere queste valute al file Cambio Valuta prima "
                        f"di rilanciare."
                    )

                prezzi["Prezzo netto EUR"] = (prezzi["Prezzo netto"]
                                              * prezzi["1 unità in EUR ≈"])

                prezzi = (prezzi[[material_col, data_doc_col, "Prezzo netto EUR"]]
                          .sort_values(data_doc_col)
                          .groupby(material_col)[["Prezzo netto EUR"]]
                          .last()
                          .rename(columns={"Prezzo netto EUR": "Prezzo netto"}))

                ss_key = f"safety_stock_p{percentile}_total"
                results = results.merge(prezzi, left_on="SKU",
                                        right_index=True, how="left")
                results["prezzo_safety"] = results["Prezzo netto"] * results[ss_key]
            except CambioValutaError:
                raise
            except Exception as e:
                log.warning("Step 4: prezzi non disponibili: %s", e)

        # 2. Merge SOLO_MODELLO
        if results is None or len(results) == 0:
            log.warning("Step 4: nessun risultato dal matching - risultati vuoti")
            results = pd.DataFrame(columns=["SKU", "SOLO_MODELLO"])
        elif "SKU" not in results.columns:
            log.warning("Step 4: colonna 'SKU' assente in results")
            results["SOLO_MODELLO"] = False
        else:
            results = results.merge(df_flag_solo_modello, on="SKU", how="left")
            results["SOLO_MODELLO"] = results["SOLO_MODELLO"].fillna(False)

        # 3. Scrittura output (se richiesto)
        output_paths = {}
        if self.config.write_final_output:
            per_sku_path = self.config.resolve_output(
                self.config.safety_stock_per_sku_filename
            )
            per_mese_path = self.config.resolve_output(
                self.config.safety_stock_per_mese_filename
            )
            log.info("Step 4: salvataggio %s e %s", per_sku_path, per_mese_path)
            save_results(results, str(per_sku_path), percentile=percentile)
            save_results_monthly(results, str(per_mese_path), percentile=percentile)
            output_paths["per_sku"] = str(per_sku_path)
            output_paths["per_mese"] = str(per_mese_path)

            # SKU esclusi per lead time mancante/zero (replica output_no_lt di
            # main_sku_v2 in SolMidTerm-ALL). Diagnostica: non entra nel calcolo SS.
            no_lt_path = self.config.resolve_output(self.config.sku_senza_lt_filename)
            mat.no_lt.to_excel(str(no_lt_path), index=False)
            output_paths["sku_senza_lt"] = str(no_lt_path)
            log.info("Step 4: salvati %d SKU senza lead time in %s",
                     len(mat.no_lt), no_lt_path)

        # 4. Costruisci DataFrame finale per SafetyStockResult
        # save_results scrive il file aggregato; in memoria, per_sku = results,
        # per_mese = vista per-mese derivata da results (stessa logica di save_results_monthly).
        per_sku = results
        per_mese = results  # save_results_monthly espande lo stesso df

        elapsed = time.time() - t0
        log.info("Step 4 completato in %.1fs: %d righe finali",
                 elapsed, len(per_sku))

        return SafetyStockResult(
            per_sku=per_sku,
            per_mese=per_mese,
            percentile=percentile,
            output_paths=output_paths,
            elapsed_seconds=elapsed,
        )

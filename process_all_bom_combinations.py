# -*- coding: utf-8 -*-
"""Shim retrocompatibile: riesporta API da mid_term.bom.

I nuovi consumer devono usare direttamente `from mid_term import ...` o
`from mid_term.bom import ...`. Questo modulo esiste per non rompere
consumer legacy (app_V3.py — che importa via `import process_all_bom_combinations
as _pabc` e chiama `_pabc.main(filter_untraced=...)`).

CONCERN/BLOCKER PARZIALE — funzioni NON migrate in mid_term/bom.py:
- `main(filter_untraced=True)`: entry point CLI orchestratore Multi-BOM,
  invocato da app_V3.py Step 1.
- `_cleanup_safety_stock_files`: helper di cleanup file stali, usato solo
  internamente da `main()`.
Per non rompere `_pabc.main(...)` di app_V3, entrambe le funzioni sono
MANTENUTE inline qui (logica IDENTICA all'originale). Le building block
(`process_combination`, `run_quality_checks`, `consolidate_results`,
`build_unified_mapping`, `COMBINAZIONI`, `MAPPATURE_ORIGINALI`,
`_MAPPATURA_UNIFICATA`, `OUTPUT_DIR`, `get_available_models`) sono importate
da `mid_term.bom` — single source of truth.

Le funzioni `_get_available_matrix_files`/`_cleanup_safety_stock_files`
dello skubundle originale NON erano importate da consumer esterni (verificato
con grep): nessuna esposizione necessaria nei rispettivi shim.
"""
from pathlib import Path
import pandas as pd

# Re-export wildcard di tutti i simboli pubblici migrati
from mid_term.bom import *  # noqa: F401,F403

# Simboli espliciti effettivamente importati dai consumer legacy / usati da main()
from mid_term.bom import (  # noqa: F401
    process_combination,
    run_quality_checks,
    consolidate_results,
    build_unified_mapping,
    get_available_models,
    AVAILABLE_MODELS,
    COMBINAZIONI,
    MAPPATURE_ORIGINALI,
    INPUT_DIR,
    OUTPUT_DIR,
    MODEL_DIR,
    PREZZI_FILE,
    STATES_FILE,
    _MAPPATURA_UNIFICATA,
    _normalize_bundle,
)


# ============================================================================
# CLEANUP + ENTRY POINT main() — non migrati in mid_term/bom.py.
# Mantenuti inline per retro-compatibilita' con app_V3.py
# (`import process_all_bom_combinations as _pabc; _pabc.main(...)`).
# Logica IDENTICA all'originale.
# ============================================================================


# ============================================================================
# CLEANUP: ELIMINA FILE STALI
# ============================================================================

def _cleanup_safety_stock_files(input_dir: str = ".\\Input") -> None:
    """
    Elimina tutti i file matrix_output_*.xlsx e tutte_righe_univoche_*.xlsx
    (esclude SA che sono trattati da crea_catalogo_sa.py).
    Eseguito all'inizio di main() per evitare file stali.
    """
    input_path = Path(input_dir)

    # Elimina matrix_output (esclude SA)
    print("\n[CLEANUP] Rimozione matrix_output precedenti (non-SA)...")
    for f in input_path.glob("matrix_output_*.xlsx"):
        if "_SA_" not in f.name:
            try:
                f.unlink()
                print(f"  Rimosso: {f.name}")
            except Exception as e:
                print(f"  [WARN] Impossibile rimuovere {f.name}: {e}")

    # Elimina tutte_righe_univoche (esclude SA)
    print("[CLEANUP] Rimozione tutte_righe_univoche precedenti (non-SA)...")
    for f in input_path.glob("tutte_righe_univoche_*.xlsx"):
        if "_SA_" not in f.name:
            try:
                f.unlink()
                print(f"  Rimosso: {f.name}")
            except Exception as e:
                print(f"  [WARN] Impossibile rimuovere {f.name}: {e}")


def main(filter_untraced: bool = True):
    """Esegui il processamento completo."""

    # =========================================================================
    # CLEANUP: Elimina file stali prima di rigenerare
    # =========================================================================
    _cleanup_safety_stock_files(input_dir=".\\Input")

    print("\n" + "="*80)
    print("PROCESSAMENTO MULTI-BOM: V21E + V22E + V24T")
    print("="*80)

    # =========================================================================
    # FASE -1: Costruisce la mappatura unificata a runtime dalle originali
    # =========================================================================
    build_unified_mapping(
        source_files=MAPPATURE_ORIGINALI,
        output_path=_MAPPATURA_UNIFICATA,
    )

    # =========================================================================
    # FASE 0: Quality checks globali
    # =========================================================================
    residual_issues, price_issues, ct_cv_issues, tr_char_options, states_set = run_quality_checks()

    # =========================================================================
    # FASE 1: Processa tutte le combinazioni
    # =========================================================================
    results = {}
    for combo in COMBINAZIONI:
        try:
            df = process_combination(combo, tr_char_options=tr_char_options, filter_untraced=filter_untraced)
            results[combo["nome"]] = df
        except Exception as e:
            print(f"\n[ERRORE] nel processamento {combo['nome']}: {e}")
            raise

    # =========================================================================
    # FASE 2: Consolida in un unico file
    # =========================================================================
    df_final = consolidate_results(results)

    # =========================================================================
    # FASE 3: Salva il file finale consolidato con Quality Summary
    # =========================================================================
    output_path = OUTPUT_DIR / "tutte_righe_univoche_V2.xlsx"

    print(f"\n{'='*80}")
    print("SALVATAGGIO FILE FINALE + QUALITY SUMMARY")
    print(f"{'='*80}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # Foglio principale: tutte le righe consolidate
        df_final.to_excel(writer, sheet_name='Tutte_Righe_Univoche', index=False)

        # Foglio per ogni versione (per tracciabilità)
        for nome, df in results.items():
            sheet_name = f"Righe_{nome}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Foglio riepilogo con deduplicazione
        riepilogo_data = {
            'Versione': list(results.keys()),
            'Righe Univoche': [len(df) for df in results.values()],
        }
        df_riepilogo = pd.DataFrame(riepilogo_data)
        df_riepilogo.loc[len(df_riepilogo)] = {
            'Versione': 'TOTALE CONSOLIDATO (dopo dedup)',
            'Righe Univoche': len(df_final)
        }
        df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)

        # Foglio deduplicazione: tracciabilità delle versioni BOM per ogni riga
        dedup_summary = df_final[['ID_Globale', 'Versioni_BOM']].copy()
        dedup_summary['Num_Versioni'] = dedup_summary['Versioni_BOM'].str.count(r'\+') + 1
        dedup_summary = dedup_summary.sort_values('Num_Versioni', ascending=False)
        dedup_summary.to_excel(writer, sheet_name='Dedup_Tracciabilita', index=False)

        # Foglio quality summary (riepilogo issues)
        quality_counts = len(residual_issues) + len(price_issues) + len(ct_cv_issues)
        quality_summary = {
            'Categoria': ['Residual Issues', 'Price Issues', 'CT/CV Issues', 'TOTALE'],
            'Conteggio': [len(residual_issues), len(price_issues), len(ct_cv_issues), quality_counts]
        }
        df_quality = pd.DataFrame(quality_summary)
        df_quality.to_excel(writer, sheet_name='Quality_Summary', index=False)

        if quality_counts > 0:
            print(f"\n[INFO] Quality issues rilevati: {quality_counts}")
            print(f"       - Residual: {len(residual_issues)}")
            print(f"       - Price: {len(price_issues)}")
            print(f"       - CT/CV: {len(ct_cv_issues)}")
        else:
            print(f"\n[OK] Nessun quality issue rilevato")

    print(f"\n[OK] File consolidato salvato:")
    print(f"  {output_path}")
    print(f"  Dimensioni: {len(df_final)} righe × {len(df_final.columns)} colonne")

    # Analisi righe univoche vs condivise
    df_final_for_analysis = df_final.copy()
    righe_univoche = (df_final_for_analysis['Versioni_BOM'].str.count(r'\+') == 0).sum()
    righe_condivise = (df_final_for_analysis['Versioni_BOM'].str.count(r'\+') > 0).sum()

    print(f"\nAnalisi Deduplicazione:")
    print(f"  - Righe UNIVOCHE (single version): {righe_univoche}")
    print(f"  - Righe CONDIVISE (multi version): {righe_condivise}")

    print(f"\nComposizione per Versione BOM Originale:")
    for nome in results.keys():
        count = len(results[nome])
        pct = (count / sum(len(df) for df in results.values()) * 100)
        print(f"  - {nome}: {count} righe ({pct:.1f}%)")

    return df_final


"""Test qty per-configurazione nel matching: più-specifica-vince + MAX tie-break.

Il vettore run_demands restituito da _compute_run_demands_for_sku è in PEZZI
(qty per-config già applicata).
"""
import numpy as np
import pandas as pd
from mid_term_supermodels import matching


def _ctx(cat, sim):
    lt = pd.DataFrame({'SKU': cat['Numero componenti'].unique(),
                       'Lead_Time_Months': 1.0, 'Lead_Time_Months_Ceil': 1.0,
                       'Description': 'c'})
    feat = [c for c in cat.columns if c not in ('Numero componenti', 'Qta_comp_UMC')]
    mapping = matching.build_column_mapping_auto(feat, list(sim.columns))
    return matching._build_matching_context(cat, lt, sim, mapping, n_runs=2)


def test_monthly_stats_in_pezzi():
    # match_skus_ultra_optimized: monthly_stats deve essere in PEZZI (qty per-config),
    # coerente con l'aggregato. Config Cast+X (10 bici) con qty 10 -> 100 pezzi/mese.
    sim = pd.DataFrame([
        {'MODEL': 'MSV4', 'RIMS': 'Cast', 'BRAKE': 'X', 'JAN 2026_Run_0': 10, 'JAN 2026_Run_1': 10},
    ])
    cat = pd.DataFrame({
        'Numero componenti': ['C', 'C'],
        'MODEL': ['MSV4', 'MSV4'],
        'RIMS': ['Cast', 'Cast'],
        'BRAKE': ['No', 'X'],
        'Qta_comp_UMC': [5.0, 10.0],
    })
    lt = pd.DataFrame({'SKU': ['C'], 'Lead_Time_Months': 1.0,
                       'Lead_Time_Months_Ceil': 1.0, 'Description': 'c'})
    mapping = matching.build_column_mapping_auto(['MODEL', 'RIMS', 'BRAKE'], list(sim.columns))
    df, _ = matching.match_skus_ultra_optimized(cat, lt, sim, mapping, n_runs=2, percentile=99)
    row = df[df['SKU'] == 'C'].iloc[0]
    # mensile in pezzi: 10 bici x qty10 = 100 (NON 10 bici, NON 10x5)
    assert abs(row['monthly_stats']['JAN 2026']['mean_demand'] - 100.0) < 1e-6
    # _total in pezzi == safety_stock (no doppia moltiplicazione)
    assert abs(row['safety_stock_p99_total'] - row['safety_stock_p99']) < 1e-6


def test_piu_specifica_vince():
    # config A = Cast+BrakeX (10 bici), config B = Cast+BrakeY (7 bici)
    sim = pd.DataFrame([
        {'MODEL': 'MSV4', 'RIMS': 'Cast', 'BRAKE': 'X', 'JAN 2026_Run_0': 10, 'JAN 2026_Run_1': 10},
        {'MODEL': 'MSV4', 'RIMS': 'Cast', 'BRAKE': 'Y', 'JAN 2026_Run_0': 7,  'JAN 2026_Run_1': 7},
    ])
    # componente C: riga generica {Cast} qty5 (spec2), riga specifica {Cast,BrakeX} qty10 (spec3)
    cat = pd.DataFrame({
        'Numero componenti': ['C', 'C'],
        'MODEL': ['MSV4', 'MSV4'],
        'RIMS': ['Cast', 'Cast'],
        'BRAKE': ['No', 'X'],          # 'No' = jolly (matcha qualsiasi brake)
        'Qta_comp_UMC': [5.0, 10.0],
    })
    ctx = _ctx(cat, sim)
    try:
        d = next(dd for sid, dd in zip(ctx.sku_ids, ctx.sku_data_list) if sid == 'C')
        rd, idx, _qpc = matching._compute_run_demands_for_sku(d, ctx)
        # config A (Cast+X): più specifica = riga {Cast,X} -> qty10 -> 10 bici x10 = 100
        # config B (Cast+Y): solo {Cast} -> qty5 -> 7 bici x5 = 35
        # totale per run = 135
        assert rd.tolist() == [135.0, 135.0]
    finally:
        ctx.cleanup()


def test_tie_break_max():
    # config MSV4 (10 bici); due righe stesse caratteristiche (solo Model) qty 4 e 5
    sim = pd.DataFrame([
        {'MODEL': 'MSV4', 'JAN 2026_Run_0': 10, 'JAN 2026_Run_1': 10},
    ])
    cat = pd.DataFrame({
        'Numero componenti': ['D', 'D'],
        'MODEL': ['MSV4', 'MSV4'],
        'Qta_comp_UMC': [4.0, 5.0],
    })
    ctx = _ctx(cat, sim)
    try:
        d = next(dd for sid, dd in zip(ctx.sku_ids, ctx.sku_data_list) if sid == 'D')
        rd, idx, _qpc = matching._compute_run_demands_for_sku(d, ctx)
        # pari specificità (solo Model) -> MAX(4,5)=5 -> 10 bici x5 = 50
        assert rd.tolist() == [50.0, 50.0]
    finally:
        ctx.cleanup()

import numpy as np
from tests._fixtures import tiny_simulation_output, tiny_sku_catalog, tiny_lead_times
import matching


def test_build_context_and_run_demands():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    ctx = matching._build_matching_context(cat, lt, sim, mapping, n_runs=4)
    try:
        # SKU1 (MODEL=MSV4) deve aggregare la domanda della config A: 10 per ogni run
        sku_data = next(d for sid, d in zip(ctx.sku_ids, ctx.sku_data_list) if sid == 'SKU1')
        run_demands, idxs = matching._compute_run_demands_for_sku(sku_data, ctx)
        assert run_demands.tolist() == [10.0, 10.0, 10.0, 10.0]
    finally:
        ctx.cleanup()


def test_match_skus_ultra_optimized_still_works():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    df, df_no_lt = matching.match_skus_ultra_optimized(cat, lt, sim, mapping, n_runs=4, percentile=99)
    skus = set(df['SKU'])
    assert {'SKU1', 'SKU2'} <= skus
    # SKU1 domanda costante 10 -> safety_stock_p99 == 0 (nessuna variabilità)
    row1 = df[df['SKU'] == 'SKU1'].iloc[0]
    assert abs(row1['safety_stock_p99']) < 1e-6

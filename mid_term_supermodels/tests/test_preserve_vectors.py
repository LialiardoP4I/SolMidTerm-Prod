from tests._fixtures import tiny_simulation_output, tiny_sku_catalog, tiny_lead_times
from mid_term_supermodels import matching


def test_preserve_run_vectors_returns_bike_demand():
    sim = tiny_simulation_output(n_runs=4)
    cat = tiny_sku_catalog()
    lt = tiny_lead_times()
    mapping = matching.build_column_mapping_auto(['MODEL'], list(sim.columns))

    vectors, df_no_lt = matching.match_skus_preserve_run_vectors(
        cat, lt, sim, mapping, n_runs=4, percentile=99
    )
    # SKU1 -> vettore bici [10,10,10,10], qty 1.0, lead_time 1.0, desc 'Comp 1'
    rd1, qty1, ltm1, desc1 = vectors['SKU1']
    assert rd1.tolist() == [10.0, 10.0, 10.0, 10.0]
    assert qty1 == 1.0
    assert ltm1 == 1.0
    assert desc1 == 'Comp 1'
    # SKU2 -> vettore già in PEZZI. Il catalogo fixture non ha Qta_comp_UMC ->
    # qty per-config = 1.0 -> pezzi = bici = [0,5,10,20]; qty nel tuple = 1.0.
    rd2, qty2, ltm2, desc2 = vectors['SKU2']
    assert rd2.tolist() == [0.0, 5.0, 10.0, 20.0]
    assert qty2 == 1.0
    assert desc2 == 'Comp 2'

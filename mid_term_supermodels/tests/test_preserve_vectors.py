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
    # SKU1 -> vettore bici [10,10,10,10], qty 1.0, lead_time 1.0
    rd1, qty1, ltm1 = vectors['SKU1']
    assert rd1.tolist() == [10.0, 10.0, 10.0, 10.0]
    assert qty1 == 1.0
    assert ltm1 == 1.0
    # SKU2 -> vettore bici [0,5,10,20], qty 2.0 (NON ancora moltiplicato)
    rd2, qty2, ltm2 = vectors['SKU2']
    assert rd2.tolist() == [0.0, 5.0, 10.0, 20.0]
    assert qty2 == 2.0

import numpy as np
from mid_term_supermodels import matching


def test_pooling_sumproduct_and_risk_pooling():
    # 2 supermodel, n_runs=4. Componente 'X' condiviso. Vettori GIÀ in PEZZI
    # (qty per-config applicata a monte); il pool non moltiplica più (qty=1.0).
    # SuperA pezzi [10,10,10,10] (varianza 0); SuperB pezzi [0,10,20,40]
    sm_results = {
        'SuperA': {'X': (np.array([10, 10, 10, 10], dtype=np.float32), 1.0, 1.0, 'desc')},
        'SuperB': {'X': (np.array([0, 10, 20, 40], dtype=np.float32), 1.0, 1.0, 'desc')},
    }
    df_pool, breakdown = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]

    # pezzi pooled per run = [10, 20, 30, 50]
    pooled = np.array([10, 20, 30, 50], dtype=np.float64)
    assert abs(row['mean_demand_pooled'] - pooled.mean()) < 1e-6
    expected_ss = max(0.0, np.percentile(pooled, 99) - pooled.mean())
    assert abs(row['safety_stock_p99_pooled'] - expected_ss) < 1e-6

    # breakdown per supermodel presente
    assert 'safety_stock_p99_standalone_SuperA' in row
    assert 'safety_stock_p99_standalone_SuperB' in row
    # SuperA pezzi costanti -> SS standalone 0
    assert abs(row['safety_stock_p99_standalone_SuperA']) < 1e-6

    # risk pooling: saving >= 0
    assert row['risk_pooling_saving'] >= -1e-9


def test_pooling_component_in_single_supermodel():
    # Componente 'Y' solo in SuperA
    sm_results = {
        'SuperA': {'Y': (np.array([1, 2, 3, 4], dtype=np.float32), 1.0, 1.0, 'desc')},
        'SuperB': {},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'Y'].iloc[0]
    pooled = np.array([1, 2, 3, 4], dtype=np.float64)
    assert abs(row['mean_demand_pooled'] - pooled.mean()) < 1e-6
    # niente contributo da SuperB
    assert row['mean_demand_SuperB'] == 0.0

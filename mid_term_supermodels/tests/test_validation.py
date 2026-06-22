import numpy as np
from mid_term_supermodels import matching


def test_single_supermodel_equals_standalone():
    # Con UN solo supermodel, SS pooled == SS standalone di quel supermodel.
    rd = np.array([0, 5, 10, 20], dtype=np.float32)
    sm_results = {'SuperA': {'X': (rd, 1.0, 1.0, 'desc')}}
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]
    assert abs(row['safety_stock_p99_pooled'] - row['safety_stock_p99_standalone_SuperA']) < 1e-6
    assert abs(row['risk_pooling_saving']) < 1e-6


def test_risk_pooling_reduces_or_equals_sum():
    # SS pooled <= somma SS standalone, per due flussi.
    sm_results = {
        'SuperA': {'X': (np.array([0, 5, 10, 20], dtype=np.float32), 1.0, 1.0, 'desc')},
        'SuperB': {'X': (np.array([20, 10, 5, 0], dtype=np.float32), 1.0, 1.0, 'desc')},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    row = df_pool[df_pool['SKU'] == 'X'].iloc[0]
    sum_standalone = (row['safety_stock_p99_standalone_SuperA']
                      + row['safety_stock_p99_standalone_SuperB'])
    assert row['safety_stock_p99_pooled'] <= sum_standalone + 1e-9
    assert row['risk_pooling_saving'] >= -1e-9

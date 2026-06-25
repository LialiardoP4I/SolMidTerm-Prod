# -*- coding: utf-8 -*-
"""Test per save_results_pooled (Task 7)."""

import numpy as np
import pandas as pd
import openpyxl
from mid_term_supermodels import matching


def test_save_results_pooled_writes_columns(tmp_path):
    """Verifica che save_results_pooled scrive colonne attese su Excel."""
    sm_results = {
        'SuperA': {'X': (np.array([10, 10, 10, 10], dtype=np.float32), 1.0, 1.0, 'desc')},
        'SuperB': {'X': (np.array([0, 5, 10, 20], dtype=np.float32), 2.0, 1.0, 'desc')},
    }
    df_pool, _ = matching.pool_safety_stock(sm_results, percentile=99, n_runs=4)
    out = tmp_path / 'pooled.xlsx'
    matching.save_results_pooled(df_pool, str(out), percentile=99)

    wb = openpyxl.load_workbook(out)
    hdr = [c.value for c in wb.active[1]]
    assert 'SKU' in hdr
    assert 'Description' in hdr
    assert 'safety_stock_p99_pooled' in hdr
    assert 'risk_pooling_saving' in hdr
    assert 'mean_demand_SuperA' in hdr

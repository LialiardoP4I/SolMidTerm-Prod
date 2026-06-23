import os

import numpy as np

from mid_term_supermodels import multi_supermodel
from mid_term_supermodels.config import SimulationConfig


def test_run_multi_supermodel_rolling_month_aligned_pooling(tmp_path, monkeypatch):
    """Orchestratore rolling + pooling per mese di lancio (tutto monkeypatchato).

    - SuperA ha due mesi di lancio ('JAN 2027', 'FEB 2027');
    - SuperB ha solo 'JAN 2027' con un vettore diverso.
    Verifica: colonna mese_lancio con entrambi i mesi; per 'JAN 2027' (entrambi
    presenti) il pooling ha le colonne dei due supermodel e risk_pooling_saving>=0;
    per 'FEB 2027' (solo SuperA) il pooled coincide con lo standalone di SuperA.
    """
    # Due cartelle supermodel fittizie (i path non vengono mai realmente letti).
    fake_a = tmp_path / 'SuperA'
    fake_b = tmp_path / 'SuperB'
    fake_a.mkdir()
    fake_b.mkdir()

    monkeypatch.setattr(
        multi_supermodel, 'discover_supermodels',
        lambda input_dir: [str(fake_a), str(fake_b)],
    )

    # Staging: niente copia reale, ritorna una dir tmp che esiste davvero
    # (cosi' os.chdir + rmtree funzionano senza toccare dati reali).
    def fake_stage(sm_dir, shared_dir, work_root):
        work_sm = tmp_path / '_work' / os.path.basename(str(sm_dir))
        (work_sm / 'Input').mkdir(parents=True, exist_ok=True)
        return work_sm
    monkeypatch.setattr(multi_supermodel, 'stage_supermodel_input', fake_stage)

    # Forecast + sim_config: nessun file reale necessario.
    monkeypatch.setattr(
        multi_supermodel, 'load_monthly_forecast',
        lambda path: ([1.0, 2.0], ['JAN 2027', 'FEB 2027']),
        raising=False,
    )
    monkeypatch.setattr(
        multi_supermodel, '_build_sim_config_for_supermodel',
        lambda json_path, month_names, k, seed_base:
            SimulationConfig(n_runs=4, random_seed=seed_base + k,
                             percentile_safety_stock=99),
    )

    # Rolling per-supermodel: vettori per-run noti, deterministici.
    a_jan = (np.array([10, 10, 10, 10], dtype=float), 1.0, 1.0, 'x')
    a_feb = (np.array([5, 5, 5, 5], dtype=float), 1.0, 1.0, 'x')
    b_jan = (np.array([0, 5, 10, 20], dtype=float), 2.0, 1.0, 'x')

    def fake_collect(config, sim_config, max_iterations=None):
        # sm_name distinto via output_dir (Output/<SM>).
        sm_name = os.path.basename(str(config.output_dir))
        if sm_name == 'SuperA':
            return {
                'JAN 2027': {'vectors': {'X': a_jan}, 'troncata': False,
                             'offset': 0, 'mesi_simulati': 2},
                'FEB 2027': {'vectors': {'X': a_feb}, 'troncata': True,
                             'offset': 1, 'mesi_simulati': 1},
            }
        return {
            'JAN 2027': {'vectors': {'X': b_jan}, 'troncata': False,
                         'offset': 0, 'mesi_simulati': 2},
        }
    monkeypatch.setattr(
        multi_supermodel, 'collect_run_vectors_rolling', fake_collect,
        raising=False,
    )

    out_dir = tmp_path / 'Output'
    result = multi_supermodel.run_multi_supermodel_rolling(
        input_dir=str(tmp_path), output_dir=str(out_dir),
        json_path='run_config.json', seed_base=42,
    )

    df = result.per_sku_pooled
    assert 'mese_lancio' in df.columns
    assert set(df['mese_lancio']) == {'JAN 2027', 'FEB 2027'}

    # JAN: entrambi i supermodel -> colonne per A e B, risk_pooling_saving >= 0.
    jan = df[df['mese_lancio'] == 'JAN 2027']
    row_jan = jan[jan['SKU'] == 'X'].iloc[0]
    assert 'safety_stock_p99_standalone_SuperA' in row_jan
    assert 'safety_stock_p99_standalone_SuperB' in row_jan
    assert row_jan['risk_pooling_saving'] >= -1e-9
    # pooled per run JAN = A[10,10,10,10] + B[0,10,20,40] = [10,20,30,50]
    pooled = np.array([10, 20, 30, 50], dtype=float)
    expected_ss = max(0.0, np.percentile(pooled, 99) - pooled.mean())
    assert abs(row_jan['safety_stock_p99_pooled'] - expected_ss) < 1e-6

    # FEB: solo SuperA -> pooled == standalone SuperA.
    feb = df[df['mese_lancio'] == 'FEB 2027']
    row_feb = feb[feb['SKU'] == 'X'].iloc[0]
    assert abs(row_feb['safety_stock_p99_pooled']
               - row_feb['safety_stock_p99_standalone_SuperA']) < 1e-6
    assert abs(row_feb['mean_demand_pooled']
               - row_feb['mean_demand_SuperA']) < 1e-6
    assert bool(row_feb['finestra_troncata']) is True

    # Output Excel scritto.
    assert 'pooled_rolling' in result.output_paths
    assert os.path.exists(result.output_paths['pooled_rolling'])

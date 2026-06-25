import numpy as np
import pandas as pd


def tiny_simulation_output(n_runs=4):
    """2 configurazioni, 1 mese, n_runs run. Colonne: MODEL + <Mese>_Run_<i>."""
    rows = []
    # config A: MODEL=MSV4  -> domanda costante 10 per run
    rowA = {'MODEL': 'MSV4'}
    # config B: MODEL=MSV4S -> domanda variabile
    rowB = {'MODEL': 'MSV4S'}
    demA = [10, 10, 10, 10][:n_runs]
    demB = [0, 5, 10, 20][:n_runs]
    for i in range(n_runs):
        rowA[f'JAN 2026_Run_{i}'] = demA[i]
        rowB[f'JAN 2026_Run_{i}'] = demB[i]
    return pd.DataFrame([rowA, rowB])


def tiny_sku_catalog():
    """SKU SKU1 usato da MSV4, SKU2 usato da MSV4S."""
    return pd.DataFrame({
        'Numero componenti': ['SKU1', 'SKU2'],
        'MODEL': ['MSV4', 'MSV4S'],
    })


def tiny_lead_times():
    return pd.DataFrame({
        'SKU': ['SKU1', 'SKU2'],
        'Lead_Time_Months': [1.0, 1.0],
        'Lead_Time_Months_Ceil': [1.0, 1.0],
        'Quantity_Per_Bike': [1.0, 2.0],
        'Description': ['Comp 1', 'Comp 2'],
    })

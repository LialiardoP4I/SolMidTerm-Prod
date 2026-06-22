import pandas as pd
import numpy as np


def _desc_from_bom(df_bom: pd.DataFrame, skus: list) -> pd.Series:
    """Replica isolata della nuova logica Description (per testarla senza l'intera pipeline)."""
    bom_desc_map = (df_bom[['Numero componenti', 'Testo breve oggetto']]
                    .dropna(subset=['Testo breve oggetto'])
                    .assign(_k=df_bom['Numero componenti'].astype(str))
                    .drop_duplicates(subset='_k', keep='first')
                    .set_index('_k')['Testo breve oggetto'])
    return pd.Series(skus, index=skus).astype(str).map(bom_desc_map).fillna('Non disponibile')


def test_description_taken_from_bom_when_present():
    df_bom = pd.DataFrame({
        'Numero componenti': ['586P3481AB', '586P3481AB', '99999XX'],
        'Testo breve oggetto': ['STGR SERB CARBURANTE GIALLO V24A', None, 'VITE M4'],
    })
    out = _desc_from_bom(df_bom, ['586P3481AB', '99999XX'])
    assert out['586P3481AB'] == 'STGR SERB CARBURANTE GIALLO V24A'
    assert out['99999XX'] == 'VITE M4'


def test_description_fallback_when_missing_in_bom():
    df_bom = pd.DataFrame({
        'Numero componenti': ['111AA'],
        'Testo breve oggetto': ['DADO'],
    })
    out = _desc_from_bom(df_bom, ['222BB'])  # SKU assente in BOM
    assert out['222BB'] == 'Non disponibile'

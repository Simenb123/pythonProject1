# -*- coding: utf-8 -*-
# src/app/services/ar_adapter.py
from __future__ import annotations
import pandas as pd

def fetch_shareholders(orgnr: str) -> pd.DataFrame:
    """
    Returner DataFrame med aksjonærer for gitt orgnr (kobles til din AR-modul).
    Forventede kolonner (fleksibelt): ["eier_navn","eier_orgnr","andel_prosent","eier_type","kilde_dato", ...]
    """
    # TODO: Koble til din AR-modul. Demo: tom DF
    return pd.DataFrame(columns=["eier_navn","eier_orgnr","andel_prosent","kilde_dato"])

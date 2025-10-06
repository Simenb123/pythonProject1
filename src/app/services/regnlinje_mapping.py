# -*- coding: utf-8 -*-
# src/app/services/regnlinje_mapping.py
from __future__ import annotations
from pathlib import Path
from typing import Tuple, Dict
import pandas as pd
import numpy as np

def _norm(s: str) -> str:
    return (
        str(s).strip().lower()
        .replace("\u00A0", " ")
        .replace("-", " ")
        .replace(".", "")
        .replace("_", " ")
    )

def _pick(df: pd.DataFrame, *cands: str) -> str | None:
    low = {_norm(c): c for c in df.columns}
    for name in cands:
        n = _norm(name)
        # eksakt
        if n in low:
            return low[n]
        # inneholder
        for k, v in low.items():
            if n == k or n in k:
                return v
    return None

def _read_intervals(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    start = _pick(df, "StartKonto", "fra", "fra konto", "start")
    end   = _pick(df, "SluttKonto", "til", "til konto", "slutt")
    regn  = _pick(df, "Regnnr.", "regnnr", "nr", "linjenr", "regnskapsnr")
    name  = _pick(df, "Regnskapslinje", "regnskapslinje", "linjenavn")
    if not (start and end and regn):
        raise ValueError("Fant ikke kolonnene for start/slutt/regnnr i mapping-filen.")
    out = pd.DataFrame({
        "start": pd.to_numeric(df[start], errors="coerce").astype("Int64"),
        "end":   pd.to_numeric(df[end], errors="coerce").astype("Int64"),
        "regnnr": pd.to_numeric(df[regn], errors="coerce").astype("Int64"),
    })
    if name:
        out["name_hint"] = df[name].astype(str)
    out = out.dropna(subset=["start","end","regnnr"]).astype({"start":"int64","end":"int64","regnnr":"int64"})
    return out

def _read_lines(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    nr = _pick(df, "nr.", "nr", "regnnr", "linjenr")
    txt = _pick(df, "Regnskapslinje", "linjenavn", "navn")
    if not (nr and txt):
        raise ValueError("Fant ikke kolonnene 'nr'/'Regnskapslinje' i regnskapslinjer-filen.")
    out = pd.DataFrame({"regnnr": pd.to_numeric(df[nr], errors="coerce").astype("Int64"),
                        "regnskapslinje": df[txt].astype(str)})
    out = out.dropna(subset=["regnnr"]).astype({"regnnr":"int64"})
    # fjern duplikate nr (ta første)
    out = out.drop_duplicates(subset=["regnnr"], keep="first")
    return out

def _expand_intervals(iv_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in iv_df.iterrows():
        s, e, nr = int(r["start"]), int(r["end"]), int(r["regnnr"])
        if e < s: s, e = e, s
        # begrens absurd store intervaller
        if e - s > 20000:
            raise ValueError(f"Uvanlig stort intervall i mapping: {s}–{e}")
        ks = np.arange(s, e+1, dtype=np.int64)
        tmp = pd.DataFrame({"konto": ks, "regnnr": nr})
        rows.append(tmp)
    if not rows:
        return pd.DataFrame(columns=["konto","regnnr"])
    out = pd.concat(rows, ignore_index=True)
    # Hvis duplikat konto i flere intervaller → behold første
    out = out.drop_duplicates(subset=["konto"], keep="first")
    return out

def attach_regnskapslinjer(df_sb: pd.DataFrame, mapping_xlsx: Path, lines_xlsx: Path) -> Tuple[pd.DataFrame, Dict[str, int]]:
    """
    Returnerer (df_med_kolonner, info):
      - df_med_kolonner har ekstra kolonner 'regnnr' og 'regnskapslinje' (hvis navneliste ble funnet)
      - info = {"mapped_accounts": X, "total_accounts": Y}
    """
    if "konto" not in df_sb.columns:
        raise ValueError("DataFrame mangler kolonnen 'konto'.")
    df = df_sb.copy()
    df["konto"] = pd.to_numeric(df["konto"], errors="coerce").astype("Int64")

    iv = _read_intervals(Path(mapping_xlsx))
    lut = _expand_intervals(iv)
    names = _read_lines(Path(lines_xlsx))

    out = df.merge(lut, how="left", left_on="konto", right_on="konto")
    out = out.merge(names, how="left", on="regnnr")

    total = int(out["konto"].dropna().nunique())
    mapped = int(out.loc[out["regnnr"].notna(), "konto"].nunique())
    return out, {"mapped_accounts": mapped, "total_accounts": total}

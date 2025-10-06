# import_pipeline.py – 2025-05-27
# -------------------------------------------------------------
# • Leser CSV/Excel robust (auto-encoding, auto-delimiter)
# • Renser desimaltall (komma → punktum, fjerner tusenskilletegn)
# • Konverterer til Parquet (“standard.parquet”) i klientroten
# -------------------------------------------------------------
from __future__ import annotations

from pathlib import Path
from typing import Mapping, Tuple, Dict, Any

import chardet
import csv
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

PARQUET_NAME = "standard.parquet"

_STD_COLS: Dict[str, str] = {
    "konto": "int32",
    "beløp": "float64",
    "dato": "datetime64[ns]",
    "bilagsnr": "string[pyarrow]",
}

_META_KEYS = {"encoding", "std_file"}  # nøkler som ikke er kolonner

# -----------------------------------------------------------------
# 1  CSV-leser med encoding– og delimiter-fallback
# -----------------------------------------------------------------
def _les_csv(p: Path) -> Tuple[pd.DataFrame, str]:
    raw = p.read_bytes()
    enc_guess = chardet.detect(raw)["encoding"]

    encodings = [
        e
        for e in [
            enc_guess,
            "utf-8-sig",
            "utf-8",
            "utf-16",
            "utf-16-le",
            "utf-16-be",
            "cp1252",
            "latin1",
        ]
        if e
    ]
    # fjern duplikater (case-uavhengig)
    seen = set()
    encodings = [e for e in encodings if not (e.lower() in seen or seen.add(e.lower()))]

    # direkte forsøk (sniff sep=None)
    for enc in encodings:
        try:
            return pd.read_csv(p, sep=None, engine="python", encoding=enc), enc
        except Exception:
            pass

    # tyd ut skilletegn fra sample
    try:
        sample = raw[:50000].decode("latin1", errors="ignore")
        delim = csv.Sniffer().sniff(sample, delimiters=";,|\t").delimiter
    except Exception:
        delim = ";"

    for enc in encodings:
        for sep in (delim, ";", ",", "|", "\t"):
            try:
                return pd.read_csv(p, sep=sep, encoding=enc), enc
            except Exception:
                pass

    raise ValueError("Kunne ikke lese CSV – skilletegn/encoding uklart.")


# -----------------------------------------------------------------
# 2  Les råfil (respekter evt. foretrukket encoding for CSV)
# -----------------------------------------------------------------
def _les_råfil(src: Path, encoding: str | None) -> pd.DataFrame:
    suf = src.suffix.lower()
    if suf in (".xlsx", ".xls"):
        return pd.read_excel(src, engine="openpyxl")

    if suf == ".csv":
        if encoding:
            return pd.read_csv(src, sep=None, engine="python", encoding=encoding)
        df, _ = _les_csv(src)
        return df

    raise ValueError("Kun .xlsx / .xls / .csv støttes")


# -----------------------------------------------------------------
# 3  Standardiser kolonnenavn og datatyper
# -----------------------------------------------------------------
def _standardiser(df_raw: pd.DataFrame, mapping: Mapping[str, str]) -> pd.DataFrame:
    # a) gi standard-kolonnene sine navn
    rename_map = {
        mapping[k]: k
        for k in mapping
        if k not in _META_KEYS and mapping[k] in df_raw.columns
    }
    df = df_raw.rename(columns=rename_map)

    # b) sikre at alle _STD_COLS finnes
    for col in _STD_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[list(_STD_COLS)]  # rekkefølge

    # ---------- verdirensing ----------------------------------
    # konto
    df["konto"] = (
        df["konto"]
        .astype("string")
        .str.replace(r"\D", "", regex=True)
        .astype("Int32")
    )

    # beløp  (fjerner tusenskilletegn « » NBSP, thin space, punktum)
    bel = (
        df["beløp"]
        .astype("string")
        .str.replace(r"[ \u00A0\u202F\u2009]", "", regex=True)  # blanke/thin space
        .pipe(lambda s: s.str.replace(".", "", regex=False) if s.str.contains(",").any() else s)
        .str.replace(",", ".", regex=False)
    )
    df["beløp"] = pd.to_numeric(bel, errors="coerce")

    # dato
    df["dato"] = pd.to_datetime(df["dato"], errors="coerce", dayfirst=True)

    return df


# -----------------------------------------------------------------
# 4  Lagre Parquet
# -----------------------------------------------------------------
def _lagre_parquet(df_std: pd.DataFrame, dst_dir: Path) -> Path:
    dst = dst_dir / PARQUET_NAME
    pq.write_table(pa.Table.from_pandas(df_std, preserve_index=False), dst)
    return dst


# -----------------------------------------------------------------
# 5  «Public API»
# -----------------------------------------------------------------
def konverter_til_parquet(
    kildefil: Path,
    dst_dir: Path,
    mapping: Mapping[str, str],
    encoding: str | None = None,
) -> Path:
    """Les *kildefil* → rens → lagre `standard.parquet` til *dst_dir*."""
    df_raw = _les_råfil(kildefil, encoding)
    df_std = _standardiser(df_raw, mapping)
    return _lagre_parquet(df_std, dst_dir)

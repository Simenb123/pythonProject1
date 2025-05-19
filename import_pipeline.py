# import_pipeline.py – med støtte for komma-desimal
from __future__ import annotations
from pathlib import Path
from typing import Mapping
import chardet, csv, pandas as pd, pyarrow as pa, pyarrow.parquet as pq

PARQUET_NAME = "standard.parquet"

_STD_COLS = {
    "konto":    "int32",
    "beløp":    "float64",
    "dato":     "datetime64[ns]",
    "bilagsnr": "string[pyarrow]",
}

# ───────────────────── robust CSV-leser (uendret) ────────────────────
def _les_csv(p: Path) -> tuple[pd.DataFrame, str]:
    raw = p.read_bytes()
    enc_guess = chardet.detect(raw)["encoding"]
    encodings = [e for e in
                 [enc_guess, "utf-8-sig", "utf-8",
                  "utf-16", "utf-16-le", "utf-16-be",
                  "cp1252", "latin1"] if e]
    seen=set(); encodings=[e for e in encodings if not (e.lower() in seen or seen.add(e.lower()))]
    for enc in encodings:
        try:
            return pd.read_csv(p, sep=None, engine="python", encoding=enc), enc
        except Exception: pass
    sample = raw[:50000].decode("latin1", errors="ignore")
    delim = csv.Sniffer().sniff(sample, delimiters=";,|\t").delimiter
    for enc in encodings:
        for sep in (delim, ";", ",", "|", "\t"):
            try:
                return pd.read_csv(p, sep=sep, encoding=enc), enc
            except Exception: pass
    raise ValueError("Kunne ikke lese CSV – skilletegn/encoding uklart.")

# ───────────────────── helpers ───────────────────────────────────────
def _les_råfil(src: Path, encoding: str | None) -> pd.DataFrame:
    suf = src.suffix.lower()
    if suf in (".xlsx", ".xls"):
        return pd.read_excel(src, engine="openpyxl")
    if suf == ".csv":
        df, _ = _les_csv(src)
        return df
    raise ValueError("Kun .xlsx / .xls / .csv støttes")

def _standardiser(df_raw: pd.DataFrame, mapping: Mapping[str, str]) -> pd.DataFrame:
    df = df_raw.rename(columns={mapping[k]: k for k in mapping})

    for col in _STD_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[list(_STD_COLS)]

    # konto
    df["konto"] = (
        df["konto"].astype("string")
        .str.replace(r"\D", "", regex=True)
        .astype("Int32")
    )
    # beløp – normaliser komma → punktum, fjern mellomrom/nbsp
    df["beløp"] = (
        df["beløp"].astype("string")
        .str.replace("\u00A0", "", regex=False)
        .str.replace(" ", "",  regex=False)
        .str.replace(",", ".", regex=False)
        .astype("float64", errors="ignore")
    )
    # dato
    df["dato"] = pd.to_datetime(df["dato"], errors="coerce", dayfirst=True)

    return df

def _lagre_parquet(df_std: pd.DataFrame, dst_dir: Path) -> Path:
    dst = dst_dir / PARQUET_NAME
    pq.write_table(pa.Table.from_pandas(df_std, preserve_index=False),
                   dst, version="2.6")
    return dst

def konverter_til_parquet(
    kildefil: Path,
    dst_dir: Path,
    mapping: Mapping[str, str],
    encoding: str | None = None,
) -> Path:
    df_raw = _les_råfil(kildefil, encoding)
    df_std = _standardiser(df_raw, mapping)
    return _lagre_parquet(df_std, dst_dir)

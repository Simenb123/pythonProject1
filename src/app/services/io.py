# src/app/services/io.py
# -----------------------------------------------------------------------------
# Robust fil-IO for CSV/XLSX:
#  - Sniffer encoding og delimiter
#  - Leser store CSV-er forutsigbart
#  - Normaliserer beløp (tusen-/desimalskilletegn)
#  - Forsøker å parse dato-kolonner
#  - Standardiserer kolonnetitler via mapping/synonymer
#  - Preview/paginering og eksport til Excel
#  - (Valgfritt) Parquet-cache med pyarrow
# -----------------------------------------------------------------------------
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Optional

import csv
import re

import chardet
import numpy as np  # type: ignore
import pandas as pd

# --- valgfrie avhengigheter ---------------------------------------------------
try:
    import openpyxl  # noqa: F401
    _HAS_OPENPYXL = True
except Exception:
    _HAS_OPENPYXL = False

try:
    import pyarrow as pa  # noqa: F401
    import pyarrow.parquet as pq  # noqa: F401
    _HAS_PYARROW = True
except Exception:
    _HAS_PYARROW = False


# ------------------------------- Datatyper ------------------------------------

@dataclass
class ReadInfo:
    """Metadata om hvordan en fil ble/kan leses."""
    path: Path
    encoding: str
    sep: Optional[str]    # kan være None (la pandas sniffe)
    rows: Optional[int]
    cols: Optional[int]
    source_type: str      # "csv" | "xlsx" | "parquet" | "unknown"


__all__ = [
    "ReadInfo",
    "sniff_csv",
    "read_csv_robust",
    "read_raw",
    "standardize",
    "detect_schema",
    "preview",
    "save_excel",
    "to_parquet",
    "convert_csv_to_parquet",
]


# ---------------------------- CSV-sniff & lesing ------------------------------

def sniff_csv(path: Path, sample_bytes: int = 200_000) -> tuple[str, Optional[str]]:
    """
    Sniff encoding og delimiter for CSV.
    Returnerer (encoding, sep). 'sep' kan være None (la pandas sniffe).
    """
    raw = path.read_bytes()
    head = raw[:sample_bytes]
    enc = "utf-8"
    try:
        enc = chardet.detect(head).get("encoding") or "utf-8"
    except Exception:
        pass

    # csv.Sniffer kan feile – fang og la sep bli None
    try:
        text = head.decode(enc, errors="replace")
        dialect = csv.Sniffer().sniff(text, delimiters=";,|\t,")
        sep = dialect.delimiter
    except Exception:
        sep = None

    return enc, sep


def read_csv_robust(
    path: Path,
    dtype: Optional[Mapping[str, Any]] = None,
    low_memory: bool = False,  # bevart parameter for bakoverkompatibilitet (IGNORERES med python-engine)
) -> tuple[pd.DataFrame, ReadInfo]:
    """
    Les CSV robust:
      - Sniffer encoding/delimiter
      - Faller tilbake til latin-1 ved uventede feil

    NB: Vi bruker pandas' python-engine når sep=None eller ukjente separators.
        Python-engine støtter **ikke** low_memory. Derfor sender vi ikke dette flagget.
    """
    enc, sep = sniff_csv(path)
    try:
        # Viktig: IKKE send low_memory til python-engine
        df = pd.read_csv(path, encoding=enc, sep=sep, engine="python", dtype=dtype)
    except Exception:
        # fallback: prøv latin-1 uten sep (la pandas sniffe)
        df = pd.read_csv(path, encoding="latin-1", engine="python", dtype=dtype)
        enc, sep = "latin-1", None

    info = ReadInfo(path=path, encoding=enc, sep=sep,
                    rows=len(df), cols=len(df.columns), source_type="csv")
    return df, info


def read_raw(path: Path) -> tuple[pd.DataFrame, ReadInfo]:
    """
    Les råfil (CSV eller XLSX). Foretrekker CSV som kanonisk format.
    """
    suf = path.suffix.lower()
    if suf == ".csv":
        return read_csv_robust(path)

    if suf in (".xlsx", ".xls"):
        if not _HAS_OPENPYXL:
            raise RuntimeError("openpyxl mangler – installer 'openpyxl' for å lese Excel-filer.")
        df = pd.read_excel(path, engine="openpyxl")
        info = ReadInfo(path=path, encoding="binary", sep=None,
                        rows=len(df), cols=len(df.columns), source_type="xlsx")
        return df, info

    if suf == ".parquet":
        if not _HAS_PYARROW:
            raise RuntimeError("pyarrow mangler – installer 'pyarrow' for å lese Parquet.")
        df = pd.read_parquet(path)
        info = ReadInfo(path=path, encoding="binary", sep=None,
                        rows=len(df), cols=len(df.columns), source_type="parquet")
        return df, info

    raise ValueError(f"Ukjent/ikke støttet filtype: {path.suffix}")


# --------------------------- Normalisering helpers ----------------------------

_NON_BREAKING_SPACE = "\u00A0"

def _normalize_numeric_series(s: pd.Series) -> pd.Series:
    """
    Gjør om beløp representert som strenger til tall:
      - fjerner NBSP/space
      - fjerner tusenskilletegn '.'
      - bytter ',' -> '.' for desimal
      - safer til float med errors='coerce'
    """
    if pd.api.types.is_numeric_dtype(s):
        return s

    out = (
        s.astype(str)
         .str.replace(_NON_BREAKING_SPACE, "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace(".", "", regex=False)      # tusenskille
         .str.replace(",", ".", regex=False)     # desimal
    )
    return pd.to_numeric(out, errors="coerce")


_DATE_RE = re.compile(r"\b(dato|date|trans|bilagsdato|post_date)\b", re.IGNORECASE)


def _maybe_parse_dates(df: pd.DataFrame, candidate_cols: Iterable[str]) -> None:
    """
    Prøv å parse dato-kolonner in-place (tåler tom/feil verdier).
    """
    for c in candidate_cols:
        if c in df.columns:
            try:
                df[c] = pd.to_datetime(df[c], errors="coerce", utc=False, dayfirst=True)
            except Exception:
                pass


# ------------------------------ Standardisering -------------------------------

_DEFAULT_SYNONYMS: dict[str, tuple[str, ...]] = {
    "konto": ("konto", "kontonr", "kontonummer", "account", "acct"),
    "beløp": ("beløp", "belop", "amount", "sum", "total", "amount_nok", "amount_sek"),
    "dato": ("dato", "date", "bilagsdato", "transdate", "post_date"),
    "bilagsnr": ("bilagsnr", "bilagsnummer", "bilagnr", "voucher", "docno", "bilag"),
    "kontonavn": ("kontonavn", "accountname", "kontonavntekst", "navn"),
}

def _normalize_colname(s: str) -> str:
    return (
        s.strip()
         .lower()
         .replace(_NON_BREAKING_SPACE, " ")
         .replace("-", " ")
         .replace(".", " ")
    )


def standardize(
    df: pd.DataFrame,
    mapping: Optional[Mapping[str, str]] = None,
    synonyms: Optional[Mapping[str, Iterable[str]]] = None,
    parse_dates: bool = True,
    numeric_fields: Iterable[str] = ("beløp",),
) -> tuple[pd.DataFrame, dict[str, str]]:
    """
    Standardiser kolonner til kanoniske navn vha. mapping/synonymer.
    Returnerer (df_kopi, mapping_brukt).

    mapping: f.eks. {"beløp": "Belop", "konto": "Konto", ...}
    synonyms: f.eks. {"beløp": ("beløp","belop","amount",...)}
    """
    syn = {k: tuple(v) for k, v in (synonyms or _DEFAULT_SYNONYMS).items()}
    used_map: dict[str, str] = {}
    col_lut = {c: _normalize_colname(c) for c in df.columns}

    # 1) Eksplisitt mapping har førsteprioritet
    if mapping:
        for std, col in mapping.items():
            if col in df.columns:
                used_map[std] = col

    # 2) Fyll inn via synonymer
    missing_std = [k for k in syn.keys() if k not in used_map]
    for std in missing_std:
        for cand in syn[std]:
            # eksakt match på normalisert navn
            for orig, norm in col_lut.items():
                if norm == _normalize_colname(cand):
                    used_map[std] = orig
                    break
            if std in used_map:
                break

    # 3) Bygg rename-kart og lag kopi
    rename_map = {}
    for std, col in used_map.items():
        if col in df.columns and col != std:
            rename_map[col] = std

    out = df.copy()
    if rename_map:
        out = out.rename(columns=rename_map)

    # 4) Dato og beløp-normalisering
    if parse_dates:
        date_candidates = [c for c in out.columns if c == "dato" or _DATE_RE.search(c)]
        _maybe_parse_dates(out, date_candidates)

    for fld in numeric_fields:
        if fld in out.columns:
            out[fld] = _normalize_numeric_series(out[fld])

    return out, {k: v for k, v in used_map.items()}


# --------------------------- Schema & utility-funksjoner -----------------------

def detect_schema(df: pd.DataFrame) -> dict[str, str]:
    """
    En enkel schema-detektor (dtype-strenger) – nyttig for cache/inspeksjon.
    """
    schema = {}
    for c in df.columns:
        dt = df[c].dtype
        if pd.api.types.is_datetime64_any_dtype(dt):
            schema[c] = "datetime"
        elif pd.api.types.is_integer_dtype(dt):
            schema[c] = "int"
        elif pd.api.types.is_float_dtype(dt):
            schema[c] = "float"
        elif pd.api.types.is_bool_dtype(dt):
            schema[c] = "bool"
        else:
            schema[c] = "string"
    return schema


def preview(df: pd.DataFrame, limit: int = 200, offset: int = 0) -> pd.DataFrame:
    """
    Returner et paginert utsnitt (for Treeview/preview i GUI).
    """
    offset = max(0, int(offset))
    limit = max(1, int(limit))
    return df.iloc[offset: offset + limit].copy()


def save_excel(df: pd.DataFrame, path: Path, sheet_name: str = "Uttrekk") -> Path:
    """
    Skriv DataFrame til .xlsx (krever openpyxl). Returnerer faktisk lagret sti.
    """
    if not _HAS_OPENPYXL:
        raise RuntimeError("openpyxl mangler – installer 'openpyxl' for Excel-eksport.")
    path = path.with_suffix(".xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name)
    return path


# ------------------------------- Parquet-cache --------------------------------

def to_parquet(df: pd.DataFrame, dst: Path) -> Path:
    """
    Skriv DataFrame til Parquet (uten index). Returnerer lagret sti.
    """
    if not _HAS_PYARROW:
        raise RuntimeError("pyarrow mangler – installer 'pyarrow' for Parquet.")
    from pyarrow import Table  # late import for tydelig feilhåndtering
    from pyarrow import parquet as _pq

    dst = dst.with_suffix(".parquet")
    table = Table.from_pandas(df, preserve_index=False)
    _pq.write_table(table, dst)
    return dst


def convert_csv_to_parquet(
    src_csv: Path,
    dst_dir: Path,
    mapping: Mapping[str, str] | None = None,
) -> Path:
    """
    Les CSV (robust) → standardiser (hvis mapping gitt) → lagre Parquet i dst_dir.
    Returnerer stien til Parquet-filen.
    """
    if src_csv.suffix.lower() != ".csv":
        raise ValueError("convert_csv_to_parquet forventer en .csv-kildefil")

    df, _ = read_csv_robust(src_csv)
    if mapping:
        df, _ = standardize(df, mapping=mapping)

    dst_dir.mkdir(parents=True, exist_ok=True)
    return to_parquet(df, dst_dir / src_csv.stem)

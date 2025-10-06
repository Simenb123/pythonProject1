# -*- coding: utf-8 -*-
# src/app/services/regnskapslinjer.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
import os
import re

import numpy as np
import pandas as pd

# Vi bruker same settings-mekanisme som ellers i appen
try:
    from app.services.clients import load_settings  # type: ignore
except Exception:  # pragma: no cover
    from services.clients import load_settings  # type: ignore


# -------------------------- Lokasjon av kildefiler --------------------------

_DEFAULT_DIRS = [
    Path(r"F:\Dokument\Kildefiler"),
]

_ENV_KEYS = ("KILDEFILER_DIR", "REGNSKAPSLINJER_DIR", "KILDEFILES_DIR")
_SETTING_KEYS = ("kildefiler_dir", "regnskapslinjer_dir", "mapping_dir")

REGNSKAPSLINJER_NAME = "Regnskapslinjer.xlsx"
MAPPING_NAME = "Mapping standard kontoplan.xlsx"


def _first_existing(paths: list[Path]) -> Optional[Path]:
    for p in paths:
        if p and Path(p).exists():
            return Path(p)
    return None


def find_kildefiler_dir() -> Optional[Path]:
    # settings.json
    try:
        st = load_settings() or {}
        cands = [Path(st[k]) for k in _SETTING_KEYS if isinstance(st.get(k), str)]
        found = _first_existing(cands)
        if found:
            return found
    except Exception:
        pass

    # miljøvariabler
    for k in _ENV_KEYS:
        v = os.getenv(k)
        if v and Path(v).exists():
            return Path(v)

    # default
    return _first_existing(_DEFAULT_DIRS)


def _find_file_case_insensitive(base: Path, wanted: str) -> Optional[Path]:
    wanted_low = wanted.lower()
    # eksakt navn først
    p = base / wanted
    if p.exists():
        return p
    # ellers: søk case-insensitivt
    for f in base.iterdir():
        if f.is_file() and f.name.lower() == wanted_low:
            return f
    # fallback: "inneholder"
    for f in base.iterdir():
        if f.is_file() and wanted_low in f.name.lower():
            return f
    return None


# -------------------------- Leser Excel-tabeller --------------------------

def _norm(s: str) -> str:
    return (
        str(s).strip().lower()
        .replace("\u00A0", " ")
        .replace("_", " ")
        .replace("-", " ")
    )


def _pick_col(df: pd.DataFrame, *choices: str) -> Optional[str]:
    low = {_norm(c): c for c in df.columns}
    for ch in choices:
        if ch in low:
            return low[ch]
    # substring match (robust mot «regnskapslinje nr» vs «regnskapsnr»)
    for ch in choices:
        for k, v in low.items():
            if ch in k:
                return v
    return None


def load_regnskapslinjer(path: Path) -> pd.DataFrame:
    """
    Leser 'Regnskapslinjer.xlsx' → DataFrame med kolonnene:
      - 'regnskapsnr' (string)
      - 'regnskapsnavn' (string)
    """
    df = pd.read_excel(path, engine="openpyxl")
    # finn kolonner robust
    c_nr = (
        _pick_col(df, "regnskapsnr", "regnskapslinjenr", "linjenr", "nr", "regnskapslinje nr")
        or df.columns[0]
    )
    c_navn = (
        _pick_col(df, "regnskapsnavn", "regnskapslinje", "linjenavn", "navn")
        or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
    )
    out = pd.DataFrame()
    out["regnskapsnr"] = df[c_nr].astype(str).str.strip()
    out["regnskapsnavn"] = df[c_navn].astype(str).str.strip()
    out = out[out["regnskapsnr"].str.len() > 0].reset_index(drop=True)
    return out


def load_konto_intervaller(path: Path) -> pd.DataFrame:
    """
    Leser 'Mapping standard kontoplan.xlsx' → DataFrame med kolonnene:
      - 'lo' (int)  | fra/fom/start
      - 'hi' (int)  | til/tom/slutt
      - 'regnskapsnr' (string)  | linjenummer
    """
    df = pd.read_excel(path, engine="openpyxl")
    # plukk kolonner robust
    c_lo = _pick_col(df, "fra", "fom", "start", "fra konto", "kontofra", "lo", "from", "konto fra")
    c_hi = _pick_col(df, "til", "tom", "slutt", "til konto", "kontotil", "hi", "to", "konto til")
    c_ln = _pick_col(df, "regnskapsnr", "linjenr", "regnskapslinjenr", "linje", "nr")

    if not c_lo or not c_hi or not c_ln:
        raise ValueError(
            "Fant ikke nødvendige kolonner i mapping-filen. "
            "Forventet noe ala 'fra/til' og 'regnskapsnr'."
        )

    out = pd.DataFrame()
    out["lo"] = pd.to_numeric(df[c_lo], errors="coerce").fillna(0).astype(int)
    out["hi"] = pd.to_numeric(df[c_hi], errors="coerce").fillna(0).astype(int)
    out["regnskapsnr"] = df[c_ln].astype(str).str.strip()
    out = out[(out["lo"] <= out["hi"]) & (out["regnskapsnr"].str.len() > 0)].copy()
    out = out.sort_values(["lo", "hi"]).reset_index(drop=True)
    return out


# -------------------------- Mapping-motor --------------------------

@dataclass
class MappingResult:
    rows: pd.DataFrame           # SB med regnskapsnr + regnskapsnavn (radnivå)
    agg: pd.DataFrame            # summer pr regnskapslinje
    meta: Dict[str, Any]         # info til manifest


def _assign_intervals_vectorized(konto: pd.Series, ranges: pd.DataFrame) -> pd.Series:
    """
    Konto (int) → regnskapsnr (string) via intervalltabell.
    Forutsetter at 'ranges' er sortert stigende på 'lo'.
    """
    if konto.empty:
        return pd.Series([], dtype="string")

    k = pd.to_numeric(konto, errors="coerce").fillna(-10**9).astype(int).to_numpy()
    lo = ranges["lo"].to_numpy()
    hi = ranges["hi"].to_numpy()
    rn = ranges["regnskapsnr"].to_numpy(dtype=object)

    # Finn siste 'lo' <= k (searchsorted), og verifiser k <= hi[idx]
    idx = np.searchsorted(lo, k, side="right") - 1
    ok = (idx >= 0) & (k <= hi[np.clip(idx, 0, len(hi) - 1)])
    out = np.empty_like(k, dtype=object)
    out[:] = None
    valid_idx = np.clip(idx[ok], 0, len(rn) - 1)
    out[ok] = rn[valid_idx]
    return pd.Series(out, index=konto.index, dtype="string")


def map_saldobalanse_to_regnskapslinjer(
    df_sb: pd.DataFrame,
    *,
    base_dir: Optional[Path] = None
) -> Optional[MappingResult]:
    """
    Legger til regnskapslinje på radnivå + aggregerer SB pr linje.
    Returnerer None hvis vi ikke finner kildefilene.
    """
    base = base_dir or find_kildefiler_dir()
    if not base:
        return None

    f_lines = _find_file_case_insensitive(base, REGNSKAPSLINJER_NAME)
    f_map = _find_file_case_insensitive(base, MAPPING_NAME)
    if not f_lines or not f_map:
        return None

    lines = load_regnskapslinjer(f_lines)
    ranges = load_konto_intervaller(f_map)

    # Normaliser kolonnenavn vi trenger i SB
    needed = [c for c in ("konto", "inngående balanse", "utgående balanse", "endring") if c in df_sb.columns]
    if "konto" not in needed:
        raise ValueError("SB mangler kolonnen 'konto' etter standardisering.")

    # 1) radnivå: legg på regnskapsnr + navn
    rn = _assign_intervals_vectorized(df_sb["konto"], ranges)
    rows = df_sb.copy()
    rows["regnskapsnr"] = rn
    rows = rows.dropna(subset=["regnskapsnr"]).copy()

    rows = rows.merge(lines, on="regnskapsnr", how="left")

    # 2) aggregat pr regnskapslinje
    agg_cols = [c for c in ("inngående balanse", "utgående balanse", "endring") if c in rows.columns]
    if not agg_cols:
        # fallback: hvis bare én saldo-kolonne finnes, aggreger den
        if "utgående balanse" in df_sb.columns:
            agg_cols = ["utgående balanse"]
        elif "endring" in df_sb.columns:
            agg_cols = ["endring"]

    if not agg_cols:
        raise ValueError("Fant ingen beløpskolonner i SB (forventet IB/UB/endring).")

    grp = rows.groupby(["regnskapsnr", "regnskapsnavn"], dropna=False)[agg_cols].sum().reset_index()
    grp = grp.sort_values("regnskapsnr").reset_index(drop=True)

    # 3) metadata til manifest
    src_accounts = int(pd.to_numeric(df_sb["konto"], errors="coerce").dropna().nunique())
    mapped_accounts = int(pd.to_numeric(rows["konto"], errors="coerce").dropna().nunique())
    unmapped = sorted(
        set(pd.to_numeric(df_sb["konto"], errors="coerce").dropna().astype(int))
        - set(pd.to_numeric(rows["konto"], errors="coerce").dropna().astype(int))
    )
    meta = {
        "base_dir": str(base),
        "source_files": [str(f_lines), str(f_map)],
        "source_accounts": src_accounts,
        "mapped_accounts": mapped_accounts,
        "unmapped_accounts": unmapped[:200],  # begrens for manifest-størrelse
        "unmapped_count": len(unmapped),
        "columns_aggregated": agg_cols,
    }

    return MappingResult(rows=rows, agg=grp, meta=meta)


def try_map_saldobalanse_to_regnskapslinjer(
    df_sb: pd.DataFrame, *, base_dir: Optional[Path] = None
) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], Dict[str, Any]]:
    """
    «Snill» wrapper som alltid returnerer meta.
    """
    try:
        res = map_saldobalanse_to_regnskapslinjer(df_sb, base_dir=base_dir)
        if not res:
            return None, None, {"active": False, "reason": "files_not_found"}
        return res.rows, res.agg, {"active": True, **res.meta}
    except Exception as exc:
        return None, None, {"active": False, "reason": f"{type(exc).__name__}: {exc}"}

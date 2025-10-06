# -*- coding: utf-8 -*-
# Robust mapping fra saldobalanse-kontoer til regnskapslinjer (regnr)
# - Leser "Regnskapslinjer.xlsx" (nr + navn) robust
# - Leser "Mapping standard kontoplan.xlsx" (Intervall-ark) robust
# - Mapper konto -> regnr etter intervaller (lo..hi)
# - Slår opp regnskapslinje-navn fra regnr
# - Lagrer/leser overstyringer pr. klient/år (sb2regnskap.json)
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, Dict

import json
import re
import pandas as pd

try:
    # for stier og årsmappestruktur
    from app.services.clients import year_paths
except Exception:
    from services.clients import year_paths  # type: ignore


# ----------------------------- helpers -----------------------------
NBSP = "\u00A0"

def _norm(s: str) -> str:
    return (
        str(s).strip().lower()
        .replace(NBSP, " ")
        .replace("\t", " ")
        .replace("-", " ")
        .replace("_", " ")
    )

def _to_int_series(s: pd.Series) -> pd.Series:
    t = (
        s.astype("string")
         .str.replace(r"\D", "", regex=True)
    )
    return pd.to_numeric(t, errors="coerce").astype("Int64")

def _first_of(df_cols: Iterable[str], *cands: str) -> Optional[str]:
    low = {c: _norm(c) for c in df_cols}
    for cand in cands:
        c = next((k for k, v in low.items() if v == _norm(cand)), None)
        if c:
            return c
    # fallback: startswith
    for cand in cands:
        c = next((k for k, v in low.items() if v.startswith(_norm(cand))), None)
        if c:
            return c
    return None


# ------------------ Regnskapslinjer.xlsx (nr + navn) ------------------
def read_regnskapslinjer(path: Path) -> pd.DataFrame:
    """
    Returnerer DF med kolonner: regnr (string) og regnskapslinje (string).
    Leser robust – finner kolonner ved navn/synonymer.
    """
    df = pd.read_excel(Path(path), engine="openpyxl")
    num_col = _first_of(df.columns, "regnr", "nr", "nummer", "linjenr", "regnskapsnr", "regnskapsnummer")
    name_col = _first_of(df.columns, "regnskapslinje", "linje", "navn", "tekst", "regnskapsnavn")

    # fallback: antak første "nummeraktige" + første tekst
    if not num_col:
        # velg den kolonnen som har mest tall
        counts = {c: pd.to_numeric(df[c], errors="coerce").notna().sum() for c in df.columns}
        num_col = max(counts, key=counts.get)

    if not name_col:
        # velg lengste tekst-kolonne
        lens = {c: df[c].astype(str).str.len().mean() for c in df.columns}
        name_col = max(lens, key=lens.get)

    out = pd.DataFrame()
    out["regnr"] = _to_int_series(df[num_col]).astype("Int64").astype("string")
    out["regnskapslinje"] = df[name_col].astype("string").str.strip()
    out = out.dropna(subset=["regnr"]).drop_duplicates(subset=["regnr"])
    return out.reset_index(drop=True)


# ------------- Mapping standard kontoplan.xlsx (Intervall) -------------
def _try_read_intervall(path: Path, *, sheet_hint: str = "Intervall") -> pd.DataFrame:
    # 1) prøv med eksplisitt "Intervall"
    try:
        return pd.read_excel(path, engine="openpyxl", sheet_name=sheet_hint, header=0)
    except Exception:
        pass
    # 2) prøv å finne et ark som *inneholder* "inter" i navnet
    try:
        x = pd.ExcelFile(path, engine="openpyxl")
        cand = next((s for s in x.sheet_names if "inter" in s.lower()), x.sheet_names[0])
        return x.parse(cand, header=0)
    except Exception as exc:
        raise RuntimeError(f"Kunne ikke lese Intervall-ark i {path.name}: {type(exc).__name__}: {exc}")

def read_konto_intervaller(path: Path) -> pd.DataFrame:
    """
    Leser intervall-arket robust og returnerer kolonner: lo, hi, regnr.
    Godtar varierende antall kolonner og header-rad.
    """
    # prøv flere header-rader (0..3)
    df_raw = None
    last_exc: Exception | None = None
    for hdr in (0, 1, 2, 3):
        try:
            df_raw = _try_read_intervall(Path(path))
            if hdr != 0:
                df_raw = pd.read_excel(Path(path), engine="openpyxl", sheet_name="Intervall", header=hdr)
            # fant vi nødvendige kolonner?
            c_lo  = _first_of(df_raw.columns, "fra", "lo", "konto fra", "konto fra nr", "start")
            c_hi  = _first_of(df_raw.columns, "til", "hi", "konto til", "konto til nr", "slutt", "stop")
            c_sum = _first_of(df_raw.columns, "sumnr", "regnr", "nr", "resultatnr", "linjenr")
            if c_lo and c_hi and c_sum:
                out = pd.DataFrame({
                    "lo":  _to_int_series(df_raw[c_lo]),
                    "hi":  _to_int_series(df_raw[c_hi]),
                    "regnr": _to_int_series(df_raw[c_sum]).astype("Int64").astype("string"),
                })
                out = out.dropna(subset=["lo", "hi", "regnr"])
                out["lo"] = out["lo"].astype(int)
                out["hi"] = out["hi"].astype(int)
                out = out[out["hi"] >= out["lo"]]
                return out.reset_index(drop=True)
        except Exception as exc:
            last_exc = exc
            continue
    raise RuntimeError(
        "Fant ikke nødvendige kolonner i 'Intervall'-arket. "
        "Sørg for at arket har kolonnene «Fra / Til / SumNr (regnr)»."
        + (f"  Siste feil: {type(last_exc).__name__}: {last_exc}" if last_exc else "")
    )


# ----------------------------- mapping motor -----------------------------
@dataclass(frozen=True)
class MapSources:
    regnskapslinjer_path: Path
    intervall_path: Path

def _build_mapper(intervals: pd.DataFrame):
    """Returnerer en funksjon konto:int -> regnr:str | None"""
    spans = [(int(r.lo), int(r.hi), str(r.regnr)) for _, r in intervals.iterrows()]
    spans.sort(key=lambda t: (t[0], t[1]))
    def fn(konto: int) -> Optional[str]:
        if pd.isna(konto): return None
        try:
            k = int(konto)
        except Exception:
            return None
        for lo, hi, reg in spans:
            if lo <= k <= hi:
                return reg
        return None
    return fn

def map_saldobalanse_df(df_sb: pd.DataFrame,
                        sources: MapSources) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Legger til kolonnene regnr + regnskapslinje på en SB-DF.
    Returnerer (df_med_mapping, reg_defs)
    """
    reg_defs = read_regnskapslinjer(sources.regnskapslinjer_path)
    intervals = read_konto_intervaller(sources.intervall_path)

    mapper = _build_mapper(intervals)
    df = df_sb.copy()
    # sikre konto som int
    if "konto" not in df.columns:
        raise ValueError("Saldobalansen mangler kolonnen «konto».")
    df["konto"] = _to_int_series(df["konto"]).astype("Int64")

    # regnr
    df["regnr"] = df["konto"].map(lambda x: None if pd.isna(x) else mapper(int(x)))
    # slå opp navn
    df = df.merge(reg_defs, on="regnr", how="left")

    # hyggelig kolonnerekkefølge (konto, kontonavn, regnr, regnskapslinje, …)
    first = [c for c in ("konto", "kontonavn", "regnr", "regnskapslinje") if c in df.columns]
    df = df[first + [c for c in df.columns if c not in first]]
    return df, reg_defs


# ----------------------------- lagring av overstyringer -----------------------------
def _overrides_path(root: Path, client: str, year: int) -> Path:
    yp = year_paths(Path(root), client, int(year))
    yp.mapping.mkdir(parents=True, exist_ok=True)
    return yp.mapping / "sb2regnskap.json"

def load_overrides(root: Path, client: str, year: int) -> Dict[str, str]:
    p = _overrides_path(root, client, year)
    if p.exists():
        try:
            d = json.loads(p.read_text("utf-8"))
            return {str(k): str(v) for k, v in (d.get("overrides") or {}).items()}
        except Exception:
            pass
    return {}

def save_overrides(root: Path, client: str, year: int, overrides: Dict[str, str]) -> Path:
    p = _overrides_path(root, client, year)
    data = {
        "overrides": {str(k): str(v) for k, v in overrides.items()}
    }
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), "utf-8")
    tmp.replace(p)
    return p

def apply_overrides(df: pd.DataFrame, overrides: Dict[str, str]) -> pd.DataFrame:
    if not overrides:
        return df
    out = df.copy()
    # konto kan være Int64 – sammenlign på streng
    konto_str = out["konto"].astype("Int64").astype("string")
    mask = konto_str.isin(list(overrides.keys()))
    out.loc[mask, "regnr"] = konto_str[mask].map(lambda k: overrides.get(str(k)))
    return out

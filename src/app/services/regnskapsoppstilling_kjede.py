# -*- coding: utf-8 -*-
"""
regnskapsoppstilling_kjede.py

Bygger regnskapsoppstilling (Resultat/Balanse) basert på **kjedet** modell:
    detalj (nr) → delsumnr → sumnr → sumnr2 → sluttsumnr

Hva er nytt i denne versjonen:
- Beholder all eksisterende funksjonalitet (I/O, CLI, KPI, m.m.).
- `compute_statement` er nå *robust* dersom den får et u-normalisert
  regnskapslinje-DataFrame (uten 'nr'); vi normaliserer selv.
- `detaljer` i retur er **konto-nivå** (saldobalanse-konti) for drilldown.
- `linje_detaljer` inneholder RL-detaljlinjene etter merge/fortegn.
- Mapping av `regnr` fra SB trigges også når regnr-kolonnen finnes,
  men er tom/tekstlig (blanke strenger → map).

Denne filen er en "superset" av tidligere utgaver slik at eksisterende
GUI-kall (inkl. CLI) fortsatt virker.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

# ------------------------- Synonymer / normalisering -------------------------

NBSP = "\\u00A0"
NNBSP = "\\u202F"

BALANCE_SYNONYMS = {
    "IB": [
        "ib", "inngående saldo", "inngående balanse", "inngaende saldo", "inngaende balanse",
        "ingående saldo", "ingående balanse", "opening balance", "opening"
    ],
    "UB": [
        "ub", "utgående saldo", "utgående balanse", "utgaende saldo", "utgaende balanse",
        "closing balance", "closing", "balance"
    ],
    "Endring": [
        "endring", "bevegelse", "diff", "change", "movement", "period", "this period"
    ],
}

KONTO_SYNONYMS = ["konto", "kontonr", "kontonummer", "account", "account no", "accountno", "account number"]
KONTONAVN_SYNONYMS = ["kontonavn", "kontotekst", "account name", "description", "tekst", "beskrivelse"]

RL_SYNONYMS = {
    "nr": ["nr", "regnr", "reg nr", "regn nr", "linjenr", "nummer", "sum nr", "sumnr", "regnskapsnr", "regnskapsnummer"],
    "regnskapslinje": ["regnskapslinje", "linje", "navn", "tekst", "beskrivelse", "regnskapslinjenavn"],
    "sumnivå": ["sumnivå", "sum_nivaa", "sumnivaa", "nivå", "nivaa", "sum nivå", "sum-nivå", "sum_nivå"],
    "sumpost": ["sumpost", "sum post", "sum-post", "sum"],
    "delsumnr": ["delsumnr", "delsum", "delsum nr", "del-sum nr", "delnummer", "delsumlinjenr"],
    "sumnr": ["sumnr", "sum nr", "sumlinje", "gruppe nr", "grupperingsnr"],
    "sumnr2": ["sumnr2", "sum nr2", "sum2", "nivå3 nr", "nivaa3 nr"],
    "sluttsumnr": ["sluttsumnr", "slutt-sum nr", "toppsumnr", "top sum nr", "sluttsum nr"],
    "regnskapstype": ["regnskapstype", "type", "art", "rt", "resultat/balanse", "kategori"],
    "fortegn": ["fortegn", "sign", "signum", "signering"],
    "formel": ["formel", "formula", "expr", "expression"],
    "med_i_sum": ["med_i_sum", "med i sum", "medisum", "exclude", "inkluder", "include"],
}

MAPPING_SYNONYMS = {
    "lo": ["fra", "from", "lo", "lower", "start"],
    "hi": ["til", "to", "hi", "upper", "slutt", "end"],
    "regnr": ["regnr", "reg nr", "regn nr", "sum", "sum nr", "nr", "nummer"],
}

# ---- små hjelpere ----

def _norm_name(s: str) -> str:
    return str(s).replace(NBSP, " ").replace(NNBSP, " ").strip().lower()

def _lower_map(cols: Iterable[str]) -> Dict[str, str]:
    return {_norm_name(c): str(c) for c in cols}

def _find_col(cols: Iterable[str], targets: List[str]) -> Optional[str]:
    low = _lower_map(cols)
    for t in targets:
        c = low.get(_norm_name(t))
        if c is not None:
            return c
    return None

def _rename_by_synonyms(df: pd.DataFrame, mapping: Dict[str, List[str]]) -> pd.DataFrame:
    low = _lower_map(df.columns)
    rename_map = {}
    for std, alts in mapping.items():
        if std in df.columns:
            continue
        for a in alts:
            c = low.get(_norm_name(a))
            if c is not None:
                rename_map[c] = std; break
    return df.rename(columns=rename_map) if rename_map else df

def _coerce_amount(s: pd.Series) -> pd.Series:
    """Tving tall uansett norsk/engelsk format til float."""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    raw = s.astype(str).str.replace(rf"[{NBSP}{NNBSP}\s]", "", regex=True)
    eu = raw.str.contains(r",\d{1,2}$")
    out = pd.Series(index=s.index, dtype="float64")
    out.loc[eu] = pd.to_numeric(
        raw.loc[eu].str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce"
    )
    out.loc[~eu] = pd.to_numeric(raw.loc[~eu], errors="coerce")
    return out.fillna(0.0)

def _to_int_safe(v) -> Optional[int]:
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    m = re.search(r"\d+", str(v))
    return int(m.group(0)) if m else None

# ------------------------- Leser saldobalanse -------------------------

def read_saldobalanse(path: Path) -> pd.DataFrame:
    """
    Leser saldobalanse fra Excel/CSV.
    Returnerer df med minst kolonner: konto, IB, Endring, UB (+ kontonavn hvis mulig).
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    if path.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(path, engine="openpyxl")
    elif path.suffix.lower() in {".csv", ".txt"}:
        try:
            df = pd.read_csv(path, sep=";", encoding="utf-8")
        except Exception:
            df = pd.read_csv(path, sep=",", encoding="utf-8")
    else:
        raise ValueError(f"Ukjent filtype for saldobalanse: {path.suffix}")

    df = _rename_by_synonyms(df, {"konto": KONTO_SYNONYMS, "kontonavn": KONTONAVN_SYNONYMS})
    if "konto" not in df.columns:
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                df = df.rename(columns={c: "konto"}); break
    if "konto" not in df.columns:
        raise ValueError("Fant ikke kolonnen for 'konto' i saldobalansen.")

    df = _rename_by_synonyms(df, BALANCE_SYNONYMS)
    for c in ("IB", "Endring", "UB"):
        if c not in df.columns:
            df[c] = 0.0
        df[c] = _coerce_amount(df[c])

    df["konto"] = pd.to_numeric(df["konto"], errors="coerce").round(0).astype("Int64")
    return df

# ------------------------- Leser Regnskapslinjer (kjedet) -------------------------

@dataclass
class RLColumns:
    nr: str
    name: str
    sumnivaa: Optional[str] = None
    sumpost: Optional[str] = None
    delsumnr: Optional[str] = None
    sumnr: Optional[str] = None
    sumnr2: Optional[str] = None
    sluttsumnr: Optional[str] = None
    regnskapstype: Optional[str] = None
    fortegn: Optional[str] = None
    formel: Optional[str] = None
    med_i_sum: Optional[str] = None

def _normalize_rl_df(df: pd.DataFrame) -> Tuple[pd.DataFrame, RLColumns]:
    """Normaliser RL‑DataFrame robust: finn 'nr'/'regnskapslinje', cast typer, sett synonymer."""
    df = _rename_by_synonyms(df, RL_SYNONYMS)

    nr_col = _find_col(df.columns, RL_SYNONYMS["nr"]) or None
    name_col = _find_col(df.columns, RL_SYNONYMS["regnskapslinje"]) or None

    # Heuristikk hvis fortsatt ikke funnet
    if not nr_col:
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                nr_col = c; break
    if not name_col:
        for c in df.columns:
            if c != nr_col and not pd.api.types.is_numeric_dtype(df[c]):
                name_col = c; break

    if not nr_col or not name_col:
        raise ValueError("Regnskapslinjer mangler tydelige kolonner for 'nr' og/eller 'regnskapslinje'.")

    cols = RLColumns(
        nr=nr_col,
        name=name_col,
        sumnivaa=_find_col(df.columns, RL_SYNONYMS["sumnivå"]),
        sumpost=_find_col(df.columns, RL_SYNONYMS["sumpost"]),
        delsumnr=_find_col(df.columns, RL_SYNONYMS["delsumnr"]),
        sumnr=_find_col(df.columns, RL_SYNONYMS["sumnr"]),
        sumnr2=_find_col(df.columns, RL_SYNONYMS["sumnr2"]),
        sluttsumnr=_find_col(df.columns, RL_SYNONYMS["sluttsumnr"]),
        regnskapstype=_find_col(df.columns, RL_SYNONYMS["regnskapstype"]),
        fortegn=_find_col(df.columns, RL_SYNONYMS["fortegn"]),
        formel=_find_col(df.columns, RL_SYNONYMS["formel"]),
        med_i_sum=_find_col(df.columns, RL_SYNONYMS["med_i_sum"]),
    )

    out = pd.DataFrame()
    out["nr"] = pd.to_numeric(df[cols.nr], errors="coerce").astype("Int64")
    out["regnskapslinje"] = df[cols.name].astype(str)

    def _opt_num(src: Optional[str]) -> pd.Series:
        return pd.to_numeric(df[src], errors="coerce").astype("Int64") if src else pd.Series([None]*len(df), dtype="Int64")

    out["sumnivå"] = _opt_num(cols.sumnivaa)
    out["delsumnr"] = _opt_num(cols.delsumnr)
    out["sumnr"] = _opt_num(cols.sumnr)
    out["sumnr2"] = _opt_num(cols.sumnr2)
    out["sluttsumnr"] = _opt_num(cols.sluttsumnr)

    out["sumpost"] = df[cols.sumpost].astype(str).str.strip().str.lower() if cols.sumpost else ""
    out["regnskapstype"] = df[cols.regnskapstype].astype(str).str.strip().str.capitalize() if cols.regnskapstype else ""
    if cols.fortegn:
        f = pd.to_numeric(df[cols.fortegn], errors="coerce"); out["fortegn"] = f.where(~f.isna(), 1.0).astype(float)
    else:
        out["fortegn"] = 1.0
    out["formel"] = df[cols.formel].astype(str) if cols.formel else ""
    out["med_i_sum"] = df[cols.med_i_sum].astype(str).str.strip().str.lower() if cols.med_i_sum else ""

    return out, cols

def read_regnskapslinjer_chain(path: Path, sheet: Optional[str] = None) -> Tuple[pd.DataFrame, RLColumns]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    xl = pd.ExcelFile(path, engine="openpyxl")
    if sheet is None:
        sheet = xl.sheet_names[0]
    df = xl.parse(sheet)
    return _normalize_rl_df(df)

# ------------------------- Leser mapping (konto → regnr) -------------------------

def read_intervals_mapping(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    path = Path(path)
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheet = sheet or (next((s for s in xl.sheet_names if _norm_name(s) == "intervall"), xl.sheet_names[0]))
    df = xl.parse(sheet)
    df = _rename_by_synonyms(df, MAPPING_SYNONYMS)
    for c in ("lo", "hi", "regnr"):
        if c not in df.columns:
            raise ValueError(f"Mapping-filen mangler kolonnen '{c}'.")
    out = pd.DataFrame({
        "lo": pd.to_numeric(df["lo"], errors="coerce"),
        "hi": pd.to_numeric(df["hi"], errors="coerce"),
        "regnr": pd.to_numeric(df["regnr"], errors="coerce"),
    }).dropna(subset=["regnr"])
    out["lo"] = out["lo"].fillna(0).astype(int)
    out["hi"] = out["hi"].fillna(out["lo"]).astype(int)
    out["regnr"] = out["regnr"].astype(int)
    out = out[out["hi"] >= out["lo"]].reset_index(drop=True)
    return out

def map_accounts_to_regnr(sb: pd.DataFrame, intervals: pd.DataFrame) -> pd.Series:
    """Vektorisert mapping av kontonummer → regnr vha intervaller."""
    if "konto" not in sb.columns:
        raise ValueError("Saldobalanse mangler 'konto'.")
    iv = intervals.sort_values(["lo", "hi"]).reset_index(drop=True)
    lo = iv["lo"].to_numpy(); hi = iv["hi"].to_numpy(); reg = iv["regnr"].to_numpy(dtype=int)
    konti = pd.to_numeric(sb["konto"], errors="coerce").fillna(-10**9).astype(int).to_numpy()
    idx = np.searchsorted(lo, konti, side="right") - 1
    valid = (idx >= 0) & (konti <= hi[np.clip(idx, 0, len(hi) - 1)])
    out = np.full_like(konti, fill_value=np.nan, dtype="float64")
    out[valid] = reg[np.clip(idx[valid], 0, len(reg) - 1)]
    return pd.Series(out, index=sb.index, dtype="Int64")

# ------------------------- Beregning (kjeden) -------------------------

@dataclass
class StatementResult:
    oppstilling: pd.DataFrame
    detaljer: pd.DataFrame             # KONTO-detaljer (for drilldown)
    kpi: Optional[pd.DataFrame] = None
    linje_detaljer: Optional[pd.DataFrame] = None  # RL-detaljer (valgfritt)

def _bool_like(s: pd.Series) -> pd.Series:
    v = s.astype(str).str.strip().str.lower()
    return v.isin(["ja", "yes", "y", "true", "1"])

def _parse_formula(expr: str) -> List[Tuple[int, int]]:
    if not isinstance(expr, str):
        return []
    s = expr.strip()
    if not s:
        return []
    if s.startswith("="):
        s = s[1:]
    token = ""
    sign = 1
    out: List[Tuple[int, int]] = []
    for ch in s:
        if ch in "+-":
            if token.strip():
                m = re.search(r"\\d+", token)
                if m:
                    out.append((sign, int(m.group(0))))
            token = ""
            sign = (1 if ch == "+" else -1)
        else:
            token += ch
    if token.strip():
        m = re.search(r"\\d+", token)
        if m:
            out.append((sign, int(m.group(0))))
    return out

def _sum_by_parent(det: pd.DataFrame, parent_col: str) -> pd.DataFrame:
    s = det.dropna(subset=[parent_col]).groupby(parent_col)[["IB", "Endring", "UB"]].sum()
    s = s.rename_axis("nr").reset_index().rename(columns={parent_col: "nr"})
    s["nr"] = pd.to_numeric(s["nr"], errors="coerce").astype("Int64")
    return s

def _ensure_rl_normalized(rl: pd.DataFrame) -> pd.DataFrame:
    """Bruker _normalize_rl_df hvis rl ikke tydelig har 'nr'/'regnskapslinje'."""
    if "nr" in rl.columns and "regnskapslinje" in rl.columns:
        return rl
    norm, _ = _normalize_rl_df(rl)
    return norm

def compute_statement(sb: pd.DataFrame,
                      rl: pd.DataFrame,
                      intervals: Optional[pd.DataFrame] = None,
                      apply_resultat_fortegn: bool = True,
                      kpi_defs: Optional[pd.DataFrame] = None) -> StatementResult:
    """
    Beregn full oppstilling (IB, Endring, UB).
    - Summer pr. nivå via groupby på detaljer
    - Formel-overstyring
    - Fortegn for resultatlinjer (hvis apply_resultat_fortegn=True)
    - Returnerer *konto-detaljer* i `detaljer` for drilldown.
    """
    # 1) Sikre 'regnr' i saldobalansen (robust: tom/tekstlig regnr → map)
    sb = sb.copy()
    if "regnr" in sb.columns:
        regnr_num = pd.to_numeric(sb["regnr"], errors="coerce")
    else:
        regnr_num = pd.Series([np.nan] * len(sb), index=sb.index)
    need_map = ("regnr" not in sb.columns) or regnr_num.isna().all()
    if need_map:
        if intervals is None:
            raise ValueError("Saldobalansen mangler 'regnr', og mapping (--map) er ikke gitt.")
        sb["regnr"] = map_accounts_to_regnr(sb, intervals)
    else:
        sb["regnr"] = regnr_num.astype("Int64")

    # 2) Konstruer KONTO-detaljer (for drilldown) og aggreger pr. regnr
    konto_det = sb.loc[pd.notna(sb["regnr"]), ["konto", "kontonavn", "regnr", "IB", "Endring", "UB"]].copy()
    konto_det["regnr"] = pd.to_numeric(konto_det["regnr"], errors="coerce").astype("Int64")
    aggr = konto_det.groupby("regnr")[["IB", "Endring", "UB"]].sum().reset_index().rename(columns={"regnr": "nr"})
    if aggr.empty:
        raise ValueError("Ingen konti ble mappet til regnr. "
                         "Sjekk intervall‑filen og at den peker mot detalj‑«nr» i Regnskapslinjer.")

    # 3) Normaliser regnskapslinjer
    rl = _ensure_rl_normalized(rl).copy()
    rl["nr"] = pd.to_numeric(rl["nr"], errors="coerce").astype("Int64")
    for col, default in [("sumnivå", None), ("sumpost", ""), ("med_i_sum", ""), ("fortegn", 1.0), ("regnskapstype", "")]:
        if col not in rl.columns:
            rl[col] = default

    # 4) RL‑detaljer (grunnlag for summer og ev. eksport)
    is_sum_row = _bool_like(rl["sumpost"])
    is_excluded = rl["med_i_sum"].astype(str).str.strip().str.lower().isin(["nei", "no", "false", "0"])
    sumnivaa_raw = pd.to_numeric(rl["sumnivå"], errors="coerce")
    has_zero = (sumnivaa_raw == 0).any()
    has_one  = (sumnivaa_raw == 1).any()
    detail_level = 0 if (has_zero and not has_one) else 1
    sumnivaa = sumnivaa_raw.fillna(detail_level)
    is_detail = (~is_sum_row) & (~is_excluded) & (sumnivaa == detail_level)

    det_rl = rl.loc[is_detail, ["nr", "regnskapslinje", "regnskapstype", "fortegn",
                                "delsumnr", "sumnr", "sumnr2", "sluttsumnr"]].copy()
    det_rl = det_rl.merge(aggr, on="nr", how="left")
    for c in ("IB", "Endring", "UB"):
        det_rl[c] = det_rl[c].fillna(0.0)

    # 5) Fortegn for resultatlinjer (kun på RL‑detaljer)
    if apply_resultat_fortegn:
        is_resultat = det_rl["regnskapstype"].astype(str).str.lower().str.startswith("resultat")
        sign = np.where(is_resultat, pd.to_numeric(det_rl["fortegn"], errors="coerce").fillna(1.0), 1.0)
        for c in ("IB", "Endring", "UB"):
            det_rl[c] = det_rl[c] * sign

    # 6) Summer pr. kjede‑nivå
    sums = []
    for parent_col in ["delsumnr", "sumnr", "sumnr2", "sluttsumnr"]:
        if parent_col in det_rl.columns:
            sums.append(_sum_by_parent(det_rl, parent_col))

    # 7) Samle verdier (detaljer + nivåsummer)
    values: Dict[int, Dict[str, float]] = {}
    for _, r in det_rl.iterrows():
        nr = _to_int_safe(r["nr"])
        if nr is None:
            continue
        values[nr] = {"IB": float(r["IB"]), "Endring": float(r["Endring"]), "UB": float(r["UB"])}
    for s in sums:
        for _, r in s.iterrows():
            nr = _to_int_safe(r["nr"])
            if nr is None:
                continue
            values[nr] = {"IB": float(r["IB"]), "Endring": float(r["Endring"]), "UB": float(r["UB"])}

    # 8) Formel-overstyringer
    if "formel" in rl.columns:
        rl_formel = rl.loc[rl["formel"].astype(str).str.strip().ne("")].copy().sort_values("nr")
        for _, row in rl_formel.iterrows():
            nr = _to_int_safe(row["nr"])
            if nr is None:
                continue
            terms = _parse_formula(row["formel"])
            sIB = sEnd = sUB = 0.0
            for sign, child in terms:
                v = values.get(child, {"IB": 0.0, "Endring": 0.0, "UB": 0.0})
                sIB += sign * v["IB"]; sEnd += sign * v["Endring"]; sUB += sign * v["UB"]
            values[nr] = {"IB": sIB, "Endring": sEnd, "UB": sUB}

    # 9) Bygg oppstilling i RL‑rekkefølge
    rows = []
    for _, r in rl.iterrows():
        nr = _to_int_safe(r["nr"])
        if nr is None:
            continue
        v = values.get(nr, {"IB": 0.0, "Endring": 0.0, "UB": 0.0})
        rows.append({
            "nr": nr,
            "regnskapslinje": str(r.get("regnskapslinje", "")),
            "sumnivå": int(r["sumnivå"]) if not pd.isna(r["sumnivå"]) else None,
            "sumpost": str(r.get("sumpost", "")),
            "regnskapstype": str(r.get("regnskapstype", "")),
            "IB": v["IB"], "Endring": v["Endring"], "UB": v["UB"],
            "formel": str(r.get("formel", "")),
        })
    oppstilling = pd.DataFrame(rows).sort_values("nr").reset_index(drop=True)

    # 10) Innrykk og retur
    def _innrykk(n):
        try: n = int(n); return max(0, (n - 1)) * 2
        except Exception: return 0
    oppstilling["innrykk"] = oppstilling["sumnivå"].apply(_innrykk)

    return StatementResult(
        oppstilling=oppstilling,
        detaljer=konto_det.reset_index(drop=True),   # KONTO‑detaljer for drilldown
        kpi=(evaluate_kpis(oppstilling, kpi_defs) if (kpi_defs is not None and not kpi_defs.empty) else None),
        linje_detaljer=det_rl.reset_index(drop=True),
    )

# ------------------------- KPI (valgfritt) -------------------------

def read_kpis(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() in {".xlsx", ".xls"}:
        xl = pd.ExcelFile(path, engine="openpyxl")
        sheet = sheet or next((s for s in xl.sheet_names if _norm_name(s) in {"kpi", "kpis"}), xl.sheet_names[0])
        df = xl.parse(sheet)
    else:
        df = pd.read_csv(path, sep=";")
    lowers = _lower_map(df.columns)
    name = lowers.get("navn", lowers.get("name"))
    expr = lowers.get("uttrykk", lowers.get("expr", lowers.get("expression")))
    felt = lowers.get("felt", None)
    fmt = lowers.get("format", lowers.get("fmt", None))
    if not (name and expr and felt):
        raise ValueError("KPI-definisjoner mangler 'navn', 'uttrykk' eller 'felt'.")
    out = pd.DataFrame({
        "navn": df[name].astype(str).str.strip(),
        "uttrykk": df[expr].astype(str).str.strip(),
        "felt": df[felt].astype(str).str.strip().str.capitalize(),
        "format": df[fmt].astype(str).str.strip() if fmt else "",
    })
    return out

def _tokenize_expr(s: str) -> List[str]:
    tokens = []; i = 0
    while i < len(s):
        ch = s[i]
        if ch.isspace(): i += 1; continue
        if ch in "+-*/()": tokens.append(ch); i += 1; continue
        m = re.match(r"\\d+", s[i:])
        if m: tokens.append(m.group(0)); i += len(m.group(0)); continue
        i += 1
    return tokens

def _to_rpn(tokens: List[str]) -> List[str]:
    prec = {"+": 1, "-": 1, "*": 2, "/": 2}
    out: List[str] = []; stack: List[str] = []
    for t in tokens:
        if t.isdigit(): out.append(t)
        elif t in prec:
            while stack and stack[-1] in prec and prec[stack[-1]] >= prec[t]:
                out.append(stack.pop())
            stack.append(t)
        elif t == "(": stack.append(t)
        elif t == ")":
            while stack and stack[-1] != "(": out.append(stack.pop())
            if stack and stack[-1] == "(": stack.pop()
    while stack: out.append(stack.pop())
    return out

def _eval_rpn(rpn: List[str], lookup: Dict[int, float]) -> float:
    st: List[float] = []
    for t in rpn:
        if t.isdigit(): st.append(float(lookup.get(int(t), 0.0)))
        elif t in {"+", "-", "*", "/"}:
            if len(st) < 2: st.append(0.0); continue
            b = st.pop(); a = st.pop()
            if t == "+": st.append(a + b)
            elif t == "-": st.append(a - b)
            elif t == "*": st.append(a * b)
            elif t == "/": st.append(0.0 if abs(b) < 1e-12 else a / b)
    return st[-1] if st else 0.0

def evaluate_kpis(oppstilling: pd.DataFrame, kpi_defs: pd.DataFrame) -> pd.DataFrame:
    lookups = {felt: oppstilling.set_index("nr")[felt].to_dict() for felt in ["IB", "Endring", "UB"]}
    rows = []
    for _, r in kpi_defs.iterrows():
        name = str(r["navn"]); expr = str(r["uttrykk"]); felt = str(r["felt"]).capitalize()
        fmt = str(r.get("format", ""))
        tokens = _tokenize_expr(expr); rpn = _to_rpn(tokens)
        value = _eval_rpn(rpn, lookups.get(felt, {}))
        rows.append({"navn": name, "felt": felt, "uttrykk": expr, "verdi": value, "format": fmt})
    return pd.DataFrame(rows)

# ------------------------- CLI (valgfritt for batch) -------------------------

def _build_parser():
    p = argparse.ArgumentParser(description="Bygg regnskapsoppstilling (kjedet modell)")
    p.add_argument("--sb", required=True, help="Saldobalanse (Excel/CSV)")
    p.add_argument("--rl", required=True, help="Regnskapslinjer - ny (Excel)")
    p.add_argument("--map", default=None, help="Mapping standard kontoplan.xlsx (valgfri hvis SB har regnr)")
    p.add_argument("--out", default="Oppstilling.xlsx", help="Output Excel (default: Oppstilling.xlsx)")
    p.add_argument("--kpi", default=None, help="KPI-definisjoner (valgfri Excel/CSV)")
    p.add_argument("--rl-sheet", default=None, help="Ark i regnskapslinjer-filen (default: første)")
    p.add_argument("--map-sheet", default=None, help="Ark i mapping-filen (default: Intervall eller første)")
    return p

def main() -> None:
    args = _build_parser().parse_args()
    sb_path = Path(args.sb); rl_path = Path(args.rl)
    map_path = Path(args.map) if args.map else None
    out_path = Path(args.out)
    kpi_path = Path(args.kpi) if args.kpi else None

    sb = read_saldobalanse(sb_path)
    rl, _ = read_regnskapslinjer_chain(rl_path, sheet=args.rl_sheet)
    intervals = read_intervals_mapping(map_path, sheet=args.map_sheet) if map_path else None
    kpi_defs = read_kpis(kpi_path) if kpi_path else None

    res = compute_statement(sb, rl, intervals=intervals, kpi_defs=kpi_defs)

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        opp = res.oppstilling.copy()
        for c in ("IB", "Endring", "UB"): opp[c] = opp[c].astype(float).round(2)
        opp.to_excel(xw, sheet_name="Oppstilling", index=False)

        det = res.detaljer.copy()
        for c in ("IB", "Endring", "UB"): det[c] = det[c].astype(float).round(2)
        det.to_excel(xw, sheet_name="Detaljer (konto)", index=False)

        if res.linje_detaljer is not None and not res.linje_detaljer.empty:
            rl_det = res.linje_detaljer.copy()
            for c in ("IB", "Endring", "UB"): rl_det[c] = rl_det[c].astype(float).round(2)
            rl_det.to_excel(xw, sheet_name="Detaljlinjer (RL)", index=False)

        if res.kpi is not None and not res.kpi.empty:
            k = res.kpi.copy(); k["verdi"] = k["verdi"].astype(float).round(4)
            k.to_excel(xw, sheet_name="KPI", index=False)

    print(f"Lagret: {out_path.resolve()}")

if __name__ == "__main__":
    main()

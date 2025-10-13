# -*- coding: utf-8 -*-
"""
regnskapsoppstilling_kjede.py

En selvstendig Python-modul som bygger en regnskapsoppstilling (Resultat/Balanse)
basert på en **kjedet** modell for regnskapslinjer:

  detalj (nr) → delsumnr → sumnr → sumnr2 → sluttsumnr

Hovedidé:
- Alle summer hentes **direkte fra detaljlinjene** ved enkle groupby pr. nivå.
- Eventuelle formler i regnskapslinje-filen overstyrer automatisk summering.
- Fortegn på resultatlinjer kan håndteres via en `fortegn`-kolonne på detaljlinjene
  (f.eks. +1 for inntekter og -1 for kostnader), slik at summer blir ren addisjon.
- Valgfritt felt `med_i_sum` lar deg ekskludere noter mv. fra summer.

I/O:
- Leser saldobalanse (Excel/CSV), regnskapslinjer (Excel) og (valgfritt) intervall-
  mapping (Excel) fra konto → regnr.
- Skriver ut en Excel med arkene: "Oppstilling", "Detaljer" (+ ev. "KPI" hvis definert).

Kjøring (CLI):
    python regnskapsoppstilling_kjede.py --sb SB.xlsx --rl "Regnskapslinjer - ny.xlsx" \
        [--map "Mapping standard kontoplan.xlsx"] [--out Oppstilling.xlsx] [--kpi kpi.xlsx]

Avhengigheter:
    pip install pandas openpyxl numpy

Merk:
- Dersom saldobalansen allerede har kolonnen "regnr", kan --map utelates.
- Kolonnenavn detekteres robust (synonymer for IB/UB/Endring, konto, m.m.).
- KPI-definisjoner er valgfrie og kan leveres i en egen fil/ark (se `read_kpis`).

(c) 2025
"""

from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

# ------------------------- Synonymer / normalisering -------------------------

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
    "nr": ["nr", "regnr", "linjenr", "nummer", "sum nr", "sumnr", "regnskapsnr"],
    "regnskapslinje": ["regnskapslinje", "linje", "navn", "tekst", "beskrivelse"],
    "sumnivå": ["sumnivå", "sum_nivaa", "nivå", "nivaa", "sum_nivå", "sum nivå", "sum-nivå"],
    "sumpost": ["sumpost", "sum post", "sum-post", "sum"],
    "delsumnr": ["delsumnr", "delsum", "delsum nr", "del-sum nr", "delnummer"],
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

NBSP = "\u00A0"


def _lower_map(cols: Iterable[str]) -> Dict[str, str]:
    return {str(c).strip().lower(): str(c) for c in cols}


def _find_col(cols: Iterable[str], targets: List[str]) -> Optional[str]:
    lowers = _lower_map(cols)
    for t in targets:
        key = str(t).strip().lower()
        if key in lowers:
            return lowers[key]
    return None


def _rename_by_synonyms(df: pd.DataFrame, mapping: Dict[str, List[str]]) -> pd.DataFrame:
    """Gi standardnavn for kolonner i df basert på synonymer."""
    lowers = _lower_map(df.columns)
    rename_map = {}
    for std, alts in mapping.items():
        if std in df.columns:
            continue
        for a in alts:
            c = lowers.get(str(a).lower())
            if c is not None:
                rename_map[c] = std
                break
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _coerce_amount(s: pd.Series) -> pd.Series:
    """Tving tall uansett norsk/engelsk format til float."""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    raw = s.astype(str).str.replace(rf"[{NBSP}\u202F\s]", "", regex=True)
    eu = raw.str.contains(r",\d{1,2}$")  # europeisk komma-desimal
    out = pd.Series(index=s.index, dtype="float64")
    out.loc[eu] = pd.to_numeric(
        raw.loc[eu].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    out.loc[~eu] = pd.to_numeric(raw.loc[~eu], errors="coerce")
    return out.fillna(0.0)


def _digits_only(v) -> Optional[str]:
    if v is None:
        return None
    s = re.sub(r"\D", "", str(v))
    return s or None


def _to_int_safe(v) -> Optional[int]:
    try:
        if v is None:
            return None
        if pd.isna(v):
            return None
    except Exception:
        pass
    m = re.search(r"\d+", str(v))
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


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
        # forsøk semikolon først, deretter komma
        try:
            df = pd.read_csv(path, sep=";", encoding="utf-8")
        except Exception:
            df = pd.read_csv(path, sep=",", encoding="utf-8")
    else:
        raise ValueError(f"Ukjent filtype for saldobalanse: {path.suffix}")

    # konto/konnavn
    df = _rename_by_synonyms(df, {"konto": KONTO_SYNONYMS, "kontonavn": KONTONAVN_SYNONYMS})
    if "konto" not in df.columns:
        # første numeriske kolonne som fallback
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                df = df.rename(columns={c: "konto"})
                break
    if "konto" not in df.columns:
        raise ValueError("Fant ikke kolonnen for 'konto' i saldobalansen.")

    # Balansesummer
    df = _rename_by_synonyms(df, BALANCE_SYNONYMS)
    for c in ("IB", "Endring", "UB"):
        if c not in df.columns:
            df[c] = 0.0

    df["IB"] = _coerce_amount(df["IB"])
    df["Endring"] = _coerce_amount(df["Endring"])
    df["UB"] = _coerce_amount(df["UB"])
    # konto som Int64 (bevarer tomme)
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


def read_regnskapslinjer_chain(path: Path, sheet: Optional[str] = None) -> Tuple[pd.DataFrame, RLColumns]:
    """Leser Regnskapslinjer-ny (kjedet) og normaliserer kolonnenavn."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    xl = pd.ExcelFile(path, engine="openpyxl")
    if sheet is None:
        sheet = xl.sheet_names[0]
    df = xl.parse(sheet)

    # gi standardnavn
    df = _rename_by_synonyms(df, RL_SYNONYMS)

    # finn påkrevde
    nr_col = _find_col(df.columns, ["nr"])
    name_col = _find_col(df.columns, ["regnskapslinje"])
    if not nr_col or not name_col:
        raise ValueError("Regnskapslinjer: mangler 'nr' og/eller 'regnskapslinje'.")

    # valgfri
    cols = RLColumns(
        nr=nr_col,
        name=name_col,
        sumnivaa=_find_col(df.columns, ["sumnivå"]),
        sumpost=_find_col(df.columns, ["sumpost"]),
        delsumnr=_find_col(df.columns, ["delsumnr"]),
        sumnr=_find_col(df.columns, ["sumnr"]),
        sumnr2=_find_col(df.columns, ["sumnr2"]),
        sluttsumnr=_find_col(df.columns, ["sluttsumnr"]),
        regnskapstype=_find_col(df.columns, ["regnskapstype"]),
        fortegn=_find_col(df.columns, ["fortegn"]),
        formel=_find_col(df.columns, ["formel"]),
        med_i_sum=_find_col(df.columns, ["med_i_sum"]),
    )

    # normaliser typer
    df["nr"] = pd.to_numeric(df[nr_col], errors="coerce").astype("Int64")
    df["regnskapslinje"] = df[name_col].astype(str)
    if cols.sumnivaa:
        df["sumnivå"] = pd.to_numeric(df[cols.sumnivaa], errors="coerce").astype("Int64")
    else:
        df["sumnivå"] = pd.Series([None] * len(df), dtype="Int64")
    if cols.sumpost:
        df["sumpost"] = df[cols.sumpost].astype(str).str.strip().str.lower()
    else:
        df["sumpost"] = ""
    for k, c in [("delsumnr", cols.delsumnr), ("sumnr", cols.sumnr),
                 ("sumnr2", cols.sumnr2), ("sluttsumnr", cols.sluttsumnr)]:
        if c:
            df[k] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
        else:
            df[k] = pd.Series([None] * len(df), dtype="Int64")
    if cols.regnskapstype:
        df["regnskapstype"] = df[cols.regnskapstype].astype(str).str.strip().str.capitalize()
    else:
        df["regnskapstype"] = ""
    if cols.fortegn:
        f = pd.to_numeric(df[cols.fortegn], errors="coerce")
        f = f.where(~f.isna(), 1.0)
        df["fortegn"] = f.astype(float)
    else:
        df["fortegn"] = 1.0
    if cols.formel:
        df["formel"] = df[cols.formel].astype(str)
    else:
        df["formel"] = ""
    if cols.med_i_sum:
        df["med_i_sum"] = df[cols.med_i_sum].astype(str).str.strip().str.lower()
    else:
        df["med_i_sum"] = ""

    return df, cols


# ------------------------- Leser mapping (konto → regnr) -------------------------

def read_intervals_mapping(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    """
    Leser "Mapping standard kontoplan.xlsx" og returnerer DataFrame med kolonnene:
        lo (int), hi (int), regnr (int)
    """
    path = Path(path)
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheet = sheet or (next((s for s in xl.sheet_names if str(s).strip().lower() == "intervall"), xl.sheet_names[0]))
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
    """
    Vektorisert mapping av kontonummer → regnr vha intervaller.
    Returnerer pd.Series (Int64) i samme index som sb.
    """
    if "konto" not in sb.columns:
        raise ValueError("Saldobalanse mangler 'konto'.")
    # sorter intervaller
    iv = intervals.sort_values(["lo", "hi"]).reset_index(drop=True)
    lo = iv["lo"].to_numpy()
    hi = iv["hi"].to_numpy()
    reg = iv["regnr"].to_numpy(dtype=int)
    # konti
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
    detaljer: pd.DataFrame
    kpi: Optional[pd.DataFrame] = None


def _bool_like(s: pd.Series) -> pd.Series:
    """Tolker 'ja/nei', 'true/false', '1/0' som bool."""
    v = s.astype(str).str.strip().str.lower()
    return v.isin(["ja", "yes", "y", "true", "1"])


def _parse_formula(expr: str) -> List[Tuple[int, int]]:
    """
    Parser en formel ala "=19+79-80" → [(+1,19), (+1,79), (-1,80)]
    Bare +, - og heltall er tillatt. Mellomrom og tekst etter tall ignoreres.
    """
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
                m = re.search(r"\d+", token)
                if m:
                    out.append((sign, int(m.group(0))))
            token = ""
            sign = (1 if ch == "+" else -1)
        else:
            token += ch
    if token.strip():
        m = re.search(r"\d+", token)
        if m:
            out.append((sign, int(m.group(0))))
    return out


def _sum_by_parent(det: pd.DataFrame, parent_col: str) -> pd.DataFrame:
    s = det.dropna(subset=[parent_col]).groupby(parent_col)[["IB", "Endring", "UB"]].sum()
    s = s.rename_axis("nr").reset_index().rename(columns={parent_col: "nr"})
    s["nr"] = pd.to_numeric(s["nr"], errors="coerce").astype("Int64")
    return s


def compute_statement(sb: pd.DataFrame,
                      rl: pd.DataFrame,
                      intervals: Optional[pd.DataFrame] = None,
                      apply_resultat_fortegn: bool = True,
                      kpi_defs: Optional[pd.DataFrame] = None) -> StatementResult:
    """
    Beregn full oppstilling (IB, Endring, UB) for alle linjer i regnskapslinjer (kjedet).
    - Summer pr. nivå via groupby på detaljer (sumnivå==1 & sumpost!=ja & med_i_sum!=nei)
    - Formel overstyrer automatisk sum der det finnes.
    - Fortegn for resultatlinjer (hvis apply_resultat_fortegn=True).
    """
    # 1) Sikre 'regnr' i saldobalansen
    if "regnr" not in sb.columns or sb["regnr"].isna().all():
        if intervals is None:
            raise ValueError("Saldobalansen mangler 'regnr', og mapping (--map) er ikke gitt.")
        sb = sb.copy()
        sb["regnr"] = map_accounts_to_regnr(sb, intervals)
    sb["regnr"] = pd.to_numeric(sb["regnr"], errors="coerce").astype("Int64")

    # 2) Aggreger kontosummer pr. regnr
    aggr = sb.groupby("regnr")[["IB", "Endring", "UB"]].sum().reset_index()
    aggr.rename(columns={"regnr": "nr"}, inplace=True)

    # 3) Normaliser regnskapslinjer
    rl = rl.copy()
    rl["nr"] = pd.to_numeric(rl["nr"], errors="coerce").astype("Int64")
    if "sumnivå" not in rl.columns:
        rl["sumnivå"] = pd.Series([None] * len(rl), dtype="Int64")
    if "sumpost" not in rl.columns:
        rl["sumpost"] = ""
    if "med_i_sum" not in rl.columns:
        rl["med_i_sum"] = ""
    if "fortegn" not in rl.columns:
        rl["fortegn"] = 1.0
    if "regnskapstype" not in rl.columns:
        rl["regnskapstype"] = ""

    # 4) Build DET (detaljer)
    is_sum_row = _bool_like(rl["sumpost"])
    is_excluded = rl["med_i_sum"].astype(str).str.strip().str.lower().isin(["nei", "no", "false", "0"])
    # Robust tolkning av sumnivå: støtt både 0- og 1-basert detaljnivå
    sumnivaa_raw = pd.to_numeric(rl["sumnivå"], errors="coerce")
    has_zero = (sumnivaa_raw == 0).any()
    has_one  = (sumnivaa_raw == 1).any()
    detail_level = 0 if (has_zero and not has_one) else 1
    sumnivaa = sumnivaa_raw.fillna(detail_level)
    is_detail = (~is_sum_row) & (~is_excluded) & (sumnivaa == detail_level)

    det = rl.loc[is_detail, ["nr", "regnskapslinje", "regnskapstype", "fortegn",
                             "delsumnr", "sumnr", "sumnr2", "sluttsumnr"]].copy()
    det = det.merge(aggr, on="nr", how="left")
    for c in ("IB", "Endring", "UB"):
        det[c] = det[c].fillna(0.0)

    # 5) Anvend fortegn for resultatlinjer
    if apply_resultat_fortegn:
        is_resultat = det["regnskapstype"].astype(str).str.lower().str.startswith("resultat")
        sign = np.where(is_resultat, pd.to_numeric(det["fortegn"], errors="coerce").fillna(1.0), 1.0)
        for c in ("IB", "Endring", "UB"):
            det[c] = det[c] * sign

    # 6) Groupby pr. kjede-nivå
    sums = []  # liste av dataframes med kolonnene: nr, IB, Endring, UB
    for parent_col in ["delsumnr", "sumnr", "sumnr2", "sluttsumnr"]:
        if parent_col in det.columns:
            sums.append(_sum_by_parent(det, parent_col))

    # 7) Samle alle verdier (detaljer + nivå-summer) i ett kart
    values = {}  # nr -> dict(IB, Endring, UB)
    # detaljer
    for _, r in det.iterrows():
        nr = int(r["nr"]) if pd.notna(r["nr"]) else None
        if nr is None:
            continue
        values[nr] = {
            "IB": float(r["IB"]),
            "Endring": float(r["Endring"]),
            "UB": float(r["UB"]),
        }
    # summer (per nivå)
    rl_levels = rl.set_index("nr")["sumnivå"].to_dict()
    for s in sums:
        for _, r in s.iterrows():
            nr = _to_int_safe(r["nr"])
            if nr is None:
                continue
            # Sett/overstyr – vi henter ALLTID sum direkte fra detaljer
            values[nr] = {
                "IB": float(r["IB"]),
                "Endring": float(r["Endring"]),
                "UB": float(r["UB"]),
            }

    # 8) Formel-overstyringer
    if "formel" in rl.columns:
        # Evaluer formler *etter* at detalj og nivåsummer er på plass
        rl_formel = rl.loc[rl["formel"].astype(str).str.strip().ne("")].copy()
        # For robusthet: evaluer i stigende rekkefølge av nr (ikke strikt nødvendig)
        rl_formel = rl_formel.sort_values("nr")
        for _, row in rl_formel.iterrows():
            nr = _to_int_safe(row["nr"])
            if nr is None:
                continue
            terms = _parse_formula(row["formel"])
            sIB = sEnd = sUB = 0.0
            for sign, child in terms:
                v = values.get(child, {"IB": 0.0, "Endring": 0.0, "UB": 0.0})
                sIB += sign * v["IB"]
                sEnd += sign * v["Endring"]
                sUB += sign * v["UB"]
            values[nr] = {"IB": sIB, "Endring": sEnd, "UB": sUB}

    # 9) Bygg oppstilling (alle linjer i regnskapslinjer i original rekkefølge)
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
            "IB": v["IB"],
            "Endring": v["Endring"],
            "UB": v["UB"],
            "formel": str(r.get("formel", "")),
        })
    oppstilling = pd.DataFrame(rows)
    # sortér på nr (men bevar opprinnelig hvis ønskelig)
    oppstilling = oppstilling.sort_values("nr").reset_index(drop=True)

    # 10) Detaljtabell for drilldown
    detaljer = det.copy().rename(columns={"nr": "regnr"})
    # legg på hjelpekoll. for drilldown nivå/gjennom
    def _innrykk(n):
        try:
            n = int(n)
            return max(0, (n - 1)) * 2
        except Exception:
            return 0
    oppstilling["innrykk"] = oppstilling["sumnivå"].apply(_innrykk)

    # 11) KPI-er (valgfritt)
    kpi_df = None
    if kpi_defs is not None and not kpi_defs.empty:
        kpi_df = evaluate_kpis(oppstilling, kpi_defs)

    return StatementResult(oppstilling=oppstilling, detaljer=detaljer, kpi=kpi_df)


# ------------------------- KPI (valgfritt) -------------------------

def read_kpis(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    """
    Leser KPI-definisjoner fra Excel/CSV.
    Forventede kolonner: navn, uttrykk, felt (IB|Endring|UB), format (valgfritt)
    Eksempel:  Driftsmargin | 80/19 | Endring | %
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() in {".xlsx", ".xls"}:
        xl = pd.ExcelFile(path, engine="openpyxl")
        sheet = sheet or next((s for s in xl.sheet_names if str(s).strip().lower() in {"kpi", "kpis"}), xl.sheet_names[0])
        df = xl.parse(sheet)
    else:
        df = pd.read_csv(path, sep=";")
    # normaliser
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
    """Tokeniserer tall/NR-navn og operatorer +-*/()"""
    tokens = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch.isspace():
            i += 1; continue
        if ch in "+-*/()":
            tokens.append(ch); i += 1; continue
        m = re.match(r"\d+", s[i:])
        if m:
            tokens.append(m.group(0)); i += len(m.group(0)); continue
        # ukjent token → stopp
        # for robusthet: hopp til neste symbol
        i += 1
    return tokens


def _to_rpn(tokens: List[str]) -> List[str]:
    """Shunting-yard for +-*/ og ()"""
    prec = {"+": 1, "-": 1, "*": 2, "/": 2}
    output: List[str] = []
    stack: List[str] = []
    for t in tokens:
        if t.isdigit():
            output.append(t)
        elif t in prec:
            while stack and stack[-1] in prec and prec[stack[-1]] >= prec[t]:
                output.append(stack.pop())
            stack.append(t)
        elif t == "(":
            stack.append(t)
        elif t == ")":
            while stack and stack[-1] != "(":
                output.append(stack.pop())
            if stack and stack[-1] == "(":
                stack.pop()
        else:
            # ignorer
            pass
    while stack:
        output.append(stack.pop())
    return output


def _eval_rpn(rpn: List[str], lookup: Dict[int, float]) -> float:
    st: List[float] = []
    for t in rpn:
        if t.isdigit():
            st.append(float(lookup.get(int(t), 0.0)))
        elif t in {"+", "-", "*", "/"}:
            if len(st) < 2:
                st.append(0.0); continue
            b = st.pop(); a = st.pop()
            if t == "+": st.append(a + b)
            elif t == "-": st.append(a - b)
            elif t == "*": st.append(a * b)
            elif t == "/": st.append(0.0 if abs(b) < 1e-12 else a / b)
        else:
            # ignorer
            pass
    return st[-1] if st else 0.0


def evaluate_kpis(oppstilling: pd.DataFrame, kpi_defs: pd.DataFrame) -> pd.DataFrame:
    """
    Evaluer KPI-er mot oppstillingen.
    - kpi_defs: kolonner navn, uttrykk, felt (IB|Endring|UB), format (valgfritt).
    Returnerer DataFrame med kolonnene: navn, verdi, felt, uttrykk, format.
    """
    # bygg opp lookup pr felt
    lookups = {}
    for felt in ["IB", "Endring", "UB"]:
        m = oppstilling.set_index("nr")[felt].to_dict()
        lookups[felt] = m
    rows = []
    for _, r in kpi_defs.iterrows():
        name = str(r["navn"])
        expr = str(r["uttrykk"])
        felt = str(r["felt"]).capitalize()
        fmt = str(r.get("format", ""))
        tokens = _tokenize_expr(expr)
        rpn = _to_rpn(tokens)
        value = _eval_rpn(rpn, lookups.get(felt, {}))
        rows.append({"navn": name, "felt": felt, "uttrykk": expr, "verdi": value, "format": fmt})
    return pd.DataFrame(rows)


# ------------------------- CLI -------------------------

def _build_parser() -> argparse.ArgumentParser:
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
    sb_path = Path(args.sb)
    rl_path = Path(args.rl)
    map_path = Path(args.map) if args.map else None
    out_path = Path(args.out)
    kpi_path = Path(args.kpi) if args.kpi else None

    # Les data
    sb = read_saldobalanse(sb_path)
    rl, _ = read_regnskapslinjer_chain(rl_path, sheet=args.rl_sheet)
    intervals = read_intervals_mapping(map_path, sheet=args.map_sheet) if map_path else None
    kpi_defs = read_kpis(kpi_path) if kpi_path else None

    # Beregn
    res = compute_statement(sb, rl, intervals=intervals, kpi_defs=kpi_defs)

    # Skriv Excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        df = res.oppstilling.copy()
        # pen visning: runde til 2 desimaler
        for c in ("IB", "Endring", "UB"):
            df[c] = df[c].astype(float).round(2)
        df.to_excel(xw, sheet_name="Oppstilling", index=False)
        det = res.detaljer.copy()
        for c in ("IB", "Endring", "UB"):
            det[c] = det[c].astype(float).round(2)
        det.to_excel(xw, sheet_name="Detaljer", index=False)
        if res.kpi is not None and not res.kpi.empty:
            k = res.kpi.copy()
            k["verdi"] = k["verdi"].astype(float).round(4)
            k.to_excel(xw, sheet_name="KPI", index=False)

    print(f"Lagret: {out_path.resolve()}")


if __name__ == "__main__":
    main()

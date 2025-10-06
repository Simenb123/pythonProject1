# src/app/services/mapping.py
from __future__ import annotations
from pathlib import Path
from typing import Mapping, Iterable, Optional, Dict, Any
import json
import re
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# sti-hjelper
try:
    from app.services.clients import mapping_file
except Exception:  # pragma: no cover
    from services.clients import mapping_file  # type: ignore

HB_REQUIRED = ("konto", "beløp", "dato", "bilagsnr")
SB_REQUIRED = ("konto", "kontonavn", "inngående balanse", "utgående balanse")

SYNONYMS: Dict[str, Iterable[str]] = {
    "konto": ("konto","kontonr","kontonummer","account","accountno","acct"),
    "kontonavn": ("kontonavn","kontotekst","account name","accountname"),
    "beløp": ("beløp","belop","amount","sum","total","amount_nok","debet","kredit"),
    "dato": ("dato","date","bilagsdato","post date","posting date","transdate"),
    "bilagsnr": (
        "bilagsnr","bilagsnummer","bilagsnum","bilag nr","bilag",
        "voucher","voucher nr","voucherno","voucher number",
        "doknr","dok nr","document no","document number",
        "verifikasjonsnr","verifnr","journalnr",
        "fakturanr","fakturanummer","invoice no","invoice number"
    ),
    "tekst": ("tekst","beskrivelse","description","post text","narrative","faktura"),
    "inngående balanse": ("inngående balanse","ib","opening balance","opening"),
    "utgående balanse": ("utgående balanse","ub","closing balance","closing"),
    "endring": ("endring","bevegelse","movement","change"),
}

NBSP = "\u00A0"
_BNR_RE = re.compile(r"[^0-9a-z]+", re.IGNORECASE)

def _norm(s: str) -> str:
    return (
        str(s).strip().lower()
        .replace(NBSP, " ")
        .replace("\t", " ")
        .replace("-", " ")
        .replace(".", " ")
    )

def _bnr_key(x: Any) -> str | None:
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    s = str(x).strip().lower()
    return _BNR_RE.sub("", s) or None

def _autoguess(df_cols: Iterable[str], source: str) -> Dict[str, str]:
    cols = list(df_cols)
    lut = {_norm(c): c for c in cols}
    target = HB_REQUIRED if source == "hovedbok" else SB_REQUIRED
    guess: Dict[str, str] = {}
    for std in target:
        for syn in SYNONYMS.get(std, (std,)):
            c = lut.get(_norm(syn))
            if c:
                guess[std] = c
                break
    if source == "hovedbok" and "tekst" not in guess:
        for syn in SYNONYMS["tekst"]:
            c = lut.get(_norm(syn))
            if c:
                guess["tekst"] = c
                break
    return guess

def _is_complete(mapping: Mapping[str, str], source: str) -> bool:
    req = set(HB_REQUIRED if source == "hovedbok" else SB_REQUIRED)
    return req.issubset({k for k, v in mapping.items() if v})

# ------------------------ I/O ------------------------
def load_mapping(root: Path, client: str, year: int, source: str) -> Dict[str, str] | None:
    p = mapping_file(root, client, year, source)
    if not p.exists(): return None
    try:
        d = json.loads(p.read_text("utf-8"))
        if isinstance(d, dict) and "mapping" in d and isinstance(d["mapping"], dict):
            return {str(k): str(v) for k, v in d["mapping"].items()}
        if isinstance(d, dict):
            return {str(k): str(v) for k, v in d.items()}
    except Exception:
        pass
    return None

def save_mapping(root: Path, client: str, year: int, source: str, mapping: Mapping[str, str]) -> Path:
    p = mapping_file(root, client, year, source); p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps({k: v for k, v in mapping.items() if v}, indent=2, ensure_ascii=False), "utf-8")
    tmp.replace(p)
    return p

# ------------------------ enkel dialog ------------------------
class _SimpleMapDialog(tk.Toplevel):
    def __init__(self, parent, source: str, columns: Iterable[str], preset: Mapping[str, str] | None = None):
        super().__init__(parent)
        self.title(f"Mapping – {source.capitalize()}"); self.resizable(False, False)
        self.transient(parent); self.grab_set()
        self.source = source
        cols = [""] + sorted(columns, key=str.casefold)
        req = HB_REQUIRED if source == "hovedbok" else SB_REQUIRED

        frm = ttk.Frame(self, padding=10); frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text="Velg kolonner i filen for hvert standardfelt:", font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))

        self.vars: Dict[str, tk.StringVar] = {}
        r = 1
        for std in req:
            ttk.Label(frm, text=std).grid(row=r, column=0, sticky="w")
            v = tk.StringVar(value=(preset or {}).get(std, ""))
            cb = ttk.Combobox(frm, values=cols, state="readonly", width=36, textvariable=v)
            cb.grid(row=r, column=1, sticky="w", pady=2)
            self.vars[std] = v
            r += 1

        if source == "hovedbok":
            for opt in ("tekst",):
                ttk.Label(frm, text=f"{opt} (valgfri)").grid(row=r, column=0, sticky="w")
                v = tk.StringVar(value=(preset or {}).get(opt, ""))
                cb = ttk.Combobox(frm, values=cols, state="readonly", width=36, textvariable=v)
                cb.grid(row=r, column=1, sticky="w", pady=2)
                self.vars[opt] = v
                r += 1

        btns = ttk.Frame(frm); btns.grid(row=r, column=0, columnspan=2, sticky="e", pady=(8,0))
        ttk.Button(btns, text="OK", command=self._ok).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Avbryt", command=self._cancel).grid(row=0, column=1)

        self.result: Dict[str, str] | None = None
        self.bind("<Return>", lambda *_: self._ok()); self.bind("<Escape>", lambda *_: self._cancel())

    def _ok(self):
        req = HB_REQUIRED if self.source == "hovedbok" else SB_REQUIRED
        mp = {k: v.get().strip() for k, v in self.vars.items() if v.get().strip()}
        missing = [f for f in req if f not in mp]
        if missing:
            messagebox.showwarning("Mangler felter", f"Mangler: {', '.join(missing)}", parent=self); return
        self.result = mp; self.destroy()

    def _cancel(self): self.result = None; self.destroy()

# ------------------------ offentlige API ------------------------
def ensure_mapping_interactive(parent, root: Path, client: str, year: int, source: str, df_sample: pd.DataFrame) -> Dict[str, str]:
    mp = load_mapping(root, client, year, source) or {}
    cols = list(df_sample.columns)
    guess = _autoguess(cols, source)
    mp = {**guess, **mp} if mp else guess

    if _is_complete(mp, source):
        save_mapping(root, client, year, source, mp); return mp

    parent = parent if (parent and getattr(parent, "winfo_exists", lambda: False)()) else None
    dlg = _SimpleMapDialog(parent, source, cols, preset=mp)
    if parent: parent.wait_window(dlg)
    else: dlg.wait_visibility(); dlg.grab_set(); dlg.wait_window(dlg)
    if dlg.result:
        save_mapping(root, client, year, source, dlg.result); return dlg.result
    save_mapping(root, client, year, source, mp); return mp

def edit_mapping_dialog(parent, root: Path, client: str, year: int, source: str, df_sample: pd.DataFrame) -> Dict[str, str] | None:
    mp0 = load_mapping(root, client, year, source) or _autoguess(df_sample.columns, source)
    parent = parent if (parent and getattr(parent, "winfo_exists", lambda: False)()) else None
    dlg = _SimpleMapDialog(parent, source, df_sample.columns, preset=mp0)
    if parent: parent.wait_window(dlg)
    else: dlg.wait_visibility(); dlg.grab_set(); dlg.wait_window(dlg)
    if dlg.result:
        save_mapping(root, client, year, source, dlg.result); return dlg.result
    return None

# ------------------------ standardisering ------------------------
_NUM_SPACES = r"[ \u00A0\u202F\u2009]"  # space, NBSP, thin, narrow
def _to_number(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s): return s
    t = (
        s.astype("string")
         .str.replace(_NUM_SPACES, "", regex=True)
         .pipe(lambda x: x.str.replace(".", "", regex=False) if x.str.contains(",", na=False).any() else x)
         .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(t, errors="coerce")

def standardize_with_mapping(
    df_raw: pd.DataFrame,
    mapping: Mapping[str, str] | None,
    *,
    source: str | None = None,
    parse_dates: bool = True,
    **_kwargs,
) -> pd.DataFrame:
    source = source or "hovedbok"
    mp = dict(mapping or {})
    # autogjett hvis noe mangler
    if not _is_complete(mp, source):
        mp = {**_autoguess(df_raw.columns, source), **mp}

    # rename
    rename = {v: k for k, v in mp.items() if v in df_raw.columns}
    df = df_raw.rename(columns=rename).copy()

    # normalisering
    if "konto" in df.columns:
        df["konto"] = (
            df["konto"].astype("string")
            .str.replace(r"\D", "", regex=True)
            .astype("Int64")
        )
    if "beløp" in df.columns:
        df["beløp"] = _to_number(df["beløp"])
    if parse_dates and "dato" in df.columns:
        df["dato"] = pd.to_datetime(df["dato"], errors="coerce", dayfirst=True)

    # HB: lag __bnr_key__ for raskt søk/drilldown
    if "bilagsnr" in df.columns:
        df["__bnr_key__"] = df["bilagsnr"].map(_bnr_key)

    # SB: beregn endring hvis mulig
    if source == "saldobalanse":
        ib = "inngående balanse"; ub = "utgående balanse"
        if ib in df.columns and ub in df.columns and "endring" not in df.columns:
            df[ib] = _to_number(df[ib]); df[ub] = _to_number(df[ub])
            df["endring"] = df[ub] - df[ib]
        # fornuftig kolonnerekkefølge
        cols = [c for c in ("konto","kontonavn",ib,"endring",ub) if c in df.columns]
        others = [c for c in df.columns if c not in cols]
        df = df[cols + others]

    return df

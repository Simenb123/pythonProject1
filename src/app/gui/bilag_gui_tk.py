# -*- coding: utf-8 -*-
# src/app/gui/bilag_gui_tk.py (fixed version)
"""
Denne filen er en tilpasset versjon av det originale `bilag_gui_tk.py` fra
pythonProject1‑repoet. Hovedendringen er at mappingen mellom konto og
regnskapslinjer (regnr) er gjort mer robust ved å tillate at regnskapsnummer
inneholder tekst (for eksempel «510 – Utvikling»). Tidligere forsøk på å
konvertere slike verdier til heltall med `int()` førte til `ValueError: invalid
literal for int() with base 10: 'Utvikling'`. Vi bruker nå et hjelpetillegg
`_extract_first_int` for å hente ut den første forekomsten av et heltall i en
streng. Denne funksjonen brukes ved lesing av lagrede mappingfiler,
oppdatering av interne mapping‑ordbøker og når DataFrame‑kolonner bygges for
visning.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import subprocess
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd

# -------------------------------------------------------------
# Synonymer for standardkolonner i saldobalanse
# -------------------------------------------------------------
# Følgende ordbok definerer alternative norske kolonnenavn som ofte
# forekommer i saldobalansefiler. Når vi leser inn en saldobalanse,
# forsøker vi å mappe disse til standardnavnene 'IB' (inngående balanse),
# 'UB' (utgående balanse) og 'Endring'. Du kan utvide listen med egne
# varianter ved å legge inn flere synonymer.
COLUMN_SYNONYMS: dict[str, list[str]] = {
    "IB": [
        "inngående saldo",
        "inngående balanse",
        "inngaende saldo",
        "inngaende balanse",
        "ingående saldo",
        "ingående balanse",
        "ib"
    ],
    "UB": [
        "utgående saldo",
        "utgående balanse",
        "utgaende saldo",
        "utgaende balanse",
        "ub"
    ],
    "Endring": [
        "endring",
        "bevegelse",
        "diff",
        "endring saldo",
        "endring beløp",
        "endringer"
    ],
}

# -------------------------------------------------------------
# Modulsøk
# -------------------------------------------------------------
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

def _import_all():
    # ingest (Parquet + manifest)
    try:
        from app.services.ingest import ensure_parquet_fresh, load_canonical_dataset
    except Exception:
        from services.ingest import ensure_parquet_fresh, load_canonical_dataset  # type: ignore

    try:
        from app.services.clients import (
            get_clients_root, load_meta, save_meta, year_paths
        )
    except Exception:
        from services.clients import get_clients_root, load_meta, save_meta, year_paths  # type: ignore

    # DataTable
    DataTable = None
    for modpath in ("app.gui.widgets.data_table", "widgets.data_table", "data_table"):
        try:
            _m = __import__(modpath, fromlist=["DataTable"])
            DataTable = getattr(_m, "DataTable", None)
            if DataTable is not None:
                break
        except Exception:
            continue
    if DataTable is None:
        raise ImportError("Fant ikke DataTable‑klassen.")
    return ensure_parquet_fresh, load_canonical_dataset, get_clients_root, load_meta, save_meta, year_paths, DataTable

(ensure_parquet_fresh,
 load_canonical_dataset,
 get_clients_root,
 load_meta,
 save_meta,
 year_paths,
 DataTable) = _import_all()

# -------------------------------------------------------------
# Global konfigurasjon: kildefiler og klientrot
# -------------------------------------------------------------
# Standardstier som brukes dersom de ikke finnes i meta eller defineres i miljø.
# Disse kan overstyres via miljøvariablene BHL_KILDEFILER_DIR og BHL_CLIENTS_DIR,
# eller via en global_config.json som ligger i prosjektmappen eller kildefil-mappen.
GLOBAL_KILDEFILER_DIR: str = os.environ.get("BHL_KILDEFILER_DIR", r"F:\\Dokument\\Kildefiler")
GLOBAL_CLIENTS_DIR: str = os.environ.get("BHL_CLIENTS_DIR", r"F:\\Dokument\\2\\BHL klienter\\Klienter")

def _load_global_config() -> None:
    """Oppdater globale stier fra global_config.json dersom den finnes."""
    global GLOBAL_KILDEFILER_DIR, GLOBAL_CLIENTS_DIR
    # Søk etter configfil i prosjektroten eller i kildefil-katalogen
    candidates = []
    try:
        base1 = Path(__file__).resolve().parents[3]
        candidates.append(base1 / "global_config.json")
    except Exception:
        pass
    try:
        base2 = Path(__file__).resolve().parents[2]
        candidates.append(base2 / "global_config.json")
    except Exception:
        pass
    try:
        candidates.append(Path(GLOBAL_KILDEFILER_DIR) / "global_config.json")
    except Exception:
        pass
    for cfg_path in candidates:
        try:
            if cfg_path.exists():
                with cfg_path.open("r", encoding="utf-8") as f:
                    cfg = json.load(f)
                if isinstance(cfg, dict):
                    GLOBAL_KILDEFILER_DIR = cfg.get("kildefiler_dir", GLOBAL_KILDEFILER_DIR)
                    GLOBAL_CLIENTS_DIR = cfg.get("clients_root", GLOBAL_CLIENTS_DIR)
                    break
        except Exception:
            continue

_load_global_config()

# --- regnskapslinjer (valgbar) ---
try:
    from app.services.regnskapslinjer import try_map_saldobalanse_to_regnskapslinjer
except Exception:
    try:
        from services.regnskapslinjer import try_map_saldobalanse_to_regnskapslinjer  # type: ignore
    except Exception:
        try_map_saldobalanse_to_regnskapslinjer = None  # type: ignore

# -------------------------------------------------------------
# Hjelpere
# -------------------------------------------------------------
def _digits_only(val) -> str | None:
    """Returner en streng med bare sifre, eller None hvis ingen sifre finnes."""
    if val is None:
        return None
    s = re.sub(r"\D", "", str(val))
    return s if s else None

def _bilag_key(val) -> str | None:
    if val is None:
        return None
    return re.sub(r"[^0-9a-z]+", "", str(val).lower()) or None

def _konto_key_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").round(0).astype("Int64").astype(str)
    return s.astype(str).str.replace(r"\D", "", regex=True)

def _extract_first_int(val) -> int | None:
    """
    Forsøk å hente ut den første forekomsten av ett eller flere sifre fra
    `val`. Returnerer None hvis ingen sifre finnes. Brukes for å tolke
    regnskapsnumre som kan inneholde tekst, f.eks. "510 - Utvikling".
    """
    if val is None:
        return None
    m = re.search(r"\d+", str(val))
    if m:
        try:
            return int(m.group(0))
        except Exception:
            return None
    return None

_CANON_PRIORITY = {
    "konto":     ["kontonr", "kontonummer", "konto nr", "konto", "accountno", "account no", "account"],
    "kontonavn": ["kontonavn", "kontonavn", "kontotekst", "account name", "accountname"],
    "dato":      ["dato", "bokføringsdato", "post date", "postdate", "transdate", "date"],
    "bilagsnr":  [
        "bilagsnr", "bilagsnummer", "bilagsnum", "bilag nr", "bilag", "voucher", "voucher nr", "voucherno",
        "voucher number", "doknr", "document no", "document number", "verifikasjonsnr", "verifnr", "journalnr"
    ],
    "tekst":     ["tekst", "beskrivelse", "description", "post text", "mottaker", "faktura", "narrative"],
}

# fil‑pref‑nøkler i meta["years"][år]["ui_prefs"]
_PREF_KILDE_DIR = "kildefiler_dir"
_PREF_RL_PATH   = "regnskapslinjer_path"
_PREF_MAP_PATH  = "kontoplan_mapping_path"

def _preferred_order(df: pd.DataFrame) -> list[str]:
    """
    Vis 'konto', 'kontonavn' først – og inkluder 'regnr' og 'regnskapslinje'
    rett etter, dersom de finnes.
    """
    cols = [c for c in df.columns if not str(c).startswith("__")]
    front = [c for c in ("konto", "kontonavn", "regnr", "regnskapslinje") if c in cols]
    rest  = [c for c in cols if c not in front]
    return front + rest

# -------------------------------------------------------------
# Argparser
# -------------------------------------------------------------
def _parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--client", required=True)
    p.add_argument("--year", type=int, required=True)
    p.add_argument("--source", choices=["hovedbok", "saldobalanse"], required=True)
    p.add_argument("--type", dest="vtype", choices=["interim", "ao", "versjon"], required=True)
    p.add_argument("--modus", choices=["analyse", "uttrekk"], default="analyse")
    p.add_argument("--konto", default=None)
    p.add_argument("--bilagsnr", default=None)
    p.add_argument("--adhoc_path", default=None)
    a = p.parse_args()
    if not getattr(a, "vtype", None) and getattr(a, "type", None):
        a.vtype = a.type
    return a

# -------------------------------------------------------------
# App
# -------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, client, year, source, vtype, modus,
                 konto: str | None = None, bilagsnr: str | None = None, adhoc_path: str | None = None):
        super().__init__()
        self.client, self.year, self.source, self.vtype, self.modus = client, int(year), source, vtype, modus
        self.prefilter_konto = _digits_only(konto) if konto else None
        self.prefilter_bkey  = _bilag_key(bilagsnr) if bilagsnr else None

        # Hent klient‑rot fra innstillinger. Dersom brukeren ennå ikke har
        # definert en root (via StartPortal), returnerer get_clients_root()
        # None. I så fall forsøker vi å falle tilbake til GLOBAL_CLIENTS_DIR,
        # men vi overstyrer ikke en verdi som allerede er satt av brukeren.
        self.root_dir = get_clients_root()
        if (not self.root_dir or not Path(self.root_dir).exists()) and GLOBAL_CLIENTS_DIR:
            try:
                self.root_dir = Path(GLOBAL_CLIENTS_DIR)
            except Exception:
                self.root_dir = None
        if not self.root_dir or not Path(self.root_dir).exists():
            messagebox.showerror(
                "Mangler klient‑rot", (
                    "Fant ikke klient‑rot i innstillinger. Gå tilbake til Start‑portalen "
                    "og bruk knappen “Bytt rot …” for å velge riktig katalog med klientmapper."
                ),
                parent=self,
            )
            self.destroy()
            return

        # meta (for persist av UI‑pref og filbaner)
        self.meta = load_meta(self.root_dir, self.client)
        # sørg for at global kildefil-rot er lagret i preferanser dersom det ikke finnes
        try:
            prefs = self.meta.setdefault("years", {}).setdefault(str(self.year), {}).setdefault("ui_prefs", {})
            if not prefs.get(_PREF_KILDE_DIR) or not Path(prefs.get(_PREF_KILDE_DIR, "")).exists():
                prefs[_PREF_KILDE_DIR] = str(GLOBAL_KILDEFILER_DIR)
                save_meta(self.root_dir, self.client, self.meta)
        except Exception:
            pass
        # sett miljøvariabler som brukes av tjenester til å finne kildefiler og klientkatalog
        try:
            os.environ.setdefault("AO7_KILDEFILER_DIR", str(GLOBAL_KILDEFILER_DIR))
            os.environ.setdefault("AO7_CLIENTS_ROOT", str(GLOBAL_CLIENTS_DIR))
        except Exception:
            pass

        self.title(f"Bilagsanalyse – {client} ({year}, {source}/{vtype})")
        self.geometry("1220x780")
        self.minsize(1024, 660)

        # 1) Sørg for kanonisk datasett (Parquet/pickle) og last det
        try:
            dataset_path, manifest = ensure_parquet_fresh(self, self.root_dir, self.client, self.year, self.source, self.vtype)
            df = load_canonical_dataset(dataset_path)
            # -----------------------------------------------------------------
            # Etter at vi har lastet inn datasettet, normaliser kolonnenavn for
            # saldobalanse ved å bruke forhåndsdefinerte synonymer. Noen
            # saldobalansefiler bruker norske feltnavn som «Inngående saldo» eller
            # «Utgående balanse» i stedet for standardnavnene «IB» og «UB». Vi
            # sjekker hver kolonne opp mot COLUMN_SYNONYMS og bygger et
            # omdøpingskart. Vi overstyrer ikke allerede eksisterende standardkolonner.
            rename_map: dict[str, str] = {}
            for col in df.columns:
                col_lower = str(col).strip().lower()
                for std_name, syns in COLUMN_SYNONYMS.items():
                    # Hopp over hvis vi allerede har en kolonne med standardnavnet
                    if std_name in df.columns:
                        continue
                    if any(col_lower == s.lower() for s in syns):
                        rename_map[col] = std_name
                        break
            if rename_map:
                try:
                    df = df.rename(columns=rename_map)
                except Exception:
                    pass
            self.src_path = str(dataset_path)
            self._manifest = manifest
        except Exception as exc:
            messagebox.showerror("Indeksering/innlasting feilet", f"{type(exc).__name__}: {exc}", parent=self)
            self.destroy()
            return

        # 2) Prefilter for drilldown (konto + bilag)
        self.df_full = df.copy()
        df_initial = df.copy()
        self.prefilter_info = ""
        if self.source == "hovedbok":
            if self.prefilter_konto and "konto" in df_initial.columns:
                ks = _konto_key_series(df_initial["konto"])
                df_initial = df_initial[ks == self.prefilter_konto]
                if len(df_initial) == 0:
                    self.prefilter_info = f" | Prefilter konto {self.prefilter_konto}: 0 rader – viser hele HB"
                    df_initial = df.copy()
                else:
                    self.prefilter_info = f" | Prefilter konto: {self.prefilter_konto} ({len(df_initial)} rader)"
            if self.prefilter_bkey:
                if "__bnr_key__" in df_initial.columns:
                    df2 = df_initial[df_initial["__bnr_key__"] == self.prefilter_bkey]
                elif "bilagsnr" in df_initial.columns:
                    keys = df_initial["bilagsnr"].map(_bilag_key)
                    df2 = df_initial[keys == self.prefilter_bkey]
                else:
                    df2 = df_initial.iloc[0:0]
                if len(df2) == 0:
                    self.prefilter_info += f" | Prefilter bilag {self.prefilter_bkey}: 0 rader – viser hele HB"
                else:
                    df_initial = df2
                    self.prefilter_info += f" | Bilag: {self.prefilter_bkey} ({len(df_initial)} rader)"

        # 3) Diagnostikk: én konto → typisk feil kilde
        self._dataset_note = ""
        if "konto" in df.columns:
            uniq = _konto_key_series(df["konto"]).dropna().unique()
            if len(uniq) <= 1:
                u = uniq[0] if len(uniq) == 1 else ""
                self._dataset_note = f" • Advarsel: datasettet har {len(uniq)} unik konto ({u}). Sjekk at aktiv HB‑fil er hele hovedboken."

        # 4) Last ev. tidligere regnr‑mapping fra disk og legg på kolonnene
        self._regnr_map_path = year_paths(self.root_dir, self.client, self.year).mapping / "sb_regnr.json"
        self._konto2regnr: dict[str, int] = self._load_regnr_map()
        self._regnr2name: dict[int, str] = {}  # fylles første gang du mapper/overstyrer

        # sørg for at self.df_full også har regnr/linje‑kolonner fra start slik at
        # de dukker opp i søkemenyen og andre operasjoner som bruker self.df_full
        self.df_full = self._with_regnskapslinjer_cols(self.df_full)
        df_initial = self._with_regnskapslinjer_cols(df_initial)

        # 5) Bygg UI
        self._build_ui(df_initial)

    # --------------------------- UI ---------------------------
    def _build_ui(self, df_initial: pd.DataFrame):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(top, text="Søk i:").pack(side="left")

        choices = ["Alle kolonner"]
        for c in ["konto", "kontonavn", "dato", "bilagsnr", "tekst", "regnr", "regnskapslinje"]:
            if c in self.df_full.columns:
                choices.append(c)
        for c in self.df_full.columns:
            if not c.startswith("__") and c not in choices:
                choices.append(c)

        self.cmb_col = ttk.Combobox(self, state="readonly", width=22, values=choices)
        self.cmb_col.set("Alle kolonner")
        self.cmb_col.pack(in_=top, side="left", padx=(4, 8))

        self.ent_q = ttk.Entry(top, width=36)
        self.ent_q.pack(side="left", padx=(0, 6))
        self.ent_q.bind("<Return>", lambda e: self._apply_search())
        self.ent_q.focus_set()

        ttk.Button(top, text="Søk", command=self._apply_search).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Tøm", command=self._reset_view).pack(side="left", padx=(0, 12))

        # SB: mappe konto -> regnskapslinjer
        if self.source == "saldobalanse" and try_map_saldobalanse_to_regnskapslinjer:
            ttk.Button(top, text="Map til regnskapslinjer …",
                       command=self._map_regn).pack(side="left", padx=(8, 0))

        ttk.Button(top, text="Sett regnr …", command=self._set_regnr_manual).pack(side="left")

        # Info‑linje
        info_txt = f"Kilde: {Path(self.src_path).name}"
        if self.prefilter_info:
            info_txt += self.prefilter_info
        if self._dataset_note:
            info_txt += self._dataset_note
        self.info = ttk.Label(self, text=info_txt, anchor="w")
        self.info.pack(fill="x", padx=8, pady=(2, 2))

        # DataTable
        self.table = DataTable(self, df=df_initial, page_size=500)
        self.table.pack(fill="both", expand=True, padx=8, pady=(2, 8))

        # Drilldown
        self._install_drilldown(self.table)

    def _reset_view(self):
        self.ent_q.delete(0, tk.END)
        self.cmb_col.set("Alle kolonner")
        cols = _preferred_order(self.df_full)
        self.table.set_dataframe(self._with_regnskapslinjer_cols(self.df_full[cols]), reset=True)
        self.table.refresh()
        info_txt = f"Kilde: {Path(self.src_path).name}"
        if self._dataset_note:
            info_txt += self._dataset_note
        self.info.config(text=info_txt)

    # --------------------------- Søk --------------------------
    def _apply_search(self):
        col = (self.cmb_col.get() or "Alle kolonner").strip()
        expr = (self.ent_q.get() or "").strip()
        df = self._with_regnskapslinjer_cols(self.df_full)

        if not expr:
            self._reset_view()
            return

        ops = ("==", "!=", ">=", "<=", ">", "<")

        if col == "Alle kolonner":
            mask = pd.Series(False, index=df.index)
            for c in df.columns:
                if c.startswith("__"):
                    continue
                if pd.api.types.is_string_dtype(df[c]) or df[c].dtype == "object":
                    mask |= df[c].astype(str).str.contains(expr, case=False, na=False, regex=False)
            cols = _preferred_order(df[mask])
            self.table.set_dataframe(df[mask][cols], reset=True)
            self.table.refresh()
            return

        if col not in df.columns:
            messagebox.showwarning("Kolonne mangler", f"Fant ikke kolonnen «{col}».")
            return

        # Bilagsnr → bruk normalisert nøkkel
        if col == "bilagsnr" and not expr.startswith(ops):
            key = _bilag_key(expr)
            keys = df["__bnr_key__"] if "__bnr_key__" in df.columns else df["bilagsnr"].map(_bilag_key)
            out = df[keys == key]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True)
            self.table.refresh()
            return

        # Konto → prefiks som default, eksakt med '=='
        if col == "konto" and not expr.startswith(ops):
            key = re.sub(r"\D", "", expr)
            ks = _konto_key_series(df[col])
            out = df[ks.str.startswith(key)]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True)
            self.table.refresh()
            return

        # Operator‑sammenligning
        op = None
        for t in ops:
            if expr.startswith(t):
                op = t
                rhs = expr[len(t):].strip()
                break

        s = df[col]
        if op is None:
            out = df[s.astype(str).str.contains(expr, case=False, na=False, regex=False)]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True)
            self.table.refresh()
            return

        # Dato?
        if pd.api.types.is_datetime64_any_dtype(s):
            rhs_dt = pd.to_datetime(rhs, errors="coerce")
            if pd.isna(rhs_dt):
                cmp = s.astype(str)
                if op == "==":
                    df2 = df[cmp == rhs]
                elif op == "!=":
                    df2 = df[cmp != rhs]
                elif op == ">=":
                    df2 = df[cmp >= rhs]
                elif op == "<=":
                    df2 = df[cmp <= rhs]
                elif op == ">":
                    df2 = df[cmp > rhs]
                elif op == "<":
                    df2 = df[cmp < rhs]
            else:
                if op == "==":
                    df2 = df[s == rhs_dt]
                elif op == "!=":
                    df2 = df[s != rhs_dt]
                elif op == ">=":
                    df2 = df[s >= rhs_dt]
                elif op == "<=":
                    df2 = df[s <= rhs_dt]
                elif op == ">":
                    df2 = df[s > rhs_dt]
                elif op == "<":
                    df2 = df[s < rhs_dt]
            self.table.set_dataframe(df2[_preferred_order(df2)], reset=True)
            self.table.refresh()
            return

        # Tall
        try:
            rhs_num = pd.to_numeric(rhs)
            s_num = pd.to_numeric(s, errors="coerce")
            if op == "==":
                df2 = df[s_num == rhs_num]
            elif op == "!=":
                df2 = df[s_num != rhs_num]
            elif op == ">=":
                df2 = df[s_num >= rhs_num]
            elif op == "<=":
                df2 = df[s_num <= rhs_num]
            elif op == ">":
                df2 = df[s_num > rhs_num]
            elif op == "<":
                df2 = df[s_num < rhs_num]
        except Exception:
            cmp = s.astype(str).str.strip()
            if op == "==":
                df2 = df[cmp == rhs]
            elif op == "!=":
                df2 = df[cmp != rhs]
            elif op == ">=":
                df2 = df[cmp >= rhs]
            elif op == "<=":
                df2 = df[cmp <= rhs]
            elif op == ">":
                df2 = df[cmp > rhs]
            elif op == "<":
                df2 = df[cmp < rhs]
        self.table.set_dataframe(df2[_preferred_order(df2)], reset=True)
        self.table.refresh()

    # ------------------------- Drilldown ----------------------
    def _install_drilldown(self, table_widget):
        if hasattr(table_widget, "bind_row_double_click"):
            if self.source == "saldobalanse":
                table_widget.bind_row_double_click(self._sb_to_hb)
            else:
                table_widget.bind_row_double_click(self._hb_to_hb)
            return

        # Fallback: tree‑bind
        tree = getattr(table_widget, "tree", None)
        if tree is None:
            return

        def _on_dclick(_):
            iid = tree.focus() or (tree.selection()[0] if tree.selection() else None)
            if not iid:
                return
            values = tree.item(iid, "values") or []
            cols = list(tree["columns"])
            row = pd.Series({c: (values[i] if i < len(values) else None) for i, c in enumerate(cols)})
            (self._sb_to_hb if self.source == "saldobalanse" else self._hb_to_hb)(row)

        tree.bind("<Double-1>", _on_dclick, add="+")

    def _sb_to_hb(self, row: pd.Series):
        # hent konto robust fra rad
        konto_val = None
        for name in _CANON_PRIORITY["konto"]:
            for c in row.index:
                if c.casefold() == name.casefold():
                    konto_val = row[c]
                    break
            if konto_val is not None:
                break
        konto = _digits_only(konto_val)
        if not konto:
            messagebox.showinfo("Drilldown", "Fant ikke kontonummer i valgt rad.")
            return
        self._open_hovedbok(konto=konto)

    def _hb_to_hb(self, row: pd.Series):
        bnr = None
        for name in _CANON_PRIORITY["bilagsnr"]:
            for c in row.index:
                if c.casefold() == name.casefold():
                    bnr = row[c]
                    break
            if bnr is not None:
                break
        key = _bilag_key(bnr)
        if not key:
            messagebox.showinfo("Drilldown", "Fant ikke bilagsnummer i valgt rad.")
            return
        self._open_hovedbok(bilagsnr=str(bnr))

    def _open_hovedbok(self, konto: str | None = None, bilagsnr: str | None = None):
        env = dict(os.environ)
        env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
        args = [sys.executable, "-m", "app.gui.bilag_gui_tk",
                "--client", self.client, "--year", str(self.year), "--source", "hovedbok", "--type", self.vtype, "--modus", "analyse"]
        if konto:
            args += ["--konto", konto]
        if bilagsnr:
            args += ["--bilagsnr", bilagsnr]
        subprocess.Popen(args, shell=False, cwd=str(SRC), env=env)

    # ------------------------- Regnskapslinjer (NYTT) ----------------------
    def _prefs_node(self) -> dict:
        """Returner (og opprett ved behov) meta->years[år]->ui_prefs."""
        years = self.meta.setdefault("years", {})
        y = years.setdefault(str(self.year), {"ui_prefs": {}})
        y.setdefault("ui_prefs", {})
        return y["ui_prefs"]

    def _save_prefs(self):
        try:
            save_meta(self.root_dir, self.client, self.meta)
        except Exception:
            pass

    def _load_regnr_map(self) -> dict[str, int]:
        p = self._regnr_map_path
        if p.exists():
            try:
                data = json.loads(p.read_text("utf-8"))
                out: dict[str, int] = {}
                for k, v in data.items():
                    num = _extract_first_int(v)
                    if num is not None:
                        out[str(k)] = num
                return out
            except Exception:
                return {}
        return {}

    def _save_regnr_map(self):
        self._regnr_map_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._regnr_map_path.with_suffix(".tmp")
        # lagre som {konto(str): regnr(int)}
        data = {str(k): int(v) for k, v in self._konto2regnr.items() if v is not None}
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
        tmp.replace(self._regnr_map_path)

    def _pick_file(self, title: str, types: tuple[tuple[str, str], ...], pref_key: str) -> Path | None:
        prefs = self._prefs_node()
        ini_dir = None
        if prefs.get(pref_key):
            ini_dir = str(Path(prefs[pref_key]).parent)
        elif prefs.get(_PREF_KILDE_DIR):
            ini_dir = prefs[_PREF_KILDE_DIR]
        p = filedialog.askopenfilename(title=title, filetypes=list(types), initialdir=ini_dir or "")
        if not p:
            return None
        prefs[pref_key] = p
        prefs[_PREF_KILDE_DIR] = str(Path(p).parent)
        self._save_prefs()
        return Path(p)

    def _ensure_regn_files(self) -> tuple[Path, Path] | None:
        """
        Sørg for at vi har stier til:
          1) Regnskapslinjer.xlsx  (regnr → regnskapslinjenavn)
          2) Mapping standard kontoplan.xlsx  (intervaller konto → regnr)
        Bruker/lager persist i meta.ui_prefs slik at samme filer brukes neste gang.
        """
        prefs = self._prefs_node()
        rl_path = Path(prefs.get(_PREF_RL_PATH, "")) if prefs.get(_PREF_RL_PATH) else None
        map_path = Path(prefs.get(_PREF_MAP_PATH, "")) if prefs.get(_PREF_MAP_PATH) else None
        if not (rl_path and rl_path.exists()):
            rl_path = self._pick_file("Velg Regnskapslinjer.xlsx",
                                      (("Excel", "*.xlsx *.xls"), ("Alle filer", "*.*")), _PREF_RL_PATH)
            if not rl_path:
                return None
        if not (map_path and map_path.exists()):
            map_path = self._pick_file("Velg Mapping standard kontoplan.xlsx",
                                       (("Excel", "*.xlsx *.xls"), ("Alle filer", "*.*")), _PREF_MAP_PATH)
            if not map_path:
                return None
        return rl_path, map_path

    def _read_regnskapslinjer_lut(self, p: Path) -> dict[int, str]:
        """
        Les 'Regnskapslinjer.xlsx' → {regnr:int -> navn:str}.
        Robust kolonnefinn: ser etter 'regnr' + tekstkolonne.
        """
        df = pd.read_excel(p, engine="openpyxl")
        low = {c.lower().strip(): c for c in df.columns}
        # finn regnr‑kolonne
        reg_col = None
        for cand in ("regnr", "reg nr", "regn nr", "nr", "nummer", "linjenr"):
            if cand in low:
                reg_col = low[cand]
                break
        if not reg_col:
            # fall‑back: første numeriske kolonne
            for c in df.columns:
                if pd.api.types.is_numeric_dtype(df[c]):
                    reg_col = c
                    break
        # finn navn‑kolonne
        name_col = None
        for cand in ("regnskapslinje", "linje", "regnskapslinjenavn", "navn", "tekst", "beskrivelse"):
            if cand in low:
                name_col = low[cand]
                break
        if not name_col:
            # fall‑back: første ikke‑numeriske kolonne ulik reg_col
            for c in df.columns:
                if c != reg_col and not pd.api.types.is_numeric_dtype(df[c]):
                    name_col = c
                    break
        if not reg_col or not name_col:
            raise ValueError("Fant ikke passende kolonner i Regnskapslinjer.xlsx")
        out: dict[int, str] = {}
        for _, r in df[[reg_col, name_col]].dropna().iterrows():
            try:
                rn = int(pd.to_numeric(r[reg_col], errors="coerce"))
                nm = str(r[name_col]).strip()
                if rn:
                    out[rn] = nm
            except Exception:
                pass
        return out

    def _read_intervals(self, p: Path) -> pd.DataFrame:
        """
        Les "Mapping standard kontoplan.xlsx" – ark «Intervall».
        Vi leser uten usecols og håndterer variable kolonneantall for å
        unngå ParserError usecols/out‑of‑bounds.
        Returnerer df med kolonnene: lo(int), hi(int), regnr(int)
        """
        try:
            xl = pd.ExcelFile(p, engine="openpyxl")
        except Exception as exc:
            raise RuntimeError(f"Kunne ikke lese mapping‑arbeidet: {type(exc).__name__}: {exc}")

        # Forsøk først å lese arket med navn "Intervall" med overskrift for å
        # identifisere kolonner basert på navn. Hvis dette lykkes, bruker vi
        # kolonnenavnene til å finne 'fra', 'til' og 'regnr'.
        sheet_name = None
        for s in xl.sheet_names:
            if str(s).strip().lower() == "intervall":
                sheet_name = s
                break
        try:
            if sheet_name is not None:
                df_named = xl.parse(sheet_name, header=0)
                lowers = {c.lower().strip(): c for c in df_named.columns}
                lo_col = None
                hi_col = None
                reg_col = None
                # Mulige navn for hver kolonne
                lo_candidates = ["fra", "from", "lo", "lower", "start"]
                hi_candidates = ["til", "to", "hi", "upper", "slutt", "end"]
                reg_candidates = ["regnr", "reg nr", "regn nr", "sum", "sum nr", "nr", "nummer"]
                for cand in lo_candidates:
                    if cand in lowers:
                        lo_col = lowers[cand]
                        break
                for cand in hi_candidates:
                    if cand in lowers:
                        hi_col = lowers[cand]
                        break
                for cand in reg_candidates:
                    if cand in lowers:
                        reg_col = lowers[cand]
                        break
                # Dersom vi har identifisert alle tre kolonner, bygg DataFrame
                if lo_col and hi_col and reg_col:
                    lo = pd.to_numeric(df_named[lo_col], errors="coerce")
                    hi = pd.to_numeric(df_named[hi_col], errors="coerce")
                    reg = pd.to_numeric(df_named[reg_col], errors="coerce")
                    df_res = pd.DataFrame({"lo": lo, "hi": hi, "regnr": reg}).dropna(subset=["regnr"])
                    df_res["lo"] = df_res["lo"].fillna(0).astype(int)
                    df_res["hi"] = df_res["hi"].fillna(df_res["lo"]).astype(int)
                    df_res["regnr"] = df_res["regnr"].astype(int)
                    df_res = df_res[df_res["hi"] >= df_res["lo"]]
                    return df_res.reset_index(drop=True)
        except Exception:
            # Fortsett med fallback under dersom noe går galt
            pass

        # Fallback: les arket uten overskrifter (gamle format) og bruk
        # posisjonsbasert parsing. Dette håndterer gamle filer med 3-4 kolonner.
        try:
            raw = xl.parse(sheet_name or 0, header=None)
        except Exception as exc:
            raise RuntimeError(f"Kunne ikke lese mapping‑arbeidet: {type(exc).__name__}: {exc}")
        n = raw.shape[1]
        def col(i):
            return raw.iloc[:, i] if i < n else pd.Series([None] * len(raw))
        lo = pd.to_numeric(col(0), errors="coerce")
        hi = pd.to_numeric(col(2), errors="coerce")
        reg = pd.to_numeric(col(3), errors="coerce")
        # fallback: finn første numeric kolonne < 10000 som regnr hvis reg er tom
        if reg.notna().sum() == 0:
            for i in range(n):
                s = pd.to_numeric(col(i), errors="coerce")
                if s.notna().sum() and (s.dropna().astype(int) == s.dropna()).all():
                    if s.max() < 10_000:
                        reg = s
                        break
        if hi.notna().sum() == 0 and n >= 2:
            hi = pd.to_numeric(col(1), errors="coerce")
        df_res = pd.DataFrame({"lo": lo, "hi": hi, "regnr": reg}).dropna(subset=["regnr"])
        df_res["lo"] = df_res["lo"].fillna(0).astype(int)
        df_res["hi"] = df_res["hi"].fillna(df_res["lo"]).astype(int)
        df_res["regnr"] = df_res["regnr"].astype(int)
        df_res = df_res[df_res["hi"] >= df_res["lo"]]
        return df_res.reset_index(drop=True)

    def _auto_map_from_intervals(self, intervals: pd.DataFrame) -> dict[str, int]:
        """Bygg {konto(str) -> regnr(int)} for konti i df_full vha intervalltabellen."""
        if "konto" not in self.df_full.columns:
            return {}
        out: dict[str, int] = {}
        konto_series = pd.to_numeric(self.df_full["konto"], errors="coerce").fillna(-1).astype(int)
        # Vi lar første treff vinne (typisk layout)
        assigned = pd.Series([False] * len(konto_series), index=self.df_full.index)
        for _, r in intervals.iterrows():
            lo, hi, rn = int(r["lo"]), int(r["hi"]), int(r["regnr"])
            mask = (~assigned) & (konto_series >= lo) & (konto_series <= hi)
            idx = konto_series[mask].index
            for ix in idx:
                k = str(int(konto_series.loc[ix]))
                if k not in out:
                    out[k] = rn
            assigned.loc[idx] = True
        return out

    def _with_regnskapslinjer_cols(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Returner df med 'regnr' (STR – for å unngå 2 desimaler i DataTable)
        og 'regnskapslinje' (STR) basert på self._konto2regnr og self._regnr2name.
        Hvis en regnr‑verdi inneholder tekst (f.eks. "510 - Utvikling"), hentes
        første tall ut via _extract_first_int.
        """
        out = df.copy()
        if "konto" not in out.columns:
            return out
        # gjør konto om til str med bare heltall (for mappingoppslag)
        ks = pd.to_numeric(out["konto"], errors="coerce").fillna(-1).astype(int).astype(str)
        rn = ks.map(self._konto2regnr).fillna("")
        def to_str_regnr(x):
            if x is None or x == "" or (isinstance(x, float) and pd.isna(x)):
                return ""
            num = _extract_first_int(x)
            return str(num) if num is not None else ""
        out["regnr"] = rn.apply(to_str_regnr)
        if self._regnr2name:
            out["regnskapslinje"] = rn.apply(
                lambda x: "" if x == "" else self._regnr2name.get(_extract_first_int(x) or 0, "")
            )
        else:
            out["regnskapslinje"] = ""
        return out

    def _map_regn(self):
        """Kjør intervall‑mapping (konto -> regnr/regnskapslinje) og oppdater tabellen."""
        if try_map_saldobalanse_to_regnskapslinjer is None:
            messagebox.showwarning("Ikke tilgjengelig", "Modul for regnskapslinje‑mapping mangler.", parent=self)
            return
        try:
            # kall tjenesten – finner 'F:\\Dokument\\Kildefiler' selv og husker sti i settings
            rows_df, _, meta = try_map_saldobalanse_to_regnskapslinjer(self.df_full)
            if rows_df is None:
                reason = (meta or {}).get("reason", "ukjent")
                messagebox.showwarning("Ingen mapping", f"Fant ikke kildefiler for regnskapslinjer.\nDetalj: {reason}", parent=self)
                return
            # rows_df har bare mappede rader; flett inn i originalen for å beholde evt. umappede
            lut = rows_df[["konto", "regnskapsnr", "regnskapsnavn"]].dropna(subset=["regnskapsnr"]).copy()
            # sørg for int konto
            lut["konto"] = pd.to_numeric(lut["konto"], errors="coerce").astype("Int64")
            # tolke regnskapsnr: kan være tekst som "510 - Utvikling"
            lut["regnskapsnr"] = lut["regnskapsnr"].apply(lambda v: _extract_first_int(v))
            base = self.df_full.copy()
            base["konto"] = pd.to_numeric(base["konto"], errors="coerce").astype("Int64")
            out = base.merge(lut, how="left", on="konto")
            # vis som tekst uten desimaler i tabellen
            out["regnr"] = out["regnskapsnr"].astype("Int64").astype("string")
            out["regnskapslinje"] = out["regnskapsnavn"].astype("string")
            # oppdater interne mappinger og lagre for senere bruk
            reg_map = lut.dropna(subset=["konto", "regnskapsnr"]).copy()
            for _, row in reg_map.iterrows():
                konto_val = row["konto"]
                regnr_val = row["regnskapsnr"]
                if pd.isna(konto_val) or pd.isna(regnr_val):
                    continue
                k_int = _extract_first_int(konto_val)
                r_int = _extract_first_int(regnr_val)
                if k_int is None or r_int is None:
                    continue
                self._konto2regnr[str(k_int)] = r_int
                if not pd.isna(row.get("regnskapsnavn")):
                    self._regnr2name[r_int] = str(row["regnskapsnavn"])
            self._save_regnr_map()
            # rydd og vis
            self.df_full = out.drop(columns=[c for c in ("regnskapsnr", "regnskapsnavn") if c in out.columns])
            cols = [c for c in ("konto", "kontonavn", "regnr", "regnskapslinje") if c in self.df_full.columns]
            cols += [c for c in self.df_full.columns if c not in cols and not str(c).startswith("__")]
            self.table.set_dataframe(self.df_full[cols], reset=True)
            self.table.refresh()
            mapped = int(lut["konto"].nunique())
            total = int(base["konto"].dropna().nunique())
            messagebox.showinfo("OK", f"Mappet {mapped} av {total} konti (lagret sti til kildefiler i settings).", parent=self)
        except Exception as exc:
            messagebox.showerror("Mapping feilet", f"{type(exc).__name__}: {exc}", parent=self)

    def _map_to_regnskapslinjer(self):
        """
        1) Sørger for stier (persist i meta.ui_prefs)
        2) Leser regnskapslinje‑navn + intervaller
        3) Mapper *manglende* konti (bevarer manuelle overstyringer)
        4) Lagrer per klient/år
        """
        files = self._ensure_regn_files()
        if not files:
            return
        rl_path, map_path = files
        try:
            self._regnr2name = self._read_regnskapslinjer_lut(rl_path)
            intervals = self._read_intervals(map_path)
            auto_map = self._auto_map_from_intervals(intervals)
        except Exception as exc:
            messagebox.showerror("Mapping‑feil", f"{type(exc).__name__}: {exc}", parent=self)
            return
        # Bevar manuelle → fyll bare hull
        added = 0
        for k, rn in auto_map.items():
            if k not in self._konto2regnr:
                self._konto2regnr[k] = int(rn)
                added += 1
        self._save_regnr_map()
        # Oppdater tabell
        self.df_full = self._with_regnskapslinjer_cols(self.df_full)
        cols = _preferred_order(self.df_full)
        self.table.set_dataframe(self.df_full[cols], reset=False)
        self.table.refresh()
        messagebox.showinfo("OK", "Mapping fullført og lagret pr. konto.\nDu kan overstyre enkeltkonti med «Sett regnr …».", parent=self)

    def _set_regnr_manual(self):
        """
        Manuell overstyring: bruker merket(e) rader i tabellen → spør etter regnr
        (tall eller '510 - Utvikling'), lagrer og oppdaterer skjermbildet.
        """
        rows = self.table.selected_rows()
        if rows is None or rows.empty:
            messagebox.showwarning("Velg rader", "Marker minst én rad i tabellen.", parent=self)
            return
        # foreslå dagens regnr hvis felles
        forslag = ""
        try:
            rns = set([str(x) for x in rows.get("regnr", "").astype(str).unique() if str(x).strip()])
            if len(rns) == 1:
                forslag = list(rns)[0]
        except Exception:
            pass
        s = simpledialog.askstring("Sett regnr", "Skriv regnr (f.eks. 510 eller '510 - Utvikling'):",
                                   initialvalue=forslag, parent=self)
        if not s:
            return
        m = re.search(r"(\d{1,4})", s)
        if not m:
            messagebox.showwarning("Ugyldig", "Fant ikke et tall i inputten.", parent=self)
            return
        rn = int(m.group(1))
        # oppdater mapping for valgte konti
        cnt = 0
        for _, r in rows.iterrows():
            k = _digits_only(r.get("konto"))
            if k:
                self._konto2regnr[str(int(k))] = rn
                cnt += 1
        self._save_regnr_map()
        # navneoppslag (valgfritt – vi bygger fra fil hvis vi har)
        if rn not in self._regnr2name:
            try:
                prefs = self._prefs_node()
                rp = prefs.get(_PREF_RL_PATH)
                if rp and Path(rp).exists():
                    self._regnr2name.update(self._read_regnskapslinjer_lut(Path(rp)))
            except Exception:
                pass
        # oppdater tabell
        self.df_full = self._with_regnskapslinjer_cols(self.df_full)
        cols = _preferred_order(self.df_full)
        self.table.set_dataframe(self.df_full[cols], reset=False)
        self.table.refresh()
        messagebox.showinfo("OK", f"Satt regnr={rn} på {cnt} konto(er).", parent=self)


# -------------------------------------------------------------
def main():
    a = _parse_args()
    App(client=a.client, year=a.year, source=a.source, vtype=a.vtype,
        modus=a.modus, konto=a.konto, bilagsnr=a.bilagsnr, adhoc_path=a.adhoc_path).mainloop()

if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
# src/app/gui/bilag_gui_tk.py
from __future__ import annotations
import argparse, os, re, sys, subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd

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
        from app.services.clients import get_clients_root
    except Exception:
        from services.clients import get_clients_root  # type: ignore

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
        raise ImportError("Fant ikke DataTable-klassen.")

    # SB→regnskap-mapping
    try:
        from app.services.sb_regnskapsmapping import (
            MapSources, map_saldobalanse_df, load_overrides, save_overrides, apply_overrides, read_regnskapslinjer
        )
    except Exception:
        from services.sb_regnskapsmapping import (  # type: ignore
            MapSources, map_saldobalanse_df, load_overrides, save_overrides, apply_overrides, read_regnskapslinjer
        )

    return (ensure_parquet_fresh, load_canonical_dataset, get_clients_root, DataTable,
            MapSources, map_saldobalanse_df, load_overrides, save_overrides, apply_overrides, read_regnskapslinjer)

(ensure_parquet_fresh, load_canonical_dataset, get_clients_root, DataTable,
 MapSources, map_saldobalanse_df, load_overrides, save_overrides, apply_overrides, read_regnskapslinjer) = _import_all()

# -------------------------------------------------------------
# Hjelpere
# -------------------------------------------------------------
def _digits_only(val) -> str | None:
    if val is None: return None
    s = re.sub(r"\D", "", str(val))
    return s if s else None

def _bilag_key(val) -> str | None:
    if val is None: return None
    return re.sub(r"[^0-9a-z]+", "", str(val).lower()) or None

def _konto_key_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").round(0).astype("Int64").astype(str)
    return s.astype(str).str.replace(r"\D", "", regex=True)

_CANON_PRIORITY = {
    "konto":     ["kontonr","kontonummer","konto nr","konto","accountno","account no","account"],
    "kontonavn": ["kontonavn","kontonamn","kontotekst","account name","accountname"],
    "dato":      ["dato","bokføringsdato","post date","postdate","transdate","date"],
    "bilagsnr":  [
        "bilagsnr","bilagsnummer","bilagsnum","bilag nr","bilag","voucher","voucher nr","voucherno",
        "voucher number","doknr","document no","document number","verifikasjonsnr","verifnr","journalnr"
    ],
    "tekst":     ["tekst","beskrivelse","description","post text","mottaker","faktura","narrative"],
}

def _preferred_order(df: pd.DataFrame) -> list[str]:
    cols = [c for c in df.columns if not str(c).startswith("__")]
    front = [c for c in ("konto","kontonavn","regnr","regnskapslinje") if c in cols]
    rest  = [c for c in cols if c not in front]
    return front + rest

# -------------------------------------------------------------
# Argparser
# -------------------------------------------------------------
def _parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--client", required=True)
    p.add_argument("--year", type=int, required=True)
    p.add_argument("--source", choices=["hovedbok","saldobalanse"], required=True)
    p.add_argument("--type", dest="vtype", choices=["interim","ao","versjon"], required=True)
    p.add_argument("--modus", choices=["analyse","uttrekk"], default="analyse")
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
                 konto: str|None=None, bilagsnr: str|None=None, adhoc_path: str|None=None):
        super().__init__()
        self.client, self.year, self.source, self.vtype, self.modus = client, int(year), source, vtype, modus
        self.prefilter_konto = _digits_only(konto) if konto else None
        self.prefilter_bkey  = _bilag_key(bilagsnr) if bilagsnr else None

        self.root_dir = get_clients_root()
        if not self.root_dir:
            messagebox.showerror("Mangler klient‑rot", "Fant ikke klient‑rot i settings.")
            self.destroy(); return

        self.title(f"Bilagsanalyse – {client} ({year}, {source}/{vtype})")
        self.geometry("1220x780"); self.minsize(1024, 660)

        # 1) Sørg for kanonisk datasett (Parquet/pickle) og last det
        try:
            dataset_path, manifest = ensure_parquet_fresh(self, self.root_dir, self.client, self.year, self.source, self.vtype)
            df = load_canonical_dataset(dataset_path)
            self.src_path = str(dataset_path)
            self._manifest = manifest
        except Exception as exc:
            messagebox.showerror("Indeksering/innlasting feilet", f"{type(exc).__name__}: {exc}", parent=self)
            self.destroy(); return

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
                u = uniq[0] if len(uniq)==1 else ""
                self._dataset_note = f" • Advarsel: datasettet har {len(uniq)} unik konto ({u}). Sjekk at aktiv HB‑fil er hele hovedboken."

        # 4) Re‑ordne kolonnene (konto/kononavn først)
        cols = _preferred_order(df_initial)
        df_initial = df_initial[cols].copy()

        # 5) Bygg UI
        self._build_ui(df_initial)

    # --------------------------- UI ---------------------------
    def _build_ui(self, df_initial: pd.DataFrame):
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=(8,2))
        ttk.Label(top, text="Søk i:").pack(side="left")

        choices = ["Alle kolonner"]
        for c in ["konto","kontonavn","dato","bilagsnr","tekst","regnr","regnskapslinje"]:
            if c in self.df_full.columns: choices.append(c)
        for c in self.df_full.columns:
            if not c.startswith("__") and c not in choices:
                choices.append(c)

        self.cmb_col = ttk.Combobox(self, state="readonly", width=22, values=choices)
        self.cmb_col.set("Alle kolonner")
        self.cmb_col.pack(in_=top, side="left", padx=(4,8))

        self.ent_q = ttk.Entry(top, width=36); self.ent_q.pack(side="left", padx=(0,6))
        self.ent_q.bind("<Return>", lambda e: self._apply_search()); self.ent_q.focus_set()

        ttk.Button(top, text="Søk", command=self._apply_search).pack(side="left", padx=(0,6))
        ttk.Button(top, text="Tøm", command=self._reset_view).pack(side="left", padx=(0,6))

        if self.source == "saldobalanse":
            ttk.Button(top, text="Map til regnskapslinjer …", command=self._map_regnskapslinjer).pack(side="left", padx=(6,0))
            ttk.Button(top, text="Sett regnr …", command=self._manual_set_regnr).pack(side="left", padx=(6,0))

        # Info-linje
        info_txt = f"Kilde: {Path(self.src_path).name}"
        if self.prefilter_info: info_txt += self.prefilter_info
        if self._dataset_note:  info_txt += self._dataset_note
        self.info = ttk.Label(self, text=info_txt, anchor="w")
        self.info.pack(fill="x", padx=8, pady=(2,2))

        # DataTable
        self.table = DataTable(self, df=df_initial, page_size=500)
        self.table.pack(fill="both", expand=True, padx=8, pady=(2,8))

        # Drilldown
        self._install_drilldown(self.table)

    def _reset_view(self):
        self.ent_q.delete(0, tk.END)
        self.cmb_col.set("Alle kolonner")
        cols = _preferred_order(self.df_full)
        self.table.set_dataframe(self.df_full[cols], reset=True)
        self.table.refresh()
        info_txt = f"Kilde: {Path(self.src_path).name}"
        if self._dataset_note: info_txt += self._dataset_note
        self.info.config(text=info_txt)

    # --------------------------- Søk --------------------------
    def _apply_search(self):
        col  = (self.cmb_col.get() or "Alle kolonner").strip()
        expr = (self.ent_q.get()  or "").strip()
        df   = self.df_full

        if not expr:
            self._reset_view(); return

        ops = ("==","!=" ,">=","<=" ,">","<")

        if col == "Alle kolonner":
            mask = pd.Series(False, index=df.index)
            for c in df.columns:
                if c.startswith("__"): continue
                if pd.api.types.is_string_dtype(df[c]) or df[c].dtype == "object":
                    mask |= df[c].astype(str).str.contains(expr, case=False, na=False, regex=False)
            cols = _preferred_order(df[mask])
            self.table.set_dataframe(df[mask][cols], reset=True); self.table.refresh(); return

        if col not in df.columns:
            messagebox.showwarning("Kolonne mangler", f"Fant ikke kolonnen «{col}»."); return

        # Bilagsnr → bruk normalisert nøkkel
        if col == "bilagsnr" and not expr.startswith(ops):
            key = _bilag_key(expr)
            keys = df["__bnr_key__"] if "__bnr_key__" in df.columns else df["bilagsnr"].map(_bilag_key)
            out = df[keys == key]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True); self.table.refresh(); return

        # Konto → prefiks som default, eksakt med '=='
        if col == "konto" and not expr.startswith(ops):
            key = re.sub(r"\D", "", expr)
            ks  = _konto_key_series(df[col])
            out = df[ks.str.startswith(key)]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True); self.table.refresh(); return

        # Operator-sammenligning
        op = None
        for t in ops:
            if expr.startswith(t):
                op = t; rhs = expr[len(t):].strip(); break

        s = df[col]
        if op is None:
            out = df[s.astype(str).str.contains(expr, case=False, na=False, regex=False)]
            self.table.set_dataframe(out[_preferred_order(out)], reset=True); self.table.refresh(); return

        # Dato?
        if pd.api.types.is_datetime64_any_dtype(s):
            rhs_dt = pd.to_datetime(rhs, errors="coerce")
            if pd.isna(rhs_dt):
                cmp = s.astype(str)
                if   op=="==": df2 = df[cmp == rhs]
                elif op=="!=": df2 = df[cmp != rhs]
                elif op==">=": df2 = df[cmp >= rhs]
                elif op=="<=": df2 = df[cmp <= rhs]
                elif op==">":  df2 = df[cmp >  rhs]
                elif op=="<":  df2 = df[cmp <  rhs]
            else:
                if   op=="==": df2 = df[s == rhs_dt]
                elif op=="!=": df2 = df[s != rhs_dt]
                elif op==">=": df2 = df[s >= rhs_dt]
                elif op=="<=": df2 = df[s <= rhs_dt]
                elif op==">":  df2 = df[s >  rhs_dt]
                elif op=="<":  df2 = df[s <  rhs_dt]
            self.table.set_dataframe(df2[_preferred_order(df2)], reset=True); self.table.refresh(); return

        # Tall
        try:
            rhs_num = pd.to_numeric(rhs)
            s_num   = pd.to_numeric(s, errors="coerce")
            if   op=="==": df2 = df[s_num == rhs_num]
            elif op=="!=": df2 = df[s_num != rhs_num]
            elif op==">=": df2 = df[s_num >= rhs_num]
            elif op=="<=": df2 = df[s_num <= rhs_num]
            elif op==">":  df2 = df[s_num >  rhs_num]
            elif op=="<":  df2 = df[s_num <  rhs_num]
        except Exception:
            cmp = s.astype(str).str.strip()
            if   op=="==": df2 = df[cmp == rhs]
            elif op=="!=": df2 = df[cmp != rhs]
            elif op==">=": df2 = df[cmp >= rhs]
            elif op=="<=": df2 = df[cmp <= rhs]
            elif op==">":  df2 = df[cmp >  rhs]
            elif op=="<":  df2 = df[cmp <  rhs]
        self.table.set_dataframe(df2[_preferred_order(df2)], reset=True); self.table.refresh()

    # ------------------------- Drilldown ----------------------
    def _install_drilldown(self, table_widget):
        if hasattr(table_widget, "bind_row_double_click"):
            if self.source == "saldobalanse":
                table_widget.bind_row_double_click(self._sb_to_hb)
            else:
                table_widget.bind_row_double_click(self._hb_to_hb)
            return

        # Fallback: tree-bind
        tree = getattr(table_widget, "tree", None)
        if tree is None: return
        def _on_dclick(_):
            iid = tree.focus() or (tree.selection()[0] if tree.selection() else None)
            if not iid: return
            values = tree.item(iid, "values") or []
            cols   = list(tree["columns"])
            row    = pd.Series({c: (values[i] if i < len(values) else None) for i,c in enumerate(cols)})
            (self._sb_to_hb if self.source=="saldobalanse" else self._hb_to_hb)(row)
        tree.bind("<Double-1>", _on_dclick, add="+")

    def _sb_to_hb(self, row: pd.Series):
        # hent konto robust fra rad
        konto_val = None
        for name in _CANON_PRIORITY["konto"]:
            for c in row.index:
                if c.casefold() == name.casefold():
                    konto_val = row[c]; break
            if konto_val is not None: break
        konto = _digits_only(konto_val)
        if not konto:
            messagebox.showinfo("Drilldown", "Fant ikke kontonummer i valgt rad."); return
        self._open_hovedbok(konto=konto)

    def _hb_to_hb(self, row: pd.Series):
        bnr = None
        for name in _CANON_PRIORITY["bilagsnr"]:
            for c in row.index:
                if c.casefold() == name.casefold():
                    bnr = row[c]; break
            if bnr is not None: break
        key = _bilag_key(bnr)
        if not key:
            messagebox.showinfo("Drilldown", "Fant ikke bilagsnummer i valgt rad.");
            return
        self._open_hovedbok(bilagsnr=str(bnr))

    def _open_hovedbok(self, konto: str|None=None, bilagsnr: str|None=None):
        env = dict(os.environ); env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH","")
        args = [sys.executable, "-m", "app.gui.bilag_gui_tk",
                "--client", self.client, "--year", str(self.year),
                "--source", "hovedbok", "--type", self.vtype, "--modus", "analyse"]
        if konto:    args += ["--konto",    konto]
        if bilagsnr: args += ["--bilagsnr", bilagsnr]
        subprocess.Popen(args, shell=False, cwd=str(SRC), env=env)

    # ------------------------- SB → Regnskap ----------------------
    def _map_regnskapslinjer(self):
        try:
            # Be om kildefiler
            p_reg = filedialog.askopenfilename(
                title="Velg «Regnskapslinjer.xlsx»",
                filetypes=[("Excel", "*.xlsx *.xls")]
            )
            if not p_reg: return
            p_int = filedialog.askopenfilename(
                title="Velg «Mapping standard kontoplan.xlsx»",
                filetypes=[("Excel", "*.xlsx *.xls")]
            )
            if not p_int: return

            sources = MapSources(Path(p_reg), Path(p_int))
            df_mapped, _defs = map_saldobalanse_df(self.df_full.copy(), sources)

            # last evt. overstyringer og anvend dem
            overrides = load_overrides(self.root_dir, self.client, self.year)
            df_mapped = apply_overrides(df_mapped, overrides)

            self.df_full = df_mapped.copy()
            cols = _preferred_order(self.df_full)
            self.table.set_dataframe(self.df_full[cols], reset=True)
            self.table.refresh()

            # lagre «basis»-mapping (uten å overskrive manuelle) – vi lagrer kun det som faktisk fikk regnr
            base_map = {str(int(k)): str(v) for k, v in zip(
                self.df_full["konto"].dropna().astype(int),
                self.df_full["regnr"].fillna("").astype(str)
            ) if v}
            # slå sammen med eksisterende overstyringer (overstyringer vinner)
            merged = {**base_map, **overrides}
            save_overrides(self.root_dir, self.client, self.year, merged)

            messagebox.showinfo("OK", "Mapping fullført og lagret pr. konto.\nDu kan overstyre enkeltkonti med «Sett regnr …».", parent=self)
        except Exception as exc:
            messagebox.showerror("Lesing av kildefiler feilet",
                                 f"{type(exc).__name__}: {exc}",
                                 parent=self)

    def _manual_set_regnr(self):
        rows = self.table.selected_rows()
        if rows is None or rows.empty:
            messagebox.showwarning("Velg rader", "Marker én eller flere konti i tabellen.", parent=self); return
        # foreslå fra første rad
        cur = str(rows.iloc[0].get("regnr") or "")
        new_reg = simpledialog.askstring("Sett regnr", "Regnr (f.eks. 585, 660 …):", initialvalue=cur, parent=self)
        if not new_reg: return
        new_reg = re.sub(r"\D", "", new_reg)
        if not new_reg:
            messagebox.showwarning("Ugyldig", "Regnr må være tall.", parent=self); return

        # Last regnskapslinje-tekst hvis vi kan
        try:
            # hent definisjoner via forrige mapping-kjøring (eller be om fil)
            # En enkel måte: spør ikke på nytt; vi setter bare regnr og lar regnskapslinje bli NaN hvis ikke krysset
            pass
        except Exception:
            pass

        # Oppdater i DataFrame og lagre overstyring
        konti = rows["konto"].dropna().astype(int).astype(str).tolist()
        self.df_full.loc[self.df_full["konto"].astype("Int64").astype(str).isin(konti), "regnr"] = str(new_reg)

        # regnskapslinje-navn: prøv å bevare hvis finnes i tabellen
        # (den sklår opp neste gang map kjøres – dette er mest for rask overstyring)
        self.table.set_dataframe(self.df_full[_preferred_order(self.df_full)], reset=True)

        # lagre/oppdater overstyringsfil
        overrides = load_overrides(self.root_dir, self.client, self.year)
        for k in konti:
            overrides[str(k)] = str(new_reg)
        save_overrides(self.root_dir, self.client, self.year, overrides)
        messagebox.showinfo("Lagret", f"Overstyrte {len(konti)} konto(er) til regnr {new_reg}.", parent=self)

# -------------------------------------------------------------
def main():
    a = _parse_args()
    App(client=a.client, year=a.year, source=a.source, vtype=a.vtype,
        modus=a.modus, konto=a.konto, bilagsnr=a.bilagsnr, adhoc_path=a.adhoc_path).mainloop()

if __name__ == "__main__":
    main()

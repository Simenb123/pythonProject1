# src/app/gui/revisjon_gui.py
# -----------------------------------------------------------------------------
# Revisjons-GUI: binder sammen Saldobalanse (SB) ↔ Hovedbok (HB)
# - Filter på kontointervall / søk
# - Klikk konto -> viser HB-linjer + statistikk
# - Kommentarer pr. konto (lagres) + tilfeldig utvalg (lagres)
# - Eksport til Excel
# -----------------------------------------------------------------------------
from __future__ import annotations

# sys.path for direkte kjøring
import sys
from pathlib import Path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import argparse
import json
import random
import statistics as stats
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

# services/widgets
from app.services.clients import get_clients_root, load_meta, processed_export_path
from app.services.versioning import resolve_active_raw_file, get_active_version
from app.services.io import read_raw, save_excel
from app.services.mapping import load_mapping, ensure_mapping_interactive, standardize_with_mapping
try:
    from app.gui.widgets.data_table import DataTable
except ModuleNotFoundError:
    try:
        from widgets.data_table import DataTable
    except ModuleNotFoundError:
        from data_table import DataTable

def _parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--client", required=True)
    p.add_argument("--year", type=int, required=True)
    return p.parse_args()


def _prefer_version(meta: dict, year: int, source: str) -> str:
    """Velg fornuftig type: ÅO hvis finnes, ellers Interim."""
    aid_ao = get_active_version(meta, year, source, "ao")
    if aid_ao:
        return "ao"
    aid_i = get_active_version(meta, year, source, "interim")
    return "interim" if aid_i else "ao"  # default hvis ingenting satt


def _load_std_df(root: Path, client: str, year: int, source: str, vtype: str) -> pd.DataFrame:
    """Les og standardiser DF for gitt kilde/type (sikrer mapping ved behov)."""
    meta = load_meta(root, client)
    p = resolve_active_raw_file(root, client, year, source, vtype, meta)
    if not p:
        # prøv andre type
        other = "ao" if vtype == "interim" else "interim"
        p = resolve_active_raw_file(root, client, year, source, other, meta)
        if not p:
            raise RuntimeError(f"Ingen aktiv {source}-versjon i {year}.")
        vtype = other

    df_raw, _ = read_raw(p)
    mp = load_mapping(root, client, year, source)
    if not mp:
        # liten sikring (dialog om noe mangler)
        mp = ensure_mapping_interactive(None, root, client, year, source, df_raw.head(200))
    df_std = standardize_with_mapping(df_raw, mp)
    return df_std


class App(tk.Tk):
    def __init__(self, client: str, year: int):
        super().__init__()
        self.title(f"Revisjonsdokumentasjon – {client} ({year})")
        self.geometry("1200x780"); self.minsize(1024, 660)

        self.client = client
        self.year = year
        self.root_dir = get_clients_root()
        if not self.root_dir:
            messagebox.showerror("Mangler klient-rot", "Fant ikke klient-rot i settings."); self.destroy(); return

        # 1) last SB (prefer ÅO), 2) last HB (prefer ÅO)
        meta = load_meta(self.root_dir, client)
        sb_vtype = _prefer_version(meta, year, "saldobalanse")
        hb_vtype = _prefer_version(meta, year, "hovedbok")

        try:
            self.df_sb = _load_std_df(self.root_dir, client, year, "saldobalanse", sb_vtype)
            self.df_hb = _load_std_df(self.root_dir, client, year, "hovedbok", hb_vtype)
        except Exception as exc:
            messagebox.showerror("Innlasting feilet", f"{type(exc).__name__}: {exc}"); self.destroy(); return

        # sørg for viktige kolonner
        for col in ("konto", "kontonavn"):
            if col not in self.df_sb.columns:
                messagebox.showerror("Saldobalanse mangler kolonner", f"Mangler '{col}' i saldobalanse."); self.destroy(); return
        for col in ("konto",):
            if col not in self.df_hb.columns:
                messagebox.showerror("Hovedbok mangler kolonner", f"Mangler '{col}' i hovedbok."); self.destroy(); return

        # normaliser typer
        self.df_sb["konto"] = pd.to_numeric(self.df_sb["konto"], errors="coerce").astype("Int64")
        self.df_hb["konto"] = pd.to_numeric(self.df_hb["konto"], errors="coerce").astype("Int64")

        # last/lager fil for notater/utvalg
        self.notes_path = (self.root_dir / self.client / "years" / f"{self.year}" /
                           "data" / "processed" / "analysis")
        self.notes_path.mkdir(parents=True, exist_ok=True)
        self.notes_file = self.notes_path / "revision_notes.json"
        self.notes = self._load_notes()

        self._build_ui()

    # ------------------ Notes/utvalg I/O ------------------
    def _load_notes(self) -> dict:
        if self.notes_file.exists():
            try:
                return json.loads(self.notes_file.read_text("utf-8"))
            except Exception:
                return {}
        return {}

    def _save_notes(self):
        tmp = self.notes_file.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.notes, indent=2, ensure_ascii=False), "utf-8")
        tmp.replace(self.notes_file)

    # ------------------ UI ------------------
    def _build_ui(self):
        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, padding=8)
        right = ttk.Frame(paned, padding=8)
        paned.add(left, weight=1)
        paned.add(right, weight=3)

        # LEFT: kontoliste + filter
        ttk.Label(left, text="Kontointervall").grid(row=0, column=0, sticky="w")
        self.k_lo = tk.StringVar(value="")
        self.k_hi = tk.StringVar(value="")
        ttk.Entry(left, textvariable=self.k_lo, width=8).grid(row=0, column=1, sticky="w")
        ttk.Entry(left, textvariable=self.k_hi, width=8).grid(row=0, column=2, sticky="w", padx=(6, 0))

        ttk.Label(left, text="Søk").grid(row=1, column=0, sticky="w")
        self.k_sok = tk.StringVar(value="")
        ttk.Entry(left, textvariable=self.k_sok, width=20).grid(row=1, column=1, columnspan=2, sticky="we")

        ttk.Button(left, text="Oppdater", command=self._refresh_accounts).grid(row=2, column=2, sticky="e", pady=(6, 4))

        self.acc_tree = ttk.Treeview(left, columns=("konto","kontonavn","inngående balanse","endring","utgående balanse"), show="headings", height=18)
        for c, w in zip(self.acc_tree["columns"], (80, 240, 120, 120, 120)):
            self.acc_tree.heading(c, text=c); self.acc_tree.column(c, width=w, anchor="w")
        self.acc_tree.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(6,0))
        self.acc_tree.bind("<<TreeviewSelect>>", self._on_select_account)

        left.grid_rowconfigure(3, weight=1)
        left.grid_columnconfigure(1, weight=1)

        self._refresh_accounts()

        # RIGHT: tabs (Transaksjoner / Statistikk / Notater / Utvalg)
        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)

        self.tab_tx = ttk.Frame(nb)
        self.tab_stat = ttk.Frame(nb)
        self.tab_notes = ttk.Frame(nb)
        self.tab_sample = ttk.Frame(nb)
        nb.add(self.tab_tx, text="Transaksjoner")
        nb.add(self.tab_stat, text="Statistikk")
        nb.add(self.tab_notes, text="Notat")
        nb.add(self.tab_sample, text="Tilfeldig utvalg")

        # Transaksjoner
        self.tx_table = DataTable(self.tab_tx, df=pd.DataFrame(), page_size=500)
        self.tx_table.pack(fill="both", expand=True)

        # Statistikk
        self.stat_txt = tk.StringVar(value="Velg en konto …")
        ttk.Label(self.tab_stat, textvariable=self.stat_txt, justify="left").pack(anchor="w", padx=8, pady=8)

        # Notater
        note_top = ttk.Frame(self.tab_notes); note_top.pack(fill="x", padx=8, pady=8)
        ttk.Label(note_top, text="Notat for konto:").pack(side="left")
        self.note_konto = tk.StringVar(value="")
        ttk.Label(note_top, textvariable=self.note_konto, font=("", 10, "bold")).pack(side="left", padx=(6,0))
        self.note_text = tk.Text(self.tab_notes, height=8)
        self.note_text.pack(fill="both", expand=True, padx=8, pady=(0,8))
        ttk.Button(self.tab_notes, text="Lagre notat", command=self._save_current_note).pack(anchor="e", padx=8, pady=(0,8))

        # Tilfeldig utvalg
        samp_top = ttk.Frame(self.tab_sample); samp_top.pack(fill="x", padx=8, pady=8)
        ttk.Label(samp_top, text="Antall:").pack(side="left")
        self.samp_n = tk.IntVar(value=10)
        ttk.Entry(samp_top, textvariable=self.samp_n, width=6).pack(side="left", padx=(6,12))
        ttk.Label(samp_top, text="Seed:").pack(side="left")
        self.samp_seed = tk.IntVar(value=42)
        ttk.Entry(samp_top, textvariable=self.samp_seed, width=8).pack(side="left", padx=(6,12))
        ttk.Button(samp_top, text="Trekk utvalg", command=self._draw_sample).pack(side="left")

        self.samp_table = DataTable(self.tab_sample, df=pd.DataFrame(), page_size=200)
        self.samp_table.pack(fill="both", expand=True, padx=8, pady=(0,8))

        ttk.Button(self.tab_sample, text="Eksporter rapport (Excel)", command=self._export_report)\
            .pack(anchor="e", padx=8, pady=(0,8))

    # ------------------ Left panel helpers ------------------
    def _refresh_accounts(self):
        try:
            df = self.df_sb.copy()
            if self.k_lo.get() or self.k_hi.get():
                lo = int(self.k_lo.get()) if self.k_lo.get() else -10**9
                hi = int(self.k_hi.get()) if self.k_hi.get() else 10**9
                df = df[(df["konto"] >= lo) & (df["konto"] <= hi)]
            if self.k_sok.get().strip():
                q = self.k_sok.get().lower()
                df = df[df["kontonavn"].astype(str).str.lower().str.contains(q, na=False) |
                        df["konto"].astype(str).str.contains(q, na=False)]
            # standardkolonnenavn for SB (hvis mappingen ga andre)
            cols = []
            for c in ("konto","kontonavn","inngående balanse","endring","utgående balanse"):
                if c in df.columns: cols.append(c)
            show = df[cols].fillna("")
            # fyll treet
            for iid in self.acc_tree.get_children(): self.acc_tree.delete(iid)
            for _, row in show.iterrows():
                vals = [row.get(c, "") for c in cols]
                self.acc_tree.insert("", "end", values=vals)
        except Exception as exc:
            messagebox.showerror("Feil i filter", f"{type(exc).__name__}: {exc}")

    def _on_select_account(self, *_):
        sel = self.acc_tree.selection()
        if not sel:
            return
        vals = self.acc_tree.item(sel[0], "values")
        # anta at første kolonne er konto, andre kontonavn
        try:
            konto = int(vals[0])
        except Exception:
            return
        self._load_account(konto, vals[1] if len(vals)>1 else "")

    # ------------------ Right panel actions ------------------
    def _load_account(self, konto: int, navn: str):
        self.note_konto.set(f"{konto} {navn}")
        # filter HB for konto
        dfk = self.df_hb[self.df_hb["konto"] == konto].copy()
        self.tx_table.set_dataframe(dfk)
        # stats
        try:
            antall = len(dfk)
            unike_bilag = len(dfk["bilagsnr"].dropna().unique()) if "bilagsnr" in dfk.columns else None
            bel = pd.to_numeric(dfk["beløp"], errors="coerce") if "beløp" in dfk.columns else pd.Series([], dtype=float)
            mmin = float(bel.min()) if not bel.empty else 0.0
            mmax = float(bel.max()) if not bel.empty else 0.0
            snitt = float(bel.mean()) if not bel.empty else 0.0
            med = float(bel.median()) if not bel.empty else 0.0
            self.stat_txt.set(
                f"Konto: {konto} {navn}\n"
                f"Transaksjoner: {antall:,}\n"
                f"Unike bilag: {unike_bilag if unike_bilag is not None else '–'}\n"
                f"Min: {mmin:,.2f}   Max: {mmax:,.2f}   Snitt: {snitt:,.2f}   Median: {med:,.2f}"
                .replace(",", " ")
            )
        except Exception:
            self.stat_txt.set("")

        # last eksisterende notat (om finnes)
        key = str(konto)
        note = self.notes.get(key, {}).get("note", "")
        self.note_text.delete("1.0","end")
        if note:
            self.note_text.insert("1.0", note)

        # vis evt. eksisterende utvalg
        sample_df = pd.DataFrame(self.notes.get(key, {}).get("sample", []))
        self.samp_table.set_dataframe(sample_df)

    def _save_current_note(self):
        sel = self.acc_tree.selection()
        if not sel:
            messagebox.showwarning("Ingen konto", "Velg en konto i saldobalansen.")
            return
        konto = str(self.acc_tree.item(sel[0], "values")[0])
        self.notes.setdefault(konto, {})
        self.notes[konto]["note"] = self.note_text.get("1.0","end").strip()
        self._save_notes()
        messagebox.showinfo("Lagret", "Notat lagret.")

    def _draw_sample(self):
        sel = self.acc_tree.selection()
        if not sel:
            messagebox.showwarning("Ingen konto", "Velg en konto i saldobalansen."); return
        konto = int(self.acc_tree.item(sel[0], "values")[0])
        dfk = self.df_hb[self.df_hb["konto"] == konto].copy()
        if dfk.empty:
            messagebox.showwarning("Tomt", "Ingen transaksjoner på valgt konto."); return
        n = max(1, int(self.samp_n.get()))
        seed = int(self.samp_seed.get())
        random.seed(seed)
        idx = list(dfk.index)
        if n > len(idx): n = len(idx)
        sel_idx = random.sample(idx, n)
        sample = dfk.loc[sel_idx].copy()
        # lagre i notes
        recs = sample.to_dict(orient="records")
        key = str(konto)
        self.notes.setdefault(key, {})
        self.notes[key]["sample"] = recs
        self.notes[key]["sample_meta"] = {"n": n, "seed": seed}
        self._save_notes()
        self.samp_table.set_dataframe(sample)

    def _export_report(self):
        # enkel rapport: SB-tabell (filtrert), notater + utvalg
        try:
            # hent konti som vises nå
            rows = [self.acc_tree.item(iid, "values") for iid in self.acc_tree.get_children()]
            df_accounts = pd.DataFrame(rows, columns=("konto","kontonavn","inngående balanse","endring","utgående balanse"))
            # bygg notat-tabell
            notes_rows = []
            for konto, d in self.notes.items():
                notes_rows.append({"konto": int(konto), "note": d.get("note",""), "sample_n": len(d.get("sample", []))})
            df_notes = pd.DataFrame(notes_rows)

            # Samples -> én ark per konto med utvalg
            out_stem = processed_export_path(self.root_dir, self.client, self.year, "Revisjonsrapport")
            out_path = Path(str(out_stem) + ".xlsx")
            with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
                df_accounts.to_excel(xw, index=False, sheet_name="Kontoliste")
                df_notes.to_excel(xw, index=False, sheet_name="Notater")
                # pr konto
                for konto, d in self.notes.items():
                    samp = pd.DataFrame(d.get("sample", []))
                    if not samp.empty:
                        # Excel-arknavn maks 31 tegn
                        sheet = f"Utvalg_{konto}"[:31]
                        samp.to_excel(xw, index=False, sheet_name=sheet)
            messagebox.showinfo("Ferdig", f"Rapport skrevet til:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Eksport feilet", f"{type(exc).__name__}: {exc}")


def main():
    args = _parse_args()
    App(args.client, args.year).mainloop()


if __name__ == "__main__":
    main()

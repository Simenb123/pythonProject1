# -*- coding: utf-8 -*-
# src/app/gui/master_import_gui.py
from __future__ import annotations
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import getpass

import pandas as pd

from app.services.clients import get_clients_root, list_clients
from app.services.registry import load_registry

def _s(x) -> str:
    try:
        return "" if pd.isna(x) else str(x)
    except Exception:
        return "" if x is None else str(x)

class MasterImportWindow(tk.Toplevel):
    """
    Master-import (klientinfo, orgnr, partner, bransje ...):
      1) Velg Excel (BHL-klientliste)
      2) Vi beregner diff mot registry (per klientmappe)
      3) Du velger hvilke klienter som skal oppdateres
      4) Lagre endringer
    """
    def __init__(self, master):
        super().__init__(master)
        self.title("Masterimport – klientinfo (orgnr-basert)")
        self.geometry("940x560"); self.minsize(820, 480)
        self.root = get_clients_root()
        self.reg = load_registry(self.root)

        # Topp
        top = ttk.Frame(self, padding=(10,8)); top.pack(fill="x")
        ttk.Button(top, text="Velg fil …", command=self._pick).pack(side="left")
        self.lbl = ttk.Label(top, text="(ingen valgt)"); self.lbl.pack(side="left", padx=(8,0))

        # Split
        paned = ttk.Panedwindow(self, orient="horizontal"); paned.pack(fill="both", expand=True, padx=10, pady=(6,10))
        left = ttk.Frame(paned); right = ttk.Frame(paned)
        paned.add(left, weight=1); paned.add(right, weight=3)

        ttk.Label(left, text="Klienter med endringer").pack(anchor="w")
        self.lb = tk.Listbox(left, selectmode="browse"); self.lb.pack(fill="both", expand=True, pady=(4,0))
        self.lb.bind("<<ListboxSelect>>", self._show_client)

        ttk.Label(right, text="felt  |  før  →  etter").pack(anchor="w")
        self.tree = ttk.Treeview(right, columns=("felt","før","etter"), show="headings", height=18)
        for c, w in zip(self.tree["columns"], (180, 320, 320)):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor=("w" if c!="felt" else "w"))
        self.tree.pack(fill="both", expand=True, pady=(4,0))

        # bunn
        bottom = ttk.Frame(self, padding=(10,8)); bottom.pack(fill="x")
        ttk.Button(bottom, text="Merk alle for valgt klient", command=self._mark_all_current).pack(side="left")
        ttk.Button(bottom, text="Lagre endringer", command=self._apply).pack(side="right")

        self._diff = pd.DataFrame()
        self._current_client = None

    def _pick(self):
        p = filedialog.askopenfilename(title="Velg master-klientliste (Excel)",
                                       filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        self.lbl.config(text=Path(p).name)
        try:
            dfm = pd.read_excel(p, engine="openpyxl", dtype="string")
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke lese fil: {type(exc).__name__}: {exc}", parent=self); return

        # diff mot registry
        folders = list_clients(self.root)
        self._diff = self.reg.import_clients_from_master(dfm, folders)
        self._fill_left()

    def _fill_left(self):
        self.lb.delete(0, tk.END)
        if self._diff is None or self._diff.empty:
            self.lb.insert(tk.END, "(ingen endringer funnet)")
            return
        for client in sorted(self._diff["client"].unique()):
            self.lb.insert(tk.END, client)
        self.lb.selection_clear(0, tk.END)
        if self.lb.size() > 0:
            self.lb.selection_set(0); self.lb.see(0); self._show_client()

    def _show_client(self, *_):
        sel = self.lb.curselection()
        if not sel: return
        client = self.lb.get(sel[0])
        self._current_client = client
        self.tree.delete(*self.tree.get_children())
        dfc = self._diff[self._diff["client"] == client]
        for _, r in dfc.iterrows():
            self.tree.insert("", "end", values=(_s(r.get("field")), _s(r.get("before")), _s(r.get("after"))))

    def _mark_all_current(self):
        # ikke nødvendig i denne versjonen (alle endringer for valgt klient er med),
        # men hook er her dersom du vil ha per-felt avhuking senere.
        messagebox.showinfo("Info", "Alle endringer for valgt klient tas med.")

    def _apply(self):
        if self._diff is None or self._diff.empty:
            messagebox.showinfo("Ingenting å gjøre", "Ingen endringer funnet."); return
        user = getpass.getuser()
        try:
            self.reg.apply_client_diffs(self._diff, who=user)
            self.reg.save()
            messagebox.showinfo("Lagret", "Endringer skrevet til _admin/registry.xlsx", parent=self)
        except Exception as exc:
            messagebox.showerror("Feil", f"Lagring feilet: {type(exc).__name__}: {exc}", parent=self)

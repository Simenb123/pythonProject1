from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess, sys
from pathlib import Path

from app.services.clients import (
    get_clients_root, load_meta, save_meta, list_years,
    open_or_create_year, default_year, set_default_year, set_year_datakilde
)
from app.gui.widgets.versions_panel import VersionsPanel
from app.services.io import read_raw
from app.services.versioning import resolve_active_raw_file
from app.services.mapping import ensure_mapping_interactive

BILAG_GUI = Path(__file__).with_name("bilag_gui_tk.py")  # du kan endre senere

class ClientHub(tk.Toplevel):
    def __init__(self, master: tk.Tk, client_name: str):
        super().__init__(master)
        self.title(f"Klienthub – {client_name}")
        self.resizable(False, False)

        self.client = client_name
        self.root_dir = getattr(master, "clients_root")
        self.meta = load_meta(self.root_dir, self.client)

        frm = ttk.Frame(self, padding=10); frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text=f"Klient: {self.client}", font=("", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")

        # ÅR
        years = list_years(self.root_dir, self.client)
        start_year = default_year(self.meta, years[-1] if years else 2025)
        ttk.Label(frm, text="Revisjonsår:").grid(row=1, column=0, sticky="w", pady=(8,2))
        self.year = tk.IntVar(value=start_year)
        self.year_cmb = ttk.Combobox(frm, values=years, textvariable=self.year, width=10, state="readonly")
        self.year_cmb.grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Åpne år …", command=self._open_year).grid(row=1, column=2, padx=6, sticky="w")
        self.year_cmb.bind("<<ComboboxSelected>>", lambda *_: self._on_year_change())

        # Datakilde
        ttk.Label(frm, text="Datakilde:").grid(row=2, column=0, sticky="w", pady=(8,2))
        self.source = tk.StringVar(value="hovedbok")
        ttk.Radiobutton(frm, text="Hovedbok", variable=self.source, value="hovedbok").grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(frm, text="Saldobalanse", variable=self.source, value="saldobalanse").grid(row=2, column=2, sticky="w")

        # Versjonspaneler
        self.hb_panel = VersionsPanel(frm, "hovedbok");     self.hb_panel.grid(row=3, column=0, columnspan=4, sticky="we")
        self.sb_panel = VersionsPanel(frm, "saldobalanse"); self.sb_panel.grid(row=4, column=0, columnspan=4, sticky="we")

        # Knapper
        btns = ttk.Frame(frm); btns.grid(row=5, column=0, columnspan=4, sticky="we", pady=(10,0))
        ttk.Button(btns, text="Analyse",        command=lambda: self._start_bilag("analyse")).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Bilagsuttrekk",  command=lambda: self._start_bilag("uttrekk")).grid(row=0, column=1, padx=4)
        ttk.Button(btns, text="Mapping …",      command=self._ensure_mapping_now).grid(row=0, column=2, padx=4)
        ttk.Button(btns, text="Åpne klientmappe", command=self._open_folder).grid(row=0, column=3, padx=4)

        self._on_year_change()

    # -------- helpers ----------
    def _open_year(self):
        y = tk.simpledialog.askinteger("Åpne nytt år", "Hvilket år vil du åpne/opprette?",
                                       parent=self, minvalue=2000, maxvalue=2100, initialvalue=self.year.get())
        if not y: return
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        years = list_years(self.root_dir, self.client)
        self.year_cmb["values"] = years
        self.year.set(y)
        self._on_year_change()
        messagebox.showinfo("År klart", f"Året {y} er opprettet.")

    def _on_year_change(self):
        y = int(self.year.get())
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        self.hb_panel.refresh(); self.sb_panel.refresh()

    def _open_folder(self):
        path = self.root_dir / self.client
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke åpne mappen: {exc}", parent=self)

    # -------- mapping nå ----------
    def _ensure_mapping_now(self):
        y = int(self.year.get())
        src = self.source.get()
        # hent aktiv versjonsfil
        from app.services.versioning import resolve_active_raw_file
        p = resolve_active_raw_file(self.root_dir, self.client, y, src, "interim", self.meta) \
            or resolve_active_raw_file(self.root_dir, self.client, y, src, "ao", self.meta)
        if not p:
            messagebox.showwarning("Mangler versjon", "Ingen aktiv versjon for valgt kilde.", parent=self); return
        df, _ = read_raw(p)
        ensure_mapping_interactive(self, self.root_dir, self.client, y, src, df.head(200))
        messagebox.showinfo("OK", "Mapping er lagret for dette året/kilden.", parent=self)

    # -------- start bilag-GUI ----
    def _start_bilag(self, modus: str):
        y = int(self.year.get())
        src = self.source.get()
        set_year_datakilde(self.meta, y, src); save_meta(self.root_dir, self.client, self.meta)
        # Løft mapping hvis mulig (minimer “første gang”-friksjon)
        from app.services.versioning import resolve_active_raw_file
        p = resolve_active_raw_file(self.root_dir, self.client, y, src, "interim", self.meta) \
            or resolve_active_raw_file(self.root_dir, self.client, y, src, "ao", self.meta)
        if p:
            df, _ = read_raw(p)
            try:
                ensure_mapping_interactive(self, self.root_dir, self.client, y, src, df.head(200))
            except Exception:
                pass

        args = [sys.executable, str(BILAG_GUI),
                f"--client={self.client}", f"--year={y}",
                f"--source={src}", f"--type=interim", f"--modus={modus}"]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte Bilag-GUI: {exc}", parent=self)

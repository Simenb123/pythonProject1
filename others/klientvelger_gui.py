# klientvelger_gui.py – 2025-06-03
from __future__ import annotations

import logging
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

import pandas as pd

from others import app_logging
from others.controllers import ClientController
from src.app.services.mapping_utils import FeltVelger
from src.app.gui.ui_theme import available_themes, load_theme, set_theme, init_style

BILAG_GUI = Path(__file__).with_name("bilag_gui_tk.py")
logger = logging.getLogger(__name__)


class KlientVelgerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Klientvelger"); self.resizable(False, False)

        # Meny – Tema ----------------------------------------------------
        menubar = tk.Menu(self)
        vis_menu = tk.Menu(menubar, tearoff=0)
        tema_menu = tk.Menu(vis_menu, tearoff=0)
        cur = load_theme()
        for t in available_themes():
            tema_menu.add_radiobutton(
                label=t, value=t, variable=tk.StringVar(value=cur), command=lambda n=t: set_theme(n)
            )
        vis_menu.add_cascade(label="Tema", menu=tema_menu)
        menubar.add_cascade(label="Vis", menu=vis_menu)
        self.config(menu=menubar)

        # Widgets --------------------------------------------------------
        ttk.Label(self, text="Søk klient:").grid(row=0, column=0, sticky="w")
        self.sok = tk.StringVar()
        ent = ttk.Entry(self, textvariable=self.sok, width=25)
        ent.grid(row=0, column=1, columnspan=3, sticky="we", pady=2)
        ent.bind("<KeyRelease>", self._filter)

        ttk.Label(self, text="Velg klient:").grid(row=1, column=0, sticky="w")
        self.cli_var = tk.StringVar()
        self.cli_cmb = ttk.Combobox(
            self, textvariable=self.cli_var, values=ClientController.list_clients(),
            state="readonly", width=30
        )
        self.cli_cmb.grid(row=1, column=1, columnspan=3, sticky="we", pady=2)
        self.cli_cmb.bind("<<ComboboxSelected>>", self._update_after_client)

        if last := ClientController.last_used():
            self.cli_cmb.set(last); self.after(50, self._update_after_client)

        ttk.Label(self, text="Datakilde:").grid(row=2, column=0, sticky="w")
        self.type_var = tk.StringVar(value="Hovedbok")
        ttk.Combobox(
            self, textvariable=self.type_var,
            values=["Hovedbok", "Saldobalanse"], state="readonly", width=27
        ).grid(row=2, column=1, columnspan=3, sticky="we", pady=2)

        ttk.Label(self, text="Bilagsfil:").grid(row=3, column=0, sticky="w")
        self.bilag = tk.StringVar()
        ttk.Entry(self, textvariable=self.bilag, width=28).grid(row=3, column=1, sticky="we")
        ttk.Button(self, text="Bla …", command=self._velg_fil)\
            .grid(row=3, column=2, sticky="e")

        ttk.Button(self, text="Analyse", command=self._analyse).grid(
            row=4, column=1, sticky="e", padx=2, pady=4
        )
        ttk.Button(self, text="Bilagsuttrekk", command=self._uttrekk).grid(
            row=4, column=2, sticky="w", padx=2, pady=4
        )
        ttk.Button(self, text="Mapping …", command=self._mapping_dialog).grid(
            row=4, column=3, sticky="w", padx=2, pady=4
        )

        ent.focus()

    # ------------------ helpers -----------------------------------------
    def _current_client(self) -> Optional[ClientController]:
        return ClientController(self.cli_var.get()) if self.cli_var.get() else None

    # ------------------ callbacks ---------------------------------------
    def _filter(self, *_):
        s = self.sok.get().lower()
        vals = [n for n in ClientController.list_clients() if s in n.lower()]
        self.cli_cmb["values"] = vals
        if vals:
            self.cli_cmb.current(0); self._update_after_client()

    def _update_after_client(self, *_):
        cli = self._current_client()
        if not cli:
            return
        meta = cli.load_meta()
        if t := meta.get("map_type"): self.type_var.set(t)
        f = meta.get("last_file")
        if f and Path(f).exists():
            self.bilag.set(f)

    def _velg_fil(self):
        p = filedialog.askopenfilename(
            title="Velg bilagsfil", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
        )
        if p: self.bilag.set(p)

    def _mapping_dialog(self):
        cli = self._current_client()
        if not cli or not self.bilag.get():
            messagebox.showerror("Feil", "Velg klient og fil først", parent=self); return
        src = Path(self.bilag.get())
        if not src.exists():
            messagebox.showerror("Feil", "Bilagsfil mangler", parent=self); return

        df = (
            pd.read_excel(src, engine="openpyxl", nrows=1)
            if src.suffix.lower() in (".xlsx", ".xls")
            else pd.read_csv(src, nrows=1)
        )
        defaults = cli.load_mapping()
        mp = FeltVelger(self, df, defaults).mapping()
        if mp:
            cli.save_mapping(mp)
            messagebox.showinfo("Lagret", "Mapping lagret.", parent=self)

    def _analyse(self):  self._start_gui("analyse")
    def _uttrekk(self):  self._start_gui("uttrekk")

    def _start_gui(self, modus: str):
        if not self.cli_var.get() or not self.bilag.get():
            messagebox.showerror("Feil", "Fyll ut alle felt"); return
        subprocess.Popen([sys.executable, str(BILAG_GUI), self.bilag.get()], shell=False)


# ---------------- main ----------------------------------------------------
if __name__ == "__main__":
    app_logging.configure()
    init_style()
    if len(sys.argv) > 1:
        logger.warning("Ignorerer CLI-argumenter til klientvelger.")
    KlientVelgerGUI().mainloop()

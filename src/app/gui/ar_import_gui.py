# src/app/gui/ar_import_gui.py
from __future__ import annotations

import sys
from pathlib import Path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import tkinter as tk
from tkinter import ttk

class ARImportWindow(tk.Toplevel):
    def __init__(self, master, clients_root: Path, client: str):
        super().__init__(master)
        self.title(f"AR‑import – {client}")
        self.geometry("520x180")
        ttk.Label(self, text="Denne dialogen er en plassholder.\n"
                             "Koble på din aksjonærregister‑modul her.").pack(fill="both", expand=True, padx=16, pady=16)
        ttk.Button(self, text="Lukk", command=self.destroy).pack(pady=(0,12))

# src/app/gui/client_info_gui.py
from __future__ import annotations

import sys, json
from pathlib import Path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import tkinter as tk
from tkinter import ttk, messagebox

FIELDS = [
    ("organisasjonsnummer", "Organisasjonsnummer"),
    ("partner",             "Partner"),
    ("klientnummer",        "Klientnummer"),
    ("bransjekode",         "Bransjekode"),
    ("bransjekodenavn",     "Bransjekodenavn"),
    ("selskapsform",        "Selskapsform"),
    ("bransje",             "Bransje"),
    ("kontaktperson",       "Kontaktperson"),
    ("epost",               "E‑post"),
    ("telefon",             "Telefon"),
    ("adresse",             "Adresse"),
    ("regnskapsslutt",      "Regnskapsår slutt (DD.MM)"),
    ("merknader",           "Merknader"),
]


class ClientInfoWindow(tk.Toplevel):
    def __init__(self, master, clients_root: Path, client: str):
        super().__init__(master)
        self.title(f"Klientinfo – {client}")
        self.geometry("640x520")
        self.resizable(True, True)

        self.root = Path(clients_root)
        self.client = client
        self.file = self.root / client / "client_info.json"

        self.vars = {k: tk.StringVar(value="") for k, _ in FIELDS}
        self._build_ui()
        self._load()

    def _build_ui(self):
        frm = ttk.Frame(self); frm.pack(fill="both", expand=True, padx=8, pady=8)
        row = 0
        for k, label in FIELDS:
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", pady=3)
            if k == "merknader":
                self.txt = tk.Text(frm, height=8)
                self.txt.grid(row=row, column=1, sticky="nsew", pady=3)
            else:
                ttk.Entry(frm, textvariable=self.vars[k]).grid(row=row, column=1, sticky="we", pady=3)
            row += 1

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(len(FIELDS)-1, weight=1)

        bar = ttk.Frame(self); bar.pack(fill="x", padx=8, pady=(0,8))
        ttk.Button(bar, text="Lagre", command=self._save).pack(side="right")
        ttk.Button(bar, text="Lukk", command=self.destroy).pack(side="right", padx=(0,6))

    def _load(self):
        if self.file.exists():
            try:
                d = json.loads(self.file.read_text("utf-8"))
                for k, _ in FIELDS:
                    if k == "merknader":
                        self.txt.delete("1.0","end")
                        self.txt.insert("1.0", d.get(k,""))
                    else:
                        self.vars[k].set(str(d.get(k, "")))
            except Exception as exc:
                messagebox.showerror("Feil", f"Kunne ikke lese klientinfo.\n{type(exc).__name__}: {exc}", parent=self)

    def _save(self):
        try:
            d = {k: self.vars[k].get().strip() for k, _ in FIELDS if k != "merknader"}
            d["merknader"] = self.txt.get("1.0","end").strip()
            self.file.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.file.with_suffix(".tmp")
            tmp.write_text(json.dumps(d, indent=2, ensure_ascii=False), "utf-8")
            tmp.replace(self.file)
            messagebox.showinfo("Lagret", "Klientinfo lagret.", parent=self)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke lagre.\n{type(exc).__name__}: {exc}", parent=self)

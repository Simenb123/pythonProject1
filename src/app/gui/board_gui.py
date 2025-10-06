# -*- coding: utf-8 -*-
# src/app/gui/board_gui.py
from __future__ import annotations
from pathlib import Path
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import pandas as pd

from app.services.clients import get_clients_root  # :contentReference[oaicite:13]{index=13}
from app.services.registry import ensure_client_org_dirs
from app.services.board import load_board, save_board, BoardMember
from app.services.audit import log_client

try:
    from app.gui.widgets.data_table import DataTable  # :contentReference[oaicite:14]{index=14}
except Exception:
    from widgets.data_table import DataTable

class BoardEditor(tk.Toplevel):
    def __init__(self, master, client: str):
        super().__init__(master)
        self.title(f"Styre – {client}")
        self.geometry("820x560"); self.minsize(740, 480)

        self.root_dir = get_clients_root()
        if not self.root_dir:
            messagebox.showerror("Mangler klient-rot", "Velg klient-rot i Launcher/Portal først.", parent=self)
            self.after(50, self.destroy); return

        self.client = client
        client_dir = Path(self.root_dir) / client
        ensure_client_org_dirs(client_dir)
        self.client_dir = client_dir

        self.df = load_board(self.client_dir)
        self.table = DataTable(self, df=self.df, page_size=500); self.table.pack(fill="both", expand=True, padx=10, pady=10)

        btns = ttk.Frame(self, padding=10); btns.pack(fill="x")
        ttk.Button(btns, text="Ny …", command=self._new).pack(side="left")
        ttk.Button(btns, text="Rediger …", command=self._edit).pack(side="left", padx=(6,0))
        ttk.Button(btns, text="Avslutt (sett til_dato)", command=self._end).pack(side="left", padx=(6,0))
        ttk.Button(btns, text="Slett", command=self._delete).pack(side="left", padx=(6,0))
        ttk.Button(btns, text="Lagre", command=self._save).pack(side="right")

    def _new(self):
        m = self._ask_member()
        if not m: return
        row = pd.Series({k: getattr(m, k) for k in ["navn","rolle","fra_dato","til_dato","kilde","oppdatert_av","oppdatert_tid"]})
        self.df = pd.concat([self.df, row.to_frame().T], ignore_index=True)
        self.table.set_dataframe(self.df, reset=True)

    def _edit(self):
        rows = self.table.selected_rows()
        if rows.empty:
            messagebox.showwarning("Velg rad", "Marker en rad du vil endre.", parent=self); return
        ix = rows.index[0]
        cur = rows.iloc[0].to_dict()
        m = self._ask_member(cur)
        if not m: return
        for k in ["navn","rolle","fra_dato","til_dato","kilde","oppdatert_av","oppdatert_tid"]:
            self.df.loc[ix, k] = getattr(m, k)
        self.table.set_dataframe(self.df, reset=True)

    def _end(self):
        rows = self.table.selected_rows()
        if rows.empty:
            messagebox.showwarning("Velg rad", "Marker en rad.", parent=self); return
        ix = rows.index[0]
        dt = simpledialog.askstring("Sett slutt", "Til-dato (YYYY-MM-DD):", parent=self) or ""
        self.df.loc[ix, "til_dato"] = dt
        self.table.set_dataframe(self.df, reset=True)

    def _delete(self):
        rows = self.table.selected_rows()
        if rows.empty: return
        if not messagebox.askyesno("Slett", "Slette markerte rader?", parent=self): return
        self.df = self.df.drop(rows.index).reset_index(drop=True)
        self.table.set_dataframe(self.df, reset=True)

    def _save(self):
        before = len(load_board(self.client_dir))
        save_board(self.client_dir, self.df)
        after = len(self.df)
        log_client(self.root_dir, self.client, area="board", action="save",
                   user="system", before={"rows": before}, after={"rows": after})
        messagebox.showinfo("Lagret", "Styre er lagret.", parent=self)

    def _ask_member(self, preset: dict | None = None) -> BoardMember | None:
        preset = preset or {}
        top = tk.Toplevel(self); top.title("Styre – rediger"); top.resizable(False, False); top.transient(self); top.grab_set()
        labels = [("navn","Navn"),("rolle","Rolle"),("fra_dato","Fra (YYYY-MM-DD)"),("til_dato","Til (YYYY-MM-DD)"),
                  ("kilde","Kilde"),("oppdatert_av","Oppdatert av"),("oppdatert_tid","Oppdatert tid (YYYY-MM-DD)")]
        vars = {}
        for r,(k,lab) in enumerate(labels):
            ttk.Label(top, text=lab).grid(row=r, column=0, sticky="w", padx=8, pady=4)
            v = tk.StringVar(value=preset.get(k,"")); ttk.Entry(top, textvariable=v, width=36).grid(row=r, column=1, sticky="w", padx=(0,8))
            vars[k] = v
        ok = {}
        def _ok():
            ok["m"] = BoardMember(**{k: vars[k].get().strip() for k,_ in labels}); top.destroy()
        ttk.Button(top, text="OK", command=_ok).grid(row=len(labels), column=1, sticky="e", padx=8, pady=8)
        ttk.Button(top, text="Avbryt", command=top.destroy).grid(row=len(labels), column=0, sticky="w", padx=8, pady=8)
        self.wait_window(top)
        return ok.get("m")

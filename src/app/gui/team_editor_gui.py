# src/app/gui/team_editor_gui.py
from __future__ import annotations

import sys
from pathlib import Path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

from app.services.registry import load_team, save_team, list_employees_df, import_employees_from_excel


class TeamEditor(tk.Toplevel):
    """Enkel redigering av team for klient (lagres i <klient>/team.json)."""
    def __init__(self, master, clients_root: Path, client: str):
        super().__init__(master)
        self.title(f"Team – {client}")
        self.geometry("720x480")
        self.resizable(True, True)

        self.root = Path(clients_root)
        self.client = client

        self.df = pd.DataFrame(columns=["email", "name", "role"])
        self._build_ui()
        self.refresh()

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self); top.pack(fill="both", expand=True, padx=8, pady=8)

        bar = ttk.Frame(top); bar.pack(fill="x")
        ttk.Button(bar, text="Legg til …", command=self._add).pack(side="left")
        ttk.Button(bar, text="Importer ansattliste …", command=self._import_employee_book).pack(side="left", padx=6)
        ttk.Button(bar, text="Fjern", command=self._remove).pack(side="left", padx=6)
        ttk.Button(bar, text="Lagre", command=self._save).pack(side="right")

        self.tree = ttk.Treeview(top, show="headings", columns=("user", "role"))
        self.tree.heading("user", text="Bruker")
        self.tree.heading("role", text="Rolle")
        self.tree.column("user", width=520, anchor="w")
        self.tree.column("role", width=120, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(8,0))

        sb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.place(in_=self.tree, relx=1.0, relheight=1.0, x=-1, y=0, anchor="ne")

    # ---------------- Data ----------------
    def refresh(self):
        t = load_team(self.root, self.client)
        rows = t.get("members", [])
        self.df = pd.DataFrame(rows, columns=["email", "name", "role"]).fillna("")
        self._fill_tree()

    def _fill_tree(self):
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for _, r in self.df.iterrows():
            name = str(r.get("name") or "").strip()
            email = str(r.get("email") or "").strip()
            role = str(r.get("role") or "editor").strip() or "editor"
            label = f"{name} [{email}]" if name else email
            self.tree.insert("", "end", values=(label, role))

    # ---------------- Actions ----------------
    def _save(self):
        out = []
        for _, r in self.df.iterrows():
            if not r.get("email"): continue
            out.append({
                "email": str(r.get("email")).strip().lower(),
                "name": str(r.get("name") or "").strip(),
                "role": str(r.get("role") or "editor").strip() or "editor",
            })
        save_team(self.root, self.client, {"members": out})
        messagebox.showinfo("Lagret", f"Lagret {len(out)} bruker(e) til team.json", parent=self)

    def _remove(self):
        sel = self.tree.selection()
        if not sel: return
        idx = [self.tree.index(i) for i in sel]
        self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True)
        self._fill_tree()

    def _import_employee_book(self):
        p = filedialog.askopenfilename(title="Velg ansattliste (Excel)",
                                       filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        try:
            n = import_employees_from_excel(Path(p))
            messagebox.showinfo("Import", f"Leste {n} ansatte til global ansattliste.", parent=self)
        except Exception as exc:
            messagebox.showerror("Feil", f"{type(exc).__name__}: {exc}", parent=self)

    def _add(self):
        df = list_employees_df()
        if df is None or df.empty:
            messagebox.showinfo("Ansattliste", "Ingen global ansattliste er lagret enda.\n"
                                               "Bruk «Importer ansattliste …».", parent=self); return
        PickUsers(self, df, on_done=self._add_rows)

    def _add_rows(self, picked: pd.DataFrame):
        if picked is None or picked.empty: return
        add = picked[["email","name"]].copy()
        add["role"] = "editor"
        current = self.df.copy()
        merged = pd.concat([current, add], ignore_index=True)
        merged = merged.drop_duplicates(subset=["email"]).reset_index(drop=True)
        self.df = merged
        self._fill_tree()


class PickUsers(tk.Toplevel):
    """Velg flere brukere fra ansattlisten."""
    def __init__(self, master, df: pd.DataFrame, on_done):
        super().__init__(master)
        self.title("Importer – velg rader")
        self.geometry("640x520"); self.transient(master); self.grab_set()
        self.on_done = on_done
        self.df = df.fillna("")

        qv = tk.StringVar(value="")
        top = ttk.Frame(self); top.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(top, text="Søk:").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(top, textvariable=qv); ent.grid(row=0, column=1, sticky="we")
        top.columnconfigure(1, weight=1)

        self.tree = ttk.Treeview(top, show="headings", columns=("user","role"), selectmode="extended")
        self.tree.heading("user", text="Bruker")
        self.tree.heading("role", text="Rolle")
        self.tree.column("user", width=460, anchor="w")
        self.tree.column("role", width=120, anchor="w")
        self.tree.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        top.rowconfigure(1, weight=1)
        sb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.place(in_=self.tree, relx=1.0, relheight=1.0, x=-1, y=0, anchor="ne")

        bar = ttk.Frame(self); bar.pack(fill="x", padx=8, pady=(0,8))
        ttk.Button(bar, text="Fjern alle", command=lambda: self.tree.selection_remove(*self.tree.get_children()))\
            .pack(side="left")
        ttk.Button(bar, text="Merk alle", command=lambda: self.tree.selection_set(*self.tree.get_children()))\
            .pack(side="left", padx=6)
        ttk.Button(bar, text="Avbryt", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="OK", command=self._ok).pack(side="right", padx=6)

        def refill(*_):
            q = qv.get().strip().lower()
            for iid in self.tree.get_children(): self.tree.delete(iid)
            rows = self.df
            if q:
                rows = rows[rows.apply(lambda r:
                    q in str(r.get("name","")).lower()
                    or q in str(r.get("email","")).lower()
                    or q in str(r.get("initials","")).lower(), axis=1)]
            for _, r in rows.iterrows():
                name = str(r.get("name") or "").strip()
                email = str(r.get("email") or "").strip().lower()
                init = str(r.get("initials") or "").strip()
                label = f"{name} [{email}]" if name else email
                if init: label = f"{label}  ({init})"
                iid = self.tree.insert("", "end", values=(label, "editor"))
                self.tree.set(iid, "user", label)
                self.tree.set(iid, "role", "editor")
                # stash e‑post i iid tags
                self.tree.item(iid, tags=(email,))
        refill()
        ent.bind("<KeyRelease>", refill)

    def _ok(self):
        rows = []
        for iid in self.tree.selection():
            label = self.tree.set(iid, "user")
            import re
            m = re.search(r"\[([^\]]+)\]", label)
            email = (m.group(1) if m else label).strip().lower()
            name = label.split(" [")[0].strip() if " [" in label else ""
            rows.append({"email": email, "name": name})
        self.on_done(pd.DataFrame(rows))
        self.destroy()

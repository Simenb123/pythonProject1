from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
from pathlib import Path

from app.services.versioning import list_versions, create_version, set_active_version, get_active_version
from app.services.clients import save_meta

class VersionsPanel(ttk.Frame):
    """Panel for versjoner av én kilde ("hovedbok" eller "saldobalanse")."""
    def __init__(self, master, source: str, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        assert source in {"hovedbok", "saldobalanse"}
        self.source = source

        ttk.Label(self, text=f"{source.capitalize()} – versjon", font=("", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(6,2), columnspan=3)

        self.var_type = tk.StringVar(value="interim")
        ttk.Radiobutton(self, text="Interim", variable=self.var_type, value="interim").grid(row=1, column=0, sticky="w")
        ttk.Radiobutton(self, text="ÅO",     variable=self.var_type, value="ao").grid(row=1, column=1, sticky="w")

        ttk.Label(self, text="Velg versjon:").grid(row=2, column=0, sticky="w", pady=(4,2))
        self.cmb_var = tk.StringVar(value="")
        self.cmb = ttk.Combobox(self, textvariable=self.cmb_var, state="readonly", width=44)
        self.cmb.grid(row=2, column=1, columnspan=2, sticky="we", pady=(4,2))

        btns = ttk.Frame(self); btns.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4,0))
        ttk.Button(btns, text="Ny …", command=self._ny_versjon).grid(row=0, column=0, padx=3)
        ttk.Button(btns, text="Aktiver", command=self._aktiver).grid(row=0, column=1, padx=3)

        self.grid_columnconfigure(2, weight=1)
        self.var_type.trace_add("write", lambda *_: self.refresh())

    def refresh(self):
        root = self.master.root_dir
        client = self.master.client
        year = int(self.master.year.get())
        vtype = self.var_type.get()

        vs = list_versions(root, client, year, self.source, vtype)
        items = [f"{v.id}  ({v.period_from}–{v.period_to})" + (f"  – {v.label}" if v.label else "") for v in vs]
        self.cmb["values"] = items

        aid = get_active_version(self.master.meta, year, self.source, vtype)
        if aid:
            for s in items:
                if s.startswith(aid + "  "):
                    self.cmb_var.set(s); break
        else:
            self.cmb_var.set("")

    def _ny_versjon(self):
        fr = simpledialog.askstring("Ny versjon", "Periode FRA (YYYY-MM-DD):", parent=self)
        to = simpledialog.askstring("Ny versjon", "Periode TIL (YYYY-MM-DD):", parent=self)
        label = simpledialog.askstring("Ny versjon", "Label (valgfri):", parent=self) or ""
        if not fr or not to: return

        p = filedialog.askopenfilename(title=f"Velg {self.source.capitalize()}-fil",
                                       filetypes=[("CSV/XLSX", "*.csv *.xlsx *.xls"), ("Alle filer", "*.*")])
        if not p: return

        v = create_version(self.master.root_dir, self.master.client, int(self.master.year.get()),
                           source=self.source, vtype=self.var_type.get(),
                           period_from=fr, period_to=to, label=label, src_file=Path(p), how="copy")
        messagebox.showinfo("Opprettet", f"Laget versjon:\n{v.id}", parent=self)
        self.refresh()

    def _aktiver(self):
        sel = self.cmb_var.get()
        if not sel:
            messagebox.showwarning("Velg", "Velg en versjon først.", parent=self); return
        vid = sel.split()[0]
        year = int(self.master.year.get())
        set_active_version(self.master.meta, year, self.source, self.var_type.get(), vid)
        save_meta(self.master.root_dir, self.master.client, self.master.meta)
        messagebox.showinfo("Aktivert",
                            f"Aktiv {self.source} ({'Interim' if self.var_type.get()=='interim' else 'ÅO'}):\n{vid}",
                            parent=self)

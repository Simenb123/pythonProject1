# a07_widgets.py
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
import csv

def fmt_amount(x: float) -> str:
    try: return f"{x:,.2f}".replace(",", " ").replace(".", ",")
    except Exception: return str(x)

class Table(ttk.Treeview):
    """Treeview med enkel sortering + CSV‑eksport."""
    def __init__(self, master, columns, **kwargs):
        ids = [c for c,_ in columns]
        super().__init__(master, columns=ids, show="headings", selectmode="extended", **kwargs)
        self._cols = columns
        for cid, header in columns:
            self.heading(cid, text=header, command=lambda c=cid: self._sort_by(c, False))
            self.column(cid, width=120, anchor=tk.W, stretch=True)
        self.tag_configure("ok", background="#e8f5e9")
        self.tag_configure("warn", background="#fff8e1")
        self.tag_configure("bad", background="#ffebee")
        self._data_cache = []
        self._formats = {}

        yscroll = ttk.Scrollbar(master, orient="vertical", command=self.yview)
        self.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def set_column_format(self, column_id, func): self._formats[column_id] = func
    def clear(self): self.delete(*self.get_children()); self._data_cache.clear()

    def insert_rows(self, rows):
        self.clear()
        for r in rows:
            values = []
            for cid,_ in self._cols:
                v = r.get(cid, "")
                if cid in self._formats:
                    v = self._formats[cid](v)
                values.append(v)
            self.insert("", "end", values=values, tags=r.get("_tags", []))
            self._data_cache.append(values)

    def export_csv(self, path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([h for _,h in self._cols])
            w.writerows(self._data_cache)

    def _sort_by(self, col, desc):
        data = [(self.set(k,col),k) for k in self.get_children("")]
        def to_num(s):
            ss = str(s).strip().replace(" ","").replace(".", "").replace(",", ".")
            try: return float(ss)
            except Exception: return s
        data.sort(key=lambda t: to_num(t[0]), reverse=desc)
        for i,(_,k) in enumerate(data): self.move(k,"",i)
        self.heading(col, command=lambda c=col: self._sort_by(c, not desc))

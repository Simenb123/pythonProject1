# -*- coding: utf-8 -*-
# src/app/gui/widgets/data_table.py
from __future__ import annotations
import math, re
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd

NBSP = "\u00A0"

_ID_LIKE_PATTERNS = [
    r"\bbilag\w*\s*(nr|no|num|nummer|nr\.)\b",
    r"\bvoucher\w*\s*(no|nr|num|number)\b",
    r"\bdok\w*\s*(nr|no|num|number)\b",
    r"\bfaktura\w*\s*(nr|no|num|nummer)\b",
    r"\binvoice\w*\s*(no|num|number)\b",
]

class DataTable(ttk.Frame):
    """
    Lett tabell for store datasett:
    - set_dataframe(df) erstatter datasettet
    - sortering ved klikk på kolonneheader
    - høyreklikk: Kopier / Eksporter / Autotilpass / Velg kolonner
    - egen knapp-bar: Autotilpass / Kolonner …
    - dra-og-slipp i header for å endre kolonnerekkefølge
    - skjuler interne kolonner som starter med "__"
    - ID-felt (konto, bilagsnr, regnr m.fl.) vises uten desimaler/tusenskiller
    - statuslinje med antall og summer
    - NYTT: dobbeltklikk på kolonneheader autosizer den ene kolonnen
    """
    def __init__(self, master, df: pd.DataFrame, page_size: int = 500):
        super().__init__(master)
        self.page_size = max(1, int(page_size))
        self.page = 0
        self._is_dragging_header = False
        self._drag_from_idx = None
        self._drag_from_col = None

        # Style: diskret striper
        self._style = ttk.Style(self)
        try:
            self._style.configure("DataTable.Treeview", rowheight=22)
            self._style.map("DataTable.Treeview", background=[("selected", "#347AE2")], foreground=[("selected", "white")])
            self._style.configure("DataTable.Treeview.Heading", padding=(4,2))
        except Exception:
            pass

        # Toppbar
        bar = ttk.Frame(self); bar.pack(fill="x", padx=8, pady=(6,4))
        ttk.Label(bar, text="Side-størrelse:").pack(side="left")
        self.spin_size = ttk.Spinbox(bar, from_=50, to=5000, increment=50, width=6)
        self.spin_size.set(str(self.page_size))
        self.spin_size.pack(side="left", padx=(4,6))
        ttk.Button(bar, text="Oppdater", command=self.refresh).pack(side="left", padx=(0,6))
        ttk.Button(bar, text="Autotilpass", command=self.autosize).pack(side="left", padx=(0,6))
        ttk.Button(bar, text="Kolonner …", command=self.choose_columns).pack(side="left")

        # Tabell
        self.tree = ttk.Treeview(self, show="headings", style="DataTable.Treeview")
        self.tree.pack(fill="both", expand=True, padx=8, pady=(2,6))

        # Scrollbars
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.place(in_=self.tree, relx=1.0, relheight=1.0, x=-1, y=0, anchor="ne")
        hsb.pack(fill="x", padx=8, pady=(0,6))

        # Status
        self.status = ttk.Label(self, anchor="e"); self.status.pack(fill="x", padx=8, pady=(0,6))

        # Meny
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Kopier valgte rader", command=self.copy_selection)
        self.menu.add_command(label="Eksporter visning (CSV)…", command=self.export_view)
        self.menu.add_separator()
        self.menu.add_command(label="Autotilpass kolonner", command=self.autosize)
        self.menu.add_command(label="Velg kolonner…", command=self.choose_columns)
        self.tree.bind("<Button-3>", self._popup_menu)

        # Header-drag for rekkefølge
        self.tree.bind("<ButtonPress-1>", self._on_header_press, add="+")
        self.tree.bind("<B1-Motion>", self._on_header_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_header_release, add="+")

        # Dobbeltklikk: enten autosize headerkolonne eller rad‑dblklikk‑callback
        self._row_dblclick_cb = None
        self.tree.bind("<Double-1>", self._on_double_click, add="+")

        # Init-data
        self.set_dataframe(df, reset=True)

    # ---------- Offentlig API ----------
    def set_dataframe(self, df: pd.DataFrame, reset: bool = False):
        self.df_full = df.reset_index(drop=True).copy()
        self.numeric_cols = [c for c in self.df_full.columns
                             if pd.api.types.is_numeric_dtype(self.df_full[c]) and not self._is_id_like(c)]
        if reset: self.page = 0
        self._build_columns(self._default_columns())
        self.refresh()

    def bind_row_double_click(self, cb): self._row_dblclick_cb = cb

    def selected_rows(self) -> pd.DataFrame:
        if getattr(self, "df_view", None) is None or self.df_view.empty:
            return self.df_view.iloc[0:0]
        idxs = []
        for iid in self.tree.selection():
            try:
                src_ix = int(self.tree.item(iid).get("tags", ["-1"])[0])
                idxs.append(src_ix)
            except Exception:
                pass
        return self.df_view.loc[idxs] if idxs else self.df_view.iloc[0:0]

    # ---------- Intern logikk ----------
    def _default_columns(self):
        cols = [c for c in self.df_full.columns if not str(c).startswith("__")]
        front = [c for c in ("konto", "kontonavn", "regnr", "regnskapslinje") if c in cols]
        rest  = [c for c in cols if c not in front]
        return front + rest

    def _is_id_like(self, colname: str) -> bool:
        n = str(colname).casefold()
        if n in {"konto","bilagsnr","regnr",
                 "leverandørnummer","leverandornummer","leverandørnr",
                 "kundenummer","kundenr"}:
            return True
        for pat in _ID_LIKE_PATTERNS:
            if re.search(pat, n):
                return True
        return False

    def _build_columns(self, columns):
        self.tree["columns"] = columns
        for c in columns:
            # heading-kommandoen sorterer bare hvis vi IKKE har dratt på headeren
            self.tree.heading(c, text=c, command=lambda cc=c: (None if self._is_dragging_header else self._sort_by(cc)))
            anchor = "e" if (c in self.numeric_cols) else "w"
            self.tree.column(c, width=140, stretch=True, anchor=anchor)

    def _sort_by(self, col):
        asc = not getattr(self, f"__sort__{col}", True)
        setattr(self, f"__sort__{col}", asc)
        try:
            self.df_full = self.df_full.sort_values(col, ascending=asc, kind="mergesort").reset_index(drop=True)
        except Exception:
            self.df_full = self.df_full.sort_values(col, ascending=asc, key=lambda s: s.astype(str)).reset_index(drop=True)
        self.refresh(keep_page=True)

    def refresh(self, keep_page: bool = False):
        try:
            self.page_size = max(1, int(self.spin_size.get()))
        except Exception:
            self.page_size = 500
        if not keep_page:
            self.page = 0

        self.df_view = self.df_full.reset_index(drop=True)
        total_pages = max(1, math.ceil(len(self.df_view) / self.page_size))
        if self.page >= total_pages: self.page = total_pages - 1
        i0, i1 = self.page * self.page_size, self.page * self.page_size + self.page_size
        page_df = self.df_view.iloc[i0:i1].copy()

        # fyll tabell
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for ix, row in page_df.iterrows():
            vals = [self._fmt(c, row[c]) for c in self.tree["columns"]]
            self.tree.insert("", "end", values=vals, tags=(str(ix),))

        self.autosize(sample_rows=min(50, len(page_df)))
        self._update_status()

    # --------- Formatering ----------
    def _fmt(self, col, val):
        n = str(col).casefold()
        if self._is_id_like(n):
            if pd.isna(val): return ""
            s = str(val).strip()
            # prøv å tolke som tall og fjerne tusenskiller/desimaler
            try:
                sn = s.replace(NBSP, "").replace(" ", "").replace("\u202F","").replace(",", ".")
                f = float(sn)
                if abs(f - round(f)) < 1e-9:
                    return str(int(round(f)))
            except Exception:
                pass
            # ellers: fjern bare ekstra whitespace
            return re.sub(r"\s+", " ", s)

        if col in self.numeric_cols:
            try:
                x = float(val)
                return f"{x:,.2f}".replace(",", " ").replace(".", ",")
            except Exception:
                return "" if pd.isna(val) else str(val)
        return "" if pd.isna(val) else str(val)

    def autosize(self, *, sample_rows: int = 50):
        """Autosize alle kolonner basert på header + et utvalg rader."""
        sample = min(sample_rows, len(self.df_view))
        for col in self.tree["columns"]:
            self._autosize_single(col, sample)

    def _autosize_single(self, col: str, sample_rows: int = 50):
        w = max(80, len(str(col))*8)
        # ta verdiene fra de første 'sample_rows' radene i treet
        children = self.tree.get_children()
        for i in range(min(sample_rows, len(children))):
            v = self.tree.set(children[i], col)
            w = max(w, min(680, int(len(str(v))*8)))
        self.tree.column(col, width=w)

    def choose_columns(self):
        cols = list(self.tree["columns"])
        top = tk.Toplevel(self); top.title("Velg kolonner"); top.resizable(False, False); top.transient(self); top.grab_set()
        info = ttk.Label(top, text="Hold Ctrl/Shift for multi‑valg. Velg de kolonnene du vil vise.", anchor="w")
        info.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 2))
        lb = tk.Listbox(top, selectmode="extended", height=16, width=36)
        for i, c in enumerate(cols):
            lb.insert(tk.END, c)
            lb.selection_set(i)  # forhåndsvelg alt
        lb.grid(row=1, column=0, padx=10, pady=6)
        def ok():
            sel = [cols[i] for i in lb.curselection()]
            if sel:
                self._build_columns(sel); self.refresh()
            top.destroy()
        ttk.Button(top, text="OK", command=ok).grid(row=2, column=0, sticky="e", padx=10, pady=(0,10))

    def copy_selection(self):
        rows = self.selected_rows()
        if rows is None or rows.empty: return
        txt = rows.to_csv(index=False, sep="\t")
        self.clipboard_clear(); self.clipboard_append(txt)

    def export_view(self):
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV", "*.csv")],
                                         title="Eksporter visning (CSV)")
        if not p: return
        self.df_view.to_csv(p, index=False, encoding="utf-8-sig")

    def _popup_menu(self, event):
        try: self.menu.tk_popup(event.x_root, event.y_root)
        finally: self.menu.grab_release()

    def _on_double_click(self, event):
        # heading-dobbeltklikk => autosize denne kolonnen
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            try:
                idx = int(self.tree.identify_column(event.x).replace("#", "")) - 1
                col = list(self.tree["columns"])[idx]
                self._autosize_single(col, sample_rows=100)
            except Exception:
                pass
            return
        # ellers: eventuelt rad-dobbeltklikk-callback
        if not self._row_dblclick_cb:
            return
        iid = self.tree.focus() or (self.tree.selection()[0] if self.tree.selection() else None)
        if not iid: return
        try:
            ix = int(self.tree.item(iid).get("tags", ["-1"])[0])
            row = self.df_view.iloc[ix]
            self._row_dblclick_cb(row)
        except Exception:
            pass

    # --------- Header-drag (endre kolonnerekkefølge) ----------
    def _on_header_press(self, e):
        if self.tree.identify_region(e.x, e.y) != "heading":
            self._drag_from_idx = None; self._drag_from_col = None; self._is_dragging_header = False; return
        colid = self.tree.identify_column(e.x)  # "#1", "#2", …
        try:
            self._drag_from_idx = int(colid.replace("#", "")) - 1
            self._drag_from_col = list(self.tree["columns"])[self._drag_from_idx]
        except Exception:
            self._drag_from_idx = None; self._drag_from_col = None

    def _on_header_motion(self, e):
        if self._drag_from_col is not None:
            self._is_dragging_header = True  # blokker sortering mens vi drar

    def _on_header_release(self, e):
        if not self._is_dragging_header or self._drag_from_col is None:
            self._is_dragging_header = False; self._drag_from_col = None; self._drag_from_idx = None; return
        if self.tree.identify_region(e.x, e.y) != "heading":
            self._is_dragging_header = False; self._drag_from_col = None; self._drag_from_idx = None; return
        try:
            to_idx = int(self.tree.identify_column(e.x).replace("#", "")) - 1
            cols = list(self.tree["columns"])
            if self._drag_from_idx is not None and 0 <= self._drag_from_idx < len(cols) and 0 <= to_idx < len(cols):
                moved = cols.pop(self._drag_from_idx)
                cols.insert(to_idx, moved)
                self._build_columns(cols); self.refresh(keep_page=True)
        finally:
            self._is_dragging_header = False; self._drag_from_col = None; self._drag_from_idx = None

    def _update_status(self):
        parts = [f"Rader: {len(self.df_view)}",
                 f" | Side {self.page+1} / {max(1, (len(self.df_view)-1)//self.page_size + 1)}"]
        sum_fields = [c for c in ["beløp", "endring", "inngående balanse", "utgående balanse"] if c in self.df_view.columns]
        if sum_fields:
            agg = []
            for c in sum_fields:
                try:
                    v = pd.to_numeric(self.df_view[c], errors="coerce").sum()
                    agg.append(f"{c}: {v:,.2f}".replace(",", " ").replace(".", ","))
                except Exception:
                    pass
            if agg: parts.append(" | Sum (filter): " + " | ".join(agg))
        self.status.config(text="".join(parts))

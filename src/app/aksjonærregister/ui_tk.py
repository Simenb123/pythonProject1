from __future__ import annotations
import os
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except Exception as e:
    raise RuntimeError("Tkinter er ikke tilgjengelig i dette miljøet.") from e

try:
    from PIL import Image, ImageTk  # type: ignore
    PIL_AVAILABLE = True
except Exception:
    Image = ImageTk = None  # type: ignore
    PIL_AVAILABLE = False

from . import settings as S
from .db import ensure_db, open_conn, search_companies, get_owners_full, list_columns
from .graph import render_graph   # nå HTML+SVG – ingen avhengigheter

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Aksjonærregister – søk og visualisering (DuckDB)")
        self.geometry("1200x840")
        self.conn = None
        self.graph_img = None

        self._build_menu()
        self._build_ui()
        self._open_db()

    # ---------- Meny ----------
    def _build_menu(self) -> None:
        m = tk.Menu(self)
        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="Velg CSV og bygg DB…", command=self._import_csv)
        fm.add_separator()
        fm.add_command(label="Avslutt", command=self.destroy)
        m.add_cascade(label="Fil", menu=fm)

        hm = tk.Menu(m, tearoff=0)
        hm.add_command(label="Vis DB-kolonner", command=self._show_columns)
        m.add_cascade(label="Hjelp", menu=hm)
        self.config(menu=m)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        row = ttk.Frame(self, padding=8); row.pack(fill=tk.X)
        ttk.Label(row, text="Søk:").pack(side=tk.LEFT)
        self.entry = ttk.Entry(row, width=40); self.entry.pack(side=tk.LEFT, padx=6)
        self.entry.bind("<Return>", lambda e: self._do_search())
        self.by = tk.StringVar(value="navn")
        ttk.Radiobutton(row, text="Navn", variable=self.by, value="navn").pack(side=tk.LEFT)
        ttk.Radiobutton(row, text="Orgnr", variable=self.by, value="orgnr").pack(side=tk.LEFT)
        ttk.Button(row, text="Søk", command=self._do_search).pack(side=tk.LEFT, padx=6)

        cols = ("company_orgnr", "company_name")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        self.tree.heading("company_orgnr", text="Orgnr")
        self.tree.heading("company_name",  text="Selskap")
        self.tree.column("company_orgnr", width=160, anchor=tk.W)
        self.tree.column("company_name",  width=560, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=False, padx=8, pady=(6, 8))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        ctrl = ttk.Frame(self, padding=(8, 0)); ctrl.pack(fill=tk.X)
        self.sel_label = ttk.Label(ctrl, text="Velg et selskap ovenfor…"); self.sel_label.pack(side=tk.LEFT)
        self.mode = tk.StringVar(value="both")
        ttk.Radiobutton(ctrl, text="Oppstrøms", variable=self.mode, value="up").pack(side=tk.RIGHT)
        ttk.Radiobutton(ctrl, text="Nedstrøms", variable=self.mode, value="down").pack(side=tk.RIGHT)
        ttk.Radiobutton(ctrl, text="Begge", variable=self.mode, value="both").pack(side=tk.RIGHT)

        depth = ttk.Frame(self, padding=(8, 0)); depth.pack(fill=tk.X)
        ttk.Label(depth, text="Dybde opp:").pack(side=tk.LEFT)
        self.spin_up = tk.Spinbox(depth, from_=0, to=8, width=3)
        self.spin_up.delete(0, tk.END); self.spin_up.insert(0, str(S.MAX_DEPTH_UP))
        self.spin_up.pack(side=tk.LEFT, padx=(4,12))
        ttk.Label(depth, text="Dybde ned:").pack(side=tk.LEFT)
        self.spin_down = tk.Spinbox(depth, from_=0, to=8, width=3)
        self.spin_down.delete(0, tk.END); self.spin_down.insert(0, str(S.MAX_DEPTH_DOWN))
        self.spin_down.pack(side=tk.LEFT, padx=(4,12))
        ttk.Button(depth, text="Vis orgkart", command=self._show_graph).pack(side=tk.RIGHT)

        owners_frame = ttk.Frame(self, padding=8); owners_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(owners_frame, text="Eiere i valgt selskap (fra CSV)", font=("", 10, "bold")).pack(anchor="w")

        # NB: alle 8 kolonner vises – disse kommer alltid fra DB nå
        columns = ("owner_orgnr","owner_name","share_class","owner_country",
                   "owner_zip_place","shares_owner_num","shares_company_num","ownership_pct")
        headings = ("Eier orgnr/fødselsår","Eier navn","Aksjeklasse","Landkode",
                    "Postnr/sted","Antall aksjer","Antall aksjer selskap","Eierandel %")

        self.owners = ttk.Treeview(owners_frame, columns=columns, show="headings")
        for c, h in zip(columns, headings):
            self.owners.heading(c, text=h)
            w = 180 if c in ("owner_orgnr","owner_name","owner_zip_place") else 120
            if c in ("shares_owner_num","shares_company_num","ownership_pct"):
                self.owners.column(c, width=160, anchor=tk.E)
            else:
                self.owners.column(c, width=w, anchor=tk.W)

        xscroll = ttk.Scrollbar(owners_frame, orient="horizontal", command=self.owners.xview)
        yscroll = ttk.Scrollbar(owners_frame, orient="vertical",   command=self.owners.yview)
        self.owners.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        self.owners.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        yscroll.pack(fill=tk.Y, side=tk.RIGHT)
        xscroll.pack(fill=tk.X)

        self.canvas = ttk.Label(self)     # beholder plass til fremtidig PNG-visning
        self.canvas.pack(fill=tk.BOTH, expand=False, padx=8, pady=8)

    # ---------- DB-hjelpere ----------
    def _open_db(self) -> None:
        if not os.path.exists(S.DB_PATH) and os.path.exists(S.CSV_PATH):
            try:
                ensure_db(S.CSV_PATH, S.DB_PATH, force=True)
            except Exception as e:
                messagebox.showerror("Importfeil", str(e)); return
        try:
            self.conn = open_conn()
        except Exception as e:
            messagebox.showerror("Feil ved åpning", str(e))

    def _import_csv(self) -> None:
        path = filedialog.askopenfilename(title="Velg CSV", filetypes=[("CSV", "*.csv"), ("Alle filer", "*.*")])
        if not path: return
        try:
            ensure_db(path, S.DB_PATH, force=True)
            if self.conn: self.conn.close()
            self.conn = open_conn()
            messagebox.showinfo("Ferdig", "Database er bygget og klar!")
        except Exception as e:
            messagebox.showerror("Importfeil", str(e))

    # ---------- Actions ----------
    def _show_columns(self) -> None:
        if not self.conn:
            messagebox.showinfo("Ingen DB", "Åpne/bygg databasen først."); return
        cols = list_columns(self.conn)
        messagebox.showinfo("DB-kolonner", "\n".join(cols) if cols else "Fant ingen kolonner.")

    def _do_search(self) -> None:
        if not self.conn:
            messagebox.showwarning("Ingen DB", "Åpne/bygg databasen først."); return
        term = self.entry.get().strip()
        rows = search_companies(self.conn, term, self.by.get())
        for i in self.tree.get_children(): self.tree.delete(i)
        for orgnr, name in rows: self.tree.insert("", tk.END, values=(orgnr, name))

    def _on_select(self, _evt=None) -> None:
        sel = self.tree.selection()
        if not sel: return
        orgnr, name = self.tree.item(sel[0], "values")
        self.sel_label.config(text=f"Valgt: {name} ({orgnr})")
        self._load_owners(orgnr)

    def _fmt_text(self, v) -> str: return "" if v is None else str(v)

    def _fmt_shares(self, v) -> str:
        try:
            if v is None: return ""
            f = float(v)
            if abs(f - round(f)) < 0.005: return f"{int(round(f)):,}".replace(",", " ")
            return f"{f:,.2f}".replace(",", " ")
        except Exception:
            return str(v) if v is not None else ""

    def _fmt_pct(self, v) -> str:
        try:    return "" if v is None else f"{float(v):.2f}"
        except Exception: return str(v) if v is not None else ""

    def _load_owners(self, company_orgnr: str) -> None:
        if not self.conn: return
        rows = get_owners_full(self.conn, company_orgnr)
        for i in self.owners.get_children(): self.owners.delete(i)
        for r in rows:
            vals = list(r)
            vals[0] = self._fmt_text(vals[0])   # owner_orgnr
            vals[1] = self._fmt_text(vals[1])   # owner_name
            vals[2] = self._fmt_text(vals[2])   # share_class
            vals[3] = self._fmt_text(vals[3])   # owner_country
            vals[4] = self._fmt_text(vals[4])   # owner_zip_place
            vals[5] = self._fmt_shares(vals[5]) # shares_owner_num
            vals[6] = self._fmt_shares(vals[6]) # shares_company_num
            vals[7] = self._fmt_pct(vals[7])    # ownership_pct
            self.owners.insert("", tk.END, values=tuple(vals))

    def _show_graph(self) -> None:
        if not self.conn:
            messagebox.showwarning("Ingen DB", "Åpne/bygg databasen først."); return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Ingen valgt", "Velg et selskap fra listen først."); return
        orgnr, name = self.tree.item(sel[0], "values")
        try:
            up = int(self.spin_up.get()); down = int(self.spin_down.get())
        except Exception:
            up, down = S.MAX_DEPTH_UP, S.MAX_DEPTH_DOWN

        path = render_graph(self.conn, orgnr, name, mode=self.mode.get(), max_up=up, max_down=down)
        if not path:
            messagebox.showwarning("Ingen graf", "Kunne ikke lage orgkart (uventet feil).")
            return

        if path.lower().endswith(".html"):
            messagebox.showinfo("Graf generert", "Interaktiv orgkart-HTML er åpnet i nettleseren.")
        elif path.lower().endswith(".png"):
            # ikke i bruk nå, men behold fallback:
            if not PIL_AVAILABLE:
                messagebox.showinfo("Graf generert", f"Lagret PNG: {path}")
                return
            try:
                from PIL import Image, ImageTk  # sikkerhetsnett
                img = Image.open(path)
                max_w = self.winfo_width() - 40; max_h = 340
                w, h = img.size; scale = min(max_w / w, max_h / h, 1.0)
                if scale < 1.0: img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
                self.graph_img = ImageTk.PhotoImage(img)
                self.canvas.configure(image=self.graph_img)
            except Exception as e:
                messagebox.showerror("Visningsfeil", str(e))

if __name__ == "__main__":
    app = App()
    app.mainloop()

# -*- coding: utf-8 -*-
"""
Faktura-GUI med PDF-visning, halvtransparent gul highlighting, tooltips og batch.
Kjør direkte i PyCharm: Run -> run_invoice_gui.py
"""

import os, sys, base64, platform, logging, csv
from decimal import Decimal
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- bootstrap ---
BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

# demp pdfminer-bråk
for name in ("pdfminer", "pdfminer.pdfpage", "pdfminer.pdfinterp", "pdfminer.pdfdocument"):
    logging.getLogger(name).setLevel(logging.ERROR)

# valgfritt: Pillow for pen alfa-highlighting (fallback til stipple hvis ikke tilgjengelig)
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

import fitz  # PyMuPDF

from app.dokumentreader import extractors
from app.dokumentreader.invoice_reader import build_invoice_model
from app.dokumentreader.parsers import parse_line_items_from_text
from app.dokumentreader.models import model_to_json_text
from app.dokumentreader.highlighter import build_invoice_highlights  # bygger highlight-map


def _open_with_default_app(path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            import subprocess; subprocess.run(["open", path])
        else:
            import subprocess; subprocess.run(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Kunne ikke åpne fil", str(e))


class TextHandler(logging.Handler):
    def __init__(self, widget: tk.Text): super().__init__(); self.widget = widget
    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self.widget.after(0, lambda: (self.widget.insert("end", msg + "\n"), self.widget.see("end")))


class _Tooltip:
    def __init__(self, canvas: tk.Canvas):
        self.canvas = canvas
        self.tw: tk.Toplevel | None = None
    def show(self, x: int, y: int, text: str):
        self.hide()
        self.tw = tk.Toplevel(self.canvas)
        self.tw.wm_overrideredirect(True)
        self.tw.attributes("-topmost", True)
        lbl = tk.Label(self.tw, text=text, bg="#ffffe0", relief="solid", bd=1, padx=6, pady=2)
        lbl.pack()
        self.tw.wm_geometry(f"+{self.canvas.winfo_rootx()+x+12}+{self.canvas.winfo_rooty()+y+8}")
    def hide(self):
        if self.tw:
            try: self.tw.destroy()
            except Exception: pass
            self.tw = None


# ---------------- PDF Viewer ----------------
class PdfViewer(ttk.Frame):
    """PDF viewer med scroll/zoom, gule highlights, tooltips og hopp/puls."""
    def __init__(self, master):
        super().__init__(master)
        self.doc: fitz.Document | None = None
        self.path: str | None = None
        self.page_index = 0
        self.zoom = 1.2
        self.photo: tk.PhotoImage | None = None

        # highlight-data
        self.page_map = {}      # {page: [(rect,label,color,kind,key), ...]}
        self.key_hits = {}      # {key: [(page, rect), ...]}
        self.key_values = {}    # {key: value-string}
        self._hover_key: str | None = None
        self._overlay_imgs = []  # holder referanser til tk-bilder (Pillow) så de ikke GC-es
        self.tooltip = _Tooltip(None)  # init etter canvas

        # Toolbar
        bar = ttk.Frame(self); bar.pack(fill="x", padx=5, pady=5)
        self.btn_prev = ttk.Button(bar, text="◀ Forrige", width=10, command=self.prev_page, state="disabled"); self.btn_prev.pack(side="left")
        self.btn_next = ttk.Button(bar, text="Neste ▶",  width=10, command=self.next_page, state="disabled"); self.btn_next.pack(side="left", padx=(5, 10))
        ttk.Label(bar, text="Side:").pack(side="left")
        self.var_page = tk.StringVar(value="0 / 0")
        self.ent_page = ttk.Entry(bar, textvariable=self.var_page, width=12); self.ent_page.pack(side="left", padx=5)
        self.ent_page.bind("<Return>", self._goto_page_from_entry)
        ttk.Button(bar, text="−", width=3, command=lambda: self.change_zoom(0.85)).pack(side="left", padx=(10, 2))
        ttk.Button(bar, text="+", width=3, command=lambda: self.change_zoom(1.15)).pack(side="left", padx=(0, 10))
        ttk.Button(bar, text="Tilpass bredde", command=self.fit_width).pack(side="left")
        ttk.Button(bar, text="100%", command=self.reset_zoom).pack(side="left", padx=5)

        self.show_values = tk.BooleanVar(value=True)
        self.show_labels = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="Felter",   variable=self.show_values, command=self.render).pack(side="left", padx=(10,0))
        ttk.Checkbutton(bar, text="Etiketter",variable=self.show_labels, command=self.render).pack(side="left")

        # Canvas
        wrap = ttk.Frame(self); wrap.pack(fill="both", expand=True, padx=5, pady=(0,5))
        self.canvas = tk.Canvas(wrap, background="#f0f0f0"); self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview); self.vsb.pack(side="left", fill="y")
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview); self.hsb.pack(fill="x")
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<Button-4>",  self._on_mouse_wheel_linux)
        self.canvas.bind("<Button-5>",  self._on_mouse_wheel_linux)
        self._fit_width_mode = False
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.tooltip = _Tooltip(self.canvas)

    # ----- API -----
    def load(self, path: str):
        if self.doc:
            try: self.doc.close()
            except Exception: pass
        self.doc = fitz.open(path)
        self.path = path
        self.page_index = 0; self.zoom = 1.2; self._fit_width_mode = False
        self._update_nav_state(); self.render()

    def set_highlights(self, page_map: dict | None, key_hits: dict | None, key_values: dict | None):
        self.page_map = page_map or {}
        self.key_hits = key_hits or {}
        self.key_values = key_values or {}
        self.render()

    def jump_to(self, key: str):
        hits = self.key_hits.get(key) or []
        if not hits: return
        page, rect = hits[0]
        if page != self.page_index:
            self.page_index = page
            self.render()
        self._flash(rect)

    def preview_key(self, key: str | None, on: bool):
        self._hover_key = key if on else None
        self.render()

    # ----- Navigasjon / zoom -----
    def next_page(self):
        if self.doc and self.page_index < self.doc.page_count-1: self.page_index += 1; self.render()
    def prev_page(self):
        if self.doc and self.page_index > 0: self.page_index -= 1; self.render()
    def change_zoom(self, f: float):
        self._fit_width_mode = False; self.zoom = max(0.2, min(6.0, self.zoom * f)); self.render()
    def reset_zoom(self):
        self._fit_width_mode = False; self.zoom = 1.0; self.render()
    def fit_width(self):
        self._fit_width_mode = True; self.render()

    # ----- intern -----
    def _update_nav_state(self):
        if not self.doc:
            self.btn_prev.config(state="disabled"); self.btn_next.config(state="disabled"); self.var_page.set("0 / 0"); return
        self.btn_prev.config(state=("normal" if self.page_index>0 else "disabled"))
        self.btn_next.config(state=("normal" if self.page_index< self.doc.page_count-1 else "disabled"))
        self.var_page.set(f"{self.page_index+1} / {self.doc.page_count}")

    def _goto_page_from_entry(self, _evt=None):
        if not self.doc: return
        try:
            idx = int(self.var_page.get().split("/")[0].strip()) - 1
            if 0 <= idx < self.doc.page_count: self.page_index = idx; self.render()
        except Exception: pass

    def _on_mouse_wheel(self, e): self.canvas.yview_scroll(int(-e.delta/120), "units")
    def _on_mouse_wheel_linux(self, e): self.canvas.yview_scroll(-3 if e.num==4 else 3, "units")
    def _on_canvas_configure(self, _evt):
        if self._fit_width_mode:
            self._compute_fit_width_zoom(); self.render(redraw_only=True)
    def _compute_fit_width_zoom(self):
        if not self.doc: return
        page = self.doc[self.page_index]; cw = max(50, self.canvas.winfo_width())
        self.zoom = max(0.1, min(8.0, (cw-20) / page.rect.width))

    def _add_alpha_highlight(self, r, color_hex="#fff59d", alpha=85, tag=None, outline="#d4af37", width=2):
        if PIL_OK:
            w = max(1, int(r.width)); h = max(1, int(r.height))
            img = Image.new("RGBA", (w, h), (255, 245, 157, alpha))  # gul
            tkimg = ImageTk.PhotoImage(img)
            self._overlay_imgs.append(tkimg)
            self.canvas.create_image(int(r.x0), int(r.y0), image=tkimg, anchor="nw", tags=tag)
            self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1, outline=outline, width=width, tags=tag)
        else:
            self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1, outline=outline, fill=color_hex, stipple="gray25", width=width, tags=tag)

    def _bind_tooltip(self, tag: str, text: str, key: str | None):
        def on_enter(ev, t=text): self.tooltip.show(ev.x, ev.y, t)
        def on_leave(_ev): self.tooltip.hide()
        self.canvas.tag_bind(tag, "<Enter>", on_enter)
        self.canvas.tag_bind(tag, "<Leave>", on_leave)
        if key:
            self.canvas.tag_bind(tag, "<Button-1>", lambda _e, k=key: self.jump_to(k))

    def _flash(self, rect: fitz.Rect):
        mat = fitz.Matrix(self.zoom, self.zoom)
        r = rect * mat
        rid = self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1, outline="#ff6f00", width=4)
        self.canvas.after(700, lambda: self.canvas.delete(rid))

    def render(self, redraw_only: bool=False):
        if not self.doc: return
        if self._fit_width_mode: self._compute_fit_width_zoom()
        page = self.doc[self.page_index]; mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        self.photo = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))
        self.canvas.delete("all"); self.canvas.create_image(0,0,image=self.photo, anchor="nw")
        self.canvas.config(scrollregion=(0,0,pix.width,pix.height))
        self._overlay_imgs.clear()

        # tegn highlights
        if self.page_map:
            for (rect, label, color, kind, key) in self.page_map.get(self.page_index, []):
                if (kind == "value" and not self.show_values.get()) or (kind == "label" and not self.show_labels.get()):
                    continue
                r = rect * mat
                tag = f"{kind}:{key or label}"
                if kind == "value":
                    self._add_alpha_highlight(r, color_hex="#fff59d", alpha=85, tag=tag, outline=color, width=2)
                    tip = f"{label}: {self.key_values.get(key or '', '')}" if key else label
                    self._bind_tooltip(tag, tip, key)
                else:
                    self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1, outline="#9e9e9e", width=1, dash=(3,2), tags=tag)
                    self._bind_tooltip(tag, label, None)

        self._update_nav_state()


# ---------------- Hovedapp ----------------
class InvoiceViewerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Faktura-leser (GUI)")
        self.geometry("1300x880")
        self.file_path = None
        self.dir_path = None
        self.raw_text = ""
        self.strict_items_var = tk.BooleanVar(value=True)
        self._build_ui(); self._setup_logging()
        self.after(100, self.choose_file)

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self, padding=(10,10,10,0)); top.pack(fill="x")
        ttk.Label(top, text="Fil:").pack(side="left")
        self.ent_file = ttk.Entry(top, width=90); self.ent_file.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(top, text="Åpne…", command=self.choose_file).pack(side="left", padx=5)
        ttk.Button(top, text="Kjør", command=self.run_extract).pack(side="left")
        ttk.Button(top, text="Åpne PDF", command=self._open_pdf).pack(side="left", padx=5)
        ttk.Button(top, text="Lagre JSON", command=self.save_json).pack(side="left")

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=10, pady=(5,10)); self.nb = nb
        self.tab_single = ttk.Frame(nb); nb.add(self.tab_single, text="Enkelt")

        # venstre info
        left = ttk.Labelframe(self.tab_single, text="Faktura", padding=10); left.pack(side="left", fill="both", padx=(0,5), pady=5)
        left.grid_columnconfigure(1, weight=1)

        self.vars = {k: tk.StringVar() for k in [
            "invoice_number","kid_number","invoice_date","due_date","seller_name","seller_org","seller_vat","buyer_name","currency"
        ]}

        def _add(lbl, key, row):
            ttk.Label(left, text=lbl+":").grid(row=row, column=0, sticky="e", padx=(0,6), pady=2)
            ent = ttk.Entry(left, textvariable=self.vars[key], state="readonly", width=45)
            ent.grid(row=row, column=1, sticky="we", pady=2)
            btn = ttk.Button(left, text="Vis", width=4, command=lambda k=key: self.pdf_view.jump_to(k))
            btn.grid(row=row, column=2, sticky="w", padx=(6,0))
            ent.bind("<Enter>", lambda _e, k=key: self.pdf_view.preview_key(k, True))
            ent.bind("<Leave>", lambda _e, k=key: self.pdf_view.preview_key(k, False))

        _add("Fakturanr", "invoice_number", 0)
        _add("KID", "kid_number", 1)
        _add("Fakturadato", "invoice_date", 2)
        _add("Forfallsdato", "due_date", 3)
        _add("Selger", "seller_name", 4)
        _add("Selger orgnr", "seller_org", 5)
        _add("Selger MVA", "seller_vat", 6)
        _add("Kjøper", "buyer_name", 7)
        _add("Valuta", "currency", 8)

        right = ttk.Labelframe(self.tab_single, text="Summer", padding=10); right.pack(side="left", fill="y", padx=(0,5), pady=5)
        self.vars_amt = {k: tk.StringVar() for k in ["subtotal","vat","total"]}
        ttk.Label(right, text="Eks. mva:").grid(row=0, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(right, textvariable=self.vars_amt["subtotal"], state="readonly", width=20).grid(row=0, column=1, sticky="w")
        ttk.Label(right, text="MVA:").grid(row=1, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(right, textvariable=self.vars_amt["vat"], state="readonly", width=20).grid(row=1, column=1, sticky="w")
        ttk.Label(right, text="Total:").grid(row=2, column=0, sticky="e", pady=2, padx=5)
        ttk.Entry(right, textvariable=self.vars_amt["total"], state="readonly", width=20).grid(row=2, column=1, sticky="w")
        ttk.Checkbutton(right, text="Streng linjepost-parsing (ignorer støy)", variable=self.strict_items_var,
                        command=self._refresh_items_view_only).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8,0))

        # underfaner
        nb2 = ttk.Notebook(self.tab_single); nb2.pack(side="left", fill="both", expand=True, padx=5, pady=(0,5))
        self.tab_pdf = ttk.Frame(nb2); nb2.add(self.tab_pdf, text="PDF"); self.pdf_view = PdfViewer(self.tab_pdf); self.pdf_view.pack(fill="both", expand=True)

        self.tab_items = ttk.Frame(nb2); nb2.add(self.tab_items, text="Linjeposter")
        cols = ("description","quantity","unit","unit_price","vat_rate","line_total")
        self.tree = ttk.Treeview(self.tab_items, columns=cols, show="headings", height=12)
        for cid, text in zip(cols, ["Beskrivelse","Antall","Enhet","Á pris","MVA %","Linjesum"]): self.tree.heading(cid, text=text)
        self.tree.column("description", width=540); self.tree.column("quantity", width=70, anchor="e")
        self.tree.column("unit", width=70); self.tree.column("unit_price", width=100, anchor="e")
        self.tree.column("vat_rate", width=80, anchor="e"); self.tree.column("line_total", width=110, anchor="e")
        vsb = ttk.Scrollbar(self.tab_items, orient="vertical", command=self.tree.yview); self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True); vsb.pack(side="left", fill="y")

        self.tab_json = ttk.Frame(nb2); nb2.add(self.tab_json, text="JSON"); self.txt_json = tk.Text(self.tab_json, wrap="none"); self.txt_json.pack(fill="both", expand=True)
        self.tab_raw  = ttk.Frame(nb2); nb2.add(self.tab_raw,  text="Råtekst"); self.txt_raw  = tk.Text(self.tab_raw,  wrap="word"); self.txt_raw.pack(fill="both", expand=True)
        self.tab_log  = ttk.Frame(nb2); nb2.add(self.tab_log,  text="Logg");   self.txt_log  = tk.Text(self.tab_log,  height=8, wrap="none"); self.txt_log.pack(fill="both", expand=True)

        # batch
        self.tab_batch = ttk.Frame(nb); nb.add(self.tab_batch, text="Batch")
        toolbar = ttk.Frame(self.tab_batch, padding=(0,8)); toolbar.pack(fill="x")
        ttk.Label(toolbar, text="Mappe:").pack(side="left")
        self.ent_dir = ttk.Entry(toolbar, width=80); self.ent_dir.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(toolbar, text="Velg mappe…", command=self.choose_folder).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Kjør batch", command=self.run_batch).pack(side="left")
        ttk.Button(toolbar, text="Eksporter CSV", command=self.export_csv).pack(side="left", padx=5)
        self.pbar = ttk.Progressbar(self.tab_batch, mode="determinate"); self.pbar.pack(fill="x", padx=5, pady=(0,5))
        cols2 = ("file","invoice_number","invoice_date","due_date","seller","buyer","subtotal","vat","total","currency","kid","error")
        self.tree_batch = ttk.Treeview(self.tab_batch, columns=cols2, show="headings", height=18)
        for cid, text in zip(cols2, ["Fil","Fakturanr","Fakturadato","Forfallsdato","Selger","Kjøper","Eks. mva","MVA","Total","Valuta","KID","Feil"]):
            self.tree_batch.heading(cid, text=text)
        self.tree_batch.column("file", width=280); self.tree_batch.column("seller", width=180); self.tree_batch.column("buyer", width=160)
        self.tree_batch.column("invoice_number", width=140); self.tree_batch.column("invoice_date", width=100); self.tree_batch.column("due_date", width=100)
        self.tree_batch.column("subtotal", width=100, anchor="e"); self.tree_batch.column("vat", width=100, anchor="e")
        self.tree_batch.column("total", width=110, anchor="e"); self.tree_batch.column("currency", width=70, anchor="center")
        self.tree_batch.column("kid", width=160); self.tree_batch.column("error", width=200)
        vsb2 = ttk.Scrollbar(self.tab_batch, orient="vertical", command=self.tree_batch.yview)
        self.tree_batch.configure(yscroll=vsb2.set); self.tree_batch.pack(side="left", fill="both", expand=True, padx=(5,0), pady=(0,5))
        vsb2.pack(side="left", fill="y", pady=(0,5))

    def _setup_logging(self):
        logging.basicConfig(level=logging.INFO)
        h = TextHandler(self.txt_log); h.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logging.getLogger().addHandler(h)

    # ---------- helpers ----------
    def _fmt_money(self, x: Decimal | None) -> str:
        return "" if x is None else f"{x:,.2f}".replace(",", " ").replace(".", ",")

    def _fmt_date_no(self, iso: str | None) -> str:
        if not iso: return ""
        try:
            y, m, d = [int(p) for p in iso.split("-")]
            return f"{d:02d}.{m:02d}.{y:04d}"
        except Exception:
            return iso or ""

    # ---------- actions ----------
    def choose_file(self):
        path = filedialog.askopenfilename(title="Velg faktura", filetypes=[("Dokumenter","*.pdf *.png *.jpg *.jpeg *.tif *.tiff")])
        if not path: return
        self.file_path = path; self.ent_file.delete(0,"end"); self.ent_file.insert(0, path)
        self.pdf_view.load(path); self.nb.select(self.tab_single); self.run_extract()

    def choose_folder(self):
        path = filedialog.askdirectory(title="Velg mappe med fakturaer")
        if not path: return
        self.dir_path = path; self.ent_dir.delete(0,"end"); self.ent_dir.insert(0, path); self.nb.select(self.tab_batch)

    def _open_pdf(self):
        if self.file_path and os.path.exists(self.file_path): _open_with_default_app(self.file_path)
        else: messagebox.showinfo("Ingen fil", "Velg en PDF først.")

    def run_extract(self):
        if not self.file_path:
            messagebox.showinfo("Ingen fil", "Velg en PDF/bilde først."); return
        try:
            logging.info("Leser: %s", self.file_path)
            inv = build_invoice_model(self.file_path)

            # råtekst (for streng linjevisning)
            et = extractors.extract_text_blocks_from_pdf(self.file_path) if self.file_path.lower().endswith(".pdf") \
                 else extractors.extract_text_from_image(self.file_path)
            self.raw_text = et.text or ""

            # felter (datoer i norsk format)
            self.vars["invoice_number"].set(inv.invoice_number or "")
            self.vars["kid_number"].set(inv.kid_number or "")
            self.vars["invoice_date"].set(self._fmt_date_no(inv.invoice_date))
            self.vars["due_date"].set(self._fmt_date_no(inv.payment_terms.due_date))
            self.vars["seller_name"].set(inv.seller.name or "")
            self.vars["seller_org"].set(inv.seller.org_number or "")
            self.vars["seller_vat"].set(inv.seller.vat_number or "")
            self.vars["buyer_name"].set(inv.buyer.name or "")
            self.vars["currency"].set(inv.amounts.currency or "NOK")
            self.vars_amt["subtotal"].set(self._fmt_money(inv.amounts.subtotal_excl_vat))
            self.vars_amt["vat"].set(self._fmt_money(inv.amounts.vat_amount))
            self.vars_amt["total"].set(self._fmt_money(inv.amounts.total_incl_vat))

            # JSON / råtekst
            self.txt_json.delete("1.0","end"); self.txt_json.insert("1.0", model_to_json_text(inv, pretty=True))
            self.txt_raw.delete("1.0","end"); self.txt_raw.insert("1.0", self.raw_text[:200000])

            # linjeposter
            self._populate_items(inv)

            # highlights
            try:
                page_map, key_hits = build_invoice_highlights(self.file_path, inv)
                key_values = {
                    "invoice_number": inv.invoice_number or "",
                    "kid_number":     inv.kid_number or "",
                    "invoice_date":   self._fmt_date_no(inv.invoice_date),
                    "due_date":       self._fmt_date_no(inv.payment_terms.due_date),
                    "subtotal":       self._fmt_money(inv.amounts.subtotal_excl_vat),
                    "vat_amount":     self._fmt_money(inv.amounts.vat_amount),
                    "total":          self._fmt_money(inv.amounts.total_incl_vat),
                    "seller_name":    inv.seller.name or "",
                    "seller_org":     inv.seller.org_number or "",
                    "seller_vat":     inv.seller.vat_number or "",
                    "buyer_name":     inv.buyer.name or "",
                    "currency":       inv.amounts.currency or "",
                }
                self.pdf_view.set_highlights(page_map, key_hits, key_values)
            except Exception as e:
                logging.info("Klarte ikke bygge markeringer: %s", e)
                self.pdf_view.set_highlights({}, {}, {})

            logging.info("Ferdig.")
        except Exception as e:
            logging.exception("Feil ved lesing: %s", e)
            messagebox.showerror("Feil", str(e))

    def _populate_items(self, inv):
        for i in self.tree.get_children(): self.tree.delete(i)
        items = parse_line_items_from_text(self.raw_text) if self.strict_items_var.get() else (inv.line_items or [])
        def s(x): return "" if x is None else str(x).replace(".", ",")
        for li in items:
            self.tree.insert("", "end", values=(li.description or "", s(li.quantity), li.unit or "", s(li.unit_price), s(li.vat_rate), s(li.line_total)))

    def _refresh_items_view_only(self):
        if not self.file_path: return
        try:
            inv = build_invoice_model(self.file_path); self._populate_items(inv)
        except Exception as e:
            logging.exception("Feil ved oppdatering av linjeposter: %s", e)

    def save_json(self):
        if not self.file_path: messagebox.showinfo("Ingen fil", "Kjør først."); return
        out = filedialog.asksaveasfilename(title="Lagre JSON", defaultextension=".json", filetypes=[("JSON","*.json")],
                                           initialfile=os.path.splitext(os.path.basename(self.file_path))[0] + ".json")
        if not out: return
        try:
            inv = build_invoice_model(self.file_path)
            with open(out, "w", encoding="utf-8") as f: f.write(model_to_json_text(inv, pretty=True))
            messagebox.showinfo("Lagret", f"Lagret til:\n{out}")
        except Exception as e:
            logging.exception("Feil ved lagring: %s", e); messagebox.showerror("Feil ved lagring", str(e))

    # ---------- batch ----------
    def run_batch(self):
        if not self.dir_path or not os.path.isdir(self.dir_path):
            messagebox.showinfo("Ingen mappe", "Velg en mappe først."); return
        for i in self.tree_batch.get_children(): self.tree_batch.delete(i)
        files = []
        for root,_,names in os.walk(self.dir_path):
            for n in names:
                if os.path.splitext(n)[1].lower() in {".pdf",".png",".jpg",".jpeg",".tif",".tiff"}:
                    files.append(os.path.join(root,n))
        files.sort(); total = len(files)
        if not total: messagebox.showinfo("Ingen filer", "Fant ingen PDF/bilder i mappen."); return
        self.pbar["maximum"] = total; self.pbar["value"] = 0; self.update_idletasks()
        for idx, path in enumerate(files, 1):
            self.tree_batch.insert("", "end", values=self._process_one_for_batch(path))
            self.pbar["value"] = idx; self.update_idletasks()
        messagebox.showinfo("Ferdig", f"Behandlet {total} filer.")

    def _process_one_for_batch(self, path: str):
        cols2 = ("file","invoice_number","invoice_date","due_date","seller","buyer","subtotal","vat","total","currency","kid","error")
        data = {k:"" for k in cols2}; data["file"] = os.path.relpath(path, self.dir_path or os.path.dirname(path))
        try:
            inv = build_invoice_model(path)
            data.update({
                "invoice_number": inv.invoice_number or "",
                "invoice_date":   self._fmt_date_no(inv.invoice_date),
                "due_date":       self._fmt_date_no(inv.payment_terms.due_date),
                "seller":         inv.seller.name or "",
                "buyer":          inv.buyer.name or "",
                "subtotal":       self._fmt_money(inv.amounts.subtotal_excl_vat),
                "vat":            self._fmt_money(inv.amounts.vat_amount),
                "total":          self._fmt_money(inv.amounts.total_incl_vat),
                "currency":       inv.amounts.currency or "",
                "kid":            inv.kid_number or "",
            })
        except Exception as e:
            data["error"] = str(e)
        return tuple(data[c] for c in cols2)

    def export_csv(self):
        if not self.tree_batch.get_children():
            messagebox.showinfo("Tomt", "Ingen batch-data å eksportere."); return
        out = filedialog.asksaveasfilename(title="Eksporter CSV", defaultextension=".csv",
                                           filetypes=[("CSV","*.csv")], initialfile="fakturakontroll.csv")
        if not out: return
        try:
            with open(out, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Fil","Fakturanr","Fakturadato","Forfallsdato","Selger","Kjøper","Eks. mva","MVA","Total","Valuta","KID","Feil"])
                for iid in self.tree_batch.get_children():
                    w.writerow(self.tree_batch.item(iid, "values"))
            messagebox.showinfo("Lagret", f"Lagret til:\n{out}")
        except Exception as e:
            logging.exception("Feil ved eksport: %s", e); messagebox.showerror("Feil ved eksport", str(e))


def main() -> int:
    app = InvoiceViewerApp(); app.mainloop(); return 0
if __name__ == "__main__":
    raise SystemExit(main())

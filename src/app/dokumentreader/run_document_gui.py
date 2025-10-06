# -*- coding: utf-8 -*-
"""
Multi-dokument GUI med side-by-side visning:
- Venstre: metadata (faktura-nøkkelfelt eller generelle felt)
- Høyre: PDF-preview med highlights (for faktura)
- Bunnfaner: JSON og Råtekst
- Admin (Lær...) for å lage klikkbare maler (anker + verdi)
"""

from __future__ import annotations
import os
import sys
import base64
import platform
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- bootstrap ---
BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

# demp pdfminer-bråk
for name in ("pdfminer", "pdfminer.pdfpage", "pdfminer.pdfinterp", "pdfminer.pdfdocument"):
    logging.getLogger(name).setLevel(logging.ERROR)

# valgfritt: Pillow for pen alfa-highlighting
try:
    from PIL import Image, ImageTk  # noqa: F401
    PIL_OK = True
except Exception:
    PIL_OK = False

import fitz  # PyMuPDF

from app.dokumentreader.document_reader import parse_document
from app.dokumentreader.doc_types import DocumentType, model_to_json_text
from app.dokumentreader.extractors import extract_text_blocks_from_pdf, extract_text_from_image
from app.dokumentreader.highlighter import build_invoice_highlights
from app.dokumentreader.template_engine import apply_templates
from app.dokumentreader.admin_ui import AdminWindow


def _open_with_default_app(path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            import subprocess
            subprocess.run(["open", path])
        else:
            import subprocess
            subprocess.run(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Kunne ikke åpne fil", str(e))


# ---------------- PDF Viewer ----------------
class PdfViewer(ttk.Frame):
    """
    PDF viewer med scroll/zoom, halvtransparent highlighting og støtte for
    Admin-plukk (klikk i PDF for å velge linje).
    """
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
        self._overlay_imgs = []  # holder referanser til tk-bilder
        self._pick_cb = None     # Admin-plukk callback
        self._debug_zone_items: list[int] = []

        # Toolbar
        bar = ttk.Frame(self); bar.pack(fill="x", padx=5, pady=5)
        self.btn_prev = ttk.Button(bar, text="◀", width=4, command=self.prev_page, state="disabled"); self.btn_prev.pack(side="left")
        self.btn_next = ttk.Button(bar, text="▶", width=4, command=self.next_page, state="disabled"); self.btn_next.pack(side="left", padx=(3,8))
        ttk.Label(bar, text="Side:").pack(side="left")
        self.var_page = tk.StringVar(value="0/0")
        ent = ttk.Entry(bar, textvariable=self.var_page, width=10); ent.pack(side="left"); ent.bind("<Return>", self._goto_page_from_entry)
        ttk.Button(bar, text="−", width=3, command=lambda: self.change_zoom(0.85)).pack(side="left", padx=(10,2))
        ttk.Button(bar, text="+", width=3, command=lambda: self.change_zoom(1.15)).pack(side="left", padx=(0,8))
        self.show_values = tk.BooleanVar(value=True)
        self.show_labels = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="Felter",   variable=self.show_values, command=self.render).pack(side="left", padx=(10,0))
        ttk.Checkbutton(bar, text="Etiketter",variable=self.show_labels, command=self.render).pack(side="left")

        # Canvas
        wrap = ttk.Frame(self); wrap.pack(fill="both", expand=True, padx=5, pady=(0,5))
        self.canvas = tk.Canvas(wrap, bg="#f0f0f0")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview); self.vsb.pack(side="left", fill="y")
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview); self.hsb.pack(fill="x")
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # Scroll-hjul (Windows/mac/Linux)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta/120), "units"))
        self.canvas.bind("<Button-4>",  lambda _e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>",  lambda _e: self.canvas.yview_scroll( 3, "units"))
        self.canvas.bind("<Button-1>", self._on_click)

    # ----- API -----
    def load(self, path: str):
        if self.doc:
            try: self.doc.close()
            except Exception: pass
        self.doc = fitz.open(path)
        self.path = path
        self.page_index = 0
        self.zoom = 1.2
        self.page_map, self.key_hits, self.key_values = {}, {}, {}
        self.render()

    def set_highlights(self, page_map: dict | None, key_hits: dict | None, key_values: dict | None):
        self.page_map = page_map or {}
        self.key_hits = key_hits or {}
        self.key_values = key_values or {}
        self.render()

    def enter_pick_mode(self, callback):
        """Admin: klikk i PDF → (page_idx, pdf_x, pdf_y). ESC for å avbryte."""
        self._pick_cb = callback
        messagebox.showinfo("Plukk", "Klikk i PDF-vinduet for å velge. Trykk ESC for å avbryte.")
        self.canvas.focus_set()
        self.canvas.bind("<Escape>", lambda _e: setattr(self, "_pick_cb", None))

    def show_debug_zones(self, zones: list[dict]):
        """Admin: tegn anker/søkesone for test-visning på gjeldende side."""
        self._debug_zone_items.clear()
        self.render()
        mat = fitz.Matrix(self.zoom, self.zoom)
        if not self.doc:
            return
        for z in zones:
            p = int(z.get("page") or 0)
            if p != self.page_index:
                continue
            anc = fitz.Rect(*z["anchor"]) * mat
            zone = fitz.Rect(*z["zone"]) * mat
            i1 = self.canvas.create_rectangle(anc.x0, anc.y0, anc.x1, anc.y1, outline="#3f51b5", width=2)
            i2 = self.canvas.create_rectangle(zone.x0, zone.y0, zone.x1, zone.y1, outline="#ef6c00", dash=(5,3), width=2)
            self._debug_zone_items += [i1, i2]

    # ----- navigasjon / zoom -----
    def next_page(self):
        if self.doc and self.page_index < self.doc.page_count - 1:
            self.page_index += 1
            self.render()

    def prev_page(self):
        if self.doc and self.page_index > 0:
            self.page_index -= 1
            self.render()

    def change_zoom(self, f: float):
        self.zoom = max(0.2, min(6.0, self.zoom * f))
        self.render()

    def _goto_page_from_entry(self, _evt=None):
        if not self.doc:
            return
        try:
            idx = int(self.var_page.get().split("/")[0].strip()) - 1
            if 0 <= idx < self.doc.page_count:
                self.page_index = idx
                self.render()
        except Exception:
            pass

    # ----- klikk / render -----
    def _on_click(self, e):
        if not self.doc:
            return
        if self._pick_cb:
            ox = self.canvas.canvasx(e.x)
            oy = self.canvas.canvasy(e.y)
            pdf_x = ox / self.zoom
            pdf_y = oy / self.zoom
            cb = self._pick_cb
            self._pick_cb = None
            try:
                cb(self.page_index, pdf_x, pdf_y)
            except Exception as ex:
                messagebox.showerror("Plukk-feil", str(ex))
            return

    def render(self):
        if not self.doc:
            return
        mat = fitz.Matrix(self.zoom, self.zoom)
        page = self.doc[self.page_index]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        self.photo = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw")
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self._overlay_imgs.clear()
        self._debug_zone_items.clear()
        self.var_page.set(f"{self.page_index+1}/{self.doc.page_count}")

        # tegn highlights (for faktura)
        if self.page_map:
            for (rect, label, color, kind, _key) in self.page_map.get(self.page_index, []):
                if (kind == "value" and not self.show_values.get()) or (kind == "label" and not self.show_labels.get()):
                    continue
                r = rect * mat
                if PIL_OK and kind == "value":
                    img = Image.new("RGBA", (max(1, int(r.width)), max(1, int(r.height))), (255, 245, 157, 85))
                    tkimg = ImageTk.PhotoImage(img)
                    self._overlay_imgs.append(tkimg)
                    self.canvas.create_image(int(r.x0), int(r.y0), image=tkimg, anchor="nw")
                    self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1, outline=color, width=2)
                else:
                    self.canvas.create_rectangle(r.x0, r.y0, r.x1, r.y1,
                                                 outline=("#9e9e9e" if kind == "label" else color),
                                                 width=(1 if kind == "label" else 2),
                                                 dash=((3, 2) if kind == "label" else None))


# ---------------- Hovedapp ----------------
class MultiDocApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Dokument-leser (faktura, regnskap, skattemelding)")
        self.geometry("1320x900")
        self.file_path: str | None = None
        self._build()

    def _build(self):
        # Topplinje
        top = ttk.Frame(self, padding=(10, 10, 10, 0)); top.pack(fill="x")
        ttk.Label(top, text="Fil:").pack(side="left")
        self.ent = ttk.Entry(top, width=90); self.ent.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(top, text="Åpne…", command=self.open_file).pack(side="left", padx=5)
        ttk.Label(top, text="Profil:").pack(side="left", padx=(10, 2))
        self.profile = tk.StringVar(value="auto")
        ttk.Combobox(top, textvariable=self.profile,
                     values=["auto", "invoice", "financials_no", "vat_return_no"],
                     state="readonly", width=18).pack(side="left")
        ttk.Button(top, text="Kjør", command=self.parse).pack(side="left", padx=5)
        ttk.Button(top, text="Åpne i system", command=self._open).pack(side="left", padx=5)
        ttk.Button(top, text="Admin (Lær…)", command=self._open_admin).pack(side="left", padx=(10,0))

        # Hovedområde: venstre metadata + høyre PDF
        main = ttk.Panedwindow(self, orient="horizontal"); main.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        # Venstre: metadata
        left = ttk.Frame(main); main.add(left, weight=1)
        docfrm = ttk.Frame(left, padding=8); docfrm.pack(fill="x")
        ttk.Label(docfrm, text="Dokumenttype:").grid(row=0, column=0, sticky="e")
        self.var_type = tk.StringVar(value=""); ttk.Entry(docfrm, textvariable=self.var_type, state="readonly", width=30)\
            .grid(row=0, column=1, sticky="w")

        self.frm_invoice = ttk.Labelframe(left, text="Faktura (nøkkelfelt)", padding=10)
        self.frm_invoice.pack(fill="x", padx=8, pady=(0,6))
        self.inv_vars = {k: tk.StringVar() for k in
                         ["invoice_number", "kid", "invoice_date", "due_date", "seller", "seller_org",
                          "seller_vat", "buyer", "currency", "subtotal", "vat", "total"]}

        def add(lbl, key, row):
            ttk.Label(self.frm_invoice, text=lbl + ":").grid(row=row, column=0, sticky="e", padx=(0, 6), pady=2)
            ttk.Entry(self.frm_invoice, textvariable=self.inv_vars[key], state="readonly", width=48)\
                .grid(row=row, column=1, sticky="w")
        add("Fakturanr", "invoice_number", 0); add("KID", "kid", 1)
        add("Fakturadato", "invoice_date", 2); add("Forfallsdato", "due_date", 3)
        add("Selger", "seller", 4); add("Selger orgnr", "seller_org", 5); add("Selger MVA", "seller_vat", 6)
        add("Kjøper", "buyer", 7); add("Valuta", "currency", 8)
        add("Eks. mva", "subtotal", 9); add("MVA", "vat", 10); add("Total", "total", 11)

        self.frm_generic = ttk.Labelframe(left, text="Generelle felter", padding=10)
        self.frm_generic.pack(fill="both", expand=True, padx=8, pady=(0,6))
        self.kv = ttk.Treeview(self.frm_generic, columns=("key", "value"), show="headings", height=18)
        self.kv.heading("key", text="Felt"); self.kv.heading("value", text="Verdi")
        self.kv.column("key", width=340); self.kv.column("value", width=240, anchor="e")
        vsb = ttk.Scrollbar(self.frm_generic, orient="vertical", command=self.kv.yview); self.kv.configure(yscroll=vsb.set)
        self.kv.pack(side="left", fill="both", expand=True); vsb.pack(side="left", fill="y")

        # Høyre: PDF
        right = ttk.Frame(main); main.add(right, weight=3)
        self.viewer = PdfViewer(right); self.viewer.pack(fill="both", expand=True)

        # Nederst: JSON / Råtekst
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=False, padx=10, pady=(0,10))
        self.tab_json = ttk.Frame(nb); nb.add(self.tab_json, text="JSON")
        self.txt_json = tk.Text(self.tab_json, wrap="none", height=14); self.txt_json.pack(fill="both", expand=True)
        self.tab_raw = ttk.Frame(nb); nb.add(self.tab_raw, text="Råtekst")
        self.txt_raw = tk.Text(self.tab_raw, wrap="word", height=14); self.txt_raw.pack(fill="both", expand=True)

    # ---- helpers ----
    def _open(self):
        if self.file_path and os.path.exists(self.file_path):
            _open_with_default_app(self.file_path)
        else:
            messagebox.showinfo("Ingen fil", "Velg en PDF først.")

    def _fmt_date(self, iso: str | None) -> str:
        if not iso:
            return ""
        try:
            y, m, d = [int(p) for p in iso.split("-")]
            return f"{d:02d}.{m:02d}.{y:04d}"
        except Exception:
            return iso or ""

    def _open_admin(self):
        if not self.file_path:
            messagebox.showinfo("Ingen fil", "Åpne en PDF først.")
            return
        try:
            AdminWindow(self, pdf_path=self.file_path,
                        doc_type=DocumentType(self.var_type.get() or "unknown"),
                        viewer=self.viewer)
        except Exception as e:
            messagebox.showerror("Admin", str(e))

    def open_file(self):
        p = filedialog.askopenfilename(
            title="Velg dokument",
            filetypes=[("Dokumenter", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff")]
        )
        if not p:
            return
        self.file_path = p
        self.ent.delete(0, "end"); self.ent.insert(0, p)
        self.viewer.load(p)
        self.parse()

    # ---- kjøring ----
    def parse(self):
        if not self.file_path:
            messagebox.showinfo("Ingen fil", "Velg en fil først.")
            return

        prof = None if self.profile.get() == "auto" else self.profile.get()
        env = parse_document(self.file_path, force_profile=prof)   # orchestrator (klassifiser + parse)

        # dokumenttype
        self.var_type.set(env.doc_type.value)

        # råtekst (debug)
        try:
            et = extract_text_blocks_from_pdf(self.file_path) if self.file_path.lower().endswith(".pdf") \
                 else extract_text_from_image(self.file_path)
            raw = et.text or ""
        except Exception:
            raw = env.raw_text_excerpt or ""
        self.txt_raw.delete("1.0", "end"); self.txt_raw.insert("1.0", raw[:200000])

        # JSON av hele "konvolutten"
        self.txt_json.delete("1.0", "end"); self.txt_json.insert("1.0", model_to_json_text(env, pretty=True))

        # nullstill venstre panel
        for iid in self.kv.get_children():
            self.kv.delete(iid)
        self.frm_invoice.pack_forget()
        self.frm_generic.pack_forget()

        # visning pr. type
        if env.doc_type == DocumentType.INVOICE and env.invoice:
            inv = env.invoice
            # Fakturafelt
            self.frm_invoice.pack(fill="x", padx=8, pady=(0,6))
            self.inv_vars["invoice_number"].set(inv.get("invoice_number") or "")
            self.inv_vars["kid"].set(inv.get("kid_number") or "")
            self.inv_vars["invoice_date"].set(self._fmt_date(inv.get("invoice_date")))
            self.inv_vars["due_date"].set(self._fmt_date((inv.get("payment_terms") or {}).get("due_date")))
            seller = inv.get("seller") or {}; buyer = inv.get("buyer") or {}
            am = inv.get("amounts") or {}
            self.inv_vars["seller"].set(seller.get("name") or "")
            self.inv_vars["seller_org"].set(seller.get("org_number") or "")
            self.inv_vars["seller_vat"].set(seller.get("vat_number") or "")
            self.inv_vars["buyer"].set(buyer.get("name") or "")
            self.inv_vars["currency"].set(am.get("currency") or "NOK")
            self.inv_vars["subtotal"].set("" if am.get("subtotal_excl_vat") is None else am.get("subtotal_excl_vat"))
            self.inv_vars["vat"].set("" if am.get("vat_amount") is None else am.get("vat_amount"))
            self.inv_vars["total"].set("" if am.get("total_incl_vat") is None else am.get("total_incl_vat"))

            # highlights i viewer (samme motor som før)
            try:
                from app.dokumentreader.models import InvoiceModel
                inv_obj = InvoiceModel.parse_obj(inv) if hasattr(InvoiceModel, "parse_obj") else InvoiceModel.model_validate(inv)
                page_map, key_hits = build_invoice_highlights(self.file_path, inv_obj)
                key_vals = {
                    "invoice_number": inv.get("invoice_number") or "",
                    "kid_number": inv.get("kid_number") or "",
                    "invoice_date": self._fmt_date(inv.get("invoice_date")),
                    "due_date": self._fmt_date((inv.get("payment_terms") or {}).get("due_date")),
                    "total": str((inv.get("amounts") or {}).get("total_incl_vat") or ""),
                }
                self.viewer.set_highlights(page_map, key_hits, key_vals)
            except Exception as e:
                logging.info("Ingen markeringer: %s", e)
                self.viewer.set_highlights({}, {}, {})

            # eventuelle mal-felt (valgfritt for faktura)
            self.frm_generic.pack(fill="both", expand=True, padx=8, pady=(0,6))
            try:
                vals = apply_templates(self.file_path, env.doc_type.value)
                for k, v in (vals or {}).items():
                    self.kv.insert("", "end", values=(k, v))
            except Exception as e:
                logging.info("Ingen malverdier: %s", e)

        else:
            # generelle felt (årsregnskap / skattemelding / ukjent)
            self.frm_generic.pack(fill="both", expand=True, padx=8, pady=(0,6))
            if env.doc_type == DocumentType.FINANCIAL_STATEMENT and env.financials:
                for k, v in (env.financials.income_statement or {}).items():
                    self.kv.insert("", "end", values=(f"Resultat: {k}", v))
                for k, v in (env.financials.balance_sheet or {}).items():
                    self.kv.insert("", "end", values=(f"Balanse: {k}", v))
            elif env.doc_type == DocumentType.TAX_RETURN and env.tax_return:
                # baseline parser fyller felt + poster (kan utvides med maler)
                for kv in env.tax_return.fields or []:
                    self.kv.insert("", "end", values=(kv.get("key"), kv.get("value")))
                for k, v in (env.tax_return.posts or {}).items():
                    self.kv.insert("", "end", values=(k, v))
            else:
                self.kv.insert("", "end", values=("Ukjent/ingen profil", ""))

            # maler for andre dokumenter
            try:
                vals = apply_templates(self.file_path, env.doc_type.value)
                if vals:
                    self.kv.insert("", "end", values=("—", "—"))
                    for k, v in vals.items():
                        self.kv.insert("", "end", values=(k, v))
            except Exception as e:
                logging.info("Ingen malverdier: %s", e)

    # ---- utility ----
    def _open_pdf(self):
        if self.file_path and os.path.exists(self.file_path):
            _open_with_default_app(self.file_path)
        else:
            messagebox.showinfo("Ingen fil", "Velg en PDF først.")


def main() -> int:
    app = MultiDocApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

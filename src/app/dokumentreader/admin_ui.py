# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Optional, Dict, Any, Tuple

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

import fitz  # PyMuPDF

from app.dokumentreader.template_engine import (
    Rule, Template, build_rule_from_clicks, save_template, load_templates,
    apply_templates, debug_test_template
)
from app.dokumentreader.doc_types import DocumentType
from app.dokumentreader.invoice_reader import build_invoice_model
from app.dokumentreader.highlighter import build_invoice_highlights


class AdminWindow(tk.Toplevel):
    """
    'Lær'-dialog. Bruk hoved-vieweren i run_document_gui for å plukke punkter i PDF.
    Støtter også 'Auto fra faktura'.
    """
    def __init__(self, master, *, pdf_path: str, doc_type: DocumentType, viewer) -> None:
        super().__init__(master)
        self.title("Admin (Lær) – bygg mal")
        self.geometry("900x560")
        self.resizable(True, True)

        self.pdf_path = pdf_path
        self.doc_type = doc_type
        self.viewer = viewer  # run_document_gui.PdfViewer
        self.doc: fitz.Document | None = None

        self.field_var = tk.StringVar()
        self.profile_var = tk.StringVar(value=doc_type.value)
        self.match_var = tk.StringVar(value="")  # kommaseparerte “match-ord”
        self.tpl_name_var = tk.StringVar(value="min-mal")

        self._picking_kind: Optional[str] = None
        self._anchor_click: Optional[Tuple[int,float,float]] = None
        self._value_click: Optional[Tuple[int,float,float]] = None
        self.rules: List[Rule] = []

        self._build_ui()
        try:
            self.doc = fitz.open(self.pdf_path)
        except Exception as e:
            messagebox.showerror("Kan ikke åpne PDF", str(e))

    # -------- UI --------
    def _build_ui(self):
        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        ttk.Label(top, text="Dokumenttype:").pack(side="left")
        cb = ttk.Combobox(top, textvariable=self.profile_var,
                          state="readonly",
                          values=[d.value for d in DocumentType], width=18)
        cb.pack(side="left", padx=(4,10))

        ttk.Label(top, text="Feltnavn:").pack(side="left")
        ttk.Entry(top, textvariable=self.field_var, width=28).pack(side="left", padx=(4,8))

        ttk.Button(top, text="1) Velg ETIKETT i PDF", command=self._pick_anchor).pack(side="left", padx=(0,6))
        ttk.Button(top, text="2) Velg VERDI i PDF", command=self._pick_value).pack(side="left", padx=(0,8))
        ttk.Button(top, text="Legg til regel", command=self._add_rule).pack(side="left")

        # Auto fra faktura (kun aktivt for invoice)
        self.btn_auto = ttk.Button(top, text="Auto fra faktura", command=self._auto_from_invoice)
        self.btn_auto.pack(side="left", padx=(12,0))
        if self.doc_type != DocumentType.INVOICE:
            self.btn_auto.configure(state="disabled")

        # tabell med regler
        self.tree = ttk.Treeview(self,
                                 columns=("field","label","dir","dx","dy","band","page","regex","hint"),
                                 show="headings", height=14)
        for cid, text, w in [
            ("field","Felt",140), ("label","Etikett",220), ("dir","Retning",70),
            ("dx","max_dx",70), ("dy","max_dy",70), ("band","Band",70),
            ("page","Side",60), ("regex","Regex",160), ("hint","Verdi-eksempel",200)
        ]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=("e" if cid in {"dx","dy","band","page"} else "w"))
        self.tree.pack(fill="both", expand=True, padx=8, pady=(4,6))
        self.tree.bind("<Double-1>", self._edit_cell)

        # bunn
        bottom = ttk.Frame(self, padding=8); bottom.pack(fill="x")
        ttk.Button(bottom, text="Fjern valgt", command=self._remove_selected).pack(side="left")
        ttk.Label(bottom, text="Malenavn:").pack(side="left", padx=(16,0))
        ttk.Entry(bottom, textvariable=self.tpl_name_var, width=18).pack(side="left", padx=(4,12))
        ttk.Label(bottom, text="Match-ord (kommaseparert):").pack(side="left")
        ttk.Entry(bottom, textvariable=self.match_var, width=40).pack(side="left", padx=(4,8))
        ttk.Button(bottom, text="Lagre mal", command=self._save_tpl).pack(side="left", padx=(8,6))
        ttk.Button(bottom, text="Test mal på dokumentet", command=self._test_tpl).pack(side="left")

        self.msg = tk.Text(self, height=5, wrap="word"); self.msg.pack(fill="x", padx=8, pady=(0,8))

    # -------- pick fra viewer --------
    def _pick_anchor(self):
        if not self.field_var.get().strip():
            messagebox.showinfo("Felt", "Skriv feltnavn først.")
            return
        self._picking_kind = "anchor"
        self.viewer.enter_pick_mode(self._on_pick_from_viewer)

    def _pick_value(self):
        if not self.field_var.get().strip():
            messagebox.showinfo("Felt", "Skriv feltnavn først.")
            return
        self._picking_kind = "value"
        self.viewer.enter_pick_mode(self._on_pick_from_viewer)

    def _on_pick_from_viewer(self, page_idx: int, pdf_x: float, pdf_y: float):
        if self._picking_kind == "anchor":
            self._anchor_click = (page_idx, pdf_x, pdf_y)
            self._log(f"Etikett valgt (side {page_idx+1}) – klikk nå på verdien.")
        elif self._picking_kind == "value":
            self._value_click = (page_idx, pdf_x, pdf_y)
            self._log(f"Verdi valgt (side {page_idx+1}). Trykk 'Legg til regel'.")
        self._picking_kind = None

    # -------- regler --------
    def _add_rule(self):
        if not (self._anchor_click and self._value_click and self.doc):
            messagebox.showinfo("Mangler", "Velg både etikett og verdi i PDF først.")
            return
        a_pg, ax, ay = self._anchor_click
        v_pg, vx, vy = self._value_click
        if a_pg != v_pg:
            messagebox.showinfo("Sidesprang", "Etikett og verdi bør være på samme side for denne regelen.")
            return
        rul = build_rule_from_clicks(self.doc, a_pg, (ax, ay), (vx, vy), self.field_var.get().strip())
        self.rules.append(rul)
        self._refresh_table()
        self._anchor_click = None; self._value_click = None
        self._log("Regel lagt til. Legg til flere felt eller lagre mal.")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.rules):
            del self.rules[idx]
            self._refresh_table()

    def _refresh_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in self.rules:
            self.tree.insert("", "end", values=(
                r.field, (r.anchor_text[:80] + ("…" if len(r.anchor_text) > 80 else "")), r.search,
                f"{r.max_dx:.0f}", f"{r.max_dy:.0f}", f"{r.band:.1f}",
                ("" if r.page is None else r.page + 1), (r.value_regex or ""), (r.value_hint or "")
            ))

    def _edit_cell(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx < 0 or idx >= len(self.rules):
            return
        r = self.rules[idx]
        edit = tk.Toplevel(self); edit.title("Endre regel"); edit.geometry("460x260")
        frm = ttk.Frame(edit, padding=8); frm.pack(fill="both", expand=True)
        v_dx = tk.StringVar(value=str(int(r.max_dx))); v_dy = tk.StringVar(value=str(int(r.max_dy)))
        v_band = tk.StringVar(value=str(round(r.band, 1))); v_rgx = tk.StringVar(value=r.value_regex or "")
        ttk.Label(frm, text="max_dx:").grid(row=0, column=0, sticky="e"); ttk.Entry(frm, textvariable=v_dx, width=12).grid(row=0, column=1, sticky="w")
        ttk.Label(frm, text="max_dy:").grid(row=1, column=0, sticky="e"); ttk.Entry(frm, textvariable=v_dy, width=12).grid(row=1, column=1, sticky="w")
        ttk.Label(frm, text="Band:").grid(row=2, column=0, sticky="e"); ttk.Entry(frm, textvariable=v_band, width=12).grid(row=2, column=1, sticky="w")
        ttk.Label(frm, text="value_regex:").grid(row=3, column=0, sticky="e"); ttk.Entry(frm, textvariable=v_rgx, width=34).grid(row=3, column=1, sticky="w")
        def ok():
            try:
                r.max_dx = float(v_dx.get() or 0); r.max_dy = float(v_dy.get() or 0); r.band = float(v_band.get() or 0)
                r.value_regex = v_rgx.get().strip() or None
                self._refresh_table(); edit.destroy()
            except Exception as e:
                messagebox.showerror("Ugyldig", str(e))
        ttk.Button(frm, text="OK", command=ok).grid(row=4, column=1, sticky="e", pady=(12,0))

    def _save_tpl(self):
        name = (self.tpl_name_var.get() or "").strip() or "min-mal"
        prof = (self.profile_var.get() or "").strip() or self.doc_type.value
        match_any = [p.strip() for p in (self.match_var.get() or "").split(",") if p.strip()]
        if not self.rules:
            messagebox.showinfo("Tom", "Ingen regler å lagre.")
            return
        path = save_template(prof, name, match_any, self.rules)
        self._log(f"Malen lagret: {path}")

    def _test_tpl(self):
        prof = (self.profile_var.get() or "").strip() or self.doc_type.value
        # velg mal (prøver navn først)
        tpls = load_templates(prof)
        tpl = None
        for t in tpls:
            if t.name == (self.tpl_name_var.get() or "").strip():
                tpl = t; break
        if tpl is None and tpls:
            tpl = tpls[0]
        if tpl is None:
            self._log("Fant ingen lagrede maler for profilen.")
            return
        dbg = debug_test_template(self.pdf_path, tpl)
        self._log("Test:\n" + _json_pretty(dbg.get("values", {})))
        try:
            zones = dbg.get("zones") or []
            self.viewer.show_debug_zones(zones)
        except Exception:
            pass

    # -------- Auto fra faktura --------
    def _auto_from_invoice(self):
        if self.doc_type != DocumentType.INVOICE:
            messagebox.showinfo("Kun for faktura", "Auto fra faktura er kun tilgjengelig for faktura.")
            return
        try:
            inv = build_invoice_model(self.pdf_path)
            page_map, key_hits = build_invoice_highlights(self.pdf_path, inv)
        except Exception as e:
            messagebox.showerror("Auto", f"Klarte ikke bygge highlight fra faktura:\n{e}")
            return

        # finn første treff per nøkkelfelt
        want = {
            "invoice_number": "Fakturanr",
            "kid_number": "KID",
            "invoice_date": "Fakturadato",
            "due_date": "Forfallsdato",
            "total": "Total",
        }
        # alle etikett-rektangler pr side
        labels_per_page: dict[int, list[Tuple[fitz.Rect,str]]] = {}
        for p, entries in (page_map or {}).items():
            for (rect, label, _color, kind, _key) in entries:
                if kind == "label":
                    labels_per_page.setdefault(p, []).append((rect, label))

        made = 0
        for key, display_name in want.items():
            hits = key_hits.get(key) or []
            if not hits:
                continue
            pidx, vrect = hits[0]
            # velg nærmeste etikett til venstre på "samme linje"
            cand_label = None
            cand_rect: Optional[fitz.Rect] = None
            for (lr, lab) in labels_per_page.get(pidx, []):
                same_line = (abs((lr.y0 + lr.y1)/2 - (vrect.y0 + vrect.y1)/2) <= max(lr.height, vrect.height)*0.7)
                left = lr.x1 <= vrect.x0 + 1
                if same_line and left:
                    cand_label, cand_rect = lab, lr
                    break
            if not cand_label or not cand_rect:
                # fallback: bruk standard-tekst for feltet som ankernavn
                cand_label = display_name

                # anslå "ankersone" som et lite rektangel til venstre for verdien
                cand_rect = fitz.Rect(max(0, vrect.x0 - 80), vrect.y0, vrect.x0, vrect.y1)

            # bygg regel til høyre for etiketten
            band = max(cand_rect.height, vrect.height) * 1.6
            dx = max(60.0, vrect.x1 - cand_rect.x1 + 20.0)
            r = Rule(
                field=display_name,
                anchor_text=cand_label,
                page=pidx,
                search="right",
                max_dx=dx,
                max_dy=0.0,
                band=band,
                anchor_hint_y=(cand_rect.y0 + cand_rect.y1)/2,
                value_regex=None,
                value_hint=None
            )
            self.rules.append(r)
            made += 1

        if made:
            self._refresh_table()
            self._log(f"Laget {made} regler automatisk fra fakturaen. (Juster ved behov, lagre så malen.)")
        else:
            self._log("Fant ingen felt å bygge automatisk. (Mangler highlights?)")

    # -------- util --------
    def _log(self, s: str):
        self.msg.insert("end", s + "\n"); self.msg.see("end")


def _json_pretty(d: Dict[str, Any]) -> str:
    try:
        import json
        return json.dumps(d, ensure_ascii=False, indent=2)
    except Exception:
        return str(d)

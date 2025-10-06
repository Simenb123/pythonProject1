# a07_board.py
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from typing import Any, Dict, List, Tuple, Optional
import re

try:
    from a07_widgets import Table, fmt_amount
except Exception:
    # Minimal fallback hvis modulen ikke er importert enda
    def fmt_amount(x: float) -> str:
        try: return f"{x:,.2f}".replace(",", " ").replace(".", ",")
        except Exception: return str(x)
    class Table(ttk.Treeview):
        def __init__(self, master, columns, **kw):
            ids = [c for c,_ in columns]
            super().__init__(master, columns=ids, show="headings", selectmode="extended", **kw)
            for cid, header in columns:
                self.heading(cid, text=header)
                self.column(cid, width=120, anchor=tk.W)
            sb = ttk.Scrollbar(master, orient="vertical", command=self.yview)
            self.configure(yscrollcommand=sb.set)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            self.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        def set_column_format(self, *a, **k): pass
        def insert_rows(self, rows):
            self.delete(*self.get_children())
            for r in rows:
                vals = [r.get(c, "") for c,_ in self["columns"]]
                self.insert("", tk.END, values=vals)

class A07Board(ttk.Frame):
    """
    Draggable board (konto → A07‑kode).
    Bruk:
        board = A07Board(parent, app=self)   # self = A07App
        board.pack(fill="both", expand=True)
        board.refresh()
    """
    def __init__(self, master, *, app):
        super().__init__(master)
        self.app = app
        self._drag_ghost: Optional[tk.Label] = None
        self._lanes: Dict[str, Tuple[int,int,int,int]] = {}  # code -> bbox på canvas
        self._lane_widgets: Dict[str, ttk.Frame] = {}
        self._lane_lists: Dict[str, tk.Listbox] = {}
        self._last_filter = ""

        # Panes
        self.pw = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        self.left = ttk.Frame(self.pw)
        self.right = ttk.Frame(self.pw)
        self.pw.add(self.left, weight=1)
        self.pw.add(self.right, weight=2)
        self.pw.pack(fill=tk.BOTH, expand=True)

        # -- Left: kontoliste + filter
        top = ttk.Frame(self.left)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(8,4))
        ttk.Label(top, text="Søk (konto/tekst):").pack(side=tk.LEFT)
        self.q = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.q, width=32)
        ent.pack(side=tk.LEFT, padx=6)
        self.q.trace_add("write", lambda *_: self._refresh_accounts())

        self.only_payroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Kun lønnsrelevante konti", variable=self.only_payroll,
                        command=self._refresh_accounts).pack(side=tk.LEFT, padx=8)

        self.tbl = Table(self.left, columns=[
            ("konto","Konto"),
            ("navn","Kontonavn"),
            ("belop","Beløp"),
            ("foreslatt","Foreslått")
        ])
        self.tbl.set_column_format("belop", fmt_amount)
        self.tbl.bind("<ButtonPress-1>", self._on_drag_start)
        self.tbl.bind("<B1-Motion>", self._on_drag_move)
        self.tbl.bind("<ButtonRelease-1>", self._on_drag_end)
        self.tbl.bind("<Double-1>", self._on_double_assign)

        # -- Right: scrollable canvas med lanes
        self.canvas = tk.Canvas(self.right, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self.right, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.bind_all("<Escape>", self._cancel_drag)

    # ---------- Public API ----------
    def refresh(self):
        """Kall når konti/A07/mapping endres."""
        self._refresh_accounts()
        self._build_lanes()

    # ---------- Internals ----------
    def _is_payroll_account(self, accno: str) -> bool:
        digits = re.sub(r"\D+","", str(accno))
        if not digits:
            return False
        v = int(digits)
        # typisk lønnsrelatert: 5xxx og 70xx–73xx + feriepengelån/trekk spesielt
        return (5000 <= v <= 5999) or (7000 <= v <= 7399) or (v in {2940, 5290})

    def _gl_amount(self, acc) -> float:
        # Bruk samme grunnlag som resten av GUI
        val, _lbl = self.app._gl_amount(acc)
        return float(val)

    def _refresh_accounts(self):
        q = (self.q.get() or "").strip().lower()
        rows = []
        sugg = self.app.auto_suggestions or {}
        for acc in self.app.gl_accounts:
            amt = self._gl_amount(acc)
            if self.app.hide_zero.get() and abs(amt) < 1e-9 and abs(float(acc.get("ub",0.0))) < 1e-9:
                continue
            if self.only_payroll.get() and not self._is_payroll_account(acc["konto"]):
                continue
            if q and (q not in str(acc["konto"]).lower()) and (q not in str(acc.get("navn","")).lower()):
                continue
            code = self.app.acc_to_code.get(acc["konto"], "")
            if not code and acc["konto"] in sugg:
                code = sugg[acc["konto"]].get("kode","")
            rows.append({"konto": acc["konto"], "navn": acc.get("navn",""),
                         "belop": amt, "foreslatt": code})
        # standard sortering: konto stigende
        def _key(r):
            s = re.sub(r"\D+","", str(r["konto"]))
            return int(s) if s else 0
        rows.sort(key=_key)
        self.tbl.insert_rows(rows)

    def _build_lanes(self):
        # Clear tidligere lanes
        for ch in list(self.inner.children.values()):
            ch.destroy()
        self._lanes.clear()
        self._lane_lists.clear()
        self._lane_widgets.clear()

        # Koder: hent fra A07 + nåværende mapping
        codes = sorted(set(self._a07_sums().keys()) |
                       set(self.app.acc_to_code.values()))
        # Bygg en lane per kode
        pad = 10
        col_w = max(320, int(self.canvas.winfo_width() * 0.40))
        x = pad; y = pad
        for code in codes:
            frame = ttk.Frame(self.inner, relief=tk.GROOVE, borderwidth=1)
            frame.place(x=x, y=y, width=col_w)
            self._lane_widgets[code] = frame

            # Header med summer og diff
            a07 = self._a07_sums().get(code, 0.0)
            gl  = self._gl_sum_for_code(code)
            diff = a07 - gl
            head = ttk.Frame(frame)
            head.pack(side=tk.TOP, fill=tk.X)
            ttk.Label(head, text=code, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=6, pady=4)
            ttk.Label(head, text=f"A07: {fmt_amount(a07)}  GL: {fmt_amount(gl)}  Diff: {fmt_amount(diff)}",
                      foreground=("#2e7d32" if abs(diff) < float(self.app.diff_threshold.get()) else "#b71c1c")
                      ).pack(side=tk.RIGHT, padx=6)

            # Liste over mappede konti
            lst = tk.Listbox(frame, selectmode="extended", height=8)
            self._lane_lists[code] = lst
            for acc in self.app.gl_accounts:
                if self._account_belongs_to_code(acc, code):
                    txt = f"{acc['konto']}  {acc.get('navn','')}  ({fmt_amount(self._gl_amount(acc))})"
                    lst.insert(tk.END, txt)
            lst.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=6, pady=(0,6))
            lst.bind("<Double-1>", lambda e, c=code: self._remove_selected_in_lane(c))

            # Boksenes geometri for drop‑deteksjon
            self._lanes[code] = (x, y, x+col_w, y + frame.winfo_reqheight())

            # Neste lane posisjon (under forrige)
            y += frame.winfo_reqheight() + pad

        self.inner.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_resize(self, _evt):
        # Rebygg lanes for å bruke ny bredde
        self.after_idle(self._build_lanes)

    # --- A07/GL hjelpeberegninger ---
    def _a07_sums(self) -> Dict[str,float]:
        from a07_core import summarize_by_code
        return summarize_by_code(self.app.rows)

    def _account_belongs_to_code(self, acc: Dict[str,Any], code: str) -> bool:
        # LP aktiv? da kan konto tilhøre flere koder (andel > 0)
        if self.app.use_lp_assignment and self.app.lp_assignment.get(str(acc["konto"])):
            return code in self.app.lp_assignment[str(acc["konto"])]
        # ellers manuell/auto mapping 1‑til‑1
        return self.app.acc_to_code.get(acc["konto"]) == code

    def _gl_sum_for_code(self, code: str) -> float:
        s = 0.0
        if self.app.use_lp_assignment and self.app.lp_assignment:
            for acc in self.app.gl_accounts:
                accno = str(acc["konto"])
                if code in self.app.lp_assignment.get(accno, {}):
                    s += self._gl_amount(acc) * float(self.app.lp_assignment[accno][code])
            # + special_add i GUI (allerede håndtert i kontrollfanen; boardet viser mapping/LP‑sum)
            return s
        for acc in self.app.gl_accounts:
            if self.app.acc_to_code.get(acc["konto"]) == code:
                s += self._gl_amount(acc)
        return s

    # --- DnD ---
    def _selected_account_nos(self) -> List[str]:
        sels = self.tbl.selection()
        out = []
        for it in sels:
            try:
                accno = str(self.tbl.item(it, "values")[0])
                out.append(accno)
            except Exception:
                pass
        return out

    def _on_drag_start(self, _evt):
        self._drag_origin = "left"
        return None

    def _on_drag_move(self, evt):
        if not hasattr(self, "_drag_origin"):
            return
        if not self._selected_account_nos():
            return
        if self._drag_ghost is None:
            self._drag_ghost = tk.Label(self, text=f"{len(self._selected_account_nos())} konto",
                                        background="#455a64", foreground="white")
            self._drag_ghost.place(x=evt.x_root - self.winfo_rootx(),
                                   y=evt.y_root - self.winfo_rooty())
        else:
            self._drag_ghost.place(x=evt.x_root - self.winfo_rootx(),
                                   y=evt.y_root - self.winfo_rooty())

    def _on_drag_end(self, evt):
        if self._drag_ghost is None:
            return
        # Hvilken lane ligger musepekeren over?
        x = evt.x_root - self.right.winfo_rootx()
        y = evt.y_root - self.right.winfo_rooty() + self.canvas.canvasy(0)
        target_code = None
        for code, (x1,y1,x2,y2) in self._lanes.items():
            # juster for venstrepanelets bredde
            lx1, ly1, lx2, ly2 = x1, y1, x2, y2
            if lx1 <= 10 and ly1 <= y <= ly2:  # lane rekker hele bredden; kun Y testes
                target_code = code
                break
        if target_code:
            self._apply_mapping(target_code, self._selected_account_nos())
        self._cancel_drag()

    def _cancel_drag(self, *_a):
        if self._drag_ghost is not None:
            self._drag_ghost.destroy()
        self._drag_ghost = None
        if hasattr(self, "_drag_origin"):
            delattr(self, "_drag_origin")

    def _double_selected_code(self) -> Optional[str]:
        # Returner koden med minst diff (nyttig når bruker dobbeltklikker i kontolista)
        best, best_abs = None, 10**18
        for code in self._a07_sums().keys():
            d = abs(self._a07_sums().get(code,0.0) - self._gl_sum_for_code(code))
            if d < best_abs:
                best, best_abs = code, d
        return best

    def _on_double_assign(self, _evt):
        accs = self._selected_account_nos()
        if not accs:
            return
        code = self._double_selected_code()
        if code:
            self._apply_mapping(code, accs)

    def _remove_selected_in_lane(self, code: str):
        lst = self._lane_lists.get(code)
        if not lst:
            return
        # Hent kontonr fra tekstlinja
        targets = []
        for i in lst.curselection() or []:
            line = lst.get(i)
            accno = line.split()[0]
            targets.append(accno)
        for a in targets:
            if a in self.app.acc_to_code and self.app.acc_to_code[a] == code:
                del self.app.acc_to_code[a]
        self.app.use_lp_assignment = False
        self.app.refresh_control_tables()
        self.refresh()

    def _apply_mapping(self, code: str, accounts: List[str]):
        for accno in accounts:
            self.app.acc_to_code[accno] = code
            # Når board brukes setter vi overstyring manuelt, så deaktiver LP‑visning
            self.app.use_lp_assignment = False
        self.app.refresh_control_tables()
        self.refresh()

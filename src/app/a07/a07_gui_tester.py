# -*- coding: utf-8 -*-
"""
A07 GUI Tester – modulært og DnD-basert kontrollpanel for mapping av A07-koder mot saldobalanse.

Hovedfunksjoner:
- Last A07-agg CSV og Saldobalanse-CSV (fleksible headere/locale støttes i a07_models)
- Last/lagre global regelbok (JSON) – alias, keywords, kontointervall, forventet fortegn
- Dra-og-slipp (TkinterDnD2) for å bygge "Buckets" (enkeltkode/grupperte koder)
- Auto-forslag (beløps-først solver) med toleranse og IB/BEV/UB-valg
- Fargemerking (grønn=diff≈0, gul=delvis, grå=brukt)
- "Sannsynlige konti": liste snevres inn basert på intervaller/teksttreff
- Høyreklikk meny for å fjerne mapping, bytte metric per konto, hoppe til rett kode

NB:
- Denne fila er kortere enn din gamle ~1848-linjers monolitt – fordi
  data-/parser-/modell-logikken er flyttet til a07_models for gjenbruk og testbarhet.
- All nødvendig funksjonalitet er her, men delt i klarere lag.
"""

from __future__ import annotations

import json
import os
import sys
import math
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple, Set

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# DnD
try:
    from tkinterdnd2 import DND_TEXT, TkinterDnD
except Exception as e:
    DND_TEXT = None
    TkinterDnD = None

# Prosjektmodeller (separate, testbare)
# Forventer a07_models i PYTHONPATH / samme prosjekt
from a07_models import (
    GLAccount, A07Entry, A07CodeDef, AmountMetric,
    read_gl_csv, read_a07_csv,
    to_money, add_money, sub_money,
    amount_for_account, parse_account_ranges, ranges_to_spec,
    jaccard, tokenize, account_in_ranges, sign_ok
)

APP_TITLE = "A07 – Matching & DnD Board"
APP_MIN_W = 1280
APP_MIN_H = 780

# Standard lagringssti for regelbok (kan overstyres i GUI)
DEFAULT_RULEBOOK_DIR = r"F:\Dokument\Kildefiler\a07"
DEFAULT_RULEBOOK_NAME = "global_a07_rulebook.json"
DEFAULT_RULEBOOK_PATH = os.path.join(DEFAULT_RULEBOOK_DIR, DEFAULT_RULEBOOK_NAME)

# --------------------------
# Verktøy / Formatering
# --------------------------

def fmt_amount(x: Decimal) -> str:
    # norsk-ish formatering, ingen tusenskilletegn for enkelthetens skyld
    # kan utvides lett til '12 345,67'
    return f"{x:.2f}".replace(".", ",")

def fmt_amount_spaced(x: Decimal) -> str:
    # grov tusenskille med mellomrom
    s = f"{x:.2f}"
    whole, frac = s.split(".")
    neg = whole.startswith("-")
    if neg:
        whole = whole[1:]
    parts = []
    while whole:
        parts.append(whole[-3:])
        whole = whole[:-3]
    whole_fmt = " ".join(reversed(parts))
    return f"{'-' if neg else ''}{whole_fmt},{frac}"

def safe_int(s: str, default: int = 0) -> int:
    try:
        return int(str(s).strip())
    except Exception:
        return default

# --------------------------
# Rulebook JSON store
# --------------------------

class RulebookJSON:
    """
    Lagrer/Laster global A07-regelbok fra JSON:
    {
      "version": 1,
      "codes": [
         {
           "code": "fastloenn",
           "name": "Fast lønn",
           "account_ranges": [[5000,5999],[2900,2949]],
           "expected_sign": 1,
           "aliases": ["lonn", "fastlønn","fast loenn"],
           "keywords": ["fast", "lønn", "loenn"]
         }, ...
      ]
    }
    """
    def __init__(self, json_path: str):
        self.json_path = json_path
        self.codes: Dict[str, A07CodeDef] = {}

    def load(self) -> None:
        p = Path(self.json_path)
        self.codes = {}
        if not p.exists():
            # tom regelbok
            return
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            for c in data.get("codes", []):
                code = c.get("code", "").strip()
                if not code:
                    continue
                name = c.get("name", code)
                ranges = []
                for r in c.get("account_ranges", []):
                    try:
                        a, b = int(r[0]), int(r[1])
                        if a > b:
                            a, b = b, a
                        ranges.append((a, b))
                    except Exception:
                        continue
                expected_sign = c.get("expected_sign", None)
                if expected_sign not in (-1, 0, 1, None):
                    expected_sign = None
                aliases = [str(x) for x in (c.get("aliases") or [])]
                keywords = [str(x) for x in (c.get("keywords") or [])]
                self.codes[code] = A07CodeDef(
                    code=code, name=name,
                    account_ranges=ranges,
                    expected_sign=expected_sign,
                    aliases=aliases, keywords=keywords
                )
        except Exception as e:
            messagebox.showerror("Regelbok", f"Kunne ikke lese JSON:\n{e}")

    def save(self) -> None:
        p = Path(self.json_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "version": 1,
            "codes": []
        }
        for code, d in sorted(self.codes.items(), key=lambda x: x[0]):
            data["codes"].append({
                "code": d.code,
                "name": d.name,
                "account_ranges": [[a, b] for (a, b) in d.account_ranges],
                "expected_sign": d.expected_sign if d.expected_sign in (-1, 0, 1) else None,
                "aliases": list(d.aliases or []),
                "keywords": list(d.keywords or [])
            })
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def upsert_code(self, defn: A07CodeDef) -> None:
        self.codes[defn.code] = defn

    def remove_code(self, code: str) -> None:
        if code in self.codes:
            del self.codes[code]

# --------------------------
# Buckets (enkel/gruppe)
# --------------------------

@dataclass
class Bucket:
    """Samleboks for én eller flere A07-koder + tildelte GL-konti."""
    bucket_id: str
    title: str
    codes: List[str] = field(default_factory=list)      # A07 code ids
    accounts: Set[int] = field(default_factory=set)     # kontonr
    target_amount: Decimal = field(default_factory=lambda: Decimal("0.00"))
    expected_sign: Optional[int] = None                 # aggregeres “svakt”: hvis alle like
    # computed
    sum_gl: Decimal = field(default_factory=lambda: Decimal("0.00"))
    diff: Decimal = field(default_factory=lambda: Decimal("0.00"))

    def compute(self,
                a07: Dict[str, A07Entry],
                glindex: Dict[int, GLAccount],
                metric_of: Dict[int, AmountMetric | None]) -> None:
        # target = sum av kodenes beløp
        tgt = Decimal("0.00")
        signs = set()
        for c in self.codes:
            e = a07.get(c)
            if not e:
                continue
            tgt += e.amount
        self.target_amount = tgt

        # sum(GL) = sum valgt metric for tildelte konti
        sgl = Decimal("0.00")
        for k in self.accounts:
            acc = glindex.get(k)
            if not acc:
                continue
            metric = metric_of.get(k)
            sgl += amount_for_account(acc, metric=metric, override_default=(metric is None))
        self.sum_gl = sgl
        self.diff = (self.target_amount - self.sum_gl).quantize(Decimal("0.01"))


# --------------------------
# Beløps-først Solver
# --------------------------

class AmountFirstSolver:
    """
    Heuristikk:
    1) Kandidater pr kode: kontointervall → tekst/alias → øvrig (scoring)
    2) Single-eksakt (abs(target - acct) <= tol)
    3) Par-kombinasjon (sum av 2 konti ≈ target)
    4) Subset (inntil K=4 konti) på topp-k kandidater
    5) Respektér brukte konti (ikke gjenbruk)
    6) Sign-krav fra regelbok kan vekte noe ned, men blir ikke hard-stop hvis None
    """

    def __init__(self,
                 gl_accounts: List[GLAccount],
                 a07: Dict[str, A07Entry],
                 defs: Dict[str, A07CodeDef],
                 metric_of: Dict[int, AmountMetric | None],
                 tolerance: Decimal = Decimal("5.00"),
                 top_k: int = 12,
                 pair_k: int = 8,
                 subset_k: int = 4):
        self.gl = gl_accounts
        self.a07 = a07
        self.defs = defs
        self.metric_of = metric_of
        self.tol = tolerance
        self.top_k = max(4, top_k)
        self.pair_k = max(4, pair_k)
        self.subset_k = max(2, subset_k)

        self.gl_index: Dict[int, GLAccount] = {int(g.konto): g for g in self.gl}
        self.gl_tokens: Dict[int, List[str]] = {int(g.konto): g.tokens() for g in self.gl}
        self.used_accounts: Set[int] = set()

    def _amount(self, acc: GLAccount) -> Decimal:
        metric = self.metric_of.get(int(acc.konto))
        return amount_for_account(acc, metric=metric, override_default=(metric is None))

    def _candidates_for_code(self, code: str) -> List[Tuple[int, float, Decimal]]:
        """Returner liste [(konto, score, amount)] sortert synkende på score."""
        d = self.defs.get(code)
        tgt = self.a07[code].amount if code in self.a07 else Decimal("0.00")
        if not d:
            d = A07CodeDef(code=code, name=code)

        cand: List[Tuple[int, float, Decimal]] = []
        code_tokens = d.tokens()

        for acc in self.gl:
            k = int(acc.konto)
            if k in self.used_accounts:
                continue
            amt = self._amount(acc)
            score = 0.0
            # Intervall gir stort løft
            in_range = d.contains_account(k)
            if in_range:
                score += 2.5
            # Tekstlig likhet
            sim = jaccard(code_tokens, self.gl_tokens.get(k) or [])
            if sim > 0.0:
                score += (1.25 * sim)
            # Beløpsnærhet
            diff = abs((tgt - amt).copy_abs())
            if diff <= self.tol:
                score += 2.0
            # Eksakt-ish treff
            if diff <= Decimal("0.00"):
                score += 3.0
            # Sign
            if d.expected_sign in (-1, 1):
                good = sign_ok(amt, d.expected_sign)
                score += (0.25 if good else -0.25)

            # Grunnscore hvis helt uten noe: liten terskel for å komme med
            score += 0.01
            cand.append((k, score, amt))

        cand.sort(key=lambda x: (-x[1], abs((tgt - x[2]).copy_abs())))
        return cand

    def _pick_single(self, code: str, tgt: Decimal, cand: List[Tuple[int, float, Decimal]]) -> Optional[int]:
        for (k, score, amt) in cand[: self.top_k]:
            if abs((tgt - amt).copy_abs()) <= self.tol:
                return k
        return None

    def _pick_pair(self, code: str, tgt: Decimal, cand: List[Tuple[int, float, Decimal]]) -> Optional[Tuple[int, int]]:
        arr = cand[: self.pair_k]
        n = len(arr)
        for i in range(n):
            ki, si, ai = arr[i]
            for j in range(i + 1, n):
                kj, sj, aj = arr[j]
                s = ai + aj
                if abs((tgt - s).copy_abs()) <= self.tol:
                    return (ki, kj)
        return None

    def _subset_sum_k(self,
                      tgt: Decimal,
                      arr: List[Tuple[int, float, Decimal]],
                      max_k: int = 4) -> Optional[List[int]]:
        """
        Enkel backtracking for inntil max_k elementer (begrenset for å holde raskt).
        Returnerer liste med kontoer.
        """
        arr2 = arr[: self.top_k]
        values = [(k, amt) for (k, _, amt) in arr2]

        best = None

        def dfs(start: int, chosen: List[int], sum_amt: Decimal, left_k: int):
            nonlocal best
            diff = (tgt - sum_amt).copy_abs()
            if diff <= self.tol:
                best = list(chosen)
                return True
            if left_k == 0 or start >= len(values):
                return False
            # Grei pruning: hvis vi allerede er forbi, stopp
            for i in range(start, len(values)):
                k, a = values[i]
                if k in self.used_accounts:
                    continue
                chosen.append(k)
                if dfs(i + 1, chosen, sum_amt + a, left_k - 1):
                    return True
                chosen.pop()
            return False

        for k in range(3, max_k + 1):
            if dfs(0, [], Decimal("0.00"), k):
                return best
        return None

    def solve_into_buckets(self, buckets: Dict[str, Bucket]) -> Dict[str, Set[int]]:
        """
        Returnerer forslag {bucket_id: {konti}}. Respekterer self.used_accounts.
        """
        # Lag arbeidsrekkefølge – start med størst beløp (vanskeligst)
        order = sorted(
            [b for b in buckets.values()],
            key=lambda b: abs(b.target_amount.copy_abs()),
            reverse=True
        )
        result: Dict[str, Set[int]] = {}

        for b in order:
            tgt = b.target_amount
            # Kandidater = union av kandidater pr kode
            all_cand: Dict[int, Tuple[float, Decimal]] = {}
            for code in b.codes:
                cand = self._candidates_for_code(code)
                for (k, s, a) in cand:
                    prev = all_cand.get(k)
                    if not prev or s > prev[0]:
                        all_cand[k] = (s, a)
            # sortert liste
            cand_list = [(k, s, a) for (k, (s, a)) in all_cand.items()]
            cand_list.sort(key=lambda x: (-x[1], abs((tgt - x[2]).copy_abs())))

            # 1) Single
            single = self._pick_single("|".join(b.codes), tgt, cand_list)
            if single is not None:
                self.used_accounts.add(single)
                result.setdefault(b.bucket_id, set()).add(single)
                continue
            # 2) Par
            pair = self._pick_pair("|".join(b.codes), tgt, cand_list)
            if pair:
                for k in pair:
                    self.used_accounts.add(k)
                    result.setdefault(b.bucket_id, set()).add(k)
                continue
            # 3) Subset inntil 4
            subset = self._subset_sum_k(tgt, cand_list, max_k=self.subset_k)
            if subset:
                for k in subset:
                    self.used_accounts.add(k)
                    result.setdefault(b.bucket_id, set()).add(k)
                continue
            # Ingen forslag – hopp
        return result

# --------------------------
# GUI
# --------------------------

class A07App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.minsize(APP_MIN_W, APP_MIN_H)

        # Dataområder
        self.gl_accounts: List[GLAccount] = []
        self.a07_entries: Dict[str, A07Entry] = {}
        self.rulebook = RulebookJSON(DEFAULT_RULEBOOK_PATH)
        self.rulebook.load()

        # Indekser
        self.gl_index: Dict[int, GLAccount] = {}

        # Mapping / board state
        self.metric_for_account: Dict[int, Optional[AmountMetric]] = {}  # None = bruk default
        self.buckets: Dict[str, Bucket] = {}   # bucket_id -> Bucket
        self.code_to_bucket: Dict[str, str] = {}  # code -> bucket_id
        self.acc_to_bucket: Dict[int, str] = {}   # konto -> bucket_id

        # GUI state
        self.only_likely = tk.BooleanVar(value=True)
        self.metric_mode = tk.StringVar(value="DEFAULT")  # DEFAULT/UB/BEV/IB
        self.tolerance = tk.StringVar(value="5,00")
        self.rulebook_path = tk.StringVar(value=self.rulebook.json_path)

        # Bygg GUI
        self._build_menu_toolbar()
        self._build_body()

        self._refresh_all()

    # --- GUI bygg ---

    def _build_menu_toolbar(self):
        top = ttk.Frame(self.root)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        # Regelbok sti
        ttk.Label(top, text="Regelbok JSON:").pack(side=tk.LEFT)
        self.ent_rule = ttk.Entry(top, width=60, textvariable=self.rulebook_path)
        self.ent_rule.pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Velg ...", command=self.on_choose_rulebook).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Last regelbok", command=self.on_load_rulebook).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Lagre regelbok", command=self.on_save_rulebook).pack(side=tk.LEFT, padx=12)

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Button(top, text="Legg til A07‑kode", command=self.on_add_code).pack(side=tk.LEFT, padx=4)

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Button(top, text="Last A07 CSV", command=self.on_load_a07).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="Last GL CSV", command=self.on_load_gl).pack(side=tk.LEFT, padx=4)

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        # Metric valg
        ttk.Label(top, text="Beløpsgrunnlag:").pack(side=tk.LEFT)
        for label, val in [("Default", "DEFAULT"), ("UB", "UB"), ("Bev", "BEV"), ("IB", "IB")]:
            ttk.Radiobutton(top, text=label, value=val, variable=self.metric_mode,
                            command=self._refresh_all).pack(side=tk.LEFT)

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Label(top, text="Toleranse (kr):").pack(side=tk.LEFT)
        self.ent_tol = ttk.Entry(top, width=6, textvariable=self.tolerance)
        self.ent_tol.pack(side=tk.LEFT)
        ttk.Checkbutton(top, text="Kun sannsynlige konti", variable=self.only_likely,
                        command=self._refresh_lists).pack(side=tk.LEFT, padx=10)

        ttk.Button(top, text="Auto‑forslag mapping", command=self.on_auto_map).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Tøm board", command=self.on_clear_board).pack(side=tk.RIGHT, padx=4)

    def _build_body(self):
        body = ttk.Frame(self.root)
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Venstre: A07‑koder
        left = ttk.Frame(body)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 4), pady=8)

        ttk.Label(left, text="A07‑koder").pack(anchor="w")
        cols = ("code", "name", "amount")
        self.tv_codes = ttk.Treeview(left, columns=cols, show="headings", selectmode="extended", height=18)
        for c, w in [("code", 120), ("name", 280), ("amount", 120)]:
            self.tv_codes.heading(c, text=c.upper())
            self.tv_codes.column(c, width=w, anchor="w" if c != "amount" else "e")
        self.tv_codes.pack(fill=tk.BOTH, expand=True)
        self.tv_codes.bind("<<TreeviewSelect>>", lambda e: self._refresh_accounts())
        self._adopt_dnd_source(self.tv_codes)

        btns = ttk.Frame(left)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="Lag gruppe av valgte", command=self.on_group_selected_codes).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Fjern valgt kode/gruppe fra board", command=self.on_remove_selected_buckets).pack(side=tk.LEFT, padx=8)

        # Midten: Board (Buckets)
        mid = ttk.Frame(body)
        mid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=8)

        ttk.Label(mid, text="Board (Buckets – koder ↔ GL)").pack(anchor="w")
        cols_b = ("bucket", "target", "sumgl", "diff", "n_codes", "n_acc")
        self.tv_board = ttk.Treeview(mid, columns=cols_b, show="headings", height=18, selectmode="browse")
        for c, w, a in [
            ("bucket", 260, "w"),
            ("target", 120, "e"),
            ("sumgl", 120, "e"),
            ("diff", 120, "e"),
            ("n_codes", 80, "center"),
            ("n_acc", 80, "center"),
        ]:
            self.tv_board.heading(c, text=c.upper())
            self.tv_board.column(c, width=w, anchor=a)
        self.tv_board.pack(fill=tk.BOTH, expand=True)

        # farger
        self.tv_board.tag_configure("ok", background="#e9ffe9")      # grønnlig
        self.tv_board.tag_configure("warn", background="#fff7d9")    # gul
        self.tv_board.tag_configure("empty", foreground="#999999")   # grå

        self._adopt_dnd_target(self.tv_board)

        # Høyre: GL-konti
        right = ttk.Frame(body)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 8), pady=8)

        ttk.Label(right, text="GL‑konti (tilgjengelige)").pack(anchor="w")
        cols_g = ("konto", "navn", "ub", "bev", "ib")
        self.tv_gl = ttk.Treeview(right, columns=cols_g, show="headings", height=18, selectmode="extended")
        self.tv_gl.heading("konto", text="KONTO")
        self.tv_gl.column("konto", width=90, anchor="center")
        self.tv_gl.heading("navn", text="NAVN")
        self.tv_gl.column("navn", width=260, anchor="w")
        self.tv_gl.heading("ub", text="UB")
        self.tv_gl.column("ub", width=120, anchor="e")
        self.tv_gl.heading("bev", text="BEV")
        self.tv_gl.column("bev", width=120, anchor="e")
        self.tv_gl.heading("ib", text="IB")
        self.tv_gl.column("ib", width=120, anchor="e")
        self.tv_gl.pack(fill=tk.BOTH, expand=True)
        self._adopt_dnd_source(self.tv_gl)

        # Context menus
        self._build_context_menus()

    def _build_context_menus(self):
        # Board: høyreklikk
        self.menu_board = tk.Menu(self.root, tearoff=0)
        self.menu_board.add_command(label="Fjern valgte GL‑konto(er) fra bucket", command=self.on_ctx_remove_accounts_from_bucket)
        self.menu_board.add_command(label="Fjern hele bucket (koder blir tilgjengelige igjen)", command=self.on_ctx_remove_bucket)

        # GL: høyreklikk
        self.menu_gl = tk.Menu(self.root, tearoff=0)
        self.menu_gl.add_command(label="Sett metric = Default", command=lambda: self._ctx_set_metric(None))
        self.menu_gl.add_command(label="Sett metric = UB", command=lambda: self._ctx_set_metric(AmountMetric.UB))
        self.menu_gl.add_command(label="Sett metric = BEV", command=lambda: self._ctx_set_metric(AmountMetric.BEV))
        self.menu_gl.add_command(label="Sett metric = IB", command=lambda: self._ctx_set_metric(AmountMetric.IB))

        # Bind
        self.tv_board.bind("<Button-3>", self._popup_board)
        self.tv_gl.bind("<Button-3>", self._popup_gl)

    # --- DnD hjelpere ---

    def _adopt_dnd_source(self, widget: ttk.Treeview):
        if TkinterDnD is None:
            return
        try:
            widget.drag_source_register(1, DND_TEXT)
            widget.dnd_bind("<<DragInitCmd>>", self._on_drag_init)
            widget.dnd_bind("<<DragEndCmd>>", self._on_drag_end)
        except Exception:
            pass

    def _adopt_dnd_target(self, widget: ttk.Treeview):
        if TkinterDnD is None:
            return
        try:
            widget.drop_target_register(DND_TEXT)
            widget.dnd_bind("<<DropEnter>>", lambda e: e.action)
            widget.dnd_bind("<<DropPosition>>", lambda e: e.action)
            widget.dnd_bind("<<DropLeave>>", lambda e: e.action)
            widget.dnd_bind("<<Drop>>", self._on_drop_to_board)
        except Exception:
            pass

    def _on_drag_init(self, event):
        # Kommer fra tv_codes eller tv_gl
        w = event.widget
        sel = w.selection()
        if not sel:
            return (None, 0, 0)
        payload_ids = ",".join(sel)
        return ((DND_TEXT, DND_TEXT), tk.DND_TEXT, payload_ids)

    def _on_drag_end(self, event):
        pass

    def _on_drop_to_board(self, event):
        # Motta id’er fra kilder. Finn bucket i board som brukes (eller bruk valgt).
        sel = self.tv_board.selection()
        if not sel:
            messagebox.showwarning("DnD", "Velg en bucket i board før du slipper.")
            return
        bucket_id = sel[0]
        bucket = self.buckets.get(bucket_id)
        if not bucket:
            return
        src = event.widget
        data = event.data or ""
        try:
            ids = [x for x in data.split(",") if x]
        except Exception:
            ids = []
        if not ids:
            return

        if src == self.tv_gl:
            # Legg konti til bucket
            konti = []
            for iid in ids:
                try:
                    konto = int(self.tv_gl.set(iid, "konto"))
                except Exception:
                    continue
                konti.append(konto)
            self._assign_accounts_to_bucket(bucket_id, konti)
        elif src == self.tv_codes:
            # Legg koder til bucket (om det er en ren kode-bucket – gruppér ellers)
            codes = []
            for iid in ids:
                code = self.tv_codes.set(iid, "code")
                if code:
                    codes.append(code)
            self._add_codes_to_bucket(bucket_id, codes)

    def _popup_board(self, event):
        try:
            iid = self.tv_board.identify_row(event.y)
            if iid:
                self.tv_board.selection_set(iid)
            self.menu_board.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu_board.grab_release()

    def _popup_gl(self, event):
        try:
            self.menu_gl.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu_gl.grab_release()

    def _ctx_set_metric(self, m: Optional[AmountMetric]):
        # Set metric for selected accounts in GL list
        sels = self.tv_gl.selection()
        for iid in sels:
            konto = safe_int(self.tv_gl.set(iid, "konto"))
            if konto:
                self.metric_for_account[konto] = m
        self._refresh_board()  # recompute diffs

    # --- I/O handlers ---

    def on_choose_rulebook(self):
        p = filedialog.asksaveasfilename(
            title="Velg/angi regelbok JSON",
            initialdir=os.path.dirname(self.rulebook_path.get()),
            initialfile=os.path.basename(self.rulebook_path.get()),
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("Alle filer", "*.*")]
        )
        if p:
            self.rulebook_path.set(p)

    def on_load_rulebook(self):
        self.rulebook = RulebookJSON(self.rulebook_path.get())
        self.rulebook.load()
        messagebox.showinfo("Regelbok", f"Lest {len(self.rulebook.codes)} koder.")
        self._refresh_all()

    def on_save_rulebook(self):
        self.rulebook.json_path = self.rulebook_path.get()
        self.rulebook.save()
        messagebox.showinfo("Regelbok", f"Lagret {len(self.rulebook.codes)} koder til\n{self.rulebook.json_path}")

    def on_add_code(self):
        dlg = CodeDialog(self.root, self.rulebook.codes)
        self.root.wait_window(dlg.top)
        if dlg.result:
            self.rulebook.upsert_code(dlg.result)
            self._refresh_codes()

    def on_load_a07(self):
        p = filedialog.askopenfilename(
            title="Velg A07 aggregert CSV (kode, navn, beløp)",
            filetypes=[("CSV", "*.csv"), ("Alle filer", "*.*")]
        )
        if not p:
            return
        try:
            entries = read_a07_csv(p)
        except Exception as e:
            messagebox.showerror("A07", f"Feil ved lesing: {e}")
            return
        self.a07_entries = {e.code: e for e in entries}
        # Opprett “default buckets” for hver kode (om ikke eksisterer)
        for code, e in self.a07_entries.items():
            if code in self.code_to_bucket:
                continue
            b_id = f"k:{code}"
            self.buckets[b_id] = Bucket(
                bucket_id=b_id,
                title=f"{code} – {e.name or code}",
                codes=[code]
            )
            self.code_to_bucket[code] = b_id
        self._refresh_all()

    def on_load_gl(self):
        p = filedialog.askopenfilename(
            title="Velg Saldobalanse CSV",
            filetypes=[("CSV", "*.csv"), ("Alle filer", "*.*")]
        )
        if not p:
            return
        try:
            rows = read_gl_csv(p)
        except Exception as e:
            messagebox.showerror("GL", f"Feil ved lesing: {e}")
            return
        self.gl_accounts = rows
        self.gl_index = {int(g.konto): g for g in self.gl_accounts}
        self.metric_for_account.clear()
        self._refresh_all()

    # --- Board/buckets ---

    def on_group_selected_codes(self):
        sels = self.tv_codes.selection()
        codes = []
        for iid in sels:
            code = self.tv_codes.set(iid, "code")
            if code:
                codes.append(code)
        if len(codes) < 2:
            messagebox.showwarning("Gruppe", "Velg minst to A07‑koder for å lage en gruppe.")
            return
        name = simpledialog.askstring("Ny gruppe", "Navn på gruppen (valgfritt):", initialvalue="Gruppe")
        if not name:
            name = f"Gruppe ({len(codes)} koder)"
        # Opprett bucket
        bid = self._new_group_bucket(codes, name)
        # Merk ny bucket valgt
        self.tv_board.selection_set(bid)
        self._refresh_board()

    def _new_group_bucket(self, codes: List[str], name: str) -> str:
        # Fjern koder fra tidligere buckets, legg i ny
        for c in codes:
            old = self.code_to_bucket.get(c)
            if old and old in self.buckets:
                self.buckets[old].codes = [x for x in self.buckets[old].codes if x != c]
                if not self.buckets[old].codes and not self.buckets[old].accounts:
                    del self.buckets[old]
            self.code_to_bucket.pop(c, None)

        bid = f"g:{abs(hash(tuple(sorted(codes))))}"
        self.buckets[bid] = Bucket(bucket_id=bid, title=name, codes=codes)
        for c in codes:
            self.code_to_bucket[c] = bid
        return bid

    def _assign_accounts_to_bucket(self, bucket_id: str, konti: Sequence[int]) -> None:
        b = self.buckets.get(bucket_id)
        if not b:
            return
        for k in konti:
            # hvis konto allerede brukt i annen bucket: fjern derfra
            old = self.acc_to_bucket.get(k)
            if old and old in self.buckets and old != bucket_id:
                self.buckets[old].accounts.discard(k)
            b.accounts.add(k)
            self.acc_to_bucket[k] = bucket_id
        self._refresh_board()
        self._refresh_lists()

    def _add_codes_to_bucket(self, bucket_id: str, codes: Sequence[str]) -> None:
        b = self.buckets.get(bucket_id)
        if not b:
            return
        for c in codes:
            old = self.code_to_bucket.get(c)
            if old and old in self.buckets and old != bucket_id:
                self.buckets[old].codes = [x for x in self.buckets[old].codes if x != c]
            if c not in b.codes:
                b.codes.append(c)
            self.code_to_bucket[c] = bucket_id
        self._refresh_board()
        self._refresh_lists()

    def on_remove_selected_buckets(self):
        # Fjern hele bucket eller code-bucket for valgte codes i venstre liste
        sels_b = self.tv_board.selection()
        if sels_b:
            for bid in sels_b:
                self._remove_bucket(bid)
        else:
            sels_c = self.tv_codes.selection()
            for iid in sels_c:
                code = self.tv_codes.set(iid, "code")
                bid = self.code_to_bucket.get(code)
                if bid:
                    self._remove_bucket(bid)
        self._refresh_all()

    def _remove_bucket(self, bucket_id: str):
        b = self.buckets.get(bucket_id)
        if not b:
            return
        # Frigi koder
        for c in b.codes:
            self.code_to_bucket.pop(c, None)
            # legg tilbake som egen bucket
            if c in self.a07_entries:
                nbid = f"k:{c}"
                self.buckets[nbid] = Bucket(bucket_id=nbid, title=f"{c} – {self.a07_entries[c].name or c}", codes=[c])
                self.code_to_bucket[c] = nbid
        # Frigi konti
        for k in list(b.accounts):
            self.acc_to_bucket.pop(k, None)
        # Slett
        del self.buckets[bucket_id]

    def on_ctx_remove_accounts_from_bucket(self):
        sels_b = self.tv_board.selection()
        if not sels_b:
            return
        bid = sels_b[0]
        b = self.buckets.get(bid)
        if not b:
            return
        # Finn markerte konti i GL-lista og fjern fra bucket
        remove_k: Set[int] = set()
        for iid in self.tv_gl.selection():
            k = safe_int(self.tv_gl.set(iid, "konto"))
            if k:
                remove_k.add(k)
        if not remove_k:
            # alternativ: fjern alle konti i bucket
            if messagebox.askyesno("Fjern", "Ingen konti valgt i GL-lista.\nFjerne alle konti fra denne bucket?"):
                remove_k = set(b.accounts)
        for k in remove_k:
            b.accounts.discard(k)
            self.acc_to_bucket.pop(k, None)
        self._refresh_board()
        self._refresh_lists()

    def on_ctx_remove_bucket(self):
        sels_b = self.tv_board.selection()
        for bid in sels_b:
            self._remove_bucket(bid)
        self._refresh_all()

    # --- Auto-map ---

    def on_auto_map(self):
        # Oppdater buckets' target før solving
        self._compute_buckets()
        # init metric-of med global preferanse
        self._apply_global_metric()
        tol = self._read_tolerance()
        solver = AmountFirstSolver(
            gl_accounts=self.gl_accounts,
            a07=self.a07_entries,
            defs=self.rulebook.codes,
            metric_of=self.metric_for_account,
            tolerance=tol,
            top_k=14,
            pair_k=10,
            subset_k=4
        )
        # marker brukte konti (allerede i buckets)
        solver.used_accounts = set(self.acc_to_bucket.keys())

        suggestions = solver.solve_into_buckets(self.buckets)
        # anvend forslag
        for bid, konti in suggestions.items():
            self._assign_accounts_to_bucket(bid, konti)
        # re-calc
        self._refresh_board()
        self._refresh_lists()
        messagebox.showinfo("Auto‑forslag", "Auto‑mapping fullført.")

    # --- Refresh UI ---

    def _apply_global_metric(self):
        mode = self.metric_mode.get().upper()
        if mode == "DEFAULT":
            # Fjern per-konto overstyring
            self.metric_for_account = {k: None for k in self.metric_for_account.keys()}
        else:
            m = AmountMetric.UB if mode == "UB" else AmountMetric.BEV if mode == "BEV" else AmountMetric.IB
            # Sett for alle *brukte* konti – tilgjengelige konti får normal default ved bruk
            for k in list(self.acc_to_bucket.keys()):
                self.metric_for_account[k] = m

    def _read_tolerance(self) -> Decimal:
        s = self.tolerance.get().strip().replace(" ", "").replace(".", "").replace(",", ".")
        try:
            x = Decimal(s)
        except Exception:
            x = Decimal("5.00")
        return x.quantize(Decimal("0.01"))

    def _compute_buckets(self):
        # Oppdater target/sum/diff for alle buckets
        for b in self.buckets.values():
            b.compute(self.a07_entries, self.gl_index, self.metric_for_account)

    def _refresh_all(self):
        self._compute_buckets()
        self._refresh_codes()
        self._refresh_board()
        self._refresh_accounts()

    def _refresh_codes(self):
        self.tv_codes.delete(*self.tv_codes.get_children())
        # vis koder (ikke grupper) + grupper som egne rader (med leading ★)
        # Vi holder tv_codes til rene A07-koder (ikke grupper) for enklere valg
        for code, e in sorted(self.a07_entries.items(), key=lambda x: x[0]):
            iid = code
            self.tv_codes.insert("", "end", iid=iid, values=(e.code, e.name, fmt_amount_spaced(e.amount)))
        # marker koder som er i en gruppe? (kun ved behov)
        # (kan evt. addere tag for "in_group")

    def _refresh_board(self):
        self._compute_buckets()
        self.tv_board.delete(*self.tv_board.get_children())
        for bid, b in sorted(self.buckets.items(), key=lambda x: x[0]):
            tags = []
            if b.accounts:
                if abs(b.diff.copy_abs()) <= self._read_tolerance():
                    tags.append("ok")
                else:
                    tags.append("warn")
            else:
                tags.append("empty")
            self.tv_board.insert(
                "", "end", iid=bid,
                values=(
                    b.title,
                    fmt_amount_spaced(b.target_amount),
                    fmt_amount_spaced(b.sum_gl),
                    fmt_amount_spaced(b.diff),
                    len(b.codes),
                    len(b.accounts)
                ),
                tags=tags
            )

    def _refresh_accounts(self):
        # Filtrer GL etter sannsynlige for *valgt* bucket (eller union av valgte koder)
        self.tv_gl.delete(*self.tv_gl.get_children())
        if not self.gl_accounts:
            return

        likely_set: Optional[Set[int]] = None
        if self.only_likely.get():
            # lag "sannsynlige" basert på markerte koder (A07-lista) eller valgt bucket
            codes = self._selected_codes()
            if not codes:
                # hvis board har valgt bucket – bruk dens koder
                bsel = self.tv_board.selection()
                if bsel:
                    bid = bsel[0]
                    b = self.buckets.get(bid)
                    if b:
                        codes = list(b.codes)
            likely_set = self._likely_accounts_for_codes(codes)

        for g in self.gl_accounts:
            k = int(g.konto)
            # skjul konti som allerede er brukt i en bucket
            if k in self.acc_to_bucket:
                continue
            if likely_set is not None and k not in likely_set:
                continue
            ub = amount_for_account(g, metric=AmountMetric.UB)
            bev = amount_for_account(g, metric=AmountMetric.BEV)
            ib = amount_for_account(g, metric=AmountMetric.IB)
            self.tv_gl.insert("", "end", iid=f"acc:{k}",
                              values=(k, g.navn, fmt_amount_spaced(ub), fmt_amount_spaced(bev), fmt_amount_spaced(ib)))

    def _refresh_lists(self):
        self._refresh_accounts()
        self._refresh_board()

    def _selected_codes(self) -> List[str]:
        codes = []
        for iid in self.tv_codes.selection():
            code = self.tv_codes.set(iid, "code")
            if code:
                codes.append(code)
        return codes

    def _likely_accounts_for_codes(self, codes: Sequence[str]) -> Set[int]:
        if not codes:
            return set(int(g.konto) for g in self.gl_accounts)
        # union av intervaller + litt tekstmatcher
        res: Set[int] = set()
        for code in codes:
            d = self.rulebook.codes.get(code)
            c_toks = d.tokens() if d else []
            for g in self.gl_accounts:
                k = int(g.konto)
                if d and d.contains_account(k):
                    res.add(k)
                    continue
                # tekst
                if c_toks:
                    sim = jaccard(c_toks, g.tokens())
                    if sim >= 0.15:
                        res.add(k)
        return res

    # --- kommandoer topp ---

    def on_clear_board(self):
        self.code_to_bucket.clear()
        self.acc_to_bucket.clear()
        self.metric_for_account.clear()
        self.buckets.clear()
        # recreate code buckets
        for code, e in self.a07_entries.items():
            b_id = f"k:{code}"
            self.buckets[b_id] = Bucket(
                bucket_id=b_id, title=f"{code} – {e.name or code}", codes=[code]
            )
            self.code_to_bucket[code] = b_id
        self._refresh_all()

# --------------------------
# Dialog for ny/endre kode
# --------------------------

class CodeDialog:
    def __init__(self, parent, existing: Dict[str, A07CodeDef], code: Optional[str] = None):
        self.top = tk.Toplevel(parent)
        self.top.title("A07‑kode – ny/endre")
        self.top.transient(parent)
        self.top.grab_set()
        self.result: Optional[A07CodeDef] = None

        self.existing = existing
        self.var_code = tk.StringVar(value=code or "")
        self.var_name = tk.StringVar(value="")
        self.var_ranges = tk.StringVar(value="")
        self.var_sign = tk.StringVar(value="none")
        self.var_alias = tk.StringVar(value="")
        self.var_keywords = tk.StringVar(value="")

        # Layout
        frm = ttk.Frame(self.top, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        r = 0
        ttk.Label(frm, text="Kode:").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_code, width=30).grid(row=r, column=1, sticky="w")
        r += 1
        ttk.Label(frm, text="Navn:").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_name, width=50).grid(row=r, column=1, sticky="we")
        r += 1
        ttk.Label(frm, text="Kontointervall (eks. 5000-5999|2900-2949|7000):").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_ranges, width=50).grid(row=r, column=1, sticky="we")
        r += 1
        ttk.Label(frm, text="Forventet fortegn:").grid(row=r, column=0, sticky="w")
        frm_sign = ttk.Frame(frm)
        frm_sign.grid(row=r, column=1, sticky="w")
        for txt, val in [("Ingen", "none"), ("Positiv (+)", "pos"), ("Negativ (−)", "neg")]:
            ttk.Radiobutton(frm_sign, text=txt, value=val, variable=self.var_sign).pack(side=tk.LEFT, padx=3)
        r += 1
        ttk.Label(frm, text="Alias (|‑separert):").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_alias, width=50).grid(row=r, column=1, sticky="we")
        r += 1
        ttk.Label(frm, text="Nøkkelord (|‑separert):").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_keywords, width=50).grid(row=r, column=1, sticky="we")
        r += 1

        btns = ttk.Frame(frm)
        btns.grid(row=r, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btns, text="Lagre", command=self.on_save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Avbryt", command=self.top.destroy).pack(side=tk.LEFT, padx=6)

        frm.columnconfigure(1, weight=1)

    def on_save(self):
        code = self.var_code.get().strip()
        if not code:
            messagebox.showwarning("Kode", "Kode mangler.")
            return
        name = self.var_name.get().strip() or code
        ranges = parse_account_ranges(self.var_ranges.get())
        sign = self.var_sign.get()
        expected_sign = None
        if sign == "pos":
            expected_sign = 1
        elif sign == "neg":
            expected_sign = -1
        aliases = [x.strip() for x in self.var_alias.get().split("|") if x.strip()]
        keywords = [x.strip() for x in self.var_keywords.get().split("|") if x.strip()]
        self.result = A07CodeDef(
            code=code, name=name, account_ranges=ranges,
            expected_sign=expected_sign, aliases=aliases, keywords=keywords
        )
        self.top.destroy()

# --------------------------
# main
# --------------------------

def main():
    # Start Tk (DnD root hvis tilgjengelig)
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = A07App(root)
    root.mainloop()

if __name__ == "__main__":
    main()

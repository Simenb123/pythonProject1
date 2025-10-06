# mapping_utils.py – 2025-06-03 (r4 – grab-release-fix)
from __future__ import annotations

import difflib
import logging
import re
import unicodedata
from typing import Dict, List

import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

logger = logging.getLogger(__name__)

# ────────── standardfelter ──────────
REQ_MAND = [
    ("konto", "Kontonummer"),
    ("beløp", "Beløp"),
    ("dato", "Dato"),
    ("bilagsnr", "Bilagsnummer"),
]
REQ_OPT = [
    ("kontonavn", "Kontonavn"),
    ("mva-kode", "MVA-kode"),
    ("mvabeløp", "MVA-beløp"),
    ("tekst", "Tekst"),
    ("beskrivelse", "Beskrivelse"),
]
REQ = REQ_MAND + REQ_OPT


# ────────── helpers ──────────
def norm(s: str) -> str:
    return re.sub(
        r"[^0-9a-z]",
        "",
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode()
        .lower(),
    )


SYNONYMS = {
    "konto": ("konto", "kontonr", "kontonummer", "account", "acct"),
    "beløp": ("beløp", "belop", "amount", "sum", "total"),
    "dato": ("dato", "date", "bilagsdato", "transdate"),
    "bilagsnr": ("bilagsnr", "bilagsnummer", "bilagnr", "voucher", "docno", "bilag"),
    "kontonavn": ("kontonavn", "accountname", "account name"),
    "mva-kode": ("mvakode", "mva-kode", "mva_kode", "vatcode", "vat_code", "mva"),
    "mvabeløp": ("mvabeløp", "mva-belop", "mva_beløp", "vatamount", "vat_amount"),
    "tekst": ("tekst", "text", "description", "desc"),
    "beskrivelse": ("beskrivelse", "description", "tekst"),
}


def infer_mapping(cols: List[str]) -> dict[str, str] | None:
    out: Dict[str, str] = {}
    low = [norm(c) for c in cols]

    for std, _ in REQ:
        cand = None
        for syn in (std, *SYNONYMS.get(std, ())):
            cand = next((c for c in cols if norm(syn) in norm(c)), None)
            if cand:
                break
        if cand is None:
            m = difflib.get_close_matches(std, low, n=1, cutoff=0.85)
            if m:
                cand = cols[low.index(m[0])]
        if cand:
            out[std] = cand

    if all(k in out for k, _ in REQ_MAND):
        return out
    return None


# ────────── FeltVelger ──────────
class FeltVelger(tk.Toplevel):
    """Modal dialog som lar bruker mappe kolonner → standardfelt."""

    def __init__(self, master: tk.Misc, df: pd.DataFrame, defaults: dict | None):
        super().__init__(master)
        self.grab_set()
        self.title("Kolonne-mapping")
        self.resizable(False, False)

        self._cols = [""] + list(df.columns)
        self.req: dict[str, ttk.Combobox] = {}
        self.extra = []

        # obligatoriske + valgfrie felt
        for r, (key, vis) in enumerate(REQ):
            ttk.Label(self, text=vis).grid(row=r, column=0, sticky="w")
            cb = ttk.Combobox(self, values=self._cols, state="readonly", width=34)
            if defaults and key in defaults:
                cb.set(defaults[key])
            else:
                for syn in (key, *SYNONYMS.get(key, ())):
                    for c in df.columns:
                        if norm(syn) in norm(c):
                            cb.set(c)
                            break
                    if cb.get():
                        break
            cb.grid(row=r, column=1, padx=3, pady=1)
            cb.bind("<Delete>", lambda e, c=cb: c.set(""))
            self.req[key] = cb

        # ekstra felt
        self.ex = ttk.Frame(self)
        self.ex.grid(row=len(REQ), column=0, columnspan=2)
        ttk.Button(self.ex, text="+ Ekstra felt", command=self._add_extra)\
            .grid(row=0, column=0, columnspan=2)

        if defaults:
            for k, v in defaults.items():
                if k not in self.req:
                    self._add_extra(prefill=(k, v))

        ttk.Button(self, text="OK", command=self._ok)\
            .grid(row=len(REQ) + 1, column=0, columnspan=2, pady=6)
        self.bind("<Escape>", lambda *_: self._cancel())

    # ----- helpers ------------------------------------------------------
    def _add_extra(self, *, prefill: tuple[str, str] | None = None):
        r = len(self.extra) + 1
        k_var = tk.StringVar(value=prefill[0] if prefill else "")
        v_var = tk.StringVar(value=prefill[1] if prefill else "")
        ttk.Entry(self.ex, textvariable=k_var, width=15)\
            .grid(row=r, column=0, padx=2, pady=1)
        cb = ttk.Combobox(
            self.ex, textvariable=v_var, values=self._cols,
            state="readonly", width=22
        )
        cb.grid(row=r, column=1, padx=2, pady=1)
        cb.bind("<Delete>", lambda e, c=cb: c.set(""))
        self.extra.append((k_var, cb))

    def _ok(self):
        missing = [vis for k, vis in REQ_MAND if not self.req[k].get()]
        if missing:
            messagebox.showerror("Mangler", ", ".join(missing), parent=self)
            return
        self._result = self._collect_mapping()
        self._close()

    def _cancel(self):
        self._result = None
        self._close()

    def _collect_mapping(self) -> dict[str, str]:
        mp = {k: cb.get() for k, cb in self.req.items() if cb.get()}
        for k_var, cb in self.extra:
            k = k_var.get().strip().lower()
            v = cb.get()
            if k and v:
                mp[k] = v
        return mp

    # ----- shutdown / public -------------------------------------------
    def _close(self):
        self.grab_release()
        self.destroy()

    def mapping(self) -> dict[str, str] | None:
        self._result = None
        self.mainloop()
        return self._result

"""
a07_board_dnd.py
=================

Dette modul definerer en ``A07Board``-widget med drag‑and‑drop støtte.
Widgeten viser GL-konti i en tabell til venstre og A07‑koder som kort
til høyre.  Brukeren kan dra en konto og slippe den på et kodekort for å
tilordne kontoen til koden.  En ``on_map`` callback meldes når
tilordningen skjer.  Dette gjør at brettet fortsatt er isolert fra
lagringslogikk – det bare informerer applikasjonen om at en mapping er
ønsket.

Ut over drag‑and‑drop er dette brettet en kopi av den enklere
``a07_board.py``; koden for klikk‑mapping er fjernet.  Kortene viser
A07-sum, GL-sum og diff per kode, sortert etter størst avvik først.

Bruk:

    from a07_board_dnd import A07Board
    board = A07Board(parent, on_map=handle_mapping)
    board.update(accounts, a07_sums, mapping, basis="endring")

"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional

from models import GLAccount, summarize_gl_by_code


class A07Board(ttk.Frame):
    """Interaktivt DnD-brett for å mappe GL-konti til A07-koder."""

    def __init__(self, master: tk.Widget, *, on_map: Callable[[GLAccount, str], None]):
        super().__init__(master)
        self.on_map = on_map
        self.accounts: List[GLAccount] = []
        self.a07_sums: Dict[str, float] = {}
        self.mapping: Dict[str, str] = {}
        self.basis: str = "endring"
        # Drag‑state
        self._drag_acc: Optional[str] = None  # kontonummer som dras
        self._drag_label: Optional[tk.Label] = None  # label som følger musa
        # Venstre panel: kontoer
        left_frame = ttk.Frame(self)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 4), pady=8)
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        ttk.Label(search_frame, text="Søk konto:").pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self._search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=(4, 0))
        search_entry.bind("<KeyRelease>", lambda e: self.refresh_accounts())
        # Treeview med kontoer
        self.tree = ttk.Treeview(left_frame, columns=("konto", "navn", "belop"), show="headings")
        self.tree.heading("konto", text="Konto")
        self.tree.heading("navn", text="Kontonavn")
        self.tree.heading("belop", text="Beløp")
        self.tree.column("konto", width=80, anchor=tk.W)
        self.tree.column("navn", width=200, anchor=tk.W)
        self.tree.column("belop", width=100, anchor=tk.E)
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Bind drag‑events på kontolisten
        self.tree.bind("<ButtonPress-1>", self._on_tree_press)
        self.tree.bind("<B1-Motion>", self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_release)
        # Høyre panel: A07-kodekort
        right_frame = ttk.Frame(self)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 8), pady=8)
        self.canvas = tk.Canvas(right_frame, highlightthickness=0)
        self.cards_container = ttk.Frame(self.canvas)
        self.scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas_window = self.canvas.create_window((0, 0), window=self.cards_container, anchor="nw")
        # Scrollregion og breddejustering
        self.cards_container.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.cards_container.bind(
            "<Configure>", lambda e: self.canvas.itemconfigure(self._canvas_window, width=self.canvas.winfo_width())
        )

    # -----------------------------------------------------------------
    # Oppdatering av data
    # -----------------------------------------------------------------
    def update(self, accounts: List[GLAccount], a07_sums: Dict[str, float], mapping: Dict[str, str], basis: str = "endring") -> None:
        """Oppdater brettet med nye data og bygg kontoer og koder.

        Args:
            accounts: Liste av ``GLAccount``-objekter.
            a07_sums: Summer pr A07-kode.
            mapping: Gjeldende mapping (konto -> kode).
            basis: Hvilket feltnavn på ``GLAccount`` som skal brukes for beløp.
        """
        self.accounts = list(accounts)
        self.a07_sums = dict(a07_sums)
        self.mapping = dict(mapping)
        self.basis = basis
        self.refresh_accounts()
        self.refresh_codes()

    # -----------------------------------------------------------------
    # Oppdater kontoer
    # -----------------------------------------------------------------
    def refresh_accounts(self) -> None:
        """Filter og vis kontoer i treet basert på søkestrengen."""
        query = (self._search_var.get() or "").strip().lower()
        # Tøm tidligere rader
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        # Legg til filtrerte kontoer
        for acc in self.accounts:
            if query and (query not in acc.konto.lower() and query not in acc.navn.lower()):
                continue
            # Velg basisfeltet
            if self.basis == "ub":
                amount = acc.ub
            elif self.basis == "belop":
                amount = acc.belop
            else:
                amount = acc.endring
            self.tree.insert(
                "", tk.END,
                values=(acc.konto, acc.navn, f"{amount:,.2f}".replace(",", " ").replace(".", ",")),
            )

    # -----------------------------------------------------------------
    # Oppdater kodekort
    # -----------------------------------------------------------------
    def refresh_codes(self) -> None:
        """Bygg kodekort basert på summer fra A07 og GL."""
        # Fjern eksisterende kort
        for child in self.cards_container.winfo_children():
            child.destroy()
        # Summér GL-pr-kode basert på mapping
        gl_sums = summarize_gl_by_code(self.accounts, self.mapping, basis=self.basis)
        # Lag en oversikt over hvilke kontoer som er mappet til hver kode, med valgt basis-beløp
        mapped_accounts: Dict[str, List[tuple[str, float]]] = {}
        for acc in self.accounts:
            code = self.mapping.get(acc.konto)
            if not code:
                continue
            if self.basis == "ub":
                amount = acc.ub
            elif self.basis == "belop":
                amount = acc.belop
            else:
                amount = acc.endring
            mapped_accounts.setdefault(code, []).append((acc.konto, float(amount)))
        codes = set(self.a07_sums) | set(gl_sums)
        def sort_key(code: str) -> float:
            a07 = float(self.a07_sums.get(code, 0.0))
            gl = float(gl_sums.get(code, 0.0))
            return -abs(a07 - gl)
        sorted_codes = sorted(codes, key=sort_key)
        row = 0
        col = 0
        for code in sorted_codes:
            card = ttk.Frame(self.cards_container, relief=tk.RIDGE, borderwidth=1)
            card.grid(row=row, column=col, padx=6, pady=6, sticky="ew")
            card.columnconfigure(0, weight=1)
            # Lagr kode på card for drop-sjekk
            card.drop_code = code  # type: ignore[attr-defined]
            a07_value = float(self.a07_sums.get(code, 0.0))
            gl_value = float(gl_sums.get(code, 0.0))
            diff = a07_value - gl_value
            # Header
            header = ttk.Frame(card)
            header.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
            ttk.Label(header, text=code, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, anchor="w")
            # Totals
            totals = ttk.Frame(card)
            totals.grid(row=1, column=0, sticky="ew", padx=4)
            ttk.Label(totals, text=f"A07: {a07_value:,.2f}".replace(",", " ").replace(".", ","), foreground="#424242").pack(side=tk.LEFT)
            ttk.Label(totals, text=f"GL: {gl_value:,.2f}".replace(",", " ").replace(".", ","), foreground="#424242").pack(side=tk.LEFT, padx=(8, 0))
            colour = "#2e7d32" if abs(diff) < 1e-2 else "#c62828"
            ttk.Label(totals, text=f"Diff: {diff:,.2f}".replace(",", " ").replace(".", ","), foreground=colour).pack(side=tk.LEFT, padx=(8, 0))
            # Tomt drop-område – farger kan endres ved drag-over
            drop_area = ttk.Frame(card)
            drop_area.grid(row=2, column=0, sticky="ew", padx=4, pady=(2, 2))
            drop_area.drop_code = code  # type: ignore[attr-defined]
            # Liste over tilordnede kontoer til denne koden
            accounts_for_code = mapped_accounts.get(code, [])
            if accounts_for_code:
                list_frame = ttk.Frame(card)
                list_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(0, 4))
                # vis inntil 5 rader i listen
                height = min(len(accounts_for_code), 5)
                lstbox = tk.Listbox(list_frame, height=height, activestyle="none")
                for accno_, amt_ in accounts_for_code:
                    lstbox.insert(tk.END, f"{accno_}: {amt_:,.2f}".replace(",", " ").replace(".", ","))
                lstbox.pack(side=tk.TOP, fill=tk.X)
            # neste kort-posisjon
            if col == 1:
                col = 0
                row += 1
            else:
                col = 1

    # -----------------------------------------------------------------
    # Drag‑and‑drop håndtering
    # -----------------------------------------------------------------
    def _on_tree_press(self, event):
        """Start potensiell drag ved å huske valgt konto."""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        vals = self.tree.item(item, "values")
        if not vals:
            return
        self._drag_acc = vals[0]
        # Lag en label som følger musa for visuell tilbakemelding
        if self._drag_label is None:
            self._drag_label = tk.Label(self, text=self._drag_acc, bg="#607d8b", fg="white")

    def _on_drag_motion(self, event):
        """Oppdater posisjonen til drage‑label."""
        if self._drag_label:
            # Flytt label i forhold til vinduets koordinater
            self._drag_label.place(x=event.x_root - self.winfo_rootx() + 10,
                                   y=event.y_root - self.winfo_rooty() + 10)

    def _on_drag_release(self, event):
        """Slipp: sjekk om vi er over et kodekort og map kontoen hvis så."""
        # Fjern drag-label om den finnes
        if self._drag_label:
            self._drag_label.destroy()
            self._drag_label = None
        accno = self._drag_acc
        self._drag_acc = None
        if not accno:
            return
        # Finn widget under pekeren
        x = event.x_root
        y = event.y_root
        target = self.winfo_containing(x, y)
        code = None
        # Gå opp i hierarkiet for å finne et card/drop-område med 'drop_code'
        while target is not None and target is not self:
            if hasattr(target, "drop_code"):
                code = getattr(target, "drop_code")
                break
            target = target.master  # type: ignore[attr-defined]
        if not code:
            return
        # Finn GLAccount-objektet med gitt kontonummer
        for acc in self.accounts:
            if acc.konto == accno:
                try:
                    self.on_map(acc, code)
                except Exception:
                    pass
                break
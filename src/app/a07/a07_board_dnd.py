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

# Import GLAccount and helper from our package.  Try relative import first;
# if that fails (e.g. when running without package context), adjust sys.path
try:
    from models import GLAccount, summarize_gl_by_code  # type: ignore
except Exception:
    try:
        from .models import GLAccount, summarize_gl_by_code  # type: ignore
    except Exception:
        import os, sys
        # Add the parent directory (../) to sys.path to locate models.py
        _cur = os.path.dirname(__file__)
        _parent = os.path.abspath(os.path.join(_cur, '..'))
        if _parent not in sys.path:
            sys.path.insert(0, _parent)
        from models import GLAccount, summarize_gl_by_code  # type: ignore

# Attempt to import TkinterDnD2 for native drag-and-drop support.
try:
    from tkinterdnd2 import DND_TEXT, COPY  # type: ignore
    HAVE_DND = True
except Exception:
    # If TkinterDnD2 is not available, fall back to manual DnD
    HAVE_DND = False


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
        # Define tags for row highlighting.  Green indicates that the
        # mapped code is fully reconciled (diff ~ 0) and yellow indicates
        # that the account is mapped but the code still has a difference.
        self.tree.tag_configure("complete", background="#e8f5e9")
        self.tree.tag_configure("partial", background="#fff8e1")
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Bind drag‑events på kontolisten
        if HAVE_DND:
            # Register the Treeview as a native drag source.  When the user
            # drags with mouse button 1, <<DragInitCmd>> will be invoked.
            try:
                # The DnD API will add methods to widgets if the root window
                # derives from TkinterDnD.Tk.  If not, registration is a no-op.
                self.tree.drag_source_register(1, DND_TEXT)
                self.tree.dnd_bind("<<DragInitCmd>>", self._on_dnd_start)
            except Exception:
                # If registration fails, fall back to manual drag events.
                self.tree.bind("<ButtonPress-1>", self._on_tree_press)
                self.tree.bind("<B1-Motion>", self._on_drag_motion)
                self.tree.bind("<ButtonRelease-1>", self._on_drag_release)
        else:
            # Manual drag-and-drop for environments without TkinterDnD2
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
        # Beregn GL-summer per kode og differanse mot A07 for tagging.
        # Vi bruker summarize_gl_by_code for å beregne GL-summer basert på
        # gjeldende mapping og valgt basis.
        try:
            gl_sums = summarize_gl_by_code(self.accounts, self.mapping, basis=self.basis)
        except Exception:
            gl_sums = {}
        # Lag diff-map for hver kode: A07-sum minus GL-sum
        diff_map: Dict[str, float] = {}
        for code, a07_val in self.a07_sums.items():
            gl_val = float(gl_sums.get(code, 0.0))
            diff_map[code] = float(a07_val) - gl_val
        # Legg til filtrerte kontoer med eventuelle highlight-tags
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
            # Bestem tag basert på mapping og diff
            tag = None
            code = self.mapping.get(acc.konto)
            if code:
                # Et beløp anses som fullstendig avstemt hvis diff for koden er
                # svært liten (mindre enn 1 kr).  Ellers markeres som delvis.
                if abs(diff_map.get(code, 0.0)) < 1.0:
                    tag = "complete"
                else:
                    tag = "partial"
            self.tree.insert(
                "",
                tk.END,
                values=(acc.konto, acc.navn, f"{amount:,.2f}".replace(",", " ").replace(".", ",")),
                tags=(tag,) if tag else (),
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
            # Add a small label showing which basis (UB, Endring, Beløp, Auto) is being used
            # Convert internal basis code to a user-friendly label
            basis_map = {
                "ub": "UB",
                "endring": "Endring",
                "belop": "Beløp",
                "ib": "IB",
                "auto": "Auto",
            }
            display_basis = basis_map.get(self.basis.lower(), self.basis)
            ttk.Label(totals, text=f"Basis: {display_basis}", foreground="#424242").pack(side=tk.LEFT, padx=(8, 0))
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
            # If TkinterDnD2 is available, register this card as a drop target
            if HAVE_DND:
                try:
                    card.drop_target_register(DND_TEXT)
                    card.dnd_bind("<<Drop>>", lambda e, c=code: self._on_dnd_drop(e, c))
                except Exception:
                    pass
            # neste kort-posisjon
            if col == 1:
                col = 0
                row += 1
            else:
                col = 1

        # Legg til en egen "fjern mapping"-boks som drop-target for unmapping.
        # Denne boksen lar brukeren dra en konto hit for å fjerne tilordningen.
        # Den vises som et eget kort nederst på brettet.
        unmap_card = ttk.Frame(self.cards_container, relief=tk.RIDGE, borderwidth=1)
        unmap_card.grid(row=row, column=col, padx=6, pady=6, sticky="ew")
        unmap_card.columnconfigure(0, weight=1)
        # Lagre en tom drop_code for å indikere unmapping
        unmap_card.drop_code = ""  # type: ignore[attr-defined]
        header_un = ttk.Frame(unmap_card)
        header_un.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        ttk.Label(header_un, text="Fjern mapping", font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, anchor="w")
        ttk.Label(unmap_card, text="Slipp konto her for å fjerne mapping", foreground="#555").grid(row=1, column=0, sticky="w", padx=4, pady=(0,4))
        # Registrer som drop target dersom TkinterDnD2 er tilgjengelig
        if HAVE_DND:
            try:
                unmap_card.drop_target_register(DND_TEXT)
                unmap_card.dnd_bind("<<Drop>>", lambda e, c="": self._on_dnd_drop(e, c))
            except Exception:
                pass

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
        if code is None:
            return
        # Hvis koden er tom streng (""), fjern mapping for denne kontoen
        if code == "":
            for acc in self.accounts:
                if acc.konto == accno:
                    try:
                        self.on_map(acc, "")
                    except Exception:
                        pass
                    break
            return
        # Ellers utfør mapping til koden
        for acc in self.accounts:
            if acc.konto == accno:
                try:
                    self.on_map(acc, code)
                except Exception:
                    pass
                break

    # --- TkinterDnD2 handler to initiate drag ---
    def _on_dnd_start(self, event):
        """Called when a drag is initiated on the treeview.

        Returns a tuple (actions, types, data) describing the drag.  The data
        is the account number of the row being dragged.  This method is only
        used when TkinterDnD2 is available.
        """
        # Determine which row is being dragged: prefer current selection.
        item_id = None
        sel = self.tree.selection()
        if sel:
            item_id = sel[0]
        else:
            item_id = self.tree.identify_row(event.y)
        if not item_id:
            return None
        vals = self.tree.item(item_id, "values")
        if not vals:
            return None
        accno = str(vals[0])
        # Remember the account number for potential fallback
        self._drag_acc = accno
        # Provide copy action, text type, and account number as data
        try:
            return (COPY, DND_TEXT, accno)
        except Exception:
            return ("copy", "text/plain", accno)

    # --- TkinterDnD2 handler to process drop ---
    def _on_dnd_drop(self, event, code: str):
        """Handle a drop on a code card when using TkinterDnD2.

        The event's data contains the account number as a string.  Maps the
        corresponding GLAccount to the specified A07 code and returns the
        desired drop action.
        """
        try:
            accno = str(event.data).strip().strip("{}")
        except Exception:
            accno = None
        if not accno:
            return "refuse_drop"
        # Find the GLAccount matching this account number
        for acc in self.accounts:
            if acc.konto == accno:
                try:
                    self.on_map(acc, code)
                except Exception:
                    pass
                break
        return "copy"
# a07_board.py
"""
a07_board.py
==========

This module defines a reusable ``A07Board`` widget for displaying and
interacting with A07 income codes and general ledger accounts.  The
component consists of two panes: a list of accounts on the left and a set
of code cards on the right.  Users can select an account in the list and
then assign it to an A07 code by clicking on a corresponding code card.

The board is intentionally kept independent of application logic.  An
``on_map`` callback is supplied by the caller to handle the mapping of
accounts to codes.  The board itself only provides the user interface and
invokes the callback with the selected account and the code that was
clicked.

Usage example::

    import tkinter as tk
    from models import A07Parser, read_gl_csv
    from board import A07Board

    def handle_mapping(account, code):
        print(f"Map {account.konto} to {code}")

    root = tk.Tk()
    board = A07Board(root, on_map=handle_mapping)
    board.pack(fill=tk.BOTH, expand=True)
    board.update(accounts, a07_sums, mapping)
    root.mainloop()

"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional
import re

# Attempt to import TkinterDnD2 for native drag-and-drop support.
try:
    from tkinterdnd2 import DND_TEXT, COPY
    HAVE_DND = True
except Exception:
    HAVE_DND = False

from models import GLAccount, summarize_gl_by_code


class A07Board(ttk.Frame):
    """Interactive board for mapping GL accounts to A07 codes.

    The board displays a list of general ledger accounts on the left and a
    set of A07 code cards on the right.  Selecting an account and then
    clicking on a code card assigns that account to the chosen code.  The
    board does not modify the mapping itself; instead it invokes an
    `on_map` callback provided by the caller.  This design keeps
    presentation separate from data handling.

    Args:
        master: Parent widget.
        on_map: Callback invoked when the user maps an account to a code.
            It should have the signature `on_map(account: GLAccount, code: str)`.
    """

    def __init__(self, master: tk.Widget, *, on_map: Callable[[GLAccount, str], None]):
        super().__init__(master)
        self.on_map = on_map
        self.accounts: List[GLAccount] = []
        self.a07_sums: Dict[str, float] = {}
        self.mapping: Dict[str, str] = {}
        self.basis: str = "endring"
        # Left pane: account list
        left_frame = ttk.Frame(self)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 4), pady=8)
        # Search box for filtering accounts (optional)
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        ttk.Label(search_frame, text="Søk konto:").pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self._search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=(4, 0))
        search_entry.bind("<KeyRelease>", lambda e: self.refresh_accounts())
        # Treeview for accounts
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
        # Right pane: code cards
        right_frame = ttk.Frame(self)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 8), pady=8)
        # Use a canvas with a vertical scrollbar to allow many cards
        self.canvas = tk.Canvas(right_frame, highlightthickness=0)
        self.cards_container = ttk.Frame(self.canvas)
        self.scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas_window = self.canvas.create_window((0, 0), window=self.cards_container, anchor="nw")
        # Adjust scroll region when the inner frame changes size
        self.cards_container.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.cards_container.bind(
            "<Configure>", lambda e: self.canvas.itemconfigure(self._canvas_window, width=self.canvas.winfo_width())
        )

    def update(self, accounts: List[GLAccount], a07_sums: Dict[str, float], mapping: Dict[str, str], basis: str = "endring") -> None:
        """Update the board with new data.

        Args:
            accounts: List of `GLAccount` objects to display.
            a07_sums: Dictionary mapping A07 codes to sums from the A07 report.
            mapping: Current mapping from account numbers to codes.
            basis: Which amount field to use from the `GLAccount` objects
                (`"endring"`, `"ub"` or `"belop"`).
        """
        self.accounts = list(accounts)
        self.a07_sums = dict(a07_sums)
        self.mapping = dict(mapping)
        self.basis = basis
        self.refresh_accounts()
        self.refresh_codes()

    def refresh_accounts(self) -> None:
        """Refresh the account list based on the current search filter."""
        query = (self._search_var.get() or "").strip().lower()
        # Clear existing rows
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        # Populate with filtered accounts
        for acc in self.accounts:
            if query and (query not in acc.konto.lower() and query not in acc.navn.lower()):
                continue
            # Choose the basis field for display
            if self.basis == "ub":
                amount = acc.ub
            elif self.basis == "belop":
                amount = acc.belop
            else:
                amount = acc.endring
            self.tree.insert("", tk.END, values=(acc.konto, acc.navn, f"{amount:,.2f}".replace(",", " ").replace(".", ",")))

    def refresh_codes(self) -> None:
        """Rebuild the A07 code cards based on sums and current mapping."""
        # Destroy existing cards
        for child in self.cards_container.winfo_children():
            child.destroy()
        # Compute GL sums per code using the selected basis
        gl_sums = summarize_gl_by_code(self.accounts, self.mapping, basis=self.basis)
        # Determine which codes to display: union of A07 codes and mapped codes
        codes = set(self.a07_sums) | set(gl_sums)
        # Sort codes by descending absolute difference
        def sort_key(code: str) -> float:
            a07 = float(self.a07_sums.get(code, 0.0))
            gl = float(gl_sums.get(code, 0.0))
            return -abs(a07 - gl)
        sorted_codes = sorted(codes, key=sort_key)
        # Build a card for each code
        row = 0
        col = 0
        for code in sorted_codes:
            card = ttk.Frame(self.cards_container, relief=tk.RIDGE, borderwidth=1)
            card.grid(row=row, column=col, padx=6, pady=6, sticky="ew")
            card.columnconfigure(0, weight=1)
            a07_value = float(self.a07_sums.get(code, 0.0))
            gl_value = float(gl_sums.get(code, 0.0))
            diff = a07_value - gl_value
            # Header: code label
            header = ttk.Frame(card)
            header.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
            ttk.Label(header, text=code, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, anchor="w")
            # Totals line
            totals = ttk.Frame(card)
            totals.grid(row=1, column=0, sticky="ew", padx=4)
            ttk.Label(totals, text=f"A07: {a07_value:,.2f}".replace(",", " ").replace(".", ","), foreground="#424242").pack(side=tk.LEFT)
            ttk.Label(totals, text=f"GL: {gl_value:,.2f}".replace(",", " ").replace(".", ","), foreground="#424242").pack(side=tk.LEFT, padx=(8, 0))
            # Diff label with colour based on whether it's within tolerance
            colour = "#2e7d32" if abs(diff) < 1e-2 else "#c62828"
            ttk.Label(totals, text=f"Diff: {diff:,.2f}".replace(",", " ").replace(".", ","), foreground=colour).pack(side=tk.LEFT, padx=(8, 0))
            # Map button
            map_btn = ttk.Button(card, text="Tilordne valgt konto", command=lambda c=code: self._map_selected(c))
            map_btn.grid(row=2, column=0, sticky="ew", padx=4, pady=(2, 4))
            # Move to next column/row
            if col == 1:
                col = 0
                row += 1
            else:
                col = 1

    def _map_selected(self, code: str) -> None:
        """Invoke the mapping callback for the currently selected account."""
        account = self.get_selected_account()
        if not account:
            return
        try:
            self.on_map(account, code)
        except Exception:
            # Swallow exceptions from the callback to avoid crashing the UI
            pass

    def get_selected_account(self) -> Optional[GLAccount]:
        """Return the `GLAccount` corresponding to the selected row in the tree."""
        selection = self.tree.selection()
        if not selection:
            return None
        item_id = selection[0]
        values = self.tree.item(item_id, "values")
        if not values:
            return None
        konto = str(values[0])
        for acc in self.accounts:
            if acc.konto == konto:
                return acc
        return None


class AssignmentBoard(ttk.Frame):
    """Interactive drag-and-drop board for mapping GL accounts to A07 codes.

    This board displays GL accounts on the left and A07 codes on the right as cards.
    Users can drag an account and drop it onto a code card to assign the account to that code.
    A special 'remove mapping' card allows unassigning an account by dropping it there.
    """

    def __init__(self, master: tk.Widget, *, get_amount_fn: Optional[Callable[[GLAccount], float]] = None,
                 on_drop: Optional[Callable[[str, str], None]] = None,
                 request_suggestions: Optional[Callable[[], None]] = None):
        super().__init__(master)
        self.get_amount_fn = get_amount_fn
        self.on_drop_cb = on_drop
        self.request_suggestions_cb = request_suggestions
        self.accounts: List[GLAccount] = []
        self.a07_sums: Dict[str, float] = {}
        self.mapping: Dict[str, Any] = {}  # account -> code or set of codes
        self.basis: str = "endring"
        # Drag-state
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
        # Bind drag-events på kontolisten
        if HAVE_DND:
            try:
                # Register the Treeview as a native drag source.  When the user
                # drags with mouse button 1, <<DragInitCmd>> will be invoked.
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

    def update(self, accounts: List[GLAccount], a07_sums: Dict[str, float], mapping: Dict[str, Any], basis: str = "endring") -> None:
        """Oppdater brettet med nye data og bygg kontoer og koder.

        Args:
            accounts: Liste av ``GLAccount``-objekter.
            a07_sums: Summer pr A07-kode.
            mapping: Gjeldende mapping (konto -> kode).
            basis: Hvilket feltnavn på ``GLAccount`` som skal brukes for beløp.
        """
        self.accounts = list(accounts)
        self.a07_sums = dict(a07_sums)
        self.mapping = mapping  # allow mapping to be dict[str, str] or dict[str, set]
        self.basis = basis
        self.refresh_accounts()
        self.refresh_codes()

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
            accno = str(acc.konto if hasattr(acc, 'konto') else acc.get('konto'))
            if query and (query not in accno.lower() and query not in (acc.navn if hasattr(acc, 'navn') else acc.get('navn', '')).lower()):
                continue
            # Velg basisfeltet
            if self.basis.lower() == "ub":
                amount = float(acc.ub) if hasattr(acc, 'ub') else float(acc.get('ub', 0.0))
            elif self.basis.lower() == "belop":
                amount = float(acc.belop) if hasattr(acc, 'belop') else float(acc.get('belop', 0.0))
            else:
                amount = float(acc.endring) if hasattr(acc, 'endring') else float(acc.get('endring', acc.get('belop', 0.0)))
            # Bestem tag basert på mapping og diff
            tag = None
            codes = self.mapping.get(acc.konto if hasattr(acc, 'konto') else accno)
            if codes:
                if isinstance(codes, (set, list)):
                    # Et beløp anses som fullstendig avstemt hvis diff for koden er
                    # svært liten (mindre enn 1 kr).  Ellers markeres som delvis.
                    tag = "partial"
                    # (Hvis flere koder per konto, markeres som partial uansett.)
                else:
                    code = str(codes)
                    if abs(diff_map.get(code, 0.0)) < 1.0:
                        tag = "complete"
                    else:
                        tag = "partial"
            self.tree.insert(
                "",
                tk.END,
                values=(accno, (acc.navn if hasattr(acc, 'navn') else acc.get('navn', '')), f"{amount:,.2f}".replace(",", " ").replace(".", ",")),
                tags=(tag,) if tag else ()
            )

    def refresh_codes(self) -> None:
        """Bygg kodekort basert på summer fra A07 og GL."""
        # Fjern eksisterende kort
        for child in self.cards_container.winfo_children():
            child.destroy()
        # Summér GL-pr-kode basert på mapping
        gl_sums: Dict[str, float] = {}
        for acc in self.accounts:
            accno = str(acc.konto if hasattr(acc, 'konto') else acc.get('konto'))
            if self.get_amount_fn:
                val = float(self.get_amount_fn(acc))
            else:
                if self.basis.lower() == "ub":
                    val = float(acc.ub) if hasattr(acc, 'ub') else float(acc.get('ub', 0.0))
                elif self.basis.lower() == "belop":
                    val = float(acc.belop) if hasattr(acc, 'belop') else float(acc.get('belop', 0.0))
                else:
                    val = float(acc.endring) if hasattr(acc, 'endring') else float(acc.get('endring', 0.0))
            codes = self.mapping.get(acc.konto if hasattr(acc, 'konto') else accno)
            if not codes:
                continue
            if isinstance(codes, (set, list)):
                for code in codes:
                    gl_sums[code] = gl_sums.get(code, 0.0) + val
            else:
                code = str(codes)
                gl_sums[code] = gl_sums.get(code, 0.0) + val
        # Lag en oversikt over hvilke kontoer som er mappet til hver kode, med valgt basis-beløp
        mapped_accounts: Dict[str, List[tuple[str, float]]] = {}
        for acc in self.accounts:
            accno = str(acc.konto if hasattr(acc, 'konto') else acc.get('konto'))
            codes = self.mapping.get(acc.konto if hasattr(acc, 'konto') else accno)
            if not codes:
                continue
            if isinstance(codes, (set, list)):
                for code in codes:
                    # Bruker full kontoverdi per kode (merk: dobbeltsummering kan forekomme hvis multi-mappet)
                    if self.basis.lower() == "ub":
                        amount = float(acc.ub) if hasattr(acc, 'ub') else float(acc.get('ub', 0.0))
                    elif self.basis.lower() == "belop":
                        amount = float(acc.belop) if hasattr(acc, 'belop') else float(acc.get('belop', 0.0))
                    else:
                        amount = float(acc.endring) if hasattr(acc, 'endring') else float(acc.get('endring', 0.0))
                    mapped_accounts.setdefault(code, []).append((accno, float(amount)))
            else:
                code = str(codes)
                if self.basis.lower() == "ub":
                    amount = float(acc.ub) if hasattr(acc, 'ub') else float(acc.get('ub', 0.0))
                elif self.basis.lower() == "belop":
                    amount = float(acc.belop) if hasattr(acc, 'belop') else float(acc.get('belop', 0.0))
                else:
                    amount = float(acc.endring) if hasattr(acc, 'endring') else float(acc.get('endring', 0.0))
                mapped_accounts.setdefault(code, []).append((accno, float(amount)))
        codes = set(self.a07_sums) | set(gl_sums)
        def sort_key(code: str) -> float:
            a07_val = float(self.a07_sums.get(code, 0.0))
            gl_val = float(gl_sums.get(code, 0.0))
            return -abs(a07_val - gl_val)
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
            # Basis label
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
        ttk.Label(unmap_card, text="Slipp konto her for å fjerne mapping", foreground="#555").grid(row=1, column=0, sticky="w", padx=4, pady=(0, 4))
        # Registrer som drop target dersom TkinterDnD2 er tilgjengelig
        if HAVE_DND:
            try:
                unmap_card.drop_target_register(DND_TEXT)
                unmap_card.dnd_bind("<<Drop>>", lambda e, c="": self._on_dnd_drop(e, c))
            except Exception:
                pass

    # -----------------------------------------------------------------
    # Drag-and-drop håndtering
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
        """Oppdater posisjonen til drage-label."""
        if self._drag_label:
            # Flytt label i forhold til vinduets koordinater
            self._drag_label.place(x=event.x_root - self.winfo_rootx() + 10,
                                   y=event.y_rooty() + 10)
# (fortsetter a07_board.py)
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
            target = getattr(target, "master", None)
        if code is None:
            return
        # Hvis koden er tom streng ntoen
        if code == "":
            if self.on_drop_cb:
                try:
                    self.on_drop_cb(accno, "")
                except Exception:
                    pass
            return
        # Ellers utfør mapping til koden
        if self.on_drop_cb:
            try:
                self.on_drop_cb(accno, str(code))
            except Exception:
                pass

    def _on_dnd_start(self, event):
        """Called when a drag is initiated on the treeview.

        Returns a tupltypes, data) describing the drag.  The data
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
  :contentReference[oaicite:3]{index=3}:contentReference[oaicite:4]{index=4}cno
        # Provide copy action, text type, and account number as data
        try:
            return (COPY, DND_TEXT, accno)
        except Exception:
            return ("copy", "text/plain", accno)

    def _on_dnd_drop(self, event, code: str):
:contentReference[oaicite:5]{index=5}dle a drop on a code card when using TkinterDnD2.

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
        # Find the GLAccount (or dict) matching :contentReference[oaicite:6]{index=6}umber
        for acc in self.accounts:
            if str(acc.konto if hasattr(acc, 'konto') else acc.get('konto')) == accno:
  :contentReference[oaicite:7]{index=7}if self.on_drop_cb:
                    try:
                        self.on_drop_cb(accno, str(code))
                    except Exception:
                        pass
                break
    :contentReference[oaicite:8]{index=8}py"

    def supply_data(self, accounts: List[Any], acc_to_code: Dict[str, Any], suggestions: Dict[str, Any] = None, a07_sums: Dict[str, float] = None, diff_threshold: float = 0.0, only_unmapped: bool = False) -> None:
        """Supply data to the board (for compatibility with A07Gui integration)."""
        # If only_unmapped flag is set, filter accounts to only those without mapping
        acc:contentReference[oaicite:9]{index=9}:contentReference[oaicite:10]{index=10}    if only_unmapped:
            acct_list = []
            for acc in accounts:
                accno = str(acc.get("konto") if isinstance(acc, dict) else getattr(acc, "konto", ""))
                if accno and not (accno in acc_to_code and acc_to_code.get(accno)):
                    acct_list.append(acc)
        # Use provided A07 sums or stored sums
        a07s = a07_sums if a07_sums is not None else self.a07_sums
        # Update board
        self.update(acct_list, a07s, acc_to_code, basis=self.basis)

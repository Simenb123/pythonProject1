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

from models import GLAccount, summarize_gl_by_code


class A07Board(ttk.Frame):
    """Interactive board for mapping GL accounts to A07 codes.

    The board displays a list of general ledger accounts on the left and a
    set of A07 code cards on the right.  Selecting an account and then
    clicking on a code card assigns that account to the chosen code.  The
    board does not modify the mapping itself; instead it invokes an
    ``on_map`` callback provided by the caller.  This design keeps
    presentation separate from data handling.

    Args:
        master: Parent widget.
        on_map: Callback invoked when the user maps an account to a code.
            It should have the signature ``on_map(account: GLAccount, code: str)``.
    """

    def __init__(self, master: tk.Widget, *, on_map: Callable[[GLAccount, str], None]):
        super().__init__(master)
        self.on_map = on_map
        self.accounts: List[GLAccount] = []
        self.a07_sums: Dict[str, float] = {}
        self.mapping: Dict[str, str] = {}
        self.basis: str = "endring"
        # Drag‑and‑drop state
        self._drag_acc: Optional[str] = None  # kontonummer som dras
        self._drag_label: Optional[tk.Label] = None  # flytende label under drag
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
        # Bind drag-n-drop events
        self.tree.bind("<ButtonPress-1>", self._on_tree_press)
        self.tree.bind("<B1-Motion>", self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_release)
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
            accounts: List of ``GLAccount`` objects to display.
            a07_sums: Dictionary mapping A07 codes to sums from the A07 report.
            mapping: Current mapping from account numbers to codes.
            basis: Which amount field to use from the ``GLAccount`` objects
                (``"endring"``, ``"ub"`` or ``"belop"``).
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
        # Build list of mapped accounts per code for display
        mapped_accounts: Dict[str, List[tuple[str, float]]] = {}
        for acc in self.accounts:
            code = self.mapping.get(acc.konto)
            if not code:
                continue
            # choose amount based on basis
            if self.basis == "ub":
                amount = acc.ub
            elif self.basis == "belop":
                amount = acc.belop
            else:
                amount = acc.endring
            mapped_accounts.setdefault(code, []).append((acc.konto, float(amount)))
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
            # Drop‑område for drag‑and‑drop
            drop_area = ttk.Frame(card)
            drop_area.grid(row=2, column=0, sticky="ew", padx=4, pady=(2, 2))
            drop_area.drop_code = code  # type: ignore[attr-defined]
            # Vis liste over tilordnede kontoer
            accounts_for_code = mapped_accounts.get(code, [])
            if accounts_for_code:
                list_frame = ttk.Frame(card)
                list_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=(0, 4))
                height = min(len(accounts_for_code), 5)
                lstbox = tk.Listbox(list_frame, height=height, activestyle="none")
                for accno_, amt_ in accounts_for_code:
                    lstbox.insert(tk.END, f"{accno_}: {amt_:,.2f}".replace(",", " ").replace(".", ","))
                lstbox.pack(side=tk.TOP, fill=tk.X)
            # Map button for click-mapping (valgfritt)
            map_btn = ttk.Button(card, text="Tilordne valgt konto", command=lambda c=code: self._map_selected(c))
            map_btn.grid(row=4, column=0, sticky="ew", padx=4, pady=(0, 4))
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

    # -----------------------------------------------------------------
    # Drag‑and‑drop event handlers
    # -----------------------------------------------------------------
    def _on_tree_press(self, event):
        """Start drag: memorize selected account."""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        vals = self.tree.item(item, "values")
        if not vals:
            return
        self._drag_acc = vals[0]
        if self._drag_label is None:
            self._drag_label = tk.Label(self, text=self._drag_acc, bg="#607d8b", fg="white")

    def _on_drag_motion(self, event):
        """Update drag label position."""
        if self._drag_label:
            self._drag_label.place(x=event.x_root - self.winfo_rootx() + 10,
                                   y=event.y_root - self.winfo_rooty() + 10)

    def _on_drag_release(self, event):
        """Handle drop: map account if dropped over a card."""
        if self._drag_label:
            self._drag_label.destroy()
            self._drag_label = None
        accno = self._drag_acc
        self._drag_acc = None
        if not accno:
            return
        # Determine widget under cursor
        target = self.winfo_containing(event.x_root, event.y_root)
        code = None
        while target is not None and target is not self:
            if hasattr(target, "drop_code"):
                code = getattr(target, "drop_code")
                break
            target = target.master  # type: ignore[attr-defined]
        if not code:
            return
        # find GLAccount by account number
        for acc in self.accounts:
            if acc.konto == accno:
                try:
                    self.on_map(acc, code)
                except Exception:
                    pass
                break

    def get_selected_account(self) -> Optional[GLAccount]:
        """Return the ``GLAccount`` corresponding to the selected row in the tree."""
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

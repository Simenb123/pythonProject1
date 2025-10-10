"""
A prototype of a drag-and-drop board that supports grouping multiple A07 codes
into a single "GroupCard".  The GroupCard displays the combined A07 sum
and GL sum for all included codes.  Accounts dropped on a group card are
allocated equally across the codes in the group.

This file is meant as a starting point for extending the existing GUI.
It does not replace the existing board, but demonstrates how groups might
be represented and interacted with.  Integration into the larger app
requires changes to mapping structures and control tables.
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Callable, Iterable


class GLAccount:
    """A minimal stand‑in for the GLAccount class used in the app."""
    def __init__(self, konto: str, navn: str, belop: float):
        self.konto = konto
        self.navn = navn
        self.endring = belop
        self.ub = belop
        self.belop = belop


class GroupCard(ttk.Frame):
    """A widget representing a group of A07 codes.

    Attributes:
        codes: list of A07 code strings in this group
        on_drop: callback invoked when a GL account is dropped on the card
        on_remove: callback invoked to remove this group
    """

    def __init__(self, master, codes: Iterable[str], *,
                 on_drop: Callable[[GLAccount, List[str]], None],
                 on_remove: Callable[[List[str]], None], **kwargs):
        super().__init__(master, relief=tk.RIDGE, borderwidth=2, **kwargs)
        self.codes = list(codes)
        self.on_drop = on_drop
        self.on_remove = on_remove
        self._build_ui()

    def _build_ui(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(2, 0))
        title = ttk.Label(header, text=" + ".join(self.codes),
                           font=("TkDefaultFont", 10, "bold"))
        title.pack(side=tk.LEFT, padx=4)
        rm_btn = ttk.Button(header, text="×", width=2,
                            command=self._remove)
        rm_btn.pack(side=tk.RIGHT, padx=4)

        self.info = ttk.Label(self, text="A07: 0.0  GL: 0.0  Diff: 0.0")
        self.info.pack(side=tk.TOP, anchor="w", padx=4, pady=(0, 4))

        drop_lbl = ttk.Label(self, text="Slipp konto her", relief=tk.GROOVE,
                             padding=4)
        drop_lbl.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        # Register drop handler.  In a real app, use tkinterdnd2 to accept drops.
        drop_lbl.bind("<ButtonRelease-1>", self._dummy_drop)

    def update_totals(self, a07_sum: float, gl_sum: float) -> None:
        diff = a07_sum - gl_sum
        self.info.config(text=f"A07: {a07_sum:,.2f}  GL: {gl_sum:,.2f}  Diff: {diff:,.2f}")

    def _dummy_drop(self, event) -> None:
        """Placeholder drop handler.  In real usage, extract account and call on_drop."""
        # Example of invoking callback with a dummy account
        acc = GLAccount("0000", "Dummy", 0.0)
        self.on_drop(acc, self.codes)

    def _remove(self) -> None:
        self.on_remove(self.codes)


class GroupBoard(ttk.Frame):
    """A simple board that can display GL accounts and group cards."""

    def __init__(self, master, *, on_map: Callable[[GLAccount, List[str]], None]):
        super().__init__(master)
        self.on_map = on_map
        self.group_cards: List[GroupCard] = []
        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=2)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        left.rowconfigure(1, weight=1)

        ttk.Label(left, text="Konti").grid(row=0, column=0, sticky="w")
        self.tree = ttk.Treeview(left, columns=("konto", "navn", "belop"),
                                 show="headings", selectmode="browse")
        for cid, hdr in zip(("konto", "navn", "belop"),
                            ("Konto", "Navn", "Beløp")):
            self.tree.heading(cid, text=hdr)
        self.tree.grid(row=1, column=0, sticky="nsew")

        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self.groups_container = ttk.Frame(right)
        self.groups_container.grid(row=0, column=0, sticky="nsew")

        add_btn = ttk.Button(right, text="Ny samlekode", command=self._add_group)
        add_btn.grid(row=1, column=0, sticky="ew", pady=(4, 0))

    def populate_accounts(self, accounts: Iterable[GLAccount]) -> None:
        for row in self.tree.get_children():
            self.tree.delete(row)
        for acc in accounts:
            self.tree.insert("", tk.END, values=(acc.konto, acc.navn, f"{acc.belop:,.2f}"))

    def _add_group(self) -> None:
        """Add a new empty group card for demonstration."""
        # In real usage, you'd ask the user which codes to group together.
        codes = ["kode1", "kode2"]
        card = GroupCard(self.groups_container, codes,
                         on_drop=self.on_map,
                         on_remove=self._remove_group)
        card.pack(side=tk.TOP, fill=tk.X, pady=4)
        self.group_cards.append(card)

    def _remove_group(self, codes: List[str]) -> None:
        for i, card in enumerate(self.group_cards):
            if card.codes == codes:
                card.destroy()
                del self.group_cards[i]
                break


# Example usage
if __name__ == "__main__":
    root = tk.Tk()
    def on_map(acc: GLAccount, codes: List[str]):
        print(f"Map account {acc.konto} to codes {codes}")

    board = GroupBoard(root, on_map=on_map)
    board.pack(fill=tk.BOTH, expand=True)

    # Populate with dummy accounts
    board.populate_accounts([
        GLAccount("5010", "Lønn til ansatte", 12825382.07),
        GLAccount("2920", "Skyldig feriepenger", -527047.02),
    ])

    root.mainloop()

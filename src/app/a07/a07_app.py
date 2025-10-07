"""
app.py
=======

This module provides a minimal but structured Tkinter application for
performing A07 income reconciliation.  It separates data parsing from
presentation by relying on the ``models`` module for reading A07 JSON and
general ledger CSV files, and the ``board`` module for the interactive
mapping user interface.  The application allows the user to load an A07
report and a general ledger, view account balances versus A07 sums and
assign accounts to codes via a simple click interface.

The design emphasises clarity and maintainability.  Each component
(parsing, summarising, presentation) is placed in its own module, and
application state is stored in instance variables rather than globals.

"""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List

from models import A07Parser, GLAccount, read_gl_csv
from board import A07Board


class A07App(tk.Tk):
    """Main application window for the A07 reconciliation tool."""

    def __init__(self) -> None:
        super().__init__()
        self.title("A07 Lønnsavstemming – Enkel versjon")
        self.geometry("1200x700")
        # Data
        self.parser = A07Parser()
        self.a07_rows: List = []
        self.a07_sums: Dict[str, float] = {}
        self.gl_accounts: List[GLAccount] = []
        self.mapping: Dict[str, str] = {}
        # Build UI
        self._build_menu()
        self._build_controls()
        # Board widget
        self.board = A07Board(self, on_map=self._handle_mapping)
        self.board.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def _build_menu(self) -> None:
        """Create a simple menu bar with file loading commands."""
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Åpne A07 JSON…", command=self._open_a07_file)
        filemenu.add_command(label="Åpne GL CSV…", command=self._open_gl_file)
        filemenu.add_separator()
        filemenu.add_command(label="Avslutt", command=self.destroy)
        menubar.add_cascade(label="Fil", menu=filemenu)
        self.config(menu=menubar)

    def _build_controls(self) -> None:
        """Create a status bar and basis selection controls."""
        control_frame = ttk.Frame(self)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=4)
        # Basis selection for GL amounts
        ttk.Label(control_frame, text="Regnskapsbasis:").pack(side=tk.LEFT)
        self._basis_var = tk.StringVar(value="endring")
        for val, txt in [("endring", "Endring"), ("ub", "UB"), ("belop", "Beløp")]:
            ttk.Radiobutton(
                control_frame,
                text=txt,
                variable=self._basis_var,
                value=val,
                command=self._update_board,
            ).pack(side=tk.LEFT, padx=(4, 0))
        # Status label
        self._status = ttk.Label(control_frame, text="Velkommen! Last inn filer.")
        self._status.pack(side=tk.RIGHT, expand=True, fill=tk.X)

    def _open_a07_file(self) -> None:
        """Prompt the user to select and load an A07 JSON file."""
        path = filedialog.askopenfilename(
            parent=self,
            title="Velg A07 JSON-fil",
            filetypes=[("JSON-filer", "*.json"), ("Alle filer", "*.*")],
        )
        if not path:
            return
        rows, errors = self.parser.parse_file(path)
        if errors:
            messagebox.showerror(
                "Feil ved lesing",
                "Det oppstod feil under parsing av A07-filen:\n" + "\n".join(errors),
            )
        self.a07_rows = rows
        self.a07_sums = self.parser.summarize_by_code(rows)
        self._status.configure(text=f"Lest {len(rows)} A07-linjer. Koder: {len(self.a07_sums)}.")
        self._update_board()

    def _open_gl_file(self) -> None:
        """Prompt the user to select and load a general ledger CSV file."""
        path = filedialog.askopenfilename(
            parent=self,
            title="Velg GL CSV-fil",
            filetypes=[("CSV-filer", "*.csv"), ("Alle filer", "*.*")],
        )
        if not path:
            return
        accounts, meta = read_gl_csv(path)
        if not accounts:
            messagebox.showwarning(
                "Tom fil",
                "CSV-filen ser ut til å være tom eller mangler nødvendige kolonner.",
            )
        self.gl_accounts = accounts
        # Reset mapping when new accounts are loaded
        self.mapping = {}
        self._status.configure(
            text=f"Lest {len(accounts)} GL-konti. Delimiter: '{meta.get('delimiter')}'. Basis: {self._basis_var.get()}"
        )
        self._update_board()

    def _handle_mapping(self, account: GLAccount, code: str) -> None:
        """Called when the user assigns a GL account to an A07 code."""
        self.mapping[account.konto] = code
        self._status.configure(
            text=f"Konto {account.konto} ble tilordnet kode {code}."
        )
        self._update_board()

    def _update_board(self) -> None:
        """Update the board widget with current data and selections."""
        basis = self._basis_var.get()
        self.board.update(self.gl_accounts, self.a07_sums, self.mapping, basis=basis)


def main() -> None:
    """Entry point to run the application standalone."""
    app = A07App()
    app.mainloop()


if __name__ == "__main__":
    main()

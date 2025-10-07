"""
a07_app_full.py
================

Dette er en utvidet versjon av den enkle A07‑applikasjonen.  Den
kombinerer det modulære oppsettet (separate filer for datamodeller,
parser og brett) med funksjoner fra den gamle GUI‑en, som:

* Mulighet til å laste regelbok (Excel eller CSV‑mappe) og generere
  automatiske mappingforslag.
* Fallback‑forslag når regelbok ikke finnes (krever en modul
  ``matcher_fallback`` – hvis denne ikke er tilgjengelig vil auto‑match
  gjøre ingenting).
* Global beløpsmatching (LP‑optimering) hvor hvert konto­beløp kan
  fordeles på flere koder.  Det brukes PuLP som solver.
* Drag‑and‑drop i brettet via ``a07_board_dnd.py``.

Filen kan brukes som drop‑in erstatning for ``a07_app.py``.  Den vil
fungerer best sammen med ``a07_board_dnd.py`` (som implementerer
drag‑and‑drop) og ``models.py`` (for datamodeller og parsere).

Bruk:

    python a07_app_full.py

Programmet oppretter et hovedvindu med meny for å åpne A07‑JSON og
GL‑CSV, laste regelbok, kjøre auto‑match og global LP.  Status
oppdateres i en statuslinje.  Mappingen kan også redigeres manuelt
ved drag‑and‑drop mellom konto og kodekort.

"""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from models import A07Parser, GLAccount, read_gl_csv
from a07_board_dnd import A07Board  # Drag-and-drop versjon av brettet


class A07App(tk.Tk):
    """Hovedapplikasjon for A07 lønnsavstemming med utvidet funksjonalitet."""

    def __init__(self) -> None:
        super().__init__()
        self.title("A07 Lønnsavstemming – Full versjon")
        self.geometry("1200x700")
        # Data
        self.parser = A07Parser()
        self.a07_rows: List = []
        self.a07_sums: Dict[str, float] = {}
        self.gl_accounts: List[GLAccount] = []
        self.mapping: Dict[str, str] = {}
        self.rulebook: Optional[Dict] = None
        # Build UI
        self._build_menu()
        self._build_controls()
        # Board widget med drag‑and‑drop
        self.board = A07Board(self, on_map=self._handle_mapping)
        self.board.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # -----------------------------------------------------------------
    # Grensesnitt bygging
    # -----------------------------------------------------------------
    def _build_menu(self) -> None:
        """Lag menylinje med filhandlinger, regelbok og matching."""
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Åpne A07 JSON…", command=self._open_a07_file)
        filemenu.add_command(label="Åpne GL CSV…", command=self._open_gl_file)
        filemenu.add_separator()
        filemenu.add_command(label="Last regelbok…", command=self._load_rulebook)
        filemenu.add_command(label="Auto‑match (regelbok/fallback)", command=self._auto_match)
        filemenu.add_command(label="Global matching (LP)", command=self._run_lp)
        filemenu.add_separator()
        filemenu.add_command(label="Avslutt", command=self.destroy)
        menubar.add_cascade(label="Fil", menu=filemenu)
        self.config(menu=menubar)

    def _build_controls(self) -> None:
        """Lag øverste kontrollpanel med basisvalg og status."""
        control_frame = ttk.Frame(self)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=4)
        # Valg av basis (påvirker GL-beløp som vises og brukes i matching)
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
        # Statuslinje
        self._status = ttk.Label(control_frame, text="Velkommen! Last inn filer.")
        self._status.pack(side=tk.RIGHT, expand=True, fill=tk.X)

    # -----------------------------------------------------------------
    # Fil- og regelbokhandlinger
    # -----------------------------------------------------------------
    def _open_a07_file(self) -> None:
        """Spør bruker om A07 JSON-fil og last den inn."""
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
        """Spør bruker om GL CSV-fil og last den inn."""
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
        # Nullstill mapping når nye kontoer lastes
        self.mapping = {}
        self._status.configure(
            text=(
                f"Lest {len(accounts)} GL-konti. Delimiter: '{meta.get('delimiter')}'. "
                f"Basis: {self._basis_var.get()}"
            )
        )
        self._update_board()

    def _load_rulebook(self) -> None:
        """Last regelbok fra Excel eller CSV-mapre og lagre i self.rulebook."""
        from a07_rulebook import load_rulebook  # Importeres her for å unngå sirkulær avhengighet
        path = filedialog.askopenfilename(
            parent=self,
            title="Velg regelbok (Excel eller mappe)",
            filetypes=[
                ("Excel-filer", "*.xlsx"),
                ("CSV/mappe", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self.rulebook = load_rulebook(path)
            self._status.configure(text=f"Regelbok lastet fra {path}.")
        except Exception as e:
            messagebox.showerror("Feil ved lesing", str(e))

    # -----------------------------------------------------------------
    # Mapping og matching
    # -----------------------------------------------------------------
    def _handle_mapping(self, account: GLAccount, code: str) -> None:
        """Kalles når bruker tilordner en konto til en kode via brettet."""
        self.mapping[account.konto] = code
        self._status.configure(text=f"Konto {account.konto} ble tilordnet kode {code}.")
        self._update_board()

    def _auto_match(self) -> None:
        """Kjør auto‑forslag basert på regelbok eller fallback."""
        if not self.gl_accounts or not self.a07_sums:
            messagebox.showinfo("Mangler data", "Last inn både A07 JSON og GL CSV først.")
            return
        # Forsøk regelbok
        suggestions = {}
        if getattr(self, "rulebook", None):
            try:
                from a07_rulebook import load_rulebook, _alias_bag_for_code  # bare for import-sjekk
                from a07_optimize import generate_candidates_for_lp
                # Generer kandidatforslag via regelbok; velg topp-score per konto
                cands = generate_candidates_for_lp(
                    self.gl_accounts,
                    self.a07_sums,
                    self.rulebook,
                    min_name=0.25,
                    min_score=0.40,
                    top_k=1,
                )
                for accno, lst in cands.items():
                    if lst:
                        code, score, amt, reason = lst[0]
                        suggestions[accno] = {"kode": code, "score": score}
            except Exception:
                suggestions = {}
        # Fallback hvis ingen regelbok-suggestions
        if not suggestions:
            try:
                from matcher_fallback import suggest_mapping_for_accounts as fallback_suggest
                suggestions = fallback_suggest(self.gl_accounts, self.a07_sums, min_score=0.60)
            except Exception:
                suggestions = {}
        # Oppdater mapping med forslag
        count = 0
        for acc in self.gl_accounts:
            accno = acc.konto
            if accno in suggestions:
                self.mapping[accno] = suggestions[accno].get("kode", "")
                count += 1
        self._status.configure(text=f"Auto‑match: {count} kontoer tilordnet.")
        self._update_board()

    def _run_lp(self) -> None:
        """Kjør LP‑matching for å fordele GL-beløp mot A07-koder."""
        if not self.gl_accounts or not self.a07_sums:
            messagebox.showinfo("Mangler data", "Last inn både A07 JSON og GL CSV først.")
            return
        from a07_optimize import generate_candidates_for_lp, solve_global_assignment_lp
        basis = self._basis_var.get()
        # Forbered beløp pr konto basert på valgt basis
        amounts: Dict[str, float] = {}
        for acc in self.gl_accounts:
            if basis == "ub":
                amounts[acc.konto] = acc.ub
            elif basis == "belop":
                amounts[acc.konto] = acc.belop
            else:
                amounts[acc.konto] = acc.endring
        # Generer kandidater (bruk regelbok hvis tilgjengelig, ellers tomt)
        rule = getattr(self, "rulebook", {}) or {}
        candidates = generate_candidates_for_lp(
            self.gl_accounts,
            self.a07_sums,
            rule,
            min_name=0.25,
            min_score=0.40,
            top_k=3,
        )
        try:
            assignment = solve_global_assignment_lp(
                amounts,
                candidates,
                self.a07_sums,
                allow_splits=True,
                lambda_score=0.25,
                lambda_sign=0.05,
                lambda_edges=0.00,
            )
        except Exception as e:
            messagebox.showerror("LP‑feil", str(e))
            return
        # Velg hovedkode per konto (høyest andel)
        count = 0
        for accno, parts in assignment.items():
            if not parts:
                continue
            best = max(parts.items(), key=lambda kv: kv[1])[0]
            self.mapping[accno] = best
            count += 1
        self._status.configure(text=f"Global matching fullført: {count} kontoer tilordnet.")
        self._update_board()

    # -----------------------------------------------------------------
    # Intern visning
    # -----------------------------------------------------------------
    def _update_board(self) -> None:
        """Oppdater brettet med gjeldende data og mapping."""
        basis = self._basis_var.get()
        self.board.update(self.gl_accounts, self.a07_sums, self.mapping, basis=basis)


def main() -> None:
    """Kjør applikasjonen standalone."""
    app = A07App()
    app.mainloop()


if __name__ == "__main__":
    main()
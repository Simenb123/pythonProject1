"""
Client overview window.

This module defines a `ClientOverview` class that provides an
overview window for a selected client.  It splits the UI into two
regions: a vertical navigation pane on the left and a main content
area on the right.  The navigation pane contains buttons for the
different functional areas (kildefiler, hovedbok, saldobalanse,
regnskap, analyse, utvalg og eierskap).  The main content area shows
general client information (such as industry code, contact details
and address) along with the team members assigned to the client.

The goal of this module is to decouple the high‑level overview from
the details of source file management.  The `KildefilerView` (see
``kildefiler_view.py``) can be displayed when the user clicks the
``Kildefiler`` button, while the other areas can be implemented
later.

Usage:
    from client_overview import ClientOverview
    root = tk.Tk(); ClientOverview(root, client_name="9762 Vitamail AS").mainloop()

This file does not make any assumptions about the surrounding
application infrastructure; it can be integrated into the existing
codebase by replacing the current `ClientHub` usage.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Dict, Callable

try:
    # VersionsPanel will be used in the Kildefiler view
    from app.gui.widgets.versions_panel import VersionsPanel  # type: ignore
except Exception:
    VersionsPanel = None  # type: ignore


class ClientOverview(tk.Toplevel):
    """Overview window for a client.

    Parameters
    ----------
    master : tk.Widget
        Parent window.
    client_name : str
        Display name of the selected client.
    year : int
        Current revision year.
    """

    def __init__(self, master: tk.Widget, client_name: str, year: int = 2024) -> None:
        super().__init__(master)
        self.title(f"Klientoversikt – {client_name}")
        self.resizable(False, False)
        self.client_name = client_name
        self.year_var = tk.IntVar(value=year)

        # Containers
        root_frame = ttk.Frame(self, padding=10)
        root_frame.grid(row=0, column=0, sticky="nsew")
        root_frame.columnconfigure(1, weight=1)

        # Navigation pane on the left
        nav_frame = ttk.Frame(root_frame)
        nav_frame.grid(row=0, column=0, sticky="ns")

        self._nav_buttons: Dict[str, ttk.Button] = {}
        for idx, label in enumerate([
            "Kildefiler",
            "Hovedbok",
            "Saldobalanse",
            "Regnskap",
            "Analyse",
            "Utvalg",
            "Eierskap",
        ]):
            btn = ttk.Button(nav_frame, text=label, command=lambda l=label: self._on_nav(l))
            btn.grid(row=idx, column=0, sticky="we", pady=2)
            self._nav_buttons[label] = btn

        # Content area on the right
        content_frame = ttk.Frame(root_frame)
        content_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(2, weight=1)

        # Header with revision year selector
        header = ttk.Frame(content_frame)
        header.grid(row=0, column=0, sticky="we")
        ttk.Label(header, text="Revisjonsår:").pack(side=tk.LEFT)
        self.year_selector = ttk.Combobox(header, width=6, state="readonly", values=list(range(2000, 2101)))
        self.year_selector.pack(side=tk.LEFT, padx=5)
        self.year_selector.set(str(year))
        self.year_selector.bind("<<ComboboxSelected>>", self._on_year_change)

        # Placeholder for client general info
        info_frame = ttk.Frame(content_frame, relief=tk.FLAT)
        info_frame.grid(row=1, column=0, sticky="we", pady=10)
        info_label = ttk.Label(
            info_frame,
            text=(
                "Diverse firmaspecifikk generell info\n"
                "som Bransjekode, Bransjenavn, Selskapsform,\n"
                "Kontaktperson, Adresse mv."
            ),
            justify=tk.CENTER,
        )
        info_label.pack(fill=tk.X)

        # Placeholder for team table
        team_frame = ttk.Frame(content_frame, relief=tk.FLAT)
        team_frame.grid(row=2, column=0, sticky="nsew")
        ttk.Label(team_frame, text="Team", font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 2))
        ttk.Label(team_frame, text="Initialer", borderwidth=1, relief=tk.SOLID, width=12).grid(row=1, column=0, sticky="we")
        ttk.Label(team_frame, text="Rolle", borderwidth=1, relief=tk.SOLID, width=12).grid(row=1, column=1, sticky="we")
        # Example team data; in a real implementation, populate from client metadata
        example_team = [
            ("SB", "Partner"),
            ("AMN", "Manager"),
            ("KS", "Medarbeider"),
            ("AH", "Medarbeider"),
        ]
        for i, (initialer, rolle) in enumerate(example_team, start=2):
            ttk.Label(team_frame, text=initialer, borderwidth=1, relief=tk.SOLID, width=12).grid(row=i, column=0, sticky="we")
            ttk.Label(team_frame, text=rolle, borderwidth=1, relief=tk.SOLID, width=12).grid(row=i, column=1, sticky="we")

        # Dictionary to hold content frames for each section; only one shown at a time
        self._views: Dict[str, tk.Frame] = {}
        self._content_frame = content_frame

        # Create Kildefiler view lazily on first use
        self._views["Kildefiler"] = None  # type: ignore

    def _on_year_change(self, event: object) -> None:
        """Handle change of the revision year."""
        try:
            self.year_var.set(int(self.year_selector.get()))
        except Exception:
            pass

    def _on_nav(self, label: str) -> None:
        """Handle navigation button clicks.

        For most sections this method simply hides/shows the corresponding
        content frame.  For the Kildefiler section it creates the view
        on demand by importing `KildefilerView`.
        """
        # Hide current view
        for view in self._views.values():
            if isinstance(view, tk.Frame) and view.winfo_ismapped():
                view.grid_forget()

        # Create view on first use
        if self._views.get(label) is None:
            if label == "Kildefiler":
                try:
                    from kildefiler_view import KildefilerView  # type: ignore
                except Exception:
                    # If import fails, show placeholder
                    fv = ttk.Frame(self._content_frame)
                    ttk.Label(fv, text="Kildefiler-modulen ikke tilgjengelig.").pack(pady=20)
                else:
                    fv = KildefilerView(
                        self._content_frame,
                        client_name=self.client_name,
                        year=self.year_var.get(),
                    )
                fv.grid(row=3, column=0, sticky="nsew")
                self._views[label] = fv
            else:
                # Placeholder for other sections
                fv = ttk.Frame(self._content_frame)
                ttk.Label(fv, text=f"{label}-funksjonen er ikke implementert ennå.").pack(pady=20)
                fv.grid(row=3, column=0, sticky="nsew")
                self._views[label] = fv
        else:
            # Already created
            fv = self._views[label]
            if fv is not None:
                fv.grid(row=3, column=0, sticky="nsew")


if __name__ == "__main__":
    # Simple test harness for interactive development
    import sys
    root = tk.Tk()
    root.withdraw()  # hide root window
    co = ClientOverview(root, client_name="1234 Eksempel AS", year=2024)
    co.mainloop()
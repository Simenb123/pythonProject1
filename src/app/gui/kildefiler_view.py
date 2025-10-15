"""
Kildefiler view module.

This module defines a `KildefilerView` class that contains the
source file (kildefil) management UI for a client.  It displays
separate version panels for hovedbok (general ledger) and saldobalanse
(balance sheet) and provides buttons to perform analysis, bilag
extraction, mapping and ownership chart actions.

The class is designed to be embedded within the `ClientOverview`
window or similar container.  It expects the caller to provide the
client name and year.  In a fully integrated application it would
also receive `root_dir`, `meta` and other context information; for
this standalone module those dependencies are optional and any
unsupported actions will display a warning message.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional

try:
    # VersionsPanel is defined in app.gui.widgets and used to manage versions
    from app.gui.widgets.versions_panel import VersionsPanel  # type: ignore
    from app.services.clients import (
        load_meta,
        save_meta,
        list_years,
        open_or_create_year,
        set_default_year,
    )
    from app.services.versioning import resolve_active_raw_file
    from app.services.io import read_raw
    from app.services.mapping import ensure_mapping_interactive
except Exception:
    # In standalone mode (e.g., when developing outside the full application)
    # these imports may fail.  We set VersionsPanel to None to disable
    # functionality gracefully.
    VersionsPanel = None  # type: ignore


class KildefilerView(ttk.Frame):
    """UI container for source file version management."""

    def __init__(
        self,
        master: tk.Widget,
        client_name: str,
        year: int,
        root_dir: Optional[str] = None,
        meta: Optional[dict] = None,
    ) -> None:
        super().__init__(master)
        self.client_name = client_name
        self.year_var = tk.IntVar(value=year)
        self.root_dir = root_dir
        self.meta = meta or {}

        # If VersionsPanel is unavailable, show a fallback message
        if VersionsPanel is None:
            ttk.Label(
                self,
                text="VersionsPanel-modulen er ikke tilgjengelig; kan ikke vise kildefiler.",
            ).pack(pady=20)
            return

        # Header: revision year selector
        hdr = ttk.Frame(self)
        hdr.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(hdr, text="Revisjonsår:").pack(side=tk.LEFT)
        self.year_selector = ttk.Combobox(hdr, width=6, state="readonly", values=list(range(2000, 2101)))
        self.year_selector.pack(side=tk.LEFT, padx=5)
        self.year_selector.set(str(year))
        self.year_selector.bind("<<ComboboxSelected>>", self._on_year_change)

        # Data source radio buttons (hovedbok vs. saldobalanse)
        ds_frame = ttk.Frame(self)
        ds_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(ds_frame, text="Datakilde:").grid(row=0, column=0, sticky="w")
        self.source_var = tk.StringVar(value="hovedbok")
        ttk.Radiobutton(ds_frame, text="Hovedbok", variable=self.source_var, value="hovedbok").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(ds_frame, text="Saldobalanse", variable=self.source_var, value="saldobalanse").grid(row=0, column=2, sticky="w")

        # Version panels
        # VersionsPanel requires master with attributes root_dir, client, meta
        # To satisfy that, we create a dummy container and assign attributes.
        # Create a dummy container to satisfy VersionsPanel's expectation of certain attributes.
        holder = tk.Frame(self)
        # Attach context attributes used by VersionsPanel: root_dir, client, meta and year
        holder.root_dir = self.root_dir  # type: ignore[attr-defined]
        holder.client = self.client_name  # type: ignore[attr-defined]
        holder.meta = self.meta  # type: ignore[attr-defined]
        # VersionsPanel checks for a `year` attribute (an IntVar).  Expose year_var via `year`.
        holder.year = self.year_var  # type: ignore[attr-defined]

        self.hb_panel = VersionsPanel(holder, "hovedbok")
        self.hb_panel.pack(fill=tk.X, pady=(0, 10))
        self.sb_panel = VersionsPanel(holder, "saldobalanse")
        self.sb_panel.pack(fill=tk.X, pady=(0, 10))

        # Action buttons
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Analyse", command=lambda: self._start_bilag("analyse")).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Bilagsuttrekk", command=lambda: self._start_bilag("uttrekk")).grid(row=0, column=1, padx=4)
        ttk.Button(btns, text="Mapping …", command=self._ensure_mapping_now).grid(row=0, column=2, padx=4)
        ttk.Button(btns, text="Eierskap", command=self._open_orgchart).grid(row=0, column=3, padx=4)

        # Initial refresh of panels
        self._refresh_panels()

    # --- Internal helpers --------------------------------------------------

    def _on_year_change(self, event: object) -> None:
        try:
            self.year_var.set(int(self.year_selector.get()))
        except Exception:
            return
        self._refresh_panels()

    def _refresh_panels(self) -> None:
        """Refresh version panels based on current year and meta."""
        if not hasattr(self.hb_panel, "refresh"):
            return
        try:
            self.hb_panel.refresh()
            self.sb_panel.refresh()
        except Exception:
            pass

    def _ensure_mapping_now(self) -> None:
        """Invoke interactive mapping for the active version.

        This requires that resolve_active_raw_file, read_raw and
        ensure_mapping_interactive are available and that root_dir/meta
        are provided.  If these dependencies are missing, a warning is
        shown.
        """
        if not (self.root_dir and resolve_active_raw_file and read_raw and ensure_mapping_interactive):
            messagebox.showwarning(
                "Ikke tilgjengelig",
                "Mapping-funksjonen er ikke tilgjengelig i denne konteksten.",
            )
            return
        src = self.source_var.get()
        year = self.year_var.get()
        # Attempt to find active version file
        p = resolve_active_raw_file(
            self.root_dir, self.client_name, year, src, "interim", self.meta
        ) or resolve_active_raw_file(
            self.root_dir, self.client_name, year, src, "ao", self.meta
        )
        if not p:
            messagebox.showwarning("Mangler versjon", "Ingen aktiv versjon for valgt kilde.")
            return
        df, _ = read_raw(p)
        try:
            ensure_mapping_interactive(self, self.root_dir, self.client_name, year, src, df.head(200))
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte mapping: {exc}")

    def _start_bilag(self, modus: str) -> None:
        """Start bilag GUI in a subprocess.

        This requires access to run bilag_gui_tk.  If unavailable,
        display a warning.
        """
        import subprocess, sys
        if not self.root_dir:
            messagebox.showwarning(
                "Ikke tilgjengelig", "Bilagsfunksjonen er ikke tilgjengelig i denne konteksten."
            )
            return
        src = self.source_var.get()
        year = self.year_var.get()
        # Build command; always pass type=interim for now
        args = [
            sys.executable,
            "-m",
            "app.gui.bilag_gui_tk",
            f"--client={self.client_name}",
            f"--year={year}",
            f"--source={src}",
            "--type=interim",
            f"--modus={modus}",
        ]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte Bilag-GUI: {exc}")

    def _open_orgchart(self) -> None:
        """Open the interactive ownership chart.

        Delegates to the existing implementation if available.  If not,
        displays a placeholder message.
        """
        try:
            from app.gui.client_hub import ClientHub  # type: ignore
        except Exception:
            messagebox.showinfo(
                "Eierskap", "Orgkart-funksjonen er ikke tilgjengelig i denne konteksten."
            )
            return
        # Use the existing ClientHub implementation to open org chart
        # by instantiating a temporary ClientHub and invoking _open_orgchart.
        # Note: this is a workaround; ideally the org chart function
        # would be exposed separately.
        tmp = ClientHub(self.master, self.client_name)  # type: ignore
        try:
            tmp._open_orgchart()  # type: ignore[attr-defined]
        finally:
            tmp.destroy()


if __name__ == "__main__":
    # Basic test for interactive development
    import sys
    root = tk.Tk()
    root.withdraw()
    # Passing None for root_dir/meta will disable certain actions
    view = KildefilerView(root, client_name="1234 Eksempel AS", year=2024)
    view.pack(fill="both", expand=True)
    root.mainloop()
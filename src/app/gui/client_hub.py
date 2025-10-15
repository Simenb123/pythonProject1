from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import subprocess
import sys
import os
from pathlib import Path

# Import services from the application.  These must match the definitions
# in the original application for the client hub to function correctly.
from app.services.clients import (
    get_clients_root,
    load_meta,
    save_meta,
    list_years,
    open_or_create_year,
    default_year,
    set_default_year,
    set_year_datakilde,
)
from app.gui.widgets.versions_panel import VersionsPanel  # type: ignore
from app.services.io import read_raw  # type: ignore
from app.services.versioning import resolve_active_raw_file  # type: ignore
from app.services.mapping import ensure_mapping_interactive  # type: ignore

# Path to the bilag GUI script.  In a packaged application this could be
# adjusted to point to the correct module entry point.
BILAG_GUI = Path(__file__).with_name("bilag_gui_tk.py")  # du kan endre senere


class ClientHub(tk.Toplevel):
    """Legacy client hub window.

    This class is retained to provide backward compatibility with existing
    functionality, such as opening the org chart from the Kildefiler view.
    It mirrors the implementation from the original GUI, including
    version management panels and buttons for analysis, bilag extraction,
    mapping and ownership charts.

    Parameters
    ----------
    master : tk.Tk
        The parent window.  Should expose attributes `clients_root` and
        optionally `kildefiler_dir` used to locate files.
    client_name : str
        The name of the client for which the hub is displayed.
    """

    def __init__(self, master: tk.Tk, client_name: str) -> None:
        super().__init__(master)
        self.title(f"Klienthub – {client_name}")
        self.resizable(False, False)

        self.client = client_name
        # root directory for this client – used by VersionsPanel and other services
        self.root_dir = getattr(master, "clients_root")
        # inherit kildefiler_dir from master if present (used to locate central client list)
        self.kildefiler_dir = getattr(master, "kildefiler_dir", None)
        # load per‑client metadata
        self.meta = load_meta(self.root_dir, self.client)

        frm = ttk.Frame(self, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text=f"Klient: {self.client}", font=("", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")

        # ÅR
        years = list_years(self.root_dir, self.client)
        start_year = default_year(self.meta, years[-1] if years else 2025)
        ttk.Label(frm, text="Revisjonsår:").grid(row=1, column=0, sticky="w", pady=(8, 2))
        self.year = tk.IntVar(value=start_year)
        self.year_cmb = ttk.Combobox(frm, values=years, textvariable=self.year, width=10, state="readonly")
        self.year_cmb.grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Åpne år …", command=self._open_year).grid(row=1, column=2, padx=6, sticky="w")
        self.year_cmb.bind("<<ComboboxSelected>>", lambda *_: self._on_year_change())

        # Datakilde
        ttk.Label(frm, text="Datakilde:").grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.source = tk.StringVar(value="hovedbok")
        ttk.Radiobutton(frm, text="Hovedbok", variable=self.source, value="hovedbok").grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(frm, text="Saldobalanse", variable=self.source, value="saldobalanse").grid(row=2, column=2, sticky="w")

        # Versjonspaneler
        # The VersionsPanel needs access to attributes like root_dir, so we use ``self`` as master
        # and place the widgets into the frame via the ``in_`` parameter.  This avoids
        # AttributeError on Frame (no root_dir).
        self.hb_panel = VersionsPanel(self, "hovedbok")
        self.hb_panel.grid(row=3, column=0, columnspan=4, sticky="we", in_=frm)
        self.sb_panel = VersionsPanel(self, "saldobalanse")
        self.sb_panel.grid(row=4, column=0, columnspan=4, sticky="we", in_=frm)

        # Knapper
        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=4, sticky="we", pady=(10, 0))
        ttk.Button(btns, text="Analyse",        command=lambda: self._start_bilag("analyse")).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Bilagsuttrekk",  command=lambda: self._start_bilag("uttrekk")).grid(row=0, column=1, padx=4)
        ttk.Button(btns, text="Mapping …",      command=self._ensure_mapping_now).grid(row=0, column=2, padx=4)
        # New: button to open interactive org chart for this client
        ttk.Button(btns, text="Eierskap", command=self._open_orgchart).grid(row=0, column=3, padx=4)
        ttk.Button(btns, text="Åpne klientmappe", command=self._open_folder).grid(row=0, column=4, padx=4)

        self._on_year_change()

    # -------- helpers ----------
    def _open_year(self) -> None:
        """Open a new year for this client."""
        y = tk.simpledialog.askinteger("Åpne nytt år", "Hvilket år vil du åpne?", parent=self)
        if not y:
            return
        open_or_create_year(self.root_dir, self.client, y)
        self.year_cmb["values"] = list_years(self.root_dir, self.client)
        self.year.set(y)
        self._on_year_change()

    def _on_year_change(self) -> None:
        """Handle selection change in the year combo box."""
        y = self.year.get()
        # update default year in meta
        set_default_year(self.meta, y)
        save_meta(self.root_dir, self.client, self.meta)
        # refresh version panels
        try:
            self.hb_panel.refresh()
            self.sb_panel.refresh()
        except Exception:
            pass

    def _ensure_mapping_now(self) -> None:
        """Ensure mapping is ready by invoking interactive mapping for the active version."""
        # Determine active version file; check interim first, then ÅO
        y = self.year.get()
        src = self.source.get()
        p = resolve_active_raw_file(
            self.root_dir, self.client, y, src, "interim", self.meta
        ) or resolve_active_raw_file(
            self.root_dir, self.client, y, src, "ao", self.meta
        )
        if not p:
            messagebox.showwarning("Ingen aktiv versjon", "Aktiver en versjon først.", parent=self)
            return
        df, _ = read_raw(p)
        try:
            ensure_mapping_interactive(self, self.root_dir, self.client, y, src, df.head(200))
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte mapping: {exc}", parent=self)

    def _start_bilag(self, modus: str) -> None:
        """Start bilag GUI in a subprocess for the selected mode."""
        y = self.year.get()
        src = self.source.get()
        args = [
            sys.executable,
            str(BILAG_GUI),
            f"--client={self.client}",
            f"--year={y}",
            f"--source={src}",
            "--type=interim",
            f"--modus={modus}",
        ]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte Bilag-GUI: {exc}", parent=self)

    def _open_orgchart(self) -> None:
        """Open the interactive ownership chart for the current client."""
        # Look up org nr from meta or prompt the user
        orgnr = self.meta.get("KLIENT_ORGNR", "")
        if not orgnr:
            orgnr = simpledialog.askstring("Org.nr", "Skriv inn organisasjonsnummer:", parent=self)
            if not orgnr:
                return
            # Save to meta so it will be remembered
            self.meta["KLIENT_ORGNR"] = orgnr
            save_meta(self.root_dir, self.client, self.meta)
        # Build command to run the org chart script
        script = Path(__file__).with_name("run_orgchart.py")
        layout = Path(__file__).with_name(f"{self.client}_{self.year.get()}_org_layout.json")
        args = [
            sys.executable,
            str(script),
            f"--orgnr={orgnr}",
            "--editable",
            f"--layout={layout}",
        ]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte orgkart: {exc}", parent=self)

    def _open_folder(self) -> None:
        """Open the client folder in the system's file explorer."""
        path = self.root_dir / self.client
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke åpne mappen: {exc}", parent=self)
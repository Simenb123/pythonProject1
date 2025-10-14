from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import subprocess, sys
import os
from pathlib import Path

from app.services.clients import (
    get_clients_root, load_meta, save_meta, list_years,
    open_or_create_year, default_year, set_default_year, set_year_datakilde
)
from app.gui.widgets.versions_panel import VersionsPanel
from app.services.io import read_raw
from app.services.versioning import resolve_active_raw_file
from app.services.mapping import ensure_mapping_interactive

BILAG_GUI = Path(__file__).with_name("bilag_gui_tk.py")  # du kan endre senere

class ClientHub(tk.Toplevel):
    def __init__(self, master: tk.Tk, client_name: str):
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

        frm = ttk.Frame(self, padding=10); frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text=f"Klient: {self.client}", font=("", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")

        # ÅR
        years = list_years(self.root_dir, self.client)
        start_year = default_year(self.meta, years[-1] if years else 2025)
        ttk.Label(frm, text="Revisjonsår:").grid(row=1, column=0, sticky="w", pady=(8,2))
        self.year = tk.IntVar(value=start_year)
        self.year_cmb = ttk.Combobox(frm, values=years, textvariable=self.year, width=10, state="readonly")
        self.year_cmb.grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Åpne år …", command=self._open_year).grid(row=1, column=2, padx=6, sticky="w")
        self.year_cmb.bind("<<ComboboxSelected>>", lambda *_: self._on_year_change())

        # Datakilde
        ttk.Label(frm, text="Datakilde:").grid(row=2, column=0, sticky="w", pady=(8,2))
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
        btns = ttk.Frame(frm); btns.grid(row=5, column=0, columnspan=4, sticky="we", pady=(10,0))
        ttk.Button(btns, text="Analyse",        command=lambda: self._start_bilag("analyse")).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Bilagsuttrekk",  command=lambda: self._start_bilag("uttrekk")).grid(row=0, column=1, padx=4)
        ttk.Button(btns, text="Mapping …",      command=self._ensure_mapping_now).grid(row=0, column=2, padx=4)
        # New: button to open interactive org chart for this client
        ttk.Button(btns, text="Eierskap", command=self._open_orgchart).grid(row=0, column=3, padx=4)
        ttk.Button(btns, text="Åpne klientmappe", command=self._open_folder).grid(row=0, column=4, padx=4)

        self._on_year_change()

    # -------- helpers ----------
    def _open_year(self):
        y = tk.simpledialog.askinteger("Åpne nytt år", "Hvilket år vil du åpne/opprette?",
                                       parent=self, minvalue=2000, maxvalue=2100, initialvalue=self.year.get())
        if not y: return
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        years = list_years(self.root_dir, self.client)
        self.year_cmb["values"] = years
        self.year.set(y)
        self._on_year_change()
        messagebox.showinfo("År klart", f"Året {y} er opprettet.")

    def _on_year_change(self):
        y = int(self.year.get())
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        self.hb_panel.refresh(); self.sb_panel.refresh()

    def _open_folder(self):
        path = self.root_dir / self.client
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke åpne mappen: {exc}", parent=self)

    # -------- mapping nå ----------
    def _ensure_mapping_now(self):
        y = int(self.year.get())
        src = self.source.get()
        # hent aktiv versjonsfil
        from app.services.versioning import resolve_active_raw_file
        p = resolve_active_raw_file(self.root_dir, self.client, y, src, "interim", self.meta) \
            or resolve_active_raw_file(self.root_dir, self.client, y, src, "ao", self.meta)
        if not p:
            messagebox.showwarning("Mangler versjon", "Ingen aktiv versjon for valgt kilde.", parent=self); return
        df, _ = read_raw(p)
        ensure_mapping_interactive(self, self.root_dir, self.client, y, src, df.head(200))
        messagebox.showinfo("OK", "Mapping er lagret for dette året/kilden.", parent=self)

    # -------- start bilag-GUI ----
    def _start_bilag(self, modus: str):
        y = int(self.year.get())
        src = self.source.get()
        set_year_datakilde(self.meta, y, src); save_meta(self.root_dir, self.client, self.meta)
        # Løft mapping hvis mulig (minimer “første gang”-friksjon)
        from app.services.versioning import resolve_active_raw_file
        p = resolve_active_raw_file(self.root_dir, self.client, y, src, "interim", self.meta) \
            or resolve_active_raw_file(self.root_dir, self.client, y, src, "ao", self.meta)
        if p:
            df, _ = read_raw(p)
            try:
                ensure_mapping_interactive(self, self.root_dir, self.client, y, src, df.head(200))
            except Exception:
                pass

        args = [sys.executable, str(BILAG_GUI),
                f"--client={self.client}", f"--year={y}",
                f"--source={src}", f"--type=interim", f"--modus={modus}"]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte Bilag-GUI: {exc}", parent=self)

    # -------- Eierskap (orgkart) ----------
    def _open_orgchart(self) -> None:
        """
        Launch the interactive ownership chart for the current client.

        If the client's organisasjonsnummer (``KLIENT_ORGNR``) is not
        already stored in the metadata, the user is prompted to enter
        it.  The value is persisted to the client's ``meta.json``
        file.  A layout file named ``orgchart_layout.json`` in the
        client's root directory is used to load/save node positions
        when editing.  The org chart is launched in a separate
        process via ``run_orgchart.py`` with appropriate command line
        arguments.
        """
        """
        Launch the interactive ownership chart for the current client.

        The organisation number is looked up automatically from metadata
        (``KLIENT_ORGNR``) or, if missing, from the central client list.
        Only if it still can't be found will the user be prompted to
        provide it.  The value is persisted to ``meta.json`` when
        entered.  The org chart is launched using ``-m app.gui.run_orgchart``
        rather than invoking the script directly; this ensures that
        relative imports in ``org_controller.py`` resolve correctly.
        """
        # Ensure metadata exists
        if not hasattr(self, 'meta') or self.meta is None:
            self.meta = {}
        orgnr = self.meta.get("KLIENT_ORGNR")
        if not orgnr:
            # Attempt to find orgnr automatically from client list
            orgnr = self._find_client_orgnr()
        if not orgnr:
            # As a last resort, ask the user
            orgnr = simpledialog.askstring(
                "Organisasjonsnummer",
                "Skriv organisasjonsnummer for klienten",
                parent=self,
            )
            if not orgnr:
                return  # user cancelled
        # Save orgnr to meta if new
        if orgnr and self.meta.get("KLIENT_ORGNR") != orgnr:
            self.meta["KLIENT_ORGNR"] = orgnr
            save_meta(self.root_dir, self.client, self.meta)
        # Determine layout path under client directory
        layout_path = Path(self.root_dir) / self.client / "orgchart_layout.json"
        try:
            # Construct the path to run_orgchart.py relative to this file
            script_path = Path(__file__).with_name("run_orgchart.py")
            # Invoke the script directly.  run_orgchart.py adjusts sys.path to
            # allow importing org_controller and db when run as a script.
            args = [
                sys.executable,
                str(script_path),
                f"--orgnr={orgnr}",
                "--editable",
                f"--layout={layout_path}",
            ]
            subprocess.Popen(args, shell=False)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke starte eierskapskart: {exc}", parent=self)

    def _find_client_orgnr(self) -> str | None:
        """
        Attempt to find the client's organisasjonsnummer from the central
        Excel client list.  Returns None if not found or an error occurs.
        """
        """
        Attempt to find the client's organisation number from the central Excel client list.
        The search uses the kildefiler directory inherited from the master (if available),
        otherwise falls back to the default ``find_kildefiler_dir``.  Returns None if
        not found or if an error occurs.
        """
        try:
            import pandas as pd  # type: ignore
        except Exception:
            return None
        # Determine the client name portion (after number) and normalise
        parts = str(self.client).split(" ", 1)
        client_name = parts[1].strip().lower() if len(parts) > 1 else ""
        if not client_name:
            return None
        # Determine base directory for "Kildefiler".  Use master-provided override
        base_dir: Path | None = None
        if getattr(self, "kildefiler_dir", None):
            base_dir = self.kildefiler_dir
        if not base_dir:
            # Fall back to the helper from regnskapslinjer
            try:
                from app.services.regnskapslinjer import find_kildefiler_dir  # type: ignore
                found = find_kildefiler_dir()
                if found:
                    base_dir = Path(found)
            except Exception:
                base_dir = None
        if not base_dir:
            return None
        # List of possible filenames for the client list
        for fn in ["BHL AS klientliste - kopi.xlsx", "BHL AS klientliste.xlsx", "BHLAS klientliste.xlsx"]:
            fpath = base_dir / fn
            if not fpath.exists():
                continue
            try:
                df = pd.read_excel(fpath)  # type: ignore
            except Exception:
                continue
            # Build a mapping of lowercase column names to actual names
            lower = {str(c).strip().lower(): c for c in df.columns}
            org_col = next((lower[c] for c in lower if c in {"klient_orgnr", "orgnr", "organisasjonsnummer"}), None)
            name_col = next((lower[c] for c in lower if c in {"klient_navn", "navn", "client_navn"}), None)
            if org_col and name_col:
                sub = df[[name_col, org_col]].dropna()
                for _, row in sub.iterrows():
                    nm = str(row[name_col]).strip().lower()
                    if nm == client_name:
                        val = str(row[org_col]).strip()
                        import re
                        digits = re.sub(r"[^0-9]", "", val)
                        return digits or val
        return None

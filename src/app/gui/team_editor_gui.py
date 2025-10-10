"""
Team editor GUI integrated with central team file.

This module defines a simple Tkinter-based TeamEditor window that reads and
updates team assignments from a central Excel file (e.g. ``BHL AS Team.xlsx``).

The original application stored team membership in JSON files within each
client's directory, which made it cumbersome to maintain and required
accessing hundreds of small files across a network.  By consolidating
team data into a single spreadsheet and loading it once, we avoid many
network calls and make it easier to manage team memberships centrally.

Usage:
    from app.gui.team_editor_gui import TeamEditor
    TeamEditor(parent, clients_root, client_name)

``client_name`` should be the folder name used for the client, typically
formatted as "<client_nr> <client_name>".  The first whitespace-separated
token is interpreted as the client number.  Team assignments are stored
per client number in the Excel file.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional

import pandas as pd


class TeamEditor(tk.Toplevel):
    """A simple editor for managing team assignments for a given client.

    This editor displays the current team members for the selected client and
    allows the user to add or remove members.  Changes are persisted back
    into the central team Excel file when the user clicks the Save button.
    """

    def __init__(self, parent: tk.Misc, clients_root: Path, client_name: str) -> None:
        super().__init__(parent)
        self.title(f"Team – {client_name}")
        self.transient(parent)
        self.grab_set()
        self.resizable(width=True, height=True)

        # Determine client number from directory name (e.g. "1234 Client AS" → 1234)
        try:
            parts = client_name.strip().split()
            self.client_nr = int(parts[0])
        except Exception:
            # Fall back to zero if parsing fails; this will show an empty list
            self.client_nr = 0

        # Locate kildefiler directory using the same logic as start_portal
        base_dir: Optional[Path] = None
        try:
            from app.services.regnskapslinjer import find_kildefiler_dir  # type: ignore
            bd = find_kildefiler_dir()
            if bd:
                base_dir = Path(bd)
        except Exception:
            pass
        if base_dir is None:
            # Fall back to a Kildefiler directory two levels up from this file
            possible = Path(__file__).resolve().parents[3] / "Kildefiler"
            if possible.exists():
                base_dir = possible

        # Paths to central team and employee lists
        self.team_path: Optional[Path] = None
        self.emp_path: Optional[Path] = None
        if base_dir:
            tp = base_dir / "BHL AS Team.xlsx"
            ep = base_dir / "Ansatte BHL.xlsx"
            if tp.exists():
                self.team_path = tp
            if ep.exists():
                self.emp_path = ep

        # Load data
        self.team_df: pd.DataFrame = pd.DataFrame()
        self.emp_df: pd.DataFrame = pd.DataFrame()
        if self.team_path is not None:
            try:
                self.team_df = pd.read_excel(self.team_path)
            except Exception as exc:
                messagebox.showwarning("Team", f"Kunne ikke lese teamfilen:\n{exc}")
        if self.emp_path is not None:
            try:
                self.emp_df = pd.read_excel(self.emp_path)
            except Exception:
                pass

        # Build initial list of team members for this client
        self.current_members: list[dict[str, str]] = []
        self._load_current_members()

        # UI layout
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=8, pady=8)

        # Team table (treeview)
        columns = ("initial", "name", "role")
        self.tree = ttk.Treeview(frm, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("initial", text="Initialer")
        self.tree.heading("name", text="Navn")
        self.tree.heading("role", text="Rolle")
        self.tree.column("initial", width=80)
        self.tree.column("name", width=200)
        self.tree.column("role", width=120)
        self.tree.pack(fill="both", expand=True)
        self._refresh_tree()

        # Buttons
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(fill="x", pady=(8, 0))
        ttk.Button(btn_frame, text="Legg til …", command=self._add_member).pack(side="left")
        ttk.Button(btn_frame, text="Fjern", command=self._remove_selected).pack(side="left", padx=(8, 0))
        ttk.Button(btn_frame, text="Lagre", command=self._save_and_close).pack(side="right")

    def _load_current_members(self) -> None:
        """Populate self.current_members from the team DataFrame."""
        self.current_members.clear()
        if self.team_df.empty:
            return
        for _, row in self.team_df.iterrows():
            try:
                if int(row.get("KLIENT_NR", -1)) != self.client_nr:
                    continue
            except Exception:
                continue
            ini = str(row.get("INITIAL") or "").strip().upper()
            role = str(row.get("Rolle") or "").strip()
            # Look up the full name in employee list
            name = ""
            if not self.emp_df.empty:
                match = self.emp_df[self.emp_df.get("IN", "").astype(str).str.strip().str.upper() == ini]
                if not match.empty:
                    name = str(match.iloc[0].get("Fullname") or match.iloc[0].get("Navn") or "").strip()
            self.current_members.append({"initial": ini, "name": name, "role": role})

    def _refresh_tree(self) -> None:
        """Refresh the displayed tree with current members."""
        for i in self.tree.get_children():
            self.tree.delete(i)
        for m in self.current_members:
            self.tree.insert("", "end", values=(m["initial"], m["name"], m["role"]))

    def _add_member(self) -> None:
        """Add a new team member.  Opens a dialog to select an employee and set their role."""
        # Build list of potential employees not already in team
        if self.emp_df.empty:
            messagebox.showinfo("Legg til", "Fant ingen ansattliste. Du kan ikke legge til ansatte.")
            return
        current_initials = {m["initial"] for m in self.current_members}
        candidates = self.emp_df[self.emp_df.get("IN", "").astype(str).str.strip().str.upper().apply(lambda x: x not in current_initials)]
        if candidates.empty:
            messagebox.showinfo("Legg til", "Alle ansatte er allerede i teamet.")
            return
        # Create selection window
        sel = tk.Toplevel(self)
        sel.title("Legg til teammedlem")
        sel.transient(self)
        sel.grab_set()
        sel.minsize(360, 420)
        qvar = tk.StringVar(value="")
        ttk.Label(sel, text="Søk:").pack(anchor="w", padx=8, pady=(8,0))
        ent = ttk.Entry(sel, textvariable=qvar)
        ent.pack(fill="x", padx=8)
        listbox = tk.Listbox(sel, height=16)
        listbox.pack(fill="both", expand=True, padx=8, pady=8)

        def refill(*_):
            q = qvar.get().strip().lower()
            listbox.delete(0, tk.END)
            rows = candidates
            if q:
                rows = rows[rows.apply(lambda r: q in str(r.get("Fullname", "")).lower()
                                       or q in str(r.get("Navn", "")).lower()
                                       or q in str(r.get("IN", "")).lower()
                                       or q in str(r.get("Email", "")).lower()
                                       or q in str(r.get("epost", "")).lower(), axis=1)]
            for _, r in rows.iterrows():
                display_name = str(r.get("Fullname") or r.get("Navn") or "").strip()
                ini = str(r.get("IN", "")).strip().upper()
                listbox.insert(tk.END, f"{ini} – {display_name}")

        qvar.trace_add("write", lambda *_: refill())
        refill()

        # Role selection
        ttk.Label(sel, text="Rolle:").pack(anchor="w", padx=8)
        role_var = tk.StringVar(value="Medarbeider")
        role_cb = ttk.Combobox(sel, textvariable=role_var, state="readonly",
                               values=["Partner", "Manager", "Medarbeider"])
        role_cb.pack(fill="x", padx=8)

        # OK and cancel buttons
        def on_ok() -> None:
            idx = listbox.curselection()
            if not idx:
                messagebox.showwarning("Legg til", "Marker en ansatt i listen.", parent=sel)
                return
            val = listbox.get(idx[0])
            ini = val.split("–", 1)[0].strip().upper()
            role = role_var.get().strip()
            # Look up full name
            match = self.emp_df[self.emp_df.get("IN", "").astype(str).str.strip().str.upper() == ini]
            name = ""
            if not match.empty:
                name = str(match.iloc[0].get("Fullname") or match.iloc[0].get("Navn") or "").strip()
            self.current_members.append({"initial": ini, "name": name, "role": role})
            self._refresh_tree()
            sel.destroy()

        def on_cancel() -> None:
            sel.destroy()

        btn_frame = ttk.Frame(sel)
        btn_frame.pack(fill="x", padx=8, pady=(0,8))
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right")
        ttk.Button(btn_frame, text="Avbryt", command=on_cancel).pack(side="right", padx=(0,8))

    def _remove_selected(self) -> None:
        """Remove the selected rows from the team list."""
        sels = self.tree.selection()
        if not sels:
            return
        # Build set of initials to remove
        initials_to_remove = set()
        for iid in sels:
            vals = self.tree.item(iid, "values")
            if vals:
                initials_to_remove.add(vals[0])
        if not initials_to_remove:
            return
        # Confirm deletion
        if not messagebox.askyesno("Fjern", "Fjern valgte teammedlemmer?", parent=self):
            return
        # Remove from list
        self.current_members = [m for m in self.current_members if m["initial"] not in initials_to_remove]
        self._refresh_tree()

    def _save_and_close(self) -> None:
        """Persist changes to the central team file and close the window."""
        if not self.team_path:
            messagebox.showerror("Team", "Fant ikke plasseringen til teamfilen.")
            self.destroy()
            return
        # Read existing file (if loaded earlier) or load fresh to avoid race
        try:
            df = pd.read_excel(self.team_path) if self.team_df.empty else self.team_df.copy()
        except Exception as exc:
            messagebox.showerror("Team", f"Kunne ikke lese teamfilen:\n{exc}")
            self.destroy()
            return
        # Remove all existing entries for this client_nr
        if not df.empty:
            try:
                df = df[df.get("KLIENT_NR", 0).astype(float).astype(int) != self.client_nr]
            except Exception:
                # fallback: remove by comparing string
                df = df[df.get("KLIENT_NR", "").astype(str) != str(self.client_nr)]
        # Build new rows
        rows = []
        for m in self.current_members:
            ini = m["initial"]
            role = m["role"] or ""
            rows.append({"KLIENT_NR": self.client_nr, "INITIAL": ini, "Rolle": role})
        new_df = pd.DataFrame(rows)
        # Append and write back
        final_df = pd.concat([df, new_df], ignore_index=True)
        try:
            final_df.to_excel(self.team_path, index=False)
        except Exception as exc:
            messagebox.showerror("Team", f"Kunne ikke lagre teamfilen:\n{exc}")
            # Even if save fails, we close the dialog
        self.destroy()
# src/app/gui/start_portal.py
from __future__ import annotations

# --- robust modulkjøring ---
import sys
from pathlib import Path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

from app.gui.klient_launcher import ClientHub                # Klienthub (analyse/mapping/versjoner) :contentReference[oaicite:3]{index=3}
from app.services.clients import (                           # Klient-rot + klientliste, samme som før  :contentReference[oaicite:4]{index=4}
    get_clients_root, set_clients_root, resolve_root_and_client, list_clients,
)
from app.services.registry import (                          # E-post-ID, ansattliste og team
    current_email, set_current_email, list_employees_df, import_employees_from_excel,
    team_has_user
)

import pandas as pd
from pathlib import Path

# (valgfrie) undermoduler – lastes «late» i knapper
#  - app.gui.client_info_gui: ClientInfoWindow
#  - app.gui.team_editor_gui: TeamEditor
#  - app.gui.ar_import_gui   : ARImportWindow


class StartPortal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Start – Klientportal")
        self.minsize(720, 420)

        # sørg for at Klienthub kan lese master.clients_root (brukes i ClientHub.__init__) :contentReference[oaicite:5]{index=5}
        self.clients_root = get_clients_root()
        # ensure we have a valid clients root; if not, prompt the user
        if not self.clients_root or not self.clients_root.exists():
            self._choose_root()
            if not self.clients_root:
                self.after(10, self.destroy)
                return

        # state
        self._email_var = tk.StringVar(value=current_email() or "")
        self.only_mine = tk.BooleanVar(value=False)
        self.q = tk.StringVar(value="")

        # header
        header = ttk.LabelFrame(self, text="Administrasjon")
        header.pack(fill="x", padx=8, pady=(8, 2))

        ttk.Label(header, text="Logget inn:").grid(row=0, column=0, sticky="w", padx=(8, 4), pady=6)
        self.lbl_user = ttk.Label(header, text=self._fmt_user())
        self.lbl_user.grid(row=0, column=1, sticky="w")
        ttk.Button(header, text="Sett e‑post …", command=self._pick_user).grid(row=0, column=2, padx=(8, 8))

        ttk.Checkbutton(header, text="Mine klienter", variable=self.only_mine, command=self._refresh_list)\
            .grid(row=0, column=3, sticky="e", padx=(8, 8))

        # søk
        wrap = ttk.Frame(self); wrap.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(wrap, text="Søk:").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(wrap, textvariable=self.q)
        ent.grid(row=0, column=1, sticky="we")
        ent.bind("<KeyRelease>", lambda *_: self._refresh_list())
        self.count_lbl = ttk.Label(wrap, text="")
        self.count_lbl.grid(row=0, column=2, sticky="e", padx=(8, 0))

        self.lb = tk.Listbox(wrap, height=18, activestyle="dotbox")
        self.lb.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        self.lb.bind("<Return>", self._open_hub)
        self.lb.bind("<Double-1>", self._open_hub)
        sb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=3, sticky="ns")

        wrap.columnconfigure(1, weight=1)
        wrap.rowconfigure(1, weight=1)

        # actions
        btns = ttk.Frame(self); btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="Åpne hub", command=self._open_hub).pack(side="left")
        ttk.Button(btns, text="Klientinfo…", command=self._open_info).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Team…", command=self._open_team).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="AR‑import…", command=self._open_ar_import).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Oppdater", command=self._update_clients).pack(side="right")

        # init list of clients
        self._all_clients: list[str] = []
        # try to load from the central client list (Excel) first; if it fails, fall back to scanning directories
        try:
            self._load_clients_from_excel()
        except Exception:
            try:
                self._all_clients = list_clients(self.clients_root)
            except Exception:
                self._all_clients = []
        # -- load team and employee data for "Mine klienter" filtering --
        # helper structures are initialised in _load_team_data()
        self.team_df: pd.DataFrame
        self.emp_df: pd.DataFrame
        self.client_to_initials: dict[int, set[str]]
        self.email_to_initial: dict[str, str]
        self._load_team_data()
        self._refresh_list()

    # ---------------- UI helpers ----------------
    def _fmt_user(self) -> str:
        e = self._email_var.get().strip()
        if not e: return "(ingen e‑post valgt)"
        alias = e.split("@")[0]
        return f"{e}  ({alias})"

    def _choose_root(self):
        """
        Prompt the user to choose a new clients root.  After the root is set
        the client list is reloaded from the central client list if present;
        otherwise we fall back to scanning the directory.
        """
        p = filedialog.askdirectory(title="Velg klient‑rot eller en klientmappe")
        if not p:
            return
        root, single = resolve_root_and_client(Path(p))
        set_clients_root(root)
        self.clients_root = root
        # reload client list: try excel first, then fallback to scanning directories
        try:
            self._load_clients_from_excel()
        except Exception:
            try:
                self._all_clients = list_clients(self.clients_root)
            except Exception:
                self._all_clients = []
        self._refresh_list()
        if single and single in self._all_clients:
            self.lb.selection_clear(0, tk.END)
            idx = self._all_clients.index(single)
            self.lb.selection_set(idx)
            self.lb.see(idx)

    # ---------------- Actions ----------------
    def _pick_user(self):
        # vis ansattliste hvis vi har en
        df = list_employees_df()
        if df is None or df.empty:
            # ingen ansattliste – la bruker skrive inn
            email = simpledialog.askstring("Sett e‑post", "Skriv e‑postadressen din:", parent=self)
            if not email: return
            set_current_email(email)
            self._email_var.set(current_email())
            self.lbl_user.config(text=self._fmt_user()); self._refresh_list(); return

        top = tk.Toplevel(self); top.title("Velg bruker"); top.transient(self); top.grab_set(); top.minsize(360, 420)
        qv = tk.StringVar(value="")
        ttk.Label(top, text="Søk:").pack(anchor="w", padx=8, pady=(8,0))
        ent = ttk.Entry(top, textvariable=qv); ent.pack(fill="x", padx=8)
        lb = tk.Listbox(top, height=16); lb.pack(fill="both", expand=True, padx=8, pady=8)

        def refill(*_):
            q = qv.get().strip().lower()
            lb.delete(0, tk.END)
            rows = df
            if q:
                rows = rows[rows.apply(lambda r:
                                       q in str(r.get("name","")).lower()
                                       or q in str(r.get("email","")).lower()
                                       or q in str(r.get("initials","")).lower(), axis=1)]
            for _, r in rows.iterrows():
                name = str(r.get("name") or "").strip()
                email = str(r.get("email") or "").strip().lower()
                init = str(r.get("initials") or "").strip()
                label = f"{name} [{email}]" if name else email
                if init: label = f"{label}  ({init})"
                lb.insert(tk.END, label)
        refill()
        ent.bind("<KeyRelease>", refill)

        def ok(*_):
            if not lb.curselection(): return
            label = lb.get(lb.curselection()[0])
            # dra ut [email] fra etiketten
            import re
            m = re.search(r"\[([^\]]+)\]", label)
            email = (m.group(1) if m else label).strip()
            set_current_email(email)
            self._email_var.set(current_email())
            self.lbl_user.config(text=self._fmt_user())
            top.destroy(); self._refresh_list()
        ttk.Button(top, text="OK", command=ok).pack(pady=(0,8))
        lb.bind("<Double-1>", ok)

    def _refresh_list(self):
        q = self.q.get().strip().lower()
        base = [n for n in self._all_clients if (q in n.lower())]
        if self.only_mine.get():
            email = self._email_var.get().strip().lower()
            if not email:
                messagebox.showinfo(
                    "Velg e‑post",
                    "Sett e‑post først for å bruke «Mine klienter».",
                    parent=self,
                )
                self.only_mine.set(False)
            else:
                # Determine the set of initials associated with the logged-in user.
                initials_to_check: set[str] = set()
                # Try to look up initials via email
                ini = self.email_to_initial.get(email)
                if ini:
                    initials_to_check.add(ini)
                else:
                    # Fallback: use alias (prefix before '@') as initial guess
                    alias = email.split("@")[0]
                    alias_upper = alias.replace(".", "").replace("_", "").upper()
                    if alias_upper:
                        initials_to_check.add(alias_upper)
                filtered_base: list[str] = []
                for n in base:
                    # Parse the client number from the beginning of the string
                    cid = None
                    try:
                        first = n.strip().split()[0]
                        # Extract leading digits from the first token (handles 0000_name or 0000-name)
                        digits = ""
                        for ch in first:
                            if ch.isdigit():
                                digits += ch
                            else:
                                break
                        if digits:
                            cid = int(digits)
                        else:
                            cid = int(first)
                    except Exception:
                        cid = None
                    if cid is None:
                        continue
                    member_inis = self.client_to_initials.get(cid)
                    if not member_inis:
                        continue
                    # Check if any of the user's initials match this client
                    match = False
                    for user_ini in initials_to_check:
                        if user_ini.upper() in member_inis:
                            match = True
                            break
                    if match:
                        filtered_base.append(n)
                base = filtered_base
        self.lb.delete(0, tk.END)
        for n in base: self.lb.insert(tk.END, n)
        self.count_lbl.config(text=f"{len(base)} av {len(self._all_clients)}")

    def _load_team_data(self) -> None:
        """Load team and employee information from central files.

        This helper populates self.team_df, self.emp_df, self.client_to_initials,
        and self.email_to_initial for use in filtering "Mine klienter".  It
        attempts to locate the files in the kildefiler directory or a local
        fallback directory.  Missing files are tolerated; in that case no
        filtering by team will occur.
        """
        self.team_df = pd.DataFrame()
        self.emp_df = pd.DataFrame()
        self.client_to_initials = {}
        self.email_to_initial = {}
        # Attempt to locate files using find_kildefiler_dir
        base_dir = None
        try:
            from app.services.regnskapslinjer import find_kildefiler_dir  # type: ignore
            bd = find_kildefiler_dir()
            if bd:
                base_dir = Path(bd)
        except Exception:
            pass
        if base_dir is None:
            # fallback: Kildefiler directory relative to src root
            possible_dir = Path(__file__).resolve().parents[2] / "Kildefiler"
            if possible_dir.exists():
                base_dir = possible_dir
        # read files if found
        if base_dir:
            tfile = base_dir / "BHL AS Team.xlsx"
            if tfile.exists():
                try:
                    self.team_df = pd.read_excel(tfile)
                except Exception:
                    pass
            efile = base_dir / "Ansatte BHL.xlsx"
            if efile.exists():
                try:
                    self.emp_df = pd.read_excel(efile)
                except Exception:
                    pass
        # Build helper maps
        if not self.emp_df.empty:
            # Normaliser e‑postkolonner til små bokstaver hvis de finnes. Vi gjør dette før mapping.
            if "epost" in self.emp_df.columns:
                self.emp_df["epost"] = self.emp_df["epost"].astype(str).str.lower().str.strip()
            if "email" in self.emp_df.columns:
                self.emp_df["email"] = self.emp_df["email"].astype(str).str.lower().str.strip()
            if "Email" in self.emp_df.columns:
                self.emp_df["Email"] = self.emp_df["Email"].astype(str).str.lower().str.strip()
            # Bygg mapping fra e‑postadresse til initialer dersom begge kolonner finnes.
            if "IN" in self.emp_df.columns:
                initials_series = self.emp_df["IN"].astype(str).str.strip().str.upper()
                # Finn første tilgjengelige e‑postkolonne
                emails = None
                for col in ("epost", "email", "Email"):
                    if col in self.emp_df.columns:
                        emails = self.emp_df[col].astype(str).str.lower().str.strip()
                        break
                if emails is not None:
                    self.email_to_initial = dict(zip(emails, initials_series))
        if not self.team_df.empty:
            for _, r in self.team_df.iterrows():
                try:
                    cid = int(r.get("KLIENT_NR"))
                except Exception:
                    continue
                ini = str(r.get("INITIAL", "")).strip().upper()
                if cid not in self.client_to_initials:
                    self.client_to_initials[cid] = set()
                if ini:
                    self.client_to_initials[cid].add(ini)

    def _load_clients_from_excel(self) -> None:
        """
        Populate self._all_clients from a central Excel client list instead of scanning
        the client folders.  This method attempts to locate the Excel file in the
        configured 'Kildefiler' directory.  Expected column names are
        'KLIENT_NR' (client number) and 'KLIENT_NAVN' (client name).  If the file
        cannot be found or read, this function does nothing.
        """
        # reset client list
        self._all_clients = []
        # find kildefiler directory using the same logic as in _load_team_data
        base_dir = None
        try:
            from app.services.regnskapslinjer import find_kildefiler_dir  # type: ignore
            bd = find_kildefiler_dir()
            if bd:
                base_dir = Path(bd)
        except Exception:
            pass
        if base_dir is None:
            # fallback: Kildefiler directory relative to src root
            possible_dir = Path(__file__).resolve().parents[2] / "Kildefiler"
            if possible_dir.exists():
                base_dir = possible_dir
        if base_dir is None:
            return
        # possible filenames for client list
        possible_files = [
            "BHL AS klientliste - kopi.xlsx",
            "BHL AS klientliste.xlsx",
            "BHLAS klientliste.xlsx",
        ]
        client_file = None
        for fn in possible_files:
            fp = base_dir / fn
            if fp.exists():
                client_file = fp
                break
        if client_file is None:
            return
        try:
            df = pd.read_excel(client_file)
        except Exception:
            return
        # Expect columns 'KLIENT_NR' and 'KLIENT_NAVN'
        nr_col, name_col = None, None
        lower_cols = {str(c).strip().lower(): c for c in df.columns}
        for key, orig in lower_cols.items():
            if key in {"klient_nr", "klientnr", "nr", "client_nr", "clientnr"}:
                nr_col = orig
            if key in {"klient_navn", "klientnavn", "navn", "client_navn", "clientnavn"}:
                name_col = orig
        if nr_col is None or name_col is None:
            return
        # build list of "<nr> <name>"
        clients: list[str] = []
        for _, row in df[[nr_col, name_col]].dropna().iterrows():
            try:
                cid = int(float(str(row[nr_col]).strip()))
            except Exception:
                continue
            name = str(row[name_col]).strip()
            if not name:
                continue
            clients.append(f"{cid} {name}")
        # remove duplicates and sort
        clients = sorted(set(clients), key=lambda s: s.lower())
        self._all_clients = clients

    def _update_clients(self) -> None:
        """Reload the client list from Excel and refresh the display."""
        try:
            self._load_clients_from_excel()
        except Exception:
            # ignore if load fails; just refresh existing list
            pass
        self._refresh_list()

    def _selected_client(self) -> str | None:
        sel = self.lb.curselection()
        return None if not sel else self.lb.get(sel[0])

    def _open_hub(self, *_):
        name = self._selected_client()
        if not name:
            messagebox.showwarning("Velg klient", "Velg en klient i listen.", parent=self); return
        # ClientHub leser master.clients_root i __init__  :contentReference[oaicite:6]{index=6}
        ClientHub(self, name)

    def _open_info(self):
        name = self._selected_client()
        if not name:
            messagebox.showwarning("Velg klient", "Velg en klient i listen.", parent=self); return
        try:
            from app.gui.client_info_gui import ClientInfoWindow
            ClientInfoWindow(self, self.clients_root, name)
        except Exception as exc:
            messagebox.showerror("Mangler modul",
                                 f"Kunne ikke laste Klientinfo‑GUI.\n{type(exc).__name__}: {exc}", parent=self)

    def _open_team(self):
        name = self._selected_client()
        if not name:
            messagebox.showwarning("Velg klient", "Velg en klient i listen.", parent=self); return
        try:
            from app.gui.team_editor_gui import TeamEditor
            win = TeamEditor(self, self.clients_root, name)
            # wait for editor to close before refreshing team data
            self.wait_window(win)
            # reload team data and refresh list (in case of changes)
            self._load_team_data()
            self._refresh_list()
        except Exception as exc:
            messagebox.showerror("Mangler modul",
                                 f"Kunne ikke laste Team‑GUI.\n{type(exc).__name__}: {exc}", parent=self)

    def _open_ar_import(self):
        name = self._selected_client()
        if not name:
            messagebox.showwarning("Velg klient", "Velg en klient i listen.", parent=self); return
        try:
            from app.gui.ar_import_gui import ARImportWindow
            ARImportWindow(self, self.clients_root, name)
        except Exception as exc:
            messagebox.showerror("Mangler modul",
                                 f"Kunne ikke laste AR‑import.\n{type(exc).__name__}: {exc}", parent=self)


if __name__ == "__main__":
    StartPortal().mainloop()

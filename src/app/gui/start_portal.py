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
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        # Knappene nederst i vinduet gir tilgang til ulike funksjoner.
        ttk.Button(btns, text="Åpne hub", command=self._open_hub).pack(side="left")
        ttk.Button(btns, text="Klientinfo…", command=self._open_info).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Team…", command=self._open_team).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="AR‑import…", command=self._open_ar_import).pack(side="left", padx=(6, 0))
        # Legg til en egen knapp for å velge/endre klient‑rot. Dette gjør det mulig
        # for ikke‑tekniske brukere å peke applikasjonen til riktig katalog med
        # klientmapper (for eksempel «F:\\Dokument\\2\\BHL klienter\\Klienter»).
        ttk.Button(btns, text="Bytt rot …", command=self._choose_root).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Oppdater", command=self._refresh_list).pack(side="right")

        # init liste
        self._all_clients: list[str] = list_clients(self.clients_root)
        self._refresh_list()

    # ---------------- UI helpers ----------------
    def _fmt_user(self) -> str:
        e = self._email_var.get().strip()
        if not e: return "(ingen e‑post valgt)"
        alias = e.split("@")[0]
        return f"{e}  ({alias})"

    def _choose_root(self):
        p = filedialog.askdirectory(title="Velg klient‑rot eller en klientmappe")
        if not p: return
        root, single = resolve_root_and_client(Path(p))
        set_clients_root(root)
        self.clients_root = root
        self._all_clients = list_clients(self.clients_root)
        self._refresh_list()
        if single and single in self._all_clients:
            self.lb.selection_clear(0, tk.END)
            idx = self._all_clients.index(single)
            self.lb.selection_set(idx); self.lb.see(idx)

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
                messagebox.showinfo("Velg e‑post", "Sett e‑post først for å bruke «Mine klienter».", parent=self)
                self.only_mine.set(False)
            else:
                base = [n for n in base if team_has_user(self.clients_root, n, email)]
        self.lb.delete(0, tk.END)
        for n in base: self.lb.insert(tk.END, n)
        self.count_lbl.config(text=f"{len(base)} av {len(self._all_clients)}")

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
            TeamEditor(self, self.clients_root, name)
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

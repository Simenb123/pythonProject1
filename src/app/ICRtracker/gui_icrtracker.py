# -*- coding: utf-8 -*-
"""
GUI-IRCTracker
--------------
Tkinter-gui for å styre og kjøre tracker, vise resultater, planlegge kjøring og sende varsel på e-post.

Plassering: src/app/ICRtracker/gui_icrtracker.py
Kjør:       Run i PyCharm, eller:  python -m app.ICRtracker.gui_irctracker  (fra mappen <prosjekt>/src)

Avhenger av:
 - app.ICRtracker.tracker        (eksisterende, fra tidligere)
 - app.ICRtracker.ar_bridge      (hvis aktiv; faller tilbake til registry_db)
 - win32com (pywin32), openpyxl  (samme som tracker)

Hva den gjør:
 - Lar deg sette innstillinger (avsender å filtrere, hvor lenge tilbake, stier, osv.)
 - Kjør "nå": Kaller tracker.main() med dine innstillinger og leser csv-resultater inn i tabeller
 - Epostvarsel: Hvis aktivert og det finnes treff, sendes epost via Outlook
 - Planlegging:
     * Intern: appen sjekker klokkeslett og kjører selv til valgt tid (så lenge GUI står åpent)
     * Windows Task Scheduler (valgfritt): opprett/slett planlagt oppgave som kjører en CLI-runner (kan legges til senere)
 - (Valgfritt) Rekursiv AR-søk (opp/ned) til valgt dybde fra orgnr funnet i vedleggene

Merk:
 - For Task Scheduler-knapper brukes 'schtasks' CLI. Det krever at Python og modulstien er kjent.
 - Du kan fint begynne med kun intern planlegging (GUI må stå åpent).
"""

from __future__ import annotations

# --- Robust bootstrap: legg mappen som inneholder "app" på sys.path når du kjører direkte ---
if __name__ == "__main__" and __package__ is None:
    import sys, pathlib

    here = pathlib.Path(__file__).resolve()
    for parent in (here.parent, *here.parents):
        if (parent / "app" / "__init__.py").exists():
            sys.path.insert(0, str(parent))
            __package__ = "app.ICRtracker"
            break

import csv
import json
import subprocess
import threading
import traceback
from dataclasses import dataclass, asdict
from datetime import datetime, time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Importer tracker + AR ---
try:
    from . import tracker  # vi setter konstanter og kjører tracker.main()
except Exception as e:
    raise RuntimeError(f"Kunne ikke importere tracker: {e}")

# AR: prøv bro med æ først, så ascii-fallback, deretter registry_db
AR_BACKEND = "none"
try:
    from .ar_bridge import open_db as ar_open_db, get_owners, companies_owned_by, normalize_orgnr

    AR_BACKEND = "ar_bridge"
except Exception:
    try:
        # Absolutt bro med æ-navn
        from app.ICRtracker.ar_bridge import open_db as ar_open_db, get_owners, companies_owned_by, normalize_orgnr

        AR_BACKEND = "ar_bridge"
    except Exception:
        try:
            # registry_db fallback
            from .registry_db import open_db as ar_open_db, get_owners, companies_owned_by, \
                normalize_orgnr  # type: ignore

            AR_BACKEND = "registry_db"
        except Exception:
            AR_BACKEND = "none"

# ------------------ Standard/lagring av innstillinger ------------------

APPDATA = Path.home() / "AppData" / "Roaming" / "ICRtracker"
APPDATA.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APPDATA / "gui_config.json"


@dataclass
class GuiSettings:
    # Tracker/Outlook
    sender_email: str = "IRC-Norway@no.gt.com"
    lookback_days: int = 14
    download_dir: str = r"F:\Dokument\Kildefiler\irc\irc_vedlegg"
    klientliste: str = r"F:\Dokument\Kildefiler\BHL AS klientliste - kopi.xlsx"
    rapport_csv: str = r"F:\Dokument\Kildefiler\irc\rapporter\match_rapport.csv"
    rapport_ar_csv: str = r"F:\Dokument\Kildefiler\irc\rapporter\ar_funn_fra_epost.csv"
    set_outlook_category: bool = True
    outlook_category_name: str = "Processed (IRC)"
    shared_mailbox_displayname: str = ""

    # AR/relasjoner
    enable_ar_recursive: bool = False
    ar_max_depth: int = 2  # rekursiv dybde (1=kun direkte)
    ar_include_up: bool = True  # eiere oppover
    ar_include_down: bool = True  # datterselskap nedover

    # Varsling
    notify_email_enabled: bool = False
    notify_email_to: str = ""  # komma-separert e-postliste
    notify_subject: str = "IRC-Tracker: treff funnet"
    notify_only_on_matches: bool = True

    # Planlegging (intern i GUI)
    schedule_enabled: bool = False
    schedule_time: str = "09:00"  # HH:MM (24t)

    # Evt. registry_db-sti brukt om ar_bridge ikke er aktiv
    registry_db_path: str = r"F:\Dokument\Kildefiler\aksjonarregister.db"

    # Annet
    min_name_score: int = 90  # brukes av fuzzy i matcher (ikke direkte her, men bevares)


def load_settings() -> GuiSettings:
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            return GuiSettings(**data)
        except Exception:
            traceback.print_exc()
    return GuiSettings()


def save_settings(s: GuiSettings):
    CONFIG_FILE.write_text(json.dumps(asdict(s), indent=2, ensure_ascii=False), encoding="utf-8")


# ------------------ Epost via Outlook (pywin32) ------------------

def send_outlook_email(to_csv: str, subject: str, body: str, attachments: Optional[List[Path]] = None):
    """
    Sender epost via lokal Outlook-klient. 'to_csv' kan være komma-separert liste.
    """
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        messagebox.showerror("Epost", f"Outlook (pywin32) ikke tilgjengelig: {e}")
        return

    recipients = [x.strip() for x in (to_csv or "").split(",") if x.strip()]
    if not recipients:
        messagebox.showwarning("Epost", "Ingen mottakere satt.")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = "; ".join(recipients)
        mail.Subject = subject
        mail.Body = body
        if attachments:
            for p in attachments:
                if Path(p).exists():
                    mail.Attachments.Add(str(p))
        mail.Send()
        messagebox.showinfo("Epost", f"Epost sendt til: {', '.join(recipients)}")
    except Exception as e:
        messagebox.showerror("Epost", f"Kunne ikke sende epost: {e}")


# ------------------ Enkle verktøy ------------------

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)


def parse_hhmm(s: str) -> Optional[time]:
    try:
        hh, mm = s.strip().split(":")
        return time(int(hh), int(mm))
    except Exception:
        return None


def read_csv_rows(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    rows: List[Dict[str, Any]] = []
    with path.open("r", encoding="utf-8", newline="") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(dict(row))
    return rows


# ------------------ AR rekursiv BFS ------------------

def ar_bfs_expand(orgs: List[str], max_depth: int, include_up: bool, include_down: bool,
                  registry_db_path: Path) -> List[Dict[str, Any]]:
    """
    Utvider relasjoner rekursivt via ar_bridge/registry_db.
    Returnerer rader med samme felt som AR-rapporten i tracker.
    """
    if AR_BACKEND == "none":
        return []

    try:
        conn = ar_open_db(registry_db_path)
    except Exception as e:
        messagebox.showerror("AR", f"Kunne ikke åpne AR: {e}")
        return []

    seen_orgs = set(orgs)
    rows_out: List[Dict[str, Any]] = []
    frontier = [(o, 0) for o in orgs]

    while frontier:
        cur, depth = frontier.pop(0)
        if depth >= max_depth:
            continue

        if include_up:
            try:
                for r in get_owners(conn, cur):
                    rel_hit = False  # flagges i tracker mot klientliste; her viser vi bare nettverket
                    rows_out.append({
                        "client_orgnr": cur,
                        "direction": "owned_by",
                        "related_orgnr": r.get("shareholder_orgnr", "") or "",
                        "related_name": r.get("shareholder_name", "") or "",
                        "related_type": r.get("shareholder_type", ""),
                        "stake_percent": r.get("stake_percent"),
                        "shares": r.get("shares"),
                        "flag_client_crosshit": "NEI",
                    })
                    nxt = normalize_orgnr(r.get("shareholder_orgnr", ""))
                    if nxt and nxt not in seen_orgs:
                        seen_orgs.add(nxt)
                        frontier.append((nxt, depth + 1))
            except Exception:
                traceback.print_exc()

        if include_down:
            try:
                for r in companies_owned_by(conn, cur):
                    rows_out.append({
                        "client_orgnr": cur,
                        "direction": "owns",
                        "related_orgnr": r.get("company_orgnr", ""),
                        "related_name": r.get("company_name", ""),
                        "related_type": "company",
                        "stake_percent": r.get("stake_percent"),
                        "shares": r.get("shares"),
                        "flag_client_crosshit": "NEI",
                    })
                    nxt = normalize_orgnr(r.get("company_orgnr", ""))
                    if nxt and nxt not in seen_orgs:
                        seen_orgs.add(nxt)
                        frontier.append((nxt, depth + 1))
            except Exception:
                traceback.print_exc()

    return rows_out


# ------------------ GUI-klassen ------------------

class ICRTrackerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IRC-Tracker – Kontrollpanel")
        self.geometry("1200x720")
        self.settings = load_settings()

        self._build_ui()
        self._load_from_settings()
        self._start_internal_scheduler()

    # ---------- UI ----------

    def _build_ui(self):
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self, padding=8)
        left.grid(row=0, column=0, sticky="ns")

        right = ttk.Frame(self, padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # --- Seksjon: Tracker/Outlook ---
        box1 = ttk.LabelFrame(left, text="Kjøring (Outlook/Tracker)", padding=8)
        box1.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        self.var_sender = tk.StringVar()
        self.var_lookback = tk.IntVar()
        self.var_download = tk.StringVar()
        self.var_klient = tk.StringVar()
        self.var_rapport = tk.StringVar()
        self.var_rapport_ar = tk.StringVar()
        self.var_setcat = tk.BooleanVar()
        self.var_catname = tk.StringVar()
        self.var_shared_mb = tk.StringVar()

        def row(parent, r, label, var, browse=False, is_int=False):
            ttk.Label(parent, text=label, anchor="w").grid(row=r, column=0, sticky="w")
            if is_int:
                e = ttk.Spinbox(parent, from_=1, to=365, textvariable=var, width=6)
                e.grid(row=r, column=1, sticky="w", padx=4)
            else:
                e = ttk.Entry(parent, textvariable=var, width=45)
                e.grid(row=r, column=1, sticky="w", padx=4)
            if browse:
                def choose():
                    p = filedialog.askopenfilename()
                    if p:
                        var.set(p)

                ttk.Button(parent, text="…", width=3, command=choose).grid(row=r, column=2, padx=2)
            return e

        row(box1, 0, "Avsender (SMTP):", self.var_sender)
        row(box1, 1, "Dager tilbake:", self.var_lookback, is_int=True)
        row(box1, 2, "Nedlastingsmappe:", self.var_download, browse=True)
        row(box1, 3, "Klientliste (xlsx/csv):", self.var_klient, browse=True)
        row(box1, 4, "Rapport (matcher).csv:", self.var_rapport, browse=True)
        row(box1, 5, "Rapport (AR).csv:", self.var_rapport_ar, browse=True)

        ttk.Checkbutton(box1, text="Merk epost med kategori ved treff",
                        variable=self.var_setcat).grid(row=6, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(box1, textvariable=self.var_catname, width=30).grid(row=6, column=1, sticky="w", pady=(6, 0))
        ttk.Label(box1, text="Delt postboks (visningsnavn, valgfri):").grid(row=7, column=0, sticky="w")
        ttk.Entry(box1, textvariable=self.var_shared_mb, width=30).grid(row=7, column=1, sticky="w")

        # --- Seksjon: AR/Relasjoner ---
        box2 = ttk.LabelFrame(left, text=f"Aksjonærregister (backend={AR_BACKEND})", padding=8)
        box2.grid(row=1, column=0, sticky="ew", padx=4, pady=4)

        self.var_ar_rec = tk.BooleanVar()
        self.var_ar_depth = tk.IntVar()
        self.var_ar_up = tk.BooleanVar()
        self.var_ar_down = tk.BooleanVar()
        self.var_registry_db = tk.StringVar()

        ttk.Checkbutton(box2, text="Rekursiv utvidelse av AR-relasjoner etter kjøring",
                        variable=self.var_ar_rec).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(box2, text="Maks dybde:").grid(row=1, column=0, sticky="w")
        ttk.Spinbox(box2, from_=1, to=6, textvariable=self.var_ar_depth, width=5).grid(row=1, column=1, sticky="w")
        ttk.Checkbutton(box2, text="Oppad (eiere)", variable=self.var_ar_up).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(box2, text="Nedad (datterselskap)", variable=self.var_ar_down).grid(row=2, column=1, sticky="w")

        ttk.Label(box2, text="registry_db (fallback-sti):").grid(row=3, column=0, sticky="w")
        ttk.Entry(box2, textvariable=self.var_registry_db, width=45).grid(row=3, column=1, sticky="w")

        # --- Seksjon: Varsling ---
        box3 = ttk.LabelFrame(left, text="Varsling (Outlook e-post)", padding=8)
        box3.grid(row=2, column=0, sticky="ew", padx=4, pady=4)

        self.var_notify_enabled = tk.BooleanVar()
        self.var_notify_to = tk.StringVar()
        self.var_notify_subject = tk.StringVar()
        self.var_notify_only_hits = tk.BooleanVar()

        ttk.Checkbutton(box3, text="Send e-post ved treff",
                        variable=self.var_notify_enabled).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(box3, text="Til (komma-separert):").grid(row=1, column=0, sticky="w")
        ttk.Entry(box3, textvariable=self.var_notify_to, width=45).grid(row=1, column=1, sticky="w")
        ttk.Label(box3, text="Emne:").grid(row=2, column=0, sticky="w")
        ttk.Entry(box3, textvariable=self.var_notify_subject, width=45).grid(row=2, column=1, sticky="w")
        ttk.Checkbutton(box3, text="Kun når det finnes treff",
                        variable=self.var_notify_only_hits).grid(row=3, column=0, columnspan=2, sticky="w")

        # --- Seksjon: Planlegging (intern) ---
        box4 = ttk.LabelFrame(left, text="Planlegging (intern i GUI)", padding=8)
        box4.grid(row=3, column=0, sticky="ew", padx=4, pady=4)

        self.var_sched_enabled = tk.BooleanVar()
        self.var_sched_time = tk.StringVar()

        ttk.Checkbutton(box4, text="Kjør daglig kl.:", variable=self.var_sched_enabled).grid(row=0, column=0,
                                                                                             sticky="w")
        ttk.Entry(box4, textvariable=self.var_sched_time, width=8).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(box4, text="(Format HH:MM, 24t)").grid(row=0, column=2, sticky="w")

        # --- Knapperekke (lagre/kjør) ---
        btns = ttk.Frame(left)
        btns.grid(row=4, column=0, sticky="ew", padx=4, pady=6)
        ttk.Button(btns, text="Lagre innstillinger", command=self.on_save).grid(row=0, column=0, padx=2)
        ttk.Button(btns, text="Kjør nå", command=self.on_run_now).grid(row=0, column=1, padx=2)
        ttk.Button(btns, text="Åpne rapportmappe", command=self.on_open_report_dir).grid(row=0, column=2, padx=2)

        # --- Resultatpanel ---
        ttk.Label(right, text="Resultater").grid(row=0, column=0, sticky="w")

        self.notebook = ttk.Notebook(right)
        self.notebook.grid(row=1, column=0, sticky="nsew")

        # Tab 1: matcher
        self.tab_matches = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_matches, text="Matcher (vedlegg)")

        self.tree_matches = self._make_tree(self.tab_matches, columns=[
            ("received", "Mottatt"), ("subject", "Emne"), ("attachment", "Vedlegg"),
            ("orgnr", "Unique ID"), ("match", "Match")
        ])

        # Tab 2: AR-relasjoner
        self.tab_ar = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_ar, text="AR-relasjoner")

        self.tree_ar = self._make_tree(self.tab_ar, columns=[
            ("received", "Mottatt"), ("subject", "Emne"), ("attachment", "Vedlegg"),
            ("client_orgnr", "Klient orgnr"), ("direction", "Retning"),
            ("related_orgnr", "Rel.orgnr"), ("related_name", "Rel.navn"),
            ("related_type", "Type"), ("stake_percent", "Andel %"),
            ("shares", "Aksjer"), ("flag_client_crosshit", "Krysstreff")
        ])

        # Statuslinje
        self.status = tk.StringVar(value=f"Klar. AR-backend: {AR_BACKEND}")
        ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken").grid(
            row=1, column=0, columnspan=2, sticky="ew"
        )

    def _make_tree(self, parent, columns: List[Tuple[str, str]]) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for key, label in columns:
            tree.heading(key, text=label)
            tree.column(key, width=120, anchor="w")
        return tree

    # ---------- Inn/ut av settings ----------

    def _load_from_settings(self):
        s = self.settings
        self.var_sender.set(s.sender_email)
        self.var_lookback.set(s.lookback_days)
        self.var_download.set(s.download_dir)
        self.var_klient.set(s.klientliste)
        self.var_rapport.set(s.rapport_csv)
        self.var_rapport_ar.set(s.rapport_ar_csv)
        self.var_setcat.set(s.set_outlook_category)
        self.var_catname.set(s.outlook_category_name)
        self.var_shared_mb.set(s.shared_mailbox_displayname)

        self.var_ar_rec.set(s.enable_ar_recursive)
        self.var_ar_depth.set(s.ar_max_depth)
        self.var_ar_up.set(s.ar_include_up)
        self.var_ar_down.set(s.ar_include_down)
        self.var_registry_db.set(s.registry_db_path)

        self.var_notify_enabled.set(s.notify_email_enabled)
        self.var_notify_to.set(s.notify_email_to)
        self.var_notify_subject.set(s.notify_subject)
        self.var_notify_only_hits.set(s.notify_only_on_matches)

        self.var_sched_enabled.set(s.schedule_enabled)
        self.var_sched_time.set(s.schedule_time)

    def _save_to_settings(self):
        s = self.settings
        s.sender_email = self.var_sender.get().strip()
        s.lookback_days = int(self.var_lookback.get())
        s.download_dir = self.var_download.get().strip()
        s.klientliste = self.var_klient.get().strip()
        s.rapport_csv = self.var_rapport.get().strip()
        s.rapport_ar_csv = self.var_rapport_ar.get().strip()
        s.set_outlook_category = bool(self.var_setcat.get())
        s.outlook_category_name = self.var_catname.get().strip()
        s.shared_mailbox_displayname = self.var_shared_mb.get().strip()

        s.enable_ar_recursive = bool(self.var_ar_rec.get())
        s.ar_max_depth = int(self.var_ar_depth.get())
        s.ar_include_up = bool(self.var_ar_up.get())
        s.ar_include_down = bool(self.var_ar_down.get())
        s.registry_db_path = self.var_registry_db.get().strip()

        s.notify_email_enabled = bool(self.var_notify_enabled.get())
        s.notify_email_to = self.var_notify_to.get().strip()
        s.notify_subject = self.var_notify_subject.get().strip()
        s.notify_only_on_matches = bool(self.var_notify_only_hits.get())

        s.schedule_enabled = bool(self.var_sched_enabled.get())
        s.schedule_time = self.var_sched_time.get().strip()

    # ---------- Callbacks ----------

    def on_save(self):
        self._save_to_settings()
        save_settings(self.settings)
        self.status.set("Innstillinger lagret.")

    def on_open_report_dir(self):
        p = Path(self.var_rapport.get() or ".")
        try:
            folder = p.parent if p.suffix else p
            if not folder.exists():
                folder = Path(self.var_download.get()).resolve()
            subprocess.Popen(f'explorer "{str(folder)}"')
        except Exception as e:
            messagebox.showerror("Åpne mappe", f"Kunne ikke åpne mappe: {e}")

    def on_run_now(self):
        self._save_to_settings()
        save_settings(self.settings)

        # Kjør i bakgrunnstråd for ikke å fryse GUI
        t = threading.Thread(target=self._run_pipeline, daemon=True)
        t.start()

    # ---------- Kjøring + opplasting av resultater ----------

    def _run_pipeline(self):
        s = self.settings
        self.status.set("Kjører tracker…")

        # Sett tracker-konstanter før kjøring
        try:
            tracker.SENDER_EMAIL = s.sender_email
            tracker.LOOKBACK_DAYS = s.lookback_days
            tracker.DOWNLOAD_DIR = Path(s.download_dir)
            tracker.KLIENTLISTE = Path(s.klientliste)
            tracker.RAPPORT_FIL = Path(s.rapport_csv)
            tracker.RAPPORT_AR_FRA_EPOST = Path(s.rapport_ar_csv)
            tracker.SET_OUTLOOK_CATEGORY = s.set_outlook_category
            tracker.OUTLOOK_CATEGORY_NAME = s.outlook_category_name
            tracker.SHARED_MAILBOX_DISPLAYNAME = s.shared_mailbox_displayname
            tracker.REGISTRY_DB_PATH = Path(s.registry_db_path)
        except Exception as e:
            messagebox.showerror("Tracker", f"Kunne ikke sette tracker-innstillinger: {e}")
            return

        # Kjør tracker.main()
        try:
            tracker.main()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Tracker", f"Kjøring feilet: {e}")
            return

        # Les rapportene og vis i tabeller
        matches = read_csv_rows(Path(s.rapport_csv))
        self._populate_tree(self.tree_matches, matches)

        ar_rows = read_csv_rows(Path(s.rapport_ar_csv))
        if s.enable_ar_recursive and AR_BACKEND != "none":
            # utvid rekursivt fra alle orgnr (fra matcher eller AR)
            seed_orgs = set()
            for r in matches:
                if r.get("orgnr"):
                    seed_orgs.add(tracker.normalize_orgnr_local(r["orgnr"]))
            for r in ar_rows:
                if r.get("client_orgnr"):
                    seed_orgs.add(tracker.normalize_orgnr_local(r["client_orgnr"]))

            if seed_orgs:
                extra = ar_bfs_expand(
                    list(seed_orgs),
                    max_depth=s.ar_max_depth,
                    include_up=s.ar_include_up,
                    include_down=s.ar_include_down,
                    registry_db_path=Path(s.registry_db_path),
                )
                # slå sammen
                ar_rows = list(ar_rows) + extra

        self._populate_tree(self.tree_ar, ar_rows)

        # Varsle på epost
        if s.notify_email_enabled:
            has_hits = any(r.get("match", "").upper() == "JA" for r in matches)
            if (not s.notify_only_on_matches) or has_hits:
                body = [
                    f"Kjøringstid: {datetime.now():%Y-%m-%d %H:%M}",
                    f"Treff (matcher.csv): {sum(1 for r in matches if (r.get('match', '').upper() == 'JA'))}",
                    f"Rader i AR-rapport: {len(ar_rows)}",
                    "",
                    f"Rapporter:",
                    f" - {s.rapport_csv}",
                    f" - {s.rapport_ar_csv}"
                ]
                try:
                    send_outlook_email(
                        to_csv=s.notify_email_to,
                        subject=s.notify_subject,
                        body="\n".join(body),
                        attachments=[Path(s.rapport_csv)] + (
                            [Path(s.rapport_ar_csv)] if Path(s.rapport_ar_csv).exists() else []),
                    )
                except Exception:
                    traceback.print_exc()

        self.status.set(f"Ferdig. {len(matches)} linjer matcher; {len(ar_rows)} AR-linjer. (Backend={AR_BACKEND})")

    def _populate_tree(self, tree: ttk.Treeview, rows: List[Dict[str, Any]]):
        tree.delete(*tree.get_children())
        if not rows:
            return
        cols = list(tree["columns"])
        for r in rows:
            values = [r.get(c, "") for c in cols]
            tree.insert("", "end", values=values)

    # ---------- Intern planlegging (GUI må stå åpen) ----------

    def _start_internal_scheduler(self):
        # sjekk hvert 30. sekund om vi skal kjøre
        self._schedule_check()

    def _schedule_check(self):
        try:
            if self.settings.schedule_enabled:
                t = parse_hhmm(self.settings.schedule_time)
                now = datetime.now()
                if t and time(now.hour, now.minute) == t and now.second < 30:
                    # unngå duplikat (enkelt): kjør kun hvis siste kjøring var > 60 sekunder siden
                    if not hasattr(self, "_last_run") or (now - getattr(self, "_last_run")).total_seconds() > 60:
                        self._last_run = now
                        threading.Thread(target=self._run_pipeline, daemon=True).start()
        except Exception:
            traceback.print_exc()
        finally:
            # sjekk igjen om 30 sek
            self.after(30_000, self._schedule_check)


# ------------------ Main ------------------

def main():
    app = ICRTrackerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()

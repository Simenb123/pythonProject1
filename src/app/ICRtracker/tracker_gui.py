# -*- coding: utf-8 -*-
"""
tracker_gui.py
----------------
Tkinter-GUI for å kjøre og overvåke IRC-tracker + aksjonærregister-søk.

Funksjoner:
- Kjør "nå" hele pipeline: hent epost, lagre vedlegg, plukk Unique ID, match mot klientliste,
  og (valgfritt) slå opp relasjoner i aksjonærregisteret opp/ned i ønsket dybde.
- Vis resultater i tabeller (grunnmatcher + AR-relasjoner).
- Eksporter CSV av begge tabeller.
- Send epostvarsel via Outlook ved treff.
- Lagre/les konfig (JSON) i %APPDATA%\IRCTracker\gui_config.json (fallback: ved siden av skriptet).
- Opprette/slette Task Scheduler-jobb (Windows) for daglig kjøring.

Krever:
  pip install pywin32 openpyxl

Struktur i prosjekt:
  src/app/ICRtracker/tracker_gui.py  (denne filen)
  src/app/ICRtracker/tracker.py      (gjenbruker mye logikk herfra)
  src/app/ICRtracker/ar_bridge.py    (for AR; fallback til registry_db hvis ikke tilgjengelig)
"""
from __future__ import annotations

import csv
import json
import os
import queue
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --------------------- Robust modul-importer ---------------------

# Finn <prosjektrot>/src automatisk dersom vi kjøres direkte
if __name__ == "__main__" and __package__ is None:
    here = Path(__file__).resolve()
    for parent in (here.parent, *here.parents):
        if (parent / "app" / "__init__.py").exists():
            sys.path.insert(0, str(parent))
            break

# Importer tracker-funksjoner
try:
    from app.ICRtracker.tracker import (
        ensure_dir, read_client_orgs, fetch_messages_from_sender,
        save_excel_attachments, extract_unique_ids, normalize_orgnr_local,
        SENDER_EMAIL as DEFAULT_SENDER_EMAIL,
        LOOKBACK_DAYS as DEFAULT_LOOKBACK,
        DOWNLOAD_DIR as DEFAULT_DOWNLOAD_DIR,
        KLIENTLISTE as DEFAULT_CLIENTS_PATH,
        RAPPORT_FIL as DEFAULT_REPORT_PATH,
        RAPPORT_AR_FRA_EPOST as DEFAULT_AR_REPORT_PATH,
        REGISTRY_DB_PATH as DEFAULT_REG_DB_PATH,
    )
    TRACKER_OK = True
except Exception as e:
    TRACKER_OK = False
    TRACKER_ERR = e

# Importer AR-bro (foretrukket), ellers fall tilbake til registry_db
AR_AVAILABLE = False
AR_BACKEND = "none"
try:
    from app.ICRtracker.ar_bridge import open_db as ar_open_db, get_owners, companies_owned_by, normalize_orgnr  # type: ignore
    AR_AVAILABLE = True
    AR_BACKEND = "ar_bridge"
except Exception:
    try:
        from app.ICRtracker.registry_db import open_db as ar_open_db, get_owners, companies_owned_by, normalize_orgnr  # type: ignore
        AR_AVAILABLE = True
        AR_BACKEND = "registry_db"
    except Exception:
        AR_AVAILABLE = False
        AR_BACKEND = "none"

# ----------------------- Konfigdatamodell -----------------------

def _default_path(p: Path) -> str:
    try:
        return str(p)
    except Exception:
        return ""

@dataclass
class GUIConfig:
    sender_email: str = DEFAULT_SENDER_EMAIL if TRACKER_OK else "IRC-Norway@no.gt.com"
    lookback_days: int = DEFAULT_LOOKBACK if TRACKER_OK else 14
    download_dir: str = _default_path(DEFAULT_DOWNLOAD_DIR) if TRACKER_OK else ""
    clients_path: str = _default_path(DEFAULT_CLIENTS_PATH) if TRACKER_OK else ""
    report_path: str = _default_path(DEFAULT_REPORT_PATH) if TRACKER_OK else ""
    ar_report_path: str = _default_path(DEFAULT_AR_REPORT_PATH) if TRACKER_OK else ""
    registry_db_path: str = _default_path(DEFAULT_REG_DB_PATH) if TRACKER_OK else ""

    ar_enable: bool = True                   # Slå på AR-relasjoner
    ar_up_depth: int = 1                     # Oppstrøms dybde (eiere)
    ar_down_depth: int = 1                   # Nedstrøms dybde (datterselskap)
    ar_min_stake: float = 0.0                # Filtrering: min eierandel (i %)

    notify_enable: bool = False              # Send epost ved treff
    notify_to: str = ""                      # Mottakere, separert med ';'

    schedule_enable: bool = False            # Opprett planlagt jobb
    schedule_time: str = "09:00"             # HH:MM (24t)

# ----------------------- Hjelpefunksjoner -----------------------

def appdata_config_path() -> Path:
    """Finn sted å lagre konfig (AppData/Roaming/IRCTracker/gui_config.json)."""
    base = os.getenv("APPDATA")
    if base:
        p = Path(base) / "IRCTracker"
        p.mkdir(parents=True, exist_ok=True)
        return p / "gui_config.json"
    # fallback: ved siden av skriptet
    return Path(__file__).resolve().with_name("gui_config.json")

def load_config() -> GUIConfig:
    p = appdata_config_path()
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            return GUIConfig(**data)
        except Exception:
            pass
    # default
    cfg = GUIConfig()
    # Forsøk å fylle bedre defaults hvis tracker er tilgjengelig
    return cfg

def save_config(cfg: GUIConfig):
    p = appdata_config_path()
    p.write_text(json.dumps(asdict(cfg), ensure_ascii=False, indent=2), encoding="utf-8")

def shell_quote(s: str) -> str:
    if '"' in s:
        return f'"{s.replace("\"", "\\\"")}"'
    return f'"{s}"'

def send_outlook_mail(to_list: List[str], subject: str, body: str, attachments: Optional[List[Path]] = None) -> str:
    """Send epost via Outlook (pywin32). Returner status-tekst."""
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        return f"Outlook COM utilgjengelig: {e}"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = ";".join(t.strip() for t in to_list if t.strip())
        mail.Subject = subject
        mail.Body = body
        for att in attachments or []:
            try:
                mail.Attachments.Add(str(att))
            except Exception as e:
                # Ignorer vedleggsfeil, men noter i tekst
                mail.Body += f"\n(Obs: kunne ikke legge ved {att.name}: {e})"
        mail.Send()
        return "Epost sendt."
    except Exception as e:
        return f"Feil ved sending: {e}"

# --------------- AR-traversering (opp/ned i dybde) ---------------

def ar_traverse(conn, base_orgnr: str, up_depth: int, down_depth: int, min_stake: float = 0.0) -> List[Dict[str, object]]:
    """
    Returnerer kant-liste (relasjonsrader) fra base_orgnr:
      direction: 'owned_by' (oppstrøms) eller 'owns' (nedstrøms)
      related_orgnr, related_name, stake_percent, shares, level (1..N)
    """
    seen_up = set([normalize_orgnr(base_orgnr)])
    seen_down = set([normalize_orgnr(base_orgnr)])
    rows: List[Dict[str, object]] = []

    # Oppstrøms (eiere)
    frontier = [normalize_orgnr(base_orgnr)]
    for level in range(1, up_depth + 1):
        new_frontier = []
        for org in frontier:
            try:
                owners = get_owners(conn, org)
            except Exception as e:
                owners = []
            for r in owners:
                pct = r.get("stake_percent")
                if pct is not None and min_stake is not None and pct < min_stake:
                    continue
                rel_org = normalize_orgnr(r.get("shareholder_orgnr") or "")
                rel_name = r.get("shareholder_name") or ""
                rows.append({
                    "direction": "owned_by",
                    "level": level,
                    "client_orgnr": org,
                    "related_orgnr": rel_org,
                    "related_name": rel_name,
                    "related_type": r.get("shareholder_type") or "",
                    "stake_percent": pct,
                    "shares": r.get("shares")
                })
                if rel_org and rel_org not in seen_up:
                    seen_up.add(rel_org)
                    new_frontier.append(rel_org)
        frontier = new_frontier

    # Nedstrøms (datterselskap)
    frontier = [normalize_orgnr(base_orgnr)]
    for level in range(1, down_depth + 1):
        new_frontier = []
        for org in frontier:
            try:
                childs = companies_owned_by(conn, org)
            except Exception as e:
                childs = []
            for r in childs:
                pct = r.get("stake_percent")
                if pct is not None and min_stake is not None and pct < min_stake:
                    continue
                rel_org = normalize_orgnr(r.get("company_orgnr") or "")
                rel_name = r.get("company_name") or ""
                rows.append({
                    "direction": "owns",
                    "level": level,
                    "client_orgnr": org,
                    "related_orgnr": rel_org,
                    "related_name": rel_name,
                    "related_type": "company",
                    "stake_percent": pct,
                    "shares": r.get("shares")
                })
                if rel_org and rel_org not in seen_down:
                    seen_down.add(rel_org)
                    new_frontier.append(rel_org)
        frontier = new_frontier

    return rows

# ----------------------- Pipeline (worker) -----------------------

@dataclass
class RunResult:
    rows_basic: List[Dict[str, object]]
    rows_ar: List[Dict[str, object]]
    saved_files: List[Path]

def run_pipeline(cfg: GUIConfig, log: queue.Queue) -> RunResult:
    """
    Kjører hele løypa. Returnerer resultater i minne (ingen filer skrives automatisk).
    Logger linjer via queue (for UI).
    """
    def qlog(msg: str):
        log.put(msg)

    if not TRACKER_OK:
        qlog(f"FEIL: tracker-moduler kunne ikke importeres: {TRACKER_ERR!r}")
        return RunResult([], [], [])

    download_dir = Path(cfg.download_dir) if cfg.download_dir else DEFAULT_DOWNLOAD_DIR
    ensure_dir(download_dir)

    # 1) Klientliste
    try:
        client_orgs = read_client_orgs(Path(cfg.clients_path))
        qlog(f"Klientliste: {len(client_orgs)} orgnr lastet.")
    except Exception as e:
        qlog(f"FEIL: Kunne ikke lese klientliste: {e}")
        return RunResult([], [], [])

    # 2) Outlook-meldinger
    try:
        msgs = fetch_messages_from_sender(cfg.sender_email)
        qlog(f"Fant {len(msgs)} meldinger fra {cfg.sender_email} (siste {cfg.lookback_days} dager).")
    except Exception as e:
        qlog(f"FEIL: Kunne ikke hente eposter: {e}")
        return RunResult([], [], [])

    rows_basic: List[Dict[str, object]] = []
    rows_ar: List[Dict[str, object]] = []
    saved_files: List[Path] = []

    # 3) AR-kobling (valgfri)
    conn = None
    if cfg.ar_enable and AR_AVAILABLE:
        try:
            conn = ar_open_db(Path(cfg.registry_db_path) if cfg.registry_db_path else None)
            qlog(f"AR-backend aktiv ({AR_BACKEND}).")
        except Exception as e:
            qlog(f"ADVARSEL: Kunne ikke åpne AR: {e}")
            conn = None
    elif cfg.ar_enable and not AR_AVAILABLE:
        qlog("ADVARSEL: AR ikke tilgjengelig (mangler ar_bridge/registry_db).")

    # 4) Gå gjennom epostene
    for msg in msgs:
        subject = getattr(msg, "Subject", "")
        received = getattr(msg, "ReceivedTime", None)
        received_iso = (received.strftime("%Y-%m-%d %H:%M") if received else "")

        files = save_excel_attachments(msg, download_dir)
        saved_files.extend(files)
        qlog(f" - '{subject}' → lagret {len(files)} vedlegg")

        for f in files:
            try:
                ids = extract_unique_ids(f)
            except Exception as e:
                qlog(f"   ! Kunne ikke lese Unique ID fra {f.name}: {e}")
                ids = []
            for org in ids:
                hit = "JA" if org in client_orgs else "NEI"
                rows_basic.append({
                    "received": received_iso,
                    "subject": subject,
                    "attachment": f.name,
                    "orgnr": org,
                    "match": hit
                })

                # AR-relasjoner for hvert orgnr
                if conn:
                    try:
                        rels = ar_traverse(conn, org, cfg.ar_up_depth, cfg.ar_down_depth, cfg.ar_min_stake)
                        # Merk crosshit mot klientliste
                        for r in rels:
                            rel_hit = r["related_orgnr"] and (normalize_orgnr(r["related_orgnr"]) in client_orgs)
                            r["flag_client_crosshit"] = "JA" if rel_hit else "NEI"
                            r["source_attachment"] = f.name
                            r["received"] = received_iso
                            r["subject"] = subject
                        rows_ar.extend(rels)
                    except Exception as e:
                        qlog(f"   ! AR-feil for {org}: {e}")

    qlog(f"Ferdig: {len(rows_basic)} grunnlinjer, {len(rows_ar)} AR-relasjoner, {len(saved_files)} vedlegg.")
    return RunResult(rows_basic, rows_ar, saved_files)

# ---------------------------- GUI ----------------------------

class Table(ttk.Treeview):
    """Enkel tabell med automatisk kolonneoppsett + eksport til CSV."""
    def __init__(self, master, columns: List[str], **kw):
        super().__init__(master, columns=columns, show="headings", **kw)
        self._columns = columns
        for c in columns:
            self.heading(c, text=c)
            self.column(c, width=120, anchor="w")
        vsb = ttk.Scrollbar(master, orient="vertical", command=self.yview)
        hsb = ttk.Scrollbar(master, orient="horizontal", command=self.xview)
        self.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def load_rows(self, rows: List[Dict[str, object]]):
        self.delete(*self.get_children())
        for r in rows:
            vals = [r.get(c, "") for c in self._columns]
            self.insert("", "end", values=vals)

    def export_csv(self, dest: Path):
        with dest.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(self._columns)
            for iid in self.get_children():
                w.writerow(list(self.item(iid, "values")))

class TrackerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IRC-Tracker – Overvåkning og søk")
        self.geometry("1200x750")
        self.minsize(1000, 650)

        self.cfg = load_config()

        self._build_ui()

        if not TRACKER_OK:
            messagebox.showerror("Feil", f"Kunne ikke importere tracker-moduler:\n{TRACKER_ERR!r}")
        elif not AR_AVAILABLE and self.cfg.ar_enable:
            self._log(f"ADVARSEL: AR ikke tilgjengelig – {AR_BACKEND}")

    # --------------------- UI bygging ---------------------

    def _build_ui(self):
        # Top: konfig
        top = ttk.LabelFrame(self, text="Konfigurasjon")
        top.pack(side="top", fill="x", padx=10, pady=10)

        # Rad 1
        r1 = ttk.Frame(top); r1.pack(fill="x", pady=4)
        ttk.Label(r1, text="Avsender epost (filter):").grid(row=0, column=0, sticky="w")
        self.e_sender = ttk.Entry(r1, width=40)
        self.e_sender.grid(row=0, column=1, sticky="w", padx=8)
        self.e_sender.insert(0, self.cfg.sender_email)

        ttk.Label(r1, text="Se tilbake (dager):").grid(row=0, column=2, sticky="e")
        self.e_lookback = ttk.Spinbox(r1, from_=1, to=60, width=6)
        self.e_lookback.grid(row=0, column=3, sticky="w", padx=8)
        self.e_lookback.delete(0, "end"); self.e_lookback.insert(0, str(self.cfg.lookback_days))

        # Rad 2
        r2 = ttk.Frame(top); r2.pack(fill="x", pady=4)
        ttk.Label(r2, text="Nedlasting/vedlegg-mappe:").grid(row=0, column=0, sticky="w")
        self.e_dl = ttk.Entry(r2, width=70); self.e_dl.grid(row=0, column=1, sticky="we", padx=8)
        self.e_dl.insert(0, self.cfg.download_dir)
        ttk.Button(r2, text="Velg…", command=self._pick_dl).grid(row=0, column=2, sticky="w")

        # Rad 3
        r3 = ttk.Frame(top); r3.pack(fill="x", pady=4)
        ttk.Label(r3, text="Klientliste (xlsx/csv):").grid(row=0, column=0, sticky="w")
        self.e_clients = ttk.Entry(r3, width=70); self.e_clients.grid(row=0, column=1, sticky="we", padx=8)
        self.e_clients.insert(0, self.cfg.clients_path)
        ttk.Button(r3, text="Velg…", command=self._pick_clients).grid(row=0, column=2, sticky="w")

        # Rad 4 – AR og varsling
        r4 = ttk.Frame(top); r4.pack(fill="x", pady=4)
        self.v_ar = tk.BooleanVar(value=self.cfg.ar_enable)
        ttk.Checkbutton(r4, text="Slå på AR-relasjoner", variable=self.v_ar).grid(row=0, column=0, sticky="w")

        ttk.Label(r4, text="Opp-dybde:").grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.e_up = ttk.Spinbox(r4, from_=0, to=5, width=5); self.e_up.grid(row=0, column=2, sticky="w")
        self.e_up.delete(0, "end"); self.e_up.insert(0, str(self.cfg.ar_up_depth))

        ttk.Label(r4, text="Ned-dybde:").grid(row=0, column=3, sticky="e")
        self.e_down = ttk.Spinbox(r4, from_=0, to=5, width=5); self.e_down.grid(row=0, column=4, sticky="w")
        self.e_down.delete(0, "end"); self.e_down.insert(0, str(self.cfg.ar_down_depth))

        ttk.Label(r4, text="Min eierandel (%):").grid(row=0, column=5, sticky="e", padx=(16, 0))
        self.e_minpct = ttk.Spinbox(r4, from_=0, to=100, increment=0.5, width=6); self.e_minpct.grid(row=0, column=6, sticky="w")
        self.e_minpct.delete(0, "end"); self.e_minpct.insert(0, str(self.cfg.ar_min_stake))

        r4b = ttk.Frame(top); r4b.pack(fill="x", pady=2)
        self.v_notify = tk.BooleanVar(value=self.cfg.notify_enable)
        ttk.Checkbutton(r4b, text="Send epostvarsel ved treff", variable=self.v_notify).grid(row=0, column=0, sticky="w")
        ttk.Label(r4b, text="Mottakere (;-separert):").grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.e_notify = ttk.Entry(r4b, width=60); self.e_notify.grid(row=0, column=2, sticky="we", padx=8)
        self.e_notify.insert(0, self.cfg.notify_to)
        ttk.Button(r4b, text="Send testvarsel", command=self._send_testmail).grid(row=0, column=3, sticky="w")

        # Rad 5 – planlegging
        r5 = ttk.Frame(top); r5.pack(fill="x", pady=6)
        self.v_sched = tk.BooleanVar(value=self.cfg.schedule_enable)
        ttk.Checkbutton(r5, text="Planlagt jobb (Windows Task Scheduler)", variable=self.v_sched).grid(row=0, column=0, sticky="w")
        ttk.Label(r5, text="Tid (HH:MM):").grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.e_time = ttk.Entry(r5, width=8); self.e_time.grid(row=0, column=2, sticky="w")
        self.e_time.insert(0, self.cfg.schedule_time)
        ttk.Button(r5, text="Opprett planlagt jobb", command=self._create_task).grid(row=0, column=3, padx=6)
        ttk.Button(r5, text="Slett planlagt jobb", command=self._delete_task).grid(row=0, column=4)

        # Separator + knapperad
        sep = ttk.Separator(self); sep.pack(fill="x", padx=10, pady=6)
        kr = ttk.Frame(self); kr.pack(fill="x", padx=10, pady=4)
        ttk.Button(kr, text="Lagre innstillinger", command=self._save_cfg).pack(side="left")
        ttk.Button(kr, text="Kjør nå", command=self._run_now).pack(side="left", padx=8)

        # Resultat-Notebook
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=10, pady=6)

        # Tab 1 – grunnmatcher
        tab_basic = ttk.Frame(nb); nb.add(tab_basic, text="Grunnmatcher")
        cols_basic = ["received", "subject", "attachment", "orgnr", "match"]
        self.tbl_basic = Table(tab_basic, columns=cols_basic)
        self.tbl_basic.pack(fill="both", expand=True, padx=6, pady=6)
        btns_basic = ttk.Frame(tab_basic); btns_basic.pack(fill="x", padx=6, pady=(0,6))
        ttk.Button(btns_basic, text="Eksporter CSV…", command=lambda: self._export_table(self.tbl_basic, "grunnmatcher.csv")).pack(side="left")

        # Tab 2 – AR-relasjoner
        tab_ar = ttk.Frame(nb); nb.add(tab_ar, text="AR-relasjoner")
        cols_ar = ["received","subject","attachment","client_orgnr","direction","level",
                   "related_orgnr","related_name","related_type","stake_percent","shares","flag_client_crosshit"]
        self.tbl_ar = Table(tab_ar, columns=cols_ar)
        self.tbl_ar.pack(fill="both", expand=True, padx=6, pady=6)
        btns_ar = ttk.Frame(tab_ar); btns_ar.pack(fill="x", padx=6, pady=(0,6))
        ttk.Button(btns_ar, text="Eksporter CSV…", command=lambda: self._export_table(self.tbl_ar, "ar_relasjoner.csv")).pack(side="left")

        # Logg
        logf = ttk.LabelFrame(self, text="Logg"); logf.pack(fill="both", expand=False, padx=10, pady=(0,10))
        self.txt = tk.Text(logf, height=8); self.txt.pack(fill="both", expand=True)
        self._log(f"Tracker-moduler: {'OK' if TRACKER_OK else f'FEIL ({TRACKER_ERR!r})'}")
        self._log(f"AR-backend: {AR_BACKEND} (tilgjengelig={AR_AVAILABLE})")

    # --------------------- UI handlers ---------------------

    def _pick_dl(self):
        d = filedialog.askdirectory(title="Velg mappe for vedlegg")
        if d:
            self.e_dl.delete(0, "end"); self.e_dl.insert(0, d)

    def _pick_clients(self):
        p = filedialog.askopenfilename(title="Velg klientliste", filetypes=[("Excel/CSV", "*.xlsx;*.xlsm;*.csv"), ("Alle", "*.*")])
        if p:
            self.e_clients.delete(0, "end"); self.e_clients.insert(0, p)

    def _send_testmail(self):
        to = [x.strip() for x in self.e_notify.get().split(";") if x.strip()]
        if not to:
            messagebox.showwarning("Varsel", "Fyll inn mottakere først.")
            return
        status = send_outlook_mail(to, "Test – IRC-Tracker", "Dette er en test fra IRC-Tracker GUI.")
        messagebox.showinfo("Send test", status)

    def _save_cfg(self):
        cfg = self._read_cfg_from_ui()
        save_config(cfg)
        self._log("Innstillinger lagret.")

    def _read_cfg_from_ui(self) -> GUIConfig:
        cfg = GUIConfig(
            sender_email=self.e_sender.get().strip(),
            lookback_days=int(self.e_lookback.get()),
            download_dir=self.e_dl.get().strip(),
            clients_path=self.e_clients.get().strip(),
            report_path=self.cfg.report_path,             # behold sti fra tracker.py defaults
            ar_report_path=self.cfg.ar_report_path,
            registry_db_path=self.cfg.registry_db_path,

            ar_enable=bool(self.v_ar.get()),
            ar_up_depth=int(self.e_up.get()),
            ar_down_depth=int(self.e_down.get()),
            ar_min_stake=float(self.e_minpct.get()),

            notify_enable=bool(self.v_notify.get()),
            notify_to=self.e_notify.get().strip(),

            schedule_enable=bool(self.v_sched.get()),
            schedule_time=self.e_time.get().strip()
        )
        return cfg

    def _run_now(self):
        cfg = self._read_cfg_from_ui()
        save_config(cfg)

        self._log("Starter kjøring …")
        self.tbl_basic.load_rows([])
        self.tbl_ar.load_rows([])

        self._q = queue.Queue()
        self._result_container = {"res": None}

        def worker():
            try:
                res = run_pipeline(cfg, self._q)
                self._result_container["res"] = res
                # Epostvarsel hvis aktuelt
                if cfg.notify_enable:
                    basic_hits = [r for r in res.rows_basic if r.get("match") == "JA"]
                    ar_hits = [r for r in res.rows_ar if r.get("flag_client_crosshit") == "JA"]
                    if basic_hits or ar_hits:
                        to = [x.strip() for x in cfg.notify_to.split(";") if x.strip()]
                        body = f"Grunnmatcher: {len(basic_hits)}\nAR-treff: {len(ar_hits)}\n" \
                               f"Kjørt: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                        status = send_outlook_mail(to, "IRC-Tracker: TREFF", body)
                        self._q.put(f"Varsel: {status}")
            except Exception as e:
                self._q.put(f"FEIL i kjøring: {e}")

        threading.Thread(target=worker, daemon=True).start()
        self._poll_queue()

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self._log(msg)
        except queue.Empty:
            pass

        res = self._result_container.get("res")
        if res is None:
            self.after(150, self._poll_queue)
            return

        # Last inn tabeller
        self.tbl_basic.load_rows(res.rows_basic)
        self.tbl_ar.load_rows(res.rows_ar)
        self._log("Kjøring ferdig.")

    def _export_table(self, table: Table, default_name: str):
        p = filedialog.asksaveasfilename(
            title="Lagre CSV",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV", "*.csv")],
        )
        if not p:
            return
        try:
            table.export_csv(Path(p))
            messagebox.showinfo("Eksport", f"Lagret: {p}")
        except Exception as e:
            messagebox.showerror("Feil", f"Kunne ikke lagre CSV: {e}")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt.insert("end", f"[{ts}] {msg}\n")
        self.txt.see("end")

    # ----------------- Task Scheduler-integrasjon -----------------

    @property
    def _task_name(self) -> str:
        return "IRCTracker_Daily"

    def _create_task(self):
        """Opprett eller oppdater planlagt jobb via schtasks.exe."""
        cfg = self._read_cfg_from_ui()
        save_config(cfg)

        # Finn python og kjørekommando
        python = sys.executable
        # Kjør tracker som modul (krever at src er på sys.path; derfor bruker vi -m fra skript-dir)
        # Sikrere: kjør direkte filsti til tracker.py
        script_path = Path(__file__).resolve().parent / "tracker.py"

        if not script_path.exists():
            messagebox.showerror("Feil", f"Fant ikke tracker.py: {script_path}")
            return

        time_str = cfg.schedule_time.strip()
        if not time_str or len(time_str) != 5 or time_str[2] != ":":
            messagebox.showwarning("Tid", "Tid må være på format HH:MM (24t).")
            return

        # Bygg schtasks-kommando
        # /TR cmd må være hel streng – vi siterer filstier
        tr = f'{shell_quote(python)} {shell_quote(str(script_path))}'
        cmd = [
            "schtasks", "/Create", "/F",
            "/SC", "DAILY",
            "/TN", self._task_name,
            "/TR", tr,
            "/ST", time_str
        ]

        try:
            out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True, shell=False)
            self._log(out.strip())
            messagebox.showinfo("Planlagt jobb", "Planlagt jobb opprettet/oppdatert.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Feil", f"Kunne ikke opprette jobb:\n{e.output}")

    def _delete_task(self):
        cmd = ["schtasks", "/Delete", "/F", "/TN", self._task_name]
        try:
            out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True, shell=False)
            self._log(out.strip())
            messagebox.showinfo("Planlagt jobb", "Planlagt jobb slettet (hvis den fantes).")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Feil", f"Kunne ikke slette jobb:\n{e.output}")


# -------------------------- main --------------------------

def main():
    app = TrackerGUI()
    app.mainloop()

if __name__ == "__main__":
    main()

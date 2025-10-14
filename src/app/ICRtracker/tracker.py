# -*- coding: utf-8 -*-
"""
IRC-Tracker: Hent Outlook-eposter, lagre Excel-vedlegg, plukk 'Unique ID',
match mot klientliste (KLIENT_ORGNR) og skriv match_rapport.csv

Krever: pip install pywin32 openpyxl
Miljø: Windows + Outlook (logget inn)
"""

import csv
import re
from datetime import datetime, timedelta
from pathlib import Path

import win32com.client  # pywin32
from openpyxl import load_workbook


# ========================== KONFIG ==========================
# Riktig avsender (OBS: sjekk domenet – ofte "no.gt.com")
SENDER_EMAIL = "IRC-Norway@no.gt.com"

# Hvor langt tilbake i tid vi leter
LOOKBACK_DAYS = 14

# Hvor lagre vedlegg + hvor skrive rapport
DOWNLOAD_DIR = Path(r"F:\Dokument\Kildefiler\irc\irc_vedlegg")
RAPPORT_FIL  = Path(r"F:\Dokument\Kildefiler\irc\rapporter\match_rapport.csv")

# Klientliste (foretrekker .xlsx med kolonnen 'KLIENT_ORGNR', men .csv støttes)
KLIENTLISTE  = Path(r"F:\Dokument\Kildefiler\BHL AS klientliste - kopi.xlsx")

# Skal vi sette Outlook-kategori på meldinger med treff?
SET_OUTLOOK_CATEGORY = True
OUTLOOK_CATEGORY_NAME = "Processed (IRC)"

# Hent fra delt postboks i stedet for din primære?
# Eksempel: "IRC-Norway" eller full e-postadresse til delt postboks.
# La stå som "" for å bruke din egen innboks.
SHARED_MAILBOX_DISPLAYNAME = ""
# ===========================================================


# --------------------- Hjelpefunksjoner ---------------------
def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def normalize_orgnr(val) -> str:
    """Til kun siffer (fjerner NO-, mellomrom, punktum osv.)."""
    if val is None:
        return ""
    return re.sub(r"\D+", "", str(val))

def read_client_orgs(path: Path) -> set:
    """Les klientliste fra .xlsx (KLIENT_ORGNR) eller .csv (orgnr/klient_orgnr/organisasjonsnummer)."""
    if not path.exists():
        raise FileNotFoundError(f"Fant ikke klientliste: {path}")

    ext = path.suffix.lower()
    orgs = set()

    if ext in {".xlsx", ".xlsm"}:
        wb = load_workbook(filename=path, data_only=True, read_only=True)
        ws = wb.active  # Første ark (i fila di 'Sheet1')
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        name_to_idx = {str(h).strip().lower(): i+1 for i, h in enumerate(header)}
        # aksepter noen varianter:
        candidates = ["klient_orgnr", "orgnr", "organisasjonsnummer", "org number", "orgnumber"]
        col = None
        for c in candidates:
            if c in name_to_idx:
                col = name_to_idx[c]
                break
        if not col:
            raise ValueError("Fant ikke kolonnen 'KLIENT_ORGNR' (eller 'orgnr/organisasjonsnummer') i klientlista.")
        for r in ws.iter_rows(min_row=2, values_only=True):
            val = r[col-1] if len(r) >= col else None
            n = normalize_orgnr(val)
            if n:
                orgs.add(n)
        if not orgs:
            raise ValueError("Ingen orgnr-verdier funnet i klientlista.")
        return orgs

    elif ext == ".csv":
        def read_csv(encoding: str) -> set:
            s = set()
            with path.open("r", encoding=encoding, newline="") as f:
                reader = csv.DictReader(f)
                if not reader.fieldnames:
                    raise ValueError("CSV mangler header.")
                fmap = {c.lower().strip(): c for c in reader.fieldnames}
                candidates = ["klient_orgnr", "orgnr", "organisasjonsnummer"]
                key = None
                for c in candidates:
                    if c in fmap:
                        key = fmap[c]; break
                if not key:
                    raise ValueError("Fant ikke kolonne for orgnr i CSV.")
                for row in reader:
                    n = normalize_orgnr(row.get(key, ""))
                    if n:
                        s.add(n)
            return s

        try:
            orgs = read_csv("utf-8-sig")
        except UnicodeDecodeError:
            orgs = read_csv("latin-1")

        return orgs

    else:
        raise ValueError(f"Ukjent klientliste-format: {ext}")


def get_namespace():
    """Hent Outlook MAPI-namespace."""
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def get_root_folder(ns):
    """Returner passende rotmappe (egen eller delt postboks)."""
    if SHARED_MAILBOX_DISPLAYNAME.strip():
        # Delt postboks ved visningsnavn/epost
        return ns.Folders[SHARED_MAILBOX_DISPLAYNAME]
    return ns.GetDefaultFolder(6).Parent  # Parent av standard Innboks = rot av din postboks

def walk_all_mail_folders(root_folder):
    """Iterer over alle mail-mapper rekursivt, returner liste med mappe-objekter som har Items."""
    folders = []

    def rec(f):
        # plukk kun mapper som sannsynligvis har meldinger (de fleste gjør)
        folders.append(f)
        try:
            for i in range(1, f.Folders.Count + 1):
                rec(f.Folders.Item(i))
        except Exception:
            pass

    rec(root_folder)
    return folders

def robust_received_filter():
    """Outlook Restrict krever US-format med AM/PM."""
    since = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    return since, since.strftime("%m/%d/%Y %I:%M %p")

def get_smtp_address(msg) -> str:
    """Returner avsenders SMTP-adresse robust (EX/SMTP/PropertyAccessor)."""
    try:
        if getattr(msg, "SenderEmailType", "") == "EX":
            exu = msg.Sender.GetExchangeUser()
            if exu:
                return (exu.PrimarySmtpAddress or "").lower()
        smtp = (getattr(msg, "SenderEmailAddress", "") or "").lower()
        if smtp:
            return smtp
    except Exception:
        pass
    try:
        pa = msg.PropertyAccessor
        prop = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"  # PR_SMTP_ADDRESS
        return (pa.GetProperty(prop) or "").lower()
    except Exception:
        return ""

def fetch_messages_from_sender(sender_email: str):
    """Hent meldinger siste LOOKBACK_DAYS fra alle mapper, filtrer på avsender i Python."""
    sender_email = sender_email.lower()
    ns = get_namespace()
    root = get_root_folder(ns)

    folders = walk_all_mail_folders(root)
    since, since_str = robust_received_filter()

    matched = []
    for f in folders:
        try:
            items = f.Items
            items.Sort("[ReceivedTime]", True)
            try:
                # tidsfilter først, det er raskt
                candidates = items.Restrict(f"[ReceivedTime] >= '{since_str}'")
            except Exception:
                # fallback uten Restrict
                candidates = [m for m in items if getattr(m, "ReceivedTime", datetime.min) >= since]

            for m in candidates:
                try:
                    smtp = get_smtp_address(m)
                    if smtp == sender_email:
                        matched.append(m)
                except Exception:
                    continue
        except Exception:
            # Noen mapper (kalender/oppgaver) kan feile – hopp over
            continue
    return matched

def save_excel_attachments(msg, dest_dir: Path):
    """Lagre .xlsx/.xlsm-vedlegg. Returnerer liste med lagrede stier."""
    saved = []
    atts = getattr(msg, "Attachments", None)
    if not atts:
        return saved
    for i in range(1, atts.Count + 1):
        att = atts.Item(i)
        name = str(att.FileName)
        if name.lower().endswith((".xlsx", ".xlsm")):
            try:
                path = dest_dir / f"{msg.EntryID[:8]}_{name}"
                att.SaveAsFile(str(path))
                saved.append(path)
            except Exception as e:
                print(f"  ! Kunne ikke lagre vedlegg '{name}': {e}")
    return saved

def mark_processed(msg):
    if not SET_OUTLOOK_CATEGORY:
        return
    try:
        cats = [c.strip() for c in (msg.Categories or "").split(",") if c.strip()]
        if OUTLOOK_CATEGORY_NAME not in cats:
            cats.append(OUTLOOK_CATEGORY_NAME)
            msg.Categories = ", ".join(cats)
            msg.Save()
    except Exception:
        pass

# --------- Excel vedlegg: finn 'Unique ID' og les verdiene --------------
def find_unique_id_col(ws):
    """Returner (header_row_index, col_index) for kolonnen med header 'Unique ID'."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=50, values_only=True), start=1):
        if not any(row):
            continue
        for j, val in enumerate(row, start=1):
            if val is None:
                continue
            if str(val).strip().lower() == "unique id":
                return i, j
    return None, None

def extract_unique_ids(xlsx_path: Path):
    """Åpne Excel og plukk sifferverdier under kolonnen 'Unique ID' (første ark som har den)."""
    ids = []
    wb = load_workbook(filename=xlsx_path, data_only=True)
    for ws in wb.worksheets:
        header_row, col = find_unique_id_col(ws)
        if not col:
            continue
        r = header_row + 1
        empty_hits = 0
        while True:
            val = ws.cell(row=r, column=col).value
            if val is None or str(val).strip() == "":
                empty_hits += 1
                if empty_hits >= 3:
                    break
            else:
                empty_hits = 0
                ids.append(normalize_orgnr(val))
            r += 1
        if ids:
            break
    return [x for x in ids if x]
# ------------------------------------------------------------------------


def main():
    ensure_dir(DOWNLOAD_DIR)
    ensure_dir(RAPPORT_FIL.parent)

    # 1) Klientliste
    client_orgs = read_client_orgs(KLIENTLISTE)
    print(f"Klientliste: {len(client_orgs)} orgnr lastet fra {KLIENTLISTE}")

    # 2) Epost
    msgs = fetch_messages_from_sender(SENDER_EMAIL)
    print(f"Fant {len(msgs)} meldinger fra {SENDER_EMAIL} siste {LOOKBACK_DAYS} dager.")

    rows = []
    total_saved = 0

    for msg in msgs:
        subject = getattr(msg, "Subject", "")
        received = getattr(msg, "ReceivedTime", None)
        received_iso = (received.strftime("%Y-%m-%d %H:%M") if received else "")

        saved_files = save_excel_attachments(msg, DOWNLOAD_DIR)
        print(f" - '{subject}' → lagret {len(saved_files)} vedlegg")
        total_saved += len(saved_files)

        matches_found = 0
        for f in saved_files:
            ids = extract_unique_ids(f)
            for org in ids:
                hit = "JA" if org in client_orgs else "NEI"
                if hit == "JA":
                    matches_found += 1
                rows.append({
                    "received": received_iso,
                    "subject": subject,
                    "attachment": f.name,
                    "orgnr": org,
                    "match": hit
                })

        if matches_found > 0:
            mark_processed(msg)

    # 3) Rapport
    with RAPPORT_FIL.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["received", "subject", "attachment", "orgnr", "match"])
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nFerdig. {len(rows)} linjer skrevet til: {RAPPORT_FIL}")
    print(f"Vedlegg lagret: {total_saved} stk → {DOWNLOAD_DIR}")


if __name__ == "__main__":
    main()

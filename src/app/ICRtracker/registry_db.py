# -*- coding: utf-8 -*-
"""
registry_db.py
Adapter rundt SQLite for et stort aksjonærregister.

- Importer CSV -> SQLite (streaming)
- Oppslag: eiere av selskap, selskaper eid av eier, navn/orgnr-søk
- Laget for å brukes av matcher.py og tracker.py

Bruk:
  python -m ICRtracker.scan_registry --import-csv "aksjonarregister.csv" --db "aksjonarregister.db"

Avhengigheter: kun stdlib (sqlite3, csv). Ingen pandas nødvendig.
Valgfri fuzzy-match håndteres i matcher.py (rapidfuzz/difflib).
"""
from __future__ import annotations

import csv
import re
import sqlite3
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

# ------------------ normalisering ------------------

def normalize_orgnr(s: object) -> str:
    if s is None:
        return ""
    return re.sub(r"\D+", "", str(s))

def normalize_name(s: object) -> str:
    if s is None:
        return ""
    return " ".join(str(s).strip().split())

# ------------------ DB-grunnmur ------------------

def open_db(db_path: Path) -> sqlite3.Connection:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn

def create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS companies (
            orgnr TEXT PRIMARY KEY,
            name  TEXT
        );

        CREATE TABLE IF NOT EXISTS holdings (
            company_orgnr        TEXT NOT NULL,
            shareholder_type     TEXT,             -- 'person' | 'company' | NULL
            shareholder_orgnr    TEXT,             -- null for person uten orgnr
            shareholder_name     TEXT,
            shareholder_birthdate TEXT,
            stake_percent        REAL,
            shares               INTEGER,
            country              TEXT,
            city                 TEXT,
            postal_code          TEXT,
            share_class          TEXT,
            source_date          TEXT,
            PRIMARY KEY (company_orgnr, COALESCE(shareholder_orgnr, ''), COALESCE(shareholder_name, ''), COALESCE(share_class, ''))
        );

        CREATE INDEX IF NOT EXISTS idx_holdings_company ON holdings(company_orgnr);
        CREATE INDEX IF NOT EXISTS idx_holdings_owner_orgnr ON holdings(shareholder_orgnr);
        CREATE INDEX IF NOT EXISTS idx_holdings_owner_name ON holdings(shareholder_name);
        """
    )
    # FTS5 for raskt navnesøk (best effort)
    try:
        conn.execute("CREATE VIRTUAL TABLE IF NOT EXISTS names_fts USING fts5(name, orgnr, content='');")
    except sqlite3.OperationalError:
        # FTS5 ikke tilgjengelig i bygget: ignorer
        pass

# ------------------ Import CSV -> DB ------------------

# Vanlige kolonnenavn i rå CSV. Vi støtter flere varianter:
SYN = {
    "company_orgnr": [
        "selskap_orgnr", "orgnr_selskap", "selskapets orgnr", "foretak_orgnr",
        "company_orgnr", "company orgnr", "orgnr"
    ],
    "company_name": [
        "selskap_navn", "foretaksnavn", "company_name", "company", "navn", "selskap"
    ],
    "shareholder_orgnr": [
        "eier_orgnr", "aksjonær_orgnr", "aksjonaer_orgnr", "owner_orgnr",
        "orgnr_eier", "aksjonar_orgnr"
    ],
    "shareholder_name": [
        "eier_navn", "aksjonær_navn", "aksjonaer_navn", "owner_name", "navn_eier", "aksjonar_namn", "navn"
    ],
    "shareholder_birthdate": ["fodselsdato", "fødselsdato", "dob", "f_dato"],
    "stake_percent": [
        "eierandel_prosent", "eierandel %", "andel %", "eierandel", "andel_prosent", "prosent"
    ],
    "shares": ["antall_aksjer", "aksjer", "beholdning", "antall"],
    "country": ["land", "landkode", "country"],
    "city": ["poststed", "by", "city"],
    "postal_code": ["postnr", "postnummer", "zip"],
    "share_class": ["aksjeklasse", "class"],
    "source_date": ["dato", "per_dato", "regdato", "kildedato", "pr_dato"],
}

def _pick(header: List[str], candidates: List[str]) -> Optional[str]:
    low = {h.strip().lower(): h for h in header}
    for c in candidates:
        if c in low:
            return low[c]
    return None

def import_csv_to_db(csv_path: Path, db_path: Path, delimiter: str=",", batch: int=20000, encoding_try=("utf-8-sig","latin-1")) -> None:
    """
    Streamer stor CSV inn i SQLite. Lager companies/holdings + indekser.
    """
    conn = open_db(db_path)
    create_schema(conn)

    # Åpne med fallback på encoding
    last_err = None
    for enc in encoding_try:
        try:
            f = open(csv_path, "r", encoding=enc, newline="")
            break
        except Exception as e:
            last_err = e
            f = None
    if not f:
        raise last_err

    with f:
        reader = csv.DictReader(f, delimiter=delimiter)
        hdr = [h or "" for h in (reader.fieldnames or [])]
        if not hdr:
            raise ValueError("CSV mangler header.")

        # Bygg mapping fra vår kanon til faktiske feltnavn
        m: Dict[str, Optional[str]] = {key: _pick(hdr, SYN[key]) for key in SYN}
        # Krev minimum:
        if not m["company_orgnr"] or not (m["shareholder_orgnr"] or m["shareholder_name"]):
            raise ValueError("Fant ikke minimumskolonner i CSV (trenger selskap_orgnr og eier_orgnr/eller eier_navn).")

        # statements
        ins_co = "INSERT OR IGNORE INTO companies(orgnr, name) VALUES (?,?)"
        ins_h  = """
            INSERT OR REPLACE INTO holdings(
                company_orgnr, shareholder_type, shareholder_orgnr, shareholder_name,
                shareholder_birthdate, stake_percent, shares, country, city, postal_code, share_class, source_date
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """

        co_buf, h_buf = [], []
        n, n_err = 0, 0

        for row in reader:
            try:
                co_org = normalize_orgnr(row.get(m["company_orgnr"], ""))
                if not co_org:
                    continue
                co_name = normalize_name(row.get(m["company_name"], "")) if m["company_name"] else ""

                sh_org = normalize_orgnr(row.get(m["shareholder_orgnr"], "")) if m["shareholder_orgnr"] else ""
                sh_name = normalize_name(row.get(m["shareholder_name"], "")) if m["shareholder_name"] else ""
                sh_bd   = row.get(m["shareholder_birthdate"], "") if m["shareholder_birthdate"] else ""
                stake   = row.get(m["stake_percent"], "") if m["stake_percent"] else ""
                shares  = row.get(m["shares"], "") if m["shares"] else ""
                country = row.get(m["country"], "") if m["country"] else ""
                city    = row.get(m["city"], "") if m["city"] else ""
                post    = row.get(m["postal_code"], "") if m["postal_code"] else ""
                klass   = row.get(m["share_class"], "") if m["share_class"] else ""
                sdate   = row.get(m["source_date"], "") if m["source_date"] else ""

                shareholder_type = "company" if sh_org else ("person" if sh_name else None)

                # parse tall skånsomt
                def _float(x):
                    try:
                        return float(str(x).replace("%", "").replace(",", "."))
                    except Exception:
                        return None
                def _int(x):
                    try:
                        return int(str(x).replace(" ", "").replace(",", ""))
                    except Exception:
                        return None

                co_buf.append((co_org, co_name))
                h_buf.append((
                    co_org, shareholder_type, sh_org or None, sh_name or None,
                    sh_bd or None, _float(stake), _int(shares), country or None, city or None,
                    post or None, klass or None, sdate or None
                ))
                n += 1

                if len(h_buf) >= batch:
                    conn.executemany(ins_co, co_buf)
                    conn.executemany(ins_h,  h_buf)
                    conn.commit()
                    co_buf.clear(); h_buf.clear()

            except Exception:
                n_err += 1

        if h_buf:
            conn.executemany(ins_co, co_buf)
            conn.executemany(ins_h,  h_buf)
            conn.commit()

    # FTS5 fylles med selskapsnavn (best effort)
    try:
        conn.execute("DELETE FROM names_fts;")
        conn.execute("INSERT INTO names_fts(name, orgnr) SELECT name, orgnr FROM companies WHERE name IS NOT NULL;")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    conn.close()

# ------------------ Oppslag ------------------

def get_company_name(conn: sqlite3.Connection, orgnr: str) -> Optional[str]:
    org = normalize_orgnr(orgnr)
    cur = conn.execute("SELECT name FROM companies WHERE orgnr=?;", (org,))
    row = cur.fetchone()
    return row["name"] if row else None

def get_owners(conn: sqlite3.Connection, company_orgnr: str) -> List[sqlite3.Row]:
    org = normalize_orgnr(company_orgnr)
    cur = conn.execute(
        """SELECT * FROM holdings WHERE company_orgnr=? ORDER BY stake_percent DESC NULLS LAST""",
        (org,)
    )
    return cur.fetchall()

def companies_owned_by(conn: sqlite3.Connection, shareholder_orgnr: str) -> List[sqlite3.Row]:
    o = normalize_orgnr(shareholder_orgnr)
    cur = conn.execute(
        """SELECT h.*, c.name AS company_name
             FROM holdings h
             LEFT JOIN companies c ON c.orgnr=h.company_orgnr
            WHERE h.shareholder_orgnr=?""",
        (o,)
    )
    return cur.fetchall()

def search_name_candidates(conn: sqlite3.Connection, name: str, limit: int=50) -> List[Tuple[str, str]]:
    """Returner (navn, orgnr) kandidater fra FTS/LIKE for fuzzy ranking."""
    q = normalize_name(name)
    if not q:
        return []
    # FTS5 hvis tilgjengelig
    try:
        cur = conn.execute("SELECT name, orgnr FROM names_fts WHERE names_fts MATCH ? LIMIT ?;", (q, limit))
        rows = cur.fetchall()
        if rows:
            return [(r["name"], r["orgnr"]) for r in rows]
    except sqlite3.OperationalError:
        pass
    # Fallback: LIKE
    pat = f"%{q}%"
    cur = conn.execute("SELECT name, orgnr FROM companies WHERE name LIKE ? LIMIT ?;", (pat, limit))
    return [(r["name"], r["orgnr"]) for r in cur.fetchall()]

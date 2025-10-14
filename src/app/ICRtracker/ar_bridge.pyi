# -*- coding: utf-8 -*-
"""
ar_bridge.py
Bro mot eksisterende "aksjonaerregister"-moduler i repoet.
Bruker DuckDB-implementasjonen deres og tilpasser returverdier/keys til det
tracker/matcher forventer (samme felt som registry_db-varianten).

Forutsetter at modulene ligger under:  app/aksjonaerregister/
og at de eksporterer funksjoner som i db.py (ensure_db, open_conn, search_companies,
get_owners_full, get_children_agg_company). Se prosjektets db.py/app.py.
"""

from __future__ import annotations
from typing import List, Dict, Tuple, Optional
import re

# Prøv import med pakkeprefiks (src/app/aksjonaerregister)
try:
    from app.aksjonaerregister.db import (
        open_conn as _open_conn,
        search_companies as _search_companies,
        get_owners_full as _get_owners_full,
        get_children_agg_company as _get_children_agg_company,
    )
except Exception:
    # Fallback hvis prosjektet kjøres med annen sys.path:
    from aksjonaerregister.db import (  # type: ignore
        open_conn as _open_conn,
        search_companies as _search_companies,
        get_owners_full as _get_owners_full,
        get_children_agg_company as _get_children_agg_company,
    )

# ---------- Normalisering ----------
def normalize_orgnr(val) -> str:
    if val is None:
        return ""
    return re.sub(r"\D+", "", str(val))

def normalize_name(val) -> str:
    if val is None:
        return ""
    return " ".join(str(val).strip().split())

# ---------- Åpne DB ----------
def open_db(_db_path_ignored=None):
    """
    Åpner DuckDB-tilkoblingen via eksisterende modul.
    _db_path_ignored er bare for signaturkompabilitet (tracker/scan sender inn sti).
    Stien styres allerede av aksjonaerregister.settings.
    """
    return _open_conn()

# ---------- Søk (navn) ----------
def search_name_candidates(conn, name: str, limit: int = 50) -> List[Tuple[str, str]]:
    """
    Returner [(navn, orgnr)] for kandidater. Bruker eksisterende search_companies(by='navn').
    """
    term = normalize_name(name)
    if not term:
        return []
    rows = _search_companies(conn, term, by="navn", limit=limit)  # -> list of (orgnr, name)
    return [(r[1], str(r[0])) for r in rows if r and len(r) >= 2]

# ---------- Hent selskapsnavn ----------
def get_company_name(conn, orgnr: str) -> Optional[str]:
    """Hent selskapsnavn fra shareholders for gitt orgnr."""
    try:
        q = """
        SELECT company_name
        FROM shareholders
        WHERE company_orgnr = ?
        LIMIT 1
        """
        row = conn.execute(q, [normalize_orgnr(orgnr)]).fetchone()
        return None if row is None else (row[0] if isinstance(row, (list, tuple)) else row["company_name"])
    except Exception:
        return None

# ---------- Eiere av selskap ----------
def get_owners(conn, company_orgnr: str) -> List[Dict[str, object]]:
    """
    Tilpasser resultatet fra get_owners_full(...) til feltene tracker/matcher bruker:
      shareholder_orgnr, shareholder_name, shareholder_type, stake_percent, shares, company_orgnr
    """
    rows = _get_owners_full(conn, normalize_orgnr(company_orgnr))
    out: List[Dict[str, object]] = []
    for r in rows:
        # r-orden i db.py: owner_orgnr, owner_name, share_class, owner_country, owner_zip_place,
        #                  shares_owner_num, shares_company_num, ownership_pct  (alle kan være None)
        owner_orgnr   = str(r[0]) if r[0] is not None else ""
        owner_name    = r[1] or ""
        shares_owner  = r[5]
        pct           = r[7]
        sh_type       = "company" if owner_orgnr else "person"
        out.append({
            "company_orgnr": normalize_orgnr(company_orgnr),
            "shareholder_orgnr": normalize_orgnr(owner_orgnr),
            "shareholder_name": normalize_name(owner_name),
            "shareholder_type": sh_type,
            "stake_percent": pct,
            "shares": shares_owner,
        })
    return out

# ---------- Selskaper eiet av gitt orgnr ----------
def companies_owned_by(conn, shareholder_orgnr: str) -> List[Dict[str, object]]:
    """
    Tilpasser resultatet fra get_children_agg_company(...) til feltene tracker/matcher bruker:
      company_orgnr, company_name, stake_percent, shares
    """
    rows = _get_children_agg_company(conn, normalize_orgnr(shareholder_orgnr))
    out: List[Dict[str, object]] = []
    for r in rows:
        # r-orden i db.py: company_orgnr, company_name, shares_owner_num, shares_company_num, ownership_pct
        co_org   = str(r[0]) if r[0] is not None else ""
        co_name  = r[1] or ""
        shares   = r[2]
        pct      = r[4]
        out.append({
            "company_orgnr": normalize_orgnr(co_org),
            "company_name": normalize_name(co_name),
            "stake_percent": pct,
            "shares": shares,
        })
    return out

# -*- coding: utf-8 -*-
"""
ar_bridge.py
----------------
Bro mot den eksisterende aksjonærregister-implementasjonen (DuckDB) i
`app.aksjonærregister.db`. Denne filen oversetter dataene til samme
feltnavn/format som ICR-tracker/matcher forventer, slik at resten av
koden kan være uendret.

Forventet prosjektstruktur:
  src/
    app/
      __init__.py
      aksjonærregister/        <-- legg merke til 'æ'
        __init__.py
        db.py                  <-- har open_conn, search_companies, get_owners_full, get_children_agg_company
      ICRtracker/
        ar_bridge.py (denne fila)
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

# ---- Importer AR-modulen (med korrekt mappenavn 'aksjonærregister').
#      Vi beholder en fallback til 'aksjonaerregister' hvis noen maskiner har ascii-navn. ----
try:
    from app.aksjonærregister.db import (  # type: ignore
        open_conn as _open_conn,
        search_companies as _search_companies,
        get_owners_full as _get_owners_full,
        get_children_agg_company as _get_children_agg_company,
    )
except Exception:
    # Fallback til ascii-mappe ('aksjonaerregister') dersom det navnet brukes i repoet på en annen maskin
    from app.aksjonaerregister.db import (  # type: ignore
        open_conn as _open_conn,
        search_companies as _search_companies,
        get_owners_full as _get_owners_full,
        get_children_agg_company as _get_children_agg_company,
    )


# ------------------------- Normalisering -------------------------

def normalize_orgnr(val) -> str:
    """Til kun sifre (fjerner NO-, mellomrom, tegn osv.)."""
    if val is None:
        return ""
    return re.sub(r"\D+", "", str(val))


def normalize_name(val) -> str:
    """Trim + kollaps mellomrom."""
    if val is None:
        return ""
    return " ".join(str(val).strip().split())


# ------------------------- DB-tilkobling -------------------------

def open_db(_db_path_ignored=None):
    """
    Signaturkompatibel med registry_db.open_db, men bruker AR-modulens open_conn.
    _db_path_ignored: beholdes kun for kompatibilitet; stien styres av AR-modulens settings.
    """
    return _open_conn()


# --------------------------- Oppslag ----------------------------

def search_name_candidates(conn, name: str, limit: int = 50) -> List[Tuple[str, str]]:
    """
    Navnesøk mot AR: returner [(navn, orgnr)] for kandidater.
    Bruker eksisterende search_companies(by='navn'), som returnerer [(company_orgnr, company_name)].
    """
    term = normalize_name(name)
    if not term:
        return []
    rows = _search_companies(conn, term, by="navn", limit=limit)
    # oversett til [(navn, orgnr)]
    out: List[Tuple[str, str]] = []
    for r in rows:
        org = str(r[0]) if r and len(r) > 0 and r[0] is not None else ""
        nam = str(r[1]) if r and len(r) > 1 and r[1] is not None else ""
        if org and nam:
            out.append((nam, org))
    return out


def get_company_name(conn, orgnr: str) -> Optional[str]:
    """
    Hent selskapsnavn for orgnr fra shareholders-tabellen (første treff).
    """
    try:
        row = conn.execute(
            "SELECT company_name FROM shareholders WHERE company_orgnr=? LIMIT 1",
            [normalize_orgnr(orgnr)],
        ).fetchone()
        if row is None:
            return None
        # DuckDB kan returnere tuple/Row – støtt begge
        return row[0] if not isinstance(row, dict) else row.get("company_name")
    except Exception:
        return None


def get_owners(conn, company_orgnr: str) -> List[Dict[str, object]]:
    """
    Eiere av gitt selskap – tilpasset til feltnavn ICR-tracker/matcher forventer.
    Bruker get_owners_full(...) fra AR-modulen.

    Forventet rekkefølge fra AR (se db.py):
      owner_orgnr, owner_name, share_class, owner_country, owner_zip_place,
      shares_owner_num, shares_company_num, ownership_pct
    """
    rows = _get_owners_full(conn, normalize_orgnr(company_orgnr))
    out: List[Dict[str, object]] = []
    for r in rows:
        owner_orgnr  = str(r[0]) if r and len(r) > 0 and r[0] is not None else ""
        owner_name   = r[1] if r and len(r) > 1 and r[1] is not None else ""
        shares_owner = r[5] if r and len(r) > 5 else None
        pct          = r[7] if r and len(r) > 7 else None
        sh_type      = "company" if owner_orgnr else "person"
        out.append({
            "company_orgnr":       normalize_orgnr(company_orgnr),
            "shareholder_orgnr":   normalize_orgnr(owner_orgnr),
            "shareholder_name":    normalize_name(owner_name),
            "shareholder_type":    sh_type,
            "stake_percent":       pct,
            "shares":              shares_owner,
        })
    return out


def companies_owned_by(conn, shareholder_orgnr: str) -> List[Dict[str, object]]:
    """
    Selskaper som eies (helt/delvis) av gitt orgnr – felter matcher matcher.py/tracker.py:
      company_orgnr, company_name, stake_percent, shares

    Bruker get_children_agg_company(...) fra AR-modulen.

    Forventet rekkefølge fra AR (se db.py):
      company_orgnr, company_name, shares_owner_num, shares_company_num, ownership_pct
    """
    rows = _get_children_agg_company(conn, normalize_orgnr(shareholder_orgnr))
    out: List[Dict[str, object]] = []
    for r in rows:
        co_org  = str(r[0]) if r and len(r) > 0 and r[0] is not None else ""
        co_name = r[1] if r and len(r) > 1 and r[1] is not None else ""
        shares  = r[2] if r and len(r) > 2 else None
        pct     = r[4] if r and len(r) > 4 else None
        out.append({
            "company_orgnr": normalize_orgnr(co_org),
            "company_name":  normalize_name(co_name),
            "stake_percent": pct,
            "shares":        shares,
        })
    return out

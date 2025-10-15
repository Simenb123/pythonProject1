# -*- coding: utf-8 -*-
"""
ar_bridge.py
-------------
Robust bro mot aksjonærregister-modulen i prosjektet.

Gir et stabilt API uansett om modulen heter:
  - app.aksjonærregister (med 'æ'), eller
  - app.aksjonaerregister (ascii-fallback)

Og uansett om den eksporterer:
  - get_owners_full / get_owners
  - companies_owned_by / get_children_agg_company / get_children_for_company
  - open_conn(db_path?) / open_conn()

Offentlig API i denne filen:
  - open_db(db_path: Optional[pathlib.Path]) -> "conn"
  - get_owners(conn, orgnr: str) -> List[Dict]
  - companies_owned_by(conn, orgnr: str) -> List[Dict]
  - normalize_orgnr(s: str) -> str
"""

from __future__ import annotations

import sys
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

# ---- Finn <prosjektrot>/src slik at "app" kan importeres når filen kjøres isolert ----
if __name__ == "__main__" and __package__ is None:
    here = Path(__file__).resolve()
    for parent in (here.parent, *here.parents):
        if (parent / "app" / "__init__.py").exists():
            sys.path.insert(0, str(parent))
            break

# ---- Importer aksjonærregister-modul – både æ/ae-varianter støttes ----
_ar_mod = None
_ar_err: Optional[BaseException] = None
for dotted in ("app.aksjonærregister.db", "app.aksjonaerregister.db", "app.aksjonærregister", "app.aksjonaerregister"):
    try:
        _ar_mod = __import__(dotted, fromlist=["*"])
        break
    except Exception as e:
        _ar_err = e
        _ar_mod = None

if _ar_mod is None:
    # Vi lar import-feilen boble først når noen kaller open_db/get_owners/companies_owned_by
    pass


# ----------------------------- Hjelpefunksjoner -----------------------------

def normalize_orgnr(s: Any) -> str:
    """Behold kun siffer (eller 4-sifret år/person-ID når det er det som brukes)."""
    if s is None:
        return ""
    return re.sub(r"\D+", "", str(s))


def _ensure_ar_loaded():
    if _ar_mod is None:
        raise RuntimeError(
            "Fant ikke aksjonærregister-modulen. "
            "Sjekk at 'src/app/aksjonærregister' (eller 'aksjonaerregister') finnes og kan importeres. "
            f"Original importfeil: {_ar_err!r}"
        )


def _as_float(x: Any) -> Optional[float]:
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None


def _as_int(x: Any) -> Optional[int]:
    try:
        if x is None or x == "":
            return None
        return int(x)
    except Exception:
        return None


# ----------------------------- Offentlig API -----------------------------

def open_db(db_path: Optional[Path] = None):
    """
    Åpner en tilkobling til aksjonærregisteret via modulens 'open_conn'.
    Støtter både signatur uten og med sti-parameter.
    Returnerer et "conn"-objekt (duckdb-tilkobling el.l.).
    """
    _ensure_ar_loaded()

    # flertydig API: noen varianter eksporterer underpakke "db" (med open_conn), noen på toppnivå
    ar = _ar_mod
    if hasattr(_ar_mod, "db"):
        ar = _ar_mod.db  # type: ignore[attr-defined]

    if not hasattr(ar, "open_conn"):
        raise RuntimeError("Aksjonærregister-modulen mangler funksjonen 'open_conn'.")

    open_conn = getattr(ar, "open_conn")
    try:
        # prøv med path-parameter
        return open_conn(db_path)
    except TypeError:
        # prøv uten path-parameter
        return open_conn()


def get_owners(conn, company_orgnr: str) -> List[Dict[str, Any]]:
    """
    Henter direkte eiere (aksjonærer) for et gitt selskap.
    Returnerer liste av dict med normaliserte felt:
      shareholder_orgnr, shareholder_name, shareholder_type, stake_percent, shares
    """
    _ensure_ar_loaded()
    ar = _ar_mod
    if hasattr(_ar_mod, "db"):
        ar = _ar_mod.db  # type: ignore[attr-defined]

    org = normalize_orgnr(company_orgnr)

    rows: List[Dict[str, Any]] = []

    # 1) Prioritér "rike" API-er først
    if hasattr(ar, "get_owners_full"):
        raw = ar.get_owners_full(conn, org)
        for r in raw:
            rows.append({
                "shareholder_orgnr": normalize_orgnr(r.get("owner_orgnr") or r.get("shareholder_orgnr") or r.get("orgnr") or ""),
                "shareholder_name": r.get("owner_name") or r.get("shareholder_name") or r.get("name") or "",
                "shareholder_type": r.get("owner_type") or r.get("shareholder_type") or r.get("type") or "",
                "stake_percent": _as_float(r.get("stake_percent") or r.get("percent") or r.get("ownership_pct")),
                "shares": _as_int(r.get("shares") or r.get("share_count")),
            })
        return rows

    # 2) Vanlig "get_owners"
    if hasattr(ar, "get_owners"):
        raw = ar.get_owners(conn, org)
        for r in raw:
            rows.append({
                "shareholder_orgnr": normalize_orgnr(r.get("owner_orgnr") or r.get("shareholder_orgnr") or r.get("orgnr") or ""),
                "shareholder_name": r.get("owner_name") or r.get("shareholder_name") or r.get("name") or "",
                "shareholder_type": r.get("owner_type") or r.get("shareholder_type") or r.get("type") or "",
                "stake_percent": _as_float(r.get("stake_percent") or r.get("percent") or r.get("ownership_pct")),
                "shares": _as_int(r.get("shares") or r.get("share_count")),
            })
        return rows

    # 3) Fallback: prøv generisk SQL (tabellnavn kan variere; dette er kun nødløsning)
    try:
        # Eksempelnavn: holdings / shareholders
        q = """
        SELECT
          COALESCE(owner_orgnr, shareholder_orgnr, '') AS owner_id,
          COALESCE(owner_name, shareholder_name, name, '') AS owner_name,
          COALESCE(owner_type, shareholder_type, type, '') AS owner_type,
          COALESCE(stake_percent, percent, ownership_pct, NULL) AS pct,
          COALESCE(shares, share_count, NULL) AS shares
        FROM holdings
        WHERE normalize_orgnr(company_orgnr) = ?
        """
        cur = conn.execute(q, [org])  # type: ignore[attr-defined]
        for owner_id, owner_name, owner_type, pct, shares in cur.fetchall():
            rows.append({
                "shareholder_orgnr": normalize_orgnr(owner_id),
                "shareholder_name": owner_name or "",
                "shareholder_type": owner_type or "",
                "stake_percent": _as_float(pct),
                "shares": _as_int(shares),
            })
        return rows
    except Exception:
        # siste utvei: tom
        return []


def companies_owned_by(conn, company_orgnr: str) -> List[Dict[str, Any]]:
    """
    Henter direkte eide selskaper for et gitt selskap.
    Returnerer liste av dict med normaliserte felt:
      company_orgnr, company_name, stake_percent, shares
    """
    _ensure_ar_loaded()
    ar = _ar_mod
    if hasattr(_ar_mod, "db"):
        ar = _ar_mod.db  # type: ignore[attr-defined]

    org = normalize_orgnr(company_orgnr)
    rows: List[Dict[str, Any]] = []

    # 1) Prefererte API-er
    if hasattr(ar, "companies_owned_by"):
        raw = ar.companies_owned_by(conn, org)
        for r in raw:
            rows.append({
                "company_orgnr": normalize_orgnr(r.get("company_orgnr") or r.get("orgnr") or ""),
                "company_name": r.get("company_name") or r.get("name") or "",
                "stake_percent": _as_float(r.get("stake_percent") or r.get("percent") or r.get("ownership_pct")),
                "shares": _as_int(r.get("shares") or r.get("share_count")),
            })
        return rows

    if hasattr(ar, "get_children_agg_company"):
        raw = ar.get_children_agg_company(conn, org)
        for r in raw:
            rows.append({
                "company_orgnr": normalize_orgnr(r.get("company_orgnr") or r.get("orgnr") or ""),
                "company_name": r.get("company_name") or r.get("name") or "",
                "stake_percent": _as_float(r.get("stake_percent") or r.get("percent") or r.get("ownership_pct")),
                "shares": _as_int(r.get("shares") or r.get("share_count")),
            })
        return rows

    if hasattr(ar, "get_children_for_company"):
        raw = ar.get_children_for_company(conn, org)
        for r in raw:
            rows.append({
                "company_orgnr": normalize_orgnr(r.get("company_orgnr") or r.get("orgnr") or ""),
                "company_name": r.get("company_name") or r.get("name") or "",
                "stake_percent": _as_float(r.get("stake_percent") or r.get("percent") or r.get("ownership_pct")),
                "shares": _as_int(r.get("shares") or r.get("share_count")),
            })
        return rows

    # 2) Fallback SQL (best-effort)
    try:
        q = """
        SELECT
          COALESCE(child_orgnr, company_orgnr, orgnr, '') AS c_org,
          COALESCE(child_name, company_name, name, '') AS c_name,
          COALESCE(stake_percent, percent, ownership_pct, NULL) AS pct,
          COALESCE(shares, share_count, NULL) AS shares
        FROM holdings
        WHERE normalize_orgnr(owner_orgnr) = ?
        """
        cur = conn.execute(q, [org])  # type: ignore[attr-defined]
        for c_org, c_name, pct, shares in cur.fetchall():
            rows.append({
                "company_orgnr": normalize_orgnr(c_org),
                "company_name": c_name or "",
                "stake_percent": _as_float(pct),
                "shares": _as_int(shares),
            })
        return rows
    except Exception:
        return []

# -*- coding: utf-8 -*-
"""
db_compat_adapter.py
Adapter for å bruke en eksisterende AR-SQLite med annet skjema (uten å bygge om DB).
Tilpass TABELL/KOLONNE-navnene i KONFIG-seksjonen så funksjonene matcher
signaturen som registry_db.py byr på:
  - open_db(db_path) -> sqlite3.Connection
  - get_owners(conn, company_orgnr) -> List[sqlite3.Row]
  - companies_owned_by(conn, shareholder_orgnr) -> List[sqlite3.Row]
  - normalize_orgnr(val) -> str
"""

from __future__ import annotations
import re
import sqlite3
from pathlib import Path
from typing import List

# ============ KONFIG: FYLL INN DINE TABELL/KOLONNE-NAVN ============

# Tabell: Selskaper
TABLE_COMPANIES = "Companies"        # f.eks. "Companies" / "company" / "selskap"
COL_CO_ORGNR    = "OrgNo"            # f.eks. "orgnr" / "OrgNr" / "OrgNo"
COL_CO_NAME     = "CompanyName"      # f.eks. "name" / "Navn" / "CompanyName"

# Tabell: Eierskap (relasjoner)
TABLE_HOLDINGS  = "Ownerships"       # f.eks. "Ownerships" / "holdings" / "Eierskap"
COL_H_CO_ORGNR  = "CompanyOrgNo"     # f.eks. "company_orgnr"
COL_H_SH_ORGNR  = "OwnerOrgNo"       # f.eks. "shareholder_orgnr" (kan være NULL for personer)
COL_H_SH_NAME   = "OwnerName"        # f.eks. "shareholder_name"
# Hvis databasen har eksplisitt eiertype (person/selskap), angi kolonnen:
COL_H_SH_TYPE   = None               # f.eks. "OwnerType" eller None for heuristikk
COL_H_STAKE     = "Percent"          # f.eks. "stake_percent" / "AndelProsent"
COL_H_SHARES    = "Shares"           # f.eks. "shares" / "AntallAksjer"

# (valgfri) felt hvis du har dem i DB-en:
COL_H_COUNTRY   = None               # f.eks. "Country"
COL_H_CITY      = None               # f.eks. "City"
COL_H_POSTAL    = None               # f.eks. "PostalCode"
COL_H_CLASS     = None               # f.eks. "ShareClass"
COL_H_SRC_DATE  = None               # f.eks. "SourceDate"

# ===================================================================

def normalize_orgnr(val) -> str:
    if val is None:
        return ""
    return re.sub(r"\D+", "", str(val))

def open_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn

def _select_or_null(col: str | None) -> str:
    return col if col else "NULL"

def get_owners(conn: sqlite3.Connection, company_orgnr: str) -> List[sqlite3.Row]:
    """
    Returner EIERE av selskapet (samme feltsett som registry_db.get_owners).
    Felt som ikke finnes i din DB returneres som NULL.
    """
    org = normalize_orgnr(company_orgnr)
    if COL_H_SH_TYPE:
        sh_type_sql = COL_H_SH_TYPE
    else:
        sh_type_sql = f"CASE WHEN {COL_H_SH_ORGNR} IS NOT NULL AND {COL_H_SH_ORGNR}!='' THEN 'company' ELSE 'person' END"

    sql = f"""
        SELECT
            {COL_H_CO_ORGNR}                               AS company_orgnr,
            {sh_type_sql}                                  AS shareholder_type,
            {COL_H_SH_ORGNR}                               AS shareholder_orgnr,
            {COL_H_SH_NAME}                                AS shareholder_name,
            NULL                                           AS shareholder_birthdate,
            {COL_H_STAKE}                                  AS stake_percent,
            {COL_H_SHARES}                                 AS shares,
            {_select_or_null(COL_H_COUNTRY)}               AS country,
            {_select_or_null(COL_H_CITY)}                  AS city,
            {_select_or_null(COL_H_POSTAL)}                AS postal_code,
            {_select_or_null(COL_H_CLASS)}                 AS share_class,
            {_select_or_null(COL_H_SRC_DATE)}              AS source_date
        FROM {TABLE_HOLDINGS}
        WHERE {COL_H_CO_ORGNR} = ?
        ORDER BY {COL_H_STAKE} DESC;
    """
    cur = conn.execute(sql, (org,))
    return cur.fetchall()

def companies_owned_by(conn: sqlite3.Connection, shareholder_orgnr: str) -> List[sqlite3.Row]:
    """
    Returner SELSKAPER som eies av gitt orgnr (samme feltnavn som registry_db.companies_owned_by).
    """
    o = normalize_orgnr(shareholder_orgnr)
    sql = f"""
        SELECT
            h.{COL_H_CO_ORGNR}     AS company_orgnr,
            c.{COL_CO_NAME}        AS company_name,
            h.{COL_H_STAKE}        AS stake_percent,
            h.{COL_H_SHARES}       AS shares
        FROM {TABLE_HOLDINGS} h
        LEFT JOIN {TABLE_COMPANIES} c ON c.{COL_CO_ORGNR} = h.{COL_H_CO_ORGNR}
        WHERE h.{COL_H_SH_ORGNR} = ?
        ORDER BY h.{COL_H_STAKE} DESC;
    """
    cur = conn.execute(sql, (o,))
    return cur.fetchall()

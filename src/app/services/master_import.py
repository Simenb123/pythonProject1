# -*- coding: utf-8 -*-
# src/app/services/master_import.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple, Dict
import re
import pandas as pd

from app.services.registry import load_registry, ExcelRegistry, digits_only
from app.services.clients import list_clients  # for map mappenavn <-> klientnr  :contentReference[oaicite:3]{index=3}

# --------------------------- kolonne-synonymer ---------------------------
_SYNONYMS: Dict[str, list[str]] = {
    "client": ["client","klient","klientnavn","navn"],
    "orgnr": ["orgnr","organisasjonsnummer","org nr","orgnummer"],
    "partner": ["partner","ansvarlig partner","ansvarlig"],
    "bransjekode": ["bransjekode","naeringskode","næringskode"],
    "bransjekodenavn": ["bransjekodenavn","næringskodenavn","naeringskodenavn"],
    "selskapsform": ["selskapsform","juridisk form"],
    "klientnummer": ["klientnummer","klientnr","klient nr","kundennr","kundenummer","clientno"],
    "industry": ["industry","bransje"],
    "contact": ["contact","kontakt","kontaktperson"],
    "email": ["email","epost","e-post"],
    "phone": ["phone","telefon","tlf"],
    "address": ["address","adresse"],
    "fiscal_year_end": ["fiscal_year_end","årsslutt","year_end","fye"],
    "notes": ["notes","merknader","kommentar"],
}
_STD_KEYS = list(_SYNONYMS.keys())


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    low = {c.strip().lower(): c for c in df.columns}
    lut: Dict[str, str] = {}
    for std, cands in _SYNONYMS.items():
        for cand in cands:
            src = low.get(cand.lower())
            if src:
                lut[src] = std
                break
    out = df.rename(columns=lut).copy()
    for k in _STD_KEYS:
        if k not in out.columns:
            out[k] = pd.NA
    return out[_STD_KEYS]


def load_master_file(p: Path) -> pd.DataFrame:
    suf = str(p).lower()
    if suf.endswith((".xlsx",".xls")):
        df = pd.read_excel(p, engine="openpyxl")
    elif suf.endswith(".csv"):
        df = pd.read_csv(p)
    else:
        raise ValueError("Støtter kun .xlsx/.xls/.csv")
    return _normalize_columns(df)


@dataclass
class FieldChange:
    key: str
    old: str
    new: str
    accept: bool = False


@dataclass
class ClientChange:
    client: str           # NØKKELEN i registry (mappenavn når vi finner match)
    orgnr: str
    proposed_name: str
    fields: List[FieldChange]


def _to_str(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def _folder_clientnr_map(root: Path) -> Dict[str, str]:
    """Bygg LUT fra klientnr (ledende tall) → mappenavn ('3171 Foo AS' -> {'3171': '3171 Foo AS'})."""
    out: Dict[str, str] = {}
    for name in list_clients(root):  # :contentReference[oaicite:4]{index=4}
        m = re.match(r"^\s*(\d{2,})\b", name)
        if m:
            out[m.group(1)] = name
    return out


def diff_against_registry(root: Path, df_imp: pd.DataFrame) -> Tuple[ExcelRegistry, List[ClientChange]]:
    reg = load_registry(root)
    out: List[ClientChange] = []

    # precompute LUT-er
    cnr_to_folder = _folder_clientnr_map(root)
    folders = set(list_clients(root))  # for «endswith»-match på navn  :contentReference[oaicite:5]{index=5}

    for _, row in df_imp.iterrows():
        orgnr = digits_only(_to_str(row.get("orgnr")))
        klnr  = digits_only(_to_str(row.get("klientnummer")))
        proposed_name = _to_str(row.get("client"))

        # 1) Finn registry-key («client» i registry): orgnr → registry, ellers klientnr → mappenavn, ellers navnefallback
        client = reg.find_client_by_orgnr(orgnr) if orgnr else None
        if not client and klnr and klnr in cnr_to_folder:
            client = cnr_to_folder[klnr]
        if not client and proposed_name:
            # forsøk navnematch på mappe-navn (håndterer «3171 Foo AS»)
            pn = proposed_name.lower()
            for f in folders:
                fn = f.lower()
                if fn == pn or fn.endswith(" " + pn):
                    client = f
                    break
        if not client:
            # siste utvei – vi lager rad under foreslått navn (kan senere kobles når orgnr settes)
            client = proposed_name or f"(ukjent_{orgnr or klnr or 'NA'})"

        # 2) Samle feltvise endringer
        current = reg.get_client_info(client)
        changes: List[FieldChange] = []
        for k in _STD_KEYS:
            new = _to_str(row.get(k))
            if k == "orgnr":
                new = orgnr or ""
            old = _to_str(current.get(k, ""))
            if new != "" and new != old:
                changes.append(FieldChange(key=k, old=old, new=new, accept=False))

        if changes:
            out.append(ClientChange(client=client, orgnr=orgnr or "", proposed_name=proposed_name, fields=changes))

    return reg, out


def apply_changes(root: Path, reg: ExcelRegistry, items: List[ClientChange], by_user: str) -> int:
    n = 0
    for it in items:
        info = reg.get_client_info(it.client)
        info.setdefault("client", it.client)

        for fc in it.fields:
            if fc.accept:
                info[fc.key] = fc.new
                n += 1

        # sikre orgnr når vi har det
        if it.orgnr and digits_only(info.get("orgnr")) != digits_only(it.orgnr):
            info["orgnr"] = it.orgnr
            n += 1

        reg.upsert_client_info(info)

    reg.save()
    return n

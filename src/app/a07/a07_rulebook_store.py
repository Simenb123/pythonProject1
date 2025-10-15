# a07_rulebook_store.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional

from a07_models import (
    A07CodeDef,
    parse_account_ranges,
    ranges_to_spec,
    normalize_text,
)

DEFAULT_RULEBOOK_DIR = Path(r"F:\Dokument\Kildefiler\a07")
DEFAULT_RULEBOOK_FILE = DEFAULT_RULEBOOK_DIR / "global_a07_rulebook.json"


def _ensure_rulebook_dir() -> None:
    DEFAULT_RULEBOOK_DIR.mkdir(parents=True, exist_ok=True)


def empty_rulebook() -> dict:
    """Tomt skjelett for regelbok."""
    return {
        "version": 1,
        "updated": datetime.now().isoformat(timespec="seconds"),
        "codes": {},    # code -> dict(...)
        "groups": {},   # group_id -> {"name": str, "codes": [code,...], "expected_sign": Optional[int]}
    }


def load_rulebook(path: Path | str = DEFAULT_RULEBOOK_FILE) -> dict:
    """Leser regelbok. Oppretter tom hvis den ikke finnes."""
    _ensure_rulebook_dir()
    p = Path(path)
    if not p.exists():
        rb = empty_rulebook()
        save_rulebook(rb, p)
        return rb
    with p.open("r", encoding="utf-8") as f:
        data = json.load(f)
    # Sanitér minimumsfelter
    data.setdefault("version", 1)
    data.setdefault("codes", {})
    data.setdefault("groups", {})
    return data


def save_rulebook(rulebook: dict, path: Path | str = DEFAULT_RULEBOOK_FILE) -> None:
    """Lagrer regelbok med oppdatert timestamp."""
    rulebook = dict(rulebook)  # shallow copy
    rulebook["updated"] = datetime.now().isoformat(timespec="seconds")
    p = Path(path)
    _ensure_rulebook_dir()
    with p.open("w", encoding="utf-8") as f:
        json.dump(rulebook, f, ensure_ascii=False, indent=2)


def codes_from_rulebook(rulebook: dict) -> Dict[str, A07CodeDef]:
    """Konverterer RB->A07CodeDef-objekter."""
    out: Dict[str, A07CodeDef] = {}
    for code, raw in (rulebook.get("codes") or {}).items():
        out[code] = A07CodeDef(
            code=code,
            name=raw.get("name") or code,
            account_ranges=[tuple(x) for x in raw.get("account_ranges") or []],
            expected_sign=raw.get("expected_sign"),
            aliases=list(raw.get("aliases") or []),
            keywords=list(raw.get("keywords") or []),
        )
    return out


def upsert_code(rulebook: dict, code_def: A07CodeDef, merge: bool = True) -> None:
    """
    Legger inn/oppdaterer én A07-kode i regelboken.
    - merge=True: eksisterende alias/keywords bevares og utvides
    """
    codes = rulebook.setdefault("codes", {})
    existing = codes.get(code_def.code)
    if not existing or not merge:
        # full overwrite
        codes[code_def.code] = {
            "name": code_def.name,
            "account_ranges": [(int(a), int(b)) for (a, b) in code_def.account_ranges],
            "expected_sign": code_def.expected_sign,
            "aliases": list(code_def.aliases or []),
            "keywords": list(code_def.keywords or []),
        }
        return

    # Merge-flyt
    merged = dict(existing)
    merged["name"] = code_def.name or existing.get("name") or code_def.code

    # ranges
    old_ranges: List[Tuple[int, int]] = [tuple(x) for x in existing.get("account_ranges") or []]
    new_ranges = list(old_ranges)
    for a, b in code_def.account_ranges:
        if (a, b) not in new_ranges:
            new_ranges.append((int(a), int(b)))
    merged["account_ranges"] = new_ranges

    # expected sign
    if code_def.expected_sign is not None:
        merged["expected_sign"] = int(code_def.expected_sign)

    # alias/keywords – unike og normaliserte
    def _merge_list(old: List[str], add: List[str]) -> List[str]:
        out: List[str] = []
        seen = set()
        for s in list(old or []) + list(add or []):
            s2 = normalize_text(s)
            if s2 and s2 not in seen:
                out.append(s)
                seen.add(s2)
        return out

    merged["aliases"] = _merge_list(existing.get("aliases") or [], code_def.aliases or [])
    merged["keywords"] = _merge_list(existing.get("keywords") or [], code_def.keywords or [])
    codes[code_def.code] = merged


def remove_code(rulebook: dict, code: str) -> None:
    codes = rulebook.get("codes") or {}
    if code in codes:
        del codes[code]


def upsert_group(rulebook: dict, group_id: str, name: str, codes: List[str],
                 expected_sign: Optional[int] = None) -> None:
    groups = rulebook.setdefault("groups", {})
    groups[group_id] = {
        "name": name,
        "codes": list(dict.fromkeys(codes)),
        "expected_sign": expected_sign,
    }


def remove_group(rulebook: dict, group_id: str) -> None:
    groups = rulebook.get("groups") or {}
    if group_id in groups:
        del groups[group_id]

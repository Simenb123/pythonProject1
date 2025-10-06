# -*- coding: utf-8 -*-
"""
matcher_fallback.py
-------------------
Konservativ fallback-matcher for konto->A07-kode når regelbok ikke er lastet.
- Gater til lønnsrelevante kontointervaller (5xxx, 70xx–73xx) + {2940, 5290}
- Ban-ord for å hindre bank/fordringer/inntekter/mva/aksjer
- Min-score (fra GUI), tegn-samsvar og størrelses-sanity må passere
"""

from __future__ import annotations
from typing import Any, Dict, List, Set
import re

# ---------- Normalisering og enkle tokenhjelpere ----------

def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    return s.replace("ø", "oe").replace("å", "aa").replace("æ", "ae")

def _split_camel_snake(s: str) -> List[str]:
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    s = s.replace("_", " ")
    return [t for t in re.split(r"[^a-zA-Z0-9]+", s) if t]

def _tokens_from_text(text: str) -> Set[str]:
    txt = _norm(text)
    txt = re.sub(r"[^a-z0-9]+", " ", txt)
    parts = [p for p in txt.split() if p]
    base = set(parts)
    replace = {
        "lonn": "loenn", "lon": "loenn",
        "feriepenger": "ferie", "fp": "ferie",
        "reiseutg": "reise", "reisekost": "reise",
        "kilometergodtgjoerelse": "kilometer", "bilgodtgjoerelse": "bil",
        "aga": "arbeidsgiveravgift",
    }
    mapped = set(replace.get(p, p) for p in base)
    if "bil" in mapped: mapped |= {"kilometer","km","godtgjorelse"}
    if "reise" in mapped: mapped |= {"diett","kost","overnatting"}
    if "trekk" in mapped: mapped |= {"trekkilonn","trekk_i_lonn"}
    return mapped

def _tokens_from_code(code: str) -> Set[str]:
    toks = [_norm(t) for t in _split_camel_snake(code)]
    base = set(toks)
    expand = set()
    for t in list(base):
        if t in {"loenn","lonn","lon"}: expand |= {"loenn","lonn","lon"}
        if "ferie" in t: expand |= {"ferie","feriepenger","fp"}
        if t in {"reise","diett","kost","overnatting"} or "reise" in t or "kost" in t:
            expand |= {"reise","reisekost","diett","overnatting","km","kilometer","bil"}
        if "km" in t or "kilo" in t or "bil" in t: expand |= {"km","kilometer","bil","bilgodtgjorelse"}
        if "trekk" in t: expand |= {"trekk","trekkilonn","trekk_i_lonn"}
        if "bonus" in t: expand |= {"bonus"}
        if "overtid" in t: expand |= {"overtid"}
    base |= expand
    return base

def _category_for_code(code_tokens: Set[str]) -> str:
    if {"reise","kilometer","km","diett","overnatting","bil"} & code_tokens: return "travel"
    if {"trekk","trekkilonn","trekk_i_lonn"} & code_tokens: return "deduction"
    if {"ferie","feriepenger"} & code_tokens: return "holiday"
    if {"arbeidsgiveravgift","aga"} & code_tokens: return "employer_tax"
    return "wage"

def _category_for_account(accno: str, name_tokens: Set[str]) -> str:
    digits = re.sub(r"\D+", "", str(accno))
    if digits:
        if digits[:2] in {"70","71","72","73"}: return "travel"
        if digits[:2] in {"54","55"}: return "employer_tax"
        if digits[:2] in {"50","51","52","53"}: return "wage"
        if digits[:2] in {"29"}:      return "provision"
        if digits[:1] in {"1","2"}:   return "balance"
    if {"reise","diett","overnatting","kilometer","km","bil"} & name_tokens: return "travel"
    if {"aga","arbeidsgiveravgift"} & name_tokens: return "employer_tax"
    if {"ferie","feriepenger","skyldig"} & name_tokens: return "provision"
    return "wage"

def _jaccard(a: Set[str], b: Set[str]) -> float:
    if not a and not b: return 0.0
    inter = len(a & b); union = len(a | b)
    return inter / union if union else 0.0

def _magnitude_score(acc_amount: float, code_amount: float) -> float:
    A = abs(acc_amount); B = abs(code_amount)
    if A == 0 and B == 0: return 1.0
    if A == 0 or B == 0: return 0.0
    return max(0.0, 1.0 - abs(A - B) / max(A, B))

def _sign_score(acc_amount: float, code_amount: float) -> float:
    if acc_amount == 0 or code_amount == 0: return 0.5
    sa = 1 if acc_amount > 0 else -1
    sb = 1 if code_amount > 0 else -1
    return 1.0 if sa == sb else 0.0

# ---------- Fallback-matcher ----------

def suggest_mapping_for_accounts(
    gl_accounts: List[Dict[str, Any]],
    a07_sums: Dict[str, float],
    *,
    min_score: float = 0.55,
) -> Dict[str, Dict[str, Any]]:
    """
    Stram fallback: kun lønnsrelevante kontointervaller + sane beløp + min-score.
    Returnerer {konto: {"kode":..., "score":..., "reason":...}}
    """
    # Lønnsrelevante intervaller
    PAYROLL_RANGES = [(5000, 5999), (7000, 7399)]
    SPECIAL_ACCOUNTS = {"2940", "5290"}  # 2940 (skyldig ferie), 5290 (trekk i lønn for ferie)

    def _is_payroll_account(accno: str) -> bool:
        digits = re.sub(r"\D+", "", str(accno))
        if not digits:
            return False
        if digits in SPECIAL_ACCOUNTS:
            return True
        v = int(digits)
        return any(lo <= v <= hi for lo, hi in PAYROLL_RANGES)

    # Ban-ord: bank, fordringer, inntekt, aksjer, mva, etc
    BAN_WORDS = {
        "bank","kundefordring","leverandør","leverandor","fordring","inntekt","salg","omsetning",
        "aksje","aksjekapital","mva","merverdiavgift","avgift","lån","lan","obligasjon","varelager",
        "porto","telefon","it","utbytte","innskudd","saldo","kund","leverand","lindorff","intercompany"
    , "reklame","markedsforing","markedsføring","marketing","annonser","ads","google","facebook","linkedin","salg","promotion" }

    def _banned(name: str) -> bool:
        t = _tokens_from_text(name)
        return any(b in t for b in BAN_WORDS)

    # Forbered A07-metadata
    code_meta = {}
    for code, amt in a07_sums.items():
        toks = _tokens_from_code(code)
        cat = _category_for_code(toks)
        code_meta[code] = {"tokens": toks, "cat": cat, "amount": float(amt)}

    out: Dict[str, Dict[str, Any]] = {}

    for acc in gl_accounts:
        accno = str(acc["konto"])
        name = acc.get("navn", "")
        amount = float(acc.get("endring", acc.get("belop", 0.0)))

        if abs(amount) < 1e-9 and abs(float(acc.get("ub", 0.0))) < 1e-9:
            continue
        if not _is_payroll_account(accno):
            continue
        if _banned(name):
            continue

        name_tokens = _tokens_from_text(name)
        acc_cat = _category_for_account(accno, name_tokens)

        best_code, best_score, best_reason = None, -1.0, ""
        for code, meta in code_meta.items():
            code_tokens = meta["tokens"]; code_cat = meta["cat"]; code_amt = meta["amount"]

            # Beløps-sanity: ikke la ekstrem ratio slippe inn (f.eks. 40x)
            if abs(code_amt) > 0 and abs(amount) / max(abs(code_amt), 1.0) > 8.0:
                continue

            s_name = _jaccard(name_tokens, code_tokens)
            s_mag  = _magnitude_score(amount, code_amt)
            s_sign = _sign_score(amount, code_amt)
            if s_sign < 0.5:
                continue

            s_cat  = 1.0 if acc_cat == code_cat else 0.0
            score = 0.60*s_name + 0.15*s_mag + 0.15*s_sign + 0.10*s_cat

            if score < min_score:
                continue

            reason_parts = []
            if s_name >= 0.4: reason_parts.append(f"navn~kode {s_name:.2f}")
            if s_cat  > 0.5:  reason_parts.append(f"kategori {acc_cat}")
            if s_sign > 0.5:  reason_parts.append("tegn samsvar")
            if s_mag  > 0.6:  reason_parts.append("størrelse nær")

            if score > best_score:
                best_code, best_score, best_reason = code, score, ", ".join(reason_parts) if reason_parts else "OK"

        if best_code:
            out[accno] = {"kode": best_code, "score": round(float(best_score), 3), "reason": best_reason}

    return out

# a07_rulebook.py
from __future__ import annotations
import csv, json, os, re
from typing import Any, Dict, List, Tuple, Set, Optional

# RapidFuzz for bedre tekstmatching (pip install rapidfuzz)
try:
    from rapidfuzz import fuzz
    HAVE_RAPIDFUZZ = True
except ImportError:
    HAVE_RAPIDFUZZ = False

# ------------------------ små utils ------------------------

def _norm(s: Optional[str]) -> str:
    if not s: return ""
    s = s.strip().lower()
    return s.replace("ø","oe").replace("å","aa").replace("æ","ae")

_WORD_RE = re.compile(r"[a-z0-9]+")

def _tokens(text: str) -> Set[str]:
    return set(_WORD_RE.findall(_norm(text)))

def jaccard(a: Set[str], b: Set[str]) -> float:
    if not a and not b: return 0.0
    return len(a & b) / len(a | b)

def magnitude_score(a: float, b: float) -> float:
    A = abs(a); B = abs(b)
    if A == 0 and B == 0: return 1.0
    if A == 0 or B == 0: return 0.0
    return max(0.0, 1.0 - abs(A-B)/max(A,B))

def sign_score(a: float, b: float) -> float:
    if a == 0 or b == 0: return 0.5
    sa = 1 if a>0 else -1; sb = 1 if b>0 else -1
    return 1.0 if sa==sb else 0.0

def _parse_expected_sign(s: str) -> int:
    """
    '+', 'pos', '1' -> +1; '-', 'neg', '-1' -> -1; blank -> 0.
    """
    s = (s or "").strip().lower()
    if s in {"+","+1","1","pos","positive","pluss"}: return 1
    if s in {"-","-1","neg","negative","minus"}: return -1
    return 0

# ------------------------ allowed ranges ------------------------

def _token_to_interval(tok: str) -> Tuple[int,int]:
    """'2940' -> (2940,2940); '50xx' -> (5000,5099); '5xxx' -> (5000,5999)"""
    tok = tok.strip().lower().replace("*","x")
    m = re.fullmatch(r"(\d+)(x+)", tok)
    if m:
        base = int(m.group(1)); k = len(m.group(2))
        low = base * (10**k); high = low + (10**k) - 1
        return low, high
    if re.fullmatch(r"\d+", tok):
        v = int(tok); return v, v
    raise ValueError(f"Ukjent områdetok: {tok}")

def _parse_range_expr(expr: str) -> List[Tuple[int,int]]:
    if not expr: return []
    parts = re.split(r"[|,;]+", expr)
    intervals: List[Tuple[int,int]] = []
    for p in parts:
        p = p.strip()
        if not p: continue
        if "-" in p:
            a,b = p.split("-",1)
            try: lo1,hi1 = _token_to_interval(a)
            except: lo1 = int(a); hi1 = int(a)
            try: lo2,hi2 = _token_to_interval(b)
            except: lo2 = int(b); hi2 = int(b)
            intervals.append((min(lo1,lo2), max(hi1,hi2)))
        else:
            try: intervals.append(_token_to_interval(p))
            except:
                m = re.fullmatch(r"(\d+)", p)
                if not m: raise
                v = int(m.group(1)); intervals.append((v,v))
    return intervals

def _in_any_interval(accno: str, intervals: List[Tuple[int,int]]) -> bool:
    digits = re.sub(r"\D+","", str(accno))
    if not digits: return False
    v = int(digits)
    return any(lo <= v <= hi for lo,hi in intervals)

# ------------------------ I/O ------------------------

def _read_csv(path: str) -> List[Dict[str, str]]:
    """Read a CSV file into a list of dictionaries.

    This helper cleans up keys and values by stripping leading/trailing
    whitespace and ignores any columns without a header.  It also
    protects against ``NoneType`` keys or values which could cause
    attribute errors when calling ``strip()`` on ``None``.

    Each row in the returned list contains only those columns that
    have a non-empty header.
    """
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        rows: List[Dict[str, str]] = []
        for row in rd:
            cleaned: Dict[str, str] = {}
            for k, v in row.items():
                # ``csv.DictReader`` returns ``None`` for extra columns if
                # the row has more fields than headers; skip those.
                key = (k or "").strip()
                if not key:
                    continue
                cleaned[key] = (v or "").strip()
            # Skip completely empty rows
            if any(cleaned.values()):
                rows.append(cleaned)
        return rows

def _read_excel(path: str) -> Dict[str, List[Dict[str,str]]]:
    try:
        import pandas as pd  # type: ignore
    except Exception:
        raise RuntimeError("Excel-lesing krever 'pandas' + 'openpyxl'. Installer eller bruk CSV.")
    xl = pd.ExcelFile(path)
    out: Dict[str, List[Dict[str,str]]] = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet).fillna("")
        out[sheet] = [{str(k).strip(): str(v).strip() for k,v in rec.items()} for rec in df.to_dict("records")]
    return out

def load_rulebook(path_or_dir: Optional[str]=None) -> Dict[str, Any]:
    """
    Laster regelbok fra mappe (CSV: a07_codes.csv, aliases.csv) eller Excel (ark: a07_codes, aliases).
    Returnerer: {'codes': {...}, 'aliases': {...}, 'source': ...}
    """
    if path_or_dir is None:
        for cand in ["rulebook.xlsx", "data/rulebook.xlsx"]:
            if os.path.exists(cand):
                path_or_dir = cand; break
        if path_or_dir is None:
            path_or_dir = "data"

    rules: Dict[str, Any] = {"codes": {}, "aliases": {}, "source": path_or_dir}

    def _load_codes(rows: List[Dict[str,str]]):
        for r in rows:
            code = r.get("a07_code","").strip()
            if not code: continue
            rules["codes"][code] = {
                "label":   r.get("label","").strip(),  # valgfri, for pen visning (f.eks. 'Fastlønn')
                "category": r.get("category","wage").strip() or "wage",
                "basis":    r.get("gl_basis_default","auto").strip() or "auto",
                "allowed":  _parse_range_expr(r.get("allowed_ranges","")),
                "keywords": set(t for t in re.split(r"[|,;]\s*", r.get("keywords","")) if t),
                "boost_accounts": set(t for t in re.split(r"[|,;]\s*", r.get("boost_accounts","")) if t),
                "special_add": json.loads(r.get("special_add","[]") or "[]"),
                "expected_sign": _parse_expected_sign(r.get("expected_sign","")),
            }

    if os.path.isdir(path_or_dir):
        codes_csv = os.path.join(path_or_dir, "a07_codes.csv")
        aliases_csv = os.path.join(path_or_dir, "aliases.csv")
        if os.path.exists(codes_csv):
            _load_codes(_read_csv(codes_csv))
        if os.path.exists(aliases_csv):
            for r in _read_csv(aliases_csv):
                can = _norm(r.get("canonical",""))
                syns = [_norm(t) for t in re.split(r"[|,;]\s*", r.get("synonyms","")) if t]
                if can: rules["aliases"][can] = set(syns)
        return rules

    if path_or_dir.lower().endswith(".xlsx"):
        sheets = _read_excel(path_or_dir)
        _load_codes(sheets.get("a07_codes", []))
        for r in sheets.get("aliases", []):
            can = _norm(r.get("canonical",""))
            syns = [_norm(t) for t in re.split(r"[|,;]\s*", r.get("synonyms","")) if t]
            if can: rules["aliases"][can] = set(syns)
        return rules

    raise FileNotFoundError(f"Fant ingen regelbok på {path_or_dir}")

# ------------------------ tekstscore v2 ------------------------

def _alias_bag_for_code(code: str, rule: Dict[str,Any], aliases_map: Dict[str,Set[str]]) -> Tuple[List[str], Set[str]]:
    """
    Returnerer (liste av alias-fraser, sett av enkelt-ord).
    """
    bag = set()
    bag.add(_norm(code))
    for k in rule.get("keywords", []):
        bag.add(_norm(k))
        bag |= set(aliases_map.get(_norm(k), []))
    bag_list = list(bag)
    bag_tokens: Set[str] = set()
    for t in bag_list:
        bag_tokens |= _tokens(t)
    return bag_list, bag_tokens

def _name_similarity_v2(name: str, bag_list: List[str], bag_tokens: Set[str]) -> float:
    """
    1) Hele-ord-treff → høy basis (0.9)
    2) Ellers fuzzy mot aliasfraser, men ignorer veldig korte termer (≤3)
    3) Returner en konservativ score (0..1)
    """
    nrm = _norm(name)
    name_tokens = _tokens(name)
    if name_tokens & bag_tokens:
        return 0.90  # klart signal: samme ord finnes

    best = 0.0
    if HAVE_RAPIDFUZZ:
        for term in bag_list:
            if len(term) <= 3:   # stopp "lon"→"telefon"
                continue
            score = max(fuzz.WRatio(nrm, term), fuzz.token_set_ratio(nrm, term)) / 100.0
            if score > best: best = score
    else:
        best = jaccard(name_tokens, bag_tokens)
    return best * 0.70

# ------------------------ forslag basert på regelbok (v2) ------------------------

def suggest_with_rulebook(
    gl_accounts: List[Dict[str,Any]],
    a07_sums: Dict[str,float],
    rulebook: Dict[str,Any],
    *,
    min_score: float = 0.55,
    min_name: float = 0.35,
    min_margin: float = 0.12,
) -> Dict[str, Dict[str,Any]]:
    """
    Returnerer forslag pr konto med streng gating + minstekrav + tie-break:
      { konto: { 'kode':..., 'score':..., 'reason':... } }
    """
    out: Dict[str, Dict[str,Any]] = {}

    # precompute meta
    code_meta: Dict[str, Dict[str,Any]] = {}
    for code, rule in rulebook.get("codes", {}).items():
        bag_list, bag_tokens = _alias_bag_for_code(code, rule, rulebook.get("aliases", {}))
        code_meta[code] = {
            "sum": a07_sums.get(code, 0.0),
            "allowed": rule.get("allowed", []),
            "bag_list": bag_list,
            "bag_tokens": bag_tokens,
            "boost_accounts": set(rule.get("boost_accounts", [])),
            "expected_sign": int(rule.get("expected_sign", 0)),
        }

    for acc in gl_accounts:
        accno = str(acc["konto"])
        name = acc.get("navn","")
        amt  = float(acc.get("endring", acc.get("belop", 0.0)))

        if abs(amt) < 1e-9 and abs(float(acc.get("ub",0.0))) < 1e-9:
            continue

        candidates: List[Tuple[str,float,str,float]] = []  # (code, total_score, reason, name_score)

        for code, meta in code_meta.items():
            allowed_intervals = meta["allowed"]
            if allowed_intervals and not _in_any_interval(accno, allowed_intervals):
                continue

            csum = meta["sum"]
            exp = int(meta.get("expected_sign", 0))
            if exp == 1 and amt <= 0:
                continue
            if exp == -1 and amt >= 0:
                continue
            s_name = _name_similarity_v2(name, meta["bag_list"], meta["bag_tokens"])
            s_mag  = magnitude_score(amt, csum)
            s_sign = sign_score(amt, csum)
            b_series = 0.10                                  # liten bonus – vi er allerede innenfor range
            b_boost  = 0.20 if accno in meta["boost_accounts"] else 0.0

            score = 0.60*s_name + 0.15*s_mag + 0.15*s_sign + b_series + b_boost
            reason = []
            if b_boost>0:   reason.append("spesialkonto-boost")
            reason.append(f"navn~kode {s_name:.2f}")
            reason.append("størrelse nær" if s_mag>0.6 else "størrelse avvik")
            reason.append("tegn samsvar" if s_sign>0.5 else "tegn ulikt")
            candidates.append((code, score, ", ".join(reason), s_name))

        if not candidates:
            continue

        candidates.sort(key=lambda t: t[1], reverse=True)
        best_code, best_score, best_reason, best_name = candidates[0]
        second_score = candidates[1][1] if len(candidates) > 1 else 0.0

        if best_name < min_name or best_score < min_score or (best_score - second_score) < min_margin:
            continue  # ikke trygt nok – heller ingen forslag

        out[accno] = {"kode": best_code, "score": round(best_score,3), "reason": best_reason}

    return out

# ------------------------ forklarende avvisninger ------------------------

def explain_account(
    gl_account: Dict[str,Any],
    a07_sums: Dict[str,float],
    rulebook: Dict[str,Any],
    *,
    min_score: float = 0.55,
    min_name: float = 0.35,
    min_margin: float = 0.12,
) -> List[str]:
    """
    Forklarer hvorfor en konto ikke fikk forslag:
    - Utenfor allowed-range
    - Forventet tegn mismatch
    - Lav navnescore
    - Lav totalscore eller for liten margin
    Returnerer en liste med korte utsagn (maks ~1 pr kode).
    """
    accno = str(gl_account.get("konto","")); name = str(gl_account.get("navn",""))
    amt = float(gl_account.get("endring", gl_account.get("belop", 0.0)))

    reasons: List[str] = []
    cand: List[Tuple[str,float,float,float]] = []  # code, s_name, score, second_gap
    all_rej = 0

    # forbered meta
    code_meta: Dict[str, Dict[str,Any]] = {}
    for code, rule in rulebook.get("codes", {}).items():
        bag_list, bag_tokens = _alias_bag_for_code(code, rule, rulebook.get("aliases", {}))
        code_meta[code] = {
            "sum": a07_sums.get(code, 0.0),
            "allowed": rule.get("allowed", []),
            "bag_list": bag_list,
            "bag_tokens": bag_tokens,
            "boost_accounts": set(rule.get("boost_accounts", [])),
            "expected_sign": int(rule.get("expected_sign", 0)),
        }

    # Finn topp-2 for marginsjekk
    scored: List[Tuple[str,float,float]] = []
    for code, meta in code_meta.items():
        allowed_intervals = meta["allowed"]
        if allowed_intervals and not _in_any_interval(accno, allowed_intervals):
            all_rej += 1; continue
        exp = int(meta.get("expected_sign", 0))
        if exp == 1 and amt <= 0: all_rej += 1; continue
        if exp == -1 and amt >= 0: all_rej += 1; continue

        s_name = _name_similarity_v2(name, meta["bag_list"], meta["bag_tokens"])
        s_mag  = magnitude_score(amt, meta["sum"])
        s_sign = sign_score(amt, meta["sum"])
        b_series = 0.10
        b_boost  = 0.20 if accno in meta["boost_accounts"] else 0.0
        score    = 0.60*s_name + 0.15*s_mag + 0.15*s_sign + b_series + b_boost
        scored.append((code, s_name, score))

    if not scored:
        # mest sannsynlig allowed-range/tegnmismatch
        reasons.append("Ingen kandidater etter gating (allowed-range/forventet tegn).")
        return reasons

    scored.sort(key=lambda t: t[2], reverse=True)
    best_code, best_name, best_score = scored[0]
    second_score = scored[1][2] if len(scored)>1 else 0.0

    if best_name < min_name:
        reasons.append(f"Beste kandidat '{best_code}': navnescore for lav ({best_name:.2f} < {min_name:.2f}).")
    if best_score < min_score:
        reasons.append(f"Beste kandidat '{best_code}': totalscore {best_score:.2f} < {min_score:.2f}.")
    if (best_score - second_score) < min_margin:
        reasons.append(f"For liten margin mellom #1 og #2 ({(best_score-second_score):.2f} < {min_margin:.2f}).")

    if not reasons:
        reasons.append("Ingen forslag av forsiktighetshensyn (terskler/margin).")
    return reasons

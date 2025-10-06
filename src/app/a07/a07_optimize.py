# -*- coding: utf-8 -*-
"""
a07_optimize.py
---------------
Kandidater for LP og global beløpsmatching (mange-til-mange) via PuLP.
Forutsetter at a07_rulebook.py er tilgjengelig når du bruker kandidatbyggingen.
"""

from __future__ import annotations
from typing import Any, Dict, List, Tuple, Optional, Set

try:
    from a07_rulebook import (
        _alias_bag_for_code, _name_similarity_v2, _in_any_interval,
        magnitude_score as rb_magnitude_score, sign_score as rb_sign_score
    )
    _HAVE_RULEBOOK_HELPERS = True
except Exception:
    _HAVE_RULEBOOK_HELPERS = False

def _rb_mag(a: float, b: float) -> float:
    if _HAVE_RULEBOOK_HELPERS and rb_magnitude_score: return rb_magnitude_score(a,b)
    A = abs(a); B = abs(b)
    if A == 0 and B == 0: return 1.0
    if A == 0 or B == 0: return 0.0
    return max(0.0, 1.0 - abs(A - B) / max(A, B))

def _rb_sign(a: float, b: float) -> float:
    if _HAVE_RULEBOOK_HELPERS and rb_sign_score: return rb_sign_score(a,b)
    if a == 0 or b == 0: return 0.5
    sa = 1 if a > 0 else -1
    sb = 1 if b > 0 else -1
    return 1.0 if sa == sb else 0.0

def generate_candidates_for_lp(
    gl_accounts: List[Dict[str,Any]],
    a07_sums: Dict[str,float],
    rulebook: Dict[str,Any],
    *,
    amounts_override: Optional[Dict[str,float]] = None,
    min_name: float = 0.25,
    min_score: float = 0.40,
    top_k: int = 3,
    skip_edges: Optional[Set[Tuple[str,str]]] = None,
) -> Dict[str, List[Tuple[str,float,float,str]]]:
    """
    Returnerer per konto: [(kode, score, amount, reason)] (topp-k).
    Bruker allowed-ranges (hard gate), alias-bag og navnescore fra regelbok-hjelpere.
    """
    if not (_HAVE_RULEBOOK_HELPERS and _alias_bag_for_code and _name_similarity_v2 and _in_any_interval):
        return {}

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

    out: Dict[str, List[Tuple[str,float,float,str]]] = {}
    for acc in gl_accounts:
        accno = str(acc["konto"])
        name  = acc.get("navn","")
        amt   = amounts_override.get(accno) if amounts_override else float(acc.get("endring", acc.get("belop", 0.0)))
        if amt is None: amt = 0.0
        if abs(amt) < 1e-9 and abs(float(acc.get("ub",0.0))) < 1e-9:
            continue

        cands: List[Tuple[str,float,float,str,float]] = []
        for code, meta in code_meta.items():
            if skip_edges and (accno,code) in skip_edges:
                continue
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
            s_mag  = _rb_mag(amt, csum)
            s_sign = _rb_sign(amt, csum)
            b_series = 0.10
            b_boost  = 0.20 if accno in meta["boost_accounts"] else 0.0
            score = 0.60*s_name + 0.15*s_mag + 0.15*s_sign + b_series + b_boost
            if s_name < min_name or score < min_score:
                continue
            reason = []
            if b_boost>0: reason.append("spesialkonto-boost")
            reason.append(f"navn~kode {s_name:.2f}")
            reason.append("størrelse nær" if s_mag>0.6 else "størrelse avvik")
            reason.append("tegn samsvar" if s_sign>0.5 else "tegn ulikt")
            cands.append((code, score, float(amt), ", ".join(reason), s_name))

        cands.sort(key=lambda t: t[1], reverse=True)
        out[accno] = [(c, s, a, r) for (c, s, a, r, _sn) in cands[:top_k]]

    return out

def solve_global_assignment_lp(
    amounts: Dict[str, float],
    candidates: Dict[str, List[Tuple[str,float,float,str]]],
    targets: Dict[str, float],
    *,
    allow_splits: bool = True,
    lambda_score: float = 0.25,
    lambda_sign: float  = 0.05,
    lambda_edges: float = 0.00,   # <--- NY: svak sparsity-straff per aktiv kant
) -> Dict[str, Dict[str, float]]:
    """
    LP: minimer Σ_c |Σ_a A[a]·x[a,c] − T[c]| + straffer.
    Returnerer {konto: {kode: andel(0..1), ...}}
    """
    try:
        import pulp
    except Exception:
        raise RuntimeError("PuLP mangler. Installer: pip install pulp")

    accounts = list(amounts.keys())
    codes = set(targets.keys())
    for a in candidates.values():
        for (code, _s, _amt, _r) in a:
            codes.add(code)
    codes = sorted(codes)

    prob = pulp.LpProblem("A07_GlobalMapping", pulp.LpMinimize)

    # x[a,c]
    X: Dict[Tuple[str,str], Tuple[Any,float]] = {}
    for acc in accounts:
        for (code, score, _amt, _reason) in candidates.get(acc, []):
            var = pulp.LpVariable(
                f"x_{acc}_{code}", lowBound=0, upBound=1,
                cat=("Continuous" if allow_splits else "Binary")
            )
            X[(acc,code)] = (var, score)

    # Hjelpevariabler Y[a,c] for sparsity- (kant-)straff
    Y: Dict[Tuple[str,str], Any] = {}
    if lambda_edges > 0.0:
        for (a,c) in X:
            y = pulp.LpVariable(f"y_{a}_{c}", lowBound=0, upBound=1, cat="Binary")
            Y[(a,c)] = y
            prob += (X[(a,c)][0] <= y), f"edge_on_{a}_{c}"

    # kapasitet per konto
    for acc in accounts:
        vars_for_acc = [X[(acc,c)][0] for (a,c) in X if a==acc]
        if vars_for_acc:
            prob += (pulp.lpSum(vars_for_acc) <= 1), f"cap_{acc}"

    # absoluttslack pr kode
    POS = {c: pulp.LpVariable(f"pos_{c}", lowBound=0) for c in codes}
    NEG = {c: pulp.LpVariable(f"neg_{c}", lowBound=0) for c in codes}
    for c in codes:
        lhs = 0
        if X:
            lhs = (sum(amounts[a] * X[(a,cc)][0] for (a,cc) in X if cc==c))
        prob += (lhs - targets.get(c,0.0) == POS[c] - NEG[c]), f"bal_{c}"

    # objekt
    obj = sum(POS.values()) + sum(NEG.values())
    for (a,c), (var, score) in X.items():
        amt = abs(amounts[a])
        obj += lambda_score * (1.0 - float(score)) * amt * var
        t = targets.get(c, 0.0)
        if amt > 0 and abs(t) > 1e-9:
            sign_mismatch = 1.0 if (amounts[a] > 0) != (t > 0) else 0.0
            if sign_mismatch:
                obj += lambda_sign * amt * var

    if lambda_edges > 0.0:
        obj += lambda_edges * pulp.lpSum(Y.values())

    prob += obj
    prob.solve(getattr(__import__("pulp"), "PULP_CBC_CMD")(msg=False))

    assignment: Dict[str, Dict[str,float]] = {}
    for (a,c), (var, _score) in X.items():
        try:
            val = var.value()
        except Exception:
            from pulp import value as _pv
            val = _pv(var)
        if val and val > 1e-6:
            assignment.setdefault(a, {})[c] = float(val)
    return assignment

# a07_matcher.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from a07_models import (
    GLAccount,
    A07Entry,
    A07CodeDef,
    AmountMetric,
    amount_for_account,
    jaccard,
    tokenize,
    account_in_ranges,
    to_money,
)

# -----------------------------------------------------------------------------
# Resultatmodeller
# -----------------------------------------------------------------------------

@dataclass
class MapHit:
    """Foreslått løsning for ett mål (kode eller gruppe)."""
    target_id: str              # 'fastloenn' eller 'bundle:1' osv
    target_name: str
    target_amount: Decimal
    accounts: List[int]         # kontonr som inngår
    sum_amount: Decimal
    diff: Decimal
    score: float                # total heuristikk-score


@dataclass
class SuggestResult:
    hits: List[MapHit]                  # aksepterte forslag
    unused_accounts: List[int]          # konti som ikke ble brukt
    unmatched_targets: List[str]        # koder/grupper vi ikke fant løsning til


# -----------------------------------------------------------------------------
# Kjernealgoritme
# -----------------------------------------------------------------------------

def _account_amount(acc: GLAccount, metric: Optional[AmountMetric]) -> Decimal:
    return amount_for_account(acc, metric)


def _candidate_score(code_def: Optional[A07CodeDef],
                     acc: GLAccount,
                     metric: Optional[AmountMetric],
                     target_tokens: Sequence[str],
                     target_sign: Optional[int]) -> float:
    """
    Vurderer hvor "lovende" en konto er før subset-søk:
    - +2.0 hvis i kontointervall
    - +1.0 hvis fortegn stemmer
    - + Jaccard(tekst)*1.0
    """
    score = 0.0
    amt = _account_amount(acc, metric)
    if code_def and code_def.contains_account(acc.konto):
        score += 2.0
    if target_sign:
        if target_sign > 0 and amt >= 0 or target_sign < 0 and amt <= 0:
            score += 1.0
    if code_def:
        score += jaccard(acc.tokens(), target_tokens) * 1.0
    return score


def _prefilter_candidates(gl_accounts: List[GLAccount],
                          code_def: Optional[A07CodeDef],
                          metric: Optional[AmountMetric],
                          target_amount: Decimal,
                          target_tokens: Sequence[str],
                          target_sign: Optional[int],
                          keep_top: int = 60) -> List[GLAccount]:
    """Filtrer/sorter kandidater før subset-søk."""
    # 1) Grovfilter: i intervall ELLER decent tekst-samsvar (>= 0.2) ELLER beløp nær mål
    def ok(acc: GLAccount) -> bool:
        amt = abs(_account_amount(acc, metric))
        near = abs(amt - abs(target_amount)) <= abs(target_amount) * Decimal("0.25") + Decimal("500.00")
        in_range = bool(code_def and code_def.contains_account(acc.konto))
        sim = jaccard(acc.tokens(), target_tokens)
        return in_range or sim >= 0.20 or near

    pool = [a for a in gl_accounts if ok(a)]

    # 2) Skår for prioritet
    def key(acc: GLAccount) -> Tuple[float, Decimal]:
        s = _candidate_score(code_def, acc, metric, target_tokens, target_sign)
        # nærmest i absolutt beløp
        amt = abs(_account_amount(acc, metric))
        dist = abs(abs(target_amount) - amt)
        return (s, -dist)

    pool.sort(key=key, reverse=True)
    return pool[:keep_top]


def _subset_sum(accounts: List[GLAccount],
                metric: Optional[AmountMetric],
                target: Decimal,
                max_diff: Decimal,
                max_k: int) -> Optional[List[GLAccount]]:
    """
    Prøver kombinasjoner (1..max_k) med begrenset branching.
    Tillater negative/positive beløp – vi bruker bruteforce++ med trimming
    på størrelse og "nærhet".
    """
    target = to_money(target)
    amounts = [(acc, _account_amount(acc, metric)) for acc in accounts]

    # Rekkefølge: nærmest mål først
    amounts.sort(key=lambda t: abs(abs(target) - abs(t[1])))

    # 1) enkelt-treff
    for acc, amt in amounts:
        if abs(amt - target) <= max_diff:
            return [acc]

    # 2) par
    if max_k >= 2:
        n = len(amounts)
        for i in range(n):
            a1, v1 = amounts[i]
            for j in range(i + 1, min(i + 1 + 200, n)):  # begrens O(n^2)
                a2, v2 = amounts[j]
                s = v1 + v2
                if abs(s - target) <= max_diff:
                    return [a1, a2]

    # 3..k – liten DFS med trimming
    def dfs(start: int, k_left: int, cur_list: List[GLAccount], cur_sum: Decimal) -> Optional[List[GLAccount]]:
        if k_left == 0:
            return None
        for idx in range(start, min(start + 60, len(amounts))):  # begrens dybde
            a, v = amounts[idx]
            new_sum = cur_sum + v
            if abs(new_sum - target) <= max_diff:
                return cur_list + [a]
            # heuristisk trimming: stopp hvis vi har for mange elementer
            if k_left > 1:
                out = dfs(idx + 1, k_left - 1, cur_list + [a], new_sum)
                if out:
                    return out
        return None

    for k in range(3, max_k + 1):
        hit = dfs(0, k, [], Decimal("0.00"))
        if hit:
            return hit
    return None


def suggest_mappings(
    a07_entries: Dict[str, A07Entry],
    code_defs: Dict[str, A07CodeDef],
    gl_accounts: List[GLAccount],
    metric: Optional[AmountMetric] = None,
    max_diff: Decimal = Decimal("5.00"),
    max_combo_size: int = 4,
    groups: Optional[Dict[str, List[str]]] = None,
) -> SuggestResult:
    """
    Returnerer forslag for alle mål (koder og ev. grupper).
    Strategi:
      1) Beløp først (1..N kontoer), prioritér kontointervall og fortegn
      2) Hvis flere muligheter, velg den med færrest konti og høyest heuristikk-score
      3) Hver konto brukes maks én gang (greedy)
    """
    # Forbered mål-liste (grupper + koder som ikke inngår i grupper)
    grouped_codes = set()
    targets: List[Tuple[str, str, Decimal, Optional[int], Sequence[str]]] = []

    if groups:
        for gid, codes in groups.items():
            amt = sum((a07_entries[c].amount for c in codes if c in a07_entries), Decimal("0.00"))
            name = f"Bundle({', '.join(codes)})"
            # expected_sign – hvis alle code_defs har likt fortegn, bruk det
            signs = {code_defs[c].expected_sign for c in codes if c in code_defs and code_defs[c].expected_sign}
            exp = list(signs)[0] if len(signs) == 1 else None
            toks = []
            for c in codes:
                if c in code_defs:
                    toks.extend(code_defs[c].tokens())
            targets.append((f"bundle:{gid}", name, amt, exp, toks))
            grouped_codes.update(codes)

    for code, entry in a07_entries.items():
        if code in grouped_codes:
            continue
        cdef = code_defs.get(code)
        toks = cdef.tokens() if cdef else []
        targets.append((code, entry.name, entry.amount, cdef.expected_sign if cdef else None, toks))

    # Størst beløp først
    targets.sort(key=lambda t: abs(t[2]), reverse=True)

    used_accounts: set[int] = set()
    hits: List[MapHit] = []
    acc_map = {a.konto: a for a in gl_accounts}

    for target_id, target_name, target_amount, exp_sign, toks in targets:
        # tilgjengelige konti (ikke brukt)
        avail = [a for a in gl_accounts if a.konto not in used_accounts]
        cdef = code_defs.get(target_id) if not target_id.startswith("bundle:") else None

        cands = _prefilter_candidates(
            avail, cdef, metric, target_amount, toks, exp_sign, keep_top=80
        )
        # Prøv rask 1..N-kombo
        for k in (1, 2, 3, max_combo_size):
            hit = _subset_sum(cands, metric, target_amount, max_diff, max_k=k)
            if hit:
                s = sum((_account_amount(x, metric) for x in hit), Decimal("0.00"))
                diff = (s - target_amount).copy_abs()
                score = float(k == 1) * 2.0 + float(k == 2) * 1.0
                mh = MapHit(
                    target_id=target_id,
                    target_name=target_name,
                    target_amount=target_amount,
                    accounts=[x.konto for x in hit],
                    sum_amount=s,
                    diff=diff,
                    score=score,
                )
                hits.append(mh)
                used_accounts.update(mh.accounts)
                break
        else:
            # ingen treff
            pass

    unused = [k for k in acc_map if k not in used_accounts]
    unmatched = [tid for (tid, _, _, _, _) in targets if tid not in {h.target_id for h in hits}]
    return SuggestResult(hits=hits, unused_accounts=unused, unmatched_targets=unmatched)

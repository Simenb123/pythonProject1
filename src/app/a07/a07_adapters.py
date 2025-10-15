# -*- coding: utf-8 -*-
"""
Adaptere mellom 'legacy' models.GLAccount og nye a07_models.GLAccount,
samt noen småhjelpere for å velge beløp (IB/BEV/UB) konsistent i GUI/matcher.

Bruk:
- Etter at du har lest GL via models.read_gl_csv(...):
    legacy_accounts, meta = read_gl_csv(gl_path)
    new_accounts = legacy_list_to_new(legacy_accounts)

- Når du trenger beløp for en konto (respekterer IB/BEV/UB eller default):
    from a07_models import AmountMetric
    amt = choose_amount(new_accounts[0], AmountMetric.BEV)  # eksempel

Hvis du vil mappe tilbake til legacy-objekter for gammel kode:
    legacy_obj = new_to_legacy(new_accounts[0])
"""

from __future__ import annotations

from typing import Iterable, List, Dict, Optional, Any, Union
from decimal import Decimal

# Nye domenemodeller (Decimal, AmountMetric, m.m.)
from a07_models import GLAccount as NewGL
from a07_models import AmountMetric, to_money

# Legacy modell (float, belop/endring/ub) – kan mangle i noen prosjekter.
try:
    from models import GLAccount as LegacyGL  # type: ignore
except Exception:  # pragma: no cover
    LegacyGL = None  # type: ignore


def _to_int(v: Any) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return 0


def legacy_to_new(one: "LegacyGL") -> NewGL:
    """
    Konverterer én legacy models.GLAccount -> a07_models.GLAccount.
    Feltmapping:
      konto(str) -> konto(int)
      ib(float) -> ib(Decimal)
      endring/bevegelse(float) -> bevegelse(Decimal)
      ub(float) eller belop(float) -> ub(Decimal) (fallback til belop)
    """
    # Hent felt med sikre fallbacks
    konto = _to_int(getattr(one, "konto", 0))
    navn = getattr(one, "navn", "") or ""
    ib = to_money(getattr(one, "ib", 0.0))
    # 'endring' i legacy tilsvarer 'bevegelse' i nye modeller
    beveg = getattr(one, "endring", getattr(one, "bevegelse", 0.0))
    bevegelse = to_money(beveg)
    # UB (eller beløp som fallback)
    ub_raw = getattr(one, "ub", None)
    if ub_raw is None:
        ub_raw = getattr(one, "belop", 0.0)
    ub = to_money(ub_raw)

    return NewGL(konto=konto, navn=navn, ib=ib, bevegelse=bevegelse, ub=ub)


def legacy_list_to_new(accounts: Iterable["LegacyGL"]) -> List[NewGL]:
    """Batch-konverter en liste av legacy GLAccounts."""
    return [legacy_to_new(a) for a in accounts]


def new_to_legacy(one: NewGL) -> "LegacyGL":
    """
    Konverterer a07_models.GLAccount -> legacy models.GLAccount.
    NB: Debet/Kredit finnes ikke i nye modeller; settes til 0.0.
    'belop' settes til UB som et greit standardvalg.
    """
    if LegacyGL is None:  # pragma: no cover
        raise RuntimeError(
            "Legacy models.GLAccount ikke tilgjengelig – kan ikke konvertere tilbake."
        )
    konto = str(one.konto)
    navn = one.navn
    ib = float(one.ib)
    endring = float(one.bevegelse)
    ub = float(one.ub)
    debet = 0.0
    kredit = 0.0
    belop = ub  # et fornuftig default-valg

    return LegacyGL(
        konto=konto,
        navn=navn,
        ib=ib,
        debet=debet,
        kredit=kredit,
        endring=endring,
        ub=ub,
        belop=belop,
    )


def choose_amount(acc: NewGL, metric: Optional[Union[str, AmountMetric]] = None) -> Decimal:
    """
    Velg beløp for en konto:
      - Hvis 'metric' settes ('IB'/'BEV'/'UB' eller AmountMetric), brukes den.
      - Ellers brukes kontoens default (UB, unntatt 1000–2999 -> BEV).
    """
    m: Optional[AmountMetric]
    if metric is None:
        m = acc.default_metric()
    elif isinstance(metric, AmountMetric):
        m = metric
    else:
        m = AmountMetric.from_str(metric)
    return acc.amount(m)


def bulk_choose_amount(accounts: Iterable[NewGL],
                       metric: Optional[Union[str, AmountMetric]] = None) -> Dict[int, Decimal]:
    """
    Velg beløp for mange konti på en gang.
    Returnerer {konto: Decimal_beløp}.
    """
    out: Dict[int, Decimal] = {}
    for a in accounts:
        out[a.konto] = choose_amount(a, metric)
    return out

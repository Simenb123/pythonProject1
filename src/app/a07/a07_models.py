# a07_models.py
# -*- coding: utf-8 -*-
"""
Grunnmodeller, parser-funksjoner og nytteverktøy for A07-prosjektet.

Denne modulen er bevisst fri for GUI-kode og regelbok-spesifikk logikk,
slik at den kan gjenbrukes, enhetstestes og brukes av flere apper
(A07 GUI, CLI-verktøy, batch-skript, osv.).

Hovedinnhold:
- Money/Decimal-håndtering (to penger m/ korrekt avrunding)
- Enum for beløpstype (IB/BEV/UB)
- Dataklasser: GLAccount, A07Entry, A07CodeDef (kun struktur)
- Parser for Saldobalanse-CSV (støtter ulike kolonnenavn/locale)
- Parser for A07-CSV (enkelt standardformat)
- Aggregering av A07-rader pr kode
- Tekst-tokenisering og kontointervall-hjelpere

Avhengigheter:
- Kun standardbibliotek. Ingen ekstra pakker kreves her.

Forventet bruk:
- Andre moduler (regelbok, matcher/solver, GUI) importerer herfra.
- Regelboken (egen modul) fyller A07CodeDef-verdier, alias, kontointervall, osv.
- Matcher/solver bruker GLAccount + A07Entry + hjelpefunksjoner for beløp og intervall.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation, getcontext
from enum import Enum
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import csv
import re

# ---------------------------------------------------------------------------
# Penger/Decimal-hjelpere
# ---------------------------------------------------------------------------

# Sett forutsigbar presisjon for penger
getcontext().prec = 28  # rikelig nok til summeringer
_TWO_PLACES = Decimal("0.01")


def to_money(value) -> Decimal:
    """
    Konverterer en hvilken som helst støttet verdi til Decimal(2 desimaler).
    Støtter norske formater som '12 345,67', '12.345,67', samt float/int/str.
    """
    if value is None:
        return Decimal("0.00")
    if isinstance(value, Decimal):
        return value.quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return Decimal("0.00")
        # Fjern tusenskilletegn (mellomrom eller punktum) og normaliser komma til punktum
        s = s.replace(" ", "")
        # Hvis både punktum og komma finnes, anta at komma er desimalskilletegn (no/nb)
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return Decimal(s).quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            # Fall back til å filtrere ut alt unntatt siffer, minus, og punktum
            s2 = re.sub(r"[^0-9\.\-]", "", s)
            if not s2:
                return Decimal("0.00")
            return Decimal(s2).quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)
    # ukjent type
    return Decimal("0.00")


def _TWOPLACES() -> Decimal:
    # liten helper for å hindre sirkulær init
    return _TWO_PLACES


def add_money(a: Decimal, b: Decimal) -> Decimal:
    return (a + b).quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)


def sub_money(a: Decimal, b: Decimal) -> Decimal:
    return (a - b).quantize(_TWOPLACES(), rounding=ROUND_HALF_UP)


# ---------------------------------------------------------------------------
# Beløpstype (IB/Bevegelse/UB)
# ---------------------------------------------------------------------------

class AmountMetric(Enum):
    IB = "IB"
    BEV = "BEV"   # bevegelse/endring i perioden
    UB = "UB"

    @classmethod
    def from_str(cls, s: str) -> "AmountMetric":
        s = (s or "").strip().upper()
        if s in ("IB", "INNGÅENDE", "INNGAENDE", "BALANSE_START", "BALSTART"):
            return cls.IB
        if s in ("BEV", "BEVEGELSE", "ENDRING", "ENDR"):
            return cls.BEV
        return cls.UB


# ---------------------------------------------------------------------------
# Datamodeller
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class GLAccount:
    """Én konto i saldobalansen."""
    konto: int
    navn: str
    ib: Decimal = field(default_factory=lambda: Decimal("0.00"))
    bevegelse: Decimal = field(default_factory=lambda: Decimal("0.00"))
    ub: Decimal = field(default_factory=lambda: Decimal("0.00"))

    def amount(self, metric: AmountMetric) -> Decimal:
        if metric == AmountMetric.IB:
            return self.ib
        if metric == AmountMetric.BEV:
            return self.bevegelse
        return self.ub

    def default_metric(self) -> AmountMetric:
        """
        Standardregel: UB på alle konti, unntak: balansekonti (1000–2999) der BEV ofte er mer korrekt
        ved sammenligning mot periodisert innberetning (f.eks. skyldig feriepenger).
        """
        if 1000 <= int(self.konto) <= 2999:
            return AmountMetric.BEV
        return AmountMetric.UB

    def tokens(self) -> List[str]:
        """Enkel tokenisering av kontonavn for tekstlig/alias match."""
        base = normalize_text(self.navn)
        # splitt på mellomrom/ikke-alfanumerisk, fjern korte tokens
        tokens = [t for t in re.split(r"[^a-z0-9æøå]+", base) if len(t) >= 2]
        # legg til konto som egen token
        tokens.append(str(self.konto))
        return list(dict.fromkeys(tokens))  # uniq, stabil rekkefølge


@dataclass
class A07Entry:
    """Aggregert A07-beløp per kode (eller pr bucket etter grouping)."""
    code: str
    name: str
    amount: Decimal

    def __post_init__(self):
        self.amount = to_money(self.amount)


@dataclass
class A07CodeDef:
    """
    Regel/definisjon for én A07-kode (fra regelbok).
    - account_ranges: Liste av (start, slutt) inklusive intervaller (eks: [(5000,5999),(2900,2949)])
    - expected_sign: +1 (positiv), -1 (negativ) eller 0/None (ikke tvang)
    - aliases/keywords: navnevarianter for tekst-søk
    """
    code: str
    name: str
    account_ranges: List[Tuple[int, int]] = field(default_factory=list)
    expected_sign: Optional[int] = None
    aliases: List[str] = field(default_factory=list)
    keywords: List[str] = field(default_factory=list)

    def contains_account(self, konto: int) -> bool:
        return account_in_ranges(konto, self.account_ranges)

    def tokens(self) -> List[str]:
        base = normalize_text(self.name)
        tk = [t for t in re.split(r"[^a-z0-9æøå]+", base) if len(t) >= 2]
        tk.extend([normalize_text(x) for x in self.aliases])
        tk.extend([normalize_text(x) for x in self.keywords])
        # uniq
        uniq = []
        seen = set()
        for t in tk:
            if t and t not in seen:
                uniq.append(t)
                seen.add(t)
        return uniq


# ---------------------------------------------------------------------------
# Parser(e) – GL CSV & A07 CSV
# ---------------------------------------------------------------------------

def read_gl_csv(path: str | Path) -> List[GLAccount]:
    """
    Leser Saldobalanse fra CSV. Støtter fleksible kolonnenavn:
        Konto/Kontonr/Kontonummer
        Navn/Kontonavn/Beskrivelse
        IB/Inngående/InnGAENDE
        Bevegelse/Endring/Endr
        UB/Utgående/Utsgaaende/Saldo
    Håndterer norske tallformater.
    """
    path = Path(path)
    rows: List[GLAccount] = []
    if not path.exists():
        raise FileNotFoundError(f"Fant ikke GL CSV: {path}")

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV mangler header.")
        # Normaliser header -> lavere, fjern mellomrom
        header_map = {h: normalize_header(h) for h in reader.fieldnames}

        def pick(row: dict, *cands: str) -> str | None:
            for k, norm in header_map.items():
                if norm in cands:
                    return row.get(k)
            return None

        for raw in reader:
            konto_str = pick(raw, "konto", "kontonr", "kontonummer")
            navn = pick(raw, "navn", "kontonavn", "beskrivelse") or ""
            if not konto_str:
                # hopp over linjer uten kontonr
                continue
            try:
                konto = int(str(konto_str).strip())
            except ValueError:
                continue

            ib = to_money(pick(raw, "ib", "inngaaende", "inngående", "balanse_start", "balstart"))
            bev = to_money(pick(raw, "bev", "bevegelse", "endring", "endr"))
            ub = to_money(pick(raw, "ub", "utgaaende", "utgående", "saldo", "balanse_slutt", "balslutt"))

            rows.append(GLAccount(konto=konto, navn=navn or "", ib=ib, bevegelse=bev, ub=ub))
    return rows


def read_a07_csv(path: str | Path) -> List[A07Entry]:
    """
    Leser A07-aggregat fra CSV i et enkelt format (en rad per kode):
        kode,kodenavn,beløp
    (kolonnenavn er fleksible – 'code', 'name', 'amount' godtas også.)
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Fant ikke A07 CSV: {path}")

    out: Dict[str, A07Entry] = {}

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("A07-CSV mangler header.")
        header_map = {h: normalize_header(h) for h in reader.fieldnames}

        def pick(row: dict, *cands: str) -> str | None:
            for k, norm in header_map.items():
                if norm in cands:
                    return row.get(k)
            return None

        for raw in reader:
            code = (pick(raw, "kode", "code") or "").strip()
            name = (pick(raw, "kodenavn", "name", "beskrivelse") or "").strip()
            amount = to_money(pick(raw, "beloep", "beløp", "amount", "sum"))
            if not code:
                continue
            if code in out:
                out[code].amount = add_money(out[code].amount, amount)
            else:
                out[code] = A07Entry(code=code, name=name or code, amount=amount)

    return list(out.values())


# ---------------------------------------------------------------------------
# Aggregering og hjelpefunksjoner
# ---------------------------------------------------------------------------

def aggregate_a07_rows(rows: Iterable[dict],
                       code_key: str = "kode",
                       name_key: str = "kodenavn",
                       amount_key: str = "beløp") -> Dict[str, A07Entry]:
    """
    Tar vilkårlige rader (eks. fra detalj A07-rapport) og aggregerer pr kode.
    Forvent at radene har minst 3 felt: kode, navn og beløp (navn på felt kan overstyres).
    """
    out: Dict[str, A07Entry] = {}
    for r in rows:
        code = str(r.get(code_key, "")).strip()
        name = str(r.get(name_key, "") or code).strip()
        amount = to_money(r.get(amount_key, 0))
        if not code:
            continue
        if code not in out:
            out[code] = A07Entry(code=code, name=name or code, amount=amount)
        else:
            out[code].amount = add_money(out[code].amount, amount)
    return out


def normalize_header(h: str) -> str:
    """Normaliserer headernavn: små bokstaver, æøå -> ascii, fjerner mellomrom/tegn."""
    h = (h or "").strip().lower()
    h = h.replace(" ", "").replace("\t", "")
    repl = {
        "ø": "o", "æ": "ae", "å": "aa",
        "é": "e", "è": "e", "ö": "o", "ä": "a",
    }
    for k, v in repl.items():
        h = h.replace(k, v)
    # typiske kolonnenavn -> standard
    synonyms = {
        "kontonr": "konto",
        "kontonummer": "konto",
        "kontonavn": "navn",
        "beskrivelse": "navn",
        "inngaaende": "inngående",
        "utgaaende": "utgående",
        "beloep": "beløp",
        "belop": "beløp",
    }
    h = synonyms.get(h, h)
    return h


def normalize_text(s: str) -> str:
    """Normaliserer tekst for sammenligning: lower, strip, æøå bevares, fjerner ekstra whitespace."""
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def account_in_ranges(konto: int, ranges: Sequence[Tuple[int, int]]) -> bool:
    for a, b in ranges:
        if a <= konto <= b:
            return True
    return False


def parse_account_ranges(spec: str) -> List[Tuple[int, int]]:
    """
    Parser en streng som '5000-5999|2900-2949|7000' til liste av intervaller.
    Enkeltkonti (uten bindestrek) håndteres som [x, x].
    Mellomrom er lov; tom streng -> tom liste.
    """
    out: List[Tuple[int, int]] = []
    if not spec:
        return out
    parts = [p.strip() for p in spec.split("|") if p.strip()]
    for p in parts:
        if "-" in p:
            a, b = p.split("-", 1)
            try:
                a_i, b_i = int(a.strip()), int(b.strip())
                if a_i > b_i:
                    a_i, b_i = b_i, a_i
                out.append((a_i, b_i))
            except ValueError:
                continue
        else:
            try:
                k = int(p)
                out.append((k, k))
            except ValueError:
                continue
    return out


def ranges_to_spec(ranges: Sequence[Tuple[int, int]]) -> str:
    """Omvendt av parse_account_ranges()."""
    parts = []
    for a, b in ranges:
        if a == b:
            parts.append(str(a))
        else:
            parts.append(f"{a}-{b}")
    return "|".join(parts)


def sign_ok(amount: Decimal, expected_sign: Optional[int]) -> bool:
    """
    Sjekker fortegn mot expected_sign:
    - None/0: ingen tvang
    - +1: må være >= 0.00 (toleranse: 0.00)
    - -1: må være <= 0.00
    """
    if not expected_sign:
        return True
    if expected_sign > 0:
        return amount >= Decimal("0.00")
    if expected_sign < 0:
        return amount <= Decimal("0.00")
    return True


# ---------------------------------------------------------------------------
# Enkle søke-/match-hjelpere (tekst)
# ---------------------------------------------------------------------------

_WORD = re.compile(r"[a-z0-9æøå]+")


def tokenize(s: str) -> List[str]:
    """Tokeniserer norsk/engelsk tekst – kun enkle ord/tegn, lower, uniq."""
    s = normalize_text(s)
    toks = _WORD.findall(s)
    uniq: List[str] = []
    seen = set()
    for t in toks:
        if len(t) < 2:
            continue
        if t not in seen:
            uniq.append(t)
            seen.add(t)
    return uniq


def jaccard(a: Sequence[str], b: Sequence[str]) -> float:
    """Jaccard-similarity for to token-sett (0..1)."""
    sa, sb = set(a), set(b)
    if not sa and not sb:
        return 0.0
    return len(sa & sb) / float(len(sa | sb))


# ---------------------------------------------------------------------------
# Eksempelhjelp for valg av beløpstype (for gjenbruk i matcher/GUI)
# ---------------------------------------------------------------------------

def amount_for_account(acc: GLAccount,
                       metric: Optional[AmountMetric] = None,
                       override_default: bool = True) -> Decimal:
    """
    Henter beløp for en konto gitt metric.
    - Hvis metric=None og override_default=True: bruk acc.default_metric()
    - Ellers: UB
    """
    if metric is None and override_default:
        metric = acc.default_metric()
    if metric is None:
        metric = AmountMetric.UB
    return acc.amount(metric)


# ---------------------------------------------------------------------------
# End-of-module
# ---------------------------------------------------------------------------

"""
models.py
===========

This module contains data classes and functions for parsing and summarising
A07 income reports and general ledger (GL) CSV files.  It is designed to be
small and self‑contained so that it can be easily tested and reused in
different contexts.  The goal of this refactoring is to separate data
handling from the graphical user interface, making the codebase easier to
maintain and understand.

The module defines two primary data structures:

* ``A07Row`` – represents a single income line from an A07 JSON file.
* ``GLAccount`` – represents a single account line from a general ledger CSV.

There are also helper functions for parsing files and summarising data.  The
parsers attempt to be tolerant of encoding and delimiter variations while
remaining easy to follow.
"""

from __future__ import annotations

import csv
import io
import json
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple


def _to_float(value: object) -> float:
    """Return a float from a string or number.

    This helper attempts to convert various string formats to a float.
    It strips whitespace, replaces comma with a dot and handles numbers
    enclosed in parentheses as negatives.  If conversion fails the
    function returns 0.0.

    Args:
        value: The input value to convert.

    Returns:
        A floating‑point representation of the input or 0.0 if parsing fails.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    # Normalize various unicode minus signs and thousand separators
    s = s.replace("\xa0", " ")  # non‑breaking space
    s = s.replace("−", "-").replace("–", "-").replace("—", "-")
    # Remove currency symbols and whitespace
    s = re.sub(r"(?i)\b(nok|kr)\b\.?", "", s).strip()
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1].strip()
    # Remove spaces and apostrophes used as thousand separators
    s = s.replace(" ", "").replace("'", "")
    # Replace European decimal comma with dot if necessary
    if "," in s and "." in s:
        # If comma appears after dot, remove dots; else remove commas
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    # Attempt conversion
    try:
        value = float(s)
    except Exception:
        # Extract only digits, dot and minus
        s2 = re.sub(r"[^0-9.\-]", "", s)
        if not s2 or s2 in {"-", "."}:
            return 0.0
        try:
            value = float(s2)
        except Exception:
            return 0.0
    return -value if negative else value


@dataclass
class A07Row:
    """Represents a single income line from an A07 JSON report."""

    orgnr: str
    fnr: str
    navn: str
    kode: str
    fordel: str
    beloep: float
    antall: Optional[int]
    trekkpliktig: bool
    aga: bool
    opptj_start: Optional[str]
    opptj_slutt: Optional[str]


@dataclass
class GLAccount:
    """Represents a single general ledger account from a CSV file."""

    konto: str
    navn: str
    ib: float  # Opening balance
    debet: float
    kredit: float
    endring: float  # Movement / period result
    ub: float  # Closing balance
    belop: float  # Amount used for matching (depends on chosen basis)


class A07Parser:
    """Parser for A07 JSON files.

    The A07 JSON structure is nested and can contain multiple employers
    (virksomhet) and employees (inntektsmottaker).  This parser extracts
    a flat list of ``A07Row`` instances from such a file.  It also
    provides helper methods for summarising data.
    """

    def parse_file(self, path: str) -> Tuple[List[A07Row], List[str]]:
        """Parse an A07 JSON file and return a list of rows and errors.

        Args:
            path: File system path to the JSON file.

        Returns:
            A tuple containing a list of ``A07Row`` instances and a list of
            error messages encountered during parsing.
        """
        rows: List[A07Row] = []
        errors: List[str] = []
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            return [], [f"Kunne ikke lese JSON: {e}"]
        # Navigate to the root object containing income data
        try:
            oppg = (data.get("mottatt", {}) or {}).get("oppgave", {}) or data
            virksomheter = oppg.get("virksomhet") or []
            if isinstance(virksomheter, dict):
                virksomheter = [virksomheter]
            for v in virksomheter:
                orgnr = str(
                    v.get("norskIdentifikator")
                    or v.get("organisasjonsnummer")
                    or v.get("orgnr")
                    or ""
                )
                inntektsmottakere = v.get("inntektsmottaker") or []
                if isinstance(inntektsmottakere, dict):
                    inntektsmottakere = [inntektsmottakere]
                for p in inntektsmottakere:
                    fnr = str(
                        p.get("norskIdentifikator")
                        or p.get("identifikator")
                        or p.get("fnr")
                        or ""
                    )
                    navn = (
                        (p.get("identifiserendeInformasjon") or {}).get("navn")
                        or p.get("navn")
                        or ""
                    )
                    inntekter = p.get("inntekt") or []
                    if isinstance(inntekter, dict):
                        inntekter = [inntekter]
                    for inc in inntekter:
                        try:
                            fordel = str(inc.get("fordel") or "").strip().lower()
                            loenn = inc.get("loennsinntekt") or {}
                            ytelse = inc.get("ytelse") or inc.get("kontantytelse") or {}
                            if not isinstance(loenn, dict):
                                loenn = {}
                            if not isinstance(ytelse, dict):
                                ytelse = {}
                            kode = (
                                loenn.get("beskrivelse")
                                or ytelse.get("beskrivelse")
                                or inc.get("type")
                                or "ukjent_kode"
                            )
                            antall = loenn.get("antall")
                            if not isinstance(antall, (int, float)):
                                antall = None
                            beloep = _to_float(inc.get("beloep"))
                            trekkpliktig = bool(inc.get("inngaarIGrunnlagForTrekk", False))
                            aga = bool(inc.get("utloeserArbeidsgiveravgift", False))
                            opptj_start = inc.get("startdatoOpptjeningsperiode")
                            opptj_slutt = inc.get("sluttdatoOpptjeningsperiode")
                            rows.append(
                                A07Row(
                                    orgnr=orgnr,
                                    fnr=fnr,
                                    navn=str(navn),
                                    kode=str(kode),
                                    fordel=fordel,
                                    beloep=float(beloep),
                                    antall=int(antall) if isinstance(antall, int) else None,
                                    trekkpliktig=trekkpliktig,
                                    aga=aga,
                                    opptj_start=opptj_start,
                                    opptj_slutt=opptj_slutt,
                                )
                            )
                        except Exception as e:
                            errors.append(f"Feil ved parsing av inntektslinje: {e}")
        except Exception as e:
            errors.append(f"Kritisk feil under parsing: {e}")
        return rows, errors

    @staticmethod
    def summarize_by_code(rows: Iterable[A07Row]) -> Dict[str, float]:
        """Summarise income by A07 code.

        Args:
            rows: Iterable of ``A07Row`` objects.

        Returns:
            Dictionary mapping A07 code to the total amount for that code.
        """
        sums: Dict[str, float] = {}
        for r in rows:
            code = str(r.kode)
            sums[code] = sums.get(code, 0.0) + float(r.beloep)
        return sums


def _find_header(fieldnames: List[str], exact: List[str], partial: List[str]) -> Optional[str]:
    """Find a header name from a list of alternatives.

    The function searches for a case‑insensitive exact match first.  If no
    exact match is found it searches for a partial match.  Returns the
    header name from the input list that matches or ``None``.
    """
    mp = {str(h or "").strip().lower(): h for h in fieldnames if h}
    for e in exact:
        if e.lower() in mp:
            return mp[e.lower()]
    for p in partial:
        for n, h in mp.items():
            if p.lower() in n:
                return h
    return None


def read_gl_csv(path: str) -> Tuple[List[GLAccount], Dict[str, str]]:
    """Read a general ledger CSV file and return a list of accounts and metadata.

    The function attempts to detect the delimiter and encoding of the CSV
    automatically.  It then identifies relevant columns such as account
    number, account name, balances and movements.  Amount fields are
    converted to floats using the ``_to_float`` helper.  Missing amounts are
    calculated when possible (e.g. endring from debet minus kredit).

    Args:
        path: Path to a CSV file.

    Returns:
        A tuple containing a list of ``GLAccount`` instances and a dictionary
        with metadata about the file (encoding, delimiter and detected
        column names).
    """
    # Try several encodings; fall back to latin‑1 with replacement
    encodings = ["utf-8-sig", "utf-16", "utf-16le", "utf-16be", "cp1252", "latin-1", "utf-8"]
    text = None
    encoding_used = "utf-8"
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                text = f.read()
                encoding_used = enc
                break
        except UnicodeDecodeError:
            continue
        except Exception:
            continue
    if text is None:
        # Last resort: read with replacement characters
        with open(path, "r", encoding="latin-1", errors="replace") as f:
            text = f.read()
            encoding_used = "latin-1"
    lines = text.splitlines()
    delimiter = None
    if lines and lines[0].strip().lower().startswith("sep="):
        delimiter = lines[0].split("=", 1)[1].strip()[:1] or ";"
        text = "\n".join(lines[1:])
    if delimiter is None:
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t")
            delimiter = dialect.delimiter
        except Exception:
            # Choose the delimiter that occurs most frequently in the sample
            semi = sample.count(";")
            comma = sample.count(",")
            tab = sample.count("\t")
            if tab >= semi and tab >= comma:
                delimiter = "\t"
            elif semi >= comma:
                delimiter = ";"
            else:
                delimiter = ","
    f = io.StringIO(text)
    reader = csv.DictReader(f, delimiter=delimiter)
    records = list(reader)
    fieldnames = reader.fieldnames or (list(records[0].keys()) if records else [])
    if not records:
        return [], {}
    # Detect column names
    acc_hdr = _find_header(fieldnames, ["konto", "kontonummer", "account", "accountno", "kontonr", "account_number"], ["konto", "account"])
    name_hdr = _find_header(fieldnames, ["kontonavn", "navn", "name", "accountname", "beskrivelse", "description", "tekst"], ["navn", "name", "tekst", "desc"])
    ib_hdr = _find_header(fieldnames, ["ib", "inngaaende", "ingaende", "opening_balance"], ["ib", "inng", "open"])
    debet_hdr = _find_header(fieldnames, ["debet", "debit"], ["debet", "debit"])
    kredit_hdr = _find_header(fieldnames, ["kredit", "credit"], ["kredit", "credit"])
    endr_hdr = _find_header(fieldnames, ["endring", "bevegelse", "movement", "ytd", "hittil", "resultat"], ["endr", "beveg", "ytd", "hittil", "period", "result"])
    ub_hdr = _find_header(fieldnames, ["ub", "utgaaende", "utgaende", "closing_balance", "ubsaldo"], ["ub", "utg", "clos"])
    amt_hdr = _find_header(fieldnames, ["saldo", "balance", "belop", "beloep", "beløp", "amount", "sum"], ["saldo", "bel", "amount", "sum"])
    accounts: List[GLAccount] = []
    for r in records:
        konto = (r.get(acc_hdr) if acc_hdr else None) or ""
        navn = (r.get(name_hdr) if name_hdr else None) or ""
        ib = _to_float(r.get(ib_hdr)) if ib_hdr else 0.0
        debet = _to_float(r.get(debet_hdr)) if debet_hdr else 0.0
        kredit = _to_float(r.get(kredit_hdr)) if kredit_hdr else 0.0
        endring = _to_float(r.get(endr_hdr)) if endr_hdr else None
        ub = _to_float(r.get(ub_hdr)) if ub_hdr else None
        belop = _to_float(r.get(amt_hdr)) if amt_hdr else None
        # Compute missing fields
        if endring is None:
            if debet_hdr and kredit_hdr:
                endring = debet - kredit
            elif (ub is not None) and (ib_hdr is not None):
                endring = ub - ib
            else:
                endring = belop if belop is not None else 0.0
        if ub is None:
            if ib_hdr is not None:
                ub = ib + endring
            else:
                ub = belop if belop is not None else endring
        if belop is None:
            belop = ub if ub is not None else endring
        # Populate GLAccount dataclass
        accounts.append(
            GLAccount(
                konto=str(konto).strip(),
                navn=str(navn).strip(),
                ib=float(ib),
                debet=float(debet),
                kredit=float(kredit),
                endring=float(endring),
                ub=float(ub),
                belop=float(belop),
            )
        )
    meta = {
        "encoding": encoding_used,
        "delimiter": delimiter,
        "account_header": acc_hdr,
        "name_header": name_hdr,
        "ib": ib_hdr,
        "debet": debet_hdr,
        "kredit": kredit_hdr,
        "endring": endr_hdr,
        "ub": ub_hdr,
        "amount_header": amt_hdr,
    }
    return accounts, meta


def summarize_gl_by_code(accounts: Iterable[GLAccount], mapping: Dict[str, str], basis: str = "endring") -> Dict[str, float]:
    """Summarise general ledger amounts per A07 code.

    Given a collection of ``GLAccount`` objects and a mapping from account
    numbers to A07 codes, this helper returns the total amount per code.

    Args:
        accounts: Iterable of ``GLAccount`` objects.
        mapping: Dictionary mapping account numbers to A07 codes.
        basis: Which field of ``GLAccount`` to use for the amount.  Valid
            values are ``"endring"``, ``"ub"`` and ``"belop"``.  Defaults to
            ``"endring"``.

    Returns:
        Dictionary mapping A07 code to the summed amount using the chosen basis.
    """
    sums: Dict[str, float] = {}
    for acc in accounts:
        code = mapping.get(acc.konto)
        if not code:
            continue
        if basis == "ub":
            amount = acc.ub
        elif basis == "belop":
            amount = acc.belop
        else:
            amount = acc.endring
        sums[code] = sums.get(code, 0.0) + float(amount)
    return sums

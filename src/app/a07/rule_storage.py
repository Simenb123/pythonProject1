"""
rule_storage.py
-----------------

This module provides data structures and helper functions for persisting and
manipulating A07 rule definitions.  It exposes two simple dataclasses,
``Rule`` and ``RuleBook``, together with convenience methods for loading and
saving rule definitions to JSON files.  The goal of this module is to
encapsulate all logic related to storing, retrieving and modifying rules so
that the rest of the application can focus on business logic.

A ``Rule`` represents the configuration for a single A07‑kode.  It captures
metadata such as the code itself, a human friendly label, which GL basis to
use when comparing beløp, allowed kontointervaller, keywords for text
matching, boost‑konti, expected sign, and any special additions.  See
``a07_rulebook.py`` for further details on the meaning of each field.

Rules are grouped into a ``RuleBook`` which internally stores a mapping
from A07‑kode to ``Rule`` objects and a mapping of canonical terms to their
aliases.  ``RuleBook`` exposes methods to add, update and delete rules,
as well as import/export the entire structure to/from disk.

Example usage::

    from rule_storage import Rule, RuleBook

    # Create a new rule book and add a rule
    book = RuleBook()
    book.add_rule(Rule(
        code="fastloenn",
        label="Fastlønn",
        category="wage",
        basis="endring",
        allowed_ranges=["5000-5999"],
        keywords=["fastlønn", "bonus"],
        boost_accounts=["5010"],
        expected_sign=1,
    ))
    # Save to JSON
    book.save_to_json("my_rules.json")

    # Load existing rules
    loaded = RuleBook.load_from_json("my_rules.json")
    # Update a rule
    loaded.update_rule("fastloenn", expected_sign=-1)
    loaded.save_to_json("my_rules.json")

This design intentionally keeps the storage format very simple – a flat JSON
file – to ensure that rules can be reviewed and edited outside of the
application if desired.  Because rules are stored as dataclasses, adding
new fields is straightforward and does not affect existing code consumers
so long as default values are provided.

"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, List, Optional, Iterable


@dataclass
class Rule:
    """Represents configuration for a single A07‑kode.

    Attributes:
        code:            The A07 code (primary key).
        label:           A human friendly label for the code (optional).
        category:        A higher level grouping (e.g. wage, travel, deduction).
        basis:           Which GL basis to use when comparing beløp; one of
                         "auto", "endring", "ub" or "belop".
        allowed_ranges:  A list of kontointervall‑expressions.  Each entry
                         should be a string on the form ``"5000-5999"`` or
                         ``"2940"``.  Parsing into numeric ranges is left
                         to higher level logic.
        keywords:        A list of phrases or tokens which should raise the
                         matching score when found in a kontonavn.
        boost_accounts:  A list of kontonummer which explicitly map to this
                         code even if they lie outside ``allowed_ranges``.
        expected_sign:   +1 if beløp is normally positive for this code,
                         -1 if beløp is normally negative, 0 otherwise.
        special_add:     A free form list of dicts used to store additional
                         configuration (for example combining codes).
    """

    code: str
    label: str = ""
    category: str = "wage"
    basis: str = "auto"
    allowed_ranges: List[str] = field(default_factory=list)
    keywords: List[str] = field(default_factory=list)
    boost_accounts: List[str] = field(default_factory=list)
    expected_sign: int = 0
    special_add: List[Any] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Serialize this rule to a plain Python dict."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Rule":
        """Instantiate a rule from a plain dictionary.

        Unknown keys are ignored and missing keys take default values.
        """
        return cls(
            code=str(data.get("code", "")).strip(),
            label=str(data.get("label", "")).strip(),
            category=str(data.get("category", "wage")),
            basis=str(data.get("basis", "auto")),
            allowed_ranges=list(data.get("allowed_ranges", [])),
            keywords=list(data.get("keywords", [])),
            boost_accounts=list(data.get("boost_accounts", [])),
            expected_sign=int(data.get("expected_sign", 0) or 0),
            special_add=list(data.get("special_add", [])),
        )


@dataclass
class RuleBook:
    """Container for storing rules and their aliases.

    The ``rules`` attribute maps an A07 code to a ``Rule``.  The ``aliases``
    attribute maps canonical terms (lower case) to a set of synonymous
    spellings.  Together they allow the rest of the application to resolve
    textual inputs into known codes.  ``source`` is an optional string
    describing where the rulebook was loaded from (e.g. file path).
    """

    rules: Dict[str, Rule] = field(default_factory=dict)
    aliases: Dict[str, List[str]] = field(default_factory=dict)
    source: str = ""

    # ------------------------------------------------------------------
    # Rule management
    #
    def add_rule(self, rule: Rule) -> None:
        """Add or replace a rule by its code."""
        self.rules[rule.code] = rule

    def update_rule(self, code: str, **updates: Any) -> None:
        """Update the fields of an existing rule.

        Any key in ``updates`` that matches a field on ``Rule`` will be
        overwritten.  A ``KeyError`` is raised if the rule does not exist.
        """
        if code not in self.rules:
            raise KeyError(f"Rule {code} does not exist")
        r = self.rules[code]
        for k, v in updates.items():
            if hasattr(r, k):
                setattr(r, k, v)

    def delete_rule(self, code: str) -> None:
        """Remove a rule from the book if present."""
        self.rules.pop(code, None)

    def get_rule(self, code: str) -> Optional[Rule]:
        """Return the rule for a given code or ``None`` if not found."""
        return self.rules.get(code)

    # ------------------------------------------------------------------
    # Alias management
    #
    def add_alias(self, canonical: str, synonym: str) -> None:
        """Add a synonym for a given canonical term.

        Both canonical and synonym are normalised to lower case.  The synonym
        will be appended to the list of existing synonyms if not already
        present.
        """
        can = str(canonical).strip().lower()
        syn = str(synonym).strip().lower()
        if not can or not syn:
            return
        self.aliases.setdefault(can, [])
        if syn not in self.aliases[can]:
            self.aliases[can].append(syn)

    def delete_alias(self, canonical: str, synonym: Optional[str] = None) -> None:
        """Delete a synonym or an entire alias group.

        If only ``canonical`` is provided, the entire group is removed.  If
        ``synonym`` is provided, only that entry is removed.
        """
        can = str(canonical).strip().lower()
        if can not in self.aliases:
            return
        if synonym is None:
            # remove the whole group
            self.aliases.pop(can, None)
        else:
            syn = str(synonym).strip().lower()
            if syn in self.aliases.get(can, []):
                self.aliases[can].remove(syn)
            if not self.aliases[can]:
                self.aliases.pop(can, None)

    # ------------------------------------------------------------------
    # Persistence helpers
    #
    @classmethod
    def load_from_json(cls, path: str) -> "RuleBook":
        """Load a rulebook from a JSON file.

        The JSON structure is expected to have the form::

            {
              "rules": {
                "fastloenn": {
                    "code": "fastloenn",
                    "label": "Fastlønn",
                    ...
                },
                ...
              },
              "aliases": {
                "ferie": ["feriepenger", "feriepenga"],
                ...
              },
              "source": "my_rules.json"
            }

        Unknown keys at the top level are ignored.
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"No such file: {path}")
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        rules_data = data.get("rules", {})
        aliases = data.get("aliases", {})
        source = data.get("source", path)
        rules: Dict[str, Rule] = {}
        for code, rdata in rules_data.items():
            rule = Rule.from_dict(rdata)
            # ensure the code key matches the rule.code field
            rules[rule.code or code] = rule
        # normalise alias values to lists of strings
        alias_map: Dict[str, List[str]] = {}
        for can, syns in aliases.items():
            if isinstance(syns, list):
                alias_map[str(can).lower()] = [str(s).lower() for s in syns]
        return cls(rules=rules, aliases=alias_map, source=source)

    def save_to_json(self, path: str) -> None:
        """Serialize this rulebook to JSON on disk.

        The resulting JSON will contain the dictionaries ``rules`` and
        ``aliases``, as well as a ``source`` field referencing the file path.
        The directory is created if it does not exist.
        """
        obj = {
            "rules": {code: rule.to_dict() for code, rule in self.rules.items()},
            "aliases": self.aliases,
            "source": self.source or path,
        }
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)

    # ------------------------------------------------------------------
    # Convenience API for iteration
    #
    def __iter__(self) -> Iterable[Rule]:
        """Iterate over all rules in the book."""
        return iter(self.rules.values())

    def __len__(self) -> int:
        return len(self.rules)

    def __contains__(self, code: str) -> bool:
        return code in self.rules

    # ------------------------------------------------------------------
    # Export helpers
    #
    def export_to_excel(self, path: str) -> None:
        """Export the rulebook to an Excel file with two sheets.

        The first sheet ``a07_codes`` will contain one row per A07 code with
        the following columns: ``a07_code``, ``label``, ``category``,
        ``basis``, ``allowed_ranges``, ``keywords``, ``boost_accounts``,
        ``special_add`` and ``expected_sign``.  The second sheet
        ``aliases`` will contain two columns: ``canonical`` and
        ``synonyms`` where synonyms is a comma‑separated string of
        alternative spellings for the canonical term.

        Requires the optional dependency ``pandas`` with an Excel writer
        engine (such as ``openpyxl``).  If pandas is not available, a
        ``RuntimeError`` will be raised.
        """
        try:
            import pandas as pd  # type: ignore
        except Exception:
            raise RuntimeError("Excel-export requires 'pandas' and an appropriate engine (e.g. openpyxl)")
        import json as _json
        # Build the codes DataFrame
        data_codes: List[Dict[str, Any]] = []
        for code, r in self.rules.items():
            data_codes.append({
                "a07_code": code,
                "label": r.label,
                "category": r.category,
                "basis": r.basis,
                # join allowed_ranges with ' | ' to match original CSV format
                "allowed_ranges": " | ".join(r.allowed_ranges or []),
                "keywords": ", ".join(r.keywords or []),
                "boost_accounts": ", ".join(r.boost_accounts or []),
                # dump special_add to JSON string for readability
                "special_add": _json.dumps(r.special_add or [], ensure_ascii=False),
                "expected_sign": r.expected_sign,
            })
        df_codes = pd.DataFrame(data_codes)
        # Build the aliases DataFrame
        alias_rows: List[Dict[str, Any]] = []
        for can, syns in self.aliases.items():
            alias_rows.append({
                "canonical": can,
                "synonyms": ", ".join(syns or []),
            })
        df_alias = pd.DataFrame(alias_rows)
        # Write to Excel file
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with pd.ExcelWriter(path) as writer:
            df_codes.to_excel(writer, sheet_name="a07_codes", index=False)
            df_alias.to_excel(writer, sheet_name="aliases", index=False)

"""
Utilities for normalising saldobalanse and generating regnskapsoppstilling.

This module centralises common logic needed when working with trial balance
(saldobalanse) and regnskapslinjer.  It includes:

* ``rename_balance_columns`` – Rename incoming column names such as ``Inngående
  saldo`` or ``Utgående balanse`` to the standard names ``IB`` (inngående
  balanse), ``UB`` (utgående balanse) and ``Endring`` (endring i perioden).
* ``summarize_regnskap`` – Given a mapped saldobalanse with regnskapsnumre
  (``regnr``) and balansekolonner (``IB``, ``Endring``, ``UB``), it
  aggregates these per regnskapslinje for use in a regnskapsoppstilling.

These helper functions can be imported and used by both GUI- og service
moduler.  They operate on pandas DataFrames and return new DataFrames, so
they do not mutate the input.
"""

from __future__ import annotations

import pandas as pd
from typing import Iterable, Tuple


# Mapping of known column synonyms to the standard names used by the
# application.  Entries are lowercased and whitespace is removed when
# comparing.
_COLUMN_SYNONYMS: dict[str, str] = {
    # Inngående balanse / saldo → IB
    "inngåendesaldo": "IB",
    "inngåendebalanse": "IB",
    "inngående saldo": "IB",
    "inngående balanse": "IB",
    "ib": "IB",
    # Utgående balanse / saldo → UB
    "utgåendesaldo": "UB",
    "utgåendebalanse": "UB",
    "utgående saldo": "UB",
    "utgående balanse": "UB",
    "ub": "UB",
    # Endring / Bevegelse → Endring
    "endring": "Endring",
    "bevegelse": "Endring",
    "endring i perioden": "Endring",
    "bevegelse i perioden": "Endring",
}


def _normalise_name(name: str) -> str:
    """Normalise a column name for comparison: lower-case and remove spaces."""
    return str(name or "").strip().lower().replace(" ", "").replace("\u00a0", "")


def rename_balance_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of ``df`` with balance columns renamed to standard names.

    The function scans the columns of ``df`` for names that match any of the
    entries in ``_COLUMN_SYNONYMS`` (ignoring case and whitespace).  When a
    match is found, the column is renamed to the corresponding standard name
    (``IB``, ``UB`` or ``Endring``).  Columns that do not match any synonym
    remain unchanged.  The returned DataFrame has the same order of
    columns.

    Parameters
    ----------
    df : pandas.DataFrame
        The input DataFrame.

    Returns
    -------
    pandas.DataFrame
        A new DataFrame with renamed columns.
    """
    rename_map: dict[str, str] = {}
    for c in df.columns:
        norm = _normalise_name(c)
        if norm in _COLUMN_SYNONYMS:
            rename_map[c] = _COLUMN_SYNONYMS[norm]
    return df.rename(columns=rename_map, inplace=False)


def summarize_regnskap(df: pd.DataFrame,
                       ib_col: str = "IB",
                       endring_col: str = "Endring",
                       ub_col: str = "UB",
                       regnr_col: str = "regnr",
                       regnskapslinje_col: str = "regnskapslinje"
                       ) -> pd.DataFrame:
    """Aggregate a mapped trial balance by regnskapslinje.

    Given a DataFrame that includes regnskapsnummer (``regnr``) and the
    balance columns ``IB``, ``Endring`` and ``UB``, this function groups by
    ``regnr`` (and includes the regnskapslinje name if available) and sums
    the balances.  Rows with missing ``regnr`` are ignored in the summary.

    Parameters
    ----------
    df : pandas.DataFrame
        The input DataFrame containing at least the columns specified by
        ``ib_col``, ``endring_col``, ``ub_col`` and ``regnr_col``.
    ib_col : str, optional
        Name of the column containing inngående balanse.  Default is ``"IB"``.
    endring_col : str, optional
        Name of the column containing endring i perioden.  Default is ``"Endring"``.
    ub_col : str, optional
        Name of the column containing utgående balanse.  Default is ``"UB"``.
    regnr_col : str, optional
        Name of the column containing regnskapsnummer.  Default is ``"regnr"``.
    regnskapslinje_col : str, optional
        Name of the column containing the regnskapslinje text.  Default is
        ``"regnskapslinje"``.

    Returns
    -------
    pandas.DataFrame
        A DataFrame with one row per regnskapsnummer, containing the
        aggregated balances and the regnskapslinje name (if available).
    """
    # Filter out rows without regnr
    if regnr_col not in df.columns:
        raise ValueError(f"DataFrame mangler kolonnen '{regnr_col}'.")
    work = df.dropna(subset=[regnr_col]).copy()
    # Ensure numeric types for balances
    for col in (ib_col, endring_col, ub_col):
        if col in work.columns:
            work[col] = pd.to_numeric(work[col], errors="coerce")
        else:
            # If any of the expected balance columns is missing, create it with zeros
            work[col] = 0.0
    # Group and aggregate
    group_cols = [regnr_col]
    if regnskapslinje_col in work.columns:
        group_cols.append(regnskapslinje_col)
    grouped = work.groupby(group_cols, dropna=True, as_index=False)[[ib_col, endring_col, ub_col]].sum()
    return grouped
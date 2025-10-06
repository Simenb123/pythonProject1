
from __future__ import annotations

from typing import Iterable, Optional, Tuple, Union, Dict, List
import numpy as np
import pandas as pd


def map_kontoplan_df(
    df_kilde: pd.DataFrame,
    mapping_df: pd.DataFrame,
    konto_col: str,
    out_cols: Tuple[str, str] = ("map_val1", "map_val2"),
    mapping_cols: Optional[Dict[str, str]] = None,
    assume_sorted: bool = False,
    validate: bool = True,
    skip_existing: bool = False,
) -> pd.DataFrame:
    """
    Mapper kontonumre i df_kilde[konto_col] mot intervaller i mapping_df.

    mapping_df må ha fire felt: 'fra', 'til', 'val1', 'val2' – enten direkte eller via alias
    gjennom mapping_cols (f.eks. {"fra": "StartKonto", "til": "SluttKonto", "val1": "Regnnr.", "val2": "Regnskapslinje"}).

    Parameters
    ----------
    df_kilde : DataFrame
        Saldobalanse/klientdata med kolonnen som inneholder konto (konto_col).
    mapping_df : DataFrame
        Grunnlag med intervaller og verdier.
    konto_col : str
        Navnet på kolonnen i df_kilde som inneholder kontonummer (f.eks. "konto").
    out_cols : Tuple[str, str]
        Navn på nye kolonner som skal inneholde mappede verdier fra mapping_df (val1, val2).
    mapping_cols : Optional[Dict[str, str]]
        Alias for kolonnenavn i mapping_df. Nøkler: 'fra','til','val1','val2'.
    assume_sorted : bool
        Sett True hvis mapping_df allerede er sortert stigende på 'fra'.
    validate : bool
        Validerer intervaller (fra<=til, sortering, overlapp).
    skip_existing : bool
        Overskriv ikke eksisterende verdier i out_cols dersom de allerede er utfylt (ikke-null).

    Returns
    -------
    DataFrame
        Kopi av df_kilde med to nye kolonner out_cols.
    """
    # Map alias -> faktiske kolonnenavn
    if mapping_cols is None:
        mapping_cols = {"fra": "fra", "til": "til", "val1": "val1", "val2": "val2"}
    required_keys = {"fra", "til", "val1", "val2"}
    missing_keys = required_keys - set(mapping_cols.keys())
    if missing_keys:
        raise ValueError(f"mapping_cols mangler nøkler: {sorted(missing_keys)}")

    _require_columns(mapping_df, [mapping_cols[k] for k in ["fra", "til", "val1", "val2"]])
    if konto_col not in df_kilde.columns:
        raise ValueError(f"Fant ikke kolonnen '{konto_col}' i df_kilde.")

    m = mapping_df[[mapping_cols["fra"], mapping_cols["til"], mapping_cols["val1"], mapping_cols["val2"]]].copy()
    m.columns = ["fra", "til", "val1", "val2"]

    # Normaliser typer
    m["fra"] = _to_numeric(m["fra"])
    m["til"] = _to_numeric(m["til"])

    if not assume_sorted:
        m = m.sort_values("fra", kind="mergesort", ignore_index=True)

    if validate:
        _validate_intervals(m)

    fra = m["fra"].to_numpy(dtype=float)
    til = m["til"].to_numpy(dtype=float)
    v1 = m["val1"].to_numpy()
    v2 = m["val2"].to_numpy()

    # Kontoer
    if konto_col not in df_kilde.columns:
        raise ValueError(f"Konto-kolonne '{konto_col}' finnes ikke i df_kilde.")
    konto_vals = _to_numeric(df_kilde[konto_col])
    k = konto_vals.to_numpy(dtype=float, copy=False)

    idx = np.searchsorted(fra, k, side="right") - 1
    valid_mask = (idx >= 0) & (k <= til[np.clip(idx, 0, len(til) - 1)])
    safe_idx = np.where(valid_mask, idx, 0)

    out1 = np.where(valid_mask, v1[safe_idx], None)
    out2 = np.where(valid_mask, v2[safe_idx], None)

    res = df_kilde.copy()
    # Lag kolonner hvis de ikke finnes fra før
    if out_cols[0] not in res.columns:
        res[out_cols[0]] = None
    if out_cols[1] not in res.columns:
        res[out_cols[1]] = None

    if skip_existing:
        # Bare skriv der det er NaN/None i eksisterende kolonner
        mask1 = res[out_cols[0]].isna()
        mask2 = res[out_cols[1]].isna()
        res.loc[mask1, out_cols[0]] = np.array(out1, dtype=object)[mask1.to_numpy()]
        res.loc[mask2, out_cols[1]] = np.array(out2, dtype=object)[mask2.to_numpy()]
    else:
        res[out_cols[0]] = out1
        res[out_cols[1]] = out2

    return res


def map_kontoplan_excel(
    kilde_xlsx: str,
    mapping_xlsx: str,
    kilde_sheet: Union[int, str] = 0,
    mapping_sheet: Union[int, str] = 0,
    konto_col: str = "konto",
    out_cols: Tuple[str, str] = ("map_val1", "map_val2"),
    mapping_cols: Optional[Dict[str, str]] = None,
    write_to: Optional[str] = None,
    write_sheet: Optional[str] = None,
    assume_sorted: bool = False,
    validate: bool = True,
    skip_existing: bool = False,
) -> pd.DataFrame:
    """
    Fil-basert hjelper som speiler VBA-flyten.

    Forventet mapping-ark (med overskrifter), via mapping_cols:
      'fra' | 'til' | 'val1' | 'val2'

    Eksempel med dine kolonnenavn i mapping-filen:
      mapping_cols={"fra":"StartKonto","til":"SluttKonto","val1":"Regnnr.","val2":"Regnskapslinje"}
    """
    df_kilde = pd.read_excel(kilde_xlsx, sheet_name=kilde_sheet, dtype=object)
    df_map = pd.read_excel(mapping_xlsx, sheet_name=mapping_sheet, dtype=object)

    res = map_kontoplan_df(
        df_kilde=df_kilde,
        mapping_df=df_map,
        konto_col=konto_col,
        out_cols=out_cols,
        mapping_cols=mapping_cols,
        assume_sorted=assume_sorted,
        validate=validate,
        skip_existing=skip_existing,
    )

    if write_to:
        sheet = write_sheet or "Mapped"
        with pd.ExcelWriter(write_to, engine="openpyxl") as xw:
            res.to_excel(xw, index=False, sheet_name=sheet)

    return res


def _require_columns(df: pd.DataFrame, cols: Iterable[str]) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Mangler kolonner: {missing}. Fant: {list(df.columns)}")


def _to_numeric(s: pd.Series) -> pd.Series:
    """Konverterer serie til numerisk (float), ikke-numeriske -> NaN."""
    return pd.to_numeric(s, errors="coerce")


def _validate_intervals(m: pd.DataFrame) -> None:
    if (m["fra"] > m["til"]).any():
        bad = m.loc[m["fra"] > m["til"]]
        raise ValueError(
            "Fant rader med fra > til i mapping_df. Eksempler:\n"
            f"{bad.head(3).to_string(index=False)}"
        )
    if not (m["fra"].is_monotonic_increasing):
        raise ValueError(
            "Intervallene må være sortert stigende på 'fra'. Sett assume_sorted=False for auto-sortering."
        )
    # Ikke-overlapp-sjekk: neste fra må være > forrige til
    til_shift = m["til"].shift(fill_value=-np.inf)
    if (m["fra"] <= til_shift).iloc[1:].any():
        raise ValueError("Intervallene ser ut til å overlappe. Rydd opp i grunnlagsfilen.")


if __name__ == "__main__":
    # Eksempel – tilpass for lokal kjøring
    KILDE = r"C:\path\til\saldobalanse.xlsx"
    GRUNNLAG = r"C:\Users\ib91\Desktop\Prosjekt\Kildefiler\Mapping standard kontoplan.xlsx"
    OUT = r"C:\path\til\saldobalanse_mapped.xlsx"

    mapping_alias = {
        "fra": "StartKonto",
        "til": "SluttKonto",
        "val1": "Regnnr.",
        "val2": "Regnskapslinje",
    }

    try:
        result = map_kontoplan_excel(
            kilde_xlsx=KILDE,
            mapping_xlsx=GRUNNLAG,
            kilde_sheet=0,
            mapping_sheet="Sheet1",
            konto_col="konto",  # som du oppga
            out_cols=("Regnr_mapped", "Regnskapslinje_mapped"),
            mapping_cols=mapping_alias,
            write_to=OUT,
            write_sheet="Mapped",
            assume_sorted=False,   # sorter mapping automatisk ved behov
            validate=True,         # streng, men trygg
            skip_existing=False,   # sett True om du etterhvert vil unngå å overskrive
        )
        print(f"Skrev resultat til: {OUT}  (rader: {len(result)})")
    except Exception as e:
        print(f"Feil: {e}")

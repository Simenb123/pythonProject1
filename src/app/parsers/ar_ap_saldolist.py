"""
ar_ap_saldolist.py – lag en samlet saldobalanse for kunder (AR) og leverandører (AP).

Dette skriptet bygger på run_saft_pro_gui.make_subledger for å generere
reskontro for både kunder og leverandører i én operasjon, og i tillegg lage
et sammendrag og en avstemmingsrapport som viser om reskontrobeløpene
stemmer med kontoplanens utgående saldo. Hvis accounts.csv mangler, hentes
closing‑netto fra trial_balance.xlsx.

Kjør som modul:
    from ar_ap_saldolist import generate_saldolist
    generate_saldolist(Path('output_dir'))

Eller fra kommandolinje:
    python ar_ap_saldolist.py [outdir] [--date_from YYYY-MM-DD] [--date_to YYYY-MM-DD]

Hvis outdir ikke angis, brukes nåværende arbeidskatalog.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Set, List
import pandas as pd

from run_saft_pro_gui import (
    make_subledger,
    _read_csv_safe,
    _to_num,
    _parse_dates,
    _norm_acc,
    _has_value,
    AR_CONTROL_ACCOUNTS,
    AP_CONTROL_ACCOUNTS,
    _compute_target_closing,
    _find_csv_file,
)


def _sum_tx_by_account(
    outdir: Path,
    which: str,
    dto: pd.Timestamp,
    ctrl_accounts: Set[str],
    with_party: bool,
) -> pd.DataFrame:
    """Summerer Debit - Credit per kontrollkonto til og med dto.

    Parametre:
        outdir: mappe med transactions.csv
        which: "AR" for kunder eller "AP" for leverandører
        dto: dato for siste transaksjon som skal tas med
        ctrl_accounts: sett med reskontrokonti
        with_party: True for å summere kun linjer med part-ID, False for
            å summere kun linjer uten part-ID (partyless)

    Returnerer en DataFrame med kolonnene AccountID og Amount.
    """
    tx = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    if tx is None or tx.empty or not ctrl_accounts:
        return pd.DataFrame(columns=["AccountID", "Amount"])
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    for col in ["AccountID", "CustomerID", "SupplierID"]:
        if col in tx.columns:
            tx[col] = tx[col].astype(str)
    tx["AccountID"] = tx["AccountID"].map(_norm_acc)
    tx = _to_num(tx, ["Debit", "Credit"])
    # Filtrer på dato
    tx = tx.loc[~tx["Date"].isna() & (tx["Date"] <= dto)].copy()
    # Filtrer på kontrollkontoer
    tx = tx.loc[tx["AccountID"].isin(ctrl_accounts)].copy()
    # Filtrer på part-ID
    if which.upper() == "AR":
        mask = _has_value(tx.get("CustomerID", pd.Series([], dtype=str)))
    else:
        mask = _has_value(tx.get("SupplierID", pd.Series([], dtype=str)))
    tx = tx.loc[mask if with_party else ~mask].copy()
    if tx.empty:
        return pd.DataFrame(columns=["AccountID", "Amount"])
    grp = tx.groupby("AccountID")[["Debit", "Credit"]].sum().reset_index()
    grp["Amount"] = grp["Debit"] - grp["Credit"]
    return grp[["AccountID", "Amount"]]


def generate_saldolist(
    outdir: Path, date_from: Optional[str] = None, date_to: Optional[str] = None
) -> Path:
    """Lag samlet AR/AP saldoliste og avstemming i angitt mappe.

    Skriptet genererer først subledger for både AR og AP ved hjelp av
    run_saft_pro_gui.make_subledger. Deretter leses balansearkene og et
    sammendrag samt en avstemmingsrapport lages. Rapporten skrives til
    ar_ap_saldolist.xlsx og en CSV-fil med samme innhold.

    outdir må inneholde SAF‑T‑CSV‑filer (transactions.csv, accounts.csv osv.).
    """
    # Generer subledger for AR og AP
    ar_path = make_subledger(outdir, "AR", date_from, date_to)
    ap_path = make_subledger(outdir, "AP", date_from, date_to)
    # Les balansearkene
    ar_df = pd.read_excel(ar_path, sheet_name="AR_Balances")
    ap_df = pd.read_excel(ap_path, sheet_name="AP_Balances")
    # Sammendrag: summer UB, IB og PR
    summary = pd.DataFrame(
        [
            {
                "Type": "AR",
                "Sum_UB": ar_df["UB_Amount"].sum(),
                "Sum_IB": ar_df["IB_Amount"].sum(),
                "Sum_PR": ar_df["PR_Amount"].sum(),
            },
            {
                "Type": "AP",
                "Sum_UB": ap_df["UB_Amount"].sum(),
                "Sum_IB": ap_df["IB_Amount"].sum(),
                "Sum_PR": ap_df["PR_Amount"].sum(),
            },
        ]
    )
    # Bestem dato for avstemming
    # Finn transactions.csv
    tx_path = _find_csv_file(outdir, "transactions.csv")
    tx = _read_csv_safe(tx_path, dtype=str) if tx_path else None
    if tx is None or tx.empty:
        raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    dto = pd.to_datetime(date_to) if date_to else tx["Date"].dropna().max()
    # Avstemmingsrapport
    rec_rows: List[dict] = []
    for typ, ctrl in [("AR", AR_CONTROL_ACCOUNTS), ("AP", AP_CONTROL_ACCOUNTS)]:
        # Closing net fra kontoplanen eller trial balance
        closing_net = _compute_target_closing(outdir, ctrl)
        closing = closing_net if closing_net is not None else 0.0
        # Sum UB fra reskontro
        res = ar_df["UB_Amount"].sum() if typ == "AR" else ap_df["UB_Amount"].sum()
        # Sum av partyless transaksjoner
        partless = _sum_tx_by_account(outdir, typ, dto, ctrl, with_party=False)[
            "Amount"
        ].sum()
        rec_rows.append(
            {
                "Type": typ,
                "ControlAccounts": ", ".join(sorted(ctrl)),
                "ClosingNet": closing,
                "Reskontro": res,
                "Partyless": partless,
                "Difference": closing - res - partless,
            }
        )
    rec_df = pd.DataFrame(rec_rows)
    # Skriv Excel og CSV
    out_path = outdir / "ar_ap_saldolist.xlsx"
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        ar_df.to_excel(writer, index=False, sheet_name="Customers_UB")
        ap_df.to_excel(writer, index=False, sheet_name="Suppliers_UB")
        summary.to_excel(writer, index=False, sheet_name="Summary")
        rec_df.to_excel(writer, index=False, sheet_name="Reconciliation")
    # Lag også CSV for begge listene med en Type-kolonne
    csv_path = outdir / "ar_ap_saldolist.csv"
    ar_df2 = ar_df.copy()
    ar_df2.insert(0, "Type", "AR")
    ap_df2 = ap_df.copy()
    ap_df2.insert(0, "Type", "AP")
    combined = pd.concat([ar_df2, ap_df2], ignore_index=True)
    combined.to_csv(csv_path, index=False)
    return out_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate combined AR/AP saldoliste and reconciliation from SAF‑T data"
    )
    parser.add_argument(
        "outdir",
        type=str,
        nargs="?",
        default=".",
        help=(
            "Directory containing SAF‑T CSV extracts. If omitted, the current working directory is used."
        ),
    )
    parser.add_argument(
        "--date_from", type=str, default=None, help="Start date (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--date_to", type=str, default=None, help="End date (YYYY-MM-DD)"
    )
    args = parser.parse_args()
    generate_saldolist(Path(args.outdir), args.date_from, args.date_to)
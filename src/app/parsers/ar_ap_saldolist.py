# -*- coding: utf-8 -*-
from __future__ import annotations
"""
ar_ap_saldolist.py — Oppdatert:
- Periode: støtte for SelectionStartDate/SelectionEndDate (v1.3)
- Strammere AR-heuristikk (kun 15xx; ikke endswith('10'))
"""

from pathlib import Path
from typing import Optional, Tuple, Set, Iterable
import pandas as pd

MONEY_COLS = ["IB_CF","IB_Amount","PR_Amount","UB_Amount","DiffCheck"]

def _read_csv_safe(path: Path, dtype=None) -> Optional[pd.DataFrame]:
    try:
        return pd.read_csv(path, dtype=dtype, keep_default_na=False)
    except Exception:
        return None

def _to_num(df: pd.DataFrame, cols: Iterable[str]):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df

def _parse_dates(df: pd.DataFrame, cols: Iterable[str]):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def _has_value(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.strip().str.lower()
    return ~t.isin(["", "nan", "none", "nat"])

def _norm_acc(acc) -> str:
    s = str(acc).strip()
    if s.endswith(".0"): s = s[:-2]
    s = s.lstrip("0") or "0"
    return s

def _norm_acc_series(s: pd.Series) -> pd.Series:
    return s.apply(_norm_acc)

def _dominant_year(tx: Optional[pd.DataFrame]) -> Optional[int]:
    if tx is None or "Date" not in tx.columns or tx.empty: return None
    years = tx["Date"].dt.year.dropna().astype(int)
    return int(years.value_counts().idxmax()) if not years.empty else None

def _range_dates(outdir: Path, date_from: Optional[str], date_to: Optional[str],
                 tx: Optional[pd.DataFrame]) -> Tuple[pd.Timestamp, pd.Timestamp]:
    dfrom = pd.to_datetime(date_from) if date_from else None
    dto   = pd.to_datetime(date_to) if date_to else None
    if dfrom is None or dto is None:
        hdr = _read_csv_safe(outdir / "header.csv", dtype=str)
        if hdr is not None and not hdr.empty:
            row = hdr.iloc[0]
            dfrom = dfrom or pd.to_datetime(
                row.get("SelectionStart") or row.get("SelectionStartDate") or row.get("StartDate"),
                errors="coerce"
            )
            dto   = dto   or pd.to_datetime(
                row.get("SelectionEnd") or row.get("SelectionEndDate") or row.get("EndDate"),
                errors="coerce"
            )
    if (dfrom is None or pd.isna(dfrom) or dto is None or pd.isna(dto)) and tx is not None:
        year = _dominant_year(tx)
        if year is not None:
            dfrom = pd.Timestamp(year=year, month=1, day=1)
            dto   = pd.Timestamp(year=year, month=12, day=31)
    if dfrom is None or pd.isna(dfrom): dfrom = pd.Timestamp.min
    if dto is None or pd.isna(dto): dto = pd.Timestamp.max
    return dfrom.normalize(), dto.normalize()

def _ctrl_accounts_from_v13(arap_ctrl: Optional[pd.DataFrame], which: str) -> Optional[Set[str]]:
    if arap_ctrl is None or arap_ctrl.empty or "PartyType" not in arap_ctrl.columns or "AccountID" not in arap_ctrl.columns:
        return None
    mask = arap_ctrl["PartyType"].eq("Customer" if which=="AR" else "Supplier")
    s = arap_ctrl.loc[mask, "AccountID"].dropna().astype(str).map(_norm_acc)
    accs = set(s.tolist())
    return accs if accs else None

def _heuristic_accounts(accounts_df: Optional[pd.DataFrame], tx: pd.DataFrame, which: str) -> Set[str]:
    accs: Set[str] = set()
    if accounts_df is not None and not accounts_df.empty:
        acc_tmp = accounts_df.copy()
        acc_tmp["__acc"] = acc_tmp.get("AccountID", "").astype(str).map(_norm_acc)
        desc_col = next((c for c in ["AccountDescription","Description","Name"] if c in acc_tmp.columns), None)
        acc_tmp["__desc"] = acc_tmp[desc_col].astype(str).str.lower() if desc_col else ""
        if which == "AR":
            accs |= set(acc_tmp.loc[acc_tmp["__acc"].str.startswith("15"), "__acc"])
            if desc_col:
                mask = acc_tmp["__desc"].str.contains("kunde|kundefordr|accounts receiv|customer", regex=True, na=False)
                accs |= set(acc_tmp.loc[mask, "__acc"])
        else:
            accs |= set(acc_tmp.loc[acc_tmp["__acc"].str.startswith("24"), "__acc"])
            if desc_col:
                mask = acc_tmp["__desc"].str.contains("leverand|accounts pay|supplier|creditor", regex=True, na=False)
                accs |= set(acc_tmp.loc[mask, "__acc"])
    # observerte konti (der partyID faktisk brukes)
    if which == "AR":
        obs = tx.loc[_has_value(tx.get("CustomerID","")), "AccountID"].astype(str).map(_norm_acc)
    else:
        obs = tx.loc[_has_value(tx.get("SupplierID","")), "AccountID"].astype(str).map(_norm_acc)
    accs |= set(obs.dropna().tolist())
    # siste filter
    if which == "AR":
        accs = {a for a in accs if a.startswith("15")}
    else:
        accs = {a for a in accs if a.startswith("24") or a in {"2410","2460"}}
    return accs

def _pick_control_accounts(outdir: Path, which: str, tx_all: pd.DataFrame) -> Set[str]:
    arap_ctrl = _read_csv_safe(outdir / "arap_control_accounts.csv", dtype=str)
    accounts_df = _read_csv_safe(outdir / "accounts.csv", dtype=str)
    from_v13 = _ctrl_accounts_from_v13(arap_ctrl, which)
    if from_v13:
        return from_v13
    return _heuristic_accounts(accounts_df, tx_all, which)

def _ib_cf_from_ctrl(arap_ctrl: Optional[pd.DataFrame], which: str) -> Optional[pd.DataFrame]:
    if arap_ctrl is None or arap_ctrl.empty: return None
    need = {"OpeningDebit","OpeningCredit","PartyID","PartyType"}
    if not need.issubset(set(arap_ctrl.columns)): return None
    tmp = arap_ctrl.loc[arap_ctrl["PartyType"].eq("Customer" if which=="AR" else "Supplier")].copy()
    tmp = _to_num(tmp, ["OpeningDebit","OpeningCredit"])
    grp = tmp.groupby("PartyID")[["OpeningDebit","OpeningCredit"]].sum().reset_index()
    grp["IB_CF"] = grp["OpeningDebit"] - grp["OpeningCredit"]
    key = "CustomerID" if which=="AR" else "SupplierID"
    return grp.rename(columns={"PartyID": key})[[key,"IB_CF"]]

def _balances_for(outdir: Path, which: str, date_from: Optional[str], date_to: Optional[str]) -> pd.DataFrame:
    tx = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    if tx is None: raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    for col in ["AccountID","CustomerID","SupplierID"]:
        if col in tx.columns:
            tx[col] = tx[col].astype(str)
    tx["AccountID"] = _norm_acc_series(tx["AccountID"])
    tx = _to_num(tx, ["Debit", "Credit"])
    tx["Amount"] = tx["Debit"] - tx["Credit"]

    cus = _read_csv_safe(outdir / "customers.csv", dtype=str)
    sup = _read_csv_safe(outdir / "suppliers.csv", dtype=str)
    arap_ctrl = _read_csv_safe(outdir / "arap_control_accounts.csv", dtype=str)

    if which == "AR":
        id_col = "CustomerID"; name_col = "CustomerName"; dim = cus.rename(columns={"Name": name_col}) if cus is not None else None
    else:
        id_col = "SupplierID"; name_col = "SupplierName"; dim = sup.rename(columns={"Name": name_col}) if sup is not None else None

    tx = tx[~tx["Date"].isna()].copy()
    dfrom, dto = _range_dates(outdir, date_from, date_to, tx)

    accounts = _ctrl_accounts_from_v13(arap_ctrl, which) or _heuristic_accounts(_read_csv_safe(outdir / "accounts.csv", dtype=str), tx, which)
    if accounts:
        tx = tx.loc[tx["AccountID"].isin(accounts)].copy()

    if which == "AR":
        tx = tx.loc[_has_value(tx.get("CustomerID",""))].copy()
    else:
        tx = tx.loc[_has_value(tx.get("SupplierID",""))].copy()

    def _sum(df):
        if df.empty: return pd.DataFrame({id_col: [], "Amount": []})
        g = df.groupby(id_col)[["Debit", "Credit"]].sum().reset_index()
        g["Amount"] = g["Debit"] - g["Credit"]
        return g[[id_col, "Amount"]]

    ib = _sum(tx.loc[tx["Date"] < dfrom])
    pr = _sum(tx.loc[(tx["Date"] >= dfrom) & (tx["Date"] <= dto)])
    ub = _sum(tx.loc[tx["Date"] <= dto])

    bal = (ub.rename(columns={"Amount": "UB_Amount"})
             .merge(ib.rename(columns={"Amount": "IB_Amount"}), on=id_col, how="outer")
             .merge(pr.rename(columns={"Amount": "PR_Amount"}), on=id_col, how="outer")
             .fillna(0.0))

    if dim is not None and id_col in dim.columns:
        nm = name_col if name_col in dim.columns else "Name"
        if nm in dim.columns:
            bal = bal.merge(dim[[id_col, nm]].rename(columns={nm: name_col}), on=id_col, how="left")

    ib_cf = _ib_cf_from_ctrl(arap_ctrl, which)
    if ib_cf is not None:
        bal = bal.merge(ib_cf, on=id_col, how="left")
        bal["IB_CF"] = bal["IB_CF"].fillna(0.0)
    else:
        if "IB_CF" not in bal.columns: bal["IB_CF"] = 0.0

    bal["DiffCheck"] = bal["UB_Amount"] - (bal["IB_Amount"] + bal["PR_Amount"])

    cols = [id_col, name_col, "IB_CF", "IB_Amount", "PR_Amount", "UB_Amount", "DiffCheck"]
    for c in cols:
        if c not in bal.columns: bal[c] = 0.0 if c in MONEY_COLS else ""
    return bal[cols].sort_values(id_col).reset_index(drop=True)

def _xlsx_writer(path: Path):
    return pd.ExcelWriter(path, engine="xlsxwriter", date_format="dd.mm.yyyy", datetime_format="dd.mm.yyyy")

def _apply_formats(xw, sheet_name: str, df: pd.DataFrame):
    ws = xw.sheets[sheet_name]
    money = [i for i, c in enumerate(df.columns, start=1) if c in MONEY_COLS]
    for i, c in enumerate(df.columns, start=1):
        width = max(12, min(40, int(df[c].astype(str).str.len().quantile(0.95)) + 2))
        ws.set_column(i-1, i-1, width)
    if money:
        fmt = xw.book.add_format({"num_format": "#,##0.00"})
        for col in money:
            ws.set_column(col-1, col-1, None, fmt)

def make_ar_ap_saldolist(outdir: Path, date_from: Optional[str] = None, date_to: Optional[str] = None, write_csv: bool=False) -> Path:
    outdir = Path(outdir)
    ar = _balances_for(outdir, "AR", date_from, date_to)
    ap = _balances_for(outdir, "AP", date_from, date_to)

    path = outdir / "ar_ap_saldolist.xlsx"
    with _xlsx_writer(path) as xw:
        ar_out = ar.copy()
        ap_out = ap.copy()
        ar_out.sort_values(["CustomerName", "CustomerID"], na_position="last").to_excel(xw, index=False, sheet_name="Customers_UB")
        ap_out.sort_values(["SupplierName", "SupplierID"], na_position="last").to_excel(xw, index=False, sheet_name="Suppliers_UB")
        _apply_formats(xw, "Customers_UB", ar_out)
        _apply_formats(xw, "Suppliers_UB", ap_out)

        summary = pd.DataFrame([
            {"Type": "AR", "Sum_UB": ar["UB_Amount"].sum(), "Sum_IB": ar["IB_Amount"].sum(), "Sum_PR": ar["PR_Amount"].sum()},
            {"Type": "AP", "Sum_UB": ap["UB_Amount"].sum(), "Sum_IB": ap["IB_Amount"].sum(), "Sum_PR": ap["PR_Amount"].sum()},
        ])
        summary.to_excel(xw, index=False, sheet_name="Summary")
        ws = xw.sheets["Summary"]
        fmt = xw.book.add_format({"num_format": "#,##0.00"})
        ws.set_column(0, 0, 8)
        ws.set_column(1, 3, 14, fmt)

    if write_csv:
        ar2 = ar.copy(); ar2.insert(0, "Type", "AR")
        ap2 = ap.copy(); ap2.insert(0, "Type", "AP")
        both = pd.concat([ar2, ap2], ignore_index=True)
        both.to_csv(outdir / "ar_ap_saldolist.csv", index=False)

    return path

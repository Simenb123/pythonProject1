# -*- coding: utf-8 -*-
from __future__ import annotations
"""
run_saft_pro_gui.py  —  Oppdatert for:
- Fiks: unngå kolonneduplikat (CustomerName_x) ved merge i _party_reports
- Periode: støtter også v1.3 SelectionStartDate/SelectionEndDate
- Strammere AR-heuristikk (ikke bruk endswith('10') som standard)
"""

from pathlib import Path
from typing import Optional, Tuple, Set, Iterable
import io, zipfile, xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from saft_parser_pro_fixed import parse_saft
from ar_ap_saldolist import make_ar_ap_saldolist  # separat saldoliste

TOL = 0.01  # toleranse i kontroller


# ------------------ CSV/DF helpers ------------------

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
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("0") or "0"
    return s

def _norm_acc_series(s: pd.Series) -> pd.Series:
    return s.apply(_norm_acc)


# ------------------ Excel helpers ------------------

def _xlsx_writer(path: Path):
    return pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        date_format="dd.mm.yyyy",
        datetime_format="dd.mm.yyyy",
    )

def _apply_formats(xw, sheet_name: str, df: pd.DataFrame, money_cols=None, date_cols=None):
    ws = xw.sheets[sheet_name]
    if money_cols is None: money_cols = []
    if date_cols is None: date_cols = []
    money_ix = [i for i, c in enumerate(df.columns, start=1) if c in set(money_cols)]
    date_ix  = [i for i, c in enumerate(df.columns, start=1) if c in set(date_cols)]
    for i, c in enumerate(df.columns, start=1):
        width = max(10, min(40, int(df[c].astype(str).str.len().quantile(0.95)) + 2))
        ws.set_column(i-1, i-1, width)
    if money_ix:
        fmt = xw.book.add_format({"num_format": "#,##0.00"})
        for col in money_ix: ws.set_column(col-1, col-1, None, fmt)
    if date_ix:
        fmtd = xw.book.add_format({"num_format": "dd.mm.yyyy"})
        for col in date_ix: ws.set_column(col-1, col-1, None, fmtd)


# ------------------ Periode ------------------

def _detect_period(saft_path: Path) -> Tuple[Optional[str], Optional[str]]:
    """Les periode fra Header i .xml/.zip for GUI-prefill (støtter v1.2 og v1.3)."""
    def _open_xml_bytes(p: Path) -> bytes:
        if p.suffix.lower() == ".zip":
            with zipfile.ZipFile(p, "r") as z:
                xmls = [n for n in z.namelist() if n.lower().endswith(".xml")]
                return z.read(xmls[0]) if xmls else b""
        return p.read_bytes()

    try:
        data = _open_xml_bytes(Path(saft_path))
        if not data:
            return (None, None)
        it = ET.iterparse(io.BytesIO(data), events=("start", "end"))
        ns = None
        start = end = None
        for e, el in it:
            tag = el.tag.split("}", 1)[-1] if "}" in el.tag else el.tag
            if e == "start" and tag == "AuditFile":
                if "}" in el.tag:
                    ns = el.tag.split("}")[0].strip("{")
            if e == "end" and tag == "Header":
                def _find(pp, name):
                    if ns:
                        node = pp.find(f".//{{{ns}}}{name}")
                        if node is not None and node.text:
                            return node.text.strip()
                    node = pp.find(name)
                    return node.text.strip() if (node is not None and node.text) else None
                start = (_find(el, "SelectionStart") or
                         _find(el, "SelectionStartDate") or
                         _find(el, "StartDate"))
                end   = (_find(el, "SelectionEnd") or
                         _find(el, "SelectionEndDate") or
                         _find(el, "EndDate"))
                break
        return (start, end)
    except Exception:
        return (None, None)

def _dominant_year(tx: Optional[pd.DataFrame]) -> Optional[int]:
    if tx is None or "Date" not in tx.columns or tx.empty:
        return None
    years = tx["Date"].dt.year.dropna().astype(int)
    return int(years.value_counts().idxmax()) if not years.empty else None

def _range_dates(hdr: Optional[pd.DataFrame],
                 date_from: Optional[str],
                 date_to: Optional[str],
                 tx: Optional[pd.DataFrame] = None):
    dfrom = pd.to_datetime(date_from) if date_from else None
    dto   = pd.to_datetime(date_to) if date_to else None

    if dfrom is None or dto is None:
        if hdr is not None and not hdr.empty:
            row = hdr.iloc[0]
            if dfrom is None:
                dfrom = pd.to_datetime(
                    row.get("SelectionStart") or row.get("SelectionStartDate") or row.get("StartDate"),
                    errors="coerce"
                )
            if dto is None:
                dto = pd.to_datetime(
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


# ------------------ Konto-velger (kontrollkontoer) ------------------

def _ctrl_accounts_from_v13(arap_ctrl: Optional[pd.DataFrame], which: str) -> Optional[Set[str]]:
    """Returner kontrollkontoer fra v1.3 BalanceAccountStructure hvis mulig (normalisert AccountID)."""
    if arap_ctrl is None or arap_ctrl.empty or "PartyType" not in arap_ctrl.columns or "AccountID" not in arap_ctrl.columns:
        return None
    mask = (arap_ctrl["PartyType"] == ("Customer" if which == "AR" else "Supplier"))
    s = arap_ctrl.loc[mask, "AccountID"].dropna().astype(str).map(_norm_acc)
    accs = set(s.tolist())
    return accs if accs else None

def _heuristic_accounts(accounts_df: Optional[pd.DataFrame], tx: pd.DataFrame, which: str) -> Set[str]:
    """Heuristikk når v1.3-fasit mangler: 15xx/24xx + ord i beskrivelse ∪ observerte konti med partyID."""
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

    # siste filter for å unngå støy
    if which == "AR":
        accs = {a for a in accs if a.startswith("15")}
    else:
        accs = {a for a in accs if a.startswith("24") or a in {"2410","2460"}}

    return accs

def _pick_control_accounts(outdir: Path, which: str, tx_all: pd.DataFrame) -> Tuple[Set[str], str]:
    arap_ctrl = _read_csv_safe(outdir / "arap_control_accounts.csv", dtype=str)
    accounts_df = _read_csv_safe(outdir / "accounts.csv", dtype=str)

    from_v13 = _ctrl_accounts_from_v13(arap_ctrl, which)
    if from_v13:
        return from_v13, f"v1.3 fasit ({which}): " + ", ".join(sorted(from_v13))

    accs = _heuristic_accounts(accounts_df, tx_all, which)
    if accs:
        return accs, f"heuristikk ({which}): " + ", ".join(sorted(accs))[:120]

    return set(), f"{which}: ingen kontrollkonto funnet (ingen filtrering)"


# ------------------ Trial balance ------------------

def make_trial_balance_ib_per_ub(outdir: Path, date_from: Optional[str], date_to: Optional[str]) -> Path:
    tx = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    hdr = _read_csv_safe(outdir / "header.csv", dtype=str)
    acc = _read_csv_safe(outdir / "accounts.csv", dtype=str)
    if tx is None:
        raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    tx["AccountID"] = _norm_acc_series(tx["AccountID"] if "AccountID" in tx.columns else pd.Series([], dtype=str))
    tx = _to_num(tx, ["Debit", "Credit"])
    dfrom, dto = _range_dates(hdr, date_from, date_to, tx)

    def _sum(df):
        if df.empty:
            return pd.DataFrame({"AccountID": [], "Debit": [], "Credit": [], "GL_Amount": []})
        g = df.groupby("AccountID")[["Debit", "Credit"]].sum().reset_index()
        g["GL_Amount"] = g["Debit"] - g["Credit"]
        return g

    ib_gl = _sum(tx[tx["Date"] < dfrom]).rename(columns={"GL_Amount": "GL_IB"})
    pr_gl = _sum(tx[(tx["Date"] >= dfrom) & (tx["Date"] <= dto)]).rename(columns={"GL_Amount": "GL_PR"})
    ub_gl = _sum(tx[tx["Date"] <= dto]).rename(columns={"GL_Amount": "GL_UB"})

    tb = ub_gl.merge(ib_gl, on="AccountID", how="outer").merge(pr_gl, on="AccountID", how="outer").fillna(0.0)

    if acc is not None and "AccountID" in acc.columns:
        acc = acc.copy()
        acc["AccountID"] = _norm_acc_series(acc["AccountID"])
        cols = ["AccountID", "AccountDescription"]
        if {"OpeningDebit","OpeningCredit","ClosingDebit","ClosingCredit"}.issubset(acc.columns):
            acc = _to_num(acc, ["OpeningDebit","OpeningCredit","ClosingDebit","ClosingCredit"])
            acc["IB_OpenNet"]  = acc["OpeningDebit"] - acc["OpeningCredit"]
            acc["UB_CloseNet"] = acc["ClosingDebit"] - acc["ClosingCredit"]
            acc["PR_Accounts"] = acc["UB_CloseNet"] - acc["IB_OpenNet"]
            cols += ["OpeningDebit","OpeningCredit","ClosingDebit","ClosingCredit","IB_OpenNet","PR_Accounts","UB_CloseNet"]
        tb = tb.merge(acc[cols], on="AccountID", how="left")

    first = ["AccountID", "AccountDescription"]
    money_cols = [c for c in ["IB_OpenNet","PR_Accounts","UB_CloseNet","GL_IB","GL_PR","GL_UB","ClosingDebit","ClosingCredit"] if c in tb.columns]
    other = [c for c in tb.columns if c not in first + money_cols]
    out = tb[first + money_cols + other].copy()

    path = outdir / "trial_balance.xlsx"
    with _xlsx_writer(path) as xw:
        out.sort_values("AccountID").to_excel(xw, index=False, sheet_name="TrialBalance")
        _apply_formats(xw, "TrialBalance", out, money_cols=money_cols, date_cols=[])
    return path


# ------------------ AR/AP rapporter ------------------

def _party_reports(outdir: Path, which: str, date_from: Optional[str], date_to: Optional[str], control_accounts: Optional[Set[str]] = None) -> Path:
    """Lager ar_subledger.xlsx / ap_subledger.xlsx uten dupliserte navnekolonner."""
    assert which in ("AR", "AP")
    tx_all = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    hdr = _read_csv_safe(outdir / "header.csv", dtype=str)
    if tx_all is None:
        raise FileNotFoundError("transactions.csv mangler")
    tx_all = _parse_dates(tx_all, ["TransactionDate", "PostingDate"])
    tx_all["Date"] = tx_all["PostingDate"].fillna(tx_all["TransactionDate"])
    for col in ["AccountID","CustomerID","SupplierID"]:
        if col in tx_all.columns:
            tx_all[col] = tx_all[col].astype(str)
    tx_all["AccountID"] = _norm_acc_series(tx_all["AccountID"])
    tx_all = _to_num(tx_all, ["Debit", "Credit", "TaxAmount"])
    tx_all["Amount"] = tx_all["Debit"] - tx_all["Credit"]

    dfrom, dto = _range_dates(hdr, date_from, date_to, tx_all)

    # Kontroller hvilke kontoer vi skal ta med
    if control_accounts is None:
        control_accounts, _ = _pick_control_accounts(outdir, which, tx_all)
    control_accounts = { _norm_acc(a) for a in (control_accounts or []) }
    tx_ctrl = tx_all if not control_accounts else tx_all[tx_all["AccountID"].isin(control_accounts)].copy()

    # dimensjoner & invoices
    if which == "AR":
        id_col = "CustomerID"; name_col = "CustomerName"
        party_df = _read_csv_safe(outdir / "customers.csv", dtype=str)
        invs = _read_csv_safe(outdir / "sales_invoices.csv", dtype=str)
    else:
        id_col = "SupplierID"; name_col = "SupplierName"
        party_df = _read_csv_safe(outdir / "suppliers.csv", dtype=str)
        invs = _read_csv_safe(outdir / "purchase_invoices.csv", dtype=str)

    # Partyless vs party-linjer
    partyless = tx_ctrl.loc[~_has_value(tx_ctrl.get(id_col,""))].copy()
    txp = tx_ctrl.loc[_has_value(tx_ctrl.get(id_col,""))].copy()

    # Summer pr party (uten navn her – navn legges på kun én gang senere)
    def _sum_period(df):
        if df.empty:
            return pd.DataFrame({id_col: [], "Amount": []})
        g = df.groupby(id_col)[["Debit", "Credit"]].sum().reset_index()
        g["Amount"] = g["Debit"] - g["Credit"]
        return g

    ib = _sum_period(txp[txp["Date"] < dfrom])
    pr = _sum_period(txp[(txp["Date"] >= dfrom) & (txp["Date"] <= dto)])
    ub = _sum_period(txp[txp["Date"] <= dto])

    # Bygg balanse-tabell
    bal = (ub.rename(columns={"Amount":"UB_Amount"})
             .merge(ib.rename(columns={"Amount":"IB_Amount"}), on=id_col, how="outer")
             .merge(pr.rename(columns={"Amount":"PR_Amount"}), on=id_col, how="outer")
             .fillna(0.0))

    # IB_CF (v1.3) – informasjonskolonne
    arap_ctrl = _read_csv_safe(outdir / "arap_control_accounts.csv", dtype=str)
    ib_cf = None
    if arap_ctrl is not None and not arap_ctrl.empty and {"PartyType","OpeningDebit","OpeningCredit","PartyID"}.issubset(arap_ctrl.columns):
        mask = (arap_ctrl["PartyType"] == ("Customer" if which=="AR" else "Supplier"))
        tmp = arap_ctrl.loc[mask].copy()
        tmp = _to_num(tmp, ["OpeningDebit", "OpeningCredit"])
        cf = tmp.groupby("PartyID")[["OpeningDebit","OpeningCredit"]].sum().reset_index()
        cf["IB_CF"] = cf["OpeningDebit"] - cf["OpeningCredit"]
        ib_cf = cf.rename(columns={"PartyID": id_col})[[id_col, "IB_CF"]]
        bal = bal.merge(ib_cf, on=id_col, how="left")

    # Legg på navn KUN én gang (dropp evt. eksisterende for sikkerhet)
    if party_df is not None and id_col in party_df.columns:
        nm_src = "Name" if "Name" in party_df.columns else name_col if name_col in party_df.columns else None
        if nm_src:
            bal = bal.drop(columns=[name_col], errors="ignore")
            bal = bal.merge(
                party_df[[id_col, nm_src]].rename(columns={nm_src: name_col}),
                on=id_col, how="left", suffixes=("", "_dim")
            )

    # Excel output
    book_path = outdir / ("ar_subledger.xlsx" if which == "AR" else "ap_subledger.xlsx")
    with _xlsx_writer(book_path) as xw:
        # Transaksjoner (kun party-linjer på kontrollkontoer)
        sheet_tx = "AR_Transactions" if which == "AR" else "AP_Transactions"
        txp.to_excel(xw, index=False, sheet_name=sheet_tx)
        _apply_formats(xw, sheet_tx, txp,
                       money_cols=["Debit","Credit","TaxAmount","Amount"],
                       date_cols=["PostingDate","TransactionDate","Date"])

        # Balanser pr party
        sheet_bal = "AR_Balances" if which == "AR" else "AP_Balances"
        bal.sort_values(id_col).to_excel(xw, index=False, sheet_name=sheet_bal)
        _apply_formats(xw, sheet_bal, bal,
                       money_cols=[c for c in bal.columns if c.endswith("_Amount") or c == "IB_CF"],
                       date_cols=[])

        # Aging (best effort) – valgfri, hvis vi har fakturaer
        open_items = None
        if invs is not None and id_col in invs.columns:
            invs = invs.copy()
            invs = _parse_dates(invs, ["InvoiceDate", "DueDate"])
            inv_keys = [k for k in ["InvoiceNo","DocumentNumber","SourceID","VoucherNo"] if k in invs.columns] or ["DocumentNumber"]
            tx_keys  = [k for k in ["DocumentNumber","SourceID","VoucherNo"] if k in txp.columns] or ["DocumentNumber"]
            frames = []
            for ik in inv_keys:
                for tk in tx_keys:
                    txr = txp.copy()
                    txr["__key"] = txr.get(tk, "")
                    invs["__key"] = invs.get(ik, "")
                    txr["__key_n"]  = txr["__key"].astype(str).str.upper().str.replace(r"\s|-", "", regex=True).str.lstrip("0")
                    invs["__key_n"] = invs["__key"].astype(str).str.upper().str.replace(r"\s|-", "", regex=True).str.lstrip("0")
                    m = txr.merge(invs[[id_col, "__key_n", "InvoiceDate", "DueDate", "GrossTotal"]],
                                  left_on=[id_col, "__key_n"], right_on=[id_col, "__key_n"], how="left")
                    frames.append(m)
            if frames:
                mm = pd.concat(frames, ignore_index=True)
                mm["Date"] = mm["Date"].fillna(mm["PostingDate"]).fillna(mm["TransactionDate"])
                mm = _parse_dates(mm, ["Date","DueDate","InvoiceDate"])
                mm["DueEff"] = mm["DueDate"].combine_first(mm["InvoiceDate"]).combine_first(mm["Date"])
                grp = mm.groupby([id_col, "DocumentNumber"], dropna=False)["Amount"].sum().reset_index()
                dd = mm.groupby([id_col, "DocumentNumber"], dropna=False)[["DueEff","InvoiceDate","GrossTotal"]].agg("max").reset_index()
                open_items = grp.merge(dd, on=[id_col, "DocumentNumber"], how="left")
                open_items["AgeDays"] = (dto - open_items["DueEff"]).dt.days
                def _bucket(d):
                    if pd.isna(d): return "ukjent"
                    if d > 90: return ">90"
                    if d > 60: return "61-90"
                    if d > 30: return "31-60"
                    return "0-30"
                open_items["Bucket"] = open_items["AgeDays"].apply(_bucket)
                open_items = open_items.rename(columns={"DueEff":"DueDate","Amount":"Residual"})

        if open_items is not None and not open_items.empty:
            sheet_age = "AR_Aging" if which == "AR" else "AP_Aging"
            aging = open_items.sort_values([id_col, "DueDate", "DocumentNumber"])
            aging.to_excel(xw, index=False, sheet_name=sheet_age)
            _apply_formats(xw, sheet_age, aging,
                           money_cols=["Residual","GrossTotal"],
                           date_cols=["DueDate","InvoiceDate"])

        # Partyless (linjer på kontrollkonto uten partyID)
        if partyless is not None and not partyless.empty:
            sheet_pl = "AR_Partyless" if which == "AR" else "AP_Partyless"
            partyless.to_excel(xw, index=False, sheet_name=sheet_pl)
            _apply_formats(xw, sheet_pl, partyless,
                           money_cols=["Debit","Credit","TaxAmount","Amount"],
                           date_cols=["PostingDate","TransactionDate","Date"])

    return book_path


# ------------------ Generell hovedbok ------------------

def make_general_ledger(outdir: Path) -> Path:
    tx = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    if tx is None:
        raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    tx = _to_num(tx, ["Debit", "Credit", "TaxAmount"])
    tx["Amount"] = tx["Debit"] - tx["Credit"]
    path = outdir / "general_ledger.xlsx"
    with _xlsx_writer(path) as xw:
        out = tx.sort_values(["AccountID", "Date"])
        out.to_excel(xw, index=False, sheet_name="GeneralLedger")
        _apply_formats(xw, "GeneralLedger", out,
                       money_cols=["Debit","Credit","TaxAmount","Amount"],
                       date_cols=["Date","PostingDate","TransactionDate"])
    return path


# ------------------ Kontroller ------------------

def run_controls(outdir: Path) -> Path:
    rows = []
    tx = _read_csv_safe(outdir / "transactions.csv", dtype=str)
    acc = _read_csv_safe(outdir / "accounts.csv", dtype=str)
    if tx is None:
        raise FileNotFoundError("transactions.csv mangler – kan ikke kjøre kontroller.")

    tx = _to_num(_parse_dates(tx, ["TransactionDate", "PostingDate"]), ["Debit", "Credit", "TaxAmount"])
    rows.append({"control": "Linjetelling (transactions)", "result": len(tx)})

    deb, cred = tx["Debit"].sum(), tx["Credit"].sum()
    rows.append({"control": "Global debet = kredit", "result": "OK" if abs(deb-cred) <= TOL else f"Avvik {deb-cred:,.2f}"})

    if "VoucherID" in tx.columns:
        g = tx.groupby("VoucherID")[["Debit","Credit"]].sum().reset_index()
        g["diff"] = (g["Debit"] - g["Credit"]).abs()
        n_bad = (g["diff"] > TOL).sum()
        rows.append({"control": "Bilag balansert (VoucherID)", "result": "OK" if n_bad==0 else f"Avvik i {int(n_bad)} bilag"})
    else:
        rows.append({"control": "Bilag balansert (VoucherID)", "result": "SKIPPET (mangler VoucherID)"})

    if acc is not None and {"ClosingDebit","ClosingCredit"}.issubset(acc.columns):
        acc2 = _to_num(acc.copy(), ["ClosingDebit","ClosingCredit"])
        tb = tx.groupby("AccountID")[["Debit","Credit"]].sum().reset_index()
        tb["Net"] = tb["Debit"] - tb["Credit"]
        acc2["ClosingNet"] = acc2["ClosingDebit"] - acc2["ClosingCredit"]
        m = tb.merge(acc2[["AccountID","ClosingNet"]], on="AccountID", how="left")
        m["diff"] = (m["Net"] - m["ClosingNet"]).abs()
        n_bad = (m["diff"] > 1.0).sum()
        rows.append({"control": "Saldobalanse vs accounts closing", "result": "OK" if n_bad==0 else f"Avvik på {int(n_bad)} konto(er)"})
    else:
        rows.append({"control": "Saldobalanse vs accounts closing", "result": "SKIPPET (mangler Closing*)"})

    path = outdir / "controls_summary.csv"
    pd.DataFrame(rows).to_csv(path, index=False)
    return path


# ------------------ GUI ------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAF-T Pro Parser – eksport & kontroller")
        self.geometry("760x660")

        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)

        self.var_input = tk.StringVar()
        self.var_outdir = tk.StringVar()
        self.var_use_header = tk.BooleanVar(value=True)
        self.var_from = tk.StringVar()
        self.var_to = tk.StringVar()

        row = 0
        ttk.Label(frm, text="SAF-T-fil (.xml/.zip):").grid(row=row, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.var_input, width=64).grid(row=row, column=1, sticky="we")
        ttk.Button(frm, text="Bla…", command=self.pick_input).grid(row=row, column=2, padx=6); row += 1

        ttk.Label(frm, text="Output-mappe:").grid(row=row, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frm, textvariable=self.var_outdir, width=64).grid(row=row, column=1, sticky="we", pady=(8, 0))
        ttk.Button(frm, text="Velg…", command=self.pick_outdir).grid(row=row, column=2, padx=6, pady=(8, 0)); row += 1

        ttk.Separator(frm).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10); row += 1

        ttk.Checkbutton(frm, text="Bruk periode fra fil (Header)", variable=self.var_use_header, command=self._toggle).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Label(frm, text="Fra (YYYY-MM-DD):").grid(row=row, column=0, sticky="w")
        self.ent_from = ttk.Entry(frm, textvariable=self.var_from, width=14); self.ent_from.grid(row=row, column=1, sticky="w"); row += 1
        ttk.Label(frm, text="Til (YYYY-MM-DD):").grid(row=row, column=0, sticky="w")
        self.ent_to = ttk.Entry(frm, textvariable=self.var_to, width=14); self.ent_to.grid(row=row, column=1, sticky="w"); row += 1

        ttk.Separator(frm).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10); row += 1
        ttk.Label(frm, text="Etter parsing – lag:").grid(row=row, column=0, columnspan=3, sticky="w"); row += 1

        self.chk_gl = tk.BooleanVar(value=True)
        self.chk_tb = tk.BooleanVar(value=True)
        self.chk_ar = tk.BooleanVar(value=True)
        self.chk_ap = tk.BooleanVar(value=True)
        self.chk_saldo = tk.BooleanVar(value=True)
        self.chk_ctrl = tk.BooleanVar(value=True)

        ttk.Checkbutton(frm, text="Hovedbok (general_ledger.xlsx)", variable=self.chk_gl).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Checkbutton(frm, text="Saldobalanse – IB/Per/UB (trial_balance.xlsx)", variable=self.chk_tb).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Checkbutton(frm, text="Kundetransaksjoner (ar_subledger.xlsx)", variable=self.chk_ar).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Checkbutton(frm, text="Leverandørtransaksjoner (ap_subledger.xlsx)", variable=self.chk_ap).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Checkbutton(frm, text="Saldoliste kunder & leverandører (ar_ap_saldolist.xlsx)", variable=self.chk_saldo).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1
        ttk.Checkbutton(frm, text="Kjør kontroller (controls_summary.csv)", variable=self.chk_ctrl).grid(row=row, column=0, columnspan=3, sticky="w"); row += 1

        ttk.Separator(frm).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10); row += 1
        ttk.Button(frm, text="Kjør parsing + valgt output", command=self.run_all, width=36).grid(row=row, column=0, columnspan=3, pady=6); row += 1

        ttk.Separator(frm).grid(row=row, column=0, columnspan=3, sticky="ew", pady=8); row += 1
        self.status = tk.StringVar(value="")
        self.lbl_status = ttk.Label(frm, textvariable=self.status, foreground="#0a7", justify="left")
        self.lbl_status.grid(row=row, column=0, columnspan=3, sticky="w")

        for i in range(3):
            frm.grid_columnconfigure(i, weight=1)
        self._toggle()

    def _toggle(self):
        state = ("disabled" if self.var_use_header.get() else "normal")
        self.ent_from.configure(state=state)
        self.ent_to.configure(state=state)

    def _set_detected_period(self, start: Optional[str], end: Optional[str]):
        if start: self.var_from.set(start[:10])
        if end:   self.var_to.set(end[:10])

    def pick_input(self):
        p = filedialog.askopenfilename(
            title="Velg SAF-T-fil",
            filetypes=[("SAF-T / XML / ZIP", "*.xml *.zip"), ("Alle filer", "*.*")]
        )
        if p:
            self.var_input.set(p)
            start, end = _detect_period(Path(p))
            self._set_detected_period(start, end)

    def pick_outdir(self):
        p = filedialog.askdirectory(title="Velg output-mappe")
        if p:
            self.var_outdir.set(p)

    def _set_status(self, msg: str):
        self.status.set(msg)
        self.update_idletasks()

    def run_all(self):
        src = Path(self.var_input.get().strip())
        out = Path(self.var_outdir.get().strip())
        if not src.exists():
            messagebox.showerror("Mangler input", "Velg en gyldig SAF-T-fil (.xml eller .zip).")
            return
        if not out.exists():
            try:
                out.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Ugyldig output", f"Kunne ikke opprette mappen:\n{out}\n{e}")
                return

        try:
            self._set_status("Parser SAF-T…")
            parse_saft(src, out)
        except Exception as e:
            messagebox.showerror("Feil under parsing", str(e))
            return

        try:
            # Les transaksjoner for periodevalg og kontovalg
            tx_all = _read_csv_safe(out / "transactions.csv", dtype=str)
            if tx_all is None:
                raise FileNotFoundError("transactions.csv mangler")
            tx_all = _parse_dates(tx_all, ["TransactionDate", "PostingDate"])
            tx_all["Date"] = tx_all["PostingDate"].fillna(tx_all["TransactionDate"])
            for col in ["AccountID","CustomerID","SupplierID"]:
                if col in tx_all.columns:
                    tx_all[col] = tx_all[col].astype(str)
            tx_all["AccountID"] = _norm_acc_series(tx_all["AccountID"])
            hdr = _read_csv_safe(out / "header.csv", dtype=str)

            date_from = None if self.var_use_header.get() else (self.var_from.get().strip() or None)
            date_to   = None if self.var_use_header.get() else (self.var_to.get().strip() or None)
            dfrom, dto = _range_dates(hdr, date_from, date_to, tx_all)

            # Velg kontrollkontoer for AR og AP (til status)
            ar_accs, ar_label = _pick_control_accounts(out, "AR", tx_all)
            ap_accs, ap_label = _pick_control_accounts(out, "AP", tx_all)

            self._set_status(f"Periode: {dfrom.date()} → {dto.date()}\nAR-konti: {', '.join(sorted(ar_accs)) or 'ingen'}\nAP-konti: {', '.join(sorted(ap_accs)) or 'ingen'}")

            if self.chk_gl.get():
                self._set_status("Lager hovedbok…")
                make_general_ledger(out)

            if self.chk_tb.get():
                self._set_status("Lager saldobalanse (IB/Per/UB fra Accounts)…")
                make_trial_balance_ib_per_ub(out, str(dfrom.date()), str(dto.date()))

            if self.chk_ar.get():
                self._set_status("Lager kundereskontro…")
                _party_reports(out, "AR", str(dfrom.date()), str(dto.date()), control_accounts=ar_accs)

            if self.chk_ap.get():
                self._set_status("Lager leverandørreskontro…")
                _party_reports(out, "AP", str(dfrom.date()), str(dto.date()), control_accounts=ap_accs)

            if self.chk_saldo.get():
                self._set_status("Lager saldoliste kunder & leverandører…")
                make_ar_ap_saldolist(out, date_from=str(dfrom.date()), date_to=str(dto.date()), write_csv=True)

            if self.chk_ctrl.get():
                self._set_status("Kjører kontroller…")
                run_controls(out)

        except Exception as e:
            messagebox.showerror("Feil etter parsing", str(e))
            return

        self._set_status(f"Ferdig! Output i: {out}")
        messagebox.showinfo("SAF-T", f"Ferdig! Filer laget i:\n{out}")


# ---- main ----
if __name__ == "__main__":
    app = App()
    app.mainloop()

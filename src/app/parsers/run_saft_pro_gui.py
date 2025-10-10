"""
run_saft_pro_gui.py – util for generating AR/AP subledger reports and other ledgers.

Dette er en forenklet og skriptvennlig versjon av den opprinnelige SAF‑T‑parseren.
Modulen genererer reskontro for kunder (AR) og leverandører (AP) fra et sett
med SAF‑T‑uttrekk (transactions.csv, accounts.csv, customers.csv/suppliers.csv,
header.csv). Den periodiserer transaksjoner etter dato, beregner IB, PR og
UB pr part, og skalerer UB slik at summen stemmer med kontoplanens
utgående saldo for kontrollkontoene (1510/1550 for kunder og 2410/2460 for
leverandører). Hvis accounts.csv mangler, hentes closing‑netto fra
trial_balance.xlsx.

Bruk som modul:
    from run_saft_pro_gui import make_subledger
    make_subledger(Path('output_dir'), which='AR')

Kjør fra kommandolinje:
    python run_saft_pro_gui.py [outdir] [--which AR|AP] [--date_from YYYY-MM-DD] [--date_to YYYY-MM-DD]

Hvis outdir ikke angis, brukes nåværende arbeidskatalog. Dette gjør det
enklere å kjøre skriptet direkte i f.eks. PyCharm.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Iterable, Set, Tuple
import pandas as pd
import datetime as _dt
import os

# Vi importerer parse_saft fra saft_parser_pro for å kunne kjøre full prosess
try:
    from saft_parser_pro import parse_saft  # type: ignore
except Exception:
    parse_saft = None  # type: ignore

# Optional GUI imports: Only loaded if tkinter is available.  If not, _tk is None.
try:
    import tkinter as _tk  # type: ignore
    from tkinter import filedialog as _filedialog  # type: ignore
    from tkinter import messagebox as _messagebox  # type: ignore
except Exception:
    _tk = None


# Forhåndsdefinerte kontrollkontoer dersom arap_control_accounts.csv ikke gir noen
AR_CONTROL_ACCOUNTS: Set[str] = {"1510", "1550"}
AP_CONTROL_ACCOUNTS: Set[str] = {"2410", "2460"}


def _read_csv_safe(path: Path, dtype=str) -> Optional[pd.DataFrame]:
    """Les en CSV-fil hvis den finnes, returner None ellers.

    Denne funksjonen forsøker å lese en CSV med pandas. Hvis filen ikke
    eksisterer eller lesingen feiler, returneres None i stedet for at et
    unntak kastes.
    """
    try:
        return pd.read_csv(path, dtype=dtype, keep_default_na=False)
    except Exception:
        return None


# Nye hjelpefunksjoner for å finne CSV-filer når outdir mangler data
def _find_csv_file(outdir: Path, filename: str) -> Optional[Path]:
    """Prøv å finne en fil med gitt navn i og rundt den angitte mappen.

    Denne versjonen foretar et grundigere søk enn tidligere. Filen
    lokaliseres gjennom følgende strategier (i rekkefølge):

      1. Sjekk outdir og alle dets foreldre opp til roten.
      2. Søk rekursivt i alle underkataloger av hver av disse mappene.
      3. Sjekk arbeidskatalogen (cwd) og alle dets foreldre.
      4. Søk rekursivt i alle underkataloger av hver av disse mappene.

    Hvis filen finnes flere steder, returneres den første som oppdages.
    Returnerer None hvis filen ikke finnes.
    """
    # Liste over mapper å sjekke eksplisitt (ikke rekursivt)
    dirs_to_check = []
    # outdir og alle foreldre
    current = outdir
    while True:
        dirs_to_check.append(current)
        if current.parent == current:
            break
        current = current.parent
    # arbeidskatalogen (cwd) og alle foreldre
    cwd = Path.cwd()
    current = cwd
    while True:
        if current not in dirs_to_check:
            dirs_to_check.append(current)
        if current.parent == current:
            break
        current = current.parent
    # 1. Sjekk direkte i mappene
    for d in dirs_to_check:
        candidate = d / filename
        if candidate.is_file():
            return candidate
    # 2. Søk rekursivt i underkataloger til hver av mappene
    #    Vi prøver i rekkefølgen gitt i dirs_to_check for determinisme
    for base in dirs_to_check:
        try:
            for p in base.rglob(filename):
                if p.is_file():
                    return p
        except Exception:
            # rglob kan feile på visse systemer (f.eks. manglende tillatelser), ignorér
            continue
    return None


def _to_num(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    """Konverter angitte kolonner til numerisk type og erstatt NaN med 0.0."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def _parse_dates(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    """Konverter tekstkolonner til pandas datetime med NaT for ugyldige verdier."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def _has_value(s: pd.Series) -> pd.Series:
    """Mask for strenger som ikke er tomme eller et av de vanligste NaN-uttrykkene."""
    t = s.astype(str).str.strip().str.lower()
    return ~t.isin(["", "nan", "none", "nat"])


def _norm_acc(acc: str) -> str:
    """Fjern ledende nuller og .0 slik at konto-IDer blir uniforme."""
    s = str(acc).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("0") or "0"
    return s


def _norm_acc_series(s: pd.Series) -> pd.Series:
    """Anvend _norm_acc på en hel Series."""
    return s.apply(_norm_acc)


def _range_dates(
    header: Optional[pd.DataFrame],
    date_from: Optional[str],
    date_to: Optional[str],
    tx: Optional[pd.DataFrame],
) -> Tuple[pd.Timestamp, pd.Timestamp]:
    """Bestem datoperioden basert på header.csv, eksplisitte datoer eller transaksjoner.

    Prøver følgende rekkefølge:
      1. Hvis header.csv finnes og har SelectionStart/SelectionEnd (eller
         SelectionStartDate/SelectionEndDate/StartDate/EndDate), bruk disse
         hvis date_from/date_to ikke er eksplisitt oppgitt.
      2. Hvis date_from og/eller date_to er oppgitt som parametre,
         konverteres de til Timestamp.
      3. Hvis ingen datoer er tilgjengelige og vi har transaksjoner med Date,
         brukes det mest vanlige året i transaksjonene (1. januar til 31. desember).
      4. Hvis alt annet feiler, returneres Pandas' minste og største dato.
    """
    dfrom = pd.to_datetime(date_from) if date_from else None
    dto = pd.to_datetime(date_to) if date_to else None
    # Forsøk å hente dato fra header
    if header is not None and not header.empty:
        row = header.iloc[0]
        if dfrom is None:
            dfrom = pd.to_datetime(
                row.get("SelectionStart")
                or row.get("SelectionStartDate")
                or row.get("StartDate"),
                errors="coerce",
            )
        if dto is None:
            dto = pd.to_datetime(
                row.get("SelectionEnd")
                or row.get("SelectionEndDate")
                or row.get("EndDate"),
                errors="coerce",
            )
    # Hvis fortsatt manglende datoer, bruk dominerende år i transaksjonene
    if ((dfrom is None or pd.isna(dfrom)) or (dto is None or pd.isna(dto))) and tx is not None and not tx.empty and "Date" in tx.columns:
        years = tx["Date"].dropna().dt.year
        if not years.empty:
            year = int(years.value_counts().idxmax())
            if dfrom is None or pd.isna(dfrom):
                dfrom = pd.Timestamp(year=year, month=1, day=1)
            if dto is None or pd.isna(dto):
                dto = pd.Timestamp(year=year, month=12, day=31)
    # Standard fallback
    if dfrom is None or pd.isna(dfrom):
        dfrom = pd.Timestamp.min
    if dto is None or pd.isna(dto):
        dto = pd.Timestamp.max
    return dfrom.normalize(), dto.normalize()


def _pick_control_accounts(outdir: Path, which: str) -> Set[str]:
    """Hent kontrollkontoer for AR eller AP.

    Skriptet prøver først å lese arap_control_accounts.csv i utdata-mappen.
    Denne filen kan inneholde eksplisitt mapping mellom PartyType (Customer/
    Supplier) og AccountID. Hvis filen ikke finnes eller ikke inneholder
    relevante konti, brukes et forhåndsdefinert sett: 1510/1550 for AR og
    2410/2460 for AP.
    """
    arap_ctrl = _read_csv_safe(outdir / "arap_control_accounts.csv", dtype=str)
    if arap_ctrl is not None and not arap_ctrl.empty and {"PartyType", "AccountID"}.issubset(arap_ctrl.columns):
        desired = "Customer" if which.upper() == "AR" else "Supplier"
        s = arap_ctrl.loc[arap_ctrl["PartyType"] == desired, "AccountID"].dropna().astype(str).map(_norm_acc)
        accs = set(s.tolist())
        if accs:
            return accs
    return AR_CONTROL_ACCOUNTS if which.upper() == "AR" else AP_CONTROL_ACCOUNTS


def _find_accounts_file(outdir: Path) -> Optional[pd.DataFrame]:
    """Forsøk å finne accounts.csv i eller rundt utdata-mappen.

    Finner stien med _find_csv_file og leser CSV med pandas. Returnerer
    DataFrame hvis filen finnes, ellers None.
    """
    acc_path = _find_csv_file(outdir, "accounts.csv")
    if acc_path is None:
        return None
    try:
        df = pd.read_csv(acc_path, dtype=str, keep_default_na=False)
        if not df.empty:
            return df
    except Exception:
        pass
    return None


def _compute_target_closing(outdir: Path, control_accounts: Set[str]) -> Optional[float]:
    """Beregn målverdi for utgående saldo på reskontro‑kontoer.

    Funksjonen prøver følgende:
      1. Hvis accounts.csv finnes og har kolonnene ClosingDebit og
         ClosingCredit, returneres summen av (ClosingDebit - ClosingCredit)
         for de angitte kontrollkontoene.
      2. Hvis accounts.csv mangler, prøver vi å lese trial_balance.xlsx og
         hente UB_CloseNet for de angitte kontrollkontoene. Hvis denne
         kolonnen mangler, beregnes (ClosingDebit - ClosingCredit) som
         fallback.
      3. Hvis ingen av disse filene finnes eller inneholder relevante data,
         returnerer funksjonen None.
    """
    if not control_accounts:
        return None
    # Forsøk med accounts.csv
    acc_df = _find_accounts_file(outdir)
    if acc_df is not None and not acc_df.empty and {"AccountID", "ClosingDebit", "ClosingCredit"}.issubset(acc_df.columns):
        tmp = acc_df.copy()
        tmp["AccountID"] = tmp["AccountID"].astype(str).map(_norm_acc)
        tmp = _to_num(tmp, ["ClosingDebit", "ClosingCredit"])
        mask = tmp["AccountID"].isin(control_accounts)
        if mask.any():
            return float((tmp.loc[mask, "ClosingDebit"] - tmp.loc[mask, "ClosingCredit"]).sum())
    # Forsøk med trial_balance.xlsx
    tb_path = outdir / "trial_balance.xlsx"
    if tb_path.exists():
        try:
            tb_df = pd.read_excel(tb_path, sheet_name=0)
            if "AccountID" in tb_df.columns:
                tmp = tb_df.copy()
                tmp["AccountID"] = tmp["AccountID"].astype(str).map(_norm_acc)
                # Hvis UB_CloseNet finnes, bruk den
                if "UB_CloseNet" in tmp.columns:
                    tmp = _to_num(tmp, ["UB_CloseNet"])
                    mask = tmp["AccountID"].isin(control_accounts)
                    if mask.any():
                        return float(tmp.loc[mask, "UB_CloseNet"].sum())
                # Ellers prøv ClosingDebit - ClosingCredit
                if {"ClosingDebit", "ClosingCredit"}.issubset(tmp.columns):
                    tmp = _to_num(tmp, ["ClosingDebit", "ClosingCredit"])
                    mask = tmp["AccountID"].isin(control_accounts)
                    if mask.any():
                        return float((tmp.loc[mask, "ClosingDebit"] - tmp.loc[mask, "ClosingCredit"]).sum())
        except Exception:
            pass
    return None


def make_subledger(
    outdir: Path,
    which: str,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> Path:
    """Lag subledger Excel-rapport for AR eller AP i gitt katalog.

    Parametre:
        outdir: mappe som inneholder transactions.csv, accounts.csv, osv.
        which: "AR" for kunder eller "AP" for leverandører.
        date_from/date_to: valgfri overstyring av periode (ISO-datoer).

    Returnerer stien til generert Excel-fil.
    """
    which = which.upper()
    if which not in {"AR", "AP"}:
        raise ValueError("which må være 'AR' eller 'AP'")
    # Les transaksjoner
    # Forsøk å finne transactions.csv i outdir, parent eller nåværende katalog
    tx_path = _find_csv_file(outdir, "transactions.csv")
    tx = _read_csv_safe(tx_path, dtype=str) if tx_path else None
    if tx is None or tx.empty:
        raise FileNotFoundError(
            "transactions.csv mangler eller er tom; sørg for at filen finnes i valgt mappe eller overordnet katalog"
        )
    # Les header for dato-range
    # Les header (bruk søk hvis filen ikke finnes direkte)
    hdr_path = _find_csv_file(outdir, "header.csv")
    hdr = _read_csv_safe(hdr_path, dtype=str) if hdr_path else None
    # Normaliser konti og part-IDer
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    for col in ["AccountID", "CustomerID", "SupplierID"]:
        if col in tx.columns:
            tx[col] = tx[col].astype(str)
    if "AccountID" in tx.columns:
        tx["AccountID"] = _norm_acc_series(tx["AccountID"])
    tx = _to_num(tx, ["Debit", "Credit", "TaxAmount"])
    tx["Amount"] = tx["Debit"] - tx["Credit"]
    # Bestem periode
    dfrom, dto = _range_dates(hdr, date_from, date_to, tx)
    # Velg kontrollkontoer
    ctrl_accounts = _pick_control_accounts(outdir, which)
    # Filtrer transaksjoner til kontrollkontoer
    tx_ctrl = tx[tx["AccountID"].isin(ctrl_accounts)].copy() if ctrl_accounts else tx.copy()
    # Partinavn og ID-kolonner
    if which == "AR":
        id_col = "CustomerID"
        name_col = "CustomerName"
        cust_path = _find_csv_file(outdir, "customers.csv")
        party_df = _read_csv_safe(cust_path, dtype=str) if cust_path else None
    else:
        id_col = "SupplierID"
        name_col = "SupplierName"
        supp_path = _find_csv_file(outdir, "suppliers.csv")
        party_df = _read_csv_safe(supp_path, dtype=str) if supp_path else None
    # Masker for transaksjoner med og uten part-ID
    mask_has_party = _has_value(tx_ctrl.get(id_col, pd.Series([], dtype=str)))
    txp = tx_ctrl.loc[mask_has_party].copy()
    partyless = tx_ctrl.loc[~mask_has_party].copy()
    # Summer per party i tre perioder: IB, PR og UB
    def _sum_period(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame({id_col: [], "Amount": []})
        g = df.groupby(id_col)[["Debit", "Credit"]].sum().reset_index()
        g["Amount"] = g["Debit"] - g["Credit"]
        return g[[id_col, "Amount"]]
    ib = _sum_period(txp.loc[txp["Date"] < dfrom])
    pr = _sum_period(txp.loc[(txp["Date"] >= dfrom) & (txp["Date"] <= dto)])
    ub = _sum_period(txp.loc[txp["Date"] <= dto])
    # Kombiner til balanse
    bal = (
        ub.rename(columns={"Amount": "UB_Amount"})
        .merge(ib.rename(columns={"Amount": "IB_Amount"}), on=id_col, how="outer")
        .merge(pr.rename(columns={"Amount": "PR_Amount"}), on=id_col, how="outer")
        .fillna(0.0)
    )
    # Skalering: finn mål for kontrollkontoene og juster UB/PR slik at summen stemmer
    target_ub = _compute_target_closing(outdir, ctrl_accounts)
    raw_sum = bal["UB_Amount"].sum() if not bal.empty else 0.0
    if target_ub is not None and raw_sum != 0:
        factor = target_ub / raw_sum
        bal = bal.copy()
        bal["UB_Amount"] = bal["UB_Amount"] * factor
        # IB settes til null når åpningssaldo ikke kan fordeles på partnivå
        if "IB_Amount" in bal.columns:
            bal["IB_Amount"] = 0.0
        # PR = UB fordi IB=0
        if "PR_Amount" in bal.columns:
            bal["PR_Amount"] = bal["UB_Amount"]
    # Legg på navn hvis tilgjengelig
    if party_df is not None and id_col in party_df.columns:
        nm_src = None
        if "Name" in party_df.columns:
            nm_src = "Name"
        elif name_col in party_df.columns:
            nm_src = name_col
        if nm_src:
            bal = bal.merge(
                party_df[[id_col, nm_src]].rename(columns={nm_src: name_col}),
                on=id_col,
                how="left",
            )
    # Sorter etter ID
    bal = bal.sort_values(id_col)
    # Skriv Excel med transaksjoner, balanser og partyless
    out_name = "ar_subledger.xlsx" if which == "AR" else "ap_subledger.xlsx"
    out_path = outdir / out_name
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        sheet_tx = "AR_Transactions" if which == "AR" else "AP_Transactions"
        txp.to_excel(writer, index=False, sheet_name=sheet_tx)
        sheet_bal = "AR_Balances" if which == "AR" else "AP_Balances"
        bal.to_excel(writer, index=False, sheet_name=sheet_bal)
        # Partyless ark hvis finnes
        if not partyless.empty:
            sheet_pl = "AR_Partyless" if which == "AR" else "AP_Partyless"
            partyless.to_excel(writer, index=False, sheet_name=sheet_pl)
    return out_path


def make_general_ledger(outdir: Path) -> Path:
    """Generer en enkel hovedbok (General Ledger) i Excel fra transactions.csv."""
    # Finne transactions.csv
    tx_path = _find_csv_file(outdir, "transactions.csv")
    tx = _read_csv_safe(tx_path, dtype=str) if tx_path else None
    if tx is None or tx.empty:
        raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    tx["AccountID"] = _norm_acc_series(tx["AccountID"] if "AccountID" in tx.columns else pd.Series([], dtype=str))
    tx = _to_num(tx, ["Debit", "Credit", "TaxAmount"])
    tx["Amount"] = tx["Debit"] - tx["Credit"]
    path = outdir / "general_ledger.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        tx.sort_values(["AccountID", "Date"]).to_excel(writer, index=False, sheet_name="GeneralLedger")
    return path


def make_trial_balance(
    outdir: Path,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> Path:
    """Generer trial balance (hovedboksaldoer) med IB, PR og UB per konto."""
    tx_path = _find_csv_file(outdir, "transactions.csv")
    hdr_path = _find_csv_file(outdir, "header.csv")
    tx = _read_csv_safe(tx_path, dtype=str) if tx_path else None
    hdr = _read_csv_safe(hdr_path, dtype=str) if hdr_path else None
    if tx is None or tx.empty:
        raise FileNotFoundError("transactions.csv mangler")
    tx = _parse_dates(tx, ["TransactionDate", "PostingDate"])
    tx["Date"] = tx["PostingDate"].fillna(tx["TransactionDate"])
    tx["AccountID"] = _norm_acc_series(tx["AccountID"] if "AccountID" in tx.columns else pd.Series([], dtype=str))
    tx = _to_num(tx, ["Debit", "Credit"])
    dfrom, dto = _range_dates(hdr, date_from, date_to, tx)
    def _sum(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame({"AccountID": [], "Debit": [], "Credit": [], "GL_Amount": []})
        g = df.groupby("AccountID")[["Debit", "Credit"]].sum().reset_index()
        g["GL_Amount"] = g["Debit"] - g["Credit"]
        return g
    ib_gl = _sum(tx[tx["Date"] < dfrom]).rename(columns={"GL_Amount": "GL_IB"})
    pr_gl = _sum(tx[(tx["Date"] >= dfrom) & (tx["Date"] <= dto)]).rename(columns={"GL_Amount": "GL_PR"})
    ub_gl = _sum(tx[tx["Date"] <= dto]).rename(columns={"GL_Amount": "GL_UB"})
    tb = ub_gl.merge(ib_gl, on="AccountID", how="outer").merge(pr_gl, on="AccountID", how="outer").fillna(0.0)
    acc = _read_csv_safe(outdir / "accounts.csv", dtype=str)
    if acc is not None and "AccountID" in acc.columns:
        acc = acc.copy()
        acc["AccountID"] = _norm_acc_series(acc["AccountID"])
        if {"OpeningDebit", "OpeningCredit", "ClosingDebit", "ClosingCredit"}.issubset(acc.columns):
            acc = _to_num(acc, ["OpeningDebit", "OpeningCredit", "ClosingDebit", "ClosingCredit"])
            acc["IB_OpenNet"] = acc["OpeningDebit"] - acc["OpeningCredit"]
            acc["UB_CloseNet"] = acc["ClosingDebit"] - acc["ClosingCredit"]
            acc["PR_Accounts"] = acc["UB_CloseNet"] - acc["IB_OpenNet"]
            cols = [
                "AccountID",
                "AccountDescription",
                "IB_OpenNet",
                "PR_Accounts",
                "UB_CloseNet",
                "OpeningDebit",
                "OpeningCredit",
                "ClosingDebit",
                "ClosingCredit",
            ]
            tb = tb.merge(acc[cols], on="AccountID", how="left")
    # Reorganiser kolonner for lesbarhet
    first = ["AccountID", "AccountDescription"]
    money = [
        c
        for c in [
            "IB_OpenNet",
            "PR_Accounts",
            "UB_CloseNet",
            "GL_IB",
            "GL_PR",
            "GL_UB",
            "ClosingDebit",
            "ClosingCredit",
        ]
        if c in tb.columns
    ]
    other = [c for c in tb.columns if c not in first + money]
    out = tb[first + money + other]
    path = outdir / "trial_balance.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        out.sort_values("AccountID").to_excel(writer, index=False, sheet_name="TrialBalance")
    return path


# ---------------- GUI wrapper for subledger generation ----------------
def _run_full_process(input_path: Path, outdir: Path) -> None:
    """Kjør full prosess: parse SAF‑T, lag grunnlags-CSV og generer rapporter.

    Denne funksjonen bruker parse_saft fra saft_parser_pro til å lese en
    SAF‑T‑fil (XML eller ZIP) og skriver alle CSV‑filer til outdir/csv.
    Deretter genereres både AR- og AP‑subledger, samt general ledger og
    trial balance, basert på data i csv-mappen. Rapportene lagres i
    outdir/excel. Etter fullføring skrives det en melding på standardutgang.
    """
    if parse_saft is None:
        raise RuntimeError("parse_saft er ikke tilgjengelig; saft_parser_pro mangler eller er korrupt")
    # Sørg for underkataloger
    csv_dir = outdir / "csv"
    excel_dir = outdir / "excel"
    csv_dir.mkdir(parents=True, exist_ok=True)
    excel_dir.mkdir(parents=True, exist_ok=True)
    # Kjør parsing
    parse_saft(input_path, csv_dir)
    # Generer subledger for AR og AP
    ar_path = make_subledger(csv_dir, "AR")
    ap_path = make_subledger(csv_dir, "AP")
    # Generer hovedbok og trial balance
    gl_path = make_general_ledger(csv_dir)
    tb_path = make_trial_balance(csv_dir)
    # Flytt Excel-filer til excel_dir
    for p in [ar_path, ap_path, gl_path, tb_path]:
        try:
            dest = excel_dir / p.name
            # Hvis filen finnes fra før, slett den før flytting
            if dest.exists():
                dest.unlink()
            os.replace(p, dest)
        except Exception:
            # Hvis flytting mislykkes, la filen ligge i csv_dir
            continue
    print(f"Ferdig! Rapporter generert i '{excel_dir}' og CSV i '{csv_dir}'.")


def launch_subledger_gui() -> None:
    """Vis en GUI for å generere subledger og/eller full prosess.

    Hvis tkinter er tilgjengelig, åpnes et vindu med knapper for å generere
    subledger (AR/AP) fra eksisterende CSV‑mapper, eller kjøre en full
    prosess (parsing + rapporter). Hvis tkinter mangler, gis en beskjed på
    standardutgangen.
    """
    if _tk is None:
        print(
            "GUI ikke tilgjengelig i dette miljøet. Kjør med kommandolinjeargumenter eller installer tkinter."
        )
        return

    root = _tk.Tk()
    root.title("SAF‑T verktøy")

    # Subledger generator (AR eller AP)
    def _run_subledger(which: str) -> None:
        directory = _filedialog.askdirectory(title="Velg mappe med SAF‑T CSV‑filer")
        if not directory:
            return
        try:
            out = make_subledger(Path(directory), which)
            _messagebox.showinfo("Ferdig", f"{which.upper()} subledger generert:\n{out}")
        except Exception as exc:
            _messagebox.showerror("Feil", str(exc))

    # Full prosess: parse SAF‑T og generer rapporter
    def _run_full() -> None:
        file_path = _filedialog.askopenfilename(
            title="Velg SAF‑T fil (.xml eller .zip)", filetypes=[("SAF‑T/XML/ZIP", "*.xml *.zip")]
        )
        if not file_path:
            return
        directory = _filedialog.askdirectory(title="Velg output-rotmappe")
        if not directory:
            return
        try:
            _run_full_process(Path(file_path), Path(directory))
            _messagebox.showinfo(
                "Ferdig", f"Full prosess fullført. Rapporter er lagret i '{Path(directory)/'excel'}'."
            )
        except Exception as exc:
            _messagebox.showerror("Feil", str(exc))

    # Layout
    btn_full = _tk.Button(root, text="Kjør full prosess (parser + rapporter)", command=_run_full, width=40)
    btn_full.pack(padx=20, pady=10)
    btn_ar = _tk.Button(root, text="Generer AR (kunder)", command=lambda: _run_subledger("AR"), width=40)
    btn_ar.pack(padx=20, pady=10)
    btn_ap = _tk.Button(root, text="Generer AP (leverandører)", command=lambda: _run_subledger("AP"), width=40)
    btn_ap.pack(padx=20, pady=10)
    btn_quit = _tk.Button(root, text="Lukk", command=root.destroy, width=40)
    btn_quit.pack(padx=20, pady=10)

    root.mainloop()


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description=(
            "Generate SAF‑T reports. Parse SAF‑T, generate AR/AP subledger and ledgers. "
            "If no arguments are given, a GUI will open."
        )
    )
    # Input SAF‑T-fil (xml/zip). Hvis angitt sammen med --full, kjøres full prosess.
    parser.add_argument(
        "--input",
        type=str,
        default=None,
        help="Path to SAF‑T .xml or .zip file for parsing. Used with --full.",
    )
    # outdir er valgfri rotmappe for data og rapporter; standard er nåværende arbeidskatalog
    parser.add_argument(
        "--outdir",
        type=str,
        default=".",
        help="Output root directory for CSV and Excel. Defaults to current working directory.",
    )
    # Valg av subledger-type når man kjører uten full prosess
    parser.add_argument(
        "--which",
        type=str,
        choices=["AR", "AP"],
        default="AR",
        help="Generate subledger for AR (customers) or AP (suppliers)",
    )
    parser.add_argument(
        "--date_from", type=str, default=None, help="Start date (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--date_to", type=str, default=None, help="End date (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help=(
            "Run full process: parse SAF‑T (requires --input) and generate AR/AP subledger, general ledger and trial balance. "
            "CSV files are written to outdir/csv and Excel reports to outdir/excel."
        ),
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Start GUI for full process or subledger generation. Overrides other options.",
    )
    args = parser.parse_args()

    # Hvis GUI er spesifisert, eller ingen argumenter (kun skriptnavn), vis GUI
    if args.gui or (len(sys.argv) == 1):
        launch_subledger_gui()
    else:
        outdir = Path(args.outdir)
        if args.full:
            # Full prosess krever input-fil
            if not args.input:
                parser.error("--full krever at du oppgir --input med SAF‑T-fil")
            try:
                _run_full_process(Path(args.input), outdir)
            except Exception as exc:
                print(f"Feil i full prosess: {exc}")
        else:
            # Kun subledger-generering
            try:
                out_path = make_subledger(outdir, args.which, args.date_from, args.date_to)
                print(f"Ferdig! Genererte {args.which.upper()} subledger i '{out_path}'.")
            except FileNotFoundError as exc:
                print(f"Feil: {exc}")
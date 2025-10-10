#!/usr/bin/env python3
# pdf2excel_hovedbok_transaksjoner_v4.py
# --------------------------------------
# Trekker Debet/Kredit‑transaksjoner fra Visma‑hovedbok‑PDF til Excel.

from __future__ import annotations
import argparse, logging, re
from pathlib import Path
from typing import Dict, Any, List

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# ── Regex’er ──────────────────────────────────────────────────────────────────
DATE_RE    = re.compile(r"^\d{2}\.\d{2}\.(\d{2}|\d{4})$")           # dd.mm.yy/yyyy
AMOUNT_RE  = re.compile(r"-?\d[\d .\u202f\xa0]*,\d{2}")             # 514 000,00
NUM_CLEAN  = str.maketrans({" ": "", "\u202f": "", "\xa0": "",
                            ".": "", ",": "."})
SKIP_RE    = re.compile(r"(Saldo pr\.|Inngående balanse)", re.I)

# ── GUI‑filvelger (Tkinter) med tekst‑fallback ───────────────────────────────
def ask_for_pdf() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        fp = filedialog.askopenfilename(title="Velg Hovedboks‑PDF",
                                        filetypes=[("PDF‑filer", "*.pdf"), ("Alle filer", "*.*")])
        root.destroy()
        return Path(fp) if fp else None
    except Exception:
        pass
    pdfs = sorted(Path.cwd().glob("*.pdf"))
    if not pdfs:
        print("Ingen PDF‑filer funnet."); return None
    for i, p in enumerate(pdfs, 1):
        print(f" {i}) {p.name}")
    try:
        idx = int(input("Nummer [0=avbryt]: "))
        return pdfs[idx-1] if idx else None
    except (ValueError, IndexError):
        return None

# ── Beløps‑hjelper ───────────────────────────────────────────────────────────
def to_float(txt: str) -> float | None:
    if AMOUNT_RE.fullmatch(txt):
        return float(txt.translate(NUM_CLEAN))
    return None

# ── Hoved‑parser for én rad ──────────────────────────────────────────────────
def parse_row(raw: str,
              col_pos: Dict[str, int]) -> Dict[str, Any] | None:
    """
    Bruk kolonne‑posisjoner fra headeren til å bestemme Debet/Kredit.
    """
    if SKIP_RE.search(raw):
        return None

    # ─ 1) Trekk ut beløp i hele raden (kan være 1–3) ────────────────────────
    amounts = list(AMOUNT_RE.finditer(raw))
    if not amounts:
        return None

    # ─ 2) Finn fire siste *heltall* før første beløp  ───────────────────────
    left_part  = raw[:amounts[0].start()].rstrip()
    left_toks  = re.split(r"\s+", left_part)
    numeric    = [t for t in left_toks if t.isdigit()]
    if len(numeric) < 4 or not DATE_RE.match(left_toks[0]):
        return None
    avd, prosj, prod, mva = map(int, numeric[-4:])

    bil_dato, bilagsnr, bil_art = left_toks[:3]
    tekst = " ".join(left_toks[3:-4])

    # ─ 3) Klassifiser hvert beløp etter x‑posisjon ──────────────────────────
    debet = kredit = saldo = None
    for m in amounts:
        x = m.start()
        if x >= col_pos["Kredit"]:
            # Beløp står i Kredit‑ eller Saldo‑kolonnen
            if x >= col_pos["Saldo"]:
                saldo = to_float(m.group())
            else:
                kredit = to_float(m.group())
        else:
            debet = to_float(m.group())

    # minst ett av Debet/Kredit må være satt for at dette skal være transaksjon
    if debet is None and kredit is None:
        return None

    return {
        "Bil.dato": bil_dato,
        "Bilagsnr": int(bilagsnr),
        "Bil.art":  int(bil_art),
        "Tekst":    tekst,
        "Avd": avd, "Prosj": prosj, "Prod": prod, "Mva": mva,
        "Debet": debet,
        "Kredit": kredit,
        "Saldo": saldo,   # droppes før Excel
    }

# ── PDF‑uttrekk  ─────────────────────────────────────────────────────────────
import pdfplumber, pandas as pd
try:
    from tqdm import tqdm
except ImportError:
    def tqdm(x, **_): return x  # type: ignore


def extract(pdf_path: Path) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    col_pos: Dict[str, int] = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in tqdm(pdf.pages, desc="Leser sider"):
            text = page.extract_text() or ""
            lines = text.splitlines()

            # 1) Finn headeren på siden for å ta kolonne‑posisjoner
            hdr = next((ln for ln in lines if "Bil.dato" in ln and "Debet" in ln), None)
            if hdr:
                col_pos = {col: hdr.index(col) for col in ("Debet", "Kredit", "Saldo")}

            if not col_pos:
                continue  # ingen header funnet ennå

            # 2) Parse hver linje
            for raw in lines:
                if "Bil.dato" in raw:       # hopp over headerlinje
                    continue
                row = parse_row(raw, col_pos)
                if row:
                    rows.append(row)

    return pd.DataFrame(rows)


# ── PDF → Excel  ─────────────────────────────────────────────────────────────
def pdf_to_excel(pdf_file: Path, out_file: Path) -> None:
    print(f"→ Leser {pdf_file} …")
    df = extract(pdf_file)

    if df.empty:
        print("⚠️  Fant ingen transaksjonslinjer – rapportlayouten avviker fortsatt.")
        return

    # Dato til datetime
    df["Bil.dato"] = (
        pd.to_datetime(df["Bil.dato"], format="%d.%m.%y",
                       errors="coerce", dayfirst=True)
          .fillna(pd.to_datetime(df["Bil.dato"], format="%d.%m.%Y",
                                 errors="coerce", dayfirst=True))
    )

    df = df.drop(columns="Saldo", errors="ignore")
    df = df.sort_values(["Bil.dato", "Bilagsnr"]).reset_index(drop=True)

    out_file.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_file, index=False, engine="openpyxl")
    print(f"✅ {len(df):,} transaksjoner skrevet til\n   {out_file.resolve()}")


# ── CLI  ─────────────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(description="PDF → Excel (kun Debet/Kredit)")
    ap.add_argument("--in",  dest="inp",  type=Path, help="PDF‑fil som skal konverteres")
    ap.add_argument("--out", dest="outp", type=Path, help="Excel‑fil som skal skrives")
    args = ap.parse_args()

    pdf_path = args.inp or ask_for_pdf()
    if not pdf_path:
        print("Ingen fil valgt – avbryter."); return

    out_path = (args.outp or pdf_path.with_suffix(".xlsx")).resolve()
    pdf_to_excel(pdf_path, out_path)


if __name__ == "__main__":
    main()

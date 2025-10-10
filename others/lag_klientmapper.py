# -*- coding: utf-8 -*-
"""
lag_klientmapper.py
-------------------
• Leser et Excel-ark som har kolonnene
    -  «klientnummer»  (eller første kolonne)
    -  «klientnavn»    (eller andre kolonne)
• Lager en mappe for hver rad med navnet
      "<klientnummer> <klientnavn>"
  under ønsket rotkatalog.
"""

from __future__ import annotations
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


def velg_fil_vindu(tittel: str, ftyper):
    """Lite Tk-dialog-vindu for å plukke fil eller mappe."""
    root = tk.Tk(); root.withdraw()
    return filedialog.askopenfilename(title=tittel, filetypes=ftyper)


def velg_mappe_vindu(tittel: str):
    root = tk.Tk(); root.withdraw()
    return filedialog.askdirectory(title=tittel)


# ------------------------------------------------------------
def safe_name(name: str) -> str:
    """Fjerner/erstatter tegn Windows ikke liker i mappenavn."""
    name = re.sub(r"[<>:\"/\\|?*]", "_", name)     # ulovlige tegn
    name = re.sub(r"\s+", " ", name.strip())       # rydd opp whitespace
    return name


def main() -> None:
    # ---- 1  velg Excel-fil ------------------------------
    xlf = velg_fil_vindu("Velg Excel-fil med klientliste",
                         [("Excel-filer", "*.xlsx *.xls *.xlsm")])
    if not xlf:
        return

    # ---- 2  velg rot-mappe -------------------------------
    base_dir = velg_mappe_vindu("Velg rotmappe der klientmapper skal lages")
    if not base_dir:
        return
    base_dir = Path(base_dir)

    # ---- 3  les Excel ------------------------------------
    df = pd.read_excel(xlf)
    # bruk eksplisitt kolonnenavn om de finnes, ellers første to kolonner
    kol_num = next((c for c in df.columns if "nummer" in str(c).lower()), df.columns[0])
    kol_navn = next((c for c in df.columns if "navn"   in str(c).lower()), df.columns[1])

    ant = 0
    for nr, navn in zip(df[kol_num], df[kol_navn]):
        if pd.isna(nr) or pd.isna(navn):
            continue
        mappenavn = safe_name(f"{int(nr)} {str(navn)}")
        dest = base_dir / mappenavn
        try:
            dest.mkdir(parents=True, exist_ok=True)
            ant += 1
        except Exception as e:
            print(f"Skippet «{mappenavn}»: {e}")

    messagebox.showinfo("Ferdig", f"Oprettet {ant} mapper i\n{base_dir}")


if __name__ == "__main__":
    main()

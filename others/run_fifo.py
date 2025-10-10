# -*- coding: utf-8 -*-
"""
fifo_report.py  –  lager full FIFO-rapport pr. verdipapir
---------------------------------------------------------

Krav: pandas, openpyxl  (pip install pandas openpyxl)
Skal ligge i samme mappe som fifo_fifo_function.py
"""

from pathlib import Path
import pandas as pd
from others.fifo_fifo_function import fifo      # funksjonen du testet tidligere

# --------------------------------------------------------------------------
# 1)  ANGI FULL FILSTI TIL ARBEIDSBOKEN HER!
# --------------------------------------------------------------------------
xl_path = Path(
    r"F:\Dokument\7\REGN\3 - Visindi regnskapskunder\3123 Atle Ronglan AS\Regnskap\2024\Transaksjoner 2024 - revisor AP – Python.xlsx"
)
# --------------------------------------------------------------------------

if not xl_path.exists():
    raise FileNotFoundError(f"Fant ikke filen:\n{xl_path}")

print(f"Leser {xl_path}")

# ---------- 2  Les fanen --------------------------------------------------
raw = pd.read_excel(xl_path, sheet_name="Tabell")

# ---------- 3  Rens BELØP + lag Netto ------------------------------------
raw["BELØP"] = (
    raw["BELØP"]
    .astype(str)
    .str.replace("kr", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

out_types = ["Tegning", "Uttak", "Belastning av utestående honorar"]
raw["Netto"] = raw.apply(
    lambda r: -r["BELØP"] if r["Type"] in out_types else r["BELØP"], axis=1
)

# ---------- 4  Bygg DataFrame til FIFO -----------------------------------
fifo_df = pd.DataFrame(
    {
        "Dato": pd.to_datetime(raw["DATO"]),
        "Selskap": raw["Produkt"],        # verdipapirnavn
        "Eier": raw["portfolio"],         # portefølje/navn
        "Netto": raw["Netto"],
        "Pris pr aksje": 1.0,             # kontant-FIFO => pris = 1
    }
)

# ---------- 5  Kjør FIFO --------------------------------------------------
realised_df, closing_df = fifo(fifo_df)

# ---------- 6  Aggreger rapport ------------------------------------------
agg_flow = (
    fifo_df.groupby("Selskap")
           .agg(
               Kjøp_Qty=("Netto", lambda s: s[s > 0].sum()),
               Salg_Qty=("Netto", lambda s: -s[s < 0].sum()),
           )
)

agg_real = (
    realised_df.groupby("Selskap")
               .agg(
                   Real_Qty=("QtySold", "sum"),
                   Real_Cost=("Cost", "sum"),
                   Real_Proceeds=("Proceeds", "sum"),
                   Real_PnL=("PnL", "sum"),
               )
)

agg_ub = (
    closing_df.groupby("Selskap")
              .agg(
                  UB_Qty=("Qty", "sum"),
                  UB_Cost=("CostPerShare",
                           lambda s: (closing_df.loc[s.index, "Qty"] * s).sum()),
              )
)
agg_ub["Snittkost_UB"] = agg_ub["UB_Cost"] / agg_ub["UB_Qty"]

report = (
    agg_flow.join(agg_real, how="outer")
            .join(agg_ub,   how="outer")
            .fillna(0)
            .reset_index()
            .sort_values("Selskap")
)

# ---------- 7  Skriv tre nye faner ---------------------------------------
with pd.ExcelWriter(
        xl_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
    realised_df.to_excel(xw, sheet_name="FIFO_Realisert", index=False)
    closing_df.to_excel(xw,  sheet_name="FIFO_Urealisert", index=False)
    report.to_excel(xw,      sheet_name="FIFO_Rapport",   index=False)

print("✓ FIFO-rapport fullført – se fanene FIFO_Realisert / Urealisert / Rapport.")

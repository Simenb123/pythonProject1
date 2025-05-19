# bilag_uttrekk_med_analyse.py
# -------------------------------------------------------------
#  Henter en Excel/CSV-fil → filtrerer → trekker bilag
#  Skriver:
#     • Bilag_uttrekk_<n>.xlsx   (3 faner)
#     • Populasjonsanalyse.xlsx  (3 faner)
#     • Histogram.png
# -------------------------------------------------------------

from pathlib import Path
import re
import pandas as pd
import chardet
import numpy as np

# matplotlib uten GUI
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from pandas import ExcelWriter


# ---------- POPULASJONS-ANALYSE -----------------------------------------
def analyse_blokk(df_pop: pd.DataFrame,
                  konto_kol: str,
                  belop_kol: str,
                  src_path: Path) -> None:

    # 1) Overordnede tall
    pop_stats = (
        df_pop[belop_kol]
        .agg(Antall="count", Sum="sum", Min="min",
             Q1=lambda s: s.quantile(0.25),
             Median="median", Mean="mean",
             Q3=lambda s: s.quantile(0.75),
             Max="max", StdAvvik="std")
        .to_frame("Verdi")
    )

    # 2) Beløpsbånd
    bins   = [-np.inf, 1_000, 10_000, 100_000, np.inf]
    labels = ["< 1 000", "1 000–9 999", "10 000–99 999", "≥ 100 000"]
    df_tmp = df_pop.copy()
    df_tmp["Beløpsbånd"] = pd.cut(df_tmp[belop_kol].abs(), bins=bins, labels=labels)

    size_bands = (
        df_tmp.groupby("Beløpsbånd")[belop_kol]
        .agg(Antall="count", Sum="sum")
        .reset_index()
        .sort_values("Beløpsbånd")
    )
    size_bands["%Antall"] = size_bands["Antall"].cumsum()/size_bands["Antall"].sum()
    size_bands["%Sum"]    = size_bands["Sum"].cumsum()/size_bands["Sum"].sum()

    # 3) Konto × måned-pivot hvis mulig
    konto_pivot = pd.DataFrame()          # tom som fallback
    if ("Dato" in df_tmp.columns and
            np.issubdtype(df_tmp["Dato"].dtype, np.datetime64)):
        df_tmp["Måned"] = df_tmp["Dato"].dt.to_period("M")
        konto_pivot = (
            df_tmp.pivot_table(index=konto_kol, columns="Måned",
                               values=belop_kol, aggfunc=["count", "sum"],
                               fill_value=0)
        )
    else:
        print("⚠️  Kolonnen 'Dato' er ikke datetime – hopper over Konto_pivot.")

    # 4) Histogram
    fig_path = src_path.with_name("Histogram.png")
    plt.figure(figsize=(7,4))
    plt.hist(np.log10(df_pop[belop_kol].abs()+1), bins=30, edgecolor="black")
    plt.title("Histogram log10(|beløp|)")
    plt.xlabel("log10(|beløp|)")
    plt.ylabel("Antall linjer")
    plt.tight_layout(); plt.savefig(fig_path, dpi=150); plt.close()

    # 5) Skriv arbeidsbok
    ana_path = src_path.with_name("Populasjonsanalyse.xlsx")
    with ExcelWriter(ana_path, engine="openpyxl") as xw:
        pop_stats.to_excel(xw, sheet_name="Pop_stats")
        size_bands.to_excel(xw, sheet_name="Size_bands", index=False)
        if not konto_pivot.empty:
            konto_pivot.to_excel(xw, sheet_name="Konto_pivot")
    print(f"✓ Analyse lagret til {ana_path}")
    print(f"  ‣ Histogram lagret til {fig_path}")


# ---------- HJELPEFUNKSJONER --------------------------------------------
def les_fil(path: Path) -> pd.DataFrame:
    suf = path.suffix.lower()
    if suf in {".xlsx", ".xls"}:
        return pd.read_excel(path, engine="openpyxl")
    if suf == ".csv":
        try:
            return pd.read_csv(path, sep=";", encoding="utf-8-sig")
        except (UnicodeDecodeError, pd.errors.ParserError):
            pass
        enc = chardet.detect(path.read_bytes())["encoding"] or "latin1"
        for sep in (";", ","):
            try:
                return pd.read_csv(path, sep=sep, encoding=enc)
            except pd.errors.ParserError:
                continue
    raise ValueError("Filen må være .xlsx, .xls eller .csv")

def gjett_kolonner(cols: list[str]) -> dict[str,str]:
    low = [c.lower() for c in cols]
    def first(pats):
        return next((c for c,l in zip(cols,low)
                     for p in pats if re.search(p,l)), "")
    return dict(
        bilag = first([r"bilag", r"voucher", r"dok"]),
        konto = first([r"konto.*nr|kontonummer"]),
        belop = first([r"bel[oø]p|amount|sum"])
    )

def til_float(s: pd.Series) -> pd.Series:
    return (s.astype(str)
             .str.replace(" ", "", regex=False)
             .str.replace("kr", "", regex=False)
             .str.replace(".", "", regex=False)
             .str.replace(",", ".", regex=False)
             .astype(float))


# ---------- MAIN --------------------------------------------------------
def main():
    # --- 1) filsti -------------------------------------------------------
    src = Path(input("Lim inn full sti til filen (Excel/CSV):\n> ").strip('" '))
    if not src.exists():
        print("Fant ikke filen – avslutter."); return

    df = les_fil(src)
    print("Kolonner funnet:", list(df.columns))

    # --- 2) kolonnevalg --------------------------------------------------
    g = gjett_kolonner(list(df.columns))
    ok = input(f"Gjetter\n  Bilag : {g['bilag']}\n  Konto : {g['konto']}\n"
               f"  Beløp : {g['belop']}\nStemmer? [Y/n] ").lower() or "y"
    if ok == "y":
        bilag_kol, konto_kol, belop_kol = g.values()
    else:
        bilag_kol = input("Kolonne for bilagsnr:\n> ").strip()
        konto_kol = input("Kolonne for kontonr:\n> ").strip()
        belop_kol = input("Kolonne for beløp:\n> ").strip()

    # --- 3) intervaller --------------------------------------------------
    lo_k, hi_k = map(int, input("Kontointervall (f.eks. 6000-7999):\n> ").split("-"))
    bel = input("Beløpsintervall (f.eks. -10000,10000 eller Enter):\n> ")
    lo_b, hi_b = (map(float, bel.split(",")) if bel else (float("-inf"), float("inf")))

    # --- 4) rens tall + konverter dato ----------------------------------
    df[belop_kol] = til_float(df[belop_kol])
    df[konto_kol] = (df[konto_kol].astype(str)
                                  .str.extract(r"(\d+)", expand=False)
                                  .astype(int))

    # konverter første kolonne som ligner «dato»
    for col in df.columns:
        if re.search(r"dat(o|e)", col.lower()):
            df[col] = pd.to_datetime(df[col], errors="coerce", format="ISO8601")
            break

    # --- 5) filtrér ------------------------------------------------------
    mask = (df[konto_kol].between(lo_k, hi_k) &
            df[belop_kol].between(lo_b, hi_b))
    df_filt = df[mask]
    if df_filt.empty:
        print("Ingen rader passer filtrene – stopper."); return

    # --- 6) trekk bilag --------------------------------------------------
    unike = df_filt[bilag_kol].drop_duplicates()
    n = int(input(f"Hvor mange bilag vil du trekke? (1–{len(unike)})\n> "))
    valgte = unike.sample(n=n, random_state=None)

    fullt  = df[df[bilag_kol].isin(valgte)].copy()
    inter  = df_filt[df_filt[bilag_kol].isin(valgte)].copy()

    summer = (
        inter.groupby(bilag_kol)[belop_kol]
        .agg(Sum_i_intervallet="sum", Linjer_i_intervallet="count")
        .reset_index()
    )

    for d in (fullt, inter, summer):
        try: d[bilag_kol] = d[bilag_kol].astype(int)
        except ValueError: pass
        d.sort_values(bilag_kol, inplace=True)

    # --- 7) lagre utvalg -------------------------------------------------
    out = src.with_name(f"Bilag_uttrekk_{n}.xlsx")
    with ExcelWriter(out, engine="openpyxl") as xw:
        fullt.to_excel(xw,  "Fullt_bilagsutvalg", index=False)
        inter.to_excel(xw,  "Kun_intervallet",    index=False)
        summer.to_excel(xw, "Bilag_summer",       index=False)
    print(f"\n✓ Bilagsuttrekk lagret til {out}")

    # --- 8) analyse ------------------------------------------------------
    analyse_blokk(df_filt, konto_kol, belop_kol, src)


# ------------------------------------------------------------------------
if __name__ == "__main__":
    main()

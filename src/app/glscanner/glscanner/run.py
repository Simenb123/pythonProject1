# run.py  (legg denne fila i prosjektroten, på samme nivå som "src")
from __future__ import annotations
import sys, json, argparse
from pathlib import Path
import pandas as pd

# --- Pek Python til pakken uten å installere noe ---
# Juster dette hvis mappestrukturen din er annerledes.
PKG_ROOT = Path(__file__).resolve().parent / "src" / "app" / "glscanner"
if not (PKG_ROOT / "glscanner" / "__init__.py").exists():
    raise SystemExit(f"Fant ikke pakken under {PKG_ROOT}. Sjekk stien i run.py.")

sys.path.insert(0, str(PKG_ROOT))  # nå funker: from glscanner import ...

from glscanner import Analyzer, GLScanConfig, ColumnMap, build_excel_report  # type: ignore

def read_df(path: str | None, sep: str | None = None) -> pd.DataFrame:
    if not path:
        # Demo-datasett hvis du ikke oppgir --input
        return pd.DataFrame({
            "Konto": ["4000","4000","4000","2710","2710","6000"],
            "Tekst": ["Faktura A","Faktura A","Faktura B","MVA","MVA","Reise"],
            "Dato": pd.to_datetime([
                "2024-01-29 21:00","2024-01-29 21:05","2024-01-15 10:00",
                "2024-01-31 22:00","2024-02-01 09:00","2024-01-28 12:00"
            ]),
            "Bilag": ["1001","1001","1002","9999","10000","3001"],
            "MVAkode": ["25","25","25","0","0","12"],
            "MVAbelop": [250,250,125,0,0,0],
            "Belop": [10000,10000,5000,0,0,1200],
            "Leverandor": ["A","A","A","-","-","B"]
        })

    p = Path(path)
    suf = p.suffix.lower()
    if suf in {".parquet",".pq",".pqt"}:
        return pd.read_parquet(p)
    if suf in {".xlsx",".xls"}:
        return pd.read_excel(p)
    # CSV
    return pd.read_csv(p, sep=(sep or None), engine="python", encoding="utf-8-sig")

def main():
    ap = argparse.ArgumentParser(description="Kjør glscanner uten install (run.py).")
    ap.add_argument("--input", help="CSV/Parquet/Excel. Hopper over for demo.", default=None)
    ap.add_argument("--excel-out", help="Excel-rapportfil.", default="gl_report.xlsx")
    ap.add_argument("--map", help='JSON mapping: {"konto":"Konto","tekst":"Tekst","dato":"Dato","bilag":"Bilag","mvakode":"MVAkode","mvabelop":"MVAbelop","belop":"Belop","leverandor":"Leverandor"}', default=None)
    ap.add_argument("--sep", help="Separator for CSV (overstyr autodetect).", default=None)
    args = ap.parse_args()

    df = read_df(args.input, sep=args.sep)

    mapping = {}
    if args.map:
        mapping = {str(k): str(v) for k, v in json.loads(args.map).items()}

    colmap = ColumnMap(
        konto=mapping.get("konto","Konto" if "Konto" in df.columns else "konto"),
        tekst=mapping.get("tekst","Tekst" if "Tekst" in df.columns else "tekst"),
        dato=mapping.get("dato","Dato" if "Dato" in df.columns else "dato"),
        bilag=mapping.get("bilag","Bilag" if "Bilag" in df.columns else "bilag"),
        mvakode=mapping.get("mvakode","MVAkode" if "MVAkode" in df.columns else "mvakode"),
        mvabelop=mapping.get("mvabelop","MVAbelop" if "MVAbelop" in df.columns else "mvabelop"),
        belop=mapping.get("belop","Belop" if "Belop" in df.columns else "belop"),
        leverandor=mapping.get("leverandor","Leverandor" if "Leverandor" in df.columns else None),
    )

    cfg = GLScanConfig(column_map=colmap)
    analyzer = Analyzer(cfg)
    findings = analyzer.analyze(df)

    print("\n=== Topp 25 funn ===")
    print(findings.sort_values("total_score", ascending=False)
                 .head(25)[["dato","bilag","konto","tekst","belop","rule","reason","total_score"]]
                 .to_string(index=False))

    build_excel_report(findings, args.excel_out)
    print(f"\n✔ Skrev rapport: {Path(args.excel_out).resolve()}")

if __name__ == "__main__":
    main()

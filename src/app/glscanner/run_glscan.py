
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
run_glscan.py — enkel kjørbar "runner" for glscanner

Bruk:
  python run_glscan.py --input hovedbok.csv --excel-out gl_report.xlsx --map '{"konto":"Konto","tekst":"Tekst","dato":"Dato","bilag":"Bilag","mvakode":"MVAkode","mvabelop":"MVAbelop","belop":"Belop"}'

  # Demo uten input
  python run_glscan.py --demo --excel-out demo_report.xlsx
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, Optional

import pandas as pd


# ------------------------------------------------------------
# Finn glscanner enten installert eller i kildekoden din
# ------------------------------------------------------------
def _try_import_glscanner():
    try:
        import glscanner  # type: ignore
        return glscanner
    except Exception:
        pass

    # Prøv å finne pakken i prosjektet (typisk src/app/glscanner/glscanner)
    here = Path(__file__).resolve()
    candidates = []
    for base in [here, *here.parents]:
        candidates.extend([
            base / "src" / "app" / "glscanner",
            base / "src",
            base,
        ])

    for candidate in candidates:
        pkg_dir = candidate / "glscanner" / "__init__.py"
        if pkg_dir.exists():
            sys.path.insert(0, str(candidate))
            try:
                import glscanner  # type: ignore
                return glscanner
            except Exception:
                continue

    print("❌ Fikk ikke importert 'glscanner'. Installer med:\n"
          "    pip install -e <path til mappen med pyproject.toml>\n"
          "eller kjør scriptet fra et sted der 'src/app/glscanner' er på PYTHONPATH.", file=sys.stderr)
    sys.exit(1)


glscanner = _try_import_glscanner()
from glscanner import GLScanConfig, ColumnMap, Analyzer, build_excel_report  # type: ignore


# ------------------------------------------------------------
# Hjelpefunksjoner
# ------------------------------------------------------------
def _read_df(input_path: Optional[str], sep: Optional[str] = None) -> pd.DataFrame:
    if not input_path:
        return _demo_df()

    p = Path(input_path)
    if not p.exists():
        print(f"❌ Finner ikke fil: {p}", file=sys.stderr)
        sys.exit(2)

    suf = p.suffix.lower()
    if suf in {".parquet", ".pq", ".pqt"}:
        return pd.read_parquet(p)
    if suf in {".xlsx", ".xls"}:
        return pd.read_excel(p)
    # CSV (forsøk auto-separator)
    if sep is None:
        return pd.read_csv(p, sep=None, engine="python", encoding="utf-8-sig")
    return pd.read_csv(p, sep=sep, encoding="utf-8-sig")


def _demo_df() -> pd.DataFrame:
    return pd.DataFrame({
        "Konto": ["4000","4000","4000","2710","2710","6000"],
        "Tekst": ["Faktura A","Faktura A","Faktura B","MVA","MVA","Reise"],
        "Dato": pd.to_datetime(["2024-01-29 21:00","2024-01-29 21:05","2024-01-15 10:00","2024-01-31 22:00","2024-02-01 09:00","2024-01-28 12:00"]),
        "Bilag": ["1001","1001","1002","9999","10000","3001"],
        "MVAkode": ["25","25","25","0","0","12"],
        "MVAbelop": [250,250,125,0,0,0],
        "Belop": [10000,10000,5000,0,0,1200],
        "Leverandor": ["A","A","A","-","-","B"]
    })


def _maybe_parse_map(map_str: Optional[str]) -> Dict[str, str]:
    if not map_str:
        return {}
    try:
        d = json.loads(map_str)
        if not isinstance(d, dict):
            raise ValueError("Mapping må være et JSON-objekt.")
        return {str(k): str(v) for k, v in d.items()}
    except Exception as e:
        print(f"❌ Kunne ikke lese --map som JSON: {e}", file=sys.stderr)
        sys.exit(3)


def _infer_colmap(df: pd.DataFrame) -> Dict[str, str]:
    """Best-effort kolonne-gjetting. Fanger vanlige norske/engelske navn."""
    import re
    def norm(s: str) -> str:
        s = s.lower().strip()
        s = s.replace("ø","o").replace("å","a").replace("æ","ae")
        return re.sub(r"[^a-z0-9]+", "", s)

    index = {norm(c): c for c in df.columns}

    cand = {
        "konto": ["konto","account","kontonr","kontonummer","acct"],
        "tekst": ["tekst","beskrivelse","description","text","posteringstekst"],
        "dato": ["dato","date","bilagsdato","postdate","transdate","bokfdato","posteringstidspunkt"],
        "bilag": ["bilag","bilagsnr","bilagsnummer","voucher","document","journal","doknr","dok"],
        "belop": ["belop","belop_","belp","amount","netto","belopinklmva","amountnok"],
        "mvakode": ["mvakode","mva_kode","vatcode","mva","mva-kode","mva%","sats"],
        "mvabelop": ["mvabelop","mvabelop_","vatamount","mva_belop","mva-belop","mvaamount"],
        "leverandor": ["leverandor","leverandorid","supplier","suppliername","leverandor_navn"],
        "kundenr": ["kundenr","customer","customerid","kundennr","kunnr"],
    }

    out = {}
    for std, names in cand.items():
        for n in names:
            if n in index:
                out[std] = index[n]
                break
    return out


# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Kjør glscanner på en CSV/Parquet/Excel-fil eller demo-data.")
    ap.add_argument("--input", help="Sti til hovedbok (CSV/Parquet/Excel). Hvis utelatt, brukes demo.", default=None)
    ap.add_argument("--excel-out", help="Skriv Excel-rapport til denne filen.", default="gl_report.xlsx")
    ap.add_argument("--config", help="YAML/JSON konfigfil for GLScanConfig.", default=None)
    ap.add_argument("--map", help="JSON mapping fra dine kolonnenavn til standard (konto,tekst,dato,bilag,mvakode,mvabelop,belop,leverandor).", default=None)
    ap.add_argument("--sep", help="Separator for CSV (overstyr auto-detect).", default=None)
    ap.add_argument("--demo", action="store_true", help="Kjør på innebygd demodatasett. Overstyrer --input.")
    args = ap.parse_args()

    # 1) Les data
    df = _read_df(None if args.demo else args.input, sep=args.sep)

    # 2) ColumnMap (fra --map eller heuristisk gjetting)
    user_map = _maybe_parse_map(args.map)
    inferred_map = _infer_colmap(df)
    # bruker-bruk går foran gjetting
    final_map = {**inferred_map, **user_map}

    colmap = ColumnMap(
        konto=final_map.get("konto", "konto"),
        tekst=final_map.get("tekst", "tekst"),
        dato=final_map.get("dato", "dato"),
        bilag=final_map.get("bilag", "bilag"),
        mvakode=final_map.get("mvakode", "mvakode"),
        mvabelop=final_map.get("mvabelop", "mvabelop"),
        belop=final_map.get("belop", "belop"),
        leverandor=final_map.get("leverandor"),
        kundenr=final_map.get("kundenr")
    )

    # 3) Config
    cfg = GLScanConfig(column_map=colmap)
    if args.config:
        # last YAML/JSON og oppdater cfg
        import yaml, json
        p = Path(args.config)
        text = p.read_text(encoding="utf-8")
        try:
            data = yaml.safe_load(text) if p.suffix.lower() in {".yml",".yaml"} else json.loads(text)
        except Exception as e:
            print(f"❌ Klarte ikke å lese config {p}: {e}", file=sys.stderr)
            sys.exit(4)
        cfg = GLScanConfig(**data, column_map=colmap)

    # 4) Analyse
    analyzer = Analyzer(cfg)
    findings = analyzer.analyze(df)

    # 5) Print kort oppsummering
    print("\n=== Topp 25 funn ===")
    print(findings.sort_values("total_score", ascending=False)
                  .head(25)[["dato","bilag","konto","tekst","belop","rule","reason","total_score"]]
                  .to_string(index=False))

    # 6) Lag Excel-rapport
    path = Path(args.excel_out)
    build_excel_report(findings, str(path))
    print(f"\n✔ Skrev rapport: {path.resolve()}")

if __name__ == "__main__":
    main()

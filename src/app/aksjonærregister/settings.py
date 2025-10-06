from __future__ import annotations
import os, json

PKG_DIR = os.path.dirname(__file__)
DB_PATH = os.path.join(PKG_DIR, "aksjonaerregister.duckdb")
META_PATH = os.path.join(PKG_DIR, "build_meta.json")

# Valgfritt: sett default CSV for auto-bygg ved første oppstart
CSV_PATH = ""  # f.eks. r"C:\data\aksjeeiebok_2024.csv"

# Brukes av detect.latest_csv_in_dir()
CSV_PATTERN = "*.csv"

# CSV-en din bruker semikolon
DELIMITER = ";"

# Kolonnemapping som passer din CSV-header:
# Orgnr;Selskap;Aksjeklasse;Navn aksjonær;Fødselsår/orgnr;Postnr/sted;Landkode;Antall aksjer;Antall aksjer selskap
COLUMN_MAP = {
    "company_orgnr": "Orgnr",
    "company_name":  "Selskap",
    "owner_orgnr":   "Fødselsår/orgnr",
    "owner_name":    "Navn aksjonær",
    # Ingen egen prosent i CSV → beregn fra antall
    "ownership_pct": "__COMPUTE_FROM_COUNTS__",
}

# Til beregning av eierandel
COUNT_COLUMNS = {
    "shares_owner":   "Antall aksjer",
    "shares_company": "Antall aksjer selskap",
}

# Standard dybder i graf
MAX_DEPTH_UP = 3
MAX_DEPTH_DOWN = 2

def load_meta() -> dict:
    try:
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_meta(meta: dict) -> None:
    tmp = META_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    os.replace(tmp, META_PATH)

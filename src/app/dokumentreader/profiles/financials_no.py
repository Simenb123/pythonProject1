from __future__ import annotations
from typing import Any, Dict, List, Tuple
import re
import os
import pandas as pd
import pdfplumber

from .base import DocProfile, as_result

HEADINGS = {
    "resultat": ["Resultatregnskap", "Resultat", "Resultat oppstilling"],
    "balanse": ["Balanse", "Balanseregnskap", "Balansens"],
    "egenkapital": ["Egenkapitaloppstilling", "Endringer i egenkapital"],
    "kontant": ["Kontantstrøm", "Kontantstrømoppstilling"],
}

# grove kolonne-mønstre for talltabeller
MONEY_RE = r"\d{1,3}(?:[ .\u00A0\u202F]\d{3})*(?:,\d{2})?"

def _find_heading_lines(text: str) -> Dict[str, List[int]]:
    idx = {}
    lines = text.splitlines()
    for key, titles in HEADINGS.items():
        pos = []
        for i, ln in enumerate(lines):
            ln_clean = ln.strip().lower()
            if any(t.lower() in ln_clean for t in titles):
                pos.append(i)
        if pos:
            idx[key] = pos
    return idx

def _extract_tables_pdfplumber(path: str) -> List[pd.DataFrame]:
    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                df = pd.DataFrame(table)
                if df.shape[0] >= 2 and df.shape[1] >= 2:
                    dfs.append(df)
    return dfs

def _pick_financial_tables(dfs: List[pd.DataFrame]) -> Dict[str, List[pd.DataFrame]]:
    """Klassifiser tabeller løst som resultat/balanse ved å se etter nøkkelord."""
    buckets = {"resultat": [], "balanse": [], "egenkapital": [], "kontant": [], "other": []}
    for df in dfs:
        text = " ".join(df.astype(str).fillna("").values.ravel().tolist()).lower()
        if any(k in text for k in ["resultat", "driftsinntekter", "driftsresultat"]):
            buckets["resultat"].append(df)
        elif any(k in text for k in ["balanse", "eiendeler", "egenkapital", "gjeld"]):
            buckets["balanse"].append(df)
        elif "egenkapital" in text:
            buckets["egenkapital"].append(df)
        elif "kontantstrøm" in text or "kontantstrom" in text:
            buckets["kontant"].append(df)
        else:
            buckets["other"].append(df)
    return buckets

def _normalize_number(s: str) -> float | None:
    if not s:
        return None
    s = str(s).replace("\u00A0", " ").replace("\u202F", " ")
    s = re.sub(r"[^\d,.\- ]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def _table_to_kv(df: pd.DataFrame) -> List[Tuple[str, float | None]]:
    # heuristikk: første kolonne = tekst, siste kolonne = beløp
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    df = df.dropna(how="all")
    out = []
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        val = _normalize_number(row.iloc[-1])
        if label and (val is not None):
            out.append((label, val))
    return out

class FinancialsNoProfile(DocProfile):
    name = "financials_no"
    description = "Norsk regnskapsrapport (resultat/balanse mm.)"

    def detect(self, page1_text: str, full_text: str) -> float:
        hits = 0
        for words in HEADINGS.values():
            if any(w.lower() in full_text.lower() for w in words):
                hits += 1
        # også typiske ord
        for w in ["årsregnskap", "noter", "eiendeler", "egenkapital", "driftsinntekter"]:
            if w in full_text.lower():
                hits += 0.3
        return min(1.0, hits / 3.0)

    def parse(self, path: str, page1_text: str, full_text: str) -> Dict[str, Any]:
        dfs = _extract_tables_pdfplumber(path)
        buckets = _pick_financial_tables(dfs)

        result = {
            "resultat": [],
            "balanse": [],
            "egenkapital": [],
            "kontantstrom": [],
        }
        for key in ["resultat", "balanse", "egenkapital", "kontant"]:
            for df in buckets.get(key, []):
                result["kontantstrom" if key == "kontant" else key].append(_table_to_kv(df))

        # “snutter” som toppsummer (hjelper på verifisering)
        snippet = "\n".join(full_text.splitlines()[:120])[:5000]

        return as_result("financials_no", {
            "file_name": os.path.basename(path),
            "tables": result,
            "text_snippet": snippet
        })

# gl_csv.py
from __future__ import annotations
from typing import Any, Dict, List, Tuple, Optional
import csv, io, re

def _to_float(x: Any) -> float:
    if x is None: return 0.0
    if isinstance(x,(int,float)): return float(x)
    s = str(x).strip().replace("\xa0"," ").replace("−","-").replace("–","-").replace("—","-")
    neg = s.startswith("(") and s.endswith(")")
    if neg: s = s[1:-1].strip()
    s = s.replace(" ","").replace("'","")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".","").replace(",",".")
        else:                            s = s.replace(",","")
    elif "," in s: s = s.replace(",",".")
    try: v = float(re.sub(r"[^0-9\.\-]","",s))
    except Exception: return 0.0
    return -v if neg else v

def _read_text_guess(path: str) -> tuple[str,str]:
    encs = ["utf-8-sig","utf-16","utf-16le","utf-16be","cp1252","latin-1","utf-8"]
    for enc in encs:
        try:
            with open(path,"r",encoding=enc,errors="strict") as f:
                return f.read(), enc
        except UnicodeDecodeError:
            continue
        except Exception:
            continue
    with open(path,"r",encoding="latin-1",errors="replace") as f:
        return f.read(),"latin-1"

def _find_header(fieldnames: List[str], exact: List[str], partial: List[str]) -> Optional[str]:
    mp = { (h or "").strip().lower(): h for h in fieldnames if h }
    for e in exact:
        if e in mp: return mp[e]
    for p in partial:
        for n,h in mp.items():
            if p in n: return h
    return None

def read_gl_csv(path: str) -> Tuple[List[Dict[str,Any]], Dict[str,Any]]:
    text, encoding = _read_text_guess(path)
    lines = text.splitlines()
    delim = None
    if lines and lines[0].strip().lower().startswith("sep="):
        delim = lines[0].split("=",1)[1].strip()[:1] or ";"
        text = "\n".join(lines[1:])
    if not delim:
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
            delim = dialect.delimiter
        except Exception:
            delim = ";" if sample.count(";") >= sample.count(",") else ","

    f = io.StringIO(text)
    rd = csv.DictReader(f, delimiter=delim)
    recs = list(rd); fns = rd.fieldnames or (list(recs[0].keys()) if recs else [])
    if not recs: raise ValueError("CSV ser tom ut.")

    acc_hdr  = _find_header(fns, ["konto","kontonummer","account","accountno","kontonr","account_number"], ["konto","account"])
    name_hdr = _find_header(fns, ["kontonavn","navn","name","accountname","beskrivelse","description","tekst"], ["navn","name","tekst","desc"])
    ib_hdr     = _find_header(fns, ["ib","inngaaende","ingaende","opening_balance"], ["ib","inng","open"])
    debet_hdr  = _find_header(fns, ["debet","debit"], ["debet","debit"])
    kredit_hdr = _find_header(fns, ["kredit","credit"], ["kredit","credit"])
    endr_hdr   = _find_header(fns, ["endring","bevegelse","movement","ytd","hittil","resultat"], ["endr","beveg","ytd","hittil","period","result"])
    ub_hdr     = _find_header(fns, ["ub","utgaaende","utgaende","closing_balance","ubsaldo"], ["ub","utg","clos"])
    amt_hdr    = _find_header(fns, ["saldo","balance","belop","beloep","beløp","amount","sum"], ["saldo","bel","amount","sum"])

    rows: List[Dict[str,Any]] = []
    for r in recs:
        konto = (r.get(acc_hdr) if acc_hdr else None) or ""
        navn  = (r.get(name_hdr) if name_hdr else None) or ""
        ib     = _to_float(r.get(ib_hdr, ""))     if ib_hdr     else 0.0
        debet  = _to_float(r.get(debet_hdr, ""))  if debet_hdr  else 0.0
        kredit = _to_float(r.get(kredit_hdr, "")) if kredit_hdr else 0.0
        endr   = _to_float(r.get(endr_hdr, ""))   if endr_hdr   else None
        ub     = _to_float(r.get(ub_hdr, ""))     if ub_hdr     else None
        bel    = _to_float(r.get(amt_hdr, ""))    if amt_hdr    else None

        if endr is None:
            if debet_hdr and kredit_hdr: endr = debet - kredit
            elif (ub is not None) and (ib_hdr is not None): endr = ub - ib
            else: endr = bel if bel is not None else 0.0

        if ub is None:
            if ib_hdr is not None: ub = ib + endr
            else: ub = bel if bel is not None else endr

        if bel is None: bel = ub if ub is not None else endr

        rows.append({
            "konto": str(konto).strip(), "navn": str(navn).strip(),
            "ib": ib, "debet": debet, "kredit": kredit, "endring": endr, "ub": ub, "belop": bel,
        })

    meta = {
        "encoding": encoding, "delimiter": delim,
        "account_header": acc_hdr, "name_header": name_hdr,
        "ib": ib_hdr, "debet": debet_hdr, "kredit": kredit_hdr, "endring": endr_hdr, "ub": ub_hdr,
        "amount_header": amt_hdr or ("UB" if ub_hdr else ("Endring" if (debet_hdr or ib_hdr) else "Beløp")),
    }
    return rows, meta

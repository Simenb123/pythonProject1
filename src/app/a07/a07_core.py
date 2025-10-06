# a07_core.py
from __future__ import annotations
from typing import Any, Dict, Iterable, List, Tuple
from collections import defaultdict
import re

def _to_float(x: Any) -> float:
    if x is None: return 0.0
    if isinstance(x,(int,float)): return float(x)
    s = str(x).strip().replace("\xa0"," ").replace("−","-").replace("–","-").replace("—","-")
    s = re.sub(r"(?i)\b(nok|kr)\b\.?", "", s).strip()
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

class A07Row(Dict[str,Any]): pass

class A07Parser:
    def parse(self, data: Dict[str, Any]) -> Tuple[List[A07Row], List[str]]:
        rows: List[A07Row] = []; errs: List[str] = []
        try:
            oppg = (data.get("mottatt", {}) or {}).get("oppgave", {}) or data
            virksomheter = oppg.get("virksomhet") or []
            if isinstance(virksomheter, dict): virksomheter = [virksomheter]
            for v in virksomheter:
                orgnr = str(v.get("norskIdentifikator") or v.get("organisasjonsnummer") or v.get("orgnr") or "")
                pers = v.get("inntektsmottaker") or []
                if isinstance(pers, dict): pers = [pers]
                for p in pers:
                    fnr = str(p.get("norskIdentifikator") or p.get("identifikator") or p.get("fnr") or "")
                    navn = (p.get("identifiserendeInformasjon") or {}).get("navn") or p.get("navn") or ""
                    inns = p.get("inntekt") or []
                    if isinstance(inns, dict): inns = [inns]
                    for inc in inns:
                        try:
                            fordel = str(inc.get("fordel") or "").strip().lower()
                            li = inc.get("loennsinntekt") or {}; alt = inc.get("ytelse") or inc.get("kontantytelse") or {}
                            if not isinstance(li, dict): li = {}
                            if not isinstance(alt, dict): alt = {}
                            kode = li.get("beskrivelse") or alt.get("beskrivelse") or inc.get("type") or "ukjent_kode"
                            antall = li.get("antall") if isinstance(li.get("antall"), (int,float)) else None
                            beloep = _to_float(inc.get("beloep"))
                            rows.append(A07Row(orgnr=orgnr, fnr=fnr, navn=str(navn), kode=str(kode),
                                               fordel=fordel, beloep=beloep, antall=antall))
                        except Exception as e:
                            errs.append(f"Feil linje: {e}")
        except Exception as e:
            errs.append(f"Kritisk feil: {e}")
        return rows, errs

    @staticmethod
    def oppsummerte_virksomheter(root: Dict[str,Any]) -> Dict[str,float]:
        res: Dict[str,float] = {}
        oppg = (root.get("mottatt", {}) or {}).get("oppgave", {}) or root
        ov = oppg.get("oppsummerteVirksomheter") or {}
        inn = ov.get("inntekt") or []
        if isinstance(inn, dict): inn = [inn]
        for it in inn:
            li = it.get("loennsinntekt") or {}
            if not isinstance(li, dict): li = {}
            alt = it.get("ytelse") or it.get("kontantytelse") or {}
            if not isinstance(alt, dict): alt = {}
            kode = li.get("beskrivelse") or alt.get("beskrivelse") or "ukjent_kode"
            res[str(kode)] = res.get(str(kode),0.0) + _to_float(it.get("beloep"))
        return res

def summarize_by_code(rows: Iterable[A07Row]) -> Dict[str,float]:
    out = defaultdict(float)
    for r in rows: out[str(r["kode"])] += float(r["beloep"])
    return dict(out)

def summarize_by_employee(rows: Iterable[A07Row]) -> Dict[str, Dict[str,Any]]:
    idx: Dict[str, Dict[str,Any]] = {}
    for r in rows:
        fnr = str(r["fnr"])
        d = idx.setdefault(fnr, {"navn": r.get("navn",""), "sum": 0.0, "per_kode": defaultdict(float), "antall_poster": 0})
        d["navn"] = d["navn"] or r.get("navn","")
        d["sum"] += float(r["beloep"])
        d["per_kode"][str(r["kode"])] += float(r["beloep"])
        d["antall_poster"] += 1
    for v in idx.values(): v["per_kode"] = dict(v["per_kode"])
    return idx

def validate_against_summary(rows: List[A07Row], json_root: Dict[str,Any]) -> List[Tuple[str,float,float,float]]:
    calc = summarize_by_code(rows); rep = A07Parser.oppsummerte_virksomheter(json_root)
    out = []
    for code in sorted(set(calc)|set(rep)):
        c = calc.get(code,0.0); r = rep.get(code,0.0); out.append((code,c,r,c-r))
    return out

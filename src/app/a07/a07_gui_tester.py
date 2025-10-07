# -*- coding: utf-8 -*-
"""
A07 Lønnsavstemming — komplett GUI (med DnD-board)
---------------------------------------------------
- Parser A07 JSON
- Leser GL (saldobalanse) fra CSV robust
- Oversikt / Ansatte / Koder / Rådata
- Kontrolloppstilling: Auto-forslag (regelbok/fallback) + LP-Optimalisering + drilldown
- Innstillinger: Laste/Redigere regelbok + lagre/lese overstyringer (JSON)
- Ny fane: "Board (DnD)" for interaktiv mapping via drag & drop

Bygger videre på eksisterende prosjektstruktur (regelbok/fallback/LP).
"""

from __future__ import annotations

import csv, io, json, os, re, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict, Counter
from typing import Any, Dict, List, Tuple, Iterable, Optional, Set

# ---------- DnD-brett ----------
# Prøv relativ import hvis vi kjøres som del av pakke; ellers fall tilbake til absolutt import.
try:
    from .a07_board import A07Board  # type: ignore[attr-defined]
except Exception:
    from a07_board import A07Board


# ---------- Regelbok / fallback / LP ----------
try:
    from a07_rulebook import load_rulebook, suggest_with_rulebook
    HAVE_RULEBOOK = True
except Exception:
    load_rulebook = None     # type: ignore
    suggest_with_rulebook = None  # type: ignore
    HAVE_RULEBOOK = False

try:
    from matcher_fallback import suggest_mapping_for_accounts as fallback_suggest
except Exception:
    def fallback_suggest(*_a, **_k): return {}

try:
    from a07_optimize import generate_candidates_for_lp, solve_global_assignment_lp
    HAVE_LP = True
except Exception:
    HAVE_LP = False
    def generate_candidates_for_lp(*_a, **_k): return {}
    def solve_global_assignment_lp(*_a, **_k): raise RuntimeError("PuLP/LP ikke tilgjengelig")

# --------------------------- Utils ---------------------------

def _to_float(x: Any) -> float:
    if x is None: return 0.0
    if isinstance(x,(int,float)): return float(x)
    s = str(x).strip()
    if s == "": return 0.0
    s = s.replace("\xa0"," ").replace("−","-").replace("–","-").replace("—","-")
    s = re.sub(r"(?i)\b(nok|kr)\b\.?", "", s).strip()
    neg = s.startswith("(") and s.endswith(")")
    if neg: s = s[1:-1].strip()
    s = s.replace(" ","").replace("'","")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".","").replace(",",".")
        else:                            s = s.replace(",","")
    elif "," in s: s = s.replace(",",".")
    try: v = float(s)
    except Exception:
        s2 = re.sub(r"[^0-9\.\-]", "", s)
        if s2 in {"","-","."}: return 0.0
        try: v = float(s2)
        except Exception: return 0.0
    return -v if neg else v

def fmt_amount(x: float) -> str:
    try: return f"{x:,.2f}".replace(",", " ").replace(".", ",")
    except Exception: return str(x)

# --------------------------- A07 ---------------------------

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
                                               fordel=fordel, beloep=beloep, antall=antall,
                                               trekkpliktig=bool(inc.get("inngaarIGrunnlagForTrekk", False)),
                                               aga=bool(inc.get("utloeserArbeidsgiveravgift", False)),
                                               opptj_start=inc.get("startdatoOpptjeningsperiode"),
                                               opptj_slutt=inc.get("sluttdatoOpptjeningsperiode")))
                        except Exception as e:
                            errs.append(f"Feil ved parsing: {e}")
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
    out: Dict[str,float] = defaultdict(float)
    for r in rows: out[str(r["kode"])] += float(r["beloep"])
    return dict(out)

def summarize_by_employee(rows: Iterable[A07Row]) -> Dict[str, Dict[str, Any]]:
    idx: Dict[str, Dict[str, Any]] = {}
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

# --------------------------- GL CSV ---------------------------

def _read_text_guess(path: str) -> tuple[str,str]:
    encs = ["utf-8-sig","utf-16","utf-16le","utf-16be","cp1252","latin-1","utf-8"]
    for enc in encs:
        try:
            with open(path,"r",encoding=enc,errors="strict") as f:
                return f.read(), enc
        except UnicodeDecodeError: continue
        except Exception: continue
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

# --------------------------- Tk table ---------------------------

class Table(ttk.Treeview):
    def __init__(self, master, columns: List[Tuple[str,str]], **kwargs):
        ids = [c for c,_ in columns]
        super().__init__(master, columns=ids, show="headings", selectmode="extended", **kwargs)
        self._cols = columns
        for cid, header in columns:
            self.heading(cid, text=header, command=lambda c=cid: self._sort_by(c, False))
            self.column(cid, width=120, anchor=tk.W, stretch=True)
        self.tag_configure("ok", background="#e8f5e9")
        self.tag_configure("warn", background="#fff8e1")
        self.tag_configure("bad", background="#ffebee")
        self.tag_configure("muted", foreground="#7f7f7f")

        self._data_cache: List[List[Any]] = []
        self._column_formats: Dict[str,Any] = {}

        yscroll = ttk.Scrollbar(master, orient="vertical", command=self.yview)
        self.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def set_column_format(self, column_id: str, fmt_func): self._column_formats[column_id] = fmt_func
    def clear(self): self.delete(*self.get_children()); self._data_cache.clear()
    def insert_rows(self, rows: Iterable[Dict[str,Any]]):
        self.clear()
        for r in rows:
            values = []
            for cid,_ in self._cols:
                v = r.get(cid,"")
                if cid in self._column_formats: v = self._column_formats[cid](v)
                values.append(v)
            tags = r.get("_tags", [])
            self.insert("", tk.END, values=values, tags=tags)
            self._data_cache.append(values)

    def export_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([h for _,h in self._cols]); w.writerows(self._data_cache)

    def _sort_by(self, col: str, desc: bool):
        data = [(self.set(k,col),k) for k in self.get_children("")]
        def to_num(s):
            ss = str(s).strip().replace(" ","").replace(".", "").replace(",", ".")
            try: return float(ss)
            except Exception: return s
        data.sort(key=lambda t: to_num(t[0]), reverse=desc)
        for i,(_,k) in enumerate(data): self.move(k,"",i)
        self.heading(col, command=lambda c=col: self._sort_by(c, not desc))

# --------------------------- GUI ---------------------------

class A07App(tk.Tk):
    # ---------- Preferanser ----------
    def _prefs_file(self) -> str:
        return os.path.join(os.path.expanduser("~"), ".a07_prefs.json")
    def _load_prefs(self) -> dict:
        try:
            with open(self._prefs_file(), "r", encoding="utf-8") as f: return json.load(f)
        except Exception: return {}
    def _save_prefs(self, **updates):
        prefs = self._load_prefs(); prefs.update(updates)
        try:
            with open(self._prefs_file(), "w", encoding="utf-8") as f: json.dump(prefs, f, ensure_ascii=False, indent=2)
        except Exception: pass

    def __init__(self):
        super().__init__()
        self.title("A07 Lønnsavstemming — GUI")
        self.geometry("1320x820"); self.minsize(1080, 720)

        self.json_root: Dict[str,Any] = {}
        self.rows: List[A07Row] = []; self.errors: List[str] = []
        self.gl_accounts: List[Dict[str,Any]] = []; self.gl_meta: Dict[str,Any] = {}
        self.acc_to_code: Dict[str,str] = {}; self.auto_suggestions: Dict[str,Dict[str,Any]] = {}

        self.gl_basis = tk.StringVar(value="auto")
        self.diff_threshold = tk.DoubleVar(value=100.0)
        self.hide_zero = tk.BooleanVar(value=True)
        self.min_score = tk.DoubleVar(value=0.60)
        self.allow_splits = tk.BooleanVar(value=True)

        # UI-valg som vi også lagrer
        self.only_diff = tk.BooleanVar(value=False)
        self.only_unmapped = tk.BooleanVar(value=False)
        self.compact_view = tk.BooleanVar(value=True)

        # LP
        self.use_lp_assignment = False
        self.lp_assignment: Dict[str, Dict[str,float]] = {}
        self.lp_fixed: Dict[str,float] = {}
        self.lp_amounts: Dict[str,float] = {}

        # Regelbok
        self.rulebook: Optional[Dict[str,Any]] = None
        self.rulebook_overrides: Dict[str,Any] = {}
        self.rulebook_source = ""

        self._build_ui()

        # Autoload regelbok + UI-innstillinger
        try:
            prefs = self._load_prefs()
            _p = prefs.get("rulebook_path")
            if _p and load_rulebook is not None and os.path.exists(_p):
                self.rulebook = load_rulebook(_p); self.rulebook_source = _p
                self._refresh_settings_tables()
                self.status.configure(text=f"Regelbok lastet automatisk: {os.path.basename(_p)}")
            self.compact_view.set(bool(prefs.get("compact_view", True)))
            self.only_diff.set(bool(prefs.get("only_diff", False)))
            self.only_unmapped.set(bool(prefs.get("only_unmapped", False)))
        except Exception:
            pass

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- UI ----------

    def _build_ui(self):
        bar = ttk.Frame(self); bar.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        ttk.Button(bar, text="Åpne A07 JSON…", command=self.on_open_json).pack(side=tk.LEFT)
        ttk.Button(bar, text="Valider mot oppsummering", command=self.on_validate).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(bar, text="Eksporter tabell → CSV", command=self.on_export_current_table).pack(side=tk.LEFT, padx=(8,0))
        self.status = ttk.Label(bar, text="Ingen fil lastet.", anchor="w"); self.status.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        self.nb = ttk.Notebook(self); self.nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        tab_overview = ttk.Frame(self.nb); self.nb.add(tab_overview, text="Oversikt"); self._build_overview(tab_overview)
        tab_emp = ttk.Frame(self.nb); self.nb.add(tab_emp, text="Ansatte"); self._build_employees(tab_emp)
        tab_codes = ttk.Frame(self.nb); self.nb.add(tab_codes, text="Koder"); self._build_codes(tab_codes)
        tab_raw = ttk.Frame(self.nb); self.nb.add(tab_raw, text="Rådata"); self._build_raw(tab_raw)
        tab_ctrl = ttk.Frame(self.nb); self.nb.add(tab_ctrl, text="Kontrolloppstilling"); self._build_control(tab_ctrl)
        tab_settings = ttk.Frame(self.nb); self.nb.add(tab_settings, text="Innstillinger"); self._build_settings(tab_settings)

    # ----- Oversikt -----
    def _build_overview(self, root: ttk.Frame):
        top = ttk.Frame(root); top.pack(side=tk.TOP, fill=tk.X)
        self.ov_file = ttk.Label(top, text="Fil: –"); self.ov_file.pack(side=tk.TOP, anchor="w", pady=(6,0))
        grid = ttk.Frame(root); grid.pack(side=tk.TOP, fill=tk.X, pady=10)
        self.ov_labels = {
            "clients": ttk.Label(grid, text="Virksomheter: –"),
            "employees": ttk.Label(grid, text="Ansatte: –"),
            "rows": ttk.Label(grid, text="Antall inntektslinjer: –"),
            "sum": ttk.Label(grid, text="Total beløp (alle koder): –"),
        }
        for i,k in enumerate(["clients","employees","rows","sum"]): self.ov_labels[k].grid(row=0, column=i, sticky="w", padx=10)
        ttk.Separator(root, orient="horizontal").pack(side=tk.TOP, fill=tk.X, pady=4)
        ttk.Label(root, text="Summer pr lønnskode").pack(side=tk.TOP, anchor="w")
        frame_tbl = ttk.Frame(root); frame_tbl.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_codes_overview = Table(frame_tbl, [("kode","Kode / beskrivelse"), ("sum","Sum (NOK)")])
        self.tbl_codes_overview.set_column_format("sum", fmt_amount)

    # ----- Ansatte -----
    def _build_employees(self, root: ttk.Frame):
        top = ttk.Frame(root); top.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(top, text="Søk (navn/fnr):").pack(side=tk.LEFT)
        self.emp_filter_var = tk.StringVar(); ttk.Entry(top, textvariable=self.emp_filter_var, width=40).pack(side=tk.LEFT, padx=6)
        self.emp_filter_var.trace_add("write", lambda *_: self._refresh_employees_table())
        frame_tbl = ttk.Frame(root); frame_tbl.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_emp = Table(frame_tbl, [("fnr","Fødselsnr"),("navn","Navn"),("antall_poster","Ant. poster"),("sum","Sum (NOK)")])
        self.tbl_emp.set_column_format("sum", fmt_amount); self.tbl_emp.bind("<Double-1>", self._on_emp_double_click)

    def _on_emp_double_click(self, _=None):
        sel = self.tbl_emp.focus()
        if not sel: return
        fnr = self.tbl_emp.item(sel,"values")[0]
        self._open_emp_detail_window(fnr)

    def _open_emp_detail_window(self, fnr: str):
        win = tk.Toplevel(self); win.title(f"Detaljer — ansatt {fnr}"); win.geometry("900x500")
        data = [r for r in self.rows if str(r["fnr"]) == str(fnr)]
        ttk.Label(win, text=f"Navn: {data[0]['navn'] if data else ''}  •  Antall poster: {len(data)}  •  Sum: {fmt_amount(sum(float(r['beloep']) for r in data))}").pack(side=tk.TOP, anchor="w", padx=8, pady=8)
        frm = ttk.Frame(win); frm.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        tbl = Table(frm, [("kode","Kode"),("fordel","Fordel"),("beloep","Beløp"),("antall","Antall"),("trekkpliktig","Trekkpl."),("aga","AGA"),("opptj_start","Opptj. start"),("opptj_slutt","Opptj. slutt"),("orgnr","Orgnr")])
        tbl.set_column_format("beloep", fmt_amount); tbl.insert_rows(data)

    def _refresh_employees_table(self):
        idx = summarize_by_employee(self.rows)
        q = getattr(self, "emp_filter_var", tk.StringVar(value="")).get().strip().lower()
        rows = []
        for fnr, d in idx.items():
            navn = str(d["navn"])
            if q and q not in fnr.lower() and q not in navn.lower(): continue
            rows.append({"fnr": fnr, "navn": navn, "antall_poster": d["antall_poster"], "sum": d["sum"]})
        rows.sort(key=lambda r: (-float(r["sum"]), r["navn"]))
        self.tbl_emp.insert_rows(rows)

    # ----- Koder -----
    def _build_codes(self, root: ttk.Frame):
        top = ttk.Frame(root); top.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(top, text="Søk (kode):").pack(side=tk.LEFT)
        self.code_filter_var = tk.StringVar(); ttk.Entry(top, textvariable=self.code_filter_var, width=40).pack(side=tk.LEFT, padx=6)
        self.code_filter_var.trace_add("write", lambda *_: self._refresh_codes_table())
        frame_tbl = ttk.Frame(root); frame_tbl.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_codes = Table(frame_tbl, [("kode","Kode / beskrivelse"),("antall_poster","Ant. poster"),("sum","Sum (NOK)")])
        self.tbl_codes.set_column_format("sum", fmt_amount); self.tbl_codes.bind("<Double-1>", self._on_code_detail)

    def _on_code_detail(self, _evt=None):
        sel = self.tbl_codes.focus()
        if not sel: return
        kode = self.tbl_codes.item(sel, "values")[0]
        self._open_code_detail_window(kode)

    def _open_code_detail_window(self, kode: str):
        win = tk.Toplevel(self); win.title(f"Detaljer — kode {kode}"); win.geometry("980x520")
        data = [r for r in self.rows if str(r["kode"]) == str(kode)]
        per_emp = summarize_by_employee(data)
        rows_idx = [{"fnr": fnr, "navn": v["navn"], "antall_poster": v["antall_poster"], "sum": v["sum"]} for fnr, v in per_emp.items()]
        rows_idx.sort(key=lambda r: (-float(r["sum"]), r["navn"]))
        ttk.Label(win, text=f"Antall poster: {len(data)}  •  Sum: {fmt_amount(sum(float(r['beloep']) for r in data))}").pack(side=tk.TOP, anchor="w", padx=8, pady=8)
        nb = ttk.Notebook(win); nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        tab1 = ttk.Frame(nb); nb.add(tab1, text="Per ansatt")
        f1 = ttk.Frame(tab1); f1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        t1 = Table(f1, [("fnr","Fødselsnr"),("navn","Navn"),("antall_poster","Ant. poster"),("sum","Sum (NOK)")])
        t1.set_column_format("sum", fmt_amount); t1.insert_rows(rows_idx)
        tab2 = ttk.Frame(nb); nb.add(tab2, text="Rå linjer")
        f2 = ttk.Frame(tab2); f2.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        t2 = Table(f2, [("fnr","Fødselsnr"),("navn","Navn"),("fordel","Fordel"),("beloep","Beløp"),("antall","Antall"),("orgnr","Orgnr")])
        t2.set_column_format("beloep", fmt_amount)
        t2.insert_rows([{ "fnr": r["fnr"], "navn": r["navn"], "fordel": r["fordel"], "beloep": r["beloep"], "antall": r["antall"], "orgnr": r["orgnr"]} for r in data])

    def _refresh_codes_table(self):
        sums = summarize_by_code(self.rows); counts = Counter([r["kode"] for r in self.rows])
        q = getattr(self, "code_filter_var", tk.StringVar(value="")).get().strip().lower()
        rows = []
        for kode, s in sums.items():
            if q and q not in str(kode).lower(): continue
            rows.append({"kode": kode, "antall_poster": counts.get(kode, 0), "sum": s})
        rows.sort(key=lambda r: (-float(r["sum"]), r["kode"]))
        self.tbl_codes.insert_rows(rows)

    # ----- Rådata -----
    def _build_raw(self, root: ttk.Frame):
        f = ttk.Frame(root); f.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_raw = Table(f, [("orgnr","Orgnr"),("fnr","Fødselsnr"),("navn","Navn"),("kode","Kode"),("fordel","Fordel"),("beloep","Beløp"),("antall","Antall"),("trekkpliktig","Trekkpl."),("aga","AGA"),("opptj_start","Opptj. start"),("opptj_slutt","Opptj. slutt")])
        self.tbl_raw.set_column_format("beloep", fmt_amount)

    # ----- Kontrolloppstilling -----
    def _build_control(self, root: ttk.Frame):
        bar = ttk.Frame(root); bar.pack(side=tk.TOP, fill=tk.X, pady=(8,6))
        ttk.Button(bar, text="Last inn saldobalanse (CSV…)", command=self.on_load_gl).pack(side=tk.LEFT)
        ttk.Button(bar, text="Auto-forslag mapping", command=self.on_auto_map).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(bar, text="Sett kode…", command=self.on_set_code).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(bar, text="Fjern mapping", command=self.on_clear_code).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(bar, text="Nullstill mapping", command=self.on_reset_mapping).pack(side=tk.LEFT, padx=(8,0))
        self.ctrl_status = ttk.Label(bar, text="Ingen saldobalanse lastet.", anchor="w"); self.ctrl_status.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        dash = ttk.Frame(root); dash.pack(side=tk.TOP, fill=tk.X, padx=8)
        self.lab_a07 = ttk.Label(dash, text="A07: –"); self.lab_a07.pack(side=tk.LEFT, padx=8)
        self.lab_gl  = ttk.Label(dash, text="GL (mappet): –"); self.lab_gl.pack(side=tk.LEFT, padx=8)
        self.lab_diff= ttk.Label(dash, text="Diff: –"); self.lab_diff.pack(side=tk.LEFT, padx=8)
        self.lab_unmapped = ttk.Label(dash, text="Uten mapping: –"); self.lab_unmapped.pack(side=tk.LEFT, padx=8)
        self.lab_code_gap = ttk.Label(dash, text="Koder uten GL: –"); self.lab_code_gap.pack(side=tk.LEFT, padx=8)

        ctrl = ttk.Frame(root); ctrl.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(4,0))
        ttk.Label(ctrl, text="Regnskapsgrunnlag:").pack(side=tk.LEFT)
        for val, txt in [("auto","Auto"),("ub","UB (utgående saldo)"),("endring","Endring (Debet−Kredit/UB−IB)"),("belop","Beløp (valgt kol.)")]:
            ttk.Radiobutton(ctrl, text=txt, variable=self.gl_basis, value=val, command=self.refresh_control_tables).pack(side=tk.LEFT, padx=(6,0))

        ttk.Label(ctrl, text="  •  Avviks-terskel:").pack(side=tk.LEFT, padx=(12,0))
        ttk.Spinbox(ctrl, from_=0, to=10_000_000, increment=50, width=8, textvariable=self.diff_threshold, command=self.refresh_control_tables).pack(side=tk.LEFT)
        ttk.Checkbutton(ctrl, text="Skjul konti med 0", variable=self.hide_zero, command=self.refresh_control_tables).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(ctrl, text="Vis kun uten mapping", variable=self.only_unmapped, command=self.refresh_control_tables).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(ctrl, text="Kompakt visning", variable=self.compact_view, command=self.refresh_control_tables).pack(side=tk.LEFT, padx=8)

        ttk.Label(ctrl, text="  •  Min‑score:").pack(side=tk.LEFT, padx=(12,0))
        ttk.Spinbox(ctrl, from_=0.00, to=1.00, increment=0.05, width=5, textvariable=self.min_score, command=self.refresh_control_tables).pack(side=tk.LEFT)
        ttk.Checkbutton(ctrl, text="Tillat splitting (LP)", variable=self.allow_splits).pack(side=tk.LEFT, padx=12)
        ttk.Button(ctrl, text="Optimaliser beløp (LP)", command=self.on_optimize_lp).pack(side=tk.LEFT, padx=8)

        self.ctrl_nb = ttk.Notebook(root); self.ctrl_nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(6,8))
        tabA = ttk.Frame(self.ctrl_nb); self.ctrl_nb.add(tabA, text="Konti & forslag")
        topA = ttk.Frame(tabA); topA.pack(side=tk.TOP, fill=tk.X)
        self.gl_search_var = tk.StringVar(); ttk.Label(topA, text="Søk konto/tekst:").pack(side=tk.LEFT)
        ttk.Entry(topA, textvariable=self.gl_search_var, width=30).pack(side=tk.LEFT, padx=6)
        self.gl_search_var.trace_add("write", lambda *_: self.refresh_control_tables())

        # venstre (tabell) + høyre (detalj)
        frame_tblA = ttk.Frame(tabA); frame_tblA.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        leftA = ttk.Frame(frame_tblA); leftA.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rightA = ttk.Frame(frame_tblA, relief=tk.GROOVE, borderwidth=1); rightA.pack(side=tk.LEFT, fill=tk.Y, padx=(8,0))
        self.tbl_gl = Table(leftA, [("konto","Konto"),("navn","Kontonavn"),("ib","IB"),("debet","Debet"),("kredit","Kredit"),("endring","Endring"),("ub","UB"),("basis","Basis"),("foreslatt","Foreslått kode"),("score","Score"),("begrunnelse","Begrunnelse")])
        for c in ("ib","debet","kredit","endring","ub"): self.tbl_gl.set_column_format(c, fmt_amount)
        self.tbl_gl.bind("<<TreeviewSelect>>", self._on_gl_row_selected)

        # detaljpanel
        ttk.Label(rightA, text="Detaljer / hurtigvalg", font=("TkDefaultFont", 10, "bold")).pack(side=tk.TOP, anchor="w", padx=8, pady=(6,2))
        self.det_account = ttk.Label(rightA, text="Konto: –"); self.det_account.pack(side=tk.TOP, anchor="w", padx=8)
        self.det_amount  = ttk.Label(rightA, text="Beløp (basis): –"); self.det_amount.pack(side=tk.TOP, anchor="w", padx=8, pady=(0,6))
        ttk.Label(rightA, text="Valgt forslag:").pack(side=tk.TOP, anchor="w", padx=8)
        self.det_best = ttk.Label(rightA, text="–", foreground="#2e7d32"); self.det_best.pack(side=tk.TOP, anchor="w", padx=8, pady=(0,6))
        ttk.Separator(rightA, orient="horizontal").pack(side=tk.TOP, fill=tk.X, padx=6, pady=4)
        ttk.Label(rightA, text="Alternativer (topp‑5):").pack(side=tk.TOP, anchor="w", padx=8)
        self.det_alt_list = tk.Listbox(rightA, height=12, exportselection=False); self.det_alt_list.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=8, pady=(2,6))
        btns = ttk.Frame(rightA); btns.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(0,8))
        ttk.Button(btns, text="Bruk valgt", command=self._apply_selected_alt).pack(side=tk.LEFT)
        ttk.Button(btns, text="Fjern mapping", command=self.on_clear_code).pack(side=tk.LEFT, padx=6)

        tabB = ttk.Frame(self.ctrl_nb); self.ctrl_nb.add(tabB, text="Avstemming pr kode")
        topB = ttk.Frame(tabB); topB.pack(side=tk.TOP, fill=tk.X)
        ttk.Checkbutton(topB, text="Vis bare avvik", variable=self.only_diff, command=self.refresh_control_tables).pack(side=tk.LEFT, padx=8, pady=(2,0))
        frame_tblB = ttk.Frame(tabB); frame_tblB.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_ctrl_codes = Table(frame_tblB, [("kode","A07-kode"),("a07","A07 sum"),("gl","Regnskap (mappet)"),("diff","Diff (A07−GL)"),("ant_konti","# konti mappet")])
        for c in ("a07","gl","diff"): self.tbl_ctrl_codes.set_column_format(c, fmt_amount)
        self.tbl_ctrl_codes.bind("<Double-1>", self._on_ctrl_code_drill)

        # --- NY fane: Board (DnD) ---
        tabDND = ttk.Frame(self.ctrl_nb)
        self.ctrl_nb.add(tabDND, text="Board (DnD)")

        def _board_get_amount(acc: Dict[str,Any]) -> float:
            return self._gl_amount(acc)[0]

        def _on_drop(accno: str, code: str):
            self.acc_to_code[str(accno)] = str(code)
            self.use_lp_assignment = False
            self.refresh_control_tables()

        def _req_suggestions():
            self.on_auto_map()

        self.board = AssignmentBoard(
            tabDND,
            get_amount_fn=_board_get_amount,
            on_drop=_on_drop,
            request_suggestions=_req_suggestions
        )
        self.board.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)

    # ----- Innstillinger -----
    def _build_settings(self, root: ttk.Frame):
        top = ttk.Frame(root); top.pack(side=tk.TOP, fill=tk.X, pady=(8,4))
        ttk.Button(top, text="Last regelbok (Excel)…", command=self.on_load_rulebook_excel).pack(side=tk.LEFT)
        ttk.Button(top, text="Last regelbok (CSV‑mappe)…", command=self.on_load_rulebook_csvdir).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(top, text="Bruk regelbok → Auto‑forslag", command=self.on_auto_map).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(top, text="Rediger valgt kode", command=self.on_edit_rule).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(top, text="Lagre endringer (JSON)…", command=self.on_save_rulebook_overrides).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(top, text="Last endringer (JSON)…", command=self.on_load_rulebook_overrides).pack(side=tk.LEFT, padx=(8,0))
        self.rulebook_info = ttk.Label(top, text=self._rulebook_status_text(), anchor="w"); self.rulebook_info.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        self.set_nb = ttk.Notebook(root); self.set_nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(4,8))
        tab1 = ttk.Frame(self.set_nb); self.set_nb.add(tab1, text="A07‑koder")
        f1 = ttk.Frame(tab1); f1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_rule_codes = Table(f1, [
            ("a07_code","A07‑kode"),("category","Kategori"),("basis","Basis"),
            ("allowed","Tillatte kontoområder"),("keywords","Nøkkelord"),
            ("boost","Boost‑konti"),("expected","Forventet tegn"),("special","Special‑add")
        ])
        tab2 = ttk.Frame(self.set_nb); self.set_nb.add(tab2, text="Aliases")
        f2 = ttk.Frame(tab2); f2.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tbl_rule_alias = Table(f2, [("canonical","Kanonisk"),("synonyms","Synonymer")])
        self._refresh_settings_tables()

    # ----- Preferanser/avslutning -----
    def _on_close(self):
        try:
            self._save_prefs(
                rulebook_path=self.rulebook_source,
                compact_view=bool(self.compact_view.get()),
                only_diff=bool(self.only_diff.get()),
                only_unmapped=bool(self.only_unmapped.get())
            )
        except Exception:
            pass
        self.destroy()

    # ---------- Regelbok-presentasjon ----------
    def _rulebook_status_text(self) -> str:
        if not self.rulebook: return "Regelbok: (ingen lastet)"
        c = len(self.rulebook.get("codes", {})); a = len(self.rulebook.get("aliases", {})); src = self.rulebook.get("source","")
        return f"Regelbok: {c} koder, {a} alias‑grupper  •  Kilde: {src or '(auto)'}"

    def _format_allowed(self, intervals: List[Tuple[int,int]]) -> str:
        if not intervals: return ""
        parts = [f"{lo}" if lo==hi else f"{lo}-{hi}" for lo,hi in intervals]
        return " | ".join(parts)

    def _refresh_settings_tables(self):
        rows = []
        if self.rulebook:
            for code, rule in self.rulebook.get("codes", {}).items():
                es = int(rule.get("expected_sign", 0))
                rows.append({
                    "a07_code": code, "category": rule.get("category",""), "basis": rule.get("basis",""),
                    "allowed": self._format_allowed(rule.get("allowed",[])),
                    "keywords": ", ".join(sorted(rule.get("keywords", []))),
                    "boost": ", ".join(sorted(rule.get("boost_accounts", []))),
                    "expected": ("+" if es==1 else ("-" if es==-1 else "")),
                    "special": json.dumps(rule.get("special_add", []), ensure_ascii=False),
                })
        rows.sort(key=lambda r: r["a07_code"]); self.tbl_rule_codes.insert_rows(rows)
        rows2 = []
        if self.rulebook:
            for can, syns in sorted(self.rulebook.get("aliases", {}).items()):
                rows2.append({"canonical": can, "synonyms": ", ".join(sorted(syns))})
        self.tbl_rule_alias.insert_rows(rows2)
        self.rulebook_info.config(text=self._rulebook_status_text())

    # ---------- Event handlers ----------

    def on_open_json(self):
        path = filedialog.askopenfilename(title="Velg A07 JSON-fil", filetypes=[("JSON","*.json"),("Alle filer","*.*")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f: data = json.load(f)
        except Exception as e:
            messagebox.showerror("Feil ved lesing", f"Kunne ikke lese JSON: {e}"); return
        parser = A07Parser()
        rows, errors = parser.parse(data)
        self.json_root = data; self.rows = rows; self.errors = errors; self._file_name = os.path.basename(path)
        self.status.configure(text=f"Lest {len(rows)} rader. Feil: {len(errors)}")
        if errors: messagebox.showwarning("Parsing", "\n".join(errors[:12]) + ("\n… (flere)" if len(errors)>12 else ""))
        self._refresh_all_tabs()

    def on_validate(self):
        if not self.rows or not self.json_root:
            messagebox.showinfo("Ingen data", "Last inn en A07 JSON først."); return
        checks = validate_against_summary(self.rows, self.json_root)
        if not checks: messagebox.showinfo("Ingen oppsummering", "Fant ikke 'oppsummerteVirksomheter' i JSON."); return
        win = tk.Toplevel(self); win.title("Validering mot oppsummering"); win.geometry("720x480")
        ttk.Label(win, text="Sammenligning av summer per kode:").pack(side=tk.TOP, anchor="w", padx=8, pady=8)
        f = ttk.Frame(win); f.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        tbl = Table(f, [("kode","Kode"),("detalj","Detaljsum"),("oppsummert","Oppsummert"),("diff","Diff")])
        for c in ("detalj","oppsummert","diff"): tbl.set_column_format(c, fmt_amount)
        rows = [{"kode":c,"detalj":d,"oppsummert":r,"diff":diff} for (c,d,r,diff) in checks]; rows.sort(key=lambda r: -abs(float(r["diff"])))
        tbl.insert_rows(rows); ttk.Label(win, text=f"Total diff: {fmt_amount(sum(float(r['diff']) for r in rows))}").pack(side=tk.TOP, anchor="w", padx=8, pady=(0,8))

    def on_export_current_table(self):
        tab = self.nb.select(); widget = self.nametowidget(tab); target: Optional[Table] = None
        def _find_table(w):
            nonlocal target
            if isinstance(w, Table): target = w; return
            for ch in w.winfo_children(): _find_table(ch)
        _find_table(widget)
        if not target: messagebox.showinfo("Ingen tabell", "Aktiv fane inneholder ingen tabell å eksportere."); return
        path = filedialog.asksaveasfilename(title="Lagre CSV", defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path: return
        try: target.export_csv(path); messagebox.showinfo("Eksport", f"Skrev fil: {os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Feil ved eksport", str(e))

    def on_load_gl(self):
        path = filedialog.askopenfilename(title="Velg saldobalanse (CSV)", filetypes=[("CSV","*.csv"),("Alle filer","*.*")])
        if not path: return
        try:
            rows, meta = read_gl_csv(path)
        except Exception as e:
            messagebox.showerror("Feil ved lesing", f"Kunne ikke lese CSV: {e}"); return
        self.gl_accounts = rows; self.gl_meta = meta
        self.acc_to_code.clear(); self.auto_suggestions.clear()
        self.use_lp_assignment = False; self.lp_assignment.clear(); self.lp_fixed.clear(); self.lp_amounts.clear()
        self.ctrl_status.configure(text=f"Saldobalanse: {len(rows)} konti.  Kolonner: IB={'ja' if meta.get('ib') else 'nei'}, D={'ja' if meta.get('debet') else 'nei'}, K={'ja' if meta.get('kredit') else 'nei'}, Endr={'ja' if meta.get('endring') else 'nei'}, UB={'ja' if meta.get('ub') else 'nei'}  • Enc: {meta.get('encoding')} • Sep: '{meta.get('delimiter')}'")
        self.refresh_control_tables()

    def on_auto_map(self):
        if not self.gl_accounts: messagebox.showinfo("Mangler saldobalanse", "Last inn saldobalanse (CSV) først."); return
        if not self.rows: messagebox.showinfo("Mangler A07", "Last inn A07 JSON først."); return
        a07_sums = summarize_by_code(self.rows)
        if self.rulebook and suggest_with_rulebook is not None:
            suggestions = suggest_with_rulebook(self.gl_accounts, a07_sums, self.rulebook, min_score=float(self.min_score.get()))
        else:
            suggestions = fallback_suggest(self.gl_accounts, a07_sums, min_score=float(self.min_score.get()))
        self.auto_suggestions = suggestions
        for acc in self.gl_accounts:
            accno = acc["konto"]
            if accno in suggestions: self.acc_to_code[accno] = suggestions[accno]["kode"]
        self.use_lp_assignment = False
        self.refresh_control_tables()

    def on_set_code(self):
        sel = self.tbl_gl.selection()
        if not sel: messagebox.showinfo("Velg konto", "Marker én eller flere konti i tabellen først."); return
        codes = sorted(list(summarize_by_code(self.rows).keys()))
        win = tk.Toplevel(self); win.title("Sett kode for valgt(e) konto(er)"); win.geometry("460x120")
        ttk.Label(win, text=f"Velg A07-kode:").pack(side=tk.TOP, anchor="w", padx=10, pady=(10,6))
        var = tk.StringVar(value=""); cb = ttk.Combobox(win, textvariable=var, values=[""]+codes, state="readonly"); cb.pack(side=tk.TOP, fill=tk.X, padx=10)
        box = ttk.Frame(win); box.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        def ok():
            val = var.get().strip()
            for item in sel:
                accno = self.tbl_gl.item(item, "values")[0]
                if val: self.acc_to_code[accno] = val
                else: self.acc_to_code.pop(accno, None)
            self.use_lp_assignment = False
            win.destroy(); self.refresh_control_tables()
        ttk.Button(box, text="OK", command=ok).pack(side=tk.RIGHT); ttk.Button(box, text="Avbryt", command=win.destroy).pack(side=tk.RIGHT, padx=6)

    def on_clear_code(self):
        sel = self.tbl_gl.selection()
        if not sel: messagebox.showinfo("Velg konto", "Marker konti du vil fjerne mapping for."); return
        for item in sel:
            accno = self.tbl_gl.item(item, "values")[0]
            self.acc_to_code.pop(accno, None)
        self.use_lp_assignment = False
        self.refresh_control_tables()

    def on_reset_mapping(self):
        if messagebox.askyesno("Nullstill mapping", "Fjern all mapping (manuell + auto + LP)?"):
            self.acc_to_code.clear(); self.auto_suggestions.clear()
            self.use_lp_assignment = False; self.lp_assignment.clear(); self.lp_fixed.clear(); self.lp_amounts.clear()
            self.refresh_control_tables()

    # ---------- LP ----------
    def on_optimize_lp(self):
        if not HAVE_LP:
            messagebox.showinfo("LP mangler", "PuLP/LP er ikke tilgjengelig i miljøet."); return
        if not self.gl_accounts: messagebox.showinfo("Mangler saldobalanse", "Last inn saldobalanse (CSV) først."); return
        if not self.rows: messagebox.showinfo("Mangler A07", "Last inn A07 JSON først."); return
        if not self.rulebook:
            messagebox.showinfo("Regelbok mangler", "Last regelbok under 'Innstillinger' først."); return

        amounts: Dict[str,float] = {}
        for acc in self.gl_accounts:
            val, _ = self._gl_amount(acc)
            amounts[str(acc["konto"])] = float(val)
        self.lp_amounts = amounts

        fixed = defaultdict(float); skip_edges = set()
        for code, rule in self.rulebook.get("codes", {}).items():
            for spec in rule.get("special_add", []):
                accno = str(spec.get("account","")); basis = str(spec.get("basis","endring")).lower(); weight = float(spec.get("weight", 1.0))
                for acc in self.gl_accounts:
                    if str(acc.get("konto")) == accno:
                        v = self._gl_amount_by_basis(acc, basis)
                        fixed[code] += weight * v
                        skip_edges.add((accno, code))
                        break
        self.lp_fixed = dict(fixed)

        a07_raw = summarize_by_code(self.rows)
        targets = {c: float(a07_raw.get(c,0.0) - fixed.get(c,0.0)) for c in set(a07_raw)|set(fixed)}

        cands = generate_candidates_for_lp(self.gl_accounts, a07_raw, self.rulebook,
                                           amounts_override=amounts,
                                           min_name=0.25, min_score=float(self.min_score.get())-0.10,
                                           top_k=3, skip_edges=skip_edges)
        if not cands:
            messagebox.showwarning("LP", "Fant ingen kandidater. Sjekk regelbok/Min‑score."); return

        try:
            assignment = solve_global_assignment_lp(
                amounts, cands, targets,
                allow_splits=bool(self.allow_splits.get()),
                lambda_score=0.25, lambda_sign=0.05
            )
        except Exception as e:
            messagebox.showerror("LP-feil", str(e)); return

        self.lp_assignment = assignment; self.use_lp_assignment = True
        self.acc_to_code.clear()
        for accno, parts in assignment.items():
            if parts:
                code = max(parts.items(), key=lambda kv: kv[1])[0]
                self.acc_to_code[accno] = code
        self.refresh_control_tables()
        messagebox.showinfo("Optimalisering", "LP‑løsning ferdig. Avstemmingstabellen er oppdatert.")

    # ---------- Beregninger/refresh ----------

    def _gl_amount(self, acc: Dict[str, Any]) -> Tuple[float, str]:
        mode = self.gl_basis.get()
        if mode == "ub": return float(acc.get("ub", 0.0)), "UB"
        if mode == "endring": return float(acc.get("endring", 0.0)), "Endring"
        if mode == "belop": return float(acc.get("belop", 0.0)), "Beløp"
        accno = str(acc.get("konto","")); digits = re.sub(r"\D+","",accno)
        if digits and digits.startswith("29"): return float(acc.get("ub", acc.get("belop",0.0))), "Auto:UB"
        return float(acc.get("endring", acc.get("belop",0.0))), "Auto:Endring"

    def _gl_amount_by_basis(self, acc: Dict[str,Any], basis: str) -> float:
        b = (basis or "endring").lower()
        if b == "ub": return float(acc.get("ub",0.0))
        if b == "belop": return float(acc.get("belop",0.0))
        return float(acc.get("endring", acc.get("belop",0.0)))

    def _set_compact_columns(self):
        compact = bool(self.compact_view.get())
        cols_hide = ["ib","debet","kredit","ub"]
        for c in cols_hide:
            try:
                self.tbl_gl.column(c, width=(1 if compact else 120), stretch=(not compact))
            except Exception:
                pass

    def _on_gl_row_selected(self, _evt=None):
        sel = self.tbl_gl.focus()
        if not sel:
            return
        vals = self.tbl_gl.item(sel, "values")
        if not vals:
            return
        accno = str(vals[0])
        self._refresh_gl_detail_panel(accno)

    def _apply_selected_alt(self):
        try:
            idx = self.det_alt_list.curselection()
            if not idx: return
            line = self.det_alt_list.get(idx[0])
            code = line.split()[0]
            if hasattr(self, "_detail_accno") and self._detail_accno and code:
                self.acc_to_code[self._detail_accno] = code
                self.use_lp_assignment = False
                self.refresh_control_tables()
        except Exception:
            pass

    def _refresh_gl_detail_panel(self, accno: str):
        self._detail_accno = accno
        acc = None
        for a in self.gl_accounts:
            if str(a.get("konto")) == str(accno):
                acc = a; break
        if not acc:
            self.det_account.config(text="Konto: –"); self.det_amount.config(text="Beløp (basis): –"); self.det_best.config(text="–")
            self.det_alt_list.delete(0, tk.END); return

        amt, lbl = self._gl_amount(acc)
        self.det_account.config(text=f"Konto {accno} — {acc.get('navn','')}")
        self.det_amount.config(text=f"{lbl}: {fmt_amount(amt)}")

        sugg = self.auto_suggestions.get(accno, {})
        chosen = self.acc_to_code.get(accno, sugg.get("kode",""))
        best_txt = chosen or "–"
        if sugg and chosen == sugg.get("kode",""):
            sc = sugg.get("score", "")
            best_txt += f"   ({sc:.3f})" if isinstance(sc,(float,int)) else ""
        self.det_best.config(text=best_txt)

        self.det_alt_list.delete(0, tk.END)
        try:
            if self.rulebook and HAVE_LP:
                a07_sums = summarize_by_code(self.rows)
                cands = generate_candidates_for_lp([acc], a07_sums, self.rulebook,
                                                   amounts_override={accno: amt},
                                                   min_name=0.20, min_score=max(0.30, float(self.min_score.get())-0.20),
                                                   top_k=5, skip_edges=None).get(accno, [])
                for (code, score, _a, reason) in cands:
                    self.det_alt_list.insert(tk.END, f"{code}    {score:.3f}  — {reason}")
        except Exception:
            pass

    def refresh_control_tables(self):
        # Konti & forslag
        q = getattr(self, "gl_search_var", tk.StringVar(value="")).get().strip().lower()
        rowsA = []
        for acc in self.gl_accounts:
            s, lbl = self._gl_amount(acc)
            if self.hide_zero.get() and abs(s) < 1e-9 and abs(float(acc.get("ub",0.0))) < 1e-9: continue
            if q and (q not in str(acc["konto"]).lower()) and (q not in str(acc.get("navn","")).lower()): continue
            sugg = self.auto_suggestions.get(acc["konto"], {})
            chosen = self.acc_to_code.get(acc["konto"], sugg.get("kode",""))
            score = sugg.get("score","") if chosen == sugg.get("kode","") else ""
            reason = sugg.get("reason","") if chosen == sugg.get("kode","") else ""
            if self.use_lp_assignment and self.lp_assignment.get(acc["konto"]):
                parts = self.lp_assignment[acc["konto"]]
                top = sorted(parts.items(), key=lambda kv: -kv[1])[:2]
                reason = (reason + " | " if reason else "") + "LP: " + ", ".join([f"{c} {p*100:.0f}%" for c,p in top])
            tags = ["ok"] if chosen else (["muted"] if abs(s) < 1e-9 else ["warn"])
            rowsA.append({
                "konto": acc["konto"], "navn": acc.get("navn",""), "ib": acc.get("ib",0.0), "debet": acc.get("debet",0.0), "kredit": acc.get("kredit",0.0),
                "endring": acc.get("endring",0.0), "ub": acc.get("ub",0.0), "basis": lbl, "foreslatt": chosen,
                "score": f"{score:.3f}" if isinstance(score,(int,float)) else "", "begrunnelse": reason, "_tags": tags
            })
        if self.only_unmapped.get():
            rowsA = [r for r in rowsA if not r.get("foreslatt")]
        rowsA.sort(key=lambda r: (r["foreslatt"] or "zzz", -abs(float(r["endring"])))); self.tbl_gl.insert_rows(rowsA)
        self._set_compact_columns()
        try:
            sel = self.tbl_gl.focus()
            if sel:
                accno = str(self.tbl_gl.item(sel, 'values')[0])
                self._refresh_gl_detail_panel(accno)
        except Exception:
            pass

        # Avstemming pr kode
        a07 = summarize_by_code(self.rows)
        gl_per_code: Dict[str,float] = defaultdict(float); code_to_accounts: Dict[str,int] = defaultdict(int)

        if self.use_lp_assignment and self.lp_assignment:
            for accno, parts in self.lp_assignment.items():
                amt = self.lp_amounts.get(accno, 0.0)
                for code, frac in parts.items():
                    gl_per_code[code] += amt * float(frac)
                    code_to_accounts[code] += 1
            for code, fx in self.lp_fixed.items():
                gl_per_code[code] += float(fx)
        else:
            for acc in self.gl_accounts:
                code = self.acc_to_code.get(acc["konto"])
                if not code: continue
                amt,_ = self._gl_amount(acc); gl_per_code[code] += amt; code_to_accounts[code] += 1
            if self.rulebook:
                for code, rule in self.rulebook.get("codes", {}).items():
                    for spec in rule.get("special_add", []):
                        accno = str(spec.get("account","")); basis = str(spec.get("basis","endring")).lower(); weight = float(spec.get("weight", 1.0))
                        if accno in self.acc_to_code: continue
                        for acc in self.gl_accounts:
                            if str(acc.get("konto")) == accno:
                                gl_per_code[code] += weight * self._gl_amount_by_basis(acc, basis)

        all_codes = set(a07)|set(gl_per_code)
        rowsB = []
        total_a07 = 0.0; total_gl = 0.0; thr = self.diff_threshold.get()
        for code in sorted(all_codes):
            a = a07.get(code,0.0); g = gl_per_code.get(code,0.0); d = a - g
            total_a07 += a; total_gl += g
            tag = "ok" if abs(d) <= thr else ("warn" if abs(d) <= 5*thr else "bad")
            rowsB.append({"kode": code, "a07": a, "gl": g, "diff": d, "ant_konti": code_to_accounts.get(code,0), "_tags":[tag]})
        if self.only_diff.get():
            rowsB = [r for r in rowsB if abs(float(r["diff"])) > thr]
        rowsB.sort(key=lambda r: -abs(float(r["diff"]))); self.tbl_ctrl_codes.insert_rows(rowsB)
        self.lab_a07.configure(text=f"A07: {fmt_amount(total_a07)}"); self.lab_gl.configure(text=f"GL (mappet): {fmt_amount(total_gl)}"); self.lab_diff.configure(text=f"Diff: {fmt_amount(total_a07-total_gl)}")
        unmapped = [acc for acc in self.gl_accounts if acc["konto"] not in self.acc_to_code]
        self.lab_unmapped.configure(text=f"Uten mapping: {len([a for a in unmapped if not(self.hide_zero.get() and abs(self._gl_amount(a)[0])<1e-9)])}")
        self.lab_code_gap.configure(text=f"Koder uten GL: {len([c for c in a07 if gl_per_code.get(c,0.0)==0.0])}")

        # --- Oppdater DnD‑brettet ---
        try:
            if hasattr(self, "board"):
                self.board.supply_data(
                    accounts=self.gl_accounts,
                    acc_to_code=self.acc_to_code,
                    suggestions=self.auto_suggestions,
                    a07_sums=a07,
                    diff_threshold=float(self.diff_threshold.get()),
                    only_unmapped=bool(self.only_unmapped.get()),
                )
        except Exception:
            pass

    def _refresh_overview(self):
        self.ov_file.configure(text=f"Fil: {getattr(self,'_file_name','–')}")
        uniq_clients = len(set(r["orgnr"] for r in self.rows)); uniq_emp = len(set(r["fnr"] for r in self.rows))
        total_rows = len(self.rows); total_amount = sum(float(r["beloep"]) for r in self.rows)
        self.ov_labels["clients"].configure(text=f"Virksomheter: {uniq_clients}")
        self.ov_labels["employees"].configure(text=f"Ansatte: {uniq_emp}")
        self.ov_labels["rows"].configure(text=f"Antall inntektslinjer: {total_rows}")
        self.ov_labels["sum"].configure(text=f"Total beløp (alle koder): {fmt_amount(total_amount)}")
        sums = summarize_by_code(self.rows); rows = [{"kode":k, "sum":v} for k,v in sums.items()]
        rows.sort(key=lambda r: (-float(r["sum"]), r["kode"])); self.tbl_codes_overview.insert_rows(rows)

    def _refresh_all_tabs(self):
        self._refresh_overview(); self._refresh_employees_table(); self._refresh_codes_table(); self.tbl_raw.insert_rows(self.rows); self.refresh_control_tables()

    def _on_ctrl_code_drill(self, _=None):
        sel = self.tbl_ctrl_codes.focus()
        if not sel: return
        kode = self.tbl_ctrl_codes.item(sel,"values")[0]
        a07_data = [r for r in self.rows if str(r["kode"]) == str(kode)]
        gl_data = []
        if self.use_lp_assignment and self.lp_assignment:
            for acc in self.gl_accounts:
                accno = str(acc["konto"])
                if kode in self.lp_assignment.get(accno, {}):
                    amt = self.lp_amounts.get(accno, 0.0) * float(self.lp_assignment[accno][kode])
                    lbl = self._gl_amount(acc)[1]
                    gl_data.append({"konto": acc["konto"], "navn": acc.get("navn",""), "belop": amt, "basis": f"{lbl}·LP"})
        else:
            for acc in self.gl_accounts:
                if self.acc_to_code.get(acc["konto"]) == kode:
                    amt,lbl = self._gl_amount(acc)
                    gl_data.append({"konto": acc["konto"], "navn": acc.get("navn",""), "belop": amt, "basis": lbl})
        win = tk.Toplevel(self); win.title(f"Drilldown — {kode}"); win.geometry("1000x560")
        header = ttk.Label(win, text=f"A07 sum: {fmt_amount(sum(float(r['beloep']) for r in a07_data))}  •  GL (mappet): {fmt_amount(sum(g['belop'] for g in gl_data))}")
        header.pack(side=tk.TOP, anchor="w", padx=8, pady=8)
        nb = ttk.Notebook(win); nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        tabA = ttk.Frame(nb); nb.add(tabA, text="GL-konti (mappet)")
        fA = ttk.Frame(tabA); fA.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        tA = Table(fA, [("konto","Konto"),("navn","Kontonavn"),("basis","Basis"),("belop","Beløp")]); tA.set_column_format("belop", fmt_amount); tA.insert_rows(gl_data)
        tabB = ttk.Frame(nb); nb.add(tabB, text="A07-detaljer")
        fB = ttk.Frame(tabB); fB.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        tB = Table(fB, [("fnr","Fødselsnr"),("navn","Navn"),("fordel","Fordel"),("beloep","Beløp"),("antall","Antall"),("orgnr","Orgnr")]); tB.set_column_format("beloep", fmt_amount)
        tB.insert_rows([{ "fnr": r["fnr"], "navn": r["navn"], "fordel": r["fordel"], "beloep": r["beloep"], "antall": r["antall"], "orgnr": r["orgnr"]} for r in a07_data])

    # ----- Regelbok -----
    def on_load_rulebook_excel(self):
        if load_rulebook is None:
            messagebox.showwarning("Regelbok", "a07_rulebook.py mangler – kan ikke laste."); return
        path = filedialog.askopenfilename(title="Velg regelbok (Excel)", filetypes=[("Excel","*.xlsx"),("Alle filer","*.*")])
        if not path: return
        try:
            self.rulebook = load_rulebook(path); self.rulebook_source = path; self._refresh_settings_tables()
            self._save_prefs(rulebook_path=path)
            messagebox.showinfo("Regelbok", f"Lest: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Feil ved lesing", str(e))

    def on_load_rulebook_csvdir(self):
        if load_rulebook is None:
            messagebox.showwarning("Regelbok", "a07_rulebook.py mangler – kan ikke laste."); return
        d = filedialog.askdirectory(title="Velg mappe med a07_codes.csv og aliases.csv")
        if not d: return
        try:
            self.rulebook = load_rulebook(d); self.rulebook_source = d; self._refresh_settings_tables()
            self._save_prefs(rulebook_path=d)
            messagebox.showinfo("Regelbok", f"Lest CSV-mappe: {d}")
        except Exception as e:
            messagebox.showerror("Feil ved lesing", str(e))

    # ----- Regelbok-editor -----
    def on_edit_rule(self):
        if not self.rulebook:
            messagebox.showinfo("Regelbok", "Last en regelbok først (Excel/CSV)."); return
        sel = self.tbl_rule_codes.focus()
        if not sel:
            messagebox.showinfo("Velg kode", "Marker en A07‑kode i tabellen."); return
        values = self.tbl_rule_codes.item(sel,"values")
        code = str(values[0]); rule = dict(self.rulebook.get("codes", {}).get(code, {}))

        win = tk.Toplevel(self); win.title(f"Rediger regel — {code}"); win.geometry("680x560")
        frm = ttk.Frame(win); frm.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=10)
        frm.columnconfigure(1, weight=1)

        def _row(r, label, var, hint=""):
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", padx=4, pady=4)
            ent = ttk.Entry(frm, textvariable=var); ent.grid(row=r, column=1, sticky="ew", padx=4, pady=4)
            if hint: ttk.Label(frm, text=hint, foreground="#777").grid(row=r, column=2, sticky="w")
            return ent

        def _format_allowed(intervals):
            if not intervals: return ""
            return " | ".join([f"{lo}-{hi}" if lo!=hi else f"{lo}" for lo,hi in intervals])

        v_category = tk.StringVar(value=rule.get("category","wage"))
        v_basis    = tk.StringVar(value=rule.get("basis","auto"))
        v_allowed  = tk.StringVar(value=_format_allowed(rule.get("allowed",[])))
        v_keywords = tk.StringVar(value=", ".join(sorted(rule.get("keywords", []))))
        v_boost    = tk.StringVar(value=", ".join(sorted(rule.get("boost_accounts", []))))
        es = int(rule.get("expected_sign", 0))
        v_exp      = tk.StringVar(value=("+" if es==1 else ("-" if es==-1 else "")))
        v_special  = tk.StringVar(value=json.dumps(rule.get("special_add", []), ensure_ascii=False))

        _row(0, "Kategori:", v_category)
        _row(1, "Basis (auto|ub|endring|beløp):", v_basis)
        _row(2, "Tillatte kontoområder:", v_allowed, "f.eks. 5000-5399 | 7100-7199")
        _row(3, "Nøkkelord (kommaseparert):", v_keywords)
        _row(4, "Boost‑konti (kommaseparert):", v_boost, "f.eks. 2940, 5290")
        _row(5, "Forventet tegn (+/−/tom):", v_exp, "bruk + eller −")
        ttk.Label(frm, text="Special‑add (JSON‑liste):").grid(row=6, column=0, sticky="nw", padx=4, pady=4)
        txt = tk.Text(frm, height=7); txt.grid(row=6, column=1, columnspan=2, sticky="nsew", padx=4, pady=4)
        txt.insert("1.0", v_special.get())

        def _parse_allowed(expr: str):
            if not expr.strip(): return []
            parts = re.split(r"[|,;]+", expr)
            out = []
            for p in parts:
                p = p.strip()
                if not p: continue
                if "-" in p:
                    a,b = p.split("-",1)
                    a = re.sub(r"\D+","",a); b = re.sub(r"\D+","",b)
                    if a and b: out.append((int(a), int(b)))
                else:
                    v = re.sub(r"\D+","",p)
                    if v: out.append((int(v), int(v)))
            return out

        def save_and_close():
            allowed  = _parse_allowed(v_allowed.get())
            keywords = set(t.strip() for t in v_keywords.get().split(",") if t.strip())
            boosts   = set(re.sub(r"\D+","",t) for t in v_boost.get().split(",") if t.strip())
            exp      = v_exp.get().strip()
            esv      = 1 if exp == "+" else (-1 if exp == "-" else 0)
            try:
                special = json.loads(txt.get("1.0","end").strip() or "[]")
            except Exception as ex:
                messagebox.showerror("JSON‑feil", f"Special‑add må være gyldig JSON‑liste. {ex}")
                return

            rb = self.rulebook
            rb["codes"].setdefault(code, {})
            rb["codes"][code].update({
                "category": v_category.get().strip() or "wage",
                "basis": v_basis.get().strip().lower() or "auto",
                "allowed": allowed,
                "keywords": keywords,
                "boost_accounts": boosts,
                "expected_sign": esv,
                "special_add": special,
            })
            self.rulebook_overrides.setdefault("codes", {})[code] = rb["codes"][code]
            self._refresh_settings_tables(); self.refresh_control_tables()
            win.destroy()

        btns = ttk.Frame(win); btns.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=10)
        ttk.Button(btns, text="Lagre", command=save_and_close).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Avbryt", command=win.destroy).pack(side=tk.RIGHT, padx=6)

    def on_save_rulebook_overrides(self):
        if not self.rulebook_overrides:
            messagebox.showinfo("Ingen endringer", "Det finnes ingen lokale endringer å lagre."); return
        path = filedialog.asksaveasfilename(title="Lagre lokale regelendringer (JSON)", defaultextension=".json", filetypes=[("JSON","*.json")])
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.rulebook_overrides, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Lagret", f"Skrev {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Feil ved lagring", str(e))

    def on_load_rulebook_overrides(self):
        path = filedialog.askopenfilename(title="Velg lokale regelendringer (JSON)", filetypes=[("JSON","*.json")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                overrides = json.load(f)
            self.rulebook_overrides = overrides
            if not self.rulebook: self.rulebook = {"codes": {}, "aliases": {}, "source": "(overrides)"}
            for code, rule in (overrides.get("codes") or {}).items():
                self.rulebook["codes"][code] = {**(self.rulebook.get("codes", {}).get(code, {})), **rule}
            for can, syns in (overrides.get("aliases") or {}).items():
                self.rulebook["aliases"][can] = set(syns)
            self._refresh_settings_tables(); self.refresh_control_tables()
            messagebox.showinfo("Lastet", f"Innlasting OK: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Feil ved innlasting", str(e))

# --------------------------- main ---------------------------

if __name__ == "__main__":
    A07App().mainloop()

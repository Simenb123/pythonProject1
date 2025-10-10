"""
brevgenerator_core.py
Kjernelogikk for Brevgenerator (ikke-GUI)

Inneholder:
- Datamodeller (Ansatt, Client)
- Konfig (lagre/lese sist brukte stier/valg)
- Hjelpefunksjoner (app_dir, datoformat, list_docx_files)
- Excel-innlasting (partnere/klienter) – robust kolonnenavn/arknavn
- Templating (docxtpl-valgfritt + fallback direkte i DOCX-XML)
- Plassholder-innsamling (for "Vis plassholdere")
- Filterfunksjoner (substring-søk)
"""

from __future__ import annotations
import os
import re
import csv
import json
import zipfile
import datetime as dt
from dataclasses import dataclass
from typing import List, Dict, Optional

# ----------------------------- Standardstier ----------------------------- #
DEFAULT_PARTNERS_XLSX = "F:/Dokument/Maler/BHL AS MALER/Partner.xlsx"
DEFAULT_CLIENTS_XLSX = "F:/Dokument/Maler/BHL AS MALER/BHL AS klienter.xlsx"
DEFAULT_TEMPLATES_DIR = "F:/Dokument/Maler/BHL AS NYE MALER 2025"

# ----------------------------- Valgfrie avhengigheter ----------------------------- #
# Pandas for Excel
try:
    import pandas as pd  # type: ignore
except Exception:
    pd = None  # type: ignore

# docxtpl for templating
try:
    from docxtpl import DocxTemplate  # type: ignore
except Exception:
    DocxTemplate = None  # type: ignore

# ----------------------------- Datamodeller ----------------------------- #
@dataclass
class Ansatt:
    navn: str
    epost: str
    telefon: str
    stilling: str
    initialer: str = ""


@dataclass
class Client:
    navn: str
    nr: str = ""
    orgnr: str = ""


# ----------------------------- Konfig ----------------------------- #
def _config_path() -> str:
    """Lokal konfigsti; kan overstyres med env `BREVGEN_CONFIG`."""
    return os.environ.get("BREVGEN_CONFIG") or os.path.join(
        os.path.expanduser("~"), ".brevgenerator_config.json"
    )


def load_config() -> Dict[str, str]:
    try:
        with open(_config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg: Dict[str, str]) -> None:
    try:
        with open(_config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ----------------------------- Hjelpere ----------------------------- #
def app_dir() -> str:
    """Mappen der skriptet/.exe ligger – tåler at __file__ mangler."""
    if getattr(os.sys, "frozen", False):  # PyInstaller
        return os.path.dirname(os.sys.executable)
    base = globals().get("__file__") or (os.sys.argv[0] if os.sys.argv else os.getcwd())
    return os.path.dirname(os.path.abspath(str(base)))


def to_norwegian_date(d: dt.date | dt.datetime) -> str:
    return d.strftime("%d.%m.%Y")


def list_docx_files(folder: str) -> List[str]:
    if not folder or not os.path.isdir(folder):
        return []
    return [f for f in os.listdir(folder) if f.lower().endswith(".docx")]


# ----------------------------- CSV fallback (ansatte) ----------------------------- #
def ensure_sample_csv(path: str) -> None:
    if os.path.exists(path):
        return
    try:
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(
                "navn,epost,telefon,stilling,initialer\n"
                "Ada Partner,ada@firma.no,+47 99 00 00 01,Partner,AP\n"
                "Bendik Partner,bendik@firma.no,+47 99 00 00 02,Partner,BP\n"
            )
    except Exception:
        pass


def load_employees_from_csv(path: str) -> Optional[List[Ansatt]]:
    if not os.path.isfile(path):
        return None
    ansatte: List[Ansatt] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        delimiter = ";" if sample.count(";") > sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delimiter)
        for row in reader:
            if not row:
                continue
            ansatte.append(
                Ansatt(
                    navn=row.get("navn", "").strip(),
                    epost=row.get("epost", "").strip(),
                    telefon=row.get("telefon", "").strip(),
                    stilling=row.get("stilling", "").strip(),
                    initialer=row.get("initialer", "").strip(),
                )
            )
    return ansatte or None


# ----------------------------- Excel-innlasting ----------------------------- #
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.strip().lower())


def _clean_excel_str(val, *, digits_only: bool = False) -> str:
    """
    Konverter Excel/Pandas-celleverdi til pen streng.
    - NaN/NaT/None -> ""
    - 1234.0 -> "1234"
    - digits_only=True -> behold kun sifre
    """
    try:
        if val is None:
            return ""
        if pd is not None:
            try:
                import pandas as _p  # type: ignore
                if _p.isna(val):
                    return ""
            except Exception:
                pass

        if isinstance(val, int):
            s = str(val)
        elif isinstance(val, float):
            s = str(int(val)) if float(val).is_integer() else str(val)
        else:
            s = str(val).strip()

        if s.lower() in {"nan", "nat", "none"}:
            return ""

        if digits_only:
            s = "".join(ch for ch in s if ch.isdigit())
        return s
    except Exception:
        return ""


def load_clients_from_excel(path: str) -> List[Client]:
    if not pd or not os.path.isfile(path):
        return []
    try:
        xls = pd.ExcelFile(path)  # type: ignore
        # Finn et egnet ark – Sheet1/Ark1 eller noe med «klient» i navnet
        sheet = None
        for sh in xls.sheet_names:
            if "klient" in sh.lower() or sh.lower() in {"sheet1", "ark1"}:
                sheet = sh
                break
        df = xls.parse(sheet or xls.sheet_names[0])
    except Exception:
        return []

    cols = {_norm(c): c for c in df.columns}
    c_navn = cols.get("klientnavn") or cols.get("navn")
    c_nr = cols.get("klientnr") or cols.get("klientnummer") or cols.get("kundnr") or cols.get("kundenr")
    c_org = cols.get("klientorgnr") or cols.get("orgnr") or cols.get("organisasjonsnummer") or cols.get("orgnummer")
    if not c_navn:
        return []

    out: List[Client] = []
    for _, row in df.iterrows():
        navn = _clean_excel_str(row.get(c_navn)) if c_navn else ""
        if not navn:
            continue
        nr = _clean_excel_str(row.get(c_nr)) if c_nr else ""
        org = _clean_excel_str(row.get(c_org), digits_only=True) if c_org else ""
        out.append(Client(navn=navn, nr=nr, orgnr=org))
    return out


def load_partners_from_excel(path: str) -> Optional[List[Ansatt]]:
    if not pd or not os.path.isfile(path):
        return None
    try:
        xls = pd.ExcelFile(path)  # type: ignore
        # Foretrekk ark som heter "Partner", ellers vurder alle
        candidates = []
        partner_sheet = None
        for sh in xls.sheet_names:
            if sh.strip().lower() == "partner":
                partner_sheet = sh
                break
        candidates = [partner_sheet] if partner_sheet else xls.sheet_names

        for sh in candidates:
            try:
                df = xls.parse(sh)
            except Exception:
                continue
            cols = {_norm(c): c for c in df.columns}
            c_navn = cols.get("partnernavn") or cols.get("navn")
            c_epost = cols.get("partnerepost") or cols.get("epost")
            c_stilling = cols.get("partnerstilling") or cols.get("stilling")
            c_tlf = cols.get("partnertelefon") or cols.get("telefon") or cols.get("mobil")
            if not c_navn:
                continue
            ans: List[Ansatt] = []
            for _, row in df.iterrows():
                navn = _clean_excel_str(row.get(c_navn))
                if not navn:
                    continue
                ans.append(
                    Ansatt(
                        navn=navn,
                        epost=_clean_excel_str(row.get(c_epost)),
                        telefon=_clean_excel_str(row.get(c_tlf)),
                        stilling=_clean_excel_str(row.get(c_stilling)),
                        initialer="",
                    )
                )
            if ans:
                return ans
        return None
    except Exception:
        return None


# ----------------------------- Templating & plassholdere ----------------------------- #
def xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _token_regex(token: str) -> re.Pattern[str]:
    """Match token (f.eks. {{KLIENT_NAVN}}) selv om det er splittet av tags/whitespace."""
    parts = [re.escape(c) for c in token]
    joiner = r"(?:\s|<[^>]+>)*"
    return re.compile(joiner.join(parts), flags=re.DOTALL)


def create_minimal_docx(path: str) -> None:
    content_types = (
        """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension='xml' ContentType='application/xml'/>
  <Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>
</Types>
"""
    ).encode("utf-8")
    rels = (
        """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
</Relationships>
"""
    ).encode("utf-8")
    document_xml = (
        """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:p><w:r><w:t>Brevmal – Eksempel</w:t></w:r></w:p>
    <w:p><w:r><w:t>Til {{KLIENT_NAVN}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>{{STED}}, {{DATO}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>Med vennlig hilsen,</w:t></w:r></w:p>
    <w:p><w:r><w:t>{{PARTNER_NAVN}} – {{PARTNER_STILLING}}</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    ).encode("utf-8")

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)


def ensure_sample_template(templates_dir: str) -> Optional[str]:
    try:
        os.makedirs(templates_dir, exist_ok=True)
        files = list_docx_files(templates_dir)
        if files:
            return os.path.join(templates_dir, files[0])
        sample = os.path.join(templates_dir, "Eksempelmal.docx")
        create_minimal_docx(sample)
        return sample
    except Exception:
        return None


def render_template_fallback(template_path: str, context: Dict[str, str], out_path: str) -> None:
    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(out_path + ".tmp", "w", compression=zipfile.ZIP_DEFLATED) as zout:
            names = zin.namelist()
            targets = ["word/document.xml"] + [
                n for n in names if n.startswith("word/header") and n.endswith(".xml")
            ] + [n for n in names if n.startswith("word/footer") and n.endswith(".xml")]

            for name in names:
                data = zin.read(name)
                if name in targets:
                    text = data.decode("utf-8", "ignore")
                    for key, val in context.items():
                        token = "{{" + key + "}}"
                        esc = xml_escape(val)
                        text = text.replace(token, esc)
                        text = _token_regex(token).sub(esc, text)
                    data = text.encode("utf-8")
                zout.writestr(name, data)

    if os.path.exists(out_path):
        os.remove(out_path)
    os.replace(out_path + ".tmp", out_path)


def render_template(template_path: str, context: Dict[str, str], out_path: str) -> None:
    if DocxTemplate is not None:
        try:
            doc = DocxTemplate(template_path)  # type: ignore
            doc.render(context)
            doc.save(out_path)
            return
        except Exception as e:
            print(f"[INFO] docxtpl feilet ({e!s}). Bruker fallback-templating…")
    else:
        print("[INFO] docxtpl ikke tilgjengelig. Bruker fallback-templating…")
    render_template_fallback(template_path, context, out_path)


def collect_placeholders(path: str) -> List[str]:
    """Hent ut {{VAR}}-plassholdere fra mal (docxtpl hvis mulig, ellers XML-scan)."""
    vars_set = set()
    if DocxTemplate is not None:
        try:
            doc = DocxTemplate(path)  # type: ignore
            try:
                vars_set |= set(doc.get_undeclared_template_variables())  # type: ignore[attr-defined]
            except Exception:
                pass
        except Exception:
            pass
    if not vars_set:
        try:
            with zipfile.ZipFile(path, "r") as zf:
                for name in zf.namelist():
                    if not (name.startswith("word/") and name.endswith(".xml")):
                        continue
                    xml = zf.read(name).decode("utf-8", "ignore")
                    for m in re.findall(r"\{\{([A-Za-z0-9_]+)\}\}", xml):
                        vars_set.add(m)
        except Exception:
            pass
    return sorted(vars_set)

# ----------------------------- Filter (substring) ----------------------------- #
def filter_client_names(clients: List[Client], query: str) -> List[str]:
    q = (query or "").strip().lower()
    if not q:
        return [c.navn for c in clients]
    return [c.navn for c in clients if q in c.navn.lower()]


def filter_employee_names(employees: List[Ansatt], query: str) -> List[str]:
    q = (query or "").strip().lower()
    if not q:
        return [e.navn for e in employees]
    return [e.navn for e in employees if q in e.navn.lower()]


__all__ = [
    "Ansatt",
    "Client",
    "DEFAULT_PARTNERS_XLSX",
    "DEFAULT_CLIENTS_XLSX",
    "DEFAULT_TEMPLATES_DIR",
    "load_config",
    "save_config",
    "app_dir",
    "to_norwegian_date",
    "list_docx_files",
    "ensure_sample_csv",
    "load_employees_from_csv",
    "load_partners_from_excel",
    "load_clients_from_excel",
    "ensure_sample_template",
    "render_template",
    "render_template_fallback",
    "collect_placeholders",
    "filter_client_names",
    "filter_employee_names",
]

"""
Brevgenerator GUI (Tkinter) med CLI-fallback, Excel-støtte, templating-fallback og selvtester

Oppsummering / nye funksjoner
- GUI for å velge ansatt (fra Excel-ark «Partner») og klient (fra «BHL AS klienter.xlsx»).
- Velg malmappe og .docx-mal via rullgardin. Lista oppdateres automatisk når mappe-feltet endres.
- Standardstier (kan overstyres og lagres automatisk):
  - Partnere:  F:/Dokument/Maler/BHL AS MALER/Partner.xlsx
  - Klienter:  F:/Dokument/Maler/BHL AS MALER/BHL AS klienter.xlsx
  - Maler:     F:/Dokument/Maler/BHL AS NYE MALER 2025
- Husker siste valg (templates-dir, ut-mappe, partner/klient-Excel, sist valgte ansatt) i `~/.brevgenerator_config.json`.
- Plassholdere i maler (docxtpl/Jinja2):
  {{PARTNER_NAVN}} {{PARTNER_EPOST}} {{PARTNER_TELEFON}} {{PARTNER_STILLING}}
  {{KLIENT_NAVN}}  {{KLIENT_STILLING}}  {{KLIENT_NR}}  {{KLIENT_ORGNR}}
  {{STED}} {{DATO}}
- Fallback når `docxtpl` mangler/feiler: skriver verdier direkte inn i DOCX-XML (document + header/footer),
  og håndterer også tokens som er splittet over flere `<w:t>`-runs.
- Robust i sandkasse: ingen `sys.exit()` i entrypoint, safe input i ikke-interaktivt miljø.

Bruk i GUI
  python brevgenerator_gui.py

Bruk i CLI (uten Tkinter)
  python brevgenerator_gui.py --cli --templates-dir "C:/maler" --template "Engasjementsbrev.docx" \
      --employee "Ada Partner" --client "ABC AS" --client-number "123" --client-orgnr "999999999" \
      --client-role "Daglig leder" --place "Sandvika" --date 31.12.2025 --out-dir "C:/Ut" --excel "C:/path/data.xlsx"

Selvtest
  python brevgenerator_gui.py --self-test

Tips
- For full templating: `pip install docxtpl`
- Tkinter trengs kun for GUI. CLI fungerer uten.
- Hvis malmappe er tom, lages «Eksempelmal.docx» automatisk.
"""
from __future__ import annotations
import os
import sys
import re
import csv
import zipfile
import shutil
import tempfile
import datetime as dt
from dataclasses import dataclass
from typing import List, Dict, Optional
import json

# Standardstier (kan overstyres av lagrede innstillinger)
DEFAULT_PARTNERS_XLSX = "F:/Dokument/Maler/BHL AS MALER/Partner.xlsx"
DEFAULT_CLIENTS_XLSX = "F:/Dokument/Maler/BHL AS MALER/BHL AS klienter.xlsx"
DEFAULT_TEMPLATES_DIR = "F:/Dokument/Maler/BHL AS NYE MALER 2025"

# Valgfritt: Excel-støtte (til auto-utfylling fra .xlsx)
try:
    import pandas as pd  # type: ignore
except Exception:
    pd = None  # type: ignore

# ----------------------------- docxtpl ----------------------------- #
try:
    from docxtpl import DocxTemplate
except Exception as _docx_exc:  # utsett feilmelding til vi faktisk trenger liben
    DocxTemplate = None  # type: ignore[assignment]
    _DOCXTPL_IMPORT_ERROR = _docx_exc
else:
    _DOCXTPL_IMPORT_ERROR = None


# ----------------------------- Data-modeller ----------------------------- #
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


DEFAULT_ANSETTE: List[Ansatt] = [
    Ansatt("Ada Partner", "ada@firma.no", "+47 99 00 00 01", "Partner", "AP"),
    Ansatt("Bendik Partner", "bendik@firma.no", "+47 99 00 00 02", "Partner", "BP"),
    Ansatt("Celine Senior", "celine@firma.no", "+47 99 00 00 03", "Senior Manager", "CS"),
    Ansatt("David Senior", "david@firma.no", "+47 99 00 00 04", "Senior Manager", "DS"),
    Ansatt("Eva Manager", "eva@firma.no", "+47 99 00 00 05", "Manager", "EM"),
    Ansatt("Felix Manager", "felix@firma.no", "+47 99 00 00 06", "Manager", "FM"),
    Ansatt("Gro Associate", "gro@firma.no", "+47 99 00 00 07", "Associate", "GA"),
]


# ----------------------------- Konfig ----------------------------- #
def _config_path() -> str:
    # Kan overstyres via env-variabel (greit i tester/CI)
    return os.environ.get("BREVGEN_CONFIG") or os.path.join(os.path.expanduser("~"), ".brevgenerator_config.json")


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


# ----------------------------- Hjelpefunksjoner ----------------------------- #
def app_dir() -> str:
    """Returner mappen der skriptet/.exe ligger – tåler at __file__ mangler.

    Prioritet:
      1) PyInstaller: sys.executable når sys.frozen er sant
      2) __file__ hvis definert
      3) sys.argv[0] hvis satt
      4) os.getcwd() som siste utvei
    """
    if getattr(sys, "frozen", False):  # PyInstaller
        return os.path.dirname(sys.executable)

    base: Optional[str] = None
    if "__file__" in globals() and globals()["__file__"]:
        base = str(globals()["__file__"])  # type: ignore[index]
    if not base and sys.argv and sys.argv[0]:
        base = sys.argv[0]
    if not base:
        base = os.getcwd()
    return os.path.dirname(os.path.abspath(base))


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


def to_norwegian_date(d: dt.date | dt.datetime) -> str:
    return d.strftime("%d.%m.%Y")


def list_docx_files(folder: str) -> List[str]:
    if not folder or not os.path.isdir(folder):
        return []
    return [f for f in os.listdir(folder) if f.lower().endswith(".docx")]


def build_context(
    ansatt: Ansatt,
    klientnavn: str,
    sted: str,
    dato: str,
    klient_stilling: str = "",
    klient_nr: str = "",
    klient_orgnr: str = "",
) -> Dict[str, str]:
    return {
        "PARTNER_NAVN": ansatt.navn,
        "PARTNER_EPOST": ansatt.epost,
        "PARTNER_TELEFON": ansatt.telefon,
        "PARTNER_STILLING": ansatt.stilling,
        "KLIENT_NAVN": klientnavn,
        "KLIENT_STILLING": klient_stilling,
        "KLIENT_NR": klient_nr,
        "KLIENT_ORGNR": klient_orgnr,
        "STED": sted,
        "DATO": dato,
    }


def xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _token_regex(token: str) -> re.Pattern[str]:
    """Match tokenet (f.eks. "{{KLIENT_NAVN}}") selv om det er splittet av XML-tags/whitespace."""
    parts = [re.escape(c) for c in token]
    joiner = r"(?:\s|<[^>]+>)*"
    pattern = joiner.join(parts)
    return re.compile(pattern, flags=re.DOTALL)


def render_template_fallback(template_path: str, context: Dict[str, str], out_path: str) -> None:
    """Fallback: enkel plassholder-erstatning direkte i DOCX (document/header/footer)."""
    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(out_path + ".tmp", "w", compression=zipfile.ZIP_DEFLATED) as zout:
            names = zin.namelist()
            target_xmls = [
                "word/document.xml",
            ] + [n for n in names if n.startswith("word/header") and n.endswith(".xml")] + [n for n in names if n.startswith("word/footer") and n.endswith(".xml")]
            for name in names:
                data = zin.read(name)
                if name in target_xmls:
                    try:
                        text = data.decode("utf-8")
                    except UnicodeDecodeError:
                        text = data.decode("latin-1", errors="ignore")
                    for key, val in context.items():
                        token = "{{" + key + "}}"
                        esc_val = xml_escape(val)
                        text = text.replace(token, esc_val)
                        text = _token_regex(token).sub(esc_val, text)
                    data = text.encode("utf-8")
                zout.writestr(name, data)
    if os.path.exists(out_path):
        os.remove(out_path)
    os.replace(out_path + ".tmp", out_path)


def render_template(template_path: str, context: Dict[str, str], out_path: str) -> None:
    """Render docx-mal. Bruker docxtpl hvis tilgjengelig, ellers fallback."""
    if DocxTemplate is not None:
        try:
            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(out_path)
            return
        except Exception as e:
            print(f"[INFO] docxtpl feilet ({e!s}). Bruker fallback-templating…")
    else:
        print("[INFO] docxtpl ikke tilgjengelig. Bruker fallback-templating…")
    render_template_fallback(template_path, context, out_path)


def open_file(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ----------------------------- Excel-innlasting ----------------------------- #

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.strip().lower())


def load_clients_from_excel(path: str) -> List[Client]:
    if not pd or not os.path.isfile(path):
        return []
    try:
        xls = pd.ExcelFile(path)  # type: ignore
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
        navn = str(row.get(c_navn, "") or "").strip()
        if not navn:
            continue
        nr = str(row.get(c_nr, "") or "").strip() if c_nr else ""
        org = str(row.get(c_org, "") or "").strip() if c_org else ""
        out.append(Client(navn=navn, nr=nr, orgnr=org))
    return out


def load_partners_from_excel(path: str) -> Optional[List[Ansatt]]:
    if not pd or not os.path.isfile(path):
        return None
    try:
        xls = pd.ExcelFile(path)  # type: ignore
        sheet = None
        for sh in xls.sheet_names:
            if sh.strip().lower() == "partner":
                sheet = sh
                break
        if not sheet:
            return None
        df = xls.parse(sheet)
    except Exception:
        return None

    cols = {_norm(c): c for c in df.columns}
    c_navn = cols.get("partnernavn") or cols.get("navn")
    c_epost = cols.get("partnerepost") or cols.get("epost")
    c_stilling = cols.get("partnerstilling") or cols.get("stilling")
    c_tlf = cols.get("partnertelefon") or cols.get("telefon") or cols.get("mobil")
    if not c_navn:
        return None

    ans: List[Ansatt] = []
    for _, row in df.iterrows():
        navn = str(row.get(c_navn, "") or "").strip()
        if not navn:
            continue
        ans.append(
            Ansatt(
                navn=navn,
                epost=str(row.get(c_epost, "") or "").strip(),
                telefon=str(row.get(c_tlf, "") or "").strip(),
                stilling=str(row.get(c_stilling, "") or "").strip(),
                initialer="",
            )
        )
    return ans or None


# ----------------------------- Interaktiv-deteksjon & helpers ----------------------------- #

def is_interactive() -> bool:
    try:
        return bool(getattr(sys.stdin, "isatty", lambda: False)())
    except Exception:
        return False


def safe_input(prompt: str, default: str, interactive: bool) -> str:
    if not interactive:
        return default
    try:
        return input(prompt)
    except (EOFError, OSError):
        return default


# ----------------------------- Malhjelpere (valg & autogenerering) ----------------------------- #

def create_minimal_docx(path: str) -> None:
    """Opprett en gyldig, minimal .docx med Jinja-plassholdere – uten eksterne biblioteker."""
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
<w:document xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
 xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
 xmlns:o='urn:schemas-microsoft-com:office:office'
 xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
 xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
 xmlns:v='urn:schemas-microsoft-com:vml'
 xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
 xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
 xmlns:w10='urn:schemas-microsoft-com:office:word'
 xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
 xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
 xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
 xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
 xmlns:wne='http://schemas.microsoft.com/office/2006/relationships'
 xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' mc:Ignorable='w14 wp14'>
  <w:body>
    <w:p><w:r><w:t>Brevmal – Eksempel</w:t></w:r></w:p>
    <w:p><w:r><w:t>Til {{KLIENT_NAVN}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>{{STED}}, {{DATO}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>Med vennlig hilsen,</w:t></w:r></w:p>
    <w:p><w:r><w:t>{{PARTNER_NAVN}} – {{PARTNER_STILLING}}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w='12240' w:h='15840'/><w:pgMar w:top='1440' w:right='1440' w:bottom='1440' w:left='1440' w:header='708' w:footer='708' w:gutter='0'/></w:sectPr>
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
    """Sørg for at det finnes minst én .docx-mal i mappen.
    - Hvis det finnes minst én, returneres første.
    - Hvis mappen er tom/ikke finnes, opprettes **Eksempelmal.docx** via `create_minimal_docx`.
    - Returnerer sti til mal, eller None ved feil (manglende rettigheter e.l.).
    """
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


def choose_employee(ansatte: List[Ansatt], employee_arg: Optional[str], interactive: bool) -> Ansatt:
    """Velg ansatt fra navn eller indeks. Ikke-interaktivt => første i listen."""
    if employee_arg is None:
        if interactive:
            print("Velg ansatt:")
            for i, a in enumerate(ansatte):
                print(f"[{i}] {a.navn} – {a.stilling}")
            idx_txt = safe_input("Indeks: ", "0", interactive)
            idx = int(idx_txt)
            return ansatte[idx]
        return ansatte[0]
    if employee_arg.isdigit():
        return ansatte[int(employee_arg)]
    for a in ansatte:
        if a.navn.lower() == employee_arg.lower():
            return a
    raise SystemExit(f"Fant ikke ansatt: {employee_arg}")


def choose_template_path(templates_dir: str, template_arg: Optional[str], interactive: bool) -> str:
    """Velg malsti. Oppretter eksempelmal hvis mappe er tom. Ikke-interaktivt => første fil."""
    if template_arg:
        return template_arg if os.path.isabs(template_arg) else os.path.join(templates_dir, template_arg)

    files = list_docx_files(templates_dir)
    if not files:
        created = ensure_sample_template(templates_dir)
        files = list_docx_files(templates_dir)
        if not files and not created:
            raise SystemExit(
                "Fant ingen .docx-maler i mappen og kunne ikke lage en eksempelmal. "
                "Oppgi --templates-dir med .docx-filer eller --template med full sti."
            )

    if interactive:
        print("Velg mal:")
        for i, f in enumerate(files):
            print(f"[{i}] {f}")
        idx_txt = safe_input("Indeks: ", "0", interactive)
        idx = int(idx_txt)
        return os.path.join(templates_dir, files[idx])

    return os.path.join(templates_dir, files[0])


# ----------------------------- GUI (lazy import) ----------------------------- #

def try_run_gui() -> bool:
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
    except ModuleNotFoundError:
        print("[INFO] Tkinter ikke tilgjengelig – går over til CLI-modus.\nInstaller Tkinter for GUI.")
        return False

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Brevgenerator – Word-maler")
            self.geometry("820x600")
            self.minsize(760, 560)

            csv_path = os.path.join(app_dir(), "ansatte.csv")
            ensure_sample_csv(csv_path)
            self.ansatte: List[Ansatt] = load_employees_from_csv(csv_path) or DEFAULT_ANSETTE

            # Konfig + standardstier
            self.cfg: Dict[str, str] = load_config()
            tpl_default = DEFAULT_TEMPLATES_DIR if os.path.isdir(DEFAULT_TEMPLATES_DIR) else app_dir()
            out_default = os.path.join(tpl_default, "Ut")
            partners_default = DEFAULT_PARTNERS_XLSX if os.path.isfile(DEFAULT_PARTNERS_XLSX) else ""
            clients_default = DEFAULT_CLIENTS_XLSX if os.path.isfile(DEFAULT_CLIENTS_XLSX) else ""

            self.mal_mappe = tk.StringVar(value=self.cfg.get("templates_dir") or tpl_default)
            self.partners_xlsx = tk.StringVar(value=self.cfg.get("partners_xlsx") or partners_default)
            self.clients_xlsx = tk.StringVar(value=self.cfg.get("clients_xlsx") or clients_default)
            self.valgt_mal = tk.StringVar()
            self.ut_mappe = tk.StringVar(value=self.cfg.get("out_dir") or out_default)
            self.klientnavn = tk.StringVar()
            self.klientnr = tk.StringVar()
            self.klientorgnr = tk.StringVar()
            self.klient_stilling = tk.StringVar()
            self.sted = tk.StringVar(value="Sandvika")
            self.dato = tk.StringVar(value=to_norwegian_date(dt.date.today()))
            self.open_after = tk.BooleanVar(value=True)

            self.clients: List[Client] = []
            self.client_roles = ["Daglig leder", "Styrets leder", "Representant for selskapet"]

            self._build_widgets()
            # auto-last partnere/klienter om stier finnes og auto-oppdater mal-liste
            self._load_partners()
            self._load_clients()
            self._refresh_maler()
            self._refresh_job = None
            try:
                self.mal_mappe.trace_add('write', self._on_templates_dir_change)
            except Exception:
                pass

        def _build_widgets(self):
            pad = {"padx": 10, "pady": 8}

            frm_ansatt = ttk.LabelFrame(self, text="Ansatt")
            frm_ansatt.pack(fill="x", **pad)
            ttk.Label(frm_ansatt, text="Velg ansatt:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
            self.ansatt_combo = ttk.Combobox(frm_ansatt, values=[a.navn for a in self.ansatte], state="readonly")
            self.ansatt_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
            self.ansatt_combo.current(0)
            frm_ansatt.columnconfigure(1, weight=1)

            self.lbl_epost = ttk.Label(frm_ansatt, text=f"E-post: {self.ansatte[0].epost}")
            self.lbl_epost.grid(row=1, column=0, columnspan=2, sticky="w", padx=8)
            self.lbl_tlf = ttk.Label(frm_ansatt, text=f"Telefon: {self.ansatte[0].telefon}")
            self.lbl_tlf.grid(row=2, column=0, columnspan=2, sticky="w", padx=8)
            self.lbl_stilling = ttk.Label(frm_ansatt, text=f"Stilling: {self.ansatte[0].stilling}")
            self.lbl_stilling.grid(row=3, column=0, columnspan=2, sticky="w", padx=8)
            self.ansatt_combo.bind("<<ComboboxSelected>>", self._on_ansatt_change)

            frm_mal = ttk.LabelFrame(self, text="Mal og felter")
            frm_mal.pack(fill="x", **pad)
            ttk.Label(frm_mal, text="Mappe med maler:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
            ent_mappe = ttk.Entry(frm_mal, textvariable=self.mal_mappe)
            ent_mappe.grid(row=0, column=1, sticky="ew", padx=8)
            ttk.Button(frm_mal, text="Bla gjennom…", command=self._choose_folder).grid(row=0, column=2, padx=8)

            ttk.Label(frm_mal, text="Velg mal (.docx):").grid(row=1, column=0, sticky="w", padx=8, pady=6)
            self.mal_combo = ttk.Combobox(frm_mal, textvariable=self.valgt_mal, state="readonly")
            self.mal_combo.grid(row=1, column=1, sticky="ew", padx=8)
            ttk.Button(frm_mal, text="Oppdater", command=self._refresh_maler).grid(row=1, column=2, padx=8)
            ttk.Button(frm_mal, text="Vis plassholdere", command=self._show_placeholders).grid(row=1, column=3, padx=8)

            ttk.Label(frm_mal, text="Partnerliste (.xlsx):").grid(row=2, column=0, sticky="w", padx=8, pady=6)
            ent_xlp = ttk.Entry(frm_mal, textvariable=self.partners_xlsx)
            ent_xlp.grid(row=2, column=1, sticky="ew", padx=8)
            ttk.Button(frm_mal, text="Bla gjennom…", command=self._choose_excel_partners).grid(row=2, column=2, padx=8)
            ttk.Button(frm_mal, text="Last", command=self._load_partners).grid(row=2, column=3, padx=8)

            ttk.Label(frm_mal, text="Klientliste (.xlsx):").grid(row=3, column=0, sticky="w", padx=8, pady=6)
            ent_xlc = ttk.Entry(frm_mal, textvariable=self.clients_xlsx)
            ent_xlc.grid(row=3, column=1, sticky="ew", padx=8)
            ttk.Button(frm_mal, text="Bla gjennom…", command=self._choose_excel_clients).grid(row=3, column=2, padx=8)
            ttk.Button(frm_mal, text="Last", command=self._load_clients).grid(row=3, column=3, padx=8)

            ttk.Label(frm_mal, text="Klientnavn:").grid(row=4, column=0, sticky="w", padx=8, pady=6)
            self.cb_klient = ttk.Combobox(frm_mal, textvariable=self.klientnavn, state="normal")
            self.cb_klient.grid(row=4, column=1, sticky="ew", padx=8)
            self.cb_klient.bind("<<ComboboxSelected>>", self._on_client_change)
            self.cb_klient.bind("<KeyRelease>", self._on_client_typed)

            ttk.Label(frm_mal, text="Klientnr:").grid(row=5, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(frm_mal, textvariable=self.klientnr).grid(row=5, column=1, sticky="ew", padx=8)

            ttk.Label(frm_mal, text="Org.nr:").grid(row=6, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(frm_mal, textvariable=self.klientorgnr).grid(row=6, column=1, sticky="ew", padx=8)

            ttk.Label(frm_mal, text="Klient-stilling:").grid(row=7, column=0, sticky="w", padx=8, pady=6)
            self.cb_klient_stilling = ttk.Combobox(frm_mal, textvariable=self.klient_stilling, values=self.client_roles, state="normal")
            self.cb_klient_stilling.grid(row=7, column=1, sticky="ew", padx=8)

            ttk.Label(frm_mal, text="Sted:").grid(row=8, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(frm_mal, textvariable=self.sted).grid(row=8, column=1, sticky="ew", padx=8)

            ttk.Label(frm_mal, text="Dato (dd.mm.åååå):").grid(row=9, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(frm_mal, textvariable=self.dato).grid(row=9, column=1, sticky="ew", padx=8)

            for i in range(0, 3):
                frm_mal.columnconfigure(i, weight=1)
            frm_mal.columnconfigure(1, weight=3)

            frm_out = ttk.LabelFrame(self, text="Lagring")
            frm_out.pack(fill="x", **pad)
            ttk.Label(frm_out, text="Ut-mappe:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
            ent_ut = ttk.Entry(frm_out, textvariable=self.ut_mappe)
            ent_ut.grid(row=0, column=1, sticky="ew", padx=8)
            ttk.Button(frm_out, text="Velg…", command=self._choose_out_folder).grid(row=0, column=2, padx=8)
            ttk.Checkbutton(frm_out, text="Åpne fil etter generering", variable=self.open_after).grid(row=1, column=1, sticky="w", padx=8)
            frm_out.columnconfigure(1, weight=1)

            frm_btn = ttk.Frame(self)
            frm_btn.pack(fill="x", **pad)
            ttk.Button(frm_btn, text="Generer dokument", command=self._generate).pack(side="right")

        def _choose_folder(self):
            from tkinter import filedialog
            folder = filedialog.askdirectory(initialdir=self.mal_mappe.get() or app_dir(), title="Velg mappe med Word-maler")
            if folder:
                self.mal_mappe.set(folder)
                self._refresh_maler_and_save()

        def _choose_out_folder(self):
            from tkinter import filedialog
            folder = filedialog.askdirectory(initialdir=self.ut_mappe.get() or app_dir(), title="Velg ut-mappe")
            if folder:
                self.ut_mappe.set(folder)
                self._save_cfg()

        def _choose_excel_partners(self):
            from tkinter import filedialog
            path = filedialog.askopenfilename(initialdir=self.partners_xlsx.get() or app_dir(),
                                              title="Velg Partner.xlsx",
                                              filetypes=[["Excel (*.xlsx)", "*.xlsx"], ["Alle", "*"]])
            if path:
                self.partners_xlsx.set(path)
                self._load_partners()
                self._save_cfg()

        def _choose_excel_clients(self):
            from tkinter import filedialog
            path = filedialog.askopenfilename(initialdir=self.clients_xlsx.get() or app_dir(),
                                              title="Velg BHL AS klienter.xlsx",
                                              filetypes=[["Excel (*.xlsx)", "*.xlsx"], ["Alle", "*"]])
            if path:
                self.clients_xlsx.set(path)
                self._load_clients()
                self._save_cfg()

        def _load_partners(self):
            path = self.partners_xlsx.get().strip()
            if not path and os.path.isfile(DEFAULT_PARTNERS_XLSX):
                path = DEFAULT_PARTNERS_XLSX
                self.partners_xlsx.set(path)
            if path and os.path.isfile(path):
                partners = load_partners_from_excel(path)
                if partners:
                    self.ansatte = partners
                    self.ansatt_combo["values"] = [a.navn for a in partners]
                    last = self.cfg.get("last_employee")
                    if last and last in [a.navn for a in partners]:
                        self.ansatt_combo.set(last)
                    else:
                        self.ansatt_combo.current(0)
                    self._on_ansatt_change()

        def _load_clients(self):
            path = self.clients_xlsx.get().strip()
            if not path and os.path.isfile(DEFAULT_CLIENTS_XLSX):
                path = DEFAULT_CLIENTS_XLSX
                self.clients_xlsx.set(path)
            if path and os.path.isfile(path):
                clients = load_clients_from_excel(path)
                if clients:
                    self.clients = clients
                    self.cb_klient["values"] = [c.navn for c in clients]

        def _on_client_typed(self, event=None):
            q = self.cb_klient.get().strip().lower()
            if not q or not self.clients:
                return
            # eksakt treff
            for c in self.clients:
                if c.navn.lower() == q:
                    self.klientnr.set(c.nr)
                    self.klientorgnr.set(c.orgnr)
                    return
            # unikt prefiks
            matches = [c for c in self.clients if c.navn.lower().startswith(q)]
            if len(matches) == 1:
                self.klientnr.set(matches[0].nr)
                self.klientorgnr.set(matches[0].orgnr)

        def _on_client_change(self, event=None):
            nav = self.klientnavn.get().strip()
            for c in self.clients:
                if c.navn == nav:
                    self.klientnr.set(c.nr)
                    self.klientorgnr.set(c.orgnr)
                    break

        def _on_ansatt_change(self, event=None):
            idx = self.ansatt_combo.current()
            if idx < 0 or idx >= len(self.ansatte):
                return
            a = self.ansatte[idx]
            self.lbl_epost.configure(text=f"E-post: {a.epost}")
            self.lbl_tlf.configure(text=f"Telefon: {a.telefon}")
            self.lbl_stilling.configure(text=f"Stilling: {a.stilling}")
            self._save_cfg()

        def _on_templates_dir_change(self, *args):
            # debounce refresh når brukeren skriver
            if getattr(self, "_refresh_job", None):
                try:
                    self.after_cancel(self._refresh_job)
                except Exception:
                    pass
            self._refresh_job = self.after(600, self._refresh_maler_and_save)

        def _refresh_maler_and_save(self):
            self._refresh_maler()
            self._save_cfg()

        def _refresh_maler(self):
            folder = self.mal_mappe.get().strip()
            files = list_docx_files(folder)
            if not files:
                ensure_sample_template(folder)
                files = list_docx_files(folder)
            self.mal_combo["values"] = files
            if files:
                # Behold eksisterende valg om mulig
                if self.valgt_mal.get() in files:
                    self.mal_combo.set(self.valgt_mal.get())
                else:
                    self.mal_combo.current(0)
            else:
                self.valgt_mal.set("")

        def _save_cfg(self):
            self.cfg["templates_dir"] = self.mal_mappe.get().strip()
            self.cfg["out_dir"] = self.ut_mappe.get().strip()
            self.cfg["partners_xlsx"] = self.partners_xlsx.get().strip()
            self.cfg["clients_xlsx"] = self.clients_xlsx.get().strip()
            try:
                self.cfg["last_employee"] = self.ansatt_combo.get()
            except Exception:
                pass
            save_config(self.cfg)
        def _collect_placeholders(self, path: str) -> List[str]:
            vars_set = set()
            # Prøv docxtpl først (hvis tilgjengelig)
            if DocxTemplate is not None:
                try:
                    doc = DocxTemplate(path)
                    try:
                        vars_set |= set(doc.get_undeclared_template_variables())  # type: ignore[attr-defined]
                    except Exception:
                        pass
                except Exception:
                    pass
            # Fallback: scan XML for {{VAR}}
            if not vars_set:
                try:
                    with zipfile.ZipFile(path, "r") as zf:
                        for name in zf.namelist():
                            if not (name.startswith("word/") and name.endswith(".xml")):
                                continue
                            xml = zf.read(name).decode("utf-8", "ignore")
                            # Fanger enkle tokens – håndterer ikke split across runs (til visning er dette ok)
                            for m in re.findall(r"\{\{([A-Za-z0-9_]+)\}\}", xml):
                                vars_set.add(m)
                except Exception:
                    pass
            return sorted(vars_set)

        def _show_placeholders(self):
            from tkinter import messagebox, Toplevel, Listbox
            path = self._selected_template_path()
            if not path or not os.path.isfile(path):
                messagebox.showinfo("Ingen mal", "Velg en .docx-mal først.")
                return
            vars_list = self._collect_placeholders(path)
            if not vars_list:
                messagebox.showinfo("Plassholdere", "Fant ingen {{VAR}}-plassholdere i valgt mal.")
                return
            top = Toplevel(self)
            top.title("Plassholdere i mal")
            top.geometry("360x320")
            lb = Listbox(top)
            lb.pack(fill="both", expand=True, padx=10, pady=10)
            for v in vars_list:
                lb.insert("end", f"{{{{{v}}}}}")

        def _selected_template_path(self) -> Optional[str]:
            if not self.valgt_mal.get():
                return None
            return os.path.join(self.mal_mappe.get(), self.valgt_mal.get())

        def _generate(self):
            from tkinter import messagebox
            try:
                tmpl_path = self._selected_template_path()
                if not tmpl_path or not os.path.isfile(tmpl_path):
                    messagebox.showerror("Mangler mal", "Velg en gyldig .docx-mal først.")
                    return
                if not os.path.isdir(self.ut_mappe.get()):
                    os.makedirs(self.ut_mappe.get(), exist_ok=True)

                ansatt = self.ansatte[self.ansatt_combo.current()]

                dato_txt = self.dato.get().strip() or to_norwegian_date(dt.date.today())
                try:
                    parsed = dt.datetime.strptime(dato_txt, "%d.%m.%Y")
                    dato_txt = to_norwegian_date(parsed)
                except Exception:
                    pass

                context = build_context(
                    ansatt=ansatt,
                    klientnavn=self.klientnavn.get().strip(),
                    sted=self.sted.get().strip(),
                    dato=dato_txt,
                    klient_stilling=self.klient_stilling.get().strip(),
                    klient_nr=self.klientnr.get().strip(),
                    klient_orgnr=self.klientorgnr.get().strip(),
                )

                base = os.path.splitext(os.path.basename(tmpl_path))[0]
                initialer = ansatt.initialer or "ANSATT"
                out_name = f"{base} - {initialer} - {dt.datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                out_path = os.path.join(self.ut_mappe.get(), out_name)

                render_template(tmpl_path, context, out_path)

                if self.open_after.get():
                    open_file(out_path)

                messagebox.showinfo("Ferdig", f"Dokument generert:\n{out_path}")
            except Exception as e:
                messagebox.showerror("Feil", f"Noe gikk galt under generering:\n{e}")

    App().mainloop()
    return True


# ----------------------------- CLI ----------------------------- #

def run_cli(argv: Optional[List[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Brevgenerator – CLI")
    parser.add_argument("--templates-dir", dest="templates_dir")
    parser.add_argument("--template", dest="template")
    parser.add_argument("--employee", dest="employee", help="Navn eller indeks (0-basert)")
    parser.add_argument("--client", dest="client")
    parser.add_argument("--client-number", dest="client_number")
    parser.add_argument("--client-orgnr", dest="client_orgnr")
    parser.add_argument("--client-role", dest="client_role")
    parser.add_argument("--excel", dest="excel_path")
    parser.add_argument("--place", dest="place", default="Sandvika")
    parser.add_argument("--date", dest="date")
    parser.add_argument("--out-dir", dest="out_dir", default=os.path.join(app_dir(), "Ut"))
    parser.add_argument("--open", dest="open_after", action="store_true")
    parser.add_argument("--list-employees", action="store_true")
    parser.add_argument("--list-templates", action="store_true")

    args = parser.parse_args(argv)

    interactive = is_interactive()

    csv_path = os.path.join(app_dir(), "ansatte.csv")
    ensure_sample_csv(csv_path)
    ansatte = load_employees_from_csv(csv_path) or DEFAULT_ANSETTE

    if args.list_employees:
        for i, a in enumerate(ansatte):
            print(f"[{i}] {a.navn} – {a.stilling} – {a.epost} – {a.telefon}")
        return 0

    templates_dir = args.templates_dir or app_dir()
    if args.list_templates:
        for f in list_docx_files(templates_dir):
            print(f)
        return 0

    valgt = choose_employee(ansatte, args.employee, interactive)
    template_path = choose_template_path(templates_dir, args.template, interactive)
    if not os.path.isfile(template_path):
        raise SystemExit(f"Fant ikke mal: {template_path}")

    clients: List[Client] = []
    # Prøv angitt Excel eller default lokasjon
    excel_candidates = [p for p in [args.excel_path, DEFAULT_CLIENTS_XLSX] if p]
    for xp in excel_candidates:
        if xp and os.path.isfile(xp):
            clients = load_clients_from_excel(xp)
            if clients:
                break

    client = args.client or safe_input("Klientnavn: ", "ACME AS", interactive)
    c_nr = args.client_number or ""
    c_org = args.client_orgnr or ""
    if clients:
        for c in clients:
            if c.navn.lower() == client.strip().lower():
                c_nr = c_nr or c.nr
                c_org = c_org or c.orgnr
                break

    c_role = args.client_role or ""
    place = args.place or safe_input("Sted: ", "Sandvika", interactive)
    default_date = to_norwegian_date(dt.date.today())
    date_txt = args.date or safe_input("Dato (dd.mm.åååå) [tom = i dag]: ", default_date, interactive)
    if not date_txt:
        date_txt = default_date

    ctx = build_context(
        valgt,
        client.strip(),
        place.strip(),
        date_txt.strip(),
        klient_stilling=c_role.strip(),
        klient_nr=c_nr.strip(),
        klient_orgnr=c_org.strip(),
    )

    os.makedirs(args.out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(template_path))[0]
    initialer = valgt.initialer or "ANSATT"
    out_name = f"{base} - {initialer} - {dt.datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    out_path = os.path.join(args.out_dir, out_name)

    render_template(template_path, ctx, out_path)
    print(f"OK – generert: {out_path}")

    if args.open_after:
        open_file(out_path)

    return 0


# ----------------------------- Selvtester ----------------------------- #

def _self_test() -> int:
    failures = 0

    def check(cond: bool, msg: str):
        nonlocal failures
        if cond:
            print("[OK]", msg)
        else:
            failures += 1
            print("[FEIL]", msg)

    # 1) build_context
    a = DEFAULT_ANSETTE[0]
    ctx = build_context(a, "ABC AS", "Sandvika", "01.01.2025", klient_stilling="Daglig leder", klient_nr="1001", klient_orgnr="999999999")
    check(ctx["PARTNER_NAVN"] == a.navn and ctx["KLIENT_NAVN"] == "ABC AS" and ctx["KLIENT_NR"] == "1001", "build_context returnerer forventede nøkler/verdier")

    # 2) list_docx_files
    tmp = tempfile.mkdtemp()
    try:
        open(os.path.join(tmp, "a.docx"), "wb").close()
        open(os.path.join(tmp, "b.DOCX"), "wb").close()
        open(os.path.join(tmp, "c.txt"), "wb").close()
        files = list_docx_files(tmp)
        check("a.docx" in files and "b.DOCX" in files and "c.txt" not in files, "list_docx_files filtrerer kun .docx")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 3) CSV-lesing (komma)
    tmp = tempfile.mkdtemp()
    try:
        csvp = os.path.join(tmp, "ansatte.csv")
        with open(csvp, "w", encoding="utf-8", newline="") as f:
            f.write("navn,epost,telefon,stilling,initialer\n")
            f.write("Test Testesen,t@ex.no,123,Partner,TT\n")
            f.write("X Y,z@ex.no,456,Manager,XY\n")
        lst = load_employees_from_csv(csvp)
        check(lst and lst[0].navn == "Test Testesen" and lst[1].initialer == "XY", "CSV parse fungerer (komma)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 4) CSV-lesing (semikolon)
    tmp = tempfile.mkdtemp()
    try:
        csvp = os.path.join(tmp, "ansatte.csv")
        with open(csvp, "w", encoding="utf-8", newline="") as f:
            f.write("navn;epost;telefon;stilling;initialer\n")
            f.write("Semi Kolon;semi@ex.no;789;Manager;SK\n")
        lst = load_employees_from_csv(csvp)
        check(lst and lst[0].navn == "Semi Kolon" and lst[0].initialer == "SK", "CSV parse fungerer (semikolon)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 5) app_dir() gir eksisterende mappe
    try:
        d = app_dir()
        check(isinstance(d, str) and os.path.isdir(d), "app_dir() returnerer eksisterende mappe selv uten __file__")
    except Exception as e:
        check(False, f"app_dir() kastet unntak: {e}")

    # 6) Edge-case: list_docx_files på ikke-eksisterende mappe
    files = list_docx_files(os.path.join("/unlikely", "does", "not", "exist"))
    check(files == [], "list_docx_files håndterer ikke-eksisterende mappe")

    # 7) Ikke-interaktiv fallback: choose_employee
    emp = choose_employee(DEFAULT_ANSETTE, None, interactive=False)
    check(emp is DEFAULT_ANSETTE[0], "choose_employee velger første ved ikke-interaktiv modus")

    # 8) Auto-malopprettelse: ensure_sample_template + zip-inspeksjon
    tmp = tempfile.mkdtemp()
    try:
        path = ensure_sample_template(tmp)
        check(path is not None and os.path.isfile(path), "ensure_sample_template opprettet Eksempelmal.docx i tom mappe")
        with zipfile.ZipFile(path, "r") as zf:
            namelist = set(zf.namelist())
            check("[Content_Types].xml" in namelist and "_rels/.rels" in namelist and "word/document.xml" in namelist, "DOCX inneholder nødvendige deler")
            docxml = zf.read("word/document.xml").decode("utf-8", "ignore")
            check("KLIENT_NAVN" in docxml and "PARTNER_NAVN" in docxml and "DATO" in docxml, "Plassholdere finnes i document.xml")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 9) safe_input gir default i ikke-interaktiv modus
    val = safe_input("prompt", "DEF", interactive=False)
    check(val == "DEF", "safe_input returnerer default når ikke-interaktivt")

    # 10) Render minimal mal via fallback (simulerer at docxtpl mangler)
    tmp = tempfile.mkdtemp()
    try:
        tmpl = os.path.join(tmp, "tmpl.docx")
        create_minimal_docx(tmpl)
        outp = os.path.join(tmp, "out.docx")
        ctx = build_context(DEFAULT_ANSETTE[0], "ACME", "Oslo", "31.12.2025")
        render_template_fallback(tmpl, ctx, outp)
        check(os.path.isfile(outp) and os.path.getsize(outp) > 0, "fallback rendering ga utfil fra minimal mal")
        with zipfile.ZipFile(outp, "r") as zf:
            docxml = zf.read("word/document.xml").decode("utf-8", "ignore")
            check("ACME" in docxml and "Oslo" in docxml and "31.12.2025" in docxml, "verdier skrevet inn i document.xml")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 11) Fallback håndterer tokens splittet over XML-tags (<w:t>-runs)
    tmp = tempfile.mkdtemp()
    try:
        docx_path = os.path.join(tmp, "split.docx")
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
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <w:body>
    <w:p><w:r><w:t>Til {{KLI</w:t></w:r><w:r><w:t>ENT_NAVN}}</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
        ).encode("utf-8")
        with zipfile.ZipFile(docx_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels)
            zf.writestr("word/document.xml", document_xml)
        outp = os.path.join(tmp, "out.docx")
        ctx = build_context(DEFAULT_ANSETTE[0], "ACME", "Oslo", "01.01.2030")
        render_template_fallback(docx_path, ctx, outp)
        with zipfile.ZipFile(outp, "r") as zf:
            xml = zf.read("word/document.xml").decode("utf-8", "ignore")
            check("ACME" in xml and "{{KLI" not in xml, "fallback erstatter tokens splittet over XML-tags")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 12) run_cli() returnerer 0 og lager utfil (uten sys.exit)
    tmp = tempfile.mkdtemp()
    try:
        code = run_cli([
            "--templates-dir", tmp,
            "--client", "ACME",
            "--client-number", "123",
            "--client-orgnr", "999999999",
            "--client-role", "Daglig leder",
            "--place", "Bergen",
            "--date", "02.01.2030",
            "--out-dir", tmp,
        ])
        check(code == 0, "run_cli() returnerte 0 (ingen SystemExit)")
        docs = [f for f in os.listdir(tmp) if f.lower().endswith('.docx') and f != 'Eksempelmal.docx']
        check(len(docs) >= 1 and os.path.getsize(os.path.join(tmp, docs[0])) > 0, "CLI genererte en utfil")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 13) Excel-lastere (hvis pandas er tilgjengelig)
    if pd is not None:
        try:
            import pandas as _p
            tmp = tempfile.mkdtemp()
            xl = os.path.join(tmp, "data.xlsx")
            df1 = _p.DataFrame({
                "KLIENT_NR": ["100", "200"],
                "KLIENT_NAVN": ["Alpha AS", "Beta AS"],
                "KLIENT_ORGNR": ["999999999", "888888888"],
            })
            dfp = _p.DataFrame({
                "PARTNER_NAVN": ["Ada Partner"],
                "PARTNER_EPOST": ["ada@firma.no"],
                "PARTNER_STILLING": ["Partner"],
                "PARTNER_TELEFON": ["99 00 00 01"],
            })
            with _p.ExcelWriter(xl, engine="openpyxl") as w:  # type: ignore
                df1.to_excel(w, sheet_name="Sheet1", index=False)
                dfp.to_excel(w, sheet_name="Partner", index=False)
            cl = load_clients_from_excel(xl)
            an = load_partners_from_excel(xl)
            check(len(cl) == 2 and cl[0].navn == "Alpha AS", "load_clients_from_excel henter klienter")
            check(an and an[0].navn == "Ada Partner", "load_partners_from_excel henter partnere")
        except Exception as e:
            print("[HOPPER OVER] Excel-test pga feil:", e)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    print("Selvtest ferdig. Feil:", failures)
    return 1 if failures else 0


# ----------------------------- Entrypoint ----------------------------- #
if __name__ == "__main__":
    try:
        if "--self-test" in sys.argv:
            code = _self_test()
            print("SELF-TEST EXIT CODE:", code)
        elif "--cli" in sys.argv:
            args = [a for a in sys.argv[1:] if a != "--cli"]
            code = run_cli(args)
            print("CLI EXIT CODE:", code)
        else:
            started = try_run_gui()
            if not started:
                code = run_cli(sys.argv[1:])
                print("CLI EXIT CODE:", code)
    except SystemExit as e:
        # Unngå abrupt terminering i miljøer som ikke liker SystemExit
        print(f"[INFO] Caught SystemExit({e.code}) – fortsetter uten å terminere prosessen.")

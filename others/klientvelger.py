# klientvelger.py – 2025-06-09  (r13 – StringVar(value=…))
# --------------------------------------------------------------------
from __future__ import annotations
import csv, json, logging, subprocess, sys, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import chardet, pandas as pd
from src.app.services.mapping_utils import FeltVelger

ROOT_DIR = Path(r"C:\Users\ib91\Desktop\Prosjekt\Klienter")
LAST_FILE = ROOT_DIR / "_last_client.txt"
BILAG_GUI = Path(__file__).with_name("bilag_gui_tk.py")
META_NAME = ".klient_meta.json"

logger = logging.getLogger(__name__)

# ---------- CSV-header helper ---------------------------------------
def _csv_header(p: Path) -> pd.DataFrame:
    raw = p.read_bytes()
    enc = chardet.detect(raw)["encoding"] or "utf-8"
    try:
        sample = raw[:50000].decode("latin1", errors="ignore")
        delim = csv.Sniffer().sniff(sample, delimiters=";,|\t").delimiter
    except Exception:
        delim = ";"
    return pd.read_csv(p, sep=delim, encoding=enc, engine="python", nrows=1)

# ---------- små-hjelpere --------------------------------------------
def list_klientmapper() -> list[str]:
    return sorted(p.name for p in ROOT_DIR.iterdir() if p.is_dir())

def les_meta(d: Path) -> dict:
    f = d / META_NAME
    if f.exists():
        try:
            return json.loads(f.read_text("utf-8"))
        except Exception:
            pass
    return {}

def skriv_meta(d: Path, meta: dict):
    (d / META_NAME).write_text(json.dumps(meta, indent=2), "utf-8")

def lagre_sist_klient(n: str):
    try:
        LAST_FILE.write_text(n, "utf-8")
    except Exception:
        pass

def hent_sist_klient() -> str | None:
    try:
        t = LAST_FILE.read_text("utf-8").strip()
        return t if t in list_klientmapper() else None
    except Exception:
        return None

# ---------- GUI ------------------------------------------------------
class KlientVelger(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Klientvelger"); self.resizable(False, False)

        ttk.Label(self, text="Søk klient:").grid(row=0, column=0, sticky="w")
        self.sok = tk.StringVar(value="")
        ent = ttk.Entry(self, textvariable=self.sok, width=25)
        ent.grid(row=0, column=1, columnspan=3, sticky="we", pady=2)
        ent.bind("<KeyRelease>", self._filter)

        ttk.Label(self, text="Velg klient:").grid(row=1, column=0, sticky="w")
        self.cli_var = tk.StringVar(value="")
        self.cli_cmb = ttk.Combobox(self, textvariable=self.cli_var,
                                    values=list_klientmapper(),
                                    state="readonly", width=30)
        self.cli_cmb.grid(row=1, column=1, columnspan=3, sticky="we", pady=2)
        self.cli_cmb.bind("<<ComboboxSelected>>", self._oppdater_type)

        if last := hent_sist_klient():
            self.cli_cmb.set(last); self.after(50, self._oppdater_type)

        ttk.Label(self, text="Datakilde:").grid(row=2, column=0, sticky="w")
        self.type_var = tk.StringVar(value="")
        ttk.Combobox(self, textvariable=self.type_var,
                     values=["Hovedbok", "Saldobalanse"],
                     state="readonly", width=27)\
            .grid(row=2, column=1, columnspan=3, sticky="we", pady=2)

        ttk.Label(self, text="Bilagsfil:").grid(row=3, column=0, sticky="w")
        self.bilag = tk.StringVar(value="")
        ttk.Entry(self, textvariable=self.bilag, width=28)\
            .grid(row=3, column=1, sticky="we")
        ttk.Button(self, text="Bla …", command=self._velg_fil)\
            .grid(row=3, column=2, sticky="e")

        ttk.Button(self, text="Analyse",
                   command=lambda: self._start("analyse"))\
            .grid(row=4, column=1, sticky="e", padx=2, pady=4)
        ttk.Button(self, text="Bilagsuttrekk",
                   command=lambda: self._start("uttrekk"))\
            .grid(row=4, column=2, sticky="w", padx=2, pady=4)
        ttk.Button(self, text="Mapping …",
                   command=self._mapping_dialog)\
            .grid(row=4, column=3, sticky="w", padx=2, pady=4)

        self.status = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.status, anchor="w")\
            .grid(row=5, column=0, columnspan=4, sticky="we")

        ent.focus()

    # ---------- søk / oppdater ---------------------------------------
    def _filter(self, *_):
        s = self.sok.get().lower()
        vals = [n for n in list_klientmapper() if s in n.lower()]
        self.cli_cmb["values"] = vals
        if vals:
            self.cli_cmb.current(0); self._oppdater_type()

    def _oppdater_type(self, *_):
        meta = les_meta(ROOT_DIR / self.cli_var.get())
        if t := meta.get("map_type"):
            self.type_var.set(t)
        if f := meta.get("last_file"):
            if Path(f).exists():
                self.bilag.set(f)

    def _velg_fil(self):
        p = filedialog.askopenfilename(
            title="Velg bilagsfil",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if p:
            self.bilag.set(p)

    # ---------- mapping-dialog ---------------------------------------
    def _mapping_dialog(self):
        if not all((self.cli_var.get(), self.type_var.get(), self.bilag.get())):
            messagebox.showerror("Feil", "Velg klient, kilde og fil først"); return
        src = Path(self.bilag.get())
        if not src.exists():
            messagebox.showerror("Feil", "Bilagsfil mangler"); return

        df = (pd.read_excel(src, engine="openpyxl", nrows=1)
              if src.suffix.lower() in (".xlsx", ".xls") else _csv_header(src))

        cli_root = ROOT_DIR / self.cli_var.get()
        map_path = cli_root / "_mapping.json"
        defaults = json.loads(map_path.read_text("utf-8")) if map_path.exists() else None

        fv = FeltVelger(self, df, defaults)
        mapping = fv.mapping()
        try:
            fv.destroy()
        except tk.TclError:
            pass

        if mapping is None:
            return
        try:
            cli_root.mkdir(exist_ok=True)
            map_path.write_text(json.dumps(mapping, indent=2,
                                           ensure_ascii=False), "utf-8")
            self.status.set(f"Mapping lagret → {map_path}")
            logger.info("Mapping lagret: %s", map_path)
        except Exception as e:
            logger.exception("Kunne ikke lagre mapping")
            messagebox.showerror("Feil", f"Kunne ikke lagre mapping:\n{e}", parent=self)

    # ---------- start analyse / uttrekk ------------------------------
    def _start(self, modus: str):
        if not all((self.cli_var.get(), self.type_var.get(), self.bilag.get())):
            messagebox.showerror("Feil", "Fyll ut alle felt"); return
        src = Path(self.bilag.get())
        if not src.exists():
            messagebox.showerror("Feil", "Bilagsfil mangler"); return
        if src.suffix.lower() not in (".xlsx", ".xls", ".csv"):
            messagebox.showerror("Feil", "Kun .xlsx / .xls / .csv"); return

        lagre_sist_klient(self.cli_var.get())
        cli = ROOT_DIR / self.cli_var.get(); cli.mkdir(exist_ok=True)
        skriv_meta(cli, {"last_file": str(src),
                         "map_type": self.type_var.get(),
                         "mapping": None})

        try:
            proc = subprocess.Popen([sys.executable, str(BILAG_GUI), str(src)],
                                    shell=False)
            self.status.set("Starter Bilagsuttrekk …")
            self.after(1200, self._check_child, proc)
        except Exception as exc:
            logger.exception("Feil ved start av bilag_gui")
            messagebox.showerror("Feil", str(exc), parent=self)
            self.status.set("Feil: kunne ikke starte Bilagsuttrekk")

    def _check_child(self, proc: subprocess.Popen):
        if proc.poll() is not None and proc.returncode != 0:
            messagebox.showerror("Feil",
                f"Bilags-GUI avsluttet med kode {proc.returncode}", parent=self)
            self.status.set("Bilagsuttrekk feilet")
        else:
            self.status.set("Bilagsuttrekk åpnet")

# ---------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    KlientVelger().mainloop()

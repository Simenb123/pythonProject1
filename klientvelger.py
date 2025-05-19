# ── klientvelger.py  (m/ «husk siste klient») ─────────────────────────────
from __future__ import annotations
import json, subprocess, sys, tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

# ▄▄▄  KONFIG  ▄▄▄
ROOT_DIR   = Path(r"C:\Users\ib91\Desktop\Prosjekt\Klienter")
LAST_FILE  = ROOT_DIR / "_last_client.txt"          # <─ NY
BILAG_GUI  = Path(__file__).with_name("bilag_gui_tk.py")
LOGISKE_FELT = ("konto", "beløp", "dato")
META_NAME = ".klient_meta.json"

# ▄▄▄  hjelpere  ▄▄▄
def list_klientmapper() -> list[str]:
    return sorted(p.name for p in ROOT_DIR.iterdir() if p.is_dir())

def les_meta(cli_dir: Path) -> dict:
    p = cli_dir / META_NAME
    if not p.exists(): return {}
    try:             return json.loads(p.read_text(encoding="utf-8"))
    except Exception:return {}

def skriv_meta(cli_dir: Path, meta: dict):
    (cli_dir / META_NAME).write_text(json.dumps(meta, indent=2), encoding="utf-8")

# —— NYTT: lagre / lese sist brukte klient ————————————————
def lagre_sist_klient(navn: str):
    try: LAST_FILE.write_text(navn, encoding="utf-8")
    except Exception: pass

def hent_sist_klient() -> str|None:
    try:
        txt = LAST_FILE.read_text(encoding="utf-8").strip()
        return txt if txt in list_klientmapper() else None
    except Exception:
        return None
# ————————————————————————————————————————————————

# ▄▄▄  GUI  ▄▄▄
class KlientVelger(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Klientvelger"); self.resizable(False, False)

        # widgets
        ttk.Label(self,text="Søk klient:").grid(row=0,column=0,sticky="w")
        self.sok=tk.StringVar()
        ent=ttk.Entry(self,textvariable=self.sok,width=25)
        ent.grid(row=0,column=1,columnspan=2,sticky="we",pady=2)
        ent.bind("<KeyRelease>", self._filter)

        ttk.Label(self,text="Velg klient:").grid(row=1,column=0,sticky="w")
        self.cli_var=tk.StringVar()
        self.cli_cmb=ttk.Combobox(self,textvariable=self.cli_var,
                                  values=list_klientmapper(),
                                  state="readonly",width=30)
        self.cli_cmb.grid(row=1,column=1,columnspan=2,sticky="we",pady=2)
        self.cli_cmb.bind("<<ComboboxSelected>>", self._oppdater_type)

        # << husk sist brukte >>
        if (last := hent_sist_klient()):
            self.cli_cmb.set(last); self.after(50, self._oppdater_type)

        ttk.Label(self,text="Datakilde:").grid(row=2,column=0,sticky="w")
        self.type_var=tk.StringVar()
        self.type_cmb=ttk.Combobox(self,textvariable=self.type_var,
                                   values=["Hovedbok","Saldobalanse"],
                                   state="readonly",width=27)
        self.type_cmb.grid(row=2,column=1,columnspan=2,sticky="we",pady=2)

        ttk.Label(self,text="Bilagsfil:").grid(row=3,column=0,sticky="w")
        self.bilag=tk.StringVar()
        ttk.Entry(self,textvariable=self.bilag,width=28)\
            .grid(row=3,column=1,sticky="we")
        ttk.Button(self,text="Bla …",command=self._velg_fil)\
            .grid(row=3,column=2,sticky="e")

        ttk.Button(self,text="Analyse",
                   command=lambda:self._start("analyse"))\
            .grid(row=4,column=1,sticky="e",padx=2,pady=4)
        ttk.Button(self,text="Bilagsuttrekk",
                   command=lambda:self._start("uttrekk"))\
            .grid(row=4,column=2,sticky="w",padx=2,pady=4)

        ent.focus()

    # ── callbacks ─────────────────────────────────────────────────
    def _filter(self,*_):
        søk=self.sok.get().lower()
        vals=[n for n in list_klientmapper() if søk in n.lower()]
        self.cli_cmb["values"]=vals
        if vals:
            self.cli_cmb.current(0); self._oppdater_type()

    def _oppdater_type(self,*_):
        cli_dir=ROOT_DIR / self.cli_var.get()
        meta=les_meta(cli_dir)
        if t:=meta.get("map_type"): self.type_var.set(t)
        if f:=meta.get("last_file"):
            if Path(f).exists(): self.bilag.set(f)

    def _velg_fil(self):
        p=filedialog.askopenfilename(title="Velg bilagsfil",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv")])
        if p: self.bilag.set(p)

    # ── start analyse/uttrekk ─────────────────────────────────────
    def _start(self, modus:str):
        if not all((self.cli_var.get(), self.type_var.get(), self.bilag.get())):
            messagebox.showerror("Feil","Fyll ut alle felt"); return
        src=Path(self.bilag.get())
        if not src.exists():
            messagebox.showerror("Feil","Bilagsfil mangler"); return
        if src.suffix.lower() not in (".xlsx",".xls",".csv"):
            messagebox.showerror("Feil","Kun .xlsx / .xls / .csv"); return

        # lagre "sist brukt klient"
        lagre_sist_klient(self.cli_var.get())

        cli_dir=ROOT_DIR / self.cli_var.get(); cli_dir.mkdir(exist_ok=True)
        meta={"last_file":str(src),
              "map_type": self.type_var.get(),
              "mapping":  None}
        skriv_meta(cli_dir, meta)

        try:
            subprocess.Popen([sys.executable, str(BILAG_GUI), str(src)])
        except FileNotFoundError:
            messagebox.showerror("Feil","Kunne ikke starte bilags-GUI")

# ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    KlientVelger().mainloop()

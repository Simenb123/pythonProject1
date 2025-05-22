"""
bilag_gui_tk.py – 2025-05-16
· Leser CSV / Excel
· Bruker _mapping.json automatisk (enten i samme mappe eller i klientroten)
· Viser kolonnedialog bare når nødvendig
· Kaller kjør_bilagsuttrekk() og viser hvor uttrekket lagres
"""
from __future__ import annotations
import json, pandas as pd, tkinter as tk, sys
from pathlib import Path
from tkinter import ttk, messagebox
from import_pipeline import _les_csv, konverter_til_parquet
from utvalg_logikk   import kjør_bilagsuttrekk

REQ = [("konto", "Kontonummer"),
       ("beløp",  "Beløp"),
       ("dato",   "Dato"),
       ("bilagsnr", "Bilagsnummer")]

# ---------- småhjelpere ---------------------------------------------------
def _les_fil(p: Path):
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, engine="openpyxl"), None
    return _les_csv(p)

def _norm(s: str) -> str:
    import re, unicodedata
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^0-9a-z]", "", s.lower())

# ---------- kolonnedialog -------------------------------------------------
class FeltVelger(tk.Toplevel):
    def __init__(self, master, df: pd.DataFrame, defaults: dict[str,str]|None):
        super().__init__(master); self.grab_set()
        self.title("Velg kolonner"); self.resizable(False, False)
        self.req: dict[str,ttk.Combobox] = {}
        self.extra: list[tuple[tk.StringVar, ttk.Combobox]] = []

        for r, (key, vis) in enumerate(REQ):
            ttk.Label(self, text=vis).grid(row=r, column=0, sticky="w")
            cb = ttk.Combobox(self, values=list(df.columns),
                              state="readonly", width=34)
            if defaults and key in defaults:
                cb.set(defaults[key])
            else:
                for c in df.columns:
                    if _norm(key) in _norm(c):
                        cb.set(c); break
            cb.grid(row=r, column=1, padx=3, pady=1)
            self.req[key] = cb

        self.ex = ttk.Frame(self); self.ex.grid(row=len(REQ), column=0, columnspan=2)
        ttk.Button(self.ex, text="+ Legg til kolonne",
                   command=lambda: self._add_extra()).grid(row=0, column=0, columnspan=2)
        if defaults:
            for k,v in defaults.items():
                if k not in self.req:
                    self._add_extra(prefill=(k,v))

        ttk.Button(self, text="OK", command=self._ok)\
            .grid(row=len(REQ)+1, column=0, columnspan=2, pady=6)
        self.bind("<Escape>", lambda *_: self.destroy())

    def _add_extra(self, *, prefill: tuple[str,str]|None=None):
        r = len(self.extra) + 1
        key_v = tk.StringVar(value=prefill[0] if prefill else "")
        val_v = tk.StringVar(value=prefill[1] if prefill else "")
        ttk.Entry(self.ex, textvariable=key_v, width=15)\
            .grid(row=r, column=0, padx=2, pady=1)
        cb = ttk.Combobox(self.ex, textvariable=val_v,
                          values=list(df.columns), state="readonly", width=22)
        cb.grid(row=r, column=1, padx=2, pady=1)
        self.extra.append((key_v, cb))

    def _ok(self):
        miss=[vis for key,vis in REQ if not self.req[key].get()]
        if miss:
            messagebox.showerror("Mangler", ", ".join(miss), parent=self); return
        self.quit(); self.withdraw()

    def mapping(self) -> dict[str,str]|None:
        self.mainloop()
        if not self.winfo_exists():       # lukket med ✕
            return None
        mp ={k:cb.get() for k,cb in self.req.items()}
        for key_v, cb in self.extra:
            k=key_v.get().strip().lower(); v=cb.get()
            if k and v: mp[k]=v
        return mp

# ---------- hoved-GUI -----------------------------------------------------
class App(tk.Tk):
    def __init__(self, fil: Path):
        super().__init__()
        self.title("Bilagsuttrekk"); self.resizable(False, False)
        self.fil = fil; self._enc=None
        ttk.Label(self, text=f"Fil: {fil}").grid(row=0, column=0, columnspan=4, sticky="w")

        self.k_lo=tk.StringVar(value="1000"); self.k_hi=tk.StringVar(value="9999")
        self.b_lo=tk.StringVar(value="0");    self.b_hi=tk.StringVar(value="1000000")
        self.ant  =tk.StringVar(value="25")
        _lbl,_ent=ttk.Label, ttk.Entry
        _lbl(self,text="Kontointervall").grid(row=1,column=0,sticky="w")
        _ent(self,textvariable=self.k_lo,width=7).grid(row=1,column=1)
        _ent(self,textvariable=self.k_hi,width=7).grid(row=1,column=2)
        _lbl(self,text="Beløps-intervall").grid(row=2,column=0,sticky="w")
        _ent(self,textvariable=self.b_lo,width=7).grid(row=2,column=1)
        _ent(self,textvariable=self.b_hi,width=7).grid(row=2,column=2)
        _lbl(self,text="Antall bilag").grid(row=3,column=0,sticky="w")
        _ent(self,textvariable=self.ant,width=7).grid(row=3,column=1)

        self.hit=tk.StringVar(value=" ")
        _lbl(self,textvariable=self.hit,foreground="blue")\
            .grid(row=4,column=0,columnspan=3,sticky="w")
        ttk.Button(self,text="Trekk utvalg",command=self._run)\
            .grid(row=5,column=2,pady=6)

        # start periodic update of hit statistics
        self.hits_job = None
        self.after(300, self._hits)

    # live-counter
    def _save_refs(self, df, m):
        self._k=pd.to_numeric(df[m["konto"]], errors="coerce")
        self._b=pd.to_numeric(df[m["beløp"]]
                              .astype(str).str.replace(r"[^0-9,.-]","",regex=True)
                              .str.replace(",","."), errors="coerce")
    def _hits(self, *_):
        if not hasattr(self, '_k'):
            return
        try:
            lo_k, hi_k = int(self.k_lo.get()), int(self.k_hi.get())
            lo_b, hi_b = float(self.b_lo.get()), float(self.b_hi.get())
            mask = self._k.between(lo_k, hi_k) & self._b.between(lo_b, hi_b)
            cnt = mask.sum()
            tot = self._b[mask].sum()
            avg = tot / cnt if cnt else 0
            self.hit.set(
                f"Linjer: {cnt}   Sum: {tot:,.0f} kr   Snitt: {avg:,.0f} kr".replace(',', ' ')
            )
        except ValueError:
            self.hit.set(" ")
        if self.hits_job is not None:
            try:
                self.after_cancel(self.hits_job)
            except Exception:
                pass
        self.hits_job = self.after(300, self._hits)

    # ----------------------------------------------------------------------
    def _run(self):
        global df
        try: df, self._enc = _les_fil(self.fil)
        except Exception as e:
            messagebox.showerror("Lesefeil", str(e), parent=self); return

        # ① finn mappingfil (samme mappe ➜ klientmappe)
        map_path = self.fil.with_name("_mapping.json")
        if not map_path.exists():
            map_path = self.fil.parent.parent / "_mapping.json"
        defaults = json.loads(map_path.read_text(encoding="utf-8")) if map_path.exists() else None

        # ② bruk dialog bare når nødvendig
        if defaults and all(v in df.columns for v in defaults.values()):
            mapping=defaults
        else:
            mapping = FeltVelger(self, df, defaults).mapping()
            if mapping is None: return
            try: map_path.write_text(json.dumps(mapping, indent=2, ensure_ascii=False), encoding="utf-8")
            except Exception: pass   # skrivemappe kan være readonly

        self._save_refs(df, mapping); self._hits()

        # ③ evt. konverter til parquet én gang
        try: konverter_til_parquet(self.fil, self.fil.parent, mapping, encoding=self._enc)
        except Exception: pass

        try:
            lo_k,hi_k=int(self.k_lo.get()),int(self.k_hi.get())
            lo_b,hi_b=float(self.b_lo.get()),float(self.b_hi.get())
            ant=int(self.ant.get())
            res=kjør_bilagsuttrekk(self.fil,(lo_k,hi_k),(lo_b,hi_b),ant,meta={**mapping,"encoding":self._enc})
            messagebox.showinfo("Ferdig",f"Uttrekk lagret til:\n{res['uttrekk']}",parent=self)
        except Exception as e:
            messagebox.showerror("Feil", str(e), parent=self)
        finally:
            self.destroy()

# ---------- entry-point ---------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv)<2:
        messagebox.showerror("Feil","Bilagsfil mangler"); sys.exit(1)
    App(Path(sys.argv[1])).mainloop()

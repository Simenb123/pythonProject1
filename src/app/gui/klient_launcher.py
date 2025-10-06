# -*- coding: utf-8 -*-
# Launcher + Klienthub (år, paneler for Hovedbok/Saldobalanse, versjoner samlet)
# - Versjoner vises samlet (ingen AO/Interim-splitt i UI)
# - Ny … / Analyse / Bilagsuttrekk / Vis mapping … / Rediger mapping … / Slett …
# - Mapping-badge (✓/–)
# - Analyse starter robust: python -m app.gui.bilag_gui_tk + PYTHONPATH=…/src
from __future__ import annotations

import os, sys, subprocess
from pathlib import Path
import datetime as _dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# sørg for …/src på sys.path
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

# (valgfritt) kalender
try:
    from tkcalendar import Calendar  # type: ignore
    _HAS_TKCAL = True
except Exception:
    _HAS_TKCAL = False

# tema (valgfritt)
try:
    from app.gui.ui_theme import init_style
except Exception:
    def init_style():  # fallback
        pass

# services
from app.services.clients import (
    get_clients_root, set_clients_root, resolve_root_and_client, list_clients,
    load_meta, save_meta, list_years, open_or_create_year, default_year, set_default_year,
    year_paths,
)
from app.services.versioning import (
    list_versions, create_version, set_active_version, get_active_version,
    delete_version,
)
from app.services.io import read_raw
from app.services.mapping import load_mapping, edit_mapping_dialog


# ---------------- Dato-hjelpere ----------------
def _iso2no(iso: str) -> str:
    try:
        d = _dt.datetime.strptime(iso, "%Y-%m-%d").date()
        return d.strftime("%d.%m.%Y")
    except Exception:
        return iso

def _no2iso(no: str) -> str:
    d = _dt.datetime.strptime(no.strip(), "%d.%m.%Y").date()
    return d.strftime("%Y-%m-%d")

def _infer_type_from_period(iso_to: str) -> str:
    """ÅO hvis TIL=31.12, ellers 'interim'."""
    try:
        d = _dt.datetime.strptime(iso_to, "%Y-%m-%d").date()
        return "ao" if (d.month, d.day) == (12, 31) else "interim"
    except Exception:
        return "interim"


# ---------------- Ny versjon-dialog ----------------
class NewVersionDialog(tk.Toplevel):
    def __init__(self, parent, year: int):
        super().__init__(parent)
        self.title("Ny versjon"); self.resizable(False, False)
        self.transient(parent); self.grab_set()
        self.year = int(year)

        today = _dt.date.today()
        fra_default = _dt.date(self.year, 1, 1)
        til_default = today if today.year == self.year else _dt.date(self.year, 12, 31)

        frm = ttk.Frame(self, padding=10); frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text="Periode").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0,6))

        ttk.Label(frm, text="FRA (DD.MM.ÅÅÅÅ):").grid(row=1, column=0, sticky="w")
        self.fra_var = tk.StringVar(value=fra_default.strftime("%d.%m.%Y"))
        ttk.Entry(frm, textvariable=self.fra_var, width=14).grid(row=1, column=1, sticky="w", padx=(4,12))

        ttk.Label(frm, text="TIL (DD.MM.ÅÅÅÅ):").grid(row=1, column=2, sticky="w")
        self.til_var = tk.StringVar(value=til_default.strftime("%d.%m.%Y"))
        ttk.Entry(frm, textvariable=self.til_var, width=14).grid(row=1, column=3, sticky="w", padx=(4,0))

        if _HAS_TKCAL:
            ttk.Button(frm, text="📅", width=3, command=lambda: self._pick(self.fra_var)).grid(row=1, column=1, sticky="e")
            ttk.Button(frm, text="📅", width=3, command=lambda: self._pick(self.til_var)).grid(row=1, column=3, sticky="e")

        btns = ttk.Frame(frm); btns.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8,4))
        ttk.Button(btns, text="01.01 – i dag", command=self._preset_today).grid(row=0, column=0, padx=(0,6))
        ttk.Button(btns, text="Årsoppgjør (01.01 – 31.12)", command=self._preset_ao).grid(row=0, column=1, padx=(0,6))

        act = ttk.Frame(frm); act.grid(row=3, column=0, columnspan=4, sticky="e", pady=(10,0))
        ttk.Button(act, text="OK", command=self._ok).grid(row=0, column=0, padx=4)
        ttk.Button(act, text="Avbryt", command=self._cancel).grid(row=0, column=1)

        self.result: tuple[str, str] | None = None
        self.bind("<Return>", lambda *_: self._ok()); self.bind("<Escape>", lambda *_: self._cancel())

    def _preset_today(self):
        today = _dt.date.today()
        self.fra_var.set(f"01.01.{self.year}")
        self.til_var.set(today.strftime("%d.%m.%Y") if today.year == self.year else f"31.12.{self.year}")

    def _preset_ao(self):
        self.fra_var.set(f"01.01.{self.year}"); self.til_var.set(f"31.12.{self.year}")

    def _pick(self, var: tk.StringVar):
        top = tk.Toplevel(self); top.title("Velg dato"); top.resizable(False, False)
        top.transient(self); top.grab_set()
        try:
            pre = _dt.datetime.strptime(var.get(), "%d.%m.%Y").date()
        except Exception:
            pre = _dt.date(self.year, 1, 1)
        cal = Calendar(top, selectmode="day", year=pre.year, month=pre.month, day=pre.day,
                       locale="nb_NO", date_pattern="dd.mm.yyyy")
        cal.grid(row=0, column=0, padx=8, pady=8)
        ttk.Button(top, text="OK", command=lambda: (var.set(cal.get_date()), top.destroy())).grid(row=1, column=0, pady=(0,8))

    def _ok(self):
        try:
            iso_fra = _no2iso(self.fra_var.get()); iso_til = _no2iso(self.til_var.get())
        except Exception:
            messagebox.showerror("Ugyldig dato", "Bruk format DD.MM.ÅÅÅÅ."); return
        self.result = (iso_fra, iso_til); self.destroy()

    def _cancel(self): self.result = None; self.destroy()


# ---------------- Panel for én kilde (HB/SB) ----------------
class VersionsPanel(ttk.Frame):
    """Viser og håndterer versjoner for én kilde, uten AO/Interim-splitt i UI."""
    def __init__(self, parent, ctx, source: str, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        assert source in {"hovedbok", "saldobalanse"}
        self.ctx = ctx; self.source = source

        ttk.Label(self, text=f"{source.capitalize()} – versjon", font=("", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(6,2), columnspan=6)

        ttk.Label(self, text="Velg versjon:").grid(row=1, column=0, sticky="w")
        self.cmb_var = tk.StringVar(value="")
        self.cmb = ttk.Combobox(self, textvariable=self.cmb_var, state="readonly", width=44)
        self.cmb.grid(row=1, column=1, columnspan=4, sticky="we", pady=(2,2))
        self.cmb.bind("<<ComboboxSelected>>", self._on_select)

        # mapping-badge
        self.map_status = tk.StringVar(value="–")
        self.map_label = ttk.Label(self, textvariable=self.map_status, width=9, anchor="center")
        self.map_label.grid(row=1, column=5, sticky="e", padx=(6,0))

        row2 = ttk.Frame(self); row2.grid(row=2, column=0, columnspan=6, sticky="we", pady=(4,0))
        ttk.Button(row2, text="Ny …", command=self._ny_versjon).grid(row=0, column=0, padx=3)
        ttk.Button(row2, text="Analyse", command=self._analyse).grid(row=0, column=1, padx=3)
        ttk.Button(row2, text="Bilagsuttrekk", command=self._uttrekk).grid(row=0, column=2, padx=3)
        ttk.Button(row2, text="Vis mapping …", command=self._vis_mapping).grid(row=0, column=3, padx=3)
        ttk.Button(row2, text="Rediger mapping …", command=self._mapping).grid(row=0, column=4, padx=3)
        ttk.Button(row2, text="Slett …", command=self._slett).grid(row=0, column=5, padx=3)

        self.grid_columnconfigure(4, weight=1)
        self._id2type: dict[str, str] = {}

    # ——— helpers ———
    def _msg(self, kind: str, title: str, msg: str):
        try:
            parent = self if self.winfo_exists() else None
            {"info": messagebox.showinfo, "warning": messagebox.showwarning}.get(kind, messagebox.showerror)(title, msg, parent=parent)
        except Exception:
            pass

    def _resolve_raw_for(self, vtype: str, vid: str) -> Path | None:
        try:
            y = int(self.ctx.year.get())
        except Exception:
            return None
        yp = year_paths(self.ctx.root_dir, self.ctx.client, y)
        base = yp.versions / self.source / vtype / vid / "raw"
        if base.exists():
            for f in base.iterdir():
                if f.is_file():
                    return f
        return None

    def _load_sample_df(self, vtype: str, vid: str, nrows: int = 200):
        p = self._resolve_raw_for(vtype, vid)
        if not p: return None, None
        try:
            df, _ = read_raw(p)
            return df.head(nrows).copy(), p
        except Exception:
            return None, p

    # ——— UI actions ———
    def refresh(self):
        root = self.ctx.root_dir; client = self.ctx.client; year = int(self.ctx.year.get())
        self._id2type.clear()
        vs_i = list_versions(root, client, year, self.source, "interim")
        vs_a = list_versions(root, client, year, self.source, "ao")
        all_vs = vs_i + vs_a

        items = []
        for v in all_vs:
            self._id2type[v.id] = v.vtype
            items.append(f"{v.id}  ({_iso2no(v.period_from)}–{_iso2no(v.period_to)})")
        self.cmb["values"] = items

        aid_i = get_active_version(self.ctx.meta, year, self.source, "interim")
        aid_a = get_active_version(self.ctx.meta, year, self.source, "ao")
        aid = aid_a or aid_i
        if aid:
            for s in items:
                if s.startswith(aid + "  "):
                    self.cmb_var.set(s); break
        else:
            self.cmb_var.set("")
        self._update_map_status()

    def _current_selection(self):
        sel = self.cmb_var.get()
        if sel:
            vid = sel.split()[0]
            return vid, self._id2type.get(vid)
        year = int(self.ctx.year.get())
        aid_a = get_active_version(self.ctx.meta, year, self.source, "ao")
        aid_i = get_active_version(self.ctx.meta, year, self.source, "interim")
        aid = aid_a or aid_i
        if aid:
            return aid, self._id2type.get(aid) or ("ao" if aid == aid_a else "interim")
        return None, None

    def _set_active(self, vid: str, vtype: str):
        try:
            y = int(self.ctx.year.get())
            set_active_version(self.ctx.meta, y, self.source, vtype, vid)
            save_meta(self.ctx.root_dir, self.ctx.client, self.ctx.meta)
        except Exception:
            pass

    def _on_select(self, *_):
        vid, vtype = self._current_selection()
        if not vid or not vtype: return
        self._set_active(vid, vtype); self._update_map_status()

    def _ny_versjon(self):
        dlg = NewVersionDialog(self, int(self.ctx.year.get())); self.wait_window(dlg)
        if not dlg.result: return
        iso_fra, iso_til = dlg.result
        p = filedialog.askopenfilename(title=f"Velg {self.source.capitalize()}-fil",
                                       filetypes=[("CSV/XLSX", "*.csv *.xlsx *.xls"), ("Alle filer", "*.*")])
        if not p: return
        vtype = _infer_type_from_period(iso_til)
        v = create_version(self.ctx.root_dir, self.ctx.client, int(self.ctx.year.get()),
                           source=self.source, vtype=vtype, period_from=iso_fra, period_to=iso_til,
                           label="", src_file=Path(p), how="copy")
        self._set_active(v.id, vtype); self.refresh(); self._update_map_status()

    def _analyse(self): self._open_gui("analyse")
    def _uttrekk(self): self._open_gui("uttrekk")

    def _vis_mapping(self):
        vid, vtype = self._current_selection()
        if not vid or not vtype:
            self._msg("warning", "Mangler versjon", "Velg eller opprett en versjon først."); return
        y = int(self.ctx.year.get())
        mp = load_mapping(self.ctx.root_dir, self.ctx.client, y, self.source) or {}

        df_sample, _ = self._load_sample_df(vtype, vid, nrows=200)
        cols = df_sample.columns.tolist() if df_sample is not None else []
        mapped_cols = set(mp.values()) if mp else set()
        unmapped = sorted(list(set(cols) - mapped_cols))

        dlg = tk.Toplevel(self); dlg.title(f"Mapping – {self.source.capitalize()} ({y})")
        dlg.resizable(False, False); dlg.transient(self); dlg.grab_set()
        frm = ttk.Frame(dlg, padding=10); frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Standardfelt → Kolonne i fil", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
        tv = ttk.Treeview(frm, show="headings", columns=("std","col"), height=10)
        tv.grid(row=1, column=0, sticky="nsew")
        tv.heading("std", text="Standardfelt"); tv.heading("col", text="Kolonne i fil")
        tv.column("std", width=180); tv.column("col", width=260)
        if mp:
            for std, col in sorted(mp.items()): tv.insert("", "end", values=(std, col))
        else:
            tv.insert("", "end", values=("–", "Ingen mapping lagret"))

        ttk.Label(frm, text="Umappede kolonner (beholdes som de er)", font=("", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(10,6))
        lb = tk.Listbox(frm, height=8, width=52); lb.grid(row=3, column=0, sticky="we")
        if unmapped:
            for c in unmapped: lb.insert(tk.END, c)
        else:
            lb.insert(tk.END, "(ingen – alle kolonner er mappet)")

        ttk.Button(frm, text="OK", command=dlg.destroy).grid(row=4, column=0, sticky="e", pady=(10,0))

    def _mapping(self):
        try:
            vid, vtype = self._current_selection()
            if not vid or not vtype:
                self._msg("warning", "Mangler versjon", "Velg eller opprett en versjon først."); return
            self._set_active(vid, vtype)
            y = int(self.ctx.year.get())
            p = self._resolve_raw_for(vtype, vid)
            if not p:
                self._msg("warning", "Mangler versjon", "Fant ikke råfil for valgt versjon."); return
            df, _ = read_raw(p)
            edit_mapping_dialog(self, self.ctx.root_dir, self.ctx.client, y, self.source, df.head(200))
            self._update_map_status()
        except Exception as exc:
            self._msg("error", "Mapping-feil", f"{type(exc).__name__}: {exc}")

    def _open_gui(self, modus: str):
        vid, vtype = self._current_selection()
        if not vtype:
            self._msg("warning", "Mangler versjon", "Velg eller opprett en versjon først."); return
        self._set_active(vid, vtype)
        self.ctx._start_bilag_for(self.source, modus, vtype)

    def _slett(self):
        vid, vtype = self._current_selection()
        if not vid or not vtype:
            self._msg("warning", "Slett versjon", "Velg en versjon først."); return
        if not messagebox.askyesno("Slett versjon", f"Slette versjonen?\n\n{vid}", parent=self if self.winfo_exists() else None):
            return
        ok = delete_version(self.ctx.root_dir, self.ctx.client, int(self.ctx.year.get()),
                            self.source, vtype, vid, meta=self.ctx.meta)
        if ok:
            save_meta(self.ctx.root_dir, self.ctx.client, self.ctx.meta)
            self.refresh()
            self._msg("info", "Slettet", f"Versjonen ble slettet:\n{vid}")
        else:
            self._msg("warning", "Ikke funnet", "Fant ikke versjonsmappen.")

    def _update_map_status(self):
        try:
            y = int(self.ctx.year.get())
            ok = bool(load_mapping(self.ctx.root_dir, self.ctx.client, y, self.source) or {})
            self.map_status.set("✓ Mappet" if ok else "– Ikke mappet")
            self.map_label.configure(foreground=("green" if ok else "grey"))
        except Exception:
            pass


# ---------------- Klienthub ----------------
class ClientHub(tk.Toplevel):
    def __init__(self, master: tk.Tk, client_name: str):
        super().__init__(master)
        self.title(f"Klienthub – {client_name}"); self.resizable(False, False)
        self.client = client_name

        # *** ROBUST: les clients_root fra parent __dict__ ELLER settings ***
        try:
            self.root_dir = getattr(master, "__dict__", {}).get("clients_root", None)
        except Exception:
            self.root_dir = None
        if not self.root_dir:
            self.root_dir = get_clients_root()
        if not self.root_dir:
            messagebox.showerror("Mangler klient-rot", "Fant ikke klient-rot i settings. Åpne Start‑portalen og velg rotmappe.")
            self.destroy(); return

        self.meta = load_meta(self.root_dir, self.client)

        frm = ttk.Frame(self, padding=10); frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text=f"Klient: {self.client}", font=("", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")

        years = list_years(self.root_dir, self.client)
        start_year = default_year(self.meta, years[-1] if years else _dt.date.today().year)
        ttk.Label(frm, text="Revisjonsår:").grid(row=1, column=0, sticky="w", pady=(8,2))
        self.year = tk.IntVar(value=start_year)
        self.year_cmb = ttk.Combobox(frm, values=years, textvariable=self.year, width=10, state="readonly")
        self.year_cmb.grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Åpne år …", command=self._open_year).grid(row=1, column=2, padx=6, sticky="w")
        self.year_cmb.bind("<<ComboboxSelected>>", lambda *_: self._on_year_change())

        self.hb_panel = VersionsPanel(frm, self, "hovedbok");     self.hb_panel.grid(row=2, column=0, columnspan=3, sticky="we")
        self.sb_panel = VersionsPanel(frm, self, "saldobalanse"); self.sb_panel.grid(row=3, column=0, columnspan=3, sticky="we")
        self._on_year_change()

    def _start_bilag_for(self, source: str, modus: str, vtype: str):
        """Start Bilags-GUI robust (module mode) og med riktig PYTHONPATH)."""
        try:
            y = int(self.year.get())
        except Exception:
            messagebox.showwarning("Ugyldig år", "Kunne ikke lese valgt år.", parent=self)
            return

        env = dict(os.environ)
        env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")

        args = [sys.executable, "-m", "app.gui.bilag_gui_tk",
                "--client", str(self.client),
                "--year", str(y),
                "--source", str(source),
                "--type", str(vtype),
                "--modus", str(modus)]
        try:
            subprocess.Popen(args, shell=False, cwd=str(SRC), env=env)
        except Exception as exc:
            messagebox.showerror(
                "Oppstart feilet",
                "Kunne ikke starte Bilags-GUI.\n\n" + " ".join(args) + "\n\n" + f"{type(exc).__name__}: {exc}",
                parent=self if self.winfo_exists() else None
            )

    def _open_year(self):
        y = simpledialog.askinteger("Åpne nytt år", "Hvilket år vil du åpne/opprette?",
                                    parent=self, minvalue=2000, maxvalue=2100, initialvalue=self.year.get())
        if not y: return
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        years = list_years(self.root_dir, self.client); self.year_cmb["values"] = years
        self.year.set(y); self._on_year_change()
        messagebox.showinfo("År klart", f"Året {y} er opprettet.", parent=self)

    def _on_year_change(self):
        y = int(self.year.get())
        self.meta = open_or_create_year(self.root_dir, self.client, y, meta=self.meta)
        set_default_year(self.meta, y); save_meta(self.root_dir, self.client, self.meta)
        self.hb_panel.refresh(); self.sb_panel.refresh()

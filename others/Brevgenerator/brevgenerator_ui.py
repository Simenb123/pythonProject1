"""
brevgenerator_ui.py
Tkinter-GUI for Brevgenerator

Start GUI:
    python brevgenerator_ui.py

Avhenger av brevgenerator_core.py i samme mappe/prosjekt.
"""
from __future__ import annotations
import os
import re
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile

from others.Brevgenerator.brevgenerator_core import (
    Ansatt, Client,
    DEFAULT_TEMPLATES_DIR, DEFAULT_PARTNERS_XLSX, DEFAULT_CLIENTS_XLSX,
    load_partners_from_excel, load_clients_from_excel,
    list_docx_files, ensure_sample_template, collect_placeholders,
    render_template, load_config, save_config, app_dir, to_norwegian_date,
    filter_employee_names, filter_client_names,
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Brevgenerator – Word-maler")
        self.geometry("900x660")
        self.minsize(860, 620)

        # Konfig/standardstier
        self.cfg: dict[str, str] = load_config()
        tpl_default = DEFAULT_TEMPLATES_DIR if os.path.isdir(DEFAULT_TEMPLATES_DIR) else app_dir()
        out_default = os.path.join(tpl_default, "Ut")
        partners_default = DEFAULT_PARTNERS_XLSX if os.path.isfile(DEFAULT_PARTNERS_XLSX) else ""
        clients_default = DEFAULT_CLIENTS_XLSX if os.path.isfile(DEFAULT_CLIENTS_XLSX) else ""

        # State
        self.mal_mappe = tk.StringVar(value=self.cfg.get("templates_dir") or tpl_default)
        self.ut_mappe = tk.StringVar(value=self.cfg.get("out_dir") or out_default)
        self.partners_xlsx = tk.StringVar(value=self.cfg.get("partners_xlsx") or partners_default)
        self.clients_xlsx = tk.StringVar(value=self.cfg.get("clients_xlsx") or clients_default)
        self.valgt_mal = tk.StringVar()

        self.klientnavn = tk.StringVar()
        self.klientnr = tk.StringVar()
        self.klientorgnr = tk.StringVar()
        # Viktig: alltid blank default og redigerbar
        self.klient_stilling = tk.StringVar(value="")
        self.sted = tk.StringVar(value="Sandvika")
        self.dato = tk.StringVar(value=to_norwegian_date(dt.date.today()))
        self.open_after = tk.BooleanVar(value=True)

        # Forslag i rullgardin – feltet er redigerbart og starter blankt
        self.client_roles = ["Daglig leder", "Styrets leder", "Representant for selskapet"]

        # Data
        self.ansatte: list[Ansatt] = []
        self.clients: list[Client] = []
        self._all_employee_names: list[str] = []
        self._all_client_names: list[str] = []

        # Build UI
        self._build_widgets()

        # Load data
        self._load_partners()
        self._load_clients()
        self._refresh_maler()

        # Auto-refresh maler når mappe endres (debounce)
        self._refresh_job = None
        try:
            self.mal_mappe.trace_add("write", self._on_templates_dir_change)
        except Exception:
            pass

    # ---------------------- UI ---------------------- #
    def _build_widgets(self) -> None:
        pad = {"padx": 10, "pady": 8}

        # Ansatt
        frm_ansatt = ttk.LabelFrame(self, text="Ansatt")
        frm_ansatt.pack(fill="x", **pad)

        ttk.Label(frm_ansatt, text="Velg ansatt:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        # Redigerbar combobox for substring-søk
        self.ansatt_combo = ttk.Combobox(frm_ansatt, values=self._all_employee_names, state="normal")
        self.ansatt_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        self.ansatt_combo.bind("<<ComboboxSelected>>", self._on_ansatt_change)
        self.ansatt_combo.bind("<KeyRelease>", self._on_employee_typed)
        frm_ansatt.columnconfigure(1, weight=1)

        self.lbl_epost = ttk.Label(frm_ansatt, text="E-post: -")
        self.lbl_epost.grid(row=1, column=0, columnspan=2, sticky="w", padx=8)
        self.lbl_tlf = ttk.Label(frm_ansatt, text="Telefon: -")
        self.lbl_tlf.grid(row=2, column=0, columnspan=2, sticky="w", padx=8)
        self.lbl_stilling = ttk.Label(frm_ansatt, text="Stilling: -")
        self.lbl_stilling.grid(row=3, column=0, columnspan=2, sticky="w", padx=8)

        # Mal & felter (ryddet – bare det viktige er synlig)
        frm_mal = ttk.LabelFrame(self, text="Mal og felter")
        frm_mal.pack(fill="x", **pad)

        # 1) Synlig viktig: Velg mal
        ttk.Label(frm_mal, text="Velg mal (.docx):").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.mal_combo = ttk.Combobox(frm_mal, textvariable=self.valgt_mal, state="readonly")
        self.mal_combo.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(frm_mal, text="Oppdater", command=self._refresh_maler).grid(row=0, column=2, padx=8)
        ttk.Button(frm_mal, text="Vis plassholdere", command=self._show_placeholders).grid(row=0, column=3, padx=8)

        # 2) Avanserte innstillinger (skjult som default)
        self._adv_shown = False
        self._adv_btn = ttk.Button(frm_mal, text="Vis avanserte ▸", command=self._toggle_advanced)
        self._adv_btn.grid(row=1, column=0, sticky="w", padx=8, pady=(0, 4))

        self._adv_frame = ttk.Frame(frm_mal)
        # Innhold i avansert
        # Mappe med maler
        ttk.Label(self._adv_frame, text="Mappe med maler:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(self._adv_frame, textvariable=self.mal_mappe).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(self._adv_frame, text="Bla gjennom…", command=self._choose_folder).grid(row=0, column=2, padx=8)

        # Partnerliste
        ttk.Label(self._adv_frame, text="Partnerliste (.xlsx):").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(self._adv_frame, textvariable=self.partners_xlsx).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(self._adv_frame, text="Bla gjennom…", command=self._choose_excel_partners).grid(row=1, column=2, padx=8)
        ttk.Button(self._adv_frame, text="Last", command=self._load_partners).grid(row=1, column=3, padx=8)

        # Klientliste
        ttk.Label(self._adv_frame, text="Klientliste (.xlsx):").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(self._adv_frame, textvariable=self.clients_xlsx).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(self._adv_frame, text="Bla gjennom…", command=self._choose_excel_clients).grid(row=2, column=2, padx=8)
        ttk.Button(self._adv_frame, text="Last", command=self._load_clients).grid(row=2, column=3, padx=8)

        # Grid-innstilling for avansert
        for i in range(0, 3):
            self._adv_frame.columnconfigure(i, weight=1)

        # 3) Klient-felter – alltid synlige
        base_row = 2  # raden etter "Vis avanserte"
        ttk.Label(frm_mal, text="Klientnavn:").grid(row=base_row, column=0, sticky="w", padx=8, pady=6)
        self.cb_klient = ttk.Combobox(frm_mal, textvariable=self.klientnavn, values=self._all_client_names, state="normal")
        self.cb_klient.grid(row=base_row, column=1, sticky="ew", padx=8)
        self.cb_klient.bind("<<ComboboxSelected>>", self._on_client_change)
        self.cb_klient.bind("<KeyRelease>", self._on_client_typed)
        ttk.Button(frm_mal, text="Søk…", command=self._open_client_picker).grid(row=base_row, column=2, padx=8)

        ttk.Label(frm_mal, text="Klientnr:").grid(row=base_row + 1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm_mal, textvariable=self.klientnr).grid(row=base_row + 1, column=1, sticky="ew", padx=8)

        ttk.Label(frm_mal, text="Org.nr:").grid(row=base_row + 2, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm_mal, textvariable=self.klientorgnr).grid(row=base_row + 2, column=1, sticky="ew", padx=8)

        ttk.Label(frm_mal, text="Klient-stilling:").grid(row=base_row + 3, column=0, sticky="w", padx=8, pady=6)
        # Redigerbar combobox – starter blank
        self.cb_klient_stilling = ttk.Combobox(
            frm_mal,
            textvariable=self.klient_stilling,
            values=self.client_roles,
            state="normal",
        )
        self.cb_klient_stilling.grid(row=base_row + 3, column=1, sticky="ew", padx=8)

        ttk.Label(frm_mal, text="Sted:").grid(row=base_row + 4, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm_mal, textvariable=self.sted).grid(row=base_row + 4, column=1, sticky="ew", padx=8)

        ttk.Label(frm_mal, text="Dato (dd.mm.åååå):").grid(row=base_row + 5, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm_mal, textvariable=self.dato).grid(row=base_row + 5, column=1, sticky="ew", padx=8)

        for i in range(0, 4):
            frm_mal.columnconfigure(i, weight=1)
        frm_mal.columnconfigure(1, weight=3)

        # Lagring
        frm_out = ttk.LabelFrame(self, text="Lagring")
        frm_out.pack(fill="x", **pad)
        ttk.Label(frm_out, text="Ut-mappe:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm_out, textvariable=self.ut_mappe).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(frm_out, text="Velg…", command=self._choose_out_folder).grid(row=0, column=2, padx=8)
        ttk.Checkbutton(frm_out, text="Åpne fil etter generering", variable=self.open_after).grid(row=1, column=1, sticky="w", padx=8)
        frm_out.columnconfigure(1, weight=1)

        # Knapper
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", **pad)
        ttk.Button(frm_btn, text="Generer dokument", command=self._generate).pack(side="right")

    # ------- Avansert-panel ------- #
    def _toggle_advanced(self) -> None:
        if self._adv_shown:
            self._adv_frame.grid_remove()
            self._adv_btn.configure(text="Vis avanserte ▸")
            self._adv_shown = False
        else:
            # Legg panelet inn under raden med knappen
            self._adv_frame.grid(row=2, column=0, columnspan=4, sticky="ew", padx=0, pady=(0, 6))
            self._adv_btn.configure(text="Skjul avanserte ▾")
            self._adv_shown = True

    # ------------------- Datakilder ------------------- #
    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.mal_mappe.get() or app_dir(), title="Velg mappe med Word-maler")
        if folder:
            self.mal_mappe.set(folder)
            self._refresh_maler_and_save()

    def _choose_out_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.ut_mappe.get() or app_dir(), title="Velg ut-mappe")
        if folder:
            self.ut_mappe.set(folder)
            self._save_cfg()

    def _choose_excel_partners(self) -> None:
        path = filedialog.askopenfilename(initialdir=self.partners_xlsx.get() or app_dir(), title="Velg Partner.xlsx",
                                          filetypes=[["Excel (*.xlsx)", "*.xlsx"], ["Alle", "*"]])
        if path:
            self.partners_xlsx.set(path)
            self._load_partners()
            self._save_cfg()

    def _choose_excel_clients(self) -> None:
        path = filedialog.askopenfilename(initialdir=self.clients_xlsx.get() or app_dir(), title="Velg BHL AS klienter.xlsx",
                                          filetypes=[["Excel (*.xlsx)", "*.xlsx"], ["Alle", "*"]])
        if path:
            self.clients_xlsx.set(path)
            self._load_clients()
            self._save_cfg()

    def _load_partners(self) -> None:
        path = self.partners_xlsx.get().strip() or DEFAULT_PARTNERS_XLSX
        if path and os.path.isfile(path):
            partners = load_partners_from_excel(path)
            if partners:
                self.ansatte = partners
                self._all_employee_names = [a.navn for a in partners]
                self.ansatt_combo["values"] = self._all_employee_names
                # Velg sist brukte hvis mulig
                last = self.cfg.get("last_employee")
                if last and last in self._all_employee_names:
                    self.ansatt_combo.set(last)
                elif self._all_employee_names:
                    self.ansatt_combo.set(self._all_employee_names[0])
                self._on_ansatt_change()

    def _load_clients(self) -> None:
        path = self.clients_xlsx.get().strip() or DEFAULT_CLIENTS_XLSX
        if path and os.path.isfile(path):
            clients = load_clients_from_excel(path)
            if clients:
                self.clients = clients
                self._all_client_names = [c.navn for c in clients]
                self.cb_klient["values"] = self._all_client_names

    # ------------------- Filtering & velging ------------------- #
    def _on_employee_typed(self, event=None) -> None:
        q = (self.ansatt_combo.get() or "").strip()
        filtered = filter_employee_names(self.ansatte, q)
        self.ansatt_combo["values"] = filtered if filtered else self._all_employee_names
        if not q:
            return
        # eksakt/unik
        low = q.lower()
        for a in self.ansatte:
            if a.navn.lower() == low:
                self.ansatt_combo.set(a.navn)
                self._on_ansatt_change()
                return
        if len(filtered) == 1:
            self.ansatt_combo.set(filtered[0])
            self._on_ansatt_change()

    def _on_ansatt_change(self, event=None) -> None:
        name = self.ansatt_combo.get().strip()
        if not name:
            return
        a = next((x for x in self.ansatte if x.navn == name), None)
        if not a:
            return
        self.lbl_epost.configure(text=f"E-post: {a.epost or '-'}")
        self.lbl_tlf.configure(text=f"Telefon: {a.telefon or '-'}")
        self.lbl_stilling.configure(text=f"Stilling: {a.stilling or '-'}")
        self._save_cfg()

    def _apply_client_selection(self, name: str) -> None:
        """Sett valgt klient i alle felter: navn, nr, orgnr."""
        c = next((x for x in self.clients if x.navn == name), None)
        if not c:
            return
        self.klientnavn.set(c.navn)
        self.cb_klient.set(c.navn)
        self.klientnr.set(c.nr or "")
        # orgnr: vis kun sifre, blank hvis mangler
        org_clean = "".join(ch for ch in str(c.orgnr or "").strip() if ch.isdigit())
        self.klientorgnr.set(org_clean)

    def _on_client_typed(self, event=None) -> None:
        q = (self.cb_klient.get() or "").strip()
        filtered = filter_client_names(self.clients, q)
        self.cb_klient["values"] = filtered if filtered else self._all_client_names
        if not q:
            return
        low = q.lower()
        # eksakt
        for c in self.clients:
            if c.navn.lower() == low:
                self._apply_client_selection(c.navn)
                return
        # unik kandidat
        if len(filtered) == 1:
            self._apply_client_selection(filtered[0])

    def _on_client_change(self, event=None) -> None:
        name = (self.klientnavn.get() or self.cb_klient.get() or "").strip()
        if name:
            self._apply_client_selection(name)

    def _open_client_picker(self) -> None:
        """Dialog: søkefelt + liste for å velge klient."""
        if not self.clients:
            messagebox.showinfo("Klienter", "Ingen klientliste lastet.")
            return
        top = tk.Toplevel(self)
        top.title("Søk klient")
        top.geometry("520x500")

        frm = ttk.Frame(top)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frm, text="Søk:").pack(anchor="w")
        qvar = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=qvar)
        ent.pack(fill="x", pady=(0, 8))

        lb = tk.Listbox(frm)
        lb.pack(fill="both", expand=True)

        def refresh():
            items = filter_client_names(self.clients, qvar.get())
            lb.delete(0, "end")
            for n in items:
                lb.insert("end", n)

        def accept():
            try:
                sel = lb.get(lb.curselection())
            except Exception:
                sel = qvar.get().strip()
            if sel:
                self._apply_client_selection(sel)
            top.destroy()

        def on_key(_=None):
            refresh()
            if lb.size() > 0:
                lb.selection_clear(0, "end")
                lb.selection_set(0)

        def on_enter(_=None):
            accept()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="OK", command=accept).pack(side="right")
        ttk.Button(btns, text="Avbryt", command=top.destroy).pack(side="right", padx=(0, 8))

        ent.bind("<KeyRelease>", on_key)
        ent.bind("<Return>", on_enter)
        lb.bind("<Double-Button-1>", lambda e: accept())

        # Pre-fill fra det du allerede har skrevet
        qvar.set(self.cb_klient.get().strip())
        refresh()
        ent.focus_set()

    # ------------------- Maler & generering ------------------- #
    def _on_templates_dir_change(self, *args) -> None:
        if getattr(self, "_refresh_job", None):
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                pass
        self._refresh_job = self.after(600, self._refresh_maler_and_save)

    def _refresh_maler_and_save(self) -> None:
        self._refresh_maler()
        self._save_cfg()

    def _refresh_maler(self) -> None:
        folder = self.mal_mappe.get().strip()
        files = list_docx_files(folder)
        if not files:
            ensure_sample_template(folder)
            files = list_docx_files(folder)
        self.mal_combo["values"] = files
        if files:
            if self.valgt_mal.get() in files:
                self.mal_combo.set(self.valgt_mal.get())
            else:
                self.mal_combo.current(0)
        else:
            self.valgt_mal.set("")

    def _show_placeholders(self) -> None:
        path = self._selected_template_path()
        if not path or not os.path.isfile(path):
            messagebox.showinfo("Ingen mal", "Velg en .docx-mal først.")
            return
        vars_list = collect_placeholders(path)
        if not vars_list:
            messagebox.showinfo("Plassholdere", "Fant ingen {{VAR}}-plassholdere i valgt mal.")
            return
        top = tk.Toplevel(self)
        top.title("Plassholdere i mal")
        top.geometry("360x320")
        lb = tk.Listbox(top)
        lb.pack(fill="both", expand=True, padx=10, pady=10)
        for v in vars_list:
            lb.insert("end", f"{{{{{v}}}}}")

    def _selected_template_path(self) -> str | None:
        if not self.valgt_mal.get():
            return None
        return os.path.join(self.mal_mappe.get(), self.valgt_mal.get())

    @staticmethod
    def _scan_unresolved_placeholders(docx_path: str) -> list[str]:
        """
        Skann ferdig .docx for {{PLACEHOLDER}} som fortsatt står igjen.
        Matcher også tokens som er splittet over XML-tags (w:t-runs).
        """
        leftovers: set[str] = set()
        pat = re.compile(r"\{\{(?:\s|<[^>]+>)*([A-Za-z0-9_]+)(?:\s|<[^>]+>)*\}\}", flags=re.DOTALL)
        try:
            with zipfile.ZipFile(docx_path, "r") as zf:
                for name in zf.namelist():
                    if not (name.startswith("word/") and name.endswith(".xml")):
                        continue
                    xml = zf.read(name).decode("utf-8", "ignore")
                    for m in pat.findall(xml):
                        leftovers.add(m)
        except Exception:
            pass
        return sorted(leftovers)

    def _generate(self) -> None:
        try:
            tmpl_path = self._selected_template_path()
            if not tmpl_path or not os.path.isfile(tmpl_path):
                messagebox.showerror("Mangler mal", "Velg en gyldig .docx-mal først.")
                return
            if not os.path.isdir(self.ut_mappe.get()):
                os.makedirs(self.ut_mappe.get(), exist_ok=True)

            # Ansatt
            name = self.ansatt_combo.get().strip()
            a = next((x for x in self.ansatte if x.navn == name), None)
            if not a:
                messagebox.showerror("Ansatt", "Velg en ansatt fra listen.")
                return

            dato_txt = (self.dato.get() or "").strip() or to_norwegian_date(dt.date.today())
            try:
                parsed = dt.datetime.strptime(dato_txt, "%d.%m.%Y")
                dato_txt = to_norwegian_date(parsed)
            except Exception:
                pass

            context = {
                "PARTNER_NAVN": a.navn,
                "PARTNER_EPOST": a.epost,
                "PARTNER_TELEFON": a.telefon,
                "PARTNER_STILLING": a.stilling,
                "KLIENT_NAVN": self.klientnavn.get().strip(),
                "KLIENT_STILLING": self.klient_stilling.get().strip(),
                "KLIENT_NR": self.klientnr.get().strip(),
                "KLIENT_ORGNR": self.klientorgnr.get().strip(),
                "STED": self.sted.get().strip(),
                "DATO": dato_txt,
            }

            base = os.path.splitext(os.path.basename(tmpl_path))[0]
            out_name = f"{base} - {a.navn} - {dt.datetime.now().strftime('%Y%m%d_%H%M')}.docx"
            out_path = os.path.join(self.ut_mappe.get(), out_name)

            render_template(tmpl_path, context, out_path)

            if self.open_after.get():
                try:
                    os.startfile(out_path)  # type: ignore[attr-defined]
                except Exception:
                    pass

            # Etter-sjekk: ubrukte plassholdere?
            leftovers = self._scan_unresolved_placeholders(out_path)
            if leftovers:
                shown = ", ".join(f"{{{{{v}}}}}" for v in leftovers[:10])
                more = "" if len(leftovers) <= 10 else f" (+{len(leftovers)-10} flere)"
                messagebox.showwarning(
                    "Ubrukte plassholdere",
                    "Det ser ut til at noen plassholdere ikke ble fylt ut i dokumentet:\n"
                    f"{shown}{more}\n\n"
                    "Sjekk malen eller verdiene. Tipset: bruk 'Vis plassholdere' for å se alle variabler i malen."
                )

            messagebox.showinfo("Ferdig", f"Dokument generert:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Feil", f"Noe gikk galt under generering:\n{e}")

    # ------------------- Konfig ------------------- #
    def _save_cfg(self) -> None:
        try:
            self.cfg["templates_dir"] = self.mal_mappe.get().strip()
            self.cfg["out_dir"] = self.ut_mappe.get().strip()
            self.cfg["partners_xlsx"] = self.partners_xlsx.get().strip()
            self.cfg["clients_xlsx"] = self.clients_xlsx.get().strip()
            self.cfg["last_employee"] = self.ansatt_combo.get().strip()
            save_config(self.cfg)
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()

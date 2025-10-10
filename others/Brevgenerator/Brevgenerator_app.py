"""
Brevgenerator – app entrypoint

Bruk:
  # GUI (default)
  python Brevgenerator_app.py

  # CLI
  python Brevgenerator_app.py --cli --templates-dir "F:/Dokument/Maler/BHL AS NYE MALER 2025" \
      --template "BHL NY MAL 20.08.2025.docx" --client "Acme AS" \
      --client-number 123 --client-orgnr 999999999 --client-role "Daglig leder" \
      --place Sandvika --date 20.08.2025 --open

Avhenger av filene:
  - brevgenerator_core.py
  - brevgenerator_ui.py  (for GUI)
"""
from __future__ import annotations
import os
import sys
import argparse
import datetime as dt

from brevgenerator_core import (
    Ansatt, Client,
    DEFAULT_TEMPLATES_DIR, DEFAULT_PARTNERS_XLSX, DEFAULT_CLIENTS_XLSX,
    load_partners_from_excel, load_clients_from_excel,
    list_docx_files, ensure_sample_template, render_template,
    load_config, save_config, app_dir, to_norwegian_date,
)


# ---------------- CLI ---------------- #

def _pick_employee(partners: list[Ansatt], name_or_index: str | None) -> Ansatt:
    if not partners:
        raise SystemExit("Fant ingen partnere i Excel-listen.")
    if not name_or_index:
        return partners[0]
    if name_or_index.isdigit():
        idx = int(name_or_index)
        if 0 <= idx < len(partners):
            return partners[idx]
        raise SystemExit(f"Ugyldig ansatt-indeks: {idx}")
    low = name_or_index.strip().lower()
    for a in partners:
        if a.navn.lower() == low:
            return a
    raise SystemExit(f"Fant ikke ansatt ved navn: {name_or_index}")


def _pick_template(templates_dir: str, template_name: str | None) -> str:
    folder = templates_dir or DEFAULT_TEMPLATES_DIR or app_dir()
    files = list_docx_files(folder)
    if not files:
        ensure_sample_template(folder)
        files = list_docx_files(folder)
    if not files:
        raise SystemExit("Fant ingen .docx-maler i mappen.")
    if template_name:
        # aksepter både filnavn og absolutte stier
        path = template_name if os.path.isabs(template_name) else os.path.join(folder, template_name)
        if os.path.isfile(path):
            return path
        raise SystemExit(f"Fant ikke mal: {path}")
    return os.path.join(folder, files[0])


def run_cli(argv: list[str]) -> int:
    p = argparse.ArgumentParser(description="Brevgenerator – CLI")
    p.add_argument("--templates-dir")
    p.add_argument("--template")
    p.add_argument("--out-dir", default=os.path.join(DEFAULT_TEMPLATES_DIR if os.path.isdir(DEFAULT_TEMPLATES_DIR) else app_dir(), "Ut"))
    p.add_argument("--partners-xlsx", default=DEFAULT_PARTNERS_XLSX)
    p.add_argument("--clients-xlsx", default=DEFAULT_CLIENTS_XLSX)

    p.add_argument("--employee", help="Navn eller indeks (0-basert)")
    p.add_argument("--client")
    p.add_argument("--client-number")
    p.add_argument("--client-orgnr")
    p.add_argument("--client-role", default="Daglig leder")
    p.add_argument("--place", default="Sandvika")
    p.add_argument("--date")
    p.add_argument("--open", action="store_true")

    args = p.parse_args(argv)

    # last partnere/klienter
    partners = load_partners_from_excel(args.partners_xlsx or DEFAULT_PARTNERS_XLSX) or []
    employee = _pick_employee(partners, args.employee)

    clients: list[Client] = load_clients_from_excel(args.clients_xlsx or DEFAULT_CLIENTS_XLSX) or []
    c_nr = (args.client_number or "").strip()
    c_org = (args.client_orgnr or "").strip()
    c_name = (args.client or "").strip()
    if c_name and clients:
        for c in clients:
            if c.navn.lower() == c_name.lower():
                c_nr = c_nr or c.nr
                c_org = c_org or c.orgnr
                break

    # mal / utmappe
    tmpl_path = _pick_template(args.templates_dir or DEFAULT_TEMPLATES_DIR, args.template)
    os.makedirs(args.out_dir, exist_ok=True)

    # dato
    dato_txt = args.date or to_norwegian_date(dt.date.today())
    try:
        parsed = dt.datetime.strptime(dato_txt, "%d.%m.%Y")
        dato_txt = to_norwegian_date(parsed)
    except Exception:
        pass

    context = {
        "PARTNER_NAVN": employee.navn,
        "PARTNER_EPOST": employee.epost,
        "PARTNER_TELEFON": employee.telefon,
        "PARTNER_STILLING": employee.stilling,
        "KLIENT_NAVN": c_name,
        "KLIENT_STILLING": (args.client_role or "").strip(),
        "KLIENT_NR": c_nr,
        "KLIENT_ORGNR": c_org,
        "STED": (args.place or "Sandvika").strip(),
        "DATO": dato_txt,
    }

    base = os.path.splitext(os.path.basename(tmpl_path))[0]
    out_name = f"{base} - {employee.navn} - {dt.datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    out_path = os.path.join(args.out_dir, out_name)
    render_template(tmpl_path, context, out_path)

    print(f"OK – generert: {out_path}")
    if args.open:
        try:
            if sys.platform.startswith("win"):
                os.startfile(out_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", out_path])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", out_path])
        except Exception:
            pass
    return 0


# ---------------- main ---------------- #

if __name__ == "__main__":
    # Lettvint modusbryter
    if "--cli" in sys.argv:
        argv = [a for a in sys.argv[1:] if a != "--cli"]
        sys.exit(run_cli(argv))

    # GUI som default
    try:
        from brevgenerator_ui import App
        App().mainloop()
    except ModuleNotFoundError as e:
        print("[INFO] Tkinter/UI ikke tilgjengelig – kjører CLI.")
        sys.exit(run_cli(sys.argv[1:]))

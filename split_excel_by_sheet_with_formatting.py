# split_excel_by_sheet_with_formatting.py
#
# Krever:
#   pip install pywin32
# (Tkinter følger med standard-Python på Windows)

import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client as win32

INVALID_CHARS = r'[<>:"/\\|?*\n\r\t]'

def safe_filename(name: str) -> str:
    """Gjør et Excel-fanenavn trygt som filnavn."""
    return re.sub(INVALID_CHARS, "_", name).strip()

def main() -> None:
    # Filvelger
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Velg en Excel-fil",
        filetypes=[("Excel filer", "*.xlsx *.xls *.xlsm")]
    )
    if not file_path:
        return

    # Lag utdatamappe
    base = os.path.splitext(os.path.basename(file_path))[0]
    out_dir = os.path.join(os.path.dirname(file_path), f"{base}_split")
    os.makedirs(out_dir, exist_ok=True)

    # Start (usynlig) Excel
    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False      # Unngå dialogbokser
    excel.Visible = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(file_path))

        for sheet in wb.Worksheets:
            sheet_name = sheet.Name
            safe_name = safe_filename(sheet_name) or "Sheet"

            # Kopier arket til en NY arbeidsbok
            sheet.Copy()                     # Nå er ActiveWorkbook = ny arbeidsbok
            new_wb = excel.ActiveWorkbook

            # Fjern eventuelle tomme ark som Excel oppretter (som regel ikke nødvendig)
            for ws in list(new_wb.Worksheets):
                if ws.Name != sheet_name:
                    ws.Delete()

            dest_path = os.path.join(out_dir, f"{safe_name}.xlsx")
            new_wb.SaveAs(dest_path, FileFormat=51)   # 51 = xlOpenXMLWorkbook (.xlsx)
            new_wb.Close(SaveChanges=False)
            print(f"Laget: {dest_path}")

        wb.Close(SaveChanges=False)
        messagebox.showinfo("Ferdig!", f"{len(wb.Worksheets)} filer er lagret i\n{out_dir}")

    except Exception as e:
        messagebox.showerror("Feil", f"Noe gikk galt:\n{e}")
        raise
    finally:
        excel.Quit()

if __name__ == "__main__":
    # Kjør bare på Windows med Excel installert
    if sys.platform != "win32":
        print("Dette scriptet krever Windows + Microsoft Excel.")
        sys.exit(1)
    main()

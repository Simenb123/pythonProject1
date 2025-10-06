"""Smoke-test for BilagController.trekk_utvalg()"""
from __future__ import annotations

from pathlib import Path

from others.controllers import BilagController, ClientController, ROOT_DIR

CSV_TXT = "konto;beløp;dato;bilagsnr\n1000;100,0;01.01.2025;A1\n"


def test_trekk_utvalg(tmp_path: Path) -> None:
    # setup dummy klient + fil
    client_dir = ROOT_DIR / "__pytest__"
    client_dir.mkdir(exist_ok=True)
    src = tmp_path / "dummy.csv"
    src.write_text(CSV_TXT, encoding="utf-8")

    bc = BilagController(src, ClientController("__pytest__"))
    out = bc.trekk_utvalg(1, (1000, 1000), [(-10**9, 10**9)])

    assert out.exists()
    # sjekk at Oppsummering-ark finnes
    import openpyxl

    wb = openpyxl.load_workbook(out)
    assert "Oppsummering" in wb.sheetnames

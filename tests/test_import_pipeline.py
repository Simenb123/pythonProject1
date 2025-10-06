"""
Tester for import_pipeline – måler ≥ 90 % dekning.

Kjør:
    pytest --cov=import_pipeline --cov-report=term-missing --cov-fail-under=90
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.app.services.import_pipeline import _STD_COLS, _les_csv, _standardiser


# ────────────────────────────────────────────────────────────────────────────
# 1  _les_csv – varierende encoding + skilletegn
# ────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "encoding, sep",
    [
        ("utf-8", ","),
        ("utf-8", ";"),
        ("utf-16", ";"),
        ("utf-16-le", ","),
    ],
)
def test_les_csv_enc_delim(tmp_path: Path, encoding: str, sep: str) -> None:
    csv_txt = f"konto{sep}beløp{sep}dato\n1234{sep}1 234,56{sep}01.01.2025\n"
    p = tmp_path / "test.csv"
    p.write_bytes(csv_txt.encode(encoding))

    df, enc = _les_csv(p)

    assert enc.lower().startswith(encoding.split("-")[0])
    assert list(df.columns) == ["konto", "beløp", "dato"]
    assert df.iloc[0]["konto"] == 1234


# ────────────────────────────────────────────────────────────────────────────
# 2  _standardiser – kolonne-mapping + rensing
# ────────────────────────────────────────────────────────────────────────────
def test_standardiser_full(tmp_path: Path) -> None:
    raw_vals = ["1 234 567,89", "42,00", "1.234,56"]
    df_raw = pd.DataFrame(
        {
            "Kontonr": [1000, 1000, 2400],
            "Amount": raw_vals,
            "Dato": ["31.12.2024", "01.01.2025", "15.02.2025"],
            "Voucher": ["A1", "A2", "A3"],
        }
    )

    mapping = {
        "konto": "Kontonr",
        "beløp": "Amount",
        "dato": "Dato",
        "bilagsnr": "Voucher",
    }

    df_std = _standardiser(df_raw, mapping)

    assert list(df_std.columns) == list(_STD_COLS)
    assert pd.api.types.is_integer_dtype(df_std["konto"])
    assert pd.api.types.is_float_dtype(df_std["beløp"])
    assert pytest.approx(df_std["beløp"].iloc[0]) == 1_234_567.89


# ────────────────────────────────────────────────────────────────────────────
# 3  Edge-case: pipe-separert CSV
# ────────────────────────────────────────────────────────────────────────────
def test_les_csv_pipe(tmp_path: Path) -> None:
    txt = "konto|beløp|dato\n123|10,0|01.01.2025\n"
    p = tmp_path / "pipe.csv"
    p.write_text(txt, encoding="utf-8")
    df, _ = _les_csv(p)
    assert df.iloc[0]["konto"] == 123

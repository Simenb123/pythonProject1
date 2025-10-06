# -*- coding: utf-8 -*-
# src/app/services/board.py
from __future__ import annotations
from dataclasses import dataclass, asdict
from pathlib import Path
import pandas as pd

from app.services.registry import ensure_client_org_dirs

BOARD_COLS = ["navn","rolle","fra_dato","til_dato","kilde","oppdatert_av","oppdatert_tid"]

@dataclass
class BoardMember:
    navn: str
    rolle: str
    fra_dato: str   # "YYYY-MM-DD"
    til_dato: str   # "" eller "YYYY-MM-DD"
    kilde: str      # "manuell" | "AR" | ...
    oppdatert_av: str
    oppdatert_tid: str

def board_paths(client_dir: Path) -> tuple[Path, Path]:
    ensure_client_org_dirs(client_dir)
    cur = client_dir / "org" / "board" / "board.xlsx"
    hist = client_dir / "org" / "board" / "history" / f"board_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
    return cur, hist

def load_board(client_dir: Path) -> pd.DataFrame:
    cur, _ = board_paths(client_dir)
    if cur.exists():
        try:
            return pd.read_excel(cur, engine="openpyxl")
        except Exception:
            pass
    return pd.DataFrame(columns=BOARD_COLS)

def save_board(client_dir: Path, df: pd.DataFrame):
    cur, hist = board_paths(client_dir)
    with pd.ExcelWriter(cur, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="board")
    with pd.ExcelWriter(hist, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="board")

def upsert_member(client_dir: Path, member: BoardMember):
    df = load_board(client_dir)
    mask = (df["navn"] == member.navn) & (df["rolle"] == member.rolle) & (df["fra_dato"] == member.fra_dato)
    rec = {k: v for k, v in asdict(member).items()}
    if mask.any():
        for k, v in rec.items():
            df.loc[mask, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([rec])], ignore_index=True)
    save_board(client_dir, df)

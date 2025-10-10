# controllers.py – 2025-06-03 (r4 – kontonavn-støtte)
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

from import_pipeline import _les_csv, konverter_til_parquet
from src.app.services.mapping_utils import infer_mapping
from src.app.services.utvalg_logikk import kjør_bilagsuttrekk

logger = logging.getLogger(__name__)

ROOT_DIR   = Path(r"C:\Users\ib91\Desktop\Prosjekt\Klienter")
LAST_FILE  = ROOT_DIR / "_last_client.txt"
META_NAME  = ".klient_meta.json"
MAPPING_FILE = "_mapping.json"


# ═════════════════ ClientController ══════════════════════════════════════
@dataclass
class ClientController:
    name: str

    def __post_init__(self):
        self.dir: Path = ROOT_DIR / self.name
        self.dir.mkdir(exist_ok=True)

    # ---- meta -----------------------------------------------------------
    def _meta_path(self) -> Path:           return self.dir / META_NAME
    def load_meta(self) -> dict[str, Any]:
        try:   return json.loads(self._meta_path().read_text("utf-8"))
        except Exception: return {}
    def save_meta(self, meta: Mapping[str, Any]) -> None:
        try:   self._meta_path().write_text(json.dumps(meta, indent=2), "utf-8")
        except Exception: logger.exception("meta write")

    # ---- mapping --------------------------------------------------------
    def mapping_path(self) -> Path:         return self.dir / MAPPING_FILE
    def load_mapping(self) -> dict[str, str] | None:
        try:   return json.loads(self.mapping_path().read_text("utf-8"))
        except Exception: return None
    def save_mapping(self, mp: Mapping[str, str]) -> None:
        try:   self.mapping_path().write_text(json.dumps(mp, indent=2, ensure_ascii=False), "utf-8")
        except Exception: logger.exception("mapping write")

    # ---- helper ---------------------------------------------------------
    @staticmethod
    def list_clients() -> list[str]:  return sorted(p.name for p in ROOT_DIR.iterdir() if p.is_dir())
    @staticmethod
    def save_last(n: str):            LAST_FILE.write_text(n, "utf-8")
    @staticmethod
    def last_used() -> str | None:
        try:
            t = LAST_FILE.read_text("utf-8").strip()
            return t if t in ClientController.list_clients() else None
        except Exception: return None


# ═════════════════ BilagController ═══════════════════════════════════════
@dataclass
class BilagController:
    src_file: Path
    client:   ClientController
    _cache:   dict[str, Any] = field(default_factory=dict)

    # ---- mapping --------------------------------------------------------
    def _ensure_mapping(self) -> dict[str, str]:
        mp = self.client.load_mapping()
        if mp: return mp
        df = self._read_file()[0]
        mp = infer_mapping(list(df.columns)) or {}
        if not mp: raise ValueError("Mangler mapping – kjør mapping-dialog")
        self.client.save_mapping(mp)
        return mp

    # ---- les råfil ------------------------------------------------------
    def _read_file(self) -> tuple[pd.DataFrame, str | None]:
        if "df" in self._cache: return self._cache["df"], self._cache.get("enc")
        if self.src_file.suffix.lower() in (".xlsx", ".xls"):
            df, enc = pd.read_excel(self.src_file, engine="openpyxl"), None
        else:
            df, enc = _les_csv(self.src_file)
        self._cache.update(df=df, enc=enc)
        return df, enc

    # ---- konto-saldo DF (for selector) ----------------------------------
    def konto_saldo_df(self) -> pd.DataFrame:
        mp = self._ensure_mapping()
        df_raw, _ = self._read_file()
        k_raw, b_raw = mp["konto"], mp["beløp"]

        bel = (
            df_raw[b_raw].astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        df_num = df_raw.copy()
        df_num[b_raw] = pd.to_numeric(bel, errors="coerce")

        saldo = (
            df_num.groupby(k_raw, as_index=False)[b_raw]
            .sum()
            .rename(columns={k_raw: "konto", b_raw: "Saldo"})
            .sort_values("konto")
        )

        # legg til kontonavn dersom mappet
        if "kontonavn" in mp and mp["kontonavn"] in df_raw.columns:
            navn = df_raw[[k_raw, mp["kontonavn"]]].drop_duplicates(k_raw)
            saldo = (
                saldo.merge(navn, left_on="konto", right_on=k_raw, how="left")
                .rename(columns={mp["kontonavn"]: "kontonavn"})
                .drop(columns=[k_raw])
            )
        return saldo

    # ---- trekk utvalg ---------------------------------------------------
    def trekk_utvalg(
        self,
        n_bilag: int,
        konto_rng: tuple[int, int],
        belop_int: list[tuple[float, float]],
    ) -> Path:
        mp  = self._ensure_mapping()
        _, enc = self._read_file()
        pq = konverter_til_parquet(self.src_file, self.client.dir, mp, enc)
        meta = {**mp, "encoding": enc, "std_file": str(pq),
                "belop_intervaller": belop_int}
        res = kjør_bilagsuttrekk(self.src_file, konto_rng, belop_int, n_bilag, meta=meta)
        return res["uttrekk"]

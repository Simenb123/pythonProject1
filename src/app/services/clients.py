from __future__ import annotations

import json, os, shutil, hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

APP_NAME = "KlientApp"
ENV_VAR_CLIENTS_ROOT = "KLIENTAPP_CLIENTS_ROOT"
SETTINGS_FILE = "settings.json"
META_NAME = ".klient_meta.json"

# ---------------- settings ----------------
def _user_config_dir() -> Path:
    if os.name == "nt":
        base = Path(os.getenv("APPDATA", Path.home() / "AppData/Roaming"))
    else:
        base = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    d = base / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d

def _settings_path() -> Path:
    return _user_config_dir() / SETTINGS_FILE

def load_settings() -> dict:
    p = _settings_path()
    if p.exists():
        try:
            return json.loads(p.read_text("utf-8"))
        except Exception:
            return {}
    return {}

def save_settings(d: dict) -> None:
    p = _settings_path()
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)

def get_clients_root(cli_arg: Optional[str] = None) -> Optional[Path]:
    if cli_arg:
        return Path(cli_arg)
    if os.getenv(ENV_VAR_CLIENTS_ROOT):
        return Path(os.getenv(ENV_VAR_CLIENTS_ROOT, ""))
    st = load_settings()
    p = st.get("clients_root")
    return Path(p) if p else None

def set_clients_root(root: Path) -> None:
    st = load_settings()
    st["clients_root"] = str(Path(root))
    save_settings(st)

def resolve_root_and_client(p: Path) -> Tuple[Path, Optional[str]]:
    p = Path(p)
    if (p / "years").exists():
        return p.parent, p.name
    return p, None

# --------------- klient-meta --------------
def list_clients(root: Path) -> list[str]:
    if not root or not root.exists():
        return []
    return sorted([p.name for p in root.iterdir() if p.is_dir()])

def client_dir(root: Path, name: str) -> Path:
    return Path(root) / name

def meta_path(root: Path, name: str) -> Path:
    return client_dir(root, name) / META_NAME

def load_meta(root: Path, name: str) -> dict:
    p = meta_path(root, name)
    if p.exists():
        try:
            return json.loads(p.read_text("utf-8"))
        except Exception:
            return {}
    return {}

def save_meta(root: Path, name: str, d: dict) -> None:
    p = meta_path(root, name)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)

# --------------- per-år modell ------------
def list_years(root: Path, client: str) -> list[int]:
    yroot = client_dir(root, client) / "years"
    if not yroot.exists():
        return []
    out: list[int] = []
    for p in yroot.iterdir():
        if p.is_dir():
            try:
                out.append(int(p.name))
            except ValueError:
                pass
    return sorted(out)

def default_year(meta: dict, fallback: int) -> int:
    try:
        return int(meta.get("default_year", fallback))
    except Exception:
        return fallback

def set_default_year(meta: dict, year: int) -> None:
    meta["default_year"] = int(year)

def get_year_meta(meta: dict, year: int, create: bool = True) -> dict:
    yrs = meta.setdefault("years", {})
    key = str(year)
    if key not in yrs and create:
        yrs[key] = {
            "datakilde": "hovedbok",
            "hovedbok_file": "",
            "saldobalanse_file": "",
            "ui_prefs": {},
        }
    return yrs.get(key, {})

def set_year_file(meta: dict, year: int, source: str, path: str) -> None:
    if source not in {"hovedbok", "saldobalanse"}:
        raise ValueError("source må være 'hovedbok' eller 'saldobalanse'")
    ym = get_year_meta(meta, year, create=True)
    ym[f"{source}_file"] = path
    ym["last_used"] = datetime.now().isoformat(timespec="seconds")

def set_year_datakilde(meta: dict, year: int, source: str) -> None:
    if source not in {"hovedbok", "saldobalanse"}:
        raise ValueError("datakilde må være 'hovedbok' eller 'saldobalanse'")
    ym = get_year_meta(meta, year, create=True)
    ym["datakilde"] = source

# -------------- standardmapper ------------
@dataclass(frozen=True)
class YearPaths:
    base: Path
    data_raw: Path
    data_processed: Path
    mapping: Path
    versions: Path
    logs: Path

def year_paths(root: Path, client: str, year: int) -> YearPaths:
    ybase = client_dir(root, client) / "years" / f"{year}"
    return YearPaths(
        base=ybase,
        data_raw=ybase / "data" / "raw",
        data_processed=ybase / "data" / "processed",
        mapping=ybase / "mapping",
        versions=ybase / "versions",
        logs=ybase / "logs",
    )

def ensure_year_folders(root: Path, client: str, year: int) -> YearPaths:
    yp = year_paths(root, client, year)
    for d in (yp.data_raw, yp.data_processed, yp.mapping, yp.versions, yp.logs):
        d.mkdir(parents=True, exist_ok=True)
    return yp

def mapping_file(root: Path, client: str, year: int, source: str) -> Path:
    if source not in {"hovedbok", "saldobalanse"}:
        raise ValueError("source må være 'hovedbok' eller 'saldobalanse'")
    yp = year_paths(root, client, year)
    yp.mapping.mkdir(parents=True, exist_ok=True)
    return yp.mapping / f"{source}_mapping.json"

def processed_export_path(root: Path, client: str, year: int, stem: str) -> Path:
    yp = year_paths(root, client, year)
    yp.data_processed.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return yp.data_processed / f"{stem}_{ts}"

# -------------- ingest av kilder ----------
def _hash_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

def target_raw_dir(root: Path, client: str, year: int, source: str) -> Path:
    yp = year_paths(root, client, year)
    sub = "Hovedbok" if source == "hovedbok" else "Saldobalanse" if source == "saldobalanse" else "Annet"
    d = yp.data_raw / sub
    d.mkdir(parents=True, exist_ok=True)
    return d

def ingest_source_file(root: Path, client: str, year: int, source: str,
                       picked: Path, how: str = "copy") -> dict:
    picked = Path(picked)
    if how == "external":
        return {"final_path": picked, "local": False, "sha256": None}

    dst_dir = target_raw_dir(root, client, year, source)
    dst = dst_dir / picked.name
    if dst.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = dst_dir / f"{picked.stem}_{ts}{picked.suffix}"

    if how == "link":
        try:
            if os.name == "nt" and picked.drive == dst.drive:
                os.link(str(picked), str(dst))
            else:
                os.link(picked, dst)
        except Exception:
            shutil.copy2(picked, dst)
    else:
        shutil.copy2(picked, dst)

    sha = _hash_file(dst)
    (dst.with_suffix(dst.suffix + ".manifest.json")).write_text(
        json.dumps({"file": dst.name, "source": source, "sha256": sha,
                    "created": datetime.now().isoformat(timespec="seconds"),
                    "origin": str(picked)}, indent=2, ensure_ascii=False),
        "utf-8"
    )
    return {"final_path": dst, "local": True, "sha256": sha}

# -------------- år-init helpers -----------
def open_or_create_year(root: Path, client: str, year: int, meta: dict | None = None) -> dict:
    ensure_year_folders(root, client, year)
    if meta is None:
        meta = load_meta(root, client)
    _ = get_year_meta(meta, year, create=True)
    if "default_year" not in meta:
        meta["default_year"] = int(year)
    save_meta(root, client, meta)
    return meta

def current_year() -> int:
    return datetime.now().year

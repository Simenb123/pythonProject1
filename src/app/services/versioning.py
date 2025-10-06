from __future__ import annotations
import json, os, shutil, hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Literal

from .clients import year_paths, get_year_meta

VersionType = Literal["interim","ao"]
SourceType  = Literal["hovedbok","saldobalanse"]

@dataclass(frozen=True)
class VersionInfo:
    id: str
    source: SourceType
    vtype: VersionType
    period_from: str
    period_to: str
    label: str
    dir: Path
    raw_file: Optional[Path]
    created_at: str

def _hash_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1<<20), b""):
            h.update(chunk)
    return h.hexdigest()

def _versions_root(root: Path, client: str, year: int, source: SourceType, vtype: VersionType) -> Path:
    yp = year_paths(root, client, year)
    d = yp.versions / source / vtype
    d.mkdir(parents=True, exist_ok=True)
    return d

def _safe_label(label: str) -> str:
    s = "".join(ch for ch in (label or "").strip().replace(" ", "_") if ch.isalnum() or ch in "_-").lower()
    return s or "versjon"

def make_version_id(fr: str, to: str, label: str) -> str:
    return f"v{fr}_{to}__{_safe_label(label)}"

def _write_manifest(vdir: Path, info: dict) -> None:
    (vdir / "manifest.json").write_text(json.dumps(info, indent=2, ensure_ascii=False), "utf-8")

def _read_manifest(vdir: Path) -> Optional[dict]:
    mf = (vdir / "manifest.json")
    if not mf.exists(): return None
    try:
        return json.loads(mf.read_text("utf-8"))
    except Exception:
        return None

def list_versions(root: Path, client: str, year: int,
                  source: SourceType, vtype: VersionType) -> list[VersionInfo]:
    base = _versions_root(root, client, year, source, vtype)
    out: list[VersionInfo] = []
    for d in sorted(p for p in base.iterdir() if p.is_dir()):
        mf = _read_manifest(d)
        if not mf: continue
        raw = d / "raw"
        raw_file = None
        if raw.exists():
            for f in raw.iterdir():
                if f.is_file():
                    raw_file = f; break
        out.append(VersionInfo(
            id=mf["id"], source=source, vtype=vtype,
            period_from=mf["period"]["from"], period_to=mf["period"]["to"],
            label=mf.get("label",""), dir=d, raw_file=raw_file,
            created_at=mf.get("created_at",""),
        ))
    return out

def _unique_dir(base: Path) -> Path:
    if not base.exists(): return base
    i = 2
    while True:
        cand = base.with_name(f"{base.name}__{i}")
        if not cand.exists(): return cand
        i += 1

def _unique_file(dst: Path) -> Path:
    if not dst.exists(): return dst
    stem, suf = dst.stem, dst.suffix
    i = 2
    while True:
        cand = dst.with_name(f"{stem}__{i}{suf}")
        if not cand.exists(): return cand
        i += 1

def create_version(root: Path, client: str, year: int, *,
                   source: SourceType, vtype: VersionType,
                   period_from: str, period_to: str,
                   label: str, src_file: Path, how: Literal["copy","link"]="copy") -> VersionInfo:
    base = _versions_root(root, client, year, source, vtype)
    vid  = make_version_id(period_from, period_to, label)
    vdir = _unique_dir(base / vid)
    (vdir / "raw").mkdir(parents=True, exist_ok=True)

    src_file = Path(src_file)
    dst = _unique_file(vdir / "raw" / src_file.name)
    try:
        if how == "link" and os.name == "nt" and src_file.drive == dst.drive:
            os.link(str(src_file), str(dst))
        elif how == "link":
            os.link(src_file, dst)
        else:
            shutil.copy2(src_file, dst)
    except Exception:
        shutil.copy2(src_file, dst)

    actual_id = vdir.name
    info = {
        "id": actual_id,
        "source": source,
        "type": vtype,
        "period": {"from": period_from, "to": period_to},
        "label": label,
        "raw": dst.name,
        "sha256": _hash_file(dst),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "origin": str(src_file)
    }
    _write_manifest(vdir, info)

    return VersionInfo(
        id=actual_id, source=source, vtype=vtype,
        period_from=period_from, period_to=period_to,
        label=label, dir=vdir, raw_file=dst, created_at=info["created_at"]
    )

# ----- aktiv versjon i meta -----

def _versions_meta_node(meta: dict, year: int) -> dict:
    ym = get_year_meta(meta, year, create=True)
    return ym.setdefault("versions", {
        "hovedbok": {"interim": "", "ao": ""},
        "saldobalanse": {"interim": "", "ao": ""}
    })

def set_active_version(meta: dict, year: int, source: SourceType, vtype: VersionType, version_id: str) -> None:
    node = _versions_meta_node(meta, year)
    node[source][vtype] = version_id

def get_active_version(meta: dict, year: int, source: SourceType, vtype: VersionType) -> str:
    node = _versions_meta_node(meta, year)
    return node.get(source, {}).get(vtype, "") or ""

def resolve_active_raw_file(root: Path, client: str, year: int,
                            source: SourceType, vtype: VersionType, meta: dict) -> Optional[Path]:
    vid = get_active_version(meta, year, source, vtype)
    if not vid: return None
    base = _versions_root(root, client, year, source, vtype) / vid / "raw"
    if not base.exists(): return None
    for f in base.iterdir():
        if f.is_file(): return f
    return None

# ----- SLETTING -----

def delete_version(root: Path, client: str, year: int,
                   source: SourceType, vtype: VersionType, version_id: str,
                   meta: Optional[dict] = None) -> bool:
    """
    Sletter en versjon (mappe) trygt. Nullstiller aktiv peker hvis den peker hit.
    Returnerer True hvis slettet.
    """
    base = _versions_root(root, client, year, source, vtype)
    vdir = base / version_id
    if not vdir.exists(): return False

    # fjern hele katalogen
    shutil.rmtree(vdir, ignore_errors=True)

    # nullstill aktiv hvis peker på denne
    if meta is not None:
        if get_active_version(meta, year, source, vtype) == version_id:
            set_active_version(meta, year, source, vtype, "")
    return True

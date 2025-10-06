# -*- coding: utf-8 -*-
# src/app/services/audit.py
from __future__ import annotations
from pathlib import Path
from typing import Any, Dict, Optional
import json, datetime as dt

def _ts() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")

def _append_jsonl(path: Path, rec: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")

def log_global(root: Path, event: str, user: str, payload: Dict[str, Any]):
    p = Path(root) / "_admin" / "audit.jsonl"
    rec = {"ts": _ts(), "event": event, "user": user, "payload": payload}
    _append_jsonl(p, rec)

def log_client(root: Path, client: str, area: str, action: str, user: str,
               before: Optional[Dict[str, Any]] = None, after: Optional[Dict[str, Any]] = None,
               extra: Optional[Dict[str, Any]] = None):
    p = Path(root) / client / "org" / area / "audit.jsonl"
    rec = {"ts": _ts(), "client": client, "area": area, "action": action, "user": user,
           "before": before or {}, "after": after or {}, "extra": extra or {}}
    _append_jsonl(p, rec)

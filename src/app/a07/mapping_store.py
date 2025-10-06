# mapping_store.py
# -*- coding: utf-8 -*-
from __future__ import annotations
import json, os, datetime as dt
from typing import Dict, Any, Optional, Tuple, List

ISO = "%Y-%m-%dT%H:%M:%S"

def _now() -> str:
    return dt.datetime.now().strftime(ISO)

def _ensure_dir(p: str) -> None:
    if p and not os.path.isdir(p):
        os.makedirs(p, exist_ok=True)

def _read_json(path: str, default: Any):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def _write_json(path: str, obj: Any) -> None:
    _ensure_dir(os.path.dirname(path))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)

class MappingStore:
    """
    Lagrer/leser:
      - Klientspesifikke mappinger i klientmappe/<lønn>/a07_mapping_<timestamp>.json
      - Global læring i global_mappe/a07_global_mappings.json
    """
    def __init__(self, global_root: Optional[str]=None, client_root: Optional[str]=None):
        self.global_root = global_root or ""
        self.client_root = client_root or ""

    # --------- klientspesifikt ---------
    def client_lonn_dir(self, client_dir: str) -> str:
        if not client_dir:
            return ""
        d = os.path.join(client_dir, "lønn")
        _ensure_dir(d)
        return d

    def save_client_mapping(self, client_dir: str, payload: Dict[str, Any]) -> str:
        """
        payload = {
          'client_id': str, 'orgnr_list': [..], 'mapping': {'5410':'arbeidsgiveravgift', ...},
          'a07_file': str, 'gl_file': str, 'basis': 'Auto|UB|Endring|Beløp', 'rulebook_source': str
        }
        """
        d = self.client_lonn_dir(client_dir)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(d, f"a07_mapping_{ts}.json")
        payload = dict(payload)
        payload["version"] = 1
        payload["saved_at"] = _now()
        _write_json(path, payload)
        # legg en "latest" for enkel autoload
        _write_json(os.path.join(d, "a07_mapping_latest.json"), payload)
        return path

    def load_client_latest(self, client_dir: str) -> Optional[Dict[str, Any]]:
        d = self.client_lonn_dir(client_dir)
        latest = os.path.join(d, "a07_mapping_latest.json")
        if os.path.exists(latest):
            return _read_json(latest, None)
        # Fallback: sist lagrede fil
        try:
            files = sorted([f for f in os.listdir(d) if f.startswith("a07_mapping_") and f.endswith(".json")])
            if files:
                return _read_json(os.path.join(d, files[-1]), None)
        except Exception:
            pass
        return None

    # --------- global læring ---------
    def global_path(self) -> str:
        if not self.global_root:
            return ""
        _ensure_dir(self.global_root)
        return os.path.join(self.global_root, "a07_global_mappings.json")

    def load_global(self) -> Dict[str, Any]:
        path = self.global_path()
        if not path:
            return {"version": 1, "entries": []}
        return _read_json(path, {"version":1, "entries":[]})

    def save_global(self, data: Dict[str, Any]) -> None:
        path = self.global_path()
        if not path:
            return
        data["version"] = 1
        data["saved_at"] = _now()
        _write_json(path, data)

    def update_global_with_mapping(self, mapping: Dict[str, str], gl_names: Dict[str, str]) -> None:
        """
        Oppdater global fil med (konto -> kode). Aggreger count/last_used.
        """
        obj = self.load_global()
        entries: List[Dict[str, Any]] = list(obj.get("entries", []))

        idx: Dict[Tuple[str,str], int] = {}
        for i, e in enumerate(entries):
            idx[(str(e.get("acc","")), str(e.get("code","")))] = i

        changed = False
        for acc, code in mapping.items():
            key = (str(acc), str(code))
            if key in idx:
                i = idx[key]
                entries[i]["count"] = int(entries[i].get("count", 0)) + 1
                entries[i]["last_used"] = _now()
                entries[i]["name"] = gl_names.get(str(acc), entries[i].get("name",""))
            else:
                entries.append({
                    "acc": str(acc),
                    "name": gl_names.get(str(acc), ""),
                    "code": str(code),
                    "count": 1,
                    "last_used": _now()
                })
            changed = True

        if changed:
            obj["entries"] = entries
            self.save_global(obj)

    def global_suggestions_for_acc(self, acc: str) -> Optional[str]:
        """
        Returnerer mest brukt kode for gitt konto (hvis noen).
        """
        obj = self.load_global()
        best_code, best_count = None, 0
        for e in obj.get("entries", []):
            if str(e.get("acc")) == str(acc):
                c = int(e.get("count", 0))
                if c > best_count:
                    best_code, best_count = str(e.get("code")), c
        return best_code

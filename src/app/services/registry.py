# src/app/services/registry.py
from __future__ import annotations

import os, json, re
from pathlib import Path
from typing import Iterable, Optional, Dict, Any, List

import pandas as pd

APP_NAME = "KlientApp"  # hold samme navn som i app.services.clients
REGISTRY_JSON = "registry.json"
EMPLOYEES_XLSX = "employees.xlsx"
TEAM_FILE = "team.json"  # lagres i klientmappen

# ------------------------- config-sti -------------------------
def _config_dir() -> Path:
    if os.name == "nt":
        base = Path(os.getenv("APPDATA", Path.home() / "AppData/Roaming"))
    else:
        base = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    d = base / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d

def _reg_path() -> Path:
    return _config_dir() / REGISTRY_JSON

def _emp_path() -> Path:
    return _config_dir() / EMPLOYEES_XLSX

# ------------------------- registry I/O ------------------------
def load_registry() -> Dict[str, Any]:
    p = _reg_path()
    if p.exists():
        try:
            d = json.loads(p.read_text("utf-8"))
            d.setdefault("current_email", "")
            d.setdefault("employees", [])  # list[ {name,email,initials} ]
            return d
        except Exception:
            pass
    return {"current_email": "", "employees": []}

def save_registry(d: Dict[str, Any]) -> None:
    p = _reg_path()
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, indent=2, ensure_ascii=False), "utf-8")
    tmp.replace(p)

# ------------------------- current user ------------------------
def _norm_email(s: str) -> str:
    s = (s or "").strip()
    return s if "@" in s else s.lower()

def current_email() -> str:
    return load_registry().get("current_email", "") or ""

def set_current_email(email: str) -> str:
    email = _norm_email(email)
    reg = load_registry()
    reg["current_email"] = email
    save_registry(reg)
    return email

# ------------------------- employees --------------------------
def list_employees_df() -> pd.DataFrame:
    reg = load_registry()
    rows = reg.get("employees", [])
    if not rows and _emp_path().exists():
        try:
            return _read_employees_excel(_emp_path())
        except Exception:
            pass
    return pd.DataFrame(rows)

def upsert_employees(rows: Iterable[Dict[str, Any]]) -> None:
    df = pd.DataFrame(rows)
    if "email" not in df.columns:
        return
    df["email"] = df["email"].astype(str).str.strip().str.lower()
    df = df.drop_duplicates(subset=["email"]).sort_values("email")
    reg = load_registry()
    reg["employees"] = df.to_dict(orient="records")
    save_registry(reg)

def _read_employees_excel(p: Path) -> pd.DataFrame:
    import pandas as pd
    df = pd.read_excel(p, engine="openpyxl")
    low = {c: c for c in df.columns}
    # prøv vanlige navn: Navn/fornavn+etternavn, IN/Initialer, epost/email
    cols = {c.lower(): c for c in df.columns}
    name_col = cols.get("navn") or cols.get("name") or None
    email_col = cols.get("epost") or cols.get("email")
    init_col = cols.get("in") or cols.get("initialer") or cols.get("initials")
    if not email_col:
        raise ValueError("Fant ikke kolonne 'epost' / 'email' i ansattlisten.")
    out = pd.DataFrame()
    out["email"] = df[email_col].astype(str).str.strip().str.lower()
    if name_col:
        out["name"] = df[name_col].astype(str).str.strip()
    else:
        # dra ut navn før <email> hvis tilgjengelig, ellers tomt
        out["name"] = ""
    if init_col:
        out["initials"] = df[init_col].astype(str).str.strip()
    else:
        out["initials"] = out["email"].str.replace(r"@.*$", "", regex=True)
    return out.dropna(subset=["email"])

def import_employees_from_excel(p: Path) -> int:
    df = _read_employees_excel(Path(p))
    upsert_employees(df.to_dict(orient="records"))
    # lag en kopi for å ha «fasit» tilgjengelig
    try: df.to_excel(_emp_path(), index=False)
    except Exception: pass
    return len(df)

# ------------------------- team per klient ---------------------
def team_file(root: Path, client: str) -> Path:
    return Path(root) / client / TEAM_FILE

def load_team(root: Path, client: str) -> Dict[str, Any]:
    p = team_file(root, client)
    if p.exists():
        try:
            return json.loads(p.read_text("utf-8"))
        except Exception:
            pass
    return {"members": []}

def save_team(root: Path, client: str, team: Dict[str, Any]) -> None:
    p = team_file(root, client)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(team, indent=2, ensure_ascii=False), "utf-8")
    tmp.replace(p)

def team_has_user(root: Path, client: str, email: str) -> bool:
    email = (email or "").strip().lower()
    if not email: return False
    t = load_team(root, client)
    for m in t.get("members", []):
        if str(m.get("email", "")).strip().lower() == email:
            return True
    return False

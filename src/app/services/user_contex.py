from __future__ import annotations

import getpass
import hashlib
import json
import os
import platform
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

APP_NAME = "KlientApp"

# ----------------------- identitet -----------------------

@dataclass(frozen=True)
class UserIdentity:
    username: str          # "ib91" eller "ola.nordmann"
    domain: Optional[str]  # "FIRMA" eller None
    display: str           # "FIRMA\\ib91" eller "ib91"
    machine: str           # maskinnavn (HOSTNAME)
    os: str                # "Windows", "Linux", "Darwin"

def _env_first(*keys: str) -> Optional[str]:
    for k in keys:
        v = os.environ.get(k)
        if v:
            return v
    return None

def _whoami_windows() -> Optional[str]:
    try:
        out = subprocess.check_output(["whoami"], text=True, shell=False).strip()
        # typisk "FIRMA\\ib91"
        return out if out else None
    except Exception:
        return None

def current_username() -> str:
    """
    Tverrplattform forsøk på å finne "fornuftig" brukernavn.
    """
    # 1) getpass er robust i konsoll/GUI
    try:
        u = getpass.getuser()
        if u:
            return u
    except Exception:
        pass

    # 2) Miljøvariabler
    u = _env_first("USERNAME", "USER", "LOGNAME")
    if u:
        return u

    # 3) whoami (Windows)
    who = _whoami_windows()
    if who and "\\" in who:
        return who.split("\\", 1)[1]

    # fallback
    return "unknown"

def current_domain() -> Optional[str]:
    """
    Windows domene hvis finnes; ellers None.
    """
    # Miljøvariabler først
    dom = _env_first("USERDOMAIN", "USERDNSDOMAIN")
    if dom:
        return dom

    # whoami (Windows)
    who = _whoami_windows()
    if who and "\\" in who:
        return who.split("\\", 1)[0]

    return None

def current_identity() -> UserIdentity:
    uname = current_username()
    dom = current_domain()
    disp = f"{dom}\\{uname}" if dom else uname
    return UserIdentity(
        username=uname,
        domain=dom,
        display=disp,
        machine=platform.node() or os.environ.get("COMPUTERNAME", ""),
        os=platform.system(),
    )

def user_id() -> str:
    """
    Stabil, anonym ID av format 12 hex-tegn (hash av domain\username).
    Kan brukes som nøkkel i databaser/innstillinger.
    """
    id_source = current_identity().display.lower().encode("utf-8", "ignore")
    return hashlib.sha256(id_source).hexdigest()[:12]

# ----------------------- per-bruker konfig -----------------------

def _user_config_dir() -> Path:
    """
    Brukers *personlige* konfig-område (ikke felles for alle brukere).
    Windows: %APPDATA%\KlientApp
    Linux:   ~/.config/KlientApp
    macOS:   ~/Library/Application Support/KlientApp  (via XDG_CONFIG_HOME fallbacks)
    """
    if os.name == "nt":
        base = Path(os.getenv("APPDATA", Path.home() / "AppData/Roaming"))
    else:
        base = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    d = base / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d

def per_user_config_dir() -> Path:
    """
    Anbefalt sted å lagre bruker-spesifikke ting:
    %APPDATA%\KlientApp\users\<user-id>\
    """
    d = _user_config_dir() / "users" / user_id()
    d.mkdir(parents=True, exist_ok=True)
    return d

def _prefs_path() -> Path:
    return per_user_config_dir() / "prefs.json"

def load_user_prefs(default: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    p = _prefs_path()
    if p.exists():
        try:
            return json.loads(p.read_text("utf-8"))
        except Exception:
            pass
    return dict(default or {})

def save_user_prefs(prefs: dict[str, Any]) -> None:
    p = _prefs_path()
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(prefs, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)

# ----------------------- eksempel: rolle/tilgang -----------------------

def user_profile() -> dict[str, Any]:
    """
    Kombinerer identitet + lokale preferanser i én struktur.
    Kan utvides med roller/tilganger fra en sentral fil eller AD.
    """
    ident = current_identity()
    prefs = load_user_prefs(default={
        "language": "nb_NO",
        "theme": "light",
        "default_year": None,
        "default_source": "hovedbok",
    })
    return {
        "id": user_id(),
        "identity": ident.__dict__,
        "prefs": prefs,
    }

# ----------------------- (valgfritt) AD/Win32 dypere info -----------------------
# Hvis dere senere trenger SID/AD-attributter:
# - Installer pywin32 og bruk win32security.LookupAccountName/ConvertSidToStringSid
# - Eller kjør 'whoami /user' og parse output
# Vi lar det være av her for å unngå ekstra avhengigheter.

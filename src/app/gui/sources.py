# SPDX-License-Identifier: MIT
"""
Centralised path resolution for shared sources/locations.

This module is copied from the `bilagsverktoy_sources` project as
described in the document "Forsalg kildefiler oppsett globalt - sources".
It provides a single place to define logical names (such as
``client_dir`` or ``versions_dir``) and resolve them to concrete
filesystem paths based on a small configuration.  By centralising
path definitions here we avoid scattering hard‑coded directory
layouts throughout the codebase.  See the accompanying
``CONFIGURATION.md`` in the docs folder for usage details and
configuration options.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, Optional, Mapping, Any
import os, sys, json

# --- Utilities ---------------------------------------------------------------

def _user_config_base() -> Path:
    """Return the base directory for user‑level configuration files."""
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "bilagsverktoy"


def _user_state_base() -> Path:
    """Return the base directory for user‑level state files."""
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state"))
    return base / "bilagsverktoy"


def _load_toml(path: Path) -> Dict[str, Any]:
    """Load a TOML file using tomllib or tomli."""
    try:
        import tomllib  # Python ≥3.11 provides tomllib by default
        with open(path, "rb") as f:
            return tomllib.load(f)  # type: ignore[no-any-return]
    except ModuleNotFoundError:
        try:
            import tomli  # optional dependency for Python <3.11
            with open(path, "rb") as f:
                return tomli.load(f)
        except ModuleNotFoundError:
            raise RuntimeError(
                f"TOML config requested but neither 'tomllib' (py311+) nor 'tomli' is available: {path}"
            )


def _maybe_load_config(path: Path) -> Dict[str, Any]:
    """Load a configuration file if it exists; otherwise return an empty dict."""
    if not path.exists():
        return {}
    if path.suffix.lower() == ".toml":
        return _load_toml(path)
    if path.suffix.lower() == ".json":
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    # Unknown format -> ignore
    return {}


# Default templates define how logical keys map to concrete paths.
DEFAULT_TEMPLATES: Dict[str, str] = {
    "client_dir": "{clients_root}/{client}",
    "year_dir": "{clients_root}/{client}/{year}",
    "versions_dir": "{clients_root}/{client}/{year}/versions",
    "clientlist": "{clients_root}/_admin/master/clients.xlsx",
    "board_excel": "{clients_root}/{client}/board.xlsx",
    "registry_db": "{state_dir}/aksjonarregister.db",
    "logs_dir": "{clients_root}/{client}/{year}/logs",
    "mapping_dir": "{clients_root}/{client}/{year}/mapping",
}


def _default_clients_root() -> Path:
    """Determine the default clients_root using environment variables, services or fallbacks."""
    # 1) Environment override
    env = os.environ.get("CLIENTS_ROOT")
    if env:
        return Path(env).expanduser().resolve()

    # 2) Try services.clients.get_clients_root() if available
    try:
        from app.services.clients import get_clients_root  # type: ignore
        root = get_clients_root()
        if root:
            return Path(root).expanduser().resolve()
    except Exception:
        pass

    # 3) Fallback: use a directory named "Clients" in the user's home
    return Path.home() / "Clients"


def _default_state_dir() -> Path:
    """Determine the default state_dir using environment variables or fallbacks."""
    env = os.environ.get("STATE_DIR")
    if env:
        return Path(env).expanduser().resolve()
    # Fallback to platform‑specific state base
    return _user_state_base()


@dataclass(slots=True)
class SourcesConfig:
    """Configuration object holding roots and templates for path resolution."""
    clients_root: Path = field(default_factory=_default_clients_root)
    state_dir: Path = field(default_factory=_default_state_dir)
    templates: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_TEMPLATES))

    def with_overrides(self, **kwargs: Any) -> "SourcesConfig":
        """Return a new SourcesConfig with specified fields overridden."""
        return replace(self, **kwargs)


def _load_config_file() -> Dict[str, Any]:
    """Load the external configuration file if specified via ENV or in the default location."""
    # Environment variable override
    env_path = os.environ.get("BVT_SOURCES_FILE")
    if env_path:
        return _maybe_load_config(Path(env_path).expanduser())

    # Search default locations in the user's config directory
    cfg_dir = _user_config_base()
    for name in ("sources.toml", "sources.json"):
        p = cfg_dir / name
        if p.exists():
            return _maybe_load_config(p)
    return {}


def _merge_config(base: SourcesConfig, cfg: Dict[str, Any]) -> SourcesConfig:
    """Merge an external config dictionary with the default configuration."""
    clients_root = Path(cfg.get("clients_root", base.clients_root))
    state_dir = Path(cfg.get("state_dir", base.state_dir))
    templates = dict(DEFAULT_TEMPLATES)
    templates.update(cfg.get("templates", {}))
    return SourcesConfig(clients_root=clients_root, state_dir=state_dir, templates=templates)


# Global singleton‑like configuration; updated via set_clients_root/set_state_dir
_CFG = _merge_config(SourcesConfig(), _load_config_file())


def current_config() -> SourcesConfig:
    """Return the current global SourcesConfig."""
    return _CFG


def set_clients_root(path: str | Path) -> None:
    """Set the clients_root at runtime (e.g., when a user chooses the root folder in the GUI)."""
    global _CFG
    _CFG = _CFG.with_overrides(clients_root=Path(path).expanduser().resolve())


def set_state_dir(path: str | Path) -> None:
    """Set the state_dir at runtime."""
    global _CFG
    _CFG = _CFG.with_overrides(state_dir=Path(path).expanduser().resolve())


@contextmanager
def override_config(**kwargs: Any):
    """Temporarily override parts of the configuration (useful in tests)."""
    global _CFG
    old = _CFG
    try:
        _CFG = _CFG.with_overrides(**kwargs)
        yield _CFG
    finally:
        _CFG = old


def resolve(key: str, ensure_parent: bool = False, **params: Any) -> Path:
    """Resolve a logical key into a concrete path using the current configuration.

    Parameters
    ----------
    key: str
        Logical key to resolve, such as ``"year_dir"`` or ``"registry_db"``.
    ensure_parent: bool
        If True, create the parent directory of the resulting path if it does not exist.
    **params: Any
        Additional parameters required by the template (e.g., client name, year).

    Returns
    -------
    Path
        The fully resolved path.
    """
    tpl = _CFG.templates.get(key)
    if not tpl:
        raise KeyError(f"Unknown source key: {key!r} (known: {sorted(_CFG.templates)})")
    values = {
        "clients_root": str(_CFG.clients_root),
        "state_dir": str(_CFG.state_dir),
        **params,
    }
    p = Path(tpl.format(**values)).expanduser()
    if ensure_parent:
        p.parent.mkdir(parents=True, exist_ok=True)
    return p


# Convenience namespaced helpers (explicit is better than implicit)
class paths:
    @staticmethod
    def client_dir(client: str) -> Path:
        return resolve("client_dir", client=client)

    @staticmethod
    def year_dir(client: str, year: int | str) -> Path:
        return resolve("year_dir", client=client, year=year)

    @staticmethod
    def versions_dir(client: str, year: int | str) -> Path:
        return resolve("versions_dir", client=client, year=year)

    @staticmethod
    def clientlist() -> Path:
        return resolve("clientlist")

    @staticmethod
    def board_excel(client: str) -> Path:
        return resolve("board_excel", client=client)

    @staticmethod
    def registry_db() -> Path:
        return resolve("registry_db")

    @staticmethod
    def logs_dir(client: str, year: int | str) -> Path:
        return resolve("logs_dir", client=client, year=year)

    @staticmethod
    def mapping_dir(client: str, year: int | str) -> Path:
        return resolve("mapping_dir", client=client, year=year)

# End of module
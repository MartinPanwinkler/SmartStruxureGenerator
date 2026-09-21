from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Any


APP_NAME = "SmartStruxure Beschriftungsgenerator"
ROOT_DIR = Path(__file__).resolve().parent
GROUP_BY = ["asp", "module", "module_type"]
BOX_ROW_SPACING = 2


def _app_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        if sys.platform == "win32" and os.environ.get("LOCALAPPDATA"):
            return Path(os.environ["LOCALAPPDATA"]) / "SmartStruxureGenerator"
        state_home = os.environ.get("XDG_STATE_HOME")
        return Path(state_home) / "smartstruxure-generator" if state_home else Path.home() / ".smartstruxure-generator"
    return ROOT_DIR


APP_DATA_DIR = _app_data_dir()
LOG_DIR = APP_DATA_DIR / "logs"


def resource_path(relative: str | Path) -> Path:
    """Return a bundled resource path both in source and in a PyInstaller build."""
    base = Path(getattr(sys, "_MEIPASS", ROOT_DIR))
    return base / relative


def editable_resource(relative: str | Path) -> Path:
    """Prefer a user-editable file next to the EXE, otherwise use bundled data."""
    external_base = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else ROOT_DIR
    external = external_base / relative
    return external if external.exists() else resource_path(relative)


DEFAULT_TEMPLATE = editable_resource(Path("templates") / "Beschriftung SmartStruxure ERR.xlsx")
DEFAULT_CONFIG = editable_resource("template_config.json")


def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(LOG_DIR / "app.log", encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )


def load_json_config(path: Path = DEFAULT_CONFIG) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Konfigurationsdatei nicht gefunden: {path}")
    with path.open("r", encoding="utf-8") as handle:
        value = json.load(handle)
    if not isinstance(value, dict):
        raise ValueError("Die Template-Konfiguration muss ein JSON-Objekt sein.")
    return value

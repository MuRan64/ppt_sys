"""Locate bundled resources and writable user data on each desktop platform."""

from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys


PROJECT_DIR = Path(__file__).resolve().parents[1]


def resource_roots() -> list[Path]:
    roots: list[Path] = []
    if getattr(sys, "frozen", False):
        executable_dir = Path(sys.executable).resolve().parent
        if sys.platform == "darwin":
            roots.append(executable_dir.parent / "Resources")
        roots.append(executable_dir)
        bundle_dir = getattr(sys, "_MEIPASS", None)
        if bundle_dir:
            roots.append(Path(bundle_dir))
    roots.append(PROJECT_DIR)
    return list(dict.fromkeys(roots))


def resource_path(relative: str | Path) -> Path:
    relative = Path(relative)
    if relative.is_absolute() or ".." in relative.parts:
        raise ValueError("Resource paths must stay inside the application directory.")
    candidates = [root / relative for root in resource_roots()]
    return next((path for path in candidates if path.is_file()), candidates[0])


def user_data_directory() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "FbdToPpt"
    if sys.platform == "win32":
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "FbdToPpt"
    return Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "FbdToPpt"


def open_directory(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])

"""Product branding constants and asset resolution (crimSun).

Single source of truth for the product name, byline, and logo so the GUI,
packaging script, and docs cannot drift. `logo_path()` resolves the logo both
from a source checkout and from a PyInstaller one-file bundle (`sys._MEIPASS`).
"""

from __future__ import annotations

import sys
from pathlib import Path

PRODUCT_NAME = "Tongkat Gaib Excel"
PRODUCT_BYLINE = "by crimSun"
LOGO_FILENAME = "crimsun_logo.png"

# Directory that holds the bundled logo, relative to the app root / bundle root.
ASSETS_DIRNAME = "assets"


def logo_path() -> Path | None:
    """Returns the logo path, or None if the asset cannot be found.

    Checks the PyInstaller extraction dir (`sys._MEIPASS`) first so frozen exes
    resolve the bundled asset, then falls back to the repo `assets/` folder.
    Callers must tolerate None (missing asset is non-fatal for the GUI).
    """
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / ASSETS_DIRNAME / LOGO_FILENAME)
    # src/branding.py -> repo root is two levels up.
    candidates.append(
        Path(__file__).resolve().parent.parent / ASSETS_DIRNAME / LOGO_FILENAME
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None

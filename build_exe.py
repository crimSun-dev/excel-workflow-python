"""PyInstaller build script (TDD Phase 4).

Bundles the pipeline into a single standalone Windows executable that runs on a
clean machine without Python installed:

    python build_exe.py

Produces: dist/TongkatGaibExcel.exe
"""

from __future__ import annotations

import os
import subprocess
import sys


# Heavy libraries that may live in a shared/global site-packages but which this
# tool never imports. Excluding them keeps the bundle small (~40-60 MB vs 490 MB+).
_EXCLUDES = [
    "torch",
    "torchvision",
    "cv2",
    "scipy",
    "numba",
    "llvmlite",
    "matplotlib",
    "pygame",
    "av",
    "IPython",
    "notebook",
    "sympy",
    "sklearn",
    "tensorflow",
    "PyQt5",
    "PySide2",
]


def build() -> int:
    # Bundle the crimSun logo into the one-file exe so the GUI/icon resolve it
    # from sys._MEIPASS without a side-car assets folder. PyInstaller expects
    # "SRC<os.pathsep>DEST_DIR"; the app looks under "assets/" at runtime.
    logo_data = f"assets{os.sep}crimsun_logo.png{os.pathsep}assets"

    args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconfirm",
        "--clean",
        "--name",
        "TongkatGaibExcel",
        # tkinter hidden import PyInstaller sometimes misses; pydantic needs
        # its compiled core collected. The modern pandas hook handles pandas.
        "--hidden-import",
        "tkinter",
        "--collect-all",
        "pydantic",
        "--add-data",
        logo_data,
    ]
    for mod in _EXCLUDES:
        args += ["--exclude-module", mod]
    args.append("main.py")
    print("Running:", " ".join(args))
    return subprocess.call(args)


if __name__ == "__main__":
    raise SystemExit(build())

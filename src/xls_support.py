"""Excel 97-2003 (.xls) support via xlrd.

pandas treats xlrd as optional and only imports it when an .xls path appears.
PyInstaller therefore cannot see the dependency, and the Windows launcher used
to skip installing it because its package check never mentioned xlrd. Giro's
original master (`DAFTAR REKENING GIRO TSPM.xls`) then fails with a pip-install
message the operator cannot act on.

Importing xlrd here makes the requirement visible to the freeze, the launcher,
and every workbook reader in this package.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

try:
    import xlrd
except ImportError:  # pragma: no cover - incomplete install / frozen miss
    xlrd = None

XLS_ENGINE = "xlrd"

MISSING_XLRD_MESSAGE = (
    "Cannot read Excel 97-2003 (.xls) files because xlrd is missing from this "
    "installation. In Excel, open the workbook and use File > Save As > "
    "Excel Workbook (*.xlsx), then select that file instead."
)


def ensure_xlrd() -> None:
    """Raises if xlrd is not importable, with an operator-facing message."""
    if xlrd is None:
        raise ModuleNotFoundError(MISSING_XLRD_MESSAGE)


def open_excel(path: Path | str) -> pd.ExcelFile:
    """`pd.ExcelFile` using xlrd for `.xls` and pandas' default otherwise."""
    path = Path(path)
    if path.suffix.lower() == ".xls":
        ensure_xlrd()
        return pd.ExcelFile(path, engine=XLS_ENGINE)
    return pd.ExcelFile(path)

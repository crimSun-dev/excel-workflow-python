"""Shared pytest fixtures: deterministic data & reference workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def raw_pipe_file(tmp_path: Path) -> Path:
    """A valid 3-field pipe-delimited raw file with 4 data rows."""
    content = (
        "KODE_UKER|SEGMEN|VOLUME_IN_IDR\n"
        "0001|Wholesale|1250000000000\n"
        "0002|Corporate|500000000.50\n"
        "0001|Wholesale|250000000000\n"
        "0003|Retail|1000000\n"
    )
    path = tmp_path / "raw_data.txt"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def raw_pipe_file_cp1252(tmp_path: Path) -> Path:
    """Raw file with a cp1252-only character (é as 0xE9) to test encoding fallback."""
    content = (
        "KODE_UKER|SEGMEN|VOLUME_IN_IDR\n"
        "0001|Wholesalé|1250000000000\n"
    )
    path = tmp_path / "raw_cp1252.txt"
    path.write_bytes(content.encode("cp1252"))
    return path


@pytest.fixture
def raw_pipe_file_whitespace(tmp_path: Path) -> Path:
    """Raw file with padded values to test whitespace trimming."""
    content = (
        "KODE_UKER|SEGMEN|VOLUME_IN_IDR\n"
        "  0001  | Wholesale |1250000000000\n"
    )
    path = tmp_path / "raw_ws.txt"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def raw_pipe_file_malformed(tmp_path: Path) -> Path:
    """Raw file with one malformed row (extra pipe -> extra column)."""
    content = (
        "KODE_UKER|SEGMEN|VOLUME_IN_IDR\n"
        "0001|Wholesale|1250000000000\n"
        "0002|Cor|porate|500000000\n"   # malformed: 4 fields
        "0003|Retail|1000000\n"
    )
    path = tmp_path / "raw_malformed.txt"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def reference_file(tmp_path: Path) -> Path:
    """Reference workbook mapping KODE_UKER -> MAIN_CODE, MAIN_BRANCH.

    Note: 0003 is intentionally omitted to exercise the UNMAPPED path.
    """
    ref = pd.DataFrame(
        {
            "KODE_UKER": ["0001", "0002"],
            "MAIN_CODE": ["MC10", "MC20"],
            "MAIN_BRANCH": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference.xlsx"
    ref.to_excel(path, index=False)
    return path

"""Shared pytest fixtures: deterministic data & reference workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def raw_pipe_file(tmp_path: Path) -> Path:
    """A valid pipe-delimited Akumulasi raw file with 4 data rows.

    Includes an FBI column (summed by the Akumulasi workflow) alongside
    VOLUME_IN_IDR. FBI totals by MAIN_CODE after enrichment:
        0001 -> MC10: 100 + 200 = 300
        0002 -> MC20: 50.25
        0003 -> UNMAPPED: 10
    Volume totals are unchanged from the original 4-row fixture.
    """
    content = (
        "KODE_UKER|SEGMEN|FBI|VOLUME_IN_IDR\n"
        "0001|Wholesale|100|1250000000000\n"
        "0002|Corporate|50.25|500000000.50\n"
        "0001|Wholesale|200|250000000000\n"
        "0003|Retail|10|1000000\n"
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


@pytest.fixture
def reference_file_kanca(tmp_path: Path) -> Path:
    """Reference workbook using aliased headers instead of canonical names.

    KANCA -> KODE_UKER, UNIQUE CODE -> MAIN_CODE, DESCRIPTION -> MAIN_BRANCH.
    0003 is intentionally omitted to exercise the UNMAPPED path.
    """
    ref = pd.DataFrame(
        {
            "KANCA": ["0001", "0002"],
            "UNIQUE CODE": ["MC10", "MC20"],
            "DESCRIPTION": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_kanca.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def reference_file_multisheet(tmp_path: Path) -> Path:
    """Multi-sheet workbook where the mapping table is not on the first sheet."""
    decoy = pd.DataFrame(
        {
            "ACCOUNT NUMBER": ["100", "200"],
            "PRODUK": ["A", "B"],
        }
    )
    mapping = pd.DataFrame(
        {
            "KANCA": ["0001", "0002"],
            "UNIQUE CODE": ["MC10", "MC20"],
            "DESCRIPTION": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_multisheet.xlsx"
    with pd.ExcelWriter(path) as writer:
        decoy.to_excel(writer, sheet_name="Cover", index=False)
        mapping.to_excel(writer, sheet_name="Mapping", index=False)
    return path


@pytest.fixture
def reference_file_unit_kerja(tmp_path: Path) -> Path:
    """Reference workbook matching BRIEFX 'Unit Kerja' / UKER sheet layout."""
    ref = pd.DataFrame(
        {
            "KODE SUB KANCA": ["0001", "0002"],
            "SUB KANCA": ["Unit A", "Unit B"],
            "KODE KANCA": ["MC10", "MC20"],
            "KANCA": ["0001", "0002"],
            "KANWIL": ["JW1", "JW1"],
            "KODE": ["MC10", "MC20"],
        }
    )
    path = tmp_path / "reference_unit_kerja.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def reference_file_uker_new(tmp_path: Path) -> Path:
    """Reference workbook matching BRIEFX 'UKER NEW' sheet layout."""
    ref = pd.DataFrame(
        {
            "KODE UNIT": ["0001", "0002"],
            "DESC UNIT": ["Unit A", "Unit B"],
            "KODE KANCA": ["MC10", "MC20"],
            "DESC KANCA": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_uker_new.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def reference_file_briefx_multisheet(tmp_path: Path) -> Path:
    """Multi-sheet workbook like the real BRIEFX file (GL + Unit Kerja + UKER NEW)."""
    gl_sheet = pd.DataFrame(
        {
            "ACCOUNT NUMBER": ["100", "200"],
            "DESCRIPTION": ["GL A", "GL B"],
            "C/C": ["x", "y"],
            "UNIQUE CODE": ["U1", "U2"],
            "PRODUK": ["P1", "P2"],
        }
    )
    unit_kerja = pd.DataFrame(
        {
            "KODE SUB KANCA": ["0001", "0002"],
            "SUB KANCA": ["Wrong A", "Wrong B"],
            "KODE KANCA": ["WRONG", "WRONG"],
            "KANCA": ["0001", "0002"],
            "KANWIL": ["JW1", "JW1"],
            "KODE": ["WRONG", "WRONG"],
        }
    )
    uker_sheet = pd.DataFrame(
        {
            "KODE UNIT": ["0001", "0002"],
            "DESC UNIT": ["Unit A", "Unit B"],
            "KODE KANCA": ["MC10", "MC20"],
            "DESC KANCA": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_briefx.xlsx"
    with pd.ExcelWriter(path) as writer:
        gl_sheet.to_excel(writer, sheet_name="Mapping GL FBI INT", index=False)
        unit_kerja.to_excel(writer, sheet_name="Unit Kerja", index=False)
        uker_sheet.to_excel(writer, sheet_name="UKER NEW 2025 (KC BATU)", index=False)
    return path


@pytest.fixture
def reference_file_unresolvable(tmp_path: Path) -> Path:
    """Reference workbook with no column resolvable to the required logical columns."""
    ref = pd.DataFrame(
        {
            "ACCOUNT NUMBER": ["100", "200"],
            "PRODUK": ["A", "B"],
            "C/C": ["x", "y"],
        }
    )
    path = tmp_path / "reference_unresolvable.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def rincian_vol_tf_file(tmp_path: Path) -> Path:
    """Pipe-delimited sample for the Rincian Vol TF workflow.

    Exercises Wholesale exclusion, a blank-SEGMEN row (must be kept), and a
    branch (B03) whose only row is Wholesale (must drop out entirely).
    Expected after processing: B01/Branch One = 2000, B02/Branch Two = 800.
    """
    content = (
        "SEGMEN|MAINBR|MBDESC|AMOUNT_IN_IDR\n"
        "Wholesale|B01|Branch One|1000\n"
        "Corporate|B01|Branch One|2000\n"
        "Retail|B02|Branch Two|500\n"
        "|B02|Branch Two|300\n"
        "Wholesale|B03|Branch Three|9999\n"
    )
    path = tmp_path / "rincian_vol_tf_sample.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def rincian_portal_bg_file(tmp_path: Path) -> Path:
    """Pipe-delimited sample for the Rincian Portal BG workflow (no filtering).

    Expected after processing: B01/Branch One = 3000, B02/Branch Two = 500.
    """
    content = (
        "MAINBR|MBNAME|AMOUNT_IN_IDR\n"
        "B01|Branch One|1000\n"
        "B01|Branch One|2000\n"
        "B02|Branch Two|500\n"
    )
    path = tmp_path / "rincian_portal_bg_sample.csv"
    path.write_text(content, encoding="utf-8")
    return path

"""Shared pytest fixtures: deterministic data & reference workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

# --------------------------------------------------------------------------- #
# Real-sample acceptance oracles (Time Series Briva / Qlola).
#
# These vendor workbooks are large and git-ignored (`*.xlsx`), so the fixtures
# below derive compact master/UKER mappings from them at runtime and SKIP when
# the samples are absent. Path A': the Qlola master ID->MAIN_CODE mapping is
# derived from the sample Sheet1 intermediate table (the post-step-4 mapping the
# manual workflow produced); the original vendor MASTER/UKER files were never
# available.
# --------------------------------------------------------------------------- #
SAMPLES_DIR = Path(__file__).parent / "fixtures" / "samples"
QLOLA_RAW = SAMPLES_DIR / "1784686455_REPORT_SUMMARY_MONTHLY_2026-07-20_QLOLA_ALL.csv"
QLOLA_XLSX = (
    SAMPLES_DIR / "1784686455_REPORT_SUMMARY_MONTHLY_2026-07-20_QLOLA_ALL.csv.xlsx"
)
BRIVA_RAW = SAMPLES_DIR / "NON_WHOLESALE_2026-07-21_1784773298.csv"
BRIVA_XLSX = SAMPLES_DIR / "NON_WHOLESALE_2026-07-21_1784773298.csv.xlsx"


@pytest.fixture(scope="session")
def qlola_sample_inputs(tmp_path_factory) -> dict[str, Path]:
    """Derives Qlola raw + UKER + master fixtures from the real sample (A')."""
    if not QLOLA_XLSX.exists() or not QLOLA_RAW.exists():
        pytest.skip("QLOLA sample workbook not present (git-ignored); oracle skipped")
    out_dir = tmp_path_factory.mktemp("qlola_sample")

    # Master ID -> MAIN_CODE from Sheet1 intermediate table (cols D, E).
    sheet1 = pd.read_excel(QLOLA_XLSX, sheet_name="Sheet1", header=None)
    inter = sheet1.iloc[:, [3, 4]].copy()
    inter.columns = ["ID", "MAIN_CODE"]
    inter = inter[inter["ID"].notna() & (inter["ID"].astype(str) != "ID")]
    master_path = out_dir / "qlola_master.csv"
    inter.to_csv(master_path, index=False)

    # UKER reference (detail context only) from the detail sheet.
    detail_sheet = pd.ExcelFile(QLOLA_XLSX).sheet_names[1]
    detail = pd.read_excel(QLOLA_XLSX, sheet_name=detail_sheet)
    uker = detail[["KODE_UKER", "MAIN_CODE", "MAIN_BRANCH"]].drop_duplicates("KODE_UKER")
    uker_path = out_dir / "qlola_uker.csv"
    uker.to_csv(uker_path, index=False)

    return {"raw": QLOLA_RAW, "uker": uker_path, "master": master_path}


@pytest.fixture(scope="session")
def briva_sample_inputs(tmp_path_factory) -> dict[str, Path]:
    """Derives Briva raw + UKER fixtures from the real sample workbook."""
    if not BRIVA_XLSX.exists() or not BRIVA_RAW.exists():
        pytest.skip("Briva sample workbook not present (git-ignored); oracle skipped")
    out_dir = tmp_path_factory.mktemp("briva_sample")

    detail_sheet = pd.ExcelFile(BRIVA_XLSX).sheet_names[1]
    detail = pd.read_excel(BRIVA_XLSX, sheet_name=detail_sheet)
    detail.columns = [str(c).strip() for c in detail.columns]
    uker = detail[["KODE UKER", "MAIN_CODE", "MAIN_BRANCH"]].drop_duplicates("KODE UKER")
    uker_path = out_dir / "briva_uker.csv"
    uker.to_csv(uker_path, index=False)

    return {"raw": BRIVA_RAW, "uker": uker_path}


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
def reference_file_uker(tmp_path: Path) -> Path:
    """Reference workbook whose join key is spelled `UKER` (no KODE_UKER column)."""
    ref = pd.DataFrame(
        {
            "UKER": ["0001", "0002"],
            "MAIN_CODE": ["MC10", "MC20"],
            "MAIN_BRANCH": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_uker.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def reference_file_kode_uker_spaced(tmp_path: Path) -> Path:
    """Reference workbook whose join key is the spaced `KODE UKER` variant."""
    ref = pd.DataFrame(
        {
            "KODE UKER": ["0001", "0002"],
            "MAIN_CODE": ["MC10", "MC20"],
            "MAIN_BRANCH": ["Jakarta Pusat", "Surabaya"],
        }
    )
    path = tmp_path / "reference_kode_uker_spaced.xlsx"
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
def timeseries_briva_file(tmp_path: Path) -> Path:
    """Pipe-delimited Time Series FBI Briva sample.

    Uses the spaced `KODE UKER` header (exercises raw-side alias resolution) and
    a WHOLESALE row that MUST be excluded by the NONWHOLESALE inclusion filter.
    NONWHOLESALE VOLUME_IDR by branch (via reference_file 0001->MC10, 0002->MC20):
        MC10 = 1481517421603.75
        MC20 = 1000000000000.00
    Grand Total = 2481517421603.75 (the sample Sheet1 oracle constant).
    """
    content = (
        "SEGMEN|KODE UKER|VOLUME_IDR\n"
        "NONWHOLESALE|0001|1481517421603.75\n"
        "NONWHOLESALE|0002|1000000000000\n"
        "WHOLESALE|0001|999999999999\n"
    )
    path = tmp_path / "timeseries_briva_sample.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def timeseries_briva_file_underscore_key(tmp_path: Path) -> Path:
    """Briva sample using the canonical underscored `KODE_UKER` header.

    Same NONWHOLESALE totals as `timeseries_briva_file`, so both header variants
    are provably interchangeable.
    """
    content = (
        "SEGMEN|KODE_UKER|VOLUME_IDR\n"
        "NONWHOLESALE|0001|1481517421603.75\n"
        "NONWHOLESALE|0002|1000000000000\n"
        "WHOLESALE|0001|999999999999\n"
    )
    path = tmp_path / "timeseries_briva_underscore.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def timeseries_briva_file_with_unmapped(tmp_path: Path) -> Path:
    """Briva sample where 0009 is absent from the reference (=> UNMAPPED row).

    NONWHOLESALE VOLUME_IDR by branch (via reference_file 0001->MC10, 0002->MC20):
        MC10 = 100, MC20 = 200, UNMAPPED (0009) = 50; Grand Total = 350.
    """
    content = (
        "SEGMEN|KODE UKER|VOLUME_IDR\n"
        "NONWHOLESALE|0001|100\n"
        "NONWHOLESALE|0002|200\n"
        "NONWHOLESALE|0009|50\n"
    )
    path = tmp_path / "timeseries_briva_unmapped.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def qlola_raw_file_numeric_ids(tmp_path: Path) -> Path:
    """Qlola sample whose IDs are numeric (int-dtype on the master side).

    U-prefixed IDs would mask the int-vs-str VLOOKUP mismatch this exercises.
    """
    content = (
        "SOURCE|ID|FREKUENSI|KODE_UKER\n"
        "QCASH|1001|6|K1\n"
        "QCASH|1002|2|K1\n"
        "QCASH|1003|7|K2\n"
    )
    path = tmp_path / "qlola_numeric_ids.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def qlola_master_numeric_file(tmp_path: Path) -> Path:
    """Master ID -> MAIN_CODE workbook storing IDs as integers (1003 omitted)."""
    master = pd.DataFrame({"ID": [1001, 1002], "MAIN_CODE": ["7", "9"]})
    path = tmp_path / "qlola_master_numeric.xlsx"
    master.to_excel(path, index=False)
    return path


@pytest.fixture
def qlola_raw_file(tmp_path: Path) -> Path:
    """Pipe-delimited Time Series Active User Qlola sample.

    Exercises: CMS exclusion (U1's CMS row and U6's CMS-only rows dropped),
    FREKUENSI summed per ID (U1 = 3 + 4 = 7), the >=5 active threshold, an
    UNMAPPED master ID (U5), and the master-vs-UKER guard (all KODE_UKER map to
    UKER MAIN_CODE 99, but the crosstab must key on master codes 7/9/UNMAPPED).
    """
    content = (
        "SOURCE|ID|FREKUENSI|KODE_UKER\n"
        "QCASH|U1|3|K1\n"
        "QIB|U1|4|K1\n"
        "CMS|U1|100|K1\n"
        "QCASH|U2|2|K2\n"
        "QCASH|U3|5|K1\n"
        "QCASH|U4|1|K2\n"
        "QCASH|U5|9|K1\n"
        "CMS|U6|10|K1\n"
    )
    path = tmp_path / "qlola_sample.csv"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def qlola_uker_reference_file(tmp_path: Path) -> Path:
    """UKER reference for Qlola; every key maps to MAIN_CODE 99 (detail context).

    Deliberately distinct from the master mapping so a test can prove the final
    crosstab uses master MAIN_CODE, not this UKER-enriched one.
    """
    ref = pd.DataFrame(
        {
            "KODE_UKER": ["K1", "K2"],
            "MAIN_CODE": ["99", "99"],
            "MAIN_BRANCH": ["Uker Branch 99", "Uker Branch 99"],
        }
    )
    path = tmp_path / "qlola_uker_reference.xlsx"
    ref.to_excel(path, index=False)
    return path


@pytest.fixture
def qlola_master_file(tmp_path: Path) -> Path:
    """Master-data ID -> MAIN_CODE mapping for Qlola (U5 omitted => UNMAPPED)."""
    master = pd.DataFrame(
        {
            "ID": ["U1", "U2", "U3", "U4", "U6"],
            "MAIN_CODE": ["7", "7", "9", "7", "9"],
        }
    )
    path = tmp_path / "qlola_master.xlsx"
    master.to_excel(path, index=False)
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

"""Phase 1 unit tests: Ingestion, Enrichment, Aggregation (TDD Section 4)."""

from __future__ import annotations

from decimal import Decimal

import pandas as pd
import pytest

from src.aggregation import AggregationEngine
from src.enrichment import ReferenceEnricher, ReferenceEnrichmentError
from src.ingestion import DataIngestionError, IngestionEngine


# --------------------------------------------------------------------------- #
# Ingestion
# --------------------------------------------------------------------------- #
def test_ingest_reads_valid_pipe_delimited_file(raw_pipe_file):
    result = IngestionEngine().read_raw_data(raw_pipe_file)
    assert list(result.data.columns) == ["KODE_UKER", "SEGMEN", "FBI", "VOLUME_IN_IDR"]
    assert result.total_rows == 4
    assert result.malformed_rows_count == 0
    assert result.data["VOLUME_IN_IDR"].dtype == "float64"


def test_ingest_handles_utf8_and_cp1252_encodings(raw_pipe_file_cp1252):
    result = IngestionEngine().read_raw_data(raw_pipe_file_cp1252)
    assert result.total_rows == 1
    assert "Wholesal" in result.data["SEGMEN"].iloc[0]


def test_ingest_strips_leading_trailing_whitespace(raw_pipe_file_whitespace):
    result = IngestionEngine().read_raw_data(raw_pipe_file_whitespace)
    assert result.data["KODE_UKER"].iloc[0] == "0001"
    assert result.data["SEGMEN"].iloc[0] == "Wholesale"


def test_ingest_skips_and_counts_malformed_rows(raw_pipe_file_malformed):
    result = IngestionEngine().read_raw_data(raw_pipe_file_malformed)
    assert result.total_rows == 2  # 3 data rows, 1 malformed skipped
    assert result.malformed_rows_count == 1


def test_ingest_empty_file_raises(tmp_path):
    empty = tmp_path / "empty.txt"
    empty.write_text("", encoding="utf-8")
    with pytest.raises(DataIngestionError, match="empty"):
        IngestionEngine().read_raw_data(empty)


def test_ingest_header_only_file_raises(tmp_path):
    """A vendor extract with titles but no data rows must not look like success."""
    header_only = tmp_path / "RINCIAN_VOL_TF_empty.csv"
    header_only.write_text(
        "TAHUN|PERIODE|POSISI|REGION|RGDESC|MAINBR|MBDESC|BRANCH|BRDESC|"
        "PRODUK|MASTERREF|CIFNO|CUST_NAME|GRUP|SEGMEN|DIVISI|CCY|"
        "AMOUNT_IN_IDR|AMOUNT_ORI|ISSUED_DATE|PNRM|RMNAME\n",
        encoding="utf-8",
    )
    with pytest.raises(DataIngestionError, match="no data rows"):
        IngestionEngine().read_raw_data(header_only)


def test_ingest_skips_title_banner_above_headers(tmp_path):
    """Headers in the middle of a file still parse as the real table."""
    path = tmp_path / "banner.txt"
    path.write_text(
        "Rincian Portal BG export\n"
        "Generated: 2026-08-01\n"
        "\n"
        "MAINBR|MBNAME|AMOUNT_IN_IDR\n"
        "B01|Branch One|1000\n"
        "B02|Branch Two|500\n",
        encoding="utf-8",
    )
    result = IngestionEngine(numeric_columns=("AMOUNT_IN_IDR",)).read_raw_data(path)
    assert list(result.data.columns) == ["MAINBR", "MBNAME", "AMOUNT_IN_IDR"]
    assert result.total_rows == 2
    assert result.header_row == 4
    assert result.data["AMOUNT_IN_IDR"].tolist() == [1000.0, 500.0]


def test_ingest_keeps_values_on_both_sides_of_a_mid_file_header(tmp_path):
    """Headers inserted in the middle of the values are not a reason to drop rows."""
    path = tmp_path / "mid_header.csv"
    path.write_text(
        "B01|Branch One|1000\n"
        "B01|Branch One|2000\n"
        "MAINBR|MBNAME|AMOUNT_IN_IDR\n"
        "B02|Branch Two|500\n"
        "B03|Branch Three|250\n",
        encoding="utf-8",
    )
    result = IngestionEngine(numeric_columns=("AMOUNT_IN_IDR",)).read_raw_data(path)
    assert list(result.data.columns) == ["MAINBR", "MBNAME", "AMOUNT_IN_IDR"]
    assert result.header_row == 3
    assert result.total_rows == 4
    assert result.data["MAINBR"].tolist() == ["B01", "B01", "B02", "B03"]
    assert result.data["AMOUNT_IN_IDR"].tolist() == [1000.0, 2000.0, 500.0, 250.0]


def test_ingest_finds_headers_on_a_deep_row(tmp_path):
    """No fixed scan window — titles on row 399 are still found."""
    lines = [f"B01|Branch One|{i}\n" for i in range(398)]
    lines.append("MAINBR|MBNAME|AMOUNT_IN_IDR\n")
    lines.append("B02|Branch Two|999\n")
    path = tmp_path / "header_row_399.csv"
    path.write_text("".join(lines), encoding="utf-8")
    result = IngestionEngine(numeric_columns=("AMOUNT_IN_IDR",)).read_raw_data(path)
    assert result.header_row == 399
    assert result.total_rows == 399  # 398 above + 1 below
    assert result.data["MAINBR"].iloc[0] == "B01"
    assert result.data["MAINBR"].iloc[-1] == "B02"
    assert result.data["AMOUNT_IN_IDR"].iloc[-1] == 999.0


def test_ingest_reads_excel_workbook_with_mid_file_headers(tmp_path):
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["B01", "Branch One", 1000])
    sheet.append(["B01", "Branch One", 2000])
    sheet.append(["MAINBR", "MBNAME", "AMOUNT_IN_IDR"])
    sheet.append(["B02", "Branch Two", 500])
    path = tmp_path / "portal_bg_mid.xlsx"
    workbook.save(path)

    result = IngestionEngine(numeric_columns=("AMOUNT_IN_IDR",)).read_raw_data(path)
    assert list(result.data.columns) == ["MAINBR", "MBNAME", "AMOUNT_IN_IDR"]
    assert result.header_row == 3
    assert result.total_rows == 3
    assert result.data["AMOUNT_IN_IDR"].tolist() == [1000.0, 2000.0, 500.0]


def test_ingest_missing_file_raises(tmp_path):
    with pytest.raises(DataIngestionError):
        IngestionEngine().read_raw_data(tmp_path / "does_not_exist.txt")


# --------------------------------------------------------------------------- #
# Enrichment
# --------------------------------------------------------------------------- #
def test_enrichment_matches_valid_kode_uker_keys(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Jakarta Pusat"
    assert result.matched_count == 3  # two 0001 rows + one 0002


def test_enrichment_flags_unmapped_kode_uker_keys(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file).enrich(raw)
    assert result.unmapped_count == 1  # 0003 not in reference
    assert "0003" in result.unmapped_keys
    orphan = result.data[result.data["KODE_UKER"] == "0003"].iloc[0]
    assert orphan["MAIN_CODE"] == "UNMAPPED"
    assert orphan["MAIN_BRANCH"] == "UNMAPPED"


def test_enrichment_missing_lookup_key_raises(reference_file):
    df = pd.DataFrame({"WRONG_KEY": ["0001"], "VOLUME_IN_IDR": [1.0]})
    with pytest.raises(ReferenceEnrichmentError):
        ReferenceEnricher(reference_file).enrich(df)


def test_enrichment_missing_reference_file_raises(tmp_path, raw_pipe_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    with pytest.raises(ReferenceEnrichmentError):
        ReferenceEnricher(tmp_path / "nope.xlsx").enrich(raw)


def test_enrichment_resolves_kanca_and_output_aliases(raw_pipe_file, reference_file_kanca):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file_kanca).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Jakarta Pusat"
    assert result.matched_count == 3
    assert result.unmapped_count == 1  # 0003 not in reference


def test_enrichment_scans_non_default_sheet(raw_pipe_file, reference_file_multisheet):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file_multisheet).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0002"].iloc[0]
    assert row["MAIN_CODE"] == "MC20"
    assert row["MAIN_BRANCH"] == "Surabaya"
    assert result.matched_count == 3


def test_enrichment_canonical_reference_unchanged(raw_pipe_file, reference_file):
    """Regression: canonical KODE_UKER reference files behave identically."""
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Jakarta Pusat"
    assert result.matched_count == 3
    assert result.unmapped_count == 1


def test_enrichment_resolves_unit_kerja_sheet(raw_pipe_file, reference_file_unit_kerja):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file_unit_kerja).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Unit A"
    assert result.matched_count == 3


def test_enrichment_resolves_uker_new_sheet(raw_pipe_file, reference_file_uker_new):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file_uker_new).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Jakarta Pusat"
    assert result.matched_count == 3


def test_enrichment_skips_gl_sheet_and_uses_uker_tab(
    raw_pipe_file, reference_file_briefx_multisheet
):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file_briefx_multisheet).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0002"].iloc[0]
    assert row["MAIN_CODE"] == "MC20"
    assert row["MAIN_BRANCH"] == "Surabaya"
    assert result.matched_count == 3


def test_enrichment_unresolvable_reference_raises_with_detail(
    raw_pipe_file, reference_file_unresolvable
):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    with pytest.raises(ReferenceEnrichmentError) as excinfo:
        ReferenceEnricher(reference_file_unresolvable).enrich(raw)
    message = str(excinfo.value)
    assert "Sheets scanned" in message
    assert "Aliases attempted" in message
    assert "ACCOUNT NUMBER" in message  # columns found are surfaced
    assert "KANCA" in message  # aliases attempted are surfaced


# --------------------------------------------------------------------------- #
# Aggregation
# --------------------------------------------------------------------------- #
def test_aggregation_groups_by_code_and_sums_volume(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine().aggregate(enriched)

    mc10 = result.summary_data[result.summary_data["MAIN_CODE"] == "MC10"].iloc[0]
    # 1_250_000_000_000 + 250_000_000_000 = 1_500_000_000_000
    assert mc10["VOLUME_IN_IDR"] == 1_500_000_000_000.0
    assert result.branch_count == 3  # MC10, MC20, UNMAPPED


def test_aggregation_monetary_precision_matches_decimal(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine().aggregate(enriched)
    reference_total = Decimal("1250000000000") + Decimal("500000000.50") + Decimal(
        "250000000000"
    ) + Decimal("1000000")
    assert result.total_volume_idr == float(round(reference_total, 2))


def test_aggregation_applies_segment_filter(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine(segment_filter="Wholesale").aggregate(enriched)
    # Only the two Wholesale rows (both 0001 -> MC10) survive the filter.
    assert result.branch_count == 1
    assert result.summary_data.iloc[0]["MAIN_CODE"] == "MC10"
    assert result.total_volume_idr == 1_500_000_000_000.0


def test_xlrd_is_installed_for_legacy_xls():
    """Giro's original master is .xls; pandas will not read it without xlrd."""
    import xlrd

    major = int(xlrd.__VERSION__.split(".")[0])
    assert major >= 2


def test_open_excel_uses_xlrd_engine_for_xls(tmp_path, monkeypatch):
    from src import xls_support

    captured: dict[str, object] = {}

    def fake_excel_file(path, engine=None):
        captured["engine"] = engine
        captured["path"] = path
        return object()

    monkeypatch.setattr(xls_support.pd, "ExcelFile", fake_excel_file)
    xls_support.open_excel(tmp_path / "DAFTAR REKENING GIRO TSPM.xls")
    assert captured["engine"] == "xlrd"


def test_open_excel_xlsx_leaves_engine_to_pandas(tmp_path, monkeypatch):
    from src import xls_support

    captured: dict[str, object] = {}

    def fake_excel_file(path, engine=None):
        captured["engine"] = engine
        return object()

    monkeypatch.setattr(xls_support.pd, "ExcelFile", fake_excel_file)
    xls_support.open_excel(tmp_path / "master.xlsx")
    assert captured["engine"] is None


def test_missing_xlrd_tells_operator_to_save_as_xlsx(monkeypatch):
    from src import xls_support

    monkeypatch.setattr(xls_support, "xlrd", None)
    with pytest.raises(ModuleNotFoundError, match="Save As"):
        xls_support.ensure_xlrd()

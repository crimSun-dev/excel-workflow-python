"""Multi-workflow tests: strategy dispatch, new workflows, and E2E pipeline."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.aggregation import AggregationEngine
from src.enrichment import (
    MasterDataEnricher,
    ReferenceEnricher,
    ReferenceEnrichmentError,
    canonicalize_join_id,
    warn_if_mostly_unmapped,
)
from src.gui import (
    default_source_text,
    parse_source_exclude,
    source_filter_enabled,
    unmapped_warning_text,
)
from src.ingestion import IngestionEngine
from src.orchestrator import PipelineOrchestrator
from src.schemas import ProcessingConfig
from src.workflows.base import WorkflowId, WorkflowValidationError, normalize_join_key
from src.workflows.registry import WORKFLOW_REGISTRY, get_definition, get_strategy


# --------------------------------------------------------------------------- #
# canonicalize_join_id unit tests (Task 1.3)
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("1001", "1001"),
        ("1001.0", "1001"),        # Excel int→float→str round-trip
        ("1001.00", "1001"),       # double trailing zeros
        ("  1001  ", "1001"),      # padded whitespace
        (" 1001.0 ", "1001"),      # padding + trailing .0
        (1001, "1001"),            # actual int input
        (1001.0, "1001"),          # actual float input
        ("U1", "U1"),             # non-numeric preserved
        ("", ""),                  # empty string
        ("1001.5", "1001.5"),      # fractional part preserved
        ("00123", "00123"),        # leading zeros preserved (distinct ID)
    ],
)
def test_canonicalize_join_id(raw, expected):
    assert canonicalize_join_id(raw) == expected


def test_canonicalize_fixes_float_master_id_mismatch(tmp_path):
    """End-to-end: master stores int 1001, raw has string '1001' → match, not UNMAPPED."""
    master = pd.DataFrame({"ID": [1001, 1002], "MAIN_CODE": ["7", "9"]})
    path = tmp_path / "master_int.xlsx"
    master.to_excel(path, index=False)
    mapping = MasterDataEnricher(path).build_id_to_main_code()
    # canonicalize_join_id collapses 1001.0 → "1001"
    assert mapping.get("1001") == "7"
    assert mapping.get("1002") == "9"


# --------------------------------------------------------------------------- #
# Registry / dispatch
# --------------------------------------------------------------------------- #
def test_registry_has_all_five_workflows():
    assert set(WORKFLOW_REGISTRY) == {
        WorkflowId.AKUMULASI,
        WorkflowId.RINCIAN_VOL_TF,
        WorkflowId.RINCIAN_PORTAL_BG,
        WorkflowId.TIMESERIES_FBI_BRIVA,
        WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA,
    }


def test_get_strategy_accepts_string_id():
    strategy = get_strategy("rincian-vol-tf")
    assert strategy.definition.workflow_id is WorkflowId.RINCIAN_VOL_TF


def test_get_strategy_resolves_new_timeseries_ids():
    briva = get_strategy("timeseries-fbi-briva")
    assert briva.definition.workflow_id is WorkflowId.TIMESERIES_FBI_BRIVA
    qlola = get_strategy("timeseries-active-user-qlola")
    assert qlola.definition.workflow_id is WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA
    assert qlola.definition.requires_master_data is True


def test_get_strategy_unknown_id_raises():
    with pytest.raises(WorkflowValidationError):
        get_strategy("does-not-exist")


# --------------------------------------------------------------------------- #
# Aggregation exclusion (unit)
# --------------------------------------------------------------------------- #
def test_aggregation_excludes_wholesale_case_insensitive():
    df = pd.DataFrame(
        {
            "SEGMEN": ["wholesale", "Corporate", "", "WHOLESALE"],
            "MAINBR": ["B01", "B01", "B02", "B03"],
            "MBDESC": ["One", "One", "Two", "Three"],
            "AMOUNT_IN_IDR": [1000.0, 2000.0, 300.0, 500.0],
        }
    )
    result = AggregationEngine(exclude_segmen=["Wholesale"]).aggregate(
        df, group_cols=["MAINBR", "MBDESC"], value_col="AMOUNT_IN_IDR"
    )
    # B01 keeps only the Corporate row (2000); B02 keeps the blank-SEGMEN row
    # (300); B03 is dropped entirely (only Wholesale).
    assert result.branch_count == 2
    assert result.total_volume_idr == 2300.0


# --------------------------------------------------------------------------- #
# Rincian Vol TF workflow
# --------------------------------------------------------------------------- #
def test_rincian_vol_tf_excludes_wholesale_and_keeps_blank(rincian_vol_tf_file, tmp_path):
    out = tmp_path / "vol_tf.xlsx"
    config = ProcessingConfig(
        raw_data_path=rincian_vol_tf_file,
        workflow_id="rincian-vol-tf",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message
    assert out.exists()

    ws = load_workbook(out)["Summary_Report"]
    rows = {
        (r[0].value, r[1].value): r[2].value
        for r in ws.iter_rows()
        if r[0].value in {"B01", "B02", "B03"}
    }
    assert rows[("B01", "Branch One")] == 2000  # Wholesale row excluded
    assert rows[("B02", "Branch Two")] == 800   # blank SEGMEN row retained
    assert ("B03", "Branch Three") not in rows  # only-Wholesale branch dropped


def test_rincian_vol_tf_runs_without_reference_file(rincian_vol_tf_file, tmp_path):
    config = ProcessingConfig(
        raw_data_path=rincian_vol_tf_file,
        workflow_id="rincian-vol-tf",
        output_report_path=tmp_path / "no_ref.xlsx",
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message


# --------------------------------------------------------------------------- #
# Rincian Portal BG workflow
# --------------------------------------------------------------------------- #
def test_rincian_portal_bg_aggregates_without_filter(rincian_portal_bg_file, tmp_path):
    out = tmp_path / "portal_bg.xlsx"
    config = ProcessingConfig(
        raw_data_path=rincian_portal_bg_file,
        workflow_id="rincian-portal-bg",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    rows = {
        (r[0].value, r[1].value): r[2].value
        for r in ws.iter_rows()
        if r[0].value in {"B01", "B02"}
    }
    assert rows[("B01", "Branch One")] == 3000
    assert rows[("B02", "Branch Two")] == 500


# --------------------------------------------------------------------------- #
# Validation & E2E structure
# --------------------------------------------------------------------------- #
def test_missing_required_column_reports_failure(tmp_path):
    # Portal BG file missing the MBNAME column.
    bad = tmp_path / "bad.csv"
    bad.write_text("MAINBR|AMOUNT_IN_IDR\nB01|1000\n", encoding="utf-8")
    config = ProcessingConfig(
        raw_data_path=bad,
        workflow_id="rincian-portal-bg",
        output_report_path=tmp_path / "out.xlsx",
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is False
    assert "MBNAME" in (report.error_message or "")


def test_new_workflow_uses_detail_sheet_name(rincian_portal_bg_file, tmp_path):
    out = tmp_path / "sheets.xlsx"
    config = ProcessingConfig(
        raw_data_path=rincian_portal_bg_file,
        workflow_id="rincian-portal-bg",
        output_report_path=out,
    )
    PipelineOrchestrator.execute(config)
    wb = load_workbook(out)
    assert set(wb.sheetnames) == {"Summary_Report", "Detail_Data"}


def test_akumulasi_workflow_unchanged_end_to_end(raw_pipe_file, reference_file, tmp_path):
    """Regression: default workflow still produces the original two sheets."""
    out = tmp_path / "akum.xlsx"
    config = ProcessingConfig(
        raw_data_path=raw_pipe_file,
        reference_data_path=reference_file,
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True
    assert report.total_records_processed == 4
    assert report.unmapped_records_count == 1
    wb = load_workbook(out)
    assert set(wb.sheetnames) == {"Summary_Report", "Enriched_Data"}


# --------------------------------------------------------------------------- #
# Akumulasi Sum of FBI
# --------------------------------------------------------------------------- #
def _read_summary_table(ws):
    """Extracts (headers, {MAIN_CODE: row-dict}, grand-total-dict) from a sheet."""
    header_row = None
    for row in ws.iter_rows():
        if row[0].value == "MAIN_CODE":
            header_row = row[0].row
            break
    assert header_row is not None, "Summary table header not found"

    headers = []
    for cell in ws[header_row]:
        if cell.value is None:
            break
        headers.append(cell.value)

    rows: dict[str, dict] = {}
    grand_total: dict = {}
    r = header_row + 1
    while True:
        first = ws.cell(row=r, column=1).value
        if first is None:
            break
        record = {
            headers[i]: ws.cell(row=r, column=i + 1).value for i in range(len(headers))
        }
        if first == "Grand Total":
            grand_total = record
            break
        rows[first] = record
        r += 1
    return headers, rows, grand_total


def test_akumulasi_summary_has_sum_of_fbi_left_of_volume(
    raw_pipe_file, reference_file, tmp_path
):
    out = tmp_path / "akum_fbi.xlsx"
    config = ProcessingConfig(
        raw_data_path=raw_pipe_file,
        reference_data_path=reference_file,
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    headers, _, _ = _read_summary_table(ws)
    assert "Sum of FBI" in headers
    assert headers.index("Sum of FBI") == headers.index("VOLUME_IN_IDR") - 1


def test_akumulasi_sums_fbi_and_volume_with_grand_totals(
    raw_pipe_file, reference_file, tmp_path
):
    out = tmp_path / "akum_totals.xlsx"
    config = ProcessingConfig(
        raw_data_path=raw_pipe_file,
        reference_data_path=reference_file,
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    _, rows, grand_total = _read_summary_table(ws)

    # FBI branch sums export as plain integers (standard round).
    assert rows["MC10"]["Sum of FBI"] == 300
    assert rows["MC10"]["VOLUME_IN_IDR"] == 1_500_000_000_000
    assert rows["MC20"]["Sum of FBI"] == 50
    assert rows["MC20"]["VOLUME_IN_IDR"] == 500_000_000
    assert rows["UNMAPPED"]["Sum of FBI"] == 10

    # Both Grand Totals present as plain ints.
    assert grand_total["Sum of FBI"] == 360
    assert grand_total["VOLUME_IN_IDR"] == int(
        round(1_500_501_000_000.50)
    )


def test_akumulasi_missing_fbi_fails_with_clear_error(tmp_path):
    bad = tmp_path / "no_fbi.txt"
    bad.write_text(
        "KODE_UKER|SEGMEN|VOLUME_IN_IDR\n0001|Wholesale|1000\n", encoding="utf-8"
    )
    config = ProcessingConfig(
        raw_data_path=bad,
        workflow_id="akumulasi",
        output_report_path=tmp_path / "out.xlsx",
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is False
    assert "FBI" in (report.error_message or "")
    assert not (tmp_path / "out.xlsx").exists()


def test_ingestion_coerces_amount_column():
    engine = IngestionEngine(numeric_columns=("AMOUNT_IN_IDR",))
    # Round-trip a small file via a DataFrame-style check on coercion.
    df = pd.DataFrame({"AMOUNT_IN_IDR": ["1,000", "2000", "bad"]})
    coerced = engine._coerce_numeric_columns(df.copy())
    assert coerced["AMOUNT_IN_IDR"].tolist() == [1000.0, 2000.0, 0.0]


# --------------------------------------------------------------------------- #
# Reference enrichment extensions (raw alias + master ID->MAIN_CODE)
# --------------------------------------------------------------------------- #
def test_enrichment_resolves_raw_kode_uker_alias(reference_file):
    # Raw frame carries the spaced `KODE UKER` while the lookup key is KODE_UKER.
    raw = pd.DataFrame({"KODE UKER": ["0001", "0002"], "VOLUME_IDR": [1.0, 2.0]})
    result = ReferenceEnricher(reference_file).enrich(raw)
    assert "MAIN_CODE" in result.data.columns
    assert "MAIN_BRANCH" in result.data.columns
    assert result.unmapped_count == 0


def test_master_enricher_maps_id_to_main_code(qlola_master_file):
    mapping = MasterDataEnricher(qlola_master_file).build_id_to_main_code()
    assert mapping["U1"] == "7"
    assert mapping["U3"] == "9"
    assert "U5" not in mapping  # absent from master => later marked UNMAPPED


def test_master_enricher_unresolvable_columns_raises(tmp_path):
    bad = pd.DataFrame({"FOO": ["1"], "BAR": ["2"]})
    path = tmp_path / "master_bad.xlsx"
    bad.to_excel(path, index=False)
    with pytest.raises(ReferenceEnrichmentError) as exc:
        MasterDataEnricher(path).build_id_to_main_code()
    message = str(exc.value)
    assert "Aliases attempted" in message
    assert "Sheets scanned" in message


def test_master_enricher_master_data_statis_headers(tmp_path):
    """Task 2.1: fixture mimicking MASTER DATA STATIS MALANG workbook headers.

    Real file has columns like ID, MAIN_CODE (or KODE INDUK) — the alias
    registry must resolve them. This verifies the current ID_ALIASES and
    MAIN_CODE_ALIASES cover the master workbook without needing index-21 API.
    """
    master = pd.DataFrame({
        "ID": ["1001", "1002", "1003"],
        "MAIN_CODE": ["7", "9", "7"],
    })
    path = tmp_path / "master_statis_malang.xlsx"
    master.to_excel(path, sheet_name="MASTER DATA STATIS MALANG", index=False)
    mapping = MasterDataEnricher(path).build_id_to_main_code()
    assert mapping["1001"] == "7"
    assert mapping["1002"] == "9"
    assert mapping["1003"] == "7"


def test_master_enricher_ambiguous_two_sheet_picks_correct_sheet(tmp_path):
    """Task 2.3: ambiguous two-sheet master → correct sheet wins.

    Sheet A resolves aliases but maps everything to WRONG codes.
    Sheet B (named "MASTER DATA") is the real mapping table.
    Preference scoring must pick sheet B over sheet A.
    """
    wrong_sheet = pd.DataFrame({
        "ID": ["1001", "1002"],
        "KODE": ["WRONG_A", "WRONG_B"],  # KODE resolves as MAIN_CODE alias
    })
    correct_sheet = pd.DataFrame({
        "ID": ["1001", "1002"],
        "MAIN_CODE": ["7", "9"],
    })
    path = tmp_path / "master_ambiguous.xlsx"
    with pd.ExcelWriter(path) as writer:
        wrong_sheet.to_excel(writer, sheet_name="Cover", index=False)
        correct_sheet.to_excel(writer, sheet_name="MASTER DATA", index=False)
    mapping = MasterDataEnricher(path).build_id_to_main_code()
    # Must pick the "MASTER DATA" sheet, not the "Cover" sheet.
    assert mapping["1001"] == "7"
    assert mapping["1002"] == "9"


def test_master_enricher_picks_id_product_over_cif(tmp_path):
    """Live MASTER DATA STATIS layout: Qlola IDs in ID_PRODUCT, not CIF."""
    master = pd.DataFrame(
        {
            "ID_PRODUCT": ["100362", "100540", "100583"],
            "CIF": ["R555823", "TK34387", "Z064299"],
            "MAIN CODE": ["9", "429", "44"],
        }
    )
    path = tmp_path / "master_statis_real_layout.xlsx"
    master.to_excel(path, sheet_name="Sheet1", index=False)
    mapping = MasterDataEnricher(path).build_id_to_main_code(
        probe_ids=["100362", "100540"]
    )
    assert mapping["100362"] == "9"
    assert mapping["100540"] == "429"
    assert "R555823" not in mapping


def test_master_enricher_picks_id_column_matching_probe_ids(tmp_path):
    """Decoy ``ID`` (CIF/account codes) must lose to the real user-ID column.

    Mirrors live ``MASTER DATA STATIS MALANG`` workbooks where a generic ``ID``
    column holds values like R555823 while Qlola user IDs live in ``ID QLOLA``.
    """
    master = pd.DataFrame(
        {
            "ID": ["R555823", "TK34387", "Z064299", "M261301"],
            "ID QLOLA": ["100362", "100540", "100583", "100586"],
            "MAIN_CODE": ["73", "110", "9", "33"],
        }
    )
    path = tmp_path / "master_decoy_id.xlsx"
    master.to_excel(path, sheet_name="MASTER DATA STATIS MALANG", index=False)
    mapping = MasterDataEnricher(path).build_id_to_main_code(
        probe_ids=["100362", "100540", "100999"]
    )
    assert mapping["100362"] == "73"
    assert mapping["100540"] == "110"
    assert "R555823" not in mapping


def test_qlola_decoy_master_id_column_maps_via_probe_ids(
    qlola_raw_file_numeric_ids, qlola_uker_reference_file, tmp_path
):
    """End-to-end: master with decoy ID column still produces a real Summary."""
    master = pd.DataFrame(
        {
            "ID": ["R555823", "TK34387", "Z064299"],
            "ID QLOLA": ["1001", "1002", "1003"],
            "MAIN_CODE": ["7", "9", "7"],
        }
    )
    master_path = tmp_path / "master_decoy.xlsx"
    master.to_excel(master_path, sheet_name="MASTER DATA STATIS MALANG", index=False)
    out = tmp_path / "qlola_decoy.xlsx"
    report = _run_qlola(
        qlola_raw_file_numeric_ids, qlola_uker_reference_file, master_path, out
    )
    assert report.success is True, report.error_message
    assert report.unmapped_records_count == 0
    assert report.unmapped_diagnostic is None

    _, rows, _ = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    assert rows["7"]["AKTIF TRX >=5x"] == 2  # 1001 + 1003
    assert rows["9"]["TIDAK AKTIF <5x"] == 1  # 1002


def test_master_enricher_integration_live_like_master_qlola(
    tmp_path, qlola_uker_reference_file
):
    """Task 2.4: live-like master + Qlola sample → Summary has real MAIN_CODEs.

    Simulates a realistic scenario where the master has int-stored IDs
    (triggering the .0 float mismatch) and the raw file has string IDs.
    After canonicalization and preference scoring, the Summary must contain
    real MAIN_CODE rows, not only UNMAPPED.
    """
    # Master with int IDs (will become "1001.0" strings in Excel)
    master_df = pd.DataFrame({
        "ID": [1001, 1002, 1003],
        "MAIN_CODE": ["7", "9", "7"],
    })
    master_path = tmp_path / "master_live_like.xlsx"
    master_df.to_excel(master_path, sheet_name="MASTER DATA STATIS", index=False)

    # Qlola raw with string IDs
    raw_path = tmp_path / "qlola_live.csv"
    raw_path.write_text(
        "SOURCE|ID|FREKUENSI|KODE_UKER\n"
        "QCASH|1001|6|K1\n"
        "QCASH|1002|2|K1\n"
        "QCASH|1003|7|K2\n",
        encoding="utf-8",
    )
    out = tmp_path / "qlola_live.xlsx"
    report = _run_qlola(raw_path, qlola_uker_reference_file, master_path, out)
    assert report.success is True, report.error_message

    _, rows, _ = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    # Must have real MAIN_CODE rows, not only UNMAPPED.
    assert "7" in rows
    assert "9" in rows


# --------------------------------------------------------------------------- #
# Time Series FBI Briva
# --------------------------------------------------------------------------- #
def test_briva_nonwholesale_volume_matches_sample_grand_total(
    timeseries_briva_file, reference_file, tmp_path
):
    out = tmp_path / "briva.xlsx"
    config = ProcessingConfig(
        raw_data_path=timeseries_briva_file,
        reference_data_path=reference_file,
        workflow_id="timeseries-fbi-briva",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    headers, rows, grand_total = _read_summary_table(ws)

    # Value header renders as the sample Sheet1 label.
    assert "Sum of VOLUME_IDR" in headers
    # NONWHOLESALE branch totals (the WHOLESALE row is excluded), ceiled to ints.
    assert rows["MC10"]["Sum of VOLUME_IDR"] == 1_481_517_421_604
    assert rows["MC20"]["Sum of VOLUME_IDR"] == 1_000_000_000_000
    # Grand Total matches Excel ROUND/CEIL of the sample oracle constant.
    assert grand_total["Sum of VOLUME_IDR"] == 2_481_517_421_604


def _grand_total_cell(ws, header: str):
    """Returns the Grand Total row's cell under `header` on a summary sheet."""
    header_row = next(r[0].row for r in ws.iter_rows() if r[0].value == "MAIN_CODE")
    col = next(
        c.column for c in ws[header_row] if c.value == header
    )
    r = header_row + 1
    while ws.cell(row=r, column=1).value is not None:
        if ws.cell(row=r, column=1).value == "Grand Total":
            return ws.cell(row=r, column=col)
        r += 1
    raise AssertionError("Grand Total row not found")


def test_briva_grand_total_rounds_fractional_cents_up(
    timeseries_briva_file, reference_file, tmp_path
):
    """WF4 exports plain integers with fractional cents rounded up (ceil).

    The raw oracle sum is ``…603.75``; manual Excel ``ROUND(..., 0)`` and our
    ``ceil`` both land on ``…604`` — stored as a plain int with format ``0``.
    """
    out = tmp_path / "briva_precision.xlsx"
    config = ProcessingConfig(
        raw_data_path=timeseries_briva_file,
        reference_data_path=reference_file,
        workflow_id="timeseries-fbi-briva",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    cell = _grand_total_cell(
        load_workbook(out)["Summary_Report"], "Sum of VOLUME_IDR"
    )
    assert cell.value == 2_481_517_421_604
    assert isinstance(cell.value, int)
    assert cell.number_format == "0"


def test_briva_missing_reference_fails_clearly(timeseries_briva_file, tmp_path):
    config = ProcessingConfig(
        raw_data_path=timeseries_briva_file,
        workflow_id="timeseries-fbi-briva",
        output_report_path=tmp_path / "out.xlsx",
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is False
    assert "reference" in (report.error_message or "").lower()
    assert not (tmp_path / "out.xlsx").exists()


# --------------------------------------------------------------------------- #
# Time Series Active User Qlola
# --------------------------------------------------------------------------- #
def _read_crosstab_table(ws):
    """Extracts ({MAIN_CODE: {col: count}}, grand-total-dict) from a crosstab sheet."""
    header_row = None
    for row in ws.iter_rows():
        if row[0].value == "MAIN_CODE":
            header_row = row[0].row
            break
    assert header_row is not None, "Crosstab header not found"

    headers = []
    for cell in ws[header_row]:
        if cell.value is None:
            break
        headers.append(cell.value)

    rows: dict[str, dict] = {}
    grand_total: dict = {}
    r = header_row + 1
    while True:
        first = ws.cell(row=r, column=1).value
        if first is None:
            break
        record = {
            headers[i]: ws.cell(row=r, column=i + 1).value for i in range(len(headers))
        }
        if first == "Grand Total":
            grand_total = record
            break
        rows[first] = record
        r += 1
    return headers, rows, grand_total


def _run_qlola(qlola_raw_file, uker, master, out):
    config = ProcessingConfig(
        raw_data_path=qlola_raw_file,
        reference_data_path=uker,
        master_data_path=master,
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
    )
    return PipelineOrchestrator.execute(config)


def test_qlola_crosstab_uses_master_main_code_not_uker(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    headers, rows, grand_total = _read_crosstab_table(ws)

    assert headers == ["MAIN_CODE", "AKTIF TRX >=5x", "TIDAK AKTIF <5x", "Grand Total"]

    # Rows are keyed on MASTER MAIN_CODE (7 / 9), NOT the UKER code 99.
    # UNMAPPED is excluded from Summary rows (mapped-only). Detail retains it.
    assert set(rows) == {"7", "9"}
    assert "99" not in rows
    assert "UNMAPPED" not in rows

    # Master-keyed distribution (CMS excluded, FREKUENSI summed per ID, >=5 active):
    #   7: U1(7)=AKTIF, U2(2)+U4(1)=TIDAK  -> 1 / 2
    #   9: U3(5)=AKTIF                      -> 1 / 0  (U6 was CMS-only => dropped)
    #   (UNMAPPED: U5(9)=AKTIF — excluded from Summary)
    assert rows["7"]["AKTIF TRX >=5x"] == 1
    assert rows["7"]["TIDAK AKTIF <5x"] == 2
    assert rows["9"]["AKTIF TRX >=5x"] == 1
    assert rows["9"]["TIDAK AKTIF <5x"] == 0

    # Detail keeps UKER MAIN_CODE (reference) separate from master MAIN_CODE.
    detail_ws = load_workbook(out)["Enriched_Data"]
    headers = [detail_ws.cell(row=1, column=c).value for c in range(1, detail_ws.max_column + 1)]
    assert headers[:5] == [
        "ID",
        "UKER_MAIN_CODE",
        "MAIN_CODE",
        "FREKUENSI",
        "USER_AKTIF",
    ]
    uker_col = headers.index("UKER_MAIN_CODE") + 1
    master_col = headers.index("MAIN_CODE") + 1
    for r in range(2, detail_ws.max_row + 1):
        assert detail_ws.cell(row=r, column=uker_col).value == "99"
        master_code = detail_ws.cell(row=r, column=master_col).value
        assert master_code in {"7", "9", "UNMAPPED"}

    # Grand Totals exclude UNMAPPED (mapped-only: 4, not 5).
    assert grand_total["AKTIF TRX >=5x"] == 2
    assert grand_total["TIDAK AKTIF <5x"] == 2
    assert grand_total["Grand Total"] == 4


def test_qlola_missing_master_fails_clearly(
    qlola_raw_file, qlola_uker_reference_file, tmp_path
):
    out = tmp_path / "qlola_no_master.xlsx"
    config = ProcessingConfig(
        raw_data_path=qlola_raw_file,
        reference_data_path=qlola_uker_reference_file,
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is False
    assert "master" in (report.error_message or "").lower()
    assert not out.exists()


def test_qlola_missing_reference_fails_clearly(
    qlola_raw_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola_no_ref.xlsx"
    config = ProcessingConfig(
        raw_data_path=qlola_raw_file,
        master_data_path=qlola_master_file,
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is False
    assert "reference" in (report.error_message or "").lower()
    assert not out.exists()


# --------------------------------------------------------------------------- #
# Reference-side UKER alias resolution
# --------------------------------------------------------------------------- #
@pytest.fixture
def _uker_raw_frame() -> pd.DataFrame:
    return pd.DataFrame({"KODE_UKER": ["0001", "0002"], "VOLUME_IDR": [1.0, 2.0]})


def test_reference_uker_alias_resolves(reference_file_uker, _uker_raw_frame):
    result = ReferenceEnricher(reference_file_uker).enrich(_uker_raw_frame)
    assert result.unmapped_count == 0
    assert result.data["MAIN_CODE"].tolist() == ["MC10", "MC20"]
    assert result.data["MAIN_BRANCH"].tolist() == ["Jakarta Pusat", "Surabaya"]


def test_reference_spaced_kode_uker_alias_resolves(
    reference_file_kode_uker_spaced, _uker_raw_frame
):
    result = ReferenceEnricher(reference_file_kode_uker_spaced).enrich(_uker_raw_frame)
    assert result.unmapped_count == 0
    assert result.data["MAIN_CODE"].tolist() == ["MC10", "MC20"]


def test_reference_canonical_kode_uker_still_resolves(reference_file, _uker_raw_frame):
    """Regression: the canonical header keeps working after the alias widening."""
    result = ReferenceEnricher(reference_file).enrich(_uker_raw_frame)
    assert result.unmapped_count == 0
    assert result.data["MAIN_CODE"].tolist() == ["MC10", "MC20"]


def test_reference_unresolvable_error_lists_columns_and_aliases(
    reference_file_unresolvable, _uker_raw_frame
):
    with pytest.raises(ReferenceEnrichmentError) as exc:
        ReferenceEnricher(reference_file_unresolvable).enrich(_uker_raw_frame)
    message = str(exc.value)
    assert "Sheets scanned" in message
    assert "ACCOUNT NUMBER" in message  # available columns of the scanned sheet
    assert "Aliases attempted" in message
    for alias in ("KODE_UKER", "KODE UKER", "UKER", "KANCA"):
        assert alias in message


# --------------------------------------------------------------------------- #
# Raw-side join key normalisation (workflow ingest step)
# --------------------------------------------------------------------------- #
def test_normalize_join_key_renames_spaced_alias():
    raw = pd.DataFrame({"KODE UKER": ["0001"], "VOLUME_IDR": [1.0]})
    out = normalize_join_key(raw, "KODE_UKER")
    assert "KODE_UKER" in out.columns
    assert "KODE UKER" not in out.columns


def test_normalize_join_key_prefers_kode_uker_over_bare_uker():
    """Real Briva exports carry both; `UKER` holds branch *names*, not codes.

    Regression for the false-UNMAPPED collapse: resolving to `UKER` matched
    'KC Tulungagung' against numeric reference keys and mapped nothing.
    """
    raw = pd.DataFrame(
        {
            "KODE UKER": ["110", "13"],
            "UKER": ["KC Tulungagung", "KC Bondowoso"],
            "VOLUME_IDR": [1.0, 2.0],
        }
    )
    out = normalize_join_key(raw, "KODE_UKER")
    assert out["KODE_UKER"].tolist() == ["110", "13"]
    assert out["UKER"].tolist() == ["KC Tulungagung", "KC Bondowoso"]


def test_enricher_prefers_kode_uker_over_bare_uker_on_both_sides(tmp_path):
    """The same code-beats-name priority applies to the reference workbook."""
    ref = pd.DataFrame(
        {
            "KODE UKER": ["110", "13"],
            "UKER": ["KC Tulungagung", "KC Bondowoso"],
            "MAIN_CODE": ["110", "13"],
            "MAIN_BRANCH": ["KC Tulungagung", "KC Bondowoso"],
        }
    )
    ref_path = tmp_path / "reference_code_and_name.xlsx"
    ref.to_excel(ref_path, index=False)

    raw = pd.DataFrame(
        {"KODE UKER": ["110", "13"], "UKER": ["KC Tulungagung", "KC Bondowoso"]}
    )
    result = ReferenceEnricher(ref_path).enrich(normalize_join_key(raw, "KODE_UKER"))
    assert result.unmapped_count == 0
    assert result.data["MAIN_CODE"].tolist() == ["110", "13"]


def test_normalize_join_key_casts_integer_codes_to_string():
    raw = pd.DataFrame({"UKER": [1, 2], "VOLUME_IDR": [1.0, 2.0]})
    out = normalize_join_key(raw, "KODE_UKER")
    assert out["KODE_UKER"].tolist() == ["1", "2"]


def test_normalize_join_key_strips_padding_and_leaves_source_untouched():
    raw = pd.DataFrame({"KODE_UKER": ["  0001 ", "0002"]})
    out = normalize_join_key(raw, "KODE_UKER")
    assert out["KODE_UKER"].tolist() == ["0001", "0002"]
    # The caller's frame is never mutated in place.
    assert raw["KODE_UKER"].tolist() == ["  0001 ", "0002"]


def test_normalize_join_key_without_any_alias_is_a_noop():
    raw = pd.DataFrame({"BRANCH_ID": ["0001"]})
    out = normalize_join_key(raw, "KODE_UKER")
    assert list(out.columns) == ["BRANCH_ID"]


def test_briva_integer_branch_codes_still_map(reference_file):
    """Numeric-typed raw codes join against the string reference keys."""
    raw = pd.DataFrame({"KODE UKER": [1, 2], "VOLUME_IDR": [1.0, 2.0]})
    normalized = normalize_join_key(raw, "KODE_UKER")
    ref = pd.DataFrame(
        {
            "KODE_UKER": ["1", "2"],
            "MAIN_CODE": ["MC10", "MC20"],
            "MAIN_BRANCH": ["A", "B"],
        }
    )
    ref_path = reference_file.parent / "reference_numeric.xlsx"
    ref.to_excel(ref_path, index=False)
    result = ReferenceEnricher(ref_path).enrich(normalized)
    assert result.unmapped_count == 0


# --------------------------------------------------------------------------- #
# WF4 end-to-end mapping
# --------------------------------------------------------------------------- #
def _run_briva(raw, reference, out):
    config = ProcessingConfig(
        raw_data_path=raw,
        reference_data_path=reference,
        workflow_id="timeseries-fbi-briva",
        output_report_path=out,
    )
    return PipelineOrchestrator.execute(config)


def test_briva_underscore_header_matches_spaced_header_totals(
    timeseries_briva_file_underscore_key, reference_file, tmp_path
):
    out = tmp_path / "briva_underscore.xlsx"
    report = _run_briva(timeseries_briva_file_underscore_key, reference_file, out)
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_summary_table(load_workbook(out)["Summary_Report"])
    assert rows["MC10"]["Sum of VOLUME_IDR"] == 1_481_517_421_604
    assert rows["MC20"]["Sum of VOLUME_IDR"] == 1_000_000_000_000
    assert grand_total["Sum of VOLUME_IDR"] == 2_481_517_421_604


def test_briva_grand_total_equals_sum_of_data_rows(
    timeseries_briva_file, reference_file, tmp_path
):
    out = tmp_path / "briva_total.xlsx"
    report = _run_briva(timeseries_briva_file, reference_file, out)
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_summary_table(load_workbook(out)["Summary_Report"])
    row_sum = sum(r["Sum of VOLUME_IDR"] for r in rows.values())
    assert grand_total["Sum of VOLUME_IDR"] == pytest.approx(row_sum)


def test_briva_unmapped_rows_are_retained_not_dropped(
    timeseries_briva_file_with_unmapped, reference_file, tmp_path
):
    out = tmp_path / "briva_unmapped.xlsx"
    report = _run_briva(timeseries_briva_file_with_unmapped, reference_file, out)
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_summary_table(load_workbook(out)["Summary_Report"])
    # Mapped branches still resolve...
    assert rows["MC10"]["Sum of VOLUME_IDR"] == 100
    assert rows["MC20"]["Sum of VOLUME_IDR"] == 200
    # ...and the genuinely unknown branch code survives as UNMAPPED.
    assert rows["UNMAPPED"]["Sum of VOLUME_IDR"] == 50
    assert grand_total["Sum of VOLUME_IDR"] == 350
    assert report.unmapped_records_count == 1


# --------------------------------------------------------------------------- #
# WF5 ID -> MAIN_CODE lookup normalisation
# --------------------------------------------------------------------------- #
def test_qlola_numeric_master_ids_map_correctly(
    qlola_raw_file_numeric_ids,
    qlola_uker_reference_file,
    qlola_master_numeric_file,
    tmp_path,
):
    """Int-dtype master IDs vs. string raw IDs must not collapse to UNMAPPED."""
    out = tmp_path / "qlola_numeric.xlsx"
    report = _run_qlola(
        qlola_raw_file_numeric_ids,
        qlola_uker_reference_file,
        qlola_master_numeric_file,
        out,
    )
    assert report.success is True, report.error_message

    _, rows, _ = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    assert rows["7"]["AKTIF TRX >=5x"] == 1       # 1001, FREKUENSI 6
    assert rows["9"]["TIDAK AKTIF <5x"] == 1      # 1002, FREKUENSI 2
    # 1003 absent from master -> UNMAPPED, excluded from Summary but on detail.
    assert "UNMAPPED" not in rows
    # Detail sheet still has UNMAPPED IDs.
    detail_ws = load_workbook(out)["Enriched_Data"]
    detail_ids = [detail_ws.cell(row=r, column=1).value for r in range(2, detail_ws.max_row + 1)]
    assert "1003" in detail_ids


def test_master_enricher_strips_whitespace_padded_ids(tmp_path):
    padded = tmp_path / "master_padded.csv"
    padded.write_text("ID,MAIN_CODE\n  U1  ,  7 \nU2,9\n", encoding="utf-8")
    mapping = MasterDataEnricher(padded).build_id_to_main_code()
    assert mapping["U1"] == "7"
    assert mapping["U2"] == "9"


def test_qlola_padding_only_id_duplicates_count_as_one_user(
    tmp_path, qlola_uker_reference_file, qlola_master_file
):
    """End-to-end invariant: ' U1' and 'U1' are one user whose FREKUENSI sums.

    Ingestion already trims fields, and the strategy re-normalises IDs before
    the per-ID groupby; this pins the resulting behaviour either way.
    """
    raw = tmp_path / "qlola_padded_ids.csv"
    raw.write_text(
        "SOURCE|ID|FREKUENSI|KODE_UKER\nQCASH| U1 |3|K1\nQCASH|U1|4|K1\n",
        encoding="utf-8",
    )
    out = tmp_path / "qlola_padded.xlsx"
    report = _run_qlola(raw, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    # One user totalling 7 => a single AKTIF row, not two TIDAK AKTIF rows.
    assert rows["7"]["AKTIF TRX >=5x"] == 1
    assert grand_total["Grand Total"] == 1


def test_qlola_missing_master_id_becomes_unmapped(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola_unmapped.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    _, rows, _ = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    # U5 is absent from the master mapping -> UNMAPPED, excluded from Summary.
    assert "UNMAPPED" not in rows
    assert report.unmapped_records_count == 1

    # Detail sheet retains UNMAPPED IDs for audit.
    detail_ws = load_workbook(out)["Enriched_Data"]
    headers = [detail_ws.cell(row=1, column=c).value for c in range(1, detail_ws.max_column + 1)]
    main_code_col = headers.index("MAIN_CODE") + 1
    detail_codes = [
        detail_ws.cell(row=r, column=main_code_col).value
        for r in range(2, detail_ws.max_row + 1)
    ]
    assert "UNMAPPED" in detail_codes


# --------------------------------------------------------------------------- #
# WF5 mapped-only Summary (Tasks 3.3 / 3.4)
# --------------------------------------------------------------------------- #
def _read_metadata(ws) -> dict[str, object]:
    """Extracts the {label: value} metadata block above a Summary crosstab."""
    meta: dict[str, object] = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label == "MAIN_CODE":  # crosstab header reached
            break
        if isinstance(label, str) and label.endswith(":"):
            meta[label] = ws.cell(row=r, column=2).value
    return meta


def test_qlola_summary_metadata_reports_unmapped_id_count(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """Task 3.3: unmapped volume is surfaced in metadata, not as a Summary row."""
    out = tmp_path / "qlola_meta.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    meta = _read_metadata(ws)
    # U5 is the only ID absent from the master mapping.
    assert meta["Unmapped IDs:"] == "1"


def test_qlola_summary_metadata_omits_unmapped_line_when_all_mapped(
    qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """No `Unmapped IDs:` line when every ID resolves — avoids a noisy zero row."""
    raw = tmp_path / "qlola_all_mapped.csv"
    raw.write_text(
        "SOURCE|ID|FREKUENSI|KODE_UKER\nQCASH|U1|6|K1\nQCASH|U2|2|K2\n",
        encoding="utf-8",
    )
    out = tmp_path / "qlola_all_mapped.xlsx"
    report = _run_qlola(raw, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    meta = _read_metadata(load_workbook(out)["Summary_Report"])
    assert "Unmapped IDs:" not in meta


def test_qlola_mapped_only_aktif_total_excludes_unmapped_at_scale(
    qlola_uker_reference_file, tmp_path
):
    """Task 3.4: Clairine's dashboard number — 3,243 AKTIF − 204 unmapped = 3,039.

    Reproduces the live shape at scale: every ID is AKTIF (FREKUENSI >= 5), but
    204 of them are absent from the master mapping. The Summary Grand Total must
    report the mapped-only 3,039, while the detail sheet keeps all 3,243 IDs and
    the metadata reports the 204 that were dropped.
    """
    total_aktif = 3243
    unmapped_aktif = 204
    mapped_aktif = total_aktif - unmapped_aktif
    assert mapped_aktif == 3039

    ids = [f"{100000 + i}" for i in range(total_aktif)]
    mapped_ids = ids[:mapped_aktif]

    raw = tmp_path / "qlola_scale.csv"
    raw.write_text(
        "SOURCE|ID|FREKUENSI|KODE_UKER\n"
        + "".join(f"QCASH|{i}|7|K1\n" for i in ids),
        encoding="utf-8",
    )
    # Two master branches so the Summary has real multi-row structure.
    master = tmp_path / "qlola_scale_master.xlsx"
    pd.DataFrame(
        {
            "ID": mapped_ids,
            "MAIN_CODE": ["7" if n % 2 == 0 else "9" for n in range(mapped_aktif)],
        }
    ).to_excel(master, index=False)

    out = tmp_path / "qlola_scale.xlsx"
    report = _run_qlola(raw, qlola_uker_reference_file, master, out)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    _, rows, grand_total = _read_crosstab_table(ws)

    # Mapped-only rows and totals: the 204 unmapped IDs never reach the Summary.
    assert "UNMAPPED" not in rows
    assert set(rows) == {"7", "9"}
    assert grand_total["AKTIF TRX >=5x"] == 3039
    assert grand_total["TIDAK AKTIF <5x"] == 0
    assert grand_total["Grand Total"] == 3039

    # Metadata and the run result still account for the dropped IDs.
    assert _read_metadata(ws)["Unmapped IDs:"] == f"{unmapped_aktif:,}"
    assert report.unmapped_records_count == unmapped_aktif

    # Detail sheet retains the full 3,243-ID population for audit.
    detail_ws = load_workbook(out)["Enriched_Data"]
    assert detail_ws.max_row - 1 == total_aktif


# --------------------------------------------------------------------------- #
# WF5 SOURCE filter config
# --------------------------------------------------------------------------- #
def _run_qlola_with_source(raw, uker, master, out, source_exclude):
    config = ProcessingConfig(
        raw_data_path=raw,
        reference_data_path=uker,
        master_data_path=master,
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
        source_exclude=source_exclude,
    )
    return PipelineOrchestrator.execute(config)


def test_qlola_definition_declares_source_filter_defaults():
    definition = get_definition(WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA)
    assert definition.has_source_filter is True
    assert definition.source_exclude == ("CMS",)


def test_qlola_default_config_excludes_cms(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola_default_source.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    # U6 is CMS-only, so it never reaches any MAIN_CODE group.
    # U5 is UNMAPPED (excluded from Summary). Mapped-only total = 4.
    assert grand_total["Grand Total"] == 4
    assert rows["9"]["AKTIF TRX >=5x"] == 1  # U3 only; U6 dropped


def test_qlola_empty_source_exclude_keeps_every_row(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola_no_source_filter.xlsx"
    report = _run_qlola_with_source(
        qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out, []
    )
    assert report.success is True, report.error_message

    _, rows, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    # U6 (CMS-only, FREKUENSI 10) now counts under master code 9.
    assert rows["9"]["AKTIF TRX >=5x"] == 2
    # U5 is UNMAPPED (excluded from Summary). Mapped-only total = 5.
    assert grand_total["Grand Total"] == 5


def test_qlola_source_exclude_match_is_case_insensitive(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    out = tmp_path / "qlola_lowercase_source.xlsx"
    report = _run_qlola_with_source(
        qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out, ["cms"]
    )
    assert report.success is True, report.error_message

    _, _, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    assert grand_total["Grand Total"] == 4  # identical to the uppercase default (mapped-only)


def test_source_exclude_ignored_by_workflows_without_the_flag(
    rincian_portal_bg_file, tmp_path
):
    """A stale SOURCE value must not alter a workflow that does not declare it."""
    out = tmp_path / "portal_bg_with_source.xlsx"
    config = ProcessingConfig(
        raw_data_path=rincian_portal_bg_file,
        workflow_id="rincian-portal-bg",
        output_report_path=out,
        source_exclude=["CMS"],
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    rows = {
        (r[0].value, r[1].value): r[2].value
        for r in ws.iter_rows()
        if r[0].value in {"B01", "B02"}
    }
    assert rows[("B01", "Branch One")] == 3000
    assert rows[("B02", "Branch Two")] == 500


def test_existing_workflows_have_no_source_filter_defaults():
    for workflow_id in (
        WorkflowId.AKUMULASI,
        WorkflowId.RINCIAN_VOL_TF,
        WorkflowId.RINCIAN_PORTAL_BG,
        WorkflowId.TIMESERIES_FBI_BRIVA,
    ):
        definition = get_definition(workflow_id)
        assert definition.has_source_filter is False
        assert definition.source_exclude == ()


# --------------------------------------------------------------------------- #
# GUI SOURCE field / config propagation
# --------------------------------------------------------------------------- #
def test_gui_source_field_enabled_only_for_qlola():
    assert source_filter_enabled(WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA) is True
    for workflow_id in (
        WorkflowId.AKUMULASI,
        WorkflowId.RINCIAN_VOL_TF,
        WorkflowId.RINCIAN_PORTAL_BG,
        WorkflowId.TIMESERIES_FBI_BRIVA,
    ):
        assert source_filter_enabled(workflow_id) is False


def test_gui_source_field_default_text():
    assert default_source_text(WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA) == "CMS"
    assert default_source_text(WorkflowId.AKUMULASI) == ""


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("CMS", ["CMS"]),
        ("CMS,MOBILE", ["CMS", "MOBILE"]),
        ("  CMS , MOBILE  ", ["CMS", "MOBILE"]),
        ("", []),
        ("   ", []),
        ("CMS,,MOBILE", ["CMS", "MOBILE"]),
    ],
)
def test_gui_parse_source_exclude(text, expected):
    assert parse_source_exclude(text) == expected


def test_gui_edited_source_field_reaches_the_workflow(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """The parsed field value is what the strategy actually filters on."""
    strategy = get_strategy(WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA)
    config = ProcessingConfig(
        raw_data_path=qlola_raw_file,
        reference_data_path=qlola_uker_reference_file,
        master_data_path=qlola_master_file,
        workflow_id="timeseries-active-user-qlola",
        output_report_path=tmp_path / "unused.xlsx",
        source_exclude=parse_source_exclude("CMS, MOBILE"),
    )
    assert strategy.resolve_source_exclude(config) == ["CMS", "MOBILE"]


def test_workflow_without_flag_resolves_empty_source_exclude(tmp_path):
    strategy = get_strategy(WorkflowId.RINCIAN_PORTAL_BG)
    config = ProcessingConfig(
        raw_data_path=tmp_path / "raw.csv",
        workflow_id="rincian-portal-bg",
        output_report_path=tmp_path / "out.xlsx",
        source_exclude=["CMS"],
    )
    assert strategy.resolve_source_exclude(config) == []


# --------------------------------------------------------------------------- #
# UNMAPPED diagnostic warning
# --------------------------------------------------------------------------- #
def _warn_args(**overrides):
    args = {
        "unmapped_count": 10,
        "total": 10,
        "raw_key_sample": ["0001", "0002"],
        "reference_key_sample": ["A1", "A2"],
        "aliases_attempted": ["KODE_UKER", "KODE UKER"],
        "context": "Reference enrichment on 'KODE_UKER'",
    }
    args.update(overrides)
    return args


def test_full_unmapped_join_emits_diagnostic_warning(caplog):
    with caplog.at_level(logging.WARNING, logger="src.enrichment"):
        warned = warn_if_mostly_unmapped(**_warn_args())
    assert warned is True
    message = caplog.text
    assert "0001" in message            # raw key sample
    assert "A1" in message              # reference key sample
    assert "KODE UKER" in message       # aliases attempted
    assert "UNMAPPED" in message


def test_half_unmapped_join_does_not_warn(caplog):
    with caplog.at_level(logging.WARNING, logger="src.enrichment"):
        warned = warn_if_mostly_unmapped(**_warn_args(unmapped_count=5, total=10))
    assert warned is False
    assert caplog.text == ""


def test_empty_frame_does_not_warn():
    assert warn_if_mostly_unmapped(**_warn_args(unmapped_count=0, total=0)) is False


def test_unmapped_warning_does_not_abort_the_pipeline(
    timeseries_briva_file, tmp_path, caplog
):
    """A reference that matches nothing still produces a full report."""
    ref = pd.DataFrame(
        {
            "KODE_UKER": ["ZZZZ"],
            "MAIN_CODE": ["MC99"],
            "MAIN_BRANCH": ["Nowhere"],
        }
    )
    ref_path = tmp_path / "reference_no_overlap.xlsx"
    ref.to_excel(ref_path, index=False)

    out = tmp_path / "briva_all_unmapped.xlsx"
    with caplog.at_level(logging.WARNING, logger="src.enrichment"):
        report = _run_briva(timeseries_briva_file, ref_path, out)

    assert report.success is True, report.error_message
    assert out.exists()
    assert "resolved to UNMAPPED" in caplog.text

    _, rows, _ = _read_summary_table(load_workbook(out)["Summary_Report"])
    assert set(rows) == {"UNMAPPED"}


# --------------------------------------------------------------------------- #
# Operator-visible near-total UNMAPPED diagnostic (Tasks 4.1 / 4.2)
# --------------------------------------------------------------------------- #
@pytest.fixture
def qlola_master_no_overlap_file(tmp_path):
    """Master workbook whose IDs match none of the raw Qlola IDs."""
    path = tmp_path / "qlola_master_no_overlap.xlsx"
    pd.DataFrame({"ID": ["ZZ1", "ZZ2"], "MAIN_CODE": ["7", "9"]}).to_excel(
        path, index=False
    )
    return path


def test_qlola_total_unmapped_join_surfaces_operator_diagnostic(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_no_overlap_file, tmp_path
):
    """Task 4.1: a 100% UNMAPPED master join reaches the operator, not just the log."""
    out = tmp_path / "qlola_no_overlap.xlsx"
    report = _run_qlola(
        qlola_raw_file, qlola_uker_reference_file, qlola_master_no_overlap_file, out
    )
    assert report.success is True, report.error_message

    diagnostic = report.unmapped_diagnostic
    assert diagnostic is not None
    assert "U1" in diagnostic                       # raw ID sample
    assert "ZZ1" in diagnostic                      # master ID sample
    assert "ID QLOLA" in diagnostic                 # aliases attempted
    assert "UNMAPPED" in diagnostic

    # The GUI turns that into an actionable warning box.
    message = unmapped_warning_text(report)
    assert message is not None
    assert "master-data file" in message
    assert diagnostic in message


def test_qlola_total_unmapped_join_still_exports(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_no_overlap_file, tmp_path
):
    """Task 4.2 / PROGRESS-003: orphan keys never abort the export."""
    out = tmp_path / "qlola_no_overlap_export.xlsx"
    report = _run_qlola(
        qlola_raw_file, qlola_uker_reference_file, qlola_master_no_overlap_file, out
    )
    assert report.success is True, report.error_message
    assert out.exists()

    workbook = load_workbook(out)
    # Summary is mapped-only, so a fully-unmapped run has no data rows...
    _, rows, _ = _read_crosstab_table(workbook["Summary_Report"])
    assert rows == {}
    # ...but every ID is still on the detail sheet for audit.
    detail_ws = workbook["Enriched_Data"]
    assert detail_ws.max_row - 1 == 5  # U1..U5 post-CMS-filter


def test_qlola_detail_numbers_are_plain_integers(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """Task 6.1/6.2: FREKUENSI stores plain ints (format ``0``); IDs stay text.

    A float 7.0 under an integer format displays as `7` but round-trips through
    paste-values / CSV as `7.0`, which is not what the manual workbook shows.
    """
    out = tmp_path / "qlola_formats.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Enriched_Data"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    id_col = headers.index("ID") + 1
    frek_col = headers.index("FREKUENSI") + 1

    for r in range(2, ws.max_row + 1):
        frek = ws.cell(row=r, column=frek_col)
        assert isinstance(frek.value, int) and not isinstance(frek.value, bool)
        assert frek.number_format == "0"
        # IDs are labels, not measures: written as text with no number format.
        id_cell = ws.cell(row=r, column=id_col)
        assert isinstance(id_cell.value, str)
        assert id_cell.number_format == "General"

    # U1's two rows (3 + 4) summed to exactly 7 — an int, not 7.0.
    frek_values = [ws.cell(row=r, column=frek_col).value for r in range(2, ws.max_row + 1)]
    assert 7 in frek_values


def test_qlola_summary_counts_are_plain_integers(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """Crosstab counts are whole users; nothing renders as a formatted decimal."""
    out = tmp_path / "qlola_summary_formats.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    header_row = next(r[0].row for r in ws.iter_rows() if r[0].value == "MAIN_CODE")
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value is None:
            break
        for c in range(2, 5):  # AKTIF / TIDAK AKTIF / Grand Total
            cell = ws.cell(row=r, column=c)
            assert isinstance(cell.value, int)
            assert cell.number_format == "0"


def test_briva_detail_exports_ceiled_plain_integers(
    timeseries_briva_file, reference_file, tmp_path
):
    """WF4 detail VOLUME_IDR cells are ceiled ints with plain ``0`` format."""
    out = tmp_path / "briva_detail_formats.xlsx"
    report = PipelineOrchestrator.execute(
        ProcessingConfig(
            raw_data_path=timeseries_briva_file,
            reference_data_path=reference_file,
            workflow_id="timeseries-fbi-briva",
            output_report_path=out,
        )
    )
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Enriched_Data"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    vol_col = headers.index("VOLUME_IDR") + 1
    cells = [ws.cell(row=r, column=vol_col) for r in range(2, ws.max_row + 1)]
    assert all(c.number_format == "0" for c in cells)
    assert all(isinstance(c.value, int) for c in cells)
    # Fractional cents in the fixture must ceil, not truncate.
    assert any(c.value % 1 == 0 for c in cells)


def test_healthy_qlola_run_shows_no_operator_diagnostic(
    qlola_raw_file, qlola_uker_reference_file, qlola_master_file, tmp_path
):
    """One unmapped ID out of five is a normal miss rate, not a join failure."""
    out = tmp_path / "qlola_healthy.xlsx"
    report = _run_qlola(qlola_raw_file, qlola_uker_reference_file, qlola_master_file, out)
    assert report.success is True, report.error_message
    assert report.unmapped_diagnostic is None
    assert unmapped_warning_text(report) is None


def test_failed_run_shows_no_unmapped_warning(tmp_path):
    """A hard failure surfaces its own error box; no stale unmapped warning."""
    report = PipelineOrchestrator.execute(
        ProcessingConfig(
            raw_data_path=tmp_path / "missing.csv",
            workflow_id="timeseries-active-user-qlola",
            output_report_path=tmp_path / "out.xlsx",
        )
    )
    assert report.success is False
    assert unmapped_warning_text(report) is None


# --------------------------------------------------------------------------- #
# Real-sample acceptance oracles (Path A' — skip if vendor samples absent)
# --------------------------------------------------------------------------- #
def test_qlola_real_sample_sheet1_oracle(qlola_sample_inputs, tmp_path):
    """Full crosstab parity with the QLOLA sample Sheet1 (3000/1820/4820; 7->130/81)."""
    out = tmp_path / "qlola_real.xlsx"
    config = ProcessingConfig(
        raw_data_path=qlola_sample_inputs["raw"],
        reference_data_path=qlola_sample_inputs["uker"],
        master_data_path=qlola_sample_inputs["master"],
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    _, rows, grand_total = _read_crosstab_table(ws)

    # Per-MAIN_CODE parity (master ID->MAIN_CODE key, not UKER): code 7 -> 130/81.
    assert rows["7"]["AKTIF TRX >=5x"] == 130
    assert rows["7"]["TIDAK AKTIF <5x"] == 81
    # Sheet1 Grand Totals.
    assert grand_total["AKTIF TRX >=5x"] == 3000
    assert grand_total["TIDAK AKTIF <5x"] == 1820
    assert grand_total["Grand Total"] == 4820


def test_briva_real_sample_sheet1_oracle(briva_sample_inputs, tmp_path):
    """Grand Total parity with Excel ROUND/ceil of the Briva sample Sheet1 oracle."""
    out = tmp_path / "briva_real.xlsx"
    config = ProcessingConfig(
        raw_data_path=briva_sample_inputs["raw"],
        reference_data_path=briva_sample_inputs["uker"],
        workflow_id="timeseries-fbi-briva",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message

    ws = load_workbook(out)["Summary_Report"]
    headers, _, grand_total = _read_summary_table(ws)
    assert "Sum of VOLUME_IDR" in headers
    assert grand_total["Sum of VOLUME_IDR"] == 2_481_517_421_604


def test_briva_real_sample_rows_are_mapped_not_unmapped(briva_sample_inputs, tmp_path):
    """The Grand Total alone is mapping-blind: assert the branches resolved too.

    The false-UNMAPPED collapse this change fixes produced a *correct* Grand
    Total on a single UNMAPPED row, so parity on the total proves nothing here.
    """
    out = tmp_path / "briva_real_mapping.xlsx"
    config = ProcessingConfig(
        raw_data_path=briva_sample_inputs["raw"],
        reference_data_path=briva_sample_inputs["uker"],
        workflow_id="timeseries-fbi-briva",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message
    assert report.unmapped_records_count == 0

    _, rows, _ = _read_summary_table(load_workbook(out)["Summary_Report"])
    assert len(rows) > 1
    assert "UNMAPPED" not in rows
    # MAIN_BRANCH must carry real branch names, not the UNMAPPED sentinel.
    assert all(r["MAIN_BRANCH"] not in (None, "", "UNMAPPED") for r in rows.values())


def test_qlola_real_sample_rows_are_mapped_not_unmapped(qlola_sample_inputs, tmp_path):
    """Companion mapping assertion for the Qlola crosstab (see the Briva twin)."""
    out = tmp_path / "qlola_real_mapping.xlsx"
    config = ProcessingConfig(
        raw_data_path=qlola_sample_inputs["raw"],
        reference_data_path=qlola_sample_inputs["uker"],
        master_data_path=qlola_sample_inputs["master"],
        workflow_id="timeseries-active-user-qlola",
        output_report_path=out,
    )
    report = PipelineOrchestrator.execute(config)
    assert report.success is True, report.error_message
    assert report.unmapped_records_count == 0

    _, rows, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    assert len(rows) > 1
    assert "UNMAPPED" not in rows
    # Column sums reconcile with the Grand Total row.
    assert sum(r["AKTIF TRX >=5x"] for r in rows.values()) == grand_total["AKTIF TRX >=5x"]
    assert (
        sum(r["TIDAK AKTIF <5x"] for r in rows.values())
        == grand_total["TIDAK AKTIF <5x"]
    )


# --------------------------------------------------------------------------- #
# Live acceptance (operator Downloads folder — skipped in CI)
# --------------------------------------------------------------------------- #
_LIVE_QLola_JUL24_RAW = Path(
    r"C:\Users\Draven Chen\Downloads"
    r"\1785047437_REPORT_SUMMARY_MONTHLY_2026-07-24_QLOLA_ALL.csv"
)
_LIVE_QLola_JUL24_UKER = Path(
    r"C:\Users\Draven Chen\Downloads"
    r"\Mapping GL FBI Vol Uker (BRIEFX) Update BO BATU.xlsx"
)
_LIVE_QLola_JUL24_MASTER = Path(
    r"C:\Users\Draven Chen\Downloads\MASTER DATA STATIS MALANG JUN (3).xlsx"
)


@pytest.mark.skipif(
    not _LIVE_QLola_JUL24_RAW.exists()
    or not _LIVE_QLola_JUL24_UKER.exists()
    or not _LIVE_QLola_JUL24_MASTER.exists(),
    reason="Jul-24 Qlola acceptance files not in Downloads",
)
def test_qlola_jul24_live_acceptance_oracle(tmp_path):
    """Parity with manual Sheet1 pivot for the Jul-24 QLOLA_ALL operator run."""
    out = tmp_path / "qlola_jul24_live.xlsx"
    report = _run_qlola(
        _LIVE_QLola_JUL24_RAW, _LIVE_QLola_JUL24_UKER, _LIVE_QLola_JUL24_MASTER, out
    )
    assert report.success is True, report.error_message
    assert report.unmapped_records_count == 372
    assert report.unmapped_diagnostic is None

    _, rows, grand_total = _read_crosstab_table(load_workbook(out)["Summary_Report"])
    assert len(rows) == 25
    assert rows["7"]["AKTIF TRX >=5x"] == 139
    assert rows["7"]["TIDAK AKTIF <5x"] == 75
    # Mapped-only Grand Totals (dashboard contract: exclude #N/A / UNMAPPED).
    assert grand_total["AKTIF TRX >=5x"] == 3039
    assert grand_total["TIDAK AKTIF <5x"] == 1626
    assert grand_total["Grand Total"] == 4665

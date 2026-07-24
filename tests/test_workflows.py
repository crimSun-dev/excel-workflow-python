"""Multi-workflow tests: strategy dispatch, new workflows, and E2E pipeline."""

from __future__ import annotations

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.aggregation import AggregationEngine
from src.ingestion import IngestionEngine
from src.orchestrator import PipelineOrchestrator
from src.schemas import ProcessingConfig
from src.workflows.base import WorkflowId, WorkflowValidationError
from src.workflows.registry import WORKFLOW_REGISTRY, get_strategy


# --------------------------------------------------------------------------- #
# Registry / dispatch
# --------------------------------------------------------------------------- #
def test_registry_has_all_three_workflows():
    assert set(WORKFLOW_REGISTRY) == {
        WorkflowId.AKUMULASI,
        WorkflowId.RINCIAN_VOL_TF,
        WorkflowId.RINCIAN_PORTAL_BG,
    }


def test_get_strategy_accepts_string_id():
    strategy = get_strategy("rincian-vol-tf")
    assert strategy.definition.workflow_id is WorkflowId.RINCIAN_VOL_TF


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

    # FBI branch sums: MC10 = 100+200, MC20 = 50.25, UNMAPPED = 10.
    assert rows["MC10"]["Sum of FBI"] == 300.0
    assert rows["MC10"]["VOLUME_IN_IDR"] == 1_500_000_000_000.0
    assert rows["MC20"]["Sum of FBI"] == 50.25
    assert rows["MC20"]["VOLUME_IN_IDR"] == 500_000_000.50
    assert rows["UNMAPPED"]["Sum of FBI"] == 10.0

    # Both Grand Totals present.
    assert grand_total["Sum of FBI"] == 360.25
    assert grand_total["VOLUME_IN_IDR"] == round(1_500_501_000_000.50, 2)


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

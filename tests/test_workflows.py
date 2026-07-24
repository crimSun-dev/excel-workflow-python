"""Multi-workflow tests: strategy dispatch, new workflows, and E2E pipeline."""

from __future__ import annotations

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.aggregation import AggregationEngine
from src.enrichment import (
    MasterDataEnricher,
    ReferenceEnricher,
    ReferenceEnrichmentError,
)
from src.ingestion import IngestionEngine
from src.orchestrator import PipelineOrchestrator
from src.schemas import ProcessingConfig
from src.workflows.base import WorkflowId, WorkflowValidationError
from src.workflows.registry import WORKFLOW_REGISTRY, get_strategy


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
    # NONWHOLESALE branch totals (the WHOLESALE row is excluded).
    assert rows["MC10"]["Sum of VOLUME_IDR"] == 1_481_517_421_603.75
    assert rows["MC20"]["Sum of VOLUME_IDR"] == 1_000_000_000_000.0
    # Grand Total matches the sample oracle constant.
    assert grand_total["Sum of VOLUME_IDR"] == 2_481_517_421_603.75


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

    # Rows are keyed on MASTER MAIN_CODE (7 / 9 / UNMAPPED), NOT the UKER code 99.
    assert set(rows) == {"7", "9", "UNMAPPED"}
    assert "99" not in rows

    # Master-keyed distribution (CMS excluded, FREKUENSI summed per ID, >=5 active):
    #   7: U1(7)=AKTIF, U2(2)+U4(1)=TIDAK  -> 1 / 2
    #   9: U3(5)=AKTIF                      -> 1 / 0  (U6 was CMS-only => dropped)
    #   UNMAPPED: U5(9)=AKTIF              -> 1 / 0
    assert rows["7"]["AKTIF TRX >=5x"] == 1
    assert rows["7"]["TIDAK AKTIF <5x"] == 2
    assert rows["9"]["AKTIF TRX >=5x"] == 1
    assert rows["9"]["TIDAK AKTIF <5x"] == 0
    assert rows["UNMAPPED"]["AKTIF TRX >=5x"] == 1

    # Grand Totals for this fixture.
    assert grand_total["AKTIF TRX >=5x"] == 3
    assert grand_total["TIDAK AKTIF <5x"] == 2
    assert grand_total["Grand Total"] == 5


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

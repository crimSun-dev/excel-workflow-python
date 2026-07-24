"""Time Series Active User Qlola workflow.

A multi-stage pipeline that does not fit the shared single groupby-sum template,
so it overrides `execute` while still reusing `IngestionEngine`,
`ReferenceEnricher`, `MasterDataEnricher`, the `ExcelReportExporter` crosstab
mode, and the orchestrator exception boundary.

Stages:
    ingest -> exclude SOURCE=CMS -> UKER enrich (detail context) ->
    sum FREKUENSI per ID -> master ID->MAIN_CODE lookup ->
    categorize AKTIF/TIDAK at FREKUENSI >= 5 ->
    crosstab Count of distinct ID by MAIN_CODE x USER_AKTIF.

The final crosstab is keyed on the *master* MAIN_CODE, never the UKER-enriched
MAIN_CODE: sample Sheet1 evidence shows UKER MAIN_CODE matches global AKTIF/TIDAK
totals but fails per-branch counts.
"""

from __future__ import annotations

import pandas as pd

from ..enrichment import UNMAPPED_SENTINEL, MasterDataEnricher, ReferenceEnricher
from ..exporter import ExcelReportExporter
from ..ingestion import IngestionEngine
from ..schemas import ProcessingConfig
from .base import (
    WorkflowDefinition,
    WorkflowId,
    WorkflowRunResult,
    WorkflowStrategy,
    WorkflowValidationError,
)

ACTIVE_LABEL = "AKTIF TRX >=5x"
INACTIVE_LABEL = "TIDAK AKTIF <5x"
USER_AKTIF_CATEGORIES = [ACTIVE_LABEL, INACTIVE_LABEL]
ACTIVE_THRESHOLD = 5

TIMESERIES_ACTIVE_USER_QLOLA_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA,
    label="Time Series Active User Qlola",
    requires_reference=True,
    requires_master_data=True,
    group_cols=("MAIN_CODE",),
    value_col="FREKUENSI",
    report_title="Time Series Active User Qlola Report",
    detail_sheet_name="Enriched_Data",
    number_format="#,##0",
    numeric_columns=("FREKUENSI",),
    required_columns=("SOURCE", "ID", "FREKUENSI"),
    exclude_source=("CMS",),
)


class TimeSeriesActiveUserQlolaStrategy(WorkflowStrategy):
    """Active-vs-inactive Qlola user crosstab keyed on master MAIN_CODE."""

    @property
    def definition(self) -> WorkflowDefinition:
        return TIMESERIES_ACTIVE_USER_QLOLA_DEFINITION

    def execute(self, config: ProcessingConfig) -> WorkflowRunResult:
        definition = self.definition

        # 0. Both reference inputs are mandatory; fail before doing any work.
        if config.reference_data_path is None:
            raise WorkflowValidationError(
                "The 'Time Series Active User Qlola' workflow requires a UKER "
                "reference mapping file, but none was provided."
            )
        if config.master_data_path is None:
            raise WorkflowValidationError(
                "The 'Time Series Active User Qlola' workflow requires a "
                "master-data file (ID -> MAIN_CODE), but none was provided."
            )

        # 1. Ingest raw pipe-delimited data, coercing FREKUENSI to numeric.
        ingestion = IngestionEngine(
            delimiter=config.delimiter,
            numeric_columns=definition.numeric_columns or ("FREKUENSI",),
        )
        ingested = ingestion.read_raw_data(config.raw_data_path)
        data = ingested.data

        # 2. Validate mandatory input columns.
        self._validate_columns(data)

        # 3. Exclude SOURCE=CMS (case-insensitive); QCASH/QIB rows are retained.
        excluded = {s.strip().casefold() for s in definition.exclude_source}
        source = data["SOURCE"].astype(str).str.strip().str.casefold()
        data = data[~source.isin(excluded)]

        # 4. UKER enrich for detail context (not the crosstab key).
        enricher = ReferenceEnricher(
            reference_path=config.reference_data_path,
            lookup_key=config.lookup_key,
        )
        enriched = enricher.enrich(data).data

        # 5. Sum FREKUENSI per unique ID.
        per_id = (
            enriched.groupby("ID", as_index=False, dropna=False)["FREKUENSI"]
            .sum()
            .reset_index(drop=True)
        )
        per_id["ID"] = per_id["ID"].astype(str).str.strip()

        # 6. Master ID -> MAIN_CODE lookup; unmatched IDs kept as UNMAPPED.
        id_to_code = MasterDataEnricher(config.master_data_path).build_id_to_main_code()
        per_id["MAIN_CODE"] = (
            per_id["ID"].map(id_to_code).fillna(UNMAPPED_SENTINEL)
        )
        unmapped_count = int((per_id["MAIN_CODE"] == UNMAPPED_SENTINEL).sum())

        # 7. Categorize active vs inactive at the >= 5 threshold.
        per_id["USER_AKTIF"] = per_id["FREKUENSI"].apply(
            lambda f: ACTIVE_LABEL if f >= ACTIVE_THRESHOLD else INACTIVE_LABEL
        )

        # 8. Crosstab: Count of distinct ID by MAIN_CODE (rows) x USER_AKTIF (cols).
        crosstab = self._build_crosstab(per_id)

        exporter = ExcelReportExporter(
            number_format=definition.number_format,
            value_column="FREKUENSI",
            report_title=definition.report_title,
            detail_sheet_name=definition.detail_sheet_name,
        )
        output_path = exporter.export_crosstab(
            crosstab_df=crosstab,
            detail_df=per_id[["ID", "FREKUENSI", "MAIN_CODE", "USER_AKTIF"]],
            output_path=config.output_report_path,
            row_label_header="MAIN_CODE",
            category_columns=USER_AKTIF_CATEGORIES,
            total_records=ingested.total_rows,
            filter_applied=f"{', '.join(definition.exclude_source)} (SOURCE excluded)",
        )

        return WorkflowRunResult(
            output_path=output_path,
            total_records_processed=ingested.total_rows,
            unmapped_records_count=unmapped_count,
        )

    @staticmethod
    def _build_crosstab(per_id: pd.DataFrame) -> pd.DataFrame:
        """Counts distinct IDs per MAIN_CODE x USER_AKTIF as a tidy frame.

        Every category column is present even when empty, and MAIN_CODE rows
        (including the UNMAPPED row) are retained and sorted for stable output.
        """
        counts = pd.crosstab(per_id["MAIN_CODE"], per_id["USER_AKTIF"])
        for category in USER_AKTIF_CATEGORIES:
            if category not in counts.columns:
                counts[category] = 0
        counts = counts[USER_AKTIF_CATEGORIES]
        counts = counts.sort_index().reset_index()
        counts.columns = ["MAIN_CODE", *USER_AKTIF_CATEGORIES]
        return counts

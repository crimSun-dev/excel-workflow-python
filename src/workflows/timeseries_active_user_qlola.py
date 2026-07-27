"""Time Series Active User Qlola workflow.

A multi-stage pipeline that does not fit the shared single groupby-sum template,
so it overrides `execute` while still reusing `IngestionEngine`,
`ReferenceEnricher`, `MasterDataEnricher`, the `ExcelReportExporter` crosstab
mode, and the orchestrator exception boundary.

Stages:
    ingest -> exclude configured SOURCE values (default CMS) ->
    UKER enrich (detail context) -> sum FREKUENSI per ID ->
    master ID->MAIN_CODE lookup -> categorize AKTIF/TIDAK at FREKUENSI >= 5 ->
    crosstab Count of distinct ID by MAIN_CODE x USER_AKTIF.

The SOURCE exclusion list comes from the workflow definition (`source_exclude`)
and can be overridden at runtime via `ProcessingConfig.source_exclude` (the GUI
SOURCE field); an empty list disables SOURCE filtering entirely.

The final crosstab is keyed on the *master* MAIN_CODE, never the UKER-enriched
MAIN_CODE: sample Sheet1 evidence shows UKER MAIN_CODE matches global AKTIF/TIDAK
totals but fails per-branch counts.
"""

from __future__ import annotations

import pandas as pd

from ..enrichment import (
    ID_ALIASES,
    UNMAPPED_SENTINEL,
    MasterDataEnricher,
    ReferenceEnricher,
    build_unmapped_diagnostic,
    canonicalize_join_id,
    warn_if_mostly_unmapped,
)
from ..exporter import ExcelReportExporter
from ..ingestion import IngestionEngine
from ..schemas import ProcessingConfig
from .base import (
    WorkflowDefinition,
    WorkflowId,
    WorkflowRunResult,
    WorkflowStrategy,
    WorkflowValidationError,
    normalize_join_key,
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
    numeric_columns=("FREKUENSI",),
    required_columns=("SOURCE", "ID", "FREKUENSI"),
    source_exclude=("CMS",),
    has_source_filter=True,
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

        # 3. Exclude the configured SOURCE values (default CMS, case-insensitive);
        # a GUI/CLI override replaces the default, and an empty list disables it.
        source_exclude = self.resolve_source_exclude(config)
        data = self.apply_source_exclude(data, source_exclude)

        # 4. UKER enrich for detail context (not the crosstab key). The branch
        # code is canonicalized here, at ingest, before the generic enricher runs.
        data = normalize_join_key(data, config.lookup_key)
        enricher = ReferenceEnricher(
            reference_path=config.reference_data_path,
            lookup_key=config.lookup_key,
        )
        enriched = enricher.enrich(data).data
        # UKER MAIN_CODE is for detail / branch context only — the Summary crosstab
        # keys on a separate master ID→MAIN_CODE VLOOKUP below.
        enriched = enriched.rename(
            columns={
                "MAIN_CODE": "UKER_MAIN_CODE",
                "MAIN_BRANCH": "UKER_MAIN_BRANCH",
            }
        )

        # 5. Sum FREKUENSI per unique ID. IDs are normalized *before* the groupby
        # so ' U1' and 'U1' collapse into one user rather than two.
        enriched["ID"] = enriched["ID"].apply(canonicalize_join_id)
        per_id = (
            enriched.groupby("ID", as_index=False)
            .agg(
                FREKUENSI=("FREKUENSI", "sum"),
                UKER_MAIN_CODE=("UKER_MAIN_CODE", "first"),
            )
            .reset_index(drop=True)
        )

        # 6. Master ID -> MAIN_CODE lookup (VLOOKUP on ID); unmatched -> UNMAPPED.
        # Probe with the aggregated raw IDs so the enricher picks the column whose
        # keys actually match the Qlola user IDs, not a decoy ID/CIF column.
        id_to_code = MasterDataEnricher(config.master_data_path).build_id_to_main_code(
            probe_ids=per_id["ID"].tolist()
        )
        per_id["MAIN_CODE"] = (
            per_id["ID"].map(id_to_code).fillna(UNMAPPED_SENTINEL)
        )
        unmapped_count = int((per_id["MAIN_CODE"] == UNMAPPED_SENTINEL).sum())
        diagnostic_args = {
            "unmapped_count": unmapped_count,
            "total": len(per_id),
            "raw_key_sample": per_id["ID"].head(5).tolist(),
            "reference_key_sample": list(id_to_code)[:5],
            "aliases_attempted": list(ID_ALIASES),
            "context": "Qlola master ID -> MAIN_CODE lookup",
        }
        warned = warn_if_mostly_unmapped(**diagnostic_args)
        # A near-total UNMAPPED join is a join *failure*, not a normal miss rate,
        # so the text is handed back for the GUI to show. The run still exports.
        diagnostic = build_unmapped_diagnostic(**diagnostic_args) if warned else None

        # 7. Categorize active vs inactive at the >= 5 threshold.
        per_id["USER_AKTIF"] = per_id["FREKUENSI"].apply(
            lambda f: ACTIVE_LABEL if f >= ACTIVE_THRESHOLD else INACTIVE_LABEL
        )

        # 8. Crosstab: Count of distinct ID by MAIN_CODE (rows) x USER_AKTIF (cols).
        # Summary excludes UNMAPPED / blank / #N/A rows; detail keeps all IDs.
        crosstab = self._build_crosstab(per_id)

        exporter = ExcelReportExporter(
            number_format=definition.number_format,
            value_column="FREKUENSI",
            report_title=definition.report_title,
            detail_sheet_name=definition.detail_sheet_name,
        )
        output_path = exporter.export_crosstab(
            crosstab_df=crosstab,
            detail_df=per_id[
                ["ID", "UKER_MAIN_CODE", "MAIN_CODE", "FREKUENSI", "USER_AKTIF"]
            ],
            output_path=config.output_report_path,
            row_label_header="MAIN_CODE",
            category_columns=USER_AKTIF_CATEGORIES,
            total_records=ingested.total_rows,
            filter_applied=(
                f"{', '.join(source_exclude)} (SOURCE excluded)"
                if source_exclude
                else None
            ),
            unmapped_ids_count=unmapped_count,
        )

        return WorkflowRunResult(
            output_path=output_path,
            total_records_processed=ingested.total_rows,
            unmapped_records_count=unmapped_count,
            unmapped_warning_emitted=warned,
            unmapped_diagnostic=diagnostic,
        )

    @staticmethod
    def _build_crosstab(per_id: pd.DataFrame) -> pd.DataFrame:
        """Counts distinct IDs per MAIN_CODE x USER_AKTIF as a tidy frame.

        Rows whose MAIN_CODE is UNMAPPED, blank, or '#N/A' are excluded from
        the Summary_Report so that Grand Totals reflect mapped-only counts.
        The detail sheet separately retains all IDs including unmapped.
        """
        # Exclude UNMAPPED / blank / #N/A from Summary rows.
        _EXCLUDED = {UNMAPPED_SENTINEL, "", "#N/A"}
        mapped = per_id[~per_id["MAIN_CODE"].isin(_EXCLUDED)]

        if mapped.empty:
            # All IDs unmapped: return an empty frame with the right columns.
            return pd.DataFrame(columns=["MAIN_CODE", *USER_AKTIF_CATEGORIES])

        counts = pd.crosstab(mapped["MAIN_CODE"], mapped["USER_AKTIF"])
        for category in USER_AKTIF_CATEGORIES:
            if category not in counts.columns:
                counts[category] = 0
        counts = counts[USER_AKTIF_CATEGORIES]
        counts = counts.sort_index().reset_index()
        counts.columns = ["MAIN_CODE", *USER_AKTIF_CATEGORIES]
        return counts


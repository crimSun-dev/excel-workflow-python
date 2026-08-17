"""Report Data Statis workflow.

Ingests raw pipe-delimited data (commonly a `.bin` export, which is plain
delimited text), enriches the unit code against a mapping workbook (VLOOKUP),
drops SEGMEN `KORPORASI`, keeps only KW `KANWIL MALANG`, then reports the
Excel PivotTable **Count of ID_PRODUCT** by MAIN_CODE + MAIN_BRANCH.

The region lives in the extract's `KW` column, so it is filtered through the one
operator-facing KW dimension rather than a separate KAWIL flag - only the *value*
is regional ("KANWIL MALANG"). Both that and the KORPORASI exclusion are the
defaults the operator gets for free, and both stay editable in the GUI/CLI
(matching the manual pivot, where unticking values is a per-run decision).
`aggfunc="count"` keeps the textual ID_PRODUCT column out of the money-summing
path. Everything else runs on the shared ingest -> enrich -> aggregate -> export
template.
"""

from __future__ import annotations

import pandas as pd

from ..enrichment import ReferenceEnricher
from ..schemas import ProcessingConfig
from .base import (
    WorkflowDefinition,
    WorkflowId,
    WorkflowStrategy,
    WorkflowValidationError,
    normalize_join_key,
)

REPORT_DATA_STATIS_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.REPORT_DATA_STATIS,
    label="Report Data Statis",
    requires_reference=True,
    group_cols=("MAIN_CODE", "MAIN_BRANCH"),
    value_col="ID_PRODUCT",
    report_title="Report Data Statis",
    detail_sheet_name="Enriched_Data",
    # ID_PRODUCT is a discrete identifier, never a monetary amount, so it is
    # deliberately absent from numeric_columns - coercing it would destroy
    # non-numeric IDs and turn the Count into a Sum of nonsense.
    required_columns=("SEGMEN", "KW", "ID_PRODUCT"),
    aggfunc="count",
    exclude_segmen=("KORPORASI",),
    # KORPORASI is only the *default* exclusion: the operator can add segments or
    # clear the field to keep corporate accounts in a one-off run.
    supports_segment_filter=True,
    has_source_filter=True,
    has_kw_filter=True,
    # The regional cut is a KW keep-list, not a KAWIL dimension of its own.
    kw_include=("KANWIL MALANG",),
    value_display_names=(("ID_PRODUCT", "Count of ID_PRODUCT"),),
)


class ReportDataStatisStrategy(WorkflowStrategy):
    """Count of ID_PRODUCT per branch for KANWIL MALANG, excluding KORPORASI."""

    @property
    def definition(self) -> WorkflowDefinition:
        return REPORT_DATA_STATIS_DEFINITION

    def _enrich(
        self, config: ProcessingConfig, data: pd.DataFrame
    ) -> tuple[pd.DataFrame, int]:
        if config.reference_data_path is None:
            raise WorkflowValidationError(
                "The 'Report Data Statis' workflow requires a reference "
                "mapping file, but none was provided."
            )
        # Raw exports spell the unit code as `KODE_UNIT`, `KODE UNIT` or
        # `KODE_UKER`; canonicalize it so the join never depends on a VLOOKUP
        # column index.
        data = normalize_join_key(data, config.lookup_key)
        enricher = ReferenceEnricher(
            reference_path=config.reference_data_path,
            lookup_key=config.lookup_key,
        )
        result = enricher.enrich(data)
        return result.data, result.unmapped_count

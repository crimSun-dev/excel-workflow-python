"""Report Summary Akumulasi workflow (the original pipeline).

Ingests raw data, enriches it against a reference workbook (VLOOKUP), then
aggregates by MAIN_CODE + MAIN_BRANCH summing VOLUME_IN_IDR into the existing
two-sheet workbook (Summary_Report + Enriched_Data).
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
)

AKUMULASI_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.AKUMULASI,
    label="Report Summary Akumulasi",
    requires_reference=True,
    group_cols=("MAIN_CODE", "MAIN_BRANCH"),
    value_col="VOLUME_IN_IDR",
    report_title="Financial Summary Report",
    detail_sheet_name="Enriched_Data",
    number_format="#,##0.00",
    numeric_columns=("FBI", "VOLUME_IN_IDR"),
    required_columns=("FBI", "VOLUME_IN_IDR"),
    supports_segment_filter=True,
    # Sum of FBI sits immediately left of volume to mirror the QLOLA pivot.
    value_cols=("FBI", "VOLUME_IN_IDR"),
    value_display_names=(("FBI", "Sum of FBI"),),
)


class AkumulasiStrategy(WorkflowStrategy):
    """The original enrichment-backed accumulation report."""

    @property
    def definition(self) -> WorkflowDefinition:
        return AKUMULASI_DEFINITION

    def _enrich(
        self, config: ProcessingConfig, data: pd.DataFrame
    ) -> tuple[pd.DataFrame, int]:
        if config.reference_data_path is None:
            raise WorkflowValidationError(
                "The 'Report Summary Akumulasi' workflow requires a reference "
                "mapping file, but none was provided."
            )
        enricher = ReferenceEnricher(
            reference_path=config.reference_data_path,
            lookup_key=config.lookup_key,
        )
        result = enricher.enrich(data)
        return result.data, result.unmapped_count

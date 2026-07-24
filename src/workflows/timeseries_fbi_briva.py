"""Time Series FBI Briva workflow.

Ingests raw pipe-delimited BRIVA data, enriches it against a UKER reference
workbook (VLOOKUP), keeps only NONWHOLESALE rows, then aggregates by
MAIN_CODE + MAIN_BRANCH summing VOLUME_IDR. The summary value column is
displayed as `Sum of VOLUME_IDR` to match the sample Sheet1 oracle.

Runs entirely on the shared ingest -> enrich -> aggregate -> export template;
only the declarative definition and the reference-enrichment hook differ.
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

TIMESERIES_FBI_BRIVA_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.TIMESERIES_FBI_BRIVA,
    label="Time Series FBI Briva",
    requires_reference=True,
    group_cols=("MAIN_CODE", "MAIN_BRANCH"),
    value_col="VOLUME_IDR",
    report_title="Time Series FBI Briva Report",
    detail_sheet_name="Enriched_Data",
    number_format="#,##0.00",
    numeric_columns=("VOLUME_IDR",),
    required_columns=("SEGMEN", "VOLUME_IDR"),
    # Definition-baked inclusion: only NONWHOLESALE rows contribute to the pivot.
    segmen_include="NONWHOLESALE",
    value_display_names=(("VOLUME_IDR", "Sum of VOLUME_IDR"),),
)


class TimeSeriesFbiBrivaStrategy(WorkflowStrategy):
    """NONWHOLESALE BRIVA volume summary with UKER reference enrichment."""

    @property
    def definition(self) -> WorkflowDefinition:
        return TIMESERIES_FBI_BRIVA_DEFINITION

    def _enrich(
        self, config: ProcessingConfig, data: pd.DataFrame
    ) -> tuple[pd.DataFrame, int]:
        if config.reference_data_path is None:
            raise WorkflowValidationError(
                "The 'Time Series FBI Briva' workflow requires a reference "
                "mapping file, but none was provided."
            )
        enricher = ReferenceEnricher(
            reference_path=config.reference_data_path,
            lookup_key=config.lookup_key,
        )
        result = enricher.enrich(data)
        return result.data, result.unmapped_count

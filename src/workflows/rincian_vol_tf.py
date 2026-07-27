"""Rincian Vol TF workflow.

No reference enrichment. Excludes rows whose SEGMEN is "Wholesale"
(case-insensitive; blank/null SEGMEN rows are kept), then groups by
MAINBR + MBDESC summing AMOUNT_IN_IDR into a tabular integer-formatted report.
"""

from __future__ import annotations

from .base import WorkflowDefinition, WorkflowId, WorkflowStrategy

RINCIAN_VOL_TF_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.RINCIAN_VOL_TF,
    label="Rincian Vol TF",
    requires_reference=False,
    group_cols=("MAINBR", "MBDESC"),
    value_col="AMOUNT_IN_IDR",
    report_title="Rincian Vol TF Report",
    detail_sheet_name="Detail_Data",
    numeric_columns=("AMOUNT_IN_IDR",),
    required_columns=("SEGMEN", "MAINBR", "MBDESC", "AMOUNT_IN_IDR"),
    exclude_segmen=("Wholesale",),
    supports_segment_filter=False,
)


class RincianVolTfStrategy(WorkflowStrategy):
    """Branch-code/description volume report with Wholesale excluded."""

    @property
    def definition(self) -> WorkflowDefinition:
        return RINCIAN_VOL_TF_DEFINITION

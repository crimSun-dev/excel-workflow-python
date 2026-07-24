"""Rincian Portal BG workflow.

No reference enrichment and no SEGMEN filtering. Groups by MAINBR + MBNAME
summing AMOUNT_IN_IDR into a tabular integer-formatted report.
"""

from __future__ import annotations

from .base import WorkflowDefinition, WorkflowId, WorkflowStrategy

RINCIAN_PORTAL_BG_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.RINCIAN_PORTAL_BG,
    label="Rincian Portal BG",
    requires_reference=False,
    group_cols=("MAINBR", "MBNAME"),
    value_col="AMOUNT_IN_IDR",
    report_title="Rincian Portal BG Report",
    detail_sheet_name="Detail_Data",
    number_format="#,##0",
    numeric_columns=("AMOUNT_IN_IDR",),
    required_columns=("MAINBR", "MBNAME", "AMOUNT_IN_IDR"),
    supports_segment_filter=False,
)


class RincianPortalBgStrategy(WorkflowStrategy):
    """Branch-code/name report with no segment filtering."""

    @property
    def definition(self) -> WorkflowDefinition:
        return RINCIAN_PORTAL_BG_DEFINITION

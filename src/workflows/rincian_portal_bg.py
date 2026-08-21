"""Rincian Portal BG workflow.

No reference enrichment and no default SEGMEN filtering. Groups by
MAINBR + MBNAME summing AMOUNT_IN_IDR into a tabular integer-formatted report.

FILTERS are skipped for this report: the extract has no confirmed SEGMENT /
SOURCE / KW vocabulary, so the boxes stay grey rather than inviting a guess.
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
    numeric_columns=("AMOUNT_IN_IDR",),
    required_columns=("MAINBR", "MBNAME", "AMOUNT_IN_IDR"),
)


class RincianPortalBgStrategy(WorkflowStrategy):
    """Branch-code/name report with no segment filtering."""

    @property
    def definition(self) -> WorkflowDefinition:
        return RINCIAN_PORTAL_BG_DEFINITION

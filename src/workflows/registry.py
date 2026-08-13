"""Dynamic workflow registry and dispatch.

Maps `WorkflowId` values to singleton strategy instances so the orchestrator can
dispatch without workflow-specific conditionals. New workflows are added here by
registering one entry, keeping the dispatch open for extension.
"""

from __future__ import annotations

from .akumulasi import AkumulasiStrategy
from .base import WorkflowDefinition, WorkflowId, WorkflowStrategy, WorkflowValidationError
from .report_data_statis import ReportDataStatisStrategy
from .report_giro import ReportGiroStrategy
from .rincian_portal_bg import RincianPortalBgStrategy
from .rincian_vol_tf import RincianVolTfStrategy
from .timeseries_active_user_qlola import TimeSeriesActiveUserQlolaStrategy
from .timeseries_fbi_briva import TimeSeriesFbiBrivaStrategy

# Insertion order mirrors `WorkflowId` (the operator's daily sequence) so the
# registry reads the same way the selector does. Dispatch is by key, so the
# order is presentational only.
WORKFLOW_REGISTRY: dict[WorkflowId, WorkflowStrategy] = {
    WorkflowId.AKUMULASI: AkumulasiStrategy(),
    WorkflowId.TIMESERIES_ACTIVE_USER_QLOLA: TimeSeriesActiveUserQlolaStrategy(),
    WorkflowId.REPORT_DATA_STATIS: ReportDataStatisStrategy(),
    WorkflowId.RINCIAN_VOL_TF: RincianVolTfStrategy(),
    WorkflowId.RINCIAN_PORTAL_BG: RincianPortalBgStrategy(),
    WorkflowId.TIMESERIES_FBI_BRIVA: TimeSeriesFbiBrivaStrategy(),
    WorkflowId.REPORT_GIRO: ReportGiroStrategy(),
}


def get_strategy(workflow_id: WorkflowId | str) -> WorkflowStrategy:
    """Returns the strategy for a workflow id (accepts the enum or its value).

    Raises:
        WorkflowValidationError: if the id is unknown/unregistered.
    """
    try:
        key = WorkflowId(workflow_id)
    except ValueError as exc:
        raise WorkflowValidationError(
            f"Unknown workflow id '{workflow_id}'. "
            f"Valid options: {[w.value for w in WorkflowId]}"
        ) from exc
    try:
        return WORKFLOW_REGISTRY[key]
    except KeyError as exc:  # pragma: no cover - registry kept in sync with enum
        raise WorkflowValidationError(
            f"No strategy registered for workflow '{key.value}'."
        ) from exc


def get_definition(workflow_id: WorkflowId | str) -> WorkflowDefinition:
    """Convenience accessor for a workflow's declarative definition."""
    return get_strategy(workflow_id).definition

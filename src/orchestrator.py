"""Pipeline Orchestrator (TDD Section 3.6).

Owns the single exception boundary and telemetry (timing, record counts) for a
pipeline run. Workflow-specific ingest/enrich/aggregate/export logic is
delegated to the selected `WorkflowStrategy` resolved from `WORKFLOW_REGISTRY`,
so adding a workflow never touches this dispatcher.
"""

from __future__ import annotations

import time

from .schemas import PipelineReport, ProcessingConfig
from .workflows.registry import get_strategy


class PipelineOrchestrator:
    """Coordinates the full ETL flow synchronously via the workflow registry."""

    @staticmethod
    def execute(config: ProcessingConfig) -> PipelineReport:
        """Executes the selected workflow with exception boundary handling.

        Never raises for expected pipeline failures; instead returns a
        PipelineReport with success=False and a user-facing error_message.
        """
        start = time.perf_counter()
        try:
            strategy = get_strategy(config.workflow_id)
            result = strategy.execute(config)

            elapsed = time.perf_counter() - start
            return PipelineReport(
                success=True,
                output_path=result.output_path,
                execution_time_seconds=round(elapsed, 4),
                total_records_processed=result.total_records_processed,
                unmapped_records_count=result.unmapped_records_count,
                error_message=None,
            )
        except Exception as exc:  # noqa: BLE001 - boundary: surface, never crash
            elapsed = time.perf_counter() - start
            return PipelineReport(
                success=False,
                output_path=config.output_report_path,
                execution_time_seconds=round(elapsed, 4),
                total_records_processed=0,
                unmapped_records_count=0,
                error_message=f"{type(exc).__name__}: {exc}",
            )

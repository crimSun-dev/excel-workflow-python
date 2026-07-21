"""Pipeline Orchestrator (TDD Section 3.6).

Coordinates execution across Ingestion, Enrichment, Aggregation, and Export
layers with a single exception boundary. Owns execution state and captures
telemetry (timing, record counts, unmapped counts) into a PipelineReport.
"""

from __future__ import annotations

import time

from .aggregation import AggregationEngine
from .enrichment import ReferenceEnricher
from .exporter import ExcelReportExporter
from .ingestion import IngestionEngine
from .schemas import PipelineReport, ProcessingConfig


class PipelineOrchestrator:
    """Coordinates the full ETL flow synchronously."""

    @staticmethod
    def execute(config: ProcessingConfig) -> PipelineReport:
        """Executes the full ETL flow with exception boundary handling.

        Never raises for expected pipeline failures; instead returns a
        PipelineReport with success=False and a user-facing error_message.
        """
        start = time.perf_counter()
        try:
            ingestion = IngestionEngine(delimiter=config.delimiter)
            ingested = ingestion.read_raw_data(config.raw_data_path)

            enricher = ReferenceEnricher(
                reference_path=config.reference_data_path,
                lookup_key=config.lookup_key,
            )
            enriched = enricher.enrich(ingested.data)

            aggregator = AggregationEngine(segment_filter=config.segmen_filter)
            aggregated = aggregator.aggregate(enriched.data)

            exporter = ExcelReportExporter(number_format=config.number_format)
            output_path = exporter.export(
                summary_df=aggregated.summary_data,
                enriched_df=enriched.data,
                output_path=config.output_report_path,
                segment_filter_applied=config.segmen_filter,
            )

            elapsed = time.perf_counter() - start
            return PipelineReport(
                success=True,
                output_path=output_path,
                execution_time_seconds=round(elapsed, 4),
                total_records_processed=ingested.total_rows,
                unmapped_records_count=enriched.unmapped_count,
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

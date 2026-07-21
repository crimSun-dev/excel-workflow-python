"""Domain data contracts & runtime configuration (TDD Section 3.1).

State within the system is strictly immutable and unidirectional:
    Config -> IngestionResult -> EnrichmentResult -> AggregationResult -> PipelineReport
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
from pydantic import BaseModel, Field


class ProcessingConfig(BaseModel):
    """Immutable runtime configuration passed to the orchestrator."""

    raw_data_path: Path = Field(..., description="Path to raw pipe-delimited file")
    reference_data_path: Path = Field(
        ..., description="Path to reference mapping Excel/CSV file"
    )
    output_report_path: Path = Field(
        default=Path("./Financial_Summary_Report.xlsx"),
        description="Destination path for formatted Excel report",
    )
    delimiter: str = Field(default="|", description="Delimiter used in raw file")
    lookup_key: str = Field(default="KODE_UKER", description="Join key column name")
    segmen_filter: Optional[str] = Field(
        default=None, description="Optional SEGMEN value to filter"
    )
    number_format: str = Field(
        default="#,##0.00", description="Excel display format for IDR volume"
    )

    model_config = {"frozen": True, "arbitrary_types_allowed": True}


@dataclass(frozen=True)
class IngestionResult:
    """Output of the ingestion phase."""

    data: pd.DataFrame
    total_rows: int
    malformed_rows_count: int


@dataclass(frozen=True)
class EnrichmentResult:
    """Output of the reference enrichment phase."""

    data: pd.DataFrame
    matched_count: int
    unmapped_count: int
    unmapped_keys: list[str]


@dataclass(frozen=True)
class AggregationResult:
    """Output of the aggregation phase (PivotTable equivalent)."""

    summary_data: pd.DataFrame
    total_volume_idr: float
    branch_count: int


@dataclass(frozen=True)
class PipelineReport:
    """Final telemetry payload returned to the caller."""

    success: bool
    output_path: Path
    execution_time_seconds: float
    total_records_processed: int
    unmapped_records_count: int
    error_message: Optional[str] = None

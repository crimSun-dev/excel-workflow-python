"""Domain data contracts & runtime configuration (TDD Section 3.1).

State within the system is strictly immutable and unidirectional:
    Config -> IngestionResult -> EnrichmentResult -> AggregationResult -> PipelineReport
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional, Union

import pandas as pd
from pydantic import BaseModel, Field, field_validator


class ProcessingConfig(BaseModel):
    """Immutable runtime configuration passed to the orchestrator."""

    raw_data_path: Path = Field(
        ...,
        description="Path to raw pipe-delimited file (.txt/.csv/.bin - a .bin "
        "extract is plain delimited text). For Report Giro this is instead the "
        "monthly giro source workbook (.xlsx/.xls/.xlsm/.csv)",
    )
    reference_data_path: Optional[Path] = Field(
        default=None,
        description="Path to reference mapping Excel/CSV file "
        "(required for Akumulasi, Report Vol Briva, Qlola, and Report Data "
        "Statis workflows)",
    )
    master_data_path: Optional[Path] = Field(
        default=None,
        description="Path to the master workbook: an Excel/CSV file mapping "
        "ID -> MAIN_CODE for Time Series Active User Qlola, or the giro master "
        "workbook to update for Report Giro (required for both)",
    )
    workflow_id: str = Field(
        default="akumulasi",
        description="Selected workflow id (akumulasi | rincian-vol-tf | "
        "rincian-portal-bg | timeseries-fbi-briva | timeseries-active-user-qlola "
        "| report-data-statis | report-giro)",
    )
    output_report_path: Path = Field(
        default=Path("./Financial_Summary_Report.xlsx"),
        description="Destination path for formatted Excel report",
    )
    delimiter: str = Field(default="|", description="Delimiter used in raw file")
    lookup_key: str = Field(default="KODE_UKER", description="Join key column name")
    segmen_filter: Optional[Union[str, list[str]]] = Field(
        default=None,
        description="Runtime override for the workflow's SEGMEN keep-list "
        "(keep only these segments). Accepts one value or a list. None keeps "
        "the workflow definition's default (e.g. 'NONWHOLESALE' for Briva); an "
        "empty string/list keeps every segment. Ignored by workflows that do "
        "not declare supports_segment_filter",
    )
    segmen_exclude: Optional[list[str]] = Field(
        default=None,
        description="Runtime override for the workflow's SEGMEN exclusion list. "
        "None keeps the workflow definition's default (e.g. ['KORPORASI'] for "
        "Report Data Statis); an empty list disables SEGMEN exclusion entirely. "
        "Ignored by workflows that do not declare supports_segment_filter",
    )
    source_exclude: Optional[list[str]] = Field(
        default=None,
        description="Runtime override for the workflow's SOURCE exclusion list. "
        "None keeps the workflow definition's default (e.g. ['CMS'] for Qlola); "
        "an empty list disables SOURCE filtering entirely. Ignored by workflows "
        "that do not declare has_source_filter",
    )
    source_include: Optional[list[str]] = Field(
        default=None,
        description="Runtime SOURCE keep-list (keep only these SOURCE values). "
        "None/empty means no inclusion filter, leaving the workflow's SOURCE "
        "exclusion policy in charge. Ignored by workflows that do not declare "
        "has_source_filter",
    )
    kw_include: Optional[list[str]] = Field(
        default=None,
        description="Runtime override for the workflow's KW keep-list. None "
        "keeps the workflow definition's default (e.g. ['KANWIL MALANG'] for "
        "Report Data Statis); an empty list keeps every KW. The KW column is "
        "resolved by header alias (KW first, then PRODUCT, GROUP_PRODUCT, "
        "KAWIL) and the filter is a no-op when none is present. Ignored by "
        "workflows that do not declare has_kw_filter",
    )
    number_format: str = Field(
        default="#,##0.00", description="Excel display format for IDR volume"
    )
    progress_callback: Optional[Callable[[str], None]] = Field(
        default=None,
        exclude=True,
        description="Optional GUI status hook. Not a CLI flag.",
    )

    model_config = {"frozen": True, "arbitrary_types_allowed": True}

    @field_validator("workflow_id")
    @classmethod
    def _validate_workflow_id(cls, value: str) -> str:
        """Normalizes/validates the workflow id against the registered enum.

        Imported lazily so this schema stays decoupled from the workflows
        package (which depends on the engines, which depend on this module).
        """
        from .workflows.base import WorkflowId

        return WorkflowId(value).value


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
    # Set when a join collapsed almost entirely to UNMAPPED. Advisory only:
    # `success` stays True and the report file is still written.
    unmapped_diagnostic: Optional[str] = None
    # Stage wall times in operator order, e.g. {"read": 0.4, "write": 10.8}.
    stage_timings: Optional[dict[str, float]] = None

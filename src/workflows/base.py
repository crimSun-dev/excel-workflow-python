"""Workflow strategy contracts and the shared execution skeleton.

Defines the `WorkflowId` enum, the declarative `WorkflowDefinition` used to
configure each pipeline, and the `WorkflowStrategy` ABC whose `execute()`
template method runs the shared ingest -> optional enrich -> aggregate -> export
flow. Concrete strategies override only the hooks that differ.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, Literal

import pandas as pd

from ..aggregation import AggregationEngine
from ..enrichment import LOOKUP_KEY_ALIASES, resolve_alias_column
from ..exporter import PLAIN_INTEGER_FORMAT, ExcelReportExporter
from ..ingestion import IngestionEngine

if TYPE_CHECKING:
    from ..schemas import ProcessingConfig


def normalize_join_key(data: pd.DataFrame, lookup_key: str) -> pd.DataFrame:
    """Canonicalizes the raw branch-code column ahead of reference enrichment.

    Real exports spell the branch code as `KODE UKER`, `KODE_UKER` or `UKER`,
    and pandas may hand back integer-typed codes (1 vs "0001") or values with
    stray padding. Each workflow owns its own column quirks, so this runs at the
    workflow ingest step rather than inside the generic enrichment module:
    the first matching alias is renamed to `lookup_key` and its values are cast
    to string and trimmed. Returns the frame untouched when no alias matches -
    the enricher then raises with its full columns-and-aliases diagnostic.
    """
    lookup_key = lookup_key.upper()
    aliases = [lookup_key, *(a for a in LOOKUP_KEY_ALIASES if a != lookup_key)]
    match = resolve_alias_column(data, aliases)
    if match is None:
        return data
    data = data.rename(columns={match: lookup_key}) if match != lookup_key else data.copy()
    data[lookup_key] = data[lookup_key].astype(str).str.strip()
    return data


class WorkflowId(str, Enum):
    """Stable identifiers for the supported report workflows."""

    AKUMULASI = "akumulasi"
    RINCIAN_VOL_TF = "rincian-vol-tf"
    RINCIAN_PORTAL_BG = "rincian-portal-bg"
    TIMESERIES_FBI_BRIVA = "timeseries-fbi-briva"
    TIMESERIES_ACTIVE_USER_QLOLA = "timeseries-active-user-qlola"


class WorkflowValidationError(Exception):
    """Raised when input data is missing columns required by the workflow."""


@dataclass(frozen=True)
class WorkflowDefinition:
    """Declarative configuration for one workflow's ingest/aggregate/export."""

    workflow_id: WorkflowId
    label: str  # human-facing name, e.g. "Report Summary Akumulasi"
    requires_reference: bool
    group_cols: tuple[str, ...]
    value_col: str
    report_title: str
    detail_sheet_name: str
    number_format: str = PLAIN_INTEGER_FORMAT
    numeric_rounding: Literal["round", "ceil"] = "round"
    numeric_columns: tuple[str, ...] = ()
    required_columns: tuple[str, ...] = ()
    exclude_segmen: tuple[str, ...] = ()
    supports_segment_filter: bool = False
    # Definition-baked SEGMEN inclusion filter (e.g. Briva keeps only NONWHOLESALE).
    # Applied ahead of the optional runtime `segmen_filter`; distinct from the
    # `exclude_segmen` exclusion list so inclusion vs exclusion never share a flag.
    segmen_include: str | None = None
    # Dedicated SOURCE exclusion list (e.g. Qlola drops CMS). Kept separate from
    # SEGMEN filtering so the two column filters never collide. Empty (the
    # default) means no SOURCE filtering at all; a runtime `ProcessingConfig.
    # source_exclude` overrides this default when supplied.
    source_exclude: tuple[str, ...] = ()
    # Whether this workflow exposes the SOURCE filter control (GUI enables its
    # SOURCE field only for workflows declaring True; all others grey it out and
    # ignore any value it holds).
    has_source_filter: bool = False
    # Whether this workflow needs a second (master-data) reference file.
    requires_master_data: bool = False
    # Ordered value columns to aggregate. Empty => single-column workflow that
    # sums only `value_col` (backward-compatible with Rincian Vol TF / Portal BG).
    value_cols: tuple[str, ...] = ()
    # Internal column -> report header overrides, e.g. ("FBI", "Sum of FBI").
    value_display_names: tuple[tuple[str, str], ...] = ()

    @property
    def resolved_value_cols(self) -> tuple[str, ...]:
        """Ordered value columns to aggregate; falls back to single `value_col`."""
        return self.value_cols if self.value_cols else (self.value_col,)

    @property
    def display_name_map(self) -> dict[str, str]:
        """Internal column -> report header overrides (e.g. FBI -> Sum of FBI)."""
        return dict(self.value_display_names)


@dataclass(frozen=True)
class WorkflowRunResult:
    """Telemetry returned by a strategy for the orchestrator to report on."""

    output_path: Path
    total_records_processed: int
    unmapped_records_count: int = 0
    unmapped_warning_emitted: bool = False
    # Operator-facing join-failure text, set only when the near-total UNMAPPED
    # threshold trips. The orchestrator forwards it so the GUI can show it
    # instead of leaving the failure buried in the log.
    unmapped_diagnostic: str | None = None


class WorkflowStrategy(ABC):
    """Base class implementing the shared workflow execution skeleton."""

    @property
    @abstractmethod
    def definition(self) -> WorkflowDefinition:
        """The declarative configuration for this workflow."""

    def execute(self, config: "ProcessingConfig") -> WorkflowRunResult:
        """Runs ingest -> optional enrich -> aggregate -> export.

        Raises:
            WorkflowValidationError / DataIngestionError / ExportError /
            ReferenceEnrichmentError: surfaced to the orchestrator's boundary.
        """
        definition = self.definition

        # 1. Ingest raw pipe-delimited data, coercing this workflow's money cols.
        ingestion = IngestionEngine(
            delimiter=config.delimiter,
            numeric_columns=definition.numeric_columns or ("VOLUME_IN_IDR",),
        )
        ingested = ingestion.read_raw_data(config.raw_data_path)
        data = ingested.data

        # 2. Validate that mandatory input columns are present up front.
        self._validate_columns(data)

        # 3. Optional reference enrichment (Akumulasi only).
        data, unmapped_count = self._enrich(config, data)

        # 4. Aggregate (group + sum, with inclusion/exclusion SEGMEN handling).
        # A definition-baked inclusion (e.g. Briva NONWHOLESALE) takes precedence
        # over the optional runtime SEGMEN filter.
        segment_filter = definition.segmen_include or (
            config.segmen_filter if definition.supports_segment_filter else None
        )
        aggregator = AggregationEngine(
            segment_filter=segment_filter,
            exclude_segmen=list(definition.exclude_segmen) or None,
        )
        aggregated = aggregator.aggregate(
            data,
            group_cols=list(definition.group_cols),
            value_col=definition.value_col,
            value_cols=list(definition.resolved_value_cols),
        )

        # 5. Export the parameterized styled workbook.
        exporter = ExcelReportExporter(
            number_format=definition.number_format,
            numeric_rounding=definition.numeric_rounding,
            value_column=definition.value_col,
            value_columns=definition.resolved_value_cols,
            display_names=definition.display_name_map,
            report_title=definition.report_title,
            detail_sheet_name=definition.detail_sheet_name,
        )
        output_path = exporter.export(
            summary_df=aggregated.summary_data,
            enriched_df=data,
            output_path=config.output_report_path,
            segment_filter_applied=self._filter_label(config),
        )

        return WorkflowRunResult(
            output_path=output_path,
            total_records_processed=ingested.total_rows,
            unmapped_records_count=unmapped_count,
        )

    # ------------------------------------------------------------------ #
    # Hooks (overridable)
    # ------------------------------------------------------------------ #
    def _enrich(
        self, config: "ProcessingConfig", data: pd.DataFrame
    ) -> tuple[pd.DataFrame, int]:
        """Default: no enrichment. Returns the data untouched, 0 unmapped."""
        return data, 0

    def resolve_source_exclude(self, config: "ProcessingConfig") -> list[str]:
        """Effective SOURCE exclusion list: runtime override beats the definition.

        Workflows that do not declare `has_source_filter` always return `[]`, so
        a stale value left in the GUI/CLI can never leak into them.
        """
        definition = self.definition
        if not definition.has_source_filter:
            return []
        if config.source_exclude is not None:
            return [s.strip() for s in config.source_exclude if s.strip()]
        return list(definition.source_exclude)

    @staticmethod
    def apply_source_exclude(
        data: pd.DataFrame, source_exclude: list[str]
    ) -> pd.DataFrame:
        """Drops rows whose SOURCE matches an excluded value (case-insensitive).

        An empty list is a no-op, so pre-existing workflows are untouched.
        """
        if not source_exclude or "SOURCE" not in data.columns:
            return data
        excluded = {s.strip().casefold() for s in source_exclude}
        source = data["SOURCE"].astype(str).str.strip().str.casefold()
        return data[~source.isin(excluded)]

    def _filter_label(self, config: "ProcessingConfig") -> str | None:
        """Human-readable filter description for the report metadata block."""
        definition = self.definition
        if definition.exclude_segmen:
            return f"{', '.join(definition.exclude_segmen)} (excluded)"
        if definition.supports_segment_filter and config.segmen_filter:
            return config.segmen_filter
        return None

    def _validate_columns(self, data: pd.DataFrame) -> None:
        """Raises WorkflowValidationError listing any missing input columns."""
        missing = [c for c in self.definition.required_columns if c not in data.columns]
        if missing:
            raise WorkflowValidationError(
                f"Workflow '{self.definition.label}' requires column(s) "
                f"{missing} which are absent from the input file. "
                f"Available columns: {list(data.columns)}"
            )

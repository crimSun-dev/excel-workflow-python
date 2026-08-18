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
from time import perf_counter
from typing import TYPE_CHECKING, Literal

import pandas as pd

from ..aggregation import (
    AggFunc,
    AggregationEngine,
    apply_kw_filters,
    apply_segmen_filters,
    apply_source_filters,
)
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
    """Stable identifiers for the supported report workflows.

    Declaration order is the *display* contract: enums iterate in declaration
    order, so this sequence drives the GUI dropdown and the CLI's valid-options
    listing directly - there is no parallel ordering list to drift out of sync.
    Members are therefore ordered by the operator's daily reporting sequence,
    not by the order the workflows were added. The string values are the public
    API (CLI flags, saved scripts) and never change when the order does.
    """

    AKUMULASI = "akumulasi"
    TIMESERIES_ACTIVE_USER_QLOLA = "timeseries-active-user-qlola"
    REPORT_DATA_STATIS = "report-data-statis"
    RINCIAN_VOL_TF = "rincian-vol-tf"
    RINCIAN_PORTAL_BG = "rincian-portal-bg"
    TIMESERIES_FBI_BRIVA = "timeseries-fbi-briva"
    REPORT_GIRO = "report-giro"


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
    # Default SEGMEN values to drop (e.g. Report Data Statis drops KORPORASI).
    # A runtime `ProcessingConfig.segmen_exclude` overrides this default.
    exclude_segmen: tuple[str, ...] = ()
    # Whether this workflow exposes the SEGMENT filter control. True for every
    # workflow that can honor it - including Report Giro, which applies it to
    # its monthly source. False greys the field out, so a disabled field always
    # means "this value would be ignored".
    supports_segment_filter: bool = False
    # Default SEGMEN inclusion filter (e.g. Briva keeps only NONWHOLESALE).
    # Applies whenever the operator leaves the SEGMENT box empty; distinct from
    # the `exclude_segmen` exclusion list so inclusion vs exclusion never share
    # a flag.
    segmen_include: str | None = None
    # How the value columns are aggregated: "sum" totals them (all money
    # workflows), "count" reproduces Excel's Count of non-blank cells (Report
    # Data Statis counts ID_PRODUCT, which is textual and must never be summed).
    aggfunc: AggFunc = "sum"
    # Dedicated SOURCE exclusion list (e.g. Qlola drops CMS). Kept separate from
    # SEGMEN filtering so the two column filters never collide. Empty (the
    # default) means no SOURCE filtering at all; a runtime `ProcessingConfig.
    # source_exclude` overrides this default when supplied.
    source_exclude: tuple[str, ...] = ()
    # Whether this workflow exposes the SOURCE filter control (GUI enables its
    # SOURCE field only for workflows declaring True; all others grey it out and
    # ignore any value it holds).
    has_source_filter: bool = False
    # Definition-baked KW keep-list (empty = keep every KW), e.g. Report Data
    # Statis keeps only "KANWIL MALANG". Regional filtering is part of this one
    # KW dimension rather than a second KAWIL flag, because `KW` is the column
    # real extracts carry. Resolved by header alias, so the same control serves
    # older extracts spelling it PRODUCT, GROUP_PRODUCT or KAWIL, and is a no-op
    # where none exist.
    kw_include: tuple[str, ...] = ()
    # Whether this workflow exposes the KW filter control, mirroring
    # `has_source_filter`: False greys the field out and discards its value.
    has_kw_filter: bool = False
    # Whether this workflow needs a second (master-data) reference file.
    requires_master_data: bool = False
    # Ordered value columns to aggregate. Empty => single-column workflow that
    # sums only `value_col` (backward-compatible with Rincian Vol TF / Portal BG).
    value_cols: tuple[str, ...] = ()
    # Internal column -> report header overrides, e.g. ("FBI", "Sum of FBI").
    value_display_names: tuple[tuple[str, str], ...] = ()
    # ---- File-picker metadata (GUI) ----------------------------------------
    # Most workflows read a pipe-delimited text extract, but Report Giro's raw
    # input is a monthly Excel workbook, so the button text, dialog title, and
    # extension filter are declarative rather than hardcoded in the GUI.
    raw_button_label: str = "1. Raw Data File..."
    raw_picker_title: str = "Select raw pipe-delimited data file"
    raw_filetypes: tuple[tuple[str, str], ...] = (
        # `.bin` raw extracts (Report Data Statis) are plain delimited text
        # despite the extension, so they belong in the same filter.
        ("Text/CSV/BIN", "*.txt *.csv *.bin"),
        ("All files", "*.*"),
    )
    master_button_label: str = "2b. Master Data File..."
    master_picker_title: str = "Select master-data workbook (ID -> MAIN_CODE)"

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
    # Named stage durations for the success dialog (read / match / write).
    stage_timings: dict[str, float] | None = None


class WorkflowStrategy(ABC):
    """Base class implementing the shared workflow execution skeleton."""

    @property
    @abstractmethod
    def definition(self) -> WorkflowDefinition:
        """The declarative configuration for this workflow."""

    def execute(self, config: "ProcessingConfig") -> WorkflowRunResult:
        """Runs ingest -> filter -> optional enrich -> aggregate -> export.

        SEGMENT / SOURCE / KW are applied *before* enrichment and the detail
        sheet, so a default Statis run never VLOOKUPs or writes rows the
        summary is going to drop. Aggregate still reapplies the same filters
        (cheap on an already-cut frame) so direct engine callers stay correct.

        Raises:
            WorkflowValidationError / DataIngestionError / ExportError /
            ReferenceEnrichmentError: surfaced to the orchestrator's boundary.
        """
        definition = self.definition
        stages: dict[str, float] = {}

        def mark(name: str, started: float) -> float:
            stages[name] = round(perf_counter() - started, 4)
            return perf_counter()

        clock = perf_counter()
        self._emit_progress(config, "Reading file...")
        ingestion = IngestionEngine(
            delimiter=config.delimiter,
            numeric_columns=definition.numeric_columns or ("VOLUME_IN_IDR",),
        )
        ingested = ingestion.read_raw_data(config.raw_data_path)
        data = ingested.data
        clock = mark("read", clock)

        self._validate_columns(data)

        self._emit_progress(config, "Applying filters...")
        # SOURCE was already pre-enrich; SEGMEN/KW join it so Statis does not
        # enrich or export the national extract just to keep KANWIL MALANG.
        data = self.apply_runtime_filters(data, config)
        clock = mark("filter", clock)

        self._emit_progress(config, "Matching reference...")
        data, unmapped_count = self._enrich(config, data)
        clock = mark("match", clock)

        aggregator = AggregationEngine(
            segment_filter=self.resolve_segment_filter(config),
            exclude_segmen=self.resolve_segmen_exclude(config),
            kw_include=self.resolve_kw_include(config),
        )
        aggregated = aggregator.aggregate(
            data,
            group_cols=list(definition.group_cols),
            value_col=definition.value_col,
            value_cols=list(definition.resolved_value_cols),
            aggfunc=definition.aggfunc,
        )

        self._emit_progress(config, "Writing Excel...")
        exporter = ExcelReportExporter(
            number_format=definition.number_format,
            numeric_rounding=definition.numeric_rounding,
            value_column=definition.value_col,
            value_columns=definition.resolved_value_cols,
            display_names=definition.display_name_map,
            report_title=definition.report_title,
            detail_sheet_name=definition.detail_sheet_name,
            format_detail_values=definition.aggfunc != "count",
        )
        output_path = exporter.export(
            summary_df=aggregated.summary_data,
            enriched_df=data,
            output_path=config.output_report_path,
            segment_filter_applied=self._filter_label(config),
        )
        mark("write", clock)

        return WorkflowRunResult(
            output_path=output_path,
            total_records_processed=ingested.total_rows,
            unmapped_records_count=unmapped_count,
            stage_timings=stages,
        )

    # ------------------------------------------------------------------ #
    # Hooks (overridable)
    # ------------------------------------------------------------------ #
    def _enrich(
        self, config: "ProcessingConfig", data: pd.DataFrame
    ) -> tuple[pd.DataFrame, int]:
        """Default: no enrichment. Returns the data untouched, 0 unmapped."""
        return data, 0

    def resolve_segment_filter(self, config: "ProcessingConfig") -> str | list[str] | None:
        """Effective SEGMENT keep-list: runtime choice beats the default.

        Returns `None` when no inclusion filter applies. A supplied but empty
        value is the caller asking for every segment - a deliberate instruction
        rather than a request to fall back to the workflow default (the GUI
        sends `None`, not an empty string, when the operator wants the default).
        Workflows that do not declare `supports_segment_filter` always keep
        their baked default, so a stale value can never leak into them.
        """
        definition = self.definition
        if definition.supports_segment_filter and config.segmen_filter is not None:
            if isinstance(config.segmen_filter, str):
                return config.segmen_filter.strip() or None
            return [s.strip() for s in config.segmen_filter if s.strip()] or None
        return definition.segmen_include

    def resolve_segmen_exclude(self, config: "ProcessingConfig") -> list[str]:
        """Effective SEGMEN exclusion list: runtime choice beats the default.

        An empty list explicitly means "exclude nothing" - that is how the
        operator re-includes segments the workflow drops by default (e.g. keeping
        KORPORASI in a Report Data Statis run).
        """
        definition = self.definition
        if definition.supports_segment_filter and config.segmen_exclude is not None:
            return [s.strip() for s in config.segmen_exclude if s.strip()]
        return list(definition.exclude_segmen)

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

    def resolve_source_include(self, config: "ProcessingConfig") -> list[str]:
        """Effective SOURCE keep-list. Runtime-only: no workflow bakes one in.

        Workflows that do not declare `has_source_filter` always return `[]`,
        matching `resolve_source_exclude`.
        """
        if not self.definition.has_source_filter or config.source_include is None:
            return []
        return [s.strip() for s in config.source_include if s.strip()]

    def resolve_kw_include(self, config: "ProcessingConfig") -> list[str]:
        """Effective KW keep-list: runtime override beats the definition.

        An empty list explicitly means "keep every KW"; `None` (nothing supplied)
        falls back to the definition. Workflows that do not declare
        `has_kw_filter` always return `[]`.
        """
        definition = self.definition
        if not definition.has_kw_filter:
            return []
        if config.kw_include is not None:
            return [s.strip() for s in config.kw_include if s.strip()]
        return list(definition.kw_include)

    def apply_runtime_filters(
        self, data: pd.DataFrame, config: "ProcessingConfig"
    ) -> pd.DataFrame:
        """Applies the resolved SEGMENT / SOURCE / KW filters to a frame.

        The one entry point the custom-`execute` workflows (Qlola, Giro) share
        with the template below, so every report answers the operator's filter
        boxes identically. Dimensions whose column is absent are a no-op.
        """
        data = apply_source_filters(
            data, self.resolve_source_include(config), self.resolve_source_exclude(config)
        )
        data = apply_segmen_filters(
            data, self.resolve_segment_filter(config), self.resolve_segmen_exclude(config)
        )
        return apply_kw_filters(data, self.resolve_kw_include(config))

    @staticmethod
    def _emit_progress(config: "ProcessingConfig", message: str) -> None:
        """Forwards a status line to the GUI when a callback was supplied."""
        callback = config.progress_callback
        if callback is not None:
            callback(message)

    @staticmethod
    def apply_source_exclude(
        data: pd.DataFrame, source_exclude: list[str]
    ) -> pd.DataFrame:
        """Drops rows whose SOURCE matches an excluded value (case-insensitive).

        An empty list is a no-op, so pre-existing workflows are untouched.
        """
        return apply_source_filters(data, exclude=source_exclude)

    def _filter_label(self, config: "ProcessingConfig") -> str | None:
        """Human-readable description of the filters this run actually applied.

        Built from the resolved values rather than the definition, so a report
        produced with operator-edited filters states what was really excluded.
        """
        parts: list[str] = []
        include = self.resolve_segment_filter(config)
        if include:
            values = [include] if isinstance(include, str) else list(include)
            parts.append(f"{', '.join(values)} (only)")
        segmen_exclude = self.resolve_segmen_exclude(config)
        if segmen_exclude:
            parts.append(f"{', '.join(segmen_exclude)} (excluded)")
        # SOURCE and KW are reported alongside the SEGMEN rules so the metadata
        # block never understates what was filtered out.
        source_include = self.resolve_source_include(config)
        if source_include:
            parts.append(f"{', '.join(source_include)} (SOURCE only)")
        source_exclude = self.resolve_source_exclude(config)
        if source_exclude:
            parts.append(f"{', '.join(source_exclude)} (SOURCE excluded)")
        kw_include = self.resolve_kw_include(config)
        if kw_include:
            parts.append(f"{', '.join(kw_include)} (KW only)")
        return "; ".join(parts) or None

    def _validate_columns(self, data: pd.DataFrame) -> None:
        """Raises WorkflowValidationError listing any missing input columns."""
        missing = [c for c in self.definition.required_columns if c not in data.columns]
        if missing:
            raise WorkflowValidationError(
                f"Workflow '{self.definition.label}' requires column(s) "
                f"{missing} which are absent from the input file. "
                f"Available columns: {list(data.columns)}"
            )

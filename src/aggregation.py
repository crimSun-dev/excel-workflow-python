"""Aggregation Engine (TDD Section 3.4).

Replicates Excel's PivotTable: optionally filters by SEGMEN, SOURCE and KW,
groups by [MAIN_CODE, MAIN_BRANCH], and sums (or counts) the value column.
Subtotals are intentionally suppressed (matching the tabular-form, no-subtotal
layout in the source workflow); only a single Grand Total is computed separately
by the exporter.

Every operator-facing dimension (SEGMENT, SOURCE, KW) runs through the one
`apply_column_filters` helper, so "keep only these" and "drop these" behave
identically everywhere and a dimension whose column is absent is always a
silent no-op rather than an error.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Literal, Optional, Union

import pandas as pd

from .enrichment import resolve_alias_column
from .schemas import AggregationResult

SEGMEN_COLUMN = "SEGMEN"
SOURCE_COLUMN = "SOURCE"

# KW is the stakeholder's label for the product/regional dimension, and `KW` is
# the header real extracts carry - it is the one operator-facing name for it, so
# regional values such as "KANWIL MALANG" are matched here rather than through a
# second KAWIL dimension. The remaining spellings stay as ranked fallbacks for
# older extracts; the control is a no-op when the source carries none of them.
KW_COLUMN_ALIASES = ["KW", "PRODUCT", "GROUP_PRODUCT", "KAWIL"]

# One value, a list of values, or nothing at all - the GUI hands over a
# keep-list while the CLI and older callers may pass a single string.
FilterValues = Union[str, Sequence[str], None]

AggFunc = Literal["sum", "count"]


def _non_blank_mask(series: pd.Series) -> pd.Series:
    """Excel's Count semantics: True for cells that are neither null nor blank.

    Whitespace-only cells are treated as blank, matching a PivotTable Count
    over a text column that was produced by a delimited export.
    """
    return series.notna() & (series.astype(str).str.strip() != "")


def normalize_filter_values(values: FilterValues) -> set[str]:
    """Trimmed, case-folded comparison set; blanks drop out.

    `None`, `""`, `[]` and `["", "  "]` all collapse to an empty set, which every
    caller reads as "no filtering on this dimension".
    """
    if values is None:
        return set()
    if isinstance(values, str):
        values = [values]
    return {str(v).strip().casefold() for v in values if str(v).strip()}


def apply_column_filters(
    df: pd.DataFrame,
    aliases: list[str],
    include: FilterValues = None,
    exclude: FilterValues = None,
) -> pd.DataFrame:
    """Keeps only `include` and drops `exclude` on the first alias column present.

    Both filters are case-insensitive, trimmed, and AND-combined. An empty
    `include` means "keep everything"; blank/null cells never match an exclusion
    so they are retained, while an inclusion filter drops them (a blank cell is
    not one of the values the operator asked to keep). A source that has none of
    the alias columns makes both filters a no-op.
    """
    included = normalize_filter_values(include)
    excluded = normalize_filter_values(exclude)
    if not included and not excluded:
        return df

    column = resolve_alias_column(df, aliases)
    if column is None:
        return df

    if included:
        values = df[column].astype(str).str.strip().str.casefold()
        df = df[values.isin(included)]
    if excluded:
        values = df[column].astype(str).str.strip().str.casefold()
        df = df[~values.isin(excluded)]
    return df


def apply_segmen_filters(
    df: pd.DataFrame,
    include: FilterValues = None,
    exclude: FilterValues = None,
) -> pd.DataFrame:
    """SEGMEN keep-list / drop-list, shared by the engine and Qlola / Giro."""
    return apply_column_filters(df, [SEGMEN_COLUMN], include, exclude)


def apply_source_filters(
    df: pd.DataFrame,
    include: FilterValues = None,
    exclude: FilterValues = None,
) -> pd.DataFrame:
    """SOURCE keep-list / drop-list (e.g. Qlola's default CMS exclusion)."""
    return apply_column_filters(df, [SOURCE_COLUMN], include, exclude)


def apply_kw_filters(df: pd.DataFrame, include: FilterValues = None) -> pd.DataFrame:
    """KW keep-list, resolved against `KW_COLUMN_ALIASES` (no-op when absent)."""
    return apply_column_filters(df, KW_COLUMN_ALIASES, include)


class AggregationEngine:
    """Summarizes enriched records by branch codes and totals the volume."""

    def __init__(
        self,
        segment_filter: FilterValues = None,
        exclude_segmen: Optional[list[str]] = None,
        kw_include: FilterValues = None,
    ):
        self.segment_filter = segment_filter
        # SEGMEN values to drop before grouping (e.g. Rincian Vol TF excludes
        # "Wholesale"). Operates independently of the inclusion filter above.
        self.exclude_segmen = exclude_segmen or []
        # Operator-facing KW keep-list, resolved by header alias and AND-combined
        # with the SEGMEN filters above. Its own knob so the two column
        # dimensions can never be conflated. Regional keep-lists live here too
        # (Report Data Statis keeps only "KANWIL MALANG"): KW is the single
        # dimension for that column, not a separate KAWIL filter.
        self.kw_include = kw_include

    def aggregate(
        self,
        enriched_df: pd.DataFrame,
        group_cols: Optional[list[str]] = None,
        value_col: str = "VOLUME_IN_IDR",
        value_cols: Optional[list[str]] = None,
        aggfunc: AggFunc = "sum",
    ) -> AggregationResult:
        """Filters, groups, and sums/counts to replicate Excel Pivot Table behavior.

        Args:
            enriched_df: DataFrame output from ReferenceEnricher.
            group_cols: Columns to group by (default MAIN_CODE, MAIN_BRANCH).
            value_col: Primary value column; drives `total_volume_idr` telemetry.
            value_cols: Ordered value columns to sum (e.g. FBI then VOLUME_IN_IDR).
                Defaults to a single-column list of `value_col`, preserving the
                original single-value behavior for the Rincian workflows.
            aggfunc: `"sum"` (default) totals the numeric value columns with IDR
                rounding; `"count"` reproduces Excel's Count of a (typically
                textual) column - the number of non-blank cells per group, with
                duplicates counted separately and no numeric coercion.

        Returns:
            AggregationResult containing the summary table and aggregate metrics.
            Summary columns are ordered as group_cols followed by value_cols.
        """
        if group_cols is None:
            group_cols = ["MAIN_CODE", "MAIN_BRANCH"]

        if not value_cols:
            value_cols = [value_col]

        # The telemetry total tracks the primary volume column when present.
        total_col = value_col if value_col in value_cols else value_cols[-1]

        df = enriched_df.copy()

        # Apply the SEGMEN include/exclude filters (Pivot "Filters" field) before
        # grouping. Both are operator-editable at runtime, so the same helper
        # backs the GUI fields and every definition-baked default.
        df = apply_segmen_filters(df, self.segment_filter, self.exclude_segmen)

        # Apply the KW keep-list (case-insensitive, trimmed). Resolved by alias
        # and skipped entirely when the source carries no KW-like column, so a
        # filter typed against an extract without that dimension can never fail
        # a run. Blank/null KW never matches an inclusion, so those rows drop -
        # the opposite of the SEGMEN exclusion above, which retains them.
        df = apply_kw_filters(df, self.kw_include)

        missing = [c for c in value_cols if c not in df.columns]
        if missing:
            raise KeyError(f"Value column(s) {missing} not present in dataset")

        if aggfunc == "count":
            # Excel Count: 1 per non-blank cell. No float conversion and no
            # fillna(0), so missing IDs are never fabricated as countable zeros.
            for col in value_cols:
                df[col] = _non_blank_mask(df[col]).astype("int64")
        else:
            for col in value_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        if df.empty:
            summary = pd.DataFrame(columns=[*group_cols, *value_cols])
            return AggregationResult(
                summary_data=summary, total_volume_idr=0.0, branch_count=0
            )

        summary = (
            df.groupby(group_cols, as_index=False, dropna=False)[value_cols]
            .sum()
            .sort_values(by=group_cols, ascending=True)
            .reset_index(drop=True)
        )

        if aggfunc == "count":
            # Counts are exact integers; rounding them would only risk a float
            # round-trip, so the IDR precision guard below is skipped.
            total_volume = float(summary[total_col].sum())
        else:
            # Explicit 2-decimal rounding guards against float sum drift (IDR precision).
            for col in value_cols:
                summary[col] = summary[col].round(2)
            total_volume = float(round(summary[total_col].sum(), 2))
        branch_count = int(len(summary))

        return AggregationResult(
            summary_data=summary,
            total_volume_idr=total_volume,
            branch_count=branch_count,
        )

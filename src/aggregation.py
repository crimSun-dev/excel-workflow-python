"""Aggregation Engine (TDD Section 3.4).

Replicates Excel's PivotTable: optionally filters by SEGMEN and KAWIL, groups by
[MAIN_CODE, MAIN_BRANCH], and sums (or counts) the value column. Subtotals are
intentionally suppressed (matching the tabular-form, no-subtotal layout in the
source workflow); only a single Grand Total is computed separately by the
exporter.
"""

from __future__ import annotations

from typing import Literal, Optional

import pandas as pd

from .schemas import AggregationResult

SEGMEN_COLUMN = "SEGMEN"
KAWIL_COLUMN = "KAWIL"

AggFunc = Literal["sum", "count"]


def _non_blank_mask(series: pd.Series) -> pd.Series:
    """Excel's Count semantics: True for cells that are neither null nor blank.

    Whitespace-only cells are treated as blank, matching a PivotTable Count
    over a text column that was produced by a delimited export.
    """
    return series.notna() & (series.astype(str).str.strip() != "")


def apply_segmen_filters(
    df: pd.DataFrame,
    include: Optional[str] = None,
    exclude: Optional[list[str]] = None,
) -> pd.DataFrame:
    """Keeps only `include` and drops `exclude`, mirroring the Pivot Filters pane.

    Both filters are case-insensitive, trimmed, and AND-combined. Shared by the
    aggregation engine and by workflows that build their own summary (Qlola), so
    "keep only X" and "drop Y" behave identically everywhere.

    An empty or whitespace-only `include` means "all segments" - that is how an
    operator clearing the field asks for no inclusion filter. Blank/null SEGMEN
    cells never match an exclusion, so they are retained. A missing SEGMEN
    column makes both filters a no-op.
    """
    if SEGMEN_COLUMN not in df.columns:
        return df

    if include is not None and include.strip():
        segmen = df[SEGMEN_COLUMN].astype(str).str.strip().str.casefold()
        df = df[segmen == include.strip().casefold()]

    excluded = {s.strip().casefold() for s in (exclude or []) if s.strip()}
    if excluded:
        segmen = df[SEGMEN_COLUMN].astype(str).str.strip().str.casefold()
        df = df[~segmen.isin(excluded)]

    return df


class AggregationEngine:
    """Summarizes enriched records by branch codes and totals the volume."""

    def __init__(
        self,
        segment_filter: Optional[str] = None,
        exclude_segmen: Optional[list[str]] = None,
        kawil_include: Optional[str] = None,
    ):
        self.segment_filter = segment_filter
        # SEGMEN values to drop before grouping (e.g. Rincian Vol TF excludes
        # "Wholesale"). Operates independently of the inclusion filter above.
        self.exclude_segmen = exclude_segmen or []
        # Regional-office (KAWIL) inclusion filter, e.g. Report Data Statis keeps
        # only "KANWIL MALANG". Deliberately a separate knob from the SEGMEN
        # filters so the two column dimensions can never be conflated; it is
        # AND-combined with them.
        self.kawil_include = kawil_include

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

        # Apply the optional KAWIL inclusion filter (case-insensitive, trimmed).
        # Blank/null KAWIL never matches, so those rows are dropped - the
        # opposite of the SEGMEN exclusion above, which retains them.
        if self.kawil_include is not None and KAWIL_COLUMN in df.columns:
            kawil = df[KAWIL_COLUMN].astype(str).str.strip().str.casefold()
            df = df[kawil == self.kawil_include.strip().casefold()]

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

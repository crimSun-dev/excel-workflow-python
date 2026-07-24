"""Aggregation Engine (TDD Section 3.4).

Replicates Excel's PivotTable: optionally filters by SEGMEN, groups by
[MAIN_CODE, MAIN_BRANCH], and sums VOLUME_IN_IDR. Subtotals are intentionally
suppressed (matching the tabular-form, no-subtotal layout in the source
workflow); only a single Grand Total is computed separately by the exporter.
"""

from __future__ import annotations

from typing import Optional

import pandas as pd

from .schemas import AggregationResult

SEGMEN_COLUMN = "SEGMEN"


class AggregationEngine:
    """Summarizes enriched records by branch codes and totals the volume."""

    def __init__(
        self,
        segment_filter: Optional[str] = None,
        exclude_segmen: Optional[list[str]] = None,
    ):
        self.segment_filter = segment_filter
        # SEGMEN values to drop before grouping (e.g. Rincian Vol TF excludes
        # "Wholesale"). Operates independently of the inclusion filter above.
        self.exclude_segmen = exclude_segmen or []

    def aggregate(
        self,
        enriched_df: pd.DataFrame,
        group_cols: Optional[list[str]] = None,
        value_col: str = "VOLUME_IN_IDR",
        value_cols: Optional[list[str]] = None,
    ) -> AggregationResult:
        """Filters, groups, and sums to replicate Excel Pivot Table behavior.

        Args:
            enriched_df: DataFrame output from ReferenceEnricher.
            group_cols: Columns to group by (default MAIN_CODE, MAIN_BRANCH).
            value_col: Primary value column; drives `total_volume_idr` telemetry.
            value_cols: Ordered value columns to sum (e.g. FBI then VOLUME_IN_IDR).
                Defaults to a single-column list of `value_col`, preserving the
                original single-value behavior for the Rincian workflows.

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

        # Apply the optional SEGMEN filter (Pivot "Filters" field) before grouping.
        if self.segment_filter is not None and SEGMEN_COLUMN in df.columns:
            df = df[
                df[SEGMEN_COLUMN].astype(str).str.strip().str.casefold()
                == self.segment_filter.strip().casefold()
            ]

        # Apply the optional SEGMEN exclusion list (case-insensitive). Blank or
        # null SEGMEN values never match an exclusion, so they are retained.
        if self.exclude_segmen and SEGMEN_COLUMN in df.columns:
            excluded = {s.strip().casefold() for s in self.exclude_segmen}
            segmen = df[SEGMEN_COLUMN].astype(str).str.strip().str.casefold()
            df = df[~segmen.isin(excluded)]

        missing = [c for c in value_cols if c not in df.columns]
        if missing:
            raise KeyError(f"Value column(s) {missing} not present in dataset")

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

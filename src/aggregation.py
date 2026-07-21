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

    def __init__(self, segment_filter: Optional[str] = None):
        self.segment_filter = segment_filter

    def aggregate(
        self,
        enriched_df: pd.DataFrame,
        group_cols: Optional[list[str]] = None,
        value_col: str = "VOLUME_IN_IDR",
    ) -> AggregationResult:
        """Filters, groups, and sums to replicate Excel Pivot Table behavior.

        Args:
            enriched_df: DataFrame output from ReferenceEnricher.
            group_cols: Columns to group by (default MAIN_CODE, MAIN_BRANCH).
            value_col: Column to sum.

        Returns:
            AggregationResult containing the summary table and aggregate metrics.
        """
        if group_cols is None:
            group_cols = ["MAIN_CODE", "MAIN_BRANCH"]

        df = enriched_df.copy()

        # Apply the optional SEGMEN filter (Pivot "Filters" field) before grouping.
        if self.segment_filter is not None and SEGMEN_COLUMN in df.columns:
            df = df[
                df[SEGMEN_COLUMN].astype(str).str.strip().str.casefold()
                == self.segment_filter.strip().casefold()
            ]

        if value_col not in df.columns:
            raise KeyError(f"Value column '{value_col}' not present in dataset")

        df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0.0)

        if df.empty:
            summary = pd.DataFrame(columns=[*group_cols, value_col])
            return AggregationResult(
                summary_data=summary, total_volume_idr=0.0, branch_count=0
            )

        summary = (
            df.groupby(group_cols, as_index=False, dropna=False)[value_col]
            .sum()
            .sort_values(by=group_cols, ascending=True)
            .reset_index(drop=True)
        )

        # Explicit 2-decimal rounding guards against float sum drift (IDR precision).
        summary[value_col] = summary[value_col].round(2)

        total_volume = float(round(summary[value_col].sum(), 2))
        branch_count = int(len(summary))

        return AggregationResult(
            summary_data=summary,
            total_volume_idr=total_volume,
            branch_count=branch_count,
        )

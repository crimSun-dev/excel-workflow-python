"""Reference Enricher (TDD Section 3.3).

Replicates Excel's VLOOKUP: left-joins the raw dataset against a reference
workbook (Excel/CSV) on KODE_UKER to attach MAIN_CODE and MAIN_BRANCH. Codes
absent from the reference are flagged as "UNMAPPED" and surfaced in an audit
payload so report generation is never blocked by minor mapping gaps.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .schemas import EnrichmentResult

MAIN_CODE = "MAIN_CODE"
MAIN_BRANCH = "MAIN_BRANCH"
UNMAPPED_SENTINEL = "UNMAPPED"


class ReferenceEnrichmentError(Exception):
    """Raised when reference workbook loading fails."""


class ReferenceEnricher:
    """Enriches the main dataset with reference mapping data (VLOOKUP equivalent)."""

    def __init__(self, reference_path: Path, lookup_key: str = "KODE_UKER"):
        self.reference_path = Path(reference_path)
        self.lookup_key = lookup_key.upper()

    def enrich(self, raw_df: pd.DataFrame) -> EnrichmentResult:
        """Performs a left outer join to attach MAIN_CODE and MAIN_BRANCH.

        Args:
            raw_df: Ingested raw DataFrame.

        Returns:
            EnrichmentResult with the merged dataset and orphan mapping stats.

        Raises:
            ReferenceEnrichmentError: If the reference file is missing or the
                lookup key column is absent in either dataset.
        """
        ref_df = self._load_reference()

        if self.lookup_key not in raw_df.columns:
            raise ReferenceEnrichmentError(
                f"Lookup key '{self.lookup_key}' absent from raw data. "
                f"Available columns: {list(raw_df.columns)}"
            )
        if self.lookup_key not in ref_df.columns:
            raise ReferenceEnrichmentError(
                f"Lookup key '{self.lookup_key}' absent from reference file. "
                f"Available columns: {list(ref_df.columns)}"
            )
        for required in (MAIN_CODE, MAIN_BRANCH):
            if required not in ref_df.columns:
                raise ReferenceEnrichmentError(
                    f"Reference file missing required column '{required}'. "
                    f"Available columns: {list(ref_df.columns)}"
                )

        # Standardize the join key to string on both sides to avoid the classic
        # int-vs-str VLOOKUP mismatch (e.g. 001 vs 1).
        merged = raw_df.copy()
        merged[self.lookup_key] = merged[self.lookup_key].astype(str).str.strip()

        ref_lookup = ref_df[[self.lookup_key, MAIN_CODE, MAIN_BRANCH]].copy()
        ref_lookup[self.lookup_key] = (
            ref_lookup[self.lookup_key].astype(str).str.strip()
        )
        # Reference tables can contain duplicate keys; VLOOKUP takes the first.
        ref_lookup = ref_lookup.drop_duplicates(subset=self.lookup_key, keep="first")

        merged = merged.merge(ref_lookup, on=self.lookup_key, how="left")

        unmapped_mask = merged[MAIN_CODE].isna()
        unmapped_keys = sorted(
            merged.loc[unmapped_mask, self.lookup_key].unique().tolist()
        )
        merged[MAIN_CODE] = merged[MAIN_CODE].fillna(UNMAPPED_SENTINEL)
        merged[MAIN_BRANCH] = merged[MAIN_BRANCH].fillna(UNMAPPED_SENTINEL)

        unmapped_count = int(unmapped_mask.sum())
        matched_count = int(len(merged) - unmapped_count)

        return EnrichmentResult(
            data=merged,
            matched_count=matched_count,
            unmapped_count=unmapped_count,
            unmapped_keys=unmapped_keys,
        )

    def _load_reference(self) -> pd.DataFrame:
        """Loads the reference workbook (.xlsx/.xls/.csv) into a DataFrame."""
        if not self.reference_path.exists():
            raise ReferenceEnrichmentError(
                f"Reference file not found: {self.reference_path}"
            )
        suffix = self.reference_path.suffix.lower()
        try:
            if suffix in (".xlsx", ".xls", ".xlsm"):
                ref_df = pd.read_excel(self.reference_path, dtype=str)
            elif suffix == ".csv":
                ref_df = pd.read_csv(self.reference_path, dtype=str)
            else:
                raise ReferenceEnrichmentError(
                    f"Unsupported reference file type: '{suffix}'"
                )
        except ReferenceEnrichmentError:
            raise
        except Exception as exc:  # noqa: BLE001 - re-wrapped for the caller
            raise ReferenceEnrichmentError(
                f"Failed to load reference file {self.reference_path}: {exc}"
            ) from exc

        ref_df.columns = [str(c).strip().upper() for c in ref_df.columns]
        return ref_df

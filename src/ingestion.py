"""Ingestion Engine (TDD Section 3.2).

Replicates Excel's "Text to Columns" feature: reads a pipe-delimited raw text
file into a structured DataFrame with automatic encoding fallback, whitespace
trimming, header sanitization, and numeric coercion of the volume column.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .schemas import IngestionResult

# Column expected to hold the monetary volume; coerced to float64.
VOLUME_COLUMN = "VOLUME_IN_IDR"

# Encodings attempted in order. utf-8-sig transparently strips a BOM if present.
_ENCODING_FALLBACKS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")


class DataIngestionError(Exception):
    """Raised when raw file parsing encounters unrecoverable structural errors."""


class IngestionEngine:
    """Parses pipe-delimited raw financial data files into structured DataFrames."""

    def __init__(
        self,
        delimiter: str = "|",
        numeric_columns: tuple[str, ...] | list[str] = (VOLUME_COLUMN,),
    ):
        self.delimiter = delimiter
        # Workflow-specific monetary columns to coerce to float64 (e.g. the
        # Akumulasi workflow uses VOLUME_IN_IDR, the Rincian workflows use
        # AMOUNT_IN_IDR). Only columns actually present in the file are coerced.
        self.numeric_columns = tuple(numeric_columns)

    def read_raw_data(self, file_path: Path) -> IngestionResult:
        """Reads pipe-delimited text file with automatic encoding fallback.

        Args:
            file_path: Absolute path to the raw data file.

        Returns:
            IngestionResult containing a clean DataFrame and parsing metrics.

        Raises:
            DataIngestionError: If the file is missing, empty, or has no columns.
        """
        file_path = Path(file_path)
        if not file_path.exists():
            raise DataIngestionError(f"Raw data file not found: {file_path}")

        frame: pd.DataFrame | None = None
        malformed = 0
        last_error: Exception | None = None
        for encoding in _ENCODING_FALLBACKS:
            skipped = 0

            def _skip_bad_line(bad_line: list[str]) -> None:
                nonlocal skipped
                skipped += 1
                return None

            try:
                frame = pd.read_csv(
                    file_path,
                    sep=self.delimiter,
                    encoding=encoding,
                    dtype=str,
                    engine="python",
                    on_bad_lines=_skip_bad_line,
                    keep_default_na=False,
                    skip_blank_lines=True,
                )
                malformed = skipped
                break
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
            except pd.errors.EmptyDataError as exc:
                raise DataIngestionError("Raw data file is empty") from exc
        else:
            raise DataIngestionError(
                f"Unable to decode {file_path} with any known encoding"
            ) from last_error

        if frame is None or (frame.empty and len(frame.columns) == 0):
            raise DataIngestionError("Raw data file is empty")

        frame.columns = self._sanitize_headers([str(c) for c in frame.columns])
        if not any(str(name).strip() for name in frame.columns):
            raise DataIngestionError("Raw data file contains no columns")

        for column in frame.columns:
            frame[column] = frame[column].astype(str).str.strip()

        nonempty = None
        for column in frame.columns:
            flag = frame[column].ne("")
            nonempty = flag if nonempty is None else nonempty | flag
        if nonempty is not None:
            frame = frame.loc[nonempty].reset_index(drop=True)

        frame = self._coerce_numeric_columns(frame)

        return IngestionResult(
            data=frame,
            total_rows=len(frame),
            malformed_rows_count=malformed,
        )

    @staticmethod
    def _sanitize_headers(headers: list[str]) -> list[str]:
        """Strips whitespace and uppercases header names for stable joins."""
        return [h.strip().upper() for h in headers]

    def _coerce_numeric_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Converts the configured numeric columns to float64 (bad values -> 0.0)."""
        for column in self.numeric_columns:
            if column in df.columns:
                cleaned = (
                    df[column]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )
                df[column] = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)
        return df

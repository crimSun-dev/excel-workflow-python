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

    def __init__(self, delimiter: str = "|"):
        self.delimiter = delimiter

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

        raw_text = self._read_with_encoding_fallback(file_path)
        if not raw_text.strip():
            raise DataIngestionError("Raw data file is empty")

        lines = [ln for ln in raw_text.splitlines() if ln.strip()]
        if not lines:
            raise DataIngestionError("Raw data file is empty")

        header = self._sanitize_headers(lines[0].split(self.delimiter))
        expected_cols = len(header)
        if expected_cols == 0:
            raise DataIngestionError("Raw data file contains no columns")

        rows: list[list[str]] = []
        malformed = 0
        for line in lines[1:]:
            fields = [f.strip() for f in line.split(self.delimiter)]
            # Malformed rows have a differing field count (e.g. extra pipes,
            # which would otherwise shift columns). Skip and count them.
            if len(fields) != expected_cols:
                malformed += 1
                continue
            rows.append(fields)

        df = pd.DataFrame(rows, columns=header)
        df = self._coerce_numeric_volume(df)

        return IngestionResult(
            data=df,
            total_rows=len(df),
            malformed_rows_count=malformed,
        )

    @staticmethod
    def _read_with_encoding_fallback(file_path: Path) -> str:
        """Attempts a list of encodings; returns decoded text of the first hit."""
        last_error: Exception | None = None
        for encoding in _ENCODING_FALLBACKS:
            try:
                return file_path.read_text(encoding=encoding)
            except (UnicodeDecodeError, UnicodeError) as exc:
                last_error = exc
                continue
        raise DataIngestionError(
            f"Unable to decode {file_path} with any known encoding"
        ) from last_error

    @staticmethod
    def _sanitize_headers(headers: list[str]) -> list[str]:
        """Strips whitespace and uppercases header names for stable joins."""
        return [h.strip().upper() for h in headers]

    @staticmethod
    def _coerce_numeric_volume(df: pd.DataFrame) -> pd.DataFrame:
        """Converts the volume column to float64, defaulting bad values to 0.0."""
        if VOLUME_COLUMN in df.columns:
            cleaned = (
                df[VOLUME_COLUMN]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[VOLUME_COLUMN] = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)
        return df

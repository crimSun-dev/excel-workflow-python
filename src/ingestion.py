"""Ingestion Engine (TDD Section 3.2).

Replicates Excel's "Text to Columns" feature: reads a pipe-delimited raw text
file (or an Excel workbook saved from one) into a structured DataFrame with
automatic encoding fallback, whitespace trimming, header sanitization, and
numeric coercion of the volume column.

Vendor extracts sometimes put column titles on a row other than the first —
a title banner above the table, or the header row inserted in the middle of
the values. Header detection scans *every* row for cells that match known
column titles (the same "column presence is the signal" rule used for
multi-sheet reference workbooks), then keeps the values both above and below
that row.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

import pandas as pd

from .schemas import IngestionResult

# Column expected to hold the monetary volume; coerced to float64.
VOLUME_COLUMN = "VOLUME_IN_IDR"

# Encodings attempted in order. utf-8-sig transparently strips a BOM if present.
_ENCODING_FALLBACKS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
_EXCEL_SUFFIXES = (".xlsx", ".xls", ".xlsm")

# A row is treated as column titles when this many cells match known headers.
# Two exact hits is enough to beat a data row (codes, amounts, dates) without
# requiring every required column to be present under one spelling.
HEADER_SCORE_THRESHOLD = 2
# Banner/title rows that survive as a single filled cell are dropped so they
# do not become a fake branch in the Summary.
_MIN_NONEMPTY_CELLS = 2

EMPTY_INPUT_MESSAGE = (
    "The file submitted for processing is empty. "
    "It has no data rows (only column titles, or nothing at all). "
    "Nothing was processed."
)

# Known extract / report headers. Exact cell match after trim+upper, so a
# title like "Rincian Volume" is not a hit and a real `VOL` / `KODE UKER`
# cell is. Workflow-required columns are unioned in at call time.
DEFAULT_HEADER_HINTS: tuple[str, ...] = (
    "KODE_UKER",
    "KODE UKER",
    "KODE_UNIT",
    "KODE UNIT",
    "KANCA",
    "MAINBR",
    "MBNAME",
    "MBDESC",
    "BRNAME",
    "MAIN_CODE",
    "MAIN_BRANCH",
    "AMOUNT_IN_IDR",
    "AMOUNT_ORI",
    "AMOUNT",
    "VOLUME_IN_IDR",
    "VOLUME_IDR",
    "VOLUME",
    "VOL",
    "SEGMEN",
    "FBI",
    "SOURCE",
    "ID",
    "FREKUENSI",
    "TAHUN",
    "PERIODE",
    "REGION",
    "CUST_NAME",
    "CUSTOMER_NAME",
    "CUSTOMER_CIF",
    "NO REK",
    "SALDO IDR",
    "SALDO UPDATE",
    "KW",
    "ID_PRODUCT",
    "BRANCH",
    "DIVISI",
    "NOMOR_BG",
    "CURRENCY",
    "STATUS",
    "ISSUEDATE",
    "PNRM",
    "RMNAME",
)


class DataIngestionError(Exception):
    """Raised when raw file parsing encounters unrecoverable structural errors."""


def header_recovery_note(header_row: int) -> str | None:
    """Operator-facing note when titles were not on the first row, else None."""
    if header_row <= 1:
        return None
    return (
        f"Column titles were found on row {header_row} instead of the first "
        "row. Values above and below that row were kept as data."
    )


def _normalize_cell(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip().upper()


def _hint_set(extra_hints: Sequence[str] = ()) -> set[str]:
    return {
        _normalize_cell(name)
        for name in (*DEFAULT_HEADER_HINTS, *extra_hints)
        if _normalize_cell(name)
    }


def _header_score(cells: Sequence[object], hints: set[str]) -> int:
    return sum(1 for cell in cells if _normalize_cell(cell) in hints)


def find_header_row_index(
    rows: Sequence[Sequence[object]], extra_hints: Sequence[str] = ()
) -> int:
    """0-based index of the first row that looks like column titles, else 0."""
    hints = _hint_set(extra_hints)
    for index, row in enumerate(rows):
        if _header_score(row, hints) >= HEADER_SCORE_THRESHOLD:
            return index
    return 0


def header_row_was_matched(
    rows: Sequence[Sequence[object]], extra_hints: Sequence[str] = ()
) -> bool:
    """True when some row actually looks like titles, not just 'use row 0'."""
    if not rows:
        return False
    hints = _hint_set(extra_hints)
    index = find_header_row_index(rows, extra_hints)
    return _header_score(rows[index], hints) >= HEADER_SCORE_THRESHOLD


def _drop_non_data_rows(
    data: pd.DataFrame, extra_hints: Sequence[str] = ()
) -> pd.DataFrame:
    """Drops stray header rows and one-cell title banners, keeping real values."""
    if data.empty:
        return data
    hints = _hint_set(extra_hints)
    normalized = data.astype(str).map(lambda value: str(value).strip().upper())
    score = normalized.isin(hints).sum(axis=1)
    blanked = normalized.replace(
        {"NAN": "", "NONE": "", "NAT": "", "<NA>": "", "NATYPE": ""}
    )
    nonempty = blanked.ne("").sum(axis=1)
    keep = (score < HEADER_SCORE_THRESHOLD) & (nonempty >= _MIN_NONEMPTY_CELLS)
    return data.loc[keep]


def promote_header_row(
    frame: pd.DataFrame, extra_hints: Sequence[str] = ()
) -> tuple[pd.DataFrame, int]:
    """Use a title row anywhere in the frame as columns; keep values on both sides.

    Returns `(frame, 1-based header row)`. When no row looks like titles, row 1
    is used — the same as pandas `header=0`.
    """
    if frame is None or frame.empty:
        return frame if frame is not None else pd.DataFrame(), 1
    rows = [tuple(row) for row in frame.itertuples(index=False, name=None)]
    header_idx = find_header_row_index(rows, extra_hints)
    columns = [
        ""
        if (value is None or (isinstance(value, float) and pd.isna(value)))
        else str(value).strip()
        for value in rows[header_idx]
    ]
    above = frame.iloc[:header_idx]
    below = frame.iloc[header_idx + 1 :]
    data = pd.concat([above, below], ignore_index=True)
    if data.empty:
        return pd.DataFrame(columns=columns), header_idx + 1
    data.columns = columns
    data = _drop_non_data_rows(data, extra_hints)
    return data.reset_index(drop=True), header_idx + 1


class IngestionEngine:
    """Parses pipe-delimited raw financial data files into structured DataFrames."""

    def __init__(
        self,
        delimiter: str = "|",
        numeric_columns: tuple[str, ...] | list[str] = (VOLUME_COLUMN,),
        header_hints: tuple[str, ...] | list[str] = (),
    ):
        self.delimiter = delimiter
        # Workflow-specific monetary columns to coerce to float64 (e.g. the
        # Akumulasi workflow uses VOLUME_IN_IDR, the Rincian workflows use
        # AMOUNT_IN_IDR). Only columns actually present in the file are coerced.
        self.numeric_columns = tuple(numeric_columns)
        self.header_hints = tuple(header_hints)

    def read_raw_data(self, file_path: Path) -> IngestionResult:
        """Reads pipe-delimited text or an Excel workbook of the same table.

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
        if file_path.suffix.lower() in _EXCEL_SUFFIXES:
            return self._read_excel(file_path)
        return self._read_delimited(file_path)

    def _read_excel(self, file_path: Path) -> IngestionResult:
        """Reads .xlsx/.xls/.xlsm with the header row discovered anywhere."""
        try:
            excel_file = pd.ExcelFile(file_path)
        except Exception as exc:  # noqa: BLE001 - re-wrapped for the caller
            raise DataIngestionError(
                f"Failed to read Excel file {file_path}: {exc}"
            ) from exc

        fallback: IngestionResult | None = None
        for sheet_name in excel_file.sheet_names:
            frame = pd.read_excel(
                excel_file, sheet_name=sheet_name, header=None, dtype=str
            )
            if frame is None or frame.empty:
                continue
            rows = [tuple(row) for row in frame.itertuples(index=False, name=None)]
            matched = header_row_was_matched(rows, self.header_hints)
            data, header_row = promote_header_row(frame, self.header_hints)
            try:
                data = self._finalize_frame(data)
            except DataIngestionError:
                continue
            result = IngestionResult(
                data=data,
                total_rows=len(data),
                malformed_rows_count=0,
                header_row=header_row,
            )
            if matched:
                return result
            if fallback is None:
                fallback = result
        if fallback is not None:
            return fallback
        raise DataIngestionError(EMPTY_INPUT_MESSAGE)

    def _read_delimited(self, file_path: Path) -> IngestionResult:
        frame: pd.DataFrame | None = None
        malformed = 0
        header_row = 1
        last_error: Exception | None = None
        for encoding in _ENCODING_FALLBACKS:
            skipped = 0

            def _skip_bad_line(bad_line: list[str]) -> None:
                nonlocal skipped
                skipped += 1
                return None

            try:
                header_idx, header_cells = self._scan_delimited_header(
                    file_path, encoding
                )
                if header_idx == 0:
                    # Happy path: titles are on row 1. Keep pandas' native
                    # header=0 + on_bad_lines behaviour so malformed extra-pipe
                    # rows still skip-and-count the way they always have.
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
                    header_row = 1
                    if frame is not None and len(frame.columns):
                        frame.columns = self._sanitize_headers(
                            [str(c) for c in frame.columns]
                        )
                else:
                    n_cols = max(len(header_cells), 1)
                    frame = pd.read_csv(
                        file_path,
                        sep=self.delimiter,
                        encoding=encoding,
                        dtype=str,
                        engine="python",
                        header=None,
                        names=list(range(n_cols)),
                        on_bad_lines=_skip_bad_line,
                        keep_default_na=False,
                        skip_blank_lines=False,
                    )
                    header_row = header_idx + 1
                    if header_idx < len(frame.index):
                        frame = frame.drop(index=frame.index[header_idx])
                    padded = list(header_cells) + [""] * (n_cols - len(header_cells))
                    frame.columns = self._sanitize_headers(padded[:n_cols])
                    frame = _drop_non_data_rows(frame, self.header_hints)
                malformed = skipped
                break
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
            except pd.errors.EmptyDataError as exc:
                raise DataIngestionError(EMPTY_INPUT_MESSAGE) from exc
        else:
            raise DataIngestionError(
                f"Unable to decode {file_path} with any known encoding"
            ) from last_error

        if frame is None or (frame.empty and len(frame.columns) == 0):
            raise DataIngestionError(EMPTY_INPUT_MESSAGE)

        frame = self._finalize_frame(frame)
        return IngestionResult(
            data=frame,
            total_rows=len(frame),
            malformed_rows_count=malformed,
            header_row=header_row,
        )

    def _scan_delimited_header(
        self, file_path: Path, encoding: str
    ) -> tuple[int, list[str]]:
        """(0-based index, cells) of the header row; falls back to the first line."""
        first_cells: list[str] | None = None
        with file_path.open(encoding=encoding, newline="") as handle:
            for index, line in enumerate(handle):
                cells = line.rstrip("\r\n").split(self.delimiter)
                if first_cells is None:
                    first_cells = cells
                if _header_score(cells, _hint_set(self.header_hints)) >= HEADER_SCORE_THRESHOLD:
                    return index, cells
        return 0, first_cells or []

    def _finalize_frame(self, frame: pd.DataFrame) -> pd.DataFrame:
        """Sanitizes headers/values, drops blank rows, coerces numeric columns."""
        frame = frame.copy()
        frame.columns = self._sanitize_headers([str(c) for c in frame.columns])
        if not any(str(name).strip() for name in frame.columns):
            raise DataIngestionError("Raw data file contains no columns")

        for column in frame.columns:
            frame[column] = frame[column].astype(str).str.strip()
            frame[column] = frame[column].replace({"nan": "", "None": "", "NaT": ""})

        nonempty = None
        for column in frame.columns:
            flag = frame[column].ne("")
            nonempty = flag if nonempty is None else nonempty | flag
        if nonempty is not None:
            frame = frame.loc[nonempty].reset_index(drop=True)
        else:
            frame = frame.reset_index(drop=True)

        if frame.empty:
            raise DataIngestionError(EMPTY_INPUT_MESSAGE)

        return self._coerce_numeric_columns(frame)

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

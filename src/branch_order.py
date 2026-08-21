"""Dashboard branch ordering for report summaries.

The dashboard the operator pastes each report into lists the 25 Malang-region
branches in one fixed sequence (alphabetical by branch name, with KC Batu last).
Summaries used to come out ordered by branch code *as text* - 110, 13, 177, 21 -
so every paste had to be re-ordered by hand before the numbers lined up with the
dashboard rows. Emitting the summary in the dashboard sequence makes the paste a
straight copy.

Codes the dashboard does not list (a national extract, `UNMAPPED`, a new unit)
are never dropped or hidden mid-table: they keep their previous ascending order
and follow the 25 known branches, so an unexpected code stays visible at the
bottom where it reads as "not part of the dashboard".
"""

from __future__ import annotations

import re

import pandas as pd

# The dashboard's NO / CABANG / KODE UKER sequence, by KODE UKER. Order is the
# contract: this tuple *is* the row order every report's Summary is written in.
DASHBOARD_BRANCH_CODES: tuple[str, ...] = (
    "7",  # Banyuwangi
    "9",  # Blitar
    "13",  # Bondowoso
    "577",  # Genteng
    "21",  # Jember
    "33",  # Kediri
    "516",  # Kepanjen
    "44",  # Lumajang
    "45",  # Madiun
    "49",  # Magetan
    "51",  # Malang Kawi
    "344",  # Malang Marthadinata
    "579",  # Malang Soekarno Hatta
    "429",  # Malang Sutoyo
    "56",  # Nganjuk
    "57",  # Ngawi
    "67",  # Pacitan
    "555",  # Pare
    "65",  # Pasuruan
    "70",  # Ponorogo
    "73",  # Probolinggo
    "90",  # Situbondo
    "177",  # Trenggalek
    "110",  # Tulungagung
    "551",  # Batu
)

# Temporary sort column. Double-underscored so it cannot collide with a real
# extract header, and always dropped before the frame is returned.
_POSITION_COLUMN = "__dashboard_position__"

_TRAILING_DOT_ZERO_RE = re.compile(r"^(\d+)\.0+$")


def canonical_branch_code(value: object) -> str:
    """Comparison form for a branch code: trimmed, no `.0` tail, no zero padding.

    The same code reaches this function as `7`, `"7"`, `" 007 "` or `"7.0"`
    depending on whether it came from a pipe extract, a reference workbook, or
    an Excel int -> float round-trip. Non-numeric values (`UNMAPPED`, `MC10`)
    are only trimmed and upper-cased.
    """
    text = str(value).strip()
    match = _TRAILING_DOT_ZERO_RE.match(text)
    if match:
        text = match.group(1)
    if text.isdigit():
        text = text.lstrip("0") or "0"
    return text.upper()


_POSITIONS: dict[str, int] = {
    canonical_branch_code(code): index
    for index, code in enumerate(DASHBOARD_BRANCH_CODES)
}

# Every unlisted code shares this one position, so the tie-break columns decide
# their order among themselves and they all land after the dashboard's 25.
UNKNOWN_POSITION = len(DASHBOARD_BRANCH_CODES)


def dashboard_position(value: object) -> int:
    """0-based dashboard row for a branch code; `UNKNOWN_POSITION` if unlisted."""
    return _POSITIONS.get(canonical_branch_code(value), UNKNOWN_POSITION)


def sort_by_dashboard_order(
    frame: pd.DataFrame,
    code_column: str,
    tie_break_columns: tuple[str, ...] | list[str] = (),
) -> pd.DataFrame:
    """Returns `frame` re-ordered so `code_column` follows the dashboard sequence.

    `tie_break_columns` order the rows that share a position - in practice the
    unlisted codes, which therefore keep the plain ascending order they had
    before. A frame that is empty or has no `code_column` is returned untouched,
    so this is a no-op for workflows without a branch-code summary.
    """
    if frame.empty or code_column not in frame.columns:
        return frame
    tie_breaks = [c for c in tie_break_columns if c in frame.columns]
    ordered = frame.assign(**{_POSITION_COLUMN: frame[code_column].map(dashboard_position)})
    ordered = ordered.sort_values(
        by=[_POSITION_COLUMN, *tie_breaks], ascending=True, kind="stable"
    )
    return ordered.drop(columns=[_POSITION_COLUMN]).reset_index(drop=True)

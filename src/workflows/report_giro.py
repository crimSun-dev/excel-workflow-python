"""Report Giro master-balance update workflow.

Replaces the operator's manual Ctrl+F reconciliation: for every account in the
Giro master workbook, look up that month's IDR balance in the monthly giro
source workbook and write it into the target column (`SALDO UPDATE`).

This workflow does not fit the shared ingest -> aggregate -> export template and
so overrides `execute()` (the precedent set by Qlola):

* the raw/source input is an Excel workbook, not pipe-delimited text;
* the output is the *master workbook itself*, updated cell-by-cell with
  openpyxl and saved to the configured output path, so historical columns,
  styling, formulas, and unrelated sheets all survive. A pandas `to_excel`
  round-trip would destroy every one of them;
* there is no Summary_Report pivot.

Business rules (see the `workflow-processing` spec):
    * `JENIS = Deposito` rows are never updated (case-insensitive).
    * Master accounts absent from the monthly source are left blank - the
      `UNMAPPED` sentinel must never land in a numeric balance cell.
    * The historical `SALDO` column is never written to.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..aggregation import apply_segmen_filters
from ..enrichment import canonicalize_join_id, normalize_header, resolve_alias_column
from ..exporter import PLAIN_INTEGER_FORMAT, ExportError
from ..schemas import ProcessingConfig
from .base import (
    WorkflowDefinition,
    WorkflowId,
    WorkflowRunResult,
    WorkflowStrategy,
    WorkflowValidationError,
)

# Ordered alias registries (first match wins), mirroring `enrichment.py`. Both
# sides are trimmed and upper-cased before comparison, so `no rek` matches
# `NO REK` while the spaced/underscored spellings stay distinct entries.
ACCOUNT_ALIASES = [
    "NO REK",
    "NO_REK",
    "NOREK",
    "NO REKENING",
    "NO_REKENING",
    "NOMOR REKENING",
    "REKENING",
]
JENIS_ALIASES = ["JENIS", "JENIS REKENING", "TIPE"]
TARGET_COLUMN_ALIASES = ["SALDO UPDATE", "SALDO_UPDATE"]
MONTHLY_BALANCE_ALIASES = [
    "SALDO IDR",
    "SALDO_IDR",
    "SALDO IN IDR",
    "SALDO_IN_IDR",
    "SALDO",
]

# JENIS value that is excluded from the update entirely.
DEPOSITO = "deposito"

# Rows scanned when hunting for the header row, so a master with a title banner
# above the table still resolves without hardcoded column letters or offsets.
_HEADER_SCAN_ROWS = 10

_EXCEL_SUFFIXES = (".xlsx", ".xls", ".xlsm")

REPORT_GIRO_DEFINITION = WorkflowDefinition(
    workflow_id=WorkflowId.REPORT_GIRO,
    label="Report Giro",
    requires_reference=False,
    # The "master" input is the Giro master workbook that gets updated; the
    # "raw" input is the monthly source it is reconciled against.
    requires_master_data=True,
    group_cols=(),
    value_col="SALDO UPDATE",
    report_title="Report Giro",
    detail_sheet_name="",
    # Concatenated digits, no thousands separator and no decimals - the same
    # plain-integer look every other report ships (Portal BG / Vol TF).
    number_format=PLAIN_INTEGER_FORMAT,
    # The monthly source usually carries only NO REK + SALDO IDR, but when it
    # does have SEGMEN/SOURCE the operator gets the same live controls as the
    # other reports. Defaults stay empty so nothing is dropped unasked.
    supports_segment_filter=True,
    has_source_filter=True,
    raw_button_label="1. Monthly Giro File...",
    raw_picker_title="Select the monthly giro source workbook",
    raw_filetypes=(("Excel Workbook", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")),
    master_button_label="2b. Giro Master File...",
    master_picker_title="Select the target giro master workbook",
)


class GiroColumnResolutionError(WorkflowValidationError):
    """Raised when the account/balance columns cannot be resolved by alias."""


class ReportGiroStrategy(WorkflowStrategy):
    """Copies monthly IDR balances into the Giro master's target month column."""

    @property
    def definition(self) -> WorkflowDefinition:
        return REPORT_GIRO_DEFINITION

    def execute(self, config: ProcessingConfig) -> WorkflowRunResult:
        # 0. The master workbook is the thing being updated; without it there is
        # nothing to write into, so fail before reading anything.
        if config.master_data_path is None:
            raise WorkflowValidationError(
                "The 'Report Giro' workflow requires a master-data file (the "
                "Giro master workbook to update), but none was provided."
            )
        master_path = Path(config.master_data_path)
        if not master_path.exists():
            raise WorkflowValidationError(
                f"Giro master workbook not found: {master_path}"
            )

        # 1. Build the monthly {account -> saldo} lookup (VLOOKUP first hit),
        # honoring the operator's SEGMEN/SOURCE filters where those columns exist.
        balances = self._load_monthly_balances(Path(config.raw_data_path), config)

        # 2. Update the master in memory, then save to the output path. Nothing
        # is written until every row succeeded, so a resolution failure can
        # never leave a partial output that looks like a successful run.
        workbook = load_workbook(master_path)
        worksheet, header_row, columns = self._resolve_master_layout(workbook)
        scanned, unmatched = self._apply_balances(
            worksheet, header_row, columns, balances
        )

        output_path = Path(config.output_report_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            workbook.save(output_path)
        except PermissionError as exc:
            raise ExportError(
                f"Cannot write '{output_path}'. The file may be open in Microsoft "
                f"Excel. Close it or choose a different output path."
            ) from exc

        return WorkflowRunResult(
            output_path=output_path,
            total_records_processed=scanned,
            unmapped_records_count=unmatched,
        )

    # ------------------------------------------------------------------ #
    # Monthly source
    # ------------------------------------------------------------------ #
    def _load_monthly_balances(
        self, monthly_path: Path, config: ProcessingConfig
    ) -> dict[str, float]:
        """Returns {canonical account: saldo} from the monthly source workbook.

        Every sheet is scanned; the first one resolving both an account column
        and a balance column wins. Duplicate accounts keep the first hit,
        matching VLOOKUP.

        The operator's SOURCE exclusion and SEGMEN keep-only/exclude filters are
        applied to the winning sheet *before* the lookup is built, so a filtered
        account simply never matches and its master cell is left as it was. A
        monthly extract without those columns makes the filters a no-op (the
        Portal BG / Qlola precedent), so typing a filter can never fail a run.
        """
        if not monthly_path.exists():
            raise WorkflowValidationError(
                f"Monthly giro source file not found: {monthly_path}"
            )

        frames = self._read_source_frames(monthly_path)
        attempts: list[tuple[str, list[str]]] = []
        for sheet_name, frame in frames:
            account_col = resolve_alias_column(frame, ACCOUNT_ALIASES)
            balance_col = resolve_alias_column(frame, MONTHLY_BALANCE_ALIASES)
            if account_col is None or balance_col is None:
                attempts.append((sheet_name, [str(c) for c in frame.columns]))
                continue
            frame = self.apply_source_exclude(
                frame, self.resolve_source_exclude(config)
            )
            frame = apply_segmen_filters(
                frame,
                self.resolve_segment_filter(config),
                self.resolve_segmen_exclude(config),
            )
            return self._mapping_from_columns(frame, account_col, balance_col)

        raise GiroColumnResolutionError(
            "Unable to resolve the account and balance columns in the monthly "
            f"giro source {monthly_path}.\n"
            f"Sheets scanned:\n{self._format_attempts(attempts)}\n"
            "Aliases attempted:\n"
            f"  - account: {ACCOUNT_ALIASES}\n"
            f"  - balance: {MONTHLY_BALANCE_ALIASES}"
        )

    @staticmethod
    def _read_source_frames(path: Path) -> list[tuple[str, pd.DataFrame]]:
        """Reads the monthly source as (sheet name, frame) pairs, all as text.

        Values are read with `dtype=str` so Excel's int/float storage cannot
        mangle the account key before `canonicalize_join_id` sees it.
        """
        suffix = path.suffix.lower()
        try:
            if suffix in _EXCEL_SUFFIXES:
                excel_file = pd.ExcelFile(path)
                frames = []
                for sheet_name in excel_file.sheet_names:
                    frame = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
                    frames.append((str(sheet_name), frame))
                return frames
            if suffix == ".csv":
                return [("(csv)", pd.read_csv(path, dtype=str))]
        except Exception as exc:  # noqa: BLE001 - re-wrapped for the caller
            raise WorkflowValidationError(
                f"Failed to read monthly giro source {path}: {exc}"
            ) from exc
        raise WorkflowValidationError(
            f"Unsupported monthly giro source file type: '{suffix}'. "
            "Expected an Excel workbook (.xlsx/.xls/.xlsm) or .csv."
        )

    @staticmethod
    def _mapping_from_columns(
        frame: pd.DataFrame, account_col: str, balance_col: str
    ) -> dict[str, float]:
        """Builds {canonical account: numeric saldo}, first hit wins."""
        pairs = frame[[account_col, balance_col]].copy()
        pairs[account_col] = pairs[account_col].apply(canonicalize_join_id)
        pairs[balance_col] = pd.to_numeric(pairs[balance_col], errors="coerce")
        pairs = pairs[
            (pairs[account_col] != "")
            & (pairs[account_col].str.lower() != "nan")
            & pairs[balance_col].notna()
        ]
        pairs = pairs.drop_duplicates(subset=account_col, keep="first")
        return {
            account: float(saldo)
            for account, saldo in zip(pairs[account_col], pairs[balance_col])
        }

    # ------------------------------------------------------------------ #
    # Master workbook
    # ------------------------------------------------------------------ #
    def _resolve_master_layout(self, workbook) -> tuple[Worksheet, int, dict[str, int]]:
        """Finds the sheet, header row, and 1-based column indexes to work with.

        The first sheet exposing both an account column and the target
        `SALDO UPDATE` column wins. Header detection scans the top rows so a title banner above
        the table does not break resolution, and no column letter is hardcoded.
        """
        attempts: list[tuple[str, list[str]]] = []
        for worksheet in workbook.worksheets:
            for header_row in range(1, min(_HEADER_SCAN_ROWS, worksheet.max_row) + 1):
                headers = self._header_map(worksheet, header_row)
                if not headers:
                    continue
                account_col = self._alias_index(headers, ACCOUNT_ALIASES)
                target_col = self._alias_index(headers, TARGET_COLUMN_ALIASES)
                if account_col is None or target_col is None:
                    continue
                return (
                    worksheet,
                    header_row,
                    {
                        "account": account_col,
                        "target": target_col,
                        # JENIS is optional: a master without it simply has no
                        # Deposito rows to skip.
                        "jenis": self._alias_index(headers, JENIS_ALIASES),
                    },
                )
            attempts.append(
                (worksheet.title, list(self._header_map(worksheet, 1).values()))
            )

        raise GiroColumnResolutionError(
            "Unable to resolve the account and target SALDO UPDATE columns in the Giro "
            "master workbook.\n"
            f"Sheets scanned:\n{self._format_attempts(attempts)}\n"
            "Aliases attempted:\n"
            f"  - account: {ACCOUNT_ALIASES}\n"
            f"  - target : {TARGET_COLUMN_ALIASES}"
        )

    @staticmethod
    def _header_map(worksheet: Worksheet, row: int) -> dict[int, str]:
        """{1-based column index: header text} for one candidate header row."""
        headers: dict[int, str] = {}
        for column in range(1, worksheet.max_column + 1):
            value = worksheet.cell(row, column).value
            if value is None or str(value).strip() == "":
                continue
            headers[column] = str(value)
        return headers

    @staticmethod
    def _alias_index(headers: dict[int, str], aliases: list[str]) -> int | None:
        """1-based index of the column matching the first alias that is present."""
        by_normalized: dict[str, int] = {}
        for column, text in headers.items():
            by_normalized.setdefault(normalize_header(text), column)
        for alias in aliases:
            match = by_normalized.get(normalize_header(alias))
            if match is not None:
                return match
        return None

    def _apply_balances(
        self,
        worksheet: Worksheet,
        header_row: int,
        columns: dict[str, int],
        balances: dict[str, float],
    ) -> tuple[int, int]:
        """Writes matched balances into the target column only.

        Returns (rows scanned, non-Deposito rows with no monthly match). Rows
        that are skipped or unmatched are left exactly as they were - the target
        cell is never blanked, and no sentinel is ever written.
        """
        account_col = columns["account"]
        target_col = columns["target"]
        jenis_col = columns["jenis"]

        scanned = 0
        unmatched = 0
        for row in range(header_row + 1, worksheet.max_row + 1):
            raw_account = worksheet.cell(row, account_col).value
            if raw_account is None or str(raw_account).strip() == "":
                continue  # trailing/blank spacer row, not a record
            scanned += 1

            if jenis_col is not None:
                jenis = worksheet.cell(row, jenis_col).value
                if str(jenis or "").strip().casefold() == DEPOSITO:
                    continue  # Deposito never receives an update

            saldo = balances.get(canonicalize_join_id(raw_account))
            if saldo is None:
                unmatched += 1
                continue  # left blank: no UNMAPPED sentinel in a balance cell

            cell = worksheet.cell(row, target_col)
            # Concatenated digits: rounded to a whole rupiah and written as a
            # real int, so the cell stays numeric rather than becoming a
            # preformatted string.
            cell.value = int(round(saldo))
            # Explicit numeric format so large IDR balances never render as
            # scientific notation or land as General-format text.
            cell.number_format = PLAIN_INTEGER_FORMAT

        return scanned, unmatched

    @staticmethod
    def _format_attempts(attempts: list[tuple[str, list[str]]]) -> str:
        return "\n".join(
            f"  - sheet '{name}': columns {cols}" for name, cols in attempts
        )

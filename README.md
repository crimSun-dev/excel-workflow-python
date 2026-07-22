# Excel Workflow Processor

Automated **Financial Data Processing & Excel Summarization Pipeline**.

This tool replaces a manual 4-step Microsoft Excel workflow (Text-to-Columns →
VLOOKUP → PivotTable → formatting) with a fast, testable, headless Python ETL
pipeline built with `pandas` + `openpyxl`. No Excel installation required.

It implements the design in [`TDD_Financial_Data_Processing.md`](./TDD_Financial_Data_Processing.md),
which itself automates the manual process described in
[`workflow_analysis_report.md`](./workflow_analysis_report.md).

## Quick start on Windows

Double-click **`Run Excel Workflow.bat`**. The launcher automatically creates
the Python environment, installs missing packages, and opens the GUI.

**Requirements:** Python **3.10–3.13** (64-bit, standard build). Avoid the
experimental **free-threaded** Python installer — it has no pre-built `pandas`
wheels on Windows and will fail with a Visual Studio / `vswhere.exe` error.
If setup fails on another PC, install [Python 3.12](https://www.python.org/downloads/),
delete the `.venv` folder, and run the batch file again. Or use the standalone
`.exe` build below (no Python required).

## What it does

| Manual Excel step | Automated equivalent | Module |
|---|---|---|
| Text to Columns (pipe `\|` delimiter) | Parse pipe-delimited raw file | `src/ingestion.py` |
| VLOOKUP against reference workbook | Left-join `KODE_UKER` → `MAIN_CODE`, `MAIN_BRANCH` | `src/enrichment.py` |
| PivotTable (Sum of `VOLUME_IN_IDR`) | Group-by + sum, optional `SEGMEN` filter | `src/aggregation.py` |
| Tabular layout, no subtotals, `#,##0.00` format | Styled `.xlsx` export + Grand Total | `src/exporter.py` |

Output is a two-sheet workbook:
- **Summary_Report** — tabular pivot view, corporate styling, Grand Total, explicit `#,##0.00` numeric format (no `E+12` scientific notation).
- **Enriched_Data** — full row-level audit trail.

## Install

```bash
pip install -r requirements.txt
```

## Usage

### CLI

```bash
python -m src.cli process --raw sample_data/raw_data.txt --ref sample_data/reference.xlsx --out report.xlsx
```

Options:

| Flag | Alias | Description |
|---|---|---|
| `--raw` | `-r` | Path to raw pipe-delimited `.txt` file (required) |
| `--ref` | `-f` | Path to reference mapping `.xlsx`/`.csv` (required) |
| `--out` | `-o` | Output report path (default `./Financial_Summary_Report.xlsx`) |
| `--segment` | `-s` | Optional `SEGMEN` filter (e.g. `Wholesale`, `Corporate`) |
| `--delimiter` | `-d` | Raw file delimiter (default `\|`) |
| `--gui` | `-g` | Launch the GUI file picker instead |

### GUI (no arguments)

```bash
python main.py
```

Launches a native Tkinter file-picker dialog — select the raw file, reference
workbook, optional filter, and output path, then click **Run Pipeline**.

## Input format

Raw file (`raw_data.txt`), pipe-delimited with a header row:

```
KODE_UKER|SEGMEN|VOLUME_IN_IDR
0001|Wholesale|1250000000000
0002|Corporate|500000000.50
```

Reference workbook (`reference.xlsx`) must contain a lookup key column plus a
code/branch pair. Canonical headers are `KODE_UKER`, `MAIN_CODE`, `MAIN_BRANCH`,
but common alternate names are resolved automatically (case-insensitive, first
match wins):

| Logical column | Accepted headers |
|---|---|
| Lookup key (`KODE_UKER`) | `KODE_UKER`, `KODE UNIT`, `KANCA`, `KODE SUB KANCA`, `KODE KANCA`, `UKER`, `KODE UKER` |
| `MAIN_CODE` | `MAIN_CODE`, `MAIN CODE`, `KODE`, `KODE KANCA`, `KODE UNIT`, `UNIQUE CODE`, `KODE INDUK` |
| `MAIN_BRANCH` | `MAIN_BRANCH`, `MAIN BRANCH`, `BRANCH`, `DESC KANCA`, `DESC UNIT`, `SUB KANCA`, `NAMA UKER`, `DESCRIPTION` |

For multi-sheet `.xlsx` workbooks, every sheet is scanned and the first sheet
whose columns resolve to all three logical columns is used. If no sheet
resolves, the error lists the sheets scanned, columns found, and aliases tried.
Codes missing from the reference are grouped under an explicit `UNMAPPED` branch
and reported (never abort the run).

## Tests

```bash
python -m pytest
```

26 tests cover ingestion (encoding fallback, whitespace, malformed rows),
enrichment (matching, unmapped flagging, column-alias/multi-sheet resolution,
error handling), aggregation
(grouping, decimal precision, segment filter), Excel export (sheets, number
format, Grand Total, column widths), and full end-to-end pipeline runs.

## Build a standalone .exe

```bash
pip install pyinstaller
python build_exe.py
```

Produces `dist/ExcelWorkflowProcessor.exe`, a single zero-dependency executable
that runs on a clean Windows machine without Python installed.

## Project layout

```
├── src/
│   ├── schemas.py        # Pydantic config + immutable result dataclasses
│   ├── ingestion.py      # Text-to-Columns equivalent
│   ├── enrichment.py     # VLOOKUP equivalent
│   ├── aggregation.py    # PivotTable equivalent
│   ├── exporter.py       # openpyxl styled report writer
│   ├── orchestrator.py   # Coordinates the ETL flow + telemetry
│   ├── cli.py            # Typer CLI
│   └── gui.py            # Tkinter file-picker fallback
├── tests/               # pytest suite (22 tests)
├── sample_data/         # Example raw + reference files
├── main.py              # Entry point (GUI when run with no args)
├── build_exe.py         # PyInstaller build script
└── requirements.txt
```

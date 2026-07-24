# Tongkat Gaib Excel

**by crimSun** — Automated **Financial Data Processing & Excel Summarization Pipeline**.

This tool replaces a manual 4-step Microsoft Excel workflow (Text-to-Columns →
VLOOKUP → PivotTable → formatting) with a fast, testable, headless Python ETL
pipeline built with `pandas` + `openpyxl`. No Excel installation required.

It implements the design in [`TDD_Financial_Data_Processing.md`](./TDD_Financial_Data_Processing.md),
which itself automates the manual process described in
[`workflow_analysis_report.md`](./workflow_analysis_report.md).

## Quick start on Windows

Double-click **`Run Tongkat Gaib Excel.bat`**. The launcher automatically creates
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
| PivotTable (Sum of value column) | Group-by + sum, optional `SEGMEN` include/exclude | `src/aggregation.py` |
| Tabular layout, no subtotals | Styled `.xlsx` export + Grand Total | `src/exporter.py` |

## Workflows

Pick a workflow in the GUI dropdown or via the CLI `--workflow` flag. Each is a
`WorkflowStrategy` (see `src/workflows/`) sharing one ingest → optional enrich →
aggregate → export skeleton, dispatched through `WORKFLOW_REGISTRY`.

| Workflow (`--workflow`) | Reference file | Grouped by | Sums | SEGMEN rule | Number format |
|---|---|---|---|---|---|
| Report Summary Akumulasi (`akumulasi`) | **Required** | `MAIN_CODE`, `MAIN_BRANCH` | `VOLUME_IN_IDR` | Optional include filter | `#,##0.00` |
| Rincian Vol TF (`rincian-vol-tf`) | Not used | `MAINBR`, `MBDESC` | `AMOUNT_IN_IDR` | Excludes `Wholesale` (blank kept) | `#,##0` |
| Rincian Portal BG (`rincian-portal-bg`) | Not used | `MAINBR`, `MBNAME` | `AMOUNT_IN_IDR` | None | `#,##0` |

Every workflow produces a two-sheet workbook: a **Summary_Report** (tabular
view, corporate styling, Grand Total, explicit numeric format — no `E+12`
scientific notation) plus a row-level detail sheet (**Enriched_Data** for
Akumulasi, **Detail_Data** for the Rincian workflows).

## Install

```bash
pip install -r requirements.txt
```

## Usage

### CLI

```bash
# Akumulasi (default) — reference file required
python -m src.cli process --raw sample_data/raw_data.txt --ref sample_data/reference.xlsx --out report.xlsx

# Rincian Vol TF — no reference file
python -m src.cli process --raw vol_tf.csv --workflow rincian-vol-tf --out vol_tf_report.xlsx

# Rincian Portal BG — no reference file
python -m src.cli process --raw portal_bg.csv --workflow rincian-portal-bg --out portal_bg_report.xlsx
```

Options:

| Flag | Alias | Description |
|---|---|---|
| `--raw` | `-r` | Path to raw pipe-delimited `.txt`/`.csv` file (required) |
| `--workflow` | `-w` | Workflow: `akumulasi` (default), `rincian-vol-tf`, `rincian-portal-bg` |
| `--ref` | `-f` | Path to reference mapping `.xlsx`/`.csv` (required for `akumulasi` only) |
| `--out` | `-o` | Output report path (default `./Financial_Summary_Report.xlsx`) |
| `--segment` | `-s` | Optional `SEGMEN` filter (Akumulasi only; e.g. `Wholesale`, `Corporate`) |
| `--delimiter` | `-d` | Raw file delimiter (default `\|`) |
| `--gui` | `-g` | Launch the GUI file picker instead |

### GUI (no arguments)

```bash
python main.py
```

Launches a native Tkinter file-picker dialog — pick a **workflow** from the
dropdown, select the raw file, (for Akumulasi) a reference workbook, an optional
filter, and an output path, then click **Run Pipeline**. The reference-file row
is shown only for the Akumulasi workflow.

## Input format

All workflows read a pipe-delimited file with a header row. The required
columns depend on the selected workflow:

**Akumulasi** (`raw_data.txt`):

```
KODE_UKER|SEGMEN|VOLUME_IN_IDR
0001|Wholesale|1250000000000
0002|Corporate|500000000.50
```

**Rincian Vol TF** (`SEGMEN`, `MAINBR`, `MBDESC`, `AMOUNT_IN_IDR`; rows where
`SEGMEN` = `Wholesale` are excluded, blank/null `SEGMEN` rows are kept):

```
SEGMEN|MAINBR|MBDESC|AMOUNT_IN_IDR
Wholesale|B01|Branch One|1000
Corporate|B01|Branch One|2000
|B02|Branch Two|300
```

**Rincian Portal BG** (`MAINBR`, `MBNAME`, `AMOUNT_IN_IDR`; no segment filter):

```
MAINBR|MBNAME|AMOUNT_IN_IDR
B01|Branch One|1000
B02|Branch Two|500
```

The Akumulasi reference workbook (`reference.xlsx`) must contain a lookup key column plus a
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

40 tests cover ingestion (encoding fallback, whitespace, malformed rows,
per-workflow numeric coercion), enrichment (matching, unmapped flagging,
column-alias/multi-sheet resolution, error handling), aggregation (grouping,
decimal precision, segment include/exclude filters), the multi-workflow registry
and the two new workflows (Wholesale exclusion, blank-segment handling, no
filter, missing-column validation), Excel export (sheets, number format, Grand
Total, column widths), and full end-to-end pipeline runs.

## Build a standalone .exe

```bash
pip install pyinstaller
python build_exe.py
```

Produces `dist/TongkatGaibExcel.exe`, a single zero-dependency executable
that runs on a clean Windows machine without Python installed. The crimSun logo
is bundled into the exe (via PyInstaller `datas`), so no side-car assets folder
is required.

## Project layout

```
├── src/
│   ├── schemas.py        # Pydantic config + immutable result dataclasses
│   ├── ingestion.py      # Text-to-Columns equivalent
│   ├── enrichment.py     # VLOOKUP equivalent
│   ├── aggregation.py    # PivotTable equivalent (include/exclude SEGMEN)
│   ├── exporter.py       # openpyxl styled report writer (parameterized)
│   ├── orchestrator.py   # Exception boundary + telemetry; dispatches to a workflow
│   ├── workflows/        # Strategy pattern: WorkflowId, definitions, registry
│   │   ├── base.py           # WorkflowId, WorkflowDefinition, WorkflowStrategy ABC
│   │   ├── akumulasi.py      # Report Summary Akumulasi (enrichment-backed)
│   │   ├── rincian_vol_tf.py # Rincian Vol TF (excludes Wholesale)
│   │   ├── rincian_portal_bg.py  # Rincian Portal BG (no filter)
│   │   └── registry.py       # WORKFLOW_REGISTRY + get_strategy()
│   ├── cli.py            # Typer CLI (--workflow)
│   ├── branding.py       # Product name / byline / logo path (crimSun)
│   └── gui.py            # Tkinter file-picker fallback (workflow dropdown)
├── assets/              # crimSun logo (bundled into the exe)
├── tests/               # pytest suite
├── sample_data/         # Example raw + reference files
├── main.py              # Entry point (GUI when run with no args)
├── build_exe.py         # PyInstaller build script
└── requirements.txt
```

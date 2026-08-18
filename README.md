# Tongkat Gaib Excel

Automated **Financial Data Processing & Excel Summarization Pipeline**.

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
| PivotTable (Sum of value column) | Group-by + sum | `src/aggregation.py` |
| Ticking/unticking segments in the Filters pane | Editable `SEGMEN` include/exclude fields, pre-set per workflow | `src/aggregation.py` |
| Tabular layout, no subtotals | Styled `.xlsx` export + Grand Total | `src/exporter.py` |

## Workflows

Pick a workflow in the GUI dropdown or via the CLI `--workflow` flag. Each is a
`WorkflowStrategy` (see `src/workflows/`) sharing one ingest → optional enrich →
aggregate → export skeleton, dispatched through `WORKFLOW_REGISTRY`.

The table lists all seven workflows in the order they appear in the GUI dropdown
and the CLI `--workflow` help — the operator's daily sequence. That order comes
from the `WorkflowId` declaration order in `src/workflows/base.py` and nowhere
else, so the selector, the CLI listing, and this table cannot drift apart.

| # | Workflow (`--workflow`) | Reference file | Grouped by | Aggregates | Default SEGMEN rule | Number format |
|---|---|---|---|---|---|---|
| 1 | Report Summary Akumulasi (`akumulasi`) | **Required** | `MAIN_CODE`, `MAIN_BRANCH` | Sum of `FBI`, Sum of `VOLUME_IN_IDR` | None | `0` |
| 2 | Time Series Active User Qlola (`timeseries-active-user-qlola`) | **Required** (plus `--master`) | `MAIN_CODE` × `USER_AKTIF` crosstab | **Count** of distinct `ID`, active at `FREKUENSI` ≥ 5 | None (excludes `SOURCE` = `CMS`) | `0` |
| 3 | Report Data Statis (`report-data-statis`) | **Required** | `MAIN_CODE`, `MAIN_BRANCH` | **Count** of `ID_PRODUCT` | Excludes `KORPORASI` (blank kept); also keeps only `KW` = `KANWIL MALANG` | `0` |
| 4 | Rincian Vol TF (`rincian-vol-tf`) | Not used | `MAINBR`, `MBDESC` | Sum of `AMOUNT_IN_IDR` | Excludes `Wholesale` (blank kept) | `0` |
| 5 | Rincian Portal BG (`rincian-portal-bg`) | Not used | `MAINBR`, `MBNAME` | Sum of `AMOUNT_IN_IDR` | None | `0` |
| 6 | Report Vol Briva (`timeseries-fbi-briva`) | **Required** | `MAIN_CODE`, `MAIN_BRANCH` | Sum of `VOLUME_IDR` (rounded up) | Keeps only `NONWHOLESALE` | `0` |
| 7 | Report Giro (`report-giro`) | Not used (needs `--master`) | n/a — fills the master's `SALDO UPDATE` | Copies monthly `SALDO IDR` per account | None (skips `JENIS` = `Deposito`) | `0` |

The SEGMEN rule above is only each workflow's **default**. Every workflow
exposes the filters as live controls, so the operator decides per run
which segments go in and which come out — see
[Choosing what to include and exclude](#choosing-what-to-include-and-exclude).

Every workflow except Report Giro produces a two-sheet workbook: a
**Summary_Report** (tabular view, corporate styling, Grand Total, explicit
numeric format — no `E+12` scientific notation) plus a row-level detail sheet
(**Enriched_Data** for Akumulasi, **Detail_Data** for the Rincian workflows).

### Report Giro (master-balance update)

Report Giro is the one workflow that does **not** build a pivot. It replaces the
manual Ctrl+F reconciliation: it reads the monthly giro workbook, then writes
each account's IDR balance into the `SALDO UPDATE` column of the giro master
workbook and saves the result to `--out`.

| Input | Flag | What it is |
|---|---|---|
| Monthly source | `--raw` | The month's giro extract (`.xlsx`/`.xls`/`.xlsm`/`.csv`), keyed by `NO REK` with a `SALDO IDR` balance column |
| Giro master | `--master` | The original giro list (e.g. `DAFTAR REKENING GIRO TSPM.xls`) or a later copy that already has `SALDO UPDATE` / `SALDO JUNI` |
| Output | `--out` | Where the updated copy is written — the master you select is **never** modified |

Rules the workflow enforces (each is covered by a regression test):

- The historical `SALDO` column is never written to. If `SALDO UPDATE` is
  missing, it is appended; leftover month-named columns such as `SALDO JUNI`
  are used only when they already exist.
- Rows whose `JENIS` is `Deposito` (case-insensitive) are skipped even when the
  account exists in the monthly file.
- Master accounts absent from the monthly file are left **blank** — the
  `UNMAPPED` sentinel never lands in a balance cell.
- Duplicate accounts in the monthly file take the first hit (VLOOKUP semantics),
  and keys match across Excel's int/float/text storage (`1234567890` matches
  `1234567890.0`).
- Other sheets, styling, and formulas in the master are preserved: the update is
  done cell-by-cell with openpyxl, never via a pandas export.
- Updated cells are numeric whole rupiah with number format `0` — concatenated
  digits, no thousands separator and no decimals, and never scientific notation.
- The SEGMEN and SOURCE filters apply to the **monthly source**: a filtered-out
  account simply never matches, so its master cell is left as it was. A monthly
  extract with only `NO REK` + `SALDO IDR` makes them a no-op.

Columns are resolved by alias (`NO REK` / `NO REKENING` / `NOREK` / `NO_REK`,
`SALDO IDR` / `SALDO_IDR` / `SALDO IN IDR`, `SALDO UPDATE` / `SALDO_UPDATE`) rather
than by fixed column letters, and a header row is located by scanning the top of
each sheet — so a title banner above the table is fine. The original five-column
master (`JENIS`, `NO REK`, `PRODUK`, `OPEN DATE`, `SALDO`) is enough: `SALDO UPDATE`
is created when absent. Legacy `.xls` masters are accepted; the updated copy is
always `.xlsx`. If no sheet resolves an account column the run fails loudly,
listing the sheets scanned and the aliases attempted, and writes no output file.

### Filter config keys

Four `WorkflowDefinition` keys (`src/workflows/base.py`) hold each workflow's
filter defaults and declare which controls it exposes.

| Key | Type | Default | Meaning |
|---|---|---|---|
| `segmen_include` | `str \| None` | `None` | SEGMEN value to keep, matched case-insensitively after trimming. Only Briva sets one (`NONWHOLESALE`). |
| `exclude_segmen` | `tuple[str, ...]` | `()` | SEGMEN values dropped before aggregation. Report Data Statis drops `KORPORASI`, Rincian Vol TF drops `Wholesale`. |
| `source_exclude` | `tuple[str, ...]` | `()` | SOURCE values dropped before enrichment. Only Qlola sets one (`CMS`). |
| `supports_segment_filter` / `has_source_filter` | `bool` | `False` | Declares that the workflow exposes that control. `True` for every workflow, including Report Giro (which applies them to its monthly source); a disabled field's value is ignored at runtime. |

At runtime the operator's choice wins over the default. `ProcessingConfig`
carries three matching overrides — `segmen_filter`, `segmen_exclude`, and
`source_exclude` — where `None` keeps the workflow default while an empty
string/list explicitly means "do not filter on this dimension".

### Reading the numbers (operator notes)

Two places where a report legitimately disagrees with a hand-built Excel cut.
Neither is a bug; both are listed here so a mismatch does not get re-litigated
every month.

**1. Qlola Summary counts mapped users only.**

The Time Series Active User Qlola **Summary_Report** counts only IDs that
resolved to a real `MAIN_CODE` in the master-data workbook. IDs the master does
not know are left out of the rows *and* out of the Grand Total — matching a
dashboard pivot that drops `#N/A`. Nothing is lost: those IDs still appear on
the **Enriched_Data** detail sheet with `MAIN_CODE = UNMAPPED`, and the Summary
metadata block carries an `Unmapped IDs:` line with the count. So a manual cut
that includes `#N/A` will read higher than the Summary by exactly that count.

If almost every ID fails to match, the GUI shows a **Master ID lookup failed**
warning after the run with sample IDs from both sides. The report is still
written — the warning means "check the master-data file you picked", not
"the run failed".

**2. Briva totals keep their cents.**

The Report Vol Briva Grand Total is stored unrounded with the `#,##0.00`
format — e.g. `2,481,517,421,603.75`. An operator who rounds the same figure in
Excel sees `…604`, because `ROUND(2481517421603.75, 0)` rounds the `.75` up. The
one-unit gap is the rounding, not a different sum. The software deliberately
stores the exact value so no cents are silently dropped; round it downstream if
your dashboard wants a whole number.

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

# Report Data Statis — reference file required; raw may be a .bin extract
python -m src.cli process --raw data_statis.bin --workflow report-data-statis --ref "Mapping GL FBI Vol Uker.xlsx" --out data_statis_report.xlsx

# Report Giro — --raw is the monthly workbook, --master is the giro master to update
python -m src.cli process --raw giro_20260630_monthly.xlsx --workflow report-giro --master "DAFTAR REKENING GIRO TSPM UPDATE JUNI.xlsx" --out giro_updated.xlsx
```

Options:

| Flag | Alias | Description |
|---|---|---|
| `--raw` | `-r` | Path to raw pipe-delimited `.txt`/`.csv`/`.bin` file (required); for `report-giro` the monthly giro `.xlsx` |
| `--workflow` | `-w` | Workflow, in selector order: `akumulasi` (default), `timeseries-active-user-qlola`, `report-data-statis`, `rincian-vol-tf`, `rincian-portal-bg`, `timeseries-fbi-briva`, `report-giro` |
| `--ref` | `-f` | Path to reference mapping `.xlsx`/`.csv` (required for `akumulasi`, `timeseries-active-user-qlola`, `report-data-statis`, and `timeseries-fbi-briva`) |
| `--master` | `-m` | Master workbook: `ID` → `MAIN_CODE` for `timeseries-active-user-qlola`, or the giro master to update for `report-giro` |
| `--out` | `-o` | Output report path (default `./Financial_Summary_Report.xlsx`) |
| `--segment` | `-s` | Comma-separated `SEGMEN` keep-list (e.g. `NONWHOLESALE`). Omit for the workflow default; pass `""` to keep every segment |
| `--exclude-segmen` | | Comma-separated `SEGMEN` values to drop (e.g. `KORPORASI,Wholesale`). Omit for the workflow default; pass `""` to drop none |
| `--source` | | Comma-separated `SOURCE` keep-list. Omit for no inclusion filter (the workflow's `SOURCE` default stays in charge) |
| `--exclude-source` | | Comma-separated `SOURCE` values to drop (e.g. `CMS`). Omit for the workflow default; pass `""` to drop none |
| `--kw` | | Comma-separated `KW` keep-list. Omit for the workflow default; pass `""` to keep every `KW` |
| `--delimiter` | `-d` | Raw file delimiter (default `\|`) |
| `--gui` | `-g` | Launch the GUI file picker instead |

### Choosing what to include and exclude

The manual pivot let the operator tick and untick segments per run, so the
automated filters are **defaults, not hardcoded rules**. Each workflow declares
what it normally does; the GUI's **FILTERS** block exposes one box per dimension
and the CLI flags above do the same headlessly.

The GUI contract is one rule, applied to all three boxes:

| Box | Empty (the default state) | Typed |
|---|---|---|
| `SEGMENT` | The workflow's baked rule runs (drop `KORPORASI`, keep `NONWHOLESALE`, …) | Keep **only** the listed segments; the baked rule is replaced for that run |
| `SOURCE` | The workflow's baked rule runs (Qlola drops `CMS`) | Keep **only** the listed `SOURCE` values |
| `KW` | The workflow's baked rule runs (Report Data Statis keeps `KANWIL MALANG`); no `KW` filtering elsewhere | Keep **only** the listed `KW` values |

Boxes are never pre-filled, because a filled box reads as something the operator
must manage and clearing it would read as "remove all filters". Instead each box
carries a hint (`auto: drops KORPORASI`) stating what the empty state does.

Every filter is case-insensitive and trimmed, they are AND-combined, and each is
a **no-op when the extract has no such column** — `KW` is resolved against the
`KW` header first, then the `PRODUCT`, `GROUP_PRODUCT`, `KAWIL` fallback aliases
for older extracts. Blank/null cells never
match an exclusion, so they survive it. Internally, inclusion and exclusion
remain separate fields (`segmen_include`, `exclude_segmen`, `source_exclude`,
`kw_include`) — there is no single filter with a negation flag.

So to include corporate accounts in a Report Data Statis run, type `KORPORASI`
in the `SEGMENT` box (or pass `--exclude-segmen ""` headlessly); leaving the box
empty keeps dropping `KORPORASI` exactly as before. The Summary sheet's
**Filter Applied** line always states the filters that run actually used.

The regional cut is part of the `KW` dimension, not a separate control: Report
Data Statis reads the region from the extract's `KW` column and keeps only
`KANWIL MALANG` when the box is empty. Type another region (or
`--kw "KANWIL SURABAYA"`) to cut a different one for a single run.

### GUI (no arguments)

```bash
python main.py
```

Launches a native Tkinter file-picker dialog — pick a **workflow** from the
dropdown, select the raw file, (where required) a reference workbook, adjust the
filters, and choose an output path, then click **Run Pipeline**. The
reference-file and master-data rows are shown only for the workflows that need
them.

The **FILTERS** block holds one box each for **SEGMENT**, **SOURCE**, and **KW**.
All three are cleared every time the dropdown changes, so an empty box always
means "this report's automatic rules" and no value can carry over between
workflows; the hint beside each box states what those rules are. They are live
for every workflow — Report Giro included, where they apply to the monthly
source and are simply a no-op when that extract has no such column. Input button
labels and file-type filters follow the selected workflow, so Report Giro offers
Excel workbook pickers instead of the text/CSV ones.

## Input format

Every workflow except Report Giro reads a pipe-delimited file with a header row
(Report Giro reads two Excel workbooks — see
[Report Giro](#report-giro-master-balance-update)). The required columns depend
on the selected workflow:

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

**Report Data Statis** (`KODE_UNIT`, `SEGMEN`, `KW`, `ID_PRODUCT`; commonly a
`.bin` extract, which is read as plain delimited text). `SEGMEN` = `KORPORASI`
rows are dropped (blank `SEGMEN` kept), only `KW` = `KANWIL MALANG` rows are
kept, and the summary is the **Count** of non-blank `ID_PRODUCT` per branch:

```
KODE_UNIT|SEGMEN|KW|ID_PRODUCT
0001|Consumer|KANWIL MALANG|P1
0001|KORPORASI|KANWIL MALANG|P2
0002|Micro|KANWIL SURABAYA|P3
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
that runs on a clean Windows machine without Python installed.

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
│   │   ├── akumulasi.py      # 1. Report Summary Akumulasi (enrichment-backed)
│   │   ├── timeseries_active_user_qlola.py  # 2. Qlola active-user crosstab
│   │   ├── report_data_statis.py # 3. Report Data Statis (Count of ID_PRODUCT)
│   │   ├── rincian_vol_tf.py # 4. Rincian Vol TF (excludes Wholesale)
│   │   ├── rincian_portal_bg.py  # 5. Rincian Portal BG (no filter)
│   │   ├── timeseries_fbi_briva.py  # 6. Briva NONWHOLESALE volume summary
│   │   ├── report_giro.py    # 7. Report Giro (openpyxl master-balance update)
│   │   └── registry.py       # WORKFLOW_REGISTRY + get_strategy()
│   ├── cli.py            # Typer CLI (--workflow)
│   ├── branding.py       # Product name
│   └── gui.py            # Tkinter file-picker fallback (workflow dropdown)
├── tests/               # pytest suite
├── sample_data/         # Example raw + reference files
├── main.py              # Entry point (GUI when run with no args)
├── build_exe.py         # PyInstaller build script
└── requirements.txt
```

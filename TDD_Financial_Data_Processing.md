# Technical Design Document (TDD)
## Automated Financial Data Processing & Excel Summarization Pipeline

**Author:** Antigravity AI  
**Date:** July 21, 2026  
**Status:** Draft / Approved for Implementation  
**Target Architecture:** Python 3.12 ETL Pipeline & Excel Exporter  

---

## TDD SECTION 1 — SYSTEM ARCHITECTURE OVERVIEW

### 1.1 Architectural Pattern

The system adopts a **Layered ETL (Extract, Transform, Load) Architecture** coupled with a **Command Pattern** for user invocation (CLI and GUI fallback). 

```
┌────────────────────────────────────────────────────────────────────────┐
│                        User Interface Layer                            │
│                 (Typer CLI / Tkinter GUI Fallback)                     │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ Invokes with ProcessingConfig
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                     Pipeline Orchestrator Layer                        │
└─────┬─────────────────────────────┬──────────────────────────────┬─────┘
      │                             │                              │
      ▼                             ▼                              ▼
┌──────────────┐           ┌─────────────────┐           ┌──────────────────┐
│   Ingestion  │           │   Enrichment    │           │   Aggregation    │
│    Engine    │──────────►│     Engine      │──────────►│      Engine      │
│ (Text/Pipe)  │           │   (VLOOKUP)     │           │  (Pivot/Filter)  │
└──────────────┘           └─────────────────┘           └────────┬─────────┘
                                                                  │
                                                                  ▼
                                                         ┌──────────────────┐
                                                         │  Report Exporter │
                                                         │  (OpenPyXL/Style)│
                                                         └──────────────────┘
```

#### Why this pattern was chosen over alternatives:
- **Over Microservices / Event-Driven:** The problem domain is fundamentally a synchronous, desktop-oriented batch processing task. An event broker or asynchronous microservice architecture would introduce network latency, deployment overhead, and state serialization costs without benefit.
- **Over Monolithic Excel COM Automation (win32com/xlwings):** Driving the Microsoft Excel desktop application programmatically via OLE/COM is slow, brittle, OS-dependent (requires Excel installed on Windows), and prone to GUI crash hangs. A pure Python data processing stack (`pandas` + `openpyxl`) runs 10x to 100x faster, is cross-platform, and runs in headless environments.
- **Why Layered ETL:** Decoupling file parsing, dictionary mapping, aggregation logic, and report formatting ensures each phase can be unit-tested independently with deterministic fixtures.

---

### 1.2 High-Level Component Diagram (ASCII)

```
                       INPUTS                                                       PIPELINE CORE                                              OUTPUT
┌───────────────────────────────────────────────────┐               ┌───────────────────────────────────────────┐           ┌───────────────────────────────────────────┐
│ Raw Data File (.txt)                              │               │ PipelineOrchestrator                      │           │ Formatted Excel Report (.xlsx)            │
│ Delimited by '|'                                  │               │                                           │           │                                           │
│ (e.g. KODE_UKER, SEGMEN, VOLUME_IN_IDR)           │               │ 1. Coordinates execution flow             │           │  Sheet 1: "Summary_Report"                │
└─────────────────────────┬─────────────────────────┘               │ 2. Captures telemetry & audit logs        │           │  - Tabular Pivot View (No subtotals)      │
                          │                                         └─────────────────────┬─────────────────────┘           │  - Standard Number Format (#,##0.00)     │
                          │                                                               │                                 │  - Auto-fitted Column Widths              │
                          ▼                                                               ▼                                 │  - Header & Total Highlight Styling       │
┌───────────────────────────────────────────────────┐               ┌───────────────────────────────────────────┐           │                                           │
│ IngestionEngine                                   │               │ ReferenceEnricher                         │           │  Sheet 2: "Enriched_Data"                 │
│                                                   │──────────────►│                                           │           │  - Complete Audit Trail Rows              │
│ - Auto-detects encoding (UTF-8 / CP1252)          │               │ - Ingests reference workbook mapping      │           └───────────────────────────────────────────┘
│ - Parses pipe-delimited columns                   │               │ - Left-joins KODE_UKER to fetch:          │                                 ▲
│ - Trims whitespace & sanitizes headers            │               │   * MAIN_CODE                             │                                 │
└───────────────────────────────────────────────────┘               │   * MAIN_BRANCH                           │                                 │
                                                                    │ - Identifies unmapped orphan codes        │                                 │
                                                                    └─────────────────────┬─────────────────────┘                                 │
                                                                                          │                                                       │
                                                                                          ▼                                                       │
                                                                    ┌───────────────────────────────────────────┐                                 │
                                                                    │ AggregationEngine                         │                                 │
                                                                    │                                           │─────────────────────────────────┘
                                                                    │ - Applies optional SEGMEN filter          │
                                                                    │ - Groups by [MAIN_CODE, MAIN_BRANCH]      │
                                                                    │ - Sums VOLUME_IN_IDR                      │
                                                                    │ - Computes Grand Total row                │
                                                                    └───────────────────────────────────────────┘
```

---

### 1.3 Data Flow Topology

1. **Extraction (Ingestion Phase):**
   - The user supplies `raw_data.txt` path via CLI arguments or GUI file picker.
   - `IngestionEngine` reads the raw file stream, handles potential BOM/encoding variations (`utf-8-sig`, `cp1252`), splits fields by `|`, strips leading/trailing whitespace from string columns, and converts `VOLUME_IN_IDR` to numeric standard (`float64`), dropping empty lines or reporting malformed rows.
2. **Transformation (Enrichment & Aggregation Phase):**
   - `ReferenceEnricher` loads the reference mapping Excel/CSV workbook (`reference.xlsx`).
   - Standardizes `KODE_UKER` data types in both datasets to string format to prevent type mismatch during join.
   - Performs a left outer join (`df_raw.merge(df_ref, on='KODE_UKER', how='left')`), appending `MAIN_CODE` and `MAIN_BRANCH`. Unmatched codes are assigned `"UNMAPPED"` and flagged in an audit payload.
   - `AggregationEngine` filters rows by `SEGMEN` if a filter is configured.
   - Groups data by `['MAIN_CODE', 'MAIN_BRANCH']` and calculates `sum('VOLUME_IN_IDR')`.
   - Sorts results by `MAIN_CODE` ascending.
3. **Loading (Report Formatting Phase):**
   - `ExcelReportExporter` writes the output to target `.xlsx` path.
   - Applies `openpyxl` custom styling: header fill (Corporate Blue `#1F4E79`), bold headers, explicit numeric formatting (`#,##0.00` to prevent scientific `E+12` representation), borders, and auto-adjusted column width logic.
   - Adds a summary metadata header block (Processing Timestamp, Total Records, Filter Applied).

---

### 1.4 State Management Contract

State within the system is **strictly immutable and unidirectional**. Stage outputs are encapsulated inside Pydantic immutable data classes.

```python
# State Flow Diagram:
# Config -> IngestionResult -> EnrichmentResult -> AggregationResult -> PipelineReport
```

- **Single Source of Truth:** The `PipelineOrchestrator` owns execution state.
- **Mutation Policy:** Pipeline components do not mutate input dataframes in place. Each component returns a new `dataclass` containing a fresh `pd.DataFrame` copy and associated metadata metrics.

---

## TDD SECTION 2 — TECHNOLOGY STACK & DEPENDENCY MATRIX

| Layer | Technology | Version | Justification |
|---|---|---|---|
| Runtime | Python | 3.12.x | Modern type hinting syntax (`X \| Y`), enhanced exception notes, fast execution |
| Data Processing | pandas | 2.2.x | Vectorized ETL engine, robust CSV parsing, fast table merging & group-by operations |
| Excel Engine | openpyxl | 3.1.x | Direct XML write engine for formatting `.xlsx` files without needing Excel installed |
| Data Validation | pydantic | 2.7.x | Schema validation for configuration data structures and execution options |
| User Interface (CLI) | typer | 0.12.x | Zero-boilerplate CLI framework with automatic `--help` flags and interactive prompts |
| User Interface (GUI) | tkinter | Built-in | Native OS drag-and-drop file picker fallback when user launches without terminal args |
| Testing Framework | pytest | 8.2.x | Industry standard testing harness for unit, integration, and parameterized tests |
| Binary Compiler | PyInstaller | 6.6.x | Bundles Python script into a single zero-dependency standalone `.exe` for end users |

### Dependency Risk Assessment:
- **No high-risk dependencies identified.** All selected packages are open-source (BSD/MIT licensed), battle-tested, and actively maintained.
- **Excel Formatting Risk:** Large numeric values in Excel default to scientific notation (`1.25E+12`) if written as general text or raw strings. `openpyxl` explicit numeric cell assignment (`cell.number_format = '#,##0.00'`) mitigates this risk completely.

---

## TDD SECTION 3 — COMPONENT INTERFACE SPECIFICATIONS

### 3.1 Domain Data Contracts & Configuration (`schemas.py`)

```python
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import pandas as pd
from pydantic import BaseModel, Field

class ProcessingConfig(BaseModel):
    """Immutable runtime configuration passed to the orchestrator."""
    raw_data_path: Path = Field(..., description="Path to raw pipe-delimited file")
    reference_data_path: Path = Field(..., description="Path to reference mapping Excel/CSV file")
    output_report_path: Path = Field(..., description="Destination path for formatted Excel report")
    delimiter: str = Field(default="|", description="Delimiter used in raw file")
    lookup_key: str = Field(default="KODE_UKER", description="Join key column name")
    segmen_filter: Optional[str] = Field(default=None, description="Optional SEGMEN value to filter")
    number_format: str = Field(default="#,##0.00", description="Excel display format for IDR volume")

@dataclass(frozen=True)
class IngestionResult:
    data: pd.DataFrame
    total_rows: int
    malformed_rows_count: int

@dataclass(frozen=True)
class EnrichmentResult:
    data: pd.DataFrame
    matched_count: int
    unmapped_count: int
    unmapped_keys: list[str]

@dataclass(frozen=True)
class AggregationResult:
    summary_data: pd.DataFrame
    total_volume_idr: float
    branch_count: int

@dataclass(frozen=True)
class PipelineReport:
    success: bool
    output_path: Path
    execution_time_seconds: float
    total_records_processed: int
    unmapped_records_count: int
    error_message: Optional[str] = None
```

---

### 3.2 Ingestion Engine (`ingestion.py`)

```python
from pathlib import Path
import pandas as pd

class DataIngestionError(Exception):
    """Raised when raw file parsing encounters unrecoverable structural errors."""
    pass

class IngestionEngine:
    """Parses pipe-delimited raw financial data files into structured DataFrames."""

    def __init__(self, delimiter: str = "|"):
        self.delimiter = delimiter

    def read_raw_data(self, file_path: Path) -> IngestionResult:
        """Reads pipe-delimited text file with automatic encoding fallback.
        
        Args:
            file_path: Absolute path to the raw data file.
            
        Returns:
            IngestionResult containing clean DataFrame and metrics.
            
        Raises:
            DataIngestionError: If file is unreadable, empty, or missing required columns.
        """
        ...
```

---

### 3.3 Reference Enricher (`enrichment.py`)

```python
from pathlib import Path
import pandas as pd

class ReferenceEnrichmentError(Exception):
    """Raised when reference workbook loading fails."""
    pass

class ReferenceEnricher:
    """Enriches main financial dataset with reference table mapping data (VLOOKUP equivalent)."""

    def __init__(self, reference_path: Path, lookup_key: str = "KODE_UKER"):
        self.reference_path = reference_path
        self.lookup_key = lookup_key

    def enrich(self, raw_df: pd.DataFrame) -> EnrichmentResult:
        """Performs left outer join to attach MAIN_CODE and MAIN_BRANCH.
        
        Args:
            raw_df: Ingested raw DataFrame.
            
        Returns:
            EnrichmentResult with merged dataset and orphan mapping stats.
            
        Raises:
            ReferenceEnrichmentError: If reference file is missing or lookup key is absent.
        """
        ...
```

---

### 3.4 Aggregation Engine (`aggregation.py`)

```python
from typing import Optional
import pandas as pd

class AggregationEngine:
    """Summarizes enriched financial records by branch codes and calculates volume totals."""

    def __init__(self, segment_filter: Optional[str] = None):
        self.segment_filter = segment_filter

    def aggregate(
        self, 
        enriched_df: pd.DataFrame, 
        group_cols: list[str] = ["MAIN_CODE", "MAIN_BRANCH"],
        value_col: str = "VOLUME_IN_IDR"
    ) -> AggregationResult:
        """Filters data and performs group-by summation replicating Excel Pivot Table behavior.
        
        Args:
            enriched_df: Dataframe output from ReferenceEnricher.
            group_cols: List of grouping columns.
            value_col: Column name to sum.
            
        Returns:
            AggregationResult containing summary table and aggregate metrics.
        """
        ...
```

---

### 3.5 Excel Report Exporter (`exporter.py`)

```python
from pathlib import Path
import pandas as pd

class ExportError(Exception):
    """Raised when writing or styling the Excel output file fails."""
    pass

class ExcelReportExporter:
    """Formats and exports summary and row-level datasets to an openpyxl Excel workbook."""

    def __init__(self, number_format: str = "#,##0.00"):
        self.number_format = number_format

    def export(
        self, 
        summary_df: pd.DataFrame, 
        enriched_df: pd.DataFrame, 
        output_path: Path,
        segment_filter_applied: str | None = None
    ) -> Path:
        """Writes Excel workbook with formatted Summary_Report and Enriched_Data tabs.
        
        Args:
            summary_df: Summarized Pivot Table dataframe.
            enriched_df: Row-level enriched dataframe.
            output_path: Destination path for .xlsx file.
            segment_filter_applied: Filter label for report header block.
            
        Returns:
            Path to written workbook.
            
        Raises:
            ExportError: If destination is locked or openpyxl fails.
        """
        ...
```

---

### 3.6 Pipeline Orchestrator & CLI Interface (`cli.py`)

```python
from pathlib import Path
from typer import Typer

app = Typer(name="excel-workflow", help="Automated Financial Data Processing Pipeline")

class PipelineOrchestrator:
    """Coordinates execution across Ingestion, Enrichment, Aggregation, and Export layers."""

    @staticmethod
    def execute(config: ProcessingConfig) -> PipelineReport:
        """Executes full ETL flow synchronously with exception boundary handling."""
        ...

@app.command()
def process(
    raw_data: Path = Option(..., "--raw", "-r", help="Path to raw_data.txt"),
    reference: Path = Option(..., "--ref", "-f", help="Path to reference.xlsx"),
    output: Path = Option(Path("./Financial_Summary_Report.xlsx"), "--out", "-o", help="Output file path"),
    segment: str | None = Option(None, "--segment", "-s", help="Filter by SEGMEN (e.g. Wholesale, Corporate)"),
    interactive: bool = Option(False, "--gui", "-g", help="Launch GUI file picker mode")
):
    """CLI entry point for running the financial pipeline."""
    ...
```

---

## TDD SECTION 4 — IMPLEMENTATION PHASES & TEST SUITE STRATEGY

```
   ┌──────────────────────────────────────────────────────────────────────────┐
   │ Phase 1: Core ETL Engine & Data Parsing                                  │
   │ Deliverable: Unit-tested Ingestion, Enrichment & Aggregation modules     │
   └────────────────────────────────────┬─────────────────────────────────────┘
                                        │
                                        ▼
   ┌──────────────────────────────────────────────────────────────────────────┐
   │ Phase 2: OpenPyXL Formatting & Excel Exporter                            │
   │ Deliverable: Styled .xlsx export with proper numeric formatting          │
   └────────────────────────────────────┬─────────────────────────────────────┘
                                        │
                                        ▼
   ┌──────────────────────────────────────────────────────────────────────────┐
   │ Phase 3: CLI & Interactive Desktop File-Picker GUI                       │
   │ Deliverable: Executable CLI app & GUI dialog fallback                    │
   └────────────────────────────────────┬─────────────────────────────────────┘
                                        │
                                        ▼
   ┌──────────────────────────────────────────────────────────────────────────┐
   │ Phase 4: Production Hardening, Edge-Cases & PyInstaller Executable      │
   │ Deliverable: Standalone zero-dependency .exe distribution package        │
   └──────────────────────────────────────────────────────────────────────────┘
```

---

### Phase 1 — Core Data Engine & Foundation

- **Scope:** IngestionEngine, ReferenceEnricher, AggregationEngine, and domain schemas.
- **Deliverable:** Runnable Python module passing pure data transformation unit tests.
- **Unit Tests:**
  1. `test_ingest_reads_valid_pipe_delimited_file`: Verifies `read_raw_data` correctly parses 3 columns from raw text.
  2. `test_ingest_handles_utf8_and_cp1252_encodings`: Verifies automatic encoding detection fallback on non-UTF8 files.
  3. `test_ingest_strips_leading_trailing_whitespace`: Ensures string values in `KODE_UKER` are trimmed.
  4. `test_enrichment_matches_valid_kode_uker_keys`: Verifies VLOOKUP equivalency attaches `MAIN_CODE` and `MAIN_BRANCH`.
  5. `test_enrichment_flags_unmapped_kode_uker_keys`: Ensures missing keys are assigned `"UNMAPPED"` and returned in audit report.
  6. `test_aggregation_groups_by_code_and_sums_volume`: Verifies correct calculation of `sum(VOLUME_IN_IDR)`.
  7. `test_aggregation_applies_segment_filter`: Verifies rows are filtered by `SEGMEN` before grouping when filter is active.

---

### Phase 2 — Excel Formatting & Report Exporter

- **Scope:** OpenPyXL export logic, tabular layout styling, subtotal suppression, numeric formatting.
- **Deliverable:** Excel report exporter producing correctly styled `.xlsx` files without scientific notation (`E+12`).
- **Integration Tests:**
  1. `test_exporter_generates_valid_xlsx_workbook`: Verifies `.xlsx` file creation and readability by openpyxl.
  2. `test_exporter_applies_explicit_number_format`: Checks that volume cells have `.number_format == '#,##0.00'`.
  3. `test_exporter_auto_adjusts_column_widths`: Verifies column widths are scaled dynamically based on string length.
  4. `test_exporter_creates_summary_and_enriched_sheets`: Asserts presence of both `"Summary_Report"` and `"Enriched_Data"` worksheets.
  5. `test_exporter_adds_grand_total_row`: Verifies presence of styled Grand Total row at bottom of summary table.

---

### Phase 3 — User Interfaces (CLI & Easy Drag-and-Drop GUI)

- **Scope:** Typer CLI commands, interactive parameter prompts, Tkinter file picker modal.
- **Deliverable:** Fully usable command-line utility with GUI drag-and-drop file fallback.
- **Validation Criteria:**
  1. **CLI Execution:** Command `python -m src.cli process -r raw.txt -f ref.xlsx` executes cleanly and exits with code 0.
  2. **Zero-Args GUI Launch:** Running `python main.py` with no CLI flags launches native Windows file selection dialog.
  3. **Error Reporting:** Clear user-facing message shown if input files do not exist.

---

### Phase 4 — Polish, Packaging & Production Hardening

- **Scope:** PyInstaller build script, edge-case hardening, performance verification under 500k-row load.
- **Deliverable:** Single-file `ExcelWorkflowProcessor.exe` executable for end users.
- **Exit Criteria:**
  1. 100% test suite pass rate in pytest.
  2. Processing time < 3.0 seconds for a dataset of 100,000 raw rows.
  3. Executable runs on clean Windows 11 machine without Python pre-installed.

---

## TDD SECTION 5 — DESIGN VERIFICATION & EDGE CASES

### 5.1 Failure Mode Catalog

| Failure Mode | System Manifestation | Detection Method | Correct Response |
|---|---|---|---|
| **Missing Lookup Key in Reference File** | `KODE_UKER` is absent from reference workbook sheet | `ReferenceEnricher.enrich()` schema check | Raise `ReferenceEnrichmentError` with explicit message listing available columns. |
| **Scientific Notation Output (`1.2E+12`)** | Numbers rendered as string scientific exponent in Excel | `test_exporter_applies_explicit_number_format` | Force openpyxl cell data type to float and assign `cell.number_format = '#,##0.00'`. |
| **Output File Open in Microsoft Excel** | `PermissionError: [Errno 13] Permission denied` | Exporter file-write try/except catch | Catch `PermissionError` and prompt user to close Excel or specify alternate output filename. |
| **Malformed Raw Data Rows (Extra Pipes)** | Extra `\|` delimiters cause column shift | Row column count check during ingestion | Log malformed row index to audit list, skip row, and report count in final telemetry. |
| **Monetary Precision Loss** | Floating-point rounding errors in IDR sum | Pytest assertion against `decimal.Decimal` reference sum | Round final aggregations explicitly to 2 decimal places using `np.round()` or `round(val, 2)`. |

---

### 5.2 Performance Thresholds

| Metric | Target | Hard Limit | Measurement Method |
|---|---|---|---|
| **Raw File Ingestion (100k rows)** | < 0.8s | < 2.0s | Pytest benchmark timing `IngestionEngine.read_raw_data` |
| **Reference Join & Merge (100k rows)** | < 0.5s | < 1.5s | Pytest benchmark timing `ReferenceEnricher.enrich` |
| **OpenPyXL Excel Render & Save** | < 1.5s | < 4.0s | Pytest benchmark timing `ExcelReportExporter.export` |
| **Total End-to-End Pipeline Latency** | < 2.8s | < 6.0s | System execution telemetry timer |
| **Peak Memory Consumption** | < 150 MB | < 400 MB | `tracemalloc` memory audit during test run |

---

### 5.3 Platform & Environment Edge Cases

- **Windows Backslash Path Formatting:** Input paths provided with mixed slashes (e.g. `C:\data/raw.txt`) normalized using Python `pathlib.Path()`.
- **Large Integer / High IDR Volume Numbers:** Indonesian Rupiah values frequently exceed 10^12 (Trillions IDR). Standard 64-bit float (`float64`) provides up to 15-17 significant digits of precision, perfectly preserving precision up to 9,000,000,000,000,000.00 IDR.
- **Empty / Blank Raw Data Files:** Ingestion raises `DataIngestionError("Raw data file is empty")` before attempting memory allocation.

---

## Next Steps & Diagnostic Question

### Next Steps
1. **Scaffold Project Directory:** Create repository structure (`src/`, `tests/`, `config/`) and install pinned dependencies from Section 2 (`pandas~=2.2.0`, `openpyxl~=3.1.0`, `typer~=0.12.0`, `pytest~=8.2.0`).
2. **Implement Phase 1 Engine:** Build `schemas.py`, `ingestion.py`, `enrichment.py`, and `aggregation.py` along with unit test suite in `tests/test_core.py`.
3. **Implement Phase 2 Exporter:** Build `exporter.py` with custom OpenPyXL styling for `Summary_Report` and `Enriched_Data` worksheets.
4. **Implement Phase 3 UI & CLI:** Build `cli.py` and `gui.py` file-picker fallback wrapper, then assemble PyInstaller build configuration.

### Diagnostic Question
*Should unmapped `KODE_UKER` records (codes missing from the reference workbook) abort the pipeline execution immediately, or should they be grouped under an explicit `"UNMAPPED"` branch in the final report while emitting an audit warning?* (Recommended default: Group under `"UNMAPPED"` with an audit sheet warning so report generation is never blocked by minor mapping gaps.)

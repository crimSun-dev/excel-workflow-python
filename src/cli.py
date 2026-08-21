"""CLI Interface (TDD Section 3.6).

Typer-based command-line entry point for the financial pipeline. Running with
no subcommand/args (or with --gui) falls back to the Tkinter file-picker GUI.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer

from .error_guidance import explain_failure
from .orchestrator import PipelineOrchestrator
from .schemas import ProcessingConfig

app = typer.Typer(
    name="excel-workflow",
    help="Automated Financial Data Processing Pipeline",
    add_completion=False,
    no_args_is_help=False,
)


def _split_list(value: Optional[str]) -> Optional[list[str]]:
    """Parses a comma-separated exclusion flag into a list, or None if omitted.

    `None` (flag absent) means "use the workflow's default", while an empty
    string means "exclude nothing" - the same contract as the GUI fields.
    """
    if value is None:
        return None
    return [token.strip() for token in value.split(",") if token.strip()]


def _print_report(report) -> None:
    """Renders a PipelineReport to the terminal."""
    if report.success:
        typer.secho("\n[SUCCESS] Pipeline completed.", fg=typer.colors.GREEN, bold=True)
        typer.echo(f"  Output report      : {report.output_path}")
        typer.echo(f"  Records processed  : {report.total_records_processed:,}")
        typer.echo(f"  Unmapped records   : {report.unmapped_records_count:,}")
        typer.echo(f"  Execution time     : {report.execution_time_seconds}s")
    else:
        # Same reasoning the GUI popup shows, so a terminal run is not the one
        # place that still reports only the exception text.
        explanation = explain_failure(report.error_message)
        typer.secho(
            f"\n[FAILED] {explanation.title}", fg=typer.colors.RED, bold=True
        )
        typer.echo(f"  {explanation.what_happened}")
        typer.echo(f"  What to do: {explanation.what_to_do}")
        typer.echo(f"  Details: {report.error_message}")


@app.command()
def process(
    raw_data: Path = typer.Option(
        ...,
        "--raw",
        "-r",
        help="Path to raw delimited data (.txt/.csv/.bin); for report-giro this "
        "is the monthly giro source workbook (.xlsx)",
    ),
    workflow: str = typer.Option(
        "akumulasi",
        "--workflow",
        "-w",
        # Listed in the operator's daily sequence, matching the GUI dropdown
        # and the valid-options listing below (both driven by WorkflowId order).
        help="Workflow: akumulasi | timeseries-active-user-qlola | "
        "report-data-statis | rincian-vol-tf | rincian-portal-bg | "
        "timeseries-fbi-briva | report-giro",
    ),
    reference: Optional[Path] = typer.Option(
        None,
        "--ref",
        "-f",
        help="Path to reference.xlsx (required for akumulasi, timeseries-fbi-briva, "
        "timeseries-active-user-qlola, and report-data-statis)",
    ),
    master: Optional[Path] = typer.Option(
        None,
        "--master",
        "-m",
        help="Path to the master workbook: ID -> MAIN_CODE for "
        "timeseries-active-user-qlola, or the giro master to update for "
        "report-giro (required for both)",
    ),
    output: Path = typer.Option(
        Path("./Financial_Summary_Report.xlsx"),
        "--out",
        "-o",
        help="Output file path",
    ),
    segment: Optional[str] = typer.Option(
        None,
        "--segment",
        "-s",
        help="Comma-separated SEGMEN keep-list (e.g. NONWHOLESALE). Omit to use "
        "the workflow default; pass an empty string to keep every segment",
    ),
    exclude_segmen: Optional[str] = typer.Option(
        None,
        "--exclude-segmen",
        help="Comma-separated SEGMEN values to drop (e.g. 'KORPORASI,Wholesale'). "
        "Omit to use the workflow default; pass an empty string to drop none",
    ),
    exclude_source: Optional[str] = typer.Option(
        None,
        "--exclude-source",
        help="Comma-separated SOURCE values to drop (e.g. 'CMS'). Omit to use "
        "the workflow default; pass an empty string to drop none",
    ),
    source: Optional[str] = typer.Option(
        None,
        "--source",
        help="Comma-separated SOURCE keep-list (keep only these). Omit for no "
        "inclusion filter, leaving the workflow's SOURCE default in charge",
    ),
    kw: Optional[str] = typer.Option(
        None,
        "--kw",
        help="Comma-separated KW keep-list (keep only these). Omit to use the "
        "workflow default (Report Data Statis keeps 'KANWIL MALANG'); pass an "
        "empty string to keep every KW. A no-op when the source has no "
        "KW/PRODUCT/GROUP_PRODUCT/KAWIL column",
    ),
    delimiter: str = typer.Option("|", "--delimiter", "-d", help="Raw file delimiter"),
    interactive: bool = typer.Option(
        False, "--gui", "-g", help="Launch GUI file picker mode"
    ),
) -> None:
    """CLI entry point for running the financial pipeline."""
    if interactive:
        from .gui import launch_gui

        launch_gui()
        raise typer.Exit(code=0)

    from .workflows.base import WorkflowId

    try:
        workflow_id = WorkflowId(workflow).value
    except ValueError:
        valid = ", ".join(w.value for w in WorkflowId)
        typer.secho(
            f"\n[FAILED] Unknown workflow '{workflow}'. Valid options: {valid}",
            fg=typer.colors.RED,
            bold=True,
        )
        raise typer.Exit(code=1)

    from .workflows.registry import get_definition

    definition = get_definition(workflow_id)
    if definition.requires_reference and reference is None:
        typer.secho(
            f"\n[FAILED] The '{workflow_id}' workflow requires --ref/-f "
            "(a reference mapping file).",
            fg=typer.colors.RED,
            bold=True,
        )
        raise typer.Exit(code=1)
    if definition.requires_master_data and master is None:
        # What "master" means differs per workflow (Qlola's ID -> MAIN_CODE
        # table vs the Giro master workbook), so the definition supplies the
        # wording rather than the flag hardcoding one workflow's meaning.
        typer.secho(
            f"\n[FAILED] The '{workflow_id}' workflow requires --master/-m: "
            f"{definition.master_picker_title}.",
            fg=typer.colors.RED,
            bold=True,
        )
        raise typer.Exit(code=1)

    config = ProcessingConfig(
        raw_data_path=raw_data,
        reference_data_path=reference,
        master_data_path=master,
        workflow_id=workflow_id,
        output_report_path=output,
        # Split like every other filter flag, so `--segment "A,B"` is a two-value
        # keep-list rather than a lookup for a segment literally named "A,B".
        segmen_filter=_split_list(segment),
        # An omitted flag stays None so the workflow default applies; a supplied
        # (even empty) value is the caller explicitly choosing what to drop.
        segmen_exclude=_split_list(exclude_segmen),
        source_exclude=_split_list(exclude_source),
        source_include=_split_list(source),
        kw_include=_split_list(kw),
        delimiter=delimiter,
    )
    report = PipelineOrchestrator.execute(config)
    _print_report(report)
    raise typer.Exit(code=0 if report.success else 1)


@app.command()
def gui() -> None:
    """Launch the interactive drag-and-drop file-picker GUI."""
    from .gui import launch_gui

    launch_gui()


def main() -> None:
    """Module entry point. With no args, launches the GUI fallback."""
    import sys

    if len(sys.argv) == 1:
        from .gui import launch_gui

        launch_gui()
        return
    app()


if __name__ == "__main__":
    main()

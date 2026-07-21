"""CLI Interface (TDD Section 3.6).

Typer-based command-line entry point for the financial pipeline. Running with
no subcommand/args (or with --gui) falls back to the Tkinter file-picker GUI.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer

from .orchestrator import PipelineOrchestrator
from .schemas import ProcessingConfig

app = typer.Typer(
    name="excel-workflow",
    help="Automated Financial Data Processing Pipeline",
    add_completion=False,
    no_args_is_help=False,
)


def _print_report(report) -> None:
    """Renders a PipelineReport to the terminal."""
    if report.success:
        typer.secho("\n[SUCCESS] Pipeline completed.", fg=typer.colors.GREEN, bold=True)
        typer.echo(f"  Output report      : {report.output_path}")
        typer.echo(f"  Records processed  : {report.total_records_processed:,}")
        typer.echo(f"  Unmapped records   : {report.unmapped_records_count:,}")
        typer.echo(f"  Execution time     : {report.execution_time_seconds}s")
    else:
        typer.secho("\n[FAILED] Pipeline error.", fg=typer.colors.RED, bold=True)
        typer.echo(f"  {report.error_message}")


@app.command()
def process(
    raw_data: Path = typer.Option(..., "--raw", "-r", help="Path to raw_data.txt"),
    reference: Path = typer.Option(..., "--ref", "-f", help="Path to reference.xlsx"),
    output: Path = typer.Option(
        Path("./Financial_Summary_Report.xlsx"),
        "--out",
        "-o",
        help="Output file path",
    ),
    segment: Optional[str] = typer.Option(
        None, "--segment", "-s", help="Filter by SEGMEN (e.g. Wholesale, Corporate)"
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

    config = ProcessingConfig(
        raw_data_path=raw_data,
        reference_data_path=reference,
        output_report_path=output,
        segmen_filter=segment,
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

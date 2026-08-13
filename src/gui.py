"""Tkinter GUI Fallback (TDD Section 3, Phase 3).

Native OS file-picker dialog launched when the user runs the app with no CLI
arguments. Lets a non-technical user pick a workflow, select the raw file,
(for Akumulasi) a reference workbook, the SEGMEN keep-only / SEGMEN exclude /
SOURCE exclude filters, and an output path, then runs the pipeline. The
reference-file row is shown only for the workflows that need it.

Each filter field is seeded with the selected workflow's default (e.g. KORPORASI
for Report Data Statis) and stays editable, so the operator can see what the
report drops by default and still decide per run what to include or exclude.
Fields are greyed out only for workflows where the filter cannot apply
(`supports_segment_filter` / `has_source_filter`).

Input button labels, dialog titles, and extension filters come from the
workflow definition, so Report Giro offers two Excel workbook pickers rather
than the pipe-delimited text labels the other workflows use.

The small pure helpers below (`source_filter_enabled`, `segment_filter_enabled`,
`default_source_text`, `default_segmen_include_text`,
`default_segmen_exclude_text`, `parse_exclude_list`, `raw_picker_config`,
`master_picker_config`) hold the fields' behavior so they are testable without a
display server; `launch_gui` only wires them to widgets.
"""

from __future__ import annotations

from pathlib import Path

from .branding import PRODUCT_NAME
from .orchestrator import PipelineOrchestrator
from .schemas import PipelineReport, ProcessingConfig
from .workflows.base import WorkflowId
from .workflows.registry import get_definition


def workflow_choices() -> dict[str, str]:
    """{label: workflow id} for the dropdown, in operator sequence.

    `WorkflowId`'s declaration order *is* the display contract - Python enums
    iterate in declaration order, so reordering the enum reorders the dropdown
    with no parallel ordering list to drift out of sync.
    """
    return {get_definition(w).label: w.value for w in WorkflowId}


def source_filter_enabled(workflow_id: WorkflowId | str) -> bool:
    """Whether the SOURCE field accepts input for this workflow."""
    return get_definition(workflow_id).has_source_filter


def segment_filter_enabled(workflow_id: WorkflowId | str) -> bool:
    """Whether the SEGMEN fields accept input for this workflow.

    Live wherever the workflow can honor them, since which segments to include
    or exclude is a per-run operator decision; Report Giro applies them to its
    monthly source like the rest. Greyed out only where the value would be
    ignored, so a disabled field always means exactly that.
    """
    return get_definition(workflow_id).supports_segment_filter


def raw_picker_config(
    workflow_id: WorkflowId | str,
) -> tuple[str, str, list[tuple[str, str]]]:
    """(button label, dialog title, filetypes) for the raw/source input.

    Report Giro reads a monthly Excel workbook rather than a pipe-delimited
    extract, so these come from the workflow definition instead of being fixed.
    """
    definition = get_definition(workflow_id)
    return (
        definition.raw_button_label,
        definition.raw_picker_title,
        [tuple(entry) for entry in definition.raw_filetypes],
    )


def master_picker_config(workflow_id: WorkflowId | str) -> tuple[str, str]:
    """(button label, dialog title) for the master-data input."""
    definition = get_definition(workflow_id)
    return definition.master_button_label, definition.master_picker_title


def default_source_text(workflow_id: WorkflowId | str) -> str:
    """The workflow's configured SOURCE exclusion list as GUI text (e.g. 'CMS')."""
    definition = get_definition(workflow_id)
    if not definition.has_source_filter:
        return ""
    return ", ".join(definition.source_exclude)


def default_segmen_include_text(workflow_id: WorkflowId | str) -> str:
    """The workflow's default SEGMEN "keep only" value (e.g. 'NONWHOLESALE')."""
    definition = get_definition(workflow_id)
    if not definition.supports_segment_filter:
        return ""
    return definition.segmen_include or ""


def default_segmen_exclude_text(workflow_id: WorkflowId | str) -> str:
    """The workflow's default SEGMEN exclusion list as GUI text (e.g. 'KORPORASI').

    Seeding the field with the default is what makes the automatic behavior
    visible: the operator can see KORPORASI is being dropped and edit or clear it
    instead of having to accept it.
    """
    definition = get_definition(workflow_id)
    if not definition.supports_segment_filter:
        return ""
    return ", ".join(definition.exclude_segmen)


def parse_exclude_list(text: str) -> list[str]:
    """Splits a comma-separated exclusion field into a trimmed list.

    An empty/whitespace-only field yields `[]`, which explicitly means "exclude
    nothing" - the operator clearing the field is a deliberate instruction, not
    a request to fall back to the workflow default.
    """
    return [token.strip() for token in text.split(",") if token.strip()]


def unmapped_warning_text(report: PipelineReport) -> str | None:
    """Operator-facing text for a near-total UNMAPPED join, or None if healthy.

    Returned only when the pipeline itself flagged the join as effectively
    failed. The report was still written, so this reads as a "check your inputs"
    warning rather than an error.
    """
    if not report.success or not report.unmapped_diagnostic:
        return None
    return (
        "The report was still written, but almost every ID failed to match the "
        "master lookup, so the Summary will be nearly empty.\n\n"
        f"{report.unmapped_diagnostic}\n\n"
        "Check that the correct master-data file was selected and that its ID "
        "column holds the same values as the raw file."
    )


def launch_gui() -> None:
    """Opens the Tkinter window. Imported lazily so headless CLI use is unaffected."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(PRODUCT_NAME)
    root.geometry("640x600")
    root.resizable(False, False)

    state: dict[str, Path | None] = {
        "raw": None,
        "ref": None,
        "master": None,
        "out": None,
    }

    # label -> workflow id, preserving the enum declaration order.
    choices = workflow_choices()
    labels = list(choices)
    default_label = get_definition(WorkflowId.AKUMULASI).label

    workflow_var = tk.StringVar(value=default_label)
    raw_var = tk.StringVar(value="No file selected")
    ref_var = tk.StringVar(value="No file selected")
    master_var = tk.StringVar(value="No file selected")
    out_var = tk.StringVar(value="./Financial_Summary_Report.xlsx")
    seg_var = tk.StringVar(value=default_segmen_include_text(WorkflowId.AKUMULASI))
    seg_exclude_var = tk.StringVar(
        value=default_segmen_exclude_text(WorkflowId.AKUMULASI)
    )
    source_var = tk.StringVar(value=default_source_text(WorkflowId.AKUMULASI))

    def selected_workflow_id() -> str:
        return choices[workflow_var.get()]

    def pick_raw() -> None:
        # Label, title, and extension filter follow the selected workflow, so
        # Report Giro offers an Excel picker instead of a text/CSV one.
        _, title, filetypes = raw_picker_config(selected_workflow_id())
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            state["raw"] = Path(path)
            raw_var.set(path)

    def pick_ref() -> None:
        path = filedialog.askopenfilename(
            title="Select reference mapping workbook",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if path:
            state["ref"] = Path(path)
            ref_var.set(path)

    def pick_master() -> None:
        _, title = master_picker_config(selected_workflow_id())
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if path:
            state["master"] = Path(path)
            master_var.set(path)

    def pick_out() -> None:
        path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            state["out"] = Path(path)
            out_var.set(path)

    def run() -> None:
        workflow_id = selected_workflow_id()
        definition = get_definition(workflow_id)
        needs_ref = definition.requires_reference
        needs_master = definition.requires_master_data
        if not state["raw"]:
            messagebox.showerror(
                "Missing input", f"{raw_picker_config(workflow_id)[1]}."
            )
            return
        if needs_ref and not state["ref"]:
            messagebox.showerror(
                "Missing input",
                "The selected workflow requires a reference file.",
            )
            return
        if needs_master and not state["master"]:
            # The master input means different things per workflow (Qlola's
            # ID -> MAIN_CODE table vs the Giro master workbook), so the
            # definition's own picker title carries the message.
            messagebox.showerror(
                "Missing input", f"{master_picker_config(workflow_id)[1]}."
            )
            return
        output = state["out"] or Path(out_var.get())
        # A disabled field must never influence the run, so its value is only read
        # for workflows that expose it. `None` means "use the workflow default";
        # a cleared but enabled field is passed through as empty, which is the
        # operator explicitly asking for no filtering on that dimension.
        segment = (
            seg_var.get().strip() if definition.supports_segment_filter else None
        )
        segmen_exclude = (
            parse_exclude_list(seg_exclude_var.get())
            if definition.supports_segment_filter
            else None
        )
        source_exclude = (
            parse_exclude_list(source_var.get())
            if definition.has_source_filter
            else None
        )
        config = ProcessingConfig(
            raw_data_path=state["raw"],
            reference_data_path=state["ref"] if needs_ref else None,
            master_data_path=state["master"] if needs_master else None,
            workflow_id=workflow_id,
            output_report_path=output,
            segmen_filter=segment,
            segmen_exclude=segmen_exclude,
            source_exclude=source_exclude,
        )
        report = PipelineOrchestrator.execute(config)
        if report.success:
            messagebox.showinfo(
                "Success",
                f"Report generated:\n{report.output_path}\n\n"
                f"Records processed: {report.total_records_processed:,}\n"
                f"Unmapped records: {report.unmapped_records_count:,}\n"
                f"Time: {report.execution_time_seconds}s",
            )
            # Shown after the success box so the operator cannot miss a join
            # that silently produced an almost-empty Summary.
            warning = unmapped_warning_text(report)
            if warning:
                messagebox.showwarning("Master ID lookup failed", warning)
        else:
            messagebox.showerror("Pipeline failed", report.error_message or "Unknown error")

    pad = {"padx": 10, "pady": 6}

    # --- Product header ---
    header = ttk.Frame(root)
    header.grid(row=0, column=0, columnspan=3, **pad)
    ttk.Label(header, text=PRODUCT_NAME, font=("Segoe UI", 16, "bold")).pack(
        anchor="w"
    )

    # --- Workflow selector ---
    ttk.Label(root, text="Workflow:").grid(row=1, column=0, sticky="e", **pad)
    workflow_combo = ttk.Combobox(
        root, textvariable=workflow_var, values=labels, state="readonly", width=42
    )
    workflow_combo.grid(row=1, column=1, columnspan=2, sticky="w", **pad)

    # --- Raw data row (label follows the workflow) ---
    raw_button = ttk.Button(root, text="1. Raw Data File...", command=pick_raw, width=22)
    raw_button.grid(row=2, column=0, **pad)
    ttk.Label(root, textvariable=raw_var, width=45, anchor="w").grid(
        row=2, column=1, columnspan=2, **pad
    )

    # --- Reference row (toggled by workflow) ---
    ref_button = ttk.Button(root, text="2. Reference File...", command=pick_ref, width=22)
    ref_button.grid(row=3, column=0, **pad)
    ref_label = ttk.Label(root, textvariable=ref_var, width=45, anchor="w")
    ref_label.grid(row=3, column=1, columnspan=2, **pad)

    # --- Master-data row (toggled by workflow; Qlola only) ---
    master_button = ttk.Button(
        root, text="2b. Master Data File...", command=pick_master, width=22
    )
    master_button.grid(row=4, column=0, **pad)
    master_label = ttk.Label(root, textvariable=master_var, width=45, anchor="w")
    master_label.grid(row=4, column=1, columnspan=2, **pad)

    # --- Output row ---
    ttk.Button(root, text="3. Output Path...", command=pick_out, width=22).grid(
        row=5, column=0, **pad
    )
    ttk.Label(root, textvariable=out_var, width=45, anchor="w").grid(
        row=5, column=1, columnspan=2, **pad
    )

    segment_field_label = ttk.Label(root, text="SEGMEN keep only:")
    segment_field_label.grid(row=6, column=0, **pad)
    segment_entry = ttk.Entry(root, textvariable=seg_var, width=30)
    segment_entry.grid(row=6, column=1, columnspan=2, sticky="w", **pad)

    segment_exclude_label = ttk.Label(root, text="SEGMEN exclude:")
    segment_exclude_label.grid(row=7, column=0, **pad)
    segment_exclude_entry = ttk.Entry(root, textvariable=seg_exclude_var, width=30)
    segment_exclude_entry.grid(row=7, column=1, columnspan=2, sticky="w", **pad)

    source_field_label = ttk.Label(root, text="SOURCE exclude:")
    source_field_label.grid(row=8, column=0, **pad)
    source_entry = ttk.Entry(root, textvariable=source_var, width=30)
    source_entry.grid(row=8, column=1, columnspan=2, sticky="w", **pad)

    ttk.Button(root, text="Run Pipeline", command=run, width=22).grid(
        row=9, column=0, columnspan=3, **pad
    )

    def toggle_optional_rows(*_args: object) -> None:
        """Show reference / master rows and the SOURCE field per workflow needs."""
        workflow_id = selected_workflow_id()
        definition = get_definition(workflow_id)
        # Input labels follow the workflow: Report Giro picks two Excel
        # workbooks, so "Raw Data File" would misdescribe both of them.
        raw_button.configure(text=raw_picker_config(workflow_id)[0])
        master_button.configure(text=master_picker_config(workflow_id)[0])
        # Filter fields are re-seeded with the selected workflow's defaults, so
        # switching workflows shows what that report filters out by default and
        # can never carry another workflow's values over.
        source_var.set(default_source_text(workflow_id))
        seg_var.set(default_segmen_include_text(workflow_id))
        seg_exclude_var.set(default_segmen_exclude_text(workflow_id))
        enabled = source_filter_enabled(workflow_id)
        source_entry.configure(state="normal" if enabled else "disabled")
        source_field_label.configure(state="normal" if enabled else "disabled")
        # Greyed out only where SEGMEN cannot apply, so a disabled field always
        # means the value would be ignored rather than silently discarded.
        segment_enabled = segment_filter_enabled(workflow_id)
        state_name = "normal" if segment_enabled else "disabled"
        segment_entry.configure(state=state_name)
        segment_field_label.configure(state=state_name)
        segment_exclude_entry.configure(state=state_name)
        segment_exclude_label.configure(state=state_name)
        if definition.requires_reference:
            ref_button.grid()
            ref_label.grid()
        else:
            ref_button.grid_remove()
            ref_label.grid_remove()
        if definition.requires_master_data:
            master_button.grid()
            master_label.grid()
        else:
            master_button.grid_remove()
            master_label.grid_remove()

    workflow_combo.bind("<<ComboboxSelected>>", toggle_optional_rows)
    toggle_optional_rows()  # apply initial visibility for the default workflow

    root.mainloop()


if __name__ == "__main__":
    launch_gui()

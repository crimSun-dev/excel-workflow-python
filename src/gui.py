"""Tkinter GUI Fallback (TDD Section 3, Phase 3).

Native OS file-picker dialog launched when the user runs the app with no CLI
arguments. Lets a non-technical user pick a workflow, select the raw file,
(for Akumulasi) a reference workbook, the FILTERS (SEGMENT / SOURCE / KW), and
an output path, then runs the pipeline. The reference-file row is shown only for
the workflows that need it.

Filter semantics, one rule for all three boxes:

* **Empty box = automatic.** The report runs its own baked rules (Report Data
  Statis drops KORPORASI, Qlola drops CMS, Briva keeps NONWHOLESALE). The boxes
  are therefore never pre-filled - a filled box would read as something the
  operator has to manage, and clearing it would read as "remove all filters".
  What each report does automatically is shown as a hint beside the box.
* **Typed or picked values = keep only those.** A comma-separated keep-list that
  replaces that dimension's automatic rule for the run. Every box has a
  `Choose...` ticklist of the values that report's extract actually carries, and
  prints the same list beneath itself: the values are meant to be *read off the
  window*, not recalled from the manual workflow, which is where mistyped
  keywords came from. Ticking nothing (or clearing the box) is still automatic.

Fields are greyed out only for workflows where the filter cannot apply
(`supports_segment_filter` / `has_source_filter` / `has_kw_filter`), so a
disabled box always means exactly "this value would be ignored". The `Choose...`
button additionally greys out where the legal values were never confirmed - the
box still accepts typed input there.

Failures are read through `error_guidance`, so the popup leads with what
happened and what to do and keeps the technical message underneath.

Input button labels, dialog titles, and extension filters come from the
workflow definition, so Report Giro offers two Excel workbook pickers rather
than the pipe-delimited text labels the other workflows use.

The small pure helpers below (`source_filter_enabled`, `segment_filter_enabled`,
`kw_filter_enabled`, `default_source_text`, `default_segmen_include_text`,
`default_segmen_exclude_text`, `default_kw_text`, `filter_hint_text`,
`filter_options`, `filter_options_note`, `join_filter_values`,
`parse_filter_input`, `raw_picker_config`, `master_picker_config`,
`pipeline_failure_dialog`, `format_run_time`) hold the fields' behavior so they
are testable without a display server; `launch_gui` only wires them to widgets.
"""

from __future__ import annotations

import threading
from pathlib import Path

from .branding import PRODUCT_NAME, contact_credit_text
from .error_guidance import failure_dialog_text
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
    or exclude is a per-run operator decision. Greyed out only where the
    stakeholder skip-list (or a missing vocabulary) means the value would be
    ignored, so a disabled field always means exactly that.
    """
    return get_definition(workflow_id).supports_segment_filter


def kw_filter_enabled(workflow_id: WorkflowId | str) -> bool:
    """Whether the KW field accepts input for this workflow."""
    return get_definition(workflow_id).has_kw_filter


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
    """The workflow's default SEGMEN exclusion list as GUI text (e.g. 'KORPORASI')."""
    definition = get_definition(workflow_id)
    if not definition.supports_segment_filter:
        return ""
    return ", ".join(definition.exclude_segmen)


def default_kw_text(workflow_id: WorkflowId | str) -> str:
    """The workflow's default KW keep-list as GUI text (Data Statis: KANWIL MALANG)."""
    definition = get_definition(workflow_id)
    if not definition.has_kw_filter:
        return ""
    return ", ".join(definition.kw_include)


def filter_hint_text(workflow_id: WorkflowId | str, dimension: str) -> str:
    """What this report does with a dimension when its box is left empty.

    The boxes ship empty so that empty reads as "automatic", which would
    otherwise hide the rule the report applies for free. This hint is what keeps
    that rule visible: "auto: drops KORPORASI" beside an empty SEGMENT box says
    the same thing the old pre-filled field did, without inviting the operator
    to manage it.
    """
    dimension = dimension.upper()
    if dimension == "SEGMENT":
        if not segment_filter_enabled(workflow_id):
            return "not used by this report"
        keeps = default_segmen_include_text(workflow_id)
        drops = default_segmen_exclude_text(workflow_id)
    elif dimension == "SOURCE":
        if not source_filter_enabled(workflow_id):
            return "not used by this report"
        keeps = ""
        drops = default_source_text(workflow_id)
    elif dimension == "KW":
        if not kw_filter_enabled(workflow_id):
            return "not used by this report"
        keeps = default_kw_text(workflow_id)
        drops = ""
    else:
        raise ValueError(f"Unknown filter dimension: {dimension!r}")

    parts = []
    if keeps:
        parts.append(f"keeps {keeps}")
    if drops:
        parts.append(f"drops {drops}")
    return f"auto: {', '.join(parts)}" if parts else "auto: keeps everything"


def filter_options(workflow_id: WorkflowId | str, dimension: str) -> tuple[str, ...]:
    """The values this report offers for a dimension, or `()` when there are none.

    `()` covers both "the report ignores this dimension" and "the legal values
    for it were never confirmed" - in either case there is nothing to pick from,
    so the box stays free-text and the picker button is greyed out.
    """
    dimension = dimension.upper()
    definition = get_definition(workflow_id)
    if dimension == "SEGMENT":
        return definition.segment_options if definition.supports_segment_filter else ()
    if dimension == "SOURCE":
        return definition.source_options if definition.has_source_filter else ()
    if dimension == "KW":
        return definition.kw_options if definition.has_kw_filter else ()
    raise ValueError(f"Unknown filter dimension: {dimension!r}")


def filter_options_note(workflow_id: WorkflowId | str, dimension: str) -> str:
    """The line under a filter box listing what the operator can choose from.

    The stakeholders' ask was that the keywords be *visible* rather than
    remembered from the manual workflow, so the values are printed beside the
    control instead of only living inside the picker popup. A dimension the
    report ignores says nothing here (its hint already says so), and one whose
    values were never confirmed says that plainly rather than implying the box
    is unusable.
    """
    dimension = dimension.upper()
    if not _dimension_enabled(workflow_id, dimension):
        return ""
    options = filter_options(workflow_id, dimension)
    if not options:
        return f"Choices: any {dimension} value in your file"
    return f"Choices ({len(options)}): {', '.join(options)}"


def _dimension_enabled(workflow_id: WorkflowId | str, dimension: str) -> bool:
    """Whether the named dimension's control accepts input for this workflow."""
    dimension = dimension.upper()
    if dimension == "SEGMENT":
        return segment_filter_enabled(workflow_id)
    if dimension == "SOURCE":
        return source_filter_enabled(workflow_id)
    if dimension == "KW":
        return kw_filter_enabled(workflow_id)
    raise ValueError(f"Unknown filter dimension: {dimension!r}")


def join_filter_values(values: list[str] | tuple[str, ...]) -> str:
    """Renders picked values as the text a filter box holds (`''` = automatic)."""
    return ", ".join(values)


def parse_filter_input(text: str) -> list[str] | None:
    """Turns one filter box into a keep-list, or `None` for "run the defaults".

    An empty/whitespace-only box yields `None`, which the workflow reads as
    "nothing supplied" and answers with its own baked rules. Any typed value
    yields the trimmed, comma-separated keep-list that replaces them for the run.
    """
    tokens = [token.strip() for token in text.split(",") if token.strip()]
    return tokens or None


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


def pipeline_failure_dialog(error_message: str | None) -> tuple[str, str]:
    """(title, body) for the failure popup: what happened, what to do, then detail.

    Every failure - not just the empty-extract case that used to be the one
    special-cased notice - is read through `error_guidance` first, so an
    operator who has never seen an exception name still learns what to try. The
    raw message is kept at the bottom for support.
    """
    return failure_dialog_text(error_message)


def format_run_time(report: PipelineReport) -> str:
    """Wall-clock time plus named stage durations for the success dialog."""
    lines = [f"Time: {report.execution_time_seconds}s"]
    timings = report.stage_timings or {}
    if timings:
        parts = [f"{name} {seconds:.1f}s" for name, seconds in timings.items()]
        lines.append("Stages: " + " · ".join(parts))
    return "\n".join(lines)


def launch_gui() -> None:
    """Opens the Tkinter window. Imported lazily so headless CLI use is unaffected."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(PRODUCT_NAME)
    # Taller and wider than the pre-picker window: each filter row now also
    # prints the values it accepts, and Report Data Statis' KW list is long.
    root.geometry("780x820")
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
    # Filter boxes start (and stay) empty: empty means "run this report's
    # automatic rules", which the hint labels spell out.
    seg_var = tk.StringVar(value="")
    source_var = tk.StringVar(value="")
    kw_var = tk.StringVar(value="")
    seg_hint_var = tk.StringVar(value="")
    source_hint_var = tk.StringVar(value="")
    kw_hint_var = tk.StringVar(value="")
    # The visible list of values each box accepts, refreshed per workflow.
    seg_note_var = tk.StringVar(value="")
    source_note_var = tk.StringVar(value="")
    kw_note_var = tk.StringVar(value="")
    status_var = tk.StringVar(value="Ready")

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
        # A disabled field must never influence the run, so its value is only
        # read for workflows that expose it. An empty box gives `None`, which the
        # workflow answers with its baked defaults; a typed value becomes that
        # dimension's keep-list and replaces the default policy for this run -
        # hence the paired empty exclusion lists below.
        segment = (
            parse_filter_input(seg_var.get())
            if definition.supports_segment_filter
            else None
        )
        source = (
            parse_filter_input(source_var.get())
            if definition.has_source_filter
            else None
        )
        kw = parse_filter_input(kw_var.get()) if definition.has_kw_filter else None
        config = ProcessingConfig(
            raw_data_path=state["raw"],
            reference_data_path=state["ref"] if needs_ref else None,
            master_data_path=state["master"] if needs_master else None,
            workflow_id=workflow_id,
            output_report_path=output,
            segmen_filter=segment,
            segmen_exclude=None if segment is None else [],
            source_include=source,
            source_exclude=None if source is None else [],
            kw_include=kw,
            progress_callback=lambda message: root.after(
                0, lambda m=message: status_var.set(m)
            ),
        )
        run_button.configure(state="disabled")
        status_var.set("Starting...")

        def worker() -> None:
            try:
                report = PipelineOrchestrator.execute(config)
            except Exception as exc:  # noqa: BLE001 - GUI must never freeze on a crash
                root.after(0, lambda: finish_with_crash(str(exc)))
                return
            root.after(0, lambda: finish_with_report(report))

        def finish_with_crash(message: str) -> None:
            run_button.configure(state="normal")
            status_var.set("Ready")
            title, body = pipeline_failure_dialog(message)
            messagebox.showerror(title, body)

        def finish_with_report(report: PipelineReport) -> None:
            run_button.configure(state="normal")
            status_var.set("Ready")
            if report.success:
                success_body = (
                    f"Report generated:\n{report.output_path}\n\n"
                    f"Records processed: {report.total_records_processed:,}\n"
                    f"Unmapped records: {report.unmapped_records_count:,}\n"
                    f"{format_run_time(report)}"
                )
                if report.operator_note:
                    success_body += f"\n\n{report.operator_note}"
                messagebox.showinfo("Success", success_body)
                # Shown after the success box so the operator cannot miss a join
                # that silently produced an almost-empty Summary.
                warning = unmapped_warning_text(report)
                if warning:
                    messagebox.showwarning("Master ID lookup failed", warning)
            else:
                title, body = pipeline_failure_dialog(report.error_message)
                messagebox.showerror(title, body)

        threading.Thread(target=worker, daemon=True).start()

    pad = {"padx": 10, "pady": 6}

    # --- Product header (name + contact credentials) ---
    header = ttk.Frame(root)
    header.grid(row=0, column=0, columnspan=3, sticky="w", **pad)
    ttk.Label(header, text=PRODUCT_NAME, font=("Segoe UI", 16, "bold")).pack(
        anchor="w"
    )
    ttk.Label(
        header,
        text=contact_credit_text(),
        font=("Segoe UI", 9),
        foreground="#555555",
        justify="left",
    ).pack(anchor="w", pady=(4, 0))

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

    def choose_values(dimension: str, var: tk.StringVar) -> None:
        """Ticklist popup for one dimension; ticking nothing means automatic.

        The options come from the *currently selected* workflow rather than
        being captured when the button was built, so switching reports switches
        the choices with no rebinding.
        """
        workflow_id = selected_workflow_id()
        options = filter_options(workflow_id, dimension)
        if not options:
            return
        popup = tk.Toplevel(root)
        popup.title(f"{dimension} filter")
        popup.transient(root)
        popup.resizable(False, False)
        ttk.Label(
            popup,
            text=f"Tick the {dimension} values to keep. Tick none to run this "
            f"report's automatic rule ({filter_hint_text(workflow_id, dimension)}).",
            wraplength=320,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(12, 6))
        picked = {value.casefold() for value in parse_filter_input(var.get()) or []}
        # Report Data Statis offers 18 KW values; one column of those would run
        # off a laptop screen, so long lists wrap into a second column.
        rows_per_column = 9
        ticks = ttk.Frame(popup)
        ticks.pack(anchor="w", padx=18)
        flags: list[tuple[str, tk.BooleanVar]] = []
        for index, option in enumerate(options):
            flag = tk.BooleanVar(value=option.casefold() in picked)
            ttk.Checkbutton(ticks, text=option, variable=flag).grid(
                row=index % rows_per_column,
                column=index // rows_per_column,
                sticky="w",
                padx=(0, 12),
                pady=1,
            )
            flags.append((option, flag))

        def confirm() -> None:
            var.set(join_filter_values([o for o, flag in flags if flag.get()]))
            popup.destroy()

        def clear() -> None:
            var.set("")
            popup.destroy()

        buttons = ttk.Frame(popup)
        buttons.pack(anchor="e", padx=12, pady=12)
        ttk.Button(buttons, text="Use automatic", command=clear, width=14).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(buttons, text="OK", command=confirm, width=10).pack(side="left")
        popup.grab_set()

    # --- FILTERS block: one box per dimension, all three on every report ---
    filters_frame = ttk.LabelFrame(root, text="FILTERS")
    filters_frame.grid(row=6, column=0, columnspan=3, sticky="we", **pad)
    ttk.Label(
        filters_frame,
        text="Leave a box empty to run this report's automatic rules. Choose "
        "values (or type them, comma-separated) to keep only those.",
        wraplength=680,
        justify="left",
    ).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(6, 2))

    # Each dimension occupies two rows: the control row, and the line listing
    # every value it can be given - what the stakeholders asked to see without
    # having to open anything.
    def build_filter_row(
        label_text: str,
        dimension: str,
        var: tk.StringVar,
        hint_var: tk.StringVar,
        note_var: tk.StringVar,
        row: int,
        bottom_pad: int,
    ) -> tuple[ttk.Label, ttk.Entry, ttk.Button]:
        field_label = ttk.Label(filters_frame, text=label_text)
        field_label.grid(row=row, column=0, sticky="e", padx=10, pady=4)
        entry = ttk.Entry(filters_frame, textvariable=var, width=30)
        entry.grid(row=row, column=1, sticky="w", padx=4, pady=4)
        button = ttk.Button(
            filters_frame,
            text="Choose...",
            width=11,
            command=lambda: choose_values(dimension, var),
        )
        button.grid(row=row, column=2, sticky="w", padx=4, pady=4)
        ttk.Label(filters_frame, textvariable=hint_var).grid(
            row=row, column=3, sticky="w", padx=6, pady=4
        )
        ttk.Label(
            filters_frame,
            textvariable=note_var,
            foreground="#555555",
            wraplength=560,
            justify="left",
        ).grid(
            row=row + 1,
            column=1,
            columnspan=3,
            sticky="w",
            padx=4,
            pady=(0, bottom_pad),
        )
        return field_label, entry, button

    segment_field_label, segment_entry, segment_button = build_filter_row(
        "SEGMENT:", "SEGMENT", seg_var, seg_hint_var, seg_note_var, 1, 4
    )
    source_field_label, source_entry, source_button = build_filter_row(
        "SOURCE:", "SOURCE", source_var, source_hint_var, source_note_var, 3, 4
    )
    kw_field_label, kw_entry, kw_button = build_filter_row(
        "KW:", "KW", kw_var, kw_hint_var, kw_note_var, 5, 8
    )

    run_button = ttk.Button(root, text="Run Pipeline", command=run, width=22)
    run_button.grid(row=7, column=0, columnspan=3, **pad)
    ttk.Label(root, textvariable=status_var, width=70, anchor="w").grid(
        row=8, column=0, columnspan=3, sticky="w", **pad
    )

    def toggle_optional_rows(*_args: object) -> None:
        """Show reference / master rows and set the FILTERS state per workflow."""
        workflow_id = selected_workflow_id()
        definition = get_definition(workflow_id)
        # Input labels follow the workflow: Report Giro picks two Excel
        # workbooks, so "Raw Data File" would misdescribe both of them.
        raw_button.configure(text=raw_picker_config(workflow_id)[0])
        master_button.configure(text=master_picker_config(workflow_id)[0])
        # Filter boxes are cleared on every switch, so an empty box always means
        # "this report's automatic rules" and no value can carry over between
        # workflows. The hints beside them state what those rules are.
        for var in (seg_var, source_var, kw_var):
            var.set("")
        # Greyed out only where the dimension cannot apply, so a disabled field
        # always means the value would be ignored rather than silently discarded.
        for dimension, enabled, entry, label, button, hint_var, note_var in (
            (
                "SEGMENT",
                segment_filter_enabled(workflow_id),
                segment_entry,
                segment_field_label,
                segment_button,
                seg_hint_var,
                seg_note_var,
            ),
            (
                "SOURCE",
                source_filter_enabled(workflow_id),
                source_entry,
                source_field_label,
                source_button,
                source_hint_var,
                source_note_var,
            ),
            (
                "KW",
                kw_filter_enabled(workflow_id),
                kw_entry,
                kw_field_label,
                kw_button,
                kw_hint_var,
                kw_note_var,
            ),
        ):
            state_name = "normal" if enabled else "disabled"
            entry.configure(state=state_name)
            label.configure(state=state_name)
            hint_var.set(filter_hint_text(workflow_id, dimension))
            note_var.set(filter_options_note(workflow_id, dimension))
            # The picker has nothing to show for a report whose values were
            # never confirmed, so it greys out rather than opening an empty box.
            button.configure(
                state="normal"
                if enabled and filter_options(workflow_id, dimension)
                else "disabled"
            )
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

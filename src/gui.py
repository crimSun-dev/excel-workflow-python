"""Tkinter GUI Fallback (TDD Section 3, Phase 3).

Native OS file-picker dialog launched when the user runs the app with no CLI
arguments. Lets a non-technical user select the raw file, reference workbook,
optional SEGMEN filter, and output path, then runs the pipeline.
"""

from __future__ import annotations

from pathlib import Path

from .orchestrator import PipelineOrchestrator
from .schemas import ProcessingConfig


def launch_gui() -> None:
    """Opens the Tkinter window. Imported lazily so headless CLI use is unaffected."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Excel Workflow Processor")
    root.geometry("640x340")
    root.resizable(False, False)

    state: dict[str, Path | None] = {"raw": None, "ref": None, "out": None}

    raw_var = tk.StringVar(value="No file selected")
    ref_var = tk.StringVar(value="No file selected")
    out_var = tk.StringVar(value="./Financial_Summary_Report.xlsx")
    seg_var = tk.StringVar(value="")

    def pick_raw() -> None:
        path = filedialog.askopenfilename(
            title="Select raw pipe-delimited data file",
            filetypes=[("Text/CSV", "*.txt *.csv"), ("All files", "*.*")],
        )
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
        if not state["raw"] or not state["ref"]:
            messagebox.showerror(
                "Missing input", "Please select both a raw data file and a reference file."
            )
            return
        output = state["out"] or Path(out_var.get())
        segment = seg_var.get().strip() or None
        config = ProcessingConfig(
            raw_data_path=state["raw"],
            reference_data_path=state["ref"],
            output_report_path=output,
            segmen_filter=segment,
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
        else:
            messagebox.showerror("Pipeline failed", report.error_message or "Unknown error")

    pad = {"padx": 10, "pady": 6}
    ttk.Label(root, text="Excel Workflow Processor", font=("Segoe UI", 14, "bold")).grid(
        row=0, column=0, columnspan=3, **pad
    )

    ttk.Button(root, text="1. Raw Data File...", command=pick_raw, width=22).grid(
        row=1, column=0, **pad
    )
    ttk.Label(root, textvariable=raw_var, width=45, anchor="w").grid(row=1, column=1, columnspan=2, **pad)

    ttk.Button(root, text="2. Reference File...", command=pick_ref, width=22).grid(
        row=2, column=0, **pad
    )
    ttk.Label(root, textvariable=ref_var, width=45, anchor="w").grid(row=2, column=1, columnspan=2, **pad)

    ttk.Button(root, text="3. Output Path...", command=pick_out, width=22).grid(
        row=3, column=0, **pad
    )
    ttk.Label(root, textvariable=out_var, width=45, anchor="w").grid(row=3, column=1, columnspan=2, **pad)

    ttk.Label(root, text="SEGMEN filter (optional):").grid(row=4, column=0, **pad)
    ttk.Entry(root, textvariable=seg_var, width=30).grid(row=4, column=1, columnspan=2, sticky="w", **pad)

    ttk.Button(root, text="Run Pipeline", command=run, width=22).grid(
        row=5, column=0, columnspan=3, **pad
    )

    root.mainloop()


if __name__ == "__main__":
    launch_gui()

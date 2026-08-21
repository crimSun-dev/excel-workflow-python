"""Plain-language explanations for pipeline failures.

The orchestrator's exception boundary produces precise but technical text -
`WorkflowValidationError: Workflow 'Report Summary Akumulasi' requires
column(s) ['FBI', 'VOLUME_IN_IDR'] which are absent from the input file`. An
operator who has never seen a stack trace reads that and learns nothing about
what to do next, which is why every failure dialog now leads with what happened
and what to try, and keeps the original message underneath for whoever supports
the tool.

The matching is deliberately done on the message text rather than the exception
class: the boundary in `orchestrator.py` flattens everything to
`"{type}: {message}"`, and the GUI/CLI only ever see that string. Each rule is
keyed on a phrase the raising site owns (`EMPTY_INPUT_MESSAGE`, the exporter's
lock message, the validator's "absent from the input file"), so a rule matches
the situation it was written for and nothing else. The final fallback means an
unrecognised failure still gets a title and a next step.
"""

from __future__ import annotations

from dataclasses import dataclass

# Kept identical to the historical dialog title/body so an empty extract still
# reads as a notice rather than a crash.
_EMPTY_FILE_TITLE = "Empty file"


@dataclass(frozen=True)
class ErrorExplanation:
    """Operator-facing reading of one failure: what happened, and what to try."""

    title: str
    what_happened: str
    what_to_do: str


# Ordered (phrases, explanation) rules; the first rule with any phrase present
# in the lower-cased message wins, so specific causes precede general ones.
_RULES: tuple[tuple[tuple[str, ...], ErrorExplanation], ...] = (
    (
        ("no data rows", "file submitted for processing is empty"),
        ErrorExplanation(
            title=_EMPTY_FILE_TITLE,
            what_happened=(
                "The file you selected has no data rows - only column titles, "
                "or nothing at all. Nothing was processed and no report was "
                "written."
            ),
            what_to_do=(
                "Open the file and check the data is really in it. This usually "
                "means the download finished early or the wrong day's extract "
                "was saved, so download the extract again and re-select it."
            ),
        ),
    ),
    (
        ("absent from the input file", "requires column(s)"),
        ErrorExplanation(
            title="This file does not match the selected report",
            what_happened=(
                "The file was read successfully, but it does not contain the "
                "columns this report needs - so it is most likely the extract "
                "for a different report."
            ),
            what_to_do=(
                "Check that the report chosen in the Workflow box is the one "
                "this file belongs to, then pick the matching file. In the "
                "details below, the first list is what the report looked for "
                "and 'Available columns' is what your file actually contains."
            ),
        ),
    ),
    (
        ("may be open in microsoft excel",),
        ErrorExplanation(
            title="The report file is still open",
            what_happened=(
                "The report could not be saved because the output file is open "
                "in Excel, which locks it against changes."
            ),
            what_to_do=(
                "Close that workbook in Excel and run again, or choose a "
                "different name under Output Path."
            ),
        ),
    ),
    (
        ("requires a reference", "requires a master-data file", "requires a master"),
        ErrorExplanation(
            title="A required file is missing",
            what_happened=(
                "This report needs a second file - a reference or master-data "
                "workbook - and none was selected."
            ),
            what_to_do=(
                "Select the missing file with the numbered button for it, then "
                "run again."
            ),
        ),
    ),
    (
        ("not found",),
        ErrorExplanation(
            title="File not found",
            what_happened=(
                "One of the selected files is no longer at the location it was "
                "picked from - it was probably moved, renamed, or deleted, or "
                "it sits on a network drive that is currently unavailable."
            ),
            what_to_do=(
                "Select the file again from its current location and run again."
            ),
        ),
    ),
    (
        (
            "unable to resolve required reference columns",
            "unable to resolve required master-data columns",
            "unable to resolve the account",
        ),
        ErrorExplanation(
            title="The reference file was not recognised",
            what_happened=(
                "The reference or master workbook was opened, but none of its "
                "sheets holds the code and name columns the lookup needs, so "
                "the matching step could not run."
            ),
            what_to_do=(
                "Check you selected the mapping workbook and not another "
                "export. The details below list every sheet that was checked "
                "and every column name that was accepted - if the file uses a "
                "different heading for the branch code, that heading has to be "
                "added to the tool."
            ),
        ),
    ),
    (
        ("unsupported",),
        ErrorExplanation(
            title="Unsupported file type",
            what_happened=(
                "One of the selected files is a type this report cannot read."
            ),
            what_to_do=(
                "Use the extract as downloaded (.txt, .csv or .bin) or an Excel "
                "workbook (.xlsx / .xls / .xlsm). Re-saving the file in Excel "
                "as .xlsx is usually enough."
            ),
        ),
    ),
    (
        ("unable to decode",),
        ErrorExplanation(
            title="The file could not be read",
            what_happened=(
                "The text in this file is not in any encoding the reader "
                "recognises, which normally means it is not the plain extract "
                "it looks like."
            ),
            what_to_do=(
                "Open it in Excel, save it as .xlsx, and select that file "
                "instead."
            ),
        ),
    ),
)

_FALLBACK = ErrorExplanation(
    title="Pipeline failed",
    what_happened=(
        "The run stopped before the report was written, for a reason the tool "
        "does not have a plain-language description for yet."
    ),
    what_to_do=(
        "Check the inputs are the files this report expects and try again. If "
        "it keeps failing, send the technical details below to whoever supports "
        "this tool - they identify the exact step that stopped."
    ),
)


def explain_failure(error_message: str | None) -> ErrorExplanation:
    """Returns the plain-language reading of a pipeline error message."""
    lowered = (error_message or "").strip().lower()
    for phrases, explanation in _RULES:
        if any(phrase in lowered for phrase in phrases):
            return explanation
    return _FALLBACK


def failure_dialog_text(error_message: str | None) -> tuple[str, str]:
    """(title, body) for a failure popup: explanation first, raw message last.

    The original message is always kept - it is the only thing that names the
    offending column, sheet, or path - but it now sits under the reasoning
    instead of being the whole dialog.
    """
    message = (error_message or "").strip() or "Unknown error"
    explanation = explain_failure(message)
    body = (
        f"{explanation.what_happened}\n\n"
        f"What to do:\n{explanation.what_to_do}\n\n"
        f"Technical details (for support):\n{message}"
    )
    return explanation.title, body

"""
Combines multiple increments (each already loaded via
ui.mock_data.MockDataStore.get_increment_for_display, at whichever
version the user picked per increment) into ONE .xlsx with the same
four sheet names a single-increment export uses -- All Data, Sum Data,
Report, Changes -- for portfolio-level reporting across a project
instead of one file per increment.

All Data / Sum Data: every selected increment's rows, concatenated in
the SAME order the increments appear in the on-screen increment list
(not the order their checkboxes happened to be clicked -- simpler to
implement and more predictable for the user, since it matches what
they're already looking at top to bottom), with a new "Increment"
column prepended identifying which increment each row came from. This
resurrects the real source template's own Inc# column, which this app
has never populated until now since it only ever handled one increment
at a time. One combined totals row at the bottom, same convention as
core.excel_export's single-increment totals row, covering every
included increment together (see _combine_all_data_totals -- summing
each increment's own already-computed totals is equivalent to, and
simpler than, recomputing from raw combined rows).

Report: each selected increment's existing grouped/totaled Report,
completely unchanged (including its own per-increment "Grand Total"
row), as a separate, visibly labeled section in sequence -- followed by
ONE overall "Grand Total (All Increments)" row summing every section
together at the very end. Deliberately NOT labeled just "Grand Total"
like the per-section rows, so it's never ambiguous which total is the
combined one when scanning the sheet.

Changes: State Revision Log and Update History (see
core.excel_export._write_changes_sheet, the single-increment version of
this sheet) are each flattened across every selected increment with a
leading Increment column -- the SAME "concatenate, don't section"
treatment as All Data/Sum Data, deliberately NOT Report's per-increment
sectioning. That choice was made explicitly (not just for consistency):
neither Changes sub-table has a per-increment OR overall total worth
visually separating for the way Report's sections exist specifically to
set off its per-section and Grand Total rows -- sectioning here would
only cost readability/filterability (a flat table sorts and filters in
Excel; per-increment blocks don't) without buying back anything
sectioning is for. Each increment's own rows stay contiguous, in
on-screen order, exactly like All Data/Sum Data; Update History is
additionally newest-first WITHIN each increment's block (matching
core.excel_export's single-increment ordering), not globally
interleaved by date across increments -- a user wanting a true
cross-increment timeline can sort the Date column themselves, which
this flattened layout enables and per-increment sectioning would not
have. No combined totals row (there is nothing to sum).

Column layout is shifted +1 versus core/excel_export.py throughout (the
new leading Increment column), so this module keeps its own column
constants and small helper functions rather than reusing
core.excel_export's directly -- those hardcode positions that would
misalign by one column here. Style constants (colors/fonts) and the
small Changes-sheet formatting helpers ARE reused from there, so this
never invents a second, subtly-different palette or phrasing.
"""

from __future__ import annotations

from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from core.excel_export import (
    _BOLD_FONT,
    _DESCRIPTION_ALIGNMENT,
    _HEADER_FONT,
    _STAGE_ALIGNMENT,
    _STAGE_FILLS,
    _STAGE_FONT,
    _STAGE_MARKS,
    _SUM_DATA_STAGE_FILLS,
    _SUM_DATA_STAGE_FONTS,
    _SUM_DATA_STAGE_LABELS,
    _TOTALS_FILL,
    _TOTALS_FONT,
    _UNSAFE_FILENAME_CHARS,
    _WHITESPACE,
    _apply_description_wrap,
    _format_changes_date,
    _history_entry_detail_text,
    _history_entry_summary,
    _wrapped_row_height,
)
from core.excel_reader import STAGE_COUNT, live_sum_data_totals

INCREMENT_COL = 1
INDEX_COL = 2
DESCRIPTION_COL = 3
AGENCY_COL = 4
STAGE_FIRST_COL = 5
STAGE_LAST_COL = STAGE_FIRST_COL + STAGE_COUNT - 1

# --- Changes sheet column layout -- see _write_combined_changes_sheet.
# Same +1 shift (leading Increment column) versus
# core.excel_export._write_changes_sheet's own column constants; the two
# mini-tables share this one 7-column grid (columns 3-7 pair up the SAME
# way core.excel_export's single-increment version pairs its 6 columns --
# e.g. col 3 holds Synopsis in one table and Detail in the other, both
# long text -- just shifted +1 here).
CHANGES_REVISION_NUM_COL = 2
CHANGES_SYNOPSIS_COL = 3
CHANGES_AOR_SIGNATURE_COL = 4
CHANGES_SEOR_SIGNATURE_COL = 5
CHANGES_EFFECTIVE_DATE_COL = 6
CHANGES_HCAI_CONCURRENCE_COL = 7

CHANGES_UPDATE_NUM_COL = 2
CHANGES_DETAIL_COL = 3
CHANGES_SUMMARY_COL = 4
CHANGES_OLD_VERSION_COL = 5
CHANGES_NEW_VERSION_COL = 6
CHANGES_UPDATE_DATE_COL = 7


def default_combined_filename(project_name: str, increment_count: int) -> str:
    """"{ProjectName}_Combined_{N}-increments.xlsx" -- same filesystem-
    unsafe-character sanitization as core.excel_export.default_filename.
    """

    def clean(text: str) -> str:
        stripped = _UNSAFE_FILENAME_CHARS.sub("", text)
        return _WHITESPACE.sub("_", stripped.strip())

    return f"{clean(project_name)}_Combined_{increment_count}-increments.xlsx"


def _stage_headers() -> list[str]:
    return [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]


def _write_header(ws: Worksheet, headers: list[str]) -> None:
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT


def _write_totals_row(ws: Worksheet, values: list[Any]) -> None:
    ws.append(values)
    for cell in ws[ws.max_row]:
        cell.font = _TOTALS_FONT
        cell.fill = _TOTALS_FILL
        if cell.column > DESCRIPTION_COL:
            cell.alignment = _STAGE_ALIGNMENT


def _write_stage_cell(ws: Worksheet, excel_row: int, stage: int, status: str | None) -> None:
    cell = ws.cell(row=excel_row, column=STAGE_FIRST_COL + stage - 1, value=_STAGE_MARKS.get(status, ""))
    fill = _STAGE_FILLS.get(status)
    if fill is not None:
        cell.fill = fill
        cell.font = _STAGE_FONT
        cell.alignment = _STAGE_ALIGNMENT


def _write_sum_data_stage_cell(ws: Worksheet, excel_row: int, stage: int, status: str) -> None:
    cell = ws.cell(row=excel_row, column=STAGE_FIRST_COL + stage - 1, value=_SUM_DATA_STAGE_LABELS[status])
    cell.fill = _SUM_DATA_STAGE_FILLS[status]
    cell.font = _SUM_DATA_STAGE_FONTS[status]
    cell.alignment = _STAGE_ALIGNMENT


def _combine_all_data_totals(increments: list[Any]) -> dict[str, Any]:
    """Sums each increment's OWN already-computed all_data_totals dict
    together, key by key. Correct and simpler than recomputing from raw
    combined rows from scratch: these are already per-column numeric
    sums (see core.excel_reader.all_data_totals), and addition is
    associative/commutative, so summing the per-increment totals gives
    exactly the same result.
    """
    combined: dict[str, Any] = {f"Stage {s}": 0 for s in range(1, STAGE_COUNT + 1)}
    combined["VCR"] = 0
    combined["SUM"] = 0
    for increment in increments:
        for key, value in increment.all_data_totals.items():
            combined[key] = combined.get(key, 0) + (value or 0)
    return combined


def _write_combined_all_data_sheet(ws: Worksheet, increments: list[Any]) -> None:
    headers = ["Increment", "Index", "Description", "Approval Agency"] + _stage_headers() + ["VCR", "SUM"]
    _write_header(ws, headers)

    for increment in increments:
        for row_data in increment.all_data:
            required = set(row_data.get("required_stages", []))
            stage_status = row_data.get("stage_status", {})
            values = [
                increment.name,
                row_data.get("index"),
                row_data.get("description") or "",
                row_data.get("approval_agency") or "",
            ]
            for stage in range(1, STAGE_COUNT + 1):
                if stage in required:
                    values.append(_STAGE_MARKS.get(stage_status.get(stage), ""))
                else:
                    # Same "0, not blank" convention as core.excel_export.
                    values.append(0)
            vcr = row_data.get("vcr")
            values.append(0 if vcr is None else vcr)
            values.append(row_data.get("sum"))
            ws.append(values)

            excel_row = ws.max_row
            _apply_description_wrap(ws, excel_row, column=DESCRIPTION_COL)
            for stage in required:
                _write_stage_cell(ws, excel_row, stage, stage_status.get(stage))
            for stage in range(1, STAGE_COUNT + 1):
                if stage not in required:
                    ws.cell(row=excel_row, column=STAGE_FIRST_COL + stage - 1).alignment = _STAGE_ALIGNMENT

    totals = _combine_all_data_totals(increments)
    totals_values = ["Totals", "", "", ""]
    for stage in range(1, STAGE_COUNT + 1):
        totals_values.append(totals.get(f"Stage {stage}", 0))
    totals_values.append(totals.get("VCR", 0))
    totals_values.append(totals.get("SUM", 0))
    _write_totals_row(ws, totals_values)

    vcr_col = STAGE_LAST_COL + 1
    sum_col = vcr_col + 1
    ws.column_dimensions[get_column_letter(INCREMENT_COL)].width = 40
    ws.column_dimensions[get_column_letter(INDEX_COL)].width = 12
    ws.column_dimensions[get_column_letter(DESCRIPTION_COL)].width = 45
    ws.column_dimensions[get_column_letter(AGENCY_COL)].width = 20
    for col in range(STAGE_FIRST_COL, STAGE_LAST_COL + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9
    ws.column_dimensions[get_column_letter(vcr_col)].width = 9
    ws.column_dimensions[get_column_letter(sum_col)].width = 9
    ws.freeze_panes = ws.cell(row=2, column=STAGE_FIRST_COL).coordinate


def _write_combined_sum_data_sheet(ws: Worksheet, increments: list[Any]) -> None:
    headers = ["Increment", "Index", "Description", "Approval Agency"] + _stage_headers() + [
        "VCR", "Open", "Done", "Total", "% Complete"
    ]
    _write_header(ws, headers)

    all_rows_combined: list[dict] = []
    for increment in increments:
        for row_data in increment.sum_data:
            required = row_data.get("required_stages", [])
            stage_status = row_data.get("stage_status", {})
            live_status = {stage: stage_status.get(stage, "Open") for stage in required}
            done_count = sum(1 for status in live_status.values() if status == "Done")
            open_count = sum(1 for status in live_status.values() if status == "Open")
            total_count = len(required)
            pct = (done_count / total_count) if required else None

            values = [
                increment.name,
                row_data.get("index"),
                row_data.get("description") or "",
                row_data.get("approval_agency") or "",
            ]
            values += ["" for _ in range(1, STAGE_COUNT + 1)]  # written properly below via _write_sum_data_stage_cell
            values.append(row_data.get("vcr"))
            values.append(open_count)
            values.append(done_count)
            values.append(total_count)
            values.append(pct)
            ws.append(values)

            excel_row = ws.max_row
            _apply_description_wrap(ws, excel_row, column=DESCRIPTION_COL)
            for stage, status in live_status.items():
                _write_sum_data_stage_cell(ws, excel_row, stage, status)

            pct_col = STAGE_LAST_COL + 5
            if pct is not None:
                ws.cell(row=excel_row, column=pct_col).number_format = "0%"

            all_rows_combined.append(row_data)

    vcr_col = STAGE_LAST_COL + 1
    open_col = vcr_col + 1
    done_col = vcr_col + 2
    total_col = vcr_col + 3
    pct_col = vcr_col + 4

    # Live totals across ALL selected increments' rows together --
    # recomputed fresh every export, same as the single-increment case
    # (see core.excel_reader.live_sum_data_totals), never cached.
    totals = live_sum_data_totals(all_rows_combined)
    totals_values = ["Totals", "", "", ""]
    for stage in range(1, STAGE_COUNT + 1):
        totals_values.append(totals["stage_open_counts"][stage])
    totals_values.append(totals["vcr_open_count"])
    totals_values.append(totals["open_total"])
    totals_values.append(totals["done_total"])
    totals_values.append(totals["grand_total"])
    totals_values.append(totals["pct_complete"])
    _write_totals_row(ws, totals_values)
    if totals["pct_complete"] is not None:
        ws.cell(row=ws.max_row, column=pct_col).number_format = "0%"

    ws.column_dimensions[get_column_letter(INCREMENT_COL)].width = 40
    ws.column_dimensions[get_column_letter(INDEX_COL)].width = 12
    ws.column_dimensions[get_column_letter(DESCRIPTION_COL)].width = 45
    ws.column_dimensions[get_column_letter(AGENCY_COL)].width = 20
    for col in range(STAGE_FIRST_COL, STAGE_LAST_COL + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9
    ws.column_dimensions[get_column_letter(vcr_col)].width = 9
    ws.column_dimensions[get_column_letter(open_col)].width = 9
    ws.column_dimensions[get_column_letter(done_col)].width = 9
    ws.column_dimensions[get_column_letter(total_col)].width = 9
    ws.column_dimensions[get_column_letter(pct_col)].width = 12
    ws.freeze_panes = ws.cell(row=2, column=STAGE_FIRST_COL).coordinate


def _write_combined_report_sheet(ws: Worksheet, increments: list[Any]) -> None:
    headers = ["Approval Agency", "Index", "Description", "Total"]
    _write_header(ws, headers)

    combined_grand_total = 0.0
    for increment in increments:
        # Section header row -- visually distinct (bold + tint), reusing
        # the same "this is a summary, not a real item" styling the
        # bottom totals rows already use elsewhere in this app.
        ws.append([increment.name, "", "", ""])
        for cell in ws[ws.max_row]:
            cell.font = _TOTALS_FONT
            cell.fill = _TOTALS_FILL

        for row_data in increment.report:
            is_grand_total = row_data.get("approval_agency") == "Grand Total"
            ws.append(
                [
                    row_data.get("approval_agency") or "",
                    row_data.get("index") or "",
                    row_data.get("description") or "",
                    row_data.get("total", 0),
                ]
            )
            _apply_description_wrap(ws, ws.max_row, column=3)
            if is_grand_total:
                for cell in ws[ws.max_row]:
                    cell.font = _BOLD_FONT
                combined_grand_total += row_data.get("total", 0) or 0

    # ONE overall total across every section -- deliberately NOT labeled
    # just "Grand Total" like each section's own row above, so it's
    # never ambiguous which total is the combined one when scanning the
    # sheet.
    ws.append(["Grand Total (All Increments)", "", "", combined_grand_total])
    for cell in ws[ws.max_row]:
        cell.font = _TOTALS_FONT
        cell.fill = _TOTALS_FILL

    ws.column_dimensions[get_column_letter(1)].width = 30
    ws.column_dimensions[get_column_letter(2)].width = 12
    ws.column_dimensions[get_column_letter(3)].width = 45
    ws.column_dimensions[get_column_letter(4)].width = 10


def _write_combined_changes_sheet(ws: Worksheet, increments: list[Any]) -> None:
    """State Revision Log then Update History, each flattened across
    every selected increment with a leading Increment column -- see the
    module docstring for why this deliberately does NOT use Report's
    per-increment sectioning.
    """
    section_row = 1
    ws.cell(row=section_row, column=1, value="State Revision Log").font = _BOLD_FONT

    header_row = section_row + 1
    revision_headers = [
        "Increment", "Rev #", "Synopsis of Change", "AOR Signature", "SEOR Signature", "Effective Date",
        "HCAI Concurrence",
    ]
    for col, label in enumerate(revision_headers, start=1):
        ws.cell(row=header_row, column=col, value=label).font = _HEADER_FONT

    row = header_row + 1
    any_revisions = False
    for increment in increments:
        for entry in increment.changes_log:
            any_revisions = True
            synopsis = entry.get("synopsis") or ""
            ws.cell(row=row, column=INCREMENT_COL, value=increment.name)
            ws.cell(row=row, column=CHANGES_REVISION_NUM_COL, value=entry.get("revision_number"))
            ws.cell(row=row, column=CHANGES_SYNOPSIS_COL, value=synopsis).alignment = _DESCRIPTION_ALIGNMENT
            ws.cell(row=row, column=CHANGES_AOR_SIGNATURE_COL, value=_format_changes_date(entry.get("aor_signature_date")))
            ws.cell(row=row, column=CHANGES_SEOR_SIGNATURE_COL, value=_format_changes_date(entry.get("seor_signature_date")))
            ws.cell(row=row, column=CHANGES_EFFECTIVE_DATE_COL, value=_format_changes_date(entry.get("effective_date")))
            ws.cell(row=row, column=CHANGES_HCAI_CONCURRENCE_COL, value=entry.get("hcai_concurrence") or "")
            ws.row_dimensions[row].height = _wrapped_row_height(synopsis)
            row += 1
    if not any_revisions:
        ws.cell(row=row, column=1, value="No revision log entries found for any selected increment.")
        row += 1

    row += 1  # blank separator row
    history_section_row = row
    ws.cell(row=history_section_row, column=1, value="Update History").font = _BOLD_FONT

    history_header_row = history_section_row + 1
    history_headers = ["Increment", "Update #", "Detail", "Summary", "Old Version", "New Version", "Date"]
    for col, label in enumerate(history_headers, start=1):
        ws.cell(row=history_header_row, column=col, value=label).font = _HEADER_FONT

    row = history_header_row + 1
    any_history = False
    for increment in increments:
        # Newest first WITHIN each increment's own block -- see module
        # docstring for why this isn't globally interleaved by date.
        for i, entry in enumerate(reversed(increment.change_history), start=1):
            any_history = True
            detail = _history_entry_detail_text(entry)
            ws.cell(row=row, column=INCREMENT_COL, value=increment.name)
            ws.cell(row=row, column=CHANGES_UPDATE_NUM_COL, value=i)
            ws.cell(row=row, column=CHANGES_DETAIL_COL, value=detail).alignment = _DESCRIPTION_ALIGNMENT
            ws.cell(row=row, column=CHANGES_SUMMARY_COL, value=_history_entry_summary(entry))
            ws.cell(row=row, column=CHANGES_OLD_VERSION_COL, value=entry.get("old_version"))
            ws.cell(row=row, column=CHANGES_NEW_VERSION_COL, value=entry.get("new_version"))
            ws.cell(row=row, column=CHANGES_UPDATE_DATE_COL, value=(entry.get("timestamp") or "").split("T", 1)[0])
            ws.row_dimensions[row].height = _wrapped_row_height(detail)
            row += 1
    if not any_history:
        ws.cell(row=row, column=1, value="No updates confirmed yet for any selected increment.")

    ws.column_dimensions[get_column_letter(INCREMENT_COL)].width = 40
    ws.column_dimensions[get_column_letter(2)].width = 10
    ws.column_dimensions[get_column_letter(3)].width = 70
    ws.column_dimensions[get_column_letter(4)].width = 45
    ws.column_dimensions[get_column_letter(5)].width = 14
    ws.column_dimensions[get_column_letter(6)].width = 14
    ws.column_dimensions[get_column_letter(7)].width = 22


def export_combined_report(increments: list[Any], path: str) -> None:
    """Writes increments (a list of ui.mock_data.Increment, each already
    populated by MockDataStore.get_increment_for_display -- one entry
    per user-selected increment, at whichever version was picked for
    it, in the order they should appear in the output) to one .xlsx
    workbook with four combined sheets: All Data, Sum Data, Report,
    Changes.
    """
    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "All Data"
    _write_combined_all_data_sheet(ws_all, increments)

    ws_sum = wb.create_sheet("Sum Data")
    _write_combined_sum_data_sheet(ws_sum, increments)

    ws_report = wb.create_sheet("Report")
    _write_combined_report_sheet(ws_report, increments)

    ws_changes = wb.create_sheet("Changes")
    _write_combined_changes_sheet(ws_changes, increments)

    wb.save(path)

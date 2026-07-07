"""
Exports a cached ui.mock_data.Increment -- All Data, Sum Data, and
Report, all loaded together by
ui.mock_data.MockDataStore.get_increment_for_display() -- to a real
.xlsx workbook with three sheets, named and ordered exactly "All Data",
"Sum Data", "Report", matching what's on screen in
ui.pages.data_view_page's three tabs (including any in-memory status
edits made this session that haven't been re-parsed from disk).

All Data's stage cell marks/colors mirror ui.pages.data_view_page's
_stage_status_icon() exactly ("X" on green = Done, "1" on red = Open)
using the same RGB values core.excel_reader._fill_color_status()
classifies as green/red, so re-uploading an exported file as a new
version is read back correctly by the app's own fill-color status
seeding (core.status_tracker.apply_file_derived_seed) -- export and
re-import round-trip the status, not just the raw data.

Sum Data's stage cells intentionally look different: literal "Done"/
"Open" text on light green/pink fills matching Excel's own built-in
"Good"/"Bad" conditional-formatting cell styles (see
_write_sum_data_stage_cell), not All Data's dark-fill "X"/"1" marks --
Sum Data is a status report, not a re-importable source sheet, so it's
free to read the way a human expects a status column to read.
"""

from __future__ import annotations

import re
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from core.excel_reader import STAGE_COUNT

_STAGE_MARKS = {"Done": "X", "Open": "1"}
_STAGE_FILLS = {
    "Done": PatternFill(start_color="FF2E7D32", end_color="FF2E7D32", fill_type="solid"),
    "Open": PatternFill(start_color="FFC62828", end_color="FFC62828", fill_type="solid"),
}
_STAGE_FONT = Font(color="FFFFFFFF", bold=True)
_STAGE_ALIGNMENT = Alignment(horizontal="center", vertical="center")

# Sum Data sheet's own required-stage cell style -- literal "Done"/"Open"
# text on a full-cell fill, matching Excel's own built-in "Good"/"Bad"
# conditional-formatting cell styles exactly (these are those styles'
# real fill/font colors), distinct from All Data's "X"/"1" dark-fill
# marks above -- Sum Data is meant to read like a status report, All
# Data like the source file's own raw markers.
_SUM_DATA_STAGE_LABELS = {"Done": "Done", "Open": "Open"}
_SUM_DATA_STAGE_FILLS = {
    "Done": PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid"),
    "Open": PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid"),
}
_SUM_DATA_STAGE_FONTS = {
    "Done": Font(color="FF006100"),
    "Open": Font(color="FF9C0006"),
}
_HEADER_FONT = Font(bold=True)
_BOLD_FONT = Font(bold=True)
_DESCRIPTION_ALIGNMENT = Alignment(wrap_text=True, vertical="top")

INDEX_COL = 1
DESCRIPTION_COL = 2
AGENCY_COL = 3
STAGE_FIRST_COL = 4
STAGE_LAST_COL = STAGE_FIRST_COL + STAGE_COUNT - 1

# Descriptions are 3 logical lines (component name, code citation, test
# description) joined with embedded \n. openpyxl can't auto-fit row height
# for wrapped text, so this is a fixed height sized for 3 lines at the
# default 11pt font (~15pt/line) plus a little padding.
_DESCRIPTION_ROW_HEIGHT = 50

_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
_WHITESPACE = re.compile(r"\s+")


def default_filename(project_name: str, increment_name: str, version: int) -> str:
    """"{ProjectName}_{IncrementName}_v{version}.xlsx", with filesystem-
    unsafe characters (from either name -- project/increment names are
    free text the user typed or pulled from a workbook, not guaranteed
    filename-safe) stripped and whitespace collapsed to underscores.
    """

    def clean(text: str) -> str:
        stripped = _UNSAFE_FILENAME_CHARS.sub("", text)
        return _WHITESPACE.sub("_", stripped.strip())

    return f"{clean(project_name)}_{clean(increment_name)}_v{version}.xlsx"


def _write_header(ws: Worksheet, headers: list[str]) -> None:
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT


def _stage_headers() -> list[str]:
    return [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]


def _set_common_column_widths(ws: Worksheet, extra_col_widths: dict[int, int]) -> None:
    ws.column_dimensions[get_column_letter(INDEX_COL)].width = 12
    ws.column_dimensions[get_column_letter(DESCRIPTION_COL)].width = 45
    ws.column_dimensions[get_column_letter(AGENCY_COL)].width = 20
    for col in range(STAGE_FIRST_COL, STAGE_LAST_COL + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9
    for col, width in extra_col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def _apply_description_wrap(ws: Worksheet, excel_row: int) -> None:
    ws.cell(row=excel_row, column=DESCRIPTION_COL).alignment = _DESCRIPTION_ALIGNMENT
    ws.row_dimensions[excel_row].height = _DESCRIPTION_ROW_HEIGHT


def _write_stage_cell(ws: Worksheet, excel_row: int, stage: int, status: str | None) -> None:
    cell = ws.cell(row=excel_row, column=STAGE_FIRST_COL + stage - 1, value=_STAGE_MARKS.get(status, ""))
    fill = _STAGE_FILLS.get(status)
    if fill is not None:
        cell.fill = fill
        cell.font = _STAGE_FONT
        cell.alignment = _STAGE_ALIGNMENT


def _write_sum_data_stage_cell(ws: Worksheet, excel_row: int, stage: int, status: str) -> None:
    cell = ws.cell(
        row=excel_row, column=STAGE_FIRST_COL + stage - 1, value=_SUM_DATA_STAGE_LABELS[status]
    )
    cell.fill = _SUM_DATA_STAGE_FILLS[status]
    cell.font = _SUM_DATA_STAGE_FONTS[status]
    cell.alignment = _STAGE_ALIGNMENT


def _write_all_data_sheet(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    headers = ["Index", "Description", "Approval Agency"] + _stage_headers() + ["VCR", "SUM"]
    _write_header(ws, headers)

    for row_data in rows:
        required = set(row_data.get("required_stages", []))
        stage_status = row_data.get("stage_status", {})
        values = [
            row_data.get("index"),
            row_data.get("description") or "",
            row_data.get("approval_agency") or "",
        ]
        for stage in range(1, STAGE_COUNT + 1):
            status = stage_status.get(stage) if stage in required else None
            values.append(_STAGE_MARKS.get(status, ""))
        values.append(row_data.get("vcr"))
        values.append(row_data.get("sum"))
        ws.append(values)

        excel_row = ws.max_row
        _apply_description_wrap(ws, excel_row)
        for stage in required:
            _write_stage_cell(ws, excel_row, stage, stage_status.get(stage))

    vcr_col = STAGE_LAST_COL + 1
    sum_col = vcr_col + 1
    _set_common_column_widths(ws, {vcr_col: 9, sum_col: 9})
    ws.freeze_panes = ws.cell(row=2, column=STAGE_FIRST_COL).coordinate


def _write_sum_data_sheet(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    """Mirrors ui.pages.data_view_page's Sum Data tab exactly: live
    status.json-backed Done/Open per required stage (an unset required
    stage defaults to "Open" -- see that module's docstring), plus live
    Open/Done/Total/% Complete columns, rather than build_sum_data()'s
    own raw-file-derived status text.
    """
    headers = ["Index", "Description", "Approval Agency"] + _stage_headers() + [
        "VCR", "Open", "Done", "Total", "% Complete"
    ]
    _write_header(ws, headers)

    for row_data in rows:
        required = row_data.get("required_stages", [])
        stage_status = row_data.get("stage_status", {})
        live_status = {stage: stage_status.get(stage, "Open") for stage in required}
        done_count = sum(1 for status in live_status.values() if status == "Done")
        open_count = sum(1 for status in live_status.values() if status == "Open")
        total_count = len(required)
        pct = (done_count / total_count) if required else None

        values = [
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
        _apply_description_wrap(ws, excel_row)
        for stage, status in live_status.items():
            _write_sum_data_stage_cell(ws, excel_row, stage, status)

        pct_col = STAGE_LAST_COL + 5
        if pct is not None:
            ws.cell(row=excel_row, column=pct_col).number_format = "0%"

    vcr_col = STAGE_LAST_COL + 1
    open_col = vcr_col + 1
    done_col = vcr_col + 2
    total_col = vcr_col + 3
    pct_col = vcr_col + 4
    _set_common_column_widths(ws, {vcr_col: 9, open_col: 9, done_col: 9, total_col: 9, pct_col: 12})
    ws.freeze_panes = ws.cell(row=2, column=STAGE_FIRST_COL).coordinate


def _write_report_sheet(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    headers = ["Approval Agency", "Index", "Description", "Total"]
    _write_header(ws, headers)

    for row_data in rows:
        is_grand_total = row_data.get("approval_agency") == "Grand Total"
        ws.append(
            [
                row_data.get("approval_agency") or "",
                row_data.get("index") or "",
                row_data.get("description") or "",
                row_data.get("total", 0),
            ]
        )
        if is_grand_total:
            for cell in ws[ws.max_row]:
                cell.font = _BOLD_FONT

    ws.column_dimensions[get_column_letter(1)].width = 30
    ws.column_dimensions[get_column_letter(2)].width = 12
    ws.column_dimensions[get_column_letter(3)].width = 45
    ws.column_dimensions[get_column_letter(4)].width = 10


def export_increment(increment: Any, path: str) -> None:
    """Writes increment (a ui.mock_data.Increment, with all_data/
    sum_data/report already populated by
    MockDataStore.get_increment_for_display) to a new .xlsx workbook at
    path with three sheets, in order: All Data, Sum Data, Report.
    """
    wb = openpyxl.Workbook()
    ws_all_data = wb.active
    ws_all_data.title = "All Data"
    _write_all_data_sheet(ws_all_data, increment.all_data)

    ws_sum_data = wb.create_sheet("Sum Data")
    _write_sum_data_sheet(ws_sum_data, increment.sum_data)

    ws_report = wb.create_sheet("Report")
    _write_report_sheet(ws_report, increment.report)

    wb.save(path)

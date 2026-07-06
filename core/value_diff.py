"""
Value-level change detection between two versions of the same TIO
increment -- complements core/structure_diff.py's compare_structure(),
which only reports Indexes added or removed entirely. This module covers
the other case: an Index present in BOTH files where one of its values
changed (e.g. a stage mark went from blank/X to a real test count) --
otherwise invisible to a human reviewer, since the item itself isn't new
or missing, just quietly different.

For every Index common to both files, across the same sheets
core.structure_diff.SHEETS_TO_COMPARE covers, this compares:
  - all 42 Stage columns -- every position, not just currently-required
    ones (per required_stages_by_index()), since a value appearing where
    there was previously none IS the interesting case
  - Description, Approval Agency, OPAA

0/""/None are treated as equivalent "blank" on both sides (reusing
core.excel_reader._blank_to_none, the same normalization
required_stages_by_index() and build_sum_data() rely on), so e.g. a
stage cell holding 0 on one side and blank on the other is not reported
as a change.

Detection only -- like compare_structure(), this never resolves or
merges anything; it just makes an otherwise-silent value change visible
to a human before the new file is confirmed.
"""

from __future__ import annotations

from typing import Any

from core.excel_reader import (
    GRID_SHEETS,
    MILESTONE_SHEET,
    STAGE_COUNT,
    BlueSheetRecord,
    _blank_to_none,
    open_workbook,
    parse_grid_sheet,
    parse_milestone_sheet,
)

SHEETS_TO_COMPARE = [*GRID_SHEETS, MILESTONE_SHEET]

# {index: {field_label: (old_value, new_value)}} -- only changed
# indexes/fields are present; unchanged data is never included.
ValueChanges = dict[str, dict[str, tuple[Any, Any]]]


def _records_by_index(wb, sheet_name: str) -> dict[str, BlueSheetRecord]:
    ws = wb[sheet_name]
    records = parse_milestone_sheet(ws) if sheet_name == MILESTONE_SHEET else parse_grid_sheet(ws)
    return {r.index: r for r in records if r.index}


def _normalized_eq(old_value: Any, new_value: Any) -> bool:
    return _blank_to_none(old_value) == _blank_to_none(new_value)


def compare_values(old_workbook_path: Any, new_workbook_path: Any) -> ValueChanges:
    """Compares two versions of the same TIO increment at the value level.

    Each argument may be a file path (str/Path) or an already-open
    Workbook (see core.excel_reader.open_workbook) -- same convention as
    compare_structure(), for callers that need to reuse an open workbook
    for something else too.
    """
    wb_old = open_workbook(old_workbook_path)
    wb_new = open_workbook(new_workbook_path)

    changes: ValueChanges = {}
    for sheet_name in SHEETS_TO_COMPARE:
        old_by_index = _records_by_index(wb_old, sheet_name)
        new_by_index = _records_by_index(wb_new, sheet_name)

        for index in sorted(set(old_by_index) & set(new_by_index)):
            old_rec = old_by_index[index]
            new_rec = new_by_index[index]
            fields: dict[str, tuple[Any, Any]] = {}

            for stage in range(1, STAGE_COUNT + 1):
                old_v = old_rec.stage_values.get(stage)
                new_v = new_rec.stage_values.get(stage)
                if not _normalized_eq(old_v, new_v):
                    fields[f"Stage {stage}"] = (old_v, new_v)

            for label, old_v, new_v in (
                ("Description", old_rec.description, new_rec.description),
                ("Approval Agency", old_rec.approval_agency, new_rec.approval_agency),
                ("OPAA", old_rec.opaa_info, new_rec.opaa_info),
            ):
                if not _normalized_eq(old_v, new_v):
                    fields[label] = (old_v, new_v)

            if fields:
                changes[index] = fields

    return changes


def has_any_value_changes(changes: ValueChanges) -> bool:
    return bool(changes)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python -m core.value_diff <old_workbook.xlsm> <new_workbook.xlsm>")
        raise SystemExit(2)

    result = compare_values(sys.argv[1], sys.argv[2])
    if not result:
        print("No value changes detected.")
    else:
        for index, fields in result.items():
            print(f"--- {index} ---")
            for field_label, (old_v, new_v) in fields.items():
                print(f"  {field_label}: {old_v!r} -> {new_v!r}")

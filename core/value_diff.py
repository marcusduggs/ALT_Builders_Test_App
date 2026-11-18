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
    BlueSheetRecord,
    _blank_to_none,
    assign_row_keys,
    open_workbook,
    parse_grid_sheet,
    parse_milestone_sheet,
)

SHEETS_TO_COMPARE = [*GRID_SHEETS, MILESTONE_SHEET]

# {index: {field_label: (old_value, new_value)}} -- only changed
# indexes/fields are present; unchanged data is never included.
ValueChanges = dict[str, dict[str, tuple[Any, Any]]]


def _records_by_row_key(wb, sheet_name: str) -> dict[str, BlueSheetRecord]:
    """Same records this module always parsed, keyed by row key (see
    core.excel_reader.assign_row_keys()) instead of bare Index.

    *** FIXED: used to be `{r.index: r for r in records if r.index}` ***
    A plain dict comprehension keyed by Index is last-write-wins for a
    repeated Index -- confirmed real on F-Cons Verif's Reference Number
    column (can repeat) and on grid sheets (HMR 10A: 39 duplicate Index
    values). That silently compared the WRONG physical rows against each
    other whenever a duplicated Index's old-file and new-file "last
    occurrence" weren't actually the same row, or hid one occurrence's
    changes entirely behind another's. Row keys disambiguate each
    physical occurrence (1st, 2nd, ... in file order) so the Nth old
    occurrence is paired against the Nth new occurrence instead.

    Scoped to just THIS sheet's own records (not the whole workbook) --
    matching this module's existing per-sheet SHEETS_TO_COMPARE loop,
    not core.excel_reader's own whole-workbook scope. For the common
    case (an Index that appears once in this sheet), the row key is
    just the bare Index, same as before.
    """
    ws = wb[sheet_name]
    if sheet_name == MILESTONE_SHEET:
        # stage_count doesn't matter here -- F-Cons Verif records carry
        # no real stage values either way, and compare_values() below
        # unions each PAIR of records' own stage_values keys rather
        # than trusting a fixed count, so an empty dict here (stage_count=0)
        # behaves identically to the old fixed-42-of-None dict for
        # comparison purposes.
        records = parse_milestone_sheet(ws, stage_count=0)
    else:
        records, _ = parse_grid_sheet(ws)
    row_keys = assign_row_keys(records)
    return {key: rec for rec, key in zip(records, row_keys) if key}


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
        old_by_key = _records_by_row_key(wb_old, sheet_name)
        new_by_key = _records_by_row_key(wb_new, sheet_name)

        for row_key in sorted(set(old_by_key) & set(new_by_key)):
            old_rec = old_by_key[row_key]
            new_rec = new_by_key[row_key]
            fields: dict[str, tuple[Any, Any]] = {}

            # Union of both records' OWN stage_values keys, not a fixed
            # 42 -- each record's dict is already sized to its own
            # file's real stage count (see core.excel_reader.parse_grid_sheet),
            # which can legitimately differ between old and new (a
            # real re-scope, not just a malformed file). A stage
            # number only one side has still needs comparing (that's
            # exactly the kind of otherwise-silent change this module
            # exists to surface), so union rather than intersect.
            for stage in sorted(set(old_rec.stage_values) | set(new_rec.stage_values)):
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
                # Reported under the plain display Index (new_rec.index),
                # not the row key -- the public return shape stays
                # exactly what it was before row keys existed, since
                # every caller (ui/mock_data.py's value_changed_items)
                # shows this key directly to the user. If a duplicated
                # Index has MULTIPLE physical occurrences that each
                # changed, their fields are merged here under that one
                # shared Index key (last-occurrence-wins per FIELD on a
                # genuine collision) -- a narrow, documented limitation
                # for an already-rare case (a duplicated Index with
                # independent per-occurrence edits), not silently wrong
                # pairing like the old bug; the real fix here is that
                # OLD and NEW are now compared occurrence-by-occurrence
                # instead of old's-last-row vs new's-last-row regardless
                # of which occurrence actually changed.
                changes.setdefault(new_rec.index, {}).update(fields)

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

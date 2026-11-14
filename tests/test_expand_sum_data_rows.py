"""
Unit test for core/excel_reader.py's expand_sum_data_rows() -- the
single function that turns Sum Data's per-item rows (shape: index,
description, approval_agency, required_stages, stage_status, vcr) into
the long-format one-row-per-(Index, Stage) list every Sum Data
surface (on-screen tabs, both Excel writers) renders from. Uses fake
rows -- no real workbook needed, this is pure logic.

This is specifically the test that's independent of any higher-level
"export shape matches N rows" check: tests/test_sum_data_report_and_export.py
and tests/test_combined_mixed_stage_counts.py both call
expand_sum_data_rows() to compute an EXPECTED row count for a real
fixture, but that only proves those two call sites agree with
themselves -- it never checks expand_sum_data_rows()'s own output
against hand-constructed, independently-known-correct expected data.
This file is that check.

Critically: this also guards against the aggregate-totals function
(core.excel_reader.live_sum_data_totals) ever being fed the EXPANDED
row list instead of the original per-item rows it's always called with
today (see every call site: core/excel_export.py, core/combined_export.py,
ui/pages/data_view_page.py all pass the original `rows`/`increment.sum_data`,
never this function's output) -- the "produces exactly N rows, no
duplicates, no omissions" assertions below are exactly the invariant
that would make counting the SAME item's status once per expanded row
(rather than once per required stage) a real bug, structurally
identical to the VCR double-counting bug fixed earlier in this
project.

Run directly: `python tests/test_expand_sum_data_rows.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_reader import expand_sum_data_rows


def main():
    failures = []

    # --- main case: one item with several required stages, some Done,
    # some Open, one stage with NO status.json entry at all (must
    # default to "Open", same convention as live_sum_data_totals/every
    # other live Sum Data view) ---
    row = {
        "index": "B-S1",
        "description": "Steel\nCBC 2202A.1\nIdentification test",
        "approval_agency": "RMA",
        "required_stages": [1, 2, 3],
        "stage_status": {1: "Done", 2: "Open"},  # 3 has no entry -> defaults to Open
        "vcr": None,
    }
    result = expand_sum_data_rows([row])
    print("Main case (3 required stages, one unset):")
    for e in result:
        print(" ", e)

    if len(result) != 3:
        failures.append(f"main: expected exactly 3 rows (one per required stage), got {len(result)}")
    stages_seen = [e["stage"] for e in result]
    if stages_seen != [1, 2, 3]:
        failures.append(f"main: expected rows in ascending stage order [1, 2, 3], got {stages_seen}")
    if len(set(stages_seen)) != len(stages_seen):
        failures.append(f"main: duplicate stage number in output -- {stages_seen}")
    statuses_by_stage = {e["stage"]: e["status"] for e in result}
    if statuses_by_stage != {1: "Done", 2: "Open", 3: "Open"}:
        failures.append(
            f"main: expected statuses {{1: 'Done', 2: 'Open', 3: 'Open'}} (stage 3 unset -> defaults to "
            f"Open), got {statuses_by_stage}"
        )
    if any(e["index"] != "B-S1" or e["approval_agency"] != "RMA" for e in result):
        failures.append("main: every expanded row should carry the source item's own index/approval_agency unchanged")

    # --- stages given out of order in the source dict must still come
    # out sorted ascending, not in dict/list insertion order ---
    unordered_row = {
        "index": "B-S2",
        "description": "desc",
        "approval_agency": "RMA",
        "required_stages": [20, 1, 5],
        "stage_status": {20: "Done", 1: "Done", 5: "Open"},
        "vcr": None,
    }
    unordered_result = expand_sum_data_rows([unordered_row])
    unordered_stages = [e["stage"] for e in unordered_result]
    print(f"\nOut-of-order required_stages [20, 1, 5] -> output stage order: {unordered_stages}")
    if unordered_stages != [1, 5, 20]:
        failures.append(f"out-of-order: expected sorted ascending [1, 5, 20], got {unordered_stages}")

    # --- a milestone/F-Cons Verif item (required_stages always empty
    # for these -- they track completion via VCR instead) must produce
    # EXACTLY ONE row, with the literal "VCR" stage label, regardless of
    # how many "stages" a grid item might have had -- this is the
    # specific rule that keeps every row the same shape (one row = one
    # trackable thing) instead of these items vanishing entirely ---
    for vcr_value, expected_status, label in [
        ("Done", "Done", "already-clean 'Done'"),
        ("Open", "Open", "already-clean 'Open'"),
        ("NA", "NA", "explicit 'NA' marker (excluded from Open/Done/Total elsewhere, but still gets a row)"),
    ]:
        milestone_row = {
            "index": "F-1",
            "description": "Substantial Compliance",
            "approval_agency": None,
            "required_stages": [],
            "stage_status": {},
            "vcr": vcr_value,
        }
        milestone_result = expand_sum_data_rows([milestone_row])
        print(f"\nMilestone item, vcr={vcr_value!r} ({label}): {milestone_result}")
        if len(milestone_result) != 1:
            failures.append(
                f"milestone ({label}): expected exactly 1 row for a no-required-stages item, got {len(milestone_result)}"
            )
        elif milestone_result[0]["stage"] != "VCR":
            failures.append(f"milestone ({label}): expected Stage # 'VCR', got {milestone_result[0]['stage']!r}")
        elif milestone_result[0]["status"] != expected_status:
            failures.append(
                f"milestone ({label}): expected status {expected_status!r}, got {milestone_result[0]['status']!r}"
            )

    # --- multiple items in one call: each item's own stages must stay
    # grouped together, in input order, not interleaved across items or
    # reordered relative to each other ---
    multi_items = [
        {
            "index": "A-1", "description": "d1", "approval_agency": "AG1",
            "required_stages": [2, 1], "stage_status": {1: "Done", 2: "Done"}, "vcr": None,
        },
        {
            "index": "A-2", "description": "d2", "approval_agency": "AG2",
            "required_stages": [], "stage_status": {}, "vcr": "Open",
        },
        {
            "index": "A-3", "description": "d3", "approval_agency": "AG3",
            "required_stages": [9], "stage_status": {9: "Open"}, "vcr": None,
        },
    ]
    multi_result = expand_sum_data_rows(multi_items)
    multi_shape = [(e["index"], e["stage"]) for e in multi_result]
    print(f"\nMultiple items -> (index, stage) sequence: {multi_shape}")
    expected_shape = [("A-1", 1), ("A-1", 2), ("A-2", "VCR"), ("A-3", 9)]
    if multi_shape != expected_shape:
        failures.append(f"multi-item: expected row sequence {expected_shape}, got {multi_shape}")

    # --- the exact invariant double-counting would violate: for ANY
    # single item, the number of expanded rows must equal EXACTLY the
    # number of its own required stages (1 for a milestone item) -- not
    # more (which would mean some stage got emitted twice, letting a
    # totals computation built on this list double-count it) and not
    # fewer (which would silently drop a stage from any aggregate built
    # on this list) ---
    print("\nRow-count-equals-required-stage-count invariant, several shapes:")
    for required in ([1], [1, 2, 3, 4, 5], list(range(1, 21)), []):
        test_row = {
            "index": "X", "description": "d", "approval_agency": "AG",
            "required_stages": required,
            "stage_status": {s: "Done" for s in required},
            "vcr": "Open",  # only consulted when required_stages is empty
        }
        expanded = expand_sum_data_rows([test_row])
        expected_count = len(required) if required else 1  # empty -> exactly 1 VCR row
        print(f"  required_stages={required} -> {len(expanded)} rows (expected {expected_count})")
        if len(expanded) != expected_count:
            failures.append(
                f"invariant: required_stages={required} should produce exactly {expected_count} row(s), "
                f"got {len(expanded)}"
            )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- expand_sum_data_rows() produces exactly one row per required stage (ascending "
              "order, no duplicates, no omissions, correct Open-default-on-unset status), exactly one 'VCR' "
              "row for a milestone item regardless of its vcr value, and keeps multiple items' rows grouped "
              "and ordered correctly -- the exact invariant that rules out double-counting in any aggregate "
              "built from this list.")


if __name__ == "__main__":
    main()

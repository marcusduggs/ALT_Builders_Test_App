"""
Tests core/value_diff.py's compare_values() against three real fixture
transitions:

  - demo_before.xlsm -> demo_after.xlsm only ADDS items (B-C18/19/20),
    so compare_values() should find ZERO value changes -- confirms this
    module doesn't overlap with or duplicate what
    core/structure_diff.py's compare_structure() already reports for
    that same transition (see test_sum_data_report_and_export.py's #6).
  - demo_after.xlsm -> demo_after_2.xlsm changes exactly one existing
    value (B-C18's Stage 19: "X" -> 2) with no items added/removed --
    the case compare_structure() is blind to and compare_values() exists
    to catch.
  - sample_increment.xlsm against itself: zero value changes, matching
    test_structure_diff.py's self-comparison sanity check.

Run directly: `python tests/test_value_diff.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.value_diff import compare_values

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
DEMO_BEFORE = os.path.join(FIXTURES, "demo_before.xlsm")
DEMO_AFTER = os.path.join(FIXTURES, "demo_after.xlsm")
DEMO_AFTER_2 = os.path.join(FIXTURES, "demo_after_2.xlsm")
SAMPLE = os.path.join(FIXTURES, "sample_increment.xlsm")


def main():
    failures = []

    # ------------------------------------------------------------
    # 1 -- demo_before.xlsm -> demo_after.xlsm: additions only, zero
    # value changes among items present in both
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- demo_before.xlsm -> demo_after.xlsm (additions only)")
    print("=" * 70)
    changes_1 = compare_values(DEMO_BEFORE, DEMO_AFTER)
    print(f"Value changes: {changes_1 or 'none'}")
    if changes_1:
        failures.append(f"1: expected zero value changes for an additions-only transition, got {changes_1}")

    # ------------------------------------------------------------
    # 2 -- demo_after.xlsm -> demo_after_2.xlsm: exactly one value
    # change, B-C18 Stage 19, "X" -> 2
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- demo_after.xlsm -> demo_after_2.xlsm (one value change)")
    print("=" * 70)
    changes_2 = compare_values(DEMO_AFTER, DEMO_AFTER_2)
    print(f"Value changes: {changes_2}")

    if list(changes_2.keys()) != ["B-C18"]:
        failures.append(f"2: expected exactly one changed index, 'B-C18', got {list(changes_2.keys())}")
    else:
        fields = changes_2["B-C18"]
        if list(fields.keys()) != ["Stage 19"]:
            failures.append(f"2: expected exactly one changed field on B-C18, 'Stage 19', got {list(fields.keys())}")
        else:
            old_v, new_v = fields["Stage 19"]
            if old_v != "X" or new_v != 2:
                failures.append(f"2: expected Stage 19 'X' -> 2, got {old_v!r} -> {new_v!r}")

    # ------------------------------------------------------------
    # 3 -- sample_increment.xlsm against itself: zero value changes
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- sample_increment.xlsm vs itself")
    print("=" * 70)
    changes_3 = compare_values(SAMPLE, SAMPLE)
    print(f"Value changes: {changes_3 or 'none'}")
    if changes_3:
        failures.append(f"3: expected zero value changes comparing a file against itself, got {changes_3}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- value-level change detection matches expectations on all 3 fixture transitions")


if __name__ == "__main__":
    main()

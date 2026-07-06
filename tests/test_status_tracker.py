"""
Unit test for core/status_tracker.py's carry_forward_status at (Index,
Stage) granularity, using fake before/after required-stages maps -- no
real workbook needed, this is pure logic.

Run directly: `python tests/test_status_tracker.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.status_tracker import carry_forward_status


def main():
    failures = []

    # --- main case: some carried over, a newly-required stage on an
    # existing item, a brand-new item, a whole item removed, and one
    # specific stage no longer required on an item that still exists ---
    previous_status = {
        "B-C1": {5: "Done", 12: "Open"},
        "B-C2": {3: "Open"},
        "B-F1": {1: "Done"},
        "C-F1": {2: "Open"},  # whole item will disappear in the new file
    }
    required_stages = {
        "B-C1": [5, 12, 20],  # 5,12 carried over; 20 is newly required
        "B-C2": [3],  # unchanged
        "B-F1": [],  # item still exists but no longer requires stage 1
        "B-C99": [7],  # brand new item
    }

    result = carry_forward_status(previous_status, required_stages)
    print("Main case:")
    print("  carried_over:", result.carried_over)
    print("  needs_status:", result.needs_status)
    print("  removed:     ", result.removed)

    expected_carried = {"B-C1": {5: "Done", 12: "Open"}, "B-C2": {3: "Open"}}
    if result.carried_over != expected_carried:
        failures.append(f"carried_over mismatch: expected {expected_carried}, got {result.carried_over}")

    expected_needs_status = {"B-C1": [20], "B-C99": [7]}
    if result.needs_status != expected_needs_status:
        failures.append(f"needs_status mismatch: expected {expected_needs_status}, got {result.needs_status}")

    expected_removed = {"B-F1": [1], "C-F1": [2]}
    if result.removed != expected_removed:
        failures.append(f"removed mismatch: expected {expected_removed}, got {result.removed}")

    if not result.has_changes:
        failures.append("has_changes should be True when there are new/removed (index, stage) pairs")

    # --- edge case: first-ever parse, no prior status at all ---
    empty_prev = carry_forward_status({}, {"X-1": [1, 2], "X-2": [3]})
    print("\nEmpty-previous-status case:", empty_prev)
    if (
        empty_prev.needs_status != {"X-1": [1, 2], "X-2": [3]}
        or empty_prev.carried_over != {}
        or empty_prev.removed != {}
    ):
        failures.append(f"empty-previous-status case failed: {empty_prev}")

    # --- edge case: every previously tracked item disappeared ---
    empty_new = carry_forward_status({"Y-1": {1: "Done", 2: "Open"}}, {})
    print("Empty-new-list case:", empty_new)
    if empty_new.removed != {"Y-1": [1, 2]} or empty_new.carried_over != {} or empty_new.needs_status != {}:
        failures.append(f"empty-new-list case failed: {empty_new}")

    # --- edge case: nothing changed at all ---
    no_change = carry_forward_status({"Z-1": {1: "Done"}}, {"Z-1": [1]})
    print("No-change case:", no_change)
    if no_change.carried_over != {"Z-1": {1: "Done"}} or no_change.has_changes:
        failures.append(f"no-change case failed: {no_change}")

    # --- edge case: an item required stage is dropped to zero required
    # stages (but the item itself still appears in required_stages) ---
    zero_required = carry_forward_status({"W-1": {1: "Done"}}, {"W-1": []})
    print("Zero-required-stages case:", zero_required)
    if zero_required.removed != {"W-1": [1]} or zero_required.carried_over != {} or zero_required.needs_status != {}:
        failures.append(f"zero-required-stages case failed: {zero_required}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- per-stage carryover, new-requirement detection, and removed-requirement detection all correct")


if __name__ == "__main__":
    main()

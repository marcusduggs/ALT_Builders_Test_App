"""
Sanity check for core/structure_diff.py: comparing sample_increment.xlsm
against itself should report zero added/removed indexes and zero column
anomalies across every blue sheet that feeds All Data.

This is a smoke test only -- there is no second real file with an actual
structural difference to test against yet. It proves compare_structure()
doesn't cry wolf on an unchanged file, which matters because the whole
point of this module is that a human trusts its output when it *does*
flag something.

Run directly: `python tests/test_structure_diff.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.structure_diff import compare_structure

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")


def main():
    print(f"Comparing fixture against itself: {FIXTURE}")
    results = compare_structure(FIXTURE, FIXTURE)

    problems = []
    for sheet_name, diff in results.items():
        print(f"\n--- {sheet_name} ---")
        print(f"  unchanged: {len(diff.unchanged_indexes)}")
        print(f"  added:     {diff.added_indexes}")
        print(f"  removed:   {diff.removed_indexes}")
        print(f"  column anomalies: {diff.column_anomalies or 'none'}")

        if diff.added_indexes:
            problems.append(f"{sheet_name}: unexpected added indexes vs itself: {diff.added_indexes}")
        if diff.removed_indexes:
            problems.append(f"{sheet_name}: unexpected removed indexes vs itself: {diff.removed_indexes}")
        if diff.column_anomalies:
            problems.append(f"{sheet_name}: unexpected column anomalies vs itself: {diff.column_anomalies}")
        if diff.unchanged_indexes == [] and sheet_name != "F-Cons Verif":
            problems.append(f"{sheet_name}: zero indexes parsed at all -- parser or fixture problem, not a diff issue")

    print("\n" + "=" * 70)
    if problems:
        print("RESULT: FAIL")
        for p in problems:
            print(" -", p)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- comparing the fixture against itself reports zero changes, as expected")


if __name__ == "__main__":
    main()

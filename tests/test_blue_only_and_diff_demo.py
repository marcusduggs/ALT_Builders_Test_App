"""
Validates the two walkthrough-video demo fixtures:

  tests/fixtures/demo_before.xlsm -- a copy of sample_increment.xlsm with
      the All Data / Sum Data / Report sheets removed entirely. Real state
      files only ever contain the blue sheets (plus Revision History and
      the hidden Inputs lookup) -- the red sheets are something this app
      generates, not something the state sends. This fixture is the
      proof that the engine doesn't secretly depend on those red sheets
      being present.

  tests/fixtures/demo_after.xlsm -- a copy of demo_before.xlsm with 3 new
      rows appended to B-Tests: B-C18, B-C19, B-C20 (continuing the
      existing B-C1..B-C17 concrete-test numbering), each with a
      realistic description, an Approval Agency, and X marks on a
      plausible subset of the 42 stage columns. Used to demonstrate
      "Upload New Version" picking up genuinely new items with zero
      false positives elsewhere.

Two things are checked, both through the real app logic (not just
"the files parse without throwing"):

  1. core.excel_reader.normalize_workbook() on demo_before.xlsm produces
     the exact same 345 named Index values as sample_increment.xlsm --
     proving blue-sheets-only input normalizes identically to the full
     file, which is what every real incoming file will look like.
  2. core.structure_diff.compare_structure() between demo_before.xlsm and
     demo_after.xlsm reports exactly {B-C18, B-C19, B-C20} added, nothing
     removed, and no column anomalies, across all 4 sheets.

Run directly: `python tests/test_blue_only_and_diff_demo.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_reader import normalize_workbook
from core.structure_diff import compare_structure

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
ORIGINAL = os.path.join(FIXTURES_DIR, "sample_increment.xlsm")
BEFORE = os.path.join(FIXTURES_DIR, "demo_before.xlsm")
AFTER = os.path.join(FIXTURES_DIR, "demo_after.xlsm")

EXPECTED_NEW_INDEXES = {"B-C18", "B-C19", "B-C20"}


def _named_indexes(all_data) -> set:
    return {idx for idx in all_data["Index"] if idx}


def main():
    failures = []

    # ------------------------------------------------------------
    # 1. Blue-sheets-only input normalizes identically to the full file
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- demo_before.xlsm (blue sheets only) vs. sample_increment.xlsm")
    print("=" * 70)

    original = normalize_workbook(ORIGINAL)
    before = normalize_workbook(BEFORE)

    original_indexes = _named_indexes(original["all_data"])
    before_indexes = _named_indexes(before["all_data"])

    print(f"sample_increment.xlsm: {len(original['all_data'])} All Data rows, {len(original_indexes)} named indexes")
    print(f"demo_before.xlsm:      {len(before['all_data'])} All Data rows, {len(before_indexes)} named indexes")
    print(f"record_name: {before['project_info']['record_name']!r}")

    if len(before["all_data"]) != 345:
        failures.append(f"demo_before.xlsm should produce 345 All Data rows, got {len(before['all_data'])}")
    if before_indexes != original_indexes:
        only_in_original = original_indexes - before_indexes
        only_in_before = before_indexes - original_indexes
        failures.append(
            "demo_before.xlsm's named indexes should exactly match sample_increment.xlsm's -- "
            f"missing from demo_before: {only_in_original or 'none'}; "
            f"extra in demo_before: {only_in_before or 'none'}"
        )
    if not before["project_info"]["record_name"]:
        failures.append("demo_before.xlsm should still have a readable record_name (A-Project Info survives sheet removal)")

    # ------------------------------------------------------------
    # 2. compare_structure detects exactly the 3 added items
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- compare_structure(demo_before.xlsm, demo_after.xlsm)")
    print("=" * 70)

    diffs = compare_structure(BEFORE, AFTER)
    all_added, all_removed, all_anomalies = [], [], []
    for sheet_name, diff in diffs.items():
        print(f"  {sheet_name}: added={diff.added_indexes} removed={diff.removed_indexes} anomalies={diff.column_anomalies}")
        all_added.extend(diff.added_indexes)
        all_removed.extend(diff.removed_indexes)
        all_anomalies.extend(diff.column_anomalies)

    print(f"\nTotal added: {all_added}")
    print(f"Total removed: {all_removed}")
    print(f"Total column anomalies: {all_anomalies}")

    if set(all_added) != EXPECTED_NEW_INDEXES:
        failures.append(f"expected exactly {EXPECTED_NEW_INDEXES} added, got {set(all_added)}")
    if len(all_added) != len(EXPECTED_NEW_INDEXES):
        failures.append(f"expected no duplicate/extra-sheet added entries, got {all_added}")
    if all_removed:
        failures.append(f"expected zero removed indexes, got {all_removed}")
    if all_anomalies:
        failures.append(f"expected zero column anomalies, got {all_anomalies}")

    # the 3 new rows should only ever show up under B-Tests specifically
    if diffs["B-Tests"].added_indexes != sorted(EXPECTED_NEW_INDEXES) and set(diffs["B-Tests"].added_indexes) != EXPECTED_NEW_INDEXES:
        failures.append(f"expected the 3 new indexes under B-Tests specifically, got {diffs['B-Tests'].added_indexes}")
    for sheet_name in ("C-On-Site Special Inspections", "D-Off-Site Special Inspections", "F-Cons Verif"):
        if diffs[sheet_name].added_indexes:
            failures.append(f"{sheet_name} should show no added indexes, got {diffs[sheet_name].added_indexes}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS")


if __name__ == "__main__":
    main()

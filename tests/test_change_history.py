"""
Tests core/project_store.py's change_history.json persistence (appended by
ui.mock_data.MockDataStore.confirm_update, see core/project_store.py's
module docstring) against three REAL sequential file versions -- never
touches the real ~/AltamiranoBuildersAppData.

Run directly: `python tests/test_change_history.py`
"""

import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
DEMO_BEFORE = os.path.join(FIXTURES, "demo_before.xlsm")
DEMO_AFTER = os.path.join(FIXTURES, "demo_after.xlsm")
DEMO_AFTER_2 = os.path.join(FIXTURES, "demo_after_2.xlsm")
PROJECT_NAME = "Change History Validation"


def _json_roundtrip(value):
    """value_changed_items' "changes" dict holds (old, new) TUPLES --
    change_history.json (like every other file in core/project_store.py)
    is plain JSON, which has no tuple type, so persisted entries come
    back with LISTS in that position instead. Round-tripping the
    in-memory ComparisonResult side through the same json.dumps/loads
    before comparing makes the comparison honest -- "does the persisted
    entry match what the review screen showed" -- rather than a false
    failure on a harmless, unavoidable JSON type coercion.
    """
    return json.loads(json.dumps(value))


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        project_store = ProjectStore(base_dir=Path(tmp) / "data")
        store = MockDataStore(project_store)
        store.add_project(PROJECT_NAME)

        increment = store.add_new_increment(PROJECT_NAME, DEMO_BEFORE)
        inc_name = increment.name

        # ------------------------------------------------------------
        # 1 -- brand-new increment (v1) should NOT write a change_history
        # entry -- only confirm_update() does, and add_new_increment()
        # never calls it.
        # ------------------------------------------------------------
        history_after_create = project_store.load_change_history(PROJECT_NAME, inc_name)
        print(f"1 -- history entries right after creating v1: {len(history_after_create)} (expected 0)")
        if history_after_create != []:
            failures.append(f"1: expected no change_history entries after initial upload, got {history_after_create}")

        # ------------------------------------------------------------
        # 2 -- first confirmed update: v1 -> v2 (demo_before -> demo_after)
        # ------------------------------------------------------------
        result_1 = store.simulate_comparison(PROJECT_NAME, inc_name, DEMO_AFTER)
        store.confirm_update(PROJECT_NAME, inc_name, DEMO_AFTER, result_1)

        history = project_store.load_change_history(PROJECT_NAME, inc_name)
        print(f"\n2 -- history entries after 1st confirmed update: {len(history)} (expected 1)")
        if len(history) != 1:
            failures.append(f"2: expected exactly 1 entry after first confirmed update, got {len(history)}")

        entry_1 = history[0] if history else {}
        print(f"   entry 1: old_version={entry_1.get('old_version')}, new_version={entry_1.get('new_version')}, "
              f"added={len(entry_1.get('added_items', []))}, removed={len(entry_1.get('removed_items', []))}, "
              f"value_changed={len(entry_1.get('value_changed_items', []))}")

        if entry_1.get("old_version") != 1 or entry_1.get("new_version") != 2:
            failures.append(f"2: expected old_version=1/new_version=2, got {entry_1.get('old_version')}/{entry_1.get('new_version')}")
        if entry_1.get("added_items") != _json_roundtrip(result_1.added_items):
            failures.append("2: entry's added_items don't exactly match what simulate_comparison computed for the review screen")
        if entry_1.get("removed_items") != _json_roundtrip(result_1.removed_items):
            failures.append("2: entry's removed_items don't exactly match what simulate_comparison computed")
        if entry_1.get("column_anomalies") != _json_roundtrip(result_1.column_anomalies):
            failures.append("2: entry's column_anomalies don't exactly match what simulate_comparison computed")
        if entry_1.get("value_changed_items") != _json_roundtrip(result_1.value_changed_items):
            failures.append("2: entry's value_changed_items don't exactly match what simulate_comparison computed")
        if "timestamp" not in entry_1 or not entry_1["timestamp"]:
            failures.append("2: entry is missing a timestamp")

        # ------------------------------------------------------------
        # 3 -- second confirmed update: v2 -> v3 (demo_after ->
        # demo_after_2) -- confirms appending, NOT overwriting entry 1
        # ------------------------------------------------------------
        result_2 = store.simulate_comparison(PROJECT_NAME, inc_name, DEMO_AFTER_2)
        store.confirm_update(PROJECT_NAME, inc_name, DEMO_AFTER_2, result_2)

        history = project_store.load_change_history(PROJECT_NAME, inc_name)
        print(f"\n3 -- history entries after 2nd confirmed update: {len(history)} (expected 2)")
        if len(history) != 2:
            failures.append(f"3: expected exactly 2 entries after second confirmed update, got {len(history)}")

        if len(history) >= 1 and history[0] != entry_1:
            failures.append("3: entry 1 changed/was overwritten by the second confirmed update -- history must be append-only")

        entry_2 = history[1] if len(history) > 1 else {}
        print(f"   entry 2: old_version={entry_2.get('old_version')}, new_version={entry_2.get('new_version')}, "
              f"added={len(entry_2.get('added_items', []))}, removed={len(entry_2.get('removed_items', []))}, "
              f"value_changed={len(entry_2.get('value_changed_items', []))}")

        if entry_2.get("old_version") != 2 or entry_2.get("new_version") != 3:
            failures.append(f"3: expected old_version=2/new_version=3, got {entry_2.get('old_version')}/{entry_2.get('new_version')}")
        if entry_2.get("added_items") != _json_roundtrip(result_2.added_items):
            failures.append("3: entry 2's added_items don't exactly match what simulate_comparison computed for the review screen")
        if entry_2.get("value_changed_items") != _json_roundtrip(result_2.value_changed_items):
            failures.append("3: entry 2's value_changed_items don't exactly match what simulate_comparison computed")

        # entries should be genuinely distinguishable (different content),
        # confirming these are two REAL diffs, not two copies of the same one
        if entry_1 == entry_2:
            failures.append("3: entry 1 and entry 2 are identical -- fixture assumption broken, these should be two distinct diffs")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- change_history.json is append-only across confirmed updates (initial upload writes "
              "nothing, each confirm appends exactly one new entry without touching prior ones), and every persisted "
              "entry's content matches exactly what simulate_comparison() computed for that update's review screen")


if __name__ == "__main__":
    main()

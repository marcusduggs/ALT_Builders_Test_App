"""
End-to-end CLOSE_MATCH regression test using two REAL fixture files,
not synthetic strings (contrast with tests/test_increment_matcher.py,
which tests match_increment() directly against hand-written name lists).

demo_inc3_new.xlsm's Record Name is "INC 3 - Sewer Systems".
demo_inc3_closematch.xlsm's Record Name is "INC 3 - Swer Systems" -- a
genuine one-character typo of the same name (edit distance 1), captured
from real use of the unified "Upload File" flow: exactly the scenario
where a typo should prompt "is this an update or a new increment?"
rather than silently creating a duplicate increment.

Run directly: `python tests/test_close_match_real_fixtures.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.increment_matcher import MatchType
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
DEMO_INC3_NEW = os.path.join(FIXTURES, "demo_inc3_new.xlsm")
DEMO_INC3_CLOSEMATCH = os.path.join(FIXTURES, "demo_inc3_closematch.xlsm")
PROJECT_NAME = "Close Match Real Fixture Validation"


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        project_store = ProjectStore(base_dir=Path(tmp) / "data")
        store = MockDataStore(project_store)
        store.add_project(PROJECT_NAME)

        # ------------------------------------------------------------
        # 1 -- demo_inc3_new.xlsm against a project with no matching
        # increment yet: NO_MATCH, then create it as a new increment
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- demo_inc3_new.xlsm: NO_MATCH, then create as new increment")
        print("=" * 70)

        match1 = store.match_upload(PROJECT_NAME, DEMO_INC3_NEW)
        print(f"match_upload result: {match1}")
        if match1.match_type != MatchType.NO_MATCH:
            failures.append(f"1: expected NO_MATCH for demo_inc3_new.xlsm against an empty project, got {match1.match_type}")

        increment = store.add_new_increment(PROJECT_NAME, DEMO_INC3_NEW)
        print(f"created increment: {increment.name!r}, version {increment.version}")
        if increment.name != "INC 3 - Sewer Systems":
            failures.append(f"1: expected increment name 'INC 3 - Sewer Systems', got {increment.name!r}")

        # ------------------------------------------------------------
        # 2 -- demo_inc3_closematch.xlsm ("INC 3 - Swer Systems", a
        # one-character typo) against that same project: CLOSE_MATCH
        # specifically -- not EXACT_MATCH, not NO_MATCH
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- demo_inc3_closematch.xlsm: CLOSE_MATCH against the increment above")
        print("=" * 70)

        match2 = store.match_upload(PROJECT_NAME, DEMO_INC3_CLOSEMATCH)
        print(f"match_upload result: {match2}")
        if match2.match_type != MatchType.CLOSE_MATCH:
            failures.append(
                f"2: expected CLOSE_MATCH for the one-character-typo Record Name, got {match2.match_type}"
            )
        if match2.matched_increment_name != "INC 3 - Sewer Systems":
            failures.append(
                f"2: expected CLOSE_MATCH to point at 'INC 3 - Sewer Systems', got {match2.matched_increment_name!r}"
            )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- real-fixture NO_MATCH -> create -> CLOSE_MATCH flow works end to end")


if __name__ == "__main__":
    main()

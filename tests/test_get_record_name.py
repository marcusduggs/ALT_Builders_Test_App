"""
Tests core/excel_reader.py's get_record_name(): the label-search-based
"Record Name (Scope of Project)" lookup that core/increment_matcher.py
uses as the identity source of truth for the unified upload flow.

sample_increment.xlsm, demo_before.xlsm, and demo_after.xlsm are all
derived from the same source project, so all three should extract the
exact same record name.

Run directly: `python tests/test_get_record_name.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_reader import get_record_name

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
EXPECTED = "INC 2 - Foundation and Underground Utilities"


def main():
    failures = []

    for fixture_name in ("sample_increment.xlsm", "demo_before.xlsm", "demo_after.xlsm"):
        path = os.path.join(FIXTURES, fixture_name)
        record_name = get_record_name(path)
        print(f"{fixture_name}: {record_name!r}")
        if record_name != EXPECTED:
            failures.append(f"{fixture_name}: expected {EXPECTED!r}, got {record_name!r}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- get_record_name() matches across all 3 fixtures derived from the same source")


if __name__ == "__main__":
    main()

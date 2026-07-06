"""
Tests core/increment_matcher.py's match_increment() against the three
outcomes the unified upload flow routes on: EXACT_MATCH, CLOSE_MATCH,
and NO_MATCH.

Run directly: `python tests/test_increment_matcher.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.increment_matcher import MatchType, match_increment

EXISTING = ["INC 2 - Foundation and Underground Utilities", "INC 1 - Site Grading"]


def main():
    failures = []

    # ------------------------------------------------------------
    # 1 -- exact match (case/whitespace aside)
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- exact match")
    print("=" * 70)
    result = match_increment(EXISTING, "INC 2 - Foundation and Underground Utilities")
    print(result)
    if result.match_type != MatchType.EXACT_MATCH:
        failures.append(f"1: expected EXACT_MATCH, got {result.match_type}")
    if result.matched_increment_name != "INC 2 - Foundation and Underground Utilities":
        failures.append(f"1: expected matched name 'INC 2 - Foundation and Underground Utilities', got {result.matched_increment_name!r}")

    # case/whitespace should still count as exact -- get_record_name()
    # already normalizes whitespace at the source, so this is the "some
    # other path produced slightly different casing/spacing" case, not
    # something a human should be asked about.
    result_casing = match_increment(EXISTING, "  inc 2 -   foundation and underground utilities  ")
    print(result_casing)
    if result_casing.match_type != MatchType.EXACT_MATCH:
        failures.append(f"1: case/whitespace-only difference should still be EXACT_MATCH, got {result_casing.match_type}")

    # ------------------------------------------------------------
    # 2 -- no match at all
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- no match")
    print("=" * 70)
    result = match_increment(EXISTING, "INC 9 - Mechanical Penthouse Equipment")
    print(result)
    if result.match_type != MatchType.NO_MATCH:
        failures.append(f"2: expected NO_MATCH, got {result.match_type}")
    if result.matched_increment_name is not None:
        failures.append(f"2: NO_MATCH should not carry a matched_increment_name, got {result.matched_increment_name!r}")

    # an empty project (no increments yet) must also be NO_MATCH, not an error
    result_empty = match_increment([], "INC 2 - Foundation and Underground Utilities")
    print(result_empty)
    if result_empty.match_type != MatchType.NO_MATCH:
        failures.append(f"2: a project with zero increments should be NO_MATCH, got {result_empty.match_type}")

    # ------------------------------------------------------------
    # 3 -- close match: same name, different dash character
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- close match (different dash character)")
    print("=" * 70)
    result = match_increment(EXISTING, "INC 2 – Foundation and Underground Utilities")  # en dash
    print(result)
    if result.match_type != MatchType.CLOSE_MATCH:
        failures.append(f"3: expected CLOSE_MATCH for a different dash character, got {result.match_type}")
    if result.matched_increment_name != "INC 2 - Foundation and Underground Utilities":
        failures.append(f"3: expected closest match 'INC 2 - Foundation and Underground Utilities', got {result.matched_increment_name!r}")

    # a small typo should also land as CLOSE_MATCH via edit distance,
    # not NO_MATCH
    result_typo = match_increment(EXISTING, "INC 2 - Foundation and Undergound Utilities")  # "Undergound"
    print(result_typo)
    if result_typo.match_type != MatchType.CLOSE_MATCH:
        failures.append(f"3: expected CLOSE_MATCH for a small typo, got {result_typo.match_type}")

    # a name close to one existing increment shouldn't get confused with
    # the OTHER existing increment
    if result.matched_increment_name == "INC 1 - Site Grading":
        failures.append("3: matched the wrong existing increment")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- EXACT_MATCH, CLOSE_MATCH, and NO_MATCH all classified correctly")


if __name__ == "__main__":
    main()

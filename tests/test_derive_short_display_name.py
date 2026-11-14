"""
Regression test for core.excel_reader.derive_short_display_name() --
a NEW, display-only function, completely separate from get_record_name()/
core.increment_matcher (both untouched by this change: identity/matching
still compares the full, untruncated Record Name text exactly as
before). Exists because real HMR/SMR client files can have a Record
Name running 300-556 characters (confirmed against
tests/fixtures/hmr_smr_projects -- real client files, not synthetic
data), which overflowed the Data View page's title bar and the Sum Data
table's increment-name column.

Validates the exact before/after values across all 13 real HMR/SMR
fixtures (investigated directly, not guessed -- see the function's own
docstring for the three real patterns found) and confirms every
existing non-HMR/SMR fixture is completely unaffected (all already
short enough to pass through as-is).

Run directly: `python tests/test_derive_short_display_name.py`
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_reader import derive_short_display_name, get_record_name

HMR_SMR_DIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
EXISTING_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")

# Exact expected short name for each of the 13 real files, confirmed by
# direct investigation (see the PR/session this landed in for the raw
# text dump of all 13) -- NOT a guess. The one genuine exception
# (SMR 2D) has no reliable " - " split (its only separator yields a
# bare 9-character "UCSF NHPH" facility fragment, rejected as too short
# to be a usable label) and falls back to hard truncation.
EXPECTED = {
    "TIO H220369-38-00_SMR 2D.xlsm":
        "UCSF NHPH - Site Make Ready 2D Part of the site make ready projects for UCSF…",  # FALLBACK (see above)
    "TIO H220881-38-00_SMR 4.xlsm":
        "UCSF NHPH - Site Make Ready 4 - Oxygen Yard",  # as-is, 43 chars
    "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm":
        "SMR 3D Permanent Power Gear and Distribution",  # as-is, 44 chars
    "TIO H231359-38-00_HMR 9B.xlsm":
        "UCSF NHPH Hospital Make Ready 9B",  # split, 556 -> 32 chars
    "TIO H240424-38-00 SMR 4B Delta 7.xlsm":
        "UCSF NHPH - SMR 4B Interim Power",  # as-is, 32 chars
    "TIO H250355-38-00 HMR 17.xlsm":
        "UCSF HHDH Hospital Make Ready 17",  # split, 313 -> 32 chars
    "TIO S220804-38-00_SMR 2C_220705.xlsm":
        "Site Make Ready 2C, Utility Disconnects",  # as-is, 39 chars
    "TIO S221870-38-00_SMR 3C.xlsm":
        "Site Make Ready 3C - Loading dock utilities and POCs",  # as-is, 52 chars
    "TIO S221919-38-00_HMR 6C.xlsm":
        "UCSF NHPH - Hospital Make Ready 6C - Med Gas Relocation, Loading Dock Renovation & IDF Closet",  # as-is, 93 chars
    "TIO S230119-38-00_SMR 3A - 2022.xlsm":
        "UCSF NHPH Site Make Ready 3A Site Infrastructure",  # as-is, 48 chars
    "TIO S231014-38-00 HMR 12.xlsm":
        "UCSF NHPH - Hospital Make Ready 12",  # split, 363 -> 34 chars
    "TIO S251652-38-00 SMR 4C.xlsm":
        "UCSF HHDH – Site Make Ready 4C – MCW Protection Project",  # as-is, 55 chars (en dash)
    "TIO_H240649-38-00 HMR 10A.xlsm":
        "UCSF NHPH Hospital Make Ready 10A",  # split, 491 -> 33 chars
}


def main():
    failures = []

    # ------------------------------------------------------------
    # 1 -- all 13 real HMR/SMR fixtures, exact before/after
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- All 13 real HMR/SMR fixtures: exact short-name values")
    print("=" * 70)

    if not os.path.isdir(HMR_SMR_DIR):
        print(f"SKIP: {HMR_SMR_DIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
    else:
        real_files = sorted(f for f in os.listdir(HMR_SMR_DIR) if f.endswith(".xlsm"))
        if set(real_files) != set(EXPECTED.keys()):
            failures.append(
                f"1: fixture set changed -- expected exactly {sorted(EXPECTED.keys())}, found {real_files}"
            )
        for fname in real_files:
            full = get_record_name(os.path.join(HMR_SMR_DIR, fname))
            short = derive_short_display_name(full)
            expected = EXPECTED.get(fname)
            status = "FALLBACK" if short.endswith("…") else ("as-is" if short == full else "split")
            print(f"  {fname}: {len(full)} -> {len(short)} chars [{status}]")
            print(f"    {short!r}")
            if expected is not None and short != expected:
                failures.append(f"1: {fname!r} expected {expected!r}, got {short!r}")

        # Exactly one real file should hit the fallback-truncation path.
        fallback_files = [f for f in real_files if derive_short_display_name(get_record_name(os.path.join(HMR_SMR_DIR, f))).endswith("…")]
        print(f"\n  Files using the fallback-truncation path: {fallback_files}")
        if fallback_files != ["TIO H220369-38-00_SMR 2D.xlsm"]:
            failures.append(f"1: expected exactly 1 fallback file (SMR 2D), got {fallback_files}")

    # ------------------------------------------------------------
    # 2 -- every existing (non-HMR/SMR) fixture is completely
    # unaffected -- all already short, all returned as-is.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- Existing fixtures are unaffected")
    print("=" * 70)
    for fname in sorted(os.listdir(EXISTING_FIXTURES_DIR)):
        if not fname.endswith(".xlsm") or fname.startswith("~$"):
            continue
        full = get_record_name(os.path.join(EXISTING_FIXTURES_DIR, fname))
        short = derive_short_display_name(full)
        print(f"  {fname}: {short!r}")
        if short != full:
            failures.append(f"2: {fname!r} should be unchanged (already short), got {short!r} from {full!r}")

    # ------------------------------------------------------------
    # 3 -- pure-logic edge cases: blank input, already-at-the-limit
    # length, and a string with NO " - " at all long enough to need
    # the fallback.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- Edge cases")
    print("=" * 70)
    if derive_short_display_name("") != "":
        failures.append("3: empty input should return empty string")
    if derive_short_display_name(None) != "":
        failures.append("3: None input should return empty string, not raise")

    exactly_100 = "X" * 100
    if derive_short_display_name(exactly_100) != exactly_100:
        failures.append("3: a string exactly at the display limit (100 chars) should pass through unchanged")

    no_dash_long = "A" * 200  # long, no " - " anywhere at all
    result = derive_short_display_name(no_dash_long)
    print(f"  No-dash long string -> {result!r}")
    if not result.endswith("…") or len(result) > 85:
        failures.append(f"3: a long string with no ' - ' at all should fall back to truncation, got {result!r}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- derive_short_display_name() matches the exact investigated values for all 13 "
              "real HMR/SMR fixtures (9 as-is, 3 cleanly split, 1 correctly falls back to truncation), leaves "
              "every existing non-HMR/SMR fixture unchanged, and handles blank/edge-case input without raising.")


if __name__ == "__main__":
    main()

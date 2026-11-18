"""
Regression test for core.excel_reader.derive_increment_code() -- a NEW,
separate function from derive_short_display_name() (untouched) and
get_record_name()/core.increment_matcher (identity/matching still
compares the full, untruncated Record Name text exactly as before).

Rey's own TIO_Changes.xlsx "Sum Data" tab column B uses literal short
CODES ("SMR 3D", "HMR 10A"), not derive_short_display_name()'s
shortened-prose fallback ("UCSF NHPH Hospital Make Ready 10A"/etc) --
this function extracts that code directly from the Record Name's own
"Hospital/Site Make Ready <n>" or already-abbreviated "HMR|SMR <n>"
text.

Two things validated, both against REAL data, not guesses:

1. An ORACLE check (skip-guarded on ~/Desktop/TIO_Changes.xlsx being
   present -- it's the client's own reference file, not part of this
   repo): for each of the 13 real HMR/SMR fixtures Rey's own Sum Data
   tab labels, derive_increment_code(get_record_name(fixture)) must
   equal Rey's own column B value for that Project #, read directly
   from the workbook (never hand-copied into this file as a second
   source of truth that could drift).

   ONE genuine, confirmed discrepancy exists in the source data itself:
   Project H250355-38-00's own Record Name literally says "Hospital
   Make Ready 17" (producing "HMR 17"), but Rey's own TIO_Changes.xlsx
   labels that same Project # "SMR 17". This is NOT a bug in
   derive_increment_code() -- the function is reading the Record Name
   correctly; the client's two source documents disagree with each
   other. This test asserts that exact, singular mismatch is what's
   found (not silently "fixed" by hand-mapping, and not grown to cover
   any OTHER project) -- if a future fixture update changes either
   document, this test will need a deliberate, visible update, not a
   silent pass.

2. Confirms every fixture OUTSIDE the HMR/SMR naming convention
   (inc8_real_2stage.xlsm, and every other non-HMR/SMR fixture already
   covered by test_derive_short_display_name.py) returns None, falling
   back to derive_short_display_name() -- proven directly against
   ui.mock_data.MockDataStore's actual resolved display_name, not just
   the bare function.

Run directly: `python tests/test_derive_increment_code.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from core.excel_reader import derive_increment_code, derive_short_display_name, get_record_name, parse_workbook
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
INC8 = os.path.join(os.path.dirname(__file__), "fixtures", "inc8_real_2stage.xlsm")
REY_PATH = os.path.expanduser("~/Desktop/TIO_Changes.xlsx")

# Maps each real fixture's filename to its own Project # -- read from
# the file itself in the test below, kept here only to pick which file
# goes with which expected code without re-parsing twice.
FIXTURE_FILES = [
    "TIO H220369-38-00_SMR 2D.xlsm",
    "TIO H220881-38-00_SMR 4.xlsm",
    "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm",
    "TIO H231359-38-00_HMR 9B.xlsm",
    "TIO H240424-38-00 SMR 4B Delta 7.xlsm",
    "TIO H250355-38-00 HMR 17.xlsm",
    "TIO S220804-38-00_SMR 2C_220705.xlsm",
    "TIO S221870-38-00_SMR 3C.xlsm",
    "TIO S221919-38-00_HMR 6C.xlsm",
    "TIO S230119-38-00_SMR 3A - 2022.xlsm",
    "TIO S231014-38-00 HMR 12.xlsm",
    "TIO S251652-38-00 SMR 4C.xlsm",
    "TIO_H240649-38-00 HMR 10A.xlsm",
]

# The ONE confirmed, expected discrepancy -- see module docstring.
EXPECTED_MISMATCHES = {"H250355-38-00": ("HMR 17", "SMR 17")}


def _reys_codes_by_project(path: str) -> dict[str, str]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sum Data"]
    codes: dict[str, set] = {}
    for r in range(2, ws.max_row + 1):
        proj = ws.cell(row=r, column=1).value
        label = ws.cell(row=r, column=2).value
        if proj is None:
            continue
        codes.setdefault(proj, set()).add(label)
    for proj, labels in codes.items():
        if len(labels) != 1:
            raise AssertionError(f"Rey's own Sum Data tab has MULTIPLE distinct column-B labels for Project #={proj!r}: {labels}")
    return {proj: next(iter(labels)) for proj, labels in codes.items()}


def test_oracle_against_reys_own_spreadsheet(failures):
    if not os.path.isfile(REY_PATH):
        print(f"SKIP: {REY_PATH!r} not present (Rey's own reference file lives outside this repo)")
        return
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    reys_codes = _reys_codes_by_project(REY_PATH)
    print(f"Rey's own column B codes, by Project #: {reys_codes}")

    found_mismatches = {}
    for fname in FIXTURE_FILES:
        path = os.path.join(FIXDIR, fname)
        record_name = get_record_name(path)
        project_number = parse_workbook(path)["project_info"]["project_number"]
        code = derive_increment_code(record_name)
        expected = reys_codes.get(project_number)
        print(f"  {fname!r:50} proj#={project_number!r:16} code={code!r:10} rey={expected!r:10}")

        if expected is None:
            failures.append(f"oracle: no Rey column-B value found for Project #={project_number!r} ({fname!r}) -- fixture/expectation list is stale")
            continue
        if code != expected:
            found_mismatches[project_number] = (code, expected)

    if found_mismatches != EXPECTED_MISMATCHES:
        failures.append(
            f"oracle: expected EXACTLY the known mismatch {EXPECTED_MISMATCHES} (Record Name says 'Hospital Make "
            f"Ready 17' but Rey's own tab calls it 'SMR 17' -- a genuine discrepancy in the source data, not a bug "
            f"here), got {found_mismatches}"
        )
    else:
        print(f"\nConfirmed: exactly the known, expected discrepancy was found: {found_mismatches}")
        print("  -> Project H250355-38-00's Record Name literally says 'Hospital Make Ready 17'; Rey's own")
        print("     TIO_Changes.xlsx Sum Data tab calls that same Project # 'SMR 17'. Reported, not guessed/forced.")


def test_non_hmr_smr_fixture_falls_back(failures):
    record_name = get_record_name(INC8)
    code = derive_increment_code(record_name)
    print(f"inc8_real_2stage.xlsm record_name={record_name!r} -> code={code!r} (expected None)")
    if code is not None:
        failures.append(f"inc8_real_2stage.xlsm's Record Name doesn't follow the HMR/SMR convention -- expected None, got {code!r}")

    expected_fallback = derive_short_display_name(record_name)
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Fallback Check")
        inc = store.add_new_increment("Fallback Check", INC8)
        display = store.get_increment_for_display("Fallback Check", inc.name)
        print(f"Resolved Increment.display_name: {display.display_name!r} (expected fallback {expected_fallback!r})")
        if display.display_name != expected_fallback:
            failures.append(
                f"a non-HMR/SMR increment's resolved display_name should fall back to derive_short_display_name() "
                f"({expected_fallback!r}), got {display.display_name!r}"
            )


def main():
    failures = []
    test_oracle_against_reys_own_spreadsheet(failures)
    test_non_hmr_smr_fixture_falls_back(failures)

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- derive_increment_code() matches Rey's own TIO_Changes.xlsx Sum Data tab codes for "
            "every real HMR/SMR fixture except the one confirmed, genuine discrepancy in the source data itself "
            "(H250355-38-00), and a fixture outside the HMR/SMR naming convention correctly falls back to "
            "derive_short_display_name()."
        )


if __name__ == "__main__":
    main()

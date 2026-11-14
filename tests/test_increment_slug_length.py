"""
Regression test for Bug A: core.project_store._slugify() had no length
cap, and increment identity is derived from core.excel_reader.
get_record_name()'s FULL free-text Record Name -- for most real files a
short title, but for several real client files (see
tests/fixtures/hmr_smr_projects/ -- real HMR/SMR TIO files, not
synthetic data) the entire "Scope of Project" paragraph, 300-550+
characters. The raw slug of that text alone exceeded the OS's 255-byte
single-path-component limit, crashing MockDataStore.add_new_increment()
with OSError: File name too long on 5 of 13 real fixture files.

Fixed by core.project_store._slugify_increment(): a long slug is
truncated (capped at _MAX_INCREMENT_SLUG_LENGTH) and disambiguated with
the file's own Project # (short, unique per real project) or a content
hash -- core.increment_matcher's identity matching is UNCHANGED (still
compares the full Record Name text; only the on-disk directory name is
capped).

Run directly: `python tests/test_increment_slug_length.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import excel_reader
from core.project_store import ProjectStore, _MAX_INCREMENT_SLUG_LENGTH, _slugify, _slugify_increment
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Slug Length Regression"

# These 5 real files are the exact ones that crashed before the fix --
# each has a Record Name long enough that _slugify() alone exceeds 255
# bytes (confirmed during investigation; re-checked below as a fixture
# assumption guard, not hardcoded as "the answer").
PREVIOUSLY_CRASHING_FILES = [
    "TIO H220369-38-00_SMR 2D.xlsm",
    "TIO H231359-38-00_HMR 9B.xlsm",
    "TIO H250355-38-00 HMR 17.xlsm",
    "TIO S231014-38-00 HMR 12.xlsm",
    "TIO_H240649-38-00 HMR 10A.xlsm",
]


def main():
    failures = []

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    # ------------------------------------------------------------
    # 1 -- every file known to have previously crashed now produces a
    # slug safely under the OS limit, via the real add_new_increment()
    # path (not just calling _slugify_increment() in isolation).
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- Previously-crashing real files now succeed end to end")
    print("=" * 70)

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)

        for fname in PREVIOUSLY_CRASHING_FILES:
            path = os.path.join(FIXDIR, fname)
            if not os.path.exists(path):
                failures.append(f"1: expected fixture {fname!r} not found at {path!r}")
                continue
            record_name = excel_reader.get_record_name(path)
            old_slug = _slugify(record_name)
            if len(old_slug.encode()) <= 255:
                failures.append(
                    f"1: fixture assumption broken -- {fname!r}'s OLD slug is only "
                    f"{len(old_slug.encode())} bytes, not actually over the 255-byte limit this "
                    "test is meant to guard against"
                )
            try:
                increment = store.add_new_increment(PROJECT_NAME, path)
            except OSError as e:
                failures.append(f"1: {fname!r} still raises OSError: {e}")
                continue
            print(f"  OK: {fname!r} (old slug {len(old_slug.encode())} bytes) -> increment created, "
                  f"version {increment.version}")

    # ------------------------------------------------------------
    # 2 -- ALL 13 real fixture files succeed via add_new_increment(),
    # not just the 5 known-bad ones (confirms no regression for the 8
    # that already worked).
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- All 13 real HMR/SMR fixture files succeed via add_new_increment()")
    print("=" * 70)

    all_files = sorted(f for f in os.listdir(FIXDIR) if f.endswith(".xlsm"))
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        succeeded = 0
        for fname in all_files:
            path = os.path.join(FIXDIR, fname)
            try:
                store.add_new_increment(PROJECT_NAME, path)
                succeeded += 1
            except Exception as e:
                failures.append(f"2: {fname!r} raised {type(e).__name__}: {e}")
        print(f"  {succeeded}/{len(all_files)} real files succeeded")
        if succeeded != len(all_files):
            failures.append(f"2: expected all {len(all_files)} real files to succeed, only {succeeded} did")

    # ------------------------------------------------------------
    # 3 -- a short, already-safe name is slugified IDENTICALLY to
    # before -- this fix must not change on-disk slugs for the common
    # case (every increment already in use anywhere).
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- Short names are completely unaffected")
    print("=" * 70)
    short_name = "INC 2 - Foundation and Underground Utilities"
    unaffected = _slugify_increment(short_name, project_number="S999999-38-00")
    expected = _slugify(short_name)
    print(f"  _slugify_increment(short_name, project_number=...) = {unaffected!r}")
    print(f"  _slugify(short_name)                                = {expected!r}")
    if unaffected != expected:
        failures.append(f"3: a short name's slug should be unaffected by this fix, got {unaffected!r} vs {expected!r}")

    # ------------------------------------------------------------
    # 4 -- disambiguation: two DIFFERENT long names that happen to
    # share the same first _MAX_INCREMENT_SLUG_LENGTH characters must
    # still produce DIFFERENT slugs (no real pair of the 13 fixtures
    # happens to collide this way, so this one case is a constructed
    # edge case, not drawn from a real file -- it locks in the
    # disambiguation guarantee itself, which the 13 real files alone
    # don't happen to exercise).
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("4 -- Two long names sharing a common prefix still get distinct slugs")
    print("=" * 70)
    common_prefix = "UCSF NHPH - Hospital Make Ready Project " + ("X" * 200)
    name_a = common_prefix + " - Scope A, a very long unique tail describing project A specifically."
    name_b = common_prefix + " - Scope B, a very long unique tail describing project B specifically."
    slug_a = _slugify_increment(name_a, project_number="S100000-38-00")
    slug_b = _slugify_increment(name_b, project_number="S200000-38-00")
    print(f"  slug_a: {slug_a}")
    print(f"  slug_b: {slug_b}")
    if slug_a == slug_b:
        failures.append(f"4: two different long names with distinct Project #s produced the SAME slug: {slug_a!r}")
    if len(slug_a.encode()) > _MAX_INCREMENT_SLUG_LENGTH + 20 or len(slug_b.encode()) > _MAX_INCREMENT_SLUG_LENGTH + 20:
        failures.append(f"4: disambiguated slugs should stay close to the {_MAX_INCREMENT_SLUG_LENGTH}-char cap")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- all 13 real HMR/SMR fixture files (including the 5 that previously crashed "
              "with OSError: File name too long) now succeed via add_new_increment(); short names are "
              "completely unaffected; two long names sharing a common prefix still get distinct, "
              "disambiguated slugs.")


if __name__ == "__main__":
    main()

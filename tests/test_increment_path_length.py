"""
Regression test for a real client Windows path-length failure: on
v0.14.0, uploading TIO H231359-38-00_HMR 9B.xlsm (one of the 13 real
HMR/SMR fixtures) failed with:

    [Errno 2] No such file or directory: 'C:\\Users\\rey\\OneDrive\\
    1. Altamirano Builders\\0. AB TIO APP\\App Files\\
    AltamiranoBuildersAppData\\projects\\hmr-projects\\increments\\
    ucsf-nhph-hospital-make-ready-9b-long-hospital-shear-wall-
    penetrations-and-retrofit-connections-reno-h231359-38-00\\
    increment.json.tmp'

shown to the user as "Could Not Read File: This file couldn't be
parsed" -- actively misleading, since the file parsed fine; SAVING it
is what failed.

Confirmed directly (see git history) with ntpath/PureWindowsPath,
before any fix: that path is 261 characters -- 2 over Windows' real
MAX_PATH failure point (not the commonly-quoted 260; creating the
242-character DIRECTORY alone succeeded, only the 261-character FILE
inside it failed). Root cause was TWO compounding things, not one:
  1. core.project_store._MAX_INCREMENT_SLUG_LENGTH (100) was sized for
     macOS's 255-byte single-PATH-COMPONENT limit, never accounting
     for the base data directory's own depth adding to a SEPARATE,
     whole-path Windows limit.
  2. The old _slugify_increment() truncated to that 100-character cap
     and then APPENDED "-{disambiguator}" (e.g. a 13-character real
     Project #) ON TOP of it -- so the real result could run up to
     ~114 characters, not 100, even though the cap's own name implied
     the opposite.

Both are fixed: _increment_slug_max_length_for_prefix() computes a
REAL, per-installation budget from the actual base data directory's
length, and the rewritten _slugify_increment() carves the disambiguator
OUT of max_length instead of appending it on top, so the final result
is a true hard cap.

Covers:
  1. The OLD, pre-fix numbers reproduced exactly (261/242/114) --
     confirms this test would have caught the real bug, not just that
     the new numbers happen to look fine in isolation.
  2. All 13 real HMR/SMR fixtures, building every path the app would
     actually create under the CLIENT'S OWN real base data directory
     (via PureWindowsPath, so this is deterministic on any host OS,
     not just Windows) -- every one stays at or under 240 characters,
     including the longest file the app writes inside each increment
     folder, not just the folder name itself.
  3. A deliberately LONGER username and project name than the
     client's own real one, confirming the fix scales down further
     rather than just barely covering this one specific real case.

Run directly: `python tests/test_increment_path_length.py`
"""

import glob
import os
import sys
from pathlib import PureWindowsPath

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import excel_reader
from core.project_store import (
    _MAX_INCREMENT_SLUG_LENGTH,
    _MIN_INCREMENT_SLUG_LENGTH,
    _TARGET_MAX_PATH_LENGTH,
    _increment_slug_max_length_for_prefix,
    _slugify,
    _slugify_increment,
)

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")

# The client's own real base data directory and project slug, exactly
# as they appeared in the failing path.
CLIENT_BASE_DIR = (
    r"C:\Users\rey\OneDrive\1. Altamirano Builders\0. AB TIO APP\App Files\AltamiranoBuildersAppData"
)
CLIENT_PROJECT_SLUG = "hmr-projects"

# Every relative path the app writes INSIDE an increment folder -- see
# core.project_store._LONGEST_RELATIVE_FILE_PATH_LENGTH's own docstring
# for how each was confirmed (increment.json/status.json/comments.json/
# change_history.json, each via _write_json()'s own ".tmp" atomic-write
# sibling, plus each stored version file under "files/" as
# "v<N>_<YYYY-MM-DD><ext>" -- never the original uploaded filename).
RELATIVE_FILE_PATHS = [
    "increment.json", "increment.json.tmp",
    "status.json", "status.json.tmp",
    "comments.json", "comments.json.tmp",
    "change_history.json", "change_history.json.tmp",
    "files\\v1_2026-11-16.xlsm",
    "files\\v10_2026-11-16.xlsm",
    "files\\v100_2026-11-16.xlsm",
]


def main():
    failures = []

    # ------------------------------------------------------------
    # 1 -- Reproduce the OLD, pre-fix numbers exactly, confirmed via
    # PureWindowsPath -- this is what actually failed on the client's
    # machine, and what this whole fix is a response to.
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- Reproducing the exact pre-fix client failure")
    print("=" * 70)

    old_slug_prefix = _slugify(
        "UCSF NHPH Hospital Make Ready 9B - Long Hospital Shear Wall Penetrations and Retrofit, "
        "Connections Renovation, and Below-grade Waterproofing Installation"
    )[:100]
    old_slug = f"{old_slug_prefix}-h231359-38-00"  # the OLD behavior: cap, THEN append the disambiguator on top
    print(f"  Old-style slug: {old_slug!r} ({len(old_slug)} chars)")
    if len(old_slug) != 114:
        failures.append(f"1: fixture assumption broken -- expected the old-style slug to be 114 chars, got {len(old_slug)}")

    increments_dir = PureWindowsPath(CLIENT_BASE_DIR, "projects", CLIENT_PROJECT_SLUG, "increments")
    old_dir_path = increments_dir / old_slug
    old_file_path = old_dir_path / "increment.json.tmp"
    print(f"  Directory-only length: {len(str(old_dir_path))} (client reported this SUCCEEDED)")
    print(f"  Full file path length: {len(str(old_file_path))} (client reported this FAILED)")
    if len(str(old_dir_path)) != 242:
        failures.append(f"1: fixture assumption broken -- expected the old directory path to be 242 chars, got {len(str(old_dir_path))}")
    if len(str(old_file_path)) != 261:
        failures.append(f"1: fixture assumption broken -- expected the old full file path to be 261 chars, got {len(str(old_file_path))}")

    # ------------------------------------------------------------
    # 2 -- All 13 real fixtures, under the client's OWN real base
    # directory, every path the app would create at or under 240 chars.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- All 13 real fixtures, client's real base directory")
    print("=" * 70)

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    prefix_length = len(str(increments_dir))
    max_slug_length = _increment_slug_max_length_for_prefix(prefix_length)
    print(f"  increments_dir prefix length: {prefix_length} (matches the client's own reported '127')")
    print(f"  computed max_slug_length: {max_slug_length}")

    files = sorted(glob.glob(os.path.join(FIXDIR, "*")))
    taken = set()
    for f in files:
        name = excel_reader.get_record_name(f)
        wb = excel_reader.open_workbook(f)
        project_number = excel_reader.parse_project_info(wb["A-Project Info"])["project_number"]

        slug = _slugify_increment(name, project_number, max_slug_length)
        n = 2
        base_slug = slug
        while slug in taken:  # mirrors _unique_slug()'s own collision handling
            slug = f"{base_slug}-{n}"
            n += 1
        taken.add(slug)

        increment_dir = increments_dir / slug
        worst_path = max((increment_dir / rel for rel in RELATIVE_FILE_PATHS), key=lambda p: len(str(p)))
        print(f"  {os.path.basename(f)[:45]:45s} slug_len={len(slug):3d} worst_path_len={len(str(worst_path))}")
        if len(str(worst_path)) > _TARGET_MAX_PATH_LENGTH:
            failures.append(
                f"2: {os.path.basename(f)}: worst-case path {len(str(worst_path))} chars exceeds the "
                f"{_TARGET_MAX_PATH_LENGTH}-char target ({worst_path})"
            )

    # ------------------------------------------------------------
    # 3a -- A REALISTICALLY longer username and project name than the
    # client's own real one (longer, not pathological) -- confirms the
    # fix scales down further, not just barely covers the one real
    # case on hand. Must stay within budget, same as section 2.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3a -- A realistically longer username and project name")
    print("=" * 70)

    realistic_longer_base = (
        r"C:\Users\alexandra.rodriguez\OneDrive - Altamirano Builders"
        r"\1. Altamirano Builders\0. AB TIO APP\App Files\AltamiranoBuildersAppData"
    )
    realistic_longer_project_slug = "hmr-smr-projects"
    realistic_increments_dir = PureWindowsPath(realistic_longer_base, "projects", realistic_longer_project_slug, "increments")
    realistic_prefix_length = len(str(realistic_increments_dir))
    realistic_max_slug_length = _increment_slug_max_length_for_prefix(realistic_prefix_length)
    print(f"  increments_dir prefix length: {realistic_prefix_length} (vs. the client's own real 127)")
    print(f"  computed max_slug_length: {realistic_max_slug_length}")
    if realistic_prefix_length <= prefix_length:
        failures.append(
            f"3a: fixture assumption broken -- expected this scenario's prefix ({realistic_prefix_length}) "
            f"to be longer than the client's own real one ({prefix_length})"
        )

    longest_name = max((excel_reader.get_record_name(f) for f in files), key=len)
    slug = _slugify_increment(longest_name, "H999999-38-00", realistic_max_slug_length)
    worst_path = max(
        (realistic_increments_dir / slug / rel for rel in RELATIVE_FILE_PATHS),
        key=lambda p: len(str(p)),
    )
    print(f"  slug_len={len(slug)} worst_path_len={len(str(worst_path))}")
    if len(str(worst_path)) > _TARGET_MAX_PATH_LENGTH:
        failures.append(
            f"3a: realistically-longer-username/project-name case: worst-case path "
            f"{len(str(worst_path))} chars exceeds the {_TARGET_MAX_PATH_LENGTH}-char target ({worst_path})"
        )

    # ------------------------------------------------------------
    # 3b -- An EXTREME, deliberately pathological base directory --
    # NOT asserted to stay under 240 (that would require shrinking the
    # slug below the point of still being a recognizable fragment of
    # the real name -- see _MIN_INCREMENT_SLUG_LENGTH's own docstring).
    # Confirms instead that the floor itself kicks in correctly (a
    # real, positive, usable length) and that the result never raises
    # or returns something nonsensical -- an extreme case fails
    # loudly later (via the "folder path may be too long" message),
    # never silently or by crashing here.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3b -- An extreme, pathologically deep base directory (floor behavior)")
    print("=" * 70)

    extreme_base = (
        r"C:\Users\alexandrathemightyandpowerful\OneDrive - Altamirano Builders Construction Group LLC"
        r"\1. Altamirano Builders Shared Documents\0. AB TIO Compliance Application\App Files"
        r"\AltamiranoBuildersApplicationData"
    )
    extreme_project_slug = "hospital-make-ready-and-site-make-ready-combined-projects-for-ucsf-nhph-2026"
    extreme_increments_dir = PureWindowsPath(extreme_base, "projects", extreme_project_slug, "increments")
    extreme_prefix_length = len(str(extreme_increments_dir))
    extreme_max_slug_length = _increment_slug_max_length_for_prefix(extreme_prefix_length)
    print(f"  increments_dir prefix length: {extreme_prefix_length} (deliberately pathological)")
    print(f"  computed max_slug_length: {extreme_max_slug_length} (expected: the floor, {_MIN_INCREMENT_SLUG_LENGTH})")
    if extreme_max_slug_length != _MIN_INCREMENT_SLUG_LENGTH:
        failures.append(
            f"3b: expected the floor ({_MIN_INCREMENT_SLUG_LENGTH}) for a pathologically deep base "
            f"directory, got {extreme_max_slug_length}"
        )
    slug = _slugify_increment(longest_name, "H999999999999999-38-00", extreme_max_slug_length)
    if len(slug) > extreme_max_slug_length:
        failures.append(f"3b: slug {slug!r} ({len(slug)} chars) exceeds its own computed max_slug_length ({extreme_max_slug_length})")
    if not slug:
        failures.append("3b: slug should never come back empty, even in the extreme case")
    worst_path = max(
        (extreme_increments_dir / slug / rel for rel in RELATIVE_FILE_PATHS),
        key=lambda p: len(str(p)),
    )
    print(f"  slug_len={len(slug)} worst_path_len={len(str(worst_path))} (NOT asserted <= {_TARGET_MAX_PATH_LENGTH} -- see above)")

    # ------------------------------------------------------------
    # 4 -- Machine independence: the client's team shares one
    # OneDrive-synced data folder across several computers. Two
    # different (non-pathological) base-directory depths -- the
    # client's own real one (127-char prefix) and the shallow
    # hypothetical from section 3a's budget computation (35) -- must
    # produce the IDENTICAL slug for the SAME increment name, since
    # both comfortably exceed the fixed 64-character cap's own
    # trigger point. Without the fixed cap, a shallower machine would
    # have produced a LONGER (and so different) slug for the exact
    # same real file.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("4 -- Machine independence: same name, different base-directory depths")
    print("=" * 70)

    slug_on_client_machine = _slugify_increment(longest_name, "H999999-38-00", max_slug_length)
    slug_on_shallow_machine = _slugify_increment(longest_name, "H999999-38-00", _increment_slug_max_length_for_prefix(20))
    print(f"  Slug computed at the client's own real depth (prefix=127): {slug_on_client_machine!r}")
    print(f"  Slug computed at a much shallower depth (prefix=20):       {slug_on_shallow_machine!r}")
    if slug_on_client_machine != slug_on_shallow_machine:
        failures.append(
            f"4: the same increment name/project number should produce the SAME slug regardless of the "
            f"local machine's own base-directory depth (once both are within the fixed cap), got "
            f"{slug_on_client_machine!r} vs {slug_on_shallow_machine!r}"
        )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- the exact pre-fix client failure (261/242/114 characters) is reproduced and "
            "confirmed as the real root cause; every path the app would create for all 13 real HMR/SMR "
            f"fixtures under the client's own real base directory stays at or under {_TARGET_MAX_PATH_LENGTH} "
            "characters, including the longest file written inside each increment folder; a realistically "
            "longer username/project name than the client's own real one still stays within budget; an "
            "extreme, pathologically deep base directory correctly hits the documented floor instead of "
            "crashing or producing a nonsensical slug; and the same increment name produces the IDENTICAL "
            "slug regardless of the local machine's own base-directory depth, once both are within the "
            "fixed 64-character cap."
        )


if __name__ == "__main__":
    main()

"""
Backward-compatibility regression test for the slug-length fix
(core.project_store._slugify_increment(), see test_increment_slug_length.py):
the client already has the app in live use, so real project folders may
already exist on disk with a long, unbounded slug created under the OLD
pre-fix _slugify() logic. This test confirms those pre-existing folders
remain fully reachable after the fix -- not orphaned by a new, separately-
truncated slug for the "same" increment going forward.

None of the 13 real HMR/SMR fixtures happen to land in the at-risk band
(a slug longer than the NEW _MAX_INCREMENT_SLUG_LENGTH cap but still
under the OS's 255-byte hard limit): they're either short enough that
neither cap ever mattered, or so long (307-546 bytes) that the OLD code
would have crashed with OSError before ever creating a real folder for
them in the first place -- meaning no real pre-existing folder could
exist for those specific names either. The at-risk band (a folder that
really did get created successfully under the old code, but whose name
is long enough that the new cap would truncate it differently for a
brand-new increment of the same name) is constructed directly instead,
using a real HMR 9B-style name truncated to a realistic, representative
length.

Confirms: discovery (list_increments/get_increment) is scan-based (reads
whatever's actually on disk, via increment.json's stored "name" field)
rather than slug-recomputation-based, so this fix cannot orphan an
existing folder -- _slugify_increment() is only ever invoked inside
create_increment(), itself guarded against ever running for a name that
already exists.

Run directly: `python tests/test_increment_folder_backward_compat.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import increment_matcher
from core.project_store import ProjectStore, _slugify, _slugify_increment, _write_json
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Backward Compat Regression"

# Representative of a real client increment name long enough to land in
# the at-risk band (a slug that landed under 255 bytes under the OLD
# code -- so a real folder could genuinely exist -- but over the NEW
# _MAX_INCREMENT_SLUG_LENGTH cap, so a brand-new increment of the same
# name would now be slugified differently).
OLD_INCREMENT_NAME = (
    "UCSF NHPH - Hospital Make Ready 9B - Long Hospital Shear Wall Penetrations and Retrofit, "
    "Connections Renovation, and Below-grade Waterproofing Installation"
)


def main():
    failures = []

    old_slug = _slugify(OLD_INCREMENT_NAME)
    if not (100 < len(old_slug) < 255):
        failures.append(
            f"setup: fixture name's old slug should land between the new cap and the OS limit "
            f"(100-255 bytes), got {len(old_slug)}"
        )

    with tempfile.TemporaryDirectory() as tmp:
        project_store = ProjectStore(base_dir=Path(tmp) / "data")
        store = MockDataStore(project_store)
        store.add_project(PROJECT_NAME)
        project = project_store.get_project(PROJECT_NAME)

        # ------------------------------------------------------------
        # 1 -- Manually construct a folder EXACTLY as the OLD, pre-fix
        # create_increment() would have (bare _slugify(), no cap/
        # disambiguation) -- simulating a real client folder that
        # already exists in production from before this fix shipped.
        # ------------------------------------------------------------
        _write_json(
            project_store._increment_json_path(project.slug, old_slug),
            {"name": OLD_INCREMENT_NAME, "version": 1, "last_updated": "2026-01-01"},
        )
        _write_json(project_store._status_json_path(project.slug, old_slug), {})
        files_dir = project_store._files_dir(project.slug, old_slug)
        files_dir.mkdir(parents=True, exist_ok=True)
        (files_dir / "v1_2026-01-01.xlsm").write_bytes(b"fake pre-existing version content")

        new_slug_if_created_today = _slugify_increment(OLD_INCREMENT_NAME, project_number="H231359-38-00")
        print(f"Pre-existing (OLD) folder slug:            {old_slug}")
        print(f"What NEW code would generate (comparison):  {new_slug_if_created_today}")
        if new_slug_if_created_today == old_slug:
            failures.append(
                "setup: expected the new and old slugs to actually differ for this name -- otherwise this "
                "test isn't exercising the at-risk scenario at all"
            )

        # ------------------------------------------------------------
        # 2 -- list_increments()/get_increment() find the pre-existing
        # folder (discovery is scan-based, not slug-recompute-based).
        # ------------------------------------------------------------
        found = project_store.get_increment(PROJECT_NAME, OLD_INCREMENT_NAME)
        print(f"\nget_increment() found: name={found.name if found else None!r}, slug={found.slug if found else None!r}")
        if found is None or found.slug != old_slug:
            failures.append(
                f"2: expected get_increment() to find the pre-existing folder (slug={old_slug!r}), "
                f"got {found}"
            )

        # ------------------------------------------------------------
        # 3 -- match_increment() (the real identity-matching function a
        # fresh upload of this same project's file would be checked
        # against) returns EXACT_MATCH, not NO_MATCH -- it must never
        # look like a brand-new increment just because the on-disk slug
        # scheme changed.
        # ------------------------------------------------------------
        existing_names = [i.name for i in project_store.list_increments(PROJECT_NAME)]
        match = increment_matcher.match_increment(existing_names, OLD_INCREMENT_NAME)
        print(f"match_increment() result: {match.match_type.name}")
        if match.match_type.name != "EXACT_MATCH":
            failures.append(f"3: expected EXACT_MATCH against the pre-existing increment, got {match.match_type.name}")

        # ------------------------------------------------------------
        # 4 -- Uploading a new version (the real path for an existing
        # increment) writes into the SAME pre-existing old-slug folder,
        # not a new sibling folder under the new scheme -- no
        # duplicate, no orphan.
        # ------------------------------------------------------------
        real_file = os.path.join(FIXDIR, "TIO H231359-38-00_HMR 9B.xlsm")
        if os.path.exists(real_file):
            record = project_store.save_new_version(PROJECT_NAME, OLD_INCREMENT_NAME, real_file)
            print(f"\nsave_new_version() -> slug={record.slug!r}, version={record.version}")
            if record.slug != old_slug:
                failures.append(f"4: save_new_version() should reuse the pre-existing slug {old_slug!r}, got {record.slug!r}")

            all_dirs = [d.name for d in project_store._increment_dirs(project.slug)]
            print(f"All increment directories on disk: {all_dirs}")
            if all_dirs != [old_slug]:
                failures.append(f"4: expected exactly one directory ({old_slug!r}), got {all_dirs}")
        else:
            print(f"\n(skipping step 4 -- real fixture {real_file!r} not present in this checkout)")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- a pre-existing increment folder created under the OLD, unbounded slug scheme "
              "remains fully reachable after the fix: found by get_increment(), matched EXACT_MATCH by "
              "increment_matcher, and a new version upload writes into the SAME folder -- no orphaning, "
              "no duplicate increment.")


if __name__ == "__main__":
    main()

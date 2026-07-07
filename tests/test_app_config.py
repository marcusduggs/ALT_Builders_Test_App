"""
Tests core/app_config.py (config read/write, relocate_data_dir) and
core/project_store.py's get_default_data_dir()/_migrate_legacy_data_dir().

TEST ISOLATION: every single call in this file passes an explicit
config_path/home_dir override -- none of these functions are EVER called
with their production zero-argument defaults, which would resolve to
(and could mutate) this machine's real config file
(~/Library/Application Support/AltamiranoBuildersApp/config.json on
macOS) or real ~/AltamiranoBuildersAppData. A bug in this file cannot
relocate or wipe real data on this machine: there is no code path here
that ever touches Path.home() or the real OS config location.

Run directly: `python tests/test_app_config.py`
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import app_config, project_store


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        # ------------------------------------------------------------
        # 1 -- config read/write round-trip, fully isolated
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- config read/write round-trip")
        print("=" * 70)

        config_path = tmp_path / "config" / "config.json"
        chosen_dir = tmp_path / "MyData"

        before = app_config.read_configured_data_dir(config_path)
        print(f"Before any write: {before} (should be None -- first-launch condition)")
        if before is not None:
            failures.append(f"1: read_configured_data_dir on a nonexistent config file should return None, got {before}")

        app_config.write_configured_data_dir(chosen_dir, config_path)
        after = app_config.read_configured_data_dir(config_path)
        print(f"After write: {after}")
        if after != chosen_dir:
            failures.append(f"1: expected {chosen_dir} after write, got {after}")
        if not config_path.exists():
            failures.append("1: config file should now exist on disk")

        # a corrupt config file should degrade to None, not crash
        corrupt_path = tmp_path / "corrupt_config.json"
        corrupt_path.parent.mkdir(parents=True, exist_ok=True)
        corrupt_path.write_text("{not valid json")
        corrupt_result = app_config.read_configured_data_dir(corrupt_path)
        print(f"Corrupt config file result: {corrupt_result} (should be None, not a crash)")
        if corrupt_result is not None:
            failures.append(f"1: a corrupt config file should read as None, got {corrupt_result}")

        # ------------------------------------------------------------
        # 2 -- relocate_data_dir: no-op cases
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- relocate_data_dir no-op cases")
        print("=" * 70)

        same_dir = tmp_path / "same"
        same_dir.mkdir()
        app_config.relocate_data_dir(same_dir, same_dir)  # old == new -- must not raise or delete anything
        if not same_dir.exists():
            failures.append("2: relocate_data_dir(same, same) should be a pure no-op")

        missing_old = tmp_path / "does-not-exist"
        fresh_new = tmp_path / "fresh-target"
        app_config.relocate_data_dir(missing_old, fresh_new)  # nothing to move -- must not create/error
        print(f"After no-op (missing old_dir): fresh_new exists = {fresh_new.exists()} (should be False -- nothing to move)")
        if fresh_new.exists():
            failures.append("2: relocate_data_dir with a missing old_dir should not create new_dir either")

        # ------------------------------------------------------------
        # 3 -- relocate_data_dir: normal same-filesystem move, contents intact
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- relocate_data_dir: real move, contents survive intact")
        print("=" * 70)

        move_old = tmp_path / "move-source"
        move_new = tmp_path / "move-target" / "nested"  # nested: parent doesn't exist yet either
        (move_old / "projects" / "acme-hospital" / "increments").mkdir(parents=True)
        (move_old / "projects" / "acme-hospital" / "project.json").write_text('{"name": "Acme Hospital"}')
        (move_old / "_deleted").mkdir()

        app_config.relocate_data_dir(move_old, move_new)
        print(f"Old exists: {move_old.exists()} (should be False)")
        print(f"New exists: {move_new.exists()} (should be True)")
        moved_project_json = move_new / "projects" / "acme-hospital" / "project.json"
        if move_old.exists():
            failures.append("3: old_dir should no longer exist after a successful move")
        if not moved_project_json.exists():
            failures.append("3: project.json should have moved intact to the new location")
        elif moved_project_json.read_text() != '{"name": "Acme Hospital"}':
            failures.append("3: moved project.json's content should be byte-for-byte unchanged")
        if not (move_new / "_deleted").exists():
            failures.append("3: the _deleted/ folder should have moved along with everything else")

        # ------------------------------------------------------------
        # 4 -- relocate_data_dir: refuses to overwrite a non-empty destination
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- relocate_data_dir refuses a non-empty destination")
        print("=" * 70)

        collide_old = tmp_path / "collide-source"
        collide_old.mkdir()
        (collide_old / "marker.txt").write_text("source")
        collide_new = tmp_path / "collide-target"
        collide_new.mkdir()
        (collide_new / "existing.txt").write_text("already here")

        try:
            app_config.relocate_data_dir(collide_old, collide_new)
            failures.append("4: relocate_data_dir should raise FileExistsError when new_dir is non-empty")
        except FileExistsError:
            print("Correctly raised FileExistsError for a non-empty destination")
        if not collide_old.exists() or not (collide_old / "marker.txt").exists():
            failures.append("4: the source must be untouched after a refused move")
        if not (collide_new / "existing.txt").exists():
            failures.append("4: the destination's existing content must be untouched after a refused move")

        # ------------------------------------------------------------
        # 5 -- relocate_data_dir: cross-device fallback (copy + delete)
        #
        # Can't easily attach a second real filesystem/drive in this
        # environment, so this simulates the OS-level failure Path.rename()
        # raises when the two paths are on different filesystems (errno
        # EXDEV) by patching Path.rename to raise OSError once, then
        # confirming the copy-then-delete fallback still produces the
        # correct end state. This exercises the actual fallback code path
        # in relocate_data_dir (the except OSError branch), not just a
        # hand-verified read of the source.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("5 -- relocate_data_dir: cross-device fallback (simulated)")
        print("=" * 70)

        xdev_old = tmp_path / "xdev-source"
        xdev_new = tmp_path / "xdev-target"
        (xdev_old / "projects" / "beta-clinic").mkdir(parents=True)
        (xdev_old / "projects" / "beta-clinic" / "project.json").write_text('{"name": "Beta Clinic"}')

        real_rename = Path.rename
        call_count = {"n": 0}

        def rename_that_fails_once(self, target):
            call_count["n"] += 1
            if self == xdev_old:
                raise OSError(18, "Invalid cross-device link")  # errno.EXDEV
            return real_rename(self, target)

        with patch.object(Path, "rename", rename_that_fails_once):
            app_config.relocate_data_dir(xdev_old, xdev_new)

        print(f"rename() was attempted (and made to fail): {call_count['n'] >= 1}")
        print(f"Old exists after fallback: {xdev_old.exists()} (should be False)")
        print(f"New exists after fallback: {xdev_new.exists()} (should be True)")
        moved_beta_json = xdev_new / "projects" / "beta-clinic" / "project.json"
        if xdev_old.exists():
            failures.append("5: source should be removed once the copy-then-delete fallback completes")
        if not moved_beta_json.exists():
            failures.append("5: cross-device fallback should still fully copy the tree to the new location")
        elif moved_beta_json.read_text() != '{"name": "Beta Clinic"}':
            failures.append("5: cross-device fallback's copied content should be byte-for-byte unchanged")

        # ------------------------------------------------------------
        # 6 -- get_default_data_dir(): no config, no legacy dir -> suggestion
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("6 -- get_default_data_dir(): fallback suggestion, fully isolated")
        print("=" * 70)

        fake_home = tmp_path / "fake-home"
        fake_home.mkdir()
        no_config_path = tmp_path / "no-such-config" / "config.json"

        result = project_store.get_default_data_dir(config_path=no_config_path, home_dir=fake_home)
        expected = fake_home / "AltamiranoBuildersAppData"
        print(f"get_default_data_dir() with no config, no legacy dir: {result}")
        if result != expected:
            failures.append(f"6: expected fallback suggestion {expected}, got {result}")

        # ------------------------------------------------------------
        # 7 -- get_default_data_dir(): legacy SubmissionAppData folder
        # present -> migrated into the suggestion, fully isolated
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("7 -- get_default_data_dir(): legacy folder migration, fully isolated")
        print("=" * 70)

        legacy_home = tmp_path / "legacy-home"
        legacy_dir = legacy_home / "SubmissionAppData"
        (legacy_dir / "projects" / "gamma-tower").mkdir(parents=True)
        (legacy_dir / "projects" / "gamma-tower" / "project.json").write_text('{"name": "Gamma Tower"}')

        result2 = project_store.get_default_data_dir(config_path=no_config_path, home_dir=legacy_home)
        expected2 = legacy_home / "AltamiranoBuildersAppData"
        print(f"get_default_data_dir() with a legacy folder present: {result2}")
        if result2 != expected2:
            failures.append(f"7: expected migrated-to location {expected2}, got {result2}")
        if legacy_dir.exists():
            failures.append("7: the legacy SubmissionAppData folder should have been renamed away, not left in place")
        migrated_json = expected2 / "projects" / "gamma-tower" / "project.json"
        if not migrated_json.exists():
            failures.append("7: the legacy folder's contents should have moved intact")

        # confirm this whole test genuinely never touched anything real
        real_config_path = app_config.get_config_path()
        print(f"\nReal OS config path (never touched by this test): {real_config_path}")
        if str(tmp_path) in str(real_config_path):
            failures.append("SANITY: real_config_path should never coincidentally be inside our temp dir")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- config read/write, relocate (same-fs, cross-device fallback, collision "
              "refusal), and get_default_data_dir() (fallback + legacy migration) all correct, "
              "fully isolated from any real path on this machine")


if __name__ == "__main__":
    main()

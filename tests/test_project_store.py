"""
Tests core/project_store.py against a temp directory -- never touches the
real ~/AltamiranoBuildersAppData.

Run directly: `python tests/test_project_store.py`
"""

import json
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.project_store import ProjectStore


def _make_fake_file(dir_path: Path, name: str) -> str:
    path = dir_path / name
    path.write_bytes(b"fake xlsm content")
    return str(path)


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        store = ProjectStore(base_dir=tmp_path / "data")
        scratch = tmp_path / "uploads"
        scratch.mkdir()

        # --- project CRUD ---
        store.create_project("Test Hospital - Wing A")
        store.create_project("Test Hospital - Wing B")

        names = {p.name for p in store.list_projects()}
        if names != {"Test Hospital - Wing A", "Test Hospital - Wing B"}:
            failures.append(f"list_projects mismatch: {names}")

        original = store.get_project("Test Hospital - Wing A")
        # home_folder is purely computed -- always THIS project's own
        # folder (base_dir/projects/<slug>/), never user-suppliable, and
        # never stale relative to the slug it's derived from.
        expected_home_folder = str(store.projects_dir / original.slug) if original else None
        if original is None or original.home_folder != expected_home_folder:
            failures.append(f"get_project mismatch: {original}, expected home_folder {expected_home_folder!r}")

        # give Wing A a real increment+file+status BEFORE renaming, so the
        # rename test below can confirm they survive the folder move intact
        wing_a_file = _make_fake_file(scratch, "wing-a-v1.xlsm")
        store.create_increment("Test Hospital - Wing A", "INC 1 - Wing A Foundation", wing_a_file)
        store.save_status("Test Hospital - Wing A", "INC 1 - Wing A Foundation", {"B-C1": {5: "Done"}})
        old_slug = original.slug
        old_project_dir = store.projects_dir / old_slug

        store.update_project("Test Hospital - Wing A", "Test Hospital - Wing A (Renamed)")
        renamed = store.get_project("Test Hospital - Wing A (Renamed)")
        # Home Folder is a live, user-visible path -- renaming MUST move
        # the folder to match, not leave it pointing at the pre-rename slug.
        if renamed is None:
            failures.append("update_project: renamed project should be resolvable by its new name")
        else:
            if renamed.slug == old_slug:
                failures.append(f"update_project should change the slug/folder to match the new name, still {old_slug!r}")
            if renamed.home_folder != str(store.projects_dir / renamed.slug):
                failures.append(f"update_project: home_folder should reflect the NEW slug, got {renamed.home_folder!r}")
            if old_project_dir.exists():
                failures.append(f"update_project should have moved the old folder, but {old_project_dir} still exists")
        if store.get_project("Test Hospital - Wing A") is not None:
            failures.append("update_project should make the old name unresolvable")

        # the increment/file/status created under the OLD slug must still
        # be reachable under the new name -- Path.rename() moves the whole
        # tree, nothing should be orphaned
        moved_increments = store.list_increments("Test Hospital - Wing A (Renamed)")
        if [i.name for i in moved_increments] != ["INC 1 - Wing A Foundation"]:
            failures.append(f"increment should survive the project rename intact, got {moved_increments}")
        moved_status = store.load_status("Test Hospital - Wing A (Renamed)", "INC 1 - Wing A Foundation")
        if moved_status != {"B-C1": {5: "Done"}}:
            failures.append(f"status.json should survive the project rename intact, got {moved_status}")
        moved_versions = store.list_version_files("Test Hospital - Wing A (Renamed)", "INC 1 - Wing A Foundation")
        if len(moved_versions) != 1 or not moved_versions[0].exists():
            failures.append(f"version file should survive the project rename intact, got {moved_versions}")

        # renaming to a name that slugifies IDENTICALLY (punctuation only)
        # must be a no-op on the folder -- same slug, no actual move. Uses
        # a separate, dedicated project (not Wing A) so it doesn't disturb
        # Wing A's name for the increment/version-history section below.
        store.create_project("No-Op Rename Test")
        no_op_before = store.get_project("No-Op Rename Test")
        store.update_project("No-Op Rename Test", "No-Op Rename Test!!!")
        no_op_after = store.get_project("No-Op Rename Test!!!")
        if no_op_after is None or no_op_after.slug != no_op_before.slug:
            failures.append(
                f"renaming to a name with the same slug should keep the same slug/folder, "
                f"got {no_op_after.slug if no_op_after else None!r} vs {no_op_before.slug!r}"
            )

        # --- collision: renaming into another EXISTING project's slug ---
        store.create_project("Zeta Clinic")
        store.create_project("Zeta Clinic Collision Source")
        zeta = store.get_project("Zeta Clinic")
        store.update_project("Zeta Clinic Collision Source", "Zeta Clinic")
        renamed_to_zeta = next(
            (p for p in store.list_projects() if p.slug != zeta.slug and p.slug.startswith("zeta-clinic")), None
        )
        if renamed_to_zeta is None:
            failures.append(
                "renaming into a name that collides with another project's slug should suffix "
                "(e.g. zeta-clinic-2), not silently merge"
            )
        if not (store.projects_dir / zeta.slug).exists():
            failures.append("the ORIGINAL 'Zeta Clinic' project's folder must be untouched by the colliding rename")

        # clean up this section's throwaway projects/increment so the
        # project-count and increment-list assertions further down (which
        # predate this rename test and assume a minimal, unrelated set of
        # projects/increments) aren't disturbed by it
        store.delete_project("No-Op Rename Test!!!")
        store.delete_project("Zeta Clinic")
        if renamed_to_zeta is not None:
            store.delete_project(renamed_to_zeta.name)
        store.delete_increment("Test Hospital - Wing A (Renamed)", "INC 1 - Wing A Foundation")

        wing_b = store.get_project("Test Hospital - Wing B")
        wing_b_slug = wing_b.slug
        wing_b_original_path = store.projects_dir / wing_b_slug

        # give Wing B a real file so we can confirm its contents survive
        # the move to _deleted/ untouched, not just its folder name
        wing_b_file = _make_fake_file(scratch, "wing-b-v1.xlsm")
        store.create_increment("Test Hospital - Wing B", "INC 1 - Wing B Foundation", wing_b_file)
        store.save_status("Test Hospital - Wing B", "INC 1 - Wing B Foundation", {"B-C1": {5: "Done"}})

        store.delete_project("Test Hospital - Wing B")
        if store.get_project("Test Hospital - Wing B") is not None:
            failures.append("delete_project did not remove the project")
        if len(store.list_projects()) != 1:
            failures.append(f"expected 1 project after delete, found {len(store.list_projects())}")
        if wing_b_original_path.exists():
            failures.append("delete_project should remove the project's folder from projects/, not just hide it")

        # --- soft-delete: folder moved intact under _deleted/, not destroyed ---
        deleted_dir = store.base_dir / "_deleted"
        deleted_candidates = [
            d for d in deleted_dir.iterdir() if d.is_dir() and d.name.endswith(f"_{wing_b_slug}")
        ] if deleted_dir.exists() else []
        if len(deleted_candidates) != 1:
            failures.append(
                f"expected exactly 1 folder under _deleted/ ending in '_{wing_b_slug}', found {deleted_candidates}"
            )
        else:
            deleted_wing_b = deleted_candidates[0]
            moved_increment_json = deleted_wing_b / "increments" / "inc-1-wing-b-foundation" / "increment.json"
            moved_status_json = deleted_wing_b / "increments" / "inc-1-wing-b-foundation" / "status.json"
            moved_files_dir = deleted_wing_b / "increments" / "inc-1-wing-b-foundation" / "files"
            if not moved_increment_json.exists():
                failures.append(f"soft-deleted project missing increment.json at {moved_increment_json}")
            if json.loads(moved_status_json.read_text()) != {"B-C1": {"5": "Done"}}:
                failures.append("soft-deleted project's status.json should be preserved exactly as-is")
            if not any(moved_files_dir.iterdir()):
                failures.append("soft-deleted project's version files/ should be preserved, found none")

        # --- deleting two similarly-named/sluggy projects must not collide in _deleted/ ---
        store.create_project("Collision Test")
        store.delete_project("Collision Test")
        store.create_project("Collision Test")
        store.delete_project("Collision Test")
        collision_dirs = [d for d in deleted_dir.iterdir() if d.is_dir() and "collision-test" in d.name]
        if len(collision_dirs) != 2:
            failures.append(
                f"expected 2 distinct _deleted/ folders for 2 deletions of same-named projects, "
                f"found {len(collision_dirs)}: {collision_dirs}"
            )

        # --- increments + version history ---
        project_name = "Test Hospital - Wing A (Renamed)"
        file_v1 = _make_fake_file(scratch, "v1.xlsm")
        record = store.create_increment(project_name, "INC 1 - Foundation", file_v1)
        if record.version != 1:
            failures.append(f"create_increment should start at version 1, got {record.version}")

        # duplicate increment name should be rejected
        try:
            store.create_increment(project_name, "INC 1 - Foundation", file_v1)
            failures.append("create_increment should reject a duplicate increment name")
        except ValueError:
            pass

        increments = store.list_increments(project_name)
        if [i.name for i in increments] != ["INC 1 - Foundation"]:
            failures.append(f"list_increments mismatch: {increments}")

        status = store.load_status(project_name, "INC 1 - Foundation")
        if status != {}:
            failures.append(f"fresh increment should have empty status.json, got {status}")

        file_v2 = _make_fake_file(scratch, "v2.xlsm")
        updated = store.save_new_version(project_name, "INC 1 - Foundation", file_v2)
        if updated.version != 2:
            failures.append(f"save_new_version should bump to version 2, got {updated.version}")

        versions = store.list_version_files(project_name, "INC 1 - Foundation")
        if len(versions) != 2:
            failures.append(f"expected 2 stored version files, got {len(versions)}: {versions}")
        elif not versions[0].name.startswith("v1_") or not versions[1].name.startswith("v2_"):
            failures.append(f"version files not in v1, v2 order: {[v.name for v in versions]}")

        current = store.current_version_path(project_name, "INC 1 - Foundation")
        if current is None or not current.name.startswith("v2_"):
            failures.append(f"current_version_path should point at v2, got {current}")

        v1_path = store.version_path(project_name, "INC 1 - Foundation", 1)
        if v1_path is None or not v1_path.name.startswith("v1_"):
            failures.append(f"version_path(1) mismatch: {v1_path}")

        # list_versions(): newest first, for a version-selector dropdown
        version_records = store.list_versions(project_name, "INC 1 - Foundation")
        if [v.version for v in version_records] != [2, 1]:
            failures.append(f"list_versions should be newest-first [2, 1], got {[v.version for v in version_records]}")
        today_str = date.today().isoformat()
        if any(v.uploaded_date != today_str for v in version_records):
            failures.append(f"list_versions dates should all be today ({today_str}), got {version_records}")

        # original upload files must be untouched (copied, not moved)
        if not Path(file_v1).exists() or not Path(file_v2).exists():
            failures.append("create_increment/save_new_version should COPY the source file, not move it")

        # --- status.json round trip, at (index, stage) granularity ---
        store.save_status(
            project_name, "INC 1 - Foundation", {"B-C1": {5: "Done", 12: "Open"}, "B-F1": {1: "Open"}}
        )
        reloaded = store.load_status(project_name, "INC 1 - Foundation")
        if reloaded != {"B-C1": {5: "Done", 12: "Open"}, "B-F1": {1: "Open"}}:
            failures.append(f"status.json round-trip mismatch: {reloaded}")

        # stage numbers must round-trip as ints even though the on-disk
        # JSON necessarily uses string keys (JSON object keys are always
        # strings) -- callers of ProjectStore should never have to think
        # about that.
        raw_on_disk = json.loads(store._status_json_path(
            store.get_project(project_name).slug,
            store.get_increment(project_name, "INC 1 - Foundation").slug,
        ).read_text())
        if raw_on_disk != {"B-C1": {"5": "Done", "12": "Open"}, "B-F1": {"1": "Open"}}:
            failures.append(f"on-disk status.json should use string stage keys, got {raw_on_disk}")

        # --- increment-level soft-delete: only the target increment
        # moves, the project and its OTHER increments are untouched ---
        file_v1_b = _make_fake_file(scratch, "inc2-v1.xlsm")
        store.create_increment(project_name, "INC 2 - Sitework", file_v1_b)
        store.save_status(project_name, "INC 2 - Sitework", {"C-F1": {3: "Done"}})
        inc2 = store.get_increment(project_name, "INC 2 - Sitework")
        inc2_slug = inc2.slug
        inc2_original_path = store._increments_dir(store.get_project(project_name).slug) / inc2_slug

        store.delete_increment(project_name, "INC 2 - Sitework")

        if store.get_increment(project_name, "INC 2 - Sitework") is not None:
            failures.append("delete_increment did not remove the increment")
        if inc2_original_path.exists():
            failures.append("delete_increment should remove the increment's folder from increments/, not just hide it")

        remaining_increments = {i.name for i in store.list_increments(project_name)}
        if remaining_increments != {"INC 1 - Foundation"}:
            failures.append(
                f"delete_increment should leave the project's OTHER increments untouched, got {remaining_increments}"
            )
        if store.get_project(project_name) is None:
            failures.append("delete_increment should never remove the parent project itself")

        # soft-delete: folder moved intact under _deleted/, prefixed with
        # BOTH project and increment slug (increment slugs alone aren't
        # unique across projects)
        deleted_project_slug = store.get_project(project_name).slug
        inc_deleted_candidates = [
            d for d in deleted_dir.iterdir()
            if d.is_dir() and d.name.endswith(f"_{deleted_project_slug}_{inc2_slug}")
        ]
        if len(inc_deleted_candidates) != 1:
            failures.append(
                f"expected exactly 1 folder under _deleted/ ending in "
                f"'_{deleted_project_slug}_{inc2_slug}', found {inc_deleted_candidates}"
            )
        else:
            deleted_inc2 = inc_deleted_candidates[0]
            moved_increment_json = deleted_inc2 / "increment.json"
            moved_status_json = deleted_inc2 / "status.json"
            moved_files_dir = deleted_inc2 / "files"
            if not moved_increment_json.exists():
                failures.append(f"soft-deleted increment missing increment.json at {moved_increment_json}")
            if not moved_status_json.exists() or json.loads(moved_status_json.read_text()) != {"C-F1": {"3": "Done"}}:
                failures.append("soft-deleted increment's status.json should be preserved exactly as-is")
            if not moved_files_dir.exists() or not any(moved_files_dir.iterdir()):
                failures.append("soft-deleted increment's version files/ should be preserved, found none")

        # deleting a nonexistent increment should be a no-op, not raise
        store.delete_increment(project_name, "does not exist")
        store.delete_increment("no such project", "does not exist")

        # --- nonexistent lookups shouldn't raise ---
        if store.get_increment(project_name, "does not exist") is not None:
            failures.append("get_increment should return None for a missing increment")
        if store.current_version_path(project_name, "does not exist") is not None:
            failures.append("current_version_path should return None for a missing increment")
        if store.load_status("no such project", "no such increment") != {}:
            failures.append("load_status should return {} for a missing project/increment")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- create/list/update/delete + version history all correct")


if __name__ == "__main__":
    main()

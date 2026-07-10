"""
Tests core/project_store.py's comments.json CRUD (add_comment/
update_comment/delete_comment/load_comments) against a temp directory --
never touches the real ~/AltamiranoBuildersAppData.

Run directly: `python tests/test_comments.py`
"""

import os
import sys
import tempfile
import time
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

        store.create_project("Comments Test Project")
        source_file = _make_fake_file(scratch, "inc1.xlsm")
        store.create_increment("Comments Test Project", "INC 1 - Test Scope", source_file)
        PROJECT, INCREMENT = "Comments Test Project", "INC 1 - Test Scope"

        # ------------------------------------------------------------
        # 1 -- empty state
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- empty state")
        print("=" * 70)
        comments = store.load_comments(PROJECT, INCREMENT)
        print(f"load_comments on a fresh increment: {comments}")
        if comments != []:
            failures.append(f"1: expected no comments yet, got {comments}")

        # ------------------------------------------------------------
        # 2 -- add_comment
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- add_comment")
        print("=" * 70)
        entry_1 = store.add_comment(PROJECT, INCREMENT, "First comment about this increment.")
        print(f"add_comment result: {entry_1}")

        for key in ("id", "timestamp", "text", "edited_timestamp"):
            if key not in entry_1:
                failures.append(f"2: add_comment's returned entry is missing key {key!r}")
        if entry_1.get("text") != "First comment about this increment.":
            failures.append(f"2: unexpected text: {entry_1.get('text')!r}")
        if entry_1.get("edited_timestamp") is not None:
            failures.append(f"2: a brand-new comment should have edited_timestamp=None, got {entry_1.get('edited_timestamp')!r}")
        if not entry_1.get("id"):
            failures.append("2: expected a non-empty id")

        time.sleep(0.01)  # ensure a second comment's timestamp is strictly later, for ordering checks below
        entry_2 = store.add_comment(PROJECT, INCREMENT, "Second comment, added later.")
        print(f"add_comment (2nd) result: {entry_2}")

        if entry_1["id"] == entry_2["id"]:
            failures.append("2: two different comments got the same id")

        loaded = store.load_comments(PROJECT, INCREMENT)
        print(f"load_comments after 2 adds: {loaded}")
        if [c["id"] for c in loaded] != [entry_1["id"], entry_2["id"]]:
            failures.append(f"2: expected oldest-first order [{entry_1['id']}, {entry_2['id']}], got {[c['id'] for c in loaded]}")

        # ------------------------------------------------------------
        # 3 -- update_comment: text changes, id and original timestamp
        # are PRESERVED, edited_timestamp gets set
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- update_comment (edit)")
        print("=" * 70)
        original_id = entry_1["id"]
        original_timestamp = entry_1["timestamp"]

        time.sleep(0.01)
        updated = store.update_comment(PROJECT, INCREMENT, original_id, "First comment -- EDITED text.")
        print(f"update_comment result: {updated}")

        if updated is None:
            failures.append("3: update_comment returned None for an existing id")
        else:
            if updated["id"] != original_id:
                failures.append(f"3: id changed on edit -- was {original_id!r}, now {updated['id']!r}")
            if updated["timestamp"] != original_timestamp:
                failures.append(
                    f"3: original creation timestamp was NOT preserved on edit -- "
                    f"was {original_timestamp!r}, now {updated['timestamp']!r}"
                )
            if updated["text"] != "First comment -- EDITED text.":
                failures.append(f"3: text was not updated, got {updated['text']!r}")
            if not updated.get("edited_timestamp"):
                failures.append("3: edited_timestamp should be set (non-empty) after an edit")
            elif updated["edited_timestamp"] == original_timestamp:
                failures.append("3: edited_timestamp should differ from the original creation timestamp")

        # Confirm the SECOND comment (never edited) is untouched by editing the first.
        reloaded = store.load_comments(PROJECT, INCREMENT)
        second_reloaded = next(c for c in reloaded if c["id"] == entry_2["id"])
        print(f"Second comment after editing the first (should be untouched): {second_reloaded}")
        if second_reloaded != entry_2:
            failures.append(f"3: editing comment 1 changed comment 2 -- expected {entry_2}, got {second_reloaded}")

        # update_comment on a NON-existent id is a no-op, not an error.
        noop_result = store.update_comment(PROJECT, INCREMENT, "not-a-real-id", "irrelevant")
        print(f"update_comment on a bogus id: {noop_result} (expected None)")
        if noop_result is not None:
            failures.append(f"3: update_comment on a bogus id should return None, got {noop_result}")

        # ------------------------------------------------------------
        # 4 -- persistence across a fresh ProjectStore instance
        # (simulates an app restart -- nothing should be cached in
        # memory that a reload wouldn't also see)
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- persistence across reload")
        print("=" * 70)
        reloaded_store = ProjectStore(base_dir=tmp_path / "data")
        reloaded_comments = reloaded_store.load_comments(PROJECT, INCREMENT)
        print(f"Comments from a freshly-constructed ProjectStore: {reloaded_comments}")
        if reloaded_comments != reloaded:
            failures.append(
                f"4: comments.json did not persist correctly across a fresh ProjectStore instance -- "
                f"expected {reloaded}, got {reloaded_comments}"
            )

        # ------------------------------------------------------------
        # 5 -- delete_comment: removes exactly the targeted comment,
        # leaves the other untouched, is a no-op on an already-deleted
        # or unknown id, and an empty list (not a missing/corrupt file)
        # once every comment is gone
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("5 -- delete_comment")
        print("=" * 70)
        store.delete_comment(PROJECT, INCREMENT, entry_2["id"])
        after_delete = store.load_comments(PROJECT, INCREMENT)
        print(f"Comments after deleting comment 2: {after_delete}")
        if [c["id"] for c in after_delete] != [original_id]:
            failures.append(f"5: expected only comment 1 (edited) to remain (id {original_id}), got {[c['id'] for c in after_delete]}")

        # Deleting an already-gone id is a no-op, not an error.
        store.delete_comment(PROJECT, INCREMENT, entry_2["id"])
        after_double_delete = store.load_comments(PROJECT, INCREMENT)
        print(f"Comments after deleting the same id again (should be unchanged): {after_double_delete}")
        if after_double_delete != after_delete:
            failures.append("5: deleting an already-deleted comment id changed something -- should be a pure no-op")

        # Deleting on a bogus, never-existed id is also a no-op.
        store.delete_comment(PROJECT, INCREMENT, "not-a-real-id")
        after_bogus_delete = store.load_comments(PROJECT, INCREMENT)
        if after_bogus_delete != after_delete:
            failures.append("5: deleting a bogus comment id changed something -- should be a pure no-op")

        # Deleting the LAST remaining comment leaves an empty list, not
        # a missing/corrupt file.
        store.delete_comment(PROJECT, INCREMENT, original_id)
        empty_after_all_deleted = store.load_comments(PROJECT, INCREMENT)
        print(f"Comments after deleting the last one: {empty_after_all_deleted}")
        if empty_after_all_deleted != []:
            failures.append(f"5: expected an empty list after deleting every comment, got {empty_after_all_deleted}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- add/load/update/delete all correct: update_comment preserves id and original creation "
              "timestamp while changing text and setting edited_timestamp, leaves other comments untouched, is a "
              "no-op on unknown ids (matching delete_comment's convention), delete_comment removes exactly the "
              "targeted comment and is a no-op on already-deleted/bogus ids, and comments.json persists correctly "
              "across a fresh ProjectStore instance")


if __name__ == "__main__":
    main()

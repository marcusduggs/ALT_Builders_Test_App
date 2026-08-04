"""
Regression test for a real reported crash: combining INC 2 (sample_increment.xlsm,
42 stages) with INC 8 (inc8_real_2stage.xlsm, 2 stages) via View Combined Data
raised `KeyError: 3` in ui.pages.combined_data_view_page._build_sum_data_tab's
per-increment subtotal row, at `totals["stage_open_counts"][stage]`.

Root cause: core.combined_export.build_combined_sum_data_rows() computes each
increment's own "subtotal_totals" with THAT increment's own (possibly narrower)
stage_count (see its docstring) -- so INC 8's stage_open_counts dict only ever
has keys {1, 2}. The Sum Data tab iterates the SHEET's wider union stage_count
(here, 42, from INC 2) and indexed into that dict directly. Two sibling code
paths already guarded against exactly this gap -- the All Data tab's subtotal
row (ui.pages.combined_data_view_page, `.get(f"Stage {stage}", 0)`) and the
combined .xlsx writer's own Sum Data subtotal row
(core.combined_export._write_sum_data_subtotal_row, `.get(stage, 0)`) -- only
the Sum Data tab's on-screen rendering was missed. Fixed by the matching
`.get(stage, 0)`.

The exact crash site (ui.pages.combined_data_view_page.CombinedDataViewPage.
_build_sum_data_tab, run from __init__) is a QWidget method, and no other
test in this suite spins up Qt (see tests/test_error_surfacing.py's module
docstring) -- but a data-layer-only reproduction of the failing dict access
can't actually regression-test THAT line, since it would just re-implement
the fix rather than exercise it. So this test instantiates the real
CombinedDataViewPage under QT_QPA_PLATFORM=offscreen (set below, before any
Qt import) and reads its actual QStandardItemModel back -- construction and
painting only, no app.exec_()/blocking event loop, matching the spirit of
"no Qt event loop" while still exercising the real buggy line. It also
separately confirms the sibling EXPORT path (core.combined_export, unaffected
by this bug, already correct) still produces the right values end to end.

Run directly: `python tests/test_combined_mixed_stage_counts.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtWidgets import QApplication

from core import combined_export
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
INC8_2STAGE = os.path.join(FIXTURES, "inc8_real_2stage.xlsm")
SAMPLE_42STAGE = os.path.join(FIXTURES, "sample_increment.xlsm")
PROJECT_NAME = "Mixed Stage Count Validation"


def main():
    failures = []
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)

        inc8 = store.add_new_increment(PROJECT_NAME, INC8_2STAGE)
        inc_sample = store.add_new_increment(PROJECT_NAME, SAMPLE_42STAGE)
        d8 = store.get_increment_for_display(PROJECT_NAME, inc8.name)
        dsample = store.get_increment_for_display(PROJECT_NAME, inc_sample.name)

        print(f"INC 8 ({d8.name}): stage_count={d8.stage_count}")
        print(f"INC 2 ({dsample.name}): stage_count={dsample.stage_count}")
        if d8.stage_count >= dsample.stage_count:
            failures.append(
                f"setup: fixture assumption broken -- expected INC 8's stage_count ({d8.stage_count}) "
                f"< INC 2's ({dsample.stage_count})"
            )

        # ------------------------------------------------------------
        # 1 -- build_combined_view() itself must not raise, and its union
        # stage_count must be the WIDER increment's, not the narrower one's
        # (i.e. the mismatch this bug depends on must actually be present).
        # ------------------------------------------------------------
        view = build_combined_view([dsample, d8])
        print(f"Combined view stage_count (union): {view.stage_count}")
        if view.stage_count != dsample.stage_count:
            failures.append(
                f"1: expected combined stage_count to be INC 2's {dsample.stage_count} (the union max), "
                f"got {view.stage_count}"
            )

        # ------------------------------------------------------------
        # 2 -- instantiate the REAL CombinedDataViewPage (the actual
        # crash site: _build_sum_data_tab runs from __init__) and read
        # its actual QStandardItemModel back. Confirms the narrower
        # increment's own stage_open_counts dict really is sparse first
        # (so this test would fail loudly, not silently pass, if the
        # underlying data shape ever stopped being sparse), then confirms
        # construction doesn't raise and the rendered cells hold the
        # right values -- not just that nothing crashed.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- CombinedDataViewPage's real Sum Data subtotal rows (the crash site)")
        print("=" * 70)

        subtotal_rows = {row["increment"]: row for row in view.sum_data_rows if row["is_subtotal"]}
        if d8.name not in subtotal_rows or dsample.name not in subtotal_rows:
            failures.append("2: expected exactly one subtotal row per increment")
        else:
            inc8_stage_open_counts = subtotal_rows[d8.name]["subtotal_totals"]["stage_open_counts"]
            print(f"INC 8 subtotal stage_open_counts keys: {sorted(inc8_stage_open_counts.keys())}")
            if 3 in inc8_stage_open_counts:
                failures.append(
                    "2: fixture/behavior assumption broken -- INC 8's own stage_open_counts should NOT "
                    "have a key for stage 3 (it only has 2 stages); if it now does, this test no longer "
                    "exercises the sparse-dict gap the bug depended on"
                )

        app = QApplication.instance() or QApplication(sys.argv)
        try:
            page = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        except KeyError as e:
            failures.append(
                f"2: constructing CombinedDataViewPage raised KeyError({e}) -- this is the exact reported "
                "crash (ui.pages.combined_data_view_page._build_sum_data_tab indexing "
                "totals['stage_open_counts'][stage] without a default)"
            )
            page = None

        if page is not None:
            model = page.sum_data_table.model()
            FROZEN_COLUMNS = 4  # Increment, Index, Description, Approval Agency

            def _find_subtotal_row(label: str):
                for r in range(model.rowCount()):
                    item = model.item(r, 0)
                    if item is not None and item.text() == label:
                        return r
                return None

            inc8_row = _find_subtotal_row(f"{d8.name} — Subtotal")
            inc2_row = _find_subtotal_row(f"{dsample.name} — Subtotal")
            if inc8_row is None or inc2_row is None:
                failures.append("2: expected a subtotal row for each increment in the rendered model")
            else:
                def _stage_cell(row, stage):
                    item = model.item(row, FROZEN_COLUMNS + (stage - 1))
                    return int(item.text())

                inc8_stage3 = _stage_cell(inc8_row, 3)
                inc2_stage14 = _stage_cell(inc2_row, 14)
                print(f"Rendered INC 8 subtotal row, Stage 3: {inc8_stage3} (expected 0)")
                print(f"Rendered INC 2 subtotal row, Stage 14: {inc2_stage14} (expected 28)")
                if inc8_stage3 != 0:
                    failures.append(f"2: rendered INC 8 subtotal Stage 3 should be 0, got {inc8_stage3}")
                if inc2_stage14 != 28:
                    failures.append(f"2: rendered INC 2 subtotal Stage 14 should be 28, got {inc2_stage14}")

        # ------------------------------------------------------------
        # 3 -- the sibling EXPORT path (core.combined_export) was already
        # correct (it has its own .get(stage, 0) at
        # _write_sum_data_subtotal_row) -- confirm it still is, end to
        # end, against the real .xlsx it writes. Not the bug that was
        # fixed, but the other half of "same underlying data, both paths
        # must agree" the user asked to re-confirm.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- combined EXPORT path (core.combined_export)")
        print("=" * 70)

        export_path = str(Path(tmp) / "combined_mixed.xlsx")
        combined_export.export_combined_report([dsample, d8], export_path)  # must not raise

        wb = openpyxl.load_workbook(export_path)
        ws = wb["Sum Data"]
        col1 = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        inc8_label = f"{d8.name} — Subtotal"
        inc2_label = f"{dsample.name} — Subtotal"
        if inc8_label not in col1 or inc2_label not in col1:
            failures.append("3: expected both increments' subtotal rows in the exported Sum Data sheet")
        else:
            inc8_row = col1.index(inc8_label) + 1
            inc2_row = col1.index(inc2_label) + 1
            stage3_col = 5 + 2  # Increment, Index, Description, Approval Agency, then stages from col 5
            stage14_col = 5 + 13
            inc8_stage3 = ws.cell(row=inc8_row, column=stage3_col).value
            inc2_stage14 = ws.cell(row=inc2_row, column=stage14_col).value
            print(f"Exported INC 8 subtotal row, Stage 3: {inc8_stage3} (expected 0)")
            print(f"Exported INC 2 subtotal row, Stage 14: {inc2_stage14} (expected 28)")
            if inc8_stage3 != 0:
                failures.append(f"3: exported INC 8 subtotal Stage 3 should be 0, got {inc8_stage3}")
            if inc2_stage14 != 28:
                failures.append(f"3: exported INC 2 subtotal Stage 14 should be 28, got {inc2_stage14}")

            totals_row = ws.max_row
            totals_stage14 = ws.cell(row=totals_row, column=stage14_col).value
            print(f"Exported combined Totals row, Stage 14: {totals_stage14} (expected 28, INC 8 contributes 0)")
            if totals_stage14 != 28:
                failures.append(f"3: exported combined Totals Stage 14 should be 28, got {totals_stage14}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- combining a 2-stage and a 42-stage increment no longer crashes the Sum Data "
              "tab's per-increment subtotal rows; stages unique to the wider increment correctly default to 0 "
              "for the narrower one instead of raising KeyError, in both the on-screen data and the export.")


if __name__ == "__main__":
    main()

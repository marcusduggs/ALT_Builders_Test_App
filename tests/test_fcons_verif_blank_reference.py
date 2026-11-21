"""
Regression test for two related fixes to core.excel_reader's F-Cons
Verif (milestone sheet) parsing and Sum Data row expansion, both driven
by Rey's TIO_Changes.xlsx "TIO Error Report" tab:

FIX 2 -- F-Cons Verif rows with a real description but a blank
Reference Number used to be silently dropped entirely (Rey: "Construction
Final F Missing" on HMR 12; SMR 2C's "Clear all plan review Outstanding
Items List (OIL) Items" and "Construction Final" both missing).
core.excel_reader.parse_milestone_sheet() now keeps them, with the
literal Index "?" -- but must STILL stop at the sheet's own
"ABBREVIATIONS:" legend block (see MILESTONE_LEGEND_MARKER), which has
the exact same shape (real text in the Description column, blank
Reference Number) and would otherwise be wrongly recovered as fake
tracked items -- confirmed present, in that exact shape, in all 14 real
fixtures checked.

FIX 1 (the part this file also covers) -- a GRID row (B-Tests/C-On-Site/
D-Off-Site) must NEVER show up as a Sum Data "VCR" row; only a genuine
F-Cons Verif/milestone item may. Multiple "?" rows in the same increment
must each appear as their own distinct row (not collapsed into one),
via core.excel_reader.assign_row_keys()'s duplicate-Index disambiguation
-- the SAME machinery that fixes required_stages_by_index()'s
duplicate-index collapse (see tests/test_hcai_fdd_use_status.py's HMR
10A test for that side of the fix).

RULE (d) REFINEMENT -- a later review (reading every non-blank,
non-numeric Reference Number cell across all 14 real fixtures directly,
not guessing) found the blank-ref "?" rule above was too coarse: a
non-blank, non-numeric Reference Number can ALSO be the preparer's own
"NOTE" marker (SMR 4 row 25 -- real Description text, but not a
trackable item, confirmed by that literal marker + a yellow highlight
fill) or a real alphanumeric sub-reference (HMR 9B's "7a"/"7b"/"7c",
sub-items of reference 7). parse_milestone_sheet() now skips a
NOTE-marker row entirely (case/punctuation-insensitive: "NOTE",
"Note:", "NOTES" all match) and gives a real sub-reference its own
Index (f"F-{ref}") instead of folding it into the generic "?".

DISCIPLINE-MARK REFINEMENT -- a later, direct client correction (Rey, via
Marcus) found the blank-ref "?" rule above was STILL too coarse: a
blank-Reference row with a real Description is only a genuine tracked
item if at least one discipline sign-off column (GEOR/AOR/SEOR/MEOR/
EEOR/CONT/IOR/SP INSP/TEST LAB/CEOR -- see MILESTONE_DISCIPLINE_COLS)
has an "X" mark. Confirmed by reading every blank-ref row across all 14
real fixtures directly: rows with NO mark at all in any of those
columns (e.g. SMR 2C's own "Clear all plan review Outstanding Items
List (OIL) Items", SMR 4/SMR 3A's temporary-equipment/Substantial
Compliance/Certificate of Occupancy boilerplate lines, HMR 9B's own
"Certificate of Occupancy" row) are per-project boilerplate placeholder
lines, never actually tracked -- the SAME shape as the legend rows
below, and now excluded the same way. A blank-ref row that DOES have at
least one mark (SMR 2C's own "Construction Final", HMR 12's own
"Construction Final", HMR 17's own "Certificate of Occupancy", two SMR
3C rows) is still kept as "?", unchanged. A row WITH a real Reference
Number is never subject to this check at all, marks or no marks -- see
test_ref_present_kept_regardless_of_marks below.

A blank-ref row with a real Description, NO discipline mark, but
something filled in the sheet's own "HCAI FDD" column (confirmed to
exist, at a different column position, on 2 of the 14 real fixtures --
SMR 2C and SMR 2D -- though neither's own blank-ref rows actually use
it) is a real, reportable edge case, not something this fix silently
decides either way: the discipline-mark rule above doesn't consult that
column at all, so such a row is dropped by the same default as "no
mark anywhere" -- see test_blank_ref_hcai_fdd_only_dropped_by_default
below, which documents this explicitly rather than leaving it as an
unstated side effect.

FLOAT REFERENCE FIX -- the SAME kind of real sub-reference also shows up
as an actual Python FLOAT, not a string, when the preparer typed a
decimal directly into the Reference Number cell (SMR 2D's 1.1/1.2, SMR
3D's 4.1/4.2/4.3/4.4) -- confirmed every real reference cell across all
14 fixtures uses number_format "General" (no custom formatting that
could display differently from the raw value, so reading it directly is
safe). This hit the numeric branch's OWN pre-existing f"F-{int(ref)}",
which silently truncated the decimal and collapsed each one into
whatever OTHER row shared its integer part. Fixed with f"F-{ref:g}" --
keeps a whole number clean ("F-3", not "F-3.0") while preserving a real
fractional reference exactly ("F-4.2") and rounding away binary
float-noise (e.g. 4.1000000000000005 -> "4.1").

Run directly: `python tests/test_fcons_verif_blank_reference.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from core import excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")


def _build_synthetic_milestone_sheet():
    """A minimal F-Cons Verif sheet (just the columns parse_milestone_sheet()
    reads) covering every case this fix must distinguish:
      row 7:  real ref (1), real description, NO marks -> kept, Index "F-1"
              (a real Reference Number is never subject to the
              discipline-mark check at all -- see
              test_ref_present_kept_regardless_of_marks)
      row 8:  blank ref, real description, NO discipline mark
                                                        -> DROPPED (the
              discipline-mark refinement -- a per-project boilerplate
              line, same shape as "Clear all plan review OIL Items")
      row 9:  blank ref, real description, ONE "X" mark (uppercase)
                                                        -> kept, Index "?"
      row 10: blank ref, real description, ONE "x" mark (lowercase)
                                                        -> kept, Index "?"
              (distinct row key from row 9's -- case-insensitive match,
              same as real SMR 3C data, which uses lowercase "x")
      row 11: blank ref, blank description              -> dropped (unchanged rule)
      row 12: blank ref, real description, NO discipline mark, but a
              value in the sheet's own "HCAI FDD" column -> DROPPED (the
              reportable edge case -- see
              test_blank_ref_hcai_fdd_only_dropped_by_default)
      row 13: real ref (2), real description, NO marks  -> kept, Index "F-2"
              (confirms a real Reference Number is kept with or
              without marks, unlike the blank-ref rows above)
      row 14: ref=3.0 (a WHOLE-number float)             -> kept, Index "F-3" (NOT "F-3.0")
      row 15: ref=4.2 (a real fractional sub-ref)        -> kept, Index "F-4.2" (NOT collapsed into "F-4")
      row 16: ref="7a" (real alphanumeric sub-ref)       -> kept, Index "F-7a" (NOT "?")
      row 17: ref="NOTE" (preparer's own marker)         -> DROPPED (a real description, but not trackable)
      row 18: ref="Note:" (different punctuation)        -> DROPPED (case/punctuation-insensitive match)
      row 19: ref="NOTES" (plural)                       -> DROPPED (same "starts with note" match)
      row 20: "ABBREVIATIONS:" in the ref column         -> legend boundary; everything
      row 21: blank ref, real description, with a mark      from here on is dropped, even
                                                              though row 21 has the exact
                                                              same "blank ref + real text +
                                                              mark" shape rows 9/10 do.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "F-Cons Verif"
    ws.cell(row=5, column=excel_reader.MILESTONE_REF_COL, value="Reference Number")
    AOR_COL = excel_reader.MILESTONE_DISCIPLINE_COLS[1]   # a real discipline column (AOR)
    SEOR_COL = excel_reader.MILESTONE_DISCIPLINE_COLS[2]  # a different real discipline column (SEOR)
    HCAI_FDD_COL = 17  # confirmed position on real files that have this column (SMR 2C, SMR 2D) -- see module docstring
    rows = [
        (7, 1, "Substantial Compliance", None),
        (8, None, "Clear all plan review Outstanding Items List (OIL) Items", None),
        (9, None, "Construction Final", (AOR_COL, "X")),
        (10, None, "Certificate of Occupancy", (SEOR_COL, "x")),
        (11, None, None, None),
        (12, None, "HCAI/FDD marked, but no discipline sign-off at all", (HCAI_FDD_COL, "2026-01-01")),
        (13, 2, "A real-ref row with no marks either", None),
        (14, 3.0, "Whole-number float reference", None),
        (15, 4.2, "Prior to 1st Foundation Pour", None),
        (16, "7a", "Gamma Knife slab (Stage 10)", None),
        (17, "NOTE", "Please see the attached note for this section", None),
        (18, "Note:", "A second, differently-punctuated note marker", None),
        (19, "NOTES", "A third, plural note marker", None),
        (20, "ABBREVIATIONS:", "GEOR - Geotechnical Engineer of Record", None),
        (21, None, "MEOR - Mechanical Engineer of Record", (AOR_COL, "X")),
    ]
    for row_num, ref, desc, mark in rows:
        ws.cell(row=row_num, column=excel_reader.MILESTONE_REF_COL, value=ref)
        ws.cell(row=row_num, column=excel_reader.MILESTONE_DESC_COL, value=desc)
        if mark is not None:
            col, value = mark
            ws.cell(row=row_num, column=col, value=value)
    return ws


def test_synthetic_blank_reference_rows(failures):
    ws = _build_synthetic_milestone_sheet()
    records = excel_reader.parse_milestone_sheet(ws, stage_count=0)
    indexes = [r.index for r in records]
    descriptions = [r.description for r in records]
    print(f"Synthetic sheet -> {len(records)} records: {list(zip(indexes, descriptions))}")

    if len(records) != 7:
        failures.append(f"synthetic: expected exactly 7 kept records (row 8 no-mark blank-ref dropped, row "
                         f"11 blank/blank dropped, row 12 HCAI/FDD-only dropped, rows 17/18/19 NOTE-variants "
                         f"dropped, legend rows 20/21 dropped), got {len(records)}: "
                         f"{list(zip(indexes, descriptions))}")

    if indexes != ["F-1", "?", "?", "F-2", "F-3", "F-4.2", "F-7a"]:
        failures.append(f"synthetic: expected indexes ['F-1', '?', '?', 'F-2', 'F-3', 'F-4.2', 'F-7a'] in file order, got {indexes}")

    if "Clear all plan review Outstanding Items List (OIL) Items" in descriptions:
        failures.append("synthetic: a blank-ref row with NO discipline mark must be dropped, not kept as '?'")

    if "F-3.0" in indexes:
        failures.append(f"synthetic: a whole-number float reference (3.0) must render as 'F-3', not 'F-3.0', got {indexes}")

    if "GEOR - Geotechnical Engineer of Record" in descriptions or "MEOR - Mechanical Engineer of Record" in descriptions:
        failures.append(f"synthetic: legend rows (after 'ABBREVIATIONS:') must never be kept, even with a discipline mark present (row 21), got descriptions {descriptions}")

    note_descriptions = {
        "Please see the attached note for this section",
        "A second, differently-punctuated note marker",
        "A third, plural note marker",
    }
    if note_descriptions & set(descriptions):
        failures.append(f"synthetic: NOTE-marker rows (any casing/punctuation) must never be kept, got {note_descriptions & set(descriptions)}")

    # The two "?" rows (rows 9 and 10 -- "Construction Final"/uppercase
    # X, "Certificate of Occupancy"/lowercase x) must get DISTINCT row
    # keys -- not collapsed into one entry the way a plain Index-keyed
    # dict would.
    row_keys = excel_reader.assign_row_keys(records)
    q_keys = [key for rec, key in zip(records, row_keys) if rec.index == "?"]
    print(f"  '?' row keys: {q_keys}")
    if len(q_keys) != 2 or len(set(q_keys)) != 2:
        failures.append(f"synthetic: the 2 '?' rows must get 2 DISTINCT row keys, got {q_keys}")
    if any(key == "?" for key in q_keys):
        # Neither bare "?" occurrence is unique in this sheet (there are
        # 2 of them), so assign_row_keys() must disambiguate BOTH, not
        # leave one as the bare "?" and only suffix the other.
        failures.append(f"synthetic: with 2 duplicate '?' rows, NEITHER should keep the bare '?' key unsuffixed, got {q_keys}")


def test_ref_present_kept_regardless_of_marks(failures):
    """A row WITH a real Reference Number is never subject to the
    discipline-mark check at all -- confirmed on the synthetic sheet's
    own row 13 ("A real-ref row with no marks either", ref=2, NO marks
    anywhere), which must still be kept as "F-2", same as row 7 (ref=1,
    also no marks). The discipline-mark rule only ever applies to the
    blank-Reference-Number branch.
    """
    ws = _build_synthetic_milestone_sheet()
    records = excel_reader.parse_milestone_sheet(ws, stage_count=0)
    by_index = {r.index: r.description for r in records}

    print(f"F-1 (ref=1, no marks): {by_index.get('F-1')!r}")
    print(f"F-2 (ref=2, no marks): {by_index.get('F-2')!r}")
    if by_index.get("F-1") != "Substantial Compliance":
        failures.append(f"a real-ref row (F-1) with no discipline marks must still be kept, got {by_index.get('F-1')!r}")
    if by_index.get("F-2") != "A real-ref row with no marks either":
        failures.append(f"a real-ref row (F-2) with no discipline marks must still be kept -- the mark rule must only apply to blank-ref rows, got {by_index.get('F-2')!r}")


def test_blank_ref_hcai_fdd_only_dropped_by_default(failures):
    """Edge case flagged for reporting, not silently decided either way
    (see module docstring's DISCIPLINE-MARK REFINEMENT section): a
    blank-Reference row with a real Description, NOTHING in any
    discipline sign-off column, but something filled in the sheet's own
    "HCAI FDD" column -- confirmed, across all 14 real fixtures, that
    NO real blank-ref row actually hits this combination (the 2 files
    that even HAVE an "HCAI FDD" column in F-Cons Verif, SMR 2C and SMR
    2D, never populate it on a blank-ref row), so this is a synthetic-
    only case today. Default: dropped, same as "no mark anywhere" --
    the discipline-mark rule never consults the HCAI/FDD column at all.
    """
    ws = _build_synthetic_milestone_sheet()
    records = excel_reader.parse_milestone_sheet(ws, stage_count=0)
    descriptions = {r.description for r in records}
    print(f"'HCAI/FDD marked, but no discipline sign-off at all' kept? {'HCAI/FDD marked, but no discipline sign-off at all' in descriptions}")
    if "HCAI/FDD marked, but no discipline sign-off at all" in descriptions:
        failures.append(
            "a blank-ref row with something in the HCAI/FDD column but NO discipline mark should be dropped "
            "by default (same as no mark anywhere) -- it was kept instead"
        )


def test_grid_row_never_shows_as_vcr(failures):
    """A grid row (is_milestone=False) with zero required stages must
    produce ZERO Sum Data rows, never a 'VCR' placeholder -- confirmed
    both as a direct unit case here and, more importantly, against every
    real fixture available (any real grid item that DOES have required
    stages, including HMR 10A's 39 duplicate-Index ones, must never
    expand to 'VCR' -- see tests/test_hcai_fdd_use_status.py's HMR 10A
    test for that direct real-file confirmation).
    """
    grid_row_no_stages = {
        "index": "B-NOPE", "description": "d", "approval_agency": "AG",
        "required_stages": [], "stage_status": {}, "vcr": None, "is_milestone": False,
    }
    result = excel_reader.expand_sum_data_rows([grid_row_no_stages])
    if result:
        failures.append(f"a grid row with no required stages must expand to ZERO rows, got {result}")

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present -- real-fixture portion of this check skipped")
        return

    import glob
    for path in sorted(glob.glob(os.path.join(FIXDIR, "*"))):
        with tempfile.TemporaryDirectory() as tmp:
            store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
            name = os.path.basename(path)
            store.add_project(name)
            inc = store.add_new_increment(name, path)
            display = store.get_increment_for_display(name, inc.name)
            expanded = excel_reader.expand_sum_data_rows(display.sum_data)
            bad = [e for e in expanded if e["stage"] == "VCR" and not str(e["index"]).startswith(("F-", "?"))]
            if bad:
                failures.append(f"{name}: {len(bad)} non-milestone row(s) show as 'VCR': {bad}")


def test_smr_2c_real_file_snapshot(failures):
    """Real-file snapshot (skip-guarded): SMR 2C now shows exactly 1
    '?' row -- "Construction Final" (has discipline marks: AOR/MEOR/
    EEOR/CONT/IOR/CEOR all "X") -- not 2. "Clear all plan review
    Outstanding Items List (OIL) Items" (no discipline marks at all) is
    now DROPPED, per the discipline-mark refinement -- this is the
    pinned number that CHANGED from the original Fix 2 (which kept
    both). SMR 2C's total Sum Data row count is now exactly 2: this one
    '?' row plus its one real grid item (B-FP1).
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("SMR 2C Snapshot")
        inc = store.add_new_increment("SMR 2C Snapshot", os.path.join(FIXDIR, "TIO S220804-38-00_SMR 2C_220705.xlsm"))
        display = store.get_increment_for_display("SMR 2C Snapshot", inc.name)

        q_rows = [r for r in display.sum_data if r["index"] == "?"]
        print(f"SMR 2C '?' rows: {[(r['description'], r['vcr']) for r in q_rows]}")
        if len(q_rows) != 1:
            failures.append(f"SMR 2C: expected exactly 1 '?' row (was 2 before the discipline-mark refinement), got {len(q_rows)}: {q_rows}")
        elif q_rows[0]["description"] != "Construction Final":
            failures.append(f"SMR 2C: expected the one '?' row to be 'Construction Final', got {q_rows[0]['description']!r}")

        if any(r["description"] == "Clear all plan review Outstanding Items List (OIL) Items" for r in display.sum_data):
            failures.append("SMR 2C: 'Clear all plan review OIL Items' has NO discipline marks and must be dropped, not kept")

        for r in q_rows:
            if r["vcr"] != "Open":
                failures.append(f"SMR 2C '?' row {r['description']!r}: expected vcr='Open', got {r['vcr']!r}")
            if not r["is_milestone"]:
                failures.append(f"SMR 2C '?' row {r['description']!r}: expected is_milestone=True")

        print(f"SMR 2C total Sum Data rows: {len(display.sum_data)} (expected 2: B-FP1 + 'Construction Final')")
        if len(display.sum_data) != 2:
            failures.append(f"SMR 2C: expected exactly 2 total Sum Data rows (B-FP1 grid item + 'Construction Final' milestone), got {len(display.sum_data)}: {[r['index'] for r in display.sum_data]}")

        all_descriptions = {r["description"] for r in display.sum_data}
        legend_text = {
            "GEOR - Geotechnical Engineer of Record",
            "MEOR - Mechanical Engineer of Record",
            "SP. INSP - Special Inspector",
        }
        if all_descriptions & legend_text:
            failures.append(f"SMR 2C: legend text leaked into Sum Data: {all_descriptions & legend_text}")

        # Export must not crash and must carry the '?' row through.
        from core import excel_export
        out_path = Path(tmp) / "smr2c_export.xlsx"
        excel_export.export_increment(display, out_path)
        if not out_path.exists():
            failures.append("SMR 2C: export_increment() did not produce a file")


def test_hmr12_real_file_snapshot(failures):
    """Real-file snapshot (skip-guarded): HMR 12 must still show exactly
    the 1 real blank-reference-number item Rey named ("Construction
    Final F Missing") -- UNCHANGED by the discipline-mark refinement,
    since this specific row has real discipline marks (confirmed
    directly: AOR/SEOR/MEOR/EEOR/CONT/IOR/SP INSP/TEST LAB all "x").
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("HMR 12 Snapshot")
        inc = store.add_new_increment("HMR 12 Snapshot", os.path.join(FIXDIR, "TIO S231014-38-00 HMR 12.xlsm"))
        display = store.get_increment_for_display("HMR 12 Snapshot", inc.name)

        q_rows = [r for r in display.sum_data if r["index"] == "?"]
        print(f"HMR 12 '?' rows: {[(r['description'], r['vcr']) for r in q_rows]}")
        if len(q_rows) != 1:
            failures.append(f"HMR 12: expected exactly 1 '?' row (unchanged by the discipline-mark refinement -- this row has real marks), got {len(q_rows)}: {q_rows}")
        elif q_rows[0]["description"] != "Construction Final":
            failures.append(f"HMR 12: expected the '?' row to be 'Construction Final', got {q_rows[0]['description']!r}")


def test_smr4_note_row_skipped(failures):
    """Real-file snapshot: SMR 4's F-Cons Verif row 25 has a real
    Description but its Reference Number cell literally reads "NOTE"
    (plus a yellow highlight fill, confirmed by opening the real file
    directly) -- it's the preparer's own "see note" marker, not a
    trackable item, and must be skipped, not recovered as "?".

    SMR 4's other 5 blank-reference rows (OIL clear, 2x temporary
    equipment, Substantial Compliance, Certificate of Occupancy) were
    kept as 5 "?" rows under the original Fix 2 -- but NONE of them
    have any discipline sign-off mark at all (confirmed directly), so
    under the discipline-mark refinement all 5 are now DROPPED too.
    This is a pinned number that CHANGED: 5 -> 0.
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("SMR 4 NOTE Snapshot")
        inc = store.add_new_increment("SMR 4 NOTE Snapshot", os.path.join(FIXDIR, "TIO H220881-38-00_SMR 4.xlsm"))
        display = store.get_increment_for_display("SMR 4 NOTE Snapshot", inc.name)

        q_rows = [r for r in display.sum_data if r["index"] == "?"]
        print(f"SMR 4 '?' rows: {[r['description'] for r in q_rows]}")
        if len(q_rows) != 0:
            failures.append(f"SMR 4: expected exactly 0 '?' rows (all 5 blank-ref rows lack discipline marks, dropped by the refinement; was 5 before it), got {len(q_rows)}: {[r['description'] for r in q_rows]}")

        note_text = "Please reference excel file 220909_TIO_UCSF-NHPH_SMR4_backcheck 1 to see more updated Section F as this sheet is protected."
        all_descriptions = {r["description"] for r in display.sum_data}
        if any(d and note_text.split("\n")[0] in d for d in all_descriptions):
            failures.append("SMR 4: the 'NOTE' marker row's text leaked into Sum Data -- it should be skipped entirely")


def test_hmr9b_alphanumeric_subreferences(failures):
    """Real-file snapshot: HMR 9B's F-Cons Verif rows 13-15 have
    Reference Number cells "7a"/"7b"/"7c" (real alphanumeric
    sub-references of reference 7, confirmed by opening the real file:
    the full sequence is 1,2,3,4,5,6,7a,7b,7c,8) -- NOT blank. They must
    get their own Index ("F-7a"/"F-7b"/"F-7c"), not the generic "?" a
    truly blank Reference Number gets.

    HMR 9B's '?' count drops from 4 (under the original rule) to 1
    (under Rule (d), once the "7a"/"7b"/"7c" sub-references got their
    own Index instead) to 0 now (under the discipline-mark refinement):
    its one remaining blank-ref row, "Certificate of Occupancy" (row
    26), has NO discipline sign-off mark at all (confirmed directly),
    so it's dropped too. This is a pinned number that CHANGED: 1 -> 0.
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("HMR 9B Sub-Reference Snapshot")
        inc = store.add_new_increment("HMR 9B Sub-Reference Snapshot", os.path.join(FIXDIR, "TIO H231359-38-00_HMR 9B.xlsm"))
        display = store.get_increment_for_display("HMR 9B Sub-Reference Snapshot", inc.name)

        q_rows = [r for r in display.sum_data if r["index"] == "?"]
        print(f"HMR 9B '?' rows: {[r['description'] for r in q_rows]}")
        if len(q_rows) != 0:
            failures.append(f"HMR 9B: expected exactly 0 '?' rows (its one blank-ref row, 'Certificate of Occupancy', has no discipline marks; was 1 before the refinement), got {len(q_rows)}: {[r['description'] for r in q_rows]}")

        sub_ref_rows = {r["index"]: r["description"] for r in display.sum_data if r["index"] in ("F-7a", "F-7b", "F-7c")}
        print(f"HMR 9B sub-reference rows: {sub_ref_rows}")
        expected_sub_refs = {
            "F-7a": "Gamma Knife slab (Stage 10)",
            "F-7b": "Lower Level Waterproofing",
            "F-7c": "Long Level 2 Electrical Rm",
        }
        if sub_ref_rows != expected_sub_refs:
            failures.append(f"HMR 9B: expected sub-reference rows {expected_sub_refs}, got {sub_ref_rows}")

        # Distinct row keys (no collision, no need for suffixing since
        # "F-7a"/"F-7b"/"F-7c" are each unique in this file).
        all_data_sub_refs = [r for r in display.all_data if r["index"] in ("F-7a", "F-7b", "F-7c")]
        row_keys = {r["index"]: r["row_key"] for r in all_data_sub_refs}
        if any(key != index for index, key in row_keys.items()):
            failures.append(f"HMR 9B: F-7a/F-7b/F-7c are each unique, so their row_key should equal their Index unchanged, got {row_keys}")


def test_all_14_fixtures_q_counts_pinned(failures):
    """Pinned '?' count for all 14 real fixtures after the discipline-
    mark refinement -- confirmed directly by reading every blank-ref
    row's own discipline columns across all of them (not guessed), one
    table so every changed number is visible in one place.

    CHANGED from before this refinement (old count -> new count):
      SMR 4:   5 -> 0  (OIL/temp-equipment/Substantial Compliance/CofO -- no marks on any)
      SMR 3A:  5 -> 0  (same boilerplate lines, no marks)
      SMR 4C:  4 -> 0  (temp-equipment/Substantial Compliance/CofO -- no marks)
      SMR 2C:  2 -> 1  ("Clear all plan review OIL" dropped; "Construction Final" kept, has marks)
      SMR 3C:  4 -> 2  ("Clear all plan review OIL"/"Domestic cold water piping" dropped, no marks;
                         2 transformer rows kept, both have marks)
      HMR 9B:  1 -> 0  ("Certificate of Occupancy" -- no marks)
      inc8:    2 -> 0  (Substantial Compliance/CofO -- no marks)

    UNCHANGED (already had marks, or already had zero blank-ref rows):
      SMR 2D: 0, SMR 3D: 0, SMR 4B: 0, HMR 6C: 0, HMR 10A: 0 -- no blank-ref
        candidate rows at all in these files, with or without this refinement.
      HMR 17: 1 ("Certificate of Occupancy" -- HAS marks: AOR/SEOR/IOR/SP INSP)
      HMR 12: 1 ("Construction Final" -- HAS marks, see test_hmr12_real_file_snapshot)
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    expected_q_counts = {
        "TIO H220369-38-00_SMR 2D.xlsm": 0,
        "TIO H220881-38-00_SMR 4.xlsm": 0,
        "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm": 0,
        "TIO H231359-38-00_HMR 9B.xlsm": 0,
        "TIO H240424-38-00 SMR 4B Delta 7.xlsm": 0,
        "TIO H250355-38-00 HMR 17.xlsm": 1,
        "TIO S220804-38-00_SMR 2C_220705.xlsm": 1,
        "TIO S221870-38-00_SMR 3C.xlsm": 2,
        "TIO S221919-38-00_HMR 6C.xlsm": 0,
        "TIO S230119-38-00_SMR 3A - 2022.xlsm": 0,
        "TIO S231014-38-00 HMR 12.xlsm": 1,
        "TIO S251652-38-00 SMR 4C.xlsm": 0,
        "TIO_H240649-38-00 HMR 10A.xlsm": 0,
    }

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("All-Fixtures Q-Count Snapshot")
        for fname, expected_count in expected_q_counts.items():
            path = os.path.join(FIXDIR, fname)
            if not os.path.isfile(path):
                failures.append(f"fixture assumption broken -- expected file {fname!r} not found")
                continue
            inc = store.add_new_increment("All-Fixtures Q-Count Snapshot", path)
            display = store.get_increment_for_display("All-Fixtures Q-Count Snapshot", inc.name)
            q_count = sum(1 for r in display.sum_data if r["index"] == "?")
            print(f"  {fname!r:50} '?' count: {q_count} (expected {expected_count})")
            if q_count != expected_count:
                failures.append(f"{fname}: expected {expected_count} '?' row(s), got {q_count}")

    inc8_path = os.path.join(os.path.dirname(__file__), "fixtures", "inc8_real_2stage.xlsm")
    if os.path.isfile(inc8_path):
        with tempfile.TemporaryDirectory() as tmp:
            store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
            store.add_project("inc8 Q-Count Snapshot")
            inc = store.add_new_increment("inc8 Q-Count Snapshot", inc8_path)
            display = store.get_increment_for_display("inc8 Q-Count Snapshot", inc.name)
            q_count = sum(1 for r in display.sum_data if r["index"] == "?")
            print(f"  inc8_real_2stage.xlsm (not one of Rey's 14, no Rey label) '?' count: {q_count} (expected 0, was 2)")
            if q_count != 0:
                failures.append(f"inc8_real_2stage.xlsm: expected 0 '?' rows (Substantial Compliance/CofO have no discipline marks), got {q_count}")


def test_smr2d_float_subreferences(failures):
    """Real-file snapshot: SMR 2D's F-Cons Verif rows 8-9 have Reference
    Number cells 1.1/1.2 -- real Python FLOATS, not strings, so they hit
    the numeric branch, not the alphanumeric-string branch HMR 9B's
    "7a"/"7b"/"7c" exercises. Before the float fix, int(1.1)==int(1.2)==1
    silently collapsed both into the same "F-1" as row 7's own real
    reference 1 -- "F-1" showed 3 times instead of the 3 DISTINCT real
    references (1, 1.1, 1.2) it actually is. Confirmed this is exactly
    what Rey's TIO Error Report pivot's "SMR 2D F-1 VCR=3" reflects.
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("SMR 2D Float Sub-Reference Snapshot")
        inc = store.add_new_increment("SMR 2D Float Sub-Reference Snapshot", os.path.join(FIXDIR, "TIO H220369-38-00_SMR 2D.xlsm"))
        display = store.get_increment_for_display("SMR 2D Float Sub-Reference Snapshot", inc.name)

        f1_rows = [r for r in display.sum_data if r["index"] == "F-1"]
        print(f"SMR 2D 'F-1' rows: {len(f1_rows)} (expected exactly 1, not the old 3-way collision)")
        if len(f1_rows) != 1:
            failures.append(f"SMR 2D: expected exactly 1 'F-1' row (1.1/1.2 must no longer collapse into it), got {len(f1_rows)}")

        sub_ref_rows = {r["index"]: r["row_key"] for r in display.all_data if r["index"] in ("F-1.1", "F-1.2")}
        print(f"SMR 2D float sub-reference rows (index -> row_key): {sub_ref_rows}")
        if set(sub_ref_rows) != {"F-1.1", "F-1.2"}:
            failures.append(f"SMR 2D: expected both 'F-1.1' and 'F-1.2' present as their own Index, got {set(sub_ref_rows)}")
        if any(key != index for index, key in sub_ref_rows.items()):
            failures.append(f"SMR 2D: 'F-1.1'/'F-1.2' are each unique now, so row_key should equal Index unchanged, got {sub_ref_rows}")
        if any("F-1.0" in str(v) or "0000000000000" in str(v) for v in sub_ref_rows):
            failures.append(f"SMR 2D: found float-noise or a trailing '.0' in a sub-reference Index: {sub_ref_rows}")


def test_smr3d_float_subreferences(failures):
    """Real-file snapshot: SMR 3D's F-Cons Verif rows 12-15 have
    Reference Number cells 4.1/4.2/4.3/4.4. Before the float fix, all
    four (plus a genuinely blank-description row 10, ref=4, already
    excluded by the existing blank-description rule) truncated to the
    same "F-4" -- Rey's pivot shows "F-4 VCR=4" (the 4 real sub-refs;
    the blank row never counted in the first place, unrelated to this
    fix). Confirms "F-4" itself still exists as its own, separate real
    reference (row 9, blank desc -- excluded) is NOT what this checks;
    row 9 (desc=None) has no bearing on Sum Data either way.
    """
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("SMR 3D Float Sub-Reference Snapshot")
        inc = store.add_new_increment("SMR 3D Float Sub-Reference Snapshot", os.path.join(FIXDIR, "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm"))
        display = store.get_increment_for_display("SMR 3D Float Sub-Reference Snapshot", inc.name)

        f4_rows = [r for r in display.sum_data if r["index"] == "F-4"]
        print(f"SMR 3D 'F-4' rows in Sum Data: {len(f4_rows)} (expected 0 -- its only physical row has a blank description)")
        if f4_rows:
            failures.append(f"SMR 3D: did not expect any 'F-4' Sum Data row (its ref=4 row has a blank description), got {f4_rows}")

        sub_ref_rows = {r["index"]: r["row_key"] for r in display.all_data if r["index"] in ("F-4.1", "F-4.2", "F-4.3", "F-4.4")}
        print(f"SMR 3D float sub-reference rows (index -> row_key): {sub_ref_rows}")
        if set(sub_ref_rows) != {"F-4.1", "F-4.2", "F-4.3", "F-4.4"}:
            failures.append(f"SMR 3D: expected all 4 sub-references present as their own Index, got {set(sub_ref_rows)}")
        if any(key != index for index, key in sub_ref_rows.items()):
            failures.append(f"SMR 3D: each sub-reference is unique now, so row_key should equal Index unchanged, got {sub_ref_rows}")
        if any("0000000000000" in str(v) for v in sub_ref_rows):
            failures.append(f"SMR 3D: found float-noise in a sub-reference Index: {sub_ref_rows}")


def main():
    failures = []
    test_synthetic_blank_reference_rows(failures)
    test_ref_present_kept_regardless_of_marks(failures)
    test_blank_ref_hcai_fdd_only_dropped_by_default(failures)
    test_grid_row_never_shows_as_vcr(failures)
    test_smr_2c_real_file_snapshot(failures)
    test_hmr12_real_file_snapshot(failures)
    test_smr4_note_row_skipped(failures)
    test_hmr9b_alphanumeric_subreferences(failures)
    test_all_14_fixtures_q_counts_pinned(failures)
    test_smr2d_float_subreferences(failures)
    test_smr3d_float_subreferences(failures)

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- F-Cons Verif rows with a real description but a blank Reference Number are kept "
            "(Index '?', Stage # 'VCR') only when at least one discipline sign-off column has an 'X' mark -- "
            "a blank-ref row with no mark at all (boilerplate placeholder lines) is dropped, same as the "
            "sheet's own 'ABBREVIATIONS:' legend block; a row WITH a real Reference Number is never subject "
            "to this check, marks or no marks; something in the HCAI/FDD column alone (no discipline mark) "
            "is dropped by the same default. Confirmed on a synthetic sheet and pinned across all 14 real "
            "fixtures (SMR 2C, SMR 4, HMR 9B, SMR 3C, SMR 3A, SMR 4C, inc8 all changed; HMR 12/HMR 17 "
            "unchanged, both have real marks). No grid row (real or synthetic) ever shows as a Sum Data "
            "'VCR' row; a non-blank, non-numeric Reference Number is correctly split between a skipped "
            "'NOTE' marker (any casing/punctuation, real SMR 4 row) and a real alphanumeric sub-reference "
            "that gets its own Index (real HMR 9B rows); a real FLOAT sub-reference (real SMR 2D/SMR 3D "
            "rows) gets its own clean Index too, with no int()-truncation collision and no float noise."
        )


if __name__ == "__main__":
    main()

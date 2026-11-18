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
      row 7:  real ref (1), real description         -> kept, Index "F-1"
      row 8:  blank ref, real description             -> kept, Index "?"
      row 9:  blank ref, real description (2nd one)   -> kept, Index "?" (distinct row key from row 8's)
      row 10: blank ref, blank description             -> dropped (unchanged rule)
      row 11: real ref (2), real description          -> kept, Index "F-2"
      row 12: ref=3.0 (a WHOLE-number float)           -> kept, Index "F-3" (NOT "F-3.0")
      row 13: ref=4.2 (a real fractional sub-ref)      -> kept, Index "F-4.2" (NOT collapsed into "F-4")
      row 14: ref="7a" (real alphanumeric sub-ref)     -> kept, Index "F-7a" (NOT "?")
      row 15: ref="NOTE" (preparer's own marker)       -> DROPPED (a real description, but not trackable)
      row 16: ref="Note:" (different punctuation)      -> DROPPED (case/punctuation-insensitive match)
      row 17: ref="NOTES" (plural)                     -> DROPPED (same "starts with note" match)
      row 18: "ABBREVIATIONS:" in the ref column       -> legend boundary; everything
      row 19: blank ref, real description (legend)        from here on is dropped, even
                                                            though row 19 has the exact
                                                            same "blank ref + real text" shape
                                                            rows 8/9 do.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "F-Cons Verif"
    ws.cell(row=5, column=excel_reader.MILESTONE_REF_COL, value="Reference Number")
    rows = [
        (7, 1, "Substantial Compliance"),
        (8, None, "Clear all plan review Outstanding Items List (OIL) Items"),
        (9, None, "Construction Final"),
        (10, None, None),
        (11, 2, "Certificate of Occupancy"),
        (12, 3.0, "Whole-number float reference"),
        (13, 4.2, "Prior to 1st Foundation Pour"),
        (14, "7a", "Gamma Knife slab (Stage 10)"),
        (15, "NOTE", "Please see the attached note for this section"),
        (16, "Note:", "A second, differently-punctuated note marker"),
        (17, "NOTES", "A third, plural note marker"),
        (18, "ABBREVIATIONS:", "GEOR - Geotechnical Engineer of Record"),
        (19, None, "MEOR - Mechanical Engineer of Record"),
    ]
    for row_num, ref, desc in rows:
        ws.cell(row=row_num, column=excel_reader.MILESTONE_REF_COL, value=ref)
        ws.cell(row=row_num, column=excel_reader.MILESTONE_DESC_COL, value=desc)
    return ws


def test_synthetic_blank_reference_rows(failures):
    ws = _build_synthetic_milestone_sheet()
    records = excel_reader.parse_milestone_sheet(ws, stage_count=0)
    indexes = [r.index for r in records]
    descriptions = [r.description for r in records]
    print(f"Synthetic sheet -> {len(records)} records: {list(zip(indexes, descriptions))}")

    if len(records) != 7:
        failures.append(f"synthetic: expected exactly 7 kept records (row 10 blank/blank dropped, rows "
                         f"15/16/17 NOTE-variants dropped, legend rows 18/19 dropped), got {len(records)}: "
                         f"{list(zip(indexes, descriptions))}")

    if indexes != ["F-1", "?", "?", "F-2", "F-3", "F-4.2", "F-7a"]:
        failures.append(f"synthetic: expected indexes ['F-1', '?', '?', 'F-2', 'F-3', 'F-4.2', 'F-7a'] in file order, got {indexes}")

    if "F-3.0" in indexes:
        failures.append(f"synthetic: a whole-number float reference (3.0) must render as 'F-3', not 'F-3.0', got {indexes}")

    if "GEOR - Geotechnical Engineer of Record" in descriptions or "MEOR - Mechanical Engineer of Record" in descriptions:
        failures.append(f"synthetic: legend rows (after 'ABBREVIATIONS:') must never be kept, got descriptions {descriptions}")

    note_descriptions = {
        "Please see the attached note for this section",
        "A second, differently-punctuated note marker",
        "A third, plural note marker",
    }
    if note_descriptions & set(descriptions):
        failures.append(f"synthetic: NOTE-marker rows (any casing/punctuation) must never be kept, got {note_descriptions & set(descriptions)}")

    # The two "?" rows (rows 8 and 9) must get DISTINCT row keys -- not
    # collapsed into one entry the way a plain Index-keyed dict would.
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
    """Real-file snapshot (skip-guarded): SMR 2C must now show exactly
    the 2 real blank-reference-number items Rey named, both "Open" (its
    F-Cons Verif sheet's VCR column (V) is genuinely blank for both --
    confirmed directly against the real file), and the legend rows must
    never appear.
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
        if len(q_rows) != 2:
            failures.append(f"SMR 2C: expected exactly 2 '?' rows, got {len(q_rows)}: {q_rows}")

        descriptions = {r["description"] for r in q_rows}
        expected_descriptions = {
            "Clear all plan review Outstanding Items List (OIL) Items",
            "Construction Final",
        }
        if descriptions != expected_descriptions:
            failures.append(f"SMR 2C: expected '?' descriptions {expected_descriptions}, got {descriptions}")

        for r in q_rows:
            if r["vcr"] != "Open":
                failures.append(f"SMR 2C '?' row {r['description']!r}: expected vcr='Open', got {r['vcr']!r}")
            if not r["is_milestone"]:
                failures.append(f"SMR 2C '?' row {r['description']!r}: expected is_milestone=True")

        all_descriptions = {r["description"] for r in display.sum_data}
        legend_text = {
            "GEOR - Geotechnical Engineer of Record",
            "MEOR - Mechanical Engineer of Record",
            "SP. INSP - Special Inspector",
        }
        if all_descriptions & legend_text:
            failures.append(f"SMR 2C: legend text leaked into Sum Data: {all_descriptions & legend_text}")

        # Export must not crash and must carry both '?' rows through.
        from core import excel_export
        out_path = Path(tmp) / "smr2c_export.xlsx"
        excel_export.export_increment(display, out_path)
        if not out_path.exists():
            failures.append("SMR 2C: export_increment() did not produce a file")


def test_hmr12_real_file_snapshot(failures):
    """Real-file snapshot (skip-guarded): HMR 12 must now show exactly
    the 1 real blank-reference-number item Rey named ("Construction
    Final F Missing").
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
            failures.append(f"HMR 12: expected exactly 1 '?' row, got {len(q_rows)}: {q_rows}")
        elif q_rows[0]["description"] != "Construction Final":
            failures.append(f"HMR 12: expected the '?' row to be 'Construction Final', got {q_rows[0]['description']!r}")


def test_smr4_note_row_skipped(failures):
    """Real-file snapshot: SMR 4's F-Cons Verif row 25 has a real
    Description but its Reference Number cell literally reads "NOTE"
    (plus a yellow highlight fill, confirmed by opening the real file
    directly) -- it's the preparer's own "see note" marker, not a
    trackable item, and must be skipped, not recovered as "?". SMR 4's
    other 5 blank-reference rows (OIL clear, 2x temporary equipment,
    Substantial Compliance, Certificate of Occupancy) ARE real and must
    still be kept -- confirmed count: 5, not 6.
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
        if len(q_rows) != 5:
            failures.append(f"SMR 4: expected exactly 5 '?' rows (NOTE row skipped), got {len(q_rows)}: {[r['description'] for r in q_rows]}")

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
    truly blank Reference Number gets. HMR 9B's '?' count drops from 4
    (under the earlier, less precise rule) to 1 -- only "Certificate of
    Occupancy" (row 26, genuinely blank Reference Number) is left.
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
        if len(q_rows) != 1:
            failures.append(f"HMR 9B: expected exactly 1 '?' row, got {len(q_rows)}: {[r['description'] for r in q_rows]}")
        elif q_rows[0]["description"] != "Certificate of Occupancy\n(New Buildings, Additions, Changes in Occupancy)":
            failures.append(f"HMR 9B: expected the one '?' row to be 'Certificate of Occupancy...', got {q_rows[0]['description']!r}")

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
    test_grid_row_never_shows_as_vcr(failures)
    test_smr_2c_real_file_snapshot(failures)
    test_hmr12_real_file_snapshot(failures)
    test_smr4_note_row_skipped(failures)
    test_hmr9b_alphanumeric_subreferences(failures)
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
            "(Index '?', distinct row keys for multiple '?' rows in the same increment) up to but not "
            "including the sheet's own 'ABBREVIATIONS:' legend block, confirmed on a synthetic sheet and on "
            "the real SMR 2C / HMR 12 files Rey named; no grid row (real or synthetic) ever shows as a Sum "
            "Data 'VCR' row; a non-blank, non-numeric Reference Number is correctly split between a skipped "
            "'NOTE' marker (any casing/punctuation, real SMR 4 row) and a real alphanumeric sub-reference "
            "that gets its own Index (real HMR 9B rows); a real FLOAT sub-reference (real SMR 2D/SMR 3D "
            "rows) gets its own clean Index too, with no int()-truncation collision and no float noise."
        )


if __name__ == "__main__":
    main()

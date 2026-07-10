"""
Real, file-backed data layer for the UI, wired to the actual backend:
core/project_store.py for persistence, core/excel_reader.py for parsing
workbooks, core/structure_diff.py for change detection, and
core/status_tracker.py for status carryover -- now at (Index, Stage)
granularity, matching how Sum Data's Done/Open status actually works.

The class is still named MockDataStore and this file is still named
mock_data.py for historical reasons -- this is where the UI-facing view
models (Project/Increment/ComparisonResult) were first defined, back when
everything here really was fake, and the UI modules already import from
here by that name. Nothing in this file is mock data anymore: every method
reads or writes real files under core.project_store.get_default_data_dir()
(or wherever the ProjectStore passed to MockDataStore is pointed).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd

from core import combined_export, excel_reader, increment_matcher, status_tracker, structure_diff, value_diff
from core.increment_matcher import MatchResult
from core.project_store import ProjectStore, VersionRecord

STAGE_COUNT = excel_reader.STAGE_COUNT


# ----------------------------------------------------------------------
# UI-facing view models
# ----------------------------------------------------------------------
@dataclass
class Increment:
    name: str
    version: int
    last_updated: str
    all_data: list[dict] = field(default_factory=list)
    # Sum Data row shape: {index, description, approval_agency,
    # required_stages, stage_status, vcr} -- "stage_status" is the SAME
    # dict object referenced by the matching all_data row (see
    # _sum_data_rows()), not a copy, so a click on the All Data tab is
    # immediately visible here too. Deliberately carries no precomputed
    # "live status" or "% complete" -- those depend on status.json state
    # that can change after this Increment was built, so
    # ui/pages/data_view_page.py derives them at render time instead of
    # trusting a value that could already be stale.
    sum_data: list[dict] = field(default_factory=list)
    # Report row shape: {approval_agency, index, description, total} --
    # static (doesn't depend on status.json), safe to compute once.
    report: list[dict] = field(default_factory=list)
    # J-Changes' state revision log, verbatim -- see
    # core.excel_reader.raw_changes_log. Row shape: {revision_number,
    # synopsis, aor_signature_date, seor_signature_date, effective_date,
    # hcai_concurrence}. Read-only reference data, same "doesn't feed All
    # Data/Sum Data/Report" scope boundary as the parser itself; static
    # per version (whatever this version's file's J-Changes sheet holds),
    # safe to compute once alongside report.
    changes_log: list[dict] = field(default_factory=list)
    # This app's own accumulated change_history.json entries (see
    # core.project_store.ProjectStore.load_change_history), OLDEST first
    # (append order) -- callers wanting newest-first (the Changes tab's
    # Update History section, and core.excel_export's Changes sheet,
    # which mirrors it) reverse this themselves. NOT actually
    # version-specific (it's the whole increment's update timeline, not
    # a snapshot of this one file) -- bundled onto Increment anyway so
    # every consumer (UI tab, export) reads from the one place this
    # object already serves as, rather than a second live store call.
    change_history: list[dict] = field(default_factory=list)
    # This app's own comments.json entries (see
    # core.project_store.ProjectStore.load_comments) -- free-text notes
    # a user typed directly onto this increment's Changes tab, OLDEST
    # first (append order), same convention as change_history. NOT
    # derived from any file, and NOT read-only the way changes_log is:
    # ui.pages.data_view_page.DataViewPage mutates this list directly
    # (append on add, remove on delete) after each MockDataStore.
    # add_comment()/delete_comment() call, same "update in-memory state
    # immediately, no full reload" pattern set_stage_status() already
    # uses for stage_status.
    comments: list[dict] = field(default_factory=list)
    # All Data's bottom totals row (see core.excel_reader.all_data_totals)
    # -- {"Stage 1": ..., ..., "Stage 42": ..., "VCR": ..., "SUM": ...}.
    # Safe to compute once and cache here, unlike Sum Data's totals:
    # these are plain numeric column sums that never depend on
    # status.json, so they can't go stale the way a precomputed Sum Data
    # total could. Sum Data's totals row is instead recomputed fresh from
    # this Increment's live sum_data every time it's rendered/exported --
    # see core.excel_reader.live_sum_data_totals.
    all_data_totals: dict = field(default_factory=dict)

    @property
    def needs_status_row_count(self) -> int:
        """Number of items with at least one required stage still lacking a status."""
        return sum(1 for row in self.all_data if row.get("needs_status"))

    @property
    def needs_status_stage_count(self) -> int:
        """Number of individual (item, stage) pairs still lacking a status."""
        return sum(len(row.get("needs_status_stages", [])) for row in self.all_data)

    @property
    def changed_count(self) -> int:
        return sum(1 for row in self.all_data if row.get("recently_added"))

    @property
    def value_changed_count(self) -> int:
        """Items present in both this version and the previous one where a
        value (a stage mark, Description, Approval Agency, or OPAA) is
        different -- distinct from changed_count (newly added items). See
        core.value_diff.compare_values.
        """
        return sum(1 for row in self.all_data if row.get("value_changed"))

    @property
    def total_count(self) -> int:
        return len(self.all_data)


@dataclass
class Project:
    name: str
    home_folder: str
    increments: list[Increment] = field(default_factory=list)


@dataclass
class CombinedView:
    """The on-screen "View Combined Data" preview's data -- see
    ui.pages.combined_data_view_page.CombinedDataViewPage. Built by
    build_combined_view() below, straight from core.combined_export's
    public build_combined_*/combine_all_data_totals functions -- the
    SAME functions core.combined_export.export_combined_report() itself
    calls -- so this preview and the exported .xlsx can never disagree
    about which rows appear, in what order, or with what totals. Nothing
    on this object is ever written to; it exists purely for read-only
    display.
    """

    increments: list[Increment]  # in on-screen order -- also what export_combined_report(...) is called with directly
    all_data_rows: list[dict]
    all_data_totals: dict
    sum_data_rows: list[dict]
    sum_data_totals: dict
    report_rows: list[dict]
    revision_log_rows: list[dict]
    update_history_rows: list[dict]
    comments_rows: list[dict]


def build_combined_view(increments: list[Increment]) -> CombinedView:
    """Builds the combined preview's data from an already-loaded list of
    Increment (same list a caller would pass straight to
    core.combined_export.export_combined_report -- see
    ui/pages/home_page.py's _on_view_combined_data/_on_export_combined_report,
    which both build this list the identical way). Does no file I/O
    itself -- everything here is pure Python over data already in
    memory, so building the preview is effectively free once the
    increments themselves are loaded.
    """
    sum_data_rows = combined_export.build_combined_sum_data_rows(increments)
    return CombinedView(
        increments=increments,
        all_data_rows=combined_export.build_combined_all_data_rows(increments),
        all_data_totals=combined_export.combine_all_data_totals(increments),
        sum_data_rows=sum_data_rows,
        sum_data_totals=excel_reader.live_sum_data_totals(sum_data_rows),
        report_rows=combined_export.build_combined_report_rows(increments),
        comments_rows=combined_export.build_combined_comments_rows(increments),
        **combined_export.build_combined_changes_data(increments),
    )


@dataclass
class ComparisonResult:
    increment_name: str
    current_version: int
    added_items: list[dict]  # [{"index": ..., "description": ...}]
    removed_items: list[dict]  # [{"index": ..., "description": ...}]
    column_anomalies: list[str]
    needs_status_items: list[dict]  # [{"index": ..., "description": ..., "stages": [1, 5, 12]}]
    # value_diff.compare_values() results for indexes present in BOTH files
    # (never overlaps added_items/removed_items -- see core/value_diff.py):
    # [{"index": ..., "description": ..., "changes": {field: (old, new)}}]
    value_changed_items: list[dict]
    report_total_before: float  # Grand Total under the CURRENTLY stored version
    report_total_after: float  # Grand Total if the new upload is confirmed
    carried_over_status: dict[str, dict[int, Any]]  # exactly what confirm_update() writes to status.json

    @property
    def has_changes(self) -> bool:
        return bool(
            self.added_items
            or self.removed_items
            or self.column_anomalies
            or self.needs_status_items
            or self.value_changed_items
        )


def _none_if_zero(value: Any) -> Any:
    """All Data represents "blank" grid-sheet cells as numeric 0 (see
    core/excel_reader.py's build_all_data docstring -- it replicates
    Excel's "formula referencing a blank cell evaluates to 0" behavior).
    The data view table only treats None as blank (`is None` checks in
    ui/pages/data_view_page.py, deliberately left untouched by this wiring
    pass), so 0 needs to be normalized to None here at the data source
    rather than changing that rendering code.
    """
    if isinstance(value, bool):
        return value
    return None if value == 0 else value


def _description_by_index(all_data: Any) -> dict[str, Any]:
    return {row["Index"]: row["Description"] for _, row in all_data.iterrows() if row["Index"]}


def _grand_total(report: Any) -> float:
    grand_total_rows = report.loc[report["Apprval Agency"] == "Grand Total", "Total"]
    return float(grand_total_rows.iloc[0]) if not grand_total_rows.empty else 0.0


def _none_if_nan(value: Any) -> Any:
    """Report's Approval Agency/Index/Description columns deliberately
    mix None (continuation-row and Grand-Total placeholders -- see
    core.excel_reader.build_report()) with real strings in the same
    column, and a pandas DataFrame column built that way can coerce
    those Nones to float NaN (the exact same coercion
    core.excel_reader.required_stages_by_index() already has to work
    around for a different column). NaN is truthy in Python, so the
    `value or ""` pattern used when rendering/exporting Report rows
    would otherwise display the literal text "nan" instead of blank.
    """
    return None if pd.isna(value) else value


def _sum_data_rows(all_data_rows: list[dict], sum_data_df: Any) -> list[dict]:
    """Sum Data tab rows: the same row subset and order
    core.excel_reader.build_sum_data() itself selects (see that
    function's docstring -- F-Cons Verif rows unconditionally, grid-sheet
    rows with SUM > 0 and a real Approval Agency), but sourced from the
    already-built All Data rows rather than re-deriving anything, so each
    row's "stage_status" is the literal same dict object as its All Data
    counterpart -- a status.json-backed live view, not the raw-file-
    derived status text build_sum_data() itself computes. Rey's team's
    actual current status (which can include a manual override that
    disagrees with the file's raw marker) is what this tab is for
    surfacing, not a frozen recomputation from the file.
    """
    all_data_by_index = {row["index"]: row for row in all_data_rows}
    rows = []
    for index in sum_data_df["Index"]:
        source = all_data_by_index.get(index)
        if source is None:
            continue
        rows.append(
            {
                "index": index,
                "description": source["description"],
                "approval_agency": source["approval_agency"],
                "required_stages": source["required_stages"],
                "stage_status": source["stage_status"],
                "vcr": source["vcr"],
            }
        )
    return rows


class MockDataStore:
    """See module docstring -- this is the real, file-backed data layer."""

    def __init__(self, store: ProjectStore | None = None):
        self.store = store or ProjectStore()

    # ------------------------------------------------------------------
    # project CRUD
    # ------------------------------------------------------------------
    def project_names(self) -> list[str]:
        return [p.name for p in self.store.list_projects()]

    def get_project(self, name: str) -> Project | None:
        record = self.store.get_project(name)
        if record is None:
            return None
        increments = [
            Increment(name=i.name, version=i.version, last_updated=i.last_updated)
            for i in self.store.list_increments(name)
        ]
        return Project(
            name=record.name,
            home_folder=record.home_folder,
            increments=increments,
        )

    def add_project(self, name: str) -> None:
        self.store.create_project(name)

    def update_project(self, old_name: str, name: str) -> None:
        self.store.update_project(old_name, name)

    def delete_project(self, name: str) -> None:
        self.store.delete_project(name)

    def delete_increment(self, project_name: str, increment_name: str) -> None:
        self.store.delete_increment(project_name, increment_name)

    # ------------------------------------------------------------------
    # increment upload / review flow
    # ------------------------------------------------------------------
    def match_upload(self, project_name: str, file_path: str) -> MatchResult:
        """The single decision behind the unified "Upload File" flow:
        reads file_path's real identity (core.excel_reader.get_record_name)
        and compares it against every increment already stored in
        project_name (core.increment_matcher.match_increment) to decide
        whether this upload is a new increment, a new version of an
        existing one, or a close-enough call that a human should decide.

        Raises whatever get_record_name() raises if the file has no
        readable Record Name -- a required field for this flow, not
        something to silently route around.

        This does real openpyxl parsing (multi-second for a real file) --
        callers on the UI thread should run it via ui.workers.run_with_progress.
        """
        record_name = excel_reader.get_record_name(file_path)
        existing_names = [i.name for i in self.store.list_increments(project_name)]
        return increment_matcher.match_increment(existing_names, record_name)

    def add_new_increment(self, project_name: str, file_path: str) -> Increment:
        """Parses file_path, names the increment from its A-Project Info
        "Record Name (Scope of Project)" field (see
        core.excel_reader.parse_project_info's "record_name"), and stores
        it as version 1. Every required stage is seeded straight from the
        file -- Rey's team's cell fill color if present, otherwise the
        raw X/1 marker value itself, trusted directly -- see
        core.status_tracker.apply_file_derived_seed. Raises ValueError if
        the file has no readable record name or an increment with that
        name already exists in this project (the caller should route
        that case to "Upload New Version" instead).

        This does real openpyxl parsing (multi-second for a real file) --
        callers on the UI thread should run it via ui.workers.run_with_progress.
        """
        parsed = excel_reader.normalize_workbook(file_path)
        increment_name = parsed["project_info"]["record_name"]
        if not increment_name:
            raise ValueError(
                "Could not read an increment name from this file's A-Project Info sheet "
                "(Record Name / Scope of Project is blank)."
            )
        record = self.store.create_increment(project_name, increment_name, file_path)

        required_stages = excel_reader.required_stages_by_index(parsed["all_data"])
        empty_carry = status_tracker.CarryForwardResult(needs_status=dict(required_stages))
        seeded = status_tracker.apply_file_derived_seed(
            empty_carry, parsed["fill_status"], parsed["raw_status"]
        )
        if seeded.carried_over:
            self.store.save_status(project_name, increment_name, seeded.carried_over)

        return Increment(name=record.name, version=record.version, last_updated=record.last_updated)

    def simulate_comparison(self, project_name: str, increment_name: str, file_path: str) -> ComparisonResult:
        """Compares file_path against the increment's current stored
        version: core.structure_diff.compare_structure() for what changed
        structurally, core.status_tracker.carry_forward_status() (at
        per-stage granularity) for what happens to Rey's team's existing
        Done/Open marks.

        This does real openpyxl parsing (multi-second for a real file) --
        callers on the UI thread should run it via ui.workers.run_with_progress.
        """
        increment = self.store.get_increment(project_name, increment_name)
        old_path = self.store.current_version_path(project_name, increment_name)

        # Load each file exactly once and reuse it for both the structural
        # diff and the record parsing below -- these are real, sometimes
        # multi-megabyte macro-enabled workbooks (multi-second loads), and
        # loading the same file twice is a real, user-visible slowdown.
        wb_old = excel_reader.open_workbook(str(old_path))
        wb_new = excel_reader.open_workbook(file_path)

        sheet_diffs = structure_diff.compare_structure(wb_old, wb_new)
        added_indexes: list[str] = []
        removed_indexes: list[str] = []
        column_anomalies: list[str] = []
        for diff in sheet_diffs.values():
            added_indexes.extend(diff.added_indexes)
            removed_indexes.extend(diff.removed_indexes)
            column_anomalies.extend(diff.column_anomalies)

        old_parsed = excel_reader.normalize_workbook(wb_old)
        old_all_data = old_parsed["all_data"]
        new_parsed = excel_reader.normalize_workbook(wb_new)
        new_all_data = new_parsed["all_data"]
        old_description_by_index = _description_by_index(old_all_data)
        new_description_by_index = _description_by_index(new_all_data)
        new_required_stages = excel_reader.required_stages_by_index(new_all_data)

        value_changes = value_diff.compare_values(wb_old, wb_new)
        value_changed_items = [
            {
                "index": idx,
                "description": new_description_by_index.get(idx) or old_description_by_index.get(idx),
                "changes": changes,
            }
            for idx, changes in value_changes.items()
        ]

        report_total_before = _grand_total(old_parsed["report"])
        report_total_after = _grand_total(new_parsed["report"])

        previous_status = self.store.load_status(project_name, increment_name)
        carry = status_tracker.carry_forward_status(previous_status, new_required_stages)
        # A brand-new stage requirement is seeded straight from the
        # uploaded file -- Rey's team's cell fill color if present,
        # otherwise the raw X/1 marker value itself -- instead of
        # surfacing as needing status. See
        # core.status_tracker.apply_file_derived_seed.
        carry = status_tracker.apply_file_derived_seed(
            carry, new_parsed["fill_status"], new_parsed["raw_status"]
        )

        added_items = [{"index": idx, "description": new_description_by_index.get(idx)} for idx in added_indexes]
        removed_items = [{"index": idx, "description": old_description_by_index.get(idx)} for idx in removed_indexes]

        # The review screen's "Items Needing Status" is specifically about
        # what THIS upload introduces (the original UI spec: "new items
        # with no prior Done/Open mark") -- not a re-dump of every
        # already-known-unmarked (item, stage) pair every time a file is
        # re-uploaded. carry.needs_status is broader by design (see
        # core/status_tracker.py -- it also covers pre-existing indexes
        # that newly require a stage, or still lack one), which is exactly
        # right for driving status.json and the Data View's badges/footer
        # counts, but would make an unchanged re-upload permanently unable
        # to show "no changes detected" the moment any (item, stage)
        # anywhere lacks a status. Restricting the review screen's list to
        # needs_status ∩ added keeps that promise while leaving
        # carry.carried_over (below) and the Data View untouched -- they
        # still reflect the true, complete picture regardless of what's
        # "new" in this particular diff.
        needs_status_items = [
            {
                "index": idx,
                "description": new_description_by_index.get(idx),
                "stages": sorted(carry.needs_status[idx]),
            }
            for idx in added_indexes
            if idx in carry.needs_status
        ]

        return ComparisonResult(
            increment_name=increment_name,
            current_version=increment.version,
            added_items=added_items,
            removed_items=removed_items,
            column_anomalies=column_anomalies,
            needs_status_items=needs_status_items,
            value_changed_items=value_changed_items,
            report_total_before=report_total_before,
            report_total_after=report_total_after,
            carried_over_status=carry.carried_over,
        )

    def confirm_update(
        self, project_name: str, increment_name: str, file_path: str, result: ComparisonResult
    ) -> Increment:
        """Saves file_path as the next version, writes the reconciled
        status.json (carried_over status from the comparison -- there's no
        merging needed beyond that: any status set via set_stage_status()
        in between simulate_comparison() and confirm_update() would be for
        the file that's about to be replaced anyway), and appends one
        change_history.json entry recording what this specific update
        changed -- persisted only now, at CONFIRM time, not when
        simulate_comparison() first computed `result` (a review the user
        backs out of leaves no trace). Built directly from `result`'s own
        fields -- the exact same added/removed/anomaly/value-change data
        the review screen already showed the user -- never recomputed.

        This copies a file and writes small JSON files -- no workbook
        parsing, so it's fast; no need to run it via
        ui.workers.run_with_progress.
        """
        record = self.store.save_new_version(project_name, increment_name, file_path)
        self.store.save_status(project_name, increment_name, result.carried_over_status)
        self.store.append_change_history_entry(
            project_name,
            increment_name,
            {
                "timestamp": datetime.now().isoformat(),
                "old_version": result.current_version,
                "new_version": record.version,
                "added_items": result.added_items,
                "removed_items": result.removed_items,
                "column_anomalies": result.column_anomalies,
                "value_changed_items": result.value_changed_items,
            },
        )
        return Increment(name=record.name, version=record.version, last_updated=record.last_updated)

    # ------------------------------------------------------------------
    # data view
    # ------------------------------------------------------------------
    def list_versions(self, project_name: str, increment_name: str) -> list[VersionRecord]:
        return self.store.list_versions(project_name, increment_name)

    def get_increment_for_display(
        self, project_name: str, increment_name: str, version: int | None = None
    ) -> Increment | None:
        """Loads ONE version's file fresh off disk and parses it ONCE via
        core.excel_reader.normalize_workbook() -- which computes All
        Data, Sum Data, and Report together in the same pass -- and
        returns all three on the one Increment. Defaults to the CURRENT
        (latest) version when `version` is omitted -- the single
        background load ui/main_window.py's show_data_view() runs via
        ui.workers.run_with_progress when the Data View first opens.
        ui/pages/data_view_page.py's version-selector dropdown calls this
        again with an explicit `version` on each switch (also via
        run_with_progress -- this is real, multi-second parsing, not a
        cheap cache lookup); switching between the All Data/Sum Data/
        Report TABS of the SAME version never calls back into this
        method, only re-renders from the Increment already in memory.

        Status marks (stage_status) always come from the CURRENT
        status.json regardless of which version is being displayed --
        status isn't tracked historically per version, so viewing an
        older version shows today's Done/Open progress overlaid on that
        version's data, not a historical snapshot of status. Callers
        showing a non-latest version should make that distinction
        visible (see ui/pages/data_view_page.py's version notice).

        "recently_added"/"value_changed" rows are computed by diffing the
        DISPLAYED version against the one immediately before it (nothing
        to diff for version 1). Each All Data row's "required_stages"
        (which of the 42 stage columns actually apply to this item) and
        "stage_status" (status.json's marks, restricted to stages that
        are still required) drive the interactive per-stage cells in
        ui/pages/data_view_page.py; "needs_status_stages" is simply
        "required stages with no entry in stage_status". See
        _sum_data_rows() for how the Sum Data rows relate to these.

        This does real openpyxl parsing (multi-second for a real file) --
        callers on the UI thread should run it via ui.workers.run_with_progress.
        """
        record = self.store.get_increment(project_name, increment_name)
        if record is None:
            return None
        selected_version = version if version is not None else record.version
        selected_path = self.store.version_path(project_name, increment_name, selected_version)
        if selected_path is None:
            return None
        wb_selected = excel_reader.open_workbook(str(selected_path))
        parsed = excel_reader.normalize_workbook(wb_selected)
        all_data_df = parsed["all_data"]
        status_map = self.store.load_status(project_name, increment_name)
        required_stages_by_index = excel_reader.required_stages_by_index(all_data_df)

        recently_added: set[str] = set()
        value_changed: set[str] = set()
        if selected_version > 1:
            previous_path = self.store.version_path(project_name, increment_name, selected_version - 1)
            if previous_path is not None:
                sheet_diffs = structure_diff.compare_structure(str(previous_path), wb_selected)
                for diff in sheet_diffs.values():
                    recently_added.update(diff.added_indexes)
                value_changed.update(value_diff.compare_values(str(previous_path), wb_selected).keys())

        rows = []
        for _, row in all_data_df.iterrows():
            index = row["Index"]
            if not index:
                continue  # blank/category rows aren't real items -- same rule Sum Data/Report use
            required = required_stages_by_index.get(index, [])
            status_for_index = status_map.get(index, {})
            stage_status = {stage: status_for_index[stage] for stage in required if stage in status_for_index}
            needs_status_stages = [stage for stage in required if stage not in stage_status]

            rows.append(
                {
                    "index": index,
                    "description": row["Description"],
                    "approval_agency": row["Apprval Agency"],
                    "required_stages": required,
                    "stage_status": stage_status,
                    "vcr": _none_if_zero(row["VCR"]),
                    "sum": row["SUM"],
                    "recently_added": index in recently_added,
                    "value_changed": index in value_changed,
                    "needs_status_stages": needs_status_stages,
                    "needs_status": bool(needs_status_stages),
                }
            )

        sum_data_rows = _sum_data_rows(rows, parsed["sum_data"])
        report_rows = [
            {
                "approval_agency": _none_if_nan(row["Apprval Agency"]),
                "index": _none_if_nan(row["Index"]),
                "description": _none_if_nan(row["Description"]),
                "total": row["Total"],
            }
            for _, row in parsed["report"].iterrows()
        ]

        changes_log = excel_reader.raw_changes_log(wb_selected["J-Changes"])
        change_history = self.store.load_change_history(project_name, increment_name)
        comments = self.store.load_comments(project_name, increment_name)

        # last_updated for the DISPLAYED version specifically -- record.last_updated
        # is only ever the latest version's date, which is wrong once
        # `version` selects something older.
        version_dates = {v.version: v.uploaded_date for v in self.store.list_versions(project_name, increment_name)}
        last_updated = version_dates.get(selected_version, record.last_updated)

        return Increment(
            name=record.name,
            version=selected_version,
            last_updated=last_updated,
            all_data=rows,
            sum_data=sum_data_rows,
            report=report_rows,
            changes_log=changes_log,
            change_history=change_history,
            comments=comments,
            all_data_totals=parsed["all_data_totals"],
        )

    def set_stage_status(self, project_name: str, increment_name: str, index: str, stage: int, status: str) -> None:
        """Sets a single (index, stage) status and saves immediately -- a
        local JSON write is cheap and instant, so there's no separate
        "Save" step for the user to remember (see the click handler in
        ui/pages/data_view_page.py). Fast (no workbook parsing involved),
        safe to call directly from the UI thread.
        """
        current = self.store.load_status(project_name, increment_name)
        current.setdefault(index, {})[stage] = status
        self.store.save_status(project_name, increment_name, current)

    def add_comment(self, project_name: str, increment_name: str, text: str) -> dict:
        """Appends one comment and saves immediately -- same "cheap local
        JSON write, no separate Save step" pattern as set_stage_status.
        Returns the persisted entry (with its generated id/timestamp) so
        the caller (ui/pages/data_view_page.py) can append it directly
        to Increment.comments in place, without a full re-fetch.
        """
        return self.store.add_comment(project_name, increment_name, text)

    def update_comment(self, project_name: str, increment_name: str, comment_id: str, new_text: str) -> dict | None:
        """Edits one comment's text in place and saves immediately --
        returns the updated entry (id/creation timestamp preserved,
        edited_timestamp set -- see core.project_store.ProjectStore.
        update_comment) so the caller can replace the matching entry in
        Increment.comments in place, without a full re-fetch. None if
        no comment with this id exists.
        """
        return self.store.update_comment(project_name, increment_name, comment_id, new_text)

    def delete_comment(self, project_name: str, increment_name: str, comment_id: str) -> None:
        self.store.delete_comment(project_name, increment_name, comment_id)

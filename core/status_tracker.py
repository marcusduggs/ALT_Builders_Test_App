"""
Status carryover across re-normalization of the same increment, at
(Index, Stage) granularity.

Sum Data's Done/Open status is hand-entered by Rey's team and cannot be
derived from the state's file -- confirmed in an earlier session (Sum
Data has no formulas at all, and in the one real fixture available it had
already drifted from the live blue sheets independently; see
core/excel_reader.py's module docstring, "SUM DATA IS STALE..." section).
It also doesn't live at the Index level: each of the 42 numbered stages
for a given Index can independently be Done, Open, or not-applicable (not
required for that Index at all, per the X/1 markers in All Data). So
whenever the app re-parses an updated file for an increment it already
has status for, that status has to be carried forward per (Index, Stage)
pair, not silently recomputed or dropped on the floor.

This module models that carryover generically -- {index: {stage: status}}
in, {index: [required stage numbers]} in, matching shapes out -- since the
real persistence layer (core/project_store.py) doesn't need to be known
here. It does exactly one thing: reconcile a previous status map against
a freshly computed required-stages map, and report what changed. It does
not guess a status for anything it can't already account for, and it does
not silently drop anything that disappeared -- both are handled the same
way core/structure_diff.py handles structural drift: surfaced for a
human, not resolved here.

Two deliberate exceptions to "cannot be derived from the state's file",
both applied by apply_file_derived_seed() below to brand-new (index,
stage) requirements only (never to anything already carried over): cell
fill color (see reys_fill_color_convention in project memory) and,
per an explicit later decision, the raw X/1 marker value itself, trusted
directly. See that function's docstring for the priority between the two
and core.excel_reader.raw_status_by_index()'s docstring for the reasoning
behind trusting the raw value despite the "SUM DATA IS STALE..." finding.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CarryForwardResult:
    carried_over: dict[str, dict[int, Any]] = field(default_factory=dict)
    needs_status: dict[str, list[int]] = field(default_factory=dict)
    removed: dict[str, list[int]] = field(default_factory=dict)

    @property
    def has_changes(self) -> bool:
        return bool(self.needs_status or self.removed)


def carry_forward_status(
    previous_status: dict[str, dict[int, Any]],
    required_stages: dict[str, list[int]],
) -> CarryForwardResult:
    """Reconciles a previous {index: {stage: status}} map against a
    freshly computed {index: [required stage numbers]} map (e.g. from
    core.excel_reader.required_stages_by_index() on a newly re-parsed
    file's All Data).

    - (index, stage) pairs required now AND present before: carried over
      unchanged in carried_over -- the common case, a stage requirement
      an update didn't touch.
    - Stage requirements that are new -- either a brand-new Index, or an
      Index that newly requires a stage it didn't require before:
      reported in needs_status as {index: [stage numbers]}. Nobody has
      ever set a status for this exact (index, stage) pair, so Rey's team
      needs to.
    - (index, stage) pairs that HAD a status before but aren't required
      anymore -- either the whole Index disappeared, or just that one
      stage requirement did: reported in removed as
      {index: [stage numbers]}, NOT silently dropped. A status just
      vanishing without a trace is exactly the kind of silent data loss
      this function exists to prevent.
    """
    carried_over: dict[str, dict[int, Any]] = {}
    needs_status: dict[str, list[int]] = {}

    for index, stages in required_stages.items():
        previous_for_index = previous_status.get(index, {})
        carried: dict[int, Any] = {}
        new_needs: list[int] = []
        for stage in stages:
            if stage in previous_for_index:
                carried[stage] = previous_for_index[stage]
            else:
                new_needs.append(stage)
        if carried:
            carried_over[index] = carried
        if new_needs:
            needs_status[index] = new_needs

    removed: dict[str, list[int]] = {}
    for index, previous_for_index in previous_status.items():
        current_required = set(required_stages.get(index, []))
        gone = [stage for stage in previous_for_index if stage not in current_required]
        if gone:
            removed[index] = gone

    return CarryForwardResult(
        carried_over=carried_over,
        needs_status=needs_status,
        removed=removed,
    )


def apply_file_derived_seed(
    carry: CarryForwardResult,
    *status_sources: dict[str, dict[int, Any]],
) -> CarryForwardResult:
    """Seeds brand-new (index, stage) requirements straight from the
    uploaded file, instead of leaving them fully unset -- from one or
    more status sources, in priority order (the first source with a
    value for a given (index, stage) wins). In practice, called with
    core.excel_reader.fill_status_by_index() first, then
    raw_status_by_index() as the fallback:

    - fill_status_by_index: Rey's team's own cell fill color (green/red)
      -- see reys_fill_color_convention in project memory. An explicit,
      deliberate mark, so it wins if a stage has one.
    - raw_status_by_index: the raw X/1 marker value itself, trusted
      directly (a deliberate call -- see that function's docstring for
      why, given this module's own "SUM DATA IS STALE..." caution above).
      Since a required grid-sheet cell is only ever "X" or "1" to begin
      with, this source alone already covers every required stage, so in
      practice needs_status ends up empty for grid-sheet items after
      this runs -- fill_status_by_index only matters for the (probably
      rare) case where Rey's team colored a cell to disagree with its
      raw marker.

    Only ever seeds carry.needs_status entries -- (index, stage) pairs
    that have NEVER had a status before. Anything already in
    carry.carried_over (a status Rey's team set through the app on a
    prior version) is left completely untouched; these sources never
    override an existing mark, they only fill the gap for a first-time
    one. A stage with no value in ANY source stays in needs_status,
    exactly as if this function didn't exist.
    """
    carried_over = {index: dict(stages) for index, stages in carry.carried_over.items()}
    needs_status: dict[str, list[int]] = {}

    for index, stages in carry.needs_status.items():
        still_needed = []
        seeded: dict[int, Any] = {}
        for stage in stages:
            resolved = None
            for source in status_sources:
                resolved = source.get(index, {}).get(stage)
                if resolved is not None:
                    break
            if resolved is not None:
                seeded[stage] = resolved
            else:
                still_needed.append(stage)
        if seeded:
            carried_over.setdefault(index, {}).update(seeded)
        if still_needed:
            needs_status[index] = still_needed

    return CarryForwardResult(
        carried_over=carried_over,
        needs_status=needs_status,
        removed=carry.removed,
    )

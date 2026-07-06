"""
Decides whether an uploaded file's Record Name (see
core.excel_reader.get_record_name) refers to a NEW increment or an
EXISTING one already stored in a project -- the single piece of logic
behind the unified "Upload File" flow (previously two separate user
choices: "Upload New Increment" vs "Upload New Version", with the user
manually picking which existing increment to update).

Takes plain increment name strings, not a ProjectStore/project id: this
codebase identifies projects and increments by their display name
everywhere else (see core/project_store.py's module docstring -- slugs
are an internal filesystem detail callers never see), so match_increment()
follows that same convention rather than introducing a separate id
concept. Callers pass `[i.name for i in store.list_increments(project_name)]`.

Three outcomes, each independently explainable without a black-box fuzzy
score:
  EXACT_MATCH  -- record_name equals an existing increment's name,
                  case/whitespace differences aside (get_record_name()
                  already trims and collapses whitespace at the source,
                  so in practice this is very close to a literal match --
                  e.g. a trailing space introduced some other way is
                  still EXACT_MATCH, deliberately: that's not a case
                  worth interrupting the upload to ask about).
  CLOSE_MATCH  -- not an exact match, but close enough that a human
                  should confirm rather than the app guessing: the
                  SAME text once punctuation is stripped too (catches a
                  different dash character, or a missing/extra space
                  around punctuation), OR a small Levenshtein distance
                  between the punctuation-stripped forms.
  NO_MATCH     -- nothing close enough; safe to treat as a new increment.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum, auto

_WHITESPACE = re.compile(r"\s+")
_PUNCTUATION = re.compile(r"[^\w\s]")

# Deliberately small and flat, not scaled to name length: this is meant
# to catch typos/formatting slips a human glancing at both names would
# immediately recognize as "the same thing", not to be a general fuzzy-
# match tolerance. Two names that differ by more than this are treated
# as NO_MATCH rather than guessed at.
CLOSE_MATCH_EDIT_DISTANCE_THRESHOLD = 3


class MatchType(Enum):
    EXACT_MATCH = auto()
    CLOSE_MATCH = auto()
    NO_MATCH = auto()


@dataclass
class MatchResult:
    match_type: MatchType
    matched_increment_name: str | None = None


def _normalize_exact(name: str) -> str:
    return _WHITESPACE.sub(" ", name).strip().lower()


def _normalize_loose(name: str) -> str:
    """Exact-normalization, plus punctuation stripped entirely -- so a
    hyphen, en dash, or em dash, or a missing/extra space around one, all
    collapse to the same string for closeness comparison.
    """
    stripped = _PUNCTUATION.sub(" ", _normalize_exact(name))
    return _WHITESPACE.sub(" ", stripped).strip()


def _levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    previous_row = list(range(len(b) + 1))
    for i, char_a in enumerate(a, start=1):
        current_row = [i] + [0] * len(b)
        for j, char_b in enumerate(b, start=1):
            current_row[j] = min(
                current_row[j - 1] + 1,  # insertion
                previous_row[j] + 1,  # deletion
                previous_row[j - 1] + (char_a != char_b),  # substitution
            )
        previous_row = current_row
    return previous_row[-1]


def match_increment(existing_increment_names: list[str], record_name: str) -> MatchResult:
    """Compares record_name (see core.excel_reader.get_record_name)
    against every name in existing_increment_names, in order, and
    returns the single best MatchResult.
    """
    if not existing_increment_names:
        return MatchResult(MatchType.NO_MATCH)

    target_exact = _normalize_exact(record_name)
    for name in existing_increment_names:
        if _normalize_exact(name) == target_exact:
            return MatchResult(MatchType.EXACT_MATCH, matched_increment_name=name)

    target_loose = _normalize_loose(record_name)
    best_name = None
    best_distance = None
    for name in existing_increment_names:
        existing_loose = _normalize_loose(name)
        if existing_loose == target_loose:
            return MatchResult(MatchType.CLOSE_MATCH, matched_increment_name=name)
        distance = _levenshtein(target_loose, existing_loose)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_name = name

    if best_distance is not None and best_distance <= CLOSE_MATCH_EDIT_DISTANCE_THRESHOLD:
        return MatchResult(MatchType.CLOSE_MATCH, matched_increment_name=best_name)

    return MatchResult(MatchType.NO_MATCH)

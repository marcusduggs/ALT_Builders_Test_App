"""
Tests core/update_check.py: version comparison (including the v0.9.0 vs
v0.10.0 ordering that a naive string comparison gets wrong) and
check_for_update()'s three outcomes, against the GET /releases LIST
endpoint (not /releases/latest, which excludes pre-releases) -- with the
GitHub API call mocked, never a real network call.

Run directly: `python tests/test_update_check.py`
"""

import json
import os
import sys
import urllib.error
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import update_check
from core.update_check import RELEASES_LIST_API_URL, UpdateStatus, _compare_versions, _parse_version, check_for_update


def _mock_response(payload):
    """A context-manager mock matching urllib.request.urlopen()'s return
    value -- .read() gives the raw JSON bytes a real HTTP response would.
    `payload` is whatever json.dumps() accepts (a list, for the real
    /releases endpoint shape -- or deliberately a non-list, to test that
    malformed-response case).
    """
    mock = MagicMock()
    mock.__enter__.return_value.read.return_value = json.dumps(payload).encode("utf-8")
    return mock


def _release(tag_name: str, prerelease: bool = False) -> dict:
    return {
        "tag_name": tag_name,
        "html_url": f"https://github.com/marcusduggs/ALT_Builders_Test_App/releases/tag/{tag_name}",
        "prerelease": prerelease,
    }


def main():
    failures = []

    # ------------------------------------------------------------
    # 1 -- version parsing/comparison, especially v0.9.0 vs v0.10.0
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- version parsing and comparison")
    print("=" * 70)

    v9 = _parse_version("v0.9.0")
    v10 = _parse_version("v0.10.0")
    print(f"_parse_version('v0.9.0') = {v9}")
    print(f"_parse_version('v0.10.0') = {v10}")
    if v9 != (0, 9, 0):
        failures.append(f"1: _parse_version('v0.9.0') should be (0, 9, 0), got {v9}")
    if v10 != (0, 10, 0):
        failures.append(f"1: _parse_version('v0.10.0') should be (0, 10, 0), got {v10}")

    naive_string_result = "0.10.0" < "0.9.0"
    print(f"Naive string comparison '0.10.0' < '0.9.0': {naive_string_result} (WRONG, this is why we don't do that)")
    if not naive_string_result:
        failures.append("1: sanity check failed -- expected the naive string comparison to be wrong here")

    cmp_result = _compare_versions(v10, v9)
    print(f"_compare_versions(v0.10.0, v0.9.0) = {cmp_result} (should be 1 -- 0.10.0 IS newer)")
    if cmp_result != 1:
        failures.append(f"1: _compare_versions should correctly rank 0.10.0 > 0.9.0, got {cmp_result}")

    if _compare_versions((0, 7, 0), (0, 7, 0)) != 0:
        failures.append("1: equal versions should compare as 0")
    if _compare_versions((0, 7), (0, 7, 0)) != 0:
        failures.append("1: (0, 7) and (0, 7, 0) should be treated as equal (zero-padding)")
    if _compare_versions((1, 0, 0), (0, 9, 9)) != 1:
        failures.append("1: 1.0.0 should be newer than 0.9.9")
    if _compare_versions((0, 6, 0), (0, 6, 1)) != -1:
        failures.append("1: 0.6.0 should be OLDER than 0.6.1")

    stray_dot = _parse_version("v.0.6.0")
    print(f"_parse_version('v.0.6.0') (this repo's actual stray-dot tag format) = {stray_dot}")
    if stray_dot != (0, 6, 0):
        failures.append(f"1: _parse_version('v.0.6.0') should still be (0, 6, 0), got {stray_dot}")

    # ------------------------------------------------------------
    # 2 -- hits the /releases LIST endpoint, not /latest
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- confirm the LIST endpoint is used, not /releases/latest")
    print("=" * 70)
    print(f"RELEASES_LIST_API_URL = {RELEASES_LIST_API_URL}")
    if RELEASES_LIST_API_URL.endswith("/latest"):
        failures.append("2: should be using the full /releases list endpoint, not /releases/latest")
    if not RELEASES_LIST_API_URL.endswith("/releases"):
        failures.append(f"2: expected the URL to end with /releases, got {RELEASES_LIST_API_URL!r}")

    captured_request = {}

    def capturing_urlopen(request, timeout=None):
        captured_request["url"] = request.full_url
        return _mock_response([_release("v0.0.1")])

    with patch("urllib.request.urlopen", side_effect=capturing_urlopen):
        check_for_update()
    print(f"Actual URL requested: {captured_request.get('url')}")
    if captured_request.get("url") != RELEASES_LIST_API_URL:
        failures.append(f"2: expected request to {RELEASES_LIST_API_URL}, got {captured_request.get('url')}")

    # ------------------------------------------------------------
    # 3 -- highest version wins, NOT list order -- and a stable (non-
    # prerelease) release is the true highest here
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- highest version by PARSED value wins, regardless of list position")
    print("=" * 70)

    out_of_order_list = [
        _release("v0.5.0"),            # listed first (e.g. most recently created) but NOT highest
        _release("v99.0.0"),           # actually highest -- listed second
        _release("v10.0.0"),           # higher than v0.5.0 but still lower than v99.0.0
    ]
    with patch("urllib.request.urlopen", return_value=_mock_response(out_of_order_list)):
        result = check_for_update()
    print(f"Result: {result}")
    if result.status != UpdateStatus.UPDATE_AVAILABLE:
        failures.append(f"3: expected UPDATE_AVAILABLE, got {result.status}")
    if result.latest_version != "99.0.0":
        failures.append(f"3: expected the highest version (99.0.0) to win regardless of list order, got {result.latest_version!r}")
    if result.is_prerelease:
        failures.append("3: v99.0.0 in this list is NOT a prerelease -- is_prerelease should be False")

    # ------------------------------------------------------------
    # 4 -- the highest version IS a prerelease -- dialog wording (Step 3)
    # depends on is_prerelease being surfaced correctly here
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("4 -- highest version is a PRE-RELEASE -- flagged as such")
    print("=" * 70)

    prerelease_is_newest_list = [
        _release("v50.0.0", prerelease=False),
        _release("v99.0.0-beta", prerelease=True),  # highest overall, and IS a prerelease
    ]
    with patch("urllib.request.urlopen", return_value=_mock_response(prerelease_is_newest_list)):
        result_pre = check_for_update()
    print(f"Result: {result_pre}")
    if result_pre.status != UpdateStatus.UPDATE_AVAILABLE:
        failures.append(f"4: expected UPDATE_AVAILABLE, got {result_pre.status}")
    if not result_pre.is_prerelease:
        failures.append("4: the highest version here IS a prerelease -- is_prerelease should be True")
    if result_pre.latest_version != "99.0.0-beta":
        failures.append(f"4: expected latest_version '99.0.0-beta', got {result_pre.latest_version!r}")

    # a prerelease existing in the list, but NOT being the highest, must
    # NOT cause a false "prerelease" label on the actual (stable) winner
    stable_is_still_newest_list = [
        _release("v99.0.0", prerelease=False),   # the true highest, stable
        _release("v50.0.0-beta", prerelease=True),  # a prerelease, but lower than v99.0.0
    ]
    with patch("urllib.request.urlopen", return_value=_mock_response(stable_is_still_newest_list)):
        result_stable = check_for_update()
    print(f"Result: {result_stable}")
    if result_stable.is_prerelease:
        failures.append("4: the highest version here is NOT a prerelease -- is_prerelease should be False, "
                         "even though a lower-versioned prerelease also exists in the list")

    # ------------------------------------------------------------
    # 5 -- UP_TO_DATE (every release in the list is <= current version)
    #
    # core/app_version.py's checked-in APP_VERSION is now a deliberate
    # "0.0.0-dev" placeholder (see .github/workflows/release.yml -- a
    # released build's real version only ever comes from the tag,
    # stamped at build time), so it parses as (0, 0, 0) and is NOT a
    # realistic "current version" to compare against here. Patching
    # core.update_check.APP_VERSION directly (the name as imported INTO
    # that module, not core.app_version.APP_VERSION itself -- patch
    # where a name is USED, not where it's defined) with a fixed,
    # realistic value keeps this test meaningful regardless of what the
    # dev placeholder happens to be.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("5 -- check_for_update(): UP_TO_DATE")
    print("=" * 70)

    with patch.object(update_check, "APP_VERSION", "1.0.0"):
        same_version_list = [_release("v1.0.0"), _release("v0.0.1")]
        with patch("urllib.request.urlopen", return_value=_mock_response(same_version_list)):
            result2 = check_for_update()
        print(f"Result: {result2}")
        if result2.status != UpdateStatus.UP_TO_DATE:
            failures.append(f"5: expected UP_TO_DATE when the highest tag matches the current version exactly, got {result2.status}")

        older_only_list = [_release("v0.0.1"), _release("v0.1.0")]
        with patch("urllib.request.urlopen", return_value=_mock_response(older_only_list)):
            result_older = check_for_update()
        print(f"Result (only older tags than current): {result_older}")
        if result_older.status != UpdateStatus.UP_TO_DATE:
            failures.append(f"5: releases all OLDER than the current version should also be UP_TO_DATE, got {result_older.status}")

    # ------------------------------------------------------------
    # 6 -- check_for_update(): CHECK_FAILED -- network error, timeout,
    # malformed/unexpected response shapes -- must never raise
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("6 -- check_for_update(): CHECK_FAILED against the new endpoint (never raises)")
    print("=" * 70)

    with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("no route to host")):
        result3 = check_for_update()
    print(f"Result (URLError): {result3}")
    if result3.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED on URLError, got {result3.status}")

    with patch("urllib.request.urlopen", side_effect=TimeoutError("timed out")):
        result4 = check_for_update()
    print(f"Result (TimeoutError): {result4}")
    if result4.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED on TimeoutError, got {result4.status}")

    malformed_mock = MagicMock()
    malformed_mock.__enter__.return_value.read.return_value = b"not valid json{{{"
    with patch("urllib.request.urlopen", return_value=malformed_mock):
        result5 = check_for_update()
    print(f"Result (malformed JSON): {result5}")
    if result5.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED on malformed JSON, got {result5.status}")

    with patch("urllib.request.urlopen", return_value=_mock_response([])):
        result_empty = check_for_update()
    print(f"Result (empty list): {result_empty}")
    if result_empty.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED on an empty releases list, got {result_empty.status}")

    with patch("urllib.request.urlopen", return_value=_mock_response({"message": "not a list at all"})):
        result_not_list = check_for_update()
    print(f"Result (response is a dict, not a list): {result_not_list}")
    if result_not_list.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED when the response isn't a list, got {result_not_list.status}")

    missing_fields_list = [{"some_other_field": "nothing useful here"}]
    with patch("urllib.request.urlopen", return_value=_mock_response(missing_fields_list)):
        result6 = check_for_update()
    print(f"Result (every entry missing tag_name/html_url): {result6}")
    if result6.status != UpdateStatus.CHECK_FAILED:
        failures.append(f"6: expected CHECK_FAILED when no entry has usable tag_name/html_url, got {result6.status}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- semantic version comparison (incl. v0.9.0 vs v0.10.0), the /releases LIST "
              "endpoint (highest-by-version wins regardless of list order), pre-release detection "
              "(from the API's own field, not the tag), and CHECK_FAILED are all correct")


if __name__ == "__main__":
    main()

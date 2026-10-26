#!/usr/bin/env python3
"""Replay one explicitly preserved ViewerSpace hover/navigation run.

This standalone standard-library analyzer reads only the pinned evidence paths
below, verifies the preservation manifest and run/build provenance, and creates
one exclusive-create JSON report. It does not run tests, alter artifacts, or
fall back to historical runs. Invoke only after the root task has preserved a
terminal filtered run and reviewed this script.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import statistics
from pathlib import Path, PurePosixPath
from typing import Any, Iterable


TEST_ROOT = Path(r"D:\RedSalamander.Perf")
EVIDENCE = TEST_ROOT / "evidence" / "viewer-space-finish-all-20261009" / "hover-navigation-current"
ARTIFACTS = EVIDENCE / "artifacts"
LAST_RUN = ARTIFACTS / "selftest" / "last_run"
PERF = LAST_RUN / "perf"
REPORT = TEST_ROOT / "evidence" / "viewer-space-finish-all-20261009" / "root-hover-navigation-current-replay.json"
EXPECTED_RECEIPT = "37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd"
EXPECTED_SOURCE = "9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554"
EXPECTED_CASE = "cmd_viewer_space_hover_does_not_rebuild_static_cache"
MARKER_TEXT = "red-salamander.test-sandbox-root.v1"


def fail(message: str) -> None:
    raise RuntimeError(message)


def read_json(path: Path) -> Any:
    try:
        with path.open("r", encoding="utf-8-sig") as stream:
            value = json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        fail(f"Cannot read valid JSON at {path}: {exc}")
    if not isinstance(value, dict):
        fail(f"Expected a JSON object at {path}")
    return value


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def checked_relative_path(raw: Any) -> PurePosixPath:
    if not isinstance(raw, str) or not raw:
        fail("Preservation manifest contains an empty/non-string path")
    path = PurePosixPath(raw.replace("\\", "/"))
    if path.is_absolute() or any(part in ("", ".", "..") for part in path.parts):
        fail(f"Unsafe preservation path: {raw!r}")
    if path.parts and ":" in path.parts[0]:
        fail(f"Drive-qualified preservation path is forbidden: {raw!r}")
    return path


def resolve_manifest_path(relative: PurePosixPath) -> Path:
    # Root may record paths relative to the copied artifact root or include
    # the leading "artifacts/" component. Accept only these two exact forms.
    parts = relative.parts
    if parts[0] == "artifacts":
        target = EVIDENCE.joinpath(*parts)
    else:
        target = ARTIFACTS.joinpath(*parts)
    resolved = target.resolve(strict=True)
    allowed = ARTIFACTS.resolve(strict=True)
    if not resolved.is_relative_to(allowed):
        fail(f"Manifest path escaped preserved artifacts: {relative}")
    if not resolved.is_file():
        fail(f"Manifest entry is not a regular file: {relative}")
    return resolved


def verify_preservation(manifest: dict[str, Any], request: dict[str, Any]) -> dict[str, dict[str, Any]]:
    expected_identity = {
        "sourceRunId": request["runId"],
        "sourceSnapshotId": EXPECTED_SOURCE,
        "receiptId": EXPECTED_RECEIPT,
    }
    for key, expected in expected_identity.items():
        if manifest.get(key) != expected:
            fail(f"preservation.json {key} does not match the requested run/build")
    entries = manifest.get("files")
    if not isinstance(entries, list) or not entries:
        fail("preservation.json must contain a nonempty files array")

    verified: dict[str, dict[str, Any]] = {}
    for entry in entries:
        if not isinstance(entry, dict):
            fail("preservation.json files entries must be objects")
        relative = checked_relative_path(entry.get("path"))
        target = resolve_manifest_path(relative)
        expected_bytes = entry.get("bytes")
        expected_hash = entry.get("sha256")
        if not isinstance(expected_bytes, int) or expected_bytes < 0:
            fail(f"Invalid byte count in preservation manifest: {relative}")
        if not isinstance(expected_hash, str) or len(expected_hash) != 64:
            fail(f"Invalid SHA-256 in preservation manifest: {relative}")
        actual_bytes = target.stat().st_size
        actual_hash = sha256(target)
        if actual_bytes != expected_bytes or actual_hash.lower() != expected_hash.lower():
            fail(f"Preserved file does not match manifest: {relative}")
        canonical = target.relative_to(ARTIFACTS.resolve(strict=True)).as_posix()
        if canonical in verified:
            fail(f"Duplicate preservation path: {relative}")
        verified[canonical] = {"bytes": actual_bytes, "sha256": actual_hash}

    required = (
        "selftest/last_run/run-all-tests-results.json",
        "selftest/last_run/workspace-final.json",
        "selftest/last_run/run-all-tests-case-history.jsonl",
        "selftest/last_run/perf/perf_metrics.jsonl",
        "selftest/last_run/perf/viewer_space_input_present_metrics.json",
        "selftest/last_run/perf/viewer_space_navigation_present_metrics.json",
    )
    missing = [name for name in required if name not in verified]
    if missing:
        fail(f"Preservation manifest omits required run evidence: {', '.join(missing)}")
    return verified


def load_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig") as stream:
        for number, line in enumerate(stream, 1):
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError as exc:
                fail(f"Invalid JSONL at {path}:{number}: {exc}")
            if not isinstance(row, dict):
                fail(f"Non-object JSONL row at {path}:{number}")
            rows.append(row)
    return rows


def durations(values: Iterable[Any], label: str) -> list[int]:
    result: list[int] = []
    try:
        iterator = iter(values)
    except TypeError:
        fail(f"{label} is not a sample sequence")
    for value in iterator:
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            fail(f"{label} contains a nonnegative-integer duration violation")
        result.append(value)
    if not result:
        fail(f"{label} contains no samples")
    return result


def distribution_summary(values: Iterable[Any], label: str) -> dict[str, int | float]:
    ordered = sorted(durations(values, label))
    count = len(ordered)
    return {
        "count": count,
        "minUs": ordered[0],
        "medianUs": statistics.median(ordered),
        "p95NearestRankUs": ordered[math.ceil(0.95 * count) - 1],
        "maxUs": ordered[-1],
    }


def metric_summary(rows: list[dict[str, Any]], name: str) -> dict[str, Any]:
    selected = [row for row in rows if row.get("metric") == name]
    if not selected:
        return {"count": 0, "status": "NO_ROWS_IN_FILTERED_RUN"}
    values = durations([row.get("durationUs") for row in selected], name)
    failed = [row for row in selected if row.get("hr") != 0]
    summary = distribution_summary(values, name)
    summary["hresultNonSuccessCount"] = len(failed)
    summary["nonSuccessHrValues"] = sorted({str(row.get("hr")) for row in failed})
    summary["detailCounts"] = {
        str(detail): sum(1 for row in selected if row.get("detail") == detail)
        for detail in sorted({row.get("detail") for row in selected}, key=lambda item: str(item))
    }
    return summary


def verify_run() -> dict[str, Any]:
    marker = TEST_ROOT / ".red-salamander-test-root"
    if marker.read_text(encoding="utf-8-sig").strip() != MARKER_TEXT:
        fail("Test-root ownership marker does not match")
    if REPORT.exists():
        fail(f"Refusing to overwrite existing replay report: {REPORT}")

    request = read_json(EVIDENCE / "request.json")
    if not isinstance(request, dict):
        fail("request.json must be an object")
    for key, expected in (("receiptId", EXPECTED_RECEIPT), ("sourceSnapshotId", EXPECTED_SOURCE)):
        if request.get(key) != expected:
            fail(f"request.json {key} does not match the pinned Release build")
    run_id = request.get("runId")
    if not isinstance(run_id, str) or not run_id or request.get("finishedUtc") in (None, ""):
        fail("request.json must identify a finished run")
    if request.get("exitCode") != 0:
        fail("The filtered run did not exit successfully")
    if request.get("case") not in (None, EXPECTED_CASE):
        fail("request.json case filter does not match the expected ViewerSpace case")

    receipt = read_json(EVIDENCE / "build-receipt.json")
    if not isinstance(receipt, dict) or receipt.get("schema") != "red-salamander.build-receipt.v1":
        fail("Preserved build receipt has an unexpected schema")
    if (receipt.get("receipt_id") != EXPECTED_RECEIPT or
            receipt.get("source_snapshot_id") != EXPECTED_SOURCE or
            receipt.get("configuration") != "Release" or
            receipt.get("platform") != "x64" or
            receipt.get("tests_enabled") is not True):
        fail("Preserved build receipt is not the pinned test-enabled x64 Release build")

    manifest = read_json(EVIDENCE / "preservation.json")
    if not isinstance(manifest, dict):
        fail("preservation.json must be an object")
    verified_files = verify_preservation(manifest, request)

    v1_path = LAST_RUN / "run-all-tests-results.json"
    v1 = read_json(v1_path)
    if (v1.get("run_id") != run_id or v1.get("suite") != "Commands" or
            v1.get("platform") != "x64" or v1.get("configuration") != "Release" or
            v1.get("case_filter") != EXPECTED_CASE or v1.get("exit_code") != 0):
        fail("V1 result does not prove the exact successful filtered Commands run")

    if (v1.get("schema") != "red-salamander.run-all-tests.v1" or
            (v1.get("passed"), v1.get("failed"), v1.get("skipped"), v1.get("total")) != (1, 0, 0, 1)):
        fail("Filtered V1 report does not retain the exact one-case outcome")
    workspace = read_json(LAST_RUN / "workspace-final.json")
    if (workspace.get("snapshot_id") != EXPECTED_SOURCE or
            workspace.get("acquisition_before") != EXPECTED_SOURCE or workspace.get("acquisition_after") != EXPECTED_SOURCE):
        fail("Exact frozen workspace-final snapshot was not preserved")

    case_history_path = LAST_RUN / "run-all-tests-case-history.jsonl"
    case_rows = [row for row in load_jsonl(case_history_path)
                 if row.get("run_id") == run_id and row.get("suite") == "Commands" and row.get("case") == EXPECTED_CASE]
    if len(case_rows) != 1 or case_rows[0].get("status") != "passed":
        fail("Case history does not contain exactly one passing target case")

    input_path = PERF / "viewer_space_input_present_metrics.json"
    navigation_path = PERF / "viewer_space_navigation_present_metrics.json"
    input_artifact = read_json(input_path)
    navigation_artifact = read_json(navigation_path)
    if input_artifact.get("scenario") != "hover-static-cache" or input_artifact.get("boundary") != "handler-to-successful-next-present-submission":
        fail("Hover artifact has unexpected scenario/boundary")
    if navigation_artifact.get("scenario") != "retained-control-wheel-navigation" or navigation_artifact.get("boundary") != "handler-call-to-new-target-layout-successful-present-submission":
        fail("Navigation artifact has unexpected scenario/boundary")
    hover = durations(input_artifact.get("samplesUs", []), "hover samples")
    zoom_in = durations(navigation_artifact.get("zoomInUs", []), "zoom-in samples")
    zoom_out = durations(navigation_artifact.get("zoomOutUs", []), "zoom-out samples")
    if len(hover) != 200 or len(zoom_in) != 200 or len(zoom_out) != 200:
        fail("Expected exactly 200 hover, 200 zoom-in, and 200 zoom-out samples")

    raw_path = PERF / "perf_metrics.jsonl"
    raw_rows = load_jsonl(raw_path)
    nav_rows = [row for row in raw_rows if row.get("metric") == "viewer.space.input.fixture_navigation_present_us"]
    in_rows = [row for row in nav_rows if row.get("detail") == "wheel-in"]
    out_rows = [row for row in nav_rows if row.get("detail") == "wheel-out"]
    hover_rows = [row for row in raw_rows if row.get("metric") == "viewer.space.input.fixture_hover_p95_us"]
    if len(nav_rows) != 400 or len(in_rows) != 200 or len(out_rows) != 200 or len(hover_rows) != 1:
        fail("Raw metric rows do not match the expected hover/navigation sample cardinality")
    if any(row.get("hr") != 0 for row in nav_rows):
        fail("A navigation sample did not report S_OK")
    if [row.get("durationUs") for row in in_rows] != zoom_in or [row.get("durationUs") for row in out_rows] != zoom_out:
        fail("Navigation JSON arrays do not exactly match raw metric rows in order")
    hover_row = hover_rows[0]
    hover_summary = distribution_summary(hover, "hover samples")
    if (hover_row.get("hr") != 0 or hover_row.get("detail") != "200-observations" or hover_row.get("value0") != 200 or
            hover_row.get("durationUs") != hover_summary["p95NearestRankUs"] or
            hover_row.get("value1") != hover_summary["maxUs"]):
        fail("Hover aggregate metric does not match its raw 200-sample JSON artifact")

    for row in raw_rows:
        if str(row.get("metric", "")).startswith("viewer.space."):
            if row.get("scenario") != "SelfTest" or row.get("build") != "Release" or row.get("machineHash") != "7d3a1247382a":
                fail("ViewerSpace metric row has unexpected scenario/build provenance")

    timing_names = (
        "viewer.space.layout.turn_us",
        "viewer.space.analysis.turn_us",
        "viewer.space.layout.publish_us",
        "viewer.space.open_to_first_paint_us",
        "viewer.space.open_to_first_meaningful_paint_us",
        "viewer.space.open.setup_us",
        "viewer.space.render.prepare_us",
    )
    timing = {name: metric_summary(raw_rows, name) for name in timing_names}
    for metric_name in ("viewer.space.layout.turn_us", "viewer.space.analysis.turn_us"):
        summary = timing[metric_name]
        if summary.get("count", 0):
            summary["target2msP95_4msMax"] = (
                "OBSERVED_WITHIN_TARGET_THIS_RUN"
                if summary["p95NearestRankUs"] <= 2000 and summary["maxUs"] <= 4000
                else "OBSERVED_ABOVE_TARGET_THIS_RUN"
            )
            summary["qualification"] = "filtered-run diagnostic only; not a corpus-wide qualification"

    def target_state(stats: dict[str, int], limit_us: int) -> str:
        return "OBSERVED_WITHIN_TARGET_THIS_RUN" if stats["p95NearestRankUs"] <= limit_us else "OBSERVED_ABOVE_TARGET_THIS_RUN"

    hover_stats = distribution_summary(hover, "hover samples")
    zoom_in_stats = distribution_summary(zoom_in, "zoom-in samples")
    zoom_out_stats = distribution_summary(zoom_out, "zoom-out samples")
    hover_stats["target16msSettled"] = target_state(hover_stats, 16000)
    zoom_in_stats["target50msSettled"] = target_state(zoom_in_stats, 50000)
    zoom_out_stats["target50msSettled"] = target_state(zoom_out_stats, 50000)

    metric_run_ids = sorted({str(row.get("runId")) for row in raw_rows if str(row.get("metric", "")).startswith("viewer.space.")})
    return {
        "schema": "red-salamander.viewer-space-hover-navigation-replay.v1",
        "reviewer": "root",
        "mismatches": 0,
        "provenance": {
            "runId": run_id,
            "finishedUtc": request["finishedUtc"],
            "exitCode": request["exitCode"],
            "receiptId": EXPECTED_RECEIPT,
            "sourceSnapshotId": EXPECTED_SOURCE,
            "workspaceSnapshotId": workspace["snapshot_id"],
            "runReportSchema": v1["schema"],
            "repositoryVerdict": "NOT_EVALUATED",
            "repositoryVerdictBasis": "Derived from the canonical focused Commands contract; original run report is v1 and has no v2 repository_verdict field.",
            "testSandboxAudit": v1["test_sandbox_audit"],
            "case": EXPECTED_CASE,
            "caseStatus": case_rows[0]["status"],
            "metricRunIds": metric_run_ids,
            "evidenceFiles": {
                name: verified_files[name]
                for name in (
                    "selftest/last_run/run-all-tests-results.json",
                    "selftest/last_run/workspace-final.json",
                    "selftest/last_run/run-all-tests-case-history.jsonl",
                    "selftest/last_run/perf/perf_metrics.jsonl",
                    "selftest/last_run/perf/viewer_space_input_present_metrics.json",
                    "selftest/last_run/perf/viewer_space_navigation_present_metrics.json",
                )
            },
        },
        "samples": {
            "hoverHandlerToNextPresentSubmission": hover_stats,
            "retainedControlWheelZoomInToTargetPresentSubmission": zoom_in_stats,
            "retainedControlWheelZoomOutToTargetPresentSubmission": zoom_out_stats,
            "rawNavigationRows": {"wheelIn": len(in_rows), "wheelOut": len(out_rows), "allSOk": True},
        },
        "uiPreparationDiagnostics": timing,
        "chromeFirstPaint100msTarget": {
            "status": "UNQUALIFIED_ENDPOINT_GAP",
            "reason": "viewer.space.open_to_first_paint_us is scan-start based; open.setup_us ends before StartScan; first-meaningful-paint is a different map endpoint.",
            "noCurrentChromeOpenToFirstPaintMeasurementClaimed": True,
        },
        "limits": [
            "This is one filtered Release Commands run, not repository-wide qualification.",
            "The repository verdict remains NOT_EVALUATED for this scoped run.",
            "No baseline comparison or gain claim is made.",
            "Hover timing covers changed targets on the settled fixture through successful Present submission.",
            "Navigation timing covers synchronous retained-tree Control-wheel target-layout submission; it does not cover Up/Back/Forward or scanning navigation.",
            "These timings are submission observations, not physical input-to-photon latency.",
            "The 100 ms chrome-first-paint endpoint remains unqualified; no target is waived.",
        ],
    }


def main() -> None:
    report = verify_run()
    REPORT.parent.mkdir(parents=False, exist_ok=True)
    payload = (json.dumps(report, indent=2, sort_keys=True) + "\n").encode("utf-8")
    fd = os.open(REPORT, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(fd, "wb") as stream:
        stream.write(payload)
    print(f"Wrote verified read-only replay report: {REPORT}")


if __name__ == "__main__":
    main()

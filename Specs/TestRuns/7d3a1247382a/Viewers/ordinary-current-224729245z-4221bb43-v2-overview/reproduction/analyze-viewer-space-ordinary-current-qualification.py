#!/usr/bin/env python3
"""Analyze a receipt-pinned ordinary ViewerSpace small/large Release comparison."""

import hashlib
import json
import math
import pathlib
import re
import statistics
import sys


EXPECTED_FILES = {"small": 10_000, "large": 100_000}
EXPECTED_DIRECTORIES = {"small": 121, "large": 501}
PHASES = ("before-open", "peak-while-open", "complete-ready", "closed-after-quiet")
QUALIFICATION_SCOPE = "ordinary-current-x64-release-large-and-small-scan-comparison"
EXPECTED_CASES = ["viewer_space_perf_small_progressive", "viewer_space_perf_large_optin"]
EXPECTED_ARCHIVED_FILES = {
    "run-all-tests-results.json", "workspace-final.json", "commands/results.json",
    "perf/viewer_space_perf_small_progressive_metrics.json",
    "perf/viewer_space_perf_large_optin_metrics.json",
}
EXPECTED_HELPERS = {
    "run-viewer-space-ordinary-current-qualification.ps1",
    "curate-viewer-space-ordinary-current-evidence.py",
    "analyze-viewer-space-ordinary-current-qualification.py",
}
TEST_SANDBOX = pathlib.Path("D:/RedSalamander.Perf").resolve()


def fail(message):
    raise RuntimeError(message)


def read_json(path):
    with path.open("r", encoding="utf-8-sig") as stream:
        return json.load(stream, object_pairs_hook=unique_object)


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def validate_dxui_runtime_identity(manifest, campaign_root):
    runtime = manifest.get("dxui_runtime")
    if not isinstance(runtime, dict) or runtime.get("status") != "matched" or \
            runtime.get("platform") != "x64" or runtime.get("configuration") != "Release" or \
            set(runtime.get("sidecars", {})) != {"baseline", "candidate"}:
        fail("campaign does not pin a matched x64 Release DxUi runtime pair")
    semantic_fields = ("repository", "commit", "apiRevision", "platform", "configuration",
                       "publicHeadersGitTree", "archiveSha256", "module")
    sidecars = {}
    for side in ("baseline", "candidate"):
        entry = runtime["sidecars"][side]
        expected_relative = f"raw-sources/dxui/{side}.DxUi.json"
        relative = pathlib.PurePosixPath(entry.get("preserved_path", ""))
        if set(entry) != {"source_path", "preserved_path", "bytes", "sha256"} or \
                relative.is_absolute() or relative.as_posix() != expected_relative or \
                not isinstance(entry.get("source_path"), str) or not pathlib.Path(entry["source_path"]).is_absolute() or \
                type(entry.get("bytes")) is not int or entry["bytes"] <= 0 or \
                not isinstance(entry.get("sha256"), str) or not re.fullmatch(r"[0-9a-f]{64}", entry["sha256"]):
            fail(f"{side} DxUi sidecar manifest pin is incomplete or noncanonical")
        expected_source = (pathlib.Path(manifest.get(f"{side}_root", "")).resolve() /
                           ".build" / "x64" / "Release" / "RedSalamander.exe.DxUi.json").resolve()
        if pathlib.Path(entry["source_path"]).resolve() != expected_source:
            fail(f"{side} DxUi source path is not its pinned Release application sidecar")
        path = campaign_root.joinpath(*relative.parts)
        if not path.is_file() or path.is_symlink():
            fail(f"preserved {side} DxUi sidecar is missing or is a link")
        payload = path.read_bytes()
        digest = hashlib.sha256(payload).hexdigest()
        if len(payload) != entry["bytes"] or digest != entry["sha256"]:
            fail(f"preserved {side} DxUi sidecar length/SHA mismatch")
        try:
            value = json.loads(payload.decode("utf-8-sig"), object_pairs_hook=unique_object)
        except (UnicodeDecodeError, ValueError, RuntimeError) as exc:
            fail(f"preserved {side} DxUi sidecar is invalid JSON: {exc}")
        if not isinstance(value, dict) or set(value) != {
                "repository", "commit", "apiRevision", "platform", "configuration", "publicHeadersGitTree",
                "buildIdentity", "archiveSha256", "module", "moduleSha256", "linkCommandSha256"}:
            fail(f"preserved {side} DxUi sidecar schema is incomplete or unsupported")
        if not isinstance(value.get("repository"), str) or not value["repository"] or \
                not isinstance(value.get("commit"), str) or not re.fullmatch(r"[0-9a-f]{40}", value["commit"]) or \
                type(value.get("apiRevision")) is not int or value["apiRevision"] <= 0 or \
                value.get("platform") != "x64" or value.get("configuration") != "Release" or \
                not isinstance(value.get("publicHeadersGitTree"), str) or not re.fullmatch(r"[0-9a-f]{40}", value["publicHeadersGitTree"]) or \
                not isinstance(value.get("archiveSha256"), str) or not re.fullmatch(r"[0-9a-fA-F]{64}", value["archiveSha256"]) or \
                value.get("module") != "RedSalamander.exe" or \
                not isinstance(value.get("moduleSha256"), str) or not re.fullmatch(r"[0-9a-fA-F]{64}", value["moduleSha256"]) or \
                not isinstance(value.get("linkCommandSha256"), str) or not re.fullmatch(r"[0-9a-fA-F]{64}", value["linkCommandSha256"]) or \
                not isinstance(value.get("buildIdentity"), dict) or set(value["buildIdentity"]) != {"Fingerprint", "Identity"} or \
                not isinstance(value["buildIdentity"].get("Fingerprint"), str) or \
                not re.fullmatch(r"[0-9a-fA-F]{64}", value["buildIdentity"]["Fingerprint"]) or \
                not isinstance(value["buildIdentity"].get("Identity"), dict) or not value["buildIdentity"]["Identity"]:
            fail(f"preserved {side} DxUi sidecar has invalid dependency/build identity fields")
        sidecars[side] = value
    baseline, candidate = sidecars["baseline"], sidecars["candidate"]
    if any(baseline.get(field) != candidate.get(field) for field in semantic_fields) or \
            baseline.get("buildIdentity") != candidate.get("buildIdentity"):
        fail("baseline/candidate DxUi dependency, API, public-header, build-identity, platform, or configuration mismatch")
    dependency = {field: baseline[field] for field in (
        "repository", "commit", "apiRevision", "archiveSha256", "publicHeadersGitTree", "buildIdentity")}
    if runtime.get("dependency") != dependency or runtime.get("platform") != baseline["platform"] or \
            runtime.get("configuration") != baseline["configuration"]:
        fail("campaign DxUi dependency identity differs from its preserved sidecars")


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def digest_bytes(payload):
    return hashlib.sha256(payload).hexdigest()


def nearest_rank(values, percentile):
    ordered = sorted(values)
    return ordered[max(0, math.ceil(percentile * len(ordered)) - 1)]


def ratio(candidate, baseline):
    return None if baseline <= 0 or candidate < 0 else candidate / baseline


def hresult_succeeded(value):
    return (int(value) & 0x80000000) == 0


def load_metrics(path):
    rows = []
    with path.open("r", encoding="utf-8-sig") as stream:
        for line_number, line in enumerate(stream, 1):
            if line.strip():
                try:
                    rows.append(json.loads(line, object_pairs_hook=unique_object))
                except (json.JSONDecodeError, ValueError) as error:
                    fail(f"Invalid curated perf row {path}:{line_number}: {error}")
    return rows


def load_metrics_payload(payload, label):
    rows = []
    for line_number, line in enumerate(payload.splitlines(), 1):
        if line.strip():
            try:
                row = json.loads(line, object_pairs_hook=unique_object)
            except (json.JSONDecodeError, ValueError) as error:
                fail(f"Invalid/duplicate-key raw metric row {label}:{line_number}: {error}")
            if not isinstance(row, dict):
                fail(f"Raw metric row {label}:{line_number} is not an object")
            rows.append(row)
    return rows


def validate_campaign_identity(manifest, campaign_root):
    marker = TEST_SANDBOX / ".red-salamander-test-root"
    if not marker.is_file() or marker.read_text(encoding="utf-8-sig").strip() != "red-salamander.test-sandbox-root.v1" or \
            not campaign_root.is_relative_to(TEST_SANDBOX / "evidence") or \
            pathlib.Path(manifest.get("test_root", "")).resolve() != TEST_SANDBOX:
        fail("campaign manifest/evidence path is outside the already marked fixed TestSandbox root")
    if manifest.get("qualification_scope") != QUALIFICATION_SCOPE or manifest.get("configuration") != "Release" or \
            manifest.get("platform") != "x64" or manifest.get("case_filter") != "viewer_space_perf_" or \
            manifest.get("expected_cases") != EXPECTED_CASES or \
            manifest.get("expected_fixture_counts") != {
                "small": {"files": 10_000, "directories": 121},
                "large": {"files": 100_000, "directories": 501},
            }:
        fail("campaign qualification scope, platform, case, or exact fixture profile mismatch")
    validate_dxui_runtime_identity(manifest, campaign_root)
    if set(manifest.get("expected_builds", {})) != {"baseline", "candidate"}:
        fail("campaign does not pin both baseline and candidate builds")
    for side in ("baseline", "candidate"):
        build = manifest["expected_builds"][side]
        expected_root = pathlib.Path(manifest[f"{side}_root"]).resolve()
        build_root = pathlib.Path(build.get("repository", "")).resolve()
        receipt_path = pathlib.Path(build.get("receipt_path", "")).resolve()
        log_path = pathlib.Path(build.get("build_log_source_path", "")).resolve()
        if build.get("side") != side or build_root != expected_root or \
                not isinstance(build.get("receipt_id"), str) or not re.fullmatch(r"[0-9a-f]{64}", build["receipt_id"]) or \
                not isinstance(build.get("source_snapshot_id"), str) or not build["source_snapshot_id"] or \
                receipt_path != (expected_root / ".build" / "x64" / "Release" / "build-receipt.json").resolve() or \
                not log_path.is_relative_to((expected_root / ".build" / "logs").resolve()):
            fail(f"explicit {side} receipt/source snapshot/exact log path is missing")
        for archive_key, digest_key, bytes_key, expected_relative in (
                ("receipt_archive_path", "receipt_sha256", "receipt_bytes", f"build-inputs/{side}/build-receipt.json"),
                ("build_log_archive_path", "build_log_sha256", "build_log_bytes", f"build-inputs/{side}/msbuild.log")):
            relative = pathlib.PurePosixPath(build.get(archive_key, ""))
            if relative.is_absolute() or relative.as_posix() != expected_relative:
                fail(f"sealed {side} build input archive path is not canonical")
            sealed = campaign_root.joinpath(*relative.parts)
            if type(build.get(bytes_key)) is not int or not 0 < build[bytes_key] <= 2 * 1024 * 1024 or \
                    not sealed.is_file() or sealed.is_symlink() or sealed.stat().st_size != build.get(bytes_key) or \
                    sha256(sealed) != build.get(digest_key):
                fail(f"sealed {side} build input length/SHA mismatch: {archive_key}")
        receipt = read_json(campaign_root / build["receipt_archive_path"])
        if receipt.get("receipt_id") != build["receipt_id"] or \
                receipt.get("source_snapshot_id") != build["source_snapshot_id"] or \
                receipt.get("configuration") != "Release" or receipt.get("platform") != "x64" or \
                receipt.get("tests_enabled") is not True or type(receipt.get("exit_code")) is not int or receipt["exit_code"] != 0:
            fail(f"sealed {side} build receipt identity/profile mismatch")
    reproduction = manifest.get("reproduction_inputs")
    if not isinstance(reproduction, dict) or set(reproduction) != EXPECTED_HELPERS:
        fail("sealed reproduction helper set is not exactly launcher, curator, analyzer")
    for name, item in reproduction.items():
        relative = pathlib.PurePosixPath(item.get("path", ""))
        if relative.as_posix() != f"reproduction/{name}" or type(item.get("bytes")) is not int or item["bytes"] <= 0:
            fail(f"sealed reproduction helper path/length mismatch: {name}")
        sealed = campaign_root.joinpath(*relative.parts)
        if not sealed.is_file() or sealed.is_symlink() or sealed.stat().st_size != item["bytes"] or sha256(sealed) != item.get("sha256"):
            fail(f"sealed reproduction helper byte identity mismatch: {name}")
    analyzer_name = "analyze-viewer-space-ordinary-current-qualification.py"
    if sha256(pathlib.Path(__file__).resolve()) != reproduction[analyzer_name]["sha256"]:
        fail("executed analyzer differs from the campaign's sealed helper")


def validate_run_identity(record, provenance, manifest):
    side = record.get("side")
    expected_build = manifest.get("expected_builds", {}).get(side)
    if side not in ("baseline", "candidate") or expected_build is None or \
            record.get("build_receipt_id") != expected_build.get("receipt_id") or \
            record.get("build_receipt_path") != expected_build.get("receipt_path") or \
            record.get("build_receipt_sha256") != expected_build.get("receipt_sha256") or \
            record.get("build_source_snapshot_id") != expected_build.get("source_snapshot_id") or \
            record.get("build_log_source_path") != expected_build.get("build_log_source_path") or \
            record.get("build_log_sha256") != expected_build.get("build_log_sha256") or \
            record.get("build_log_bytes") != expected_build.get("build_log_bytes"):
        fail("run build identity differs from the exact side-specific receipt/source/log pins")
    if provenance.get("qualification_scope") != QUALIFICATION_SCOPE or provenance.get("side") != side or \
            provenance.get("mode") != "OrdinaryCurrent" or provenance.get("caseFilter") != "viewer_space_perf_" or \
            provenance.get("caseNames") != EXPECTED_CASES or \
            provenance.get("buildReceiptId") != expected_build.get("receipt_id") or \
            provenance.get("buildSourceSnapshotId") != expected_build.get("source_snapshot_id") or \
            provenance.get("buildReceiptSha256") != expected_build.get("receipt_sha256") or \
            provenance.get("buildLogSourcePath") != expected_build.get("build_log_source_path") or \
            provenance.get("buildLogBytes") != expected_build.get("build_log_bytes") or \
            provenance.get("buildLogSha256") != expected_build.get("build_log_sha256"):
        fail("run provenance side/mode/case/receipt/source/log mismatch")


def one_metric(rows, metric, predicate=lambda row: True):
    matches = [row for row in rows if row.get("metric") == metric and predicate(row)]
    if len(matches) != 1:
        fail(f"Expected one {metric} row, found {len(matches)}")
    return matches[0]


def load_run(record, campaign_root, manifest):
    if record.get("status") != "PASSED" or record.get("role") not in ("measurement", "warmup"):
        fail(f"Run record is not a passed warmup or measurement: {record.get('label')}")
    curated = pathlib.Path(record["curated_path"]).resolve()
    expected_curated = (campaign_root / "curated" / record.get("side", "") / record.get("label", "")).resolve()
    if curated != expected_curated or not curated.is_relative_to((campaign_root / "curated").resolve()):
        fail(f"Curated run is not at its exact side/label campaign path: {curated}")
    evidence = read_json(curated / "evidence.json")
    provenance = read_json(curated / "provenance.json")
    summary = read_json(curated / "run-all-tests-results.json")
    evidence_path = curated / "evidence.json"
    provenance_path = curated / "provenance.json"
    if evidence_path.stat().st_size != record.get("curated_evidence_bytes") or \
            sha256(evidence_path) != record.get("curated_evidence_sha256") or \
            provenance_path.stat().st_size != record.get("curated_provenance_bytes") or \
            sha256(provenance_path) != record.get("curated_provenance_sha256"):
        fail(f"Curated evidence/provenance does not match the launcher-pinned hash: {record.get('label')}")
    if type(record.get("curated_total_bytes")) is not int or not 0 < record["curated_total_bytes"] <= 5 * 1024 * 1024:
        fail(f"Curated run evidence exceeds or omits the 5 MiB run bound: {record.get('label')}")
    side = record.get("side")
    expected_build = manifest.get("expected_builds", {}).get(side)
    if evidence.get("sourceRunId") != record.get("run_id") or provenance.get("sourceRunId") != record.get("run_id"):
        fail(f"Curated run ID does not match manifest: {record.get('label')}")
    validate_run_identity(record, provenance, manifest)
    if provenance.get("qualificationStatus") != "qualified" or evidence.get("sourceRunId") != record.get("run_id"):
        fail(f"Curated evidence did not qualify: {record.get('label')}")
    for stream in ("stdout", "stderr"):
        expected_log = (campaign_root / "runner-logs" / f"{record.get('label')}.{stream}.txt").resolve()
        supplied_log = pathlib.Path(record.get(f"runner_{stream}_path", "")).resolve()
        if supplied_log != expected_log or not expected_log.is_file() or expected_log.is_symlink() or \
                expected_log.stat().st_size != record.get(f"runner_{stream}_bytes") or \
                expected_log.stat().st_size > 2 * 1024 * 1024 or \
                sha256(expected_log) != record.get(f"runner_{stream}_sha256"):
            fail(f"Campaign runner {stream} log identity/size/SHA mismatch: {record.get('label')}")
    if not isinstance(evidence.get("rawSourceIndex"), dict) or evidence.get("sourceSha256") != provenance.get("rawMetricsSha256") or \
            evidence.get("sourceBytes") != provenance.get("rawMetricsBytes") or \
            evidence["rawSourceIndex"].get("path") != provenance.get("rawSourceIndexPath") or \
            evidence.get("sourceSha256") != record.get("raw_metrics_sha256") or \
            evidence.get("files") != provenance.get("curatedFiles") or \
            provenance.get("workspaceSnapshotId") != record.get("workspace_snapshot_id") or \
            provenance.get("workspaceSnapshotSha256") != record.get("workspace_snapshot_sha256"):
        fail(f"Curated raw-source/file/workspace provenance mismatch: {record.get('label')}")
    if not isinstance(evidence.get("files"), dict) or set(evidence["files"]) != EXPECTED_ARCHIVED_FILES or \
            set(evidence.get("artifacts", {})) != set(EXPECTED_CASES):
        fail(f"Curated file/artifact set is incomplete or unexpected: {record.get('label')}")
    snapshot = read_json(curated / "workspace-final.json")
    if snapshot.get("snapshot_id") != provenance.get("workspaceSnapshotId") or \
            sha256(curated / "workspace-final.json") != provenance.get("workspaceSnapshotSha256"):
        fail(f"Curated workspace snapshot id/hash mismatch: {record.get('label')}")
    if summary.get("run_id") != record.get("run_id") or summary.get("case_filter") != "viewer_space_perf_" or \
            summary.get("platform") != "x64" or summary.get("configuration") != "Release":
        fail(f"Archived canonical summary run ID mismatch: {record.get('label')}")
    if sha256(curated / "run-all-tests-results.json") != record.get("summary_sha256"):
        fail(f"Archived canonical summary digest mismatch: {record.get('label')}")
    if read_json(campaign_root / expected_build["receipt_archive_path"]).get("receipt_id") != record.get("build_receipt_id"):
        fail(f"Sealed per-side receipt does not match run: {record.get('label')}")
    for logical_name, file_record in evidence.get("files", {}).items():
        relative = pathlib.PurePosixPath(file_record.get("path", ""))
        expected_relative = pathlib.PurePosixPath("curated", side, record.get("label", ""), logical_name)
        if relative.is_absolute() or ".." in relative.parts or relative != expected_relative:
            fail(f"Curated file path differs from its exact side/label location: {logical_name}")
        archived = campaign_root.joinpath(*relative.parts)
        if not archived.is_file() or archived.is_symlink() or archived.stat().st_size != file_record.get("bytes") or \
                sha256(archived) != file_record.get("sha256"):
            fail(f"Curated file length/SHA mismatch: {logical_name}")
    commands = read_json(curated / "commands" / "results.json")
    cases = commands.get("cases", [])
    expected_cases = {"viewer_space_perf_small_progressive", "viewer_space_perf_large_optin"}
    case_names = [case.get("name") for case in cases]
    if len(cases) != 2 or set(case_names) != expected_cases or case_names != [
            "viewer_space_perf_small_progressive", "viewer_space_perf_large_optin"]:
        fail(f"Archived Commands cases are not the exact small-then-large sequence for {record.get('label')}")
    if any(case.get("status") != "passed" for case in cases):
        fail(f"Archived Commands case did not pass for {record.get('label')}")

    source_index_record = evidence.get("rawSourceIndex")
    if not isinstance(source_index_record, dict):
        fail(f"Exact raw-source index is missing: {record.get('label')}")
    source_index_relative = pathlib.PurePosixPath(source_index_record.get("path", ""))
    if source_index_relative.is_absolute() or source_index_relative.parts != ("raw-sources", record.get("run_id"), "source-index.json"):
        fail("Raw-source index path is not campaign-relative")
    source_index_path = campaign_root.joinpath(*source_index_relative.parts)
    if not source_index_path.is_file() or source_index_path.is_symlink() or \
            source_index_path.stat().st_size != source_index_record.get("bytes") or \
            sha256(source_index_path) != source_index_record.get("sha256"):
        fail(f"Raw-source index length/SHA mismatch: {record.get('label')}")
    source_index = read_json(source_index_path)
    raw_parts = []
    for shard in source_index.get("shards", []):
        relative = pathlib.PurePosixPath(shard.get("path", ""))
        if relative.is_absolute() or relative.parts[:2] != ("raw-sources", record.get("run_id")) or \
                ".." in relative.parts:
            fail("Raw metric shard path is not in its exact run-owned campaign directory")
        shard_path = campaign_root.joinpath(*relative.parts)
        if not shard_path.is_file() or shard_path.is_symlink() or shard_path.stat().st_size != shard.get("bytes") or \
                sha256(shard_path) != shard.get("sha256"):
            fail(f"Raw metric shard length/SHA mismatch: {shard.get('path')}")
        raw_parts.append(shard_path.read_bytes())
    calculated_evidence_bytes = evidence_path.stat().st_size + provenance_path.stat().st_size + \
        source_index_record["bytes"] + sum(shard["bytes"] for shard in source_index.get("shards", [])) + \
        sum(item["bytes"] for item in evidence.get("files", {}).values())
    if calculated_evidence_bytes != record.get("curated_total_bytes"):
        fail(f"Curated run byte total does not match manifest: {record.get('label')}")
    raw_payload = b"".join(raw_parts)
    if source_index.get("sourceRunId", record.get("run_id")) != record.get("run_id") or \
            pathlib.Path(source_index.get("source_path", "")).resolve() != \
                (TEST_SANDBOX / "runs" / record.get("run_id") / "artifacts" / "selftest" / "last_run" / "perf" / "perf_metrics.jsonl").resolve() or \
            source_index.get("source_bytes") != evidence.get("sourceBytes") or \
            len(raw_payload) != evidence.get("sourceBytes") or digest_bytes(raw_payload) != evidence.get("sourceSha256") or \
            source_index.get("source_sha256") != evidence.get("sourceSha256"):
        fail(f"Exact raw metric source does not reconstruct for {record.get('label')}")
    rows = load_metrics_payload(raw_payload, record.get("label"))
    if len(rows) != evidence.get("sourceRows"):
        fail(f"Raw metric source row count mismatch: {record.get('label')}")
    if {row.get("machineHash") for row in rows if row.get("machineHash")} != {evidence.get("machineHash")}:
        fail(f"Raw metric machine identity differs from curated provenance: {record.get('label')}")
    scan_rows = [row for row in rows if row.get("metric") == "viewer.space.scan.total_us" and row.get("detail") == "done"]
    if len(scan_rows) != 2:
        fail(f"Expected one completed scan for each fixture in {record.get('label')}; got {len(scan_rows)}")
    scans = {}
    for row in scan_rows:
        file_count = int(row.get("value0", -1))
        folder_count = int(row.get("value1", -1))
        scenario = "small" if file_count == EXPECTED_FILES["small"] else "large" if file_count == EXPECTED_FILES["large"] else None
        if scenario is None or scenario in scans:
            fail(f"Unexpected/duplicate completed scan file count in {record.get('label')}: {file_count}")
        scans[scenario] = {"scan_total_us": int(row["durationUs"]), "scan_file_count": file_count,
                           "scan_folder_count": folder_count, "scan_detail": row.get("detail"),
                           "scan_hr": row.get("hr")}

    first_paint_rows = [row for row in rows if row.get("metric") == "viewer.space.open_to_first_paint_us"]
    if len(first_paint_rows) != 2:
        fail(f"Expected two production open-to-first-paint rows, got {len(first_paint_rows)} in {record.get('label')}")
    # The self-test case results below verify small then large order; this historical
    # production event carries no scenario tag, so preserve its observed order.
    for scenario, row in zip(("small", "large"), first_paint_rows):
        if "hr" not in row:
            fail(f"Historical first-paint event has no HRESULT in {record.get('label')}")
        scans[scenario]["historical_open_to_first_paint_us"] = int(row["durationUs"])
        scans[scenario]["historical_open_to_first_paint_hr"] = int(row["hr"])
        scans[scenario]["historical_open_to_first_paint_succeeded"] = hresult_succeeded(row["hr"])
        scans[scenario]["historical_first_paint_detail"] = row.get("detail")

    meaningful_rows = []
    if record["side"] == "candidate":
        meaningful_rows = [row for row in rows
                           if row.get("metric") == "viewer.space.open_to_first_meaningful_paint_us"
                           and row.get("detail") == "published-map"]
        if len(meaningful_rows) != 2:
            fail(f"Expected two candidate meaningful-paint events, found {len(meaningful_rows)} in {record.get('label')}")

    scenarios = {}
    artifact_digests = {}
    for scenario, artifact_name in (
        ("small", "viewer_space_perf_small_progressive_metrics.json"),
        ("large", "viewer_space_perf_large_optin_metrics.json"),
    ):
        artifact = read_json(curated / "perf" / artifact_name)
        artifact_digests[scenario] = sha256(curated / "perf" / artifact_name)
        artifact_record = evidence.get("artifacts", {}).get(
            "viewer_space_perf_small_progressive" if scenario == "small" else "viewer_space_perf_large_optin")
        if not isinstance(artifact_record, dict) or artifact_digests[scenario] != artifact_record.get("sha256") or \
                artifact_record.get("bytes") != (curated / "perf" / artifact_name).stat().st_size:
            fail(f"Scenario artifact does not match exact curated artifact record for {scenario} in {record.get('label')}")
        if artifact.get("real_directory_count") != EXPECTED_DIRECTORIES[scenario] or \
                artifact.get("file_candidate_count") != EXPECTED_FILES[scenario]:
            fail(f"Retained fixture counts differ for {scenario} in {record.get('label')}")
        if record["side"] == "candidate":
            meaningful_row = meaningful_rows[0 if scenario == "small" else 1]
            if "hr" not in meaningful_row or not hresult_succeeded(meaningful_row["hr"]):
                fail(f"Candidate meaningful-paint event did not succeed for {scenario} in {record.get('label')}")
            meaningful_us = int(artifact.get("first_meaningful_paint_us", 0))
            if meaningful_us <= 0 or int(meaningful_row.get("durationUs", -1)) != meaningful_us:
                fail(f"Candidate meaningful-paint artifact/metric disagree for {scenario} in {record.get('label')}")
            scans[scenario]["first_meaningful_paint_us"] = meaningful_us
            scans[scenario]["first_meaningful_paint_hr"] = int(meaningful_row["hr"])
            scans[scenario]["first_meaningful_paint_target_us"] = 250_000
            scans[scenario]["first_meaningful_paint_target_met"] = meaningful_us <= 250_000
        witness = artifact.get("cleanup_witness", {})
        if witness.get("satisfied") is not True or witness.get("scanWorkerCompletedCount") != 1 or \
                witness.get("scanWorkerAbandonedEventCount") != 0 or \
                witness.get("postCloseQuietUs", 0) < 100_000:
            fail(f"Cleanup/provider-lifetime witness failed for {scenario} in {record.get('label')}")
        if artifact.get("settled_before_deadline") is not True:
            fail(f"Historical scan-complete predicate failed for {scenario} in {record.get('label')}")
        if record["side"] == "baseline":
            if witness.get("readyKind") != "legacy-layout" or "baseline_layout_settle_wait_us" not in artifact:
                fail(f"Baseline complete-ready does not identify its legacy layout predicate: {record.get('label')}")
            product_ready_wait = int(artifact["baseline_layout_settle_wait_us"])
            product_ready_semantics = "legacy layout settled after scan: scan done, pending queue bytes zero, animation idle, non-empty draw list, and completed layout."
        else:
            if witness.get("readyKind") != "full-projection" or \
                    artifact.get("projections_settled_before_deadline") is not True or \
                    witness.get("deferredDisposalCompleted") is not True or \
                    witness.get("deferredDisposalCompletedCount", 0) < 1:
                fail(f"Candidate complete-ready full-projection predicate failed: {record.get('label')}")
            product_ready_wait = int(artifact["projection_settle_wait_us"])
            product_ready_semantics = "full projection settled after scan: scan done, pending queue bytes zero, cache restore idle, layout and analysis inactive, animation idle, and non-empty analysis/draw lists."
        memory = artifact.get("process_memory", {})
        if set(memory) != {"beforeOpen", "peakWhileOpen", "completeReady", "closedAfterQuiet"}:
            fail(f"Unexpected process_memory phases for {scenario} in {record.get('label')}")
        for phase in memory.values():
            if "workingSetBytes" not in phase or "privateBytes" not in phase:
                fail(f"Incomplete process memory sample for {scenario} in {record.get('label')}")
        before = memory["beforeOpen"]
        values = scans[scenario]
        values.update({
            "fixture_creation_us": int(artifact["fixture_creation_us"]),
            "harness_first_paint_wait_us": int(artifact["first_paint_wait_us"]),
            "historical_scan_settle_wait_us": int(artifact["settle_wait_us"]),
            "product_ready_wait_us": product_ready_wait,
            "product_ready_semantics": product_ready_semantics,
            "worker_completion_wait_us": int(artifact["worker_completion_wait_us"]),
            "first_paint_semantics": "legacy open_to_first_paint_us is scan-start based and retains its event HRESULT; harness_first_paint_wait_us is the test's separate bounded wait for a successful first-paint snapshot.",
            "real_directory_count": int(artifact["real_directory_count"]),
            "file_candidate_count": int(artifact["file_candidate_count"]),
            "cleanup_witness": witness,
            "process_memory": memory,
            "memory_deltas_bytes": {
                phase: {
                    key: int(memory[phase][key]) - int(before[key])
                    for key in ("workingSetBytes", "privateBytes")
                }
                for phase in ("peakWhileOpen", "completeReady", "closedAfterQuiet")
            },
        })
        scenarios[scenario] = values

    phase_metric_rows = [row for row in rows if row.get("metric") == "viewer.space.memory.process_phase"]
    for scenario in ("small", "large"):
        for phase in PHASES:
            expected_detail = f"viewer_space_perf_{'small_progressive' if scenario == 'small' else 'large_optin'}:{phase}"
            matches = [row for row in phase_metric_rows if row.get("detail") == expected_detail]
            if len(matches) != 1:
                fail(f"Missing/duplicate process-phase metric '{expected_detail}' in {record.get('label')}")
            metric_row = matches[0]
            artifact_phase = {
                "before-open": "beforeOpen", "peak-while-open": "peakWhileOpen",
                "complete-ready": "completeReady", "closed-after-quiet": "closedAfterQuiet",
            }[phase]
            values = scenarios[scenario]["process_memory"][artifact_phase]
            if int(metric_row.get("value0", -1)) != int(values["privateBytes"]) or \
                    int(metric_row.get("value1", -1)) != int(values["workingSetBytes"]):
                fail(f"Artifact/process-phase memory disagreement for '{expected_detail}' in {record.get('label')}")

    return {
        "label": record["label"], "side": record["side"], "run_id": record["run_id"],
        "block": record["block"], "position": record["position"],
        "build_receipt_id": record["build_receipt_id"],
        "summary_sha256": record["summary_sha256"],
        "machine_hash": evidence.get("machineHash"),
        "perf_metrics_preserved_sha256": evidence.get("sourceSha256"),
        "commands_results_sha256": sha256(curated / "commands" / "results.json"),
        "scenario_artifact_sha256": artifact_digests,
        "small": scenarios["small"], "large": scenarios["large"],
    }


def summarize(values, expected_count=6):
    if expected_count is not None and len(values) != expected_count:
        fail(f"Expected {expected_count} values, got {len(values)}")
    if not values:
        return {"n": 0, "median": None, "p95_nearest_rank": None,
                "p95_definition": "No valid observations were available."}
    return {"n": len(values), "median": statistics.median(values),
            "p95_nearest_rank": nearest_rank(values, 0.95),
            "p95_definition": "nearest-rank ceil(0.95*n); with n=6 this is the maximum observation, not a stable tail estimate."
            if len(values) == 6 else "nearest-rank ceil(0.95*n) over the reported valid-observation count."}


def metric_value(run, scenario, key):
    phase, field = key.split(".", 1)
    if field.endswith("Delta"):
        field = field[:-5]
        return int(run[scenario]["memory_deltas_bytes"][phase][field])
    return int(run[scenario]["process_memory"][phase][field])


def main():
    if len(sys.argv) != 2:
        fail("usage: analyze-viewer-space-apprelease-comparison.py <campaign-manifest.json>")
    manifest_path = pathlib.Path(sys.argv[1]).resolve()
    manifest = read_json(manifest_path)
    if manifest.get("schema") != "viewer-space.apprelease-comparison-campaign.v1":
        fail("Unsupported campaign manifest schema")
    campaign_root = manifest_path.parent.resolve()
    validate_campaign_identity(manifest, campaign_root)
    warmups = [record for record in manifest.get("runs", []) if record.get("role") == "warmup"]
    if len(warmups) != 2 or [(r.get("side"), r.get("status")) for r in warmups] != [
            ("baseline", "PASSED"), ("candidate", "PASSED")]:
        fail("Campaign must contain exactly the passed baseline then candidate warmups")
    for warmup in warmups:
        load_run(warmup, campaign_root, manifest)
    measurement_records = [record for record in manifest.get("runs", []) if record.get("role") == "measurement"]
    measurement_plan = manifest.get("measurements", [])
    if len(measurement_records) != 12 or len(measurement_plan) != 12:
        fail("Campaign must contain twelve passing measurements (six per side)")
    run_records_by_label = {record["label"]: record for record in measurement_records}
    if len(run_records_by_label) != 12 or len({record["label"] for record in measurement_plan}) != 12:
        fail("Campaign archive labels are not unique")
    loaded = {}
    for planned in measurement_plan:
        record = run_records_by_label.get(planned["label"])
        if record is None or record.get("run_id") != planned.get("run_id") or \
                record.get("side") != planned.get("side") or record.get("block") != planned.get("block") or \
                record.get("position") != planned.get("position"):
            fail(f"Measurement plan does not match archived runner record: {planned.get('label')}")
        loaded[record["label"]] = load_run(record, campaign_root, manifest)
    if sum(item["side"] == "baseline" for item in loaded.values()) != 6 or \
            sum(item["side"] == "candidate" for item in loaded.values()) != 6:
        fail("Campaign measurement side counts are not six baseline and six candidate")
    machine_hashes = {item["machine_hash"] for item in loaded.values()}
    if len(machine_hashes) != 1 or None in machine_hashes:
        fail(f"Expected one shared machine identity across the comparison, got {sorted(str(value) for value in machine_hashes)}")
    for side in ("baseline", "candidate"):
        receipts = {item["build_receipt_id"] for item in loaded.values() if item["side"] == side}
        if len(receipts) != 1:
            fail(f"{side} build receipt changed during measured samples")

    metric_keys = ("scan_total_us", "harness_first_paint_wait_us",
                   "peakWhileOpen.privateBytesDelta", "peakWhileOpen.workingSetBytesDelta",
                   "completeReady.privateBytesDelta", "completeReady.workingSetBytesDelta",
                   "closedAfterQuiet.privateBytesDelta", "closedAfterQuiet.workingSetBytesDelta")
    output = {
        "schema": "viewer-space.apprelease-comparison-analysis.v1",
        "campaign_id": manifest["campaign_id"],
        "configuration": "x64 Release",
        "candidate_functionality_note": "Candidate complete-ready includes its full delivered projection; baseline complete-ready is its complete product view. This comparison measures each product's complete ready view and does not attribute that additional candidate functionality to a regression or gain.",
        "first_paint_note": "Historical open_to_first_paint_us is retained with its HRESULT because it is scan-start based and may capture an unsuccessful paint attempt; only SUCCEEDED HRESULT observations enter its summary. Candidate-only open_to_first_meaningful_paint_us starts at the valid IViewerOpen entry (including control creation) and ends at the first S_OK Present of non-empty published geometry with layout complete; it is evaluated against the absolute 250 ms target without inventing a baseline value.",
        "p95_note": "Nearest-rank p95 with six samples equals the maximum; it is reported descriptively and should not be read as a statistically stable tail estimate.",
        "measurement_order": [record["label"] for record in measurement_plan],
        "build_receipts": {side: next(iter({item["build_receipt_id"] for item in loaded.values() if item["side"] == side}))
                           for side in ("baseline", "candidate")},
        "build_identities": {side: {
            "receipt_id": manifest["expected_builds"][side]["receipt_id"],
            "source_snapshot_id": manifest["expected_builds"][side]["source_snapshot_id"],
            "build_log_source_path": manifest["expected_builds"][side]["build_log_source_path"],
            "build_log_bytes": manifest["expected_builds"][side]["build_log_bytes"],
            "build_log_sha256": manifest["expected_builds"][side]["build_log_sha256"],
        } for side in ("baseline", "candidate")},
        "runs": list(loaded.values()),
        "scenario_summaries": {},
        "abba_adjacent_pairs": [],
        "historical_open_to_first_paint": {},
        "candidate_only_first_meaningful_paint": {
            "metric": "viewer.space.open_to_first_meaningful_paint_us",
            "target_us": 250_000,
            "target_ms": 250,
            "baseline_value": None,
            "baseline_note": "Not measured by the baseline build; no value is inferred.",
            "scenarios": {},
        },
        "claim": "Descriptive measurements only. No positive candidate performance gain is asserted by this script.",
    }

    for scenario in ("small", "large"):
        output["scenario_summaries"][scenario] = {}
        for side in ("baseline", "candidate"):
            side_runs = [item for item in loaded.values() if item["side"] == side]
            output["scenario_summaries"][scenario][side] = {}
            for key in metric_keys:
                if "." in key:
                    vals = [metric_value(item, scenario, key) for item in side_runs]
                else:
                    vals = [int(item[scenario][key]) for item in side_runs]
                output["scenario_summaries"][scenario][side][key] = summarize(vals)

            paint_events = [{"run_id": item["run_id"],
                             "duration_us": int(item[scenario]["historical_open_to_first_paint_us"]),
                             "hr": int(item[scenario]["historical_open_to_first_paint_hr"]),
                             "succeeded": bool(item[scenario]["historical_open_to_first_paint_succeeded"])}
                            for item in side_runs]
            successful_values = [event["duration_us"] for event in paint_events if event["succeeded"]]
            output["historical_open_to_first_paint"].setdefault(scenario, {})[side] = {
                "events": paint_events,
                "successful_observation_summary": ({
                    "n": len(successful_values),
                    "median": statistics.median(successful_values),
                    "p95_nearest_rank": nearest_rank(successful_values, 0.95),
                    "p95_definition": "nearest-rank ceil(0.95*n); unsuccessful HRESULT rows remain in events but are excluded from this timing summary.",
                } if successful_values else {"n": 0, "median": None, "p95_nearest_rank": None}),
            }

        candidate_values = [int(item[scenario]["first_meaningful_paint_us"])
                            for item in loaded.values() if item["side"] == "candidate"]
        output["candidate_only_first_meaningful_paint"]["scenarios"][scenario] = {
            "baseline": None,
            "candidate": {
                **summarize(candidate_values),
                "target_us": 250_000,
                "target_met_count": sum(value <= 250_000 for value in candidate_values),
                "target_met_rate": sum(value <= 250_000 for value in candidate_values) / len(candidate_values),
                "measurements_us": candidate_values,
            },
        }

    measurements_by_block = {}
    for planned in measurement_plan:
        measurements_by_block.setdefault(int(planned["block"]), []).append(planned)
    for block in range(1, 4):
        block_entries = sorted(measurements_by_block.get(block, []), key=lambda item: int(item["position"]))
        if len(block_entries) != 4 or [int(item["position"]) for item in block_entries] != [1, 2, 3, 4] or \
                [item["side"] for item in block_entries] != ["baseline", "candidate", "candidate", "baseline"]:
            fail(f"Block {block} does not follow the prescribed ABBA order")
        for pair_name, baseline_plan, candidate_plan in (
                ("left-adjacent", block_entries[0], block_entries[1]),
                ("right-adjacent", block_entries[3], block_entries[2])):
            baseline = loaded[baseline_plan["label"]]
            candidate = loaded[candidate_plan["label"]]
            pair = {"block": block, "pair": pair_name,
                    "baseline_label": baseline["label"], "candidate_label": candidate["label"],
                    "baseline_run_id": baseline["run_id"], "candidate_run_id": candidate["run_id"],
                    "scenarios": {}}
            for scenario in ("small", "large"):
                pair["scenarios"][scenario] = {}
                for key in metric_keys:
                    if "." in key:
                        before = metric_value(baseline, scenario, key)
                        after = metric_value(candidate, scenario, key)
                    else:
                        before = int(baseline[scenario][key])
                        after = int(candidate[scenario][key])
                    pair["scenarios"][scenario][key] = {
                        "baseline": before, "candidate": after,
                        "candidate_over_baseline": ratio(after, before),
                        "relative_change": None if before <= 0 or after < 0 else (after - before) / before,
                    }
                baseline_paint = baseline[scenario]
                candidate_paint = candidate[scenario]
                both_historical_succeeded = baseline_paint["historical_open_to_first_paint_succeeded"] and \
                    candidate_paint["historical_open_to_first_paint_succeeded"]
                baseline_historical_us = int(baseline_paint["historical_open_to_first_paint_us"])
                candidate_historical_us = int(candidate_paint["historical_open_to_first_paint_us"])
                pair["scenarios"][scenario]["historical_open_to_first_paint"] = {
                    "baseline_us": baseline_historical_us,
                    "baseline_hr": baseline_paint["historical_open_to_first_paint_hr"],
                    "baseline_succeeded": baseline_paint["historical_open_to_first_paint_succeeded"],
                    "candidate_us": candidate_historical_us,
                    "candidate_hr": candidate_paint["historical_open_to_first_paint_hr"],
                    "candidate_succeeded": candidate_paint["historical_open_to_first_paint_succeeded"],
                    "candidate_over_baseline": ratio(candidate_historical_us, baseline_historical_us)
                        if both_historical_succeeded else None,
                    "ratio_omission_reason": None if both_historical_succeeded else "At least one historical event had an unsuccessful HRESULT.",
                }
                pair["scenarios"][scenario]["candidate_only_first_meaningful_paint"] = {
                    "baseline_us": None,
                    "candidate_us": int(candidate_paint["first_meaningful_paint_us"]),
                    "target_us": 250_000,
                    "target_met": bool(candidate_paint["first_meaningful_paint_target_met"]),
                }
            output["abba_adjacent_pairs"].append(pair)

    output["paired_ratio_summaries"] = {}
    for scenario in ("small", "large"):
        output["paired_ratio_summaries"][scenario] = {}
        for key in metric_keys:
            ratios = [pair["scenarios"][scenario][key]["candidate_over_baseline"]
                      for pair in output["abba_adjacent_pairs"]]
            valid_ratios = [value for value in ratios if value is not None]
            if not valid_ratios:
                output["paired_ratio_summaries"][scenario][key] = {
                    "n": 0, "median": None, "p95_nearest_rank": None,
                    "n_pairs": len(ratios), "ratios": ratios,
                    "note": "All pairs have a non-positive baseline or negative candidate measurement; ratios are not reported."}
            else:
                output["paired_ratio_summaries"][scenario][key] = {
                    **summarize(valid_ratios, expected_count=len(valid_ratios)), "n_pairs": len(ratios),
                    "omitted_nonpositive_pairs": len(ratios) - len(valid_ratios), "ratios": ratios,
                    "direction": "candidate / baseline; values below 1 are lower measurements and are not by themselves a performance-gain claim."}

        historical_ratios = [pair["scenarios"][scenario]["historical_open_to_first_paint"]["candidate_over_baseline"]
                             for pair in output["abba_adjacent_pairs"]]
        valid_historical_ratios = [value for value in historical_ratios if value is not None]
        output["historical_open_to_first_paint"][scenario]["paired_candidate_over_baseline"] = {
            **({**summarize(valid_historical_ratios, expected_count=len(valid_historical_ratios)), "n_pairs": len(historical_ratios),
                "omitted_unsuccessful_hresult_pairs": len(historical_ratios) - len(valid_historical_ratios)}
               if valid_historical_ratios else
               {"n": 0, "n_pairs": len(historical_ratios), "median": None, "p95_nearest_rank": None,
                "omitted_unsuccessful_hresult_pairs": len(historical_ratios)}),
            "ratios": historical_ratios,
            "definition": "candidate / baseline only when both historical first-paint rows had SUCCEEDED HRESULTs; unsuccessful attempts remain listed in per-run events.",
        }

    machine_hash = next(iter(machine_hashes))
    destination = campaign_root / "analysis.json"
    if destination.exists():
        fail(f"Refusing to replace immutable campaign analysis: {destination}")
    destination.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"analysis": str(destination), "measurement_runs": len(loaded),
                      "paired_contrasts": len(output["abba_adjacent_pairs"]),
                      "claim": output["claim"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()

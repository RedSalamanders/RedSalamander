#!/usr/bin/env python3
"""Fail-closed analysis for controlled same-IViewer four-root retained-reference memory ABBA runs."""
import hashlib
import json
import pathlib
import re
import statistics
import sys


def fail(message):
    raise RuntimeError(message)


def read_json(path):
    def unique_object(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate JSON object key: {key}")
            result[key] = value
        return result
    return json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_object)


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


REFERENCE_SCOPE = "controlled-same-IViewer-reference-not-historical-S0"
REFERENCE_CASE = "cmd_viewer_space_memory_heavy_four_root_retained_reference_optin"
REFERENCE_CONFIGURATION = "topK512_threads8_volume1_cacheTTL600_entries4"
REFERENCE_HELPERS = {
    "run-viewer-space-retained-reference-qualification.ps1",
    "curate-viewer-space-retained-reference-run.py",
    "analyze-viewer-space-retained-reference-campaign.py",
}
SCAN_STATE_MIN = 0
SCAN_STATE_MAX = 5  # Common::WindowMessages.h: NotStarted through Canceled, inclusive.


def validate_dxui_runtime_identity(manifest, root):
    runtime = manifest.get("dxui_runtime")
    if not isinstance(runtime, dict) or runtime.get("status") != "matched" or \
            runtime.get("platform") != "x64" or runtime.get("configuration") != "Release" or \
            set(runtime.get("sidecars", {})) != {"baseline", "candidate"}:
        fail("campaign does not pin a matched x64 Release DxUi runtime pair")

    def unique_object(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate JSON object key: {key}")
            result[key] = value
        return result

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
        log_path = manifest.get("expected_build_log_paths", {}).get(side)
        if not isinstance(log_path, str) or not pathlib.Path(log_path).is_absolute():
            fail(f"{side} DxUi source cannot be tied to an exact pinned Release build log")
        expected_repo = pathlib.Path(log_path).resolve().parents[2]
        expected_source = (expected_repo / ".build" / "x64" / "Release" / "RedSalamander.exe.DxUi.json").resolve()
        if pathlib.Path(entry["source_path"]).resolve() != expected_source:
            fail(f"{side} DxUi source path is not its pinned Release application sidecar")
        path = root.joinpath(*relative.parts)
        if not path.is_file() or path.is_symlink():
            fail(f"preserved {side} DxUi sidecar is missing or is a link")
        payload = path.read_bytes()
        if len(payload) != entry["bytes"] or hashlib.sha256(payload).hexdigest() != entry["sha256"]:
            fail(f"preserved {side} DxUi sidecar length/SHA mismatch")
        try:
            value = json.loads(payload.decode("utf-8-sig"), object_pairs_hook=unique_object)
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
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


def is_positive_integer(value):
    return type(value) is int and value > 0


def is_scan_state(value):
    return type(value) is int and SCAN_STATE_MIN <= value <= SCAN_STATE_MAX


def validate_reference_manifest_identity(manifest, root):
    validate_dxui_runtime_identity(manifest, root)
    if manifest.get("mode") != "RetainedReference" or \
            manifest.get("qualification_scope") != REFERENCE_SCOPE or \
            manifest.get("case_filter") != REFERENCE_CASE or \
            manifest.get("platform") != "x64" or manifest.get("configuration") != "Release":
        fail("controlled-reference mode, scope, case, platform, or configuration identity mismatch")
    for key in ("baseline", "candidate"):
        if not isinstance(manifest.get("expected_receipts", {}).get(key), str) or \
                not manifest["expected_receipts"][key] or \
                not isinstance(manifest.get("expected_build_source_snapshot_ids", {}).get(key), str) or \
                not manifest["expected_build_source_snapshot_ids"][key] or \
                not isinstance(manifest.get("expected_build_log_paths", {}).get(key), str) or \
                not manifest["expected_build_log_paths"][key]:
            fail(f"pinned {key} receipt, source snapshot, or build-log path is missing")

    reproduction = manifest.get("reproduction")
    if not isinstance(reproduction, dict) or set(reproduction) != REFERENCE_HELPERS:
        fail("sealed reproduction helper set does not match the exact three controlled-reference inputs")
    campaign_root = root.resolve()
    for name in sorted(REFERENCE_HELPERS):
        item = reproduction[name]
        if not isinstance(item, dict) or item.get("path") != f"reproduction\\{name}" or \
                not is_positive_integer(item.get("bytes")) or \
                not isinstance(item.get("sha256"), str) or not re.fullmatch(r"[0-9a-f]{64}", item["sha256"]):
            fail(f"sealed reproduction helper record is incomplete or has a noncanonical path: {name}")
        helper_path = campaign_root / "reproduction" / name
        if not helper_path.is_file() or helper_path.is_symlink():
            fail(f"sealed reproduction helper is missing or linked: {name}")
        stat_result = helper_path.stat()
        if getattr(stat_result, "st_file_attributes", 0) & 0x400:
            fail(f"sealed reproduction helper is a reparse point: {name}")
        if stat_result.st_size != item["bytes"] or sha256(helper_path) != item["sha256"]:
            fail(f"sealed reproduction helper size or SHA-256 mismatch: {name}")
    executed_analyzer = pathlib.Path(__file__).resolve()
    if sha256(executed_analyzer) != reproduction["analyze-viewer-space-retained-reference-campaign.py"]["sha256"]:
        fail("executed analyzer differs from the campaign's sealed analyzer helper")


def validate_reference_record_identity(manifest, record, provenance):
    side = record.get("side")
    if side not in ("baseline", "candidate") or \
            record.get("receipt_id") != manifest.get("expected_receipts", {}).get(side) or \
            provenance.get("side") != side or provenance.get("mode") != "RetainedReference" or \
            provenance.get("case") != REFERENCE_CASE or \
            provenance.get("buildReceiptId") != record.get("receipt_id"):
        fail(f"controlled-reference record/provenance/pinned receipt identity mismatch: {record.get('label')}")


def validate_reference_profile(artifact, label):
    if artifact.get("configuration") != REFERENCE_CONFIGURATION or \
            artifact.get("rootCount") != 4 or artifact.get("filesPerRoot") != 100000 or \
            artifact.get("directoriesPerRoot") != 501 or artifact.get("effectiveScanThreads") != 8:
        fail(f"fixture/profile mismatch in {label}")
    for stage in ("processBefore", "processPeak", "processAfter"):
        sample = artifact.get(stage)
        if not isinstance(sample, dict) or any(not is_positive_integer(sample.get(field)) for field in ("privateBytes", "workingSetBytes")):
            fail(f"{stage} must contain positive integer private/working-set samples in {label}")
    for field in ("privateBytes", "workingSetBytes"):
        if artifact["processPeak"][field] < artifact["processBefore"][field]:
            fail(f"peak {field} is below the before sample in {label}")
    if not is_positive_integer(artifact.get("rootTotalBytes")):
        fail(f"root byte total must be a positive integer in {label}")


def validate_reference_scan_state(value, label):
    if not is_scan_state(value):
        fail(f"reference F5 scan state must be an integer enum value 0..5 in {label}")


def nearest_rank(values, percentile):
    ordered = sorted(values)
    if not ordered:
        return None
    index = max(0, min(len(ordered) - 1, int((percentile * len(ordered) + 0.999999999999) - 1)))
    return ordered[index]


def run_record(manifest, record, root):
    if record.get("status") != "PASSED":
        fail(f"run was not passed: {record.get('label')}")
    archive = (root / "runs" / record["label"]).resolve()
    if not archive.is_relative_to(root.resolve()):
        fail("run archive escapes campaign root")
    provenance = read_json(archive / "provenance.json")
    if manifest.get("mode") == "RetainedReference":
        validate_reference_record_identity(manifest, record, provenance)
    if provenance.get("runId") != record.get("run_id") or provenance.get("buildReceiptId") != record.get("receipt_id"):
        fail(f"run provenance identity mismatch: {record.get('label')}")
    if provenance.get("qualificationStatus") != "qualified":
        fail(f"run evidence was preserved but did not qualify: {record.get('label')}")
    if provenance.get("buildSourceSnapshotId") != manifest["expected_build_source_snapshot_ids"][record["side"]]:
        fail(f"build source_snapshot_id changed: {record.get('label')}")
    if not provenance.get("workspaceSnapshotId"):
        fail(f"independent workspace-final snapshot_id is missing: {record.get('label')}")
    receipt_path = archive / "build-receipt.json"
    receipt = read_json(receipt_path)
    if receipt.get("receipt_id") != record.get("receipt_id") or receipt.get("source_snapshot_id") != provenance.get("buildSourceSnapshotId"):
        fail(f"archived immutable build receipt mismatch: {record.get('label')}")
    if sha256(receipt_path) != record.get("receipt_sha256"):
        fail(f"archived build receipt SHA mismatch: {record.get('label')}")
    archived_build_index = read_json(archive / "build-log-index.json")
    expected_log_path = pathlib.Path(manifest.get("expected_build_log_paths", {}).get(record["side"], "")).resolve()
    recorded_log_path = pathlib.Path(record.get("build_log_source_path", "")).resolve()
    indexed_log_path = pathlib.Path(archived_build_index.get("source_path", "")).resolve()
    if archived_build_index.get("source_sha256") != record.get("build_log_sha256") or \
            archived_build_index.get("source_bytes") != record.get("build_log_bytes") or \
            recorded_log_path != expected_log_path or indexed_log_path != expected_log_path:
        fail(f"exact expected build-log path/length/SHA does not match the per-run manifest: {record.get('label')}")
    build_index_path = root / "build-inputs" / record["side"] / "msbuild-log-index.json"
    build_index = read_json(build_index_path)
    build_payload = b"".join(pathlib.Path(chunk["path"]).read_bytes() for chunk in build_index.get("chunks", []))
    if hashlib.sha256(build_payload).hexdigest() != build_index.get("source_sha256") or len(build_payload) != build_index.get("source_bytes"):
        fail(f"sealed MSBuild log is not byte-exact: {record.get('label')}")
    if sha256(build_index_path) != sha256(archive / "build-log-index.json"):
        fail(f"archived build-log index digest mismatch: {record.get('label')}")
    for key, stem in (("stdout", "runner-stdout"), ("stderr", "runner-stderr")):
        evidence = provenance.get("runner" + key.title(), {})
        log_payload = b"".join((archive / name).read_bytes() for name in evidence.get("shards", []))
        if hashlib.sha256(log_payload).hexdigest() != record.get(f"{key}_sha256") or \
                len(log_payload) != record.get(f"{key}_bytes"):
            fail(f"runner {key} log chunks/digest mismatch: {record.get('label')}")
    workspace_path = archive / "workspace-final.json"
    workspace = read_json(workspace_path)
    if workspace.get("snapshot_id") != provenance.get("workspaceSnapshotId") or sha256(workspace_path) != provenance.get("workspaceSnapshotSha256"):
        fail(f"independent workspace snapshot evidence mismatch: {record.get('label')}")
    metric_info = provenance.get("metrics", {})
    raw_source = provenance.get("rawMetricsSourceIndex")
    if not isinstance(raw_source, dict):
        fail(f"campaign raw-source identity is missing: {record.get('label')}")
    preserved_path = pathlib.Path(provenance.get("rawMetricsPreservedSourcePath", "")).resolve()
    expected_preserved = (root / "raw-sources" / record["label"] / "perf_metrics.jsonl").resolve()
    source_index_path = (archive / "raw-metrics-source-index.json").resolve()
    if preserved_path != expected_preserved or not preserved_path.is_file() or not source_index_path.is_file():
        fail(f"campaign-preserved raw metrics/source index is missing or misplaced: {record.get('label')}")
    preserved = preserved_path.read_bytes()
    archived_source_index = read_json(source_index_path)
    expected_original = pathlib.Path(provenance.get("rawMetricsOriginalSourcePath", "")).resolve()
    if archived_source_index != raw_source or pathlib.Path(raw_source.get("original_path", "")).resolve() != expected_original or \
            pathlib.Path(raw_source.get("preserved_path", "")).resolve() != preserved_path or \
            pathlib.Path(record.get("raw_metrics_original_path", "")).resolve() != expected_original or \
            pathlib.Path(record.get("raw_metrics_preserved_path", "")).resolve() != preserved_path or \
            raw_source.get("bytes") != len(preserved) or raw_source.get("bytes") != record.get("raw_metrics_bytes") or \
            raw_source.get("sha256") != hashlib.sha256(preserved).hexdigest() or raw_source.get("sha256") != record.get("raw_metrics_sha256") or \
            len(preserved) != provenance.get("rawMetricsBytes") or hashlib.sha256(preserved).hexdigest() != provenance.get("rawMetricsSha256"):
        fail(f"campaign-preserved raw metrics fail their original/preserved path, size, or digest identity: {record.get('label')}")
    if metric_info.get("preservation") == "full-byte-exact":
        shard_paths = sorted(archive.glob("perf-metrics*.jsonl"),
                             key=lambda p: 0 if p.name == "perf-metrics.jsonl" else int(p.stem.removeprefix("perf-metrics-part")))
        payload = b"".join(path.read_bytes() for path in shard_paths)
        if hashlib.sha256(payload).hexdigest() != provenance.get("rawMetricsSha256"):
            fail(f"full raw metric payload digest mismatch: {record.get('label')}")
    elif metric_info.get("preservation") == "curated-with-source-replay-map":
        replay = read_json(archive / "metric-replay-map.json")
        curated = (archive / "perf-metrics-curated.jsonl").read_bytes()
        if hashlib.sha256(curated).hexdigest() != replay.get("retainedSha256") or \
                replay.get("sourceSha256") != provenance.get("rawMetricsSha256") or \
                replay.get("retainedRows") != len(replay.get("retainedSourceLineNumbers", [])):
            fail(f"curated metric source replay map mismatch: {record.get('label')}")
        payload = curated
        original = preserved
        original_lines = original.splitlines()
        mapped = [original_lines[int(number) - 1] for number in replay["retainedSourceLineNumbers"]]
        if hashlib.sha256(original).hexdigest() != replay.get("sourceSha256") or \
                b"\n".join(mapped) + (b"\n" if mapped else b"") != curated:
            fail(f"curated metrics do not replay exactly from the complete raw source: {record.get('label')}")
    else:
        fail(f"metric evidence preservation mode is unknown: {record.get('label')}")
    artifact_name = "viewer_space_memory_heavy_four_root_retained_reference_metrics.json" if manifest["mode"] == "RetainedReference" else "viewer_space_cached_reopen_200_first_map.json"
    artifact_path = archive / artifact_name
    if sha256(artifact_path) != provenance.get("artifactSha256"):
        fail(f"scenario artifact digest mismatch: {record.get('label')}")
    return provenance, read_json(artifact_path), payload


def write_cached_report(manifest, root, output):
    records = manifest.get("runs", [])
    if len(records) != 1 or records[0].get("side") != "candidate" or records[0].get("role") != "measurement":
        fail("cached-200 qualification expects exactly one candidate measurement")
    provenance, artifact, _ = run_record(manifest, records[0], root)
    samples = artifact.get("samples", [])
    if artifact.get("status") != "passed" or len(samples) != 200 or artifact.get("completedSampleCount") != 200 or \
            artifact.get("cacheHitCount") != 200 or artifact.get("fixtureFiles") != 10000 or \
            artifact.get("fixtureDirectories") != 121 or artifact.get("fixtureRootBytes", 0) <= 0:
        fail("cached-200 artifact lacks 200 completed samples")
    values = [int(row["firstMeaningfulPaintUs"]) for row in samples]
    if any(row.get("status") != "passed" or row.get("layoutGeneration", 0) <= 0 or
           row.get("layoutGeneration") != row.get("presentedLayoutGeneration") or
           row.get("cacheHitCount") != index or row.get("cacheRestoreBytes", 0) <= 0 or
           row.get("cacheResidentBytes", 0) <= 0 or
           row.get("firstMeaningfulPaintUs", 0) <= 0 for index, row in enumerate(samples, 1)):
        fail("cached-200 contains a sample without a clean current presented generation and cache hit")
    report = {
        "schema": "viewer-space.cached-reopen-200-analysis.v1", "status": "qualified",
        "runId": records[0]["run_id"], "receiptId": records[0]["receipt_id"],
        "workspaceSnapshotId": provenance["workspaceSnapshotId"], "fixtureFiles": artifact["fixtureFiles"],
        "fixtureDirectories": artifact["fixtureDirectories"], "fixtureRootBytes": artifact["fixtureRootBytes"],
        "cacheHitCount": artifact["cacheHitCount"], "sampleCount": len(values),
        "firstMeaningfulPaintUs": {"min": min(values), "median": statistics.median(values),
                                    "p95": nearest_rank(values, .95), "max": max(values),
                                    "targetUs": 100000, "targetMetCount": sum(v <= 100000 for v in values)},
        "allSamplesCurrentPresented": True, "observerMode": artifact.get("observerMode"),
        "observerModeInterpretation": "Observer mode is reported as recorded; samples from different observer modes are not combined.",
        "interpretation": "candidate-only cached reopen samples; no baseline comparison or cache-phase-cap override",
        "sourceArtifactSha256": provenance["artifactSha256"],
    }
    output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def write_candidate_heavy_report(manifest, root, output):
    records = manifest.get("runs", [])
    if len(records) != 1 or records[0].get("side") != "candidate" or records[0].get("role") != "candidate-heavy" or \
            records[0].get("block") != 0 or records[0].get("position") != 1:
        fail("CandidateHeavy requires exactly one candidate-only candidate-heavy run at block 0, position 1")
    if manifest.get("qualification_scope") != "candidate-only-four-root-lifecycle" or \
            records[0].get("receipt_id") != manifest.get("expected_receipts", {}).get("candidate"):
        fail("CandidateHeavy scope or pinned candidate receipt identity is missing/mismatched")
    if manifest.get("measurements") or any(record.get("role") == "warmup" for record in records):
        fail("CandidateHeavy must not contain warmups or comparative measurement entries")
    if manifest.get("expected_receipts", {}).get("baseline") is not None or \
            manifest.get("expected_build_source_snapshot_ids", {}).get("baseline") is not None or \
            manifest.get("expected_build_log_paths", {}).get("baseline") is not None:
        fail("CandidateHeavy must not pin baseline build inputs")
    record = records[0]
    provenance, artifact, _ = run_record(manifest, record, root)
    digests = artifact.get("rootFixtureDigests", [])
    if artifact.get("scenario") != "viewer_space_memory_heavy_four_root_cache_refresh" or \
            artifact.get("rootCount") != 4 or artifact.get("filesPerRoot") != 100000 or \
            artifact.get("directoriesPerRoot") != 501 or artifact.get("effectiveScanThreads") != 8 or \
            artifact.get("configuration") != "topK512_threads8_volume1_cacheTTL600_entries4":
        fail("CandidateHeavy artifact fixture/profile mismatch")
    if len(digests) != 4 or len(set(digests)) != 1 or artifact.get("cacheHitCount") != 4 or \
            len(artifact.get("phases", [])) != 24 or int(artifact.get("rootTotalBytes", 0)) <= 0:
        fail("CandidateHeavy artifact lacks four identical fixtures, four cache hits, or all 24 lifecycle rows")
    phase_names = ("seed-complete", "seed-closed", "cached-reopen", "refresh-retained-pending", "refresh-adopted", "refresh-closed")
    phases = artifact["phases"]
    for root_number in range(1, 5):
        rows = [row for row in phases if row.get("root") == root_number]
        if len(rows) != 6 or {row.get("phase") for row in rows} != set(phase_names):
            fail(f"CandidateHeavy root {root_number} does not have the six expected phase observations")
        if any(int(row.get(field, -1)) < 0 for row in rows for field in ("phaseUs", "privateBytes", "workingSetBytes", "rootBytes", "firstMeaningfulPaintUs")):
            fail(f"CandidateHeavy root {root_number} contains a negative or missing phase observation")
        if any(int(row.get("rootBytes", -1)) != int(artifact["rootTotalBytes"]) for row in rows):
            fail(f"CandidateHeavy root {root_number} phase byte total differs from the validated corpus")
    setup_per_root = artifact.get("rootSetupPerRootUs")
    if not isinstance(setup_per_root, list) or len(setup_per_root) != 4 or any(not isinstance(value, int) or value <= 0 for value in setup_per_root):
        fail("CandidateHeavy per-root fixture setup observations are incomplete")
    for row in phases:
        if row.get("phase") in {"seed-complete", "cached-reopen", "refresh-adopted"} and int(row.get("firstMeaningfulPaintUs", 0)) <= 0:
            fail("CandidateHeavy meaningful-paint phase lacks a positive recorded value")
    memory = {}
    for stage in ("processBefore", "processPeak", "processAfter"):
        sample = artifact.get(stage)
        if not isinstance(sample, dict) or any(not isinstance(sample.get(field), int) or sample[field] <= 0 for field in ("privateBytes", "workingSetBytes")):
            fail(f"CandidateHeavy {stage} process-memory observation is missing")
        memory[stage] = sample
    deltas = {}
    for field in ("privateBytes", "workingSetBytes"):
        before, peak, after = (memory[stage][field] for stage in ("processBefore", "processPeak", "processAfter"))
        if peak < before or peak < after:
            fail(f"CandidateHeavy {field} peak is below a before/after sample")
        deltas[field] = {"peakMinusBeforeBytes": peak - before, "afterMinusBeforeBytes": after - before,
                         "peakMinusAfterBytes": peak - after}
    first_paint = [{"root": row["root"], "phase": row["phase"], "firstMeaningfulPaintUs": row["firstMeaningfulPaintUs"]}
                   for row in phases if row.get("phase") in {"seed-complete", "cached-reopen", "refresh-adopted"}]
    report = {
        "schema": "viewer-space.four-root-memory-candidate-only-analysis.v1",
        "status": "candidate-run-qualified", "campaignId": manifest["campaign_id"],
        "runId": record["run_id"], "receiptId": record["receipt_id"],
        "workspaceSnapshotId": provenance["workspaceSnapshotId"],
        "fixtureDigests": digests, "fixture": {"roots": 4, "directoriesPerRoot": 501, "filesPerRoot": 100000,
            "topK": 512, "uncachedScanWorkers": 8, "cacheTtlSeconds": 600, "cacheMaxEntries": 4,
            "rootTotalBytes": artifact["rootTotalBytes"]},
        "cacheHitCount": artifact["cacheHitCount"], "phaseCount": len(phases), "phases": phases,
        "processMemoryBytes": memory, "processMemoryDeltas": deltas,
        "rootSetupUs": artifact.get("rootSetupUs"), "rootSetupPerRootUs": setup_per_root,
        "lifecycleUs": artifact.get("lifecycleUs"), "firstMeaningfulPaintPhaseRows": first_paint,
        "memoryReductionTarget": {"targetPercent": 70, "status": "UNMEASURED",
            "reason": "Candidate-only execution has no baseline denominator; no reduction ratio is computed."},
        "comparison": {"baselineIncluded": False, "pairedContrasts": 0, "candidateBaselineRatios": None},
        "interpretation": "This qualifies only the candidate's four-root lifecycle guard and records its absolute process samples. It is not a baseline comparison, memory-gain result, target waiver, or performance conclusion."
    }
    output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def write_retained_reference_report(manifest, root, output):
    validate_reference_manifest_identity(manifest, root)
    measurements = manifest.get("measurements", [])
    if len(measurements) != 12:
        fail(f"expected 12 measurement runs (six mirrored pairs), got {len(measurements)}")
    records = {(r.get("block"), r.get("position")): r for r in manifest.get("runs", []) if r.get("role") == "measurement"}
    if len(records) != 12:
        fail("measurement positions are duplicate or missing")
    warmups = [r for r in manifest.get("runs", []) if r.get("role") == "warmup"]
    if len(warmups) != 2 or [(r.get("side"), r.get("status")) for r in warmups] != [("baseline", "PASSED"), ("candidate", "PASSED")]:
        fail("expected exactly two passed warmups in baseline-then-candidate order")
    expected_sequence = [(1, 1, "baseline"), (1, 2, "candidate"), (1, 3, "candidate"), (1, 4, "baseline"),
                         (2, 1, "baseline"), (2, 2, "candidate"), (2, 3, "candidate"), (2, 4, "baseline"),
                         (3, 1, "baseline"), (3, 2, "candidate"), (3, 3, "candidate"), (3, 4, "baseline")]
    actual_sequence = [(r.get("block"), r.get("position"), r.get("side")) for r in manifest.get("runs", []) if r.get("role") == "measurement"]
    if actual_sequence != expected_sequence:
        fail("measurement sequence is not the exact three-block serial ABBA order")
    data, fixture_digests = {}, set()
    for key, record in records.items():
        if key[0] not in (1, 2, 3) or key[1] not in (1, 2, 3, 4):
            fail(f"invalid block/position: {key}")
        provenance, artifact, _ = run_record(manifest, record, root)
        digest_tuple = tuple(artifact.get("rootFixtureDigests", []))
        if len(digest_tuple) != 4 or len(set(digest_tuple)) != 1:
            fail(f"inconsistent four-root BLAKE3 corpus in {record['label']}")
        fixture_digests.add(digest_tuple)
        validate_reference_profile(artifact, record["label"])
        expected_phase_rows = 25
        if artifact.get("scenario") != "viewer_space_memory_heavy_four_root_retained_reference" or artifact.get("cacheHitCount") != 4 or len(artifact.get("phases", [])) != expected_phase_rows:
            fail(f"cache/lifecycle evidence incomplete in {record['label']}")
        phases = artifact.get("phases", [])
        common = {"open-loaded", "seed-closed-object-retained", "reopened-cached", "refresh-inflight", "refresh-done", "refresh-closed-object-retained"}
        for root_number in range(1, 5):
            rows = [row for row in phases if row.get("root") == root_number]
            if len(rows) != 6 or {row.get("phase") for row in rows} != common:
                fail(f"reference lifecycle phases incomplete for {record['label']} root {root_number}")
            inflight = next(row for row in rows if row.get("phase") == "refresh-inflight")
            validate_reference_scan_state(inflight.get("scanState"), f"{record['label']} root {root_number}")
            stable_phases = common - {"refresh-inflight"}
            if any(int(row.get("rootBytes", -1)) != int(artifact["rootTotalBytes"]) for row in rows if row.get("phase") in stable_phases):
                fail(f"reference completed/closed detail total mismatch for {record['label']} root {root_number}")
            # The in-flight rootBytes field is deliberately excluded: baseline may report zero while scanning,
            # and candidate may still show its previous presentation. Compare only completed/closed detail rows.
        release = [row for row in phases if row.get("root") == 0 and row.get("phase") == "final-owner-release"]
        if len(release) != 1:
            fail(f"reference final-owner-release sample missing for {record['label']}")
        data[key] = {"record": record, "artifact": artifact, "provenance": provenance}
    if len(fixture_digests) != 1:
        fail("baseline/candidate runs did not use the same four deterministic root fixtures")
    if len({int(value["artifact"]["rootTotalBytes"]) for value in data.values()}) != 1:
        fail("baseline/candidate root total bytes differ despite matching fixture digests")
    pairs = []
    for block in (1, 2, 3):
        for pair_index, (baseline_position, candidate_position) in enumerate(((1, 2), (4, 3)), 1):
            base = data[(block, baseline_position)]
            cand = data[(block, candidate_position)]
            if base["record"]["side"] != "baseline" or cand["record"]["side"] != "candidate":
                fail(f"ABBA pair ordering invalid at block {block}, pair {pair_index}")
            b_art, c_art = base["artifact"], cand["artifact"]
            if b_art.get("rootFixtureDigests") != c_art.get("rootFixtureDigests") or b_art.get("rootTotalBytes") != c_art.get("rootTotalBytes"):
                fail(f"matched baseline/candidate pair has different fixture identity or bytes at block {block}, pair {pair_index}")
            base_peak_delta = b_art["processPeak"]["privateBytes"] - b_art["processBefore"]["privateBytes"]
            if base_peak_delta <= 0:
                fail(f"baseline peak private-memory delta is not positive for block {block}, pair {pair_index}")
            pairs.append({
                "block": block, "pair": pair_index,
                "baselineRunId": base["record"]["run_id"], "candidateRunId": cand["record"]["run_id"],
                "baselinePeakPrivateBytes": b_art["processPeak"]["privateBytes"],
                "candidatePeakPrivateBytes": c_art["processPeak"]["privateBytes"],
                "baselinePeakWorkingSetBytes": b_art["processPeak"]["workingSetBytes"],
                "candidatePeakWorkingSetBytes": c_art["processPeak"]["workingSetBytes"],
                "baselinePeakPrivateDeltaBytes": base_peak_delta,
                "candidatePeakPrivateDeltaBytes": c_art["processPeak"]["privateBytes"] - c_art["processBefore"]["privateBytes"],
                "baselineRootTotalBytes": b_art["rootTotalBytes"], "candidateRootTotalBytes": c_art["rootTotalBytes"],
                "baselineRootSetupUs": b_art["rootSetupUs"], "candidateRootSetupUs": c_art["rootSetupUs"],
                "baselineLifecycleUs": b_art["lifecycleUs"], "candidateLifecycleUs": c_art["lifecycleUs"],
                "pairedPeakPrivateDeltaRatioCandidateOverBaseline":
                    (c_art["processPeak"]["privateBytes"] - c_art["processBefore"]["privateBytes"]) / base_peak_delta,
            })
    def pair_summary(field):
        values = [int(pair[field]) for pair in pairs]
        return {"n": len(values), "median": statistics.median(values), "p95": nearest_rank(values, .95), "max": max(values)}
    candidate_paints = {"open-loaded": [], "reopened-cached": [], "refresh-done": []}
    for value in data.values():
        if value["record"]["side"] != "candidate":
            continue
        for row in value["artifact"]["phases"]:
            phase = row.get("phase")
            paint = int(row.get("firstMeaningfulPaintUs", 0))
            if phase in candidate_paints and paint > 0:
                candidate_paints[phase].append(paint)
    phase_paint_summary = {}
    for phase, values in candidate_paints.items():
        target = None
        phase_paint_summary[phase] = {
            "n": len(values), "median": statistics.median(values) if values else None,
            "p95": nearest_rank(values, .95), "max": max(values) if values else None,
            "targetUs": target, "targetMetCount": sum(value <= target for value in values) if target is not None else None,
        }
    report = {
        "schema": "viewer-space.four-root-retained-reference-analysis.v1", "status": "qualified",
        "campaignId": manifest["campaign_id"], "fixtureDigests": next(iter(fixture_digests)),
        "fixture": {"roots": 4, "directoriesPerRoot": 501, "filesPerRoot": 100000, "topK": 512,
                    "uncachedScanWorkers": 8, "cacheTtlSeconds": 600, "cacheMaxEntries": 4},
        "warmups": [{"side": r["side"], "runId": r["run_id"]} for r in warmups],
        "pairedContrasts": pairs,
        "candidateVsBaselineRatio": {
            "pairedPeakPrivateDeltaRatioMedian": statistics.median(pair["pairedPeakPrivateDeltaRatioCandidateOverBaseline"] for pair in pairs),
            "ratioOfUnpairedMedians_peakPrivateDelta": (statistics.median(pair["candidatePeakPrivateDeltaBytes"] for pair in pairs) /
                                        statistics.median(pair["baselinePeakPrivateDeltaBytes"] for pair in pairs)
                                        if statistics.median(pair["baselinePeakPrivateDeltaBytes"] for pair in pairs) else None),
            "ratioOfUnpairedMedians_peakPrivateTotal": (statistics.median(pair["candidatePeakPrivateBytes"] for pair in pairs) /
                                       statistics.median(pair["baselinePeakPrivateBytes"] for pair in pairs)
                                       if statistics.median(pair["baselinePeakPrivateBytes"] for pair in pairs) else None),
        },
        "pairedSampleSummaries": {
            field: pair_summary(field) for field in ("baselinePeakPrivateDeltaBytes", "candidatePeakPrivateDeltaBytes",
                                                        "baselinePeakWorkingSetBytes", "candidatePeakWorkingSetBytes",
                                                        "baselineLifecycleUs", "candidateLifecycleUs")
        },
        "candidateFirstMeaningfulPaintUsByPhase": phase_paint_summary,
        "firstMeaningfulPaintSemantics": "Candidate firstMeaningfulPaintUs is measured from each IViewer::Open entry to the first successful present of a published nonempty map. On later refresh rows it remains that Open-scoped value and is not an F5-specific latency. Baseline zero/unavailable values are not imputed or compared.",
        "cohort": "same-IViewer controlled reference; not historical fresh-instance S0",
        "memoryReductionTarget": {"targetPercent": 70, "status": "MEASURED", "pairedMedianRatio": statistics.median(pair["pairedPeakPrivateDeltaRatioCandidateOverBaseline"] for pair in pairs), "met": statistics.median(pair["pairedPeakPrivateDeltaRatioCandidateOverBaseline"] for pair in pairs) <= 0.70, "denominator": "controlled-reference paired peak private-byte delta; not historical S0"},
        "interpretation": "Six mirrored ABBA contrasts of the controlled-reference lifecycle. Fixture setup is separate. Closed-HWND rows retain the caller COM owner; only the final owner-release row records caller release. This comparison does not qualify historical fresh-instance S0.",
    }
    output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        fail("usage: analyzer CAMPAIGN-MANIFEST OUTPUT-JSON")
    manifest_path, output = pathlib.Path(sys.argv[1]).resolve(), pathlib.Path(sys.argv[2]).resolve()
    root = manifest_path.parent.resolve()
    manifest = read_json(manifest_path)
    if manifest.get("schema") != "viewer-space.memory-qualification.v1":
        fail("unsupported campaign manifest")
    if any(record.get("status") != "PASSED" for record in manifest.get("runs", [])):
        fail("campaign contains incomplete/failed runs; refusing qualified comparison")
    if manifest.get("mode") == "RetainedReference":
        write_retained_reference_report(manifest, root, output)
    elif manifest.get("mode") == "FourRootAbba":
        write_heavy_report(manifest, root, output)
    elif manifest.get("mode") == "CandidateHeavy":
        write_candidate_heavy_report(manifest, root, output)
    elif manifest.get("mode") == "CachedReopen200":
        write_cached_report(manifest, root, output)
    else:
        fail(f"unknown mode {manifest.get('mode')}")
    print(output)


if __name__ == "__main__":
    main()

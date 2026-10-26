#!/usr/bin/env python3
"""Preserve one ViewerSpace memory qualification run with strict evidence bounds."""
import collections
import hashlib
import json
import math
import pathlib
import sys

LIMIT = 2 * 1024 * 1024
RUN_LIMIT = 5 * 1024 * 1024
OMIT_METRICS = {
    "viewer.space.queue.posted_count", "viewer.space.queue.coalesced_count", "viewer.space.queue.pending_bytes",
    "viewer.space.queue.lock_hold_us", "viewer.space.scan.active_workers", "viewer.space.scan.frontier_depth",
    "viewer.space.scan.summary_stage_bytes", "viewer.space.scan.summary_chunk_max_bytes",
    "viewer.space.scan.summary_committed_count", "viewer.space.scan.enumerate_us",
    "viewer.space.scan.frontier_lock_wait_us", "viewer.space.model.directory_count",
    "viewer.space.model.file_candidate_count", "viewer.space.model.synthetic_count",
    "viewer.space.queue.pending_peak_bytes", "viewer.space.render.tile_draw_count",
    "viewer.space.render.text_draw_count", "viewer.space.layout.visible_tiles",
    "viewer.space.scan.progress_bytes", "viewer.space.scan.progress_files", "viewer.space.scan.progress_folders",
}


def fail(message):
    raise RuntimeError(message)


def sha_bytes(payload):
    return hashlib.sha256(payload).hexdigest()


def sha_file(path):
    return sha_bytes(path.read_bytes())


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def file_record(path, name):
    return {"path": name, "bytes": path.stat().st_size, "sha256": sha_file(path)}


def copy_limited(source, archive, name, errors):
    if not source.is_file():
        errors.append(f"missing evidence: {source.name}")
        return None
    target = archive / name
    payload = source.read_bytes()
    if len(payload) > LIMIT:
        errors.append(f"source exceeds 2 MiB per-file limit: {source.name} ({len(payload)} bytes)")
        return None
    target.write_bytes(payload)
    return file_record(target, name)


def copy_log_shards(source, archive, stem, errors):
    if not source.is_file():
        errors.append(f"missing runner log: {source.name}")
        return [], {"path": str(source), "bytes": 0, "sha256": None}
    payload = source.read_bytes()
    records = []
    for index, offset in enumerate(range(0, len(payload), LIMIT)):
        part = payload[offset:offset + LIMIT]
        name = f"{stem}.log" if index == 0 else f"{stem}-part{index + 1}.log"
        target = archive / name
        target.write_bytes(part)
        records.append(file_record(target, name))
    if b"".join((archive / item["path"]).read_bytes() for item in records) != payload:
        fail(f"runner log chunks do not reconstruct source bytes: {source}")
    return records, {"path": str(source), "bytes": len(payload), "sha256": sha_bytes(payload), "shards": [item["path"] for item in records]}


def nearest(values, p):
    ordered = sorted(values)
    return ordered[max(0, math.ceil(p * len(ordered)) - 1)] if ordered else None


def retain_metric_rows(metric_payload):
    lines = metric_payload.splitlines()
    rows = []
    for number, line in enumerate(lines, 1):
        if line.strip():
            rows.append((number, line, json.loads(line)))
    chosen = [item for item in rows if item[2].get("metric", "").startswith("viewer.space.")
              and item[2].get("metric") not in OMIT_METRICS]
    sampling = "all claim-bearing ViewerSpace rows; high-frequency diagnostic counters omitted by named family policy"
    encoded = b"\n".join(item[1] for item in chosen) + (b"\n" if chosen else b"")
    if len(encoded) <= LIMIT:
        return chosen, rows, sampling
    groups = collections.defaultdict(list)
    for item in chosen:
        groups[item[2].get("metric")].append(item)
    sampled = {}
    for metric, items in groups.items():
        if len(items) <= 200 or metric in {"viewer.space.scan.total_us", "viewer.space.memory.process_phase"}:
            for item in items:
                sampled[item[0]] = item
            continue
        witnesses = items[:3] + items[-3:] + [item for item in items if item[2].get("hr", 0) != 0]
        for field in ("durationUs", "value0", "value1"):
            ordered = sorted(items, key=lambda item: item[2].get(field, 0))
            for ordinal in (0, math.ceil(.50 * len(ordered)) - 1, math.ceil(.95 * len(ordered)) - 1, len(ordered) - 1):
                witnesses.append(ordered[ordinal])
        for item in witnesses:
            sampled[item[0]] = item
    chosen = [sampled[index] for index in sorted(sampled)]
    sampling = "high-frequency families retain endpoints, HRESULT failures, and min/p50/p95/max source witnesses; source line map and full-source digest enable replay"
    return chosen, rows, sampling


def write_metric_archive(payload, archive, base_bytes, errors):
    if not payload:
        errors.append("raw performance metrics are empty")
        return [], {"preservation": "missing", "sourceBytes": 0, "sourceSha256": sha_bytes(b"")}
    source_hash = sha_bytes(payload)
    source_rows = sum(bool(line.strip()) for line in payload.splitlines())
    full_rows = payload.splitlines(keepends=True)
    available = RUN_LIMIT - base_bytes - 64 * 1024
    if len(payload) <= max(0, available):
        if any(len(line) > LIMIT for line in full_rows):
            fail("raw metric row exceeds the 2 MiB per-file limit")
        chunks, current, current_bytes = [], [], 0
        for line in full_rows:
            if current_bytes + len(line) > LIMIT:
                chunks.append(b"".join(current)); current, current_bytes = [], 0
            current.append(line); current_bytes += len(line)
        if current:
            chunks.append(b"".join(current))
        files = []
        for index, chunk in enumerate(chunks):
            name = "perf-metrics.jsonl" if index == 0 else f"perf-metrics-part{index + 1}.jsonl"
            target = archive / name; target.write_bytes(chunk); files.append(file_record(target, name))
        if b"".join((archive / item["path"]).read_bytes() for item in files) != payload:
            fail("raw metric shards do not reconstruct source bytes exactly")
        return files, {"preservation": "full-byte-exact", "sourceBytes": len(payload), "sourceSha256": source_hash,
                       "sourceRows": source_rows, "retainedRows": source_rows}

    selected, rows, policy = retain_metric_rows(payload)
    retained = b"\n".join(item[1] for item in selected) + (b"\n" if selected else b"")
    if len(retained) > LIMIT:
        fail("curated metric payload still exceeds 2 MiB after existing-family retention")
    name = "perf-metrics-curated.jsonl"
    (archive / name).write_bytes(retained)
    file_rows = [file_record(archive / name, name)]
    retained_lines = [item[0] for item in selected]
    summary = {
        "schema": "viewer-space.raw-metric-replay-map.v1", "preservation": "curated-with-source-replay-map",
        "sourceSha256": source_hash, "sourceBytes": len(payload), "sourceRows": source_rows,
        "retainedSha256": sha_bytes(retained), "retainedBytes": len(retained), "retainedRows": len(selected),
        "retainedSourceLineNumbers": retained_lines, "retentionPolicy": policy,
        "metricCounts": dict(collections.Counter(item[2].get("metric") for item in rows)),
        "metricSummary": {},
    }
    for metric in sorted(summary["metricCounts"]):
        values = [item[2] for item in rows if item[2].get("metric") == metric]
        summary["metricSummary"][metric] = {}
        for field in ("durationUs", "value0", "value1"):
            samples = [row[field] for row in values if isinstance(row.get(field), (int, float))]
            if samples:
                summary["metricSummary"][metric][field] = {"n": len(samples), "min": min(samples),
                    "p50": nearest(samples, .50), "p95": nearest(samples, .95), "max": max(samples)}
    replay_path = archive / "metric-replay-map.json"
    replay_path.write_text(json.dumps(summary, separators=(",", ":")) + "\n", encoding="utf-8")
    if replay_path.stat().st_size > LIMIT:
        fail("metric source-line replay map exceeds 2 MiB; refusing non-replayable archive")
    file_rows.append(file_record(replay_path, replay_path.name))
    return file_rows, summary


def main():
    if len(sys.argv) != 15:
        fail("usage: curator RAW ARCHIVE RUN_ID SIDE RECEIPT_PATH RECEIPT_ID SOURCE_SNAPSHOT_ID CASE MODE STDOUT STDERR BUILD_LOG_INDEX PRESERVED_RAW_METRICS SOURCE_INDEX")
    raw, archive = pathlib.Path(sys.argv[1]).resolve(), pathlib.Path(sys.argv[2]).resolve()
    run_id, side, receipt_path, expected_receipt, expected_source, case_name, mode, stdout_path, stderr_path, build_log_index_path, preserved_metrics_path, source_index_path = sys.argv[3:]
    receipt_path = pathlib.Path(receipt_path).resolve()
    stdout_path, stderr_path, build_log_index_path, preserved_metrics_path, source_index_path = map(pathlib.Path, (stdout_path, stderr_path, build_log_index_path, preserved_metrics_path, source_index_path))
    preserved_metrics_path = preserved_metrics_path.resolve()
    source_index_path = source_index_path.resolve()
    if archive.exists():
        fail(f"refusing to overwrite campaign archive: {archive}")
    sandbox = pathlib.Path("D:/RedSalamander.Perf").resolve()
    if not raw.is_relative_to(sandbox / "runs") or not archive.is_relative_to(sandbox / "evidence"):
        fail("raw and curated evidence must remain under the fixed marked TestSandbox runs/evidence trees")
    if side not in {"baseline", "candidate"} or mode not in {"FourRootAbba", "CachedReopen200"}:
        fail("invalid side or qualification mode")
    expected_suffix = pathlib.Path("runs") / run_id / "artifacts" / "selftest" / "last_run"
    if raw.name != "last_run" or raw.parts[-5:] != expected_suffix.parts:
        fail(f"raw path is not the canonical run-owned artifact directory: {raw}")

    # Create the archive first, then preserve every available raw item before
    # deciding whether this execution qualifies; failures and partial samples remain inspectable.
    archive.mkdir(parents=True)
    errors, files = [], []
    summary_path = raw / "run-all-tests-results.json"
    snapshot_path = raw / "workspace-final.json"
    command_path = raw / "commands" / "results.json"
    metric_path = raw / "perf" / "perf_metrics.jsonl"
    artifact_path = raw / "perf" / ("viewer_space_memory_heavy_four_root_metrics.json" if mode == "FourRootAbba"
                                    else "viewer_space_cached_reopen_200_first_map.json")
    receipt_record = None
    if receipt_path.is_file():
        try:
            receipt_record = read_json(receipt_path)
            if receipt_record.get("receipt_id") != expected_receipt or receipt_record.get("source_snapshot_id") != expected_source:
                errors.append("build receipt identity/source snapshot did not match requested frozen inputs")
        except (OSError, json.JSONDecodeError) as error:
            errors.append(f"build receipt could not be parsed: {error}")
        copied = copy_limited(receipt_path, archive, "build-receipt.json", errors)
        if copied: files.append(copied)
    else:
        errors.append("build receipt input is missing")
    build_log_index = None
    if build_log_index_path.is_file():
        try:
            build_log_index = read_json(build_log_index_path)
            copied = copy_limited(build_log_index_path, archive, "build-log-index.json", errors)
            if copied: files.append(copied)
        except (OSError, json.JSONDecodeError) as error:
            errors.append(f"build log index could not be parsed: {error}")
    else:
        errors.append("sealed build log index is missing")
    stdout_files, stdout_record = copy_log_shards(stdout_path, archive, "runner-stdout", errors)
    stderr_files, stderr_record = copy_log_shards(stderr_path, archive, "runner-stderr", errors)
    files.extend(stdout_files); files.extend(stderr_files)

    summary = snapshot = command_results = artifact = None
    for path, variable in ((summary_path, "summary"), (snapshot_path, "snapshot"), (command_path, "commands"), (artifact_path, "artifact")):
        if path.is_file():
            try:
                value = read_json(path)
                if variable == "summary": summary = value
                elif variable == "snapshot": snapshot = value
                elif variable == "commands": command_results = value
                else: artifact = value
            except (OSError, json.JSONDecodeError) as error:
                errors.append(f"{path.name} could not be parsed: {error}")
        else:
            errors.append(f"missing evidence: {path.name}")
        copied = copy_limited(path, archive, path.name, errors)
        if copied: files.append(copied)

    case_rows = []
    if isinstance(summary, dict):
        if summary.get("schema") != "red-salamander.run-all-tests.v1" or summary.get("run_id") != run_id or \
                summary.get("platform") != "x64" or summary.get("configuration") != "Release":
            errors.append("canonical run-all-tests summary identity/schema mismatch")
        for suite in summary.get("suites", []):
            case_rows.extend(row for row in suite.get("cases", []) if row.get("name") == case_name)
        if summary.get("case_filter") != case_name:
            errors.append("runner summary case_filter is not the exact requested case")
    if isinstance(command_results, dict):
        command_rows = command_results.get("cases", [])
        if len(command_rows) != 1 or command_rows[0].get("name") != case_name:
            errors.append("focused Commands results do not contain exactly the requested case")
        elif command_rows[0].get("status") != "passed":
            errors.append(f"focused Commands case status is {command_rows[0].get('status')}")
        if command_results.get("repeat_count") != 1:
            errors.append("focused Commands evidence does not prove exactly one in-process repeat")
    if len(case_rows) != 1 or case_rows[0].get("status") != "passed":
        errors.append("canonical runner did not record exactly one passing focused case")
    if isinstance(snapshot, dict) and not snapshot.get("snapshot_id"):
        errors.append("workspace-final.json lacks its independent snapshot_id")

    metric_payload = metric_path.read_bytes() if metric_path.is_file() else b""
    if not metric_path.is_file(): errors.append("raw perf_metrics.jsonl missing")
    raw_source_record = None
    campaign_root = archive.parent.parent.resolve()
    expected_source_dir = (campaign_root / "raw-sources" / archive.name).resolve()
    expected_preserved_metrics = (expected_source_dir / "perf_metrics.jsonl").resolve()
    expected_source_index = (expected_source_dir / "source-index.json").resolve()
    if preserved_metrics_path != expected_preserved_metrics or source_index_path != expected_source_index:
        errors.append("preserved raw metrics/index escape this campaign run's raw-sources directory")
    elif not preserved_metrics_path.is_file() or not source_index_path.is_file():
        errors.append("preserved campaign raw metrics or source index are missing")
    else:
        try:
            raw_source_record = read_json(source_index_path)
            if pathlib.Path(raw_source_record.get("original_path", "")).resolve() != metric_path.resolve() or \
                    pathlib.Path(raw_source_record.get("preserved_path", "")).resolve() != preserved_metrics_path or \
                    raw_source_record.get("bytes") != len(metric_payload) or \
                    raw_source_record.get("sha256") != sha_bytes(metric_payload):
                errors.append("preserved raw source index does not match original raw metric bytes/path")
            preserved_payload = preserved_metrics_path.read_bytes()
            if preserved_payload != metric_payload or len(preserved_payload) != raw_source_record.get("bytes") or \
                    sha_bytes(preserved_payload) != raw_source_record.get("sha256"):
                errors.append("campaign-preserved raw metrics differ from the exact original source")
            copied = copy_limited(source_index_path, archive, "raw-metrics-source-index.json", errors)
            if copied: files.append(copied)
        except (OSError, json.JSONDecodeError, TypeError, ValueError, AttributeError) as error:
            errors.append(f"campaign raw-source preservation could not be verified: {error}")
    existing_bytes = sum(item["bytes"] for item in files)
    metric_files, metric_provenance = write_metric_archive(metric_payload, archive, existing_bytes, errors)
    files.extend(metric_files)
    if isinstance(artifact, dict):
        if mode == "FourRootAbba":
            expected_phases = 20 if side == "baseline" else 24
            if artifact.get("rootCount") != 4 or artifact.get("filesPerRoot") != 100000 or artifact.get("directoriesPerRoot") != 501:
                errors.append("four-root fixture count fields mismatch")
            if artifact.get("effectiveScanThreads") != 8 or artifact.get("configuration") != "topK512_threads8_volume1_cacheTTL600_entries4":
                errors.append("four-root scan/cache profile mismatch")
            digests = artifact.get("rootFixtureDigests", [])
            if len(digests) != 4 or len(set(digests)) != 1:
                errors.append("four-root BLAKE3 fixture digests missing or unequal")
            if len(artifact.get("phases", [])) != expected_phases:
                errors.append(f"four-root phase rows incomplete: got {len(artifact.get('phases', []))}, expected {expected_phases} for {side}")
            if artifact.get("cacheHitCount") != 4:
                errors.append("four-root cache-hit count differs from four")
        else:
            samples = artifact.get("samples", [])
            if artifact.get("expectedSampleCount") != 200 or artifact.get("sampleCount") != 200 or \
                    artifact.get("completedSampleCount") != 200 or artifact.get("status") != "passed":
                errors.append("cached-reopen artifact is incomplete or failed")
            for sample_index, row in enumerate(samples, 1):
                if row.get("status") != "passed" or row.get("firstMeaningfulPaintUs", 0) <= 0 or \
                        row.get("layoutGeneration", 0) <= 0 or row.get("layoutGeneration") != row.get("presentedLayoutGeneration") or \
                        row.get("cacheHitCount") != sample_index or row.get("cacheRestoreBytes", 0) <= 0 or \
                        row.get("cacheResidentBytes", 0) <= 0:
                    errors.append("cached-reopen artifact contains a failed/non-current presentation sample")
                    break
            if artifact.get("cacheHitCount") != 200 or artifact.get("fixtureFiles") != 10000 or artifact.get("fixtureDirectories") != 121:
                errors.append("cached-reopen fixture/cache-hit count mismatch")
    else:
        errors.append("scenario artifact is unavailable for qualification")

    if artifact_path.is_file():
        artifact_sha, artifact_bytes = sha_file(artifact_path), artifact_path.stat().st_size
    else:
        artifact_sha, artifact_bytes = None, 0
    provenance = {
        "schema": "viewer-space.memory-run-evidence.v2", "runId": run_id, "side": side, "mode": mode,
        "case": case_name, "caseStatus": case_rows[0].get("status") if len(case_rows) == 1 else "missing-or-duplicate",
        "buildReceiptId": (receipt_record or {}).get("receipt_id"), "buildSourceSnapshotId": (receipt_record or {}).get("source_snapshot_id"),
        "buildReceiptSha256": sha_file(receipt_path) if receipt_path.is_file() else None,
        "workspaceSnapshotId": (snapshot or {}).get("snapshot_id"),
        "workspaceSnapshotSha256": sha_file(snapshot_path) if snapshot_path.is_file() else None,
        "sealedBuildLog": build_log_index, "runnerStdout": stdout_record, "runnerStderr": stderr_record,
        "rawMetricsOriginalSourcePath": str(metric_path.resolve()) if metric_path.is_file() else None,
        "rawMetricsPreservedSourcePath": str(preserved_metrics_path) if preserved_metrics_path.is_file() else None,
        "rawMetricsSourceIndex": raw_source_record,
        "rawMetricsSha256": sha_bytes(metric_payload), "rawMetricsBytes": len(metric_payload),
        "rawMetricRowCount": sum(bool(line.strip()) for line in metric_payload.splitlines()), "metrics": metric_provenance,
        "artifactSha256": artifact_sha, "artifactBytes": artifact_bytes,
        "fixtureDigests": (artifact or {}).get("rootFixtureDigests"),
        "fixtureFiles": (artifact or {}).get("filesPerRoot", (artifact or {}).get("fixtureFiles")),
        "fixtureDirectories": (artifact or {}).get("directoriesPerRoot", (artifact or {}).get("fixtureDirectories")),
        "files": files, "qualificationStatus": "qualified" if not errors else "unqualified",
        "qualificationErrors": errors, "limits": {"fileBytes": LIMIT, "runBytes": RUN_LIMIT},
        "fullRawMetricsPreserved": metric_provenance.get("preservation") == "full-byte-exact",
    }
    prov_path = archive / "provenance.json"
    prov_path.write_text(json.dumps(provenance, indent=2) + "\n", encoding="utf-8")
    final_files = [path for path in archive.iterdir() if path.is_file()]
    total = sum(path.stat().st_size for path in final_files)
    if any(path.stat().st_size > LIMIT for path in final_files):
        fail("a final run-evidence file, including provenance, exceeds 2 MiB")
    if total > RUN_LIMIT:
        fail(f"run archive exceeds 5 MiB: {total} bytes")
    print(json.dumps({"archive": str(archive), "bytes": total, "qualificationStatus": provenance["qualificationStatus"],
                      "errors": errors, "rawMetricsSha256": provenance["rawMetricsSha256"]}))


if __name__ == "__main__":
    main()

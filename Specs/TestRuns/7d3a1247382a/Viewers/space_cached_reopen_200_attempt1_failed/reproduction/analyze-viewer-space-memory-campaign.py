#!/usr/bin/env python3
"""Fail-closed analysis for the four-root memory ABBA or cached-200 run."""
import hashlib
import json
import pathlib
import statistics
import sys


def fail(message):
    raise RuntimeError(message)


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


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
    artifact_name = "viewer_space_memory_heavy_four_root_metrics.json" if manifest["mode"] == "FourRootAbba" else "viewer_space_cached_reopen_200_first_map.json"
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
    if artifact.get("status") != "passed" or len(samples) != 200 or artifact.get("completedSampleCount") != 200:
        fail("cached-200 artifact lacks 200 completed samples")
    values = [int(row["firstMeaningfulPaintUs"]) for row in samples]
    if any(row.get("status") != "passed" or row.get("layoutGeneration", 0) <= 0 or
           row.get("layoutGeneration") != row.get("presentedLayoutGeneration") or
           row.get("cacheHitCount") != index or row.get("cacheRestoreBytes", 0) <= 0 or
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
        "allSamplesCurrentPresented": True, "interpretation": "candidate-only cached reopen samples; no baseline comparison or cache-phase-cap override",
        "sourceArtifactSha256": provenance["artifactSha256"],
    }
    output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def write_heavy_report(manifest, root, output):
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
        if artifact.get("rootCount") != 4 or artifact.get("filesPerRoot") != 100000 or artifact.get("directoriesPerRoot") != 501 or artifact.get("effectiveScanThreads") != 8:
            fail(f"fixture/profile mismatch in {record['label']}")
        expected_phase_rows = 20 if record.get("side") == "baseline" else 24
        if artifact.get("cacheHitCount") != 4 or len(artifact.get("phases", [])) != expected_phase_rows:
            fail(f"cache/lifecycle evidence incomplete in {record['label']}")
        if int(artifact.get("rootTotalBytes", -1)) <= 0:
            fail(f"root byte total missing in {record['label']}")
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
            base_peak_delta = int(b_art["processPeak"]["privateBytes"]) - int(b_art["processBefore"]["privateBytes"])
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
    candidate_paints = {"seed-complete": [], "cached-reopen": [], "refresh-adopted": []}
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
        target = 100000 if phase == "cached-reopen" else (250000 if phase == "seed-complete" else None)
        phase_paint_summary[phase] = {
            "n": len(values), "median": statistics.median(values) if values else None,
            "p95": nearest_rank(values, .95), "max": max(values) if values else None,
            "targetUs": target, "targetMetCount": sum(value <= target for value in values) if target is not None else None,
        }
    report = {
        "schema": "viewer-space.four-root-memory-abba-analysis.v1", "status": "qualified",
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
        "firstMeaningfulPaintSemantics": "Candidate-only measurements. Seed-complete is uncached open, cached-reopen is a cache-hit open with a separate 100 ms target, refresh-adopted is reported without inventing a target. Baseline has no comparable first-meaningful-paint metric.",
        "interpretation": "Six mirrored ABBA contrasts; fixture setup is reported separately from lifecycle. Peak process samples are application-process observations, not OS-wide or multi-process memory totals. Corpus, actions and final counts match. Candidate retains the displayed model during F5; baseline clears it, so refresh ownership lifetimes differ and this is a complete-version comparison.",
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
    if manifest.get("mode") == "FourRootAbba":
        write_heavy_report(manifest, root, output)
    elif manifest.get("mode") == "CachedReopen200":
        write_cached_report(manifest, root, output)
    else:
        fail(f"unknown mode {manifest.get('mode')}")
    print(output)


if __name__ == "__main__":
    main()

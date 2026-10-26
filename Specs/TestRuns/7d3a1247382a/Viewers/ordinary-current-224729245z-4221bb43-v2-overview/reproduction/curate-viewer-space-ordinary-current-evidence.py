#!/usr/bin/env python3
"""Bounded, byte-exact curation for one ordinary ViewerSpace qualification run."""
import hashlib
import json
import pathlib
import re
import sys

LIMIT = 2 * 1024 * 1024
RUN_LIMIT = 5 * 1024 * 1024
SANDBOX = pathlib.Path("D:/RedSalamander.Perf").resolve()
CASE_FILTER = "viewer_space_perf_"
EXPECTED_CASES = ["viewer_space_perf_small_progressive", "viewer_space_perf_large_optin"]
FIXTURES = {
    "viewer_space_perf_small_progressive": (10_000, 121),
    "viewer_space_perf_large_optin": (100_000, 501),
}


def fail(message):
    raise RuntimeError(message)


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_object)


def digest(payload):
    return hashlib.sha256(payload).hexdigest()


def sha_file(path):
    return digest(path.read_bytes())


def validate_case_sequence(cases, label):
    if [case.get("name") for case in cases] != EXPECTED_CASES or \
            any(case.get("status") != "passed" for case in cases):
        fail(f"{label} does not contain the exact passing small-then-large case sequence")


def validate_fixture_artifact(artifact, case_name):
    if case_name not in FIXTURES:
        fail(f"unknown ordinary scan fixture case: {case_name}")
    files, directories = FIXTURES[case_name]
    if type(artifact.get("file_candidate_count")) is not int or artifact["file_candidate_count"] != files or \
            type(artifact.get("real_directory_count")) is not int or artifact["real_directory_count"] != directories:
        fail(f"ordinary scan fixture mismatch for {case_name}")


def main():
    if len(sys.argv) != 7:
        fail("usage: curator CAMPAIGN_ROOT RUN_ID LABEL SIDE EXPECTED_RECEIPT_ID EXPECTED_SOURCE_SNAPSHOT_ID")
    campaign = pathlib.Path(sys.argv[1]).resolve()
    run_id, label, side, receipt_id, source_snapshot_id = sys.argv[2:]
    marker = SANDBOX / ".red-salamander-test-root"
    if not marker.is_file() or marker.read_text(encoding="utf-8-sig").strip() != "red-salamander.test-sandbox-root.v1":
        fail("fixed TestSandbox root is not already marked")
    if not campaign.is_relative_to(SANDBOX / "evidence"):
        fail("campaign evidence is outside the fixed marked TestSandbox evidence tree")
    if side not in {"baseline", "candidate"} or not re.fullmatch(r"[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}", run_id):
        fail("side or canonical run ID is invalid")
    if not label or pathlib.Path(label).name != label or any(ch in label for ch in "\\/:"):
        fail("curation label is not a safe single path component")
    manifest_path = campaign / "campaign-manifest.json"
    manifest = read_json(manifest_path)
    if manifest.get("qualification_scope") != "ordinary-current-x64-release-large-and-small-scan-comparison":
        fail("campaign qualification scope mismatch")
    expected = manifest.get("expected_builds", {}).get(side)
    if not isinstance(expected, dict) or expected.get("receipt_id") != receipt_id or \
            expected.get("source_snapshot_id") != source_snapshot_id:
        fail("curator receipt/source snapshot does not match the manifest-pinned side")
    sealed_receipt = campaign / expected["receipt_archive_path"]
    sealed_log = campaign / expected["build_log_archive_path"]
    receipt = read_json(sealed_receipt)
    if receipt.get("schema") != "red-salamander.build-receipt.v1" or receipt.get("configuration") != "Release" or \
            receipt.get("platform") != "x64" or receipt.get("tests_enabled") is not True or \
            type(receipt.get("exit_code")) is not int or receipt.get("exit_code") != 0 or \
            receipt.get("receipt_id") != receipt_id or receipt.get("source_snapshot_id") != source_snapshot_id or \
            sha_file(sealed_receipt) != expected.get("receipt_sha256") or \
            sealed_receipt.stat().st_size != expected.get("receipt_bytes"):
        fail("sealed build receipt identity, length, or SHA mismatch")
    if sha_file(sealed_log) != expected.get("build_log_sha256") or sealed_log.stat().st_size != expected.get("build_log_bytes"):
        fail("sealed exact MSBuild log length or SHA mismatch")

    raw = SANDBOX / "runs" / run_id / "artifacts" / "selftest" / "last_run"
    if not raw.is_dir() or not raw.resolve().is_relative_to(SANDBOX / "runs" / run_id):
        fail("canonical run-owned raw evidence path is missing or escapes TestSandbox")
    target = campaign / "curated" / side / label
    raw_source = campaign / "raw-sources" / run_id
    if target.exists() or raw_source.exists():
        fail("refusing to overwrite immutable curation or raw-source evidence")
    target.mkdir(parents=True)
    raw_source.mkdir(parents=True)
    total_bytes = 0

    def write_bounded(destination, payload):
        nonlocal total_bytes
        if len(payload) > LIMIT:
            fail(f"evidence file exceeds 2 MiB: {destination.name}")
        if total_bytes + len(payload) > RUN_LIMIT:
            fail("combined raw source and curated run evidence exceeds 5 MiB")
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(payload)
        total_bytes += len(payload)
        return {"path": destination.relative_to(campaign).as_posix(), "bytes": len(payload), "sha256": digest(payload)}

    raw_metrics_path = raw / "perf" / "perf_metrics.jsonl"
    if not raw_metrics_path.is_file():
        fail("raw perf_metrics.jsonl is missing before canonical-run cleanup")
    raw_payload = raw_metrics_path.read_bytes()
    raw_hash = digest(raw_payload)
    if len(raw_payload) > RUN_LIMIT:
        fail("exact raw performance source exceeds the 5 MiB per-run evidence bound")
    raw_shards = []
    for ordinal, start in enumerate(range(0, len(raw_payload), LIMIT), 1):
        payload = raw_payload[start:start + LIMIT]
        name = "perf_metrics.jsonl" if ordinal == 1 else f"perf_metrics-part{ordinal}.jsonl"
        raw_shards.append(write_bounded(raw_source / name, payload))
    if b"".join((campaign / item["path"]).read_bytes() for item in raw_shards) != raw_payload:
        fail("preserved raw performance shards do not reconstruct source bytes exactly")
    source_index = {
        "schema": "viewer-space.ordinary-current-raw-source.v1",
        "sourceRunId": run_id,
        "source_path": str(raw_metrics_path.resolve()), "source_bytes": len(raw_payload),
        "source_sha256": raw_hash, "shards": raw_shards,
    }
    source_index_payload = (json.dumps(source_index, indent=2) + "\n").encode("utf-8")
    source_index_record = write_bounded(raw_source / "source-index.json", source_index_payload)

    rows = []
    machines = set()
    for line_number, line in enumerate(raw_payload.splitlines(), 1):
        if line.strip():
            try:
                row = json.loads(line, object_pairs_hook=unique_object)
            except (json.JSONDecodeError, ValueError) as error:
                fail(f"invalid/duplicate-key raw metric JSON at line {line_number}: {error}")
            if not isinstance(row, dict):
                fail(f"raw metric line {line_number} is not a JSON object")
            rows.append(row)
            if row.get("machineHash"):
                machines.add(row["machineHash"])
    if len(machines) != 1:
        fail("expected one machine identity in raw performance metrics")

    names = ["run-all-tests-results.json", "workspace-final.json", "commands/results.json",
             "perf/viewer_space_perf_small_progressive_metrics.json",
             "perf/viewer_space_perf_large_optin_metrics.json"]
    copied = {}
    for name in names:
        source = raw / name
        if not source.is_file():
            fail(f"required raw evidence is missing: {name}")
        copied[name] = write_bounded(target / name, source.read_bytes())

    summary = read_json(target / "run-all-tests-results.json")
    commands = read_json(target / "commands" / "results.json")
    snapshot = read_json(target / "workspace-final.json")
    command_cases = commands.get("cases", [])
    runner_cases = [case for suite in summary.get("suites", []) for case in suite.get("cases", [])
                    if case.get("name") in EXPECTED_CASES]
    if summary.get("run_id") != run_id or summary.get("schema") != "red-salamander.run-all-tests.v1" or \
            summary.get("platform") != "x64" or summary.get("configuration") != "Release" or \
            summary.get("case_filter") != CASE_FILTER or type(summary.get("exit_code")) is not int or \
            summary.get("exit_code") != 0:
        fail("canonical runner summary identity/platform/configuration/case filter mismatch")
    if commands.get("repeat_count") != 1:
        fail("Commands evidence does not prove exactly one in-process repeat")
    validate_case_sequence(command_cases, "Commands evidence")
    validate_case_sequence(runner_cases, "canonical runner case records")
    if not snapshot.get("snapshot_id"):
        fail("workspace-final.json lacks its independent snapshot_id")

    artifacts = {}
    for case_name, (files, directories) in FIXTURES.items():
        artifact_name = "viewer_space_perf_small_progressive_metrics.json" if files == 10_000 else "viewer_space_perf_large_optin_metrics.json"
        artifact_path = target / "perf" / artifact_name
        artifact = read_json(artifact_path)
        validate_fixture_artifact(artifact, case_name)
        artifacts[case_name] = {"path": copied[f"perf/{artifact_name}"]["path"],
                                "bytes": copied[f"perf/{artifact_name}"]["bytes"],
                                "sha256": copied[f"perf/{artifact_name}"]["sha256"]}

    evidence = {
        "schema": "viewer-space.ordinary-current-curated-evidence.v1",
        "sourceRunId": run_id, "machineHash": next(iter(machines)),
        "sourceSha256": raw_hash, "sourceBytes": len(raw_payload), "sourceRows": len(rows),
        "rawSourceIndex": source_index_record, "rawSourceShards": raw_shards,
        "artifacts": artifacts, "files": copied,
    }
    write_bounded(target / "evidence.json", (json.dumps(evidence, indent=2) + "\n").encode("utf-8"))
    provenance = {
        "schema": "viewer-space.ordinary-current-run-provenance.v1",
        "qualification_scope": manifest["qualification_scope"],
        "sourceRunId": run_id, "side": side, "mode": "OrdinaryCurrent",
        "caseFilter": CASE_FILTER, "caseNames": EXPECTED_CASES,
        "buildReceiptId": receipt_id, "buildSourceSnapshotId": source_snapshot_id,
        "buildReceiptSha256": expected["receipt_sha256"],
        "buildLogSourcePath": expected["build_log_source_path"],
        "buildLogBytes": expected["build_log_bytes"], "buildLogSha256": expected["build_log_sha256"],
        "workspaceSnapshotId": snapshot["snapshot_id"],
        "workspaceSnapshotSha256": copied["workspace-final.json"]["sha256"],
        "rawSourceIndexPath": source_index_record["path"],
        "rawMetricsSha256": raw_hash, "rawMetricsBytes": len(raw_payload),
        "curatedFiles": copied, "artifactFiles": artifacts,
        "qualificationStatus": "qualified",
    }
    write_bounded(target / "provenance.json", (json.dumps(provenance, indent=2) + "\n").encode("utf-8"))
    if total_bytes > RUN_LIMIT:
        fail(f"curated run plus exact raw source exceeds 5 MiB: {total_bytes} bytes")
    evidence_file = target / "evidence.json"
    provenance_file = target / "provenance.json"
    print(json.dumps({"path": str(target.resolve()), "bytes": total_bytes,
                      "qualificationStatus": provenance["qualificationStatus"],
                      "rawMetricsSha256": raw_hash,
                      "evidenceBytes": evidence_file.stat().st_size,
                      "evidenceSha256": sha_file(evidence_file),
                      "provenanceBytes": provenance_file.stat().st_size,
                      "provenanceSha256": sha_file(provenance_file)}))


if __name__ == "__main__":
    main()

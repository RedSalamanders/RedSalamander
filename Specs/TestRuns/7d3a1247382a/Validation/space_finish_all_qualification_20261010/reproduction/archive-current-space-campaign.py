#!/usr/bin/env python3
"""Preview or archive one already-curated ViewerSpace qualification campaign.

This ignored helper is intentionally narrow. It reads one exact campaign manifest
under the fixed marked TestSandbox evidence root, copies only manifest-selected
curated run leaves and bounded overview inputs, and never launches a test or analyzer.
Default operation is preview-only; writes require --execute.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import pathlib
import re
import shutil
import stat
import sys
import uuid
from dataclasses import dataclass, field
from typing import Any


PERF_ROOT = pathlib.Path("D:/RedSalamander.Perf")
EVIDENCE_ROOT = PERF_ROOT / "evidence"
MARKER_NAME = ".red-salamander-test-root"
MARKER_VALUE = "red-salamander.test-sandbox-root.v1"
REPOSITORY_RELATIVE_TARGET = pathlib.Path("Specs/TestRuns/7d3a1247382a/Viewers")
MAX_FILE_BYTES = 2 * 1024 * 1024
MAX_ARCHIVE_BYTES = 5 * 1024 * 1024
MAX_ARCHIVE_METADATA_BYTES = 128 * 1024
MAX_ARCHIVE_PAYLOAD_BYTES = MAX_ARCHIVE_BYTES - MAX_ARCHIVE_METADATA_BYTES
MAX_RUN_FILES = 256
MAX_TOTAL_FILES = 512
MAX_TREE_DEPTH = 12
MAX_RAW_SOURCE_HASH_BYTES = 128 * 1024 * 1024
REPARSE_ATTRIBUTE = 0x400
SAFE_LABEL = re.compile(r"^[a-z0-9][a-z0-9-]{0,47}$", re.ASCII)
HEX_256 = re.compile(r"^[0-9a-f]{64}$", re.ASCII)
RUN_ID = re.compile(r"^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$", re.ASCII)
MEMORY_SCHEMA = "viewer-space.memory-qualification.v1"
ORDINARY_SCHEMA = "viewer-space.apprelease-comparison-campaign.v1"
ORDINARY_SCOPE = "ordinary-current-x64-release-large-and-small-scan-comparison"
MEMORY_MODES = {"CachedReopen200", "CandidateHeavy", "RetainedReference"}
FORBIDDEN_RUN_DIRS = {"raw-sources", "rawsource", "fixtures", "fixture", "work"}


class ArchiveError(RuntimeError):
    pass


@dataclass(frozen=True)
class CopyItem:
    source: pathlib.Path
    destination: pathlib.PurePosixPath
    size: int
    sha256: str
    group: str
    run_key: str | None = None


@dataclass
class ArchiveLeaf:
    name: str
    kind: str
    copy_items: list[CopyItem]
    run_record: dict[str, Any] | None = None
    generated_files: list[tuple[str, bytes, str]] = field(default_factory=list)


@dataclass
class CampaignPlan:
    repo_root: pathlib.Path
    campaign_root: pathlib.Path
    manifest_path: pathlib.Path
    manifest_payload: bytes
    manifest_sha256: str
    manifest: dict[str, Any]
    label: str
    target_base: pathlib.Path
    leaves: list[ArchiveLeaf]
    run_records: list[dict[str, Any]]
    scope_label: str


def fail(message: str) -> None:
    raise ArchiveError(message)


def _path_text(path: pathlib.Path) -> str:
    return str(path).replace("\\", "/")


def _norm(path: pathlib.Path | str) -> str:
    return os.path.normcase(os.path.abspath(os.fspath(path)))


def _is_within(path: pathlib.Path, root: pathlib.Path) -> bool:
    try:
        return os.path.commonpath((_norm(path), _norm(root))) == _norm(root)
    except ValueError:
        return False


def _reject_reparse(st: os.stat_result, path: pathlib.Path) -> None:
    if stat.S_ISLNK(st.st_mode) or (getattr(st, "st_file_attributes", 0) & REPARSE_ATTRIBUTE):
        fail(f"reparse point or symbolic link is forbidden: {path}")


def _check_existing_ancestors(path: pathlib.Path, *, allow_missing_leaf: bool = False) -> pathlib.Path:
    """Lstat each lexical component; never resolve through a link/reparse point."""
    absolute = pathlib.Path(os.path.abspath(os.fspath(path)))
    current = pathlib.Path(absolute.anchor)
    parts = absolute.parts[1:]
    for index, part in enumerate(parts):
        current = current / part
        try:
            st = current.lstat()
        except FileNotFoundError:
            if allow_missing_leaf:
                break
            raise ArchiveError(f"required path component is missing: {current}") from None
        _reject_reparse(st, current)
        is_leaf = index == len(parts) - 1
        if not is_leaf and not stat.S_ISDIR(st.st_mode):
            fail(f"path ancestor is not a directory: {current}")
    return absolute


def _check_tree_no_reparse(root: pathlib.Path, *, max_files: int = MAX_RUN_FILES) -> list[pathlib.Path]:
    """Enumerate only a selected curated leaf, with strict depth/count/type bounds."""
    root = _check_existing_ancestors(root)
    if not root.is_dir():
        fail(f"selected curated run is not a directory: {root}")
    files: list[pathlib.Path] = []

    def visit(directory: pathlib.Path, depth: int) -> None:
        if depth > MAX_TREE_DEPTH:
            fail(f"selected curated run exceeds directory-depth bound: {root}")
        with os.scandir(directory) as scan:
            children = sorted(scan, key=lambda entry: entry.name.casefold())
        for entry in children:
            path = pathlib.Path(entry.path)
            st = path.lstat()
            _reject_reparse(st, path)
            if stat.S_ISDIR(st.st_mode):
                if entry.name.casefold() in FORBIDDEN_RUN_DIRS:
                    fail(f"raw/fixture/work trees cannot be copied as curated run leaves: {path}")
                visit(path, depth + 1)
            elif stat.S_ISREG(st.st_mode):
                files.append(path)
                if len(files) > max_files:
                    fail(f"selected curated run exceeds file-count bound: {root}")
            else:
                fail(f"unsupported non-regular object in selected curated run: {path}")

    visit(root, 0)
    return files


def _safe_manifest_file(raw_path: str, evidence_root: pathlib.Path) -> pathlib.Path:
    if not raw_path or "\x00" in raw_path:
        fail("manifest path is empty or contains NUL")
    candidate = pathlib.Path(raw_path)
    if not candidate.is_absolute() or ".." in candidate.parts:
        fail("--manifest must be an absolute path without parent traversal")
    candidate = _check_existing_ancestors(candidate)
    if not candidate.is_file() or not _is_within(candidate, evidence_root):
        fail("--manifest must name a regular file beneath the fixed marked evidence root")
    if candidate.name.casefold() != "campaign-manifest.json" or _norm(candidate.parent.parent) != _norm(evidence_root):
        fail("--manifest must be exactly one evidence campaign's campaign-manifest.json")
    return candidate


def _safe_source_path(raw_path: str | pathlib.Path, allowed_root: pathlib.Path, *, must_exist: bool = True) -> pathlib.Path:
    candidate = pathlib.Path(raw_path)
    if not candidate.is_absolute() or ".." in candidate.parts:
        fail(f"source path must be absolute and traversal-free: {raw_path}")
    candidate = _check_existing_ancestors(candidate, allow_missing_leaf=not must_exist)
    if not _is_within(candidate, allowed_root):
        fail(f"source path escapes its approved root: {candidate}")
    if must_exist and not candidate.exists():
        fail(f"selected source is missing: {candidate}")
    return candidate


def _read_json_bytes(payload: bytes, description: str) -> dict[str, Any]:
    def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                fail(f"duplicate JSON property {key!r} in {description}")
            result[key] = value
        return result

    try:
        value = json.loads(payload.decode("utf-8-sig"), object_pairs_hook=unique_object)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        fail(f"invalid JSON in {description}: {exc}")
    if not isinstance(value, dict):
        fail(f"expected a JSON object in {description}")
    return value


def _read_json_file(path: pathlib.Path, *, maximum: int = MAX_FILE_BYTES) -> tuple[dict[str, Any], bytes, str]:
    path = _check_existing_ancestors(path)
    st = path.stat()
    if not stat.S_ISREG(st.st_mode) or st.st_size > maximum:
        fail(f"JSON payload is not a bounded regular file: {path}")
    payload = path.read_bytes()
    if len(payload) != st.st_size:
        fail(f"file changed while reading: {path}")
    return _read_json_bytes(payload, str(path)), payload, hashlib.sha256(payload).hexdigest()


def _sha_file(path: pathlib.Path, *, maximum: int) -> tuple[int, str]:
    path = _check_existing_ancestors(path)
    st = path.stat()
    if not stat.S_ISREG(st.st_mode) or st.st_size > maximum:
        fail(f"payload is not a bounded regular file: {path}")
    digest = hashlib.sha256()
    total = 0
    with path.open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                break
            total += len(block)
            if total > maximum:
                fail(f"payload exceeded its bound while being read: {path}")
            digest.update(block)
    if total != st.st_size:
        fail(f"file changed while hashing: {path}")
    return total, digest.hexdigest()


def _check_marker_first() -> None:
    root = _check_existing_ancestors(PERF_ROOT)
    if not root.is_dir():
        fail(f"fixed TestSandbox root is not a directory: {root}")
    marker = _check_existing_ancestors(root / MARKER_NAME)
    if not marker.is_file():
        fail(f"fixed TestSandbox marker is not a regular file: {marker}")
    if marker.stat().st_size > 4096:
        fail("fixed TestSandbox marker exceeds its small fixed bound")
    payload = marker.read_text(encoding="utf-8-sig").strip()
    if payload != MARKER_VALUE:
        fail("fixed TestSandbox marker does not match the required ownership token")


def _repo_root_after_marker() -> pathlib.Path:
    script = pathlib.Path(os.path.abspath(__file__))
    _check_existing_ancestors(script)
    root = script.parents[2]
    _check_existing_ancestors(root)
    if not (root / "AGENTS.md").is_file():
        fail(f"candidate repository root was not identified: {root}")
    return root


def _expected_id(value: Any, description: str) -> str:
    if not isinstance(value, str) or not HEX_256.fullmatch(value):
        fail(f"invalid {description} SHA-256 identifier")
    return value


def _check_campaign_schema(manifest: dict[str, Any], campaign_root: pathlib.Path) -> tuple[str, str]:
    schema = manifest.get("schema")
    mode = manifest.get("mode")
    if schema == MEMORY_SCHEMA and mode in MEMORY_MODES:
        mode_name = str(mode)
    elif schema == ORDINARY_SCHEMA and manifest.get("qualification_scope") == ORDINARY_SCOPE:
        mode_name = "OrdinaryCurrent"
    else:
        fail(f"unsupported campaign schema/mode: {schema!r}/{mode!r}")
    campaign_id = manifest.get("campaign_id")
    if not isinstance(campaign_id, str) or not re.fullmatch(r"[0-9]{8}T[0-9]{9}Z-[0-9a-f]{8}", campaign_id):
        fail("campaign_id is missing or noncanonical")
    if not campaign_root.name.endswith("-" + campaign_id):
        fail("campaign directory name does not end in its manifest campaign_id")
    if os.path.normcase(os.path.abspath(str(manifest.get("test_root", "")))) != _norm(PERF_ROOT):
        fail("campaign test_root does not match the fixed marked TestSandbox")
    if manifest.get("platform") != "x64" or manifest.get("configuration") != "Release":
        fail("only the pinned x64 Release campaign profiles are supported")
    if mode_name == "OrdinaryCurrent" and (
        manifest.get("case_filter") != "viewer_space_perf_" or
        manifest.get("expected_cases") != ["viewer_space_perf_small_progressive", "viewer_space_perf_large_optin"] or
        manifest.get("expected_fixture_counts") != {
            "small": {"files": 10_000, "directories": 121},
            "large": {"files": 100_000, "directories": 501},
        }
    ):
        fail("ordinary-current campaign scope/case/fixture profile is not the exact supported qualification")
    runs = manifest.get("runs")
    if not isinstance(runs, list) or any(not isinstance(run, dict) for run in runs):
        fail("campaign runs must be an array of objects before archive-state validation")
    ancillary_pending_states = {"starting", "pending", "running", "in_progress", "in-progress", "curation_pending"}
    terminal_run_states = {"passed", "failed", "skipped"}
    terminal_failed = False
    for ordinal, run in enumerate(runs):
        primary = run.get("status")
        if not isinstance(primary, str) or primary.casefold() not in terminal_run_states:
            fail(f"refusing campaign with unsupported or nonterminal primary run status at index {ordinal}")
        ancillary = [run.get("case_status", run.get("caseStatus")),
                     run.get("curation_status"), run.get("qualification_status")]
        has_ancillary_pending = any(
            isinstance(value, str) and value.casefold() in ancillary_pending_states
            for value in ancillary
        )
        if has_ancillary_pending and not (primary.casefold() == "failed" and bool(run.get("finished_utc"))):
            fail(f"refusing run with pending ancillary state that is not a sealed FAILED record at index {ordinal}")
        if primary.casefold() == "failed" and bool(run.get("finished_utc")):
            terminal_failed = True
    finished = bool(manifest.get("finished_utc"))
    if not finished and not terminal_failed:
        fail(f"refusing mutable/incomplete {mode_name} campaign without finished_utc or a terminal FAILED run")
    if not finished and mode_name == "RetainedReference":
        scope_label = "incomplete-retained-reference-history"
    elif not finished:
        scope_label = "incomplete-terminal-failure-history"
    else:
        scope_label = "campaign-state-preserved"
    if campaign_id == "20261009T224059615Z-9e53a14b":
        pins = manifest.get("expected_receipts", {})
        snapshots = manifest.get("expected_build_source_snapshot_ids", {})
        if mode_name != "CachedReopen200" or \
                pins.get("candidate") != "37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd" or \
                snapshots.get("candidate") != "9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554":
            fail("the selected cached-reopen campaign does not preserve the root-pinned receipt/source snapshot")
    return mode_name, scope_label


def _destination_relative(value: str) -> pathlib.PurePosixPath:
    if not isinstance(value, str) or not value or any(character in value for character in ("\x00", "\\", "\t", "\r", "\n")):
        fail(f"unsafe archive-relative path: {value!r}")
    rel = pathlib.PurePosixPath(value)
    if rel.is_absolute() or any(part in {"", ".", ".."} for part in rel.parts) or ":" in value:
        fail(f"unsafe archive-relative path: {value!r}")
    return rel


def _add_file_item(
    items: list[CopyItem],
    destinations: set[str],
    source: pathlib.Path,
    destination: str,
    group: str,
    *,
    expected_bytes: int | None = None,
    expected_sha256: str | None = None,
    run_key: str | None = None,
) -> CopyItem:
    rel = _destination_relative(destination)
    dest_key = rel.as_posix().casefold()
    if dest_key in destinations:
        fail(f"duplicate archive destination: {destination}")
    destinations.add(dest_key)
    source = _check_existing_ancestors(source)
    st = source.stat()
    if not stat.S_ISREG(st.st_mode):
        fail(f"selected copy source is not a regular file: {source}")
    if st.st_size > MAX_FILE_BYTES:
        fail(f"selected file exceeds the 2 MiB per-file cap: {source}")
    size, sha256 = _sha_file(source, maximum=MAX_FILE_BYTES)
    if expected_bytes is not None and size != expected_bytes:
        fail(f"selected file byte length differs from its manifest pin: {source}")
    if expected_sha256 is not None and sha256 != expected_sha256.lower():
        fail(f"selected file SHA-256 differs from its manifest pin: {source}")
    item = CopyItem(source, rel, size, sha256, group, run_key)
    items.append(item)
    if len(items) > MAX_TOTAL_FILES:
        fail("campaign archive exceeds the bounded total-file count")
    return item


def _add_tree_items(
    items: list[CopyItem],
    destinations: set[str],
    source_root: pathlib.Path,
    destination_root: str,
    run_key: str,
) -> tuple[list[CopyItem], int]:
    files = _check_tree_no_reparse(source_root)
    run_items: list[CopyItem] = []
    total = 0
    for source in files:
        rel = source.relative_to(source_root).as_posix()
        destination = f"{destination_root}/{rel}" if destination_root else rel
        item = _add_file_item(items, destinations, source, destination, "run", run_key=run_key)
        run_items.append(item)
        total += item.size
        if total > MAX_ARCHIVE_PAYLOAD_BYTES:
            fail(f"selected curated run exceeds the 5 MiB archive payload budget after metadata reserve: {source_root}")
    return run_items, total


def _safe_run_source(raw_path: Any, campaign_root: pathlib.Path, repo_root: pathlib.Path, ordinary: bool) -> pathlib.Path | None:
    if raw_path is None or raw_path == "":
        return None
    if not isinstance(raw_path, str):
        fail("manifest run archive path is not a string")
    root = campaign_root / "curated" if ordinary else campaign_root / "runs"
    source = _safe_source_path(raw_path, root)
    if not source.is_dir() or source == root:
        fail(f"manifest run pointer is not a selected curated leaf: {source}")
    expected_parent = root / source.parent.name if ordinary else root
    if ordinary and source.parent.name not in {"baseline", "candidate"}:
        fail(f"ordinary run pointer has an invalid side directory: {source}")
    if _norm(source.parent) != _norm(expected_parent):
        fail(f"manifest run pointer must select one direct curated leaf: {source}")
    return source


def _safe_pinned_campaign_relative(campaign_root: pathlib.Path, raw: Any, expected_prefix: str) -> pathlib.Path:
    if not isinstance(raw, str) or not raw:
        fail("campaign manifest lacks a pinned relative input path")
    posix = pathlib.PurePosixPath(raw.replace("\\", "/"))
    if posix.is_absolute() or any(part in {"", ".", ".."} for part in posix.parts):
        fail(f"unsafe manifest-relative input path: {raw}")
    rel = posix.as_posix()
    if not rel.startswith(expected_prefix):
        fail(f"manifest input path does not use the expected prefix {expected_prefix!r}: {rel}")
    source = _safe_source_path(campaign_root.joinpath(*posix.parts), campaign_root)
    return source


def _validate_memory_run(
    manifest: dict[str, Any], run: dict[str, Any], run_root: pathlib.Path, campaign_root: pathlib.Path
) -> None:
    provenance_path = run_root / "provenance.json"
    if not provenance_path.exists():
        # Failed/partial runs are still archived as recorded; absence cannot upgrade them.
        return
    provenance, _, _ = _read_json_file(provenance_path)
    if provenance.get("schema") != "viewer-space.memory-run-evidence.v2":
        fail(f"unsupported run provenance schema: {provenance_path}")
    for prop, expected in (("runId", run.get("run_id")), ("side", run.get("side")), ("mode", manifest.get("mode"))):
        if expected is not None and provenance.get(prop) != expected:
            fail(f"run provenance {prop} differs from campaign manifest: {run_root}")
    for prop, run_prop in (("buildReceiptId", "receipt_id"), ("buildSourceSnapshotId", "receipt_source_snapshot_id"),
                           ("buildReceiptSha256", "receipt_sha256"), ("workspaceSnapshotId", "workspace_snapshot_id")):
        expected = run.get(run_prop)
        if expected is not None and provenance.get(prop) != expected:
            fail(f"run provenance {prop} differs from campaign run record: {run_root}")
    receipt_path = run_root / "build-receipt.json"
    if receipt_path.exists():
        receipt, receipt_bytes, receipt_sha = _read_json_file(receipt_path)
        if receipt.get("receipt_id") != run.get("receipt_id") or \
                receipt.get("source_snapshot_id") != run.get("receipt_source_snapshot_id") or \
                (run.get("receipt_sha256") and receipt_sha != run.get("receipt_sha256")):
            fail(f"curated run build receipt does not match its manifest pins: {receipt_path}")
    expected_sides = manifest.get("expected_receipts", {})
    expected_sources = manifest.get("expected_build_source_snapshot_ids", {})
    side = run.get("side")
    if expected_sides.get(side) and run.get("receipt_id") != expected_sides[side]:
        fail(f"run receipt is not the campaign-pinned {side} receipt")
    if expected_sources.get(side) and run.get("receipt_source_snapshot_id") != expected_sources[side]:
        fail(f"run source snapshot is not the campaign-pinned {side} snapshot")

    index_path = run_root / "raw-metrics-source-index.json"
    replay_path = run_root / "metric-replay-map.json"
    retained_path = run_root / "perf-metrics-curated.jsonl"
    if index_path.exists() and replay_path.exists() and retained_path.exists():
        index, _, _ = _read_json_file(index_path)
        replay, _, _ = _read_json_file(replay_path)
        retained_size, retained_sha = _sha_file(retained_path, maximum=MAX_FILE_BYTES)
        if replay.get("schema") != "viewer-space.raw-metric-replay-map.v1" or \
                replay.get("preservation") != "curated-with-source-replay-map" or \
                index.get("sha256") != replay.get("sourceSha256") or \
                index.get("bytes") != replay.get("sourceBytes") or \
                retained_sha != replay.get("retainedSha256") or retained_size != replay.get("retainedBytes"):
            fail(f"curated metric replay metadata is internally inconsistent: {run_root}")
        raw_path = _safe_source_path(index.get("preserved_path", ""), campaign_root / "raw-sources")
        raw_size, raw_sha = _sha_file(raw_path, maximum=MAX_RAW_SOURCE_HASH_BYTES)
        if raw_size != index.get("bytes") or raw_sha != index.get("sha256"):
            fail(f"single provenance-selected preserved raw source differs from its index: {raw_path}")
        for prop, expected in (("rawMetricsBytes", raw_size), ("rawMetricsSha256", raw_sha)):
            if provenance.get(prop) != expected:
                fail(f"provenance {prop} differs from preserved raw-source index: {run_root}")


def _validate_ordinary_run(
    run: dict[str, Any], run_root: pathlib.Path, campaign_root: pathlib.Path
) -> tuple[Any, list[dict[str, Any]], dict[str, Any]]:
    evidence_path = run_root / "evidence.json"
    provenance_path = run_root / "provenance.json"
    if not evidence_path.is_file() or not provenance_path.is_file():
        fail(f"ordinary-current curated run lacks evidence/provenance records: {run_root}")
    evidence, evidence_bytes, evidence_sha = _read_json_file(evidence_path)
    provenance, provenance_bytes, provenance_sha = _read_json_file(provenance_path)
    pins = (("curated_evidence_sha256", evidence_sha), ("curated_evidence_bytes", len(evidence_bytes)),
            ("curated_provenance_sha256", provenance_sha), ("curated_provenance_bytes", len(provenance_bytes)))
    for prop, actual in pins:
        expected = run.get(prop)
        if expected is not None and expected != actual:
            fail(f"ordinary-current {prop} differs from campaign manifest: {run_root}")
    if provenance.get("sourceRunId") != run.get("run_id") or evidence.get("sourceRunId") != run.get("run_id"):
        fail(f"ordinary-current curated evidence does not match the selected run id: {run_root}")
    if provenance.get("side") != run.get("side"):
        fail(f"ordinary-current curated evidence side differs from campaign manifest: {run_root}")
    run_id = run.get("run_id")
    index_record = evidence.get("rawSourceIndex")
    shard_records = evidence.get("rawSourceShards")
    if not isinstance(index_record, dict) or not isinstance(shard_records, list) or not shard_records or len(shard_records) > 8:
        fail(f"ordinary-current raw-source index/shards are missing or malformed: {run_root}")
    canonical_prefix = f"raw-sources/{run_id}/"
    index_path = _safe_pinned_campaign_relative(campaign_root, index_record.get("path"), canonical_prefix)
    raw_source_root = campaign_root / "raw-sources" / run_id
    if index_path.name != "source-index.json" or index_path.parent != raw_source_root:
        fail(f"ordinary-current source-index path is not the canonical direct run child: {index_path}")
    index, index_payload, index_sha = _read_json_file(index_path)
    if len(index_payload) != index_record.get("bytes") or index_sha != index_record.get("sha256") or \
            provenance.get("rawSourceIndexPath") != index_record.get("path"):
        fail(f"ordinary-current source-index bytes/SHA/path differ from evidence or provenance: {index_path}")
    if index.get("schema") != "viewer-space.ordinary-current-raw-source.v1" or \
            index.get("sourceRunId") != run_id or index.get("shards") != shard_records:
        fail(f"ordinary-current source-index identity/order differs from curated evidence: {index_path}")
    source_bytes = evidence.get("sourceBytes")
    source_sha = _expected_id(evidence.get("sourceSha256"), "ordinary-current raw source")
    if type(source_bytes) is not int or source_bytes <= 0 or source_bytes > MAX_ARCHIVE_PAYLOAD_BYTES or \
            index.get("source_bytes") != source_bytes or index.get("source_sha256") != source_sha or \
            provenance.get("rawMetricsBytes") != source_bytes or provenance.get("rawMetricsSha256") != source_sha or \
            run.get("raw_metrics_sha256") != source_sha:
        fail(f"ordinary-current exact source byte/hash pins disagree: {run_root}")
    expected_raw_source_path = PERF_ROOT / "runs" / run_id / "artifacts" / "selftest" / "last_run" / "perf" / "perf_metrics.jsonl"
    index_source_path = index.get("source_path")
    if not isinstance(index_source_path, str) or _norm(index_source_path) != _norm(expected_raw_source_path):
        fail(f"ordinary-current source index points at a noncanonical raw journal path: {index_path}")

    raw_fragments: list[dict[str, Any]] = []
    archive_sources: list[dict[str, Any]] = [{
        "source_path": str(index_path), "archive_path": "metrics/source-index.json",
        "bytes": len(index_payload), "sha256": index_sha,
    }]
    source_parts: list[bytes] = []
    source_paths: set[str] = set()
    reconstructed_bytes = 0
    for ordinal, shard in enumerate(shard_records, 1):
        if not isinstance(shard, dict):
            fail(f"ordinary-current raw shard record {ordinal} is malformed")
        shard_path = _safe_pinned_campaign_relative(campaign_root, shard.get("path"), canonical_prefix)
        if shard_path.parent != raw_source_root:
            fail(f"ordinary-current raw shard is not a direct run child: {shard_path}")
        shard_key = _norm(shard_path)
        if shard_key in source_paths or shard_path == index_path:
            fail(f"ordinary-current raw-source index/shard paths are duplicated: {shard_path}")
        source_paths.add(shard_key)
        expected_size = shard.get("bytes")
        expected_sha = _expected_id(shard.get("sha256"), f"ordinary-current raw shard {ordinal}")
        if type(expected_size) is not int or expected_size <= 0 or expected_size > MAX_FILE_BYTES:
            fail(f"ordinary-current raw shard {ordinal} exceeds its per-file bound")
        shard_payload = shard_path.read_bytes()
        if len(shard_payload) != expected_size or hashlib.sha256(shard_payload).hexdigest() != expected_sha:
            fail(f"ordinary-current raw shard bytes/SHA differ from source-index pin: {shard_path}")
        reconstructed_bytes += len(shard_payload)
        if reconstructed_bytes > MAX_ARCHIVE_PAYLOAD_BYTES:
            fail("ordinary-current reconstructed raw journal exceeds its bounded archive payload")
        source_parts.append(shard_payload)
        archive_path = f"metrics/{shard_path.name}.part"
        raw_fragments.append({
            "ordinal": ordinal, "source_path": shard.get("path"), "archive_path": archive_path,
            "bytes": expected_size, "sha256": expected_sha,
        })
        archive_sources.append({"source_path": str(shard_path), "archive_path": archive_path,
                                "bytes": expected_size, "sha256": expected_sha})
    raw_payload = b"".join(source_parts)
    if reconstructed_bytes != source_bytes or hashlib.sha256(raw_payload).hexdigest() != source_sha:
        fail(f"ordinary-current raw fragments do not reconstruct the exact source journal: {run_root}")

    rows: list[dict[str, Any]] = []
    machines: set[str] = set()

    def unique_metric_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                fail(f"duplicate JSON property {key!r} in ordinary-current raw metric journal")
            result[key] = value
        return result

    for line_number, line in enumerate(raw_payload.splitlines(), 1):
        if not line.strip():
            continue
        try:
            metric = json.loads(line, object_pairs_hook=unique_metric_object)
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
            fail(f"invalid ordinary-current JSONL row {line_number}: {exc}")
        if not isinstance(metric, dict):
            fail(f"ordinary-current raw metric row {line_number} is not a JSON object")
        rows.append(metric)
        if metric.get("machineHash"):
            machines.add(metric["machineHash"])
    expected_rows = evidence.get("sourceRows")
    if type(expected_rows) is not int or len(rows) != expected_rows:
        fail(f"ordinary-current raw metric row count differs from evidence: {run_root}")
    if machines != {evidence.get("machineHash")}:
        fail(f"ordinary-current raw metric machine identity differs from evidence: {run_root}")
    replay_map = {
        "schema": "viewer-space.ordinary-current-raw-replay-map.v1",
        "evidence": {"source_path": str(evidence_path), "archive_path": "evidence.json",
                     "bytes": len(evidence_bytes), "sha256": evidence_sha},
        "source_index": {"source_path": index_record["path"], "archive_path": "metrics/source-index.json",
                         "bytes": len(index_payload), "sha256": index_sha},
        "source_path": index_source_path, "source_bytes": source_bytes,
        "source_sha256": source_sha, "source_rows": expected_rows,
        "machine_hash": evidence.get("machineHash"), "fragments": raw_fragments,
    }
    return provenance.get("qualificationStatus"), archive_sources, replay_map


def _add_memory_build_inputs(
    manifest: dict[str, Any], campaign_root: pathlib.Path, items: list[CopyItem], destinations: set[str]
) -> None:
    by_side: dict[str, dict[str, Any]] = {}
    for run in manifest.get("runs", []):
        if not isinstance(run, dict):
            fail("campaign runs must be JSON objects")
        side = run.get("side")
        if side not in {"baseline", "candidate"}:
            continue
        if run.get("receipt_id") and run.get("receipt_sha256"):
            previous = by_side.setdefault(side, {})
            receipt_identity = (run.get("receipt_id"), run.get("receipt_sha256"), run.get("receipt_source_snapshot_id"))
            if previous.get("receipt_identity", receipt_identity) != receipt_identity:
                fail(f"campaign changed its {side} build receipt within the selected runs")
            previous["receipt_identity"] = receipt_identity
            previous.setdefault("runs", []).append(run)
        index_raw = run.get("build_log_index_path")
        if index_raw:
            previous = by_side.setdefault(side, {})
            index_path = _safe_source_path(index_raw, campaign_root / "build-inputs")
            expected_index = campaign_root / "build-inputs" / side / "msbuild-log-index.json"
            if _norm(index_path) != _norm(expected_index):
                fail(f"memory campaign log index path is not canonical for {side}")
            previous["index_path"] = index_path
            previous["run"] = run

    for side, selection in by_side.items():
        side_root = campaign_root / "build-inputs" / side
        receipt_source = side_root / "build-receipt.json"
        receipt_identity = selection.get("receipt_identity")
        if receipt_identity:
            receipt_source = _safe_source_path(receipt_source, campaign_root / "build-inputs")
            receipt, receipt_payload, receipt_sha = _read_json_file(receipt_source)
            receipt_id, receipt_sha_pin, snapshot_id = receipt_identity
            if receipt.get("receipt_id") != receipt_id or receipt.get("source_snapshot_id") != snapshot_id or receipt_sha != receipt_sha_pin:
                fail(f"selected {side} build input receipt differs from run pins")
            _add_file_item(items, destinations, receipt_source, f"build-inputs/{side}/build-receipt.json",
                           "build-input", expected_bytes=len(receipt_payload), expected_sha256=receipt_sha_pin)
        index_path = selection.get("index_path")
        if not index_path:
            continue
        index, index_payload, index_sha = _read_json_file(index_path)
        _add_file_item(items, destinations, index_path, f"build-inputs/{side}/msbuild-log-index.json",
                       "build-input", expected_bytes=len(index_payload), expected_sha256=index_sha)
        run = selection.get("run", {})
        for selected_run in selection.get("runs", [run]):
            if selected_run.get("build_log_sha256") and index.get("source_sha256") != selected_run.get("build_log_sha256"):
                fail(f"{side} build-log index digest differs from run manifest")
            if selected_run.get("build_log_bytes") is not None and index.get("source_bytes") != selected_run.get("build_log_bytes"):
                fail(f"{side} build-log index length differs from run manifest")
            run_archive = selected_run.get("archive_path")
            if isinstance(run_archive, str):
                run_root = _safe_source_path(run_archive, campaign_root / "runs")
                local_index = run_root / "build-log-index.json"
                if local_index.exists():
                    local_value, _, _ = _read_json_file(local_index)
                    if local_value != index:
                        fail(f"curated {side} run log index differs from the overview build-input index")
                provenance_path = run_root / "provenance.json"
                if provenance_path.exists():
                    provenance, _, _ = _read_json_file(provenance_path)
                    sealed = provenance.get("sealedBuildLog")
                    if isinstance(sealed, dict) and (
                        sealed.get("source_sha256") != index.get("source_sha256") or
                        sealed.get("source_bytes") != index.get("source_bytes")
                    ):
                        fail(f"curated {side} run build-log provenance differs from overview index")
        chunks = index.get("chunks")
        if not isinstance(chunks, list) or not chunks:
            fail(f"{side} build-log index has no explicit chunks")
        seen_chunks: set[str] = set()
        for chunk in chunks:
            if not isinstance(chunk, dict):
                fail(f"{side} build-log chunk entry is malformed")
            source = _safe_source_path(chunk.get("path", ""), side_root)
            if source.parent != side_root:
                fail(f"build-log shard must be a direct selected child of {side_root}")
            key = _norm(source)
            if key in seen_chunks:
                fail(f"duplicate build-log shard selected for {side}")
            seen_chunks.add(key)
            _add_file_item(items, destinations, source,
                           f"build-inputs/{side}/{source.name}", "build-input",
                           expected_bytes=chunk.get("bytes"), expected_sha256=chunk.get("sha256"))


def _add_ordinary_overview_inputs(
    manifest: dict[str, Any], campaign_root: pathlib.Path, items: list[CopyItem], destinations: set[str]
) -> None:
    expected_builds = manifest.get("expected_builds")
    if not isinstance(expected_builds, dict) or set(expected_builds) != {"baseline", "candidate"}:
        fail("ordinary-current manifest must pin both baseline and candidate build inputs")
    for side in ("baseline", "candidate"):
        build = expected_builds[side]
        if not isinstance(build, dict):
            fail(f"ordinary-current {side} build pin is malformed")
        for field, expected_path, digest_field, bytes_field in (
            ("receipt_archive_path", f"build-inputs/{side}/build-receipt.json", "receipt_sha256", "receipt_bytes"),
            ("build_log_archive_path", f"build-inputs/{side}/msbuild.log", "build_log_sha256", "build_log_bytes"),
        ):
            relative = build.get(field)
            if relative != expected_path:
                fail(f"ordinary-current {side} {field} is not its exact pinned path")
            source = _safe_pinned_campaign_relative(campaign_root, relative, f"build-inputs/{side}/")
            expected_sha = _expected_id(build.get(digest_field), f"ordinary-current {side} {digest_field}")
            size = build.get(bytes_field)
            if type(size) is not int or size <= 0:
                fail(f"ordinary-current {side} {bytes_field} is invalid")
            _add_file_item(items, destinations, source, expected_path, "build-input",
                           expected_bytes=size, expected_sha256=expected_sha)
        receipt_path = campaign_root / pathlib.PurePosixPath(build["receipt_archive_path"])
        receipt, _, _ = _read_json_file(receipt_path)
        if receipt.get("receipt_id") != build.get("receipt_id") or receipt.get("source_snapshot_id") != build.get("source_snapshot_id") or \
                receipt.get("configuration") != "Release" or receipt.get("platform") != "x64" or receipt.get("tests_enabled") is not True:
            fail(f"ordinary-current {side} build receipt identity/profile does not match its manifest pin")

    runtime = manifest.get("dxui_runtime")
    if not isinstance(runtime, dict) or runtime.get("status") != "matched" or \
            not isinstance(runtime.get("sidecars"), dict) or set(runtime["sidecars"]) != {"baseline", "candidate"}:
        fail("ordinary-current campaign lacks the exact matched baseline/candidate DxUi sidecar pins")
    for side in ("baseline", "candidate"):
        pin = runtime["sidecars"][side]
        expected = f"raw-sources/dxui/{side}.DxUi.json"
        if not isinstance(pin, dict) or pin.get("preserved_path") != expected:
            fail(f"ordinary-current {side} DxUi sidecar path is not canonical")
        source = _safe_pinned_campaign_relative(campaign_root, expected, "raw-sources/dxui/")
        _add_file_item(items, destinations, source, f"build-inputs/dxui/{side}.DxUi.json", "build-input",
                       expected_bytes=pin.get("bytes"), expected_sha256=_expected_id(pin.get("sha256"), f"{side} DxUi sidecar"))


def _add_reproduction_inputs(
    manifest: dict[str, Any], campaign_root: pathlib.Path, items: list[CopyItem], destinations: set[str]
) -> None:
    if manifest.get("schema") == MEMORY_SCHEMA:
        reproduction = manifest.get("reproduction")
        if not isinstance(reproduction, dict):
            fail("memory campaign has no manifest-pinned reproduction helper map")
        entries = []
        for name, item in reproduction.items():
            if not isinstance(item, dict):
                fail("memory reproduction helper pin is malformed")
            entries.append((item.get("path"), item.get("bytes"), item.get("sha256"), name))
    else:
        reproduction = manifest.get("reproduction_inputs")
        if not isinstance(reproduction, dict):
            fail("ordinary-current campaign has no manifest-pinned reproduction helper map")
        entries = []
        for name, item in reproduction.items():
            if not isinstance(item, dict):
                fail("ordinary-current reproduction helper pin is malformed")
            preserved = item.get("path")
            if not isinstance(preserved, str):
                fail("ordinary-current reproduction helper lacks path")
            entries.append((preserved, item.get("bytes"), item.get("sha256"), name))

    seen: set[str] = set()
    for raw_path, size, sha256, name in entries:
        if not isinstance(name, str) or pathlib.Path(name).name != name or name in {"", ".", ".."}:
            fail("reproduction helper name is unsafe")
        if name.casefold() in seen:
            fail(f"duplicate reproduction helper name: {name}")
        seen.add(name.casefold())
        if manifest.get("schema") in {MEMORY_SCHEMA, ORDINARY_SCHEMA}:
            source = _safe_pinned_campaign_relative(campaign_root, raw_path, "reproduction/")
        else:
            source = _safe_source_path(raw_path, campaign_root / "reproduction")
            if source.parent != campaign_root / "reproduction":
                fail(f"ordinary-current reproduction helper must be a direct pinned child: {source}")
        _add_file_item(items, destinations, source, f"reproduction/{name}", "reproduction",
                       expected_bytes=size, expected_sha256=_expected_id(sha256, f"reproduction helper {name}"))


def _add_overview_file(
    items: list[CopyItem], destinations: set[str], source: pathlib.Path, destination: str, group: str
) -> None:
    _add_file_item(items, destinations, source, destination, group)


def _safe_slug(value: Any, description: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", value):
        fail(f"unsafe {description}: {value!r}")
    return value.lower().replace("_", "-")


def build_plan(manifest_arg: str, label: str, repo_root: pathlib.Path) -> CampaignPlan:
    if not SAFE_LABEL.fullmatch(label):
        fail("--label must be 1-48 lowercase ASCII letters/digits/hyphens, beginning alphanumeric")
    manifest_path = _safe_manifest_file(manifest_arg, EVIDENCE_ROOT)
    campaign_root = manifest_path.parent
    if campaign_root.parent != EVIDENCE_ROOT:
        fail("campaign directory is not a direct child of evidence")
    manifest_stat = manifest_path.stat()
    if manifest_stat.st_size > MAX_FILE_BYTES:
        fail("campaign manifest exceeds the 2 MiB per-file cap")
    manifest_payload = manifest_path.read_bytes()
    if len(manifest_payload) != manifest_stat.st_size:
        fail("campaign manifest changed while being read")
    manifest = _read_json_bytes(manifest_payload, str(manifest_path))
    mode, scope_label = _check_campaign_schema(manifest, campaign_root)
    target_base = repo_root / REPOSITORY_RELATIVE_TARGET
    target_base = _check_existing_ancestors(target_base)
    if not target_base.is_dir():
        fail(f"archive target base must already exist: {target_base}")
    overview_items: list[CopyItem] = []
    destinations: set[str] = set()
    _add_overview_file(overview_items, destinations, manifest_path, "campaign-manifest.json", "campaign")
    analysis_path = campaign_root / "analysis.json"
    _check_existing_ancestors(analysis_path, allow_missing_leaf=True)
    if analysis_path.exists():
        _add_overview_file(overview_items, destinations, analysis_path, "analysis.json", "analysis")
    _add_reproduction_inputs(manifest, campaign_root, overview_items, destinations)
    if mode in MEMORY_MODES:
        _add_memory_build_inputs(manifest, campaign_root, overview_items, destinations)
    else:
        _add_ordinary_overview_inputs(manifest, campaign_root, overview_items, destinations)

    runs = manifest.get("runs")
    if not isinstance(runs, list):
        fail("campaign runs must be an array")
    if len(runs) > 24:
        fail("campaign manifest exceeds the supported run-record bound")
    if mode in {"CachedReopen200", "CandidateHeavy"} and len(runs) != 1:
        fail(f"{mode} campaign must preserve exactly one run record")
    if mode == "OrdinaryCurrent" and len(runs) != 14:
        fail("ordinary-current campaign must preserve its two warmups and twelve measurements")
    run_records: list[dict[str, Any]] = []
    seen_archive_names = {f"{label}-overview".casefold()}
    run_leaves: list[ArchiveLeaf] = []
    for ordinal, run in enumerate(runs, 1):
        if not isinstance(run, dict):
            fail(f"campaign run at index {ordinal - 1} is not an object")
        side = run.get("side")
        if side not in {"baseline", "candidate"}:
            fail(f"campaign run has invalid side at index {ordinal - 1}")
        role = _safe_slug(run.get("role", run.get("stage", "run")), "run role")
        block = run.get("block", 0)
        position = run.get("position", ordinal)
        if type(block) is not int or block < 0 or type(position) is not int or position < 0:
            fail(f"campaign run has invalid block/position at index {ordinal - 1}")
        is_warmup = role.casefold() == "warmup"
        run_key = f"{label}-warmup-{side}" if is_warmup else f"{label}-b{block:02d}-p{position:02d}-{side}"
        if run_key.casefold() in seen_archive_names:
            fail(f"duplicate direct-child archive name: {run_key}")
        seen_archive_names.add(run_key.casefold())
        run_items: list[CopyItem] = []
        run_destinations: set[str] = set()
        ordinary = mode == "OrdinaryCurrent"
        raw_source = run.get("curated_path") if ordinary else run.get("archive_path")
        source_root = _safe_run_source(raw_source, campaign_root, repo_root, ordinary)
        copied_files = 0
        copied_bytes = 0
        source_names: list[str] = []
        qualification_status = run.get("qualification_status")
        generated_files: list[tuple[str, bytes, str]] = []
        if source_root is not None:
            _check_tree_no_reparse(source_root)
            if ordinary:
                qualification_status, raw_archive_sources, replay_map = _validate_ordinary_run(run, source_root, campaign_root)
            else:
                _validate_memory_run(manifest, run, source_root, campaign_root)
            run_items, copied_bytes = _add_tree_items(run_items, run_destinations, source_root, "", run_key)
            if ordinary:
                for pin in raw_archive_sources:
                    _add_file_item(run_items, run_destinations, pathlib.Path(pin["source_path"]),
                                   pin["archive_path"], "run", expected_bytes=pin["bytes"],
                                   expected_sha256=pin["sha256"], run_key=run_key)
                    copied_bytes += pin["bytes"]
                replay_map_payload = (json.dumps(replay_map, indent=2, ensure_ascii=False) + "\n").encode("utf-8")
                generated_files.append(("raw-metric-replay-map.json", replay_map_payload,
                                        "curator:ordinary-raw-metric-replay-map"))
            copied_files = len(run_items)
            source_names = [item.destination.as_posix() for item in run_items]
            run_leaves.append(ArchiveLeaf(run_key, "run", run_items, generated_files=generated_files))
        run_records.append({
            "archive_folder": run_key if source_root is not None else None,
            "manifest_run_index": ordinal - 1,
            "archived_file_count": copied_files,
            "archived_payload_bytes": copied_bytes,
            "block": block,
            "label": run.get("label"),
            "mode": run.get("mode", mode),
            "position": position,
            "role": run.get("role", run.get("stage")),
            "run_id": run.get("run_id"),
            "side": side,
            "source_path": str(source_root) if source_root is not None else None,
            "status": run.get("status"),
            "case_status": run.get("case_status", run.get("caseStatus")),
            "curation_status": run.get("curation_status"),
            "qualification_status": qualification_status,
            "curation_errors": run.get("curation_errors", []),
            "source_files": source_names,
            "generated_replay_map": "raw-metric-replay-map.json" if generated_files else None,
        })
        if source_root is not None:
            run_leaves[-1].run_record = run_records[-1]
    overview = ArchiveLeaf(f"{label}-overview", "overview", overview_items)
    leaves = [overview, *run_leaves]
    if sum(len(leaf.copy_items) + len(leaf.generated_files) for leaf in leaves) + 5 * len(leaves) > MAX_TOTAL_FILES:
        fail("campaign archives plus per-archive metadata exceed the bounded total-file count")
    for leaf in leaves:
        target = target_base / leaf.name
        if not _is_within(target, target_base):
            fail(f"archive target escapes the approved Viewers directory: {target}")
        _require_target_absent(target, target_base)
    plan = CampaignPlan(repo_root, campaign_root, manifest_path, manifest_payload,
                        hashlib.sha256(manifest_payload).hexdigest(), manifest, label,
                        target_base, leaves, run_records, scope_label)
    _validate_leaf_budgets(plan)
    return plan


def _copy_exact(item: CopyItem, stage: pathlib.Path) -> dict[str, Any]:
    destination = stage.joinpath(*item.destination.parts)
    if not _is_within(destination, stage):
        fail(f"archive destination escaped its private staging directory: {destination}")
    parent = destination.parent
    current = stage
    for part in pathlib.PurePosixPath(item.destination.parent.as_posix()).parts:
        current = current / part
        if not current.exists():
            current.mkdir()
        _reject_reparse(current.lstat(), current)
        if not current.is_dir():
            fail(f"archive destination ancestor is not a directory: {current}")
    source = _check_existing_ancestors(item.source)
    payload = source.read_bytes()
    if len(payload) != item.size or hashlib.sha256(payload).hexdigest() != item.sha256:
        fail(f"source payload changed after preview planning: {source}")
    with destination.open("xb") as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())
    return {"path": item.destination.as_posix(), "bytes": item.size,
            "sha256": item.sha256, "source": str(source), "group": item.group,
            "run_key": item.run_key}


def _make_readme(plan: CampaignPlan, leaf: ArchiveLeaf) -> bytes:
    manifest = plan.manifest
    mode = manifest.get("mode", "OrdinaryCurrent")
    sibling_names = [candidate.name for candidate in plan.leaves if candidate.name != leaf.name]
    lines = [
        "# ViewerSpace campaign evidence archive leaf",
        "",
        f"Archive folder: `{leaf.name}`",
        f"Archive kind: `{leaf.kind}`",
        f"Campaign ID: `{manifest['campaign_id']}`",
        f"Campaign schema/mode: `{manifest['schema']}` / `{mode}`",
        f"Source campaign: `{plan.campaign_root}`",
        f"Scope state: `{plan.scope_label}`",
        f"Campaign manifest SHA-256: `{plan.manifest_sha256}`",
        "",
        "This folder is one independently bounded archive leaf. Source status, curation, and qualification values are preserved as recorded; archiving does not upgrade a failed or partial run. Repository qualification remains NOT_EVALUATED.",
        "",
        "Selected files are copied byte-for-byte, with a 2 MiB per-file limit. Each archive leaf is limited to 5 MiB including generated metadata, reserving at most 128 KiB for metadata. Raw source trees, fixture/work trees, and unselected workspace content are not copied wholesale.",
        "",
        f"Sibling archive folders: {', '.join(f'`{name}`' for name in sibling_names) if sibling_names else '`none`'}.",
        "`manifest.tsv` and verifier-compatible `manifest.txt` list files in this folder by relative path, byte length, SHA-256, and exact source path.",
        "Independent archive replay can use `.build/verify-space-campaign-archive.py <archive-folder> <record-under-D:/RedSalamander.Perf>`. A case pass is not converted into a performance-target pass; incomplete retained-reference history remains explicitly incomplete.",
    ]
    if leaf.generated_files:
        lines.extend([
            "",
            "This ordinary-current leaf includes only the manifest-pinned raw source index and journal fragments; each fragment is explicitly suffixed `.part` because a byte boundary may split a JSON row.",
            "`raw-metric-replay-map.json` binds the unchanged source-index and evidence shard paths to copied `.part` fragments in source-index order. Reconstruct the whole journal and verify its byte count, SHA-256, row count, and machine identity before passing it to the canonical analyzer.",
        ])
    else:
        lines.extend(["", "A campaign-specific preserved raw source remains at its exact source path for independent replay."])
    analysis_item = next((item for item in plan.leaves[0].copy_items if item.destination.as_posix() == "analysis.json"), None)
    if leaf.kind == "overview" and analysis_item is not None:
        try:
            analysis = _read_json_file(analysis_item.source)[0]
        except ArchiveError:
            analysis = {}
        fields = []
        for key in ("status", "targetMetCount", "sampleCount", "firstMeaningfulPaintUs", "claim"):
            if key in analysis:
                value = json.dumps(analysis[key], ensure_ascii=False, separators=(",", ":"))
                if len(value) > 600:
                    value = value[:597] + "..."
                fields.append(f"- Analysis `{key}` (copied verbatim): `{value}`")
        if fields:
            lines.extend(["## Analysis fields copied verbatim", "", *fields, ""])
        if manifest.get("campaign_id") == "20261009T224059615Z-9e53a14b":
            paint = analysis.get("firstMeaningfulPaintUs", {})
            if isinstance(paint, dict):
                lines.extend([
                    "## Cached-reopen target disposition",
                    "",
                    f"The copied analysis records targetMetCount={paint.get('targetMetCount')} of sampleCount={analysis.get('sampleCount')} at targetUs={paint.get('targetUs')}. The focused functional case passed, but the original 100 ms latency target remains a MISS; archive copying does not change that result.",
                    "",
                ])
    return ("\n".join(lines) + "\n").encode("utf-8")


def _make_archive_index(plan: CampaignPlan, leaf: ArchiveLeaf) -> bytes:
    selected_runs = plan.run_records if leaf.kind == "overview" else ([leaf.run_record] if leaf.run_record else [])
    document = {
        "schema": "viewer-space.campaign-archive-copy.v1",
        "archive_folder": leaf.name,
        "archive_kind": leaf.kind,
        "campaign_id": plan.manifest.get("campaign_id"),
        "campaign_mode": plan.manifest.get("mode", "OrdinaryCurrent"),
        "campaign_schema": plan.manifest.get("schema"),
        "campaign_scope_state": plan.scope_label,
        "campaign_manifest_sha256": plan.manifest_sha256,
        "expected_receipts": plan.manifest.get("expected_receipts", {
            side: plan.manifest.get("expected_builds", {}).get(side, {}).get("receipt_id")
            for side in ("baseline", "candidate")
        }),
        "expected_build_source_snapshot_ids": plan.manifest.get("expected_build_source_snapshot_ids", {
            side: plan.manifest.get("expected_builds", {}).get(side, {}).get("source_snapshot_id")
            for side in ("baseline", "candidate")
        }),
        "repository_verdict": "NOT_EVALUATED",
        "sibling_archive_folders": [candidate.name for candidate in plan.leaves if candidate.name != leaf.name],
        "generated_files": [path for path, _payload, _source in leaf.generated_files],
        "runs": selected_runs,
    }
    return (json.dumps(document, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def _manifest_bytes(rows: list[dict[str, Any]]) -> bytes:
    from io import StringIO

    stream = StringIO(newline="")
    writer = csv.writer(stream, delimiter="\t", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
    writer.writerow(("path", "bytes", "sha256", "source"))
    for row in rows:
        writer.writerow((row["path"], row["bytes"], row["sha256"], row["source"]))
    return stream.getvalue().encode("utf-8")


def _metadata_payloads(plan: CampaignPlan, leaf: ArchiveLeaf) -> tuple[bytes, bytes, bytes, int]:
    reserved = {"readme.md", "archive-index.json", ".gitattributes", "manifest.tsv", "manifest.txt"}
    destinations = {item.destination.as_posix().casefold() for item in leaf.copy_items}
    generated_names: set[str] = set()
    for path, _payload, _source in leaf.generated_files:
        rel = _destination_relative(path)
        key = rel.as_posix().casefold()
        if key in destinations or key in generated_names:
            fail(f"{leaf.name} generated file collides with a copied/generated destination: {path}")
        generated_names.add(key)
    if destinations.intersection(reserved) or generated_names.intersection(reserved):
        fail(f"{leaf.name} copied payload collides with generated archive metadata")
    readme = _make_readme(plan, leaf)
    index_payload = _make_archive_index(plan, leaf)
    rows = [{"path": item.destination.as_posix(), "bytes": item.size,
             "sha256": item.sha256, "source": str(item.source)} for item in leaf.copy_items]
    rows.extend([
        {"path": "README.md", "bytes": len(readme), "sha256": hashlib.sha256(readme).hexdigest(), "source": "curator:archive-readme"},
        {"path": "archive-index.json", "bytes": len(index_payload), "sha256": hashlib.sha256(index_payload).hexdigest(), "source": "curator:archive-index"},
        {"path": ".gitattributes", "bytes": len(b"* -text whitespace=cr-at-eol\n"),
         "sha256": hashlib.sha256(b"* -text whitespace=cr-at-eol\n").hexdigest(),
         "source": "curator:byte-exact-gitattributes"},
    ])
    rows.extend({"path": path, "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest(),
                 "source": source} for path, payload, source in leaf.generated_files)
    manifest_payload = _manifest_bytes(rows)
    gitattributes = b"* -text whitespace=cr-at-eol\n"
    metadata_bytes = len(readme) + len(index_payload) + 2 * len(manifest_payload) + len(gitattributes) + \
        sum(len(payload) for _path, payload, _source in leaf.generated_files)
    if metadata_bytes > MAX_ARCHIVE_METADATA_BYTES:
        fail(f"archive metadata exceeds its 128 KiB reserve: {metadata_bytes} bytes")
    return readme, index_payload, manifest_payload, metadata_bytes


def _require_target_absent(target: pathlib.Path, target_base: pathlib.Path) -> None:
    _check_existing_ancestors(target_base)
    if not _is_within(target, target_base) or target.parent != target_base:
        fail(f"archive target is not a direct child of the approved Viewers directory: {target}")
    if os.path.lexists(target):
        fail(f"archive target already exists; no overwrite is allowed: {target}")


def _leaf_payload_bytes(leaf: ArchiveLeaf) -> int:
    return sum(item.size for item in leaf.copy_items)


def _require_archive_size(name: str, payload_bytes: int, metadata_bytes: int) -> None:
    if metadata_bytes > MAX_ARCHIVE_METADATA_BYTES:
        fail(f"{name} metadata exceeds its 128 KiB reserve: {metadata_bytes} bytes")
    if payload_bytes > MAX_ARCHIVE_PAYLOAD_BYTES:
        fail(f"{name} copied payload exceeds 5 MiB minus the 128 KiB metadata reserve: {payload_bytes} bytes")
    if payload_bytes + metadata_bytes > MAX_ARCHIVE_BYTES:
        fail(f"{name} exceeds the 5 MiB archive cap including metadata")


def _validate_leaf_budgets(plan: CampaignPlan) -> None:
    for leaf in plan.leaves:
        payload_bytes = _leaf_payload_bytes(leaf)
        if len(leaf.copy_items) + len(leaf.generated_files) + 5 > MAX_RUN_FILES:
            fail(f"{leaf.name} exceeds the per-archive file-count bound including metadata")
        _, _, _, metadata_bytes = _metadata_payloads(plan, leaf)
        _require_archive_size(leaf.name, payload_bytes, metadata_bytes)


def _validate_source_snapshot(plan: CampaignPlan, rows: list[dict[str, Any]]) -> None:
    current_manifest = plan.manifest_path.read_bytes()
    if hashlib.sha256(current_manifest).hexdigest() != plan.manifest_sha256 or current_manifest != plan.manifest_payload:
        fail("campaign manifest changed while staging the archive")
    for row in rows:
        if row["source"].startswith("curator:"):
            continue
        source = pathlib.Path(row["source"])
        size, digest = _sha_file(source, maximum=MAX_FILE_BYTES)
        if size != row["bytes"] or digest != row["sha256"]:
            fail(f"source payload changed while staging archive: {source}")


def _remove_stage_contained(stage: pathlib.Path, target_base: pathlib.Path) -> None:
    stage_abs = pathlib.Path(os.path.abspath(stage))
    if not _is_within(stage_abs, target_base) or stage_abs.parent != target_base:
        fail(f"refusing cleanup outside the exact archive target parent: {stage_abs}")
    _check_tree_no_reparse(stage_abs, max_files=MAX_TOTAL_FILES)
    shutil.rmtree(stage_abs)


def _measure_tree_bytes(root: pathlib.Path) -> tuple[int, int]:
    files = _check_tree_no_reparse(root, max_files=MAX_TOTAL_FILES)
    total = 0
    for path in files:
        st = path.lstat()
        if not stat.S_ISREG(st.st_mode):
            fail(f"archive contains a non-regular file during size verification: {path}")
        total += st.st_size
    return total, len(files)


def execute_plan(plan: CampaignPlan) -> dict[str, Any]:
    target_base = plan.target_base
    staged: list[tuple[ArchiveLeaf, pathlib.Path, pathlib.Path]] = []
    published: list[pathlib.Path] = []
    published_sizes: dict[str, int] = {}
    staged_metadata_bytes: dict[str, int] = {}
    result: dict[str, Any] | None = None
    failure: ArchiveError | OSError | None = None
    cleanup_errors: list[str] = []
    try:
        _check_existing_ancestors(target_base)
        for leaf in plan.leaves:
            target = target_base / leaf.name
            _require_target_absent(target, target_base)
            stage = target_base / f".{leaf.name}.staging-{uuid.uuid4().hex}"
            if os.path.lexists(stage):
                fail(f"unique staging path unexpectedly exists: {stage}")
            stage.mkdir()
            staged.append((leaf, stage, target))
            _check_existing_ancestors(stage)

        all_rows: list[dict[str, Any]] = []
        for leaf, stage, _target in staged:
            rows: list[dict[str, Any]] = []
            for item in leaf.copy_items:
                rows.append(_copy_exact(item, stage))
            readme, index_payload, _preview_manifest, _preview_metadata = _metadata_payloads(plan, leaf)
            gitattributes = b"* -text whitespace=cr-at-eol\n"
            generated = [
                ("README.md", readme, "curator:archive-readme"),
                ("archive-index.json", index_payload, "curator:archive-index"),
                (".gitattributes", gitattributes, "curator:byte-exact-gitattributes"),
                *leaf.generated_files,
            ]
            for dest, payload, origin in generated:
                path = stage / dest
                with path.open("xb") as stream:
                    stream.write(payload)
                    stream.flush()
                    os.fsync(stream.fileno())
                rows.append({"path": dest, "bytes": len(payload),
                             "sha256": hashlib.sha256(payload).hexdigest(), "source": origin})
            manifest_payload = _manifest_bytes(rows)
            metadata_bytes = len(readme) + len(index_payload) + 2 * len(manifest_payload) + len(gitattributes) + \
                sum(len(payload) for _dest, payload, _origin in leaf.generated_files)
            staged_metadata_bytes[leaf.name] = metadata_bytes
            payload_bytes = sum(item.size for item in leaf.copy_items)
            _require_archive_size(leaf.name, payload_bytes, metadata_bytes)
            for dest in ("manifest.tsv", "manifest.txt"):
                path = stage / dest
                with path.open("xb") as stream:
                    stream.write(manifest_payload)
                    stream.flush()
                    os.fsync(stream.fileno())
            total_bytes, file_count = _measure_tree_bytes(stage)
            _require_archive_size(leaf.name, total_bytes - metadata_bytes, metadata_bytes)
            if file_count > MAX_RUN_FILES:
                fail(f"{leaf.name} final archive exceeds its per-archive file-count bound")
            all_rows.extend(rows)

        _validate_source_snapshot(plan, all_rows)
        for leaf, stage, _target in staged:
            _check_tree_no_reparse(stage, max_files=MAX_TOTAL_FILES)
            total_bytes, _ = _measure_tree_bytes(stage)
            metadata_bytes = staged_metadata_bytes[leaf.name]
            _require_archive_size(leaf.name, total_bytes - metadata_bytes, metadata_bytes)
            published_sizes[leaf.name] = total_bytes

        for leaf, stage, target in staged:
            _require_target_absent(target, target_base)
            _check_existing_ancestors(stage)
            os.rename(stage, target)
            published.append(target)
        result = {
            "status": "ARCHIVED",
            "archives": [{"target": str(target), "kind": leaf.kind,
                          "bytes": published_sizes[leaf.name]} for leaf, _stage, target in staged],
            "campaign_id": plan.manifest.get("campaign_id"),
            "manifest_sha256": plan.manifest_sha256,
            "run_count": len(plan.run_records),
            "repository_verdict": "NOT_EVALUATED",
        }
    except (ArchiveError, OSError) as exc:
        if published:
            result = {
                "status": "PARTIAL_PUBLISH",
                "published_archives": [str(path) for path in published],
                "error": str(exc),
                "campaign_id": plan.manifest.get("campaign_id"),
                "repository_verdict": "NOT_EVALUATED",
            }
        else:
            failure = exc
    finally:
        for _leaf, stage, _target in staged:
            if os.path.lexists(stage):
                try:
                    _remove_stage_contained(stage, target_base)
                except (ArchiveError, OSError) as exc:
                    cleanup_errors.append(f"{stage}: {exc}")
    if result is not None:
        if cleanup_errors:
            result["cleanup_errors"] = cleanup_errors
        return result
    if failure is not None:
        if cleanup_errors:
            fail(f"archive staging failed ({failure}); contained staging cleanup also failed: {'; '.join(cleanup_errors)}")
        if isinstance(failure, ArchiveError):
            raise failure
        fail(f"archive staging or publication failed: {failure}")
    if cleanup_errors:
        fail(f"staging cleanup failed: {'; '.join(cleanup_errors)}")
    fail("archive operation ended without a result")


def preview_plan(plan: CampaignPlan) -> dict[str, Any]:
    archive_previews = []
    for leaf in plan.leaves:
        _, _, _, metadata_bytes = _metadata_payloads(plan, leaf)
        payload_bytes = _leaf_payload_bytes(leaf)
        archive_previews.append({
            "folder": leaf.name,
            "kind": leaf.kind,
            "payload_bytes": payload_bytes,
            "metadata_bytes": metadata_bytes,
            "total_bytes_estimate": payload_bytes + metadata_bytes,
            "selected_files": len(leaf.copy_items),
            "generated_files": [{"path": path, "bytes": len(payload)} for path, payload, _source in leaf.generated_files],
            "run_id": leaf.run_record.get("run_id") if leaf.run_record else None,
            "manifest_run_index": leaf.run_record.get("manifest_run_index") if leaf.run_record else None,
        })
    return {
        "status": "PREVIEW_ONLY",
        "would_write": False,
        "target_base": str(plan.target_base),
        "archives": archive_previews,
        "campaign_id": plan.manifest.get("campaign_id"),
        "campaign_schema": plan.manifest.get("schema"),
        "campaign_mode": plan.manifest.get("mode", "OrdinaryCurrent"),
        "scope_state": plan.scope_label,
        "manifest_sha256": plan.manifest_sha256,
        "selected_files": sum(len(leaf.copy_items) for leaf in plan.leaves),
        "run_count": len(plan.run_records),
        "runs": plan.run_records,
        "repository_verdict": "NOT_EVALUATED",
        "next_step": "The preview writes nothing; --execute stages and publishes only these listed direct-child archives.",
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, help="Exact absolute path to one evidence campaign-manifest.json")
    parser.add_argument("--label", required=True, help="Safe unique destination label under Specs/TestRuns/.../Viewers")
    parser.add_argument("--execute", action="store_true", help="Explicitly stage and publish the listed direct-child archives")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        # Marker/ownership is verified before any campaign, repository, or output access.
        _check_marker_first()
        repo_root = _repo_root_after_marker()
        plan = build_plan(args.manifest, args.label, repo_root)
        result = execute_plan(plan) if args.execute else preview_plan(plan)
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 0
    except ArchiveError as exc:
        print(f"archive-current-space-campaign: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

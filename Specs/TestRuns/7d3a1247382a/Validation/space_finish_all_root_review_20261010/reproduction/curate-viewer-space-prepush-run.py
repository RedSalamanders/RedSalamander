#!/usr/bin/env python3
"""Bounded, byte-exact curation of one ViewerSpace PrePush run."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import sys
import uuid
from pathlib import Path, PurePosixPath

SANDBOX = Path("D:/RedSalamander.Perf").resolve()
RUN_PATTERN = re.compile(r"[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}")
SCOPES = ("reports", "native-captures", "commands-captures", "all")
FILE_LIMIT = 2 * 1024 * 1024
RUN_LIMIT = 5 * 1024 * 1024
RUN_FILES = (
    "selftest/last_run/run-all-tests-results.json",
    "selftest/last_run/run-all-tests-results-v1.json",
    "selftest/last_run/results.json",
    "selftest/last_run/trace.txt",
    "selftest/last_run/workspace-initial.json",
    "selftest/last_run/workspace-final.json",
    "selftest/last_run/workspace-after-selftest.commands.json",
    "selftest/last_run/validation-result.json",
    "selftest/last_run/validation-ledger.json",
    "selftest/last_run/disk-audit.json",
    "selftest/last_run/commands/results.json",
    "selftest/last_run/commands/trace.txt",
)
REPORT_NAME = re.compile(
    r"^(?:run-all-tests(?:-[a-z0-9-]+)?\.(?:json|jsonl|md)|"
    r"(?:case-records?|validation(?:-.*)?|run-plan(?:-.*)?|coverage-plan(?:-.*)|provenance)\.(?:json|jsonl|txt)|"
    r"workspace(?:-.*)?\.json|disk-audit\.json)$", re.IGNORECASE
)
COMMAND_CAPTURE_DIR = PurePosixPath("selftest/last_run/perf/viewer_space_captures")
ARCHIVED_COMMANDS_ROOT = PurePosixPath("archive/7d3a1247382a/Commands")
ARCHIVE_DATE_DIR = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{6}")
ARCHIVED_COMMAND_REPORTS = (
    "commands_results.json",
    "commands_trace.txt",
    "compare_trace.txt",
    "fileops_trace.txt",
    "selftest_run_results.json",
    "selftest_run_trace.txt",
    "selftest.txt",
)

class CurationError(RuntimeError):
    pass

def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def beneath(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False

def is_reparse(path: Path) -> bool:
    try:
        info = path.lstat()
    except FileNotFoundError:
        return False
    attributes = getattr(info, "st_file_attributes", 0)
    reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    junction_check = getattr(path, "is_junction", None)
    return stat.S_ISLNK(info.st_mode) or bool(attributes & reparse_flag) or bool(junction_check and junction_check())

def reject_reparse_path(path: Path, stop: Path) -> None:
    path = Path(os.path.abspath(path))
    stop = Path(os.path.abspath(stop))
    if not beneath(path, stop) and path != stop:
        raise CurationError(f"path is outside the expected workspace: {path}")
    cursor = path
    while True:
        if is_reparse(cursor):
            raise CurationError(f"reparse point is not accepted: {cursor}")
        if cursor == stop:
            break
        cursor = cursor.parent

def require_sandbox() -> Path:
    marker = SANDBOX / ".red-salamander-test-root"
    reject_reparse_path(marker, SANDBOX)
    if not marker.is_file() or marker.read_text(encoding="utf-8-sig").strip() != "red-salamander.test-sandbox-root.v1":
        raise CurationError("marked performance-root ownership marker is invalid")
    return SANDBOX

def exact_artifacts_root(run_id: str, sandbox: Path | None = None) -> Path:
    if not RUN_PATTERN.fullmatch(run_id):
        raise CurationError("run ID is not canonical")
    sandbox = sandbox or require_sandbox()
    runs = sandbox / "runs"
    run_root = runs / run_id
    artifacts = run_root / "artifacts"
    for path, stop in ((runs, sandbox), (run_root, runs), (artifacts, run_root)):
        reject_reparse_path(path, stop)
        if not path.is_dir():
            raise CurationError(f"required run directory is missing: {path}")
    runs_real, run_real, art_real = runs.resolve(strict=True), run_root.resolve(strict=True), artifacts.resolve(strict=True)
    if not beneath(run_real, runs_real) or not beneath(art_real, run_real):
        raise CurationError("run artifacts resolve outside the exact marked run")
    return art_real

def safe_rel(value: str) -> str:
    pure = PurePosixPath(value)
    if (not value or pure.is_absolute() or Path(value).is_absolute() or re.match(r"^[A-Za-z]:", value)
            or any(part in ("", ".", "..") for part in pure.parts)):
        raise CurationError(f"unsafe archive-relative path: {value}")
    return value

def relpath(path: Path, root: Path) -> str:
    try:
        return safe_rel(path.relative_to(root).as_posix())
    except ValueError as exc:
        raise CurationError(f"source path escaped run artifacts: {path}") from exc

def regular_files_no_reparse(root: Path):
    """Walk only the requested subtree; reject junctions before descending."""
    root = Path(os.path.abspath(root))
    stack = [root]
    while stack:
        directory = stack.pop()
        reject_reparse_path(directory, root)
        try:
            with os.scandir(directory) as entries:
                for entry in entries:
                    path = Path(entry.path)
                    if is_reparse(path):
                        raise CurationError(f"reparse point is not accepted: {path}")
                    if entry.is_dir(follow_symlinks=False):
                        stack.append(path)
                    elif entry.is_file(follow_symlinks=False):
                        yield path
        except OSError as exc:
            raise CurationError(f"cannot enumerate bounded evidence directory {directory}: {exc}") from exc

def add(files: set[Path], path: Path, root: Path) -> None:
    if not path.exists():
        return
    reject_reparse_path(path, root)
    real = path.resolve(strict=True)
    if not beneath(real, root) or not real.is_file():
        raise CurationError(f"evidence is not a regular file beneath exact run artifacts: {path}")
    files.add(real)

def commands_capture_dirs(root: Path) -> list[Path]:
    """Resolve only the two governed Commands capture layouts; never walk perf roots."""
    result = []
    current = root.joinpath(*COMMAND_CAPTURE_DIR.parts)
    if current.is_dir():
        reject_reparse_path(current, root)
        result.append(current)

    archived_commands = root.joinpath(*ARCHIVED_COMMANDS_ROOT.parts)
    if archived_commands.is_dir():
        reject_reparse_path(archived_commands, root)
        for entry in archived_commands.iterdir():
            reject_reparse_path(entry, archived_commands)
            if not ARCHIVE_DATE_DIR.fullmatch(entry.name) or not entry.is_dir():
                continue
            capture_dir = entry / "perf" / "viewer_space_captures"
            if capture_dir.is_dir():
                reject_reparse_path(capture_dir, root)
                result.append(capture_dir)
    return sorted(result, key=lambda path: relpath(path, root))

def selected_files(root: Path, scope: str) -> list[Path]:
    if scope not in SCOPES:
        raise CurationError("scope must be reports, native-captures, commands-captures, or all")
    files: set[Path] = set()
    want_reports = scope in ("reports", "all")
    want_native = scope in ("native-captures", "all")
    want_commands = scope in ("commands-captures", "all")

    if want_reports:
        for rel in RUN_FILES:
            add(files, root / PurePosixPath(rel), root)
        for name in ("provenance.json", "provenance.txt"):
            add(files, root / name, root)
        # The completed PrePush run moves canonical suite summaries out of last_run.
        # Preserve only these known report/log files, never the large archived perf tree.
        for name in ARCHIVED_COMMAND_REPORTS:
            add(files, root / name, root)
        last = root / "selftest" / "last_run"
        if last.is_dir():
            reject_reparse_path(last, root)
            for path in last.iterdir():
                if path.is_file() and REPORT_NAME.match(path.name):
                    add(files, path, root)
            for subname in ("commands", "previous_run", "validation"):
                sub = last / subname
                if sub.is_dir():
                    reject_reparse_path(sub, last)
                    for path in sub.iterdir():
                        if path.is_file() and REPORT_NAME.match(path.name):
                            add(files, path, root)

    if want_native:
        viewer_pe = root / "viewer-pe"
        if viewer_pe.is_dir():
            for path in regular_files_no_reparse(viewer_pe):
                name, lower = path.name, path.name.lower()
                if ((name.startswith("viewer-space-") and
                     (lower.endswith(".png") or lower.endswith(".provenance.json") or lower.endswith(".provenance.txt"))) or
                    (name.startswith("TestViewerSpace") and name.endswith(".output.log")) or
                    lower == "viewer-space-focused-diagnostic.json"):
                    add(files, path, root)

    if want_commands:
        for capture_dir in commands_capture_dirs(root):
            for path in regular_files_no_reparse(capture_dir):
                if (path.suffix.lower() == ".png" and path.name.endswith(".png")) or path.name == "provenance.json":
                    add(files, path, root)

    if not files:
        raise CurationError(f"no selected evidence files exist for scope {scope}")
    return sorted(files, key=lambda path: relpath(path, root))

def record(path: Path, root: Path) -> dict[str, object]:
    data = path.read_bytes()
    return {"path": relpath(path, root), "sourcePath": str(path), "bytes": len(data), "sha256": sha(data)}

def parse_unique_json(payload: bytes, path: str) -> object:
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result
    try:
        return json.loads(payload.decode("utf-8-sig"), object_pairs_hook=unique)
    except (UnicodeError, json.JSONDecodeError, ValueError) as exc:
        raise CurationError(f"invalid JSON evidence {path}: {exc}") from exc

def is_commands_capture_dir(parent: PurePosixPath) -> bool:
    if parent == COMMAND_CAPTURE_DIR:
        return True
    parts = parent.parts
    return (len(parts) == len(ARCHIVED_COMMANDS_ROOT.parts) + 3
            and parts[:len(ARCHIVED_COMMANDS_ROOT.parts)] == ARCHIVED_COMMANDS_ROOT.parts
            and ARCHIVE_DATE_DIR.fullmatch(parts[len(ARCHIVED_COMMANDS_ROOT.parts)]) is not None
            and parts[-2:] == ("perf", "viewer_space_captures"))

def is_commands_capture(rel: str) -> bool:
    pure = PurePosixPath(rel)
    return (is_commands_capture_dir(pure.parent)
            and pure.name.startswith(("explorer-", "analyst-", "boundary-"))
            and pure.suffix.lower() == ".png")

def is_native_capture(rel: str) -> bool:
    pure = PurePosixPath(rel)
    return "viewer-pe" in pure.parts and pure.name.startswith("viewer-space-") and pure.suffix.lower() == ".png"

def build_pairs(files: list[Path], root: Path, scope: str) -> dict[str, object]:
    selected = {relpath(path, root): path for path in files}
    command_pngs = sorted(rel for rel in selected if is_commands_capture(rel))
    if command_pngs:
        dirs = {str(PurePosixPath(rel).parent) for rel in command_pngs}
        command_groups = {}
        for directory in dirs:
            pure_dir = PurePosixPath(directory)
            if not is_commands_capture_dir(pure_dir):
                raise CurationError("Commands captures are outside the exact governed capture directories")
            manifest_rel = (pure_dir / "provenance.json").as_posix()
            manifest_path = selected.get(manifest_rel)
            frame_names = {PurePosixPath(rel).name for rel in command_pngs if PurePosixPath(rel).parent == pure_dir}
            if manifest_path is None:
                command_groups[directory] = None
                continue
            parsed = parse_unique_json(manifest_path.read_bytes(), manifest_rel)
            frames = parsed.get("frames") if isinstance(parsed, dict) else None
            if not isinstance(frames, list):
                raise CurationError("Commands shared provenance has no frames array")
            names = [frame.get("file") for frame in frames if isinstance(frame, dict)]
            if len(names) != len(frames) or len(names) != len(set(names)) or set(names) != frame_names:
                raise CurationError("Commands PNG set does not exactly match shared provenance frame names")
            command_groups[directory] = record(manifest_path, root)

    captures = []
    for rel in command_pngs:
        img = selected[rel]
        shared = command_groups[str(PurePosixPath(rel).parent)]
        captures.append({"group": "commands", "png": record(img, root),
                         "provenancePolicy": "shared-directory-manifest",
                         "provenance": [shared] if shared else [],
                         "provenanceStatus": "PRESENT" if shared else "MISSING_IN_SOURCE"})
    for rel, img in sorted(selected.items()):
        if not is_native_capture(rel):
            continue
        sidecars = []
        for suffix in (".provenance.json", ".provenance.txt"):
            companion_rel = rel[:-4] + suffix
            if companion_rel in selected:
                sidecars.append(record(selected[companion_rel], root))
        captures.append({"group": "native", "png": record(img, root),
                         "provenancePolicy": "same-stem-sidecars",
                         "provenance": sidecars,
                         "provenanceStatus": "PRESENT" if sidecars else "MISSING_IN_SOURCE"})
    captures.sort(key=lambda item: (item["group"], item["png"]["path"]))
    return {"schema": "viewer-space.prepush-capture-pairs.v2", "scope": scope,
            "captureCount": len(captures), "captures": captures,
            "policies": {
                "native": "same-stem existing .provenance.json/.provenance.txt; no sidecar is synthesized",
                "commands": "each PNG is bound to the existing shared sibling provenance.json; manifest is preserved byte-exact",
            }}

def enforce_budget(records: list[dict[str, object]], generated: list[bytes]) -> int:
    for row in records:
        if int(row["bytes"]) > FILE_LIMIT:
            raise CurationError(f"source exceeds 2 MiB/file: {row['path']} ({row['bytes']} bytes)")
    names = ("manifest.tsv", "capture-pairs.json", "README.md", ".gitattributes")
    for name, data in zip(names, generated):
        if len(data) > FILE_LIMIT:
            raise CurationError(f"generated file exceeds 2 MiB: {name}")
    total = sum(int(row["bytes"]) for row in records) + sum(map(len, generated))
    if total > RUN_LIMIT:
        raise CurationError(f"selected scope would exceed 5 MiB/archive: {total} bytes")
    return total

def parse_manifest(path: Path) -> list[dict[str, str]]:
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    if not lines or lines[0] != "path\tbytes\tsha256\tabsolute_source":
        raise CurationError("manifest header is invalid")
    rows, seen = [], set()
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != 4:
            raise CurationError("manifest row must contain four tab-separated fields")
        rel, size, digest, source = fields
        safe_rel(rel)
        if rel in seen or not size.isdecimal() or not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise CurationError(f"invalid or duplicate manifest row: {rel}")
        if not Path(source).is_absolute():
            raise CurationError(f"manifest source path is not absolute: {source}")
        rows.append({"path": rel, "bytes": size, "sha256": digest, "absolute_source": source})
        seen.add(rel)
    return rows

def verify_archive(archive: Path, require_sources: bool = True) -> dict[str, object]:
    reject_reparse_path(archive, Path(archive.anchor))
    archive = archive.resolve(strict=True)
    if not archive.is_dir():
        raise CurationError("archive path is not a directory")
    rows = parse_manifest(archive / "manifest.tsv")
    pairs_path = archive / "capture-pairs.json"
    pairs = parse_unique_json(pairs_path.read_bytes(), str(pairs_path))
    if not isinstance(pairs, dict) or not RUN_PATTERN.fullmatch(str(pairs.get("runId", ""))):
        raise CurationError("capture-pairs.json has no canonical run ID")
    scope = pairs.get("scope")
    if scope not in SCOPES:
        raise CurationError("capture-pairs.json has an invalid scope")
    source_root = SANDBOX / "runs" / pairs["runId"] / "artifacts"
    reject_reparse_path(source_root, SANDBOX)
    row_map = {row["path"]: row for row in rows}
    all_sources = True
    for row in rows:
        archived = archive.joinpath(*PurePosixPath(row["path"]).parts)
        reject_reparse_path(archived, archive)
        if not archived.is_file():
            raise CurationError(f"archive payload missing: {archived}")
        data = archived.read_bytes()
        if len(data) > FILE_LIMIT or len(data) != int(row["bytes"]) or sha(data) != row["sha256"]:
            raise CurationError(f"archive payload size/hash mismatch: {archived}")
        source = Path(row["absolute_source"])
        if not beneath(source.resolve(strict=False), source_root.resolve(strict=False)):
            raise CurationError(f"manifest source escaped the exact run: {source}")
        if source.exists():
            reject_reparse_path(source, source_root)
            original = source.read_bytes()
            if len(original) != int(row["bytes"]) or sha(original) != row["sha256"]:
                raise CurationError(f"raw source no longer replays archived bytes: {source}")
        else:
            all_sources = False
            if require_sources:
                raise CurationError(f"raw source unavailable for replay: {source}")

    expected_captures = []
    command_pngs = [row for row in rows if is_commands_capture(row["path"])]
    if command_pngs:
        dirs = {PurePosixPath(row["path"]).parent for row in command_pngs}
        shared_records = {}
        for capture_dir in dirs:
            if not is_commands_capture_dir(capture_dir):
                raise CurationError("archived Commands captures are outside governed directories")
            shared_rel = (capture_dir / "provenance.json").as_posix()
            shared = row_map.get(shared_rel)
            if shared:
                manifest_bytes = archive.joinpath(*PurePosixPath(shared_rel).parts).read_bytes()
                parsed = parse_unique_json(manifest_bytes, shared_rel)
                frames = parsed.get("frames") if isinstance(parsed, dict) else None
                names = [frame.get("file") for frame in frames] if isinstance(frames, list) and all(isinstance(frame, dict) for frame in frames) else []
                expected_names = {PurePosixPath(row["path"]).name for row in command_pngs if PurePosixPath(row["path"]).parent == capture_dir}
                if len(names) != len(set(names)) or set(names) != expected_names:
                    raise CurationError("archived Commands PNG set does not match its shared provenance")
                shared_records[capture_dir] = {"path": shared_rel, "sourcePath": shared["absolute_source"],
                                               "bytes": int(shared["bytes"]), "sha256": shared["sha256"]}
            else:
                shared_records[capture_dir] = None
        for row in command_pngs:
            capture_dir = PurePosixPath(row["path"]).parent
            shared_record = shared_records[capture_dir]
            expected_captures.append({
                "group": "commands",
                "png": {"path": row["path"], "sourcePath": row["absolute_source"],
                        "bytes": int(row["bytes"]), "sha256": row["sha256"]},
                "provenancePolicy": "shared-directory-manifest",
                "provenance": [shared_record] if shared_record else [],
                "provenanceStatus": "PRESENT" if shared_record else "MISSING_IN_SOURCE",
            })
    for row in rows:
        if not is_native_capture(row["path"]):
            continue
        rel, pure = row["path"], PurePosixPath(row["path"])
        sidecars = []
        for suffix in (".provenance.json", ".provenance.txt"):
            companion = rel[:-4] + suffix
            if companion in row_map:
                item = row_map[companion]
                sidecars.append({"path": companion, "sourcePath": item["absolute_source"],
                                 "bytes": int(item["bytes"]), "sha256": item["sha256"]})
        expected_captures.append({
            "group": "native",
            "png": {"path": rel, "sourcePath": row["absolute_source"],
                    "bytes": int(row["bytes"]), "sha256": row["sha256"]},
            "provenancePolicy": "same-stem-sidecars", "provenance": sidecars,
            "provenanceStatus": "PRESENT" if sidecars else "MISSING_IN_SOURCE",
        })
    expected_captures.sort(key=lambda item: (item["group"], item["png"]["path"]))
    if pairs.get("captureCount") != len(expected_captures) or pairs.get("captures") != expected_captures:
        raise CurationError("capture-pairs.json does not replay manifest capture/source pairs")
    entries = list(regular_files_no_reparse(archive))
    archive_bytes = sum(path.stat().st_size for path in entries)
    if archive_bytes > RUN_LIMIT or any(path.stat().st_size > FILE_LIMIT for path in entries):
        raise CurationError("archive violates per-file or per-run size limits")
    return {"archive": str(archive), "sourceRunId": pairs["runId"], "scope": scope,
            "payloadCount": len(rows), "captureCount": len(expected_captures),
            "allRawSourcesAvailable": all_sources, "archiveBytes": archive_bytes}

def curate_run(repo_root: Path, run_id: str, label: str, scope: str = "all",
               sandbox: Path | None = None) -> dict[str, object]:
    if not RUN_PATTERN.fullmatch(run_id) or not re.fullmatch(r"[A-Za-z0-9_-]+", label):
        raise CurationError("run ID or archive label is unsafe")
    sandbox = sandbox or require_sandbox()
    root = exact_artifacts_root(run_id, sandbox)
    files = selected_files(root, scope)
    rows = [record(path, root) for path in files]
    report = root / "selftest" / "last_run" / "run-all-tests-results.json"
    if report.is_file():
        parsed_report = parse_unique_json(report.read_bytes(), str(report))
        if isinstance(parsed_report, dict) and parsed_report.get("run_id") not in (None, run_id):
            raise CurationError("run report ID differs from selected run")

    pairs = build_pairs(files, root, scope)
    pairs["runId"] = run_id
    pair_bytes = (json.dumps(pairs, indent=2, ensure_ascii=False) + "\n").encode()
    manifest_bytes = ("path\tbytes\tsha256\tabsolute_source\n" + "".join(
        f"{row['path']}\t{row['bytes']}\t{row['sha256']}\t{row['sourcePath']}\n" for row in rows)).encode()
    readme = (
        "# ViewerSpace PrePush evidence curation\n\n"
        f"Source run: {run_id}. Scope: {scope}. This archive preserves selected exact run reports/logs and "
        "captures under their original artifacts-relative paths. Existing PNG provenance remains byte-exact: "
        "native frames use same-stem JSON/text sidecars; Commands frames share the sibling provenance.json. "
        "Missing source sidecars are recorded, never synthesized. No test pass or whole-frame acceptance is inferred.\n\n"
        "manifest.tsv binds copied payloads to absolute source paths, byte counts, and SHA-256 values. "
        "capture-pairs.json binds every copied ViewerSpace PNG to its source hash and actual provenance-source hash.\n"
    ).encode()
    attributes = b"* -text whitespace=cr-at-eol\n"
    total = enforce_budget(rows, [manifest_bytes, pair_bytes, readme, attributes])

    repo_root = repo_root.resolve(strict=True)
    test_runs = repo_root / "Specs" / "TestRuns"
    reject_reparse_path(test_runs, repo_root)
    if not test_runs.is_dir():
        raise CurationError("repository Specs/TestRuns directory is missing")
    archive = test_runs / "7d3a1247382a" / "Validation" / label
    parent = archive.parent
    if not beneath(parent.resolve(strict=False), test_runs.resolve(strict=True)):
        raise CurationError("archive parent escapes Specs/TestRuns")
    # Inspect all existing ancestors before mkdir: an existing Validation junction
    # must not redirect even the creation of the staging directory.
    reject_reparse_path(parent, test_runs)
    parent.mkdir(parents=True, exist_ok=True)
    reject_reparse_path(parent, test_runs)
    if archive.exists() or is_reparse(archive):
        raise CurationError(f"refusing to overwrite archive: {archive}")
    staging = parent / f".{label}.partial-{uuid.uuid4().hex}"
    if staging.exists() or is_reparse(staging):
        raise CurationError("unique staging path unexpectedly exists")
    created = False
    try:
        staging.mkdir(parents=False, exist_ok=False)
        created = True
        for source, row in zip(files, rows):
            reject_reparse_path(staging, parent)
            destination = staging.joinpath(*PurePosixPath(row["path"]).parts)
            destination.parent.mkdir(parents=True, exist_ok=True)
            reject_reparse_path(destination.parent, staging)
            data = source.read_bytes()
            if len(data) != int(row["bytes"]) or sha(data) != row["sha256"]:
                raise CurationError(f"raw source changed after preflight: {source}")
            destination.write_bytes(data)
            if destination.read_bytes() != data:
                raise CurationError(f"copied bytes differ from raw source: {source}")
        (staging / "manifest.tsv").write_bytes(manifest_bytes)
        (staging / "capture-pairs.json").write_bytes(pair_bytes)
        (staging / "README.md").write_bytes(readme)
        (staging / ".gitattributes").write_bytes(attributes)
        verified = verify_archive(staging, require_sources=True)
        if verified["archiveBytes"] != total:
            raise CurationError("final archive size differs from preflight")
        if archive.exists() or is_reparse(archive):
            raise CurationError("archive target appeared during curation")
        staging.rename(archive)
        return {**verified, "archive": str(archive), "bytes": total}
    except Exception:
        if created and staging.exists():
            reject_reparse_path(staging, parent)
            if not beneath(staging.resolve(strict=False), parent.resolve(strict=True)):
                raise CurationError("owned staging directory escaped the workspace")
            shutil.rmtree(staging)
        raise

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    curate = sub.add_parser("curate")
    curate.add_argument("--repo-root", required=True, type=Path)
    curate.add_argument("--run-id", required=True)
    curate.add_argument("--label", required=True)
    curate.add_argument("--scope", choices=SCOPES, default="all")
    verify = sub.add_parser("verify")
    verify.add_argument("--archive", required=True, type=Path)
    verify.add_argument("--allow-missing-sources", action="store_true")
    args = parser.parse_args()
    try:
        result = (curate_run(args.repo_root, args.run_id, args.label, args.scope) if args.command == "curate"
                  else verify_archive(args.archive, not args.allow_missing_sources))
    except (CurationError, OSError, ValueError, TypeError) as exc:
        print(f"viewer-space-prepush-curator: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())


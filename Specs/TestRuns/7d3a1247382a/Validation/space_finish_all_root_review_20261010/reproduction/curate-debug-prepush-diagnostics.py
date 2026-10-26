#!/usr/bin/env python3
"""Explicit, bounded curation proposal for the 2026-10-09 Debug PrePush diagnostics.

Default mode prints a read-only inventory/preflight. Pass --execute only after root review.
No recursive discovery is used: inputs are exact report, entry-checkpoint, build archive,
and diagnostic-log paths for one canonical run.
"""
from __future__ import annotations
import argparse
import csv
import hashlib
import json
import os
import shutil
import stat
import uuid
from pathlib import Path, PurePosixPath

RUN_ID = "20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f"
SANDBOX = Path("D:/RedSalamander.Perf").resolve()
REPO = Path(r"C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander").resolve()
EVIDENCE = SANDBOX / "evidence" / "viewer-space-finish-all-20261009"
PRESERVED = EVIDENCE / "prepush-debug-queued-path"
LAST_RUN = PRESERVED / "artifacts" / "selftest" / "last_run"
PRESERVATION_MANIFEST = PRESERVED / "artifacts" / "selftest" / "preservation-manifest.tsv"
CHECKPOINT = SANDBOX / "evidence" / "runs" / RUN_ID
OUTER_RUN_LOG = EVIDENCE / "prepush-debug-queued-path.log"
BUILD_ARCHIVE = REPO / "Specs" / "TestRuns" / "7d3a1247382a" / "Build" / "space_candidate_debug_prepush_reattest_20261009T202815Z"
ARCHIVE_PARENT = REPO / "Specs" / "TestRuns" / "7d3a1247382a" / "Validation"
FILE_LIMIT = 2 * 1024 * 1024
RUN_LIMIT = 5 * 1024 * 1024

# Only completed report, exact suite result/trace, and workspace provenance paths.
LAST_REPORTS = (
    "run-all-tests-results.json", "run-all-tests-results-v1.json",
    "run-all-tests-case-history.jsonl", "run-all-tests-dashboard.md",
    "results.json", "trace.txt",
    "workspace-after-perf.monitor-etw-latency.json",
    "workspace-after-perf.performance-tests2.json",
    "workspace-after-selftest.file-operations.json",
    "workspace-after-standalone.filesystem-curl.json",
    "workspace-after-viewer-pe.interactive.json",
    "workspace-after-viewer-sqlite.interactive.json", "workspace-final.json",
    "commands/trace.txt", "compare/trace.txt",
    "fileops/results.json", "fileops/trace.txt",
)
FAILED_ENTRIES = (
    "selftest.commands", "selftest.compare-directories", "selftest.file-operations",
    "viewer-pe.interactive", "standalone.plugin-contract",
)
ENTRY_COMPONENTS = {
    "selftest.commands": (
        "evidence/e41e498ff41b91d2116571abc825b98047657e633a9f41153d7107615ac4b789.json",
        "results/results.json",
        "terminal-results/d4d4d8ba596789a10baef1d4390186a5c222a1f56a2808b7d4ac32b0c7477182.json",
    ),
    "selftest.compare-directories": (
        "evidence/71eba9c62907437a5c1a537882b2dc5155d906202795c1c95a8e0a38b51e7cdb.json",
        "results/results.json",
        "terminal-results/cad320e3b40bd3dca0d9a7640a64ac3e5461bd999505c072e9873decce810d9d.json",
    ),
    "selftest.file-operations": (
        "evidence/2a2ff6f5a4bd6e09c2b0c6992ca036aa1907d40397a713adab037bda0d22621f.json",
        "results/results.json",
        "terminal-results/9a76df4211567dbf3550e808f39319cdf2db1778d7563407680e806b19577332.json",
    ),
    "viewer-pe.interactive": (
        "evidence/68c10f05a5035e59947bc687800d0d3a8fce0f63c887e38300da54a2c83c001e.json",
        "results/ViewerPETests.Interactive.output.log",
        "terminal-results/66e4cf44a288fe32bd1505e562018e9446d60344b840d9912944a2ab0972417f.json",
    ),
    "standalone.plugin-contract": (
        "evidence/bbfb1474fae73bef8252dbb2dd25762c67ded72a0f755acf85f854c8386df4e6.json",
        "results/PluginContractTests.output.log",
        "terminal-results/a92186103be3368b899cba76e2c223f9114416d35d7187fabb0af91b6eb0042e.json",
    ),
}
DECISIONS = (
    "decisions/167542f4a3c3655c647aeb85b4b0c4aede5933c8f0f06fdcb22fa98b4011876b.json",
    "decisions/836ccb1355554c10b4ae901ad71d9ce3d24edeccfa333e89fe1734b436e01395.json",
)
REPORT_LABEL = "space_prepush_debug_failure_reports_20261009T201507Z"
DIAGNOSTIC_LABEL = "space_prepush_debug_interactive_s3_diagnostics_20261009T201507Z"
PLUGIN_LOG = "entries/standalone.plugin-contract/results/PluginContractTests.output.log"
PE_LOG = "entries/viewer-pe.interactive/results/ViewerPETests.Interactive.output.log"

class CurationError(RuntimeError):
    pass

def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def within(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False

def reparse(path: Path) -> bool:
    try:
        info = path.lstat()
    except FileNotFoundError:
        return False
    flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    junction = getattr(path, "is_junction", None)
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info, "st_file_attributes", 0) & flag) or bool(junction and junction())

def reject_reparse(path: Path, stop: Path) -> None:
    path, stop = Path(os.path.abspath(path)), Path(os.path.abspath(stop))
    if path != stop and not within(path, stop):
        raise CurationError(f"path escaped expected root: {path}")
    cursor = path
    while True:
        if reparse(cursor):
            raise CurationError(f"reparse point is not accepted: {cursor}")
        if cursor == stop:
            break
        cursor = cursor.parent

def require_marker() -> None:
    marker = SANDBOX / ".red-salamander-test-root"
    reject_reparse(marker, SANDBOX)
    if marker.read_text(encoding="utf-8-sig").strip() != "red-salamander.test-sandbox-root.v1":
        raise CurationError("Droot ownership marker is invalid")

def read_preservation_map() -> dict[str, dict[str, str]]:
    result = {}
    reject_reparse(PRESERVATION_MANIFEST, SANDBOX)
    if not within(PRESERVATION_MANIFEST.resolve(strict=True), SANDBOX):
        raise CurationError("preservation manifest escaped the marked Droot")
    with PRESERVATION_MANIFEST.open(encoding="utf-8-sig", newline="") as stream:
        for fields in csv.reader(stream, delimiter="\t"):
            if not fields or fields[0].startswith("#") or fields[0].lower() in ("path", "relative_path"):
                continue
            if len(fields) < 4:
                raise CurationError(f"malformed preserved-source manifest row: {fields}")
            rel, size, digest, original = fields[:4]
            pure = PurePosixPath(rel)
            if not rel or pure.is_absolute() or ".." in pure.parts or rel in result:
                raise CurationError(f"unsafe or duplicate preservation path: {rel}")
            result[rel] = {"bytes": size, "sha256": digest, "original": original}
    return result

def payload(group: str, archive_rel: str, preserved: Path, original: str | None = None) -> dict[str, object]:
    pure = PurePosixPath(archive_rel)
    if pure.is_absolute() or ".." in pure.parts:
        raise CurationError(f"unsafe archive path: {archive_rel}")
    allowed_root = SANDBOX if within(preserved, SANDBOX) else REPO
    reject_reparse(preserved, allowed_root)
    resolved = preserved.resolve(strict=True)
    if not within(resolved, allowed_root):
        raise CurationError(f"preserved source escaped its exact root: {preserved}")
    if not preserved.is_file():
        raise CurationError(f"required preserved source missing: {preserved}")
    data = preserved.read_bytes()
    if len(data) > FILE_LIMIT:
        raise CurationError(f"source exceeds 2 MiB/file: {preserved} ({len(data)} bytes)")
    return {"group": group, "archivePath": archive_rel, "preservedSource": str(preserved),
            "originalSource": original or "", "bytes": len(data), "sha256": sha(data)}

def build_inventory() -> dict[str, list[dict[str, object]]]:
    preserved_map = read_preservation_map()
    reports = []
    for rel in LAST_REPORTS:
        info = preserved_map.get(rel)
        source = LAST_RUN.joinpath(*PurePosixPath(rel).parts)
        if info is None:
            raise CurationError(f"selected last_run file is not bound by preservation manifest: {rel}")
        reject_reparse(source, SANDBOX)
        if not within(source.resolve(strict=True), SANDBOX):
            raise CurationError(f"preserved last_run source escaped the marked Droot: {rel}")
        data = source.read_bytes()
        if len(data) != int(info["bytes"]) or sha(data) != info["sha256"]:
            raise CurationError(f"preserved last_run bytes differ from preservation manifest: {rel}")
        reports.append(payload("reports", f"preserved-last-run/{rel}", source, info["original"]))
    reports.append(payload("reports", "preservation-manifest.tsv", PRESERVATION_MANIFEST))

    # Immutable validation checkpoint: global plan/state and both decision snapshots.
    for rel in ("validation-plan.json", "run-state.json", *DECISIONS):
        source = CHECKPOINT.joinpath(*PurePosixPath(rel).parts)
        reports.append(payload("reports", f"checkpoint/{rel}", source))
    for entry in FAILED_ENTRIES:
        entry_root = CHECKPOINT / "entries" / entry
        reject_reparse(entry_root, SANDBOX)
        for subpath in ENTRY_COMPONENTS[entry]:
            source = entry_root.joinpath(*PurePosixPath(subpath).parts)
            directory = source.parent
            reject_reparse(directory, SANDBOX)
            direct_files = sorted(p.name for p in directory.iterdir() if p.is_file())
            if direct_files != [source.name]:
                raise CurationError(f"checkpoint component directory is missing or ambiguous for {entry}/{subpath}: {direct_files}")
            reject_reparse(source, SANDBOX)
            if not source.is_file():
                raise CurationError(f"required checkpoint result component missing for {entry}/{subpath}")
            reports.append(payload("reports", f"checkpoint/entries/{entry}/{subpath}", source))

    # Bind the run to its already archived exact re-attestation receipt, sidecar and log.
    build_manifest = BUILD_ARCHIVE / "manifest.tsv"
    reject_reparse(build_manifest, REPO)
    if not within(build_manifest.resolve(strict=True), REPO):
        raise CurationError("build archive manifest escaped the repository")
    with build_manifest.open(encoding="utf-8-sig", newline="") as stream:
        build_rows = {row["path"]: row for row in csv.DictReader(stream, delimiter="\t")}
    for rel in ("build-receipt.json", "msbuild.log", "outer.log", "RedSalamander.exe.DxUi.json"):
        row = build_rows.get(rel)
        if row is None:
            raise CurationError(f"exact Debug re-attestation archive lacks {rel}")
        source = BUILD_ARCHIVE / rel
        original = row["absolute_source"]
        reject_reparse(source, REPO)
        data = source.read_bytes()
        if len(data) != int(row["bytes"]) or sha(data) != row["sha256"]:
            raise CurationError(f"build archive payload does not match its manifest: {rel}")
        reports.append(payload("reports", f"build/{rel}", source, original))

    diagnostics = []
    # PE Interactive exists in the preserved raw run; keep its recorded original path distinct.
    info = preserved_map.get("ViewerPETests.Interactive.output.log")
    pe_source = LAST_RUN / "ViewerPETests.Interactive.output.log"
    if info is None:
        raise CurationError("PE Interactive log is absent from preservation manifest")
    reject_reparse(pe_source, SANDBOX)
    pe_bytes = pe_source.read_bytes()
    if len(pe_bytes) != int(info["bytes"]) or sha(pe_bytes) != info["sha256"]:
        raise CurationError("preserved PE Interactive log differs from preservation manifest")
    diagnostics.append(payload("diagnostics", "ViewerPETests.Interactive.output.log", pe_source, info["original"]))
    # PluginContract/S3 output survived in the immutable entry checkpoint, but its original run path was removed.
    plugin_source = CHECKPOINT / PLUGIN_LOG
    original_plugin = str(SANDBOX / "runs" / RUN_ID / "artifacts" / "selftest" / "last_run" / "PluginContractTests.output.log")
    diagnostics.append(payload("diagnostics", "checkpoint/PluginContractTests.output.log", plugin_source, original_plugin))
    diagnostics.append(payload("diagnostics", "prepush-debug-queued-path.log", OUTER_RUN_LOG, str(OUTER_RUN_LOG)))
    return {"reports": reports, "diagnostics": diagnostics}

def verify_and_size(items: list[dict[str, object]]) -> int:
    total = 0
    for item in items:
        path = str(item["archivePath"])
        size = int(item["bytes"])
        if size > FILE_LIMIT:
            raise CurationError(f"per-file cap exceeded: {path}")
        total += size
    # Generated manifest/readme/attributes reserve 128 KiB of the strict 5 MiB cap.
    total += 128 * 1024
    if total > RUN_LIMIT:
        raise CurationError(f"selected group exceeds 5 MiB with metadata reserve: {total} bytes")
    return total

def plan() -> dict[str, object]:
    require_marker()
    inventory = build_inventory()
    return {group: {"payloadCount": len(items), "payloadBytes": sum(int(x["bytes"]) for x in items),
                    "preflightBytesWithMetadataReserve": verify_and_size(items), "files": [
                        {"archivePath":x["archivePath"],"bytes":x["bytes"],"sha256":x["sha256"],
                         "originalSource":x["originalSource"],"preservedSource":x["preservedSource"]} for x in items]}
            for group, items in inventory.items()}

def write_group(repo_root: Path, label: str, group: str, items: list[dict[str, object]]) -> dict[str, object]:
    verify_and_size(items)
    target = ARCHIVE_PARENT / label
    if target.exists() or reparse(target):
        raise CurationError(f"refusing to overwrite existing archive: {target}")
    reject_reparse(ARCHIVE_PARENT, REPO)
    staging = ARCHIVE_PARENT / f".{label}.partial-{uuid.uuid4().hex}"
    staging.mkdir(parents=False, exist_ok=False)
    try:
        rows = []
        for item in items:
            src = Path(str(item["preservedSource"]))
            if not src.is_file():
                raise CurationError(f"preserved source disappeared: {src}")
            reject_reparse(src, SANDBOX if within(src, SANDBOX) else REPO)
            data = src.read_bytes()
            if len(data) != int(item["bytes"]) or sha(data) != item["sha256"]:
                raise CurationError(f"preserved source changed after preflight: {src}")
            rel = str(item["archivePath"])
            dest = staging.joinpath(*PurePosixPath(rel).parts)
            dest.parent.mkdir(parents=True, exist_ok=True)
            reject_reparse(dest.parent, staging)
            dest.write_bytes(data)
            if dest.read_bytes() != data:
                raise CurationError(f"copy mismatch: {rel}")
            original = str(item["originalSource"])
            original_available = False
            if original:
                original_path = Path(original)
                if original_path.is_file():
                    reject_reparse(original_path, SANDBOX if within(original_path, SANDBOX) else REPO)
                    original_available = original_path.read_bytes() == data
            rows.append({"path": rel, "bytes": len(data), "sha256": sha(data),
                         "original_source_path": str(item["originalSource"]),
                         "preserved_source_path": str(src),
                         "original_source_available": original_available})
        manifest = "path\tbytes\tsha256\toriginal_source_path\tpreserved_source_path\toriginal_source_available\n" + "".join(
            f"{r['path']}\t{r['bytes']}\t{r['sha256']}\t{r['original_source_path']}\t{r['preserved_source_path']}\t{str(r['original_source_available']).lower()}\n" for r in rows)
        (staging / "manifest.tsv").write_text(manifest, encoding="utf-8", newline="")
        if group == "reports":
            details = (
                "The exact V2 run report records 2,127 passed, 17 failed, 261 skipped (2,405 total), with 20 entries executed, "
                "14 promoted and 6 incomplete. Terminal failures and incomplete entries were not promoted; the recorded reason "
                "codes are preserved verbatim and do not establish that source inputs changed during execution. They are not suite passes. "
                "The report, V1/V2 histories, dashboard, workspace snapshots, "
                "validation plan/state, decision snapshots, selected entry fingerprints/results, and exact Debug build provenance are retained.\n\n"
                f"[Interactive/S3 diagnostic archive](../{DIAGNOSTIC_LABEL}/README.md).\n"
            )
        else:
            details = (
                "ViewerPETests.Interactive output includes failures in restored foreground/keyboard focus and stopping before another child. "
                "PluginContractTests output records three FileSystemS3 failures: recursive Delete did not list the fixture prefix; "
                "a request with a body still arriving did not return after Cancel; and Cancel did not return a streaming request within a few body callbacks. "
                "The logged 15,000 ms return is supporting timing evidence for the cancellation failures. The outer runner log is retained byte-exact. These are preserved diagnostics, "
                "not a new reproduction or a claim that every overall failure has been assigned a cause.\n\n"
                f"[Run reports and validation checkpoint](../{REPORT_LABEL}/README.md).\n"
            )
        readme = (
            f"# Debug PrePush {group} evidence\n\nRun `{RUN_ID}`. Original overall verdict: **FAILED**. "
            "This archive preserves scoped reports or diagnostic logs only; it is not a qualification pass. "
            "Original source paths and surviving preserved copies are separate manifest columns. Availability is true only "
            "when the original path still exists with identical bytes; unavailable sources are never reconstructed.\n\n"
            + details +
            "The 2 MiB/file and 5 MiB/archive limits are enforced before writing. The oversized 2,269,578-byte "
            "perf_metrics.jsonl is intentionally omitted; fixture/work trees are not selected.\n"
        ).encode()
        (staging / "README.md").write_bytes(readme)
        (staging / ".gitattributes").write_bytes(b"* -text\n")
        files = [p for p in staging.rglob("*") if p.is_file()]
        if any(p.stat().st_size > FILE_LIMIT for p in files) or sum(p.stat().st_size for p in files) > RUN_LIMIT:
            raise CurationError("final archive exceeds the enforced caps")
        for row in rows:
            dest = staging.joinpath(*PurePosixPath(row["path"]).parts)
            data = dest.read_bytes()
            if len(data) != row["bytes"] or sha(data) != row["sha256"]:
                raise CurationError(f"archive payload replay failed: {dest}")
            preserved = Path(row["preserved_source_path"])
            if data != preserved.read_bytes():
                raise CurationError(f"preserved source replay failed: {preserved}")
        if target.exists() or reparse(target):
            raise CurationError(f"archive target appeared during curation: {target}")
        staging.rename(target)
        return {"archive": str(target), "payloadCount": len(rows),
                "payloadBytes": sum(r["bytes"] for r in rows),
                "archiveBytes": sum(p.stat().st_size for p in target.rglob("*") if p.is_file())}
    except Exception:
        if staging.exists():
            reject_reparse(staging, ARCHIVE_PARENT)
            shutil.rmtree(staging)
        raise

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true", help="write the two preflighted archives after review")
    args = parser.parse_args()
    try:
        require_marker()
        inventory = build_inventory()
        summary = {group: {"payloadCount":len(items),"payloadBytes":sum(int(x["bytes"]) for x in items),
                           "preflightBytesWithMetadataReserve":verify_and_size(items)} for group,items in inventory.items()}
        if not args.execute:
            print(json.dumps({"mode":"read-only-plan","runId":RUN_ID,"groups":summary},indent=2))
            return 0
        reject_reparse(ARCHIVE_PARENT,REPO)
        ARCHIVE_PARENT.mkdir(parents=True,exist_ok=True)
        reject_reparse(ARCHIVE_PARENT,REPO)
        result=[]
        for group,label in (("reports",REPORT_LABEL),("diagnostics",DIAGNOSTIC_LABEL)):
            result.append(write_group(REPO,label,group,inventory[group]))
        print(json.dumps({"runId":RUN_ID,"archives":result},indent=2))
        return 0
    except (CurationError,OSError,ValueError,KeyError) as exc:
        print(f"debug-prepush-curator: {exc}",file=__import__("sys").stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Explicit, bounded curation proposal for the 2026-10-09 Release PrePush diagnostics.

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

RUN_ID = "20261009T213050Z-43264-4bde93ea855e44c885a3277a68e89038"
SANDBOX = Path("D:/RedSalamander.Perf").resolve()
REPO = Path(r"C:/Users/PVB/.codex/worktrees/viewer-space-explorer-analyst/RedSalamander").resolve()
EVIDENCE = SANDBOX / "evidence" / "viewer-space-finish-all-20261009"
PRESERVED = EVIDENCE / "prepush-release-queued-path"
LAST_RUN = PRESERVED / "artifacts" / "selftest" / "last_run"
PRESERVATION_MANIFEST = PRESERVED / "artifacts" / "selftest" / "preservation-manifest.tsv"
CHECKPOINT = SANDBOX / "evidence" / "runs" / RUN_ID
OUTER_RUN_LOG = EVIDENCE / "prepush-release-queued-path.log"
BUILD_ARCHIVE = REPO / "Specs" / "TestRuns" / "7d3a1247382a" / "Build" / "space_candidate_release_prepush_reattest_20261009T214710Z"
ARCHIVE_PARENT = REPO / "Specs" / "TestRuns" / "7d3a1247382a" / "Validation"
FILE_LIMIT = 2 * 1024 * 1024
RUN_LIMIT = 5 * 1024 * 1024
EXPECTED_RECEIPT = "37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd"
EXPECTED_SOURCE = "9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554"

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
    "viewer-pe.interactive", "tooling.pester",
)
ENTRY_COMPONENTS = {
    "selftest.commands": (
        "evidence/f9d54cb39a23e80bd03999315e99d36699fc6841415e02da20c9f315e97370b1.json",
        "results/results.json",
        "terminal-results/d4d4d8ba596789a10baef1d4390186a5c222a1f56a2808b7d4ac32b0c7477182.json",
    ),
    "selftest.compare-directories": (
        "evidence/8a3f71759c098120cc0dff79daacdb9f476c7d6bbefdabe067898f781399dbeb.json",
        "results/results.json",
        "terminal-results/fc7273c12e0ed7dfe6d13c336bc546713e737bc87c38cd0e0807e0b8e0a4fe9a.json",
    ),
    "selftest.file-operations": (
        "evidence/89772b83f9ac119abfdfbab0fc7c2554ea5df0424f249a63a8c21ac046b2359a.json",
        "results/results.json",
        "terminal-results/9a76df4211567dbf3550e808f39319cdf2db1778d7563407680e806b19577332.json",
    ),
    "viewer-pe.interactive": (
        "evidence/96430cb7a19648a42bf1b46c69a178a0a209a73eef9c78996ef15d5954a84a66.json",
        "results/ViewerPETests.Interactive.output.log",
        "terminal-results/66e4cf44a288fe32bd1505e562018e9446d60344b840d9912944a2ab0972417f.json",
    ),
    "tooling.pester": (
        "evidence/4fb5cd35c5abe2e95a2e9d598b95c5e04dd7e97b157d44682f407ae2f463eed7.json",
        "terminal-results/3a31fa2bf1e0ace39801a9c9b6a5e6c77d1a8be47b0c579faa307f48ae9e2fc6.json",
    ),
}
NONINTERACTIVE_COMPONENTS = (
    "evidence/61b2b9e0e85303275c9ab912102c1d7775372866b93fbb14520ee10c3bb07bda.json",
    "results/ViewerPETests.Noninteractive.output.log",
    "promotions/0c5085c73a179fa61f1c64ba1f40c3011cbc7e23e3ef59055633a7224b6834af.json",
    "terminal-results/fc192dc0585e256a0fc3fbf2a35d5e9131ffa4299502e0b2bb8ec14fce9b3467.json",
)
DECISIONS = (
    "decisions/32218234e2aab4647165e7bd476b3010836a98b567fa324eb97638e12a48dc64.json",
    "decisions/6a50ad68d4fc600bbf4633f3f37370301e883b19751aefa02d837d67077a271e.json",
)
REPORT_LABEL = "space_prepush_release_failure_reports_20261009T213050Z"
DIAGNOSTIC_LABEL = "space_prepush_release_viewer_diagnostics_20261009T213050Z"
PLUGIN_LOG = "entries/standalone.plugin-contract/results/PluginContractTests.output.log"
PE_LOG = "entries/viewer-pe.interactive/results/ViewerPETests.Interactive.output.log"
PLUGIN_COMPONENTS = (
    "evidence/3c361108a93f6caf2598904bd67792ffab3ff1632223b24b8dd00ecf80b60f6e.json",
    "promotions/04ecbce5c3be1636c7b42b64af256e47a03d4bc54581e3c2f21cd2726c263754.json",
    "terminal-results/ff500477519044648df06e162efd19813b0bcc2b86dacb41220205a404dc4675.json",
)
TERMINAL_EXPECTATIONS = {
    "selftest.commands": ("terminal-results/d4d4d8ba596789a10baef1d4390186a5c222a1f56a2808b7d4ac32b0c7477182.json", "FAILED", {"passed": 938, "failed": 8, "skipped": 4, "total": 950}),
    "selftest.compare-directories": ("terminal-results/fc7273c12e0ed7dfe6d13c336bc546713e737bc87c38cd0e0807e0b8e0a4fe9a.json", "FAILED", {"passed": 241, "failed": 1, "skipped": 30, "total": 272}),
    "selftest.file-operations": ("terminal-results/9a76df4211567dbf3550e808f39319cdf2db1778d7563407680e806b19577332.json", "FAILED", {"passed": 3, "failed": 1, "skipped": 229, "total": 233}),
    "tooling.pester": ("terminal-results/3a31fa2bf1e0ace39801a9c9b6a5e6c77d1a8be47b0c579faa307f48ae9e2fc6.json", "FAILED", {"passed": 916, "failed": 5, "skipped": 1, "total": 922}),
    "viewer-pe.interactive": ("terminal-results/66e4cf44a288fe32bd1505e562018e9446d60344b840d9912944a2ab0972417f.json", "FAILED", {"passed": 0, "failed": 1, "skipped": 0, "total": 1}),
    "viewer-pe.noninteractive": ("terminal-results/fc192dc0585e256a0fc3fbf2a35d5e9131ffa4299502e0b2bb8ec14fce9b3467.json", "PASSED", {"passed": 1, "failed": 0, "skipped": 0, "total": 1}),
    "standalone.plugin-contract": ("terminal-results/ff500477519044648df06e162efd19813b0bcc2b86dacb41220205a404dc4675.json", "PASSED", {"passed": 1, "failed": 0, "skipped": 0, "total": 1}),
}

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

def validate_terminal(entry_id: str) -> None:
    rel, expected_status, expected_counts = TERMINAL_EXPECTATIONS[entry_id]
    path = CHECKPOINT / "entries" / entry_id / Path(*PurePosixPath(rel).parts)
    reject_reparse(path, SANDBOX)
    record = json.loads(path.read_text(encoding="utf-8-sig"))
    if (record.get("entry_id") != entry_id or record.get("status") != expected_status or
            record.get("counts") != expected_counts):
        raise CurationError(f"terminal result does not match the reviewed Release run for {entry_id}")

def validate_result_binding(entry_id: str, result_relative: str, result_path: Path) -> None:
    reject_reparse(result_path, SANDBOX)
    if not within(result_path.resolve(strict=True), SANDBOX):
        raise CurationError("checkpoint result escaped the marked sandbox")
    if entry_id == "standalone.plugin-contract":
        evidence_rel = PLUGIN_COMPONENTS[0]
    elif entry_id == "viewer-pe.noninteractive":
        evidence_rel = NONINTERACTIVE_COMPONENTS[0]
    else:
        evidence_rel = ENTRY_COMPONENTS[entry_id][0]
    evidence_path = CHECKPOINT / "entries" / entry_id / Path(*PurePosixPath(evidence_rel).parts)
    reject_reparse(evidence_path, SANDBOX)
    evidence = json.loads(evidence_path.read_text(encoding="utf-8-sig"))
    expected_path = f"entries/{entry_id}/{result_relative}"
    result_bytes = result_path.read_bytes()
    expected = [item for item in evidence.get("result_artifacts", []) if item.get("path") == expected_path]
    if len(expected) != 1 or expected[0].get("size_bytes") != len(result_bytes) or expected[0].get("sha256") != sha(result_bytes):
        raise CurationError(f"checkpoint evidence does not bind exact result bytes for {entry_id}")

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
    run_summary = json.loads((LAST_RUN / "run-all-tests-results.json").read_text(encoding="utf-8-sig"))
    if (run_summary.get("run_id") != RUN_ID or run_summary.get("build_receipt_id") != EXPECTED_RECEIPT or
            run_summary.get("repository_verdict") != "FAILED" or run_summary.get("change_verdict") != "FAILED" or run_summary.get("requested_gate") != "Full" or
            run_summary.get("validation_mode") != "Fresh" or
            run_summary.get("case_counts") != {"passed": 2125, "failed": 16, "skipped": 264, "total": 2405} or
            run_summary.get("counts") != {"executed": 20, "incomplete": 5, "invalidated": 0, "promoted": 15, "provisional": 0, "reused": 0}):
        raise CurationError("preserved Release run report does not match the reviewed run/count pins")
    v1_summary = json.loads((LAST_RUN / "run-all-tests-results-v1.json").read_text(encoding="utf-8-sig"))
    if (v1_summary.get("passed"), v1_summary.get("failed"), v1_summary.get("skipped"), v1_summary.get("total")) != (2125, 16, 264, 2405):
        raise CurationError("preserved Release V1 summary does not match the reviewed case counts")
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
            if subpath.startswith("results/"):
                validate_result_binding(entry, subpath, source)
            reports.append(payload("reports", f"checkpoint/entries/{entry}/{subpath}", source))
        validate_terminal(entry)

    # Bind this run to its already archived exact Release re-attestation receipt, sidecar and log.
    build_manifest = BUILD_ARCHIVE / "manifest.tsv"
    reject_reparse(build_manifest, REPO)
    if not within(build_manifest.resolve(strict=True), REPO):
        raise CurationError("build archive manifest escaped the repository")
    with build_manifest.open(encoding="utf-8-sig", newline="") as stream:
        build_rows = {row["path"]: row for row in csv.DictReader(stream, delimiter="\t")}
    for rel in ("build-receipt.json", "msbuild.log", "outer.log", "RedSalamander.exe.DxUi.json"):
        row = build_rows.get(rel)
        if row is None:
            raise CurationError(f"exact Release re-attestation archive lacks {rel}")
        source = BUILD_ARCHIVE / rel
        original = row["absolute_source"]
        reject_reparse(source, REPO)
        data = source.read_bytes()
        if len(data) != int(row["bytes"]) or sha(data) != row["sha256"]:
            raise CurationError(f"build archive payload does not match its manifest: {rel}")
        if rel == "build-receipt.json":
            receipt = json.loads(data.decode("utf-8-sig"))
            if (receipt.get("receipt_id") != EXPECTED_RECEIPT or receipt.get("source_snapshot_id") != EXPECTED_SOURCE or
                    receipt.get("configuration") != "Release" or receipt.get("platform") != "x64" or
                    receipt.get("tests_enabled") is not True or receipt.get("exit_code") != 0):
                raise CurationError("archived Release build receipt does not match the reviewed receipt/source pins")
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
        raise CurationError("preserved Release PE Interactive log differs from preservation manifest")
    diagnostics.append(payload("diagnostics", "ViewerPETests.Interactive.output.log", pe_source, info["original"]))
    pe_checkpoint = CHECKPOINT / PE_LOG
    reject_reparse(pe_checkpoint, SANDBOX)
    if pe_checkpoint.read_bytes() != pe_bytes:
        raise CurationError("preserved PE Interactive log differs from the exact checkpoint output")
    validate_result_binding("viewer-pe.interactive", "results/ViewerPETests.Interactive.output.log", CHECKPOINT / PE_LOG)
    # The Release PluginContract entry passed; its checkpoint log records which Debug-only gates were skipped.
    plugin_source = CHECKPOINT / PLUGIN_LOG
    original_plugin = str(SANDBOX / "runs" / RUN_ID / "artifacts" / "selftest" / "last_run" / "PluginContractTests.output.log")
    reject_reparse(plugin_source, SANDBOX)
    plugin_dir = plugin_source.parent
    if sorted(p.name for p in plugin_dir.iterdir() if p.is_file()) != [plugin_source.name]:
        raise CurationError("Release PluginContract checkpoint output is missing or ambiguous")
    diagnostics.append(payload("diagnostics", "checkpoint/PluginContractTests.output.log", plugin_source, original_plugin))
    validate_result_binding("standalone.plugin-contract", "results/PluginContractTests.output.log", plugin_source)
    validate_terminal("standalone.plugin-contract")
    plugin_root = CHECKPOINT / "entries" / "standalone.plugin-contract"
    for subpath in PLUGIN_COMPONENTS:
        source = plugin_root.joinpath(*PurePosixPath(subpath).parts)
        directory = source.parent
        reject_reparse(directory, SANDBOX)
        direct_files = sorted(p.name for p in directory.iterdir() if p.is_file())
        if direct_files != [source.name]:
            raise CurationError(f"Release PluginContract checkpoint component is missing or ambiguous for {subpath}: {direct_files}")
        reject_reparse(source, SANDBOX)
        diagnostics.append(payload("diagnostics", f"checkpoint/standalone.plugin-contract/{subpath}", source))
    # Keep the successful noninteractive parent output beside its immutable entry binders.
    noninteractive_root = CHECKPOINT / "entries" / "viewer-pe.noninteractive"
    for subpath in NONINTERACTIVE_COMPONENTS:
        source = noninteractive_root.joinpath(*PurePosixPath(subpath).parts)
        directory = source.parent
        reject_reparse(directory, SANDBOX)
        direct_files = sorted(p.name for p in directory.iterdir() if p.is_file())
        if direct_files != [source.name]:
            raise CurationError(f"successful PE noninteractive component is missing or ambiguous for {subpath}: {direct_files}")
        reject_reparse(source, SANDBOX)
        if subpath.startswith("results/"):
            validate_result_binding("viewer-pe.noninteractive", subpath, source)
        original = ""
        if subpath == "results/ViewerPETests.Noninteractive.output.log":
            original = str(SANDBOX / "runs" / RUN_ID / "artifacts" / "selftest" / "last_run" / "ViewerPETests.Noninteractive.output.log")
        diagnostics.append(payload("diagnostics", f"checkpoint/viewer-pe.noninteractive/{subpath}", source, original))
    validate_terminal("viewer-pe.noninteractive")
    diagnostics.append(payload("diagnostics", "prepush-release-queued-path.log", OUTER_RUN_LOG, str(OUTER_RUN_LOG)))
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
                "The exact V2 run report records 2,125 passed, 16 failed, 264 skipped (2,405 total), with 20 entries executed, "
                "15 promoted and 5 incomplete. The failed terminal entries account for Commands 938 passed/8 failed/4 skipped, "
                "CompareDirectories 241/1/30, FileOperations 3/1/229, Tools Pester 916/5/1, and ViewerPETests.Interactive 0/1/0. "
                "The 16 failed witnesses are preserved in case history: Commands reports eight named failures; CompareDirectories has "
                "local_search_callback_contract (expected ERROR_CANCELLED, got S_OK); FileOperations has FileOps_SafetyRemediationContracts; "
                "Pester has five named reparse/sandbox failures; ViewerPE Interactive has one terminal failure plus its assertion log. "
                "Incomplete entries were not promoted; SOURCE_INPUT_CHANGED is preserved as the recorded reason code, not evidence of a runtime source mutation. "
                "The report, V1/V2 histories, dashboard, workspace snapshots, validation plan/state, decision snapshots, exact failed entry "
                "fingerprints/results, and Release build provenance are retained.\n\n"
                "Commands failures: file_action_selected_paths_identity_cleanup; file_action_selected_paths_recovery_bounds; "
                "file_operations_phase0_typed_contract_source_guard; cmd_preferences_dialog_shell_footer_live_dx_interaction; "
                "cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions; "
                "cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics; "
                "cmd_pane_fileops_conflict_remote_case_sensitive_nested_inspection; folderView_thumbnail_sort_popup_slider. "
                "Pester failures: rejects a final-file reparse destination without changing its external target; rejects a real directory "
                "symbolic-link ancestor; rejects a real file symbolic link and a dangling symbolic-link component; refuses a test link "
                "that would live or point outside its marked sandbox, before creating it; rejects reparse inputs and unsafe impact-base arguments without fetching.\n\n"
                f"[Viewer and PluginContract diagnostics](../{DIAGNOSTIC_LABEL}/README.md).\n"
            )
        else:
            details = (
                "The broad ViewerPETests.Interactive log reached ViewerSpace activation at output line 315: the keyboard fixture activated only its owned window "
                "under the directed-input warning and passed its focus, F5, Escape, and provider teardown assertions. Later the full PE entry "
                "failed during Web/Shell churn: restored foreground/keyboard focus failed for TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo; "
                "TestViewerShellComboHostsLongRunOpenCloseStayStable reported child completion/restoration failures; the log then stopped before another child. "
                "The terminal entry is 0 passed/1 failed/0 skipped; five [FAIL] witness lines are at 1724, 1725, 1748, 1750, and 1753. "
                "The separate ViewerPETests.Noninteractive entry passed and its exact seven ViewerSpace child harness outcomes, empty-job teardown, "
                "promotion, terminal result, and evidence binding are retained together: TestViewerSpaceScanPolicyHelpers, "
                "TestViewerSpaceAnalysisHostVirtualGridAndLifetime, TestViewerSpaceLifecycleCaptures, TestViewerSpaceListingValidation, "
                "TestViewerSpaceBoundedHostileProviderScanning, TestViewerSpaceBlockedProviderCloseAndPostUpdateRace, and "
                "TestViewerSpaceDirectorySummaryProtocol. The plan records this entry as noninteractive (`requires_interactive_desktop=false`, "
                "`--group=noninteractive`); the log records the seven launches/pass/empty-job outcomes at lines 692–1239. The original last_run copy of this "
                "noninteractive log is unavailable; the immutable checkpoint output and binders are the preserved source.\n\n"
                "Release standalone.plugin-contract passed. Its log explicitly marks the general FileSystemS3 Debug selftest export as skipped because "
                "that export is not expected in this configuration (entry terminal count 1 passed, 0 failed, 0 skipped). S3 multipart and directory-transfer probe groups did pass; neither constitutes "
                "execution of the skipped general Debug FileSystemS3 selftest gate.\n\n"
                "The outer runner log is retained byte-exact. These are preserved diagnostics, not a new reproduction or an overall qualification pass.\n\n"
                f"[Run reports and validation checkpoint](../{REPORT_LABEL}/README.md).\n"
            )
        readme = (
            f"# Release PrePush {group} evidence\n\nRun `{RUN_ID}`. Original overall verdict: **FAILED**. "
            "This archive preserves scoped reports or diagnostic logs only; it is not a qualification pass. "
            "Original source paths and surviving preserved copies are separate manifest columns. Availability is true only "
            "when the original path still exists with identical bytes; unavailable sources are never reconstructed.\n\n"
            + details +
            "The 2 MiB/file and 5 MiB/archive limits are enforced before writing. The oversized 2,285,134-byte "
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
    parser.add_argument("--execute", action="store_true", help="write the two preflighted archives after root review")
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
        print(f"release-prepush-curator: {exc}",file=__import__("sys").stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())

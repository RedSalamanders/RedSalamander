"""DRAFT ONLY: preflight and curate direct-Issues Release evidence after archival.

The baseline gallery is expected to fail its new Issues witnesses. This script
never upgrades that failure to a pass and makes no Debug, UIA, foreground,
Commands, PR, Full, or final performance qualification claim.
"""

import argparse
import hashlib
import html
import importlib.util
import json
import math
import re
import shutil
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = Path("Z:/src/RedSalamander")
EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
FAILED_COMPILE = EVIDENCE / "direct-issues-baseline-20260923"
BASELINE = EVIDENCE / "direct-issues-baseline-v2-20260923"
DEFAULT_CANDIDATE = EVIDENCE / "direct-issues-qualified-20260923"
DEFAULT_PACKET = ROOT / "Specs/TestRuns/4cb089111a23/Validation/2026-09-23_direct-issues"
DEFAULT_GALLERY = ROOT / "Specs/Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-direct-issues"
BASELINE_PRIMARY_HEAD = "199728b88a43afb55f1b26f2426c2b0fc1c535a1"
CANDIDATE_PRIMARY_HEAD = "bf5506bd00babf55e9d9ef80f0af854b0903199c"
ISOLATED_BUILD_HEAD = "ccebd5af68eb45f2e9deed64a7063780e71eb1ce"
ADDED = {
    "RedSalamander/FolderWindow.FileOperations.IssuesPane.h",
    "RedSalamander/FolderWindow.FileOperationsInternal.h",
    "RedSalamander/FolderWindow.FileOperations.State.Diagnostics.cpp",
    "RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp",
}
CHANGED = {
    "RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp",
    "RedSalamander/FolderWindow.FileOperations.Popup.cpp",
    "RedSalamander/FolderWindow.FileOperations.Popup.h",
}
COMMON_CHECKS = {"StatusChipPaintFit": 12, "DoneHeaderPaint": 8,
                 "CompletedGroupBadgePaint": 4, "HostedProgressValue": 18,
                 "CompletedGroupLabelFit": 28, "TerminalFactsFit": 8,
                 "TerminalFactsTransition": 24, "TerminalFactsStableRepaint": 2,
                 "SpeedLimitLabelFit": 148}
STAGES = {"row-removed-counters-retained", "row-arrived", "row-without-counters",
          "issues-without-more", "collapsed", "expanded", "footer-hidden",
          "footer-restored", "row-removed-zero-counters", "original-restored"}


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def file_at(path):
    require(path.is_file(), f"Missing required evidence: {path}")
    return path


def raw_text(path):
    raw = file_at(path).read_bytes()
    return raw.decode("utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def read_json(path):
    return json.loads(raw_text(path))


def digest(path):
    return hashlib.sha256(file_at(path).read_bytes()).hexdigest()


def exit_file(path, expected):
    require(raw_text(path).strip() == str(expected), f"Wrong exit value: {path}")


def manifest(folder, expected_count):
    original = read_json(folder / "candidate-inputs.json")
    require(isinstance(original, list) and len(original) == expected_count,
            f"Unexpected source manifest size in {folder}")
    result = {}
    for row in original:
        name, value = row["path"], row["sha256"].lower()
        require(name not in result and re.fullmatch(r"[0-9a-f]{64}", value),
                f"Duplicate or malformed input: {name}")
        require(digest(folder / "inputs" / name) == value, f"Frozen input drift: {name}")
        result[name] = value
    return result


def check_build(folder):
    receipt = read_json(folder / "build-receipt.json")
    require(receipt["configuration"] == "Release" and receipt["platform"] == "x64"
            and receipt["tests_enabled"] is True and receipt["exit_code"] == 0,
            f"Wrong build profile/outcome: {folder}")
    for key, size in (("git_head", 40), ("source_snapshot_id", 64), ("receipt_id", 64)):
        require(re.fullmatch(rf"[0-9a-f]{{{size}}}", receipt[key]), f"Malformed {key}: {folder}")
    log = raw_text(folder / "build.log")
    diagnostic = re.search(r"Diagnostics:\s*(\d+) warning\(s\),\s*(\d+) error\(s\)", log)
    require(diagnostic is not None and diagnostic.groups() == ("0", "0"),
            f"Unexpected build diagnostics: {folder}")
    exit_file(folder / "build-exit.txt", 0)
    head = raw_text(folder / "primary-head.txt").strip()
    require(re.fullmatch(r"[0-9a-f]{40}", head), f"Malformed primary HEAD: {folder}")
    # Primary source and isolated build HEADs are separate identities. Do not equate them.
    return receipt, head


def check_failed_compile():
    exit_file(FAILED_COMPILE / "build-exit.txt", 1)
    log = raw_text(FAILED_COMPILE / "build.log")
    require("error C2513" in log and "warning C4365" in log
            and "Diagnostics: 1 warning(s), 1 error(s)" in log,
            "Initial near-macro/UINT compile failure differs")
    require(not (FAILED_COMPILE / "build-receipt.json").exists(),
            "Failed compile unexpectedly has successful build receipt")
    return {"head": raw_text(FAILED_COMPILE / "primary-head.txt").strip(),
            "manifest": manifest(FAILED_COMPILE, 25)}


def check_localization(candidate):
    log = raw_text(candidate / "localization-tests.log")
    require(re.search(r"(?m)^Passed:\s*9\s+Failed:\s*0\s+Skipped:\s*0\s+Pending:\s*0\s+Inconclusive:\s*0\s*$", log)
            and re.search(r"(?m)^TotalCount\s*:\s*9\s*$", log)
            and re.search(r"(?m)^PassedCount\s*:\s*9\s*$", log)
            and re.search(r"(?m)^FailedCount\s*:\s*0\s*$", log)
            and re.search(r"(?m)^SkippedCount\s*:\s*0\s*$", log),
            "Raw Pester localization transcript does not show 9/9 passes")
    note = raw_text(candidate / "localization-wrapper-note.md")
    require("exit 1" in note and "redirected the PassThru object" in note
            and "no test failure or passing wrapper exit is claimed" in note,
            "Localization wrapper-failure provenance note changed")


def find_gallery(folder):
    matches = []
    for root in folder.glob("*-runner-artifacts/artifacts/selftest/last_run"):
        result_path = root / "run-all-tests-results.json"
        if result_path.is_file():
            result = read_json(result_path)
            if result.get("case_filter") == "FileOps_VisualGallery":
                matches.append((root, result))
    require(len(matches) == 1, f"Expected exactly one archived gallery in {folder}: {matches}")
    root, result = matches[0]
    require(result["schema"] == "red-salamander.run-all-tests.v1"
            and result["suite"] == "FileOps" and result["platform"] == "x64"
            and result["configuration"] == "Release"
            and result["test_sandbox_audit"]["is_clean"] is True
            and result["test_sandbox_audit"]["issue_count"] == 0,
            f"Wrong run profile or disk audit: {folder}")
    cases = [row for suite in result["suites"] for row in suite["cases"]]
    selected = [row for row in cases if row["name"] == "FileOps_VisualGallery"]
    require(len(selected) == 1, f"Filtered gallery case missing: {folder}")
    logs = [path for path in folder.glob("*.log")
            if "FileOps_VisualGallery" in raw_text(path) and result["run_id"] in raw_text(path)]
    require(len(logs) == 1, f"Gallery log ambiguity in {folder}: {logs}")
    log = logs[0]
    plan = read_json(folder / f"{log.stem}-plan.json")
    require(plan["Id"] == "selftest.file-operations"
            and plan["RequiresInteractiveDesktop"] is False,
            f"Wrong runner plan: {folder}")
    require(f"verified receipt {read_json(folder / 'build-receipt.json')['receipt_id']}" in raw_text(log),
            f"Run log does not bind build receipt: {log}")
    exit_file(folder / f"{log.stem}-exit.txt", result["exit_code"])
    nested = read_json(root / "fileops/results.json")
    require(nested["passed"] == result["passed"] and nested["failed"] == result["failed"],
            f"Nested FileOps result differs: {folder}")
    return {"root": root, "result": result, "case": selected[0], "log": log}


def check_trace(path, baseline):
    lines = raw_text(path).splitlines()
    for key, expected in COMMON_CHECKS.items():
        matching = [line for line in lines if f"{key}:" in line]
        require(len(matching) == expected and all(re.search(r"\bfit=true\b", line) for line in matching),
                f"Existing witness count/fit drift: {key} in {path}")
    for key, expected in (("DirectIssuesActionFit", 16), ("DirectIssuesTransition", 20)):
        matching = [line for line in lines if f"{key}:" in line]
        require(len(matching) == expected, f"Direct Issues witness count differs: {key} in {path}")
        outcomes = [re.search(r"\bfit=(true|false)\b", line) for line in matching]
        require(all(outcomes), f"Unparseable {key} outcome in {path}")
        require((any(match.group(1) == "false" for match in outcomes) if baseline else
                 all(match.group(1) == "true" for match in outcomes)),
                f"Wrong expected {key} outcome in {path}")
    actions = [line for line in lines if "DirectIssuesActionFit:" in line]
    identities = [re.search(r"DirectIssuesActionFit: (\S+) dpi=(\d+) width=(\d+) theme=(\S+)", line)
                  for line in actions]
    require(all(identities) and {match.groups() for match in identities} ==
            {(scene, str(dpi), str(width), theme)
             for dpi in (96, 144) for width in (480, 760)
             for scene, themes in (("145-direct-issues-row", ("dark", "high-contrast")),
                                   ("146-diagnostic-counters-no-issues-row", ("dark",)),
                                   ("147-clean-no-issues", ("dark",))) for theme in themes},
            "Direct action capture identities differ")
    transitions = [line for line in lines if "DirectIssuesTransition:" in line]
    identities = [re.search(r"DirectIssuesTransition: dpi=(\d+) width=(\d+) stage=(\S+)", line)
                  for line in transitions]
    require(all(identities) and {match.groups() for match in identities} ==
            {(str(dpi), "480", stage) for dpi in (96, 144) for stage in STAGES},
            "Direct Issues transition identities differ")


def add_copy(copies, source, target):
    copies.append((file_at(source), target))


def generated_bytes(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")


def budget(copies, generated, root):
    entries = [(target, source.stat().st_size) for source, target in copies if target.is_relative_to(root)]
    entries += [(target, len(data)) for target, data in generated.items() if target.is_relative_to(root)]
    require(len({str(target) for target, _ in entries}) == len(entries), f"Duplicate payload target: {root}")
    require(all(size <= 2*1024*1024 for _, size in entries), f"File exceeds 2 MiB: {root}")
    require(sum(size for _, size in entries)+65536 <= 5*1024*1024, f"Packet exceeds 5 MiB: {root}")


def preflight(candidate, packet, gallery):
    require(not packet.exists() and not gallery.exists(), "Output exists; inspect before retrying")
    spec = importlib.util.spec_from_file_location("direct_issues_comparator", file_at(HERE / "compare-direct-issues-release.py"))
    comparator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(comparator)
    fresh_layout, fresh_capture = comparator.compare(candidate)
    require(read_json(candidate / "layout-comparison.json") == fresh_layout
            and read_json(candidate / "capture-comparison.json") == fresh_capture,
            "Saved comparison disagrees with fresh raw trace/PNG calculation")
    failed = check_failed_compile()
    folders = {"baseline": BASELINE, "candidate": candidate}
    receipts, heads, manifests, runs = {}, {}, {}, {}
    for label, folder in folders.items():
        receipts[label], heads[label] = check_build(folder)
        manifests[label] = manifest(folder, 25 if label == "baseline" else 29)
        runs[label] = find_gallery(folder)
        check_trace(runs[label]["root"] / "fileops/trace.txt", label == "baseline")
    require(failed["head"] == heads["baseline"] == BASELINE_PRIMARY_HEAD
            and heads["candidate"] == CANDIDATE_PRIMARY_HEAD,
            "Primary HEAD differs from the reviewed before/after facts commit")
    require(receipts["baseline"]["git_head"] == receipts["candidate"]["git_head"] == ISOLATED_BUILD_HEAD,
            "Isolated build HEAD differs from the matched frozen checkout")
    require(receipts["baseline"]["source_snapshot_id"] != receipts["candidate"]["source_snapshot_id"],
            "Source snapshots unexpectedly identical")
    for key in ("repo_identity", "project_graph_digest", "runtime_dependency_digest",
                "toolchain_identity_digest", "vcpkg_identity_digest"):
        require(receipts["baseline"][key] == receipts["candidate"][key],
                f"Dependency/toolchain drift: {key}")
    require(set(failed["manifest"]) == set(manifests["baseline"])
            and {name for name in failed["manifest"] if failed["manifest"][name] != manifests["baseline"][name]} ==
            {"RedSalamander/FolderWindow.FileOperations.Popup.cpp",
             "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallery.cpp"},
            "Compile correction changed unexpected sources")
    old, new = manifests["baseline"], manifests["candidate"]
    changed = {name for name in old if old[name] != new.get(name)}
    require(set(new)-set(old) == ADDED and not (set(old)-set(new))
            and changed == CHANGED,
            f"Candidate overlay delta needs review: added={sorted(set(new)-set(old))}, changed={sorted(changed)}")
    check_localization(candidate)
    for label, run in runs.items():
        result, case = run["result"], run["case"]
        expected = (0, 1, 1, "failed") if label == "baseline" else (1, 0, 0, "passed")
        require((result["passed"], result["failed"], result["exit_code"], case["status"]) == expected
                and result["total"] == 1,
                f"Wrong expected gallery outcome: {label}")
        if label == "baseline":
            require("Direct Issues" in case.get("reason", ""),
                    "Baseline gallery failed for an unrelated or unreported reason")
    selected = []
    for label, run in runs.items():
        cap_dir = run["root"] / "fileops/visual-gallery"
        for scene in ("145-direct-issues-row", "146-diagnostic-counters-no-issues-row", "147-clean-no-issues"):
            for display, dpi in ((1, 144), (2, 96)):
                matches = sorted(cap_dir.glob(f"{scene}__dark__480dip__fr-FR__display{display}__{dpi}dpi.png"))
                require(len(matches) == 1, f"Missing focused {label} capture: {scene}/{dpi}")
                selected.append((label, matches[0]))
    # Retain all new actual candidate captures, including both widths and high
    # contrast, beyond the six narrow paired baseline comparisons above.
    candidate_captures = runs["candidate"]["root"] / "fileops/visual-gallery"
    already_selected = {path.name for label, path in selected if label == "candidate"}
    for name in fresh_capture["newSceneNames"]:
        if name not in already_selected:
            selected.append(("candidate", file_at(candidate_captures / name)))
    require(len(selected) == 22, "Expected six baseline focus captures and all 16 candidate new-scene captures")
    copies = []
    for label, folder in folders.items():
        top = ["build.log", "build-exit.txt", "build-receipt.json", "candidate-inputs.json",
               "primary-head.txt", "build-and-capture.ps1", "capture-candidate-gallery.ps1"]
        if label == "baseline":
            top += ["red-baseline-witness.patch", "witness-correction.md"]
        else:
            top += ["production-and-witness.patch", "localization-tests.log", "localization-wrapper-note.md"]
        for name in top:
            add_copy(copies, folder / name, packet / "raw" / label / name)
        run = runs[label]
        run_dir = packet / "raw" / label / "gallery"
        for path in (run["log"], folder / f"{run['log'].stem}-plan.json", folder / f"{run['log'].stem}-exit.txt"):
            add_copy(copies, path, run_dir / path.name)
        for name in ("run-all-tests-results.json", "fileops/results.json", "fileops/trace.txt"):
            add_copy(copies, run["root"] / name, run_dir / name)
        for name in ("scenarios.tsv", "displays.tsv"):
            add_copy(copies, run["root"] / "fileops/visual-gallery" / name,
                     gallery / "captures" / label / name)
    for name in ("build.log", "build-exit.txt", "candidate-inputs.json", "primary-head.txt",
                 "red-baseline-witness.patch"):
        add_copy(copies, FAILED_COMPILE / name, packet / "raw/failed-compile" / name)
    for name in ("layout-comparison.json", "capture-comparison.json"):
        add_copy(copies, candidate / name, packet / "raw" / name)
    for name in ("compare-direct-issues-release.py", "curate-direct-issues-release.py"):
        add_copy(copies, HERE / name, packet / "raw" / name)
    add_copy(copies, candidate / "capture-comparison.json", gallery / "baseline-release-comparison.json")
    for label, path in selected:
        add_copy(copies, path, gallery / "captures" / label / path.name)
    changes = [row["usPerRedraw"]["changePercent"] for row in fresh_layout["pairedMedians"].values()]
    private = [row["privateBytes"]["delta"]/(1024*1024) for row in fresh_layout["pairedMedians"].values()]
    require(all(value is not None and math.isfinite(value) for value in changes+private),
            "Non-finite paired metric")
    narrative = (
        "# Direct Issues: matched Release evidence\n\n"
        f"The corrected baseline gallery `{runs['baseline']['result']['run_id']}` fails its new Direct Issues witnesses by design; "
        f"the candidate gallery `{runs['candidate']['result']['run_id']}` passes. Both have 532 filename-matched actual "
        "captures, including 16 new scene-145–147 PNGs. The initial baseline compile failure (C2513 near-macro "
        "collision and C4365 UINT conversion) remains archived separately; the corrected baseline build is the "
        "paired comparison source. The primary HEAD changed from the reviewed 199728b8 baseline to bf5506bd "
        "after terminal facts were committed between attempts; both isolated build receipts retain ccebd5af. "
        "The frozen input manifests bind the exact source delta, and the VisualGallery witness bytes are unchanged.\n\n"
        f"RGBA pixel comparison finds {fresh_capture['commonChangedImages']} changed images and "
        f"{fresh_capture['commonDimensionChanges']} dimension changes. The 24 shared five-round synchronous redraw "
        f"fixtures show paired median redraw-time changes from {min(changes):+.2f}% to {max(changes):+.2f}% and "
        f"private-byte changes from {min(private):+.2f} to {max(private):+.2f} MiB. These are in-process "
        "observations, not presented FPS, heap attribution, GPU memory, or final performance acceptance.\n\n"
        "The gallery trace checks 16 actual action-fit observations and 20 same-HWND transition observations "
        "across two native displays, including Issues without More. The raw localization transcript records "
        "9/9 passes; its first PowerShell wrapper exited 1 after redirecting the PassThru object, so no passing "
        "wrapper exit is claimed. The direct Issues Commands API case is "
        "candidate-only and requires an interactive desktop lease; this Release gallery packet does not validate it. "
        "Debug, UIA, manual input, PR, and Full validation remain separate.\n"
    )
    cards = "\n".join(
        f'<article><h2>{html.escape(label+": "+path.name)}</h2>'
        f'<img loading="lazy" src="captures/{label}/{html.escape(path.name, quote=True)}" '
        f'alt="{html.escape(label+": "+path.name, quote=True)}"></article>'
        for label, path in selected)
    generated = {
        packet / "README.md": generated_bytes(narrative +
            "\n[Focused actual captures](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-direct-issues/README.md).\n"),
        gallery / "README.md": generated_bytes(narrative +
            "\n[Gallery](gallery.html) · [All capture comparisons](baseline-release-comparison.json).\n"),
        gallery / "gallery.html": generated_bytes(
            '<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>Direct Issues Release captures</title><style>body{background:#15181d;color:#eee;font:16px system-ui;margin:24px}'
            'main{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:20px}'
            'article{padding:12px;background:#222831}h2{font-size:14px;overflow-wrap:anywhere}'
            'img{width:100%;height:650px;object-fit:contain}a{color:#acdaff}</style>'
            '<h1>Actual direct Issues captures</h1><p>Six focused baseline captures and all sixteen new candidate captures.</p>'
            '<p><a href="README.md">Evidence and limits</a></p><main>'+cards+'</main></html>\n'),
        packet / ".gitattributes": b"raw/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
        gallery / ".gitattributes": b"captures/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
    }
    for output in (packet, gallery):
        budget(copies, generated, output)
    return copies, generated, fresh_capture, fresh_layout


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=DEFAULT_CANDIDATE)
    parser.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    parser.add_argument("--gallery", type=Path, default=DEFAULT_GALLERY)
    args = parser.parse_args()
    copies, generated, capture, layout = preflight(args.candidate, args.packet, args.gallery)
    # Mutate outputs only after all provenance, witness, comparison and byte guards pass.
    for source, target in copies:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
        require(digest(source) == digest(target), f"Copy drift: {target}")
    for target, content in generated.items():
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(content)
    for output in (args.packet, args.gallery):
        paths = sorted(path for path in output.rglob("*") if path.is_file() and path.name != "SHA256SUMS.txt")
        content = "".join(f"{digest(path)}  {path.relative_to(output).as_posix()}\n" for path in paths)
        (output / "SHA256SUMS.txt").write_bytes(content.encode("utf-8"))
        require(all(path.stat().st_size <= 2*1024*1024 for path in output.rglob("*") if path.is_file())
                and sum(path.stat().st_size for path in output.rglob("*") if path.is_file()) <= 5*1024*1024,
                f"Post-write archive budget failed: {output}")
    print(json.dumps({"commonPairs": capture["commonPairs"], "changed": capture["commonChangedImages"],
                      "pairedRedrawFixtures": layout["fixtureCount"], "focusedActualCaptures": 22}, indent=2))


if __name__ == "__main__":
    main()

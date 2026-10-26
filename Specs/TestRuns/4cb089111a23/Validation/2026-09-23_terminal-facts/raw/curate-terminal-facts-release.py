"""DRAFT ONLY: preflight and curate the terminal-facts Release evidence pair.

Do not run until the final candidate (possibly v2) has archived both projection
and gallery runs and compare-terminal-facts-release.py has completed. The
baseline's *projection* fails deliberately; its 508-capture gallery passes.
No Debug, PR, Full, UIA, interactive, or final performance acceptance is claimed.
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

from PIL import Image, ImageChops

HERE = Path(__file__).resolve().parent
ROOT = Path("Z:/src/RedSalamander")
EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
BASELINE = EVIDENCE / "terminal-facts-baseline-20260923"
INTERMEDIATE = EVIDENCE / "terminal-facts-qualified-20260923"
DEFAULT_CANDIDATE = EVIDENCE / "terminal-facts-qualified-v2-20260923"
DEFAULT_PACKET = ROOT / "Specs/TestRuns/4cb089111a23/Validation/2026-09-23_terminal-facts"
DEFAULT_GALLERY = ROOT / "Specs/Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-terminal-facts"
CHANGED = {
    "RedSalamander/FolderWindow.FileOperations.Popup.cpp",
    "RedSalamander/FolderWindow.FileOperations.Popup.h",
    "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallery.cpp",
    "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp",
}
ADDED = {
    "RedSalamander/Lang/cs-CZ/RedSalamander-cs-CZ.rc",
    "RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc",
    "RedSalamander/Lang/ja-JP/RedSalamander-ja-JP.rc",
    "RedSalamander/Lang/sk-SK/RedSalamander-sk-SK.rc",
    "RedSalamander/RedSalamander.rc", "RedSalamander/resource.h",
    "Tools/Tests/ResourceLocalizationContracts.Tests.ps1", "Tools/tool-inventory.json",
}
COMMON_CHECKS = {"StatusChipPaintFit": 12, "DoneHeaderPaint": 8,
                 "CompletedGroupBadgePaint": 4, "HostedProgressValue": 18,
                 "CompletedGroupLabelFit": 28, "TerminalSummaryFit": 56,
                 "TerminalSummaryChange": 12, "SpeedLimitLabelFit": 148}
NEW_CHECKS = {"TerminalFactsFit": 8, "TerminalFactsTransition": 24,
              "TerminalFactsStableRepaint": 2}
STAGES = {"count-18", "unchanged-repaint", "collapsed", "expanded", "footer-hidden",
          "footer-restored", "count-restored", "count-1009", "settings-refreshed",
          "settings-unchanged", "summary-empty", "original-restored"}
FOCUS_COMMON = ("21-move-source-kept", "23-move-source-unknown",
                "31-completed-group-expanded")


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


def manifest(folder, baseline):
    original = read_json(folder / "candidate-inputs.json")
    require(isinstance(original, list), f"Malformed source manifest: {folder}")
    unique = {}
    duplicates = []
    for row in original:
        name, value = row["path"], row["sha256"].lower()
        require(re.fullmatch(r"[0-9a-f]{64}", value), f"Malformed source SHA: {name}")
        if name in unique:
            require(unique[name] == value, f"Conflicting duplicate input: {name}")
            duplicates.append(name)
        else:
            unique[name] = value
        require(digest(folder / "inputs" / name) == value, f"Frozen input drift: {name}")
    if baseline:
        require(len(original) == 18 and len(unique) == 17 and
                duplicates == ["RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp"],
                "Baseline's documented duplicate-manifest shape changed")
        corrected = read_json(folder / "candidate-inputs-unique.json")
        require(len(corrected) == 17 and {row["path"]: row["sha256"].lower() for row in corrected} == unique,
                "Corrected 17-source baseline view disagrees with original manifest")
        require("18 entries for 17 unique source files" in raw_text(folder / "manifest-note.md"),
                "Missing duplicate-manifest provenance note")
    else:
        require(len(original) == len(unique) == 25 and not duplicates,
                "Candidate source manifest needs explicit review for this script")
    return unique


def check_build(folder, baseline):
    receipt = read_json(folder / "build-receipt.json")
    require(receipt["configuration"] == "Release" and receipt["platform"] == "x64"
            and receipt["tests_enabled"] is True and receipt["exit_code"] == 0,
            f"Wrong build outcome/profile: {folder}")
    for key, size in (("git_head", 40), ("source_snapshot_id", 64), ("receipt_id", 64)):
        require(re.fullmatch(rf"[0-9a-f]{{{size}}}", receipt[key]), f"Malformed {key}: {folder}")
    log = raw_text(folder / "build.log")
    diagnostic = re.search(r"Diagnostics:\s*(\d+) warning\(s\),\s*(\d+) error\(s\)", log)
    require(diagnostic is not None and int(diagnostic.group(2)) == 0,
            f"Build errors or missing diagnostic summary: {folder}")
    if baseline:
        warnings = re.findall(r"warning C5258: explicit capture of 'taskId' is not required", log)
        require(diagnostic.group(1) == "2" and len(warnings) == 2,
                "Baseline warning identity/count differs from the archived two C5258 witness warnings")
    else:
        require(diagnostic.group(1) == "0", f"Candidate build warnings remain: {folder}")
    exit_file(folder / "build-exit.txt", 0)
    head = raw_text(folder / "primary-head.txt").strip()
    require(re.fullmatch(r"[0-9a-f]{40}", head), f"Malformed primary HEAD: {folder}")
    return receipt, head


def find_run(folder, case):
    roots = list(folder.glob("*-runner-artifacts/artifacts/selftest/last_run"))
    matches = []
    for root in roots:
        result_path = root / "run-all-tests-results.json"
        if result_path.is_file():
            result = read_json(result_path)
            if result.get("case_filter") == case:
                matches.append((root, result))
    require(len(matches) == 1, f"Expected one archived {case} run in {folder}; found {len(matches)}")
    root, result = matches[0]
    require(result["schema"] == "red-salamander.run-all-tests.v1"
            and result["suite"] == "FileOps" and result["platform"] == "x64"
            and result["configuration"] == "Release"
            and result["test_sandbox_audit"]["is_clean"] is True
            and result["test_sandbox_audit"]["issue_count"] == 0,
            f"Wrong run profile or disk audit: {root}")
    cases = [row for suite in result["suites"] for row in suite["cases"]]
    selected = [row for row in cases if row["name"] == case]
    require(len(selected) == 1, f"Filtered case missing from {root}")
    log_candidates = [p for p in folder.glob("*.log")
                      if case in raw_text(p) and result["run_id"] in raw_text(p)]
    require(len(log_candidates) == 1, f"Run log ambiguity for {case}: {log_candidates}")
    log = log_candidates[0]
    prefix = log.stem
    plan, exit_path = folder / f"{prefix}-plan.json", folder / f"{prefix}-exit.txt"
    plan_json = read_json(plan)
    require(plan_json["Id"] == "selftest.file-operations"
            and plan_json["RequiresInteractiveDesktop"] is False,
            f"Runner plan is interactive or wrong: {plan}")
    require(f"verified receipt {read_json(folder / 'build-receipt.json')['receipt_id']}" in raw_text(log),
            f"Runner log does not bind build receipt: {log}")
    exit_file(exit_path, result["exit_code"])
    nested = read_json(root / "fileops/results.json")
    require(nested["passed"] == result["passed"] and nested["failed"] == result["failed"],
            f"Nested FileOps result differs: {root}")
    return {"root": root, "result": result, "case": selected[0], "log": log,
            "plan": plan, "exit": exit_path}


def check_runs(folder, baseline):
    projection = find_run(folder, "FileOps_TerminalFactsProjection")
    gallery = find_run(folder, "FileOps_VisualGallery")
    expected = (2, 1, 1, "failed") if baseline else (3, 0, 0, "passed")
    p, f, x, status = expected
    result = projection["result"]
    require((result["passed"], result["failed"], result["exit_code"], result["total"],
             projection["case"]["status"]) == (p, f, x, 3, status),
            f"Unexpected projection result: {folder}")
    if baseline:
        require("lost independent completed publication" in projection["case"]["reason"],
                "Baseline projection failed for an unrelated reason")
    g = gallery["result"]
    require((g["passed"], g["failed"], g["exit_code"], g["total"],
             gallery["case"]["status"]) == (1, 0, 0, 1, "passed"),
            f"Gallery must pass separately from projection: {folder}")
    require(projection["result"]["run_id"] != g["run_id"], "Projection/gallery run IDs reused")
    return {"projection": projection, "gallery": gallery}


def check_trace(path, baseline):
    trace = raw_text(path)
    for key, expected in COMMON_CHECKS.items():
        lines = [line for line in trace.splitlines() if f"{key}:" in line]
        if key == "TerminalSummaryFit" and not baseline:
            expected += 8  # Each new scene-144 capture adds its own summary-fit observation.
        require(len(lines) == expected and all(re.search(r"\bfit=true\b", line) for line in lines),
                f"Existing gallery witness count/fit changed: {key} in {path}")
    for key, expected in NEW_CHECKS.items():
        lines = [line for line in trace.splitlines() if f"{key}:" in line]
        require(len(lines) == (0 if baseline else expected), f"New witness count differs: {key} in {path}")
        require(all(re.search(r"\bfit=true\b", line) for line in lines),
                f"New terminal-facts witness failed: {key} in {path}")
    if not baseline:
        lines = [line for line in trace.splitlines() if "TerminalFactsFit:" in line]
        identities = [re.search(r"TerminalFactsFit: (\S+) dpi=(\d+) width=(\d+) theme=(\S+)", line)
                      for line in lines]
        require(all(identities) and {match.groups() for match in identities} ==
                {("144-terminal-move-mixed-facts", str(dpi), str(width), theme)
                 for dpi in (96, 144) for width in (480, 760) for theme in ("dark", "high-contrast")},
                "Scene 144 terminal-facts observation identities differ")
        transitions = [line for line in trace.splitlines() if "TerminalFactsTransition:" in line]
        identities = [re.search(r"TerminalFactsTransition: dpi=(\d+) width=(\d+) stage=(\S+)", line)
                      for line in transitions]
        require(all(identities) and {match.groups() for match in identities} ==
                {(str(dpi), "480", stage) for dpi in (96, 144) for stage in STAGES},
                "The two native displays must each witness all twelve same-HWND stages")
        refresh = [line for line in trace.splitlines() if "TerminalFactsSettingsRefresh:" in line]
        require(len(refresh) == 2 and all(re.search(r"rebuilds=1 repeats=0 stable=true\b", line)
                                          for line in refresh),
                "Settings notification must rebuild once and reuse on repeat")
        refresh_ids = [re.search(r"TerminalFactsSettingsRefresh: dpi=(\d+) width=(\d+)", line)
                       for line in refresh]
        require(all(refresh_ids) and {match.groups() for match in refresh_ids} ==
                {("96", "480"), ("144", "480")}, "Settings refresh native-display identities differ")
    else:
        require("TerminalFactsSettingsRefresh:" not in trace, "Baseline unexpectedly has new settings witness")


def check_localization(candidate):
    result = read_json(candidate / "localization-results.json")
    require(result == {"TotalCount": 9, "PassedCount": 9, "FailedCount": 0, "SkippedCount": 0},
            f"Structured localization contract result differs: {candidate}")
    log = raw_text(candidate / "localization-tests.log")
    require(re.search(r"(?m)^Passed:\s*9\s+Failed:\s*0\s+Skipped:\s*0\s+Pending:\s*0\s+Inconclusive:\s*0\s*$", log),
            f"Human localization transcript disagrees with structured result: {candidate}")


def check_intermediate(final_receipt, final_manifest):
    """Retain a concise v1 receipt trail, without treating v1 as final qualification."""
    receipt, head = check_build(INTERMEDIATE, False)
    inputs = manifest(INTERMEDIATE, False)
    runs = check_runs(INTERMEDIATE, False)
    log = raw_text(INTERMEDIATE / "localization-tests.log")
    require(re.search(r"(?m)^\s*TotalCount\s*:\s*9\s*$", log)
            and re.search(r"(?m)^\s*PassedCount\s*:\s*8\s*$", log)
            and re.search(r"(?m)^\s*FailedCount\s*:\s*1\s*$", log)
            and "IDS_FMT_FILEOPS_TERMINAL_FACT_AXIS" in log,
            "Intermediate localization 8/9 result drifted")
    require(receipt["git_head"] == final_receipt["git_head"]
            and receipt["source_snapshot_id"] != final_receipt["source_snapshot_id"],
            "Intermediate/final source provenance differs unexpectedly")
    for key in ("repo_identity", "project_graph_digest", "runtime_dependency_digest",
                "toolchain_identity_digest", "vcpkg_identity_digest"):
        require(receipt[key] == final_receipt[key], f"Intermediate/final {key} drift")
    require(set(inputs) == set(final_manifest) and
            {path for path in inputs if inputs[path] != final_manifest[path]} ==
            {"RedSalamander/FolderWindow.FileOperations.Popup.cpp",
             "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallery.cpp"},
            "V2 must change only Popup.cpp and the gallery witness from v1")
    return {"receipt": receipt, "head": head, "runs": runs}


def add_copy(plan, source, target):
    plan.append((file_at(source), target))


def generated_bytes(text):
    # Generated prose only. Raw logs, traces, receipts, and patches retain their bytes.
    return text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")


def payload_budget(copies, generated, folder):
    payload = [(target, source.stat().st_size) for source, target in copies if target.is_relative_to(folder)]
    payload += [(target, len(data)) for target, data in generated.items() if target.is_relative_to(folder)]
    require(len({str(target) for target, _ in payload}) == len(payload), f"Duplicate payload target in {folder}")
    require(all(size <= 2 * 1024 * 1024 for _, size in payload), f"File exceeds 2 MiB in {folder}")
    require(sum(size for _, size in payload) + 65536 <= 5 * 1024 * 1024,
            f"Packet exceeds 5 MiB with SHA reserve: {folder}")


def preflight(candidate, packet, gallery):
    require(not packet.exists() and not gallery.exists(), "Output exists; inspect before retrying")
    spec = importlib.util.spec_from_file_location("terminal_facts_comparator", file_at(HERE / "compare-terminal-facts-release.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fresh_layout, fresh_capture = module.compare(candidate)
    require(read_json(candidate / "layout-comparison.json") == fresh_layout and
            read_json(candidate / "capture-comparison.json") == fresh_capture,
            "Saved comparison does not match fresh raw trace/PNG recalculation")

    folders = {"baseline": BASELINE, "candidate": candidate}
    receipts, heads, manifests, runs = {}, {}, {}, {}
    for label, folder in folders.items():
        receipts[label], heads[label] = check_build(folder, label == "baseline")
        manifests[label] = manifest(folder, label == "baseline")
        runs[label] = check_runs(folder, label == "baseline")
        check_trace(runs[label]["gallery"]["root"] / "fileops/trace.txt", label == "baseline")
    check_localization(candidate)
    require(heads["baseline"] == heads["candidate"], "Primary HEAD changed between builds; review source provenance")
    require(receipts["baseline"]["git_head"] == receipts["candidate"]["git_head"],
            "Isolated build HEAD changed; review before calling matched")
    require(receipts["baseline"]["source_snapshot_id"] != receipts["candidate"]["source_snapshot_id"],
            "Source snapshots unexpectedly identical")
    for key in ("repo_identity", "project_graph_digest", "runtime_dependency_digest",
                "toolchain_identity_digest", "vcpkg_identity_digest"):
        require(receipts["baseline"][key] == receipts["candidate"][key],
                f"Dependency/toolchain identity drift: {key}")
    old, new = manifests["baseline"], manifests["candidate"]
    require(set(new) - set(old) == ADDED and not (set(old) - set(new)) and
            {name for name in old if old[name] != new[name]} == CHANGED,
            "Source delta differs from the reviewed four changed/eight added inputs")
    intermediate = check_intermediate(receipts["candidate"], manifests["candidate"])
    require(intermediate["head"] == heads["candidate"],
            "Intermediate and final primary HEADs differ")

    selected = []
    for label in folders:
        cap_dir = runs[label]["gallery"]["root"] / "fileops/visual-gallery"
        for scene in FOCUS_COMMON:
            for dpi, display in ((144, 1), (96, 2)):
                matches = sorted(cap_dir.glob(f"{scene}__*__480dip__fr-FR__display{display}__{dpi}dpi.png"))
                # Scene 31 is high contrast; the two Move scenes are dark.
                expected_theme = "high-contrast" if scene.startswith("31-") else "dark"
                matches = [p for p in matches if f"__{expected_theme}__" in p.name]
                require(len(matches) == 1, f"Missing focused {label} capture: {scene}/{dpi}")
                selected.append((label, matches[0]))
    candidate_gallery = runs["candidate"]["gallery"]["root"] / "fileops/visual-gallery"
    selected += [("candidate", candidate_gallery / row["file"])
                 for row in fresh_capture["candidateOnlyImages"]]
    require(len(selected) == 20, "Expected 12 paired common and 8 candidate-only actual captures")

    copies = []
    for label, folder in folders.items():
        top = ["build.log", "build-exit.txt", "build-receipt.json", "candidate-inputs.json",
               "primary-head.txt", "build-and-capture.ps1"]
        if label == "baseline":
            top += ["candidate-inputs-unique.json", "manifest-note.md", "baseline-witness.patch",
                    "witness-review.md", "capture-baseline-gallery.ps1"]
        else:
            top += ["localization-tests.log", "localization-results.json", "settings-followup.patch", "capture-candidate.ps1",
                    "capture-candidate-gallery.ps1"]
        for name in top:
            add_copy(copies, folder / name, packet / "raw" / label / name)
        for kind, run in runs[label].items():
            run_dir = packet / "raw" / label / kind
            for path in (run["log"], run["plan"], run["exit"]):
                add_copy(copies, path, run_dir / path.name)
            root = run["root"]
            for name in ("run-all-tests-results.json", "fileops/results.json", "fileops/trace.txt"):
                add_copy(copies, root / name, run_dir / name)
        cap_dir = runs[label]["gallery"]["root"] / "fileops/visual-gallery"
        for name in ("scenarios.tsv", "displays.tsv"):
            add_copy(copies, cap_dir / name, gallery / "captures" / label / name)
    for name in ("layout-comparison.json", "capture-comparison.json"):
        add_copy(copies, candidate / name, packet / "raw" / name)
    for name in ("build-receipt.json", "candidate-inputs.json", "localization-tests.log",
                 "production-and-witness.patch"):
        add_copy(copies, INTERMEDIATE / name, packet / "raw/intermediate-v1" / name)
    for kind, run in intermediate["runs"].items():
        add_copy(copies, run["root"] / "run-all-tests-results.json",
                 packet / "raw/intermediate-v1" / kind / "run-all-tests-results.json")
    add_copy(copies, candidate / "capture-comparison.json", gallery / "baseline-release-comparison.json")
    add_copy(copies, HERE / "compare-terminal-facts-release.py", packet / "raw/compare-terminal-facts-release.py")
    add_copy(copies, HERE / "curate-terminal-facts-release.py", packet / "raw/curate-terminal-facts-release.py")
    for label, path in selected:
        add_copy(copies, path, gallery / "captures" / label / path.name)

    time_changes = [row["usPerRedraw"]["changePercent"]
                    for row in fresh_layout["pairedMedians"].values()]
    private_changes = [row["privateBytes"]["delta"] / (1024 * 1024)
                       for row in fresh_layout["pairedMedians"].values()]
    require(all(x is not None and math.isfinite(x) for x in time_changes + private_changes),
            "Non-finite paired metric")
    narrative = (
        "# Terminal facts: matched Release evidence\n\n"
        f"Baseline projection `{runs['baseline']['projection']['result']['run_id']}` fails its count projection "
        "(2 passed, 1 failed) by design; its separate gallery "
        f"`{runs['baseline']['gallery']['result']['run_id']}` passes with 508 captures. Candidate projection "
        f"`{runs['candidate']['projection']['result']['run_id']}` passes 3/3, and gallery "
        f"`{runs['candidate']['gallery']['result']['run_id']}` passes with 516 captures. "
        "Both test runs per build have clean disk audits and receipt-bound logs. The baseline build "
        "had two C5258 test-witness capture warnings and no errors; candidate diagnostics must be zero.\n\n"
        f"The comparator checks 508 common PNG pairs by SHA-256, dimensions, and RGBA difference bounds "
        f"({fresh_capture['commonChangedImages']} differ). Eight scene-144 PNGs exist only in the candidate. "
        "The retained 24 shared redraw fixtures have five trace-bound rounds each; their paired median "
        f"redraw-time changes range {min(time_changes):+.2f}% to {max(time_changes):+.2f}%, and private-byte "
        f"changes range {min(private_changes):+.2f} to {max(private_changes):+.2f} MiB. "
        "These are in-process observations, not presented FPS, heap attribution, GPU memory, or final performance acceptance.\n\n"
        "The intermediate v1 candidate passed projection and gallery, but its 8/9 localization-contract result "
        "and missing settings-cache invalidation prevent it from being treated as final. Its build and run "
        "receipts, manifest, localization result, and initial production patch are retained concisely; the "
        "v2 settings follow-up patch and final receipts are retained separately. V2 adds twelve same-HWND "
        "transition stages per display, including a settings notification that rebuilds once and reuses "
        "the retained paragraph on the repeated paint.\n\n"
        "The original baseline manifest has 18 rows for 17 distinct sources because SelfTest.cpp is repeated "
        "with the same hash; the corrected unique view and original metadata are both retained. "
        "The final source delta is four changed existing files and eight added localization/test inputs; "
        "fixture bytes are therefore not identical as a whole. Both primary and isolated build HEADs and "
        "dependency digests are explicitly compared by the curator.\n\n"
        "The focused gallery has 12 shared actual captures (Move source kept, source unknown, and the "
        "completed group at both native displays) plus all eight new scene-144 captures. This packet covers "
        "Release only; Debug, PR, Full, UIA, native-user interaction, and final consumer acceptance remain open.\n"
    )
    cards = "\n".join(
        f'<article><h2>{html.escape(label + ": " + path.name)}</h2>'
        f'<img loading="lazy" src="captures/{label}/{html.escape(path.name, quote=True)}" '
        f'alt="{html.escape(label + ": " + path.name, quote=True)}"></article>'
        for label, path in selected)
    generated = {
        packet / "README.md": generated_bytes(narrative +
            "\n[Focused actual captures](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-terminal-facts/README.md).\n"),
        gallery / "README.md": generated_bytes(narrative +
            "\n[Gallery](gallery.html) · [All capture comparisons](baseline-release-comparison.json).\n"),
        gallery / "gallery.html": generated_bytes(
            '<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>Terminal facts Release captures</title><style>body{background:#15181d;color:#eee;font:16px system-ui;margin:24px}'
            'main{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:20px}'
            'article{padding:12px;background:#222831}h2{font-size:14px;overflow-wrap:anywhere}'
            'img{width:100%;height:650px;object-fit:contain}a{color:#acdaff}</style>'
            '<h1>Actual terminal facts captures</h1><p>Matched Release and eight candidate-only scene-144 captures.</p>'
            '<p><a href="README.md">Evidence and limits</a></p><main>' + cards + '</main></html>\n'),
        packet / ".gitattributes": b"raw/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
        gallery / ".gitattributes": b"captures/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
    }
    for output in (packet, gallery):
        payload_budget(copies, generated, output)
    return copies, generated, fresh_capture, fresh_layout


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=DEFAULT_CANDIDATE)
    parser.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    parser.add_argument("--gallery", type=Path, default=DEFAULT_GALLERY)
    args = parser.parse_args()
    copies, generated, captures, layout = preflight(args.candidate, args.packet, args.gallery)
    # No output directories are created until every source, receipt, trace, image and byte guard passes.
    for source, target in copies:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
        require(digest(source) == digest(target), f"Copy drift: {target}")
    for target, content in generated.items():
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(content)
    for output in (args.packet, args.gallery):
        paths = sorted(path for path in output.rglob("*") if path.is_file() and path.name != "SHA256SUMS.txt")
        manifest = "".join(f"{digest(path)}  {path.relative_to(output).as_posix()}\n" for path in paths)
        (output / "SHA256SUMS.txt").write_bytes(manifest.encode("utf-8"))
        require(all(path.stat().st_size <= 2 * 1024 * 1024 for path in output.rglob("*") if path.is_file())
                and sum(path.stat().st_size for path in output.rglob("*") if path.is_file()) <= 5 * 1024 * 1024,
                f"Post-write archive budget failed: {output}")
    print(json.dumps({"commonPairs": captures["commonPairs"], "changed": captures["commonChangedImages"],
                      "candidateOnly": captures["candidateOnly"], "pairedRedrawFixtures": layout["fixtureCount"],
                      "focusedActualCaptures": 20}, indent=2))


if __name__ == "__main__":
    main()

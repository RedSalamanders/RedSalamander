"""Draft only: curate the final semantic-detail Release pair after v6 archival.

Never upgrades retained v1/v2 compile failures or v3/v4/v5 gallery failures into a final pass.
The author has not executed this script or written into either repository.
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
BASELINE = EVIDENCE / "direct-issues-qualified-20260923"
V1 = EVIDENCE / "semantic-detail-qualified-20260923"
V2 = EVIDENCE / "semantic-detail-qualified-v2-20260923"
V3 = EVIDENCE / "semantic-detail-qualified-v3-20260923"
V4 = EVIDENCE / "semantic-detail-qualified-v4-20260923"
V5 = EVIDENCE / "semantic-detail-qualified-v5-20260923"
V6 = EVIDENCE / "semantic-detail-qualified-v6-20260923"
PACKET = ROOT / "Specs/TestRuns/4cb089111a23/Validation/2026-09-23_semantic-detail"
GALLERY = ROOT / "Specs/Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-semantic-detail"
BASELINE_HEAD = "bf5506bd00babf55e9d9ef80f0af854b0903199c"
ISOLATED_HEAD = "ccebd5af68eb45f2e9deed64a7063780e71eb1ce"
DEPENDENCY_KEYS = ("repo_identity", "project_graph_digest", "runtime_dependency_digest",
                   "toolchain_identity_digest", "vcpkg_identity_digest")
FOCUSED_COMMON = ("125-issues-failed-replacement", "129-issues-reappeared-identity")
RECONSTRUCTION_NOTE_SHA = "a7bd86c546557549a9ecafcfe1445631d6c6eb2b68c66fd583d92991791cbd82"
VISUAL_MD_SHA = "17b7cd89b30ad89e5b741baf989f64e37766eb5808ddf2be0e4013f0f756a0db"
VISUAL_JSON_SHA = "0ce3f72b168f19aa7bdff130521a55626f3b318014bffd9d055caca05c555598"


def require(value, explanation):
    if not value:
        raise RuntimeError(explanation)


def raw_text(path):
    raw = path.read_bytes()
    return raw.decode("utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def read_json(path):
    return json.loads(raw_text(path))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def frozen_manifest(folder):
    rows = read_json(folder / "candidate-inputs.json")
    require(len(rows) == 29, f"Expected 29 frozen source inputs in {folder}")
    result = {}
    for row in rows:
        name, expected = row["path"], row["sha256"].lower()
        require(name not in result and re.fullmatch(r"[0-9a-f]{64}", expected), f"Bad source row: {name}")
        require(sha(folder / "inputs" / name) == expected, f"Frozen source changed: {folder}/{name}")
        result[name] = expected
    return result


def exit_value(path):
    return int(raw_text(path).strip())


def build(folder, expected_success):
    code = exit_value(folder / "build-exit.txt")
    log = raw_text(folder / "build.log")
    head_path = folder / "primary-head.txt"
    head = raw_text(head_path).strip() if head_path.is_file() else None
    require(head is None or re.fullmatch(r"[0-9a-f]{40}", head), f"Invalid primary HEAD: {folder}")
    if expected_success:
        require(code == 0, f"Required Release build failed: {folder}")
        receipt = read_json(folder / "build-receipt.json")
        require(receipt["configuration"] == "Release" and receipt["platform"] == "x64"
                and receipt["tests_enabled"] is True and receipt["exit_code"] == 0,
                f"Wrong build profile/outcome: {folder}")
        for key, size in (("git_head", 40), ("source_snapshot_id", 64), ("receipt_id", 64)):
            require(re.fullmatch(rf"[0-9a-f]{{{size}}}", receipt[key]), f"Malformed {key}: {folder}")
        diagnostic = re.search(r"Diagnostics:\s*(\d+) warning\(s\),\s*(\d+) error\(s\)", log)
        require(diagnostic and diagnostic.groups() == ("0", "0"), f"Build diagnostics in {folder}")
        return receipt, head, code
    if code == 0:
        receipt = read_json(folder / "build-receipt.json")
        require(receipt["exit_code"] == 0, f"V2 successful build receipt differs: {folder}")
    else:
        require(not (folder / "build-receipt.json").exists(), "Failed intermediate unexpectedly has successful receipt")
        receipt = None
    return receipt, head, code


def gallery(folder, build_receipt, required):
    matches = []
    for root in folder.glob("*-runner-artifacts/artifacts/selftest/last_run"):
        path = root / "run-all-tests-results.json"
        if path.is_file() and read_json(path).get("case_filter") == "FileOps_VisualGallery":
            matches.append((root, read_json(path)))
    if not required and not matches:
        return None
    require(len(matches) == 1, f"Expected exactly one archived gallery in {folder}: {matches}")
    root, result = matches[0]
    require(result["schema"] == "red-salamander.run-all-tests.v1"
            and result["suite"] == "FileOps" and result["platform"] == "x64"
            and result["configuration"] == "Release"
            and result["test_sandbox_audit"]["is_clean"] is True
            and result["test_sandbox_audit"]["issue_count"] == 0,
            f"Wrong gallery profile/disk audit: {folder}")
    cases = [row for suite in result["suites"] for row in suite["cases"]]
    selected = [row for row in cases if row["name"] == "FileOps_VisualGallery"]
    require(len(selected) == 1 and result["total"] == 1, f"Gallery case mismatch: {folder}")
    logs = [path for path in folder.glob("*.log")
            if "FileOps_VisualGallery" in raw_text(path) and result["run_id"] in raw_text(path)]
    require(len(logs) == 1, f"Gallery run log ambiguity: {folder}")
    log = logs[0]
    plan = read_json(folder / f"{log.stem}-plan.json")
    require(plan["Id"] == "selftest.file-operations" and plan["RequiresInteractiveDesktop"] is False,
            f"Gallery unexpectedly requires desktop: {folder}")
    require(build_receipt is not None and f"verified receipt {build_receipt['receipt_id']}" in raw_text(log),
            f"Gallery log does not bind build receipt: {folder}")
    require(exit_value(folder / f"{log.stem}-exit.txt") == result["exit_code"], f"Gallery exit differs: {folder}")
    nested = read_json(root / "fileops/results.json")
    require(nested["passed"] == result["passed"] and nested["failed"] == result["failed"],
            f"Nested gallery count differs: {folder}")
    return {"root": root, "result": result, "case": selected[0], "log": log}


def check_retained_issues_trace(root):
    lines = raw_text(root / "fileops/trace.txt").splitlines()
    for marker, count in (("DirectIssuesActionFit:", 16), ("DirectIssuesTransition:", 20)):
        rows = [line for line in lines if marker in line]
        require(len(rows) == count and all(re.search(r"\bfit=true\b", line) for line in rows),
                f"Retained Issues witness changed: {marker} in {root}")


def target_name(path):
    return path.name + ".receipt.txt" if path.suffix.lower() == ".json" else path.name


def add_copy(copies, source, destination):
    require(source.is_file(), f"Missing required raw evidence: {source}")
    copies.append((source, destination))


def add_folder_files(copies, folder, names, destination):
    for name in names:
        source = folder / name
        add_copy(copies, source, destination / target_name(source))


def generated(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")


def budget(copies, content, root):
    entries = [(target, source.stat().st_size) for source, target in copies if target.is_relative_to(root)]
    entries += [(target, len(data)) for target, data in content.items() if target.is_relative_to(root)]
    require(len({str(path) for path, _ in entries}) == len(entries), f"Duplicate archive path: {root}")
    require(all(size <= 2*1024*1024 for _, size in entries), f"File over 2 MiB: {root}")
    require(sum(size for _, size in entries) + 65536 <= 5*1024*1024, f"Packet over 5 MiB: {root}")


def preflight(candidate, packet, gallery_target):
    require(not packet.exists() and not gallery_target.exists(), "Output exists; inspect before retry")
    module_path = HERE / "compare-semantic-detail-release-draft.py"
    spec = importlib.util.spec_from_file_location("semantic_detail_comparator", module_path)
    comparator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(comparator)
    fresh_layout, fresh_capture = comparator.compare(candidate)
    require(read_json(candidate / "layout-comparison.json") == fresh_layout
            and read_json(candidate / "capture-comparison.json") == fresh_capture,
            "Saved comparison differs from raw PNG/trace recalculation")
    note = HERE / "semantic-detail-reconstruction-note-sol.md"
    visual_md = HERE / "semantic-detail-v6-visual-review-sol.md"
    visual_json = HERE / "semantic-detail-v6-visual-review-sol.json"
    require(sha(note) == RECONSTRUCTION_NOTE_SHA and sha(visual_md) == VISUAL_MD_SHA
            and sha(visual_json) == VISUAL_JSON_SHA,
            "Reviewed reconstruction note or visual-review artifact changed")
    review = read_json(visual_json)
    require(review["candidate"] == candidate.name
            and review["captureComparisonSha256"] == sha(candidate / "capture-comparison.json")
            and review["reviewedChangedCommon"] == [row for row in fresh_capture["common"] if row["visualChanged"]]
            and review["reviewedNew"] == fresh_capture["new"],
            "Visual review no longer matches actual 20 changed/six new captures")

    folders = {"baseline": BASELINE, "v1-failed-compile": V1, "v2-failed-compile": V2,
               "v3-failed-gallery": V3, "v4-failed-gallery": V4,
               "v5-failed-gallery": V5, "candidate-v6": candidate}
    manifests = {label: frozen_manifest(folder) for label, folder in folders.items()}
    require(set(manifests["baseline"]) == set(manifests["candidate-v6"]), "Final source path set differs")
    surfaces = "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallerySurfaces.cpp"
    for before, after in (("v3-failed-gallery", "v4-failed-gallery"),
                          ("v4-failed-gallery", "v5-failed-gallery"),
                          ("v5-failed-gallery", "candidate-v6")):
        old, new = manifests[before], manifests[after]
        require(set(old) == set(new) and
                {name for name in old if old[name] != new[name]} == {surfaces},
                f"{before} to {after} changed more than the gallery witness source")
    b_receipt, b_head, _ = build(BASELINE, True)
    c_receipt, c_head, _ = build(candidate, True)
    v3_receipt, v3_head, _ = build(V3, True)
    v4_receipt, v4_head, _ = build(V4, True)
    v5_receipt, v5_head, _ = build(V5, True)
    require(v4_head is None and v5_head is None, "v4/v5 primary HEAD receipts changed; review provenance")
    v2_receipt, v2_head, v2_exit = build(V2, False)
    require(exit_value(V1 / "build-exit.txt") == 1 and not (V1 / "build-receipt.json").exists(),
            "Initial compile-failure record changed")
    v1_log = raw_text(V1 / "build.log")
    require("error C2039: 'successText'" in v1_log and "warning C5246" in v1_log
            and "Diagnostics: 16 warning(s), 1 error(s)" in v1_log,
            "Initial compile-failure diagnostic changed")
    require(v2_exit == 1 and v2_receipt is None
            and "error C2664" in raw_text(V2 / "build.log")
            and "Diagnostics: 2 warning(s), 5 error(s)" in raw_text(V2 / "build.log"),
            "Second compile-failure diagnostic changed")
    require(b_head == BASELINE_HEAD and b_receipt["git_head"] == c_receipt["git_head"] == ISOLATED_HEAD,
            "Reviewed baseline/isolated build HEAD changed")
    require(all(receipt["git_head"] == ISOLATED_HEAD
                for receipt in (v3_receipt, v4_receipt, v5_receipt)),
            "Intermediate isolated build HEAD drifted")
    require(c_head == "33b42f40fd6e8c1d037a748e335781e59210513b",
            "Candidate primary HEAD changed")
    for key in DEPENDENCY_KEYS:
        require(b_receipt[key] == c_receipt[key], f"Build dependency/toolchain drift: {key}")
        require(all(receipt[key] == c_receipt[key]
                    for receipt in (v3_receipt, v4_receipt, v5_receipt)),
                f"Intermediate dependency/toolchain drift: {key}")
    require(b_receipt["source_snapshot_id"] != c_receipt["source_snapshot_id"],
            "Final source snapshots unexpectedly identical")

    b_run = gallery(BASELINE, b_receipt, True)
    c_run = gallery(candidate, c_receipt, True)
    v3_run = gallery(V3, v3_receipt, True)
    v4_run = gallery(V4, v4_receipt, True)
    v5_run = gallery(V5, v5_receipt, True)
    require((v3_run["result"]["passed"], v3_run["result"]["failed"],
             v3_run["result"]["exit_code"], v3_run["case"]["status"])
            == (0, 1, 1, "failed")
            and "Scenario capture failed" in v3_run["case"].get("reason", "")
            and "123-issues-legacy-missing-facts" in raw_text(v3_run["root"] / "fileops/trace.txt")
            and "matched=0 required=1" in raw_text(v3_run["root"] / "fileops/trace.txt"),
            "Retained v3 gallery failure changed")
    for label, run, filename, witness in (
            ("v4", v4_run, "129b-issues-confirmed-publication__dark__fr-FR__display1__144dpi.png",
             "matched=0 required=4"),
            ("v5", v5_run, "129b-issues-confirmed-publication__dark__fr-FR__display2__96dpi.png",
             "matched=1 required=4")):
        require((run["result"]["passed"], run["result"]["failed"],
                 run["result"]["exit_code"], run["case"]["status"]) == (0, 1, 1, "failed")
                and "Scenario capture failed" in run["case"].get("reason", "")
                and filename in raw_text(run["root"] / "fileops/trace.txt")
                and witness in raw_text(run["root"] / "fileops/trace.txt"),
                f"Retained {label} gallery failure changed")
    require(v5_receipt["receipt_id"] == "32b84b3d3087ddc905283281609c38d99a1b6b1cb273c1a8389e6e0943b23eb3"
            and v5_run["result"]["run_id"] == "20260923T201341Z-86572-05f42345b2d54a0cb9d522fdd0640f94",
            "V5 retained receipt identity changed")
    for label, run in (("baseline", b_run), ("candidate-v6", c_run)):
        result = run["result"]
        require((result["passed"], result["failed"], result["exit_code"], run["case"]["status"])
                == (1, 0, 0, "passed"), f"Final {label} gallery did not pass")
        check_retained_issues_trace(run["root"])
    require(len(list((b_run["root"] / "fileops/visual-gallery").glob("*.png"))) == 532
            and len(list((c_run["root"] / "fileops/visual-gallery").glob("*.png"))) == 538,
            "Actual capture counts changed")
    copies = []
    common_top = ("build.log", "build-exit.txt", "candidate-inputs.json",
                  "build-and-capture.ps1", "capture-candidate-gallery.ps1")
    for label, folder in folders.items():
        names = list(common_top)
        if (folder / "primary-head.txt").is_file():
            names.append("primary-head.txt")
        if label in ("baseline", "v3-failed-gallery", "v4-failed-gallery",
                     "v5-failed-gallery", "candidate-v6"):
            names.append("build-receipt.json")
        if label in ("v4-failed-gallery", "v5-failed-gallery", "candidate-v6"):
            names += ["candidate-overlay.patch", "correction.txt"]
        else:
            names.append("production-and-witness.patch")
        if label.startswith("v") and (folder / "baseline-note.md").is_file():
            names.append("baseline-note.md")
        add_folder_files(copies, folder, names, packet / "raw" / label)
    for label, folder, run in (("baseline", BASELINE, b_run), ("v3-failed-gallery", V3, v3_run),
                               ("v4-failed-gallery", V4, v4_run),
                               ("v5-failed-gallery", V5, v5_run),
                               ("candidate-v6", candidate, c_run)):
        if run is None:
            continue
        target = packet / "raw" / label / "gallery"
        add_folder_files(copies, folder,
                         (run["log"].name, f"{run['log'].stem}-plan.json", f"{run['log'].stem}-exit.txt"), target)
        for name in ("run-all-tests-results.json", "fileops/results.json", "fileops/trace.txt"):
            source = run["root"] / name
            add_copy(copies, source, target / (name.replace("/", "__") + ".receipt.txt" if name.endswith(".json") else name.replace("/", "__")))
    for label, run, name in (
            ("v3-failed-gallery", v3_run,
             "123-issues-legacy-missing-facts__dark__fr-FR__display1__144dpi.png"),
            ("v4-failed-gallery", v4_run,
             "129b-issues-confirmed-publication__dark__fr-FR__display1__144dpi.png"),
            ("v5-failed-gallery", v5_run,
             "129b-issues-confirmed-publication__dark__fr-FR__display2__96dpi.png")):
        failed_capture = run["root"] / "fileops/visual-gallery" / name
        add_copy(copies, failed_capture, packet / "raw" / label / failed_capture.name)
    add_folder_files(copies, candidate, ("layout-comparison.json", "capture-comparison.json"), packet / "raw")
    for name in ("compare-semantic-detail-release-draft.py", "curate-semantic-detail-release-draft.py"):
        add_copy(copies, HERE / name, packet / "raw" / name)
    add_copy(copies, note, packet / "raw" / "post-run-reconstruction.md")
    # Keep the browsable review Markdown only in the gallery, beside its linked
    # JSON. A second byte-exact copy under packet/raw would have a broken link
    # because non-benchmark JSON there must carry the .receipt.txt suffix.
    add_copy(copies, visual_json, packet / "raw" / "visual-review-sol.json.receipt.txt")
    add_copy(copies, visual_md, gallery_target / "visual-review-sol.md")
    # Preserve the original filename beside the byte-exact Markdown review:
    # its relative JSON link must work in the browsable gallery. The packet's
    # raw non-benchmark JSON keeps its .receipt.txt archival name.
    add_copy(copies, visual_json, gallery_target / "semantic-detail-v6-visual-review-sol.json")

    selected = []
    for label, run in (("baseline", b_run), ("candidate-v6", c_run)):
        folder = run["root"] / "fileops/visual-gallery"
        for scene in FOCUSED_COMMON:
            for display, dpi in ((1, 144), (2, 96)):
                name = f"{scene}__dark__fr-FR__display{display}__{dpi}dpi.png"
                add_copy(copies, folder / name, gallery_target / "captures" / label / name)
                selected.append((label, name))
    folder = c_run["root"] / "fileops/visual-gallery"
    for name in fresh_capture["newSceneNames"]:
        add_copy(copies, folder / name, gallery_target / "captures" / "candidate-v6" / name)
        selected.append(("candidate-v6", name))
    for label, run in (("baseline", b_run), ("candidate-v6", c_run)):
        source = run["root"] / "fileops/visual-gallery/scenarios.tsv"
        add_copy(copies, source, gallery_target / "captures" / label / "scenarios.tsv")
    add_copy(copies, candidate / "capture-comparison.json", gallery_target / "capture-comparison.json")

    changes = [row["usPerRedraw"]["changePercent"] for row in fresh_layout["pairedMedians"].values()]
    private = [row["privateBytes"]["delta"] / (1024*1024)
               for row in fresh_layout["pairedMedians"].values()]
    require(all(value is not None and math.isfinite(value) for value in changes + private),
            "Invalid paired redraw measurement")
    narrative = (
        "# Semantic Issues detail: matched Release evidence\n\n"
        f"Baseline gallery `{b_run['result']['run_id']}` and final v6 gallery `{c_run['result']['run_id']}` "
        "each pass their filtered nonactivating FileOps_VisualGallery case with clean sandbox audits. "
        "The retained direct-Issues baseline has 532 actual PNGs; v6 has 538. Exactly 532 names are paired "
        "by SHA-256, dimensions and RGBA pixels, while six new 129b/129c/129d captures are recorded separately. "
        f"{fresh_capture['commonChangedImages']} common captures differ; "
        f"{fresh_capture['commonDimensionChanges']} change dimensions. This automated comparison does not "
        "replace independent visual review.\n\n"
        f"The shared 24 same-HWND redraw fixtures use byte-identical `VisualGallery.cpp` source "
        f"`{fresh_layout['sourceSha256']}` and five eight-frame synchronous samples each. "
        f"Paired median redraw time ranges {min(changes):+.2f}% to {max(changes):+.2f}%; "
        f"private-byte differences range {min(private):+.2f} to {max(private):+.2f} MiB. "
        "These are short in-process diagnostics, not presented FPS, allocator attribution, or whole-gallery "
        "memory acceptance. The ordinary paired performance and consumer gates remain open.\n\n"
        "The v1 compile failure (C2039 `successText`, 16 C5246 warnings) and v2 compile failure "
        "(C2664 conversion, two warnings/five errors) remain separate. V3 built cleanly but failed scene 123 "
        f"warning-icon probe (zero/one; run `{v3_run['result']['run_id']}`). V4 built cleanly but failed "
        f"129b display 1 green-check probe (zero/four; run `{v4_run['result']['run_id']}`). "
        f"V5 built cleanly but failed 129b display 2 green-check probe (one/four; run `{v5_run['result']['run_id']}`). "
        "The three failed captures, traces, build receipts and gallery receipts remain separate in raw/. "
        "V4/v5 frozen input copies were reconstructed after their runs from v6 bytes and retained manifest hashes; "
        "the post-run repair and its limits are documented in raw/post-run-reconstruction.md. "
        "The new 129c Delete fixture retains a Copy destination, and 129d Rename uses a cross-folder same-leaf "
        "destination. Their synthetic rows witness typed result/icon mapping, not executed filesystem operations "
        "or representative final documentation capture data; a later qualified gallery must correct both paths. "
        "Each v3-to-v6 correction changes only the gallery-surfaces witness source; no intermediate outcome is upgraded to a pass. "
        f"Recorded primary HEADs are baseline `{b_head}` and final `{c_head}`; v4/v5 did not record "
        "`primary-head.txt`, and their isolated build-receipt `git_head` is not a primary HEAD. "
        f"Both baseline/final isolated build receipts use `{ISOLATED_HEAD}` with equal dependency/toolchain "
        "digests and distinct source snapshots. "
        "Frozen input manifests bind the exact four-file source change. Debug, Commands, UIA, interactive input, "
        "PR, Full, and final performance qualification are separate.\n"
    )
    cards = "".join(
        f'<article><h2>{html.escape(label + ": " + name)}</h2>'
        f'<img loading="lazy" src="captures/{html.escape(label)}/{html.escape(name, quote=True)}" '
        f'alt="{html.escape(label + ": " + name, quote=True)}"></article>'
        for label, name in selected)
    content = {
        packet / "README.md": generated(narrative +
            "\n[Focused actual captures](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-semantic-detail/README.md).\n"),
        gallery_target / "README.md": generated(narrative +
            "\n[Gallery](gallery.html) · [Visual review](visual-review-sol.md) · "
            "[All capture comparisons](capture-comparison.json).\n"),
        gallery_target / "gallery.html": generated(
            '<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>Semantic Issues detail Release captures</title><style>body{background:#15181d;color:#eee;font:16px system-ui;margin:24px}'
            'main{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:20px}'
            'article{padding:12px;background:#222831}h2{font-size:14px;overflow-wrap:anywhere}'
            'img{width:100%;height:650px;object-fit:contain}a{color:#acdaff}</style>'
            '<h1>Actual semantic Issues detail captures</h1><p>Fourteen focused originals; all common pairs are indexed in JSON.</p>'
            '<p><a href="README.md">Evidence and limits</a></p><main>' + cards + '</main></html>\n'),
        packet / ".gitattributes": b"raw/** -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
        gallery_target / ".gitattributes": b"captures/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n",
    }
    for output in (packet, gallery_target):
        budget(copies, content, output)
    return copies, content, fresh_capture, fresh_layout


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=V6)
    parser.add_argument("--packet", type=Path, default=PACKET)
    parser.add_argument("--gallery", type=Path, default=GALLERY)
    args = parser.parse_args()
    copies, content, captures, layout = preflight(args.candidate, args.packet, args.gallery)
    # All provenance, source, outcome, image and size checks occur before any output write.
    for source, target in copies:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
        require(sha(source) == sha(target), f"Copy drift: {target}")
    for target, data in content.items():
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    for root in (args.packet, args.gallery):
        paths = sorted(path for path in root.rglob("*") if path.is_file() and path.name != "SHA256SUMS.txt")
        payload = "".join(f"{sha(path)}  {path.relative_to(root).as_posix()}\n" for path in paths)
        (root / "SHA256SUMS.txt").write_bytes(payload.encode("utf-8"))
        require(all(path.stat().st_size <= 2*1024*1024 for path in root.rglob("*") if path.is_file())
                and sum(path.stat().st_size for path in root.rglob("*") if path.is_file()) <= 5*1024*1024,
                f"Post-write archive budget failed: {root}")
    print(json.dumps({"commonPairs": captures["commonPairs"], "changed": captures["commonChangedImages"],
                      "new": captures["newCandidateImages"], "redrawFixtures": layout["fixtureCount"]}, indent=2))


if __name__ == "__main__":
    main()

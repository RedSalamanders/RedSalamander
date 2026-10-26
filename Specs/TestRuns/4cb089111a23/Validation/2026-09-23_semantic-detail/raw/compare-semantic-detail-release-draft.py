"""Draft only: pair every common Release PNG and 24 unchanged redraw fixtures.

No native execution, visual acceptance, or whole-gallery memory qualification.
The author has not executed this script against the unfinished v6 candidate.
"""

import argparse
import csv
import hashlib
import io
import json
import re
from pathlib import Path
from statistics import median

from PIL import Image, ImageChops

EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
BASELINE = EVIDENCE / "direct-issues-qualified-20260923"
CANDIDATE = EVIDENCE / "semantic-detail-qualified-v6-20260923"
REDRAW_SOURCE = "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallery.cpp"
EXPECTED_DELTA = {
    "RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp",
    "RedSalamander/FolderWindow.FileOperations.IssuesPane.h",
    "RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp",
    "RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallerySurfaces.cpp",
}
NEW_SCENES = ("129b-issues-confirmed-publication", "129c-issues-confirmed-delete", "129d-issues-confirmed-rename")
NEW_FIXTURES = dict(zip(NEW_SCENES, ("production-issues-typed-published-result",
                                     "production-issues-typed-removed-result",
                                     "production-issues-typed-renamed-result")))
REDRAW = re.compile(
    r"TerminalLayoutRedraw: (\S+) dpi=(\d+) width=(\d+) round=(\d+) frames=(\d+) "
    r"elapsedUs=(\d+) privateBytes=(\d+) workingSetBytes=(\d+) gdi=(\d+) user=(\d+)"
)
CAPTURE = re.compile(r"VisualGallery: (\S+\.png)")
METRICS = ("usPerRedraw", "privateBytes", "workingSetBytes", "gdi", "user")


def require(value, message):
    if not value:
        raise RuntimeError(message)


def raw_text(path):
    raw = path.read_bytes()
    return raw.decode("utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def manifest(folder):
    rows = json.loads(raw_text(folder / "candidate-inputs.json"))
    require(len(rows) == 29, f"Expected 29 frozen inputs in {folder}")
    result = {}
    for row in rows:
        name, expected = row["path"], row["sha256"].lower()
        require(name not in result and re.fullmatch(r"[0-9a-f]{64}", expected), f"Bad input row: {name}")
        require(sha(folder / "inputs" / name) == expected, f"Frozen input changed: {folder}/{name}")
        result[name] = expected
    return result


def gallery_root(folder):
    matches = []
    for root in folder.glob("*-runner-artifacts/artifacts/selftest/last_run"):
        result_path = root / "run-all-tests-results.json"
        if result_path.is_file() and json.loads(raw_text(result_path)).get("case_filter") == "FileOps_VisualGallery":
            matches.append(root)
    require(len(matches) == 1, f"Expected one archived FileOps gallery in {folder}: {matches}")
    return matches[0]


def capture_files(root, expected):
    files = {path.name: path for path in (root / "fileops/visual-gallery").glob("*.png")}
    require(len(files) == expected, f"Expected {expected} PNGs in {root}, found {len(files)}")
    return files


def parse_redraw(path):
    samples, pending = {}, []
    for line in raw_text(path).splitlines():
        found = REDRAW.search(line)
        if found:
            scene, dpi, width, round_, frames, elapsed, private, working, gdi, user = found.groups()
            require(int(frames) == 8, f"Expected eight synchronous redraws: {line}")
            pending.append({"scene": scene, "dpi": int(dpi), "width": int(width),
                            "round": int(round_), "frames": 8,
                            "usPerRedraw": int(elapsed) / 8, "privateBytes": int(private),
                            "workingSetBytes": int(working), "gdi": int(gdi), "user": int(user)})
        found = CAPTURE.search(line)
        if found and pending:
            name = found.group(1)
            require(name not in samples and len(pending) == 5, f"Bad redraw group for {name}")
            require([row["round"] for row in pending] == list(range(5)), f"Bad redraw rounds for {name}")
            for row in pending:
                require(name.startswith(row["scene"] + "__")
                        and f"__{row['width']}dip__" in name
                        and name.endswith(f"__{row['dpi']}dpi.png"), f"Misbound redraw sample: {name}")
            samples[name] = pending
            pending = []
    require(not pending and len(samples) == 24, f"Expected 24 complete redraw fixtures in {path}")
    return samples


def image_row(name, old, new):
    with Image.open(old) as a_file, Image.open(new) as b_file:
        a, b = a_file.convert("RGBA"), b_file.convert("RGBA")
        bounds = ImageChops.difference(a, b).getbbox(alpha_only=False) if a.size == b.size else None
        return {"file": name, "baselineSha256": sha(old), "candidateSha256": sha(new),
                "baselineDimensions": list(a.size), "candidateDimensions": list(b.size),
                "changedBounds": list(bounds) if bounds else None,
                "visualChanged": a.size != b.size or bounds is not None}


def compare(candidate=CANDIDATE):
    manifests = {"baseline": manifest(BASELINE), "candidate": manifest(candidate)}
    a, b = manifests.values()
    require(set(a) == set(b), f"Frozen source path set changed: {sorted(set(a) ^ set(b))}")
    delta = {name for name in a if a[name] != b[name]}
    require(delta == EXPECTED_DELTA, f"Source delta requires review: {sorted(delta)}")
    require(a[REDRAW_SOURCE] == b[REDRAW_SOURCE], "24-redraw fixture source changed")
    roots = {"baseline": gallery_root(BASELINE), "candidate": gallery_root(candidate)}
    files = {"baseline": capture_files(roots["baseline"], 532),
             "candidate": capture_files(roots["candidate"], 538)}
    old, new = files["baseline"], files["candidate"]
    require(set(old) <= set(new), f"Baseline PNG names removed: {sorted(set(old)-set(new))}")
    added = set(new) - set(old)
    expected = {f"{scene}__dark__fr-FR__display{display}__{dpi}dpi.png"
                for scene in NEW_SCENES for display, dpi in ((1, 144), (2, 96))}
    require(added == expected, f"Unexpected new PNG names: {sorted(added ^ expected)}")
    manifest_text = raw_text(roots["candidate"] / "fileops/visual-gallery/scenarios.tsv")
    manifest_rows = list(csv.DictReader(io.StringIO(manifest_text), delimiter="\t"))
    by_name = {row["file"]: row for row in manifest_rows}
    require(len(by_name) == len(manifest_rows) and set(by_name) == set(new), "Scenario manifest/PNG set differs")
    for name in added:
        row = by_name[name]
        require(row["scenario"] in NEW_FIXTURES and row["fixture"] == NEW_FIXTURES[row["scenario"]]
                and row["theme"] == "dark" and row["dpi"] == ("144" if "__display1__" in name else "96"),
                f"New scene provenance differs: {name}")
    common = [image_row(name, old[name], new[name]) for name in sorted(old)]
    new_rows = []
    for name in sorted(added):
        with Image.open(new[name]) as image:
            new_rows.append({"file": name, "sha256": sha(new[name]), "dimensions": list(image.size)})

    samples = {label: parse_redraw(root / "fileops/trace.txt") for label, root in roots.items()}
    require(set(samples["baseline"]) == set(samples["candidate"]), "Redraw fixture keys changed")
    paired = {}
    for name in sorted(samples["baseline"]):
        paired[name] = {}
        for metric in METRICS:
            before = median(row[metric] for row in samples["baseline"][name])
            after = median(row[metric] for row in samples["candidate"][name])
            paired[name][metric] = {"baseline": before, "candidate": after, "delta": after-before,
                                    "changePercent": (after/before-1)*100 if before else None}
    layout = {
        "scope": "24 same-source, same-HWND Release redraw fixtures: five rounds of eight synchronous paints each. Diagnostic only; not presented FPS or memory acceptance.",
        "sourcePath": REDRAW_SOURCE, "sourceSha256": a[REDRAW_SOURCE],
        "traceSha256": {label: sha(root / "fileops/trace.txt") for label, root in roots.items()},
        "fixtureCount": 24, "roundCount": 5, "framesPerRound": 8,
        "rawSamples": samples, "pairedMedians": paired,
    }
    capture = {
        "scope": "All 532 common actual application PNGs paired by exact name and RGBA; six new candidate PNGs recorded separately. Visual review remains separate.",
        "commonPairs": 532, "newCandidateImages": 6,
        "commonChangedImages": sum(row["visualChanged"] for row in common),
        "commonDimensionChanges": sum(row["baselineDimensions"] != row["candidateDimensions"] for row in common),
        "newSceneNames": sorted(added), "new": new_rows, "common": common,
    }
    return layout, capture


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=CANDIDATE)
    args = parser.parse_args()
    layout, capture = compare(args.candidate)
    for name, data in (("layout-comparison.json", layout), ("capture-comparison.json", capture)):
        path = args.candidate / name
        require(not path.exists(), f"Output exists; inspect before rerun: {path}")
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps({"commonPairs": 532, "commonChanged": capture["commonChangedImages"],
                      "newCandidate": 6, "pairedRedrawFixtures": 24}, indent=2))


if __name__ == "__main__":
    main()

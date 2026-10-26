"""Draft only. Compare frozen direct-Issues Release galleries after both runs archive.

This script does not run native tests and does not certify UIA, input, or final perf.
"""

import argparse
import hashlib
import json
import re
from pathlib import Path
from statistics import median

from PIL import Image, ImageChops

EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
BASELINE = EVIDENCE / "direct-issues-baseline-v2-20260923"
DEFAULT_CANDIDATE = EVIDENCE / "direct-issues-qualified-20260923"
PRIOR = EVIDENCE / "terminal-facts-qualified-v2-20260923"
REDRAW = re.compile(
    r"TerminalLayoutRedraw: (\S+) dpi=(\d+) width=(\d+) round=(\d+) frames=(\d+) "
    r"elapsedUs=(\d+) privateBytes=(\d+) workingSetBytes=(\d+) gdi=(\d+) user=(\d+)"
)
CAPTURE = re.compile(r"VisualGallery: (\S+\.png)")
METRICS = ("usPerRedraw", "privateBytes", "workingSetBytes", "gdi", "user")


def require(condition, explanation):
    if not condition:
        raise RuntimeError(explanation)


def read_text(path):
    raw = path.read_bytes()
    return raw.decode("utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def gallery_root(folder):
    matches = []
    for root in folder.glob("*-runner-artifacts/artifacts/selftest/last_run"):
        result = root / "run-all-tests-results.json"
        if result.is_file() and json.loads(read_text(result)).get("case_filter") == "FileOps_VisualGallery":
            matches.append(root)
    require(len(matches) == 1, f"Exactly one archived gallery required in {folder}: {matches}")
    return matches[0]


def capture_files(root):
    gallery = root / "fileops/visual-gallery"
    files = {path.name: path for path in gallery.glob("*.png")}
    require(len(files) == 532, f"Expected 532 actual gallery PNGs in {gallery}; found {len(files)}")
    return files


def parse_redraw(path):
    samples, pending = {}, []
    for line in read_text(path).splitlines():
        found = REDRAW.search(line)
        if found:
            scene, dpi, width, round_, frames, elapsed, private, working, gdi, user = found.groups()
            frames = int(frames)
            require(frames == 8, f"Expected eight synchronous redraws per sample: {line}")
            pending.append({"scene": scene, "dpi": int(dpi), "width": int(width),
                            "round": int(round_), "frames": frames,
                            "usPerRedraw": int(elapsed) / frames, "privateBytes": int(private),
                            "workingSetBytes": int(working), "gdi": int(gdi), "user": int(user)})
        capture = CAPTURE.search(line)
        if capture and pending:
            name = capture.group(1)
            require(name not in samples and len(pending) == 5, f"Invalid redraw group for {name}")
            require([part["round"] for part in pending] == list(range(5)), f"Bad rounds: {name}")
            for part in pending:
                require(name.startswith(part["scene"] + "__")
                        and f"__{part['width']}dip__" in name
                        and name.endswith(f"__{part['dpi']}dpi.png"), f"Misbound redraw: {name}")
            samples[name] = pending
            pending = []
    require(not pending and len(samples) == 24, f"Expected 24 complete common redraw fixtures: {path}")
    return samples


def image_row(old, new):
    with Image.open(old) as old_image, Image.open(new) as new_image:
        a, b = old_image.convert("RGBA"), new_image.convert("RGBA")
        bounds = ImageChops.difference(a, b).getbbox(alpha_only=False) if a.size == b.size else None
        return {"file": old.name, "baselineSha256": sha(old), "candidateSha256": sha(new),
                "baselineDimensions": list(a.size), "candidateDimensions": list(b.size),
                "changedBounds": list(bounds) if bounds else None,
                "visualChanged": a.size != b.size or bounds is not None}


def compare(candidate):
    roots = {"baseline": gallery_root(BASELINE), "candidate": gallery_root(candidate)}
    files = {label: capture_files(root) for label, root in roots.items()}
    old, new = files["baseline"], files["candidate"]
    require(set(old) == set(new), f"Gallery filename drift: missing={sorted(set(old)-set(new))}; extra={sorted(set(new)-set(old))}")
    prior_files = {path.name for path in (gallery_root(PRIOR) / "fileops/visual-gallery").glob("*.png")}
    require(len(prior_files) == 516 and prior_files <= set(old), "Prior 516-capture lineage differs")
    added = set(old) - prior_files
    require(len(added) == 16, f"Expected 16 scene-145–147 captures: {sorted(added)}")
    expected_counts = {"145-direct-issues-row": 8, "146-diagnostic-counters-no-issues-row": 4,
                       "147-clean-no-issues": 4}
    require({scene: sum(name.startswith(scene + "__") for name in added)
             for scene in expected_counts} == expected_counts
            and all(name.startswith(tuple(scene + "__" for scene in expected_counts)) for name in added),
            f"New capture identities differ: {sorted(added)}")
    images = [image_row(old[name], new[name]) for name in sorted(old)]
    samples = {label: parse_redraw(root / "fileops/trace.txt") for label, root in roots.items()}
    require(set(samples["baseline"]) == set(samples["candidate"]), "Redraw fixture keys differ")
    paired = {}
    for name in sorted(samples["baseline"]):
        paired[name] = {}
        for metric in METRICS:
            a = median(row[metric] for row in samples["baseline"][name])
            b = median(row[metric] for row in samples["candidate"][name])
            paired[name][metric] = {"baseline": a, "candidate": b, "delta": b-a,
                                    "changePercent": (b/a-1)*100 if a else None}
    layout = {"scope": "24 paired same-HWND Release redraw fixtures, five rounds of eight synchronous paints each; not presented FPS.",
              "sampleCountPerFixture": 5, "fixtureCount": 24,
              "rawSamples": samples, "pairedMedians": paired}
    capture = {"scope": "532 matched actual application PNGs; SHA-256, dimensions, and RGBA difference bounds only. Visual review remains separate.",
               "commonPairs": 532, "commonChangedImages": sum(row["visualChanged"] for row in images),
               "commonDimensionChanges": sum(row["baselineDimensions"] != row["candidateDimensions"] for row in images),
               "newSceneNames": sorted(added), "common": images}
    return layout, capture


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=DEFAULT_CANDIDATE)
    args = parser.parse_args()
    layout, capture = compare(args.candidate)
    for name, data in (("layout-comparison.json", layout), ("capture-comparison.json", capture)):
        (args.candidate / name).write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                                             encoding="utf-8", newline="\n")
    print(json.dumps({"commonPairs": capture["commonPairs"], "changed": capture["commonChangedImages"],
                      "pairedRedrawFixtures": layout["fixtureCount"]}, indent=2))


if __name__ == "__main__":
    main()

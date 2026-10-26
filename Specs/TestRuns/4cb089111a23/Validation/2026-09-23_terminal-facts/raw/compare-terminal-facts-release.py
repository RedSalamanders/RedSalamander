"""Draft Release-only comparator. Run only after both archived galleries are final.

The 508 shared captures and 24 shared redraw fixtures are separate comparisons.
Scene 144 is candidate-only and is never represented as a paired observation.
"""

import argparse
import hashlib
import json
import re
from pathlib import Path
from statistics import median

from PIL import Image, ImageChops

EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
BASELINE = EVIDENCE / "terminal-facts-baseline-20260923"
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


def gallery_root(folder, case_filter="FileOps_VisualGallery"):
    roots = list(folder.glob("*-runner-artifacts/artifacts/selftest/last_run"))
    matching = []
    for root in roots:
        result = root / "run-all-tests-results.json"
        if result.is_file() and json.loads(read_text(result)).get("case_filter") == case_filter:
            matching.append(root)
    require(len(matching) == 1, f"Exactly one {case_filter} archived run required in {folder}: {matching}")
    return matching[0]


def parse_redraw(path):
    samples, pending = {}, []
    for line in read_text(path).splitlines():
        found = REDRAW.search(line)
        if found:
            scene, dpi, width, round_, frames, elapsed, private, working, gdi, user = found.groups()
            frames = int(frames)
            require(frames > 0, f"Nonpositive frame count: {line}")
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
    require(not pending and len(samples) == 24, f"Expected 24 complete redraw fixtures: {path}")
    return samples


def image_row(old, new):
    with Image.open(old) as a_image, Image.open(new) as b_image:
        a, b = a_image.convert("RGBA"), b_image.convert("RGBA")
        # A newly visible facts block can change popup height in a shared scene.
        # Record that geometry change; a pixel bound is meaningful only at equal size.
        bounds = ImageChops.difference(a, b).getbbox(alpha_only=False) if a.size == b.size else None
        return {"file": old.name, "baselineSha256": sha(old), "candidateSha256": sha(new),
                "baselineDimensions": list(a.size), "candidateDimensions": list(b.size),
                "changedBounds": list(bounds) if bounds else None,
                "visualChanged": a.size != b.size or bounds is not None}


def compare(candidate):
    roots = {"baseline": gallery_root(BASELINE), "candidate": gallery_root(candidate)}
    galleries = {label: root / "fileops/visual-gallery" for label, root in roots.items()}
    files = {label: {p.name: p for p in gallery.glob("*.png")}
             for label, gallery in galleries.items()}
    old, new = files["baseline"], files["candidate"]
    require(len(old) == 508, f"Baseline capture count {len(old)}")
    require(len(new) == 516, f"Candidate capture count {len(new)}")
    require(set(old) <= set(new), f"Missing common captures: {sorted(set(old) - set(new))}")
    extra = sorted(set(new) - set(old))
    expected_extra = {
        f"144-terminal-move-mixed-facts__{theme}__{width}dip__fr-FR__display{display}__{dpi}dpi.png"
        for theme in ("dark", "high-contrast")
        for width in (480, 760)
        for display, dpi in ((1, 144), (2, 96))
    }
    require(set(extra) == expected_extra, f"Unexpected candidate-only captures: {extra}")

    images = [image_row(old[name], new[name]) for name in sorted(old)]
    candidate_only = []
    for name in extra:
        with Image.open(new[name]) as image:
            dimensions = list(image.size)
        candidate_only.append({"file": name, "sha256": sha(new[name]), "dimensions": dimensions})

    samples = {label: parse_redraw(root / "fileops/trace.txt") for label, root in roots.items()}
    require(set(samples["baseline"]) == set(samples["candidate"]), "Redraw fixture keys differ")
    paired = {}
    for name in sorted(samples["baseline"]):
        paired[name] = {}
        for metric in METRICS:
            a = median(row[metric] for row in samples["baseline"][name])
            b = median(row[metric] for row in samples["candidate"][name])
            paired[name][metric] = {"baseline": a, "candidate": b, "delta": b - a,
                                    "changePercent": (b / a - 1) * 100 if a else None}
    layout = {"scope": "24 paired same-HWND Release redraw fixtures only; scene 144 is not paired.",
              "sampleCountPerFixture": 5, "fixtureCount": 24,
              "rawSamples": samples, "pairedMedians": paired}
    capture = {"scope": "508 filename-matched captures plus 8 candidate-only scene-144 captures; "
                       "dimensions and RGBA pixel differences (where dimensions match) are recorded, "
                       "not a UIA or whole-gallery perf claim.",
               "commonPairs": len(images), "commonChangedImages": sum(x["visualChanged"] for x in images),
               "commonDimensionChanges": sum(x["baselineDimensions"] != x["candidateDimensions"] for x in images),
               "candidateOnly": len(candidate_only), "common": images, "candidateOnlyImages": candidate_only}
    return layout, capture


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", type=Path, default=EVIDENCE / "terminal-facts-qualified-v2-20260923")
    args = parser.parse_args()
    layout, captures = compare(args.candidate)
    for name, data in (("layout-comparison.json", layout), ("capture-comparison.json", captures)):
        (args.candidate / name).write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                                             encoding="utf-8", newline="\n")
    print(json.dumps({"pairedRedrawFixtures": 24, "commonPairs": captures["commonPairs"],
                      "commonChangedImages": captures["commonChangedImages"],
                      "candidateOnly": captures["candidateOnly"]}, indent=2))


if __name__ == "__main__":
    main()

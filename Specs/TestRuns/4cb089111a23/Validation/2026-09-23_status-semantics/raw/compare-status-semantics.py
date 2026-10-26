"""Intermediate numerical comparison only; archived provenance is separately validated by the curator."""

from pathlib import Path
from statistics import median
import hashlib
import json
import re

from PIL import Image, ImageChops


EVIDENCE = Path("C:/RedSalamander.Perf/evidence/i26-ui")
FOLDERS = {
    "baseline": EVIDENCE / "status-semantics-baseline-20260923",
    "candidate": EVIDENCE / "status-semantics-qualified-20260923",
}
CHECKS = (
    "StatusChipPaintFit", "DoneHeaderPaint", "CompletedGroupBadgePaint", "HostedProgressValue",
    "CompletedGroupLabelFit", "TerminalSummaryFit", "TerminalSummaryChange", "SpeedLimitLabelFit",
)
REDRAW = re.compile(
    r"TerminalLayoutRedraw: (\S+) dpi=(\d+) width=(\d+) round=(\d+) frames=(\d+) "
    r"elapsedUs=(\d+) privateBytes=(\d+) workingSetBytes=(\d+) gdi=(\d+) user=(\d+)"
)


def read_text(path):
    raw = path.read_bytes()
    return raw.decode("utf-16" if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def parse_trace(path):
    lines = read_text(path).splitlines()
    checks = {key: [] for key in CHECKS}
    samples = {}
    pending = []
    for line in lines:
        for key in CHECKS:
            if f"{key}:" in line:
                fit = re.search(r"\bfit=(true|false)\b", line)
                if not (fit):
                    raise RuntimeError((path, key, line))
                checks[key].append(fit.group(1) == "true")
        found = REDRAW.search(line)
        if found:
            scene, dpi, width, round_, frames, elapsed, private, working, gdi, user = found.groups()
            if not (int(frames) > 0):
                raise RuntimeError('Comparison guard failed at original line 50')
            pending.append({"scene": scene, "dpi": int(dpi), "width": int(width), "round": int(round_),
                            "usPerRedraw": int(elapsed) / int(frames), "privateBytes": int(private),
                            "workingSetBytes": int(working), "gdi": int(gdi), "user": int(user)})
        capture = re.search(r"VisualGallery: (\S+\.png)", line)
        if capture and pending:
            name = capture.group(1)
            if not (len(pending) == 5 and [part["round"] for part in pending] == list(range(5))):
                raise RuntimeError(name)
            if not (name not in samples):
                raise RuntimeError(name)
            for part in pending:
                if not (name.startswith(part["scene"] + "__") and f"__{part['width']}dip__" in name):
                    raise RuntimeError('Comparison guard failed at original line 60')
                if not (name.endswith(f"__{part['dpi']}dpi.png")):
                    raise RuntimeError('Comparison guard failed at original line 61')
            samples[name] = pending
            pending = []
    if not (not pending and len(samples) == 24):
        raise RuntimeError((path, len(samples)))
    return {"checks": {key: {"count": len(values), "failures": values.count(False)} for key, values in checks.items()},
            "samples": samples}


def main():
    data = {}
    galleries = {}
    for label, folder in FOLDERS.items():
        run = folder / "runner-artifacts/artifacts/selftest/last_run"
        data[label] = parse_trace(run / "fileops/trace.txt")
        galleries[label] = run / "fileops/visual-gallery"
    expected = {
        "StatusChipPaintFit": 12, "DoneHeaderPaint": 8, "CompletedGroupBadgePaint": 4,
        "HostedProgressValue": 18, "CompletedGroupLabelFit": 28, "TerminalSummaryFit": 56,
        "TerminalSummaryChange": 12, "SpeedLimitLabelFit": 148,
    }
    baseline_failures = {"StatusChipPaintFit": 12, "DoneHeaderPaint": 8,
                         "CompletedGroupBadgePaint": 4, "HostedProgressValue": 6}
    for label in FOLDERS:
        for key, count in expected.items():
            if not (data[label]["checks"][key] == {"count": count,
                                                     "failures": baseline_failures.get(key, 0) if label == "baseline" else 0}):
                raise RuntimeError((label, key))
    if not (set(data["baseline"]["samples"]) == set(data["candidate"]["samples"])):
        raise RuntimeError('Comparison guard failed at original line 87')
    paired = {}
    for name, old in data["baseline"]["samples"].items():
        new = data["candidate"]["samples"][name]
        if not (len(old) == len(new) == 5):
            raise RuntimeError('Comparison guard failed at original line 91')
        paired[name] = {}
        for metric in ("usPerRedraw", "privateBytes", "workingSetBytes", "gdi", "user"):
            a, b = median(part[metric] for part in old), median(part[metric] for part in new)
            paired[name][metric] = {"baseline": a, "candidate": b, "delta": b-a,
                                    "changePercent": (b/a-1)*100 if a else None}
    data["pairedMedians"] = paired
    data["scope"] = ("Matched same-machine Release in-process redraw and process observations; "
                     "not presented FPS, live heap attribution, GPU surface bytes, Debug/PR/Full or UIA acceptance.")
    names = {path.name for path in galleries["baseline"].glob("*.png")}
    if not (len(names) == 508 and names == {path.name for path in galleries["candidate"].glob("*.png")}):
        raise RuntimeError('Comparison guard failed at original line 101')
    images = []
    for name in sorted(names):
        before, after = galleries["baseline"] / name, galleries["candidate"] / name
        with Image.open(before) as old, Image.open(after) as new:
            a, b = old.convert("RGBA"), new.convert("RGBA")
            if not (a.size == b.size):
                raise RuntimeError(name)
            box = ImageChops.difference(a, b).getbbox(alpha_only=False)
            images.append({"file": name, "baselineSha256": sha(before), "candidateSha256": sha(after),
                           "baselineDimensions": list(a.size), "candidateDimensions": list(b.size),
                           "changedBounds": list(box) if box else None})
    data["captures"] = len(images)
    data["changedImages"] = sum(row["changedBounds"] is not None for row in images)
    candidate = FOLDERS["candidate"]
    (candidate / "layout-comparison.json").write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8", newline="\n")
    (candidate / "capture-comparison.json").write_text(json.dumps(images, indent=2) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps({"checks": {key: data[key]["checks"] for key in FOLDERS},
                      "captures": data["captures"], "changedImages": data["changedImages"]}, indent=2))


if __name__ == "__main__":
    main()

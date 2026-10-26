# Speed-selector gallery visual review

Compared the retained `speed-selector-baseline-20260923` and `speed-selector-qualified-20260923` galleries under `C:/RedSalamander.Perf/evidence/i26-ui/`. These are application harness captures; no desktop interaction, native job, or repository edit was performed.

The qualified gallery contains 438 PNGs. I independently verified the SHA-256 values, image dimensions, and pixel difference bounds of all 438 entries in `capture-comparison.json` against both PNG sets (zero mismatches). Of those, 313 files are byte-identical and 125 changed. The only dimension changes are the scene-96 speed flyouts: 160x288 to 166x288 at 96 DPI and 241x432 to 249x432 at 144 DPI.

Visual review covered the supplied changed-captures contact, all 55 changed dark/96-DPI task-frame regions as baseline/candidate detail crops (analysis sheets `speed-selector-diff-crops-01.png` through `-05.png` in this directory), key original qualified scenes 11/133/134/135 at 480 and 760 DIP and both DPIs, scene 10 in light/high-contrast/rainbow, the 144-DPI-only changed scene 74, and both baseline/qualified scene-96 flyout originals. The 144-DPI full frames were sampled at the four speed-specific scenes rather than every changed 144-DPI file.

No new visual regression was found in the inspected action regions. The selector now shows the full French default `Limite : illimitée`, 80.0 MB/s, and extreme 999 GB/s values. Adjacent Pause, Démarrer, Ignorer la découverte, and Annuler labels remain readable, with distinct nonoverlapping controls in the narrow 480-DIP views; the 760-DIP views retain generous separation. In high contrast, the widened selector and adjacent controls keep complete focus/button outlines. The two flyouts retain every label and separator; their width increase is modest and height unchanged. Scene 134's discovery status chip is visibly abbreviated at 480 DIP, but its pixels are outside the changed bounds and it is present in the baseline; this is not a speed-selector regression.

The isolated flyout PNGs prove their own contents and dimensions, not the screen-space anchor relationship to the selector. These screenshots alone also do not prove keyboard/UIA behavior or a resource/performance result. The qualified run's separate assertions and measurements should be used for those claims.

Analysis crops in this directory are derived review aids, not application captures or proposed replacement screenshots.

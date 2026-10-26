# French open-discovery progress rows — 2026-09-21

These are actual, non-activating application captures from the deterministic
`FileOps_VisualGallery` harness, not proposal mockups. They use the I26 consumer
checkout at `Z:/src/RedSalamander`, qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`,
French locale, dark theme and a 480-DIP popup. The gallery run ID is
`20260921T090803Z-110040-97186e7e2c4440c79fe01f84cf15d81d`; its complete
manifest, synthetic scenario data and result receipt were produced under
`C:/RedSalamander.Perf/runs/<run ID>/artifacts/selftest/last_run/fileops/visual-gallery/`.

| Scenario | Display 1, 144 DPI | Display 2, 96 DPI |
|---|---|---|
| Discovery and transfer | [capture](07-discovery-and-transfer__dark__480dip__fr-FR__display1__144dpi.png) | [capture](07-discovery-and-transfer__dark__480dip__fr-FR__display2__96dpi.png) |
| Discovery on demand after Skip | [capture](08-discovering-as-needed__dark__480dip__fr-FR__display1__144dpi.png) | [capture](08-discovering-as-needed__dark__480dip__fr-FR__display2__96dpi.png) |

The [prior 96-DPI capture](../2026-09-20-common-name/native__07-discovery-and-transfer__dark__480dip__fr-FR__display2__96dpi.png)
placed transferred/discovered bytes and completed/discovered item counts in one
sentence. The candidate presents `Découverts` and `Transférés` as two short,
separately localized rows in the same reserved space. Both rows and their values
are fully visible at the two measured scales; the on-demand row remains legible.
The counter denominator remains open until traversal closes. The screenshots do
not establish all-card localization, screen-reader acceptance, or performance
non-regression.

Validation: the x64 Debug application-project harness build completed with zero
warnings/errors, the focused `cmd_pane_fileops_popup_progress_contracts` case
passed 1/1 with no skips, and `FileOps_VisualGallery` passed 1/1 with no skips.
The localization resource contract and spec inventory checks passed with zero
blocking findings. Full logs are under
`C:/RedSalamander.Perf/evidence/i26-ui/progress-rows-*-20260921.log`.
The runner reported 229 pre-existing scratch disk-audit issues; no clean-sandbox
claim is made from these receipts.

The runner has since cleaned the raw run directory; the four reviewed captures
above and retained external build/test logs remain. A project-target build is
not a reusable full-solution receipt.

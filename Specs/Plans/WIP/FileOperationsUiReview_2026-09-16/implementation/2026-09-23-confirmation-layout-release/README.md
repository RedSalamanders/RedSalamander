# Confirmation presentation and adaptive layout

[Open all 376 original application captures](gallery.html). The gallery opens filtered to the 76 confirmations. These are deterministic, nonactivating French application captures at native 96/144 DPI, not proposal mockups.

The old renderer produced black captured confirmations despite successful paint results. Composition presentation now makes the content visible. Long French option labels wrap at word boundaries; no options are dropped to fit a fixed panel. Copy/Move/Delete body prose uses readable neutral text, while severity remains in the title/icon/accent. A short window scrolls its body/options and preserves the action footer. Scroll input is enabled only for the host that routes it.

All 76 confirmation images were reviewed:20 wide,20 narrow,20 short top and16 short scrolled. No clipped controls, missing options, overlap or black panels remained. Both Delete variants fit at short height. Copy/Move options are fully visible after scrolling. The six French mid-word breaks found in the previous attempt are corrected.

| Example | Top | Scrolled |
|---|---|---|
| Copy, 96 DPI | [Top](candidate/102-confirmation-0-short__dark__fr-FR__display2__96dpi.png) | [Bottom](candidate/102-confirmation-0-short-scrolled__dark__fr-FR__display2__96dpi.png) |
| Move, 144 DPI | [Top](candidate/102-confirmation-7-short__dark__fr-FR__display1__144dpi.png) | [Bottom](candidate/102-confirmation-7-short-scrolled__dark__fr-FR__display1__144dpi.png) |
| Permanent Delete, 96 DPI | [Short host](candidate/102-confirmation-9-short__dark__fr-FR__display2__96dpi.png) | All content fits |

[Build, test and review receipts](../../../../../TestRuns/4cb089111a23/Validation/2026-09-23_confirmation-layout-release/README.md) bind the exact source and result. Gallery passes 1/0/0 with zero disk-audit findings. Fresh Release PR also passes all 13 executed/promoted entries, 824 reported cases, zero failures/skips and zero disk-audit findings. Capture coverage is dark-theme French 96/144 DPI; it does not establish native keyboard/UIA/drag, device-loss or performance/resource acceptance. Those gates remain open in I26.

The original scenario/display manifests and PNG bytes are retained in `candidate/`; `SHA256SUMS.txt` records their hashes. Earlier galleries and failed attempts are preserved. The other 300 images retain existing scenario coverage and are byte-identical to the preceding passing gallery.

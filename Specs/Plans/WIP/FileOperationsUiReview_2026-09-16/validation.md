# Validation record — revised 2026-09-16

The pack preserves the original English Debug evidence and adds French static and interaction Debug evidence. Captures from different receipts are identified individually; this is not a cross-configuration pixel-parity claim.

| Check | Result | Evidence / limit |
|---|---|---|
| x64 Debug / english-baseline | Passed native capture; 128 PNGs | [Receipt](evidence/build-receipt.json), [native results](evidence/fileops-results.json), [build log](evidence/build.log) |
| x64 Debug / french-static | Passed native capture; 300 PNGs | [Receipt](evidence/french-static/build-receipt.json), [native results](evidence/french-static/native-results.json), [build log](evidence/french-static/build.log) |
| x64 Debug / french-interaction | Passed native capture; 9 PNGs | [Receipt](evidence/french-interaction/build-receipt.json), [native results](evidence/french-interaction/native-results.json), [build log](evidence/french-interaction/build.log) |
| Final Debug compilation | 0 warnings, 0 errors | Final build logs retained; tests enabled; earlier English Release run also passed |
| Desktop input | Tab/Shift+Tab, actual target hover and owned mouse capture asserted at 144 and 96 DPI | [Interaction record](evidence/french-interaction/interaction.tsv); warning/lease and foreground/pointer ownership checks precede input |
| Monitor transition | Same HWND moves 144 → 96 → 144 DPI | Captures 109–113; [actual work areas](evidence/french-static/displays.tsv) |
| Runner plan tests | 56 source checks + 1 native inventory check passed | [Source tests](evidence/revision-plan-source-tests.json), [native inventory](evidence/revision-plan-runtime-tests.json); exact interaction case alone requests the lease |
| Test inventory / tooling governance | 7 + 12 passed | [Focused results](evidence/revision-inventory-tests.json); both gallery cases excluded from broad FileOps totals |
| Specification inventory | 0 blocking findings | [Validation log](evidence/revision-spec-inventory.log) |
| Tool inventory | 0 blocking findings | [Validation log](evidence/revision-tool-inventory.log) |
| Artifact integrity | All PNG hashes/dimensions and local links pass; gallery JavaScript syntax passes | [Integrity result](evidence/revision-artifact-check.json), [manifest](capture-manifest.json) |
| Visual inspection | All published application captures and all three revised concepts inspected | Contact sheets plus full-size risk, dialog and interaction screens; findings remain visible |
| Formatting / whitespace | Changed native sources formatted; diff whitespace checked; [one post-capture line-wrap record](evidence/post-capture-formatting.json) | No baseline reformat required |
| DxUi scope | Guidance only; no library implementation or consumer pin changed | Pinned `b129956d01b725cca131d1369580f61ba3def76b` |

The initial input driver failed foreground acquisition, then exposed coordinate scaling and pointer-release sequencing defects. Those attempts were excluded from the published manifest; the final lane requires successful native assertions before images are accepted. Input cleanup is drained before moving the HWND to the next monitor.

The final revised Release rerun could not build because an independently launched Release application owned the exact output executable ([build refusal](evidence/revision-release.log)). It was left running. Final French capture coverage uses the validated Debug binary; the earlier English Release capture passed, but it does not qualify the revised source in Release.

The previously run DxUi dependency validator rejected the pre-existing `src/DxUi.vcxproj.user`. This unrelated local file was not modified or deleted. DxUi skills/spec validation passed in the initial review. No library behavior changed here; this is not a capture blocker.

The native runner retains its existing sandbox audit findings (224 old run directories in the retained results). This work does not claim a globally clean sandbox or delete unrelated evidence.

## Reproduction and provenance

Use the two governed commands in the [README](README.md). The final source patch is [retained here](evidence/capture-source-revision.patch); the manifest records source-file SHA-256 values and each receipt's source snapshot. French resource loading is asserted by an actual translated Cancel label. Per-file DPI comes from `GetDpiForWindow`, not a scaled image.

- **english-baseline**: run `20260916T085559Z-81200-4b9ac2f15139453cb017b0aac343aee7`, receipt `cc2fb03b1470797b043fd577b2a3205b83e24c6d2f6ed441848b43100e47055f`.
- **french-static**: run `20260916T132245Z-112164-60aea098d9e2492eaded02eb7511e4de`, receipt `316ddd3f631a0ad6f3cf144697c3f89cc271d7be430bdee69e0207551ae0fec8`.
- **french-interaction**: run `20260916T131714Z-98580-340fc0efb8c049edb9c75d061d2097e5`, receipt `798f4c1f4567f455acaee810355027f9132d14e3ed35f2b4429839c4c242ef3a`.

This is focused harness and UI-review evidence. It does not establish Fresh Full, native ARM64/ASan, provider behavior, comprehensive tooltip/UIA/assistive-technology coverage or performance acceptance. Existing platform qualification remains with its current owners. 200% and Japanese/RTL require future native execution. Capture-driver completion is distinct from the product's open no-clipping requirements.

Native source hashes and the source patch describe the tested capture sources. One later clang-format string-literal line wrap is recorded separately; concatenated literal content and every non-string token were verified unchanged.

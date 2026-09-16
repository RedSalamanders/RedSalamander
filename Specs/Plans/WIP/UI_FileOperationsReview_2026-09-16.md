# File Operations UX design decision and screenshot evidence

Status: **DECISION, I26**. Priority: **P1**. Revised 2026-09-16.

Planned at `9c68026dd`. Drift: `git diff 9c68026dd -- RedSalamander/FolderWindow.FileOperations* RedSalamander/SelfTest Tests/TestSupport Specs/UI/UI_FileOperationsPopup.md`.

## Ownership

The harness implementation and requested capture blockers are complete. This WIP owner now routes the unresolved **product design decision** and subsequent implementation scope. It does not authorize the redesign. Review the [proposal](FileOperationsUiReview_2026-09-16/proposal.md), [gallery](FileOperationsUiReview_2026-09-16/gallery.html) and [validation](FileOperationsUiReview_2026-09-16/validation.md).

I18 retains confirmation policy and metadata outcomes; I3 retains stress qualification; I25 retains its provider/harness residuals. DxUi remains the canonical owner of shared controls. Proposed improvements may extend the library and need an explicit later consumer pin update.

## Completed capture work

- [x] Persist harness-only documentation screenshots and the explicitly authorized notified focus exception in both repositories.
- [x] Implement reusable owned-window PNG capture and named synthetic presentation scenarios.
- [x] Capture all 24 conflict classes, related task phases, results, representative themes, widths and long text.
- [x] Close destination history, Issues empty/populated/selected/scrolled, recovery menu and live confirmation driver gaps.
- [x] Load genuine French UI resources and long French sentences; capture both physical displays at 144 and 96 DPI.
- [x] Capture real keyboard focus, hover and pressed states, with owned input and a same-HWND monitor round trip.
- [x] Inspect 437 native PNGs / 124 named scenarios and three revised French mockups; preserve per-batch provenance.
- [x] Replace the rejected horizontal comparison with mandatory Incoming-above-Existing at every width.
- [x] Update [UI contract](../../UI/UI_FileOperationsPopup.md), [selftest workflow](../../Testing/Testing_SelfTests.md) and [helper catalog](../../Core/Core_SharedHelpers.md).
- [x] Validate focused native Debug interaction / Debug static capture, 76 runner/inventory/governance checks and spec/tool inventories.

## Open design decision

- [ ] Review the revised layout/copy and action placement with the user; reconcile proposed direct Ignorer with the current More policy.
- [ ] Route an implementation plan covering visible task controls, readable French actions, visible checkbox state, dialog DPI reflow, vertical path comparison and Issues detail.
- [ ] Define paired library/consumer resource/performance and accessibility qualification before implementing rendering changes.

The requested capture infrastructure has no remaining blocker. Later 200%, Japanese/RTL and full assistive-technology/platform qualification are explicit scope limits, not claims of coverage. Screenshots show declared fixtures rather than successful provider operations. Current clipping remains product work documented by the evidence.

## Safety and closeout

Default capture does not activate windows. The exact interaction case uses the existing warning and runner lease, owned-window/foreground checks and cursor/focus restoration. Abort on wrong process, unsafe paths, lost ownership or missing renderer content. Never substitute desktop control or manually edited application screenshots.

Archive this design owner when the decision is resolved and the approved implementation scope is routed; do not represent the proposed UI as implemented. The completed capture record and reproducible commands are retained in the review folder's validation and README.

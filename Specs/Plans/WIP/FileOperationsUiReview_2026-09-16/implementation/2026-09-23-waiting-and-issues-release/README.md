# Waiting cards and Issues decisions/results

[Open the 72 focused captures](gallery.html). These are actual French application screenshots from the nonactivating test harness at native 96/144 DPI. The complete run has 424 PNGs; [all capture hashes and baseline mapping](capture-comparison.json) retain provenance without duplicating every historical image.

Waiting cards now reserve only their rendered header, paths, optional totals/concurrency and actions. Running/paused cards keep graph space. Expanding or changing an expanded body reserves its full height on the first paint, preventing its content from running through the action row.

Issues scenes 123–132 show legacy unknown facts, a pending accepted decision, failure after Remplacer, partial Move, cached scope, removal, result-first reappearance, real history-cap eviction and reintroduction. The original conflict text remains separate from result prose; a new identity replaces an evicted one. Missing provider display facts use the existing exact-path fallback.

Root reviewed all 20 new Issues images and six waiting/footer/menu samples; Sol reviewed the 46 changed/new non-Issues images. Together these reviews cover all 66 changed/new captures. The selected row is visible and fixed action captions fit; long details scroll. Existing long-path clipping in the 107 destination flyout remains a separate open menu gate. The 358 byte-identical baseline images were not needlessly reviewed again. Geometry tests additionally cover 48 transition observations and 36 static waiting layouts.

| Example | Capture |
|---|---|
| Compact queued task, French 144 DPI | [Queued](candidate/04-queued-start-now__dark__480dip__fr-FR__display1__144dpi.png) |
| Running and queued, 96 DPI | [Mixed](candidate/30-running-and-queued__dark__760dip__fr-FR__display2__96dpi.png) |
| Accepted decision followed by failure | [Distinct facts](candidate/125-issues-failed-replacement__dark__fr-FR__display1__144dpi.png) |
| Selected history entry evicted | [Selection cleared](candidate/131-issues-history-evicted__dark__fr-FR__display2__96dpi.png) |
| Identical record reintroduced | [New identity](candidate/132-issues-history-reintroduced__dark__fr-FR__display1__144dpi.png) |

[Build and test receipts](../../../../../TestRuns/4cb089111a23/Validation/2026-09-23_waiting-and-issues-release/README.md): build zero warnings/errors, gallery 1/0/0, all layout observations pass. Fresh Release PR run `20260923T151747Z-4280-b086854ce22d4a3c94010dae62565664` passes **824/0/0 reported cases**, 541,148 ms, disk audit zero. Exact executed entries, capability omissions and reuse classifications are preserved in the original runner receipts.

Required Debug/Full, English/high-contrast/native 192-DPI, foreground keyboard/hover/drag/UIA/AT, device recovery, paired product resources and combined qualified DxUi pin adoption remain open. Foreground timing remains unanswered; no desktop input was attempted. This slice does not close I26 or the independent library resource gates.

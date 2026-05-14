# FolderView Column Scroll Gap Fix Plan

## Progress Checklist

- [x] Identify root cause of the first right-scroll stopping on the left gutter.
- [x] Add failing real-pane regression coverage for first right/left line scroll.
- [x] Add pure scroll-stop helper tests before production scroll changes.
- [x] Implement shared scroll-stop helper so first-column gutter is not a scroll stop.
- [x] Route horizontal scrollbar, mouse wheel, and page-key navigation through the shared helper.
- [x] Run build, pure tests, command regression, audit, and focused perf comparison.
- [x] Update authoritative specs/test docs with the scroll-stop contract.
- [x] Move this plan to `Specs/Plans/Done/` after verification.

## Contract

- Horizontal offset `0` is the canonical first-column stop and preserves the left gutter.
- A single right line-scroll from offset `0` must scroll past the first column to the second column's left edge, not to the first column's left gutter.
- A single left line-scroll from the second column's left edge must return to offset `0`, restoring the first gutter in one step.
- Mouse wheel, scrollbar line scroll, and keyboard/page navigation must share the same stop semantics.
- Existing per-column width and crazy-folder audit coverage must remain green.

## Verification Evidence

- Red phase: `folderView_visible_column_widths` failed before the fix with `offset=18.0 expected=238.5 firstLeft=18.0`.
- Red phase: `.\build.ps1 -ProjectName PerformanceTests2` failed before helper implementation because `ResolveNextScrollStop` / `ResolvePreviousScrollStop` did not exist.
- Build: `.\build.ps1 -ProjectName PerformanceTests2` passed with 0 warnings / 0 errors.
- Pure tests: `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll /Tests:FolderViewColumnLayoutTests` passed 4/4.
- Build: `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings / 0 errors.
- Real-pane regression: `folderView_visible_column_widths` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_160835`.
- Audit: `folderView_column_widths_audit` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_161257` with 48 records and zero integrity failures.
- Perf against original current-layout baseline `2026-05-14_123634`: `render.frame_us` P95 `55465 us -> 56894 us` (`+2.6%`, noise), `render.begin_to_enddraw_us` P95 `48058 us -> 49112 us` (`+2.2%`, noise), `render.draw_item_us` P95 `164 us -> 151 us` (`-7.9%`, noise/improvement).

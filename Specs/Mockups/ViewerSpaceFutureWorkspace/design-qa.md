# ViewerSpace future-workspace design QA

Date: 2026-10-22

**final result: passed** for the browser proposal and exercised flows. This result does not qualify native ViewerSpace, providers, accessibility technology, performance or the repository test matrix.

## Evidence and comparison

The visual reference was the user's supplied 1261 × 674 Explorer screenshot while scanning, plus the selected-folder inspector crop. Those environment-derived references were inspected alongside the rendered proposal and were not copied into this repository. Their paths, names and account data were not reused.

The proposal intentionally changes the reference: always-visible discovered size, a distinct shown total after exclusion/navigation, one name and parent location in the inspector, exceptional states only, one Layout control, a shorter filter disclosure and distinct cleanup marks. It retains the Windows desktop font family, green folder icons, nested rectangular map, linked dense analysis and menu/navigation anatomy. This is a future interaction proposal, not a pixel clone or a native application capture. The prototype-only top strip and outer margin consume height that would not be part of the product window.

Browser captures use invented Orion workspace data. The scan-wide measured baseline is 94,231,775,746 bytes (87.8 GB), 124,293 files and 16 folders; an unreadable folder explicitly has unknown contents. Counts and scan rates are fixture values, not benchmarks. Unknown allocation and summarized names are represented separately from measured zero.

| Capture | CSS viewport | State |
|---|---|---|
| [Desktop proposal](qa/proposal-desktop.jpg) | 1440 × 900 | Light, normal Size, panel right, selected folder |
| [Reference-size proposal](qa/proposal-reference-size.jpg) | 1261 × 674 | Light, scanning, visible discovered size and compact inspector |
| [Exclusions](qa/proposal-exclusions.jpg) | 1261 × 674 | Discovered unchanged; build excluded; cache marked for cleanup |
| [Cleanup review](qa/proposal-cleanup-review.jpg) | 1261 × 674 | Whole-folder review and explicit simulated host handoff |
| [Compact dark proposal](qa/proposal-compact-dark.jpg) | 840 × 760 | Automatic bottom arrangement and dark theme |

The browser also initially rendered at its natural 683 × 744 panel size. Viewport overrides were temporary. Captures are compressed browser JPEGs at the listed pixel dimensions, with no image editing, compositing or rescaling applied. Compression limits fine text/pixel comparison; these are flow/layout evidence, not native-DPI or font-rasterization evidence.

## Core design review

- **Typography and density:** Segoe UI Variable / Segoe UI with 12 px table text, compact secondary text and a prominent discovered figure. Increased the initial small text and merged the selection identity into its action row. Exact bytes are secondary. Labels truncate with complete tooltips; essential size/share columns fit without horizontal scrolling in the inspected arrangements.
- **Layout and spacing:** dominant map with a separated right panel; bottom layout places analysis and the inspector side by side. The single layout glyph stays available in map-only mode. Summary and navigation occupy distinct rows. Popup placement is clamped to the viewport and escapes scrolling panel clipping.
- **Color and tokens:** subdued folder-group colors, neutral mixed folders in file-type mode, per-file category colors, explicit type/age legends, green selected/focus cues and amber unreadable status. Exclusion retains the row's real size and adds dimming/overlay/state. Dark controls and scrollbars follow the dark color scheme.
- **Icons and assets:** actual Fluent UI React Icons throughout, with accessible action/type names. No raster product imagery was required. Map rectangles are quantitative geometry computed from fixture weights, rather than illustration placeholders. The summary pattern identifies summarized detail.
- **Content:** no repeated leaf path/name fields, Kind text or routine Done. Unreadable size/share are Unavailable. No volume/free-space bar, Git status, administrator action, saved scans, automatic cleanup suggestions or growth comparison controls were added.
- **Resilience:** inspected desktop, reference-sized and compact layouts. Persistent toolbar/footer actions remain reachable. Lists/expanded details scroll within their own region. At small desktop heights, fewer ranked rows are visible; scrolling preserves access to the complete retained list.
- **Keyboard/accessibility:** semantic labelled buttons, checkboxes, selects and tables; independent icon hit targets; keyboard context route; visible focus; menu arrow navigation; modal focus containment/restore; reduced-motion handling. Keyboard layout aliases and Escape hierarchy were exercised. Human screen-reader/UIA, high-contrast, native DPI and localization qualification remain future product work.

## Interactions exercised

| Flow | Observed result |
|---|---|
| Exclude build / restore | Shown falls from 87.8 GB to 60.1 GB; Discovered stays 87.8 GB. Map reflows and shares change. Excluded row remains real-sized; Restore all returns the original projection. |
| Filter by assets | One current-folder row remains; shown/discovered totals stay unchanged; map branches dim rather than reweight. |
| Navigate and refresh with exclusions | Folder header drills in; Up returns to root. Exclusion survives both navigation and simulated refresh; cached-result age is shown during refresh. |
| Three layout choices | Right, Bottom and Map only switch through the same menu. Map-only hides the panel and leaves the restoration control. Narrow windows use Bottom without changing the preferred Right choice. Ctrl+2 / Ctrl+1 work. |
| Cleanup marks and filtered bulk marking | Included retained real rows are snapshotted into review, separate from exclusions. Review shows target size, whole-folder consequences, removal and operation choice. Simulated Send clears marks and reports a simulated handoff, with no filesystem operation. |
| Size / File count / Size on disk | Areas and ranked values change with the selected metric. Count mode labels file counts. Disk mode shows incomplete known allocation and unavailable entries without substituting logical bytes. |
| Color and breakdown | File-type and age/newest-contained-file modes switch, legends update, breakdown expands and shows size/counts. Type/date controls apply to the current folder with an explicit date basis. |
| Age filter | Over-one-year / Modified retains the old backup file; other current-folder rows do not match. |
| Export | JSON and CSV previews contain current filtered rows and exclusion/summary/unavailable fields. No automated download was initiated. |
| Problems / Locate | Unreadable-folder reason and summary limits are visible; Locate selects the affected item. No administrator-rescan action appears. |
| Selection and pane menus | Copy path / Alt+Insert, Exclude/Restore and explicit Open folder in other pane are present. The latter reports a simulated host action and has no dedicated shortcut. |
| Advanced settings | Everyday display choices remain editable in the proposal. Advanced explains configured workers, memory admission, retention and in-memory reuse; no native settings are written. |
| Escape and fixtures | Running scan cancels and retains partial results; idle viewer closes to a reopen state. Empty fixture shows zero entries with a useful inspector instruction and no scan problems. |
| Theme / reset | Light and dark themes render; Reset returns the projection, marks, filter and list scroll to the initial fixture. Theme preview is session-local. |

Clipboard success/failure prompts, text-field Ctrl+C and OS download dialogs were not exercised against the user's clipboard/download area. Their frontend paths are implemented; the full native clipboard/host contract remains unqualified. No live file/provider integration, hard-link identity scenario, actual scan, memory budget enforcement, minimized process priority or performance measurement is simulated as passing evidence.

## Corrections made during QA

1. Nested folder buttons originally covered their children's centers. Restricted a nested folder's actionable header so clicking its labelled header reliably drills into the folder while children retain their own selection targets.
2. Popup menus could have been clipped by scrollable panels. Positioned and bounded them against the browser viewport; added menu keyboard focus/arrow behavior.
3. Unreadable and summarized inspector values initially risked displaying zero or fabricated metadata. Replaced unknown values with Unavailable and disabled synthetic aggregate path copying/exported path values.
4. Kept group colors stable when exclusions reorder map rectangles; current-folder name filtering highlights whole matching branches without recursively fading their children.
5. Normalized nested exclusions, required restoring an excluded ancestor explicitly and removed conflicting cleanup marks. Review now states that a folder target includes all its contents, even files outside a filter.
6. Increased small text, compressed redundant selection chrome, fixed dark scrollbar styling, and reset list scroll with the demo/filter/navigation.

## Build and console evidence

- `npm run build`: passed on the finished prototype; React/Vite produced the client bundle and retained hosting scaffold output. No publication occurred.
- Browser console was checked. A transient Vite reload error occurred during an intermediate source edit; the final build passed and the finished page was then freshly reloaded and checked again. No remaining runtime error is accepted as part of this pass.
- `Test-Changes.ps1 -Explain -TestRoot D:\RedSalamander.Perf` generated an advisory plan, with unknown/new inputs selecting broader native obligations. It did not run those suites or establish any repository verdict. Native implementation and qualification remain open in the roadmap.
- Changelog fragment validation/folding and the specification inventory are repository documentation checks; they do not constitute native behavior validation.

The prototype is reviewable as a local browser artifact. No C++ product sources, normative shipped-UI contracts or existing validation gates were changed for this prototype.

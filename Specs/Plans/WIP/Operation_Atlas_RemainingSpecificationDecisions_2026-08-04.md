# Operation Atlas — Remaining Specification Decisions

> **FILE OPERATIONS RESOLUTION (2026-08-21).** `ATLAS-DEC-CLIP-01` is closed: a cut payload is consumed when the complete Move request is accepted and is never restored, including `Copied; source kept`. `ATLAS-DEC-BR-01` is closed: Batch Rename submits a normal `RenamePlan` and uses ordinary File Operations queue/card behavior. The durable contracts are `Specs/FileSystem/FileSystem_FileOperations.md` and the FolderView/Batch Rename UI specs.


> **NON-NORMATIVE DECISION QUEUE.** This file records questions that cannot be
> answered from current implementation, tests, public behavior, and accepted
> contracts without making a new product or architecture choice. It does not
> change current behavior. Current contracts remain in `Specs/<Domain>/`.

## Status

- **State:** DECISION
- **Priority:** P3 unless a row is promoted by its named owner
- **Planned at:** `e14ba96e30d7f767b4020f271f07a0936fe56b7d`
- **Ownership boundary:** unresolved proposals removed from normative prose by
  Operation Atlas; no implementation is authorized by this file
- **Drift check:**
  `git diff e14ba96e30d7f767b4020f271f07a0936fe56b7d..HEAD -- Specs RedSalamander Common Plugins Tests Tools`

## Decision ledger

| ID | Question | Current contract and evidence | Decision needed | Promotion criteria |
|---|---|---|---|---|
| ATLAS-DEC-VFS-01 | Should directory enumeration gain a paged/streaming ABI? | `IFileSystem::ReadDirectoryInfo` and the cancellable variant return a complete result; `Specs/Plugins/Plugins_VirtualFileSystem.md` owns the current ABI. | Keep the complete-result model, or approve a new COM IID and host incremental-update semantics. | Representative 100K-item traces, memory/latency budgets, cancellation and compatibility design, then a separate perf-first implementation plan. |
| ATLAS-DEC-VFS-02 | Should mutable remote providers detect changes not initiated through RedSalamander? | Synthetic directory-watch notifications cover successful host/plugin mutations; unrelated backend changes require refresh unless a provider has its own current mechanism. | Choose polling, provider subscription, per-provider behavior, or continued manual refresh. | Provider capability matrix, request/power budgets, reconnect/cancellation semantics, and an owner-specific plan. |
| ATLAS-DEC-FV-01 | Which additional FolderView presentation and interaction modes belong in the product? | Current FolderView uses its documented column layout, selection model, rename prompt, single-key sorting, and pane-to-pane file operations. | Decide independently on list/details table mode, grouping, marquee selection, inline rename, Quick Look, resizable detail columns, and multi-key sorting. | UX contract, keyboard/UIA behavior, settings migration, performance scenarios, and one separately approved plan per selected slice. |
| ATLAS-DEC-NAV-01 | Which optional NavigationView features belong in the product? | Current behavior includes cache-first path suggestions, breadcrumb/path editing, history, provider menus, and disk information. | Decide independently on breadcrumb icons/colors/animation/badges, recent files, bookmarks, search or copy-path integration, a network-locations section, breadcrumb drag/drop/context menus, segment keyboard traversal, and touch gestures. | UX and accessibility contract, command/settings ownership, provider boundaries, performance scenarios, and one separately approved plan per selected slice. |
| ATLAS-DEC-PREF-01 | Should Preferences add full-text settings search? | The current category tree and handcrafted DxUi pages expose the implemented settings; schema `title`/`description` metadata is required for visible settings but is not a complete search index. | Reject search, or choose a localized static index versus runtime aggregation and define result navigation. | Search corpus/locale contract, stable setting-to-page mapping, keyboard/UIA behavior, performance budget, and a separate implementation plan. |
| ATLAS-DEC-PREF-02 | Should the embedded plugin editor support nested or otherwise richer JSON schemas? | Current embedded pages consume the common flat scalar/options shape and retain the advanced editor path for unsupported schemas. | Keep the bounded embedded shape or approve specific nested collection/object controls. | Schema vocabulary, validation/round-trip ownership, accessibility contract, compatibility fixtures, and a separate implementation plan. |
| ATLAS-DEC-MON-01 | Should Monitor add alternate selection/export modes? | Current behavior supports ordinary source-coordinate selection, `Ctrl+C`, `Ctrl+A`, transactional UTF-8 file export, and filtered visible-line rendering. | Decide on payload/metadata/CSV/JSON clipboard modes, quick-copy commands, pause-on-select, block selection, or multi-range copy independently. | Exact user semantics, command/resources, source-versus-visible coordinate rules, large-copy bounds, UIA behavior, and deterministic tests. |
| ATLAS-DEC-MON-02 | Should Monitor add content and metadata filtering beyond the six type bits? | The current filter is the six-bit `InfoParam::Type` mask; find/navigation and retention are separately bounded. | Decide on include/exclude tokens, regular expressions, hide-versus-highlight, PID/TID/time filters, and named presets. | Query grammar, settings schema/migration, filter/search interaction, retention semantics, performance budgets, and a separate implementation plan. |
| ATLAS-DEC-MON-03 | Is another Monitor rendering/storage optimization justified? | Current tail/scrollback modes, retained-history bounds, async slice layout, and metric gates are implemented. Old notes proposed adaptive tails, compressed inactive storage, and different incremental/prefetch policies without current comparative evidence. | Reject, defer, or select one evidence-backed optimization target. | A measured bottleneck, same-machine baseline, metric/sample contract, deterministic correctness guard, and perf-first implementation plan. |
| ATLAS-DEC-CMD-01 | Should registered parameterized command families replace the remaining resource/pane-specific dispatch routes? | `CommandRegistry.cpp` registers base families for known-folder navigation, plugin configure/toggle, and pane navigation; current menus and NavigationView still use their documented concrete routes. | Keep the dual model or approve one consolidated parameterized dispatch model. | Exact parameter grammar, enablement/error semantics, localization, shortcut compatibility, and deterministic command/menu tests. |
| ATLAS-DEC-WIN-01 | Should window placement persist continuously while moving or resizing? | Current owners persist accepted placement through their close/save paths and application shutdown; no repository-wide debounce contract exists. | Keep close/shutdown persistence or adopt bounded debounced persistence per eligible top-level window. | Save-coordinator impact, cross-process/settings contention, DPI/monitor semantics, and deterministic timer/shutdown tests. |
| ATLAS-DEC-FV-02 | Should FolderView show the complete filename in a hover tooltip when row text is clipped? | FolderView currently renders ellipsis but owns no tooltip path; DxUi has a shared tooltip surface. | Keep ellipsis-only behavior or approve a retained DxUi tooltip with keyboard/UIA semantics. | Trigger/dismiss rules, clipping detection, keyboard/UIA access, DPI/theme behavior, and hover-performance coverage. |
| ATLAS-DEC-DXUI-01 | Which broader bidirectional-text, grid-editing, and composition capabilities should DxUi add? | Current contracts cover logical-order text editing, the verified DirectWrite/TSF/UIA geometry set, readonly/toggle/range grid cells, relative layout, flip-sequential presentation, and overlay-only measured animation. | Decide independently on visual-order mixed-script navigation, remaining multiline/wrapped BiDi and TSF/UIA geometry, broader Grid mirroring, editable cells, or another composition pilot. | Exact semantics and accessibility surface, deterministic multilingual/runtime tests, and same-machine rendering/input evidence for performance-sensitive slices. |
| ATLAS-DEC-WIN-02 | Which currently fixed-size dialogs or tool surfaces should become resizable? | `UI_TopLevelToolWindows.md` owns the current top-level placement/minimum-size contract; individual modal and plugin windows have surface-specific behavior. A blanket “all windows” rule is not implemented. | Select exact windows and desired persistence/minimum-size behavior, or retain their current sizing model. | Per-window UX need, layout/DPI/accessibility contract, settings ownership, and focused resize tests. |
| ATLAS-DEC-HOST-01 | Should plugins receive a broader pane inspection and mutation host ABI? | Current public host services include `IHostPaneExecute` for active-pane navigation/focus/command execution and narrower alert/prompt/connections services. They do not expose arbitrary left/right selection indexes or pane-state enumeration. | Keep the narrow intent-based service or design versioned pane query/mutation interfaces. | Threat/ownership model, stable item identity instead of volatile indexes, UI-thread and callback lifetime, new IID/versioning, and multi-provider tests. |

## Resolved decisions

| ID | Decision | Owner and rationale | Reopen trigger |
|---|---|---|---|
| ATLAS-DEC-BR-01 | **ACCEPTED 2026-08-21:** Batch Rename submits `RenamePlan(BatchRename)` and uses ordinary File Operations queue/card behavior; it never inherits the Inline-F2 silent-success exception merely because it has one step. | `Specs/FileSystem/FileSystem_FileOperations.md` and `Specs/UI/UI_BatchRenameWindow.md` own the typed plan, confirmation, progress, cancellation, and presentation contract. | Reopen only if Batch Rename is intentionally removed from central File Operations execution or a separately controlling second task surface is proposed. |
| ATLAS-DEC-CLIP-01 | **ACCEPTED 2026-08-21:** RedSalamander consumes only its unchanged captured cut sequence when the complete Move request is accepted for queue admission, and never restores it after cancel, failure, partial, indeterminate, or `Copied; source kept`. | `Specs/FileSystem/FileSystem_FileOperations.md` and `Specs/UI/UI_FolderView.md` own sequence validation, accepted-queue timing, external clipboard ownership, and result UX. | Reopen only if RedSalamander adopts a different multi-paste cut model or takes ownership of external clipboard payloads it did not create. |
| ATLAS-DEC-ABI-01 | **ACCEPTED 2026-08-09: current source-tree/release lockstep only.** The host and shipped plugins are built and packaged together. Existing Viewer, Host, and Terminal methods require their full current records; oversized unknown tails remain accepted, but no pre-transition or mixed-release binary is supported. Retained IIDs and `sizeBytes` must never be used to infer an older shifted layout. | Maintainer decision: every current ABI producer and consumer is in this source tree, so ABI changes may be atomic for the current generation. `Specs/Plugins/Plugins_PluginAPI.md`, `Plugins_ViewerPlugins.md`, and `Specs/Terminal/Terminal_EmbeddedPlugin.md` own the durable boundary. Reserved Terminal values/fields remain reserved and are not a shorter-prefix promise. No old-header fixture or compatibility adapter is required for an unsupported binary generation. | Before supporting third-party or mixed-release binaries, or promising compatibility with any older host/plugin release, approve a separate ABI-generation plan. It must negotiate or reject generation before interface invocation and add new IIDs/adapters plus frozen old/new x64 and ARM64 fixtures where layouts or semantics changed. |

## Rules

- A decision is not approval to implement.
- Record **accept**, **reject**, or **defer** with rationale and named owner.
- Accepted work moves to a separate executable WIP plan with authoritative spec
  targets, deterministic tests, performance validation where applicable, and
  STOP conditions.
- Rejected work is recorded here once and is not reintroduced into a normative
  spec as a future section.
- Deferred rows remain non-normative and must state a concrete reopen trigger.

## Verification

```powershell
pwsh -File Tools/Get-SpecInventory.ps1 -Root Specs -OutputRoot .build/SpecInventory -FailOnFindings
```

The queue is correctly classified only when it is indexed once as `DECISION`
in `Specs/Plans/WIP/README.md` and no normative file claims an option from an
unresolved ledger row as current or planned behavior. Resolved rows may point
to the authoritative specification that records their selected current policy.

## STOP conditions

Stop and request a maintainer decision if implementation would alter public ABI,
file-operation semantics, default keyboard behavior, accessibility, persistent
settings, or a measured performance budget. Do not infer approval from old plan
or history text.

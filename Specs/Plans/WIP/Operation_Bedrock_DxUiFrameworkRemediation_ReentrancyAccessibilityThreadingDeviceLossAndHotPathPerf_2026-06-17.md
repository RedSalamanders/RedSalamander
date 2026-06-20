# Operation Bedrock — DxUi Framework Remediation: Reentrancy UAF, Accessibility Threading, Device-Loss Recovery & Hot-Path Performance

**Status:** WIP
**Date:** 2026-06-17
**Author:** Independent adversarial multi-agent review of the entire `Common/DxUi` framework at `master` `a68274ade`. Workflow run `wf_ddf495f4-3e9` (115 agents, run in 6 throttled waves): 17 review lenses — per-component deep dives on every translation unit (core object model, Controls ×2 safety+arch, Accessibility ×2 correctness+arch, Menu, Grid, WindowHost, TextInput, TSF/TextStoreACP, ComboBox, Tree/Typeahead, Theme/Typography, PointerInput/Scrollbar/FocusRestore) **plus three whole-framework cross-cutting lenses** (D2D/DXGI device-loss resource lifetime; reentrancy / delete-on-stack / mutate-while-iterating; WIL-RAII & COM refcount compliance) → per-finding adversarial refutation pass that re-opened the real source at each cited line and tried to *disprove* it. **56 findings survived verification (1 critical, 6 high, 14 medium, 35 low); 41 refuted** as guarded / by-design / impossible.

**Scope of this pass (~45,634 LOC under `Common/DxUi`):**
- `DxUi.h` / `DxUi.Internal.h` / `DxUi.cpp` — core `Element`/`Control`/`Panel` retained tree, frame runtime, invalidation.
- `DxUi.WindowHost.cpp` / `DxUi.Win32Hooks.h` — HWND host, Win32 message loop, swap-chain/device + device-loss, DPI, focus, animation dispatcher.
- `DxUi.Accessibility.cpp` — UI Automation provider tree (the largest correctness surface in this plan).
- `DxUi.Controls.cpp` — Button/Toggle/RadioButton/ScrollPanel/MenuBar/TabControl etc.
- `DxUi.Grid.cpp` — virtualized grid/list. `DxUi.Tree.cpp` / `DxUi.Typeahead.cpp` — tree + type-to-select.
- `DxUi.TextInput.cpp` / `DxUi.SingleLineTextEditing.cpp` — editing engine. `DxUi.NativeTextInput.cpp` / `DxUi.TextStoreACP.cpp` — TSF `ITextStoreACP`.
- `DxUi.ComboBox.cpp`, `DxUi.Menu.cpp` / `DxUiNativeMenuInterop.h`, `DxUi.Theme.cpp` / `DxUi.Typography.h`, `DxUi.Scrollbar.cpp`, `DxUi.PointerInput.{cpp,h}`, `DxUi.FocusRestore.h`.
- Consumers touched by Theme findings: `RedSalamander/DxUiThemePalette.h`, `RedSalamander/SplashScreen.cpp`, `RedConfigure/RedConfigureSplashScreen.cpp`.
- Tests: `Tests/DxUiTests/*` (Accessibility, Animation, ComboBox, Controls, Grid, Menu, MultilineText, NativeTextInput, Rendering, TextField, Theme, Tooltip, Tree, WindowHost). Registered in the `Full` suite via `Tools/Run-AllTests.ps1`.

> **Anchors are relative to `a68274ade`. Re-grep every `~line` before editing — line numbers drift.**

---

## Why this plan exists

**Verdict: DxUi is NOT currently rock-solid — but it is very fixable, and almost every P0/P1 defect already has a correct sibling in the tree.** The architecture is sound and the team clearly understands the hazards: `Control::GetLifetimeToken()`/`std::weak_ptr<int> _lifetimeToken` is used correctly in `Button::InvokeDropDown`; UIA *mutations* are correctly marshalled to the UI thread via `SendMessageTimeout`; `DebugSimulateDeviceLoss` calls `Invalidate()` after discarding resources; `wil::BeginPaint` already exists in Helpers.h; `GetOrCreateMultilineLayout` and `WindowHost::GetTextFormat` already show the layout/format caching pattern. **The defects are that these correct patterns are applied inconsistently — and the sites that miss them are the load-bearing ones.** Three classes block sign-off:

1. **Reentrancy use-after-free on the hottest gestures.** The primary input-dispatch path dereferences a raw `Control*` target after a user callback that routinely destroys the control tree (double-click folder → navigate; tree-item invoke; context-menu action). **This is a reproducible heap UAF on everyday file-manager use, in production, today.** `UpdateHover` and several `ScrollPanel`/`TabControl` sites share the same unguarded shape. The fix pattern already exists in-tree; it is simply not applied here.

2. **The accessibility threading model is broken.** UIA read methods run on RPC worker threads and walk the live, UI-thread-affine `Control`/`Panel`/model tree with **no** synchronization; the one mutex that exists is taken only by providers, never by the UI thread during mutation/teardown — a *false* safety guarantee. `GetBoundingRectangles`/`Compare` don't even take that lock, and `Detach` destroys the tree after releasing it (shutdown UAF). **Any Narrator / Inspect / UI-automation user can crash the app during normal navigation.** Reachable, not theoretical.

3. **Device-loss recovery is incomplete.** On a real GPU TDR / driver reset / RDP transition the production render path discards resources and returns **without `Invalidate()`**, so a static window stays blank/stale until an unrelated event repaints it. The debug-only simulation path proves the fix is a single missing call.

Beyond those: a consistent performance anti-pattern (O(n²) DirectWrite layout creation per frame / per mouse-move in MenuBar/TabControl; uncached per-cell/per-row `GetCellData` + text layouts in Grid/Tree/TextInput/ComboBox; redundant UIA control-tree re-resolution), plus a cluster of bounded correctness bugs (grid bottom-row drop, tree typeahead jam, surrogate-pair/RTL/masked-caret text handling, dark-theme palette gaps). None of the perf/correctness items is individually catastrophic, but collectively they keep DxUi below the project's strict-correctness bar.

**Fix the class, not the instance.** Seven systemic patterns drive the 56 findings — fixing each once kills a cluster (see "Cross-cutting themes" below). In particular:
- The **reentrancy guard** (Theme 1) is *one* convention; apply it uniformly instead of patching each site.
- The **UIA thread-affinity** fix (Theme 2) is *one* model decision enforced globally, not N method patches.
- The **single-line DWrite layout cache** (Theme 3) is *one* shared facility that subsumes the MenuBar/TabControl/Grid/Tree/ComboBox/TextInput churn.

---

## Implementation Tracking Checklist (update first, before editing code in a slice)

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` needs a product decision.

| State | Slice | Sev | Component | Required proof before `[x]` |
|-------|-------|-----|-----------|-----------------------------|
| [ ] | B-S0-1 | CRIT | WindowHost / Controls | A double-click whose handler resets `_root` does NOT read freed memory: no `target->` deref after `OnMouseDown/OnMouseDoubleClick` without a re-validated lifetime token / `ControlBelongsToTree`. RED under AppVerifier/ASAN today. Mirror on `WM_LBUTTONUP`, `UpdateHover`, `ScrollPanel`, `TabControl::CloseTab`. |
| [ ] | B-S0-2 | HIGH | Accessibility | Every provider *read* either marshals to the UI thread (snapshot) or holds `GetAccessibilityTargetMutex()` for its full body AND the UI thread takes the same mutex around every tree/model mutation + the tree-destroying section of `Detach`. Concurrent Inspect/Narrator traversal during navigation/list-refresh shows no AV (AppVerifier). |
| [ ] | B-S0-3 | HIGH | WindowHost | After the device-loss discard branch in BOTH `Render` overloads, a full repaint is scheduled (`Invalidate()`), the device is recreated, and the first post-recreate present is forced full. RED: today a static window stays blank after simulated device loss with no further input. |
| [ ] | B-S1-1 | HIGH | Grid | With a viewport height that is a non-integer multiple of row height, the last partially-visible row is painted AND hit-testable (click/hover/context-menu/select); `_hoveredRow` reset on data change. |
| [ ] | B-S1-2 | HIGH | TextInput | Masked caret/selection/IME correct for emoji/CJK (dot count tracks code units); caret+selection rects correct under RTL/BiDi; undo never leaves a lone surrogate. |
| [ ] | B-S1-3 | HIGH | TextInput | Ctrl+Z undoes a typing run (not per-char); no full `_text` heap copy per keystroke; single-line layout + grapheme count cached and correctly invalidated. |
| [ ] | B-S1-4 | HIGH | Grid | Ordered groups + content rect computed once per frame; per-cell `GridCellData` not reconstructed per frame and not re-scanned every animation tick; geometry matches a freshly-computed value after each invalidation trigger. |
| [ ] | B-S1-5 | HIGH | TSF / Accessibility | `GetTextExt` computes the union rect without O(n) layouts; Tree UIA hit-test does not allocate a full `TreeItemData` per row; control/cell/item resolved once per public UIA method. |
| [ ] | B-S1-6 | HIGH | Menu / ComboBox | Closing the owner while a context menu is open breaks the modal loop (no stuck/no-capture state); `ComboBox::HitTestOverlay` only claims points inside field/popup; class-registered flag set only after `RegisterClassExW` succeeds; HRGN freed on `SetWindowRgn` failure. |
| [ ] | B-S1-7 | HIGH | Tree | Typeahead single-char fallback (one mistype doesn't poison the buffer); UIA-driven expand schedules an animation tick; keyboard context-menu scrolls selection into view first. |
| [ ] | B-S1-8 | HIGH | Theme | `MakeDefaultThemePalette(dark=true)` sets dark focus/pressed/knob/accent-variant/smokeOverlay; `accentHover/Pressed` recomputed whenever `accent` is reassigned (splash screens + rainbow header tints correct). |
| [ ] | B-S2-1 | MED | Controls / ComboBox / Tree / Typography | Shared single-line measurement/layout cache; MenuBar/TabControl/ComboBox/Tree-badge/Typography measurement no longer O(n²)/per-frame; rect-correctness selftests pass after each invalidation trigger. |
| [ ] | B-S2-2 | MED | Accessibility / WindowHost | Single `QueryPattern` source of truth for QueryInterface+GetPatternProvider; `BuildRuntimeId`/`MakeProvider<T>` helpers; `ScopedPaint`→`wil::BeginPaint`; root pane exposes exactly one element (Inspect check). |
| [ ] | B-S3-1 | LOW | framework-wide | Batch of behavior-preserving simplifications (Toggle/RadioButton dedup, const/non-const overload collapse, `ComputeScrollbarPageStepDip`, DPI-bitmap recreate, hidden-host timer release, dead-code removal). Each independently revertable; existing DxUiTests stay green. |

---

## Guiding Principle

> **Never deref a control across a user callback that could free it; never touch the retained tree or client models off the UI thread; never drop the recovery frame after device loss; geometry that matches the glyphs DWrite actually draws; one balanced Release per AddRef; WIL RAII for every handle.** Prove the dangerous / dead / wrong path on the user's actual route, not just the happy path.

## Validation Contract (mandatory)

- **Build gate:** `.\build.ps1 -ProjectName DxUi` then `.\build.ps1` (0 warnings / 0 errors, x64 Debug; spot-check ARM64 for the SIMD/DPI-sensitive slices).
- **Functional gate:** `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` (long; serialized by a machine-wide mutex — exit 3 means another selftest run holds it, retry). Register every new case in `Tests/DxUiTests/*` so it actually runs in `Full`.
- **For each P0/P1 slice the proof must make the dangerous/dead/wrong condition observable and FAIL (RED) on the current tree before the fix, then pass after.** Where a UAF/race is impractical to force deterministically (B-S0-1 teardown race, B-S0-2 UIA threading), the proof is (a) a line-by-line lifetime/lock trace showing "no raw `Control*`/`self->_member` deref after the callback / off the UI thread" **plus** (b) a structural assertion in DxUiTests **plus** (c) a run under **Application Verifier** (heap + lock checks) / ASAN driving the repro gesture.
- Archive perf before/after evidence (frame time, allocation counts) under `Specs/TestRuns/<machine>/DxUi/<timestamp>/`.

---

## Cross-cutting themes (fix once)

1. **Lifetime-token guard around reentrant user callbacks.** Pattern exists (`Button::InvokeDropDown`). Missing at WindowHost pointer-down, `UpdateHover`, `ScrollPanel` hover/capture, `TabControl::CloseTab`. Establish one "capture token before user callback → re-validate or bail after" convention; `ControlBelongsToTree(_root.get(), target)` is the existing validation helper. → **B-S0-1**.
2. **UI-thread affinity for the retained tree and client models.** `DxUi.h:2660` says "Model state accessed only on UI thread." UIA reads violate this on RPC worker threads. Make ALL provider reads either marshal to the UI thread (snapshot) or hold `GetAccessibilityTargetMutex()` for the full body *and* have the UI thread take the same mutex around every mutation + teardown. → **B-S0-2**.
3. **Single-line DWrite layout / model-callback caching on the paint+hit-test hot path.** `IDWriteTextLayout`/`IDWriteTextFormat` recreated (and `GetCellData`/`GetVisibleItem` materialized with owning `wstring`s) per frame / per cell / per mouse-move across Grid, Tree, TextInput, ComboBox, MenuBar, TabControl, Typography. Build one shared cache keyed on `(text, width, role, readingDirection, dpi)` + a per-frame content-rect/displayText memo; route all consumers through it. → **B-S1-3, B-S1-4, B-S1-5, B-S2-1**.
4. **Redundant UIA control-tree re-resolution.** Provider methods re-run `ResolveControl*`/`ResolveGridCellData`/`ResolveTreeItemData` several times per public call; `Supports*` predicates re-resolve from root on every probe. Resolve once per public method, thread it down. → **B-S1-5, B-S2-2**.
5. **Divergent magic-number policies for shared widgets.** Scrollbar track-click paging and pointer-input routing partially unified but each consumer still uses its own constants. Add `ComputeScrollbarPageStepDip`; finish-or-trim the central pointer-input router. (Mirrors the project's known "divergent fold schemes" family.) → **B-S3-1**.
6. **Manual COM/Win32 resource accounting the RAII rule forbids.** Hand-coded AddRef/Release-on-OOM across ~12 provider factories; 6 near-identical SAFEARRAY runtime-id builders; hand-rolled `BeginPaint/EndPaint`; an HRGN released-before-success-check. Factor `MakeProvider<T>`, `BuildRuntimeId`; use `wil::BeginPaint`. → **B-S1-6, B-S2-2**.
7. **Activation/notification logic copy-pasted across input paths.** Toggle/RadioButton state-transition+invalidate+notify duplicated 5×/3×; const/non-const hit-test overloads byte-identical. Centralize into `ApplyCheckedState`/`SelectSelf`; delegate const → non-const. → **B-S3-1**.

---

## Phase P0 — Ship-blockers: reentrancy UAF, accessibility threading, device-loss recovery

### B-S0-1. `[CRITICAL / reentrancy — pointer-down UAF]` WindowHost dispatch dereferences the hit-test target after a handler that destroyed it

**Symptom:** A heap use-after-free on the most common file-manager gestures — double-click a folder row, click a navigation-tree item, right-click + context-menu action — leading to crashes or silent memory corruption.

**Root cause:** `DxUi.WindowHost.cpp:2412-2472`. The `WM_LBUTTONDOWN/WM_RBUTTONDOWN/WM_*BUTTONDBLCLK` handler dispatches to a raw, non-owning `Control* target`:
```
const bool controlHandled = target && (doubleClick ? target->OnMouseDoubleClick(...) : target->OnMouseDown(...));
...
if (controlHandled) {
    if (target->IsFocusable())            // <-- direct deref of possibly-freed target (line ~2451)
        if (target->SupportsTextInput())  // <-- direct deref (line ~2453)
            SetFocusControl(target);
        ...
    CaptureMouse(target);
    RememberPointerButtonDown(target, ...);
}
```
`SetFocusControl`/`CaptureMouse` internally re-validate via `ControlBelongsToTree`, but `IsFocusable()`/`SupportsTextInput()` at `:2451/:2453` deref the raw pointer with **no** guard. The dispatched handler routinely tears down the tree synchronously: `Grid::OnMouseDoubleClick` → `OnGridRowActivated` (folder navigate → panel rebuilt), `Tree::OnMouseDown` → `OnTreeItemInvoked`/`OnTreeToggleExpanded`, right-click → `Control::OnContextMenu` → Delete/Open rebuilds the view. All return `true`, after which `target` is dangling.

**Fix:** Capture the lifetime token (`Control` already exposes `GetLifetimeToken()`/`std::weak_ptr<int> _lifetimeToken`) **before** the `OnMouseDown/OnMouseDoubleClick` call; bail out of the entire `controlHandled` post-dispatch block if it expired — or re-validate with `ControlBelongsToTree(_root.get(), target)` before any post-callback deref, routing `IsFocusable`/`SupportsTextInput`/`CaptureMouse`/`RememberPointerButtonDown` through that guard. **Apply the identical guard to `WM_LBUTTONUP`.** Establish this as the framework-wide convention.

**Sibling sites folded into this slice (same shape):**
- `WindowHost::UpdateHover` `:3785-3811` — re-validate `target` against `ControlBelongsToTree`/token before assigning `_hoveredControl` and before `OnHoverChanged(true)`/`OnMouseMove`. *(LOW today: not reachable with stock controls, but same defect.)*
- `ScrollPanel` `DxUi.Controls.cpp:6204-6210, 6432-6454, 6662-6731, 6825-6861` — raw `_innerHoveredChild`/`_innerCapturedChild` can dangle when a child handler replaces children; recompute from the live `_children` span after any dispatch instead of caching raw pointers across a callback. *(MED.)*
- `Controls.cpp:5197-5218, 6733-6742, 6688-6701, 2929-2933` and `TabControl::CloseTab` — snapshot a lifetime token before user callbacks (`onTabClosed` etc.) and skip post-callback member access/`Invalidate` if expired.

**Required proof:** A DxUiTests case double-clicks a row whose delegate resets `_root`; assert no UAF under AppVerifier/ASAN and that no `target->` deref occurs after the callback without re-validation (structural + grep). RED on current tree.

---

### B-S0-2. `[HIGH ×6 / concurrency + resource-lifetime — UIA threading]` Provider reads walk the live tree/models off the UI thread; teardown races in-flight reads

**Symptom:** Intermittent access violations and corrupted property reads whenever a screen reader / UIA client (Narrator, Inspect.exe, Accessibility Insights, automated UI tests) traverses the tree while the app adds/removes/replaces controls (navigation, dialog open/close, list refresh) — reachable in completely normal use by any assistive-technology user. Plus a shutdown UAF when a window/dialog closes mid-traversal.

**Root cause:** `AccessibilityProvider::get_ProviderOptions` (`DxUi.Accessibility.cpp:3026`) returns `ProviderOptions_ServerSideProvider` only (no `UseComThreading`), so UIA calls providers on its own RPC worker threads. Mutating actions are correctly marshalled to the UI thread (`IsCurrentThreadWindowThread()` + `SendMessageTimeoutW`, e.g. `:3981-3992, 4134-4145`). But **all read methods run directly on the worker thread**:
- **Tree walk without sync** (`:168-190, 3160-3479, 3497-3806, 5119-5232`): `GetPropertyValue`, `Navigate`, `get_BoundingRectangle`, `GetSelection`, `GetFocus`, `ElementProviderFromPoint`, `ResolveControl*`, `IsControlPathVisible`, `FindFirstSemanticControl` walk `Panel::GetChildren()` (a `std::span` over `std::vector<unique_ptr<Control>> _children`, `DxUi.h:1406-1429`) while the UI thread mutates that same vector (`Panel::AddChild`/`ClearChildren`, `WindowHost::SetRoot` reset `DxUi.WindowHost.cpp:1299-1306`) with **no lock**. `GetAccessibilityTargetMutex()` is referenced only inside Accessibility.cpp and **never taken by the UI thread during mutation** — it serializes providers against each other, not against the UI thread. A reallocating `push_back` or a `reset()`/`ClearChildren` during iteration is a textbook UAF/torn read.
- **Client model off-thread** (`:3195-3228, 3256-3360, 3457-3476, 4626-4666, 5234-5300, 5366-5383`): violates the explicit contract `DxUi.h:2660` "Model state accessed only on UI thread — no synchronization needed." `GetPropertyValue` (Tree) → `model->GetVisibleItem`; (Grid) → `BuildGridRowAccessibleName` loops `model->GetCellData` per column; `GetSelection` iterates `GetOrderedSelection`/`FindRowByStableId`; `FindTreeItemAtPoint` → `GetItemLayoutMetrics` per row — all on the UIA thread. Client model code (e.g. FolderView) mutates its store on the UI thread without locking *because the contract promised single-thread access*.
- **`GetBoundingRectangles` omits the lock entirely** (`:2513-2553`): unlike every sibling text-range method it takes no `scoped_lock(GetAccessibilityTargetMutex())`; calls `ResolveHost()`/`ResolveControl()`/`ResolveText()` + geometry on the worker thread, so it isn't even serialized against `Detach`. *(Also: it reads/clamps `_rangeStart`/`_rangeEnd` without the lock — MED data race.)*
- **`Compare` omits the lock** (`:2413-2426`): reads another range's mutable endpoints without the accessibility lock. *(LOW.)*
- **Teardown UAF** (`:116-124, 6002-6027` + `DxUi.WindowHost.cpp:1247-1263`): `UnregisterWindowHostAccessibilityTarget` stores `nullptr` into `target->host` under the mutex, but `Detach` then — **outside** the mutex — runs `_root->PropagateHost(nullptr)` and `_root.reset()`, destroying every `Control`. The atomic keeps neither the `WindowHost` nor its tree alive; a provider that captured a non-null host and is still walking (especially the unlocked `GetBoundingRectangles`) derefs freed memory.

**Fix:** Pick **one** coherent model and enforce it globally.
- *Preferred:* marshal **every** provider read to the UI thread via `SendMessageTimeout` (the mechanism already used for mutations), snapshotting needed strings/flags/rects into the request struct, so the off-thread provider never touches live `Control`/`Panel`/model memory.
- *Alternative (cheaper, broader):* keep reads off-thread but take `GetAccessibilityTargetMutex()` for the **full body** of every read AND have the UI thread take the same mutex around every tree/model mutation (`Panel::AddChild`/`ClearChildren`, `WindowHost::SetRoot` reset) and around the tree-destroying section of `Detach` (hold it across `PropagateHost(nullptr)` + `_root.reset()`, not just inside `Unregister`).
- *Immediate sub-fixes regardless of model:* add `scoped_lock(GetAccessibilityTargetMutex())` to `GetBoundingRectangles` and `Compare`; make `ResolveHost` hand out a lifetime-pinned host handle for the call duration; stop calling Tree/Grid model methods off the UI thread.

**Watch for:** added `SendMessageTimeout` latency under heavy screen-reader use (bound with timeouts + snapshot granularity), and **deadlock** if a UI-thread mutation path can be re-entered by a synchronous provider call — audit for that.

**Required proof:** Line-by-line lock/marshal trace ("no live tree/model deref off the UI thread; teardown holds the lock until `_root.reset()` returns") + a DxUiTests/Accessibility case that drives UIA traversal (Inspect-style provider walk) concurrently with navigation/list-refresh/window-close under **AppVerifier** (lock + heap) with no AV.

---

### B-S0-3. `[HIGH ×2 / resource-lifetime — device loss]` Render discards GPU resources without re-invalidating; static window stays blank after device loss

**Symptom:** After a real device-lost event (GPU TDR/hang recovery, driver update/restart, RDP/remote-session transition, hybrid-GPU/adapter reset — all surfacing as `D2DERR_RECREATE_TARGET`/`DXGI_ERROR_DEVICE_REMOVED`/`DXGI_ERROR_DEVICE_RESET` from `EndDraw`/`Present`), a static window (idle file pane, a modal/popup host with no animation) renders **blank or stale** until an unrelated event (mouse move, resize, focus, theme change) happens to invalidate it. The window looks hung. Because the device is shared across all WindowHosts on the thread and siblings aren't invalidated either, every sibling host also stays stale. A secondary hazard: when the swap chain is later recreated, the next paint may be a partial `WM_PAINT` whose dirty-rect-only `Present1` leaves the rest of the uninitialized FLIP back buffer as garbage.

**Root cause:** Both `WindowHost::Render` overloads — retail `DxUi.WindowHost.cpp:3314-3323` and the `ENABLE_TESTS` twin `:3526-3535` — handle device loss by `DiscardSizeDependentResources` / `DiscardDeviceResources` / `ResetSharedWindowHostGraphicsResources` then `return;` with **no `Invalidate()`**. Render is driven by `WM_PAINT` (`:2280-2286` via `ScopedPaint`/`BeginPaint`, which validates the update region `:1125-1150`); nothing schedules a new paint after teardown, so no further `WM_PAINT` arrives and the device is never recreated. The debug-only `DebugSimulateDeviceLoss` (`:2098-2104`) *does* call `Invalidate()` after the identical discard sequence — confirming the production path is simply missing it. The animation dispatcher only ticks while something animates (`:3875-3905`), so an idle window has no render loop to recover it. `WindowHost::Invalidate()` (`:1433`) already does the right thing (`InvalidateRect(_hwnd, nullptr, FALSE)`).

**Fix:** In the device-loss branch of **both** `Render` overloads, after the discard sequence, call `Invalidate()` (`InvalidateRect(_hwnd, nullptr, FALSE)`) before returning so a fresh `WM_PAINT` re-runs `EnsureSizeDependentResources`→`EnsureDeviceResources`. Set a "first frame after recreate" flag so the next present is forced **full** (non-partial) regardless of `rcPaint`. Optionally iterate the existing attached-hosts registry to invalidate sibling hosts so the whole UI recovers in one cycle.

**Required proof:** Extend the device-loss selftest: after the discard sequence assert a repaint is scheduled and the device is recreated on the next paint; assert the first post-recreate present is full. RED today (static window with no further input stays blank). *Risk: very low — this is exactly what the debug-simulation and resize/focus paths already do.*

---

## Phase P1 — Serious reliability / correctness / performance

### B-S1-1. `[HIGH / grid correctness]` Bottom partially-visible row is un-painted AND un-hittable

**Symptom:** Every grid whose viewport height isn't an exact multiple of row height (essentially all — `kScrollbarThicknessDip=12`, rows ~28dip) renders a blank strip at the bottom and leaves the last partially-visible row un-paintable **and** un-hittable: it can't be clicked, hovered, context-menued, or selected by mouse. Keyboard nav that scrolls a row flush to the bottom can land selection on an undrawn row.

**Root cause:** `DxUi.Grid.cpp:157-162, 3938-3950` — `BuildVisibleBodyItems` uses a floor-based end-row count; `HitTestPoint` (`:3468`), `GetVisibleRowCount`, `GetVisibleRowAt` all consume the same output, so the dropped row is invisible to both paint and hit-test.

**Fix:** For the **END** boundary only, compute `std::min(sectionRowCount, ceil((visibleBottomDip - sectionTopDip - epsilon) / safeRowHeightDip))` (add a dedicated ceil helper; keep the floor helper for `beginRowOffset`). Also reset `_hoveredRow`/`_hoveredColumn` in `NotifyDataChanged`/`SetModel` (already done after group toggles), or store hover as a stable row id via `FindRowByStableId` consistent with the selection model (folds in the MED stale-hover finding `:2068-2069, 2689`).

**Required proof:** DxUiTests/Grid: with a viewport height that is a non-integer multiple of row height, assert the last visible index == `floor((scroll+viewportHeight)/rowHeight)` and is hit-testable. RED today.

### B-S1-2. `[HIGH / text correctness]` Masked caret, BiDi/RTL geometry, and surrogate-pair undo are wrong

**Symptom:** (a) Masked password fields containing emoji/CJK supplementary/combining characters mis-position caret, drag-selection, and the IME composition window (clamped left) — display dot count ≠ text code-unit count. (b) Caret/selection rendering is visibly wrong for RTL/BiDi single-line input (Arabic/Hebrew, or LTR with embedded RTL runs). (c) Undo after typing an emoji can leave a lone half-surrogate → replacement glyph + corrupt copy/paste.

**Root cause:** (a) `DxUi.TextInput.cpp:3030-3039, 1127-1154, 2517-2555` index `displayText` (one dot per grapheme) with `_text` (code-unit) indices. (b) `DxUi.SingleLineTextEditing.cpp:535-559, 828-865` + `DxUi.TextInput.cpp:1525-1527` use left-anchored scalar X math instead of `HitTestTextRange`. (c) `DxUi.TextInput.cpp:2429-2437, 2908-2917` snapshot undo mid-surrogate-pair.

**Fix:** (a) make masked display text track `_text` in **code units** (one dot per UTF-16 code unit) or build an explicit code-unit↔dot index map; for the Concealed policy route caret/selection through a fixed dot position. (b) use `IDWriteTextLayout::HitTestTextRange` for selection and derive the caret rect from hit-test metrics (as the multiline path does); fix the left-anchored clamp in `TextField::Paint`. (c) suppress the undo snapshot for a lead-surrogate-only insertion (defer until the pair completes) — coordinate with B-S1-3's typing-run coalescing — and validate/strip lone surrogates at the buffer boundary.

**Required proof:** DxUiTests cases: masked caret with emoji; RTL caret/selection rects; undo-after-emoji leaves well-formed UTF-16. *Risk: medium — easy to regress the common LTR/ASCII case; add those as regression guards.*

### B-S1-3. `[HIGH / text undo UX + perf]` No undo coalescing; per-keystroke full-buffer copies and uncached layouts

**Symptom:** Ctrl+Z undoes one character at a time instead of by typing run; every keypress in a large multiline field does a full-text heap copy on the UI thread; every paint/blink recomputes `GetDisplayText()` copies, an O(n) grapheme scan, and a fresh single-line layout.

**Root cause:** `DxUi.TextInput.cpp:2404-2455, 2908-2917, 2805-2809` push a full-buffer snapshot per keystroke; `:1447-1538, 3030-3039, 3073-3098` recompute display text + layout per frame; `GetSecretVisibleDotCount` (`:1127-1154`) rescans graphemes per `GetDisplayText`.

**Fix:** Coalesce consecutive same-kind edits (insert-run / delete-run) into one undo transaction, broken by caret jump, selection change, focus change, or edit-kind change; treat a surrogate pair as one unit (subsumes B-S1-2c). Compute `displayText` once per Paint and thread it through; cache the grapheme/dot count, invalidated only when `_text` changes (alongside the existing concealed-mask epoch). Add a single-line `IDWriteTextLayout` cache keyed on `(text, width, height, role, readingDirection)` analogous to `GetOrCreateMultilineLayout`, reused across measure/hit-test/draw within a frame.

**Required proof:** DxUiTests: typing run is one undo unit; delete-run is one unit; boundary breaks (caret jump) split units; allocation count per keystroke drops; layout cache invalidates on text/width/dpi/role change (no stale glyphs).

### B-S1-4. `[HIGH / grid perf]` Per-frame content-rect/group recompute and per-cell `GetCellData` allocation storm

**Symptom:** For grouped grids a single frame does ~a dozen redundant `CollectOrderedGroups` allocations+sorts + `GetBodyContentHeight` scans; for wide grids each frame constructs O(visibleRows×visibleCols) `GridCellData` (5 `std::wstring` + `com_ptr`) via model callbacks — and animation frames double it via `HasAnimatedVisibleCells`. Avoidable UI-thread churn that hurts scroll/animation frame pacing.

**Root cause:** `DxUi.Grid.cpp:4014-4053, 4055-4067, 4115-4169, 2285-2316` (`GetContentRect`/`CollectOrderedGroups` recomputed many times per Paint); `:2122-2123, 2359-2378` (per-cell `GridCellData` rebuilt every frame + again every tick).

**Fix:** Compute ordered groups + content rect **once per frame** and thread them through, or memoize `GetContentRect` keyed on bounds/scroll/header/column state, invalidated by `NotifyDataChanged`/`SetModel`/resize; extend the existing `_cachedGroups` to the scrollbar geometry helpers. Reuse a single `GridCellData` per loop iteration (clear/reset rather than reconstruct) or populate `string_view`s/an arena; for animation detection, have the model expose a cheap "has animated cells" hint or cache the animated-cell set computed during Paint instead of a second full scan in `Tick`.

**Required proof:** DxUiTests/Grid: assert geometry equals a freshly-computed value after each invalidation trigger (data change, sort, filter, resize, dpi, column/header change); allocation/callback count per frame drops. *Risk: medium — cache must invalidate on every layout-affecting state.*

### B-S1-5. `[HIGH / TSF + UIA perf]` `GetTextExt` O(n) layouts; UIA tree-walk allocation churn; redundant resolution

**Symptom:** TSF/IME/screen-reader extent queries on a large range in a multiline field stall the UI thread (O(n) DirectWrite layouts per character); UIA hit-testing a Tree allocates a full `TreeItemData` per visible row per query; `Supports*` predicates and `Navigate` re-resolve the control tree from root multiple times per single UIA query.

**Root cause:** `DxUi.TextStoreACP.cpp:104-139, 680-693` (`GetTextExt` multiline = O(n) layouts); `DxUi.Accessibility.cpp:5366-5383` (per-row `TreeItemData` in `FindTreeItemAtPoint`); `:5302-5364` (`Supports*` re-resolve from root); `:3520-3546` (`Navigate` re-resolves up to 3×); `:3160-3479` (`ResolveControl`/`ResolveGridCellData` run multiple times per public call).

**Fix:** `GetTextExt` — use the existing `Control::TryGetTextInputRangeRects` to get the union rect in one pass, or compute the bounding rect from just start+end caret rects + line metrics; cache multiline layout/line-metrics for the call duration. Tree hit-test — add a geometry helper returning only `rowRect` for a `visibleIndex` without filling text, or compute the index arithmetically (as `DebugGetFirstVisibleIndex` does) + one `GetItemLayoutMetrics` validation. UIA — resolve the control (and grid-cell/tree-item data) **once per public method** and pass it into predicate checks; add a lightweight tree-item existence predicate that doesn't allocate the four `wstring`s. (Coordinates with B-S0-2: resolve-once shrinks the surface that must be marshalled/locked.)

**Required proof:** `GetTextExt` union-rect matches the per-index result for wrapped multiline text (selftest); UIA tree hit-test allocation count is O(1) not O(N); resolution count per public method == 1.

### B-S1-6. `[HIGH / menu+combo lifetime, hit-test, RAII]` Stuck modal menu, over-greedy combo hit-test, class-flag + HRGN bugs

**Symptom:** (a) The modal menu loop becomes stuck (no visible popups, no mouse capture) if the owner window is destroyed while a context menu is open — recovers only on the next mouse-down on the same thread or WM_QUIT. (b) While a combo dropdown is open, an outside click that should dismiss-and-act is swallowed → users click twice. (c) A one-time menu class-registration failure permanently disables every DxUi context menu for the process. (d) Each failed `SetWindowRgn` leaks an HRGN; popups apply a region on every create and DPI relayout.

**Root cause:** (a) `DxUi.Menu.cpp:4826-4848, 3120, 2032-2052`. (b) `DxUi.ComboBox.cpp:2187-2195` — `HitTestOverlay` ignores the hit point and claims every point while the popup is open. (c) `DxUi.Menu.cpp:2106-2119, 1307` — `s_classRegistered=true` set **before** `RegisterClassExW` and its result ignored. (d) `DxUi.Menu.cpp:2942-2945` — `region.release()` before checking `SetWindowRgn`'s result.

**Fix:** (a) in the modal loop detect the root popup HWND becoming invalid (`!GetRootPopup() || !IsWindow(root->hwnd)`) and call `controller.Dismiss()`, mirroring the async `WM_NCDESTROY` teardown. (b) hit-test against the actual interactive region (`GetHitBounds`/`GetPopupBounds` union); return `this` only when inside field or popup; let outside-down fall through (still dismiss the popup, just stop swallowing the click). (c) set the flag only **after** `RegisterClassExW` succeeds (treat `ERROR_CLASS_ALREADY_EXISTS` as success), as `NativeMenuBarHost::EnsureWindowClass` already does (`DxUiNativeMenuInterop.h:551`); check + log the return. (d) pass `region.get()` to `SetWindowRgn`; `region.release()` only after it returns non-zero so the `unique_hrgn` dtor frees on failure.

**Required proof:** Selftests — close the owner while a context menu is up (loop breaks); an outside click both dismisses the combo and reaches the underlying control; `SetWindowRgn`-failure path doesn't leak (grep + structural). *(c)/(d) are tiny and obviously correct.*

### B-S1-7. `[HIGH / tree interaction]` Typeahead jam, frozen UIA-expand animation, off-screen context-menu anchor

**Symptom:** (a) One mistyped/non-matching key (or repeated-letter cycling) jams tree type-to-select for ~1s. (b) Screen-reader/automation-driven expand shows a stuck partial-expand frame + stale animation state until an incidental repaint. (c) Shift+F10/Menu-key context menu appears clamped at the wrong place when the selected node is scrolled off-screen.

**Root cause:** (a) `DxUi.Tree.cpp:954-977` — no single-char fallback after the accumulated-prefix lookup fails (diverges from `ComboBox::OnChar`). (b) `:382-406` — `RequestExpandedState` from the UIA path never schedules an animation tick. (c) `:988-1000` — keyboard branch anchors before scrolling the row into view.

**Fix:** (a) mirror `ComboBox::OnChar` — after the prefix lookup fails, retry with `_typeaheadBuffer.assign(1u, ch)` and `FindNextTypeaheadMatch` again; only return false if the single-char retry also fails. (b) have `AccessibilityProvider::ExecuteExpandOnWindowThread` also call `host->RequestAnimation()`, or (better) make `RequestExpandedState` take the `WindowHost` and call `RequestAnimation()` itself so all callers benefit. (c) call `EnsureVisibleIndex(visibleIndex)` and recompute scroll-derived `rowTop` before computing the anchor.

**Required proof:** Selftests — repeated-letter cycling works; UIA-expand animation completes; keyboard context-menu anchors correctly on an off-screen selection. *Risk: low — all three match existing in-framework patterns.*

### B-S1-8. `[HIGH / theme correctness]` Dark default palette gaps + stale accent variants

**Symptom:** (a) Anything rendering with `MakeDefaultThemePalette(dark=true)` shows light-theme focus stroke / pressed fill / toggle-knob / smoke overlay — confirmed at `RedSalamander/SplashScreen.cpp:350` and `RedConfigure/RedConfigureSplashScreen.cpp:310` (both `SetTheme(MakeDefaultThemePalette(true))` with no overrides). (b) `accentHover`/`accentPressed` are derived once from the default accent and go stale when a caller reassigns `accent` — `RedSalamander/DxUiThemePalette.h:28` sets `palette.accent=theme.accent` but `:99-100` mixes the **default**-derived variants for rainbow header hover/press tints → wrong hue.

**Root cause:** `DxUi.Theme.cpp:901-958` (dark branch leaves `focusStroke`/`pressedFill`/`toggleKnobFill`/`toggleKnobCheckedFill`/`smokeOverlay`/accent at light defaults); `:908-909` (variants derived once).

**Fix:** In the dark branch explicitly set dark-appropriate `focusStroke`, `pressedFill`, `toggleKnobFill`, `toggleKnobCheckedFill`, `smokeOverlay`, and a dark default accent (recompute variants **after** any accent change) — or give `ThemePalette` a dark-aware default constructor. Provide `RefreshAccentVariants(palette)` (recomputes `accentHover`/`accentPressed` from `accent`, as `MakeThemePaletteFromViewerTheme:1017-1018,1030-1034` already does correctly) and call it wherever `accent` is set; fix the App and Monitor palette builders.

**Required proof:** Screenshot/selftest of the splash screens + rainbow header tints in dark mode; assert variants track a reassigned accent. *Risk: low — color-value corrections; eyeball the dark values.*

---

## Phase P2 — Performance hot-path caching & architecture consolidation

### B-S2-1. `[MED / perf]` O(n²) and per-frame DirectWrite measurement across MenuBar/TabControl/ComboBox/Tree/Typography

Establish the shared single-line measurement/layout cache from Theme 3 (keyed on `text, role, readingDirection, dpi[, width]`) and route all consumers through it.
- `DxUi.Controls.cpp:4331-4368, 4626-4693, 4609-4624` — **MenuBar::Paint** O(n²) layout creation per frame/hover: cache item widths + laid-out rects, rebuilt only on `SetItems`/`OnBoundsChanged`/`OnDensityChanged`/`OnHostDpiChanged`/`OnFlowDirectionChanged`; hoist `GetVisualHighlightIndex` out of the loop.
- `DxUi.Controls.cpp:4991-5031, 5033-5044, 5312-5455, 4939-4965` — **TabControl::Paint** O(n²); `GetTabRect` re-measures all preceding tabs per call: cache `MeasureTabWidthDip` + tab rects + `GetTotalTabWidthDip`/`NeedsOverflowButtons`, recomputed only on `_tabs`/bounds/scroll/dpi/density/flow change.
- `DxUi.ComboBox.cpp:814-865` — editable combo recreates 4–5 layouts/frame (and computes `EnsureEditableCaretVisible` twice): compute caret offset/scroll once per paint; cache the single-line layout.
- `DxUi.Tree.cpp:1054-1063` — per-row/per-frame badge layout + per-mousemove tooltip layout: cache badge width by `(badgeText, fontRole, dpi)`; cache tooltip overflow per hovered index.
- `DxUi.Typography.h:298-362` — measurement helpers query the system font collection on every call, in per-row dialog loops (`RedSalamander/ManagePluginsDialog.cpp:3292-3293,3386-3387`; `Preferences.*.cpp` per card): cache `IDWriteTextFormat` per `FontRole` + the font-collection/family-availability results, mirroring `WindowHost::GetTextFormat`.
- `DxUi.ComboBox.cpp:193-253, 2349-2389` — synchronous full-screen GDI BitBlt on the UI thread for Mica/Acrylic dropdown backdrop (already emits `dxui.combo.backdrop_capture_us`): derive backdrop from the rendered D2D surface or bound the capture to the visible popup region.

**Required proof:** before/after frame-time + allocation counts; a rect-correctness selftest after each invalidation trigger. *Risk: medium — many invalidation sites; convert incrementally.*

### B-S2-2. `[MED / architecture + RAII]` Single source of truth for UIA patterns, factored COM/Win32 helpers, honest root-pane shape

- `DxUi.Accessibility.cpp:2888-2996` — pattern eligibility encoded twice (QueryInterface vs GetPatternProvider as parallel dispatch tables that must stay in lock-step): drive both from one `QueryPattern(IID/PATTERNID, const Control*)` resolving the control once.
- `DxUi.Accessibility.cpp:1608-2012` — six near-identical RuntimeId SAFEARRAY builders: add `BuildRuntimeId(SAFEARRAY**, std::span<const LONG>)` + a prefix helper; name the `1001..1005` discriminators.
- `DxUi.Accessibility.cpp:5752-5844` — seven `Create*Provider` factories duplicate `AddRefTarget`/`new(nothrow)`/release-on-failure (a single omitted `Release` leaks a target + host for the window's life): add `template<class... Args> T* MakeProvider(Args&&...)`; collapse all factories + the ~5 open-coded sites.
- `DxUi.WindowHost.cpp:1125-1150, 2280-2286` — hand-rolled `ScopedPaint` violates the mandatory-RAII rule: replace with `PAINTSTRUCT ps{}; wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd,&ps);` and only `Render` when `hdc` is valid (matches `OnThemedMessageBoxPaint`).
- `DxUi.Accessibility.cpp:417-439, 3170-3193, 3380-3478, 3515-3519, 5130-5156` — root pane provider duplicates a single semantic child as both its own properties and a navigable child (UIA tree announces the control twice, claims Invoke/Toggle/Value it routes to the child, runtime IDs differ): pick ONE representation (pure pane OR collapsed child).

**Required proof:** QueryInterface and GetPatternProvider still agree; single-control dialogs expose exactly one element (Inspect.exe); `BeginPaint`-failure path guarded; existing DxUiTests green. *Risk: low but broad; F31/root-pane carry UIA-behavior risk — verify with Inspect.*

---

## Phase P3 — Low-risk simplifications (batch opportunistically once structure is stable)

Behavior-preserving; each independently revertable; keep existing DxUiTests green.

| Item | File:line | Action |
|------|-----------|--------|
| Toggle/RadioButton activation duplicated 5×/3× | `Controls.cpp:2505-2595` | Extract `Toggle::ApplyCheckedState(WindowHost&,bool,bool)` + `RadioButton::SelectSelf(WindowHost&)`; call from all paths (also where reentrancy guards land). |
| Byte-identical const/non-const hit-test overloads | `Controls.cpp:6384-6430, 1195-1277` | Delegate const → non-const (small `const_cast` wrapper / shared templated helper); collapse `FindChildAtContent`/`FindOverlayChildAtContent`. |
| Divergent scrollbar paging constants | `Grid.cpp:2876, 2899` (+ Tree/ScrollPanel/Combo) | Add `ComputeScrollbarPageStepDip(trackRect, viewport, content)`; route all four consumers through it (DPI/viewport-aware). |
| `ResolveScrollbarVisuals` 4-arg recomputes what 6-arg recomputes | `Scrollbar.cpp:22-37` | Have the 6-arg overload reuse the resolved targets. |
| Chevron column reserved globally vs per-item | `Menu.cpp:2257-2258, 2343-2346` | Reserve the chevron column uniformly when any submenu exists (cosmetic). |
| Stale target-bitmap DPI after `WM_DPICHANGED_AFTERPARENT` (size unchanged) | `WindowHost.cpp:2154-2190, 3050-3082` | Force size-dependent bitmap recreation in `OnDpiChanged` after `SetDpi`. |
| 8ms animation timer kept hot while host hidden | `WindowHost.cpp:3875-3905` | Drop the subscription when hidden (re-subscribe on next visible/Invalidate), or skip wakeups when no subscriber has work. |
| Redundant deep copy of `beforeItems` on tree expand | `Tree.cpp:389-404` | Make `beforeItems` non-const; `std::move` into `BeginTreeExpansionAnimation`. |
| PointerInput source kinds defined+tested but never produced | `PointerInput.h:12-19` | Finish routing modal-loop/popup/forwarded-child through the module, or trim the unused kinds to keep the API honest. |
| Dead code | `Tree.cpp:1358-1365` | Remove `HasActiveExpanderAnimation`; wire or delete `NormalizeTypeaheadChar` (+ its test). |

---

## Appendix — base commit & re-verification

- All anchors relative to `master` `a68274ade` (the worktree branch and `master` were at the same commit at review time). **Re-grep every `~line` before editing.**
- The review **refuted 41 of 97 raw findings** (guarded / by-design / impossible); only the 56 above survived an adversarial re-read of the real source. The single-UI-thread assumption was applied when judging "race" claims — the surviving concurrency findings (B-S0-2) are specifically the cross-thread UIA cases.
- Workflow run `wf_ddf495f4-3e9` (115 agents). Standing rule per the Validation Contract: a slice is not done because a selftest is green — for each P0/P1 slice the dangerous/dead/wrong condition must FAIL (RED) on the current tree before the fix, proven on the user's actual route (UAF/race slices: line-by-line lifetime/lock trace + structural assert + AppVerifier/ASAN run).

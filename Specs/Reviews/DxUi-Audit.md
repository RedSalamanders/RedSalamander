> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


I'll produce the final audit report directly. The findings are already verified with corrected severities, so I'll synthesize, deduplicate by root cause, and prioritize.

# DxUi UI-Framework Audit — Final Report (RedSalamander)

## 1. Verdict

**The DxUi framework has a systemic reentrancy/lifetime defect that is the single most dangerous theme in this audit: raw, non-owning `Control*` pointers are dereferenced *after* a synchronous user callback that is free to destroy them.** This pattern recurs across the central input dispatcher (`WindowHost::HandleMessage`) and individual controls (`RadioButton`, `ScrollPanel`, `PageHost`, `MenuBar`, `NativeMenuBarHost`). It is "RED under AppVerifier/ASAN today" per the project's own remediation plan (Operation Bedrock, item B-S0-1, still unchecked). The second-most dangerous theme is a genuine **cross-thread use-after-free in the UIA accessibility provider**: read methods walk the *live* control tree and grid/tree models on UIA RPC worker threads, guarded by a mutex that no UI-thread mutation path ever acquires — a dangling `std::span` deref reachable by any running screen reader. Beyond memory safety, the TSF/IME surface is fragile (ACP cache desync, O(n)/O(n²) UI-thread stalls inside TSF locks, IME composition committed on abnormal teardown), the UIA contract is violated in two ways (unstable tree runtime-ids, non-expanding text units), and device-removed recovery leaves windows blank. None of these are theoretical: the dispatcher UAF and the UIA UAF are reachable in normal interactive use.

**Most dangerous themes, ranked:**
1. **Post-callback dereference of dangling `Control*`** (dispatcher + controls) — the dominant memory-safety cluster.
2. **UIA reads racing live tree/model mutation** — cross-thread UAF + data race.
3. **TSF/ITextStoreACP UI-thread stalls and state desync** — DoS-grade hangs and IME corruption.
4. **Unhandled-on-static-window device-removed** + **noexcept-boundary `std::terminate`** — robustness cliffs that turn recoverable conditions into blank windows / hard aborts.

---

## 2. Findings by Severity

### A. Memory-safety / Use-after-free / Crash

---

**A1. [MERGED] Input dispatcher dereferences hit-test `target` after its own handler may have freed it**
`Common/DxUi/DxUi.WindowHost.cpp:2457` (pointer-down), `:2503`/`:2534` (pointer-up diagnostics)
**Severity: use-after-free — CONFIRMED (down-path unconditional; up-path diagnostics-gated).**

This is the keystone finding; two verified reports (`windowhost-lifecycle` and `input-scrollbar-focus`) describe the same root cause and are merged here.

- **Down path (unconditional, normal builds):** `WM_*BUTTONDOWN/DBLCLK` calls `target->OnMouseDown(...)` / `OnMouseDoubleClick(...)` at `:2420-2421`, then on `controlHandled` unconditionally reads `target->IsFocusable()` (`:2459`), `target->SupportsTextInput()` (`:2461`), `CaptureMouse(target)` (`:2471`), `RememberPointerButtonDown(target, ...)` (`:2478`). `IsFocusable()` is a real member read (`DxUi.cpp:100`), so `:2459` is a vtable/field read of freed memory. `OnMouseDown` can synchronously free `target` via `SetRoot` (`_root = std::move(root)`, `:1313`), a nested `ContextMenu::Show` modal loop whose chosen command swaps the root, `ComboBox::CommitSelection` (which the code itself documents may "re-enter and mutate this control"), or `Tree::OnMouseDoubleClick → OnTreeItemInvoked` rebuilding the panel. `CaptureMouse`'s `ControlBelongsToTree` guard (`:1707`) runs *after* the `:2459` read and only pointer-compares — it does not protect the earlier virtual calls.
- **Up path (lower exposure):** `target->OnMouseUp(...)` at `:2505`; the only post-callback deref is `DescribeWindowHostTraceControl(target)` at `:2534`, gated behind `IsContextMenuDiagnosticsEnabled()` (compile flag + opt-in env var — dead code in release). `ReleaseMouseCapture()` at `:2540` is pointer-compare-only and safe.

**Failure scenario:** User invokes a tree/grid row or toolbar action whose down-handler navigates and replaces the root; on return `target->IsFocusable()` reads a freed control → UAF / vtable crash, intermittent depending on heap reuse.

**Fix:** After the handler returns, re-validate before any further use: `if (controlHandled && ControlBelongsToTree(_root.get(), target)) { ...focus/capture/remember... }`. Mirror the existing pre-dispatch guard at `:2406`. Apply the same `ControlBelongsToTree` gate to the up-path diagnostics block (`:2534`) and the down-path diagnostics block (`:2449`). This is project item **B-S0-1**.

---

**A2. RadioButton reads `_onSelected` after the group selection callback may have destroyed `this`**
`Common/DxUi/DxUi.Controls.cpp:2817` (also `OnKeyDown:2841`, `OnMnemonic:2875`)
**Severity: use-after-free — CONFIRMED.**

`RadioButton::OnMouseUp` calls `_group->SelectItem(this)` (`:2810`) *first*; `RadioButtons::SelectItem` fires `_onSelectionChanged(_selectedIndex)` synchronously (`:2959`). That callback can tear down the page that owns the button via `PageHost::SetPage`'s non-animated path (`_currentPage = std::move(page)`, `:1301`), freeing `this`. Control then returns and reads `const std::function<void()> onSelected = _onSelected;` (`:2817`) from freed memory, copy-constructing a `std::function` from freed bytes and invoking it → UAF / arbitrary execution. `Toggle::OnMouseUp` (`:2526`) demonstrates the correct discipline (callback copied to local, invoked as the last statement, no member touched after). All three RadioButton entry points violate it.

**Fix:** Never read `_onSelected` after `_group->SelectItem(this)`. Cleanest: have `SelectItem` defer firing `onSelectionChanged` until after `OnMouseUp` unwinds, or merge the two notifications so only one callback (which must not assume the button survives) runs.

---

**A3. ScrollPanel stores the captured-child raw pointer *after* dispatching the child's mouse-down**
`Common/DxUi/DxUi.Controls.cpp:6611` (overlay), `:6655` (standard child)
**Severity: use-after-free — CONFIRMED.**

`OnMouseDown` calls `child->OnMouseDown(...)` and only on `true` executes `_innerCapturedChild = child;` + `host.CaptureMouse(this)`. If the child's handler synchronously rebuilds the panel via `ScrollPanel::ClearChildren` (which nulls `_innerCapturedChild` at `:6232` then frees children via `Panel::ClearChildren` at `:6236/1084`), the freshly-cleared pointer is overwritten with the now-freed `child` and capture is armed. The next `WM_MOUSEMOVE` dereferences it: `_innerCapturedChild->OnMouseMove(...)` (`:6693`) → UAF. Reachable concretely: a `TabControl` in a `ScrollPanel` whose select handler (`SelectTab → onSelectionChanged`, `:5221`) rebuilds the panel. The author already used the *safe* ordering in `OnMouseUp` (`:6764-6766`: snapshot + null *before* dispatch); `OnMouseDown` does the opposite.

**Fix:** Set `_innerCapturedChild`/capture *before* dispatch and clear on `false`; or re-validate `child` against `GetChildren()` before committing. Have any teardown null the pointer last.

---

**A4. NativeMenuBarHost::OpenPopup null-derefs `_menuBar` after the nested `ContextMenu::Show` loop tears down the host window**
`Common/DxUi/DxUiNativeMenuInterop.h:751`
**Severity: use-after-free (null-deref of torn-down member) — CONFIRMED.**

`OpenPopup` guards `_menuBar` at entry (`:676`), runs the blocking nested message loop `ContextMenu::Show` (`:749`), then unconditionally calls `_menuBar->SetSelectedIndex(std::nullopt)` (`:751`). If the owner viewer window (ViewerWeb/Text/Space/ImgRaw/PE — all parent a `_menuBarHost` to the viewer hwnd) is closed during the loop (Alt+F4, tab/document close dispatched inside the loop, app shutdown), the child host receives `WM_NCDESTROY` and the handler nulls `self->_menuBar` (`:786`/`:808`). The C++ object survives (only members nulled), so `:751` dereferences a null `_menuBar` → crash. Every sibling method guards `_menuBar`; only `:751` does not.

**Fix:** After `Show` returns, re-check teardown before touching members: `if (!_menuBar || !_hwnd || IsWindow(_hwnd.get()) == FALSE) return;`, mirroring the entry guard. Apply the same pattern at every member access after a nested-loop call.

---

**A5. Tree double-click / context-menu reuses a stale visible index after a reentrant model mutation**
`Common/DxUi/DxUi.Tree.cpp:799` (double-click), `:1019` (context-menu)
**Severity: crash (OOB / null model deref) — verified verdict PLAUSIBLE; reachability latent today.**

`OnMouseDoubleClick` calls `SelectVisibleIndex(hit.visibleIndex, true)` (`:796`), which runs `_delegate->OnTreeSelectionChanged(...)` synchronously, then re-reads `_model->GetVisibleItem(hit.visibleIndex, item)` (`:799`) with the *pre-mutation* index and **no null/bounds re-check**. A delegate that rebuilds or nulls the model leaves `:799` indexing a shorter list (`std::out_of_range` in `.at()`-based models) or calling through a null `_model`. Same shape at `:1019`. This breaks the file's own invariant — `SelectVisibleIndex`, `ToggleExpanded`, `RequestSelectVisibleItem` all guard `if (!_model || visibleIndex >= GetVisibleItemCount())` immediately before indexing. *No current shipping delegate triggers it* (production Preferences `OnTreeSelectionChanged` only updates page text; its model `GetVisibleItem` is bounds-guarded), so this is a latent contract gap, not a reproducible crash today — but it is a real missing guard inconsistent with every sibling, and B-S0-1 does **not** cover these two Tree-internal re-reads.

**Fix:** Re-validate after `SelectVisibleIndex`: `if (!_model || hit.visibleIndex >= _model->GetVisibleItemCount()) { Invalidate(host); return true; }` before `:799` and `:1019`. Better: capture `TreeItemData` (id) *before* notifying and re-resolve by id.

---

### B. Robustness (latent reentrancy / device-lost / exception-safety)

These were downgraded from the reporters' initial UAF/crash labels by verification — the structural hazard is real but not reachable by a concrete current input, or the abort is clean-by-design. They remain high-leverage because they harden the same reentrancy/lifetime surface as Section A.

---

**B1. PageHost::SetPage destroys the outgoing page synchronously in the non-animated path**
`Common/DxUi/DxUi.Controls.cpp:1301`
**Severity: robustness (latent reentrancy UAF) — verdict PLAUSIBLE.**

Non-animated branch: `_currentPage->PropagateHost(nullptr); _currentPage = std::move(page);` (`:1299-1301`) runs the outgoing page's destructor *on-stack* when navigation is triggered from within that page. The animated branch deliberately stashes it in `_outgoingPage` (`:1313`) for safety; the non-animated branch has no such protection. The concrete `Button`-click scenario does **not** fault today only because controls are reentrancy-hardened (`Button::OnMouseUp` copies callback, Invalidates before, returns a local) — i.e., safety rests entirely on every control never touching `this` after its callback, with **no guard at the PageHost layer**. Note: this is the exact mechanism that makes **A2 (RadioButton)** an actual UAF.

**Fix:** Mirror the animated path — move outgoing `_currentPage` into a holding member / local and reset it after the current input dispatch completes (FinishTransition / next Tick), so a page that initiated navigation is never destroyed while its handler is on the stack.

---

**B2. MenuBar::ActivateItem calls `RequestInvalidate()` (reads `_host`) after the open-item callback may have freed the MenuBar**
`Common/DxUi/DxUi.Controls.cpp:4322`
**Severity: robustness (latent reentrancy UAF) — verdict PLAUSIBLE.**

`_onOpenItem(...)` (`:4321`) then `RequestInvalidate()` (`:4322`) → `_host->Invalidate()` (`DxUi.cpp:425-427`). If the callback freed the MenuBar, `_host` is read from a dangling `this`. No current `_onOpenItem` handler frees the MenuBar synchronously (both production handlers run `ContextMenu::Show` and themselves deref `_menuBar` afterward; rebuilds are deferred via `PostMessageW(WM_COMMAND)`), so not reachable today — but the post-callback member access is fragile.

**Fix:** Touch no instance member after `_onOpenItem`. Invalidate via the stable `WindowHost& host` argument already in scope (`host.Invalidate()`), or snapshot and return immediately.

---

**B3. Control virtuals invoked from `noexcept` boundaries `std::terminate` on any throw**
`Common/DxUi/DxUi.WindowHost.cpp:3278` (Render Paint), `:2420-2421/2505/2555/2583` (HandleMessage input), `:3910` (OnAnimationTick → Tick)
**Severity: robustness (exception-safety) — CONFIRMED; reclassified from "crash".**

`Render`, `HandleMessage`, `OnAnimationTick` are `noexcept`, but the `Paint`/`OnMouseDown`/`OnKeyDown`/`OnChar`/`Tick` virtuals they call are not (`DxUi.h:1286-1298`). Any throw (e.g. `std::bad_alloc` from `std::format`/`std::wstring` in a virtualized grid row's Paint under memory pressure — `DxUi.Grid.cpp` has ~45 such allocs) crosses the `noexcept` boundary → `std::terminate` → hard process abort with no logging. The only try/catch in the 3921-line file (`TraceWindowHostDiagnostics:352-364`) already models the degrade-vs-terminate tradeoff. This is a clean abort (no corruption), so robustness, not memory-safety — but it converts any control's recoverable OOM into a whole-app crash.

**Fix:** Wrap each `noexcept` dispatch site (Render's draw block, each `HandleMessage` input dispatch, `OnAnimationTick`) in `try/catch (const std::exception&)`, log via `Debug::Error`, and degrade the frame — applying the existing degrade path to the real dispatch sites.

---

**B4. Device-removed/reset recovery never re-invalidates — static windows stay blank**
`Common/DxUi/DxUi.WindowHost.cpp:3329-3338` (Render), `:3541-3550` (ENABLE_TESTS Render)
**Severity: robustness — CONFIRMED.**

On `DXGI_ERROR_DEVICE_REMOVED/RESET` / `D2DERR_RECREATE_TARGET`, both Render overloads discard all device resources and `return;` with **no `Invalidate()`**. Render was called from `WM_PAINT` inside `ScopedPaint` (BeginPaint/EndPaint validates the update region), so a non-animating window (modal dialog, preferences host, static dual-pane view) receives no further `WM_PAINT` until an unrelated event (mouse-move, focus change, resize, theme switch) reinvalidates it — leaving an indefinitely blank composition surface. `DebugSimulateDeviceLoss` (`:2106-2112`) runs the identical discard trio but deliberately calls `Invalidate()` afterward, confirming the omission.

**Fix:** Add `Invalidate()` at the end of the device-lost branch in both Render overloads, immediately after `ResetSharedWindowHostGraphicsResources()`, mirroring `DebugSimulateDeviceLoss`.

---

### C. Cross-thread race / Data race

---

**C1. UIA provider read methods race against UI-thread tree/model mutation (dangling `GetChildren()` span)**
`Common/DxUi/DxUi.Accessibility.cpp:3500` (and all read methods walking the live tree)
**Severity: use-after-free + data race — CONFIRMED. Project item B-S0-2 (HIGH, unchecked).**

Providers return only `ProviderOptions_ServerSideProvider` (`:3027`) — no `UseComThreading` — so UIA invokes read methods (`Navigate`, `GetPropertyValue`, `get_BoundingRectangle`, `GetSelection`, …) on arbitrary RPC worker threads. *Writes* are correctly marshaled to the UI thread via `SendMessageTimeoutW` (`:5409`), proving reads are not. Read methods take `GetAccessibilityTargetMutex()` and then walk `Panel::GetChildren()` — a `std::span` over the **live** `_children` vector (`DxUi.h:1406`) — and read live grid/tree models (`BuildGridRowAccessibleName → model->GetCellData`). **The mutex is referenced only within `DxUi.Accessibility.cpp`; no control-tree mutation, layout, or model code acquires it.** `Panel::AddChild`/`RemoveChild` take no lock; a reallocating `push_back` frees the span's backing storage while a worker thread iterates it → dangling `unique_ptr` deref → heap corruption / UAF. Reachable by any Narrator/Inspect client during normal navigation or list refresh.

**Fix:** Establish one lock discipline — all UI-thread tree/model mutation and layout must acquire the same mutex the providers hold, **or** marshal every provider read onto the UI thread (as writes already are). A copied `GetChildren()` vector alone is insufficient because deeper model reads remain unsynchronized.

---

### D. Denial-of-service / UI-thread stall

---

**D1. [MERGED] TSF `GetTextExt` multiline path does O(n) (effectively O(n²)) DirectWrite work per call, on the UI thread inside the TSF lock**
`Common/DxUi/DxUi.TextStoreACP.cpp:110` (`TryResolveMultilineTextStoreRangeRect`), reached from `GetTextExt:696-699`
**Severity: dos-oom — CONFIRMED (two reports merged).**

`GetTextExt` clamps the range only to `text.size()` (`ClampAcpRange`, no width cap), then for multiline loops `for (size_t index = range.start;; ++index)` from start to end **inclusive**, calling `control.TryGetTextInputCaretRect(host, index)` per code unit. Each call routes to `MeasureMultilineCaretRectDip`, which builds a **fresh, uncached** `IDWriteTextLayout` over the whole text and runs `HitTestTextPosition` every call (`DxUi.TextInput.cpp:773-792`) — so a full-document range is O(n) layout creations × O(n) each ≈ **O(n²)** DirectWrite work, synchronously under the TSF read lock. TSF calls `GetTextExt` routinely to position candidate/reading windows; an IME or screen reader over a multi-thousand-char multiline field stalls the UI thread for seconds, freezing input and rendering (CLAUDE.md forbids blocking the UI thread). The single-line path (two `MeasureCaretOffsetDip` calls) confirms the asymmetry.

**Fix:** Resolve the range rect in O(1)/O(visible-lines): measure only the start and end caret rects and union them, or reuse `IDWriteTextLayout::HitTestTextRange` on the cached multiline layout (as `TextField::GetTextInputRangeRects` already does). Optionally short-circuit to `TS_E_NOLAYOUT`/clipped rect above a threshold.

---

**D2. O(n²) grapheme counting on every masked-field repaint (`StepToNextTextElement` rescans from index 0)**
`Common/DxUi/DxUi.SingleLineTextEditing.cpp:213` (also `StepToPreviousTextElement`)
**Severity: dos-oom — CONFIRMED.**

`StepToNextTextElement` ignores `caretIndex` as a start hint — `size_t cursor = 0u;` (`:232`) and walks boundaries from the string start every call → O(n). `CountTextElements` calls it once per grapheme (`DxUi.TextInput.cpp:39-44`) → O(n²). `CountTextElements → GetSecretVisibleDotCount → GetDisplayText` is invoked on essentially every paint/measure/hit-test of a masked field (~14 sites: `DxUi.TextInput.cpp:1358,1448,1490,1754,…`). There is **no `MaxLength`** on the control (`ReplaceSelectionAndNotify` inserts with no cap, `:1031`) and **no caching** of the element count. Pasting a ~50–100KB string into a masked field makes every repaint (caret-blink timer, drag-select) do billions of boundary comparisons → multi-second-to-minute hang. (Adversarial input, so impact severity is conditional, but the algorithm and call chain are confirmed.)

**Fix:** Start `StepTo{Next,Previous}TextElement` scans near `caretIndex` instead of 0 (reducing each call to O(grapheme size)); cache the grapheme count per `_text` in `TextField` (invalidate on mutation) so `GetSecretVisibleDotCount` does not recount on every paint; and impose a sane maximum text length.

---

**D3. `GetMenuItemLayoutRects` is O(n) per item during paint → O(n²) paint for long/dynamic menus**
`Common/DxUi/DxUi.Menu.cpp:2516` (`GetItemRect` rewalk at `:2280-2284`)
**Severity: dos-oom — verdict PLAUSIBLE (quadratic confirmed; DoS impact conditional on item count).**

`MenuContentControl::Paint` iterates all items (`:2481`, no virtualization/off-screen skip) and per header/slider/info/standard item calls `GetMenuItemLayoutRects → GetVisibleItemRect → GetItemRect`, which re-walks from index 0 summing row heights every call → sum(0..n) = O(n²) per repaint, repeated on every hover/scroll `InvalidatePopup`. The Paint loop already tracks the running `y` offset that `GetItemRect` redundantly recomputes. Real for thousands of dynamically generated items (file/bookmark/history lists), negligible for typical menus — no thousand-item call site was located, hence the conditional severity.

**Fix:** Reuse the running `y` already tracked in the Paint loop (pass the precomputed top into the layout helper), or cache cumulative per-item tops in `MenuPopup` computed once at layout time and reused by `GetItemRect`/hit-test. Avoid re-walking from index 0 per item.

---

### E. Incorrect-behavior / Accessibility / Contract

---

**E1. `ReplaceSelectionAndNotify` mutates the buffer without syncing the native TSF/IME ACP cache**
`Common/DxUi/DxUi.TextInput.cpp:1020`
**Severity: contract (ACP/TSF state desync) — CONFIRMED.**

Every sibling mutator re-publishes the buffer/offsets to the native session: `SetText` (`:954-957`), `SetSelectionRange` (`:1013-1016`) both call `host->SyncTextInput(this)` under a focus guard; input-event paths are synced by the central dispatcher (`RouteFocusedCharInput`, WM_KEYDOWN). `ReplaceSelectionAndNotify` (a *programmatic* mutator outside input dispatch) updates `_text`/`_caretIndex`, invalidates, calls `NotifyChanged()` — but **never** `SyncTextInput`. The `WindowHost::_nativeTextInputStateCache` (served verbatim to TSF/ITextStoreACP via `TryReadNativeTextInputState`) keeps pre-edit text and offsets. An in-progress composition or candidate-window `GetTextExt` then maps offsets against the wrong length → garbled commit or `TS_E_INVALIDPOS`. Two callers only avoid this by manually re-syncing afterward; `BatchRenameWindow.cpp:4096-4101` re-syncs *only* inside `if (applied.text == targetField.GetText())`, leaving the cache stale otherwise.

**Fix:** Make `ReplaceSelectionAndNotify` self-consistent: after updating `_text`/`_caretIndex` and resetting the anchor, add the focus-guarded `host->SyncTextInput(this)` that `SetText`/`SetSelectionRange` use, and remove the now-redundant manual syncs at call sites.

---

**E2. IME composition preview committed into the control when the session is torn down on non-`WM_IME_ENDCOMPOSITION` focus loss**
`Common/DxUi/DxUi.NativeTextInput.cpp:593`
**Severity: incorrect-behavior (data integrity) — CONFIRMED.**

`applyCompositionPayload` imports the *uncommitted* preview into the control but deliberately leaves `_nativeTextInputImeBaseState` holding only pre-composition text (`:1246-1257`). The correct abandon path (`WM_IME_ENDCOMPOSITION → restoreCompositionBaseState`, `:1363-1369`) re-imports that base state. But `DeactivateNativeTextInputSession` (`:586/593`) and `OnKillFocus` call `ClearNativeTextInputCompositionState` — which `reset()`s the base state (`:822`) **without** restoring it. Reachable without an OS `WM_IME_ENDCOMPOSITION`: `SetTextInputBackend` switch, `PruneStaleInteractionState` (control goes non-interactive mid-composition, `WindowHost.cpp:3715`), `ResetRootInteractionState`. Result: half-typed kana/pinyin stays in the field as if committed → corrupted field value.

**Fix:** In `DeactivateNativeTextInputSession` (and any teardown clearing composition state), if `_nativeTextInputImeComposing`, call `restoreCompositionBaseState()` *before* `ClearNativeTextInputCompositionState()`, mirroring the `WM_IME_ENDCOMPOSITION` abandon path.

---

**E3. UIA tree-item runtime IDs encode the volatile visible index, not a stable id**
`Common/DxUi/DxUi.Accessibility.cpp:3818`
**Severity: incorrect-behavior / accessibility — CONFIRMED.**

`GetRuntimeId` for `TreeItem` calls `SetTreeItemRuntimeId(..., _treeVisibleIndex)` (a position in the virtualized list). Every TreeItem op (`get_BoundingRectangle:3887`, `ResolveTreeItemData:5249`, `ExecuteExpandOnWindowThread:5768`, Select/SetFocus) dereferences `_treeVisibleIndex` against the *current* `GetVisibleItemCount()` with only a bounds check. Grid rows deliberately use a stable id (`SetGridRowRuntimeId` + `FindRowByStableId`), proving the asymmetry. When a node above a retained TreeItem is expanded/collapsed, its runtime id changes (violating UIA's stable-runtime-id contract, breaking caching/focus tracking) **and** a later `Expand()/Select()/SetFocus()` from the OS-retained provider silently acts on whatever logical item now occupies that index. Bounds-checked, so no crash — silent mis-targeting.

**Fix:** Encode `TreeItemData::id` in the tree-item runtime id (mirroring grid rows) and resolve providers via `FindVisibleItemById(storedId)` rather than `_treeVisibleIndex`. The primitives already exist (`TreeItemData::id`, `FindVisibleItemById`).

---

**E4. `ExpandToEnclosingUnit` does not expand for Character/Word/Line — degenerate range, silent reader**
`Common/DxUi/DxUi.Accessibility.cpp:2459`
**Severity: accessibility / contract — CONFIRMED.**

`ITextRangeProvider::ExpandToEnclosingUnit` handles only `TextUnit_Document` (0..size). For Character/Word/Line it merely re-clamps the existing endpoints (`ClampCurrentRange`), never growing a collapsed caret range. A collapsed range yields empty `GetText` and empty `GetBoundingRectangles`, so Narrator per-word/per-character navigation reads nothing. The needed helpers exist and are used in `Move` (`GetTextRangeWordSpanAtPosition:950`, `GetTextRangeLineSpanAtPosition:1076`) but are not invoked here.

**Fix:** For Word use `GetTextRangeWordSpanAtPosition(text, _rangeStart)`; for Line use `GetTextRangeLineSpanAtPosition`; for Character snap to a grapheme boundary and extend by one element; assign the span to `_rangeStart`/`_rangeEnd`.

---

**E5. Slider mouse drag/click ignores `_step` and tick marks (continuous value vs quantized keyboard nav)**
`Common/DxUi/DxUi.Controls.cpp:3321`
**Severity: contract — CONFIRMED.**

`UpdateValueFromPoint` computes `nextValue = _minimum + clamp(normalized) * (_maximum - _minimum)` and passes it straight to `SetValueInternal` (clamps only, no snap). All three mouse handlers (`OnMouseDown:3392`, `OnMouseMove:3403`, `OnMouseUp:3416`) route through it. Keyboard moves strictly by `_step`/`_largeStep` (`OnKeyDown:3437/3451`). `SetStep` clamps step to `>=0.0001` (`:3169`), confirming step is meant to be honored. So a slider with `SetStep(1.0)` emits e.g. `3.7194` on drag — consumers indexing a discrete option list or formatting integers get wrong values. No crash; inconsistent value domain by input device.

**Fix:** In `UpdateValueFromPoint`, quantize to `_step` (round `(nextValue - _minimum)/_step`, multiply back, clamp) and optionally snap to nearest tick when `_tickMarks` is non-empty, matching the keyboard path.

---

**E6. Masked multiline TextField paints/selects control-text offsets against a shorter masked display string**
`Common/DxUi/DxUi.TextInput.cpp:1499`
**Severity: incorrect-behavior — CONFIRMED.**

`SetMasked` and `SetMultiline` are independent setters with no mutual exclusion. `GetDisplayText()` masks unconditionally, and `GetSecretVisibleDotCount` derives length from grapheme count (and bucketized Concealed policy), so the dot string can be strictly shorter than `_text.size()`. The multiline Paint branch hit-tests `GetSelectionRange()`/`_caretIndex` (control-text coords) against the masked display layout (`DrawMultilineSelection:1499`, `MeasureMultilineCaretRectDip:1536`). DirectWrite clamps over-range args (the buffer is sized from the larger control range), so **no OOB write** — but the highlighted dots and caret don't correspond to the actual selected characters, leaking incorrect structural feedback about a secret value.

**Fix:** Map selection/caret offsets from control-text space into display (dot) space before hit-testing, clamp `selectionStart/End` to `displayText.size()` in `DrawMultilineSelection`, or disallow masked+multiline at the API level (mirroring how the reveal/clear buttons already require `!_multiline`).

---

**E7. Double-click word selection uses raw UTF-16 code-unit indices — can split a surrogate pair**
`Common/DxUi/DxUi.SingleLineTextEditing.cpp:797`
**Severity: incorrect-behavior (data corruption) — CONFIRMED.**

`SelectSingleLineWordAt` expands purely over `GetWordSelectionClass(text[i])` per code unit (`:809/814-818`); `GetWordSelectionClass` classifies a lone surrogate via `GetStringTypeW` on one isolated `wchar_t`, so the boundary can fall between the two halves of a surrogate pair. It never snaps to text-element boundaries (unlike the arrow-key/backspace paths that use `StepTo*TextElement`). Double-clicking an emoji word then pressing Delete runs `DeleteSelection` first, erasing only one half → a lone unpaired surrogate in `_text` that flows to DWrite, the clipboard, and the TSF ACP store → rendering glitches and inconsistent ACP offsets. `std::wstring::erase` with valid in-range offsets is safe, so corruption, not a crash.

**Fix:** After computing `selectionStart/End`, snap both to grapheme/text-element boundaries (`SnapCaretIndexToTextElementBoundary` / `StepTo{Previous,Next}TextElement`), matching the keyboard paths.

---

### F. Robustness / Contract (cosmetic, latent)

---

**F1. Tree expansion-animation index mapping silently requires node-id stability across model rebuild**
`Common/DxUi/DxUi.Tree.cpp:532`
**Severity: robustness / incorrect-behavior — verdict PLAUSIBLE (cosmetic-only, contingent on a hypothetical model).**

Before/after snapshot rows are matched purely by `item.id` via `FindVisibleItemIndexById` (`:532-534,564-575`). `NotifyDataChanged` only clears the animation on a size mismatch (`:344`), so a size-preserving id renumber is undetected. A model whose ids encode position (or are reassigned on `Rebuild`) makes rows teleport/render at stale offsets during expand/collapse. All index derivations stay snapshot-bounded (verified), so **visual corruption only**, no crash. The "id is stable" contract is undocumented (`TreeItemData::id`, `IDxTreeModel` carry no such comment) yet silently required by selection, `FindVisibleItemById`, and toggle.

**Fix:** Document and enforce node-id stability as a hard `IDxTreeModel` contract; in `NotifyDataChanged`, also clear the expansion animation when the animated `itemId` is no longer findable or when before/after id multisets diverge beyond the toggled subtree; fall back to a non-animated update when id matching fails.

---

## 3. Cross-cutting Themes (highest-leverage fixes)

1. **"Dereference-after-callback" on raw `Control*` — the dominant defect class.** Affects the central dispatcher (**A1**) and individual controls (**A2 RadioButton**, **A3 ScrollPanel**, **A4 NativeMenuBarHost**, **B1 PageHost**, **B2 MenuBar**, **A5 Tree**). The codebase already has the correct discipline in places (`Toggle::OnMouseUp`, `ScrollPanel::OnMouseUp`, `ControlBelongsToTree` pre-dispatch guard at `WindowHost.cpp:2406`) — it is applied **inconsistently**. **Highest-leverage fix:** (a) a single post-callback re-validation pattern in the dispatcher (`ControlBelongsToTree(_root.get(), target)` after every `OnMouseDown/Up/DoubleClick`), and (b) a control-side rule enforced by review/lint: *copy callbacks to locals, invoke them last, and touch no member or `this`-derived pointer after any callback that can reenter*. Consider a per-control lifetime/generation token the dispatcher can cheaply re-check, so individual controls don't each have to be perfectly hardened. This single theme closes A1–A5, B1, B2.

2. **One lock discipline (or full UI-thread marshaling) for the UIA provider (C1).** Reads currently walk the live tree/model off-thread under a mutex no mutator holds. The write path already marshals to the UI thread via `SendMessageTimeoutW` — extending that to reads (or taking the accessibility mutex in every mutation/layout path) is the structural fix. This is the only cross-thread defect but it is reachable by any running screen reader.

3. **TSF/ITextStoreACP must never do unbounded work under a document lock (D1) and the ACP cache must stay in sync with every buffer mutation (E1).** The push-based native-input cache model is sound; the leak is *programmatic* mutators (`ReplaceSelectionAndNotify`) bypassing the publish contract, and the range-extent path iterating per code unit instead of using `HitTestTextRange`. Audit every ITextStoreACP method for (a) O(1)/O(visible) bounds and (b) cache freshness.

4. **Grapheme/text-element boundary discipline is applied to keyboard paths but not mouse/word-selection paths (E7), and grapheme counting is recomputed quadratically (D2).** A shared, caret-anchored, cached text-element iterator used by *all* selection/measurement paths fixes correctness (surrogate splitting) and performance (O(n²) masked repaint) together. Add a `MaxLength` to bound adversarial input.

5. **Tree virtualization identity uses volatile visible indices in two subsystems (E3 UIA, F1 animation).** Grid already uses stable ids end-to-end; tree should mirror it (`TreeItemData::id` + `FindVisibleItemById`) for runtime-ids and animation matching, and the id-stability contract should be documented on `IDxTreeModel`.

6. **`noexcept` boundaries and device-lost recovery are robustness cliffs (B3, B4).** A small, uniform "degrade the frame and reschedule a paint" wrapper at the three `noexcept` dispatch boundaries plus an `Invalidate()` in the device-lost branch converts two whole-app failure modes (terminate / indefinite blank window) into recoverable ones.

---

## 4. Coverage & Residual Risk

**What was statically verified vs. what needs interactive/stress testing:**

- **Reentrancy UAF cluster (A1–A5, B1, B2):** The dispatcher down-path UAF (**A1**) is "RED under AppVerifier/ASAN today" per the project's own plan and should be reproduced under **AppVerifier + ASAN with click-handlers that navigate/replace the root** (file-manager tree/grid item invoke, toolbar pane swap, ComboBox commit). **A2/A3** need targeted repros (radio group driving page switch; TabControl-in-ScrollPanel whose select handler rebuilds the panel). **A5/B1/B2** are latent today — they need a deliberately adversarial delegate/handler to fault, so they won't surface in normal QA; cover them with unit tests that mutate the model/tree from inside the callback.

- **TSF/IME (D1, E1, E2):** Requires **interactive IME testing** (Japanese/Chinese/Korean): in-progress composition + programmatic insertion (E1), backend switch / focus-steal mid-composition (E2), and candidate-window placement over a large multiline buffer (D1). These cross the **OS-side TSF trust boundary** — TSF/IME and AT clients are out-of-process and can issue arbitrary, pathological `GetTextExt`/lock sequences; the framework must treat all ACP offsets and range widths as adversarial input (clamp + bound work), which D1's unbounded loop currently does not.

- **UIA (C1, E3, E4):** Requires **live screen-reader testing** (Narrator + Inspect/Accessibility Insights) under concurrent UI mutation. **C1** specifically needs a **stress race**: a UIA client walking children while the UI thread rapidly adds/removes rows (list refresh, scroll) — best caught with ASAN + a tight mutation loop, as the window between span acquisition and reallocation is small. **E3** needs expand/collapse-above-a-retained-element scenarios; **E4** needs per-word/per-character Narrator navigation. UIA is also an **OS-side trust boundary**: providers run on RPC worker threads the app does not control, so thread-safety cannot be assumed away.

- **Device-lost (B4):** Requires **interactive device-removed testing** — TDR injection (`dxcap`/`DDI`), driver update, RDP connect/disconnect, lock-screen transition — specifically on a **static, non-animating window** (modal dialog, preferences host) where no incidental paint trigger hides the missing `Invalidate()`. An animating window will mask the bug.

- **DoS (D2, D3):** Stress with adversarial input sizes — paste ~50–100KB into a masked field (D2), generate a several-thousand-item dynamic menu (D3) — and observe UI-thread responsiveness. Severity is input-size-dependent; both are confirmed-quadratic in code.

- **Exception-safety (B3):** Best validated with **fault injection** (allocation-failure hooks) under low-memory to confirm the `std::terminate` path and the post-fix degrade behavior.

**Residual risk note:** This audit is static-analysis-based. The two highest-severity items (A1 dispatcher UAF, C1 UIA UAF) are corroborated by the project's own independent adversarial review (Operation Bedrock B-S0-1 / B-S0-2) and remain unchecked — they are real and unfixed. The latent items (A5, B1, B2, F1) have no current triggering caller, so they will not regress visibly until a future handler/model adds the reentrant or id-renumbering pattern; they are best fixed defensively now and guarded by tests, since the cost of discovering them later (as a field crash) is high.

**Key files:** `Common/DxUi/DxUi.WindowHost.cpp`, `Common/DxUi/DxUi.Controls.cpp`, `Common/DxUi/DxUi.Accessibility.cpp`, `Common/DxUi/DxUi.TextInput.cpp`, `Common/DxUi/DxUi.TextStoreACP.cpp`, `Common/DxUi/DxUi.NativeTextInput.cpp`, `Common/DxUi/DxUi.SingleLineTextEditing.cpp`, `Common/DxUi/DxUi.Tree.cpp`, `Common/DxUi/DxUi.Menu.cpp`, `Common/DxUi/DxUiNativeMenuInterop.h`.

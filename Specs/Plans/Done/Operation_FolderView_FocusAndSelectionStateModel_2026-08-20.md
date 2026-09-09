# Operation FolderView Focus and Selection State Model

> **NON-NORMATIVE OPERATION RECORD.** The classified contracts under `Specs/` are authoritative. This file records the decisions, implementation sequence, verification, and closeout evidence for the FolderView focus/selection state-model change.

## Execution directive

Before changing production behavior, rewrite the owning normative specifications so they contain the complete current-item, selection, anchor, pointer, drag-source, command-target, refresh, filtering, navigation-memory, and accessibility model defined by this operation. Resolve conflicting legacy clauses as one coherent contract; do not leave required behavior solely in this WIP plan.

After that spec-first rewrite, implement the behavior in the order defined by this plan and keep code, deterministic tests, performance evidence, and the authoritative specifications synchronized throughout execution. When a discovered implementation constraint requires a behavioral decision, update the normative contract and this plan before continuing implementation rather than silently choosing behavior in code.

## Master implementation checklist

- [x] **Phase 0:** Rewrite every owning normative specification and eliminate conflicting legacy clauses before production code changes.
- [x] **Phase 1:** Reconcile drift, add bounded test/perf observations, and archive an executable same-machine baseline.
- [x] **Phase 2:** Centralize current resolution, implement multi-removal successor/predecessor repair, and separate current-ownership epochs from selection.
- [x] **Phase 3:** Implement provider-aware bounded per-pane focus memory with deterministic dual-cap LRU eviction.
- [x] **Phase 4:** Implement selection, pointer/drag, Space/Insert, command fallback, filtering persistence, and the separate empty-parent action.
- [x] **Phase 5:** Complete and pass the deterministic correctness/invariant matrix.
- [x] **Phase 6:** Complete aggregate instrumentation, archive candidate evidence, compare performance, and make an evidence-backed budget decision.
- [x] **Phase 7:** Pass Fresh Full and documentation gates, record evidence, and move the completed plan to Done.

Phases are dependency ordered. Do not mark a phase complete until every checkbox in that phase and its verification gate passes. Do not start Phase 1 production/test-source edits until Phase 0 is complete.

## Status

- **State:** COMPLETE
- **Priority:** P1 correctness and interaction consistency
- **Risk:** Medium-high; the change touches refresh, filtering, sorting, navigation, pointer input, keyboard input, command targeting, and a FolderView hot path.
- **Estimated effort:** L; the change includes a normative rewrite, central model repair, bounded provider-aware memory, pointer/drag lifecycle, selection persistence, Space size-workflow preservation, new Commands cases, and perf evidence on a FolderView hot path.
- **Planned at:** `8313c1950e69f7ef1e5e057c5b0850381f410d33` (Phase 0 base, 2026-08-20; inspect the working-tree diff as well as commit drift)
- **Drift check:**

  ```powershell
  git diff --name-status 8313c1950e69f7ef1e5e057c5b0850381f410d33..HEAD -- `
    Specs/UI/UI_FolderView.md `
    Specs/UI/UI_CommandMenuKeyboard.md `
    Specs/Testing/Testing_SelfTests.md `
    Specs/Testing/Testing_PerformanceValidation.md `
    Specs/Testing/FolderViewPerfBudgets.json5 `
    RedSalamander/FolderView.h `
    RedSalamander/FolderView.cpp `
    RedSalamander/FolderView.Enumeration.cpp `
    RedSalamander/FolderView.Interaction.cpp `
    RedSalamander/FolderView.Selection.cpp `
    RedSalamander/FolderView.Rendering.cpp `
    RedSalamander/FolderView.Menus.cpp `
    RedSalamander/FolderView.DragDrop.cpp `
    RedSalamander/SelfTest/Commands
  ```

- **Ownership:** This plan is the sole owner of the FolderView current-item invariant, bounded per-pane focus memory, and the selection-state model defined here.
- **Coordination:**
  - Operation WarpDrive remains the owner of broader FolderView performance work.
  - `Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md` remains the owner of the already specified host-owned Delete/Move removal-focus repair. This plan incorporates that behavior as a specialized focus intent and must not replace its proven-removal safety rules.
  - Thumbnail enrichment remains out of scope.

## Decision state

The requester accepted the focus model on 2026-08-20 and asked that it be retained for later implementation. This plan also proposes and locks a coherent selection model so implementation does not need to infer behavior from contradictory existing clauses.

The 2026-08-20 hardening pass additionally locks these previously ambiguous product decisions:

- Space retains the established **Select + Calculate Directory Size + Next** command behavior, with toggle semantics on the captured current item: an unselected current item becomes selected, contributes its known file size or asynchronous folder-subtree size, and then current advances; an already-selected current item becomes unselected, is removed from size work, and current advances. The advance does not wrap. Insert toggles and advances without starting explicit folder subtree-size work.
- Generic disappearance chooses the first surviving successor under the resulting UI policy, otherwise the nearest predecessor; raw old-index reuse is not the contract.
- The standard empty-folder **Go to parent** cue remains actionable through Enter/double-click/Backspace but is not a real item or current-item exception.
- Empty-background selection clearing retains current and therefore does not invalidate removal-focus ownership; only an actual user current-identity change does.

The production behavior, deterministic coverage, Release performance evidence, and final repository gates are complete. The lasting behavior is recorded in the owning normative specifications.

## Remaining outcomes

1. A non-empty FolderView always has exactly one current item.
2. If the current item disappears, FolderView chooses its first surviving successor under the resulting visible UI order, or its nearest surviving predecessor when no successor exists; it does not blindly reuse the old numeric index.
3. Navigation remembers and restores current items across folders and file-system providers without unbounded memory growth.
4. Selection is an independent set containing from zero through all currently displayed real items.
5. Refresh, sort, filter, hide, pointer, keyboard, and command-target behavior follows one deterministic state model.
6. Focus migration never changes the target of a pending destructive or deferred command.
7. Deterministic correctness tests, performance instrumentation, and archived evidence protect the new behavior.
8. The lasting contract is moved into authoritative specifications before this plan moves to `Specs/Plans/Done/`.

## Scope

- FolderView's retained current-item model and its visual flag/index consistency.
- Current-item resolution after refresh, sort, filter, hide/show, and navigation.
- Per-pane current-item memory across folder and provider/context navigation.
- Selection membership, persistence, range anchor, pointer gestures, keyboard gestures, and command fallback semantics.
- Debug observations, deterministic Commands selftests, FolderView performance instrumentation, and archived closeout evidence.
- Required updates to the authoritative FolderView, keyboard, selftest, and performance specifications.

## Out of scope

- Persisting current-item memory or selection in settings or across process restarts.
- Changing Win32 pane focus, Tab traversal, address-bar focus, or application-reactivation policy.
- Redesigning provider enumeration, directory caching, sort algorithms, display layouts, scrolling, thumbnails, or icon enrichment except where they call the central current/selection resolver.
- Weakening or redesigning File Operations success, cancellation, epoch, containment, or immutable-target rules.
- Restoring selections on navigation history.
- Introducing a reusable repository-wide LRU library without a separate shared-helper review and a second semantically compatible consumer.

## Current authoritative state and implementation evidence

### Phase 0 conflict resolution

The operation opened with contradictory clauses for empty-background focus,
generic disappearance, Space, Shift ranges, and plain-click selection. Phase 0
rewrote `UI_FolderView.md`, `UI_CommandMenuKeyboard.md`, `UI_FolderWindow.md`,
the File Operations contract, and the owning testing/performance specifications
as one contract before production behavior changed. The final conflict scan and
spec-inventory gate verify that the superseded clauses are gone.

### Implementation outcome

- Every accepted non-empty model normalizes through one current resolver and exposes exactly one matching focused flag/index.
- Generic disappearance uses the production keyed comparator or prior Sort-None order to choose the first surviving successor, then predecessor.
- Per-pane current memory uses provider-aware location/item identity with deterministic O(1)-average LRU touch/eviction and 512-entry/512-KiB caps.
- Selection remains independent from current, is pruned by display membership, transfers only through proven rename chains, and is never restored by navigation memory.
- Pointer, keyboard, context-menu, drag, Space/Insert, selection-aware command fallback, and empty-parent behavior implement the transition tables below.
- Failed navigation restores the last accepted displayed model; successful navigation clears selection and notifies FolderWindow's cached selection/status state.
- Aggregate instrumentation and deterministic Commands coverage protect the invariant, repair reasons, cache bounds, state transitions, and performance direction.

## Target normative model

### 1. Terms and core invariants

- A **displayed real item** is an entry in the accepted, filtered, hidden-state-adjusted, and sorted FolderView item model. Busy, error, watermark, and empty-folder placeholders are not real items.
- The **current item** is FolderView's retained item cursor. It is distinct from Win32 keyboard focus and remains visible with the inactive-pane current-item treatment when the pane does not own keyboard focus.
- The **selection** is an independent set of displayed real items.
- If the displayed real-item count is greater than zero:
  - `_focusedIndex` identifies one valid displayed row;
  - exactly one `FolderItem::focused` flag is true;
  - both representations identify the same item.
- If the displayed real-item count is zero, no real item is current or selected. The standard empty-folder **Go to parent** cue is a separate empty-state action visual, not a `FolderItem` and not an exception to this invariant.
- Selection cardinality is always in `[0, displayedRealItemCount]`.
- The current item may be selected or unselected. Moving or repairing current item does not select it.
- Selection changes do not remove the current item. A current-item change occurs only when an input/navigation operation requests it or the current item is no longer displayable.

Implementation must centralize invariant normalization. Public model-replacement paths may compute a preferred current item, but they must finish through one operation that clears stale `focused` flags and establishes exactly one valid current item for a non-empty model.

### 2. Current-item resolution

#### Same-location model changes

“Same logical location” means equality of the logical file-system plugin ID, stable provider-instance/mount/connection context, and provider-normalized folder key produced through `FileSystemPathIdentity`. FolderView does not guess that alternate spellings, junctions, reparse paths, or remounted connections are equivalent. If stable recreation identity is unavailable, equality is scoped to the exact live instance context and exact ordinal path text.

For refresh, filter, hide/show, visibility-policy change, and any other rebuild of the same logical folder:

1. Honor a validated one-shot host intent, such as a successful rename/create target or the specialized host-owned Delete/Move removal-focus result.
2. Otherwise preserve the current item by provider-aware item identity when it survives.
3. Otherwise create a non-actionable repair probe from the prior current row and resolve its successor under the **resulting** directories-first sort, direction, filter, and visibility policy:
   - for an active keyed sort, use the exact production comparator and stable tie-breaker to find the first resulting row that orders after the probe;
   - for Sort None, scan the prior accepted order after the old current for the first provider-aware identity that survives in the resulting display.
4. If no successor exists, choose the nearest predecessor by the corresponding comparator/insertion position, or by reverse prior-order survivor scan for Sort None.
5. If no prior neighbor survives but the resulting display is nonempty, choose its first row. If there was no valid prior current context, also choose the first displayed row.

This defines “next item in the current UI” without blindly reusing the old numeric index. Removing several rows before or around current cannot skip the first surviving successor. The repair probe is resolver-only state: it is never selected, remembered, exposed through status/UIA, or accepted as a command target.

Additional rules:

- A pure sort change preserves the current item by identity across reordering.
- Passive remembered focus must not override a surviving current item during same-location refresh.
- If filtering or hiding removes the current item, the replacement becomes the remembered item for that location. Widening the filter later does not resurrect the older hidden current item.
- If a non-empty model reaches normalization without a valid candidate because of stale or corrupt state, the first displayed item is the fail-safe current item.
- After resolution, FolderView ensures the current item is visible using the existing viewport policy.

#### Navigation to another logical location

For folder, history, pane-provider, or provider-instance navigation:

1. Honor a validated explicit navigation target, if one exists.
2. Otherwise restore the remembered current item for the destination location when that item is present.
3. Otherwise choose the first displayed item.

The location's selection is not restored. Navigation begins with an empty selection even when current-item memory succeeds.

#### Empty and later populated displays

- An accepted empty model has no current item and no selection.
- When that logical display later becomes non-empty, normal resolution applies: explicit target, remembered item when entering a location, otherwise first displayed item.
- This empty-list exception is the only stable state in which FolderView has no current item.

The standard unfiltered empty-folder **Go to parent** cue has a separate contract:

- it is a row-shaped empty-state action visual and MUST NOT be stored in `_items`, `_focusedIndex`, `FolderItem::focused`, selection, anchor, focus memory, command fallback, or drag-source state;
- its visual focus treatment follows pane keyboard focus, but it does not count as a real current or selected item and does not appear in item/selection counts or status summaries;
- Enter while this cue is active and double-click anywhere in that standard empty-folder surface invoke the existing `NavigateUp()` route; Backspace remains unchanged;
- a filter-empty display, host-provided empty message, busy/cancel/error overlay, or background watermark does not acquire this parent action unless its owning normative contract explicitly enables it;
- if accessibility exposes the cue as a child action, it is an Invoke action rather than a SelectionItem and is never reported as the FolderView current item.

#### Specialized removal intent and command safety

- The existing host-owned Delete/Move contract remains more specific than generic refresh: only exact per-source `S_OK`, matching provider/folder/sort/focus epochs, and a newer accepted enumeration may authorize its positional replacement.
- The focus-ownership epoch represents ownership of the current-item identity, not selection ownership. Increment it when explicit user input changes current item; do not increment it for a selection-only mutation, including an empty-background click that retains current. Provider, folder, and sort invalidation epochs remain unchanged.
- Therefore, clearing selection on empty background does not invalidate an otherwise matching pending Delete/Move rehome. The operation's immutable source snapshot remains authoritative, and any later user gesture that actually changes current still invalidates the pending rehome.
- A current-item replacement remains unselected unless it already survived in the selection.
- Context menus and deferred/destructive commands continue to snapshot immutable item names/identities before nested UI or asynchronous work and revalidate them at dispatch.
- Automatic current-item repair must never retarget an already captured command to the replacement item.

### 3. Bounded focus memory

Focus memory is a per-FolderView-pane, process-lifetime convenience cache. It is not a setting and is not serialized.

#### Capacity

- Maximum remembered locations per pane: **512 entries**.
- Maximum tracked owned UTF-16 key/value payload per pane: **512 KiB**, counting the code units owned by provider id, instance context, normalized folder key, and remembered item name.
- Container/node overhead is additionally bounded by the 512-entry limit.
- An individual entry larger than the payload limit is not cached; it must not evict the entire useful cache.
- On insert/update, evict least-recently-used entries until both caps are satisfied.

#### LRU rules

- Successful restore lookup touches recency.
- Recording or updating a location touches recency and does not increase entry count for an equivalent key.
- Eviction is deterministic: least recently used first; ties follow the cache's stable insertion order.
- Cache size and accounted payload must remain within bounds after every successful mutating operation. Follow the repository rule that `std::bad_alloc` is fatal; do not catch it to continue with partially mutated cache state.
- Cache lifetime ends with the pane. Switching roots or providers does not globally clear it.

#### Identity key

The key comprises:

1. logical file-system plugin id;
2. provider instance context or mount/connection identity;
3. logical folder path normalized through the provider's stable `FileSystemPathIdentity` contract.

The value is the remembered leaf/display identity normalized and compared through the same provider component relation.

Rules:

- Never use a COM pointer, HWND, or transient object address as cache identity.
- Local Windows paths use the canonical ordinal-ignore-case path identity.
- A provider that exposes stable path-text identity uses `TryMakePathKey(...)` / `TryMakeComponentKey(...)`, with helper equality verification where a collision could affect behavior.
- If a provider cannot provide a stable path identity, memory is scoped to its exact live instance context and exact ordinal path/item text; it must not claim restoration across instance recreation.
- Changing provider or connection cannot collide with an identically spelled path in another provider/context.
- Missing remembered items are ordinary cache misses. Resolution falls back to the first displayed item and updates the entry after the accepted model is current.

This is intentionally FolderView-owned state, not a new generic repository LRU utility: the key, accounting, provider identity, and restoration rules are pane-specific. Re-run the shared-helper search before implementation; extract only if a semantically identical bounded cache has become canonical.

### 4. Selection model

#### Set membership and persistence

- Only displayed real items may be selected. Hidden, filtered-out, removed, placeholder, or stale identities are not latent selections and cannot be command targets.
- Navigation to another logical folder/provider/context clears selection. Returning through Back/Forward restores current item only, never selection.
- Same-location refresh and sorting preserve selection for surviving provider-aware item identities.
- A proven chained rename hint transfers selection from the selected original identity to the final identity. An unselected original never becomes selected because of a rename hint.
- Items removed, filtered out, or hidden are removed from selection. Clearing the filter or showing the item later does not restore that prior selection.
- Newly enumerated or newly visible items start unselected.
- **Select All is scoped to the displayed set at the instant it runs.** If Select All is followed by a filter that removes some selected items from the display, those filtered-out items leave the selection. Clearing that filter reveals those items as unselected; the FolderView is therefore no longer in a Select All state. The user must invoke Select All again to select the newly visible complete set.
- Current-item repair, memory restoration, sorting, and scrolling never add to or remove from selection except for removal of items no longer displayed.

#### Display-membership transition table

Selection is a concrete set, never a sticky Select All, mask, or invert mode. Apply these outcomes after an accepted transition:

| Transition | Required selection result |
|---|---|
| Accepted same-location refresh with the same identities | Preserve the selected surviving identities |
| Accepted refresh, Copy, Create, Paste, or provider update adds identities | Preserve selected survivors; every newly displayed identity is unselected, so a prior Select All is no longer Select All |
| Accepted refresh removes a selected identity | Remove it from selection immediately; an ordinary later reappearance is unselected |
| Select All → filter narrows → filter clears/widens | Items continuously displayed remain selected; reappearing items are unselected |
| Hide Selected Names → Show Hidden Names | Hidden selected identities leave selection and return unselected; selection may become empty |
| Hide Unselected Names → Show Hidden Names | Continuously displayed selected identities remain selected; the reappearing formerly unselected identities remain unselected |
| Hidden Files or System Files is disabled then re-enabled | Selected identities excluded by the visibility flag leave selection and return unselected |
| Select/Unselect mask, Select Same Name/Extension, or Invert while filtered | Mutate only identities displayed when the command runs; later-visible matching identities remain unselected |
| Successful Delete or Move removes sources | Remove exactly the proven-removed identities; selected surviving or failed/cancelled sources remain selected when their identities survive |
| Actual folder, history destination, provider, or provider-context navigation | Clear selection; returning later restores current item only |
| Accepted empty model followed by repopulation | Empty model has no selection; all repopulated identities begin unselected |
| Failed, cancelled, stale, or rejected enumeration/result | Preserve the current displayed model and its selection because no replacement model was accepted |
| Sort, display-layout change, Quick Search, scrolling, pane activation, or window focus change | Preserve selection exactly because displayed membership did not change |

The two explicit identity-transfer mechanisms are:

1. A proven same-folder rename chain transfers selected state from the selected source identity to its final target. Short-lived missing-selection bookkeeping may retain only the provenance needed for that proven chain. It is not active selection, must not appear in status/commands, and must never reselect an ordinary reappearing identity.
2. **Save/Restore Selection is an explicit user-requested exception to automatic no-resurrection.** Save records the currently selected displayed identities, or the current-item fallback when selection is empty. Restore replaces the live selection with saved identities that match items displayed at the instant Restore runs. Saved identities excluded by a filter/visibility rule remain unmatched and unselected; clearing the filter does not apply them later. Invoking Restore again after widening visibility is the explicit action that may select those now-visible matches.

#### Anchor

- The selection anchor is internal range-selection state, not a second focus.
- A non-range current-item move establishes the destination as the new anchor.
- A range gesture keeps the starting anchor and moves the current item to the range endpoint.
- At the start of a Shift pointer gesture, FolderView resolves the range anchor before changing current item for that button press. A valid prior anchor wins; otherwise the valid pre-press current item becomes the anchor. The inclusive range contains both that current item and the clicked endpoint.
- Ctrl+Shift uses the same anchor resolution but adds the inclusive range to the existing selection instead of replacing it.
- After a model rebuild that invalidates the old anchor, the resolved current item becomes the anchor.
- Empty displays have no anchor.

#### Pointer and drag target state

- Current item, selection, and the transient pointer/drag target are independent state dimensions. A pointer gesture may clear its item target without removing the keyboard-navigation current item.
- Every left-button-down that hits a displayed real item makes that item current during button-down handling, before drag initiation or later command routing can observe the gesture. FolderView first captures any pre-press current item needed for no-anchor Shift/Ctrl+Shift resolution, then moves current to the hit item.
- A left-button miss on empty background with no modifier, Ctrl, Shift, or Ctrl+Shift clears selection, clears the transient pointer/drag target, disarms any armed drag source, and resets the anchor to the retained current item. The modifiers have no endpoint on a miss. An empty display continues to have neither current item nor anchor.
- A pointer right-click miss performs the same selection-clear/retain-current/no-ownership-change transition, then opens the background context route with no item pointer target. A keyboard context-menu request is not a miss: it preserves selection, uses the selected-or-current fallback, and must not hit-test a synthetic pane-center point.
- Movement after an empty-background press cannot promote the retained current item or a stale selection into a drag source. Starting a later drag requires a fresh button press that hits an item and establishes a new pointer/drag target.
- Item context-menu and drag routing consume the gesture's established pointer target plus the command-target snapshot rules; they MUST NOT infer that an empty-background gesture targets the retained current item.
- An item left-button-down may arm a potential drag only for that gesture. Its source policy is the same selected-or-current fallback used by selection-aware commands: the selected displayed identities when selection is nonempty, otherwise the current item established by the press. The armed source is an immutable identity snapshot for that gesture. If a modifier transition leaves the pressed item outside that resolved source set, the gesture does not arm a drag; this prevents Ctrl+Click deselection from dragging some other remaining selection.
- Pointer movement starts the drag only after leaving the effective system drag-threshold rectangle centered on the button-down point. Movement inside that rectangle is still a click gesture. Mouse-up inside the rectangle completes the click without starting a drag. The item's visual bounds are not an additional drag threshold; a large row must not require traversing the whole item rectangle before dragging can start.
- Button release, capture loss, cancellation, teardown, navigation/model replacement, or a new pointer press disarms the prior potential drag. A drag can never inherit source state from an earlier gesture.
- A plain press on an already-selected item preserves the multi-selection and therefore arms that selected set. A plain press on an unselected item first replaces selection with that item and arms it. Modifier gestures apply their selection transition before the selected-or-current drag-source snapshot is captured.
- A double-click follows the first press's normal current/selection rules and activates the resulting current item; it does not bypass target establishment or reuse an earlier armed drag source.

This preserves the defensive intent of the prior empty-background rule—no stale drag or pointer-routed command target—without conflating that transient state with the persistent current item.

#### Input transition table

| Input in normal item-navigation mode | Current item after input | Selection after input | Anchor after input |
|---|---|---|---|
| Plain click on an unselected item | Clicked item | Replace with clicked item | Clicked item |
| Plain click on an already-selected item | Clicked item | Preserve existing selection | Clicked item |
| Ctrl+Click | Clicked item | Toggle clicked item only | Clicked item |
| Shift+Click with a valid anchor | Clicked endpoint, established on left-button-down | Replace with inclusive anchor-to-endpoint range | Preserve prior anchor |
| Shift+Click with no valid anchor | Clicked endpoint, established on left-button-down | Replace with inclusive pre-press-current-to-endpoint range, including both items | Pre-press current item |
| Ctrl+Shift+Click with a valid anchor | Clicked endpoint, established on left-button-down | Add inclusive anchor-to-endpoint range to existing selection | Preserve prior anchor |
| Ctrl+Shift+Click with no valid anchor | Clicked endpoint, established on left-button-down | Add inclusive pre-press-current-to-endpoint range, including both items | Pre-press current item |
| Shift+navigation | Range endpoint | Replace with inclusive anchor-to-endpoint range; if anchor is invalid, use and include the pre-move current item | Preserve or establish starting anchor |
| Ctrl+Shift+navigation | Range endpoint | Add inclusive anchor-to-endpoint range; if anchor is invalid, use and include the pre-move current item | Preserve or establish starting anchor |
| Unmodified Arrow/Home/End/Page navigation | Destination | Preserve selection | Destination |
| Quick-search current-item jump | Matching item | Preserve selection | Matching item |
| Alt+Up / Alt+Down | Previous/next selected displayed identity, with wrap | Preserve selection | Destination |
| Space | Next displayed item if one exists, otherwise unchanged; no wrap | Toggle the pre-advance current item (unselected becomes selected; selected becomes unselected), then recompute size from the post-toggle selected set | Resulting current item |
| Insert | Next displayed item if one exists, otherwise unchanged | Toggle the pre-advance current item | Resulting current item |
| Esc | Unchanged | Clear selection | Current item |
| Ctrl+A | Unchanged | Select all displayed items | Unchanged |
| Invert selection / extension masks | Unchanged | Mutate displayed items only | Unchanged |
| Empty-background left miss with any modifier combination | Unchanged | Clear selection and disarm drag | Current item |
| Empty-background pointer right-click miss | Unchanged | Clear selection; route background menu with no item pointer target | Current item |
| Keyboard context-menu request | Unchanged | Preserve selection; use selected-or-current command fallback | Unchanged |
| Right-click on an unselected item | Clicked item | Replace with clicked item before menu snapshot | Clicked item |
| Right-click on an already-selected item | Clicked item | Preserve selection before menu snapshot | Clicked item |

Transient modes own Escape first. For example, active incremental search consumes the first Escape to exit that mode; the selection-clear rule applies when FolderView is in normal item-navigation mode.

This table deliberately resolves current specification/implementation conflicts:

- plain click on an unselected item selects it as the sole target;
- item left-button-down establishes current immediately, while no-anchor Shift/Ctrl+Shift still use the captured pre-press current item as the inclusive range anchor;
- empty-background click clears selection but retains the current item;
- drag source and selection-aware commands use the same selected-or-current fallback, and drag initiation additionally requires an item press followed by movement outside the system drag-threshold rectangle;
- Space preserves the established **Select + Calculate Directory Size + Next** workflow;
- Insert toggles and advances without requesting folder subtree-size work;
- range navigation moves the current item to its endpoint, while current and selection remain independent state dimensions.

#### Space and Insert workflow

The product decision is locked as follows; implementation MUST NOT turn Space into toggle-in-place:

1. In normal item-navigation mode, Space captures the operated current item before movement, toggles that item's selection, requests selection-size recomputation from the post-toggle selected set, then advances current to the next displayed row when one exists. At the last row it remains there; it does not wrap, and the destination is not implicitly selected.
2. If the operated current item was unselected, Space selects it before recomputation. Its known file size contributes directly, or its folder subtree enters the existing asynchronous, cancelable size path. Current then advances while that operated item remains selected.
3. If the operated current item was already selected, Space deselects it before recomputation. Pending folder-size work for it is canceled, stale completion is discarded, and it is excluded from the new selected-size total. Current then advances.
4. The selection-size request always consumes an immutable snapshot of all post-toggle selected identities, not the advanced-to current item and not merely the operated item.
5. Insert performs the same toggle-then-advance sequence but does not initiate explicit folder subtree-size work.
6. The advance is a current-item move, so it updates anchor and current ownership normally without changing any other selected rows.
7. While Quick Search owns input, Space remains query text and performs none of the selection, size, or advance steps.

#### Command targeting

- Focus-only commands such as Enter/Open-current and F2/Rename consume the current item.
- Selection-aware commands such as Copy, Cut, Move, Delete, and selected-path actions consume the selection when non-empty and fall back to the current item only when selection is empty.
- A command must snapshot the identities selected by its target policy before opening a nested menu, prompt, or asynchronous operation.
- Later current-item or selection repair cannot alter that captured target set.
- Status-bar selection summaries report the selection only; an unselected current-item fallback is not displayed as selected.

### 5. Notifications, rendering, accessibility, and diagnostics

- Current-item and selection notifications are independent and emit only when their corresponding observable state changes.
- Current/path notifications refresh preview and current-item details. Selection-only notifications update selection statistics, status/command state, and size-work cancellation/recomputation but do not refresh an unchanged current-item preview or change current ownership.
- Every transition invalidates both the old and new current-item bounds as needed and cannot leave a stale `FolderItem::focused` flag.
- The focused and inactive-pane current border continues to render independently from selection fill.
- Debug snapshots used by behavioral tests must expose at least:
  - displayed real-item count;
  - current index and current display name;
  - count of `FolderItem::focused` flags;
  - current-resolution reason and focus-ownership epoch;
  - selected count and selected display names or a deterministic digest;
  - anchor index;
  - whether the separate empty-folder parent action is active;
  - whether a potential drag is armed and its immutable source count/digest;
  - focus-memory entry count, accounted payload bytes, and eviction count.
- UI Automation selection exposure must remain selection-based. The retained current-item cue must not be reported as selected merely because it is current.
- The empty-folder parent action is not the retained current-item cue. If exposed through UI Automation, it uses Invoke semantics and remains outside Selection/SelectionItem state.

## Implementation plan

### Phase 0 — Rewrite every owning normative contract

**Gate:** No production source file may change until every checkbox in this phase is complete. This phase replaces the split contract with the target model above; the executor must not treat the WIP plan itself as lasting authority.

Normative files in scope:

- `Specs/UI/UI_FolderView.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_FolderWindow.md`
- `Specs/UI/UI_VisualStyle.md` (consistency audit; edit only if its current/inactive visual language conflicts)
- `Specs/FileSystem/FileSystem_FileOperations.md` (host-owned removal-focus cross-contract)
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/Testing/FolderViewPerfBudgets.json5` (policy/inventory audit only in Phase 0; do not invent a numeric budget before Phase 6 evidence)

Checklist:

- [x] Run the Status drift command, `git status --short`, and `git diff --check`; reconcile overlapping I2/I12 edits before changing any normative clause.
- [x] Search all authoritative specs, not only the files already known, for `current item`, `focused item`, `selection`, `Space`, `Insert`, `Shift`, `empty folder`, `Go to parent`, `drag`, and `removal focus`. Add any newly discovered owning spec to this phase before editing.
- [x] Rewrite `UI_FolderView.md` as one coherent state contract covering: exactly-one-current for nonempty real models; zero current for empty real models; successor/predecessor repair; provider-aware bounded memory; independent selection; no resurrection; range anchor; left-button-down timing; pointer/drag lifecycle; command fallback; empty-background behavior; the separate empty-folder parent action; ownership epochs; notifications; rendering; accessibility; and diagnostics.
- [x] Delete or replace the `generic refresh ... leave focus unset` clause. State the exact comparator/probe successor rule and the multiple-removal behavior; do not describe repair as raw old-index reuse.
- [x] Delete or replace every `Shift ... without moving focus` clause. State that range endpoints become current, a valid anchor is preserved, and a missing anchor captures/includes the pre-gesture current item.
- [x] Rewrite all duplicated `UI_FolderView.md` input summaries so plain, Ctrl, Shift, Ctrl+Shift, right-click, empty background, double-click, Space, Insert, and drag semantics agree with the transition table.
- [x] In `UI_CommandMenuKeyboard.md`, retain **Select + Calculate Directory Size + Next** for normal-mode Space and explicitly cover both transitions: unselected current becomes selected/measured before advance, while selected current becomes deselected/removed from size work before advance. Retain toggle+advance without size work for Insert and Space as text while Quick Search owns input. Update the shortcut table, command detail, command/menu text, and folder-size section together.
- [x] In `UI_FolderWindow.md`, replace the incorrect statement that Insert explicitly starts folder subtree-size work. Specify that the Space workflow requests it and that post-toggle selection changes cancel/recompute against immutable selected identities.
- [x] In `UI_FolderView.md` and, where cross-domain wording is needed, `FileSystem_FileOperations.md`, state that the removal-focus ownership epoch changes only when user input changes current identity. Selection-only background clearing does not invalidate a matching pending Delete/Move rehome; folder/provider/sort epochs and exact-`S_OK` proof remain unchanged.
- [x] Specify the standard **Go to parent** cue as a separate empty-state Invoke action, never a real item/current/selection/anchor/drag/command fallback. Preserve Enter, double-click, Backspace, root callback, and the exclusions for filter-empty/host-message/busy/error surfaces.
- [x] Add the exact deterministic case names and assertions from Phase 5 to `Testing_SelfTests.md`, including multi-removal repair, Space size+advance, empty-parent action, empty-background epoch preservation, pointer/drag snapshots, selection visibility, and bounded LRU.
- [x] Add the metric/scenario/evidence contract from Phase 6 to `Testing_PerformanceValidation.md`. Record that a machine-keyed budget is added to `FolderViewPerfBudgets.json5` only after stable same-machine Release samples justify it.
- [x] Verify `UI_VisualStyle.md` still describes current-item visuals independently from selection and requires no contradictory change.
- [x] Keep mutable test counts source-derived. Do not paste current inventory totals into normative prose.
- [x] Run the conflict scan below and manually inspect every match; allowed matches may describe the new behavior, but no legacy winner may remain.
- [x] Run spec inventory and diff checks; resolve every blocking finding.
- [x] Confirm `git diff --name-only HEAD -- RedSalamander Common Plugins` prints nothing at the Phase 0 gate.

Phase 0 evidence: `Get-SpecInventory.ps1 -FailOnFindings` reported zero blocking findings; `git diff --check` passed; the production-source diff gate was empty. `UI_VisualStyle.md` already separates current-item border/fill from active/inactive selection and required no edit. `FolderViewPerfBudgets.json5` remains unchanged pending Phase 6 same-machine evidence.

Verification:

```powershell
rg -n -i "leave focus unset|without moving focus|clear selection, focus|select \+ calc|Space|Insert|Shift\+Click|Go to parent|drag-threshold|removal.focus|focus memory|512 KiB|512 entries" Specs/UI Specs/FileSystem Specs/Testing
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
git diff --name-only HEAD -- RedSalamander Common Plugins
```

Expected result: spec inventory reports zero blocking findings; diff check exits 0; the production-source diff command has no output; every search match agrees with the target model.

### Phase 1 — Reconcile drift, establish executable baselines, and add observations

Primary files:

- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.Selection.cpp`
- `RedSalamander/FolderView.Enumeration.cpp`
- `RedSalamander/FolderView.Interaction.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- the Commands registration in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

- [x] Re-run the Status drift command after Phase 0 and STOP if I2/I12 changed the same resolver, epoch, or removal-intent paths without a reviewed merge boundary.
- [x] Read `.github/skills/cpp-build/SKILL.md`, `.github/skills/cpp-modern-style/SKILL.md`, `.github/skills/perf-validation/SKILL.md`, `Specs/Core/Core_SharedHelpers.md`, `Specs/Testing/Testing_SelfTests.md`, and `Specs/Testing/Testing_PerformanceValidation.md` before code edits.
- [x] Capture the current test-enabled x64 Debug behavior for the existing empty-folder, Space navigation-shell, and removal-focus cases. Record failures as baseline facts; do not weaken the target contract to match them.
- [x] Add or extend test-only FolderView observations for current index/name, focused-flag count, selection digest, anchor, current-resolution reason, focus-ownership epoch, empty-parent action, potential-drag source digest, and focus-memory count/payload/evictions. Do not expose mutable production internals.
- [x] Add `folder.focus.resolve_us` instrumentation and the deterministic `folderView_perf_focus_selection_state` scenario before changing resolver behavior so the old implementation can provide a comparable same-machine baseline. Emit aggregate rows only.
- [x] Build test-enabled x64 Release, run the new perf scenario through the governed runner, validate its archive, and record baseline archive/machine/build flavor in Closeout evidence.
- [x] Confirm the observations and scenario themselves do not add per-row metrics, full-list UI-thread name copies, or file-backed trace writes inside measured hot paths.

Baseline commands:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}

.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_empty_folder_state --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs --selftest-timeout-multiplier=4
```

When PowerShell automation launches the GUI-subsystem executable, use `Start-Process -Wait -PassThru` and inspect `ExitCode`; the direct commands above are shorthand for an interactive developer shell.

Phase 1 baseline evidence (2026-08-20, before production behavior changes): test-enabled x64 Debug build receipt
`e7383f87ccbf4490ffa556bcc7bc21086f67edcad389870269cfd71ff21831a2` completed with zero warnings and zero errors.
Separate `Start-Process` runs of `folderView_empty_folder_state`,
`cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable`, and
`cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs` each exited `0`.

The standard Release output remained locked by an independently launched app, so the exact tracked working-tree snapshot was built and exercised in an isolated temporary worktree without touching that process. The test-enabled x64 Release build exposed and then passed after a behavior-neutral `ENABLE_TESTS` guard correction in `DxUi.Menu.cpp`; governed build receipt
`d4857ae72bd280dfd07b6aa5e7ecd403d0e425d3fe0ab077a9f6b25f20eed317` completed with zero warnings and zero errors. The governed Commands run passed 1/1 and its compacted archive at
`Specs/TestRuns/4cb089111a23/Commands/2026-08-20_154800_folder_view_focus_selection_baseline_release/` passed explicit pre-stage archive validation. The baseline contains 10,000 displayed items, 206 accepted model changes and `folder.focus.resolve_us` rows, the expected legacy two current-invariant violations, and the expected legacy memory-cap violation (601 entries, 103,400 payload bytes, zero evictions). Observations use test-only read-only snapshots; the production metric is one aggregate scope row per resolver invocation, does not clone the displayed-name list, and uses only the shared perf sink after the timed resolver work—there is no bespoke per-row metric or trace write.

### Phase 2 — Centralize current resolution, generic rehome, and ownership epochs

Files:

- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.Selection.cpp`
- `RedSalamander/FolderView.Enumeration.cpp`
- `RedSalamander/FolderView.Rendering.cpp`

Checklist:

- [x] Introduce one FolderView-private current-resolution input/result model covering explicit host target, surviving identity, generic repair probe, navigation restore, host-owned removal intent, fail-safe first row, and empty state. Include a resolution-reason enum for diagnostics/perf.
- [x] Make accepted enumeration, refresh, sort, filter/hide/visibility rebuild, and fallback repair finish through the same invariant-normalization operation.
- [x] Clear all item `focused` flags before setting the one resolved current item; publish `_focusedIndex` and the item flag as one UI-thread transition.
- [x] Preserve a surviving current item by provider-aware identity across sort and same-location refresh before consulting memory or positional repair.
- [x] Implement generic disappearance as the exact successor/predecessor rule in the target model: production-comparator insertion for keyed sorts, prior-order surviving-neighbor scan for Sort None, then first-row fail-safe. Do not use the raw old index as the decision.
- [x] Integrate successor lookup with the existing rebuild/sort identity work so several disappeared predecessors/current/successors are handled without cloning every display name.
- [x] Keep validated explicit rename/create targets and I12's exact-success host-owned removal result as higher-priority one-shot intents passed into the resolver; do not duplicate their proof logic.
- [x] Refactor focus-ownership epoch updates so only an actual user-owned current-identity change increments the epoch. Selection-only changes, including empty-background clear with retained current, do not increment it.
- [x] Keep folder/provider/sort epochs and the pending-removal exact-`S_OK`/newer-generation checks unchanged. Prove an actual current change still invalidates pending rehome.
- [x] Do not choose the first selected item as an implicit current fallback and do not mutate selection during repair.
- [x] Notify current-item consumers, remember the final identity, update preview/status, and ensure visibility only after final normalization.
- [x] Preserve the zero-real-item invariant: `_focusedIndex` invalid, zero focused flags, zero selection, and no anchor.
- [x] Add/enable `folder_view_focus_invariant_generic_rebuild` with single- and multi-row disappearance plus keyed-sort/Sort-None successor/predecessor cases before considering this phase complete.

Constraints:

- Do not copy all visible names merely to repair one current item.
- Integrate identity lookup with existing rebuild/sort passes where possible.
- Do not add per-item performance metric rows.

Verification:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_invariant_generic_rebuild --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs --selftest-timeout-multiplier=4
```

Expected result: build exits 0; the generic case proves multi-row successor/predecessor behavior and the invariant; the existing removal-focus case retains exact-success/epoch safety.

### Phase 3 — Replace unbounded focus memory with the bounded provider-aware LRU

Files:

- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Enumeration.cpp`

Checklist:

- [x] Re-run the shared-helper search in `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/`. Keep this cache FolderView-private unless an exact provider-keyed, dual-cap, LRU semantic match already exists.
- [x] Replace `_focusMemoryRootKey` plus the unbounded map with FolderView-owned bounded state using O(1)-average key lookup/touch and deterministic least-recently-used eviction.
- [x] Key entries by plugin id, provider instance/mount context, and provider-normalized location identity; normalize the remembered component through the same provider contract.
- [x] Never key by COM pointer, object address, or HWND. Use exact live-instance scoping only when a provider cannot promise stable recreation identity.
- [x] Account owned UTF-16 provider/context/folder/item payload and enforce both 512 entries and 512 KiB after every insert/update.
- [x] Reject an individually oversized or unsupported identity as a no-cache event without evicting the useful cache or failing navigation.
- [x] Touch recency on successful restore and record/update; preserve stable insertion-order tie breaking.
- [x] Stop globally clearing memory in `SetFileSystem(...)` or root transitions; the provider/context key prevents cross-provider collision.
- [x] Keep `std::bad_alloc` fatal per repository policy. Do not add `catch (...)` or a catch-and-continue partial-cache path.
- [x] Expose read-only debug observations for entry count, accounted payload, evictions, hit/miss/no-cache reason, and deterministic LRU tests.
- [x] Add/enable `folder_view_focus_memory_bounded_lru` and `folder_view_focus_memory_navigation_roundtrip` before considering this phase complete.

Verification:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_memory_bounded_lru --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_memory_navigation_roundtrip --selftest-timeout-multiplier=4
```

Expected result: build exits 0; both cases pass; every sampled mutation remains within both caps; provider/context round trips restore current without restoring selection.

### Phase 4 — Implement selection, pointer/drag, Space, commands, and empty-state action

Files:

- `RedSalamander/FolderView.Selection.cpp`
- `RedSalamander/FolderView.Interaction.cpp`
- `RedSalamander/FolderView.DragDrop.cpp`
- `RedSalamander/FolderView.Menus.cpp`
- `RedSalamander/FolderView.Enumeration.cpp`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderWindow.SelectionSize.cpp` only if the existing post-toggle resnapshot/cancellation path needs correction

Checklist:

- [x] Split current-only movement, replace-selection, toggle-selection, replace-range, additive-range, and toggle-then-advance into operations with explicit current/selection/anchor effects.
- [x] Make every real-item left-button-down establish that item as current during down handling. For no-anchor Shift/Ctrl+Shift, capture the valid pre-press current first, then move current to the clicked endpoint and include both endpoints.
- [x] Implement the complete input table: plain selected/unselected, Ctrl, Shift, Ctrl+Shift, arrows/Home/End/Page, Alt+Up/Down, Esc, Ctrl+A, masks/invert, item right-click target establishment, every modified left background miss, pointer right background miss, and keyboard context-menu routing.
- [x] Preserve **Space = toggle operated current + request post-toggle selection-size recomputation + advance/clamp**, and **Insert = toggle operated current + advance/clamp without explicit size work**. Keep Quick Search Space routing unchanged.
- [x] Ensure size work consumes a post-toggle immutable selected-identity snapshot, cancels discarded/deselected work, ignores stale completions, and never treats the advanced-to current item as implicitly selected.
- [x] On empty-background left-button-down, clear selection, anchor-to-current, pointer target, and potential drag while retaining current and leaving focus-ownership epoch unchanged.
- [x] Replace `_drag.dragging`'s ambiguous armed/active meaning with explicit potential-gesture state as needed. A fresh item press snapshots source identities after modifier selection handling.
- [x] Resolve selection-aware commands and drag source through one selected-or-current policy: selected displayed identities when nonempty, otherwise current item. Status summaries remain selection-only.
- [x] Arm drag only when the pressed item belongs to the resolved source. Ctrl+Click that removes the pressed item while another selection remains must not drag that leftover set.
- [x] Use only the effective system drag-threshold rectangle around the press point. Remove the entire-row `startItemRect` gate; a wide item must not require crossing its visual bounds.
- [x] Disarm potential drag on release without drag, cancellation/Escape ownership, capture loss, teardown, navigation/model replacement, and every new press. Double-click must not inherit an older source.
- [x] Keep plain/right-click item target establishment deterministic before drag or menu snapshot. Background context routing must not infer an item pointer target from retained current.
- [x] Preserve selection only for surviving provider-aware visible identities across same-location rebuild/sort, plus the explicit proven rename-chain transfer. Removed/filtered/hidden/ordinary-reappearing identities remain unselected.
- [x] Apply Hide Selected/Unselected, visibility flags, filter changes, Select All, masks/invert, failed/stale results, and Save/Restore Selection exactly as the display-membership table specifies.
- [x] Clear selection on actual folder/provider/context/history navigation and never store it in focus memory.
- [x] Keep current-item and selection notifications independent; do not increment focus ownership or refresh preview/current-detail consumers for selection-only changes.
- [x] Implement the separate standard-empty **Go to parent** action without creating a pseudo `FolderItem`: Enter and double-click call `NavigateUp()`, item commands/drag remain unavailable, and filter-empty/host-message/busy/error states do not gain this action.
- [x] Preserve existing root/connection-root `NavigateUp()` callback behavior and existing empty-state rendering geometry/resources.
- [x] Add/enable the directly dependent input, background, drag, Space, empty-state, and selection-persistence cases named in Phase 5 before considering this phase complete; Phase 5 then runs the complete matrix and closes cross-case gaps.

Verification:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_selection_input_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_empty_background_keeps_current --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_drag_source_target_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_empty_folder_state --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable --selftest-timeout-multiplier=4
```

Expected result: build exits 0 and each exact case passes with the transition, size-workflow, drag-source, empty-parent, and epoch assertions defined below.

### Phase 5 — Complete deterministic behavioral and invariant coverage

Primary files:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` only for the existing Quick Search Space route
- the Commands registration owner if a new family split is necessary

Required deterministic cases:

- [x] `folder_view_focus_invariant_generic_rebuild`
   - Build sorted `a,b,c,d`; focus `b`; externally remove it without host-owned operation intent; require `c` current and unchanged selection.
   - Remove the last current row; require the prior row.
   - Build sorted `a,b,c,d,e`; focus `c`; remove `a` and `c` in one accepted rebuild; require `d`, not raw-old-index `e`. Then remove the available successors and require the nearest surviving predecessor.
   - Change active sort/filter while removing current and prove the production comparator/probe chooses the successor in the resulting visible order.
   - Exercise filter/hide removal and require the replacement in resulting UI order.
   - Exercise a zero-row result and repopulation.
   - At every checkpoint assert `itemCount == 0 ? focusedFlagCount == 0 : focusedFlagCount == 1` and index/name agreement.
- [x] `folder_view_empty_background_keeps_current`
   - Use the real pointer route for unmodified/Ctrl/Shift/Ctrl+Shift left misses and pointer right-click miss; use the real keyboard context-menu route separately.
   - Require selection cleared, current retained, anchor reset, drag disarmed, no stale second focus flag, and unchanged focus-ownership epoch.
   - Begin a valid pending host-owned removal-focus intent, click empty background without changing current, complete exact `S_OK`, and require the intent to remain eligible. Repeat with an actual user current change and require invalidation.
   - Move the pointer after the background press and prove that neither the retained current item nor the former selection becomes a drag source; require a fresh item press before drag initiation.
   - Open the pointer background context route and prove it clears selection and does not infer an item target from retained current. Prove the keyboard context-menu request instead preserves selection and uses selected-or-current without a synthetic hit-test.
- [x] `folder_view_focus_memory_bounded_lru`
   - Exercise repeated insert, update, lookup/touch, 513th-entry eviction, payload-cap eviction, oversized-entry rejection, and count/payload bounds.
   - Prove the touched oldest entry survives while the true least-recently-used entry is evicted.
- [x] `folder_view_focus_memory_navigation_roundtrip`
   - Navigate A → B → A and require exact current restoration.
   - Switch between two provider/context identities with the same textual path and require independent restoration.
   - Return through Back/Forward and assert current identity, not only path/item counts.
   - Require selection empty after every actual location transition.
- [x] `folder_view_selection_input_contract`
   - Cover every row in the input transition table, including selected versus unselected plain/right click, replace versus additive ranges, Space versus Insert, Esc, Ctrl+A, inversion, and extension masks.
   - Assert that item left-button-down changes current before button-up and before drag/command observation.
   - Clear the anchor, Shift+Click another item, and require the pre-press current item to become the anchor and both endpoints to be selected. Repeat for Ctrl+Shift additive selection.
   - For Space starting on an unselected current item, require that item to become selected, enter the post-toggle file/folder size snapshot, remain selected after current advances, and contribute known file bytes or asynchronous folder-subtree bytes. Starting on a selected current item, require deselection, size-work cancellation/stale-result rejection, and exclusion from the new total. In both cases require advance to next with no wrap at the last row, anchor at resulting current, and no implicit selection of that resulting current. Require Insert to advance without requesting subtree-size work.
- [x] Extend `cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable`
   - Preserve its existing navigation-shell quiet-point assertions.
   - Assert the full selected-item/current-item progression, the operated folder identity passed to post-toggle size resnapshot, stale-generation rejection, and final status bytes/calculating state.
   - Prove Quick Search Space remains text and emits no selection-size request.
- [x] Extend `folderView_empty_folder_state`
   - Assert zero real items, invalid current index, zero focused flags, zero selection, invalid anchor, and active separate parent action.
   - Assert Enter, double-click, and Backspace route upward; the action never appears in item command fallback or drag source.
   - Assert filtered-empty, host-message, busy/cancel/error, and root callback behavior remain distinct and correct.
- [x] `folder_view_selection_refresh_filter_contract`
   - Preserve surviving selection across sort/refresh.
   - Transfer only a selected item through a proven rename chain.
   - Drop removed/filtered/hidden selection and prove it does not resurrect when visibility returns.
   - Invoke Select All, apply a filter that hides a subset, then clear the filter; require only the continuously displayed subset to remain selected and require the reappearing items to be unselected.
   - Invoke Select All, add a new item through an accepted refresh, and require the new item unselected while prior survivors remain selected.
   - Remove a selected item, accept the refresh, recreate the same ordinary identity without a rename hint, and require it unselected.
   - Inject a failed, cancelled, stale, or rejected result and require the current displayed selection unchanged.
   - Prove automatic current repair does not select its replacement.
- [x] `folder_view_drag_source_target_contract`
   - With empty selection, press the current item, move inside the system drag-threshold rectangle, and require no drag; then leave the rectangle and require a drag snapshot containing only current.
   - With nonempty selection, press a selected item and require the immutable drag snapshot to contain the selected set. Press an unselected item without modifiers and require sole-selection replacement before the snapshot.
   - Ctrl+Click a selected item off while other selection remains and require no drag to arm from the now-unselected pressed item.
   - Prove release, cancellation, capture loss, navigation/model replacement, and a new press disarm the prior potential drag.
   - Prove empty-background movement and double-click cannot inherit an earlier drag source.
   - Prove selection-aware commands and drag-source resolution agree on selection-or-current fallback while status summaries remain based on actual selection only.
- [x] `folder_view_selection_visibility_restore_contract`
   - Hide Selected Names then Show Hidden Names; require returning identities unselected.
   - Hide Unselected Names then Show Hidden Names; require only continuously displayed selected identities to remain selected.
   - Disable/re-enable Hidden Files and System Files around selected fixtures; require returning identities unselected.
   - Apply Select/Unselect masks, same-name/extension commands, and Invert while filtered; require later-visible identities unaffected.
   - Save a selection, narrow visibility, invoke Restore, then widen visibility; require only matches visible during Restore selected. Invoke Restore again and require now-visible saved matches selected as an explicit user action.
   - Assert short-lived missing-selection rename provenance never contributes to selected count, status summaries, or command targets.
- [x] Extend existing host-owned Delete/Move tests to assert the global one-current-item invariant and selection-only epoch behavior without weakening exact-`S_OK`, provider/folder/sort epochs, newer-enumeration, or immutable-target conditions.

Test requirements:

- Use existing Commands sandbox and bounded message/snapshot polling helpers.
- Wait for path, item model, and focus state to remain stable across repeated samples before assertions.
- Pointer cases require a foreground-capable desktop and must follow the directed-input warning/restore contract when desktop-global input is unavoidable.
- Prefer compiled behavioral/debug-seam assertions over source-shape tests.

Focused Debug commands after a test-enabled build:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_invariant_generic_rebuild --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_empty_background_keeps_current --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_memory_bounded_lru --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_focus_memory_navigation_roundtrip --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_selection_input_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_drag_source_target_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_selection_refresh_filter_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folder_view_selection_visibility_restore_contract --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_empty_folder_state --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs --selftest-timeout-multiplier=4
```

When launched from PowerShell automation, use `Start-Process -Wait -PassThru` rather than `$LASTEXITCODE` for the GUI-subsystem executable.

### Phase 6 — Validate performance and bounded memory with archived evidence

Protected user-visible scenarios:

- large same-folder refresh/filter/sort when the current item survives or disappears;
- navigation restore across many visited locations;
- selection preservation across large accepted enumerations.

Metric contract:

- Reuse the one-row-per-refresh `folder.refresh.*` family and `folder.refresh.request_to_paint_us`.
- Add `folder.focus.resolve_us`, one row per accepted model rebuild, with `value0 = displayed item count` and `value1 = resolution-reason enum`.
- Emit scenario-checkpoint aggregates for `folder.focus_memory.entry_count`, `folder.focus_memory.payload_bytes`, and `folder.focus_memory.eviction_count`; do not emit a row for every normal focus move.
- Extend the selftest artifact with resolution-reason counts, maximum observed cache entries/payload, eviction count, and the invariant failure count.

Deterministic performance case:

- Add `folderView_perf_focus_selection_state` using the dummy provider.
- Use at least 10,000 displayed items and at least 200 model-change samples when making p95 claims.
- Include current-survives, current-disappears, sparse selection, dense selection, filter hide/show, and sort-toggle phases.
- Assert the focus-memory caps independently; do not perform hundreds of real provider navigations merely to test the LRU data structure.
- Expected direction: bounded memory; no new per-row instrumentation; no material regression in same-folder refresh request-to-paint or selection-preservation cost.

Checklist:

- [x] Confirm `folder.focus.resolve_us` emits exactly once per accepted model rebuild with displayed count and resolution reason, and that existing `folder.refresh.*`/request-to-paint rows remain one-per-refresh aggregates.
- [x] Confirm focus-memory observations emit at scenario checkpoints only and report maximum entries, maximum payload, evictions, hits/misses/no-cache reasons, and zero cap violations.
- [x] Complete `folderView_perf_focus_selection_state` with at least 10,000 displayed items and 200 accepted model-change samples for every p95 claim.
- [x] Exercise current survives/disappears, multiple-neighbor disappearance, keyed sort and Sort None repair, sparse/dense selection, filter hide/show, selection preservation, and bounded-memory churn.
- [x] Include pointer current/selection transitions in the scenario only through aggregate input-to-paint observations; do not synthesize per-mouse-move JSONL rows.
- [x] Build test-enabled x64 Release and run at least one valid same-machine baseline and candidate archive with the same suite, scenario parameters, and build flavor.
- [x] Validate each curated archive with `Test-TestRunArchive.ps1` and compare using `Show-PerfRuns.ps1 -FolderViewPreset -FailOnQuality -ShowBuildFlavor`.
- [x] Record baseline/candidate paths, machine hash, build flavor, sample counts, resolution-reason counts, cache maximums, changed metrics, analyzer verdict, and caveats in Closeout evidence.
- [x] Add a machine-keyed hard budget to `FolderViewPerfBudgets.json5` only if repeated stable same-machine Release evidence supports it. Otherwise document why the scenario remains visible but unbudgeted; do not invent a threshold.
- [x] Reject the candidate if it copies the full displayed-name set on the UI thread, emits per-item metrics, violates either memory cap, or materially regresses refresh request-to-paint/input-to-paint without an explicitly reviewed tradeoff.

Build test-enabled Release:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

Run through the governed runner and archive each run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_focus_selection_state -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -TimeoutMultiplier 8
.\Tools\Show-PerfRuns.ps1 -FolderViewPreset -CompareRun <baseline-run>,<candidate-run> -FailOnQuality -ShowBuildFlavor
```

Closeout evidence must state baseline/candidate paths, same-machine and same-suite status, build flavor, sample counts, changed metrics, cache maximums, and caveats. Add a machine-keyed hard budget only when a stable same-machine Release baseline supports it; unknown machines remain visible and non-silent per the existing budget contract.

Expected archive location:

```text
Specs/TestRuns/<MachineHash>/Commands/<RunId>/
```

Validate every curated archive before closeout:

```powershell
.\Tools\Test-TestRunArchive.ps1 -RunPath Specs\TestRuns\<MachineHash>\Commands\<RunId>
```

### Phase 7 — Run final gates and close the operation

- [x] Re-run every focused Debug case listed in Phase 5 in a test-enabled build.
- [x] Run the test-enabled Release performance case and analyzer comparison from Phase 6.
- [x] Run the canonical final gate:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
   ```

- [x] Validate documentation/index integrity:

   ```powershell
   .\Tools\Get-SpecInventory.ps1 -FailOnFindings
   git diff --check
   ```

- [x] Re-run the Phase 0 conflict scan and verify no normative behavior lives only in this plan.
- [x] Review the final diff for scope, unrelated user changes, forbidden exception/ownership patterns, and accidental weakening of I12 removal proof.
- [x] Fill every Closeout evidence field with actual paths/results.
- [x] Move this plan to `Specs/Plans/Done/` and update `Specs/Plans/WIP/README.md` in the same change only after every Done criterion is checked.

## Verification matrix

| Surface | Required proof |
|---|---|
| Non-empty invariant | Exactly one valid current index and one matching item flag after every tested transition |
| Empty display | No real current/selection/anchor/drag; standard Go-to-parent remains a separate Invoke action with Enter/double-click/Backspace coverage |
| Generic disappearance | Comparator/probe or Sort-None survivor scan chooses the first successor, otherwise predecessor, including multiple simultaneous removals; replacement not auto-selected |
| Sort | Current and selection preserved by provider-aware identity |
| Filter/hide | Hidden current rehomed; hidden selection removed; neither resurrects on widening; Select All → filter → unfilter is no longer Select All |
| Display membership | New/reappearing identities unselected; removed identities pruned; membership-preserving transitions retain selection exactly |
| Save/Restore Selection | Restore affects only matching identities visible when explicitly invoked; widening visibility has no deferred effect |
| Failed/stale model result | Existing displayed selection remains unchanged because no replacement model was accepted |
| Navigation memory | Folder/provider/context round-trip restore with no selection restore |
| Memory bound | At most 512 entries and 512 KiB tracked payload after every mutation |
| Pointer | Button-down current timing and Plain/Ctrl/Shift/Ctrl+Shift/right/background transitions match the state model |
| Drag source | Selection-or-current fallback, pressed-item membership, immutable snapshot, system threshold, and every disarm path are proven |
| Keyboard | Navigation, Space select+size+advance, Insert toggle+advance without size, Quick Search Space, Esc, range, select-all, invert, and masks match the table |
| Command safety | Captured targets remain immutable across focus repair and refresh |
| Host-owned removal | Existing exact-success/provider/folder/sort/current-ownership semantics preserved; selection-only empty background retains eligibility while a real current change invalidates |
| Accessibility | Current and selection remain distinct; empty-parent action is Invoke-only if exposed and never SelectionItem/current |
| Performance | Aggregate instrumentation only, sufficient samples, bounded memory, archived Release evidence |

## Done criteria

- [x] Phase 0 updated every owning UI/FileSystem/testing/performance normative spec before production behavior changes, with no contradictory legacy text.
- [x] Every non-empty accepted FolderView model has exactly one current item logically and visually.
- [x] Generic disappearance resolves comparator/probe or Sort-None first-surviving-successor then predecessor, including simultaneous removal of rows before/at/after current.
- [x] A zero-real-item model has no current/selection/anchor/drag, while standard Go-to-parent remains a separate Enter/double-click/Backspace Invoke action.
- [x] Navigation restores per-location current item across provider/context switches.
- [x] Focus memory enforces both 512-entry and 512-KiB payload limits with deterministic LRU eviction.
- [x] Selection follows the complete transition and persistence model.
- [x] Space preserves select+post-toggle-size+advance with no wrap; Insert advances without explicit subtree-size work; Quick Search Space remains text.
- [x] Empty-background selection clearing retains current and does not change focus-ownership epoch or invalidate a matching pending removal rehome.
- [x] Pointer/drag behavior uses current-on-button-down, pre-press no-anchor ranges, selected-or-current immutable sources, the system threshold only, and complete disarm paths.
- [x] Every display-membership transition in the normative table has deterministic coverage, including refresh additions, ordinary reappearance, Hide/Show commands, Hidden/System visibility, failed/stale results, and explicit Save/Restore behavior.
- [x] Current repair never mutates selection or pending command targets.
- [x] Required deterministic Commands cases pass and are present in native case inventory.
- [x] Test-enabled x64 Release perf evidence is archived and analyzed.
- [x] `Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` passes.
- [x] `Get-SpecInventory.ps1 -FailOnFindings` and `git diff --check` pass.
- [x] Evidence paths and measured results are recorded below.
- [x] This plan is moved to `Specs/Plans/Done/` and removed from the active WIP index.

## STOP conditions

Stop and reconcile before implementation continues if:

- Phase 0 cannot eliminate the `leave focus unset`, range-without-current-move, Space toggle-in-place, or Insert-starts-size-work contradictions without changing another product owner's accepted contract;
- provider path/item identity cannot be obtained without guessing semantics or colliding provider contexts;
- the 512-entry or 512-KiB policy conflicts with an existing accepted settings/cache contract;
- a change would weaken immutable destructive-command target validation;
- I12's host-owned removal-focus implementation has drifted so the specialized intent cannot feed the central resolver without changing its proof rules;
- the existing removal-focus epoch is consumed as selection ownership by another proven contract, so selection-only empty-background behavior cannot be separated safely;
- the production sort policy cannot place a removed-current repair probe using the same comparator/tie-breaker without reimplementing or diverging from canonical sort semantics;
- the resolver requires copying every displayed name or introduces per-item metric emission on the UI thread;
- selection preservation would retain hidden/filtered identities as actionable targets;
- focused behavioral tests cannot distinguish retained current item from Win32 focus or UIA selection;
- the standard empty-folder parent action cannot remain outside `_items`/`_focusedIndex` while preserving its existing Enter/double-click/root route;
- Space size recomputation cannot consume the post-toggle immutable selected set or reject stale completions without redesigning the separately owned selection-size worker;
- baseline/candidate performance evidence is not same-machine/same-suite but is being used for a definitive regression or improvement claim;
- required interactive desktop coverage is unavailable and no valid lower-level behavioral seam proves the same routing contract;
- an unrelated dirty-worktree change overlaps a touched file and cannot be preserved safely.

## Closeout evidence

Fill during implementation; do not mark this plan complete with placeholders.

- Baseline archive: `Specs/TestRuns/4cb089111a23/Commands/2026-08-20_154800_folder_view_focus_selection_baseline_release/`; governed run `20260820T134407Z-101228-2c620b6b43944aac82c22e3ed0da606a`, 1 passed / 0 failed / 0 skipped.
- Candidate archive: `Specs/TestRuns/4cb089111a23/Commands/2026-08-20_183441_folder_view_focus_selection_final_candidate_release/`; governed run `20260820T163056Z-82244-416313433334403ab57fd61921519116`, build receipt `d80e658f1df02cb5b42be5c82fa349f697118c9261ec579e31018bc56258ce03`, source snapshot `2f496c461d1d61509532f5f760015352c7d28a953af41e47d0ef8e30a7b70ae7`, 1 passed / 0 failed / 0 skipped. Both final comparison archives passed explicit `Test-TestRunArchive.ps1` validation.
- Machine hash/build flavor: `4cb089111a23`; both test-enabled x64 Release, local console, 150% DPI, 120 Hz, NVIDIA GeForce RTX 5080. WARP/RDP variants were not run.
- Analyzer result: `folder.focus.resolve_us` p95 `4700 us` baseline (206 rows) versus `4406 us` final candidate (209 archived rows), `-6.3%`, noise/no regression, with the explicit 200-sample quality gate passing. FolderView preset comparison reported 0 regressions and 4 noise metrics plus one candidate-only aggregate input-to-paint observation; frame total changed `340321 us` to `332523 us` (`-2.3%`, 16 versus 19 sparse rows) and request-to-paint changed `878498 us` to `871710 us` (`-0.8%`, 4 versus 5 sparse rows). The preset quality command used an explicit one-sample override because this correctness-oriented scenario is not the 200-frame stress scenario. No hard budget was added: one same-machine pair and sparse frame/refresh rows are insufficient for a stable threshold.
- Candidate state metrics: 10,000 displayed items, 208 explicitly observed model changes, 207 resolver rows inside the measured scenario interval, resolution reasons `{surviving-current: 203, generic-successor: 3, generic-predecessor: 1, navigation-memory: 1}`, zero invariant/cap violations, maximum/final cache `512 entries / 132096 bytes`, 89 evictions, 0 no-cache events. Two additional resolver rows occur during runner setup/teardown and explain the 209 analyzer rows.
- Phase 0 conflict scan/spec inventory: conflict scan clean; `Get-SpecInventory.ps1 -FailOnFindings` reported 0 findings; `git diff --check` passed apart from informational LF-to-CRLF warnings.
- Generic multi-removal resolver cases: `folder_view_focus_invariant_generic_rebuild` passed, including keyed-sort and Sort-None successor/predecessor repair.
- Space/Insert/Quick Search cases: `folder_view_selection_input_contract`, `cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable`, and `cmd_pane_quickSearch_integrated_navigation` passed; the exact Space case also passed 10/10 repeated during stabilization.
- Empty-parent action cases: extended `folderView_empty_folder_state` passed.
- Empty-background/removal-epoch cases: `folder_view_empty_background_keeps_current` and `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs` passed.
- Pointer/drag cases: `folder_view_selection_input_contract` and `folder_view_drag_source_target_contract` passed.
- Focused correctness results: final clean test-enabled Debug build (`27f5c3401a06f62e99c60413f8daf9c58661df7237f6c6642b0299f331caf97f`, zero warnings/errors) followed by 14/14 focused Commands cases and the Debug performance scenario passing.
- Fresh Full run id/result: governed Fresh run `20260820T203027Z-32116-0f41f1276c4640a79310e3ecd6ab11df`, Debug build receipt `b50f95cd8a23ad69dca94cd6728352f670fc46f736143c00b10c5ff9c3a57fcc0`, repository verdict **PASSED**, 1912 total / 1859 passed / 0 failed / 53 documented skips, all 18 governed entries promoted on a stable source snapshot, wall time 56m 46.7s. Commands passed 865/865 with 2 documented skips; File Operations passed 118/118 with 20 documented skips; Compare Directories passed 227/227 with 31 documented skips; tooling Pester passed 634/634 plus build-toolchain Pester 2/2.
- Pre-acceptance transient evidence: unchanged-snapshot Fresh run `20260820T193115Z-100756-91eeeb1b58f1448fb1081343d53074d9` failed only three unrelated, load-sensitive Commands cases (`theme_cycle_overlay_timer_fallback`, `cmd_pane_batchRename_window_invokes_success_callback`, and `cmd_pane_batchRename_window_executes_three_member_cycle`). Each case passed immediately in an isolated rerun with the same binaries; the accepted Fresh run then passed the complete Commands family. This run is diagnostic only and is not acceptance evidence.
- Final authoritative specification state: `UI_FolderView.md`, `UI_CommandMenuKeyboard.md`, `UI_FolderWindow.md`, `FileSystem_FileOperations.md`, `Testing_SelfTests.md`, and `Testing_PerformanceValidation.md` are updated in the working tree at HEAD `8313c1950e69f7ef1e5e057c5b0850381f410d33`; no commit was created because none was requested. Final `Get-SpecInventory.ps1 -FailOnFindings` reported 0 findings, `Get-ToolInventory.ps1 -FailOnFindings` classified 172/172 files with 0 findings, the Phase 0 retired-clause scan was clean, whole TestRuns inventory validated 541 files, and `git diff --check` passed apart from informational LF-to-CRLF notices.
- Interrupted-run diagnostic: Fresh run `20260820T174041Z-57644-06d74b151c8e4af48012a81dc0011c43` was explicitly preserved as non-acceptance evidence after an unrelated intermittent `PluginContractTests.exe` stall. A subsequent focused visible execution passed in about 18 seconds, and the accepted Fresh run passed the same plugin contract in 19.5 seconds. The diagnostic archive remains under `Specs/TestRuns/4cb089111a23/Continuation/20260820_201700_folderview_closeout_plugincontract_hang/` and was not used as acceptance evidence.

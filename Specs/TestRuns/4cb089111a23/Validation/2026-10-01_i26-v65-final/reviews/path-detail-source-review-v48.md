# V48 inline path detail — source review, native qualification pending

Root is the sole repository writer. Sol reviewed the pinned API/lifecycle/test routes; Luna drafted
the gallery extension. This record is not a build, native pass or performance acceptance.

## Implementation boundary

- Consumer-owned `FileOperationsPathDetailPanel` uses the existing popup host and pinned DxUi 49
  multiline, read-only `TextField`. It creates no HWND, thread, timer or shared-library capability.
- Path row focus retains its existing tooltip and exact Ctrl+C. Enter/Space/click/UIA Invoke opens
  detail. Initial full selection has caret zero so the visible document starts at the root.
- While detail is active, root children outside it are invisible; task callbacks/default shortcuts
  cannot resolve a decision. DxUi's existing visible-tree Tab traversal owns forward/reverse keys.
- Escape returns to the invoking row without native activation. Editor destruction is deferred beyond
  the close callback. Hide retires it immediately; HWND teardown clears borrowed pointers before detach.
- Same-row values refresh. Row disappearance closes detail. Unchanged values retain selection and
  hosted focus across retained-tree rebuild; bounds/theme/DPI continue through the existing host.
- Closed synchronization does not search the path descriptors or allocate detail children. One small
  empty panel may remain after first use; the four children, copied path and editor layouts are released.

## Review corrections applied by root

1. Pinned `Control::SetVisible` already refreshes UIA snapshots synchronously. Sol withdrew the
   proposed missing-explicit-refresh finding after root checked `DxUi.cpp:389–402`.
2. Wheel deltas are raw Windows 120-unit values, not normalized ticks.
3. Removed a redundant manual Tab loop using native `GetKeyState`; the host already tracks synthetic
   and native modifier state and excludes invisible subtrees.
4. Gallery running destination rows expose `destinationFolder`, not `currentDestinationPath`; the
   >4Ki fixture now opens the exact source-row value.
5. `UIA_IsReadOnlyAttributeId` is a text attribute, not a property. Pinned 49 returns an empty variant
   for text attributes, so the draft's assertion was unsupported. Tests use actual read-only state,
   rejection of edit input and exact full TextPattern text. External AT support is not newly claimed.
6. Kept the gallery's six-column TSV schema; removed the draft's extra language column.
7. Gallery coverage runs on every actual monitor and records actual DPI. It does not hard-code a
   failure for machines lacking 96/144 DPI. This workstation's required two-monitor receipts remain
   an acceptance gate, with no synthetic-DPI substitution.
8. Unicode substring selection crosses Hebrew and an emoji surrogate pair, not only an ASCII suffix.
   Every close checks editor absence, task/scope/default stability, host and attachment counts.

## Validation defined

Existing Commands `cmd_pane_fileops_popup_throughput_disclosure` now exercises >4Ki French UNC text,
raw UIA TextPattern, full and selected copy using the existing clipboard ownership guard, PageDown
viewport movement, Tab/reverse Tab, edit rejection, stale provider retirement, live root replacement,
hide/removal and 40 open/close cycles. Memory samples and open/update/release metrics are retained.
The actual gallery adds root and scrolled detail captures for dark/light at 480 DIP on each monitor.

ResourceLocalizationContracts: first attempt 8/1/0 identified missing cs-CZ/ja-JP/sk-SK entries.
After adding those translations, `path-detail-resource-contracts-r2.xml` records **9/0/0**.
clang-format 22.1.3 and `git diff --check` passed for the changed C++ files.

V47 is frozen and still running Full/Fresh in a separate checkout. Its result is pre-feature evidence.
V48 build, native journeys, real monitor captures, final Full and paired resources remain pending.

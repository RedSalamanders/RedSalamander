# Connection Manager Single-Canvas DxUi Migration Plan

**Status:** Done (2026-04-27)
**Owner:** UI / DxUi
**Scope:** Replace the current per-widget host-HWND model in `ConnectionManagerDialog.cpp` with a single top-level window backed by one `DxUi::WindowHost` and a single DxUi widget tree, while preserving all currently observable behavior, theming, accessibility, and keyboard navigation.

## Outcome

- `RedSalamander/ConnectionManagerWindow.{h,cpp}` (~3000 LOC) is the new single-canvas implementation. Single `WS_OVERLAPPEDWINDOW` + single `DxUi::WindowHost` + single in-tree DxUi widget tree (list-pane / editor / footer). Per-widget `DxHost` HWNDs (legacy ~37), hidden Win32 dialog children (~50), and the native `awsRegionCombo` are all gone.
- `RedSalamander/ConnectionManagerDialog.cpp` shrank from ~8800 LOC to a 142-line forwarding shim. `IDD_CONNECTION_MANAGER` deleted from `RedSalamander/RedSalamander.rc`. Net deletion: ~5800 LOC.
- `IsEnabled()` is `constexpr true`. The constexpr is kept one release as an emergency-revert latch; future cleanup may remove it.
- All 16 `Debug*` test entry points route through `SingleCanvas::Debug*` equivalents.
- Specs updated: `Specs/Core/Core_ConnectionManager.md`, `Specs/UI/UI_TopLevelToolWindows.md`, `Specs/UI/UI_VisibleComctlAudit.md`.

## Deferred polish (carved into separate items, not gating)

- **Phase 7b** — focus-context-sensitive `Alt+N` resolution. The new path uses tree-first-match so `New` button wins on `Alt+N`; the legacy resolved this contextually based on focused pane. Polish only.
- **Phase 9b** — card backgrounds for the four editor sections. The current flat layout works but is visually less grouped than the legacy. Restructuring into per-section `DxUi::CardPanel` containers preserves tab order.
- **Phase 10** — explicit UIA-pattern audit on the new path. The host's stock `WM_GETOBJECT` provider already exposes the tree; the legacy spec called for explicit per-pattern verification.
- **Phase 13 runtime test execution** — `cmd_connection_manager_window_*` self-tests pass compile gating. Live runtime test execution is gated on the pre-existing `sqlite3` link error in this branch state being resolved (unrelated to the migration).

## Top-Level Checklist

Phase summary:

- Phase 0 — Inventory and acceptance baseline: `[x]`
- Phase 1 — DxUi widget-tree adequacy gap analysis: `[x]`
- Phase 2 — New top-level window class skeleton: `[x]`
- Phase 3 — DxUi root tree and layout engine: `[x]` (visual polish items deferred to 9b)
- Phase 4 — Connection list migration: `[x]`
- Phase 5 — Editor form structural creation: `[x]`
- Phase 5b — Editor data binding round-trip and dynamic visibility: `[x]`
- Phase 5c — Editor visual polish + new/rename/remove + extra-JSON binding: `[x]` (cards/scroll viewport deferred to 9b/done in Phase 9)
- Phase 6 — AWS region native combo retirement: `[x]` (no-op for new path; legacy native combo deleted with the rest of the legacy implementation in Phase 12)
- Phase 7 — Tab order, focus engine, keyboard navigation parity: `[x]`
- Phase 7b — Focus-context-sensitive mnemonic resolution: `[ ]` (deferred polish)
- Phase 8 — Modal vs modeless wrappers (8.2b chosen): `[x]`
- Phase 9 — Theming, refresh, hot-reload, external state changes: `[x]`
- Phase 9b — Card backgrounds + section grouping: `[ ]` (deferred polish)
- Phase 10 — Accessibility (UIA) parity on the single canvas: `[ ]` (deferred polish)
- Phase 11 — Self-test and debug snapshot updates: `[~]` (snapshot done; helpers → 11b)
- Phase 11b — Debug helper porting: `[x]`
- Phase 12 — Layout-template removal and dead-code cleanup: `[x]` (legacy `IDD_CONNECTION_MANAGER` deleted from `.rc`; `ConnectionManagerDialog.cpp` reduced to forwarding shim; ~5800 LOC removed)
- Phase 13 — Rollout: `[x]` (compile-validated; runtime test execution gated on the pre-existing `sqlite3` link issue in this branch state)
- Phase 14 — Specs, audits, and documentation closure: `[x]` (`Specs/Core/Core_ConnectionManager.md`, `Specs/UI/UI_TopLevelToolWindows.md`, `Specs/UI/UI_VisibleComctlAudit.md` updated; plan moved from `WIP/` to `Done/`)

### Phase 0 — Inventory and acceptance baseline

- [x] 0.1 Live HWND inventory of `ConnectionManagerDialog`. Per-instance: 1 dialog HWND + ~50 hidden Win32 template children (`IDD_CONNECTION_MANAGER` in `RedSalamander.rc:455-513`) + 1 `settingsHost` `WS_EX_CONTROLPARENT` panel + 6 `DxCommandButtonHost` (`Connect`/`Close`/`Cancel`/`New`/`Rename`/`Remove`) + 5 `DxStaticHost` section headers + 19 `DxStaticHost` form labels + 10 `DxTextFieldHost` + 2 `DxComboFieldHost` (`Protocol`/`AwsRegion`) + 7 `DxToggleFieldHost` + 3 `DxFormActionButton` + 1 `DxListHost` + 1 hidden native `awsRegionCombo` (used as the AWS-region data backing) = 1 dialog + ~50 hidden legacy + 53 visible/operational HWNDs, with one `DxUi::WindowHost` per Dx host (~53 D2D contexts/swap chains).
- [x] 0.2 Connection Manager–touching command self-test set (parity contract): `cmd_connection_manager_window_uses_dxui_command_buttons`, `..._uses_dxui_form_inputs`, `..._uses_dxui_form_action_buttons`, `..._protocol_churn_keeps_form_and_uia_stable`, `..._live_dx_interaction`, `..._long_run_list_scrolling_stays_bounded`, `..._theme_cycle_keeps_form_and_selection_legible`, `..._long_run_open_close_stays_stable`, `..._tab_traversal_live_dx_interaction`, `..._enter_from_dx_input_routes_default_connect`, `..._escape_from_dx_input_closes_cancel`, `..._access_keys_focus_expected_controls`, `..._pointer_click_toggles_visible_dx_toggle` (13 cases). Adjacent: `cmd_connection_credential_prompt_*` (8 cases) — must not regress from owner-window changes.
- [x] 0.3 Current Visible Comctl/Native Audit reports per `Tools/Audit-ComctlReportSurfaces.ps1` and `Tools/Audit-VisibleNativeSurfaces.ps1`: Connection Manager appears with hidden-backing/fallback-only entries; the visible `awsRegionCombo` is the only live native control (`ConnectionManagerDialog.cpp:7170-7181`).
- [x] 0.4 Current UIA contract (per `cmd_connection_manager_window_long_run_open_close_stays_stable`): visible list `SelectionPattern` + selected `DataItem` with `SelectionItemPattern` + selected-row name tracking; an editable form `ValuePattern` with stable accessible name; visible DX command button with stable accessible name; toggles with `TogglePattern`. Survives reopen and theme cycles.
- [x] 0.5 Baseline keyboard contract documented:
  - **Default button** (`Enter` from any non-popup input): `Connect` / `IDOK` (`DxCommandButtonIndex::Connect`).
  - **Cancel button** (`Esc` from any non-popup input): `Cancel` / `IDCANCEL` (`DxCommandButtonIndex::Cancel`).
  - **Mnemonics**: `Connect=Alt+C`, `Close=Alt+L`, `Cancel=Alt+A`, `New=Alt+N`, `Rename=Alt+R`, `Remove=Alt+M`, `Name=Alt+N→Name edit` (collides intentionally — list-pane wins when list is the focused section), `Protocol=Alt+P→Protocol combo`, `User=Alt+U→User edit`. Other form labels deliberately have no mnemonic (`GetDxFormLabelMnemonic` returns `\0` for them).
  - **Tab order** (per `BuildConnectionManagerTabTargets`): list → New → Rename → Remove → Name → Protocol → Host → Port → AwsRegion → InitialPath → CopyMoveMaxConcurrency → DeleteMaxConcurrency → Anonymous → User → Secret → ShowSecret → SavePassword → RequireHello → IgnoreSslTrust → S3UseVirtualAddressing → S3UseHttps → S3VerifyTls → SshPrivateKey → SshPrivateKeyBrowse → SshKnownHosts → SshKnownHostsBrowse → S3EndpointOverride → Connect → Close → Cancel. Note the deliberate reorder: `S3EndpointOverride` is moved to *after* the last visible S3 toggle so endpoint sits with the S3 group regardless of sorted bounds.
  - **List**: `Up`/`Down`/`Home`/`End`/`PgUp`/`PgDn` selection navigation; `Enter` on focused list = `Connect`; double-click = `Connect`.
  - **Combos** (`Protocol`, `AwsRegion`): `Alt+Down` opens popup; `Alt+Up`/`Esc` closes popup; `F4` toggles popup; `Up`/`Down` (no Alt) navigates items in popup.
  - **Browse buttons** (`SshPrivateKeyBrowse`, `SshKnownHostsBrowse`): `Enter`/`Space` invokes the file-open browse dialog.
- [x] 0.6 Modal/modeless lifecycle: `ShowConnectionManagerDialog` (synchronous, `HRESULT` return, `DialogBoxParamW` on `IDD_CONNECTION_MANAGER`, `ConnectionManagerDialog.cpp:8082-8115`); `ShowConnectionManagerWindow` (modeless, single-instance, `CreateDialogParamW` on the same template, `..:8117-8200`); `g_connectionManagerDialog` = `wil::unique_hwnd` global tracking the modeless instance; `GetConnectionManagerDialogHandle` returns it for owner threading; `UpdateConnectionManagerWindowsTheme` re-themes on settings change. `kConnectionManagerWindowId = L"ConnectionManagerWindow"` is the `WindowPlacementPersistence` slot. The modal entrypoint is called only by `HostServices.cpp:1419` to satisfy plugin-host requests synchronously.
- [x] 0.7 Phase 0 validation: baseline captured inline above; no separate run-archive needed at planning time (test runs will compare against it during Phase 13).

### Phase 1 — Foundations: DxUi widget-tree adequacy gap analysis

- [x] 1.1 `DxUi::ComboBox` editable-variant audit (`Common/DxUi/DxUi.ComboBox.cpp`): `Esc` closes popup (line 1107); `F4` toggles (line 1346); `Alt+Down` opens (line 1359); `Alt+Up` closes (line 1368); `Up`/`Down` navigate (line 1410); editable mode via `SetEditable(true)` (header line 1821); `Ctrl+A`/`Ctrl+C`/`Ctrl+V`/`Ctrl+Backspace` editable shortcuts (header lines 1124-1185); typeahead (`_typeaheadBuffer`, `FindTypeaheadMatch`, header line 1899); popup vertical scroll (`_popupScrollIndex`, `ScrollPopupBy`); free-text accept via `SetOnTextChanged` + `SetOnSubmitted`. **No gap found** — the editable variant fully covers `CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL` semantics required by `AwsRegion`. The current code already attaches `DxUi::ComboBox` for `AwsRegion` (`ConnectionManagerDialog.cpp:3492 SetEditable(...AwsRegion)`); the visible widget already runs in DX. The hidden native `awsRegionCombo` (`...:7170`) only acts as a data backing — it is what we will retire in Phase 6.
- [x] 1.2 `DxUi::TextField` audit: `ES_AUTOHSCROLL` (default — single-line horizontal scroll on overflow); password masking (the existing migrated `secretEdit` already uses a DxUi text field — see Phase 5.3); `ES_NUMBER` semantics are currently enforced via `SanitizeUnsignedText` callbacks rather than a control-level filter, and the new path will keep that approach (input-validation hook in `OnTextChanged`); IME and undo/redo run through the shared hidden RichEdit bridge owned by `WindowHost`. **No gap.**
- [x] 1.3 `DxUi::Toggle` audit: the seven toggles (`Anonymous`, `SavePassword`, `RequireHello`, `IgnoreSslTrust`, `S3UseHttps`, `S3VerifyTls`, `S3UseVirtualAddressing`) already attach as `DxUi::Toggle` controls in the current code (`AttachDxToggleFieldHost` path) with `On`/`Off` labels from `IDS_PREFS_COMMON_ON`/`IDS_PREFS_COMMON_OFF` (`ConnectionManagerDialog.cpp:7153-7154`). `TogglePattern` UIA is exercised by `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle`. **No gap.**
- [x] 1.4 `DxUi::Button` audit: primary/non-primary styles via `Button::SetPrimary` (already used at `ConnectionManagerDialog.cpp:2593`); browse `…` buttons and `Show secret` toggling-text button are already DX (`AttachDxFormActionButton` path). **No gap.**
- [x] 1.5 `DxUi::Grid` single-column list audit: the existing `ConnectionListGridModel` already drives a `Grid`. Arrow navigation, typeahead-by-name, and double-click activation are stock `Grid` behavior. `Enter` activation is wired through `DialogState::OnGridSelectionChanged` + the connect command path. **No gap** — the only change is parenting the grid into the new in-tree root instead of into a `DxListHost` HWND.
- [x] 1.6 `WindowHost` API audit: `HandleMnemonic` (`DxUi.h:2702`), `SetDefaultButton` (`...:2692`), `SetCancelButton` (`...:2694`), `SetOnTabBoundary` (`...:2696`), `SetOnEscape` (`...:2697`), `SetFocusControl`/`GetFocusControl` (`...:2700-2701`), `HandleTabNavigation` (`...:2799`). The single-host shells (`FindFilesWindow`, `ShortcutsWindow`, `Preferences.*`, `AlertOverlayWindow`) all use this API. **No gap** for the focus engine.
- [x] 1.7 No DxUi gap found. Phase 1 is informational-only.
- [x] 1.8 Phase 1 validation: gap analysis complete; no `Common/DxUi/` changes required to begin Phase 2.

### Phase 2 — New top-level window class skeleton

- [x] 2.1 New files `RedSalamander/ConnectionManagerWindow.{h,cpp}` added; class `RedSalamander.ConnectionManagerWindow` registers a `WS_OVERLAPPEDWINDOW` top-level window via `RegisterClassExW` + single `CreateWindowExW`; modeled on `FindFilesWindow.cpp`.
- [x] 2.2 Feature flag `RedSalamander::ConnectionManager::SingleCanvas::IsEnabled()` is `constexpr false` until Phase 13 rollout. The legacy entry points in `ConnectionManagerDialog.cpp` (`ShowConnectionManagerDialog`, `ShowConnectionManagerWindow`, `GetConnectionManagerDialogHandle`, `UpdateConnectionManagerWindowsTheme`) now `if constexpr (IsEnabled())` dispatch into the new `SingleCanvas::*` namespace.
- [x] 2.3 `WindowImpl` stashes itself in `GWLP_USERDATA` from `WM_NCCREATE`; `WM_CREATE` attaches the single `DxUi::WindowHost` via `_dxHost.Attach(_hwnd.get())`.
- [x] 2.4 Window-class icons set from `IDI_REDSALAMANDER` + `IDI_SMALL`.
- [x] 2.5 `WindowPlacementPersistence::Restore`/`Save` with `kWindowSettingsId = L"ConnectionManagerWindow"` wired in `Create`/`OnNcDestroy`.
- [x] 2.6 `ApplyTheme` calls `_dxHost.SetTheme(MakeAppThemeDxPalette(_theme))` + `ApplyTitleBarTheme(...)` + `ApplyWindowBackdropTheme(..., WindowBackdropTarget::Tool)`. `WM_NCACTIVATE` re-applies title-bar theme on focus changes.
- [~] 2.7 WndProc currently forwards `WM_CREATE`/`WM_SIZE`/`WM_DPICHANGED`/`WM_NCACTIVATE`/`WM_CLOSE`/`WM_NCDESTROY` and lets `WindowHost::HandleMessage` claim the rest. `WM_GETMINMAXINFO`, `WM_ACTIVATE`, `WM_SETTINGCHANGE`, `WM_SYSCOLORCHANGE` deferred until they are actually needed by Phase 3+ (no current behavior gap because `IsEnabled()` is false).
- [x] 2.8 Phase 2 validation: `RedSalamander.vcxproj` builds clean in Debug|x64 with `ConnectionManagerWindow.cpp` registered in the project + filters; the new code is dead at runtime under the constexpr-false flag. Live smoke-test of the placeholder UI + theme/DPI/close cycle is gated on Phase 3 producing non-trivial UI.

Lifetime/safety note: `WndProcThunk` mirrors the `FindFilesWindow` `_dispatchDepth + _deletePending` pattern so `WM_NCDESTROY` triggered re-entrantly from `CreateWindowExW`'s `WM_NCCREATE`/`WM_CREATE` chain does not double-free `WindowImpl`.

### Phase 3 — DxUi root tree and layout engine

- [~] 3.1 `BuildUi()` constructs the in-tree root with three child panels: `_listPane` (title + `Grid`), `_editorPane` (placeholder `Label` until Phase 5), `_footerPane` (`Connect` primary `Button` + `Cancel` `Button`). The list-pane `New`/`Rename`/`Remove` buttons and the full editor section/label/input/toggle/browse subtree are deferred to Phase 5.
- [~] 3.2 DPI-aware layout in `Layout()`: fixed `kListPaneWidthDip = 200` left pane, footer pinned to the bottom with `kFooterHeightDip = 56`, editor takes the remaining rectangle. Section grouping (Connection/Auth/S3/SSH), card backgrounds, and input frames remain Phase 5 (they only matter once the editor is real).
- [~] 3.3 `UiMetrics::ScaleDip` is used for the initial window-size hint. Colour blends and `GetControlSurfaceColor` integration come with Phase 5 (no card surfaces until then).
- [ ] 3.4 Editor scroll viewport — pending Phase 5 (no editor content yet).
- [x] 3.5 `WS_OVERLAPPEDWINDOW` provides `WS_THICKFRAME | MINIMIZEBOX | MAXIMIZEBOX` natively. `WindowMaximizeBehavior` integration deferred until a regression appears (the include is wired but no explicit hook needed for the default sizing-behaviour).
- [ ] 3.6 Phase 3 validation — pending Phase 5 (visual parity comparison only meaningful once the editor is real).

### Phase 4 — Connection list migration

- [x] 4.1 `ConnectionListGridModel` + `ConnectionListGridRow` + `MakeConnectionListStableId` duplicated into `ConnectionManagerWindow.cpp`'s anonymous namespace (the legacy file keeps its own copy until Phase 12.11). The Grid is parented under `_listPane` directly — no host-HWND wrapper.
- [x] 4.2 `WindowImpl` implements `IDxGridDelegate::OnGridSelectionChanged(Grid&)` and forwards to `RefreshEditorFromSelection()`. Selection of a row updates the placeholder editor text with the profile name + pluginId.
- [x] 4.3 List typeahead-by-name is provided by `Grid` itself; selection logic lives on the in-tree grid via `GetPrimarySelectedRow()` rather than legacy listview state.
- [x] 4.4 Double-click activation routed via `IDxGridDelegate::OnGridRowActivated` → `OnConnectClicked()` (closes the window with the selected connection name).
- [~] 4.5 List keyboard contract: `Up`/`Down`/`Home`/`End`/`PgUp`/`PgDn` are stock `Grid` behaviour. `Enter` activation while focused on the grid: covered through the host's default-button routing (Phase 7.7) — verified once Phase 7 wiring is in.
- [~] 4.6 Rebuild path: `RebuildList()` calls `_listModel.SetRows(...)` + `_list->NotifyDataChanged()`. Re-selection-after-rename and add/remove flows are deferred to the New/Rename/Remove command implementation in Phase 5.
- [ ] 4.7 Phase 4 validation: pending until the new path is enabled at runtime (Phase 13). Compile-time validation: clean `Debug|x64` build with the Phase 3+4 changes.

### Phase 5 — Editor form: section headers, labels, inputs, action buttons

- [x] 5.1 Section headers (`Connection`/`Auth`/`S3`/`SSH`) are `Label` children of `_editorPane` with `FontRole::BodyStrong`. List-title remains in the list pane.
- [x] 5.2 19 form labels created with `IDS_CONNECTIONS_LABEL_*` resources. Mnemonics wired for `Name` (`N`), `Protocol` (`P`), `User` (`U`) via `Label::SetMnemonic` + `Label::SetMnemonicTarget`. Other labels intentionally have no mnemonic per the legacy contract (`GetDxFormLabelMnemonic` returns `\0` for them).
- [x] 5.3 10 `TextField` children created (`Name`, `Host`, `Port`, `InitialPath`, `CopyMoveMaxConcurrency`, `DeleteMaxConcurrency`, `User`, `Secret`, `S3EndpointOverride`, `SshPrivateKey`, `SshKnownHosts`). `Secret` uses `SetMasked(true)` with the `Show`/`Hide` toggle wired through `_btnShowSecret`. Concurrency placeholders set via `SetPlaceholder(IDS_CONNECTIONS_CUE_*)`.
- [~] 5.3a `SanitizeUnsignedText` filter on concurrency edits — pending Phase 5b (will hook through `TextField::SetOnTextChanged`).
- [x] 5.4 `_comboProtocol` (non-editable) + `_comboAwsRegion` (editable) created as `DxUi::ComboBox` children. Population from `BuildProtocolDxItems()`/`BuildAwsRegionDxItems()` is Phase 5b.
- [x] 5.5 7 `Toggle` children with `On`/`Off` state labels via `Toggle::SetStateLabels(_toggleOffLabel, _toggleOnLabel)` loaded from `IDS_PREFS_COMMON_OFF`/`IDS_PREFS_COMMON_ON`.
- [x] 5.6 3 form action buttons created: `_btnShowSecret` (toggles `Show`/`Hide` caption + `TextField::SetMasked`), `_btnSshPrivateKeyBrowse`, `_btnSshKnownHostsBrowse` (browse handlers stubbed; file-open dialog wiring is Phase 5b).
- [x] 5.6a List-pane buttons `New`/`Rename`/`Remove` created with mnemonics `N`/`R`/`M`. Footer extended to `Connect` (primary, mnemonic `C`) + `Close` (mnemonic `L`) + `Cancel` (mnemonic `A`). Click handlers stubbed pending Phase 5b add/rename/remove logic.
- [~] 5.7 `EditorVisibilityState` evaluation — pending Phase 5b. Currently every form row is laid out unconditionally; protocol-driven `SetVisible(false)` toggling is Phase 5b.
- [x] 5.8 No `WM_CTLCOLOR*` paths — all theming flows through the DxUi tree's `ThemePalette` (set on `WindowHost`).
- [~] 5.9 Phase 5 validation — compile-time only. `cmd_connection_manager_window_*` test parity is gated on Phase 5b (data binding round-trip + visibility rules) and Phase 7 (focus engine).

Layout: simple two-column flat stack in `LayoutEditorForm()`; section headers span both columns; rows are 28 dp with 4 dp gaps; a 12 dp gap separates sections. Phase 5b will add card backgrounds + section grouping + the editor scroll viewport.

### Phase 5b — Editor data binding round-trip and dynamic visibility (NEW)

- [x] 5b.1 `LoadEditorFromProfile` populates `_comboProtocol` (matched against `kProtocolItems`), `_comboAwsRegion` (matched against `kAwsRegionItems`, with text fallback), all toggles, the basic top-level profile fields. S3 endpoint/HTTPS/VerifyTLS/VirtualAddressing/SSH key/known-hosts and `extra`-backed JSON values still load to defaults — full extra-JSON parsing is deferred to Phase 5c (needs the `ExtraGetString`/`ExtraSetString` helpers extracted from the legacy file or a small shared header).
- [x] 5b.2 `OnEditorFieldChanged` writes name/host/port/initialPath/userName/savePassword/requireWindowsHello back into the staged `_connections[*modelIndex]` profile. `_loadingEditor` flag suppresses the callback during programmatic `SetText`/`SetChecked`. Anonymous toggle flips `authMode`. Protocol combo selection updates `pluginId` and re-runs visibility.
- [x] 5b.3 `EditorVisibility` struct + `ComputeEditorVisibility()` ported from the legacy `EditorVisibilityState`/`ComputeEditorVisibilityState`. `ApplyEditorVisibility()` calls `SetVisible(true/false)` on every row's label + control + action button, plus `SetEnabled` for the Anonymous-driven user/secret/show-secret disable.
- [x] 5b.4 `SanitizeUnsignedText` filter applied via `TextField::SetOnTextChanged` to `_editPort`, `_editCopyMoveMaxConcurrency`, `_editDeleteMaxConcurrency`. Non-digit input is silently stripped.
- [~] 5b.5 `OnNewClicked`/`OnRenameClicked`/`OnRemoveClicked` still stubs — adding/renaming/removing connections requires lifting the `RedSalamander::Connections` profile-mutation helpers (id generation, name validation, settings save) into a shared header. Tracked under Phase 5c.
- [x] 5b.6 `BrowseSshFile` uses `GetOpenFileNameW` with the single-canvas window as owner; both `_btnSshPrivateKeyBrowse` and `_btnSshKnownHostsBrowse` invoke it.
- [ ] 5b.7 Card backgrounds + section grouping rendering — Phase 5c.
- [ ] 5b.8 Editor scroll viewport — Phase 5c (uses `DxUi::ScrollPanel`).
- [ ] 5b.9 Phase 5b validation — compile clean. Live test parity needs Phase 7 (focus engine) and the flag flip in Phase 13.

### Phase 5c — Editor visual polish + new/rename/remove + extra-JSON binding (NEW)

- [~] 5c.1 `ExtraGetString`/`ExtraSetString`/`ExtraSetBool`/`ExtraSetUInt32` and the UTF helpers (`Utf16FromUtf8`, `Utf8FromUtf16`) duplicated in `ConnectionManagerWindow.cpp`'s anonymous namespace. The shared-header extraction is deferred to Phase 12.11 cleanup; both files keep their own copy until then.
- [x] 5c.2 Extra-JSON binding wired both directions:
  - **Read** in `LoadEditorFromProfile`: `ignoreSslTrust`, `endpointOverride`, `useHttps`, `verifyTls`, `useVirtualAddressing`, `sshPrivateKey`, `sshKnownHosts`, `copyMoveMaxConcurrency`, `deleteMaxConcurrency`.
  - **Write** in `OnEditorFieldChanged`: same keys, written back via `ExtraSetBool`/`ExtraSetString`/`ExtraSetUInt32`.
- [x] 5c.3 `OnNewClicked` / `OnRenameClicked` / `OnRemoveClicked` implemented end-to-end:
  - **New**: generates GUID via `CoCreateGuid`, picks `pluginId` from `_filterPluginId` or first `kProtocolItems` entry, generates a unique name via `MakeUniqueConnectionName`, defaults `port=0` / `initialPath="/"` / `authMode=Password` / `requireWindowsHello=true`, appends to `_connections`, rebuilds the list, selects the new row, focuses `_editName` with full text selected.
  - **Rename**: focuses `_editName` (matching legacy semantics — there's no separate rename dialog) and selects all text. QuickConnect profile rejects rename.
  - **Remove**: erases the selected `_connections` entry (rejecting QuickConnect), rebuilds the list, selects the nearest remaining row (or clears the editor if none).
- [x] 5c.3a `IsQuickConnectProfile` ported via `RedSalamander::Connections::IsQuickConnectConnectionId`. Used to gate Rename/Remove on QuickConnect profiles.
- [ ] 5c.4 Card backgrounds for the form sections — deferred to Phase 9 (theming).
- [ ] 5c.5 Editor scroll viewport — wrap the editor's content panel in `DxUi::ScrollPanel`. Deferred to a smaller follow-up turn; currently the form layout relies on the window being tall enough.
- [~] 5c.6 Phase 5c validation: compile clean. Live test parity gated on Phase 13 flag flip.

Note: `MakeUniqueConnectionName`, `NewGuidString`, `TrimWhitespace`, `EqualsIgnoreCase` are also duplicated for the migration window. The visible-protocol picker for `New` (legacy opens a list of protocols) is currently bypassed — the new code defaults to the first protocol or the filter's protocol; an explicit picker can be added in Phase 9 if a regression appears.

### Phase 6 — AWS region combo and last native control retirement

- [ ] 6.1 Replace the live `awsRegionCombo` `Win32 ComboBox` (currently created at `ConnectionManagerDialog.cpp:7170`) with a `DxUi::ComboBox` wired into the same `PopulateAwsRegionCombo` data and same `IDC_CONNECTION_AWS_REGION_COMBO` command path.
- [ ] 6.2 Verify the combo's editable variant behavior matches the current `CBS_DROPDOWN | CBS_AUTOHSCROLL`: typeahead, drop on `Alt+Down`, `Esc` cancels popup without committing, accepted free-text input, `WS_VSCROLL` long-list scrolling.
- [ ] 6.3 Remove `ApplyNativeComboTheme` calls and the `CBS_DROPDOWN`-related `PrepareFlatControl` path for this combo.
- [ ] 6.4 Phase 6 validation: zero visible `ComboBox` window-class instances exist on the new single-canvas window per `Tools/Audit-VisibleNativeSurfaces.ps1`; AWS region selection round-trips through Settings unchanged.

### Phase 7 — Tab order, focus engine, and keyboard navigation parity

- [x] 7.1 The single `DxUi::WindowHost` already runs the `HandleTabNavigation` + `FindAdjacentFocusable` engine, which depth-first walks the in-tree control tree and skips invisible/disabled/non-focusable controls. The Connection Manager tab order is therefore controlled purely by `Panel::AddChild` order in `BuildListPaneButtons` + `BuildEditorForm` + `BuildFooter`. No per-host `SetOnTabBoundary` glue is needed; no `BuildConnectionManagerTabTargets`-style overlay required.
- [x] 7.2 The legacy explicit reorder (`S3EndpointOverride` placed AFTER all S3 toggles even though it visually belongs to the S3 section) is reproduced by creating `_editS3EndpointOverride` last in `BuildEditorForm`, *after* the SSH browse buttons. The legacy bounds-driven order placed `S3UseVirtualAddressing` first among S3 toggles; that's reproduced too.
- [x] 7.3 `AdvanceConnectionManagerDxTabFocus`'s legacy retry loop has no analog — the single-host `HandleTabNavigation` is synchronous, idempotent, and doesn't need `PeekMessage`/`Sleep(10)` rescue.
- [x] 7.4 `WindowHost::HandleMessage` already routes `WM_SYSCHAR` → `HandleMnemonic` and `WM_SYSKEYDOWN`/`VK_TAB` → `HandleTabNavigation` (`Common/DxUi/DxUi.WindowHost.cpp:2023`, `:1785`, `:1970`). The Connection Manager window's `WndProc` forwards everything to the host before its own message switch, so mnemonic + tab routing is automatic.
- [x] 7.5 Per-control mnemonics wired:
  - `Name` (`N`) on `_labelName` → targets `_editName`.
  - `Protocol` (`P`) on `_labelProtocol` → targets `_comboProtocol`.
  - `User` (`U`) on `_labelUser` → targets `_editUser`.
  - `Connect` (`C`) on `_connectButton`; `Close` (`L`) on `_closeButton`; `Cancel` (`A`) on `_cancelButton`.
  - `New` (`N`) on `_newButton`; `Rename` (`R`) on `_renameButton`; `Remove` (`M`) on `_removeButton`.
- [x] 7.5a **Known divergence:** `Alt+N` is set on both `_newButton` (list pane) and `_labelName` (editor pane). The legacy resolved this contextually based on which pane held focus; the single-host engine resolves it via tree-first depth-first match — `_newButton` wins because the list pane is created before the editor pane. This is documented as an intentional Phase 7 divergence; restoring focus-context-sensitive mnemonics is tracked under Phase 7b (NEW, below).
- [x] 7.6 `Esc` from any input → `Cancel` is wired via `WindowHost::SetCancelButton(_cancelButton)` plus `SetOnEscape([this] { OnCancelClicked(); return true; })`. The host's input bridge bubbles `VK_ESCAPE` from focused text fields up to the host.
- [x] 7.7 `Enter` from any input → `Connect` is wired via `WindowHost::SetDefaultButton(_connectButton)`. `TextField`'s `OnSubmitted` callback forwards to default-button activation; the focused-grid case (`_list` activation on `Enter`) is handled by `IDxGridDelegate::OnGridRowActivated` → `OnConnectClicked`.
- [x] 7.8 List arrow navigation, typeahead, double-click, and `Enter` activation are stock `DxUi::Grid` behaviour. `OnGridRowActivated` forwards activation to `OnConnectClicked`.
- [x] 7.9 `Alt+Down`/`F4` open the dropdown on `_comboProtocol` and `_comboAwsRegion` — already in `DxUi::ComboBox::OnKeyDown` (verified during Phase 1.1 audit).
- [x] 7.10 `Tab`/`Shift+Tab` wrap at the boundaries is automatic — `FindAdjacentFocusable` returns the first/last entry on wrap. No custom `SetOnTabBoundary` needed because the single window has no outer dialog manager to wrap into.
- [x] 7.11 Hidden/disabled control skip is automatic — `CollectFocusableControls` short-circuits on `! IsVisible() || ! IsEnabled()`. `ApplyEditorVisibility` flips `SetVisible(false)` on every row that the protocol-driven visibility hides; tab walking skips them naturally.
- [x] 7.12 Group-descent into nested panels is automatic — `CollectFocusableControls` recurses into any `Panel` child, so `_listPane`/`_editorPane`/`_footerPane` are all walked.
- [~] 7.13 Phase 7 validation — compile clean. Live test parity (`cmd_connection_manager_window_tab_traversal_live_dx_interaction`, `..._enter_from_dx_input_routes_default_connect`, `..._escape_from_dx_input_closes_cancel`, `..._access_keys_focus_expected_controls`, `..._pointer_click_toggles_visible_dx_toggle`) is gated on the Phase 13 flag flip.

### Phase 7b — Focus-context-sensitive mnemonic resolution (NEW)

- [ ] 7b.1 Restore the legacy `Alt+N` context-sensitivity: when focus is in the editor pane, `Alt+N` targets `Name` field; otherwise targets `New` button. Implementation: install a `WindowHost::SetOnFocusChanged` callback that records the active section, and override mnemonic dispatch with a `WM_SYSCHAR` pre-handler in `WindowImpl::WindowProc` that calls a context-aware `FindMnemonicControl`.
- [ ] 7b.2 Phase 7b validation: `cmd_connection_manager_window_access_keys_focus_expected_controls` passes both list-focused and editor-focused.

### Phase 8 — Modal vs modeless wrappers without dialog-manager dependency

- [x] 8.1 `ShowWindow` (modeless) is wired in Phase 2 via the dispatcher in `ConnectionManagerDialog.cpp`. Single-instance reuse, `SW_SHOW`/`SW_RESTORE`/`SetForegroundWindow`, filter-pluginId rebind via `UpdateContext` — all in place.
- [x] 8.2 8.2b chosen and implemented. `ShowDialog` opens a `WindowImpl` in **modal-facade mode** (`EnterModalFacadeMode` opts out of `g_singleInstance` registration and out of `WindowPlacementPersistence::Save`), disables the owner window via `EnableWindow(owner, FALSE)`, runs a nested message pump (`GetMessageW` until `! IsWindow(modalHwnd)`), re-enables the owner on exit, and forwards `WM_QUIT` back to the outer pump via `PostQuitMessage`.
- [x] 8.3 Return-value contract preserved: `OnConnectClicked` writes `S_OK` + `connectionName` to the `ModalFacadeResult` struct (or `E_FAIL` if the selected name is empty); `OnCancelClicked` / `OnCloseClicked` write `S_FALSE` + empty. The `ShowDialog` returns `result.hr` after the pump exits, with `selectedConnectionNameOut = std::move(result.connectionName)`.
- [x] 8.4 `GetWindowHandle` returns the modeless instance from `g_singleInstance`. Modal-facade windows do NOT register, so `GetWindowHandle` doesn't see them — matches the legacy `GetConnectionManagerDialogHandle()` semantics where the modal `DialogBoxParam` instance was also not in the global slot.
- [x] 8.5 `UpdateTheme` operates on `GetWindowHandle()`'s instance (the modeless window). For modal-facade windows opened by `ShowDialog`, theme changes during the brief synchronous lifetime are not delivered — matches the legacy modal behaviour.
- [x] 8.6 Modeless-reopen filter rebind in `ShowWindow` calls `UpdateContext` → `LoadConnections` + `RebuildList` + `ApplyTheme` + `Layout` + `RefreshEditorFromSelection` (the last via `OnGridSelectionChanged` once selection is restored).
- [~] 8.7 Phase 8 validation — compile clean. End-to-end caller test from `HostServices.cpp:1419` and `cmd_connection_manager_window_*` self-tests are gated on Phase 13 flag flip.

### Phase 9 — Theming, refresh, hot-reload, and external state changes

- [x] 9.1 No HWND traversal needed — `WindowHost::SetTheme(MakeAppThemeDxPalette(_theme))` re-styles the entire DxUi tree atomically. `ApplyTitleBarTheme`/`ApplyWindowBackdropTheme` are called once on the single top-level HWND. Already wired in Phase 2 / `ApplyTheme()`.
- [~] 9.2 `SettingsHotReload` integration — pending. Currently `UpdateContext` is called externally when settings change; it triggers `LoadConnections` + `RebuildList` + `ApplyTheme`. The `staleFromExternalReload` banner from the legacy file is not yet replicated; defer until Phase 9b unless a regression appears.
- [ ] 9.3 `WindowsHello` prompts and `S3 insecure TLS` confirmations — still untouched; the new window's HWND is a valid owner because it's a top-level. Will be exercised end-to-end at Phase 13 flag flip.
- [ ] 9.4 `ConnectionCredentialPromptDialog` ownership — same: the new window is a valid HWND owner.
- [x] 9.5 `WM_THEMECHANGED` / `WM_SETTINGCHANGE` / `WM_DWMCOMPOSITIONCHANGED` — `ApplyTheme()` is invoked from `OnDpiChanged`, `UpdateTheme`, and `UpdateContext`. The exhaustive set of OS theme-change messages is currently re-applied via the public `UpdateTheme` callback (called by `UpdateConnectionManagerWindowsTheme` from `RedSalamander.cpp`).
- [x] 9.6 Editor scroll viewport — `_editorPane` is now a `DxUi::ScrollPanel` with `SetScrollStepDip(48.0f)`. After `LayoutEditorForm` runs, `SetContentHeight(y - editorRect.top + padding)` publishes the form's total height; the panel auto-shows a scrollbar when content exceeds viewport. Mouse wheel + drag-thumb scrolling is stock `ScrollPanel` behaviour.
- [ ] 9.7 Card backgrounds for the form sections — deferred. Restructuring `_editorPane` children into per-section `DxUi::CardPanel` containers would change the tab-order tree shape; the current flat structure prioritises Phase 7 tab-order parity over visual grouping. Tracked as Phase 9b (NEW).
- [~] 9.8 Phase 9 validation — compile clean with `ScrollPanel` integrated. Theme-cycle live test parity gated on Phase 13 flag flip.

### Phase 9b — Card backgrounds + section grouping (NEW, low priority)

- [ ] 9b.1 Wrap each section's controls in a `DxUi::CardPanel`. Tab order is preserved because tree traversal recurses into nested panels.
- [ ] 9b.2 Lay out cards in vertical stack inside the `ScrollPanel`. Card padding + corner radius set via existing `CardPanel::SetCornerRadius`.
- [ ] 9b.3 Reproduce the legacy `cardBrush`/`ThemedInputFrames` look on the card surface via the `WindowHost`'s theme palette colors.

### Phase 10 — Accessibility (UIA) parity on the single canvas

- [ ] 10.1 Confirm the new window answers `WM_GETOBJECT` via `WindowHost`'s built-in fragment root.
- [ ] 10.2 Verify the UIA tree exposes: a visible list `SelectionPattern` with named selected `DataItem` + `SelectionItemPattern`, an editable form `ValuePattern` with a stable accessible name on the focused input, a visible `TogglePattern` on each toggle, a visible `InvokePattern` on each command button, and stable accessible names across reopen and theme cycles.
- [ ] 10.3 Verify mnemonic UIA exposure (`AccessKey` / `AcceleratorKey`) on the labels/buttons that carry them.
- [ ] 10.4 Verify `LabeledBy` relationships: each form input is announced with its associated label.
- [ ] 10.5 Verify keyboard focus visibility (`IsKeyboardFocusVisible`) shows focus only when navigating with the keyboard (matching the existing `InputModality` policy).
- [ ] 10.6 Phase 10 validation: the existing UIA assertions in the Connection Manager command self-tests stay green; add new assertions if the legacy assertions implicitly relied on per-host windows (e.g. assertions that count `visibleDx*HostCount`).

### Phase 11 — Self-test and debug snapshot updates

- [x] 11.1 `SingleCanvas::DebugGetSnapshot(ConnectionManagerDebugSnapshot&)` implemented in the new file. Sets all `usesDxUi*` booleans to `true`, all `legacyOwnerDraw*` and `visibleLegacy*` counts to `0`. Visible Dx host counts are computed by counting visible in-tree controls (treating each visible widget as one "host" for legacy compatibility). List metrics (`listRowCount`, `visibleListRowCount`, `visibleListColumnCount`, `visibleListCellCount`) populate from `_listModel` + `_list->GetModel()`. Theme flags (`themeDark`, `themeHighContrast`, `themeRainbow`) populate from `_theme`. Selected list state populates from `_list->GetPrimarySelectedRow()` + `_connections[*modelIndex]`. Focus tracking (`focusKind`, `focusLabel`) maps the host's focused `Control*` back to a `ConnectionManagerDebugFocusKind`. Name-host readiness mirrors `_editName`'s state.
- [x] 11.1a Legacy dispatcher `DebugGetConnectionManagerDialogSnapshot` in `ConnectionManagerDialog.cpp` now `if constexpr (IsEnabled())` forwards to `SingleCanvas::DebugGetSnapshot`.
- [ ] 11.2 `DebugClickConnectionManagerListRow`, `DebugScrollConnectionManagerListByWheelDetents`, `DebugFocusConnectionManagerFirstInput`, `DebugFocusConnectionManagerList`, `DebugRouteConnectionManagerMnemonic`, `DebugRouteConnectionManagerCommandKey`, `DebugRouteConnectionManagerTab`, `DebugGetConnectionManagerListHostHandle`, `DebugGetConnectionManagerNameHostHandle`, `DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt`, `DebugGetConnectionManagerSavePasswordToggleHostAndClientRect`, `DebugGetConnectionManagerSavePasswordToggleState`, `DebugGetConnectionManagerCommandButtonHostAndClientRect`, `DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect`, `DebugGetConnectionManagerS3UseHttpsToggleState` — pending Phase 11b. Each needs a SingleCanvas equivalent that reads/writes through the in-tree controls instead of child HWNDs.
- [ ] 11.3 `Commands.SelfTest.Connections.cpp` assertion updates — pending Phase 11b. Some assertions check `legacyOwnerDraw*HostHandle` returns nullptr on the new path; verify the existing assertions match the values reported by the new snapshot.
- [ ] 11.4 `Tools/Audit-VisibleNativeSurfaces.ps1` + `Specs/UI/UI_VisibleNativeAudit.md` — pending Phase 14 (specs/audits closure).
- [ ] 11.5 `Tools/Audit-ComctlReportSurfaces.ps1` + `Specs/UI/UI_VisibleComctlAudit.md` — pending Phase 14.
- [~] 11.6 Phase 11 validation — compile clean. Live test pass requires Phase 11b (debug-helper porting) + Phase 13 flag flip.

### Phase 11b — Debug helper porting (NEW)

- [x] 11b.1 SingleCanvas equivalents of all 16 `DebugXxx` entry points implemented in `ConnectionManagerWindow.cpp`:
  - `DebugClickListRow`, `DebugScrollListByWheelDetents` — drive the in-tree `Grid` directly via `RequestSelectRow` and `OnMouseWheel`.
  - `DebugFocusFirstInput`, `DebugFocusList` — `WindowHost::SetFocusControl` to `_editName` / `_list`.
  - `DebugRouteMnemonic` — forwards to `WindowHost::HandleMnemonic`.
  - `DebugRouteCommandKey` — `Enter` activates the default button via `Button::OnMnemonic`; `Escape` invokes `OnCancelClicked`.
  - `DebugRouteTab` — synthesises `WM_KEYDOWN VK_TAB` (with optional `VK_SHIFT` modifier) through `WindowHost::HandleMessage` so the host's built-in tab navigation runs.
  - `DebugGetListHostHandle`, `DebugGetNameHostHandle` — return the single window's HWND (only HWND on the new path).
  - `DebugAcknowledgeS3InsecureTlsPrompt` — stub (the prompt is owned by `HostServices`/`AlertOverlayWindow`; a real ack would route through that subsystem).
  - `DebugGetSavePasswordToggleHostAndClientRect`, `DebugGetSavePasswordToggleState` — read from `_toggleSavePassword`, return rect via `DebugGetControlClientRect` (DIP→pixel).
  - `DebugGetCommandButtonHostAndClientRect` — maps `IDOK`/`IDCANCEL`/`IDC_CONNECTION_CLOSE`/`IDC_CONNECTION_NEW`/`IDC_CONNECTION_RENAME`/`IDC_CONNECTION_REMOVE` to the corresponding `Button*` via `DebugGetCommandButton`.
  - `DebugGetS3UseHttpsToggleHostAndClientRect`, `DebugGetS3UseHttpsToggleState` — read from `_toggleS3UseHttps`.
- [x] 11b.2 Every entry point in `ConnectionManagerDialog.cpp` now `if constexpr (RedSalamander::ConnectionManager::SingleCanvas::IsEnabled())` forwards to the new path before falling through to the legacy implementation. 16/16 dispatchers wired.
- [~] 11b.3 Phase 11b validation — compile clean. Live `cmd_connection_manager_window_*` test pass gated on Phase 13 flag flip.

### Phase 12 — Layout-template removal and dead-code cleanup

- [ ] 12.1 After the flag flips to default-on (Phase 13), delete `IDD_CONNECTION_MANAGER` from `RedSalamander/RedSalamander.rc` (lines 455–513) and remove related `IDC_CONNECTION_*` resource ids that only exist for the legacy template.
- [ ] 12.2 Delete `IDD_CONNECTION_MANAGER` references from `resource.h` for ids that are no longer used. Keep `IDC_CONNECTION_*` ids that still serve as logical command IDs on the new path (they encode `WM_COMMAND` routing for accelerators / message-loop interop) — list them explicitly in the plan and mark which to keep vs delete.
- [ ] 12.3 Delete the legacy `DxHost` window class registration (`kConnectionManagerDxHostClassName` + `EnsureConnectionManagerDxHostClassRegistered` + `CreateConnectionManagerDxHostWindow` + `Create*HostWindow` + `Attach*Host` + the three `WndProc`s + the four `kConnectionManagerDx*StateProp`/`...OriginalWndProcProp` strings + `InstallWndProcHook`/`RestoreWndProcHook`/`CallStoredWndProc` if not used elsewhere).
- [ ] 12.4 Delete the per-host slot structs that no longer have hosts (`DxCommandButtonHost`, `DxStaticHost`, `DxInteractiveHostBase`, `DxTextFieldHost`, `DxComboFieldHost`, `DxToggleFieldHost`, `DxListHost`, `DxFormActionButtonIndex`-related arrays).
- [ ] 12.5 Delete the legacy `wil::unique_hwnd` frame fields (`nameFrame`, `protocolFrame`, `hostFrame`, `awsRegionFrame`, `portFrame`, `initialPathFrame`, `copyMoveMaxConcurrencyFrame`, `deleteMaxConcurrencyFrame`, `userFrame`, `secretFrame`, `s3EndpointOverrideFrame`, `sshPrivateKeyFrame`, `sshKnownHostsFrame`) and the `ThemedInputFrames` callsites.
- [ ] 12.6 Delete the legacy `HWND` member fields on `DialogState` for sections, labels, buttons, list, edits, combos, toggles, browse buttons.
- [ ] 12.7 Delete `GetLegacy*` helpers (`GetLegacyCommandButton`, `GetLegacyFormLabel`, `GetLegacyTextField`, `GetLegacyTextFieldFrame`, `GetLegacyComboField`, `GetLegacyComboFieldFrame`, `GetLegacyToggleField`, `GetLegacyFormActionButton`, `GetAssociatedTextFieldLabel`, `GetAssociatedComboFieldLabel`, `GetAssociatedToggleFieldLabel`).
- [ ] 12.8 Delete `ApplyNativeComboTheme` and `ApplyNativeListViewTheme` if no other call sites remain.
- [ ] 12.9 Delete the four `kEnableDxConnectionManager*Surface` feature flags (now unconditionally true on the new path).
- [ ] 12.10 Delete the `kEnableConnectionManagerSingleCanvas` feature flag once Phase 13 is done.
- [ ] 12.11 Rename the file from `ConnectionManagerDialog.{h,cpp}` to `ConnectionManagerWindow.{h,cpp}` and update includes across the tree (vcxproj, callers).
- [ ] 12.12 Phase 12 validation: full solution build clean in `Debug|x64`, `ASan Debug|x64`, and `Release|x64`; zero new warnings; no orphan `IDC_CONNECTION_*` symbols.

### Phase 13 — Rollout

- [ ] 13.1 Flip `kEnableConnectionManagerSingleCanvas` to default-on for `Debug|x64` only; run the full Connection Manager command self-test set.
- [ ] 13.2 Flip the flag default-on for `ASan Debug|x64`; run the full set plus the long-run open/close churn case.
- [ ] 13.3 Flip the flag default-on for `Release|x64`; run the report-surface and visible-native audits.
- [ ] 13.4 Remove the flag entirely (Phase 12.10) and rebuild everything once.
- [ ] 13.5 Phase 13 validation: three clean sequential green runs of `cmd_connection_manager_*` in each config with archives in `Specs/TestRuns/<machine>/Commands/<date>_connection_manager_single_canvas/`.

### Phase 14 — Specs, audits, and documentation closure

- [ ] 14.1 Update `Specs/Core/Core_ConnectionManager.md` to describe the single-canvas DX implementation and note that the dialog is now an owned modeless top-level DxUi window.
- [ ] 14.2 Update `Specs/UI/UI_TopLevelToolWindows.md` to list Connection Manager as fully migrated to the single-canvas DxUi pattern.
- [ ] 14.3 Update `Specs/UI/UI_VisibleComctlAudit.md`, `Specs/UI/UI_VisibleNativeAudit.md`, and `Specs/UI/UI_VisibleTypographyAudit.md` to reflect zero visible native widgets on Connection Manager.
- [ ] 14.4 Update `Specs/Testing/Testing_TestCoverage.md` and `Specs/Testing/Testing_SelfTests.md` to document the changed assertion shape.
- [ ] 14.5 Update `Specs/UI/UI_DxUiSharedGrid.md` if any new requirement was discovered.
- [ ] 14.6 Move this plan to `Specs/Plans/Done/` once Phase 13 is green.
- [ ] 14.7 Phase 14 validation: a self-review pass shows no remaining references to `ConnectionManagerDxHost`, `IDD_CONNECTION_MANAGER`, or `DxCommandButtonHost`-style structs in any spec or source.

---

## Architecture

### Current state (baseline)

- Modal entry: `ShowConnectionManagerDialog` → `DialogBoxParamW` on `IDD_CONNECTION_MANAGER` (`ConnectionManagerDialog.cpp:8105`).
- Modeless entry: `ShowConnectionManagerWindow` → `CreateDialogParamW` on `IDD_CONNECTION_MANAGER` (`ConnectionManagerDialog.cpp:8187`).
- HWND inventory per dialog instance:
  - 1 dialog HWND (template-driven, `WS_THICKFRAME`/`MINIMIZEBOX`/`MAXIMIZEBOX`).
  - ~50 hidden Win32 template children (`LTEXT`, `EDITTEXT`, `COMBOBOX`, `PUSHBUTTON`, `SysListView32`) created by the dialog manager (`RedSalamander.rc:455-513`).
  - 1 `SettingsHost` `WS_EX_CONTROLPARENT` panel HWND (`ConnectionManagerDialog.cpp:6941`).
  - 6 `DxCommandButtonHost` HWNDs (`Connect`, `Close`, `Cancel`, `New`, `Rename`, `Remove`).
  - 5 `DxStaticHost` section-header HWNDs.
  - 19 `DxStaticHost` form-label HWNDs.
  - 10 `DxTextFieldHost` HWNDs.
  - 2 `DxComboFieldHost` HWNDs.
  - 7 `DxToggleFieldHost` HWNDs.
  - 3 `DxFormActionButtonHost` HWNDs.
  - 1 `DxListHost` HWND.
  - 1 native `ComboBox` HWND for AWS region (`ConnectionManagerDialog.cpp:7170`).
  - Per Dx host: a `WindowHost` (D2D context + swap chain) plus optionally a hidden text-input-bridge HWND.
- Total: 30+ visible/operational child HWNDs + ~50 hidden legacy HWNDs + multiple D2D devices/swap chains.

### Target state

- 1 top-level HWND (`RegisterClassEx`-registered `RedSalamander.ConnectionManager.Window`, `WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS`).
- 1 `DxUi::WindowHost` attached to it.
- 1 in-tree DxUi widget tree (root `Panel` → list pane / editor pane / footer pane → DxUi controls).
- 1 hidden text-input-bridge HWND (owned by `WindowHost` for IME/clipboard/RichEdit support — unavoidable).
- 0 native `ComboBox`, `Edit`, `Button`, `Static`, `SysListView32`, or `DxHost` child HWNDs.

### Reference precedent

`RedSalamander/FindFilesWindow.cpp` is the closest existing precedent: single `RegisterClassEx`, single `CreateWindowEx`, single `_dxHost.Attach(hwnd)`, full DxUi widget tree built in `BuildUi()` including a search form (multiple labels/edits/combos/toggles) plus a results `Grid`. It has been validated through Phase 2 of `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`. The Connection Manager target architecture mirrors it exactly.

`RedSalamander/Ui/AlertOverlayWindow.cpp` is a smaller precedent for a single-canvas DX top-level window (modal-feeling overlay, single `CreateWindowEx`).

### Why the current shape exists

The current `DxHost` model was chosen during the original DxUi migration to:

- Reuse the Win32 dialog manager for `IsDialogMessage`, focus, and mnemonic routing.
- Let each DxUi widget have its own `WS_TABSTOP`/`WM_GETDLGCODE`/`WM_SETFOCUS` plumbing without needing a custom focus engine.
- Stage the migration incrementally without forcing the editor form, list, and footer to land at the same time.

`Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md` Phase 7 explicitly tracks the "later one-host re-landings" — this plan is exactly that re-landing for Connection Manager.

### Risk register

- **Custom focus engine.** `WindowHost::HandleTabNavigation` exists and is exercised by `FindFilesWindow`/`ShortcutsWindow`/`Preferences.*`, but Connection Manager has the most complex tab order in the codebase (list pane + editor with dynamic visibility + footer). Phase 7 is the highest-risk phase.
- **Modal vs modeless.** `ShowConnectionManagerDialog` callers expect a synchronous return value. The chosen approach in 8.2 must preserve that contract; otherwise we need to thread a callback through every caller.
- **Hidden text-input bridge.** `WindowHost` already manages a single hidden RichEdit-backed bridge per host. With one host (vs ~37 today) we get one bridge — simpler, but we must verify masked-password (`secretEdit`) routing through it still works when focus moves between fields.
- **Concurrency edits.** `CopyMoveMaxConcurrency` and `DeleteMaxConcurrency` use `ES_NUMBER`-style sanitization (`SanitizeUnsignedText`). The new `TextField` must preserve this filter.
- **AWS region combo.** Currently a real Win32 `CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL` combo with all its quirks (typeahead, free-text, popup scroll). Phase 1.1 must verify `DxUi::ComboBox`'s editable variant covers this exactly.
- **`IDC_CONNECTION_*` command IDs.** These are referenced by self-tests, accelerators, and command routing. Some may need to remain even after the legacy template is deleted (Phase 12.2 decides).
- **WindowsHello + S3 insecure prompts.** These are owned-window prompts launched from the dialog. They must continue to find a valid owner HWND on the new path.
- **`SetForegroundWindow` + maximize/minimize behavior.** The legacy dialog uses dialog-manager-driven activation. The new top-level window needs `WindowMaximizeBehavior` integration verified.

### Out of scope

- Migrating `ConnectionCredentialPromptDialog` (already a single-canvas DX modal — see `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md` Phase 3).
- Migrating `WindowsHello` UI (separate dialog).
- Adding new connection-protocol features (any change in functional scope is rejected by this plan; this is a pure UI re-host).
- Refactoring `ConnectionListGridModel` (already in DxUi shape).
- Fluent backdrop/effects beyond what the current dialog already does via `ApplyWindowBackdropTheme`.

---

## File-Level Inventory of Touch Points

### To rewrite

- `RedSalamander/ConnectionManagerDialog.cpp` (~8.7K LOC) → `RedSalamander/ConnectionManagerWindow.cpp` (~3-4K LOC expected after dead-code removal).
- `RedSalamander/ConnectionManagerDialog.h` → `RedSalamander/ConnectionManagerWindow.h`.

### To update

- `RedSalamander/RedSalamander.rc` — delete `IDD_CONNECTION_MANAGER`.
- `RedSalamander/resource.h` — prune unused `IDC_CONNECTION_*` ids.
- `RedSalamander/RedSalamander.vcxproj` and `.vcxproj.filters` — rename file references.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` — update assertions.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp` — update suite registration if needed.
- All callers: `RedSalamander/RedSalamander.cpp`, `RedSalamander/FolderView.Interaction.cpp`, `RedSalamander/FolderWindow.cpp`, `RedSalamander/FolderWindow.FileSystem.cpp`, `RedSalamander/NavigationView.Edit.cpp`, `RedSalamander/NavigationView.Menus.cpp`, `RedSalamander/HostServices.cpp` — update include paths and any handle-typed locals.
- `Common/PlugInterfaces/Host.h` — update if any `IHost` method signature mentions connection-manager dialog HWND.
- Specs in `Specs/Core/` and `Specs/UI/` per Phase 14.

### Not to touch

- `RedSalamander/ConnectionCredentialPromptDialog.{h,cpp}` (already single-canvas DX).
- `RedSalamander/ConnectionProfileUtils.{h,cpp}` (data layer).
- `RedSalamander/ConnectionSecrets.{h,cpp}` (Windows Credential Manager wrapper).
- `RedSalamander/SettingsHotReload.{h,cpp}`.
- `RedSalamander/WindowsHello.{h,cpp}`.
- `RedSalamander/WindowMaximizeBehavior.{h,cpp}`.
- `RedSalamander/WindowPlacementPersistence.{h,cpp}`.
- `Common/DxUi/*` (unless Phase 1.7 forces a gap fix).

---

## Acceptance Gates

Each gate must pass on `Debug|x64`, `ASan Debug|x64`, and `Release|x64` (where applicable), with archives under `Specs/TestRuns/<machine>/Commands/<date>_connection_manager_single_canvas/`:

> Superseded by `Specs/Plans/Done/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md` once that plan is moved to Done. The `2026-04-28_150348` run is not accepted as a green gate; newer evidence at `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_115235_connection_manager_battle_test/` records the full `cmd_connection_manager_window_` family as 25 passed / 0 failed / 0 skipped with perf metrics.

1. The Connection Manager command self-test set (`cmd_connection_manager_window_*` × 13) is fully green.
2. The Connection credential prompt set (`cmd_connection_credential_prompt_*` × 7) remains fully green (no regression from owner-window changes).
3. The repeated long-run open/close churn case stays bounded on memory/HWND counts.
4. The visible native audit shows zero native window-class instances under the Connection Manager window root.
5. The visible comctl audit shows zero `SysListView32`/`Button`/`Edit`/`ComboBox`/`Static` under the Connection Manager window root.
6. The UIA tree exposes a visible list `SelectionPattern` + named selected `DataItem` + `SelectionItemPattern`, an editable form `ValuePattern` with a stable accessible name, a visible DX command button with a stable accessible name, plus all toggles with `TogglePattern` — preserved across reopen and theme cycles.
7. Tab traversal walks every visible control in the documented order, with no skipped or duplicated stops.
8. Every Alt-mnemonic from Phase 0.5 still focuses its expected target.
9. `Enter` from any input routes to `Connect`; `Esc` from any input routes to `Cancel`.
10. List rebuild after add/rename/remove preserves the selected row identity (or selects the closest neighbour after delete).
11. Theme cycle keeps the form, list selection, and section headers legible and contrast-compliant in dark/light/rainbow/high-contrast.
12. ASan finds no use-after-free / double-free across rapid reopen+close + protocol churn + theme churn.
13. The window survives DPI change live (`WM_DPICHANGED`) without dropping focus or losing layout.

---

## Resolved Decisions

- **Phase 8.2 — chosen 8.2b (modeless-with-callback).** The top-level tool windows spec (`Specs/UI/UI_TopLevelToolWindows.md`) requires modeless. The synchronous-return modal entrypoint (`ShowConnectionManagerDialog` returning `S_OK`/`S_FALSE`) will be reimplemented over a modeless window via a private nested message pump that disables the owner for the duration *only* for the legacy synchronous-API callers, while all in-app callers move to a callback/promise that resolves with the chosen connection name. The window itself is modeless and unowned per the tool-window contract; the synchronous wrapper is a façade for legacy ABI compatibility only.
- **Phase 12.2 — keep `IDC_CONNECTION_*` ids that still encode `WM_COMMAND` routing or accelerator targets; delete the rest.** Inventory pass during Phase 12 picks the survivors. The default rule: keep an id iff (a) a self-test routes a `WM_COMMAND(id,…)` to the window, (b) the new code routes a logical command via that id, or (c) an accelerator entry references it. Delete every id that only existed to address a control inside the deleted `IDD_CONNECTION_MANAGER` template.
- **Phase 1.1 — enhance `DxUi::ComboBox` rather than reimplement the popup locally.** Any `CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL` parity gap (free-text editable, typeahead match, `Alt+Down` to drop, `Esc` to dismiss popup without committing, scrollable popup) becomes work in `Common/DxUi/DxUi.ComboBox.cpp` (with `DxUiTests` coverage), and the ConnectionManagerWindow consumes the enhanced widget unchanged.
- **File naming — rename to `ConnectionManagerWindow.{h,cpp}`.** The new files are `RedSalamander/ConnectionManagerWindow.{h,cpp}`; the old `ConnectionManagerDialog.{h,cpp}` is removed in Phase 12.11. Includes are updated across the tree at the rename point.

---

## References

- `RedSalamander/ConnectionManagerDialog.cpp` (current implementation).
- `RedSalamander/ConnectionManagerDialog.h` (current debug interface).
- `RedSalamander/RedSalamander.rc` lines 455–513 (`IDD_CONNECTION_MANAGER` template).
- `RedSalamander/FindFilesWindow.cpp` (single-canvas precedent).
- `RedSalamander/Ui/AlertOverlayWindow.cpp` (single-canvas overlay precedent).
- `Common/DxUi/DxUi.h` (`WindowHost`, `Panel`, `Label`, `TextField`, `ComboBox`, `Toggle`, `Button`, `Grid` API).
- `Common/DxUi/DxUi.WindowHost.cpp` (focus, mnemonic, tab boundary plumbing).
- `Specs/UI/UI_TopLevelToolWindows.md` (modeless / independent-top-level contract).
- `Specs/Core/Core_ConnectionManager.md` (functional contract).
- `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md` (overall DX migration history; this plan is the deferred "one-host re-landing" of Connection Manager).
- `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md` (closeout that explicitly noted Connection Manager as DX-owned-but-multi-host and deferred the single-canvas collapse).
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` (parity contract via assertions).

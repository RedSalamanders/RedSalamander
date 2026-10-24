# UI Connection Manager Battle Test And Dialog Retirement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the single-canvas Connection Manager window reliable enough to ship as the only implementation, then remove `ConnectionManagerDialog.h` and `ConnectionManagerDialog.cpp` from source and project files.

**Architecture:** `RedSalamander/ConnectionManagerWindow.{h,cpp}` becomes the only public and internal Connection Manager surface. The window owns modeless navigation notification, profile validation, settings hot-reload participation, localized strings, debug/test hooks, and UIA/focus/list behavior; callers include the window header directly. The implementation must be backed by selftests and archived evidence before this WIP plan can move to `Specs/Plans/Done/`.

**Tech Stack:** Windows C++23, MSVC/MSBuild, WIL RAII, Win32 messages, DxUi `WindowHost`, SettingsHotReload, RedSalamander command selftests, resource string localization.

---

## Top Checklist

- [x] Confirm the worktree scope and keep unrelated dirty files out of every commit.
- [x] Add red tests for modeless Connect navigation, duplicate-name validation, hot reload, localization-backed captions, UIA/list/focus behavior, and dialog-file retirement. *(modeless Connect red/green complete; profile-name validation red/green complete; hot reload red/green complete; localization red/green complete; UIA/list/focus red/green complete; dialog retirement red/green complete)*
- [x] Fix modeless `Connect` so `ShowConnectionManagerWindow(...)` posts `WndMsg::kConnectionManagerConnect` with the selected profile name and target pane.
- [x] Add editor validation that trims profile names, rejects blank/reserved/duplicate names case-insensitively, and prevents saving invalid profiles.
- [x] Integrate `SettingsHotReload` with clean reload, dirty conflict prompt, stale-save prompt, registration, unregistration, and teardown-safe message handling.
- [x] Replace hardcoded Connection Manager user-facing strings with `.rc` resources, including secret Show/Hide captions, default new profile name, file filters, and validation messages.
- [x] Make the full `cmd_connection_manager_window_` suite green, including UIA visibility, list virtualization bounds, theme cycle, tab traversal, Enter routing, access keys, pointer toggles, and open/close stability. *(2026-04-29 local run after Task 7: 25 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111643/`; perf metrics included in `perf/perf_metrics.jsonl`.)*
- [x] Move the public API and debug declarations from `ConnectionManagerDialog.h` into `ConnectionManagerWindow.h`.
- [x] Delete `ConnectionManagerDialog.cpp` and `ConnectionManagerDialog.h`, remove them from `RedSalamander.vcxproj` and `.filters`, and update every include.
- [x] Update authoritative specs and testing docs so durable behavior is not stranded in this WIP plan.
- [x] Archive green Debug, Release, and focused command selftest evidence under `Specs/TestRuns/`. *(Final archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_115235_connection_manager_battle_test/`.)*
- [x] Move this plan to `Specs/Plans/Done/` only after all gates below pass without red or skipped Connection Manager cases.

## Definition Of Rock Solid

The window is considered rock solid when all of these are true:

- Connect from the normal modeless application entry point opens the selected connection in the requested pane, or leaves the window open with a visible/logged failure if the owner cannot be notified.
- A saved connection name uniquely identifies exactly one profile after trimming and case folding; selecting a row can never navigate with another row's credential set.
- External settings edits cannot be silently overwritten. Clean windows reload automatically; dirty windows prompt; stale saves prompt before overwriting disk state.
- Every user-visible Connection Manager string comes from `RedSalamander.rc`.
- `ConnectionManagerDialog.h` and `ConnectionManagerDialog.cpp` no longer exist in the project source list, and source/project `git grep` has no live dependency on them.
- The full `cmd_connection_manager_window_` selftest family passes in one archived run.
- No known red acceptance gate remains in `Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md` without a newer green archive that covers the same case family.

## Current Evidence To Respect

- The review found a P1 modeless Connect gap in `RedSalamander/ConnectionManagerWindow.cpp`: `OnConnectClicked` saves and closes but does not post `WndMsg::kConnectionManagerConnect` when `_modalResult` is null.
- The review found a P1 profile identity gap: the editor writes `_editName->GetText()` directly into `ConnectionProfile::name`, while `HostServices.cpp` resolves the first case-insensitive name match.
- The review found a P2 settings contract gap: the single-canvas window does not register with `SettingsHotReload` and does not handle `WndMsg::kSettingsReloadedFromDisk`.
- The committed archive `Specs/TestRuns/4cb089111a23/Commands/2026-04-28_150348/selftest_run_results.json` recorded 7 passed and 7 failed for `cmd_connection_manager_window_`.
- A later run only covered `cmd_connection_manager_window_close_persists_new_profile`; it does not supersede the red family gate.
- A fresh local `Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_connection_manager_window_'` attempt exited before test execution with `0xC0000135`; implementation closeout must include a successful run on a machine where the built executable can start.
- A 2026-04-29 local `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2` run executed successfully and reported 16 passed / 7 failed. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094231/`. Remaining failing cases: live UIA provider exposure, long-run list virtualization bounds, theme-cycle form/selection readiness, tab traversal setup, Enter/default routing setup, access-key focus targeting, and pointer toggle state routing.
- A later 2026-04-29 local `cmd_connection_manager_window_` run after localization reported 17 passed / 7 failed. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_095415/`. Remaining failing cases are unchanged: live UIA provider exposure, long-run list virtualization bounds, theme-cycle form/selection readiness, tab traversal setup, Enter/default routing setup, access-key focus targeting, and pointer toggle state routing.
- A 2026-04-29 local Task 6 run after UIA/list/focus/pointer hardening reported 24 passed / 0 failed / 0 skipped for `cmd_connection_manager_window_`. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103331/`; perf metrics were archived at `perf/perf_metrics.jsonl`.
- A 2026-04-29 local Task 7 run after retiring the dialog files reported 25 passed / 0 failed / 0 skipped for `cmd_connection_manager_window_`. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111643/`; perf metrics were archived at `perf/perf_metrics.jsonl`.

## Files And Responsibilities

- Modify `RedSalamander/ConnectionManagerWindow.h`: public API declarations, debug snapshot declarations, and single-canvas namespace declarations after retiring the dialog header.
- Modify `RedSalamander/ConnectionManagerWindow.cpp`: modeless connect notification, validation helpers, hot reload, localization use, debug hooks, UIA/list/focus repairs, and public wrapper definitions after deleting the shim.
- Delete `RedSalamander/ConnectionManagerDialog.h`: public declarations move to `ConnectionManagerWindow.h`.
- Delete `RedSalamander/ConnectionManagerDialog.cpp`: forwarding definitions move to `ConnectionManagerWindow.cpp`.
- Modify `RedSalamander/RedSalamander.cpp`, `RedSalamander/FolderWindow.FileSystem.cpp`, `RedSalamander/HostServices.cpp`, and every other live source include found by `git grep -n "ConnectionManagerDialog" -- RedSalamander`: include `ConnectionManagerWindow.h`.
- Modify `RedSalamander/RedSalamander.vcxproj` and `RedSalamander/RedSalamander.vcxproj.filters`: remove dialog header/source entries.
- Modify `RedSalamander/resource.h` and `RedSalamander/RedSalamander.rc`: add missing Connection Manager string resources and use existing secret Show/Hide IDs.
- Modify `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`: add/repair modeless connect, validation, reload, localization, UIA, list, focus, and pointer tests.
- Modify `RedSalamander/SelfTest/Commands/Commands.SelfTest.h`: register every new command case that is not auto-registered by the existing connection selftest dispatch table.
- Modify `Specs/Core/Core_ConnectionManager.md`: persist the single-canvas contract, validation rules, navigation message contract, and hot reload behavior.
- Modify `Specs/UI/UI_TopLevelToolWindows.md`: document modeless top-level behavior and Close/Connect behavior for this surface.
- Modify `Specs/UI/UI_VisibleComctlAudit.md`: record that no legacy dialog template or visible comctl surface remains for Connection Manager.
- Modify `Specs/Testing/Testing_TestCoverage.md`: list the Connection Manager command selftest gates and expected archive path.
- Add archived evidence under `Specs/TestRuns/<run-id>/Commands/<timestamp>_connection_manager_battle_test/`.

## Task 0: Protect Workspace And Establish Baseline

**Files:**
- Read: `RedSalamander/ConnectionManagerWindow.cpp`
- Read: `RedSalamander/ConnectionManagerWindow.h`
- Read: `RedSalamander/ConnectionManagerDialog.h`
- Read: `RedSalamander/ConnectionManagerDialog.cpp`
- Read: `Specs/TestRuns/4cb089111a23/Commands/2026-04-28_150348/selftest_run_results.json`

- [x] **Step 0.1: Capture the dirty-worktree boundary**

Run:

```powershell
git status --short
```

Expected: the output can include unrelated file-operation/vcpkg changes. Do not stage or modify unrelated files unless a later task explicitly names them.

- [x] **Step 0.2: Record current live dialog references**

Run:

```powershell
git grep -n "ConnectionManagerDialog" -- RedSalamander
```

Expected: hits in `ConnectionManagerDialog.{h,cpp}`, `ConnectionManagerWindow.{h,cpp}`, caller includes, and project files before removal.

- [x] **Step 0.3: Confirm the known red family gate**

Run:

```powershell
Get-Content .\Specs\TestRuns\4cb089111a23\Commands\2026-04-28_150348\selftest_run_results.json
```

Expected: the historical archive shows `case_filter` as `cmd_connection_manager_window_`, with 7 passed and 7 failed. Keep this evidence in the final closeout until a newer full-family green archive exists.

- [x] **Step 0.4: Create an implementation branch before code changes**

Run:

```powershell
git switch -c codex/connection-manager-battle-test
```

Expected: branch switches successfully. If the branch already exists, use `git switch codex/connection-manager-battle-test` and verify `git status --short` still shows only expected pre-existing changes plus this plan.

## Task 1: Add Modeless Connect Navigation Tests

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.h`
- Modify: `RedSalamander/RedSalamander.cpp`

- [x] **Step 1.1: Add a helper that opens the modeless window for a specific pane**

Insert near the existing Connection Manager test helpers in `Commands.SelfTest.Connections.cpp`:

```cpp
[[nodiscard]] bool OpenConnectionManagerForPane(HWND mainWindow, FolderWindow::Pane pane) noexcept
{
    if (mainWindow == nullptr || IsWindow(mainWindow) == FALSE)
    {
        return false;
    }

    g_folderWindow.SetActivePane(pane);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    return true;
}
```

- [x] **Step 1.2: Add a test-only last-navigation probe**

In `RedSalamander.cpp`, inside `#ifdef ENABLE_TESTS`, add:

```cpp
namespace
{
#ifdef ENABLE_TESTS
std::mutex g_debugConnectionManagerConnectMutex;
bool g_debugConnectionManagerConnectSeen = false;
uint8_t g_debugConnectionManagerConnectPane = 0u;
std::wstring g_debugConnectionManagerConnectName;
#endif
}

#ifdef ENABLE_TESTS
void DebugResetConnectionManagerConnectNavigation() noexcept
{
    const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
    g_debugConnectionManagerConnectSeen = false;
    g_debugConnectionManagerConnectPane = 0u;
    g_debugConnectionManagerConnectName.clear();
}

[[nodiscard]] bool DebugGetConnectionManagerConnectNavigation(uint8_t& outPane, std::wstring& outName) noexcept
{
    const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
    if (! g_debugConnectionManagerConnectSeen)
    {
        return false;
    }

    outPane = g_debugConnectionManagerConnectPane;
    outName = g_debugConnectionManagerConnectName;
    return true;
}
#endif
```

At the top of `OnMainWindowConnectionManagerConnect(...)`, after the payload is validated, add:

```cpp
#ifdef ENABLE_TESTS
{
    const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
    g_debugConnectionManagerConnectSeen = true;
    g_debugConnectionManagerConnectPane = static_cast<uint8_t>(wParam == 1u ? 1u : 0u);
    g_debugConnectionManagerConnectName = *name;
}
#endif
```

Declare the two probe functions in `Commands.SelfTest.h` next to the other command selftest debug declarations.

- [x] **Step 1.3: Add a helper that waits for the probe**

Add in `Commands.SelfTest.Connections.cpp`:

```cpp
[[nodiscard]] bool WaitForConnectionManagerConnectNavigation(uint8_t expectedPane,
                                                             std::wstring_view expectedName,
                                                             std::chrono::milliseconds timeout) noexcept
{
    return WaitUntil(
        [&]() noexcept {
            uint8_t actualPane = 0u;
            std::wstring actualName;
            return DebugGetConnectionManagerConnectNavigation(actualPane, actualName) &&
                   actualPane == expectedPane &&
                   actualName == expectedName;
        },
        timeout);
}
```

- [x] **Step 1.4: Add a red selftest for left-pane modeless Connect**

Add a case named `cmd_connection_manager_window_modeless_connect_posts_left_navigation`:

```cpp
[[nodiscard]] bool TestConnectionManagerWindowModelessConnectPostsLeftNavigation(HWND mainWindow, CaseState& state) noexcept
{
    SelfTest::AppendSelfTestTrace(L"ConnectionManager modeless-connect-left: begin");

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitUntil([] noexcept { return GetConnectionManagerDialogHandle() == nullptr; }, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager did not close before modeless-connect-left.");
    }

    DebugResetConnectionManagerConnectNavigation();
    state.Require(OpenConnectionManagerForPane(mainWindow, FolderWindow::Pane::Left), L"Failed to open Connection Manager for left pane.");
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr, L"Connection Manager did not open for left pane.");

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitUntil([&] noexcept {
                      return DebugGetConnectionManagerDialogSnapshot(snapshot) && snapshot.listRowCount > 0u && snapshot.selectedListIndex >= 0 &&
                             ! snapshot.currentNameText.empty();
                  }, SelfTest::Scale(5000ms)),
                  L"Connection Manager did not reach a selectable profile before left-pane Connect.");

    const std::wstring expectedName = snapshot.currentNameText;
    state.Require(DebugRouteConnectionManagerCommandKey(VK_RETURN), L"Enter did not route to the Connection Manager default Connect button.");
    state.Require(WaitUntil([] noexcept { return GetConnectionManagerDialogHandle() == nullptr; }, SelfTest::Scale(5000ms)),
                  L"Connection Manager stayed open after left-pane Connect.");

    state.Require(WaitForConnectionManagerConnectNavigation(0u, expectedName, SelfTest::Scale(5000ms)),
                  std::format(L"Left-pane Connect did not navigate to '{}'.", expectedName));

    SelfTest::AppendSelfTestTrace(L"ConnectionManager modeless-connect-left: complete");
    return true;
}
```

- [x] **Step 1.5: Add a red selftest for right-pane modeless Connect**

Add a case named `cmd_connection_manager_window_modeless_connect_posts_right_navigation` using the same body as Step 1.4 with `FolderWindow::Pane::Right`, expected pane `1u`, and command text `right-pane`.

- [x] **Step 1.6: Run the new tests and confirm the existing bug**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_modeless_connect_posts_" -TimeoutMultiplier 2
```

Expected before implementation: both cases fail because modeless `OnConnectClicked` closes without posting `WndMsg::kConnectionManagerConnect`.

## Task 2: Implement Modeless Connect Notification

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Test: `RedSalamander/RedSalamander.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.h`

- [x] **Step 2.1: Add a private notify helper**

In `ConnectionManagerWindow.cpp`, add a method to the single-canvas window class:

```cpp
[[nodiscard]] bool NotifyOwnerToConnectSelectedProfile() noexcept
{
    if (_notifyOwner == nullptr || IsWindow(_notifyOwner) == FALSE)
    {
        Debug::Error(L"Connection Manager Connect failed because the owner window is no longer valid.");
        return false;
    }

    auto payload = std::make_unique<std::wstring>(_selectedConnectionName);
    if (! PostMessagePayload(_notifyOwner,
                             WndMsg::kConnectionManagerConnect,
                             static_cast<WPARAM>(_targetPane),
                             std::move(payload)))
    {
        Debug::ErrorWithLastError(L"Connection Manager failed to post navigation request to owner.");
        return false;
    }

    return true;
}
```

Use the existing `PostMessagePayload(...)` and `WndMsg::kConnectionManagerConnect` include pattern already used by other UI-host cross-thread payloads. Do not use raw `new`, `payload.release()`, or `PostMessageW` with a raw pointer.

- [x] **Step 2.2: Route modeless Connect through the helper**

Change `OnConnectClicked()` to this control flow:

```cpp
void OnConnectClicked() noexcept
{
    if (! HasValidSelection())
    {
        return;
    }

    if (! TryValidateAndSaveCurrentStateForConnect())
    {
        return;
    }

    if (_modalResult)
    {
        *_modalResult = DialogResult{S_OK, _selectedConnectionName};
        RequestCloseWindow();
        return;
    }

    if (! NotifyOwnerToConnectSelectedProfile())
    {
        ShowConnectionManagerAlert(IDS_CONNECTIONS_ERR_CONNECT_NOTIFY_FAILED);
        return;
    }

    RequestCloseWindow();
}
```

If `TryValidateAndSaveCurrentStateForConnect()` and `ShowConnectionManagerAlert(...)` are introduced in Task 3, create them there and temporarily keep the existing `SaveConnectionsSettings()` call in this step. The final control flow after Task 3 must match this snippet.

- [x] **Step 2.3: Run the focused connect tests**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_modeless_connect_posts_" -TimeoutMultiplier 2
```

Expected: both new modeless connect cases pass, and the window closes only after successful owner notification.

- [x] **Step 2.4: Commit the modeless connect fix**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.cpp RedSalamander/RedSalamander.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.h
git commit -m "fix: route modeless connection manager connect"
```

Expected: commit contains only modeless connect implementation and focused tests.

## Task 3: Enforce Profile Name Validation

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Modify: `RedSalamander/resource.h`
- Modify: `RedSalamander/RedSalamander.rc`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`

- [x] **Step 3.1: Add resource strings for validation**

Add IDs near the existing `IDS_CONNECTIONS_*` block in `resource.h`:

```cpp
#define IDS_CONNECTIONS_CAPTION 1036
#define IDS_CONNECTIONS_ERR_NAME_RESERVED 1037
#define IDS_CONNECTIONS_ERR_CONNECT_NOTIFY_FAILED 1038
```

Add strings near the existing Connection Manager strings in `RedSalamander.rc`:

```rc
IDS_CONNECTIONS_CAPTION "Connections"
IDS_CONNECTIONS_ERR_NAME_RESERVED "The name '{0}' is reserved. Choose another name."
IDS_CONNECTIONS_ERR_CONNECT_NOTIFY_FAILED "The selected connection could not be opened because the main window was not available."
```

Use the existing `IDS_CONNECTIONS_ERR_NAME_REQUIRED` for blank names and `IDS_CONNECTIONS_ERR_NAME_UNIQUE` for duplicate names.

- [x] **Step 3.2: Add trim and comparison helpers**

Add helpers in the anonymous namespace in `ConnectionManagerWindow.cpp`:

```cpp
[[nodiscard]] std::wstring TrimConnectionText(std::wstring_view value)
{
    const auto first = std::ranges::find_if_not(value, iswspace);
    if (first == value.end())
    {
        return {};
    }

    const auto last = std::find_if_not(value.rbegin(), value.rend(), iswspace).base();
    return std::wstring(first, last);
}

[[nodiscard]] bool EqualsConnectionNameInsensitive(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    return _wcsicmp(std::wstring(lhs).c_str(), std::wstring(rhs).c_str()) == 0;
}

[[nodiscard]] bool IsReservedConnectionName(std::wstring_view name) noexcept
{
    return EqualsConnectionNameInsensitive(name, L"@quick") || EqualsConnectionNameInsensitive(name, L"Quick Connect");
}
```

Add `#include <ranges>` to `ConnectionManagerWindow.cpp` together with the other standard-library includes.

- [x] **Step 3.3: Add a validation result type**

Add this local type:

```cpp
enum class ConnectionProfileValidationError : uint8_t
{
    None,
    EmptyName,
    DuplicateName,
    ReservedName,
};

struct ConnectionProfileValidationResult
{
    ConnectionProfileValidationError error = ConnectionProfileValidationError::None;
    std::wstring normalizedName;
    size_t conflictingIndex = static_cast<size_t>(-1);
};
```

- [x] **Step 3.4: Validate against every other profile**

Add this helper:

```cpp
[[nodiscard]] ConnectionProfileValidationResult ValidateConnectionProfileName(
    std::span<const ConnectionProfile> profiles,
    size_t editedIndex,
    std::wstring_view proposedName) noexcept
{
    ConnectionProfileValidationResult result{};
    result.normalizedName = TrimConnectionText(proposedName);

    if (result.normalizedName.empty())
    {
        result.error = ConnectionProfileValidationError::EmptyName;
        return result;
    }

    if (IsReservedConnectionName(result.normalizedName))
    {
        result.error = ConnectionProfileValidationError::ReservedName;
        return result;
    }

    for (size_t index = 0; index < profiles.size(); ++index)
    {
        if (index == editedIndex)
        {
            continue;
        }

        if (EqualsConnectionNameInsensitive(profiles[index].name, result.normalizedName))
        {
            result.error            = ConnectionProfileValidationError::DuplicateName;
            result.conflictingIndex = index;
            return result;
        }
    }

    return result;
}
```

- [x] **Step 3.5: Gate name edits and saves through validation**

Replace direct name assignment in the editor save path with:

```cpp
const auto validation = ValidateConnectionProfileName(_connections, static_cast<size_t>(_selectedIndex), _editName->GetText());
if (validation.error != ConnectionProfileValidationError::None)
{
    ShowNameValidationError(validation);
    return false;
}

profile.name = validation.normalizedName;
```

Use this from a new method:

```cpp
[[nodiscard]] bool TryApplyEditorToSelectedProfile() noexcept;
```

Then require `TryApplyEditorToSelectedProfile()` before `SaveConnectionsSettings()`, `OnConnectClicked()`, `OnCloseClicked()`, protocol switches, row switches, and every reload action that can persist or discard the current editor row.

- [x] **Step 3.6: Keep the UI state honest after trimming**

After accepting a trimmed name, update the text field and list row:

```cpp
if (_editName && _editName->GetText() != validation.normalizedName)
{
    _editName->SetText(validation.normalizedName);
}
RefreshListItems();
```

Expected: entering `"  Prod FTP  "` persists as `"Prod FTP"` and the visible row matches the saved value.

- [x] **Step 3.7: Add validation selftests**

Add these cases in `Commands.SelfTest.Connections.cpp`:

```text
cmd_connection_manager_window_rejects_blank_profile_name
cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive
cmd_connection_manager_window_rejects_reserved_quick_profile_name
cmd_connection_manager_window_trims_profile_name_before_save
```

Each case must:

- Open the modeless window.
- Create a dedicated editable row.
- Set the name through the DxUi test hook or keyboard input.
- Press Close or Connect.
- Assert the window remains open and the settings file is unchanged for invalid names.
- Assert the settings file contains the trimmed name for the trim case.

- [x] **Step 3.8: Run validation tests**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_rejects_" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_trims_profile_name_before_save" -TimeoutMultiplier 2
```

Expected: all validation cases pass.

Passing evidence:

- `.\build.ps1 -ProjectName RedSalamander` passed after the production validation fix.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_rejects_" -TimeoutMultiplier 2` passed 3 passed / 0 failed / 0 skipped.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_trims_profile_name_before_save" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_modeless_connect_posts_" -TimeoutMultiplier 2` passed 2 passed / 0 failed / 0 skipped after validation was added to Connect.

Earlier red evidence:

- `.\build.ps1 -ProjectName RedSalamander` passed after adding the validation cases.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_rejects_" -TimeoutMultiplier 2` failed 0 passed / 3 failed because invalid blank, duplicate, and `@quick` names still close the window.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_trims_profile_name_before_save" -TimeoutMultiplier 2` failed 0 passed / 1 failed because padded names are not persisted as trimmed names yet.

- [x] **Step 3.9: Commit validation**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.cpp RedSalamander/resource.h RedSalamander/RedSalamander.rc RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
git commit -m "fix: validate connection profile names"
```

Expected: commit contains profile validation, resources, and tests only.

## Task 4: Add Settings Hot Reload Participation

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Reference: `RedSalamander/Preferences.Dialog.cpp`
- Reference: `RedSalamander/CompareDirectoriesWindow.Options.cpp`
- Reference: `RedSalamander/SettingsHotReload.h`

- [x] **Step 4.1: Add state fields**

Add fields to the Connection Manager window class:

```cpp
bool _settingsHotReloadRegistered = false;
bool _dirtySinceLastSettingsLoad  = false;
bool _staleExternalSettings       = false;
bool _loadingFromSettings         = false;
```

The Connection Manager follows the existing Preferences and Compare Directories pattern: dirty/stale booleans plus `SettingsHotReload` prompts, with no custom generation token.

- [x] **Step 4.2: Register after the HWND is valid**

In the successful `OnCreate` path, after the window is fully initialized:

```cpp
SettingsHotReload::RegisterParticipant(_hwnd);
_settingsHotReloadRegistered = true;
```

- [x] **Step 4.3: Unregister during teardown**

In `WM_NCDESTROY` or the existing teardown method:

```cpp
if (_settingsHotReloadRegistered)
{
    SettingsHotReload::UnregisterParticipant(_hwnd);
    _settingsHotReloadRegistered = false;
}
```

Keep this before destroying child UI state that hot-reload handlers might inspect.

- [x] **Step 4.4: Mark editor changes dirty**

In every field-change path, including `OnEditorFieldChanged`, protocol combo changes, toggle changes, and add/remove row commands:

```cpp
if (! _loadingFromSettings)
{
    _dirtySinceLastSettingsLoad = true;
}
```

Do not mark dirty while populating fields from settings; wrap `LoadConnectionsFromSettings()` and `PopulateEditorFromSelection()` with this scope guard:

```cpp
const bool wasLoading = _loadingFromSettings;
_loadingFromSettings = true;
const auto restoreLoading = wil::scope_exit([&] noexcept { _loadingFromSettings = wasLoading; });
```

- [x] **Step 4.5: Handle reload messages**

Add `WndMsg::kSettingsReloadedFromDisk` handling in `WindowProc`:

```cpp
case WndMsg::kSettingsReloadedFromDisk:
    self->OnSettingsReloadedFromDisk();
    return 0;
```

Add:

```cpp
void OnSettingsReloadedFromDisk() noexcept
{
    if (! _dirtySinceLastSettingsLoad)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _staleExternalSettings   = false;
        return;
    }

    SettingsHotReload::ExternalReloadChoice choice = SettingsHotReload::ExternalReloadChoice::KeepEditing;
    const HRESULT promptHr = SettingsHotReload::PromptExternalReloadConflict(
        _hwnd,
        LoadStringResource(nullptr, IDS_CONNECTIONS_CAPTION),
        choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"Connection Manager: failed to prompt for external reload conflict (hr=0x{:08X})",
                       static_cast<unsigned long>(promptHr));
        return;
    }

    if (choice == SettingsHotReload::ExternalReloadChoice::ReloadFromDisk)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _dirtySinceLastSettingsLoad = false;
        _staleExternalSettings      = false;
        return;
    }

    _staleExternalSettings = true;
}
```

- [x] **Step 4.6: Add stale-save prompt before writing**

At the start of `SaveConnectionsSettings()`:

```cpp
if (_staleExternalSettings)
{
    SettingsHotReload::StaleSaveChoice choice = SettingsHotReload::StaleSaveChoice::Cancel;
    const HRESULT promptHr = SettingsHotReload::PromptStaleSaveConflict(
        _hwnd,
        LoadStringResource(nullptr, IDS_CONNECTIONS_CAPTION),
        choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"Connection Manager: failed to prompt for stale save conflict (hr=0x{:08X})",
                       static_cast<unsigned long>(promptHr));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::Cancel)
    {
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::ReloadFromDisk)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _dirtySinceLastSettingsLoad = false;
        _staleExternalSettings      = false;
        return false;
    }
}
```

On successful save:

```cpp
_dirtySinceLastSettingsLoad = false;
_staleExternalSettings      = false;
```

- [x] **Step 4.7: Add hot reload selftests**

Add these cases:

```text
cmd_connection_manager_window_clean_external_reload_refreshes_list
cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing
cmd_connection_manager_window_stale_save_prompts_before_overwrite
```

Red evidence:

- `.\build.ps1 -ProjectName RedSalamander` passed after adding the tests and prompt-result selftest hook.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_clean_external_reload_refreshes_list" -TimeoutMultiplier 2` failed 0 passed / 1 failed because clean external reload did not refresh the selected row.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing" -TimeoutMultiplier 2` failed 0 passed / 1 failed because dirty external reload did not prompt.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_stale_save_prompts_before_overwrite" -TimeoutMultiplier 2` failed 0 passed / 1 failed because dirty external reload did not enter the stale-save prompt path.

Each case must edit the settings file through the same test settings fixture used by Preferences tests, call `SettingsHotReload::NotifyParticipants()`, and assert the visible list/editor and saved file reflect the chosen branch.

- [x] **Step 4.8: Run hot reload tests**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_clean_external_reload_refreshes_list" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_stale_save_prompts_before_overwrite" -TimeoutMultiplier 2
```

Expected: all hot reload cases pass and no settings edit is silently lost.

Green evidence (2026-04-29):

- `.\build.ps1 -ProjectName RedSalamander` passed after hot reload implementation and the active-pane modeless-launch correction.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_clean_external_reload_refreshes_list" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094113/`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094119/`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_stale_save_prompts_before_overwrite" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094127/`.
- Regression checks after this slice: modeless left/right navigation passed 2 passed / 0 failed / 0 skipped (`Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094048/`); name rejection passed 3 passed / 0 failed / 0 skipped (`Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094102/`); profile-name trim passed 1 passed / 0 failed / 0 skipped (`Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094108/`).
- The full `cmd_connection_manager_window_` family now executes locally but still fails 7 UIA/list/focus/pointer cases: 16 passed / 7 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_094231/`.

- [x] **Step 4.9: Commit hot reload**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
git commit -m "fix: add connection manager settings hot reload"
```

Expected: commit contains hot reload code and tests only.

## Task 5: Complete Localization For Connection Manager Strings

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Modify: `RedSalamander/resource.h`
- Modify: `RedSalamander/RedSalamander.rc`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`

- [x] **Step 5.1: Use existing Show/Hide resources**

Replace hardcoded secret button captions with:

```cpp
LoadStringResource(IDS_CONNECTIONS_BTN_SHOW_SECRET)
LoadStringResource(IDS_CONNECTIONS_BTN_HIDE_SECRET)
```

Apply this in initial button construction, mask toggling, and editor reset paths.

- [x] **Step 5.2: Use the existing default new profile string**

Replace hardcoded `L"New connection"` with:

```cpp
LoadStringResource(IDS_CONNECTIONS_DEFAULT_NEW_NAME)
```

Use it in both `MakeUniqueConnectionName(...)` fallback and `OnNewClicked()`.

- [x] **Step 5.3: Add and use a file filter resource**

Add:

```cpp
#define IDS_CONNECTIONS_FILE_FILTER_ALL_FILES 1039
```

Add to `RedSalamander.rc`:

```rc
IDS_CONNECTIONS_FILE_FILTER_ALL_FILES "All Files (*.*)\0*.*\0"
```

Replace the private-key/known-hosts `OPENFILENAMEW` filter assignment with:

```cpp
const auto filter = LoadStringResource(IDS_CONNECTIONS_FILE_FILTER_ALL_FILES);
ofn.lpstrFilter = filter.c_str();
```

Keep the filter string alive for the full `GetOpenFileNameW(...)` call.

- [x] **Step 5.4: Add a hardcoded-string audit test**

Add `cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels` that:

- Opens the window.
- Reads the Show/Hide button label through `DebugGetConnectionManagerCommandButtonHostAndClientRect(...)` or the toggle-specific debug hook.
- Toggles the secret mask twice.
- Asserts labels equal `FormatStringResource(IDS_CONNECTIONS_BTN_SHOW_SECRET)` and `FormatStringResource(IDS_CONNECTIONS_BTN_HIDE_SECRET)`.
- Creates a row and asserts the default name equals `FormatStringResource(IDS_CONNECTIONS_DEFAULT_NEW_NAME)` or its unique-numbered derivative.

Red evidence:

- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels" -TimeoutMultiplier 2` failed 0 passed / 1 failed because the secret visibility form-action label was not exposed through the debug label hook. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_095031/`.

- [x] **Step 5.5: Run static localization audit**

Run:

```powershell
Select-String -Path .\RedSalamander\ConnectionManagerWindow.cpp -Pattern 'L"Show"|L"Hide"|L"New connection"|L"All Files'
```

Expected: no matches for those hardcoded user-facing literals in `ConnectionManagerWindow.cpp`.

- [x] **Step 5.6: Run localization test**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels" -TimeoutMultiplier 2
```

Expected: test passes.

Green evidence (2026-04-29):

- `.\build.ps1 -ProjectName RedSalamander` passed after replacing hardcoded dynamic strings with resources.
- `Select-String -Path .\RedSalamander\ConnectionManagerWindow.cpp -Pattern 'L"Show"|L"Hide"|L"New connection"|L"All Files'` returned no matches.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_095314/`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2` reported 17 passed / 7 failed / 0 skipped, with the same remaining UIA/list/focus/pointer failures. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_095415/`.

- [x] **Step 5.7: Commit localization**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.cpp RedSalamander/resource.h RedSalamander/RedSalamander.rc RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
git commit -m "fix: localize connection manager dynamic strings"
```

Expected: commit contains only resource, code, and focused test changes.

## Task 6: Repair UIA, List, Focus, Keyboard, And Pointer Gates

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`

- [x] **Step 6.1: Fix UIA visibility for the single window host**

Inspect `DebugGetConnectionManagerListHostHandle(...)`, `DebugGetConnectionManagerNameHostHandle(...)`, and the UIA provider path used by the single `WindowHost`. The live-dx test must observe visible descendants for list rows, editable fields, command buttons, and toggles.

Expected code shape in the debug snapshot:

```cpp
out.visibleDxListHostCount = (_list != nullptr && _list->IsVisible()) ? 1u : 0u;
out.visibleListRowCount    = _list != nullptr ? _list->VisibleRowCount() : 0u;
out.visibleListCellCount   = _list != nullptr ? _list->VisibleCellCount() : 0u;
```

The provider must expose the same visible controls that the debug snapshot reports.

- [x] **Step 6.2: Fix list virtualization accounting**

The long-run list test failed with `visible rows=304 total rows=304`. Change the list or snapshot logic so `visibleListRowCount` reports realized rows in the viewport, not total model rows.

Expected invariant:

```cpp
out.visibleListRowCount > 0u;
out.visibleListRowCount < out.listRowCount;
out.visibleListRowCount <= 80u;
```

Keep the upper bound at the existing test value when it is stricter than `80u`.

- [x] **Step 6.3: Fix theme-cycle selection persistence**

Ensure `UpdateConnectionManagerWindowsTheme(...)` and the single-canvas `UpdateTheme(...)` path preserve:

```cpp
selectedListIndex >= 0
!selectedListRowName.empty()
!currentNameText.empty()
selectedListRowFillArgb != selectedListRowTextArgb
```

Do not rebuild the list in a way that clears the editor before restoring selection.

- [x] **Step 6.4: Fix tab traversal and mnemonics**

Make `DebugRouteConnectionManagerTab(...)` and `DebugRouteConnectionManagerMnemonic(...)` route through the same focus manager used by live keyboard messages. The expected focus sequence must include:

```text
List -> New -> Rename -> Remove -> Name -> Protocol -> Host/Region -> Port -> user/auth fields -> S3/SFTP fields -> Connect -> Close -> Cancel
```

Hidden controls must be skipped. Disabled OAuth/password controls must be skipped.

- [x] **Step 6.5: Fix Enter and Escape command routing**

Ensure:

```cpp
VK_RETURN -> OnConnectClicked()
VK_ESCAPE -> RequestCloseWindow()
```

When a multiline edit is not active, Enter must use the default Connect action. If validation fails, Enter keeps the window open and focuses the invalid field.

- [x] **Step 6.6: Fix pointer toggle state and pre-ack path**

Keep `DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt()` deterministic for the S3 insecure TLS prompt, and make the pointer-toggle debug state report stable resource-backed labels plus the live checked state. The pointer-toggle test must verify that the visible HTTPS/TLS toggle changes state after a real canvas click without depending on transient On/Off label text.

- [x] **Step 6.7: Run each previously red case individually**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_live_dx_interaction" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_long_run_list_scrolling_stays_bounded" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_tab_traversal_live_dx_interaction" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_enter_from_dx_input_routes_default_connect" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_access_keys_focus_expected_controls" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle" -TimeoutMultiplier 2
```

Expected: every command passes individually.

Green evidence (2026-04-29):

- `.\build.ps1 -ProjectName RedSalamander` passed. Log: `.build/logs/msbuild-20260429_103007_968.log`.
- `cmd_connection_manager_window_live_dx_interaction` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_101647/`.
- `cmd_connection_manager_window_long_run_list_scrolling_stays_bounded` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_102121/`.
- `cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_102354/`.
- `cmd_connection_manager_window_tab_traversal_live_dx_interaction` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103209/`.
- `cmd_connection_manager_window_enter_from_dx_input_routes_default_connect` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103228/`.
- `cmd_connection_manager_window_access_keys_focus_expected_controls` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103232/`.
- `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103222/`.

- [x] **Step 6.8: Run the full connection manager family**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2
```

Expected: all Connection Manager window cases pass in a single run, with zero failures and zero skipped cases unless a skip is explicitly documented by test infrastructure outside Connection Manager.

Green evidence (2026-04-29):

- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2` passed 24 passed / 0 failed / 0 skipped in one run. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_103331/`.
- The full-family archive includes `perf/perf_metrics.jsonl` for the run.

- [ ] **Step 6.9: Commit UI hardening**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
git commit -m "fix: harden connection manager dxui behavior"
```

Expected: commit contains UIA/list/focus/pointer repairs and tests only.

## Task 7: Retire ConnectionManagerDialog Files

**Files:**
- Modify: `RedSalamander/ConnectionManagerWindow.h`
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Delete: `RedSalamander/ConnectionManagerDialog.h`
- Delete: `RedSalamander/ConnectionManagerDialog.cpp`
- Modify: `RedSalamander/RedSalamander.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.cpp`
- Modify: `RedSalamander/HostServices.cpp`
- Modify all additional files reported by `git grep -n "ConnectionManagerDialog" -- RedSalamander`
- Modify: `RedSalamander/RedSalamander.vcxproj`
- Modify: `RedSalamander/RedSalamander.vcxproj.filters`

- [x] **Step 7.1: Move public declarations into ConnectionManagerWindow.h**

Replace the migration-only header content with public declarations from `ConnectionManagerDialog.h`, followed by the internal namespace declarations. Keep public function names stable:

```cpp
#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

HRESULT ShowConnectionManagerDialog(HWND owner,
                                    std::wstring_view appId,
                                    Common::Settings::Settings& settings,
                                    const AppTheme& theme,
                                    std::wstring_view filterPluginId,
                                    std::wstring& selectedConnectionNameOut) noexcept;

[[nodiscard]] bool ShowConnectionManagerWindow(HWND owner,
                                               std::wstring_view appId,
                                               Common::Settings::Settings& settings,
                                               const AppTheme& theme,
                                               std::wstring_view filterPluginId,
                                               uint8_t targetPane) noexcept;

[[nodiscard]] HWND GetConnectionManagerDialogHandle() noexcept;
void UpdateConnectionManagerWindowsTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
enum class ConnectionManagerDebugFocusKind : uint8_t
{
    None,
    List,
    CommandButton,
    Edit,
    Combo,
    Toggle,
    FormActionButton,
};

struct ConnectionManagerDebugSnapshot
{
    bool usesDxUiCommandButtons               = false;
    bool usesDxUiSectionHeaders               = false;
    bool usesDxUiFormLabels                   = false;
    bool usesDxUiFormInputs                   = false;
    bool usesDxUiFormActionButtons            = false;
    bool usesDxUiList                         = false;
    size_t legacyOwnerDrawCommandButtonCount  = 0u;
    size_t legacyOwnerDrawFormInputCount      = 0u;
    size_t legacyOwnerDrawFormActionButtonCount = 0u;
    size_t visibleLegacyCommandButtonCount    = 0u;
    size_t visibleLegacySectionHeaderCount    = 0u;
    size_t visibleLegacyFormLabelCount        = 0u;
    size_t visibleLegacyFormInputCount        = 0u;
    size_t visibleLegacyFormActionButtonCount = 0u;
    size_t visibleLegacyListCount             = 0u;
    size_t visibleDxSectionHeaderHostCount    = 0u;
    size_t visibleDxFormInputHostCount        = 0u;
    size_t visibleDxFormActionButtonHostCount = 0u;
    size_t visibleDxListHostCount             = 0u;
    size_t listRowCount                       = 0u;
    size_t visibleListRowCount                = 0u;
    size_t visibleListColumnCount             = 0u;
    size_t visibleListCellCount               = 0u;
    bool listHasVerticalScrollbar             = false;
    bool themeDark                            = false;
    bool themeHighContrast                    = false;
    bool themeRainbow                         = false;
    uint64_t dxListRenderCount                = 0u;
    uint64_t dxListResizeCount                = 0u;
    uint64_t dxListResizeFailureCount         = 0u;
    int selectedListIndex                     = -1;
    std::wstring selectedListRowName;
    uint32_t selectedListRowFillArgb          = 0u;
    uint32_t selectedListRowTextArgb          = 0u;
    bool selectedListRowUsesRainbow           = false;
    ConnectionManagerDebugFocusKind focusKind = ConnectionManagerDebugFocusKind::None;
    std::wstring focusLabel;
    std::wstring currentNameText;
    std::wstring currentPluginId;
    bool nameHostPresent             = false;
    bool nameHostVisible             = false;
    bool nameHostEnabled             = false;
    bool nameLegacyVisible           = false;
    bool nameTextFieldPresent        = false;
    bool nameTextFieldVisible        = false;
    bool nameTextFieldEnabled        = false;
    bool nameHostFocusControlMatches = false;
    bool nameHostOwnsFocus           = false;
};

[[nodiscard]] bool DebugGetConnectionManagerDialogSnapshot(ConnectionManagerDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugClickConnectionManagerListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugScrollConnectionManagerListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerFirstInput() noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerList() noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerMnemonic(wchar_t mnemonic) noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerCommandKey(WPARAM virtualKey) noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerTab(bool reverse) noexcept;
[[nodiscard]] bool DebugSetConnectionManagerProtocolPluginId(std::wstring_view pluginId) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerListHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerNameHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt() noexcept;
[[nodiscard]] bool DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerSavePasswordToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerCommandButtonHostAndClientRect(UINT commandId, HWND& outHost, RECT& outRect, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerS3UseHttpsToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
#endif
```

Keep the `RedSalamander::ConnectionManager::SingleCanvas` declarations below these public declarations, or remove the namespace declaration if the cpp no longer needs it externally.

- [x] **Step 7.2: Move shim functions into ConnectionManagerWindow.cpp**

At the bottom of `ConnectionManagerWindow.cpp`, keep or add the public wrappers that used to live in `ConnectionManagerDialog.cpp`:

```cpp
HRESULT ShowConnectionManagerDialog(HWND owner,
                                    std::wstring_view appId,
                                    Common::Settings::Settings& settings,
                                    const AppTheme& theme,
                                    std::wstring_view filterPluginId,
                                    std::wstring& selectedConnectionNameOut) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::ShowDialog(owner, appId, settings, theme, filterPluginId, selectedConnectionNameOut);
}

bool ShowConnectionManagerWindow(HWND owner,
                                 std::wstring_view appId,
                                 Common::Settings::Settings& settings,
                                 const AppTheme& theme,
                                 std::wstring_view filterPluginId,
                                 uint8_t targetPane) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::ShowWindow(owner, appId, settings, theme, filterPluginId, targetPane);
}

HWND GetConnectionManagerDialogHandle() noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::GetWindowHandle();
}

void UpdateConnectionManagerWindowsTheme(const AppTheme& theme) noexcept
{
    RedSalamander::ConnectionManager::SingleCanvas::UpdateTheme(theme);
}
```

Move every `Debug*` wrapper from `ConnectionManagerDialog.cpp` into `ConnectionManagerWindow.cpp` under `#ifdef ENABLE_TESTS`.

- [x] **Step 7.3: Update includes**

Replace:

```cpp
#include "ConnectionManagerDialog.h"
```

with:

```cpp
#include "ConnectionManagerWindow.h"
```

in every source file reported by:

```powershell
git grep -n '#include "ConnectionManagerDialog.h"' -- RedSalamander
```

Expected known callers include `RedSalamander.cpp`, `FolderWindow.FileSystem.cpp`, and `HostServices.cpp`.

- [x] **Step 7.4: Remove migration comments that reference the deleted shim as live code**

Update comments in `ConnectionManagerWindow.h` and `ConnectionManagerWindow.cpp` so they describe the current state:

```cpp
// The Connection Manager runs as a single-canvas DxUi top-level window.
```

Expected: no source comment claims the public API remains in `ConnectionManagerDialog.h` or that `ConnectionManagerDialog.cpp` still forwards entry points.

- [x] **Step 7.5: Remove project entries**

Delete these entries from `RedSalamander.vcxproj`:

```xml
<ClInclude Include="ConnectionManagerDialog.h" />
<ClCompile Include="ConnectionManagerDialog.cpp" />
```

Delete the same include/compile entries from `RedSalamander.vcxproj.filters`.

- [x] **Step 7.6: Delete the files**

Run:

```powershell
git rm RedSalamander\ConnectionManagerDialog.h RedSalamander\ConnectionManagerDialog.cpp
```

Expected: both files are staged as deleted.

- [x] **Step 7.7: Verify no live dialog-file references remain**

Run:

```powershell
git grep -n "ConnectionManagerDialog" -- RedSalamander
```

Expected: no hits in live source, headers, project files, or filters. If function names such as `ShowConnectionManagerDialog` remain for compatibility, this command will still find them; in that case run the narrower file-name checks below and confirm they have no hits:

```powershell
git grep -n "ConnectionManagerDialog\.h\|ConnectionManagerDialog\.cpp" -- RedSalamander
git grep -n '#include "ConnectionManagerDialog.h"' -- RedSalamander
```

Expected: both narrower checks return no hits.

- [x] **Step 7.8: Build after file retirement**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Expected: build succeeds with no missing include, duplicate symbol, or unresolved external errors.

- [x] **Step 7.9: Commit dialog retirement**

Run:

```powershell
git add RedSalamander/ConnectionManagerWindow.h RedSalamander/ConnectionManagerWindow.cpp RedSalamander/RedSalamander.cpp RedSalamander/FolderWindow.FileSystem.cpp RedSalamander/HostServices.cpp RedSalamander/RedSalamander.vcxproj RedSalamander/RedSalamander.vcxproj.filters
git add -u RedSalamander/ConnectionManagerDialog.h RedSalamander/ConnectionManagerDialog.cpp
git commit -m "refactor: retire connection manager dialog shim"
```

Expected: commit removes both dialog files from disk and project files.

Red evidence:

- `.\build.ps1 -ProjectName RedSalamander` passed after adding the retirement guard test. Log: `.build/logs/msbuild-20260429_110706_806.log`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_retired_dialog_files_absent" -TimeoutMultiplier 2` failed 0 passed / 1 failed / 0 skipped because the retired header still existed. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_110858/`.

Green evidence (2026-04-29):

- `Get-ChildItem -Path RedSalamander -Recurse -File | Select-String -Pattern 'ConnectionManagerDialog\.h|ConnectionManagerDialog\.cpp'` returned no matches after file retirement.
- `.\build.ps1 -ProjectName RedSalamander` passed after retiring the files. Log: `.build/logs/msbuild-20260429_111100_292.log`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_retired_dialog_files_absent" -TimeoutMultiplier 2` passed 1 passed / 0 failed / 0 skipped. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111238/`.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2` passed 25 passed / 0 failed / 0 skipped in one run. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111643/`.
- The full-family archive includes `perf/perf_metrics.jsonl` for the run.

## Task 8: Update Specs And Test Coverage Contracts

**Files:**
- Modify: `Specs/Core/Core_ConnectionManager.md`
- Modify: `Specs/UI/UI_TopLevelToolWindows.md`
- Modify: `Specs/UI/UI_VisibleComctlAudit.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md`

- [x] **Step 8.1: Update the Core Connection Manager spec**

Add or update a section in `Specs/Core/Core_ConnectionManager.md`:

```markdown
## Single-Canvas Connection Manager Window

- The live implementation is `RedSalamander/ConnectionManagerWindow.{h,cpp}`.
- `ShowConnectionManagerWindow(...)` is modeless and posts `WndMsg::kConnectionManagerConnect` to the owner with the selected connection name payload and the requested target pane.
- `ShowConnectionManagerDialog(...)` is a synchronous facade over the same single-canvas implementation and returns the selected connection name to host-service callers.
- Profile names are trimmed before save and must be non-empty, unique case-insensitively, and not reserved for Quick Connect.
- The settings hot-reload contract matches Preferences: clean windows reload, dirty windows prompt, and stale saves prompt before overwrite.
- The deleted `ConnectionManagerDialog.{h,cpp}` files must not be reintroduced.
```

- [x] **Step 8.2: Update top-level tool window spec**

Add:

```markdown
### Connection Manager

The Connection Manager is a top-level modeless tool window for normal application commands. Connect validates and saves the current profile, posts a typed payload to the owner, and closes only after the owner notification is queued. Close validates and saves dirty edits or leaves the window open when validation fails.
```

- [x] **Step 8.3: Update visible comctl audit spec**

Add:

```markdown
Connection Manager has no live `IDD_CONNECTION_MANAGER` dialog template path and no visible legacy comctl owner-draw controls. The single DxUi window host is responsible for UIA, keyboard traversal, pointer interaction, and theme repaint behavior.
```

- [x] **Step 8.4: Update test coverage spec**

Add:

```markdown
Connection Manager closeout requires:

- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_credential_prompt_" -TimeoutMultiplier 2`
- Debug and Release `.\build.ps1 -ProjectName RedSalamander`
- A `Specs/TestRuns/.../Commands/...` archive for the full Connection Manager window family with zero failures.
```

- [x] **Step 8.5: Annotate the old Done plan with superseding evidence**

In `Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md`, add a closeout note near the red acceptance gate:

```markdown
> Superseded by `Specs/Plans/Done/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md` once that plan is moved to Done with a full green `cmd_connection_manager_window_` archive. The 2026-04-28_150348 run is not accepted as a green gate.
```

Do this only after the green archive exists, so the note can point to the exact archive path.

- [x] **Step 8.6: Commit spec updates**

Run:

```powershell
git add Specs/Core/Core_ConnectionManager.md Specs/UI/UI_TopLevelToolWindows.md Specs/UI/UI_VisibleComctlAudit.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md
git commit -m "docs: lock connection manager window contract"
```

Expected: commit contains only docs/spec changes.

Green evidence:

- Committed as `58d8a3a8e` (`docs: lock connection manager window contract`).

## Task 9: Final Verification And Evidence Archive

**Files:**
- Add: `Specs/TestRuns/<run-id>/Commands/<timestamp>_connection_manager_battle_test/`
- Move after success: `Specs/Plans/WIP/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md` to `Specs/Plans/Done/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md`

- [x] **Step 9.1: Build Debug**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Expected: Debug build succeeds.

- [x] **Step 9.2: Build Release**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Release
```

Expected: Release build succeeds.

- [x] **Step 9.3: Run the full Connection Manager family**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2
```

Expected: all cases pass in one run, with zero failures.

- [x] **Step 9.4: Run the credential prompt neighbor family**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_credential_prompt_" -TimeoutMultiplier 2
```

Expected: all credential prompt cases pass. This guards the shared secret Show/Hide resource and password masking behavior.

- [x] **Step 9.5: Run visible-surface audits**

Run the audit scripts present in the current tree:

```powershell
.\Tools\Audit-VisibleNativeSurfaces.ps1
.\Tools\Audit-ComctlReportSurfaces.ps1
.\Tools\Get-VisibleTypographyAudit.ps1
```

Expected: no Connection Manager regressions, no visible legacy dialog controls, and no typography overlap/focus-label regression.

- [x] **Step 9.6: Run source removal checks**

Run:

```powershell
Test-Path .\RedSalamander\ConnectionManagerDialog.h
Test-Path .\RedSalamander\ConnectionManagerDialog.cpp
git grep -n "ConnectionManagerDialog\.h\|ConnectionManagerDialog\.cpp" -- RedSalamander
git grep -n '#include "ConnectionManagerDialog.h"' -- RedSalamander
Select-String -Path .\RedSalamander\RedSalamander.vcxproj,.\RedSalamander\RedSalamander.vcxproj.filters -Pattern "ConnectionManagerDialog"
```

Expected: both `Test-Path` commands print `False`; grep/search commands return no live project/source hits.

- [x] **Step 9.7: Archive the evidence**

Copy the command selftest result folder into:

```text
Specs/TestRuns/<run-id>/Commands/<timestamp>_connection_manager_battle_test/
```

The archive must include:

```text
selftest_run_results.json
stdout/stderr or transcript
build configuration notes
git commit hash
case filter: cmd_connection_manager_window_
pass/fail counts
```

- [x] **Step 9.8: Move the plan to Done**

Run:

```powershell
git mv Specs\Plans\WIP\UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md Specs\Plans\Done\UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md
```

Expected: the plan moves only after every gate above passes.

- [x] **Step 9.9: Commit evidence and plan closeout**

Run:

```powershell
git add Specs/TestRuns Specs/Plans/Done/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md
git add -u Specs/Plans/WIP/UI_ConnectionManagerBattleTestAndDialogRetirementPlan.md
git commit -m "test: archive connection manager battle test evidence"
```

Expected: final commit contains the evidence archive and the WIP-to-Done move.

Final green evidence (2026-04-29):

- Debug build passed: `.\build.ps1 -ProjectName RedSalamander`. Log: `.build/logs/msbuild-20260429_114459_402.log`.
- Release build passed: `.\build.ps1 -ProjectName RedSalamander -Configuration Release`. Log: `.build/logs/msbuild-20260429_114318_137.log`.
- Full Connection Manager family passed: `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2`, 25 passed / 0 failed / 0 skipped. Archive copied into `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_115235_connection_manager_battle_test/`.
- Credential prompt neighbor family passed: `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_credential_prompt_" -TimeoutMultiplier 2`, 8 passed / 0 failed / 0 skipped. Archive copied into `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_115235_connection_manager_battle_test/credential_prompt_2026-04-29_115039/`.
- Visible native, comctl report-surface, and typography audits passed. Outputs are archived in `audit_visible_native_surfaces.txt`, `audit_comctl_report_surfaces.txt`, and `audit_visible_typography.txt`.
- Source-removal checks passed: `ConnectionManagerDialog.h` and `.cpp` are absent, retired filename/include greps have no hits, and project files have no `ConnectionManagerDialog` entries. Output: `source_removal_checks.txt`.
- The final archive includes `closeout_notes.txt`, `selftest_run_results.json`, `commands_results.json`, and `perf/perf_metrics.jsonl`.

## Final Commit Stack

- `fix: route modeless connection manager connect`
- `fix: validate connection profile names`
- `fix: add connection manager settings hot reload`
- `fix: localize connection manager dynamic strings`
- `fix: harden connection manager dxui behavior`
- `refactor: retire connection manager dialog shim`
- `docs: lock connection manager window contract`
- `test: archive connection manager battle test evidence`

## Final Reviewer Checklist

- [x] `git status --short` shows no unrelated user changes staged with this work.
- [x] The final full-family test archive is newer than `2026-04-28_150348` and records zero Connection Manager failures.
- [x] `ConnectionManagerDialog.h` and `ConnectionManagerDialog.cpp` are absent from disk and project files.
- [x] `ConnectionManagerWindow.h` is the only public header for the Connection Manager surface.
- [x] Modeless Connect posts a typed payload instead of silently closing.
- [x] Duplicate names, blank names, and reserved names cannot be saved.
- [x] Hot reload protects dirty and stale edits.
- [x] Hardcoded `Show`, `Hide`, `New connection`, and file filter strings are absent from `ConnectionManagerWindow.cpp`.
- [x] UIA/list/focus/theme/pointer tests all pass in one `cmd_connection_manager_window_` run.
- [x] Specs describe the final behavior, not the retired migration shim.

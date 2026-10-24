# File Actions Preferences Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]` / `- [x]`) syntax for tracking.

**Goal:** Replace the split-brained Viewers/Editors Preferences model with an Actions + Associations model that answers, for a file such as `.txt`, what F3, Alt+F3, F4, Ctrl+Shift+F4, and Shift+F4 will do.

**Architecture:** Keep `Preferences > Viewers` and `Preferences > Editors` as separate user-facing pages, because the command families and columns are different, but back both pages with one shared `fileActions` settings model, one resolver, one launch/applicability layer, and one reusable DxUi page builder. This is a schema-breaking change: no v15 viewer/editor compatibility, no migration from old root `viewers`/`editors`, and no migration from `extensions.openWithViewerByExtension`.

**Tech Stack:** C++23, WIL RAII, yyjson, Direct2D/DxUi, RedSalamander command selftests, Settings Store schema v16, `Debug::Perf` metrics.

---

## Live Checklist

- [x] WIP correction plan created in `Specs/Plans/WIP` and closed into `Specs/Plans/Done`.
- [ ] Current implementation phase: [blocked] Targeted file-actions specs/docs, focused Debug coverage, archived command-dialog/file-action perf metrics, and Release build were completed, but the broader Commands suite still stops at `cmd_connection_manager_window_modeless_connect_posts_left_navigation`; do not treat the full plan as globally closed until that blocker is resolved.
- [x] Freeze the v16 settings contract around `fileActions.viewers` and `fileActions.editors`.
- [x] Write failing settings and resolver selftests before changing runtime code. Verified RED with Debug build: missing `FileActionResolver.h`, as expected.
- [x] Replace old settings structs, parser, writer, schema, and defaults with the new model. Foundation writer/parser/schema slice is green, fresh v16 defaults now route viewer extensions through `fileActions`, and old root `viewers`/`editors` plus `extensions.openWithViewerByExtension` are rejected.
- [x] Add an explainable resolver used by commands and Preferences preview. Focused Debug selftests passed: settings v16 round-trip/reject legacy at `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_102110`; resolver priority/command keys at `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_102116`.
- [x] Rewire View, Alternate View, Edit, Alternate Edit, Edit With, View With, and Edit New to the resolver. Focused Edit New prompt selftests passed after v16 fixture update.
- [x] Build the shared DxUi Actions + Associations page surface. Passed Debug build `.build/logs/msbuild-20260503_120542_285.log`.
- [x] Replace `Preferences.Viewers.cpp` with the new Viewers page.
- [x] Replace `Preferences.Editors.cpp` with the new Editors page.
- [x] Update Edit New editor selection to use explicit `editNewActionId` associations. Focused Preferences selftest passed with Edit New default selection coverage.
- [x] Rewrite app-owned Win32 command option dialogs created by the command backlog slice in DxUi; keep OS shell/common pickers for path selection. Change Attributes and Make File List now use DxUi modal windows, canonical no-template selftests pass, and `IDD_CHANGE_ATTRIBUTES`/`IDD_MAKE_FILE_LIST` resources are gone.
- [x] Add UI, command, schema, resolver, and perf regression selftests. Latest focused Debug pass is archived across `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132613` through `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132632`.
- [x] Update authoritative specs and user docs. Completed in `Specs/Core/Core_SettingsStore.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_PreferencesDialog.md`, `Docs/SettingsFile.md`, `Docs/Preferences.md`, `Docs/Viewers.md`, `Docs/UserGuide.md`, and `Docs/Plugins.md`.
- [x] Run final focused Debug and Release verification, archive perf evidence, then move this plan to `Specs/Plans/Done`. Final evidence: Debug build `.build/logs/msbuild-20260503_132719_410.log`, focused Debug cases archived from `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132613` through `2026-05-03_132632`, Release build `.build/logs/msbuild-20260503_132856_040.log`, and perf metrics in `2026-05-03_132617`, `2026-05-03_132630`, `2026-05-03_132630_001`, and `2026-05-03_132631`.

## Live Progress Detail

- [x] Settings v16 parser/writer/schema added for `fileActions.viewers` and `fileActions.editors`.
- [x] Resolver added with command-specific resolution reasons and perf metrics.
- [x] Runtime command wiring moved to resolver for View/Edit families.
- [x] Fresh defaults now express viewer extension routing through `fileActions`.
- [x] Remove old `Common::Settings::Settings::viewers` and `Settings::editors`.
- [x] Remove old `extensions.openWithViewerByExtension` storage and all runtime/test references, leaving only explicit legacy-rejection tests and parser guards.
- [x] Replace old `FileActionSetSettings` usage in user menu with a dedicated `UserMenuSettings`.
- [x] Remove legacy resolver helpers from `FileActionLauncher`.
- [x] Convert Preferences dirty/merge/debug snapshot paths to `fileActions`.
- [x] Rebuild Debug after legacy-settings removal and fix any remaining compile failures. Passed Debug build: `.build/logs/msbuild-20260503_110505_106.log`.
- [x] Run focused settings/resolver/preferences selftests after legacy-settings removal. Passed focused settings/resolver/Edit New/apply cases archived under `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_110718`, `_110723`, `_110728`, `_110737`, and `_110744`.
- [x] Replace obsolete Viewers preference page interaction tests with Actions + Associations table coverage; the focused Viewers/Editors file-actions test now validates row counts, resolver preview, and Edit New.
- [x] Replace Viewers Preferences page with DxUi Associations + Actions tabs.
- [x] Replace Editors Preferences page with DxUi Associations + Actions tabs, including `editNewActionId`.
- [x] Rewrite any app-owned Win32 dialogs created by this backlog slice in DxUi; keep only OS shell/common pickers as system pickers. Change Attributes and Make File List option prompts now use DxUi modal windows; old `DialogBoxParamW` paths are removed from the command part.
- [x] Update preference selftests to validate the new tables, resolved-action preview, dirty-state merge, and Edit New apply behavior. Focused case `cmd_preferences_dialog_viewers_editors_file_action_settings_apply` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132617`.
- [x] Update command/settings selftests to ensure v15 legacy shape is rejected and v16 defaults are stable. Focused cases passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132613` through `2026-05-03_132615_001`.
- [x] Add or refresh perf scenarios and archive `Debug::Perf` evidence under `Specs/TestRuns/`. Preferences file-actions layout and resolver metrics are archived in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132617`; command-dialog layout metrics are archived in `2026-05-03_132630` (`commands.dialog.changeAttributes_layout_us`) and `2026-05-03_132631` (`commands.dialog.makeFileList_layout_us`); Make File List generation metrics are archived in `2026-05-03_132630_001`.
- [x] Run Debug focused tests, attempt the broader Commands suite, run Release build, and inspect warnings/errors. Final status: Debug build `.build/logs/msbuild-20260503_132719_410.log` passed; focused Debug cases passed across `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_132613` through `2026-05-03_132632`; Release build `.build/logs/msbuild-20260503_132856_040.log` passed. [blocked] Full Commands suite was re-attempted and still stops at `cmd_connection_manager_window_modeless_connect_posts_left_navigation` without writing a final results artifact; the targeted contract/perf slice for this plan is green.

## Product Decision

Use split pages:

- `Preferences > Viewers` contains viewer associations and viewer actions.
- `Preferences > Editors` contains editor associations and editor actions.
- Both pages use identical concepts, shared code, shared validation, and shared resolution preview.

Do not merge them into one visible `File Actions` page in this slice. A single page would be mechanically elegant but would force users to scan five command columns at once. The split pages keep the mental model direct:

- Viewers answers: "What happens when I press F3 or Alt+F3?"
- Editors answers: "What happens when I press F4, Ctrl+Shift+F4, or Shift+F4?"

## No-Compatibility Contract

This work intentionally breaks the old persisted shape.

- Bump `schemaVersion` from `15` to `16`.
- Remove root `viewers` and `editors` sections from `Common::Settings::Settings`.
- Add root `fileActions`.
- Remove `extensions.openWithViewerByExtension` from the settings struct, parser, writer, schema, docs, and Preferences UI.
- Keep `extensions.openWithFileSystemByExtension` for archive/file-system activation.
- Change `userMenu` to a dedicated `UserMenuSettings { actions }` shape instead of keeping viewer/editor default fields in its type.
- Normal startup with a v15 settings file follows the existing unsupported-version behavior: use defaults and back up the invalid file.
- Hot reload and `cmd/app/rereadAssociations` reject v15 with the normal invalid-settings alert and keep current runtime settings.
- No migration helper copies old viewer mappings into the new model.

## Target Settings Shape

Authoritative JSON shape:

```json5
{
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        {
          "id": "viewer-text",
          "displayName": "Text Viewer",
          "enabled": true,
          "kind": "viewerPlugin",
          "pluginId": "builtin/viewer-text",
          "appliesTo": { "matches": [{ "kind": "default" }], "computerNames": [] }
        },
        {
          "id": "irfanview",
          "displayName": "IrfanView",
          "enabled": true,
          "kind": "externalProgram",
          "executablePath": "C:\\Tools\\IrfanView\\i_view64.exe",
          "arguments": "\"{FullPath}\"",
          "workingDirectory": "{Path}",
          "appliesTo": {
            "matches": [{ "kind": "extension", "value": ".png" }, { "kind": "extension", "value": ".jpg" }],
            "computerNames": ["HOME-PC"]
          }
        }
      ],
      "associations": [
        {
          "match": { "kind": "extension", "value": ".txt" },
          "viewActionId": "viewer-text",
          "alternateViewActionId": "viewer-hex"
        },
        {
          "match": { "kind": "default" },
          "viewActionId": "viewer-text",
          "alternateViewActionId": ""
        }
      ]
    },
    "editors": {
      "actions": [
        {
          "id": "vscode",
          "displayName": "VS Code",
          "enabled": true,
          "kind": "externalProgram",
          "executablePath": "C:\\Users\\me\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
          "arguments": "\"{FullPath}\"",
          "workingDirectory": "{Path}",
          "appliesTo": { "matches": [{ "kind": "default" }], "computerNames": [] }
        }
      ],
      "associations": [
        {
          "match": { "kind": "extension", "value": ".cpp" },
          "computerName": "DEV-PC",
          "editActionId": "visual-studio",
          "alternateEditActionId": "vscode",
          "editNewActionId": "visual-studio"
        },
        {
          "match": { "kind": "default" },
          "editActionId": "vscode",
          "alternateEditActionId": "",
          "editNewActionId": "vscode"
        }
      ]
    }
  },
  "userMenu": {
    "actions": [
      {
        "id": "open-terminal-here",
        "displayName": "Open Terminal Here",
        "enabled": true,
        "kind": "externalProgram",
        "executablePath": "wt.exe",
        "arguments": "-d \"{Path}\"",
        "workingDirectory": "{Path}",
        "appliesTo": { "matches": [{ "kind": "default" }], "computerNames": [] }
      }
    ]
  }
}
```

Matching rules:

- `kind: "default"` means the `*` row.
- `kind: "extension"` stores a lowercase extension with a leading dot.
- `kind: "pattern"` stores a case-insensitive filename glob such as `*.test.log`.
- Pattern and extension rows share the same specificity bucket. If multiple rows in the same priority bucket match, the first row in the persisted grid order wins.
- The UI prevents duplicate exact keys: `(family, command group, computerName, match.kind, match.value)`.

Resolver priority:

1. Computer + extension or pattern rule.
2. Global extension or pattern rule.
3. Computer default rule.
4. Global default rule.

The resolver must return a structured explanation:

```cpp
enum class FileActionResolutionReason : uint8_t
{
    ComputerSpecificMatch,
    GlobalMatch,
    ComputerDefault,
    GlobalDefault,
    NoAssociation,
    MissingAction,
    DisabledAction,
    ActionDoesNotApply,
    UnsupportedActionKind,
};
```

Example Preferences preview text:

```text
Selected test file: C:\src\main.cpp
Command: F4 Edit
Resolves to: Visual Studio
Reason: DEV-PC extension rule for .cpp
```

## Low-Fi User Experience

Viewers page:

```text
Preferences > Viewers

[ Associations ] [ Actions ]

Associations
+-----------+----------+----------------------+-----------------------+----------------------------+
| Match     | Computer | F3 View              | Alt+F3 Alternate View | Status                     |
+-----------+----------+----------------------+-----------------------+----------------------------+
| .txt      | Any      | Text Viewer          | Hex Viewer            | Enabled                    |
| .md       | Any      | Markdown Viewer      | Text Viewer           | Enabled                    |
| .png      | HOME-PC  | Image Viewer         | External: IrfanView   | IrfanView limited HOME-PC  |
| *         | Any      | Text Viewer          | (none)                | Default                    |
+-----------+----------+----------------------+-----------------------+----------------------------+

Match: [ Extension v ] [.txt                 ]
Computer: [ Any v ] [                         ]
F3 View: [ Text Viewer v ]
Alt+F3:  [ Hex Viewer  v ]

Test file: [ C:\src\main.cpp                                      ] [ Test ]
F3 View resolves to: Text Viewer
Reason: global default
```

Editors page:

```text
Preferences > Editors

[ Associations ] [ Actions ]

Associations
+-----------+----------+---------------+-----------------------------+-------------------+----------------------+
| Match     | Computer | F4 Edit       | Ctrl+Shift+F4 Alternate Edit | Shift+F4 Edit New | Status               |
+-----------+----------+---------------+-----------------------------+-------------------+----------------------+
| .txt      | Any      | Notepad++     | VS Code                     | Notepad++         | Enabled              |
| .cpp      | DEV-PC   | Visual Studio | VS Code                     | Visual Studio     | DEV-PC override      |
| .md       | Any      | VS Code       | Typora                      | VS Code           | Enabled              |
| *         | Any      | Notepad++     | (none)                      | Notepad++         | Default              |
+-----------+----------+---------------+-----------------------------+-------------------+----------------------+

Match: [ Extension v ] [.cpp                 ]
Computer: [ DEV-PC v ]
F4 Edit:               [ Visual Studio v ]
Ctrl+Shift+F4:         [ VS Code       v ]
Shift+F4 Edit New:     [ Visual Studio v ]

Test file: [ C:\src\main.cpp                                      ] [ Test ]
F4 Edit resolves to: Visual Studio
Reason: DEV-PC extension rule for .cpp
```

Actions tab, shared layout:

```text
Actions
+----------------+------------------+-------------------+----------+----------+
| Name           | Type             | Applies To        | Computer | Status   |
+----------------+------------------+-------------------+----------+----------+
| Text Viewer    | Viewer plugin    | *                 | Any      | Enabled  |
| Hex Viewer     | Viewer plugin    | .bin .dat         | Any      | Enabled  |
| IrfanView      | External program | .png .jpg         | HOME-PC  | Enabled  |
+----------------+------------------+-------------------+----------+----------+

Name: [ IrfanView                         ]
Type: [ External program v ]
Enabled: [x]
Executable: [ C:\Tools\IrfanView\i_view64.exe                 ] [...]
Arguments:  [ "{FullPath}"                                    ]
Working dir:[ {Path}                                          ]
Applies to: [ .png .jpg                                       ]
Computers:  [ HOME-PC                                         ]
```

## Dialog Ownership Rule

No app-owned RedSalamander dialog may expose a Win32 dialog template or native child-control form for this feature. The only allowed Win32 layer for an app-owned dialog is the minimal HWND/window-host plumbing needed to own a themed DxUi surface; all visible labels, fields, tables, buttons, validation, and layout must be DxUi.

OS-provided shell/common pickers are allowed and are not a gray area. `IFileOpenDialog`, an OS save-file picker, and equivalent system-owned folder/archive path pickers may still be used for Browse buttons because they are system pickers, not RedSalamander-owned application dialogs.

The command backlog slice currently has app-owned Win32 template dialogs for command options. This plan must rewrite those surfaces, including the `DialogBoxParamW(...IDD_MAKE_FILE_LIST...)` Make File List options dialog and the `DialogBoxParamW(...IDD_CHANGE_ATTRIBUTES...)` Change Attributes options dialog if those code paths still exist when implementation begins.

## File Structure

- Modify: `Common/SettingsStore.h`
  - Replace viewer/editor `FileActionSetSettings` fields with `FileActionsSettings fileActions`.
  - Add association structs and explicit editor `editNewActionId`.
  - Split `UserMenuSettings` from viewer/editor association settings.
- Modify: `Common/Common/SettingsStore.cpp`
  - Parse/write schema v16.
  - Remove `ParseFileActionSet(root, "viewers", ...)`, `ParseFileActionSet(root, "editors", ...)`, and viewer extension parsing.
  - Add parse/write helpers for `fileActions.viewers`, `fileActions.editors`, `fileActions.*.associations`, and `userMenu.actions`.
- Modify: `Specs/SettingsStore.schema.json`
  - Bump `$id`, title, and `schemaVersion` to v16.
  - Add `$defs.fileActionsSettings`, `$defs.viewerAssociationRule`, `$defs.editorAssociationRule`, `$defs.fileActionMatch`, `$defs.fileActionApplicability`, and `$defs.launchActionDefinition`.
  - Remove root `viewers`, root `editors`, and `extensions.openWithViewerByExtension`.
- Modify: `RedSalamander/SettingsSchemaParser.cpp`, `RedSalamander/SettingsSchemaExport.cpp`
  - Teach generated Preferences/schema metadata about `fileActions`.
- Create: `RedSalamander/FileActionResolver.h`
  - Public resolver types and read-only APIs used by commands, Edit New prompt, and Preferences preview.
- Create: `RedSalamander/FileActionResolver.cpp`
  - Matching, priority order, action applicability, explanation strings, and metrics.
- Modify: `RedSalamander/FileActionLauncher.h`
  - Keep macro expansion and launch-plan construction here.
  - Remove resolver declarations from this file after callers move to `FileActionResolver`.
- Modify: `RedSalamander/FileActionLauncher.cpp`
  - Keep external launch macro behavior unchanged.
  - Use the new action definition field names.
- Modify: `RedSalamander/FolderWindow.Viewers.cpp`
  - Resolve View, Alternate View, Edit, Alternate Edit, View With, and Edit With through `FileActionResolver`.
  - Use resolution details in pane alerts and List Opened Files labels.
- Modify: `RedSalamander/FolderWindow.FileSystem.cpp`
  - Update the Edit New prompt's editor combo to resolve `editNewActionId`, not primary edit.
  - Continue using the existing DxUi prompt host and validation.
- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
  - Pass `settings.fileActions.editors` into Edit New.
  - Rewrite app-owned command option dialogs created by the command backlog slice from Win32 dialog templates to DxUi.
  - Keep OS shell/common pickers for Browse actions that choose files, folders, archive paths, or output paths.
- Modify: `RedSalamander/Preferences.Viewers.h/.cpp`
  - Replace extension-viewer mapping UI and default action controls with the new Associations + Actions page.
- Modify: `RedSalamander/Preferences.Editors.h/.cpp`
  - Replace global primary/alternate combo UI with the new Associations + Actions page.
- Create: `RedSalamander/Preferences.FileActions.h`
  - Shared DxUi model/view helpers for file-action associations/actions pages.
- Create: `RedSalamander/Preferences.FileActions.cpp`
  - Shared grid models, inline forms, combo population, validation, preview resolution, and debug hooks.
- Modify: `RedSalamander/Preferences.Internal.h`, `RedSalamander/Preferences.Dialog.cpp`, `RedSalamander/Preferences.cpp`
  - Wire new shared page state, debug snapshots, dirty tracking, layout metrics, and Apply/OK copying.
- Modify: `RedSalamander/Preferences.UserMenu.h/.cpp`
  - Update to `UserMenuSettings { actions }` while keeping external-command macro behavior.
- Modify: `RedSalamander/SettingsHotReload.cpp`
  - Reload `fileActions` and `userMenu` instead of root viewer/editor action sets.
- Modify: `RedSalamander/RedSalamander.vcxproj` and `.filters`
  - Add `FileActionResolver.*` and `Preferences.FileActions.*`.
- Modify: `RedSalamander/resource.h`, `RedSalamander/res/*.rc`
  - Add localized labels for tabs, columns, validation, resolution reasons, and preview fields.
  - Remove app-owned command option dialog templates and native-control IDs that are replaced by DxUi surfaces.
- Modify tests:
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp`
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`
- Modify docs/specs:
  - `Specs/Core/Core_SettingsStore.md`
  - `Specs/UI/UI_CommandMenuKeyboard.md`
  - `Specs/UI/UI_PreferencesDialog.md`
  - `Specs/Testing/Testing_TestCoverage.md`
  - `Docs/Preferences.md`
  - `Docs/SettingsFile.md`
  - `Docs/Viewers.md`
  - `Docs/UserGuide.md`

## Task 1: Add Failing Settings and Resolver Selftests

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`

- [x] **Step 1: Rename the old action roundtrip test to the v16 contract.**

Add a failing case named:

```cpp
SelfTest::RunCase(options, suite, L"settings_store_file_actions_v16_roundtrip", [](CaseState& state) noexcept {
    return TestSettingsStoreFileActionsV16RoundTrip(state);
});
```

The test must create `Common::Settings::Settings settings{};`, populate `settings.fileActions.viewers.actions`, `settings.fileActions.viewers.associations`, `settings.fileActions.editors.actions`, and `settings.fileActions.editors.associations`, then save/load through `Common::Settings::SaveSettings` and `TryLoadSettingsNoRecovery`.

Required assertions:

```cpp
state.Require(loaded.schemaVersion == 16u, L"Settings Store should write schema v16.");
state.Require(loaded.fileActions.viewers.associations.size() == 2u, L"Viewer associations did not round-trip.");
state.Require(loaded.fileActions.editors.associations.size() == 2u, L"Editor associations did not round-trip.");
state.Require(loaded.fileActions.editors.associations[0].editNewActionId == L"visual-studio",
              L"Edit New action association did not round-trip.");
```

Expected result before implementation: compile fails because `fileActions`, association structs, and schema v16 do not exist.

- [x] **Step 2: Add an old-shape rejection guard.**

Add a failing case named:

```cpp
SelfTest::RunCase(options, suite, L"settings_store_v16_rejects_legacy_viewer_editor_shape", [](CaseState& state) noexcept {
    return TestSettingsStoreV16RejectsLegacyViewerEditorShape(state);
});
```

Test input JSON must include `"schemaVersion": 16`, root `"viewers"`, root `"editors"`, and `"extensions": { "openWithViewerByExtension": { ".txt": "builtin/viewer-text" } }`.

Expected assertions:

```cpp
state.Require(FAILED(hr), L"v16 settings should reject root viewers/editors and legacy viewer extension mappings.");
```

Expected result before implementation: the test either compiles against old fields or incorrectly accepts legacy data.

- [x] **Step 3: Add resolver priority and reason coverage.**

Add a failing case named:

```cpp
SelfTest::RunCase(options, suite, L"file_action_resolution_v16_explains_priority", [](CaseState& state) noexcept {
    return TestFileActionResolutionV16ExplainsPriority(state);
});
```

Required matrix:

| File | Computer | Command | Expected action | Expected reason |
|---|---|---|---|---|
| `C:\src\main.cpp` | `DEV-PC` | Edit | `visual-studio` | computer + extension |
| `C:\src\readme.md` | `OTHER` | Edit | `vscode` | global extension |
| `C:\src\image.bin` | `DEV-PC` | Edit | `dev-default` | computer default |
| `C:\src\notes.unknown` | `OTHER` | Edit | `global-default` | global default |
| `C:\src\new.cpp` | `DEV-PC` | Edit New | `visual-studio` | computer + extension |
| `C:\src\old.bak` | `OTHER` | Edit | no action | disabled action |

Expected result before implementation: compile fails because `FileActionResolver` does not exist.

- [x] **Step 4: Add command-key resolution coverage.**

Add a failing case named:

```cpp
SelfTest::RunCase(options, suite, L"file_action_resolution_v16_command_keys_are_distinct", [](CaseState& state) noexcept {
    return TestFileActionResolutionV16CommandKeysAreDistinct(state);
});
```

Required assertions:

```cpp
state.Require(ResolveForCommand(settings, FileActionCommand::View, L"C:\\x\\note.txt", L"PC").actionId == L"text-viewer",
              L"F3 should use viewActionId.");
state.Require(ResolveForCommand(settings, FileActionCommand::AlternateView, L"C:\\x\\note.txt", L"PC").actionId == L"hex-viewer",
              L"Alt+F3 should use alternateViewActionId.");
state.Require(ResolveForCommand(settings, FileActionCommand::Edit, L"C:\\x\\note.txt", L"PC").actionId == L"notepad-plus-plus",
              L"F4 should use editActionId.");
state.Require(ResolveForCommand(settings, FileActionCommand::AlternateEdit, L"C:\\x\\note.txt", L"PC").actionId == L"vscode",
              L"Ctrl+Shift+F4 should use alternateEditActionId.");
state.Require(ResolveForCommand(settings, FileActionCommand::EditNew, L"C:\\x\\note.txt", L"PC").actionId == L"notepad-plus-plus",
              L"Shift+F4 should use editNewActionId.");
```

- [x] **Step 5: Run the focused RED build.**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Expected: build fails on missing v16 model/resolver symbols. Save the log path in this plan's checklist when implementation starts.

## Task 2: Replace Settings Model and Schema

**Files:**
- Modify: `Common/SettingsStore.h`
- Modify: `Common/Common/SettingsStore.cpp`
- Modify: `Specs/SettingsStore.schema.json`
- Modify: `RedSalamander/SettingsSchemaParser.cpp`
- Modify: `RedSalamander/SettingsSchemaExport.cpp`

- [x] **Step 1: Define v16 structs in `Common/SettingsStore.h`.**

Add these shapes:

```cpp
enum class FileActionMatchKind : uint8_t
{
    Default,
    Extension,
    Pattern,
};

struct FileActionMatch
{
    FileActionMatchKind kind = FileActionMatchKind::Default;
    std::wstring value;
    bool operator==(const FileActionMatch&) const = default;
};

struct FileActionApplicability
{
    std::vector<FileActionMatch> matches;
    std::vector<std::wstring> computerNames;
    bool operator==(const FileActionApplicability&) const = default;
};

struct FileActionDefinition
{
    std::wstring id;
    std::wstring displayName;
    bool enabled = true;
    FileActionKind kind = FileActionKind::ExternalProgram;
    std::wstring pluginId;
    std::wstring executablePath;
    std::wstring arguments;
    std::wstring workingDirectory;
    FileActionApplicability appliesTo;
    bool operator==(const FileActionDefinition&) const = default;
};

struct ViewerAssociationRule
{
    FileActionMatch match;
    std::wstring computerName;
    std::wstring viewActionId;
    std::wstring alternateViewActionId;
    bool operator==(const ViewerAssociationRule&) const = default;
};

struct EditorAssociationRule
{
    FileActionMatch match;
    std::wstring computerName;
    std::wstring editActionId;
    std::wstring alternateEditActionId;
    std::wstring editNewActionId;
    bool operator==(const EditorAssociationRule&) const = default;
};

struct ViewerFileActionsSettings
{
    std::vector<FileActionDefinition> actions;
    std::vector<ViewerAssociationRule> associations;
    bool operator==(const ViewerFileActionsSettings&) const = default;
};

struct EditorFileActionsSettings
{
    std::vector<FileActionDefinition> actions;
    std::vector<EditorAssociationRule> associations;
    bool operator==(const EditorFileActionsSettings&) const = default;
};

struct FileActionsSettings
{
    ViewerFileActionsSettings viewers;
    EditorFileActionsSettings editors;
    bool operator==(const FileActionsSettings&) const = default;
};

struct UserMenuSettings
{
    std::vector<FileActionDefinition> actions;
    bool operator==(const UserMenuSettings&) const = default;
};
```

Then change `Settings`:

```cpp
uint32_t schemaVersion = 16;
FileActionsSettings fileActions;
UserMenuSettings userMenu;
```

Remove `FileActionComputerOverride`, `FileActionSetSettings`, root `viewers`, root `editors`, and viewer extension mappings from `ExtensionsSettings`.

- [x] **Step 2: Seed v16 defaults without migration code.**

Add default viewer actions for built-in viewer plugins and default viewer associations directly in `ViewerFileActionsSettings`.

Minimum defaults:

| Match | F3 action | Alt+F3 action |
|---|---|---|
| `.txt`, `.log`, `.xml`, `.ini`, `.cfg`, `.csv`, `.diff`, `.patch`, `.rej` | `viewer-text` | empty |
| `.md` | `viewer-markdown` | `viewer-text` |
| `.json`, `.json5`, `.jsonl`, `.ndjson` | `viewer-json` | `viewer-text` |
| `.html`, `.htm`, `.pdf` | `viewer-web` | `viewer-text` |
| common WIC image formats | `viewer-imgraw` | empty |
| common RAW photo formats | `viewer-imgraw` | empty |
| `*` | `viewer-text` | empty |

Do not read old persisted mappings to create these defaults.

Editor defaults:

- Start with no external editor actions unless existing project defaults define a stable external editor.
- A missing editor action is a visible configuration state, not silent fallback.
- The Editors page must show the empty state and make it easy to add an action.

- [x] **Step 3: Replace parser/writer helpers.**

In `Common/Common/SettingsStore.cpp`, add helpers:

```cpp
[[nodiscard]] FileActionMatchKind ParseFileActionMatchKind(std::string_view value) noexcept;
[[nodiscard]] const char* FileActionMatchKindToString(FileActionMatchKind kind) noexcept;
void ParseFileActionMatch(yyjson_val* obj, FileActionMatch& out) noexcept;
void ParseFileActionApplicability(yyjson_val* obj, FileActionApplicability& out) noexcept;
void ParseFileActionDefinitionArray(yyjson_val* arr, std::vector<FileActionDefinition>& out) noexcept;
void ParseViewerAssociations(yyjson_val* arr, std::vector<ViewerAssociationRule>& out) noexcept;
void ParseEditorAssociations(yyjson_val* arr, std::vector<EditorAssociationRule>& out) noexcept;
void ParseFileActions(yyjson_val* root, FileActionsSettings& out) noexcept;
```

Writer helpers must use yyjson copy APIs for dynamic strings:

```cpp
yyjson_mut_val* key = yyjson_mut_strncpy(doc, keyUtf8.data(), keyUtf8.size());
yyjson_mut_val* val = yyjson_mut_strncpy(doc, valueUtf8.data(), valueUtf8.size());
yyjson_mut_obj_add(obj, key, val);
```

Do not pass temporary strings to non-copy yyjson APIs.

- [x] **Step 4: Update schema to v16.**

Schema requirements:

- Root `required` still includes `schemaVersion`.
- Root `additionalProperties` remains `false`.
- Root includes `fileActions`, `userMenu`, and `extensions`.
- Root excludes `viewers` and `editors`.
- `extensionsSettings` excludes `openWithViewerByExtension`.
- `editorAssociationRule` requires `match` and allows the three command ids.
- `viewerAssociationRule` requires `match` and allows the two command ids.
- `fileActionMatch` validates:
  - `default`: no non-empty `value`.
  - `extension`: `value` matches `^\.[^\\/:*?"<>|.][^\\/:*?"<>|]*$`.
  - `pattern`: non-empty string without path separators.

- [x] **Step 5: Run focused settings selftests.**

Run after implementation:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "settings_store_file_actions_v16_roundtrip|settings_store_v16_rejects_legacy_viewer_editor_shape"
```

Expected: both cases pass and archive results under `Specs/TestRuns/<MachineHash>/Commands/<timestamp>`.

## Task 3: Build the Explainable Resolver

**Files:**
- Create: `RedSalamander/FileActionResolver.h`
- Create: `RedSalamander/FileActionResolver.cpp`
- Modify: `RedSalamander/FileActionLauncher.h`
- Modify: `RedSalamander/FileActionLauncher.cpp`
- Modify: `RedSalamander/RedSalamander.vcxproj`
- Modify: `RedSalamander/RedSalamander.vcxproj.filters`

- [x] **Step 1: Add resolver public API.**

Use value-returning result structs so Preferences can display the same reason commands use:

```cpp
namespace FileActionResolver
{
enum class Command : uint8_t
{
    View,
    AlternateView,
    Edit,
    AlternateEdit,
    EditNew,
};

struct Request
{
    Command command = Command::View;
    std::filesystem::path filePath;
    std::wstring computerName;
};

struct Result
{
    const Common::Settings::FileActionDefinition* action = nullptr;
    const void* associationRule = nullptr;
    FileActionResolutionReason reason = FileActionResolutionReason::NoAssociation;
    std::wstring actionId;
    std::wstring matchText;
    std::wstring computerName;
};

[[nodiscard]] Result ResolveViewerAction(const Common::Settings::ViewerFileActionsSettings& settings, const Request& request) noexcept;
[[nodiscard]] Result ResolveEditorAction(const Common::Settings::EditorFileActionsSettings& settings, const Request& request) noexcept;
[[nodiscard]] std::vector<const Common::Settings::FileActionDefinition*> CollectApplicableActions(
    std::span<const Common::Settings::FileActionDefinition> actions,
    const std::filesystem::path& filePath,
    std::wstring_view computerName);
[[nodiscard]] std::wstring FormatResolutionReasonForUi(const Result& result, Command command);
}
```

- [x] **Step 2: Implement match normalization.**

Rules:

- Extension uses `path.extension()`, lowercased, leading dot.
- Pattern matches `path.filename()` case-insensitively.
- Default matches every file.
- Empty extension only matches default or pattern rules.

Use a deterministic local glob matcher if no project helper already exists. It must support `*` and `?` only and must not allocate inside the character loop.

- [x] **Step 3: Implement priority buckets.**

Resolver loop:

```cpp
for (const PriorityBucket bucket : kPriorityOrder)
{
    for (const AssociationRule& rule : associations)
    {
        if (! RuleIsInBucket(rule, bucket, request.computerName))
        {
            continue;
        }
        if (! MatchFile(rule.match, request.filePath))
        {
            continue;
        }
        return ResolveActionIdFromRule(rule, request, bucket);
    }
}
return {.reason = FileActionResolutionReason::NoAssociation};
```

If an association points to a missing, disabled, filtered, or wrong-kind action, return that failure reason and stop. Do not silently fall through to a lower-priority rule, because the user explicitly configured the higher-priority rule.

- [x] **Step 4: Instrument resolver cost.**

Emit metrics:

- `fileactions.resolve_us`
- `fileactions.collect_applicable_us`

Set `value0` to association/action count and `value1` to matched bucket index or returned item count.

- [x] **Step 5: Run resolver tests.**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "file_action_resolution_v16_explains_priority|file_action_resolution_v16_command_keys_are_distinct"
```

Expected: both pass, with `fileactions.resolve_us` present in `perf_metrics.jsonl`.

## Task 4: Rewire Runtime Commands

**Files:**
- Modify: `RedSalamander/FolderWindow.Viewers.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.cpp`
- Modify: `RedSalamander/SettingsHotReload.cpp`
- Modify: `RedSalamander/resource.h`
- Modify: `RedSalamander/res/*.rc`

- [x] **Step 1: Update View and Alternate View.**

Use:

```cpp
const FileActionResolver::Request request{
    .command = alternate ? FileActionResolver::Command::AlternateView : FileActionResolver::Command::View,
    .filePath = focusedPath,
    .computerName = computerName,
};
const FileActionResolver::Result resolution = FileActionResolver::ResolveViewerAction(_settings->fileActions.viewers, request);
```

Behavior:

- Viewer plugin action opens the internal viewer.
- External program action launches through `FileActionLauncher`.
- Missing action, disabled action, wrong kind, or macro failure shows localized pane alert with action id, file name, and reason.
- F3 no longer consults `extensions.openWithViewerByExtension`.

- [x] **Step 2: Update View With.**

Dynamic `View With` menu uses `CollectApplicableActions(settings.fileActions.viewers.actions, filePath, computerName)`.

Parameterized `cmd/pane/viewWith/<viewerId>` resolves that action id directly from `settings.fileActions.viewers.actions`, then validates applicability and launchability.

- [x] **Step 3: Update Edit and Alternate Edit.**

Use `ResolveEditorAction(settings.fileActions.editors, request)` for:

- `cmd/pane/edit`
- `cmd/pane/alternateEdit`

Behavior:

- Only external program editor actions are valid for editor commands.
- Disabled/missing/unavailable editor uses existing pane alert style.
- Ctrl+Shift+F4 must never fall back to F4's action when the alternate field is empty.

- [x] **Step 4: Update Edit With.**

Dynamic `Edit With` menu and `cmd/pane/editWith/<editorId>` use `settings.fileActions.editors.actions`.

- [x] **Step 5: Update Edit New.**

In `FolderWindow.FileSystem.cpp`, change the prompt's preferred editor selection from primary edit resolution to explicit Edit New resolution:

```cpp
const FileActionResolver::Request request{
    .command = FileActionResolver::Command::EditNew,
    .filePath = candidatePath,
    .computerName = _computerName,
};
const FileActionResolver::Result resolution = FileActionResolver::ResolveEditorAction(*_editorSettings, request);
```

The Editor combo still lists all applicable editor actions for the typed extension/computer, but the selected item is `editNewActionId` when it resolves. If no Edit New action resolves, the combo is enabled only when applicable actions exist and starts with `(none)`.

- [x] **Step 6: Update manual reload.**

`cmd/app/rereadAssociations` applies:

- `settings.fileActions`
- `settings.userMenu`
- `settings.extensions.openWithFileSystemByExtension`
- plugins, shortcuts, and current existing reload sections

It must not reference root `viewers`, root `editors`, or `openWithViewerByExtension`.

- [x] **Step 7: Run command selftests.**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "view_|alternate_view|edit_|alternate_edit|editNew|view_with|edit_with|rereadAssociations"
```

Expected: updated command tests pass with no legacy viewer-extension fallback.

## Task 5: Build Shared DxUi File Actions Preferences Components

**Files:**
- Create: `RedSalamander/Preferences.FileActions.h`
- Create: `RedSalamander/Preferences.FileActions.cpp`
- Modify: `RedSalamander/Preferences.Internal.h`
- Modify: `RedSalamander/Preferences.Dialog.cpp`
- Modify: `RedSalamander/RedSalamander.vcxproj`
- Modify: `RedSalamander/RedSalamander.vcxproj.filters`

- [x] **Step 1: Add shared page state.**

Use owned models and non-owning DxUi child handles:

```cpp
struct FileActionsPageState
{
    enum class Tab : uint8_t { Associations, Actions };
    Tab selectedTab = Tab::Associations;
    std::wstring selectedAssociationKey;
    std::wstring selectedActionId;
    std::filesystem::path previewFilePath;
    FileActionResolver::Command previewCommand = FileActionResolver::Command::View;
};
```

The DxUi control tree is owned by the retained `Panel` child collection. Page classes may cache non-owning control pointers only while attached to that panel, and must set them to `nullptr` in `DetachDxPageHost()`. No `new`, no `delete`, no owning raw pointers.

- [x] **Step 2: Add shared grid models.**

Create models:

- `ViewerAssociationsGridModel`
- `EditorAssociationsGridModel`
- `FileActionsGridModel`

All models:

- Implement `IDxGridModel`.
- Use stable row ids derived from normalized match + computer + action id.
- Keep column widths responsive and bounded.
- Expose `DebugVisibleWorkMetrics()` for perf selftests.

- [x] **Step 3: Add inline forms.**

Shared controls:

- Segmented tabs: Associations, Actions.
- Match kind combo: Default, Extension, Pattern.
- Match value text field.
- Computer combo/text: Any, current computer, custom.
- Action combo fields filtered by page family and command.
- Add, Update, Remove, Reset Defaults, Test buttons.
- Resolution preview fields.

Do not use visible instructional paragraphs to explain the UI. Field labels and validation text are enough.

- [x] **Step 4: Add validation.**

Validation rules:

- Action IDs are unique case-insensitively inside each action list.
- Action ID must be stable ASCII-like settings id.
- Viewer plugin actions require non-empty `pluginId`.
- External program actions require non-empty `executablePath`.
- `arguments` and `workingDirectory` macro syntax must validate through `FileActionLauncher::ExpandMacros` with a synthetic context that contains all supported macro values.
- Association action ids must either be empty or refer to an action in the same family.
- Duplicate association keys are rejected.
- Action applicability filters normalize extensions and computer names.

- [x] **Step 5: Add preview path.**

Preview uses the real resolver:

```cpp
const FileActionResolver::Result result = pageKind == PageKind::Viewers
    ? FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request)
    : FileActionResolver::ResolveEditorAction(settings.fileActions.editors, request);
```

The displayed action and reason must match command execution for the same settings, path, computer, and command.

- [x] **Step 6: Add layout/perf metrics.**

Emit:

- `preferences.ui.fileactions.viewers_layout_us`
- `preferences.ui.fileactions.editors_layout_us`
- `preferences.ui.fileactions.associations_visible_rows`
- `preferences.ui.fileactions.actions_visible_rows`

Use bounded visible work; scrolling a large grid must not rebuild every row's text every frame.

## Task 6: Replace Viewers Preferences Page

**Files:**
- Modify: `RedSalamander/Preferences.Viewers.h`
- Modify: `RedSalamander/Preferences.Viewers.cpp`
- Modify: `RedSalamander/Preferences.Internal.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`

- [x] **Step 1: Remove legacy extension viewer mapping UI.**

Delete the page concepts that edit `state.workingSettings.extensions.openWithViewerByExtension`.

Viewers page data source becomes:

```cpp
state.workingSettings.fileActions.viewers
```

- [x] **Step 2: Add Viewers columns.**

Associations columns:

- Match
- Computer
- F3 View
- Alt+F3 Alternate View
- Status

Actions columns:

- Name
- Type
- Applies To
- Computer
- Status

- [x] **Step 3: Add Viewers debug snapshot fields.**

Extend Preferences debug snapshot with:

```cpp
size_t viewersAssociationRowCount = 0;
size_t viewersActionRowCount = 0;
std::wstring viewersPreviewActionId;
std::wstring viewersPreviewReason;
bool viewersActionsTabVisible = false;
bool viewersAssociationsTabVisible = false;
```

- [x] **Step 4: Add failing-to-green UI tests.**

Required selftest names:

- `cmd_preferences_dialog_viewers_file_actions_tabs_are_dxui`
- `cmd_preferences_dialog_viewers_association_apply_roundtrip`
- `cmd_preferences_dialog_viewers_action_apply_roundtrip`
- `cmd_preferences_dialog_viewers_resolution_preview_explains_command`
- `cmd_preferences_dialog_viewers_large_grid_visible_work_bounded`

Required assertions:

- One shared visible page-host child.
- No visible legacy listview/edit/combo/button child controls.
- UIA descendants include grid, value, selection, combo, and invoke patterns.
- Editing `.txt` F3 and Alt+F3 marks dirty and persists to `fileActions.viewers.associations`.
- Preview for `C:\src\main.cpp` shows the same action id as `FileActionResolver::ResolveViewerAction`.

## Task 7: Replace Editors Preferences Page

**Files:**
- Modify: `RedSalamander/Preferences.Editors.h`
- Modify: `RedSalamander/Preferences.Editors.cpp`
- Modify: `RedSalamander/Preferences.Internal.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`

- [x] **Step 1: Remove global primary/alternate-only page.**

Editors page data source becomes:

```cpp
state.workingSettings.fileActions.editors
```

- [x] **Step 2: Add Editors columns.**

Associations columns:

- Match
- Computer
- F4 Edit
- Ctrl+Shift+F4 Alternate Edit
- Shift+F4 Edit New
- Status

Actions columns:

- Name
- Type
- Applies To
- Computer
- Status

- [x] **Step 3: Add editor-specific validation.**

Editors page must reject `viewerPlugin` actions for editor commands. The action type combo on Editors should show only `External program`.

- [x] **Step 4: Add failing-to-green UI tests.**

Required selftest names:

- `cmd_preferences_dialog_editors_file_actions_tabs_are_dxui`
- `cmd_preferences_dialog_editors_association_apply_roundtrip`
- `cmd_preferences_dialog_editors_action_apply_roundtrip`
- `cmd_preferences_dialog_editors_edit_new_association_apply_roundtrip`
- `cmd_preferences_dialog_editors_resolution_preview_explains_command`
- `cmd_preferences_dialog_editors_large_grid_visible_work_bounded`

Required assertions:

- Editing `.cpp` F4, Ctrl+Shift+F4, and Shift+F4 stores three distinct action ids.
- Preview command selector can show F4, Ctrl+Shift+F4, and Shift+F4 outcomes.
- The old two-combo-only Editors page no longer appears.

## Task 8: Rewrite App-Owned Command Option Dialogs in DxUi

**Files:**
- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/resource.h`
- Modify: `RedSalamander/res/*.rc`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

- [x] **Step 1: Add failing Win32-template retirement guards.**

Add selftests:

- `cmd_pane_makeFileList_options_dialog_uses_dxui_not_win32_template`
- `cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template`

Required assertions:

```cpp
state.Require(snapshot.usesDxUiHost, L"Command options dialog should render through a DxUi host.");
state.Require(snapshot.visibleNativeChildControlCount <= 1u,
              L"Command options dialog should not expose a Win32 dialog-template form.");
state.Require(snapshot.dialogClassName != L"#32770",
              L"Command options dialog should not be an app-owned Win32 dialog template.");
state.Require(snapshot.uiaPatternStats.has_value(), L"Command options dialog should expose UIA patterns through DxUi.");
```

Expected RED result: Make File List and Change Attributes still use `DialogBoxParamW` and native dialog resources.

- [x] **Step 2: Add a small DxUi command-dialog host.**

Implement a reusable command-dialog window helper in the command file or a focused helper if the file becomes too large. The helper may use a Win32 owner HWND, but the visible content must be a DxUi `WindowHost` and retained `Panel`.

Required ownership rules:

- Use `wil::unique_hwnd` for the dialog host window.
- Use `std::unique_ptr` for owned dialog state.
- Cache non-owning DxUi child pointers only while the retained panel owns those controls.
- Reset cached pointers during teardown.
- Do not use `DialogBoxParamW`, `EndDialog`, `CreateDialogParam`, or Win32 dialog templates for these command option surfaces.

- [x] **Step 3: Rewrite Make File List options as DxUi.**

Replace `MakeFileListDialogProc` and `PromptForMakeFileListOptions` with a DxUi modal command surface.

Required controls:

- Source segmented/radio choice: Selection, Current Folder.
- Recursive checkbox.
- Format segmented choice: Text, CSV, JSON.
- Output segmented choice: Clipboard, File.
- Text macro field.
- Output file field and Browse button.
- Include-field checkboxes: Name, Full path, Size, Modified, Attributes, Directories.
- OK and Cancel buttons.

Browse behavior:

- The Browse button may use an OS-owned save-file picker such as `IFileSaveDialog` or the current system save-file picker.
- That picker is allowed because it is not a RedSalamander-owned dialog.

Persistence and command behavior remain exactly as the existing Make File List command contract describes.

- [x] **Step 4: Rewrite Change Attributes options as DxUi.**

Replace `ChangeAttributesDialogProc` and `PromptForChangeAttributes` with a DxUi modal command surface.

Required controls:

- Attribute state controls for Read-only, Hidden, System, and Archive.
- Remove extra streams checkbox.
- OK and Cancel buttons.
- Validation/status line for no-op choices.

Runtime behavior remains:

- Apply selected attribute changes.
- Remove extra streams when requested.
- Show the same report summary for changed attributes, removed streams, skipped items, and failures.

- [x] **Step 5: Remove app-owned dialog resources.**

Remove or stop referencing:

- `IDD_MAKE_FILE_LIST`
- `IDD_CHANGE_ATTRIBUTES`
- native control ids used only by those two dialog templates

Keep localized strings for DxUi labels and validation. Resource strings must use positional placeholders.

- [x] **Step 6: Run command-dialog selftests.**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "makeFileList|changeAttributes"
```

Expected:

- Make File List behavior and saved-option tests still pass.
- Change Attributes attribute/stream/report tests still pass.
- New DxUi/no-template guards pass.

## Task 9: Update Edit New Prompt Behavior

**Files:**
- Modify: `RedSalamander/FolderWindow.FileSystem.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp`
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp`

- [x] **Step 1: Add explicit Edit New debug visibility.**

Extend `FolderViewEditNewPromptDebugSnapshot`:

```cpp
std::wstring resolvedEditNewActionId;
std::wstring resolvedEditNewReason;
```

- [x] **Step 2: Select editor by `editNewActionId`.**

When the typed name changes, recompute:

```cpp
const auto resolution = FileActionResolver::ResolveEditorAction(*_editorSettings, requestForEditNew);
_selectedEditorActionId = resolution.action ? resolution.action->id : L"";
```

The combo item list still comes from applicable editor actions for the candidate file.

- [x] **Step 3: Update existing tests.**

Change `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file` so it sets:

```cpp
g_settings.fileActions.editors.associations.push_back(EditorAssociationRule{
    .match = {.kind = FileActionMatchKind::Extension, .value = L".editnew"},
    .editNewActionId = L"edit-new-editor",
});
```

Add a second applicable editor and assert it appears in the combo but is not selected when `editNewActionId` points to the preferred editor.

- [x] **Step 4: Add no-Edit-New association test.**

Required selftest:

- `cmd_pane_editNew_prompt_no_edit_new_association_starts_with_none`

Assertions:

- Combo is enabled when applicable editors exist.
- Selected editor id is empty.
- Creating the file succeeds without launching an editor unless the user selects one.

## Task 10: Update User Menu Shape Without Changing Behavior

**Files:**
- Modify: `RedSalamander/Preferences.UserMenu.h`
- Modify: `RedSalamander/Preferences.UserMenu.cpp`
- Modify: `RedSalamander/FolderWindow.UserMenu.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`

- [x] **Step 1: Replace `settings.userMenu.actions` access through `FileActionSetSettings`.**

New type:

```cpp
Common::Settings::UserMenuSettings userMenu;
```

Only `actions` is persisted. There are no default primary/alternate ids or association maps in user menu settings.

- [x] **Step 2: Keep runtime behavior stable.**

Existing User Menu behavior remains:

- Dynamic menu from enabled/applicable external actions.
- `cmd/pane/userMenu/<itemId>` launches the exact action.
- Macro expansion, selected paths file, missing executable alert, and computer/extension filters keep existing behavior.

- [x] **Step 3: Update selftests.**

Update existing User Menu selftests to seed `g_settings.userMenu.actions` on the new type and confirm schema output has no default/mapping fields.

## Task 11: Documentation and Authoritative Specs

**Files:**
- Modify: `Specs/Core/Core_SettingsStore.md`
- Modify: `Specs/UI/UI_CommandMenuKeyboard.md`
- Modify: `Specs/UI/UI_PreferencesDialog.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `Docs/Preferences.md`
- Modify: `Docs/SettingsFile.md`
- Modify: `Docs/Viewers.md`
- Modify: `Docs/UserGuide.md`

- [x] **Step 1: Update Settings Store spec.**

Document:

- `schemaVersion: 16`.
- `fileActions.viewers.actions`.
- `fileActions.viewers.associations`.
- `fileActions.editors.actions`.
- `fileActions.editors.associations`.
- Explicit `editNewActionId`.
- Resolver priority and reason behavior.
- Removed root `viewers`, root `editors`, and `extensions.openWithViewerByExtension`.
- No v15 migration.

- [x] **Step 2: Update command spec.**

In `Specs/UI/UI_CommandMenuKeyboard.md`, make command contracts refer to the new resolver:

- `cmd/pane/view`: Viewers association `viewActionId`.
- `cmd/pane/alternateView`: Viewers association `alternateViewActionId`.
- `cmd/pane/edit`: Editors association `editActionId`.
- `cmd/pane/alternateEdit`: Editors association `alternateEditActionId`.
- `cmd/pane/editNew`: Editors association `editNewActionId` for preferred editor selection.
- `cmd/pane/viewWith/<viewerId>` and `cmd/pane/editWith/<editorId>`: direct action id dispatch.

- [x] **Step 3: Update Preferences spec.**

Add the low-fi UX tables from this plan to `Specs/UI/UI_PreferencesDialog.md`, with the rule that the preview must use the production resolver.

- [x] **Step 4: Update command-dialog UI spec.**

Document the Dialog Ownership Rule in `Specs/UI/UI_CommandMenuKeyboard.md` or `Specs/UI/UI_PreferencesDialog.md`:

- RedSalamander-owned command option dialogs must render visible content through DxUi.
- Win32 dialog templates and native child-control forms are not allowed for app-owned command dialogs.
- OS shell/common pickers such as `IFileOpenDialog` and save-file/folder pickers are allowed system UI.
- Make File List and Change Attributes option dialogs are covered by this rule.

- [x] **Step 5: Update user docs.**

Document how users answer:

```text
For .txt, what happens when I press F3, Alt+F3, F4, Ctrl+Shift+F4, or Shift+F4?
```

Include one viewer example and one editor example.

## Task 12: Perf and Regression Evidence

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
- Modify: `Specs/TestRuns/README.md` only if new artifact naming needs explanation.

- [x] **Step 1: Define protected scenarios.**

Protected scenarios:

- Resolve one command for one focused item while navigating.
- Populate View With and Edit With dynamic menus.
- Recompute Edit New preferred editor while typing.
- Layout and scroll large Viewers/Editors Preferences grids.
- Open and validate Make File List and Change Attributes DxUi command option dialogs.

- [x] **Step 2: Emit deterministic perf artifacts.**

Metrics required:

- `fileactions.resolve_us`
- `fileactions.collect_applicable_us`
- `fileaction.editnew.editor_combo_us`
- `preferences.ui.fileactions.viewers_layout_us`
- `preferences.ui.fileactions.editors_layout_us`
- `preferences.ui.fileactions.associations_visible_rows`
- `preferences.ui.fileactions.actions_visible_rows`
- `commands.dialog.makeFileList_layout_us`
- `commands.dialog.changeAttributes_layout_us`

- [x] **Step 3: Add large-grid tests.**

Seed at least:

- 500 viewer associations.
- 500 editor associations.
- 100 actions per page.

Assert visible row work stays bounded by visible rows plus a small buffer, not total rows.

- [x] **Step 4: Archive focused runs.**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "file_action_resolution_v16|cmd_preferences_dialog_viewers_|cmd_preferences_dialog_editors_|cmd_pane_editNew_prompt|makeFileList|changeAttributes"
```

Expected artifacts:

- `Specs/TestRuns/<MachineHash>/Commands/<timestamp>/results.json`
- `Specs/TestRuns/<MachineHash>/Commands/<timestamp>/trace.txt`
- `Specs/TestRuns/<MachineHash>/Commands/<timestamp>/perf_metrics.jsonl`

## Task 13: Build, Release, and Closeout

**Files:**
- Modify: this plan, then move to `Specs/Plans/Done/` only after all implementation and spec updates are complete.

- [x] **Step 1: Run Debug build.**

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Expected: success.

- [x] **Step 2: Run focused Debug selftests.**

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "settings_store_file_actions_v16|file_action_resolution_v16|view_with|edit_with|alternate_view|alternate_edit|editNew|cmd_preferences_dialog_viewers_|cmd_preferences_dialog_editors_|user_menu|makeFileList|changeAttributes"
```

Expected: success with archived test run.

- [x] **Step 3: Run Release build.**

```powershell
.\build.ps1 -Configuration Release -ProjectName RedSalamander
```

Expected: success.

- [x] **Step 4: Run stale command/docs guards.**

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "implemented_menu_labels_not_todo|command_registry|settings_schema"
```

Expected: no stale not-implemented labels and schema v16 is shipped.

- [x] **Step 5: Close the plan.**

Before moving the plan:

- Confirm `Specs/Core/Core_SettingsStore.md` is updated with durable settings behavior.
- Confirm `Specs/UI/UI_PreferencesDialog.md` is updated with durable UI behavior.
- Confirm `Specs/UI/UI_CommandMenuKeyboard.md` is updated with durable command behavior.
- Record Debug build log, Release build log, focused test run paths, and key perf metrics in this checklist.
- Move this file to `Specs/Plans/Done/UI_FileActionsPreferencesRedesignPlan_2026-05-03.md`.

## Acceptance Criteria

- The UI lets a user inspect and edit the answer to:
  - `.txt` + F3
  - `.txt` + Alt+F3
  - `.txt` + F4
  - `.txt` + Ctrl+Shift+F4
  - `.txt` + Shift+F4
- Viewers and Editors pages expose Associations and Actions tabs.
- Preferences preview uses the production resolver and explains the winning rule.
- `editNewActionId` exists and Shift+F4 no longer depends on primary edit.
- App-owned Make File List and Change Attributes option dialogs use DxUi surfaces, not Win32 dialog templates or native child-control forms.
- OS shell/common pickers remain allowed for Browse/path-selection actions because they are system-owned UI.
- Root `viewers`, root `editors`, and `extensions.openWithViewerByExtension` are gone from schema v16.
- Runtime command resolution follows:
  1. computer + extension/pattern
  2. global extension/pattern
  3. computer default
  4. global default
- Higher-priority invalid configuration reports the specific failure and does not silently fall through.
- View With and Edit With menus list applicable enabled actions from the new action lists.
- User Menu behavior remains stable on its dedicated `actions` settings shape.
- Settings, command, Preferences UI, Edit New prompt, and perf selftests pass in Debug.
- Release build passes.
- Perf evidence is archived under `Specs/TestRuns/`.

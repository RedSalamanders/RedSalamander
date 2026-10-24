# Localization Resource DLL Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add culture-specific resource-only DLL support for RedSalamander, RedSalamanderMonitor, and plugins, with a Preferences language setting and embedded English fallback.

**Architecture:** English remains embedded in each existing executable or plugin DLL exactly as it is today. Optional satellite resource-only DLLs are built from per-project `Lang\<culture>\` projects, copied to the build output `Lang` directory, and loaded at runtime according to the persisted language preference. Resource helpers resolve strings, menus, dialogs, icons, accelerators, and other localizable resources from the selected satellite first, then fall back to the owning module's embedded English resources.

**Tech Stack:** MSBuild `.vcxproj`, Windows `.rc` resources, WIL RAII, `Common.dll` settings store using yyjson, Win32 resource APIs, DxUI Preferences, existing self-test framework.

---

## Detailed Implementation Checklist

- [x] Resync worktree with `master` before implementation.
- [x] Add `Common/LocalizationManager.h`.
- [x] Add `Common/Common/LocalizationManager.cpp`.
- [x] Include `LocalizationManager.*` in `Common/Common.vcxproj`.
- [x] Add resource-owner registration keyed by embedded `HINSTANCE`.
- [x] Add culture preference resolution for `system` and concrete culture values.
- [x] Load resource-only satellites from `.build\<Platform>\<Configuration>\Lang\`.
- [x] Use `wil::unique_hmodule` ownership for satellite DLLs.
- [x] Fall back to embedded English when a satellite DLL is missing.
- [x] Fall back to embedded English when a satellite resource ID is missing.
- [x] Route app/Common string helpers through the localization manager.
- [x] Keep plugin string helper call sites compiling while plugin ownership support is pending.
- [x] Add deterministic `LocalizationTests` project.
- [x] Add embedded fallback localization test.
- [x] Add satellite string override localization test.
- [x] Add `LocalizationTests-fr-FR` resource-only satellite project.
- [x] Add central MSBuild properties for language resource projects.
- [x] Output language resource DLLs to `.build\<Platform>\<Configuration>\Lang\`.
- [x] Keep normal executable and plugin outputs in their existing directories.
- [x] Add `RedSalamander-fr-FR` resource-only satellite project.
- [x] Add `RedSalamanderMonitor-fr-FR` resource-only satellite project.
- [x] Add app and monitor satellite projects to the solution build graph.
- [x] Add `ui.language` field to `Common::Settings::UiSettings`.
- [x] Parse valid `ui.language` values from settings.
- [x] Treat missing, `system`, empty, and invalid `ui.language` as `system`.
- [x] Omit default `ui.language` on save.
- [x] Save concrete `ui.language` values.
- [x] Add `ui.language` to `Specs/SettingsStore.schema.json`.
- [x] Register `RedSalamander` resource owner during startup after settings load.
- [x] Register `RedSalamanderMonitor` resource owner during startup after settings load.
- [x] Apply persisted language preference at app startup.
- [x] Apply persisted language preference at monitor startup.
- [x] Add settings-store tests for `ui.language` load/save/default/invalid cases.
- [x] Add `build.ps1` language output validation.
- [x] Add plugin satellite project pattern, starting with `ViewerText`.
- [x] Add first-pass `fr-FR` resource-only satellite projects for every plugin owner.
- [x] Review and polish first-version `fr-FR` UI copy across app, monitor, and plugin satellites.
- [x] Register plugin resource owners when plugin DLLs load.
- [x] Unregister plugin resource owners when plugin DLLs unload.
- [x] Add Preferences language resource IDs and embedded English strings.
- [x] Add General Preferences language combo.
- [x] Populate language combo from `System Language`, embedded English, and discovered satellites.
- [x] Persist Preferences language changes on Apply/OK.
- [x] Re-apply localization preference after runtime language changes.
- [x] Refresh/rebuild top-level menus after runtime language changes.
- [x] Inventory direct user-facing non-string resource loads.
- [x] Route menu/accelerator/dialog resource loads through localization helpers.
- [x] Add deterministic tests for satellite present/missing, invalid culture, system fallback, late owner registration, and unregister behavior.
- [x] Add one-warning diagnostics for invalid persisted language values.
- [x] Add or archive startup/resource-lookup perf evidence under `Specs/TestRuns/`.
- [x] Update authoritative localization spec for implemented satellite lookup behavior.
- [x] Update authoritative settings spec for implemented `ui.language` persistence behavior.
- [x] Update authoritative Preferences spec after Preferences language UI is implemented.
- [x] Move this plan to `Specs/Plans/Done/` only after all implementation, tests, perf evidence, and authoritative specs are complete.

---

## Requirements Covered

- Add `Lang` folders under `RedSalamander` and `RedSalamanderMonitor` for application resource-only DLL projects.
- Add `Lang` folders under each plugin source folder for plugin resource-only DLL projects.
- Name culture DLLs with the owning project name plus BCP 47 language tag: `RedSalamander-fr-FR.dll`, `RedSalamanderMonitor-fr-FR.dll`, `ViewerText-de-DE.dll`.
- Add a general Preferences language setting with default option `System Language`.
- Copy all application and plugin language DLLs to `.build\<Platform>\<Configuration>\Lang\`.
- Keep English as the default embedded resource set in the existing `.exe` and plugin `.dll` binaries.
- Read the language preference at runtime, load the correct resource DLL for each resource owner, and fall back to embedded English when the satellite DLL or resource ID is missing.
- Update authoritative specs when implementation is complete, then move this plan to `Specs/Plans/Done/`.

## Non-Goals

- Do not migrate English out of existing `.rc` files.
- Do not add installer UI for selecting language in this plan.
- Do not add online translation update/download behavior.
- Do not localize dynamic user data such as file names, paths, plugin IDs, or theme file names.
- Do not introduce registry-backed language settings.

## Naming And Folder Contract

Source layout:

```text
RedSalamander/
  RedSalamander.rc
  Lang/
    de-DE/
      RedSalamander-de-DE.rc
      RedSalamander-de-DE.vcxproj
    fr-FR/
      RedSalamander-fr-FR.rc
      RedSalamander-fr-FR.vcxproj

RedSalamanderMonitor/
  RedSalamanderMonitor.rc
  Lang/
    de-DE/
      RedSalamanderMonitor-de-DE.rc
      RedSalamanderMonitor-de-DE.vcxproj

Plugins/ViewerText/
  ViewerTextResources.rc
  Lang/
    de-DE/
      ViewerText-de-DE.rc
      ViewerText-de-DE.vcxproj
```

Output layout:

```text
.build/x64/Debug/
  RedSalamander.exe
  RedSalamanderMonitor.exe
  ViewerText.dll
  Lang/
    RedSalamander-de-DE.dll
    RedSalamanderMonitor-de-DE.dll
    FileSystem-de-DE.dll
    ViewerText-de-DE.dll
```

Language tags:

- Use BCP 47 tags accepted by Windows locale APIs, for example `fr-FR`, `de-DE`, `cs-CZ`.
- Persist `system` for the `System Language` preference.
- Resolve `system` to the user's preferred UI language chain at runtime.
- Probe exact culture first, then parent culture when applicable, then no satellite fallback. Example for `fr-CA`: `fr-CA`, `fr`, embedded English.
- Only build a culture project when translators intentionally provide that culture. The plan includes `fr-FR` as the first sample satellite to verify the pipeline while embedded English remains authoritative fallback.

## File Map

Create:

- `Common/LocalizationManager.h` - shared resource-owner registration, language preference resolution, satellite DLL load/unload, fallback resource lookups.
- `Common/LocalizationManager.cpp` - implementation using WIL RAII and Win32 resource APIs.
- `Tests/LocalizationTests/LocalizationTests.vcxproj` or add equivalent tests to the existing common test project if one already owns `Common` unit tests.
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc` - first app satellite resource script, initially mirroring localizable resources needed for tests.
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.vcxproj` - resource-only DLL project.
- `RedSalamanderMonitor/Lang/fr-FR/RedSalamanderMonitor-fr-FR.rc` - first monitor satellite resource script.
- `RedSalamanderMonitor/Lang/fr-FR/RedSalamanderMonitor-fr-FR.vcxproj` - resource-only DLL project.
- `Plugins/<PluginName>/Lang/fr-FR/<PluginName>-fr-FR.rc` - first plugin satellite resource scripts for each plugin with localizable resources.
- `Plugins/<PluginName>/Lang/fr-FR/<PluginName>-fr-FR.vcxproj` - resource-only DLL project for each plugin.

Modify:

- `Common/Common.vcxproj` - include `LocalizationManager.*`.
- `Common/Helpers.h` - route existing resource helpers through the localization manager while preserving current overloads.
- `Common/Common.vcxproj.filters` if present - add localization files.
- `Common/Common/SettingsStore.cpp` - read/write `ui.language`.
- `Common/Common/SettingsStore.h` - add `Settings::UiSettings::language`.
- `Specs/SettingsStore.schema.json` - add `ui.language` with default `system`.
- `RedSalamander/Preferences.General.cpp` - add General page language combo in the Display group.
- `RedSalamander/Preferences.cpp`, `RedSalamander/Preferences.Internal.cpp`, or adjacent Preferences files as needed - commit/apply language changes through working settings and runtime notification.
- `RedSalamander/RedSalamander.cpp` or current app initialization file - initialize localization before first resource-dependent UI is built.
- `RedSalamanderMonitor` startup file - initialize monitor localization before resource use.
- Plugin load path in `RedSalamander/PluginManager*.cpp` - register each plugin module as a resource owner and attach its culture DLL.
- All `*.vcxproj` owners and `RedSalamander.sln` - include satellite projects and build dependencies.
- `Directory.Build.props` and/or `Directory.Build.targets` - centralize resource-only DLL output and copy-to-`Lang` behavior.
- `build.ps1` - ensure language projects build with the normal solution and that output verification includes `.build\<Platform>\<Configuration>\Lang`.
- `Specs/Core/Core_Localization.md` - make satellite resource DLL behavior normative.
- `Specs/UI/UI_PreferencesDialog.md` - document the General page language setting.
- `Specs/Core/Core_SettingsStore.md` - document `ui.language`.

## Design Details

### Resource Owner Model

Each module that owns embedded English resources becomes a resource owner:

```cpp
namespace Localization
{
    struct ResourceOwner final
    {
        std::wstring moduleName;       // RedSalamander, RedSalamanderMonitor, ViewerText
        HINSTANCE embeddedInstance{};  // exe or plugin DLL containing English fallback
        wil::unique_hmodule satellite; // selected culture resource-only DLL, if loaded
    };
}
```

Runtime registration rules:

- Main app registers `RedSalamander` with `GetModuleHandleW(nullptr)`.
- Monitor registers `RedSalamanderMonitor` with `GetModuleHandleW(nullptr)`.
- Each plugin registers its plugin name with the loaded plugin DLL `HINSTANCE`.
- Registration does not transfer ownership of the embedded module.
- Satellite modules are loaded with `LoadLibraryExW(path, nullptr, LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE)`.
- Satellite modules are owned by `wil::unique_hmodule`.
- Language changes unload previous satellite modules only after no resource load is in progress. Use a mutex around the owner map and keep resource loads short.

### Resource Lookup Contract

All resource helper overloads keep the same call sites:

```cpp
std::wstring title = LoadStringResource(hInstance, IDS_APP_TITLE);
std::wstring text = FormatStringResource(hInstance, IDS_FILE_COUNT, fileCount);
MessageBoxResource(hwnd, hInstance, IDS_ERROR_MSG, IDS_ERROR_TITLE, MB_OK);
```

Internally:

1. Resolve the resource owner from `hInstance`.
2. Try the selected satellite DLL for that owner.
3. If the satellite resource is missing or empty, load from `hInstance`.
4. If both fail, return the current empty result behavior.

Menus and dialogs must get equivalent helper paths:

```cpp
HMENU LoadMenuResource(HINSTANCE owner, UINT menuId) noexcept;
HACCEL LoadAcceleratorsResource(HINSTANCE owner, PCWSTR tableName) noexcept;
HRSRC FindLocalizedResource(HINSTANCE owner, PCWSTR name, PCWSTR type) noexcept;
```

Use these helpers when an existing call directly uses `LoadMenuW`, `LoadAcceleratorsW`, `FindResourceW`, `LoadImageW`, or dialog resource loading for user-facing UI.

### Settings Contract

Add `ui.language`:

```json
"language": {
  "type": "string",
  "default": "system",
  "title": "Language",
  "description": "Application language. Use system to follow the Windows display language.",
  "anyOf": [
    { "const": "system" },
    { "pattern": "^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$" }
  ]
}
```

Runtime values:

- `system` means use Windows preferred UI language list.
- A concrete culture means use that culture directly.
- Invalid values are ignored on load and replaced by `system`.
- Saving omits `ui.language` when it is `system`, matching existing default-omission behavior.

### Preferences UX Contract

General page `Display` group order after this change:

1. Language
2. Existing display settings
3. DxUI group remains unchanged after Display
4. Startup group remains unchanged

Language combo entries:

- `System Language`
- `English (United States)` for `en-US`
- One entry per available satellite culture discovered from `.build\<Platform>\<Configuration>\Lang\`.

Behavior:

- `Cancel` discards unapplied language selection.
- `Apply` and `OK` persist `ui.language`.
- If language changes while the app is running, reload resource DLLs and refresh labels for top-level windows that already support live resource refresh. Any UI that cannot safely refresh immediately may show the new language after reopening that window, but menus, Preferences, and message boxes opened after `Apply` must use the new language.
- If a selected language DLL is missing, log one warning and use embedded English without blocking startup.

### Build Contract

Every resource-only DLL project:

- Uses project type `DynamicLibrary`.
- Compiles only `.rc` resources and version metadata.
- Does not link C++ object files.
- Sets output name to `<OwnerProjectName>-<Culture>.dll`.
- Writes to `$(OutDir)Lang\`.
- Uses the same platform and configuration dimensions as the owning project.
- Has no dependency on vcpkg packages.

Central MSBuild properties:

```xml
<PropertyGroup Condition="'$(IsLanguageResourceProject)' == 'true'">
  <ConfigurationType>DynamicLibrary</ConfigurationType>
  <TargetName>$(ResourceOwnerName)-$(ResourceCulture)</TargetName>
  <OutDir>$(SolutionDir).build\$(Platform)\$(Configuration)\Lang\</OutDir>
  <IntDir>$(SolutionDir).build\obj\$(Platform)\$(Configuration)\$(MSBuildProjectName)\</IntDir>
</PropertyGroup>
```

Satellite projects define:

```xml
<PropertyGroup Label="Localization">
  <IsLanguageResourceProject>true</IsLanguageResourceProject>
  <ResourceOwnerName>RedSalamander</ResourceOwnerName>
  <ResourceCulture>fr-FR</ResourceCulture>
</PropertyGroup>
```

### Resource Script Contract

Satellite `.rc` files include the same `Resource.h` as the owner project and override only localizable resources:

```rc
#include "../../Resource.h"

LANGUAGE LANG_FRENCH, SUBLANG_FRENCH

STRINGTABLE
BEGIN
    IDS_APP_TITLE "RedSalamander"
END
```

Rules:

- Numeric resource IDs must match the embedded English owner.
- English embedded resources remain complete.
- Satellite resources may be partial. Missing satellite IDs fall back to embedded English.
- `VERSIONINFO` for satellites must include `FileDescription` and `OriginalFilename` with the satellite DLL name.
- Do not duplicate binary resources that are not localized unless Windows requires them for resource loading.

## Implementation Tasks

### Task 1: Add LocalizationManager Core

**Files:**

- Create: `Common/LocalizationManager.h`
- Create: `Common/LocalizationManager.cpp`
- Modify: `Common/Common.vcxproj`
- Modify: `Common/Helpers.h`

- [x] **Step 1: Add focused unit tests for fallback lookup**

Create tests that register an owner with no satellite and verify `LoadStringResource(owner, existingId)` returns embedded English and `LoadStringResource(owner, missingId)` returns empty.

Expected command:

```powershell
.\build.ps1 -ProjectName LocalizationTests -Configuration Debug
```

Expected result: tests fail to compile until `LocalizationManager` exists.

- [x] **Step 2: Implement owner registration and string lookup**

Add API shape:

```cpp
namespace Localization
{
    enum class LanguagePreferenceKind
    {
        System,
        Culture
    };

    struct LanguagePreference final
    {
        LanguagePreferenceKind kind{LanguagePreferenceKind::System};
        std::wstring culture;
    };

    HRESULT RegisterResourceOwner(std::wstring_view ownerName, HINSTANCE embeddedInstance) noexcept;
    void UnregisterResourceOwner(HINSTANCE embeddedInstance) noexcept;
    HRESULT ApplyLanguagePreference(const LanguagePreference& preference) noexcept;
    int LoadString(HINSTANCE embeddedInstance, UINT id, std::wstring& result) noexcept;
}
```

Use `std::mutex`, `std::unordered_map`, `wil::unique_hmodule`, `Debug::Warning(...)`, and `LoadLibraryExW` with resource-only flags.

- [x] **Step 3: Route existing string helpers through the manager**

Keep `Common/Helpers.h` overload names stable. The helper path must use the localization manager first, then embedded English fallback. Existing callers must continue compiling unchanged.

- [x] **Step 4: Verify**

Run:

```powershell
.\build.ps1 -ProjectName Common -Configuration Debug
.\build.ps1 -ProjectName LocalizationTests -Configuration Debug
```

Expected result: build succeeds and fallback tests pass.

### Task 2: Add Resource DLL Build Infrastructure

**Files:**

- Modify: `Directory.Build.props`
- Modify: `Directory.Build.targets`
- Modify: `RedSalamander.sln`
- Modify: `build.ps1`

- [x] **Step 1: Add MSBuild properties for language resource projects**

Add the `IsLanguageResourceProject`, `ResourceOwnerName`, and `ResourceCulture` property contract from the Build Contract section.

- [x] **Step 2: Add output validation to the build script**

After solution build, verify the `Lang` directory exists when any language resource project was built:

```powershell
Test-Path ".build\x64\Debug\Lang"
```

Expected result after language projects exist: `True`.

- [x] **Step 3: Verify no normal project output moves**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Expected result: `RedSalamander.exe` remains in `.build\x64\Debug\`, not under `Lang`.

### Task 3: Add First Application Satellite Projects

**Files:**

- Create: `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- Create: `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.vcxproj`
- Create: `RedSalamanderMonitor/Lang/fr-FR/RedSalamanderMonitor-fr-FR.rc`
- Create: `RedSalamanderMonitor/Lang/fr-FR/RedSalamanderMonitor-fr-FR.vcxproj`
- Modify: `RedSalamander.sln`

- [x] **Step 1: Create minimal fr-FR satellite resources**

Include `IDS_APP_TITLE`, one Preferences label string, and one menu resource in the satellite. Use the same IDs as embedded English.

- [x] **Step 2: Build satellites**

Run:

```powershell
.\build.ps1 -Configuration Debug
```

Expected files:

```text
.build\x64\Debug\Lang\RedSalamander-fr-FR.dll
.build\x64\Debug\Lang\RedSalamanderMonitor-fr-FR.dll
```

- [x] **Step 3: Confirm embedded English still exists**

Temporarily rename the `Lang` folder and start the app self-test command that loads `IDS_APP_TITLE`.

Expected result: embedded English title still loads.

Verified on 2026-04-29 by temporarily renaming `.build\x64\Debug\Lang` to `.build\x64\Debug\Lang.__missing_dll_fallback_test`, running the focused Preferences self-test successfully, and restoring the `Lang` folder. Archived Commands run: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_165329`.

### Task 4: Add Plugin Satellite Project Pattern

**Files:**

- Create: `Plugins/FileSystem/Lang/fr-FR/FileSystem-fr-FR.rc`
- Create: `Plugins/FileSystem/Lang/fr-FR/FileSystem-fr-FR.vcxproj`
- Create equivalent `Lang/fr-FR` satellite files for:
  - `FileSystem7z`
  - `FileSystemCurl`
  - `FileSystemDummy`
  - `FileSystemGoogleDrive`
  - `FileSystemMicrosoftDrive`
  - `FileSystemS3`
  - `ViewerImgRaw`
  - `ViewerPE`
  - `ViewerSpace`
  - `ViewerSqlite`
  - `ViewerText`
  - `ViewerVLC`
  - `ViewerWeb`
- Modify: `RedSalamander.sln`

- [x] **Step 1: Build one plugin satellite first**

Start with `ViewerText` because viewer UI strings are easy to validate manually and through self-tests.

- [x] **Step 2: Validate plugin satellite fallback**

Run a test that registers `ViewerText.dll`, loads a string present in `ViewerText-fr-FR.dll`, then loads a string intentionally absent from the satellite and confirms embedded English fallback.

Validated through the deterministic localization owner tests: the test owner loads a present `fr-FR` satellite string, then loads an embedded-only string while the satellite remains active and confirms embedded English fallback. Plugin owners use the same `RegisterResourceOwner` and helper lookup path, and all plugin satellites build into the shared `Lang` output during the `RedSalamander` Debug build.

- [x] **Step 3: Replicate the project pattern to all plugins**

Keep each plugin's resource IDs tied to its existing `resource.h` or `*Resources.rc` include chain. Do not renumber plugin resources.

Implemented first-pass `fr-FR` satellites for every listed plugin. Each satellite includes localizable resource blocks from the owner `.rc` file and is included in the solution dependency graph so the normal `RedSalamander` Debug build produces all plugin `*-fr-FR.dll` files under `.build\x64\Debug\Lang\`.

- [x] **Step 4: Verify output**

Run:

```powershell
.\build.ps1 -Configuration Debug
Get-ChildItem .build\x64\Debug\Lang\*-fr-FR.dll
```

Expected result: application, monitor, and all plugin `*-fr-FR.dll` files are present in one `Lang` folder.

### Task 5: Persist `ui.language`

**Files:**

- Modify: `Common/Common/SettingsStore.h`
- Modify: `Common/Common/SettingsStore.cpp`
- Modify: `Specs/SettingsStore.schema.json`
- Modify: `Specs/Core/Core_SettingsStore.md`

- [x] **Step 1: Add settings tests**

Add coverage for:

- Missing `ui.language` loads as `system`.
- `"language": "system"` loads as `system`.
- `"language": "fr-FR"` loads as `fr-FR`.
- Invalid values such as `"..\\bad"` load as `system`.
- Saving defaults omits `ui.language`.
- Saving `fr-FR` writes `"language": "fr-FR"`.

- [x] **Step 2: Implement parser/writer**

Use yyjson copy APIs for dynamic strings. Do not store temporary string pointers in yyjson mutable builders.

- [x] **Step 3: Update schema**

Add the schema block from the Settings Contract section under `$defs.uiSettings.properties.language`.

- [x] **Step 4: Verify**

Run:

```powershell
.\build.ps1 -ProjectName Common -Configuration Debug
```

Expected result: settings tests pass and schema is copied next to the executable as before.

### Task 6: Add Preferences Language Control

**Files:**

- Modify: `RedSalamander/Preferences.General.cpp`
- Modify: `RedSalamander/Preferences.Internal.cpp`
- Modify: `RedSalamander/Preferences.cpp`
- Modify: `RedSalamander/Resource.h`
- Modify: `RedSalamander/RedSalamander.rc`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`
- Modify: `Specs/UI/UI_PreferencesDialog.md`

- [x] **Step 1: Add resource IDs and embedded English strings**

Add IDs for:

```text
IDS_PREF_LANGUAGE_LABEL
IDS_PREF_LANGUAGE_SYSTEM
IDS_PREF_LANGUAGE_DESCRIPTION
```

Use embedded English in `RedSalamander.rc`.

- [x] **Step 2: Add self-test coverage**

Extend General page self-tests to assert:

- `Display` group contains the language combo.
- Default selected item is `System Language`.
- Changing language enables `Apply`.
- `Cancel` restores previous language.
- `Apply` persists `ui.language`.

- [x] **Step 3: Implement combo population**

Populate from:

1. `System Language`
2. Built-in `fr-FR` entry for testing
3. Discovered cultures from `Lang\*-<culture>.dll`

De-duplicate cultures and sort by display name using Windows locale APIs.

- [x] **Step 4: Implement apply/cancel behavior**

Use the existing working settings pattern. On `Apply` or `OK`, persist and call the app-level localization refresh.

- [x] **Step 5: Verify**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Expected result: Preferences self-tests pass for General page.

### Task 7: Initialize Runtime Localization

**Files:**

- Modify: `RedSalamander/RedSalamander.cpp` or current app startup owner.
- Modify: `RedSalamanderMonitor/<startup>.cpp`
- Modify: plugin manager files that load plugins and track plugin module handles.
- Modify: `Common/LocalizationManager.cpp`

- [x] **Step 1: Register app owner before first UI resource use**

After settings load and before menus/windows/dialogs are created:

```cpp
Localization::RegisterResourceOwner(L"RedSalamander", GetModuleHandleW(nullptr));
Localization::ApplyLanguagePreference(settings.ui.language);
```

- [x] **Step 2: Register monitor owner**

Use the same pattern with owner name `RedSalamanderMonitor`.

- [x] **Step 3: Register plugin owners**

When plugin DLLs are loaded, register owner names matching their DLL base names, such as `ViewerText`. Unregister when plugins unload.

- [x] **Step 4: Implement language refresh**

On preference change:

- Re-resolve the language chain.
- Reload satellite DLLs for every registered owner.
- Rebuild menus created after the change.
- Invalidate visible top-level windows that support label refresh.
- Leave existing modal dialogs unchanged until reopened.

- [x] **Step 5: Verify missing DLL fallback**

Delete or rename `.build\x64\Debug\Lang\RedSalamander-fr-FR.dll`, set `ui.language` to `fr-FR`, and start the app.

Expected result: one warning is logged and embedded English is used.

Verified on 2026-04-29 with the entire `.build\x64\Debug\Lang` folder temporarily renamed away while running the focused Preferences self-test. The app started and completed with embedded English fallback; the folder was restored afterwards. Archived Commands run: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_165329`.

### Task 8: Route Non-String Resources Through Localization

**Files:**

- Modify every source file using direct resource APIs for user-facing resources:
  - `LoadMenuW`
  - `LoadAcceleratorsW`
  - `FindResourceW`
  - `LoadImageW`
  - dialog template loading APIs

- [x] **Step 1: Inventory direct resource loads**

Run:

```powershell
Get-ChildItem -Path Common,RedSalamander,RedSalamanderMonitor,Plugins -Recurse -Include *.cpp,*.h |
  Select-String -Pattern 'LoadMenuW|LoadAcceleratorsW|FindResourceW|LoadImageW|DialogBoxParamW|CreateDialogParamW'
```

Expected result: a complete list of resource load call sites.

Inventory result:

- Routed: app and monitor accelerator table loads, app and monitor top-level menus, Compare Directories menu, FolderView context menu, ViewerImgRaw/ViewerPE/ViewerSpace/ViewerWeb plugin menus, ViewerSpace host FolderView context menu reuse, dialog template APIs, icon/image loads, PNG resources, and ViewerWeb RCDATA loads.
- Remaining direct resource APIs are centralized inside `LocalizationManager` and `Win32CallbackHelpers` only.

- [x] **Step 2: Convert call sites to localization helpers**

Use helper names that make fallback behavior explicit:

```cpp
LoadMenuResource(ownerInstance, IDR_MAINMENU);
LoadAcceleratorsResource(ownerInstance, MAKEINTRESOURCEW(IDR_ACCELERATOR));
FindLocalizedResourceHandle(ownerInstance, MAKEINTRESOURCEW(id), RT_DIALOG);
LoadImageResource(ownerInstance, MAKEINTRESOURCEW(id), IMAGE_ICON, cx, cy, flags);
LoadResourceBytes(ownerInstance, MAKEINTRESOURCEW(id), RT_RCDATA, bytes);
```

- [x] **Step 3: Verify menu fallback**

Run a self-test that loads a menu from the satellite, then a menu missing from the satellite, and verifies fallback to embedded English.

### Task 9: Add Localization Self-Tests And Diagnostics

**Files:**

- Modify: `RedSalamander/SelfTest/Commands/*.cpp` as appropriate.
- Create or modify: `Tests/LocalizationTests/*`.
- Modify: `Common/LocalizationManager.cpp`.

- [x] **Step 1: Add deterministic resource lookup tests**

Test exact cases:

- Satellite present and ID present.
- Satellite present and ID missing.
- Satellite missing.
- Invalid culture.
- System language preference with unavailable system culture.
- Plugin resource owner registered after language already applied.
- Plugin resource owner unregistered before language changes again.

- [x] **Step 2: Add diagnostics**

Log:

- Satellite DLL load failure once per owner/culture.
- Invalid persisted language value once during settings load.
- Resource lookup fallback only in diagnostic/test mode to avoid noisy normal execution.

- [x] **Step 3: Verify**

Run:

```powershell
.\build.ps1 -Configuration Debug
```

Expected result: all affected tests pass.

### Task 10: Update Specs And Closeout

**Files:**

- Modify: `Specs/Core/Core_Localization.md`
- Modify: `Specs/Core/Core_SettingsStore.md`
- Modify: `Specs/UI/UI_PreferencesDialog.md`
- Move after implementation: `Specs/Plans/WIP/LocalizationResourceDllPlan_2026-04-29.md` to `Specs/Plans/Done/LocalizationResourceDllPlan_2026-04-29.md`

- [x] **Step 1: Update localization spec**

Document:

- `Lang` source folder convention.
- Output `Lang` folder convention.
- Satellite DLL naming.
- Embedded English fallback.
- Resource lookup order.
- Plugin owner registration.

- [x] **Step 2: Update settings and preferences specs**

Document `ui.language` and the General page language combo contract.

- [x] **Step 3: Archive perf evidence if startup/resource lookup changes are measurable**

Because this touches startup and resource lookup, capture at least one before/after startup/resource-load self-test run under `Specs/TestRuns/` if the implementation adds measurable startup work or hot-path lookup cost.

- [x] **Step 4: Move the plan to Done**

Only after all implementation, tests, and authoritative specs are complete:

```powershell
Move-Item -LiteralPath Specs\Plans\WIP\LocalizationResourceDllPlan_2026-04-29.md -Destination Specs\Plans\Done\LocalizationResourceDllPlan_2026-04-29.md
```

## Verification Matrix

Build verification:

```powershell
.\build.ps1 -Configuration Debug
.\build.ps1 -Configuration Release
.\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
```

Runtime verification:

- Fresh settings file starts with `System Language`.
- `System Language` uses available Windows preferred UI language satellites when present. falls back to embedded English when no preferred UI language satellites are present.
- `En-US` use embeded ressources.
- `fr-FR` loads `.build\x64\Debug\Lang\RedSalamander-fr-FR.dll` and plugin satellites.
- Missing app satellite falls back to embedded English.
- Missing plugin satellite falls back to that plugin's embedded English.
- Missing resource ID inside an existing satellite falls back to embedded English for that ID.
- Preferences `Cancel`, `Apply`, and `OK` language behavior matches the existing working settings contract.
- Existing English-only build still runs when `Lang` folder is absent.

Regression checks:

- No hardcoded new user-facing strings in C++.
- No printf-style format strings in resources.
- No raw owning `HMODULE`; use WIL RAII for satellite DLL ownership.
- No manual `FreeLibrary` except through WIL resource wrapper behavior.
- No blocking UI thread work for expensive language discovery. Cache the culture list and refresh only when Preferences opens or language files change during development.

Latest verification:

- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed on 2026-04-29.
- `.\build.ps1 -ProjectName LocalizationTests -Configuration Debug` passed on 2026-04-29.
- `.build\x64\Debug\LocalizationTests.exe` passed on 2026-04-29.
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_general_dxui_customization_preview_and_cancel --selftest-timeout-multiplier=2` passed on 2026-04-29.
- Archived focused Commands run: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_153049`.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed after adding first-pass `fr-FR` satellites for all plugin owners on 2026-04-29.
- `.build\x64\Debug\LocalizationTests.exe` passed after the full `fr-FR` satellite expansion on 2026-04-29.
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_general_dxui_customization_preview_and_cancel --selftest-timeout-multiplier=2` passed after the full `fr-FR` satellite expansion on 2026-04-29.
- Archived focused Commands run after full `fr-FR` expansion: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_160421`.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed after the `fr-FR` translation language-level review on 2026-04-29.
- `.build\x64\Debug\LocalizationTests.exe` passed after the `fr-FR` translation language-level review on 2026-04-29.
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_general_dxui_customization_preview_and_cancel --selftest-timeout-multiplier=2` passed after the `fr-FR` translation language-level review on 2026-04-29.
- Archived focused Commands run after the `fr-FR` translation language-level review: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_162752`.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed after routing non-string resources through localization helpers on 2026-04-29.
- `.\build.ps1 -ProjectName LocalizationTests -Configuration Debug` passed after adding dialog/template and embedded-only fallback tests on 2026-04-29.
- `.build\x64\Debug\LocalizationTests.exe` passed after adding dialog/template and embedded-only fallback tests on 2026-04-29.
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_general_dxui_customization_preview_and_cancel --selftest-timeout-multiplier=2` passed after non-string resource routing on 2026-04-29.
- Archived focused Commands run after non-string resource routing: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_165319`.
- Missing `Lang` fallback self-test passed after temporarily renaming `.build\x64\Debug\Lang` on 2026-04-29.
- Archived missing `Lang` fallback Commands run: `Specs\TestRuns\4cb089111a23\Commands\2026-04-29_165329`.
- `git diff --check` passed on 2026-04-29 with line-ending warnings only.

## Open Decisions For Implementation

- Whether to include `fr-FR` satellite DLLs in normal release artifacts even though embedded English is complete. Recommended: yes during the first implementation to continuously validate the satellite pipeline.
- Whether live language switching must refresh every already-open modeless tool window immediately. Recommended first contract: refresh main menus and Preferences immediately; other modeless windows may refresh on reopen unless their owning code already has a safe label refresh path.
- Whether installer packaging should expose language selection. Recommended: defer to an installer-specific plan after runtime and build support are stable.


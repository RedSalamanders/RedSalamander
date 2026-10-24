# RedConfigure Localization Review Workbench Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace RedConfigure's single-owner/single-culture localization editor with one review workbench that can show every application/plugin owner and multiple target languages at the same time, with English as the read-only master.

**Architecture:** Keep workspace discovery as the source of owners and satellite `.rc` files, but introduce a localization review model keyed by `(owner, resource ID)` with one read-only English source cell and zero or more editable target-language cells. The UI projects that model through owner and language check states plus text/ID/status filters; export writes one satellite `.rc` per changed owner/culture.

**Tech Stack:** C++23, Win32/DxUi, RedConfigure `.rc` resources, `RedConfigureTests`, MSBuild via `build.ps1`, existing RC parser/writer and placeholder validation.

---

## User Contract

- English embedded `.rc` resources are the master source and are never editable in RedConfigure.
- All discovered first-party resource owners can be reviewed in one localization grid: applications and plugins together.
- Owners are selectable with check-style controls, including an all-owners state. Unchecked owners are hidden from the review grid and export preview.
- Target languages are selectable with check-style controls, including an all-target-languages state. Multiple checked target languages are visible at the same time.
- Existing target languages come from `Lang\<culture>\` satellite resource paths discovered across all owners. New target cultures can still be chosen from official Windows locale names and remain in memory until export.
- Search matches owner, ID, English source text, and visible target-language text.
- ID filter narrows resource IDs.
- Status filter supports `All`, `OK`, and `Problems`. A row has problems when any visible target-language cell fails placeholder validation.
- Editing is available only for checked non-English target languages. English cells and unsupported resource inventory entries are read-only.
- The focused editor shows the selected owner, resource ID, English source, and a target editor for the selected non-English language cell.
- Export writes only target-language satellite `.rc` files. It never rewrites embedded English source files.
- Export must preserve deterministic ordering and must block invalid placeholder edits.

## Protected Performance Scenario

Scenario: `redconfigure.localization.review.all_owners_all_languages`.

Why protected: the localization workbench can project every discovered app/plugin owner and several languages at once. Typing in search, toggling owners/languages, and editing a target cell must stay responsive with thousands of rows.

Metrics/evidence:

- Correctness: `RedConfigureTests` must include deterministic fixtures with at least two owners and at least two target cultures.
- Filter performance: add an in-process deterministic measurement in `RedConfigureTests` for building the review projection over a synthetic multi-owner/multi-language data set. Record elapsed time and visible row count to stdout for archived runs.
- Archive: when the full feature closes, run the targeted test and place the stdout/log under `Specs/TestRuns/<MachineHash>/RedConfigure/LocalizationReview/<timestamp>/`.
- Completion caveat: until a checked-in perf archive exists, the feature can be marked implemented but not perf-closed.

## File Structure

- Modify `RedConfigure/RedConfigureSession.h`
  - Add review-row/cell types for owner + source + per-culture targets.
  - Add owner/language check-state APIs.
  - Keep existing single-owner APIs temporarily where the current UI still uses them during migration.
- Modify `RedConfigure/RedConfigureSession.cpp`
  - Parse all discovered owners during workspace load.
  - Build and cache review rows for every source `STRINGTABLE` entry.
  - Merge each owner/culture target from satellite `.rc` files.
  - Update target cells by `(rowIndex, cultureName)`.
  - Build export text per owner/culture.
- Modify `RedConfigure/RedConfigureGridModels.h`
  - Replace the translation-grid projection with a dynamic localization review grid: columns `Owner`, `ID`, `English`, one column per checked target culture, and `Status`.
  - Keep stable row IDs based on owner/source ID rather than transient vector offsets.
- Modify `RedConfigure/RedConfigureRoot.cpp`
  - Replace the single culture combo and active owner combo in the localization page with check-list style owner/language surfaces.
  - Update selection/edit flow to track selected row plus selected target culture.
  - Update scope/count labels for multi-owner/multi-language review.
- Modify `RedConfigure/RedConfigure.rc`, `RedConfigure/resource.h`, and `RedConfigure/Lang/*/RedConfigure-*.rc`
  - Add localized labels for `Languages`, `Owners`, `All owners`, `All languages`, `English source`, `Target language`, `Selected cell`, `Export changed RC files`, and empty-state copy for filtered multi-owner review.
- Modify `Specs/Core/Core_RedConfigure.md`
  - Make multi-owner/multi-language localization review normative.
  - Define English-read-only and target-culture export behavior.
- Modify `Specs/UI/UI_RedConfigure.md`
  - Replace the single culture/active owner localization contract with checkable owners/languages and dynamic target columns.
- Modify `Tests/RedConfigureTests/RedConfigureTests.cpp`
  - Add failing tests before implementation for review row assembly, owner/language filtering, editing a target language, and multi-file export preview.

## Data Model

Add these public types in `RedConfigureSession.h`:

```cpp
struct LocalizationTargetCell
{
    std::wstring cultureName;
    std::wstring targetText;
    Localization::PlaceholderValidationResult validation;
    bool hasExistingTranslation = false;
    bool dirty = false;
};

struct LocalizationReviewRow
{
    std::wstring ownerName;
    std::wstring id;
    std::wstring sourceText;
    std::vector<LocalizationTargetCell> targets;
};

struct LocalizationReviewSelection
{
    size_t rowIndex = 0u;
    std::wstring cultureName;
};

struct LocalizationReviewViewOptions
{
    std::wstring searchText;
    std::wstring idFilterText;
    std::vector<std::wstring> visibleOwnerNames;
    std::vector<std::wstring> visibleCultureNames;
    LocalizationStatusFilter statusFilter = LocalizationStatusFilter::All;
    LocalizationViewColumn sortColumn = LocalizationViewColumn::Id;
    std::wstring sortCultureName;
    LocalizationSortDirection sortDirection = LocalizationSortDirection::None;
};
```

Keep English out of `targets`; it is represented by `sourceText` and always read-only.

## Task 1: Add Multi-Owner/Multi-Language Model Tests

**Files:**
- Modify: `Tests/RedConfigureTests/RedConfigureTests.cpp`

- [x] **Step 1: Write the failing review assembly test**

Add a test fixture that creates:

- `App/App.rc` with `IDS_HELLO` and `IDS_APP_ONLY`
- `App/Lang/fr-FR/App-fr-FR.rc` with `IDS_HELLO`
- `App/Lang/cs-CZ/App-cs-CZ.rc` with `IDS_HELLO`
- `Plugin/Plugin.rc` with `IDS_HELLO` and `IDS_PLUGIN_ONLY`
- `Plugin/Lang/fr-FR/Plugin-fr-FR.rc` with `IDS_PLUGIN_ONLY`

Expected assertions:

```cpp
RedConfigure::RedConfigureSession session;
ok = Require(SUCCEEDED(session.LoadWorkspace(tempRoot, L"fr-FR")), L"Session should load the review fixture.") && ok;
ok = Require(session.GetLocalizationReviewRows().size() == 4u, L"Review rows should include all string IDs from all owners.") && ok;
ok = Require(ContainsCulture(session.GetLocalizationReviewCultures(), L"fr-FR"), L"Review cultures should include fr-FR.") && ok;
ok = Require(ContainsCulture(session.GetLocalizationReviewCultures(), L"cs-CZ"), L"Review cultures should include cs-CZ.") && ok;
ok = Require(FindReviewRow(session, L"App", L"IDS_APP_ONLY") != nullptr, L"Review rows should preserve app-only IDs.") && ok;
ok = Require(FindReviewRow(session, L"Plugin", L"IDS_PLUGIN_ONLY") != nullptr, L"Review rows should preserve plugin-only IDs.") && ok;
```

- [x] **Step 2: Run the test and capture RED**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
```

Expected: compile failure because `GetLocalizationReviewRows()` and `GetLocalizationReviewCultures()` do not exist yet.

## Task 2: Implement Review Row Assembly

**Files:**
- Modify: `RedConfigure/RedConfigureSession.h`
- Modify: `RedConfigure/RedConfigureSession.cpp`

- [x] **Step 1: Add model storage and public getters**

Add:

```cpp
[[nodiscard]] std::span<const LocalizationReviewRow> GetLocalizationReviewRows() const noexcept;
[[nodiscard]] std::span<const std::wstring> GetLocalizationReviewCultures() const noexcept;
```

Store:

```cpp
std::vector<LocalizationReviewRow> _localizationReviewRows;
std::vector<std::wstring> _localizationReviewCultures;
```

- [x] **Step 2: Build review rows during `LoadWorkspace`**

After workspace and theme load, call a new private method:

```cpp
[[nodiscard]] HRESULT LoadLocalizationReview();
```

`LoadLocalizationReview()` reads each `ResourceOwner::embeddedResourcePath`, parses source strings, discovers all culture names from every owner satellite path, reads matching satellite files per owner/culture, and builds one row per source string.

- [x] **Step 3: Keep existing active-owner API working**

After `LoadLocalizationReview()` succeeds, call the current `LoadLocalizationForActiveOwner()` so the current UI remains usable while the review UI is migrated.

- [x] **Step 4: Run GREEN verification**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
.\.build\x64\Debug\RedConfigureTests.exe
```

Expected: new review assembly test passes, and existing single-owner tests still pass.

Implementation-start evidence from 2026-06-08:

- RED: direct MSBuild failed on missing `LocalizationReviewRow`, `LocalizationTargetCell`, and review getters.
- GREEN: direct MSBuild with `BuildProjectReferences=false` built `RedConfigureTests` with 0 warnings/0 errors.
- GREEN: `.\.build\x64\Debug\RedConfigureTests.exe` printed `RedConfigureTests passed.`
- Note: `.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug` was blocked before compile because a running Debug `RedConfigure.exe` held `.build\x64\Debug\Common.dll` open.

## Task 3: Add Review Projection Tests

**Files:**
- Modify: `Tests/RedConfigureTests/RedConfigureTests.cpp`
- Modify: `RedConfigure/RedConfigureSession.h`
- Modify: `RedConfigure/RedConfigureSession.cpp`

- [x] **Step 1: Write failing projection tests**

Add tests for:

- owner filtering hides unchecked owners
- culture filtering hides unchecked target-language text from search/status evaluation
- search matches owner, ID, English source, and visible target text
- problem status is set when any visible target language has placeholder mismatch
- sorting by a visible target culture orders by that culture's target text

Expected API:

```cpp
std::vector<size_t> view = RedConfigure::BuildLocalizationReviewView(session.GetLocalizationReviewRows(), options);
```

- [x] **Step 2: Implement `BuildLocalizationReviewView`**

Use the existing `BuildTranslationView` patterns:

- ASCII-insensitive substring matching can reuse the existing helper.
- Status checks must iterate only visible target cultures.
- Stable source order is the default.
- Sort ties by owner then ID then original index.

- [x] **Step 3: Run GREEN verification**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
.\.build\x64\Debug\RedConfigureTests.exe
```

Expected: projection tests pass.

Implementation-start evidence from 2026-06-08:

- RED: direct MSBuild failed on missing `LocalizationReviewViewOptions` and `BuildLocalizationReviewView`.
- GREEN: direct MSBuild with `BuildProjectReferences=false` built `RedConfigureTests` with 0 warnings/0 errors.
- GREEN: `.\.build\x64\Debug\RedConfigureTests.exe` printed `RedConfigureTests passed.`

## Task 4: Add Editing and Export Tests

**Files:**
- Modify: `Tests/RedConfigureTests/RedConfigureTests.cpp`
- Modify: `RedConfigure/RedConfigureSession.h`
- Modify: `RedConfigure/RedConfigureSession.cpp`

- [x] **Step 1: Write failing target edit test**

Expected API:

```cpp
ok = Require(session.UpdateLocalizationReviewTarget(rowIndex, L"cs-CZ", L"Ahoj {0}"), L"Review target edit should accept valid placeholders.") && ok;
ok = Require(! session.UpdateLocalizationReviewTarget(rowIndex, L"cs-CZ", L"Ahoj"), L"Review target edit should reject placeholder mismatches.") && ok;
```

- [x] **Step 2: Write failing per-owner/per-culture export preview test**

Expected API:

```cpp
std::vector<RedConfigure::LocalizationExportPreview> previews;
ok = Require(SUCCEEDED(session.BuildLocalizationReviewExportPreviews(previews)), L"Review export previews should build.") && ok;
ok = Require(ContainsPreview(previews, L"App", L"cs-CZ", L"Ahoj {0}"), L"Edited cs-CZ app preview should exist.") && ok;
ok = Require(ContainsPreview(previews, L"Plugin", L"fr-FR", L"Plugin seul"), L"Existing plugin fr-FR preview should exist.") && ok;
```

- [x] **Step 3: Implement edit and export APIs**

Add:

```cpp
[[nodiscard]] bool UpdateLocalizationReviewTarget(size_t rowIndex, std::wstring_view cultureName, std::wstring_view targetText);
[[nodiscard]] HRESULT BuildLocalizationReviewExportPreviews(std::vector<LocalizationExportPreview>& outPreviews) const;
[[nodiscard]] HRESULT ExportLocalizationReview(const std::filesystem::path& outputRoot) const;
```

`LocalizationExportPreview` contains owner name, culture, path, and generated RC text. Paths follow:

```text
<workspace>\RedConfigureOutput\<Owner>-<Culture>.rc
```

- [x] **Step 4: Run GREEN verification**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
.\.build\x64\Debug\RedConfigureTests.exe
```

Expected: edit/export tests pass.

Implementation-start evidence from 2026-06-08:

- RED: direct MSBuild failed on missing `LocalizationExportPreview`, `UpdateLocalizationReviewTarget`, and `BuildLocalizationReviewExportPreviews`.
- GREEN: direct MSBuild with `BuildProjectReferences=false` built `RedConfigureTests` with 0 warnings/0 errors.
- GREEN: `.\.build\x64\Debug\RedConfigureTests.exe` printed `RedConfigureTests passed.`
- Compile-only app check: direct MSBuild `RedConfigure\RedConfigure.vcxproj /t:ClCompile /p:BuildProjectReferences=false` completed with 0 warnings/0 errors.

## Task 5: Build Dynamic Review Grid

**Files:**
- Modify: `RedConfigure/RedConfigureGridModels.h`
- Modify: `RedConfigure/RedConfigureRoot.cpp`

- [x] **Step 1: Add `LocalizationReviewGridModel`**

Columns:

1. `owner`
2. `id`
3. `english`
4. one column per visible target culture
5. `status`

Cell behavior:

- English column uses `row.sourceText`.
- Target culture columns use `LocalizationTargetCell::targetText`, falling back to English source text only for display when a target is absent.
- Status column shows `OK` only when all visible target cells are valid.
- Problem rows use `GridRowTone::Warning`.

- [x] **Step 2: Preserve stable selection**

Stable row ID should hash owner + ID deterministically with no persisted ABI guarantee. A simple FNV-1a over UTF-16 code units is sufficient for in-memory grid selection.

- [x] **Step 3: Wire grid projection**

Replace `_translationModel` on the localization page with `_localizationReviewModel`. Keep old `_translationModel` members until export preview migration is complete, then remove.

- [x] **Step 4: Run build**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigure -Configuration Debug
```

Expected: RedConfigure builds.

Implementation evidence from 2026-06-08:

- Added `LocalizationReviewGridModel` with owner, ID, English, dynamic target-language, and status columns.
- Added deterministic owner+ID FNV-1a stable row IDs.
- Wired the localization page to `_localizationReviewModel` and `BuildLocalizationReviewView(...)`.
- GREEN: `.\build.ps1 -ProjectName RedConfigure -Configuration Debug` completed with 0 warnings/0 errors.

## Task 6: Owner and Language Check Controls

**Files:**
- Modify: `RedConfigure/RedConfigureRoot.cpp`
- Modify: `RedConfigure/RedConfigure.rc`
- Modify: `RedConfigure/resource.h`
- Modify: `RedConfigure/Lang/cs-CZ/RedConfigure-cs-CZ.rc`
- Modify: `RedConfigure/Lang/ja-JP/RedConfigure-ja-JP.rc`
- Modify: `RedConfigure/Lang/sk-SK/RedConfigure-sk-SK.rc`

- [x] **Step 1: Add resource strings**

Add source strings with positional placeholders where needed:

```rc
IDS_REDCONFIGURE_LABEL_OWNERS "Owners"
IDS_REDCONFIGURE_LABEL_LANGUAGES "Languages"
IDS_REDCONFIGURE_TOGGLE_ALL_OWNERS "All owners"
IDS_REDCONFIGURE_TOGGLE_ALL_LANGUAGES "All languages"
IDS_REDCONFIGURE_LABEL_ENGLISH_SOURCE "English source"
IDS_REDCONFIGURE_LABEL_TARGET_LANGUAGE "Target language"
IDS_REDCONFIGURE_LOCALIZATION_REVIEW_EMPTY "No localization rows match the current owner, language, and text filters."
IDS_REDCONFIGURE_STATUS_REVIEW_EXPORT_DONE "{0} RC file(s) exported"
```

- [x] **Step 2: Implement owner/language toggles**

Use existing DxUi controls. If DxUi does not provide a compact checkbox list, use a grid with checkbox columns because the grid already supports checkbox cell toggling.

- [x] **Step 3: Update layout**

Desktop layout:

- first row: owners selector region and languages selector region
- second row: search, ID, status filters
- main band: review grid
- editor band: English source and selected target language editor

Compact layout:

- owners, languages, search, ID, and status stack into reserved rows
- grid and editor keep minimum heights and do not overlap the status line

- [x] **Step 4: Run build**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigure -Configuration Debug
```

Expected: RedConfigure builds.

Implementation evidence from 2026-06-08:

- Added localized owner/language/all-state/source/target/export/empty/status strings to English, cs-CZ, ja-JP, and sk-SK RedConfigure resources.
- Note: this repo does not contain `RedConfigure/Lang/fr-FR/RedConfigure-fr-FR.rc`, so no fr-FR RedConfigure satellite was updated.
- Added `LocalizationFilterGridModel` checkbox grids for owners and target languages.
- Added the target-language picker path for creating a new in-memory culture from official Windows locale names.
- Updated the desktop and compact localization layout to show owner/language check grids, search/ID/status filters, dynamic review grid, English source, and selected target-language editor.
- GREEN: `.\build.ps1 -ProjectName RedConfigure -Configuration Debug` completed with 0 warnings/0 errors.

## Task 7: Export Review Files from UI

**Files:**
- Modify: `RedConfigure/RedConfigureRoot.cpp`

- [x] **Step 1: Update review/export preview**

Review & Export should list generated `.rc` previews for every visible or dirty owner/culture pair. Keep theme preview unchanged.

- [x] **Step 2: Update export button**

`Export RC` calls `ExportLocalizationReview(...)` and reports the number of written files.

- [x] **Step 3: Preserve old single-owner export only as fallback**

Implementation evidence from 2026-06-08:

- Added `RedConfigureSession::ExportLocalizationReview(...)`.
- Review & Export concatenates generated per-owner/per-culture `.rc` previews for visible owner/language pairs and dirty hidden files.
- `Export RC` writes the multi-file review export set and reports the count with positional resource formatting.
- Empty review rows still fall back to the legacy active-owner export path.

If no review rows exist, fall back to the current `ExportLocalization(...)` behavior so partial/broken workspaces can still export the active owner.

## Task 8: Specs and Closeout

**Files:**
- Modify: `Specs/Core/Core_RedConfigure.md`
- Modify: `Specs/UI/UI_RedConfigure.md`
- Move on closeout: `Specs/Plans/WIP/RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md` to `Specs/Plans/Done/RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md`

- [x] **Step 1: Update authoritative specs**

Document:

- English embedded resources are read-only master strings.
- Localization review can show multiple owners and multiple target languages at once.
- Owner and language filters are checkable and support all-states.
- Target edits validate placeholders immediately.
- Export writes satellite `.rc` files per owner/culture.

- [x] **Step 2: Run validation**

Run:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
.\.build\x64\Debug\RedConfigureTests.exe
.\build.ps1 -ProjectName RedConfigure -Configuration Debug
```

- [x] **Step 3: Archive perf evidence**

Create a run directory under `Specs/TestRuns/<MachineHash>/RedConfigure/LocalizationReview/<timestamp>/` containing:

- `redconfigure-tests-stdout.txt`
- `redconfigure-tests-stderr.txt`
- `build-log-path.txt`
- `localization-review-metrics.txt`

- [x] **Step 4: Move the completed plan**

Only after tests and spec updates are complete:

```powershell
Move-Item -LiteralPath Specs\Plans\WIP\RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md -Destination Specs\Plans\Done\RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md
```

Closeout evidence from 2026-06-08:

- GREEN: `.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug` completed with 0 warnings/0 errors.
- GREEN: `.\.build\x64\Debug\RedConfigureTests.exe` completed with `RedConfigureTests passed.`
- GREEN: `.\build.ps1 -ProjectName RedConfigure -Configuration Debug` completed with 0 warnings/0 errors.
- Archived performance evidence: `Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926/`.
- Metric: `redconfigure.localization.review.all_owners_all_languages rows=4800 visibleRows=4800 cultures=4 elapsedMicros=4852`.
- Additional resource-contract check: `Tools\Tests\ResourceLocalizationContracts.Tests.ps1` passed the positional-placeholder safety check and reported unrelated pre-existing language-neutral resource violations in `RedSalamander/Lang/*` plus localized helper uses in `FolderWindow.FileOperations.Popup.cpp`.

## Self-Review

- Spec coverage: user-visible owner/language multi-select, English read-only, editable target languages, search/ID/status filtering, multi-language display, app/plugin scope, export behavior, validation, and performance evidence are covered.
- Placeholder scan: no `TBD`, `TODO`, or deferred implementation placeholders remain.
- Type consistency: public types consistently use `LocalizationReview*`; existing `TranslationEntry` APIs remain for migration compatibility until the UI is fully switched.

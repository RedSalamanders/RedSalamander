# Advisor Plan 006 - SettingsSchemaParser unit tests

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/006-settings-schema-parser-tests.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/006-settings-schema-parser-tests.md`
- **Live owner:** none (implemented as Tests/SettingsSchemaTests)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 006: Unit-test SettingsSchemaParser against the real settings schema and malformed inputs

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat 551e35655..HEAD -- RedSalamander/SettingsSchemaParser.h RedSalamander/SettingsSchemaParser.cpp Specs/SettingsStore.schema.json`
> If any of those three files changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.
> The scope is deliberately limited to the files this plan quotes. `Tests\` and
> `RedSalamander.sln` are modification *targets* (every new test project changes
> them), not drift-sensitive excerpts — including them here would false-trip the
> check every time.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW (new test code; possible small build-wiring change to reuse the parser source)
- **Depends on**: none
- **Category**: tests
- **Planned at**: commit `d6bfccc42`, 2026-06-11; **re-verified & tightened against `551e35655`, 2026-06-21** (parser source byte-identical to plan time; scaffolding facts re-confirmed against the now-merged plan-004 test project)

## Why this matters

`SettingsSchemaParser` turns `Specs/SettingsStore.schema.json` into the field model that drives the Preferences UI and settings handling. It is pure string-in/struct-out logic — the cheapest kind of code to unit-test — yet has zero direct tests; it is exercised only indirectly through UI-driven Commands selftests. A parsing regression (a field silently dropped, min/max misread, enum lost) corrupts how user settings are presented and persisted, and today nothing would catch it. The pure API makes this a high-leverage, low-risk test plan.

## Current state

- `RedSalamander\SettingsSchemaParser.h` (2,001 bytes, read at plan time) — the full public API:
  ```cpp
  namespace SettingsSchemaParser
  {
  struct SettingField
  {
      std::wstring jsonPath;      // "mainMenu.menuBarVisible"
      std::wstring paneName;      // "General", "Advanced", "Keyboard", ...
      std::wstring title;
      std::wstring description;
      std::wstring controlType;   // "toggle", "edit", "number", "combo", "custom"
      std::wstring sectionHeader;
      int displayOrder = 0;
      std::wstring schemaType;    // "boolean", "string", "integer", "number", "array", "object"
      bool hasMin = false; bool hasMax = false;
      int64_t minValue = 0; int64_t maxValue = 0;
      std::vector<std::wstring> enumValues;
      std::wstring defaultValue;
  };
  [[nodiscard]] std::vector<SettingField> ParseSettingsSchema(std::string_view schemaJsonUtf8) noexcept;
  [[nodiscard]] std::vector<SettingField> LoadAndParseSettingsSchema(std::wstring_view schemaFilePath) noexcept;
  [[nodiscard]] std::vector<SettingField> GetFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept;
  [[nodiscard]] std::vector<SettingField> GetNonCustomFieldsForPane(const std::vector<SettingField>& allFields, std::wstring_view paneName) noexcept;
  }
  ```
- `RedSalamander\SettingsSchemaParser.cpp` (~12.7 KB) — implementation (yyjson-based). Two facts that shape the tests and the build wiring:
  - **It uses yyjson** (`#include <yyjson.h>`, calls `yyjson_read`/`yyjson_obj_get`/…). The test exe therefore **must link `yyjson.lib`** (see "Build wiring" below) — this is the #1 thing a naive scaffold gets wrong.
  - **It includes the app-local `"Framework.h"`** (`RedSalamander\Framework.h`), which in turn includes `"targetver.h"` (also in `RedSalamander\`) and `"Win32CallbackHelpers.h"` (in **`Common\`**). So the test project needs **both `RedSalamander\` and `Common\`** on its include path. `Framework.h` also declares `extern PCWSTR REDSALAMANDER_TEXT_VERSION;` — the parser never references it, so there is **no extra link dependency** from that.
  - It walks **both** the top-level `properties` tree **and** the `$defs` section. A `$defs`-level definition that carries `x-ui-pane` defaults `controlType` to **`"custom"`** (`SettingsSchemaParser.cpp:261`), whereas a `properties`-level field defaults to **`"edit"`** (`SettingsSchemaParser.cpp:140`). This is why `GetNonCustomFieldsForPane` has something to exclude, and why a field's default `controlType` depends on where it sits — keep it in mind for Steps 2.5 and 3.
- The real schema: `Specs/SettingsStore.schema.json` (path verified to exist at plan time via `Glob Specs/**/*.schema.json` → exactly `Specs\SettingsStore.schema.json`). It uses `x-ui-pane`/`x-ui-control`/`x-ui-section`/`x-ui-order` attributes and has a `$defs` section (verified at `:54`).
- The parser source lives in the **RedSalamander app project**, not Common — a standalone test exe must compile `..\..\RedSalamander\SettingsSchemaParser.cpp` into itself. **The proven template is plan 004's `Tests\PluginContractTests\PluginContractTests.vcxproj`** (now merged at HEAD): it is a console `Application`, uses **no PCH**, compiles an *external* source file directly (`..\..\Common\FileSystemPathIdentity.cpp`), and — critically — already shows the exact yyjson link block. Copy THAT project, not `PerformanceTests2` (which uses a PCH + a `main=__ignored_*` entry-point-rename trick that this plan does **not** need, because `SettingsSchemaParser.cpp` defines no `main`/`WinMain`).
- **Build wiring you inherit for free** from `Tests\Directory.Build.props` (read it — `:8-26`): any first-party `.vcxproj` under `Tests\` automatically gets Console subsystem, the vcpkg include dir (`$(VcpkgInstalledDir)$(RSVcpkgTriplet)\include` via `ExternalIncludePath`), and **`$(SolutionDir)Common` on the include path** (`RSTestAdditionalIncludeDirectories` default). It does **NOT** add `RedSalamander\`, and it does **NOT** link yyjson — those two are the only additions you must make (Step 1).
- Test-exe output pattern: `Tests\LocalizationTests\LocalizationTests.cpp` (and `PluginContractTests.cpp`) use a `Check(bool condition, const wchar_t* message, bool& success)` helper that prints `[       OK ]` on pass and `[ FAILED  ]` on fail (exact bracket spacing — `LocalizationTests.cpp:33-43`), and `wmain` returns 0/1. Reuse that exact helper shape so the runner's pass/fail grep matches the other suites.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build | `.\build.ps1` | exit 0 |
| Run new tests | `.\.build\x64\Debug\SettingsSchemaTests.exe` | exit 0; `[       OK ]` lines, no `[ FAILED  ]` |
| Schema exists | `Test-Path Specs\SettingsStore.schema.json` | `True` |
| yyjson lib present (debug) | `Get-ChildItem -Recurse -Filter yyjson.lib vcpkg_installed` (run from repo root; or just trust that `Tests\PluginContractTests` already links it and builds) | at least one `yyjson.lib` under a `*\debug\lib\` path |

(Note: the exact pass/fail strings are `[       OK ]` and `[ FAILED  ]` with that bracket spacing — copy the `Check` helper verbatim from `LocalizationTests.cpp:33-43` so they match.)

## Scope

**In scope**:
- `Tests\SettingsSchemaTests\` (create: `SettingsSchemaTests.cpp`, `SettingsSchemaTests.vcxproj`)
- `RedSalamander.sln` (add project)
- `Tools\TestRunPlan.ps1` (OPTIONAL — add the new exe to the `-Suite Full` plan; see Step 4). There is no `test-all.ps1`; the local runner is `Tools\Run-AllTests.ps1`.

**Out of scope** (do NOT touch):
- `SettingsSchemaParser.h/.cpp` themselves — if a test exposes a real bug, report it; do not fix parser behavior in this plan (characterize current behavior instead and mark the check with a `// documents current behavior, suspected bug:` comment).
- `SettingsStore.cpp` / registry persistence — registry-level tests were considered and deferred (the store isn't obviously parameterizable by root key; see plans/README.md).
- `Specs/SettingsStore.schema.json` — read-only fixture.

## Git workflow

- Branch: `advisor/006-settings-schema-tests`
- Subject style: "Add SettingsSchemaParser unit tests".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Scaffold the test project

Copy **`Tests\PluginContractTests\PluginContractTests.vcxproj`** to `Tests\SettingsSchemaTests\SettingsSchemaTests.vcxproj` (it is the cleanest in-repo template: console `Application`, no PCH, already compiles an external `.cpp` directly and already links yyjson). Then make exactly these edits:

1. **New GUID + name**: generate a fresh `<ProjectGuid>` (PowerShell: `[guid]::NewGuid().ToString('B').ToUpper()`), and set `<RootNamespace>SettingsSchemaTests</RootNamespace>`. Do not reuse PluginContractTests' GUID `{AC2ECB20-389B-4B35-A717-3E63A4FA42A0}` — a duplicate GUID breaks the solution load.

2. **Sources**: replace the `<ClCompile>` items with your test file plus the parser source compiled directly:
   ```xml
   <ItemGroup>
     <ClCompile Include="SettingsSchemaTests.cpp" />
     <!-- Parser lives in the RedSalamander app project; compile it directly so the
          test exe has no dependency on the app binary. -->
     <ClCompile Include="..\..\RedSalamander\SettingsSchemaParser.cpp" />
   </ItemGroup>
   ```

3. **Include path — add `RedSalamander\`** (Common is already provided by `Tests\Directory.Build.props`; `RedSalamander\` is NOT, and `SettingsSchemaParser.cpp` includes `Framework.h`/`targetver.h` from there). Add to the existing `<ItemDefinitionGroup><ClCompile>`:
   ```xml
   <AdditionalIncludeDirectories>$(SolutionDir)RedSalamander;$(SolutionDir)Common;%(AdditionalIncludeDirectories)</AdditionalIncludeDirectories>
   ```

4. **Link yyjson** — keep the two `<ItemDefinitionGroup>` link blocks that PluginContractTests already has (they are the load-bearing part most likely to be dropped):
   ```xml
   <!-- Debug and ASan Debug: vcpkg debug lib -->
   <ItemDefinitionGroup Condition="'$(Configuration)'=='Debug' or '$(Configuration)'=='ASan Debug'">
     <Link>
       <AdditionalDependencies>yyjson.lib;%(AdditionalDependencies)</AdditionalDependencies>
       <AdditionalLibraryDirectories>$(VcpkgInstalledDir)$(RSVcpkgTriplet)\debug\lib;%(AdditionalLibraryDirectories)</AdditionalLibraryDirectories>
     </Link>
   </ItemDefinitionGroup>
   <!-- Release: vcpkg release lib -->
   <ItemDefinitionGroup Condition="'$(Configuration)'=='Release'">
     <Link>
       <AdditionalDependencies>yyjson.lib;%(AdditionalDependencies)</AdditionalDependencies>
       <AdditionalLibraryDirectories>$(VcpkgInstalledDir)$(RSVcpkgTriplet)\lib;%(AdditionalLibraryDirectories)</AdditionalLibraryDirectories>
     </Link>
   </ItemDefinitionGroup>
   ```
   Keep the `<ProjectReference>` to `..\..\Common\Common\Common.vcxproj` that PluginContractTests has (harmless and matches the template; the parser itself needs no Common symbols, but it keeps the include/build environment consistent).

5. **Add to the solution**: register `SettingsSchemaTests.vcxproj` in `RedSalamander.sln` with the standard `Project(...)`/`EndProject` block and per-config entries in `GlobalSection(ProjectConfigurationPlatforms)`. Easiest reliable approach: copy the two PluginContractTests blocks in the `.sln` and substitute the new project name + new GUID.

6. **Minimal `SettingsSchemaTests.cpp`** for this step: include `<iostream>` and the `Check` helper copied verbatim from `LocalizationTests.cpp:33-43`, plus `#include "SettingsSchemaParser.h"`, and a `wmain` that calls `SettingsSchemaParser::ParseSettingsSchema("{}")` once (proves the parser links) and returns 0.

**Verify**: `.\build.ps1` → exit 0; `.\.build\x64\Debug\SettingsSchemaTests.exe` exists and exits 0. If the link fails with unresolved `yyjson_*` externals, item 4 is wrong/missing; if it fails to find `Framework.h`/`targetver.h`, item 3 is wrong.

### Step 2: Real-schema characterization tests

Using `LoadAndParseSettingsSchema` pointed at the repo schema (the test runs from `.build\x64\Debug`, so resolve the path robustly: walk up from the exe dir to the repo root looking for `Specs\SettingsStore.schema.json`, or accept the path as `argv[1]` with that walk as default):

1. Parse returns a non-empty vector.
2. Pick 2–3 anchor fields by READING the schema file first. **Verified anchor at plan time** (`Specs\SettingsStore.schema.json:1316-1340`): the property is `mainMenuSettings` and the field is `mainMenuSettings.menuBarVisible` — **NOT** `mainMenu.menuBarVisible` (the parser header comment at `SettingsSchemaParser.h:14` says `mainMenu.menuBarVisible`, but that comment is stale/illustrative; the real schema key is `mainMenuSettings`). Its values: `jsonPath = L"mainMenuSettings.menuBarVisible"`, `paneName = L"General"`, `schemaType = L"boolean"`, `controlType = L"toggle"`, `sectionHeader = L"Display"`, `displayOrder = 10`, `defaultValue = L"true"`. A good second anchor is its sibling `mainMenuSettings.functionBarVisible` (same pane/control/section, `displayOrder = 20`). For each anchor assert jsonPath, paneName, schemaType, and controlType against these literals. **Re-read the schema before trusting these numbers** — if they differ, the schema drifted (a STOP condition), do not silently "fix" the asserts to whatever the parser returns.
3. Sorting contract: the result is grouped by pane and non-descending by displayOrder within a pane+section (assert the invariant over the whole vector, not hardcoded positions).
4. `GetFieldsForPane` with a known pane (e.g. `L"General"`) returns only that pane's fields and at least one; with a garbage pane name (`L"NoSuchPane"`) returns empty.
5. `GetNonCustomFieldsForPane` excludes every field whose controlType is `custom` (assert none in the result, AND that for some pane it returns fewer fields than `GetFieldsForPane` — proving the filter actually drops something). Custom-control fields in this schema come from `$defs`-level definitions, which default to `controlType == "custom"` (`SettingsSchemaParser.cpp:261`); read the parsed set to find which pane owns one rather than hardcoding a pane name, since `$defs` membership is the schema's to change.

**Verify**: run exe → all `[       OK ]`, no `[ FAILED  ]`, exit 0.

### Step 3: Synthetic-input tests (inline JSON literals in the test)

1. Minimal valid schema with one `properties`-level boolean field carrying `x-ui-pane` but NO `x-ui-control` → exactly one field; assert these parser defaults (verified in `SettingsSchemaParser.cpp`): `controlType == L"edit"` (`:140`), `title == jsonPath` when `title` is absent (`:150`), `schemaType == L"boolean"`, and `defaultValue` round-trips (`true`→`L"true"`, `:198-201`).
2. Number field with min/max → `hasMin/hasMax` true and values correct; number field without bounds → both false.
3. Enum/combo field → `enumValues` populated in schema order.
4. Field WITHOUT `x-ui-pane` → not returned (parser contract: only pane-attributed fields).
5. Malformed JSON (truncated, wrong types e.g. `"minimum": "abc"`, empty string, huge-but-valid `minimum` beyond int64 if the parser claims int64) → no crash, empty vector or field skipped; these are characterization checks: first RUN them to see what the parser actually does, then assert that behavior (`noexcept` functions returning empty on bad input is the expected shape).
6. `LoadAndParseSettingsSchema` with a nonexistent path → empty vector, no crash.
7. Non-ASCII content in titles/descriptions (UTF-8 → wstring round trip) → preserved.

**Verify**: run exe → all pass, exit 0.

### Step 4: Optionally register in the `-Suite Full` runner, then finish

Optional: to make `SettingsSchemaTests` run under the repo's one-command runner (`Tools\Run-AllTests.ps1`), add `'SettingsSchemaTests'` to the executable array inside the `if ($Suite -eq 'Full')` block of `Tools\TestRunPlan.ps1` — at plan time the array is at **`:412`**: `foreach ($exeName in @('DxUiTests', 'FileSystemCurlTests', 'ViewerPETests', 'ViewerSqliteTests', 'MonitorTest', 'LocalizationTests', 'RedConfigureTests', 'PluginContractTests'))`. `'PluginContractTests'` (added by plan 004) is the model entry — append yours the same way. There is no `test-all.ps1`, and the runner has no per-exe `-Suite` value.

Then finish. **Format only your new source file, NOT the whole tree** — `.\format-all.ps1` reformats every file in the repo and is known to reorder includes in unrelated files (e.g. `RedConfigure\Workspace\WorkspaceDiscovery.cpp`) and break the build. Instead run `clang-format --style=file -i Tests\SettingsSchemaTests\SettingsSchemaTests.cpp`, then `git status --porcelain` and confirm ONLY your 4 in-scope files are dirty (revert anything else with `git checkout -- <file>`). Then: full `.\build.ps1` → exit 0; run the exe; run `LocalizationTests.exe` to confirm nothing else broke via the `.sln` edits. If you registered the suite, also confirm `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` lists a `SettingsSchemaTests` row (note: that runner executes the full ~18-20 min suite — checking the row is enough; you do not need to run it to completion).

## Test plan

This plan is the test plan: ~12-15 checks across Steps 2–3, modeled on `Tests\LocalizationTests\LocalizationTests.cpp` `Check` style. Verification: `.\.build\x64\Debug\SettingsSchemaTests.exe` → exit 0.

## Done criteria

- (archived, not live)  New test exe builds in the solution and exits 0
- (archived, not live)  Tests cover: real-schema anchors, sort invariant, pane filtering, custom exclusion, min/max, enums, missing x-ui-pane, malformed JSON, missing file, non-ASCII (grep the test source for one check per item)
- (archived, not live)  No modifications to `SettingsSchemaParser.*` or `Specs/**` (`git status`)
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- `SettingsSchemaParser.cpp` does not compile/link standalone **after** the two documented additions (Step 1 items 3 and 4: `RedSalamander\`+`Common\` on the include path, `yyjson.lib` linked) — i.e. it pulls in some *further* app-only dependency beyond `Framework.h`/`targetver.h`/`Win32CallbackHelpers.h`/yyjson that was not anticipated here. Report exactly what symbol/header is unresolved; the fallback (moving the parser to Common) is a maintainer decision. (Do NOT treat a yyjson unresolved-external or a missing-`Framework.h` error as this STOP — those mean Step 1 item 4 or 3 was done wrong; fix the vcxproj and retry.)
- A malformed-input test CRASHES the parser (not just returns empty) — that is a real bug in `noexcept` code; report it with the input, mark the check disabled with a comment, continue with the rest.
- `Specs\SettingsStore.schema.json` doesn't exist at that path — find the real path (`Glob Specs/**/*.schema.json`), use it, and note the doc drift.

## Maintenance notes

- When schema fields/panes are renamed, the anchor checks in Step 2 need the same rename — keep anchors few and stable.
- The `SettingsSchemaParser.h:14` comment says `mainMenu.menuBarVisible`, but the real schema key is `mainMenuSettings.menuBarVisible`. Don't "trust the comment" when picking anchors; it's illustrative, not authoritative. (Fixing that stale comment is out of scope here — it's a source edit — but worth a one-line follow-up.)
- The yyjson link (Step 1 item 4) and the `RedSalamander\` include dir (item 3) are the two non-obvious build requirements; a reviewer should confirm both survive in the committed `.vcxproj`, and that no PCH (`pch.h`) crept in from copying the wrong template.
- Follow-up (deferred): round-trip test with `SettingsSchemaExport.cpp` (export → parse → compare); registry-level SettingsStore tests if/when the store gains an injectable root key.
- Reviewer focus: that characterization checks for malformed input assert real observed behavior, not wishful behavior.

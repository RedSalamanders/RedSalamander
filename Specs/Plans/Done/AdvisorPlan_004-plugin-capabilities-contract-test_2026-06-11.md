# Advisor Plan 004 - Plugin factory and capabilities contract test

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/004-plugin-capabilities-contract-test.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/004-plugin-capabilities-contract-test.md`
- **Live owner:** none (implemented as Tests/PluginContractTests)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 004: Add a plugin contract test that loads every plugin DLL and validates its factory exports and capabilities JSON

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- Common/PlugInterfaces Tests Plugins/*/Factory.cpp RedSalamander.sln`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW (new test code only; no production changes)
- **Depends on**: none (but plan 005 depends on THIS — it is the safety net for the factory refactor)
- **Category**: tests
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

`GetCapabilities` was recently made a mandatory contract (commit `6fe57868d` "Remove legacy GetCapabilities fallback; enforce mandatory contract"), and every plugin DLL exports four factory functions that the host trusts. Nothing in CI loads the built plugin DLLs and verifies they honor the contract — a malformed capabilities document or a broken export is discovered at app runtime, not in CI. This test is also the prerequisite safety net for plan 005 (de-duplicating the 14 near-identical Factory.cpp files): with this green before and after, that refactor becomes low-risk.

## Current state

- The factory ABI (exemplar: `Plugins\FileSystemDummy\Factory.cpp`, 105 lines; all 14 plugins follow the same shape):
  - `extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)` — returns `E_NOINTERFACE` for a non-matching IID; on match returns a static metadata array and count (FileSystemDummy returns 1; some DLLs may enumerate multiple plugins — e.g. ViewerWeb has Web/Json/Markdown kinds; verify per-DLL at runtime, don't assume 1).
  - `extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions*, IHost* host, const wchar_t* pluginId, void** result)`
  - `extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)`
  - Declarations live in `Common\PlugInterfaces\Factory.h` (consumed with `#define PLUGFACTORY_EXPORTS` in plugins; the host side uses `GetProcAddress` — see `RedSalamander\FileSystemPluginManager.cpp:150-175`, which loads with `LoadLibraryExW(path, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH)` and resolves `RedSalamanderEnumeratePlugins`).
- The capabilities contract, documented at `Common\PlugInterfaces\FileSystem.h:472-493`:
  - `virtual HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept = 0;`
  - Doc comment: implementations MUST return `S_OK` and a non-empty UTF-8 JSON document with `"version": 1`, `"operations"`, `"concurrency"`, and `"crossFileSystem"`. Returned pointer owned by the plugin, valid until next call or release.
- The two plugin interface IIDs: `__uuidof(IFileSystem)` (`Common\PlugInterfaces\FileSystem.h`) and `__uuidof(IViewer)` (`Common\PlugInterfaces\Viewer.h`).
- Plugin DLLs in the build output (`.build\x64\Debug\`): `FileSystem.dll`, `FileSystem7z.dll`, `FileSystemCurl.dll`, `FileSystemDummy.dll`, `FileSystemGoogleDrive.dll`, `FileSystemMicrosoftDrive.dll`, `FileSystemS3.dll`, `ViewerText.dll`, `ViewerSqlite.dll`, `ViewerSpace.dll`, `ViewerImgRaw.dll`, `ViewerVLC.dll`, `ViewerPE.dll`, `ViewerWeb.dll`. (Confirm exact filenames by listing the output dir; some may differ — adapt the hardcoded list to what is actually produced by the corresponding `.vcxproj` names.)
- Existing test-exe pattern to copy: `Tests\LocalizationTests\` — a console exe with a local `Check(bool condition, const wchar_t* message, bool& success)` helper printing `[ OK ]`/`[ FAILED ]` lines and returning a process exit code; it includes `Helpers.h` with `#define REDSAL_DEFINE_TRACE_PROVIDER` and links Common. Model the new project on its `.vcxproj` (copy + rename + new GUID).
- Host stubs that already exist for tests: `Tests\PerformanceTests2\PluginManager.TestStubs.cpp` — check it before writing your own IHost stub.
- JSON parsing: yyjson is a repo dependency; follow `.github\skills\yyjson\SKILL.md` (ownership/cleanup rules) when validating documents.
- CI test invocation pattern (`.github/workflows/ci.yml:131-145`): plain `.\SomeTests.exe` from `.build/x64/Debug`, non-zero exit fails the job.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build all | `.\build.ps1` | exit 0 |
| Run new test | `.\.build\x64\Debug\PluginContractTests.exe` | `[ OK ]` lines, exit 0 |
| Existing tests still green | `.\.build\x64\Debug\LocalizationTests.exe` | exit 0 |

## Suggested executor toolkit

- Read `.github\skills\yyjson\SKILL.md` before writing the JSON validation.
- Read `.github\skills\plugin-callbacks\SKILL.md` for the host/plugin lifetime rules.

## Scope

**In scope**:
- `Tests\PluginContractTests\` (create: `PluginContractTests.cpp`, `PluginContractTests.vcxproj`, minimal resource files if the LocalizationTests pattern requires them)
- `RedSalamander.sln` (add the new project; mirror how `Tests\LocalizationTests` is declared, including configuration mappings for x64/ARM64 × Debug/Release/ASan Debug)
- `Tools\TestRunPlan.ps1` (OPTIONAL — add the new exe to the `-Suite Full` plan so `Tools\Run-AllTests.ps1 -Suite Full` runs it; see Step 5). There is no `test-all.ps1`; the local runner is `Tools\Run-AllTests.ps1`.

**Out of scope** (do NOT touch):
- Any file under `Plugins\` or `Common\PlugInterfaces\` — if a plugin FAILS the contract, that is a *finding to report*, not something to fix in this plan (fixing might be trivial, but it changes shipping code and belongs in its own change).
- `.github/workflows/ci.yml` — wiring into CI is plan 011.

## Git workflow

- Branch: `advisor/004-plugin-contract-tests`
- Subject style: "Add plugin factory and capabilities contract tests".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Scaffold the project

Copy `Tests\LocalizationTests\LocalizationTests.vcxproj` → `Tests\PluginContractTests\PluginContractTests.vcxproj`. Replace project name, output name, and `<ProjectGuid>` (generate: `powershell -Command "[guid]::NewGuid().ToString('B').ToUpper()"`). Strip LocalizationTests-specific sources/resources; add `PluginContractTests.cpp`. Add the project to `RedSalamander.sln` next to the other test projects (one `Project(...)` block + per-configuration entries in `GlobalSection(ProjectConfigurationPlatforms)` — copy LocalizationTests' lines and substitute the new GUID). Create a minimal `PluginContractTests.cpp` with a `wmain` returning 0.

**Verify**: `.\build.ps1` → exit 0 and `.build\x64\Debug\PluginContractTests.exe` exists and exits 0.

### Step 2: Enumerate-and-validate pass (metadata + schema — no instance creation)

In `PluginContractTests.cpp`, implement:

1. A hardcoded table of plugin DLL names split into filesystem vs viewer sets (adjust to the actual filenames you find in the output dir; the test runs from the output dir, so use relative names like `FileSystem7z.dll`).
2. For each DLL: `LoadLibraryExW(name, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH)` (match the host's flags, `FileSystemPluginManager.cpp:154`); `Check` it loaded.
3. Resolve all three exports with `GetProcAddress`; `Check` each is non-null.
4. Call `RedSalamanderEnumeratePlugins` with the matching IID; `Check` it returns `S_OK`, `count >= 1`, and for each of the `count` metadata entries: non-empty `id`, non-empty `name`, non-null `version`. Call it with the *wrong* IID; `Check` it returns `E_NOINTERFACE` and `count == 0`.
5. For each enumerated plugin id, call `RedSalamanderGetConfigurationSchema`; if it returns `S_OK`, `Check` the schema parses as JSON via yyjson (a plugin MAY legitimately have no schema — `ERROR_NOT_FOUND` HRESULT is acceptable; treat only other failures and unparseable JSON as failures). Use the wide→UTF-8 helpers from `Helpers.h` rather than hand-rolling conversions.

**Verify**: `.\build.ps1 -ProjectName PluginContractTests` (or full build) then run the exe → all checks `[ OK ]`, exit 0.

### Step 3: GetCapabilities pass (filesystem plugins, instance creation)

1. Look at `Tests\PerformanceTests2\PluginManager.TestStubs.cpp` for an existing IHost stub; reuse its approach (copy the minimal stub class into this test if it isn't in a shareable location — do NOT move/refactor the original).
2. For each *filesystem* DLL and each enumerated plugin id: call `RedSalamanderCreate(__uuidof(IFileSystem), nullptr /*FactoryOptions*/, stubHost, pluginId, &instance)`. `Check` `S_OK` and non-null instance.
3. Call `instance->GetCapabilities(&json)`; `Check`: `S_OK`, non-null, non-empty; parses with yyjson; root object has `version` == integer `1`; has `operations` (object), `concurrency` (object), `crossFileSystem` (object) — exactly the contract from `FileSystem.h:476-477`.
4. Release the instance (`instance->Release()` or `wil::com_ptr` attach), free the JSON per the documented ownership (plugin-owned pointer — do NOT free it), then `FreeLibrary` via `wil::unique_hmodule` at scope end.
5. IMPORTANT lifetime note: keep the module loaded for the whole time any interface pointer or returned `const char*` from it is alive (declare the `wil::unique_hmodule` first in scope).

If a specific plugin crashes or hangs on instance creation with the stub host (plausible for plugins needing real host services — e.g. cloud plugins), catch that case by FIRST running step 3 plugin-by-plugin manually; any such plugin goes on an explicit skip-list in the test source with a comment naming the failure, and your final report lists it. Do not silently skip.

**Verify**: run the exe → all checks pass for non-skip-listed plugins, exit 0. Run it twice in a row (idempotence / unload hygiene) → still 0.

### Step 4: Negative-contract test

Add one synthetic check: calling `RedSalamanderCreate` with a bogus plugin id (e.g. `L"no/such-plugin"`) must return `HRESULT_FROM_WIN32(ERROR_NOT_FOUND)` (the pattern at `Plugins\FileSystemDummy\Factory.cpp:95-98`). Run against FileSystemDummy.dll only.

**Verify**: exe exit 0.

### Step 5: Optional — register in the `-Suite Full` local test runner

The repo's one-command runner is `Tools\Run-AllTests.ps1`; its `-Suite Full` plan is built by `Get-RSTestRunPlan` in `Tools\TestRunPlan.ps1`. To make `PluginContractTests` run under `-Suite Full`, add `'PluginContractTests'` to the `foreach ($exeName in @('DxUiTests', …, 'RedConfigureTests'))` array inside the `if ($Suite -eq 'Full')` block (around `Tools\TestRunPlan.ps1:405-412`). There is NO per-exe `-Suite` value (`-Suite` is a fixed ValidateSet `All`/`Compare`/`Commands`/`FileOps`/`Full`) and NO `test-all.ps1`. (CI registration is plan 011's separate concern — do not touch `ci.yml` here.)

**Verify**: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → the summary lists a `PluginContractTests` row marked `[PASS]` (requires the exe already built).

## Test plan

This plan IS the test. Case inventory (all in `PluginContractTests.cpp`, `Check`-style like `Tests\LocalizationTests\LocalizationTests.cpp:33-43`):
happy-path load/enumerate/schema for ~14 DLLs; wrong-IID rejection; capabilities JSON schema fields for all filesystem plugins; bogus-id rejection; double-run hygiene.

## Done criteria

- (archived, not live)  `.\build.ps1` exit 0 with the new project in the solution
- (archived, not live)  `.\.build\x64\Debug\PluginContractTests.exe` exit 0; output contains an `[ OK ]` line for every non-skip-listed plugin DLL
- (archived, not live)  Skip-list (if any) is empty OR each entry has a comment + is named in your report
- (archived, not live)  No files under `Plugins\` or `Common\` modified (`git status`)
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- More than 3 plugins need the skip-list in Step 3 — the stub-host approach is then wrong for this codebase; report what host services they require instead of building an elaborate fake.
- A plugin genuinely violates the documented capabilities contract (missing mandatory field) — the TEST is then correctly failing; report the violation, put the plugin on the skip-list with a `// CONTRACT VIOLATION:` comment, and let the maintainer decide. Do not fix the plugin.
- The .sln edit breaks loading in any configuration (build errors mentioning project GUIDs) twice in a row.

## Maintenance notes

- Plan 005 (factory dedup) MUST run this test before and after its refactor; that's the main reason this plan exists.
- When a new plugin is added, its DLL name must be added to the test table — a follow-up could derive the list from the build output dir pattern (`*.dll` with a `RedSalamanderEnumeratePlugins` export) to remove the hardcoding; deferred to keep this deterministic.
- Reviewer focus: module-vs-pointer lifetimes in Step 3 (interface pointer must not outlive its DLL).

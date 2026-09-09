# Advisor Plan 007 - Crash quarantine and crash-handler tests

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/007-crash-quarantine-tests.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/007-crash-quarantine-tests.md`
- **Live owner:** none (implemented as Tests/CrashHandlingTests)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 007: Test coverage for crash quarantine and crash-handler path logic

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- RedSalamander/CrashHandler.h RedSalamander/CrashHandler.cpp RedSalamander/CrashQuarantine.h RedSalamander/CrashQuarantine.cpp RedSalamander/AppDataPaths.cpp`
> If any of those five production files changed since this plan was written,
> compare the "Current state" excerpts against the live code before proceeding;
> on a mismatch, treat it as a STOP condition. (`Tests\` was removed from the
> scope — it is a modification *target*, not a drift-sensitive excerpt, so it
> would false-trip on every new test project.)

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED (may need to add small testability seams to production code — tightly bounded below)
- **Depends on**: none
- **Category**: tests
- **Planned at**: commit `d6bfccc42`, 2026-06-11; **re-verified & tightened against `551e35655`, 2026-06-21** (the five production crash files are byte-identical to plan time — no drift; scaffolding + seam guidance updated below from a direct read of the source)

## Why this matters

Crash handling is the code that runs when everything else already failed — and it has zero test coverage. `CrashQuarantine` decides whether to disable a user's filesystem plugin after a crash (it mutates user settings — a wrong decision silently turns off functionality); `CrashHandler` constructs dump/marker paths under `%LOCALAPPDATA%`. Regressions here are invisible in normal runs and only surface during real crashes, which is the worst possible time. The logic is small and mostly file-I/O + decision logic, so targeted tests are cheap once the path roots are controllable — and `%LOCALAPPDATA%` is a per-process environment variable, which gives us the seam for free.

## Current state

- `RedSalamander\CrashQuarantine.h` (read at plan time, full contents):
  ```cpp
  namespace CrashQuarantine
  {
  // Best-effort: if a crash marker exists, offers to disable the last-active filesystem plugin in settings.
  // Must be called after settings load and before plugin initialization.
  void OfferPluginDisableIfPreviousCrashDetected(Common::Settings::Settings& settings) noexcept;
  }
  ```
  Note "offers to" — the implementation likely shows UI (a prompt). Your Step 1 confirms.
- `RedSalamander\CrashQuarantine.cpp` — the audit reported the crash marker lives at `%LOCALAPPDATA%\RedSalamander\Crashes\last_crash.txt` (around lines 25-33); verify by reading the file.
- `RedSalamander\CrashHandler.h` (741 bytes — small API; read it) and `CrashHandler.cpp` — installs the handler, writes minidumps, shows previous-crash UI (`Install()`, `WriteDumpForException()`, `ShowPreviousCrashUiIfPresent()` are called from `RedSalamander.cpp` startup/shutdown).
- `RedSalamander\AppDataPaths.cpp` — **VERIFIED at re-tighten time** (`AppDataPaths.cpp:17-41`): `GetLocalAppDataPath()` calls `SHGetKnownFolderPath(FOLDERID_LocalAppData, …)` **first** and returns immediately on success; it only falls back to `GetEnvironmentVariableW(L"LOCALAPPDATA")` if the known-folder API *fails*. **Consequence: a per-process `%LOCALAPPDATA%` env override does NOT reliably reach the marker path** (SHGetKnownFolderPath normally succeeds and wins). So do **not** build the quarantine tests on real marker files at the real path — use the pure-decision seam (below). Do not modify `AppDataPaths.cpp` (out of scope).
- **VERIFIED — the quarantine decision and the UI are FUSED** in `OfferPluginDisableIfPreviousCrashDetected` (`CrashQuarantine.cpp:67-147`): it (1) builds the marker path and checks `exists`, (2) calls `SessionState::TryRead()` for `activeFileSystemPluginIds`, (3) computes the `toDisable` list (skipping ids already in `settings.plugins.disabledPluginIds` via `IsDisabledPluginId`), (4) shows a prompt via `HostShowPrompt`, (5) on `YES` applies `DisablePluginIdInSettings` to each. The pure, testable core is step 3 + the two anonymous-namespace helpers `IsDisabledPluginId` / `DisablePluginIdInSettings` (`:36-64`) — they take only a `std::wstring_view` and `Common::Settings::Settings&`, no I/O, no UI.
- **Dependency-closure WARNING (this is the plan's main risk).** Compiling `CrashQuarantine.cpp` into a test exe drags in everything its functions reference at link time, even if the test never calls them: `AppDataPaths::GetLocalAppDataPath` (AppDataPaths.cpp), `SessionState::TryRead` (SessionState.cpp), `HostShowPrompt` (HostServices — likely heavy), and `LoadStringResource`/`FormatStringResource` (resource helpers). The repo's established way to satisfy such awkward externals without pulling the whole app is a **stubs translation unit** — see `Tests\PerformanceTests2\PluginManager.TestStubs.cpp` (a sibling `.cpp` in the test project that defines the externals the imported app source needs). The executor should plan for a `CrashQuarantine.TestStubs.cpp` (or equivalent) and STOP per the conditions below if the closure still explodes past ~4 app `.cpp` files or pulls window/D2D.
- **Test-only exposure mechanism — VERIFIED precedent:** `#ifdef ENABLE_TESTS` blocks in BOTH header and `.cpp` are the repo idiom for exposing internals to tests (e.g. `CompareDirectoriesEngine.h:261`, `ConnectionCredentialPromptDialog.h:42` + `.cpp:233`, `ConnectionSecrets.h:64`). Expose the extracted pure decision function this way; the test project defines `ENABLE_TESTS`.
- **`CrashHandler.h` (VERIFIED, `:1-22`) exposes only SEH/UI/minidump entry points** — `Install()`, `WriteDumpForException(EXCEPTION_POINTERS*)`, `ShowPreviousCrashUiIfPresent(HWND)`, `TriggerCrashTest()`. There is **no pure path function in the public API**; any testable marker/path helper lives in the `.cpp`'s anonymous namespace (533 lines, SEH/minidump-heavy). Expect Step 5 to yield little — "half this plan (quarantine tests) is fine" per Step 5's own note.
- These files are part of the RedSalamander app project (not Common). **Scaffolding template: copy `Tests\PluginContractTests\PluginContractTests.vcxproj`** (console `Application`, no PCH, compiles an external app `.cpp` directly, `Common` ProjectReference) — the same clean template plan 006 used. Use `Tests\PerformanceTests2\PerformanceTests2.vcxproj` ONLY as the reference for the `*.TestStubs.cpp` pattern, not as the project base.
- Test-exe output pattern: `Tests\LocalizationTests\LocalizationTests.cpp:33-43` — `Check(bool, const wchar_t*, bool&)` printing `[       OK ]` / `[ FAILED  ]` (exact spacing), `wmain` returns 0/1. Copy it verbatim so the runner grep matches.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build (scoped, fast) | `.\build.ps1 -ProjectName CrashHandlingTests` | exit 0 |
| Build (full) | `.\build.ps1` | exit 0 |
| Run new tests | `.\.build\x64\Debug\CrashHandlingTests.exe` | exit 0, `[       OK ]` lines, no `[ FAILED  ]` |
| Selftests unaffected | full `--selftest` via Start-Process pattern | exit 0 — **BUT see the worktree caveat below** |

> **Worktree `--selftest` caveat (read before relying on the done-criteria):** in
> an isolated worktree the full `--selftest` is known to abort with
> `STATUS_FATAL_APP_EXIT` from a **pre-existing** issue unrelated to this plan (a
> removed `--plugin-path` option referenced by the CompareDirectories selftest
> harness). If `--selftest` crashes, first reproduce it at the **unchanged base
> commit** (`git stash` your changes or check the base) — if it crashes there too,
> it is pre-existing, NOT introduced by your seam; document that and do **not**
> treat it as a failure or a STOP. The substantive "production behavior unchanged"
> gate is then: full `.\build.ps1` green **and** the seam `git diff` is inert
> (defaulted params / extracted pure function / `#ifdef ENABLE_TESTS` only) **and**
> the new test exe passes. If you can run any narrower, non-crashing selftest that
> exercises settings/plugin load, do so and report it.

## Scope

**In scope**:
- `Tests\CrashHandlingTests\` (create)
- `RedSalamander.sln` (add project)
- `RedSalamander\CrashQuarantine.cpp` / `CrashHandler.cpp` — ONLY the bounded seam changes listed in Step 2 (nothing else)
- `Tools\TestRunPlan.ps1` (OPTIONAL — register `CrashHandlingTests` in the `-Suite Full` plan by appending its name to the executable array in the `if ($Suite -eq 'Full')` block, at **`:412`**: `@('DxUiTests', 'FileSystemCurlTests', 'ViewerPETests', 'ViewerSqliteTests', 'MonitorTest', 'LocalizationTests', 'RedConfigureTests', 'PluginContractTests')` — append `'CrashHandlingTests'` the same way `'PluginContractTests'` was added). There is no `test-all.ps1`.

**Out of scope** (do NOT touch):
- `RedSalamander.cpp` call sites — the seams must not change how production calls these functions (default arguments / overloads only).
- Actually triggering real SEH crashes / writing real minidumps in tests (a controlled end-to-end crash test is explicitly deferred — see Maintenance notes).
- `AppDataPaths.cpp` behavior — read it, don't change it (if it blocks testability, that's a STOP).

## Git workflow

- Branch: `advisor/007-crash-handling-tests`
- Subjects: "Add crash quarantine path/decision tests" etc.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Read the implementation and map the seams

Read `CrashQuarantine.cpp`, `CrashHandler.h/.cpp`, and the path-construction part of `AppDataPaths.cpp`. Produce (in your report) a 10-line map: where the marker path comes from, whether `%LOCALAPPDATA%` env override reaches it, where the "offer" UI happens (MessageBox? task dialog? Debug-build bypass?), and which functions are pure-ish (path construction, marker parse, decision logic).

**Verify**: you can name (a) the exact function that decides "should we offer to disable" and (b) the exact function that shows UI. If they are the same function with no separation, Step 2's seam applies.

### Step 2: Add minimal seams (only as needed)

**RECOMMENDED seam (given the verified facts in "Current state"):** seam #2 — extract a **pure** decision function and expose it via `#ifdef ENABLE_TESTS`. Concretely, factor the step-3 computation out of `OfferPluginDisableIfPreviousCrashDetected` into something like:

```cpp
// CrashQuarantine.h, inside namespace CrashQuarantine, test-only:
#ifdef ENABLE_TESTS
// Pure: given whether a crash marker exists and the active FS plugin ids from the
// previous session, returns the ids that should be offered for disable (those not
// already disabled). No file I/O, no UI.
[[nodiscard]] std::vector<std::wstring> SelectPluginsToDisable(
    bool markerExists,
    const std::vector<std::wstring>& activeFileSystemPluginIds,
    const Common::Settings::Settings& settings) noexcept;
#endif
```

`OfferPluginDisableIfPreviousCrashDetected` then computes `markerExists`/`activeIds`, calls `SelectPluginsToDisable`, shows the prompt, and applies `DisablePluginIdInSettings` on YES — i.e. **its observable behavior is unchanged**. The test calls `SelectPluginsToDisable` directly (no marker file, no `SessionState`, no `HostShowPrompt`, no UI) and separately can call `DisablePluginIdInSettings` (also expose it under `ENABLE_TESTS` if needed) to assert the apply step. This keeps the *call* graph the test exercises free of UI even though the *translation unit* still links the heavy externals (handle those via the stubs TU in Step 3).

Avoid seam #1 (the `overrideRoot` path param) unless you specifically test `GetCrashMarkerPath` — it is not needed for the decision tests and `AppDataPaths` is out of scope.

**Verify**: `.\build.ps1` → exit 0; `--selftest` per the worktree caveat in "Commands" (inert seam diff is the real gate).

### Step 3: Scaffold the test project

Copy `Tests\PluginContractTests\PluginContractTests.vcxproj` (console Application, no PCH, fresh GUID, add `Project`/config rows/NestedProjects to `RedSalamander.sln`; add `$(SolutionDir)RedSalamander` to `AdditionalIncludeDirectories` and define `ENABLE_TESTS` in `<PreprocessorDefinitions>`). Compile in `..\..\RedSalamander\CrashQuarantine.cpp` plus a sibling **`CrashQuarantine.TestStubs.cpp`** that defines the externals the imported source needs but the test must not really invoke (`HostShowPrompt`, `SessionState::TryRead`, `AppDataPaths::GetLocalAppDataPath`, `LoadStringResource`/`FormatStringResource` — provide trivial stubs returning empty/none) — model the file on `Tests\PerformanceTests2\PluginManager.TestStubs.cpp`. Prefer stubs over compiling the real `AppDataPaths.cpp`/`SessionState.cpp`/`HostServices.cpp` so the closure stays tiny. If you cannot satisfy the externals with ≤ ~4 app `.cpp` files **or** a stubs TU and it keeps dragging in window/D2D machinery, **STOP** per below.

**Verify**: `.\build.ps1 -ProjectName CrashHandlingTests` → exit 0; empty test exe runs and exits 0.

### Step 4: Quarantine tests

Under a temp directory root (`[System.IO.Path]::GetTempPath()` equivalent in C++: `std::filesystem::temp_directory_path()` + unique subdir, cleaned up in the test):
1. No marker file → decision function returns "no offer"; settings untouched.
2. Marker present + settings have a last-active filesystem plugin → decision is "offer"; simulate accept → the plugin is disabled in the `Settings` struct; simulate decline → settings untouched. (How accept/decline is simulated depends on Step 2's seam — decision function returns data, test applies it, or callback injection.)
3. Marker present but unreadable/garbage content → no crash (`noexcept`), behavior characterized (assert what you observe; comment it).
4. Marker is deleted/consumed after handling if that's the production behavior (verify in Step 1; assert it).
5. Settings with NO last-active plugin + marker present → no offer, no crash.

**Verify**: `CrashHandlingTests.exe` → exit 0.

### Step 5: Crash-handler path tests

For whatever pure path/marker functions Step 1 found in `CrashHandler.cpp` (dump file naming, crash-marker write/read round-trip):
1. Dump/marker path is under the expected root and contains no invalid filename characters even with hostile module names/timestamps if those feed in.
2. Marker write → `ShowPreviousCrashUiIfPresent`-adjacent *detection* logic (not the UI) sees it; absent → doesn't.

If `CrashHandler.cpp`'s logic is too entangled with SEH/minidump APIs to test any pure part, document that in the report and keep only the quarantine tests — half this plan delivered is fine.

**Verify**: test exe exit 0.

### Step 6: Finish

`.\format-all.ps1`, full build, full `--selftest`, run the new exe twice (temp-dir cleanup hygiene).

## Test plan

Steps 4–5 enumerate the cases (~8-10 checks), `Check`-style per `Tests\LocalizationTests\LocalizationTests.cpp:33-43`. Verification: new exe exit 0 AND full selftest exit 0 (proves seams changed nothing).

## Done criteria

- (archived, not live)  `CrashHandlingTests.exe` builds and exits 0 with at least the 5 quarantine cases from Step 4
- (archived, not live)  Production behavior unchanged: full `.\build.ps1` exit 0, AND `git diff` on `CrashQuarantine.cpp`/`CrashHandler.cpp` shows only seam-shaped changes (defaulted params / extracted pure function / `#ifdef ENABLE_TESTS` blocks). For `--selftest`: exit 0 if it runs; if it hits the pre-existing worktree `STATUS_FATAL_APP_EXIT` (see the caveat in "Commands"), the criterion is satisfied by confirming the SAME crash at the unchanged base + an inert seam diff.
- (archived, not live)  No UI appears when running the test exe in a normal console session (guaranteed if the tests call only the extracted pure decision function, never `OfferPluginDisableIfPreviousCrashDetected` with a marker present)
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- The dependency closure to compile `CrashQuarantine.cpp` into a test exe exceeds ~4 app .cpp files or pulls in window/D2D machinery — the right fix is then moving the logic, which is a maintainer decision.
- Seam work would change any production call-site signature in `RedSalamander.cpp`.
- The quarantine flow turns out to be interactive-only with no separable decision (pure UI) — report; a UI-driven Commands-selftest case would be the right vehicle instead, and that's a different plan.
- Tests require Administrator or write outside the temp dir.

## Maintenance notes

- Deferred deliberately: an end-to-end crash test (child process raises SEH, parent asserts a dump+marker appear). Valuable but heavier; revisit once this lands.
- Reviewer focus: the seams must be inert in production (defaulted args, ENABLE_TESTS gating consistent with existing usage at `ConnectionCredentialPromptDialog.cpp:232-234`).

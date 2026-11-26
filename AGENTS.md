# RedSalamander development guidance

## Project Overview

**RedSalamander** is a Windows-native file manager and monitoring application built in MSVC `stdcpplatest` mode. It features dual-pane file management, plugin-based virtual file systems, advanced text visualization with Direct2D/DirectWrite rendering, and real-time debugging via ETW (Event Tracing for Windows).

## Requirements

- **Platform**: Windows 11 build 22000.2600 minimum
- **Language Standard**: `stdcpplatest` (latest MSVC mode)
- **Character Encoding**: Unicode UTF-16
- **Build System**: Visual Studio 2026 with MSBuild
- **Package Manager**: vcpkg
- **Graphics APIs**: Direct2D, DirectWrite, Direct3D 11, DXGI

## Drive-root output policy

Do not create scratch, caches, evidence, copied checkouts, or other output directories directly at a drive root.
Use repository-local build outputs or an existing approved workspace, subject to the stricter test-sandbox
contract below. The only permitted new drive-root output name is `<Project>.Perf` (for example
`RedSalamander.Perf`, `RedPrism.Perf`, or another project's own name), and the user must explicitly authorize
the exact path before its first creation or initialization. A matching name, environment variable, low disk
space, or authorization for another drive/project is not authorization. Keep outputs beneath the approved
root, never in a sibling root folder. Test runners still require their ownership marker and dedicated layout.
Filesystem-specific tests inspect their acquired sandbox and skip when its filesystem is unsuitable; they
must not discover another volume and create an unapproved root there. Historical cleanup inventories and
preserves useful evidence/recovery data, checks process use, repairs live junction/worktree references, and
verifies preservation before removal. Never follow junctions during cleanup or leave compatibility junctions
at forbidden old drive-root paths.

## Canonical guidance

- `AGENTS.md` (this file) is the canonical agent guidance for this repository. `CLAUDE.md` and `.github/copilot-instructions.md` are thin pointers to it; do not let them accumulate divergent content.
- Project skills live in `.agents/skills/` (one directory per skill; table below), the skill folder of every RedSalamanders repository. Cross-repository C++, WIL, `yyjson`, localization, audit (`improve`), and `spec-workflow` skills live in the sibling `RedHome` repository.
- Every project skill, including the crash-forensics meta-skill (`red-salamander-crash-forensics`), lives only in `.agents/skills/` with `name` and `description` front matter; do not create copies under `.claude/`, `.github/skills/` or other tool-specific directories.
- Agent/skill authority, package/example integrity, compatibility routes and truthful evaluation are governed by [Testing_AgentGuidance.md](Specs/Testing/Testing_AgentGuidance.md). Keep mutable machine facts and delivery history with their owners rather than copying them into instructions.

## Skills

Shared skill index: `../RedHome/AGENTS.md`; packages: `../RedHome/.agents/skills/`. If `../RedHome` is missing, clone https://github.com/RedSalamanders/RedHome.git to that path. This repository's guidance and domain specs own product behavior and validation.

Use RedHome's shared `native-build` package for build, restore and packaging. Use `test-evidence` for testing and performance validation.

### Skills Reference

| Skill | Description |
|-------|-------------|
| [direct2d-rendering](.agents/skills/direct2d-rendering/SKILL.md) | Direct2D/DirectWrite graphics patterns |
| [icon-cache](.agents/skills/icon-cache/SKILL.md) | Shell icon management via IconCache |
| [win32-wndproc](.agents/skills/win32-wndproc/SKILL.md) | Window procedure and message handling |
| [plugin-callbacks](.agents/skills/plugin-callbacks/SKILL.md) | Plugin callback pattern with cookie |
| [file-actions](.agents/skills/file-actions/SKILL.md) | File-action ID semantics, settings validation, direct dispatch, and regression coverage |
| [build-preferences-pages](.agents/skills/build-preferences-pages/SKILL.md) | End-to-end Preferences page workflow: draft state, dirty tracking, persistence, accessibility, specs, and tests |
| [tooling-governance](.agents/skills/tooling-governance/SKILL.md) | Inventory, placement, help, compatibility, caller, cache, and spec rules for scripts and tooling |
| [theming](.agents/skills/theming/SKILL.md) | Theme color keys and JSON5 themes |
| [async-threading](.agents/skills/async-threading/SKILL.md) | Threading model and async patterns |
| [red-salamander-crash-forensics](.agents/skills/red-salamander-crash-forensics/SKILL.md) | Dumps, sidecar stacks, fatal dialogs, hangs, timeouts and failed selftest runs |

## Core Principles

### RAII is Mandatory
All Windows resources MUST use WIL RAII wrappers. Manual cleanup (`DestroyIcon`, `DeleteObject`, `EndPaint`, etc.) is **PROHIBITED** (see `RedHome`'s `wil-raii` skill).

### Performance Validation is Mandatory
Any new feature, hot-path change, or optimization that can affect responsiveness, throughput, queueing, rendering, startup, search, Compare Directories, File Operations, plugin I/O, or memory retention MUST integrate:
- scenario definition,
- instrumentation,
- deterministic selftest coverage,
- archived perf evidence under `Specs/TestRuns/`

from the beginning. Follow the [performance and resource contract](Specs/Core/Core_PerformanceAndResources.md) and the [performance validation procedures](Specs/Testing/Testing_PerformanceValidation.md). Use RedHome's `test-evidence` workflow to record validation evidence.

### Spec Closeout is Mandatory

Application screenshots for specifications, documentation, and UI reviews MUST come from deterministic test-harness
scenarios with no desktop takeover and realistic invented fixture data, never personal data or redaction labels.
[Testing_SelfTests.md](Specs/Testing/Testing_SelfTests.md#documentation-screenshot-scenarios) owns provenance, mockup
labeling, and the explicitly authorized brief-focus exception for interaction captures.
Completed WIP plans MUST be moved to `Specs/Plans/Done/`. Any durable behavior, UI contract, validation rule, or workflow requirement discovered during implementation MUST be merged into the authoritative domain spec under `Specs/<Domain>/` (or repo-level guidance such as `AGENTS.md` / `Specs/Testing/*` when appropriate) before the work is considered closed. Do not leave normative requirements stranded only in `Specs/Plans/WIP/` or `Specs/Plans/Done/`.

### Tooling Governance is Mandatory
Before adding, moving, renaming, or materially changing a file under `Tools/`, read
[`Tools/README.md`](Tools/README.md),
[`Specs/Testing/Testing_ToolingGovernance.md`](Specs/Testing/Testing_ToolingGovernance.md),
and the [tooling-governance skill](.agents/skills/tooling-governance/SKILL.md).
Update `Tools/tool-inventory.json`, public help, callers, focused tests, and the
owning normative spec in the same change. Keep stable entrypoints at the Tools
root, put reusable implementation in explicit modules, and preserve documented
paths with thin compatibility shims when they may have external consumers.

### Regression Guards (Common Violations)
- **Tools inventory and command surface are enforced:** every tracked file under
  `Tools/` must have one effective classification in `Tools/tool-inventory.json`.
  New public commands use `Verb-Noun.ps1`, include complete comment-based help,
  and declare side effects and outputs. Function-only implementation belongs in
  `Tools/Modules/` with explicit exports. Cache keys and workflow manifests must
  cover the full behavioral dependency closure. See
  `Specs/Testing/Testing_ToolingGovernance.md`.
- **Build/test coordination is profile-scoped and non-destructive:** artifact
  operations serialize by repository root plus platform/configuration; do not
  block disjoint profile lifecycles or worktrees. MSBuild/vcpkg phases also use
  the narrower repository+platform dependency lock because configurations share
  one triplet; packaging alone uses a repository-wide lock because platforms
  share installer inputs and `.build/AppPackages`. Process names alone are never ownership evidence. An independently
  launched executable at an exact output path blocks
  replacement with diagnostics but is never killed; automatic termination is
  limited to descendants contained by the launching operation. The session-wide
  interactive mutex covers only plan entries marked `RequiresInteractiveDesktop`,
  including their quarantine repair attempts. See
  `Specs/Build/Build_Toolchain.md` and `Specs/Testing/Testing_SelfTests.md`.
- **Tests preserve desktop focus whenever possible:** GUI coverage uses the
  governed no-activation mode unless its assertions require real foreground
  input. Focus-taking suites use the existing `DirectedSelfTestInputWarning`
  from `Tests/TestSupport/` before and during execution inside the runner's
  desktop lease. Its large, non-activating, click-through surface stays centered
  on the test app. Do not add duplicate warnings or weaken focus assertions.
  See `Specs/Testing/Testing_SelfTests.md` for the activation classes and warning contract.
- **Test paths are fail-closed:** all test-generated scratch, temp, logs, archives, and validation evidence stay
  beneath exact `X:\RedSalamander.Perf` (any selected fixed local drive) with its ownership marker.
  `REDSALAMANDER_TEST_ROOT` is a request, not authorization, and repository `.build` is never a test-data sandbox.
  Tests never mount, link, share, or `subst` anything outside a marked sandbox, sandbox cleanup never follows links,
  and each case purges what it moved to the Recycle Bin. [Testing_SelfTests.md](Specs/Testing/Testing_SelfTests.md)
  owns the full rules, the guarded link helpers, and the `TestHarnessSourceContracts.Tests.ps1` allowlist.
- **Shared-helper reuse is mandatory:** before adding a consumer-local utility, search
  `Specs/Core/Core_SharedHelpers.md`, `Common/`, and (for test code) `Tests/TestSupport/`. Reuse or extend the
  canonical helper when its semantics match; do not reimplement it under a different name. A local variant is
  allowed only when policy, dependency layer, ABI, ownership, error handling, or performance semantics differ.
  Name that difference, document it at the definition, and cover it in the relevant reviewed allowlist or
  source-contract test. Adding a new shared helper also requires updating the shared-helper catalog and the
  authoritative domain spec that owns its behavior.
- **Runtime dependency staging is centralized:** add or change app-local plugin DLLs in
  root `RuntimeDependencies.props`; do not reintroduce plugin-local `PostBuildEvent`/`xcopy` batches. Required
  inputs must fail through the shared MSBuild tasks, obsolete outputs must be removed declaratively, and portable
  packaging must consume the same manifest and pass the clean-extraction smoke contract in
  `Specs/Installer/Installer_PortableZip.md`.
- **Ban `sprintf_s` / `swprintf_s`** in non-PoC code:
  - Diagnostics: `std::format` / `std::format_to_n` + `OutputDebugStringA/W`
  - User-facing/localized: `.rc` resources + `FormatStringResource(...)` with **positional** placeholders
- **Resource strings use `std::format` positional placeholders** (`{0}`, `{1:08X}`), introduced in argument order, never bare `{}`, unindexed specs, or printf syntax; translations keep the source tokens. [Core_Localization.md](Specs/Core/Core_Localization.md) owns the full placeholder and translation rules.
- **COM ownership:** never store owning raw COM interface pointers (no manual `Release()`); use `wil::com_ptr<T>` for members and locals.
- **COM ref-counting:** never do `obj->AddRef(); ptr.attach(obj);` (two-step hazard); prefer `ptr = obj;` / `wil::com_ptr<T> ptr = obj;`.
- **`wil::unique_hwnd` ownership:** never call `DestroyWindow(_hWnd.get())` on a `wil::unique_hwnd` owner; use `_hWnd.reset()` (or `.release()` only when transferring ownership explicitly).
- **Cross-thread `PostMessageW` payloads:** post with `PostMessagePayload(...)` and receive with `TakeMessagePayload<T>(hwnd, message, lParam)`; never post `payload.release()` or a raw `new`. A nonzero `lParam` is an opaque token, never a pointer, and a null take is stale input to ignore. Receiving windows call `InitPostedPayloadWindow(hwnd)` on create and `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`. The async-threading skill owns the token, argument-capture, and drain rules and the review checklist for every UI-host cross-thread change.
- **IconCache COM contract:** `IconCache::Initialize(...)` stays UI-thread/STA responsibility; any worker thread calling `IconCache::ExtractSystemIcon()` must initialize COM as MTA (`wil::CoInitializeEx(COINIT_MULTITHREADED)`).
- **MTP/WPD worker contract:** portable-device workers must initialize COM as MTA before using WPD, serialize each device session, honor watchdog/cancel/quarantine teardown, and keep `GetCapabilities()` honest for read-only, writable, invalid-profile, and disconnected states.
- **File-action IDs:** treat viewer/editor/user-menu action IDs as stable but case-insensitive identifiers. Uniqueness and lookup are case-insensitive, IDs that differ only by case are invalid duplicates, and parameterized command dispatch or association resolution must never require exact casing.
- **yyjson mutable builders:** never pass temporary/stack strings to non-copy APIs (`yyjson_mut_obj_add_str`, `yyjson_mut_str`); for dynamic keys use `yyjson_mut_strncpy` + `yyjson_mut_obj_add`, and for string values prefer `*_strcpy`/`*_strncpy` (`RedHome` `yyjson` skill).
- **Thread-safety:** never read/write shared non-atomic state without a lock (or use `std::atomic` with correct memory ordering).
- **Exceptions / `catch (...)`:** `catch (...)` is **FORBIDDEN**. If exception handling is mandatory at an ABI boundary (`noexcept` methods, Win32 callbacks, thread entrypoints), catch only explicitly named exception types and add a short comment explaining why catching is required there and what the fallback is. Prefer non-throwing APIs (e.g. `std::filesystem` overloads with `std::error_code`). If you catch-and-fail, log once with `Debug::Error(...)`; if you catch-and-continue, document the fallback (avoid empty catch without an explanation). Treat `std::bad_alloc` as fatal (`std::terminate()`), and never let it be swallowed by a broader `catch (const std::exception&)`.
- **Detached threads:** avoid `std::thread(...).detach()` in plugin DLLs; prefer `TrySubmitThreadpoolCallback` or `std::jthread` + `stop_token`. If detach is unavoidable, pin module lifetime with `AcquireModuleReferenceFromAddress(...)` and make thread start exception-safe (especially in `noexcept` code).
- **Threadpool module pins:** a plugin threadpool callback that owns an `AcquireModuleReferenceFromAddress(...)` handle MUST transfer it at callback entry with `TransferModulePinToCallbackReturn(instance, pin)`. Destroying the last `wil::unique_hmodule` inside the callback can unmap the callback's own code before it returns.
- **Plugin unload quiet point:** stop producers → request shutdown/cancel → stop posting payload messages → `SetCallback(nullptr, nullptr)` → release instances → unload module only when no callbacks can still run.

### Modern C++
- Smart pointers over raw pointers
- `std::format` over string concatenation
- `std::optional` with `.value()` or `.value_or()` (never `*`)
- Range-based for loops and structured bindings
- `WIN32_LEAN_AND_MEAN` and `NOMINMAX` before `<windows.h>` in its own include block before `bcrypt.h`/`wincodec.h`, and `objbase.h`/`objidl.h` in its own block before `UIAutomation.h`/`XmlLite.h` (`RedHome` `modern-cpp-windows` skill)

### Error Handling
- Use HRESULT for Windows API calls
- `Debug::ErrorWithLastError(...)` for Win32 failures
- `Debug::Error(...)` for unexpected failures
- `Debug::Warning(...)` for recoverable failures
- Don't log normal control-flow (window size 0, cancellation, device recreation)

## Patterns to Avoid

- Raw `new`/`delete`
- C-style casts
- `goto` (use early returns + RAII / `wil::scope_exit`)
- Raw Windows handles without WIL
- Manual resource cleanup
- `sprintf_s` / `swprintf_s` in non-PoC code
- `DestroyWindow(...get())` on `wil::unique_hwnd` owners
- `AddRef()` + `attach()` (use `wil::com_ptr` assignment)
- Empty `catch (...) {}` without explaining why it’s safe to ignore
- Global state and singletons
- Blocking UI thread
- Hardcoded UI strings (use `.rc` resources)

## Architecture

The RedSalamander section of `../RedHome/.agents/skills/native-build/references/repository-facts.md` describes
the solution layout, project dependencies, components, plugin factory entry point and output locations.
[Plugins_PluginAPI.md](Specs/Plugins/Plugins_PluginAPI.md) owns the plugin ABI.

## Build

Use `build.ps1` for command-line builds. The [build toolchain contract](Specs/Build/Build_Toolchain.md) owns RedSalamander's build rules.

```powershell
.\build.ps1                              # Build all (Debug)
.\build.ps1 -Configuration Release       # Build all (Release)
.\build.ps1 -ProjectName RedSalamander   # Build specific project
```

Use `test.ps1 -Explain` and affected execution while editing. Use `test.ps1 -Mode PrePush` for the full local obligation. Use `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` when Full is explicitly requested or required for the changed behavior's qualification; `-SkipBuild` requires a byte-verified compatible receipt. Exact Full continuation uses `-ValidationMode Resume -ResumeFrom <run-id>`. An affected/family pass or pending CI never establishes complete repository qualification. See `Specs/Testing/Testing_ValidationEvidence.md` and `Tests/README.md` for the current acceptance and reuse boundaries.

## Dependencies

- **WIL**: Windows Implementation Library (RAII wrappers)
- **yyjson**: JSON parsing and serialization
- **DirectX**: Graphics APIs (D2D, D3D11, DXGI)
- Other key vcpkg packages: libraw, libjpeg-turbo, 7zip, pe-parse

## Specifications

Detailed component specs live under `Specs/`:
- `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_FolderWindow.md`, `Specs/UI/UI_NavigationView.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Plugins/Plugins_ViewerPlugins.md`, `Specs/Plugins/Plugins_PluginAPI.md`
- `Specs/UI/UI_PreferencesDialog.md`, `Specs/SettingsStore.schema.json`
- `Specs/Core/Core_RedSalamanderMonitor.md`

## LLM Assistant Guidelines

When working with this codebase:
- Always use WIL RAII wrappers for Windows resources
- Prioritize performance and responsiveness
- Treat performance validation as part of the feature contract: add or reuse metrics, add deterministic selftests, and archive runs when the scenario is perf-sensitive
- When finishing a plan, move it to `Specs/Plans/Done/` and update the authoritative spec or repo guidance so the lasting contract does not live only in the plan
- Consider DPI awareness in UI components
- Respect the Windows-specific architecture
- Suggest modern C++ patterns
- Reference appropriate skills for detailed patterns

## Scoped testing policy (2026-10-05)

Use `test.ps1 -Explain` and the affected default while editing. Use `-Mode PrePush` for the full local
obligation and `-Full` to request canonical Full acceptance. Pull request checks are compiler-free. Active native
test files use `Scope.Tests.Something.h/.cpp` and the native file inventory. Reuse exact prior evidence only under
[the validation contract](Specs/Testing/Testing_ValidationEvidence.md) and follow [the test guide](Tests/README.md).

## Interactive desktop tests

A test that takes real foreground focus, the cursor or the clipboard requires explicit agreement from the person at
the desktop for that time. Use the canonical interactive lease, warn before and during execution, and drive only
windows the fixture owns. Follow the [self-test contract](Specs/Testing/Testing_SelfTests.md) and restore focus only
while the fixture still owns it. Keep the hosted and manual desktop acceptance scenarios explicit.

## Changes workflow

Record each change as one dated Markdown fragment under `Changes/`; follow [Changes/README.md](Changes/README.md)
and [Specs/Core/Core_Changes.md](Specs/Core/Core_Changes.md). Do not hand-edit generated `CHANGELOG.md` entries.
Keep the fragment in the change: never fold it in a feature change and never edit `CHANGELOG.md`. Fragments are
folded on `main` after merge by `Tools/Submit-ChangelogFold.ps1`, which opens a fold pull request of its own. Run
`Tools/Test-Changelog.ps1` while editing; PrePush and CI run it with `-Base`, which refuses any `CHANGELOG.md` change
that is not an exact fold of fragments already on `main`. Validation is read-only.

## Shared native build policy (2026-10-28)

The [shared native build profiles](https://github.com/RedSalamanders/.github/blob/e513a50b3aea60d86f41054188537fb7b418bac6/Specs/NativeBuildProfiles.md) define the common six-profile baseline and product exceptions.

Use /Wall with /WX disabled in every first-party profile. A build is accepted
only with zero warnings and zero errors; fix actionable diagnostics rather than
adding blanket suppressions. Ordinary Debug uses /ZI, ASan Debug uses /Zi.
Release explicitly enables /OPT:ICF; Debug/ASan explicitly disable it.
Follow Specs/Build/Build_ConsoleAndPaths.md for the common green result, safe
terminal launch links and short dependency folders with full identity checking.

## Unified CI policy

Use [Specs/Build/Build_UnifiedCI.md](Specs/Build/Build_UnifiedCI.md) for hosted CI. PRs run compiler-free checks and formatting feedback. Main native qualification starts after five continuous quiet minutes; both native Release test branches must pass before either production Release/ZIP branch starts. Scheduled dependency/security/status checks remain compiler-free. Formatting commits, extended qualification and ZIP/winget publication are explicit manual actions. PR checks do not satisfy native acceptance or waive applicable local scoped tests. Shared-engine pins in .ci/project.json and workflow callers must agree. Preserve dependency pins, native toolchain, evidence and desktop ownership contracts.

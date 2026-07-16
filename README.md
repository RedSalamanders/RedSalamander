# RedSalamander

RedSalamander is a Windows dual-pane file manager with:

- A fast DirectX-based folder view
- A modeless Find Files and Directories workflow with native, indexed, and fallback search backends
- Virtual file systems (archives, FTP/SFTP/SCP/IMAP, S3, …)
- Viewer plugins (Text/Hex, SQLite, Images/RAW, WebView2-based viewers, …)
- A themed Preferences experience (themes, plugins, shortcuts, associations)

![GitHub Downloads (all assets, all releases)](https://img.shields.io/github/downloads/RedSalamanders/RedSalamander/total?style=plastic)

![RedSalamander main window](docs/res/main-window.png)

[Complete User Documentation](docs/UserGuide.md)

## Why another file manager?

This project is a tribute to Servant Salamander and especially [Open Salamander](https://github.com/OpenSalamander/salamander).
Thanks to the Open Salamander contributors:

- David Andrš
- Lukáš Cerman
- Jakub Červený
- Tomáš Jelínek
- Milan Kaše
- Tomáš Kopal
- Jan Patera
- Martin Přikryl
- Juraj Rojko
- Jan Ryšavý
- Petr Šolín

You've done an amazing job.

Open Salamander appears to be quiet these days, so I wanted to start fresh with a modern, themeable, plugin-based successor.

RedSalamander is not yet at the level of other file managers. This is a work in progress.

Enjoy!

## For Developers

RedSalamander is a Windows-native C++23 application featuring advanced text visualization, real-time debugging (ETW), and high-performance graphics rendering (DirectX / Direct2D / DirectWrite).

Key engineering specs for the current search stack and self-test contract:

- `Specs/Core/Core_Search.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/Testing/Testing_SelfTests.md`

## Building the Project

### Prerequisites

- **Visual Studio 2026** with C++ toolset **v145** (Desktop development with C++)
- **vcpkg** (manifest mode) for dependencies
- **Windows 11 build 22000.2600 or later**

### Quick Start

#### Command Line Build

Use the `build.ps1` PowerShell script for easy building:

```powershell
# One-time: install dependencies (writes to .build\vcpkg_installed)
.\vcpkg-install.ps1

# Build entire solution in Debug configuration (default)
.\build.ps1

# Build in Release configuration
.\build.ps1 -Configuration Release

# Build for ARM64
.\build.ps1 -Platform ARM64

# Build specific project only
.\build.ps1 -ProjectName RedSalamanderMonitor
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName Common

# Clean build
.\build.ps1 -Clean

# Rebuild all
.\build.ps1 -Rebuild

# Combined options
.\build.ps1 -Configuration Release -ProjectName RedSalamanderMonitor -Rebuild
```

**Build Script Parameters:**

- `-Configuration` : `Debug` (default), `Release`, or `ASan Debug`
- `-Platform` : `x64` (default) or `ARM64`
- `-ProjectName` : Specific project name (builds entire solution if not specified)
- `-Clean` : Perform clean build
- `-Rebuild` : Rebuild all projects
- `-Msix` : Build an MSIX package after a successful Release build
- `-Msi` : Build an MSI package after a successful Release build

The build preflight closes an ordinary interactive instance only when it is
running from the exact output path being rebuilt. If that process is running a
self-test, or Windows cannot expose its command line safely, the build aborts
with the blocking PID, path, and command line instead of terminating the test.
Wait for self-tests to finish before rebuilding the same configuration.

### AddressSanitizer (`ASan Debug`)

Build an ASan-instrumented binary from the console with:

```powershell
.\build.ps1 -Configuration 'ASan Debug' -ProjectName RedSalamander
.\build.ps1 -Configuration 'ASan Debug' -ProjectName RedSalamanderSearchService
```

Outputs land in:

```text
.build\x64\ASan Debug\
.build\ARM64\ASan Debug\
```

Run them directly from the output folder:

```powershell
.\.build\x64\ASan Debug\RedSalamander.exe
.\.build\x64\ASan Debug\RedSalamanderSearchService.exe --run-foreground
```

The build now copies the required `clang_rt.asan_dynamic-*.dll` next to each ASan executable automatically. If you still see a missing ASan runtime DLL, rebuild after updating Visual Studio C++ tools and verify the ASan runtime exists under `$(VCToolsInstallDir)\bin\Host*\`.

#### Visual Studio Build

1. Open `RedSalamander.sln` in Visual Studio 2026
2. Select configuration (Debug/Release) and platform (x64/ARM64)
3. Build → Build Solution (Ctrl+Shift+B)

### Solution Structure

The solution contains the following projects:

- **Applications**: `RedSalamander`, `RedSalamanderMonitor`, `RedSalamanderSearchService`
- **Standalone tools**: `RedConfigure`, `RedLauncher`
- **Installer**: `RedSalamanderInstaller` (MSIX packaging)
- **Shared library**: `Common`
- **Plugins** (all under `Plugins\`):
  - **File-system plugins**: `FileSystem`, `FileSystem7z`, `FileSystemCurl`, `FileSystemS3`, `FileSystemGoogleDrive`, `FileSystemMicrosoftDrive`, `FileSystemDummy`
  - **Viewer plugins**: `ViewerText`, `ViewerSqlite`, `ViewerSpace`, `ViewerImgRaw`, `ViewerVLC`, `ViewerPE`, `ViewerWeb`
- **Tests**: 8 standalone test projects (e.g. `DxUiTests`, `PerformanceTests2`)
- **Tools**: PowerShell tooling (+ Pester tests in `Tools\Tests`)
- **PoC projects**: `ls1`, `ls2`, `ls3`, `ls4`, `FlipSequentialDiscard`, `MonitorTest`

The solution also contains per-language localization satellite projects, so Solution Explorer shows far more than the core projects listed above.

### Output

Built executables and libraries are located in:

```text
.build\x64\Debug\     (Debug builds)
.build\x64\Release\   (Release builds)
.build\ARM64\Debug\   (Debug builds)
.build\ARM64\Release\ (Release builds)
```

All build outputs and intermediate files are written under `.build\` to keep the source tree clean.

## Self-tests (Debug only)

RedSalamander includes three debug-only self-test suites:

- CompareDirectories self-test (`--compare-selftest`)
- Commands self-test (`--commands-selftest`)
- FileOperations self-test (`--fileops-selftest`)

### Run self-tests

Build a Debug binary and run the suites (recommended: run all suites in one process so the aggregated `results.json` includes everything):

```powershell
# Build Debug (x64)
.\build.ps1 -Configuration Debug -ProjectName RedSalamander

# Run all suites in one process (recommended)
.\.build\x64\Debug\RedSalamander.exe --selftest --selftest-timeout-multiplier=2.0

# Run suites individually
.\.build\x64\Debug\RedSalamander.exe --compare-selftest
.\.build\x64\Debug\RedSalamander.exe --commands-selftest
.\.build\x64\Debug\RedSalamander.exe --fileops-selftest

# Optional: fail-fast for CompareDirectories
.\.build\x64\Debug\RedSalamander.exe --compare-selftest --selftest-fail-fast
```

The process exit code is `0` on success and non-zero on failure.

Every declared self-test case must report `passed`, `failed`, or `skipped`. Conditional coverage stays part of the suite and uses `skipped` with a reason when prerequisites are absent. See `Specs/Testing/Testing_SelfTests.md`.

Note: the Commands self-test is UI-driven and may take longer to run.

Note: `RedSalamander.exe` is a GUI app, so PowerShell may return to the prompt immediately after launching it. Use `Start-Process -Wait` if you need to wait for completion and capture the exit code.

### Run all tests with one command

`Tools\Run-AllTests.ps1` is the canonical local test runner. It builds (unless `-SkipBuild`), runs the selected suite(s), prints a color pass/fail/skip summary, writes `run-all-tests-results.json`, `run-all-tests-case-history.jsonl`, and `run-all-tests-dashboard.md`, and exits with code `0` when everything is green.

```powershell
# The GitHub Actions PR gate, through the same unified runner CI uses:
.\Tools\Run-AllTests.ps1 -Suite CI

# Full local/closeout gate — builds the whole solution with tests enabled:
.\Tools\Run-AllTests.ps1 -Suite Full

# Just the three in-process selftest suites (default), reusing an existing Debug build:
.\Tools\Run-AllTests.ps1 -SkipBuild

# One suite, one case, on a slow machine:
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_ -TimeoutMultiplier 2.0

# Repeat one FileOps case twice and record a reproducible shuffle seed:
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -SelfTestRepeat 2 -SelfTestShuffleSeed 789 -TimeoutMultiplier 0.1
```

`-Suite CI` runs the PR gate through `Run-AllTests.ps1`: the three in-product self-test suites as separate processes, DxUiTests split per suite, FileSystemCurlTests, ViewerPETests plus the explicit prompt cases, ViewerSqliteTests, MonitorTest, LocalizationTests, RedConfigureTests, PerformanceTests2, the artifact-only `Tools\Tests` Pester pass, and the vcpkg-merge synthetic test.

`-Suite CI` also enables failure classification automatically. A standalone pass-on-rerun is reported as blocking `FLAKY`, and fail-again is blocking `REGRESSION`. For broad in-product suite failures, failed cases first rerun through `--selftest-case`; if those pass, the runner performs shuffle triage across three seeds. Pass-all-shuffle evidence becomes blocking `FLAKY`, fail-any-shuffle evidence becomes blocking `REGRESSION`, and missing shuffle evidence remains blocking `ISOLATION_SUSPECT`. Use `-ClassifyFailures` to enable the same retry evidence on non-CI suites.

Debug self-test builds expose `-SelfTestFlakyProofCase NAME` / `--selftest-flaky-proof-case=NAME` and `-SelfTestOrderProofCase NAME` / `--selftest-order-proof-case=NAME` only to prove the classifier path: flaky proof must classify as blocking `FLAKY`, and order-dependent proof must classify as blocking `REGRESSION` after shuffle triage.

`-SelfTestRepeat N` forwards `--selftest-repeat=N` to in-product self-test suites and repeat result rows include `repeat_index`; Commands and CompareDirectories repeat through the shared case runner, and FileOperations expands its phase-state run plan per repeat attempt. `-SelfTestShuffleSeed SEED` forwards `--selftest-shuffle=SEED` to Commands, CompareDirectories, and FileOperations. FileOperations uses explicit seeded phase ordering when a shuffle seed is supplied, while preserving `Setup` and `Cleanup_RestorePluginConfig` rows in the aggregate result.

The expensive repeat+shuffle lane lives in `.github\workflows\nightly-flake.yml`, not the PR gate. It runs `-Suite All -SelfTestRepeat 5 -SelfTestShuffleSeed <seed> -ClassifyFailures` on schedule or manual dispatch and uploads `selftest-artifacts-nightly-shuffle`.

Known flaky tests are tracked only through `Tools\test-quarantine.jsonl`. Each JSONL entry must name the harness/case, owner, opened/expires dates, issue/spec link, root-cause hypothesis, and fix-or-replace plan. Active or invalid entries keep the runner red. Active entries that match a runner harness adapter execute again in a separate repair lane, and `run-all-tests-results.json` records the repair attempt evidence. In GitHub Actions, the runner also appends classification counts, active quarantine owner/expiry, and repair-lane results to `GITHUB_STEP_SUMMARY`.

Self-test crashes remain failures, but they should not erase evidence: suite `results.json` is flushed after every case, and an in-flight crash is written as `status: "crashed"` in partial aggregate results for the runner to report. Debug self-test builds expose `--selftest-crash-case=NAME` only for proving that crash-signal path.

`-Suite Full` builds the full solution (test projects included) and additionally runs the broader closeout surface, including PluginContractTests, SettingsSchemaTests, CrashHandlingTests, and RedSalamanderMonitorEtwLatency. Results land under `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run\`; by default the runner sets `REDSALAMANDER_TEST_ROOT` to `.build\TestSandbox`, sets `REDSALAMANDER_TEST_RUN_ID`, and ignores/clears inherited `REDSALAMANDER_SELFTEST_ROOT` values for normal runs.

Before launching child tests, the runner invokes `Tools\Clean-TestSandbox.ps1 -Apply -Confirm:$false` to remove known legacy self-test/temp roots that predate the unified sandbox. Run `Tools\Clean-TestSandbox.ps1` without `-Apply` for a dry-run listing, or pass `-SkipLegacySandboxCleanup` to the runner only for diagnosis.

The runner defaults to Debug configuration; self-tests only run in Debug builds.

### Self-test artifacts and results

When launched through `Tools\Run-AllTests.ps1`, self-test output is written under:

```text
<repoRoot>\.build\TestSandbox\runs\<runId>\artifacts\selftest\last_run\
```

The runner writes a per-invocation aggregate summary, and native self-tests resolve their `last_run` writer directly from `REDSALAMANDER_TEST_ROOT` plus `REDSALAMANDER_TEST_RUN_ID`:

- `<repoRoot>\.build\TestSandbox\runs\<runId>\artifacts\selftest\last_run\`
- `<repoRoot>\.build\TestSandbox\runs\<runId>\artifacts\selftest\last_run\run-all-tests-results.json`
- `<repoRoot>\.build\TestSandbox\runs\<runId>\artifacts\selftest\last_run\run-all-tests-case-history.jsonl`
- `<repoRoot>\.build\TestSandbox\runs\<runId>\artifacts\selftest\last_run\run-all-tests-dashboard.md`

Key files:

- `trace.txt` (host trace)
- `compare\trace.txt`
- `compare\results.json`
- `commands\trace.txt`
- `commands\results.json`
- `fileops\trace.txt`
- `fileops\results.json`
- `run-all-tests-results.json` (runner-owned aggregate summary)
- `run-all-tests-case-history.jsonl` (per-case and retry history rows)
- `run-all-tests-dashboard.md` (slow-case and failure/retry dashboard)

## Search Service (Developer)

`RedSalamanderSearchService.exe` supports terminal-friendly management commands.

Default identities:

- Debug: `RedSalamanderSearchService.Debug` on `\\.\pipe\RedSalamander.SearchService.Debug.v3`
- Release: `RedSalamanderSearchService` on `\\.\pipe\RedSalamander.SearchService.v3`

Useful commands:

```powershell
# Show all supported options
.\.build\x64\Debug\RedSalamanderSearchService.exe --help

# Run in the current terminal
.\.build\x64\Debug\RedSalamanderSearchService.exe --run-foreground

# Run offline SQLite maintenance
.\.build\x64\Debug\RedSalamanderSearchService.exe --compact

# Ask a running SQLite-backed service to compact in place
.\.build\x64\Debug\RedSalamanderSearchService.exe --request-compact

# Register / unregister the Windows service for the current build
.\.build\x64\Debug\RedSalamanderSearchService.exe --register
.\.build\x64\Debug\RedSalamanderSearchService.exe --unregister
```

Notes:

- `--register` and `--unregister` typically require an elevated terminal.
- `--compact` targets the SQLite store, acquires the single-instance guard, truncates WAL, runs `VACUUM`, and exits with a before/after space summary. Stop any running service instance first.
- `--request-compact` asks the running service to run live SQLite maintenance in-process and prints the refreshed DB/WAL/free-page state after the request completes.
- `--run-foreground` now auto-attaches to the parent terminal when needed, prints a startup banner with PID/build/mode details, shows a live status dashboard in a console, and falls back to readable lifecycle log lines when output is redirected.
- Press `Ctrl+C` to stop foreground mode cleanly.
- `--pipe-name=...` can target a non-default running service for `--request-compact`, and `--protocol-version=...`, `--max-requests=...`, and `--disconnect-after-batches=...` are intended for foreground-mode development and self-tests.
- `--storage-root=...`, `--store-backend=...`, and `--sqlite-path=...` can be used with `--run-foreground` or `--compact`.
- `--storage-root=...` overrides the snapshot/storage root, which is useful for isolated self-tests and local service debugging.

### MSIX Installer

Build Release + MSIX in one command:

```powershell
.\build.ps1 -Msix
```

Or build/package separately:

```powershell
.\build.ps1 -Configuration Release
msbuild Installer\msix\RedSalamanderInstaller.wapproj /p:Configuration=Release /p:Platform=x64
```

MSIX output is written to:

```text
.build\AppPackages\
```

See `Specs/Installer/Installer_Msix.md` for signing and deployment details.

### MSI Installer

Build Release + MSI in one command (requires WiX Toolset v6+):

```powershell
.\build.ps1 -Msi
```

MSI output is written to:

```text
.build\AppPackages\
```

See `Specs/Installer/Installer_Msi.md` for details.

## vcpkg

### Installing dependencies

Install all libraries from `vcpkg.json` into `.build\vcpkg_installed`:

```powershell
.\vcpkg-install.ps1

# ARM64:
.\vcpkg-install.ps1 -Platform ARM64
```

### (Optional) Enable user-wide MSBuild integration

If your Visual Studio/MSBuild setup does not pick up vcpkg manifest dependencies automatically, enable vcpkg's MSBuild integration:

```powershell
vcpkg.exe integrate install
```

### Adding New Libraries

To add a new library:

```powershell
vcpkg.exe add port <library-name>
```

Then re-run `.\vcpkg-install.ps1`.

### Using vcpkg in Visual Studio

1. Open Visual Studio
2. Open the project you want to use vcpkg with
3. Ensure vcpkg integration is enabled:
   - Right-click on the project in Solution Explorer
   - Select "Properties"
   - Under "Configuration Properties", check if "Vcpkg" is listed
   - If listed, vcpkg integration is enabled

## Technology Stack

- **Language**: C++23
- **Build System**: MSBuild / Visual Studio 2026 (v145)
- **Package Manager**: vcpkg
- **Graphics**: Direct2D, DirectWrite, Direct3D 11, DXGI
- **Platform**: Windows 11 build 22000.2600 or later (Unicode UTF-16)

## ETW Tracing (RedSalamanderMonitor)

[RedSalamanderMonitor](RedSalamanderMonitor/) is a real-time ETW (Event Tracing for Windows) viewer for RedSalamander events.
On some machines it may need extra privileges to start its ETW listener.

### One-time permission setup (avoid UAC prompts)

Run the helper script once to add your account to the local **Performance Log Users** group (the script will self-elevate):

```powershell
.\init-etw-trace.ps1
```

Then **sign out/in (or reboot)** so your access token picks up the new group membership, and launch:

- `.build\x64\Debug\RedSalamanderMonitor.exe` (Debug), or
- `.build\x64\Release\RedSalamanderMonitor.exe` (Release)

To undo the change later:

```powershell
.\init-etw-trace.ps1 -Remove
```

### Optional: capture an `.etl` trace file

If you need an external ETW session (e.g. Windows Performance Analyzer), use:

- `.\start-etw-trace.ps1`
- `.\stop-etw-trace.ps1`
- `.\clean-etw-trace.ps1`

Note: `RedSalamanderMonitor.exe` has its own built-in ETW listener and does not require an external session for normal use.

Normal Release builds keep Info/Perf/debug-style ETW diagnostics quiet, and RedSalamanderMonitor filters out its own ETW messages. To build a dedicated Release set that emits and displays those diagnostics in RedSalamanderMonitor, use:

```powershell
.\build.ps1 -Configuration Release -MonitorDiagnostics
```

## Additional Documentation

- **docs/**: User and developer documentation (start at `docs/UserGuide.md` or `docs/DeveloperGuide.md`)
- `docs/RemoteFileSystems.md`: Remote file systems (FTP/SFTP/SCP/IMAP)
- `docs/DxUi.md`: Shared DirectX UI developer guide
- **AGENTS.md**: Comprehensive development guidelines for AI assistants and developers
- **.github/copilot-instructions.md**: GitHub Copilot specific guidelines
- **Specs/**: Detailed specifications for various components

## License

See `LICENSE.txt` for license information.

# Developer Guide

This guide is the developer entry point for the public documentation set. The
authoritative engineering rules still live in [../AGENTS.md](../AGENTS.md) and
the durable subsystem contracts live under `Specs/`.

## Repository layout

| Path | Purpose |
| --- | --- |
| `RedSalamander/` | Main dual-pane shell, dialogs, settings, command routing, and app theming. |
| `Common/` | Shared libraries, plugin interfaces, settings helpers, Win32 helpers, and `Common/DxUi`. |
| `Plugins/` | Built-in file-system and viewer plugins. |
| `Tests/` | Deterministic selftests and component test executables, including `Tests/DxUiTests`. |
| `Specs/` | Authoritative product, UI, file-system, testing, and implementation-plan specs. |
| `docs/` | User and developer documentation that can be published with GitHub Pages. |

## Build and test

Use the repository build wrapper from the repo root:

```powershell
.\build.ps1
.\build.ps1 -Configuration Release
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName DxUi
.\build.ps1 -ProjectName DxUiTests
```

Useful outputs:

- `.build\x64\Debug\RedSalamander.exe`
- `.build\x64\Debug\RedSalamanderMonitor.exe`
- `.build\x64\Debug\DxUiTests.exe`
- `.build\x64\Release\DxUiTests.exe`

Run focused DxUi suites while working on the shared UI layer:

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=WindowHost
.\build\x64\Debug\DxUiTests.exe --suite=Control
.\build\x64\Debug\DxUiTests.exe --suite=Grid
.\build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
.\build\x64\Debug\DxUiTests.exe --suite=Accessibility
```

Generate the public theme/control gallery screenshots:

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=Gallery --gallery-output-directory=docs\res
.\build\x64\Debug\DxUiTests.exe --suite=ButtonContrast --button-audit-output=docs\res\theme-button-states-after-fix.png
```

## Development rules to remember

- Use WIL RAII wrappers for Windows resources and COM pointers.
- Keep shared UI work on the UI thread unless a spec explicitly says otherwise.
- Treat perf validation as part of the feature, not a later cleanup.
- Localized user-facing strings belong in resources with positional
  `std::format` placeholders.
- Update durable specs and the relevant `docs/` page before closing a user-facing
  change.
- Do not leave completed behavior only in `Specs/Plans/WIP/` or
  `Specs/Plans/Done/`.

## Technical guides

- [DxUi Technical Guide](DxUi.md) explains the shared DirectX UI layer, host
  lifecycle, theme/background model, controls, examples, and tests.
- [Winget Integration](WingetIntegration.md) explains package manifest
  generation and validation.
- [Documentation Coverage Map](DocumentationMap.md) tracks what the public docs
  cover and which areas are still thin.
- [Monitor.md](Monitor.md) introduces RedSalamanderMonitor for ETW/log
  inspection.

## Specs worth reading first

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_VisualStyle.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/Testing/Testing_TestCoverage.md`

Use specs for normative behavior and `docs/` for onboarding, workflows, and
public explanations.

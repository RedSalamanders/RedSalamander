# Specs

This folder contains **engineering specs**, **WIP plans**, **completed plan history**, **RFCs**, and **test artifacts**.

## What to read

- **Authoritative specs** (behavior/contracts to follow): `Specs/<Domain>/*.md`
- **Plans (WIP)** (active implementation checklists): `Specs/Plans/WIP/`
- **Plans (Done)** (completed plans kept for sequencing history and closeout evidence): `Specs/Plans/Done/`
- **RFCs** (design proposals): `Specs/Plans/WIP/RFC_*.md` (move to `Done/` once implemented)

## Pinned paths (do not move)

These locations are referenced by the build/runtime:

- `Specs/SettingsStore.schema.json`
- `Specs/Themes/`
- `Specs/TestRuns/`

## Naming conventions

- Domain specs: `Domain_Title.md` (no spaces)
- RFCs: `RFC_Domain_Title.md`
- Plans: `Domain_Title.md` under `Specs/Plans/WIP/` (move to `Specs/Plans/Done/` when completed).

## Hard rules

- Any new feature work must have a plan in `Specs/Plans/WIP/`.
- When a plan is finished, move it to `Specs/Plans/Done/` and merge any enduring behavior, UI contract, verification rule, or workflow requirement into the authoritative domain spec under `Specs/<Domain>/` (or repo-level guidance such as `AGENTS.md` / `Specs/Testing/*` when appropriate).
- RFCs must live in `Specs/Plans/WIP/` (or `Done/`) and start with `RFC_`.
- `Specs/Plans/Done/` is historical and non-authoritative. Use it for rollout history and closeout evidence, not as the current source of truth for behavior contracts.

## Start here

### Core

- `Specs/Core/Core_SharedHelpers.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Core/Core_Search.md`
- `Specs/Core/Core_ConnectionManager.md`
- `Specs/Core/Core_StartupBootstrap.md`

### UI

- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_FolderWindow.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_ManagePluginsDialog.md`
- `Specs/UI/UI_TopLevelToolWindows.md`
- `Specs/UI/UI_PreferencesDialog.md`

### FileSystem

- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/FileSystem/FileSystem_FtpSftpScp.md`
- `Specs/Plans/Done/FileSystem_RemediationPlan_2026-02-26.md`

### Plugins

- `Specs/Plugins/Plugins_PluginAPI.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Plugins/Plugins_ViewerPlugins.md`
- `Specs/Plugins/Plugins_ViewerSqlite.md`

### Installer

- `Specs/Installer/Installer_Msix.md`
- `Specs/Installer/Installer_Msi.md`

### Testing

- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`
- `Specs/TestRuns/README.md`

# Specs

This folder contains **engineering specs**, **WIP plans**, **RFCs**, and **test artifacts**.

## What to read

- **Authoritative specs** (behavior/contracts to follow): `Specs/<Domain>/*.md`
- **Plans (WIP)** (active implementation checklists): `Specs/Plans/WIP/`
- **Plans (Done)** (completed plans kept for history): `Specs/Plans/Done/`
- **RFCs** (design proposals): `Specs/Plans/WIP/RFC_*.md` (move to `Done/` once implemented)
- **Notes/Scratch** (non-normative): `Specs/Notes/`

## Pinned paths (do not move)

These locations are referenced by the build/runtime:

- `Specs/SettingsStore.schema.json`
- `Specs/Themes/`
- `Specs/TestRuns/`

## Naming conventions

- Domain specs: `Domain_Title.md` (no spaces)
- RFCs: `RFC_Domain_Title.md`
- Plans: `Domain_Title.md` under `Specs/Plans/WIP/` (move to `Specs/Plans/Done/` when completed)

## Hard rules

- Any new feature work must have a plan in `Specs/Plans/WIP/` and be moved to `Specs/Plans/Done/` when finished.
- RFCs must live in `Specs/Plans/WIP/` (or `Done/`) and start with `RFC_`.
- Notes in `Specs/Notes/` are **not** authoritative unless explicitly promoted into a domain spec.

## Start here

### Core
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Core/Core_ConnectionManager.md`
- `Specs/Core/Core_StartupBootstrap.md`

### UI
- `Specs/UI/UI_FolderWindow.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/UI/UI_PreferencesDialog.md`

### FileSystem
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/FileSystem/FileSystem_FtpSftpScp.md`
- `Specs/Plans/WIP/FileSystem_RemediationPlan_2026-02-26.md`

### Plugins
- `Specs/Plugins/Plugins_PluginAPI.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Plugins/Plugins_ViewerPlugins.md`

### Installer
- `Specs/Installer/Installer_Msix.md`
- `Specs/Installer/Installer_Msi.md`

### Testing
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`
- `Specs/TestRuns/README.md`


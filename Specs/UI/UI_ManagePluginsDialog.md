# Manage Plugins Dialog Specification

Last updated: 2026-03-28

## Purpose

This document is the authoritative contract for the reusable Manage Plugins dialog that edits a plugin's schema-driven configuration payload.

Shared `DxUi` hosting, accessibility, visible-native retirement, and migrated-window acceptance rules live in `Specs/UI/UI_DxUiSharedGrid.md`.

## Scope

This specification applies to:

- `RedSalamander/ManagePluginsDialog.cpp`
- `RedSalamander/ManagePluginsDialog.h`

Related specs:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_PreferencesDialog.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Core/Core_SettingsStore.md`

## Dialog Contract

- The dialog is a modal, host-owned configuration editor for a single plugin.
- The dialog title MUST identify the target plugin.
- The visible surface consists of:
  - a scrollable schema form body,
  - section labels and field descriptions,
  - schema-driven controls,
  - `OK` and `Cancel` command buttons.
- `OK` MUST validate and commit the edited configuration payload for the target plugin.
- `Cancel` MUST close without committing the in-dialog edits.
- The scrollable form body MUST remain usable at small window sizes through wheel scrolling, scrollbar interaction, and focus movement into off-screen fields.

## Schema-Driven Field Contract

- The dialog renders plugin schema fields using the plugin configuration schema as the source of truth.
- The current supported field families are:
  - text,
  - numeric value,
  - boolean,
  - single-choice option,
  - multi-selection choice groups.
- `x-ui-section` controls grouping under section headers.
- `x-ui-order` controls stable display order inside a section.
- `x-ui-hidden: true` keeps a field in the JSON payload but MUST suppress it from the rendered editor.
- Unknown or unsupported field metadata MAY fall back to the existing JSON-preserving behavior, but the dialog MUST NOT silently drop persisted values.

## Persistence And Reload Contract

- Committing the dialog MUST update the owning settings model and persist through the existing settings-save path.
- The dialog MUST surface save failures to the user instead of failing silently.
- If the underlying settings source changes externally while the dialog is open, the user MUST be offered an explicit reload-or-keep-edit decision before stale data is overwritten.

## DXUI Contract

- The current live dialog keeps its visible schema-form statics, interactive inputs, and command row on the shared `DxUi` path.
- Any remaining hidden legacy helpers are temporary backing scaffolding only and are not accepted visible end-state UI.
- Accessibility and visible-child-fallback rules for this dialog are normative through `Specs/UI/UI_DxUiSharedGrid.md`.

## Verification Requirements

Changes to this dialog MUST keep the following behavior green:

- the schema form opens on the shared DX path with no accepted visible native fallback,
- scrolling remains stable for large schemas,
- visible inputs expose the expected UI Automation patterns,
- repeated open/close churn stays clean across both `OK` and `Cancel` flows,
- stale-settings reload prompts preserve either the reloaded source state or the kept in-flight edits according to the user's choice.

# Preview Properties Default Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the plain no-viewer preview text with a scrollable, card-based file/folder properties preview that matches the Properties dialog and gains subtle rainbow treatment in Rainbow theme.

**Architecture:** Keep embedded preview viewers unchanged. When no embedded viewer matches, parse the existing `IFileSystemIO::GetItemProperties` JSON into the existing properties document model, build DxUi cards inside the preview content host, and expose deterministic debug state for layout, scrollbar, and rainbow validation. Keep the old plain text fallback only for unsupported/non-properties failures.

**Tech Stack:** C++23, Win32, WIL, Direct2D/DirectWrite-backed DxUi, RedSalamander command self-tests.

---

### Task 1: Red Test and Debug Surface

**Files:**
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`

- [x] Add `PreviewPaneDebugSnapshot` fields for properties-card mode, section count, field count, scrollbar presence, scroll offset, and rainbow mode.
- [x] Add a `DebugScrollPreviewPropertiesByWheelDetents(...)` helper so tests can exercise the preview scrollbar without stealing source-pane focus.
- [x] Add a command self-test named `pane_view_options_preview_properties_card_scrolls_and_uses_rainbow_theme`.
- [x] Run the new self-test and confirm it fails because the current preview is still a plain label without a properties scroll panel.

### Task 2: Properties Preview UI

**Files:**
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/FolderWindow.cpp`
- Modify: `RedSalamander/FolderWindow.Layout.cpp`
- Modify: `RedSalamander/FolderWindow.ItemProperties.cpp`

- [x] Create the preview content root with both the legacy fallback label and a hidden `DxUi::ScrollPanel`.
- [x] Add preview card control state to `PaneState`.
- [x] Add `ShowPreviewPropertiesForPath(...)` that loads properties JSON, parses it with the existing item-properties parser, builds compact cards, sets `previewText` for copy/debug parity, and returns `false` to the legacy fallback when properties are unavailable.
- [x] Add layout code that positions card section titles, key/value rows, wraps long values, and sets `ScrollPanel::SetContentHeight(...)`.
- [x] Use the existing Rainbow theme flag to color section headers with a restrained hue cycle, while high-contrast mode keeps system-safe colors.
- [x] Reset scroll position on focused item changes and preserve source-pane keyboard focus.

### Task 3: Specs and Verification

**Files:**
- Modify: `Specs/UI/UI_FolderWindow.md`
- Modify: `Specs/UI/UI_CommandMenuKeyboard.md`
- Modify: `Docs/Viewers.md`
- Modify: `Docs/UserGuide.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] Document that default preview is card-based, scrollable, theme-aware, and rainbow-accented in Rainbow theme.
- [x] Build `RedSalamander`.
- [x] Run the new red/green self-test, the grouped `pane_view_options_preview_` tests, and the preview tab selection test.
- [x] Archive perf/correctness evidence under `Specs/TestRuns/`.
- [x] Move this plan to `Specs/Plans/Done/` after the implementation and spec updates are verified.

#pragma once

#ifdef ENABLE_TESTS

#include <cstdint>
#include <string>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

// Debug-only hook for automation/self-tests: dispatch a command by its canonical command id (e.g., "cmd/pane/refresh").
[[nodiscard]] bool DebugDispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept;

// Debug-only hook for menu-contract self-tests: returns the Fluent icon glyph for a menu command, or 0 when none is assigned.
[[nodiscard]] wchar_t DebugGetMainMenuIconGlyph(UINT menuCommandId) noexcept;

// Debug-only hook for menu-contract self-tests: returns the retained main-menu model even when the native menu is detached.
[[nodiscard]] HMENU DebugGetMainMenuModelHandle() noexcept;

// Debug-only hook for menu-bar self-tests: reports whether the DxUI menu-bar surface is currently visible for the main window.
[[nodiscard]] bool DebugIsMainMenuBarSurfaceVisible(HWND mainWindow) noexcept;

// Debug-only hook for menu-bar self-tests: returns the selected top-level menu index, or -1 when nothing is selected.
[[nodiscard]] int DebugGetMainMenuBarSelectedIndex() noexcept;

// Debug-only hook for menu-bar self-tests: returns how many top-level menu-bar items are visually highlighted.
[[nodiscard]] int DebugGetMainMenuBarVisualHighlightCount() noexcept;

// Debug-only hook for menu-bar self-tests: returns the current retained menu-bar render count.
[[nodiscard]] uint64_t DebugGetMainMenuBarRenderCount() noexcept;

// Debug-only hook for menu-bar self-tests: returns the retained label text for a top-level menu-bar item.
[[nodiscard]] bool DebugGetMainMenuBarItemLabel(size_t index, std::wstring& outText) noexcept;

// Debug-only hook for menu-bar self-tests: returns the raw HMENU source index backing a retained top-level menu-bar item.
[[nodiscard]] bool DebugGetMainMenuBarItemSourceIndex(size_t index, size_t& outSourceIndex) noexcept;

// Debug-only hook for menu-bar self-tests: returns the screen rect of a visible top-level menu-bar item.
[[nodiscard]] bool DebugGetMainMenuBarItemScreenRect(HWND mainWindow, size_t index, RECT& rectPx) noexcept;

// Debug-only hook for menu-bar self-tests: hit-tests a screen point against the retained DxUI menu bar.
[[nodiscard]] bool DebugHitTestMainMenuBarScreenPoint(HWND mainWindow, POINT screenPoint, size_t& outIndex) noexcept;

#endif // ENABLE_TESTS

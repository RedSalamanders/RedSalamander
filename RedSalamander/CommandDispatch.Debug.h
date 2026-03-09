#pragma once

#ifdef _DEBUG

#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

// Debug-only hook for automation/self-tests: dispatch a command by its canonical command id (e.g., "cmd/pane/refresh").
[[nodiscard]] bool DebugDispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept;

// Debug-only hook for menu-contract self-tests: returns the Fluent icon glyph for a menu command, or 0 when none is assigned.
[[nodiscard]] wchar_t DebugGetMainMenuIconGlyph(UINT menuCommandId) noexcept;

#endif // _DEBUG

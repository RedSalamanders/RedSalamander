#pragma once

#ifdef _DEBUG

#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

// Debug-only hook for automation/self-tests: dispatch a command by its canonical command id (e.g., "cmd/pane/refresh").
[[nodiscard]] bool DebugDispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept;

#endif // _DEBUG


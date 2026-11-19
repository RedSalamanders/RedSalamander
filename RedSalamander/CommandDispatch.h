#pragma once

#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "Ui/ThemeCycleOverlayWindow.h"

// Dispatches a canonical application command through the same runtime-state and
// implementation path used by shortcuts, menus, and the command palette.
[[nodiscard]] bool DispatchApplicationCommand(HWND ownerWindow, std::wstring_view commandId, RedSalamander::Ui::CommandInvocationSource source) noexcept;

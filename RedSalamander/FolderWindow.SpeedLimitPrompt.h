#pragma once

#include <cstdint>
#include <string>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

enum class SpeedLimitPromptDebugFocusTarget : uint8_t
{
    None,
    ValueField,
    OkButton,
    CancelButton,
};

struct SpeedLimitPromptDebugSnapshot
{
    bool usesDxUiHost                            = false;
    size_t visibleChildWindowCount               = 0u;
    SpeedLimitPromptDebugFocusTarget focusTarget = SpeedLimitPromptDebugFocusTarget::None;
    std::wstring valueText;
    std::wstring validationText;
};

[[nodiscard]] HWND GetSpeedLimitPromptDialogHandle() noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugGetSpeedLimitPromptSnapshot(SpeedLimitPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetSpeedLimitPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmSpeedLimitPrompt() noexcept;
#endif

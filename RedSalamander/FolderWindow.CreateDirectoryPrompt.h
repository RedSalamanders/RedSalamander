#pragma once

#include <cstdint>
#include <string>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

enum class CreateDirectoryPromptDebugFocusTarget : uint8_t
{
    None,
    NameField,
    OkButton,
    CancelButton,
};

struct CreateDirectoryPromptDebugSnapshot
{
    bool usesDxUiHost                                 = false;
    size_t visibleChildWindowCount                    = 0u;
    CreateDirectoryPromptDebugFocusTarget focusTarget = CreateDirectoryPromptDebugFocusTarget::None;
    std::wstring pathText;
    std::wstring nameText;
    std::wstring validationText;
};

[[nodiscard]] HWND GetCreateDirectoryPromptDialogHandle() noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugGetCreateDirectoryPromptSnapshot(CreateDirectoryPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetCreateDirectoryPromptName(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmCreateDirectoryPrompt() noexcept;
#endif

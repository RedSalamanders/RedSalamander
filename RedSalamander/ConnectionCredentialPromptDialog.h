#pragma once

#include <cstdint>
#include <string>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "AppTheme.h"

// Themed prompt for a connection secret (password/passphrase).
// Returns:
// - S_OK: secretOut set (may be empty if allowEmptySecret == true)
// - S_FALSE: user cancelled (secretOut cleared)
// - failure HRESULT: unexpected error
HRESULT PromptForConnectionSecret(HWND ownerWindow,
                                  const AppTheme& theme,
                                  std::wstring_view caption,
                                  std::wstring_view message,
                                  std::wstring_view secretLabel,
                                  bool allowEmptySecret,
                                  std::wstring& secretOut) noexcept;

// Themed prompt for a user name + password (FTP anonymous rejected).
// Returns:
// - S_OK: userNameOut + passwordOut set
// - S_FALSE: user cancelled (outputs cleared)
// - failure HRESULT: unexpected error
HRESULT PromptForConnectionUserAndPassword(HWND ownerWindow,
                                           const AppTheme& theme,
                                           std::wstring_view caption,
                                           std::wstring_view message,
                                           std::wstring_view initialUserName,
                                           std::wstring& userNameOut,
                                           std::wstring& passwordOut) noexcept;

[[nodiscard]] HWND GetConnectionCredentialPromptDialogHandle() noexcept;
void UpdateConnectionCredentialPromptWindowsTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
enum class ConnectionCredentialPromptDebugFocusTarget : uint8_t
{
    None,
    UserField,
    SecretField,
    ToggleSecretButton,
    OkButton,
    CancelButton,
};

struct ConnectionCredentialPromptDebugSnapshot
{
    bool usesDxUiHost                                      = false;
    bool showUserName                                      = false;
    bool allowEmptySecret                                  = false;
    bool secretVisible                                     = false;
    bool themeDark                                         = false;
    bool themeHighContrast                                 = false;
    bool themeRainbow                                      = false;
    size_t secretLength                                    = 0u;
    size_t visibleChildWindowCount                         = 0u;
    ConnectionCredentialPromptDebugFocusTarget focusTarget = ConnectionCredentialPromptDebugFocusTarget::None;
    std::wstring userNameText;
    std::wstring validationText;
};

[[nodiscard]] bool DebugGetConnectionCredentialPromptSnapshot(ConnectionCredentialPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetConnectionCredentialPromptUserName(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetConnectionCredentialPromptSecret(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugGetConnectionCredentialPromptToggleSecretButtonHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugToggleConnectionCredentialPromptSecretVisibility() noexcept;
[[nodiscard]] bool DebugConfirmConnectionCredentialPrompt() noexcept;
#endif

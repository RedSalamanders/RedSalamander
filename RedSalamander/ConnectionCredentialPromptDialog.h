#pragma once

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

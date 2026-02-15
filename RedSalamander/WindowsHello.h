#pragma once

#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace RedSalamander::Security
{
// Returns:
// - S_OK: verified
// - HRESULT_FROM_WIN32(ERROR_CANCELLED): user cancelled / not verified
// - HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED): Windows Hello unavailable
// - other failures as HRESULT
HRESULT VerifyWindowsHelloForWindow(HWND ownerWindow, std::wstring_view message) noexcept;
} // namespace RedSalamander::Security

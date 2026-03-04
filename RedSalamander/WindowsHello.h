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

#ifdef _DEBUG
using WindowsHelloTestVerifier = HRESULT (*)(HWND ownerWindow, std::wstring_view message) noexcept;
// Sets a test hook for Windows Hello verification. Returns the previously installed verifier (if any).
// When installed, the verifier is called instead of the real Windows Hello flow.
WindowsHelloTestVerifier SetWindowsHelloTestVerifier(WindowsHelloTestVerifier verifier) noexcept;
#endif
} // namespace RedSalamander::Security

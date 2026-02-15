#pragma once
#include <chrono>
#include <string_view>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace SplashScreen
{
void BeginDelayedOpen(std::chrono::milliseconds delay, HINSTANCE instance) noexcept;
void CloseIfExist() noexcept;
bool Exist() noexcept;
HWND GetHwnd() noexcept;
void SetOwner(HWND owner) noexcept;
void IfExistSetText(std::wstring_view text) noexcept;
}; // namespace SplashScreen

#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include <chrono>
#include <stop_token>
#include <string_view>

namespace RedConfigure::SplashScreen
{
namespace Detail
{
[[nodiscard]] inline bool ShouldAbortPendingOpen(std::stop_token stopToken, HANDLE closeEvent) noexcept
{
    if (stopToken.stop_requested())
    {
        return true;
    }

    return closeEvent && ::WaitForSingleObject(closeEvent, 0) == WAIT_OBJECT_0;
}
} // namespace Detail

void BeginOpen(std::chrono::milliseconds delay, HINSTANCE instance, std::wstring_view initialStatus) noexcept;
void BeginImmediateOpen(HINSTANCE instance, std::wstring_view initialStatus) noexcept;
void RequestCloseIfExist() noexcept;
void CloseIfExist() noexcept;
bool Exist() noexcept;
HWND GetHwnd() noexcept;
void SetOwner(HWND owner) noexcept;
void IfExistSetText(std::wstring_view text) noexcept;

#ifdef ENABLE_TESTS
struct DebugSnapshot
{
    bool threadStarted      = false;
    bool hasHwnd            = false;
    unsigned long stage     = 0;
    unsigned long lastError = 0;
    long comHr              = 0;
};

[[nodiscard]] DebugSnapshot DebugGetSnapshot() noexcept;
#endif
} // namespace RedConfigure::SplashScreen

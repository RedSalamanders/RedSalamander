#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <filesystem>
#include <format>
#include <fstream>
#include <iostream>
#include <memory>
#include <mutex>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>

#include <UIAutomation.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "ViewerText/resource.h"
#include "WindowMessages.h"

namespace
{
using namespace std::chrono_literals;
using RedSalamanderCreateFn = HRESULT(__stdcall*)(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);

constexpr wchar_t kViewerPEWindowClassName[]         = L"RedSalamander.ViewerPE";
constexpr wchar_t kViewerWebWindowClassName[]        = L"RedSalamander.ViewerWeb";
constexpr wchar_t kViewerImgRawWindowClassName[]     = L"RedSalamander.ViewerImgRaw";
constexpr wchar_t kViewerSpaceWindowClassName[]      = L"RedSalamander.ViewerSpace";
constexpr wchar_t kViewerVLCWindowClassName[]        = L"RedSalamander.ViewerVLC";
constexpr wchar_t kViewerTextWindowClassName[]       = L"RedSalamander.ViewerText";
constexpr wchar_t kViewerTextPromptWindowClassName[] = L"RedSalamander.ViewerText.Prompt";
constexpr wchar_t kDxNativeMenuBarWindowClassName[]  = L"RedSalamander.DxNativeMenuBar";
constexpr wchar_t kViewerPEPluginId[]                = L"builtin/viewer-pe";
constexpr wchar_t kViewerWebPluginId[]               = L"builtin/viewer-web";
constexpr wchar_t kViewerImgRawPluginId[]            = L"builtin/viewer-imgraw";
constexpr wchar_t kViewerTextPluginId[]              = L"builtin/viewer-text";
constexpr wchar_t kViewerSpacePluginId[]             = L"builtin/viewer-space";
constexpr wchar_t kViewerVlcPluginId[]               = L"builtin/viewer-vlc";
constexpr int kViewerPEFileComboId                   = 1001;
constexpr int kViewerWebFileComboId                  = 1001;
constexpr int kViewerImgRawFileComboId               = 2001;
constexpr int kViewerTextFileComboId                 = 1003;
constexpr UINT kViewerTextFindCommandId              = 40101u;
constexpr UINT kViewerTextGotoCommandId              = 40203u;

struct UiaHostPatternStats
{
    size_t visibleElementCount    = 0u;
    size_t comboBoxControlCount   = 0u;
    size_t editControlCount       = 0u;
    size_t buttonControlCount     = 0u;
    bool hasValuePattern          = false;
    bool hasExpandCollapsePattern = false;
    std::wstring comboName;
    std::wstring valueText;
};

void PumpPendingMessages() noexcept
{
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

template <typename TPredicate> [[nodiscard]] bool PumpUntil(TPredicate&& predicate, std::chrono::milliseconds timeout) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (predicate())
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return predicate();
}

[[nodiscard]] std::vector<HWND> CollectVisibleWindowsByClass(std::wstring_view className) noexcept
{
    std::vector<HWND> windows;
    struct EnumArgs
    {
        std::wstring_view className;
        DWORD processId            = 0;
        std::vector<HWND>* windows = nullptr;
    } args{className, GetCurrentProcessId(), &windows};

    static_cast<void>(EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* args = reinterpret_cast<EnumArgs*>(lParam);
        if (! args || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        DWORD processId = 0;
        static_cast<void>(GetWindowThreadProcessId(hwnd, &processId));
        if (processId != args->processId)
        {
            return TRUE;
        }

        wchar_t classBuffer[128] = {};
        const int classLength    = GetClassNameW(hwnd, classBuffer, static_cast<int>(std::size(classBuffer)));
        if (classLength <= 0)
        {
            return TRUE;
        }

        if (args->className == std::wstring_view(classBuffer, static_cast<size_t>(classLength)))
        {
            args->windows->push_back(hwnd);
        }
        return TRUE;
    },
        reinterpret_cast<LPARAM>(&args)));
    return windows;
}

[[nodiscard]] HWND FindNewVisibleWindowByClass(std::wstring_view className, std::span<const HWND> existingWindows) noexcept
{
    const std::vector<HWND> currentWindows = CollectVisibleWindowsByClass(className);
    const auto isExisting = [&](HWND hwnd) noexcept { return std::find(existingWindows.begin(), existingWindows.end(), hwnd) != existingWindows.end(); };

    const auto it = std::find_if(currentWindows.begin(), currentWindows.end(), [&](HWND hwnd) noexcept { return ! isExisting(hwnd); });
    return (it != currentWindows.end()) ? *it : nullptr;
}

[[nodiscard]] bool RequestCloseWindow(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return true;
    }

    DWORD_PTR sendResult = 0;
    static_cast<void>(SendMessageTimeoutW(hwnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG | SMTO_NORMAL, 2000, &sendResult));
    if (PumpUntil([&]() noexcept { return IsWindow(hwnd) == FALSE; }, timeout))
    {
        return true;
    }

    static_cast<void>(PostMessageW(hwnd, WM_CLOSE, 0, 0));
    return PumpUntil([&]() noexcept { return IsWindow(hwnd) == FALSE; }, timeout);
}

void CloseVisibleWindowsByClass(std::wstring_view className) noexcept
{
    for (HWND hwnd : CollectVisibleWindowsByClass(className))
    {
        static_cast<void>(RequestCloseWindow(hwnd, 3000ms));
    }
}

[[nodiscard]] size_t CountVisibleChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (IsWindowVisible(child) != FALSE)
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

#ifdef _DEBUG
[[nodiscard]] bool TryGetViewerTextDebugSnapshot(HWND hwnd, WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
{
    snapshot = {};
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto remoteSnapshot = std::make_unique<WndMsg::ViewerTextDebugSnapshot>();
    if (! remoteSnapshot)
    {
        return false;
    }

    if (SendMessageW(hwnd, WndMsg::kViewerTextDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(remoteSnapshot.get())) == FALSE)
    {
        return false;
    }

    snapshot = *remoteSnapshot;
    return true;
}

template <typename Predicate>
[[nodiscard]] bool WaitForViewerTextSnapshot(HWND hwnd,
                                             Predicate&& predicate,
                                             std::chrono::milliseconds timeout,
                                             WndMsg::ViewerTextDebugSnapshot* outSnapshot = nullptr) noexcept
{
    WndMsg::ViewerTextDebugSnapshot snapshot{};
    WndMsg::ViewerTextDebugSnapshot lastSnapshot{};
    bool sawSnapshot = false;
    const bool ready = PumpUntil(
        [&]() noexcept
    {
        if (! TryGetViewerTextDebugSnapshot(hwnd, snapshot))
        {
            return false;
        }

        lastSnapshot = snapshot;
        sawSnapshot  = true;
        if (! predicate(snapshot))
        {
            return false;
        }

        if (outSnapshot)
        {
            *outSnapshot = snapshot;
        }
        return true;
    },
        timeout);

    if (! ready && outSnapshot && sawSnapshot)
    {
        *outSnapshot = lastSnapshot;
    }

    return ready;
}

ViewerTheme MakeViewerTextTestTheme(bool highContrast, bool rainbowMode = false) noexcept
{
    ViewerTheme theme{};
    theme.version                    = 4u;
    theme.dpi                        = 96u;
    theme.backgroundArgb             = 0xFFF7F4EEu;
    theme.textArgb                   = 0xFF1D1D1Du;
    theme.selectionBackgroundArgb    = 0xFFCCE2FFu;
    theme.selectionTextArgb          = 0xFF0E0E0Eu;
    theme.accentArgb                 = 0xFF0D7CC9u;
    theme.alertErrorBackgroundArgb   = 0xFF7A1C1Cu;
    theme.alertErrorTextArgb         = 0xFFFFFFFFu;
    theme.alertWarningBackgroundArgb = 0xFF8A5A00u;
    theme.alertWarningTextArgb       = 0xFFFFFFFFu;
    theme.alertInfoBackgroundArgb    = 0xFF124B78u;
    theme.alertInfoTextArgb          = 0xFFFFFFFFu;
    theme.darkMode                   = FALSE;
    theme.highContrast               = highContrast ? TRUE : FALSE;
    theme.rainbowMode                = rainbowMode ? TRUE : FALSE;
    theme.darkBase                   = FALSE;
    if (highContrast)
    {
        theme.diffAddedBackgroundArgb       = 0x55338A3Eu;
        theme.diffRemovedBackgroundArgb     = 0x55B92C39u;
        theme.diffContextBackgroundArgb     = 0x302A8FFFu;
        theme.diffHeaderBackgroundArgb      = 0x661F6FEBu;
        theme.diffBannerBackgroundArgb      = 0x881F6FEBu;
        theme.diffPlaceholderBackgroundArgb = 0x77124B78u;
        theme.diffDividerArgb               = 0xE0000000u;
        return theme;
    }

    if (rainbowMode)
    {
        theme.accentArgb                    = 0xFFB537FFu;
        theme.diffAddedBackgroundArgb       = 0x3352D67Bu;
        theme.diffRemovedBackgroundArgb     = 0x33FF627Au;
        theme.diffContextBackgroundArgb     = 0x18B537FFu;
        theme.diffHeaderBackgroundArgb      = 0x24B537FFu;
        theme.diffBannerBackgroundArgb      = 0x34B537FFu;
        theme.diffPlaceholderBackgroundArgb = 0x28B537FFu;
        theme.diffDividerArgb               = 0xB8C454FFu;
        return theme;
    }

    theme.diffAddedBackgroundArgb       = 0x332E7D32u;
    theme.diffRemovedBackgroundArgb     = 0x33D11A2Au;
    theme.diffContextBackgroundArgb     = 0x121F6FEBu;
    theme.diffHeaderBackgroundArgb      = 0x141F6FEBu;
    theme.diffBannerBackgroundArgb      = 0x221F6FEBu;
    theme.diffPlaceholderBackgroundArgb = 0x1A1F6FEBu;
    theme.diffDividerArgb               = 0xB8D0D0C8u;
    return theme;
}
#endif

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
    if (region)
    {
        const int rgnType = GetWindowRgn(hwnd, region.get());
        if (rgnType == NULLREGION)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool IsTextInputBridgeChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    wchar_t classBuffer[32] = {};
    const int classLength   = GetClassNameW(hwnd, classBuffer, static_cast<int>(std::size(classBuffer)));
    if (classLength <= 0)
    {
        return false;
    }

    const std::wstring_view className(classBuffer, static_cast<size_t>(classLength));
    return className == L"RICHEDIT50W" || className == L"Edit";
}

[[nodiscard]] size_t CountActuallyVisibleChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (IsActuallyVisibleChildWindow(child) && ! IsTextInputBridgeChildWindow(child))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[nodiscard]] size_t CountVisibleChildWindowsByClass(HWND hwnd, std::wstring_view className) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        std::wstring_view className;
        size_t count = 0u;
    } counter{className, 0u};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (IsWindowVisible(child) == FALSE)
        {
            return TRUE;
        }

        wchar_t classBuffer[64] = {};
        const int classLength   = GetClassNameW(child, classBuffer, static_cast<int>(std::size(classBuffer)));
        if (classLength > 0 && counterRef.className == std::wstring_view(classBuffer, static_cast<size_t>(classLength)))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[maybe_unused]] [[nodiscard]] HWND FindFirstChildWindowByClass(HWND hwnd, std::wstring_view className, bool requireVisible = true) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    struct MatchState
    {
        std::wstring_view className;
        bool requireVisible = true;
        HWND found          = nullptr;
    } state{className, requireVisible, nullptr};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& stateRef = *reinterpret_cast<MatchState*>(lParam);
        if (stateRef.requireVisible && IsWindowVisible(child) == FALSE)
        {
            return TRUE;
        }

        wchar_t classBuffer[64] = {};
        const int classLength   = GetClassNameW(child, classBuffer, static_cast<int>(std::size(classBuffer)));
        if (classLength > 0 && stateRef.className == std::wstring_view(classBuffer, static_cast<size_t>(classLength)))
        {
            stateRef.found = child;
            return FALSE;
        }

        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&state)));

    return state.found;
}

[[nodiscard]] HWND FindFirstVisibleChildWindowWithStyle(HWND hwnd, LONG_PTR requiredStyle) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    struct MatchState
    {
        LONG_PTR requiredStyle = 0;
        HWND found             = nullptr;
    } state{requiredStyle, nullptr};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& stateRef = *reinterpret_cast<MatchState*>(lParam);
        if (IsWindowVisible(child) == FALSE)
        {
            return TRUE;
        }

        const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
        if ((style & stateRef.requiredStyle) == stateRef.requiredStyle)
        {
            stateRef.found = child;
            return FALSE;
        }

        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&state)));

    return state.found;
}

void PrintVisibleChildWindowClasses(HWND hwnd, std::wstring_view prefix) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    std::wcout << std::format(L"[INFO] {} visible child classes:\n", prefix);
    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM) noexcept -> BOOL
    {
        if (IsWindowVisible(child) == FALSE)
        {
            return TRUE;
        }

        wchar_t classBuffer[64] = {};
        const int classLength   = GetClassNameW(child, classBuffer, static_cast<int>(std::size(classBuffer)));
        if (classLength > 0)
        {
            RECT rc{};
            static_cast<void>(GetWindowRect(child, &rc));
            std::wcout << std::format(
                L"  - {} ({}, {}, {}, {})\n", std::wstring_view(classBuffer, static_cast<size_t>(classLength)), rc.left, rc.top, rc.right, rc.bottom);
        }
        return TRUE;
    },
                                       0));
}

bool Check(bool condition, std::wstring_view message, bool& success);

void CheckDxNativeMenuBar(HWND viewerWindow, std::wstring_view viewerName, bool& success) noexcept
{
    Check(GetMenu(viewerWindow) == nullptr, std::format(L"{} detaches the native window menu after attaching the DxUi menu bar", viewerName), success);

    const size_t dxMenuBarCount = CountVisibleChildWindowsByClass(viewerWindow, kDxNativeMenuBarWindowClassName);
    if (dxMenuBarCount == 0u)
    {
        PrintVisibleChildWindowClasses(viewerWindow, viewerName);
    }

    Check(dxMenuBarCount >= 1u, std::format(L"{} exposes a visible DxUi menu bar host child", viewerName), success);
}

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &itemInfo) == 0)
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }

        if (itemInfo.hSubMenu)
        {
            ownerDrawCount += CountOwnerDrawMenuItems(itemInfo.hSubMenu);
        }
    }

    return ownerDrawCount;
}

void CheckPlainMenuModelContract(HWND viewerWindow, std::wstring_view viewerName, bool& success) noexcept
{
    WndMsg::ViewerNativeMenuModelDebugSnapshot snapshot{};
    const LRESULT queryResult = SendMessageW(viewerWindow, WndMsg::kViewerDebugGetNativeMenuModelSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot));
    Check(queryResult != FALSE, std::format(L"{} answers the hidden menu-model debug contract", viewerName), success);
    if (queryResult == FALSE)
    {
        return;
    }

    Check(snapshot.hasHiddenMenuModel, std::format(L"{} keeps a hidden HMENU model for DxUi menu conversion", viewerName), success);
    Check(snapshot.ownerDrawItemCount == 0u, std::format(L"{} hidden HMENU model keeps zero owner-draw items after DxUi menu migration", viewerName), success);
}

bool Check(bool condition, std::wstring_view message, bool& success)
{
    if (condition)
    {
        std::wcout << std::format(L"[PASS] {}\n", message);
        return true;
    }

    std::wcout << std::format(L"[FAIL] {}\n", message);
    success = false;
    return false;
}

[[nodiscard]] bool RunFilteredSelfExecutable(std::wstring_view testName, std::chrono::milliseconds timeout, bool& success) noexcept
{
    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), std::format(L"{} harness path resolves", testName), success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    std::wstring commandLine = std::format(L"\"{}\" \"{}\"", modulePath.data(), testName);
    STARTUPINFOW startupInfo{};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo{};
    const wil::unique_handle jobObject(CreateJobObjectW(nullptr, nullptr));
    Check(jobObject.is_valid(), std::format(L"{} fresh harness kill-on-close job object is created", testName), success);
    if (! jobObject.is_valid())
    {
        return false;
    }

    JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobInfo{};
    jobInfo.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
    const BOOL configuredJob                 = SetInformationJobObject(jobObject.get(), JobObjectExtendedLimitInformation, &jobInfo, sizeof(jobInfo));
    Check(configuredJob != FALSE, std::format(L"{} fresh harness kill-on-close job object is configured", testName), success);
    if (configuredJob == FALSE)
    {
        return false;
    }

    const BOOL created =
        CreateProcessW(modulePath.data(), commandLine.data(), nullptr, nullptr, FALSE, CREATE_SUSPENDED, nullptr, nullptr, &startupInfo, &processInfo);
    Check(created != FALSE, std::format(L"{} fresh harness process launches", testName), success);
    if (created == FALSE)
    {
        return false;
    }

    wil::unique_handle processHandle(processInfo.hProcess);
    wil::unique_handle threadHandle(processInfo.hThread);
    const BOOL assignedJob = AssignProcessToJobObject(jobObject.get(), processHandle.get());
    Check(assignedJob != FALSE, std::format(L"{} fresh harness process joins the kill-on-close job", testName), success);
    if (assignedJob == FALSE)
    {
        static_cast<void>(TerminateProcess(processHandle.get(), 1u));
        return false;
    }

    const DWORD resumedThreadCount = ResumeThread(threadHandle.get());
    Check(resumedThreadCount != DWORD(-1), std::format(L"{} fresh harness process resumes", testName), success);
    if (resumedThreadCount == DWORD(-1))
    {
        static_cast<void>(TerminateJobObject(jobObject.get(), 1u));
        return false;
    }

    const DWORD waitResult = WaitForSingleObject(processHandle.get(), static_cast<DWORD>(timeout.count()));
    if (waitResult == WAIT_TIMEOUT)
    {
        static_cast<void>(TerminateJobObject(jobObject.get(), 1u));
    }
    Check(waitResult == WAIT_OBJECT_0, std::format(L"{} fresh harness process exits before timeout", testName), success);
    if (waitResult != WAIT_OBJECT_0)
    {
        return false;
    }

    DWORD exitCode         = 1u;
    const BOOL gotExitCode = GetExitCodeProcess(processHandle.get(), &exitCode);
    Check(gotExitCode != FALSE, std::format(L"{} fresh harness process reports an exit code", testName), success);
    if (gotExitCode == FALSE)
    {
        return false;
    }

    Check(exitCode == 0u, std::format(L"{} fresh harness run passes", testName), success);
    return exitCode == 0u;
}

void CheckDxComboHostClickActivation(HWND viewerWindow, int comboControlId, std::wstring_view viewerName, bool& success)
{
    const HWND comboHost = GetDlgItem(viewerWindow, comboControlId);
    Check(comboHost != nullptr && IsWindow(comboHost) != FALSE, std::format(L"{} exposes a DxUi combo host child", viewerName), success);
    if (! comboHost || IsWindow(comboHost) == FALSE)
    {
        return;
    }

    Check(IsWindowVisible(comboHost) != FALSE, std::format(L"{} DxUi combo host is visible", viewerName), success);

    const LONG_PTR style = GetWindowLongPtrW(comboHost, GWL_STYLE);
    Check((style & SS_NOTIFY) != 0, std::format(L"{} DxUi combo host uses SS_NOTIFY so mouse clicks reach the shared host", viewerName), success);

    const LRESULT mouseActivate = SendMessageW(comboHost, WM_MOUSEACTIVATE, reinterpret_cast<WPARAM>(viewerWindow), MAKELPARAM(HTCLIENT, WM_LBUTTONDOWN));
    Check(mouseActivate == MA_ACTIVATE, std::format(L"{} DxUi combo host activates on mouse click", viewerName), success);

    RECT initialClient{};
    GetClientRect(comboHost, &initialClient);
    const int initialHeight = static_cast<int>(std::max<LONG>(0, initialClient.bottom - initialClient.top));
    const int clickX        = static_cast<int>(std::max<LONG>(4, initialClient.right - 12));
    const int clickY        = (std::max)(4, initialHeight / 2);

    static_cast<void>(SendMessageW(comboHost, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY)));
    static_cast<void>(SendMessageW(comboHost, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY)));

    RECT expandedClient{};
    const bool popupExpanded = PumpUntil(
        [&]() noexcept
    {
        GetClientRect(comboHost, &expandedClient);
        return (expandedClient.bottom - expandedClient.top) > initialHeight;
    },
        2000ms);
    Check(popupExpanded, std::format(L"{} DxUi combo host expands to show its popup rows on click", viewerName), success);

    if (popupExpanded)
    {
        static_cast<void>(SendMessageW(comboHost, WM_KEYDOWN, VK_ESCAPE, 0));
        static_cast<void>(SendMessageW(comboHost, WM_KEYUP, VK_ESCAPE, 0));

        RECT collapsedClient{};
        const bool popupCollapsed = PumpUntil(
            [&]() noexcept
        {
            GetClientRect(comboHost, &collapsedClient);
            return (collapsedClient.bottom - collapsedClient.top) == initialHeight;
        },
            2000ms);
        Check(popupCollapsed, std::format(L"{} DxUi combo host collapses back to header height after Escape", viewerName), success);
    }
}

[[nodiscard]] bool TryCollectVisibleUiaHostPatternStats(HWND hwnd, UiaHostPatternStats& stats) noexcept
{
    stats = {};
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    wil::com_ptr<IUIAutomation> automation;
    if (FAILED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation))) || ! automation)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> root;
    if (FAILED(automation->ElementFromHandle(hwnd, root.put())) || ! root)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationCondition> trueCondition;
    if (FAILED(automation->CreateTrueCondition(trueCondition.put())) || ! trueCondition)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElementArray> elements;
    if (FAILED(root->FindAll(TreeScope_Subtree, trueCondition.get(), elements.put())) || ! elements)
    {
        return false;
    }

    int length = 0;
    if (FAILED(elements->get_Length(&length)) || length <= 0)
    {
        return false;
    }

    for (int index = 0; index < length; ++index)
    {
        wil::com_ptr<IUIAutomationElement> element;
        if (FAILED(elements->GetElement(index, element.put())) || ! element)
        {
            continue;
        }

        VARIANT offscreen{};
        VariantInit(&offscreen);
        const bool hasOffscreen = SUCCEEDED(element->GetCurrentPropertyValue(UIA_IsOffscreenPropertyId, &offscreen));
        const bool isOffscreen  = hasOffscreen && offscreen.vt == VT_BOOL && offscreen.boolVal == VARIANT_TRUE;
        VariantClear(&offscreen);
        if (isOffscreen)
        {
            continue;
        }

        ++stats.visibleElementCount;

        int controlType = 0;
        if (SUCCEEDED(element->get_CurrentControlType(&controlType)))
        {
            if (controlType == UIA_ComboBoxControlTypeId)
            {
                ++stats.comboBoxControlCount;

                if (stats.comboName.empty())
                {
                    BSTR currentName           = nullptr;
                    const auto freeCurrentName = wil::scope_exit([&]() noexcept
                    {
                        if (currentName)
                        {
                            SysFreeString(currentName);
                        }
                    });

                    if (SUCCEEDED(element->get_CurrentName(&currentName)) && currentName && currentName[0] != L'\0')
                    {
                        stats.comboName = currentName;
                    }
                }
            }
            else if (controlType == UIA_EditControlTypeId)
            {
                ++stats.editControlCount;
            }
            else if (controlType == UIA_ButtonControlTypeId)
            {
                ++stats.buttonControlCount;
            }
        }

        wil::com_ptr<IUIAutomationValuePattern> valuePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
        {
            stats.hasValuePattern = true;

            if (stats.valueText.empty())
            {
                BSTR currentValue           = nullptr;
                const auto freeCurrentValue = wil::scope_exit([&]() noexcept
                {
                    if (currentValue)
                    {
                        SysFreeString(currentValue);
                    }
                });

                if (SUCCEEDED(valuePattern->get_CurrentValue(&currentValue)) && currentValue && currentValue[0] != L'\0')
                {
                    stats.valueText = currentValue;
                }
            }
        }

        wil::com_ptr<IUIAutomationExpandCollapsePattern> expandCollapsePattern;
        if (SUCCEEDED(
                element->GetCurrentPatternAs(UIA_ExpandCollapsePatternId, __uuidof(IUIAutomationExpandCollapsePattern), expandCollapsePattern.put_void())) &&
            expandCollapsePattern)
        {
            stats.hasExpandCollapsePattern = true;
        }
    }

    return true;
}

void CheckDxComboHostAccessibility(HWND viewerWindow, int comboControlId, std::wstring_view viewerName, std::wstring_view expectedSelectionText, bool& success)
{
    const HWND comboHost = GetDlgItem(viewerWindow, comboControlId);
    Check(comboHost != nullptr && IsWindow(comboHost) != FALSE, std::format(L"{} exposes a DxUi combo host child for UIA checks", viewerName), success);
    if (! comboHost || IsWindow(comboHost) == FALSE)
    {
        return;
    }

    UiaHostPatternStats stats{};
    bool providerReady            = false;
    bool subtreeReady             = false;
    const auto probeAccessibility = [&]() noexcept
    {
        providerReady = false;
        subtreeReady  = false;
        stats         = {};
        return PumpUntil(
            [&]() noexcept
        {
            providerReady = SendMessageW(comboHost, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId)) != 0;
            if (! providerReady)
            {
                subtreeReady = false;
                return false;
            }

            subtreeReady = TryCollectVisibleUiaHostPatternStats(comboHost, stats);
            return subtreeReady;
        },
            5000ms);
    };

    bool accessibilityReady = probeAccessibility();
    if (! accessibilityReady)
    {
        static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
        accessibilityReady = probeAccessibility();
    }

    Check(providerReady, std::format(L"{} DxUi combo host answers WM_GETOBJECT", viewerName), success);
    Check(accessibilityReady && subtreeReady, std::format(L"{} DxUi combo host exposes a traversable UIA subtree", viewerName), success);
    if (! accessibilityReady || ! subtreeReady)
    {
        return;
    }

    Check(stats.visibleElementCount > 0u, std::format(L"{} DxUi combo host exposes visible UIA elements", viewerName), success);
    Check(stats.comboBoxControlCount > 0u, std::format(L"{} DxUi combo host exposes a visible UIA ComboBox descendant", viewerName), success);
    Check(stats.hasExpandCollapsePattern || stats.hasValuePattern, std::format(L"{} DxUi combo host exposes live combo UIA patterns", viewerName), success);
    Check((! stats.comboName.empty() && stats.comboName.find(expectedSelectionText) != std::wstring::npos) ||
              (! stats.valueText.empty() && stats.valueText.find(expectedSelectionText) != std::wstring::npos),
          std::format(L"{} DxUi combo host exposes the current selection text through UIA", viewerName),
          success);
}

void CheckViewerTextPromptAccessibility(
    HWND promptWindow, std::wstring_view promptName, [[maybe_unused]] std::wstring_view expectedInitialText, std::chrono::milliseconds timeout, bool& success)
{
    Check(promptWindow != nullptr && IsWindow(promptWindow) != FALSE, std::format(L"{} prompt window exists", promptName), success);
    if (! promptWindow || IsWindow(promptWindow) == FALSE)
    {
        return;
    }

    const size_t visibleChildCount = CountActuallyVisibleChildWindows(promptWindow);
    if (visibleChildCount != 0u)
    {
        PrintVisibleChildWindowClasses(promptWindow, promptName);
    }
    Check(visibleChildCount == 0u, std::format(L"{} prompt does not expose visible child fallback controls", promptName), success);

    UiaHostPatternStats stats{};
    bool providerReady            = false;
    bool subtreeReady             = false;
    const auto probeAccessibility = [&]() noexcept
    {
        providerReady = false;
        subtreeReady  = false;
        stats         = {};
        return PumpUntil(
            [&]() noexcept
        {
            providerReady = SendMessageW(promptWindow, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId)) != 0;
            if (! providerReady)
            {
                subtreeReady = false;
                return false;
            }

            subtreeReady = TryCollectVisibleUiaHostPatternStats(promptWindow, stats);
            return subtreeReady;
        },
            timeout);
    };

    bool accessibilityReady = probeAccessibility();
    if (! accessibilityReady)
    {
        static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
        accessibilityReady = probeAccessibility();
    }

    Check(providerReady, std::format(L"{} prompt answers WM_GETOBJECT", promptName), success);
    Check(accessibilityReady && subtreeReady, std::format(L"{} prompt exposes a traversable UIA subtree", promptName), success);
    if (! accessibilityReady || ! subtreeReady)
    {
        return;
    }

    Check(stats.visibleElementCount > 0u, std::format(L"{} prompt exposes visible UIA elements", promptName), success);
    Check(stats.editControlCount >= 1u, std::format(L"{} prompt exposes a visible editable field", promptName), success);
    Check(stats.buttonControlCount >= 2u, std::format(L"{} prompt exposes visible DX command buttons", promptName), success);
    Check(stats.hasValuePattern, std::format(L"{} prompt exposes a live ValuePattern", promptName), success);
}

[[nodiscard]] std::filesystem::path WriteUtf8TextFile(const std::filesystem::path& path, std::string_view text)
{
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    stream.write(text.data(), static_cast<std::streamsize>(text.size()));
    stream.flush();
    return path;
}

[[nodiscard]] std::filesystem::path WriteBinaryFile(const std::filesystem::path& path, std::span<const std::byte> bytes)
{
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    stream.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    stream.flush();
    return path;
}

[[nodiscard]] std::filesystem::path WriteTinyJpegFile(const std::filesystem::path& path)
{
    static constexpr auto kTinyJpeg = std::to_array<unsigned char>({
        0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43, 0x00,
        0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0A, 0x0C, 0x14, 0x0D, 0x0C, 0x0B, 0x0B, 0x0C, 0x19, 0x12, 0x13, 0x0F,
        0x14, 0x1D, 0x1A, 0x1F, 0x1E, 0x1D, 0x1A, 0x1C, 0x1C, 0x20, 0x24, 0x2E, 0x27, 0x20, 0x22, 0x2C, 0x23, 0x1C, 0x1C, 0x28, 0x37, 0x29, 0x2C, 0x30, 0x31,
        0x34, 0x34, 0x34, 0x1F, 0x27, 0x39, 0x3D, 0x38, 0x32, 0x3C, 0x2E, 0x33, 0x34, 0x32, 0xFF, 0xDB, 0x00, 0x43, 0x01, 0x09, 0x09, 0x09, 0x0C, 0x0B, 0x0C,
        0x18, 0x0D, 0x0D, 0x18, 0x32, 0x21, 0x1C, 0x21, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32,
        0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0xFF, 0xC0,
        0x00, 0x11, 0x08, 0x00, 0x01, 0x00, 0x01, 0x03, 0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01, 0xFF, 0xC4, 0x00, 0x1F, 0x00, 0x00, 0x01, 0x05,
        0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0xFF,
        0xC4, 0x00, 0xB5, 0x10, 0x00, 0x02, 0x01, 0x03, 0x03, 0x02, 0x04, 0x03, 0x05, 0x05, 0x04, 0x04, 0x00, 0x00, 0x01, 0x7D, 0x01, 0x02, 0x03, 0x00, 0x04,
        0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08, 0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52,
        0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38,
        0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
        0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99,
        0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7,
        0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3,
        0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA, 0xFF, 0xC4, 0x00, 0x1F, 0x01, 0x00, 0x03, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0xFF, 0xC4, 0x00, 0xB5, 0x11, 0x00, 0x02, 0x01, 0x02, 0x04,
        0x04, 0x03, 0x04, 0x07, 0x05, 0x04, 0x04, 0x00, 0x01, 0x02, 0x77, 0x00, 0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07,
        0x61, 0x71, 0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xA1, 0xB1, 0xC1, 0x09, 0x23, 0x33, 0x52, 0xF0, 0x15, 0x62, 0x72, 0xD1, 0x0A, 0x16, 0x24,
        0x34, 0xE1, 0x25, 0xF1, 0x17, 0x18, 0x19, 0x1A, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
        0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79,
        0x7A, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7,
        0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5,
        0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA, 0xFF, 0xDA,
        0x00, 0x0C, 0x03, 0x01, 0x00, 0x02, 0x11, 0x03, 0x11, 0x00, 0x3F, 0x00, 0xF7, 0xFA, 0x28, 0xA2, 0x80, 0x3F, 0xFF, 0xD9,
    });

    return WriteBinaryFile(path, std::as_bytes(std::span(kTinyJpeg)));
}

[[nodiscard]] std::filesystem::path WriteTinyWaveFile(const std::filesystem::path& path)
{
    static constexpr auto kTinyWave = std::to_array<unsigned char>({
        'R',  'I',  'F',  'F',  0x2C, 0x00, 0x00, 0x00, 'W',  'A',  'V', 'E', 'f', 'm', 't',  ' ',  0x10, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x44, 0xAC,
        0x00, 0x00, 0x44, 0xAC, 0x00, 0x00, 0x01, 0x00, 0x08, 0x00, 'd', 'a', 't', 'a', 0x08, 0x00, 0x00, 0x00, 0x80, 0x90, 0xA0, 0xB0, 0xA0, 0x90, 0x80, 0x70,
    });

    return WriteBinaryFile(path, std::as_bytes(std::span(kTinyWave)));
}

class LocalFileReader final : public IFileReader
{
public:
    using RecordReadBytesFn = void (*)(void* cookie, std::wstring_view pathKey, size_t bytesRead) noexcept;

    explicit LocalFileReader(wil::unique_handle file, std::wstring pathKey, void* readCookie, RecordReadBytesFn recordReadBytes) noexcept
        : _file(std::move(file)),
          _pathKey(std::move(pathKey)),
          _readCookie(readCookie),
          _recordReadBytes(recordReadBytes)
    {
    }
    LocalFileReader(const LocalFileReader&)            = delete;
    LocalFileReader(LocalFileReader&&)                 = delete;
    LocalFileReader& operator=(const LocalFileReader&) = delete;
    LocalFileReader& operator=(LocalFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return static_cast<ULONG>(_refCount.fetch_add(1u, std::memory_order_relaxed) + 1u);
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = static_cast<ULONG>(_refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u);
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }

        LARGE_INTEGER size{};
        if (GetFileSizeEx(_file.get(), &size) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *sizeBytes = static_cast<uint64_t>(size.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        LARGE_INTEGER move{};
        move.QuadPart = offset;
        LARGE_INTEGER outPosition{};
        if (SetFilePointerEx(_file.get(), move, &outPosition, static_cast<DWORD>(origin)) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (newPosition)
        {
            *newPosition = static_cast<uint64_t>(outPosition.QuadPart);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! buffer && bytesToRead != 0u)
        {
            return E_POINTER;
        }

        DWORD localBytesRead = 0u;
        if (ReadFile(_file.get(), buffer, static_cast<DWORD>(bytesToRead), &localBytesRead, nullptr) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (bytesRead)
        {
            *bytesRead = static_cast<unsigned long>(localBytesRead);
        }
        if (localBytesRead != 0u && _recordReadBytes)
        {
            _recordReadBytes(_readCookie, _pathKey, static_cast<size_t>(localBytesRead));
        }
        return S_OK;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    std::wstring _pathKey;
    void* _readCookie                  = nullptr;
    RecordReadBytesFn _recordReadBytes = nullptr;
};

class BuiltinFileSystemStub final : public IFileSystem, public IInformations, public IFileSystemIO
{
public:
    BuiltinFileSystemStub()                                        = default;
    BuiltinFileSystemStub(const BuiltinFileSystemStub&)            = delete;
    BuiltinFileSystemStub(BuiltinFileSystemStub&&)                 = delete;
    BuiltinFileSystemStub& operator=(const BuiltinFileSystemStub&) = delete;
    BuiltinFileSystemStub& operator=(BuiltinFileSystemStub&&)      = delete;

    void ResetCreateFileReaderCounts() noexcept
    {
        std::scoped_lock lock(_createFileReaderCountsMutex, _readByteCountsMutex);
        _createFileReaderCounts.clear();
        _readByteCounts.clear();
    }

    [[nodiscard]] size_t GetCreateFileReaderCount(const std::filesystem::path& path) const noexcept
    {
        const std::wstring key = path.lexically_normal().wstring();
        std::scoped_lock lock(_createFileReaderCountsMutex);
        const auto it = _createFileReaderCounts.find(key);
        return it != _createFileReaderCounts.end() ? it->second : 0u;
    }

    [[nodiscard]] size_t GetReadByteCount(const std::filesystem::path& path) const noexcept
    {
        const std::wstring key = path.lexically_normal().wstring();
        std::scoped_lock lock(_readByteCountsMutex);
        const auto it = _readByteCounts.find(key);
        return it != _readByteCounts.end() ? it->second : 0u;
    }

    void RecordReadBytes(std::wstring_view pathKey, size_t bytesRead) noexcept
    {
        std::scoped_lock lock(_readByteCountsMutex);
        _readByteCounts[std::wstring(pathKey)] += bytesRead;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
        }
        else if (riid == __uuidof(IInformations))
        {
            *ppvObject = static_cast<IInformations*>(this);
        }
        else if (riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = static_cast<IFileSystemIO*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return static_cast<ULONG>(_refCount.fetch_add(1u, std::memory_order_relaxed) + 1u);
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        return static_cast<ULONG>(_refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u);
    }

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override
    {
        if (! metaData)
        {
            return E_POINTER;
        }

        *metaData = &kMetaData;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override
    {
        if (! schemaJsonUtf8)
        {
            return E_POINTER;
        }

        *schemaJsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* /*configurationJsonUtf8*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override
    {
        if (! configurationJsonUtf8)
        {
            return E_POINTER;
        }

        *configurationJsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints([[maybe_unused]] const wchar_t* path,
                                               [[maybe_unused]] FileSystemOperation operationType,
                                               [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        if (! path || path[0] == L'\0' || ! hints)
        {
            return E_INVALIDARG;
        }
        if (hints->sizeBytes < sizeof(FileSystemTransferHints))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                        FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        if (! path || path[0] == L'\0' || ! characteristics)
        {
            return E_INVALIDARG;
        }
        if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override
    {
        if (! pSomethingToSave)
        {
            return E_POINTER;
        }

        *pSomethingToSave = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* /*path*/, IFilesInformation** /*ppFilesInformation*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* /*sourcePath*/,
                                       const wchar_t* /*destinationPath*/,
                                       FileSystemFlags /*flags*/,
                                       const FileSystemOptions* /*options*/,
                                       IFileSystemCallback* /*callback*/,
                                       void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* /*sourcePath*/,
                                       const wchar_t* /*destinationPath*/,
                                       FileSystemFlags /*flags*/,
                                       const FileSystemOptions* /*options*/,
                                       IFileSystemCallback* /*callback*/,
                                       void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t* /*path*/,
                                         FileSystemFlags /*flags*/,
                                         const FileSystemOptions* /*options*/,
                                         IFileSystemCallback* /*callback*/,
                                         void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* /*sourcePath*/,
                                         const wchar_t* /*destinationPath*/,
                                         FileSystemFlags /*flags*/,
                                         const FileSystemOptions* /*options*/,
                                         IFileSystemCallback* /*callback*/,
                                         void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* /*sourcePaths*/,
                                        unsigned long /*count*/,
                                        const wchar_t* /*destinationFolder*/,
                                        FileSystemFlags /*flags*/,
                                        const FileSystemOptions* /*options*/,
                                        IFileSystemCallback* /*callback*/,
                                        void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* /*sourcePaths*/,
                                        unsigned long /*count*/,
                                        const wchar_t* /*destinationFolder*/,
                                        FileSystemFlags /*flags*/,
                                        const FileSystemOptions* /*options*/,
                                        IFileSystemCallback* /*callback*/,
                                        void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* /*paths*/,
                                          unsigned long /*count*/,
                                          FileSystemFlags /*flags*/,
                                          const FileSystemOptions* /*options*/,
                                          IFileSystemCallback* /*callback*/,
                                          void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* /*items*/,
                                          unsigned long /*count*/,
                                          FileSystemFlags /*flags*/,
                                          const FileSystemOptions* /*options*/,
                                          IFileSystemCallback* /*callback*/,
                                          void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }

        *jsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! path || ! fileAttributes)
        {
            return E_POINTER;
        }

        const DWORD attributes = GetFileAttributesW(path);
        if (attributes == INVALID_FILE_ATTRIBUTES)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *fileAttributes = static_cast<unsigned long>(attributes);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (! path || ! reader)
        {
            return E_POINTER;
        }

        *reader                = nullptr;
        const std::wstring key = std::filesystem::path(path).lexically_normal().wstring();
        {
            std::scoped_lock lock(_createFileReaderCountsMutex);
            _createFileReaderCounts[key] += 1u;
        }
        wil::unique_handle file(
            CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        auto localReader = std::make_unique<LocalFileReader>(std::move(file),
                                                             key,
                                                             this,
                                                             [](void* cookie, std::wstring_view pathKey, size_t bytesRead) noexcept
        {
            auto* self = static_cast<BuiltinFileSystemStub*>(cookie);
            if (self)
            {
                self->RecordReadBytes(pathKey, bytesRead);
            }
        });
        *reader          = localReader.release();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* /*path*/, FileSystemFlags /*flags*/, IFileWriter** writer) noexcept override
    {
        if (! writer)
        {
            return E_POINTER;
        }

        *writer = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! path || ! info)
        {
            return E_POINTER;
        }

        WIN32_FILE_ATTRIBUTE_DATA attributes{};
        if (GetFileAttributesExW(path, GetFileExInfoStandard, &attributes) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        info->sizeBytes = sizeof(FileSystemBasicInformation);
        ULARGE_INTEGER creation{};
        creation.LowPart   = attributes.ftCreationTime.dwLowDateTime;
        creation.HighPart  = attributes.ftCreationTime.dwHighDateTime;
        info->creationTime = static_cast<__int64>(creation.QuadPart);
        ULARGE_INTEGER access{};
        access.LowPart       = attributes.ftLastAccessTime.dwLowDateTime;
        access.HighPart      = attributes.ftLastAccessTime.dwHighDateTime;
        info->lastAccessTime = static_cast<__int64>(access.QuadPart);
        ULARGE_INTEGER write{};
        write.LowPart       = attributes.ftLastWriteTime.dwLowDateTime;
        write.HighPart      = attributes.ftLastWriteTime.dwHighDateTime;
        info->lastWriteTime = static_cast<__int64>(write.QuadPart);
        info->attributes    = attributes.dwFileAttributes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* /*path*/, const FileSystemBasicInformation* /*info*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* /*path*/, const char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }

        *jsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

private:
    inline static const PluginMetaData kMetaData{
        L"builtin/file-system",
        L"file",
        L"File System",
        L"Test built-in file system stub",
        L"ViewerPETests",
        L"1",
    };

    std::atomic_ulong _refCount{1};
    mutable std::mutex _createFileReaderCountsMutex;
    std::unordered_map<std::wstring, size_t> _createFileReaderCounts;
    mutable std::mutex _readByteCountsMutex;
    std::unordered_map<std::wstring, size_t> _readByteCounts;
};

[[nodiscard]] bool TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerPE.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    const std::filesystem::path appPath    = buildDir / L"RedSalamander.exe";
    Check(std::filesystem::exists(pluginPath), L"ViewerPE.dll is present for runtime validation", success);
    Check(std::filesystem::exists(appPath), L"RedSalamander.exe is present as an alternate PE input", success);
    if (! std::filesystem::exists(pluginPath) || ! std::filesystem::exists(appPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerPE.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerPE factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerPEPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerPE factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerPEWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const std::wstring pluginPathText = pluginPath.wstring();
    const std::wstring appPathText    = appPath.wstring();
    const wchar_t* otherFiles[]       = {pluginPathText.c_str(), appPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = pluginPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerPE window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerPEWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"ViewerPE window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(CountVisibleChildWindows(viewerWindow) >= 1u, L"ViewerPE exposes a visible child DxUi host when alternate files exist", success);
    Check(CountVisibleChildWindowsByClass(viewerWindow, L"ComboBox") == 0u, L"ViewerPE does not expose a visible legacy ComboBox child", success);
    CheckDxNativeMenuBar(viewerWindow, L"ViewerPE", success);
    CheckDxComboHostClickActivation(viewerWindow, kViewerPEFileComboId, L"ViewerPE", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerPEFileComboId, L"ViewerPE", pluginPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerPE window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerPE window closes cleanly", success);
    return success;
}

[[nodiscard]] bool TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerWeb test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerWeb.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerWeb.dll is present for runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerWeb.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerWeb factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerWebPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerWeb factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerWebTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path firstPath  = WriteUtf8TextFile(tempDir / L"viewerweb-first.html", "<!doctype html><html><body>first</body></html>");
    const std::filesystem::path secondPath = WriteUtf8TextFile(tempDir / L"viewerweb-second.html", "<!doctype html><html><body>second</body></html>");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(firstPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(secondPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring firstPathText  = firstPath.wstring();
    const std::wstring secondPathText = secondPath.wstring();
    const wchar_t* otherFiles[]       = {firstPathText.c_str(), secondPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = firstPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerWeb window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerWebWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        8000ms);
    Check(openedWindow, L"ViewerWeb window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(CountVisibleChildWindows(viewerWindow) >= 1u, L"ViewerWeb exposes visible child host content when alternate files exist", success);
    Check(CountVisibleChildWindowsByClass(viewerWindow, L"ComboBox") == 0u, L"ViewerWeb does not expose a visible legacy ComboBox child", success);
    CheckDxNativeMenuBar(viewerWindow, L"ViewerWeb", success);
    CheckPlainMenuModelContract(viewerWindow, L"ViewerWeb", success);
    CheckDxComboHostClickActivation(viewerWindow, kViewerWebFileComboId, L"ViewerWeb", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerWebFileComboId, L"ViewerWeb", firstPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerWeb window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerWeb window closes cleanly", success);
    return success;
}

[[nodiscard]] bool TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerImgRaw test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerImgRaw factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerImgRawPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerImgRaw factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerImgRawTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path firstPath  = WriteTinyJpegFile(tempDir / L"viewerimgraw-first.jpg");
    const std::filesystem::path secondPath = WriteTinyJpegFile(tempDir / L"viewerimgraw-second.jpg");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(firstPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(secondPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring firstPathText  = firstPath.wstring();
    const std::wstring secondPathText = secondPath.wstring();
    const wchar_t* otherFiles[]       = {firstPathText.c_str(), secondPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = firstPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerImgRaw window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        8000ms);
    Check(openedWindow, L"ViewerImgRaw window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(CountVisibleChildWindows(viewerWindow) >= 1u, L"ViewerImgRaw exposes visible child host content when alternate files exist", success);
    Check(CountVisibleChildWindowsByClass(viewerWindow, L"ComboBox") == 0u, L"ViewerImgRaw does not expose a visible legacy ComboBox child", success);
    CheckDxNativeMenuBar(viewerWindow, L"ViewerImgRaw", success);
    CheckPlainMenuModelContract(viewerWindow, L"ViewerImgRaw", success);
    CheckDxComboHostClickActivation(viewerWindow, kViewerImgRawFileComboId, L"ViewerImgRaw", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerImgRawFileComboId, L"ViewerImgRaw", firstPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerImgRaw window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerImgRaw window closes cleanly", success);
    return success;
}

[[nodiscard]] bool TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerText test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerText.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerTextTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path firstPath  = WriteUtf8TextFile(tempDir / L"viewertext-first.txt", "first viewer text file\n");
    const std::filesystem::path secondPath = WriteUtf8TextFile(tempDir / L"viewertext-second.txt", "second viewer text file\n");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(firstPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(secondPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring firstPathText  = firstPath.wstring();
    const std::wstring secondPathText = secondPath.wstring();
    const wchar_t* otherFiles[]       = {firstPathText.c_str(), secondPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = firstPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerText window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerTextWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"ViewerText window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(CountVisibleChildWindows(viewerWindow) >= 1u, L"ViewerText exposes visible child host content when alternate files exist", success);
    const size_t visibleLegacyComboCount = CountVisibleChildWindowsByClass(viewerWindow, L"ComboBox");
    if (visibleLegacyComboCount != 0u)
    {
        PrintVisibleChildWindowClasses(viewerWindow, L"ViewerText");
    }
    Check(visibleLegacyComboCount == 0u, L"ViewerText does not expose a visible legacy ComboBox child", success);
    CheckDxNativeMenuBar(viewerWindow, L"ViewerText", success);
    CheckPlainMenuModelContract(viewerWindow, L"ViewerText", success);
    CheckDxComboHostClickActivation(viewerWindow, kViewerTextFileComboId, L"ViewerText", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerTextFileComboId, L"ViewerText", firstPath.filename().wstring(), success);
#ifdef _DEBUG
    WndMsg::ViewerTextDebugSnapshot snapshot{};
    const bool snapshotReady = WaitForViewerTextSnapshot(viewerWindow, [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
        return value.renderCount > 0u && value.legacyVisibleGdiTextSurfaceCount == 0u && value.legacyVisibleHfontSurfaceCount == 0u;
    }, 5000ms, &snapshot);
    Check(snapshotReady, L"ViewerText exposes no visible legacy GDI or HFONT text surfaces in the active shell", success);
#endif
    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerText window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerText window closes cleanly", success);
    return success;
}

#ifdef _DEBUG
[[nodiscard]] bool TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerText hex color test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for hex color runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerText.dll loads successfully for hex color runtime validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available for hex color runtime validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates an IViewer instance for hex color runtime validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
    Check(SUCCEEDED(infoHr) && info != nullptr, L"ViewerText exposes IInformations for hex color runtime validation", success);
    if (FAILED(infoHr) || ! info)
    {
        return false;
    }

    const HRESULT configHr = info->SetConfiguration(R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1"})json");
    Check(SUCCEEDED(configHr), L"ViewerText accepts configuration without an explicit hex byte color key", success);
    if (FAILED(configHr))
    {
        return false;
    }

    const char* configurationJson    = nullptr;
    const HRESULT getDefaultConfigHr = info->GetConfiguration(&configurationJson);
    Check(SUCCEEDED(getDefaultConfigHr) && configurationJson != nullptr, L"ViewerText returns canonical configuration before opening in hex mode", success);
    if (configurationJson != nullptr)
    {
        const std::string_view configurationView(configurationJson);
        Check(configurationView.find("\"hexByteColorMode\":\"leadingNibble\"") != std::string_view::npos,
              L"ViewerText fills in leadingNibble when the hex byte color key is absent",
              success);
    }

    const ViewerTheme normalTheme = MakeViewerTextTestTheme(false);
    const HRESULT themeHr         = viewer->SetTheme(&normalTheme);
    Check(SUCCEEDED(themeHr), L"ViewerText accepts the normal theme before opening in hex mode", success);
    if (FAILED(themeHr))
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerTextTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    static constexpr auto kHexColorFixture = std::to_array<std::byte>({
        std::byte{0x00}, std::byte{0x0F}, std::byte{0x10}, std::byte{0x1A}, std::byte{0x2B}, std::byte{0x3C}, std::byte{0x4D}, std::byte{0x5E},
        std::byte{0x6F}, std::byte{0x70}, std::byte{0x7F}, std::byte{0x80}, std::byte{0x8A}, std::byte{0x90}, std::byte{0xA5}, std::byte{0xBF},
        std::byte{0xC1}, std::byte{0xD2}, std::byte{0xE3}, std::byte{0xF4}, std::byte{0xFF}, std::byte{0x00}, std::byte{0x11}, std::byte{0x22},
        std::byte{0x33}, std::byte{0x44}, std::byte{0x55}, std::byte{0x66}, std::byte{0x77}, std::byte{0x88}, std::byte{0x99}, std::byte{0xAA},
    });
    const std::filesystem::path samplePath = WriteBinaryFile(tempDir / L"viewertext-hex-colors.bin", std::span(kHexColorFixture));
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.focusedPath    = samplePathText.c_str();
    context.flags          = VIEWER_OPEN_FLAG_START_HEX;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerText opens the crafted binary fixture in hex mode", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow     = nullptr;
    auto emergencyCleanup = wil::scope_exit([&]() noexcept
    {
        if (viewerWindow && IsWindow(viewerWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(viewerWindow, 5000ms));
        }
    });

    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerTextWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"ViewerText hex color runtime validation window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot initialSnapshot{};
    const bool initialSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Hex &&
               snapshot.hexByteColorMode == WndMsg::ViewerTextDebugHexByteColorMode::LeadingNibble && snapshot.renderCount > 0u &&
               snapshot.visibleByteCount > 0u && snapshot.visibleColorizedByteCount > 0u && snapshot.visibleUniqueColorBucketCount > 1u &&
               ! snapshot.highContrastFallback;
    },
                                                                5000ms,
                                                                &initialSnapshot);
    Check(initialSnapshotReady, L"ViewerText reports visible byte colors in hex mode through the debug snapshot", success);
    if (initialSnapshotReady)
    {
        Check(initialSnapshot.visibleColorizedByteCount <= initialSnapshot.visibleByteCount, L"ViewerText reports a bounded colorized byte count", success);
    }

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_HEX_BYTE_COLORS_OFF), 0));

    WndMsg::ViewerTextDebugSnapshot offSnapshot{};
    const bool offSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Hex && snapshot.hexByteColorMode == WndMsg::ViewerTextDebugHexByteColorMode::Off &&
               snapshot.renderCount > initialSnapshot.renderCount && snapshot.visibleByteCount > 0u && snapshot.visibleColorizedByteCount == 0u &&
               ! snapshot.highContrastFallback;
    },
                                                            5000ms,
                                                            &offSnapshot);
    Check(offSnapshotReady, L"ViewerText menu can switch hex byte colors off at runtime", success);

    configurationJson            = nullptr;
    const HRESULT getOffConfigHr = info->GetConfiguration(&configurationJson);
    Check(
        SUCCEEDED(getOffConfigHr) && configurationJson != nullptr, L"ViewerText returns canonical configuration after switching hex byte colors off", success);
    if (configurationJson != nullptr)
    {
        const std::string_view configurationView(configurationJson);
        Check(configurationView.find("\"hexByteColorMode\":\"off\"") != std::string_view::npos,
              L"ViewerText persists the off hex byte color mode after the menu command",
              success);
    }

    const uint64_t renderBeforeRestore = offSnapshotReady ? offSnapshot.renderCount : initialSnapshot.renderCount;
    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LEADING_NIBBLE), 0));

    WndMsg::ViewerTextDebugSnapshot restoredSnapshot{};
    const bool restoredSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                 [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Hex &&
               snapshot.hexByteColorMode == WndMsg::ViewerTextDebugHexByteColorMode::LeadingNibble && snapshot.renderCount > renderBeforeRestore &&
               snapshot.visibleByteCount > 0u && snapshot.visibleColorizedByteCount > 0u && snapshot.visibleUniqueColorBucketCount > 1u &&
               ! snapshot.highContrastFallback;
    },
                                                                 5000ms,
                                                                 &restoredSnapshot);
    Check(restoredSnapshotReady, L"ViewerText menu can restore leading-nibble byte colors at runtime", success);

    configurationJson                 = nullptr;
    const HRESULT getRestoredConfigHr = info->GetConfiguration(&configurationJson);
    Check(SUCCEEDED(getRestoredConfigHr) && configurationJson != nullptr,
          L"ViewerText returns canonical configuration after restoring leading-nibble byte colors",
          success);
    if (configurationJson != nullptr)
    {
        const std::string_view configurationView(configurationJson);
        Check(configurationView.find("\"hexByteColorMode\":\"leadingNibble\"") != std::string_view::npos,
              L"ViewerText persists the leading-nibble hex byte color mode after the menu command",
              success);
    }

    const ViewerTheme highContrastTheme = MakeViewerTextTestTheme(true);
    const HRESULT highContrastHr        = viewer->SetTheme(&highContrastTheme);
    Check(SUCCEEDED(highContrastHr), L"ViewerText accepts a high-contrast theme update in hex mode", success);
    if (SUCCEEDED(highContrastHr))
    {
        WndMsg::ViewerTextDebugSnapshot highContrastSnapshot{};
        const bool highContrastReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                 [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Hex &&
                   snapshot.hexByteColorMode == WndMsg::ViewerTextDebugHexByteColorMode::LeadingNibble && snapshot.renderCount > 0u &&
                   snapshot.highContrastFallback && snapshot.visibleColorizedByteCount == 0u;
        },
                                                                 5000ms,
                                                                 &highContrastSnapshot);
        Check(highContrastReady, L"ViewerText falls back to monochrome in high-contrast mode", success);
        if (highContrastReady)
        {
            Check(highContrastSnapshot.visibleByteCount >= initialSnapshot.visibleByteCount,
                  L"ViewerText keeps the visible byte accounting stable after the high-contrast theme update",
                  success);
        }
    }

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerText hex color runtime validation window closes cleanly", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerText hex color runtime validation window is destroyed after close",
          success);
    return success;
}

[[nodiscard]] bool TestViewerTextDiffModesAndPlaceholders() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerText diff test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for diff runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerText.dll loads successfully for diff runtime validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available for diff runtime validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates an IViewer instance for diff runtime validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
    Check(SUCCEEDED(infoHr) && info != nullptr, L"ViewerText exposes IInformations for diff runtime validation", success);
    if (FAILED(infoHr) || ! info)
    {
        return false;
    }

    const ViewerTheme normalTheme       = MakeViewerTextTestTheme(false);
    const ViewerTheme rainbowTheme      = MakeViewerTextTestTheme(false, true);
    const ViewerTheme highContrastTheme = MakeViewerTextTestTheme(true);
    const HRESULT themeHr               = viewer->SetTheme(&normalTheme);
    Check(SUCCEEDED(themeHr), L"ViewerText accepts the normal theme before opening in diff mode", success);
    if (FAILED(themeHr))
    {
        return false;
    }

    const std::filesystem::path tempDir = buildDir / L"ViewerTextTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));

    const auto makeStableFileBody = [](std::string_view stablePrefix,
                                       size_t lineCount,
                                       std::initializer_list<std::pair<size_t, std::string_view>> overrides,
                                       std::string_view stableSuffix = {}) -> std::string
    {
        std::string text;
        text.reserve(lineCount * (40u + stableSuffix.size()));

        for (size_t lineIndex = 1u; lineIndex <= lineCount; ++lineIndex)
        {
            std::string_view lineText;
            for (const auto& [overrideLine, overrideValue] : overrides)
            {
                if (overrideLine == lineIndex)
                {
                    lineText = overrideValue;
                    break;
                }
            }

            if (lineText.empty())
            {
                text += std::format("{} {:03}{}\n", stablePrefix, lineIndex, stableSuffix);
            }
            else
            {
                text += lineText;
                text.push_back('\n');
            }
        }

        return text;
    };
    const std::filesystem::path oldPath = WriteUtf8TextFile(tempDir / L"viewertext-diff-old.txt", "alpha\nbeta\ngamma\ndelta\nepsilon\nzeta\neta\ntheta\n");
    const std::filesystem::path newPath =
        WriteUtf8TextFile(tempDir / L"viewertext-diff-new.txt", "alpha\nbeta updated\ngamma\ndelta\nepsilon\nzeta\neta updated\ntheta\n");
    const std::filesystem::path diffPath          = WriteUtf8TextFile(tempDir / L"viewertext-layout.diff",
                                                                      "diff --git a/viewertext-diff-old.txt b/viewertext-diff-new.txt\n"
                                                                      "--- a/viewertext-diff-old.txt\n"
                                                                      "+++ b/viewertext-diff-new.txt\n"
                                                                      "@@ -2 +2 @@\n"
                                                                      "-beta\n"
                                                                      "+beta updated\n"
                                                                      "@@ -7 +7 @@\n"
                                                                      "-eta\n"
                                                                      "+eta updated\n");
    const std::filesystem::path missingDiffPath   = WriteUtf8TextFile(tempDir / L"viewertext-missing.diff",
                                                                      "diff --git a/viewertext-missing-left.txt b/viewertext-missing-right.txt\n"
                                                                      "--- a/viewertext-missing-left.txt\n"
                                                                      "+++ b/viewertext-missing-right.txt\n"
                                                                      "@@ -2 +2 @@\n"
                                                                      "-beta\n"
                                                                      "+beta updated\n"
                                                                      "@@ -7 +7 @@\n"
                                                                      "-eta\n"
                                                                      "+eta updated\n");
    const std::filesystem::path multiFileDiffPath = WriteUtf8TextFile(tempDir / L"viewertext-parser-cases.diff",
                                                                      "diff --git a/viewertext-created.txt b/viewertext-created.txt\n"
                                                                      "new file mode 100644\n"
                                                                      "--- /dev/null\n"
                                                                      "+++ b/viewertext-created.txt\n"
                                                                      "@@ -0,0 +1,1 @@\n"
                                                                      "+created line\n"
                                                                      "diff --git a/viewertext-deleted.txt b/viewertext-deleted.txt\n"
                                                                      "deleted file mode 100644\n"
                                                                      "--- a/viewertext-deleted.txt\n"
                                                                      "+++ /dev/null\n"
                                                                      "@@ -1,1 +0,0 @@\n"
                                                                      "-deleted line\n"
                                                                      "diff --git a/viewertext-renamed-old.txt b/viewertext-renamed-new.txt\n"
                                                                      "similarity index 60%\n"
                                                                      "rename from viewertext-renamed-old.txt\n"
                                                                      "rename to viewertext-renamed-new.txt\n"
                                                                      "--- a/viewertext-renamed-old.txt\n"
                                                                      "+++ b/viewertext-renamed-new.txt\n"
                                                                      "@@ -1,1 +1,1 @@\n"
                                                                      "-old name\n"
                                                                      "+new name\n"
                                                                      "diff --git a/viewertext-second-old.txt b/viewertext-second-new.txt\n"
                                                                      "--- a/viewertext-second-old.txt\n"
                                                                      "+++ b/viewertext-second-new.txt\n"
                                                                      "@@ -1,1 +1,1 @@\n"
                                                                      "-second old\n"
                                                                      "+second new\n");
    const std::filesystem::path malformedDiffPath = WriteUtf8TextFile(tempDir / L"viewertext-malformed.diff",
                                                                      "diff --git a/viewertext-bad.txt b/viewertext-bad.txt\n"
                                                                      "--- a/viewertext-bad.txt\n"
                                                                      "+++ b/viewertext-bad.txt\n"
                                                                      "@@ -bad +1 @@\n"
                                                                      "-broken\n"
                                                                      "+fixed\n");
    const std::filesystem::path sectionFirstOldPath =
        WriteUtf8TextFile(tempDir / L"viewertext-section-first-old.txt", "first keep 1\nfirst keep 2\nfirst keep 3\nfirst old value\nfirst keep 5\n");
    const std::filesystem::path sectionFirstNewPath =
        WriteUtf8TextFile(tempDir / L"viewertext-section-first-new.txt", "first keep 1\nfirst keep 2\nfirst keep 3\nfirst new value\nfirst keep 5\n");
    const std::filesystem::path sectionSecondOldPath =
        WriteUtf8TextFile(tempDir / L"viewertext-section-second-old.txt", "second keep 1\nsecond keep 2\nsecond old value\nsecond keep 4\nsecond keep 5\n");
    const std::filesystem::path sectionSecondNewPath =
        WriteUtf8TextFile(tempDir / L"viewertext-section-second-new.txt", "second keep 1\nsecond keep 2\nsecond new value\nsecond keep 4\nsecond keep 5\n");
    const std::filesystem::path sectionScopedDiffPath = WriteUtf8TextFile(tempDir / L"viewertext-section-scoped.diff",
                                                                          "diff --git a/viewertext-section-first-old.txt b/viewertext-section-first-new.txt\n"
                                                                          "--- a/viewertext-section-first-old.txt\n"
                                                                          "+++ b/viewertext-section-first-new.txt\n"
                                                                          "@@ -4 +4 @@\n"
                                                                          "-first old value\n"
                                                                          "+first new value\n"
                                                                          "diff --git a/viewertext-section-second-old.txt b/viewertext-section-second-new.txt\n"
                                                                          "--- a/viewertext-section-second-old.txt\n"
                                                                          "+++ b/viewertext-section-second-new.txt\n"
                                                                          "@@ -3 +3 @@\n"
                                                                          "-second old value\n"
                                                                          "+second new value\n");
    const std::string viewportLineSuffix(192u, 'x');
    const std::filesystem::path viewportOldPath = WriteUtf8TextFile(
        tempDir / L"viewertext-viewport-old.txt",
        makeStableFileBody(
            "viewport stable line", 1024u, {{20u, "viewport changed line 020 old"}, {920u, "viewport changed line 920 old"}}, viewportLineSuffix));
    const std::filesystem::path viewportNewPath = WriteUtf8TextFile(
        tempDir / L"viewertext-viewport-new.txt",
        makeStableFileBody(
            "viewport stable line", 1024u, {{20u, "viewport changed line 020 new"}, {920u, "viewport changed line 920 new"}}, viewportLineSuffix));
    const std::filesystem::path viewportLazyDiffPath      = WriteUtf8TextFile(tempDir / L"viewertext-viewport-lazy.diff",
                                                                              "diff --git a/viewertext-viewport-old.txt b/viewertext-viewport-new.txt\n"
                                                                              "--- a/viewertext-viewport-old.txt\n"
                                                                              "+++ b/viewertext-viewport-new.txt\n"
                                                                              "@@ -18,5 +18,5 @@\n"
                                                                              " viewport stable line 018\n"
                                                                              " viewport stable line 019\n"
                                                                              "-viewport changed line 020 old\n"
                                                                              "+viewport changed line 020 new\n"
                                                                              " viewport stable line 021\n"
                                                                              " viewport stable line 022\n"
                                                                              "@@ -918,5 +918,5 @@\n"
                                                                              " viewport stable line 918\n"
                                                                              " viewport stable line 919\n"
                                                                              "-viewport changed line 920 old\n"
                                                                              "+viewport changed line 920 new\n"
                                                                              " viewport stable line 921\n"
                                                                              " viewport stable line 922\n");
    constexpr uint32_t kLargeBufferedDiffChangedLineCount = 30000u;
    std::string largeBufferedDiffText;
    largeBufferedDiffText.reserve(5u * 1024u * 1024u);
    largeBufferedDiffText += "diff --git a/viewertext-large-old.txt b/viewertext-large-new.txt\n";
    largeBufferedDiffText += "--- a/viewertext-large-old.txt\n";
    largeBufferedDiffText += "+++ b/viewertext-large-new.txt\n";
    largeBufferedDiffText += std::format("@@ -1,{} +1,{} @@\n", kLargeBufferedDiffChangedLineCount, kLargeBufferedDiffChangedLineCount);
    for (uint32_t lineIndex = 0u; lineIndex < kLargeBufferedDiffChangedLineCount; ++lineIndex)
    {
        largeBufferedDiffText += std::format("-large old line {:05} payload payload payload payload payload payload payload\n", lineIndex);
        largeBufferedDiffText += std::format("+large new line {:05} payload payload payload payload payload payload payload\n", lineIndex);
    }
    const std::filesystem::path largeBufferedDiffPath     = WriteUtf8TextFile(tempDir / L"viewertext-large-buffered.diff", largeBufferedDiffText);
    const std::filesystem::path largeSniffedDiffPath      = WriteUtf8TextFile(tempDir / L"viewertext-large-sniffed.txt", largeBufferedDiffText);
    constexpr uint32_t kLargeStreamedDiffChangedLineCount = 30000u;
    std::string largeStreamedDiffText;
    largeStreamedDiffText.reserve(24u * 1024u * 1024u);
    for (const std::string_view sectionName : {std::string_view{"first"}, std::string_view{"second"}, std::string_view{"third"}, std::string_view{"fourth"}})
    {
        largeStreamedDiffText += std::format("diff --git a/viewertext-stream-{}-old.txt b/viewertext-stream-{}-new.txt\n", sectionName, sectionName);
        largeStreamedDiffText += std::format("--- a/viewertext-stream-{}-old.txt\n", sectionName);
        largeStreamedDiffText += std::format("+++ b/viewertext-stream-{}-new.txt\n", sectionName);
        largeStreamedDiffText += std::format("@@ -1,{} +1,{} @@\n", kLargeStreamedDiffChangedLineCount, kLargeStreamedDiffChangedLineCount);
        for (uint32_t lineIndex = 0u; lineIndex < kLargeStreamedDiffChangedLineCount; ++lineIndex)
        {
            largeStreamedDiffText +=
                std::format("-{} streamed old line {:05} payload payload payload payload payload payload payload\n", sectionName, lineIndex);
            largeStreamedDiffText +=
                std::format("+{} streamed new line {:05} payload payload payload payload payload payload payload\n", sectionName, lineIndex);
        }
    }
    const std::filesystem::path largeStreamedDiffPath = WriteUtf8TextFile(tempDir / L"viewertext-large-streamed-multifile.diff", largeStreamedDiffText);

    auto cleanupTemp = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(oldPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(newPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(diffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(missingDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(multiFileDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(malformedDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(sectionFirstOldPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(sectionFirstNewPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(sectionSecondOldPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(sectionSecondNewPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(sectionScopedDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(viewportOldPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(viewportNewPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(viewportLazyDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(largeBufferedDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(largeSniffedDiffPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(largeStreamedDiffPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;

    auto openViewerForPath = [&](const std::filesystem::path& path, std::string_view configurationJson, HWND& viewerWindowOut) noexcept -> bool
    {
        viewerWindowOut = nullptr;

        const std::string configurationText(configurationJson);
        const HRESULT configHr = info->SetConfiguration(configurationText.c_str());
        Check(SUCCEEDED(configHr), std::format(L"ViewerText accepts diff configuration for '{}'", path.filename().wstring()), success);
        if (FAILED(configHr))
        {
            return false;
        }

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);

        const std::wstring pathText = path.wstring();
        ViewerOpenContext context{};
        context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName = L"File System";
        context.focusedPath    = pathText.c_str();

        const HRESULT openHr = viewer->Open(&context);
        Check(SUCCEEDED(openHr), std::format(L"ViewerText opens '{}'", path.filename().wstring()), success);
        if (FAILED(openHr))
        {
            return false;
        }

        const bool openedWindow = PumpUntil(
            [&]() noexcept
        {
            viewerWindowOut = FindNewVisibleWindowByClass(kViewerTextWindowClassName, existingWindows);
            return viewerWindowOut != nullptr;
        },
            5000ms);
        Check(openedWindow, std::format(L"ViewerText window becomes visible for '{}'", path.filename().wstring()), success);
        return openedWindow && viewerWindowOut != nullptr;
    };

    auto closeViewerWindow = [&](HWND viewerWindow) noexcept
    {
        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), L"ViewerText diff runtime validation window closes cleanly", success);
        Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
              L"ViewerText diff runtime validation window is destroyed after close",
              success);
    };

    HWND viewerWindow = nullptr;
    fileSystem.ResetCreateFileReaderCounts();
    if (! openViewerForPath(
            diffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    auto emergencyCleanup = wil::scope_exit([&]() noexcept
    {
        if (viewerWindow && IsWindow(viewerWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(viewerWindow, 5000ms));
        }
    });

    WndMsg::ViewerTextDebugSnapshot initialSnapshot{};
    const bool initialSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Text && snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.fileSectionCount == 1u && ! snapshot.diffExpandedContext && ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > 0u;
    },
                                                                5000ms,
                                                                &initialSnapshot);
    Check(initialSnapshotReady, L"ViewerText opens recognized diff files in parsed side-by-side mode", success);
    if (initialSnapshotReady)
    {
        const std::wstring_view firstLine(initialSnapshot.firstTextLine);
        const std::wstring_view secondLine(initialSnapshot.secondTextLine);
        const std::wstring_view preview(initialSnapshot.textPreview);
        Check(firstLine.find(L"old path: viewertext-diff-old.txt") != std::wstring_view::npos,
              L"ViewerText parsed diff view shows the old path as text at the top of the document",
              success);
        Check(secondLine.find(L"new path: viewertext-diff-new.txt") != std::wstring_view::npos,
              L"ViewerText parsed diff view shows the new path as text at the top of the document",
              success);
        Check(preview.find(L"Show 4 hidden lines") != std::wstring_view::npos,
              L"ViewerText hunks-only parsed diff view renders clickable hidden-context banners with explicit show text",
              success);
        Check(
            preview.find(L"@@") == std::wstring_view::npos, L"ViewerText parsed diff view hides raw hunk anchor text from the rendered document rows", success);
    }
    Check(initialSnapshot.diffParseCount == 1u, L"ViewerText keeps a single parsed diff document model after the initial parsed open", success);
    Check(initialSnapshot.diffHunkCount == 2u && initialSnapshot.activeDiffHunkIndex == 0u,
          L"ViewerText debug snapshots expose the parsed diff hunk count and the active hunk index",
          success);
    Check(initialSnapshot.addedRowCount >= 2u && initialSnapshot.removedRowCount >= 2u && initialSnapshot.bannerRowCount >= 2u &&
              initialSnapshot.headerRowCount >= 2u && initialSnapshot.styledRowCount >= 6u,
          L"ViewerText parsed diff metadata tracks semantic added, removed, header, and hunk rows",
          success);
    Check(initialSnapshot.visibleStyledRowCount > 0u && initialSnapshot.visibleAddedRowCount > 0u && initialSnapshot.visibleRemovedRowCount > 0u &&
              initialSnapshot.visibleBannerRowCount > 0u && initialSnapshot.textLastPaintUs > 0u,
          L"ViewerText debug snapshots expose visible add/remove/banner row counts and text paint timing",
          success);
    Check(initialSnapshot.paneLocalSideBySideLayout && initialSnapshot.sideBySideLeftPaneColumns > 0u && initialSnapshot.sideBySideRightPaneColumns > 0u &&
              initialSnapshot.sideBySideSeparatorColumns == 3u,
          L"ViewerText side-by-side parsed diff snapshots expose pane-local layout metadata and pane column widths",
          success);
    Check(
        initialSnapshot.diffContextUsesBaseBackground, L"ViewerText parsed diff rendering uses the normal text background for unchanged content rows", success);
    Check(initialSnapshot.diffMarkerArgb != 0u, L"ViewerText parsed diff snapshots expose a dimmed leading diff-marker color", success);
    Check(initialSnapshot.firstClickableBannerLogicalLine != static_cast<size_t>(-1),
          L"ViewerText debug snapshots expose the first clickable hidden-context banner row",
          success);
    Check(initialSnapshot.diffAddedBackgroundArgb == normalTheme.diffAddedBackgroundArgb &&
              initialSnapshot.diffRemovedBackgroundArgb == normalTheme.diffRemovedBackgroundArgb &&
              initialSnapshot.diffContextBackgroundArgb == normalTheme.diffContextBackgroundArgb &&
              initialSnapshot.diffHeaderBackgroundArgb == normalTheme.diffHeaderBackgroundArgb &&
              initialSnapshot.diffBannerBackgroundArgb == normalTheme.diffBannerBackgroundArgb,
          L"ViewerText debug snapshots expose the active semantic diff colors from the applied theme",
          success);
    Check(! initialSnapshot.themeRainbow, L"ViewerText debug snapshots report rainbow mode disabled for the normal diff theme", success);
    Check(GetMenu(viewerWindow) == nullptr,
          L"ViewerText detaches the native window menu while routing parsed diff raw-text fallback through the DxUi menu bar",
          success);
    Check(fileSystem.GetCreateFileReaderCount(oldPath) == 0u && fileSystem.GetCreateFileReaderCount(newPath) == 0u,
          L"ViewerText parsed diff open does not eagerly read referenced files while unchanged text is hidden",
          success);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_NEXT_HUNK), 0));

    WndMsg::ViewerTextDebugSnapshot firstHunkSnapshot{};
    const bool firstHunkSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                  [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > initialSnapshot.renderCount && snapshot.activeDiffHunkIndex == 0u &&
               snapshot.topVisibleLogicalLine > initialSnapshot.topVisibleLogicalLine &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"@@") == std::wstring_view::npos;
    },
                                                                  5000ms,
                                                                  &firstHunkSnapshot);
    Check(firstHunkSnapshotReady,
          L"ViewerText next-hunk commands jump from diff file headers to the first parsed hunk without exposing raw anchor rows",
          success);
    WndMsg::ViewerTextDebugSnapshot firstSplitRowSnapshot = firstHunkSnapshot;
    if (firstHunkSnapshotReady)
    {
        const HWND diffTextView = FindWindowExW(viewerWindow, nullptr, L"RedSalamander.ViewerText.TextView", nullptr);
        Check(diffTextView != nullptr, L"ViewerText hunk-visible diff validation locates the text child window for split-row inspection", success);
        if (diffTextView)
        {
            static_cast<void>(SendMessageW(diffTextView, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0), 0));

            const bool firstSplitRowReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                      [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
            {
                return snapshot.renderCount > firstHunkSnapshot.renderCount && snapshot.visibleSplitRowCount > 0u &&
                       snapshot.topVisibleLogicalLine > firstHunkSnapshot.topVisibleLogicalLine;
            },
                                                                      5000ms,
                                                                      &firstSplitRowSnapshot);
            Check(firstSplitRowReady,
                  L"ViewerText hunk-visible side-by-side diff rows expose pane-local split-row metadata when the changed row reaches the top of the viewport",
                  success);
        }
    }

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_NEXT_HUNK), 0));

    WndMsg::ViewerTextDebugSnapshot secondHunkSnapshot{};
    const bool secondHunkSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                   [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > firstSplitRowSnapshot.renderCount && snapshot.activeDiffHunkIndex == 1u &&
               snapshot.topVisibleLogicalLine > firstSplitRowSnapshot.topVisibleLogicalLine &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"@@") == std::wstring_view::npos;
    },
                                                                   5000ms,
                                                                   &secondHunkSnapshot);
    Check(secondHunkSnapshotReady, L"ViewerText next-hunk commands navigate to later parsed hunks without reparsing or surfacing raw anchor rows", success);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_PREVIOUS_HUNK), 0));

    WndMsg::ViewerTextDebugSnapshot previousHunkSnapshot{};
    const bool previousHunkSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                     [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > secondHunkSnapshot.renderCount && snapshot.activeDiffHunkIndex == 0u &&
               snapshot.topVisibleLogicalLine < secondHunkSnapshot.topVisibleLogicalLine &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"@@") == std::wstring_view::npos;
    },
                                                                     5000ms,
                                                                     &previousHunkSnapshot);
    Check(previousHunkSnapshotReady, L"ViewerText previous-hunk commands navigate back to earlier parsed hunks without exposing raw anchor rows", success);
    if (previousHunkSnapshotReady)
    {
        initialSnapshot = previousHunkSnapshot;
    }

    Check(rainbowTheme.rainbowMode != FALSE && (rainbowTheme.diffAddedBackgroundArgb != normalTheme.diffAddedBackgroundArgb ||
                                                rainbowTheme.diffRemovedBackgroundArgb != normalTheme.diffRemovedBackgroundArgb),
          L"ViewerText diff runtime validation defines a distinct rainbow-mode semantic palette",
          success);

    const HRESULT rainbowThemeHr = viewer->SetTheme(&rainbowTheme);
    Check(SUCCEEDED(rainbowThemeHr), L"ViewerText accepts the rainbow diff theme after opening", success);
    if (SUCCEEDED(rainbowThemeHr))
    {
        WndMsg::ViewerTextDebugSnapshot rainbowDiffSnapshot{};
        const bool rainbowDiffReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > initialSnapshot.renderCount && snapshot.themeRainbow &&
                   snapshot.diffAddedBackgroundArgb == rainbowTheme.diffAddedBackgroundArgb &&
                   snapshot.diffRemovedBackgroundArgb == rainbowTheme.diffRemovedBackgroundArgb &&
                   snapshot.diffContextBackgroundArgb == rainbowTheme.diffContextBackgroundArgb &&
                   snapshot.diffBannerBackgroundArgb == rainbowTheme.diffBannerBackgroundArgb && snapshot.visibleStyledRowCount > 0u;
        },
                                                                5000ms,
                                                                &rainbowDiffSnapshot);
        Check(rainbowDiffReady, L"ViewerText updates parsed diff semantic colors when switching into rainbow mode after opening", success);
        if (rainbowDiffReady)
        {
            Check(rainbowDiffSnapshot.paneLocalSideBySideLayout && rainbowDiffSnapshot.sideBySideLeftPaneColumns > 0u &&
                      rainbowDiffSnapshot.sideBySideRightPaneColumns > 0u && rainbowDiffSnapshot.sideBySideSeparatorColumns == 3u,
                  L"ViewerText keeps pane-local side-by-side layout metadata populated after switching into rainbow mode",
                  success);
        }

        const HRESULT highContrastThemeHr = viewer->SetTheme(&highContrastTheme);
        Check(SUCCEEDED(highContrastThemeHr), L"ViewerText accepts the high-contrast diff theme after opening", success);
        if (SUCCEEDED(highContrastThemeHr))
        {
            WndMsg::ViewerTextDebugSnapshot highContrastDiffSnapshot{};
            const bool highContrastDiffReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                         [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
            {
                return snapshot.renderCount > rainbowDiffSnapshot.renderCount &&
                       snapshot.diffAddedBackgroundArgb == highContrastTheme.diffAddedBackgroundArgb &&
                       snapshot.diffRemovedBackgroundArgb == highContrastTheme.diffRemovedBackgroundArgb &&
                       snapshot.diffContextBackgroundArgb == highContrastTheme.diffContextBackgroundArgb && snapshot.visibleStyledRowCount > 0u;
            },
                                                                         5000ms,
                                                                         &highContrastDiffSnapshot);
            Check(highContrastDiffReady, L"ViewerText updates parsed diff semantic colors when the theme changes after opening", success);

            const HRESULT restoreThemeHr = viewer->SetTheme(&normalTheme);
            Check(SUCCEEDED(restoreThemeHr), L"ViewerText can restore the normal diff theme after a theme switch", success);
            if (SUCCEEDED(restoreThemeHr))
            {
                WndMsg::ViewerTextDebugSnapshot restoredDiffSnapshot{};
                const bool restoredDiffReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                         [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
                {
                    return snapshot.renderCount > highContrastDiffSnapshot.renderCount &&
                           snapshot.diffAddedBackgroundArgb == normalTheme.diffAddedBackgroundArgb &&
                           snapshot.diffRemovedBackgroundArgb == normalTheme.diffRemovedBackgroundArgb &&
                           snapshot.diffContextBackgroundArgb == normalTheme.diffContextBackgroundArgb;
                },
                                                                         5000ms,
                                                                         &restoredDiffSnapshot);
                Check(restoredDiffReady, L"ViewerText restores the original semantic diff colors after returning to the normal theme", success);
                if (restoredDiffReady)
                {
                    initialSnapshot = restoredDiffSnapshot;
                }
            }
        }
    }

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_SHOW_UNCHANGED), 0));

    WndMsg::ViewerTextDebugSnapshot expandedSnapshot{};
    const bool expandedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                 [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffExpandedContext &&
               snapshot.diffReferencedFilesResolved && ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > initialSnapshot.renderCount;
    },
                                                                 5000ms,
                                                                 &expandedSnapshot);
    Check(expandedSnapshotReady, L"ViewerText can expand unchanged text from referenced files", success);
    Check(expandedSnapshot.diffParseCount == initialSnapshot.diffParseCount,
          L"ViewerText unchanged-text expansion reuses the parsed diff model instead of reparsing raw text",
          success);
    Check(fileSystem.GetCreateFileReaderCount(oldPath) > 0u && fileSystem.GetCreateFileReaderCount(newPath) > 0u,
          L"ViewerText reads referenced files when unchanged text expansion is requested",
          success);
    const size_t oldReadCountAfterExpanded = fileSystem.GetCreateFileReaderCount(oldPath);
    const size_t newReadCountAfterExpanded = fileSystem.GetCreateFileReaderCount(newPath);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_INLINE), 0));

    WndMsg::ViewerTextDebugSnapshot inlineSnapshot{};
    const bool inlineSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                               [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::Inline && snapshot.diffExpandedContext &&
               snapshot.renderCount > expandedSnapshot.renderCount;
    },
                                                               5000ms,
                                                               &inlineSnapshot);
    Check(inlineSnapshotReady, L"ViewerText can switch parsed diffs to inline mode", success);
    Check(inlineSnapshot.diffParseCount == expandedSnapshot.diffParseCount,
          L"ViewerText switching between parsed diff layouts reuses the cached diff document",
          success);
    Check(fileSystem.GetCreateFileReaderCount(oldPath) == oldReadCountAfterExpanded &&
              fileSystem.GetCreateFileReaderCount(newPath) == newReadCountAfterExpanded,
          L"ViewerText switching expanded diff layouts reuses cached referenced-file content without rereading files",
          success);

    const char* configurationJson   = nullptr;
    const HRESULT getInlineConfigHr = info->GetConfiguration(&configurationJson);
    Check(SUCCEEDED(getInlineConfigHr) && configurationJson != nullptr, L"ViewerText returns canonical configuration after switching diff layout", success);
    if (configurationJson != nullptr)
    {
        const std::string_view configurationView(configurationJson);
        Check(configurationView.find("\"diffDefaultLayout\":\"inline\"") != std::string_view::npos,
              L"ViewerText persists the inline diff layout after the menu command",
              success);
        Check(configurationView.find("\"diffContextMode\":\"fullFileWhenAvailable\"") != std::string_view::npos,
              L"ViewerText persists the unchanged-text expansion mode after the menu command",
              success);
    }

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_TEXT), 0));

    WndMsg::ViewerTextDebugSnapshot rawSnapshot{};
    const bool rawSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::RawText && snapshot.diffParsedAvailable && snapshot.visibleRowCount > 0u &&
               snapshot.renderCount > inlineSnapshot.renderCount;
    },
                                                            5000ms,
                                                            &rawSnapshot);
    Check(rawSnapshotReady, L"ViewerText routes parsed diffs back to raw text through the normal View / Text command", success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    fileSystem.ResetCreateFileReaderCounts();
    if (! openViewerForPath(
            diffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot bannerSnapshot{};
    const bool bannerSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                               [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               ! snapshot.diffExpandedContext && snapshot.firstClickableBannerLogicalLine != static_cast<size_t>(-1) && snapshot.renderCount > 0u;
    },
                                                               5000ms,
                                                               &bannerSnapshot);
    Check(bannerSnapshotReady, L"ViewerText exposes a clickable hidden-context banner in hunks-only parsed diff mode", success);
    if (bannerSnapshotReady)
    {
        static_cast<void>(
            SendMessageW(viewerWindow, WndMsg::kViewerTextDebugClickTextLogicalLine, static_cast<WPARAM>(bannerSnapshot.firstClickableBannerLogicalLine), 0));

        WndMsg::ViewerTextDebugSnapshot bannerExpandedSnapshot{};
        const bool bannerExpandedReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                   [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
                   snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffExpandedContext &&
                   snapshot.diffReferencedFilesResolved && snapshot.renderCount > bannerSnapshot.renderCount;
        },
                                                                   5000ms,
                                                                   &bannerExpandedSnapshot);
        Check(bannerExpandedReady, L"ViewerText clicking a hidden-context banner reveals unchanged text on demand", success);
        Check(fileSystem.GetCreateFileReaderCount(oldPath) > 0u && fileSystem.GetCreateFileReaderCount(newPath) > 0u,
              L"ViewerText hidden-context banner clicks trigger referenced-file reads only when unchanged text is revealed",
              success);

        const char* clickRevealConfigurationJson = nullptr;
        const HRESULT clickRevealConfigurationHr = info->GetConfiguration(&clickRevealConfigurationJson);
        Check(SUCCEEDED(clickRevealConfigurationHr) && clickRevealConfigurationJson != nullptr,
              L"ViewerText returns canonical configuration after hidden-context banner reveal",
              success);
        if (clickRevealConfigurationJson != nullptr)
        {
            const std::string_view configurationView(clickRevealConfigurationJson);
            Check(configurationView.find("\"diffContextMode\":\"fullFileWhenAvailable\"") != std::string_view::npos,
                  L"ViewerText hidden-context banner clicks persist the expanded unchanged-text mode through the normal configuration path",
                  success);
        }
    }

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    if (! openViewerForPath(
            missingDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"fullFileWhenAvailable","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot missingSnapshot{};
    const bool missingSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.diffHasPlaceholderRows && snapshot.placeholderRowCount > 0u && snapshot.placeholderBandCount > 0u &&
               ! snapshot.diffReferencedFilesResolved && snapshot.renderCount > 0u;
    },
                                                                5000ms,
                                                                &missingSnapshot);
    Check(missingSnapshotReady, L"ViewerText shows placeholder rows and background bands when referenced files cannot be resolved", success);
    Check(missingSnapshot.visibleGapHatchCount > 0u && missingSnapshot.diffGapHatchArgb != 0u,
          L"ViewerText unresolved placeholder rows render subtle hatched gap treatment instead of flat empty fills",
          success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    if (! openViewerForPath(
            multiFileDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot multiFileSnapshot{};
    const bool multiFileSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                  [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.fileSectionCount == 4u && snapshot.fileComboUsesDiffSections && snapshot.fileComboEntryCount == 4u &&
               snapshot.activeDiffSectionIndex == 0u && ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > 0u;
    },
                                                                  5000ms,
                                                                  &multiFileSnapshot);
    Check(multiFileSnapshotReady, L"ViewerText parses multi-file diffs into separate file sections", success);
    if (multiFileSnapshotReady)
    {
        const std::wstring_view preview(multiFileSnapshot.textPreview);
        Check(preview.find(L"new file mode 100644") != std::wstring_view::npos, L"ViewerText parsed diff text keeps create-file metadata visible", success);
        Check(preview.find(L"deleted file mode 100644") != std::wstring_view::npos, L"ViewerText parsed diff text keeps delete-file metadata visible", success);
        Check(preview.find(L"rename from viewertext-renamed-old.txt") != std::wstring_view::npos &&
                  preview.find(L"rename to viewertext-renamed-new.txt") != std::wstring_view::npos,
              L"ViewerText parsed diff text keeps rename metadata visible",
              success);
    }

    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSelectDiffSection, 3u, 0));

    WndMsg::ViewerTextDebugSnapshot navigatedSectionSnapshot{};
    const bool navigatedSectionReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                 [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff && snapshot.fileComboUsesDiffSections &&
               snapshot.activeDiffSectionIndex == 3u &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"old path: viewertext-second-old.txt") != std::wstring_view::npos;
    },
                                                                 5000ms,
                                                                 &navigatedSectionSnapshot);
    Check(navigatedSectionReady, L"ViewerText diff section navigation can jump to later file sections", success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    fileSystem.ResetCreateFileReaderCounts();
    if (! openViewerForPath(
            sectionScopedDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"fullFileWhenAvailable","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot sectionScopedSnapshot{};
    const bool sectionScopedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                      [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.fileSectionCount == 2u && snapshot.fileComboUsesDiffSections && snapshot.fileComboEntryCount == 2u &&
               snapshot.activeDiffSectionIndex == 0u && snapshot.diffExpandedContext && snapshot.diffReferencedFilesResolved &&
               ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > 0u;
    },
                                                                      5000ms,
                                                                      &sectionScopedSnapshot);
    Check(sectionScopedSnapshotReady, L"ViewerText expanded parsed diffs hydrate unchanged text only for the active file section on first open", success);
    Check(fileSystem.GetCreateFileReaderCount(sectionFirstOldPath) > 0u && fileSystem.GetCreateFileReaderCount(sectionFirstNewPath) > 0u,
          L"ViewerText expanded parsed diff open reads referenced files for the initially active file section",
          success);
    Check(fileSystem.GetCreateFileReaderCount(sectionSecondOldPath) == 0u && fileSystem.GetCreateFileReaderCount(sectionSecondNewPath) == 0u,
          L"ViewerText expanded parsed diff open defers referenced-file reads for later file sections",
          success);
    const size_t firstOldReadCountAfterSectionOpen = fileSystem.GetCreateFileReaderCount(sectionFirstOldPath);
    const size_t firstNewReadCountAfterSectionOpen = fileSystem.GetCreateFileReaderCount(sectionFirstNewPath);

    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSelectDiffSection, 1u, 0));

    WndMsg::ViewerTextDebugSnapshot secondSectionSnapshot{};
    const bool secondSectionSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                      [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.activeDiffSectionIndex == 1u &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"old path: viewertext-section-second-old.txt") != std::wstring_view::npos;
    },
                                                                      5000ms,
                                                                      &secondSectionSnapshot);
    Check(secondSectionSnapshotReady, L"ViewerText hydrates the newly selected parsed diff section when section navigation changes in expanded mode", success);
    Check(secondSectionSnapshot.diffParseCount == sectionScopedSnapshot.diffParseCount,
          L"ViewerText section-scoped unchanged-text hydration reuses the cached parsed diff document",
          success);
    Check(fileSystem.GetCreateFileReaderCount(sectionSecondOldPath) > 0u && fileSystem.GetCreateFileReaderCount(sectionSecondNewPath) > 0u,
          L"ViewerText reads referenced files when a later expanded diff section becomes active",
          success);
    Check(fileSystem.GetCreateFileReaderCount(sectionFirstOldPath) == firstOldReadCountAfterSectionOpen &&
              fileSystem.GetCreateFileReaderCount(sectionFirstNewPath) == firstNewReadCountAfterSectionOpen,
          L"ViewerText section-scoped expansion does not reread the previously hydrated section when switching sections",
          success);
    const size_t secondOldReadCountAfterSectionJump = fileSystem.GetCreateFileReaderCount(sectionSecondOldPath);
    const size_t secondNewReadCountAfterSectionJump = fileSystem.GetCreateFileReaderCount(sectionSecondNewPath);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_INLINE), 0));

    WndMsg::ViewerTextDebugSnapshot sectionScopedInlineSnapshot{};
    const bool sectionScopedInlineSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::Inline && snapshot.diffExpandedContext &&
               snapshot.activeDiffSectionIndex == 1u && snapshot.renderCount > secondSectionSnapshot.renderCount;
    },
                                                                            5000ms,
                                                                            &sectionScopedInlineSnapshot);
    Check(
        sectionScopedInlineSnapshotReady, L"ViewerText keeps the active expanded diff section when switching layouts after section-scoped hydration", success);
    Check(fileSystem.GetCreateFileReaderCount(sectionSecondOldPath) == secondOldReadCountAfterSectionJump &&
              fileSystem.GetCreateFileReaderCount(sectionSecondNewPath) == secondNewReadCountAfterSectionJump,
          L"ViewerText layout switches reuse the active section's referenced-file cache after on-demand section hydration",
          success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    fileSystem.ResetCreateFileReaderCounts();
    const uintmax_t viewportOldSize = std::filesystem::file_size(viewportOldPath, ec);
    const uintmax_t viewportNewSize = std::filesystem::file_size(viewportNewPath, ec);
    if (! openViewerForPath(
            viewportLazyDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"fullFileWhenAvailable","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot viewportInitialSnapshot{};
    const bool viewportInitialSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                        [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.diffExpandedContext && snapshot.diffReferencedFilesResolved && snapshot.deferredContextRowCount > 0u &&
               snapshot.hydratedLogicalLineEndExclusive > snapshot.hydratedLogicalLineStart && snapshot.hydratedLogicalLineEndExclusive < 260u &&
               snapshot.renderCount > 0u;
    },
                                                                        5000ms,
                                                                        &viewportInitialSnapshot);
    Check(viewportInitialSnapshotReady,
          L"ViewerText bounds expanded unchanged-text materialization to a viewport-sized logical window inside the active section",
          success);
    if (viewportInitialSnapshotReady)
    {
        const size_t hydratedSpan = viewportInitialSnapshot.hydratedLogicalLineEndExclusive - viewportInitialSnapshot.hydratedLogicalLineStart;
        Check(hydratedSpan < 128u, L"ViewerText viewport-lazy hydration keeps the initial logical hydration window bounded", success);
        Check(viewportInitialSnapshot.contextRowCount > 0u && viewportInitialSnapshot.visibleContextRowCount > 0u &&
                  viewportInitialSnapshot.diffContextUsesBaseBackground,
              L"ViewerText expanded unchanged-text snapshots expose semantic context-row counts while rendering unchanged rows on the base text background",
              success);
    }
    const size_t viewportOldBytesAfterOpen = fileSystem.GetReadByteCount(viewportOldPath);
    const size_t viewportNewBytesAfterOpen = fileSystem.GetReadByteCount(viewportNewPath);
    Check(viewportOldBytesAfterOpen > 0u && viewportOldBytesAfterOpen < viewportOldSize && viewportNewBytesAfterOpen > 0u &&
              viewportNewBytesAfterOpen < viewportNewSize,
          L"ViewerText viewport-lazy open reads only a bounded prefix of large referenced files instead of loading them fully",
          success);
    Check(viewportInitialSnapshot.referencedBytesRead >= (viewportOldBytesAfterOpen + viewportNewBytesAfterOpen),
          L"ViewerText debug snapshot reports referenced-file bytes read for viewport-lazy hydration",
          success);

    HWND viewportTextView = FindFirstChildWindowByClass(viewerWindow, L"RedSalamander.ViewerText.TextView");
    if (! viewportTextView)
    {
        viewportTextView = FindFirstVisibleChildWindowWithStyle(viewerWindow, WS_VSCROLL);
    }
    Check(viewportTextView != nullptr, L"ViewerText viewport-lazy hydration test locates the visible text child window", success);
    if (! viewportTextView)
    {
        return false;
    }

    const uint64_t viewportRenderBeforeScroll = viewportInitialSnapshot.renderCount;
    for (int step = 0; step < 8; ++step)
    {
        static_cast<void>(SendMessageW(viewportTextView, WM_VSCROLL, MAKEWPARAM(SB_PAGEDOWN, 0), 0));
    }

    WndMsg::ViewerTextDebugSnapshot viewportScrolledSnapshot{};
    const bool viewportScrolledSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                         [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.renderCount > viewportRenderBeforeScroll &&
               snapshot.topVisibleLogicalLine > viewportInitialSnapshot.topVisibleLogicalLine &&
               snapshot.hydratedLogicalLineStart > viewportInitialSnapshot.hydratedLogicalLineStart;
    },
                                                                         5000ms,
                                                                         &viewportScrolledSnapshot);
    Check(viewportScrolledSnapshotReady, L"ViewerText scroll-driven viewport hydration rehydrates later unchanged rows when they enter view", success);
    Check(viewportScrolledSnapshot.deferredContextRowCount > 0u,
          L"ViewerText viewport-driven rehydration keeps far-away unchanged rows deferred after scrolling",
          success);
    Check(fileSystem.GetReadByteCount(viewportOldPath) > viewportOldBytesAfterOpen && fileSystem.GetReadByteCount(viewportNewPath) > viewportNewBytesAfterOpen,
          L"ViewerText viewport-driven rehydration reads additional referenced-file bytes when the visible logical window advances",
          success);
    Check(fileSystem.GetReadByteCount(viewportOldPath) < viewportOldSize && fileSystem.GetReadByteCount(viewportNewPath) < viewportNewSize,
          L"ViewerText viewport-driven rehydration still avoids full referenced-file reads when the viewport stays far from EOF",
          success);

    const size_t viewportOldBytesAfterScroll     = fileSystem.GetReadByteCount(viewportOldPath);
    const size_t viewportNewBytesAfterScroll     = fileSystem.GetReadByteCount(viewportNewPath);
    const uint64_t viewportRenderBeforeBacktrack = viewportScrolledSnapshot.renderCount;
    for (int step = 0; step < 8; ++step)
    {
        static_cast<void>(SendMessageW(viewportTextView, WM_VSCROLL, MAKEWPARAM(SB_PAGEUP, 0), 0));
    }

    WndMsg::ViewerTextDebugSnapshot viewportBacktrackedSnapshot{};
    const bool viewportBacktrackedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.renderCount > viewportRenderBeforeBacktrack &&
               snapshot.topVisibleLogicalLine < viewportScrolledSnapshot.topVisibleLogicalLine;
    },
                                                                            5000ms,
                                                                            &viewportBacktrackedSnapshot);
    Check(viewportBacktrackedSnapshotReady, L"ViewerText can revisit earlier viewport-hydrated rows after scrolling back up", success);
    Check(fileSystem.GetReadByteCount(viewportOldPath) == viewportOldBytesAfterScroll &&
              fileSystem.GetReadByteCount(viewportNewPath) == viewportNewBytesAfterScroll,
          L"ViewerText revisiting an already hydrated viewport range reuses cached referenced-file content without rereading files",
          success);
    Check(viewportBacktrackedSnapshot.referencedBytesRead == viewportScrolledSnapshot.referencedBytesRead,
          L"ViewerText debug snapshot keeps referenced-file byte counts stable when revisiting an already read viewport range",
          success);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_WRAP), 0));

    WndMsg::ViewerTextDebugSnapshot wrapDisabledSnapshot{};
    const bool wrapDisabledSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                     [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > viewportBacktrackedSnapshot.renderCount &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.textLeftColumn == 0u;
    },
                                                                     5000ms,
                                                                     &wrapDisabledSnapshot);
    Check(wrapDisabledSnapshotReady, L"ViewerText can disable wrapping for expanded side-by-side diffs before horizontal scrolling", success);
    SCROLLINFO sideBySideHScrollInfo{};
    sideBySideHScrollInfo.cbSize = sizeof(sideBySideHScrollInfo);
    sideBySideHScrollInfo.fMask  = SIF_PAGE | SIF_RANGE | SIF_POS;
    static_cast<void>(GetScrollInfo(viewportTextView, SB_HORZ, &sideBySideHScrollInfo));
    Check(sideBySideHScrollInfo.nPage > 0u, L"ViewerText exposes a live horizontal scrollbar page size for wrap-disabled side-by-side diffs", success);

    static_cast<void>(SendMessageW(viewportTextView, WM_HSCROLL, MAKEWPARAM(SB_PAGERIGHT, 0), 0));

    WndMsg::ViewerTextDebugSnapshot hScrolledSnapshot{};
    const bool hScrolledSnapshotReady = WaitForViewerTextSnapshot(viewerWindow, [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept {
        return snapshot.renderCount > wrapDisabledSnapshot.renderCount && snapshot.textLeftColumn > 0u;
    }, 5000ms, &hScrolledSnapshot);
    Check(hScrolledSnapshotReady, L"ViewerText horizontal scrolling advances the pane-local left column when wrapping is disabled", success);
    Check(hScrolledSnapshot.firstVisibleSplitLeftPaneColumnStart > 0u && hScrolledSnapshot.firstVisibleSplitRightPaneColumnStart > 0u,
          L"ViewerText side-by-side horizontal scrolling advances both diff panes together",
          success);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_INLINE), 0));

    WndMsg::ViewerTextDebugSnapshot inlineWrapDisabledSnapshot{};
    const bool inlineWrapDisabledSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                           [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > hScrolledSnapshot.renderCount && snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::Inline &&
               snapshot.textLeftColumn == 0u;
    },
                                                                           5000ms,
                                                                           &inlineWrapDisabledSnapshot);
    Check(inlineWrapDisabledSnapshotReady, L"ViewerText resets the horizontal scroll position when switching from side-by-side to inline diff mode", success);

    SCROLLINFO inlineHScrollInfo{};
    inlineHScrollInfo.cbSize = sizeof(inlineHScrollInfo);
    inlineHScrollInfo.fMask  = SIF_PAGE | SIF_RANGE | SIF_POS;
    static_cast<void>(GetScrollInfo(viewportTextView, SB_HORZ, &inlineHScrollInfo));
    Check(inlineHScrollInfo.nPage > sideBySideHScrollInfo.nPage,
          L"ViewerText refreshes the horizontal scrollbar page size when switching from side-by-side to inline diff mode",
          success);

    static_cast<void>(SendMessageW(viewportTextView, WM_HSCROLL, MAKEWPARAM(SB_PAGERIGHT, 0), 0));

    WndMsg::ViewerTextDebugSnapshot inlineHScrolledSnapshot{};
    const bool inlineHScrolledSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                        [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > inlineWrapDisabledSnapshot.renderCount && snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::Inline &&
               snapshot.textLeftColumn > 0u && snapshot.topVisibleSegmentColumnStart > 0u;
    },
                                                                        5000ms,
                                                                        &inlineHScrolledSnapshot);
    Check(inlineHScrolledSnapshotReady, L"ViewerText horizontal scrolling advances inline diffs after switching presentation modes", success);

    static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(IDM_VIEWER_VIEW_DIFF_SIDE_BY_SIDE), 0));

    WndMsg::ViewerTextDebugSnapshot sideBySideResyncedSnapshot{};
    const bool sideBySideResyncedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                           [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.renderCount > inlineHScrolledSnapshot.renderCount && snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide &&
               snapshot.textLeftColumn == 0u;
    },
                                                                           5000ms,
                                                                           &sideBySideResyncedSnapshot);
    Check(sideBySideResyncedSnapshotReady, L"ViewerText refreshes horizontal scrolling again when switching back to side-by-side diff mode", success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    const uintmax_t largeBufferedDiffSize = std::filesystem::file_size(largeBufferedDiffPath, ec);
    Check(largeBufferedDiffSize > (4u * 1024u * 1024u), L"ViewerText large buffered diff test file exceeds the old 4 MiB parsed-diff threshold", success);

    if (! openViewerForPath(
            largeBufferedDiffPath,
            R"json({"textBufferMiB":4,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot largeBufferedSnapshot{};
    const bool largeBufferedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                      [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.fileSectionCount == 1u && ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > 0u;
    },
                                                                      15000ms,
                                                                      &largeBufferedSnapshot);
    Check(
        largeBufferedSnapshotReady, L"ViewerText parses fully buffered extension-recognized diffs even when they exceed the normal text buffer size", success);
    if (largeBufferedSnapshotReady)
    {
        Check(std::wstring_view(largeBufferedSnapshot.firstTextLine).find(L"old path: viewertext-large-old.txt") != std::wstring_view::npos,
              L"ViewerText keeps path summary text when opening a fully buffered larger diff in parsed mode",
              success);
    }

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    if (! openViewerForPath(
            largeSniffedDiffPath,
            R"json({"textBufferMiB":4,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot largeSniffedSnapshot{};
    const bool largeSniffedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                     [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
               snapshot.fileSectionCount == 1u && ! snapshot.diffHasPlaceholderRows && snapshot.renderCount > 0u;
    },
                                                                     15000ms,
                                                                     &largeSniffedSnapshot);
    Check(largeSniffedSnapshotReady,
          L"ViewerText parses large diff-like text files detected by header sniff even when they exceed the normal text buffer size",
          success);
    if (largeSniffedSnapshotReady)
    {
        Check(std::wstring_view(largeSniffedSnapshot.firstTextLine).find(L"old path: viewertext-large-old.txt") != std::wstring_view::npos,
              L"ViewerText keeps path summary text when opening a header-sniffed fully buffered larger diff in parsed mode",
              success);
    }

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    const uintmax_t largeStreamedDiffSize = std::filesystem::file_size(largeStreamedDiffPath, ec);
    Check(largeStreamedDiffSize > (16u * 1024u * 1024u),
          L"ViewerText large streamed multi-file diff test file exceeds the fully buffered parsed-diff cap",
          success);

    if (! openViewerForPath(
            largeStreamedDiffPath,
            R"json({"textBufferMiB":4,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot largeStreamedSnapshot{};
    const bool largeStreamedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                      [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::RawText && ! snapshot.diffParsedAvailable &&
               snapshot.fileComboUsesDiffSections && snapshot.fileSectionCount == 4u && snapshot.fileComboEntryCount == 4u &&
               snapshot.activeDiffSectionIndex == 0u && snapshot.visibleRowCount > 0u && snapshot.renderCount > 0u;
    },
                                                                      15000ms,
                                                                      &largeStreamedSnapshot);
    Check(largeStreamedSnapshotReady, L"ViewerText keeps a bounded file-section index for large streamed diffs that fall back to raw text", success);
    if (largeStreamedSnapshotReady)
    {
        Check(std::wstring_view(largeStreamedSnapshot.topVisibleTextLine).find(L"diff --git a/viewertext-stream-first-old.txt") != std::wstring_view::npos,
              L"ViewerText raw streamed diff section index opens at the first indexed file section",
              success);
    }

    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSelectDiffSection, 3u, 0));

    WndMsg::ViewerTextDebugSnapshot largeStreamedNavigatedSnapshot{};
    const bool largeStreamedNavigatedReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                       [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::RawText && ! snapshot.diffParsedAvailable &&
               snapshot.fileComboUsesDiffSections && snapshot.activeDiffSectionIndex == 3u &&
               std::wstring_view(snapshot.topVisibleTextLine).find(L"diff --git a/viewertext-stream-fourth-old.txt") != std::wstring_view::npos;
    },
                                                                       15000ms,
                                                                       &largeStreamedNavigatedSnapshot);
    Check(largeStreamedNavigatedReady, L"ViewerText raw streamed diff section navigation can jump to later indexed file sections", success);

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;

    info.reset();
    viewer.reset();

    const HRESULT recreateHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(recreateHr) && viewer != nullptr, L"ViewerText can recreate a fresh viewer instance before malformed diff fallback validation", success);
    if (FAILED(recreateHr) || ! viewer)
    {
        return false;
    }

    const HRESULT requeryInfoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
    Check(SUCCEEDED(requeryInfoHr) && info != nullptr,
          L"ViewerText recreated viewer instance exposes IInformations for malformed diff fallback validation",
          success);
    if (FAILED(requeryInfoHr) || ! info)
    {
        return false;
    }

    const HRESULT malformedThemeHr = viewer->SetTheme(&normalTheme);
    Check(SUCCEEDED(malformedThemeHr), L"ViewerText accepts the normal theme on the recreated viewer before malformed diff fallback validation", success);
    if (FAILED(malformedThemeHr))
    {
        return false;
    }

    if (! openViewerForPath(
            malformedDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot malformedSnapshot{};
    const bool malformedSnapshotReady = WaitForViewerTextSnapshot(viewerWindow,
                                                                  [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::RawText && ! snapshot.diffParsedAvailable &&
               snapshot.visibleRowCount > 0u && snapshot.renderCount > 0u;
    },
                                                                  5000ms,
                                                                  &malformedSnapshot);
    Check(malformedSnapshotReady, L"ViewerText falls back to raw text when a diff hunk header is malformed", success);
    if (malformedSnapshotReady)
    {
        const std::wstring_view preview(malformedSnapshot.textPreview);
        Check(preview.find(L"@@ -bad +1 @@") != std::wstring_view::npos,
              L"ViewerText raw-text fallback keeps the malformed hunk header visible for inspection",
              success);
    }

    closeViewerWindow(viewerWindow);
    viewerWindow = nullptr;
    return success;
}
#endif

[[nodiscard]] bool TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerSpace test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerSpace.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerSpace.dll is present for runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSpace.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSpace factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSpacePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSpace factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSpaceWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerSpaceTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path childDir = tempDir / L"folder-a";
    static_cast<void>(std::filesystem::create_directories(childDir, ec));
    const std::filesystem::path samplePath = WriteUtf8TextFile(tempDir / L"sample.txt", "viewer space sample\n");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
        static_cast<void>(std::filesystem::remove(childDir, cleanupEc));
        static_cast<void>(std::filesystem::remove(tempDir, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring tempDirText = tempDir.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.focusedPath    = tempDirText.c_str();

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerSpace window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSpaceWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        8000ms);
    Check(openedWindow, L"ViewerSpace window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    CheckDxNativeMenuBar(viewerWindow, L"ViewerSpace", success);
    CheckPlainMenuModelContract(viewerWindow, L"ViewerSpace", success);
    Check(CountActuallyVisibleChildWindows(viewerWindow) == 1u,
          L"ViewerSpace only exposes the DxUi menu bar child and no visible fallback child surface",
          success);
    Check(SetFocus(viewerWindow) != nullptr, L"ViewerSpace window accepts keyboard focus", success);
    const LRESULT escapeHandled = SendMessageW(viewerWindow, WM_KEYDOWN, VK_ESCAPE, 0);
    static_cast<void>(escapeHandled);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerSpace Escape closes the window cleanly", success);
    return success;
}

[[nodiscard]] bool TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerVLC test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerVLC.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerVLC.dll is present for runtime validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerVLC.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerVLC factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerVlcPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerVLC factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerVLCWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerVLCTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path samplePath = WriteTinyWaveFile(tempDir / L"sample.wav");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
        static_cast<void>(std::filesystem::remove(tempDir, cleanupEc));
    });

    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.focusedPath = samplePathText.c_str();

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerVLC window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerVLCWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        8000ms);
    Check(openedWindow, L"ViewerVLC window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const HWND hudWindow = FindFirstVisibleChildWindowWithStyle(viewerWindow, WS_TABSTOP);
    Check(hudWindow != nullptr && IsWindow(hudWindow) != FALSE, L"ViewerVLC exposes a visible HUD child that can take focus", success);
    Check(CountActuallyVisibleChildWindows(viewerWindow) >= 2u, L"ViewerVLC exposes visible video and HUD child surfaces", success);
    if (hudWindow)
    {
        Check(SetFocus(viewerWindow) != nullptr, L"ViewerVLC window accepts keyboard focus", success);

        HWND initialFocus          = nullptr;
        const bool gotInitialFocus = PumpUntil(
            [&]() noexcept
        {
            initialFocus = GetFocus();
            return initialFocus != nullptr && initialFocus != viewerWindow;
        },
            2000ms);
        Check(gotInitialFocus, L"ViewerVLC routes initial focus into a visible child surface", success);

        if (gotInitialFocus && initialFocus)
        {
            const LRESULT tabHandled = SendMessageW(initialFocus, WM_KEYDOWN, VK_TAB, 0);
            static_cast<void>(tabHandled);
            const bool movedToHud = PumpUntil([&]() noexcept { return GetFocus() == hudWindow; }, 2000ms);
            Check(movedToHud, L"ViewerVLC Tab transfers focus from the video surface into the HUD", success);

            const LRESULT secondTabHandled = SendMessageW(hudWindow, WM_KEYDOWN, VK_TAB, 0);
            static_cast<void>(secondTabHandled);
            const bool stayedOnHud = PumpUntil([&]() noexcept { return GetFocus() == hudWindow; }, 2000ms);
            Check(stayedOnHud, L"ViewerVLC HUD keeps keyboard focus while cycling its internal controls", success);
        }
    }

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerVLC window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerVLC window closes cleanly", success);
    return success;
}

[[nodiscard]] bool TestViewerVlcConfigurationPersistsLastVolumeAndMute() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerVLC configuration test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerVLC.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerVLC.dll is present for configuration validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerVLC.dll loads successfully for configuration validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerVLC factory export is available for configuration validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerVlcPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerVLC factory creates an IViewer for configuration validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), informations.put_void());
    Check(SUCCEEDED(infoHr) && informations != nullptr, L"ViewerVLC exposes IInformations for configuration validation", success);
    if (FAILED(infoHr) || ! informations)
    {
        return false;
    }

    const char* schema = nullptr;
    const HRESULT schemaHr = informations->GetConfigurationSchema(&schema);
    Check(SUCCEEDED(schemaHr) && schema != nullptr, L"ViewerVLC configuration schema is available", success);
    if (SUCCEEDED(schemaHr) && schema)
    {
        const std::string_view schemaText(schema);
        Check(schemaText.find("\"lastVolumePercent\"") != std::string_view::npos,
              L"ViewerVLC schema exposes last volume persistence",
              success);
        Check(schemaText.find("\"muted\"") != std::string_view::npos, L"ViewerVLC schema exposes mute persistence", success);
    }

    constexpr char kSavedVolumeJson[] = R"json({"lastVolumePercent":37,"muted":true})json";
    const HRESULT setHr               = informations->SetConfiguration(kSavedVolumeJson);
    Check(SUCCEEDED(setHr), L"ViewerVLC accepts persisted volume and mute configuration", success);

    const char* savedJson = nullptr;
    const HRESULT getHr = informations->GetConfiguration(&savedJson);
    Check(SUCCEEDED(getHr) && savedJson != nullptr, L"ViewerVLC returns normalized persisted configuration", success);
    if (SUCCEEDED(getHr) && savedJson)
    {
        const std::string_view savedText(savedJson);
        Check(savedText.find("\"lastVolumePercent\":37") != std::string_view::npos,
              L"ViewerVLC keeps the last volume in its persisted configuration",
              success);
        Check(savedText.find("\"muted\":true") != std::string_view::npos,
              L"ViewerVLC keeps mute state in its persisted configuration",
              success);
    }

    BOOL somethingToSave = FALSE;
    const HRESULT saveHr = informations->SomethingToSave(&somethingToSave);
    Check(SUCCEEDED(saveHr) && somethingToSave != FALSE, L"ViewerVLC marks non-default volume state for persistence", success);
    return success;
}

[[nodiscard]] bool TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerVLC HUD contract test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerVLC.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerVLC.dll is present for HUD contract validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerVLC.dll loads successfully for HUD contract validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerVLC factory export is available for HUD contract validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerVlcPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerVLC factory creates an IViewer for HUD contract validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> informations;
    static_cast<void>(viewer->QueryInterface(__uuidof(IInformations), informations.put_void()));
    if (informations)
    {
        static_cast<void>(informations->SetConfiguration(R"json({"autoDetectVlc":false})json"));
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerVLCWindowClassName);

    const std::filesystem::path tempDir = buildDir / L"ViewerVLCTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path samplePath = WriteTinyWaveFile(tempDir / L"sample-hud.wav");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
        static_cast<void>(std::filesystem::remove(tempDir, cleanupEc));
    });

    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.focusedPath = samplePathText.c_str();

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerVLC HUD contract window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerVLCWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        8000ms);
    Check(openedWindow, L"ViewerVLC HUD contract window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerVlcDebugSnapshot snapshot{};
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE,
          L"ViewerVLC debug snapshot is available",
          success);
    Check(snapshot.hasVolumeMuteButton, L"ViewerVLC HUD exposes a speaker mute button", success);
    Check(snapshot.hasVolumeSlider, L"ViewerVLC HUD exposes a horizontal volume slider", success);
    Check(snapshot.snapshotWidth > 0 && snapshot.snapshotHeight > 0, L"ViewerVLC snapshot uses the visible video surface size", success);

    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugForceLoadingVisible, 0, 0) != FALSE,
          L"ViewerVLC debug can force delayed loading overlay",
          success);
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.loadingActive && snapshot.loadingVisible, L"ViewerVLC shows a loading overlay once VLC init exceeds the delay", success);

    WndMsg::ViewerVlcDebugPlaybackState playbackState{};
    playbackState.lengthMs = 120'000;
    playbackState.timeMs   = 60'000;
    playbackState.volume   = 37;
    playbackState.muted    = false;
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetPlaybackState, 0, reinterpret_cast<LPARAM>(&playbackState)) != FALSE,
          L"ViewerVLC debug playback state can seed wheel and volume checks",
          success);

    WndMsg::ViewerVlcDebugWheel wheel{};
    wheel.wheelDelta = WHEEL_DELTA;
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugWheel, 0, reinterpret_cast<LPARAM>(&wheel)));
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.timeMs == 70'000, L"ViewerVLC mouse wheel seeks by the normal ten-second step", success);

    wheel.wheelDelta = -WHEEL_DELTA;
    wheel.shift      = true;
    wheel.ctrl       = false;
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugWheel, 0, reinterpret_cast<LPARAM>(&wheel)));
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.timeMs == 67'000, L"ViewerVLC Shift+wheel seeks by the small three-second step", success);

    wheel.wheelDelta = WHEEL_DELTA;
    wheel.shift      = false;
    wheel.ctrl       = true;
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugWheel, 0, reinterpret_cast<LPARAM>(&wheel)));
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.timeMs == 120'000, L"ViewerVLC Ctrl+wheel seeks by the large one-minute step and clamps to the end", success);

    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugToggleMute, 0, 0));
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.muted && snapshot.volume == 37, L"ViewerVLC speaker click mutes without forgetting the slider volume", success);

    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugToggleMute, 0, 0));
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(! snapshot.muted && snapshot.volume == 37, L"ViewerVLC speaker click restores the previous volume", success);

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerVLC HUD contract window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerVLC HUD contract window closes cleanly", success);
    return success;
}

[[nodiscard]] bool CheckViewerTextPromptUsesDxUiHostAndClosesCleanly(UINT commandId,
                                                                     std::wstring_view promptName,
                                                                     std::wstring_view expectedInitialText,
                                                                     UINT closeVirtualKey) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), std::format(L"{} test executable path resolves", promptName), success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), std::format(L"ViewerText.dll is present for {} runtime validation", promptName), success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), std::format(L"ViewerText.dll loads successfully for {}", promptName), success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, std::format(L"ViewerText factory export is available for {}", promptName), success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, std::format(L"ViewerText factory creates an IViewer instance for {}", promptName), success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    CloseVisibleWindowsByClass(kViewerTextPromptWindowClassName);
    CloseVisibleWindowsByClass(kViewerTextWindowClassName);
    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));

    const std::vector<HWND> existingViewerWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);
    const std::vector<HWND> existingPromptWindows = CollectVisibleWindowsByClass(kViewerTextPromptWindowClassName);
    HWND viewerWindow                             = nullptr;
    HWND promptWindow                             = nullptr;
    auto emergencyCleanup                         = wil::scope_exit([&]() noexcept
    {
        if (promptWindow && IsWindow(promptWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(promptWindow, 3000ms));
        }

        CloseVisibleWindowsByClass(kViewerTextPromptWindowClassName);

        if (viewerWindow && IsWindow(viewerWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(viewerWindow, 5000ms));
        }

        CloseVisibleWindowsByClass(kViewerTextWindowClassName);
    });

    const std::filesystem::path tempDir = buildDir / L"ViewerTextTests";
    std::error_code ec;
    static_cast<void>(std::filesystem::create_directories(tempDir, ec));
    const std::filesystem::path firstPath  = WriteUtf8TextFile(tempDir / L"viewertext-first.txt", "first viewer text file\n");
    const std::filesystem::path secondPath = WriteUtf8TextFile(tempDir / L"viewertext-second.txt", "second viewer text file\n");
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(firstPath, cleanupEc));
        static_cast<void>(std::filesystem::remove(secondPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring firstPathText  = firstPath.wstring();
    const std::wstring secondPathText = secondPath.wstring();
    const wchar_t* otherFiles[]       = {firstPathText.c_str(), secondPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = firstPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), std::format(L"ViewerText window open succeeds for {}", promptName), success);
    if (FAILED(openHr))
    {
        return false;
    }

    const bool openedViewerWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerTextWindowClassName, existingViewerWindows);
        return viewerWindow != nullptr;
    },
        10000ms);
    Check(openedViewerWindow, std::format(L"ViewerText window becomes visible for {}", promptName), success);
    if (! openedViewerWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    HWND comboHost            = nullptr;
    const bool comboHostReady = PumpUntil(
        [&]() noexcept
    {
        comboHost = GetDlgItem(viewerWindow, kViewerTextFileComboId);
        return comboHost != nullptr && IsWindowVisible(comboHost) != FALSE;
    },
        5000ms);
    Check(comboHostReady, std::format(L"ViewerText DX combo host is ready before {} prompt dispatch", promptName), success);
    if (! comboHostReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 750ms));

    struct PromptProbeResult
    {
        bool dispatchedCommand   = false;
        bool openedPromptWindow  = false;
        bool accessibilityPassed = false;
        bool closedPromptWindow  = false;
        HWND promptWindow        = nullptr;
    } probeResult{};
    std::atomic<bool> probeFinished = false;

    std::thread promptProbeThread([&]() noexcept
    {
        const auto releaseCom = wil::scope_exit([]() noexcept { CoUninitialize(); });
        static_cast<void>(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));

        const auto dispatchPromptCommand = [&]() noexcept
        {
            if (PostMessageW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(commandId), 0) != FALSE)
            {
                return true;
            }

            DWORD_PTR sendResult = 0;
            return SendMessageTimeoutW(viewerWindow, WM_COMMAND, static_cast<WPARAM>(commandId), 0, SMTO_ABORTIFHUNG | SMTO_NORMAL, 2000, &sendResult) != 0;
        };

        for (size_t attempt = 0; attempt < 3u && ! probeResult.openedPromptWindow; ++attempt)
        {
            probeResult.dispatchedCommand = dispatchPromptCommand() || probeResult.dispatchedCommand;

            const bool foundPrompt = PumpUntil(
                [&]() noexcept
            {
                probeResult.promptWindow = FindNewVisibleWindowByClass(kViewerTextPromptWindowClassName, existingPromptWindows);
                return probeResult.promptWindow != nullptr;
            },
                attempt == 0u ? 6000ms : 8000ms);
            probeResult.openedPromptWindow = probeResult.openedPromptWindow || foundPrompt;

            if (! probeResult.openedPromptWindow)
            {
                const std::vector<HWND> currentPromptWindows = CollectVisibleWindowsByClass(kViewerTextPromptWindowClassName);
                if (! currentPromptWindows.empty())
                {
                    probeResult.promptWindow       = currentPromptWindows.back();
                    probeResult.openedPromptWindow = true;
                }
            }

            if (! probeResult.openedPromptWindow)
            {
                std::this_thread::sleep_for(250ms);
            }
        }

        if (probeResult.openedPromptWindow && probeResult.promptWindow)
        {
            bool promptSuccess = true;
            CheckViewerTextPromptAccessibility(probeResult.promptWindow, promptName, expectedInitialText, 2000ms, promptSuccess);
            probeResult.accessibilityPassed = promptSuccess;
            if (closeVirtualKey != 0)
            {
                static_cast<void>(PostMessageW(probeResult.promptWindow, WM_KEYDOWN, static_cast<WPARAM>(closeVirtualKey), 0));
                static_cast<void>(PostMessageW(probeResult.promptWindow, WM_KEYUP, static_cast<WPARAM>(closeVirtualKey), 0));
            }

            probeResult.closedPromptWindow =
                PumpUntil([&]() noexcept { return ! probeResult.promptWindow || IsWindow(probeResult.promptWindow) == FALSE; }, 1500ms);
            if (! probeResult.closedPromptWindow)
            {
                probeResult.closedPromptWindow = RequestCloseWindow(probeResult.promptWindow, 3000ms);
            }
        }
        else
        {
            CloseVisibleWindowsByClass(kViewerTextPromptWindowClassName);
        }

        probeFinished.store(true);
    });

    const bool finishedProbe = PumpUntil([&]() noexcept { return probeFinished.load(); }, 30000ms);
    if (promptProbeThread.joinable())
    {
        promptProbeThread.join();
    }

    Check(finishedProbe, std::format(L"ViewerText {} prompt probe finishes", promptName), success);
    Check(probeResult.dispatchedCommand, std::format(L"ViewerText dispatches {} command", promptName), success);
    Check(probeResult.openedPromptWindow, std::format(L"ViewerText {} prompt becomes visible", promptName), success);
    if (! probeResult.openedPromptWindow || ! probeResult.promptWindow)
    {
        return false;
    }

    success = probeResult.accessibilityPassed && success;
    Check(probeResult.closedPromptWindow, std::format(L"ViewerText {} prompt closes cleanly", promptName), success);

    static_cast<void>(PumpUntil([&]() noexcept { return IsWindowEnabled(viewerWindow) != FALSE; }, 2000ms));

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), std::format(L"ViewerText window close succeeds after {} prompt", promptName), success);
    bool closedViewerWindow = PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms);
    if (! closedViewerWindow && viewerWindow && IsWindow(viewerWindow) != FALSE)
    {
        closedViewerWindow = RequestCloseWindow(viewerWindow, 5000ms);
    }
    Check(closedViewerWindow, std::format(L"ViewerText window closes cleanly after {} prompt", promptName), success);
    emergencyCleanup.release();
    return success;
}

[[nodiscard]] bool TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly() noexcept
{
    return CheckViewerTextPromptUsesDxUiHostAndClosesCleanly(kViewerTextFindCommandId, L"Find", L"", VK_ESCAPE);
}

[[nodiscard]] bool TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly() noexcept
{
    return CheckViewerTextPromptUsesDxUiHostAndClosesCleanly(kViewerTextGotoCommandId, L"Go To", L"", VK_RETURN);
}

constexpr std::wstring_view kViewerTextFindPromptExplicitTestName = L"TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly";
constexpr std::wstring_view kViewerTextGotoPromptExplicitTestName = L"TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly";
constexpr std::wstring_view kViewerTextFindPromptInternalTestName = L"__Internal_TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly";
constexpr std::wstring_view kViewerTextGotoPromptInternalTestName = L"__Internal_TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly";

[[nodiscard]] bool TestViewerShellComboHostsLongRunOpenCloseStayStable() noexcept
{
    bool success = true;

    for (size_t cycleIndex = 0; cycleIndex < 6u; ++cycleIndex)
    {
        std::wcout << std::format(L"[INFO] Viewer shell churn cycle {}/6\n", cycleIndex + 1u);
        success = RunFilteredSelfExecutable(L"TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerVlcConfigurationPersistsLastVolumeAndMute", 120000ms, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts", 120000ms, success) && success;
    }

    return success;
}

[[nodiscard]] bool RunFullSuiteInFreshProcesses() noexcept
{
    bool success = true;

    std::vector<std::wstring_view> isolatedTests{
        L"TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo",
        L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo",
        L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo",
        L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo",
        L"TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses",
        L"TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly",
        L"TestViewerVlcConfigurationPersistsLastVolumeAndMute",
        L"TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts",
        L"TestViewerShellComboHostsLongRunOpenCloseStayStable",
    };
#ifdef _DEBUG
    isolatedTests.push_back(L"TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback");
    isolatedTests.push_back(L"TestViewerTextDiffModesAndPlaceholders");
#endif

    for (const std::wstring_view testName : isolatedTests)
    {
        success = RunFilteredSelfExecutable(testName, 120000ms, success) && success;
    }

    return success;
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    const HRESULT coHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(coHr) && coHr != RPC_E_CHANGED_MODE)
    {
        std::wcout << std::format(L"[FAIL] COM apartment initialization failed (hr=0x{:08X}).\n", static_cast<unsigned long>(coHr));
        return 1;
    }
    const auto uninitializeCom = wil::scope_exit([&]() noexcept
    {
        if (SUCCEEDED(coHr))
        {
            CoUninitialize();
        }
    });

    const std::wstring filter = (argc >= 2 && argv != nullptr && argv[1] != nullptr) ? std::wstring(argv[1]) : std::wstring{};

    bool success = true;
    if (filter.empty())
    {
        success = RunFullSuiteInFreshProcesses();
        std::wcout << (success ? L"ViewerPETests passed.\n" : L"ViewerPETests failed.\n");
        return success ? 0 : 1;
    }

    const auto shouldRun = [&](std::wstring_view testName) noexcept { return filter.empty() || filter == testName; };

    if (shouldRun(L"TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
    if (shouldRun(L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
    if (shouldRun(L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
    if (shouldRun(L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
#ifdef _DEBUG
    if (shouldRun(L"TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback"))
    {
        success = TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback() && success;
    }
    if (shouldRun(L"TestViewerTextDiffModesAndPlaceholders"))
    {
        success = TestViewerTextDiffModesAndPlaceholders() && success;
    }
#endif
    if (shouldRun(L"TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses"))
    {
        success = TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses() && success;
    }
    if (shouldRun(L"TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly"))
    {
        success = TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly() && success;
    }
    if (shouldRun(L"TestViewerVlcConfigurationPersistsLastVolumeAndMute"))
    {
        success = TestViewerVlcConfigurationPersistsLastVolumeAndMute() && success;
    }
    if (shouldRun(L"TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts"))
    {
        success = TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts() && success;
    }
    if (shouldRun(kViewerTextFindPromptExplicitTestName))
    {
        success = RunFilteredSelfExecutable(kViewerTextFindPromptInternalTestName, 30000ms, success) && success;
    }
    if (shouldRun(kViewerTextGotoPromptExplicitTestName))
    {
        success = RunFilteredSelfExecutable(kViewerTextGotoPromptInternalTestName, 30000ms, success) && success;
    }
    if (shouldRun(kViewerTextFindPromptInternalTestName))
    {
        success = TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly() && success;
    }
    if (shouldRun(kViewerTextGotoPromptInternalTestName))
    {
        success = TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly() && success;
    }
    if (shouldRun(L"TestViewerShellComboHostsLongRunOpenCloseStayStable"))
    {
        success = TestViewerShellComboHostsLongRunOpenCloseStayStable() && success;
    }
    std::wcout << (success ? L"ViewerPETests passed.\n" : L"ViewerPETests failed.\n");
    return success ? 0 : 1;
}

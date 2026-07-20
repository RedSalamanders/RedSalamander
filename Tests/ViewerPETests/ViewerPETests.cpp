#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstdint>
#include <exception>
#include <filesystem>
#include <format>
#include <fstream>
#include <iostream>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>
#include <vector>

#include <UIAutomation.h>
#include <shlwapi.h>
#include <wincodec.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "TestSupport/TestSupport.h"
#include "ViewerImgRaw/ViewerImgRaw.AsyncProtocol.h"
#include "ViewerImgRaw/ViewerImgRaw.ResourcePolicy.h"
#include "ViewerSpace/ViewerSpace.ScanPolicy.h"
#include "ViewerText/ViewerText.SafetyHelpers.h"
#include "ViewerText/resource.h"
#include "ViewerWeb/JsStringEscape.h"
#include "ViewerWeb/ViewerWebCleanupTracker.h"
#include "ViewerWeb/ViewerWebSecurity.h"
#include "WindowMessages.h"

#pragma comment(lib, "windowscodecs")
#pragma comment(lib, "shlwapi")

namespace
{
using namespace std::chrono_literals;
using RedSalamanderCreateFn = HRESULT(__stdcall*)(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);

constexpr wchar_t kViewerPEWindowClassName[]         = L"RedSalamander.ViewerPE";
constexpr wchar_t kViewerWebWindowClassName[]        = L"RedSalamander.ViewerWeb";
constexpr wchar_t kViewerImgRawWindowClassName[]     = L"RedSalamander.ViewerImgRaw";
constexpr wchar_t kViewerSpaceWindowClassName[]      = L"RedSalamander.ViewerSpace";
constexpr wchar_t kViewerVLCWindowClassName[]        = L"RedSalamander.ViewerVLC";
constexpr wchar_t kViewerVLCVideoWindowClassName[]   = L"RedSalamander.ViewerVLC.Video";
constexpr wchar_t kViewerVLCHudWindowClassName[]     = L"RedSalamander.ViewerVLC.Hud";
constexpr wchar_t kViewerTextWindowClassName[]       = L"RedSalamander.ViewerText";
constexpr wchar_t kViewerTextViewWindowClassName[]   = L"RedSalamander.ViewerText.TextView";
constexpr wchar_t kViewerTextPromptWindowClassName[] = L"RedSalamander.ViewerText.Prompt";
constexpr wchar_t kDxNativeMenuBarWindowClassName[]  = L"RedSalamander.DxNativeMenuBar";
constexpr wchar_t kNativeTooltipWindowClassName[]    = L"tooltips_class32";
constexpr wchar_t kViewerPEPluginId[]                = L"builtin/viewer-pe";
constexpr wchar_t kViewerWebPluginId[]               = L"builtin/viewer-web";
constexpr wchar_t kViewerMarkdownPluginId[]          = L"builtin/viewer-markdown";
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
constexpr std::wstring_view kViewerPEHarnessSegment{L"viewer-pe"};
constexpr char kReadOnlyFileSystemCapabilitiesJson[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": false,
    "rename": false,
    "properties": false,
    "read": true,
    "write": false
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": [], "move": [] },
    "import": { "copy": [], "move": [] }
  }
}
)json";
constexpr auto kViewerHarnessDefaultTimeout          = 120000ms;
constexpr auto kViewerShellComboLongRunTimeout       = 600000ms;
static_assert(kViewerShellComboLongRunTimeout > kViewerHarnessDefaultTimeout);

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

struct IsolatedViewerTest
{
    std::wstring_view name;
    std::chrono::milliseconds timeout;
};

[[nodiscard]] std::wstring GetEnvironmentString(std::wstring_view name)
{
    return RedSalamander::TestSupport::GetEnvironmentString(name);
}

[[nodiscard]] std::optional<std::filesystem::path> FindInstalledVlcForLoaderTest() noexcept
{
    for (const std::wstring_view variable : {std::wstring_view(L"ProgramFiles"), std::wstring_view(L"ProgramFiles(x86)")})
    {
        const std::wstring root = GetEnvironmentString(variable);
        if (root.empty())
        {
            continue;
        }

        const std::filesystem::path installDir = std::filesystem::path(root) / L"VideoLAN" / L"VLC";
        std::error_code ec;
        const bool hasLibVlc = std::filesystem::is_regular_file(installDir / L"libvlc.dll", ec);
        if (ec || ! hasLibVlc)
        {
            continue;
        }

        const bool hasVlcExe = std::filesystem::is_regular_file(installDir / L"vlc.exe", ec);
        if (ec || ! hasVlcExe)
        {
            continue;
        }

        const bool hasPlugins = std::filesystem::is_directory(installDir / L"plugins", ec);
        if (! ec && hasPlugins)
        {
            return installDir;
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring ReadProcessDllDirectoryForTest() noexcept
{
    const DWORD required = GetDllDirectoryW(0u, nullptr);
    if (required == 0u || required >= 32768u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = GetDllDirectoryW(required, value.data());
    if (written == 0u || written >= required)
    {
        return {};
    }

    value.resize(written);
    return value;
}

[[nodiscard]] std::filesystem::path AcquireViewerPETestSandbox(std::wstring_view caseName, std::error_code& ec) noexcept
{
    return RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = kViewerPEHarnessSegment, .leafSegment = caseName, .fallbackRunIdPrefix = L"viewer-pe"}, ec);
}

void PumpPendingMessages() noexcept
{
    static_cast<void>(RedSalamander::TestSupport::PumpPendingMessages());
}

template <typename TPredicate> [[nodiscard]] bool PumpUntil(TPredicate&& predicate, std::chrono::milliseconds timeout) noexcept
{
    return RedSalamander::TestSupport::PumpMessagesUntil(std::forward<TPredicate>(predicate), {.timeout = timeout, .operationName = L"ViewerPE test condition"})
        .conditionMet;
}

[[nodiscard]] std::vector<HWND> CollectWindowsByClass(std::wstring_view className, bool requireVisible) noexcept
{
    std::vector<HWND> windows;
    struct EnumArgs
    {
        std::wstring_view className;
        DWORD processId            = 0;
        bool requireVisible        = true;
        std::vector<HWND>* windows = nullptr;
    } args{className, GetCurrentProcessId(), requireVisible, &windows};

    static_cast<void>(EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* args = reinterpret_cast<EnumArgs*>(lParam);
        if (! args || (args->requireVisible && IsWindowVisible(hwnd) == FALSE))
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

[[nodiscard]] std::vector<HWND> CollectVisibleWindowsByClass(std::wstring_view className) noexcept
{
    return CollectWindowsByClass(className, true);
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
    std::wstring timeoutDiagnostic;
    const bool ready = RedSalamander::TestSupport::WaitForSnapshot<WndMsg::ViewerTextDebugSnapshot>([hwnd](WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept {
        return TryGetViewerTextDebugSnapshot(hwnd, snapshot);
    }, std::forward<Predicate>(predicate), {.timeout = timeout, .operationName = L"ViewerText debug snapshot"}, outSnapshot, &timeoutDiagnostic);
    if (! ready)
    {
        std::wcerr << timeoutDiagnostic << L'\n';
    }
    return ready;
}
#endif

#ifdef _DEBUG
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

[[nodiscard]] RECT GetChildRectInParent(HWND parent, HWND child) noexcept
{
    RECT rect{};
    if (! parent || ! child || IsWindow(parent) == FALSE || IsWindow(child) == FALSE || GetWindowRect(child, &rect) == 0)
    {
        return {};
    }

    POINT points[] = {{rect.left, rect.top}, {rect.right, rect.bottom}};
    static_cast<void>(MapWindowPoints(HWND_DESKTOP, parent, points, static_cast<UINT>(std::size(points))));
    rect.left   = points[0].x;
    rect.top    = points[0].y;
    rect.right  = points[1].x;
    rect.bottom = points[1].y;
    return rect;
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

void PrintChildWindowDiagnostics(HWND hwnd, std::wstring_view prefix) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    std::wcout << std::format(L"[INFO] {} child window diagnostics:\n", prefix);
    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM parentParam) noexcept -> BOOL
    {
        const HWND parent = reinterpret_cast<HWND>(parentParam);

        wchar_t classBuffer[128] = {};
        const int classLength    = GetClassNameW(child, classBuffer, static_cast<int>(std::size(classBuffer)));
        const std::wstring_view className =
            (classLength > 0) ? std::wstring_view(classBuffer, static_cast<size_t>(classLength)) : std::wstring_view(L"<unknown>");

        RECT rect{};
        static_cast<void>(GetWindowRect(child, &rect));
        POINT points[] = {{rect.left, rect.top}, {rect.right, rect.bottom}};
        static_cast<void>(MapWindowPoints(HWND_DESKTOP, parent, points, static_cast<UINT>(std::size(points))));

        const LONG_PTR style   = GetWindowLongPtrW(child, GWL_STYLE);
        const LONG_PTR exStyle = GetWindowLongPtrW(child, GWL_EXSTYLE);
        std::wcout << std::format(L"  - hwnd=0x{:X} class='{}' visible={} style=0x{:X} exStyle=0x{:X} parent=0x{:X} rect=({}, {}, {}, {})\n",
                                  reinterpret_cast<uintptr_t>(child),
                                  className,
                                  IsWindowVisible(child) != FALSE ? 1 : 0,
                                  static_cast<uint64_t>(style),
                                  static_cast<uint64_t>(exStyle),
                                  reinterpret_cast<uintptr_t>(GetParent(child)),
                                  points[0].x,
                                  points[0].y,
                                  points[1].x,
                                  points[1].y);
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(hwnd)));
}

bool Check(bool condition, std::wstring_view message, bool& success);
void CheckViewerSpaceTooltipOverlay(HWND viewerWindow, bool& success) noexcept;

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
        Check(IsWindow(viewerWindow) != FALSE, std::format(L"{} stays open when Escape leaves the file combo chrome", viewerName), success);

        const bool focusReturned = PumpUntil(
            [&]() noexcept
        {
            const HWND focused = GetFocus();
            return focused != comboHost && (! focused || IsChild(comboHost, focused) == FALSE);
        },
            1000ms);
        Check(focusReturned, std::format(L"{} DxUi combo host returns keyboard focus after Escape", viewerName), success);
    }

    static_cast<void>(SetFocus(comboHost));
    static_cast<void>(SendMessageW(comboHost, WM_KEYDOWN, VK_TAB, 0));
    static_cast<void>(SendMessageW(comboHost, WM_KEYUP, VK_TAB, 0));
    const bool tabFocusReturned = PumpUntil(
        [&]() noexcept
    {
        const HWND focused = GetFocus();
        return focused != comboHost && (! focused || IsChild(comboHost, focused) == FALSE);
    },
        1000ms);
    Check(tabFocusReturned, std::format(L"{} DxUi combo host returns keyboard focus after Tab wraps", viewerName), success);
}

void CheckDxComboHostCompactChrome(HWND viewerWindow, int comboControlId, std::wstring_view viewerName, bool& success) noexcept
{
    const HWND comboHost = GetDlgItem(viewerWindow, comboControlId);
    Check(comboHost != nullptr && IsWindowVisible(comboHost) != FALSE,
          std::format(L"{} exposes a visible combo host for compact chrome validation", viewerName),
          success);
    if (! comboHost || IsWindowVisible(comboHost) == FALSE)
    {
        return;
    }

    RECT comboClient{};
    static_cast<void>(GetClientRect(comboHost, &comboClient));
    const UINT dpi      = GetDpiForWindow(viewerWindow);
    const int maxHeight = MulDiv(30, static_cast<int>(dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI);
    const int height    = static_cast<int>(std::max<LONG>(0, comboClient.bottom - comboClient.top));
    Check(height <= maxHeight,
          std::format(L"{} file combo chrome uses a compact 28 DIP control height (height={}, max={})", viewerName, height, maxHeight),
          success);
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

void CheckImageRawComboHostIsInsetInsideHeader(HWND viewerWindow, bool& success) noexcept
{
    const HWND comboHost = GetDlgItem(viewerWindow, kViewerImgRawFileComboId);
    Check(comboHost != nullptr && IsWindowVisible(comboHost) != FALSE, L"ViewerImgRaw exposes a visible combo host for header inset validation", success);
    if (! comboHost || IsWindowVisible(comboHost) == FALSE)
    {
        return;
    }

    const HWND menuHost = FindFirstChildWindowByClass(viewerWindow, kDxNativeMenuBarWindowClassName);
    Check(menuHost != nullptr && IsWindowVisible(menuHost) != FALSE, L"ViewerImgRaw exposes a visible menu host for header inset validation", success);
    if (! menuHost || IsWindowVisible(menuHost) == FALSE)
    {
        return;
    }

    const RECT comboRect  = GetChildRectInParent(viewerWindow, comboHost);
    const RECT menuRect   = GetChildRectInParent(viewerWindow, menuHost);
    const UINT dpi        = GetDpiForWindow(viewerWindow);
    const int minInset    = (std::max)(1, MulDiv(1, static_cast<int>(dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI));
    const int maxInset    = (std::max)(1, MulDiv(4, static_cast<int>(dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI));
    const int actualInset = comboRect.top - menuRect.bottom;
    Check(actualInset >= minInset && actualInset <= maxInset,
          std::format(L"ViewerImgRaw combo host is tightly inset within the header background (top={}, menuBottom={}, inset={}, max={})",
                      comboRect.top,
                      menuRect.bottom,
                      actualInset,
                      maxInset),
          success);
}

void CheckEmbeddedViewerHidesStandaloneFileCombo(RedSalamanderCreateFn createFn,
                                                 const wchar_t* pluginId,
                                                 const ViewerOpenContext& sourceContext,
                                                 std::wstring_view windowClassName,
                                                 int comboControlId,
                                                 std::wstring_view viewerName,
                                                 bool& success) noexcept
{
    if (! createFn || ! pluginId)
    {
        Check(false, std::format(L"{} embedded combo normalization has a valid factory", viewerName), success);
        return;
    }

    wil::com_ptr<IViewer> embeddedViewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, pluginId, embeddedViewer.put_void());
    Check(SUCCEEDED(createHr) && embeddedViewer != nullptr, std::format(L"{} factory creates an embedded IViewer instance", viewerName), success);
    if (FAILED(createHr) || ! embeddedViewer)
    {
        return;
    }

    const std::wstring hostTitle = std::format(L"{} embedded host", viewerName);
    wil::unique_hwnd hostWindow(CreateWindowExW(0,
                                                L"STATIC",
                                                hostTitle.c_str(),
                                                WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                                                CW_USEDEFAULT,
                                                CW_USEDEFAULT,
                                                640,
                                                480,
                                                nullptr,
                                                nullptr,
                                                GetModuleHandleW(nullptr),
                                                nullptr));
    Check(hostWindow.is_valid(), std::format(L"{} embedded host window is created", viewerName), success);
    if (! hostWindow.is_valid())
    {
        return;
    }

    ViewerOpenContext embeddedContext = sourceContext;
    embeddedContext.ownerWindow       = hostWindow.get();
    embeddedContext.flags             = VIEWER_OPEN_FLAG_EMBEDDED;

    const HRESULT openHr = embeddedViewer->Open(&embeddedContext);
    Check(SUCCEEDED(openHr), std::format(L"{} embedded viewer open succeeds", viewerName), success);
    if (FAILED(openHr))
    {
        return;
    }

    HWND embeddedWindow      = nullptr;
    const bool embeddedReady = PumpUntil(
        [&]() noexcept
    {
        embeddedWindow = FindFirstChildWindowByClass(hostWindow.get(), windowClassName);
        return embeddedWindow != nullptr;
    },
        8000ms);
    Check(embeddedReady && embeddedWindow != nullptr, std::format(L"{} embedded child window becomes visible", viewerName), success);
    if (! embeddedReady || ! embeddedWindow)
    {
        static_cast<void>(embeddedViewer->Close());
        return;
    }

    if (comboControlId != 0)
    {
        const HWND comboHost = GetDlgItem(embeddedWindow, comboControlId);
        Check(comboHost == nullptr || IsWindowVisible(comboHost) == FALSE,
              std::format(L"{} embedded viewer hides the standalone filename combo host", viewerName),
              success);
    }
    Check(CountVisibleChildWindowsByClass(embeddedWindow, L"ComboBox") == 0u,
          std::format(L"{} embedded viewer has no visible legacy ComboBox child", viewerName),
          success);
    Check(GetMenu(embeddedWindow) == nullptr, std::format(L"{} embedded viewer has no top-level native menu", viewerName), success);

    const HRESULT closeHr = embeddedViewer->Close();
    Check(SUCCEEDED(closeHr), std::format(L"{} embedded viewer close succeeds", viewerName), success);
    Check(PumpUntil([&]() noexcept { return IsWindow(embeddedWindow) == FALSE; }, 5000ms),
          std::format(L"{} embedded child window closes cleanly", viewerName),
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

[[nodiscard]] std::optional<std::vector<std::byte>> ReadBinaryFile(const std::filesystem::path& path)
{
    std::ifstream stream(path, std::ios::binary | std::ios::ate);
    const std::streampos end = stream.tellg();
    if (! stream || end < static_cast<std::streampos>(0))
    {
        return std::nullopt;
    }

    std::vector<std::byte> bytes(static_cast<size_t>(end));
    stream.seekg(0, std::ios::beg);
    if (! bytes.empty())
    {
        stream.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    }
    if (! stream)
    {
        return std::nullopt;
    }
    return bytes;
}

[[nodiscard]] size_t CountViewerTextSaveTemps(const std::filesystem::path& directory, std::error_code& ec) noexcept
{
    constexpr std::wstring_view kPrefix = L".redsalamander-viewertext-save-";
    size_t count                        = 0u;
    std::filesystem::directory_iterator iterator(directory, ec);
    const std::filesystem::directory_iterator end;
    while (! ec && iterator != end)
    {
        if (iterator->path().filename().wstring().starts_with(kPrefix))
        {
            ++count;
        }
        iterator.increment(ec);
    }
    return count;
}

[[nodiscard]] size_t CountViewerWebSaveTemps(const std::filesystem::path& directory, std::error_code& ec) noexcept
{
    constexpr std::wstring_view kPrefix = L".rsw-save-";
    constexpr std::wstring_view kSuffix = L".tmp";
    size_t count                        = 0u;
    std::filesystem::directory_iterator iterator(directory, ec);
    const std::filesystem::directory_iterator end;
    while (! ec && iterator != end)
    {
        const std::wstring filename = iterator->path().filename().wstring();
        if (filename.starts_with(kPrefix) && filename.ends_with(kSuffix))
        {
            ++count;
        }
        iterator.increment(ec);
    }
    return count;
}

[[nodiscard]] std::span<const unsigned char> TinyJpegBytes() noexcept
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

    return kTinyJpeg;
}

[[nodiscard]] std::filesystem::path WriteTinyJpegFile(const std::filesystem::path& path)
{
    return WriteBinaryFile(path, std::as_bytes(TinyJpegBytes()));
}

// A minimal, valid 1x1 PNG (RGBA). Decoding PNG goes through the WIC codec path in
// ViewerImgRaw (CoCreateInstance(CLSID_WICImagingFactory...)), unlike JPEG which uses
// turbojpeg and never touches COM -- so this fixture exercises the WIC decode path that
// plan 015 fixes. These exact bytes were generated via GDI+ and round-trip verified through
// the WIC-backed PngBitmapDecoder (confirmed 1x1 Bgra32), so the system WIC PNG codec accepts
// them; a decode failure therefore indicates a regression, not a malformed fixture.
[[nodiscard]] std::filesystem::path WriteTinyPngFile(const std::filesystem::path& path)
{
    static constexpr auto kTinyPng = std::to_array<unsigned char>({
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xAE, 0xCE, 0x1C, 0xE9, 0x00, 0x00,
        0x00, 0x04, 0x67, 0x41, 0x4D, 0x41, 0x00, 0x00, 0xB1, 0x8F, 0x0B, 0xFC, 0x61, 0x05, 0x00, 0x00, 0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00,
        0x0E, 0xC3, 0x00, 0x00, 0x0E, 0xC3, 0x01, 0xC7, 0x6F, 0xA8, 0x64, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x18, 0x57, 0x63, 0x10, 0x32,
        0x09, 0xFB, 0x0F, 0x00, 0x02, 0x94, 0x01, 0x9C, 0x32, 0x9E, 0xE0, 0x1F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    });

    return WriteBinaryFile(path, std::as_bytes(std::span(kTinyPng)));
}

void AppendU16Le(std::vector<unsigned char>& bytes, uint16_t value)
{
    bytes.push_back(static_cast<unsigned char>(value & 0xFFu));
    bytes.push_back(static_cast<unsigned char>((value >> 8u) & 0xFFu));
}

void AppendU32Le(std::vector<unsigned char>& bytes, uint32_t value)
{
    bytes.push_back(static_cast<unsigned char>(value & 0xFFu));
    bytes.push_back(static_cast<unsigned char>((value >> 8u) & 0xFFu));
    bytes.push_back(static_cast<unsigned char>((value >> 16u) & 0xFFu));
    bytes.push_back(static_cast<unsigned char>((value >> 24u) & 0xFFu));
}

void AppendTiffIfdEntry(std::vector<unsigned char>& bytes, uint16_t tag, uint16_t type, uint32_t count, uint32_t valueOrOffset)
{
    AppendU16Le(bytes, tag);
    AppendU16Le(bytes, type);
    AppendU32Le(bytes, count);
    AppendU32Le(bytes, valueOrOffset);
}

[[nodiscard]] std::filesystem::path WriteOrientedRgbTiffFile(const std::filesystem::path& path, uint16_t orientation)
{
    constexpr uint32_t kWidth       = 2u;
    constexpr uint32_t kHeight      = 3u;
    constexpr uint16_t kEntryCount  = 11u;
    constexpr uint32_t kIfdOffset   = 8u;
    constexpr uint32_t kBitsOffset  = kIfdOffset + 2u + (static_cast<uint32_t>(kEntryCount) * 12u) + 4u;
    constexpr uint32_t kPixelOffset = kBitsOffset + 6u;
    constexpr uint32_t kPixelBytes  = kWidth * kHeight * 3u;
    constexpr uint32_t kShortType   = 3u;
    constexpr uint32_t kLongType    = 4u;

    std::vector<unsigned char> bytes;
    bytes.reserve(static_cast<size_t>(kPixelOffset + kPixelBytes));
    bytes.push_back('I');
    bytes.push_back('I');
    AppendU16Le(bytes, 42u);
    AppendU32Le(bytes, kIfdOffset);
    AppendU16Le(bytes, kEntryCount);
    AppendTiffIfdEntry(bytes, 256u, kLongType, 1u, kWidth);
    AppendTiffIfdEntry(bytes, 257u, kLongType, 1u, kHeight);
    AppendTiffIfdEntry(bytes, 258u, kShortType, 3u, kBitsOffset);
    AppendTiffIfdEntry(bytes, 259u, kShortType, 1u, 1u);
    AppendTiffIfdEntry(bytes, 262u, kShortType, 1u, 2u);
    AppendTiffIfdEntry(bytes, 273u, kLongType, 1u, kPixelOffset);
    AppendTiffIfdEntry(bytes, 274u, kShortType, 1u, orientation);
    AppendTiffIfdEntry(bytes, 277u, kShortType, 1u, 3u);
    AppendTiffIfdEntry(bytes, 278u, kLongType, 1u, kHeight);
    AppendTiffIfdEntry(bytes, 279u, kLongType, 1u, kPixelBytes);
    AppendTiffIfdEntry(bytes, 284u, kShortType, 1u, 1u);
    AppendU32Le(bytes, 0u);
    AppendU16Le(bytes, 8u);
    AppendU16Le(bytes, 8u);
    AppendU16Le(bytes, 8u);

    static constexpr std::array<unsigned char, kPixelBytes> kPixels{{
        0xFF,
        0x00,
        0x00,
        0x00,
        0xFF,
        0x00,
        0x00,
        0x00,
        0xFF,
        0xFF,
        0xFF,
        0x00,
        0x00,
        0xFF,
        0xFF,
        0xFF,
        0x00,
        0xFF,
    }};
    bytes.insert(bytes.end(), kPixels.begin(), kPixels.end());
    return WriteBinaryFile(path, std::as_bytes(std::span(bytes)));
}

[[nodiscard]] std::filesystem::path WriteTinyWicJpegFile(const std::filesystem::path& path)
{
    wil::com_ptr<IWICImagingFactory> factory;
    if (FAILED(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()))) || ! factory)
    {
        return {};
    }

    wil::com_ptr<IWICStream> stream;
    if (FAILED(factory->CreateStream(stream.put())) || ! stream || FAILED(stream->InitializeFromFilename(path.c_str(), GENERIC_WRITE)))
    {
        return {};
    }

    wil::com_ptr<IWICBitmapEncoder> encoder;
    if (FAILED(factory->CreateEncoder(GUID_ContainerFormatJpeg, nullptr, encoder.put())) || ! encoder ||
        FAILED(encoder->Initialize(stream.get(), WICBitmapEncoderNoCache)))
    {
        return {};
    }

    wil::com_ptr<IWICBitmapFrameEncode> frame;
    if (FAILED(encoder->CreateNewFrame(frame.put(), nullptr)) || ! frame || FAILED(frame->Initialize(nullptr)) || FAILED(frame->SetSize(2u, 3u)))
    {
        return {};
    }

    WICPixelFormatGUID pixelFormat = GUID_WICPixelFormat24bppBGR;
    if (FAILED(frame->SetPixelFormat(&pixelFormat)) || pixelFormat != GUID_WICPixelFormat24bppBGR)
    {
        return {};
    }

    std::array<BYTE, 18> pixels{{
        0x00,
        0x00,
        0xFF,
        0x00,
        0xFF,
        0x00,
        0xFF,
        0x00,
        0x00,
        0x00,
        0xFF,
        0xFF,
        0xFF,
        0xFF,
        0x00,
        0xFF,
        0x00,
        0xFF,
    }};
    if (FAILED(frame->WritePixels(3u, 6u, static_cast<UINT>(pixels.size()), pixels.data())) || FAILED(frame->Commit()) || FAILED(encoder->Commit()))
    {
        return {};
    }

    return path;
}

[[nodiscard]] std::filesystem::path WriteTinyJpegWithExifOrientation(const std::filesystem::path& path,
                                                                     uint16_t orientationType,
                                                                     uint32_t orientationCount,
                                                                     uint16_t orientation)
{
    std::vector<unsigned char> app1;
    app1.reserve(32u);
    app1.insert(app1.end(), {'E', 'x', 'i', 'f', 0x00, 0x00, 'I', 'I'});
    AppendU16Le(app1, 42u);
    AppendU32Le(app1, 8u);
    AppendU16Le(app1, 1u);
    AppendTiffIfdEntry(app1, 0x0112u, orientationType, orientationCount, orientation);
    AppendU32Le(app1, 0u);

    if (WriteTinyWicJpegFile(path).empty())
    {
        return {};
    }

    std::ifstream input(path, std::ios::binary | std::ios::ate);
    const std::streampos inputSize = input.tellg();
    if (! input || inputSize < static_cast<std::streampos>(2))
    {
        return {};
    }
    std::vector<unsigned char> jpeg(static_cast<size_t>(inputSize));
    input.seekg(0, std::ios::beg);
    input.read(reinterpret_cast<char*>(jpeg.data()), static_cast<std::streamsize>(jpeg.size()));
    if (! input || jpeg[0] != 0xFFu || jpeg[1] != 0xD8u)
    {
        return {};
    }

    const uint16_t segmentLength = static_cast<uint16_t>(app1.size() + 2u);
    std::vector<unsigned char> bytes;
    bytes.reserve(jpeg.size() + app1.size() + 4u);
    bytes.push_back(0xFFu);
    bytes.push_back(0xD8u);
    bytes.push_back(0xFFu);
    bytes.push_back(0xE1u);
    bytes.push_back(static_cast<unsigned char>((segmentLength >> 8u) & 0xFFu));
    bytes.push_back(static_cast<unsigned char>(segmentLength & 0xFFu));
    bytes.insert(bytes.end(), app1.begin(), app1.end());
    bytes.insert(bytes.end(), jpeg.cbegin() + 2, jpeg.cend());
    return WriteBinaryFile(path, std::as_bytes(std::span(bytes)));
}

[[nodiscard]] std::filesystem::path WriteTinyWaveFile(const std::filesystem::path& path)
{
    static constexpr auto kTinyWave = std::to_array<unsigned char>({
        'R',  'I',  'F',  'F',  0x2C, 0x00, 0x00, 0x00, 'W',  'A',  'V', 'E', 'f', 'm', 't',  ' ',  0x10, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x44, 0xAC,
        0x00, 0x00, 0x44, 0xAC, 0x00, 0x00, 0x01, 0x00, 0x08, 0x00, 'd', 'a', 't', 'a', 0x08, 0x00, 0x00, 0x00, 0x80, 0x90, 0xA0, 0xB0, 0xA0, 0x90, 0x80, 0x70,
    });

    return WriteBinaryFile(path, std::as_bytes(std::span(kTinyWave)));
}

[[nodiscard]] std::filesystem::path WriteTinyPdfFile(const std::filesystem::path& path)
{
    std::string pdf = "%PDF-1.4\n";
    std::array<size_t, 5u> offsets{};
    const auto appendObject = [&](size_t objectNumber, std::string_view body)
    {
        offsets[objectNumber] = pdf.size();
        pdf += std::format("{} 0 obj\n{}\nendobj\n", objectNumber, body);
    };

    appendObject(1u, "<< /Type /Catalog /Pages 2 0 R >>");
    appendObject(2u, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
    appendObject(3u, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>");
    appendObject(4u, "<< /Length 0 >>\nstream\n\nendstream");

    const size_t xrefOffset = pdf.size();
    pdf += "xref\n0 5\n0000000000 65535 f \n";
    for (size_t objectNumber = 1u; objectNumber < offsets.size(); ++objectNumber)
    {
        pdf += std::format("{:010} 00000 n \n", offsets[objectNumber]);
    }
    pdf += std::format("trailer\n<< /Size 5 /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n", xrefOffset);
    return WriteUtf8TextFile(path, pdf);
}

enum class LocalFileReaderFault : uint8_t
{
    None = 0u,
    SeekPositionMismatch,
    OverReportedRead,
    ReadFailure,
};

class LocalFileReader final : public IFileReader
{
public:
    using RecordReadBytesFn = void (*)(void* cookie, std::wstring_view pathKey, size_t bytesRead) noexcept;

    explicit LocalFileReader(wil::unique_handle file,
                             std::wstring pathKey,
                             void* readCookie,
                             RecordReadBytesFn recordReadBytes,
                             std::optional<uint64_t> reportedSize = std::nullopt,
                             LocalFileReaderFault fault           = LocalFileReaderFault::None) noexcept
        : _file(std::move(file)),
          _pathKey(std::move(pathKey)),
          _readCookie(readCookie),
          _recordReadBytes(recordReadBytes),
          _reportedSize(reportedSize),
          _fault(fault)
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

        if (_reportedSize.has_value())
        {
            *sizeBytes = _reportedSize.value();
            return S_OK;
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
            const uint64_t actualPosition = static_cast<uint64_t>(outPosition.QuadPart);
            *newPosition                  = _fault == LocalFileReaderFault::SeekPositionMismatch ? actualPosition + 1u : actualPosition;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! buffer && bytesToRead != 0u)
        {
            return E_POINTER;
        }

        if (_fault == LocalFileReaderFault::ReadFailure)
        {
            if (bytesRead)
            {
                *bytesRead = 0u;
            }
            return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
        }

        DWORD localBytesRead = 0u;
        if (ReadFile(_file.get(), buffer, static_cast<DWORD>(bytesToRead), &localBytesRead, nullptr) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (bytesRead)
        {
            *bytesRead = _fault == LocalFileReaderFault::OverReportedRead && bytesToRead < (std::numeric_limits<unsigned long>::max)()
                             ? bytesToRead + 1u
                             : static_cast<unsigned long>(localBytesRead);
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
    std::optional<uint64_t> _reportedSize;
    LocalFileReaderFault _fault = LocalFileReaderFault::None;
};

// Test-only IFileReader that reports the real on-disk file size via GetSize() but stops yielding bytes
// once the absolute file position reaches a cap, simulating a short read (truncation race, or normal
// behaviour on virtual/remote filesystems where Read can return fewer bytes than requested). The cap is
// enforced by absolute file POSITION (not a cumulative byte counter) so it survives the encoding BOM
// probe and the Seek(0, FILE_BEGIN) the hex loader performs before its read loop — a monotonic counter
// would be partly consumed by the BOM probe and miscalibrate the simulated end-of-data.
class ShortReadFileReader final : public IFileReader
{
public:
    ShortReadFileReader(wil::unique_handle file, uint64_t capBytes) noexcept : _file(std::move(file)), _capBytes(capBytes)
    {
    }
    ShortReadFileReader(const ShortReadFileReader&)            = delete;
    ShortReadFileReader(ShortReadFileReader&&)                 = delete;
    ShortReadFileReader& operator=(const ShortReadFileReader&) = delete;
    ShortReadFileReader& operator=(ShortReadFileReader&&)      = delete;

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

        if (bytesRead)
        {
            *bytesRead = 0u;
        }

        LARGE_INTEGER zero{};
        LARGE_INTEGER position{};
        if (SetFilePointerEx(_file.get(), zero, &position, FILE_CURRENT) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const uint64_t current   = static_cast<uint64_t>(position.QuadPart);
        const uint64_t remaining = (current >= _capBytes) ? 0u : (_capBytes - current);
        if (remaining == 0u)
        {
            return S_OK; // simulate end-of-data past the cap (read == 0, success)
        }

        unsigned long effective = bytesToRead;
        if (static_cast<uint64_t>(effective) > remaining)
        {
            effective = static_cast<unsigned long>(remaining);
        }

        DWORD localBytesRead = 0u;
        if (ReadFile(_file.get(), buffer, static_cast<DWORD>(effective), &localBytesRead, nullptr) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (bytesRead)
        {
            *bytesRead = static_cast<unsigned long>(localBytesRead);
        }
        return S_OK;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    uint64_t _capBytes = 0;
};

struct BlockingReadControl final
{
    BlockingReadControl()
        : entered(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          release(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          exited(CreateEventW(nullptr, TRUE, FALSE, nullptr))
    {
    }
    BlockingReadControl(const BlockingReadControl&)            = delete;
    BlockingReadControl(BlockingReadControl&&)                 = delete;
    BlockingReadControl& operator=(const BlockingReadControl&) = delete;
    BlockingReadControl& operator=(BlockingReadControl&&)      = delete;
    ~BlockingReadControl()                                     = default;

    [[nodiscard]] bool IsValid() const noexcept
    {
        return entered && release && exited;
    }

    wil::unique_handle entered;
    wil::unique_handle release;
    wil::unique_handle exited;
    std::atomic_ulong maxRequestedBytes{0u};
};

class BlockingFileReader final : public IFileReader
{
public:
    BlockingFileReader(wil::com_ptr<IFileReader> inner, std::shared_ptr<BlockingReadControl> control) noexcept
        : _inner(std::move(inner)),
          _control(std::move(control))
    {
    }
    BlockingFileReader(const BlockingFileReader&)            = delete;
    BlockingFileReader(BlockingFileReader&&)                 = delete;
    BlockingFileReader& operator=(const BlockingFileReader&) = delete;
    BlockingFileReader& operator=(BlockingFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid != __uuidof(IUnknown) && riid != __uuidof(IFileReader))
        {
            return E_NOINTERFACE;
        }
        *ppvObject = static_cast<IFileReader*>(this);
        AddRef();
        return S_OK;
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
        return _inner ? _inner->GetSize(sizeBytes) : E_UNEXPECTED;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        return _inner ? _inner->Seek(offset, origin, newPosition) : E_UNEXPECTED;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! _inner || ! _control)
        {
            return E_UNEXPECTED;
        }

        const bool firstRead  = ! _blocked.exchange(true, std::memory_order_acq_rel);
        const auto signalExit = wil::scope_exit([&]() noexcept
        {
            if (firstRead)
            {
                static_cast<void>(SetEvent(_control->exited.get()));
            }
        });
        if (firstRead)
        {
            unsigned long previous = _control->maxRequestedBytes.load(std::memory_order_relaxed);
            while (previous < bytesToRead && ! _control->maxRequestedBytes.compare_exchange_weak(previous, bytesToRead, std::memory_order_relaxed))
            {
            }

            static_cast<void>(SetEvent(_control->entered.get()));
            const DWORD waitResult = WaitForSingleObject(_control->release.get(), 10000u);
            if (waitResult != WAIT_OBJECT_0)
            {
                if (bytesRead)
                {
                    *bytesRead = 0u;
                }
                return waitResult == WAIT_TIMEOUT ? HRESULT_FROM_WIN32(ERROR_TIMEOUT) : HRESULT_FROM_WIN32(GetLastError());
            }
        }

        return _inner->Read(buffer, bytesToRead, bytesRead);
    }

private:
    std::atomic_ulong _refCount{1u};
    wil::com_ptr<IFileReader> _inner;
    std::shared_ptr<BlockingReadControl> _control;
    std::atomic_bool _blocked{false};
};

class BuiltinFileSystemStub final : public IFileSystem, public IInformations, public IFileSystemIO
{
public:
    using DirectoryReadCallback = HRESULT (*)(void* context, const wchar_t* path, IFilesInformation** filesInformation) noexcept;

    BuiltinFileSystemStub()                                        = default;
    BuiltinFileSystemStub(const BuiltinFileSystemStub&)            = delete;
    BuiltinFileSystemStub(BuiltinFileSystemStub&&)                 = delete;
    BuiltinFileSystemStub& operator=(const BuiltinFileSystemStub&) = delete;
    BuiltinFileSystemStub& operator=(BuiltinFileSystemStub&&)      = delete;

    void SetDirectoryReadCallback(DirectoryReadCallback callback, void* context) noexcept
    {
        _directoryReadCallback = callback;
        _directoryReadContext  = context;
    }

    void UseSyntheticFileSystemMetadata() noexcept
    {
        _useSyntheticFileSystemMetadata = true;
    }

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

    // Opt-in: subsequent CreateFileReader calls return a ShortReadFileReader that reports the real file
    // size via GetSize() but only yields up to capBytes. Off by default so every other test keeps the
    // normal full-read LocalFileReader.
    void EnableShortRead(uint64_t capBytes) noexcept
    {
        _shortReadCapBytes.store(capBytes, std::memory_order_relaxed);
        _shortReadEnabled.store(true, std::memory_order_relaxed);
    }

    void DisableShortRead() noexcept
    {
        _shortReadEnabled.store(false, std::memory_order_relaxed);
    }

    void EnableBlockingRead(std::shared_ptr<BlockingReadControl> control) noexcept
    {
        std::scoped_lock lock(_blockingReadMutex);
        _blockingReadControl = std::move(control);
    }

    void DisableBlockingRead() noexcept
    {
        std::scoped_lock lock(_blockingReadMutex);
        _blockingReadControl.reset();
    }

    void EnableReportedSize(uint64_t reportedSize) noexcept
    {
        _reportedSize.store(reportedSize, std::memory_order_relaxed);
        _reportedSizeEnabled.store(true, std::memory_order_relaxed);
    }

    void DisableReportedSize() noexcept
    {
        _reportedSizeEnabled.store(false, std::memory_order_relaxed);
    }

    void EnableLocalReaderFault(LocalFileReaderFault fault) noexcept
    {
        _localReaderFault.store(static_cast<uint32_t>(fault), std::memory_order_release);
    }

    void DisableLocalReaderFault() noexcept
    {
        _localReaderFault.store(static_cast<uint32_t>(LocalFileReaderFault::None), std::memory_order_release);
    }

    [[nodiscard]] ULONG GetReferenceCount() const noexcept
    {
        return static_cast<ULONG>(_refCount.load(std::memory_order_acquire));
    }

    void MapVirtualPath(std::wstring virtualPath, std::filesystem::path backingPath)
    {
        const std::wstring key = std::filesystem::path(virtualPath).lexically_normal().wstring();
        std::scoped_lock lock(_virtualPathMappingsMutex);
        _virtualPathMappings.insert_or_assign(key, std::move(backingPath));
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

        *metaData = _useSyntheticFileSystemMetadata ? &kSyntheticMetaData : &kMetaData;
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

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (_directoryReadCallback)
        {
            return _directoryReadCallback(_directoryReadContext, path, ppFilesInformation);
        }
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

        *jsonUtf8 = kReadOnlyFileSystemCapabilitiesJson;
        return S_OK;
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

        std::filesystem::path backingPath(path);
        {
            std::scoped_lock lock(_virtualPathMappingsMutex);
            const auto mapping = _virtualPathMappings.find(key);
            if (mapping != _virtualPathMappings.end())
            {
                backingPath = mapping->second;
            }
        }
        wil::unique_handle file(CreateFileW(
            backingPath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (_shortReadEnabled.load(std::memory_order_relaxed))
        {
            auto shortReader = std::make_unique<ShortReadFileReader>(std::move(file), _shortReadCapBytes.load(std::memory_order_relaxed));
            *reader          = shortReader.release();
            return S_OK;
        }

        std::shared_ptr<BlockingReadControl> blockingControl;
        {
            std::scoped_lock lock(_blockingReadMutex);
            blockingControl = _blockingReadControl;
        }
        if (blockingControl)
        {
            auto localReader = std::make_unique<LocalFileReader>(
                std::move(file),
                key,
                this,
                [](void* cookie, std::wstring_view pathKey, size_t bytesRead) noexcept
            {
                auto* self = static_cast<BuiltinFileSystemStub*>(cookie);
                if (self)
                {
                    self->RecordReadBytes(pathKey, bytesRead);
                }
            },
                _reportedSizeEnabled.load(std::memory_order_relaxed) ? std::optional<uint64_t>(_reportedSize.load(std::memory_order_relaxed)) : std::nullopt,
                static_cast<LocalFileReaderFault>(_localReaderFault.load(std::memory_order_acquire)));
            wil::com_ptr<IFileReader> inner;
            inner.attach(localReader.release());
            auto blockingReader = std::make_unique<BlockingFileReader>(std::move(inner), std::move(blockingControl));
            *reader             = blockingReader.release();
            return S_OK;
        }

        auto localReader = std::make_unique<LocalFileReader>(
            std::move(file),
            key,
            this,
            [](void* cookie, std::wstring_view pathKey, size_t bytesRead) noexcept
        {
            auto* self = static_cast<BuiltinFileSystemStub*>(cookie);
            if (self)
            {
                self->RecordReadBytes(pathKey, bytesRead);
            }
        },
            _reportedSizeEnabled.load(std::memory_order_relaxed) ? std::optional<uint64_t>(_reportedSize.load(std::memory_order_relaxed)) : std::nullopt,
            static_cast<LocalFileReaderFault>(_localReaderFault.load(std::memory_order_acquire)));
        *reader = localReader.release();
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
    inline static const PluginMetaData kSyntheticMetaData{
        L"builtin/synthetic-file-system",
        L"synthetic",
        L"Synthetic File System",
        L"Test synthetic file system stub",
        L"ViewerPETests",
        L"1",
    };

    std::atomic_ulong _refCount{1};
    mutable std::mutex _createFileReaderCountsMutex;
    std::unordered_map<std::wstring, size_t> _createFileReaderCounts;
    mutable std::mutex _readByteCountsMutex;
    std::unordered_map<std::wstring, size_t> _readByteCounts;
    mutable std::mutex _virtualPathMappingsMutex;
    std::unordered_map<std::wstring, std::filesystem::path> _virtualPathMappings;
    std::atomic<bool> _shortReadEnabled{false};
    std::atomic<uint64_t> _shortReadCapBytes{0};
    mutable std::mutex _blockingReadMutex;
    std::shared_ptr<BlockingReadControl> _blockingReadControl;
    std::atomic<bool> _reportedSizeEnabled{false};
    std::atomic<uint64_t> _reportedSize{0u};
    std::atomic_uint32_t _localReaderFault{static_cast<uint32_t>(LocalFileReaderFault::None)};
    DirectoryReadCallback _directoryReadCallback = nullptr;
    void* _directoryReadContext                  = nullptr;
    bool _useSyntheticFileSystemMetadata         = false;
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
    CheckDxComboHostCompactChrome(viewerWindow, kViewerPEFileComboId, L"ViewerPE", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerPEFileComboId, L"ViewerPE", pluginPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerPE window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerPE window closes cleanly", success);
    CheckEmbeddedViewerHidesStandaloneFileCombo(createFn, kViewerPEPluginId, context, kViewerPEWindowClassName, kViewerPEFileComboId, L"ViewerPE", success);
    return success;
}

[[nodiscard]] bool TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerPE blocked-read test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerPE.dll";
    const std::filesystem::path activePath = buildDir / L"RedSalamander.exe";
    const std::filesystem::path middlePath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path latestPath = buildDir / L"Common.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    const std::wstring pluginModuleName    = pluginPath.filename().wstring();
    Check(std::filesystem::exists(pluginPath) && std::filesystem::exists(activePath) && std::filesystem::exists(middlePath) &&
              std::filesystem::exists(latestPath),
          L"ViewerPE blocked-read fixtures are present",
          success);
    if (! success)
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    const DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    const DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    const auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            static_cast<void>(RemoveDllDirectory(pluginCookie));
        }
        if (buildCookie)
        {
            static_cast<void>(RemoveDllDirectory(buildCookie));
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerPE.dll loads for blocked-read validation", success);
    if (! pluginModule)
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerPE factory resolves for blocked-read validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerPEPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer, L"ViewerPE instance is created for blocked-read validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    BuiltinFileSystemStub fileSystem;
    const std::wstring oversizeText = pluginPath.wstring();
    const std::wstring activeText   = activePath.wstring();
    const std::wstring middleText   = middlePath.wstring();
    const std::wstring latestText   = latestPath.wstring();
    const wchar_t* otherFiles[]     = {oversizeText.c_str(), activeText.c_str(), middleText.c_str(), latestText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerPEWindowClassName);
    fileSystem.EnableReportedSize((256ull * 1024ull * 1024ull) + 1ull);
    context.focusedPath           = oversizeText.c_str();
    context.focusedOtherFileIndex = 0u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts an oversized-provider parse request", success);
    HWND viewerWindow = nullptr;
    WndMsg::ViewerPeDebugSnapshot oversizeSnapshot{};
    Check(PumpUntil(
              [&]() noexcept
    {
        viewerWindow     = FindNewVisibleWindowByClass(kViewerPEWindowClassName, existingWindows);
        oversizeSnapshot = {};
        return viewerWindow && SendMessageW(viewerWindow, WndMsg::kViewerPeDebugGetSnapshot, 0u, reinterpret_cast<LPARAM>(&oversizeSnapshot)) == TRUE &&
               ! oversizeSnapshot.isLoading && oversizeSnapshot.requestId > 0u && oversizeSnapshot.windowIdentity > 0u && oversizeSnapshot.bodyLength > 0u &&
               FAILED(oversizeSnapshot.parseHr);
    },
              5000ms),
          L"ViewerPE rejects the reported 256 MiB plus one byte source with a terminal error",
          success);
    Check(fileSystem.GetCreateFileReaderCount(pluginPath) == 1u, L"ViewerPE queries the oversized source exactly once", success);
    Check(fileSystem.GetReadByteCount(pluginPath) == 0u, L"ViewerPE rejects an oversized source before issuing any Read", success);
    fileSystem.DisableReportedSize();

    WndMsg::ViewerPeDebugSnapshot lastTerminalSnapshot = oversizeSnapshot;
    const auto waitForTerminalAfter                    = [&](uint64_t previousRequestId, bool expectFailure, std::wstring_view label) noexcept
    {
        WndMsg::ViewerPeDebugSnapshot snapshot{};
        const bool terminal = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            return viewerWindow && SendMessageW(viewerWindow, WndMsg::kViewerPeDebugGetSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) == TRUE &&
                   snapshot.requestId > previousRequestId && ! snapshot.isLoading && snapshot.bodyLength > 0u &&
                   (expectFailure ? FAILED(snapshot.parseHr) : SUCCEEDED(snapshot.parseHr));
        },
            8000ms);
        Check(terminal, label, success);
        if (terminal)
        {
            lastTerminalSnapshot = snapshot;
        }
        return terminal;
    };

    std::error_code fileSizeEc;
    const uint64_t activeFileSize = static_cast<uint64_t>(std::filesystem::file_size(activePath, fileSizeEc));
    Check(! fileSizeEc && activeFileSize > 1u && activeFileSize < 256ull * 1024ull * 1024ull,
          L"ViewerPE exact-reader fixture has a bounded nonempty size",
          success);
    if (! fileSizeEc && activeFileSize > 1u)
    {
        context.focusedPath           = activeText.c_str();
        context.focusedOtherFileIndex = 1u;

        uint64_t previousRequestId = lastTerminalSnapshot.requestId;
        fileSystem.ResetCreateFileReaderCounts();
        fileSystem.EnableLocalReaderFault(LocalFileReaderFault::SeekPositionMismatch);
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts the seek-position mismatch test request", success);
        static_cast<void>(waitForTerminalAfter(previousRequestId, true, L"ViewerPE rejects a provider that lies about the initial seek position"));
        Check(fileSystem.GetReadByteCount(activePath) == 0u, L"ViewerPE rejects the seek mismatch before reading bytes", success);
        fileSystem.DisableLocalReaderFault();

        previousRequestId = lastTerminalSnapshot.requestId;
        fileSystem.ResetCreateFileReaderCounts();
        fileSystem.EnableReportedSize(activeFileSize + 1u);
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts the premature-EOF test request", success);
        static_cast<void>(waitForTerminalAfter(previousRequestId, true, L"ViewerPE rejects premature EOF before the committed GetSize byte count"));
        fileSystem.DisableReportedSize();

        previousRequestId = lastTerminalSnapshot.requestId;
        fileSystem.ResetCreateFileReaderCounts();
        fileSystem.EnableReportedSize(activeFileSize - 1u);
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts the trailing-byte test request", success);
        static_cast<void>(waitForTerminalAfter(previousRequestId, true, L"ViewerPE rejects bytes beyond the committed GetSize byte count"));
        fileSystem.DisableReportedSize();

        previousRequestId = lastTerminalSnapshot.requestId;
        fileSystem.ResetCreateFileReaderCounts();
        fileSystem.EnableLocalReaderFault(LocalFileReaderFault::OverReportedRead);
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts the impossible over-return test request", success);
        static_cast<void>(waitForTerminalAfter(previousRequestId, true, L"ViewerPE rejects Read counts larger than the requested buffer"));
        fileSystem.DisableLocalReaderFault();
    }

    auto firstControl = std::make_shared<BlockingReadControl>();
    Check(firstControl->IsValid(), L"ViewerPE first blocking-read events are created", success);
    fileSystem.EnableBlockingRead(firstControl);
    const auto releaseFirst = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(firstControl->release.get())); });

    context.focusedPath           = activeText.c_str();
    context.focusedOtherFileIndex = 1u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE starts the first blocked parse", success);
    Check(PumpUntil(
              [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerPEWindowClassName, existingWindows);
        return viewerWindow != nullptr && WaitForSingleObject(firstControl->entered.get(), 0u) == WAIT_OBJECT_0;
    },
              5000ms),
          L"ViewerPE first parse reaches the blocking Read while its window stays responsive",
          success);

    context.focusedPath           = middleText.c_str();
    context.focusedOtherFileIndex = 2u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts a replaceable middle parse request", success);
    context.focusedPath           = latestText.c_str();
    context.focusedOtherFileIndex = 3u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE accepts the latest parse request", success);
    Check(fileSystem.GetCreateFileReaderCount(middlePath) == 0u, L"ViewerPE does not start the replaceable middle request while one Read is active", success);

    static_cast<void>(SetEvent(firstControl->release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(firstControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
          L"ViewerPE first blocking Read exits after release",
          success);
    fileSystem.DisableBlockingRead();
    Check(PumpUntil([&]() noexcept { return fileSystem.GetReadByteCount(latestPath) > 0u; }, 5000ms),
          L"ViewerPE runs the latest pending request after the blocked request exits",
          success);
    Check(fileSystem.GetCreateFileReaderCount(middlePath) == 0u, L"ViewerPE permanently drops the superseded middle request", success);
    Check(firstControl->maxRequestedBytes.load(std::memory_order_relaxed) <= 1u * 1024u * 1024u,
          L"ViewerPE caps each file-system Read request at one MiB",
          success);

    static_cast<void>(waitForTerminalAfter(lastTerminalSnapshot.requestId, false, L"ViewerPE latest-wins request reaches a successful terminal parse"));

    for (const auto [fault, label] : std::array{
             std::pair{WndMsg::ViewerPeDebugAsyncFault::ResultAllocation, std::wstring_view(L"result-allocation")},
             std::pair{WndMsg::ViewerPeDebugAsyncFault::PayloadPost, std::wstring_view(L"payload-post")},
         })
    {
        const uint64_t previousRequestId = lastTerminalSnapshot.requestId;
        Check(SendMessageW(viewerWindow, WndMsg::kViewerPeDebugReloadWithAsyncFault, static_cast<WPARAM>(fault), 0u) == TRUE,
              std::format(L"ViewerPE arms the {} terminal-delivery fault", label),
              success);
        static_cast<void>(
            waitForTerminalAfter(previousRequestId, true, std::format(L"ViewerPE {} failure leaves loading through the scheduler fallback", label)));
        Check(viewerWindow && IsWindow(viewerWindow) != FALSE, std::format(L"ViewerPE {} terminal fallback preserves the current window", label), success);
    }

    auto closeControl = std::make_shared<BlockingReadControl>();
    Check(closeControl->IsValid(), L"ViewerPE close-path blocking-read events are created", success);
    fileSystem.EnableBlockingRead(closeControl);
    const auto releaseClose = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(closeControl->release.get())); });

    context.focusedPath           = activeText.c_str();
    context.focusedOtherFileIndex = 1u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerPE starts the close-path blocked parse", success);
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(closeControl->entered.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
          L"ViewerPE close-path parse reaches the blocking Read",
          success);

    const auto closeStarted = std::chrono::steady_clock::now();
    const HRESULT closeHr   = viewer->Close();
    const auto closeElapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - closeStarted);
    Check(SUCCEEDED(closeHr), L"ViewerPE Close succeeds while Read remains blocked", success);
    Check(closeElapsed < 500ms, L"ViewerPE Close returns within 500 ms instead of joining the blocked Read", success);
    Check(! viewerWindow || IsWindow(viewerWindow) == FALSE, L"ViewerPE window is destroyed before the blocked Read is released", success);
    Check(closeControl->maxRequestedBytes.load(std::memory_order_relaxed) <= 1u * 1024u * 1024u,
          L"ViewerPE close-path Read request remains capped at one MiB",
          success);

    viewer.reset();
    pluginModule.reset();
    Check(GetModuleHandleW(pluginModuleName.c_str()) != nullptr,
          L"ViewerPE worker keeps its DLL module loaded after the window and caller references are released",
          success);

    static_cast<void>(SetEvent(closeControl->release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(closeControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
          L"ViewerPE cancelled worker exits after the provider releases Read",
          success);
    fileSystem.DisableBlockingRead();
    Check(PumpUntil([&]() noexcept { return GetModuleHandleW(pluginModuleName.c_str()) == nullptr; }, 5000ms),
          L"ViewerPE worker releases its module pin after cancelled work exits",
          success);
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerweb_combo", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerWeb fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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
    CheckDxComboHostCompactChrome(viewerWindow, kViewerWebFileComboId, L"ViewerWeb", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerWebFileComboId, L"ViewerWeb", firstPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerWeb window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerWeb window closes cleanly", success);
    CheckEmbeddedViewerHidesStandaloneFileCombo(createFn, kViewerWebPluginId, context, kViewerWebWindowClassName, kViewerWebFileComboId, L"ViewerWeb", success);
    return success;
}

// Regression guard for Plan 018 (script injection): ViewerWeb embeds JSON / JSONL / Markdown file
// content into INLINE <script> string literals. The HTML tokenizer ends a <script> element on the
// byte sequence "</script" regardless of JS quoting, so the shared escaper MUST neutralize '<' (and
// the JS line/paragraph separators U+2028 / U+2029) or an untrusted .json/.jsonl/.md file can break
// out of the script and inject arbitrary HTML/JS. This drives the exact escaper that the generated
// documents use, so it is immune to the WebView2/modal-render observability wall and fails on the
// pre-fix escaper (which left '<' and U+2028/U+2029 unescaped).
[[nodiscard]] bool TestViewerWebEscapesScriptBreakoutInGeneratedDocuments() noexcept
{
    using ViewerWebDetail::EscapeJavaScriptString;
    using ViewerWebDetail::EscapeJavaScriptStringUtf8;

    bool success = true;

    // The core breakout payload an untrusted data file could carry.
    const std::string payload = "</script><!-- --><img src=x onerror=alert(1)><script>window.__pwn=1</script>";
    const std::string escaped = EscapeJavaScriptStringUtf8(payload);
    Check(escaped.find('<') == std::string::npos, L"UTF-8 escaper emits no literal '<'", success);
    Check(escaped.find("</script") == std::string::npos, L"UTF-8 escaper output cannot terminate <script>", success);
    Check(escaped.find("<!--") == std::string::npos, L"UTF-8 escaper neutralizes the <!-- comment opener", success);
    Check(escaped.find("\\x3C") != std::string::npos, L"UTF-8 escaper encodes '<' as the inert \\x3C escape", success);

    // U+2028 (E2 80 A8) and U+2029 (E2 80 A9) also terminate a JS string literal.
    const std::string separators        = std::string("a") + "\xE2\x80\xA8" + "b" + "\xE2\x80\xA9" + "c";
    const std::string separatorsEscaped = EscapeJavaScriptStringUtf8(separators);
    Check(separatorsEscaped.find("\xE2\x80\xA8") == std::string::npos, L"UTF-8 escaper neutralizes raw U+2028", success);
    Check(separatorsEscaped.find("\xE2\x80\xA9") == std::string::npos, L"UTF-8 escaper neutralizes raw U+2029", success);
    Check(separatorsEscaped.find("\\u2028") != std::string::npos && separatorsEscaped.find("\\u2029") != std::string::npos,
          L"UTF-8 escaper encodes U+2028/U+2029 as \\u escapes",
          success);

    // A lone 0xE2 that is not part of a separator must pass through untouched (no false positive).
    const std::string loneLead = std::string("x") + "\xE2" + "y";
    Check(EscapeJavaScriptStringUtf8(loneLead) == loneLead, L"UTF-8 escaper leaves a non-separator 0xE2 byte intact", success);

    // Legitimate content still round-trips: quotes, backslashes and newlines are escaped, nothing lost.
    const std::string normal        = "{\"name\":\"a\\b\",\n\"v\":1}";
    const std::string normalEscaped = EscapeJavaScriptStringUtf8(normal);
    Check(normalEscaped.find("\\\"") != std::string::npos, L"UTF-8 escaper escapes double quotes", success);
    Check(normalEscaped.find("\\\\") != std::string::npos, L"UTF-8 escaper escapes backslashes", success);
    Check(normalEscaped.find("\\n") != std::string::npos && normalEscaped.find('\n') == std::string::npos,
          L"UTF-8 escaper escapes newlines with no raw newline left",
          success);

    // The wide variant (find-query path) carries the same '<' and U+2028/U+2029 hardening.
    std::wstring wideInput = L"</script>";
    wideInput.push_back(static_cast<wchar_t>(0x2028)); // U+2028 line separator, built numerically
    wideInput.push_back(L'x');
    const std::wstring wideEscaped = EscapeJavaScriptString(wideInput);
    Check(wideEscaped.find(L'<') == std::wstring::npos, L"Wide escaper emits no literal '<'", success);
    Check(wideEscaped.find(L"\\x3C") != std::wstring::npos, L"Wide escaper encodes '<' as the inert \\x3C escape", success);
    Check(wideEscaped.find(static_cast<wchar_t>(0x2028)) == std::wstring::npos, L"Wide escaper neutralizes raw U+2028", success);

    return success;
}

[[nodiscard]] bool TestViewerWebSecurityPolicyAndBounds() noexcept
{
    using ViewerWebSecurity::DocumentRoute;
    using ViewerWebSecurity::EvaluateNavigation;
    using ViewerWebSecurity::NavigationAction;
    using ViewerWebSecurity::NavigationSurface;
    using ViewerWebSecurity::NormalizeTextResult;

    bool success = true;

    const std::wstring_view rawHeaders(ViewerWebSecurity::kRawHtmlResponseHeaders);
    Check(rawHeaders.find(L"sandbox") != std::wstring_view::npos, L"ViewerWeb raw response is CSP sandboxed", success);
    Check(rawHeaders.find(L"default-src 'none'") != std::wstring_view::npos, L"ViewerWeb raw response denies resources by default", success);
    Check(rawHeaders.find(L"script-src 'none'") != std::wstring_view::npos, L"ViewerWeb raw response denies document script", success);
    Check(rawHeaders.find(L"connect-src 'none'") != std::wstring_view::npos && rawHeaders.find(L"frame-src 'none'") != std::wstring_view::npos &&
              rawHeaders.find(L"object-src 'none'") != std::wstring_view::npos && rawHeaders.find(L"base-uri 'none'") != std::wstring_view::npos &&
              rawHeaders.find(L"form-action 'none'") != std::wstring_view::npos,
          L"ViewerWeb raw response denies network, frame, object, base, and form capabilities",
          success);
    Check(rawHeaders.find(L"script-src 'unsafe-inline'") == std::wstring_view::npos, L"ViewerWeb raw response never enables inline script", success);
    Check(rawHeaders.find(L"charset=") == std::wstring_view::npos, L"ViewerWeb raw response leaves encoding selection to BOM/meta sniffing", success);

    const std::wstring_view generatedHeaders(ViewerWebSecurity::kGeneratedDocumentResponseHeaders);
    Check(generatedHeaders.find(L"script-src 'unsafe-inline'") != std::wstring_view::npos,
          L"ViewerWeb generated response retains its bundled inline renderer",
          success);
    Check(generatedHeaders.find(L"connect-src 'none'") != std::wstring_view::npos && generatedHeaders.find(L"frame-src 'none'") != std::wstring_view::npos,
          L"ViewerWeb generated response still denies network and frames",
          success);
    Check(! ViewerWebSecurity::kDefaultAllowExternalNavigation, L"ViewerWeb external navigation defaults to blocked", success);
    Check(ViewerWebSecurity::kMaximumDocumentMiB == 64u, L"ViewerWeb configuration has a finite 64 MiB hard ceiling", success);

    constexpr std::wstring_view allowedDocument = L"https://viewer.redsalamander.invalid/web/42.html";
    Check(EvaluateNavigation(allowedDocument, NavigationSurface::TopLevel, false, false, allowedDocument) == NavigationAction::AllowInViewer,
          L"ViewerWeb permits only its exact active top-level document",
          success);
    Check(EvaluateNavigation(L"https://viewer.redsalamander.invalid/web/42.html#section", NavigationSurface::TopLevel, true, false, allowedDocument) ==
              NavigationAction::AllowInViewer,
          L"ViewerWeb permits a fragment on the exact active document",
          success);
    Check(EvaluateNavigation(L"https://viewer.redsalamander.invalid/web/42.html?escape=1", NavigationSurface::TopLevel, true, true, allowedDocument) ==
              NavigationAction::Block,
          L"ViewerWeb blocks query changes on the active document URL",
          success);
    Check(EvaluateNavigation(allowedDocument, NavigationSurface::Frame, true, true, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb blocks frame navigation even to the active document URL",
          success);
    Check(EvaluateNavigation(L"file:///C:/secret.txt", NavigationSurface::TopLevel, true, true, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb blocks non-allowlisted file navigation",
          success);
    Check(EvaluateNavigation(L"data:text/html,pwn", NavigationSurface::NewWindow, true, true, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb blocks non-http new windows",
          success);
    Check(EvaluateNavigation(L"https://example.invalid/", NavigationSurface::TopLevel, true, false, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb blocks external navigation under the default setting",
          success);
    Check(EvaluateNavigation(L"https://example.invalid/", NavigationSurface::TopLevel, false, true, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb blocks non-user-initiated external navigation",
          success);
    Check(EvaluateNavigation(L"https://example.invalid/", NavigationSurface::TopLevel, true, true, allowedDocument) == NavigationAction::OpenExternal,
          L"ViewerWeb permits an explicit user-initiated https system-browser open",
          success);
    Check(EvaluateNavigation(L"http://example.invalid/", NavigationSurface::Frame, true, true, allowedDocument) == NavigationAction::Block,
          L"ViewerWeb never opens frame navigation externally",
          success);

    uint64_t nextBytes = 0u;
    Check(ViewerWebSecurity::TryAccumulateWithinLimit(3u, 5u, 8u, nextBytes) && nextBytes == 8u,
          L"ViewerWeb running-byte guard accepts an exact-limit read",
          success);
    Check(! ViewerWebSecurity::TryAccumulateWithinLimit(8u, 1u, 8u, nextBytes), L"ViewerWeb running-byte guard rejects the first byte over limit", success);
    Check(! ViewerWebSecurity::TryAccumulateWithinLimit((std::numeric_limits<uint64_t>::max)() - 1u, 2u, (std::numeric_limits<uint64_t>::max)(), nextBytes),
          L"ViewerWeb running-byte guard rejects integer overflow",
          success);

    constexpr uint64_t kOneMiB = 1024ull * 1024ull;
    const uint64_t outputLimit = ViewerWebSecurity::GeneratedOutputLimit(kOneMiB);
    Check(outputLimit == 3ull * kOneMiB, L"ViewerWeb generated output limit has deterministic fixed and expansion allowances", success);
    Check(ViewerWebSecurity::IsGeneratedOutputWithinLimit(outputLimit, kOneMiB), L"ViewerWeb accepts generated output at the exact limit", success);
    Check(! ViewerWebSecurity::IsGeneratedOutputWithinLimit(outputLimit + 1u, kOneMiB),
          L"ViewerWeb rejects generated output before publishing the first byte over limit",
          success);

    const std::string utf16Le = std::string("\xFF\xFE", 2u) + std::string("A\0\xAC\x20", 4u);
    std::string normalized;
    Check(ViewerWebSecurity::NormalizeTextUtf8Bounded(utf16Le, 4u, normalized) == NormalizeTextResult::Ok && normalized == "A\xE2\x82\xAC",
          L"ViewerWeb converts UTF-16LE incrementally within the output bound",
          success);
    Check(ViewerWebSecurity::NormalizeTextUtf8Bounded(utf16Le, 3u, normalized) == NormalizeTextResult::TooLarge,
          L"ViewerWeb rejects UTF-16 expansion before output exceeds its bound",
          success);
    const std::string oddUtf16("\xFF\xFE\x41", 3u);
    Check(ViewerWebSecurity::NormalizeTextUtf8Bounded(oddUtf16, 16u, normalized) == NormalizeTextResult::InvalidEncoding,
          L"ViewerWeb rejects a truncated UTF-16 code unit",
          success);

    std::string boundaryUtf16("\xFF\xFE", 2u);
    boundaryUtf16.reserve(2u + 4097u * 2u);
    for (size_t index = 0u; index < 4095u; ++index)
    {
        boundaryUtf16.append("A\0", 2u);
    }
    boundaryUtf16.append("=\xD8\0\xDE", 4u); // U+1F600 split across the converter's nominal chunk boundary.
    Check(ViewerWebSecurity::NormalizeTextUtf8Bounded(boundaryUtf16, 4099u, normalized) == NormalizeTextResult::Ok && normalized.size() == 4099u,
          L"ViewerWeb chunked conversion preserves a surrogate pair at a chunk boundary",
          success);

    struct CleanupFaultState
    {
        bool deleteSucceeds   = false;
        bool scheduleSucceeds = false;
        size_t deleteCalls    = 0u;
        size_t scheduleCalls  = 0u;
    } cleanupState;
    const ViewerWebSecurity::StagedCleanupTracker::Operations cleanupOperations{
        .context = &cleanupState,
        .deleteNow =
            [](void* context, [[maybe_unused]] std::wstring_view path) noexcept
    {
        auto* state = static_cast<CleanupFaultState*>(context);
        ++state->deleteCalls;
        return state->deleteSucceeds;
    },
        .scheduleLater =
            [](void* context, [[maybe_unused]] std::wstring_view path) noexcept
    {
        auto* state = static_cast<CleanupFaultState*>(context);
        ++state->scheduleCalls;
        return state->scheduleSucceeds;
    },
    };
    ViewerWebSecurity::StagedCleanupTracker cleanupTracker;
    cleanupTracker.Track(L"C:\\temp\\viewerweb-staged.pdf");
    cleanupTracker.Retry(cleanupOperations);
    Check(cleanupTracker.PendingCount() == 1u && cleanupState.deleteCalls == 1u && cleanupState.scheduleCalls == 1u,
          L"ViewerWeb retains a staged path when immediate and delayed cleanup both fail",
          success);
    cleanupState.deleteSucceeds = true;
    cleanupTracker.Retry(cleanupOperations);
    Check(
        cleanupTracker.PendingCount() == 0u && cleanupState.deleteCalls == 2u, L"ViewerWeb later quiet-point retry removes the retained staged path", success);

    ViewerWebSecurity::DebugSnapshot snapshot{};
    Check(snapshot.route == DocumentRoute::None, L"ViewerWeb debug snapshot starts with no document route", success);
    return success;
}

[[nodiscard]] bool TestViewerWebVirtualHtmlUsesPrivateOriginAndEnforcesByteCaps() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1u> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerWeb security harness resolves its executable path", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerWeb.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerWeb.dll is present for the security harness", success);
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
    Check(pluginModule.is_valid(), L"ViewerWeb.dll loads for the security harness", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerWeb security harness resolves the factory", success);
    if (! createFn)
    {
        return false;
    }

    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const auto createConfiguredViewer = [&](const wchar_t* pluginId) -> wil::com_ptr<IViewer>
    {
        wil::com_ptr<IViewer> viewer;
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, pluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerWeb security harness creates a viewer", success);
        if (FAILED(createHr) || ! viewer)
        {
            return {};
        }

        wil::com_ptr<IInformations> information;
        const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), information.put_void());
        Check(SUCCEEDED(infoHr) && information != nullptr, L"ViewerWeb security harness obtains configuration", success);
        if (FAILED(infoHr) || ! information)
        {
            return {};
        }

        const HRESULT configHr = information->SetConfiguration(R"json({"maxDocumentMiB":1,"allowExternalNavigation":false,"devToolsEnabled":false})json");
        Check(SUCCEEDED(configHr), L"ViewerWeb security harness applies a 1 MiB deny-by-default configuration", success);
        return viewer;
    };

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerweb_security", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerWeb security harness TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    constexpr std::string_view hostileHtml = "<!doctype html><script>window.__viewerWebPwned=1</script><iframe src='file:///C:/Windows/win.ini'></iframe>"
                                             "<img src='https://example.invalid/leak'><form action='https://example.invalid/post'></form>"
                                             "<a target='_blank' href='data:text/html,pwn'>escape</a>";
    std::vector<std::byte> hostileUtf16Bytes;
    hostileUtf16Bytes.reserve(2u + hostileHtml.size() * 2u);
    hostileUtf16Bytes.push_back(std::byte{0xFF});
    hostileUtf16Bytes.push_back(std::byte{0xFE});
    for (const char value : hostileHtml)
    {
        hostileUtf16Bytes.push_back(static_cast<std::byte>(static_cast<unsigned char>(value)));
        hostileUtf16Bytes.push_back(std::byte{0x00});
    }
    const std::filesystem::path hostileBacking = WriteBinaryFile(tempDir / L"hostile-utf16-bom.bin", hostileUtf16Bytes);
    const std::wstring hostileVirtualPath      = L"vault\\hostile.html";

    BuiltinFileSystemStub acceptedFileSystem;
    acceptedFileSystem.MapVirtualPath(hostileVirtualPath, hostileBacking);
    wil::com_ptr<IViewer> acceptedViewer = createConfiguredViewer(kViewerWebPluginId);
    if (! acceptedViewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> acceptedInformation;
    static_cast<void>(acceptedViewer->QueryInterface(__uuidof(IInformations), acceptedInformation.put_void()));
    const char* schema = nullptr;
    Check(acceptedInformation && SUCCEEDED(acceptedInformation->GetConfigurationSchema(&schema)) && schema != nullptr,
          L"ViewerWeb security harness obtains the Web schema",
          success);
    if (schema)
    {
        const std::string_view schemaText(schema);
        Check(schemaText.find(R"("max": 64)") != std::string_view::npos, L"ViewerWeb schema publishes the 64 MiB hard ceiling", success);
        Check(schemaText.find(R"("default": "0")") != std::string_view::npos, L"ViewerWeb schema publishes blocked external navigation", success);
    }

    const std::vector<HWND> acceptedExistingWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);
    const wchar_t* acceptedFiles[]                  = {hostileVirtualPath.c_str()};
    ViewerOpenContext acceptedContext{};
    acceptedContext.fileSystem            = static_cast<IFileSystem*>(&acceptedFileSystem);
    acceptedContext.fileSystemName        = L"Virtual File System";
    acceptedContext.focusedPath           = hostileVirtualPath.c_str();
    acceptedContext.otherFiles            = acceptedFiles;
    acceptedContext.otherFileCount        = 1u;
    acceptedContext.focusedOtherFileIndex = 0u;
    Check(SUCCEEDED(acceptedViewer->Open(&acceptedContext)), L"ViewerWeb opens hostile HTML through a virtual filesystem", success);

    HWND acceptedWindow = nullptr;
    Check(PumpUntil(
              [&]() noexcept
    {
        acceptedWindow = FindNewVisibleWindowByClass(kViewerWebWindowClassName, acceptedExistingWindows);
        return acceptedWindow != nullptr;
    },
              8000ms),
          L"ViewerWeb hostile virtual HTML window becomes visible",
          success);

    ViewerWebSecurity::DebugSnapshot acceptedSnapshot{};
    const bool observedPrivateOrigin = acceptedWindow && PumpUntil(
                                                             [&]() noexcept
    {
        acceptedSnapshot = {};
        const LRESULT snapshotResult =
            SendMessageW(acceptedWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&acceptedSnapshot));
        return snapshotResult == TRUE && acceptedSnapshot.route == ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin &&
               ViewerWebSecurity::StartsWithNoCase(acceptedSnapshot.webViewSourceUrl.data(), ViewerWebSecurity::kInternalDocumentOrigin);
    },
                                                             12000ms);
    Check(observedPrivateOrigin, L"ViewerWeb WebView2 source observes the private viewer origin", success);
    Check(acceptedSnapshot.scriptsEnabled == FALSE, L"ViewerWeb disables the actual WebView2 script setting for raw HTML", success);
    Check(acceptedSnapshot.privateOrigin != FALSE, L"ViewerWeb classifies hostile HTML as private-origin content", success);
    Check(acceptedSnapshot.stagedFileTracked == FALSE, L"ViewerWeb raw HTML does not create a staging file", success);
    Check(acceptedSnapshot.loadedSourceBytes == hostileUtf16Bytes.size(), L"ViewerWeb reports accepted UTF-16 BOM raw source bytes", success);
    Check(ViewerWebSecurity::StartsWithNoCase(acceptedSnapshot.allowedDocumentUrl.data(), ViewerWebSecurity::kInternalDocumentOrigin),
          L"ViewerWeb allowlists only the generated private document URL",
          success);
    Check(! ViewerWebSecurity::StartsWithNoCase(acceptedSnapshot.webViewSourceUrl.data(), L"file:"),
          L"ViewerWeb never navigates virtual raw HTML through file://",
          success);
    Check(acceptedFileSystem.GetReadByteCount(hostileVirtualPath) == hostileUtf16Bytes.size(),
          L"ViewerWeb reads the accepted UTF-16 BOM virtual HTML exactly once within the cap",
          success);
    if (acceptedWindow)
    {
        static_cast<void>(acceptedViewer->Close());
        Check(PumpUntil([&]() noexcept { return IsWindow(acceptedWindow) == FALSE; }, 5000ms), L"ViewerWeb private-origin harness closes cleanly", success);
    }

    const std::filesystem::path pdfBacking = WriteTinyPdfFile(tempDir / L"viewerweb-local.pdf");
    const auto runPdfCase                  = [&](bool virtualFileSystem) -> std::optional<std::filesystem::path>
    {
        BuiltinFileSystemStub fileSystem;
        const std::wstring focusedPath = virtualFileSystem ? L"vault\\viewerweb.pdf" : pdfBacking.wstring();
        if (virtualFileSystem)
        {
            fileSystem.MapVirtualPath(focusedPath, pdfBacking);
        }

        wil::com_ptr<IViewer> viewer = createConfiguredViewer(kViewerWebPluginId);
        if (! viewer)
        {
            return std::nullopt;
        }

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);
        const wchar_t* files[]                  = {focusedPath.c_str()};
        ViewerOpenContext context{};
        context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName        = virtualFileSystem ? L"Virtual File System" : L"File System";
        context.focusedPath           = focusedPath.c_str();
        context.otherFiles            = files;
        context.otherFileCount        = 1u;
        context.focusedOtherFileIndex = 0u;
        Check(SUCCEEDED(viewer->Open(&context)), virtualFileSystem ? L"ViewerWeb opens a virtual PDF" : L"ViewerWeb opens a local PDF", success);

        HWND window = nullptr;
        Check(PumpUntil(
                  [&]() noexcept
        {
            window = FindNewVisibleWindowByClass(kViewerWebWindowClassName, existingWindows);
            return window != nullptr;
        },
                  8000ms),
              virtualFileSystem ? L"ViewerWeb virtual PDF window becomes visible" : L"ViewerWeb local PDF window becomes visible",
              success);

        ViewerWebSecurity::DebugSnapshot snapshot{};
        const bool completed = window && PumpUntil(
                                             [&]() noexcept
        {
            snapshot = {};
            static_cast<void>(SendMessageW(window, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&snapshot)));
            return snapshot.route == ViewerWebSecurity::DocumentRoute::StagedPdf && snapshot.navigationCompleted != FALSE;
        },
                                             12000ms);
        Check(completed, virtualFileSystem ? L"ViewerWeb virtual PDF navigation completes" : L"ViewerWeb local PDF navigation completes", success);
        Check(snapshot.navigationSucceeded != FALSE,
              virtualFileSystem ? L"ViewerWeb virtual PDF navigation succeeds" : L"ViewerWeb local PDF navigation succeeds",
              success);
        Check(snapshot.scriptsEnabled == FALSE,
              virtualFileSystem ? L"ViewerWeb virtual PDF keeps document script disabled" : L"ViewerWeb local PDF keeps document script disabled",
              success);
        Check(ViewerWebSecurity::StartsWithNoCase(snapshot.webViewSourceUrl.data(), L"file:"),
              virtualFileSystem ? L"ViewerWeb virtual PDF uses its constrained staged file URL" : L"ViewerWeb local PDF uses its constrained file URL",
              success);
        Check(snapshot.stagedFileTracked != FALSE,
              virtualFileSystem ? L"ViewerWeb tracks the virtual PDF staging path" : L"ViewerWeb also stages local PDFs through the provider trust boundary",
              success);
        Check(std::wstring_view(snapshot.allowedDocumentUrl.data()).find(L".pdf") != std::wstring_view::npos,
              L"ViewerWeb PDF allowlist always names a collision-safe .pdf path",
              success);

        std::optional<std::filesystem::path> stagedPath;
        {
            std::array<wchar_t, MAX_PATH + 1u> decodedPath{};
            DWORD decodedLength = static_cast<DWORD>(decodedPath.size());
            if (SUCCEEDED(PathCreateFromUrlW(snapshot.allowedDocumentUrl.data(), decodedPath.data(), &decodedLength, 0u)))
            {
                stagedPath = std::filesystem::path(decodedPath.data());
            }
            Check(stagedPath.has_value() && std::filesystem::exists(stagedPath.value()),
                  virtualFileSystem ? L"ViewerWeb virtual PDF staging artifact exists while active"
                                    : L"ViewerWeb local PDF staging artifact exists while active",
                  success);
        }

        if (window)
        {
            static_cast<void>(viewer->Close());
            static_cast<void>(PumpUntil([&]() noexcept { return IsWindow(window) == FALSE; }, 5000ms));
        }
        return stagedPath;
    };

    const std::optional<std::filesystem::path> virtualStagedPdfPath = runPdfCase(true);
    const std::optional<std::filesystem::path> localStagedPdfPath   = runPdfCase(false);
    for (const auto& stagedPdfPath : std::array{virtualStagedPdfPath, localStagedPdfPath})
    {
        if (! stagedPdfPath.has_value())
        {
            continue;
        }
        Check(PumpUntil(
                  [&]() noexcept
        {
            std::error_code existsEc;
            return ! std::filesystem::exists(stagedPdfPath.value(), existsEc);
        },
                  3000ms),
              L"ViewerWeb deletes every staged PDF after the WebView2 quiet point",
              success);
    }

    const auto runRejectedCase = [&](std::wstring_view caseName,
                                     const wchar_t* pluginId,
                                     std::wstring virtualPath,
                                     const std::filesystem::path& backingPath,
                                     std::optional<uint64_t> reportedSize,
                                     size_t minimumReadBytes,
                                     bool expectZeroReads,
                                     bool expectOutputRejection) noexcept
    {
        BuiltinFileSystemStub fileSystem;
        fileSystem.MapVirtualPath(virtualPath, backingPath);
        if (reportedSize.has_value())
        {
            fileSystem.EnableReportedSize(reportedSize.value());
        }

        wil::com_ptr<IViewer> viewer = createConfiguredViewer(pluginId);
        if (! viewer)
        {
            return;
        }

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);
        const wchar_t* files[]                  = {virtualPath.c_str()};
        ViewerOpenContext context{};
        context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName        = L"Virtual File System";
        context.focusedPath           = virtualPath.c_str();
        context.otherFiles            = files;
        context.otherFileCount        = 1u;
        context.focusedOtherFileIndex = 0u;
        Check(SUCCEEDED(viewer->Open(&context)), std::format(L"ViewerWeb {} case starts", caseName), success);

        HWND window = nullptr;
        static_cast<void>(PumpUntil(
            [&]() noexcept
        {
            window = FindNewVisibleWindowByClass(kViewerWebWindowClassName, existingWindows);
            return window != nullptr;
        },
            8000ms));
        ViewerWebSecurity::DebugSnapshot snapshot{};
        const bool completed = PumpUntil(
            [&]() noexcept
        {
            const size_t readBytes = fileSystem.GetReadByteCount(virtualPath);
            const bool readBoundaryReached =
                fileSystem.GetCreateFileReaderCount(virtualPath) >= 1u && (expectZeroReads ? readBytes == 0u : readBytes >= minimumReadBytes);
            if (! readBoundaryReached || ! expectOutputRejection)
            {
                return readBoundaryReached;
            }
            snapshot = {};
            return window && SendMessageW(window, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&snapshot)) == TRUE &&
                   snapshot.generatedOutputRejected != FALSE;
        },
            8000ms);
        Check(completed, std::format(L"ViewerWeb {} case reaches its byte guard", caseName), success);
        const size_t readBytes = fileSystem.GetReadByteCount(virtualPath);
        Check(expectZeroReads ? readBytes == 0u : readBytes == minimumReadBytes,
              std::format(L"ViewerWeb {} case stops at the expected read boundary", caseName),
              success);

        if (window)
        {
            static_cast<void>(SendMessageW(window, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&snapshot)));
            Check(snapshot.route == ViewerWebSecurity::DocumentRoute::None,
                  std::format(L"ViewerWeb {} case never publishes a document route", caseName),
                  success);
            if (expectOutputRejection)
            {
                Check(snapshot.generatedOutputRejected != FALSE && snapshot.generatedOutputBytes > snapshot.generatedOutputLimit,
                      L"ViewerWeb generated-output case rejects the expanded page before publication",
                      success);
            }
            static_cast<void>(viewer->Close());
            static_cast<void>(PumpUntil([&]() noexcept { return IsWindow(window) == FALSE; }, 5000ms));
        }
    };

    runRejectedCase(L"advertised-size rejection", kViewerWebPluginId, L"vault\\advertised.html", hostileBacking, 2u * 1024u * 1024u, 0u, true, false);

    constexpr size_t kOneMiB = 1024u * 1024u;
    std::vector<std::byte> oversizedBytes(kOneMiB + 1u, std::byte{0x41});
    const std::filesystem::path oversizedBacking = WriteBinaryFile(tempDir / L"running-overflow.bin", oversizedBytes);
    runRejectedCase(L"trailing-size mismatch", kViewerWebPluginId, L"vault\\running.html", oversizedBacking, 1u, 2u, false, false);

    std::vector<std::byte> expansionBytes(kOneMiB, std::byte{0x3C});
    const std::filesystem::path expansionBacking = WriteBinaryFile(tempDir / L"generated-expansion.md", expansionBytes);
    runRejectedCase(L"generated-output rejection", kViewerMarkdownPluginId, L"vault\\expansion.md", expansionBacking, std::nullopt, kOneMiB, false, true);

    std::error_code cleanupEc;
    static_cast<void>(std::filesystem::remove(hostileBacking, cleanupEc));
    static_cast<void>(std::filesystem::remove(oversizedBacking, cleanupEc));
    static_cast<void>(std::filesystem::remove(expansionBacking, cleanupEc));
    static_cast<void>(std::filesystem::remove(pdfBacking, cleanupEc));
    return success;
}

[[nodiscard]] bool TestViewerWebTransactionalSaveAsAndCloseSafety() noexcept
{
#ifndef _DEBUG
    std::wcout << L"[SKIP] ViewerWeb transactional Save As validation requires ENABLE_TESTS hooks.\n";
    return true;
#else
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1u> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerWeb Save As test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerWeb.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    const std::wstring pluginModuleName    = pluginPath.filename().wstring();
    Check(std::filesystem::exists(pluginPath), L"ViewerWeb.dll is present for Save As data-safety validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    const DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    const DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    const auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            static_cast<void>(RemoveDllDirectory(pluginCookie));
        }
        if (buildCookie)
        {
            static_cast<void>(RemoveDllDirectory(buildCookie));
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerWeb.dll loads for Save As data-safety validation", success);
    if (! pluginModule)
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    const FARPROC shutdownProc     = GetProcAddress(pluginModule.get(), "RedSalamanderPluginShutdown");
    const FARPROC canUnloadProc    = GetProcAddress(pluginModule.get(), "RedSalamanderPluginCanUnloadNow");
    RedSalamanderCreateFn createFn = nullptr;
    using PluginShutdownFn         = void(__stdcall*)();
    using PluginCanUnloadFn        = BOOL(__stdcall*)();
    PluginShutdownFn shutdownFn    = nullptr;
    PluginCanUnloadFn canUnloadFn  = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    static_assert(sizeof(shutdownFn) == sizeof(shutdownProc));
    static_assert(sizeof(canUnloadFn) == sizeof(canUnloadProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    std::memcpy(&shutdownFn, &shutdownProc, sizeof(shutdownFn));
    std::memcpy(&canUnloadFn, &canUnloadProc, sizeof(canUnloadFn));
    Check(createFn && shutdownFn && canUnloadFn, L"ViewerWeb factory and unload-gate exports are available for Save As data-safety validation", success);
    if (! createFn || ! shutdownFn || ! canUnloadFn)
    {
        return false;
    }
    bool moduleShutdownPending         = true;
    const auto shutdownModuleOnFailure = wil::scope_exit([&]() noexcept
    {
        if (moduleShutdownPending)
        {
            shutdownFn();
        }
    });

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerweb_save_as_safety", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerWeb Save As fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    constexpr std::string_view kSourceText       = "<!doctype html><html><body>ViewerWeb transactional Save As source</body></html>\r\n";
    constexpr std::string_view kDestinationText  = "pre-existing ViewerWeb destination bytes must survive\r\n";
    const std::filesystem::path sourcePath       = WriteUtf8TextFile(tempDir / L"source.html", kSourceText);
    const std::filesystem::path destinationPath  = WriteUtf8TextFile(tempDir / L"destination.html", kDestinationText);
    const std::filesystem::path postFallbackPath = tempDir / L"post-fallback.html";
    const auto sourceBytes                       = ReadBinaryFile(sourcePath);
    const auto destinationBytes                  = ReadBinaryFile(destinationPath);
    Check(sourceBytes.has_value() && destinationBytes.has_value(), L"ViewerWeb Save As baseline bytes are readable", success);
    if (! sourceBytes.has_value() || ! destinationBytes.has_value())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupError;
        static_cast<void>(std::filesystem::remove(sourcePath, cleanupError));
        cleanupError.clear();
        static_cast<void>(std::filesystem::remove(destinationPath, cleanupError));
        cleanupError.clear();
        static_cast<void>(std::filesystem::remove(postFallbackPath, cleanupError));
    });

    BuiltinFileSystemStub fileSystem;
    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerWebPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer, L"ViewerWeb factory creates the Save As test viewer", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    HWND viewerWindow                       = nullptr;
    const auto emergencyClose               = wil::scope_exit([&]() noexcept
    {
        if (viewerWindow && IsWindow(viewerWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(viewerWindow, 5000ms));
        }
        viewer.reset();
    });
    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);
    const std::wstring sourcePathText       = sourcePath.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.focusedPath    = sourcePathText.c_str();
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerWeb opens the Save As source", success);
    Check(PumpUntil(
              [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerWebWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
              8000ms),
          L"ViewerWeb Save As test window becomes visible",
          success);
    if (! viewerWindow)
    {
        return false;
    }

    ViewerWebSecurity::DebugSnapshot loadSnapshot{};
    Check(PumpUntil(
              [&]() noexcept
    {
        loadSnapshot = {};
        return SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&loadSnapshot)) == TRUE &&
               loadSnapshot.loadedSourceBytes == sourceBytes.value().size() && loadSnapshot.route == ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin;
    },
              12000ms),
          L"ViewerWeb fully consumes the source before Save As fault injection",
          success);

    const UINT controlMessage = ViewerWebSecurity::GetDebugControlMessage();
    Check(controlMessage != 0u, L"ViewerWeb exposes the registered Save As debug-control message", success);

    const auto checkBytes = [&](const std::filesystem::path& path, const std::vector<std::byte>& expected, std::wstring_view description) noexcept
    {
        const auto actual = ReadBinaryFile(path);
        Check(actual.has_value() && actual.value() == expected, description, success);
    };
    const auto checkNoTemps = [&]() noexcept
    {
        std::error_code countError;
        const size_t count = CountViewerWebSaveTemps(tempDir, countError);
        Check(! countError && count == 0u, L"ViewerWeb Save As leaves no sibling transaction temp behind", success);
    };
    const auto resetDestination = [&]() { static_cast<void>(WriteUtf8TextFile(destinationPath, kDestinationText)); };
    const auto invokeSave       = [&](const std::filesystem::path& destination, uint32_t faultMask = 0u) noexcept -> bool
    {
        const std::wstring destinationText = destination.wstring();
        ViewerWebSecurity::DebugSaveAsRequest request{};
        request.destinationPath = destinationText.c_str();
        request.faultMask       = faultMask;
        const LRESULT handled   = SendMessageW(
            viewerWindow, controlMessage, static_cast<WPARAM>(ViewerWebSecurity::DebugControlAction::SaveAsToPath), reinterpret_cast<LPARAM>(&request));
        Check(handled == TRUE && SUCCEEDED(request.submissionHr), L"ViewerWeb submits Save As work asynchronously", success);
        if (handled != TRUE || FAILED(request.submissionHr))
        {
            return false;
        }

        ViewerWebSecurity::DebugSnapshot activeSnapshot{};
        const bool active = SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&activeSnapshot)) == TRUE &&
                            activeSnapshot.saveInProgress != FALSE;
        Check(active, L"ViewerWeb reports Save As in progress before dispatching its posted completion", success);
        const bool completed = PumpUntil(
            [&]() noexcept
        {
            ViewerWebSecurity::DebugSnapshot snapshot{};
            return IsWindow(viewerWindow) != FALSE &&
                   SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&snapshot)) == TRUE &&
                   snapshot.saveInProgress == FALSE;
        },
            10000ms);
        Check(completed, L"ViewerWeb Save As always reaches a terminal UI state", success);
        return completed;
    };

    Check(invokeSave(sourcePath), L"ViewerWeb same-path Save As completes", success);
    checkBytes(sourcePath, sourceBytes.value(), L"ViewerWeb same-path Save As preserves every source byte");
    checkNoTemps();

    resetDestination();
    Check(invokeSave(destinationPath), L"ViewerWeb transactional Save As replaces an existing destination", success);
    checkBytes(destinationPath, sourceBytes.value(), L"ViewerWeb successful Save As commits the exact provider bytes");
    checkNoTemps();

    struct InjectedFailure
    {
        std::wstring_view name;
        uint32_t faultMask;
    };
    constexpr std::array injectedFailures{
        InjectedFailure{L"write", ViewerWebSecurity::DebugSaveFaultWrite},
        InjectedFailure{L"flush", ViewerWebSecurity::DebugSaveFaultFlush},
        InjectedFailure{L"commit", ViewerWebSecurity::DebugSaveFaultCommit},
    };
    for (const InjectedFailure& failure : injectedFailures)
    {
        resetDestination();
        Check(invokeSave(destinationPath, failure.faultMask), std::format(L"ViewerWeb injected {} failure reaches a terminal UI state", failure.name), success);
        checkBytes(
            destinationPath, destinationBytes.value(), std::format(L"ViewerWeb injected {} failure preserves pre-existing destination bytes", failure.name));
        checkNoTemps();
    }

    resetDestination();
    wil::unique_handle lockedDestination(
        CreateFileW(destinationPath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    Check(lockedDestination.is_valid(), L"ViewerWeb locks the pre-existing destination against atomic replacement", success);
    Check(invokeSave(destinationPath), L"ViewerWeb locked-destination failure reaches a terminal UI state", success);
    lockedDestination.reset();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb locked destination preserves every pre-existing byte");
    checkNoTemps();

    resetDestination();
    fileSystem.EnableLocalReaderFault(LocalFileReaderFault::SeekPositionMismatch);
    Check(invokeSave(destinationPath), L"ViewerWeb mismatched provider seek position reaches a terminal UI state", success);
    fileSystem.DisableLocalReaderFault();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb rejects a mismatched provider seek without replacing the destination");
    checkNoTemps();

    resetDestination();
    fileSystem.EnableLocalReaderFault(LocalFileReaderFault::ReadFailure);
    Check(invokeSave(destinationPath), L"ViewerWeb provider Read failure reaches a terminal UI state", success);
    fileSystem.DisableLocalReaderFault();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb failed provider Read preserves pre-existing destination bytes");
    checkNoTemps();

    resetDestination();
    fileSystem.EnableReportedSize(static_cast<uint64_t>(sourceBytes.value().size()) + 1u);
    Check(invokeSave(destinationPath), L"ViewerWeb premature provider EOF reaches a terminal UI state", success);
    fileSystem.DisableReportedSize();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb rejects premature provider EOF without replacing the destination");
    checkNoTemps();

    resetDestination();
    fileSystem.EnableReportedSize(static_cast<uint64_t>(sourceBytes.value().size() - 1u));
    Check(invokeSave(destinationPath), L"ViewerWeb trailing provider data reaches a terminal UI state", success);
    fileSystem.DisableReportedSize();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb rejects trailing provider data without replacing the destination");
    checkNoTemps();

    resetDestination();
    fileSystem.EnableLocalReaderFault(LocalFileReaderFault::OverReportedRead);
    Check(invokeSave(destinationPath), L"ViewerWeb over-reported provider read reaches a terminal UI state", success);
    fileSystem.DisableLocalReaderFault();
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb rejects an impossible provider read count without replacing the destination");
    checkNoTemps();

    static_cast<void>(WriteUtf8TextFile(postFallbackPath, kDestinationText));
    ViewerWebSecurity::DebugSnapshot beforePostFailure{};
    static_cast<void>(SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&beforePostFailure)));
    Check(SendMessageW(viewerWindow, controlMessage, static_cast<WPARAM>(ViewerWebSecurity::DebugControlAction::FailNextAsyncSaveCompletionPost), 0u) == TRUE,
          L"ViewerWeb arms the Save As completion-post failure seam",
          success);
    Check(invokeSave(postFallbackPath), L"ViewerWeb Save As completion-post failure reaches the allocation-free UI fallback", success);
    ViewerWebSecurity::DebugSnapshot afterPostFailure{};
    static_cast<void>(SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&afterPostFailure)));
    Check(afterPostFailure.asyncSavePostFailures == beforePostFailure.asyncSavePostFailures + 1u,
          L"ViewerWeb records exactly one forced Save As completion-post failure",
          success);
    checkBytes(postFallbackPath, sourceBytes.value(), L"ViewerWeb commits exact bytes even when completion posting fails");
    checkNoTemps();

    Check(SUCCEEDED(viewer->Close()), L"ViewerWeb closes before callback-release epoch validation", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerWeb destroys its first window before module quiet-point validation",
          success);
    viewer.reset();
    shutdownFn();
    size_t unloadProbeCount = 0u;
    Check(PumpUntil(
              [&]() noexcept
    {
        ++unloadProbeCount;
        return canUnloadFn() == TRUE;
    },
              10000ms) &&
              unloadProbeCount >= 2u,
          L"ViewerWeb callback release epoch requires a later quiescent unload observation",
          success);

    viewerWindow             = nullptr;
    const HRESULT recreateHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerWebPluginId, viewer.put_void());
    Check(SUCCEEDED(recreateHr) && viewer, L"ViewerWeb recreates a viewer after the first module quiet point", success);
    const std::vector<HWND> existingBlockedWindows = CollectVisibleWindowsByClass(kViewerWebWindowClassName);
    Check(viewer && SUCCEEDED(viewer->Open(&context)), L"ViewerWeb reopens the source for blocked-provider pin validation", success);
    Check(PumpUntil(
              [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerWebWindowClassName, existingBlockedWindows);
        ViewerWebSecurity::DebugSnapshot snapshot{};
        return viewerWindow && SendMessageW(viewerWindow, ViewerWebSecurity::GetDebugSnapshotMessage(), 0u, reinterpret_cast<LPARAM>(&snapshot)) == TRUE &&
               snapshot.loadedSourceBytes == sourceBytes.value().size() && snapshot.route == ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin;
    },
              12000ms),
          L"ViewerWeb recreated window reaches a loaded source before blocked Save As",
          success);

    resetDestination();
    const auto blockingControl = std::make_shared<BlockingReadControl>();
    Check(blockingControl && blockingControl->IsValid(), L"ViewerWeb blocked Save As control initializes", success);
    fileSystem.EnableBlockingRead(blockingControl);
    const auto releaseBlockedRead = wil::scope_exit([&]() noexcept
    {
        if (blockingControl && blockingControl->release)
        {
            static_cast<void>(SetEvent(blockingControl->release.get()));
        }
        fileSystem.DisableBlockingRead();
    });

    const std::wstring blockedDestinationText = destinationPath.wstring();
    ViewerWebSecurity::DebugSaveAsRequest blockedRequest{};
    blockedRequest.destinationPath = blockedDestinationText.c_str();
    const LRESULT blockedHandled   = SendMessageW(
        viewerWindow, controlMessage, static_cast<WPARAM>(ViewerWebSecurity::DebugControlAction::SaveAsToPath), reinterpret_cast<LPARAM>(&blockedRequest));
    Check(blockedHandled == TRUE && SUCCEEDED(blockedRequest.submissionHr), L"ViewerWeb submits the blocked Save As request", success);
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(blockingControl->entered.get(), 0u) == WAIT_OBJECT_0; }, 8000ms),
          L"ViewerWeb Save As worker blocks inside the provider read",
          success);
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb blocked Save As leaves the destination untouched before close");

    const auto closeStarted = std::chrono::steady_clock::now();
    const HRESULT closeHr   = viewer->Close();
    const auto closeElapsed = std::chrono::steady_clock::now() - closeStarted;
    Check(SUCCEEDED(closeHr), L"ViewerWeb closes while Save As remains blocked in the provider", success);
    Check(closeElapsed < 500ms, L"ViewerWeb close never waits for a blocked Save As provider", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerWeb destroys its window while the Save As provider remains blocked",
          success);
    viewer.reset();
    shutdownFn();
    moduleShutdownPending = false;
    Check(canUnloadFn() == FALSE, L"ViewerWeb unload gate rejects refresh while Save As remains blocked", success);
    pluginModule.reset();
    Check(GetModuleHandleW(pluginModuleName.c_str()) != nullptr,
          L"ViewerWeb Save As worker keeps its DLL mapped after the caller releases the module handle",
          success);
    static_cast<void>(SetEvent(blockingControl->release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(blockingControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
          L"ViewerWeb blocked Save As provider exits after release",
          success);
    fileSystem.DisableBlockingRead();
    Check(PumpUntil(
              [&]() noexcept
    {
        std::error_code countError;
        return fileSystem.GetReferenceCount() == 1u && CountViewerWebSaveTemps(tempDir, countError) == 0u && ! countError;
    },
              5000ms),
          L"ViewerWeb cancelled Save As releases provider state and removes its sibling temp",
          success);
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerWeb close during blocked Save As never commits partial bytes");
    checkNoTemps();
    Check(PumpUntil([&]() noexcept { return GetModuleHandleW(pluginModuleName.c_str()) == nullptr; }, 8000ms),
          L"ViewerWeb module unloads only after the blocked Save As callback returns",
          success);
    return success;
#endif
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_combo", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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
    CheckImageRawComboHostIsInsetInsideHeader(viewerWindow, success);
    CheckDxComboHostClickActivation(viewerWindow, kViewerImgRawFileComboId, L"ViewerImgRaw", success);
    CheckDxComboHostCompactChrome(viewerWindow, kViewerImgRawFileComboId, L"ViewerImgRaw", success);
    CheckDxComboHostAccessibility(viewerWindow, kViewerImgRawFileComboId, L"ViewerImgRaw", firstPath.filename().wstring(), success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerImgRaw window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerImgRaw window closes cleanly", success);
    CheckEmbeddedViewerHidesStandaloneFileCombo(
        createFn, kViewerImgRawPluginId, context, kViewerImgRawWindowClassName, kViewerImgRawFileComboId, L"ViewerImgRaw", success);
    return success;
}

// Test-only host stub that records the decode-outcome alerts ViewerImgRaw raises. ViewerImgRaw
// obtains its alert sink purely via host->QueryInterface(IHostAlerts) (ViewerImgRaw.cpp SetHost),
// so passing this stub as the factory 'host' argument makes the async decode outcome observable:
// a successful decode calls ClearAlert, while a decode failure (e.g. the CO_E_NOTINITIALIZED that
// plan 015 fixes) calls ShowAlert(HOST_ALERT_WARNING). Stack-allocated, so Release never deletes.
class AlertRecordingHostStub final : public IHost, public IHostAlerts
{
public:
    AlertRecordingHostStub()                                         = default;
    AlertRecordingHostStub(const AlertRecordingHostStub&)            = delete;
    AlertRecordingHostStub(AlertRecordingHostStub&&)                 = delete;
    AlertRecordingHostStub& operator=(const AlertRecordingHostStub&) = delete;
    AlertRecordingHostStub& operator=(AlertRecordingHostStub&&)      = delete;

    [[nodiscard]] unsigned int WarningAlertCount() const noexcept
    {
        return _warningAlerts.load(std::memory_order_acquire);
    }
    [[nodiscard]] unsigned int ClearAlertCount() const noexcept
    {
        return _clearAlerts.load(std::memory_order_acquire);
    }
    [[nodiscard]] bool DecodeCompleted() const noexcept
    {
        return WarningAlertCount() != 0u || ClearAlertCount() != 0u;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IHost))
        {
            *ppvObject = static_cast<IHost*>(this);
        }
        else if (riid == __uuidof(IHostAlerts))
        {
            *ppvObject = static_cast<IHostAlerts*>(this);
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

    HRESULT STDMETHODCALLTYPE ShowAlert(const HostAlertRequest* request, void* /*cookie*/) noexcept override
    {
        if (request != nullptr && request->severity == HOST_ALERT_WARNING)
        {
            _warningAlerts.fetch_add(1u, std::memory_order_acq_rel);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ClearAlert(HostAlertScope /*scope*/, void* /*cookie*/) noexcept override
    {
        _clearAlerts.fetch_add(1u, std::memory_order_acq_rel);
        return S_OK;
    }

private:
    std::atomic<ULONG> _refCount{1u};
    std::atomic<unsigned int> _warningAlerts{0u};
    std::atomic<unsigned int> _clearAlerts{0u};
};

[[nodiscard]] bool TestViewerImgRawMenuOwnershipOrientationAndExifScalarGuards() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerImgRaw Farsight test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for Farsight validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads for Farsight validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerImgRaw factory export is available for Farsight validation", success);
    if (! createFn)
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_farsight", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw Farsight TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    const std::filesystem::path fallbackMenuImage   = WriteTinyPngFile(tempDir / L"native-menu-fallback.png");
    const std::filesystem::path orientedTiff        = WriteOrientedRgbTiffFile(tempDir / L"oriented-wic.tiff", 6u);
    const std::filesystem::path orientedDngFallback = WriteOrientedRgbTiffFile(tempDir / L"oriented-wic-fallback.dng", 6u);
    const std::filesystem::path validExifJpeg       = WriteTinyJpegWithExifOrientation(tempDir / L"valid-orientation.jpg", 3u, 1u, 6u);
    const std::filesystem::path wrongTypeExifJpeg   = WriteTinyJpegWithExifOrientation(tempDir / L"wrong-type-orientation.jpg", 4u, 1u, 6u);
    const std::filesystem::path hugeCountExifJpeg   = WriteTinyJpegWithExifOrientation(tempDir / L"huge-count-orientation.jpg", 3u, 0x80000001u, 6u);

    auto cleanupTemp = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(fallbackMenuImage, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(orientedTiff, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(orientedDngFallback, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(validExifJpeg, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(wrongTypeExifJpeg, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(hugeCountExifJpeg, cleanupEc));
    });

    const bool fixturesReady = std::filesystem::exists(fallbackMenuImage) && std::filesystem::exists(orientedTiff) &&
                               std::filesystem::exists(orientedDngFallback) && std::filesystem::exists(validExifJpeg) &&
                               std::filesystem::exists(wrongTypeExifJpeg) && std::filesystem::exists(hugeCountExifJpeg);
    Check(fixturesReady, L"ViewerImgRaw Farsight image fixtures are valid files", success);
    if (! fixturesReady)
    {
        return false;
    }

    struct ProbeResult final
    {
        bool observed = false;
        WndMsg::ViewerImgRawDecodeDebugSnapshot decode{};
    };

    constexpr wchar_t kForceMenuAttachFailureEnvVar[] = L"REDSALAMANDER_VIEWERIMGRAW_FORCE_MENU_ATTACH_FAILURE";
    const auto runProbe = [&](const std::filesystem::path& path, std::wstring_view label, bool forceNativeMenuFallback) noexcept -> ProbeResult
    {
        ProbeResult probe{};
        AlertRecordingHostStub hostStub;
        wil::com_ptr<IViewer> viewer;
        const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&hostStub), kViewerImgRawPluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, std::format(L"{} creates a ViewerImgRaw instance", label), success);
        if (FAILED(createHr) || ! viewer)
        {
            return probe;
        }

        const std::wstring previousFaultValue = GetEnvironmentString(kForceMenuAttachFailureEnvVar);
        const BOOL setFaultOk                 = SetEnvironmentVariableW(kForceMenuAttachFailureEnvVar, forceNativeMenuFallback ? L"1" : nullptr);
        Check(setFaultOk != FALSE, std::format(L"{} configures the menu attach fault seam", label), success);
        auto restoreFault = wil::scope_exit([&]() noexcept
        { static_cast<void>(SetEnvironmentVariableW(kForceMenuAttachFailureEnvVar, previousFaultValue.empty() ? nullptr : previousFaultValue.c_str())); });

        BuiltinFileSystemStub fileSystem;
        const std::wstring pathText = path.wstring();
        const wchar_t* otherFiles[] = {pathText.c_str()};
        ViewerOpenContext context{};
        context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName        = L"File System";
        context.focusedPath           = pathText.c_str();
        context.otherFiles            = otherFiles;
        context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
        context.focusedOtherFileIndex = 0u;

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);
        const HRESULT openHr                    = viewer->Open(&context);
        Check(SUCCEEDED(openHr), std::format(L"{} opens", label), success);
        if (FAILED(openHr))
        {
            return probe;
        }

        HWND viewerWindow = nullptr;
        const bool opened = PumpUntil(
            [&]() noexcept
        {
            viewerWindow = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingWindows);
            return viewerWindow != nullptr;
        },
            8000ms);
        Check(opened, std::format(L"{} shows a viewer window", label), success);
        if (! opened || ! viewerWindow)
        {
            static_cast<void>(viewer->Close());
            return probe;
        }

        WndMsg::ViewerNativeMenuModelDebugSnapshot menuSnapshot{};
        const bool menuSnapshotRead =
            SendMessageW(viewerWindow, WndMsg::kViewerDebugGetNativeMenuModelSnapshot, 0, reinterpret_cast<LPARAM>(&menuSnapshot)) != FALSE;
        Check(menuSnapshotRead, std::format(L"{} exposes the menu ownership snapshot", label), success);
        if (forceNativeMenuFallback)
        {
            Check(GetMenu(viewerWindow) != nullptr, std::format(L"{} keeps the native window-owned menu", label), success);
            Check(CountVisibleChildWindowsByClass(viewerWindow, kDxNativeMenuBarWindowClassName) == 0u,
                  std::format(L"{} does not leave a partial DxUi menu host", label),
                  success);
            Check(! menuSnapshot.hasHiddenMenuModel, std::format(L"{} does not adopt the window-owned fallback menu", label), success);
        }
        else
        {
            Check(GetMenu(viewerWindow) == nullptr, std::format(L"{} detaches the native menu after DxUi attach", label), success);
            Check(menuSnapshot.hasHiddenMenuModel, std::format(L"{} has one hidden RAII menu owner", label), success);
        }

        const bool decodeCompleted = PumpUntil([&]() noexcept { return hostStub.DecodeCompleted(); }, 8000ms);
        Check(decodeCompleted, std::format(L"{} completes decode", label), success);
        Check(hostStub.WarningAlertCount() == 0u, std::format(L"{} decodes without a warning", label), success);

        probe.observed = SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0, reinterpret_cast<LPARAM>(&probe.decode)) != FALSE;
        Check(probe.observed, std::format(L"{} exposes the decode snapshot", label), success);
        Check(probe.decode.hasImage, std::format(L"{} retains a decoded image", label), success);

        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), std::format(L"{} closes", label), success);
        Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), std::format(L"{} destroys its viewer window", label), success);
        viewer.reset();
        return probe;
    };

    for (int iteration = 0; iteration < 3; ++iteration)
    {
        const ProbeResult fallback = runProbe(fallbackMenuImage, std::format(L"ViewerImgRaw native-menu fallback iteration {}", iteration + 1), true);
        Check(fallback.observed, L"ViewerImgRaw native-menu fallback completes without an ownership crash", success);
    }

    const ProbeResult wic = runProbe(orientedTiff, L"ViewerImgRaw oriented WIC TIFF", false);
    Check(wic.decode.baseOrientation == 6u && wic.decode.viewOrientation == 6u,
          std::format(L"WIC TIFF orientation 6 propagates to render geometry (base={}, view={})", wic.decode.baseOrientation, wic.decode.viewOrientation),
          success);
    Check(wic.decode.sourceWidth == 2u && wic.decode.sourceHeight == 3u && wic.decode.orientedWidth == 3u && wic.decode.orientedHeight == 2u,
          std::format(L"WIC TIFF orientation swaps displayed axes (source={}x{}, oriented={}x{})",
                      wic.decode.sourceWidth,
                      wic.decode.sourceHeight,
                      wic.decode.orientedWidth,
                      wic.decode.orientedHeight),
          success);

    const ProbeResult wicFallback = runProbe(orientedDngFallback, L"ViewerImgRaw oriented WIC fallback", false);
    Check(wicFallback.decode.viewOrientation == 6u && wicFallback.decode.orientedWidth == 3u && wicFallback.decode.orientedHeight == 2u,
          std::format(L"WIC fallback preserves orientation (view={}, oriented={}x{})",
                      wicFallback.decode.viewOrientation,
                      wicFallback.decode.orientedWidth,
                      wicFallback.decode.orientedHeight),
          success);

    const ProbeResult validExif = runProbe(validExifJpeg, L"ViewerImgRaw valid EXIF orientation", false);
    Check(validExif.decode.baseOrientation == 6u, L"Scalar SHORT count=1 EXIF orientation is accepted", success);

    const ProbeResult wrongTypeExif = runProbe(wrongTypeExifJpeg, L"ViewerImgRaw wrong-type EXIF orientation", false);
    Check(wrongTypeExif.decode.baseOrientation == 1u, L"Wrong-type EXIF orientation is rejected", success);

    const ProbeResult hugeCountExif = runProbe(hugeCountExifJpeg, L"ViewerImgRaw huge-count EXIF orientation", false);
    Check(hugeCountExif.decode.baseOrientation == 1u, L"Overflowing/non-scalar EXIF orientation count is rejected", success);

    return success;
}

[[nodiscard]] bool TestViewerImgRawResourcePolicyHelpers() noexcept
{
    bool success = true;
    using ViewerImgRawResource::DecodedImageLayout;
    using ViewerImgRawResource::DecodedImagePolicy;
    using ViewerImgRawResource::ValidationError;

    DecodedImageLayout layout{};
    Check(ViewerImgRawResource::ValidateDecodedImage(8000u, 6000u, ViewerImgRawResource::kProductionDecodedImagePolicy, layout) == ValidationError::None &&
              layout.pixels == 48'000'000u && layout.bgraBytes == 192'000'000u,
          L"ViewerImgRaw production policy accepts a practical 48-megapixel WIC/JPEG/RAW image",
          success);
    Check(ViewerImgRawResource::ValidateDecodedImage(10'000u, 8'000u, ViewerImgRawResource::kProductionDecodedImagePolicy, layout) ==
              ValidationError::PixelLimit,
          L"ViewerImgRaw production policy rejects excessive decoded pixels before BGRA allocation",
          success);
    Check(ViewerImgRawResource::ValidateDecodedImage(16'385u, 1u, ViewerImgRawResource::kProductionDecodedImagePolicy, layout) ==
              ValidationError::DimensionLimit,
          L"ViewerImgRaw production policy rejects an excessive single dimension",
          success);

    DecodedImagePolicy smallPolicy{};
    smallPolicy.maxDimension         = 4u;
    smallPolicy.maxPixels            = 4u;
    smallPolicy.maxBgraBytes         = 16u;
    smallPolicy.maxEmbeddedJpegBytes = 32u;
    Check(ViewerImgRawResource::ValidateDecodedImage(2u, 2u, smallPolicy, layout) == ValidationError::None && layout.bgraBytes == 16u,
          L"ViewerImgRaw deterministic small policy accepts its exact WIC/JPEG/RAW boundary",
          success);
    Check(ViewerImgRawResource::ValidateDecodedImage(2u, 3u, smallPolicy, layout) == ValidationError::PixelLimit,
          L"ViewerImgRaw deterministic small policy rejects one pixel beyond its RAW/WIC/JPEG limit",
          success);

    Check(ViewerImgRawResource::ValidateEmbeddedJpeg(2u, 2u, 16u, 64u, smallPolicy, layout) == ValidationError::None,
          L"ViewerImgRaw embedded-JPEG guard accepts bounded declared dimensions and compressed length",
          success);
    Check(ViewerImgRawResource::ValidateEmbeddedJpeg(2u, 2u, 65u, 64u, smallPolicy, layout) == ValidationError::SourceLength,
          L"ViewerImgRaw embedded-JPEG guard rejects tlength beyond the RAW source",
          success);
    Check(ViewerImgRawResource::ValidateEmbeddedJpeg(2u, 2u, 33u, 64u, smallPolicy, layout) == ValidationError::ByteLimit,
          L"ViewerImgRaw embedded-JPEG guard enforces its fixed compressed-byte ceiling",
          success);

    uint64_t expectedPackedBytes = 0u;
    Check(ViewerImgRawResource::ValidatePackedBitmap(2u, 2u, 3u, 8u, 12u, 64u, smallPolicy, layout, expectedPackedBytes) == ValidationError::None &&
              expectedPackedBytes == 12u,
          L"ViewerImgRaw embedded-bitmap guard accepts the exact packed source size",
          success);
    Check(ViewerImgRawResource::ValidatePackedBitmap(2u, 2u, 3u, 8u, 11u, 64u, smallPolicy, layout, expectedPackedBytes) == ValidationError::SourceLength,
          L"ViewerImgRaw embedded-bitmap guard rejects truncated tlength",
          success);
    Check(ViewerImgRawResource::ValidatePackedBitmap(2u, 2u, 5u, 8u, 20u, 64u, smallPolicy, layout, expectedPackedBytes) == ValidationError::InvalidFormat,
          L"ViewerImgRaw embedded-bitmap guard rejects unsupported channel counts",
          success);

    uint64_t multiplied = 0u;
    Check(! ViewerImgRawResource::TryMultiply((std::numeric_limits<uint64_t>::max)(), 2u, multiplied),
          L"ViewerImgRaw shared byte arithmetic rejects uint64 overflow",
          success);
    return success;
}

[[nodiscard]] bool TestViewerImgRawResourceBudgetAndLongPathExport() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerImgRaw resource test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for resource validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads for resource validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerImgRaw resource validation resolves the factory", success);
    if (! createFn)
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_resource", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw resource TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    const std::filesystem::path wicOne   = WriteOrientedRgbTiffFile(tempDir / L"resource-one.tiff", 1u);
    const std::filesystem::path wicTwo   = WriteOrientedRgbTiffFile(tempDir / L"resource-two.tiff", 1u);
    const std::filesystem::path wicThree = WriteOrientedRgbTiffFile(tempDir / L"resource-three.tiff", 1u);
    const std::filesystem::path jpeg     = WriteTinyWicJpegFile(tempDir / L"resource-jpeg.jpg");
    const bool fixturesReady =
        std::filesystem::exists(wicOne) && std::filesystem::exists(wicTwo) && std::filesystem::exists(wicThree) && std::filesystem::exists(jpeg);
    Check(fixturesReady, L"ViewerImgRaw resource fixtures are valid files", success);
    if (! fixturesReady)
    {
        return false;
    }

    auto cleanupFixtures = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(wicOne, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(wicTwo, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(wicThree, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(jpeg, cleanupEc));
    });

    constexpr wchar_t kSmallPolicyEnv[]         = L"REDSALAMANDER_VIEWERIMGRAW_FORCE_SMALL_DECODE_POLICY";
    constexpr wchar_t kSmallBudgetEnv[]         = L"REDSALAMANDER_VIEWERIMGRAW_FORCE_SMALL_PREFETCH_BUDGET";
    constexpr wchar_t kPausePrefetchEnv[]       = L"REDSALAMANDER_VIEWERIMGRAW_PAUSE_PREFETCH_AFTER_RESERVE";
    constexpr wchar_t kPausePrefetchCommitEnv[] = L"REDSALAMANDER_VIEWERIMGRAW_PAUSE_PREFETCH_BEFORE_COMMIT";
    const std::wstring previousSmallPolicy      = GetEnvironmentString(kSmallPolicyEnv);
    const std::wstring previousSmallBudget      = GetEnvironmentString(kSmallBudgetEnv);
    const std::wstring previousPause            = GetEnvironmentString(kPausePrefetchEnv);
    const std::wstring previousCommitPause      = GetEnvironmentString(kPausePrefetchCommitEnv);
    auto restoreEnvironment                     = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, previousSmallPolicy.empty() ? nullptr : previousSmallPolicy.c_str()));
        static_cast<void>(SetEnvironmentVariableW(kSmallBudgetEnv, previousSmallBudget.empty() ? nullptr : previousSmallBudget.c_str()));
        static_cast<void>(SetEnvironmentVariableW(kPausePrefetchEnv, previousPause.empty() ? nullptr : previousPause.c_str()));
        static_cast<void>(SetEnvironmentVariableW(kPausePrefetchCommitEnv, previousCommitPause.empty() ? nullptr : previousCommitPause.c_str()));
    });
    static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSmallBudgetEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchCommitEnv, nullptr));

    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const auto runViewer = [&](const std::vector<std::wstring>& paths, size_t focusedIndex, std::wstring_view label, auto&& inspect) noexcept
    {
        BuiltinFileSystemStub fileSystem;
        AlertRecordingHostStub hostStub;
        wil::com_ptr<IViewer> viewer;
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&hostStub), kViewerImgRawPluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, std::format(L"{} creates a viewer", label), success);
        if (FAILED(createHr) || ! viewer || paths.empty() || focusedIndex >= paths.size())
        {
            return;
        }

        std::vector<const wchar_t*> otherFiles;
        otherFiles.reserve(paths.size());
        for (const std::wstring& path : paths)
        {
            otherFiles.push_back(path.c_str());
        }

        ViewerOpenContext context{};
        context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName        = L"File System";
        context.focusedPath           = paths[focusedIndex].c_str();
        context.otherFiles            = otherFiles.data();
        context.otherFileCount        = static_cast<unsigned long>(otherFiles.size());
        context.focusedOtherFileIndex = static_cast<unsigned long>(focusedIndex);

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);
        const HRESULT openHr                    = viewer->Open(&context);
        Check(SUCCEEDED(openHr), std::format(L"{} opens", label), success);
        if (FAILED(openHr))
        {
            return;
        }

        HWND viewerWindow = nullptr;
        const bool opened = PumpUntil(
            [&]() noexcept
        {
            viewerWindow = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingWindows);
            return viewerWindow != nullptr;
        },
            8000ms);
        Check(opened, std::format(L"{} shows a viewer window", label), success);
        if (! opened || ! viewerWindow)
        {
            static_cast<void>(viewer->Close());
            return;
        }

        const bool decodeCompleted = PumpUntil([&]() noexcept { return hostStub.DecodeCompleted(); }, 8000ms);
        Check(decodeCompleted, std::format(L"{} reaches a decode result", label), success);
        inspect(viewerWindow, hostStub);

        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), std::format(L"{} closes", label), success);
        Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), std::format(L"{} destroys its window", label), success);
    };

    static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, L"1"));
    runViewer({wicOne.wstring()},
              0u,
              L"ViewerImgRaw small-policy WIC",
              [&](HWND window, const AlertRecordingHostStub& host) noexcept
    {
        WndMsg::ViewerImgRawDecodeDebugSnapshot snapshot{};
        const bool read = SendMessageW(window, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
        Check(read && ! snapshot.hasImage, L"Small decoded-image policy rejects WIC before retaining BGRA", success);
        Check(host.WarningAlertCount() > 0u, L"Small decoded-image policy reports the WIC rejection", success);
    });
    runViewer({jpeg.wstring()},
              0u,
              L"ViewerImgRaw small-policy JPEG",
              [&](HWND window, const AlertRecordingHostStub& host) noexcept
    {
        WndMsg::ViewerImgRawDecodeDebugSnapshot snapshot{};
        const bool read = SendMessageW(window, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
        Check(read && ! snapshot.hasImage, L"Small decoded-image policy rejects libjpeg before retaining BGRA", success);
        Check(host.WarningAlertCount() > 0u, L"Small decoded-image policy reports the JPEG rejection", success);
    });
    static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, nullptr));

    static_cast<void>(SetEnvironmentVariableW(kSmallBudgetEnv, L"1"));
    runViewer({wicOne.wstring(), wicTwo.wstring(), wicThree.wstring()},
              1u,
              L"ViewerImgRaw bounded speculative cache",
              [&](HWND window, const AlertRecordingHostStub& host) noexcept
    {
        Check(host.WarningAlertCount() == 0u, L"Speculative budget rejection does not fail the main image", success);
        WndMsg::ViewerImgRawResourceDebugSnapshot snapshot{};
        const bool budgetSettled = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            if (SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) == FALSE)
            {
                return false;
            }
            return snapshot.budgetAcceptedCount >= 1u && snapshot.budgetRejectedCount >= 1u && snapshot.inflightDecodeCount == 0u;
        },
            8000ms);
        Check(budgetSettled, L"Speculative prefetch records deterministic admission and rejection", success);
        Check(snapshot.speculativeBytesLimit == 32u && snapshot.speculativeBytes <= snapshot.speculativeBytesLimit &&
                  snapshot.speculativeBytesPeak <= snapshot.speculativeBytesLimit,
              L"Speculative decoded bytes never exceed the configured small cap",
              success);

        static_cast<void>(SendMessageW(window, WndMsg::kViewerImgRawDebugClearImageCache, 0u, 0));
        const bool released = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            return SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE &&
                   snapshot.speculativeBytes == 0u && snapshot.cachedImageCount == 0u;
        },
            5000ms);
        Check(released, L"Clearing the cache releases every speculative decoded-byte reservation", success);
    });

    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchEnv, L"1"));
    runViewer({wicOne.wstring(), wicTwo.wstring()},
              0u,
              L"ViewerImgRaw stale-generation prefetch cleanup",
              [&](HWND window, const AlertRecordingHostStub&) noexcept
    {
        WndMsg::ViewerImgRawResourceDebugSnapshot snapshot{};
        const bool reservationObserved = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            return SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE &&
                   snapshot.speculativeBytes > 0u && snapshot.inflightDecodeCount > 0u;
        },
            5000ms);
        Check(reservationObserved, L"Stale-generation test observes an in-flight speculative reservation", success);
        static_cast<void>(SendMessageW(window, WndMsg::kViewerImgRawDebugClearImageCache, 0u, 0));
        const bool staleReleased = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            return SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE &&
                   snapshot.speculativeBytes == 0u;
        },
            5000ms);
        Check(staleReleased, L"Invalidating the request generation releases the in-flight reservation", success);
    });
    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchEnv, nullptr));

    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchCommitEnv, L"1"));
    runViewer({wicOne.wstring(), wicTwo.wstring()},
              0u,
              L"ViewerImgRaw stale prefetch commit rejection",
              [&](HWND window, const AlertRecordingHostStub&) noexcept
    {
        ViewerImgRawAsyncProtocol::DebugStateSnapshot debugSnapshot{};
        const bool commitPaused = PumpUntil(
            [&]() noexcept
        {
            debugSnapshot = {};
            return SendMessageW(window,
                                WndMsg::kViewerImgRawDebugGetResourceSnapshot,
                                ViewerImgRawAsyncProtocol::kDebugStateSnapshotSelector,
                                reinterpret_cast<LPARAM>(&debugSnapshot)) == TRUE &&
                   debugSnapshot.prefetchCommitPaused;
        },
            5000ms);
        Check(commitPaused, L"Stale prefetch commit test reaches the post-generation-check pause", success);

        static_cast<void>(SendMessageW(window, WndMsg::kViewerImgRawDebugClearImageCache, 0u, 0));
        static_cast<void>(SetEnvironmentVariableW(kPausePrefetchCommitEnv, nullptr));

        WndMsg::ViewerImgRawResourceDebugSnapshot resourceSnapshot{};
        const bool staleCommitRejected = PumpUntil(
            [&]() noexcept
        {
            debugSnapshot    = {};
            resourceSnapshot = {};
            return SendMessageW(window,
                                WndMsg::kViewerImgRawDebugGetResourceSnapshot,
                                ViewerImgRawAsyncProtocol::kDebugStateSnapshotSelector,
                                reinterpret_cast<LPARAM>(&debugSnapshot)) == TRUE &&
                   SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&resourceSnapshot)) == TRUE &&
                   ! debugSnapshot.prefetchCommitPaused && resourceSnapshot.cachedImageCount == 0u && resourceSnapshot.inflightDecodeCount == 0u &&
                   resourceSnapshot.speculativeBytes == 0u;
        },
            5000ms);
        Check(staleCommitRejected, L"Generation invalidation prevents a paused prefetch from repopulating the cleared cache or retaining budget", success);
    });
    static_cast<void>(SetEnvironmentVariableW(kPausePrefetchCommitEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSmallBudgetEnv, nullptr));

    const auto toExtendedPath = [](std::wstring_view path)
    {
        if (path.rfind(LR"(\\?\)", 0u) == 0u)
        {
            return std::wstring(path);
        }
        if (path.rfind(LR"(\\)", 0u) == 0u)
        {
            return std::wstring(LR"(\\?\UNC\)") + std::wstring(path.substr(2u));
        }
        return std::wstring(LR"(\\?\)") + std::wstring(path);
    };

    std::filesystem::path absoluteLongRoot = std::filesystem::absolute(tempDir / L"long-export", ec);
    Check(! ec, L"ViewerImgRaw long export root resolves absolutely", success);
    std::wstring longDirectory = absoluteLongRoot.wstring();
    for (unsigned int index = 0u; index < 8u; ++index)
    {
        longDirectory.push_back(L'\\');
        longDirectory.append(32u, static_cast<wchar_t>(L'a' + index));
    }
    const std::wstring extendedLongDirectory = toExtendedPath(longDirectory);
    ec.clear();
    static_cast<void>(std::filesystem::create_directories(std::filesystem::path(extendedLongDirectory), ec));
    Check(! ec, L"ViewerImgRaw creates a long-path export directory", success);
    const std::wstring outputPath         = longDirectory + L"\\exported.png";
    const std::wstring extendedOutputPath = toExtendedPath(outputPath);
    Check(outputPath.size() > MAX_PATH, L"ViewerImgRaw export test destination actually exceeds MAX_PATH", success);

    runViewer({wicOne.wstring()},
              0u,
              L"ViewerImgRaw long-path export",
              [&](HWND window, const AlertRecordingHostStub& host) noexcept
    {
        Check(host.WarningAlertCount() == 0u, L"Long-path export source decodes successfully", success);
        WndMsg::ViewerImgRawDebugExportRequest request{};
        request.destinationPath = outputPath.c_str();
        const bool requested    = SendMessageW(window, WndMsg::kViewerImgRawDebugExportToPath, 0u, reinterpret_cast<LPARAM>(&request)) != FALSE;
        Check(requested && request.queued, L"ViewerImgRaw queues the real long-path export", success);

        WIN32_FILE_ATTRIBUTE_DATA attributes{};
        const bool committed = PumpUntil(
            [&]() noexcept
        {
            attributes = {};
            if (GetFileAttributesExW(extendedOutputPath.c_str(), GetFileExInfoStandard, &attributes) == FALSE)
            {
                return false;
            }
            ULARGE_INTEGER size{};
            size.LowPart  = attributes.nFileSizeLow;
            size.HighPart = attributes.nFileSizeHigh;
            return size.QuadPart > 0u;
        },
            10'000ms);
        Check(committed, L"ViewerImgRaw atomically commits a non-empty image beyond MAX_PATH", success);

        WIN32_FIND_DATAW findData{};
        const std::wstring tempPattern = extendedLongDirectory + L"\\.rsi-*.tmp";
        wil::unique_hfind tempFind(FindFirstFileW(tempPattern.c_str(), &findData));
        Check(! tempFind, L"ViewerImgRaw long-path export leaves no sibling staging file", success);
    });

    ec.clear();
    static_cast<void>(std::filesystem::remove_all(std::filesystem::path(toExtendedPath(absoluteLongRoot.wstring())), ec));
    return success;
}

[[nodiscard]] bool TestViewerImgRawLatestWinsExactReaderAndCloseSafety() noexcept
{
#ifndef _DEBUG
    std::wcout << L"[SKIP] ViewerImgRaw scheduler safety requires ENABLE_TESTS hooks.\n";
    return true;
#else
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerImgRaw scheduler test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    const std::wstring pluginModuleName    = pluginPath.filename().wstring();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for scheduler safety", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    const DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    const DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    const auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            static_cast<void>(RemoveDllDirectory(pluginCookie));
        }
        if (buildCookie)
        {
            static_cast<void>(RemoveDllDirectory(buildCookie));
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads for scheduler safety", success);
    if (! pluginModule)
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    const FARPROC shutdownProc     = GetProcAddress(pluginModule.get(), "RedSalamanderPluginShutdown");
    const FARPROC canUnloadProc    = GetProcAddress(pluginModule.get(), "RedSalamanderPluginCanUnloadNow");
    RedSalamanderCreateFn createFn = nullptr;
    using PluginShutdownFn         = void(__stdcall*)();
    using PluginCanUnloadFn        = BOOL(__stdcall*)();
    PluginShutdownFn shutdownFn    = nullptr;
    PluginCanUnloadFn canUnloadFn  = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    static_assert(sizeof(shutdownFn) == sizeof(shutdownProc));
    static_assert(sizeof(canUnloadFn) == sizeof(canUnloadProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    std::memcpy(&shutdownFn, &shutdownProc, sizeof(shutdownFn));
    std::memcpy(&canUnloadFn, &canUnloadProc, sizeof(canUnloadFn));
    Check(createFn && shutdownFn && canUnloadFn, L"ViewerImgRaw exposes factory, shutdown, and runtime unload-gate exports", success);
    if (! createFn || ! shutdownFn || ! canUnloadFn)
    {
        return false;
    }
    bool moduleShutdownPending         = true;
    const auto shutdownModuleOnFailure = wil::scope_exit([&]() noexcept
    {
        if (moduleShutdownPending)
        {
            shutdownFn();
        }
    });

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_scheduler_safety", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw scheduler TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    const std::filesystem::path activePath = WriteOrientedRgbTiffFile(tempDir / L"active.tiff", 1u);
    const std::filesystem::path middlePath = WriteOrientedRgbTiffFile(tempDir / L"middle.tiff", 1u);
    const std::filesystem::path latestPath = WriteOrientedRgbTiffFile(tempDir / L"latest.tiff", 1u);
    Check(std::filesystem::exists(activePath) && std::filesystem::exists(middlePath) && std::filesystem::exists(latestPath),
          L"ViewerImgRaw scheduler fixtures are valid TIFF files",
          success);
    if (! success)
    {
        return false;
    }
    const auto cleanupFixtures = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(activePath, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(middlePath, cleanupEc));
        cleanupEc.clear();
        static_cast<void>(std::filesystem::remove(latestPath, cleanupEc));
    });

    constexpr wchar_t kAllocationFailureEnv[]  = L"REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_ALLOCATION_FAILURE";
    constexpr wchar_t kPostFailureEnv[]        = L"REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_POST_FAILURE";
    const std::wstring previousAllocationValue = GetEnvironmentString(kAllocationFailureEnv);
    const std::wstring previousPostValue       = GetEnvironmentString(kPostFailureEnv);
    const auto restoreEnvironment              = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kAllocationFailureEnv, previousAllocationValue.empty() ? nullptr : previousAllocationValue.c_str()));
        static_cast<void>(SetEnvironmentVariableW(kPostFailureEnv, previousPostValue.empty() ? nullptr : previousPostValue.c_str()));
    });
    static_cast<void>(SetEnvironmentVariableW(kAllocationFailureEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kPostFailureEnv, nullptr));

    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    {
        AlertRecordingHostStub retainedHost;
        wil::com_ptr<IViewer> retainedViewer;
        Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&retainedHost), kViewerImgRawPluginId, retainedViewer.put_void())) &&
                  retainedViewer,
              L"ViewerImgRaw creates an unopened viewer for live-instance unload gating",
              success);
        if (retainedViewer)
        {
            shutdownFn();
            Check(canUnloadFn() == FALSE, L"ViewerImgRaw runtime unload gate remains false while an unopened COM instance is retained", success);
            retainedViewer.reset();
            Check(canUnloadFn() == TRUE, L"ViewerImgRaw runtime unload gate becomes true after the last unopened COM instance is released", success);
        }
    }

    BuiltinFileSystemStub fileSystem;
    AlertRecordingHostStub host;
    wil::com_ptr<IViewer> viewer;
    Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&host), kViewerImgRawPluginId, viewer.put_void())) && viewer,
          L"ViewerImgRaw creates a viewer for scheduler safety",
          success);
    if (! viewer)
    {
        return false;
    }
    wil::com_ptr<IInformations> information;
    static_cast<void>(viewer->QueryInterface(__uuidof(IInformations), information.put_void()));
    Check(information && SUCCEEDED(information->SetConfiguration(R"json({"preferThumbnail":false,"prevCache":0,"nextCache":0})json")),
          L"ViewerImgRaw accepts deterministic no-prefetch configuration",
          success);

    const std::wstring activeText = activePath.wstring();
    const std::wstring middleText = middlePath.wstring();
    const std::wstring latestText = latestPath.wstring();
    const wchar_t* otherFiles[]   = {activeText.c_str(), middleText.c_str(), latestText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.otherFiles     = otherFiles;
    context.otherFileCount = static_cast<unsigned long>(std::size(otherFiles));

    auto firstControl = std::make_shared<BlockingReadControl>();
    Check(firstControl->IsValid(), L"ViewerImgRaw latest-wins blocking events are created", success);
    fileSystem.EnableBlockingRead(firstControl);
    const auto releaseFirst = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(firstControl->release.get())); });

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);
    context.focusedPath                     = activeText.c_str();
    context.focusedOtherFileIndex           = 0u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerImgRaw starts the first blocked decode", success);
    HWND viewerWindow = nullptr;
    Check(PumpUntil(
              [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingWindows);
        return viewerWindow && WaitForSingleObject(firstControl->entered.get(), 0u) == WAIT_OBJECT_0;
    },
              8000ms),
          L"ViewerImgRaw first decode reaches a blocked provider Read while the window remains responsive",
          success);

    WndMsg::ViewerImgRawResourceDebugSnapshot blockedSnapshot{};
    Check(viewerWindow && SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&blockedSnapshot)) == TRUE &&
              blockedSnapshot.activeMainDecodeCount == 1u && blockedSnapshot.pendingMainDecodeCount == 0u && blockedSnapshot.loading &&
              blockedSnapshot.currentRequestId != 0u,
          L"ViewerImgRaw exposes exactly one active main decode while the provider is blocked",
          success);

    ViewerImgRawAsyncProtocol::DebugStateSnapshot progressBefore{};
    Check(viewerWindow && SendMessageW(viewerWindow,
                                       WndMsg::kViewerImgRawDebugGetResourceSnapshot,
                                       ViewerImgRawAsyncProtocol::kDebugStateSnapshotSelector,
                                       reinterpret_cast<LPARAM>(&progressBefore)) == TRUE,
          L"ViewerImgRaw exposes its progress-generation debug state",
          success);
    constexpr int kCurrentProgressStage   = 1;
    constexpr int kCurrentProgressPercent = 37;
    Check(viewerWindow && PostMessageW(viewerWindow,
                                       WndMsg::kViewerImgRawAsyncProgress,
                                       static_cast<WPARAM>(blockedSnapshot.currentRequestId),
                                       ViewerImgRawAsyncProtocol::PackProgress(kCurrentProgressStage, kCurrentProgressPercent)) != FALSE,
          L"ViewerImgRaw accepts a queued current-generation progress probe",
          success);
    PumpPendingMessages();

    ViewerImgRawAsyncProtocol::DebugStateSnapshot progressCurrent{};
    Check(viewerWindow &&
              SendMessageW(viewerWindow,
                           WndMsg::kViewerImgRawDebugGetResourceSnapshot,
                           ViewerImgRawAsyncProtocol::kDebugStateSnapshotSelector,
                           reinterpret_cast<LPARAM>(&progressCurrent)) == TRUE &&
              progressCurrent.progressApplyCount == progressBefore.progressApplyCount + 1u && progressCurrent.progressStage == kCurrentProgressStage &&
              progressCurrent.progressPercent == kCurrentProgressPercent,
          L"ViewerImgRaw applies progress only when the queued request generation is current",
          success);

    Check(viewerWindow && PostMessageW(viewerWindow,
                                       WndMsg::kViewerImgRawAsyncProgress,
                                       static_cast<WPARAM>(blockedSnapshot.currentRequestId - 1u),
                                       ViewerImgRawAsyncProtocol::PackProgress(2, 88)) != FALSE,
          L"ViewerImgRaw accepts a queued stale-generation progress probe",
          success);
    PumpPendingMessages();

    ViewerImgRawAsyncProtocol::DebugStateSnapshot progressStale{};
    Check(viewerWindow &&
              SendMessageW(viewerWindow,
                           WndMsg::kViewerImgRawDebugGetResourceSnapshot,
                           ViewerImgRawAsyncProtocol::kDebugStateSnapshotSelector,
                           reinterpret_cast<LPARAM>(&progressStale)) == TRUE &&
              progressStale.progressApplyCount == progressCurrent.progressApplyCount && progressStale.progressStage == progressCurrent.progressStage &&
              progressStale.progressPercent == progressCurrent.progressPercent,
          L"ViewerImgRaw drops queued progress from a stale request generation",
          success);

    context.focusedPath           = middleText.c_str();
    context.focusedOtherFileIndex = 1u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerImgRaw accepts a replaceable middle decode request", success);
    context.focusedPath           = latestText.c_str();
    context.focusedOtherFileIndex = 2u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerImgRaw accepts the latest decode request", success);

    WndMsg::ViewerImgRawResourceDebugSnapshot pendingSnapshot{};
    Check(viewerWindow && SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&pendingSnapshot)) == TRUE &&
              pendingSnapshot.activeMainDecodeCount == 1u && pendingSnapshot.pendingMainDecodeCount == 1u && pendingSnapshot.replacedMainDecodeCount == 1u,
          L"ViewerImgRaw bounds main scheduling to one active and one replaceable pending decode",
          success);
    Check(fileSystem.GetCreateFileReaderCount(middlePath) == 0u && fileSystem.GetCreateFileReaderCount(latestPath) == 0u,
          L"ViewerImgRaw does not start pending providers while the active Read is blocked",
          success);

    fileSystem.DisableBlockingRead();
    static_cast<void>(SetEvent(firstControl->release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(firstControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
          L"ViewerImgRaw first blocked Read exits after release",
          success);

    WndMsg::ViewerImgRawResourceDebugSnapshot latestResource{};
    WndMsg::ViewerImgRawDecodeDebugSnapshot latestDecode{};
    Check(PumpUntil(
              [&]() noexcept
    {
        latestResource = {};
        latestDecode   = {};
        return viewerWindow &&
               SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&latestResource)) == TRUE &&
               SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0u, reinterpret_cast<LPARAM>(&latestDecode)) == TRUE &&
               latestResource.finalSuccessCount == 1u && latestResource.activeMainDecodeCount == 0u && latestResource.pendingMainDecodeCount == 0u &&
               ! latestResource.loading && latestDecode.hasImage;
    },
              10000ms),
          L"ViewerImgRaw runs the latest pending request to one successful terminal result",
          success);
    Check(fileSystem.GetCreateFileReaderCount(middlePath) == 0u && fileSystem.GetCreateFileReaderCount(latestPath) == 1u,
          L"ViewerImgRaw permanently drops the superseded middle request and reads only the latest",
          success);
    Check(firstControl->maxRequestedBytes.load(std::memory_order_relaxed) <= 1u * 1024u * 1024u,
          L"ViewerImgRaw caps each provider Read request at one MiB",
          success);
    Check(host.WarningAlertCount() == 0u, L"ViewerImgRaw latest-wins decode raises no warning", success);

    const auto waitForFailureAfter = [&](uint64_t previousFailures, unsigned int previousWarnings, std::wstring_view label) noexcept
    {
        WndMsg::ViewerImgRawResourceDebugSnapshot snapshot{};
        const bool failed = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            return viewerWindow && SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) == TRUE &&
                   snapshot.finalFailureCount == previousFailures + 1u && ! snapshot.loading && snapshot.inflightDecodeCount == 0u &&
                   snapshot.activeMainDecodeCount == 0u && snapshot.pendingMainDecodeCount == 0u && host.WarningAlertCount() == previousWarnings + 1u;
        },
            10000ms);
        Check(failed, label, success);
        return snapshot;
    };

    std::error_code sizeEc;
    const uint64_t actualSize = static_cast<uint64_t>(std::filesystem::file_size(activePath, sizeEc));
    Check(! sizeEc && actualSize > 1u, L"ViewerImgRaw exact-reader fixture has a bounded nonempty size", success);

    auto runReaderFailure = [&](std::wstring_view requestLabel, std::wstring_view terminalLabel, auto&& arm, auto&& disarm) noexcept
    {
        const uint64_t priorFailures     = latestResource.finalFailureCount;
        const unsigned int priorWarnings = host.WarningAlertCount();
        fileSystem.ResetCreateFileReaderCounts();
        arm();
        context.focusedPath           = activeText.c_str();
        context.focusedOtherFileIndex = 0u;
        Check(SUCCEEDED(viewer->Open(&context)), requestLabel, success);
        latestResource = waitForFailureAfter(priorFailures, priorWarnings, terminalLabel);
        disarm();
    };

    runReaderFailure(L"ViewerImgRaw accepts an oversized-provider decode request",
                     L"ViewerImgRaw rejects a reported one-GiB-plus-one source before reading",
                     [&]() noexcept { fileSystem.EnableReportedSize((1024ull * 1024ull * 1024ull) + 1ull); },
                     [&]() noexcept { fileSystem.DisableReportedSize(); });
    Check(fileSystem.GetReadByteCount(activePath) == 0u, L"ViewerImgRaw source cap rejects before issuing Read", success);

    runReaderFailure(L"ViewerImgRaw accepts the seek-mismatch decode request",
                     L"ViewerImgRaw rejects a provider that lies about the initial seek position",
                     [&]() noexcept { fileSystem.EnableLocalReaderFault(LocalFileReaderFault::SeekPositionMismatch); },
                     [&]() noexcept { fileSystem.DisableLocalReaderFault(); });
    Check(fileSystem.GetReadByteCount(activePath) == 0u, L"ViewerImgRaw seek mismatch rejects before issuing Read", success);

    if (! sizeEc && actualSize > 1u)
    {
        runReaderFailure(L"ViewerImgRaw accepts the premature-EOF decode request",
                         L"ViewerImgRaw rejects premature EOF before the committed GetSize byte count",
                         [&]() noexcept { fileSystem.EnableReportedSize(actualSize + 1u); },
                         [&]() noexcept { fileSystem.DisableReportedSize(); });
        runReaderFailure(L"ViewerImgRaw accepts the trailing-byte decode request",
                         L"ViewerImgRaw rejects bytes beyond the committed GetSize byte count",
                         [&]() noexcept { fileSystem.EnableReportedSize(actualSize - 1u); },
                         [&]() noexcept { fileSystem.DisableReportedSize(); });
    }
    runReaderFailure(L"ViewerImgRaw accepts the impossible over-return decode request",
                     L"ViewerImgRaw rejects Read counts larger than the requested buffer",
                     [&]() noexcept { fileSystem.EnableLocalReaderFault(LocalFileReaderFault::OverReportedRead); },
                     [&]() noexcept { fileSystem.DisableLocalReaderFault(); });

    for (const auto& [environmentName, label] : std::array{
             std::pair{std::wstring_view(kAllocationFailureEnv), std::wstring_view(L"result-allocation")},
             std::pair{std::wstring_view(kPostFailureEnv), std::wstring_view(L"payload-post")},
         })
    {
        const uint64_t priorFailures     = latestResource.finalFailureCount;
        const unsigned int priorWarnings = host.WarningAlertCount();
        Check(SetEnvironmentVariableW(environmentName.data(), L"1") != FALSE, std::format(L"ViewerImgRaw arms the {} terminal-delivery fault", label), success);
        context.focusedPath           = activeText.c_str();
        context.focusedOtherFileIndex = 0u;
        Check(SUCCEEDED(viewer->Open(&context)), std::format(L"ViewerImgRaw starts the {} fault request", label), success);
        latestResource = waitForFailureAfter(
            priorFailures, priorWarnings, std::format(L"ViewerImgRaw {} failure exits loading through the allocation-free fallback", label));
        static_cast<void>(SetEnvironmentVariableW(environmentName.data(), nullptr));
        Check(viewerWindow && IsWindow(viewerWindow) != FALSE, std::format(L"ViewerImgRaw {} fallback preserves the live window", label), success);
    }

    const uint64_t successBeforeRecovery      = latestResource.finalSuccessCount;
    const unsigned int warningsBeforeRecovery = host.WarningAlertCount();
    context.focusedPath                       = activeText.c_str();
    context.focusedOtherFileIndex             = 0u;
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerImgRaw starts a normal decode after terminal faults", success);
    Check(PumpUntil(
              [&]() noexcept
    {
        latestResource = {};
        latestDecode   = {};
        return SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&latestResource)) == TRUE &&
               SendMessageW(viewerWindow, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0u, reinterpret_cast<LPARAM>(&latestDecode)) == TRUE &&
               latestResource.finalSuccessCount == successBeforeRecovery + 1u && ! latestResource.loading && latestResource.activeMainDecodeCount == 0u &&
               latestDecode.hasImage;
    },
              10000ms),
          L"ViewerImgRaw recovers with a successful decode after terminal-delivery faults",
          success);
    Check(host.WarningAlertCount() == warningsBeforeRecovery, L"ViewerImgRaw recovery adds no warning", success);

    const HRESULT settledCloseHr = viewer->Close();
    Check(SUCCEEDED(settledCloseHr), L"ViewerImgRaw closes the settled scheduler fixture", success);
    Check(! viewerWindow || PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerImgRaw destroys the settled scheduler window",
          success);
    information.reset();
    viewer.reset();

    struct CloseCounter final : IViewerCallback
    {
        CloseCounter()                               = default;
        CloseCounter(const CloseCounter&)            = delete;
        CloseCounter(CloseCounter&&)                 = delete;
        CloseCounter& operator=(const CloseCounter&) = delete;
        CloseCounter& operator=(CloseCounter&&)      = delete;
        HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override
        {
            cookieMatched.store(cookie == this, std::memory_order_release);
            count.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }
        std::atomic_uint32_t count{0u};
        std::atomic_bool cookieMatched{false};
    } closeCounter;

    BuiltinFileSystemStub blockedFileSystem;
    AlertRecordingHostStub blockedHost;
    auto closeControl = std::make_shared<BlockingReadControl>();
    Check(closeControl->IsValid(), L"ViewerImgRaw close-path blocking events are created", success);
    blockedFileSystem.EnableBlockingRead(closeControl);
    const auto releaseClose = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(closeControl->release.get())); });

    wil::com_ptr<IViewer> blockedViewer;
    Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&blockedHost), kViewerImgRawPluginId, blockedViewer.put_void())) &&
              blockedViewer,
          L"ViewerImgRaw creates a viewer for blocked-provider close",
          success);
    if (! blockedViewer)
    {
        return false;
    }
    static_cast<void>(blockedViewer->SetCallback(&closeCounter, &closeCounter));
    wil::com_ptr<IInformations> blockedInformation;
    static_cast<void>(blockedViewer->QueryInterface(__uuidof(IInformations), blockedInformation.put_void()));
    if (blockedInformation)
    {
        static_cast<void>(blockedInformation->SetConfiguration(R"json({"preferThumbnail":false,"prevCache":0,"nextCache":0})json"));
    }

    const wchar_t* blockedFiles[] = {activeText.c_str()};
    ViewerOpenContext blockedContext{};
    blockedContext.fileSystem                      = static_cast<IFileSystem*>(&blockedFileSystem);
    blockedContext.fileSystemName                  = L"File System";
    blockedContext.focusedPath                     = activeText.c_str();
    blockedContext.otherFiles                      = blockedFiles;
    blockedContext.otherFileCount                  = static_cast<unsigned long>(std::size(blockedFiles));
    blockedContext.focusedOtherFileIndex           = 0u;
    const std::vector<HWND> existingBlockedWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);
    Check(SUCCEEDED(blockedViewer->Open(&blockedContext)), L"ViewerImgRaw opens the blocked-provider fixture", success);
    HWND blockedWindow = nullptr;
    Check(PumpUntil(
              [&]() noexcept
    {
        blockedWindow = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingBlockedWindows);
        return blockedWindow && WaitForSingleObject(closeControl->entered.get(), 0u) == WAIT_OBJECT_0;
    },
              8000ms),
          L"ViewerImgRaw reaches blocked Read while its window remains responsive",
          success);

    const auto closeStarted = std::chrono::steady_clock::now();
    const HRESULT closeHr   = blockedViewer->Close();
    const auto closeElapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - closeStarted);
    Check(SUCCEEDED(closeHr) && closeElapsed < 500ms, L"ViewerImgRaw Close returns within 500 ms while the provider remains blocked", success);
    Check(! blockedWindow || IsWindow(blockedWindow) == FALSE, L"ViewerImgRaw destroys its HWND before the blocked provider returns", success);
    Check(WaitForSingleObject(closeControl->exited.get(), 0u) == WAIT_TIMEOUT, L"ViewerImgRaw provider remains blocked after non-waiting Close", success);
    Check(closeCounter.count.load(std::memory_order_acquire) == 1u && closeCounter.cookieMatched.load(std::memory_order_acquire),
          L"ViewerImgRaw reports ViewerClosed exactly once with the registered cookie",
          success);
    Check(SUCCEEDED(blockedViewer->SetCallback(nullptr, nullptr)),
          L"ViewerImgRaw clears the raw lifecycle callback before releasing the blocked viewer",
          success);

    blockedInformation.reset();
    blockedViewer.reset();
    Check(blockedFileSystem.GetReferenceCount() > 1u, L"ViewerImgRaw detached worker retains the provider and viewer lifetime", success);
    shutdownFn();
    moduleShutdownPending = false;
    Check(canUnloadFn() == FALSE, L"ViewerImgRaw runtime unload gate rejects refresh while a provider Read is blocked", success);

    pluginModule.reset();
    Check(GetModuleHandleW(pluginModuleName.c_str()) != nullptr,
          L"ViewerImgRaw worker keeps its DLL mapped after the caller releases the module handle",
          success);
    static_cast<void>(SetEvent(closeControl->release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(closeControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 8000ms),
          L"ViewerImgRaw blocked provider exits after release",
          success);
    blockedFileSystem.DisableBlockingRead();
    Check(PumpUntil([&]() noexcept { return blockedFileSystem.GetReferenceCount() == 1u; }, 8000ms),
          L"ViewerImgRaw worker and provider self-references retire exactly once",
          success);
    Check(PumpUntil([&]() noexcept { return GetModuleHandleW(pluginModuleName.c_str()) == nullptr; }, 8000ms),
          L"ViewerImgRaw callback-return module pin unloads only after blocked work returns",
          success);
    Check(closeCounter.count.load(std::memory_order_acquire) == 1u, L"ViewerImgRaw does not report a second close after worker retirement", success);
    return success;
#endif
}

[[nodiscard]] bool TestViewerImgRawEmbeddedThumbnailTerminalSequencing() noexcept
{
#ifndef _DEBUG
    std::wcout << L"[SKIP] ViewerImgRaw embedded-thumbnail sequencing requires ENABLE_TESTS hooks.\n";
    return true;
#else
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerImgRaw thumbnail-sequence executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for thumbnail sequencing", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    const DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    const DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    const auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            static_cast<void>(RemoveDllDirectory(pluginCookie));
        }
        if (buildCookie)
        {
            static_cast<void>(RemoveDllDirectory(buildCookie));
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads for thumbnail sequencing", success);
    if (! pluginModule)
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    const FARPROC shutdownProc     = GetProcAddress(pluginModule.get(), "RedSalamanderPluginShutdown");
    const FARPROC canUnloadProc    = GetProcAddress(pluginModule.get(), "RedSalamanderPluginCanUnloadNow");
    RedSalamanderCreateFn createFn = nullptr;
    using PluginShutdownFn         = void(__stdcall*)();
    using PluginCanUnloadFn        = BOOL(__stdcall*)();
    PluginShutdownFn shutdownFn    = nullptr;
    PluginCanUnloadFn canUnloadFn  = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    static_assert(sizeof(shutdownFn) == sizeof(shutdownProc));
    static_assert(sizeof(canUnloadFn) == sizeof(canUnloadProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    std::memcpy(&shutdownFn, &shutdownProc, sizeof(shutdownFn));
    std::memcpy(&canUnloadFn, &canUnloadProc, sizeof(canUnloadFn));
    Check(createFn && shutdownFn && canUnloadFn, L"ViewerImgRaw factory and quiet-point exports resolve for thumbnail sequencing", success);
    if (! createFn || ! shutdownFn || ! canUnloadFn)
    {
        return false;
    }
    bool moduleShutdownPending         = true;
    const auto shutdownModuleOnFailure = wil::scope_exit([&]() noexcept
    {
        if (moduleShutdownPending)
        {
            shutdownFn();
        }
    });

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_thumbnail_sequence", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw thumbnail-sequence TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    const std::filesystem::path embeddedPath = WriteTinyWicJpegFile(tempDir / L"synthetic-embedded-thumbnail.dng");
    Check(std::filesystem::exists(embeddedPath), L"ViewerImgRaw synthetic embedded-thumbnail payload is available", success);
    if (! std::filesystem::exists(embeddedPath))
    {
        return false;
    }
    const auto cleanupFixture = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(embeddedPath, cleanupEc));
    });

    constexpr wchar_t kEmbeddedPayloadEnv[]  = L"REDSALAMANDER_VIEWERIMGRAW_TEST_EMBEDDED_JPEG_PAYLOAD";
    const std::wstring previousEmbeddedValue = GetEnvironmentString(kEmbeddedPayloadEnv);
    const auto restoreEnvironment            = wil::scope_exit([&]() noexcept
    { static_cast<void>(SetEnvironmentVariableW(kEmbeddedPayloadEnv, previousEmbeddedValue.empty() ? nullptr : previousEmbeddedValue.c_str())); });
    Check(SetEnvironmentVariableW(kEmbeddedPayloadEnv, L"1") != FALSE, L"ViewerImgRaw synthetic embedded-thumbnail seam is enabled", success);

    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const std::wstring embeddedText = embeddedPath.wstring();
    const wchar_t* otherFiles[]     = {embeddedText.c_str()};
    const auto runMode              = [&](bool preferThumbnail, std::wstring_view label) noexcept
    {
        AlertRecordingHostStub host;
        BuiltinFileSystemStub fileSystem;
        wil::com_ptr<IViewer> viewer;
        Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&host), kViewerImgRawPluginId, viewer.put_void())) && viewer,
              std::format(L"{} creates a viewer", label),
              success);
        if (! viewer)
        {
            return;
        }

        wil::com_ptr<IInformations> information;
        static_cast<void>(viewer->QueryInterface(__uuidof(IInformations), information.put_void()));
        Check(information && SUCCEEDED(information->SetConfiguration(preferThumbnail ? R"json({"preferThumbnail":true,"prevCache":0,"nextCache":0})json"
                                                                                     : R"json({"preferThumbnail":false,"prevCache":0,"nextCache":0})json")),
              std::format(L"{} accepts deterministic configuration", label),
              success);

        ViewerOpenContext context{};
        context.fileSystem                      = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName                  = L"File System";
        context.focusedPath                     = embeddedText.c_str();
        context.otherFiles                      = otherFiles;
        context.otherFileCount                  = static_cast<unsigned long>(std::size(otherFiles));
        context.focusedOtherFileIndex           = 0u;
        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);
        Check(SUCCEEDED(viewer->Open(&context)), std::format(L"{} opens the synthetic embedded payload", label), success);

        HWND window = nullptr;
        WndMsg::ViewerImgRawResourceDebugSnapshot resource{};
        WndMsg::ViewerImgRawDecodeDebugSnapshot decode{};
        const bool terminal = PumpUntil(
            [&]() noexcept
        {
            window   = FindNewVisibleWindowByClass(kViewerImgRawWindowClassName, existingWindows);
            resource = {};
            decode   = {};
            if (! window || SendMessageW(window, WndMsg::kViewerImgRawDebugGetResourceSnapshot, 0u, reinterpret_cast<LPARAM>(&resource)) == FALSE ||
                SendMessageW(window, WndMsg::kViewerImgRawDebugGetDecodeSnapshot, 0u, reinterpret_cast<LPARAM>(&decode)) == FALSE)
            {
                return false;
            }
            const bool expectedSequence =
                preferThumbnail ? resource.finalSuccessCount == 1u && resource.previewSuccessCount == 0u && resource.lastPreviewApplyOrdinal == 0u &&
                                      resource.lastFinalApplyOrdinal > 0u && decode.displayingThumbnail
                                : resource.finalSuccessCount == 1u && resource.previewSuccessCount == 1u && resource.lastPreviewApplyOrdinal > 0u &&
                                      resource.lastPreviewApplyOrdinal < resource.lastFinalApplyOrdinal && ! decode.displayingThumbnail;
            return expectedSequence && resource.finalFailureCount == 0u && ! resource.loading && resource.activeMainDecodeCount == 0u &&
                   resource.pendingMainDecodeCount == 0u && decode.hasImage;
        },
            10000ms);
        Check(terminal,
              preferThumbnail ? L"ViewerImgRaw embedded thumbnail is a single final success with no preview"
                              : L"ViewerImgRaw raw mode applies the embedded preview before one final full-image success",
              success);
        Check(host.WarningAlertCount() == 0u, std::format(L"{} raises no terminal warning", label), success);
        Check(fileSystem.GetCreateFileReaderCount(embeddedPath) == 1u, std::format(L"{} reads the source exactly once", label), success);

        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), std::format(L"{} closes", label), success);
        Check(! window || PumpUntil([&]() noexcept { return IsWindow(window) == FALSE; }, 5000ms), std::format(L"{} destroys its window", label), success);
    };

    runMode(true, L"ViewerImgRaw thumbnail-mode terminal sequence");
    runMode(false, L"ViewerImgRaw raw-mode preview/final sequence");
    shutdownFn();
    moduleShutdownPending = false;
    Check(PumpUntil([&]() noexcept { return canUnloadFn() == TRUE; }, 5000ms), L"ViewerImgRaw thumbnail sequencing reaches its module quiet point", success);
    const std::wstring pluginModuleName = pluginPath.filename().wstring();
    pluginModule.reset();
    Check(PumpUntil([&]() noexcept { return GetModuleHandleW(pluginModuleName.c_str()) == nullptr; }, 5000ms),
          L"ViewerImgRaw thumbnail sequencing unloads after explicit shutdown",
          success);
    return success;
#endif
}

// Regression test for plan 015: ViewerImgRaw must initialize COM on its decode threadpool
// workers so WIC-backed formats (PNG/GIF/BMP/non-RAW TIFF/HEIC) decode. Before the fix,
// DecodeImageToBgraWic's CoCreateInstance returned CO_E_NOTINITIALIZED on the worker thread and
// every WIC format silently failed to open. JPEG masked this because it uses turbojpeg (no COM),
// which is why the existing TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo case (which
// opens JPEG) never caught it. Here we open a 1x1 PNG and assert ViewerImgRaw did NOT raise a
// decode-failure warning alert. Note: a child-window count check would NOT work -- ViewerImgRaw
// renders both image content and the "no image" status via Direct2D (no content HWND), so the
// combo-host chrome appears whether or not decode succeeds.
[[nodiscard]] bool TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerImgRaw PNG test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerImgRaw.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerImgRaw.dll is present for PNG decode validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerImgRaw.dll loads for PNG decode validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerImgRaw factory export is available for PNG decode validation", success);
    if (! createFn)
    {
        return false;
    }

    AlertRecordingHostStub hostStub;

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&hostStub), kViewerImgRawPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerImgRaw factory creates an IViewer instance (PNG decode)", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerImgRawWindowClassName);

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerimgraw_png_decode", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerImgRaw PNG fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
    const std::filesystem::path pngPath = WriteTinyPngFile(tempDir / L"viewerimgraw-decode.png");
    auto cleanupTemp                    = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(pngPath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    const std::wstring pngPathText = pngPath.wstring();
    const wchar_t* otherFiles[]    = {pngPathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = pngPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerImgRaw open succeeds for PNG", success);
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
    Check(openedWindow, L"ViewerImgRaw window becomes visible for PNG", success);

    // Wait for the async decode to finish. Each alert site fires only at decode completion
    // (success -> ClearAlert; failure -> ShowAlert(HOST_ALERT_WARNING)), never at open-start, so
    // the first alert activity is the terminal decode signal.
    const bool decodeCompleted = PumpUntil([&]() noexcept { return hostStub.DecodeCompleted(); }, 8000ms);
    Check(decodeCompleted, L"ViewerImgRaw PNG decode completes (host alert observed)", success);

    // Regression assertion: a successful WIC decode raises no warning alert. On the pre-fix code
    // DecodeImageToBgraWic returns CO_E_NOTINITIALIZED and the async-failure path raises
    // ShowAlert(HOST_ALERT_WARNING), so WarningAlertCount() would be non-zero.
    Check(hostStub.WarningAlertCount() == 0u, L"ViewerImgRaw decodes PNG via WIC without a decode-failure alert", success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerImgRaw window close succeeds (PNG)", success);
    if (viewerWindow != nullptr)
    {
        Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerImgRaw window closes cleanly (PNG)", success);
    }
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_combo", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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
    CheckDxComboHostCompactChrome(viewerWindow, kViewerTextFileComboId, L"ViewerText", success);
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
    CheckEmbeddedViewerHidesStandaloneFileCombo(
        createFn, kViewerTextPluginId, context, kViewerTextWindowClassName, kViewerTextFileComboId, L"ViewerText", success);
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_hex_byte_colors", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText hex-color fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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

// Regression for plan 020: on a short read (IFileReader::Read yields fewer bytes than the reported file
// size), the hex view must show only the bytes actually read — never a fabricated zero-padded tail. The
// reader reports 16 bytes via GetSize() but only yields the first 8; the visible byte count must settle
// at 8, not 16. This exercises the dominant async-preload load path (ViewerText.cpp), which shares the
// resize-to-bytes-read fix with LoadHexData.
[[nodiscard]] bool TestViewerTextHexShortReadDropsPhantomTailBytes() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"ViewerText hex short-read test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for hex short-read validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerText.dll loads successfully for hex short-read validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available for hex short-read validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates an IViewer instance for hex short-read validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
    Check(SUCCEEDED(infoHr) && info != nullptr, L"ViewerText exposes IInformations for hex short-read validation", success);
    if (FAILED(infoHr) || ! info)
    {
        return false;
    }

    const HRESULT configHr = info->SetConfiguration(R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1"})json");
    Check(SUCCEEDED(configHr), L"ViewerText accepts configuration for hex short-read validation", success);
    if (FAILED(configHr))
    {
        return false;
    }

    const ViewerTheme normalTheme = MakeViewerTextTestTheme(false);
    const HRESULT themeHr         = viewer->SetTheme(&normalTheme);
    Check(SUCCEEDED(themeHr), L"ViewerText accepts the normal theme for hex short-read validation", success);
    if (FAILED(themeHr))
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);

    // One full 16-byte hex row on disk (GetSize reports 16), but the reader only yields the first 8 bytes.
    static constexpr uint64_t kReportedSize = 16u;
    static constexpr uint64_t kReadableSize = 8u;
    static constexpr auto kShortReadFixture = std::to_array<std::byte>({
        std::byte{0x11},
        std::byte{0x22},
        std::byte{0x33},
        std::byte{0x44},
        std::byte{0x55},
        std::byte{0x66},
        std::byte{0x77},
        std::byte{0x88},
        std::byte{0x99},
        std::byte{0xAA},
        std::byte{0xBB},
        std::byte{0xCC},
        std::byte{0xDD},
        std::byte{0xEE},
        std::byte{0xF0},
        std::byte{0x0F},
    });
    static_assert(kShortReadFixture.size() == static_cast<size_t>(kReportedSize));

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_hex_short_read", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText short-read fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
    const std::filesystem::path samplePath = WriteBinaryFile(tempDir / L"viewertext-hex-short-read.bin", std::span(kShortReadFixture));
    auto cleanupTemp                       = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
    });

    BuiltinFileSystemStub fileSystem;
    fileSystem.EnableShortRead(kReadableSize);

    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.focusedPath    = samplePathText.c_str();
    context.flags          = VIEWER_OPEN_FLAG_START_HEX;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerText opens the short-read fixture in hex mode", success);
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
    Check(openedWindow, L"ViewerText hex short-read window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    // The painted/visible byte count must equal the bytes actually read (8), never the reported file
    // size (16). Pre-fix the zero-padded tail is treated as real data and this settles at 16.
    WndMsg::ViewerTextDebugSnapshot snapshot{};
    const bool snapshotReady = WaitForViewerTextSnapshot(viewerWindow, [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
        return value.viewMode == WndMsg::ViewerTextDebugViewMode::Hex && value.renderCount > 0u && value.visibleByteCount == static_cast<size_t>(kReadableSize);
    }, 5000ms, &snapshot);
    Check(snapshotReady, L"ViewerText hex view shows only the bytes actually read on a short read (no phantom zero tail)", success);
    Check(snapshot.visibleByteCount == static_cast<size_t>(kReadableSize),
          L"ViewerText hex visible byte count equals bytes read, not the reported file size",
          success);
    Check(snapshot.visibleByteCount < static_cast<size_t>(kReportedSize),
          L"ViewerText hex visible byte count is strictly below the reported file size on a short read",
          success);

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"ViewerText hex short-read window closes cleanly", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"ViewerText hex short-read window is destroyed after close", success);
    return success;
}

[[nodiscard]] bool TestViewerTextDecodeAndClipboardSafetyHelpers() noexcept
{
    bool success = true;

    struct DbcsFixture
    {
        UINT codePage = 0u;
        std::array<uint8_t, 4> bytes{};
        std::wstring_view name;
    };

    const std::array<DbcsFixture, 3> fixtures{{
        {ViewerTextSafety::kShiftJisCodePage, {{0x41u, 0x82u, 0xA0u, 0x42u}}, L"Shift-JIS"},
        {ViewerTextSafety::kGbkCodePage, {{0x41u, 0xC4u, 0xE3u, 0x42u}}, L"GBK"},
        {ViewerTextSafety::kBig5CodePage, {{0x41u, 0xA4u, 0x40u, 0x42u}}, L"Big5"},
    }};

    const auto decode = [](const UINT codePage, const uint8_t* bytes, const size_t size) -> std::wstring
    {
        if (! bytes || size == 0u || size > static_cast<size_t>((std::numeric_limits<int>::max)()))
        {
            return {};
        }

        const int sourceLength = static_cast<int>(size);
        const int required     = MultiByteToWideChar(codePage, 0u, reinterpret_cast<LPCCH>(bytes), sourceLength, nullptr, 0);
        if (required <= 0)
        {
            return {};
        }

        std::wstring decoded(static_cast<size_t>(required), L'\0');
        const int written = MultiByteToWideChar(codePage, 0u, reinterpret_cast<LPCCH>(bytes), sourceLength, decoded.data(), required);
        if (written <= 0)
        {
            return {};
        }
        decoded.resize(static_cast<size_t>(written));
        return decoded;
    };

    for (const DbcsFixture& fixture : fixtures)
    {
        const std::wstring expected = decode(fixture.codePage, fixture.bytes.data(), fixture.bytes.size());
        Check(! expected.empty(), std::format(L"ViewerText {} fixture decodes with the selected Windows code page", fixture.name), success);

        for (size_t split = 0u; split <= fixture.bytes.size(); ++split)
        {
            const size_t carry = ViewerTextSafety::IncompleteDbcsTailSize(fixture.bytes.data(), split, fixture.codePage);
            Check(carry <= 1u && carry <= split, std::format(L"ViewerText {} split {} reports a bounded DBCS carry", fixture.name, split), success);
            if (split == 2u)
            {
                Check(carry == 1u, std::format(L"ViewerText {} preserves the lead byte at its only incomplete-pair boundary", fixture.name), success);
            }

            const size_t prefixSize    = split - carry;
            std::wstring streamedText  = decode(fixture.codePage, fixture.bytes.data(), prefixSize);
            const size_t pendingOffset = split - carry;
            streamedText.append(decode(fixture.codePage, fixture.bytes.data() + pendingOffset, fixture.bytes.size() - pendingOffset));
            Check(streamedText == expected, std::format(L"ViewerText {} split {} decodes identically to the unsplit fixture", fixture.name, split), success);
        }
    }

    static constexpr std::array<uint8_t, 4> kGrinningFaceUtf8{{0xF0u, 0x9Fu, 0x98u, 0x80u}};
    const ViewerTextSafety::Utf8Scalar scalar = ViewerTextSafety::DecodeUtf8Scalar(kGrinningFaceUtf8.data(), kGrinningFaceUtf8.size());
    Check(scalar.valid && scalar.codePoint == 0x1F600u && scalar.consumed == kGrinningFaceUtf8.size(),
          L"ViewerText UTF-8 scalar decoding preserves U+1F600 instead of replacing it",
          success);

    const uint64_t cap                                = ViewerTextSafety::kMaxHexClipboardBytes;
    const ViewerTextSafety::HexClipboardPlan accepted = ViewerTextSafety::ComputeHexClipboardPlan(cap, 0u, (std::numeric_limits<uint64_t>::max)());
    Check(accepted.hasData && ! accepted.truncated && accepted.copiedBytes == cap && accepted.rejectedBytes == 0u,
          L"ViewerText hex clipboard accepts exactly the synchronous byte cap",
          success);

    const ViewerTextSafety::HexClipboardPlan firstTruncated = ViewerTextSafety::ComputeHexClipboardPlan(cap + 1u, 0u, (std::numeric_limits<uint64_t>::max)());
    Check(firstTruncated.hasData && firstTruncated.truncated && firstTruncated.copiedBytes == cap && firstTruncated.rejectedBytes == 1u,
          L"ViewerText hex clipboard truncates the first byte above the cap",
          success);

    const ViewerTextSafety::HexClipboardPlan maximum =
        ViewerTextSafety::ComputeHexClipboardPlan((std::numeric_limits<uint64_t>::max)(), 0u, (std::numeric_limits<uint64_t>::max)());
    Check(maximum.hasData && maximum.truncated && maximum.copiedBytes == cap && maximum.requestedBytes == (std::numeric_limits<uint64_t>::max)() &&
              maximum.lastLine < (cap / ViewerTextSafety::kHexBytesPerLine),
          L"ViewerText hex clipboard computes a bounded plan for UINT64_MAX without arithmetic wrap",
          success);

    return success;
}

[[nodiscard]] bool TestViewerTextAsyncOpenAndUtf8HexTerminalContracts() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerText terminal-contract test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for terminal-contract validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerText.dll loads for terminal-contract validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available for terminal-contract validation", success);
    if (! createFn)
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_terminal_decode", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText terminal-contract fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    std::vector<std::byte> fixture(8u * 1024u * 1024u + 17u, std::byte{0x2Eu});
    static constexpr std::array<std::byte, 4> kFace{{std::byte{0xF0}, std::byte{0x9F}, std::byte{0x98}, std::byte{0x80}}};
    for (const size_t start : std::array<size_t, 3>{{0u, 16u + 5u, 32u + 12u}})
    {
        std::copy(kFace.begin(), kFace.end(), fixture.begin() + static_cast<ptrdiff_t>(start));
    }
    static constexpr std::array<std::byte, 10> kGeometryText{{std::byte{0x41},
                                                              std::byte{0x09},
                                                              std::byte{0xE4},
                                                              std::byte{0xB8},
                                                              std::byte{0xAD},
                                                              std::byte{0xF0},
                                                              std::byte{0x9F},
                                                              std::byte{0x98},
                                                              std::byte{0x80},
                                                              std::byte{0x42}}};
    std::copy(kGeometryText.begin(), kGeometryText.end(), fixture.begin() + 4);
    fixture[256] = std::byte{0x0A};

    const std::filesystem::path samplePath = WriteBinaryFile(tempDir / L"viewertext-utf8-nonbmp.txt", std::span(fixture));
    auto cleanupFixture                    = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(samplePath, cleanupEc));
    });

    AlertRecordingHostStub hostStub;
    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, static_cast<IHost*>(&hostStub), kViewerTextPluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates a viewer for terminal-contract validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> information;
    const HRESULT informationHr = viewer->QueryInterface(__uuidof(IInformations), information.put_void());
    Check(SUCCEEDED(informationHr) && information != nullptr, L"ViewerText exposes IInformations for streamed text-window validation", success);
    if (FAILED(informationHr) || ! information)
    {
        return false;
    }

    const HRESULT configHr = information->SetConfiguration(R"json({"textBufferMiB":2,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1"})json");
    Check(SUCCEEDED(configHr), L"ViewerText accepts a two-MiB streamed text-window configuration", success);

    const ViewerTheme normalTheme = MakeViewerTextTestTheme(false);
    const HRESULT themeHr         = viewer->SetTheme(&normalTheme);
    Check(SUCCEEDED(themeHr), L"ViewerText accepts a theme for terminal-contract validation", success);

    BuiltinFileSystemStub fileSystem;
    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName = L"File System";
    context.focusedPath    = samplePathText.c_str();

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);
    const HRESULT openHr                    = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"ViewerText opens the terminal-contract fixture", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    auto emergencyCleanup   = wil::scope_exit([&]() noexcept
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
    Check(openedWindow, L"ViewerText terminal-contract window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot snapshot{};
    const bool initialTerminal = WaitForViewerTextSnapshot(viewerWindow, [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
        return ! value.isLoading && value.asyncOpenTerminalCount >= 1u && SUCCEEDED(value.asyncOpenLastTerminalHr);
    }, 5000ms, &snapshot);
    Check(initialTerminal, L"ViewerText initial async open reaches one successful terminal result", success);

    const HWND textViewWindow = FindWindowExW(viewerWindow, nullptr, kViewerTextViewWindowClassName, nullptr);
    Check(textViewWindow != nullptr, L"ViewerText exposes its DirectWrite text surface for geometry and stream validation", success);
    if (textViewWindow)
    {
        static_cast<void>(SetWindowPos(textViewWindow, nullptr, 0, 0, 48, 200, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE));
        InvalidateRect(textViewWindow, nullptr, TRUE);
        WndMsg::ViewerTextDebugSnapshot sparseSnapshot{};
        const bool sparseReady = WaitForViewerTextSnapshot(viewerWindow,
                                                           [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
        {
            return value.textSparseWrapActive && ! value.textVisualLineCountExact && value.textVisualLineCount > 65536u &&
                   value.textMaterializedVisualLineCount == 0u && value.textSparseLogicalSummaryCount == 2u && value.renderCount > 0u;
        },
                                                           5000ms,
                                                           &sparseSnapshot);
        Check(sparseReady,
              L"ViewerText represents multi-million-code-unit wrapped text with one sparse summary per logical line and no wrap-row vector",
              success);
        if (sparseReady)
        {
            WndMsg::ViewerTextDebugGeometryRequest wrappedCoverage{};
            wrappedCoverage.operation   = WndMsg::ViewerTextDebugGeometryOperation::ProbeWrappedCoverage;
            wrappedCoverage.logicalLine = 0u;
            const bool wrappedCoverageProbed =
                SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&wrappedCoverage)) != FALSE &&
                SUCCEEDED(wrappedCoverage.result);
            std::wcout << std::format(
                L"[INFO] ViewerText wrapped coverage: probed={} segments={} start={} end={} second=[{},{}] gap={} fit={} width={} line_height={}\n",
                wrappedCoverageProbed,
                wrappedCoverage.wrappedSegmentCount,
                wrappedCoverage.wrappedCoveredStart,
                wrappedCoverage.wrappedCoveredEnd,
                wrappedCoverage.wrappedSecondSegmentStart,
                wrappedCoverage.wrappedSecondSegmentEnd,
                wrappedCoverage.wrappedHasGapOrOverlap,
                wrappedCoverage.wrappedAllSegmentsFit,
                wrappedCoverage.wrappedWidthDip,
                wrappedCoverage.wrappedLineHeightDip);
            Check(wrappedCoverageProbed && wrappedCoverage.wrappedSegmentCount > 1u &&
                      wrappedCoverage.wrappedCoveredEnd > wrappedCoverage.wrappedCoveredStart && ! wrappedCoverage.wrappedHasGapOrOverlap &&
                      wrappedCoverage.wrappedAllSegmentsFit && wrappedCoverage.wrappedWidthDip > 0.0f,
                  L"ViewerText sparse wrapping covers the mixed-width logical line with contiguous DirectWrite-fit segments",
                  success);

            const uint32_t initialTopColumn = sparseSnapshot.topVisibleSegmentColumnStart;
            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0u), 0u));
            WndMsg::ViewerTextDebugSnapshot lineDownSnapshot{};
            const bool lineDownAdvanced =
                TryGetViewerTextDebugSnapshot(viewerWindow, lineDownSnapshot) && lineDownSnapshot.topVisibleSegmentColumnStart > initialTopColumn;
            Check(lineDownAdvanced, L"ViewerText sparse line-down advances to the next DirectWrite-fit segment", success);
            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEUP, 0u), 0u));
            WndMsg::ViewerTextDebugSnapshot lineRoundTripSnapshot{};
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, lineRoundTripSnapshot) && lineRoundTripSnapshot.topVisibleSegmentColumnStart == initialTopColumn,
                  L"ViewerText sparse line-down then line-up returns to the identical segment anchor",
                  success);

            const WPARAM wheelDown = MAKEWPARAM(0u, static_cast<WORD>(static_cast<SHORT>(-WHEEL_DELTA)));
            const WPARAM wheelUp   = MAKEWPARAM(0u, static_cast<WORD>(static_cast<SHORT>(WHEEL_DELTA)));
            static_cast<void>(SendMessageW(textViewWindow, WM_MOUSEWHEEL, wheelDown, 0u));
            WndMsg::ViewerTextDebugSnapshot wheelDownSnapshot{};
            const bool wheelAdvanced =
                TryGetViewerTextDebugSnapshot(viewerWindow, wheelDownSnapshot) && wheelDownSnapshot.topVisibleSegmentColumnStart > initialTopColumn;
            Check(wheelAdvanced, L"ViewerText sparse wheel-down advances through physical wrapped rows", success);
            static_cast<void>(SendMessageW(textViewWindow, WM_MOUSEWHEEL, wheelUp, 0u));
            WndMsg::ViewerTextDebugSnapshot wheelRoundTripSnapshot{};
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, wheelRoundTripSnapshot) &&
                      wheelRoundTripSnapshot.topVisibleSegmentColumnStart == initialTopColumn,
                  L"ViewerText sparse wheel-down then wheel-up returns to the identical segment anchor",
                  success);

            const UINT textDpi       = GetDpiForWindow(textViewWindow);
            const auto pixelsFromDip = [textDpi](float dip) noexcept
            { return static_cast<int>(std::lround(static_cast<double>(dip) * static_cast<double>(textDpi) / 96.0)); };
            const int wrappedClickX        = pixelsFromDip(6.5f);
            const int wrappedClickY        = pixelsFromDip(6.0f + (std::max)(1.0f, wrappedCoverage.wrappedLineHeightDip) * 1.5f);
            const LPARAM wrappedClickPoint = MAKELPARAM(wrappedClickX, wrappedClickY);
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONDOWN, MK_LBUTTON, wrappedClickPoint));
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONUP, 0u, wrappedClickPoint));
            WndMsg::ViewerTextDebugSnapshot wrappedClickSnapshot{};
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, wrappedClickSnapshot) &&
                      wrappedClickSnapshot.textCaretIndex >= wrappedCoverage.wrappedSecondSegmentStart &&
                      wrappedClickSnapshot.textCaretIndex <= wrappedCoverage.wrappedSecondSegmentEnd,
                  L"ViewerText production hit-test maps a click on wrapped row two into that row's DirectWrite segment",
                  success);
            static_cast<void>(SetFocus(textViewWindow));
            static_cast<void>(SendMessageW(textViewWindow, WM_KEYDOWN, VK_DOWN, 0u));
            WndMsg::ViewerTextDebugSnapshot keyDownSnapshot{};
            const bool keyDownAdvanced =
                TryGetViewerTextDebugSnapshot(viewerWindow, keyDownSnapshot) && keyDownSnapshot.textCaretIndex > wrappedClickSnapshot.textCaretIndex;
            Check(keyDownAdvanced, L"ViewerText sparse down-arrow advances through the next physical wrapped row", success);
            static_cast<void>(SendMessageW(textViewWindow, WM_KEYDOWN, VK_UP, 0u));
            WndMsg::ViewerTextDebugSnapshot keyRoundTripSnapshot{};
            static_cast<void>(TryGetViewerTextDebugSnapshot(viewerWindow, keyRoundTripSnapshot));
            std::wcout << std::format(L"[INFO] ViewerText key round trip: before={} down={} up={}\n",
                                      wrappedClickSnapshot.textCaretIndex,
                                      keyDownSnapshot.textCaretIndex,
                                      keyRoundTripSnapshot.textCaretIndex);
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, keyRoundTripSnapshot) &&
                      keyRoundTripSnapshot.textCaretIndex == wrappedClickSnapshot.textCaretIndex,
                  L"ViewerText sparse down-arrow then up-arrow returns to the identical DirectWrite caret stop",
                  success);
            static_cast<void>(SendMessageW(textViewWindow, WM_KEYDOWN, VK_NEXT, 0u));
            WndMsg::ViewerTextDebugSnapshot pageDownSnapshot{};
            const bool pageDownAdvanced =
                TryGetViewerTextDebugSnapshot(viewerWindow, pageDownSnapshot) && pageDownSnapshot.textCaretIndex > keyRoundTripSnapshot.textCaretIndex;
            Check(pageDownAdvanced, L"ViewerText sparse Page Down advances the caret beyond one wrapped viewport", success);
            static_cast<void>(SendMessageW(textViewWindow, WM_KEYDOWN, VK_PRIOR, 0u));
            WndMsg::ViewerTextDebugSnapshot pageRoundTripSnapshot{};
            static_cast<void>(TryGetViewerTextDebugSnapshot(viewerWindow, pageRoundTripSnapshot));
            std::wcout << std::format(L"[INFO] ViewerText page round trip: before={} down={} up={}\n",
                                      wrappedClickSnapshot.textCaretIndex,
                                      pageDownSnapshot.textCaretIndex,
                                      pageRoundTripSnapshot.textCaretIndex);
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, pageRoundTripSnapshot) &&
                      pageRoundTripSnapshot.textCaretIndex == wrappedClickSnapshot.textCaretIndex,
                  L"ViewerText sparse Page Down then Page Up returns to the identical DirectWrite caret stop",
                  success);

            static_cast<void>(SetWindowPos(textViewWindow, nullptr, 0, 0, 48, 2400, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE));
            InvalidateRect(textViewWindow, nullptr, TRUE);
            UpdateWindow(textViewWindow);
            WndMsg::ViewerTextDebugSnapshot tallViewportSnapshot{};
            const bool tallViewportRendered = WaitForViewerTextSnapshot(viewerWindow, [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
                return value.visibleRowCount > 128u && value.textMaterializedVisualLineCount == 0u;
            }, 5000ms, &tallViewportSnapshot);
            Check(tallViewportRendered, L"ViewerText sparse viewport paints more than 128 physical rows without a hidden truncation cap", success);
            static_cast<void>(SetWindowPos(textViewWindow, nullptr, 0, 0, 48, 200, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE));
            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_TOP, 0u), 0u));
            InvalidateRect(textViewWindow, nullptr, TRUE);
            UpdateWindow(textViewWindow);

            Check(sparseSnapshot.textLastPaintUs < 250000u, L"ViewerText huge-line visible paint latency stays below 250 ms", success);

            SCROLLINFO sparseScroll{};
            sparseScroll.cbSize = sizeof(sparseScroll);
            sparseScroll.fMask  = SIF_RANGE;
            static_cast<void>(GetScrollInfo(textViewWindow, SB_VERT, &sparseScroll));
            const int middlePosition = sparseScroll.nMin + (sparseScroll.nMax - sparseScroll.nMin) / 2;
            static_cast<void>(SetScrollPos(textViewWindow, SB_VERT, middlePosition, TRUE));
            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_THUMBPOSITION, 0u), 0u));
            WndMsg::ViewerTextDebugSnapshot middleSnapshot{};
            const bool middleReachable = TryGetViewerTextDebugSnapshot(viewerWindow, middleSnapshot) &&
                                         middleSnapshot.topVisibleSegmentColumnStart > sparseSnapshot.topVisibleSegmentColumnStart;
            Check(middleReachable, L"ViewerText sparse wrapping keeps the middle of the huge line reachable", success);

            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_BOTTOM, 0u), 0u));
            WndMsg::ViewerTextDebugSnapshot endSnapshot{};
            const bool endReachable = TryGetViewerTextDebugSnapshot(viewerWindow, endSnapshot) &&
                                      endSnapshot.topVisibleSegmentColumnStart > middleSnapshot.topVisibleSegmentColumnStart;
            Check(endReachable, L"ViewerText sparse wrapping keeps the end of the huge line reachable", success);
        }

        static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_BOTTOM, 0u), 0u));
        for (size_t request = 0u; request < 8u; ++request)
        {
            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0u), 0u));
        }

        WndMsg::ViewerTextDebugSnapshot streamSnapshot{};
        const bool streamSettled = WaitForViewerTextSnapshot(viewerWindow,
                                                             [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
        {
            return ! value.textStreamLoadPending && value.textStreamAcceptedCount >= 2u && value.textStreamStaleCount >= 1u &&
                   value.textStreamTerminalCount >= 1u && SUCCEEDED(value.textStreamLastTerminalHr) && value.textStreamStartOffset > 0u &&
                   value.textStreamEndOffset == value.textFileSize;
        },
                                                             10000ms,
                                                             &streamSnapshot);
        Check(streamSettled, L"ViewerText rapid stream navigation rejects stale completions and accepts one identity-bound terminal window", success);
        if (streamSettled)
        {
            std::wcout << std::format(L"[INFO] ViewerText stream metrics: accepted={} stale={} terminal={} worker_us={} ui_apply_us={} visual_rows={} "
                                      L"sparse_summaries={} materialized_rows={} paint_us={}\n",
                                      streamSnapshot.textStreamAcceptedCount,
                                      streamSnapshot.textStreamStaleCount,
                                      streamSnapshot.textStreamTerminalCount,
                                      streamSnapshot.textStreamLastElapsedUs,
                                      streamSnapshot.textStreamLastUiApplyUs,
                                      streamSnapshot.textVisualLineCount,
                                      streamSnapshot.textSparseLogicalSummaryCount,
                                      streamSnapshot.textMaterializedVisualLineCount,
                                      streamSnapshot.textLastPaintUs);
            Check(streamSnapshot.textStreamLastUiApplyUs < 250000u, L"ViewerText streamed navigation UI apply latency stays below 250 ms", success);
            Check(streamSnapshot.textStreamLastElapsedUs > 0u, L"ViewerText reports worker decode/index and UI apply latency metrics", success);

            static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_TOP, 0u), 0u));
            for (size_t request = 0u; request < 4u; ++request)
            {
                static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEUP, 0u), 0u));
            }
            const uint64_t previousTerminalCount = streamSnapshot.textStreamTerminalCount;
            const bool beginningRestored         = WaitForViewerTextSnapshot(viewerWindow,
                                                                             [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
            {
                return ! value.textStreamLoadPending && value.textStreamStartOffset == 0u && value.textStreamTerminalCount > previousTerminalCount &&
                       SUCCEEDED(value.textStreamLastTerminalHr);
            },
                                                                     10000ms);
            Check(beginningRestored, L"ViewerText streamed navigation keeps the beginning window reachable after a latest-wins end jump", success);

            if (beginningRestored)
            {
                WndMsg::ViewerTextDebugSnapshot settledAtBeginning{};
                const bool allPriorRequestsSettled = WaitForViewerTextSnapshot(viewerWindow, [](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
                    return ! value.textStreamLoadPending && value.textStreamAcceptedCount == value.textStreamStaleCount + value.textStreamTerminalCount;
                }, 10000ms, &settledAtBeginning);
                Check(allPriorRequestsSettled, L"ViewerText retires every rapid navigation request before stream fault validation", success);

                const auto setNextStreamFault = [&](WndMsg::ViewerTextDebugAsyncStreamFault fault) noexcept
                {
                    WndMsg::ViewerTextDebugGeometryRequest faultRequest{};
                    faultRequest.operation        = WndMsg::ViewerTextDebugGeometryOperation::SetAsyncStreamFault;
                    faultRequest.asyncStreamFault = fault;
                    return SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&faultRequest)) != FALSE &&
                           SUCCEEDED(faultRequest.result);
                };
                const auto requestNextStreamWindow = [&]() noexcept
                { static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0u), 0u)); };

                struct StreamFaultCase final
                {
                    WndMsg::ViewerTextDebugAsyncStreamFault fault;
                    std::wstring_view label;
                    bool acceptedByWorker = false;
                };
                const std::array streamFaultCases{
                    StreamFaultCase{WndMsg::ViewerTextDebugAsyncStreamFault::Allocation, L"allocation", false},
                    StreamFaultCase{WndMsg::ViewerTextDebugAsyncStreamFault::Submit, L"submit", false},
                    StreamFaultCase{WndMsg::ViewerTextDebugAsyncStreamFault::Worker, L"worker", true},
                    StreamFaultCase{WndMsg::ViewerTextDebugAsyncStreamFault::PayloadPost, L"payload-post", true},
                };
                for (const StreamFaultCase& faultCase : streamFaultCases)
                {
                    WndMsg::ViewerTextDebugSnapshot beforeFault{};
                    Check(TryGetViewerTextDebugSnapshot(viewerWindow, beforeFault),
                          std::format(L"ViewerText captures metrics before the {} stream fault", faultCase.label),
                          success);
                    Check(setNextStreamFault(faultCase.fault), std::format(L"ViewerText accepts the deterministic {} stream fault", faultCase.label), success);
                    requestNextStreamWindow();

                    WndMsg::ViewerTextDebugSnapshot afterFault{};
                    const bool terminalFailure = WaitForViewerTextSnapshot(viewerWindow,
                                                                           [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
                    {
                        return ! value.textStreamLoadPending && value.textStreamTerminalCount == beforeFault.textStreamTerminalCount + 1u &&
                               FAILED(value.textStreamLastTerminalHr) && value.textStreamStartOffset == beforeFault.textStreamStartOffset;
                    },
                                                                           10000ms,
                                                                           &afterFault);
                    Check(terminalFailure,
                          std::format(L"ViewerText {} stream fault reaches exactly one terminal failure without replacing the visible window", faultCase.label),
                          success);
                    if (terminalFailure)
                    {
                        const uint64_t acceptedDelta = afterFault.textStreamAcceptedCount - beforeFault.textStreamAcceptedCount;
                        const uint64_t rejectedDelta = afterFault.textStreamRejectedCount - beforeFault.textStreamRejectedCount;
                        Check(acceptedDelta == (faultCase.acceptedByWorker ? 1u : 0u) && rejectedDelta == (faultCase.acceptedByWorker ? 0u : 1u),
                              std::format(L"ViewerText {} stream fault updates accepted/rejected metrics exactly once", faultCase.label),
                              success);
                    }
                }

                auto staleFailureControl = std::make_shared<BlockingReadControl>();
                Check(staleFailureControl->IsValid(), L"ViewerText creates a blocking provider seam for stale stream-failure validation", success);
                if (staleFailureControl->IsValid())
                {
                    WndMsg::ViewerTextDebugSnapshot beforeStaleFailure{};
                    static_cast<void>(TryGetViewerTextDebugSnapshot(viewerWindow, beforeStaleFailure));
                    fileSystem.EnableBlockingRead(staleFailureControl);
                    Check(setNextStreamFault(WndMsg::ViewerTextDebugAsyncStreamFault::Worker), L"ViewerText accepts a delayed worker stream fault", success);
                    requestNextStreamWindow();
                    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(staleFailureControl->entered.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
                          L"ViewerText delayed stream worker reaches the blocking provider",
                          success);

                    Check(setNextStreamFault(WndMsg::ViewerTextDebugAsyncStreamFault::Allocation),
                          L"ViewerText accepts a newer terminal stream fault while an older worker is blocked",
                          success);
                    requestNextStreamWindow();
                    WndMsg::ViewerTextDebugSnapshot currentFailure{};
                    const bool newerFailureTerminal = WaitForViewerTextSnapshot(viewerWindow,
                                                                                [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
                    {
                        return ! value.textStreamLoadPending && value.textStreamTerminalCount == beforeStaleFailure.textStreamTerminalCount + 1u &&
                               FAILED(value.textStreamLastTerminalHr);
                    },
                                                                                5000ms,
                                                                                &currentFailure);
                    Check(newerFailureTerminal, L"ViewerText newer stream failure becomes the sole current terminal state", success);

                    static_cast<void>(SetEvent(staleFailureControl->release.get()));
                    fileSystem.DisableBlockingRead();
                    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(staleFailureControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
                          L"ViewerText delayed failing stream worker exits after provider release",
                          success);
                    const bool staleFailureIgnored = WaitForViewerTextSnapshot(viewerWindow,
                                                                               [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
                    {
                        return value.textStreamStaleCount > beforeStaleFailure.textStreamStaleCount &&
                               value.textStreamTerminalCount == currentFailure.textStreamTerminalCount &&
                               value.textStreamLastTerminalHr == currentFailure.textStreamLastTerminalHr;
                    },
                                                                               5000ms);
                    Check(staleFailureIgnored, L"ViewerText stale worker failure cannot overwrite the newer terminal state", success);
                }
            }
        }
    }

    if (textViewWindow)
    {
        static_cast<void>(SetWindowPos(textViewWindow, nullptr, 0, 0, 320, 400, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE));
        static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_TOP, 0u), 0u));
        InvalidateRect(textViewWindow, nullptr, TRUE);
        UpdateWindow(textViewWindow);
    }

    const std::wstring_view preview(snapshot.textPreview);
    const size_t geometryStart = preview.find(L"A\t\u4E2D\U0001F600B");
    Check(geometryStart != std::wstring_view::npos, L"ViewerText geometry fixture contains tab, CJK, and a surrogate pair", success);
    if (geometryStart != std::wstring_view::npos)
    {
        WndMsg::ViewerTextDebugGeometryRequest budget{};
        budget.operation       = WndMsg::ViewerTextDebugGeometryOperation::SetCacheBudget;
        budget.cacheMaxEntries = 2u;
        budget.cacheMaxBytes   = 8192u;
        Check(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&budget)) != FALSE &&
                  SUCCEEDED(budget.result),
              L"ViewerText accepts a tiny deterministic visible-layout cache budget",
              success);

        std::array<float, 7> caretPositions{};
        for (size_t offset = 0u; offset < caretPositions.size(); ++offset)
        {
            WndMsg::ViewerTextDebugGeometryRequest probe{};
            probe.operation    = WndMsg::ViewerTextDebugGeometryOperation::ProbeLayout;
            probe.segmentStart = 0u;
            probe.segmentEnd   = geometryStart + 6u;
            probe.textPosition = geometryStart + offset;
            probe.rangeStart   = geometryStart + 1u;
            probe.rangeEnd     = geometryStart + 5u;
            probe.widthDip     = 320.0f;
            probe.hitX         = 0.0f;
            const bool probed  = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&probe)) != FALSE;
            Check(probed && SUCCEEDED(probe.result), std::format(L"ViewerText probes DirectWrite caret position {}", offset), success);
            caretPositions[offset] = probe.caretX;
            if (offset == 4u)
            {
                Check(probe.normalizedTextPosition == geometryStart + 5u,
                      L"ViewerText never exposes the interior of the emoji surrogate pair as a caret stop",
                      success);
            }
            if (offset == 3u)
            {
                Check(probe.nextTextPosition == geometryStart + 5u, L"ViewerText right-arrow movement crosses the emoji as one caret stop", success);
            }
            if (offset == 5u)
            {
                Check(probe.previousTextPosition == geometryStart + 3u, L"ViewerText left-arrow movement crosses the emoji as one caret stop", success);
            }
            if (offset == 1u)
            {
                Check(probe.rangeLeft == probe.caretX, L"ViewerText selection/search range start shares the caret DirectWrite mapping", success);
            }
            if (offset == 5u)
            {
                Check(probe.rangeRight == probe.caretX, L"ViewerText selection/search range end shares the caret DirectWrite mapping", success);
            }
        }
        Check(caretPositions[0] < caretPositions[1] && caretPositions[1] < caretPositions[2] && caretPositions[2] < caretPositions[3] &&
                  caretPositions[3] < caretPositions[5] && caretPositions[5] < caretPositions[6],
              L"ViewerText DirectWrite geometry advances monotonically across tab, CJK, emoji, and ASCII glyphs",
              success);

        for (const size_t offset : std::array<size_t, 4>{{1u, 2u, 3u, 5u}})
        {
            WndMsg::ViewerTextDebugGeometryRequest hit{};
            hit.operation        = WndMsg::ViewerTextDebugGeometryOperation::ProbeLayout;
            hit.segmentStart     = 0u;
            hit.segmentEnd       = geometryStart + 6u;
            hit.textPosition     = geometryStart + offset;
            hit.rangeStart       = geometryStart;
            hit.rangeEnd         = geometryStart + 6u;
            hit.widthDip         = 320.0f;
            hit.hitX             = caretPositions[offset];
            const bool hitProbed = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&hit)) != FALSE;
            Check(hitProbed && hit.hitTextPosition != geometryStart + 4u,
                  std::format(L"ViewerText click mapping {} never splits the emoji surrogate pair", offset),
                  success);
        }

        if (textViewWindow)
        {
            const UINT textDpi       = GetDpiForWindow(textViewWindow);
            const auto pixelsFromDip = [textDpi](float dip) noexcept
            { return static_cast<int>(std::lround(static_cast<double>(dip) * static_cast<double>(textDpi) / 96.0)); };
            const int geometryY               = pixelsFromDip(12.0f);
            const auto pointForGeometryOffset = [&](size_t offset) noexcept { return MAKELPARAM(pixelsFromDip(6.0f + caretPositions[offset]), geometryY); };

            const size_t clickOffset = 2u;
            const LPARAM clickPoint  = pointForGeometryOffset(clickOffset);
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint));
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONUP, 0u, clickPoint));
            WndMsg::ViewerTextDebugSnapshot productionClickSnapshot{};
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, productionClickSnapshot) && productionClickSnapshot.textCaretIndex == geometryStart + clickOffset,
                  L"ViewerText production mouse hit-test lands on the DirectWrite-probed CJK caret boundary",
                  success);

            const size_t dragStartOffset = 1u;
            const size_t dragEndOffset   = 5u;
            const LPARAM dragStartPoint  = pointForGeometryOffset(dragStartOffset);
            const LPARAM dragEndPoint    = pointForGeometryOffset(dragEndOffset);
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONDOWN, MK_LBUTTON, dragStartPoint));
            static_cast<void>(SendMessageW(textViewWindow, WM_MOUSEMOVE, MK_LBUTTON, dragEndPoint));
            static_cast<void>(SendMessageW(textViewWindow, WM_LBUTTONUP, 0u, dragEndPoint));
            WndMsg::ViewerTextDebugSnapshot productionDragSnapshot{};
            static_cast<void>(TryGetViewerTextDebugSnapshot(viewerWindow, productionDragSnapshot));
            std::wcout << std::format(L"[INFO] ViewerText drag selection: expected=[{},{}] actual=[{},{}] caret={}\n",
                                      geometryStart + dragStartOffset,
                                      geometryStart + dragEndOffset,
                                      productionDragSnapshot.textSelectionAnchor,
                                      productionDragSnapshot.textSelectionActive,
                                      productionDragSnapshot.textCaretIndex);
            Check(TryGetViewerTextDebugSnapshot(viewerWindow, productionDragSnapshot) &&
                      productionDragSnapshot.textSelectionAnchor == geometryStart + dragStartOffset &&
                      productionDragSnapshot.textSelectionActive == geometryStart + dragEndOffset,
                  L"ViewerText production drag selection shares the DirectWrite tab/CJK/emoji geometry mapping",
                  success);
        }

        WndMsg::ViewerTextDebugSnapshot geometrySnapshot{};
        Check(TryGetViewerTextDebugSnapshot(viewerWindow, geometrySnapshot) && geometrySnapshot.textLayoutCacheEntryCount <= 2u &&
                  geometrySnapshot.textLayoutCacheBytes <= 8192u,
              L"ViewerText keeps visible DirectWrite layouts within the tiny entry and byte budget",
              success);

        WndMsg::ViewerTextDebugGeometryRequest uncachedProbe{};
        uncachedProbe.operation    = WndMsg::ViewerTextDebugGeometryOperation::ProbeLayout;
        uncachedProbe.segmentStart = 0u;
        uncachedProbe.segmentEnd   = 4096u;
        uncachedProbe.textPosition = 1u;
        uncachedProbe.rangeStart   = 0u;
        uncachedProbe.rangeEnd     = 2u;
        uncachedProbe.widthDip     = 320.0f;
        const bool uncachedProbeSucceeded =
            SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSetLayoutCacheBudget, 0u, reinterpret_cast<LPARAM>(&uncachedProbe)) != FALSE &&
            SUCCEEDED(uncachedProbe.result);
        Check(uncachedProbeSucceeded, L"ViewerText keeps over-budget DirectWrite geometry correct through one uncached layout", success);
        WndMsg::ViewerTextDebugSnapshot uncachedSnapshot{};
        Check(TryGetViewerTextDebugSnapshot(viewerWindow, uncachedSnapshot) && uncachedSnapshot.textLayoutCacheEntryCount <= 2u &&
                  uncachedSnapshot.textLayoutCacheBytes <= 8192u,
              L"ViewerText never charges an oversized transient layout into the bounded cache",
              success);
        std::wcout << std::format(L"[INFO] ViewerText layout metrics: entries={} bytes={} max_entries={} max_bytes={} evictions={} hits={} misses={}\n",
                                  uncachedSnapshot.textLayoutCacheEntryCount,
                                  uncachedSnapshot.textLayoutCacheBytes,
                                  uncachedSnapshot.textLayoutCacheMaxEntries,
                                  uncachedSnapshot.textLayoutCacheMaxBytes,
                                  uncachedSnapshot.textLayoutCacheEvictions,
                                  uncachedSnapshot.textLayoutCacheHits,
                                  uncachedSnapshot.textLayoutCacheMisses);
    }

    constexpr wchar_t kHighSurrogate = static_cast<wchar_t>(0xD83Du);
    constexpr wchar_t kLowSurrogate  = static_cast<wchar_t>(0xDE00u);
    const std::array<std::pair<uint64_t, size_t>, 3> lineCases{{{0u, 0u}, {16u, 5u}, {32u, 12u}}};
    for (const auto& [offset, column] : lineCases)
    {
        WndMsg::ViewerTextDebugHexLineRequest request{};
        request.offset       = offset;
        const bool formatted = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugFormatUtf8HexLine, 0u, reinterpret_cast<LPARAM>(&request)) != FALSE;
        Check(formatted && request.validBytes == 16u, std::format(L"ViewerText formats the UTF-8 hex row at offset {}", offset), success);
        Check(request.text[column] == kHighSurrogate && request.text[column + 1u] == kLowSurrogate,
              std::format(L"ViewerText emits a valid U+1F600 surrogate pair at byte column {}", column),
              success);
        Check(request.sourceLengths[column] == 2u && request.sourceLengths[column + 1u] == 0u && request.columnStarts[column] == column &&
                  request.columnStarts[column + 1u] == column + 1u,
              std::format(L"ViewerText preserves byte-cell span alignment for U+1F600 at column {}", column),
              success);
    }

    const unsigned int warningsBeforeCopy = hostStub.WarningAlertCount();
    WndMsg::ViewerTextDebugHexCopyRequest acceptedCopyRequest{};
    acceptedCopyRequest.anchorOffset = 0u;
    acceptedCopyRequest.activeOffset = ViewerTextSafety::kMaxHexClipboardBytes - 1u;
    const bool acceptedCopyDispatched =
        SendMessageW(viewerWindow, WndMsg::kViewerTextDebugCopyHexSelection, 0u, reinterpret_cast<LPARAM>(&acceptedCopyRequest)) != FALSE;
    Check(acceptedCopyDispatched && acceptedCopyRequest.dispatched, L"ViewerText dispatches the real hex clipboard path at the exact source-byte cap", success);
    Check(hostStub.WarningAlertCount() == warningsBeforeCopy, L"ViewerText does not warn when a hex clipboard selection is exactly at the cap", success);

    WndMsg::ViewerTextDebugHexCopyRequest copyRequest{};
    copyRequest.anchorOffset  = 0u;
    copyRequest.activeOffset  = static_cast<uint64_t>(fixture.size() - 1u);
    const bool copyDispatched = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugCopyHexSelection, 0u, reinterpret_cast<LPARAM>(&copyRequest)) != FALSE;
    Check(copyDispatched && copyRequest.dispatched, L"ViewerText dispatches the real bounded hex clipboard path for an over-cap selection", success);
    Check(hostStub.WarningAlertCount() == warningsBeforeCopy + 1u,
          L"ViewerText shows one localized warning after truncating an over-cap hex clipboard selection",
          success);

    const std::array faults{
        WndMsg::ViewerTextDebugAsyncOpenFault::FileSystemIo,
        WndMsg::ViewerTextDebugAsyncOpenFault::OpenReader,
        WndMsg::ViewerTextDebugAsyncOpenFault::GetSize,
        WndMsg::ViewerTextDebugAsyncOpenFault::InitialSeek,
        WndMsg::ViewerTextDebugAsyncOpenFault::InitialRead,
        WndMsg::ViewerTextDebugAsyncOpenFault::DataSeek,
        WndMsg::ViewerTextDebugAsyncOpenFault::DataRead,
        WndMsg::ViewerTextDebugAsyncOpenFault::Decode,
        WndMsg::ViewerTextDebugAsyncOpenFault::PayloadPost,
        WndMsg::ViewerTextDebugAsyncOpenFault::Submit,
    };

    uint64_t expectedTerminalCount = snapshot.asyncOpenTerminalCount;
    for (const WndMsg::ViewerTextDebugAsyncOpenFault fault : faults)
    {
        const bool requested = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugReloadWithOpenFault, static_cast<WPARAM>(fault), 0u) != FALSE;
        Check(requested, std::format(L"ViewerText accepts deterministic async-open fault {}", static_cast<unsigned int>(fault)), success);
        ++expectedTerminalCount;

        WndMsg::ViewerTextDebugSnapshot faultSnapshot{};
        const bool terminal = WaitForViewerTextSnapshot(viewerWindow, [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
            return ! value.isLoading && value.asyncOpenTerminalCount == expectedTerminalCount && FAILED(value.asyncOpenLastTerminalHr);
        }, 5000ms, &faultSnapshot);
        Check(terminal, std::format(L"ViewerText async-open fault {} reaches exactly one failed terminal result", static_cast<unsigned int>(fault)), success);
    }

    const bool restoreRequested =
        SendMessageW(viewerWindow, WndMsg::kViewerTextDebugReloadWithOpenFault, static_cast<WPARAM>(WndMsg::ViewerTextDebugAsyncOpenFault::None), 0u) != FALSE;
    Check(restoreRequested, L"ViewerText accepts a clean reload after every injected failure", success);
    ++expectedTerminalCount;
    const bool restored = WaitForViewerTextSnapshot(viewerWindow, [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept {
        return ! value.isLoading && value.asyncOpenTerminalCount == expectedTerminalCount && SUCCEEDED(value.asyncOpenLastTerminalHr);
    }, 5000ms);
    Check(restored, L"ViewerText recovers with one successful terminal result after the fault sequence", success);

    auto closeControl = std::make_shared<BlockingReadControl>();
    Check(closeControl->IsValid(), L"ViewerText creates a blocking provider seam for non-blocking stream close validation", success);
    bool closeWorkerEntered = false;
    if (closeControl->IsValid() && textViewWindow && IsWindow(textViewWindow) != FALSE)
    {
        fileSystem.EnableBlockingRead(closeControl);
        static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_BOTTOM, 0u), 0u));
        static_cast<void>(SendMessageW(textViewWindow, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0u), 0u));
        closeWorkerEntered = PumpUntil([&]() noexcept { return WaitForSingleObject(closeControl->entered.get(), 0u) == WAIT_OBJECT_0; }, 5000ms);
        Check(closeWorkerEntered, L"ViewerText streamed navigation worker reaches the blocking provider before close", success);
        WndMsg::ViewerTextDebugSnapshot pendingCloseSnapshot{};
        Check(closeWorkerEntered && TryGetViewerTextDebugSnapshot(viewerWindow, pendingCloseSnapshot) && pendingCloseSnapshot.textStreamLoadPending,
              L"ViewerText exposes a pending streamed navigation request before close",
              success);
        fileSystem.DisableBlockingRead();
    }

    bool unloadStressQueued = true;
    for (size_t index = 0u; index < 32u; ++index)
    {
        unloadStressQueued =
            unloadStressQueued &&
            SendMessageW(viewerWindow, WndMsg::kViewerTextDebugReloadWithOpenFault, static_cast<WPARAM>(WndMsg::ViewerTextDebugAsyncOpenFault::None), 0u) !=
                FALSE;
    }
    Check(unloadStressQueued, L"ViewerText queues async-open unload stress work", success);

    const auto closeStartedAt = std::chrono::steady_clock::now();
    const HRESULT closeHr     = viewer->Close();
    const uint64_t closeElapsedUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - closeStartedAt).count());
    Check(SUCCEEDED(closeHr), L"ViewerText terminal-contract window closes while async-open callbacks are pending", success);
    Check(closeElapsedUs < 500000u, L"ViewerText close does not wait for a blocked streamed navigation callback", success);
    viewer.reset();
    information.reset();
    pluginModule.reset();
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerText terminal-contract window is destroyed with pending callbacks",
          success);
    if (closeWorkerEntered)
    {
        Check(GetModuleHandleW(pluginPath.filename().c_str()) != nullptr,
              L"ViewerText blocked stream callback retains the plugin module after explicit handles are released",
              success);
    }
    if (closeControl->IsValid())
    {
        static_cast<void>(SetEvent(closeControl->release.get()));
    }
    if (closeWorkerEntered)
    {
        Check(PumpUntil([&]() noexcept { return WaitForSingleObject(closeControl->exited.get(), 0u) == WAIT_OBJECT_0; }, 5000ms),
              L"ViewerText blocked stream callback exits after the provider is released",
              success);
    }
    Check(PumpUntil([&]() noexcept { return GetModuleHandleW(pluginPath.filename().c_str()) == nullptr; }, 10000ms),
          L"ViewerText async-open and stream callbacks defer the final module release until callback return",
          success);
    return success;
}

[[nodiscard]] bool TestViewerTextSaveAsPreservesDataOnFailures() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerText Save As test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerText.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerText.dll is present for Save As data-safety validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerText.dll loads for Save As data-safety validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerText factory export is available for Save As data-safety validation", success);
    if (! createFn)
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_save_as_safety", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText Save As fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

    constexpr std::string_view kSourceText            = "ViewerText transactional Save As source\r\nsecond line\r\n";
    constexpr std::string_view kDestinationText       = "pre-existing destination bytes must survive\r\n";
    const std::filesystem::path sourcePath            = WriteUtf8TextFile(tempDir / L"source.txt", kSourceText);
    const std::filesystem::path destinationPath       = WriteUtf8TextFile(tempDir / L"destination.txt", kDestinationText);
    const std::filesystem::path streamPath            = tempDir / L"streamed.txt";
    const std::filesystem::path streamDestinationPath = tempDir / L"stream-destination.txt";

    auto cleanupFiles = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupError;
        static_cast<void>(std::filesystem::remove(sourcePath, cleanupError));
        static_cast<void>(std::filesystem::remove(destinationPath, cleanupError));
        static_cast<void>(std::filesystem::remove(streamPath, cleanupError));
        static_cast<void>(std::filesystem::remove(streamDestinationPath, cleanupError));
    });

    const auto sourceBytes      = ReadBinaryFile(sourcePath);
    const auto destinationBytes = ReadBinaryFile(destinationPath);
    Check(sourceBytes.has_value() && destinationBytes.has_value(), L"ViewerText Save As baseline bytes are readable", success);
    if (! sourceBytes.has_value() || ! destinationBytes.has_value())
    {
        return false;
    }

    BuiltinFileSystemStub fileSystem;
    wil::com_ptr<IViewer> viewer;
    HWND viewerWindow = nullptr;

    auto closeViewer = [&]() noexcept
    {
        if (viewer)
        {
            const HRESULT closeHr = viewer->Close();
            Check(SUCCEEDED(closeHr), L"ViewerText Save As test viewer closes cleanly", success);
        }
        if (viewerWindow)
        {
            Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
                  L"ViewerText Save As test window is destroyed after close",
                  success);
        }
        viewerWindow = nullptr;
        viewer.reset();
    };
    auto emergencyCleanup = wil::scope_exit([&]() noexcept
    {
        if (viewerWindow && IsWindow(viewerWindow) != FALSE)
        {
            static_cast<void>(RequestCloseWindow(viewerWindow, 5000ms));
        }
        viewer.reset();
    });

    const auto openViewer = [&](const std::filesystem::path& path, const char* configuration, bool expectStream) noexcept -> bool
    {
        const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerTextPluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerText factory creates a viewer for Save As data-safety validation", success);
        if (FAILED(createHr) || ! viewer)
        {
            return false;
        }

        if (configuration)
        {
            wil::com_ptr<IInformations> information;
            const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), information.put_void());
            Check(SUCCEEDED(infoHr) && information != nullptr, L"ViewerText Save As test obtains configuration interface", success);
            if (FAILED(infoHr) || ! information)
            {
                return false;
            }
            const HRESULT configHr = information->SetConfiguration(configuration);
            Check(SUCCEEDED(configHr), L"ViewerText Save As test applies deterministic text buffer configuration", success);
            if (FAILED(configHr))
            {
                return false;
            }
        }

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerTextWindowClassName);
        const std::wstring pathText             = path.wstring();
        ViewerOpenContext context{};
        context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName = L"File System";
        context.focusedPath    = pathText.c_str();
        const HRESULT openHr   = viewer->Open(&context);
        Check(SUCCEEDED(openHr), L"ViewerText opens the Save As data-safety fixture", success);
        if (FAILED(openHr))
        {
            return false;
        }

        const bool opened = PumpUntil(
            [&]() noexcept
        {
            viewerWindow = FindNewVisibleWindowByClass(kViewerTextWindowClassName, existingWindows);
            return viewerWindow != nullptr;
        },
            5000ms);
        Check(opened, L"ViewerText Save As data-safety window becomes visible", success);
        if (! opened || ! viewerWindow)
        {
            return false;
        }

        WndMsg::ViewerTextDebugSnapshot snapshot{};
        const bool loaded = WaitForViewerTextSnapshot(viewerWindow,
                                                      [&](const WndMsg::ViewerTextDebugSnapshot& value) noexcept
        {
            if (value.isLoading || value.textStreamActive != expectStream || value.renderCount == 0u)
            {
                return false;
            }
            return expectStream || std::wstring_view(value.textPreview).find(L"ViewerText transactional Save As source") != std::wstring_view::npos;
        },
                                                      15000ms,
                                                      &snapshot);
        Check(loaded,
              expectStream ? L"ViewerText enters streamed mode for the oversized Save As fixture"
                           : L"ViewerText fully loads the Save As source before test dispatch",
              success);
        return loaded;
    };

    const auto invokeSave = [&](const std::filesystem::path& path,
                                UINT encodingSelection,
                                WndMsg::ViewerTextDebugSaveFault fault = WndMsg::ViewerTextDebugSaveFault::None,
                                bool simulateLoading                   = false) noexcept -> HRESULT
    {
        WndMsg::ViewerTextDebugSaveRequest request{};
        request.destinationPath   = path.c_str();
        request.encodingSelection = encodingSelection;
        request.fault             = fault;
        request.simulateLoading   = simulateLoading;
        const LRESULT handled     = SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSaveAs, 0u, reinterpret_cast<LPARAM>(&request));
        Check(handled == TRUE, L"ViewerText Debug Save As seam handles the request", success);
        return request.result;
    };

    const auto checkBytes = [&](const std::filesystem::path& path, const std::vector<std::byte>& expected, std::wstring_view description) noexcept
    {
        const auto actual = ReadBinaryFile(path);
        Check(actual.has_value() && actual.value() == expected, description, success);
    };
    const auto checkNoTemps = [&]() noexcept
    {
        std::error_code countError;
        const size_t count = CountViewerTextSaveTemps(tempDir, countError);
        Check(! countError && count == 0u, L"ViewerText Save As leaves no sibling transaction temp behind", success);
    };
    const auto resetDestination = [&]() { static_cast<void>(WriteUtf8TextFile(destinationPath, kDestinationText)); };

    if (! openViewer(sourcePath, nullptr, false))
    {
        closeViewer();
        return false;
    }

    const HRESULT samePathHr = invokeSave(sourcePath, IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL);
    Check(SUCCEEDED(samePathHr),
          std::format(L"ViewerText keep-original Save As succeeds when source and destination are the same path (hr=0x{:08X})",
                      static_cast<unsigned long>(samePathHr)),
          success);
    checkBytes(sourcePath, sourceBytes.value(), L"ViewerText same-path keep-original Save As preserves every source byte");
    checkNoTemps();

    struct FailureCase
    {
        std::wstring_view name;
        UINT encodingSelection;
        WndMsg::ViewerTextDebugSaveFault fault;
    };
    constexpr std::array failureCases{
        FailureCase{L"source-open", IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL, WndMsg::ViewerTextDebugSaveFault::SourceOpen},
        FailureCase{L"source-read", IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL, WndMsg::ViewerTextDebugSaveFault::SourceRead},
        FailureCase{L"encode", IDM_VIEWER_ENCODING_SAVE_UTF8, WndMsg::ViewerTextDebugSaveFault::Encode},
        FailureCase{L"write", IDM_VIEWER_ENCODING_SAVE_UTF8, WndMsg::ViewerTextDebugSaveFault::Write},
        FailureCase{L"flush", IDM_VIEWER_ENCODING_SAVE_UTF8, WndMsg::ViewerTextDebugSaveFault::Flush},
        FailureCase{L"commit", IDM_VIEWER_ENCODING_SAVE_UTF8, WndMsg::ViewerTextDebugSaveFault::Commit},
    };
    for (const FailureCase& failureCase : failureCases)
    {
        resetDestination();
        const HRESULT failureHr = invokeSave(destinationPath, failureCase.encodingSelection, failureCase.fault);
        Check(FAILED(failureHr), std::format(L"ViewerText injected {} Save As fault fails the transaction", failureCase.name), success);
        checkBytes(sourcePath, sourceBytes.value(), std::format(L"ViewerText injected {} Save As fault preserves source bytes", failureCase.name));
        checkBytes(destinationPath,
                   destinationBytes.value(),
                   std::format(L"ViewerText injected {} Save As fault preserves pre-existing destination bytes", failureCase.name));
        checkNoTemps();
    }

    resetDestination();
    fileSystem.EnableShortRead(static_cast<uint64_t>(sourceBytes.value().size() - 1u));
    const HRESULT shortReadHr = invokeSave(destinationPath, IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL);
    fileSystem.DisableShortRead();
    Check(shortReadHr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF),
          L"ViewerText rejects a successful-but-short source read instead of committing truncated bytes",
          success);
    checkBytes(sourcePath, sourceBytes.value(), L"ViewerText short-read rejection preserves source bytes");
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerText short-read rejection preserves pre-existing destination bytes");
    checkNoTemps();

    resetDestination();
    const HRESULT keepOriginalHr = invokeSave(destinationPath, IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL);
    Check(SUCCEEDED(keepOriginalHr), L"ViewerText transactional keep-original Save As succeeds", success);
    checkBytes(destinationPath, sourceBytes.value(), L"ViewerText successful keep-original Save As copies every source byte");
    checkNoTemps();

    resetDestination();
    const HRESULT reencodeHr = invokeSave(destinationPath, IDM_VIEWER_ENCODING_SAVE_UTF8);
    Check(SUCCEEDED(reencodeHr), L"ViewerText transactional UTF-8 re-encode Save As succeeds", success);
    checkBytes(destinationPath, sourceBytes.value(), L"ViewerText successful UTF-8 re-encode Save As writes the complete document");
    checkNoTemps();

    resetDestination();
    const HRESULT loadingHr = invokeSave(destinationPath, IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL, WndMsg::ViewerTextDebugSaveFault::None, true);
    Check(loadingHr == HRESULT_FROM_WIN32(ERROR_BUSY), L"ViewerText refuses Save As while loading before destination mutation", success);
    checkBytes(sourcePath, sourceBytes.value(), L"ViewerText loading refusal preserves source bytes");
    checkBytes(destinationPath, destinationBytes.value(), L"ViewerText loading refusal preserves pre-existing destination bytes");
    checkNoTemps();

    closeViewer();

    std::string streamText;
    constexpr size_t kStreamFixtureBytes   = (2u * 1024u * 1024u) + 257u;
    constexpr std::string_view kStreamLine = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef\n";
    streamText.reserve(kStreamFixtureBytes + kStreamLine.size());
    while (streamText.size() < kStreamFixtureBytes)
    {
        streamText.append(kStreamLine);
    }
    static_cast<void>(WriteUtf8TextFile(streamPath, streamText));
    static_cast<void>(WriteUtf8TextFile(streamDestinationPath, kDestinationText));
    const auto streamBytes = ReadBinaryFile(streamPath);
    Check(streamBytes.has_value(), L"ViewerText streamed Save As source bytes are readable", success);
    if (! streamBytes.has_value())
    {
        return false;
    }

    constexpr char kStreamConfiguration[] = R"json({"textBufferMiB":1,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1"})json";
    if (! openViewer(streamPath, kStreamConfiguration, true))
    {
        closeViewer();
        return false;
    }

    const HRESULT streamedReencodeHr = invokeSave(streamDestinationPath, IDM_VIEWER_ENCODING_SAVE_UTF8);
    Check(streamedReencodeHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), L"ViewerText refuses streamed re-encode before destination mutation", success);
    checkBytes(streamPath, streamBytes.value(), L"ViewerText streamed re-encode refusal preserves source bytes");
    checkBytes(streamDestinationPath, destinationBytes.value(), L"ViewerText streamed re-encode refusal preserves pre-existing destination bytes");
    checkNoTemps();

    closeViewer();
    emergencyCleanup.release();
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewertext_diff_modes", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerText diff fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }

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
    std::string sparsePaneDiffText;
    sparsePaneDiffText.reserve(150000u);
    sparsePaneDiffText += "diff --git a/viewertext-sparse-pane-old.txt b/viewertext-sparse-pane-new.txt\n";
    sparsePaneDiffText += "--- a/viewertext-sparse-pane-old.txt\n";
    sparsePaneDiffText += "+++ b/viewertext-sparse-pane-new.txt\n";
    sparsePaneDiffText += "@@ -1 +1 @@\n-";
    sparsePaneDiffText.append(70000u, 'i');
    sparsePaneDiffText += "\n+";
    sparsePaneDiffText.append(70000u, '\t');
    sparsePaneDiffText.push_back('\n');
    const std::filesystem::path sparsePaneDiffPath = WriteUtf8TextFile(tempDir / L"viewertext-sparse-pane-checkpoints.diff", sparsePaneDiffText);

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
        static_cast<void>(std::filesystem::remove(sparsePaneDiffPath, cleanupEc));
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

    if (! openViewerForPath(
            sparsePaneDiffPath,
            R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
            viewerWindow))
    {
        return false;
    }

    WndMsg::ViewerTextDebugSnapshot sparsePaneSnapshot{};
    const bool sparsePaneReady = WaitForViewerTextSnapshot(viewerWindow,
                                                           [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
    {
        return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff && snapshot.diffParsedAvailable &&
               snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.paneLocalSideBySideLayout &&
               snapshot.textSparseWrapActive && ! snapshot.textVisualLineCountExact && snapshot.builtLogicalLineCount > 0u;
    },
                                                           10000ms,
                                                           &sparsePaneSnapshot);
    Check(sparsePaneReady, L"ViewerText activates sparse side-by-side layout for a single huge unequal-width diff row", success);
    const HWND sparsePaneTextView = FindWindowExW(viewerWindow, nullptr, kViewerTextViewWindowClassName, nullptr);
    Check(sparsePaneTextView != nullptr, L"ViewerText sparse pane checkpoint fixture exposes its text surface", success);
    bool foundSparseSplitRow = false;
    WndMsg::ViewerTextDebugSnapshot sparseSplitStart{};
    if (sparsePaneReady && sparsePaneTextView)
    {
        const size_t logicalProbeCount = std::min<size_t>(sparsePaneSnapshot.builtLogicalLineCount, 32u);
        for (size_t logicalLine = 0u; logicalLine < logicalProbeCount; ++logicalLine)
        {
            if (SendMessageW(viewerWindow, WndMsg::kViewerTextDebugClickTextLogicalLine, static_cast<WPARAM>(logicalLine), 0u) == FALSE)
            {
                continue;
            }
            WndMsg::ViewerTextDebugSnapshot candidate{};
            if (TryGetViewerTextDebugSnapshot(viewerWindow, candidate) && candidate.topVisibleLogicalLine == logicalLine &&
                candidate.topVisibleRightPaneColumnStart > candidate.topVisibleLeftPaneColumnStart)
            {
                sparseSplitStart    = candidate;
                foundSparseSplitRow = true;
                break;
            }
        }
    }
    Check(foundSparseSplitRow, L"ViewerText can address the first sparse split-pane checkpoint directly", success);
    if (foundSparseSplitRow && sparsePaneTextView)
    {
        static_cast<void>(SendMessageW(sparsePaneTextView, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0u), 0u));
        WndMsg::ViewerTextDebugSnapshot sparseSplitNext{};
        const bool advancedBothPanes = TryGetViewerTextDebugSnapshot(viewerWindow, sparseSplitNext) &&
                                       sparseSplitNext.topVisibleLogicalLine == sparseSplitStart.topVisibleLogicalLine &&
                                       sparseSplitNext.topVisibleLeftPaneColumnStart > sparseSplitStart.topVisibleLeftPaneColumnStart &&
                                       sparseSplitNext.topVisibleRightPaneColumnStart > sparseSplitStart.topVisibleRightPaneColumnStart;
        const uint32_t leftDelta     = advancedBothPanes ? sparseSplitNext.topVisibleLeftPaneColumnStart - sparseSplitStart.topVisibleLeftPaneColumnStart : 0u;
        const uint32_t rightDelta = advancedBothPanes ? sparseSplitNext.topVisibleRightPaneColumnStart - sparseSplitStart.topVisibleRightPaneColumnStart : 0u;
        Check(advancedBothPanes && leftDelta != rightDelta, L"ViewerText advances unequal sparse diff panes with independent DirectWrite checkpoints", success);

        static_cast<void>(SendMessageW(sparsePaneTextView, WM_VSCROLL, MAKEWPARAM(SB_LINEUP, 0u), 0u));
        WndMsg::ViewerTextDebugSnapshot sparseSplitRoundTrip{};
        Check(TryGetViewerTextDebugSnapshot(viewerWindow, sparseSplitRoundTrip) &&
                  sparseSplitRoundTrip.topVisibleLogicalLine == sparseSplitStart.topVisibleLogicalLine &&
                  sparseSplitRoundTrip.topVisibleLeftPaneColumnStart == sparseSplitStart.topVisibleLeftPaneColumnStart &&
                  sparseSplitRoundTrip.topVisibleRightPaneColumnStart == sparseSplitStart.topVisibleRightPaneColumnStart,
              L"ViewerText sparse split-pane line-down then line-up restores both pane checkpoints exactly",
              success);
    }

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

enum class ViewerSpaceProviderFault : uint8_t
{
    None = 0,
    MisalignedNextOffset,
    TruncatedUsedBytes,
    CountMismatch,
    AllocatedTooSmall,
    AllocatedTooLarge,
};

struct ViewerSpaceProviderEntry final
{
    std::wstring name;
    uint64_t bytes      = 0u;
    DWORD attributes    = 0u;
    unsigned long index = 0u;
};

struct ViewerSpaceProviderDirectory final
{
    std::vector<ViewerSpaceProviderEntry> entries;
    ViewerSpaceProviderFault fault = ViewerSpaceProviderFault::None;
    HRESULT readResult             = S_OK;
};

class ViewerSpaceSyntheticFilesInformation final : public IFilesInformation
{
public:
    ViewerSpaceSyntheticFilesInformation(std::vector<std::byte> buffer, unsigned long usedBytes, unsigned long allocatedBytes, unsigned long count) noexcept
        : _buffer(std::move(buffer)),
          _usedBytes(usedBytes),
          _allocatedBytes(allocatedBytes),
          _count(count)
    {
    }
    ~ViewerSpaceSyntheticFilesInformation() = default;

    ViewerSpaceSyntheticFilesInformation(const ViewerSpaceSyntheticFilesInformation&)            = delete;
    ViewerSpaceSyntheticFilesInformation& operator=(const ViewerSpaceSyntheticFilesInformation&) = delete;
    ViewerSpaceSyntheticFilesInformation(ViewerSpaceSyntheticFilesInformation&&)                 = delete;
    ViewerSpaceSyntheticFilesInformation& operator=(ViewerSpaceSyntheticFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (! object)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (riid != __uuidof(IUnknown) && riid != __uuidof(IFilesInformation))
        {
            return E_NOINTERFACE;
        }
        *object = static_cast<IFilesInformation*>(this);
        AddRef();
        return S_OK;
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

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** fileInfo) noexcept override
    {
        if (! fileInfo)
        {
            return E_POINTER;
        }
        *fileInfo = _buffer.empty() ? nullptr : reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* size) noexcept override
    {
        if (! size)
        {
            return E_POINTER;
        }
        *size = _usedBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* size) noexcept override
    {
        if (! size)
        {
            return E_POINTER;
        }
        *size = _allocatedBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* count) noexcept override
    {
        if (! count)
        {
            return E_POINTER;
        }
        *count = _count;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long /*index*/, FileInfo** entry) noexcept override
    {
        if (! entry)
        {
            return E_POINTER;
        }
        *entry = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

private:
    std::atomic_ulong _refCount{1u};
    std::vector<std::byte> _buffer;
    unsigned long _usedBytes      = 0u;
    unsigned long _allocatedBytes = 0u;
    unsigned long _count          = 0u;
};

class ViewerSpaceSyntheticProvider final
{
public:
    ViewerSpaceSyntheticProvider()  = default;
    ~ViewerSpaceSyntheticProvider() = default;

    ViewerSpaceSyntheticProvider(const ViewerSpaceSyntheticProvider&)            = delete;
    ViewerSpaceSyntheticProvider& operator=(const ViewerSpaceSyntheticProvider&) = delete;
    ViewerSpaceSyntheticProvider(ViewerSpaceSyntheticProvider&&)                 = delete;
    ViewerSpaceSyntheticProvider& operator=(ViewerSpaceSyntheticProvider&&)      = delete;

    void AddDirectory(std::wstring path, ViewerSpaceProviderDirectory directory)
    {
        std::scoped_lock lock(_mutex);
        _directories.insert_or_assign(std::move(path), std::move(directory));
    }

    void AddAlias(std::wstring aliasPath, std::wstring targetPath)
    {
        std::scoped_lock lock(_mutex);
        _aliases.insert_or_assign(std::move(aliasPath), std::move(targetPath));
    }

    void UseNonWin32Metadata() noexcept
    {
        _useNonWin32Metadata = true;
    }

    [[nodiscard]] bool UsesNonWin32Metadata() const noexcept
    {
        return _useNonWin32Metadata;
    }

    [[nodiscard]] uint64_t ReadCount(std::wstring_view path) const noexcept
    {
        std::scoped_lock lock(_mutex);
        const auto it = _readCounts.find(std::wstring(path));
        return it == _readCounts.end() ? 0u : it->second;
    }

    static HRESULT ReadCallback(void* context, const wchar_t* path, IFilesInformation** filesInformation) noexcept
    {
        if (! context || ! path || ! filesInformation)
        {
            return E_POINTER;
        }
        return static_cast<ViewerSpaceSyntheticProvider*>(context)->Read(path, filesInformation);
    }

private:
    static size_t EntrySize(std::wstring_view name) noexcept
    {
        const size_t nameBytes = name.size() * sizeof(wchar_t);
        size_t nameStorage     = 0u;
        size_t unaligned       = 0u;
        size_t aligned         = 0u;
        if (! ViewerSpaceScan::TryAddSize(nameBytes, sizeof(wchar_t), nameStorage) ||
            ! ViewerSpaceScan::TryAddSize(offsetof(FileInfo, FileName), nameStorage, unaligned) ||
            ! ViewerSpaceScan::TryAlignUp(unaligned, alignof(FileInfo), aligned))
        {
            return 0u;
        }
        return aligned;
    }

    HRESULT Read(std::wstring_view path, IFilesInformation** filesInformation) noexcept
    {
        *filesInformation = nullptr;
        ViewerSpaceProviderDirectory directory;
        {
            std::scoped_lock lock(_mutex);
            const std::wstring requestedPath(path);
            auto it = _directories.find(requestedPath);
            if (it == _directories.end())
            {
                const auto aliasIt = _aliases.find(requestedPath);
                if (aliasIt != _aliases.end())
                {
                    it = _directories.find(aliasIt->second);
                }
            }
            if (it == _directories.end())
            {
                return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
            }
            directory                  = it->second;
            _readCounts[requestedPath] = ViewerSpaceScan::SaturatingAdd(_readCounts[requestedPath], uint64_t{1u});
        }
        if (FAILED(directory.readResult))
        {
            return directory.readResult;
        }

        size_t totalBytes = 0u;
        for (const ViewerSpaceProviderEntry& source : directory.entries)
        {
            const size_t entrySize = EntrySize(source.name);
            if (entrySize == 0u || ! ViewerSpaceScan::TryAddSize(totalBytes, entrySize, totalBytes) || totalBytes > (std::numeric_limits<unsigned long>::max)())
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
        }

        std::vector<std::byte> buffer(totalBytes, std::byte{0});
        size_t offset = 0u;
        for (size_t index = 0u; index < directory.entries.size(); ++index)
        {
            const ViewerSpaceProviderEntry& source = directory.entries[index];
            const size_t entrySize                 = EntrySize(source.name);
            auto* entry                            = reinterpret_cast<FileInfo*>(buffer.data() + offset);
            entry->FileIndex                       = source.index;
            entry->FileAttributes                  = source.attributes;
            entry->EndOfFile      = source.bytes > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()) ? (std::numeric_limits<__int64>::max)()
                                                                                                                : static_cast<__int64>(source.bytes);
            entry->AllocationSize = entry->EndOfFile;
            entry->FileNameSize   = static_cast<unsigned long>(source.name.size() * sizeof(wchar_t));
            if (! source.name.empty())
            {
                std::memcpy(entry->FileName, source.name.data(), entry->FileNameSize);
            }
            entry->FileName[source.name.size()] = L'\0';
            entry->NextEntryOffset              = index + 1u < directory.entries.size() ? static_cast<unsigned long>(entrySize) : 0u;
            offset += entrySize;
        }

        unsigned long usedBytes      = static_cast<unsigned long>(totalBytes);
        unsigned long allocatedBytes = usedBytes;
        unsigned long count          = static_cast<unsigned long>(directory.entries.size());
        switch (directory.fault)
        {
            case ViewerSpaceProviderFault::MisalignedNextOffset:
                if (directory.entries.size() >= 2u)
                {
                    auto* first = reinterpret_cast<FileInfo*>(buffer.data());
                    first->NextEntryOffset += 2u;
                }
                break;
            case ViewerSpaceProviderFault::TruncatedUsedBytes:
                usedBytes = static_cast<unsigned long>(std::min<size_t>(buffer.size(), offsetof(FileInfo, FileName) - 1u));
                break;
            case ViewerSpaceProviderFault::CountMismatch: count = ViewerSpaceScan::SaturatingAdd(count, uint32_t{1u}); break;
            case ViewerSpaceProviderFault::AllocatedTooSmall: allocatedBytes = usedBytes == 0u ? 0u : usedBytes - 1u; break;
            case ViewerSpaceProviderFault::AllocatedTooLarge: allocatedBytes = (64u * 1024u) + 1u; break;
            case ViewerSpaceProviderFault::None:
            default: break;
        }

        auto result       = std::make_unique<ViewerSpaceSyntheticFilesInformation>(std::move(buffer), usedBytes, allocatedBytes, count);
        *filesInformation = result.release();
        return S_OK;
    }

    mutable std::mutex _mutex;
    std::unordered_map<std::wstring, ViewerSpaceProviderDirectory> _directories;
    std::unordered_map<std::wstring, std::wstring> _aliases;
    mutable std::unordered_map<std::wstring, uint64_t> _readCounts;
    bool _useNonWin32Metadata = false;
};

[[nodiscard]] bool TestViewerSpaceScanPolicyHelpers() noexcept
{
    bool success                           = true;
    ViewerSpaceScan::ResourcePolicy policy = ViewerSpaceScan::kProductionResourcePolicy;
    policy.maxProviderBufferBytes          = 1024u;
    policy.maxProviderEntriesPerFolder     = 4u;
    policy.maxProviderNameBytes            = 64u;

    ViewerSpaceSyntheticProvider provider;
    ViewerSpaceProviderDirectory validDirectory;
    validDirectory.entries.push_back(ViewerSpaceProviderEntry{L"folder", 0u, FILE_ATTRIBUTE_DIRECTORY, 1u});
    validDirectory.entries.push_back(ViewerSpaceProviderEntry{L"file.bin", 42u, FILE_ATTRIBUTE_NORMAL, 2u});
    provider.AddDirectory(L"V:\\helper", std::move(validDirectory));

    BuiltinFileSystemStub fileSystem;
    fileSystem.SetDirectoryReadCallback(&ViewerSpaceSyntheticProvider::ReadCallback, &provider);
    wil::com_ptr<IFilesInformation> files;
    const HRESULT readHr = fileSystem.ReadDirectoryInfo(L"V:\\helper", files.put());
    Check(SUCCEEDED(readHr) && files != nullptr, L"ViewerSpace helper fixture creates a provider buffer", success);
    if (FAILED(readHr) || ! files)
    {
        return false;
    }

    FileInfo* buffer             = nullptr;
    unsigned long usedBytes      = 0u;
    unsigned long allocatedBytes = 0u;
    unsigned long count          = 0u;
    const HRESULT bufferHr       = files->GetBuffer(&buffer);
    const HRESULT usedHr         = files->GetBufferSize(&usedBytes);
    const HRESULT allocatedHr    = files->GetAllocatedSize(&allocatedBytes);
    const HRESULT countHr        = files->GetCount(&count);
    Check(SUCCEEDED(bufferHr) && SUCCEEDED(usedHr) && SUCCEEDED(allocatedHr) && SUCCEEDED(countHr),
          L"ViewerSpace helper fixture exposes provider metadata",
          success);

    using ViewerSpaceScan::ValidationError;
    Check(ViewerSpaceScan::ValidateProviderBufferContract(buffer, usedBytes, allocatedBytes, count, policy) == ValidationError::None,
          L"ViewerSpace provider contract accepts a bounded aligned buffer",
          success);

    const std::span<const std::byte> bytes(reinterpret_cast<const std::byte*>(buffer), usedBytes);
    ViewerSpaceScan::ProviderEntryView first{};
    const ValidationError firstError = ViewerSpaceScan::ValidateProviderEntry(bytes, 0u, policy, first);
    Check(firstError == ValidationError::None && first.name == L"folder" && ! first.isLast,
          L"ViewerSpace provider parser accepts the first aligned record",
          success);
    ViewerSpaceScan::ProviderEntryView second{};
    const ValidationError secondError = firstError == ValidationError::None ? ViewerSpaceScan::ValidateProviderEntry(bytes, first.nextOffset, policy, second)
                                                                            : ValidationError::TruncatedHeader;
    Check(secondError == ValidationError::None && second.name == L"file.bin" && second.isLast,
          L"ViewerSpace provider parser reaches the exact terminal record",
          success);
    Check(secondError == ValidationError::None && ViewerSpaceScan::ValidateProviderTerminalExtent(bytes, first.nextOffset, second) == ValidationError::None,
          L"ViewerSpace provider parser reconciles the terminal record extent with used bytes",
          success);

    Check(ViewerSpaceScan::ValidateProviderBufferContract(buffer, usedBytes, usedBytes - 1u, count, policy) == ValidationError::UsedExceedsAllocated,
          L"ViewerSpace provider contract rejects used bytes beyond allocation",
          success);
    Check(ViewerSpaceScan::ValidateProviderBufferContract(buffer, usedBytes, allocatedBytes, policy.maxProviderEntriesPerFolder + 1u, policy) ==
              ValidationError::EntryCountLimit,
          L"ViewerSpace provider contract rejects excessive advertised counts",
          success);
    Check(ViewerSpaceScan::ValidateProviderBufferContract(buffer, usedBytes, policy.maxProviderBufferBytes + 1u, count, policy) ==
              ValidationError::AllocatedBufferLimit,
          L"ViewerSpace provider contract rejects excessive advertised allocation",
          success);
    Check(ViewerSpaceScan::ValidateProviderChildName(L"safe-name") == ValidationError::None &&
              ViewerSpaceScan::ValidateProviderChildName(L".") == ValidationError::UnsafeName &&
              ViewerSpaceScan::ValidateProviderChildName(L"..") == ValidationError::UnsafeName &&
              ViewerSpaceScan::ValidateProviderChildName(L"cycle\\..") == ValidationError::UnsafeName &&
              ViewerSpaceScan::ValidateProviderChildName(std::wstring_view(L"bad\0name", 8u)) == ValidationError::UnsafeName,
          L"ViewerSpace child-name policy rejects path injection, cycles, and embedded NULs",
          success);
    Check(ViewerSpaceScan::SaturatingAdd((std::numeric_limits<uint64_t>::max)() - 2u, uint64_t{9u}) == (std::numeric_limits<uint64_t>::max)(),
          L"ViewerSpace accounting saturates instead of wrapping",
          success);
    Check(ViewerSpaceScan::kProductionResourcePolicy.maxChildArenaSlots >= ViewerSpaceScan::kProductionResourcePolicy.maxChildReferences * 4u,
          L"ViewerSpace child arena covers geometric live blocks plus every predecessor block at the child-reference ceiling",
          success);

    ViewerSpaceScan::ProviderEntryView misaligned{};
    Check(bytes.size() > 1u && ViewerSpaceScan::ValidateProviderEntry(bytes.subspan(1u), 0u, policy, misaligned) == ValidationError::MisalignedEntry,
          L"ViewerSpace provider parser rejects a misaligned record base",
          success);
    return success;
}

[[nodiscard]] bool TestViewerSpaceBoundedHostileProviderScanning() noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerSpace bounded scan executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerSpace.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerSpace.dll is present for bounded scan validation", success);
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
    Check(pluginModule.is_valid(), L"ViewerSpace.dll loads for bounded scan validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSpace bounded scan validation resolves the factory", success);
    if (! createFn)
    {
        return false;
    }

    constexpr wchar_t kSmallPolicyEnv[]    = L"REDSALAMANDER_VIEWERSPACE_FORCE_SMALL_SCAN_POLICY";
    const std::wstring previousSmallPolicy = GetEnvironmentString(kSmallPolicyEnv);
    auto restoreEnvironment                = wil::scope_exit([&]() noexcept
    { static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, previousSmallPolicy.empty() ? nullptr : previousSmallPolicy.c_str())); });
    static_cast<void>(SetEnvironmentVariableW(kSmallPolicyEnv, L"1"));

    const auto runScenario = [&](ViewerSpaceSyntheticProvider& provider,
                                 std::wstring rootPath,
                                 std::wstring_view label,
                                 WndMsg::ViewerSpacePerfScanState expectedState,
                                 auto&& inspect) noexcept
    {
        BuiltinFileSystemStub fileSystem;
        if (provider.UsesNonWin32Metadata())
        {
            fileSystem.UseSyntheticFileSystemMetadata();
        }
        fileSystem.SetDirectoryReadCallback(&ViewerSpaceSyntheticProvider::ReadCallback, &provider);

        wil::com_ptr<IViewer> viewer;
        const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSpacePluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, std::format(L"{} creates a viewer", label), success);
        if (FAILED(createHr) || ! viewer)
        {
            return;
        }

        wil::com_ptr<IInformations> information;
        const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), information.put_void());
        Check(SUCCEEDED(infoHr) && information != nullptr, std::format(L"{} exposes configuration", label), success);
        if (SUCCEEDED(infoHr) && information)
        {
            const HRESULT configHr = information->SetConfiguration(
                R"json({"topFilesPerDirectory":16,"scanThreads":1,"maxConcurrentScansPerVolume":1,"cacheEnabled":false,"cacheMaxEntries":0})json");
            Check(SUCCEEDED(configHr), std::format(L"{} accepts deterministic scan configuration", label), success);
        }

        ViewerOpenContext context{};
        context.fileSystem     = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName = L"Synthetic";
        context.focusedPath    = rootPath.c_str();

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSpaceWindowClassName);
        const HRESULT openHr                    = viewer->Open(&context);
        Check(SUCCEEDED(openHr), std::format(L"{} opens", label), success);
        if (FAILED(openHr))
        {
            return;
        }

        HWND viewerWindow = nullptr;
        const bool opened = PumpUntil(
            [&]() noexcept
        {
            viewerWindow = FindNewVisibleWindowByClass(kViewerSpaceWindowClassName, existingWindows);
            return viewerWindow != nullptr;
        },
            8000ms);
        Check(opened, std::format(L"{} shows a viewer window", label), success);
        if (! opened || ! viewerWindow)
        {
            static_cast<void>(viewer->Close());
            return;
        }

        WndMsg::ViewerSpacePerfDebugSnapshot snapshot{};
        const bool settled = PumpUntil(
            [&]() noexcept
        {
            snapshot = {};
            if (SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetPerfSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) == FALSE)
            {
                return false;
            }
            return snapshot.scanState == expectedState && snapshot.pendingQueueCount == 0u && snapshot.modelTraversedDirectories > 0u &&
                   (expectedState != WndMsg::ViewerSpacePerfScanState::Done || snapshot.drawItemCount > 0u);
        },
            10'000ms);
        Check(settled, std::format(L"{} reaches a stable terminal model", label), success);
        if (! settled)
        {
            std::wcout << std::format(
                L"[INFO] {} final snapshot: state={}, pending={}, validation={}, realDirs={}, traversed={}, rootBytes={}, folders={}, files={}\n",
                label,
                static_cast<uint32_t>(snapshot.scanState),
                snapshot.pendingQueueCount,
                snapshot.modelValidationError,
                snapshot.realDirectoryCount,
                snapshot.modelTraversedDirectories,
                snapshot.rootTotalBytes,
                snapshot.scannedFolders,
                snapshot.scannedFiles);
        }
        inspect(snapshot, viewerWindow);

        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), std::format(L"{} closes", label), success);
        Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), std::format(L"{} destroys its window", label), success);
    };

    ViewerSpaceSyntheticProvider wideProvider;
    ViewerSpaceProviderDirectory wideRoot;
    for (uint32_t index = 0u; index < 6u; ++index)
    {
        const std::wstring name = std::format(L"d{}", index);
        wideRoot.entries.push_back(ViewerSpaceProviderEntry{name, 0u, FILE_ATTRIBUTE_DIRECTORY, index + 1u});
        ViewerSpaceProviderDirectory child;
        child.entries.push_back(ViewerSpaceProviderEntry{L"payload.bin", static_cast<uint64_t>(index + 1u) * 10u, FILE_ATTRIBUTE_NORMAL, 1u});
        wideProvider.AddDirectory(std::format(L"V:\\wide\\{}", name), std::move(child));
    }
    wideProvider.AddDirectory(L"V:\\wide", std::move(wideRoot));
    runScenario(wideProvider,
                L"V:\\wide",
                L"ViewerSpace tiny-cap wide tree",
                WndMsg::ViewerSpacePerfScanState::Done,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND window) noexcept
    {
        Check(snapshot.rootTotalBytes == 210u && snapshot.scannedFolders == 6u && snapshot.scannedFiles == 6u,
              L"Wide capped scan preserves exact rolled-up totals and counts",
              success);
        Check(snapshot.realDirectoryCount == 4u && snapshot.fileCandidateCount == 3u && snapshot.aggregateFolders == 3u && snapshot.aggregateFiles == 3u &&
                  snapshot.aggregateBytes == 150u,
              L"Wide capped scan itemizes only the tiny policy and aggregates the omitted subtrees exactly",
              success);
        Check(snapshot.childReferenceCount <= 8u && snapshot.childArenaSlots == 10u && snapshot.childArenaFreeSlots == 0u &&
                  snapshot.childArenaSlots <= snapshot.modelChildArenaSlotLimit,
              L"ViewerSpace reuses the retired child block and keeps arena slots bounded",
              success);
        Check(snapshot.modelCappedDirectories == 3u && snapshot.modelCappedFiles == 3u && snapshot.modelRejectedEntries == 0u &&
                  snapshot.modelRetainedNameBytes <= 512u,
              L"Wide capped scan exposes deterministic capped/resource accounting",
              success);
        Check(snapshot.drawItemCount > 0u && IsWindow(window) != FALSE, L"Wide capped scan leaves a stable renderable UI", success);
    });

    ViewerSpaceSyntheticProvider deepProvider;
    std::wstring deepPath = L"V:\\deep";
    for (uint32_t depth = 0u; depth < 6u; ++depth)
    {
        ViewerSpaceProviderDirectory directory;
        const std::wstring childName = std::format(L"level{}", depth);
        directory.entries.push_back(ViewerSpaceProviderEntry{childName, 0u, FILE_ATTRIBUTE_DIRECTORY, depth + 1u});
        deepProvider.AddDirectory(deepPath, std::move(directory));
        deepPath.append(L"\\");
        deepPath.append(childName);
    }
    deepProvider.AddDirectory(deepPath, ViewerSpaceProviderDirectory{});
    runScenario(deepProvider,
                L"V:\\deep",
                L"ViewerSpace tiny-cap deep tree",
                WndMsg::ViewerSpacePerfScanState::Error,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
    {
        Check(snapshot.modelValidationError == static_cast<uint32_t>(ViewerSpaceScan::ValidationError::DepthLimit) && snapshot.realDirectoryCount <= 4u &&
                  snapshot.childArenaSlots <= snapshot.modelChildArenaSlotLimit,
              L"Deep provider traversal stops at the deterministic depth cap without unbounded model growth",
              success);
    });

    const auto runRejectedRoot =
        [&](std::wstring root, ViewerSpaceProviderDirectory directory, ViewerSpaceScan::ValidationError expectedError, std::wstring_view label) noexcept
    {
        ViewerSpaceSyntheticProvider provider;
        provider.AddDirectory(root, std::move(directory));
        runScenario(provider,
                    std::move(root),
                    label,
                    WndMsg::ViewerSpacePerfScanState::Error,
                    [&](const auto& snapshot, HWND) noexcept
        {
            Check(snapshot.modelValidationError == static_cast<uint32_t>(expectedError) && snapshot.modelRejectedEntries > 0u &&
                      snapshot.realDirectoryCount <= snapshot.modelRetainedDirectoryLimit,
                  std::format(L"{} rejects safely with bounded model state", label),
                  success);
        });
    };

    ViewerSpaceProviderDirectory cycleDirectory;
    cycleDirectory.entries.push_back(ViewerSpaceProviderEntry{L"loop\\..", 0u, FILE_ATTRIBUTE_DIRECTORY, 1u});
    runRejectedRoot(L"V:\\cycle", std::move(cycleDirectory), ViewerSpaceScan::ValidationError::UnsafeName, L"ViewerSpace cycle/path injection");

    ViewerSpaceSyntheticProvider aliasCycleProvider;
    aliasCycleProvider.UseNonWin32Metadata();
    ViewerSpaceProviderDirectory aliasCycleRoot;
    aliasCycleRoot.entries.push_back(ViewerSpaceProviderEntry{L"branch", 0u, FILE_ATTRIBUTE_DIRECTORY, 77u});
    aliasCycleProvider.AddDirectory(L"V:\\alias-cycle", std::move(aliasCycleRoot));
    ViewerSpaceProviderDirectory aliasCycleBranch;
    aliasCycleBranch.entries.push_back(ViewerSpaceProviderEntry{L"back", 0u, FILE_ATTRIBUTE_DIRECTORY, 77u});
    aliasCycleProvider.AddDirectory(L"V:\\alias-cycle\\branch", std::move(aliasCycleBranch));
    aliasCycleProvider.AddAlias(L"V:\\alias-cycle\\branch\\back", L"V:\\alias-cycle");
    runScenario(aliasCycleProvider,
                L"V:\\alias-cycle",
                L"ViewerSpace non-reparse provider alias cycle",
                WndMsg::ViewerSpacePerfScanState::Error,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
    {
        Check(snapshot.modelValidationError == static_cast<uint32_t>(ViewerSpaceScan::ValidationError::AncestorCycle) && snapshot.modelRejectedEntries > 0u &&
                  aliasCycleProvider.ReadCount(L"V:\\alias-cycle\\branch\\back") == 0u,
              L"ViewerSpace rejects a repeated nonzero provider identity before descending through an ancestor alias",
              success);
    });

    ViewerSpaceProviderDirectory misalignedDirectory;
    misalignedDirectory.entries.push_back(ViewerSpaceProviderEntry{L"one.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    misalignedDirectory.entries.push_back(ViewerSpaceProviderEntry{L"two.bin", 2u, FILE_ATTRIBUTE_NORMAL, 2u});
    misalignedDirectory.fault = ViewerSpaceProviderFault::MisalignedNextOffset;
    runRejectedRoot(
        L"V:\\misaligned", std::move(misalignedDirectory), ViewerSpaceScan::ValidationError::InvalidNextOffset, L"ViewerSpace misaligned provider buffer");

    ViewerSpaceProviderDirectory countDirectory;
    countDirectory.entries.push_back(ViewerSpaceProviderEntry{L"only.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    countDirectory.fault = ViewerSpaceProviderFault::CountMismatch;
    runRejectedRoot(L"V:\\count", std::move(countDirectory), ViewerSpaceScan::ValidationError::CountMismatch, L"ViewerSpace count mismatch");

    ViewerSpaceProviderDirectory truncatedDirectory;
    truncatedDirectory.entries.push_back(ViewerSpaceProviderEntry{L"truncated.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    truncatedDirectory.fault = ViewerSpaceProviderFault::TruncatedUsedBytes;
    runRejectedRoot(
        L"V:\\truncated", std::move(truncatedDirectory), ViewerSpaceScan::ValidationError::TruncatedHeader, L"ViewerSpace truncated used-buffer bytes");

    ViewerSpaceProviderDirectory allocatedTooSmallDirectory;
    allocatedTooSmallDirectory.entries.push_back(ViewerSpaceProviderEntry{L"allocation.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    allocatedTooSmallDirectory.fault = ViewerSpaceProviderFault::AllocatedTooSmall;
    runRejectedRoot(L"V:\\allocated-too-small",
                    std::move(allocatedTooSmallDirectory),
                    ViewerSpaceScan::ValidationError::UsedExceedsAllocated,
                    L"ViewerSpace allocated-size contract");

    ViewerSpaceProviderDirectory allocatedTooLargeDirectory;
    allocatedTooLargeDirectory.entries.push_back(ViewerSpaceProviderEntry{L"allocation-cap.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    allocatedTooLargeDirectory.fault = ViewerSpaceProviderFault::AllocatedTooLarge;
    runRejectedRoot(L"V:\\allocated-too-large",
                    std::move(allocatedTooLargeDirectory),
                    ViewerSpaceScan::ValidationError::AllocatedBufferLimit,
                    L"ViewerSpace allocated-buffer cap");

    ViewerSpaceProviderDirectory duplicateNameDirectory;
    duplicateNameDirectory.entries.push_back(ViewerSpaceProviderEntry{L"Report.bin", 1u, FILE_ATTRIBUTE_NORMAL, 1u});
    duplicateNameDirectory.entries.push_back(ViewerSpaceProviderEntry{L"report.BIN", 2u, FILE_ATTRIBUTE_NORMAL, 2u});
    runRejectedRoot(L"V:\\duplicate-name",
                    std::move(duplicateNameDirectory),
                    ViewerSpaceScan::ValidationError::DuplicateName,
                    L"ViewerSpace Win32-folded duplicate logical name");

    ViewerSpaceProviderDirectory duplicateDirectory;
    duplicateDirectory.entries.push_back(ViewerSpaceProviderEntry{L"one.bin", 1u, FILE_ATTRIBUTE_NORMAL, 7u});
    duplicateDirectory.entries.push_back(ViewerSpaceProviderEntry{L"two.bin", 2u, FILE_ATTRIBUTE_NORMAL, 7u});
    runRejectedRoot(L"V:\\duplicate", std::move(duplicateDirectory), ViewerSpaceScan::ValidationError::DuplicateId, L"ViewerSpace duplicate provider id");

    constexpr wchar_t kNextModelIdEnv[]    = L"REDSALAMANDER_VIEWERSPACE_TEST_NEXT_MODEL_ID";
    const std::wstring previousNextModelId = GetEnvironmentString(kNextModelIdEnv);
    static_cast<void>(SetEnvironmentVariableW(kNextModelIdEnv, L"1073741823"));
    {
        auto restoreNextModelId = wil::scope_exit([&]() noexcept
        { static_cast<void>(SetEnvironmentVariableW(kNextModelIdEnv, previousNextModelId.empty() ? nullptr : previousNextModelId.c_str())); });
        ViewerSpaceSyntheticProvider exhaustedProvider;
        ViewerSpaceProviderDirectory exhaustedDirectory;
        exhaustedDirectory.entries.push_back(ViewerSpaceProviderEntry{L"child-a", 0u, FILE_ATTRIBUTE_DIRECTORY, 1u});
        exhaustedDirectory.entries.push_back(ViewerSpaceProviderEntry{L"child-b", 0u, FILE_ATTRIBUTE_DIRECTORY, 2u});
        exhaustedProvider.AddDirectory(L"V:\\item-id-exhaustion", std::move(exhaustedDirectory));
        runScenario(exhaustedProvider,
                    L"V:\\item-id-exhaustion",
                    L"ViewerSpace packed item-id exhaustion",
                    WndMsg::ViewerSpacePerfScanState::Error,
                    [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
        {
            Check(snapshot.modelValidationError == static_cast<uint32_t>(ViewerSpaceScan::ValidationError::ItemIdLimit) && snapshot.realDirectoryCount == 1u &&
                      snapshot.childReferenceCount == 0u && snapshot.modelRetainedChildReferences == 0u && snapshot.modelTraversedDirectories == 1u &&
                      snapshot.modelRetainedNameBytes == std::wstring_view(L"V:\\item-id-exhaustion").size() * sizeof(wchar_t),
                  L"ViewerSpace preflights the whole directory-id range before retaining or publishing any child",
                  success);
        });
    }

    constexpr wchar_t kMaxOutstandingEnv[]         = L"REDSALAMANDER_VIEWERSPACE_TEST_MAX_OUTSTANDING_DIRECTORIES";
    constexpr wchar_t kMaxRetainedDirectoriesEnv[] = L"REDSALAMANDER_VIEWERSPACE_TEST_MAX_RETAINED_DIRECTORIES";
    const std::wstring previousMaxOutstanding      = GetEnvironmentString(kMaxOutstandingEnv);
    const std::wstring previousMaxRetained         = GetEnvironmentString(kMaxRetainedDirectoriesEnv);
    static_cast<void>(SetEnvironmentVariableW(kMaxOutstandingEnv, L"4"));
    static_cast<void>(SetEnvironmentVariableW(kMaxRetainedDirectoriesEnv, L"8"));
    {
        auto restoreOutstandingPolicy = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(SetEnvironmentVariableW(kMaxOutstandingEnv, previousMaxOutstanding.empty() ? nullptr : previousMaxOutstanding.c_str()));
            static_cast<void>(SetEnvironmentVariableW(kMaxRetainedDirectoriesEnv, previousMaxRetained.empty() ? nullptr : previousMaxRetained.c_str()));
        });

        ViewerSpaceSyntheticProvider outstandingProvider;
        outstandingProvider.UseNonWin32Metadata();
        ViewerSpaceProviderDirectory outstandingRoot;
        outstandingRoot.entries.push_back(ViewerSpaceProviderEntry{L"d0", 0u, FILE_ATTRIBUTE_DIRECTORY, 1u});
        outstandingRoot.entries.push_back(ViewerSpaceProviderEntry{L"d1", 0u, FILE_ATTRIBUTE_DIRECTORY, 2u});
        outstandingRoot.entries.push_back(ViewerSpaceProviderEntry{L"d2", 0u, FILE_ATTRIBUTE_DIRECTORY, 3u});
        outstandingProvider.AddDirectory(L"V:\\outstanding", std::move(outstandingRoot));
        ViewerSpaceProviderDirectory firstChild;
        firstChild.entries.push_back(ViewerSpaceProviderEntry{L"nested", 0u, FILE_ATTRIBUTE_DIRECTORY, 4u});
        outstandingProvider.AddDirectory(L"V:\\outstanding\\d0", std::move(firstChild));
        outstandingProvider.AddDirectory(L"V:\\outstanding\\d1", ViewerSpaceProviderDirectory{});
        outstandingProvider.AddDirectory(L"V:\\outstanding\\d2", ViewerSpaceProviderDirectory{});
        outstandingProvider.AddDirectory(L"V:\\outstanding\\d0\\nested", ViewerSpaceProviderDirectory{});
        runScenario(outstandingProvider,
                    L"V:\\outstanding",
                    L"ViewerSpace ancestor-inclusive outstanding admission",
                    WndMsg::ViewerSpacePerfScanState::Error,
                    [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
        {
            Check(snapshot.modelValidationError == static_cast<uint32_t>(ViewerSpaceScan::ValidationError::OutstandingLimit) &&
                      snapshot.realDirectoryCount == 4u && snapshot.childReferenceCount == 3u && snapshot.modelRetainedChildReferences == 3u &&
                      snapshot.modelTraversedDirectories == 4u &&
                      snapshot.modelRetainedNameBytes == (std::wstring_view(L"V:\\outstanding").size() + 6u) * sizeof(wchar_t) &&
                      outstandingProvider.ReadCount(L"V:\\outstanding\\d0\\nested") == 0u,
                  L"ViewerSpace counts ancestor completions and publishes no child before outstanding admission",
                  success);
        });
    }

    ViewerSpaceSyntheticProvider overflowProvider;
    ViewerSpaceProviderDirectory overflowDirectory;
    for (uint32_t index = 0u; index < 3u; ++index)
    {
        overflowDirectory.entries.push_back(ViewerSpaceProviderEntry{
            std::format(L"huge{}.bin", index), static_cast<uint64_t>((std::numeric_limits<__int64>::max)()), FILE_ATTRIBUTE_NORMAL, index + 1u});
    }
    overflowProvider.AddDirectory(L"V:\\overflow", std::move(overflowDirectory));
    runScenario(overflowProvider,
                L"V:\\overflow",
                L"ViewerSpace saturating totals",
                WndMsg::ViewerSpacePerfScanState::Done,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
    {
        Check(snapshot.rootTotalBytes == (std::numeric_limits<uint64_t>::max)() && snapshot.scannedFiles == 3u,
              L"ViewerSpace saturates overflowing provider bytes while retaining exact file count",
              success);
    });

    ViewerSpaceSyntheticProvider deniedProvider;
    ViewerSpaceProviderDirectory deniedRoot;
    deniedRoot.entries.push_back(ViewerSpaceProviderEntry{L"denied", 0u, FILE_ATTRIBUTE_DIRECTORY, 1u});
    deniedRoot.entries.push_back(ViewerSpaceProviderEntry{L"visible.bin", 5u, FILE_ATTRIBUTE_NORMAL, 2u});
    deniedProvider.AddDirectory(L"V:\\denied", std::move(deniedRoot));
    ViewerSpaceProviderDirectory deniedChild;
    deniedChild.readResult = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    deniedProvider.AddDirectory(L"V:\\denied\\denied", std::move(deniedChild));
    runScenario(deniedProvider,
                L"V:\\denied",
                L"ViewerSpace access-denied leaf",
                WndMsg::ViewerSpacePerfScanState::Done,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
    {
        Check(snapshot.rootTotalBytes == 5u && snapshot.realDirectoryCount == 2u && snapshot.modelRejectedEntries == 0u &&
                  snapshot.modelValidationError == static_cast<uint32_t>(ViewerSpaceScan::ValidationError::None) &&
                  deniedProvider.ReadCount(L"V:\\denied\\denied") == 1u,
              L"ViewerSpace treats access-denied directories as normal visible zero-sized leaves",
              success);
    });

    ViewerSpaceSyntheticProvider reparseProvider;
    ViewerSpaceProviderDirectory reparseRoot;
    reparseRoot.entries.push_back(ViewerSpaceProviderEntry{L"link", 0u, FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT, 1u});
    reparseRoot.entries.push_back(ViewerSpaceProviderEntry{L"plain.bin", 5u, FILE_ATTRIBUTE_NORMAL, 2u});
    reparseProvider.AddDirectory(L"V:\\reparse", std::move(reparseRoot));
    reparseProvider.AddAlias(L"V:\\reparse\\link", L"V:\\reparse");
    runScenario(reparseProvider,
                L"V:\\reparse",
                L"ViewerSpace reparse non-descent",
                WndMsg::ViewerSpacePerfScanState::Done,
                [&](const WndMsg::ViewerSpacePerfDebugSnapshot& snapshot, HWND) noexcept
    {
        Check(snapshot.rootTotalBytes == 5u && snapshot.realDirectoryCount == 2u && snapshot.scannedFolders == 1u &&
                  reparseProvider.ReadCount(L"V:\\reparse\\link") == 0u,
              L"ViewerSpace keeps an ancestor-alias reparse directory visible and non-descending",
              success);
    });

    return success;
}

struct ViewerSpaceBlockingDirectoryControl final
{
    ViewerSpaceBlockingDirectoryControl()
        : entered(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          release(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          exited(CreateEventW(nullptr, TRUE, FALSE, nullptr))
    {
    }
    ViewerSpaceBlockingDirectoryControl(const ViewerSpaceBlockingDirectoryControl&)            = delete;
    ViewerSpaceBlockingDirectoryControl(ViewerSpaceBlockingDirectoryControl&&)                 = delete;
    ViewerSpaceBlockingDirectoryControl& operator=(const ViewerSpaceBlockingDirectoryControl&) = delete;
    ViewerSpaceBlockingDirectoryControl& operator=(ViewerSpaceBlockingDirectoryControl&&)      = delete;

    [[nodiscard]] bool IsValid() const noexcept
    {
        return entered && release && exited;
    }

    wil::unique_handle entered;
    wil::unique_handle release;
    wil::unique_handle exited;
    std::atomic_uint32_t calls{0u};
};

HRESULT ViewerSpaceBlockingDirectoryRead(void* context, const wchar_t* /*path*/, IFilesInformation** filesInformation) noexcept
{
    auto* control = static_cast<ViewerSpaceBlockingDirectoryControl*>(context);
    if (! control || ! filesInformation)
    {
        return E_POINTER;
    }

    *filesInformation = nullptr;
    control->calls.fetch_add(1u, std::memory_order_acq_rel);
    static_cast<void>(SetEvent(control->entered.get()));
    const auto signalExit  = wil::scope_exit([control]() noexcept { static_cast<void>(SetEvent(control->exited.get())); });
    const DWORD waitResult = WaitForSingleObject(control->release.get(), 30000u);
    if (waitResult != WAIT_OBJECT_0)
    {
        return waitResult == WAIT_TIMEOUT ? HRESULT_FROM_WIN32(ERROR_TIMEOUT) : HRESULT_FROM_WIN32(GetLastError());
    }
    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] bool TestViewerSpaceBlockedProviderCloseAndPostUpdateRace() noexcept
{
#ifndef _DEBUG
    std::wcout << L"[SKIP] ViewerSpace blocked-provider close requires ENABLE_TESTS hooks.\n";
    return true;
#else
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0u && pathLength < modulePath.size(), L"ViewerSpace blocked-provider test executable path resolves", success);
    if (pathLength == 0u || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path buildDir   = std::filesystem::path(modulePath.data()).parent_path();
    const std::filesystem::path pluginPath = buildDir / L"Plugins" / L"ViewerSpace.dll";
    const std::filesystem::path pluginDir  = pluginPath.parent_path();
    Check(std::filesystem::exists(pluginPath), L"ViewerSpace.dll is present for blocked-provider validation", success);
    if (! std::filesystem::exists(pluginPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    const DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(buildDir.c_str());
    const DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(pluginDir.c_str());
    const auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            static_cast<void>(RemoveDllDirectory(pluginCookie));
        }
        if (buildCookie)
        {
            static_cast<void>(RemoveDllDirectory(buildCookie));
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(pluginPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSpace.dll loads for blocked-provider validation", success);
    if (! pluginModule)
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    const FARPROC shutdownProc     = GetProcAddress(pluginModule.get(), "RedSalamanderPluginShutdown");
    const FARPROC canUnloadProc    = GetProcAddress(pluginModule.get(), "RedSalamanderPluginCanUnloadNow");
    RedSalamanderCreateFn createFn = nullptr;
    using PluginShutdownFn         = void(__stdcall*)();
    using PluginCanUnloadFn        = BOOL(__stdcall*)();
    PluginShutdownFn shutdownFn    = nullptr;
    PluginCanUnloadFn canUnloadFn  = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    static_assert(sizeof(shutdownFn) == sizeof(shutdownProc));
    static_assert(sizeof(canUnloadFn) == sizeof(canUnloadProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    std::memcpy(&shutdownFn, &shutdownProc, sizeof(shutdownFn));
    std::memcpy(&canUnloadFn, &canUnloadProc, sizeof(canUnloadFn));
    Check(createFn && shutdownFn && canUnloadFn, L"ViewerSpace exposes factory, shutdown, and runtime unload-gate exports", success);
    if (! createFn || ! shutdownFn || ! canUnloadFn)
    {
        return false;
    }

    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};

    // Force the precise race that the second generation check closes: the old producer passes
    // its fast check, a new scan invalidates and clears under the UI thread, then the producer resumes.
    {
        ViewerSpaceSyntheticProvider provider;
        provider.UseNonWin32Metadata();
        provider.AddDirectory(L"V:\\race", ViewerSpaceProviderDirectory{});
        BuiltinFileSystemStub fileSystem;
        fileSystem.UseSyntheticFileSystemMetadata();
        fileSystem.SetDirectoryReadCallback(&ViewerSpaceSyntheticProvider::ReadCallback, &provider);

        wil::com_ptr<IViewer> viewer;
        Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSpacePluginId, viewer.put_void())) && viewer,
              L"ViewerSpace creates a viewer for the PostUpdate race",
              success);
        wil::com_ptr<IInformations> information;
        if (viewer)
        {
            static_cast<void>(viewer->QueryInterface(__uuidof(IInformations), information.put_void()));
        }
        if (information)
        {
            static_cast<void>(information->SetConfiguration(
                R"json({"topFilesPerDirectory":4,"scanThreads":1,"maxConcurrentScansPerVolume":1,"cacheEnabled":false,"cacheMaxEntries":0})json"));
        }

        const std::wstring rootPath = L"V:\\race";
        ViewerOpenContext context{};
        context.fileSystem                      = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName                  = L"Synthetic";
        context.focusedPath                     = rootPath.c_str();
        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSpaceWindowClassName);
        Check(viewer && SUCCEEDED(viewer->Open(&context)), L"ViewerSpace opens the PostUpdate race fixture", success);

        HWND window = nullptr;
        Check(PumpUntil(
                  [&]() noexcept
        {
            window = FindNewVisibleWindowByClass(kViewerSpaceWindowClassName, existingWindows);
            return window != nullptr;
        },
                  8000ms),
              L"ViewerSpace PostUpdate race window becomes visible",
              success);

        wil::unique_handle pauseEntered(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        wil::unique_handle pauseRelease(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        WndMsg::ViewerSpacePostUpdatePauseDebugControl pauseControl{pauseEntered.get(), pauseRelease.get()};
        Check(window && pauseEntered && pauseRelease &&
                  SendMessageW(window, WndMsg::kViewerSpaceDebugPauseNextPostUpdate, 0u, reinterpret_cast<LPARAM>(&pauseControl)) == TRUE,
              L"ViewerSpace arms the PostUpdate generation-race pause",
              success);
        Check(viewer && SUCCEEDED(viewer->Open(&context)), L"ViewerSpace starts the paused scan generation", success);
        Check(PumpUntil([&]() noexcept { return WaitForSingleObject(pauseEntered.get(), 0u) == WAIT_OBJECT_0; }, 8000ms),
              L"ViewerSpace producer pauses after the fast generation check",
              success);
        Check(viewer && SUCCEEDED(viewer->Open(&context)), L"ViewerSpace invalidates and clears while the stale producer is paused", success);
        static_cast<void>(SetEvent(pauseRelease.get()));

        WndMsg::ViewerSpacePerfDebugSnapshot raceSnapshot{};
        Check(PumpUntil(
                  [&]() noexcept
        {
            raceSnapshot = {};
            return window && SendMessageW(window, WndMsg::kViewerSpaceDebugGetPerfSnapshot, 0u, reinterpret_cast<LPARAM>(&raceSnapshot)) == TRUE &&
                   raceSnapshot.postUpdateInnerGenerationRejects >= 1u && raceSnapshot.scanState == WndMsg::ViewerSpacePerfScanState::Done &&
                   raceSnapshot.pendingQueueCount == 0u && raceSnapshot.pendingQueueBytes == 0u;
        },
                  10000ms),
              L"ViewerSpace inner generation check rejects the resumed stale update with an empty queue",
              success);
        if (viewer)
        {
            static_cast<void>(viewer->Close());
        }
        Check(
            ! window || PumpUntil([&]() noexcept { return IsWindow(window) == FALSE; }, 5000ms), L"ViewerSpace PostUpdate race window closes cleanly", success);
    }

    struct CloseCounter final : IViewerCallback
    {
        CloseCounter()                               = default;
        CloseCounter(const CloseCounter&)            = delete;
        CloseCounter(CloseCounter&&)                 = delete;
        CloseCounter& operator=(const CloseCounter&) = delete;
        CloseCounter& operator=(CloseCounter&&)      = delete;
        HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override
        {
            cookieMatched.store(cookie == this, std::memory_order_release);
            count.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }
        std::atomic_uint32_t count{0u};
        std::atomic_bool cookieMatched{false};
    } closeCounter;

    ViewerSpaceBlockingDirectoryControl block;
    Check(block.IsValid(), L"ViewerSpace blocked-provider events are created", success);
    BuiltinFileSystemStub blockedFileSystem;
    blockedFileSystem.UseSyntheticFileSystemMetadata();
    blockedFileSystem.SetDirectoryReadCallback(&ViewerSpaceBlockingDirectoryRead, &block);

    wil::com_ptr<IViewer> blockedViewer;
    Check(SUCCEEDED(createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSpacePluginId, blockedViewer.put_void())) && blockedViewer,
          L"ViewerSpace creates a viewer for blocked-provider close",
          success);
    if (! blockedViewer)
    {
        return false;
    }
    static_cast<void>(blockedViewer->SetCallback(&closeCounter, &closeCounter));

    const std::wstring blockedPath = L"V:\\blocked";
    ViewerOpenContext blockedContext{};
    blockedContext.fileSystem                      = static_cast<IFileSystem*>(&blockedFileSystem);
    blockedContext.fileSystemName                  = L"Synthetic";
    blockedContext.focusedPath                     = blockedPath.c_str();
    const std::vector<HWND> existingBlockedWindows = CollectVisibleWindowsByClass(kViewerSpaceWindowClassName);
    Check(SUCCEEDED(blockedViewer->Open(&blockedContext)), L"ViewerSpace opens the blocked provider fixture", success);

    HWND blockedWindow = nullptr;
    Check(PumpUntil(
              [&]() noexcept
    {
        blockedWindow = FindNewVisibleWindowByClass(kViewerSpaceWindowClassName, existingBlockedWindows);
        return blockedWindow && WaitForSingleObject(block.entered.get(), 0u) == WAIT_OBJECT_0;
    },
              8000ms),
          L"ViewerSpace reaches blocked ReadDirectoryInfo while its window remains responsive",
          success);

    const auto closeStarted = std::chrono::steady_clock::now();
    const HRESULT closeHr   = blockedViewer->Close();
    const auto closeElapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - closeStarted);
    Check(SUCCEEDED(closeHr) && closeElapsed < 500ms, L"ViewerSpace Close returns within 500 ms while the provider remains blocked", success);
    Check(! blockedWindow || IsWindow(blockedWindow) == FALSE, L"ViewerSpace destroys its HWND before the blocked provider returns", success);
    Check(WaitForSingleObject(block.exited.get(), 0u) == WAIT_TIMEOUT, L"ViewerSpace provider remains blocked after non-waiting Close", success);
    Check(closeCounter.count.load(std::memory_order_acquire) == 1u && closeCounter.cookieMatched.load(std::memory_order_acquire),
          L"ViewerSpace reports ViewerClosed exactly once with the registered cookie",
          success);

    blockedViewer.reset();
    Check(blockedFileSystem.GetReferenceCount() > 1u, L"ViewerSpace detached worker retains the provider and viewer lifetime", success);
    shutdownFn();
    Check(canUnloadFn() == FALSE, L"ViewerSpace runtime unload gate rejects refresh while an abandoned worker exists", success);
    Check(GetModuleHandleW(pluginPath.filename().c_str()) != nullptr, L"ViewerSpace module remains mapped while runtime unload is deferred", success);

    static_cast<void>(SetEvent(block.release.get()));
    Check(PumpUntil([&]() noexcept { return WaitForSingleObject(block.exited.get(), 0u) == WAIT_OBJECT_0; }, 8000ms),
          L"ViewerSpace blocked provider exits after the test releases it",
          success);
    Check(PumpUntil([&]() noexcept { return blockedFileSystem.GetReferenceCount() == 1u; }, 8000ms),
          L"ViewerSpace worker and provider self-references retire exactly once after unblock",
          success);
    Check(canUnloadFn() == FALSE, L"ViewerSpace intentionally quarantines same-path runtime reload for the process after forced abandonment", success);
    Check(closeCounter.count.load(std::memory_order_acquire) == 1u, L"ViewerSpace does not report a second close after worker retirement", success);
    return success;
#endif
}

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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewerspace_window", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerSpace fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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
    CheckViewerSpaceTooltipOverlay(viewerWindow, success);
    CheckEmbeddedViewerHidesStandaloneFileCombo(createFn, kViewerSpacePluginId, context, kViewerSpaceWindowClassName, 0, L"ViewerSpace", success);
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewervlc_focus_hud", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerVLC fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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

    const auto installedVlc                   = FindInstalledVlcForLoaderTest();
    const std::wstring dllDirectoryBeforeOpen = ReadProcessDllDirectoryForTest();
    const HRESULT openHr                      = viewer->Open(&context);
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

    HWND videoWindow                = nullptr;
    HWND hudWindow                  = nullptr;
    const bool visibleSurfacesReady = PumpUntil(
        [&]() noexcept
    {
        videoWindow = FindFirstChildWindowByClass(viewerWindow, kViewerVLCVideoWindowClassName);
        hudWindow   = FindFirstChildWindowByClass(viewerWindow, kViewerVLCHudWindowClassName, false);
        return videoWindow != nullptr && hudWindow != nullptr;
    },
        10000ms);
    if (! visibleSurfacesReady)
    {
        PrintChildWindowDiagnostics(viewerWindow, L"ViewerVLC");
#ifdef _DEBUG
        WndMsg::ViewerVlcDebugSnapshot snapshot{};
        if (SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE)
        {
            std::wcout << std::format(L"[INFO] ViewerVLC snapshot: loadingActive={} loadingVisible={} missingVisible={} hasVideoChild={} "
                                      L"videoChildIsChildWindow={} videoChildParentIsViewer={} hasVolumeMuteButton={} hasVolumeSlider={}\n",
                                      snapshot.loadingActive ? 1 : 0,
                                      snapshot.loadingVisible ? 1 : 0,
                                      snapshot.missingVisible ? 1 : 0,
                                      snapshot.hasVideoChild ? 1 : 0,
                                      snapshot.videoChildIsChildWindow ? 1 : 0,
                                      snapshot.videoChildParentIsViewer ? 1 : 0,
                                      snapshot.hasVolumeMuteButton ? 1 : 0,
                                      snapshot.hasVolumeSlider ? 1 : 0);
        }
#endif
    }
    Check(videoWindow != nullptr && IsWindow(videoWindow) != FALSE, L"ViewerVLC exposes a visible video child that can take initial focus", success);
    Check(hudWindow != nullptr && IsWindow(hudWindow) != FALSE, L"ViewerVLC exposes a HUD child that can take focus", success);
    Check(CountActuallyVisibleChildWindows(viewerWindow) >= 2u, L"ViewerVLC exposes visible video and HUD child surfaces", success);
#ifdef _DEBUG
    if (installedVlc.has_value())
    {
        WndMsg::ViewerVlcDebugSnapshot loadSnapshot{};
        const bool loadReachedTerminalState = PumpUntil(
            [&]() noexcept
        {
            loadSnapshot = {};
            if (SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&loadSnapshot)) == FALSE)
            {
                return false;
            }

            return loadSnapshot.vlcPlayerCreated || (! loadSnapshot.loadingActive && loadSnapshot.missingVisible);
        },
            20'000ms);
        Check(loadReachedTerminalState, L"ViewerVLC reaches a terminal libVLC load state", success);
        Check(loadSnapshot.vlcModuleLoaded, L"ViewerVLC loads libvlc.dll with its sibling dependencies", success);
        Check(loadSnapshot.vlcInstanceLoaded, L"ViewerVLC creates a libVLC instance", success);
        Check(loadSnapshot.vlcPlayerCreated && ! loadSnapshot.missingVisible, L"ViewerVLC starts playback for the valid WAV fixture", success);
        Check(ReadProcessDllDirectoryForTest() == dllDirectoryBeforeOpen, L"ViewerVLC dependency loading leaves the process DLL directory unchanged", success);
    }
    else
    {
        std::wcout << L"[SKIP] ViewerVLC dependency-load/playback proof requires VLC in a standard Program Files location.\n";
    }
#else
    static_cast<void>(installedVlc);
    static_cast<void>(dllDirectoryBeforeOpen);
#endif
    if (videoWindow && hudWindow)
    {
        static_cast<void>(SetFocus(viewerWindow));

        const bool productRoutedVideoFocus = PumpUntil([&]() noexcept { return GetFocus() == videoWindow; }, 2000ms);
        bool videoFocusReady               = productRoutedVideoFocus;
        if (! videoFocusReady)
        {
            std::wcout << L"[INFO] ViewerVLC did not route parent focus into the video surface; using the explicit video-focus fallback for keyboard-contract "
                          L"coverage.\n";
            static_cast<void>(SetFocus(videoWindow));
            videoFocusReady = PumpUntil([&]() noexcept { return GetFocus() == videoWindow; }, 2000ms);
        }
        Check(videoFocusReady, L"ViewerVLC video surface can take keyboard focus", success);
        if (productRoutedVideoFocus)
        {
            std::wcout << L"[INFO] ViewerVLC parent focus routed into the video surface without fallback.\n";
        }

        if (videoFocusReady)
        {
            const LRESULT tabHandled = SendMessageW(videoWindow, WM_KEYDOWN, VK_TAB, 0);
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

    const char* schema     = nullptr;
    const HRESULT schemaHr = informations->GetConfigurationSchema(&schema);
    Check(SUCCEEDED(schemaHr) && schema != nullptr, L"ViewerVLC configuration schema is available", success);
    if (SUCCEEDED(schemaHr) && schema)
    {
        const std::string_view schemaText(schema);
        Check(schemaText.find("\"lastVolumePercent\"") != std::string_view::npos, L"ViewerVLC schema exposes last volume persistence", success);
        Check(schemaText.find("\"muted\"") != std::string_view::npos, L"ViewerVLC schema exposes mute persistence", success);
        Check(schemaText.find("Trusted-user libVLC options") != std::string_view::npos,
              L"ViewerVLC schema identifies extra arguments as trusted-user configuration",
              success);
    }

    constexpr char kSafeOptionsJson[] =
        R"json({"avcodecHw":"d3d11va","videoOutput":"direct3d11","audioOutput":"mmdevice","audioVisualization":"spectrum","extraArgs":"--no-sub-autodetect-file --file-caching=250"})json";
    Check(SUCCEEDED(informations->SetConfiguration(kSafeOptionsJson)), L"ViewerVLC accepts bounded module tokens and ordinary trusted-user options", success);

    static constexpr std::array<std::string_view, 11> kRejectedConfigurations{{
        R"json({"videoOutput":"direct3d11/..\\evil"})json",
        R"json({"audioOutput":"mmdevice --intf=lua"})json",
        "{\"avcodecHw\":\"any\\n--sout=x\"}",
        R"json({"extraArgs":"--plugin-path=C:\\evil"})json",
        R"json({"extraArgs":"--intf=lua"})json",
        R"json({"extraArgs":"--video-filter=scene"})json",
        R"json({"extraArgs":"--no-logfile=C:\\temp\\vlc.log"})json",
        R"json({"extraArgs":"--No-logfile=C:\\temp\\vlc.log"})json",
        R"json({"extraArgs":"--file-logging"})json",
        R"json({"extraArgs":"--pidfile=C:\\temp\\vlc.pid"})json",
        R"json({"extraArgs":"--no-sub-autodetect-file=\"unterminated"})json",
    }};
    for (const std::string_view rejected : kRejectedConfigurations)
    {
        Check(informations->SetConfiguration(std::string(rejected).c_str()) == E_INVALIDARG,
              L"ViewerVLC rejects separator, control, option, and external-module injection",
              success);
    }

    constexpr char kMaxUnsignedJson[] =
        R"json({"fileCachingMs":18446744073709551615,"networkCachingMs":18446744073709551615,"defaultPlaybackRatePercent":18446744073709551615,"lastVolumePercent":18446744073709551615})json";
    Check(SUCCEEDED(informations->SetConfiguration(kMaxUnsignedJson)), L"ViewerVLC clamps maximum unsigned JSON values before narrowing", success);
    const char* clampedJson = nullptr;
    Check(SUCCEEDED(informations->GetConfiguration(&clampedJson)) && clampedJson != nullptr,
          L"ViewerVLC returns the safely clamped unsigned configuration",
          success);
    if (clampedJson)
    {
        const std::string_view clampedText(clampedJson);
        Check(clampedText.find("\"fileCachingMs\":60000") != std::string_view::npos &&
                  clampedText.find("\"networkCachingMs\":60000") != std::string_view::npos &&
                  clampedText.find("\"defaultPlaybackRatePercent\":400") != std::string_view::npos &&
                  clampedText.find("\"lastVolumePercent\":100") != std::string_view::npos,
              L"ViewerVLC clamps each unsigned setting to its documented maximum",
              success);
    }

    constexpr char kSavedVolumeJson[] = R"json({"lastVolumePercent":37,"muted":true})json";
    const HRESULT setHr               = informations->SetConfiguration(kSavedVolumeJson);
    Check(SUCCEEDED(setHr), L"ViewerVLC accepts persisted volume and mute configuration", success);

    const char* savedJson = nullptr;
    const HRESULT getHr   = informations->GetConfiguration(&savedJson);
    Check(SUCCEEDED(getHr) && savedJson != nullptr, L"ViewerVLC returns normalized persisted configuration", success);
    if (SUCCEEDED(getHr) && savedJson)
    {
        const std::string_view savedText(savedJson);
        Check(savedText.find("\"lastVolumePercent\":37") != std::string_view::npos, L"ViewerVLC keeps the last volume in its persisted configuration", success);
        Check(savedText.find("\"muted\":true") != std::string_view::npos, L"ViewerVLC keeps mute state in its persisted configuration", success);
    }

    BOOL somethingToSave = FALSE;
    const HRESULT saveHr = informations->SomethingToSave(&somethingToSave);
    Check(SUCCEEDED(saveHr) && somethingToSave != FALSE, L"ViewerVLC marks non-default volume state for persistence", success);
    return success;
}

[[nodiscard]] bool TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts() noexcept
{
#ifndef _DEBUG
    std::wcout << L"[SKIP] ViewerVLC HUD snapshot contract requires ViewerVLC test hooks; Release plugin builds omit ENABLE_TESTS.\n";
    return true;
#else
    bool success = true;

    struct CloseCounter final : IViewerCallback
    {
        CloseCounter()                               = default;
        CloseCounter(const CloseCounter&)            = delete;
        CloseCounter(CloseCounter&&)                 = delete;
        CloseCounter& operator=(const CloseCounter&) = delete;
        CloseCounter& operator=(CloseCounter&&)      = delete;

        HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override
        {
            cookieMatched = cookie == this;
            count.fetch_add(1, std::memory_order_acq_rel);
            return S_OK;
        }

        std::atomic_uint32_t count{0};
        bool cookieMatched = false;
    } closeCounter;

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
    Check(SUCCEEDED(viewer->SetCallback(&closeCounter, &closeCounter)), L"ViewerVLC accepts a lifecycle callback for close-once validation", success);

    wil::com_ptr<IInformations> informations;
    static_cast<void>(viewer->QueryInterface(__uuidof(IInformations), informations.put_void()));
    if (informations)
    {
        static_cast<void>(informations->SetConfiguration(R"json({"autoDetectVlc":false})json"));
    }

    const ViewerTheme rainbowTheme = MakeViewerTextTestTheme(false, true);
    const HRESULT themeHr          = viewer->SetTheme(&rainbowTheme);
    Check(SUCCEEDED(themeHr), L"ViewerVLC accepts a rainbow viewer theme for loading-overlay validation", success);

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerVLCWindowClassName);

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(L"viewervlc_hud_snapshot", ec);
    Check(! ec && ! tempDir.empty(), L"ViewerVLC HUD fixture TestSandbox root is available", success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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

    const bool initialLoadTerminal = PumpUntil(
        [&]() noexcept
    {
        snapshot = {};
        return SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE && ! snapshot.loadingActive;
    },
        5000ms);
    Check(initialLoadTerminal, L"ViewerVLC reaches a terminal state before async fault injection", success);

    WndMsg::ViewerVlcDebugAsyncControl asyncControl{};
    asyncControl.failNextLoadSubmit = true;
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)) != FALSE,
          L"ViewerVLC debug hook arms a deterministic loader-submit failure",
          success);
    const uint64_t rejectedBefore     = snapshot.loadQueueRejected;
    const auto submitFailureStartedAt = std::chrono::steady_clock::now();
    const HRESULT submitFailureOpenHr = viewer->Open(&context);
    const uint64_t submitFailureReturnUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - submitFailureStartedAt).count());
    Check(SUCCEEDED(submitFailureOpenHr), L"ViewerVLC accepts an open whose loader submission is fault-injected", success);
    Check(submitFailureReturnUs < 100'000u, L"ViewerVLC loader-submit failure never runs loader work synchronously on the UI thread", success);
    const bool submitFailureTerminal = PumpUntil(
        [&]() noexcept
    {
        snapshot = {};
        return SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE && ! snapshot.loadingActive &&
               snapshot.missingVisible && snapshot.loadQueueRejected > rejectedBefore;
    },
        5000ms);
    Check(submitFailureTerminal, L"ViewerVLC posts a terminal UI error after loader-submit failure", success);

    const uint64_t loadPostFallbacksBefore  = snapshot.loadPostFallbacks;
    const uint64_t postFailuresBefore       = snapshot.asyncResultPostFailures;
    asyncControl                            = {};
    asyncControl.failNextLoadCompletionPost = true;
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)) != FALSE,
          L"ViewerVLC debug hook arms a deterministic load-completion payload-post failure",
          success);
    Check(SUCCEEDED(viewer->Open(&context)), L"ViewerVLC accepts an open whose completion payload post is fault-injected", success);
    const bool loadPostFallbackTerminal = PumpUntil(
        [&]() noexcept
    {
        snapshot = {};
        return SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE && ! snapshot.loadingActive &&
               snapshot.missingVisible && snapshot.loadPostFallbacks > loadPostFallbacksBefore && snapshot.asyncResultPostFailures > postFailuresBefore;
    },
        5000ms);
    Check(loadPostFallbackTerminal, L"ViewerVLC load-completion post failure reaches the identity-bound terminal UI fallback", success);

    const uint64_t staleBefore    = snapshot.staleLoadResults;
    const uint64_t acceptedBefore = snapshot.loadQueueAccepted;
    asyncControl                  = {};
    asyncControl.loadDelayMs      = 250;
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)));
    Check(SUCCEEDED(viewer->Open(&context)) && SUCCEEDED(viewer->Open(&context)), L"ViewerVLC accepts rapid superseding delayed opens", success);
    const bool staleRequestRejected = PumpUntil(
        [&]() noexcept
    {
        snapshot = {};
        return SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE && ! snapshot.loadingActive &&
               snapshot.loadQueueAccepted >= acceptedBefore + 2u && snapshot.staleLoadResults > staleBefore;
    },
        8000ms);
    Check(staleRequestRejected, L"ViewerVLC rejects a delayed stale load result by request and window identity", success);
    Check(! snapshot.vlcPlayerCreated, L"ViewerVLC never attaches or plays a stale delayed load result", success);

    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugForceLoadingVisible, 0, 0) != FALSE, L"ViewerVLC debug can force delayed loading overlay", success);
    snapshot = {};
    static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
    Check(snapshot.loadingActive && snapshot.loadingVisible, L"ViewerVLC shows a loading overlay once VLC init exceeds the delay", success);
    Check(snapshot.loadingSpinnerDotCount >= 12, L"ViewerVLC loading spinner uses a denser, easier-to-see dot ring", success);
    Check(snapshot.loadingSpinnerOrbitPx >= 18, L"ViewerVLC loading spinner is larger than the old compact wheel", success);
    Check(snapshot.loadingSpinnerActiveDotRadiusPx > snapshot.loadingSpinnerDotRadiusPx,
          L"ViewerVLC loading spinner gives the active dot extra visual weight",
          success);
    Check(snapshot.loadingSpinnerUsesRainbow, L"ViewerVLC loading spinner honors rainbow viewer themes", success);
    Check(snapshot.loadingSpinnerFirstDotArgb != snapshot.loadingSpinnerSecondDotArgb,
          L"ViewerVLC rainbow loading spinner uses multiple dot colors instead of one accent color",
          success);

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

    WndMsg::ViewerVlcDebugStopDelay stopDelay{};
    stopDelay.delayMs = 300;
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetStopDelay, 0, reinterpret_cast<LPARAM>(&stopDelay)) != FALSE,
          L"ViewerVLC debug hook arms delayed stop/release",
          success);
    const uint64_t cleanupPostFallbacksBefore      = snapshot.cleanupPostFallbacks;
    const uint64_t closePostFailuresBefore         = snapshot.asyncResultPostFailures;
    const uint64_t cleanupAllocationFailuresBefore = snapshot.cleanupAllocationFailures;
    const uint64_t cleanupSubmitFailuresBefore     = snapshot.cleanupSubmitFailures;
    asyncControl                                   = {};
    asyncControl.failNextCloseCompletionPost       = true;
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)) != FALSE,
          L"ViewerVLC debug hook arms a deterministic close-completion payload-post failure",
          success);
    const HWND retainedVideoWindow = FindFirstChildWindowByClass(viewerWindow, kViewerVLCVideoWindowClassName);
    const auto closeStartedAt      = std::chrono::steady_clock::now();
    const HRESULT closeHr          = viewer->Close();
    const uint64_t closeReturnUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - closeStartedAt).count());
    Check(SUCCEEDED(closeHr), L"ViewerVLC HUD contract window close succeeds", success);
    Check(closeReturnUs < 100'000u, L"ViewerVLC Close returns without waiting for delayed stop/release", success);
    Check(IsWindow(viewerWindow) != FALSE && IsWindow(retainedVideoWindow) != FALSE,
          L"ViewerVLC retains parent and video HWND identities while stop/release is outstanding",
          success);
    Check(IsWindowVisible(viewerWindow) == FALSE && IsWindowVisible(retainedVideoWindow) == FALSE,
          L"ViewerVLC hides retained parent and video surfaces during asynchronous close",
          success);
    snapshot = {};
    Check(SendMessageW(viewerWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE && snapshot.closePending,
          L"ViewerVLC reports an identity-bound close pending state",
          success);
    uint64_t previousWindowIdentity = snapshot.windowIdentity;
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"ViewerVLC destroys retained HWNDs after asynchronous cleanup completion",
          success);
    Check(closeCounter.count.load(std::memory_order_acquire) == 1u && closeCounter.cookieMatched,
          L"ViewerVLC notifies close exactly once with the registered cookie",
          success);

    constexpr size_t kRapidCycles = 6u;
    for (size_t cycle = 0; cycle < kRapidCycles; ++cycle)
    {
        const std::vector<HWND> existingCycleWindows = CollectVisibleWindowsByClass(kViewerVLCWindowClassName);
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerVLC rapidly reopens after asynchronous close", success);
        HWND cycleWindow       = nullptr;
        const bool cycleOpened = PumpUntil(
            [&]() noexcept
        {
            cycleWindow = FindNewVisibleWindowByClass(kViewerVLCWindowClassName, existingCycleWindows);
            return cycleWindow != nullptr;
        },
            5000ms);
        Check(cycleOpened, L"ViewerVLC rapid-reopen cycle creates a visible window", success);
        if (! cycleOpened || ! cycleWindow)
        {
            break;
        }

        snapshot = {};
        static_cast<void>(SendMessageW(cycleWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
        if (cycle == 0)
        {
            Check(snapshot.cleanupPostFallbacks > cleanupPostFallbacksBefore && snapshot.asyncResultPostFailures > closePostFailuresBefore,
                  L"ViewerVLC close-completion post failure drained through the identity-bound UI fallback",
                  success);
        }
        if (cycle == 1)
        {
            Check(snapshot.cleanupAllocationFailures > cleanupAllocationFailuresBefore,
                  L"ViewerVLC pre-created dispatcher drains cleanup after an injected allocation failure",
                  success);
        }
        if (cycle == 2)
        {
            Check(snapshot.cleanupSubmitFailures > cleanupSubmitFailuresBefore,
                  L"ViewerVLC pre-created dispatcher drains cleanup after an injected submit failure",
                  success);
        }
        Check(snapshot.windowIdentity != 0 && snapshot.windowIdentity != previousWindowIdentity,
              L"ViewerVLC assigns a new identity to every reopened window even under HWND reuse pressure",
              success);
        previousWindowIdentity = snapshot.windowIdentity;

        asyncControl                           = {};
        asyncControl.loadDelayMs               = 75;
        asyncControl.failNextCleanupAllocation = cycle == 0;
        asyncControl.failNextCleanupSubmit     = cycle == 1;
        static_cast<void>(SendMessageW(cycleWindow, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)));
        stopDelay.delayMs = 40;
        static_cast<void>(SendMessageW(cycleWindow, WndMsg::kViewerVlcDebugSetStopDelay, 0, reinterpret_cast<LPARAM>(&stopDelay)));
        Check(SUCCEEDED(viewer->Open(&context)), L"ViewerVLC queues a delayed superseding load in a rapid-reopen cycle", success);
        const auto rapidCloseStartedAt = std::chrono::steady_clock::now();
        const HRESULT rapidCloseHr     = viewer->Close();
        const uint64_t rapidCloseReturnUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - rapidCloseStartedAt).count());
        Check(SUCCEEDED(rapidCloseHr), L"ViewerVLC closes immediately after a delayed rapid-reopen load", success);
        Check(rapidCloseReturnUs < 100'000u, L"ViewerVLC cleanup allocation/submit fault injection never blocks Close on stop/release", success);
        Check(IsWindow(cycleWindow) != FALSE && IsWindowVisible(cycleWindow) == FALSE,
              L"ViewerVLC retains and hides the rapid-reopen window until delayed work is drained",
              success);
        Check(PumpUntil([&]() noexcept { return IsWindow(cycleWindow) == FALSE; }, 5000ms),
              L"ViewerVLC rapid-reopen window closes after identity-bound completions drain",
              success);
        Check(closeCounter.count.load(std::memory_order_acquire) == cycle + 2u,
              L"ViewerVLC emits one close notification per accepted rapid-reopen generation",
              success);
    }

    wil::unique_hwnd forcedParent(CreateWindowExW(0,
                                                  L"STATIC",
                                                  L"ViewerVLC forced-parent teardown host",
                                                  WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                                                  CW_USEDEFAULT,
                                                  CW_USEDEFAULT,
                                                  640,
                                                  480,
                                                  nullptr,
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    Check(forcedParent.is_valid(), L"ViewerVLC forced-parent teardown host is created", success);
    HWND forcedChild = nullptr;
    if (forcedParent)
    {
        ViewerOpenContext embeddedContext = context;
        embeddedContext.ownerWindow       = forcedParent.get();
        embeddedContext.flags             = VIEWER_OPEN_FLAG_EMBEDDED;
        Check(SUCCEEDED(viewer->Open(&embeddedContext)), L"ViewerVLC opens under the forced-destroy parent", success);
        Check(PumpUntil(
                  [&]() noexcept
        {
            forcedChild = FindFirstChildWindowByClass(forcedParent.get(), kViewerVLCWindowClassName);
            return forcedChild != nullptr;
        },
                  5000ms),
              L"ViewerVLC forced-parent child becomes visible",
              success);
    }

    if (forcedChild)
    {
        Check(PumpUntil(
                  [&]() noexcept
        {
            WndMsg::ViewerVlcDebugSnapshot forcedSnapshot{};
            return SendMessageW(forcedChild, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&forcedSnapshot)) != FALSE &&
                   ! forcedSnapshot.loadingActive;
        },
                  5000ms),
              L"ViewerVLC forced-parent child drains loader work before teardown",
              success);

        // Keep a wide separation between the injected cleanup delay and the
        // parent-destroy budget so ordinary scheduler jitter cannot mimic a
        // synchronous cleanup wait.
        stopDelay.delayMs = 1000u;
        static_cast<void>(SendMessageW(forcedChild, WndMsg::kViewerVlcDebugSetStopDelay, 0, reinterpret_cast<LPARAM>(&stopDelay)));
        asyncControl                           = {};
        asyncControl.failNextCleanupAllocation = true;
        static_cast<void>(SendMessageW(forcedChild, WndMsg::kViewerVlcDebugSetAsyncControl, 0, reinterpret_cast<LPARAM>(&asyncControl)));
        static_cast<void>(viewer->SetCallback(nullptr, nullptr));

        const auto parentDestroyStartedAt = std::chrono::steady_clock::now();
        forcedParent.reset();
        const uint64_t parentDestroyReturnUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - parentDestroyStartedAt).count());
        Check(parentDestroyReturnUs < 250'000u, L"ViewerVLC forced parent destruction does not wait for delayed cleanup", success);
        Check(IsWindow(forcedChild) == FALSE, L"ViewerVLC child HWND is destroyed immediately with its forced parent", success);
    }

    const bool expectDispatcherPin = forcedChild != nullptr;
    informations.reset();
    viewer.reset();
    pluginModule.reset();
    if (expectDispatcherPin)
    {
        Check(GetModuleHandleW(L"ViewerVLC.dll") != nullptr, L"ViewerVLC dispatcher pin keeps the plugin mapped after forced parent destruction", success);
    }
    Check(PumpUntil([]() noexcept { return GetModuleHandleW(L"ViewerVLC.dll") == nullptr; }, 5000ms),
          L"ViewerVLC unloads after the HWND-independent dispatcher completes stop/release",
          success);
    return success;
#endif
}

void CheckViewerSpaceTooltipOverlay(HWND viewerWindow, bool& success) noexcept
{
    static_cast<void>(UpdateWindow(viewerWindow));
    const std::vector<HWND> existingTooltipWindows = CollectWindowsByClass(kNativeTooltipWindowClassName, false);

    WndMsg::ViewerSpaceTooltipDebugSnapshot before{};
    const LRESULT beforeResult = SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetTooltipSnapshot, 0, reinterpret_cast<LPARAM>(&before));
    Check(beforeResult != FALSE, L"ViewerSpace answers the tooltip-overlay debug contract", success);
    if (beforeResult == FALSE)
    {
        return;
    }

    Check(before.hasRenderTarget, L"ViewerSpace has a Direct2D render target before tooltip overlay validation", success);
    Check(before.hasTooltipFormat, L"ViewerSpace has a DirectWrite tooltip format before tooltip overlay validation", success);
    Check(std::abs(before.tooltipMaxWidthDip - 420.0f) <= 0.01f, L"ViewerSpace tooltip overlay uses a 420 DIP maximum text width", success);

    const LRESULT showResult = SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugShowTooltipOverlay, 0, 0);
    Check(showResult != FALSE, L"ViewerSpace debug hook shows the Direct2D tooltip overlay", success);
    if (showResult == FALSE)
    {
        return;
    }

    WndMsg::ViewerSpaceTooltipDebugSnapshot after{};
    const bool painted = PumpUntil(
        [&]() noexcept
    {
        static_cast<void>(UpdateWindow(viewerWindow));
        const LRESULT queryResult = SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetTooltipSnapshot, 0, reinterpret_cast<LPARAM>(&after));
        return queryResult != FALSE && after.tooltipNodeId != 0u && after.tooltipTextLength > 0u && after.tooltipPaintCount > before.tooltipPaintCount;
    },
        5000ms);
    Check(painted, L"ViewerSpace Direct2D tooltip overlay paints after it is shown", success);
    if (painted)
    {
        Check(after.tooltipAnchorXDip > 0.0f && after.tooltipAnchorYDip > 0.0f, L"ViewerSpace tooltip overlay tracks a client DIP anchor", success);
        Check(std::abs(after.tooltipMaxWidthDip - 420.0f) <= 0.01f, L"ViewerSpace tooltip overlay keeps the high-DPI-safe width cap", success);
    }

    const std::vector<HWND> currentTooltipWindows = CollectWindowsByClass(kNativeTooltipWindowClassName, false);
    const auto isExisting                         = [&](HWND hwnd) noexcept
    { return std::find(existingTooltipWindows.begin(), existingTooltipWindows.end(), hwnd) != existingTooltipWindows.end(); };
    const bool createdNativeTooltip = std::find_if(currentTooltipWindows.begin(), currentTooltipWindows.end(), [&](HWND hwnd) noexcept {
        return ! isExisting(hwnd);
    }) != currentTooltipWindows.end();
    Check(! createdNativeTooltip, L"ViewerSpace tooltip overlay does not create a native tooltip window", success);
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

    std::error_code ec;
    const std::filesystem::path tempDir = AcquireViewerPETestSandbox(promptName == L"Find" ? L"viewertext_find_prompt" : L"viewertext_goto_prompt", ec);
    Check(! ec && ! tempDir.empty(), std::format(L"ViewerText {} prompt fixture TestSandbox root is available", promptName), success);
    if (ec || tempDir.empty())
    {
        return false;
    }
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
        success = RunFilteredSelfExecutable(L"TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses", kViewerHarnessDefaultTimeout, success) &&
                  success;
        success = RunFilteredSelfExecutable(L"TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly", kViewerHarnessDefaultTimeout, success) && success;
        success = RunFilteredSelfExecutable(L"TestViewerVlcConfigurationPersistsLastVolumeAndMute", kViewerHarnessDefaultTimeout, success) && success;
#ifdef _DEBUG
        success = RunFilteredSelfExecutable(L"TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts", kViewerHarnessDefaultTimeout, success) && success;
#endif
    }

    return success;
}

[[nodiscard]] bool RunFullSuiteInFreshProcesses() noexcept
{
    bool success = true;

    std::vector<IsolatedViewerTest> isolatedTests{
        {L"TestViewerPEUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout},
        {L"TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead", kViewerHarnessDefaultTimeout},
        {L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout},
        {L"TestViewerWebEscapesScriptBreakoutInGeneratedDocuments", kViewerHarnessDefaultTimeout},
        {L"TestViewerWebSecurityPolicyAndBounds", kViewerHarnessDefaultTimeout},
        {L"TestViewerWebVirtualHtmlUsesPrivateOriginAndEnforcesByteCaps", kViewerHarnessDefaultTimeout},
        {L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout},
        {L"TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert", kViewerHarnessDefaultTimeout},
        {L"TestViewerImgRawMenuOwnershipOrientationAndExifScalarGuards", kViewerHarnessDefaultTimeout},
        {L"TestViewerImgRawResourcePolicyHelpers", kViewerHarnessDefaultTimeout},
        {L"TestViewerImgRawResourceBudgetAndLongPathExport", kViewerHarnessDefaultTimeout},
        {L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo", kViewerHarnessDefaultTimeout},
        {L"TestViewerSpaceScanPolicyHelpers", kViewerHarnessDefaultTimeout},
        {L"TestViewerSpaceBoundedHostileProviderScanning", 120000ms},
        {L"TestViewerSpaceBlockedProviderCloseAndPostUpdateRace", kViewerHarnessDefaultTimeout},
        {L"TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses", kViewerHarnessDefaultTimeout},
        {L"TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly", kViewerHarnessDefaultTimeout},
        {L"TestViewerVlcConfigurationPersistsLastVolumeAndMute", kViewerHarnessDefaultTimeout},
        {L"TestViewerShellComboHostsLongRunOpenCloseStayStable", kViewerShellComboLongRunTimeout},
    };
#ifdef _DEBUG
    isolatedTests.push_back({L"TestViewerWebTransactionalSaveAsAndCloseSafety", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerImgRawLatestWinsExactReaderAndCloseSafety", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerImgRawEmbeddedThumbnailTerminalSequencing", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerVlcHudLoadingWheelSnapshotAndVolumeContracts", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextHexShortReadDropsPhantomTailBytes", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextDecodeAndClipboardSafetyHelpers", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextAsyncOpenAndUtf8HexTerminalContracts", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextSaveAsPreservesDataOnFailures", kViewerHarnessDefaultTimeout});
    isolatedTests.push_back({L"TestViewerTextDiffModesAndPlaceholders", kViewerHarnessDefaultTimeout});
#endif

    for (const IsolatedViewerTest& test : isolatedTests)
    {
        success = RunFilteredSelfExecutable(test.name, test.timeout, success) && success;
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
    if (shouldRun(L"TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead"))
    {
        success = TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead() && success;
    }
    if (shouldRun(L"TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
    if (shouldRun(L"TestViewerWebEscapesScriptBreakoutInGeneratedDocuments"))
    {
        success = TestViewerWebEscapesScriptBreakoutInGeneratedDocuments() && success;
    }
    if (shouldRun(L"TestViewerWebSecurityPolicyAndBounds"))
    {
        success = TestViewerWebSecurityPolicyAndBounds() && success;
    }
    if (shouldRun(L"TestViewerWebVirtualHtmlUsesPrivateOriginAndEnforcesByteCaps"))
    {
        success = TestViewerWebVirtualHtmlUsesPrivateOriginAndEnforcesByteCaps() && success;
    }
#ifdef _DEBUG
    if (shouldRun(L"TestViewerWebTransactionalSaveAsAndCloseSafety"))
    {
        success = TestViewerWebTransactionalSaveAsAndCloseSafety() && success;
    }
#endif
    if (shouldRun(L"TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
    if (shouldRun(L"TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert"))
    {
        success = TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert() && success;
    }
    if (shouldRun(L"TestViewerImgRawMenuOwnershipOrientationAndExifScalarGuards"))
    {
        success = TestViewerImgRawMenuOwnershipOrientationAndExifScalarGuards() && success;
    }
    if (shouldRun(L"TestViewerImgRawResourcePolicyHelpers"))
    {
        success = TestViewerImgRawResourcePolicyHelpers() && success;
    }
    if (shouldRun(L"TestViewerImgRawResourceBudgetAndLongPathExport"))
    {
        success = TestViewerImgRawResourceBudgetAndLongPathExport() && success;
    }
#ifdef _DEBUG
    if (shouldRun(L"TestViewerImgRawLatestWinsExactReaderAndCloseSafety"))
    {
        success = TestViewerImgRawLatestWinsExactReaderAndCloseSafety() && success;
    }
    if (shouldRun(L"TestViewerImgRawEmbeddedThumbnailTerminalSequencing"))
    {
        success = TestViewerImgRawEmbeddedThumbnailTerminalSequencing() && success;
    }
#endif
    if (shouldRun(L"TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo"))
    {
        success = TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo() && success;
    }
#ifdef _DEBUG
    if (shouldRun(L"TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback"))
    {
        success = TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback() && success;
    }
    if (shouldRun(L"TestViewerTextHexShortReadDropsPhantomTailBytes"))
    {
        success = TestViewerTextHexShortReadDropsPhantomTailBytes() && success;
    }
    if (shouldRun(L"TestViewerTextDecodeAndClipboardSafetyHelpers"))
    {
        success = TestViewerTextDecodeAndClipboardSafetyHelpers() && success;
    }
    if (shouldRun(L"TestViewerTextAsyncOpenAndUtf8HexTerminalContracts"))
    {
        success = TestViewerTextAsyncOpenAndUtf8HexTerminalContracts() && success;
    }
    if (shouldRun(L"TestViewerTextSaveAsPreservesDataOnFailures"))
    {
        success = TestViewerTextSaveAsPreservesDataOnFailures() && success;
    }
    if (shouldRun(L"TestViewerTextDiffModesAndPlaceholders"))
    {
        success = TestViewerTextDiffModesAndPlaceholders() && success;
    }
#endif
    if (shouldRun(L"TestViewerSpaceScanPolicyHelpers"))
    {
        success = TestViewerSpaceScanPolicyHelpers() && success;
    }
    if (shouldRun(L"TestViewerSpaceBoundedHostileProviderScanning"))
    {
        success = TestViewerSpaceBoundedHostileProviderScanning() && success;
    }
    if (shouldRun(L"TestViewerSpaceBlockedProviderCloseAndPostUpdateRace"))
    {
        success = TestViewerSpaceBlockedProviderCloseAndPostUpdateRace() && success;
    }
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

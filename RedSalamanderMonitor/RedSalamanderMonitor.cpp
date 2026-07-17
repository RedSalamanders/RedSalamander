#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <filesystem>
#include <format>
#include <fstream>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>
#include <vector>

#include <Windows.h>

#include <bcrypt.h>
#include <commdlg.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <shellapi.h>
#include <shellscalingapi.h>
#pragma comment(lib, "Bcrypt.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shcore.lib")

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <wincodec.h>
#pragma comment(lib, "Windowscodecs.lib")

// Define the ETW provider for RedSalamanderMonitor.exe
// Each executable must have its own provider instance with the same GUID
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h" // Must be included before EtwListener.h for InfoParam definition

#include "ColorTextView.h"
#include "Configuration.h"
#include "DxUi/DxUi.h"
#include "EtwListener.h"
#include "ExceptionHelpers.h" // Shared exception handling utilities
#include "LocalizationManager.h"
#include "MinimumOsVersion.h"
#include "MonitorDiagnostics.h"
#include "MonitorFileReader.h"
#include "RedSalamanderMonitor.h"
#include "SettingsStore.h"
#include "Version.h"
#include "WindowBackdropPolicy.h"
#include "WindowSizing.h"
#include "resource.h"

// Global Variables:
// All globals below are accessed exclusively from the UI thread (message loop).
// The only cross-thread interaction is EtwListener's worker thread calling
// g_colorView.QueueEtwEvent(), which is thread-safe via atomic HWND + critical section.
HINSTANCE g_hInstance = NULL;    // current instance
HWND g_hMainWindow    = nullptr; // monitor top-level window
ColorTextView g_colorView;       // ColorTextView instance for the right panel
wil::unique_hwnd g_hColorView;   // ColorTextView window handle
wil::unique_hwnd g_hToolbar;     // Toolbar window handle
wil::unique_hwnd g_hStatusBar;   // Status bar window handle
RedSalamander::DxUi::WindowHost g_toolbarDxHost;
RedSalamander::DxUi::WindowHost g_statusDxHost;
bool g_showIds            = true;  // Show Process/Thread IDs in output
bool g_alwaysOnTop        = false; // Main window always-on-top flag
bool g_toolbarVisible     = true;  // Toolbar visibility (menu state)
bool g_lineNumbersVisible = true;  // Line numbers menu state
bool g_autoScrollEnabled  = true;  // Auto-scroll menu state
// Auto-scroll state is now managed by ColorTextView (_autoScrollEnabled member)
static std::unique_ptr<EtwListener> g_etwListener; // ETW real-time event listener
Common::Settings::Settings g_settings;

// Filter state: bitmask where each visible message type maps to a dedicated bit.
static uint32_t g_filterMask  = Debug::InfoParam::Type::All; // All visible types enabled by default
static int g_lastFilterPreset = -1;                          // -1 = custom, 0 = Errors Only, 1 = Errors+Warnings, 2 = All, 3 = Errors+Perf+Debug

// Status bar update timer
static constexpr UINT_PTR kStatusBarTimerId      = 100;
static constexpr UINT kStatusBarUpdateIntervalMs = 500; // Update every 500ms
static uint64_t g_lastMessageCount               = 0;   // Track message rate for adaptive refresh

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::BlendColor;
using RedSalamander::DxUi::ColorFromArgb;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::StatusStrip;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::Toolbar;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kAppId[]                                = L"RedSalamanderMonitor";
constexpr wchar_t kWindowId[]                             = L"MonitorWindow";
constexpr wchar_t kDxHostClassName[]                      = L"RedSalamanderMonitor.DxHost";
constexpr std::wstring_view kMonitorChromeSelfTestArg     = L"--chrome-selftest";
constexpr std::wstring_view kMonitorScrollbackSelfTestArg = L"--monitor-scrollback-selftest";
constexpr std::wstring_view kMonitorEtwBurstModePrefix    = L"--monitor-etw-burst-mode=";
constexpr std::wstring_view kMonitorEtwBurstCountPrefix   = L"--monitor-etw-burst-count=";
constexpr std::wstring_view kMonitorEtwBurstSizePrefix    = L"--monitor-etw-burst-size=";
constexpr std::wstring_view kMonitorEtwBurstLatencyMode   = L"latency";
constexpr std::wstring_view kMonitorAreaName              = L"Monitor";
constexpr std::wstring_view kMonitorScenarioName          = L"monitor.chrome.dxui_toolbar_statusstrip";
constexpr UINT kMsgRunMonitorChromeSelfTest               = WM_APP + 0x61C;
constexpr UINT kMsgMonitorFileOpenProgress                 = WM_APP + 0x61D;
constexpr UINT kMsgMonitorFileOpenCompleted                = WM_APP + 0x61E;

#if defined(_DEBUG)
constexpr std::wstring_view kMonitorBuildFlavor = L"Debug";
#else
constexpr std::wstring_view kMonitorBuildFlavor = L"Release";
#endif

constexpr wchar_t kMonitorHelpText[] = L"RedSalamanderMonitor\r\n"
                                       L"\r\n"
                                       L"Usage:\r\n"
                                       L"  RedSalamanderMonitor.exe [options]\r\n"
                                       L"\r\n"
                                       L"Options:\r\n"
                                       L"  -h, --help, /?                 Show this help.\r\n"
                                       L"  --etw                           Enable RedSalamanderMonitor self Info/Perf/Debug ETW diagnostics.\r\n"
                                       L"  --perf                          Write RedSalamanderMonitor perf metrics to the default JSONL path.\r\n"
                                       L"  --perf=PATH                     Write RedSalamanderMonitor perf metrics to a custom JSONL path.\r\n"
                                       L"  --monitor-scrollback-selftest    Run monitor selftest scrollback frame scenario.\r\n"
                                       L"  --monitor-etw-burst-mode=latency Run monitor selftest ETW latency burst scenario.\r\n"
                                       L"  --monitor-etw-burst-count=N      Set latency burst sample count.\r\n"
                                       L"  --monitor-etw-burst-size=N       Set latency burst payload characters.\r\n"
                                       L"\r\n";

HMENU g_viewThemeMenu = nullptr;

constexpr UINT kCustomThemeMenuIdFirst = 32900u;
constexpr UINT kCustomThemeMenuIdLast  = 33099u;

std::unordered_map<UINT, std::wstring> g_customThemeMenuIdToThemeId;
std::unordered_map<std::wstring, UINT> g_customThemeIdToMenuId;
std::vector<Common::Settings::ThemeDefinition> g_fileThemes;

struct MonitorFilterUiEntry final
{
    Debug::InfoParam::Type type;
    UINT stringId;
    UINT menuId;
};

constexpr std::array<MonitorFilterUiEntry, 6> kMonitorFilterEntries = {{
    {Debug::InfoParam::Type::Text, IDS_FILTER_TYPE_TEXT, IDM_FILTER_TEXT},
    {Debug::InfoParam::Type::Error, IDS_FILTER_TYPE_ERROR, IDM_FILTER_ERROR},
    {Debug::InfoParam::Type::Warning, IDS_FILTER_TYPE_WARNING, IDM_FILTER_WARNING},
    {Debug::InfoParam::Type::Info, IDS_FILTER_TYPE_INFO, IDM_FILTER_INFO},
    {Debug::InfoParam::Type::Perf, IDS_FILTER_TYPE_PERF, IDM_FILTER_PERF},
    {Debug::InfoParam::Type::Debug, IDS_FILTER_TYPE_DEBUG, IDM_FILTER_DEBUG},
}};

constexpr uint32_t kMonitorFilterAllMask        = Debug::InfoParam::Type::All;
constexpr uint32_t kMonitorPresetErrorsOnlyMask = Debug::FilterBitForType(Debug::InfoParam::Type::Error);
constexpr uint32_t kMonitorPresetErrorsWarningsMask =
    Debug::FilterBitForType(Debug::InfoParam::Type::Error) | Debug::FilterBitForType(Debug::InfoParam::Type::Warning);
constexpr uint32_t kMonitorPresetErrorsPerfDebugMask = Debug::FilterBitForType(Debug::InfoParam::Type::Error) |
                                                       Debug::FilterBitForType(Debug::InfoParam::Type::Perf) |
                                                       Debug::FilterBitForType(Debug::InfoParam::Type::Debug);

[[nodiscard]] const MonitorFilterUiEntry* FindMonitorFilterEntry(UINT menuId) noexcept
{
    for (const auto& entry : kMonitorFilterEntries)
    {
        if (entry.menuId == menuId)
        {
            return &entry;
        }
    }

    return nullptr;
}

void SyncMonitorFilterMenuChecks(HMENU menu, uint32_t mask) noexcept
{
    if (! menu)
    {
        return;
    }

    const uint32_t clampedMask = mask & kMonitorFilterAllMask;
    for (const auto& entry : kMonitorFilterEntries)
    {
        const uint32_t bit = Debug::FilterBitForType(entry.type);
        CheckMenuItem(menu, entry.menuId, static_cast<UINT>(MF_BYCOMMAND | ((clampedMask & bit) ? MF_CHECKED : MF_UNCHECKED)));
    }
}

Toolbar* g_toolbarRoot         = nullptr;
Button* g_toolbarNewButton     = nullptr;
Button* g_toolbarOpenButton    = nullptr;
Button* g_toolbarSaveButton    = nullptr;
Button* g_toolbarCopyButton    = nullptr;
Toggle* g_toolbarShowIdsToggle = nullptr;
StatusStrip* g_statusStrip     = nullptr;
struct MonitorChromeSelfTestCheck final
{
    std::wstring name;
    bool passed = false;
    std::wstring detail;
};

struct MonitorChromeMetricPresence final
{
    std::wstring_view metric;
    uint64_t count = 0;
};

struct MonitorChromeMetricSummary final
{
    std::wstring_view metric;
    uint64_t count = 0;
    uint64_t p50   = 0;
    uint64_t p95   = 0;
    uint64_t p99   = 0;
    uint64_t max   = 0;
};

struct MonitorEtwBurstOptions final
{
    bool latencyMode    = false;
    size_t count        = 60u;
    size_t payloadChars = 260u;
};

struct MonitorScrollbackSelfTestOptions final
{
    bool enabled = false;
};

struct MonitorChromeSelfTestContext final
{
    bool enabled   = false;
    bool completed = false;
    int exitCode   = 0;
    std::wstring machineHash;
    std::wstring runId;
    std::filesystem::path repoRoot;
    std::filesystem::path runRoot;
    std::filesystem::path tracePath;
    std::filesystem::path resultsPath;
    std::filesystem::path perfPath;
    std::vector<MonitorChromeSelfTestCheck> checks;
};

MonitorChromeSelfTestContext g_monitorChromeSelfTest;
MonitorEtwBurstOptions g_monitorEtwBurstOptions;
MonitorScrollbackSelfTestOptions g_monitorScrollbackSelfTestOptions;

struct MonitorFileOpenProgress final
{
    uint64_t generation = 0u;
    uint64_t bytesRead  = 0u;
    uint64_t totalBytes = 0u;
};

struct MonitorFileOpenCompletion final
{
    RedSalamanderMonitor::MonitorFileReadResult result;
    uint64_t generation = 0u;
    uint64_t durationUs = 0u;
};

std::jthread g_monitorFileOpenThread;
std::atomic<bool> g_monitorFileOpenActive{false};
uint64_t g_monitorFileOpenGeneration = 0u;

constexpr std::array<std::wstring_view, 6> kRequiredMonitorFrameMetrics{{
    L"monitor.frame.total_us",
    L"monitor.frame.present_us",
    L"monitor.frame.append_to_visible_us",
    L"monitor.frame.tail_layout_us",
    L"monitor.frame.mode",
    L"monitor.etw.batch_drain_us",
}};

constexpr std::array<std::wstring_view, 8> kRequiredMonitorEtwBurstLatencyMetrics{{
    L"monitor.etw.batch_drain_us",
    L"monitor.etw.selftest_burst_drain_us",
    L"monitor.frame.append_to_visible_us",
    L"monitor.frame.total_us",
    L"monitor.frame.present_us",
    L"monitor.frame.tail_layout_us",
    L"monitor.etw.queue_depth",
    L"monitor.etw.batch_repost_count",
}};

constexpr std::array<std::wstring_view, 4> kSummarizedMonitorEtwBurstLatencyMetrics{{
    L"monitor.frame.append_to_visible_us",
    L"monitor.etw.batch_drain_us",
    L"monitor.frame.total_us",
    L"monitor.frame.present_us",
}};

constexpr std::array<std::wstring_view, 4> kRequiredMonitorScrollbackMetrics{{
    L"monitor.frame.scrollback_slice_us",
    L"monitor.frame.mode",
    L"monitor.frame.total_us",
    L"monitor.frame.present_us",
}};

constexpr std::array<std::wstring_view, 4> kSummarizedMonitorScrollbackMetrics{{
    L"monitor.frame.scrollback_slice_us",
    L"monitor.frame.mode",
    L"monitor.frame.total_us",
    L"monitor.frame.present_us",
}};

void SyncToolbarState() noexcept;
void UpdateStatusBar();
void LayoutToolbarControls() noexcept;
void AdjustLayout(HWND hWnd);

struct MonitorResolvedTheme final
{
    ColorTextView::Theme textView;
    bool dark         = false;
    bool highContrast = false;
    bool rainbow      = false;
    bool compactMode  = false;
    std::optional<bool> reducedMotionOverride;
    Common::Settings::WindowBackdropMode windowBackdrop = Common::Settings::WindowBackdropMode::Default;
};

struct MonitorChromeMetrics final
{
    float toolbarHeightDip         = 42.0f;
    float statusStripHeightDip     = 24.0f;
    float toolbarPaddingDip        = 8.0f;
    float toolbarButtonHeightDip   = 32.0f;
    float toolbarGapDip            = 6.0f;
    float toolbarSeparatorWidthDip = 12.0f;
    float toolbarMinButtonWidthDip = 64.0f;
    float toolbarMinToggleWidthDip = 108.0f;
    float toolbarLabelCharWidthDip = 8.0f;
    float toolbarTextPaddingDip    = 24.0f;
    float toolbarToggleChromeDip   = 52.0f;
    float statusAutoWidthDip       = 90.0f;
    float statusFilterWidthDip     = 220.0f;
    float statusVisibleWidthDip    = 150.0f;
    float statusTotalWidthDip      = 150.0f;
};

struct ModalMessageDialogState
{
    const wchar_t* caption = nullptr;
    const wchar_t* message = nullptr;
};

struct ModalConfirmDialogState
{
    const wchar_t* caption = nullptr;
    const wchar_t* message = nullptr;
};

INT_PTR CALLBACK ModalMessageDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<ModalMessageDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
        case WM_INITDIALOG:
        {
            state = reinterpret_cast<ModalMessageDialogState*>(lp);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

            if (state)
            {
                if (state->caption && state->caption[0] != L'\0')
                {
                    SetWindowTextW(dlg, state->caption);
                }

                if (state->message && state->message[0] != L'\0')
                {
                    SetDlgItemTextW(dlg, IDC_MODAL_MESSAGE_TEXT, state->message);
                }
            }

            wchar_t okText[64]{};
            const int okLength = LoadStringW(GetModuleHandleW(nullptr), IDS_BTN_OK, okText, static_cast<int>(sizeof(okText) / sizeof(okText[0])));
            if (okLength > 0)
            {
                SetDlgItemTextW(dlg, IDOK, okText);
            }

            return static_cast<INT_PTR>(TRUE);
        }
        case WM_COMMAND:
        {
            const WORD id = LOWORD(wp);
            if (id == IDOK || id == IDCANCEL)
            {
                EndDialog(dlg, id);
                return static_cast<INT_PTR>(TRUE);
            }
            break;
        }
    }
    return static_cast<INT_PTR>(FALSE);
}

INT_PTR CALLBACK ModalConfirmDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<ModalConfirmDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
        case WM_INITDIALOG:
        {
            state = reinterpret_cast<ModalConfirmDialogState*>(lp);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

            if (state)
            {
                if (state->caption && state->caption[0] != L'\0')
                {
                    SetWindowTextW(dlg, state->caption);
                }

                if (state->message && state->message[0] != L'\0')
                {
                    SetDlgItemTextW(dlg, IDC_MODAL_MESSAGE_TEXT, state->message);
                }
            }

            wchar_t yesText[64]{};
            const int yesLength = LoadStringW(GetModuleHandleW(nullptr), IDS_BTN_YES, yesText, static_cast<int>(sizeof(yesText) / sizeof(yesText[0])));
            if (yesLength > 0)
            {
                SetDlgItemTextW(dlg, IDYES, yesText);
            }

            wchar_t noText[64]{};
            const int noLength = LoadStringW(GetModuleHandleW(nullptr), IDS_BTN_NO, noText, static_cast<int>(sizeof(noText) / sizeof(noText[0])));
            if (noLength > 0)
            {
                SetDlgItemTextW(dlg, IDNO, noText);
            }

            SetFocus(GetDlgItem(dlg, IDYES));
            return static_cast<INT_PTR>(FALSE);
        }
        case WM_COMMAND:
        {
            const WORD id = LOWORD(wp);
            if (id == IDYES || id == IDNO)
            {
                EndDialog(dlg, id);
                return static_cast<INT_PTR>(TRUE);
            }

            if (id == IDCANCEL)
            {
                EndDialog(dlg, IDNO);
                return static_cast<INT_PTR>(TRUE);
            }
            break;
        }
        case WM_CLOSE: EndDialog(dlg, IDNO); return static_cast<INT_PTR>(TRUE);
    }
    return static_cast<INT_PTR>(FALSE);
}

void ShowModalMessageDialog(HINSTANCE instance, HWND owner, const wchar_t* caption, const wchar_t* message) noexcept
{
    ModalMessageDialogState state{};
    state.caption = caption;
    state.message = message;

    static_cast<void>(RedSalamander::Win32Callback::DialogBoxParamResourceNoThrow(
        instance, MAKEINTRESOURCEW(IDD_MODAL_MESSAGE), owner, ModalMessageDialogProc, reinterpret_cast<LPARAM>(&state)));
}

bool ShowModalConfirmDialog(HINSTANCE instance, HWND owner, const wchar_t* caption, const wchar_t* message) noexcept
{
    ModalConfirmDialogState state{};
    state.caption = caption;
    state.message = message;

    const INT_PTR result = RedSalamander::Win32Callback::DialogBoxParamResourceNoThrow(
        instance, MAKEINTRESOURCEW(IDD_MODAL_CONFIRM), owner, ModalConfirmDialogProc, reinterpret_cast<LPARAM>(&state));

    return result == IDYES;
}

std::filesystem::path GetThemesDirectory() noexcept
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }

    return std::filesystem::path(modulePath.get()).parent_path() / L"Themes";
}

std::wstring EscapeMenuLabel(std::wstring_view raw)
{
    std::wstring result;
    result.reserve(raw.size());

    for (const wchar_t ch : raw)
    {
        result.push_back(ch);
        if (ch == L'&')
        {
            result.push_back(L'&');
        }
    }

    return result;
}

// Convert a value expressed in DIPs (96-DPI) to physical pixels for a given window
inline int DipsToPx(HWND hwnd, int dip)
{
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    if (hwnd)
    {
        dpi = GetDpiForWindow(hwnd);
    }
    return Common::WindowSizing::DipToPixelRounded(dpi, dip);
}

bool IsProcessElevated() noexcept
{
    wil::unique_handle token;
    if (! ::OpenProcessToken(::GetCurrentProcess(), TOKEN_QUERY, token.put()))
    {
        return false;
    }

    TOKEN_ELEVATION elevation{};
    DWORD size = 0;
    if (! ::GetTokenInformation(token.get(), TokenElevation, &elevation, sizeof(elevation), &size))
    {
        return false;
    }

    return elevation.TokenIsElevated != 0;
}

std::wstring QuoteCommandLineArg(std::wstring_view arg)
{
    if (arg.empty())
    {
        return L"\"\"";
    }

    const bool needsQuotes = arg.find_first_of(L" \t\n\v\"") != std::wstring_view::npos;
    if (! needsQuotes)
    {
        return std::wstring{arg};
    }

    std::wstring result;
    result.reserve(arg.size() + 2);
    result.push_back(L'"');

    size_t backslashCount = 0;
    for (const wchar_t ch : arg)
    {
        if (ch == L'\\')
        {
            ++backslashCount;
            continue;
        }

        if (ch == L'"')
        {
            result.append(backslashCount * 2 + 1, L'\\');
            result.push_back(L'"');
            backslashCount = 0;
            continue;
        }

        if (backslashCount != 0)
        {
            result.append(backslashCount, L'\\');
            backslashCount = 0;
        }

        result.push_back(ch);
    }

    if (backslashCount != 0)
    {
        result.append(backslashCount * 2, L'\\');
    }

    result.push_back(L'"');
    return result;
}

std::wstring BuildRelaunchParameters(std::wstring_view extraArg) noexcept
{
    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));

    std::wstring params;
    bool alreadyHasExtra = false;

    if (argv && argc > 1)
    {
        for (int i = 1; i < argc; ++i)
        {
            if (! extraArg.empty() && extraArg == argv.get()[i])
            {
                alreadyHasExtra = true;
            }

            if (! params.empty())
            {
                params.push_back(L' ');
            }
            params.append(QuoteCommandLineArg(argv.get()[i]));
        }
    }

    if (! extraArg.empty() && ! alreadyHasExtra)
    {
        if (! params.empty())
        {
            params.push_back(L' ');
        }
        params.append(QuoteCommandLineArg(extraArg));
    }

    return params;
}

bool HasCommandLineArg(std::wstring_view arg) noexcept
{
    if (arg.empty())
    {
        return false;
    }

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv || argc <= 1)
    {
        return false;
    }

    for (int i = 1; i < argc; ++i)
    {
        if (arg == argv.get()[i])
        {
            return true;
        }
    }

    return false;
}

bool TryGetCommandLineArgValue(std::wstring_view prefix, std::wstring& value) noexcept
{
    value.clear();
    if (prefix.empty())
    {
        return false;
    }

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv || argc <= 1)
    {
        return false;
    }

    for (int i = 1; i < argc; ++i)
    {
        const wchar_t* arg = argv.get()[i];
        if (! arg)
        {
            continue;
        }

        if (wcsncmp(arg, prefix.data(), prefix.size()) != 0)
        {
            continue;
        }

        value = arg + prefix.size();
        return true;
    }

    return false;
}

std::optional<size_t> TryParseSizeT(std::wstring_view value) noexcept
{
    if (value.empty())
    {
        return std::nullopt;
    }

    uint64_t parsed = 0u;
    for (const wchar_t ch : value)
    {
        if (ch < L'0' || ch > L'9')
        {
            return std::nullopt;
        }

        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        if (parsed > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
        {
            return std::nullopt;
        }
        parsed = (parsed * 10u) + digit;
    }

    if (parsed > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return std::nullopt;
    }

    return static_cast<size_t>(parsed);
}

MonitorEtwBurstOptions ReadMonitorEtwBurstOptions() noexcept
{
    MonitorEtwBurstOptions options;

    std::wstring mode;
    if (! TryGetCommandLineArgValue(kMonitorEtwBurstModePrefix, mode) || mode != kMonitorEtwBurstLatencyMode)
    {
        options.latencyMode = false;
        return options;
    }

    options.latencyMode = true;

    std::wstring value;
    if (TryGetCommandLineArgValue(kMonitorEtwBurstCountPrefix, value))
    {
        if (const std::optional<size_t> parsed = TryParseSizeT(value); parsed.has_value())
        {
            options.count = std::clamp(parsed.value(), size_t{1u}, size_t{2'000u});
        }
    }

    if (TryGetCommandLineArgValue(kMonitorEtwBurstSizePrefix, value))
    {
        if (const std::optional<size_t> parsed = TryParseSizeT(value); parsed.has_value())
        {
            options.payloadChars = std::clamp(parsed.value(), size_t{1u}, size_t{8'192u});
        }
    }

    return options;
}

void WriteMonitorHelpText(HINSTANCE hInstance) noexcept
{
    auto tryWrite = [](HANDLE handle, std::wstring_view msg) -> bool
    {
        if (! handle || handle == INVALID_HANDLE_VALUE)
        {
            return false;
        }

        DWORD mode = 0;
        if (GetConsoleMode(handle, &mode) != FALSE)
        {
            DWORD written = 0;
            return WriteConsoleW(handle, msg.data(), static_cast<DWORD>(msg.size()), &written, nullptr) != FALSE;
        }

        const std::string utf8 = Common::Strings::Utf8FromUtf16ReplacingInvalid(msg);
        if (utf8.empty() && ! msg.empty())
        {
            return false;
        }

        DWORD written = 0;
        return WriteFile(handle, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr) != FALSE;
    };

    if (tryWrite(GetStdHandle(STD_OUTPUT_HANDLE), kMonitorHelpText))
    {
        return;
    }

    if (AttachConsole(ATTACH_PARENT_PROCESS) == FALSE)
    {
        const DWORD err = GetLastError();
        if (err != ERROR_ACCESS_DENIED)
        {
            static_cast<void>(AllocConsole());
        }
    }

    wil::unique_handle conout(CreateFileW(L"CONOUT$", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr));
    if (conout && tryWrite(conout.get(), kMonitorHelpText))
    {
        return;
    }

    ShowModalMessageDialog(hInstance, nullptr, L"RedSalamanderMonitor Help", kMonitorHelpText);
}

std::wstring GetEnvironmentVariableString(std::wstring_view name) noexcept
{
    return EnvironmentVariables::Read(name).value_or(std::wstring{});
}

bool PathExists(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::exists(path, ec) && ! ec;
}

bool EnsureDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::create_directories(path, ec);
    return ! ec && PathExists(path);
}

std::filesystem::path TryFindRepoRoot() noexcept
{
    const std::wstring configuredRoot = GetEnvironmentVariableString(L"REDSALAMANDER_REPO_ROOT");
    if (! configuredRoot.empty())
    {
        const std::filesystem::path candidate(configuredRoot);
        if (PathExists(candidate / L"Specs" / L"TestRuns"))
        {
            return candidate;
        }
    }

    wil::unique_cotaskmem_string modulePath;
    if (FAILED(wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath)) || ! modulePath)
    {
        return {};
    }

    std::filesystem::path candidate(modulePath.get());
    candidate = candidate.parent_path();
    while (! candidate.empty())
    {
        if (PathExists(candidate / L"Specs" / L"TestRuns"))
        {
            return candidate;
        }

        const std::filesystem::path parent = candidate.parent_path();
        if (parent == candidate)
        {
            break;
        }
        candidate = parent;
    }

    return {};
}

std::string Utf8FromWide(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

bool WriteUtf8TextFile(const std::filesystem::path& path, std::string_view text, bool append) noexcept
{
    if (path.empty())
    {
        return false;
    }

    if (path.has_parent_path() && ! EnsureDirectory(path.parent_path()))
    {
        return false;
    }

    const DWORD desiredAccess = append ? FILE_APPEND_DATA : GENERIC_WRITE;
    const DWORD disposition   = append ? OPEN_ALWAYS : CREATE_ALWAYS;
    wil::unique_handle file(::CreateFileW(path.c_str(), desiredAccess, FILE_SHARE_READ, nullptr, disposition, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    if (append)
    {
        ::SetFilePointer(file.get(), 0, nullptr, FILE_END);
    }

    DWORD written = 0u;
    return ::WriteFile(file.get(), text.data(), static_cast<DWORD>(text.size()), &written, nullptr) && written == text.size();
}

std::string EscapeJsonUtf8(std::string_view text)
{
    std::string escaped;
    escaped.reserve(text.size() + 16u);
    for (const char rawCh : text)
    {
        const auto ch = static_cast<unsigned char>(rawCh);
        switch (ch)
        {
            case '\\': escaped += "\\\\"; break;
            case '"': escaped += "\\\""; break;
            case '\r': escaped += "\\r"; break;
            case '\n': escaped += "\\n"; break;
            case '\t': escaped += "\\t"; break;
            default:
                if (ch < 0x20u)
                {
                    escaped += std::format("\\u{:04x}", static_cast<unsigned>(ch));
                }
                else
                {
                    escaped.push_back(static_cast<char>(ch));
                }
                break;
        }
    }

    return escaped;
}

std::string EscapeJsonWide(std::wstring_view text)
{
    return EscapeJsonUtf8(Utf8FromWide(text));
}

void AppendMonitorChromeSelfTestTrace(std::wstring_view message) noexcept
{
    if (! g_monitorChromeSelfTest.enabled || g_monitorChromeSelfTest.tracePath.empty())
    {
        return;
    }

    std::wstring line(message);
    line.append(L"\r\n");
    WriteUtf8TextFile(g_monitorChromeSelfTest.tracePath, Utf8FromWide(line), true);
}

void RecordMonitorChromeSelfTestCheck(std::wstring_view name, bool passed, std::wstring_view detail = {}) noexcept
{
    if (! g_monitorChromeSelfTest.enabled)
    {
        return;
    }

    g_monitorChromeSelfTest.checks.push_back(MonitorChromeSelfTestCheck{std::wstring(name), passed, std::wstring(detail)});
    AppendMonitorChromeSelfTestTrace(std::format(L"{}: {}{}", passed ? L"PASS" : L"FAIL", name, detail.empty() ? L"" : std::format(L" ({})", detail)));
}

std::wstring TryReadRegistryStringValue(HKEY hive, const wchar_t* subKey, const wchar_t* valueName) noexcept
{
    if (! hive || ! subKey || ! valueName)
    {
        return {};
    }

    DWORD type            = 0u;
    DWORD size            = 0u;
    const LONG sizeStatus = ::RegGetValueW(hive, subKey, valueName, RRF_RT_REG_SZ, &type, nullptr, &size);
    if (sizeStatus != ERROR_SUCCESS || type != REG_SZ || size < sizeof(wchar_t))
    {
        return {};
    }

    std::wstring value(size / sizeof(wchar_t), L'\0');
    DWORD readSize        = size;
    const LONG readStatus = ::RegGetValueW(hive, subKey, valueName, RRF_RT_REG_SZ, nullptr, value.data(), &readSize);
    if (readStatus != ERROR_SUCCESS || readSize < sizeof(wchar_t))
    {
        return {};
    }

    value.resize((readSize / sizeof(wchar_t)) - 1u);
    return value;
}

std::wstring GetComputerNameString() noexcept
{
    wchar_t buffer[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD size = static_cast<DWORD>(std::size(buffer));
    if (! ::GetComputerNameW(buffer, &size) || size == 0u)
    {
        return {};
    }

    return std::wstring(buffer, size);
}

#pragma warning(push)
#pragma warning(disable : 4625 4626)
bool TryComputeSha256(std::span<const std::byte> data, std::array<std::byte, 32>& outHash) noexcept
{
    BCRYPT_ALG_HANDLE algHandleRaw = nullptr;
    const NTSTATUS openStatus      = BCryptOpenAlgorithmProvider(&algHandleRaw, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
    if (! BCRYPT_SUCCESS(openStatus) || ! algHandleRaw)
    {
        return false;
    }

    wil::unique_bcrypt_algorithm closeAlg(algHandleRaw);

    DWORD objLen              = 0u;
    DWORD hashLen             = 0u;
    DWORD cb                  = sizeof(DWORD);
    const NTSTATUS propObject = BCryptGetProperty(closeAlg.get(), BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&objLen), cb, &cb, 0);
    const NTSTATUS propHash   = BCryptGetProperty(closeAlg.get(), BCRYPT_HASH_LENGTH, reinterpret_cast<PUCHAR>(&hashLen), cb, &cb, 0);
    if (! BCRYPT_SUCCESS(propObject) || ! BCRYPT_SUCCESS(propHash) || hashLen != static_cast<DWORD>(outHash.size()))
    {
        return false;
    }

    std::vector<std::byte> hashObject(static_cast<size_t>(objLen));
    BCRYPT_HASH_HANDLE hashHandleRaw = nullptr;
    const NTSTATUS createStatus      = BCryptCreateHash(closeAlg.get(), &hashHandleRaw, reinterpret_cast<PUCHAR>(hashObject.data()), objLen, nullptr, 0, 0);
    if (! BCRYPT_SUCCESS(createStatus) || ! hashHandleRaw)
    {
        return false;
    }

    wil::unique_bcrypt_hash destroyHash(hashHandleRaw);
    const NTSTATUS hashStatus =
        BCryptHashData(destroyHash.get(), reinterpret_cast<PUCHAR>(const_cast<std::byte*>(data.data())), static_cast<ULONG>(data.size()), 0);
    if (! BCRYPT_SUCCESS(hashStatus))
    {
        return false;
    }

    const NTSTATUS finishStatus = BCryptFinishHash(destroyHash.get(), reinterpret_cast<PUCHAR>(outHash.data()), hashLen, 0);
    return BCRYPT_SUCCESS(finishStatus);
}
#pragma warning(pop)

std::wstring GetComputerHashName() noexcept
{
    std::wstring seed = TryReadRegistryStringValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Cryptography", L"MachineGuid");
    if (seed.empty())
    {
        seed = GetComputerNameString();
    }
    if (seed.empty())
    {
        return L"unknown";
    }

    const std::string seedUtf8 = Utf8FromWide(seed);
    if (seedUtf8.empty())
    {
        return L"unknown";
    }

    std::array<std::byte, 32> hash{};
    if (! TryComputeSha256(std::as_bytes(std::span(seedUtf8.data(), seedUtf8.size())), hash))
    {
        return L"unknown";
    }

    constexpr wchar_t kHex[] = L"0123456789abcdef";
    std::wstring out;
    out.reserve(12u);
    for (size_t i = 0; i < 6u; ++i)
    {
        const uint8_t byte = std::to_integer<uint8_t>(hash[i]);
        out.push_back(kHex[(byte >> 4) & 0x0Fu]);
        out.push_back(kHex[byte & 0x0Fu]);
    }
    return out;
}

std::wstring GetTimestampFolderNameLocal() noexcept
{
    SYSTEMTIME st{};
    ::GetLocalTime(&st);
    return std::format(L"{0:04}-{1:02}-{2:02}_{3:02}{4:02}{5:02}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

uint64_t CountMetricRowsInJsonl(std::string_view jsonl, std::wstring_view metric) noexcept
{
    const std::string metricUtf8 = Utf8FromWide(metric);
    if (metricUtf8.empty())
    {
        return 0;
    }

    const std::string token = std::format("\"metric\":\"{}\"", metricUtf8);
    uint64_t count          = 0;
    size_t position         = 0;
    while ((position = jsonl.find(token, position)) != std::string_view::npos)
    {
        ++count;
        position += token.size();
    }

    return count;
}

std::optional<uint64_t> TryReadUnsignedJsonField(std::string_view row, std::string_view fieldName) noexcept
{
    const std::string token  = std::format("\"{}\":", fieldName);
    const size_t tokenOffset = row.find(token);
    if (tokenOffset == std::string_view::npos)
    {
        return std::nullopt;
    }

    size_t position = tokenOffset + token.size();
    while (position < row.size() && row[position] == ' ')
    {
        ++position;
    }

    if (position == row.size() || row[position] < '0' || row[position] > '9')
    {
        return std::nullopt;
    }

    uint64_t value = 0u;
    while (position < row.size() && row[position] >= '0' && row[position] <= '9')
    {
        const uint64_t digit = static_cast<uint64_t>(row[position] - '0');
        if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
        {
            return std::nullopt;
        }
        value = (value * 10u) + digit;
        ++position;
    }

    return value;
}

std::vector<uint64_t> ReadMetricValuesInJsonl(std::string_view jsonl, std::wstring_view metric) noexcept
{
    std::vector<uint64_t> values;
    const std::string metricUtf8 = Utf8FromWide(metric);
    if (metricUtf8.empty())
    {
        return values;
    }

    const std::string token = std::format("\"metric\":\"{}\"", metricUtf8);
    size_t rowStart         = 0u;
    while (rowStart < jsonl.size())
    {
        const size_t rowEnd        = jsonl.find('\n', rowStart);
        const std::string_view row = rowEnd == std::string_view::npos ? jsonl.substr(rowStart) : jsonl.substr(rowStart, rowEnd - rowStart);
        if (row.find(token) != std::string_view::npos)
        {
            if (const std::optional<uint64_t> value = TryReadUnsignedJsonField(row, "value"); value.has_value())
            {
                values.push_back(value.value());
            }
        }

        if (rowEnd == std::string_view::npos)
        {
            break;
        }
        rowStart = rowEnd + 1u;
    }

    return values;
}

uint64_t NearestRankPercentile(std::span<const uint64_t> sortedValues, uint64_t percentile) noexcept
{
    if (sortedValues.empty())
    {
        return 0u;
    }

    const uint64_t count = static_cast<uint64_t>(sortedValues.size());
    const uint64_t rank  = (count * percentile + 99u) / 100u;
    const size_t index   = static_cast<size_t>((std::max)(uint64_t{1u}, rank) - 1u);
    return sortedValues[(std::min)(index, sortedValues.size() - 1u)];
}

MonitorChromeMetricSummary BuildMetricSummary(std::string_view jsonl, std::wstring_view metric) noexcept
{
    MonitorChromeMetricSummary summary{.metric = metric};
    std::vector<uint64_t> values = ReadMetricValuesInJsonl(jsonl, metric);
    if (values.empty())
    {
        return summary;
    }

    std::sort(values.begin(), values.end());
    summary.count = static_cast<uint64_t>(values.size());
    summary.p50   = NearestRankPercentile(values, 50u);
    summary.p95   = NearestRankPercentile(values, 95u);
    summary.p99   = NearestRankPercentile(values, 99u);
    summary.max   = values.back();
    return summary;
}

std::string ReadMonitorPerfJsonl() noexcept
{
    if (g_monitorChromeSelfTest.perfPath.empty())
    {
        return {};
    }

    std::ifstream input(g_monitorChromeSelfTest.perfPath, std::ios::binary);
    if (! input)
    {
        return {};
    }

    return std::string((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
}

template <size_t Count>
std::vector<MonitorChromeMetricPresence> BuildMetricPresence(std::string_view jsonl, const std::array<std::wstring_view, Count>& metrics)
{
    std::vector<MonitorChromeMetricPresence> presence;
    presence.reserve(metrics.size());
    for (const std::wstring_view metric : metrics)
    {
        presence.push_back(MonitorChromeMetricPresence{metric, CountMetricRowsInJsonl(jsonl, metric)});
    }

    return presence;
}

template <size_t Count>
std::vector<MonitorChromeMetricSummary> BuildMetricSummaries(std::string_view jsonl, const std::array<std::wstring_view, Count>& metrics)
{
    std::vector<MonitorChromeMetricSummary> summaries;
    summaries.reserve(metrics.size());
    for (const std::wstring_view metric : metrics)
    {
        summaries.push_back(BuildMetricSummary(jsonl, metric));
    }

    return summaries;
}

bool AreAllMetricsPresent(std::span<const MonitorChromeMetricPresence> metricPresence) noexcept
{
    return std::all_of(metricPresence.begin(), metricPresence.end(), [](const MonitorChromeMetricPresence& metric) noexcept { return metric.count > 0u; });
}

std::filesystem::path GetDefaultMonitorPerfJsonlPath() noexcept
{
    const std::filesystem::path repoRoot = TryFindRepoRoot();
    if (! repoRoot.empty())
    {
        return repoRoot / L"Specs" / L"TestRuns" / GetComputerHashName() / L"RedSalamanderMonitor" / GetTimestampFolderNameLocal() / L"perf_metrics.jsonl";
    }

    std::error_code ec;
    std::filesystem::path root = std::filesystem::temp_directory_path(ec);
    if (ec)
    {
        root = L".";
    }
    return root / L"RedSalamander" / L"Perf" / (std::wstring(kAppId) + L"_" + GetTimestampFolderNameLocal() + L".jsonl");
}

void FinalizeMonitorChromeSelfTest(bool passed, std::wstring_view summary) noexcept
{
    if (! g_monitorChromeSelfTest.enabled || g_monitorChromeSelfTest.completed)
    {
        return;
    }

    const std::string perfJsonl                                        = ReadMonitorPerfJsonl();
    const std::vector<MonitorChromeMetricPresence> metricPresence      = BuildMetricPresence(perfJsonl, kRequiredMonitorFrameMetrics);
    const bool allRequiredMetricsPresent                               = AreAllMetricsPresent(metricPresence);
    const std::vector<MonitorChromeMetricPresence> burstMetricPresence = g_monitorEtwBurstOptions.latencyMode
                                                                             ? BuildMetricPresence(perfJsonl, kRequiredMonitorEtwBurstLatencyMetrics)
                                                                             : std::vector<MonitorChromeMetricPresence>{};
    const bool allBurstMetricsPresent                                  = ! g_monitorEtwBurstOptions.latencyMode || AreAllMetricsPresent(burstMetricPresence);
    const std::vector<MonitorChromeMetricSummary> burstMetricSummaries = g_monitorEtwBurstOptions.latencyMode
                                                                             ? BuildMetricSummaries(perfJsonl, kSummarizedMonitorEtwBurstLatencyMetrics)
                                                                             : std::vector<MonitorChromeMetricSummary>{};
    const std::vector<MonitorChromeMetricPresence> scrollbackMetricPresence = g_monitorScrollbackSelfTestOptions.enabled
                                                                                  ? BuildMetricPresence(perfJsonl, kRequiredMonitorScrollbackMetrics)
                                                                                  : std::vector<MonitorChromeMetricPresence>{};
    const bool allScrollbackMetricsPresent = ! g_monitorScrollbackSelfTestOptions.enabled || AreAllMetricsPresent(scrollbackMetricPresence);
    const std::vector<MonitorChromeMetricSummary> scrollbackMetricSummaries = g_monitorScrollbackSelfTestOptions.enabled
                                                                                  ? BuildMetricSummaries(perfJsonl, kSummarizedMonitorScrollbackMetrics)
                                                                                  : std::vector<MonitorChromeMetricSummary>{};

    RecordMonitorChromeSelfTestCheck(
        L"required monitor frame metrics", allRequiredMetricsPresent, allRequiredMetricsPresent ? L"" : L"one or more required monitor frame metrics missing");
    if (g_monitorEtwBurstOptions.latencyMode)
    {
        RecordMonitorChromeSelfTestCheck(
            L"monitor etw burst latency metrics", allBurstMetricsPresent, allBurstMetricsPresent ? L"" : L"one or more etw burst latency metrics missing");
    }
    if (g_monitorScrollbackSelfTestOptions.enabled)
    {
        RecordMonitorChromeSelfTestCheck(L"monitor scrollback frame metrics",
                                         allScrollbackMetricsPresent,
                                         allScrollbackMetricsPresent ? L"" : L"one or more scrollback frame metrics missing");
    }

    const bool finalPassed = passed && allRequiredMetricsPresent && allBurstMetricsPresent && allScrollbackMetricsPresent;
    const std::wstring finalSummary =
        finalPassed
            ? std::wstring(summary)
            : (passed && ! allRequiredMetricsPresent
                   ? L"Monitor chrome selftest passed functionally, but required frame metrics were missing."
                   : (passed && ! allBurstMetricsPresent ? L"Monitor chrome selftest passed functionally, but ETW burst latency metrics were missing."
                                                         : (passed && ! allScrollbackMetricsPresent
                                                                ? L"Monitor chrome selftest passed functionally, but scrollback frame metrics were missing."
                                                                : std::wstring(summary))));

    g_monitorChromeSelfTest.completed = true;
    g_monitorChromeSelfTest.exitCode  = finalPassed ? 0 : 1;

    RecordMonitorChromeSelfTestCheck(L"overall", finalPassed, finalSummary);

    std::string json;
    json += "{\n";
    json += std::format("  \"scenario\": \"{}\",\n", EscapeJsonWide(kMonitorScenarioName));
    json += std::format("  \"status\": \"{}\",\n", finalPassed ? "passed" : "failed");
    json += std::format("  \"summary\": \"{}\",\n", EscapeJsonWide(finalSummary));
    json += std::format("  \"machineHash\": \"{}\",\n", EscapeJsonWide(g_monitorChromeSelfTest.machineHash));
    json += std::format("  \"runId\": \"{}\",\n", EscapeJsonWide(g_monitorChromeSelfTest.runId));
    json += "  \"checks\": [\n";
    for (size_t index = 0; index < g_monitorChromeSelfTest.checks.size(); ++index)
    {
        const auto& check = g_monitorChromeSelfTest.checks[index];
        json += std::format("    {{\"name\": \"{}\", \"status\": \"{}\", \"detail\": \"{}\"}}{}\n",
                            EscapeJsonWide(check.name),
                            check.passed ? "passed" : "failed",
                            EscapeJsonWide(check.detail),
                            (index + 1u) == g_monitorChromeSelfTest.checks.size() ? "" : ",");
    }
    json += "  ],\n";
    json += "  \"monitorFrameMetricPresence\": {\n";
    json += std::format("    \"allPresent\": {},\n", allRequiredMetricsPresent ? "true" : "false");
    json += "    \"metrics\": [\n";
    for (size_t index = 0; index < metricPresence.size(); ++index)
    {
        const MonitorChromeMetricPresence& metric = metricPresence[index];
        json += std::format("      {{\"metric\": \"{}\", \"present\": {}, \"count\": {}}}{}\n",
                            EscapeJsonWide(metric.metric),
                            metric.count > 0u ? "true" : "false",
                            metric.count,
                            (index + 1u) == metricPresence.size() ? "" : ",");
    }
    json += "    ]\n";
    json += "  },\n";
    json += "  \"monitorEtwBurstLatency\": {\n";
    json += std::format("    \"enabled\": {},\n", g_monitorEtwBurstOptions.latencyMode ? "true" : "false");
    json += std::format("    \"count\": {},\n", static_cast<uint64_t>(g_monitorEtwBurstOptions.count));
    json += std::format("    \"payloadChars\": {},\n", static_cast<uint64_t>(g_monitorEtwBurstOptions.payloadChars));
    json += "    \"metricPresence\": {\n";
    json += std::format("      \"allPresent\": {},\n", allBurstMetricsPresent ? "true" : "false");
    json += "      \"metrics\": [\n";
    for (size_t index = 0; index < burstMetricPresence.size(); ++index)
    {
        const MonitorChromeMetricPresence& metric = burstMetricPresence[index];
        json += std::format("        {{\"metric\": \"{}\", \"present\": {}, \"count\": {}}}{}\n",
                            EscapeJsonWide(metric.metric),
                            metric.count > 0u ? "true" : "false",
                            metric.count,
                            (index + 1u) == burstMetricPresence.size() ? "" : ",");
    }
    json += "      ]\n";
    json += "    },\n";
    json += "    \"metricSummary\": [\n";
    for (size_t index = 0; index < burstMetricSummaries.size(); ++index)
    {
        const MonitorChromeMetricSummary& metric = burstMetricSummaries[index];
        json += std::format("      {{\"metric\": \"{}\", \"count\": {}, \"p50\": {}, \"p95\": {}, \"p99\": {}, \"max\": {}}}{}\n",
                            EscapeJsonWide(metric.metric),
                            metric.count,
                            metric.p50,
                            metric.p95,
                            metric.p99,
                            metric.max,
                            (index + 1u) == burstMetricSummaries.size() ? "" : ",");
    }
    json += "    ]\n";
    json += "  },\n";
    json += "  \"monitorScrollbackSelfTest\": {\n";
    json += std::format("    \"enabled\": {},\n", g_monitorScrollbackSelfTestOptions.enabled ? "true" : "false");
    json += "    \"metricPresence\": {\n";
    json += std::format("      \"allPresent\": {},\n", allScrollbackMetricsPresent ? "true" : "false");
    json += "      \"metrics\": [\n";
    for (size_t index = 0; index < scrollbackMetricPresence.size(); ++index)
    {
        const MonitorChromeMetricPresence& metric = scrollbackMetricPresence[index];
        json += std::format("        {{\"metric\": \"{}\", \"present\": {}, \"count\": {}}}{}\n",
                            EscapeJsonWide(metric.metric),
                            metric.count > 0u ? "true" : "false",
                            metric.count,
                            (index + 1u) == scrollbackMetricPresence.size() ? "" : ",");
    }
    json += "      ]\n";
    json += "    },\n";
    json += "    \"metricSummary\": [\n";
    for (size_t index = 0; index < scrollbackMetricSummaries.size(); ++index)
    {
        const MonitorChromeMetricSummary& metric = scrollbackMetricSummaries[index];
        json += std::format("      {{\"metric\": \"{}\", \"count\": {}, \"p50\": {}, \"p95\": {}, \"p99\": {}, \"max\": {}}}{}\n",
                            EscapeJsonWide(metric.metric),
                            metric.count,
                            metric.p50,
                            metric.p95,
                            metric.p99,
                            metric.max,
                            (index + 1u) == scrollbackMetricSummaries.size() ? "" : ",");
    }
    json += "    ]\n";
    json += "  }\n";
    json += "}\n";

    WriteUtf8TextFile(g_monitorChromeSelfTest.resultsPath, json, false);
    Debug::Perf::ClearJsonlOutput();
}

bool InitializeMonitorChromeSelfTestArtifacts() noexcept
{
    if (! g_monitorChromeSelfTest.enabled)
    {
        return true;
    }

    g_monitorChromeSelfTest.completed = false;
    g_monitorChromeSelfTest.exitCode  = 1;
    g_monitorChromeSelfTest.checks.clear();
    g_monitorChromeSelfTest.machineHash = GetComputerHashName();
    g_monitorChromeSelfTest.runId       = GetTimestampFolderNameLocal();
    g_monitorChromeSelfTest.repoRoot    = TryFindRepoRoot();
    if (g_monitorChromeSelfTest.repoRoot.empty())
    {
        return false;
    }

    g_monitorChromeSelfTest.runRoot = g_monitorChromeSelfTest.repoRoot / L"Specs" / L"TestRuns" / g_monitorChromeSelfTest.machineHash /
                                      std::wstring(kMonitorAreaName) / g_monitorChromeSelfTest.runId;
    if (! EnsureDirectory(g_monitorChromeSelfTest.runRoot))
    {
        return false;
    }

    g_monitorChromeSelfTest.tracePath   = g_monitorChromeSelfTest.runRoot / L"trace.txt";
    g_monitorChromeSelfTest.resultsPath = g_monitorChromeSelfTest.runRoot / L"results.json";
    g_monitorChromeSelfTest.perfPath    = g_monitorChromeSelfTest.runRoot / L"perf_metrics.jsonl";

    WriteUtf8TextFile(g_monitorChromeSelfTest.tracePath, {}, false);
    AppendMonitorChromeSelfTestTrace(std::format(L"Scenario: {}", kMonitorScenarioName));
    AppendMonitorChromeSelfTestTrace(std::format(L"ArchiveToRepo: {}", g_monitorChromeSelfTest.runRoot.native()));

    Debug::Perf::ConfigureJsonlOutput(
        g_monitorChromeSelfTest.perfPath, L"MonitorChromeSelfTest", L"Debug", {}, {}, g_monitorChromeSelfTest.machineHash, g_monitorChromeSelfTest.runId);

    return true;
}

bool RelaunchSelfElevated(HWND owner, std::wstring_view extraArg) noexcept
{
    wil::unique_cotaskmem_string exePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, exePath);
    if (FAILED(hr) || ! exePath)
    {
        return false;
    }

    const std::wstring params = BuildRelaunchParameters(extraArg);

    SHELLEXECUTEINFOW execInfo{};
    execInfo.cbSize       = sizeof(execInfo);
    execInfo.fMask        = SEE_MASK_NOCLOSEPROCESS;
    execInfo.hwnd         = owner;
    execInfo.lpVerb       = L"runas";
    execInfo.lpFile       = exePath.get();
    execInfo.lpParameters = params.empty() ? nullptr : params.c_str();
    execInfo.nShow        = SW_SHOWNORMAL;

    if (! ::ShellExecuteExW(&execInfo))
    {
        return false;
    }

    wil::unique_handle launchedProcess(execInfo.hProcess);
    return true;
}

bool IsHighContrastEnabled() noexcept
{
    HIGHCONTRASTW hc{};
    hc.cbSize = sizeof(hc);
    if (! SystemParametersInfoW(SPI_GETHIGHCONTRAST, hc.cbSize, &hc, 0))
    {
        return false;
    }
    return (hc.dwFlags & HCF_HIGHCONTRASTON) != 0;
}

bool IsSystemDarkModeEnabled() noexcept
{
    DWORD appsUseLightTheme = 1;
    DWORD dataSize          = sizeof(appsUseLightTheme);
    const LSTATUS status    = RegGetValueW(HKEY_CURRENT_USER,
                                           L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                                           L"AppsUseLightTheme",
                                           RRF_RT_REG_DWORD,
                                           nullptr,
                                           &appsUseLightTheme,
                                           &dataSize);
    if (status != ERROR_SUCCESS)
    {
        return false;
    }
    return appsUseLightTheme == 0;
}

Common::Settings::MonitorFilterPreset PresetFromLegacy(int legacy) noexcept
{
    switch (legacy)
    {
        case 0: return Common::Settings::MonitorFilterPreset::ErrorsOnly;
        case 1: return Common::Settings::MonitorFilterPreset::ErrorsWarnings;
        case 2: return Common::Settings::MonitorFilterPreset::AllTypes;
        default: return Common::Settings::MonitorFilterPreset::Custom;
    }
}

int LegacyFromPreset(Common::Settings::MonitorFilterPreset preset) noexcept
{
    switch (preset)
    {
        case Common::Settings::MonitorFilterPreset::ErrorsOnly: return 0;
        case Common::Settings::MonitorFilterPreset::ErrorsWarnings: return 1;
        case Common::Settings::MonitorFilterPreset::AllTypes: return 2;
        case Common::Settings::MonitorFilterPreset::Custom:
        default: return -1;
    }
}

int InferLegacyPresetFromMask(uint32_t mask) noexcept
{
    mask &= kMonitorFilterAllMask;
    switch (mask)
    {
        case kMonitorPresetErrorsOnlyMask: return 0;      // Errors only
        case kMonitorPresetErrorsWarningsMask: return 1;  // Errors+Warnings
        case kMonitorFilterAllMask: return 2;             // All types
        case kMonitorPresetErrorsPerfDebugMask: return 3; // Errors+Perf+Debug
        default: return -1;                               // Custom
    }
}

const Common::Settings::ThemeDefinition* FindThemeById(std::wstring_view id) noexcept
{
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    for (const auto& def : g_fileThemes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    return nullptr;
}

void ApplyMonitorThemeOverrides(ColorTextView::Theme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto apply = [&](std::wstring_view key, D2D1_COLOR_F& target) noexcept
    {
        const auto it = colors.find(std::wstring(key));
        if (it == colors.end())
        {
            return;
        }
        target = ColorFromArgb(it->second);
    };

    apply(L"monitor.textView.bg", theme.bg);
    apply(L"monitor.textView.fg", theme.fg);
    apply(L"monitor.textView.caret", theme.caret);
    apply(L"monitor.textView.selection", theme.selection);
    apply(L"monitor.textView.searchHighlight", theme.searchHighlight);
    apply(L"monitor.textView.gutterBg", theme.gutterBg);
    apply(L"monitor.textView.gutterFg", theme.gutterFg);
    apply(L"monitor.textView.metaText", theme.metaText);
    apply(L"monitor.textView.metaError", theme.metaError);
    apply(L"monitor.textView.metaWarning", theme.metaWarning);
    apply(L"monitor.textView.metaInfo", theme.metaInfo);
    apply(L"monitor.textView.metaPerf", theme.metaPerf);
    apply(L"monitor.textView.metaDebug", theme.metaDebug);
}

[[nodiscard]] std::optional<uint32_t> FindMonitorThemeColorArgb(const ColorTextView::Theme& theme, std::wstring_view key) noexcept
{
    const auto argb = [](const D2D1_COLOR_F& color) noexcept
    {
        const auto channel = [](float value) noexcept
        { return static_cast<uint32_t>(std::clamp(std::lround(std::clamp(value, 0.0f, 1.0f) * 255.0f), 0l, 255l)); };
        return (channel(color.a) << 24u) | (channel(color.r) << 16u) | (channel(color.g) << 8u) | channel(color.b);
    };
    if (key == L"monitor.textView.bg") return argb(theme.bg);
    if (key == L"monitor.textView.fg") return argb(theme.fg);
    if (key == L"monitor.textView.caret") return argb(theme.caret);
    if (key == L"monitor.textView.selection") return argb(theme.selection);
    if (key == L"monitor.textView.searchHighlight") return argb(theme.searchHighlight);
    if (key == L"monitor.textView.gutterBg") return argb(theme.gutterBg);
    if (key == L"monitor.textView.gutterFg") return argb(theme.gutterFg);
    if (key == L"monitor.textView.metaText") return argb(theme.metaText);
    if (key == L"monitor.textView.metaError") return argb(theme.metaError);
    if (key == L"monitor.textView.metaWarning") return argb(theme.metaWarning);
    if (key == L"monitor.textView.metaInfo") return argb(theme.metaInfo);
    if (key == L"monitor.textView.metaPerf") return argb(theme.metaPerf);
    if (key == L"monitor.textView.metaDebug") return argb(theme.metaDebug);
    return std::nullopt;
}

ColorTextView::Theme MakeMonitorThemeHighContrast() noexcept
{
    const auto sys = [](int idx, float alpha = 1.0f) noexcept
    {
        const COLORREF c = GetSysColor(idx);
        return D2D1::ColorF(
            static_cast<float>(GetRValue(c)) / 255.0f, static_cast<float>(GetGValue(c)) / 255.0f, static_cast<float>(GetBValue(c)) / 255.0f, alpha);
    };

    ColorTextView::Theme t;
    t.bg              = sys(COLOR_WINDOW);
    t.fg              = sys(COLOR_WINDOWTEXT);
    t.caret           = t.fg;
    t.selection       = sys(COLOR_HIGHLIGHT, 0.45f);
    t.searchHighlight = sys(COLOR_HIGHLIGHT, 0.35f);
    t.gutterBg        = sys(COLOR_BTNFACE);
    t.gutterFg        = sys(COLOR_WINDOWTEXT);
    t.metaText        = sys(COLOR_WINDOWTEXT, 0.85f);
    t.metaError       = sys(COLOR_HIGHLIGHTTEXT);
    t.metaWarning     = sys(COLOR_HIGHLIGHTTEXT);
    t.metaInfo        = sys(COLOR_HIGHLIGHTTEXT);
    t.metaPerf        = sys(COLOR_HIGHLIGHTTEXT);
    t.metaDebug       = sys(COLOR_HIGHLIGHTTEXT);
    return t;
}

ColorTextView::Theme MakeMonitorThemeLight() noexcept
{
    return ColorTextView::Theme{};
}

ColorTextView::Theme MakeMonitorThemeDark() noexcept
{
    ColorTextView::Theme t;
    t.bg              = D2D1::ColorF(0.08f, 0.08f, 0.08f);
    t.fg              = D2D1::ColorF(0.92f, 0.92f, 0.92f);
    t.caret           = t.fg;
    t.selection       = D2D1::ColorF(0.20f, 0.55f, 0.95f, 0.35f);
    t.searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.35f);
    t.gutterBg        = D2D1::ColorF(0.12f, 0.12f, 0.12f);
    t.gutterFg        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
    t.metaText        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
    t.metaError       = D2D1::ColorF(D2D1::ColorF::Red);
    t.metaWarning     = D2D1::ColorF(D2D1::ColorF::Orange);
    t.metaInfo        = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    t.metaPerf        = D2D1::ColorF(D2D1::ColorF::MediumSeaGreen);
    t.metaDebug       = D2D1::ColorF(D2D1::ColorF::MediumPurple);
    return t;
}

MonitorResolvedTheme ResolveMonitorTheme() noexcept
{
    MonitorResolvedTheme resolved{};
    const Common::Settings::UiSettings ui = g_settings.ui.value_or(Common::Settings::UiSettings{});
    if (IsHighContrastEnabled())
    {
        resolved.textView       = MakeMonitorThemeHighContrast();
        resolved.dark           = false;
        resolved.highContrast   = true;
        resolved.compactMode    = ui.compactMode;
        resolved.windowBackdrop = ui.windowBackdrop;
        switch (ui.reducedMotion)
        {
            case Common::Settings::ReducedMotionMode::On: resolved.reducedMotionOverride = true; break;
            case Common::Settings::ReducedMotionMode::Off: resolved.reducedMotionOverride = false; break;
            case Common::Settings::ReducedMotionMode::System: resolved.reducedMotionOverride.reset(); break;
        }
        return resolved;
    }

    std::wstring_view themeId                       = g_settings.theme.currentThemeId;
    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        custom = FindThemeById(themeId);
    }

    std::wstring_view baseThemeId                               = themeId;
    const std::unordered_map<std::wstring, uint32_t>* overrides = nullptr;
    Common::Settings::ResolvedThemeColors resolvedOverrides;
    if (custom)
    {
        baseThemeId = custom->baseThemeId;
    }

    const bool systemDark = IsSystemDarkModeEnabled();
    resolved.highContrast = false;
    resolved.rainbow      = baseThemeId == L"builtin/rainbow";

    if (baseThemeId == L"builtin/highContrast")
    {
        resolved.textView     = MakeMonitorThemeHighContrast();
        resolved.dark         = false;
        resolved.highContrast = true;
    }
    else if (baseThemeId == L"builtin/dark")
    {
        resolved.textView = MakeMonitorThemeDark();
        resolved.dark     = true;
    }
    else if (baseThemeId == L"builtin/light")
    {
        resolved.textView = MakeMonitorThemeLight();
        resolved.dark     = false;
    }
    else if (baseThemeId == L"builtin/rainbow")
    {
        resolved.textView = systemDark ? MakeMonitorThemeDark() : MakeMonitorThemeLight();
        resolved.dark     = systemDark;
    }
    else
    {
        resolved.textView = systemDark ? MakeMonitorThemeDark() : MakeMonitorThemeLight();
        resolved.dark     = systemDark;
    }

    if (custom)
    {
        auto context = Common::Settings::MakeSystemThemeResolutionContext(resolved.dark);
        context.highContrast = resolved.highContrast;
        const ColorTextView::Theme baseTextView = resolved.textView;
        context.baseColor = [baseTextView](std::wstring_view key) { return FindMonitorThemeColorArgb(baseTextView, key); };
        if (SUCCEEDED(Common::Settings::ResolveThemeDefinition(*custom, context, resolvedOverrides)))
        {
            overrides = &resolvedOverrides.colors;
        }
    }

    if (overrides)
    {
        ApplyMonitorThemeOverrides(resolved.textView, *overrides);
    }

    resolved.compactMode    = ui.compactMode;
    resolved.windowBackdrop = ui.windowBackdrop;
    switch (ui.reducedMotion)
    {
        case Common::Settings::ReducedMotionMode::On: resolved.reducedMotionOverride = true; break;
        case Common::Settings::ReducedMotionMode::Off: resolved.reducedMotionOverride = false; break;
        case Common::Settings::ReducedMotionMode::System: resolved.reducedMotionOverride.reset(); break;
    }

    return resolved;
}

[[nodiscard]] ThemePalette MakeMonitorDxPalette(const MonitorResolvedTheme& theme) noexcept
{
    ThemePalette palette          = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);
    palette.dark                  = theme.dark;
    palette.highContrast          = theme.highContrast;
    palette.rainbowMode           = theme.rainbow;
    palette.accent                = D2D1::ColorF(theme.textView.selection.r, theme.textView.selection.g, theme.textView.selection.b, 1.0f);
    palette.windowBackground      = theme.textView.bg;
    palette.surfaceBackground     = BlendColor(theme.textView.bg, theme.textView.gutterBg, theme.dark ? 0.28f : 0.40f);
    palette.cardBackground        = BlendColor(palette.windowBackground, palette.surfaceBackground, theme.dark ? 0.68f : 0.82f);
    palette.headerBackground      = palette.cardBackground;
    palette.headerHovered         = BlendColor(palette.cardBackground, palette.accent, theme.dark ? 0.20f : 0.12f);
    palette.headerPressed         = BlendColor(palette.cardBackground, palette.accent, theme.dark ? 0.28f : 0.18f);
    palette.border                = BlendColor(theme.textView.bg, theme.textView.fg, theme.dark ? 0.28f : 0.18f);
    palette.borderDefault         = palette.border;
    palette.borderStrong          = BlendColor(theme.textView.bg, theme.textView.fg, theme.dark ? 0.42f : 0.28f);
    palette.gridLine              = BlendColor(palette.windowBackground, palette.borderDefault, theme.dark ? 0.72f : 0.48f);
    palette.text                  = theme.textView.fg;
    palette.subduedText           = theme.textView.metaText;
    palette.disabledText          = BlendColor(theme.textView.bg, theme.textView.metaText, 0.48f);
    palette.selectionFill         = palette.accent;
    palette.selectionText = Common::Colors::WeightedSrgbLuminanceWithoutLinearization(palette.accent.r, palette.accent.g, palette.accent.b) < 0.56
                                ? D2D1::ColorF(D2D1::ColorF::White)
                                : D2D1::ColorF(D2D1::ColorF::Black);
    palette.selectionInactiveFill = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.highContrast ? 1.0f : 0.55f);
    palette.focusStroke           = theme.textView.metaInfo;
    palette.hoverFill             = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.16f : 0.10f);
    palette.pressedFill           = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.22f : 0.16f);
    palette.buttonFill            = palette.cardBackground;
    palette.buttonBorder          = palette.borderDefault;
    palette.buttonHotFill         = palette.headerHovered;
    palette.buttonPressedFill     = palette.headerPressed;
    palette.inputFill             = BlendColor(palette.cardBackground, palette.windowBackground, theme.dark ? 0.70f : 0.78f);
    palette.inputBorder           = palette.borderDefault;
    palette.scrollbarTrack        = D2D1::ColorF(theme.textView.fg.r, theme.textView.fg.g, theme.textView.fg.b, theme.dark ? 0.12f : 0.06f);
    palette.scrollbarThumb        = D2D1::ColorF(theme.textView.fg.r, theme.textView.fg.g, theme.textView.fg.b, theme.dark ? 0.28f : 0.18f);
    palette.scrollbarThumbHot     = D2D1::ColorF(theme.textView.fg.r, theme.textView.fg.g, theme.textView.fg.b, theme.dark ? 0.40f : 0.28f);
    palette.infoFill              = BlendColor(theme.textView.bg, theme.textView.metaInfo, theme.dark ? 0.22f : 0.14f);
    palette.infoText              = theme.textView.metaInfo;
    palette.warningFill           = BlendColor(theme.textView.bg, theme.textView.metaWarning, theme.dark ? 0.22f : 0.14f);
    palette.warningText           = theme.textView.metaWarning;
    palette.errorFill             = BlendColor(theme.textView.bg, theme.textView.metaError, theme.dark ? 0.20f : 0.12f);
    palette.errorText             = theme.textView.metaError;
    palette.density               = theme.compactMode ? RedSalamander::DxUi::Density::Compact : RedSalamander::DxUi::Density::Standard;
    if (theme.reducedMotionOverride.has_value())
    {
        palette.reducedMotion = theme.reducedMotionOverride.value();
    }
    return palette;
}

[[nodiscard]] MonitorChromeMetrics ResolveMonitorChromeMetrics(bool compactMode) noexcept
{
    if (! compactMode)
    {
        return {};
    }

    return MonitorChromeMetrics{
        .toolbarHeightDip         = 36.0f,
        .statusStripHeightDip     = 20.0f,
        .toolbarPaddingDip        = 6.0f,
        .toolbarButtonHeightDip   = 28.0f,
        .toolbarGapDip            = 4.0f,
        .toolbarSeparatorWidthDip = 10.0f,
        .toolbarMinButtonWidthDip = 56.0f,
        .toolbarMinToggleWidthDip = 96.0f,
        .toolbarLabelCharWidthDip = 7.5f,
        .toolbarTextPaddingDip    = 20.0f,
        .toolbarToggleChromeDip   = 46.0f,
        .statusAutoWidthDip       = 74.0f,
        .statusFilterWidthDip     = 188.0f,
        .statusVisibleWidthDip    = 124.0f,
        .statusTotalWidthDip      = 124.0f,
    };
}

void ApplyMonitorTheme() noexcept
{
    const MonitorResolvedTheme theme = ResolveMonitorTheme();
    g_colorView.SetTheme(theme.textView);

    const ThemePalette palette = MakeMonitorDxPalette(theme);
    g_toolbarDxHost.SetTheme(palette);
    g_statusDxHost.SetTheme(palette);
    if (g_statusStrip)
    {
        const MonitorChromeMetrics metrics = ResolveMonitorChromeMetrics(theme.compactMode);
        g_statusStrip->SetSections({
            StatusStrip::Section{.text = std::wstring(g_statusStrip->GetSectionText(0u)), .widthDip = metrics.statusAutoWidthDip},
            StatusStrip::Section{.text = std::wstring(g_statusStrip->GetSectionText(1u)), .widthDip = metrics.statusFilterWidthDip},
            StatusStrip::Section{.text = std::wstring(g_statusStrip->GetSectionText(2u)), .widthDip = metrics.statusVisibleWidthDip},
            StatusStrip::Section{.text = std::wstring(g_statusStrip->GetSectionText(3u)), .widthDip = metrics.statusTotalWidthDip},
            StatusStrip::Section{.text = std::wstring(g_statusStrip->GetSectionText(4u)), .widthDip = 0.0f},
        });
    }
    if (g_hMainWindow && IsWindow(g_hMainWindow) != FALSE)
    {
        Common::WindowBackdrop::ApplyResolvedWindowBackdrop(g_hMainWindow, theme.windowBackdrop, Common::WindowBackdrop::Target::Primary, theme.highContrast);
    }
    SyncToolbarState();
    UpdateStatusBar();
    LayoutToolbarControls();
    if (g_hMainWindow && IsWindow(g_hMainWindow) != FALSE)
    {
        AdjustLayout(g_hMainWindow);
    }
}

bool TryFindMenuPathToCommand(HMENU menu, UINT commandId, std::vector<HMENU>& path) noexcept
{
    if (! menu)
    {
        return false;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return false;
    }

    path.push_back(menu);

    for (int pos = 0; pos < count; ++pos)
    {
        const UINT id = GetMenuItemID(menu, pos);
        if (id == commandId)
        {
            return true;
        }

        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (TryFindMenuPathToCommand(subMenu, commandId, path))
        {
            return true;
        }
    }

    path.pop_back();
    return false;
}

int FindMenuItemPosById(HMENU menu, UINT id) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return -1;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        if (GetMenuItemID(menu, pos) == id)
        {
            return pos;
        }
    }

    return -1;
}

void DeleteMenuItemsFromPosition(HMENU menu, int startPos) noexcept
{
    if (! menu)
    {
        return;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return;
    }

    for (int pos = count - 1; pos >= startPos; --pos)
    {
        DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);
    }
}

void EnsureThemeMenuHandle(HWND hWnd) noexcept
{
    if (g_viewThemeMenu)
    {
        return;
    }

    HMENU mainMenu = GetMenu(hWnd);
    if (! mainMenu)
    {
        return;
    }

    std::vector<HMENU> themePath;
    if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_THEME_SYSTEM, themePath) && ! themePath.empty())
    {
        g_viewThemeMenu = themePath.back();
    }
}

void RebuildThemeMenuDynamicItems(HWND hWnd)
{
    EnsureThemeMenuHandle(hWnd);
    if (! g_viewThemeMenu)
    {
        return;
    }

    const int lastBuiltInPos = FindMenuItemPosById(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP);
    if (lastBuiltInPos < 0)
    {
        return;
    }

    DeleteMenuItemsFromPosition(g_viewThemeMenu, lastBuiltInPos + 1);

    g_customThemeMenuIdToThemeId.clear();
    g_customThemeIdToMenuId.clear();

    std::unordered_map<std::wstring, const Common::Settings::ThemeDefinition*> settingsThemesById;
    settingsThemesById.reserve(g_settings.theme.themes.size());
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        settingsThemesById[def.id] = &def;
    }

    std::vector<const Common::Settings::ThemeDefinition*> settingsThemes;
    settingsThemes.reserve(settingsThemesById.size());
    for (const auto& [id, def] : settingsThemesById)
    {
        settingsThemes.push_back(def);
    }

    std::vector<const Common::Settings::ThemeDefinition*> fileThemes;
    fileThemes.reserve(g_fileThemes.size());
    for (const auto& def : g_fileThemes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        if (settingsThemesById.contains(def.id))
        {
            continue; // settings version wins
        }
        fileThemes.push_back(&def);
    }

    const auto byNameThenId = [](const Common::Settings::ThemeDefinition* a, const Common::Settings::ThemeDefinition* b)
    {
        if (a->name == b->name)
        {
            return a->id < b->id;
        }
        return a->name < b->name;
    };

    std::sort(fileThemes.begin(), fileThemes.end(), byNameThenId);
    std::sort(settingsThemes.begin(), settingsThemes.end(), byNameThenId);

    if (! fileThemes.empty() || ! settingsThemes.empty())
    {
        if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
        {
            return;
        }

        UINT nextId = kCustomThemeMenuIdFirst;

        const auto addThemes = [&](const std::vector<const Common::Settings::ThemeDefinition*>& themes)
        {
            for (const auto* def : themes)
            {
                if (nextId > kCustomThemeMenuIdLast)
                {
                    break;
                }

                std::wstring label = def->name.empty() ? def->id : def->name;
                label              = EscapeMenuLabel(label);

                if (! AppendMenuW(g_viewThemeMenu, MF_STRING, nextId, label.c_str()))
                {
                    return false;
                }

                g_customThemeMenuIdToThemeId[nextId] = def->id;
                g_customThemeIdToMenuId[def->id]     = nextId;
                ++nextId;
            }

            return true;
        };

        if (! addThemes(fileThemes))
        {
            return;
        }

        if (! fileThemes.empty() && ! settingsThemes.empty() && nextId <= kCustomThemeMenuIdLast)
        {
            if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
            {
                return;
            }
        }

        if (! addThemes(settingsThemes))
        {
            return;
        }
    }

    DrawMenuBar(hWnd);
}

void UpdateThemeMenuChecks() noexcept
{
    if (! g_viewThemeMenu)
    {
        return;
    }

    const bool highContrastEnabled = IsHighContrastEnabled();

    EnableMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, MF_BYCOMMAND | MF_GRAYED);
    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, static_cast<UINT>(MF_BYCOMMAND | (highContrastEnabled ? MF_CHECKED : MF_UNCHECKED)));

    for (const auto& [id, _] : g_customThemeMenuIdToThemeId)
    {
        CheckMenuItem(g_viewThemeMenu, id, MF_BYCOMMAND | MF_UNCHECKED);
    }

    const std::wstring currentThemeId = g_settings.theme.currentThemeId;
    const auto customIt               = g_customThemeIdToMenuId.find(currentThemeId);
    if (customIt != g_customThemeIdToMenuId.end())
    {
        CheckMenuItem(g_viewThemeMenu, customIt->second, MF_BYCOMMAND | MF_CHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);
        return;
    }

    if (currentThemeId == L"builtin/highContrast")
    {
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_CHECKED);
        return;
    }

    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);

    UINT checkedId = IDM_VIEW_THEME_SYSTEM;
    if (currentThemeId == L"builtin/light")
    {
        checkedId = IDM_VIEW_THEME_LIGHT;
    }
    else if (currentThemeId == L"builtin/dark")
    {
        checkedId = IDM_VIEW_THEME_DARK;
    }
    else if (currentThemeId == L"builtin/rainbow")
    {
        checkedId = IDM_VIEW_THEME_RAINBOW;
    }

    CheckMenuRadioItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, IDM_VIEW_THEME_RAINBOW, checkedId, MF_BYCOMMAND);
}

void SaveMonitorSettings(HWND hWnd) noexcept
{
    if (hWnd)
    {
        WINDOWPLACEMENT placement{};
        placement.length = sizeof(placement);
        if (GetWindowPlacement(hWnd, &placement))
        {
            Common::Settings::WindowPlacement wp;
            wp.state = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;

            const RECT rc         = placement.rcNormalPosition;
            wp.bounds.x           = rc.left;
            wp.bounds.y           = rc.top;
            const int savedWidth  = static_cast<int>(rc.right - rc.left);
            const int savedHeight = static_cast<int>(rc.bottom - rc.top);
            wp.bounds.width       = std::max(1, savedWidth);
            wp.bounds.height      = std::max(1, savedHeight);
            wp.dpi                = GetDpiForWindow(hWnd);

            g_settings.windows[kWindowId] = std::move(wp);
        }
    }

    if (! g_settings.monitor)
    {
        g_settings.monitor = Common::Settings::MonitorSettings{};
    }

    g_settings.monitor->menu.toolbarVisible     = g_toolbarVisible;
    g_settings.monitor->menu.lineNumbersVisible = g_colorView.IsLineNumbersEnabled();
    g_settings.monitor->menu.alwaysOnTop        = g_alwaysOnTop;
    g_settings.monitor->menu.showIds            = g_showIds;
    g_settings.monitor->menu.autoScroll         = g_colorView.GetAutoScroll();

    g_settings.monitor->filter.mask   = g_filterMask & kMonitorFilterAllMask;
    g_settings.monitor->filter.preset = PresetFromLegacy(g_lastFilterPreset);

    const HRESULT saveHr = Common::Settings::SaveSettings(kAppId, g_settings);
    if (FAILED(saveHr))
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
        DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
    }
}

LRESULT CALLBACK DxHostWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* host = reinterpret_cast<WindowHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        host     = reinterpret_cast<WindowHost*>(cs->lpCreateParams);
        if (! host || ! host->Attach(hwnd))
        {
            return FALSE;
        }

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(host));
    }

    if (! host)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled         = false;
    const LRESULT result = host->HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    }

    return handled ? result : DefWindowProcW(hwnd, msg, wp, lp);
}

bool RegisterDxHostClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = DxHostWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kDxHostClassName;

    atom = RegisterClassExW(&wc);
    return atom != 0;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text)
{
    return StringUtils::TrimWhitespaceCopy(text);
}

[[nodiscard]] std::wstring StripMenuDecorations(std::wstring_view rawText)
{
    std::wstring cleaned;
    cleaned.reserve(rawText.size());

    bool skipAmpersand = false;
    for (const wchar_t ch : rawText)
    {
        if (ch == L'\t')
        {
            break;
        }

        if (skipAmpersand)
        {
            cleaned.push_back(ch);
            skipAmpersand = false;
            continue;
        }

        if (ch == L'&')
        {
            skipAmpersand = true;
            continue;
        }

        cleaned.push_back(ch);
    }

    return TrimWhitespace(cleaned);
}

[[nodiscard]] std::wstring GetMenuCommandLabel(HWND hWnd, UINT commandId)
{
    HMENU mainMenu = GetMenu(hWnd);
    if (! mainMenu)
    {
        return {};
    }

    std::vector<HMENU> path;
    if (! TryFindMenuPathToCommand(mainMenu, commandId, path) || path.empty())
    {
        return {};
    }

    wchar_t buffer[256]{};
    const int copied = GetMenuStringW(path.back(), commandId, buffer, static_cast<int>(std::size(buffer)), MF_BYCOMMAND);
    if (copied <= 0)
    {
        return {};
    }

    return StripMenuDecorations(std::wstring_view(buffer, static_cast<size_t>(copied)));
}

[[nodiscard]] float MeasureToolbarTextWidthDip(std::wstring_view text, bool toggle, const MonitorChromeMetrics& metrics) noexcept
{
    const float baseWidth = metrics.toolbarTextPaddingDip + (static_cast<float>(text.size()) * metrics.toolbarLabelCharWidthDip);
    return (std::max)(toggle ? metrics.toolbarMinToggleWidthDip : metrics.toolbarMinButtonWidthDip,
                      baseWidth + (toggle ? metrics.toolbarToggleChromeDip : 0.0f));
}

void SyncToolbarState() noexcept
{
    if (g_toolbarShowIdsToggle)
    {
        g_toolbarShowIdsToggle->SetChecked(g_showIds);
    }
}

void LayoutToolbarControls() noexcept
{
    if (! g_toolbarRoot || ! g_hToolbar)
    {
        return;
    }

    RECT client{};
    GetClientRect(g_hToolbar.get(), &client);
    const float widthDip               = g_toolbarDxHost.PixelsToDip(static_cast<float>((std::max)(0L, client.right - client.left)));
    const float heightDip              = g_toolbarDxHost.PixelsToDip(static_cast<float>((std::max)(0L, client.bottom - client.top)));
    const MonitorChromeMetrics metrics = ResolveMonitorChromeMetrics(g_toolbarDxHost.GetTheme().density == RedSalamander::DxUi::Density::Compact);
    g_toolbarRoot->SetBounds(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip));

    const float topDip    = (std::max)(0.0f, (heightDip - metrics.toolbarButtonHeightDip) * 0.5f);
    const float bottomDip = topDip + metrics.toolbarButtonHeightDip;
    float leftDip         = metrics.toolbarPaddingDip;

    const auto layoutControl = [&](RedSalamander::DxUi::Control* control, const float width) noexcept
    {
        if (! control)
        {
            return;
        }

        control->SetBounds(D2D1::RectF(leftDip, topDip, leftDip + width, bottomDip));
        leftDip += width + metrics.toolbarGapDip;
    };

    const auto layoutSeparator = [&](Label* separator) noexcept
    {
        if (! separator)
        {
            return;
        }

        separator->SetBounds(D2D1::RectF(leftDip, topDip, leftDip + metrics.toolbarSeparatorWidthDip, bottomDip));
        leftDip += metrics.toolbarSeparatorWidthDip + metrics.toolbarGapDip;
    };

    const auto children    = g_toolbarRoot->GetChildren();
    Label* firstSeparator  = children.size() > 3u ? dynamic_cast<Label*>(children[3u].get()) : nullptr;
    Label* secondSeparator = children.size() > 5u ? dynamic_cast<Label*>(children[5u].get()) : nullptr;

    layoutControl(g_toolbarNewButton, MeasureToolbarTextWidthDip(g_toolbarNewButton ? g_toolbarNewButton->GetText() : L"", false, metrics));
    layoutControl(g_toolbarOpenButton, MeasureToolbarTextWidthDip(g_toolbarOpenButton ? g_toolbarOpenButton->GetText() : L"", false, metrics));
    layoutControl(g_toolbarSaveButton, MeasureToolbarTextWidthDip(g_toolbarSaveButton ? g_toolbarSaveButton->GetText() : L"", false, metrics));
    layoutSeparator(firstSeparator);
    layoutControl(g_toolbarCopyButton, MeasureToolbarTextWidthDip(g_toolbarCopyButton ? g_toolbarCopyButton->GetText() : L"", false, metrics));
    layoutSeparator(secondSeparator);
    layoutControl(g_toolbarShowIdsToggle, MeasureToolbarTextWidthDip(g_toolbarShowIdsToggle ? g_toolbarShowIdsToggle->GetText() : L"", true, metrics));
}

void CreateToolbarHost(HWND hWnd)
{
    if (! hWnd || g_hToolbar)
    {
        return;
    }

    if (! RegisterDxHostClass(g_hInstance))
    {
        return;
    }

    g_hToolbar.reset(CreateWindowExW(
        0, kDxHostClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 0, 0, hWnd, nullptr, g_hInstance, &g_toolbarDxHost));
    if (! g_hToolbar)
    {
        return;
    }

    auto root     = std::make_unique<Toolbar>();
    g_toolbarRoot = root.get();

    const std::wstring newLabel = GetMenuCommandLabel(hWnd, IDM_FILE_NEW);
    g_toolbarNewButton          = g_toolbarRoot->AddChild<Button>(newLabel.empty() ? L"New" : newLabel);
    g_toolbarNewButton->SetOnClick([hWnd] { SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_FILE_NEW, 0), 0); });

    const std::wstring openLabel = GetMenuCommandLabel(hWnd, IDM_FILE_OPEN);
    g_toolbarOpenButton          = g_toolbarRoot->AddChild<Button>(openLabel.empty() ? L"Open" : openLabel);
    g_toolbarOpenButton->SetOnClick([hWnd] { SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_FILE_OPEN, 0), 0); });

    const std::wstring saveLabel = GetMenuCommandLabel(hWnd, IDM_FILE_SAVE_AS);
    g_toolbarSaveButton          = g_toolbarRoot->AddChild<Button>(saveLabel.empty() ? L"Save As..." : saveLabel);
    g_toolbarSaveButton->SetOnClick([hWnd] { SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_FILE_SAVE_AS, 0), 0); });

    g_toolbarRoot->AddSeparator();

    const std::wstring copyLabel = GetMenuCommandLabel(hWnd, IDM_EDIT_COPY);
    g_toolbarCopyButton          = g_toolbarRoot->AddChild<Button>(copyLabel.empty() ? L"Copy" : copyLabel);
    g_toolbarCopyButton->SetOnClick([hWnd] { SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_EDIT_COPY, 0), 0); });

    g_toolbarRoot->AddSeparator();

    const std::wstring showIdsLabel = GetMenuCommandLabel(hWnd, IDM_OPTION_ID);
    g_toolbarShowIdsToggle          = g_toolbarRoot->AddChild<Toggle>(showIdsLabel.empty() ? L"Show IDs" : showIdsLabel);
    g_toolbarShowIdsToggle->SetOnToggled([hWnd](bool /*checked*/) { SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_ID, 0), 0); });
    g_toolbarShowIdsToggle->SetChecked(g_showIds);

    g_toolbarDxHost.SetRoot(std::move(root));
    LayoutToolbarControls();
    static_cast<void>(g_toolbarDxHost.PrimeForShow());
}

void CreateStatusStripHost(HWND hWnd)
{
    if (! hWnd || g_hStatusBar)
    {
        return;
    }

    if (! RegisterDxHostClass(g_hInstance))
    {
        return;
    }

    g_hStatusBar.reset(CreateWindowExW(
        0, kDxHostClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 0, 0, hWnd, nullptr, g_hInstance, &g_statusDxHost));
    if (! g_hStatusBar)
    {
        return;
    }

    auto strip                         = std::make_unique<StatusStrip>();
    g_statusStrip                      = strip.get();
    const MonitorChromeMetrics metrics = ResolveMonitorChromeMetrics(g_statusDxHost.GetTheme().density == RedSalamander::DxUi::Density::Compact);
    g_statusStrip->SetSections({
        StatusStrip::Section{.text = {}, .widthDip = metrics.statusAutoWidthDip},
        StatusStrip::Section{.text = {}, .widthDip = metrics.statusFilterWidthDip},
        StatusStrip::Section{.text = {}, .widthDip = metrics.statusVisibleWidthDip},
        StatusStrip::Section{.text = {}, .widthDip = metrics.statusTotalWidthDip},
        StatusStrip::Section{.text = {}, .widthDip = 0.0f},
    });
    g_statusDxHost.SetRoot(std::move(strip));
    static_cast<void>(g_statusDxHost.PrimeForShow());
}

void RequestMonitorFileOpenCancellation() noexcept
{
    if (! g_monitorFileOpenThread.joinable())
    {
        return;
    }

    g_monitorFileOpenThread.request_stop();
    if (CancelSynchronousIo(g_monitorFileOpenThread.native_handle()) == FALSE)
    {
        const DWORD error = GetLastError();
        if (error != ERROR_NOT_FOUND)
        {
            Debug::Warning(L"Monitor file open: CancelSynchronousIo failed (gle=0x{:08X})", error);
        }
    }
}

bool DoFileOpen(HWND owner)
{
    if (g_monitorFileOpenActive.load(std::memory_order_acquire))
    {
        RequestMonitorFileOpenCancellation();
        if (g_statusStrip)
        {
            g_statusStrip->SetSectionText(4u, LoadStringResource(g_hInstance, IDS_MSG_OPEN_CANCEL_REQUESTED));
        }
        return false;
    }
    if (g_monitorFileOpenThread.joinable())
    {
        g_monitorFileOpenThread = std::jthread{};
    }

    wchar_t file[MAX_PATH]        = L"";
    const std::wstring filter     = LoadStringResource(g_hInstance, IDS_FILE_FILTER_OPEN);
    const std::wstring defaultExt = LoadStringResource(g_hInstance, IDS_FILE_DEFAULT_EXT_TXT);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFilter = filter.empty() ? nullptr : filter.c_str();
    ofn.lpstrFile   = file;
    ofn.nMaxFile    = static_cast<DWORD>(std::size(file));
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_EXPLORER;
    ofn.lpstrDefExt = defaultExt.empty() ? nullptr : defaultExt.c_str();
    if (! GetOpenFileNameW(&ofn))
        return false;

    const Common::Settings::MonitorRetentionSettings retention =
        g_settings.monitor.value_or(Common::Settings::MonitorSettings{}).retention;
    const RedSalamanderMonitor::MonitorFileReadLimits limits{
        .maxBytes = retention.maxRetainedTextBytes,
        .maxLines = retention.maxRetainedLines,
    };
    const std::filesystem::path path(file);
    const uint64_t generation = ++g_monitorFileOpenGeneration;
    g_monitorFileOpenActive.store(true, std::memory_order_release);
    g_monitorFileOpenThread = std::jthread([owner, path, limits, generation](std::stop_token stopToken)
    {
        const auto started = std::chrono::steady_clock::now();
        uint64_t lastPostedBytes = 0u;
        RedSalamanderMonitor::MonitorFileReadResult result = RedSalamanderMonitor::ReadMonitorTextFile(
            path,
            stopToken,
            limits,
            [owner, generation, &lastPostedBytes](uint64_t bytesRead, uint64_t totalBytes)
            {
                constexpr uint64_t kProgressStepBytes = 1u * 1024u * 1024u;
                if (bytesRead != totalBytes && bytesRead - lastPostedBytes < kProgressStepBytes)
                {
                    return;
                }
                lastPostedBytes = bytesRead;
                auto progress = std::make_unique<MonitorFileOpenProgress>(MonitorFileOpenProgress{
                    .generation = generation,
                    .bytesRead  = bytesRead,
                    .totalBytes = totalBytes,
                });
                static_cast<void>(PostMessagePayload(owner, kMsgMonitorFileOpenProgress, 0, std::move(progress)));
            });

        auto completion = std::make_unique<MonitorFileOpenCompletion>(MonitorFileOpenCompletion{
            .result     = std::move(result),
            .generation = generation,
            .durationUs = Debug::Perf::ElapsedUs(started),
        });
        g_monitorFileOpenActive.store(false, std::memory_order_release);
        static_cast<void>(PostMessagePayload(owner, kMsgMonitorFileOpenCompleted, 0, std::move(completion)));
    });
    return true;
}

LRESULT OnMonitorFileOpenProgress(LPARAM lParam)
{
    auto progress = TakeMessagePayload<MonitorFileOpenProgress>(lParam);
    if (progress && progress->generation == g_monitorFileOpenGeneration && g_statusStrip)
    {
        g_statusStrip->SetSectionText(
            4u, FormatStringResource(g_hInstance, IDS_STATUS_OPEN_PROGRESS_FMT, progress->bytesRead, progress->totalBytes));
    }
    return 0;
}

LRESULT OnMonitorFileOpenCompleted(HWND owner, LPARAM lParam)
{
    auto completion = TakeMessagePayload<MonitorFileOpenCompletion>(lParam);
    if (! completion)
    {
        return 0;
    }
    if (completion->generation != g_monitorFileOpenGeneration)
    {
        return 0;
    }

    Debug::Perf::EmitDurationUs(L"monitor.file_open.total_us",
                                completion->durationUs,
                                completion->result.bytesRead,
                                static_cast<uint64_t>(completion->result.lineCount),
                                completion->result.hr);
    if (completion->result.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        UpdateStatusBar();
        return 0;
    }
    if (FAILED(completion->result.hr))
    {
        Debug::Warning(L"Monitor file open failed before publication: 0x{:08X}", static_cast<uint32_t>(completion->result.hr));
        const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_OPEN);
        const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_OPEN_FAILED_READ);
        ShowModalMessageDialog(g_hInstance, owner, caption.c_str(), message.c_str());
        UpdateStatusBar();
        return 0;
    }

    g_colorView.ClearColoring();
    g_colorView.SetText(completion->result.text);
    UpdateStatusBar();
    return 0;
}

bool DoFileSaveAs(HWND owner)
{
    wchar_t file[MAX_PATH]        = L"";
    const std::wstring filter     = LoadStringResource(g_hInstance, IDS_FILE_FILTER_SAVE);
    const std::wstring defaultExt = LoadStringResource(g_hInstance, IDS_FILE_DEFAULT_EXT_TXT);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFilter = filter.empty() ? nullptr : filter.c_str();
    ofn.lpstrFile   = file;
    ofn.nMaxFile    = static_cast<DWORD>(std::size(file));
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT | OFN_EXPLORER;
    ofn.lpstrDefExt = defaultExt.empty() ? nullptr : defaultExt.c_str();
    if (! GetSaveFileNameW(&ofn))
        return false;

    return g_colorView.SaveTextToFile(file);
}

void AdjustLayout(HWND hWnd)
{
    if (! hWnd)
        return;

    RECT clientRect{};
    GetClientRect(hWnd, &clientRect);

    const MonitorChromeMetrics metrics = ResolveMonitorChromeMetrics(g_toolbarDxHost.GetTheme().density == RedSalamander::DxUi::Density::Compact);
    const int toolbarHeight            = (g_hToolbar && g_toolbarVisible) ? DipsToPx(hWnd, static_cast<int>(std::lround(metrics.toolbarHeightDip))) : 0;
    const int statusBarHeight          = g_hStatusBar ? DipsToPx(hWnd, static_cast<int>(std::lround(metrics.statusStripHeightDip))) : 0;
    const bool toolbarWindowVisible    = g_hToolbar && g_toolbarVisible;

    const int width  = clientRect.right - clientRect.left;
    const int height = clientRect.bottom - clientRect.top;

    if (toolbarWindowVisible)
    {
        MoveWindow(g_hToolbar.get(), 0, 0, width, toolbarHeight, TRUE);
    }
    else if (g_hToolbar)
    {
        MoveWindow(g_hToolbar.get(), 0, 0, width, 0, FALSE);
    }

    const int contentHeight = std::max(0, height - toolbarHeight - statusBarHeight);

    if (g_hColorView)
    {
        MoveWindow(g_hColorView.get(), 0, toolbarHeight, width, contentHeight, TRUE);
    }

    if (g_hStatusBar)
    {
        MoveWindow(g_hStatusBar.get(), 0, toolbarHeight + contentHeight, width, statusBarHeight, TRUE);
    }

    LayoutToolbarControls();
}

// Update status bar with current statistics and synchronize auto-scroll menu state
void UpdateStatusBar()
{
    if (! g_statusStrip)
        return;

    // Synchronize auto-scroll menu checkmark with ColorTextView's auto-scroll state
    // ColorTextView manages its own auto-scroll state and we query it here
    const bool isAutoScrollEnabled = g_colorView.GetAutoScroll();

    HWND hMainWnd = GetParent(g_hStatusBar.get());
    if (hMainWnd)
    {
        HMENU hMenu = GetMenu(hMainWnd);
        if (hMenu)
        {
            // Always sync menu to match actual auto-scroll state (no static variable needed)
            UINT checkState = isAutoScrollEnabled ? MF_CHECKED : MF_UNCHECKED;
            CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, MF_BYCOMMAND | checkState);
        }
    }

    // Get line counts from ColorTextView
    const size_t visibleLines = g_colorView.GetVisibleLineCount();
    const size_t totalLines   = g_colorView.GetTotalLineCount();

    // Get ETW stats (listener-side, not writer-side)
    uint64_t etwReceived = 0;
    if (g_etwListener)
    {
        const auto s = g_etwListener->GetStatistics();
        etwReceived  = static_cast<uint64_t>(s.eventsProcessed);
    }

    // Format status bar text with specific filter names
    std::wstring filterText;
    const uint32_t filterMask = g_filterMask & kMonitorFilterAllMask;
    if (filterMask == kMonitorFilterAllMask)
    {
        filterText = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_ALL);
    }
    else
    {
        // Build filter text showing enabled types
        std::vector<std::wstring> enabledTypeNames;
        for (const auto& entry : kMonitorFilterEntries)
        {
            const uint32_t bit = Debug::FilterBitForType(entry.type);
            if ((filterMask & bit) != 0u)
            {
                enabledTypeNames.push_back(LoadStringResource(g_hInstance, entry.stringId));
            }
        }

        if (enabledTypeNames.empty())
        {
            filterText = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_NONE);
        }
        else if (enabledTypeNames.size() == 1)
        {
            filterText = FormatStringResource(g_hInstance, IDS_STATUS_FILTER_ONE_FMT, enabledTypeNames[0]);
        }
        else
        {
            // Join with "+" separator for multiple types
            filterText                = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_MULTI_PREFIX);
            const std::wstring joiner = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_MULTI_JOINER);
            for (size_t i = 0; i < enabledTypeNames.size(); ++i)
            {
                if (i > 0)
                    filterText += joiner;
                filterText += enabledTypeNames[i];
            }
        }
    }

    const std::wstring autoText    = LoadStringResource(g_hInstance, isAutoScrollEnabled ? IDS_STATUS_AUTOSCROLL_ON : IDS_STATUS_AUTOSCROLL_OFF);
    const std::wstring visibleText = FormatStringResource(g_hInstance, IDS_STATUS_VISIBLE_FMT, visibleLines);
    const std::wstring totalText   = FormatStringResource(g_hInstance, IDS_STATUS_TOTAL_FMT, totalLines);
    const std::wstring etwText =
        FormatStringResource(g_hInstance, IDS_STATUS_ETW_RECEIVED_FMT, etwReceived, g_colorView.GetDroppedEventCount());

    g_statusStrip->SetSectionText(0u, autoText);
    g_statusStrip->SetSectionText(1u, filterText);
    g_statusStrip->SetSectionText(2u, visibleText);
    g_statusStrip->SetSectionText(3u, totalText);
    g_statusStrip->SetSectionText(4u, etwText);

    // Track message count for adaptive refresh
    g_lastMessageCount = etwReceived;
}

[[maybe_unused]] void RedrawMonitorChrome(HWND hWnd)
{
    if (g_hToolbar && ::IsWindowVisible(g_hToolbar.get()))
    {
        ::RedrawWindow(g_hToolbar.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    }
    if (g_hStatusBar && ::IsWindowVisible(g_hStatusBar.get()))
    {
        ::RedrawWindow(g_hStatusBar.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    }
    ::RedrawWindow(hWnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
}

#if defined(ENABLE_TESTS)
std::wstring BuildMonitorEtwBurstMessage(size_t index, size_t payloadChars)
{
    std::wstring message = std::format(L"Monitor selftest ETW batch event {}", index);
    if (message.size() < payloadChars)
    {
        message.append(payloadChars - message.size(), L'x');
    }
    return message;
}

bool PumpMonitorColorViewUntilLineCount(size_t targetLineCount, int maxPumpCount) noexcept
{
    MSG msg{};
    for (int pump = 0; pump < maxPumpCount && g_colorView.GetTotalLineCount() < targetLineCount; ++pump)
    {
        while (::PeekMessageW(&msg, g_hColorView.get(), 0, 0, PM_REMOVE))
        {
            ::TranslateMessage(&msg);
            ::DispatchMessageW(&msg);
        }
        if (g_colorView.GetTotalLineCount() >= targetLineCount)
        {
            break;
        }
        ::Sleep(1);
    }

    if (g_hColorView)
    {
        ::RedrawWindow(g_hColorView.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    }

    return g_colorView.GetTotalLineCount() >= targetLineCount;
}
#endif

LRESULT RunMonitorChromeSelfTest(HWND hWnd)
{
#if ! defined(ENABLE_TESTS)
    if (g_monitorChromeSelfTest.enabled)
    {
        FinalizeMonitorChromeSelfTest(false, L"Monitor DxUI chrome selftest requires ENABLE_TESTS.");
        ::DestroyWindow(hWnd);
    }
    return 0;
#else
    if (! g_monitorChromeSelfTest.enabled)
    {
        return 0;
    }

    const auto require = [&](std::wstring_view name, bool condition, std::wstring_view detail = {}) noexcept
    {
        RecordMonitorChromeSelfTestCheck(name, condition, detail);
        return condition;
    };

    const auto scenarioStarted = std::chrono::steady_clock::now();

    Common::Settings::UiSettings baselineUi = g_settings.ui.value_or(Common::Settings::UiSettings{});
    baselineUi.compactMode                  = false;
    baselineUi.reducedMotion                = Common::Settings::ReducedMotionMode::Off;
    g_settings.ui                           = baselineUi;
    ApplyMonitorTheme();

    AdjustLayout(hWnd);
    SyncToolbarState();
    UpdateStatusBar();
    RedrawMonitorChrome(hWnd);

    Debug::Perf::EmitDurationUs(L"monitor.ui.chrome_ready_us",
                                Debug::Perf::ElapsedUs(scenarioStarted),
                                g_toolbarDxHost.DebugGetRenderCount(),
                                g_statusDxHost.DebugGetRenderCount(),
                                S_OK);

    bool passed = true;
    passed &= require(L"toolbar host created", g_hToolbar && ::IsWindow(g_hToolbar.get()));
    passed &= require(L"status host created", g_hStatusBar && ::IsWindow(g_hStatusBar.get()));
    passed &= require(L"toolbar host visible", g_hToolbar && ::IsWindowVisible(g_hToolbar.get()));
    passed &= require(L"toolbar root type", dynamic_cast<Toolbar*>(g_toolbarDxHost.GetRoot()) == g_toolbarRoot);
    passed &= require(L"status root type", dynamic_cast<StatusStrip*>(g_statusDxHost.GetRoot()) == g_statusStrip);
    passed &= require(L"toolbar rendered", g_toolbarDxHost.DebugGetRenderCount() > 0u);
    passed &= require(L"status rendered", g_statusDxHost.DebugGetRenderCount() > 0u);
    passed &= require(L"toolbar present failures", g_toolbarDxHost.DebugGetPresentFailureCount() == 0u);
    passed &= require(L"status present failures", g_statusDxHost.DebugGetPresentFailureCount() == 0u);
    passed &= require(L"toolbar new label", g_toolbarNewButton && ! g_toolbarNewButton->GetText().empty());
    passed &= require(L"toolbar open label", g_toolbarOpenButton && ! g_toolbarOpenButton->GetText().empty());
    passed &= require(L"toolbar save label", g_toolbarSaveButton && ! g_toolbarSaveButton->GetText().empty());
    passed &= require(L"toolbar copy label", g_toolbarCopyButton && ! g_toolbarCopyButton->GetText().empty());
    passed &= require(L"toolbar show ids toggle sync", g_toolbarShowIdsToggle && g_toolbarShowIdsToggle->IsChecked() == g_showIds);
    passed &= require(L"status section count", g_statusStrip && g_statusStrip->GetSectionCount() == 5u);
    passed &= require(L"status auto section", g_statusStrip && ! g_statusStrip->GetSectionText(0u).empty());
    passed &= require(L"status filter section", g_statusStrip && ! g_statusStrip->GetSectionText(1u).empty());
    passed &= require(L"status visible section", g_statusStrip && ! g_statusStrip->GetSectionText(2u).empty());
    passed &= require(L"status total section", g_statusStrip && ! g_statusStrip->GetSectionText(3u).empty());
    passed &= require(L"status etw section", g_statusStrip && ! g_statusStrip->GetSectionText(4u).empty());

    g_colorView.SetText(L"clear state regression");
    g_colorView.DebugSetSelectionState(2u, 7u, 5u);
    g_colorView.ClearText();
    const auto [clearSelectionStart, clearSelectionEnd, clearCaretPosition] = g_colorView.DebugGetSelectionState();
    passed &= require(L"clear resets selection and caret",
                      clearSelectionStart == 0u && clearSelectionEnd == 0u && clearCaretPosition == 0u && g_colorView.GetTotalLineCount() == 0u);

    const uint64_t droppedBeforeOverload = g_colorView.GetDroppedEventCount();
    g_colorView.SetRetentionLimits(ColorTextView::RetentionLimits{
        .maxQueuedEvents       = 32u,
        .maxRetainedLines      = 24u,
        .maxRetainedTextBytes  = 1u * 1024u * 1024u,
        .maxSearchMatches      = 1'000u,
    });
    for (size_t i = 0u; i < 80u; ++i)
    {
        Debug::InfoParam info{};
        info.processID = ::GetCurrentProcessId();
        info.threadID  = ::GetCurrentThreadId();
        info.type      = Debug::InfoParam::Type::Warning;
        g_colorView.QueueEtwEvent(info, std::format(L"Monitor bounded overload event {}", i));
    }
    MSG overloadMessage{};
    for (int pump = 0; pump < 64 && g_colorView.DebugGetQueuedEventCount() != 0u; ++pump)
    {
        while (::PeekMessageW(&overloadMessage, g_hColorView.get(), 0, 0, PM_REMOVE))
        {
            ::TranslateMessage(&overloadMessage);
            ::DispatchMessageW(&overloadMessage);
        }
        if (g_colorView.DebugGetQueuedEventCount() != 0u)
        {
            ::Sleep(1);
        }
    }
    const uint64_t overloadDropped = g_colorView.GetDroppedEventCount() - droppedBeforeOverload;
    passed &= require(L"etw overload stays bounded",
                      g_colorView.DebugGetQueuedEventCount() == 0u && g_colorView.GetTotalLineCount() == 24u && overloadDropped == 56u,
                      std::format(L"queued={} retained={} dropped={}",
                                  g_colorView.DebugGetQueuedEventCount(),
                                  g_colorView.GetTotalLineCount(),
                                  overloadDropped));
    g_colorView.ClearText();
    const Common::Settings::MonitorRetentionSettings configuredRetention =
        g_settings.monitor.value_or(Common::Settings::MonitorSettings{}).retention;
    g_colorView.SetRetentionLimits(ColorTextView::RetentionLimits{
        .maxQueuedEvents       = configuredRetention.maxQueuedEvents,
        .maxRetainedLines      = configuredRetention.maxRetainedLines,
        .maxRetainedTextBytes  = configuredRetention.maxRetainedTextBytes,
        .maxSearchMatches      = configuredRetention.maxSearchMatches,
    });

    if (passed)
    {
        RECT baselineToolbarRect{};
        RECT baselineStatusRect{};
        GetClientRect(g_hToolbar.get(), &baselineToolbarRect);
        GetClientRect(g_hStatusBar.get(), &baselineStatusRect);
        const float baselineToolbarHeightDip =
            g_toolbarDxHost.PixelsToDip(static_cast<float>((std::max)(0L, baselineToolbarRect.bottom - baselineToolbarRect.top)));
        const float baselineStatusHeightDip =
            g_statusDxHost.PixelsToDip(static_cast<float>((std::max)(0L, baselineStatusRect.bottom - baselineStatusRect.top)));

        const auto compactAppliedStarted       = std::chrono::steady_clock::now();
        Common::Settings::UiSettings compactUi = g_settings.ui.value_or(Common::Settings::UiSettings{});
        compactUi.compactMode                  = true;
        compactUi.reducedMotion                = Common::Settings::ReducedMotionMode::On;
        compactUi.windowBackdrop               = Common::Settings::WindowBackdropMode::MicaAlt;
        g_settings.ui                          = compactUi;
        ApplyMonitorTheme();
        RedrawMonitorChrome(hWnd);

        RECT compactToolbarRect{};
        RECT compactStatusRect{};
        GetClientRect(g_hToolbar.get(), &compactToolbarRect);
        GetClientRect(g_hStatusBar.get(), &compactStatusRect);
        const float compactToolbarHeightDip =
            g_toolbarDxHost.PixelsToDip(static_cast<float>((std::max)(0L, compactToolbarRect.bottom - compactToolbarRect.top)));
        const float compactStatusHeightDip = g_statusDxHost.PixelsToDip(static_cast<float>((std::max)(0L, compactStatusRect.bottom - compactStatusRect.top)));
        const auto appliedBackdrop         = Common::WindowBackdrop::TryGetAppliedWindowBackdropKind(hWnd);

        passed &= require(L"toolbar compact density", g_toolbarDxHost.GetTheme().density == RedSalamander::DxUi::Density::Compact);
        passed &= require(L"status compact density", g_statusDxHost.GetTheme().density == RedSalamander::DxUi::Density::Compact);
        passed &= require(L"toolbar reduced motion override", g_toolbarDxHost.GetTheme().reducedMotion);
        passed &= require(L"status reduced motion override", g_statusDxHost.GetTheme().reducedMotion);
        passed &= require(L"toolbar compact height shrinks",
                          compactToolbarHeightDip < baselineToolbarHeightDip,
                          std::format(L"baseline={} compact={}", baselineToolbarHeightDip, compactToolbarHeightDip));
        passed &= require(L"status compact height shrinks",
                          compactStatusHeightDip < baselineStatusHeightDip,
                          std::format(L"baseline={} compact={}", baselineStatusHeightDip, compactStatusHeightDip));
        passed &= require(L"monitor backdrop applied",
                          appliedBackdrop.has_value() && appliedBackdrop.value() == Common::WindowBackdrop::Kind::MicaAlt,
                          appliedBackdrop.has_value() ? std::format(L"kind={}", static_cast<int>(appliedBackdrop.value())) : L"readback unavailable");
        Debug::Perf::EmitDurationUs(L"monitor.ui.compact_apply_us",
                                    Debug::Perf::ElapsedUs(compactAppliedStarted),
                                    static_cast<uint64_t>(std::lround(compactToolbarHeightDip * 100.0f)),
                                    static_cast<uint64_t>(std::lround(compactStatusHeightDip * 100.0f)),
                                    S_OK);

        const uint64_t toolbarRenderBefore = g_toolbarDxHost.DebugGetRenderCount();
        const auto toolbarToggleStarted    = std::chrono::steady_clock::now();
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_VIEW_TOOLBAR, 0), 0);
        passed &= require(L"toolbar hides via command", g_hToolbar && ! ::IsWindowVisible(g_hToolbar.get()));
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_VIEW_TOOLBAR, 0), 0);
        RedrawMonitorChrome(hWnd);
        passed &= require(L"toolbar re-shows via command", g_hToolbar && ::IsWindowVisible(g_hToolbar.get()));
        passed &= require(L"toolbar rerenders after re-show", g_toolbarDxHost.DebugGetRenderCount() > toolbarRenderBefore);
        Debug::Perf::EmitDurationUs(L"monitor.ui.toolbar_toggle_us",
                                    Debug::Perf::ElapsedUs(toolbarToggleStarted),
                                    g_toolbarDxHost.DebugGetRenderCount() - toolbarRenderBefore,
                                    0u,
                                    S_OK);

        const bool initialShowIds       = g_showIds;
        const auto showIdsToggleStarted = std::chrono::steady_clock::now();
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_ID, 0), 0);
        passed &= require(L"show ids toggled", g_showIds != initialShowIds);
        passed &= require(L"toolbar toggle tracks show ids", g_toolbarShowIdsToggle && g_toolbarShowIdsToggle->IsChecked() == g_showIds);
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_ID, 0), 0);
        passed &= require(L"show ids restored", g_showIds == initialShowIds);
        Debug::Perf::EmitDurationUs(L"monitor.ui.show_ids_toggle_us", Debug::Perf::ElapsedUs(showIdsToggleStarted), g_showIds ? 1u : 0u, 0u, S_OK);

        const std::wstring initialFilterText = g_statusStrip ? std::wstring(g_statusStrip->GetSectionText(1u)) : std::wstring{};
        const auto filterSyncStarted         = std::chrono::steady_clock::now();
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_FILTER_PRESET_ERRORS_ONLY, 0), 0);
        const std::wstring errorsOnlyFilterText = g_statusStrip ? std::wstring(g_statusStrip->GetSectionText(1u)) : std::wstring{};
        passed &= require(L"status filter section changes", ! errorsOnlyFilterText.empty() && errorsOnlyFilterText != initialFilterText);
        ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_FILTER_PRESET_ALL, 0), 0);
        passed &= require(L"status filter section restores", g_statusStrip && std::wstring(g_statusStrip->GetSectionText(1u)) == initialFilterText);
        Debug::Perf::EmitDurationUs(L"monitor.ui.status_filter_sync_us",
                                    Debug::Perf::ElapsedUs(filterSyncStarted),
                                    g_statusStrip ? g_statusStrip->GetSectionCount() : 0u,
                                    g_filterMask,
                                    S_OK);

        constexpr size_t kDefaultSelfTestEtwBurstCount = 260u;
        const size_t etwBurstCount                     = g_monitorEtwBurstOptions.latencyMode ? g_monitorEtwBurstOptions.count : kDefaultSelfTestEtwBurstCount;
        const size_t payloadChars                      = g_monitorEtwBurstOptions.latencyMode ? g_monitorEtwBurstOptions.payloadChars : 0u;
        const size_t linesBeforeBurst                  = g_colorView.GetTotalLineCount();
        const auto etwBurstStarted                     = std::chrono::steady_clock::now();
        if (g_monitorEtwBurstOptions.latencyMode)
        {
            for (size_t i = 0; i < etwBurstCount; ++i)
            {
                Debug::InfoParam info{};
                info.processID = ::GetCurrentProcessId();
                info.threadID  = ::GetCurrentThreadId();
                info.type      = (i % 2u) == 0u ? Debug::InfoParam::Type::Warning : Debug::InfoParam::Type::Error;
                g_colorView.QueueEtwEvent(info, BuildMonitorEtwBurstMessage(i, payloadChars));
                const size_t targetLineCount = linesBeforeBurst + i + 1u;
                passed &= require(L"etw latency sample drained",
                                  PumpMonitorColorViewUntilLineCount(targetLineCount, 64),
                                  std::format(L"sample={} target={}", i, targetLineCount));
            }
        }
        else
        {
            for (size_t i = 0; i < etwBurstCount; ++i)
            {
                Debug::InfoParam info{};
                info.processID = ::GetCurrentProcessId();
                info.threadID  = ::GetCurrentThreadId();
                info.type      = (i % 2u) == 0u ? Debug::InfoParam::Type::Warning : Debug::InfoParam::Type::Error;
                g_colorView.QueueEtwEvent(info, std::format(L"Monitor selftest ETW batch event {}", i));
            }

            MSG msg{};
            for (int pump = 0; pump < 64 && g_colorView.GetTotalLineCount() < linesBeforeBurst + etwBurstCount; ++pump)
            {
                while (::PeekMessageW(&msg, g_hColorView.get(), 0, 0, PM_REMOVE))
                {
                    ::TranslateMessage(&msg);
                    ::DispatchMessageW(&msg);
                }
                if (g_colorView.GetTotalLineCount() >= linesBeforeBurst + etwBurstCount)
                    break;
                ::Sleep(1);
            }

            if (g_hColorView)
            {
                ::RedrawWindow(g_hColorView.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
            }
        }

        UpdateStatusBar();
        const size_t linesAfterBurst = g_colorView.GetTotalLineCount();
        const bool burstDrained      = linesAfterBurst >= linesBeforeBurst + etwBurstCount;
        passed &= require(
            L"etw batch queue drained", burstDrained, std::format(L"before={} after={} expectedDelta={}", linesBeforeBurst, linesAfterBurst, etwBurstCount));
        Debug::Perf::EmitDurationUs(L"monitor.etw.selftest_burst_drain_us",
                                    Debug::Perf::ElapsedUs(etwBurstStarted),
                                    static_cast<uint64_t>(etwBurstCount),
                                    static_cast<uint64_t>(linesAfterBurst - linesBeforeBurst),
                                    burstDrained ? S_OK : E_FAIL);

        if (g_monitorScrollbackSelfTestOptions.enabled)
        {
            constexpr size_t kScrollbackTargetLineCount = 320u;
            if (g_colorView.GetTotalLineCount() < kScrollbackTargetLineCount)
            {
                const size_t linesBeforeFill = g_colorView.GetTotalLineCount();
                const auto fillStarted       = std::chrono::steady_clock::now();
                const size_t fillCount       = kScrollbackTargetLineCount - linesBeforeFill;
                for (size_t i = 0; i < fillCount; ++i)
                {
                    Debug::InfoParam info{};
                    info.processID = ::GetCurrentProcessId();
                    info.threadID  = ::GetCurrentThreadId();
                    info.type      = Debug::InfoParam::Type::Perf;
                    g_colorView.QueueEtwEvent(info, std::format(L"Monitor scrollback selftest fill event {}", i));
                }

                const bool fillDrained = PumpMonitorColorViewUntilLineCount(kScrollbackTargetLineCount, 96);
                passed &= require(L"scrollback fill queue drained",
                                  fillDrained,
                                  std::format(L"before={} after={} target={}", linesBeforeFill, g_colorView.GetTotalLineCount(), kScrollbackTargetLineCount));
                Debug::Perf::EmitDurationUs(L"monitor.etw.selftest_scrollback_fill_us",
                                            Debug::Perf::ElapsedUs(fillStarted),
                                            static_cast<uint64_t>(fillCount),
                                            static_cast<uint64_t>(g_colorView.GetTotalLineCount() - linesBeforeFill),
                                            fillDrained ? S_OK : E_FAIL);
            }

            if (! g_colorView.GetAutoScroll())
            {
                ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_AUTO_SCROLL, 0), 0);
            }

            passed &= require(L"scrollback starts from auto-scroll", g_colorView.GetAutoScroll());

            const auto scrollbackStarted = std::chrono::steady_clock::now();
            ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_AUTO_SCROLL, 0), 0);
            passed &= require(L"scrollback disables auto-scroll via command", ! g_colorView.GetAutoScroll());

            if (g_hColorView)
            {
                ::SendMessageW(g_hColorView.get(), WM_VSCROLL, MAKEWPARAM(SB_PAGEUP, 0), 0);
                ::RedrawWindow(g_hColorView.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
            }

            passed &= require(L"scrollback renders after page-up input", ! g_colorView.GetAutoScroll());

            ::SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(IDM_OPTION_AUTO_SCROLL, 0), 0);
            passed &= require(L"scrollback restores auto-scroll via command", g_colorView.GetAutoScroll());
            Debug::Perf::EmitDurationUs(L"monitor.selftest.scrollback_frame_us",
                                        Debug::Perf::ElapsedUs(scrollbackStarted),
                                        static_cast<uint64_t>(g_colorView.GetTotalLineCount()),
                                        g_colorView.GetAutoScroll() ? 1u : 0u,
                                        g_colorView.GetAutoScroll() ? S_OK : E_FAIL);
        }
    }

    Debug::Perf::EmitValue(L"monitor.ui.toolbar_render_count", g_toolbarDxHost.DebugGetRenderCount(), S_OK);
    Debug::Perf::EmitValue(L"monitor.ui.status_render_count", g_statusDxHost.DebugGetRenderCount(), S_OK);
    Debug::Perf::EmitValue(L"monitor.ui.toolbar_present_failure_count", g_toolbarDxHost.DebugGetPresentFailureCount(), S_OK);
    Debug::Perf::EmitValue(L"monitor.ui.status_present_failure_count", g_statusDxHost.DebugGetPresentFailureCount(), S_OK);
    Debug::Perf::EmitValue(L"monitor.ui.dxhost_attached_count", static_cast<uint64_t>(RedSalamander::DxUi::DebugGetAttachedWindowHostCount()), S_OK);

    FinalizeMonitorChromeSelfTest(passed,
                                  passed ? L"Monitor DxUI toolbar/status strip selftest passed." : L"Monitor DxUI toolbar/status strip selftest failed.");
    ::DestroyWindow(hWnd);
    return 0;
#endif
}
} // namespace

// Forward declarations of functions included in this code module:
std::optional<HWND> InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

// Ensure the process runs Per-Monitor (V2 if available) so GetDpiForWindow returns actual DPI
// this information is also in the manifest
// <dpiAware>true/PM</dpiAware>
// <dpiAwareness> PerMonitorV2</ dpiAwareness>
static void InitializeDpiAwareness()
{
#ifdef _DEBUG
    OutputDebugStringA("=== DPI Awareness Diagnostics ===\n");

    // Check what the thread DPI awareness context reports FIRST
    DPI_AWARENESS_CONTEXT currentContext = GetThreadDpiAwarenessContext();
    if (currentContext)
    {
        // Get the raw pointer value for comparison
        auto rawValue = reinterpret_cast<uintptr_t>(currentContext);
        auto msg      = std::format("Initial context raw value: 0x{:X}\n", rawValue);
        OutputDebugStringA(msg.c_str());

        OutputDebugStringA("Initial thread DPI awareness context: ");

        // Check against all known context values
        if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
            OutputDebugStringA("PER_MONITOR_AWARE_V2: Ok\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE))
            OutputDebugStringA("PER_MONITOR_AWARE: Ok\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_SYSTEM_AWARE))
            OutputDebugStringA("SYSTEM_AWARE\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_UNAWARE))
            OutputDebugStringA("UNAWARE\n");
        else
        {
            // Try to get DPI awareness from the context using the conversion function
            DPI_AWARENESS awareness  = GetAwarenessFromDpiAwarenessContext(currentContext);
            const char* awarenessStr = "Unknown";
            switch (awareness)
            {
                case DPI_AWARENESS_UNAWARE: awarenessStr = "UNAWARE"; break;
                case DPI_AWARENESS_SYSTEM_AWARE: awarenessStr = "SYSTEM_AWARE"; break;
                case DPI_AWARENESS_PER_MONITOR_AWARE: awarenessStr = "PER_MONITOR_AWARE (likely V2 from manifest)"; break;
                case DPI_AWARENESS_INVALID:
                default: awarenessStr = "Invalid"; break;
            }
            auto awarenessMsg = std::format("Decoded awareness: {}: Ok\n", awarenessStr);
            OutputDebugStringA(awarenessMsg.c_str());
        }

        // Also show what the constants are for comparison
        auto v2Value      = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        auto v1Value      = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
        auto systemValue  = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE);
        auto unawareValue = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_UNAWARE);

        auto constantsMsg = std::format("Constants: V2={}, V1={}, SYS={}, UN={}\n", v2Value, v1Value, systemValue, unawareValue);
        OutputDebugStringA(constantsMsg.c_str());
    }

    OutputDebugStringA("Attempting to set DPI awareness programmatically\n");
#endif

    // Try the DPI-specific API first on the supported Windows 11 baseline.
    if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2\n");
#endif
        return;
    }

    auto err = GetLastError();
    if (err == ERROR_ACCESS_DENIED)
    {
#ifdef _DEBUG
        OutputDebugStringA("DPI awareness already set (ACCESS_DENIED) - manifest is working!\n");

        // The legacy API may not report correctly after windows are created
        PROCESS_DPI_AWARENESS awareness = PROCESS_DPI_UNAWARE;
        auto hr                         = GetProcessDpiAwareness(GetCurrentProcess(), &awareness);
        if (SUCCEEDED(hr))
        {
            const char* awarenessStr = "Unknown";
            switch (awareness)
            {
                case PROCESS_DPI_UNAWARE: awarenessStr = "UNAWARE (legacy API - may be inaccurate)"; break;
                case PROCESS_SYSTEM_DPI_AWARE: awarenessStr = "SYSTEM_AWARE"; break;
                case PROCESS_PER_MONITOR_DPI_AWARE: awarenessStr = "PER_MONITOR_AWARE"; break;
            }
            auto msg = std::format("Legacy API reports: {}\n", awarenessStr);
            OutputDebugStringA(msg.c_str());
        }

        // Use the newer, more accurate API to decode the actual awareness level
        DPI_AWARENESS_CONTEXT context = GetThreadDpiAwarenessContext();
        if (context)
        {
            DPI_AWARENESS contextAwareness = GetAwarenessFromDpiAwarenessContext(context);
            OutputDebugStringA("Actual DPI awareness: ");
            switch (contextAwareness)
            {
                case DPI_AWARENESS_UNAWARE: OutputDebugStringA("UNAWARE (Something is wrong)\n"); break;
                case DPI_AWARENESS_SYSTEM_AWARE: OutputDebugStringA("SYSTEM_AWARE\n"); break;
                case DPI_AWARENESS_PER_MONITOR_AWARE: OutputDebugStringA("PER_MONITOR_AWARE (V1 or V2 - Manifest working!)\n"); break;
                case DPI_AWARENESS_INVALID: OutputDebugStringA("INVALID\n"); break;
                default: OutputDebugStringA("Invalid/Unknown\n"); break;
            }
        }

        OutputDebugStringA("=== End DPI Diagnostics ===\n");
#endif
        return; // Success - DPI awareness is already set
    }

    // ... rest of fallback code remains the same
#ifdef _DEBUG
    auto errorMsg = std::format("SetProcessDpiAwarenessContext V2 failed: {}\n", err);
    OutputDebugStringA(errorMsg.c_str());
#endif

    // Fallback to Per-Monitor V1 (Windows 8.1+)
    if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE))
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE\n");
#endif
        return;
    }

    err = GetLastError();
    if (err == ERROR_ACCESS_DENIED)
    {
#ifdef _DEBUG
        OutputDebugStringA("DPI awareness already set (V1 also ACCESS_DENIED)\n");
#endif
        return; // Success - DPI awareness is already set
    }

#ifdef _DEBUG
    auto errorMsg2 = std::format("SetProcessDpiAwarenessContext V1 failed: {}\n", err);
    OutputDebugStringA(errorMsg2.c_str());
#endif

    // Final fallback to legacy API (Windows Vista+)
    if (SetProcessDPIAware())
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set legacy DPI awareness\n");
#endif
        return;
    }

#ifdef _DEBUG
    OutputDebugStringA("All DPI awareness methods failed - this is unexpected\n");
#endif
}

Localization::LanguagePreference GetLanguagePreferenceFromSettings(const Common::Settings::Settings& settings)
{
    Localization::LanguagePreference preference;
    if (! settings.ui || settings.ui->language.empty() || settings.ui->language == L"system")
    {
        preference.kind = Localization::LanguagePreferenceKind::System;
        return preference;
    }

    preference.kind    = Localization::LanguagePreferenceKind::Culture;
    preference.culture = settings.ui->language;
    return preference;
}

// Separate function with C++ objects (cannot use __try/__except)
static int RunApplication(HINSTANCE hInstance, int nCmdShow)
{
    constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";

    constexpr std::wstring_view kWaitInstanceArg = L"--wait-instance";
    constexpr ULONGLONG kWaitInstanceTimeoutMs   = 5000;
    constexpr DWORD kWaitInstancePollMs          = 50;

    if (HasCommandLineArg(L"--help") || HasCommandLineArg(L"-h") || HasCommandLineArg(L"/?"))
    {
        WriteMonitorHelpText(hInstance);
        return 0;
    }

    if (HasCommandLineArg(L"--etw"))
    {
        Debug::detail::SetRuntimeMonitorDiagnosticsEnabled(true);
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_DIAGNOSTICS_ETW", L"1"));
    }

    std::wstring perfJsonlPath;
    const bool customPerfPath = TryGetCommandLineArgValue(L"--perf=", perfJsonlPath);
    if (customPerfPath || HasCommandLineArg(L"--perf"))
    {
        const std::filesystem::path perfPath =
            (customPerfPath && ! perfJsonlPath.empty()) ? std::filesystem::path(perfJsonlPath) : GetDefaultMonitorPerfJsonlPath();
        Debug::Perf::ConfigureJsonlOutput(perfPath, L"RedSalamanderMonitor", kMonitorBuildFlavor);
    }

    wil::unique_handle instanceMutex;
    DWORD mutexCreationError                   = ERROR_SUCCESS;
    g_monitorChromeSelfTest.enabled            = HasCommandLineArg(kMonitorChromeSelfTestArg);
    g_monitorEtwBurstOptions                   = ReadMonitorEtwBurstOptions();
    g_monitorScrollbackSelfTestOptions.enabled = HasCommandLineArg(kMonitorScrollbackSelfTestArg);

    const auto tryCreateInstanceMutex = [&]() -> bool
    {
        instanceMutex.reset(::CreateMutexW(nullptr, FALSE, kInstanceMutexName));
        mutexCreationError = ::GetLastError();
        return static_cast<bool>(instanceMutex);
    };

    if (! tryCreateInstanceMutex())
    {
        const std::wstring caption = LoadStringResource(hInstance, IDS_APP_TITLE);
        const std::wstring message = LoadStringResource(hInstance, IDS_MSG_INSTANCE_GUARD_FAILED);
        ShowModalMessageDialog(hInstance, nullptr, caption.c_str(), message.c_str());
        return g_monitorChromeSelfTest.enabled ? 1 : FALSE;
    }

    if (mutexCreationError == ERROR_ALREADY_EXISTS)
    {
        if (! HasCommandLineArg(kWaitInstanceArg))
        {
            OutputDebugStringW(L"RedSalamander Monitor is already running.");
            return g_monitorChromeSelfTest.enabled ? 1 : FALSE;
        }

        instanceMutex.reset();

        const ULONGLONG startTick = ::GetTickCount64();
        while (true)
        {
            ::Sleep(kWaitInstancePollMs);

            if (! tryCreateInstanceMutex())
            {
                const std::wstring caption = LoadStringResource(hInstance, IDS_APP_TITLE);
                const std::wstring message = LoadStringResource(hInstance, IDS_MSG_INSTANCE_GUARD_FAILED);
                ShowModalMessageDialog(hInstance, nullptr, caption.c_str(), message.c_str());
                return g_monitorChromeSelfTest.enabled ? 1 : FALSE;
            }

            if (mutexCreationError != ERROR_ALREADY_EXISTS)
            {
                break;
            }

            instanceMutex.reset();

            if ((::GetTickCount64() - startTick) > kWaitInstanceTimeoutMs)
            {
                OutputDebugStringW(L"Timed out waiting for previous instance to exit.");
                return g_monitorChromeSelfTest.enabled ? 1 : FALSE;
            }
        }
    }

    // Set DPI awareness before creating any windows
    InitializeDpiAwareness();

    g_config.Load();

    const std::filesystem::path themesDir = GetThemesDirectory();
    if (! themesDir.empty())
    {
        Common::Settings::LoadThemeDefinitionsFromDirectory(themesDir, g_fileThemes);
    }

    Common::Settings::LoadSettings(kAppId, g_settings);
    Localization::RegisterResourceOwner(kAppId, hInstance);
    Localization::ApplyLanguagePreference(GetLanguagePreferenceFromSettings(g_settings));

    if (g_settings.monitor)
    {
        g_toolbarVisible     = g_settings.monitor->menu.toolbarVisible;
        g_lineNumbersVisible = g_settings.monitor->menu.lineNumbersVisible;
        g_alwaysOnTop        = g_settings.monitor->menu.alwaysOnTop;
        g_showIds            = g_settings.monitor->menu.showIds;
        g_autoScrollEnabled  = g_settings.monitor->menu.autoScroll;

        g_filterMask       = g_settings.monitor->filter.mask & kMonitorFilterAllMask;
        g_lastFilterPreset = LegacyFromPreset(g_settings.monitor->filter.preset);
        if (g_lastFilterPreset == -1)
        {
            const int inferred = InferLegacyPresetFromMask(g_filterMask);
            if (inferred != -1)
                g_lastFilterPreset = inferred;
        }
    }
    else
    {
        // Migration/defaults: prefer existing registry values for filter settings if present.
        g_filterMask       = g_config.filterMask & kMonitorFilterAllMask;
        g_lastFilterPreset = g_config.lastFilterPreset;
        if (g_lastFilterPreset == -1)
        {
            const int inferred = InferLegacyPresetFromMask(g_filterMask);
            if (inferred != -1)
                g_lastFilterPreset = inferred;
        }
        else
        {
            switch (g_lastFilterPreset)
            {
                case 0: g_filterMask = kMonitorPresetErrorsOnlyMask; break;
                case 1: g_filterMask = kMonitorPresetErrorsWarningsMask; break;
                case 2: g_filterMask = kMonitorFilterAllMask; break;
                case 3: g_filterMask = kMonitorPresetErrorsPerfDebugMask; break;
                default: break;
            }
        }

        Common::Settings::MonitorSettings monitorSettings;
        monitorSettings.menu.toolbarVisible     = g_toolbarVisible;
        monitorSettings.menu.lineNumbersVisible = g_lineNumbersVisible;
        monitorSettings.menu.alwaysOnTop        = g_alwaysOnTop;
        monitorSettings.menu.showIds            = g_showIds;
        monitorSettings.menu.autoScroll         = g_autoScrollEnabled;
        monitorSettings.filter.mask             = g_filterMask & kMonitorFilterAllMask;
        monitorSettings.filter.preset           = PresetFromLegacy(g_lastFilterPreset);
        g_settings.monitor                      = std::move(monitorSettings);
    }

    if (g_monitorChromeSelfTest.enabled)
    {
        g_toolbarVisible                = true;
        g_alwaysOnTop                   = false;
        g_lineNumbersVisible            = true;
        g_showIds                       = true;
        g_autoScrollEnabled             = true;
        g_filterMask                    = Debug::InfoParam::Type::All;
        g_lastFilterPreset              = 2;
        g_settings.theme.currentThemeId = L"builtin/light";
        g_settings.ui                   = Common::Settings::UiSettings{};

        if (! InitializeMonitorChromeSelfTestArtifacts())
        {
            return 1;
        }
    }

    // Perform application initialization:
    auto hWnd = InitInstance(hInstance, nCmdShow);
    if (! hWnd)
    {
        if (g_monitorChromeSelfTest.enabled)
        {
            FinalizeMonitorChromeSelfTest(false, L"InitInstance failed before selftest window creation.");
        }
        return g_monitorChromeSelfTest.enabled ? 1 : FALSE;
    }

    wil::unique_haccel hAccelTable(Localization::LoadAcceleratorsResource(hInstance, MAKEINTRESOURCEW(IDC_REDSALAMANDERMONITOR)));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (! TranslateAccelerator(hWnd.value(), hAccelTable.get(), &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    if (g_monitorChromeSelfTest.enabled && ! g_monitorChromeSelfTest.completed)
    {
        FinalizeMonitorChromeSelfTest(false, L"Monitor selftest ended before completion.");
    }

    return static_cast<int>(msg.wParam);
}

namespace
{
void BuildFatalExceptionMessage(HINSTANCE hInstance, const wchar_t* exceptionName, DWORD exceptionCode, wchar_t* outMessage, size_t outMessageChars) noexcept
{
    if (! outMessage || outMessageChars == 0)
    {
        return;
    }
    outMessage[0] = L'\0';

    const std::wstring msg = FormatStringResource(hInstance, IDS_FATAL_EXCEPTION_FMT, exceptionName, static_cast<unsigned>(exceptionCode));
    if (! msg.empty())
    {
        static_cast<void>(wcsncpy_s(outMessage, outMessageChars, msg.c_str(), _TRUNCATE));
        return;
    }

    const auto max = std::min<std::size_t>(outMessageChars - 1, static_cast<std::size_t>(std::numeric_limits<std::ptrdiff_t>::max()));
    const auto r   = std::format_to_n(
        outMessage, static_cast<std::ptrdiff_t>(max), L"Fatal Exception ({0}, 0x{1:08X}).", exceptionName, static_cast<unsigned>(exceptionCode));
    const std::ptrdiff_t cap                                       = static_cast<std::ptrdiff_t>(max);
    const std::ptrdiff_t written                                   = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
    outMessage[(written <= 0) ? 0u : static_cast<size_t>(written)] = L'\0';
}
} // namespace

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE /*hPrevInstance*/, _In_ LPWSTR /*lpCmdLine*/, _In_ int nCmdShow)
{
    if (! Common::MinimumOsVersion::EnsureCurrentWindowsVersionSupported(nullptr))
    {
        return 1;
    }

    // Use SEH to catch all exceptions (no C++ objects in this scope)
    __try
    {
        return RunApplication(hInstance, nCmdShow);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);

        wchar_t errorMsg[512]{};
        BuildFatalExceptionMessage(hInstance, exceptionName, exceptionCode, errorMsg, std::size(errorMsg));
        OutputDebugStringW(errorMsg);

        wchar_t caption[256]{};
        const int captionLength = LoadStringW(hInstance, IDS_FATAL_ERROR_CAPTION, caption, static_cast<int>(sizeof(caption) / sizeof(caption[0])));
        ShowModalMessageDialog(hInstance, nullptr, captionLength > 0 ? caption : L"", errorMsg);

        return -1;
    }
}

// Saves instance handle and creates main window
std::optional<HWND> InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    g_hInstance = hInstance; // Store instance handle in our global variable

    WNDCLASSEXW wcex{};
    wcex.cbSize        = sizeof(WNDCLASSEX);
    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.cbClsExtra    = 0;
    wcex.cbWndExtra    = 0;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REDSALAMANDERMONITOR));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = nullptr;
    wcex.lpszClassName = g_redSalamanderMonitorClassName;
    wcex.hIconSm       = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    auto atom = RegisterClassExW(&wcex);
    if (! atom)
    {
        Debug::ErrorWithLastError(L"RedSalamanderMonitor: RegisterClassExW failed");
        return std::nullopt;
    }

    wchar_t title[256]{};
    const int titleLength = LoadStringW(hInstance, IDS_APP_TITLE, title, static_cast<int>(sizeof(title) / sizeof(title[0])));
    PCWSTR windowTitle    = titleLength > 0 ? title : g_redSalamanderMonitor;

    wil::unique_hwnd hWnd(CreateWindowExW(WS_EX_NOACTIVATE,
                                          g_redSalamanderMonitorClassName,
                                          windowTitle,
                                          WS_OVERLAPPEDWINDOW,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          640,
                                          480,
                                          nullptr,
                                          nullptr,
                                          hInstance,
                                          nullptr));
    if (! hWnd)
    {
        Debug::ErrorWithLastError(L"RedSalamanderMonitor: CreateWindowExW failed");
        return std::nullopt;
    }

    // Remove WS_EX_NOACTIVATE after creation
    SetWindowLongPtr(hWnd.get(), GWL_EXSTYLE, GetWindowLongPtr(hWnd.get(), GWL_EXSTYLE) & ~WS_EX_NOACTIVATE);
    g_hMainWindow = hWnd.get();

    wil::unique_hmenu menu(Localization::LoadMenuResource(hInstance, IDC_REDSALAMANDERMONITOR));
    if (menu && SetMenu(hWnd.get(), menu.get()) != FALSE)
    {
        menu.release();
    }

    int showCmd   = nCmdShow;
    const auto it = g_settings.windows.find(kWindowId);
    if (it != g_settings.windows.end())
    {
        const UINT dpi                                     = GetDpiForWindow(hWnd.get());
        const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, dpi);

        SetWindowPos(
            hWnd.get(), nullptr, normalized.bounds.x, normalized.bounds.y, normalized.bounds.width, normalized.bounds.height, SWP_NOZORDER | SWP_NOACTIVATE);

        showCmd = normalized.state == Common::Settings::WindowState::Maximized ? SW_MAXIMIZE : SW_SHOWNORMAL;
    }

    ShowWindow(hWnd.get(), showCmd);
    UpdateWindow(hWnd.get());

    return hWnd.release();
}

static std::wstring NormalizeLineEndings(const std::wstring& input)
{
    if (input.empty())
    {
        return L"\n";
    }

    std::wstring result;
    result.reserve(input.length() + 1); // Reserve space to avoid reallocations

    for (size_t i = 0; i < input.length(); ++i)
    {
        if (input[i] == L'\r')
        {
            // Check if this is \r\n sequence
            if (i + 1 < input.length() && input[i + 1] == L'\n')
            {
                result += L'\n';
                ++i; // Skip the \n as we've handled the \r\n pair
            }
            else
            {
                result += L'\n'; // Convert standalone \r to \n
            }
        }
        else
        {
            result += input[i];
        }
    }

    // Ensure text ends with exactly one \n
    while (! result.empty() && result.back() == L'\n')
    {
        result.pop_back();
    }
    result += L'\n';

    return result;
}

void AddLine(PCWSTR line)
{
    // OutputDebugStringW(std::format(L"-- {}\n", line).c_str());

    if (! g_hColorView)
        return;

    std::wstring text = NormalizeLineEndings(line);
    g_colorView.AppendText(text); // ColorTextView handles scroll-to-bottom internally
}

namespace
{
LRESULT OnCreateMainWindow(HWND hWnd)
{
    InitPostedPayloadWindow(hWnd);
    g_hColorView.reset(g_colorView.Create(hWnd, 0, 0, 0, 0));
    if (! g_hColorView)
    {
        const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_ERROR);
        const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_CREATE_COLORTEXTVIEW_FAILED);
        ShowModalMessageDialog(g_hInstance, hWnd, caption.c_str(), message.c_str());
        return -1;
    }

    CreateToolbarHost(hWnd);
    CreateStatusStripHost(hWnd);

    if (RedSalamanderMonitor::ShouldDisplayInitialMonitorStatus())
    {
        const std::wstring sampleText = LoadStringResource(g_hInstance, IDS_SAMPLE_TEXT);
        g_colorView.SetText(sampleText);
        g_colorView.ColorizeWord(L"ColorTextView", D2D1::ColorF(D2D1::ColorF::Blue));
        g_colorView.ColorizeWord(L"right", D2D1::ColorF(D2D1::ColorF::Green));
    }

    g_colorView.EnableShowIds(g_showIds);
    g_colorView.EnableLineNumbers(g_lineNumbersVisible);
    g_colorView.SetAutoScroll(g_autoScrollEnabled);
    g_colorView.SetFilterMask(g_filterMask);
    const Common::Settings::MonitorRetentionSettings retention =
        g_settings.monitor.value_or(Common::Settings::MonitorSettings{}).retention;
    g_colorView.SetRetentionLimits(ColorTextView::RetentionLimits{
        .maxQueuedEvents       = retention.maxQueuedEvents,
        .maxRetainedLines      = retention.maxRetainedLines,
        .maxRetainedTextBytes  = retention.maxRetainedTextBytes,
        .maxSearchMatches      = retention.maxSearchMatches,
    });
    ApplyMonitorTheme();

    if (g_hToolbar)
    {
        ShowWindow(g_hToolbar.get(), g_toolbarVisible ? SW_SHOW : SW_HIDE);
    }

    if (g_alwaysOnTop)
    {
        SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    }

    HMENU hMenu = GetMenu(hWnd);
    if (hMenu)
    {
        CheckMenuItem(hMenu, IDM_VIEW_TOOLBAR, static_cast<UINT>(MF_BYCOMMAND | (g_toolbarVisible ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_OPTION_TOP, static_cast<UINT>(MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_OPTION_ID, static_cast<UINT>(MF_BYCOMMAND | (g_showIds ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (g_colorView.IsLineNumbersEnabled() ? MF_CHECKED : MF_UNCHECKED)));

        SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

        CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, static_cast<UINT>(MF_BYCOMMAND | (g_colorView.GetAutoScroll() ? MF_CHECKED : MF_UNCHECKED)));

        RebuildThemeMenuDynamicItems(hWnd);
        UpdateThemeMenuChecks();
    }

    AdjustLayout(hWnd);
    UpdateStatusBar();

    if (! g_monitorChromeSelfTest.enabled)
    {
        g_etwListener                = std::make_unique<EtwListener>();
        const DWORD monitorProcessId = GetCurrentProcessId();
        const bool etwStarted        = g_etwListener->Start([monitorProcessId](const Debug::InfoParam& info, const std::wstring& message)
        {
            if (! RedSalamanderMonitor::ShouldAcceptEtwEventForDisplay(info, monitorProcessId))
            {
                return;
            }

            std::wstring normalizedMsg = message;
            while (! normalizedMsg.empty() && (normalizedMsg.back() == L'\n' || normalizedMsg.back() == L'\r'))
            {
                normalizedMsg.pop_back();
            }

            g_colorView.QueueEtwEvent(info, std::move(normalizedMsg));
        });

        if (! etwStarted)
        {
            constexpr std::wstring_view kWaitInstanceArg = L"--wait-instance";

            const ULONG etwErrorCode = g_etwListener ? g_etwListener->GetLastErrorCode() : ERROR_SUCCESS;
            if (etwErrorCode == ERROR_ACCESS_DENIED && ! IsProcessElevated())
            {
                const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_ETW_WARNING);
                const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_ETW_ELEVATE_PROMPT);

                const bool elevateNow = ShowModalConfirmDialog(g_hInstance, hWnd, caption.empty() ? L"" : caption.c_str(), message.c_str());
                if (elevateNow && RelaunchSelfElevated(hWnd, kWaitInstanceArg))
                {
                    return -1;
                }
            }

            std::wstring errorMsg = FormatStringResource(g_hInstance, IDS_FMT_ETW_START_FAILED, g_etwListener->GetLastError());
            AddLine(errorMsg.c_str());

#ifdef _DEBUG
            OutputDebugString(errorMsg.c_str());
            OutputDebugStringA("\n");
#endif
        }
        else
        {
            if (RedSalamanderMonitor::ShouldDisplayInitialMonitorStatus())
            {
                const std::wstring startedText = LoadStringResource(g_hInstance, IDS_MSG_ETW_STARTED);
                AddLine(startedText.empty() ? L"" : startedText.c_str());
            }
        }
    }

    SetTimer(hWnd, kStatusBarTimerId, kStatusBarUpdateIntervalMs, nullptr);
    if (g_monitorChromeSelfTest.enabled)
    {
        PostMessageW(hWnd, kMsgRunMonitorChromeSelfTest, 0, 0);
    }
    return 0;
}

LRESULT OnTimerMainWindow(HWND hWnd, UINT_PTR timerId)
{
    if (timerId != kStatusBarTimerId)
    {
        return DefWindowProcW(hWnd, WM_TIMER, static_cast<WPARAM>(timerId), 0);
    }

    uint64_t currentCount = 0;
    if (g_etwListener)
    {
        const auto s = g_etwListener->GetStatistics();
        currentCount = static_cast<uint64_t>(s.eventsProcessed);
    }

    if (currentCount != g_lastMessageCount || (currentCount % 10 == 0))
    {
        UpdateStatusBar();
    }

    return 0;
}

LRESULT OnSizeMainWindow(HWND hWnd, UINT /*width*/, UINT /*height*/)
{
    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnDpiChangedMainWindow(HWND hWnd, [[maybe_unused]] UINT newDpi, const RECT* suggestedRect)
{
    if (g_hToolbar)
    {
        SendMessageW(g_hToolbar.get(), WM_DPICHANGED, static_cast<WPARAM>(static_cast<ULONG>(MAKELONG(newDpi, newDpi))), 0);
        ShowWindow(g_hToolbar.get(), g_toolbarVisible ? SW_SHOW : SW_HIDE);
    }
    if (g_hStatusBar)
    {
        SendMessageW(g_hStatusBar.get(), WM_DPICHANGED, static_cast<WPARAM>(static_cast<ULONG>(MAKELONG(newDpi, newDpi))), 0);
    }

    if (suggestedRect != nullptr)
    {
        SetWindowPos(hWnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnSystemThemeChangedMainWindow([[maybe_unused]] HWND hWnd)
{
    LocaleFormatting::InvalidateFormatLocaleCache();
    ApplyMonitorTheme();
    UpdateThemeMenuChecks();
    return 0;
}

LRESULT OnCommandMainWindow(HWND hWnd, UINT id, UINT codeNotify, HWND hwndCtl)
{
    switch (id)
    {
        case IDM_ABOUT:
#pragma warning(suppress : 5039)
            DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, reinterpret_cast<DLGPROC>(About));
            break;
        case IDM_EXIT: DestroyWindow(hWnd); break;
        case IDM_FILE_NEW:
            g_colorView.ClearColoring();
            g_colorView.ClearText();
            break;
        case IDM_FILE_OPEN: DoFileOpen(hWnd); break;
        case IDM_FILE_SAVE_AS: DoFileSaveAs(hWnd); break;
        case IDM_EDIT_COPY: g_colorView.CopySelection(); break;
        case IDM_EDIT_FIND: g_colorView.ShowFind(); break;
        case IDM_EDIT_FIND_NEXT: g_colorView.FindNext(false); break;
        case IDM_EDIT_FIND_PREV: g_colorView.FindNext(true); break;
        case IDM_FILE_PRINT:
        {
            const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_PRINT);
            const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_PRINT_NOT_IMPLEMENTED);
            ShowModalMessageDialog(g_hInstance, hWnd, caption.c_str(), message.c_str());
            break;
        }
        case IDM_VIEW_LINE_NUMBERS:
            if (g_hColorView)
            {
                const bool currentState = g_colorView.IsLineNumbersEnabled();
                g_colorView.EnableLineNumbers(! currentState);
                g_lineNumbersVisible = ! currentState;

                HMENU hMenu = GetMenu(hWnd);
                CheckMenuItem(hMenu, IDM_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (currentState ? MF_UNCHECKED : MF_CHECKED)));
            }
            break;
        case IDM_VIEW_THEME_SYSTEM:
            g_settings.theme.currentThemeId = L"builtin/system";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_LIGHT:
            g_settings.theme.currentThemeId = L"builtin/light";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_DARK:
            g_settings.theme.currentThemeId = L"builtin/dark";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_RAINBOW:
            g_settings.theme.currentThemeId = L"builtin/rainbow";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_HIGH_CONTRAST_APP:
            g_settings.theme.currentThemeId = L"builtin/highContrast";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_OPTION_TOP:
        {
            g_alwaysOnTop = ! g_alwaysOnTop;
            SetWindowPos(hWnd, g_alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_TOP, static_cast<UINT>(MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED)));
            break;
        }
        case IDM_OPTION_ID:
        {
            g_showIds   = ! g_showIds;
            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_ID, static_cast<UINT>(MF_BYCOMMAND | (g_showIds ? MF_CHECKED : MF_UNCHECKED)));
            g_colorView.EnableShowIds(g_showIds);
            SyncToolbarState();
            break;
        }
        case IDM_OPTION_AUTO_SCROLL:
        {
            // From the End-key accelerator we treat this as "go to end and follow" (force enable),
            // not as a toggle.
            const bool invokedByAccelerator = (codeNotify == 1);
            const bool newState             = invokedByAccelerator ? true : ! g_colorView.GetAutoScroll();

            if (invokedByAccelerator)
            {
                g_colorView.GoToEnd(true);
            }
            else
            {
                g_colorView.SetAutoScroll(newState);
            }

            g_autoScrollEnabled = newState;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, static_cast<UINT>(MF_BYCOMMAND | (newState ? MF_CHECKED : MF_UNCHECKED)));
            UpdateStatusBar();
            break;
        }
        case IDM_VIEW_TOOLBAR:
        {
            HMENU hMenu            = GetMenu(hWnd);
            const bool nextVisible = ! g_toolbarVisible;
            if (g_hToolbar)
            {
                ShowWindow(g_hToolbar.get(), nextVisible ? SW_SHOW : SW_HIDE);
            }
            g_toolbarVisible = nextVisible;
            CheckMenuItem(hMenu, IDM_VIEW_TOOLBAR, static_cast<UINT>(MF_BYCOMMAND | (nextVisible ? MF_CHECKED : MF_UNCHECKED)));
            AdjustLayout(hWnd);
            break;
        }
        case IDM_FILTER_TEXT:
        case IDM_FILTER_ERROR:
        case IDM_FILTER_WARNING:
        case IDM_FILTER_INFO:
        case IDM_FILTER_PERF:
        case IDM_FILTER_DEBUG:
        {
            const MonitorFilterUiEntry* entry = FindMonitorFilterEntry(id);
            if (! entry)
            {
                break;
            }

            const uint32_t bitMask = Debug::FilterBitForType(entry->type);

            g_filterMask ^= bitMask;

            HMENU hMenu = GetMenu(hWnd);
            SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

            g_lastFilterPreset = -1;

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_ONLY:
        {
            g_filterMask       = kMonitorPresetErrorsOnlyMask;
            g_lastFilterPreset = 0;

            HMENU hMenu = GetMenu(hWnd);
            SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_WARNINGS:
        {
            g_filterMask       = kMonitorPresetErrorsWarningsMask;
            g_lastFilterPreset = 1;

            HMENU hMenu = GetMenu(hWnd);
            SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_DEBUG:
        {
            g_filterMask       = kMonitorPresetErrorsPerfDebugMask;
            g_lastFilterPreset = 3;

            HMENU hMenu = GetMenu(hWnd);
            SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ALL:
        {
            g_filterMask       = kMonitorFilterAllMask;
            g_lastFilterPreset = 2;

            HMENU hMenu = GetMenu(hWnd);
            SyncMonitorFilterMenuChecks(hMenu, g_filterMask);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        default:
        {
            const auto it = g_customThemeMenuIdToThemeId.find(id);
            if (it != g_customThemeMenuIdToThemeId.end())
            {
                g_settings.theme.currentThemeId = it->second;
                ApplyMonitorTheme();
                UpdateThemeMenuChecks();
                break;
            }

            return DefWindowProcW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(id), static_cast<WORD>(codeNotify)), reinterpret_cast<LPARAM>(hwndCtl));
        }
    }

    return 0;
}

LRESULT OnPaintMainWindow(HWND hWnd)
{
    wil::unique_hdc_paint paint_dc = wil::BeginPaint(hWnd);
    return 0;
}

LRESULT OnDestroyMainWindow(HWND hWnd)
{
    if (! g_monitorChromeSelfTest.enabled)
    {
        SaveMonitorSettings(hWnd);
    }
    KillTimer(hWnd, kStatusBarTimerId);

    // IMPORTANT: Shutdown order matters for thread safety.
    // 1. Stop ETW listener first (stops worker thread that calls QueueEtwEvent on g_colorView)
    // 2. Then destroy the color view (safe because no more cross-thread PostMessage calls)
    // Reversing this order risks use-after-free: worker thread could PostMessage to destroyed HWND.
    if (g_etwListener)
    {
        g_etwListener->Stop();
        g_etwListener.reset();
    }

    if (g_monitorFileOpenThread.joinable())
    {
        RequestMonitorFileOpenCancellation();
        g_monitorFileOpenThread = std::jthread{};
    }
    ++g_monitorFileOpenGeneration;
    g_monitorFileOpenActive.store(false, std::memory_order_release);

    g_hColorView.reset();
    if (g_hToolbar)
    {
        g_toolbarDxHost.Detach();
        g_hToolbar.reset();
    }
    if (g_hStatusBar)
    {
        g_statusDxHost.Detach();
        g_hStatusBar.reset();
    }

    g_toolbarRoot          = nullptr;
    g_toolbarNewButton     = nullptr;
    g_toolbarOpenButton    = nullptr;
    g_toolbarSaveButton    = nullptr;
    g_toolbarCopyButton    = nullptr;
    g_toolbarShowIdsToggle = nullptr;
    g_statusStrip          = nullptr;
    g_hMainWindow          = nullptr;

    PostQuitMessage(g_monitorChromeSelfTest.enabled ? g_monitorChromeSelfTest.exitCode : 0);
    return 0;
}
} // namespace

// Processes messages for the main window.
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_CREATE: return OnCreateMainWindow(hWnd);
        case WM_TIMER: return OnTimerMainWindow(hWnd, static_cast<UINT_PTR>(wParam));
        case WM_SIZE: return OnSizeMainWindow(hWnd, LOWORD(lParam), HIWORD(lParam));
        case WM_DPICHANGED: return OnDpiChangedMainWindow(hWnd, static_cast<UINT>(HIWORD(wParam)), reinterpret_cast<const RECT*>(lParam));
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hWnd, *info, 640, 360);
            }
            return 0;
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE: return OnSystemThemeChangedMainWindow(hWnd);
        case kMsgRunMonitorChromeSelfTest: return RunMonitorChromeSelfTest(hWnd);
        case kMsgMonitorFileOpenProgress: return OnMonitorFileOpenProgress(lParam);
        case kMsgMonitorFileOpenCompleted: return OnMonitorFileOpenCompleted(hWnd, lParam);
        case WM_COMMAND: return OnCommandMainWindow(hWnd, LOWORD(wParam), HIWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WM_PAINT: return OnPaintMainWindow(hWnd);
        case WM_DESTROY: return OnDestroyMainWindow(hWnd);
        case WM_NCDESTROY:
            static_cast<void>(DrainPostedPayloadsForWindow(hWnd));
            return DefWindowProc(hWnd, message, wParam, lParam);
        default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, [[maybe_unused]] LPARAM lParam)
{
    switch (message)
    {
        case WM_INITDIALOG: SetDlgItemTextW(hDlg, IDC_ABOUT_VERSION, VERSINFO_VERSION_LABEL); return static_cast<INT_PTR>(TRUE);

        case WM_COMMAND:
            if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
            {
                EndDialog(hDlg, LOWORD(wParam));
                return static_cast<INT_PTR>(TRUE);
            }
            break;
    }
    return static_cast<INT_PTR>(FALSE);
}

#include "ViewerPE.h"
#include "HandleIo.h"

#include "LocalizationManager.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <format>
#include <functional>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <shobjidl.h>
#include <uxtheme.h>

#define _PEPARSE_WINDOWS_CONFLICTS
#include <pe-parse/parse.h>

#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "uxtheme")

#include "DxUi/DxUi.Typography.h"
#include "Helpers.h"
#include "ViewerFileComboHost.h"
#include "ViewerTitleBarTheme.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace Typography = RedSalamander::DxUi::Typography;

using RedSalamander::DxUi::ColorFromArgb;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::MakeThemePaletteFromViewerTheme;

namespace
{
constexpr UINT kAsyncParseCompleteMessage             = WndMsg::kViewerPeAsyncParseComplete;
constexpr UINT_PTR kAsyncParseFallbackTimerId         = 0x5045u;
constexpr UINT kAsyncParseFallbackPollMs              = 100u;
constexpr uint64_t kMaxPeFileBytes                    = 256ull * 1024ull * 1024ull;
constexpr wchar_t kFileComboHostOriginalWndProcProp[] = L"RS.ViewerPE.FileComboHostOriginalWndProc";
constexpr wchar_t kFileComboHostStateProp[]           = L"RS.ViewerPE.FileComboHostState";
constexpr char kViewerPeModuleAnchor                  = 0;
std::atomic_uint64_t g_nextViewerPeWindowIdentity{1u};

enum class AsyncParseInjectedFault : unsigned int
{
    None = 0u,
    ResultAllocation,
    PayloadPost,
};

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

void UnhookFileComboHostWindow(HWND hwnd) noexcept
{
    RedSalamander::ViewerFileComboHost::UnhookFileComboHostWindow(hwnd, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    return RedSalamander::ViewerFileComboHost::DispatchFileComboHostWndProc<ViewerPE>(
        hwnd, msg, wp, lp, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

constexpr float kOuterPaddingDip    = 12.0f;
constexpr float kCardRadiusDip      = 10.0f;
constexpr float kInnerPaddingDip    = 12.0f;
constexpr float kHeaderGapDip       = 4.0f;
constexpr float kScrollWheelStepDip = 48.0f;
constexpr float kScrollLineStepDip  = 24.0f;

constexpr char kViewerPESchemaJson[] = R"json({
  "version": 1,
  "title": "PE Viewer",
  "fields": []
})json";

[[nodiscard]] const char* GetViewerPEStaticConfigurationSchemaImpl() noexcept
{
    return kViewerPESchemaJson;
}

constexpr std::array<std::wstring_view, 16> kDataDirectoryNames = {
    std::wstring_view(L"Export"),
    std::wstring_view(L"Import"),
    std::wstring_view(L"Resource"),
    std::wstring_view(L"Exception"),
    std::wstring_view(L"Security"),
    std::wstring_view(L"Base Relocation"),
    std::wstring_view(L"Debug"),
    std::wstring_view(L"Architecture"),
    std::wstring_view(L"GlobalPtr"),
    std::wstring_view(L"TLS"),
    std::wstring_view(L"Load Config"),
    std::wstring_view(L"Bound Import"),
    std::wstring_view(L"IAT"),
    std::wstring_view(L"Delay Import"),
    std::wstring_view(L"COM Descriptor"),
    std::wstring_view(L"Reserved"),
};

using Common::Colors::BlendColorRefTruncate;
using Common::Colors::ColorRefFromArgb;
using Common::Colors::ColorRefFromHsvClampedNegativeHueToZero;
using Common::Colors::StableVisualHash32Utf16V1;

[[nodiscard]] COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableVisualHash32Utf16V1(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorRefFromHsvClampedNegativeHueToZero(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

[[nodiscard]] D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

enum class ExportFormat
{
    Text,
    Markdown,
};

[[nodiscard]] std::optional<std::filesystem::path> ShowExportSaveDialog(HWND owner, const std::wstring& defaultFileName, ExportFormat format) noexcept
{
    if (! owner)
    {
        return std::nullopt;
    }

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(coinitHr))
    {
        Debug::Error(L"ViewerPE export: CoInitializeEx(COINIT_APARTMENTTHREADED) failed: 0x{:08X}", coinitHr);
        FAIL_FAST_IF_FAILED(coinitHr);
    }
    [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT;
    static_cast<void>(dialog->SetOptions(options));

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_DIALOG_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    if (! defaultFileName.empty())
    {
        static_cast<void>(dialog->SetFileName(defaultFileName.c_str()));
    }

    COMDLG_FILTERSPEC spec{};
    if (format == ExportFormat::Markdown)
    {
        const std::wstring name = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_FILTER_MARKDOWN);
        spec.pszName            = name.c_str();
        spec.pszSpec            = L"*.md";
        static_cast<void>(dialog->SetDefaultExtension(L"md"));
        static_cast<void>(dialog->SetFileTypes(1, &spec));
    }
    else
    {
        const std::wstring name = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_FILTER_TEXT);
        spec.pszName            = name.c_str();
        spec.pszSpec            = L"*.txt";
        static_cast<void>(dialog->SetDefaultExtension(L"txt"));
        static_cast<void>(dialog->SetFileTypes(1, &spec));
    }

    const HRESULT showHr = dialog->Show(owner);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    return std::filesystem::path(path.get());
}

[[nodiscard]] HRESULT WriteUtf8FileWithBom(const std::filesystem::path& path, std::wstring_view content) noexcept
{
    const std::string utf8 = Utf8FromUtf16(content);
    if (utf8.empty() && ! content.empty())
    {
        return E_FAIL;
    }

    constexpr size_t kBomSize = 3;
    if (utf8.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()) - kBomSize)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    wil::unique_handle outFile(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! outFile)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    static constexpr BYTE kBom[kBomSize] = {0xEF, 0xBB, 0xBF};
    if (const HRESULT writeHr = Common::HandleIo::WriteAll(outFile.get(), kBom, kBomSize); FAILED(writeHr))
    {
        return writeHr;
    }

    if (! utf8.empty())
    {
        if (const HRESULT writeHr = Common::HandleIo::WriteAll(outFile.get(), utf8.data(), utf8.size()); FAILED(writeHr))
        {
            return writeHr;
        }
    }

    return S_OK;
}
[[nodiscard]] float ClampScroll(float scrollDip, float contentHeightDip, float viewportHeightDip) noexcept
{
    const float maxScroll = std::max(0.0f, contentHeightDip - viewportHeightDip);
    return std::clamp(scrollDip, 0.0f, maxScroll);
}

struct ParsedPeDeleter
{
    void operator()(peparse::parsed_pe* pe) const noexcept
    {
        if (pe)
        {
            peparse::DestructParsedPE(pe);
        }
    }
};

[[nodiscard]] std::wstring MachineText(peparse::parsed_pe* pe) noexcept
{
    if (! pe)
    {
        return {};
    }
    if (const char* s = peparse::GetMachineAsString(pe); s && s[0] != '\0')
    {
        return Utf16FromUtf8(s);
    }
    return {};
}

[[nodiscard]] std::wstring SubsystemText(peparse::parsed_pe* pe) noexcept
{
    if (! pe)
    {
        return {};
    }
    if (const char* s = peparse::GetSubsystemAsString(pe); s && s[0] != '\0')
    {
        return Utf16FromUtf8(s);
    }
    return {};
}

[[nodiscard]] std::wstring_view PeKindName(std::uint16_t optionalMagic) noexcept
{
    switch (optionalMagic)
    {
        case 0x10Bu: return L"PE32";
        case 0x20Bu: return L"PE32+";
        default: return L"PE";
    }
}
} // namespace

struct ViewerPE::AsyncParseScheduler final
{
    AsyncParseScheduler()                                      = default;
    AsyncParseScheduler(const AsyncParseScheduler&)            = delete;
    AsyncParseScheduler(AsyncParseScheduler&&)                 = delete;
    AsyncParseScheduler& operator=(const AsyncParseScheduler&) = delete;
    AsyncParseScheduler& operator=(AsyncParseScheduler&&)      = delete;
    ~AsyncParseScheduler()                                     = default;

    struct WorkItem final
    {
        uint64_t requestId = 0u;
        std::function<void()> run;
    };

    struct WorkerContext final
    {
        WorkerContext()                                    = default;
        WorkerContext(const WorkerContext&)                = delete;
        WorkerContext(WorkerContext&&) noexcept            = default;
        WorkerContext& operator=(const WorkerContext&)     = delete;
        WorkerContext& operator=(WorkerContext&&) noexcept = default;
        ~WorkerContext()                                   = default;

        std::shared_ptr<AsyncParseScheduler> scheduler;
        wil::unique_hmodule moduleKeepAlive;
    };

    std::mutex mutex;
    std::optional<WorkItem> pending;
    bool workerRunning = false;
    std::atomic_uint64_t latestRequestId{0u};
    std::atomic_uint64_t latestWindowIdentity{0u};
    std::atomic_uint64_t terminalFallbackRequestId{0u};
    std::atomic_uint64_t terminalFallbackWindowIdentity{0u};
    std::atomic_long terminalFallbackHr{E_FAIL};
};

const char* GetViewerPEStaticConfigurationSchema() noexcept
{
    return GetViewerPEStaticConfigurationSchemaImpl();
}

ViewerPE::ViewerPE() : _parseScheduler(std::make_shared<AsyncParseScheduler>())
{
    _metaId          = L"builtin/viewer-pe";
    _metaShortId     = L"pe";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERPE_DESCRIPTION);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    _configurationJson = "{}";
}

ViewerPE::~ViewerPE()
{
    CancelAsyncParse();

    if (! _hFileComboHost)
    {
        return;
    }

    _fileComboControl            = nullptr;
    _fileComboHostPreExpandPopup = false;
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.reset();
}

void ViewerPE::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerPE::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerPE::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerPE::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerPE::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerPEStaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson = configurationJsonUtf8;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    *pSomethingToSave = FALSE;
    return S_OK;
}

ATOM ViewerPE::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.hInstance     = instance;
    wc.lpszClassName = kClassName;
    wc.lpfnWndProc   = WndProcThunk;
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK ViewerPE::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        if (cs && cs->lpCreateParams)
        {
            auto* self            = static_cast<ViewerPE*>(cs->lpCreateParams);
            self->_windowIdentity = g_nextViewerPeWindowIdentity.fetch_add(1u, std::memory_order_relaxed);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    auto* self = reinterpret_cast<ViewerPE*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerPE::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
#ifdef ENABLE_TESTS
        case WndMsg::kViewerPeDebugGetSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerPeDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                = {};
            snapshot->isLoading      = _isLoading;
            snapshot->requestId      = _parseRequestId.load(std::memory_order_acquire);
            snapshot->windowIdentity = _windowIdentity;
            snapshot->parseHr        = _lastParseHr;
            snapshot->bodyLength     = _bodyText.size();
            const size_t copyLength  = (std::min)(_bodyText.size(), std::size(snapshot->bodyPreview) - 1u);
            std::copy_n(_bodyText.data(), copyLength, snapshot->bodyPreview);
            return TRUE;
        }
        case WndMsg::kViewerPeDebugReloadWithAsyncFault:
        {
            const auto fault = static_cast<WndMsg::ViewerPeDebugAsyncFault>(wp);
            if (fault != WndMsg::ViewerPeDebugAsyncFault::ResultAllocation && fault != WndMsg::ViewerPeDebugAsyncFault::PayloadPost)
            {
                return FALSE;
            }
            _debugNextAsyncFault = fault == WndMsg::ViewerPeDebugAsyncFault::ResultAllocation
                                       ? static_cast<unsigned int>(AsyncParseInjectedFault::ResultAllocation)
                                       : static_cast<unsigned int>(AsyncParseInjectedFault::PayloadPost);
            StartAsyncParse(hwnd, _fileSystem, _currentPath);
            return TRUE;
        }
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_DPICHANGED: OnDpiChanged(static_cast<UINT>(LOWORD(wp)), *reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
            }
            return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp)); return 0;
        case WM_CONTEXTMENU: OnContextMenu(hwnd, RedSalamander::DxUi::ResolveNativeContextMenuScreenPoint(hwnd, lp)); return 0;
        case WM_SYSKEYDOWN:
            if ((wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            break;
        case WM_PAINT: OnPaint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_MOUSEWHEEL: OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wp)); return 0;
        case WM_VSCROLL: OnVScroll(LOWORD(wp), HIWORD(wp)); return 0;
        case WM_KEYDOWN: OnKeyDown(static_cast<UINT>(wp)); return 0;
        case WM_TIMER:
            if (wp == kAsyncParseFallbackTimerId)
            {
                PollAsyncParseTerminalFallback();
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_NCACTIVATE: ApplyTitleBarTheme(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_CLOSE: static_cast<void>(Close()); return 0;
        case WM_NCDESTROY:
        {
            CancelAsyncParse();
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            ResetDeviceResources();
            _menuBarHost.Detach();
            _menuHandle.reset();
            UnhookFileComboHostWindow(_hFileComboHost.get());
            _fileComboControl            = nullptr;
            _fileComboHostPreExpandPopup = false;
            _fileComboHost.Detach();
            _hFileComboHost.release();
            _hWnd.release();
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

            NotifyViewerClosed();

            Release();
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case kAsyncParseCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncParseResult>(lp);
            OnAsyncParseComplete(std::move(result));
            return 0;
        }
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerPE::HandleFileComboHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    const bool popupWasOpen      = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    const bool preExpandForPopup = ! popupWasOpen && _fileComboControl && RedSalamander::ViewerFileComboHost::MessageMayOpenWindowComboPopup(msg, wp);
    if (preExpandForPopup)
    {
        _fileComboHostPreExpandPopup = true;
        if (_hWnd)
        {
            Layout(_hWnd.get());
        }
    }

    const LRESULT dxResult = _fileComboHost.HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        handled = true;
        _fileComboHost.ReleaseMouseCapture();
        _fileComboControl            = nullptr;
        _fileComboHostPreExpandPopup = false;
        _hFileComboHost.release();
        return dxResult;
    }

    const bool popupIsOpen = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    if (popupIsOpen != popupWasOpen || (preExpandForPopup && ! popupIsOpen))
    {
        _fileComboHostPreExpandPopup = false;
        if (_hWnd)
        {
            Layout(_hWnd.get());
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    return dxResult;
}

void ViewerPE::FocusMainSurfaceFromFileCombo(HWND hwnd) noexcept
{
    if (_embeddedMode)
    {
        return;
    }

    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);

    const DWORD comboHostStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY;
    _hFileComboHost.reset(CreateWindowExW(
        0, L"Static", L"", comboHostStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERPE_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileComboHost)
    {
        Debug::ErrorWithLastError(L"ViewerPE: CreateWindowExW failed for DxUi file combo host.");
    }
    else if (! _fileComboHost.Attach(_hFileComboHost.get()))
    {
        Debug::Error(L"ViewerPE: failed to attach DxUi host for file combo.");
        _hFileComboHost.reset();
    }
    else if (! RedSalamander::ViewerFileComboHost::InstallFileComboHostWindow(
                 _hFileComboHost.get(), this, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc))
    {
        Debug::ErrorWithLastError(L"ViewerPE: failed to install WNDPROC hook for DxUi file combo host.");
        _fileComboHost.Detach();
        _hFileComboHost.reset();
    }
    else
    {
        auto combo        = std::make_unique<ComboBox>();
        _fileComboControl = combo.get();
        _fileComboControl->SetVariant(ComboBoxVariant::Window);
        _fileComboControl->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncingFileCombo || selectedIndex >= _otherFiles.size() || ! _hWnd)
            {
                return;
            }

            const HWND hwnd = _hWnd.get();
            _otherIndex     = selectedIndex;
            _currentPath    = _otherFiles[_otherIndex];
            StartAsyncParse(hwnd, _fileSystem, _currentPath);
            UpdateMenuState(hwnd);
            if (! _embeddedMode)
            {
                SetFocus(hwnd);
            }
        });
        RedSalamander::ViewerFileComboHost::ConfigureFileComboKeyboard(_fileComboHost,
                                                                       [this]() noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
        });
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _fileComboHost.SetRoot(std::move(combo));
    }

    if (! _menuHandle)
    {
        _menuHandle.reset(GetMenu(hwnd));
    }
    if (_menuHandle)
    {
        _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this]
        {
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get(), false);
            }
        });
        _menuBarHost.SetOnTabBoundary([this](bool) noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
            return true;
        });
        _menuBarHost.SetOnEscape([this]() noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
            return true;
        });
        static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));
    }

    if (! _embeddedMode)
    {
        ApplyTitleBarTheme(true);
    }
    RefreshFileCombo(hwnd);
    Layout(hwnd);
}

void ViewerPE::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    if (_renderTarget)
    {
        RECT rc{};
        if (GetClientRect(_hWnd.get(), &rc) != 0)
        {
            const D2D1_SIZE_U size = D2D1::SizeU(static_cast<UINT>(std::max(1l, rc.right - rc.left)), static_cast<UINT>(std::max(1l, rc.bottom - rc.top)));
            _renderTarget->Resize(size);
        }
    }

    if (_hWnd)
    {
        Layout(_hWnd.get());
    }

    _textLayout.reset();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::OnDpiChanged(UINT dpi, const RECT& suggestedRect) noexcept
{
    _dpi = dpi;

    if (_hWnd)
    {
        SetWindowPos(_hWnd.get(),
                     nullptr,
                     suggestedRect.left,
                     suggestedRect.top,
                     suggestedRect.right - suggestedRect.left,
                     suggestedRect.bottom - suggestedRect.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    _textLayout.reset();
    if (_renderTarget)
    {
        _renderTarget->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));
    }

    if (_hWnd)
    {
        Layout(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::OnPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);

    EnsureDeviceResources(hwnd);
    if (! _renderTarget || ! _textBrush || ! _bgBrush || ! _cardBrush || ! _cardBorderBrush)
    {
        return;
    }

    const D2D1_SIZE_F size = _renderTarget->GetSize();

    const float cardLeft         = kOuterPaddingDip;
    const float menuBarHeightDip = (_menuBarHost.GetHwnd() != nullptr)
                                       ? (static_cast<float>(_menuBarHost.GetVisibleHeightPx()) * 96.0f / static_cast<float>(std::max<UINT>(1u, _dpi)))
                                       : 0.0f;
    const float cardTop          = menuBarHeightDip + kOuterPaddingDip;
    const float cardRight        = std::max(cardLeft + 1.0f, size.width - kOuterPaddingDip);
    const float cardBottom       = std::max(cardTop + 1.0f, size.height - kOuterPaddingDip);
    const D2D1_ROUNDED_RECT card{D2D1::RectF(cardLeft, cardTop, cardRight, cardBottom), kCardRadiusDip, kCardRadiusDip};

    const float contentLeft       = cardLeft + kInnerPaddingDip;
    const float contentTop        = cardTop + kInnerPaddingDip + std::max(0.0f, _headerHeightDip);
    const float contentRight      = std::max(contentLeft + 1.0f, cardRight - kInnerPaddingDip);
    const float contentBottom     = std::max(contentTop + 1.0f, cardBottom - kInnerPaddingDip);
    const float contentWidthDip   = std::max(1.0f, contentRight - contentLeft);
    const float viewportHeightDip = std::max(1.0f, contentBottom - contentTop);

    EnsureTextLayout(contentWidthDip, viewportHeightDip);
    UpdateScrollBars(hwnd, viewportHeightDip);

    const bool themed         = _hasTheme && ! _theme.highContrast;
    const COLORREF clearColor = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);

    _renderTarget->BeginDraw();
    _renderTarget->Clear(themed ? ColorFromArgb(_theme.backgroundArgb) : ColorFFromColorRef(clearColor));

    _renderTarget->FillRoundedRectangle(card, _cardBrush.get());
    _renderTarget->DrawRoundedRectangle(card, _cardBorderBrush.get(), 1.0f);

    const D2D1_RECT_F clipRc = D2D1::RectF(contentLeft, contentTop, contentRight, contentBottom);
    _renderTarget->PushAxisAlignedClip(clipRc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

    if (_textLayout)
    {
        const D2D1_POINT_2F origin = D2D1::Point2F(contentLeft, contentTop - _scrollDip);
        _renderTarget->DrawTextLayout(origin, _textLayout.get(), _textBrush.get());
    }

    _renderTarget->PopAxisAlignedClip();

    const HRESULT hr = _renderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        _textBrush.reset();
        _cardBorderBrush.reset();
        _cardBrush.reset();
        _bgBrush.reset();
        _renderTarget.reset();
    }
}

void ViewerPE::OnMouseWheel(short delta) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const float steps = static_cast<float>(delta) / static_cast<float>(WHEEL_DELTA);
    ScrollByDip(_hWnd.get(), -steps * kScrollWheelStepDip);
}

void ViewerPE::OnVScroll(WORD request, WORD /*position*/) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (GetScrollInfo(_hWnd.get(), SB_VERT, &si) == 0)
    {
        return;
    }

    int newPos = si.nPos;
    switch (request)
    {
        case SB_TOP: newPos = si.nMin; break;
        case SB_BOTTOM: newPos = si.nMax; break;
        case SB_LINEUP: newPos = si.nPos - static_cast<int>(MulDiv(static_cast<int>(std::lround(kScrollLineStepDip)), static_cast<int>(_dpi), 96)); break;
        case SB_LINEDOWN: newPos = si.nPos + static_cast<int>(MulDiv(static_cast<int>(std::lround(kScrollLineStepDip)), static_cast<int>(_dpi), 96)); break;
        case SB_PAGEUP: newPos = si.nPos - static_cast<int>(si.nPage); break;
        case SB_PAGEDOWN: newPos = si.nPos + static_cast<int>(si.nPage); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: newPos = si.nTrackPos; break;
        default: return;
    }

    const float scrollDip = static_cast<float>(newPos) * 96.0f / static_cast<float>(_dpi);
    SetScrollDip(_hWnd.get(), scrollDip);
}

void ViewerPE::OnKeyDown(UINT vk) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
        return;
    }

    if (vk == VK_F5)
    {
        CommandRefresh(_hWnd.get());
        return;
    }

    if (vk == VK_SPACE)
    {
        CommandOtherNext(_hWnd.get());
        return;
    }

    if (vk == VK_BACK)
    {
        CommandOtherPrevious(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_UP)
    {
        CommandOtherPrevious(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_DOWN)
    {
        CommandOtherNext(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_HOME)
    {
        CommandOtherFirst(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_END)
    {
        CommandOtherLast(_hWnd.get());
        return;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        const UINT cmd = shift ? IDM_VIEWERPE_FILE_EXPORT_MARKDOWN : IDM_VIEWERPE_FILE_EXPORT_TEXT;
        SendMessageW(_hWnd.get(), WM_COMMAND, static_cast<WPARAM>(cmd), 0);
        return;
    }

    switch (vk)
    {
        case VK_UP: ScrollByDip(_hWnd.get(), -kScrollLineStepDip); break;
        case VK_DOWN: ScrollByDip(_hWnd.get(), kScrollLineStepDip); break;
        case VK_PRIOR: ScrollByDip(_hWnd.get(), -std::max(1.0f, _viewportHeightDip * 0.9f)); break;
        case VK_NEXT: ScrollByDip(_hWnd.get(), std::max(1.0f, _viewportHeightDip * 0.9f)); break;
        case VK_HOME: SetScrollDip(_hWnd.get(), 0.0f); break;
        case VK_END: SetScrollDip(_hWnd.get(), std::numeric_limits<float>::infinity()); break;
        default: break;
    }
}

void ViewerPE::OnCommand(HWND hwnd, UINT commandId, [[maybe_unused]] UINT notifyCode, [[maybe_unused]] HWND control) noexcept
{
    if (! hwnd)
    {
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWERPE_FILE_EXPORT_TEXT: CommandExportText(hwnd); break;
        case IDM_VIEWERPE_FILE_EXPORT_MARKDOWN: CommandExportMarkdown(hwnd); break;
        case IDM_VIEWERPE_FILE_REFRESH: CommandRefresh(hwnd); break;
        case IDM_VIEWERPE_OTHER_NEXT: CommandOtherNext(hwnd); break;
        case IDM_VIEWERPE_OTHER_PREVIOUS: CommandOtherPrevious(hwnd); break;
        case IDM_VIEWERPE_OTHER_FIRST: CommandOtherFirst(hwnd); break;
        case IDM_VIEWERPE_OTHER_LAST: CommandOtherLast(hwnd); break;
        case IDM_VIEWERPE_VIEW_GOTO_TOP: SetScrollDip(hwnd, 0.0f); break;
        case IDM_VIEWERPE_VIEW_GOTO_BOTTOM: SetScrollDip(hwnd, std::numeric_limits<float>::infinity()); break;
        case IDM_VIEWERPE_FILE_EXIT: CommandExit(); break;
        default: break;
    }
}

void ViewerPE::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
    if (! _menuHandle)
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERPE_MENU));
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    UpdateMenuState(hwnd, false);
    static constexpr std::array<int, 5> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWERPE_FILE_EXIT,
        IDM_VIEWERPE_OTHER_NEXT,
        IDM_VIEWERPE_OTHER_PREVIOUS,
        IDM_VIEWERPE_OTHER_FIRST,
        IDM_VIEWERPE_OTHER_LAST,
    }};
    RedSalamander::DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    const auto result = RedSalamander::DxUi::ShowNativeHMenuContextMenu(
        hwnd, screenPt, menu, _hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false), previewMenuOptions);
    if (result.has_value() && result.value() > 0)
    {
        OnCommand(hwnd, static_cast<UINT>(result.value()), 0, nullptr);
    }
}

void ViewerPE::Layout(HWND hwnd) noexcept
{
    if (! hwnd || ! _hFileComboHost)
    {
        _headerHeightDip = 0.0f;
        return;
    }

    RECT client{};
    GetClientRect(hwnd, &client);
    _menuBarHost.UpdateLayout();

    const UINT dpi            = GetDpiForWindow(hwnd);
    const int menuBarHeightPx = _menuBarHost.GetHwnd() ? _menuBarHost.GetVisibleHeightPx() : 0;
    const bool showCombo      = (! _embeddedMode && _otherFiles.size() > 1);
    const int outerPaddingPx  = Common::WindowSizing::DipToPixelRounded(kOuterPaddingDip, dpi);
    const int innerPaddingPx  = Common::WindowSizing::DipToPixelRounded(kInnerPaddingDip, dpi);

    const int cardLeft  = outerPaddingPx;
    const int cardTop   = menuBarHeightPx + outerPaddingPx;
    const int cardRight = std::max(cardLeft + 1, static_cast<int>(client.right) - outerPaddingPx);

    const int contentLeft  = cardLeft + innerPaddingPx;
    const int contentRight = std::max(contentLeft + 1, cardRight - innerPaddingPx);

    ShowWindow(_hFileComboHost.get(), showCombo ? SW_SHOW : SW_HIDE);
    EnableWindow(_hFileComboHost.get(), showCombo ? TRUE : FALSE);
    if (_fileComboControl)
    {
        _fileComboControl->SetEnabled(showCombo);
        _fileComboControl->SetVisible(showCombo);
    }
    if (! showCombo)
    {
        _fileComboHostPreExpandPopup = false;
    }

    float newHeaderHeightDip = 0.0f;
    if (showCombo)
    {
        const int comboHeight      = std::max(1L, Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboHeightDip));
        const int comboX           = contentLeft;
        const int comboW           = std::max(1, contentRight - contentLeft);
        const int comboY           = cardTop + innerPaddingPx;
        const bool expandPopupHost = _fileComboHostPreExpandPopup || (_fileComboControl && _fileComboControl->DebugIsPopupOpen());
        const int hostHeight =
            comboHeight + (expandPopupHost ? RedSalamander::ViewerFileComboHost::ComputeStandaloneComboPopupHeightPx(_otherFiles.size(), dpi) : 0);

        SetWindowPos(_hFileComboHost.get(), HWND_TOP, comboX, comboY, comboW, hostHeight, SWP_NOACTIVATE);
        if (_fileComboControl)
        {
            _fileComboControl->SetBounds(D2D1::RectF(0.0f,
                                                     0.0f,
                                                     Common::WindowSizing::PixelToDip(static_cast<float>(comboW), static_cast<float>(dpi)),
                                                     Common::WindowSizing::PixelToDip(static_cast<float>(comboHeight), static_cast<float>(dpi))));
            _fileComboHost.Invalidate();
        }

        newHeaderHeightDip = Common::WindowSizing::PixelToDip(static_cast<float>(comboHeight), static_cast<float>(dpi)) + kHeaderGapDip;
    }

    if (std::fabs(newHeaderHeightDip - _headerHeightDip) > 0.25f)
    {
        _headerHeightDip = newHeaderHeightDip;
        _textLayout.reset();
    }
}

void ViewerPE::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    if (_otherFiles.size() <= 1)
    {
        _fileComboControl->SetItems({});
        _fileComboControl->SetSelectedIndex(std::nullopt);
        if (hwnd)
        {
            Layout(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return;
    }

    std::vector<ComboBox::Item> items;
    items.reserve(_otherFiles.size());
    for (const auto& path : _otherFiles)
    {
        std::wstring itemText = std::filesystem::path(path).filename().wstring();
        if (itemText.empty())
        {
            itemText = path;
        }
        items.push_back(ComboBox::Item{path, std::move(itemText)});
    }
    _fileComboControl->SetItems(std::move(items));

    if (_otherIndex >= _otherFiles.size())
    {
        _otherIndex = 0;
    }

    _fileComboControl->SetSelectedIndex(_otherIndex);
    _fileComboHost.Invalidate();

    if (hwnd)
    {
        Layout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerPE::SyncFileComboSelection() noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    if (_otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex >= _otherFiles.size())
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    _fileComboControl->SetSelectedIndex(_otherIndex);
    _fileComboHost.Invalidate();
}

void ViewerPE::UpdateMenuState([[maybe_unused]] HWND hwnd, bool syncDxMenuBar) noexcept
{
    if (! hwnd)
    {
        return;
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    const bool hasOther   = _otherFiles.size() > 1;
    const UINT otherState = static_cast<UINT>(MF_BYCOMMAND | (hasOther ? MF_ENABLED : MF_GRAYED));

    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_NEXT, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_PREVIOUS, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_FIRST, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_LAST, otherState);

    const bool canRefresh = ! _currentPath.empty();
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_REFRESH, static_cast<UINT>(MF_BYCOMMAND | (canRefresh ? MF_ENABLED : MF_GRAYED)));

    const bool canExport = ! _isLoading && (! _subtitleText.empty() || ! _bodyText.empty());
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_EXPORT_TEXT, static_cast<UINT>(MF_BYCOMMAND | (canExport ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_EXPORT_MARKDOWN, static_cast<UINT>(MF_BYCOMMAND | (canExport ? MF_ENABLED : MF_GRAYED)));

    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else
    {
        DrawMenuBar(hwnd);
    }
}

void ViewerPE::CommandExit() noexcept
{
    static_cast<void>(Close());
}

void ViewerPE::CommandRefresh(HWND hwnd) noexcept
{
    if (! hwnd || _currentPath.empty())
    {
        return;
    }

    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::CommandOtherNext(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.size() <= 1)
    {
        return;
    }

    _otherIndex  = (_otherIndex + 1) % _otherFiles.size();
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::CommandOtherPrevious(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex == 0)
    {
        _otherIndex = _otherFiles.size() - 1;
    }
    else
    {
        _otherIndex -= 1;
    }

    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::CommandOtherFirst(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.empty())
    {
        return;
    }

    _otherIndex  = 0;
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::CommandOtherLast(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.empty())
    {
        return;
    }

    _otherIndex  = _otherFiles.size() - 1;
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }
}

void ViewerPE::CommandExportText(HWND hwnd) noexcept
{
    if (! hwnd || _isLoading)
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring content;
    content = std::format(L"{}\n{}\n\n{}", _titleText, _subtitleText, _bodyText);
    if (content.empty())
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring defaultName;
    defaultName     = _titleText.empty() ? L"pe-report.pe.txt" : (_titleText + L".pe.txt");
    const auto dest = ShowExportSaveDialog(hwnd, defaultName, ExportFormat::Text);
    if (! dest.has_value())
    {
        return;
    }

    const HRESULT hr = WriteUtf8FileWithBom(dest.value(), content);
    if (FAILED(hr))
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_FAILED, IDS_VIEWERPE_NAME, MB_OK | MB_ICONERROR);
    }
}

void ViewerPE::CommandExportMarkdown(HWND hwnd) noexcept
{
    if (! hwnd || _isLoading)
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring content;
    content = _markdownText;
    if (content.empty())
    {
        const std::wstring mdTitle = _titleText.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME) : _titleText;
        content                    = std::format(L"# {}\n\n{}\n\n```text\n{}\n```\n", mdTitle, _subtitleText, _bodyText);
    }

    if (content.empty())
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring defaultName;
    defaultName     = _titleText.empty() ? L"pe-report.pe.md" : (_titleText + L".pe.md");
    const auto dest = ShowExportSaveDialog(hwnd, defaultName, ExportFormat::Markdown);
    if (! dest.has_value())
    {
        return;
    }

    const HRESULT hr = WriteUtf8FileWithBom(dest.value(), content);
    if (FAILED(hr))
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_FAILED, IDS_VIEWERPE_NAME, MB_OK | MB_ICONERROR);
    }
}

void ViewerPE::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    RedSalamander::ViewerChrome::ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive, L"title");
}

void ViewerPE::ResetDeviceResources() noexcept
{
    _textLayout.reset();
    _baseTextFormat.reset();
    _textBrush.reset();
    _cardBorderBrush.reset();
    _cardBrush.reset();
    _bgBrush.reset();
    _renderTarget.reset();
    _writeFactory.reset();
    _d2dFactory.reset();
}

void ViewerPE::EnsureDeviceResources(HWND /*hwnd*/) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const HWND hwnd = _hWnd.get();

    if (! _d2dFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        options.debugLevel = D2D1_DEBUG_LEVEL_NONE;
        const HRESULT hr   = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, _d2dFactory.put());
        if (FAILED(hr))
        {
            return;
        }
    }

    if (! _writeFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_writeFactory.put()));
        if (FAILED(hr))
        {
            return;
        }
    }

    if (! _renderTarget)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const D2D1_SIZE_U size = D2D1::SizeU(static_cast<UINT>(std::max(1l, rc.right - rc.left)), static_cast<UINT>(std::max(1l, rc.bottom - rc.top)));

        D2D1_RENDER_TARGET_PROPERTIES rtProps        = D2D1::RenderTargetProperties(D2D1_RENDER_TARGET_TYPE_DEFAULT,
                                                                                    D2D1::PixelFormat(DXGI_FORMAT_UNKNOWN, D2D1_ALPHA_MODE_IGNORE),
                                                                                    static_cast<float>(_dpi),
                                                                                    static_cast<float>(_dpi));
        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(rtProps, hwndProps, _renderTarget.put());
        if (FAILED(hr))
        {
            return;
        }
    }

    const bool themed            = _hasTheme && ! _theme.highContrast;
    const COLORREF bg            = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg            = themed ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const COLORREF accent        = themed ? ResolveAccentColor(_theme, L"content") : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF cardBg        = themed ? BlendColorRefTruncate(bg, fg, themed && _theme.darkMode ? 24u : 18u) : GetSysColor(COLOR_WINDOW);
    const COLORREF cardBorder    = themed ? BlendColorRefTruncate(cardBg, accent, 92u) : GetSysColor(COLOR_WINDOWFRAME);
    const D2D1_COLOR_F bgColor   = themed ? ColorFromArgb(_theme.backgroundArgb) : ColorFFromColorRef(bg);
    const D2D1_COLOR_F textColor = themed ? ColorFromArgb(_theme.textArgb) : ColorFFromColorRef(fg);

    if (! _bgBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(bgColor, _bgBrush.put()));
    }
    if (! _cardBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardBg), _cardBrush.put()));
    }
    if (! _cardBorderBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardBorder), _cardBorderBrush.put()));
    }
    if (! _textBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(textColor, _textBrush.put()));
    }
}

void ViewerPE::EnsureTextLayout(float viewportWidthDip, float viewportHeightDip) noexcept
{
    if (! _writeFactory)
    {
        return;
    }

    _viewportHeightDip = std::max(1.0f, viewportHeightDip);

    if (_textLayout && std::fabs(_layoutWidthDip - viewportWidthDip) <= 0.5f)
    {
        return;
    }

    if (! _baseTextFormat)
    {
        wil::com_ptr<IDWriteTextFormat> fmt;
        HRESULT hr = Typography::CreateTextFormat(_writeFactory.get(), Typography::MakeUiMonospaceSpec(11.0f), fmt.put(), L"en-us");
        if (FAILED(hr))
        {
            hr = Typography::CreateTextFormat(_writeFactory.get(), Typography::MakeUiTextSpec(11.0f), fmt.put(), L"en-us");
        }
        if (FAILED(hr))
        {
            return;
        }

        fmt->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
        _baseTextFormat = std::move(fmt);
    }

    std::wstring text;
    text = std::format(L"{}\n{}\n\n{}", _titleText, _subtitleText, _bodyText);

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT hr = _writeFactory->CreateTextLayout(
        text.c_str(), static_cast<UINT32>(text.size()), _baseTextFormat.get(), std::max(1.0f, viewportWidthDip), 1000000.0f, layout.put());
    if (FAILED(hr))
    {
        return;
    }

    const UINT32 titleLen      = static_cast<UINT32>(std::min<size_t>(_titleText.size(), std::numeric_limits<UINT32>::max()));
    const UINT32 subtitleStart = titleLen + 1u;
    const UINT32 subtitleLen = static_cast<UINT32>(std::min<size_t>(_subtitleText.size(), (subtitleStart <= text.size()) ? (text.size() - subtitleStart) : 0u));

    if (titleLen > 0)
    {
        const DWRITE_TEXT_RANGE r{0u, titleLen};
        layout->SetFontFamilyName(Typography::kSegoeUiVariableTextFamily, r);
        layout->SetFontSize(20.0f, r);
        layout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, r);
    }
    if (subtitleLen > 0 && subtitleStart < text.size())
    {
        const DWRITE_TEXT_RANGE r{subtitleStart, subtitleLen};
        layout->SetFontFamilyName(Typography::kSegoeUiVariableTextFamily, r);
        layout->SetFontSize(12.0f, r);
        layout->SetFontStyle(DWRITE_FONT_STYLE_ITALIC, r);
    }

    DWRITE_TEXT_METRICS metrics{};
    if (SUCCEEDED(layout->GetMetrics(&metrics)))
    {
        _contentHeightDip = metrics.height;
    }
    else
    {
        _contentHeightDip = 0.0f;
    }

    _layoutWidthDip = viewportWidthDip;
    _textLayout     = std::move(layout);
}

void ViewerPE::UpdateScrollBars(HWND hwnd, float viewportHeightDip) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const float maxScrollDip = std::max(0.0f, _contentHeightDip - viewportHeightDip);
    _scrollDip               = std::clamp(_scrollDip, 0.0f, maxScrollDip);

    const float pxPerDip = static_cast<float>(_dpi) / 96.0f;
    const int viewportPx = std::max(1, static_cast<int>(std::lround(viewportHeightDip * pxPerDip)));
    const int contentPx  = std::max(viewportPx, static_cast<int>(std::lround(_contentHeightDip * pxPerDip)));
    const int scrollPx   = std::clamp(static_cast<int>(std::lround(_scrollDip * pxPerDip)), 0, std::max(0, contentPx - viewportPx));

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentPx - 1);
    si.nPage  = static_cast<UINT>(viewportPx);
    si.nPos   = scrollPx;
    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

void ViewerPE::SetScrollDip(HWND hwnd, float scrollDip) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! std::isfinite(scrollDip))
    {
        scrollDip = std::numeric_limits<float>::max();
    }

    const float clamped = ClampScroll(scrollDip, _contentHeightDip, _viewportHeightDip);
    if (std::fabs(clamped - _scrollDip) <= 0.25f)
    {
        return;
    }

    _scrollDip = clamped;
    UpdateScrollBars(hwnd, _viewportHeightDip);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerPE::ScrollByDip(HWND hwnd, float deltaDip) noexcept
{
    SetScrollDip(hwnd, _scrollDip + deltaDip);
}

bool ViewerPE::QueueAsyncParseWorker(const std::shared_ptr<AsyncParseScheduler>& scheduler) noexcept
{
    if (! scheduler)
    {
        return false;
    }

    auto context             = std::make_unique<AsyncParseScheduler::WorkerContext>();
    context->scheduler       = scheduler;
    context->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerPeModuleAnchor);
    if (! context->moduleKeepAlive)
    {
        return false;
    }

    const BOOL queued = TrySubmitThreadpoolCallback(&ViewerPE::AsyncParseThreadpoolCallback, context.get(), nullptr);
    if (queued == FALSE)
    {
        return false;
    }

    static_cast<void>(context.release());
    return true;
}

void CALLBACK ViewerPE::AsyncParseThreadpoolCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept
{
    std::unique_ptr<AsyncParseScheduler::WorkerContext> worker(static_cast<AsyncParseScheduler::WorkerContext*>(context));
    if (worker && worker->moduleKeepAlive)
    {
        // Releasing the last module reference from inside this callback can unmap the
        // callback's own code before it returns. Let the threadpool release it only
        // after the callback has reached its safe return boundary.
        TransferModulePinToCallbackReturn(instance, worker->moduleKeepAlive);
    }
    if (! worker || ! worker->scheduler)
    {
        return;
    }

    const std::shared_ptr<AsyncParseScheduler> scheduler = worker->scheduler;
    for (;;)
    {
        std::optional<AsyncParseScheduler::WorkItem> work;
        {
            std::scoped_lock lock(scheduler->mutex);
            if (! scheduler->pending.has_value())
            {
                scheduler->workerRunning = false;
                return;
            }

            work = std::move(scheduler->pending);
            scheduler->pending.reset();
        }

        if (work->run)
        {
            work->run();
        }
    }
}

void ViewerPE::CancelAsyncParse() noexcept
{
    const uint64_t requestId = _parseRequestId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kAsyncParseFallbackTimerId));
    }
    if (! _parseScheduler)
    {
        return;
    }

    _parseScheduler->latestWindowIdentity.store(0u, std::memory_order_release);
    _parseScheduler->latestRequestId.store(requestId, std::memory_order_release);
    _parseScheduler->terminalFallbackRequestId.store(0u, std::memory_order_release);
    std::scoped_lock lock(_parseScheduler->mutex);
    _parseScheduler->pending.reset();
}

void ViewerPE::StartAsyncParse(HWND hwnd, wil::com_ptr<IFileSystem> fileSystem, std::wstring path) noexcept
{
    const uint64_t requestId                             = _parseRequestId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    const uint64_t windowIdentity                        = _windowIdentity;
    const std::shared_ptr<AsyncParseScheduler> scheduler = _parseScheduler;
#ifdef ENABLE_TESTS
    const auto injectedFault = static_cast<AsyncParseInjectedFault>(std::exchange(_debugNextAsyncFault, 0u));
#else
    constexpr auto injectedFault = AsyncParseInjectedFault::None;
#endif
    _titleText = std::filesystem::path(path).filename().wstring();

    _subtitleText = LoadStringResource(g_hInstance, IDS_VIEWERPE_STATUS_LOADING);
    _bodyText.clear();
    _markdownText.clear();
    _isLoading   = true;
    _lastParseHr = E_PENDING;
    _textLayout.reset();
    _scrollDip = 0.0f;
    InvalidateRect(hwnd, nullptr, TRUE);

    if (scheduler)
    {
        scheduler->terminalFallbackRequestId.store(0u, std::memory_order_release);
        scheduler->latestWindowIdentity.store(windowIdentity, std::memory_order_release);
        scheduler->latestRequestId.store(requestId, std::memory_order_release);
    }

    static_cast<void>(KillTimer(hwnd, kAsyncParseFallbackTimerId));
    if (SetTimer(hwnd, kAsyncParseFallbackTimerId, kAsyncParseFallbackPollMs, nullptr) == 0u)
    {
        const DWORD error = GetLastError();
        CompleteAsyncParseFailure(error == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(error));
        return;
    }

    if (! fileSystem)
    {
        if (scheduler)
        {
            std::scoped_lock lock(scheduler->mutex);
            scheduler->pending.reset();
        }
        CompleteAsyncParseFailure(E_POINTER);
        return;
    }

    if (! scheduler)
    {
        CompleteAsyncParseFailure(E_UNEXPECTED);
        return;
    }

    std::function<void()> run = [hwnd,
                                 requestId,
                                 windowIdentity,
                                 scheduler,
#ifdef ENABLE_TESTS
                                 injectedFault,
#endif
                                 fileSystem = std::move(fileSystem),
                                 path       = std::move(path)]() noexcept
    {
        Debug::Perf::Scope parseScope(L"viewer.pe.parse");
        parseScope.SetDetail(L"threadpool-latest-wins");
        const auto cancelled = [&]() noexcept
        {
            const bool isCancelled = scheduler->latestRequestId.load(std::memory_order_acquire) != requestId ||
                                     scheduler->latestWindowIdentity.load(std::memory_order_acquire) != windowIdentity;
            if (isCancelled)
            {
                parseScope.SetHr(HRESULT_FROM_WIN32(ERROR_CANCELLED));
            }
            return isCancelled;
        };

        auto postResult = [&](HRESULT hr, std::wstring title, std::wstring subtitle, std::wstring body, std::wstring markdown) noexcept
        {
            parseScope.SetHr(hr);
            if (cancelled())
            {
                return;
            }

            if (! hwnd || IsWindow(hwnd) == 0)
            {
                return;
            }

            auto result = injectedFault == AsyncParseInjectedFault::ResultAllocation ? std::unique_ptr<AsyncParseResult>{}
                                                                                     : std::unique_ptr<AsyncParseResult>(new (std::nothrow) AsyncParseResult());
            if (! result)
            {
                if (! cancelled())
                {
                    scheduler->terminalFallbackHr.store(E_OUTOFMEMORY, std::memory_order_relaxed);
                    scheduler->terminalFallbackWindowIdentity.store(windowIdentity, std::memory_order_relaxed);
                    scheduler->terminalFallbackRequestId.store(requestId, std::memory_order_release);
                }
                return;
            }

            result->requestId      = requestId;
            result->windowIdentity = windowIdentity;
            result->hr             = hr;
            result->title          = std::move(title);
            result->subtitle       = std::move(subtitle);
            result->body           = std::move(body);
            result->markdown       = std::move(markdown);
            const bool posted =
                injectedFault != AsyncParseInjectedFault::PayloadPost && PostMessagePayload(hwnd, kAsyncParseCompleteMessage, 0, std::move(result));
            if (! posted && ! cancelled())
            {
                const DWORD error = injectedFault == AsyncParseInjectedFault::PayloadPost ? ERROR_NOT_ENOUGH_MEMORY : GetLastError();
                scheduler->terminalFallbackHr.store(error == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(error), std::memory_order_relaxed);
                scheduler->terminalFallbackWindowIdentity.store(windowIdentity, std::memory_order_relaxed);
                scheduler->terminalFallbackRequestId.store(requestId, std::memory_order_release);
            }
        };

        if (cancelled())
        {
            return;
        }

        wil::com_ptr<IFileSystemIO> fsio;
        HRESULT hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fsio.put_void());
        if (FAILED(hr) || ! fsio)
        {
            postResult(E_NOINTERFACE, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_NO_FILEIO), {});
            return;
        }

        wil::com_ptr<IFileReader> reader;
        hr = fsio->CreateFileReader(path.c_str(), reader.put());
        if (FAILED(hr) || ! reader)
        {
            postResult(hr, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED), {});
            return;
        }

        uint64_t sizeBytes = 0;
        hr                 = reader->GetSize(&sizeBytes);
        parseScope.SetValue1(sizeBytes);
        if (FAILED(hr) || sizeBytes == 0)
        {
            postResult(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED), {});
            return;
        }

        if (sizeBytes > kMaxPeFileBytes || sizeBytes > static_cast<uint64_t>((std::numeric_limits<std::uint32_t>::max)()) ||
            sizeBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
        {
            postResult(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_TOO_LARGE), {});
            return;
        }

        uint64_t position = 0u;
        hr                = reader->Seek(0, FILE_BEGIN, &position);
        if (FAILED(hr) || position != 0u)
        {
            postResult(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
            return;
        }

        std::vector<std::uint8_t> bytes;
        bytes.resize(static_cast<size_t>(sizeBytes));

        size_t offset = 0;
        while (offset < bytes.size())
        {
            if (cancelled())
            {
                return;
            }

            const size_t remaining   = bytes.size() - offset;
            const unsigned long want = static_cast<unsigned long>((std::min)(remaining, static_cast<size_t>(1u * 1024u * 1024u)));
            unsigned long read       = 0;
            hr                       = reader->Read(bytes.data() + offset, want, &read);
            if (FAILED(hr))
            {
                postResult(hr, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
                return;
            }
            if (read > want)
            {
                postResult(HRESULT_FROM_WIN32(ERROR_INVALID_DATA), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
                return;
            }
            if (read == 0)
            {
                postResult(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
                return;
            }
            offset += static_cast<size_t>(read);
            parseScope.SetValue0(offset);
        }

        if (offset != bytes.size())
        {
            postResult(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
            return;
        }

        if (cancelled())
        {
            return;
        }

        std::uint8_t trailingByte  = 0u;
        unsigned long trailingRead = 0u;
        hr                         = reader->Read(&trailingByte, 1u, &trailingRead);
        if (FAILED(hr) || trailingRead > 1u || trailingRead != 0u)
        {
            postResult(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
            return;
        }

        std::unique_ptr<peparse::parsed_pe, ParsedPeDeleter> pe(peparse::ParsePEFromPointer(bytes.data(), static_cast<std::uint32_t>(bytes.size())));
        if (! pe)
        {
            std::wstring err          = LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_PARSE_FAILED);
            const std::wstring detail = Utf16FromUtf8(peparse::GetPEErrString());
            const std::wstring loc    = Utf16FromUtf8(peparse::GetPEErrLoc());
            if (! detail.empty())
            {
                err += L"\n\n";
                err += detail;
                if (! loc.empty())
                {
                    err += L"\n";
                    err += loc;
                }
            }

            postResult(E_FAIL, {}, {}, std::move(err), {});
            return;
        }

        const auto& nt               = pe->peHeader.nt;
        const std::uint16_t optMagic = nt.OptionalMagic;
        const bool is64              = optMagic == 0x20Bu;

        peparse::VA entryPoint   = 0;
        const bool hasEntryPoint = peparse::GetEntryPoint(pe.get(), entryPoint);

        const std::wstring machine       = MachineText(pe.get());
        const std::wstring subsystem     = SubsystemText(pe.get());
        const std::wstring_view kindName = PeKindName(optMagic);

        std::wstring subtitle;
        subtitle = FormatStringResource(g_hInstance,
                                        IDS_VIEWERPE_REPORT_SUBTITLE_FORMAT,
                                        kindName,
                                        machine.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_UNKNOWN_MACHINE) : machine,
                                        nt.FileHeader.NumberOfSections);

        const auto& dos        = pe->peHeader.dos;
        const auto& fileHeader = nt.FileHeader;
        const auto& opt32      = nt.OptionalHeader;
        const auto& opt64      = nt.OptionalHeader64;

        const uint64_t imageBase = is64 ? opt64.ImageBase : opt32.ImageBase;

        std::wstring entryPointText;
        if (hasEntryPoint)
        {
            entryPointText = std::format(L"0x{:016X}", static_cast<uint64_t>(entryPoint));
        }
        else
        {
            entryPointText = LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_VALUE_NONE);
        }

        std::wstring body;
        const std::wstring unknownValue = LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_VALUE_UNKNOWN);
        body                            = FormatStringResource(g_hInstance,
                                                               IDS_VIEWERPE_REPORT_SUMMARY_FORMAT,
                                                               path,
                                                               sizeBytes,
                                                               FormatBytesCompact(sizeBytes),
                                                               kindName,
                                                               machine.empty() ? unknownValue : machine,
                                                               subsystem.empty() ? unknownValue : subsystem,
                                                               static_cast<std::uint32_t>(fileHeader.TimeDateStamp),
                                                               fileHeader.Characteristics,
                                                               imageBase,
                                                               entryPointText,
                                                               dos.e_magic,
                                                               dos.e_lfanew);

        body += LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_SECTION_COLUMNS);
        body += L'\n';

        struct SectionRow
        {
            std::wstring name;
            std::uint32_t rva     = 0;
            std::uint32_t vsize   = 0;
            std::uint32_t rawPtr  = 0;
            std::uint32_t rawSize = 0;
            std::uint32_t chars   = 0;
        };

        std::vector<SectionRow> sections;
        peparse::IterSec(
            pe.get(),
            [](void* ctx, const peparse::VA& /*va*/, const std::string& name, const peparse::image_section_header& hdr, const peparse::bounded_buffer* /*buf*/)
                -> int
        {
            auto* out = static_cast<std::vector<SectionRow>*>(ctx);
            if (! out)
            {
                return 0;
            }

            SectionRow row{};
            row.name.assign(name.begin(), name.end());
            row.rva     = hdr.VirtualAddress;
            row.vsize   = hdr.Misc.VirtualSize;
            row.rawPtr  = hdr.PointerToRawData;
            row.rawSize = hdr.SizeOfRawData;
            row.chars   = hdr.Characteristics;

            out->push_back(std::move(row));
            return 0;
        },
            &sections);

        for (const auto& sec : sections)
        {
            if (cancelled())
            {
                return;
            }

            std::wstring line;
            line = std::format(L"{:<10} 0x{:08X} 0x{:08X} 0x{:08X} 0x{:08X} 0x{:08X}\n", sec.name, sec.rva, sec.vsize, sec.rawPtr, sec.rawSize, sec.chars);
            body += line;
        }

        const auto safeAppend = [&](std::wstring_view text) noexcept { body.append(text); };

        const auto safeAppendLine = [&](std::wstring_view text) noexcept
        {
            safeAppend(text);
            safeAppend(L"\n");
        };

        const auto safeAppendFormat = [&]<typename... Args>(std::wformat_string<Args...> fmt, Args&&... args) noexcept
        { safeAppend(std::format(fmt, std::forward<Args>(args)...)); };

        const auto safeAppendResourceLine = [&](UINT resourceId) noexcept { safeAppendLine(LoadStringResource(g_hInstance, resourceId)); };

        const auto safeAppendResourceFormat = [&]<typename... Args>(UINT resourceId, Args&&... args) noexcept
        { safeAppend(FormatStringResource(g_hInstance, resourceId, std::forward<Args>(args)...)); };

        const auto safeAppendField = [&]<typename... Args>(UINT formatResourceId, UINT labelResourceId, Args&&... args) noexcept
        { safeAppendResourceFormat(formatResourceId, LoadStringResource(g_hInstance, labelResourceId), std::forward<Args>(args)...); };

        const auto safeAppendBlankLine = [&]() noexcept { safeAppend(L"\n"); };

        safeAppendBlankLine();
        safeAppendResourceLine(IDS_VIEWERPE_REPORT_RICH_HEADER);
        const std::wstring yesText = LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_VALUE_YES);
        const std::wstring noText  = LoadStringResource(g_hInstance, IDS_VIEWERPE_REPORT_VALUE_NO);
        safeAppendResourceFormat(IDS_VIEWERPE_REPORT_RICH_PRESENT_FORMAT, pe->peHeader.rich.isPresent ? yesText : noText);
        safeAppendResourceFormat(IDS_VIEWERPE_REPORT_RICH_VALID_FORMAT, pe->peHeader.rich.isValid ? yesText : noText);
        if (pe->peHeader.rich.isPresent)
        {
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_DECRYPTION_KEY, pe->peHeader.rich.DecryptionKey);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_CHECKSUM, pe->peHeader.rich.Checksum);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_ENTRIES, pe->peHeader.rich.Entries.size());

            static constexpr size_t kMaxRichEntries = 256;
            if (! pe->peHeader.rich.Entries.empty())
            {
                safeAppendBlankLine();
                safeAppendResourceLine(IDS_VIEWERPE_REPORT_RICH_COLUMNS);
                safeAppendLine(L"-------- -----     ---------  ------------------------------");

                size_t shown = 0;
                for (const auto& entry : pe->peHeader.rich.Entries)
                {
                    if (cancelled())
                    {
                        return;
                    }

                    if (shown >= kMaxRichEntries)
                    {
                        safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, kMaxRichEntries);
                        break;
                    }

                    const std::string& prod = peparse::GetRichProductName(entry.BuildNumber);
                    const std::string& obj  = peparse::GetRichObjectType(entry.ProductId);
                    std::wstring label      = Utf16FromUtf8(prod);
                    if (! obj.empty())
                    {
                        label += L" ";
                        label += Utf16FromUtf8(obj);
                    }

                    safeAppendFormat(L"{:>8} {:>5}     {:>9}  {}\n", entry.ProductId, entry.BuildNumber, entry.Count, label);
                    shown += 1;
                }
            }
        }

        safeAppendBlankLine();
        safeAppendResourceLine(IDS_VIEWERPE_REPORT_FILE_HEADER);
        safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_NUMBER_OF_SECTIONS, fileHeader.NumberOfSections);
        safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_OPTIONAL_HEADER, fileHeader.SizeOfOptionalHeader);
        safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_POINTER_TO_SYMBOL_TABLE, fileHeader.PointerToSymbolTable);
        safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_NUMBER_OF_SYMBOLS, fileHeader.NumberOfSymbols);

        safeAppendBlankLine();
        safeAppendResourceLine(IDS_VIEWERPE_REPORT_OPTIONAL_HEADER);
        if (is64)
        {
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX4_FORMAT, IDS_VIEWERPE_REPORT_FIELD_MAGIC, opt64.Magic);
            safeAppendField(
                IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT, IDS_VIEWERPE_REPORT_FIELD_LINKER_VERSION, opt64.MajorLinkerVersion, opt64.MinorLinkerVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_IMAGE, opt64.SizeOfImage);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEADERS, opt64.SizeOfHeaders);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_CHECKSUM, opt64.CheckSum);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX4_FORMAT, IDS_VIEWERPE_REPORT_FIELD_DLL_CHARACTERISTICS, opt64.DllCharacteristics);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SECTION_ALIGNMENT, opt64.SectionAlignment);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_FILE_ALIGNMENT, opt64.FileAlignment);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT,
                            IDS_VIEWERPE_REPORT_FIELD_OS_VERSION,
                            opt64.MajorOperatingSystemVersion,
                            opt64.MinorOperatingSystemVersion);
            safeAppendField(
                IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT, IDS_VIEWERPE_REPORT_FIELD_IMAGE_VERSION, opt64.MajorImageVersion, opt64.MinorImageVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT,
                            IDS_VIEWERPE_REPORT_FIELD_SUBSYSTEM_VERSION,
                            opt64.MajorSubsystemVersion,
                            opt64.MinorSubsystemVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX16_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_STACK_RESERVE, opt64.SizeOfStackReserve);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX16_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_STACK_COMMIT, opt64.SizeOfStackCommit);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX16_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEAP_RESERVE, opt64.SizeOfHeapReserve);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX16_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEAP_COMMIT, opt64.SizeOfHeapCommit);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_NUMBER_OF_RVA_AND_SIZES, opt64.NumberOfRvaAndSizes);
        }
        else
        {
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX4_FORMAT, IDS_VIEWERPE_REPORT_FIELD_MAGIC, opt32.Magic);
            safeAppendField(
                IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT, IDS_VIEWERPE_REPORT_FIELD_LINKER_VERSION, opt32.MajorLinkerVersion, opt32.MinorLinkerVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_IMAGE, opt32.SizeOfImage);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEADERS, opt32.SizeOfHeaders);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_CHECKSUM, opt32.CheckSum);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX4_FORMAT, IDS_VIEWERPE_REPORT_FIELD_DLL_CHARACTERISTICS, opt32.DllCharacteristics);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SECTION_ALIGNMENT, opt32.SectionAlignment);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_FILE_ALIGNMENT, opt32.FileAlignment);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT,
                            IDS_VIEWERPE_REPORT_FIELD_OS_VERSION,
                            opt32.MajorOperatingSystemVersion,
                            opt32.MinorOperatingSystemVersion);
            safeAppendField(
                IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT, IDS_VIEWERPE_REPORT_FIELD_IMAGE_VERSION, opt32.MajorImageVersion, opt32.MinorImageVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_VERSION_FORMAT,
                            IDS_VIEWERPE_REPORT_FIELD_SUBSYSTEM_VERSION,
                            opt32.MajorSubsystemVersion,
                            opt32.MinorSubsystemVersion);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_STACK_RESERVE, opt32.SizeOfStackReserve);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_STACK_COMMIT, opt32.SizeOfStackCommit);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEAP_RESERVE, opt32.SizeOfHeapReserve);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_HEX8_FORMAT, IDS_VIEWERPE_REPORT_FIELD_SIZE_OF_HEAP_COMMIT, opt32.SizeOfHeapCommit);
            safeAppendField(IDS_VIEWERPE_REPORT_FIELD_DEC_FORMAT, IDS_VIEWERPE_REPORT_FIELD_NUMBER_OF_RVA_AND_SIZES, opt32.NumberOfRvaAndSizes);
        }

        safeAppendBlankLine();
        safeAppendResourceLine(IDS_VIEWERPE_REPORT_DATA_DIRECTORIES);
        safeAppendResourceLine(IDS_VIEWERPE_REPORT_DATA_DIRECTORY_COLUMNS);
        safeAppendLine(L"-------------------  ---------  ---------");
        for (size_t i = 0; i < kDataDirectoryNames.size(); ++i)
        {
            const peparse::data_directory dir = is64 ? opt64.DataDirectory[i] : opt32.DataDirectory[i];
            safeAppendFormat(L"{:<19}  0x{:08X}  0x{:08X}\n", kDataDirectoryNames[i], dir.VirtualAddress, dir.Size);
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxImportRows  = 600;
        static constexpr size_t kHardMaxImports = 6000;

        struct ImportRow
        {
            uint64_t va = 0;
            std::wstring module;
            std::wstring name;
        };

        std::vector<ImportRow> imports;
        imports.reserve(128);

        struct ImportCollect
        {
            std::vector<ImportRow>* out = nullptr;
            size_t maxRows              = 0;
            size_t hardMax              = 0;
            size_t seen                 = 0;
            bool truncated              = false;
        };

        ImportCollect importCollect{&imports, kMaxImportRows, kHardMaxImports, 0, false};
        peparse::IterImpVAString(pe.get(),
                                 [](void* ctx, const peparse::VA& va, const std::string& module, const std::string& name) -> int
        {
            auto* c = static_cast<ImportCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                ImportRow row{};
                row.va     = static_cast<uint64_t>(va);
                row.module = Utf16FromUtf8(module);
                row.name   = Utf16FromUtf8(name);

                c->out->push_back(std::move(row));
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
                                 &importCollect);

        safeAppendBlankLine();
        if (importCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_IMPORTS_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_IMPORTS_COUNT_FORMAT, importCollect.seen, importCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_IMPORTS_COLUMNS);
            safeAppendLine(L"-----------------  -----------------------  ------------------------------");
            for (const auto& imp : imports)
            {
                safeAppendFormat(L"0x{:016X} {:<23}  {}\n", imp.va, imp.module, imp.name);
            }
            if (importCollect.seen > imports.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, imports.size());
            }
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxExportRows  = 600;
        static constexpr size_t kHardMaxExports = 6000;

        struct ExportRow
        {
            uint64_t va       = 0;
            std::uint16_t ord = 0;
            std::wstring name;
            std::wstring forward;
        };

        std::vector<ExportRow> exports;
        exports.reserve(128);

        struct ExportCollect
        {
            std::vector<ExportRow>* out = nullptr;
            size_t maxRows              = 0;
            size_t hardMax              = 0;
            size_t seen                 = 0;
            bool truncated              = false;
        };

        ExportCollect exportCollect{&exports, kMaxExportRows, kHardMaxExports, 0, false};
        peparse::IterExpFull(
            pe.get(),
            [](void* ctx, const peparse::VA& va, std::uint16_t ord, const std::string& name, const std::string& /*module*/, const std::string& forwardStr)
                -> int
        {
            auto* c = static_cast<ExportCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                ExportRow row{};
                row.va      = static_cast<uint64_t>(va);
                row.ord     = ord;
                row.name    = Utf16FromUtf8(name);
                row.forward = Utf16FromUtf8(forwardStr);

                c->out->push_back(std::move(row));
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
            &exportCollect);

        safeAppendBlankLine();
        if (exportCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_EXPORTS_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_EXPORTS_COUNT_FORMAT, exportCollect.seen, exportCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_EXPORTS_COLUMNS);
            safeAppendLine(L"----  -----------------  --------------------------------------------");
            for (const auto& exp : exports)
            {
                if (exp.va != 0)
                {
                    safeAppendFormat(L"{:>4} 0x{:016X}  {}\n", exp.ord, exp.va, exp.name);
                }
                else if (! exp.forward.empty())
                {
                    safeAppendResourceFormat(IDS_VIEWERPE_REPORT_EXPORT_FORWARDED_FORMAT, exp.ord, exp.name, exp.forward);
                }
                else
                {
                    safeAppendResourceFormat(IDS_VIEWERPE_REPORT_EXPORT_UNAVAILABLE_FORMAT, exp.ord, exp.name);
                }
            }
            if (exportCollect.seen > exports.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, exports.size());
            }
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxResourceRows  = 600;
        static constexpr size_t kHardMaxResources = 6000;

        struct ResourceRow
        {
            std::wstring type;
            std::wstring name;
            std::wstring lang;
            std::uint32_t codepage = 0;
            std::uint32_t rva      = 0;
            std::uint32_t size     = 0;
        };

        std::vector<ResourceRow> resources;
        resources.reserve(128);

        struct ResourceCollect
        {
            std::vector<ResourceRow>* out = nullptr;
            size_t maxRows                = 0;
            size_t hardMax                = 0;
            size_t seen                   = 0;
            bool truncated                = false;
        };

        ResourceCollect rsrcCollect{&resources, kMaxResourceRows, kHardMaxResources, 0, false};
        peparse::IterRsrc(pe.get(),
                          [](void* ctx, const peparse::resource& res) -> int
        {
            auto* c = static_cast<ResourceCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                ResourceRow row{};
                row.type     = Utf16FromUtf8(res.type_str);
                row.name     = Utf16FromUtf8(res.name_str);
                row.lang     = Utf16FromUtf8(res.lang_str);
                row.codepage = res.codepage;
                row.rva      = res.RVA;
                row.size     = res.size;

                if (row.type.empty())
                {
                    row.type = std::to_wstring(res.type);
                }
                if (row.name.empty())
                {
                    row.name = std::to_wstring(res.name);
                }
                if (row.lang.empty())
                {
                    row.lang = std::to_wstring(res.lang);
                }

                c->out->push_back(std::move(row));
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
                          &rsrcCollect);

        safeAppendBlankLine();
        if (rsrcCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_RESOURCES_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_RESOURCES_COUNT_FORMAT, rsrcCollect.seen, rsrcCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_RESOURCES_COLUMNS);
            safeAppendLine(L"---------  ---------  --------  ---------------------------------------");
            for (const auto& res : resources)
            {
                safeAppendFormat(L"0x{:08X} 0x{:08X} {:>8}  {}/{}/{}\n", res.rva, res.size, res.codepage, res.type, res.name, res.lang);
            }
            if (rsrcCollect.seen > resources.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, resources.size());
            }
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxRelocRows  = 600;
        static constexpr size_t kHardMaxRelocs = 20000;

        struct RelocRow
        {
            uint64_t va              = 0;
            peparse::reloc_type type = peparse::RELOC_ABSOLUTE;
        };

        std::vector<RelocRow> relocs;
        relocs.reserve(256);

        struct RelocCollect
        {
            std::vector<RelocRow>* out = nullptr;
            size_t maxRows             = 0;
            size_t hardMax             = 0;
            size_t seen                = 0;
            bool truncated             = false;
        };

        RelocCollect relocCollect{&relocs, kMaxRelocRows, kHardMaxRelocs, 0, false};
        peparse::IterRelocs(pe.get(),
                            [](void* ctx, const peparse::VA& va, const peparse::reloc_type& type) -> int
        {
            auto* c = static_cast<RelocCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                RelocRow row{};
                row.va   = static_cast<uint64_t>(va);
                row.type = type;

                c->out->push_back(row);
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
                            &relocCollect);

        safeAppendBlankLine();
        if (relocCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_RELOCATIONS_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_RELOCATIONS_COUNT_FORMAT, relocCollect.seen, relocCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_RELOCATIONS_COLUMNS);
            safeAppendLine(L"-----------------  ----");
            for (const auto& rel : relocs)
            {
                safeAppendFormat(L"0x{:016X}  {:L}\n", rel.va, static_cast<unsigned long>(rel.type));
            }
            if (relocCollect.seen > relocs.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, relocs.size());
            }
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxDebugRows  = 256;
        static constexpr size_t kHardMaxDebugs = 2000;

        struct DebugRow
        {
            std::uint32_t type = 0;
            std::uint32_t size = 0;
        };

        std::vector<DebugRow> debugs;
        debugs.reserve(32);

        struct DebugCollect
        {
            std::vector<DebugRow>* out = nullptr;
            size_t maxRows             = 0;
            size_t hardMax             = 0;
            size_t seen                = 0;
            bool truncated             = false;
        };

        DebugCollect debugCollect{&debugs, kMaxDebugRows, kHardMaxDebugs, 0, false};
        peparse::IterDebugs(pe.get(),
                            [](void* ctx, const std::uint32_t& type, const peparse::bounded_buffer* buf) -> int
        {
            auto* c = static_cast<DebugCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                DebugRow row{};
                row.type = type;
                row.size = buf ? buf->bufLen : 0;
                c->out->push_back(row);
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
                            &debugCollect);

        safeAppendBlankLine();
        if (debugCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_DEBUG_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_DEBUG_COUNT_FORMAT, debugCollect.seen, debugCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_DEBUG_COLUMNS);
            safeAppendLine(L"---------  ---------");
            for (const auto& dbg : debugs)
            {
                safeAppendFormat(L"{:>9} 0x{:08X}\n", dbg.type, dbg.size);
            }
            if (debugCollect.seen > debugs.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, debugs.size());
            }
        }

        if (cancelled())
        {
            return;
        }

        static constexpr size_t kMaxSymbolRows  = 600;
        static constexpr size_t kHardMaxSymbols = 6000;

        struct SymbolRow
        {
            std::wstring name;
            std::uint32_t value  = 0;
            std::int16_t section = 0;
            std::uint16_t type   = 0;
            std::uint8_t storage = 0;
            std::uint8_t aux     = 0;
        };

        std::vector<SymbolRow> symbols;
        symbols.reserve(128);

        struct SymbolCollect
        {
            std::vector<SymbolRow>* out = nullptr;
            size_t maxRows              = 0;
            size_t hardMax              = 0;
            size_t seen                 = 0;
            bool truncated              = false;
        };

        SymbolCollect symbolCollect{&symbols, kMaxSymbolRows, kHardMaxSymbols, 0, false};
        peparse::IterSymbols(pe.get(),
                             [](void* ctx,
                                const std::string& name,
                                const std::uint32_t& value,
                                const std::int16_t& section,
                                const std::uint16_t& type,
                                const std::uint8_t& storage,
                                const std::uint8_t& aux) -> int
        {
            auto* c = static_cast<SymbolCollect*>(ctx);
            if (! c)
            {
                return 0;
            }

            c->seen += 1;
            if (c->out && c->out->size() < c->maxRows)
            {
                SymbolRow row{};
                row.name    = Utf16FromUtf8(name);
                row.value   = value;
                row.section = section;
                row.type    = type;
                row.storage = storage;
                row.aux     = aux;

                c->out->push_back(std::move(row));
            }

            if (c->seen >= c->hardMax)
            {
                c->truncated = true;
                return 1;
            }

            return 0;
        },
                             &symbolCollect);

        safeAppendBlankLine();
        if (symbolCollect.seen == 0)
        {
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_SYMBOLS_NONE);
        }
        else
        {
            safeAppendResourceFormat(IDS_VIEWERPE_REPORT_SYMBOLS_COUNT_FORMAT, symbolCollect.seen, symbolCollect.truncated ? L"+" : L"");
            safeAppendResourceLine(IDS_VIEWERPE_REPORT_SYMBOLS_COLUMNS);
            safeAppendLine(L"---------  ---- ----  ---- --- --------------------------------");
            for (const auto& sym : symbols)
            {
                safeAppendFormat(L"0x{:08X} {:>4} 0x{:04X} {:>4} {:>3} {}\n", sym.value, sym.section, sym.type, sym.storage, sym.aux, sym.name);
            }
            if (symbolCollect.seen > symbols.size())
            {
                safeAppendResourceFormat(IDS_VIEWERPE_REPORT_TRUNCATED_FORMAT, symbols.size());
            }
        }

        std::wstring title;
        title = std::filesystem::path(path).filename().wstring();

        std::wstring markdown;
        const std::wstring mdTitle = title.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME) : title;
        markdown                   = std::format(L"# {}\n\n{}\n\n```text\n{}\n```\n", mdTitle, subtitle, body);

        postResult(S_OK, std::move(title), std::move(subtitle), std::move(body), std::move(markdown));
    };

    bool queueWorker     = false;
    bool replacedPending = false;
    {
        std::scoped_lock lock(scheduler->mutex);
        replacedPending    = scheduler->pending.has_value();
        scheduler->pending = AsyncParseScheduler::WorkItem{.requestId = requestId, .run = std::move(run)};
        if (! scheduler->workerRunning)
        {
            scheduler->workerRunning = true;
            queueWorker              = true;
        }
    }

    Debug::Perf::Emit(L"viewer.pe.queue",
                      queueWorker       ? L"worker-start"
                      : replacedPending ? L"pending-replaced"
                                        : L"pending",
                      0u,
                      1u,
                      replacedPending ? 1u : 0u,
                      S_OK);

    if (! queueWorker || QueueAsyncParseWorker(scheduler))
    {
        return;
    }

    Debug::Perf::Emit(L"viewer.pe.queue", L"submit-failed", 0u, 0u, 0u, HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY));

    {
        std::scoped_lock lock(scheduler->mutex);
        if (scheduler->pending.has_value() && scheduler->pending->requestId == requestId)
        {
            scheduler->pending.reset();
        }
        scheduler->workerRunning = false;
    }

    if (scheduler->latestRequestId.load(std::memory_order_acquire) == requestId &&
        scheduler->latestWindowIdentity.load(std::memory_order_acquire) == windowIdentity)
    {
        CompleteAsyncParseFailure(HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY));
    }
}

void ViewerPE::OnAsyncParseComplete(std::unique_ptr<AsyncParseResult> result) noexcept
{
    if (! result)
    {
        return;
    }

    if (result->requestId != _parseRequestId.load(std::memory_order_acquire) || result->windowIdentity != _windowIdentity)
    {
        return;
    }

    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kAsyncParseFallbackTimerId));
    }

    _titleText    = std::move(result->title);
    _subtitleText = std::move(result->subtitle);
    _bodyText     = std::move(result->body);
    _markdownText = std::move(result->markdown);
    _isLoading    = false;
    _lastParseHr  = result->hr;

    if (_hWnd)
    {
        std::wstring title;
        const std::wstring viewerName = LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME);
        title                         = _titleText.empty() ? viewerName : std::format(L"{} — {}", _titleText, viewerName);
        SetWindowTextW(_hWnd.get(), title.c_str());
    }

    _textLayout.reset();
    _scrollDip = 0.0f;
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::CompleteAsyncParseFailure(HRESULT hr) noexcept
{
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kAsyncParseFallbackTimerId));
    }

    _subtitleText.clear();
    _bodyText = LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED);
    _markdownText.clear();
    _isLoading   = false;
    _lastParseHr = hr;
    _textLayout.reset();
    _scrollDip = 0.0f;

    Debug::Error(std::format(L"ViewerPE: asynchronous parse failed before a result could be delivered (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::PollAsyncParseTerminalFallback() noexcept
{
    const std::shared_ptr<AsyncParseScheduler> scheduler = _parseScheduler;
    if (! _isLoading || ! scheduler)
    {
        return;
    }

    const uint64_t requestId = scheduler->terminalFallbackRequestId.load(std::memory_order_acquire);
    if (requestId == 0u || requestId != _parseRequestId.load(std::memory_order_acquire) ||
        scheduler->terminalFallbackWindowIdentity.load(std::memory_order_relaxed) != _windowIdentity)
    {
        return;
    }

    uint64_t expected = requestId;
    if (! scheduler->terminalFallbackRequestId.compare_exchange_strong(expected, 0u, std::memory_order_acq_rel))
    {
        return;
    }

    CompleteAsyncParseFailure(static_cast<HRESULT>(scheduler->terminalFallbackHr.load(std::memory_order_relaxed)));
}

HRESULT STDMETHODCALLTYPE ViewerPE::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || ! context->fileSystem)
    {
        return E_INVALIDARG;
    }

    if (! RegisterWndClass(g_hInstance))
    {
        return E_FAIL;
    }

    const std::filesystem::path focused(context->focusedPath);
    const std::wstring fileName = focused.filename().wstring();

    const bool embeddedMode   = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (embeddedParent == nullptr || IsWindow(embeddedParent) == FALSE))
    {
        Debug::Error(L"ViewerPE: embedded Open requires a valid ownerWindow parent.");
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    if (! _hWnd)
    {
        RECT ownerRect{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 900;
        int h = 700;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) == 0)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerPE: GetClientRect failed for embedded preview parent.");
                return HRESULT_FROM_WIN32(lastError);
            }
            x = 0;
            y = 0;
            w = std::max(1L, client.right - client.left);
            h = std::max(1L, client.bottom - client.top);
        }
        else if (context->ownerWindow && GetWindowRect(context->ownerWindow, &ownerRect) != 0)
        {
            x = static_cast<int>(ownerRect.left);
            y = static_cast<int>(ownerRect.top);
            w = std::max(1, static_cast<int>(ownerRect.right - ownerRect.left));
            h = std::max(1, static_cast<int>(ownerRect.bottom - ownerRect.top));
        }

        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(embeddedMode ? nullptr
                                                                                          : Localization::LoadMenuResource(g_hInstance, IDR_VIEWERPE_MENU));
        const DWORD style =
            embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VSCROLL) : (WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VSCROLL);
        const std::wstring initialTitle = embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME) : _metaName);
        HWND window =
            CreateWindowExW(0, kClassName, initialTitle.c_str(), style, x, y, w, h, embeddedMode ? embeddedParent : nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ViewerPE: CreateWindowExW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }

        if (! embeddedMode)
        {
            menu.release();
        }
        _hWnd.reset(window);
        _embeddedMode = embeddedMode;

        if (! embeddedMode)
        {
            ApplyTitleBarTheme(true);
        }

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        if (_embeddedMode)
        {
            RECT client{};
            if (embeddedParent != nullptr && GetClientRect(embeddedParent, &client) != 0)
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             0,
                             0,
                             std::max(1L, client.right - client.left),
                             std::max(1L, client.bottom - client.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }

    _fileSystem = context->fileSystem;

    _currentPath = context->focusedPath;
    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        _otherFiles.reserve(static_cast<size_t>(context->otherFileCount));

        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* s = context->otherFiles[i];
            if (! s)
            {
                continue;
            }

            _otherFiles.emplace_back(s);
        }
    }

    if (_otherFiles.size() > 1 && context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }
    else
    {
        _otherIndex = 0;
    }

    if (_otherFiles.empty())
    {
        _otherFiles.push_back(_currentPath);
        _otherIndex = 0;
    }

    if (_hWnd)
    {
        RefreshFileCombo(_hWnd.get());
        UpdateMenuState(_hWnd.get());
    }

    _titleText    = fileName;
    _subtitleText = LoadStringResource(g_hInstance, IDS_VIEWERPE_STATUS_LOADING);
    _bodyText.clear();
    _markdownText.clear();
    _textLayout.reset();
    _scrollDip = 0.0f;

    StartAsyncParse(_hWnd.get(), _fileSystem, _currentPath);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::Close() noexcept
{
    Debug::Perf::Scope closeScope(L"viewer.pe.close");
    closeScope.SetDetail(L"cancel-without-wait");
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    CancelAsyncParse();
    _hWnd.reset();
    _embeddedMode = false;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version < 2u || theme->version > 4u)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    if (_hWnd)
    {
        _fileComboHost.SetTheme(MakeThemePaletteFromViewerTheme(_theme));
        const bool active = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(active);
        _textBrush.reset();
        _cardBorderBrush.reset();
        _cardBrush.reset();
        _bgBrush.reset();
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }

    return S_OK;
}

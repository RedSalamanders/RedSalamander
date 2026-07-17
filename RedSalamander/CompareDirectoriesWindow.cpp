#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "DxUi/DxUi.Typography.h"
#include "FileMetadataFormatting.h"
#include "LocalizationManager.h"

#include <windowsx.h>

namespace CompareDirectoriesWindowInternal
{
// UI-thread-only registry for theme refresh.
std::vector<HWND> g_compareDirectoriesWindows;

#ifdef ENABLE_TESTS
std::atomic<bool> g_failNextCompareDirectoriesWindowCreate{false};
#endif

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

constexpr wchar_t kCompareProgressSpinnerStateProp[] = L"RedSalamander.CompareDirectories.ProgressSpinnerState";
constexpr wchar_t kCompareProgressSpinnerClassName[] = L"RedSalamander.CompareDirectories.ProgressSpinner";

[[nodiscard]] bool EnsureCompareProgressSpinnerWindowClass(HINSTANCE instance) noexcept
{
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kCompareProgressSpinnerClassName, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc   = &CompareProgressSpinnerWndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kCompareProgressSpinnerClassName;
    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

void LogComparePerfStats(std::wstring_view reason, const std::shared_ptr<CompareDirectoriesSession>& session, HRESULT resultHr) noexcept
{
    if (! session)
    {
        return;
    }

    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
    const auto toMiB                        = [](uint64_t bytes) noexcept -> uint64_t { return bytes / (1024u * 1024u); };

    Debug::Info(L"ComparePerf({}): hr=0x{:08X} v={} ui={} scanActive={} scanQ(H={},L={},sched={},inflight={},pendingSubdir={}) "
                L"content(pending={},done={}/{},QH={},QL={},inflight={},updates={}) decisionCache(entries={},~{}MiB/{}MiB) "
                L"dirCache(cur={}MiB,max={}MiB,hits={},misses={},enum={},evict={})",
                reason,
                static_cast<unsigned>(resultHr),
                stats.version,
                stats.uiVersion,
                stats.scanActiveScans,
                stats.scanQueueHighSize,
                stats.scanQueueLowSize,
                stats.scanScheduledKeys,
                stats.scanInFlightKeys,
                stats.pendingSubdirUpdates,
                stats.contentPendingCompares,
                stats.contentCompletedCompares,
                stats.contentTotalCompares,
                stats.contentQueueHighSize,
                stats.contentQueueLowSize,
                stats.contentInFlightSize,
                stats.pendingContentUpdates,
                stats.decisionCacheEntries,
                toMiB(stats.decisionCacheEstimatedBytes),
                toMiB(stats.decisionCacheBudgetBytes),
                toMiB(stats.directoryInfoCache.currentBytes),
                toMiB(stats.directoryInfoCache.maxBytes),
                stats.directoryInfoCache.cacheHits,
                stats.directoryInfoCache.cacheMisses,
                stats.directoryInfoCache.enumerations,
                stats.directoryInfoCache.evictions);
}

template <typename T> void QueueCompareCleanup(std::unique_ptr<T> cleanup, std::wstring_view label) noexcept
{
    if (! cleanup)
    {
        return;
    }

    if (TrySubmitUniqueToThreadpool(cleanup))
    {
        return;
    }

    const DWORD lastError = GetLastError();
    static_cast<void>(cleanup.release());
    Debug::Warning(L"CompareDirectories: {} scheduling failed (gle=0x{:08X}); abandoning deferred cleanup to keep UI teardown non-blocking.",
                   label,
                   static_cast<unsigned long>(lastError));
}

struct CreatedFileSystemInstance
{
    wil::unique_hmodule module;
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginShortId;

    CreatedFileSystemInstance()                                            = default;
    CreatedFileSystemInstance(const CreatedFileSystemInstance&)            = delete;
    CreatedFileSystemInstance& operator=(const CreatedFileSystemInstance&) = delete;
    CreatedFileSystemInstance(CreatedFileSystemInstance&&)                 = default;
    CreatedFileSystemInstance& operator=(CreatedFileSystemInstance&&)      = default;
};

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindFileSystemPluginById(std::wstring_view pluginId) noexcept
{
    return FileSystemPluginManager::GetInstance().FindPluginById(pluginId);
}

[[nodiscard]] FolderWindow::Pane ToFolderWindowPane(ComparePane pane) noexcept
{
    return pane == ComparePane::Left ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
}

[[nodiscard]] std::optional<CreatedFileSystemInstance> TryCreateFileSystemInstance(std::wstring_view pluginId, const std::wstring& instanceContext) noexcept
{
    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || ! entry->fileSystem || entry->path.empty())
    {
        return std::nullopt;
    }

    wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        const DWORD lastError =
            Debug::ErrorWithLastError(L"CompareDirectories: LoadLibraryExW failed for '{}' (pluginId={})", entry->path.wstring(), entry->id);
        static_cast<void>(lastError);
        return std::nullopt;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        Debug::Error(L"CompareDirectories: Missing export RedSalamanderCreate in '{}' (pluginId={}).", entry->path.wstring(), entry->id);
        return std::nullopt;
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        return std::nullopt;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), fileSystem.put_void());

    if (FAILED(createHr) || ! fileSystem)
    {
        Debug::Error(L"CompareDirectories: RedSalamanderCreate failed for '{}' (pluginId={} hr=0x{:08X}).",
                     entry->path.wstring(),
                     requestedPluginId,
                     static_cast<unsigned long>(createHr));
        return std::nullopt;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
    if (FAILED(qiInfos) || ! informations)
    {
        Debug::Error(L"CompareDirectories: IInformations not supported by '{}' (pluginId={} hr=0x{:08X}).",
                     entry->path.wstring(),
                     entry->id,
                     static_cast<unsigned long>(qiInfos));
        return std::nullopt;
    }

    if (entry->informations)
    {
        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    if (! instanceContext.empty())
    {
        wil::com_ptr<IFileSystemInitialize> initializer;
        const HRESULT qiInit = fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
        if (FAILED(qiInit) || ! initializer)
        {
            Debug::Error(L"CompareDirectories: IFileSystemInitialize not supported (pluginId={} hr=0x{:08X}).", entry->id, static_cast<unsigned long>(qiInit));
            return std::nullopt;
        }

        const HRESULT initHr = initializer->Initialize(instanceContext.c_str(), nullptr);
        if (FAILED(initHr))
        {
            Debug::Error(L"CompareDirectories: Initialize failed (pluginId={} hr=0x{:08X}).", entry->id, static_cast<unsigned long>(initHr));
            return std::nullopt;
        }
    }

    CreatedFileSystemInstance created{};
    created.module        = std::move(module);
    created.fileSystem    = std::move(fileSystem);
    created.pluginShortId = entry->shortId;
    return created;
}

[[nodiscard]] std::wstring BuildMetadataDetailsText(bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes) noexcept
{
    std::wstring result;
    result.reserve(64);

    const auto fields = Common::FileMetadata::FormatDisplayFields(
        {.lastWriteTime100nsSince1601 = lastWriteTime, .fileAttributes = fileAttributes}, Common::FileMetadata::DisplayProfile::CompactDetails);

    auto appendToken = [&](std::wstring_view token) noexcept
    {
        if (token.empty())
        {
            return;
        }

        if (! result.empty())
        {
            result.append(L" • ");
        }
        result.append(token);
    };

    appendToken(fields.localTime);
    if (! isDirectory)
    {
        appendToken(FormatBytesCompact(sizeBytes));
    }
    appendToken(fields.attributes);

    return result;
}

struct CompareDetailsTextStrings
{
    std::wstring identical;
    std::wstring onlyInLeft;
    std::wstring onlyInRight;
    std::wstring typeMismatch;
    std::wstring bigger;
    std::wstring smaller;
    std::wstring newer;
    std::wstring older;
    std::wstring attributesDiffer;
    std::wstring contentDiffer;
    std::wstring contentComparing;
    std::wstring subdirAttributesDiffer;
    std::wstring subdirContentDiffer;
    std::wstring subdirComputing;
};

[[nodiscard]] const CompareDetailsTextStrings& GetCompareDetailsTextStrings() noexcept
{
    static const CompareDetailsTextStrings strings{
        .identical              = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_IDENTICAL),
        .onlyInLeft             = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_ONLY_IN_LEFT),
        .onlyInRight            = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_ONLY_IN_RIGHT),
        .typeMismatch           = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_TYPE_MISMATCH),
        .bigger                 = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_BIGGER),
        .smaller                = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_SMALLER),
        .newer                  = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_NEWER),
        .older                  = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_OLDER),
        .attributesDiffer       = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_ATTRIBUTES_DIFFER),
        .contentDiffer          = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_CONTENT_DIFFER),
        .contentComparing       = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_CONTENT_COMPARING),
        .subdirAttributesDiffer = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_SUBDIR_ATTRIBUTES_DIFFER),
        .subdirContentDiffer    = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_SUBDIR_CONTENT_DIFFER),
        .subdirComputing        = LoadStringResource(nullptr, IDS_COMPARE_DETAILS_SUBDIR_COMPUTING),
    };
    return strings;
}

CompareDirectoriesWindow::CompareDirectoriesWindow(FolderWindow& applicationFolderWindow,
                                                   Common::Settings::Settings& settings,
                                                   AppTheme theme,
                                                   const ShortcutManager* shortcuts,
                                                   CompareDirectoriesPaneContext left,
                                                   CompareDirectoriesPaneContext right) noexcept
    : _settings(&settings),
      _applicationFolderWindow(&applicationFolderWindow),
      _theme(std::move(theme)),
      _shortcuts(shortcuts),
      _leftContext(std::move(left)),
      _rightContext(std::move(right))
{
}

ATOM CompareDirectoriesWindow::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kCompareDirectoriesWindowClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK CompareDirectoriesWindow::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self     = cs ? reinterpret_cast<CompareDirectoriesWindow*>(cs->lpCreateParams) : nullptr;
        if (self)
        {
            self->_hWnd.reset(hwnd);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    ++self->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([self]() noexcept
    {
        if (self->_dispatchDepth > 0u)
        {
            --self->_dispatchDepth;
        }
        if (self->_dispatchDepth == 0u && self->_deletePending && ! self->_createInProgress)
        {
            delete self;
        }
    });

    return self->WndProc(hwnd, msg, wp, lp);
}

LRESULT CompareDirectoriesWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: OnNcDestroy(); return 0;
        case WM_SIZE: OnSize(); return 0;
        case WM_DPICHANGED: OnDpiChanged(HIWORD(wp), reinterpret_cast<RECT*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 760, 480);
            }
            return 0;
        case WM_COMMAND: OnCommand(LOWORD(wp)); return 0;
        case WndMsg::kFunctionBarInvoke: return OnFunctionBarInvoke(wp, lp);
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_TIMER:
            if (wp == kCompareTaskAutoDismissTimerId)
            {
                KillTimer(hwnd, kCompareTaskAutoDismissTimerId);
                DismissCompareTaskCard();
                return 0;
            }
            if (wp == kCompareTaskCardTrailingFlushTimerId)
            {
                OnCompareTaskCardTrailingFlushTimer();
                return 0;
            }
            if (wp == kCompareProgressPulseTimerId)
            {
                OnProgressPulseTimer();
                return 0;
            }
            if (wp == kCompareDecisionRefreshTimerId)
            {
                OnDecisionRefreshTimer();
                return 0;
            }
            break;
        case WM_ACTIVATE:
            if (_hWnd)
            {
                const bool windowActive = LOWORD(wp) != WA_INACTIVE;
                ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive);
            }
            return 0;
        case WM_CTLCOLORSTATIC:
        {
            const LRESULT result = OnCtlColorStatic(reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
            if (result != 0)
            {
                return result;
            }
            break;
        }
        case WM_SYSKEYDOWN:
            if ((wp == VK_F10 || wp == VK_MENU) && _chrome.usesMenuBar)
            {
                return FocusFirstDxMenuBarItem() ? 0 : DefWindowProcW(hwnd, msg, wp, lp);
            }
            break;
        case WM_SYSCHAR:
            if (_chrome.usesMenuBar && wp >= 0x20u && ActivateDxMenuBarMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            break;
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONUP: OnLButtonUp(); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(); return 0;
        case WM_SETCURSOR:
        {
            const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
            POINT pt{GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
            ScreenToClient(hwnd, &pt);
            if (OnSetCursor(pt))
            {
                return TRUE;
            }
            break;
        }
        case WndMsg::kCompareDirectoriesDeferredStart: return OnDeferredBeginOrRescanCompare(wp);
        case WndMsg::kCompareDirectoriesScanProgress: return OnScanProgress(wp, lp);
        case WndMsg::kCompareDirectoriesContentProgress: return OnContentProgress(wp, lp);
        case WndMsg::kCompareDirectoriesDecisionUpdated:
            if (_compareActive && _session && static_cast<uint64_t>(wp) == _compareRunId)
            {
                ScheduleDecisionRefresh();
            }
            return 0;
        case WndMsg::kCompareDirectoriesDecisionRefreshNow:
            _decisionRefreshFallbackPending = false;
            if (_compareActive && _session && static_cast<uint64_t>(wp) == _compareRunId)
            {
                OnDecisionRefreshTimer();
            }
            return 0;
        case WndMsg::kCompareDirectoriesExecuteCommand: return OnExecuteShortcutCommand(lp);
        case WndMsg::kCompareDirectoriesLeaveScopePrompt: return OnLeaveScopePrompt();
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

bool CompareDirectoriesWindow::Create(HWND owner) noexcept
{
    UpdateOwnerWindow(owner);

    const HINSTANCE instance = GetModuleHandleW(nullptr);
    if (! RegisterWndClass(instance))
    {
        return false;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_COMPARE_DIRECTORIES_TITLE);

    _hasSavedPlacement = _settings && _settings->windows.find(std::wstring(kCompareDirectoriesWindowId)) != _settings->windows.end();

    HWND placementOwner = owner;
    if (placementOwner && IsWindow(placementOwner))
    {
        placementOwner = GetAncestor(placementOwner, GA_ROOT);
    }
    else
    {
        placementOwner = nullptr;
    }

    wil::unique_hmenu menu(Localization::LoadMenuResource(instance, IDR_COMPARE_DIRECTORIES_MENU));

    int x = CW_USEDEFAULT;
    int y = CW_USEDEFAULT;
    int w = 1100;
    int h = 700;
    if (! _hasSavedPlacement && placementOwner)
    {
        WINDOWPLACEMENT wp{};
        wp.length = sizeof(wp);
        if (GetWindowPlacement(placementOwner, &wp) != 0)
        {
            const RECT rc = wp.rcNormalPosition;
            x             = rc.left;
            y             = rc.top;
            w             = std::max(1, static_cast<int>(rc.right - rc.left));
            h             = std::max(1, static_cast<int>(rc.bottom - rc.top));
        }
        else
        {
            RECT rc{};
            if (GetWindowRect(placementOwner, &rc) != 0)
            {
                x = rc.left;
                y = rc.top;
                w = std::max(1, static_cast<int>(rc.right - rc.left));
                h = std::max(1, static_cast<int>(rc.bottom - rc.top));
            }
        }

        _restoreShowCmd = IsZoomed(placementOwner) ? SW_MAXIMIZE : SW_SHOWNORMAL;
    }

    _createInProgress                = true;
    const auto clearCreateInProgress = wil::scope_exit([&]() noexcept { _createInProgress = false; });
    HWND created                     = CreateWindowExW(0,
                                                       kCompareDirectoriesWindowClassName,
                                                       title.c_str(),
                                                       WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                                       x,
                                                       y,
                                                       w,
                                                       h,
                                                       nullptr,
                                                       menu.get(),
                                                       instance,
                                                       this);
    if (created || _deletePending || _hWnd)
    {
        menu.release();
    }
    if (! created || _deletePending || ! _hWnd)
    {
        return false;
    }

    ShowWindow(created, _restoreShowCmd);
    return true;
}

HWND CompareDirectoriesWindow::ResolveRestoreFolderViewWindow() const noexcept
{
    const HWND folderWindow = _applicationFolderWindow->GetHwnd();
    if (! folderWindow || IsWindow(folderWindow) == FALSE)
    {
        return nullptr;
    }

    HWND focusedFolderView = _applicationFolderWindow->GetFocusedFolderViewHwnd();
    if (! focusedFolderView)
    {
        focusedFolderView = _applicationFolderWindow->GetFolderViewHwnd(_applicationFolderWindow->GetFocusedPane());
    }

    if (! focusedFolderView || IsWindow(focusedFolderView) == FALSE)
    {
        return nullptr;
    }

    return focusedFolderView;
}

void CompareDirectoriesWindow::UpdateOwnerWindow(HWND owner) noexcept
{
    _ownerWindow             = (owner && IsWindow(owner) != FALSE) ? GetAncestor(owner, GA_ROOT) : nullptr;
    _restoreFolderViewWindow = ResolveRestoreFolderViewWindow();
}

void CompareDirectoriesWindow::RestoreOwnerFocusAfterClose() noexcept
{
    const HWND restoreOwner = (_ownerWindow && IsWindow(_ownerWindow) != FALSE) ? _ownerWindow : nullptr;
    const HWND restoreFocus =
        (_restoreFolderViewWindow && IsWindow(_restoreFolderViewWindow) != FALSE) ? _restoreFolderViewWindow : ResolveRestoreFolderViewWindow();

    if (restoreOwner)
    {
        if (IsIconic(restoreOwner) != FALSE)
        {
            ShowWindow(restoreOwner, SW_RESTORE);
        }
        static_cast<void>(SetActiveWindow(restoreOwner));
        static_cast<void>(SetForegroundWindow(restoreOwner));
    }

    if (restoreFocus)
    {
        _applicationFolderWindow->RequestRestoreFolderViewFocus(restoreFocus);
    }
    else if (restoreOwner)
    {
        static_cast<void>(SetFocus(restoreOwner));
    }
}

bool CompareDirectoriesWindow::OnCreate(HWND hwnd) noexcept
{
#ifdef ENABLE_TESTS
    if (g_failNextCompareDirectoriesWindowCreate.exchange(false, std::memory_order_acq_rel))
    {
        return false;
    }
#endif

    _dpi = GetDpiForWindow(hwnd);
    g_compareDirectoriesWindows.push_back(hwnd);
    if (_settings && _hasSavedPlacement)
    {
        _restoreShowCmd = WindowPlacementPersistence::Restore(*_settings, kCompareDirectoriesWindowId, hwnd);
    }

    ApplyTheme();
    CreateChildWindows(hwnd);
    UpdateViewMenuChecks();
    ApplyTheme();
    Layout();
    ShowOptionsPanel(true);
    return true;
}

void CompareDirectoriesWindow::OnDestroy() noexcept
{
    _leaveScopePromptRequest.reset();
    _leaveScopePromptPending = false;

    if (_session)
    {
        LogComparePerfStats(L"window_destroy", _session, _progress.compareRunResultHr);
    }

    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kCompareTaskAutoDismissTimerId);
        KillTimer(_hWnd.get(), kCompareTaskCardTrailingFlushTimerId);
        KillTimer(_hWnd.get(), kCompareProgressPulseTimerId);
        KillTimer(_hWnd.get(), kCompareDecisionRefreshTimerId);
        _progress.pulseTimerActive             = false;
        _progress.taskCardTrailingFlushPending = false;
        _decisionRefreshTimerActive            = false;
        _decisionRefreshFallbackPending        = false;
        _decisionRefreshPending                = false;
    }
    DismissCompareTaskCard();

    if (_settings && _hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kCompareDirectoriesWindowId, _hWnd.get());
    }

    if (_session)
    {
        _session->SetScanProgressCallback({});
        _session->SetContentProgressCallback({});
        _session->SetDecisionUpdatedCallback({});
        _session->SetBackgroundWorkEnabled(false);
    }

    _folderWindow.SetShowSortMenuCallback({});
    _folderWindow.SetPanePathChangedCallback({});
    _folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
    _folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right, {});
    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Left, {});
    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Right, {});
    _folderWindow.RemoveFileOperationCompletedCallback(std::exchange(_fileOperationCompletedCallbackToken, 0ull));

    DetachOptionsDxButtonHosts();
    DetachOptionsDxStaticHosts();
    _optionsPanel.dxUi.reset();
    _optionsPanel.ui = {};
    _optionsPanel.cards.clear();
    _optionsPanel.scrollOffset = 0;
    _optionsPanel.scrollMax    = 0;

    _optionsPanel.dlg.reset();
    _progress.scanProgressText.reset();
    if (_progress.scanProgressBar && IsWindow(_progress.scanProgressBar.get()) != FALSE)
    {
        RemovePropW(_progress.scanProgressBar.get(), kCompareProgressSpinnerStateProp);
    }
    _progress.scanProgressBar.reset();
    _chrome.bannerTitle.reset();
    DetachDxChromeHosts();
    _chrome.bannerOptionsButton.reset();
    _chrome.bannerRescanButton.reset();
    _chrome.menuHandle.reset();

    if (_session || _fsLeft || _fsRight || _leftBaseFs || _rightBaseFs || _leftBaseModule || _rightBaseModule)
    {
        struct CompareDestroyCleanup final
        {
            wil::unique_hmodule leftModule;
            wil::unique_hmodule rightModule;
            wil::com_ptr<IFileSystem> leftBaseFs;
            wil::com_ptr<IFileSystem> rightBaseFs;
            wil::com_ptr<IFileSystem> fsLeft;
            wil::com_ptr<IFileSystem> fsRight;
            std::shared_ptr<CompareDirectoriesSession> session;

            CompareDestroyCleanup()                                        = default;
            CompareDestroyCleanup(const CompareDestroyCleanup&)            = delete;
            CompareDestroyCleanup(CompareDestroyCleanup&&)                 = delete;
            CompareDestroyCleanup& operator=(const CompareDestroyCleanup&) = delete;
            CompareDestroyCleanup& operator=(CompareDestroyCleanup&&)      = delete;
        };

        auto cleanup         = std::make_unique<CompareDestroyCleanup>();
        cleanup->session     = std::move(_session);
        cleanup->fsLeft      = std::move(_fsLeft);
        cleanup->fsRight     = std::move(_fsRight);
        cleanup->leftBaseFs  = std::move(_leftBaseFs);
        cleanup->rightBaseFs = std::move(_rightBaseFs);
        cleanup->leftModule  = std::move(_leftBaseModule);
        cleanup->rightModule = std::move(_rightBaseModule);

        QueueCompareCleanup(std::move(cleanup), L"window destroy cleanup");
    }

    _folderWindow.Destroy();
}

void CompareDirectoriesWindow::OnNcDestroy() noexcept
{
    if (_hWnd)
    {
        std::erase(g_compareDirectoriesWindows, _hWnd.get());
    }

    if (_hWnd)
    {
        static_cast<void>(DrainPostedPayloadsForWindow(_hWnd.get()));
        SetWindowLongPtrW(_hWnd.get(), GWLP_USERDATA, 0);
        _hWnd.release();
    }
    RestoreOwnerFocusAfterClose();
    _ownerWindow             = nullptr;
    _restoreFolderViewWindow = nullptr;
    _deletePending           = true;
    if (_dispatchDepth == 0u && ! _createInProgress)
    {
        delete this;
    }
}

void CompareDirectoriesWindow::OnSize() noexcept
{
    Layout();
}

void CompareDirectoriesWindow::OnDpiChanged(UINT newDpi, const RECT* newRect) noexcept
{
    _dpi = newDpi;

    if (newRect && _hWnd)
    {
        SetWindowPos(
            _hWnd.get(), nullptr, newRect->left, newRect->top, newRect->right - newRect->left, newRect->bottom - newRect->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    _folderWindow.OnDpiChanged(static_cast<float>(_dpi));
    ApplyTheme();
    Layout();
}

void CompareDirectoriesWindow::OnCommand(UINT id) noexcept
{
    switch (id)
    {
        case IDM_VIEW_SWITCH_PANE_FOCUS:
        {
            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);
            if (const HWND view = _folderWindow.GetFolderViewHwnd(pane))
            {
                SendMessageW(view, WM_KEYDOWN, static_cast<WPARAM>(VK_TAB), 0);
            }
            break;
        }
        case IDM_PANE_RENAME:
        case IDM_PANE_VIEW:
        case IDM_PANE_ALTERNATE_VIEW:
        case IDM_PANE_VIEW_SPACE:
        case IDM_PANE_COPY_TO_OTHER:
        case IDM_PANE_MOVE_TO_OTHER:
        case IDM_PANE_CREATE_DIR:
        case IDM_PANE_DELETE:
        case IDM_PANE_PERMANENT_DELETE:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);

            switch (id)
            {
                case IDM_PANE_RENAME: _folderWindow.CommandRename(pane); break;
                case IDM_PANE_VIEW: _folderWindow.CommandView(pane); break;
                case IDM_PANE_ALTERNATE_VIEW: _folderWindow.CommandAlternateView(pane); break;
                case IDM_PANE_VIEW_SPACE: _folderWindow.CommandViewSpace(pane); break;
                case IDM_PANE_COPY_TO_OTHER: static_cast<void>(StartCompareSyncToOtherPane(pane, FILESYSTEM_COPY)); break;
                case IDM_PANE_MOVE_TO_OTHER: static_cast<void>(StartCompareSyncToOtherPane(pane, FILESYSTEM_MOVE)); break;
                case IDM_PANE_CREATE_DIR: _folderWindow.CommandCreateDirectory(pane); break;
                case IDM_PANE_DELETE: _folderWindow.CommandDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE: _folderWindow.CommandPermanentDelete(pane); break;
            }
            break;
        }
        case IDM_LEFT_SORT_NAME:
        case IDM_LEFT_SORT_EXTENSION:
        case IDM_LEFT_SORT_TIME:
        case IDM_LEFT_SORT_SIZE:
        case IDM_LEFT_SORT_ATTRIBUTES:
        case IDM_RIGHT_SORT_NAME:
        case IDM_RIGHT_SORT_EXTENSION:
        case IDM_RIGHT_SORT_TIME:
        case IDM_RIGHT_SORT_SIZE:
        case IDM_RIGHT_SORT_ATTRIBUTES:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id >= IDM_RIGHT_SORT_NAME ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
            _folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (id)
            {
                case IDM_LEFT_SORT_NAME:
                case IDM_RIGHT_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_LEFT_SORT_EXTENSION:
                case IDM_RIGHT_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_LEFT_SORT_TIME:
                case IDM_RIGHT_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_LEFT_SORT_SIZE:
                case IDM_RIGHT_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_LEFT_SORT_ATTRIBUTES:
                case IDM_RIGHT_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            _folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_LEFT_SORT_NONE:
        case IDM_RIGHT_SORT_NONE:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id == IDM_RIGHT_SORT_NONE ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
            _folderWindow.SetActivePane(pane);
            _folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_SORT_NAME:
        case IDM_PANE_SORT_EXTENSION:
        case IDM_PANE_SORT_TIME:
        case IDM_PANE_SORT_SIZE:
        case IDM_PANE_SORT_ATTRIBUTES:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (id)
            {
                case IDM_PANE_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_PANE_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_PANE_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_PANE_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_PANE_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            _folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_PANE_SORT_NONE:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);
            _folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_DISPLAY_BRIEF:
        case IDM_PANE_DISPLAY_DETAILED:
        case IDM_PANE_DISPLAY_EXTRA_DETAILED:
        {
            if (! _compareStarted)
            {
                break;
            }

            FolderView::DisplayMode mode = FolderView::DisplayMode::Detailed;
            switch (id)
            {
                case IDM_PANE_DISPLAY_BRIEF: mode = FolderView::DisplayMode::Brief; break;
                case IDM_PANE_DISPLAY_DETAILED: mode = FolderView::DisplayMode::Detailed; break;
                case IDM_PANE_DISPLAY_EXTRA_DETAILED: mode = FolderView::DisplayMode::ExtraDetailed; break;
            }

            _compareDisplayMode = mode;
            _folderWindow.SetDisplayMode(FolderWindow::Pane::Left, mode);
            _folderWindow.SetDisplayMode(FolderWindow::Pane::Right, mode);
            _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Left);
            _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Right);
            UpdateViewMenuChecks();
            break;
        }
        case IDM_LEFT_REFRESH:
        case IDM_RIGHT_REFRESH:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id == IDM_LEFT_REFRESH ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
            _folderWindow.SetActivePane(pane);
            if (const auto currentPath = _folderWindow.GetCurrentPluginPath(pane))
            {
                ApplySelectionForFolder(pane == FolderWindow::Pane::Left ? ComparePane::Left : ComparePane::Right, currentPath.value());
            }
            _folderWindow.CommandRefresh(pane);
            break;
        }
        case IDM_COMPARE_OPTIONS:
            if (_optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != 0)
            {
                break;
            }
            ShowOptionsPanel(true);
            break;
        case IDM_COMPARE_RESCAN:
            if (_compareActive && (_compareRunPending || _progress.banner.scanActiveScans > 0u || _progress.banner.contentPendingCompares > 0u) && _session)
            {
                _progress.compareRunResultHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                _session->SetBackgroundWorkEnabled(false);
                _session->Invalidate();
                UpdateCompareWatermark();
            }
            else
            {
                ScheduleBeginOrRescanCompare();
            }
            break;
        case IDM_COMPARE_TOGGLE_IDENTICAL:
        {
            if (! _settings)
            {
                break;
            }

            Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
            if (! s.keepIdenticalItems)
            {
                break;
            }

            s.showIdenticalItems          = ! s.showIdenticalItems;
            _settings->compareDirectories = s;
            if (_session)
            {
                _session->SetSettings(s);
            }

            UpdateViewMenuChecks();
            RefreshBothPanes();
            break;
        }
        case IDM_COMPARE_RESTORE_DIFFERENCES_SELECTION:
        {
            if (! _compareStarted)
            {
                break;
            }

            _selectionMode = CompareSelectionMode::Default;
            if (const auto leftPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left))
            {
                ApplySelectionForFolder(ComparePane::Left, leftPath.value());
            }
            if (const auto rightPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right))
            {
                ApplySelectionForFolder(ComparePane::Right, rightPath.value());
            }
            break;
        }
        case IDM_COMPARE_INVERT_DIFFERENCES_SELECTION:
        {
            if (! _compareStarted || ! _session)
            {
                break;
            }

            _selectionMode = CompareSelectionMode::Inverted;
            if (const auto leftPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left))
            {
                ApplySelectionForFolder(ComparePane::Left, leftPath.value());
            }
            if (const auto rightPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right))
            {
                ApplySelectionForFolder(ComparePane::Right, rightPath.value());
            }
            _selectionMode = CompareSelectionMode::Default;
            break;
        }
        case IDM_COMPARE_CLOSE:
            if (_hWnd)
            {
                PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
            }
            break;
    }
}

LRESULT CompareDirectoriesWindow::OnFunctionBarInvoke(WPARAM wParam, LPARAM lParam) noexcept
{
    if (! _hWnd || ! _shortcuts)
    {
        return 0;
    }

    const uint32_t vk        = static_cast<uint32_t>(wParam);
    const uint32_t modifiers = static_cast<uint32_t>(lParam) & 0x7u;

    const std::optional<std::wstring_view> commandOpt = _shortcuts->FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return 0;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (ShortcutIds::IsUnassignedCommandId(commandId))
    {
        return 0;
    }
    if (commandId.starts_with(L"cmd/app/"))
    {
        // App-scoped commands are handled by the main window's message loop.
        return 0;
    }

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (! wmCommandOpt.has_value())
    {
        ExecuteShortcutCommand(commandId);
        return 0;
    }

    const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
    SendMessageW(_hWnd.get(), WM_COMMAND, wp, 0);
    return 0;
}

void CompareDirectoriesWindow::OnPaint() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);

    HBRUSH bg = _backgroundBrush ? _backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc.get(), &ps.rcPaint, bg);

    if (_splitterBrush)
    {
        RECT splitter = _splitterRect;
        RECT paint    = ps.rcPaint;
        RECT intersect{};
        if (IntersectRect(&intersect, &splitter, &paint))
        {
            FillRect(hdc.get(), &intersect, _splitterBrush.get());

            if (_splitterGripBrush)
            {
                const int dpi            = static_cast<int>(_dpi);
                const int dotSize        = std::max(1, MulDiv(kSplitterGripDotSizeDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int dotGap         = std::max(1, MulDiv(kSplitterGripDotGapDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int gripHeight     = (dotSize * kSplitterGripDotCount) + (dotGap * (kSplitterGripDotCount - 1));
                const int splitterWidth  = splitter.right - splitter.left;
                const int splitterHeight = splitter.bottom - splitter.top;

                if (splitterWidth > 0 && splitterHeight >= gripHeight)
                {
                    const int left = splitter.left + (splitterWidth - dotSize) / 2;
                    const int top  = splitter.top + (splitterHeight - gripHeight) / 2;

                    for (int i = 0; i < kSplitterGripDotCount; ++i)
                    {
                        const int dotTop = top + i * (dotSize + dotGap);
                        RECT dotRect{};
                        dotRect.left   = left;
                        dotRect.top    = dotTop;
                        dotRect.right  = left + dotSize;
                        dotRect.bottom = dotTop + dotSize;
                        FillRect(hdc.get(), &dotRect, _splitterGripBrush.get());
                    }
                }
            }
        }
    }
}

void CompareDirectoriesWindow::ExecuteShortcutCommand(std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return;
    }

    const std::wstring_view originalCommandId = commandId;
    std::optional<wchar_t> driveRootLetter;
    {
        constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
        if (originalCommandId.starts_with(kGoDriveRootPrefix) && originalCommandId.size() > kGoDriveRootPrefix.size())
        {
            const wchar_t rawLetter = originalCommandId[kGoDriveRootPrefix.size()];
            if (std::iswalpha(static_cast<wint_t>(rawLetter)) != 0)
            {
                const wchar_t upper = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(rawLetter)));
                if (upper >= L'A' && upper <= L'Z')
                {
                    driveRootLetter = upper;
                }
            }
        }
    }

    commandId = CanonicalizeCommandId(commandId);
    if (ShortcutIds::IsUnassignedCommandId(commandId))
    {
        return;
    }

    if (commandId == L"cmd/pane/menu")
    {
        if (_hWnd)
        {
            SendMessageW(_hWnd.get(), WM_SYSCOMMAND, SC_KEYMENU, 0);
        }
        return;
    }

    const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
    _folderWindow.SetActivePane(pane);

    const auto sendKeyToPaneFolderView = [&](uint32_t vk) noexcept
    {
        if (const HWND view = _folderWindow.GetFolderViewHwnd(pane))
        {
            SendMessageW(view, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
        }
    };

    if (commandId == L"cmd/pane/focusAddressBar")
    {
        _folderWindow.CommandFocusAddressBar(pane);
        return;
    }
    if (commandId == L"cmd/pane/upOneDirectory")
    {
        sendKeyToPaneFolderView(VK_BACK);
        return;
    }
    if (commandId == L"cmd/pane/switchPaneFocus")
    {
        sendKeyToPaneFolderView(VK_TAB);
        return;
    }
    if (commandId == L"cmd/pane/zoomPanel")
    {
        _folderWindow.ToggleZoomPanel(pane);
        return;
    }
    if (commandId == L"cmd/pane/refresh")
    {
        _folderWindow.CommandRefresh(pane);
        return;
    }
    if (commandId == L"cmd/pane/executeOpen")
    {
        sendKeyToPaneFolderView(VK_RETURN);
        return;
    }
    if (commandId == L"cmd/pane/selectCalculateDirectorySizeNext")
    {
        sendKeyToPaneFolderView(VK_SPACE);
        return;
    }
    if (commandId == L"cmd/pane/selectNext")
    {
        sendKeyToPaneFolderView(VK_INSERT);
        return;
    }
    if (commandId == L"cmd/pane/moveToRecycleBin")
    {
        sendKeyToPaneFolderView(VK_DELETE);
        return;
    }
    if (commandId == L"cmd/pane/goDriveRoot")
    {
        const auto getDefaultRoot = []() noexcept -> std::filesystem::path
        {
            wchar_t buffer[MAX_PATH] = {};
            const UINT bufferSize    = static_cast<UINT>(ARRAYSIZE(buffer));
            const UINT length        = GetWindowsDirectoryW(buffer, bufferSize);
            if (length > 0 && length < bufferSize)
            {
                const std::filesystem::path root = std::filesystem::path(buffer).root_path();
                if (! root.empty())
                {
                    return root;
                }
            }
            return std::filesystem::path(L"C:\\");
        };

        if (! driveRootLetter.has_value())
        {
            _folderWindow.SetFolderPath(pane, getDefaultRoot());
            return;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveRootLetter.value());
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_NO_ROOT_DIR)
        {
            return;
        }

        _folderWindow.SetFolderPath(pane, std::filesystem::path(driveRoot));
        return;
    }
}

LRESULT CompareDirectoriesWindow::OnCtlColorStatic(HDC hdc, HWND /*control*/) noexcept
{
    if (! hdc || ! _backgroundBrush)
    {
        return 0;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, _theme.menu.text);
    SetBkColor(hdc, _theme.windowBackground);
    return reinterpret_cast<LRESULT>(_backgroundBrush.get());
}

void CompareDirectoriesWindow::OnLButtonDown(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (PtInRect(&_splitterRect, pt))
    {
        _draggingSplitter     = true;
        _splitterDragOffsetPx = pt.x - _splitterRect.left;
        SetCapture(_hWnd.get());
    }
}

void CompareDirectoriesWindow::OnLButtonDblClk(POINT pt) noexcept
{
    if (! PtInRect(&_splitterRect, pt))
    {
        return;
    }

    _draggingSplitter = false;
    ReleaseCapture();
    SetSplitRatio(0.5f);
}

void CompareDirectoriesWindow::OnLButtonUp() noexcept
{
    if (_draggingSplitter)
    {
        _draggingSplitter = false;
        ReleaseCapture();
    }
}

void CompareDirectoriesWindow::OnMouseMove(POINT pt) noexcept
{
    if (! _draggingSplitter)
    {
        return;
    }

    const int splitterWidth  = _splitterRect.right - _splitterRect.left;
    const int availableWidth = std::max(0L, _clientSize.cx - splitterWidth);
    if (availableWidth <= 0)
    {
        return;
    }

    int desiredLeftWidth = pt.x - _splitterDragOffsetPx;
    desiredLeftWidth     = std::clamp(desiredLeftWidth, 0, availableWidth);

    const float ratio = static_cast<float>(desiredLeftWidth) / static_cast<float>(availableWidth);
    SetSplitRatio(ratio);

    if (_hWnd)
    {
        UpdateWindow(_hWnd.get());
    }
}

void CompareDirectoriesWindow::OnCaptureChanged() noexcept
{
    _draggingSplitter = false;
}

bool CompareDirectoriesWindow::OnSetCursor(POINT pt) noexcept
{
    if (PtInRect(&_splitterRect, pt))
    {
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        return true;
    }
    return false;
}

void CompareDirectoriesWindow::SetSplitRatio(float ratio) noexcept
{
    const RECT oldSplitter = _splitterRect;
    _splitRatio            = std::clamp(ratio, kMinSplitRatio, kMaxSplitRatio);
    Layout();
    if (_hWnd)
    {
        RECT invalid = oldSplitter;
        if (IsRectEmpty(&invalid))
        {
            invalid = _splitterRect;
        }
        else if (! IsRectEmpty(&_splitterRect))
        {
            UnionRect(&invalid, &invalid, &_splitterRect);
        }

        if (! IsRectEmpty(&invalid))
        {
            InvalidateRect(_hWnd.get(), &invalid, TRUE);
        }
    }
}

void CompareDirectoriesWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    Layout();
}

void CompareDirectoriesWindow::ApplyTheme() noexcept
{
    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    _menuBackgroundBrush.reset(CreateSolidBrush(_theme.menu.background));
    _optionsPanel.backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));

    const COLORREF surface = UiMetrics::GetControlSurfaceColor(_theme);
    _optionsPanel.cardBrush.reset(CreateSolidBrush(surface));
    _optionsPanel.inputBackgroundColor         = UiMetrics::BlendColorRefWeightedTruncate(surface, _theme.windowBackground, _theme.dark ? 50 : 30, 255);
    _optionsPanel.inputFocusedBackgroundColor  = UiMetrics::BlendColorRefWeightedTruncate(_optionsPanel.inputBackgroundColor, _theme.menu.text, _theme.dark ? 20 : 16, 255);
    _optionsPanel.inputDisabledBackgroundColor = UiMetrics::BlendColorRefWeightedTruncate(_theme.windowBackground, _optionsPanel.inputBackgroundColor, _theme.dark ? 70 : 40, 255);
    _optionsPanel.inputBrush.reset(CreateSolidBrush(_optionsPanel.inputBackgroundColor));
    _optionsPanel.inputFocusedBrush.reset(CreateSolidBrush(_optionsPanel.inputFocusedBackgroundColor));
    _optionsPanel.inputDisabledBrush.reset(CreateSolidBrush(_optionsPanel.inputDisabledBackgroundColor));

    if (_progress.scanProgressBar)
    {
        InvalidateRect(_progress.scanProgressBar.get(), nullptr, FALSE);
    }

    _folderWindow.ApplyTheme(_theme);

    const wchar_t* folderViewThemeName = L"";
    if (! _theme.highContrast)
    {
        folderViewThemeName = _theme.dark ? L"DarkMode_Explorer" : L"Explorer";
    }

    if (const HWND leftView = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left))
    {
        SetWindowTheme(leftView, folderViewThemeName, nullptr);
        SendMessageW(leftView, WM_THEMECHANGED, 0, 0);
    }
    if (const HWND rightView = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right))
    {
        SetWindowTheme(rightView, folderViewThemeName, nullptr);
        SendMessageW(rightView, WM_THEMECHANGED, 0, 0);
    }

    ApplyOptionsDialogTheme();
    ApplyDxChromeTheme();
    SyncDxChrome();

    if (_hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive);
        ApplyWindowBackdropTheme(_hWnd.get(), _theme, WindowBackdropTarget::Primary);
        RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
    }
}

void CompareDirectoriesWindow::OnPanePathChanged(ComparePane pane, const std::optional<std::filesystem::path>& newPath) noexcept
{
    std::optional<std::filesystem::path>& last          = (pane == ComparePane::Left) ? _lastLeftPluginPath : _lastRightPluginPath;
    const std::optional<std::filesystem::path> previous = last;
    last                                                = newPath;
    SyncOtherPanePath(pane, previous, newPath);
}

void CompareDirectoriesWindow::CreateChildWindows(HWND hwnd) noexcept
{
    FolderView::RegisterWndClass(GetModuleHandleW(nullptr));

    const HINSTANCE instance             = GetModuleHandleW(nullptr);
    const std::wstring bannerTitleText   = LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE);
    const std::wstring bannerOptionsText = LoadStringResource(nullptr, IDS_COMPARE_BANNER_OPTIONS_ELLIPSIS);
    const std::wstring bannerRescanText  = LoadStringResource(nullptr, IDS_COMPARE_BANNER_RESCAN);
    _chrome.bannerTitle.reset(CreateWindowExW(
        0, L"Static", bannerTitleText.c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE | SS_NOPREFIX, 0, 0, 10, 10, hwnd, nullptr, instance, nullptr));

    _chrome.bannerOptionsButton.reset(CreateWindowExW(0,
                                                      L"Button",
                                                      bannerOptionsText.c_str(),
                                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      hwnd,
                                                      reinterpret_cast<HMENU>(IDM_COMPARE_OPTIONS),
                                                      instance,
                                                      nullptr));

    _chrome.bannerRescanButton.reset(CreateWindowExW(0,
                                                     L"Button",
                                                     bannerRescanText.c_str(),
                                                     WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     hwnd,
                                                     reinterpret_cast<HMENU>(IDM_COMPARE_RESCAN),
                                                     instance,
                                                     nullptr));

    _progress.scanProgressText.reset(CreateWindowExW(0,
                                                     L"Static",
                                                     L"",
                                                     WS_CHILD | SS_LEFT | SS_NOPREFIX | SS_PATHELLIPSIS,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     hwnd,
                                                     reinterpret_cast<HMENU>(kScanProgressTextId),
                                                     instance,
                                                     nullptr));
    if (EnsureCompareProgressSpinnerWindowClass(instance))
    {
        _progress.scanProgressBar.reset(CreateWindowExW(
            0, kCompareProgressSpinnerClassName, nullptr, WS_CHILD, 0, 0, 10, 10, hwnd, reinterpret_cast<HMENU>(kScanProgressBarId), instance, nullptr));
    }
    if (_progress.scanProgressBar)
    {
        if (SetPropW(_progress.scanProgressBar.get(), kCompareProgressSpinnerStateProp, this) == 0)
        {
            Debug::ErrorWithLastError(L"CompareDirectories: failed to attach progress spinner host state.");
            _progress.scanProgressBar.reset();
        }
    }

    if (_progress.scanProgressText)
    {
        ShowWindow(_progress.scanProgressText.get(), SW_HIDE);
    }
    if (_progress.scanProgressBar)
    {
        ShowWindow(_progress.scanProgressBar.get(), SW_HIDE);
    }

    static_cast<void>(EnsureDxChromeHosts());

    _folderWindow.Create(hwnd, 0, 0, 10, 10);
    _folderWindow.SetSettings(_settings);
    _folderWindow.SetShortcutManager(_shortcuts);
    _folderWindow.SetShowSortMenuCallback([this](FolderWindow::Pane pane, POINT screenPoint) noexcept { ShowSortMenuPopup(pane, screenPoint); });

    bool functionBarVisible = true;
    if (_settings && _settings->mainMenu.has_value())
    {
        functionBarVisible = _settings->mainMenu->functionBarVisible;
    }
    _folderWindow.SetFunctionBarVisible(functionBarVisible);

    _folderWindow.SetPanePathChangedCallback([this](FolderWindow::Pane pane, const std::optional<std::filesystem::path>& pluginPath)
    { OnPanePathChanged(pane == FolderWindow::Pane::Left ? ComparePane::Left : ComparePane::Right, pluginPath); });

    _folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                      [this](const std::filesystem::path& folder)
    {
        ApplySelectionForFolder(ComparePane::Left, folder);
        UpdateEmptyStateForFolder(ComparePane::Left, folder);
    });
    _folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right,
                                                      [this](const std::filesystem::path& folder)
    {
        ApplySelectionForFolder(ComparePane::Right, folder);
        UpdateEmptyStateForFolder(ComparePane::Right, folder);
    });

    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Left,
                                             [this](const std::filesystem::path& folder,
                                                    std::wstring_view displayName,
                                                    bool isDirectory,
                                                    uint64_t sizeBytes,
                                                    int64_t lastWriteTime,
                                                    DWORD fileAttributes) noexcept -> std::wstring
    { return BuildDetailsTextForCompareItem(ComparePane::Left, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Right,
                                             [this](const std::filesystem::path& folder,
                                                    std::wstring_view displayName,
                                                    bool isDirectory,
                                                    uint64_t sizeBytes,
                                                    int64_t lastWriteTime,
                                                    DWORD fileAttributes) noexcept -> std::wstring
    { return BuildDetailsTextForCompareItem(ComparePane::Right, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _folderWindow.SetPaneMetadataTextProvider(FolderWindow::Pane::Left,
                                              [this](const std::filesystem::path& folder,
                                                     std::wstring_view displayName,
                                                     bool isDirectory,
                                                     uint64_t sizeBytes,
                                                     int64_t lastWriteTime,
                                                     DWORD fileAttributes) noexcept -> std::wstring
    { return BuildMetadataTextForCompareItem(ComparePane::Left, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _folderWindow.SetPaneMetadataTextProvider(FolderWindow::Pane::Right,
                                              [this](const std::filesystem::path& folder,
                                                     std::wstring_view displayName,
                                                     bool isDirectory,
                                                     uint64_t sizeBytes,
                                                     int64_t lastWriteTime,
                                                     DWORD fileAttributes) noexcept -> std::wstring
    { return BuildMetadataTextForCompareItem(ComparePane::Right, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _fileOperationCompletedCallbackToken = _folderWindow.AddFileOperationCompletedCallback([this](const FolderWindow::FileOperationCompletedEvent& e)
    { OnFolderWindowFileOperationCompleted(e); });

    _optionsPanel.dlg.reset(RedSalamander::Win32Callback::CreateDialogParamResourceNoThrow(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_COMPARE_DIRECTORIES_OPTIONS), hwnd, OptionsDlgProc, reinterpret_cast<LPARAM>(this)));

    if (_optionsPanel.dlg)
    {
        ShowWindow(_optionsPanel.dlg.get(), SW_HIDE);
        LoadOptionsControlsFromSettings();
        ApplyOptionsDialogTheme();
    }
}

LRESULT CALLBACK CompareProgressSpinnerWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetPropW(hwnd, kCompareProgressSpinnerStateProp));
    if (! self)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);

            RECT rc{};
            GetClientRect(hwnd, &rc);
            self->DrawProgressSpinner(hdc.get(), rc);
            return 0;
        }
        case WM_NCDESTROY:
        {
            RemovePropW(hwnd, kCompareProgressSpinnerStateProp);
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void CompareDirectoriesWindow::Layout() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    RECT rc{};
    if (GetClientRect(_hWnd.get(), &rc) == 0)
    {
        return;
    }

    const int w = std::max(0, static_cast<int>(rc.right - rc.left));
    const int h = std::max(0, static_cast<int>(rc.bottom - rc.top));

    _clientSize                = {w, h};
    const int dpi              = static_cast<int>(_dpi);
    const int menuBarHeight    = _chrome.usesMenuBar ? std::clamp(GetDxMenuBarVisibleHeightPx(), 0, h) : 0;
    const int bannerBaseHeight = std::clamp(MulDiv(42, dpi, USER_DEFAULT_SCREEN_DPI), 0, std::max(0, h - menuBarHeight));
    const bool showStatus      = (_progress.scanProgressTextHostHwnd && IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) != 0) ||
                                 (_progress.scanProgressText && IsWindowVisible(_progress.scanProgressText.get()) != 0) ||
                                 (_progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0);
    const int statusHeight =
        showStatus ? std::clamp(MulDiv(kScanStatusHeightDip, dpi, USER_DEFAULT_SCREEN_DPI), 0, std::max(0, h - menuBarHeight - bannerBaseHeight)) : 0;
    const int bannerHeight  = bannerBaseHeight + statusHeight;
    const int contentHeight = std::max(0, h - menuBarHeight - bannerHeight);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    {
        Debug::Perf::Scope bannerLayoutPerf(L"compare.ui.banner_layout_us");
        bannerLayoutPerf.SetValue0(static_cast<uint64_t>(w));
        bannerLayoutPerf.SetValue1(static_cast<uint64_t>(h));

        if (_chrome.menuBarHostHwnd)
        {
            SetWindowPos(_chrome.menuBarHostHwnd.get(), nullptr, 0, 0, w, menuBarHeight, flags);
            ShowWindow(_chrome.menuBarHostHwnd.get(), _chrome.usesMenuBar ? SW_SHOWNA : SW_HIDE);
        }

        // Banner layout
        const int bannerPaddingX = std::max(0, MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI));
        const int bannerPaddingY = std::max(0, MulDiv(6, dpi, USER_DEFAULT_SCREEN_DPI));
        const int buttonW        = std::max(1, MulDiv(110, dpi, USER_DEFAULT_SCREEN_DPI));
        const int buttonH        = std::max(1, MulDiv(28, dpi, USER_DEFAULT_SCREEN_DPI));
        const int buttonGap      = std::max(0, MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI));
        const int bannerTop      = menuBarHeight;
        const int buttonY        = std::max(bannerTop, bannerTop + bannerPaddingY + (std::max(0, bannerHeight - (2 * bannerPaddingY) - buttonH) / 2));

        int rightX                    = std::max(0, w - bannerPaddingX);
        const bool useDxRescanButton  = _chrome.usesBannerButtons && _chrome.bannerRescanHostHwnd;
        const bool useDxOptionsButton = _chrome.usesBannerButtons && _chrome.bannerOptionsHostHwnd;
        if (useDxRescanButton || _chrome.bannerRescanButton)
        {
            rightX = std::max(0, rightX - buttonW);
            if (useDxRescanButton)
            {
                if (_chrome.bannerRescanButtonDx)
                {
                    _chrome.bannerRescanButtonDx->SetBounds(D2D1::RectF(0.0f,
                                                                        0.0f,
                                                                        _chrome.bannerRescanHost.PixelsToDip(static_cast<float>(buttonW)),
                                                                        _chrome.bannerRescanHost.PixelsToDip(static_cast<float>(buttonH))));
                }
                SetWindowPos(_chrome.bannerRescanHostHwnd.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
                ShowWindow(_chrome.bannerRescanHostHwnd.get(), SW_SHOWNA);
                _chrome.bannerRescanHost.Invalidate();
                if (_chrome.bannerRescanButton)
                {
                    ShowWindow(_chrome.bannerRescanButton.get(), SW_HIDE);
                }
            }
            else if (_chrome.bannerRescanButton)
            {
                SetWindowPos(_chrome.bannerRescanButton.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
                ShowWindow(_chrome.bannerRescanButton.get(), SW_SHOWNA);
            }
            rightX = std::max(0, rightX - buttonGap);
        }
        else if (_chrome.bannerRescanHostHwnd)
        {
            ShowWindow(_chrome.bannerRescanHostHwnd.get(), SW_HIDE);
        }
        if (_progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0)
        {
            const int spinnerSize = std::clamp(buttonH, 1, std::max(1, bannerHeight));
            rightX                = std::max(0, rightX - spinnerSize);
            SetWindowPos(_progress.scanProgressBar.get(), nullptr, rightX, buttonY, spinnerSize, spinnerSize, flags);
            rightX = std::max(0, rightX - buttonGap);
        }
        if (useDxOptionsButton || _chrome.bannerOptionsButton)
        {
            rightX = std::max(0, rightX - buttonW);
            if (useDxOptionsButton)
            {
                if (_chrome.bannerOptionsButtonDx)
                {
                    _chrome.bannerOptionsButtonDx->SetBounds(D2D1::RectF(0.0f,
                                                                         0.0f,
                                                                         _chrome.bannerOptionsHost.PixelsToDip(static_cast<float>(buttonW)),
                                                                         _chrome.bannerOptionsHost.PixelsToDip(static_cast<float>(buttonH))));
                }
                SetWindowPos(_chrome.bannerOptionsHostHwnd.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
                ShowWindow(_chrome.bannerOptionsHostHwnd.get(), SW_SHOWNA);
                _chrome.bannerOptionsHost.Invalidate();
                if (_chrome.bannerOptionsButton)
                {
                    ShowWindow(_chrome.bannerOptionsButton.get(), SW_HIDE);
                }
            }
            else if (_chrome.bannerOptionsButton)
            {
                SetWindowPos(_chrome.bannerOptionsButton.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
                ShowWindow(_chrome.bannerOptionsButton.get(), SW_SHOWNA);
            }
            rightX = std::max(0, rightX - buttonGap);
        }
        else if (_chrome.bannerOptionsHostHwnd)
        {
            ShowWindow(_chrome.bannerOptionsHostHwnd.get(), SW_HIDE);
        }
        const bool useDxBannerTitle = _chrome.usesBannerText && _chrome.bannerTitleHostHwnd && _chrome.bannerTitleLabel;
        const int titleW            = std::max(0, rightX - bannerPaddingX);
        if (useDxBannerTitle)
        {
            if (_chrome.bannerTitle)
            {
                ShowWindow(_chrome.bannerTitle.get(), SW_HIDE);
            }
            _chrome.bannerTitleLabel->SetBounds(D2D1::RectF(0.0f,
                                                            0.0f,
                                                            _chrome.bannerTitleHost.PixelsToDip(static_cast<float>(titleW)),
                                                            _chrome.bannerTitleHost.PixelsToDip(static_cast<float>(bannerBaseHeight))));
            SetWindowPos(_chrome.bannerTitleHostHwnd.get(), nullptr, bannerPaddingX, bannerTop, titleW, bannerBaseHeight, flags);
            ShowWindow(_chrome.bannerTitleHostHwnd.get(), SW_SHOWNA);
            _chrome.bannerTitleHost.Invalidate();
        }
        else if (_chrome.bannerTitle)
        {
            SetWindowPos(_chrome.bannerTitle.get(), nullptr, bannerPaddingX, bannerTop, titleW, bannerBaseHeight, flags);
            ShowWindow(_chrome.bannerTitle.get(), SW_SHOWNA);
        }
        else if (_chrome.bannerTitleHostHwnd)
        {
            ShowWindow(_chrome.bannerTitleHostHwnd.get(), SW_HIDE);
        }

        if (showStatus)
        {
            const int statusTop = menuBarHeight + bannerBaseHeight;
            const int paddingX  = std::max(0, MulDiv(kScanStatusPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI));
            const int textX     = paddingX;
            const int textW     = std::max(0, rightX - paddingX);

            if (_chrome.usesBannerText && _progress.scanProgressTextHostHwnd && _progress.scanProgressTextLabel &&
                IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) != 0)
            {
                if (_progress.scanProgressText)
                {
                    ShowWindow(_progress.scanProgressText.get(), SW_HIDE);
                }
                _progress.scanProgressTextLabel->SetBounds(D2D1::RectF(0.0f,
                                                                       0.0f,
                                                                       _progress.scanProgressTextHost.PixelsToDip(static_cast<float>(textW)),
                                                                       _progress.scanProgressTextHost.PixelsToDip(static_cast<float>(statusHeight))));
                SetWindowPos(_progress.scanProgressTextHostHwnd.get(), nullptr, textX, statusTop, textW, statusHeight, flags);
                _progress.scanProgressTextHost.Invalidate();
            }
            else if (_progress.scanProgressText)
            {
                SetWindowPos(_progress.scanProgressText.get(), nullptr, textX, statusTop, textW, statusHeight, flags);
            }
        }
    }

    if (const HWND fw = _folderWindow.GetHwnd())
    {
        SetWindowPos(fw, nullptr, 0, menuBarHeight + bannerHeight, w, contentHeight, flags);
    }

    if (_optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != 0)
    {
        const int outerMargin = std::max(0, MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI));
        const int maxDw       = std::max(1, w - 2 * outerMargin);
        const int maxDh       = std::max(1, contentHeight - 2 * outerMargin);
        const int dw          = maxDw;
        const int dh          = maxDh;
        const int x           = std::max(0, (w - dw) / 2);
        const int y           = std::max(menuBarHeight + bannerHeight, menuBarHeight + bannerHeight + (contentHeight - dh) / 2);
        SetWindowPos(_optionsPanel.dlg.get(), nullptr, x, y, dw, dh, SWP_NOZORDER | SWP_NOACTIVATE);
        LayoutOptionsControls();
    }
}

void CompareDirectoriesWindow::EnsureCompareSession() noexcept
{
    if (_session)
    {
        return;
    }

    if (_leftContext.pluginId.empty() || _rightContext.pluginId.empty())
    {
        return;
    }

    if (! _leftBaseFs)
    {
        std::optional<CreatedFileSystemInstance> created = TryCreateFileSystemInstance(_leftContext.pluginId, _leftContext.instanceContext);
        if (! created.has_value())
        {
            return;
        }

        _leftBaseModule    = std::move(created->module);
        _leftBaseFs        = std::move(created->fileSystem);
        _leftPluginShortId = std::move(created->pluginShortId);
    }

    if (! _rightBaseFs)
    {
        std::optional<CreatedFileSystemInstance> created = TryCreateFileSystemInstance(_rightContext.pluginId, _rightContext.instanceContext);
        if (! created.has_value())
        {
            return;
        }

        _rightBaseModule    = std::move(created->module);
        _rightBaseFs        = std::move(created->fileSystem);
        _rightPluginShortId = std::move(created->pluginShortId);
    }

    Common::Settings::CompareDirectoriesSettings settings = GetEffectiveCompareSettings();
    _session = std::make_shared<CompareDirectoriesSession>(_leftBaseFs, _rightBaseFs, _leftContext.rootPluginPath, _rightContext.rootPluginPath, settings);

    _fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, _session);
    _fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, _session);
}

void CompareDirectoriesWindow::StartCompare() noexcept
{
    EnsureCompareSession();
    if (! _session || ! _fsLeft || ! _fsRight)
    {
        return;
    }

    if (_compareStarted)
    {
        ShowOptionsPanel(false);
        return;
    }

    static_cast<void>(
        _folderWindow.SetFileSystemInstanceForPane(FolderWindow::Pane::Left, _fsLeft, _leftContext.pluginId, _leftPluginShortId, _leftContext.instanceContext));
    static_cast<void>(_folderWindow.SetFileSystemInstanceForPane(
        FolderWindow::Pane::Right, _fsRight, _rightContext.pluginId, _rightPluginShortId, _rightContext.instanceContext));

    _folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, true);
    _folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, true);

    _folderWindow.SetDisplayMode(FolderWindow::Pane::Left, _compareDisplayMode);
    _folderWindow.SetDisplayMode(FolderWindow::Pane::Right, _compareDisplayMode);
    _folderWindow.SetSplitRatio(0.5f);

    _compareStarted = true;
    ShowOptionsPanel(false);

    if (const HWND fw = _folderWindow.GetHwnd())
    {
        ShowWindow(fw, SW_SHOW);
    }
    Layout();

    _syncingPaths = true;
    _folderWindow.SetFolderPath(FolderWindow::Pane::Left, _leftContext.rootPluginPath);
    _folderWindow.SetFolderPath(FolderWindow::Pane::Right, _rightContext.rootPluginPath);
    _syncingPaths = false;

    SetFocus(_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left));
}

void CompareDirectoriesWindow::UpdateCompareRootsFromCurrentPanes() noexcept
{
    if (! _compareStarted)
    {
        return;
    }

    if (const auto leftCurrent = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left); leftCurrent.has_value())
    {
        _leftContext.rootPluginPath = leftCurrent.value();
    }
    if (const auto rightCurrent = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right); rightCurrent.has_value())
    {
        _rightContext.rootPluginPath = rightCurrent.value();
    }
}

void CompareDirectoriesWindow::BeginOrRescanCompare() noexcept
{
    const auto prepared = PrepareCompareRun();
    if (! prepared.has_value())
    {
        return;
    }

    ExecutePreparedCompareRun(prepared->runId, prepared->startedBefore);
}

void CompareDirectoriesWindow::ScheduleBeginOrRescanCompare() noexcept
{
    if (! _hWnd)
    {
        BeginOrRescanCompare();
        return;
    }

    if (_deferredCompareStartPhase != DeferredStartPhase::None)
    {
        return;
    }

    _deferredCompareStartPhase = DeferredStartPhase::Scheduled;
    if (PostMessageW(_hWnd.get(), WndMsg::kCompareDirectoriesDeferredStart, 0, 0) == 0)
    {
        _deferredCompareStartPhase = DeferredStartPhase::None;
    }
}

std::optional<CompareDirectoriesWindow::PreparedCompareRun> CompareDirectoriesWindow::PrepareCompareRun() noexcept
{
    PreparedCompareRun prepared{};
    prepared.startedBefore = _compareStarted;

    ++_compareRunId;
    prepared.runId = _compareRunId;

    EnsureCompareSession();
    if (! _session)
    {
        return std::nullopt;
    }

    SetSessionCallbacksForRun(prepared.runId);
    _session->SetBackgroundWorkEnabled(true);

    UpdateCompareRootsFromCurrentPanes();

    _compareActive                      = true;
    _compareRunPending                  = true;
    _selectionMode                      = CompareSelectionMode::Default;
    _progress.compareRunSawScanProgress = false;
    _progress.compareRunResultHr        = S_OK;
    _session->SetCompareEnabled(true);

    if (_settings && _settings->compareDirectories.has_value())
    {
        _session->SetSettings(GetEffectiveCompareSettings());
    }

    _session->SetRoots(_leftContext.rootPluginPath, _rightContext.rootPluginPath);
    _session->StartScan();
    _session->SetPinnedFolders({}, {});

    _progress.banner                        = {};
    _progress.scanStartTickMs               = GetTickCount64();
    _progress.contentEtaLastTickMs          = 0;
    _progress.contentEtaLastCompletedBytes  = 0;
    _progress.contentEtaSmoothedBytesPerSec = 0.0;
    _progress.contentEtaSeconds.reset();

    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kCompareTaskAutoDismissTimerId);
    }
    DismissCompareTaskCard();
    UpdateCompareTaskCard(false);
    UpdateRescanButtonText();
    UpdateProgressControls();

    return prepared;
}

void CompareDirectoriesWindow::ExecutePreparedCompareRun(uint64_t runId, bool startedBefore) noexcept
{
    if (runId == 0 || runId != _compareRunId)
    {
        return;
    }

    StartCompare();

    if (startedBefore)
    {
        _syncingPaths = true;
        _folderWindow.SetFolderPath(FolderWindow::Pane::Left, _leftContext.rootPluginPath);
        _folderWindow.SetFolderPath(FolderWindow::Pane::Right, _rightContext.rootPluginPath);
        _syncingPaths = false;
    }

    RefreshBothPanes();
}

LRESULT CompareDirectoriesWindow::OnDeferredBeginOrRescanCompare(WPARAM wp) noexcept
{
    if (! _hWnd)
    {
        _deferredCompareStartPhase         = DeferredStartPhase::None;
        _deferredCompareStartRunId         = 0;
        _deferredCompareStartStartedBefore = false;
        return 0;
    }

    if (wp == 0)
    {
        if (_deferredCompareStartPhase != DeferredStartPhase::Scheduled)
        {
            return 0;
        }

        const auto prepared = PrepareCompareRun();
        if (! prepared.has_value())
        {
            _deferredCompareStartPhase = DeferredStartPhase::None;
            return 0;
        }

        _deferredCompareStartRunId         = prepared->runId;
        _deferredCompareStartStartedBefore = prepared->startedBefore;
        _deferredCompareStartPhase         = DeferredStartPhase::Prepared;

        if (PostMessageW(_hWnd.get(), WndMsg::kCompareDirectoriesDeferredStart, static_cast<WPARAM>(_deferredCompareStartRunId), 0) == 0)
        {
            _deferredCompareStartPhase         = DeferredStartPhase::None;
            _deferredCompareStartRunId         = 0;
            _deferredCompareStartStartedBefore = false;
        }
        return 0;
    }

    const uint64_t runId = static_cast<uint64_t>(wp);
    if (_deferredCompareStartPhase != DeferredStartPhase::Prepared || runId == 0 || runId != _deferredCompareStartRunId)
    {
        return 0;
    }

    ExecutePreparedCompareRun(runId, _deferredCompareStartStartedBefore);

    _deferredCompareStartPhase         = DeferredStartPhase::None;
    _deferredCompareStartRunId         = 0;
    _deferredCompareStartStartedBefore = false;
    return 0;
}

void CompareDirectoriesWindow::CancelCompareMode() noexcept
{
    if (! _compareActive)
    {
        return;
    }

    _leaveScopePromptRequest.reset();
    _leaveScopePromptPending = false;

    if (_compareRunPending)
    {
        _progress.compareRunResultHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        UpdateCompareTaskCard(true);
        if (_hWnd)
        {
            SetTimer(_hWnd.get(), kCompareTaskAutoDismissTimerId, kCompareTaskAutoDismissDelayMs, nullptr);
        }
    }

    _compareActive                      = false;
    _compareRunPending                  = false;
    _selectionMode                      = CompareSelectionMode::Default;
    _progress.compareRunSawScanProgress = false;
    UpdateRescanButtonText();

    if (_session)
    {
        LogComparePerfStats(L"cancel", _session, _progress.compareRunResultHr);
        _session->SetBackgroundWorkEnabled(false);
        _session->SetCompareEnabled(false);
        _session->Invalidate();
    }

    _progress.banner.scanActiveScans = 0;
    _progress.banner.scanRelativeFolder.clear();
    _progress.banner.scanEntryName.clear();
    _progress.banner.contentPendingCompares = 0;
    _progress.banner.contentRelativeFolder.clear();
    _progress.banner.contentEntryName.clear();
    _progress.banner.contentFileTotalBytes     = 0;
    _progress.banner.contentFileCompletedBytes = 0;
    for (auto& slot : _progress.banner.contentInFlight)
    {
        slot = {};
    }
    _progress.scanStartTickMs               = 0;
    _progress.contentEtaLastTickMs          = 0;
    _progress.contentEtaLastCompletedBytes  = 0;
    _progress.contentEtaSmoothedBytesPerSec = 0.0;
    _progress.contentEtaSeconds.reset();
    UpdateProgressControls();

    auto clearSelection = [](std::wstring_view) noexcept { return false; };
    _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, clearSelection, true);
    _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, clearSelection, true);
    _folderWindow.SetPaneEmptyStateMessage(FolderWindow::Pane::Left, {});
    _folderWindow.SetPaneEmptyStateMessage(FolderWindow::Pane::Right, {});

    RefreshBothPanes();
}

void CompareDirectoriesWindow::SyncOtherPanePath(ComparePane changedPane,
                                                 const std::optional<std::filesystem::path>& previousPath,
                                                 const std::optional<std::filesystem::path>& newPath) noexcept
{
    if (! _compareStarted || ! _compareActive || _syncingPaths || ! _session || ! newPath.has_value())
    {
        return;
    }

    const auto hasContextChanged = [&](ComparePane pane) noexcept -> bool
    {
        const FolderWindow::Pane fwPane               = (pane == ComparePane::Left) ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
        const CompareDirectoriesPaneContext& expected = (pane == ComparePane::Left) ? _leftContext : _rightContext;

        const std::wstring_view currentPluginId = _folderWindow.GetFileSystemPluginId(fwPane);
        if (! OrdinalString::EqualsNoCase(currentPluginId, expected.pluginId))
        {
            return true;
        }

        const std::optional<std::filesystem::path> currentPath = _folderWindow.GetCurrentPath(fwPane);
        if (! currentPath.has_value())
        {
            return ! expected.instanceContext.empty();
        }

        NavigationLocation::Location loc{};
        if (! NavigationLocation::TryParseLocation(currentPath->native(), loc))
        {
            return true;
        }

        return loc.instanceContext != expected.instanceContext;
    };

    if (hasContextChanged(ComparePane::Left) || hasContextChanged(ComparePane::Right))
    {
        // Compare session scope is defined by (pluginId, instanceContext, roots). If the pane context changes,
        // cancel compare mode to avoid continuing with the wrong filesystem instance.
        CancelCompareMode();
        return;
    }

    const auto relOpt = _session->TryMakeRelative(changedPane, newPath.value());
    if (! relOpt.has_value())
    {
        // User navigated outside the compare scope: cancel compare mode and allow independent browsing.
        if (_compareRunPending && _hWnd)
        {
            const std::filesystem::path rootPath = _session->ResolveAbsolute(changedPane, {});
            std::filesystem::path revertPath     = previousPath.value_or(rootPath);
            if (! _session->TryMakeRelative(changedPane, revertPath).has_value())
            {
                revertPath = rootPath;
            }

            SetComparePanePathSilently(changedPane, revertPath);
            QueueLeaveScopePrompt(changedPane, newPath.value(), revertPath);
            return;
        }

        CancelCompareMode();
        return;
    }

    const ComparePane other              = changedPane == ComparePane::Left ? ComparePane::Right : ComparePane::Left;
    const std::filesystem::path otherAbs = _session->ResolveAbsolute(other, relOpt.value());

    _syncingPaths = true;
    if (other == ComparePane::Left)
    {
        _folderWindow.SetFolderPath(FolderWindow::Pane::Left, otherAbs);
    }
    else
    {
        _folderWindow.SetFolderPath(FolderWindow::Pane::Right, otherAbs);
    }
    _syncingPaths = false;

    _session->SetPinnedFolders(relOpt.value(), relOpt.value());
}

void CompareDirectoriesWindow::SetComparePanePathSilently(ComparePane pane, const std::filesystem::path& path) noexcept
{
    _syncingPaths                = true;
    const auto resetSyncingPaths = wil::scope_exit([&]() noexcept { _syncingPaths = false; });
    _folderWindow.SetFolderPath(ToFolderWindowPane(pane), path);
}

void CompareDirectoriesWindow::QueueLeaveScopePrompt(ComparePane changedPane,
                                                     const std::filesystem::path& attemptedPath,
                                                     const std::filesystem::path& revertPath) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    _leaveScopePromptRequest = LeaveScopePromptRequest{
        .changedPane   = changedPane,
        .runId         = _compareRunId,
        .attemptedPath = attemptedPath,
        .revertPath    = revertPath,
    };

    if (_leaveScopePromptPending)
    {
        return;
    }

    _leaveScopePromptPending = true;
    if (PostMessageW(_hWnd.get(), WndMsg::kCompareDirectoriesLeaveScopePrompt, static_cast<WPARAM>(_compareRunId), 0) == 0)
    {
        _leaveScopePromptPending = false;
        _leaveScopePromptRequest.reset();
        Debug::ErrorWithLastError(L"CompareDirectories: failed to post leave-scope prompt.");
        _progress.compareRunResultHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        CancelCompareMode();
    }
}

LRESULT CompareDirectoriesWindow::OnLeaveScopePrompt() noexcept
{
    const auto request = _leaveScopePromptRequest;
    _leaveScopePromptRequest.reset();
    const auto clearPending = wil::scope_exit([&]() noexcept
    {
        _leaveScopePromptPending = false;
        if (_leaveScopePromptRequest.has_value() && _hWnd)
        {
            _leaveScopePromptPending = true;
            if (PostMessageW(_hWnd.get(), WndMsg::kCompareDirectoriesLeaveScopePrompt, static_cast<WPARAM>(_leaveScopePromptRequest->runId), 0) == 0)
            {
                _leaveScopePromptPending = false;
                _leaveScopePromptRequest.reset();
                Debug::ErrorWithLastError(L"CompareDirectories: failed to repost leave-scope prompt.");
                _progress.compareRunResultHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                CancelCompareMode();
            }
        }
    });
    if (! request.has_value() || ! _hWnd || ! _session || ! _compareActive || request->runId != _compareRunId)
    {
        return 0;
    }

    const std::optional<std::filesystem::path> currentPath = _folderWindow.GetCurrentPluginPath(ToFolderWindowPane(request->changedPane));
    if (! currentPath.has_value() || ! OrdinalString::EqualsNoCasePath(currentPath.value(), request->revertPath))
    {
        return 0;
    }

    const int result = MessageBoxCentered(
        _hWnd.get(), GetModuleHandleW(nullptr), IDS_COMPARE_LEAVE_SCOPE_MESSAGE, IDS_COMPARE_LEAVE_SCOPE_TITLE, MB_OKCANCEL | MB_ICONWARNING);
    if (result != IDOK || ! _compareActive || request->runId != _compareRunId)
    {
        return 0;
    }

    _progress.compareRunResultHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    CancelCompareMode();
    SetComparePanePathSilently(request->changedPane, request->attemptedPath);
    return 0;
}

void CompareDirectoriesWindow::ApplySelectionForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept
{
    if (! _compareStarted || ! _compareActive || ! _session)
    {
        return;
    }

    const auto relOpt = _session->TryMakeRelative(pane, folder);
    if (! relOpt.has_value())
    {
        return;
    }

    const auto decision = _session->TryGetCachedDecision(relOpt.value());
    if (! decision || FAILED(decision->hr))
    {
        return;
    }

    const bool isLeft = pane == ComparePane::Left;
    auto shouldSelect = [&](std::wstring_view name) noexcept -> bool
    {
        const auto it = decision->items.find(name);
        if (it == decision->items.end())
        {
            return false;
        }

        const bool defaultSelected = isLeft ? it->second.selectLeft : it->second.selectRight;
        return _selectionMode == CompareSelectionMode::Inverted ? ! defaultSelected : defaultSelected;
    };

    if (isLeft)
    {
        _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, shouldSelect, true);
    }
    else
    {
        _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, shouldSelect, true);
    }
}

void CompareDirectoriesWindow::UpdateEmptyStateForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept
{
    if (! _compareStarted)
    {
        return;
    }

    const FolderWindow::Pane fwPane = pane == ComparePane::Left ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;

    if (! _compareActive || ! _session)
    {
        _folderWindow.SetPaneEmptyStateMessage(fwPane, {});
        return;
    }

    const auto relOpt = _session->TryMakeRelative(pane, folder);
    if (! relOpt.has_value())
    {
        _folderWindow.SetPaneEmptyStateMessage(fwPane, {});
        return;
    }

    const auto decision = _session->TryGetCachedDecision(relOpt.value());
    if (! decision || FAILED(decision->hr))
    {
        _folderWindow.SetPaneEmptyStateMessage(fwPane, {});
        return;
    }

    const bool missing = pane == ComparePane::Left ? decision->leftFolderMissing : decision->rightFolderMissing;
    if (missing)
    {
        _folderWindow.SetPaneEmptyStateMessage(fwPane, LoadStringResource(nullptr, IDS_COMPARE_FOLDER_NOT_FOUND));
        return;
    }

    const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
    const bool runBusy = _compareRunPending || _progress.banner.scanActiveScans > 0u || _progress.banner.contentPendingCompares > 0u;
    if (! s.showIdenticalItems && ! runBusy && _progress.compareRunResultHr == S_OK && ! decision->anyDifferent && ! decision->anyPending &&
        decision->pendingContentCompareCount == 0)
    {
        _folderWindow.SetPaneEmptyStateMessage(fwPane, LoadStringResource(nullptr, IDS_COMPARE_NO_DIFFERENCES));
        return;
    }

    _folderWindow.SetPaneEmptyStateMessage(fwPane, {});
}

std::wstring CompareDirectoriesWindow::BuildDetailsTextForCompareItem(ComparePane pane,
                                                                      const std::filesystem::path& folder,
                                                                      std::wstring_view displayName,
                                                                      bool isDirectory,
                                                                      uint64_t sizeBytes,
                                                                      int64_t lastWriteTime,
                                                                      DWORD fileAttributes) noexcept
{
    if (! _compareStarted)
    {
        return {};
    }

    if (_compareDisplayMode == FolderView::DisplayMode::Brief)
    {
        return {};
    }

    const std::wstring metaText = BuildMetadataDetailsText(isDirectory, sizeBytes, lastWriteTime, fileAttributes);

    if (! _compareActive || ! _session)
    {
        return metaText;
    }

    DetailsDecisionCache& cache     = pane == ComparePane::Left ? _detailsCacheLeft : _detailsCacheRight;
    const uint64_t currentUiVersion = _session->GetUiVersion();

    if (cache.sessionUiVersion != currentUiVersion || cache.folder != folder)
    {
        cache.sessionUiVersion = currentUiVersion;
        cache.folder           = folder;
        cache.decision.reset();

        if (const auto relOpt = _session->TryMakeRelative(pane, folder))
        {
            cache.decision = _session->TryGetCachedDecision(relOpt.value());
        }
    }

    const auto decision = cache.decision;
    if (! decision || FAILED(decision->hr))
    {
        return metaText;
    }

    const auto it = decision->items.find(displayName);
    if (it == decision->items.end())
    {
        return metaText;
    }

    const CompareDirectoriesItemDecision& item = it->second;
    const uint32_t diffMask                    = item.differenceMask;
    const auto& strings                        = GetCompareDetailsTextStrings();

    std::wstring statusText;

    if (diffMask == 0u)
    {
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::OnlyInLeft))
    {
        statusText.assign(strings.onlyInLeft);
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::OnlyInRight))
    {
        statusText.assign(strings.onlyInRight);
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::TypeMismatch))
    {
        statusText.assign(strings.typeMismatch);
    }
    if (statusText.empty() && diffMask != 0u)
    {
        statusText.reserve(64);

        auto appendToken = [&](std::wstring_view token) noexcept
        {
            if (token.empty())
            {
                return;
            }

            if (! statusText.empty())
            {
                statusText.append(L" • ");
            }
            statusText.append(token);
        };

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Size))
        {
            const bool thisBigger = pane == ComparePane::Left ? (item.leftSizeBytes > item.rightSizeBytes) : (item.rightSizeBytes > item.leftSizeBytes);
            appendToken(thisBigger ? strings.bigger : strings.smaller);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::DateTime))
        {
            const bool thisNewer =
                pane == ComparePane::Left ? (item.leftLastWriteTime > item.rightLastWriteTime) : (item.rightLastWriteTime > item.leftLastWriteTime);
            appendToken(thisNewer ? strings.newer : strings.older);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Attributes))
        {
            appendToken(strings.attributesDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Content))
        {
            appendToken(strings.contentDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::ContentPending))
        {
            appendToken(strings.contentComparing);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirAttributes))
        {
            appendToken(strings.subdirAttributesDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirContent))
        {
            appendToken(strings.subdirContentDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirPending))
        {
            appendToken(strings.subdirComputing);
        }
    }

    if (_compareDisplayMode == FolderView::DisplayMode::ExtraDetailed)
    {
        return statusText;
    }

    return statusText.empty() ? metaText : statusText;
}

std::wstring CompareDirectoriesWindow::BuildMetadataTextForCompareItem(ComparePane pane,
                                                                       const std::filesystem::path& folder,
                                                                       std::wstring_view displayName,
                                                                       bool isDirectory,
                                                                       uint64_t sizeBytes,
                                                                       int64_t lastWriteTime,
                                                                       DWORD fileAttributes) noexcept
{
    UNREFERENCED_PARAMETER(pane);
    UNREFERENCED_PARAMETER(folder);
    UNREFERENCED_PARAMETER(displayName);

    if (! _compareStarted || ! _compareActive || _compareDisplayMode != FolderView::DisplayMode::ExtraDetailed)
    {
        return {};
    }

    return BuildMetadataDetailsText(isDirectory, sizeBytes, lastWriteTime, fileAttributes);
}

void CompareDirectoriesWindow::RefreshBothPanes() noexcept
{
    if (! _compareStarted)
    {
        return;
    }

    const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
    _folderWindow.SetActivePane(pane);
    _folderWindow.CommandRefresh(FolderWindow::Pane::Left);
    _folderWindow.CommandRefresh(FolderWindow::Pane::Right);
}

bool CompareDirectoriesWindow::StartCompareSyncToOtherPane(FolderWindow::Pane sourcePane, FileSystemOperation operation) noexcept
{
    if (! _compareStarted || ! _compareActive || ! _session || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return false;
    }

    const ComparePane compareSourcePane                = sourcePane == FolderWindow::Pane::Left ? ComparePane::Left : ComparePane::Right;
    const CompareDirectoriesPaneContext& sourceContext = compareSourcePane == ComparePane::Left ? _leftContext : _rightContext;

    const std::vector<std::filesystem::path> selectedPaths = _folderWindow.GetSelectedOrFocusedPaths(sourcePane);
    if (selectedPaths.empty())
    {
        return true;
    }

    std::vector<std::filesystem::path> selectedRelativePaths;
    selectedRelativePaths.reserve(selectedPaths.size());
    for (const std::filesystem::path& selectedPath : selectedPaths)
    {
        const auto relativePath = _session->TryMakeRelative(compareSourcePane, selectedPath);
        if (! relativePath.has_value())
        {
            Debug::Warning(L"CompareDirectories: sync ignored out-of-scope selected path '{}'.", selectedPath.native());
            return true;
        }
        selectedRelativePaths.push_back(relativePath.value());
    }

    CompareSyncManifest manifest{};
    CompareSyncManifestBlocker blocker{};
    const CompareSyncManifestStatus status = _session->TryBuildSyncManifest(compareSourcePane, selectedRelativePaths, manifest, blocker);
    if (status == CompareSyncManifestStatus::Empty)
    {
        return true;
    }

    if (status == CompareSyncManifestStatus::NotReady)
    {
        _session->RequestScanForFolder(blocker.relativePath);
        const std::filesystem::path parent = blocker.relativePath.parent_path().lexically_normal();
        if (parent != blocker.relativePath)
        {
            _session->RequestScanForFolder(parent);
        }
        Debug::Perf::Emit(L"compare.sync.manifest.blocker_scheduled", L"not-ready", 0, static_cast<uint64_t>(blocker.reason), 0, blocker.hr);
        std::wstring blockerText = blocker.relativePath.empty() ? std::wstring(L".") : blocker.relativePath.native();
        for (wchar_t& ch : blockerText)
        {
            if (ch == L'/')
            {
                ch = L'\\';
            }
        }
        _folderWindow.ShowPaneAlertOverlay(sourcePane,
                                           FolderView::ErrorOverlayKind::Operation,
                                           FolderView::OverlaySeverity::Information,
                                           LoadStringResource(nullptr, IDS_COMPARE_SYNC_NOT_READY_TITLE),
                                           FormatStringResource(nullptr, IDS_FMT_COMPARE_SYNC_NOT_READY, blockerText),
                                           S_OK,
                                           true,
                                           false);
        ScheduleDecisionRefresh();
        UpdateCompareWatermark();
        Debug::Warning(L"CompareDirectories: sync deferred because compare decision is not ready for '{}'.", blocker.relativePath.native());
        return true;
    }

    if (status != CompareSyncManifestStatus::Ready)
    {
        Debug::Warning(L"CompareDirectories: sync manifest unavailable status={} path='{}' hr=0x{:08X}.",
                       static_cast<unsigned int>(status),
                       blocker.relativePath.native(),
                       static_cast<unsigned long>(blocker.hr));
        return true;
    }

    std::vector<FolderWindow::ResolvedFileOperationItem> items;
    items.reserve(manifest.items.size());
    for (const CompareSyncManifestItem& manifestItem : manifest.items)
    {
        FolderWindow::ResolvedFileOperationItem item{};
        item.sourcePath      = manifestItem.sourceAbsolutePath;
        item.destinationPath = manifestItem.destinationAbsolutePath;
        item.flags           = manifestItem.flags;
        item.kind            = manifestItem.kind == CompareSyncManifestItemKind::DirectoryShell ? FolderWindow::ResolvedFileOperationItemKind::DirectoryShell
                                                                                                : FolderWindow::ResolvedFileOperationItemKind::Normal;
        items.push_back(std::move(item));
    }

    std::optional<std::filesystem::path> destinationFolder;
    std::wstring confirmationMessage;
    if (operation == FILESYSTEM_MOVE)
    {
        uint64_t wholeSubtreeCount = 0;
        for (const CompareSyncManifestItem& manifestItem : manifest.items)
        {
            if (manifestItem.kind == CompareSyncManifestItemKind::WholeSubtree)
            {
                ++wholeSubtreeCount;
            }
        }
        confirmationMessage = FormatStringResource(nullptr, IDS_FMT_COMPARE_SYNC_MOVE_CONFIRM, static_cast<uint64_t>(manifest.items.size()), wholeSubtreeCount);
    }

    const HRESULT hr = _folderWindow.StartFileOperationForResolvedItemsToOtherPane(
        sourceContext.pluginId, sourceContext.instanceContext, operation, std::move(items), &destinationFolder, nullptr, std::move(confirmationMessage));
    if (FAILED(hr))
    {
        Debug::Error(
            L"CompareDirectories: sync operation failed to start op={} hr=0x{:08X}.", static_cast<unsigned int>(operation), static_cast<unsigned long>(hr));
    }

    return true;
}

void CompareDirectoriesWindow::OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept
{
    if (! _compareStarted || ! _compareActive || ! _session)
    {
        return;
    }

    // Invalidate affected paths so the forced refresh performed by FolderWindow updates the compare decisions.
    const bool hasResolvedDestinations = e.destinationPaths.size() == e.sourcePaths.size();
    for (size_t index = 0; index < e.sourcePaths.size(); ++index)
    {
        const std::filesystem::path& src = e.sourcePaths[index];
        _session->InvalidateForAbsolutePath(src, true);

        if (hasResolvedDestinations)
        {
            _session->InvalidateForAbsolutePath(e.destinationPaths[index], true);
        }
        else if (e.destinationFolder.has_value())
        {
            const std::filesystem::path dst = e.destinationFolder.value() / src.filename();
            _session->InvalidateForAbsolutePath(dst, true);
        }
    }
}

LRESULT CompareDirectoriesWindow::OnExecuteShortcutCommand(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<std::wstring>(lp);
    if (! payload || payload->empty())
    {
        return 0;
    }

    ExecuteShortcutCommand(*payload);
    return 0;
}

} // namespace CompareDirectoriesWindowInternal

bool ShowCompareDirectoriesWindow(HWND owner,
                                  FolderWindow& applicationFolderWindow,
                                  Common::Settings::Settings& settings,
                                  const AppTheme& theme,
                                  const ShortcutManager* shortcuts,
                                  CompareDirectoriesPaneContext left,
                                  CompareDirectoriesPaneContext right) noexcept
{
    SessionState::UpdateActiveFileSystemPluginIdsAndOperation({left.pluginId, right.pluginId}, SessionState::OperationKind::Compare);

    auto window = std::make_unique<CompareDirectoriesWindowInternal::CompareDirectoriesWindow>(
        applicationFolderWindow, settings, theme, shortcuts, std::move(left), std::move(right));
    if (! window->Create(owner))
    {
        return false;
    }

    static_cast<void>(window.release());
    return true;
}

HWND GetCompareDirectoriesWindowHandle() noexcept
{
    HWND fallback = nullptr;
    for (auto it = CompareDirectoriesWindowInternal::g_compareDirectoriesWindows.rbegin();
         it != CompareDirectoriesWindowInternal::g_compareDirectoriesWindows.rend();
         ++it)
    {
        const HWND hwnd = *it;
        if (hwnd && IsWindow(hwnd) != FALSE)
        {
            if (IsWindowVisible(hwnd) != FALSE)
            {
                return hwnd;
            }

            if (! fallback)
            {
                fallback = hwnd;
            }
        }
    }

    return fallback;
}

void UpdateCompareDirectoriesWindowsTheme(const AppTheme& theme) noexcept
{
    const auto windows = CompareDirectoriesWindowInternal::g_compareDirectoriesWindows;
    for (HWND hwnd : windows)
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            continue;
        }

        auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (window)
        {
            window->UpdateTheme(theme);
        }
    }
}

#ifdef ENABLE_TESTS
namespace
{

[[nodiscard]] CompareDirectoriesWindowInternal::CompareDirectoriesWindow* ResolveCompareDirectoriesWindowForDebug(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    return reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

} // namespace

bool CompareDirectoriesWindowInternal::CompareDirectoriesWindow::DebugGetRunSnapshot(CompareDirectoriesRunDebugSnapshot& out) const noexcept
{
    out                           = {};
    out.windowVisible             = _hWnd && IsWindowVisible(_hWnd.get()) != FALSE;
    out.optionsDialogVisible      = _optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != FALSE;
    out.compareStarted            = _compareStarted;
    out.compareActive             = _compareActive;
    out.compareRunPending         = _compareRunPending;
    out.compareRunSawScanProgress = _progress.compareRunSawScanProgress;
    out.leaveScopePromptPending   = _leaveScopePromptPending;
    out.usesDxUiMenuBar           = _chrome.usesMenuBar && _chrome.menuBarHostHwnd && IsWindowVisible(_chrome.menuBarHostHwnd.get()) != FALSE;
    out.usesDxUiBannerButtons = _chrome.usesBannerButtons && _chrome.bannerOptionsHostHwnd && _chrome.bannerRescanHostHwnd &&
                                IsWindowVisible(_chrome.bannerOptionsHostHwnd.get()) != FALSE && IsWindowVisible(_chrome.bannerRescanHostHwnd.get()) != FALSE;
    out.usesDxUiBannerText =
        _chrome.usesBannerText && _chrome.bannerTitleHostHwnd && _chrome.bannerTitleLabel && IsWindowVisible(_chrome.bannerTitleHostHwnd.get()) != FALSE;
    out.hasNativeUiFontState      = false;
    out.nativeMenuAttached        = _hWnd && GetMenu(_hWnd.get()) != nullptr;
    out.bannerOptionsEnabled      = _chrome.usesBannerButtons && _chrome.bannerOptionsButtonDx
                                        ? _chrome.bannerOptionsButtonDx->IsEnabled()
                                        : (_chrome.bannerOptionsButton && IsWindowEnabled(_chrome.bannerOptionsButton.get()) != FALSE);
    out.bannerRescanEnabled       = _chrome.usesBannerButtons && _chrome.bannerRescanButtonDx
                                        ? _chrome.bannerRescanButtonDx->IsEnabled()
                                        : (_chrome.bannerRescanButton && IsWindowEnabled(_chrome.bannerRescanButton.get()) != FALSE);
    out.scanActiveScans           = _progress.banner.scanActiveScans;
    out.scanFolderCount           = _progress.banner.scanFolderCount;
    out.scanEntryCount            = _progress.banner.scanEntryCount;
    out.contentPendingCompares    = _progress.banner.contentPendingCompares;
    out.contentCompletedCompares  = _progress.banner.contentCompletedCompares;
    out.contentTotalCompares      = _progress.banner.contentTotalCompares;
    out.dxMenuBarRenderCount      = _chrome.usesMenuBar ? _chrome.menuBarHost.DebugGetRenderCount() : 0u;
    out.menuBarItemCount          = _chrome.menuBar ? _chrome.menuBar->GetItems().size() : 0u;
    out.visibleDxMenuBarHostCount = (_chrome.menuBarHostHwnd && IsWindowVisible(_chrome.menuBarHostHwnd.get()) != FALSE) ? 1u : 0u;
    out.visibleDxBannerButtonHostCount =
        static_cast<size_t>((_chrome.bannerOptionsHostHwnd && IsWindowVisible(_chrome.bannerOptionsHostHwnd.get()) != FALSE) ? 1u : 0u) +
        static_cast<size_t>((_chrome.bannerRescanHostHwnd && IsWindowVisible(_chrome.bannerRescanHostHwnd.get()) != FALSE) ? 1u : 0u);
    out.visibleDxBannerTextHostCount =
        static_cast<size_t>((_chrome.bannerTitleHostHwnd && IsWindowVisible(_chrome.bannerTitleHostHwnd.get()) != FALSE) ? 1u : 0u) +
        static_cast<size_t>((_progress.scanProgressTextHostHwnd && IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) != FALSE) ? 1u : 0u);
    out.visibleLegacyBannerButtonCount =
        static_cast<size_t>((_chrome.bannerOptionsButton && IsWindowVisible(_chrome.bannerOptionsButton.get()) != FALSE) ? 1u : 0u) +
        static_cast<size_t>((_chrome.bannerRescanButton && IsWindowVisible(_chrome.bannerRescanButton.get()) != FALSE) ? 1u : 0u);
    out.visibleLegacyBannerTextCount =
        static_cast<size_t>((_chrome.bannerTitle && IsWindowVisible(_chrome.bannerTitle.get()) != FALSE) ? 1u : 0u) +
        static_cast<size_t>((_progress.scanProgressText && IsWindowVisible(_progress.scanProgressText.get()) != FALSE) ? 1u : 0u);
    out.leftPaneItemCount      = _folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    out.rightPaneItemCount     = _folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    out.leftPaneSelectedCount  = _folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    out.rightPaneSelectedCount = _folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    if (const auto leftPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left); leftPath.has_value())
    {
        out.leftPanePluginPath = leftPath->generic_wstring();
    }
    if (const auto rightPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right); rightPath.has_value())
    {
        out.rightPanePluginPath = rightPath->generic_wstring();
    }
    out.leftPaneEmptyStateMessage  = std::wstring(_folderWindow.DebugGetPaneEmptyStateMessage(FolderWindow::Pane::Left));
    out.rightPaneEmptyStateMessage = std::wstring(_folderWindow.DebugGetPaneEmptyStateMessage(FolderWindow::Pane::Right));
    return true;
}

bool CompareDirectoriesWindowInternal::CompareDirectoriesWindow::DebugSetPanePath(bool leftPane, const std::filesystem::path& path) noexcept
{
    _folderWindow.SetFolderPath(leftPane ? FolderWindow::Pane::Left : FolderWindow::Pane::Right, path);
    return true;
}

bool CompareDirectoriesWindowInternal::CompareDirectoriesWindow::DebugSetRunPending(bool pending) noexcept
{
    if (! _compareStarted || ! _compareActive || ! _session)
    {
        return false;
    }

    _compareRunPending = pending;
    return true;
}

bool CompareDirectoriesWindowInternal::CompareDirectoriesWindow::DebugGetMenuBarItemLabel(size_t index, std::wstring& outText) const noexcept
{
    outText.clear();
    if (! _chrome.menuBar)
    {
        return false;
    }

    const auto items = _chrome.menuBar->GetItems();
    if (index >= items.size())
    {
        return false;
    }

    outText.assign(items[index].text);
    return true;
}

bool CompareDirectoriesWindowInternal::CompareDirectoriesWindow::DebugGetMenuBarItemScreenRect(size_t index, RECT& outRect) const noexcept
{
    return _chrome.menuBar && _chrome.menuBarHostHwnd && IsWindow(_chrome.menuBarHostHwnd.get()) != FALSE &&
           _chrome.menuBar->TryGetItemScreenRect(_chrome.menuBarHost, index, outRect);
}

bool DebugGetCompareDirectoriesOptionsSnapshot(CompareDirectoriesOptionsDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return DebugGetCompareDirectoriesOptionsSnapshotForWindow(hwnd, out);
}

bool DebugGetCompareDirectoriesOptionsSnapshotForWindow(HWND compareWindow, CompareDirectoriesOptionsDebugSnapshot& out) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugGetOptionsSnapshot(out) : false;
}

bool DebugGetCompareDirectoriesRunSnapshot(CompareDirectoriesRunDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return DebugGetCompareDirectoriesRunSnapshotForWindow(hwnd, out);
}

bool DebugGetCompareDirectoriesRunSnapshotForWindow(HWND compareWindow, CompareDirectoriesRunDebugSnapshot& out) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugGetRunSnapshot(out) : false;
}

bool DebugSetCompareDirectoriesPanePathForWindow(HWND compareWindow, bool leftPane, const std::filesystem::path& path) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugSetPanePath(leftPane, path) : false;
}

bool DebugSetCompareDirectoriesRunPendingForWindow(HWND compareWindow, bool pending) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugSetRunPending(pending) : false;
}

void DebugFailNextCompareDirectoriesWindowCreate() noexcept
{
    CompareDirectoriesWindowInternal::g_failNextCompareDirectoriesWindowCreate.store(true, std::memory_order_release);
}

bool DebugGetCompareDirectoriesMenuBarItemLabel(size_t index, std::wstring& outText) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        outText.clear();
        return false;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugGetMenuBarItemLabel(index, outText) : false;
}

bool DebugGetCompareDirectoriesMenuBarItemScreenRect(size_t index, RECT& outRect) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugGetMenuBarItemScreenRect(index, outRect) : false;
}

bool DebugFocusCompareDirectoriesOptionsFirstControl() noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugFocusOptionsFirstControl() : false;
}

bool DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget target) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return DebugFocusCompareDirectoriesOptionsTargetForWindow(hwnd, target);
}

bool DebugFocusCompareDirectoriesOptionsTargetForWindow(HWND compareWindow, CompareDirectoriesOptionsDebugFocusTarget target) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugFocusOptionsTarget(target) : false;
}

bool DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(const bool enabled) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugSetOptionsIgnoreFilesEnabled(enabled) : false;
}

bool DebugSetCompareDirectoriesOptionsIgnoreDirectoriesEnabled(const bool enabled) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugSetOptionsIgnoreDirectoriesEnabled(enabled) : false;
}

HWND DebugGetCompareDirectoriesOptionsDialogHandle() noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    auto* window = reinterpret_cast<CompareDirectoriesWindowInternal::CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugGetOptionsDialogHandle() : nullptr;
}

bool DebugScrollCompareDirectoriesOptionsBodyPages(const int pageDelta) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return DebugScrollCompareDirectoriesOptionsBodyPagesForWindow(hwnd, pageDelta);
}

bool DebugScrollCompareDirectoriesOptionsBodyPagesForWindow(HWND compareWindow, const int pageDelta) noexcept
{
    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugScrollOptionsBodyPages(pageDelta) : false;
}

bool DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(const CompareDirectoriesOptionsDebugFocusTarget target, HWND& outHost, RECT& outRect) noexcept
{
    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(hwnd, target, outHost, outRect);
}

bool DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(HWND compareWindow,
                                                                       const CompareDirectoriesOptionsDebugFocusTarget target,
                                                                       HWND& outHost,
                                                                       RECT& outRect) noexcept
{
    outHost = nullptr;
    outRect = {};

    auto* window = ResolveCompareDirectoriesWindowForDebug(compareWindow);
    return window ? window->DebugGetOptionsTargetHostAndClientRect(target, outHost, outRect) : false;
}
#endif

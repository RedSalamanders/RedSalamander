#include "FolderWindowInternal.h"

#include <algorithm>
#include <array>
#include <cwctype>
#include <limits>
#include <utility>

#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "Helpers.h"
#include "SettingsStore.h"
#include "ViewerPluginManager.h"
#include "resource.h"

namespace
{
constexpr std::wstring_view kFallbackPreviewViewerId = L"builtin/viewer-text";

bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(a, b);
}

// Hard-coded embedded preview-pane allowlist. A new embedded-capable viewer must be added here to appear
// in the preview pane (see Specs/Plugins/Plugins_ViewerPlugins.md "Embedded preview allowlist is hard-coded").
[[nodiscard]] bool SupportsEmbeddedPreviewViewer(std::wstring_view pluginId) noexcept
{
    return EqualsNoCase(pluginId, kFallbackPreviewViewerId) || EqualsNoCase(pluginId, L"builtin/viewer-space") ||
           EqualsNoCase(pluginId, L"builtin/viewer-imgraw") || EqualsNoCase(pluginId, L"builtin/viewer-vlc") || EqualsNoCase(pluginId, L"builtin/viewer-web") ||
           EqualsNoCase(pluginId, L"builtin/viewer-json") || EqualsNoCase(pluginId, L"builtin/viewer-markdown") ||
           EqualsNoCase(pluginId, L"builtin/viewer-pe") || EqualsNoCase(pluginId, L"builtin/viewer-sqlite");
}

[[nodiscard]] bool IsDefaultFileActionResolution(FileActionResolver::Reason reason) noexcept
{
    return reason == FileActionResolver::Reason::ComputerDefaultRule || reason == FileActionResolver::Reason::GlobalDefaultRule;
}

[[nodiscard]] HWND ResolveFileActionOwnerWindow(HWND preferredOwner, HWND fallbackWindow) noexcept
{
    if (preferredOwner && IsWindow(preferredOwner) != FALSE)
    {
        return preferredOwner;
    }

    HWND owner = fallbackWindow ? GetAncestor(fallbackWindow, GA_ROOT) : nullptr;
    if (! owner)
    {
        owner = fallbackWindow;
    }
    return owner;
}

[[nodiscard]] bool HasExplicitPathSyntax(std::wstring_view path) noexcept
{
    return path.find(L'\\') != std::wstring_view::npos || path.find(L'/') != std::wstring_view::npos || path.find(L':') != std::wstring_view::npos;
}

[[nodiscard]] bool ExecutablePathLooksAvailable(std::wstring_view executablePath) noexcept
{
    if (executablePath.empty())
    {
        return false;
    }

    if (executablePath.front() == L'"' || executablePath.back() == L'"')
    {
        return false;
    }

    if (HasExplicitPathSyntax(executablePath))
    {
        std::error_code ec;
        return std::filesystem::exists(std::filesystem::path(executablePath), ec);
    }

    std::array<wchar_t, MAX_PATH> buffer{};
    const std::wstring executable(executablePath);
    const DWORD found = SearchPathW(nullptr, executable.c_str(), nullptr, static_cast<DWORD>(buffer.size()), buffer.data(), nullptr);
    return found > 0u && found < buffer.size();
}

[[nodiscard]] std::wstring FileActionDisplayName(const Common::Settings::FileActionDefinition& action)
{
    if (! action.displayName.empty())
    {
        return action.displayName;
    }
    return action.id;
}

[[nodiscard]] bool UserMenuActionMatchesContext(const Common::Settings::FileActionDefinition& action,
                                                const std::filesystem::path& itemPath,
                                                std::wstring_view computerName) noexcept
{
    return FileActionResolver::ActionAppliesToContext(action, itemPath, computerName);
}

[[nodiscard]] HRESULT CheckUserMenuActionAvailability(const Common::Settings::FileActionDefinition& action,
                                                      const FileActionLauncher::MacroContext& macroContext) noexcept
{
    if (! action.enabled)
    {
        return E_ACCESSDENIED;
    }
    if (action.kind != Common::Settings::FileActionKind::ExternalProgram || action.executablePath.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring expandedExecutable;
    if (const HRESULT hr = FileActionLauncher::ExpandMacros(action.executablePath, macroContext, expandedExecutable); FAILED(hr))
    {
        return hr;
    }
    if (! ExecutablePathLooksAvailable(expandedExecutable))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    return S_OK;
}

[[nodiscard]] std::wstring UserMenuTargetLabel(const std::filesystem::path& itemPath, const std::filesystem::path& currentDirectory)
{
    if (! itemPath.empty())
    {
        const std::wstring fileName = itemPath.filename().wstring();
        if (! fileName.empty())
        {
            return fileName;
        }
        return itemPath.wstring();
    }
    if (! currentDirectory.empty())
    {
        return currentDirectory.wstring();
    }
    return LoadStringResource(nullptr, IDS_CMD_USER_MENU);
}

void ShowUserMenuUnavailableOverlay(
    FolderWindow& window, FolderWindow::Pane pane, std::wstring_view actionId, const std::wstring& targetLabel, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"usermenu.feedback_us");
    perf.SetDetail(L"unavailable");
    perf.SetValue0(static_cast<uint64_t>(actionId.size()));
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_USER_MENU_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_USER_MENU_ITEM_UNAVAILABLE, std::wstring(actionId), targetLabel);
    if (message.empty())
    {
        message = targetLabel;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
}

void ShowUserMenuLaunchFailedOverlay(
    FolderWindow& window, FolderWindow::Pane pane, std::wstring_view actionId, const std::wstring& targetLabel, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"usermenu.feedback_us");
    perf.SetDetail(L"launch-failed");
    perf.SetValue0(static_cast<uint64_t>(actionId.size()));
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_USER_MENU_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_USER_MENU_LAUNCH_FAILED, std::wstring(actionId), targetLabel, static_cast<unsigned long>(hr));
    if (message.empty())
    {
        message = targetLabel;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr, true, false);
}

std::wstring GetComputerNameText() noexcept
{
    wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD computerNameLength = static_cast<DWORD>(std::size(computerName));
    if (GetComputerNameW(computerName, &computerNameLength) == FALSE || computerNameLength == 0u)
    {
        return {};
    }
    return std::wstring(computerName, computerNameLength);
}

uint8_t ClampByte(float value) noexcept
{
    const float scaled  = value * 255.0f;
    const float clamped = std::clamp(scaled, 0.0f, 255.0f);
    return static_cast<uint8_t>(std::lround(clamped));
}

uint32_t ArgbFromColorF(const D2D1::ColorF& color) noexcept
{
    const uint8_t a = ClampByte(color.a);
    const uint8_t r = ClampByte(color.r);
    const uint8_t g = ClampByte(color.g);
    const uint8_t b = ClampByte(color.b);
    return (static_cast<uint32_t>(a) << 24) | (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(g) << 8) | static_cast<uint32_t>(b);
}
} // namespace

HRESULT FolderWindow::ViewerCallbackState::ViewerClosed(void* cookie) noexcept
{
    if (! owner)
    {
        return S_OK;
    }

    ViewerInstance* instance = static_cast<ViewerInstance*>(cookie);
    return owner->OnViewerClosed(instance);
}

void FolderWindow::PersistViewerConfiguration(ViewerInstance& instance) noexcept
{
    if (! _settings || ! instance.viewer || instance.viewerPluginId.empty())
    {
        return;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = instance.viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr) || ! infos)
    {
        return;
    }

    BOOL something            = FALSE;
    const HRESULT saveCheckHr = infos->SomethingToSave(&something);
    if (FAILED(saveCheckHr))
    {
        return;
    }

    std::string currentConfigurationJson;
    bool hasCurrentConfigurationJson = false;
    if (instance.hasInitialConfigurationJson)
    {
        const char* config  = nullptr;
        const HRESULT getHr = infos->GetConfiguration(&config);
        if (SUCCEEDED(getHr))
        {
            currentConfigurationJson    = config ? config : "";
            hasCurrentConfigurationJson = true;
        }
    }

    const bool shouldPersistConfiguration =
        ! instance.hasInitialConfigurationJson || ! hasCurrentConfigurationJson || currentConfigurationJson != instance.initialConfigurationJson;
    if (! shouldPersistConfiguration)
    {
        return;
    }

    if (! something)
    {
        _settings->plugins.configurationByPluginId.erase(instance.viewerPluginId);
        return;
    }

    const char* config = nullptr;
    if (hasCurrentConfigurationJson)
    {
        config = currentConfigurationJson.c_str();
    }
    else
    {
        const HRESULT getHr = infos->GetConfiguration(&config);
        if (FAILED(getHr))
        {
            config = nullptr;
        }
    }

    if (config == nullptr && ! hasCurrentConfigurationJson)
    {
        return;
    }

    Common::Settings::JsonValue persistedValue;
    const std::string_view configText = (config && config[0] != '\0') ? std::string_view(config) : std::string_view("{}");

    const HRESULT parseHr = Common::Settings::ParseJsonValue(configText, persistedValue);
    if (SUCCEEDED(parseHr))
    {
        _settings->plugins.configurationByPluginId[instance.viewerPluginId] = std::move(persistedValue);
    }
    else
    {
        Debug::Warning(L"FolderWindow::PersistViewerConfiguration: failed to parse viewer config JSON for '{}' (hr=0x{:08X}).",
                       instance.viewerPluginId,
                       static_cast<unsigned long>(parseHr));
    }
}

HRESULT FolderWindow::OnViewerClosed(ViewerInstance* instance) noexcept
{
    if (! instance)
    {
        return S_OK;
    }

    for (auto it = _viewerInstances.begin(); it != _viewerInstances.end(); ++it)
    {
        if (it->get() != instance)
        {
            continue;
        }

        PersistViewerConfiguration(*(*it));

        // Do not clear the callback from inside ViewerClosed itself.
        // RegistrationCallbackState::Set(nullptr, ...) drains in-flight
        // callbacks, which would deadlock here because this callback is the
        // in-flight invocation being drained. ShutdownViewers() still clears
        // callbacks before forcing Close() from the host side.
        if (_leftPane.previewViewerInstance == instance)
        {
            _leftPane.previewViewerInstance = nullptr;
            _leftPane.previewViewerPluginId.clear();
        }
        if (_rightPane.previewViewerInstance == instance)
        {
            _rightPane.previewViewerInstance = nullptr;
            _rightPane.previewViewerPluginId.clear();
        }
        _viewerInstances.erase(it);
        break;
    }

    return S_OK;
}

ViewerTheme FolderWindow::BuildViewerTheme() const noexcept
{
    ViewerTheme theme{};
    theme.version                       = 4;
    theme.dpi                           = static_cast<unsigned int>(_dpi);
    theme.backgroundArgb                = ArgbFromColorF(_theme.folderView.backgroundColor);
    theme.textArgb                      = ArgbFromColorF(_theme.folderView.textNormal);
    theme.selectionBackgroundArgb       = ArgbFromColorF(_theme.folderView.itemBackgroundSelected);
    theme.selectionTextArgb             = ArgbFromColorF(_theme.folderView.textSelected);
    theme.accentArgb                    = ArgbFromColorF(_theme.accent);
    theme.alertErrorBackgroundArgb      = ArgbFromColorF(_theme.folderView.errorBackground);
    theme.alertErrorTextArgb            = ArgbFromColorF(_theme.folderView.errorText);
    theme.alertWarningBackgroundArgb    = ArgbFromColorF(_theme.folderView.warningBackground);
    theme.alertWarningTextArgb          = ArgbFromColorF(_theme.folderView.warningText);
    theme.alertInfoBackgroundArgb       = ArgbFromColorF(_theme.folderView.infoBackground);
    theme.alertInfoTextArgb             = ArgbFromColorF(_theme.folderView.infoText);
    theme.darkMode                      = _theme.dark ? TRUE : FALSE;
    theme.highContrast                  = _theme.highContrast ? TRUE : FALSE;
    theme.rainbowMode                   = _theme.menu.rainbowMode ? TRUE : FALSE;
    theme.darkBase                      = _theme.menu.darkBase ? TRUE : FALSE;
    theme.diffAddedBackgroundArgb       = ArgbFromColorF(_theme.viewerDiff.addedBackground);
    theme.diffRemovedBackgroundArgb     = ArgbFromColorF(_theme.viewerDiff.removedBackground);
    theme.diffContextBackgroundArgb     = ArgbFromColorF(_theme.viewerDiff.contextBackground);
    theme.diffHeaderBackgroundArgb      = ArgbFromColorF(_theme.viewerDiff.headerBackground);
    theme.diffBannerBackgroundArgb      = ArgbFromColorF(_theme.viewerDiff.bannerBackground);
    theme.diffPlaceholderBackgroundArgb = ArgbFromColorF(_theme.viewerDiff.placeholderBackground);
    theme.diffDividerArgb               = ArgbFromColorF(_theme.viewerDiff.divider);
    return theme;
}

void FolderWindow::ApplyViewerTheme() noexcept
{
    const ViewerTheme theme = BuildViewerTheme();
    for (const auto& instance : _viewerInstances)
    {
        if (! instance || ! instance->viewer)
        {
            continue;
        }

        static_cast<void>(instance->viewer->SetTheme(&theme));
    }
}

void FolderWindow::ShutdownViewers() noexcept
{
    _leftPane.previewViewerInstance  = nullptr;
    _rightPane.previewViewerInstance = nullptr;
    _leftPane.previewViewerPluginId.clear();
    _rightPane.previewViewerPluginId.clear();

    for (const auto& instance : _viewerInstances)
    {
        if (! instance || ! instance->viewer)
        {
            continue;
        }

        PersistViewerConfiguration(*instance);
        static_cast<void>(instance->viewer->SetCallback(nullptr, nullptr));
        static_cast<void>(instance->viewer->Close());
    }

    _viewerInstances.clear();
}

void FolderWindow::CloseAllViewers() noexcept
{
    ShutdownViewers();
}

HRESULT FolderWindow::OpenViewerWithPlugin(std::wstring_view pluginId, const ViewerOpenContext& context, std::wstring_view openedBy, Pane pane) noexcept
{
    return OpenViewerWithPluginInternal(pluginId, context, openedBy, pane, OpenedFileSourceKind::Viewer, nullptr);
}

void FolderWindow::UpdateViewerInstanceContext(
    ViewerInstance& instance, const ViewerOpenContext& context, std::wstring_view openedBy, Pane pane, OpenedFileSourceKind source) noexcept
{
    instance.openedBy.assign(openedBy);
    instance.source     = source;
    instance.pane       = pane;
    instance.fileSystem = context.fileSystem;
    instance.fileSystemName.assign(context.fileSystemName ? context.fileSystemName : L"");
    instance.focusedPath.assign(context.focusedPath ? context.focusedPath : L"");

    instance.selectionStorage.clear();
    instance.selectionPointers.clear();
    if (context.selectionPaths && context.selectionCount > 0)
    {
        instance.selectionStorage.reserve(context.selectionCount);
        for (unsigned long i = 0; i < context.selectionCount; ++i)
        {
            const wchar_t* path = context.selectionPaths[i];
            if (path && path[0] != L'\0')
            {
                instance.selectionStorage.emplace_back(path);
            }
        }
    }

    instance.selectionPointers.reserve(instance.selectionStorage.size());
    for (const auto& selectionPath : instance.selectionStorage)
    {
        instance.selectionPointers.push_back(selectionPath.c_str());
    }

    instance.otherFilesStorage.clear();
    instance.otherFilePointers.clear();
    if (context.otherFiles && context.otherFileCount > 0)
    {
        instance.otherFilesStorage.reserve(context.otherFileCount);
        for (unsigned long i = 0; i < context.otherFileCount; ++i)
        {
            const wchar_t* path = context.otherFiles[i];
            if (path && path[0] != L'\0')
            {
                instance.otherFilesStorage.emplace_back(path);
            }
        }
    }

    if (instance.otherFilesStorage.empty() && ! instance.focusedPath.empty())
    {
        instance.otherFilesStorage.push_back(instance.focusedPath);
    }

    size_t focusedOtherIndex = static_cast<size_t>(context.focusedOtherFileIndex);
    if (focusedOtherIndex >= instance.otherFilesStorage.size())
    {
        focusedOtherIndex = 0;
        for (size_t i = 0; i < instance.otherFilesStorage.size(); ++i)
        {
            if (EqualsNoCase(instance.otherFilesStorage[i], instance.focusedPath))
            {
                focusedOtherIndex = i;
                break;
            }
        }
    }

    bool focusedPathPresent = false;
    for (const auto& otherPath : instance.otherFilesStorage)
    {
        if (EqualsNoCase(otherPath, instance.focusedPath))
        {
            focusedPathPresent = true;
            break;
        }
    }

    if (! focusedPathPresent && ! instance.focusedPath.empty())
    {
        instance.otherFilesStorage.insert(instance.otherFilesStorage.begin(), instance.focusedPath);
        focusedOtherIndex = 0;
    }

    instance.otherFilePointers.reserve(instance.otherFilesStorage.size());
    for (const auto& otherPath : instance.otherFilesStorage)
    {
        instance.otherFilePointers.push_back(otherPath.c_str());
    }

    HWND ownerWindow = context.ownerWindow;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    }
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    instance.openContext                       = {};
    instance.openContext.ownerWindow           = ownerWindow;
    instance.openContext.fileSystem            = instance.fileSystem.get();
    instance.openContext.fileSystemName        = instance.fileSystemName.empty() ? nullptr : instance.fileSystemName.c_str();
    instance.openContext.focusedPath           = instance.focusedPath.c_str();
    instance.openContext.selectionPaths        = instance.selectionPointers.empty() ? nullptr : instance.selectionPointers.data();
    instance.openContext.selectionCount        = static_cast<unsigned long>(instance.selectionPointers.size());
    instance.openContext.otherFiles            = instance.otherFilePointers.empty() ? nullptr : instance.otherFilePointers.data();
    instance.openContext.otherFileCount        = static_cast<unsigned long>(instance.otherFilePointers.size());
    instance.openContext.focusedOtherFileIndex = static_cast<unsigned long>(focusedOtherIndex);
    instance.openContext.flags                 = context.flags;
}

HRESULT FolderWindow::ReopenViewerInstance(
    ViewerInstance& instance, const ViewerOpenContext& context, std::wstring_view openedBy, Pane pane, OpenedFileSourceKind source) noexcept
{
    if (! instance.viewer || ! context.fileSystem || ! context.focusedPath || context.focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    UpdateViewerInstanceContext(instance, context, openedBy, pane, source);
    const ViewerTheme theme = BuildViewerTheme();
    static_cast<void>(instance.viewer->SetTheme(&theme));
    const HRESULT callbackHr = instance.viewer->SetCallback(&_viewerCallback, &instance);
    if (FAILED(callbackHr))
    {
        return callbackHr;
    }
    return instance.viewer->Open(&instance.openContext);
}

HRESULT FolderWindow::OpenViewerWithPluginInternal(std::wstring_view pluginId,
                                                   const ViewerOpenContext& context,
                                                   std::wstring_view openedBy,
                                                   Pane pane,
                                                   OpenedFileSourceKind source,
                                                   ViewerInstance** outInstance) noexcept
{
    if (! _settings)
    {
        return E_UNEXPECTED;
    }
    if (outInstance)
    {
        *outInstance = nullptr;
    }

    if (pluginId.empty() || ! context.fileSystem || ! context.focusedPath || context.focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    ViewerPluginManager& pluginManager = ViewerPluginManager::GetInstance();

    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = pluginManager.CreateViewerInstance(pluginId, *_settings, viewer);
    if (FAILED(createHr) || ! viewer)
    {
        return FAILED(createHr) ? createHr : E_FAIL;
    }

    auto instance            = std::make_unique<ViewerInstance>();
    instance->viewerPluginId = std::wstring(pluginId);
    instance->viewer         = viewer;

    wil::com_ptr<IInformations> infos;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (SUCCEEDED(infoHr) && infos)
    {
        const char* configurationJson = nullptr;
        const HRESULT getHr           = infos->GetConfiguration(&configurationJson);
        if (SUCCEEDED(getHr))
        {
            instance->hasInitialConfigurationJson = true;
            instance->initialConfigurationJson    = configurationJson ? configurationJson : "";
        }
    }

    UpdateViewerInstanceContext(*instance, context, openedBy, pane, source);

    ViewerInstance* cookie = instance.get();
    _viewerInstances.push_back(std::move(instance));

    const ViewerTheme theme = BuildViewerTheme();
    static_cast<void>(viewer->SetTheme(&theme));
    static_cast<void>(viewer->SetCallback(&_viewerCallback, cookie));

    const HRESULT openHr = viewer->Open(&cookie->openContext);
    if (FAILED(openHr))
    {
        static_cast<void>(viewer->SetCallback(nullptr, nullptr));
        static_cast<void>(viewer->Close());

        for (auto instanceIt = _viewerInstances.begin(); instanceIt != _viewerInstances.end(); ++instanceIt)
        {
            if (instanceIt->get() == cookie)
            {
                _viewerInstances.erase(instanceIt);
                break;
            }
        }
    }
    else if (outInstance)
    {
        *outInstance = cookie;
    }

    return openHr;
}

void FolderWindow::ClosePreviewViewer(Pane hostPane) noexcept
{
    PaneState& host                 = hostPane == Pane::Left ? _leftPane : _rightPane;
    ViewerInstance* previewInstance = host.previewViewerInstance;
    host.previewViewerInstance      = nullptr;
    host.previewViewerPluginId.clear();

    if (! previewInstance)
    {
        return;
    }

    for (auto it = _viewerInstances.begin(); it != _viewerInstances.end(); ++it)
    {
        if (it->get() != previewInstance)
        {
            continue;
        }

        if ((*it)->viewer)
        {
            PersistViewerConfiguration(*(*it));
            static_cast<void>((*it)->viewer->SetCallback(nullptr, nullptr));
            static_cast<void>((*it)->viewer->Close());
        }
        _viewerInstances.erase(it);
        break;
    }
}

bool FolderWindow::OpenPreviewFocusedPathWithViewer(Pane sourcePane, Pane hostPane) noexcept
{
    Debug::Perf::Scope previewPerf(L"preview.switch_us");
    previewPerf.SetDetail(hostPane == Pane::Left ? L"host-left" : L"host-right");

    PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& hostState   = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! _settings || ! sourceState.fileSystem || ! hostState.hPreviewContent || hostState.previewedPath.empty())
    {
        previewPerf.SetHr(E_INVALIDARG);
        return false;
    }

    std::wstring pluginIdStorage;
    std::wstring openedBy;
    std::wstring resolutionSource               = L"fallback";
    const std::wstring computerName             = GetComputerNameText();
    FileActionResolver::Reason configuredReason = FileActionResolver::Reason::None;
    if (! _settings->fileActions.viewers.associations.empty())
    {
        FileActionResolver::Request resolveRequest{};
        resolveRequest.command                          = FileActionResolver::Command::View;
        resolveRequest.filePath                         = hostState.previewedPath;
        resolveRequest.computerName                     = computerName;
        const FileActionResolver::Resolution resolution = FileActionResolver::ResolveViewerAction(_settings->fileActions.viewers, resolveRequest);
        configuredReason                                = resolution.reason;
        if (resolution.action && resolution.action->kind == Common::Settings::FileActionKind::ViewerPlugin && ! resolution.action->pluginId.empty())
        {
            pluginIdStorage  = resolution.action->pluginId;
            openedBy         = resolution.action->displayName;
            resolutionSource = resolution.reasonText.empty() ? L"saved viewer association" : resolution.reasonText;
        }
    }

    Common::Settings::ViewerFileActionsSettings builtInViewers;
    bool builtInViewersInitialized   = false;
    auto resolveBuiltInPreviewViewer = [&]() -> FileActionResolver::Resolution
    {
        if (! builtInViewersInitialized)
        {
            builtInViewers            = Common::Settings::DefaultViewerFileActionsSettings();
            builtInViewersInitialized = true;
        }

        FileActionResolver::Request resolveRequest{};
        resolveRequest.command      = FileActionResolver::Command::View;
        resolveRequest.filePath     = hostState.previewedPath;
        resolveRequest.computerName = computerName;
        return FileActionResolver::ResolveViewerAction(builtInViewers, resolveRequest);
    };

    const bool configuredOnlyTextDefault  = EqualsNoCase(pluginIdStorage, kFallbackPreviewViewerId) && IsDefaultFileActionResolution(configuredReason);
    bool builtInPreviewResolutionSelected = false;
    if (pluginIdStorage.empty() || configuredOnlyTextDefault)
    {
        const FileActionResolver::Resolution builtInResolution = resolveBuiltInPreviewViewer();
        const bool builtInOnlyTextDefault = builtInResolution.action && EqualsNoCase(builtInResolution.action->pluginId, kFallbackPreviewViewerId) &&
                                            IsDefaultFileActionResolution(builtInResolution.reason);
        if (builtInResolution.action && builtInResolution.action->kind == Common::Settings::FileActionKind::ViewerPlugin &&
            ! builtInResolution.action->pluginId.empty() && SupportsEmbeddedPreviewViewer(builtInResolution.action->pluginId) && ! builtInOnlyTextDefault)
        {
            pluginIdStorage = builtInResolution.action->pluginId;
            openedBy        = builtInResolution.action->displayName;
            resolutionSource =
                builtInResolution.reasonText.empty() ? L"built-in viewer defaults" : L"built-in viewer defaults: " + builtInResolution.reasonText;
            builtInPreviewResolutionSelected = true;
        }
    }

    if (configuredOnlyTextDefault && ! builtInPreviewResolutionSelected)
    {
        pluginIdStorage.clear();
        openedBy.clear();
        resolutionSource = L"properties fallback: only the default text viewer matched";
    }

    if (pluginIdStorage.empty())
    {
        Debug::Info(L"FolderWindow::OpenPreviewFocusedPathWithViewer: preview '{}' has no specific embedded viewer match; using item properties fallback.",
                    hostState.previewedPath.wstring());
        previewPerf.SetDetail(L"properties-fallback:no-viewer");
        return false;
    }
    else if (! SupportsEmbeddedPreviewViewer(pluginIdStorage))
    {
        Debug::Warning(
            L"FolderWindow::OpenPreviewFocusedPathWithViewer: preview viewer '{}' for '{}' does not support embedded hosting; using item properties fallback.",
            pluginIdStorage,
            hostState.previewedPath.wstring());
        previewPerf.SetDetail(L"properties-fallback:not-embedded");
        return false;
    }

    if (! OrdinalString::EqualsNoCase(hostState.previewViewerPluginId, pluginIdStorage))
    {
        Debug::Info(L"FolderWindow::OpenPreviewFocusedPathWithViewer: preview '{}' resolved to viewer '{}' ({})",
                    hostState.previewedPath.wstring(),
                    pluginIdStorage,
                    resolutionSource);
    }
    previewPerf.SetDetail(pluginIdStorage);

    if (hostState.previewViewerInstance && OrdinalString::EqualsNoCase(hostState.previewViewerPluginId, pluginIdStorage) &&
        OrdinalString::EqualsNoCase(hostState.previewedPath.wstring(), hostState.previewViewerInstance->focusedPath))
    {
        previewPerf.SetValue0(2u);
        SetPreviewPlaceholder(hostPane, {});
        LayoutEmbeddedPreviewViewer(hostPane);
        FocusPaneFolderView(sourcePane);
        return true;
    }

    std::wstring fileSystemName;
    wil::com_ptr<IInformations> fileSystemInfo;
    if (SUCCEEDED(sourceState.fileSystem->QueryInterface(__uuidof(IInformations), fileSystemInfo.put_void())) && fileSystemInfo)
    {
        const PluginMetaData* meta = nullptr;
        if (SUCCEEDED(fileSystemInfo->GetMetaData(&meta)) && meta && meta->name && meta->name[0] != L'\0')
        {
            fileSystemName = meta->name;
        }
    }
    if (fileSystemName.empty())
    {
        if (! sourceState.pluginShortId.empty())
        {
            fileSystemName = sourceState.pluginShortId;
        }
        else if (! sourceState.pluginId.empty())
        {
            fileSystemName = sourceState.pluginId;
        }
    }

    std::vector<std::filesystem::path> selectedPaths = sourceState.folderView.GetSelectedPaths();
    std::vector<std::wstring> selectionStorage;
    selectionStorage.reserve(selectedPaths.size());
    for (const auto& path : selectedPaths)
    {
        selectionStorage.push_back(path.wstring());
    }

    std::vector<const wchar_t*> selectionPointers;
    selectionPointers.reserve(selectionStorage.size());
    for (const auto& selection : selectionStorage)
    {
        selectionPointers.push_back(selection.c_str());
    }

    const std::wstring focusedPathText = hostState.previewedPath.wstring();
    const wchar_t* otherFilePointer    = focusedPathText.c_str();

    ViewerOpenContext context{};
    context.ownerWindow           = hostState.hPreviewContent.get();
    context.fileSystem            = sourceState.fileSystem.get();
    context.fileSystemName        = fileSystemName.empty() ? nullptr : fileSystemName.c_str();
    context.focusedPath           = focusedPathText.c_str();
    context.selectionPaths        = selectionPointers.empty() ? nullptr : selectionPointers.data();
    context.selectionCount        = static_cast<unsigned long>(selectionPointers.size());
    context.otherFiles            = &otherFilePointer;
    context.otherFileCount        = 1;
    context.focusedOtherFileIndex = 0;
    context.flags                 = static_cast<ViewerOpenFlags>(VIEWER_OPEN_FLAG_EMBEDDED);

    ViewerInstance* instance = nullptr;
    HRESULT openHr           = S_OK;
    if (hostState.previewViewerInstance && OrdinalString::EqualsNoCase(hostState.previewViewerPluginId, pluginIdStorage))
    {
        previewPerf.SetValue0(1u);
        instance = hostState.previewViewerInstance;
        openHr   = ReopenViewerInstance(*instance, context, openedBy, sourcePane, OpenedFileSourceKind::Preview);
        if (SUCCEEDED(openHr))
        {
            SetPreviewPlaceholder(hostPane, {});
            LayoutEmbeddedPreviewViewer(hostPane);
            FocusPaneFolderView(sourcePane);
            return true;
        }

        Debug::Warning(L"FolderWindow::OpenPreviewFocusedPathWithViewer: preview viewer '{}' failed to refresh '{}' (hr=0x{:08X}); reopening.",
                       pluginIdStorage,
                       hostState.previewedPath.wstring(),
                       static_cast<unsigned long>(openHr));
        previewPerf.SetHr(openHr);
        ClosePreviewViewer(hostPane);
        instance = nullptr;
    }
    else
    {
        ClosePreviewViewer(hostPane);
    }

    openHr = OpenViewerWithPluginInternal(pluginIdStorage, context, openedBy, sourcePane, OpenedFileSourceKind::Preview, &instance);
    if (FAILED(openHr) || ! instance)
    {
        Debug::Warning(
            L"FolderWindow::OpenPreviewFocusedPathWithViewer: failed to open preview viewer '{}' for '{}' (hr=0x{:08X}); using item properties fallback.",
            pluginIdStorage,
            hostState.previewedPath.wstring(),
            static_cast<unsigned long>(openHr));
        previewPerf.SetHr(openHr);
        return false;
    }

    hostState.previewViewerInstance = instance;
    hostState.previewViewerPluginId = pluginIdStorage;
    SetPreviewPlaceholder(hostPane, {});
    LayoutEmbeddedPreviewViewer(hostPane);
    FocusPaneFolderView(sourcePane);
    return true;
}

bool FolderWindow::TryViewFileWithViewer(Pane pane, const FolderView::ViewFileRequest& request) noexcept
{
    ClearFileActionFailure();
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! _settings)
    {
        return false;
    }

    if (! state.fileSystem)
    {
        Debug::Error(L"FolderWindow::TryViewFileWithViewer: file system unavailable");
        return false;
    }

    if (request.focusedPath.empty())
    {
        return false;
    }

    constexpr std::wstring_view kFallbackViewerId = L"builtin/viewer-text";

    std::wstring pluginIdStorage;
    const std::wstring computerName = GetComputerNameText();
    auto launchExternalAction       = [&](const Common::Settings::FileActionDefinition& action) noexcept -> bool
    {
        PaneState& oppositeState = pane == Pane::Left ? _rightPane : _leftPane;

        FileActionLauncher::MacroContext macroContext{};
        macroContext.itemPath         = request.focusedPath;
        macroContext.currentDirectory = request.focusedPath.parent_path();
        macroContext.selectedPaths    = request.selectionPaths;
        if (oppositeState.currentPath.has_value())
        {
            macroContext.oppositePanePath = oppositeState.currentPath.value();
        }
        macroContext.computerName = computerName;

        FileActionLauncher::LaunchPlan plan{};
        const HRESULT buildHr = FileActionLauncher::BuildExternalLaunchPlan(action, macroContext, plan);
        if (FAILED(buildHr))
        {
            RecordFileActionLaunchFailure(true, action.id, request.focusedPath, buildHr);
            Debug::Warning(L"FolderWindow::TryViewFileWithViewer: failed to build external viewer action '{}' for '{}' (hr=0x{:08X}).",
                           action.id,
                           request.focusedPath.wstring(),
                           static_cast<unsigned long>(buildHr));
            return false;
        }

        FileActionLauncher::LaunchOptions options{};
        options.ownerWindow          = ResolveFileActionOwnerWindow(request.ownerWindow, _hWnd.get());
        options.captureProcessHandle = true;

        FileActionLauncher::LaunchResult launchResult{};
        const HRESULT launchHr = FileActionLauncher::LaunchExternalPlan(plan, options, &launchResult);
        if (FAILED(launchHr))
        {
            RecordFileActionLaunchFailure(true, action.id, request.focusedPath, launchHr);
            Debug::Warning(L"FolderWindow::TryViewFileWithViewer: failed to launch external viewer action '{}' for '{}' (hr=0x{:08X}).",
                           action.id,
                           request.focusedPath.wstring(),
                           static_cast<unsigned long>(launchHr));
            return false;
        }

        const std::wstring_view openedBy = action.displayName.empty() ? std::wstring_view(action.id) : std::wstring_view(action.displayName);
        RegisterOpenedExternalFile(
            OpenedFileSourceKind::Viewer, request.focusedPath, openedBy, pane, std::move(launchResult.processHandle), launchResult.processId);
        return true;
    };

    const Common::Settings::FileActionDefinition* configuredAction = nullptr;
    if (! request.actionId.empty())
    {
        configuredAction =
            FileActionResolver::FindApplicableActionById(_settings->fileActions.viewers.actions, request.actionId, request.focusedPath, computerName);
        if (! configuredAction)
        {
            return false;
        }

        if (configuredAction->kind == Common::Settings::FileActionKind::ExternalProgram)
        {
            return launchExternalAction(*configuredAction);
        }
        if (configuredAction->kind != Common::Settings::FileActionKind::ViewerPlugin || configuredAction->pluginId.empty())
        {
            return false;
        }

        pluginIdStorage = configuredAction->pluginId;
    }
    else if (! _settings->fileActions.viewers.associations.empty())
    {
        FileActionResolver::Request resolveRequest{};
        resolveRequest.command =
            request.role == FolderView::ViewFileRole::Alternate ? FileActionResolver::Command::AlternateView : FileActionResolver::Command::View;
        resolveRequest.filePath                         = request.focusedPath;
        resolveRequest.computerName                     = computerName;
        const FileActionResolver::Resolution resolution = FileActionResolver::ResolveViewerAction(_settings->fileActions.viewers, resolveRequest);
        configuredAction                                = resolution.action;
        if (configuredAction)
        {
            if (configuredAction->kind == Common::Settings::FileActionKind::ExternalProgram)
            {
                return launchExternalAction(*configuredAction);
            }
            if (configuredAction->kind != Common::Settings::FileActionKind::ViewerPlugin || configuredAction->pluginId.empty())
            {
                return false;
            }

            pluginIdStorage = configuredAction->pluginId;
        }
        else if (request.role == FolderView::ViewFileRole::Alternate)
        {
            return false;
        }
    }

    if (pluginIdStorage.empty())
    {
        if (request.role == FolderView::ViewFileRole::Alternate)
        {
            return false;
        }

        pluginIdStorage.assign(kFallbackViewerId);
    }

    const std::wstring_view pluginId = pluginIdStorage;

    std::vector<std::filesystem::path> otherFiles;
    otherFiles.reserve(request.displayedFilePaths.size());

    const bool isTextViewer = EqualsNoCase(pluginId, kFallbackViewerId);
    for (const auto& candidate : request.displayedFilePaths)
    {
        if (configuredAction)
        {
            FileActionResolver::Request resolveRequest{};
            resolveRequest.command =
                request.role == FolderView::ViewFileRole::Alternate ? FileActionResolver::Command::AlternateView : FileActionResolver::Command::View;
            resolveRequest.filePath                                  = candidate;
            resolveRequest.computerName                              = computerName;
            const FileActionResolver::Resolution candidateResolution = FileActionResolver::ResolveViewerAction(_settings->fileActions.viewers, resolveRequest);
            if (candidateResolution.action && EqualsNoCase(candidateResolution.action->id, configuredAction->id))
            {
                otherFiles.push_back(candidate);
            }
            continue;
        }

        if (isTextViewer)
        {
            otherFiles.push_back(candidate);
        }
    }

    if (otherFiles.empty())
    {
        otherFiles.push_back(request.focusedPath);
    }

    size_t focusedOtherIndex = static_cast<size_t>(-1);
    for (size_t i = 0; i < otherFiles.size(); ++i)
    {
        if (OrdinalString::EqualsNoCasePath(otherFiles[i], request.focusedPath))
        {
            focusedOtherIndex = i;
            break;
        }
    }

    if (focusedOtherIndex == static_cast<size_t>(-1))
    {
        otherFiles.insert(otherFiles.begin(), request.focusedPath);
        focusedOtherIndex = 0;
    }

    std::wstring fileSystemName;
    wil::com_ptr<IInformations> fileSystemInfo;
    if (SUCCEEDED(state.fileSystem->QueryInterface(__uuidof(IInformations), fileSystemInfo.put_void())) && fileSystemInfo)
    {
        const PluginMetaData* meta = nullptr;
        if (SUCCEEDED(fileSystemInfo->GetMetaData(&meta)) && meta && meta->name && meta->name[0] != L'\0')
        {
            fileSystemName = meta->name;
        }
    }
    if (fileSystemName.empty())
    {
        if (! state.pluginShortId.empty())
        {
            fileSystemName = state.pluginShortId;
        }
        else if (! state.pluginId.empty())
        {
            fileSystemName = state.pluginId;
        }
    }

    std::vector<std::wstring> selectionStorage;
    selectionStorage.reserve(request.selectionPaths.size());
    for (const auto& path : request.selectionPaths)
    {
        selectionStorage.push_back(path.wstring());
    }

    std::vector<const wchar_t*> selectionPointers;
    selectionPointers.reserve(selectionStorage.size());
    for (const auto& s : selectionStorage)
    {
        selectionPointers.push_back(s.c_str());
    }

    std::vector<std::wstring> otherFileStorage;
    otherFileStorage.reserve(otherFiles.size());
    for (const auto& path : otherFiles)
    {
        otherFileStorage.push_back(path.wstring());
    }

    std::vector<const wchar_t*> otherFilePointers;
    otherFilePointers.reserve(otherFileStorage.size());
    for (const auto& s : otherFileStorage)
    {
        otherFilePointers.push_back(s.c_str());
    }

    const HWND ownerWindow = ResolveFileActionOwnerWindow(request.ownerWindow, _hWnd.get());

    ViewerOpenContext context{};
    context.ownerWindow            = ownerWindow;
    context.fileSystem             = state.fileSystem.get();
    context.fileSystemName         = fileSystemName.empty() ? nullptr : fileSystemName.c_str();
    const std::wstring focusedPath = request.focusedPath.wstring();
    context.focusedPath            = focusedPath.c_str();
    context.selectionPaths         = selectionPointers.empty() ? nullptr : selectionPointers.data();
    context.selectionCount         = static_cast<unsigned long>(selectionPointers.size());
    context.otherFiles             = otherFilePointers.empty() ? nullptr : otherFilePointers.data();
    context.otherFileCount         = static_cast<unsigned long>(otherFilePointers.size());
    context.focusedOtherFileIndex  = static_cast<unsigned long>(focusedOtherIndex);
    context.flags                  = VIEWER_OPEN_FLAG_NONE;

    const std::wstring_view openedBy =
        configuredAction && ! configuredAction->displayName.empty() ? std::wstring_view(configuredAction->displayName) : std::wstring_view{};
    HRESULT openHr = OpenViewerWithPlugin(pluginIdStorage, context, openedBy, pane);
    if (FAILED(openHr) && ! EqualsNoCase(pluginIdStorage, kFallbackViewerId))
    {
        pluginIdStorage.assign(kFallbackViewerId);
        openHr = OpenViewerWithPlugin(pluginIdStorage, context, {}, pane);
    }

    if (FAILED(openHr))
    {
        Debug::Error(L"FolderWindow::TryViewFileWithViewer: failed to open viewer '{}' for '{}' (hr=0x{:08X}).",
                     pluginIdStorage,
                     request.focusedPath.wstring(),
                     static_cast<unsigned long>(openHr));
    }

    return SUCCEEDED(openHr);
}

bool FolderWindow::TryEditFileWithEditor(Pane pane,
                                          const std::filesystem::path& filePath,
                                          const std::vector<std::filesystem::path>& selectedPaths,
                                          std::wstring_view actionId,
                                          bool alternate,
                                          HWND ownerWindow) noexcept
{
    ClearFileActionFailure();

    if (! _settings)
    {
        return false;
    }

    if (filePath.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> effectiveSelectedPaths = selectedPaths;
    if (effectiveSelectedPaths.empty())
    {
        effectiveSelectedPaths.push_back(filePath);
    }

    const std::wstring computerName                      = GetComputerNameText();
    const Common::Settings::FileActionDefinition* action = nullptr;
    if (! actionId.empty())
    {
        action = FileActionResolver::FindApplicableActionById(_settings->fileActions.editors.actions, actionId, filePath, computerName);
        if (! action)
        {
            return false;
        }
    }
    else
    {
        FileActionResolver::Request resolveRequest{};
        resolveRequest.command                          = alternate ? FileActionResolver::Command::AlternateEdit : FileActionResolver::Command::Edit;
        resolveRequest.filePath                         = filePath;
        resolveRequest.computerName                     = computerName;
        const FileActionResolver::Resolution resolution = FileActionResolver::ResolveEditorAction(_settings->fileActions.editors, resolveRequest);
        action                                          = resolution.action;
    }

    if (! action || action->kind != Common::Settings::FileActionKind::ExternalProgram)
    {
        return false;
    }

    PaneState& oppositeState = pane == Pane::Left ? _rightPane : _leftPane;

    FileActionLauncher::MacroContext macroContext{};
    macroContext.itemPath         = filePath;
    macroContext.currentDirectory = filePath.parent_path();
    macroContext.selectedPaths    = effectiveSelectedPaths;
    if (oppositeState.currentPath.has_value())
    {
        macroContext.oppositePanePath = oppositeState.currentPath.value();
    }
    macroContext.computerName = computerName;

    FileActionLauncher::LaunchPlan plan{};
    const HRESULT buildHr = FileActionLauncher::BuildExternalLaunchPlan(*action, macroContext, plan);
    if (FAILED(buildHr))
    {
        RecordFileActionLaunchFailure(false, action->id, filePath, buildHr);
        Debug::Warning(L"FolderWindow::TryEditFocusedFileWithEditor: failed to build external editor action '{}' for '{}' (hr=0x{:08X}).",
                       action->id,
                       filePath.wstring(),
                       static_cast<unsigned long>(buildHr));
        return false;
    }

    FileActionLauncher::LaunchOptions options{};
    options.ownerWindow          = ResolveFileActionOwnerWindow(ownerWindow, _hWnd.get());
    options.captureProcessHandle = true;

    FileActionLauncher::LaunchResult launchResult{};
    const HRESULT launchHr = FileActionLauncher::LaunchExternalPlan(plan, options, &launchResult);
    if (FAILED(launchHr))
    {
        RecordFileActionLaunchFailure(false, action->id, filePath, launchHr);
        Debug::Warning(L"FolderWindow::TryEditFocusedFileWithEditor: failed to launch external editor action '{}' for '{}' (hr=0x{:08X}).",
                       action->id,
                       filePath.wstring(),
                       static_cast<unsigned long>(launchHr));
        return false;
    }

    const std::wstring_view openedBy = action->displayName.empty() ? std::wstring_view(action->id) : std::wstring_view(action->displayName);
    RegisterOpenedExternalFile(OpenedFileSourceKind::Editor, filePath, openedBy, pane, std::move(launchResult.processHandle), launchResult.processId);
    return true;
}

bool FolderWindow::TryEditFocusedFileWithEditor(Pane pane, std::wstring_view actionId, bool alternate) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    if (! focusedPath.has_value() || focusedPath.value().empty())
    {
        ClearFileActionFailure();
        return false;
    }

    std::vector<std::filesystem::path> selectedPaths = state.folderView.GetSelectedOrFocusedPaths();
    if (selectedPaths.empty())
    {
        selectedPaths.push_back(focusedPath.value());
    }

    return TryEditFileWithEditor(pane, focusedPath.value(), selectedPaths, actionId, alternate);
}

bool FolderWindow::TryLaunchResolvedFileAction(std::wstring_view sourcePluginId,
                                               std::wstring_view sourceInstanceContext,
                                               const std::filesystem::path& focusedPath,
                                               std::vector<std::filesystem::path> selectedPaths,
                                               std::vector<std::filesystem::path> displayedFilePaths,
                                               unsigned int commandId,
                                               HWND ownerWindow) noexcept
{
    if (focusedPath.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourceProbe;
    sourceProbe.push_back(focusedPath);
    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourceProbe);
    if (! resolvedSourcePane.has_value())
    {
        return false;
    }

    if (selectedPaths.empty())
    {
        selectedPaths.push_back(focusedPath);
    }
    if (displayedFilePaths.empty())
    {
        displayedFilePaths.push_back(focusedPath);
    }

    const Pane pane = resolvedSourcePane.value();
    ClearFileActionFailure();

    switch (commandId)
    {
        case IDM_PANE_VIEW:
        case IDM_PANE_ALTERNATE_VIEW:
        {
            FolderView::ViewFileRequest request{};
            request.role               = commandId == IDM_PANE_ALTERNATE_VIEW ? FolderView::ViewFileRole::Alternate : FolderView::ViewFileRole::Primary;
            request.ownerWindow        = ownerWindow;
            request.focusedPath        = focusedPath;
            request.selectionPaths     = std::move(selectedPaths);
            request.displayedFilePaths = std::move(displayedFilePaths);

            const bool handled = TryViewFileWithViewer(pane, request);
            if (! handled)
            {
                static_cast<void>(ShowRecordedFileActionFailureOverlay(pane));
            }
            return handled;
        }

        case IDM_PANE_EDIT:
        case IDM_PANE_ALTERNATE_EDIT:
        {
            const bool alternate = commandId == IDM_PANE_ALTERNATE_EDIT;
            const bool handled   = TryEditFileWithEditor(pane, focusedPath, selectedPaths, {}, alternate, ownerWindow);
            if (! handled)
            {
                static_cast<void>(ShowRecordedFileActionFailureOverlay(pane));
            }
            return handled;
        }

        default: return false;
    }
}

std::vector<FolderWindow::UserMenuItem> FolderWindow::CollectUserMenuItems(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _settings)
    {
        return {};
    }

    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    std::filesystem::path currentDirectory;
    if (state.currentPath.has_value())
    {
        currentDirectory = state.currentPath.value();
    }
    else if (const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath(); folderPath.has_value())
    {
        currentDirectory = folderPath.value();
    }

    std::filesystem::path itemPath;
    if (focusedPath.has_value())
    {
        itemPath = focusedPath.value();
    }
    if (itemPath.empty())
    {
        itemPath = currentDirectory;
    }

    FileActionLauncher::MacroContext macroContext{};
    macroContext.itemPath         = itemPath;
    macroContext.currentDirectory = currentDirectory;
    if (macroContext.currentDirectory.empty() && ! itemPath.empty())
    {
        macroContext.currentDirectory = itemPath.parent_path();
    }

    const PaneState& oppositeState = pane == Pane::Left ? _rightPane : _leftPane;
    if (oppositeState.currentPath.has_value())
    {
        macroContext.oppositePanePath = oppositeState.currentPath.value();
    }
    macroContext.computerName = GetComputerNameText();

    std::vector<UserMenuItem> items;
    items.reserve(_settings->userMenu.actions.size());
    for (const Common::Settings::FileActionDefinition& action : _settings->userMenu.actions)
    {
        if (action.id.empty() || ! UserMenuActionMatchesContext(action, itemPath, macroContext.computerName))
        {
            continue;
        }

        UserMenuItem item{};
        item.id             = action.id;
        item.displayName    = FileActionDisplayName(action);
        item.availabilityHr = CheckUserMenuActionAvailability(action, macroContext);
        item.enabled        = SUCCEEDED(item.availabilityHr);
        items.push_back(std::move(item));
    }

    return items;
}

void FolderWindow::CommandUserMenu(Pane pane, std::wstring_view actionId)
{
    Debug::Perf::Scope perf(L"usermenu.launch_us");
    perf.SetDetail(actionId);
    perf.SetValue0(static_cast<uint64_t>(actionId.size()));

    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    ClearFileActionFailure();

    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    std::filesystem::path currentDirectory;
    if (state.currentPath.has_value())
    {
        currentDirectory = state.currentPath.value();
    }
    else if (const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath(); folderPath.has_value())
    {
        currentDirectory = folderPath.value();
    }

    std::filesystem::path itemPath;
    if (focusedPath.has_value())
    {
        itemPath = focusedPath.value();
    }
    if (itemPath.empty())
    {
        itemPath = currentDirectory;
    }

    const std::wstring targetLabel = UserMenuTargetLabel(itemPath, currentDirectory);
    if (! _settings || actionId.empty())
    {
        constexpr HRESULT hr = E_INVALIDARG;
        perf.SetHr(hr);
        ShowUserMenuUnavailableOverlay(*this, pane, actionId, targetLabel, hr);
        return;
    }

    const std::wstring computerName = GetComputerNameText();
    const auto actionIt = std::find_if(_settings->userMenu.actions.begin(),
                                       _settings->userMenu.actions.end(),
                                       [&](const Common::Settings::FileActionDefinition& action) noexcept { return EqualsNoCase(action.id, actionId); });

    if (actionIt == _settings->userMenu.actions.end() || ! UserMenuActionMatchesContext(*actionIt, itemPath, computerName))
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        perf.SetHr(hr);
        ShowUserMenuUnavailableOverlay(*this, pane, actionId, targetLabel, hr);
        return;
    }

    std::vector<std::filesystem::path> selectedPaths = state.folderView.GetSelectedOrFocusedPaths();
    if (selectedPaths.empty() && ! itemPath.empty())
    {
        selectedPaths.push_back(itemPath);
    }

    FileActionLauncher::MacroContext macroContext{};
    macroContext.itemPath         = itemPath;
    macroContext.currentDirectory = currentDirectory;
    if (macroContext.currentDirectory.empty() && ! itemPath.empty())
    {
        macroContext.currentDirectory = itemPath.parent_path();
    }
    macroContext.selectedPaths = std::move(selectedPaths);

    PaneState& oppositeState = pane == Pane::Left ? _rightPane : _leftPane;
    if (oppositeState.currentPath.has_value())
    {
        macroContext.oppositePanePath = oppositeState.currentPath.value();
    }
    macroContext.computerName = computerName;

    if (const HRESULT availabilityHr = CheckUserMenuActionAvailability(*actionIt, macroContext); FAILED(availabilityHr))
    {
        perf.SetHr(availabilityHr);
        ShowUserMenuUnavailableOverlay(*this, pane, actionIt->id, targetLabel, availabilityHr);
        return;
    }

    FileActionLauncher::LaunchPlan plan{};
    const HRESULT buildHr = FileActionLauncher::BuildExternalLaunchPlan(*actionIt, macroContext, plan);
    if (FAILED(buildHr))
    {
        perf.SetHr(buildHr);
        ShowUserMenuLaunchFailedOverlay(*this, pane, actionIt->id, targetLabel, buildHr);
        Debug::Warning(L"FolderWindow::CommandUserMenu: failed to build external user-menu action '{}' for '{}' (hr=0x{:08X}).",
                       actionIt->id,
                       itemPath.wstring(),
                       static_cast<unsigned long>(buildHr));
        return;
    }

    FileActionLauncher::LaunchOptions options{};
    options.ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! options.ownerWindow)
    {
        options.ownerWindow = _hWnd.get();
    }

    const HRESULT launchHr = FileActionLauncher::LaunchExternalPlan(plan, options);
    if (FAILED(launchHr))
    {
        perf.SetHr(launchHr);
        ShowUserMenuLaunchFailedOverlay(*this, pane, actionIt->id, targetLabel, launchHr);
        Debug::Warning(L"FolderWindow::CommandUserMenu: failed to launch external user-menu action '{}' for '{}' (hr=0x{:08X}).",
                       actionIt->id,
                       itemPath.wstring(),
                       static_cast<unsigned long>(launchHr));
        return;
    }

    perf.SetValue1(static_cast<uint64_t>(macroContext.selectedPaths.size()));
}

bool FolderWindow::TryViewSpaceWithViewer(Pane pane, const std::filesystem::path& folderPath) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! _settings)
    {
        Debug::Error(L"FolderWindow::TryViewSpaceWithViewer: settings unavailable");
        return false;
    }

    if (folderPath.empty())
    {
        Debug::Error(L"FolderWindow::TryViewSpaceWithViewer: empty folder path");
        return false;
    }

    std::wstring fileSystemName;
    if (state.fileSystem)
    {
        wil::com_ptr<IInformations> fileSystemInfo;
        if (SUCCEEDED(state.fileSystem->QueryInterface(__uuidof(IInformations), fileSystemInfo.put_void())) && fileSystemInfo)
        {
            const PluginMetaData* meta = nullptr;
            if (SUCCEEDED(fileSystemInfo->GetMetaData(&meta)) && meta && meta->name && meta->name[0] != L'\0')
            {
                fileSystemName = meta->name;
            }
        }
    }
    if (fileSystemName.empty())
    {
        if (! state.pluginShortId.empty())
        {
            fileSystemName = state.pluginShortId;
        }
        else if (! state.pluginId.empty())
        {
            fileSystemName = state.pluginId;
        }
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::wstring focusedPath = folderPath.wstring();
    ViewerOpenContext context{};
    context.ownerWindow           = ownerWindow;
    context.fileSystem            = state.fileSystem.get();
    context.fileSystemName        = fileSystemName.empty() ? nullptr : fileSystemName.c_str();
    context.focusedPath           = focusedPath.c_str();
    context.selectionPaths        = nullptr;
    context.selectionCount        = 0;
    context.otherFiles            = nullptr;
    context.otherFileCount        = 0;
    context.focusedOtherFileIndex = 0;
    context.flags                 = VIEWER_OPEN_FLAG_NONE;

    const HRESULT openHr = OpenViewerWithPlugin(L"builtin/viewer-space", context, L"ViewerSpace", pane);
    if (FAILED(openHr))
    {
        Debug::Error(L"FolderWindow::TryViewSpaceWithViewer: failed to open viewer instance (hr=0x{:08X}).", static_cast<unsigned long>(openHr));
        return false;
    }

    return true;
}

#ifdef ENABLE_TESTS
size_t FolderWindow::DebugGetViewerInstanceCount() const noexcept
{
    return _viewerInstances.size();
}

bool FolderWindow::DebugHasViewerPluginId(std::wstring_view viewerPluginId) const noexcept
{
    if (viewerPluginId.empty())
    {
        return false;
    }

    for (const auto& instance : _viewerInstances)
    {
        if (! instance)
        {
            continue;
        }

        if (instance->viewerPluginId == viewerPluginId)
        {
            return true;
        }
    }

    return false;
}
#endif

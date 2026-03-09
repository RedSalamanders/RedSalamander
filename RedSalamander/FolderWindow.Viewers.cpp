#include "FolderWindowInternal.h"

#include <algorithm>
#include <cwctype>
#include <limits>
#include <utility>

#include "Helpers.h"
#include "SettingsStore.h"
#include "ViewerPluginManager.h"

namespace
{
std::wstring ToLowerInvariant(std::wstring_view text)
{
    std::wstring lowered(text);
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });
    return lowered;
}

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

        if (_settings && (*it)->viewer && ! (*it)->viewerPluginId.empty())
        {
            wil::com_ptr<IInformations> infos;
            const HRESULT qiHr = (*it)->viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
            if (SUCCEEDED(qiHr) && infos)
            {
                BOOL something            = FALSE;
                const HRESULT saveCheckHr = infos->SomethingToSave(&something);
                if (SUCCEEDED(saveCheckHr))
                {
                    if (! something)
                    {
                        _settings->plugins.configurationByPluginId.erase((*it)->viewerPluginId);
                    }
                    else
                    {
                        const char* config  = nullptr;
                        const HRESULT getHr = infos->GetConfiguration(&config);
                        if (SUCCEEDED(getHr))
                        {
                            Common::Settings::JsonValue persistedValue;
                            const std::string_view configText = (config && config[0] != '\0') ? std::string_view(config) : std::string_view("{}");

                            const HRESULT parseHr = Common::Settings::ParseJsonValue(configText, persistedValue);
                            if (SUCCEEDED(parseHr))
                            {
                                _settings->plugins.configurationByPluginId[(*it)->viewerPluginId] = std::move(persistedValue);
                            }
                            else
                            {
                                Debug::Warning(L"FolderWindow::OnViewerClosed: failed to parse viewer config JSON for '{}' (hr=0x{:08X}).",
                                               (*it)->viewerPluginId,
                                               static_cast<unsigned long>(parseHr));
                            }
                        }
                    }
                }
            }
        }

        if ((*it)->viewer)
        {
            static_cast<void>((*it)->viewer->SetCallback(nullptr, nullptr));
        }

        _viewerInstances.erase(it);
        break;
    }

    return S_OK;
}

ViewerTheme FolderWindow::BuildViewerTheme() const noexcept
{
    ViewerTheme theme{};
    theme.version                    = 2;
    theme.dpi                        = static_cast<unsigned int>(_dpi);
    theme.backgroundArgb             = ArgbFromColorF(_theme.folderView.backgroundColor);
    theme.textArgb                   = ArgbFromColorF(_theme.folderView.textNormal);
    theme.selectionBackgroundArgb    = ArgbFromColorF(_theme.folderView.itemBackgroundSelected);
    theme.selectionTextArgb          = ArgbFromColorF(_theme.folderView.textSelected);
    theme.accentArgb                 = ArgbFromColorF(_theme.accent);
    theme.alertErrorBackgroundArgb   = ArgbFromColorF(_theme.folderView.errorBackground);
    theme.alertErrorTextArgb         = ArgbFromColorF(_theme.folderView.errorText);
    theme.alertWarningBackgroundArgb = ArgbFromColorF(_theme.folderView.warningBackground);
    theme.alertWarningTextArgb       = ArgbFromColorF(_theme.folderView.warningText);
    theme.alertInfoBackgroundArgb    = ArgbFromColorF(_theme.folderView.infoBackground);
    theme.alertInfoTextArgb          = ArgbFromColorF(_theme.folderView.infoText);
    theme.darkMode                   = _theme.dark ? TRUE : FALSE;
    theme.highContrast               = _theme.highContrast ? TRUE : FALSE;
    theme.rainbowMode                = _theme.menu.rainbowMode ? TRUE : FALSE;
    theme.darkBase                   = _theme.menu.darkBase ? TRUE : FALSE;
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
    for (const auto& instance : _viewerInstances)
    {
        if (! instance || ! instance->viewer)
        {
            continue;
        }

        static_cast<void>(instance->viewer->SetCallback(nullptr, nullptr));
        static_cast<void>(instance->viewer->Close());
    }

    _viewerInstances.clear();
}

void FolderWindow::CloseAllViewers() noexcept
{
    ShutdownViewers();
}

HRESULT FolderWindow::OpenViewerWithPlugin(std::wstring_view pluginId, const ViewerOpenContext& context) noexcept
{
    if (! _settings)
    {
        return E_UNEXPECTED;
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
    instance->fileSystem     = context.fileSystem;
    instance->fileSystemName.assign(context.fileSystemName ? context.fileSystemName : L"");
    instance->focusedPath.assign(context.focusedPath);

    if (context.selectionPaths && context.selectionCount > 0)
    {
        instance->selectionStorage.reserve(context.selectionCount);
        for (unsigned long i = 0; i < context.selectionCount; ++i)
        {
            const wchar_t* path = context.selectionPaths[i];
            if (path && path[0] != L'\0')
            {
                instance->selectionStorage.emplace_back(path);
            }
        }
    }

    instance->selectionPointers.reserve(instance->selectionStorage.size());
    for (const auto& selectionPath : instance->selectionStorage)
    {
        instance->selectionPointers.push_back(selectionPath.c_str());
    }

    if (context.otherFiles && context.otherFileCount > 0)
    {
        instance->otherFilesStorage.reserve(context.otherFileCount);
        for (unsigned long i = 0; i < context.otherFileCount; ++i)
        {
            const wchar_t* path = context.otherFiles[i];
            if (path && path[0] != L'\0')
            {
                instance->otherFilesStorage.emplace_back(path);
            }
        }
    }

    if (instance->otherFilesStorage.empty())
    {
        instance->otherFilesStorage.push_back(instance->focusedPath);
    }

    size_t focusedOtherIndex = static_cast<size_t>(context.focusedOtherFileIndex);
    if (focusedOtherIndex >= instance->otherFilesStorage.size())
    {
        focusedOtherIndex = 0;
        for (size_t i = 0; i < instance->otherFilesStorage.size(); ++i)
        {
            if (EqualsNoCase(instance->otherFilesStorage[i], instance->focusedPath))
            {
                focusedOtherIndex = i;
                break;
            }
        }
    }

    bool focusedPathPresent = false;
    for (const auto& otherPath : instance->otherFilesStorage)
    {
        if (EqualsNoCase(otherPath, instance->focusedPath))
        {
            focusedPathPresent = true;
            break;
        }
    }

    if (! focusedPathPresent)
    {
        instance->otherFilesStorage.insert(instance->otherFilesStorage.begin(), instance->focusedPath);
        focusedOtherIndex = 0;
    }

    instance->otherFilePointers.reserve(instance->otherFilesStorage.size());
    for (const auto& otherPath : instance->otherFilesStorage)
    {
        instance->otherFilePointers.push_back(otherPath.c_str());
    }

    ViewerInstance* cookie = instance.get();
    _viewerInstances.push_back(std::move(instance));

    const ViewerTheme theme = BuildViewerTheme();
    static_cast<void>(viewer->SetTheme(&theme));
    static_cast<void>(viewer->SetCallback(&_viewerCallback, cookie));

    HWND ownerWindow = context.ownerWindow;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    }
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    cookie->openContext                       = {};
    cookie->openContext.ownerWindow           = ownerWindow;
    cookie->openContext.fileSystem            = cookie->fileSystem.get();
    cookie->openContext.fileSystemName        = cookie->fileSystemName.empty() ? nullptr : cookie->fileSystemName.c_str();
    cookie->openContext.focusedPath           = cookie->focusedPath.c_str();
    cookie->openContext.selectionPaths        = cookie->selectionPointers.empty() ? nullptr : cookie->selectionPointers.data();
    cookie->openContext.selectionCount        = static_cast<unsigned long>(cookie->selectionPointers.size());
    cookie->openContext.otherFiles            = cookie->otherFilePointers.empty() ? nullptr : cookie->otherFilePointers.data();
    cookie->openContext.otherFileCount        = static_cast<unsigned long>(cookie->otherFilePointers.size());
    cookie->openContext.focusedOtherFileIndex = static_cast<unsigned long>(focusedOtherIndex);
    cookie->openContext.flags                 = context.flags;

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

    return openHr;
}

bool FolderWindow::TryViewFileWithViewer(Pane pane, const FolderView::ViewFileRequest& request) noexcept
{
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

    std::wstring pluginIdStorage(kFallbackViewerId);

    std::wstring ext = request.focusedPath.extension().wstring();
    if (! ext.empty())
    {
        ext = ToLowerInvariant(ext);
        if (ext.front() != L'.')
        {
            ext.insert(ext.begin(), L'.');
        }

        const auto it = _settings->extensions.openWithViewerByExtension.find(ext);
        if (it != _settings->extensions.openWithViewerByExtension.end())
        {
            if (it->second.empty())
            {
                return false;
            }

            pluginIdStorage = it->second;
        }
    }

    const std::wstring_view pluginId = pluginIdStorage;

    std::vector<std::filesystem::path> otherFiles;
    otherFiles.reserve(request.displayedFilePaths.size());

    const bool isTextViewer = EqualsNoCase(pluginId, kFallbackViewerId);
    for (const auto& candidate : request.displayedFilePaths)
    {
        std::wstring candidateExt = candidate.extension().wstring();
        if (! candidateExt.empty())
        {
            candidateExt = ToLowerInvariant(candidateExt);
            if (candidateExt.front() != L'.')
            {
                candidateExt.insert(candidateExt.begin(), L'.');
            }

            const auto mapIt = _settings->extensions.openWithViewerByExtension.find(candidateExt);
            if (mapIt != _settings->extensions.openWithViewerByExtension.end())
            {
                if (mapIt->second.empty())
                {
                    continue;
                }

                if (isTextViewer)
                {
                    if (! EqualsNoCase(mapIt->second, kFallbackViewerId))
                    {
                        continue;
                    }
                }
                else if (! EqualsNoCase(mapIt->second, pluginId))
                {
                    continue;
                }

                otherFiles.push_back(candidate);
                continue;
            }
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

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    ViewerOpenContext context{};
    context.ownerWindow           = ownerWindow;
    context.fileSystem            = state.fileSystem.get();
    context.fileSystemName        = fileSystemName.empty() ? nullptr : fileSystemName.c_str();
    const std::wstring focusedPath = request.focusedPath.wstring();
    context.focusedPath           = focusedPath.c_str();
    context.selectionPaths        = selectionPointers.empty() ? nullptr : selectionPointers.data();
    context.selectionCount        = static_cast<unsigned long>(selectionPointers.size());
    context.otherFiles            = otherFilePointers.empty() ? nullptr : otherFilePointers.data();
    context.otherFileCount        = static_cast<unsigned long>(otherFilePointers.size());
    context.focusedOtherFileIndex = static_cast<unsigned long>(focusedOtherIndex);
    context.flags                 = VIEWER_OPEN_FLAG_NONE;

    HRESULT openHr = OpenViewerWithPlugin(pluginIdStorage, context);
    if (FAILED(openHr) && ! EqualsNoCase(pluginIdStorage, kFallbackViewerId))
    {
        pluginIdStorage.assign(kFallbackViewerId);
        openHr = OpenViewerWithPlugin(pluginIdStorage, context);
    }

    return SUCCEEDED(openHr);
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

    const HRESULT openHr = OpenViewerWithPlugin(L"builtin/viewer-space", context);
    if (FAILED(openHr))
    {
        Debug::Error(L"FolderWindow::TryViewSpaceWithViewer: failed to open viewer instance (hr=0x{:08X}).", static_cast<unsigned long>(openHr));
        return false;
    }

    return true;
}

#ifdef _DEBUG
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

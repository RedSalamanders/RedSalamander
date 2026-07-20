#include "FolderViewInternal.h"

namespace
{
constexpr HRESULT kMissingFileOperationHostHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
constexpr uint32_t kMaxDropPathCount          = 65'536u;
constexpr size_t kMaxDropPathChars            = 32'768u;
constexpr size_t kMaxDropTotalChars           = 16u * 1024u * 1024u;

[[nodiscard]] std::wstring_view TrimTrailingPathSeparators(std::wstring_view path) noexcept
{
    while (path.size() > 1u && (path.back() == L'\\' || path.back() == L'/'))
    {
        path.remove_suffix(1u);
    }
    return path;
}

// FolderView provider paths use Windows-style case-insensitive identity but may use either separator.
// Keep this lexical: resolving symlinks/reparse points belongs to the operation host/provider.
[[nodiscard]] bool IsSameOrDescendantPath(std::wstring_view root, std::wstring_view candidate) noexcept
{
    root      = TrimTrailingPathSeparators(root);
    candidate = TrimTrailingPathSeparators(candidate);
    if (root.empty() || candidate.size() < root.size() || ! OrdinalString::StartsWithNoCase(candidate, root))
    {
        return false;
    }
    if (candidate.size() == root.size())
    {
        return true;
    }
    const wchar_t boundary = candidate[root.size()];
    return boundary == L'\\' || boundary == L'/';
}

[[nodiscard]] bool AreSameLocalVolume(const std::filesystem::path& source, const std::filesystem::path& destination) noexcept
{
    std::array<wchar_t, MAX_PATH> sourceVolume{};
    std::array<wchar_t, MAX_PATH> destinationVolume{};
    if (GetVolumePathNameW(source.c_str(), sourceVolume.data(), static_cast<DWORD>(sourceVolume.size())) == FALSE ||
        GetVolumePathNameW(destination.c_str(), destinationVolume.data(), static_cast<DWORD>(destinationVolume.size())) == FALSE)
    {
        return false;
    }
    return CompareStringOrdinal(sourceVolume.data(), -1, destinationVolume.data(), -1, TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool IsDirectFileOperationFallbackEnabledForDropSelfTest() noexcept
{
#ifdef ENABLE_TESTS
    return FolderView::DebugIsDirectFileOperationFallbackEnabledForSelfTest();
#else
    return false;
#endif
}
} // namespace

class FolderView::DropTarget final : public IDropTarget
{
public:
    explicit DropTarget(FolderView& owner) : _refCount(1), _owner(owner)
    {
    }

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    DropTarget(const DropTarget&)            = delete;
    DropTarget(DropTarget&&)                 = delete;
    DropTarget& operator=(const DropTarget&) = delete;
    DropTarget& operator=(DropTarget&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropTarget)
        {
            *ppvObject = static_cast<IDropTarget*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override
    {
        if (! effect)
        {
            return E_POINTER;
        }

        _currentDataObject.reset();
        if (! dataObject || ! _owner.HasFileDrop(dataObject))
        {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        _currentDataObject = dataObject;
        _allowedEffects    = *effect;
        _lastEffect        = _owner.ResolveDropEffect(keyState, _allowedEffects);
        *effect            = _lastEffect;

        EnsureHelper();
        if (_helper)
        {
            POINT pt{point.x, point.y};
            _helper->DragEnter(_owner.GetHWND(), dataObject, &pt, _lastEffect);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL point, DWORD* effect) override
    {
        if (! effect)
        {
            return E_POINTER;
        }

        if (! _currentDataObject)
        {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        _lastEffect = _owner.ResolveDropEffect(keyState, _allowedEffects);
        *effect     = _lastEffect;

        if (_helper)
        {
            POINT pt{point.x, point.y};
            _helper->DragOver(&pt, _lastEffect);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DragLeave() override
    {
        if (_helper)
        {
            _helper->DragLeave();
        }
        _currentDataObject.reset();
        _allowedEffects = DROPEFFECT_NONE;
        _lastEffect     = DROPEFFECT_NONE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Drop(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override
    {
        if (! effect)
        {
            return E_POINTER;
        }
        if (! dataObject)
        {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        EnsureHelper();
        if (_helper)
        {
            POINT pt{point.x, point.y};
            _helper->Drop(dataObject, &pt, _lastEffect);
        }

        const DWORD allowed = _allowedEffects ? _allowedEffects : *effect;
        POINT clientPoint{point.x, point.y};
        static_cast<void>(ScreenToClient(_owner.GetHWND(), &clientPoint));
        DWORD performed = DROPEFFECT_NONE;
        HRESULT hr      = _owner.PerformDrop(dataObject, keyState, allowed, clientPoint, &performed);
        _currentDataObject.reset();
        _allowedEffects = DROPEFFECT_NONE;
        _lastEffect     = performed;
        *effect         = performed;
        return hr;
    }

private:
    void EnsureHelper()
    {
        if (! _helper)
        {
            wil::com_ptr<IDropTargetHelper> helper;
            if (SUCCEEDED(CoCreateInstance(CLSID_DragDropHelper, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(helper.addressof()))))
            {
                _helper = std::move(helper);
            }
        }
    }

    std::atomic<ULONG> _refCount;
    FolderView& _owner;
    wil::com_ptr<IDataObject> _currentDataObject;
    wil::com_ptr<IDropTargetHelper> _helper;
    DWORD _allowedEffects = DROPEFFECT_NONE;
    DWORD _lastEffect     = DROPEFFECT_NONE;
};

void FolderView::EnsureDropTarget()
{
    if (! _hWnd || _dropTarget || _dropTargetRegistered || ! _oleInitialized)
    {
        return;
    }

    auto* target = new (std::nothrow) DropTarget(*this);
    if (! target)
    {
        return;
    }

    _dropTarget.attach(target);
    const HRESULT hrRegister = RegisterDragDrop(_hWnd.get(), _dropTarget.get());
    if (SUCCEEDED(hrRegister))
    {
        _dropTargetRegistered = true;
        return;
    }

    ReportError(L"RegisterDragDrop", hrRegister);
    _dropTarget.reset();
}

void FolderView::BeginDragDrop()
{
    if (_drag.anchorIndex == static_cast<size_t>(-1))
    {
        return;
    }

    auto paths = GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    std::wstring pluginId = _fileSystemPluginId;
    if (pluginId.empty() && _fileSystemMetadata && _fileSystemMetadata->id && _fileSystemMetadata->id[0] != L'\0')
    {
        pluginId = _fileSystemMetadata->id;
    }
    std::wstring instanceContext = _fileSystemInstanceContext;
    const bool includeHDrop      = CompareStringOrdinal(pluginId.c_str(), -1, L"builtin/file-system", -1, TRUE) == CSTR_EQUAL;

    auto* dataObjectRaw =
        new (std::nothrow) FolderViewDataObject(std::move(paths), std::move(pluginId), std::move(instanceContext), DROPEFFECT_COPY, includeHDrop);
    if (! dataObjectRaw)
    {
        return;
    }
    wil::com_ptr<IDataObject> dataObject;
    dataObject.attach(dataObjectRaw);

    auto* dropSourceRaw = new (std::nothrow) FolderViewDropSource();
    if (! dropSourceRaw)
    {
        return;
    }
    wil::com_ptr<IDropSource> dropSource;
    dropSource.attach(dropSourceRaw);

    ReleaseCapture();

    wil::com_ptr<IDragSourceHelper> helper;
    if (SUCCEEDED(CoCreateInstance(CLSID_DragDropHelper, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(helper.addressof()))))
    {
        POINT screenPt = _drag.startPoint;
        ClientToScreen(_hWnd.get(), &screenPt);
        helper->InitializeFromWindow(_hWnd.get(), &screenPt, dataObject.get());
    }

    DWORD effect = DROPEFFECT_NONE;
    HRESULT hr   = DoDragDrop(dataObject.get(), dropSource.get(), DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK, &effect);
    if (FAILED(hr))
    {
        ReportError(L"DoDragDrop", hr);
        return;
    }

    if (effect == DROPEFFECT_MOVE)
    {
        EnumerateFolder();
    }

    _drag.dragging = false;
}

#ifdef ENABLE_TESTS
HRESULT FolderView::DebugPerformFileDropForSelfTest(const std::vector<std::filesystem::path>& paths,
                                                    DWORD effect,
                                                    DWORD* performedEffect,
                                                    POINT clientPoint,
                                                    bool externalSource,
                                                    DWORD allowedEffects,
                                                    DWORD keyState) noexcept
{
    DWORD localPerformed = DROPEFFECT_NONE;
    if (! performedEffect)
    {
        performedEffect = &localPerformed;
    }
    *performedEffect = DROPEFFECT_NONE;

    auto* dataObjectRaw = new (std::nothrow) FolderViewDataObject(paths, L"builtin/file-system", std::wstring{}, effect, true, ! externalSource);
    if (! dataObjectRaw)
    {
        return E_OUTOFMEMORY;
    }

    wil::com_ptr<IDataObject> dataObject;
    dataObject.attach(dataObjectRaw);
    const DWORD effectiveAllowedEffects = allowedEffects == DROPEFFECT_NONE ? effect : allowedEffects;
    return PerformDrop(dataObject.get(), keyState, effectiveAllowedEffects, clientPoint, performedEffect);
}

HRESULT FolderView::DebugPerformDropFromDataObjectForSelfTest(
    IDataObject* dataObject, DWORD effect, DWORD* performedEffect, POINT clientPoint, DWORD keyState) noexcept
{
    DWORD localPerformed = DROPEFFECT_NONE;
    if (! performedEffect)
    {
        performedEffect = &localPerformed;
    }
    *performedEffect = DROPEFFECT_NONE;
    return PerformDrop(dataObject, keyState, effect, clientPoint, performedEffect);
}
#endif

DWORD FolderView::ResolveDropEffect(DWORD keyState, DWORD allowedEffects) const
{
    auto chooseEffect = [&](DWORD desired) -> DWORD { return (allowedEffects & desired) ? desired : DROPEFFECT_NONE; };

    const bool ctrl  = (keyState & MK_CONTROL) != 0;
    const bool shift = (keyState & MK_SHIFT) != 0;
    const bool alt   = (GetKeyState(VK_MENU) & 0x8000) != 0;

    if ((ctrl && shift) || alt)
    {
        if (auto effect = chooseEffect(DROPEFFECT_LINK); effect != DROPEFFECT_NONE)
        {
            return effect;
        }
    }
    if (shift)
    {
        if (auto effect = chooseEffect(DROPEFFECT_MOVE); effect != DROPEFFECT_NONE)
        {
            return effect;
        }
    }
    if (ctrl)
    {
        if (auto effect = chooseEffect(DROPEFFECT_COPY); effect != DROPEFFECT_NONE)
        {
            return effect;
        }
    }
    if (auto effect = chooseEffect(DROPEFFECT_COPY); effect != DROPEFFECT_NONE)
    {
        return effect;
    }
    if (auto effect = chooseEffect(DROPEFFECT_MOVE); effect != DROPEFFECT_NONE)
    {
        return effect;
    }
    if (auto effect = chooseEffect(DROPEFFECT_LINK); effect != DROPEFFECT_NONE)
    {
        return effect;
    }
    return DROPEFFECT_NONE;
}

bool FolderView::HasFileDrop(IDataObject* dataObject) const
{
    if (! dataObject)
    {
        return false;
    }

    {
        FORMATETC format{};
        format.cfFormat = static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat());
        format.dwAspect = DVASPECT_CONTENT;
        format.lindex   = -1;
        format.tymed    = TYMED_HGLOBAL;
        if (SUCCEEDED(dataObject->QueryGetData(&format)))
        {
            return true;
        }
    }

    FORMATETC format{};
    format.cfFormat = CF_HDROP;
    format.dwAspect = DVASPECT_CONTENT;
    format.lindex   = -1;
    format.tymed    = TYMED_HGLOBAL;
    return SUCCEEDED(dataObject->QueryGetData(&format));
}

HRESULT FolderView::PerformDrop(IDataObject* dataObject, DWORD keyState, DWORD allowedEffects, POINT clientPoint, DWORD* performedEffect)
{
    if (! dataObject || ! performedEffect)
    {
        return E_INVALIDARG;
    }
    *performedEffect = DROPEFFECT_NONE;

    if (! _currentFolder)
    {
        return E_FAIL;
    }
    if (! IsCurrentFolderEnumerated())
    {
        return DRAGDROP_S_CANCEL;
    }

    std::filesystem::path destination = _currentFolder.value();
    if (const auto hit = HitTest(clientPoint); hit.has_value() && hit.value() < _items.size() && _items[hit.value()].isDirectory)
    {
        destination = GetItemFullPath(_items[hit.value()]);
    }

    DWORD effect = ResolveDropEffect(keyState, allowedEffects);
    if (effect == DROPEFFECT_NONE)
    {
        return S_OK;
    }

    bool internalDrop = false;
    std::wstring sourcePluginId;
    std::wstring sourceInstanceContext;
    std::vector<std::filesystem::path> paths;

    const auto tryReadInternalDrop = [&]() noexcept -> HRESULT
    {
        FORMATETC internal{};
        internal.cfFormat = static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat());
        internal.dwAspect = DVASPECT_CONTENT;
        internal.lindex   = -1;
        internal.tymed    = TYMED_HGLOBAL;

        if (FAILED(dataObject->QueryGetData(&internal)))
        {
            return S_FALSE;
        }

        STGMEDIUM medium{};
        HRESULT hr = dataObject->GetData(&internal, &medium);
        if (FAILED(hr))
        {
            Debug::Warning(L"FolderView::PerformDrop: IDataObject::GetData(InternalFileDrop) failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return hr;
        }

        auto release = wil::scope_exit([&]() { ReleaseStgMedium(&medium); });

        if (medium.tymed != TYMED_HGLOBAL || ! medium.hGlobal)
        {
            return DV_E_TYMED;
        }

        const SIZE_T bytesAvailable = GlobalSize(medium.hGlobal);
        if (bytesAvailable < sizeof(uint32_t) * 4u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const void* raw = GlobalLock(medium.hGlobal);
        if (! raw)
        {
            return E_FAIL;
        }
        auto unlock = wil::scope_exit([&]() { GlobalUnlock(medium.hGlobal); });

        struct Header
        {
            uint32_t version              = 0;
            uint32_t pluginIdChars        = 0;
            uint32_t instanceContextChars = 0;
            uint32_t pathCount            = 0;
        };

        const auto* header = static_cast<const Header*>(raw);
        if (header->version != 1)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const std::byte* base = static_cast<const std::byte*>(raw);
        size_t offset         = sizeof(Header);

        auto readString = [&](uint32_t chars, std::wstring& out) noexcept -> bool
        {
            if (chars > kMaxDropPathChars)
            {
                return false;
            }
            const size_t bytes = (static_cast<size_t>(chars) + 1u) * sizeof(wchar_t);
            if (offset > bytesAvailable || bytes > bytesAvailable - offset)
            {
                return false;
            }

            const auto* wide = reinterpret_cast<const wchar_t*>(base + offset);
            if (wide[chars] != L'\0')
            {
                return false;
            }

            out.assign(wide, static_cast<size_t>(chars));
            offset += bytes;
            return true;
        };

        if (! readString(header->pluginIdChars, sourcePluginId) || ! readString(header->instanceContextChars, sourceInstanceContext))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const size_t remainingBytes              = offset <= bytesAvailable ? bytesAvailable - offset : 0u;
        constexpr size_t kMinimumPathRecordBytes = sizeof(uint32_t) + sizeof(wchar_t);
        if (header->pathCount > kMaxDropPathCount || static_cast<size_t>(header->pathCount) > remainingBytes / kMinimumPathRecordBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        paths.clear();
        paths.reserve(header->pathCount);
        size_t totalPathChars = 0u;
        for (uint32_t i = 0; i < header->pathCount; ++i)
        {
            if (offset > bytesAvailable || sizeof(uint32_t) > bytesAvailable - offset)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            uint32_t chars = 0;
            memcpy(&chars, base + offset, sizeof(chars));
            offset += sizeof(chars);

            if (chars > kMaxDropPathChars || totalPathChars > kMaxDropTotalChars - static_cast<size_t>(chars))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            totalPathChars += chars;

            std::wstring text;
            if (! readString(chars, text))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            if (! text.empty())
            {
                paths.emplace_back(std::move(text));
            }
        }

        if (paths.empty())
        {
            sourcePluginId.clear();
            sourceInstanceContext.clear();
            internalDrop = false;
            return S_FALSE;
        }

        internalDrop = true;
        return S_OK;
    };

    HRESULT hr = tryReadInternalDrop();
    if (hr == S_FALSE)
    {
        FORMATETC format{};
        format.cfFormat = CF_HDROP;
        format.dwAspect = DVASPECT_CONTENT;
        format.lindex   = -1;
        format.tymed    = TYMED_HGLOBAL;

        STGMEDIUM medium{};
        hr = dataObject->GetData(&format, &medium);
        if (FAILED(hr))
        {
            Debug::Warning(L"FolderView::PerformDrop: IDataObject::GetData(CF_HDROP) failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return hr;
        }

        auto releaseMedium = wil::scope_exit([&]() { ReleaseStgMedium(&medium); });

        if (medium.tymed != TYMED_HGLOBAL || ! medium.hGlobal)
        {
            return DV_E_TYMED;
        }

        const SIZE_T bytesAvailable = GlobalSize(medium.hGlobal);
        if (bytesAvailable < sizeof(DROPFILES))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const auto* dropFiles = static_cast<const DROPFILES*>(GlobalLock(medium.hGlobal));
        if (! dropFiles)
        {
            return E_FAIL;
        }
        auto unlock = wil::scope_exit([&]() { GlobalUnlock(medium.hGlobal); });

        if (! dropFiles->fWide || dropFiles->pFiles < sizeof(DROPFILES) || dropFiles->pFiles > bytesAvailable || (dropFiles->pFiles % alignof(wchar_t)) != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const size_t pathBytes = bytesAvailable - dropFiles->pFiles;
        if (pathBytes < 2u * sizeof(wchar_t) || (pathBytes % sizeof(wchar_t)) != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const auto* pathBuffer = reinterpret_cast<const wchar_t*>(reinterpret_cast<const std::byte*>(dropFiles) + dropFiles->pFiles);
        const size_t pathChars = pathBytes / sizeof(wchar_t);
        if (pathBuffer[pathChars - 1u] != L'\0' || pathBuffer[pathChars - 2u] != L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        // The structural bounds above make the shell parser safe to use while retaining its canonical
        // handling of DROPFILES details.
        const HDROP drop     = static_cast<HDROP>(medium.hGlobal);
        const UINT fileCount = DragQueryFileW(drop, 0xFFFFFFFFu, nullptr, 0u);
        if (fileCount > kMaxDropPathCount)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        paths.clear();
        paths.reserve(fileCount);
        size_t totalPathChars = 0u;
        for (UINT index = 0u; index < fileCount; ++index)
        {
            const UINT length = DragQueryFileW(drop, index, nullptr, 0u);
            if (length == 0u || length > kMaxDropPathChars || totalPathChars > kMaxDropTotalChars - static_cast<size_t>(length))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            std::wstring path(static_cast<size_t>(length) + 1u, L'\0');
            if (DragQueryFileW(drop, index, path.data(), length + 1u) != length)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            totalPathChars += length;
            path.resize(length);
            paths.emplace_back(std::move(path));
        }
    }
    else if (FAILED(hr))
    {
        return hr;
    }

    if (paths.empty())
    {
        return S_OK;
    }

    const bool hasModifier = (keyState & (MK_CONTROL | MK_SHIFT)) != 0 || (GetKeyState(VK_MENU) & 0x8000) != 0;
    if (! hasModifier && effect == DROPEFFECT_COPY && (allowedEffects & DROPEFFECT_MOVE) != 0 &&
        std::ranges::all_of(paths, [&](const std::filesystem::path& source) noexcept { return AreSameLocalVolume(source, destination); }))
    {
        effect = DROPEFFECT_MOVE;
    }

    const std::filesystem::path normalizedDestination = destination.lexically_normal();
    for (const auto& source : paths)
    {
        const std::filesystem::path normalizedSource = source.lexically_normal();
        const std::filesystem::path sourceParent     = normalizedSource.parent_path().lexically_normal();
        if ((! sourceParent.empty() && OrdinalString::EqualsNoCasePath(sourceParent, normalizedDestination)) ||
            IsSameOrDescendantPath(normalizedSource.native(), normalizedDestination.native()))
        {
            return DRAGDROP_S_CANCEL;
        }
    }

    HRESULT operationHr  = S_OK;
    DWORD reportedEffect = effect;
    switch (effect)
    {
        case DROPEFFECT_MOVE:
        case DROPEFFECT_COPY:
        {
            if (! _fileSystem)
            {
                operationHr = E_FAIL;
                break;
            }

            const FileSystemOperation operationType = effect == DROPEFFECT_COPY ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
            const FileSystemFlags flags             = FILESYSTEM_FLAG_RECURSIVE;

            if (_fileOperationRequestCallback)
            {
                FileOperationRequest request{};
                request.operation              = operationType;
                request.sourcePaths            = std::move(paths);
                request.sourceContextSpecified = internalDrop;
                request.sourcePluginId         = std::move(sourcePluginId);
                request.sourceInstanceContext  = std::move(sourceInstanceContext);
                request.destinationFolder      = destination;
                request.flags                  = flags;

                operationHr = _fileOperationRequestCallback(std::move(request));
                if (SUCCEEDED(operationHr) && effect == DROPEFFECT_MOVE && ! internalDrop)
                {
                    // The host operation is asynchronous. Reporting MOVE here would authorize an external
                    // source to delete before our worker has committed the move.
                    reportedEffect = DROPEFFECT_COPY;
                }
                break;
            }

            if (! IsDirectFileOperationFallbackEnabledForDropSelfTest())
            {
                Debug::Error(L"FolderView::PerformDrop missing FileOperationRequestCallback; refusing direct file-operation fallback");
                operationHr = kMissingFileOperationHostHr;
                break;
            }

            if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), operationType, paths, destination))
            {
                return DRAGDROP_S_CANCEL;
            }

            FileSystemArenaOwner arenaOwner;
            const wchar_t** sourcePaths = nullptr;
            unsigned long count         = 0;
            HRESULT buildHr             = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
            if (FAILED(buildHr))
            {
                operationHr = buildHr;
                break;
            }

            if (effect == DROPEFFECT_COPY)
            {
                operationHr = _fileSystem->CopyItems(sourcePaths, count, destination.c_str(), flags, nullptr, nullptr, nullptr);
            }
            else
            {
                operationHr = _fileSystem->MoveItems(sourcePaths, count, destination.c_str(), flags, nullptr, nullptr, nullptr);
            }
            break;
        }
        case DROPEFFECT_LINK:
        {
            for (const auto& path : paths)
            {
                wil::com_ptr<IShellLinkW> shellLink;
                HRESULT hrLink = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.addressof()));
                if (FAILED(hrLink))
                {
                    operationHr = hrLink;
                    break;
                }
                hrLink = shellLink->SetPath(path.c_str());
                if (FAILED(hrLink))
                {
                    operationHr = hrLink;
                    break;
                }
                if (const auto parent = path.parent_path(); ! parent.empty())
                {
                    shellLink->SetWorkingDirectory(parent.c_str());
                }
                std::wstring description = path.filename().wstring();
                if (! description.empty())
                {
                    shellLink->SetDescription(description.c_str());
                }

                wil::com_ptr<IPersistFile> persist;
                hrLink = shellLink->QueryInterface(IID_PPV_ARGS(persist.addressof()));
                if (FAILED(hrLink))
                {
                    operationHr = hrLink;
                    break;
                }

                std::error_code ec;
                std::filesystem::path linkPath;
                bool foundSlot = false;
                for (int attempt = 0; attempt < 256; ++attempt)
                {
                    linkPath = GenerateShortcutPath(destination, path, attempt);
                    if (! std::filesystem::exists(linkPath, ec))
                    {
                        ec.clear();
                        foundSlot = true;
                        break;
                    }
                }
                if (ec)
                {
                    operationHr = HrFromErrorCode(ec);
                    break;
                }
                if (! foundSlot)
                {
                    operationHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    break;
                }

                hrLink = SaveShellShortcutExactPath(persist.get(), linkPath);
                if (FAILED(hrLink))
                {
                    operationHr = hrLink;
                    break;
                }
            }
            break;
        }
        default: break;
    }

    if (FAILED(operationHr))
    {
        ReportError(L"Drop operation", operationHr);
        return operationHr;
    }

    if (effect == DROPEFFECT_MOVE || effect == DROPEFFECT_COPY || effect == DROPEFFECT_LINK)
    {
        EnumerateFolder();
    }

    *performedEffect = reportedEffect;
    return S_OK;
}

#include "FolderViewInternal.h"

void FolderView::CommandRename()
{
    RenameFocusedItem();
}

void FolderView::CommandView()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
    {
        return;
    }

    const auto& item = _items[_focusedIndex];
    if (item.isDirectory)
    {
        SetFolderPath(GetItemFullPath(item));
        return;
    }

    bool handled = false;
    if (_viewFileRequestCallback)
    {
        ViewFileRequest request;
        request.focusedPath = GetItemFullPath(item);

        for (const auto& candidate : _items)
        {
            if (candidate.isDirectory)
            {
                continue;
            }

            request.displayedFilePaths.push_back(GetItemFullPath(candidate));
            if (candidate.selected)
            {
                request.selectionPaths.push_back(GetItemFullPath(candidate));
            }
        }

        handled = _viewFileRequestCallback(request);
    }

    if (! handled)
    {
        ActivateFocusedItem();
    }
}

void FolderView::CommandDelete()
{
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
        const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
        if (root && PostMessageW(root, WM_COMMAND, MAKEWPARAM(IDM_PANE_DELETE, 0), 0) != 0)
        {
            return;
        }
    }

    DeleteSelectedItems();
}

HRESULT FolderView::CopySelectedItemsToFolder(const std::filesystem::path& destinationFolder)
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return S_FALSE;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_COPY, paths, destinationFolder))
    {
        return S_FALSE;
    }

    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = FILESYSTEM_COPY;
        request.sourcePaths       = std::move(paths);
        request.destinationFolder = destinationFolder;
        request.flags             = FILESYSTEM_FLAG_RECURSIVE;

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Copy", hrStart);
            return hrStart;
        }

        return hrStart;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return hr;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    hr                          = _fileSystem->CopyItems(sourcePaths, count, destinationFolder.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return hr;
    }

    return S_OK;
}

HRESULT FolderView::MoveSelectedItemsToFolder(const std::filesystem::path& destinationFolder)
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return S_FALSE;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_MOVE, paths, destinationFolder))
    {
        return S_FALSE;
    }

    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = FILESYSTEM_MOVE;
        request.sourcePaths       = std::move(paths);
        request.destinationFolder = destinationFolder;
        request.flags             = FILESYSTEM_FLAG_RECURSIVE;

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Move", hrStart);
            return hrStart;
        }

        return hrStart;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return hr;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    hr                          = _fileSystem->MoveItems(sourcePaths, count, destinationFolder.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return hr;
    }

    return S_OK;
}

void FolderView::DeleteSelectedItems()
{
    if (! _fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation   = FILESYSTEM_DELETE;
        request.sourcePaths = std::move(paths);
        request.flags       = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Delete", hrStart);
        }
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** pathArray = nullptr;
    unsigned long count       = 0;
    HRESULT hr                = BuildPathArrayArena(paths, arenaOwner, &pathArray, &count);
    if (FAILED(hr))
    {
        ReportError(L"Delete", hr);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    hr                          = _fileSystem->DeleteItems(pathArray, count, flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Delete", hr);
        return;
    }

    if (! _currentFolder || ! DirectoryInfoCache::GetInstance().IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

void FolderView::CopySelectionToClipboard()
{
    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
        return;

    std::wstring multiSz = BuildMultiSz(paths);
    SIZE_T bytes         = (multiSz.size()) * sizeof(wchar_t) + sizeof(DROPFILES);

    wil::unique_hglobal hMem(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! hMem)
        return;

    auto* data = static_cast<BYTE*>(GlobalLock(hMem.get()));
    if (! data)
        return;

    auto* drop   = reinterpret_cast<DROPFILES*>(data);
    drop->pFiles = sizeof(DROPFILES);
    drop->pt     = POINT{};
    drop->fNC    = FALSE;
    drop->fWide  = TRUE;

    auto* files = reinterpret_cast<wchar_t*>(data + drop->pFiles);
    memcpy(files, multiSz.c_str(), (multiSz.size()) * sizeof(wchar_t));
    GlobalUnlock(hMem.get());

    if (OpenClipboard(_hWnd.get()))
    {
        EmptyClipboard();
        SetClipboardData(CF_HDROP, hMem.release());
        CloseClipboard();
    }
}

void FolderView::PasteItemsFromClipboard()
{
    if (! OpenClipboard(_hWnd.get()))
        return;

    HANDLE handle = GetClipboardData(CF_HDROP);
    if (! handle)
    {
        CloseClipboard();
        return;
    }

    auto* drop = static_cast<DROPFILES*>(GlobalLock(handle));
    if (! drop)
    {
        CloseClipboard();
        return;
    }

    const wchar_t* current = reinterpret_cast<const wchar_t*>(reinterpret_cast<const BYTE*>(drop) + drop->pFiles);
    std::vector<std::filesystem::path> sources;
    while (*current)
    {
        sources.emplace_back(current);
        current += wcslen(current) + 1;
    }
    GlobalUnlock(handle);
    CloseClipboard();

    if (! _currentFolder || ! _fileSystem)
    {
        return;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_COPY, sources, _currentFolder.value()))
    {
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** paths = nullptr;
    unsigned long count   = 0;
    HRESULT hr            = BuildPathArrayArena(sources, arenaOwner, &paths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    hr                          = _fileSystem->CopyItems(paths, count, _currentFolder->c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return;
    }

    if (! _currentFolder || ! DirectoryInfoCache::GetInstance().IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

void FolderView::RenameFocusedItem()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
        return;

    if (! _fileSystem)
        return;

    const auto& item = _items[_focusedIndex];
    auto prompt      = PromptForRename(_hWnd.get(), std::wstring(item.displayName), item.isDirectory);
    if (! prompt || prompt->empty())
        return;

    const std::filesystem::path fullPath = GetItemFullPath(item);
    std::filesystem::path target         = fullPath.parent_path() / *prompt;
    const FileSystemFlags flags          = FILESYSTEM_FLAG_NONE;
    const HRESULT hr                     = _fileSystem->RenameItem(fullPath.c_str(), target.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Rename", hr);
        return;
    }

    if (! _currentFolder || ! DirectoryInfoCache::GetInstance().IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

void FolderView::ShowProperties()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
        return;

    const auto& item                     = _items[_focusedIndex];
    const std::filesystem::path fullPath = GetItemFullPath(item);

    if (_propertiesRequestCallback)
    {
        const HRESULT hr = _propertiesRequestCallback(fullPath);
        if (FAILED(hr))
        {
            ReportError(L"Properties", hr);
        }
        return;
    }

    SHObjectProperties(_hWnd.get(), SHOP_FILEPATH, fullPath.c_str(), nullptr);
}

void FolderView::MoveSelectedItems()
{
    if (! _fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return;
    }

    wil::com_ptr<IFileOpenDialog> dialog;
    if (FAILED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.addressof()))))
        return;

    DWORD options = 0;
    if (FAILED(dialog->GetOptions(&options)))
        return;
    dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
    if (FAILED(dialog->Show(_hWnd.get())))
        return;

    wil::com_ptr<IShellItem> result;
    if (FAILED(dialog->GetResult(result.addressof())))
        return;

    wil::unique_cotaskmem_string selectedPath;
    if (FAILED(result->GetDisplayName(SIGDN_FILESYSPATH, selectedPath.addressof())))
        return;

    std::filesystem::path destination(selectedPath.get());

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_MOVE, paths, destination))
    {
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    hr                          = _fileSystem->MoveItems(sourcePaths, count, destination.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return;
    }

    if (! _currentFolder || ! DirectoryInfoCache::GetInstance().IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

#include "FolderViewInternal.h"
#include "StartupMetrics.h"

namespace
{
struct WStringViewHash
{
    using is_transparent = void;

    size_t operator()(std::wstring_view value) const noexcept
    {
        // Case-insensitive: extensions should be treated as case-insensitive on Windows.
        // (Avoid duplicate extension queries for ".TXT" vs ".txt".)
        uint64_t hash = 14695981039346656037ull; // FNV-1a 64-bit offset basis
        for (const wchar_t ch : value)
        {
            const wchar_t lower = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            hash ^= static_cast<uint64_t>(lower);
            hash *= 1099511628211ull; // FNV-1a 64-bit prime
        }
        return static_cast<size_t>(hash);
    }

    size_t operator()(const std::wstring& value) const noexcept
    {
        return (*this)(std::wstring_view{value});
    }
};

struct WStringViewEq
{
    using is_transparent = void;

    bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        return wil::compare_string_ordinal(left, right, true) == wistd::weak_ordering::equivalent;
    }
};

[[nodiscard]] bool LooksLikeWindowsDrivePath(std::wstring_view text) noexcept
{
    if (text.size() < 2)
    {
        return false;
    }

    const wchar_t first = text[0];
    if (! ((first >= L'A' && first <= L'Z') || (first >= L'a' && first <= L'z')))
    {
        return false;
    }

    return text[1] == L':';
}

[[nodiscard]] bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\", 0) == 0 || text.rfind(L"//", 0) == 0;
}

[[nodiscard]] bool LooksLikeExtendedPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\?\\", 0) == 0 || text.rfind(L"\\\\.\\", 0) == 0 || text.rfind(L"//?/", 0) == 0 || text.rfind(L"//./", 0) == 0;
}

[[nodiscard]] bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return false;
    }

    if (LooksLikeExtendedPath(text))
    {
        return true;
    }

    if (LooksLikeUncPath(text))
    {
        return true;
    }

    return LooksLikeWindowsDrivePath(text);
}

std::filesystem::path NormalizeFolderPathForFocusMemory(std::filesystem::path folder)
{
    folder = folder.lexically_normal();
    while (! folder.empty() && ! folder.has_filename() && folder != folder.root_path())
    {
        folder = folder.parent_path();
    }
    return folder;
}

std::wstring NormalizeFocusMemoryKey(std::filesystem::path path)
{
    path             = path.lexically_normal();
    std::wstring key = path.generic_wstring();
    if (LooksLikeWindowsAbsolutePath(key))
    {
        for (auto& ch : key)
        {
            ch = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
        }
    }
    return key;
}

std::wstring NormalizeFocusMemoryFolderKey(const std::filesystem::path& folder)
{
    return NormalizeFocusMemoryKey(NormalizeFolderPathForFocusMemory(folder));
}

std::wstring NormalizeFocusMemoryRootKey(const std::filesystem::path& folder)
{
    const std::filesystem::path normalized = NormalizeFolderPathForFocusMemory(folder);
    const std::filesystem::path root       = normalized.root_path();
    if (root.empty())
    {
        return {};
    }
    return NormalizeFocusMemoryKey(root);
}
} // namespace

void FolderView::EnsureEnumerationThread()
{
    if (_enumerationThreadStarted)
    {
        return;
    }

    _enumerationThread        = std::jthread([this](std::stop_token stopToken) { EnumerationWorker(stopToken); });
    _enumerationThreadStarted = true;
}

void FolderView::StopEnumerationThread() noexcept
{
    if (! _enumerationThread.joinable())
    {
        _enumerationThreadStarted = false;
        return;
    }

    _enumerationThread.request_stop();
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath.reset();
        _iconLoadQueue.clear();
        _iconLoadingActive.store(false, std::memory_order_release);
    }
    _enumerationCv.notify_all();
    _enumerationThread        = std::jthread{};
    _enumerationThreadStarted = false;
}

void FolderView::EnumerationWorker(std::stop_token stopToken)
{
    // Icon extraction calls COM (IImageList::GetIcon via IconCache::ExtractSystemIcon). This worker runs in the background,
    // so initialize COM as MTA here.
    [[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);

    [[maybe_unused]] const std::stop_callback notifyOnStop(stopToken, [this]() noexcept { _enumerationCv.notify_all(); });
    while (! stopToken.stop_requested())
    {
        std::filesystem::path folder;
        uint64_t generation     = 0;
        bool hasEnumerationWork = false;

        {
            std::unique_lock lock(_enumerationMutex);
            _enumerationCv.wait(
                lock,
                [&]() { return stopToken.stop_requested() || _pendingEnumerationPath.has_value() || _iconLoadingActive.load(std::memory_order_acquire); });

            if (stopToken.stop_requested())
            {
                break;
            }

            hasEnumerationWork = _pendingEnumerationPath.has_value();
            if (hasEnumerationWork)
            {
                folder     = std::move(_pendingEnumerationPath.value());
                generation = _pendingEnumerationGeneration;
                _pendingEnumerationPath.reset();
            }
        }

        // Process folder enumeration if requested
        if (hasEnumerationWork && ! folder.empty())
        {
            auto payload = ExecuteEnumeration(folder, generation, stopToken);
            if (payload && ! stopToken.stop_requested() && generation == _enumerationGeneration.load(std::memory_order_acquire))
            {
                if (_hWnd)
                {
                    static_cast<void>(PostMessagePayload(_hWnd.get(), WndMsg::kFolderViewEnumerateComplete, 0, std::move(payload)));
                }
            }
        }

        // Process icon loading queue (if active)
        const bool iconActive = _iconLoadingActive.load(std::memory_order_acquire);
        Debug::Info(L"EnumerationWorker: checking icon loading, active={}", iconActive ? L"true" : L"false");
        if (iconActive)
        {
            ProcessIconLoadQueue();
        }
    }
}

std::unique_ptr<FolderView::EnumerationPayload>
FolderView::ExecuteEnumeration(const std::filesystem::path& folder, uint64_t generation, std::stop_token stopToken)
{
    TRACER_CTX(folder.c_str());

    using UniqueThreadpoolWork = wil::unique_any<PTP_WORK, decltype(&::CloseThreadpoolWork), ::CloseThreadpoolWork>;

    auto payload        = std::make_unique<EnumerationPayload>();
    payload->generation = generation;
    payload->status     = S_OK;

    if (! _fileSystem)
    {
        payload->status = HRESULT_FROM_WIN32(ERROR_DLL_NOT_FOUND);
        return payload;
    }

    auto borrowed = DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(_fileSystem.get(), folder, DirectoryInfoCache::BorrowMode::AllowEnumerate);
    if (borrowed.Status() != S_OK)
    {
        payload->status = borrowed.Status();
        return payload;
    }

    IFilesInformation* filesInformation = borrowed.Get();
    if (! filesInformation)
    {
        payload->status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return payload;
    }

    // Zero-copy: take a COM ref to keep arena buffer alive
    // This allows FolderItems to use string_view pointing into the buffer
    payload->arenaBuffer = filesInformation;
    payload->folder      = folder;

    unsigned long entryCount = 0;
    HRESULT hr               = filesInformation->GetCount(&entryCount);
    if (FAILED(hr))
    {
        entryCount = 0;
    }

    std::vector<FolderItem> directories;
    std::vector<FolderItem> files;
    // Pre-allocate with better estimates (reduces reallocations)
    const size_t estimatedDirs  = static_cast<size_t>(entryCount) / 4u; // Estimate ~25% directories
    const size_t estimatedFiles = static_cast<size_t>(entryCount);      // Upper bound for files
    directories.reserve(std::max(estimatedDirs, static_cast<size_t>(128u)));
    files.reserve(std::max(estimatedFiles, static_cast<size_t>(256u)));

    // Best-effort: this runs on a background worker; translate exceptions into a failed payload.
    try
    {
        FileInfo* entry = nullptr;
        hr              = filesInformation->GetBuffer(&entry);
        if (FAILED(hr))
        {
            payload->status = hr;
            return payload;
        }

        if (entry != nullptr)
        {
            unsigned long bufferSize = 0;
            hr                       = filesInformation->GetBufferSize(&bufferSize);
            if (FAILED(hr))
            {
                payload->status = hr;
                return payload;
            }

            unsigned long allocatedSize = 0;
            hr                          = filesInformation->GetAllocatedSize(&allocatedSize);
            if (FAILED(hr))
            {
                payload->status = hr;
                return payload;
            }

            if (allocatedSize < bufferSize || allocatedSize < sizeof(FileInfo))
            {
                payload->status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                return payload;
            }

            std::byte* base = reinterpret_cast<std::byte*>(entry);
            std::byte* end  = base + bufferSize;

            {
                Debug::Perf::Scope perf(L"FolderView.ExecuteEnumeration.BuildItems");
                const std::wstring_view folderText = folder.native();
                perf.SetDetail(folderText);
                perf.SetValue0(entryCount);

                const auto appendStableHash32 = [](uint32_t hash, std::wstring_view text) noexcept -> uint32_t
                {
                    static constexpr uint32_t kFnvPrime32 = 16777619u;
                    for (const wchar_t ch : text)
                    {
                        const uint16_t value = static_cast<uint16_t>(ch);

                        hash ^= static_cast<uint8_t>(value & 0xFFu);
                        hash *= kFnvPrime32;

                        hash ^= static_cast<uint8_t>((value >> 8) & 0xFFu);
                        hash *= kFnvPrime32;
                    }
                    return hash;
                };

                static constexpr std::wstring_view kStableHashSeparator = L"|";
                const uint32_t folderStableHashSeed                     = appendStableHash32(StableHash32(folderText), kStableHashSeparator);

                while (! stopToken.stop_requested())
                {
                    if (_enumerationGeneration.load(std::memory_order_acquire) != generation)
                    {
                        return nullptr;
                    }

                    const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);

                    // Zero-copy: create string_view pointing into arena buffer for displayName
                    FolderItem item{};
                    item.displayName = std::wstring_view(entry->FileName, nameChars);

                    // Stable hash used for rainbow rendering (avoid storing full paths per item).
                    {
                        item.stableHash32 = appendStableHash32(folderStableHashSeed, item.displayName);
                    }

                    item.isDirectory    = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                    item.fileAttributes = entry->FileAttributes;
                    item.lastWriteTime  = entry->LastWriteTime;
                    if (! item.isDirectory && entry->EndOfFile > 0)
                    {
                        item.sizeBytes = static_cast<uint64_t>(entry->EndOfFile);
                    }

                    // Compute extension offset for files (zero-copy)
                    if (! item.isDirectory && ! item.displayName.empty())
                    {
                        const size_t dotPos = item.displayName.rfind(L'.');
                        if (dotPos != std::wstring_view::npos && dotPos > 0)
                        {
                            item.extensionOffset = static_cast<uint16_t>(dotPos);
                            // Detect .lnk shortcuts
                            const auto ext  = item.displayName.substr(dotPos);
                            item.isShortcut = (ext.size() == 4 && (ext[1] == L'l' || ext[1] == L'L') && (ext[2] == L'n' || ext[2] == L'N') &&
                                               (ext[3] == L'k' || ext[3] == L'K'));
                        }
                    }

                    if (item.isDirectory)
                    {
                        directories.emplace_back(std::move(item));
                    }
                    else
                    {
                        files.emplace_back(std::move(item));
                    }

                    if (entry->NextEntryOffset == 0)
                    {
                        break;
                    }

                    if (entry->NextEntryOffset < sizeof(FileInfo))
                    {
                        payload->status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                        break;
                    }

                    std::byte* next = reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset;
                    if (next < base || next + sizeof(FileInfo) > end)
                    {
                        payload->status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                        break;
                    }

                    entry = reinterpret_cast<FileInfo*>(next);
                }

                perf.SetValue1(directories.size() + files.size());
            }
        }

        if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
        {
            return nullptr;
        }

        if (SUCCEEDED(payload->status))
        {
            Debug::Perf::Scope perf(L"FolderView.ExecuteEnumeration.SortMerge");
            perf.SetDetail(folder.native());
            perf.SetValue0(directories.size());
            perf.SetValue1(files.size());

            // Use CompareStringOrdinal for string_view comparison (handles non-null-terminated strings)
            auto compare = [](const FolderItem& a, const FolderItem& b)
            {
                const int result = CompareStringOrdinal(
                    a.displayName.data(), static_cast<int>(a.displayName.size()), b.displayName.data(), static_cast<int>(b.displayName.size()), TRUE);
                return result == CSTR_LESS_THAN;
            };

            // Use parallel sorting for large directories (threshold: 1000 items)
            constexpr size_t kParallelSortThreshold = 1000;
            if (directories.size() >= kParallelSortThreshold)
            {
                std::sort(std::execution::par, directories.begin(), directories.end(), compare);
            }
            else
            {
                std::sort(directories.begin(), directories.end(), compare);
            }

            if (files.size() >= kParallelSortThreshold)
            {
                std::sort(std::execution::par, files.begin(), files.end(), compare);
            }
            else
            {
                std::sort(files.begin(), files.end(), compare);
            }

            payload->items.reserve(directories.size() + files.size());
            payload->items.insert(payload->items.end(), std::make_move_iterator(directories.begin()), std::make_move_iterator(directories.end()));
            payload->items.insert(payload->items.end(), std::make_move_iterator(files.begin()), std::make_move_iterator(files.end()));

            Debug::Info(L"FolderView enumeration completed: {} directories, {} files (total: {})", directories.size(), files.size(), payload->items.size());

            // Step 1: Collect unique extensions that need icon queries (parallel optimization)
            struct ExtensionQuery
            {
                std::wstring extension;
                std::wstring queryPath;
                DWORD fileAttributes = 0;
            };
            std::unordered_map<std::wstring, ExtensionQuery, WStringViewHash, WStringViewEq> uniqueExtensions;
            std::vector<size_t> perFileIconIndices; // Items needing per-file icon lookup

            Debug::Perf::Scope iconPreparePerf(L"FolderView.ExecuteEnumeration.IconIndex.Prepare");
            iconPreparePerf.SetDetail(folder.native());
            for (size_t i = 0; i < payload->items.size(); ++i)
            {
                if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
                {
                    break;
                }

                auto& item = payload->items[i];
                std::wstring_view extension;
                DWORD fileAttributes = 0;

                if (item.isDirectory)
                {
                    // Check if this is a special folder that needs custom icon
                    if (IconCache::IsSpecialFolder((folder / item.displayName).wstring()))
                    {
                        // Special folders need per-file icon lookup
                        perFileIconIndices.push_back(i);
                        continue;
                    }

                    extension      = L"<directory>";
                    fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
                }
                else
                {
                    extension      = item.GetExtension();
                    fileAttributes = FILE_ATTRIBUTE_NORMAL;
                }

                // Check cache first
                auto cachedIndex = IconCache::GetInstance().GetIconIndexByExtension(extension);
                if (cachedIndex.has_value())
                {
                    item.iconIndex = cachedIndex.value();
                    continue;
                }

                // Check if per-file lookup required (only for whitelist like .exe, .dll, .ico, .lnk, .url)
                // Files without extensions should use extension-based caching with empty string key
                if (IconCache::GetInstance().RequiresPerFileLookup(extension))
                {
                    perFileIconIndices.push_back(i);
                    continue;
                }

                // Add to unique extensions for batch query
                if (uniqueExtensions.find(extension) == uniqueExtensions.end())
                {
                    // Use SHGFI_USEFILEATTRIBUTES with dummy paths - Windows looks up file associations by extension
                    // Folders need backslash-terminated path, files need path with extension
                    std::wstring ownedExtension(extension);
                    std::wstring queryPath = (ownedExtension == L"<directory>") ? L"C:\\DummyFolder\\" : (L"C:\\Dummy" + ownedExtension);
                    uniqueExtensions.emplace(ownedExtension, ExtensionQuery{ownedExtension, std::move(queryPath), fileAttributes});
                }
            }

            iconPreparePerf.SetValue0(uniqueExtensions.size());
            iconPreparePerf.SetValue1(perFileIconIndices.size());

            Debug::Info(L"FolderView: {} unique extensions to query, {} per-file icons", uniqueExtensions.size(), perFileIconIndices.size());

            // Step 2: Parallel query unique extensions using Windows Thread Pool
            if (! uniqueExtensions.empty() && ! stopToken.stop_requested() && _enumerationGeneration.load(std::memory_order_acquire) == generation)
            {
                Debug::Perf::Scope extQueryPerf(L"FolderView.ExecuteEnumeration.IconIndex.QueryExtensions");
                extQueryPerf.SetDetail(folder.native());
                extQueryPerf.SetValue0(uniqueExtensions.size());

                TRACER_CTX(L"FolderView: Parallel extension query");

                // Thread-safe result storage
                std::mutex resultsMutex;
                std::unordered_map<std::wstring, int, WStringViewHash, WStringViewEq> extensionResults;

                // Thread pool work callback context
                struct QueryWork
                {
                    ExtensionQuery query;
                    std::mutex* resultsMutex                 = nullptr;
                    decltype(extensionResults)* results      = nullptr;
                    std::atomic<bool>* stopRequested         = nullptr;
                    std::atomic<uint64_t>* generationCounter = nullptr;
                    uint64_t generation                      = 0;
                };

                std::atomic<bool> queryStopRequested{false};
                std::vector<std::unique_ptr<QueryWork>> workItems;
                std::vector<UniqueThreadpoolWork> threadPoolWorks;
                threadPoolWorks.reserve(uniqueExtensions.size());

                // Prepare work items
                for (const auto& [ext, query] : uniqueExtensions)
                {
                    auto work               = std::make_unique<QueryWork>();
                    work->query             = query;
                    work->resultsMutex      = &resultsMutex;
                    work->results           = &extensionResults;
                    work->stopRequested     = &queryStopRequested;
                    work->generationCounter = &_enumerationGeneration;
                    work->generation        = generation;

                    // Create thread pool work item
                    UniqueThreadpoolWork tpWork(::CreateThreadpoolWork(
                        [](PTP_CALLBACK_INSTANCE, PVOID context, PTP_WORK) noexcept
                        {
                            auto* work = static_cast<QueryWork*>(context);
                            if (work->stopRequested->load() ||
                                (work->generationCounter && work->generationCounter->load(std::memory_order_acquire) != work->generation))
                            {
                                return;
                            }

                            const auto iconIndex = IconCache::GetInstance().GetOrQueryIconIndexByExtension(work->query.extension, work->query.fileAttributes);
                            if (iconIndex.has_value())
                            {
                                std::lock_guard lock(*work->resultsMutex);
                                (*work->results)[work->query.extension] = iconIndex.value();
                            }
                        },
                        work.get(),
                        nullptr));

                    if (tpWork)
                    {
                        threadPoolWorks.push_back(std::move(tpWork));
                        workItems.push_back(std::move(work));
                    }
                }

                // Submit all work items
                for (auto& tpWork : threadPoolWorks)
                {
                    SubmitThreadpoolWork(tpWork.get());
                }

                // Wait for completion
                for (auto& tpWork : threadPoolWorks)
                {
                    const bool cancelPending = stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation;
                    if (cancelPending)
                    {
                        queryStopRequested.store(true, std::memory_order_release);
                    }
                    WaitForThreadpoolWorkCallbacks(tpWork.get(), cancelPending ? TRUE : FALSE);
                }

                if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
                {
                    return nullptr;
                }

                // Apply results to items
                for (auto& item : payload->items)
                {
                    if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
                    {
                        return nullptr;
                    }

                    if (item.iconIndex >= 0)
                    {
                        continue; // Already set from cache
                    }

                    const std::wstring_view ext = item.isDirectory ? std::wstring_view(L"<directory>") : item.GetExtension();

                    auto it = extensionResults.find(ext);
                    if (it != extensionResults.end())
                    {
                        item.iconIndex = it->second;
                    }
                }

                extQueryPerf.SetValue1(extensionResults.size());
            }

            // Step 3: Parallel query per-file icons using thread pool
            if (! perFileIconIndices.empty() && ! stopToken.stop_requested() && _enumerationGeneration.load(std::memory_order_acquire) == generation)
            {
                Debug::Perf::Scope perFileQueryPerf(L"FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons");
                perFileQueryPerf.SetDetail(folder.native());
                perFileQueryPerf.SetValue0(perFileIconIndices.size());

                TRACER_CTX(L"FolderView : Parallel per - file query");

                // Thread-safe result storage
                std::mutex perFileResultsMutex;
                std::unordered_map<size_t, int> perFileResults;

                struct PerFileWork
                {
                    size_t itemIndex;
                    std::wstring fullPath;
                    std::mutex* resultsMutex                 = nullptr;
                    std::unordered_map<size_t, int>* results = nullptr;
                    std::atomic<bool>* stopRequested         = nullptr;
                    std::atomic<uint64_t>* generationCounter = nullptr;
                    uint64_t generation                      = 0;
                };

                std::atomic<bool> perFileStopRequested{false};
                std::vector<std::unique_ptr<PerFileWork>> perFileWorks;
                perFileWorks.reserve(perFileIconIndices.size());

                uint64_t perFilePathChars = 0;
                {
                    Debug::Perf::Scope perFilePathsPerf(L"FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths");
                    perFilePathsPerf.SetDetail(folder.native());
                    perFilePathsPerf.SetValue0(perFileIconIndices.size());

                    for (size_t idx : perFileIconIndices)
                    {
                        auto work       = std::make_unique<PerFileWork>();
                        work->itemIndex = idx;
                        work->fullPath  = (folder / payload->items[idx].displayName).wstring();
                        perFilePathChars += static_cast<uint64_t>(work->fullPath.size());
                        work->resultsMutex      = &perFileResultsMutex;
                        work->results           = &perFileResults;
                        work->stopRequested     = &perFileStopRequested;
                        work->generationCounter = &_enumerationGeneration;
                        work->generation        = generation;
                        perFileWorks.push_back(std::move(work));
                    }

                    perFilePathsPerf.SetValue1(perFilePathChars);
                }

                std::vector<UniqueThreadpoolWork> perFileWorkItems;
                perFileWorkItems.reserve(perFileWorks.size());

                for (auto& work : perFileWorks)
                {
                    UniqueThreadpoolWork workItem(::CreateThreadpoolWork(
                        [](PTP_CALLBACK_INSTANCE, PVOID context, PTP_WORK) noexcept
                        {
                            auto* work = static_cast<PerFileWork*>(context);
                            if (work->stopRequested->load() ||
                                (work->generationCounter && work->generationCounter->load(std::memory_order_acquire) != work->generation))
                            {
                                return;
                            }

                            const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(work->fullPath.c_str(), 0, false);

                            std::lock_guard<std::mutex> lock(*work->resultsMutex);
                            (*work->results)[work->itemIndex] = iconIndex.value_or(-1);
                        },
                        work.get(),
                        nullptr));

                    if (workItem)
                    {
                        SubmitThreadpoolWork(workItem.get());
                        perFileWorkItems.push_back(std::move(workItem));
                    }
                }

                // Wait for all work items
                for (auto& workItem : perFileWorkItems)
                {
                    const bool cancelPending = stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation;
                    if (cancelPending)
                    {
                        perFileStopRequested.store(true, std::memory_order_release);
                    }
                    WaitForThreadpoolWorkCallbacks(workItem.get(), cancelPending ? TRUE : FALSE);
                }

                if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
                {
                    return nullptr;
                }

                // Apply results to items
                uint64_t perFileFailures = 0;
                for (const auto& [idx, iconIndex] : perFileResults)
                {
                    if (iconIndex < 0)
                    {
                        ++perFileFailures;
                    }
                    payload->items[idx].iconIndex = iconIndex;
                }

                perFileQueryPerf.SetValue1(perFileFailures);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FolderView::ExecuteEnumeration: Exception while enumerating '{}'", folder.c_str());
        payload->status = E_FAIL;
    }

    return payload;
}

void FolderView::CancelPendingEnumeration()
{
    _pendingExternalCommandAfterEnumeration.reset();
    _enumerationGeneration.fetch_add(1, std::memory_order_release);
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath.reset();
        _iconLoadQueue.clear();
        _iconLoadingActive.store(false, std::memory_order_release);
    }
    _enumerationCv.notify_one();

    _pendingBusyOverlay.reset();

    bool clearedBusyOverlay = false;
    bool hasOverlay         = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay && _errorOverlay->kind == ErrorOverlayKind::Enumeration && _errorOverlay->severity == OverlaySeverity::Busy)
        {
            _errorOverlay.reset();
            clearedBusyOverlay = true;
        }
        hasOverlay = _errorOverlay.has_value();
    }

    if (clearedBusyOverlay && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    StopOverlayTimer();

    if (! hasOverlay)
    {
        const uint64_t nowTickMs = GetTickCount64();
        if (! UpdateIncrementalSearchIndicatorAnimation(nowTickMs))
        {
            StopOverlayAnimation();
        }
    }
}

void FolderView::EnumerateFolder()
{
    // Stop idle layout pre-creation from previous folder
    if (_idleLayoutTimer != 0 && _hWnd)
    {
        KillTimer(_hWnd.get(), kIdleLayoutTimerId);
        _idleLayoutTimer = 0;
    }

    _items.clear();
    _columnCounts.clear();
    _columnPrefixSums.clear();
    _scrollOffset      = 0.0f;
    _horizontalOffset  = 0.0f;
    _itemMetricsCached = false;
    _focusedIndex      = static_cast<size_t>(-1);
    _anchorIndex       = static_cast<size_t>(-1);
    _hoveredIndex      = static_cast<size_t>(-1);

    LayoutItems();
    UpdateScrollMetrics();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (! _currentFolder || ! _hWnd)
    {
        return;
    }

    ClearErrorOverlay(ErrorOverlayKind::Enumeration);

    EnsureEnumerationThread();
    const uint64_t generation = _enumerationGeneration.fetch_add(1, std::memory_order_release) + 1;
    if (_pendingExternalCommandAfterEnumeration && _currentFolder)
    {
        const std::wstring currentKey = NormalizeFocusMemoryFolderKey(_currentFolder.value());
        const std::wstring targetKey  = NormalizeFocusMemoryFolderKey(_pendingExternalCommandAfterEnumeration->targetFolder);
        if (! currentKey.empty() && currentKey == targetKey)
        {
            _pendingExternalCommandAfterEnumeration->generation = generation;
        }
        else
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }
    }
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath       = *_currentFolder;
        _pendingEnumerationGeneration = generation;
    }
    _enumerationCv.notify_one();

    ScheduleBusyOverlay(generation, *_currentFolder);
}

void FolderView::OnDirectoryCacheDirty()
{
    if (! _currentFolder || ! _hWnd)
    {
        return;
    }

    const ULONGLONG now             = GetTickCount64();
    constexpr ULONGLONG kDebounceMs = 200;
    if (_lastDirectoryCacheRefreshTick != 0 && now - _lastDirectoryCacheRefreshTick < kDebounceMs)
    {
        return;
    }

    _lastDirectoryCacheRefreshTick = now;
    RequestRefreshFromCache();
}

void FolderView::RequestRefreshFromCache()
{
    if (! _currentFolder || ! _hWnd)
    {
        return;
    }

    EnsureEnumerationThread();
    const uint64_t generation = _enumerationGeneration.fetch_add(1, std::memory_order_release) + 1;
    if (_pendingExternalCommandAfterEnumeration && _currentFolder)
    {
        const std::wstring currentKey = NormalizeFocusMemoryFolderKey(_currentFolder.value());
        const std::wstring targetKey  = NormalizeFocusMemoryFolderKey(_pendingExternalCommandAfterEnumeration->targetFolder);
        if (! currentKey.empty() && currentKey == targetKey)
        {
            _pendingExternalCommandAfterEnumeration->generation = generation;
        }
        else
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }
    }
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath       = *_currentFolder;
        _pendingEnumerationGeneration = generation;
    }
    _enumerationCv.notify_one();
}

void FolderView::ApplyCurrentSort()
{
    ApplyCurrentSort({}, static_cast<size_t>(-1));
}

void FolderView::ApplyCurrentSort(std::wstring_view preferredFocusedPath, size_t fallbackFocusIndex)
{
    constexpr auto invalidIndex = static_cast<size_t>(-1);

    if (_items.empty())
    {
        _focusedIndex = invalidIndex;
        _anchorIndex  = invalidIndex;
        _hoveredIndex = invalidIndex;
        return;
    }

    Debug::Perf::Scope perf(L"FolderView.ApplyCurrentSort");
    perf.SetValue0(_items.size());

    std::wstring_view focusedName = preferredFocusedPath;
    if (focusedName.empty() && _focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        focusedName = _items[_focusedIndex].displayName;
    }

    std::unordered_set<std::wstring_view> selectedNames;
    selectedNames.reserve(_items.size());
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            selectedNames.insert(item.displayName);
        }
    }

    auto compareInt = [&](int cmp) noexcept
    {
        if (_sortDirection == SortDirection::Ascending)
        {
            return cmp < 0;
        }
        return cmp > 0;
    };

    auto compareName = [&](const FolderItem& a, const FolderItem& b) noexcept
    {
        const int cmpResult = CompareStringOrdinal(
            a.displayName.data(), static_cast<int>(a.displayName.size()), b.displayName.data(), static_cast<int>(b.displayName.size()), TRUE);
        const int cmp = (cmpResult == CSTR_LESS_THAN) ? -1 : ((cmpResult == CSTR_GREATER_THAN) ? 1 : 0);
        if (cmp != 0)
        {
            return compareInt(cmp);
        }

        const int caseCmpResult = CompareStringOrdinal(
            a.displayName.data(), static_cast<int>(a.displayName.size()), b.displayName.data(), static_cast<int>(b.displayName.size()), FALSE);
        const int caseCmp = (caseCmpResult == CSTR_LESS_THAN) ? -1 : ((caseCmpResult == CSTR_GREATER_THAN) ? 1 : 0);
        if (caseCmp != 0)
        {
            return compareInt(caseCmp);
        }

        return a.unsortedOrder < b.unsortedOrder;
    };

    auto compare = [&](const FolderItem& a, const FolderItem& b) noexcept
    {
        if (a.isDirectory != b.isDirectory)
        {
            return a.isDirectory && ! b.isDirectory;
        }

        switch (_sortBy)
        {
            case SortBy::Name: return compareName(a, b);
            case SortBy::Extension:
            {
                const auto extA        = a.GetExtension();
                const auto extB        = b.GetExtension();
                const int extCmpResult = CompareStringOrdinal(extA.data(), static_cast<int>(extA.size()), extB.data(), static_cast<int>(extB.size()), TRUE);
                const int extCmp       = (extCmpResult == CSTR_LESS_THAN) ? -1 : ((extCmpResult == CSTR_GREATER_THAN) ? 1 : 0);
                if (extCmp != 0)
                {
                    return compareInt(extCmp);
                }
                return compareName(a, b);
            }
            case SortBy::Time:
            {
                if (a.lastWriteTime != b.lastWriteTime)
                {
                    return _sortDirection == SortDirection::Ascending ? (a.lastWriteTime < b.lastWriteTime) : (a.lastWriteTime > b.lastWriteTime);
                }
                return compareName(a, b);
            }
            case SortBy::Size:
            {
                if (! a.isDirectory && ! b.isDirectory && a.sizeBytes != b.sizeBytes)
                {
                    return _sortDirection == SortDirection::Ascending ? (a.sizeBytes < b.sizeBytes) : (a.sizeBytes > b.sizeBytes);
                }
                return compareName(a, b);
            }
            case SortBy::Attributes:
            {
                if (a.fileAttributes != b.fileAttributes)
                {
                    return _sortDirection == SortDirection::Ascending ? (a.fileAttributes < b.fileAttributes) : (a.fileAttributes > b.fileAttributes);
                }
                return compareName(a, b);
            }
            case SortBy::None:
            {
                return a.unsortedOrder < b.unsortedOrder;
            }
        }

        return compareName(a, b);
    };

    // Use parallel sorting for large directories (threshold: 1000 items)
    constexpr size_t kParallelSortThreshold = 1000;
    if (_items.size() >= kParallelSortThreshold)
    {
        std::stable_sort(std::execution::par, _items.begin(), _items.end(), compare);
    }
    else
    {
        std::stable_sort(_items.begin(), _items.end(), compare);
    }

    size_t newFocusedIndex = invalidIndex;
    size_t firstSelected   = invalidIndex;
    SelectionStats stats{};
    const FolderItem* singleSelected = nullptr;
    uint32_t selectedTotal           = 0;
    for (size_t i = 0; i < _items.size(); ++i)
    {
        auto& item    = _items[i];
        item.selected = selectedNames.contains(item.displayName);
        item.focused  = false;

        if (item.selected)
        {
            ++selectedTotal;
            if (selectedTotal == 1)
            {
                singleSelected = &item;
            }
            else
            {
                singleSelected = nullptr;
            }
            if (item.isDirectory)
            {
                ++stats.selectedFolders;
            }
            else
            {
                ++stats.selectedFiles;
                stats.selectedFileBytes += item.sizeBytes;
            }
        }

        if (firstSelected == invalidIndex && item.selected)
        {
            firstSelected = i;
        }

        if (! focusedName.empty() && item.displayName == focusedName)
        {
            newFocusedIndex = i;
        }
    }

    if (newFocusedIndex == invalidIndex)
    {
        if (firstSelected != invalidIndex)
        {
            newFocusedIndex = firstSelected;
        }
        else if (fallbackFocusIndex != invalidIndex)
        {
            newFocusedIndex = std::min(fallbackFocusIndex, _items.size() - 1);
        }
        else
        {
            newFocusedIndex = 0;
        }
    }

    _focusedIndex = newFocusedIndex;
    _anchorIndex  = newFocusedIndex;

    if (_focusedIndex < _items.size())
    {
        _items[_focusedIndex].focused = true;
    }

    if (selectedTotal == 1 && singleSelected)
    {
        SelectionStats::SelectedItemDetails details{};
        details.isDirectory    = singleSelected->isDirectory;
        details.sizeBytes      = singleSelected->sizeBytes;
        details.lastWriteTime  = singleSelected->lastWriteTime;
        details.fileAttributes = singleSelected->fileAttributes;
        stats.singleItem       = details;
    }

    _hoveredIndex   = static_cast<size_t>(-1);
    _selectionStats = stats;
    NotifySelectionChanged();
    RememberFocusedItemForDisplayedFolder();
}

void FolderView::RememberFocusedItemForDisplayedFolder() noexcept
{
    constexpr auto invalidIndex = static_cast<size_t>(-1);

    if (! _displayedFolder)
    {
        return;
    }

    if (_items.empty())
    {
        return;
    }

    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return;
    }

    EnsureFocusMemoryRootForFolder(_displayedFolder.value());

    const std::wstring folderKey = NormalizeFocusMemoryFolderKey(_displayedFolder.value());
    if (folderKey.empty())
    {
        return;
    }

    _focusMemory.insert_or_assign(folderKey, std::wstring(_items[_focusedIndex].displayName));
}

void FolderView::RememberFocusedItemForFolder(const std::filesystem::path& folder, std::wstring_view itemDisplayName) noexcept
{
    if (itemDisplayName.empty())
    {
        return;
    }

    const std::wstring rootKey = NormalizeFocusMemoryRootKey(folder);
    if (_focusMemoryRootKey != rootKey)
    {
        _focusMemory.clear();
        _focusMemoryRootKey = rootKey;
    }

    const std::wstring folderKey = NormalizeFocusMemoryFolderKey(folder);
    if (folderKey.empty())
    {
        return;
    }

    _focusMemory.insert_or_assign(folderKey, std::wstring(itemDisplayName));
}

void FolderView::QueueCommandAfterNextEnumeration(UINT commandId,
                                                 const std::filesystem::path& targetFolder,
                                                 std::wstring_view expectedFocusDisplayName) noexcept
{
    if (commandId == 0 || targetFolder.empty())
    {
        _pendingExternalCommandAfterEnumeration.reset();
        return;
    }

    PendingExternalCommand pending{};
    pending.commandId   = commandId;
    pending.generation  = 0;
    pending.targetFolder = targetFolder;
    if (! expectedFocusDisplayName.empty())
    {
        pending.expectedFocusDisplayName.assign(expectedFocusDisplayName);
    }

    _pendingExternalCommandAfterEnumeration = std::move(pending);
}

void FolderView::EnsureFocusMemoryRootForFolder(const std::filesystem::path& folder) noexcept
{
    const std::wstring rootKey = NormalizeFocusMemoryRootKey(folder);
    if (_focusMemoryRootKey != rootKey)
    {
        _focusMemory.clear();
        _focusMemoryRootKey = rootKey;
    }
}

std::wstring FolderView::GetRememberedFocusedItemPathForFolder(const std::filesystem::path& folder) noexcept
{
    EnsureFocusMemoryRootForFolder(folder);

    const std::wstring folderKey = NormalizeFocusMemoryFolderKey(folder);
    if (folderKey.empty())
    {
        return {};
    }

    const auto it = _focusMemory.find(folderKey);
    if (it == _focusMemory.end())
    {
        return {};
    }

    return it->second;
}

void FolderView::ProcessEnumerationResult(std::unique_ptr<EnumerationPayload> payload)
{
    TRACER;
    if (! payload)
    {
        return;
    }

    const uint64_t currentGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    if (payload->generation != currentGeneration)
    {
        if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }
        return;
    }

    CancelBusyOverlay(payload->generation);

    if (FAILED(payload->status))
    {
        if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }

        ReportError(L"EnumerateFolder", payload->status);
        UpdateScrollMetrics();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    ClearErrorOverlay(ErrorOverlayKind::Enumeration);

    ExitIncrementalSearch();

    const auto invalidIndex     = static_cast<size_t>(-1);
    size_t previousFocusedIndex = invalidIndex;
    std::wstring previousFocusName;
    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        previousFocusedIndex = _focusedIndex;
        previousFocusName.assign(_items[_focusedIndex].displayName);
    }

    bool isRefresh = false;
    if (_displayedFolder && _currentFolder)
    {
        isRefresh = NormalizeFocusMemoryFolderKey(_displayedFolder.value()) == NormalizeFocusMemoryFolderKey(_currentFolder.value());
    }

    std::wstring preferredFocusPath;
    size_t fallbackFocusIndex = invalidIndex;
    if (isRefresh)
    {
        preferredFocusPath = previousFocusName;
        fallbackFocusIndex = previousFocusedIndex;

        if (_currentFolder)
        {
            const std::wstring rememberedFocusPath = GetRememberedFocusedItemPathForFolder(_currentFolder.value());
            if (! rememberedFocusPath.empty() && rememberedFocusPath != previousFocusName)
            {
                preferredFocusPath = rememberedFocusPath;
            }
        }
    }
    else if (_currentFolder)
    {
        preferredFocusPath = GetRememberedFocusedItemPathForFolder(_currentFolder.value());
        fallbackFocusIndex = invalidIndex;
    }

    // Incremental refresh: preserve rendering state for unchanged items
    size_t itemsPreserved = 0;
    if (isRefresh && ! _items.empty())
    {
        // Build lookup map of old items by path for O(1) access
        std::unordered_map<std::wstring_view, size_t> oldItemsByPath;
        oldItemsByPath.reserve(_items.size());
        for (size_t i = 0; i < _items.size(); ++i)
        {
            oldItemsByPath[_items[i].displayName] = i;
        }

        // Transfer rendering state from matching old items to new items
        for (auto& newItem : payload->items)
        {
            auto it = oldItemsByPath.find(newItem.displayName);
            if (it == oldItemsByPath.end())
            {
                continue; // New item, no state to transfer
            }

            const auto& oldItem = _items[it->second];

            // Check if item data is unchanged (same size, time, attributes)
            const bool dataUnchanged = (oldItem.sizeBytes == newItem.sizeBytes && oldItem.lastWriteTime == newItem.lastWriteTime &&
                                        oldItem.fileAttributes == newItem.fileAttributes && oldItem.isDirectory == newItem.isDirectory);
            if (! dataUnchanged)
            {
                continue; // Item modified, needs fresh rendering
            }

            // Transfer rendering state from old item
            newItem.labelLayout    = oldItem.labelLayout;
            newItem.labelMetrics   = oldItem.labelMetrics;
            newItem.detailsText    = oldItem.detailsText;
            newItem.detailsLayout  = oldItem.detailsLayout;
            newItem.detailsMetrics = oldItem.detailsMetrics;
            // Only preserve D2D bitmap if icon index matches (icons are shared by extension)
            if (oldItem.iconIndex == newItem.iconIndex && oldItem.icon)
            {
                newItem.icon = oldItem.icon;
            }

            // Preserve selection state
            newItem.selected = oldItem.selected;

            ++itemsPreserved;
        }

        if (itemsPreserved > 0)
        {
            Debug::Info(L"FolderView: Incremental refresh preserved {} of {} items", itemsPreserved, payload->items.size());
        }
    }

    _items            = std::move(payload->items);
    _itemsArenaBuffer = std::move(payload->arenaBuffer); // Keep arena alive for string_views
    _itemsFolder      = std::move(payload->folder);      // For computing full paths
    for (size_t i = 0; i < _items.size(); ++i)
    {
        _items[i].unsortedOrder = i;
    }
    _displayedFolder = _currentFolder;
    _focusedIndex    = invalidIndex;
    _anchorIndex     = invalidIndex;
    _hoveredIndex    = invalidIndex;
    ApplyCurrentSort(preferredFocusPath, fallbackFocusIndex);

    // Only reset scroll position on folder navigation, not on refresh
    if (! isRefresh)
    {
        _scrollOffset     = 0.0f;
        _horizontalOffset = 0.0f;
    }
    _itemMetricsCached = false;

    if (_detailsTextProvider && _displayMode == DisplayMode::Detailed)
    {
        for (auto& item : _items)
        {
            if (item.displayName.empty())
            {
                continue;
            }

            std::wstring details = _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            if (details != item.detailsText)
            {
                item.detailsText    = std::move(details);
                item.detailsLayout.reset();
                item.detailsMetrics = {};
            }
        }
    }

    // Items already have iconIndex populated from ExecuteEnumeration background thread
    // Now queue icon loading to convert HICON to D2D bitmaps on UI thread
    LayoutItems();
    UpdateScrollMetrics();
    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        EnsureVisible(_focusedIndex);
    }

    // Queue icon loading after layout - only for items without D2D bitmaps
    Debug::Info(L"FolderView: About to queue icons for {} items", _items.size());
    QueueIconLoading();

    if (! _items.empty())
    {
        const std::wstring_view folderDetail = _itemsFolder.empty() ? std::wstring_view{} : std::wstring_view(_itemsFolder.native());
        StartupMetrics::MarkFirstPanePopulated(folderDetail.empty() ? std::wstring_view(L"(unknown)") : folderDetail, static_cast<uint64_t>(_items.size()));
    }

    // Schedule idle-time layout pre-creation for off-screen items
    // This creates layouts gradually during UI idle periods for smoother scrolling
    ScheduleIdleLayoutCreation();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (_enumerationCompletedCallback)
    {
        _enumerationCompletedCallback(_itemsFolder);
    }

    if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
    {
        PendingExternalCommand pending = std::move(_pendingExternalCommandAfterEnumeration.value());
        _pendingExternalCommandAfterEnumeration.reset();

        bool focusMatches = true;
        if (! pending.expectedFocusDisplayName.empty())
        {
            focusMatches = false;
            if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
            {
                const std::wstring_view focusedName = _items[_focusedIndex].displayName;
                focusMatches = (focusedName == pending.expectedFocusDisplayName) || OrdinalString::EqualsNoCase(focusedName, pending.expectedFocusDisplayName);
            }
        }

        if (focusMatches && pending.commandId != 0u && _hWnd)
        {
            PostMessageW(_hWnd.get(), WM_COMMAND, MAKEWPARAM(pending.commandId, 0), 0);
        }
    }
}

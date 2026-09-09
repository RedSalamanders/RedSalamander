#include "FolderViewInternal.h"
#include "FolderViewSortPolicy.h"
#include "PathUtils.h"
#include "StartupMetrics.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

namespace
{
struct WStringViewHash
{
    using is_transparent = void;
    bool ignoreCase = true;

    size_t operator()(std::wstring_view value) const noexcept
    {
        // Case-insensitive: extensions should be treated as case-insensitive on Windows.
        // (Avoid duplicate extension queries for ".TXT" vs ".txt".)
        uint64_t hash = 14695981039346656037ull; // FNV-1a 64-bit offset basis
        for (const wchar_t ch : value)
        {
            const wchar_t effective = ignoreCase ? static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))) : ch;
            hash ^= static_cast<uint64_t>(effective);
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
    bool ignoreCase = true;

    bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        return wil::compare_string_ordinal(left, right, ignoreCase) == wistd::weak_ordering::equivalent;
    }
};

[[nodiscard]] bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(text);
    return pathClass != Common::Paths::WindowsPathClass::Relative && pathClass != Common::Paths::WindowsPathClass::Rooted;
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

} // namespace

void FolderView::UpdateCompareNoDifferencesState() noexcept
{
    const auto resetCompareNoDifferencesLayouts = [&]() noexcept
    {
        _emptyMessageIconLayout.reset();
        _emptyMessageTitleLayout.reset();
        _emptyMessageFunLayout.reset();
        _emptyMessageLayoutClientSizePx = {};
        _emptyMessageLayoutDpi          = 0.0f;
        _emptyMessageLayoutMessageId    = 0;
        _emptyMessageIconFontSizeDip    = 0.0f;
        _emptyMessageIconMetrics        = {};
        _emptyMessageTitleMetrics       = {};
        _emptyMessageFunMetrics         = {};
    };

    const bool shouldShowNoDifferences =
        _items.empty() && _emptyStateMessageKind == EmptyStateMessageKind::CompareNoDifferences && _displayedFolder.has_value();
    if (! shouldShowNoDifferences)
    {
        if (_compareNoDifferencesState.has_value())
        {
            _compareNoDifferencesState.reset();
            resetCompareNoDifferencesLayouts();
        }
        return;
    }

    constexpr UINT kFunMessages[] = {
        IDS_COMPARE_NO_DIFFERENCES_FUN_1,
        IDS_COMPARE_NO_DIFFERENCES_FUN_2,
        IDS_COMPARE_NO_DIFFERENCES_FUN_3,
        IDS_COMPARE_NO_DIFFERENCES_FUN_4,
        IDS_COMPARE_NO_DIFFERENCES_FUN_5,
        IDS_COMPARE_NO_DIFFERENCES_FUN_6,
        IDS_COMPARE_NO_DIFFERENCES_FUN_7,
        IDS_COMPARE_NO_DIFFERENCES_FUN_8,
        IDS_COMPARE_NO_DIFFERENCES_FUN_9,
        IDS_COMPARE_NO_DIFFERENCES_FUN_10,
    };

    const std::wstring folderKey = NormalizeFocusMemoryFolderKey(_displayedFolder.value());
    const bool needsNewMessage =
        ! _compareNoDifferencesState.has_value() || _compareNoDifferencesState->folderKey != folderKey || _compareNoDifferencesState->funMessageResourceId == 0;
    if (! needsNewMessage)
    {
        return;
    }

    const ULONGLONG tick  = GetTickCount64();
    const uint32_t tick32 = static_cast<uint32_t>(tick ^ (tick >> 32));
    const uint32_t seed   = StableHash32(std::wstring_view(folderKey)) ^ tick32;
    const UINT messageId  = kFunMessages[static_cast<size_t>(seed % static_cast<uint32_t>(std::size(kFunMessages)))];

    std::wstring raw = LoadStringResource(nullptr, messageId);
    std::wstring emoji;
    std::wstring funMessage;

    const size_t breakPos = raw.find_first_of(L"\r\n");
    if (breakPos == std::wstring::npos)
    {
        funMessage = StringUtils::TrimWhitespaceCopy(raw);
    }
    else
    {
        emoji = StringUtils::TrimWhitespaceCopy(std::wstring_view(raw).substr(0, breakPos));

        size_t messageStart = breakPos;
        while (messageStart < raw.size() && (raw[messageStart] == L'\r' || raw[messageStart] == L'\n'))
        {
            ++messageStart;
        }
        funMessage = StringUtils::TrimWhitespaceCopy(std::wstring_view(raw).substr(messageStart));
    }

    CompareNoDifferencesState state{};
    state.folderKey            = folderKey;
    state.funMessageResourceId = messageId;
    state.emoji                = std::move(emoji);
    state.funMessage           = std::move(funMessage);
    _compareNoDifferencesState = std::move(state);

    resetCompareNoDifferencesLayouts();
}

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
        _pendingEnumerationFileSystem.reset();
        _pendingEnumerationPluginId.clear();
        _pendingEnumerationInstanceContext.clear();
        _iconLoadQueue.clear();
        _thumbnailLoadQueue.clear();
        _iconLoadingActive.store(false, std::memory_order_release);
        _thumbnailLoadingActive.store(false, std::memory_order_release);
    }
    _enumerationCv.notify_all();
    _enumerationThread        = std::jthread{};
    _enumerationThreadStarted = false;
}

void FolderView::DrainPendingEnumerationPayloadMessages() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    size_t drainedCount = 0u;
    MSG msg{};
    while (PeekMessageW(&msg, _hWnd.get(), WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_REMOVE) != FALSE)
    {
        auto payload = TakeMessagePayload<EnumerationPayload>(msg.lParam);
        ++drainedCount;
    }

#ifdef ENABLE_TESTS
    if (drainedCount > 0u)
    {
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::DrainPendingEnumerationPayloadMessages: drained={}", drainedCount));
    }
#endif
}

void FolderView::EnumerationWorker(std::stop_token stopToken)
{
    // Icon extraction calls COM (IImageList::GetIcon via IconCache::ExtractSystemIcon). This worker runs in the background,
    // so initialize COM as MTA here.
    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(coinitHr))
    {
        Debug::Error(L"FolderView enumeration: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
        FAIL_FAST_IF_FAILED(coinitHr);
    }
    [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

    [[maybe_unused]] const std::stop_callback notifyOnStop(stopToken, [this]() noexcept { _enumerationCv.notify_all(); });
    while (! stopToken.stop_requested())
    {
        std::filesystem::path folder;
        uint64_t generation     = 0;
        bool hasEnumerationWork = false;
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring pluginId;
        std::wstring instanceContext;

        {
            std::unique_lock lock(_enumerationMutex);
            _enumerationCv.wait(lock,
                                [&]()
            {
                return stopToken.stop_requested() || _pendingEnumerationPath.has_value() || _iconLoadingActive.load(std::memory_order_acquire) ||
                       _thumbnailLoadingActive.load(std::memory_order_acquire);
            });

            if (stopToken.stop_requested())
            {
                break;
            }

            hasEnumerationWork = _pendingEnumerationPath.has_value();
            if (hasEnumerationWork)
            {
                folder     = std::move(_pendingEnumerationPath.value());
                generation = _pendingEnumerationGeneration;
                fileSystem = std::move(_pendingEnumerationFileSystem);
                pluginId = std::move(_pendingEnumerationPluginId);
                instanceContext = std::move(_pendingEnumerationInstanceContext);
                _pendingEnumerationPath.reset();
            }
        }

        // Process folder enumeration if requested
        if (hasEnumerationWork && ! folder.empty())
        {
            auto payload = ExecuteEnumeration(folder,
                                              generation,
                                              stopToken,
                                              std::move(fileSystem),
                                              std::move(pluginId),
                                              std::move(instanceContext));
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

        const bool thumbnailActive = _thumbnailLoadingActive.load(std::memory_order_acquire);
        if (thumbnailActive)
        {
            ProcessThumbnailLoadQueue(stopToken);
        }
    }
}

std::unique_ptr<FolderView::EnumerationPayload> FolderView::ExecuteEnumeration(const std::filesystem::path& folder,
                                                                               uint64_t generation,
                                                                               std::stop_token stopToken,
                                                                               wil::com_ptr<IFileSystem> fileSystem,
                                                                               std::wstring pluginId,
                                                                               std::wstring instanceContext)
{
    TRACER_CTX(folder.c_str());

    using UniqueThreadpoolWork = wil::unique_any<PTP_WORK, decltype(&::CloseThreadpoolWork), ::CloseThreadpoolWork>;

    auto payload        = std::make_unique<EnumerationPayload>();
    payload->generation = generation;
    payload->status     = S_OK;

    if (! fileSystem)
    {
        payload->status = HRESULT_FROM_WIN32(ERROR_DLL_NOT_FOUND);
        return payload;
    }

    auto borrowed =
        DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(fileSystem.get(), folder, DirectoryInfoCache::BorrowMode::AllowEnumerate, stopToken);
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

    // Best-effort: this runs on a background worker; translate exceptions into a failed payload.
    try
    {
        constexpr size_t kEnumerationReserveCeiling = 1u << 20;
        unsigned long bufferSizeHint                = 0;
        if (FAILED(filesInformation->GetBufferSize(&bufferSizeHint)))
        {
            bufferSizeHint = 0;
        }
        const size_t countHint = static_cast<size_t>(entryCount);
        const size_t maxByBuffer = bufferSizeHint > 0u ? static_cast<size_t>(bufferSizeHint) / sizeof(FileInfo) : countHint;
        const size_t clampedCount = std::min(countHint, std::min(maxByBuffer, kEnumerationReserveCeiling));

        // GetCount is an untrusted allocation hint; reserve only a bounded amount derived from the real buffer.
        const size_t estimatedDirs  = clampedCount / 4u;
        const size_t estimatedFiles = clampedCount;
        directories.reserve(std::max(estimatedDirs, static_cast<size_t>(128u)));
        files.reserve(std::max(estimatedFiles, static_cast<size_t>(256u)));

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

                const uint32_t folderStableHashSeed = AppendStableHash32(StableHash32(folderText), L"|");

                const bool showHiddenFiles = _showHiddenFiles.load(std::memory_order_acquire);
                const bool showSystemFiles = _showSystemFiles.load(std::memory_order_acquire);

                const auto nameFilter   = _nameFilter.load(std::memory_order_acquire);
                const bool filterActive = nameFilter && nameFilter->state.enabled && nameFilter->hasMask;

                const auto hiddenNames  = _hiddenNames.load(std::memory_order_acquire);
                const bool hiddenActive = hiddenNames && ! hiddenNames->names.empty();

                while (! stopToken.stop_requested())
                {
                    if (_enumerationGeneration.load(std::memory_order_acquire) != generation)
                    {
                        return nullptr;
                    }

                    const DWORD fileAttributes = entry->FileAttributes;
                    bool include               = (showHiddenFiles || (fileAttributes & FILE_ATTRIBUTE_HIDDEN) == 0) &&
                                                 (showSystemFiles || (fileAttributes & FILE_ATTRIBUTE_SYSTEM) == 0);

                    std::wstring_view displayName;
                    if (include)
                    {
                        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                        displayName            = std::wstring_view(entry->FileName, nameChars);

                        if (filterActive && ! MaskSyntax::MatchesWildcardMask(displayName, nameFilter->mask))
                        {
                            include = false;
                        }
                        else if (hiddenActive && hiddenNames->names.contains(displayName))
                        {
                            include = false;
                        }
                    }

                    if (include)
                    {
                        // Zero-copy: create string_view pointing into arena buffer for displayName
                        FolderItem item{};
                        item.displayName = displayName;

                        // Stable hash used for rainbow rendering (avoid storing full paths per item).
                        {
                            item.stableHash32 = AppendStableHash32(folderStableHashSeed, item.displayName);
                        }

                        item.isDirectory    = (fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                        item.fileAttributes = fileAttributes;
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

            auto compare = [](const FolderItem& a, const FolderItem& b) { return OrdinalString::Compare(a.displayName, b.displayName, true) < 0; };

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

            size_t artifactCandidateCount = 0u;
            size_t artifactProjectedCount = 0u;
            const auto artifactStartedAt = std::chrono::steady_clock::now();
            for (FolderItem& item : payload->items)
            {
                if (stopToken.stop_requested() || _enumerationGeneration.load(std::memory_order_acquire) != generation)
                {
                    return nullptr;
                }
                if (! FileOperationArtifacts::HasPossibleArtifactName(item.displayName))
                {
                    continue;
                }
                ++artifactCandidateCount;
                FileOperationArtifacts::Projection projection{};
                const HRESULT projectHr = FileOperationArtifacts::ProjectProviderChildObject(fileSystem.get(),
                                                                                               folder.native(),
                                                                                               item.displayName,
                                                                                               pluginId,
                                                                                               instanceContext,
                                                                                               projection);
                if (projectHr == S_OK && projection.classification != FileOperationArtifacts::Classification::Ordinary)
                {
                    item.artifactProjection = std::make_shared<FileOperationArtifacts::Projection>(std::move(projection));
                    ++artifactProjectedCount;
                }
            }
            Debug::Perf::Emit(L"fileops.artifact.folder_projection.us",
                              L"name-shape-only",
                              Debug::Perf::ElapsedUs(artifactStartedAt),
                              static_cast<uint64_t>(artifactProjectedCount),
                              static_cast<uint64_t>(artifactCandidateCount),
                              S_OK);
            Debug::Perf::EmitValue(L"fileops.artifact.folder_projection.probe_candidates",
                                   static_cast<uint64_t>(artifactCandidateCount),
                                   S_OK);

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
                        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast(COINIT_MULTITHREADED);
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

                struct PerFileWork
                {
                    std::vector<size_t> itemIndices;
                    std::wstring fullPath;
                    DWORD fileAttributes                    = 0;
                    std::atomic<bool>* stopRequested         = nullptr;
                    std::atomic<uint64_t>* generationCounter = nullptr;
                    uint64_t generation                      = 0;
                    int resolvedIconIndex                    = -1;
                };

                std::atomic<bool> perFileStopRequested{false};
                std::vector<std::unique_ptr<PerFileWork>> perFileWorks;
                perFileWorks.reserve(perFileIconIndices.size());
                std::unordered_map<std::wstring, size_t> perFileWorkIndexByPath;
                perFileWorkIndexByPath.reserve(perFileIconIndices.size());

                uint64_t perFilePathChars = 0;
                {
                    Debug::Perf::Scope perFilePathsPerf(L"FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths");
                    perFilePathsPerf.SetDetail(folder.native());
                    perFilePathsPerf.SetValue0(perFileIconIndices.size());

                    for (size_t idx : perFileIconIndices)
                    {
                        std::wstring fullPath = (folder / payload->items[idx].displayName).wstring();
                        perFilePathChars += static_cast<uint64_t>(fullPath.size());

                        const auto [it, inserted] = perFileWorkIndexByPath.emplace(fullPath, perFileWorks.size());
                        if (! inserted)
                        {
                            perFileWorks[it->second]->itemIndices.push_back(idx);
                            continue;
                        }

                        auto work = std::make_unique<PerFileWork>();
                        work->itemIndices.push_back(idx);
                        work->fullPath          = std::move(fullPath);
                        work->fileAttributes    = payload->items[idx].fileAttributes;
                        work->stopRequested     = &perFileStopRequested;
                        work->generationCounter = &_enumerationGeneration;
                        work->generation        = generation;
                        perFileWorks.push_back(std::move(work));
                    }

                    perFilePathsPerf.SetValue1(perFileWorks.size());
                }

                std::vector<UniqueThreadpoolWork> perFileWorkItems;
                perFileWorkItems.reserve(perFileWorks.size());

                for (auto& work : perFileWorks)
                {
                    UniqueThreadpoolWork workItem(::CreateThreadpoolWork(
                        [](PTP_CALLBACK_INSTANCE, PVOID context, PTP_WORK) noexcept
                    {
                        auto* work = static_cast<PerFileWork*>(context);
                        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast(COINIT_MULTITHREADED);
                        if (work->stopRequested->load() ||
                            (work->generationCounter && work->generationCounter->load(std::memory_order_acquire) != work->generation))
                        {
                            return;
                        }

                        const auto iconIndex    = IconCache::GetInstance().QuerySysIconIndexForPath(work->fullPath.c_str(), work->fileAttributes, false);
                        work->resolvedIconIndex = iconIndex.value_or(-1);
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
                for (const auto& work : perFileWorks)
                {
                    if (work->resolvedIconIndex < 0)
                    {
                        perFileFailures += static_cast<uint64_t>(work->itemIndices.size());
                    }

                    for (const size_t idx : work->itemIndices)
                    {
                        payload->items[idx].iconIndex = work->resolvedIconIndex;
                    }
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
        _pendingEnumerationFileSystem.reset();
        _pendingEnumerationPluginId.clear();
        _pendingEnumerationInstanceContext.clear();
        _iconLoadQueue.clear();
        _thumbnailLoadQueue.clear();
        _iconLoadingActive.store(false, std::memory_order_release);
        _thumbnailLoadingActive.store(false, std::memory_order_release);
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
    DisarmPotentialDrag();

    // Stop idle layout pre-creation from previous folder
    if (_idleLayoutTimer != 0 && _hWnd)
    {
        KillTimer(_hWnd.get(), kIdleLayoutTimerId);
        _idleLayoutTimer = 0;
    }

    if (! _pendingNavigationDisplayedModel.has_value() && _displayedFolder.has_value())
    {
        _pendingNavigationDisplayedModel = PendingNavigationDisplayedModel{
            .displayedFolder = _displayedFolder,
            .displayedLocationKey = _displayedLocationKey,
            .items = std::move(_items),
            .itemsArenaBuffer = std::move(_itemsArenaBuffer),
            .itemsFolder = std::move(_itemsFolder),
            .selectionStats = _selectionStats,
            .focusedIndex = _focusedIndex,
            .hoveredIndex = _hoveredIndex,
            .anchorIndex = _anchorIndex,
            .resolutionReason = _lastCurrentResolutionReason,
            .scrollOffset = _scrollOffset,
            .horizontalOffset = _horizontalOffset,
        };
    }
    else
    {
        _items.clear();
        _itemsArenaBuffer.reset();
        _itemsFolder.clear();
    }
    _columnLayout.clear();
    _columnCounts.clear();
    _columnPrefixSums.clear();
    _scrollOffset      = 0.0f;
    _horizontalOffset  = 0.0f;
    _itemMetricsCached = false;
    _focusedIndex      = static_cast<size_t>(-1);
    _anchorIndex       = static_cast<size_t>(-1);
    _hoveredIndex      = static_cast<size_t>(-1);
    _selectionStats    = {};

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
    if (_pendingExternalSelectionAfterEnumeration && _currentFolder)
    {
        const std::wstring currentKey = NormalizeFocusMemoryFolderKey(_currentFolder.value());
        const std::wstring targetKey  = NormalizeFocusMemoryFolderKey(_pendingExternalSelectionAfterEnumeration->targetFolder);
        if (! currentKey.empty() && currentKey == targetKey)
        {
            _pendingExternalSelectionAfterEnumeration->generation = generation;
        }
        else
        {
            _pendingExternalSelectionAfterEnumeration.reset();
        }
    }
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath       = *_currentFolder;
        _pendingEnumerationGeneration = generation;
        _pendingEnumerationFileSystem = _fileSystem;
        _pendingEnumerationPluginId = _fileSystemPluginId;
        _pendingEnumerationInstanceContext = _fileSystemInstanceContext;
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
        const ULONGLONG remaining = kDebounceMs - (now - _lastDirectoryCacheRefreshTick);
        const UINT intervalMs     = static_cast<UINT>(std::clamp<ULONGLONG>(remaining, 1u, kDebounceMs));
        const UINT_PTR timer      = SetTimer(_hWnd.get(), kDirectoryCacheRefreshTimerId, intervalMs, nullptr);
        if (timer != 0)
        {
            _directoryCacheRefreshTimer = timer;
            _pendingRefreshDebounceDelayMs =
                (std::max)(_pendingRefreshDebounceDelayMs, static_cast<uint64_t>(intervalMs));
            return;
        }
    }

    if (_directoryCacheRefreshTimer != 0)
    {
        KillTimer(_hWnd.get(), kDirectoryCacheRefreshTimerId);
        _directoryCacheRefreshTimer = 0;
    }

    _lastDirectoryCacheRefreshTick = now;
    _pendingRefreshDebounceDelayMs = 0u;
    RequestRefreshFromCache();
}

void FolderView::OnDirectoryImpact(std::unique_ptr<DirectoryInfoCache::DirectoryImpact> impact)
{
    if (! impact)
    {
        return;
    }

    switch (impact->kind)
    {
        case DirectoryInfoCache::DirectoryImpact::Kind::RefreshCurrentFolder:
            if (! impact->renamedFromDisplayName.empty() && ! impact->renamedToDisplayName.empty())
            {
                const auto now       = std::chrono::steady_clock::now();
                const auto expiresAt = now + std::chrono::seconds{30};
                std::erase_if(_recentlyMissingRefreshSelections,
                              [&](const RecentlyMissingRefreshSelection& selection) noexcept { return selection.expiresAt <= now; });

                bool fromWasSelected = false;
                for (const auto& item : _items)
                {
                    if (item.selected && EquivalentProviderComponent(item.displayName, impact->renamedFromDisplayName))
                    {
                        fromWasSelected = true;
                        break;
                    }
                }

                if (! fromWasSelected)
                {
                    for (const auto& rename : _pendingRefreshSelectionRenames)
                    {
                        if (rename.fromWasSelected && EquivalentProviderComponent(rename.toDisplayName, impact->renamedFromDisplayName))
                        {
                            fromWasSelected = true;
                            break;
                        }
                    }
                }

                if (! fromWasSelected)
                {
                    const auto missingIt = std::find_if(_recentlyMissingRefreshSelections.begin(),
                                                        _recentlyMissingRefreshSelections.end(),
                                                        [&](const RecentlyMissingRefreshSelection& selection) noexcept
                    { return EquivalentProviderComponent(selection.displayName, impact->renamedFromDisplayName); });
                    if (missingIt != _recentlyMissingRefreshSelections.end())
                    {
                        fromWasSelected = true;
                        _recentlyMissingRefreshSelections.erase(missingIt);
                    }
                }

                bool collapsedChain  = false;
                for (auto& rename : _pendingRefreshSelectionRenames)
                {
                    if (EquivalentProviderComponent(rename.toDisplayName, impact->renamedFromDisplayName))
                    {
                        rename.toDisplayName.assign(impact->renamedToDisplayName);
                        rename.fromWasSelected = rename.fromWasSelected || fromWasSelected;
                        rename.expiresAt = expiresAt;
                        collapsedChain = true;
                    }
                }

                if (! collapsedChain)
                {
                    _pendingRefreshSelectionRenames.push_back({.fromDisplayName = impact->renamedFromDisplayName,
                                                               .toDisplayName = impact->renamedToDisplayName,
                                                               .fromWasSelected = fromWasSelected,
                                                               .expiresAt = expiresAt});
                }
            }
            OnDirectoryCacheDirty();
            return;
        case DirectoryInfoCache::DirectoryImpact::Kind::RelocateCurrentFolder:
        case DirectoryInfoCache::DirectoryImpact::Kind::RetargetInstanceContext:
        case DirectoryInfoCache::DirectoryImpact::Kind::ExitInstanceContext:
            if (_directoryImpactCallback)
            {
                _directoryImpactCallback(*impact);
                return;
            }

            if (impact->kind == DirectoryInfoCache::DirectoryImpact::Kind::RelocateCurrentFolder && ! impact->targetFolder.empty())
            {
                if (! impact->focusDisplayName.empty())
                {
                    RememberFocusedItemForFolder(impact->targetFolder, impact->focusDisplayName);
                }
                SetFolderPath(impact->targetFolder);
            }
            return;
    }
}

void FolderView::RequestRefreshFromCache()
{
    if (! _currentFolder || ! _hWnd)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::RequestRefreshFromCache: skipped (no current folder or hwnd)");
#endif
        return;
    }

    EnsureEnumerationThread();
    const uint64_t generation = _enumerationGeneration.fetch_add(1, std::memory_order_release) + 1;
    const uint64_t debounceDelayMs = _pendingRefreshDebounceDelayMs;
    _pendingRefreshDebounceDelayMs = 0u;
    RecordPendingRefreshToPaintStart(generation, debounceDelayMs);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::RequestRefreshFromCache: generation={} folder='{}'", generation, _currentFolder->native()));
#endif
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
    if (_pendingExternalSelectionAfterEnumeration && _currentFolder)
    {
        const std::wstring currentKey = NormalizeFocusMemoryFolderKey(_currentFolder.value());
        const std::wstring targetKey  = NormalizeFocusMemoryFolderKey(_pendingExternalSelectionAfterEnumeration->targetFolder);
        if (! currentKey.empty() && currentKey == targetKey)
        {
            _pendingExternalSelectionAfterEnumeration->generation = generation;
        }
        else
        {
            _pendingExternalSelectionAfterEnumeration.reset();
        }
    }
    {
        std::lock_guard guard(_enumerationMutex);
        _pendingEnumerationPath       = *_currentFolder;
        _pendingEnumerationGeneration = generation;
        _pendingEnumerationFileSystem = _fileSystem;
        _pendingEnumerationPluginId = _fileSystemPluginId;
        _pendingEnumerationInstanceContext = _fileSystemInstanceContext;
    }
    _enumerationCv.notify_one();
}

void FolderView::ApplyCurrentSort()
{
    DisarmPotentialDrag();
    CurrentResolutionInput input{};
    input.sameLogicalLocation = true;
    input.previousSelectionStats = _selectionStats;
    if (_focusedIndex < _items.size())
    {
        const FolderItem& current = _items[_focusedIndex];
        input.survivingIdentity = std::wstring(current.displayName);
        input.previousCurrentIdentity = input.survivingIdentity;
        input.genericRepairProbe = CurrentRepairProbe{
            .displayName = std::wstring(current.displayName),
            .extensionOffset = current.extensionOffset,
            .isDirectory = current.isDirectory,
            .sizeBytes = current.sizeBytes,
            .lastWriteTime = current.lastWriteTime,
            .fileAttributes = current.fileAttributes,
            .unsortedOrder = current.unsortedOrder,
        };
    }
    if (_anchorIndex < _items.size())
    {
        input.previousAnchorIdentity = std::wstring(_items[_anchorIndex].displayName);
    }
    ApplyCurrentSort(std::move(input));
}

std::wstring FolderView::BuildLogicalLocationKey(const std::filesystem::path& folder) const noexcept
{
    std::wstring key;
    const auto appendSegment = [&](std::wstring_view segment)
    {
        key.append(std::to_wstring(segment.size()));
        key.push_back(L':');
        key.append(segment);
        key.push_back(L'|');
    };

    appendSegment(_fileSystemPluginId);
    appendSegment(_fileSystemInstanceContext);

    if (_fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity)
    {
        const std::optional<std::wstring> pathKey = TryMakePathKey(_fileSystemPathIdentity.value(), folder.native());
        if (pathKey.has_value())
        {
            key.append(L"stable|");
            appendSegment(pathKey.value());
            return key;
        }
    }

    key.append(L"live|");
    // Unstable path text is valid only for this live provider object. Context changes on the
    // same object remain independently addressable and restorable through the context segment.
    key.append(std::to_wstring(_fileSystemLiveInstanceEpoch));
    key.push_back(L'|');
    appendSegment(folder.native());
    return key;
}

bool FolderView::EquivalentProviderComponent(std::wstring_view left, std::wstring_view right) const noexcept
{
    if (_fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity)
    {
        return EquivalentComponent(_fileSystemPathIdentity.value(), left, right);
    }
    return left == right;
}

std::optional<size_t> FolderView::FindItemByProviderIdentity(std::wstring_view displayName) const noexcept
{
    if (displayName.empty())
    {
        return std::nullopt;
    }
    for (size_t index = 0u; index < _items.size(); ++index)
    {
        if (EquivalentProviderComponent(_items[index].displayName, displayName))
        {
            return index;
        }
    }
    return std::nullopt;
}

bool FolderView::ItemOrdersBeforeForCurrentSort(const FolderItem& left, const FolderItem& right) const noexcept
{
    const auto compareInt = [&](const int comparison) noexcept
    {
        return _sortDirection == SortDirection::Ascending ? comparison < 0 : comparison > 0;
    };

    const auto compareName = [&](const FolderItem& lhs, const FolderItem& rhs) noexcept
    {
        const int comparison = OrdinalString::Compare(lhs.displayName, rhs.displayName, true);
        if (comparison != 0)
        {
            return compareInt(comparison);
        }

        const int caseComparison = OrdinalString::Compare(lhs.displayName, rhs.displayName, false);
        if (caseComparison != 0)
        {
            return compareInt(caseComparison);
        }

        return lhs.unsortedOrder < rhs.unsortedOrder;
    };

    if (left.isDirectory != right.isDirectory)
    {
        return left.isDirectory && ! right.isDirectory;
    }

    switch (_sortBy)
    {
        case SortBy::Name: return compareName(left, right);
        case SortBy::Extension:
        {
            const int extensionComparison = OrdinalString::Compare(left.GetExtension(), right.GetExtension(), true);
            return extensionComparison != 0 ? compareInt(extensionComparison) : compareName(left, right);
        }
        case SortBy::Time:
            if (left.lastWriteTime != right.lastWriteTime)
            {
                return _sortDirection == SortDirection::Ascending ? left.lastWriteTime < right.lastWriteTime
                                                                  : left.lastWriteTime > right.lastWriteTime;
            }
            return compareName(left, right);
        case SortBy::Size:
            if (! left.isDirectory && ! right.isDirectory && left.sizeBytes != right.sizeBytes)
            {
                return _sortDirection == SortDirection::Ascending ? left.sizeBytes < right.sizeBytes : left.sizeBytes > right.sizeBytes;
            }
            return compareName(left, right);
        case SortBy::Attributes:
            if (left.fileAttributes != right.fileAttributes)
            {
                return _sortDirection == SortDirection::Ascending ? left.fileAttributes < right.fileAttributes
                                                                  : left.fileAttributes > right.fileAttributes;
            }
            return compareName(left, right);
        case SortBy::None: return left.unsortedOrder < right.unsortedOrder;
    }

    return compareName(left, right);
}

FolderView::CurrentResolutionResult FolderView::ResolveCurrentAfterSort(const CurrentResolutionInput& input) const noexcept
{
    if (_items.empty())
    {
        return {.reason = CurrentResolutionReason::Empty};
    }

    const auto resolveIdentity = [&](const std::optional<std::wstring>& identity,
                                     const CurrentResolutionReason reason) noexcept -> std::optional<CurrentResolutionResult>
    {
        if (! identity.has_value())
        {
            return std::nullopt;
        }
        if (const std::optional<size_t> index = FindItemByProviderIdentity(identity.value()); index.has_value())
        {
            return CurrentResolutionResult{.index = index.value(), .reason = reason};
        }
        return std::nullopt;
    };

    if (const auto explicitTarget = resolveIdentity(input.explicitHostTarget, CurrentResolutionReason::ExplicitHostTarget);
        explicitTarget.has_value())
    {
        return explicitTarget.value();
    }

    const auto resolveRepairProbe = [&](const CurrentRepairProbe& probe,
                                        const CurrentResolutionReason successorReason,
                                        const CurrentResolutionReason predecessorReason) noexcept
        -> std::optional<CurrentResolutionResult>
    {
        const auto isExcluded = [&](std::wstring_view identity) noexcept
        {
            return std::ranges::any_of(probe.excludedIdentities, [&](const std::wstring& excluded) noexcept
            { return EquivalentProviderComponent(identity, excluded); });
        };

        if (_sortBy == SortBy::None)
        {
            if (probe.sortNoneSuccessor.has_value() && ! isExcluded(probe.sortNoneSuccessor.value()))
            {
                if (const auto successor = resolveIdentity(probe.sortNoneSuccessor, successorReason); successor.has_value())
                {
                    return successor.value();
                }
            }
            if (probe.sortNonePredecessor.has_value() && ! isExcluded(probe.sortNonePredecessor.value()))
            {
                if (const auto predecessor = resolveIdentity(probe.sortNonePredecessor, predecessorReason); predecessor.has_value())
                {
                    return predecessor.value();
                }
            }
            return std::nullopt;
        }

        FolderItem probeItem{};
        probeItem.displayName = probe.displayName;
        probeItem.extensionOffset = probe.extensionOffset;
        probeItem.isDirectory = probe.isDirectory;
        probeItem.sizeBytes = probe.sizeBytes;
        probeItem.lastWriteTime = probe.lastWriteTime;
        probeItem.fileAttributes = probe.fileAttributes;
        probeItem.unsortedOrder = probe.unsortedOrder;

        const auto insertion = std::lower_bound(_items.begin(), _items.end(), probeItem,
                                                [&](const FolderItem& item, const FolderItem& value) noexcept
        { return ItemOrdersBeforeForCurrentSort(item, value); });
        for (auto successor = insertion; successor != _items.end(); ++successor)
        {
            if (! isExcluded(successor->displayName))
            {
                return CurrentResolutionResult{
                    .index = static_cast<size_t>(successor - _items.begin()),
                    .reason = successorReason,
                };
            }
        }
        for (auto predecessor = insertion; predecessor != _items.begin();)
        {
            --predecessor;
            if (! isExcluded(predecessor->displayName))
            {
                return CurrentResolutionResult{
                    .index = static_cast<size_t>(predecessor - _items.begin()),
                    .reason = predecessorReason,
                };
            }
        }
        return std::nullopt;
    };

    if (input.genericRepairProbe.has_value() && input.genericRepairProbe->hostRemovalRepair)
    {
        if (const auto hostRepair = resolveRepairProbe(input.genericRepairProbe.value(),
                                                       CurrentResolutionReason::HostRemovalIntent,
                                                       CurrentResolutionReason::HostRemovalIntent);
            hostRepair.has_value())
        {
            return hostRepair.value();
        }
    }

    if (const auto surviving = resolveIdentity(input.survivingIdentity, CurrentResolutionReason::SurvivingIdentity); surviving.has_value())
    {
        return surviving.value();
    }

    if (input.genericRepairProbe.has_value() && ! input.genericRepairProbe->hostRemovalRepair)
    {
        if (const auto genericRepair = resolveRepairProbe(input.genericRepairProbe.value(),
                                                          CurrentResolutionReason::GenericSuccessor,
                                                          CurrentResolutionReason::GenericPredecessor);
            genericRepair.has_value())
        {
            return genericRepair.value();
        }
    }

    if (const auto restored = resolveIdentity(input.navigationMemoryRestore, CurrentResolutionReason::NavigationMemoryRestore);
        restored.has_value())
    {
        return restored.value();
    }

    return {.index = 0u, .reason = CurrentResolutionReason::FirstItem};
}

void FolderView::ApplyCurrentSort(CurrentResolutionInput resolutionInput)
{
    constexpr auto invalidIndex = static_cast<size_t>(-1);
    Debug::Perf::Scope focusResolvePerf(L"folder.focus.resolve_us");
    focusResolvePerf.SetValue0(static_cast<uint64_t>(_items.size()));

    if (_items.empty())
    {
        const bool focusChanged     = _focusedIndex != invalidIndex || resolutionInput.previousCurrentIdentity.has_value();
        const bool selectionChanged = resolutionInput.selectionMembershipChanged ||
            resolutionInput.previousSelectionStats.selectedFolders != 0u || resolutionInput.previousSelectionStats.selectedFiles != 0u ||
            resolutionInput.previousSelectionStats.singleItem.has_value();
        _focusedIndex               = invalidIndex;
        _anchorIndex                = invalidIndex;
        _hoveredIndex               = invalidIndex;
        _selectionStats             = {};
        _lastCurrentResolutionReason = CurrentResolutionReason::Empty;
        focusResolvePerf.SetValue1(static_cast<uint64_t>(_lastCurrentResolutionReason));
        if (selectionChanged)
        {
            NotifySelectionChanged();
        }
        if (focusChanged)
        {
            NotifyFocusedItemChanged();
        }
        return;
    }

    Debug::Perf::Scope perf(L"FolderView.ApplyCurrentSort");
    perf.SetValue0(_items.size());

    const auto compare = [&](const FolderItem& left, const FolderItem& right) noexcept
    { return ItemOrdersBeforeForCurrentSort(left, right); };

    // Parallel sorting has enough scheduling overhead that medium interactive folders are faster
    // on the sequential path; keep parallelism for genuinely large directories.
    if (FolderViewSortPolicy::ShouldUseParallelSort(_items.size()))
    {
        std::sort(std::execution::par, _items.begin(), _items.end(), compare);
    }
    else
    {
        std::sort(_items.begin(), _items.end(), compare);
    }

    SelectionStats stats{};
    const FolderItem* singleSelected = nullptr;
    uint32_t selectedTotal           = 0;
    for (size_t i = 0; i < _items.size(); ++i)
    {
        auto& item    = _items[i];
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

    }

    const CurrentResolutionResult resolution = ResolveCurrentAfterSort(resolutionInput);
    _focusedIndex = resolution.index;
    _anchorIndex = _focusedIndex;
    if (resolutionInput.sameLogicalLocation && resolutionInput.previousAnchorIdentity.has_value())
    {
        if (const std::optional<size_t> survivingAnchor = FindItemByProviderIdentity(resolutionInput.previousAnchorIdentity.value());
            survivingAnchor.has_value())
        {
            _anchorIndex = survivingAnchor.value();
        }
    }

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

    const auto selectedDetailsEqual = [](const std::optional<SelectionStats::SelectedItemDetails>& left,
                                         const std::optional<SelectionStats::SelectedItemDetails>& right) noexcept
    {
        if (left.has_value() != right.has_value())
        {
            return false;
        }
        if (! left.has_value())
        {
            return true;
        }
        return left->isDirectory == right->isDirectory && left->sizeBytes == right->sizeBytes &&
            left->lastWriteTime == right->lastWriteTime && left->fileAttributes == right->fileAttributes;
    };
    const bool selectionMembershipChanged = resolutionInput.selectionMembershipChanged ||
        resolutionInput.previousSelectionStats.selectedFiles != stats.selectedFiles ||
        resolutionInput.previousSelectionStats.selectedFolders != stats.selectedFolders;
    const bool selectionChanged = selectionMembershipChanged ||
        resolutionInput.previousSelectionStats.selectedFileBytes != stats.selectedFileBytes ||
        ! selectedDetailsEqual(resolutionInput.previousSelectionStats.singleItem, stats.singleItem);

    _hoveredIndex   = static_cast<size_t>(-1);
    _selectionStats = stats;
    _lastCurrentResolutionReason = resolution.reason;
    focusResolvePerf.SetValue1(static_cast<uint64_t>(_lastCurrentResolutionReason));
    if (selectionChanged)
    {
        NotifySelectionChanged(selectionMembershipChanged ? SelectionChangeKind::Membership : SelectionChangeKind::ItemMetadata);
    }
    bool focusChanged = ! resolutionInput.sameLogicalLocation || ! resolutionInput.previousCurrentIdentity.has_value();
    if (! focusChanged && _focusedIndex < _items.size())
    {
        focusChanged = ! EquivalentProviderComponent(resolutionInput.previousCurrentIdentity.value(), _items[_focusedIndex].displayName);
    }
    if (focusChanged)
    {
        NotifyFocusedItemChanged();
    }
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

    static_cast<void>(RecordFocusMemoryEntry(_displayedFolder.value(), _items[_focusedIndex].displayName));
}

void FolderView::RememberFocusedItemForFolder(const std::filesystem::path& folder, std::wstring_view itemDisplayName) noexcept
{
    if (itemDisplayName.empty())
    {
        return;
    }

    const std::wstring locationKey = BuildLogicalLocationKey(folder);
    if (! locationKey.empty())
    {
        _pendingExplicitCurrentTarget = PendingExplicitCurrentTarget{
            .locationKey = locationKey,
            .displayName = std::wstring(itemDisplayName),
        };
    }

    static_cast<void>(RecordFocusMemoryEntry(folder, itemDisplayName));
}

uint64_t FolderView::BeginRemovalFocusTracking(const std::vector<std::filesystem::path>& removedPaths,
                                               const FileSystemPathIdentity& pathIdentity)
{
    constexpr auto invalidIndex = static_cast<size_t>(-1);

    if (! pathIdentity.pathTextStableIdentity || ! _displayedFolder || ! _currentFolder ||
        ! EquivalentPath(pathIdentity, _displayedFolder->native(), _currentFolder->native()) ||
        _focusedIndex == invalidIndex || _focusedIndex >= _items.size() || removedPaths.empty())
    {
        return 0u;
    }

    PendingRemovalFocus entry{};
    entry.folderPath = _displayedFolder.value();
    entry.pathIdentity = pathIdentity;
    entry.removedFocusDisplayName.assign(_items[_focusedIndex].displayName);
    entry.folderPathGeneration = _folderPathGeneration;
    entry.focusOwnershipEpoch = _removalFocusOwnershipEpoch;
    entry.sortEpoch = _removalFocusSortEpoch;
    entry.providerEpoch = _removalFocusProviderEpoch;

    Debug::Perf::Scope beginPerf(L"folder.removal_focus.begin_us");
    beginPerf.SetValue0(static_cast<uint64_t>(_items.size()));
    beginPerf.SetValue1(static_cast<uint64_t>(removedPaths.size()));

    bool focusedSourceFound = false;
    entry.sources.reserve(removedPaths.size());
    for (size_t sourceIndex = 0u; sourceIndex < removedPaths.size(); ++sourceIndex)
    {
        const std::filesystem::path& path = removedPaths[sourceIndex];
        if (! EquivalentPath(pathIdentity, path.parent_path().native(), _itemsFolder.native()))
        {
            continue;
        }

        PendingRemovalFocus::Source source{.sourceIndex = sourceIndex, .displayName = path.filename().wstring()};
        if (EquivalentComponent(pathIdentity, source.displayName, entry.removedFocusDisplayName))
        {
            entry.focusedSourceIndex = sourceIndex;
            focusedSourceFound = true;
        }
        entry.sources.push_back(std::move(source));
    }

    if (! focusedSourceFound)
    {
        return 0u;
    }

    uint64_t token = _nextRemovalFocusToken++;
    if (token == 0u)
    {
        token = _nextRemovalFocusToken++;
    }
    entry.token = token;
    _pendingRemovalFocus.push_back(std::move(entry));
    return token;
}

void FolderView::CompleteRemovalFocusTracking(uint64_t token, std::span<const RemovalDisposition> sourceDispositions) noexcept
{
    if (token == 0u)
    {
        return;
    }

    const auto pending = std::find_if(_pendingRemovalFocus.begin(), _pendingRemovalFocus.end(),
                                      [token](const PendingRemovalFocus& entry) noexcept { return entry.token == token; });
    if (pending == _pendingRemovalFocus.end())
    {
        return;
    }

    if (pending->folderPathGeneration != _folderPathGeneration || pending->focusOwnershipEpoch != _removalFocusOwnershipEpoch ||
        pending->sortEpoch != _removalFocusSortEpoch || pending->providerEpoch != _removalFocusProviderEpoch ||
        pending->focusedSourceIndex >= sourceDispositions.size() ||
        sourceDispositions[pending->focusedSourceIndex] != RemovalDisposition::Removed)
    {
        _pendingRemovalFocus.erase(pending);
        return;
    }

    for (PendingRemovalFocus::Source& source : pending->sources)
    {
        source.provenRemoved = source.sourceIndex < sourceDispositions.size() &&
                               sourceDispositions[source.sourceIndex] == RemovalDisposition::Removed;
    }

    pending->committed = true;
    pending->commitEnumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    pending->commitSequence = _nextRemovalFocusCommitSequence++;
    if (pending->commitSequence == 0u)
    {
        pending->commitSequence = _nextRemovalFocusCommitSequence++;
    }
    pending->expiresAt = std::chrono::steady_clock::now() + std::chrono::seconds{30};
}

std::optional<std::vector<std::wstring>> FolderView::ResolvePendingRemovalFocusForEnumeration(
    uint64_t generation,
    const std::filesystem::path& currentFolder,
    const std::vector<FolderItem>& items)
{
    const auto now = std::chrono::steady_clock::now();
    std::erase_if(_pendingRemovalFocus,
                  [&](const PendingRemovalFocus& pending) noexcept
    {
        return (pending.committed && pending.expiresAt <= now) || pending.folderPathGeneration != _folderPathGeneration ||
               pending.focusOwnershipEpoch != _removalFocusOwnershipEpoch || pending.sortEpoch != _removalFocusSortEpoch ||
               pending.providerEpoch != _removalFocusProviderEpoch ||
               ! EquivalentPath(pending.pathIdentity, pending.folderPath.native(), currentFolder.native());
    });

    uint64_t newestReflectedSequence = 0u;
    std::optional<std::vector<std::wstring>> newestProvenRemovedIdentities;
    for (auto pending = _pendingRemovalFocus.begin(); pending != _pendingRemovalFocus.end();)
    {
        if (! pending->committed || generation <= pending->commitEnumerationGeneration)
        {
            ++pending;
            continue;
        }

        const bool removedFocusSurvived = std::ranges::any_of(items, [&](const FolderItem& item) noexcept
        { return EquivalentComponent(pending->pathIdentity, item.displayName, pending->removedFocusDisplayName); });
        if (removedFocusSurvived)
        {
            ++pending;
            continue;
        }

        std::vector<std::wstring> provenRemovedIdentities;
        provenRemovedIdentities.reserve(pending->sources.size());
        for (const PendingRemovalFocus::Source& source : pending->sources)
        {
            if (source.provenRemoved)
            {
                provenRemovedIdentities.push_back(source.displayName);
            }
        }
        if (pending->commitSequence >= newestReflectedSequence)
        {
            newestReflectedSequence = pending->commitSequence;
            newestProvenRemovedIdentities = std::move(provenRemovedIdentities);
        }
        pending = _pendingRemovalFocus.erase(pending);
    }
    return newestProvenRemovedIdentities;
}

#if defined(ENABLE_TESTS)
std::wstring FolderView::DebugValidateRemovalFocusContractsForSelfTest()
{
    const std::filesystem::path folder = LR"(C:\focus-contract)";
    const FileSystemPathIdentity insensitive = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    FileSystemPathIdentity sensitive = insensitive;
    sensitive.componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive;

    const auto makeItems = [](const std::vector<std::wstring>& names)
    {
        std::vector<FolderItem> items(names.size());
        for (size_t index = 0u; index < names.size(); ++index)
        {
            items[index].displayName = names[index];
            items[index].unsortedOrder = index;
        }
        return items;
    };

    struct ResolutionObservation final
    {
        bool intentAccepted = false;
        std::wstring currentDisplayName;
        CurrentResolutionReason reason = CurrentResolutionReason::Unresolved;
        std::vector<std::wstring> provenRemovedIdentities;
    };

    const auto resolveExact = [&](const std::vector<std::wstring>& ordered,
                                  size_t focusedOrderedIndex,
                                  const std::vector<std::wstring>& requested,
                                  std::span<const RemovalDisposition> dispositions,
                                  const FileSystemPathIdentity& identity,
                                  const std::vector<std::wstring>& refreshed) -> ResolutionObservation
    {
        FolderView view;
        view._fileSystemPathIdentity = identity;
        view._items = makeItems(ordered);
        view._focusedIndex = focusedOrderedIndex;

        PendingRemovalFocus pending{};
        pending.token = 1u;
        pending.folderPath = folder;
        pending.pathIdentity = identity;
        pending.removedFocusDisplayName = ordered[focusedOrderedIndex];
        for (size_t sourceIndex = 0u; sourceIndex < requested.size(); ++sourceIndex)
        {
            pending.sources.push_back(PendingRemovalFocus::Source{
                .sourceIndex = sourceIndex,
                .displayName = requested[sourceIndex],
            });
            if (EquivalentComponent(identity, requested[sourceIndex], pending.removedFocusDisplayName))
            {
                pending.focusedSourceIndex = sourceIndex;
            }
        }
        view._pendingRemovalFocus.push_back(std::move(pending));
        view.CompleteRemovalFocusTracking(1u, dispositions);
        if (view._pendingRemovalFocus.empty() || ! view._pendingRemovalFocus.front().committed)
        {
            return {};
        }

        const uint64_t reflectedGeneration = view._pendingRemovalFocus.front().commitEnumerationGeneration + 1u;
        std::optional<std::vector<std::wstring>> provenRemoved =
            view.ResolvePendingRemovalFocusForEnumeration(reflectedGeneration, folder, makeItems(refreshed));
        if (! provenRemoved.has_value())
        {
            return {};
        }

        const FolderItem& priorCurrent = view._items[focusedOrderedIndex];
        CurrentResolutionInput input{};
        input.sameLogicalLocation = true;
        input.genericRepairProbe = CurrentRepairProbe{
            .displayName = std::wstring(priorCurrent.displayName),
            .extensionOffset = priorCurrent.extensionOffset,
            .isDirectory = priorCurrent.isDirectory,
            .sizeBytes = priorCurrent.sizeBytes,
            .lastWriteTime = priorCurrent.lastWriteTime,
            .fileAttributes = priorCurrent.fileAttributes,
            .unsortedOrder = priorCurrent.unsortedOrder,
            .excludedIdentities = std::move(provenRemoved.value()),
            .hostRemovalRepair = true,
        };

        view._items = makeItems(refreshed);
        std::stable_sort(view._items.begin(), view._items.end(), [&](const FolderItem& left, const FolderItem& right) noexcept
        { return view.ItemOrdersBeforeForCurrentSort(left, right); });
        const CurrentResolutionResult result = view.ResolveCurrentAfterSort(input);
        return ResolutionObservation{
            .intentAccepted = true,
            .currentDisplayName = result.index < view._items.size() ? std::wstring(view._items[result.index].displayName) : std::wstring{},
            .reason = result.reason,
            .provenRemovedIdentities = std::move(input.genericRepairProbe.value().excludedIdentities),
        };
    };

    const std::vector<std::wstring> ordered{L"A", L"B", L"C", L"D"};
    const std::vector<std::wstring> requested{L"B", L"C"};
    const std::array<RemovalDisposition, 2> focusedSuccessOtherFailure{
        RemovalDisposition::Removed, RemovalDisposition::Retained};
    const ResolutionObservation retainedFailure =
        resolveExact(ordered, 1u, requested, focusedSuccessOtherFailure, insensitive, {L"D", L"A", L"C"});
    if (! retainedFailure.intentAccepted || retainedFailure.currentDisplayName != L"C" ||
        retainedFailure.reason != CurrentResolutionReason::HostRemovalIntent)
    {
        return L"focused Removed/other Retained did not keep survivor C under resulting keyed order";
    }
    const std::array<RemovalDisposition, 2> focusedFailureOtherSuccess{
        RemovalDisposition::Retained, RemovalDisposition::Removed};
    if (resolveExact(ordered, 1u, requested, focusedFailureOtherSuccess, insensitive, {L"D", L"A", L"B"}).intentAccepted)
    {
        return L"focused Retained/other Removed incorrectly committed focus migration";
    }
    const std::array<RemovalDisposition, 2> falseOutcome{
        RemovalDisposition::Retained, RemovalDisposition::Removed};
    if (resolveExact(ordered, 1u, requested, falseOutcome, insensitive, {L"D", L"A", L"B"}).intentAccepted)
    {
        return L"a Retained disposition incorrectly proved removal of the focused source";
    }
    const std::array<RemovalDisposition, 1> missingFocusedOutcome{RemovalDisposition::Retained};
    if (resolveExact(ordered, 1u, requested, missingFocusedOutcome, insensitive, {L"D", L"A", L"B"}).intentAccepted)
    {
        return L"a missing focused-source outcome did not fail closed";
    }
    const std::array<RemovalDisposition, 2> canceledOutcome{
        RemovalDisposition::Retained, RemovalDisposition::Removed};
    if (resolveExact(ordered, 1u, requested, canceledOutcome, insensitive, {L"D", L"A", L"B"}).intentAccepted)
    {
        return L"cancellation incorrectly committed focus migration";
    }
    const std::array<RemovalDisposition, 2> allSuccess{
        RemovalDisposition::Removed, RemovalDisposition::Removed};
    const ResolutionObservation allRemoved = resolveExact(ordered, 1u, requested, allSuccess, insensitive, {L"D", L"A"});
    if (! allRemoved.intentAccepted || allRemoved.currentDisplayName != L"D" ||
        allRemoved.reason != CurrentResolutionReason::HostRemovalIntent)
    {
        return L"all-Removed completion did not select canonical successor D from unsorted enumeration input";
    }
    const ResolutionObservation removedBeforeCurrent = resolveExact({L"A", L"B", L"C", L"D", L"E"},
                                                                     2u,
                                                                     {L"A", L"C"},
                                                                     allSuccess,
                                                                     insensitive,
                                                                     {L"E", L"B", L"D"});
    if (! removedBeforeCurrent.intentAccepted || removedBeforeCurrent.currentDisplayName != L"D")
    {
        return L"successful removals before current did not use the current probe's canonical successor D";
    }
    const std::vector<std::wstring> caseOrdered{L"A", L"README", L"Readme", L"Z"};
    const std::vector<std::wstring> caseRequested{L"README"};
    const std::array<RemovalDisposition, 1> caseSuccess{RemovalDisposition::Removed};
    const ResolutionObservation caseSensitive =
        resolveExact(caseOrdered, 1u, caseRequested, caseSuccess, sensitive, {L"Z", L"Readme", L"A"});
    if (! caseSensitive.intentAccepted || caseSensitive.currentDisplayName != L"Readme")
    {
        return L"case-sensitive provider identity conflated case-only siblings";
    }
    const std::array<RemovalDisposition, 1> focusedOnlySuccess{RemovalDisposition::Removed};
    const ResolutionObservation fartherSuccessor =
        resolveExact(ordered, 1u, {L"B"}, focusedOnlySuccess, insensitive, {L"D", L"A"});
    if (! fartherSuccessor.intentAccepted || fartherSuccessor.currentDisplayName != L"D")
    {
        return L"a disappeared pre-delete neighbor hid a farther refreshed survivor";
    }
    const ResolutionObservation skipProvenRemoved = resolveExact(ordered, 1u, requested, allSuccess, insensitive, {L"C", L"D", L"A"});
    if (! skipProvenRemoved.intentAccepted || skipProvenRemoved.currentDisplayName != L"D" ||
        skipProvenRemoved.reason != CurrentResolutionReason::HostRemovalIntent)
    {
        return L"canonical removal repair did not skip a proven-removed identity retained by a stale payload";
    }
    const ResolutionObservation onlyProvenRemoved = resolveExact(ordered, 1u, requested, allSuccess, insensitive, {L"C"});
    if (! onlyProvenRemoved.intentAccepted || onlyProvenRemoved.currentDisplayName != L"C" ||
        onlyProvenRemoved.reason != CurrentResolutionReason::FirstItem)
    {
        return L"specialized removal intent nominated a proven-removed fallback instead of using the invariant fallback";
    }

    {
        FolderView view;
        constexpr size_t kBeginSnapshotCount = 4096u;
        std::vector<std::wstring> names;
        names.reserve(kBeginSnapshotCount);
        for (size_t index = 0u; index < kBeginSnapshotCount; ++index)
        {
            names.push_back(std::format(L"removal-focus-snapshot-item-{:04}", index));
        }
        view._displayedFolder = folder;
        view._currentFolder = folder;
        view._itemsFolder = folder;
        view._items = makeItems(names);
        view._focusedIndex = 100u;
        const uint64_t token = view.BeginRemovalFocusTracking({folder / names[100]}, insensitive);
        if (token == 0u || view._pendingRemovalFocus.empty())
        {
            return L"BeginRemovalFocusTracking failed on a large synthetic folder";
        }
        const PendingRemovalFocus& pending = view._pendingRemovalFocus.front();
        if (pending.removedFocusDisplayName != names[100] || pending.sources.size() != 1u ||
            pending.sources.front().displayName != names[100])
        {
            return L"BeginRemovalFocusTracking did not retain only the focused source proof snapshot";
        }
        const std::array<RemovalDisposition, 1> beginSuccess{RemovalDisposition::Removed};
        view.CompleteRemovalFocusTracking(token, beginSuccess);
        if (view._pendingRemovalFocus.empty() || ! view._pendingRemovalFocus.front().committed)
        {
            return L"sources-only Begin snapshot did not commit after explicit focused Removed";
        }
        const uint64_t reflectedGeneration = view._pendingRemovalFocus.front().commitEnumerationGeneration + 1u;
        std::vector<std::wstring> refreshed = names;
        refreshed.erase(refreshed.begin() + 100);
        const std::optional<std::vector<std::wstring>> provenRemoved =
            view.ResolvePendingRemovalFocusForEnumeration(reflectedGeneration, folder, makeItems(refreshed));
        if (! provenRemoved.has_value() || provenRemoved->size() != 1u || provenRemoved.value()[0] != names[100])
        {
            return L"sources-only Begin snapshot did not produce the focused exact-removal proof";
        }
    }

    const auto addCommitted = [&](FolderView& view,
                                  std::wstring removedFocusDisplayName,
                                  uint64_t sequence,
                                  uint64_t commitGeneration,
                                  std::vector<std::wstring> provenRemoved = {})
    {
        PendingRemovalFocus pending{};
        pending.token = sequence;
        pending.folderPath = folder;
        pending.pathIdentity = insensitive;
        pending.removedFocusDisplayName = std::move(removedFocusDisplayName);
        if (provenRemoved.empty())
        {
            provenRemoved.push_back(pending.removedFocusDisplayName);
        }
        for (size_t sourceIndex = 0u; sourceIndex < provenRemoved.size(); ++sourceIndex)
        {
            pending.sources.push_back(PendingRemovalFocus::Source{
                .sourceIndex = sourceIndex,
                .displayName = std::move(provenRemoved[sourceIndex]),
                .provenRemoved = true,
            });
        }
        pending.folderPathGeneration = view._folderPathGeneration;
        pending.focusOwnershipEpoch = view._removalFocusOwnershipEpoch;
        pending.sortEpoch = view._removalFocusSortEpoch;
        pending.providerEpoch = view._removalFocusProviderEpoch;
        pending.commitEnumerationGeneration = commitGeneration;
        pending.commitSequence = sequence;
        pending.expiresAt = std::chrono::steady_clock::now() + std::chrono::seconds{30};
        pending.committed = true;
        view._pendingRemovalFocus.push_back(std::move(pending));
    };

    {
        FolderView view;
        addCommitted(view, L"B", 1u, 10u);
        ++view._removalFocusOwnershipEpoch;
        if (view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"C"})).has_value() ||
            ! view._pendingRemovalFocus.empty())
        {
            return L"newer user focus ownership did not cancel an older removal intent";
        }
    }
    {
        FolderView view;
        view._displayedFolder = folder;
        view._currentFolder = folder;
        view._itemsFolder = folder;
        view._items = makeItems({L"A", L"B", L"C"});
        view._focusedIndex = 1u;
        view._anchorIndex = 1u;
        view._items[1].focused = true;
        view._items[0].selected = true;
        view._items[2].selected = true;
        view._selectionStats.selectedFiles = 2u;
        addCommitted(view, L"B", 1u, 10u);
        const uint64_t ownershipBeforeSelectionClear = view._removalFocusOwnershipEpoch;
        view.ClearSelection();
        const std::optional<std::vector<std::wstring>> provenRemoved =
            view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"C"}));
        if (view._removalFocusOwnershipEpoch != ownershipBeforeSelectionClear || ! provenRemoved.has_value() ||
            provenRemoved->size() != 1u || provenRemoved.value()[0] != L"B")
        {
            return L"selection-only clear changed current ownership or invalidated a committed removal intent";
        }
    }
    {
        FolderView view;
        view._displayedFolder = folder;
        view._currentFolder = folder;
        view._itemsFolder = folder;
        view._items = makeItems({L"A", L"B", L"C"});
        view._focusedIndex = 1u;
        view._anchorIndex = 1u;
        view._items[1].focused = true;
        addCommitted(view, L"B", 1u, 10u);
        view.FocusItem(2u, false);
        if (view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"C"})).has_value() ||
            ! view._pendingRemovalFocus.empty())
        {
            return L"actual current-item ownership change did not invalidate a committed removal intent";
        }
    }
    for (const uint32_t changedEpoch : {0u, 1u, 2u})
    {
        FolderView view;
        addCommitted(view, L"B", 1u, 10u);
        if (changedEpoch == 0u)
        {
            ++view._folderPathGeneration;
        }
        else if (changedEpoch == 1u)
        {
            ++view._removalFocusSortEpoch;
        }
        else
        {
            ++view._removalFocusProviderEpoch;
        }
        if (view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"C"})).has_value() ||
            ! view._pendingRemovalFocus.empty())
        {
            return L"navigation, sort, or provider epoch change did not invalidate an older removal intent";
        }
    }
    {
        FolderView view;
        addCommitted(view, L"B", 1u, 10u);
        if (view.ResolvePendingRemovalFocusForEnumeration(10u, folder, makeItems({L"A", L"C"})).has_value() ||
            view._pendingRemovalFocus.size() != 1u)
        {
            return L"an in-flight enumeration was accepted as post-completion evidence";
        }
        if (view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"B", L"C"})).has_value() ||
            view._pendingRemovalFocus.size() != 1u)
        {
            return L"an unreflected newer snapshot consumed the pending removal intent";
        }
        const std::optional<std::vector<std::wstring>> provenRemoved =
            view.ResolvePendingRemovalFocusForEnumeration(12u, folder, makeItems({L"A", L"C"}));
        if (! provenRemoved.has_value() || provenRemoved->size() != 1u || provenRemoved.value()[0] != L"B" ||
            ! view._pendingRemovalFocus.empty())
        {
            return L"a reflected post-completion snapshot did not resolve the retained intent";
        }
    }
    for (const bool reverseInsertion : {false, true})
    {
        FolderView view;
        if (reverseInsertion)
        {
            addCommitted(view, L"D", 2u, 10u);
            addCommitted(view, L"B", 1u, 10u);
        }
        else
        {
            addCommitted(view, L"B", 1u, 10u);
            addCommitted(view, L"D", 2u, 10u);
        }
        const std::optional<std::vector<std::wstring>> provenRemoved =
            view.ResolvePendingRemovalFocusForEnumeration(11u, folder, makeItems({L"A", L"C", L"E"}));
        if (! provenRemoved.has_value() || provenRemoved->size() != 1u || provenRemoved.value()[0] != L"D" ||
            ! view._pendingRemovalFocus.empty())
        {
            return L"coalesced removal intents were not all consumed with newest-sequence precedence";
        }
    }

    return {};
}
#endif

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
    pending.commandId    = commandId;
    pending.generation   = 0;
    pending.targetFolder = targetFolder;
    if (! expectedFocusDisplayName.empty())
    {
        pending.expectedFocusDisplayName.assign(expectedFocusDisplayName);
    }

    _pendingExternalCommandAfterEnumeration = std::move(pending);
}

void FolderView::SelectDisplayNamesAfterNextEnumeration(const std::filesystem::path& targetFolder,
                                                        std::vector<std::wstring> displayNames) noexcept
{
    if (targetFolder.empty() || displayNames.empty())
    {
        _pendingExternalSelectionAfterEnumeration.reset();
        return;
    }

    PendingExternalSelection pending{};
    pending.targetFolder = targetFolder;
    pending.displayNames = std::move(displayNames);
    _pendingExternalSelectionAfterEnumeration = std::move(pending);
}

std::optional<std::wstring> FolderView::TryBuildFocusMemoryLocationKey(const std::filesystem::path& folder) const noexcept
{
    if (_fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity &&
        ! TryMakePathKey(_fileSystemPathIdentity.value(), folder.native()).has_value())
    {
        return std::nullopt;
    }

    std::wstring locationKey = BuildLogicalLocationKey(folder);
    if (locationKey.empty())
    {
        return std::nullopt;
    }
    return locationKey;
}

std::optional<std::wstring> FolderView::TryBuildFocusMemoryItemIdentityKey(const std::wstring_view itemDisplayName) const noexcept
{
    if (itemDisplayName.empty())
    {
        return std::nullopt;
    }

    if (_fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity)
    {
        return TryMakeComponentKey(_fileSystemPathIdentity.value(), itemDisplayName);
    }
    return std::wstring(itemDisplayName);
}

void FolderView::RecordFocusMemoryNoCache(const FocusMemoryNoCacheReason reason) noexcept
{
    ++_focusMemoryNoCacheCount;
    _focusMemoryLastNoCacheReason = reason;
    switch (reason)
    {
    case FocusMemoryNoCacheReason::LocationIdentity:
        ++_focusMemoryLocationIdentityNoCacheCount;
        break;
    case FocusMemoryNoCacheReason::ItemIdentity:
        ++_focusMemoryItemIdentityNoCacheCount;
        break;
    case FocusMemoryNoCacheReason::OversizedEntry:
        ++_focusMemoryOversizedEntryNoCacheCount;
        break;
    case FocusMemoryNoCacheReason::None:
        break;
    }
}

void FolderView::TouchFocusMemoryEntry(FocusMemoryEntry& entry) noexcept
{
    _focusMemoryRecency.splice(_focusMemoryRecency.begin(), _focusMemoryRecency, entry.recencyIterator);
    entry.recencyIterator = _focusMemoryRecency.begin();
}

void FolderView::EvictFocusMemoryEntry(const FocusMemoryRecencyList::iterator recencyIterator) noexcept
{
    const std::wstring* const locationKey = *recencyIterator;
    if (! locationKey)
    {
        std::terminate();
    }
    const auto mapIterator = _focusMemory.find(*locationKey);
    if (mapIterator == _focusMemory.end() || mapIterator->second.payloadBytes > _focusMemoryPayloadBytes)
    {
        std::terminate();
    }

    _focusMemoryPayloadBytes -= mapIterator->second.payloadBytes;
    _focusMemoryRecency.erase(recencyIterator);
    _focusMemory.erase(mapIterator);
    ++_focusMemoryEvictionCount;
}

void FolderView::ClearFocusMemory() noexcept
{
    _focusMemoryRecency.clear();
    _focusMemory.clear();
    _focusMemoryPayloadBytes = 0u;
}

bool FolderView::RecordFocusMemoryEntry(const std::filesystem::path& folder, const std::wstring_view itemDisplayName) noexcept
{
    std::optional<std::wstring> locationKey = TryBuildFocusMemoryLocationKey(folder);
    if (! locationKey.has_value())
    {
        RecordFocusMemoryNoCache(FocusMemoryNoCacheReason::LocationIdentity);
        return false;
    }

    std::optional<std::wstring> itemIdentityKey = TryBuildFocusMemoryItemIdentityKey(itemDisplayName);
    if (! itemIdentityKey.has_value())
    {
        RecordFocusMemoryNoCache(FocusMemoryNoCacheReason::ItemIdentity);
        return false;
    }

    constexpr size_t maximumPayloadCodeUnits = kFocusMemoryMaxPayloadBytes / sizeof(wchar_t);
    if (locationKey->size() > maximumPayloadCodeUnits || itemIdentityKey->size() > maximumPayloadCodeUnits - locationKey->size())
    {
        RecordFocusMemoryNoCache(FocusMemoryNoCacheReason::OversizedEntry);
        return false;
    }
    const size_t payloadBytes = (locationKey->size() + itemIdentityKey->size()) * sizeof(wchar_t);

    const auto existing = _focusMemory.find(locationKey.value());
    if (existing != _focusMemory.end())
    {
        if (existing->second.payloadBytes > _focusMemoryPayloadBytes)
        {
            std::terminate();
        }
        _focusMemoryPayloadBytes -= existing->second.payloadBytes;
        existing->second.itemIdentityKey = std::move(itemIdentityKey.value());
        existing->second.payloadBytes = payloadBytes;
        _focusMemoryPayloadBytes += payloadBytes;
        TouchFocusMemoryEntry(existing->second);
    }
    else
    {
        auto [inserted, didInsert] = _focusMemory.emplace(std::move(locationKey.value()), FocusMemoryEntry{});
        if (! didInsert)
        {
            std::terminate();
        }
        inserted->second.itemIdentityKey = std::move(itemIdentityKey.value());
        inserted->second.payloadBytes = payloadBytes;
        _focusMemoryRecency.push_front(std::addressof(inserted->first));
        inserted->second.recencyIterator = _focusMemoryRecency.begin();
        _focusMemoryPayloadBytes += payloadBytes;
    }

    while (_focusMemory.size() > kFocusMemoryMaxEntries || _focusMemoryPayloadBytes > kFocusMemoryMaxPayloadBytes)
    {
        EvictFocusMemoryEntry(std::prev(_focusMemoryRecency.end()));
    }
    _focusMemoryLastNoCacheReason = FocusMemoryNoCacheReason::None;
    return true;
}

std::wstring FolderView::GetRememberedFocusedItemPathForFolder(const std::filesystem::path& folder) noexcept
{
    const std::optional<std::wstring> locationKey = TryBuildFocusMemoryLocationKey(folder);
    if (! locationKey.has_value())
    {
        RecordFocusMemoryNoCache(FocusMemoryNoCacheReason::LocationIdentity);
        return {};
    }

    const auto it = _focusMemory.find(locationKey.value());
    if (it == _focusMemory.end())
    {
        ++_focusMemoryMissCount;
        _focusMemoryLastNoCacheReason = FocusMemoryNoCacheReason::None;
        return {};
    }

    ++_focusMemoryHitCount;
    _focusMemoryLastNoCacheReason = FocusMemoryNoCacheReason::None;
    TouchFocusMemoryEntry(it->second);
    return it->second.itemIdentityKey;
}

void FolderView::ProcessEnumerationResult(std::unique_ptr<EnumerationPayload> payload)
{
    TRACER;
    if (! payload)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: null payload");
#endif
        return;
    }

    const uint64_t currentGeneration = _enumerationGeneration.load(std::memory_order_acquire);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::ProcessEnumerationResult: payload generation={} current={} status=0x{:08X} itemCount={}",
                                              payload->generation,
                                              currentGeneration,
                                              static_cast<unsigned>(payload->status),
                                              payload->items.size()));
#endif
    if (payload->generation != currentGeneration)
    {
        CancelPendingRefreshToPaint(payload->generation);
        if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }
        if (_pendingExternalSelectionAfterEnumeration && _pendingExternalSelectionAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalSelectionAfterEnumeration.reset();
        }
        return;
    }

    CancelBusyOverlay(payload->generation);

    if (FAILED(payload->status))
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: failed payload status");
#endif
        CancelPendingRefreshToPaint(payload->generation);
        if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalCommandAfterEnumeration.reset();
        }
        if (_pendingExternalSelectionAfterEnumeration && _pendingExternalSelectionAfterEnumeration->generation == payload->generation)
        {
            _pendingExternalSelectionAfterEnumeration.reset();
        }

        if (_pendingNavigationDisplayedModel.has_value())
        {
            PendingNavigationDisplayedModel displayed = std::move(_pendingNavigationDisplayedModel.value());
            _pendingNavigationDisplayedModel.reset();
            _displayedFolder = std::move(displayed.displayedFolder);
            _displayedLocationKey = std::move(displayed.displayedLocationKey);
            _items = std::move(displayed.items);
            _itemsArenaBuffer = std::move(displayed.itemsArenaBuffer);
            _itemsFolder = std::move(displayed.itemsFolder);
            _selectionStats = displayed.selectionStats;
            _focusedIndex = displayed.focusedIndex;
            _hoveredIndex = displayed.hoveredIndex;
            _anchorIndex = displayed.anchorIndex;
            _lastCurrentResolutionReason = displayed.resolutionReason;
            _scrollOffset = displayed.scrollOffset;
            _horizontalOffset = displayed.horizontalOffset;
            _columnLayout.clear();
            _columnCounts.clear();
            _columnPrefixSums.clear();
            _itemMetricsCached = false;
            LayoutItems();
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
    DisarmPotentialDrag();

    const auto invalidIndex = static_cast<size_t>(-1);
    const std::wstring targetLocationKey = _currentFolder ? BuildLogicalLocationKey(_currentFolder.value()) : std::wstring{};
    const bool isRefresh = ! targetLocationKey.empty() && targetLocationKey == _displayedLocationKey;
    if (isRefresh && _pendingNavigationDisplayedModel.has_value())
    {
        PendingNavigationDisplayedModel displayed = std::move(_pendingNavigationDisplayedModel.value());
        _pendingNavigationDisplayedModel.reset();
        _displayedFolder = std::move(displayed.displayedFolder);
        _displayedLocationKey = std::move(displayed.displayedLocationKey);
        _items = std::move(displayed.items);
        _itemsArenaBuffer = std::move(displayed.itemsArenaBuffer);
        _itemsFolder = std::move(displayed.itemsFolder);
        _selectionStats = displayed.selectionStats;
        _focusedIndex = displayed.focusedIndex;
        _hoveredIndex = displayed.hoveredIndex;
        _anchorIndex = displayed.anchorIndex;
        _lastCurrentResolutionReason = displayed.resolutionReason;
        _scrollOffset = displayed.scrollOffset;
        _horizontalOffset = displayed.horizontalOffset;
    }

    CurrentResolutionInput currentResolution{};
    currentResolution.sameLogicalLocation = isRefresh;
    currentResolution.previousSelectionStats = ! isRefresh && _pendingNavigationDisplayedModel.has_value()
        ? _pendingNavigationDisplayedModel->selectionStats
        : _selectionStats;
    const bool hadSelectionBeforeResult = currentResolution.previousSelectionStats.selectedFiles != 0u ||
        currentResolution.previousSelectionStats.selectedFolders != 0u;
    currentResolution.selectionMembershipChanged = ! isRefresh && hadSelectionBeforeResult;
    size_t previousFocusedIndex = invalidIndex;
    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        previousFocusedIndex = _focusedIndex;
        const FolderItem& previousCurrent = _items[_focusedIndex];
        currentResolution.previousCurrentIdentity = std::wstring(previousCurrent.displayName);
        if (isRefresh)
        {
            currentResolution.genericRepairProbe = CurrentRepairProbe{
                .displayName = std::wstring(previousCurrent.displayName),
                .extensionOffset = previousCurrent.extensionOffset,
                .isDirectory = previousCurrent.isDirectory,
                .sizeBytes = previousCurrent.sizeBytes,
                .lastWriteTime = previousCurrent.lastWriteTime,
                .fileAttributes = previousCurrent.fileAttributes,
                .unsortedOrder = previousCurrent.unsortedOrder,
            };
        }
    }
    if (isRefresh && _anchorIndex != invalidIndex && _anchorIndex < _items.size())
    {
        currentResolution.previousAnchorIdentity = std::wstring(_items[_anchorIndex].displayName);
    }

    if (! isRefresh)
    {
        ExitIncrementalSearch();
    }

    if (_pendingExplicitCurrentTarget.has_value() && _pendingExplicitCurrentTarget->locationKey == targetLocationKey)
    {
        currentResolution.explicitHostTarget = std::move(_pendingExplicitCurrentTarget->displayName);
        _pendingExplicitCurrentTarget.reset();
    }

    if (isRefresh && currentResolution.previousCurrentIdentity.has_value())
    {
        const auto survivingCurrent = std::find_if(payload->items.begin(), payload->items.end(), [&](const FolderItem& item) noexcept
        { return EquivalentProviderComponent(item.displayName, currentResolution.previousCurrentIdentity.value()); });
        if (survivingCurrent != payload->items.end())
        {
            currentResolution.survivingIdentity = std::wstring(survivingCurrent->displayName);
        }
    }
    if (! isRefresh && _currentFolder)
    {
        std::wstring remembered = GetRememberedFocusedItemPathForFolder(_currentFolder.value());
        if (! remembered.empty())
        {
            currentResolution.navigationMemoryRestore = std::move(remembered);
        }
    }

    if (isRefresh && _currentFolder)
    {
        if (std::optional<std::vector<std::wstring>> provenRemoved =
                ResolvePendingRemovalFocusForEnumeration(payload->generation, _currentFolder.value(), payload->items);
            provenRemoved.has_value() && currentResolution.genericRepairProbe.has_value())
        {
            CurrentRepairProbe& probe = currentResolution.genericRepairProbe.value();
            probe.hostRemovalRepair = true;
            probe.excludedIdentities = std::move(provenRemoved.value());
        }
    }

    if (isRefresh && currentResolution.genericRepairProbe.has_value() && ! currentResolution.survivingIdentity.has_value() &&
        _sortBy == SortBy::None && previousFocusedIndex != invalidIndex)
    {
        std::vector<std::wstring_view> resultingIdentities;
        resultingIdentities.reserve(payload->items.size());
        for (const FolderItem& item : payload->items)
        {
            resultingIdentities.push_back(item.displayName);
        }
        const bool ignoreCase = _fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity &&
            _fileSystemPathIdentity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase;
        const auto identityBefore = [ignoreCase](std::wstring_view left, std::wstring_view right) noexcept
        { return OrdinalString::Compare(left, right, ignoreCase) < 0; };
        std::sort(resultingIdentities.begin(), resultingIdentities.end(), identityBefore);
        const auto survives = [&](std::wstring_view identity) noexcept
        {
            const auto candidate = std::lower_bound(resultingIdentities.begin(), resultingIdentities.end(), identity, identityBefore);
            return candidate != resultingIdentities.end() && EquivalentProviderComponent(*candidate, identity);
        };

        CurrentRepairProbe& probe = currentResolution.genericRepairProbe.value();
        const auto isExcluded = [&](std::wstring_view identity) noexcept
        {
            return std::ranges::any_of(probe.excludedIdentities, [&](const std::wstring& excluded) noexcept
            { return EquivalentProviderComponent(identity, excluded); });
        };
        for (size_t index = previousFocusedIndex + 1u; index < _items.size(); ++index)
        {
            if (survives(_items[index].displayName) && ! isExcluded(_items[index].displayName))
            {
                probe.sortNoneSuccessor = std::wstring(_items[index].displayName);
                break;
            }
        }
        for (size_t index = previousFocusedIndex; index-- > 0u;)
        {
            if (survives(_items[index].displayName) && ! isExcluded(_items[index].displayName))
            {
                probe.sortNonePredecessor = std::wstring(_items[index].displayName);
                break;
            }
        }
    }

    std::vector<PendingRefreshSelectionRename> refreshSelectionRenames;
    if (isRefresh)
    {
        refreshSelectionRenames = std::move(_pendingRefreshSelectionRenames);
    }
    _pendingRefreshSelectionRenames.clear();
    std::vector<bool> refreshSelectionRenameTargetsObserved(refreshSelectionRenames.size(), false);

    // Incremental refresh: preserve rendering state for unchanged items
    size_t itemsPreserved = 0;
    size_t selectionPreserved = 0;
    size_t renameSelectionTransferCount = 0;
    if (isRefresh && ! _items.empty())
    {
        // Build lookup map of old items by path for O(1) access
        const bool ignoreCase = _fileSystemPathIdentity.has_value() && _fileSystemPathIdentity->pathTextStableIdentity &&
            _fileSystemPathIdentity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase;
        std::unordered_map<std::wstring_view, size_t, WStringViewHash, WStringViewEq> oldItemsByPath(
            0u, WStringViewHash{.ignoreCase = ignoreCase}, WStringViewEq{.ignoreCase = ignoreCase});
        oldItemsByPath.reserve(_items.size());
        std::vector<bool> oldItemsObserved(_items.size(), false);
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
            oldItemsObserved[it->second] = true;
            newItem.selected    = oldItem.selected;
            if (oldItem.selected)
            {
                ++selectionPreserved;
            }

            // Check if item data is unchanged (same size, time, attributes)
            const bool dataUnchanged = (oldItem.sizeBytes == newItem.sizeBytes && oldItem.lastWriteTime == newItem.lastWriteTime &&
                                        oldItem.fileAttributes == newItem.fileAttributes && oldItem.isDirectory == newItem.isDirectory);
            if (! dataUnchanged)
            {
                continue; // Item modified, needs fresh rendering
            }

            // Transfer rendering state from old item
            newItem.labelLayout     = oldItem.labelLayout;
            newItem.labelMetrics    = oldItem.labelMetrics;
            newItem.detailsText     = oldItem.detailsText;
            newItem.detailsLayout   = oldItem.detailsLayout;
            newItem.detailsMetrics  = oldItem.detailsMetrics;
            newItem.metadataText    = oldItem.metadataText;
            newItem.metadataLayout  = oldItem.metadataLayout;
            newItem.metadataMetrics = oldItem.metadataMetrics;
            // Only preserve D2D bitmap if icon index matches (icons are shared by extension)
            if (oldItem.iconIndex == newItem.iconIndex && oldItem.icon)
            {
                newItem.icon = oldItem.icon;
            }

            ++itemsPreserved;
        }

        const auto now       = std::chrono::steady_clock::now();
        const auto expiresAt = now + std::chrono::seconds{30};
        std::erase_if(_recentlyMissingRefreshSelections,
                      [&](const RecentlyMissingRefreshSelection& selection) noexcept
        {
            if (selection.expiresAt <= now)
            {
                return true;
            }

            return std::ranges::any_of(payload->items,
                                       [&](const FolderItem& item) noexcept
            { return EquivalentProviderComponent(item.displayName, selection.displayName); });
        });

        for (size_t oldIndex = 0; oldIndex < _items.size(); ++oldIndex)
        {
            const FolderItem& oldItem = _items[oldIndex];
            if (! oldItem.selected || oldItemsObserved[oldIndex])
            {
                continue;
            }

            currentResolution.selectionMembershipChanged = true;

            const auto missingIt = std::find_if(_recentlyMissingRefreshSelections.begin(),
                                                _recentlyMissingRefreshSelections.end(),
                                                [&](const RecentlyMissingRefreshSelection& selection) noexcept
            { return EquivalentProviderComponent(selection.displayName, oldItem.displayName); });
            if (missingIt != _recentlyMissingRefreshSelections.end())
            {
                missingIt->expiresAt = expiresAt;
            }
            else
            {
                _recentlyMissingRefreshSelections.push_back({.displayName = std::wstring(oldItem.displayName), .expiresAt = expiresAt});
            }
        }

        if (! refreshSelectionRenames.empty())
        {
            std::unordered_map<std::wstring_view, std::vector<const PendingRefreshSelectionRename*>, WStringViewHash, WStringViewEq> renamesByTarget(
                0u, WStringViewHash{.ignoreCase = ignoreCase}, WStringViewEq{.ignoreCase = ignoreCase});
            renamesByTarget.reserve(refreshSelectionRenames.size());
            for (const auto& rename : refreshSelectionRenames)
            {
                renamesByTarget[rename.toDisplayName].push_back(&rename);
            }

            for (auto& newItem : payload->items)
            {
                const auto targetIt = renamesByTarget.find(newItem.displayName);
                if (targetIt == renamesByTarget.end())
                {
                    continue;
                }

                for (const auto* rename : targetIt->second)
                {
                    if (! rename)
                    {
                        continue;
                    }

                    const size_t renameIndex = static_cast<size_t>(rename - refreshSelectionRenames.data());
                    if (renameIndex < refreshSelectionRenameTargetsObserved.size())
                    {
                        refreshSelectionRenameTargetsObserved[renameIndex] = true;
                    }

                    if (newItem.selected)
                    {
                        continue;
                    }

                    bool renameSourceSelected = rename->fromWasSelected;
                    if (! renameSourceSelected)
                    {
                        const auto oldIt     = oldItemsByPath.find(rename->fromDisplayName);
                        renameSourceSelected = oldIt != oldItemsByPath.end() && _items[oldIt->second].selected;
                    }

                    if (renameSourceSelected)
                    {
                        newItem.selected = true;
                        ++selectionPreserved;
                        ++renameSelectionTransferCount;
                        currentResolution.selectionMembershipChanged = true;
                        break;
                    }
                }
            }
        }

        if (itemsPreserved > 0)
        {
            Debug::Info(L"FolderView: Incremental refresh preserved {} of {} items", itemsPreserved, payload->items.size());
        }
    }

    for (size_t renameIndex = 0; renameIndex < refreshSelectionRenames.size(); ++renameIndex)
    {
        PendingRefreshSelectionRename& rename = refreshSelectionRenames[renameIndex];
        if (! refreshSelectionRenameTargetsObserved[renameIndex] && rename.fromWasSelected && std::chrono::steady_clock::now() < rename.expiresAt)
        {
            _pendingRefreshSelectionRenames.push_back(std::move(rename));
        }
    }

    if (isRefresh)
    {
        const uint64_t itemCount       = static_cast<uint64_t>(payload->items.size());
        const uint64_t preservedCount  = static_cast<uint64_t>(itemsPreserved);
        const uint64_t rebuildCount    = itemCount >= preservedCount ? itemCount - preservedCount : 0u;
        const uint64_t debounceDelayMs = PendingRefreshDebounceDelayMs(payload->generation);
        Debug::Perf::Emit(L"folder.refresh.preserve_count", L"", 0, preservedCount, itemCount, S_OK);
        Debug::Perf::Emit(L"folder.refresh.rebuild_count", L"", 0, rebuildCount, itemCount, S_OK);
        Debug::Perf::Emit(L"folder.refresh.selection_preserve_count", L"", 0, static_cast<uint64_t>(selectionPreserved), itemCount, S_OK);
        Debug::Perf::Emit(
            L"folder.refresh.rename_transfer_count", L"", 0, static_cast<uint64_t>(renameSelectionTransferCount), static_cast<uint64_t>(refreshSelectionRenames.size()), S_OK);
        Debug::Perf::Emit(L"folder.refresh.missing_selection_count",
                          L"",
                          0,
                          static_cast<uint64_t>(_recentlyMissingRefreshSelections.size()),
                          itemCount,
                          S_OK);
        Debug::Perf::Emit(L"folder.refresh.debounce_delay_ms", L"", 0, debounceDelayMs, itemCount, S_OK);
        Debug::Perf::Emit(L"folder.refresh.enumeration_count", L"", 0, 1u, itemCount, S_OK);
        UpdatePendingRefreshToPaintResult(payload->generation, itemCount);
    }
    else
    {
        CancelPendingRefreshToPaint(payload->generation);
    }

    _items            = std::move(payload->items);
    _itemsArenaBuffer = std::move(payload->arenaBuffer); // Keep arena alive for string_views
    _itemsFolder      = std::move(payload->folder);      // For computing full paths
    _pendingNavigationDisplayedModel.reset();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::ProcessEnumerationResult: assigned items count={}", _items.size()));
#endif
    for (size_t i = 0; i < _items.size(); ++i)
    {
        _items[i].unsortedOrder = i;
    }
    _displayedFolder = _currentFolder;
    _displayedLocationKey = targetLocationKey;
    _focusedIndex    = invalidIndex;
    _anchorIndex     = invalidIndex;
    _hoveredIndex    = invalidIndex;
    ApplyCurrentSort(std::move(currentResolution));
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::ProcessEnumerationResult: after sort focusedIndex={} displayedFolder='{}'",
                                              _focusedIndex,
                                              _displayedFolder.has_value() ? _displayedFolder->native() : std::wstring(L"<none>")));
#endif

    // Only reset scroll position on folder navigation, not on refresh
    if (! isRefresh)
    {
        _scrollOffset     = 0.0f;
        _horizontalOffset = 0.0f;
    }
    _itemMetricsCached = false;

    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    if (_detailsTextProvider && includeDetailsLine)
    {
        for (auto& item : _items)
        {
            if (item.displayName.empty())
            {
                continue;
            }

            std::wstring details =
                _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            if (details != item.detailsText)
            {
                item.detailsText = std::move(details);
                item.detailsLayout.reset();
                item.detailsMetrics = {};
            }
        }
    }

    if (_metadataTextProvider && includeMetadataLine)
    {
        for (auto& item : _items)
        {
            if (item.displayName.empty())
            {
                continue;
            }

            std::wstring metadata =
                _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            if (metadata != item.metadataText)
            {
                item.metadataText = std::move(metadata);
                item.metadataLayout.reset();
                item.metadataMetrics = {};
            }
        }
    }

    // Items already have iconIndex populated from ExecuteEnumeration background thread
    // Now queue icon loading to convert HICON to D2D bitmaps on UI thread
    LayoutItems();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after LayoutItems");
#endif
    UpdateScrollMetrics();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after UpdateScrollMetrics");
#endif
    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        EnsureVisible(_focusedIndex);
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after EnsureVisible");
#endif
    }

    // Queue icon loading after layout - only for items without D2D bitmaps
    Debug::Info(L"FolderView: About to queue icons for {} items", _items.size());
    QueueIconLoading();
    if (_thumbnailsVisible)
    {
        QueueThumbnailLoading();
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after QueueIconLoading");
#endif

    const auto resetEmptyFolderLayouts = [&]() noexcept
    {
        _emptyFolderIconLayout.reset();
        _emptyFolderTitleLayout.reset();
        _emptyFolderFunLayout.reset();
        _emptyFolderLayoutClientSizePx = {};
        _emptyFolderLayoutDpi          = 0.0f;
        _emptyFolderLayoutMessageId    = 0;
        _emptyFolderIconFontSizeDip    = 0.0f;
        _emptyFolderIconMetrics        = {};
        _emptyFolderTitleMetrics       = {};
        _emptyFolderFunMetrics         = {};
    };

    if (! _items.empty())
    {
        const std::wstring_view folderDetail = _itemsFolder.empty() ? std::wstring_view{} : std::wstring_view(_itemsFolder.native());
        StartupMetrics::MarkFirstPanePopulated(folderDetail.empty() ? std::wstring_view(L"(unknown)") : folderDetail, static_cast<uint64_t>(_items.size()));

        if (_emptyFolderState.has_value())
        {
            _emptyFolderState.reset();
            resetEmptyFolderLayouts();
        }
    }
    else
    {
        if (_emptyStateMessage.empty() && _displayedFolder)
        {
            constexpr UINT kFunMessages[] = {
                IDS_EMPTY_FOLDER_FUN_1,
                IDS_EMPTY_FOLDER_FUN_2,
                IDS_EMPTY_FOLDER_FUN_3,
                IDS_EMPTY_FOLDER_FUN_4,
                IDS_EMPTY_FOLDER_FUN_5,
                IDS_EMPTY_FOLDER_FUN_6,
                IDS_EMPTY_FOLDER_FUN_7,
                IDS_EMPTY_FOLDER_FUN_8,
                IDS_EMPTY_FOLDER_FUN_9,
                IDS_EMPTY_FOLDER_FUN_10,
            };

            const std::wstring folderKey = NormalizeFocusMemoryFolderKey(_displayedFolder.value());
            const bool needsNewMessage =
                ! _emptyFolderState.has_value() || _emptyFolderState->folderKey != folderKey || _emptyFolderState->funMessageResourceId == 0;
            if (needsNewMessage)
            {
                const ULONGLONG tick  = GetTickCount64();
                const uint32_t tick32 = static_cast<uint32_t>(tick ^ (tick >> 32));
                const uint32_t seed   = StableHash32(std::wstring_view(folderKey)) ^ tick32;
                const UINT messageId  = kFunMessages[static_cast<size_t>(seed % static_cast<uint32_t>(std::size(kFunMessages)))];

                std::wstring raw = LoadStringResource(nullptr, messageId);
                std::wstring emoji;
                std::wstring funMessage;

                const size_t breakPos = raw.find_first_of(L"\r\n");
                if (breakPos == std::wstring::npos)
                {
                    funMessage = StringUtils::TrimWhitespaceCopy(raw);
                }
                else
                {
                    emoji = StringUtils::TrimWhitespaceCopy(std::wstring_view(raw).substr(0, breakPos));

                    size_t messageStart = breakPos;
                    while (messageStart < raw.size() && (raw[messageStart] == L'\r' || raw[messageStart] == L'\n'))
                    {
                        ++messageStart;
                    }
                    funMessage = StringUtils::TrimWhitespaceCopy(std::wstring_view(raw).substr(messageStart));
                }

                EmptyFolderState state{};
                state.folderKey            = folderKey;
                state.funMessageResourceId = messageId;
                state.emoji                = std::move(emoji);
                state.funMessage           = std::move(funMessage);
                _emptyFolderState          = std::move(state);

                resetEmptyFolderLayouts();
            }
        }
        else if (_emptyFolderState.has_value())
        {
            _emptyFolderState.reset();
            resetEmptyFolderLayouts();
        }
    }

    UpdateCompareNoDifferencesState();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after UpdateCompareNoDifferencesState");
#endif

    // Schedule idle-time layout pre-creation for off-screen items
    // This creates layouts gradually during UI idle periods for smoother scrolling
    ScheduleIdleLayoutCreation();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after ScheduleIdleLayoutCreation");
#endif

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after InvalidateRect");
#endif
    }

    if (_enumerationCompletedCallback)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: before enumeration callback");
#endif
        _enumerationCompletedCallback(_itemsFolder);
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: after enumeration callback");
#endif
    }

    if (_pendingExternalCommandAfterEnumeration && _pendingExternalCommandAfterEnumeration->generation == payload->generation)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: entering pending external command block");
#endif
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
#ifdef ENABLE_TESTS
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::ProcessEnumerationResult: dispatching pending command id={}", pending.commandId));
#endif
            // Execute against the focus identity just validated above. Posting would reopen a window in
            // which another refresh could migrate the command to a different positional neighbor.
            OnCommandMessage(pending.commandId);
        }
    }
    if (_pendingExternalSelectionAfterEnumeration && _pendingExternalSelectionAfterEnumeration->generation == payload->generation)
    {
        PendingExternalSelection pending = std::move(_pendingExternalSelectionAfterEnumeration.value());
        _pendingExternalSelectionAfterEnumeration.reset();
        SetSelectionByDisplayNamePredicate(
            [&pending](std::wstring_view displayName) noexcept
            {
                return std::ranges::any_of(pending.displayNames, [displayName](const std::wstring& wanted) noexcept
                { return displayName == wanted || OrdinalString::EqualsNoCase(displayName, wanted); });
            },
            true);
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"FolderView::ProcessEnumerationResult: end");
#endif
}

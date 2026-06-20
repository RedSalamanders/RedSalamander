// Commands.SelfTest.BatchRename.cpp
// Included from Commands.SelfTest.cpp - NOT compiled standalone.
// Batch Rename test family: pure planning engine first, UI/command cases as the window lands.

void CloseBatchRenameWindowIfOpen() noexcept;
[[nodiscard]] bool SetupBatchRenamePaneFixture(CaseState& state,
                                               const std::filesystem::path& root,
                                               std::initializer_list<std::wstring_view> files,
                                               std::initializer_list<std::wstring_view> directories) noexcept;
[[nodiscard]] bool WaitForPaneDisplayNames(FolderWindow::Pane pane,
                                           std::initializer_list<std::wstring_view> expectedPresent,
                                           std::initializer_list<std::wstring_view> expectedAbsent,
                                           std::chrono::milliseconds timeout) noexcept;

[[nodiscard]] uint64_t CountBatchRenamePerfRowsWithMetric(std::string_view metricText) noexcept
{
    const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    if (path.empty() || metricText.empty())
    {
        return 0u;
    }

    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return 0u;
    }

    std::string needle{"\"metric\":\""};
    needle.append(metricText);
    needle.push_back('"');

    uint64_t count = 0u;
    std::string line;
    while (std::getline(input, line))
    {
        if (line.find(needle) != std::string::npos)
        {
            ++count;
        }
    }
    return count;
}

[[nodiscard]] std::optional<uint64_t> TryReadMaxBatchRenamePerfDurationUs(std::string_view metricText, uint64_t skipMatchingRows = 0u) noexcept
{
    const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    if (path.empty() || metricText.empty())
    {
        return std::nullopt;
    }

    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return std::nullopt;
    }

    std::string metricNeedle{"\"metric\":\""};
    metricNeedle.append(metricText);
    metricNeedle.push_back('"');

    constexpr std::string_view kDurationNeedle{"\"durationUs\":"};
    std::optional<uint64_t> maxDurationUs;
    std::string line;
    while (std::getline(input, line))
    {
        if (line.find(metricNeedle) == std::string::npos)
        {
            continue;
        }

        if (skipMatchingRows != 0u)
        {
            --skipMatchingRows;
            continue;
        }

        const size_t durationPos = line.find(kDurationNeedle);
        if (durationPos == std::string::npos)
        {
            continue;
        }

        size_t cursor = durationPos + kDurationNeedle.size();
        uint64_t value = 0u;
        bool sawDigit = false;
        while (cursor < line.size() && line[cursor] >= '0' && line[cursor] <= '9')
        {
            sawDigit = true;
            const uint64_t digit = static_cast<uint64_t>(line[cursor] - '0');
            if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
            {
                value = (std::numeric_limits<uint64_t>::max)();
                break;
            }

            value = value * 10u + digit;
            ++cursor;
        }

        if (sawDigit && (! maxDurationUs.has_value() || value > maxDurationUs.value()))
        {
            maxDurationUs = value;
        }
    }

    return maxDurationUs;
}

[[nodiscard]] std::optional<uint64_t> TryReadMaxBatchRenamePerfUintField(std::string_view metricText,
                                                                         std::string_view fieldName,
                                                                         uint64_t skipMatchingRows = 0u) noexcept
{
    const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    if (path.empty() || metricText.empty() || fieldName.empty())
    {
        return std::nullopt;
    }

    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return std::nullopt;
    }

    std::string metricNeedle{"\"metric\":\""};
    metricNeedle.append(metricText);
    metricNeedle.push_back('"');

    std::string fieldNeedle{"\""};
    fieldNeedle.append(fieldName);
    fieldNeedle.append("\":");

    std::optional<uint64_t> maxValue;
    std::string line;
    while (std::getline(input, line))
    {
        if (line.find(metricNeedle) == std::string::npos)
        {
            continue;
        }

        if (skipMatchingRows != 0u)
        {
            --skipMatchingRows;
            continue;
        }

        const size_t fieldPos = line.find(fieldNeedle);
        if (fieldPos == std::string::npos)
        {
            continue;
        }

        size_t cursor = fieldPos + fieldNeedle.size();
        uint64_t value = 0u;
        bool sawDigit = false;
        while (cursor < line.size() && line[cursor] >= '0' && line[cursor] <= '9')
        {
            sawDigit = true;
            const uint64_t digit = static_cast<uint64_t>(line[cursor] - '0');
            if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
            {
                value = (std::numeric_limits<uint64_t>::max)();
                break;
            }

            value = value * 10u + digit;
            ++cursor;
        }

        if (sawDigit && (! maxValue.has_value() || value > maxValue.value()))
        {
            maxValue = value;
        }
    }

    return maxValue;
}

class BatchRenameCountingReadDirectoryFileSystem final : public IFileSystem
{
public:
    BatchRenameCountingReadDirectoryFileSystem(wil::com_ptr<IFileSystem> base,
                                               std::atomic_uint32_t* readDirectoryInfoCounter,
                                               std::atomic_uint32_t* renameItemCounter = nullptr,
                                               std::atomic_uint32_t* renameItemsCounter = nullptr,
                                               bool failRenameItemsAsUnsupported = false,
                                               std::atomic_uint32_t* shouldCancelCounter = nullptr,
                                               bool cancelRenameItemsAfterShouldCancel = false,
                                               uint32_t cancelReadDirectoryInfoAtCall = 0u) noexcept
        : _base(std::move(base)),
          _readDirectoryInfoCounter(readDirectoryInfoCounter),
          _renameItemCounter(renameItemCounter),
          _renameItemsCounter(renameItemsCounter),
          _failRenameItemsAsUnsupported(failRenameItemsAsUnsupported),
          _shouldCancelCounter(shouldCancelCounter),
          _cancelRenameItemsAfterShouldCancel(cancelRenameItemsAfterShouldCancel),
          _cancelReadDirectoryInfoAtCall(cancelReadDirectoryInfoAtCall)
    {
    }

    BatchRenameCountingReadDirectoryFileSystem(const BatchRenameCountingReadDirectoryFileSystem&)            = delete;
    BatchRenameCountingReadDirectoryFileSystem(BatchRenameCountingReadDirectoryFileSystem&&)                 = delete;
    BatchRenameCountingReadDirectoryFileSystem& operator=(const BatchRenameCountingReadDirectoryFileSystem&) = delete;
    BatchRenameCountingReadDirectoryFileSystem& operator=(BatchRenameCountingReadDirectoryFileSystem&&)      = delete;

    // Switch RenameItems to a deterministic per-item provider model: each pair
    // is renamed individually through the base RenameItem, reported through
    // FileSystemItemCompleted, and the pair whose source leaf matches failLeaf
    // fails with failHr. With cancelAfterFirstItem the provider cancels after
    // completing (and reporting) the first pair only.
    void SetScriptedPerItemRenameItems(std::wstring failLeaf, const HRESULT failHr, const bool cancelAfterFirstItem) noexcept
    {
        _scriptedPerItemRenameItems = true;
        _failRenameItemLeaf         = std::move(failLeaf);
        _failRenameItemHr           = failHr;
        _cancelAfterFirstRenameItem = cancelAfterFirstItem;
    }

    void SetOmittedCompletionLeaf(std::wstring leaf) noexcept
    {
        _scriptedPerItemRenameItems = true;
        _omitCompletionLeaf         = std::move(leaf);
    }

    void SetReportedFailureWithOverallSuccess(std::wstring leaf, const HRESULT status) noexcept
    {
        _scriptedPerItemRenameItems      = true;
        _reportedFailureLeaf             = std::move(leaf);
        _reportedFailureHr               = status;
        _forceRenameItemsOverallSuccess  = true;
    }

    void SetNoOpRenameItems() noexcept
    {
        _noOpRenameItems = true;
    }

    // Block RenameItems on a test-owned event (bounded) before delegating, so
    // tests can observe the in-flight busy state deterministically.
    void SetRenameItemsGate(const HANDLE gate) noexcept
    {
        _renameItemsGate = gate;
    }

    void SetCapabilitiesResponse(std::string capabilitiesJson, const HRESULT hr = S_OK) noexcept
    {
        _capabilitiesJson = std::move(capabilitiesJson);
        _capabilitiesHr   = hr;
        _overrideCapabilities = true;
    }

    void SetReadDirectoryInfoFailure(const HRESULT hr) noexcept
    {
        _readDirectoryInfoFailureHr = hr;
        _failReadDirectoryInfo      = true;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (current == 0u)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        const uint32_t callIndex = _readDirectoryInfoCounter ? _readDirectoryInfoCounter->fetch_add(1u, std::memory_order_relaxed) + 1u
                                                             : _readDirectoryInfoCalls.fetch_add(1u, std::memory_order_relaxed) + 1u;
        if (_cancelReadDirectoryInfoAtCall != 0u && callIndex >= _cancelReadDirectoryInfoAtCall)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (_failReadDirectoryInfo)
        {
            return _readDirectoryInfoFailureHr;
        }

        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(path, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        if (_renameItemCounter)
        {
            static_cast<void>(_renameItemCounter->fetch_add(1u, std::memory_order_relaxed));
        }
        return _base ? _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (_renameItemsCounter)
        {
            static_cast<void>(_renameItemsCounter->fetch_add(1u, std::memory_order_relaxed));
        }
        if (_failRenameItemsAsUnsupported)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        if (_cancelRenameItemsAfterShouldCancel)
        {
            if (! callback)
            {
                return E_POINTER;
            }
            BOOL cancel = FALSE;
            const HRESULT cancelHr = callback->FileSystemShouldCancel(&cancel, cookie);
            if (_shouldCancelCounter)
            {
                static_cast<void>(_shouldCancelCounter->fetch_add(1u, std::memory_order_relaxed));
            }
            if (FAILED(cancelHr))
            {
                return cancelHr;
            }
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (_renameItemsGate)
        {
            const DWORD waitResult = WaitForSingleObject(_renameItemsGate, 30000u);
            if (waitResult != WAIT_OBJECT_0)
            {
                return HRESULT_FROM_WIN32(WAIT_TIMEOUT);
            }
            return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
        }
        if (_noOpRenameItems)
        {
            if (! items)
            {
                return E_POINTER;
            }

            for (unsigned long index = 0u; index < count; ++index)
            {
                const std::filesystem::path sourcePath{items[index].sourcePath ? items[index].sourcePath : L""};
                const std::filesystem::path destinationPath =
                    sourcePath.parent_path() / (items[index].newName ? items[index].newName : L"");
                if (callback)
                {
                    static_cast<void>(callback->FileSystemItemCompleted(
                        FILESYSTEM_RENAME, index, sourcePath.c_str(), destinationPath.c_str(), S_OK, nullptr, cookie));
                }
            }
            return S_OK;
        }
        if (_scriptedPerItemRenameItems)
        {
            if (! items || ! _base)
            {
                return E_POINTER;
            }

            HRESULT firstFailure = S_OK;
            for (unsigned long index = 0u; index < count; ++index)
            {
                const std::filesystem::path sourcePath{items[index].sourcePath ? items[index].sourcePath : L""};
                const std::filesystem::path destinationPath =
                    sourcePath.parent_path() / (items[index].newName ? items[index].newName : L"");

                HRESULT itemHr = S_OK;
                if (! _failRenameItemLeaf.empty() && sourcePath.filename().native() == _failRenameItemLeaf)
                {
                    itemHr = _failRenameItemHr;
                }
                else if (! _reportedFailureLeaf.empty() && sourcePath.filename().native() == _reportedFailureLeaf)
                {
                    itemHr = _reportedFailureHr;
                }
                else
                {
                    itemHr = _base->RenameItem(sourcePath.c_str(), destinationPath.c_str(), flags, options, nullptr, cookie);
                }

                if (callback && (_omitCompletionLeaf.empty() || sourcePath.filename().native() != _omitCompletionLeaf))
                {
                    static_cast<void>(callback->FileSystemItemCompleted(
                        FILESYSTEM_RENAME, index, sourcePath.c_str(), destinationPath.c_str(), itemHr, nullptr, cookie));
                }
                if (FAILED(itemHr) && SUCCEEDED(firstFailure))
                {
                    firstFailure = itemHr;
                }
                if (_cancelAfterFirstRenameItem)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
            }
            return _forceRenameItemsOverallSuccess ? S_OK : firstFailure;
        }
        return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (_overrideCapabilities)
        {
            if (! jsonUtf8)
            {
                return E_POINTER;
            }
            *jsonUtf8 = SUCCEEDED(_capabilitiesHr) ? _capabilitiesJson.c_str() : nullptr;
            return _capabilitiesHr;
        }
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _base ? _base->GetTransferHints(path, operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _base ? _base->GetStorageCharacteristics(path, characteristics) : E_POINTER;
    }

private:
    ~BatchRenameCountingReadDirectoryFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    std::atomic_uint32_t* _readDirectoryInfoCounter = nullptr;
    std::atomic_uint32_t* _renameItemCounter = nullptr;
    std::atomic_uint32_t* _renameItemsCounter = nullptr;
    bool _failRenameItemsAsUnsupported = false;
    std::atomic_uint32_t* _shouldCancelCounter = nullptr;
    bool _cancelRenameItemsAfterShouldCancel = false;
    uint32_t _cancelReadDirectoryInfoAtCall = 0u;
    bool _scriptedPerItemRenameItems = false;
    std::wstring _failRenameItemLeaf;
    HRESULT _failRenameItemHr = S_OK;
    std::wstring _omitCompletionLeaf;
    std::wstring _reportedFailureLeaf;
    HRESULT _reportedFailureHr = S_OK;
    bool _forceRenameItemsOverallSuccess = false;
    bool _cancelAfterFirstRenameItem = false;
    HANDLE _renameItemsGate = nullptr;
    bool _noOpRenameItems = false;
    bool _overrideCapabilities = false;
    HRESULT _capabilitiesHr = S_OK;
    std::string _capabilitiesJson;
    bool _failReadDirectoryInfo = false;
    HRESULT _readDirectoryInfoFailureHr = E_FAIL;
    std::atomic_uint32_t _readDirectoryInfoCalls{0u};
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameCountingReadDirectoryFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                        std::atomic_uint32_t* counter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, counter);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameCancelingReadDirectoryFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                          std::atomic_uint32_t* counter,
                                                                                          uint32_t cancelAtCall) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(
        base, counter, nullptr, nullptr, false, nullptr, false, cancelAtCall);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameFailingReadDirectoryFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                        std::atomic_uint32_t* counter,
                                                                                        const HRESULT hr) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, counter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetReadDirectoryInfoFailure(hr);
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameBulkUnsupportedFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                   std::atomic_uint32_t* renameItemCounter,
                                                                                   std::atomic_uint32_t* renameItemsCounter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, renameItemCounter, renameItemsCounter, true);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameCancelOnShouldCancelFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                        std::atomic_uint32_t* renameItemsCounter,
                                                                                        std::atomic_uint32_t* shouldCancelCounter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(
        base, nullptr, nullptr, renameItemsCounter, false, shouldCancelCounter, true);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameRenameCountingFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                  std::atomic_uint32_t* renameItemsCounter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameScriptedPerItemFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                   std::atomic_uint32_t* renameItemsCounter,
                                                                                   std::wstring failLeaf,
                                                                                   const HRESULT failHr,
                                                                                   const bool cancelAfterFirstItem) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetScriptedPerItemRenameItems(std::move(failLeaf), failHr, cancelAfterFirstItem);
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameOmittedCompletionFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                     std::atomic_uint32_t* renameItemsCounter,
                                                                                     std::wstring omittedLeaf) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetOmittedCompletionLeaf(std::move(omittedLeaf));
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameReportedFailureOverallSuccessFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                                 std::atomic_uint32_t* renameItemsCounter,
                                                                                                 std::wstring failedLeaf,
                                                                                                 const HRESULT status) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetReportedFailureWithOverallSuccess(std::move(failedLeaf), status);
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameGatedRenameItemsFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                    const HANDLE gate) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetRenameItemsGate(gate);
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameCapabilitiesOverrideFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                        std::atomic_uint32_t* renameItemsCounter,
                                                                                        std::string capabilitiesJson) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetCapabilitiesResponse(std::move(capabilitiesJson));
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateBatchRenameNoOpRenameItemsFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                                   std::atomic_uint32_t* renameItemsCounter,
                                                                                   std::string capabilitiesJson) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) BatchRenameCountingReadDirectoryFileSystem(base, nullptr, nullptr, renameItemsCounter);
    if (! wrapper)
    {
        return {};
    }
    wrapper->SetNoOpRenameItems();
    wrapper->SetCapabilitiesResponse(std::move(capabilitiesJson));
    wrapped.attach(wrapper);
    return wrapped;
}

// Resolves (enabling on demand) the local builtin file-system plugin shared by
// the Batch Rename execution self-tests.
[[nodiscard]] wil::com_ptr<IFileSystem> GetBatchRenameLocalFileSystem(CaseState& state) noexcept
{
    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename selftest: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename selftest.");
    return fileSystem;
}

constexpr std::string_view kBatchRenameNoOpProviderCapabilities = R"json({
  "version": 1,
  "operations": {
    "rename": true
  },
  "concurrency": {},
  "crossFileSystem": {},
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "normalization": "none",
    "componentComparison": "ordinalIgnoreCase",
    "preferredSeparator": "\\",
    "acceptedSeparators": ["\\", "/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
})json";

[[nodiscard]] std::string ReadBatchRenameFileText(const std::filesystem::path& path) noexcept
{
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return {};
    }
    std::string text;
    char buffer[256];
    while (input.read(buffer, sizeof(buffer)) || input.gcount() > 0)
    {
        text.append(buffer, static_cast<size_t>(input.gcount()));
    }
    return text;
}

[[nodiscard]] bool DirectoryHasBatchRenameTempLeftovers(const std::filesystem::path& root) noexcept
{
    std::error_code ec;
    for (std::filesystem::recursive_directory_iterator it(root, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (it->path().filename().native().find(L".rsren-") != std::wstring::npos)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] std::vector<std::wstring> ListBatchRenameDirectoryLeaves(const std::filesystem::path& root) noexcept
{
    std::vector<std::wstring> leaves;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(root, ec), end; ! ec && it != end; it.increment(ec))
    {
        leaves.push_back(it->path().filename().native());
    }
    std::ranges::sort(leaves);
    return leaves;
}

[[nodiscard]] bool HasBatchRenameIssue(const BatchRename::PreviewRow& row,
                                       const BatchRename::IssueSeverity severity,
                                       const std::wstring_view message) noexcept
{
    return std::ranges::any_of(row.issues,
                               [severity, message](const BatchRename::Issue& issue) noexcept
    { return issue.severity == severity && issue.message == message; });
}

struct BatchRenameLocalStampParts final
{
    int year         = 0;
    unsigned month   = 0;
    unsigned day     = 0;
    long long hour   = 0;
    long long minute = 0;
    long long second = 0;
};

// Mirrors the engine timestamp policy: {date}/{time}/{created} macros format local wall-clock
// fields, falling back to UTC when the time-zone database is unavailable.
[[nodiscard]] BatchRenameLocalStampParts GetBatchRenameExpectedLocalParts(const std::chrono::sys_seconds timestamp) noexcept
{
    std::chrono::local_seconds localTime{timestamp.time_since_epoch()};
    try
    {
        localTime = std::chrono::current_zone()->to_local(timestamp);
    }
    catch (const std::runtime_error&)
    {
        // Mirrors the engine fallback: tzdb lookup throws std::runtime_error when the time-zone
        // database is unavailable; keep the UTC wall-clock fields.
    }

    const std::chrono::local_days date = std::chrono::floor<std::chrono::days>(localTime);
    const std::chrono::year_month_day ymd{date};
    const std::chrono::hh_mm_ss timeOfDay{localTime - date};

    BatchRenameLocalStampParts parts{};
    parts.year   = static_cast<int>(ymd.year());
    parts.month  = static_cast<unsigned>(ymd.month());
    parts.day    = static_cast<unsigned>(ymd.day());
    parts.hour   = timeOfDay.hours().count();
    parts.minute = timeOfDay.minutes().count();
    parts.second = timeOfDay.seconds().count();
    return parts;
}

[[nodiscard]] FILETIME BatchRenameFileTimeFromSysSeconds(const std::chrono::sys_seconds timestamp) noexcept
{
    constexpr int64_t kWindowsToUnixEpoch100Ns = 116444736000000000LL;
    constexpr int64_t kFileTimeTicksPerSecond  = 10000000LL;
    const int64_t ticks = kWindowsToUnixEpoch100Ns + (timestamp.time_since_epoch().count() * kFileTimeTicksPerSecond);

    ULARGE_INTEGER value{};
    value.QuadPart = static_cast<ULONGLONG>(ticks);

    FILETIME fileTime{};
    fileTime.dwLowDateTime  = value.LowPart;
    fileTime.dwHighDateTime = value.HighPart;
    return fileTime;
}

[[nodiscard]] bool SetBatchRenameFileCreationTime(const std::filesystem::path& path, const std::chrono::sys_seconds timestamp) noexcept
{
    const FILETIME creationTime = BatchRenameFileTimeFromSysSeconds(timestamp);
    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return false;
    }

    return ::SetFileTime(file.get(), &creationTime, nullptr, nullptr) != FALSE;
}

#ifndef SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE
#define SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE 0x2u
#endif

[[nodiscard]] bool TryCreateBatchRenameDirectorySymlink(const std::filesystem::path& linkPath,
                                                        const std::filesystem::path& targetPath,
                                                        DWORD& outLastError) noexcept
{
    outLastError = ERROR_SUCCESS;
    const DWORD flags = SYMBOLIC_LINK_FLAG_DIRECTORY;
    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags | SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE) != FALSE)
    {
        return true;
    }

    outLastError = ::GetLastError();
    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags) != FALSE)
    {
        outLastError = ERROR_SUCCESS;
        return true;
    }

    outLastError = ::GetLastError();
    return false;
}

[[nodiscard]] bool TestBatchRenameCommandRegistered(CaseState& state) noexcept
{
    const CommandInfo* const command = FindCommandInfo(L"cmd/pane/batchRename");
    state.Require(command != nullptr, L"Batch Rename command should be registered for command lookup.");
    if (command == nullptr)
    {
        return false;
    }

    state.Require(command->wmCommandId == IDM_PANE_BATCH_RENAME, L"Batch Rename command should dispatch through IDM_PANE_BATCH_RENAME.");
    state.Require(command->displayNameStringId == IDS_CMD_BATCH_RENAME, L"Batch Rename command should use its localized display string.");
    state.Require(command->descriptionStringId == IDS_CMD_DESC_BATCH_RENAME, L"Batch Rename command should use its localized description string.");

    const std::wstring displayName = LoadStringResource(nullptr, command->displayNameStringId);
    const std::wstring description = LoadStringResource(nullptr, command->descriptionStringId);
    state.Require(displayName == L"Batch Rename", L"Batch Rename display resource should be present and stable.");
    state.Require(description.find(L"Batch Rename") != std::wstring::npos, L"Batch Rename description resource should be present.");

    state.Require(WndMsg::kBatchRenameTaskUpdate == WM_APP + 0x544, L"Batch Rename task update message should use the reserved central ID.");
    state.Require(WndMsg::kBatchRenameCompleted == WM_APP + 0x545, L"Batch Rename completed message should use the reserved central ID.");
    state.Require(WndMsg::kBatchRenameWindowDebug == WM_APP + 0x546, L"Batch Rename debug message should use the reserved central ID.");

    const std::optional<unsigned int> wmCommandId = TryGetWmCommandId(L"cmd/pane/batchRename");
    state.Require(wmCommandId.has_value() && wmCommandId.value() == IDM_PANE_BATCH_RENAME,
                  L"Batch Rename command should resolve through TryGetWmCommandId.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowOpensFromPaneContext(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"one.txt", root / L"two.md"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open from a pane context.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename window handle should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename window should expose a debug snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.usesDxUiHost, L"Batch Rename window should use a DxUi WindowHost shell.");
    state.Require(snapshot.rootText == root.native(), L"Batch Rename root text should be seeded from the pane context.");
    state.Require(snapshot.rootNavigationVisible, L"Batch Rename root navigation header should be visible.");
    state.Require(snapshot.rootNavigationUsesNavigationView, L"Batch Rename root header should use the shared NavigationView control.");
    state.Require(snapshot.rootNavigationPathText == root.native(),
                  L"Batch Rename NavigationView header should be populated with the pane context root.");
    state.Require(snapshot.visibleChildWindowCount >= 1u, L"Batch Rename window should expose the NavigationView child window.");
    state.Require(snapshot.previewRowCount == 2u, L"Batch Rename preview should seed one row per initial path.");
    state.Require(snapshot.previewColumnIds == std::vector<std::wstring>{L"original", L"new", L"size", L"date", L"time", L"path"},
                  L"Batch Rename preview columns should keep the v1 order.");
    state.Require(snapshot.previewIconCellCount == snapshot.previewRowCount,
                  L"Batch Rename preview should render one icon-enabled Original Name cell per row.");
    state.Require(snapshot.originalIconIndices.size() == snapshot.previewRowCount,
                  L"Batch Rename preview should expose one resolved shell icon index per row.");
    state.Require(std::all_of(snapshot.originalIconIndices.begin(),
                              snapshot.originalIconIndices.end(),
                              [](const int iconIndex) noexcept { return iconIndex >= 0; }),
                  L"Batch Rename preview Original Name icons should resolve through the system icon cache.");
    state.Require(snapshot.originalNames == std::vector<std::wstring>{L"one.txt", L"two.md"},
                  L"Batch Rename preview should show original leaf names.");
    state.Require(snapshot.newNames == snapshot.originalNames, L"Batch Rename initial preview should keep names unchanged.");
    state.Require(snapshot.fullPaths == std::vector<std::wstring>{(root / L"one.txt").native(), (root / L"two.md").native()},
                  L"Batch Rename preview should retain full source paths.");

    PostMessageW(batchWindow, WM_CLOSE, 0, 0);
    PumpPendingMessages();
    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowFolderScopeCollectsLocalChildrenMetadata(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_folder_scope_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root / L"scope-sub"), L"Failed to create Batch Rename folder-scope subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "payload"), L"Failed to create Batch Rename folder-scope file.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.md", "markdown"), L"Failed to create second Batch Rename folder-scope file.");
    state.Require(SelfTest::WriteTextFile(root / L"scope-sub" / L"nested.txt", "nested"), L"Failed to create nested Batch Rename folder-scope file.");
    state.Require(SelfTest::WriteTextFile(root / L"scope-sub" / L"nested.md", "nested-md"), L"Failed to create nested non-matching Batch Rename file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRenamePaneContext context{};
    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin for Batch Rename folder-scope collection: 0x{:08X}.",
                                                       static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename folder-scope collection.");
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t readDirectoryInfoCalls{0u};
    wil::com_ptr<IFileSystem> countingFileSystem = CreateBatchRenameCountingReadDirectoryFileSystem(fileSystem, &readDirectoryInfoCalls);
    state.Require(countingFileSystem != nullptr, L"Batch Rename folder-scope selftest should create a counting file-system wrapper.");
    if (! countingFileSystem)
    {
        return false;
    }

    const uint64_t collectDurationBefore = CountBatchRenamePerfRowsWithMetric("batchrename.collect.us");
    const uint64_t collectTargetsBefore  = CountBatchRenamePerfRowsWithMetric("batchrename.collect.targets");

    context.fileSystem      = countingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-folder-scope-selftest";
    context.rootPluginPath  = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-folder-scope-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for folder-scope collection testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename folder-scope test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename folder-scope snapshot should be available.");
    state.Require(snapshot.rootText == root.native(), L"Folder-scope Batch Rename should display the scope root.");
    state.Require(snapshot.previewRowCount == 2u, L"Folder-scope Batch Rename should collect immediate local files by default.");
    state.Require(snapshot.originalNames == std::vector<std::wstring>{L"alpha.txt", L"beta.md"},
                  L"Default folder-scope collection should include files and exclude folders.");
    state.Require(snapshot.newNames == snapshot.originalNames, L"Folder-scope initial preview should keep names unchanged.");
    state.Require(snapshot.fullPaths == std::vector<std::wstring>{(root / L"alpha.txt").native(), (root / L"beta.md").native()},
                  L"Folder-scope collection should retain full source paths.");
    state.Require(snapshot.sizeTexts.size() == 2u && snapshot.sizeTexts[0] == L"7" && snapshot.sizeTexts[1] == L"8",
                  L"Folder-scope metadata should expose file byte sizes.");
    state.Require(snapshot.dateTexts.size() == 2u && ! snapshot.dateTexts[0].empty(),
                  L"Folder-scope metadata should expose a date for collected local files.");
    state.Require(snapshot.timeTexts.size() == 2u && ! snapshot.timeTexts[0].empty(),
                  L"Folder-scope metadata should expose a time for collected local files.");

    state.Require(DebugSetBatchRenameWindowScope(L"*.txt", true, true, false),
                  L"Batch Rename folder-scope debug hook should update mask and recursive file options.");

    BatchRenameDebugSnapshot recursiveFilesSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(recursiveFilesSnapshot), L"Batch Rename recursive scope snapshot should be available.");
    state.Require(recursiveFilesSnapshot.originalNames == std::vector<std::wstring>{L"alpha.txt", L"nested.txt"},
                  L"Recursive folder-scope collection should include matching files in stable path order.");
    state.Require(recursiveFilesSnapshot.fullPaths == std::vector<std::wstring>{(root / L"alpha.txt").native(), (root / L"scope-sub" / L"nested.txt").native()},
                  L"Recursive folder-scope collection should retain nested source paths.");

    BatchRename::Rules relativeRules{};
    relativeRules.nameTemplate     = L"rel-{relativeFolderFlat}-{name}";
    relativeRules.flattenSeparator = L"__";
    state.Require(DebugSetBatchRenameWindowRules(relativeRules),
                  L"Batch Rename folder-scope debug hook should update root-relative folder macro rules.");

    BatchRenameDebugSnapshot recursiveRelativeSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(recursiveRelativeSnapshot),
                  L"Batch Rename recursive relative-folder macro snapshot should be available.");
    state.Require(recursiveRelativeSnapshot.newNames == std::vector<std::wstring>{L"rel--alpha.txt", L"rel-scope-sub-nested.txt"},
                  L"Recursive folder-scope previews should expand relative-folder macros from the active pane root.");

    state.Require(DebugSetBatchRenameWindowScope(L"scope-*", false, false, true),
                  L"Batch Rename folder-scope debug hook should update folder-only scope options.");

    BatchRenameDebugSnapshot folderOnlySnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(folderOnlySnapshot), L"Batch Rename folder-only scope snapshot should be available.");
    state.Require(folderOnlySnapshot.originalNames == std::vector<std::wstring>{L"scope-sub"},
                  L"Folder-only scope collection should include matching immediate folders.");
    state.Require(folderOnlySnapshot.sizeTexts.size() == 1u && folderOnlySnapshot.sizeTexts[0].empty(),
                  L"Folder-only scope collection should keep folder size blank.");
    state.Require(readDirectoryInfoCalls.load(std::memory_order_relaxed) > 0u,
                  L"Batch Rename folder-scope collection should enumerate through IFileSystem::ReadDirectoryInfo when a provider is available.");

    const uint64_t collectDurationAfter = CountBatchRenamePerfRowsWithMetric("batchrename.collect.us");
    const uint64_t collectTargetsAfter  = CountBatchRenamePerfRowsWithMetric("batchrename.collect.targets");
    state.Require(collectDurationAfter > collectDurationBefore,
                  std::format(L"Batch Rename folder-scope collection should emit batchrename.collect.us; before={} after={}.",
                              collectDurationBefore,
                              collectDurationAfter));
    state.Require(collectTargetsAfter > collectTargetsBefore,
                  std::format(L"Batch Rename folder-scope collection should emit batchrename.collect.targets; before={} after={}.",
                              collectTargetsBefore,
                              collectTargetsAfter));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameTargetCollectionRespectsProviderCancellation(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_collect_cancel_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root / L"nested"), L"Failed to create Batch Rename cancelable collection subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create Batch Rename cancelable collection root file.");
    state.Require(SelfTest::WriteTextFile(root / L"nested" / L"child.txt", "child"),
                  L"Failed to create Batch Rename cancelable collection nested file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename cancelable collection: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename cancelable collection.");
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t readDirectoryInfoCalls{0u};
    wil::com_ptr<IFileSystem> cancelingFileSystem =
        CreateBatchRenameCancelingReadDirectoryFileSystem(fileSystem, &readDirectoryInfoCalls, 2u);
    state.Require(cancelingFileSystem != nullptr, L"Batch Rename cancelable collection test should create a canceling file-system wrapper.");
    if (! cancelingFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = cancelingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-collect-cancel-selftest";
    context.rootPluginPath  = root;

    BatchRenameDebugCollectionResult result{};
    bool collectReturned = false;
    HRESULT workerCoInitHr = S_OK;
    std::jthread worker([&](std::stop_token) noexcept
    {
        workerCoInitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(workerCoInitHr))
        {
            return;
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;
        collectReturned = DebugCollectBatchRenameTargetsForTests(std::move(context), L"*.*", true, true, false, result);
    });
    worker = std::jthread{};

    state.Require(SUCCEEDED(workerCoInitHr),
                  std::format(L"Batch Rename collection worker should initialize COM MTA, hr=0x{:08X}.",
                              static_cast<unsigned long>(workerCoInitHr)));
    state.Require(collectReturned, L"Batch Rename debug collection helper should run on the worker thread.");
    state.Require(result.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Provider cancellation during Batch Rename collection should return ERROR_CANCELLED; saw 0x{:08X}.",
                              static_cast<unsigned long>(result.hr)));
    state.Require(readDirectoryInfoCalls.load(std::memory_order_relaxed) == 2u,
                  std::format(L"Cancelable collection should stop at the provider cancellation point; saw {} ReadDirectoryInfo calls.",
                              readDirectoryInfoCalls.load(std::memory_order_relaxed)));
    state.Require(result.fullPaths == std::vector<std::wstring>{(root / L"root.txt").native()},
                  L"Canceled recursive collection should keep only targets collected before provider cancellation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameRecursiveCollectionSkipsSymlinkLoops(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_symlink_loop_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root / L"real"), L"Failed to create Batch Rename symlink-loop real folder.");
    state.Require(SelfTest::WriteTextFile(root / L"needle.txt", "root"), L"Failed to create Batch Rename symlink-loop root file.");
    state.Require(SelfTest::WriteTextFile(root / L"real" / L"child.txt", "child"), L"Failed to create Batch Rename symlink-loop child file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const std::filesystem::path loopPath = root / L"loop";
    DWORD symlinkError = ERROR_SUCCESS;
    if (! TryCreateBatchRenameDirectorySymlink(loopPath, root, symlinkError))
    {
        if (symlinkError == ERROR_PRIVILEGE_NOT_HELD || symlinkError == ERROR_ACCESS_DENIED || symlinkError == ERROR_INVALID_PARAMETER)
        {
            Debug::Warning(L"BatchRename selftest: skipping recursive symlink-loop guard test (CreateSymbolicLinkW failed: {}).",
                           static_cast<unsigned long>(symlinkError));
            return true;
        }

        state.Require(false,
                      std::format(L"CreateSymbolicLinkW failed unexpectedly for Batch Rename symlink-loop guard: {}.",
                                  static_cast<unsigned long>(symlinkError)));
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-symlink-loop-selftest";
    context.rootPluginPath  = root;

    BatchRenameDebugCollectionResult result{};
    bool collectReturned = false;
    HRESULT workerCoInitHr = S_OK;
    std::jthread worker([&](std::stop_token) noexcept
    {
        workerCoInitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(workerCoInitHr))
        {
            return;
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;
        collectReturned = DebugCollectBatchRenameTargetsForTests(std::move(context), L"*.txt", true, true, false, result);
    });
    worker = std::jthread{};

    state.Require(SUCCEEDED(workerCoInitHr),
                  std::format(L"Batch Rename symlink-loop collection worker should initialize COM MTA, hr=0x{:08X}.",
                              static_cast<unsigned long>(workerCoInitHr)));
    state.Require(collectReturned, L"Batch Rename symlink-loop debug collection helper should return.");
    state.Require(SUCCEEDED(result.hr),
                  std::format(L"Recursive Batch Rename collection with a symlink loop should succeed, hr=0x{:08X}.",
                              static_cast<unsigned long>(result.hr)));
    state.Require(result.fullPaths == std::vector<std::wstring>{(root / L"needle.txt").native(), (root / L"real" / L"child.txt").native()},
                  L"Recursive Batch Rename collection should collect real files once and skip the symlink loop.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameProviderSelectionFallbackMarksMetadataUnknown(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_nonlocal_selection_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename non-local selection root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t readDirectoryInfoCalls{0u};
    wil::com_ptr<IFileSystem> failingFileSystem =
        CreateBatchRenameFailingReadDirectoryFileSystem(fileSystem, &readDirectoryInfoCalls, HRESULT_FROM_WIN32(ERROR_NOT_FOUND));
    state.Require(failingFileSystem != nullptr, L"Batch Rename non-local selection test should create a failing file-system wrapper.");
    if (! failingFileSystem)
    {
        return false;
    }

    const std::filesystem::path selectedPath = root / L"ghost.remote";
    BatchRenamePaneContext context{};
    context.fileSystem      = failingFileSystem;
    context.pluginId        = L"test/non-local-provider";
    context.pluginShortId   = L"nonlocal";
    context.instanceContext = L"batch-rename-nonlocal-selection-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {selectedPath};

    BatchRenameDebugCollectionResult result{};
    const bool collectReturned = DebugCollectBatchRenameTargetsForTests(std::move(context), L"*.*", false, true, false, result);
    state.Require(collectReturned, L"Batch Rename non-local selection debug collection should return.");
    state.Require(SUCCEEDED(result.hr),
                  std::format(L"Non-local undescribed selection fallback should keep collection successful, hr=0x{:08X}.",
                              static_cast<unsigned long>(result.hr)));
    state.Require(readDirectoryInfoCalls.load(std::memory_order_relaxed) == 1u,
                  std::format(L"Non-local selection should ask the provider to describe the parent once; saw {} calls.",
                              readDirectoryInfoCalls.load(std::memory_order_relaxed)));
    state.Require(result.originalNames == std::vector<std::wstring>{L"ghost.remote"} &&
                      result.fullPaths == std::vector<std::wstring>{selectedPath.native()},
                  L"Non-local undescribed selection should preserve the selected path identity.");
    state.Require(result.isDirectories == std::vector<bool>{false},
                  L"Undescribed non-local selection should not invent a directory classification.");
    state.Require(result.sizeBytes == std::vector<uint64_t>{0u},
                  L"Undescribed non-local selection should keep size bytes at the neutral value.");
    state.Require(result.metadataUnknowns == std::vector<bool>{true},
                  L"Undescribed non-local selection should be marked metadata-unknown instead of fabricated as a local 0-byte file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowRulesRecomputePreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameRulesSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-rules-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Alpha File.TXT", root / L"beta file.md"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-rules-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for rules recompute testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename rules test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate        = L"{counter:000}_{stem}{ext}";
    rules.searchFor           = L"file";
    rules.replaceWith         = L"clip";
    rules.caseSensitive       = false;
    rules.excludeExtension    = true;
    rules.fileNameCaseStyle   = BatchRename::CaseTransform::Upper;
    rules.extensionCaseStyle  = BatchRename::CaseTransform::Lower;

    const uint64_t recomputeRowsBefore      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t visibleRefreshRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename debug rules hook should update the active window.");
    const uint64_t recomputeRowsAfter      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t visibleRefreshRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");
    state.Require(recomputeRowsAfter > recomputeRowsBefore,
                  std::format(L"Batch Rename rule recompute should emit batchrename.preview.recompute.us; before={} after={}.",
                              recomputeRowsBefore,
                              recomputeRowsAfter));
    state.Require(visibleRefreshRowsAfter > visibleRefreshRowsBefore,
                  std::format(L"Batch Rename rule recompute should emit batchrename.preview.visible_refresh.us; before={} after={}.",
                              visibleRefreshRowsBefore,
                              visibleRefreshRowsAfter));

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename rules snapshot should be available.");
    state.Require(snapshot.previewRowCount == 2u, L"Rules recompute should keep the original target count.");
    state.Require(snapshot.originalNames == std::vector<std::wstring>{L"Alpha File.TXT", L"beta file.md"},
                  L"Rules recompute should preserve original names.");
    state.Require(snapshot.newNames == std::vector<std::wstring>{L"001_ALPHA CLIP.txt", L"002_BETA CLIP.md"},
                  L"Rules recompute should apply macros, search/replace, and case transforms through the preview engine.");
    state.Require(snapshot.changedRowCount == 2u, L"Rules recompute should report changed preview rows.");
    state.Require(snapshot.errorRowCount == 0u, L"Valid rules should report no blocking preview errors.");
    state.Require(snapshot.renameButtonEnabled, L"Rename should be enabled when all preview rows are valid changed rows.");

    BatchRename::Rules noOpRules{};
    noOpRules.nameTemplate = L"{name}";
    state.Require(DebugSetBatchRenameWindowRules(noOpRules), L"Batch Rename debug rules hook should accept no-op rules for warning status testing.");

    BatchRenameDebugSnapshot warningSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(warningSnapshot), L"Batch Rename warning-status snapshot should be available.");
    state.Require(warningSnapshot.warningRowCount == 2u, L"No-op preview rows should be counted as warnings.");
    state.Require(warningSnapshot.newNameStatusIconTexts.size() == 2u &&
                      std::ranges::all_of(warningSnapshot.newNameStatusIconTexts, [](const std::wstring& iconText) noexcept { return ! iconText.empty(); }),
                  L"Warning preview rows should show a status glyph in the New Name column.");
    const std::wstring warningStatusIcon = warningSnapshot.newNameStatusIconTexts.empty() ? std::wstring{} : warningSnapshot.newNameStatusIconTexts.front();
    state.Require(std::ranges::all_of(warningSnapshot.newNameStatusIconTexts, [&warningStatusIcon](const std::wstring& iconText) noexcept {
                      return iconText == warningStatusIcon;
                  }),
                  L"Warning preview rows should use a stable warning status glyph.");
    state.Require(warningSnapshot.newNameTooltips.size() == 2u &&
                      std::ranges::all_of(warningSnapshot.newNameTooltips, [](const std::wstring& tooltip) noexcept {
                          return tooltip.find(L"name_unchanged") != std::wstring::npos;
                      }),
                  L"Warning preview row tooltips should include the stable warning issue id.");
    state.Require(warningSnapshot.statusText.find(L"2 warnings") != std::wstring::npos,
                  L"Batch Rename footer status should include warning counts.");
    state.Require(warningSnapshot.statusText.find(L"0 changed") != std::wstring::npos &&
                      warningSnapshot.statusText.find(L"2 unchanged") != std::wstring::npos,
                  L"Batch Rename footer status should include changed and unchanged counts.");

    BatchRename::Rules mixedNoOpRules{};
    mixedNoOpRules.nameTemplate = L"{name}";
    mixedNoOpRules.searchFor    = L"Alpha";
    mixedNoOpRules.replaceWith  = L"Alpha Renamed";
    state.Require(DebugSetBatchRenameWindowRules(mixedNoOpRules),
                  L"Batch Rename debug rules hook should accept mixed changed/unchanged rules for preview filtering.");

    BatchRenameDebugSnapshot mixedNoOpSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(mixedNoOpSnapshot), L"Batch Rename mixed no-op snapshot should be available.");
    state.Require(! mixedNoOpSnapshot.hideUnchangedRows, L"Batch Rename should show unchanged rows by default.");
    state.Require(mixedNoOpSnapshot.previewRowCount == 2u, L"Mixed preview should show changed and unchanged rows by default.");
    state.Require(mixedNoOpSnapshot.changedRowCount == 1u, L"Mixed preview stats should count one changed row.");
    state.Require(mixedNoOpSnapshot.warningRowCount == 1u, L"Mixed preview stats should count the unchanged row warning.");
    state.Require(mixedNoOpSnapshot.statusText.find(L"1 changed") != std::wstring::npos &&
                      mixedNoOpSnapshot.statusText.find(L"1 unchanged") != std::wstring::npos,
                  L"Mixed preview footer status should keep changed and unchanged counts distinct.");

    state.Require(DebugSetBatchRenameWindowHideUnchanged(true), L"Batch Rename debug helper should enable hiding unchanged preview rows.");

    BatchRenameDebugSnapshot filteredNoOpSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(filteredNoOpSnapshot), L"Batch Rename filtered no-op snapshot should be available.");
    state.Require(filteredNoOpSnapshot.hideUnchangedRows, L"Batch Rename snapshot should report hide-unchanged mode.");
    state.Require(filteredNoOpSnapshot.previewRowCount == 1u, L"Hide unchanged should leave only changed rows visible.");
    state.Require(filteredNoOpSnapshot.originalNames == std::vector<std::wstring>{L"Alpha File.TXT"},
                  L"Hide unchanged should preserve the changed preview row.");
    state.Require(filteredNoOpSnapshot.newNames == std::vector<std::wstring>{L"Alpha Renamed File.TXT"},
                  L"Hide unchanged should preserve the changed row's proposed name.");
    state.Require(filteredNoOpSnapshot.changedRowCount == 1u && filteredNoOpSnapshot.warningRowCount == 1u,
                  L"Hide unchanged should not change underlying preview statistics.");

    state.Require(DebugSetBatchRenameWindowHideUnchanged(false), L"Batch Rename debug helper should restore unchanged preview rows.");

    BatchRenameDebugSnapshot restoredNoOpSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(restoredNoOpSnapshot), L"Batch Rename restored no-op snapshot should be available.");
    state.Require(! restoredNoOpSnapshot.hideUnchangedRows, L"Batch Rename snapshot should report visible unchanged rows after restore.");
    state.Require(restoredNoOpSnapshot.previewRowCount == 2u, L"Restoring unchanged rows should show the full preview again.");

    rules.nameTemplate = L"{unknown}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename debug rules hook should accept invalid rule input for validation.");

    BatchRenameDebugSnapshot invalidSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(invalidSnapshot), L"Invalid Batch Rename rules snapshot should be available.");
    state.Require(invalidSnapshot.errorRowCount == 2u, L"Unknown macro should be a blocking error on every preview row.");
    state.Require(invalidSnapshot.newNameStatusIconTexts.size() == 2u &&
                      std::ranges::all_of(invalidSnapshot.newNameStatusIconTexts, [](const std::wstring& iconText) noexcept { return ! iconText.empty(); }),
                  L"Blocking preview rows should show a status glyph in the New Name column.");
    state.Require(std::ranges::all_of(invalidSnapshot.newNameStatusIconTexts, [&warningStatusIcon](const std::wstring& iconText) noexcept {
                      return iconText != warningStatusIcon;
                  }),
                  L"Blocking preview rows should use an error status glyph distinct from the warning glyph.");
    state.Require(invalidSnapshot.newNameTooltips.size() == 2u &&
                      std::ranges::all_of(invalidSnapshot.newNameTooltips, [](const std::wstring& tooltip) noexcept {
                          return tooltip.find(L"macro_unknown") != std::wstring::npos;
                      }),
                  L"Blocking preview row tooltips should include the stable error issue id.");
    state.Require(! invalidSnapshot.renameButtonEnabled, L"Rename should be disabled while blocking preview errors are present.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowPreviewContextMenuCopiesRows(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameContextMenuSelfTest";

    uint32_t revealCalls = 0u;
    std::filesystem::path revealedPath;

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-context-menu-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Alpha.txt", root / L"Beta.md"};
    context.onRevealPath    = [&](const std::filesystem::path& path) noexcept
    {
        ++revealCalls;
        revealedPath = path;
        return true;
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-context-menu-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for preview context-menu copy testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename context-menu copy test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto clearClipboard = wil::scope_exit([batchWindow]() noexcept { ClearClipboardContents(batchWindow); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"{stem}_renamed{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename context-menu copy test should set preview rules.");

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename context-menu copy snapshot should be available.");
    state.Require(snapshot.newNames == std::vector<std::wstring>{L"Alpha_renamed.txt", L"Beta_renamed.md"},
                  L"Batch Rename context-menu copy preview should contain renamed rows.");

    const auto expectCopy = [&](const BatchRenameDebugPreviewCopyKind kind, const size_t rowIndex, const std::wstring_view expected, const wchar_t* label)
    {
        ClearClipboardContents(batchWindow);
        state.Require(DebugCopyBatchRenameWindowPreview(kind, rowIndex), std::format(L"{} copy command should succeed.", label));
        const std::wstring copied = ReadClipboardUnicodeText(batchWindow);
        state.Require(copied == expected,
                      std::format(L"{} copy command wrote unexpected text. Expected '{}', saw '{}'.", label, expected, copied));
    };

    expectCopy(BatchRenameDebugPreviewCopyKind::OriginalName, 0u, L"Alpha.txt", L"Batch Rename original-name");
    expectCopy(BatchRenameDebugPreviewCopyKind::NewName, 1u, L"Beta_renamed.md", L"Batch Rename new-name");
    expectCopy(BatchRenameDebugPreviewCopyKind::SourcePath, 0u, (root / L"Alpha.txt").native(), L"Batch Rename source-path");

    const std::wstring expectedPreviewRows =
        L"Original Name\tNew Name\tSize\tDate\tTime\tPath\r\n"
        L"Alpha.txt\tAlpha_renamed.txt\t0\t\t\t" +
        (root / L"Alpha.txt").native() +
        L"\r\n"
        L"Beta.md\tBeta_renamed.md\t0\t\t\t" +
        (root / L"Beta.md").native();
    expectCopy(BatchRenameDebugPreviewCopyKind::PreviewRows, 0u, expectedPreviewRows, L"Batch Rename preview-rows");

    state.Require(DebugRevealBatchRenameWindowPreview(1u), L"Batch Rename reveal command should invoke the preview-row reveal callback.");
    state.Require(revealCalls == 1u, std::format(L"Batch Rename reveal callback should run once; saw {}.", revealCalls));
    state.Require(revealedPath == root / L"Beta.md",
                  std::format(L"Batch Rename reveal callback should receive the selected source path; saw '{}'.", revealedPath.native()));

    state.Require(DebugActivateBatchRenameWindowPreview(0u), L"Batch Rename row activation should invoke the preview-row reveal callback.");
    state.Require(revealCalls == 2u, std::format(L"Batch Rename reveal callback should run twice after row activation; saw {}.", revealCalls));
    state.Require(revealedPath == root / L"Alpha.txt",
                  std::format(L"Batch Rename row activation should reveal the activated source path; saw '{}'.", revealedPath.native()));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowPreviewClipboardHonorsDisplayOrder(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameClipboardOrderSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-clipboard-order-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Alpha.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-clipboard-order-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for preview clipboard column-order testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename clipboard order test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto clearClipboard = wil::scope_exit([batchWindow]() noexcept { ClearClipboardContents(batchWindow); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"{stem}_renamed{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename clipboard order test should set preview rules.");
    state.Require(DebugReorderBatchRenameWindowPreviewColumn(L"path", 0u), L"Batch Rename clipboard order test should move Path first.");
    state.Require(DebugReorderBatchRenameWindowPreviewColumn(L"new", 1u), L"Batch Rename clipboard order test should move New Name second.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowPreview(BatchRenameDebugPreviewCopyKind::PreviewRows, 0u),
                  L"Batch Rename preview row copy should succeed after column reordering.");

    const std::wstring expected =
        L"Path\tNew Name\tOriginal Name\tSize\tDate\tTime\r\n" +
        (root / L"Alpha.txt").native() +
        L"\tAlpha_renamed.txt\tAlpha.txt\t0\t\t";
    const std::wstring copied = ReadClipboardUnicodeText(batchWindow);
    state.Require(copied == expected,
                  std::format(L"Preview-row TSV should follow display column order. Expected '{}', saw '{}'.", expected, copied));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowIgnoresStaleGenerationPayloads(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameStaleGenerationSelfTest";
    size_t successCallbackCount      = 0u;

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-stale-generation-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Alpha.txt"};
    context.onSuccessfulRename =
        [&](std::span<const std::filesystem::path> sourcePaths, std::span<const std::filesystem::path> targetPaths) noexcept
    {
        successCallbackCount += std::min(sourcePaths.size(), targetPaths.size());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-stale-generation-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for stale-generation testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename stale-generation test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"{stem}_renamed{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename stale-generation test should set preview rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename stale-generation test should capture the initial snapshot.");
    state.Require(before.originalNames == std::vector<std::wstring>{L"Alpha.txt"},
                  L"Batch Rename stale-generation test should start with the original target row.");

    state.Require(DebugInjectStaleBatchRenameWindowCollectionPayload(root / L"Injected.txt"),
                  L"Stale collection payload injection should run.");
    BatchRenameDebugSnapshot afterCollection{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterCollection), L"Batch Rename stale collection snapshot should be available.");
    state.Require(afterCollection.originalNames == before.originalNames && afterCollection.newNames == before.newNames,
                  L"Stale collection payload should not replace collected targets or preview rows.");

    state.Require(DebugInjectStaleBatchRenameWindowExecutionPayload(root / L"Alpha.txt", root / L"Alpha_renamed.txt"),
                  L"Stale execution payload injection should run.");
    BatchRenameDebugSnapshot afterExecution{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterExecution), L"Batch Rename stale execution snapshot should be available.");
    state.Require(afterExecution.originalNames == before.originalNames && afterExecution.newNames == before.newNames,
                  L"Stale execution payload should not refresh targets or preview rows.");
    state.Require(! afterExecution.hasExecutionReport, L"Stale execution payload should not store an execution report.");
    state.Require(successCallbackCount == 0u,
                  std::format(L"Stale execution payload should not invoke success callbacks; saw {}.", successCallbackCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowThemeAccessibilitySnapshot(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameAccessibilitySelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-accessibility-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Alpha.txt"};

    const AppTheme darkTheme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-accessibility-dark-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, darkTheme, std::move(context)),
                  L"Batch Rename window should open for accessibility snapshot testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename accessibility test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"{stem}_v2{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename accessibility test should enable the Rename button through a valid change.");

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename accessibility snapshot should be available.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::vector<std::wstring> expectedFocusableNames{
        L"Mask:",
        L"Include subdirectories",
        L"Files",
        L"Folders",
        L"Rules",
        L"Manual",
        L"New name:",
        L"Insert name template helper",
        L"Search for:",
        L"Insert regular expression helper",
        L"Replace with:",
        L"Insert replacement helper",
        L"Regular expression",
        L"Case sensitive",
        L"Whole words",
        L"Only once in each name",
        L"Exclude extension",
        L"File name:",
        L"Extension:",
        L"Batch Rename preview",
        L"Hide unchanged",
        L"Rename",
    };

    state.Require(snapshot.focusableAccessibleNames == expectedFocusableNames,
                  std::format(L"Batch Rename focusable accessible names should match Rules-mode tab order. Expected {} names, saw {}.",
                              expectedFocusableNames.size(),
                              snapshot.focusableAccessibleNames.size()));

    const AppTheme lightTheme = ResolveAppTheme(ThemeMode::Light, L"batch-rename-accessibility-light-selftest");
    UpdateBatchRenameWindowsTheme(lightTheme);
    PumpPendingMessages();

    BatchRenameDebugSnapshot lightSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(lightSnapshot), L"Batch Rename accessibility snapshot should remain available after theme update.");
    state.Require(lightSnapshot.ruleControlsVisible, L"Batch Rename rules controls should remain visible after theme update.");
    state.Require(lightSnapshot.focusableAccessibleNames == expectedFocusableNames,
                  L"Batch Rename focusable accessible names should remain stable after theme update.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowRuleControlsDrivePreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameControlsSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-controls-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Movie SAMPLE.MKV", root / L"Other SAMPLE.SRT"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-controls-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for rule-control testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename controls test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRenameDebugSnapshot initialSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(initialSnapshot), L"Batch Rename controls initial snapshot should be available.");
    state.Require(initialSnapshot.ruleControlsVisible, L"Batch Rename should expose visible rule controls above the preview grid.");
    state.Require(initialSnapshot.nameTemplateText == L"{name}", L"Name template control should start from the default rules.");
    state.Require(initialSnapshot.caseSensitive, L"Case-sensitive control should start enabled.");
    state.Require(initialSnapshot.fileNameCaseText == L"Do not change", L"File-name case dropdown should show the default option.");
    state.Require(initialSnapshot.extensionCaseText == L"Do not change", L"Extension case dropdown should show the default option.");

    BatchRename::Rules rules{};
    rules.nameTemplate        = L"{stem}{ext}";
    rules.searchFor           = L"sample";
    rules.replaceWith         = L"archive";
    rules.caseSensitive       = false;
    rules.replaceOnce         = true;
    rules.excludeExtension    = true;
    rules.fileNameCaseStyle   = BatchRename::CaseTransform::Mixed;
    rules.extensionCaseStyle  = BatchRename::CaseTransform::Lower;

    state.Require(DebugSetBatchRenameWindowRuleControls(rules),
                  L"Batch Rename debug helper should drive visible rule controls like user edits.");

    BatchRenameDebugSnapshot editedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(editedSnapshot), L"Batch Rename edited controls snapshot should be available.");
    state.Require(editedSnapshot.nameTemplateText == L"{stem}{ext}", L"Name template control should reflect edited text.");
    state.Require(editedSnapshot.searchForText == L"sample", L"Search control should reflect edited text.");
    state.Require(editedSnapshot.replaceWithText == L"archive", L"Replace control should reflect edited text.");
    state.Require(! editedSnapshot.caseSensitive, L"Case-sensitive checkbox should reflect edited state.");
    state.Require(editedSnapshot.replaceOnce, L"Replace-once checkbox should reflect edited state.");
    state.Require(editedSnapshot.excludeExtension, L"Exclude-extension checkbox should reflect edited state.");
    state.Require(editedSnapshot.fileNameCaseText == L"Mixed case", L"File-name case dropdown should reflect edited state.");
    state.Require(editedSnapshot.extensionCaseText == L"Lower case", L"Extension case dropdown should reflect edited state.");
    state.Require(editedSnapshot.newNames == std::vector<std::wstring>{L"Movie Archive.mkv", L"Other Archive.srt"},
                  L"Visible rule-control edits should recompute the preview through the engine.");
    state.Require(editedSnapshot.changedRowCount == 2u, L"Visible rule-control edits should update changed-row stats.");
    state.Require(editedSnapshot.errorRowCount == 0u, L"Visible rule-control edits should keep valid preview rows error-free.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowDebouncesTextPreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameDebounceSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-debounce-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt", root / L"beta.md"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-debounce-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for debounce testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename debounce test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRenameDebugSnapshot initialSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(initialSnapshot), L"Batch Rename debounce initial snapshot should be available.");
    state.Require(initialSnapshot.newNames == std::vector<std::wstring>{L"alpha.txt", L"beta.md"},
                  L"Debounce test should start from unchanged default preview names.");
    state.Require(! initialSnapshot.previewRebuildPending, L"Initial Batch Rename preview should not have a pending rebuild.");

    state.Require(DebugSetBatchRenameWindowRuleFieldText(BatchRenameDebugRuleField::NameTemplate, L"renamed_{counter:000}{ext}"),
                  L"Batch Rename debounce test should edit the real name-template field.");

    BatchRenameDebugSnapshot pendingSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(pendingSnapshot), L"Batch Rename debounce pending snapshot should be available.");
    state.Require(pendingSnapshot.nameTemplateText == L"renamed_{counter:000}{ext}",
                  L"Debounced field text should update immediately.");
    state.Require(pendingSnapshot.previewRebuildPending, L"Text edits should schedule a debounced preview rebuild.");
    state.Require(pendingSnapshot.newNames == std::vector<std::wstring>{L"alpha.txt", L"beta.md"},
                  L"Debounced preview should remain unchanged until the pending rebuild is flushed.");

    state.Require(DebugFlushBatchRenameWindowPendingPreview(), L"Batch Rename debounce test should flush the pending preview timer.");

    BatchRenameDebugSnapshot flushedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(flushedSnapshot), L"Batch Rename debounce flushed snapshot should be available.");
    state.Require(! flushedSnapshot.previewRebuildPending, L"Flushed debounce should clear the pending preview flag.");
    state.Require(flushedSnapshot.newNames == std::vector<std::wstring>{L"renamed_001.txt", L"renamed_002.md"},
                  L"Flushed debounce should recompute the preview with the latest field text.");
    state.Require(flushedSnapshot.changedRowCount == 2u, L"Flushed debounce should update preview stats.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowUsesAndPersistsSettings(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameSettingsSelfTest";

    const std::optional<Common::Settings::BatchRenameSettings> previousBatchRename = g_settings.batchRename;
    const auto previousPlacementIt = g_settings.windows.find(L"BatchRenameWindow");
    const std::optional<Common::Settings::WindowPlacement> previousPlacement =
        previousPlacementIt == g_settings.windows.end()
            ? std::nullopt
            : std::optional<Common::Settings::WindowPlacement>{previousPlacementIt->second};
    const auto restoreSettings = wil::scope_exit([previousBatchRename, previousPlacement]() {
        CloseBatchRenameWindowIfOpen();
        g_settings.batchRename = previousBatchRename;
        if (previousPlacement.has_value())
        {
            g_settings.windows[L"BatchRenameWindow"] = previousPlacement.value();
        }
        else
        {
            g_settings.windows.erase(L"BatchRenameWindow");
        }
    });

    CloseBatchRenameWindowIfOpen();
    g_settings.windows.erase(L"BatchRenameWindow");

    Common::Settings::BatchRenameSettings persisted{};
    persisted.lastRoot              = L"C:\\old-batch-root";
    persisted.recentMasks           = {L"*.mkv", L"*.srt"};
    persisted.recentNameTemplates   = {L"{stem}_persisted{ext}", L"{counter}_{name}"};
    persisted.recentSearchPatterns  = {L"sample", L"draft"};
    persisted.recentReplacePatterns = {L"archive", L"final"};
    persisted.includeSubdirectories = true;
    persisted.includeFiles          = true;
    persisted.includeFolders        = true;
    persisted.regexEnabled          = true;
    persisted.caseSensitive         = false;
    persisted.wholeWords            = true;
    persisted.replaceOnce           = true;
    persisted.excludeExtension      = true;
    persisted.flattenSeparator      = L"__";
    persisted.fileNameCaseStyle     = Common::Settings::BatchRenameCaseStyle::Mixed;
    persisted.extensionCaseStyle    = Common::Settings::BatchRenameCaseStyle::Lower;
    g_settings.batchRename          = persisted;

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-settings-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Movie SAMPLE.MKV", root / L"Other SAMPLE.SRT"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-settings-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for settings persistence testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename settings test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    BatchRenameDebugSnapshot initialSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(initialSnapshot), L"Batch Rename settings snapshot should be available.");
    state.Require(initialSnapshot.nameTemplateText == L"{stem}_persisted{ext}",
                  L"Batch Rename should seed the name-template control from persisted history.");
    state.Require(initialSnapshot.scopeMaskText == L"*.mkv", L"Batch Rename should seed the mask control from persisted history.");
    state.Require(initialSnapshot.includeSubdirectories, L"Batch Rename should restore the include-subdirectories scope option.");
    state.Require(initialSnapshot.includeFiles, L"Batch Rename should restore the include-files scope option.");
    state.Require(initialSnapshot.includeFolders, L"Batch Rename should restore the include-folders scope option.");
    state.Require(initialSnapshot.searchForText == L"sample", L"Batch Rename should seed the search control from persisted history.");
    state.Require(initialSnapshot.replaceWithText == L"archive", L"Batch Rename should seed the replace control from persisted history.");
    state.Require(initialSnapshot.regexEnabled, L"Batch Rename should restore the regex option.");
    state.Require(! initialSnapshot.caseSensitive, L"Batch Rename should restore the case-sensitive option.");
    state.Require(initialSnapshot.wholeWords, L"Batch Rename should restore the whole-words option.");
    state.Require(initialSnapshot.replaceOnce, L"Batch Rename should restore the replace-once option.");
    state.Require(initialSnapshot.excludeExtension, L"Batch Rename should restore the exclude-extension option.");
    state.Require(initialSnapshot.fileNameCaseText == L"Mixed case", L"Batch Rename should restore the file-name case style.");
    state.Require(initialSnapshot.extensionCaseText == L"Lower case", L"Batch Rename should restore the extension case style.");

    BatchRename::Rules editedRules{};
    editedRules.nameTemplate       = L"{stem}_saved{ext}";
    editedRules.searchFor          = L"movie";
    editedRules.replaceWith        = L"clip";
    editedRules.regexEnabled       = false;
    editedRules.caseSensitive      = false;
    editedRules.wholeWords         = false;
    editedRules.replaceOnce        = true;
    editedRules.excludeExtension   = true;
    editedRules.flattenSeparator   = L"--";
    editedRules.fileNameCaseStyle  = BatchRename::CaseTransform::Upper;
    editedRules.extensionCaseStyle = BatchRename::CaseTransform::Lower;
    state.Require(DebugSetBatchRenameWindowRuleControls(editedRules),
                  L"Batch Rename settings test should edit the visible rule controls.");
    state.Require(DebugSetBatchRenameWindowScope(L"*.saved", false, true, false),
                  L"Batch Rename settings test should edit the visible scope controls.");
    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename settings test should switch to manual mode before close.");
    state.Require(DebugSetBatchRenameWindowManualText(L"manual-one.txt\nmanual-two.txt"),
                  L"Batch Rename settings test should edit manual names before close.");

    CloseBatchRenameWindowIfOpen();
    state.Require(GetBatchRenameWindowHandle() == nullptr, L"Batch Rename settings test window should close.");
    state.Require(g_settings.batchRename.has_value(), L"Batch Rename should persist a settings object on close.");
    if (! g_settings.batchRename.has_value())
    {
        return false;
    }

    const Common::Settings::BatchRenameSettings& saved = g_settings.batchRename.value();
    state.Require(saved.lastRoot == root.native(), L"Batch Rename should persist the current root path.");
    state.Require(! saved.recentNameTemplates.empty() && saved.recentNameTemplates.front() == L"{stem}_saved{ext}",
                  L"Batch Rename should MRU-persist the edited name template.");
    state.Require(! saved.recentMasks.empty() && saved.recentMasks.front() == L"*.saved",
                  L"Batch Rename should MRU-persist the edited mask.");
    state.Require(! saved.recentSearchPatterns.empty() && saved.recentSearchPatterns.front() == L"movie",
                  L"Batch Rename should MRU-persist the edited search pattern.");
    state.Require(! saved.recentReplacePatterns.empty() && saved.recentReplacePatterns.front() == L"clip",
                  L"Batch Rename should MRU-persist the edited replacement pattern.");
    state.Require(saved.recentNameTemplates.size() < 2u || saved.recentNameTemplates[1] == L"{stem}_persisted{ext}",
                  L"Batch Rename should preserve older template history behind the newest entry.");
    state.Require(! std::ranges::contains(saved.recentNameTemplates, L"manual-one.txt"),
                  L"Batch Rename must not persist manual multiline names as history.");
    state.Require(! std::ranges::contains(saved.recentNameTemplates, L"manual-two.txt"),
                  L"Batch Rename must not persist manual multiline names as history.");
    state.Require(! saved.includeSubdirectories, L"Batch Rename should persist the edited include-subdirectories option.");
    state.Require(saved.includeFiles, L"Batch Rename should persist the edited include-files option.");
    state.Require(! saved.includeFolders, L"Batch Rename should persist the edited include-folders option.");
    state.Require(! saved.regexEnabled, L"Batch Rename should persist the edited regex option.");
    state.Require(! saved.caseSensitive, L"Batch Rename should persist the edited case-sensitive option.");
    state.Require(! saved.wholeWords, L"Batch Rename should persist the edited whole-words option.");
    state.Require(saved.replaceOnce, L"Batch Rename should persist the edited replace-once option.");
    state.Require(saved.excludeExtension, L"Batch Rename should persist the edited exclude-extension option.");
    state.Require(saved.flattenSeparator == L"--", L"Batch Rename should persist the flatten separator.");
    state.Require(saved.fileNameCaseStyle == Common::Settings::BatchRenameCaseStyle::Upper,
                  L"Batch Rename should persist the edited file-name case style.");
    state.Require(saved.extensionCaseStyle == Common::Settings::BatchRenameCaseStyle::Lower,
                  L"Batch Rename should persist the edited extension case style.");
    state.Require(g_settings.windows.contains(L"BatchRenameWindow"),
                  L"Batch Rename should persist window placement under settings.windows.BatchRenameWindow.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowManualModeControlsDrivePreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameManualControlsSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-manual-controls-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"two.md", root / L"one.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-manual-controls-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for manual-mode control testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename manual controls test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename debug helper should switch the visible window into manual mode.");

    BatchRenameDebugSnapshot manualSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(manualSnapshot), L"Batch Rename manual snapshot should be available.");
    state.Require(manualSnapshot.manualModeSelected, L"Manual mode selector should be selected after switching modes.");
    state.Require(! manualSnapshot.ruleControlsVisible, L"Rule controls should hide while manual mode is active.");
    state.Require(manualSnapshot.manualControlsVisible, L"Manual mode should expose the multiline manual names editor.");
    state.Require(manualSnapshot.manualText == L"two.md\none.txt",
                  L"Switching to manual mode should seed the editor from current preview names.");
    state.Require(manualSnapshot.newNames == std::vector<std::wstring>{L"two.md", L"one.txt"},
                  L"Seeded manual names should keep the initial preview unchanged.");
    state.Require(! manualSnapshot.renameButtonEnabled, L"Unchanged seeded manual names should not enable Rename.");

    state.Require(DebugSetBatchRenameWindowManualText(L"dos.md\nuno.txt"),
                  L"Batch Rename debug helper should edit the visible manual multiline field.");

    BatchRenameDebugSnapshot editedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(editedSnapshot), L"Batch Rename edited manual snapshot should be available.");
    state.Require(editedSnapshot.manualText == L"dos.md\nuno.txt", L"Manual editor text should reflect debug/user edits.");
    state.Require(editedSnapshot.newNames == std::vector<std::wstring>{L"dos.md", L"uno.txt"},
                  L"Manual editor lines should map one-for-one to preview rows.");
    state.Require(editedSnapshot.changedRowCount == 2u, L"Manual editor edits should update changed-row stats.");
    state.Require(editedSnapshot.errorRowCount == 0u, L"Valid manual editor lines should report no blocking errors.");
    state.Require(editedSnapshot.renameButtonEnabled, L"Valid changed manual names should enable Rename.");

    state.Require(DebugSetBatchRenameWindowPreviewSort(L"original", false),
                  L"Batch Rename debug helper should sort the preview grid by original name.");

    BatchRenameDebugSnapshot sortedPreviewSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(sortedPreviewSnapshot), L"Batch Rename sorted manual preview snapshot should be available.");
    state.Require(sortedPreviewSnapshot.originalNames == std::vector<std::wstring>{L"one.txt", L"two.md"},
                  L"Sorting by original name should reorder the visible preview rows.");
    state.Require(sortedPreviewSnapshot.newNames == std::vector<std::wstring>{L"uno.txt", L"dos.md"},
                  L"Preview sorting should keep each target paired with its current manual name.");
    state.Require(sortedPreviewSnapshot.manualText == L"dos.md\nuno.txt",
                  L"Preview sorting should not rewrite the manual editor until Sort like preview is pressed.");

    state.Require(DebugClickBatchRenameWindowManualSortLikePreview(),
                  L"Batch Rename Manual Sort like preview command should be invokable from selftests.");

    BatchRenameDebugSnapshot sortedManualSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(sortedManualSnapshot), L"Batch Rename sorted manual snapshot should be available.");
    state.Require(sortedManualSnapshot.manualText == L"uno.txt\ndos.md",
                  L"Sort like preview should rewrite manual lines into the current visible preview order.");
    state.Require(sortedManualSnapshot.originalNames == std::vector<std::wstring>{L"one.txt", L"two.md"},
                  L"Sort like preview should keep the sorted preview order visible.");
    state.Require(sortedManualSnapshot.newNames == std::vector<std::wstring>{L"uno.txt", L"dos.md"},
                  L"Sort like preview should keep each sorted target paired with its manual line.");

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Rules),
                  L"Batch Rename debug helper should switch back to Rules mode without discarding manual text.");
    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename debug helper should switch back to Manual mode with preserved manual text.");

    BatchRenameDebugSnapshot restoredManualSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(restoredManualSnapshot), L"Batch Rename restored manual snapshot should be available.");
    state.Require(restoredManualSnapshot.manualText == L"uno.txt\ndos.md",
                  L"Manual editor text should survive switching away and back while the target set is unchanged.");
    state.Require(restoredManualSnapshot.newNames == std::vector<std::wstring>{L"uno.txt", L"dos.md"},
                  L"Manual preview should survive switching away and back while the target set is unchanged.");

    BatchRename::Rules nonTargetRules{};
    nonTargetRules.nameTemplate = L"{stem}-rule{ext}";
    nonTargetRules.searchFor    = L"rule";
    nonTargetRules.replaceWith  = L"edited";
    state.Require(DebugSetBatchRenameWindowRuleControls(nonTargetRules),
                  L"Batch Rename debug helper should edit rule controls while manual text is preserved.");
    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename debug helper should return to Manual after non-target rule edits.");

    BatchRenameDebugSnapshot afterRuleEditSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterRuleEditSnapshot), L"Batch Rename post-rule-edit manual snapshot should be available.");
    state.Require(afterRuleEditSnapshot.manualText == L"uno.txt\ndos.md",
                  L"Manual editor text should survive non-target rule edits while the target set is unchanged.");
    state.Require(afterRuleEditSnapshot.newNames == std::vector<std::wstring>{L"uno.txt", L"dos.md"},
                  L"Manual preview should survive non-target rule edits while the target set is unchanged.");

    state.Require(DebugSetBatchRenameWindowManualText(L"only-one.txt"),
                  L"Batch Rename debug helper should accept mismatched manual line count for validation.");

    BatchRenameDebugSnapshot mismatchSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(mismatchSnapshot), L"Batch Rename mismatched manual snapshot should be available.");
    state.Require(mismatchSnapshot.errorRowCount == 2u, L"Manual line-count mismatch should block every preview row.");
    state.Require(! mismatchSnapshot.renameButtonEnabled, L"Manual line-count mismatch should disable Rename.");

    RedSalamander::DxUi::DebugClearClipboardFallbackText();
    state.Require(RedSalamander::DxUi::DebugWriteClipboardUnicodeText(batchWindow, L"paste-one.txt\r\npaste-two.md"),
                  L"Batch Rename manual paste test should seed multiline Unicode clipboard text.");
    state.Require(DebugClickBatchRenameWindowManualPaste(),
                  L"Batch Rename manual Paste command should be invokable from selftests.");

    BatchRenameDebugSnapshot pastedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(pastedSnapshot), L"Batch Rename pasted manual snapshot should be available.");
    state.Require(pastedSnapshot.manualText == L"paste-one.txt\r\npaste-two.md",
                  L"Manual Paste should preserve clipboard line breaks in the editor text.");
    state.Require(pastedSnapshot.newNames == std::vector<std::wstring>{L"paste-one.txt", L"paste-two.md"},
                  L"Manual Paste should map multiline clipboard text to one preview row per pasted line.");
    state.Require(pastedSnapshot.errorRowCount == 0u, L"Manual Paste with matching lines should clear manual validation errors.");
    state.Require(pastedSnapshot.renameButtonEnabled, L"Manual Paste with valid changed names should enable Rename.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowManualModeTargetChangeBlocksUntilReconciled(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_manual_scope_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename manual-scope root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for manual-scope test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for manual-scope test.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.md", "gamma"), L"Failed to create gamma.md for manual-scope test.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-manual-scope-selftest";
    context.rootPluginPath  = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-manual-scope-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for manual target-change testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename manual target-change test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSetBatchRenameWindowScope(L"*.txt", false, true, false),
                  L"Batch Rename manual target-change test should start from a two-file scope.");
    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename manual target-change test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"uno.txt\ndos.txt"),
                  L"Batch Rename manual target-change test should set valid two-line manual text.");

    BatchRenameDebugSnapshot twoTargetSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(twoTargetSnapshot), L"Batch Rename two-target manual snapshot should be available.");
    state.Require(twoTargetSnapshot.previewRowCount == 2u, L"Initial *.txt scope should contain two preview rows.");
    state.Require(twoTargetSnapshot.errorRowCount == 0u, L"Matching manual line count should have no blocking errors.");
    state.Require(twoTargetSnapshot.renameButtonEnabled, L"Changed matching manual names should enable Rename.");

    state.Require(DebugSetBatchRenameWindowScope(L"*.*", false, true, false),
                  L"Batch Rename manual target-change test should expand the scope while Manual mode is active.");

    BatchRenameDebugSnapshot expandedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(expandedSnapshot), L"Batch Rename expanded manual snapshot should be available.");
    state.Require(expandedSnapshot.previewRowCount == 3u, L"Expanded folder scope should contain three preview rows.");
    state.Require(expandedSnapshot.manualText == L"uno.txt\ndos.txt",
                  L"Manual text should be preserved when target collection changes underneath it.");
    state.Require(expandedSnapshot.errorRowCount == 3u,
                  L"Manual target-set changes should block every preview row until line count is reconciled.");
    state.Require(! expandedSnapshot.renameButtonEnabled,
                  L"Manual target-set mismatch should disable Rename until the user reconciles the line count.");

    state.Require(DebugSetBatchRenameWindowManualText(L"uno.txt\ndos.txt\ntres.md"),
                  L"Batch Rename manual target-change test should accept reconciled three-line manual text.");

    BatchRenameDebugSnapshot reconciledSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(reconciledSnapshot), L"Batch Rename reconciled manual snapshot should be available.");
    state.Require(reconciledSnapshot.previewRowCount == 3u, L"Reconciled manual snapshot should keep the expanded target set.");
    state.Require(reconciledSnapshot.errorRowCount == 0u, L"Reconciled manual line count should clear blocking errors.");
    state.Require(reconciledSnapshot.renameButtonEnabled, L"Reconciled changed manual names should re-enable Rename.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesLocalRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_execute_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename execution root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create first Batch Rename execution input.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create second Batch Rename execution input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename execution.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-execute-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt", root / L"beta.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-execute-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for local execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename execution test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename execution test should set valid rename rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename execution preview should be available before executing.");
    state.Require(before.renameButtonEnabled, L"Valid changed preview should enable execution.");
    state.Require(before.newNames == std::vector<std::wstring>{L"renamed_001.txt", L"renamed_002.txt"},
                  L"Execution preview should propose the expected new names.");

    const uint64_t executeDurationBefore = CountBatchRenamePerfRowsWithMetric("batchrename.execute.us");
    const uint64_t executeRowsBefore     = CountBatchRenamePerfRowsWithMetric("batchrename.execute.rows");
    const uint64_t executeDoneBefore     = CountBatchRenamePerfRowsWithMetric("batchrename.execute.completed");
    const uint64_t executeFailedBefore   = CountBatchRenamePerfRowsWithMetric("batchrename.execute.failed");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    const uint64_t executeDurationAfter = CountBatchRenamePerfRowsWithMetric("batchrename.execute.us");
    const uint64_t executeRowsAfter     = CountBatchRenamePerfRowsWithMetric("batchrename.execute.rows");
    const uint64_t executeDoneAfter     = CountBatchRenamePerfRowsWithMetric("batchrename.execute.completed");
    const uint64_t executeFailedAfter   = CountBatchRenamePerfRowsWithMetric("batchrename.execute.failed");
    state.Require(executeDurationAfter > executeDurationBefore, L"Batch Rename execution should emit batchrename.execute.us.");
    state.Require(executeRowsAfter > executeRowsBefore, L"Batch Rename execution should emit batchrename.execute.rows.");
    state.Require(executeDoneAfter > executeDoneBefore, L"Batch Rename execution should emit batchrename.execute.completed.");
    state.Require(executeFailedAfter > executeFailedBefore, L"Batch Rename execution should emit batchrename.execute.failed.");

    state.Require(! SelfTest::PathExists(root / L"alpha.txt"), L"Batch Rename execution should remove the first original path.");
    state.Require(! SelfTest::PathExists(root / L"beta.txt"), L"Batch Rename execution should remove the second original path.");
    state.Require(SelfTest::PathExists(root / L"renamed_001.txt"), L"Batch Rename execution should create the first renamed path.");
    state.Require(SelfTest::PathExists(root / L"renamed_002.txt"), L"Batch Rename execution should create the second renamed path.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename execution preview should be available after executing.");
    state.Require(after.originalNames == std::vector<std::wstring>{L"renamed_001.txt", L"renamed_002.txt"},
                  L"Successful execution should refresh the window target list to the renamed paths.");
    state.Require(after.changedRowCount == 0u, L"Successful execution should leave the refreshed preview with no pending changes.");
    state.Require(! after.renameButtonEnabled, L"Successful execution should disable Rename after refreshing changed rows.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowRefreshesPaneAfterSuccess([[maybe_unused]] HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_pane_refresh_" + NewGuidText());
    if (! SetupBatchRenamePaneFixture(state, root, {L"pane-alpha.batch", L"pane-beta.batch"}, {}))
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"pane-alpha.batch"),
                  L"Failed to focus first file for pane refresh Batch Rename test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left,
        [](std::wstring_view name) noexcept { return name == L"pane-alpha.batch" || name == L"pane-beta.batch"; },
        true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u,
                  L"Batch Rename pane refresh fixture should have exactly two selected files.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t refreshBefore = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    g_folderWindow.CommandBatchRename(FolderWindow::Pane::Left);
    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Pane command should open Batch Rename for refresh testing.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    BatchRename::Rules rules{};
    rules.nameTemplate = L"pane-refreshed-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename pane refresh test should set valid rename rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename pane refresh preview should be available before executing.");
    state.Require(before.renameButtonEnabled, L"Batch Rename pane refresh preview should enable Rename.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename pane refresh execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) > refreshBefore,
                  L"Successful Batch Rename execution launched from a pane should refresh that pane.");
    state.Require(WaitForPaneDisplayNames(FolderWindow::Pane::Left,
                                          {L"pane-refreshed-001.batch", L"pane-refreshed-002.batch"},
                                          {L"pane-alpha.batch", L"pane-beta.batch"},
                                          SelfTest::Scale(std::chrono::milliseconds{5000})),
                  L"Successful Batch Rename execution should update the visible pane display names.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowInvokesSuccessCallback(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_success_callback_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename success-callback root.");
    state.Require(SelfTest::WriteTextFile(root / L"callback-alpha.txt", "alpha"), L"Failed to create first Batch Rename success-callback input.");
    state.Require(SelfTest::WriteTextFile(root / L"callback-beta.txt", "beta"), L"Failed to create second Batch Rename success-callback input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename success-callback: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename success-callback execution.");
    if (! fileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-success-callback-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"callback-alpha.txt", root / L"callback-beta.txt"};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-success-callback-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for success-callback testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename success-callback test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"callback-renamed-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename success-callback test should set valid rename rules.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename success-callback execution should succeed: 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(callbackCalls == 1u, std::format(L"Batch Rename should invoke the success callback once; saw {}.", callbackCalls));
    state.Require(callbackSources == std::vector<std::filesystem::path>{root / L"callback-alpha.txt", root / L"callback-beta.txt"},
                  L"Batch Rename success callback should report original source paths in execution order.");
    state.Require(callbackTargets == std::vector<std::filesystem::path>{root / L"callback-renamed-001.txt", root / L"callback-renamed-002.txt"},
                  L"Batch Rename success callback should report target paths in execution order.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowSuccessCallbackParentChildExecutionOrder(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root        = suiteRoot / L"work" / (L"batch_rename_success_callback_parent_child_" + NewGuidText());
    const std::filesystem::path sourceDir   = root / L"scope";
    const std::filesystem::path sourceChild = sourceDir / L"item.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create Batch Rename parent/child callback directory.");
    state.Require(SelfTest::WriteTextFile(sourceChild, "child"), L"Failed to create Batch Rename parent/child callback file.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename parent/child callback execution.");
    if (! fileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-parent-child-callback-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {sourceDir, sourceChild};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-parent-child-callback-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for parent/child callback testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename parent/child callback test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{index}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename parent/child callback test should set valid rename rules.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename parent/child callback execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    const std::filesystem::path renamedDir   = root / L"renamed_0";
    const std::filesystem::path renamedChild = renamedDir / L"renamed_1.txt";
    state.Require(SelfTest::PathExists(renamedDir), L"Parent/child callback execution should create the renamed parent directory.");
    state.Require(SelfTest::PathExists(renamedChild), L"Parent/child callback execution should create the renamed child under the renamed parent.");

    state.Require(callbackCalls == 1u, std::format(L"Batch Rename should invoke the success callback once; saw {}.", callbackCalls));
    state.Require(callbackSources == std::vector<std::filesystem::path>{sourceChild, sourceDir},
                  L"Parent/child success callback should report original source paths in execution order (deepest first).");
    // Contract pin (see onSuccessfulRename in BatchRenameWindow.h): targets are each
    // row's path as of its own rename, so the child's pair is intentionally NOT folded
    // through the parent's later rename even though the file now lives under the
    // renamed parent (asserted via PathExists above). Consumers replay the pairs
    // sequentially to compute final paths.
    state.Require(callbackTargets == std::vector<std::filesystem::path>{sourceDir / L"renamed_1.txt", root / L"renamed_0"},
                  L"Parent/child success callback should report as-of-execution target paths (child under the pre-rename parent).");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowParentChildDirectoryCacheNotifyRetargetsPinnedDescendant(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root            = suiteRoot / L"work" / (L"batch_rename_cache_parent_child_" + NewGuidText());
    const std::filesystem::path sourceParent    = root / L"scope";
    const std::filesystem::path sourceChild     = sourceParent / L"child";
    const std::filesystem::path pinnedDescendant = sourceChild / L"pinned";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(pinnedDescendant), L"Failed to create Batch Rename cache-notify descendant directory.");
    state.Require(SelfTest::WriteTextFile(pinnedDescendant / L"inside.txt", "inside"),
                  L"Failed to create Batch Rename cache-notify descendant file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const DirectoryInfoCache::Stats cacheStatsBefore       = DirectoryInfoCache::GetInstance().GetStats();
    const auto restorePane                                 = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
        DirectoryInfoCache::GetInstance().SetLimits(cacheStatsBefore.maxBytes, cacheStatsBefore.maxWatchers, cacheStatsBefore.mruWatched);
    });
    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    DirectoryInfoCache::GetInstance().SetLimits(cacheStatsBefore.maxBytes, 0u, 0u);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set right pane local file-system plugin for Batch Rename cache-notify test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, pinnedDescendant);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, pinnedDescendant, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane did not pin the Batch Rename descendant directory.");
    state.Require(WaitForPaneDisplayNames(FolderWindow::Pane::Right,
                                          {L"inside.txt"},
                                          {},
                                          SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane did not enumerate the pinned descendant directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = g_folderWindow.GetFileSystem(FolderWindow::Pane::Right);
    state.Require(static_cast<bool>(fileSystem), L"Local file-system plugin should be available for Batch Rename cache-notify test.");
    if (! fileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-cache-parent-child-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {sourceParent, sourceChild};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-cache-parent-child-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for parent/child cache-notify testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename cache-notify test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{index}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename cache-notify test should set valid rename rules.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename parent/child cache-notify execution should succeed: 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    const std::filesystem::path renamedParent     = root / L"renamed_0";
    const std::filesystem::path renamedChild      = renamedParent / L"renamed_1";
    const std::filesystem::path renamedDescendant = renamedChild / L"pinned";
    state.Require(SelfTest::PathExists(renamedDescendant / L"inside.txt"),
                  L"Parent/child directory Batch Rename should preserve the descendant file under the final path.");
    state.Require(! SelfTest::PathExists(pinnedDescendant),
                  L"Parent/child directory Batch Rename should remove the original descendant path.");

    state.Require(callbackCalls == 1u, std::format(L"Batch Rename cache-notify test should invoke the success callback once; saw {}.", callbackCalls));
    state.Require(callbackSources == std::vector<std::filesystem::path>{sourceChild, sourceParent},
                  L"Parent/child directory success callback should keep raw deepest-first source order.");
    state.Require(callbackTargets == std::vector<std::filesystem::path>{sourceParent / L"renamed_1", renamedParent},
                  L"Parent/child directory success callback should keep raw as-of-execution targets.");

    state.Require(WaitForPanePath(FolderWindow::Pane::Right, renamedDescendant, SelfTest::Scale(std::chrono::milliseconds{5000})),
                  std::format(L"DirectoryInfoCache should retarget the pinned descendant pane to '{}'; current '{}'.",
                              renamedDescendant.wstring(),
                              g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right).value_or(std::filesystem::path{}).wstring()));
    state.Require(WaitForPaneDisplayNames(FolderWindow::Pane::Right,
                                          {L"inside.txt"},
                                          {},
                                          SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Retargeted descendant pane should enumerate the final directory contents.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowReportsExecutionSummary(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_report_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename report root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create first Batch Rename report input.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create unchanged Batch Rename report input.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma"), L"Failed to create second Batch Rename report input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename report: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename report execution.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-report-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt", root / L"beta.txt", root / L"gamma.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-report-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for execution report testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename report test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename report test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"renamed-alpha.txt\nbeta.txt\nrenamed-gamma.txt"),
                  L"Batch Rename report test should set manual names with one unchanged row.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename report preview should be available before executing.");
    state.Require(before.previewRowCount == 3u, L"Batch Rename report preview should include all three targets.");
    state.Require(before.changedRowCount == 2u, L"Batch Rename report preview should have two changed rows.");
    state.Require(before.renameButtonEnabled, L"Batch Rename report preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename report execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename report snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Batch Rename should retain an execution report after execution.");
    state.Require(after.lastExecutionTotalRows == 3u,
                  std::format(L"Batch Rename report should record three planned rows; saw {}.", after.lastExecutionTotalRows));
    state.Require(after.lastExecutionCompletedRows == 2u,
                  std::format(L"Batch Rename report should record two completed rows; saw {}.", after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionSkippedRows == 1u,
                  std::format(L"Batch Rename report should record one skipped no-op row; saw {}.", after.lastExecutionSkippedRows));
    state.Require(after.lastExecutionFailedRows == 0u,
                  std::format(L"Batch Rename report should record zero failed rows; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionUndoRowCount == 2u,
                  std::format(L"Batch Rename undo report should include only the two successful rows; saw {}.", after.lastExecutionUndoRowCount));
    state.Require(! after.lastExecutionCanceled, L"Successful Batch Rename report should not be marked canceled.");
    state.Require(SUCCEEDED(after.lastExecutionFirstFailure),
                  std::format(L"Successful Batch Rename report should have no first failure; saw 0x{:08X}.",
                              static_cast<unsigned long>(after.lastExecutionFirstFailure)));
    state.Require(after.lastExecutionFirstFailureText.empty(), L"Successful Batch Rename report should not carry a first-failure message.");
    state.Require(after.statusText.contains(L"2 renamed") && after.statusText.contains(L"1 skipped") && after.statusText.contains(L"0 failed"),
                  std::format(L"Batch Rename status should show execution summary; saw '{}'.", after.statusText));

    state.Require(! SelfTest::PathExists(root / L"alpha.txt"), L"Batch Rename report execution should remove first original path.");
    state.Require(SelfTest::PathExists(root / L"beta.txt"), L"Batch Rename report execution should keep unchanged path.");
    state.Require(! SelfTest::PathExists(root / L"gamma.txt"), L"Batch Rename report execution should remove second original path.");
    state.Require(SelfTest::PathExists(root / L"renamed-alpha.txt"), L"Batch Rename report execution should create first renamed path.");
    state.Require(SelfTest::PathExists(root / L"renamed-gamma.txt"), L"Batch Rename report execution should create second renamed path.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Batch Rename should copy a retained undo report after successful execution.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains(L"Current Path\tRestore Name\tOriginal Path"),
                  std::format(L"Batch Rename undo report should include TSV headers; saw '{}'.", undoReport));
    state.Require(undoReport.contains((root / L"renamed-alpha.txt").native()) && undoReport.contains(L"\talpha.txt\t") &&
                      undoReport.contains((root / L"alpha.txt").native()),
                  std::format(L"Batch Rename undo report should include the first successful inverse row; saw '{}'.", undoReport));
    state.Require(undoReport.contains((root / L"renamed-gamma.txt").native()) && undoReport.contains(L"\tgamma.txt\t") &&
                      undoReport.contains((root / L"gamma.txt").native()),
                  std::format(L"Batch Rename undo report should include the second successful inverse row; saw '{}'.", undoReport));
    state.Require(! undoReport.contains(L"beta.txt\tbeta.txt"), L"Batch Rename undo report should exclude skipped no-op rows.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameCollisionExistingAndDuplicateNames(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_collision_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename collision root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create first Batch Rename collision input.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create second Batch Rename collision input.");
    state.Require(SelfTest::WriteTextFile(root / L"occupied_001.txt", "occupied"), L"Failed to create Batch Rename existing destination.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename collision: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename collision validation.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-collision-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt", root / L"beta.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-collision-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for collision validation testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename collision test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules duplicateRules{};
    duplicateRules.nameTemplate = L"duplicate.txt";
    state.Require(DebugSetBatchRenameWindowRules(duplicateRules), L"Batch Rename collision test should set duplicate rules.");

    BatchRenameDebugSnapshot duplicateSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(duplicateSnapshot), L"Batch Rename duplicate collision snapshot should be available.");
    state.Require(duplicateSnapshot.errorRowCount == 2u,
                  std::format(L"Duplicate proposed names should block both preview rows; saw {} errors.",
                              duplicateSnapshot.errorRowCount));
    state.Require(! duplicateSnapshot.renameButtonEnabled, L"Duplicate proposed names should disable Rename.");

    BatchRename::Rules existingRules{};
    existingRules.nameTemplate = L"occupied_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(existingRules), L"Batch Rename collision test should set existing-destination rules.");

    BatchRenameDebugSnapshot existingSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(existingSnapshot), L"Batch Rename existing-destination snapshot should be available.");
    state.Require(existingSnapshot.newNames == std::vector<std::wstring>{L"occupied_001.txt", L"occupied_002.txt"},
                  L"Existing-destination collision test should propose deterministic occupied names.");
    state.Require(existingSnapshot.errorRowCount == 1u,
                  std::format(L"Existing unselected local destination should block the colliding preview row; saw {} errors.",
                              existingSnapshot.errorRowCount));
    state.Require(! existingSnapshot.renameButtonEnabled, L"Existing unselected local destination should disable Rename before execution.");
    state.Require(existingSnapshot.newNameTooltips.size() == 2u &&
                      existingSnapshot.newNameTooltips[0].contains(L"name_destination_exists") &&
                      ! existingSnapshot.newNameTooltips[1].contains(L"name_destination_exists"),
                  L"Existing-destination collision should expose the stable name_destination_exists issue ID only on the colliding row.");

    state.Require(SelfTest::PathExists(root / L"alpha.txt"), L"Collision preview validation should not mutate the first source.");
    state.Require(SelfTest::PathExists(root / L"beta.txt"), L"Collision preview validation should not mutate the second source.");
    state.Require(SelfTest::PathExists(root / L"occupied_001.txt"), L"Collision preview validation should preserve the existing destination.");
    state.Require(! SelfTest::PathExists(root / L"occupied_002.txt"), L"Collision preview validation should not create a second destination.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameDestinationProbeFailureIssueId(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_probe_failure_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename probe-failure root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create Batch Rename probe-failure input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-probe-failure-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-probe-failure-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for destination-probe failure testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename probe-failure test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupProbeFault  = wil::scope_exit([]() noexcept { DebugClearBatchRenameWindowDestinationProbeFailurePath(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const std::filesystem::path destination = root / L"probe_001.txt";
    DebugSetBatchRenameWindowDestinationProbeFailurePath(destination, ERROR_ACCESS_DENIED);

    BatchRename::Rules rules{};
    rules.nameTemplate = L"probe_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename probe-failure test should set destination-probing rules.");

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename probe-failure snapshot should be available.");
    state.Require(snapshot.newNames == std::vector<std::wstring>{L"probe_001.txt"},
                  L"Destination-probe failure test should propose the injected destination path.");
    state.Require(snapshot.errorRowCount == 1u,
                  std::format(L"Destination-probe failure should block the row; saw {} errors.", snapshot.errorRowCount));
    state.Require(! snapshot.renameButtonEnabled, L"Destination-probe failure should disable Rename.");
    state.Require(snapshot.newNameTooltips.size() == 1u && snapshot.newNameTooltips[0].contains(L"name_destination_probe_failed"),
                  L"Destination-probe failure should expose the stable name_destination_probe_failed issue ID.");

    state.Require(SelfTest::PathExists(root / L"alpha.txt"), L"Probe-failure preview validation should not mutate the source.");
    state.Require(! SelfTest::PathExists(destination), L"Probe-failure preview validation should not create the destination.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameCancelDoesNotApplyRemainingRows(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_cancel_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename cancel root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create first Batch Rename cancel input.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create second Batch Rename cancel input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for Batch Rename cancel: 0x{:08X}.",
                                  static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename cancel validation.");
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    std::atomic_uint32_t shouldCancelCalls{0u};
    wil::com_ptr<IFileSystem> cancelFileSystem =
        CreateBatchRenameCancelOnShouldCancelFileSystem(fileSystem, &renameItemsCalls, &shouldCancelCalls);
    state.Require(cancelFileSystem != nullptr, L"Batch Rename cancel selftest should create a cancel-aware file-system wrapper.");
    if (! cancelFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = cancelFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-cancel-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"alpha.txt", root / L"beta.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-cancel-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for cancel validation testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename cancel test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"cancelled_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename cancel test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename cancel preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid cancel preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Batch Rename cancel execution should return ERROR_CANCELLED; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  L"Batch Rename cancel test should call provider RenameItems exactly once.");
    state.Require(shouldCancelCalls.load(std::memory_order_relaxed) == 1u,
                  L"Batch Rename execution should provide a FileSystemShouldCancel callback to the provider.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename cancel snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Canceled Batch Rename should retain an execution report.");
    state.Require(after.lastExecutionCanceled, L"Canceled Batch Rename report should mark canceled state.");
    state.Require(after.lastExecutionFailedRows == 2u,
                  std::format(L"Canceled Batch Rename should report remaining failed rows; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionCompletedRows == 0u,
                  std::format(L"Canceled Batch Rename should not report completed rows; saw {}.", after.lastExecutionCompletedRows));

    state.Require(SelfTest::PathExists(root / L"alpha.txt"), L"Canceled Batch Rename should preserve first source.");
    state.Require(SelfTest::PathExists(root / L"beta.txt"), L"Canceled Batch Rename should preserve second source.");
    state.Require(! SelfTest::PathExists(root / L"cancelled_001.txt"), L"Canceled Batch Rename should not create first destination.");
    state.Require(! SelfTest::PathExists(root / L"cancelled_002.txt"), L"Canceled Batch Rename should not create second destination.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowExecutionReport(), L"Batch Rename should copy a retained execution report after cancellation.");
    const std::wstring executionReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(executionReport.contains(L"Total Rows\tCompleted Rows\tSkipped Rows\tFailed Rows\tCanceled\tFirst Failure"),
                  std::format(L"Batch Rename execution report should include TSV headers; saw '{}'.", executionReport));
    state.Require(executionReport.contains(L"2\t0\t0\t2\ttrue\t"),
                  std::format(L"Canceled Batch Rename execution report should include row counts and canceled state; saw '{}'.", executionReport));
    state.Require(executionReport.contains(L"0x800704C7"),
                  std::format(L"Canceled Batch Rename execution report should include the cancellation HRESULT; saw '{}'.", executionReport));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowFallsBackToRenameItemWhenBulkUnsupported(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_fallback_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename fallback root.");
    state.Require(SelfTest::WriteTextFile(root / L"first.txt", "first"), L"Failed to create first Batch Rename fallback input.");
    state.Require(SelfTest::WriteTextFile(root / L"second.txt", "second"), L"Failed to create second Batch Rename fallback input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin for Batch Rename fallback: 0x{:08X}.",
                                                       static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename fallback execution.");
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemCalls{0u};
    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> fallbackFileSystem = CreateBatchRenameBulkUnsupportedFileSystem(fileSystem, &renameItemCalls, &renameItemsCalls);
    state.Require(fallbackFileSystem != nullptr, L"Batch Rename fallback selftest should create a bulk-unsupported file-system wrapper.");
    if (! fallbackFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fallbackFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-fallback-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"first.txt", root / L"second.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-fallback-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for RenameItem fallback testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename fallback test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"fallback-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename fallback test should set valid rename rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename fallback preview should be available before executing.");
    state.Require(before.renameButtonEnabled, L"Fallback preview should be executable.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename fallback execution should succeed, hr=0x{:08X}.", static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  L"Batch Rename fallback should try provider bulk RenameItems exactly once.");
    state.Require(renameItemCalls.load(std::memory_order_relaxed) == 2u,
                  L"Batch Rename fallback should execute one RenameItem call per changed row after unsupported bulk rename.");

    state.Require(! std::filesystem::exists(root / L"first.txt"), L"Fallback execution should remove first original path.");
    state.Require(! std::filesystem::exists(root / L"second.txt"), L"Fallback execution should remove second original path.");
    state.Require(std::filesystem::exists(root / L"fallback-001.txt"), L"Fallback execution should create first renamed path.");
    state.Require(std::filesystem::exists(root / L"fallback-002.txt"), L"Fallback execution should create second renamed path.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename fallback preview should refresh after executing.");
    state.Require(after.originalNames == std::vector<std::wstring>{L"fallback-001.txt", L"fallback-002.txt"},
                  L"Fallback execution should refresh preview targets to renamed paths.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesCaseOnlyLocalRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root   = suiteRoot / L"work" / (L"batch_rename_case_only_" + NewGuidText());
    const std::filesystem::path source = root / L"MiXeD.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename case-only root.");
    state.Require(SelfTest::WriteTextFile(source, "case"), L"Failed to create Batch Rename case-only input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename case-only execution.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-case-only-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {source};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-case-only-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for case-only execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename case-only execution test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate       = L"{name}";
    rules.fileNameCaseStyle  = BatchRename::CaseTransform::Upper;
    rules.extensionCaseStyle = BatchRename::CaseTransform::Upper;
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename case-only execution test should set upper-case rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename case-only preview should be available before executing.");
    state.Require(before.renameButtonEnabled, L"Case-only preview should be executable.");
    state.Require(before.newNames == std::vector<std::wstring>{L"MIXED.TXT"}, L"Case-only preview should propose upper-case leaf casing.");
    state.Require(before.errorRowCount == 0u, L"Case-only preview should not report a destination conflict.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename case-only execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    std::vector<std::wstring> leaves;
    for (std::filesystem::directory_iterator it(root, ec), end; it != end && ! ec; it.increment(ec))
    {
        leaves.push_back(it->path().filename().wstring());
    }
    state.Require(! ec, L"Batch Rename case-only test should enumerate the destination directory.");
    state.Require(std::ranges::find(leaves, L"MIXED.TXT") != leaves.end(), L"Case-only execution should preserve the requested upper-case leaf.");
    state.Require(std::ranges::find(leaves, L"MiXeD.txt") == leaves.end(), L"Case-only execution should not leave the original casing in the directory entry.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename case-only preview should refresh after executing.");
    state.Require(after.originalNames == std::vector<std::wstring>{L"MIXED.TXT"},
                  L"Case-only execution should refresh the preview to the new leaf casing.");
    state.Require(after.changedRowCount == 0u, L"Case-only execution should leave no pending changed rows after refresh.");
    state.Require(! after.renameButtonEnabled, L"Case-only execution should disable Rename after refreshing changed rows.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowBlocksInvalidPreviewExecution(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_invalid_execute_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path firstSource  = root / L"alpha.txt";
    const std::filesystem::path secondSource = root / L"beta.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename invalid-preview root.");
    state.Require(SelfTest::WriteTextFile(firstSource, "alpha"), L"Failed to create first Batch Rename invalid-preview input.");
    state.Require(SelfTest::WriteTextFile(secondSource, "beta"), L"Failed to create second Batch Rename invalid-preview input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename invalid-preview execution.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-invalid-execute-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {firstSource, secondSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-invalid-execute-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for invalid-preview execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename invalid-preview execution window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules invalidRules{};
    invalidRules.nameTemplate = L"{unknown_macro}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(invalidRules), L"Batch Rename invalid-preview execution test should set invalid rules.");

    BatchRenameDebugSnapshot invalidPreview{};
    state.Require(DebugGetBatchRenameWindowSnapshot(invalidPreview), L"Batch Rename invalid-preview snapshot should be available.");
    state.Require(invalidPreview.errorRowCount == 2u, L"Invalid preview should report blocking errors on every row.");
    state.Require(! invalidPreview.renameButtonEnabled, L"Invalid preview should disable Rename before execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(FAILED(executeHr), L"Batch Rename execution should reject blocking preview errors before provider dispatch.");
    state.Require(SelfTest::PathExists(firstSource), L"Invalid preview execution should leave the first source untouched.");
    state.Require(SelfTest::PathExists(secondSource), L"Invalid preview execution should leave the second source untouched.");
    state.Require(! SelfTest::PathExists(root / L".txt"), L"Invalid preview execution should not create a malformed destination.");
    state.Require(! SelfTest::PathExists(root / L"unknown_macro.txt"), L"Invalid preview execution should not create a guessed destination.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowBlocksProviderWithoutPathIdentity(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_missing_identity_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path source = root / L"provider-alpha.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename missing-identity root.");
    state.Require(SelfTest::WriteTextFile(source, "provider-alpha"), L"Failed to create Batch Rename missing-identity input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    constexpr std::string_view kMissingPathIdentityCapabilities = R"json({
  "version": 1,
  "operations": {
    "rename": true
  },
  "concurrency": {},
  "crossFileSystem": {}
})json";
    wil::com_ptr<IFileSystem> missingIdentityFileSystem =
        CreateBatchRenameCapabilitiesOverrideFileSystem(fileSystem, &renameItemsCalls, std::string(kMissingPathIdentityCapabilities));
    state.Require(missingIdentityFileSystem != nullptr,
                  L"Batch Rename missing-identity selftest should create a capabilities override filesystem.");
    if (! missingIdentityFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = missingIdentityFileSystem;
    context.pluginId        = L"selftest/missing-path-identity";
    context.pluginShortId   = L"missing-identity";
    context.instanceContext = L"batch-rename-missing-identity-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {source};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-missing-identity-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for missing provider path identity testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename missing-identity window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename missing-identity test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"provider-renamed.txt"),
                  L"Batch Rename missing-identity test should set a changed manual name.");

    BatchRenameDebugSnapshot preview{};
    state.Require(DebugGetBatchRenameWindowSnapshot(preview), L"Batch Rename missing-identity preview should be available.");
    state.Require(preview.previewRowCount == 1u, L"Missing-identity preview should still show the provider row.");
    state.Require(preview.changedRowCount == 1u, L"Missing-identity preview should preserve the proposed rename.");
    state.Require(preview.errorRowCount == 1u, L"Missing provider path identity should block the preview row.");
    state.Require(! preview.renameButtonEnabled, L"Missing provider path identity should disable Rename.");
    if (! preview.newNameTooltips.empty())
    {
        state.Require(preview.newNameTooltips[0].contains(L"provider_path_identity_unknown"),
                      std::format(L"Missing provider path identity should surface provider_path_identity_unknown; saw '{}'.",
                                  preview.newNameTooltips[0]));
    }

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(FAILED(executeHr), L"Missing provider path identity should reject execution before provider dispatch.");
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 0u,
                  std::format(L"Missing provider path identity should not call RenameItems; saw {} calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));
    state.Require(SelfTest::PathExists(source), L"Missing provider path identity should leave the source untouched.");
    state.Require(! SelfTest::PathExists(root / L"provider-renamed.txt"),
                  L"Missing provider path identity should not create the requested destination.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesParentChildDeepestFirst(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root        = suiteRoot / L"work" / (L"batch_rename_parent_child_" + NewGuidText());
    const std::filesystem::path sourceDir   = root / L"scope";
    const std::filesystem::path sourceChild = sourceDir / L"item.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create Batch Rename parent/child directory.");
    state.Require(SelfTest::WriteTextFile(sourceChild, "child"), L"Failed to create Batch Rename parent/child file.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename parent/child execution.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-parent-child-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {sourceDir, sourceChild};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-parent-child-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for parent/child execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename parent/child execution window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{index}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename parent/child execution test should set valid rename rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename parent/child preview should be available before executing.");
    state.Require(before.renameButtonEnabled, L"Valid parent/child preview should enable execution.");
    state.Require(before.newNames == std::vector<std::wstring>{L"renamed_0", L"renamed_1.txt"},
                  L"Parent/child preview should propose expected new names.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename parent/child execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    const std::filesystem::path renamedDir   = root / L"renamed_0";
    const std::filesystem::path renamedChild = renamedDir / L"renamed_1.txt";
    state.Require(! SelfTest::PathExists(sourceDir), L"Parent/child execution should remove the original parent directory path.");
    state.Require(! SelfTest::PathExists(sourceChild), L"Parent/child execution should remove the original child path.");
    state.Require(SelfTest::PathExists(renamedDir), L"Parent/child execution should create the renamed parent directory.");
    state.Require(SelfTest::PathExists(renamedChild), L"Parent/child execution should create the renamed child under the renamed parent.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename parent/child preview should refresh after executing.");
    state.Require(after.fullPaths == std::vector<std::wstring>{renamedDir.native(), renamedChild.native()},
                  L"Parent/child execution should refresh child preview paths under the renamed parent.");
    state.Require(after.originalNames == std::vector<std::wstring>{L"renamed_0", L"renamed_1.txt"},
                  L"Parent/child execution should refresh original names to the renamed leaves.");
    state.Require(after.changedRowCount == 0u, L"Parent/child execution should leave the refreshed preview with no pending changes.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowBlocksDestinationCreatedAfterPreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_dest_revalidate_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path firstSource  = root / L"very-long-source-name.txt";
    const std::filesystem::path secondSource = root / L"b.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename destination-revalidation root.");
    state.Require(SelfTest::WriteTextFile(firstSource, "first"), L"Failed to create first Batch Rename destination-revalidation input.");
    state.Require(SelfTest::WriteTextFile(secondSource, "second"), L"Failed to create second Batch Rename destination-revalidation input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename destination revalidation.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-dest-revalidate-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {firstSource, secondSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-dest-revalidate-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for destination revalidation testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename destination revalidation test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename destination revalidation test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename destination revalidation preview should be available.");
    state.Require(before.renameButtonEnabled, L"Destination revalidation preview should initially be executable.");
    state.Require(before.newNames == std::vector<std::wstring>{L"renamed_001.txt", L"renamed_002.txt"},
                  L"Destination revalidation preview should propose expected new names.");

    state.Require(SelfTest::WriteTextFile(root / L"renamed_002.txt", "conflict"),
                  L"Failed to create conflicting destination after Batch Rename preview.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(FAILED(executeHr), L"Batch Rename should fail pre-dispatch revalidation when a destination appears after preview.");
    state.Require(SelfTest::PathExists(firstSource), L"Destination revalidation should not partially rename the first source.");
    state.Require(SelfTest::PathExists(secondSource), L"Destination revalidation should leave the conflicting source unchanged.");
    state.Require(! SelfTest::PathExists(root / L"renamed_001.txt"), L"Destination revalidation should not create earlier-row destinations.");
    state.Require(SelfTest::PathExists(root / L"renamed_002.txt"), L"Destination revalidation should preserve the conflicting destination.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowBlocksSourceMissingAfterPreview(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_source_revalidate_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path firstSource  = root / L"very-long-source-name.txt";
    const std::filesystem::path secondSource = root / L"b.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename source-revalidation root.");
    state.Require(SelfTest::WriteTextFile(firstSource, "first"), L"Failed to create first Batch Rename source-revalidation input.");
    state.Require(SelfTest::WriteTextFile(secondSource, "second"), L"Failed to create second Batch Rename source-revalidation input.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    if (! fileSystem)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(L"builtin/file-system", g_settings);
        state.Require(SUCCEEDED(enableHr), std::format(L"Failed to enable local file-system plugin: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fileSystem = SelfTest::GetFileSystem(L"builtin/file-system");
    }
    state.Require(fileSystem != nullptr, L"Local file-system plugin should be available for Batch Rename source revalidation.");
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-source-revalidate-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {firstSource, secondSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-source-revalidate-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for source revalidation testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename source revalidation test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"renamed_{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename source revalidation test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename source revalidation preview should be available.");
    state.Require(before.renameButtonEnabled, L"Source revalidation preview should initially be executable.");

    std::filesystem::remove(secondSource, ec);
    state.Require(! ec, L"Failed to delete a source after Batch Rename preview.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(FAILED(executeHr), L"Batch Rename should fail pre-dispatch revalidation when a source disappears after preview.");
    state.Require(SelfTest::PathExists(firstSource), L"Source revalidation should not partially rename the first source.");
    state.Require(! SelfTest::PathExists(secondSource), L"Source revalidation should preserve the missing-source state.");
    state.Require(! SelfTest::PathExists(root / L"renamed_001.txt"), L"Source revalidation should not create earlier-row destinations.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameHelperMenusExposeCanonicalInsertions(CaseState& state) noexcept
{
    using namespace BatchRenameMenus;

    const std::vector<RedSalamander::DxUi::MenuFlyoutItem> templateItems = BuildTemplateHelperMenuItems();
    const std::vector<RedSalamander::DxUi::MenuFlyoutItem> regexItems    = BuildRegexSearchHelperMenuItems();
    const std::vector<RedSalamander::DxUi::MenuFlyoutItem> replaceItems  = BuildReplacementHelperMenuItems();

    const std::function<bool(const std::vector<RedSalamander::DxUi::MenuFlyoutItem>&, std::wstring_view)> hasInsertion =
        [&](const std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items, std::wstring_view expected) noexcept
    {
        for (const RedSalamander::DxUi::MenuFlyoutItem& item : items)
        {
            if (item.commandId != 0)
            {
                const std::optional<std::wstring_view> insertion = TryGetHelperInsertionText(item.commandId);
                if (insertion.has_value() && insertion.value() == expected)
                {
                    return true;
                }
            }
            if (! item.children.empty() && hasInsertion(item.children, expected))
            {
                return true;
            }
        }
        return false;
    };

    const std::function<bool(const std::vector<RedSalamander::DxUi::MenuFlyoutItem>&, int)> hasCommandId =
        [&](const std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items, const int commandId) noexcept
    {
        for (const RedSalamander::DxUi::MenuFlyoutItem& item : items)
        {
            if (item.commandId == commandId)
            {
                return true;
            }
            if (! item.children.empty() && hasCommandId(item.children, commandId))
            {
                return true;
            }
        }
        return false;
    };

    state.Require(hasInsertion(templateItems, L"{name}"), L"Template helper menu should insert canonical {name} macro syntax.");
    state.Require(hasInsertion(templateItems, L"{counter:000}"), L"Template helper menu should expose padded counter macro syntax.");
    state.Require(hasInsertion(templateItems, L"{date:yyyy-MM-dd}"), L"Template helper menu should expose date macro syntax.");
    state.Require(hasInsertion(templateItems, L"{{"), L"Template helper menu should expose literal opening brace escape.");
    state.Require(hasInsertion(regexItems, L"."), L"Regex helper menu should expose any-character syntax.");
    state.Require(hasInsertion(regexItems, L"\\d+"), L"Regex helper menu should expose decimal-number syntax.");
    state.Require(hasInsertion(regexItems, L"^(.+?)(\\.[^.]+)?$"), L"Regex helper menu should expose file-name split syntax.");
    state.Require(! hasInsertion(regexItems, L"(?:)"), L"Regex helper menu should omit unsupported non-capturing group syntax.");
    state.Require(hasCommandId(regexItems, BatchRenameMenus::RegexEscapedLiteralHelperCommandId()),
                  L"Regex helper menu should expose a selected-text escaped-literal command.");
    state.Require(hasInsertion(replaceItems, L"$$"), L"Replacement helper menu should expose literal-dollar syntax.");
    state.Require(hasInsertion(replaceItems, L"$&"), L"Replacement helper menu should expose whole-match syntax.");
    state.Require(hasInsertion(replaceItems, L"$1"), L"Replacement helper menu should expose first matched subexpression syntax.");
    state.Require(hasCommandId(replaceItems, BatchRenameMenus::ReplacementCustomSubexpressionHelperCommandId()),
                  L"Replacement helper menu should expose a custom matched-subexpression command.");

    const std::optional<HelperCommandInsertion> escapedLiteral =
        TryBuildDynamicHelperInsertion(RegexEscapedLiteralHelperCommandId(), L"file (1).txt+[ok]\\end");
    state.Require(escapedLiteral.has_value(), L"Regex escaped-literal helper should build insertion text from selected text.");
    if (escapedLiteral.has_value())
    {
        state.Require(escapedLiteral->insertionText == L"file \\(1\\)\\.txt\\+\\[ok\\]\\\\end",
                      L"Regex escaped-literal helper should escape ECMAScript metacharacters.");
        state.Require(escapedLiteral->selectionStart == escapedLiteral->selectionEnd &&
                          escapedLiteral->selectionStart == escapedLiteral->insertionText.size(),
                      L"Regex escaped-literal helper should leave the caret after the escaped literal.");
    }

    const std::optional<HelperCommandInsertion> emptyEscapedLiteral =
        TryBuildDynamicHelperInsertion(RegexEscapedLiteralHelperCommandId(), L"");
    state.Require(emptyEscapedLiteral.has_value(), L"Regex escaped-literal helper should have an empty-selection insertion.");
    if (emptyEscapedLiteral.has_value())
    {
        state.Require(emptyEscapedLiteral->insertionText == L"\\\\",
                      L"Regex escaped-literal helper should insert a literal-backslash skeleton when no text is selected.");
        state.Require(emptyEscapedLiteral->selectionStart == 1u && emptyEscapedLiteral->selectionEnd == 2u,
                      L"Regex escaped-literal helper should select the editable escape target for quick overwrite.");
    }

    const std::optional<HelperCommandInsertion> customDefault =
        TryBuildDynamicHelperInsertion(ReplacementCustomSubexpressionHelperCommandId(), L"");
    state.Require(customDefault.has_value(), L"Custom matched-subexpression helper should build a default token.");
    if (customDefault.has_value())
    {
        state.Require(customDefault->insertionText == L"$1", L"Custom matched-subexpression helper should default to $1.");
        state.Require(customDefault->selectionStart == 1u && customDefault->selectionEnd == 2u,
                      L"Custom matched-subexpression helper should select the default capture number for quick overwrite.");
    }

    const std::optional<HelperCommandInsertion> customDigits =
        TryBuildDynamicHelperInsertion(ReplacementCustomSubexpressionHelperCommandId(), L"12");
    state.Require(customDigits.has_value(), L"Custom matched-subexpression helper should accept selected digits.");
    if (customDigits.has_value())
    {
        state.Require(customDigits->insertionText == L"$12", L"Custom matched-subexpression helper should prefix selected digits with dollar.");
        state.Require(customDigits->selectionStart == customDigits->selectionEnd && customDigits->selectionStart == 3u,
                      L"Custom matched-subexpression helper should leave the caret after selected custom token.");
    }

    const HelperInsertionResult caretInsert = ApplyHelperInsertion(L"alpha.txt", 5u, 5u, L"{counter}");
    state.Require(caretInsert.text == L"alpha{counter}.txt", L"Helper insertion should insert at the caret.");
    state.Require(caretInsert.selectionStart == caretInsert.selectionEnd && caretInsert.selectionStart == 14u,
                  L"Helper insertion should leave the caret after inserted text.");

    const HelperInsertionResult selectionReplace = ApplyHelperInsertion(L"alpha.txt", 0u, 5u, L"{stem}");
    state.Require(selectionReplace.text == L"{stem}.txt", L"Helper insertion should replace selected text.");
    state.Require(selectionReplace.selectionStart == selectionReplace.selectionEnd && selectionReplace.selectionStart == 6u,
                  L"Helper replacement should leave the caret after inserted text.");

    const HelperInsertionResult reverseSelection = ApplyHelperInsertion(L"abc", 3u, 1u, L"X");
    state.Require(reverseSelection.text == L"aX", L"Helper insertion should normalize and clamp reverse selections.");
    state.Require(reverseSelection.selectionStart == 2u && reverseSelection.selectionEnd == 2u,
                  L"Reverse-selection helper insertion should report the normalized caret position.");

    return state.failure.empty();
}

[[nodiscard]] std::optional<int> FindBatchRenameHelperCommandId(const std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items,
                                                                std::wstring_view expectedInsertion) noexcept
{
    for (const RedSalamander::DxUi::MenuFlyoutItem& item : items)
    {
        if (item.commandId != 0)
        {
            const std::optional<std::wstring_view> insertion = BatchRenameMenus::TryGetHelperInsertionText(item.commandId);
            if (insertion.has_value() && insertion.value() == expectedInsertion)
            {
                return item.commandId;
            }
        }
        if (! item.children.empty())
        {
            if (const std::optional<int> childCommand = FindBatchRenameHelperCommandId(item.children, expectedInsertion); childCommand.has_value())
            {
                return childCommand;
            }
        }
    }
    return std::nullopt;
}

[[nodiscard]] bool TestBatchRenameWindowHelperButtonsInsertIntoRuleFields(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameHelperButtonsSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-helper-buttons-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"Clip.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-helper-buttons-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for helper-button insertion testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename helper-buttons test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    const std::optional<int> counterCommand =
        FindBatchRenameHelperCommandId(BatchRenameMenus::BuildTemplateHelperMenuItems(), L"{counter:000}");
    state.Require(counterCommand.has_value(), L"Template helper menu should expose a padded-counter command id.");
    const std::optional<int> digitCommand = FindBatchRenameHelperCommandId(BatchRenameMenus::BuildRegexSearchHelperMenuItems(), L"\\d");
    state.Require(digitCommand.has_value(), L"Regex helper menu should expose a digit command id.");
    const std::optional<int> firstCaptureCommand = FindBatchRenameHelperCommandId(BatchRenameMenus::BuildReplacementHelperMenuItems(), L"$1");
    state.Require(firstCaptureCommand.has_value(), L"Replacement helper menu should expose a first-capture command id.");
    const int escapedLiteralCommand = BatchRenameMenus::RegexEscapedLiteralHelperCommandId();
    const int customCaptureCommand  = BatchRenameMenus::ReplacementCustomSubexpressionHelperCommandId();
    if (! counterCommand.has_value() || ! digitCommand.has_value() || ! firstCaptureCommand.has_value())
    {
        return false;
    }

    BatchRenameDebugSnapshot initialSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(initialSnapshot), L"Batch Rename helper-buttons initial snapshot should be available.");
    state.Require(initialSnapshot.ruleControlsVisible, L"Batch Rename helper-buttons test should start in rules mode.");
    state.Require(initialSnapshot.ruleHelperButtonsVisible, L"Batch Rename rule fields should expose helper menu buttons.");

    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::NameTemplate, 0u, 6u),
                  L"Debug helper should select the current template token.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::NameTemplate, counterCommand.value()),
                  L"Template helper command should insert into the visible name-template field.");

    BatchRenameDebugSnapshot templateSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(templateSnapshot), L"Batch Rename helper template snapshot should be available.");
    state.Require(templateSnapshot.nameTemplateText == L"{counter:000}",
                  L"Template helper insertion should replace the selected template text.");
    state.Require(templateSnapshot.newNames == std::vector<std::wstring>{L"001"},
                  L"Template helper insertion should recompute preview through the window controls.");
    state.Require(templateSnapshot.changedRowCount == 1u, L"Template helper insertion should update changed-row stats.");
    state.Require(templateSnapshot.renameButtonEnabled, L"Valid helper-driven template edits should enable Rename.");

    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::SearchFor, 0u, 0u),
                  L"Debug helper should place the caret in the search field.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::SearchFor, digitCommand.value()),
                  L"Regex helper command should insert into the visible search field.");
    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::ReplaceWith, 0u, 0u),
                  L"Debug helper should place the caret in the replace field.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::ReplaceWith, firstCaptureCommand.value()),
                  L"Replacement helper command should insert into the visible replacement field.");

    BatchRenameDebugSnapshot regexSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(regexSnapshot), L"Batch Rename helper regex snapshot should be available.");
    state.Require(regexSnapshot.searchForText == L"\\d", L"Regex helper insertion should update the search field text.");
    state.Require(regexSnapshot.replaceWithText == L"$1", L"Replacement helper insertion should update the replacement field text.");

    BatchRename::Rules selectedTextRules{};
    selectedTextRules.nameTemplate   = L"{name}";
    selectedTextRules.searchFor      = L"file (1).txt";
    selectedTextRules.replaceWith    = L"12";
    selectedTextRules.regexEnabled   = true;
    selectedTextRules.caseSensitive  = true;
    selectedTextRules.replaceOnce    = false;
    selectedTextRules.excludeExtension = false;
    state.Require(DebugSetBatchRenameWindowRuleControls(selectedTextRules),
                  L"Batch Rename helper-buttons test should set selected-text rule fields.");
    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::SearchFor, 0u, selectedTextRules.searchFor.size()),
                  L"Debug helper should select regex text before escaped-literal insertion.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::SearchFor, escapedLiteralCommand),
                  L"Regex escaped-literal helper command should replace the selected search field text.");
    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::ReplaceWith, 0u, selectedTextRules.replaceWith.size()),
                  L"Debug helper should select replacement digits before custom subexpression insertion.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::ReplaceWith, customCaptureCommand),
                  L"Custom matched-subexpression helper command should replace the selected replacement field text.");

    BatchRenameDebugSnapshot selectedTextSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(selectedTextSnapshot), L"Batch Rename dynamic helper snapshot should be available.");
    state.Require(selectedTextSnapshot.searchForText == L"file \\(1\\)\\.txt",
                  L"Regex escaped-literal helper should update the search field with escaped selected text.");
    state.Require(selectedTextSnapshot.replaceWithText == L"$12",
                  L"Custom matched-subexpression helper should update the replacement field from selected digits.");

    state.Require(DebugSetBatchRenameWindowRuleFieldText(BatchRenameDebugRuleField::SearchFor, L""),
                  L"Batch Rename helper-buttons test should clear the search field before empty escaped-literal insertion.");
    state.Require(DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField::SearchFor, 0u, 0u),
                  L"Debug helper should place the caret in the empty search field.");
    state.Require(DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField::SearchFor, escapedLiteralCommand),
                  L"Regex escaped-literal helper command should insert a skeleton when no text is selected.");

    BatchRenameDebugSnapshot emptySelectionSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(emptySelectionSnapshot), L"Batch Rename empty-selection helper snapshot should be available.");
    state.Require(emptySelectionSnapshot.searchForText == L"\\\\",
                  L"Regex escaped-literal helper should not silently no-op when invoked with an empty selection.");

    return state.failure.empty();
}

[[nodiscard]] bool SetupBatchRenamePaneFixture(CaseState& state,
                                               const std::filesystem::path& root,
                                               std::initializer_list<std::wstring_view> files,
                                               std::initializer_list<std::wstring_view> directories = {}) noexcept
{
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), std::format(L"Failed to create Batch Rename fixture root '{}'.", root.wstring()));

    for (std::wstring_view name : directories)
    {
        state.Require(SelfTest::EnsureDirectory(root / std::wstring(name)),
                      std::format(L"Failed to create Batch Rename fixture folder '{}'.", std::wstring(name)));
    }

    for (std::wstring_view name : files)
    {
        state.Require(SelfTest::WriteTextFile(root / std::wstring(name), "batch rename fixture"),
                      std::format(L"Failed to create Batch Rename fixture file '{}'.", std::wstring(name)));
    }

    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Batch Rename pane fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Batch Rename fixture pane path did not load.");

    std::vector<std::wstring> expectedNames;
    expectedNames.reserve(files.size() + directories.size());
    for (std::wstring_view name : files)
    {
        expectedNames.emplace_back(name);
    }
    for (std::wstring_view name : directories)
    {
        expectedNames.emplace_back(name);
    }
    const auto itemsDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{3000});
    bool itemsReady          = false;
    while (std::chrono::steady_clock::now() < itemsDeadline)
    {
        PumpPendingMessages();
        itemsReady = std::ranges::all_of(expectedNames, [](const std::wstring& name) noexcept
        { return g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, name); });
        if (itemsReady)
        {
            break;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds{20});
    }
    state.Require(itemsReady, L"Batch Rename fixture pane items did not load.");
    FocusFolderViewPane(FolderWindow::Pane::Left);
    return state.failure.empty();
}

[[nodiscard]] bool WaitForPaneDisplayNames(FolderWindow::Pane pane,
                                           std::initializer_list<std::wstring_view> expectedPresent,
                                           std::initializer_list<std::wstring_view> expectedAbsent,
                                           std::chrono::milliseconds timeout) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        const bool presentReady = std::ranges::all_of(expectedPresent, [pane](std::wstring_view name) noexcept
        { return g_folderWindow.DebugHasItemDisplayName(pane, name); });
        const bool absentReady = std::ranges::none_of(expectedAbsent, [pane](std::wstring_view name) noexcept
        { return g_folderWindow.DebugHasItemDisplayName(pane, name); });
        if (presentReady && absentReady)
        {
            return true;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds{20});
    }

    return false;
}

void CloseBatchRenameWindowIfOpen() noexcept
{
    if (const HWND batchWindow = GetBatchRenameWindowHandle(); batchWindow && IsWindow(batchWindow) != FALSE)
    {
        PostMessageW(batchWindow, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(batchWindow, SelfTest::Scale(std::chrono::milliseconds{3000})));
    }
}

void CloseRenamePromptIfOpen() noexcept
{
    if (const HWND prompt = GetFolderViewRenamePromptHandle(); prompt && IsWindow(prompt) != FALSE)
    {
        PostMessageW(prompt, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
    }
}

class BatchRenameSettingsTestScope final
{
public:
    BatchRenameSettingsTestScope()
        : _previousBatchRename(g_settings.batchRename)
    {
        const auto placementIt = g_settings.windows.find(L"BatchRenameWindow");
        if (placementIt != g_settings.windows.end())
        {
            _previousPlacement = placementIt->second;
        }
    }

    BatchRenameSettingsTestScope(const BatchRenameSettingsTestScope&)            = delete;
    BatchRenameSettingsTestScope& operator=(const BatchRenameSettingsTestScope&) = delete;

    ~BatchRenameSettingsTestScope()
    {
        CloseBatchRenameWindowIfOpen();
        g_settings.batchRename = _previousBatchRename;
        if (_previousPlacement.has_value())
        {
            g_settings.windows[L"BatchRenameWindow"] = _previousPlacement.value();
        }
        else
        {
            g_settings.windows.erase(L"BatchRenameWindow");
        }
    }

    void ResetForDefaultWindow() noexcept
    {
        CloseBatchRenameWindowIfOpen();
        g_settings.batchRename.reset();
        g_settings.windows.erase(L"BatchRenameWindow");
    }

private:
    std::optional<Common::Settings::BatchRenameSettings> _previousBatchRename;
    std::optional<Common::Settings::WindowPlacement> _previousPlacement;
};

[[nodiscard]] bool TestPaneRenameMultiSelectionOpensBatchRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_multi_" + NewGuidText());
    if (! SetupBatchRenamePaneFixture(state, root, {L"one.batch", L"two.batch", L"three.keep"}))
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupPrompt      = wil::scope_exit([]() noexcept { CloseRenamePromptIfOpen(); });
    CloseBatchRenameWindowIfOpen();
    CloseRenamePromptIfOpen();

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"one.batch"),
                  L"Failed to focus first file for multi-selection Batch Rename test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"one.batch" || name == L"two.batch"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u,
                  L"Batch Rename multi-selection fixture should have exactly two selected files.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_RENAME, 0), 0);
    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Multi-selection Rename should open Batch Rename.");
    state.Require(GetFolderViewRenamePromptHandle() == nullptr, L"Multi-selection Rename should bypass the standard rename prompt.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename multi-selection snapshot should be available.");
    state.Require(snapshot.rootText == root.native(), L"Multi-selection Batch Rename should use the active pane root.");
    state.Require(snapshot.previewRowCount == 2u, L"Multi-selection Batch Rename should preview exactly the selected files.");
    state.Require(snapshot.originalNames == std::vector<std::wstring>{L"one.batch", L"two.batch"},
                  L"Multi-selection Batch Rename should preserve selected file order.");
    state.Require(snapshot.fullPaths == std::vector<std::wstring>{(root / L"one.batch").native(), (root / L"two.batch").native()},
                  L"Multi-selection Batch Rename should seed selected full paths.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenameSingleFileUsesStandardPrompt(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_single_file_" + NewGuidText());
    if (! SetupBatchRenamePaneFixture(state, root, {L"solo.batch"}))
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupPrompt      = wil::scope_exit([]() noexcept { CloseRenamePromptIfOpen(); });
    CloseBatchRenameWindowIfOpen();
    CloseRenamePromptIfOpen();

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"solo.batch"),
                  L"Failed to focus file for single-file rename prompt test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"solo.batch"; }, true);
    if (! state.failure.empty())
    {
        return false;
    }

    struct PromptProbe final
    {
        bool sawPrompt = false;
        bool canceled  = false;
        FolderViewRenamePromptDebugSnapshot snapshot{};
    } probe;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt =
            WaitForWindow([]() noexcept { return GetFolderViewRenamePromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        static_cast<void>(DebugGetFolderViewRenamePromptSnapshot(probe.snapshot));
        probe.canceled = DebugCancelFolderViewRenamePrompt();
        static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_RENAME, 0), 0);
    worker.join();

    state.Require(probe.sawPrompt, L"Single-file Rename should keep opening the standard rename prompt.");
    state.Require(probe.canceled, L"Single-file rename prompt should be cancelable through the debug helper.");
    state.Require(probe.snapshot.batchButtonVisible, L"Single-file rename prompt should expose the Batch action.");
    state.Require(GetBatchRenameWindowHandle() == nullptr, L"Single-file Rename should not open Batch Rename until Batch is invoked.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenameFilePromptBatchButtonOpensBatchRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_file_prompt_" + NewGuidText());
    if (! SetupBatchRenamePaneFixture(state, root, {L"prompt-file.batch"}))
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupPrompt      = wil::scope_exit([]() noexcept { CloseRenamePromptIfOpen(); });
    CloseBatchRenameWindowIfOpen();
    CloseRenamePromptIfOpen();

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"prompt-file.batch"),
                  L"Failed to focus file for Batch button rename prompt test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"prompt-file.batch"; }, true);
    if (! state.failure.empty())
    {
        return false;
    }

    struct FileBatchProbe final
    {
        bool sawPrompt     = false;
        bool invokedBatch  = false;
        bool promptClosed  = false;
        HWND batchWindow   = nullptr;
        bool batchSnapshot = false;
        FolderViewRenamePromptDebugSnapshot promptSnapshot{};
        BatchRenameDebugSnapshot batchRenameSnapshot{};
    } probe;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt =
            WaitForWindow([]() noexcept { return GetFolderViewRenamePromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        static_cast<void>(DebugGetFolderViewRenamePromptSnapshot(probe.promptSnapshot));
        probe.invokedBatch = DebugInvokeFolderViewRenamePromptBatch();
        probe.promptClosed = WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000}));
        probe.batchWindow =
            WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        if (probe.batchWindow && IsWindow(probe.batchWindow) != FALSE)
        {
            // The window handle is published before SetContext seeds the explicit selection,
            // so poll until the seeded row is observable rather than racing the main thread.
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{5000});
            do
            {
                probe.batchSnapshot = DebugGetBatchRenameWindowSnapshot(probe.batchRenameSnapshot);
                if (probe.batchSnapshot && probe.batchRenameSnapshot.previewRowCount > 0u)
                {
                    break;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds{20});
            } while (std::chrono::steady_clock::now() < deadline);
        }
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_RENAME, 0), 0);
    worker.join();

    state.Require(probe.sawPrompt, L"File Rename should open the standard rename prompt first.");
    state.Require(probe.promptSnapshot.batchButtonVisible, L"File rename prompt should expose the Batch action.");
    state.Require(probe.invokedBatch, L"File rename prompt Batch debug action should be invokable.");
    state.Require(probe.promptClosed, L"File rename prompt should close after invoking Batch.");
    state.Require(probe.batchWindow != nullptr && IsWindow(probe.batchWindow) != FALSE,
                  L"File rename prompt Batch action should open Batch Rename.");
    state.Require(probe.batchSnapshot, L"File Batch Rename snapshot should be available.");
    state.Require(probe.batchRenameSnapshot.rootText == root.native(), L"File Batch Rename should use the active pane root.");
    std::wstring previewNames;
    for (const std::wstring& name : probe.batchRenameSnapshot.originalNames)
    {
        if (! previewNames.empty())
        {
            previewNames += L", ";
        }
        previewNames += name;
    }
    state.Require(probe.batchRenameSnapshot.previewRowCount == 1u,
                  std::format(L"File Batch Rename should preview exactly the prompted file (rows={}, names=[{}]).",
                              probe.batchRenameSnapshot.previewRowCount,
                              previewNames));
    state.Require(probe.batchRenameSnapshot.originalNames == std::vector<std::wstring>{L"prompt-file.batch"},
                  L"File Batch Rename should seed the prompted file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenameFolderPromptBatchButtonOpensBatchRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"batch_rename_folder_prompt_" + NewGuidText());
    const std::filesystem::path folderRoot = root / L"scope-folder";
    if (! SetupBatchRenamePaneFixture(state, root, {}, {L"scope-folder"}))
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupPrompt      = wil::scope_exit([]() noexcept { CloseRenamePromptIfOpen(); });
    CloseBatchRenameWindowIfOpen();
    CloseRenamePromptIfOpen();

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"scope-folder"),
                  L"Failed to focus folder for Batch button rename prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct FolderBatchProbe final
    {
        bool sawPrompt      = false;
        bool invokedBatch   = false;
        bool promptClosed   = false;
        HWND batchWindow    = nullptr;
        bool batchSnapshot  = false;
        FolderViewRenamePromptDebugSnapshot promptSnapshot{};
        BatchRenameDebugSnapshot batchRenameSnapshot{};
    } probe;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt =
            WaitForWindow([]() noexcept { return GetFolderViewRenamePromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        static_cast<void>(DebugGetFolderViewRenamePromptSnapshot(probe.promptSnapshot));
        probe.invokedBatch = DebugInvokeFolderViewRenamePromptBatch();
        probe.promptClosed = WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000}));
        probe.batchWindow =
            WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        if (probe.batchWindow && IsWindow(probe.batchWindow) != FALSE)
        {
            probe.batchSnapshot = DebugGetBatchRenameWindowSnapshot(probe.batchRenameSnapshot);
        }
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_RENAME, 0), 0);
    worker.join();

    state.Require(probe.sawPrompt, L"Folder Rename should open the standard rename prompt first.");
    state.Require(probe.promptSnapshot.batchButtonVisible, L"Folder rename prompt should expose the Batch action.");
    state.Require(probe.invokedBatch, L"Folder rename prompt Batch debug action should be invokable.");
    state.Require(probe.promptClosed, L"Folder rename prompt should close after invoking Batch.");
    state.Require(probe.batchWindow != nullptr && IsWindow(probe.batchWindow) != FALSE,
                  L"Folder rename prompt Batch action should open Batch Rename.");
    state.Require(probe.batchSnapshot, L"Folder Batch Rename snapshot should be available.");
    state.Require(probe.batchRenameSnapshot.rootText == folderRoot.native(), L"Folder Batch Rename should be rooted at the chosen folder.");
    state.Require(probe.batchRenameSnapshot.previewRowCount == 0u,
                  L"Folder Batch Rename should wait for scope enumeration instead of previewing the renamed folder itself.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineMacroRegexCaseAndValidation(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target alpha{};
    alpha.sourcePath = L"C:\\root\\Alpha File.txt";
    alpha.isDirectory = false;
    alpha.sizeBytes = 1234;

    Target duplicate{};
    duplicate.sourcePath = L"C:\\root\\duplicate.txt";
    duplicate.isDirectory = false;
    duplicate.sizeBytes = 10;

    Target duplicateCase{};
    duplicateCase.sourcePath = L"C:\\root\\Duplicate.TXT";
    duplicateCase.isDirectory = false;
    duplicateCase.sizeBytes = 20;

    Rules rules{};
    rules.nameTemplate = L"{counter:000}_{stem}{ext}";
    rules.searchFor = L"file";
    rules.replaceWith = L"clip";
    rules.caseSensitive = false;
    rules.excludeExtension = true;
    rules.fileNameCaseStyle = CaseTransform::Upper;
    rules.extensionCaseStyle = CaseTransform::Lower;

    const std::vector<Target> targets{alpha, duplicate, duplicateCase};
    const Plan plan = BuildPlan(targets, rules);

    state.Require(plan.rows.size() == 3u, L"Batch rename plan should contain one preview row per target.");
    if (plan.rows.size() != 3u)
    {
        return false;
    }

    state.Require(plan.rows[0].originalName == L"Alpha File.txt", L"Original name should be the source leaf name.");
    state.Require(plan.rows[0].newName == L"001_ALPHA CLIP.txt", L"Macro, search/replace, and case transforms should compose in order.");
    state.Require(plan.rows[0].issues.empty(), L"Valid transformed row should not report issues.");

    state.Require(plan.rows[1].newName == L"002_DUPLICATE.txt", L"Counter should use preview order for each row.");
    state.Require(plan.rows[2].newName == L"003_DUPLICATE.txt", L"Extension case transform should normalize the extension only.");
    state.Require(plan.stats.totalRows == 3u, L"Stats should count every preview row.");
    state.Require(plan.stats.changedRows == 3u, L"Stats should count changed rows.");
    state.Require(plan.stats.errorRows == 0u, L"Distinct generated names should not produce duplicate errors.");

    rules.nameTemplate = L"same{ext}";
    const Plan duplicatePlan = BuildPlan({duplicate, duplicateCase}, rules);
    state.Require(duplicatePlan.stats.errorRows == 2u,
                  L"Rows targeting names that differ only by case in the same parent should be blocking duplicate errors.");

    rules.nameTemplate = L"{unknown}";
    const Plan unknownMacroPlan = BuildPlan({alpha}, rules);
    state.Require(unknownMacroPlan.stats.errorRows == 1u, L"Unknown macros should be blocking validation errors.");

    rules.nameTemplate = L"literal.txt";
    rules.searchFor = L"(";
    rules.regexEnabled = true;
    const Plan invalidRegexPlan = BuildPlan({alpha}, rules);
    state.Require(invalidRegexPlan.stats.errorRows == 1u, L"Invalid regular expressions should be reported as validation errors.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineUsesProviderPathIdentity(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target first{};
    first.sourcePath = L"C:\\root\\A.txt";

    Target second{};
    second.sourcePath = L"C:\\root\\a.txt";

    Rules rules{};
    rules.nameTemplate = L"{name}";

    const FileSystemPathIdentity caseInsensitive = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    const Plan caseInsensitivePlan              = BuildPlan({first, second}, rules, caseInsensitive);
    state.Require(caseInsensitivePlan.stats.errorRows == 2u,
                  L"Case-insensitive provider identity should reject sibling targets that differ only by case.");
    if (caseInsensitivePlan.rows.size() == 2u)
    {
        state.Require(HasBatchRenameIssue(caseInsensitivePlan.rows[0], IssueSeverity::Error, L"name_duplicate") &&
                          HasBatchRenameIssue(caseInsensitivePlan.rows[1], IssueSeverity::Error, L"name_duplicate"),
                      L"Case-insensitive provider duplicate rows should carry name_duplicate.");
    }

    FileSystemPathIdentity caseSensitive = caseInsensitive;
    caseSensitive.componentComparison    = FileSystemPathComponentComparison::OrdinalCaseSensitive;
    const Plan caseSensitivePlan         = BuildPlan({first, second}, rules, caseSensitive);
    state.Require(caseSensitivePlan.stats.errorRows == 0u,
                  L"Case-sensitive provider identity should allow sibling targets that differ only by case.");
    state.Require(caseSensitivePlan.stats.warningRows == 2u,
                  L"Case-sensitive provider identity should treat exact no-op rows as warnings only.");

    Target mixed{};
    mixed.sourcePath = L"C:\\root\\Readme.txt";
    rules.fileNameCaseStyle  = CaseTransform::Upper;
    rules.extensionCaseStyle = CaseTransform::Upper;

    const Plan caseOnlyInsensitivePlan = BuildPlan({mixed}, rules, caseInsensitive);
    state.Require(caseOnlyInsensitivePlan.rows.size() == 1u &&
                      HasBatchRenameIssue(caseOnlyInsensitivePlan.rows[0], IssueSeverity::Warning, L"name_case_only"),
                  L"Case-insensitive provider identity should flag text-only case changes as case-only renames.");

    const Plan caseOnlySensitivePlan = BuildPlan({mixed}, rules, caseSensitive);
    state.Require(caseOnlySensitivePlan.rows.size() == 1u &&
                      ! HasBatchRenameIssue(caseOnlySensitivePlan.rows[0], IssueSeverity::Warning, L"name_case_only"),
                  L"Case-sensitive provider identity should not flag case-only text changes as same-item renames.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenamePathIdentityParserRejectsUnplannableProfiles(CaseState& state) noexcept
{
    constexpr std::string_view kIgnoreCaseJson = R"json(
{
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalIgnoreCase",
    "normalization": "none",
    "preferredSeparator": "\\",
    "acceptedSeparators": ["\\", "/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";

    const std::optional<FileSystemPathIdentity> ignoreCase = TryParseFileSystemPathIdentity(kIgnoreCaseJson, L"test-ignore");
    state.Require(ignoreCase.has_value(), L"Valid ordinalIgnoreCase path identity should parse.");
    if (ignoreCase.has_value())
    {
        state.Require(EquivalentComponent(ignoreCase.value(), L"Alpha.txt", L"alpha.TXT"),
                      L"ordinalIgnoreCase profile should compare components case-insensitively.");
        state.Require(EquivalentPath(ignoreCase.value(), L"Root\\Alpha/Leaf.txt", L"root/alpha\\leaf.TXT"),
                      L"ordinalIgnoreCase profile should normalize every accepted separator before path comparison.");

        const std::optional<std::wstring> lhsPathKey = TryMakePathKey(ignoreCase.value(), L"Root\\Alpha/Leaf.txt");
        const std::optional<std::wstring> rhsPathKey = TryMakePathKey(ignoreCase.value(), L"root/alpha\\leaf.TXT");
        state.Require(lhsPathKey.has_value() && rhsPathKey.has_value() && lhsPathKey.value() == rhsPathKey.value(),
                      L"ordinalIgnoreCase path keys should normalize accepted separators and ASCII case.");

        const std::optional<std::wstring> lhsComponentKey = TryMakeComponentKey(ignoreCase.value(), L"Alpha.txt");
        const std::optional<std::wstring> rhsComponentKey = TryMakeComponentKey(ignoreCase.value(), L"alpha.TXT");
        state.Require(lhsComponentKey.has_value() && rhsComponentKey.has_value() && lhsComponentKey.value() == rhsComponentKey.value(),
                      L"ordinalIgnoreCase component keys should normalize ASCII case.");

        state.Require(! TryMakePathKey(ignoreCase.value(), L"Root\\R\u00E9sum\u00E9.txt").has_value(),
                      L"ordinalIgnoreCase path keys should decline non-ASCII text so callers verify collisions with EquivalentPath.");
    }

    constexpr std::string_view kRenameJson = R"json(
{
  "version": 1,
  "operations": {
    "rename": true
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalIgnoreCase",
    "normalization": "none",
    "preferredSeparator": "\\",
    "acceptedSeparators": ["\\", "/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";
    state.Require(TryParseFileSystemRenamePathIdentity(kRenameJson, L"test-rename").has_value(),
                  L"Rename-capable provider identity should parse only when operations.rename is true.");

    constexpr std::string_view kRenameFalseJson = R"json(
{
  "version": 1,
  "operations": {
    "rename": false
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalIgnoreCase",
    "normalization": "none",
    "preferredSeparator": "\\",
    "acceptedSeparators": ["\\", "/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";
    state.Require(! TryParseFileSystemRenamePathIdentity(kRenameFalseJson, L"test-rename-false").has_value(),
                  L"Provider identity should be non-plannable for Batch Rename when operations.rename is false.");

    constexpr std::string_view kCaseSensitiveJson = R"json(
{
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";

    const std::optional<FileSystemPathIdentity> caseSensitive = TryParseFileSystemPathIdentity(kCaseSensitiveJson, L"test-sensitive");
    state.Require(caseSensitive.has_value(), L"Valid ordinalCaseSensitive path identity should parse.");
    if (caseSensitive.has_value())
    {
        state.Require(! EquivalentComponent(caseSensitive.value(), L"Alpha.txt", L"alpha.TXT"),
                      L"ordinalCaseSensitive profile should not fold component case.");
        state.Require(EquivalentPath(caseSensitive.value(), L"Root/Alpha/Leaf.txt", L"Root/Alpha/Leaf.txt"),
                      L"ordinalCaseSensitive profile should compare exact paths with the provider separator.");
        state.Require(! EquivalentPath(caseSensitive.value(), L"Root\\Alpha\\Leaf.txt", L"Root/Alpha/Leaf.txt"),
                      L"ordinalCaseSensitive profile should not accept separators the provider did not advertise.");

        const std::optional<std::wstring> exactPathKey = TryMakePathKey(caseSensitive.value(), L"Root/Alpha/Leaf.txt");
        const std::optional<std::wstring> mismatchedSeparatorKey = TryMakePathKey(caseSensitive.value(), L"Root\\Alpha\\Leaf.txt");
        state.Require(exactPathKey.has_value(), L"ordinalCaseSensitive path keys should build for accepted provider separators.");
        state.Require(! mismatchedSeparatorKey.has_value(),
                      L"ordinalCaseSensitive path keys should reject separators the provider did not advertise.");
    }

    constexpr std::string_view kMissingSeparatorJson = R"json(
{
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalIgnoreCase",
    "normalization": "none",
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";
    state.Require(! TryParseFileSystemPathIdentity(kMissingSeparatorJson, L"test-missing-separators").has_value(),
                  L"Path identity should reject profiles that omit preferredSeparator/acceptedSeparators.");

    constexpr std::string_view kUnknownJson = R"json(
{
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "unknown",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";
    state.Require(! TryParseFileSystemPathIdentity(kUnknownJson, L"test-unknown").has_value(),
                  L"Plugin-emitted componentComparison=unknown should be rejected as a contract violation.");

    constexpr std::string_view kUnstableJson = R"json(
{
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": false,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "notApplicable"
  }
}
)json";
    state.Require(! TryParseFileSystemPathIdentity(kUnstableJson, L"test-unstable").has_value(),
                  L"Unstable path text identity should be rejected for Batch Rename planning.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineMacroAliasDateTimeAndRegexReplacementTokens(CaseState& state) noexcept
{
    using namespace BatchRename;
    using namespace std::chrono_literals;

    Target alpha{};
    alpha.sourcePath    = L"C:\\root\\Alpha File.txt";
    alpha.isDirectory   = false;
    alpha.sizeBytes     = 1234;
    alpha.lastWriteTime = std::chrono::sys_days{std::chrono::year{2026} / std::chrono::June / 10} + 14h + 5min + 9s;

    Rules macroRules{};
    macroRules.nameTemplate = L"$(Name)__{{literal}}__{date:yyyy-MM-dd}_{time:HH-mm-ss}";

    const Plan macroPlan = BuildPlan({alpha}, macroRules);
    state.Require(macroPlan.rows.size() == 1u, L"Macro alias/date-time plan should contain one preview row.");
    if (macroPlan.rows.size() != 1u)
    {
        return false;
    }

    const BatchRenameLocalStampParts aliasParts = GetBatchRenameExpectedLocalParts(alpha.lastWriteTime.value());
    const std::wstring expectedAliasName        = std::format(L"Alpha File.txt__{{literal}}__{:04}-{:02}-{:02}_{:02}-{:02}-{:02}",
                                                       aliasParts.year,
                                                       aliasParts.month,
                                                       aliasParts.day,
                                                       aliasParts.hour,
                                                       aliasParts.minute,
                                                       aliasParts.second);
    state.Require(macroPlan.rows[0].newName == expectedAliasName,
                  L"Alias macros, escaped braces, and local-time date/time macros should expand together.");
    state.Require(macroPlan.stats.errorRows == 0u, L"Supported alias/date-time macros should not report validation errors.");

    Target numbered{};
    numbered.sourcePath = L"C:\\root\\Alpha 42.txt";

    Rules regexRules{};
    regexRules.nameTemplate   = L"{name}";
    regexRules.searchFor      = L"^([A-Za-z]+) (\\d+)(\\.txt)$";
    regexRules.replaceWith    = L"$2-$1$$$&";
    regexRules.regexEnabled   = true;
    regexRules.caseSensitive  = true;
    regexRules.replaceOnce    = true;

    const Plan regexPlan = BuildPlan({numbered}, regexRules);
    state.Require(regexPlan.rows.size() == 1u, L"Regex replacement token plan should contain one preview row.");
    if (regexPlan.rows.size() != 1u)
    {
        return false;
    }

    state.Require(regexPlan.rows[0].newName == L"42-Alpha$Alpha 42.txt",
                  L"Regex replacement should honor $1, $&, and $$ ECMAScript replacement tokens.");
    state.Require(regexPlan.stats.errorRows == 0u, L"Valid regex replacement tokens should not report validation errors.");

    Target nested{};
    nested.sourcePath     = L"C:\\root\\Series\\Season 01\\Episode 01.mkv";
    nested.relativeFolder = L"Series\\Season 01";
    nested.sizeBytes      = 42;
    nested.createdTime    = std::chrono::sys_days{std::chrono::year{2025} / std::chrono::December / 31} + 23h + 59min + 58s;

    Rules relativeRules{};
    relativeRules.flattenSeparator = L"__";
    relativeRules.nameTemplate     = L"{relativeFolderFlat}_{filename}_{extNoDot}_{size}_{created:yyyyMMdd}_{counter}_{index}";

    const Plan relativePlan = BuildPlan({nested}, relativeRules);
    state.Require(relativePlan.rows.size() == 1u, L"Relative-folder macro plan should contain one preview row.");
    if (relativePlan.rows.size() == 1u)
    {
        const BatchRenameLocalStampParts createdParts = GetBatchRenameExpectedLocalParts(nested.createdTime.value());
        const std::wstring expectedRelativeName       = std::format(L"Series__Season 01_Episode 01.mkv_mkv_42_{:04}{:02}{:02}_1_0",
                                                              createdParts.year,
                                                              createdParts.month,
                                                              createdParts.day);
        state.Require(relativePlan.rows[0].newName == expectedRelativeName,
                      L"Relative-folder flat, filename, extNoDot, size, created, counter, and index macros should expand from target metadata.");
        state.Require(relativePlan.stats.errorRows == 0u, L"Flattened relative-folder macros should be valid leaf names.");
    }

    Rules rawRelativeRules{};
    rawRelativeRules.nameTemplate = L"{relativeFolder}_{name}";
    const Plan rawRelativePlan = BuildPlan({nested}, rawRelativeRules);
    state.Require(rawRelativePlan.rows.size() == 1u, L"Raw relative-folder macro plan should contain one preview row.");
    if (rawRelativePlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(rawRelativePlan.rows[0], IssueSeverity::Error, L"name_separator"),
                      L"Raw relative-folder macro output should be known but blocked when it contains path separators.");
        state.Require(! HasBatchRenameIssue(rawRelativePlan.rows[0], IssueSeverity::Error, L"macro_unknown"),
                      L"Raw relative-folder macro should not be reported as an unknown macro.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineLeafSplitterMatchesFilesystem(CaseState& state) noexcept
{
    using namespace BatchRename;

    const std::vector<std::wstring> leaves = {
        L"alpha.txt",
        L"archive.tar.gz",
        L".gitignore",
        L".profile.txt",
        L"trailing.",
        L"noext",
        L"..foo",
        L"...",
        L"a..b",
        L"a.b.",
    };

    std::vector<Target> targets;
    targets.reserve(leaves.size());
    for (const std::wstring& leaf : leaves)
    {
        Target target{};
        target.sourcePath = std::filesystem::path(L"C:\\root") / leaf;
        targets.push_back(std::move(target));
    }

    Rules rules{};
    rules.nameTemplate = L"{stem}__{ext}__{extNoDot}";

    const Plan plan = BuildPlan(targets, rules);
    state.Require(plan.rows.size() == leaves.size(), L"Leaf splitter corpus should keep one row per leaf.");
    if (plan.rows.size() != leaves.size())
    {
        return false;
    }

    for (size_t index = 0u; index < leaves.size(); ++index)
    {
        const std::filesystem::path leafPath(leaves[index]);
        std::wstring extension      = leafPath.extension().wstring();
        std::wstring extensionNoDot = extension;
        if (! extensionNoDot.empty() && extensionNoDot.front() == L'.')
        {
            extensionNoDot.erase(extensionNoDot.begin());
        }
        const std::wstring expected = leafPath.stem().wstring() + L"__" + extension + L"__" + extensionNoDot;
        state.Require(plan.rows[index].newName == expected,
                      std::format(L"Leaf splitter should match std::filesystem for '{}': expected '{}', saw '{}'.",
                                  leaves[index],
                                  expected,
                                  plan.rows[index].newName));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineManualModeLineValidation(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target first{};
    first.sourcePath = L"C:\\root\\one.txt";
    Target second{};
    second.sourcePath = L"C:\\root\\two.txt";

    Rules rules{};
    rules.mode = Mode::Manual;
    rules.manualNames = {L"uno.txt", L"dos.txt"};

    const Plan plan = BuildPlan({first, second}, rules);
    state.Require(plan.rows.size() == 2u, L"Manual mode should keep one row per target.");
    if (plan.rows.size() != 2u)
    {
        return false;
    }

    state.Require(plan.rows[0].newName == L"uno.txt", L"Manual mode should use the first manual line for the first row.");
    state.Require(plan.rows[1].newName == L"dos.txt", L"Manual mode should use the second manual line for the second row.");
    state.Require(plan.stats.errorRows == 0u, L"Manual mode with matching line count and valid names should pass validation.");

    rules.manualNames = {L"uno.txt"};
    const Plan tooFew = BuildPlan({first, second}, rules);
    state.Require(tooFew.stats.errorRows == 2u, L"Manual mode should block every row when the manual line count is too small.");
    if (tooFew.rows.size() == 2u)
    {
        state.Require(HasBatchRenameIssue(tooFew.rows[0], IssueSeverity::Error, L"manual_line_count") &&
                          HasBatchRenameIssue(tooFew.rows[1], IssueSeverity::Error, L"manual_line_count"),
                      L"Too-few manual names should carry the stable manual_line_count error on every row.");
    }

    rules.manualNames = {L"uno.txt", L"dos.txt", L"tres.txt"};
    const Plan tooMany = BuildPlan({first, second}, rules);
    state.Require(tooMany.stats.errorRows == 2u, L"Manual mode should block every row when the manual line count is too large.");
    if (tooMany.rows.size() == 2u)
    {
        state.Require(HasBatchRenameIssue(tooMany.rows[0], IssueSeverity::Error, L"manual_line_count") &&
                          HasBatchRenameIssue(tooMany.rows[1], IssueSeverity::Error, L"manual_line_count"),
                      L"Too-many manual names should carry the stable manual_line_count error on every row.");
    }

    rules.manualNames = {L"uno.txt", L""};
    const Plan emptyLine = BuildPlan({first, second}, rules);
    state.Require(emptyLine.stats.errorRows == 1u, L"Manual mode should block empty target names.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineRecomputeStatsAfterContextualIssue(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target target{};
    target.sourcePath = L"C:\\root\\alpha.txt";

    Rules rules{};
    rules.nameTemplate = L"beta{ext}";

    Plan plan = BuildPlan({target}, rules);
    state.Require(plan.rows.size() == 1u, L"Stats recompute test should create one preview row.");
    if (plan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(plan.stats.changedRows == 1u && plan.stats.errorRows == 0u,
                  L"Initial stats should report one changed row and no errors.");

    AddIssue(plan.rows[0], IssueSeverity::Error, L"name_destination_exists");
    RecomputeStats(plan);

    state.Require(plan.stats.totalRows == 1u, L"Recomputed stats should preserve total row count.");
    state.Require(plan.stats.changedRows == 1u, L"Recomputed stats should preserve changed-row count.");
    state.Require(plan.stats.unchangedRows == 0u, L"Recomputed stats should preserve unchanged-row count.");
    state.Require(plan.stats.errorRows == 1u, L"Recomputed stats should count the contextual destination error.");
    state.Require(plan.stats.warningRows == 0u, L"Error rows should not also count as warning rows.");
    state.Require(HasIssueSeverity(plan.rows[0], IssueSeverity::Error),
                  L"Shared issue helper should detect the contextual error issue.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowDateTimeColumnsMatchMacroExpansion(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root   = suiteRoot / L"work" / (L"batch_rename_datetime_grid_" + NewGuidText());
    const std::filesystem::path source = root / L"timestamped.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename date/time root.");
    state.Require(SelfTest::WriteTextFile(source, "timestamp"), L"Failed to create Batch Rename date/time input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const std::chrono::sys_seconds desiredStamp =
        std::chrono::sys_days{std::chrono::year{2026} / std::chrono::January / 1} + 23h + 30min + 5s;
    const auto nowFile = std::filesystem::file_time_type::clock::now();
    const auto nowSys =
        std::chrono::time_point_cast<std::filesystem::file_time_type::duration>(std::chrono::system_clock::now());
    const auto desiredSys =
        std::chrono::time_point_cast<std::filesystem::file_time_type::duration>(desiredStamp);
    std::filesystem::last_write_time(source, nowFile + (desiredSys - nowSys), ec);
    state.Require(! ec, L"Failed to set Batch Rename date/time input timestamp.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-datetime-grid-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {source};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-datetime-grid-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for date/time grid testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename date/time grid test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"{date:yyyy-MM-dd}_{time:HH-mm-ss}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename date/time grid test should set date/time macro rules.");

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename date/time grid snapshot should be available.");
    state.Require(snapshot.previewRowCount == 1u, L"Batch Rename date/time grid preview should contain one row.");
    state.Require(snapshot.newNames.size() == 1u && snapshot.dateTexts.size() == 1u && snapshot.timeTexts.size() == 1u,
                  L"Batch Rename date/time grid snapshot should expose macro, date, and time text.");
    if (snapshot.newNames.empty() || snapshot.dateTexts.empty() || snapshot.timeTexts.empty())
    {
        return false;
    }

    const std::wstring& macroName = snapshot.newNames[0];
    state.Require(macroName.size() >= 19u && macroName[10] == L'_',
                  std::format(L"Date/time macro name should start with yyyy-MM-dd_HH-mm-ss; saw '{}'.", macroName));
    if (macroName.size() >= 19u)
    {
        const std::wstring macroDate = macroName.substr(0u, 10u);
        std::wstring macroTime       = macroName.substr(11u, 8u);
        std::ranges::replace(macroTime, L'-', L':');
        state.Require(snapshot.dateTexts[0] == macroDate,
                      std::format(L"Date column should match {{date}} macro output; grid='{}' macro='{}'.",
                                  snapshot.dateTexts[0],
                                  macroDate));
        state.Require(snapshot.timeTexts[0] == macroTime,
                      std::format(L"Time column should match {{time}} macro output; grid='{}' macro='{}'.",
                                  snapshot.timeTexts[0],
                                  macroTime));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowCreatedMacroUsesCollectedCreationTime(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root   = suiteRoot / L"work" / (L"batch_rename_created_macro_" + NewGuidText());
    const std::filesystem::path source = root / L"created-source.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename created-macro root.");
    state.Require(SelfTest::WriteTextFile(source, "created"), L"Failed to create Batch Rename created-macro input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const std::chrono::sys_seconds creationStamp =
        std::chrono::sys_days{std::chrono::year{2026} / std::chrono::March / 15} + 12h + 0min + 0s;
    state.Require(SetBatchRenameFileCreationTime(source, creationStamp),
                  L"Failed to set Batch Rename created-macro file creation time.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-created-macro-selftest";
    context.rootPluginPath  = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-created-macro-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for created-macro testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename created-macro test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"created_{created:yyyyMMdd}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename created-macro test should set created-date macro rules.");

    BatchRenameDebugSnapshot snapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(snapshot), L"Batch Rename created-macro snapshot should be available.");
    state.Require(snapshot.previewRowCount == 1u, L"Batch Rename created-macro preview should contain one row.");
    state.Require(snapshot.newNames.size() == 1u, L"Batch Rename created-macro snapshot should expose one new name.");
    if (snapshot.newNames.empty())
    {
        return false;
    }

    const BatchRenameLocalStampParts parts = GetBatchRenameExpectedLocalParts(creationStamp);
    const std::wstring expectedName =
        std::format(L"created_{:04}{:02}{:02}.txt", parts.year, parts.month, parts.day);
    state.Require(snapshot.newNames[0] == expectedName,
                  std::format(L"{{created}} macro should use collected creation time. Expected '{}', saw '{}'.",
                              expectedName,
                              snapshot.newNames[0]));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineWarningsAndRemainingTransforms(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target noOp{};
    noOp.sourcePath = L"C:\\root\\same.txt";

    Rules rules{};
    rules.nameTemplate = L"{name}";

    const Plan noOpPlan = BuildPlan({noOp}, rules);
    state.Require(noOpPlan.rows.size() == 1u, L"No-op warning plan should contain one row.");
    if (noOpPlan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(noOpPlan.stats.warningRows == 1u, L"No-op rows should be reported as warning rows.");
    state.Require(HasBatchRenameIssue(noOpPlan.rows[0], IssueSeverity::Warning, L"name_unchanged"),
                  L"No-op rows should carry a stable unchanged-name warning.");

    Target caseOnly{};
    caseOnly.sourcePath = L"C:\\root\\MiXeD.TXT";
    rules.fileNameCaseStyle = CaseTransform::Lower;
    rules.extensionCaseStyle = CaseTransform::Lower;

    const Plan caseOnlyPlan = BuildPlan({caseOnly}, rules);
    state.Require(caseOnlyPlan.rows.size() == 1u, L"Case-only warning plan should contain one row.");
    if (caseOnlyPlan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(caseOnlyPlan.rows[0].newName == L"mixed.txt", L"Lower-case transform should apply to stem and extension.");
    state.Require(caseOnlyPlan.stats.warningRows == 1u, L"Case-only renames should be warning rows.");
    state.Require(HasBatchRenameIssue(caseOnlyPlan.rows[0], IssueSeverity::Warning, L"name_case_only"),
                  L"Case-only renames should carry a stable warning.");

    Target spaced{};
    spaced.sourcePath = L"C:\\root\\spaced.txt";
    rules = {};
    rules.mode = Mode::Manual;
    rules.manualNames = {L" leading.txt"};

    const Plan spacedPlan = BuildPlan({spaced}, rules);
    state.Require(spacedPlan.rows.size() == 1u, L"Leading-space warning plan should contain one row.");
    if (spacedPlan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(spacedPlan.stats.warningRows == 1u, L"Names with leading spaces should be warning rows.");
    state.Require(HasBatchRenameIssue(spacedPlan.rows[0], IssueSeverity::Warning, L"name_edge_space_or_dot"),
                  L"Names with leading spaces should carry an edge-space warning.");

    Target trailingDot{};
    trailingDot.sourcePath = L"C:\\root\\trailing.txt";
    rules.manualNames = {L"trailing."};
    const Plan trailingDotPlan = BuildPlan({trailingDot}, rules);
    state.Require(trailingDotPlan.rows.size() == 1u, L"Trailing-dot warning plan should contain one row.");
    if (trailingDotPlan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(trailingDotPlan.stats.warningRows == 1u, L"Names with trailing dots should be warning rows.");
    state.Require(HasBatchRenameIssue(trailingDotPlan.rows[0], IssueSeverity::Warning, L"name_edge_space_or_dot"),
                  L"Names with trailing dots should carry an edge-space warning.");

    Target title{};
    title.sourcePath = L"C:\\root\\alpha-beta_gamma.txt";
    rules = {};
    rules.nameTemplate = L"{stem}{ext}";
    rules.fileNameCaseStyle = CaseTransform::Mixed;
    const Plan mixedPlan = BuildPlan({title}, rules);
    state.Require(mixedPlan.rows.size() == 1u, L"Mixed-case transform plan should contain one row.");
    if (mixedPlan.rows.size() != 1u)
    {
        return false;
    }
    state.Require(mixedPlan.rows[0].newName == L"Alpha-Beta_Gamma.txt",
                  L"Mixed-case transform should title-case words separated by punctuation.");

    Target replace{};
    replace.sourcePath = L"C:\\root\\cat cat scatter.txt";
    rules = {};
    rules.nameTemplate = L"{name}";
    rules.searchFor = L"cat";
    rules.replaceWith = L"dog";
    rules.replaceOnce = true;
    const Plan replaceOncePlan = BuildPlan({replace}, rules);
    state.Require(replaceOncePlan.rows.size() == 1u && replaceOncePlan.rows[0].newName == L"dog cat scatter.txt",
                  L"Replace-once should update only the first literal match.");

    rules.replaceOnce = false;
    rules.wholeWords = true;
    const Plan wholeWordPlan = BuildPlan({replace}, rules);
    state.Require(wholeWordPlan.rows.size() == 1u && wholeWordPlan.rows[0].newName == L"dog dog scatter.txt",
                  L"Whole-word literal replacement should skip embedded word fragments.");

    std::wstring emoji;
    emoji.push_back(static_cast<wchar_t>(0xD83D));
    emoji.push_back(static_cast<wchar_t>(0xDE00));
    Target emojiTarget{};
    emojiTarget.sourcePath = std::filesystem::path(L"C:\\root") / (emoji + L".txt");
    rules = {};
    rules.nameTemplate = L"{name}";
    rules.searchFor = std::wstring(1u, emoji.front());
    rules.replaceWith = L"X";
    rules.wholeWords = true;
    const Plan surrogateBoundaryPlan = BuildPlan({emojiTarget}, rules);
    state.Require(surrogateBoundaryPlan.rows.size() == 1u && surrogateBoundaryPlan.rows[0].newName == emoji + L".txt",
                  L"Whole-word literal replacement must not split a supplementary-plane surrogate pair.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineRemainingValidationAndTransformCoverage(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target regexTarget{};
    regexTarget.sourcePath = L"C:\\root\\cat scatter cat.txt";

    Rules regexRules{};
    regexRules.nameTemplate = L"{name}";
    regexRules.searchFor = L"cat";
    regexRules.replaceWith = L"dog";
    regexRules.regexEnabled = true;
    regexRules.wholeWords = true;
    regexRules.excludeExtension = true;

    const Plan wholeWordRegexPlan = BuildPlan({regexTarget}, regexRules);
    state.Require(wholeWordRegexPlan.rows.size() == 1u && wholeWordRegexPlan.rows[0].newName == L"dog scatter dog.txt",
                  L"Whole-word regex replacement should update standalone matches and skip embedded word fragments.");
    state.Require(wholeWordRegexPlan.stats.errorRows == 0u, L"Whole-word regex replacement should not report validation errors.");

    Target stemOnly{};
    stemOnly.sourcePath = L"C:\\root\\MiXeD.TXT";

    Rules stemCaseRules{};
    stemCaseRules.nameTemplate = L"{name}";
    stemCaseRules.fileNameCaseStyle = CaseTransform::Lower;

    const Plan lowerStemPlan = BuildPlan({stemOnly}, stemCaseRules);
    state.Require(lowerStemPlan.rows.size() == 1u && lowerStemPlan.rows[0].newName == L"mixed.TXT",
                  L"Lower-case stem-only transform should leave the extension spelling unchanged.");
    state.Require(lowerStemPlan.stats.errorRows == 0u, L"Lower-case stem-only transform should not report validation errors.");

    Target noExtension{};
    noExtension.sourcePath = L"C:\\root\\README";

    Rules extensionNoOpRules{};
    extensionNoOpRules.nameTemplate = L"{name}";
    extensionNoOpRules.extensionCaseStyle = CaseTransform::Lower;

    const Plan noExtensionPlan = BuildPlan({noExtension}, extensionNoOpRules);
    state.Require(noExtensionPlan.rows.size() == 1u && noExtensionPlan.rows[0].newName == L"README",
                  L"Extension-only case transform should not add punctuation to names without extensions.");
    state.Require(noExtensionPlan.stats.warningRows == 1u, L"No-extension extension-only transform should remain a no-op warning.");
    if (noExtensionPlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(noExtensionPlan.rows[0], IssueSeverity::Warning, L"name_unchanged"),
                      L"No-extension extension-only transform should carry the unchanged-name warning.");
    }

    Target invalidManual{};
    invalidManual.sourcePath = L"C:\\root\\source.txt";

    Rules invalidRules{};
    invalidRules.mode = Mode::Manual;
    invalidRules.manualNames = {L"bad\\leaf.txt"};

    const Plan invalidSeparatorPlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(invalidSeparatorPlan.rows.size() == 1u, L"Invalid separator validation plan should contain one row.");
    state.Require(invalidSeparatorPlan.stats.errorRows == 1u, L"Manual target leaves containing path separators should be blocking errors.");
    if (invalidSeparatorPlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(invalidSeparatorPlan.rows[0], IssueSeverity::Error, L"name_separator"),
                      L"Manual target leaves containing path separators should carry the stable name_separator error.");
    }

    invalidRules.manualNames = {L"bad:name.txt"};
    const Plan invalidCharacterPlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(invalidCharacterPlan.rows.size() == 1u, L"Invalid-character validation plan should contain one row.");
    state.Require(invalidCharacterPlan.stats.errorRows == 1u, L"Manual target leaves containing Windows-invalid characters should be blocking errors.");
    if (invalidCharacterPlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(invalidCharacterPlan.rows[0], IssueSeverity::Error, L"name_invalid_character"),
                      L"Manual target leaves containing Windows-invalid characters should carry the stable name_invalid_character error.");
    }

    invalidRules.manualNames = {std::wstring(255u, L'a')};
    const Plan maxLengthPlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(maxLengthPlan.rows.size() == 1u, L"Max-length target leaf validation plan should contain one row.");
    state.Require(maxLengthPlan.stats.errorRows == 0u, L"Manual target leaves with exactly 255 UTF-16 code units should be valid.");
    if (maxLengthPlan.rows.size() == 1u)
    {
        state.Require(! HasBatchRenameIssue(maxLengthPlan.rows[0], IssueSeverity::Error, L"name_too_long"),
                      L"Manual target leaves with exactly 255 UTF-16 code units should not carry name_too_long.");
    }

    invalidRules.manualNames = {std::wstring(256u, L'a')};
    const Plan tooLongPlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(tooLongPlan.rows.size() == 1u, L"Overlong target leaf validation plan should contain one row.");
    state.Require(tooLongPlan.stats.errorRows == 1u, L"Manual target leaves with 256 UTF-16 code units should be blocking errors.");
    if (tooLongPlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(tooLongPlan.rows[0], IssueSeverity::Error, L"name_too_long"),
                      L"Manual target leaves with 256 UTF-16 code units should carry the stable name_too_long error.");
    }

    std::wstring surrogatePair;
    surrogatePair.push_back(static_cast<wchar_t>(0xD83D));
    surrogatePair.push_back(static_cast<wchar_t>(0xDE00));

    invalidRules.manualNames = {std::wstring(253u, L'a') + surrogatePair};
    const Plan maxLengthSurrogatePlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(maxLengthSurrogatePlan.rows.size() == 1u, L"Max-length surrogate target leaf plan should contain one row.");
    state.Require(maxLengthSurrogatePlan.stats.errorRows == 0u,
                  L"A 255-code-unit leaf ending in a surrogate pair should remain valid.");
    if (maxLengthSurrogatePlan.rows.size() == 1u)
    {
        state.Require(! HasBatchRenameIssue(maxLengthSurrogatePlan.rows[0], IssueSeverity::Error, L"name_too_long"),
                      L"A 255-code-unit leaf ending in a surrogate pair should not carry name_too_long.");
    }

    invalidRules.manualNames = {std::wstring(254u, L'a') + surrogatePair};
    const Plan tooLongSurrogatePlan = BuildPlan({invalidManual}, invalidRules);
    state.Require(tooLongSurrogatePlan.rows.size() == 1u, L"Overlong surrogate target leaf plan should contain one row.");
    state.Require(tooLongSurrogatePlan.stats.errorRows == 1u,
                  L"A 256-code-unit leaf ending in a surrogate pair should be a blocking error.");
    if (tooLongSurrogatePlan.rows.size() == 1u)
    {
        state.Require(HasBatchRenameIssue(tooLongSurrogatePlan.rows[0], IssueSeverity::Error, L"name_too_long"),
                      L"A 256-code-unit leaf ending in a surrogate pair should carry name_too_long.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineRegexMatchFailureAndWholeWordCaptures(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target pathological{};
    pathological.sourcePath = std::filesystem::path(L"C:\\root") / (std::wstring(64u, L'a') + L".txt");

    Rules pathologicalRules{};
    pathologicalRules.nameTemplate     = L"{name}";
    pathologicalRules.searchFor        = L"(a+)+b";
    pathologicalRules.replaceWith      = L"x";
    pathologicalRules.regexEnabled     = true;
    pathologicalRules.excludeExtension = true;

    const Plan pathologicalPlan = BuildPlan({pathological}, pathologicalRules);
    state.Require(pathologicalPlan.rows.size() == 1u, L"Pathological regex plan should contain one preview row.");
    if (pathologicalPlan.rows.size() != 1u)
    {
        return false;
    }

    state.Require(HasBatchRenameIssue(pathologicalPlan.rows[0], IssueSeverity::Error, L"regex_match_failed"),
                  L"Backtracking-heavy regex patterns that fail at match time should report the stable regex_match_failed error instead of crashing the preview.");
    state.Require(pathologicalPlan.rows[0].newName == pathologicalPlan.rows[0].originalName,
                  L"Rows with match-time regex failures should keep the original leaf name.");
    state.Require(pathologicalPlan.stats.errorRows == 1u, L"Match-time regex failures should be blocking validation errors.");

    Target grouped{};
    grouped.sourcePath = L"C:\\root\\12-34.txt";

    Rules groupRules{};
    groupRules.nameTemplate     = L"{name}";
    groupRules.searchFor        = L"(\\d+)-(\\d+)";
    groupRules.replaceWith      = L"$2_$1";
    groupRules.regexEnabled     = true;
    groupRules.wholeWords       = true;
    groupRules.excludeExtension = true;

    const Plan groupPlan = BuildPlan({grouped}, groupRules);
    state.Require(groupPlan.rows.size() == 1u, L"Whole-word capture-group plan should contain one preview row.");
    if (groupPlan.rows.size() == 1u)
    {
        state.Require(groupPlan.rows[0].newName == L"34_12.txt",
                      L"Whole-word regex wrapping should be non-capturing so $1/$2 keep their user-pattern indexes.");
        state.Require(groupPlan.stats.errorRows == 0u, L"Whole-word capture-group replacement should not report validation errors.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineReservedDeviceNamesAndCounterWidth(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target source{};
    source.sourcePath = L"C:\\root\\source.txt";

    Rules manualRules{};
    manualRules.mode = Mode::Manual;

    manualRules.manualNames = {L"CON.txt"};
    const Plan conPlan = BuildPlan({source}, manualRules);
    state.Require(conPlan.rows.size() == 1u && HasBatchRenameIssue(conPlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"CON with an extension should carry the stable name_reserved_device blocking error.");

    manualRules.manualNames = {L"NUL"};
    const Plan nulPlan = BuildPlan({source}, manualRules);
    state.Require(nulPlan.rows.size() == 1u && HasBatchRenameIssue(nulPlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"Bare NUL should carry the stable name_reserved_device blocking error.");

    manualRules.manualNames = {L"com1.log"};
    const Plan comPlan = BuildPlan({source}, manualRules);
    state.Require(comPlan.rows.size() == 1u && HasBatchRenameIssue(comPlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"Lower-case COM1 with an extension should carry the stable name_reserved_device blocking error.");

    manualRules.manualNames = {L"CONIN$.txt"};
    const Plan coninPlan = BuildPlan({source}, manualRules);
    state.Require(coninPlan.rows.size() == 1u && HasBatchRenameIssue(coninPlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"CONIN$ with an extension should carry the stable name_reserved_device blocking error.");

    manualRules.manualNames = {L"conout$"};
    const Plan conoutPlan = BuildPlan({source}, manualRules);
    state.Require(conoutPlan.rows.size() == 1u && HasBatchRenameIssue(conoutPlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"Lower-case CONOUT$ should carry the stable name_reserved_device blocking error.");

    manualRules.manualNames = {L"CONSOLE.txt"};
    const Plan consolePlan = BuildPlan({source}, manualRules);
    state.Require(consolePlan.rows.size() == 1u && ! HasBatchRenameIssue(consolePlan.rows[0], IssueSeverity::Error, L"name_reserved_device"),
                  L"Names that merely start with a reserved token should not be blocked.");
    state.Require(consolePlan.stats.errorRows == 0u, L"CONSOLE.txt should remain a valid leaf name.");

    Target first{};
    first.sourcePath = L"C:\\root\\one.txt";
    Target second{};
    second.sourcePath = L"C:\\root\\two.txt";

    Rules counterRules{};
    counterRules.nameTemplate = L"{counter:3}_{name}";

    const Plan counterPlan = BuildPlan({first, second}, counterRules);
    state.Require(counterPlan.rows.size() == 2u, L"Counter-width plan should contain one row per target.");
    if (counterPlan.rows.size() == 2u)
    {
        state.Require(counterPlan.rows[0].newName == L"001_one.txt" && counterPlan.rows[1].newName == L"002_two.txt",
                      L"A plain numeric counter width such as {counter:3} should zero-pad to that width.");
    }
    state.Require(counterPlan.stats.errorRows == 0u, L"Counter-width macros should not report validation errors.");

    counterRules.nameTemplate = L"bad-{counter:abc}_{index:x}_{name}";
    const Plan invalidFormatPlan = BuildPlan({first, second}, counterRules);
    state.Require(invalidFormatPlan.rows.size() == 2u, L"Invalid counter/index format plan should contain one row per target.");
    if (invalidFormatPlan.rows.size() == 2u)
    {
        state.Require(HasBatchRenameIssue(invalidFormatPlan.rows[0], IssueSeverity::Error, L"macro_invalid_format") &&
                          HasBatchRenameIssue(invalidFormatPlan.rows[1], IssueSeverity::Error, L"macro_invalid_format"),
                      L"Nonnumeric counter/index formats should carry the stable macro_invalid_format blocking error.");
    }
    state.Require(invalidFormatPlan.stats.errorRows == 2u, L"Nonnumeric counter/index formats should be blocking validation errors.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineLargePreviewEmitsPerfMetrics(CaseState& state) noexcept
{
    using namespace BatchRename;
    using namespace std::chrono_literals;

    constexpr size_t kTargetCount = 10'000u;
    constexpr uint64_t kMaxLargePreviewBuildPlanBaseMs = 5'000u;

    std::vector<Target> targets;
    targets.reserve(kTargetCount);
    for (size_t i = 0; i < kTargetCount; ++i)
    {
        Target target{};
        target.sourcePath    = std::filesystem::path(L"C:\\batch-preview") / std::format(L"Episode {:05}.RAW", i);
        target.sizeBytes     = 1024u + i;
        target.lastWriteTime = std::chrono::sys_days{std::chrono::year{2026} / std::chrono::June / 10} + std::chrono::seconds{i};
        targets.push_back(std::move(target));
    }

    const uint64_t buildPerfRowsBefore   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t countPerfRowsBefore   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.rows");
    const uint64_t changedPerfRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.preview.changed");
    const uint64_t errorPerfRowsBefore   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.errors");

    Rules rules{};
    rules.nameTemplate       = L"{counter:00000}_{stem}_{date:yyyyMMdd}_{time:HHmmss}{ext}";
    rules.searchFor          = L"episode";
    rules.replaceWith        = L"clip";
    rules.caseSensitive      = false;
    rules.excludeExtension   = true;
    rules.fileNameCaseStyle  = CaseTransform::Mixed;
    rules.extensionCaseStyle = CaseTransform::Lower;

    const Plan plan = BuildPlan(targets, rules);

    state.Require(plan.rows.size() == kTargetCount, L"Large Batch Rename preview should preserve one row per synthetic target.");
    state.Require(plan.stats.totalRows == kTargetCount, L"Large Batch Rename preview stats should count every synthetic target.");
    state.Require(plan.stats.changedRows == kTargetCount, L"Large Batch Rename preview should transform every synthetic target.");
    state.Require(plan.stats.errorRows == 0u, L"Large Batch Rename preview should remain free of blocking errors.");
    if (plan.rows.size() == kTargetCount)
    {
        const auto expectedRowName = [](const size_t index)
        {
            const std::chrono::sys_seconds stamp =
                std::chrono::sys_days{std::chrono::year{2026} / std::chrono::June / 10} + std::chrono::seconds{index};
            const BatchRenameLocalStampParts parts = GetBatchRenameExpectedLocalParts(stamp);
            return std::format(L"{:05}_Clip {:05}_{:04}{:02}{:02}_{:02}{:02}{:02}.raw",
                               index + 1u,
                               index,
                               parts.year,
                               parts.month,
                               parts.day,
                               parts.hour,
                               parts.minute,
                               parts.second);
        };
        state.Require(plan.rows.front().newName == expectedRowName(0u),
                      L"Large preview first row should combine counter, literal replacement, local date/time, and case transforms.");
        state.Require(plan.rows.back().newName == expectedRowName(kTargetCount - 1u),
                      L"Large preview last row should keep deterministic counter and local-time expansion.");
    }

    const uint64_t buildPerfRowsAfter   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t countPerfRowsAfter   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.rows");
    const uint64_t changedPerfRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.preview.changed");
    const uint64_t errorPerfRowsAfter   = CountBatchRenamePerfRowsWithMetric("batchrename.preview.errors");
    state.Require(buildPerfRowsAfter > buildPerfRowsBefore,
                  std::format(L"Large Batch Rename preview should emit batchrename.preview.build_plan_us perf metrics; before={} after={}.",
                              buildPerfRowsBefore,
                              buildPerfRowsAfter));
    state.Require(countPerfRowsAfter > countPerfRowsBefore,
                  std::format(L"Large Batch Rename preview should emit batchrename.preview.rows perf metrics; before={} after={}.",
                              countPerfRowsBefore,
                              countPerfRowsAfter));
    state.Require(changedPerfRowsAfter > changedPerfRowsBefore,
                  std::format(L"Large Batch Rename preview should emit batchrename.preview.changed perf metrics; before={} after={}.",
                              changedPerfRowsBefore,
                              changedPerfRowsAfter));
    state.Require(errorPerfRowsAfter > errorPerfRowsBefore,
                  std::format(L"Large Batch Rename preview should emit batchrename.preview.errors perf metrics; before={} after={}.",
                              errorPerfRowsBefore,
                              errorPerfRowsAfter));
    // Consider only the build_plan_us rows emitted by this test (skip rows that existed before),
    // and scale the threshold with the suite timeout scale used for other timing-sensitive waits.
    const uint64_t maxBuildPlanThresholdUs = SelfTest::ScaleTimeout(kMaxLargePreviewBuildPlanBaseMs) * 1'000u;
    const std::optional<uint64_t> maxBuildPlanUs =
        TryReadMaxBatchRenamePerfDurationUs("batchrename.preview.build_plan_us", buildPerfRowsBefore);
    state.Require(maxBuildPlanUs.has_value() && maxBuildPlanUs.value() <= maxBuildPlanThresholdUs,
                  std::format(L"Large Batch Rename preview should stay within the bounded preview metric threshold; maxUs={} thresholdUs={}.",
                              maxBuildPlanUs.value_or(0u),
                              maxBuildPlanThresholdUs));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineRegexCompileEmitsPerfMetric(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target first{};
    first.sourcePath = L"C:\\root\\Episode 01.txt";
    Target second{};
    second.sourcePath = L"C:\\root\\Episode 02.txt";

    Rules rules{};
    rules.nameTemplate  = L"{name}";
    rules.searchFor     = L"Episode (\\d+)";
    rules.replaceWith   = L"Clip $1";
    rules.regexEnabled  = true;
    rules.caseSensitive = true;

    const uint64_t regexCompileRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.regex.compile.us");
    const Plan plan = BuildPlan({first, second}, rules);
    const uint64_t regexCompileRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.regex.compile.us");

    state.Require(plan.rows.size() == 2u, L"Regex compile perf test should keep one row per target.");
    if (plan.rows.size() == 2u)
    {
        state.Require(plan.rows[0].newName == L"Clip 01.txt" && plan.rows[1].newName == L"Clip 02.txt",
                      L"Regex compile perf test should still apply replacement with one compiled regex.");
    }
    state.Require(plan.stats.errorRows == 0u, L"Valid regex compile perf test should not report validation errors.");
    state.Require(regexCompileRowsAfter > regexCompileRowsBefore,
                  std::format(L"Regex-enabled Batch Rename preview should emit batchrename.regex.compile.us; before={} after={}.",
                              regexCompileRowsBefore,
                              regexCompileRowsAfter));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameEngineValidationEmitsPerfMetric(CaseState& state) noexcept
{
    using namespace BatchRename;

    Target first{};
    first.sourcePath = L"C:\\root\\alpha.txt";
    Target second{};
    second.sourcePath = L"C:\\root\\beta.txt";

    Rules rules{};
    rules.nameTemplate = L"same.txt";

    const uint64_t validationRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.validation.us");
    const Plan plan = BuildPlan({first, second}, rules);
    const uint64_t validationRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.validation.us");

    state.Require(plan.rows.size() == 2u, L"Validation perf test should keep one row per target.");
    state.Require(plan.stats.errorRows == 2u, L"Validation perf test should still mark duplicate target names as errors.");
    state.Require(validationRowsAfter > validationRowsBefore,
                  std::format(L"Batch Rename preview validation should emit batchrename.validation.us; before={} after={}.",
                              validationRowsBefore,
                              validationRowsAfter));

    return state.failure.empty();
}

struct BatchRenameExecutionProgressEvent final
{
    uint64_t completed = 0u;
    uint64_t total     = 0u;
    bool forcePost     = false;
};

struct BatchRenameExecutionProgressRecorder final
{
    std::vector<BatchRenameExecutionProgressEvent> events;
};

void RecordBatchRenameExecutionProgress(void* context, const uint64_t completedItems, const uint64_t totalItems, const bool forcePost) noexcept
{
    auto* recorder = static_cast<BatchRenameExecutionProgressRecorder*>(context);
    if (! recorder)
    {
        return;
    }

    recorder->events.push_back(BatchRenameExecutionProgressEvent{
        .completed = completedItems,
        .total     = totalItems,
        .forcePost = forcePost,
    });
}

[[nodiscard]] bool TestBatchRenameExecutionEngineDirectScenarios(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_execution_engine_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename execution-engine root.");
    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    const FileSystemPathIdentity pathIdentity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    std::atomic_bool cancelRequested{false};
    const auto makeOp = [](const std::filesystem::path& source, std::wstring finalLeaf, const bool isDirectory = false) noexcept
    {
        BatchRenameExecutionOp op{};
        op.currentSource  = source;
        op.originalSource = source;
        op.finalLeaf      = std::move(finalLeaf);
        op.depth          = PathDepthKey(source);
        op.isDirectory    = isDirectory;
        return op;
    };
    const auto sortOps = [](std::vector<BatchRenameExecutionOp>& ops) noexcept
    {
        std::sort(ops.begin(), ops.end(), [](const BatchRenameExecutionOp& lhs, const BatchRenameExecutionOp& rhs) noexcept
        {
            if (lhs.depth != rhs.depth)
            {
                return lhs.depth > rhs.depth;
            }
            if (lhs.currentSource.native().size() != rhs.currentSource.native().size())
            {
                return lhs.currentSource.native().size() > rhs.currentSource.native().size();
            }
            return lhs.currentSource.native() > rhs.currentSource.native();
        });
    };
    const auto runEngine = [&](wil::com_ptr<IFileSystem> fs,
                               std::vector<BatchRenameExecutionOp> ops,
                               BatchRenameExecutionOptions options = {}) noexcept
    {
        sortOps(ops);
        cancelRequested.store(false, std::memory_order_release);
        return RunBatchRenameExecutionEngine(cancelRequested, *fs.get(), pathIdentity, std::move(ops), options);
    };

    const std::filesystem::path swapRoot = root / L"swap";
    state.Require(SelfTest::EnsureDirectory(swapRoot), L"Failed to create direct engine swap root.");
    state.Require(SelfTest::WriteTextFile(swapRoot / L"swap-a.txt", "alpha-content"), L"Failed to create direct engine swap-a.");
    state.Require(SelfTest::WriteTextFile(swapRoot / L"swap-b.txt", "beta-content"), L"Failed to create direct engine swap-b.");
    std::atomic_uint32_t swapRenameItemsCalls{0u};
    wil::com_ptr<IFileSystem> swapFs = CreateBatchRenameRenameCountingFileSystem(fileSystem, &swapRenameItemsCalls);
    state.Require(swapFs != nullptr, L"Direct engine swap should create a counting provider.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenameExecutionProgressRecorder swapProgress{};
    BatchRenameExecutionOptions swapOptions{
        .progressCallback = RecordBatchRenameExecutionProgress,
        .progressContext  = &swapProgress,
    };
    BatchRenameExecutionResult swapResult =
        runEngine(swapFs, {makeOp(swapRoot / L"swap-a.txt", L"swap-b.txt"), makeOp(swapRoot / L"swap-b.txt", L"swap-a.txt")}, swapOptions);
    state.Require(SUCCEEDED(swapResult.hr), std::format(L"Direct engine swap should succeed: 0x{:08X}.", static_cast<unsigned long>(swapResult.hr)));
    state.Require(swapResult.report.completedRows == 2u && swapResult.report.failedRows == 0u,
                  L"Direct engine swap should report two completed rows and no failures.");
    state.Require(swapResult.report.undoEntries.size() == 2u, L"Direct engine swap should record one undo entry per row.");
    state.Require(swapRenameItemsCalls.load(std::memory_order_relaxed) == 3u,
                  std::format(L"Direct engine swap should use one temp hop plus two layers; saw {} RenameItems calls.",
                              swapRenameItemsCalls.load(std::memory_order_relaxed)));
    state.Require(ReadBatchRenameFileText(swapRoot / L"swap-a.txt") == "beta-content",
                  L"Direct engine swap should move beta content to swap-a.");
    state.Require(ReadBatchRenameFileText(swapRoot / L"swap-b.txt") == "alpha-content",
                  L"Direct engine swap should move alpha content to swap-b.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(swapRoot), L"Direct engine swap must not leave temp-hop leaves.");
    state.Require(! swapProgress.events.empty(), L"Direct engine swap should report execution progress.");
    state.Require(std::ranges::all_of(swapProgress.events, [](const BatchRenameExecutionProgressEvent& event) noexcept {
                      return event.completed <= event.total && event.total == 2u;
                  }),
                  L"Direct engine swap progress should clamp temp-hop progress to the visible row total.");
    state.Require(std::ranges::any_of(swapProgress.events, [](const BatchRenameExecutionProgressEvent& event) noexcept {
                      return event.completed == 2u && event.total == 2u && event.forcePost;
                  }),
                  L"Direct engine swap progress should post a terminal completed/total update.");

    const std::filesystem::path cycleRoot = root / L"cycle";
    state.Require(SelfTest::EnsureDirectory(cycleRoot), L"Failed to create direct engine cycle root.");
    state.Require(SelfTest::WriteTextFile(cycleRoot / L"cycle-a.txt", "a"), L"Failed to create direct engine cycle-a.");
    state.Require(SelfTest::WriteTextFile(cycleRoot / L"cycle-b.txt", "b"), L"Failed to create direct engine cycle-b.");
    state.Require(SelfTest::WriteTextFile(cycleRoot / L"cycle-c.txt", "c"), L"Failed to create direct engine cycle-c.");
    std::atomic_uint32_t cycleRenameItemsCalls{0u};
    wil::com_ptr<IFileSystem> cycleFs = CreateBatchRenameRenameCountingFileSystem(fileSystem, &cycleRenameItemsCalls);
    state.Require(cycleFs != nullptr, L"Direct engine cycle should create a counting provider.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenameExecutionResult cycleResult = runEngine(cycleFs,
                                                       {makeOp(cycleRoot / L"cycle-a.txt", L"cycle-b.txt"),
                                                        makeOp(cycleRoot / L"cycle-b.txt", L"cycle-c.txt"),
                                                        makeOp(cycleRoot / L"cycle-c.txt", L"cycle-a.txt")});
    state.Require(SUCCEEDED(cycleResult.hr), std::format(L"Direct engine three-cycle should succeed: 0x{:08X}.",
                                                         static_cast<unsigned long>(cycleResult.hr)));
    state.Require(cycleResult.report.completedRows == 3u && cycleResult.report.failedRows == 0u,
                  L"Direct engine three-cycle should report three completed rows and no failures.");
    state.Require(cycleRenameItemsCalls.load(std::memory_order_relaxed) == 4u,
                  std::format(L"Direct engine three-cycle should use one temp hop plus three layers; saw {} RenameItems calls.",
                              cycleRenameItemsCalls.load(std::memory_order_relaxed)));
    state.Require(ReadBatchRenameFileText(cycleRoot / L"cycle-a.txt") == "c", L"Direct engine cycle should rotate c to cycle-a.");
    state.Require(ReadBatchRenameFileText(cycleRoot / L"cycle-b.txt") == "a", L"Direct engine cycle should rotate a to cycle-b.");
    state.Require(ReadBatchRenameFileText(cycleRoot / L"cycle-c.txt") == "b", L"Direct engine cycle should rotate b to cycle-c.");

    const std::filesystem::path partialRoot = root / L"partial";
    state.Require(SelfTest::EnsureDirectory(partialRoot), L"Failed to create direct engine partial root.");
    state.Require(SelfTest::WriteTextFile(partialRoot / L"ok.txt", "ok"), L"Failed to create direct engine partial ok input.");
    state.Require(SelfTest::WriteTextFile(partialRoot / L"fail.txt", "fail"), L"Failed to create direct engine partial fail input.");
    std::atomic_uint32_t partialRenameItemsCalls{0u};
    const HRESULT failHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    wil::com_ptr<IFileSystem> partialFs =
        CreateBatchRenameScriptedPerItemFileSystem(fileSystem, &partialRenameItemsCalls, L"fail.txt", failHr, false);
    state.Require(partialFs != nullptr, L"Direct engine partial failure should create a scripted provider.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenameExecutionResult partialResult = runEngine(partialFs,
                                                         {makeOp(partialRoot / L"ok.txt", L"ok-renamed.txt"),
                                                          makeOp(partialRoot / L"fail.txt", L"fail-renamed.txt")});
    state.Require(partialResult.hr == failHr,
                  std::format(L"Direct engine partial failure should surface access denied; saw 0x{:08X}.",
                              static_cast<unsigned long>(partialResult.hr)));
    state.Require(partialResult.report.completedRows == 1u && partialResult.report.failedRows == 1u,
                  L"Direct engine partial failure should report one completed row and one failed row.");
    state.Require(partialResult.report.undoEntries.size() == 1u, L"Direct engine partial failure should record undo only for completed rows.");
    state.Require(ReadBatchRenameFileText(partialRoot / L"ok-renamed.txt") == "ok",
                  L"Direct engine partial failure should keep the successful rename.");
    state.Require(ReadBatchRenameFileText(partialRoot / L"fail.txt") == "fail",
                  L"Direct engine partial failure should leave the failed source in place.");

    const std::filesystem::path directoryRoot = root / L"directory";
    const std::filesystem::path parentSource = directoryRoot / L"parent";
    state.Require(SelfTest::EnsureDirectory(parentSource), L"Failed to create direct engine parent directory.");
    state.Require(SelfTest::WriteTextFile(parentSource / L"child.txt", "child"), L"Failed to create direct engine child input.");
    std::atomic_uint32_t directoryRenameItemsCalls{0u};
    wil::com_ptr<IFileSystem> directoryFs = CreateBatchRenameRenameCountingFileSystem(fileSystem, &directoryRenameItemsCalls);
    state.Require(directoryFs != nullptr, L"Direct engine directory rewrite should create a counting provider.");
    if (! state.failure.empty())
    {
        return false;
    }

    BatchRenameExecutionResult directoryResult =
        runEngine(directoryFs, {makeOp(parentSource / L"child.txt", L"grandchild.txt"), makeOp(parentSource, L"renamed-parent", true)});
    const std::filesystem::path finalChild = directoryRoot / L"renamed-parent" / L"grandchild.txt";
    state.Require(SUCCEEDED(directoryResult.hr),
                  std::format(L"Direct engine parent/child rename should succeed: 0x{:08X}.",
                              static_cast<unsigned long>(directoryResult.hr)));
    state.Require(ReadBatchRenameFileText(finalChild) == "child", L"Direct engine parent/child rename should move the child under the final parent.");
    state.Require(directoryResult.executedDirectoryMoves.size() == 1u, L"Direct engine should record the parent directory move.");
    state.Require(directoryResult.report.undoEntries.size() == 2u, L"Direct engine parent/child rename should record two undo entries.");
    const bool childUndoRewritten = std::ranges::any_of(directoryResult.report.undoEntries,
                                                        [&finalChild](const BatchRenameUndoEntry& entry) noexcept
    {
        return entry.originalPath.filename().native() == L"child.txt" && entry.currentPath == finalChild;
    });
    state.Require(childUndoRewritten, L"Direct engine should rewrite child undo currentPath after the parent directory move.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesSwapRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_swap_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename swap root.");
    state.Require(SelfTest::WriteTextFile(root / L"swap-a.txt", "alpha-content"), L"Failed to create first Batch Rename swap input.");
    state.Require(SelfTest::WriteTextFile(root / L"swap-b.txt", "beta-content"), L"Failed to create second Batch Rename swap input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> countingFileSystem = CreateBatchRenameRenameCountingFileSystem(fileSystem, &renameItemsCalls);
    state.Require(countingFileSystem != nullptr, L"Batch Rename swap selftest should create a rename-counting file-system wrapper.");
    if (! countingFileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = countingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-swap-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"swap-a.txt", root / L"swap-b.txt"};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-swap-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for swap execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename swap test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual), L"Batch Rename swap test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"swap-b.txt\nswap-a.txt"),
                  L"Batch Rename swap test should set crossed manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename swap preview should be available before executing.");
    state.Require(before.changedRowCount == 2u, L"Swap preview should treat both crossed rows as changed.");
    state.Require(before.errorRowCount == 0u, L"Swap preview must not flag planned-source destinations as collisions.");
    state.Require(before.renameButtonEnabled, L"Valid swap preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename swap execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(ReadBatchRenameFileText(root / L"swap-a.txt") == "beta-content",
                  L"Swap execution should leave the second file's content under the first name.");
    state.Require(ReadBatchRenameFileText(root / L"swap-b.txt") == "alpha-content",
                  L"Swap execution should leave the first file's content under the second name.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Swap execution must not leave .rsren- temp files behind.");
    state.Require(ListBatchRenameDirectoryLeaves(root) == std::vector<std::wstring>{L"swap-a.txt", L"swap-b.txt"},
                  L"Swap execution should leave exactly the two swapped names in the directory.");
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 3u,
                  std::format(L"Swap execution should break the cycle with one temp hop and two dependency layers; saw {} RenameItems calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename swap snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 2u,
                  std::format(L"Swap execution should report two completed rows; saw {}.", after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionFailedRows == 0u, L"Swap execution should report zero failed rows.");
    state.Require(after.lastExecutionUndoRowCount == 2u, L"Swap execution should record one undo entry per row.");
    state.Require(! after.lastExecutionCanceled, L"Swap execution should not be marked canceled.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Batch Rename swap test should copy the retained undo plan.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains((root / L"swap-b.txt").native() + L"\tswap-a.txt\t" + (root / L"swap-a.txt").native()),
                  std::format(L"Swap undo plan should record the net first-row transition; saw '{}'.", undoReport));
    state.Require(undoReport.contains((root / L"swap-a.txt").native() + L"\tswap-b.txt\t" + (root / L"swap-b.txt").native()),
                  std::format(L"Swap undo plan should record the net second-row transition; saw '{}'.", undoReport));
    state.Require(! undoReport.contains(L".rsren-"), L"Swap undo plan must not leak temp hop names.");

    state.Require(callbackCalls == 1u, std::format(L"Swap execution should invoke the success callback once; saw {}.", callbackCalls));
    state.Require(callbackSources.size() == 2u && callbackTargets.size() == 2u,
                  L"Swap success callback should report both renamed rows.");
    const auto hasPair = [&](const std::filesystem::path& source, const std::filesystem::path& target) noexcept
    {
        for (size_t index = 0u; index < callbackSources.size() && index < callbackTargets.size(); ++index)
        {
            if (callbackSources[index] == source && callbackTargets[index] == target)
            {
                return true;
            }
        }
        return false;
    };
    state.Require(hasPair(root / L"swap-a.txt", root / L"swap-b.txt"),
                  L"Swap success callback should pair the first original with its swapped target.");
    state.Require(hasPair(root / L"swap-b.txt", root / L"swap-a.txt"),
                  L"Swap success callback should pair the second original with its swapped target.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesChainRename(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_chain_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename chain root.");
    state.Require(SelfTest::WriteTextFile(root / L"chain-a.txt", "first-content"), L"Failed to create first Batch Rename chain input.");
    state.Require(SelfTest::WriteTextFile(root / L"chain-b.txt", "second-content"), L"Failed to create second Batch Rename chain input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> countingFileSystem = CreateBatchRenameRenameCountingFileSystem(fileSystem, &renameItemsCalls);
    state.Require(countingFileSystem != nullptr, L"Batch Rename chain selftest should create a rename-counting file-system wrapper.");
    if (! countingFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = countingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-chain-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"chain-a.txt", root / L"chain-b.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-chain-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for chain execution testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename chain test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual), L"Batch Rename chain test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"chain-b.txt\nchain-c.txt"),
                  L"Batch Rename chain test should set chained manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename chain preview should be available before executing.");
    state.Require(before.errorRowCount == 0u, L"Chain preview must not flag the vacated destination as a collision.");
    state.Require(before.renameButtonEnabled, L"Valid chain preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr), std::format(L"Batch Rename chain execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(! SelfTest::PathExists(root / L"chain-a.txt"), L"Chain execution should vacate the chain head name.");
    state.Require(ReadBatchRenameFileText(root / L"chain-b.txt") == "first-content",
                  L"Chain execution should move the first file's content to the vacated middle name.");
    state.Require(ReadBatchRenameFileText(root / L"chain-c.txt") == "second-content",
                  L"Chain execution should move the second file's content to the tail name.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Chain execution must not leave .rsren- temp files behind.");
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 2u,
                  std::format(L"Chain execution should run in two dependency layers without temp hops; saw {} RenameItems calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename chain snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 2u, L"Chain execution should report two completed rows.");
    state.Require(after.lastExecutionFailedRows == 0u, L"Chain execution should report zero failed rows.");
    state.Require(after.lastExecutionUndoRowCount == 2u, L"Chain execution should record one undo entry per row.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Batch Rename chain test should copy the retained undo plan.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains((root / L"chain-b.txt").native() + L"\tchain-a.txt\t" + (root / L"chain-a.txt").native()),
                  std::format(L"Chain undo plan should record the net head transition; saw '{}'.", undoReport));
    state.Require(undoReport.contains((root / L"chain-c.txt").native() + L"\tchain-b.txt\t" + (root / L"chain-b.txt").native()),
                  std::format(L"Chain undo plan should record the net tail transition; saw '{}'.", undoReport));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowDirectoryChainUndoPlan(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_dir_chain_undo_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root / L"chain-a"), L"Failed to create first Batch Rename directory-chain input.");
    state.Require(SelfTest::EnsureDirectory(root / L"chain-b"), L"Failed to create second Batch Rename directory-chain input.");
    state.Require(SelfTest::WriteTextFile(root / L"chain-a" / L"marker-a.txt", "marker-a"),
                  L"Failed to create first Batch Rename directory-chain marker.");
    state.Require(SelfTest::WriteTextFile(root / L"chain-b" / L"marker-b.txt", "marker-b"),
                  L"Failed to create second Batch Rename directory-chain marker.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-dir-chain-undo-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"chain-a", root / L"chain-b"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-dir-chain-undo-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for directory-chain undo testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename directory-chain undo test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual), L"Batch Rename directory-chain undo test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"chain-b\nchain-c"),
                  L"Batch Rename directory-chain undo test should set chained manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename directory-chain preview should be available before executing.");
    state.Require(before.errorRowCount == 0u, L"Directory-chain preview must not flag the vacated destination as a collision.");
    state.Require(before.renameButtonEnabled, L"Valid directory-chain preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename directory-chain execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(! SelfTest::PathExists(root / L"chain-a"), L"Directory-chain execution should vacate the chain head name.");
    state.Require(SelfTest::PathExists(root / L"chain-b" / L"marker-a.txt"),
                  L"Directory-chain execution should move the first directory to the vacated middle name.");
    state.Require(SelfTest::PathExists(root / L"chain-c" / L"marker-b.txt"),
                  L"Directory-chain execution should move the second directory to the tail name.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename directory-chain snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 2u, L"Directory-chain execution should report two completed rows.");
    state.Require(after.lastExecutionFailedRows == 0u, L"Directory-chain execution should report zero failed rows.");
    state.Require(after.lastExecutionUndoRowCount == 2u, L"Directory-chain execution should record one undo entry per row.");

    // Regression pin: the undo entry for the head rename (chain-a -> chain-b) must
    // keep currentPath = chain-b. The sibling move chain-b -> chain-c relocated the
    // ORIGINAL chain-b directory, not the new occupant of that name, so finalizing
    // undo paths must not rewrite the head entry's chain-b into chain-c.
    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Batch Rename directory-chain undo test should copy the retained undo plan.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains((root / L"chain-b").native() + L"\tchain-a\t" + (root / L"chain-a").native()),
                  std::format(L"Directory-chain undo plan should keep the head entry at the vacated middle name; saw '{}'.", undoReport));
    state.Require(undoReport.contains((root / L"chain-c").native() + L"\tchain-b\t" + (root / L"chain-b").native()),
                  std::format(L"Directory-chain undo plan should record the net tail transition; saw '{}'.", undoReport));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameCaseOnlyRenameNotTreatedAsDependency(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_case_dep_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename case-dependency root.");
    state.Require(SelfTest::WriteTextFile(root / L"CaseSwap.TXT", "case-content"), L"Failed to create Batch Rename case-only input.");
    state.Require(SelfTest::WriteTextFile(root / L"plain.txt", "plain-content"), L"Failed to create Batch Rename plain input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> countingFileSystem = CreateBatchRenameRenameCountingFileSystem(fileSystem, &renameItemsCalls);
    state.Require(countingFileSystem != nullptr, L"Batch Rename case-dependency selftest should create a rename-counting file-system wrapper.");
    if (! countingFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = countingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-case-dep-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"CaseSwap.TXT", root / L"plain.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-case-dep-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for case-only dependency testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename case-only dependency test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename case-only dependency test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"caseswap.txt\nplain-renamed.txt"),
                  L"Batch Rename case-only dependency test should set a case-only name beside a normal rename.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename case-only dependency preview should be available.");
    state.Require(before.changedRowCount == 2u, L"Case-only dependency preview should report both rows changed.");
    state.Require(before.errorRowCount == 0u, L"Case-only self rename must not be flagged as a destination conflict.");
    state.Require(before.renameButtonEnabled, L"Valid case-only dependency preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename case-only dependency execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  std::format(L"Case-only self rename must execute in a single dependency layer (no temp-hop cycle); saw {} RenameItems calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Case-only dependency execution must not leave .rsren- temp files behind.");
    state.Require(ListBatchRenameDirectoryLeaves(root) == std::vector<std::wstring>{L"caseswap.txt", L"plain-renamed.txt"},
                  L"Case-only dependency execution should apply the exact requested casing and the sibling rename.");
    state.Require(ReadBatchRenameFileText(root / L"caseswap.txt") == "case-content",
                  L"Case-only rename should preserve the file content.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename case-only dependency snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 2u, L"Case-only dependency execution should report two completed rows.");
    state.Require(after.lastExecutionFailedRows == 0u, L"Case-only dependency execution should report zero failed rows.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenamePartialBatchFailureTracksCompletedRows(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_partial_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename partial-failure root.");
    state.Require(SelfTest::WriteTextFile(root / L"part-a.txt", "a"), L"Failed to create first Batch Rename partial-failure input.");
    state.Require(SelfTest::WriteTextFile(root / L"part-b.txt", "b"), L"Failed to create second Batch Rename partial-failure input.");
    state.Require(SelfTest::WriteTextFile(root / L"part-c.txt", "c"), L"Failed to create third Batch Rename partial-failure input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    const HRESULT failHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> partialFileSystem =
        CreateBatchRenameScriptedPerItemFileSystem(fileSystem, &renameItemsCalls, L"part-b.txt", failHr, false);
    state.Require(partialFileSystem != nullptr, L"Batch Rename partial-failure selftest should create a scripted per-item file-system wrapper.");
    if (! partialFileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = partialFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-partial-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"part-a.txt", root / L"part-b.txt", root / L"part-c.txt"};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-partial-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for partial-failure testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename partial-failure test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"partial-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename partial-failure test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename partial-failure preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid partial-failure preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == failHr,
                  std::format(L"Partial batch failure should surface the per-item failure; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));

    state.Require(SelfTest::PathExists(root / L"partial-001.txt"), L"Partial failure should keep the first successful rename.");
    state.Require(SelfTest::PathExists(root / L"partial-003.txt"), L"Partial failure should keep the third successful rename.");
    state.Require(SelfTest::PathExists(root / L"part-b.txt"), L"Partial failure should preserve the failed row's source.");
    state.Require(! SelfTest::PathExists(root / L"partial-002.txt"), L"Partial failure should not create the failed row's destination.");
    state.Require(! SelfTest::PathExists(root / L"part-a.txt"), L"Partial failure should remove the first renamed source.");
    state.Require(! SelfTest::PathExists(root / L"part-c.txt"), L"Partial failure should remove the third renamed source.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename partial-failure snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Partial failure should retain an execution report.");
    state.Require(after.lastExecutionCompletedRows == 2u,
                  std::format(L"Partial failure should report the two rows that actually renamed; saw {}.", after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionFailedRows == 1u,
                  std::format(L"Partial failure should count only the real failure; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionUndoRowCount == 2u,
                  std::format(L"Partial failure should record undo entries for the completed rows only; saw {}.", after.lastExecutionUndoRowCount));
    state.Require(after.lastExecutionFirstFailure == failHr,
                  std::format(L"Partial failure report should carry the per-item failure HRESULT; saw 0x{:08X}.",
                              static_cast<unsigned long>(after.lastExecutionFirstFailure)));
    state.Require(! after.lastExecutionCanceled, L"Partial failure must not be reported as canceled.");

    state.Require(callbackCalls == 1u,
                  std::format(L"Partial failure should still invoke the success callback for completed rows; saw {} calls.", callbackCalls));
    state.Require(callbackSources.size() == 2u && callbackTargets.size() == 2u,
                  L"Partial-failure success callback should report exactly the two completed rows.");
    const auto hasPair = [&](const std::filesystem::path& source, const std::filesystem::path& target) noexcept
    {
        for (size_t index = 0u; index < callbackSources.size() && index < callbackTargets.size(); ++index)
        {
            if (callbackSources[index] == source && callbackTargets[index] == target)
            {
                return true;
            }
        }
        return false;
    };
    state.Require(hasPair(root / L"part-a.txt", root / L"partial-001.txt"),
                  L"Partial-failure success callback should pair the first completed row.");
    state.Require(hasPair(root / L"part-c.txt", root / L"partial-003.txt"),
                  L"Partial-failure success callback should pair the third completed row.");
    state.Require(std::ranges::find(callbackSources, root / L"part-b.txt") == callbackSources.end(),
                  L"Partial-failure success callback must not report the failed row.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Partial failure should retain a copyable undo plan for the completed rows.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains((root / L"partial-001.txt").native()) && undoReport.contains((root / L"partial-003.txt").native()),
                  std::format(L"Partial-failure undo plan should include both completed rows; saw '{}'.", undoReport));
    state.Require(! undoReport.contains(L"partial-002.txt"), L"Partial-failure undo plan must not include the failed row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameOmittedItemCompletionCountsRowFailed(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_omitted_completion_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename omitted-completion root.");
    state.Require(SelfTest::WriteTextFile(root / L"omit-a.txt", "a"), L"Failed to create first Batch Rename omitted-completion input.");
    state.Require(SelfTest::WriteTextFile(root / L"omit-b.txt", "b"), L"Failed to create second Batch Rename omitted-completion input.");
    state.Require(SelfTest::WriteTextFile(root / L"omit-c.txt", "c"), L"Failed to create third Batch Rename omitted-completion input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> omittedFileSystem =
        CreateBatchRenameOmittedCompletionFileSystem(fileSystem, &renameItemsCalls, L"omit-b.txt");
    state.Require(omittedFileSystem != nullptr, L"Batch Rename omitted-completion selftest should create a scripted file-system wrapper.");
    if (! omittedFileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = omittedFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-omitted-completion-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"omit-a.txt", root / L"omit-b.txt", root / L"omit-c.txt"};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-omitted-completion-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for omitted-completion testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename omitted-completion test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"omitted-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename omitted-completion test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename omitted-completion preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid omitted-completion preview should enable execution.");

    const HRESULT expectedFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    const HRESULT executeHr       = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == expectedFailure,
                  std::format(L"Omitted item completion should surface ERROR_INVALID_DATA; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  std::format(L"Omitted-completion test should use one bulk RenameItems call; saw {}.",
                              renameItemsCalls.load(std::memory_order_relaxed)));

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename omitted-completion snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Omitted item completion should retain an execution report.");
    state.Require(after.lastExecutionCompletedRows == 2u,
                  std::format(L"Omitted item completion should report only the two callback-confirmed rows as completed; saw {}.",
                              after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionFailedRows == 1u,
                  std::format(L"Omitted item completion should count the unreported row as failed; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionUndoRowCount == 2u,
                  std::format(L"Omitted item completion should record undo only for callback-confirmed rows; saw {}.",
                              after.lastExecutionUndoRowCount));
    state.Require(after.lastExecutionFirstFailure == expectedFailure,
                  std::format(L"Omitted item completion should retain ERROR_INVALID_DATA; saw 0x{:08X}.",
                              static_cast<unsigned long>(after.lastExecutionFirstFailure)));

    state.Require(callbackCalls == 1u,
                  std::format(L"Omitted item completion should still invoke the success callback for confirmed rows; saw {} calls.",
                              callbackCalls));
    state.Require(callbackSources.size() == 2u && callbackTargets.size() == 2u,
                  L"Omitted item completion should report exactly the two callback-confirmed rows.");
    state.Require(std::ranges::find(callbackSources, root / L"omit-b.txt") == callbackSources.end(),
                  L"Omitted item completion must not report the omitted row as successful.");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"Omitted item completion should retain a copyable undo plan.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains((root / L"omitted-001.txt").native()) && undoReport.contains((root / L"omitted-003.txt").native()),
                  std::format(L"Omitted item completion undo plan should include only confirmed rows; saw '{}'.", undoReport));
    state.Require(! undoReport.contains(L"omitted-002.txt"), L"Omitted item completion undo plan must not include the unreported row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameTempHopReportedFailureCountsRowFailed(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_temp_reported_failure_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path longSource  = root / L"swap-aaaa.txt";
    const std::filesystem::path shortSource = root / L"b.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename temp-hop reported-failure root.");
    state.Require(SelfTest::WriteTextFile(longSource, "long"), L"Failed to create long Batch Rename swap input.");
    state.Require(SelfTest::WriteTextFile(shortSource, "short"), L"Failed to create short Batch Rename swap input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    const HRESULT reportedFailure = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> reportedFailureFileSystem =
        CreateBatchRenameReportedFailureOverallSuccessFileSystem(fileSystem, &renameItemsCalls, L"swap-aaaa.txt", reportedFailure);
    state.Require(reportedFailureFileSystem != nullptr,
                  L"Batch Rename temp-hop reported-failure selftest should create a scripted file-system wrapper.");
    if (! reportedFailureFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = reportedFailureFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-temp-reported-failure-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {longSource, shortSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-temp-reported-failure-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for temp-hop reported-failure testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename temp-hop reported-failure window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename temp-hop reported-failure test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"b.txt\nswap-aaaa.txt"),
                  L"Batch Rename temp-hop reported-failure test should set crossed manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename temp-hop reported-failure preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid swap preview should enable execution before the provider reports temp-hop failure.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == reportedFailure,
                  std::format(L"Temp-hop reported failure should surface the item failure despite overall S_OK; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) >= 1u,
                  L"Temp-hop reported failure should attempt at least the temp-hop RenameItems call.");

    state.Require(SelfTest::PathExists(longSource), L"Temp-hop reported failure should leave the original long source in place.");
    state.Require(SelfTest::PathExists(shortSource), L"Temp-hop reported failure should leave the short source in place after the blocked swap.");
    state.Require(ReadBatchRenameFileText(longSource) == "long", L"Temp-hop reported failure should preserve long source content.");
    state.Require(ReadBatchRenameFileText(shortSource) == "short", L"Temp-hop reported failure should preserve short source content.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Temp-hop reported failure must not leave .rsren- temp files behind.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename temp-hop reported-failure snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Temp-hop reported failure should retain an execution report.");
    state.Require(after.lastExecutionCompletedRows == 0u,
                  std::format(L"Temp-hop reported failure should not count any row completed; saw {}.", after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionFailedRows == 2u,
                  std::format(L"Temp-hop reported failure should count both swap rows failed; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionUndoRowCount == 0u,
                  std::format(L"Temp-hop reported failure should not record phantom undo entries; saw {}.", after.lastExecutionUndoRowCount));
    state.Require(after.lastExecutionFirstFailure == reportedFailure,
                  std::format(L"Temp-hop reported failure should retain the reported item HRESULT; saw 0x{:08X}.",
                              static_cast<unsigned long>(after.lastExecutionFirstFailure)));

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameSwapFailureRollsBackOrphanTemp(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_swap_rollback_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path longSource  = root / L"swap-aaaa.txt";
    const std::filesystem::path shortSource = root / L"b.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename swap-rollback root.");
    state.Require(SelfTest::WriteTextFile(longSource, "long-content"), L"Failed to create long Batch Rename swap-rollback input.");
    state.Require(SelfTest::WriteTextFile(shortSource, "short-content"), L"Failed to create short Batch Rename swap-rollback input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    const HRESULT failHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> partialFileSystem =
        CreateBatchRenameScriptedPerItemFileSystem(fileSystem, &renameItemsCalls, L"b.txt", failHr, false);
    state.Require(partialFileSystem != nullptr, L"Batch Rename swap-rollback selftest should create a scripted file-system wrapper.");
    if (! partialFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = partialFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-swap-rollback-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {longSource, shortSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-swap-rollback-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for swap rollback testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename swap rollback window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual), L"Batch Rename swap-rollback test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"b.txt\nswap-aaaa.txt"),
                  L"Batch Rename swap-rollback test should set crossed manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename swap-rollback preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid swap preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == failHr,
                  std::format(L"Swap rollback should surface the failing non-temp row; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 2u,
                  std::format(L"Swap rollback should execute temp hop plus failing layer; saw {} RenameItems calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));

    state.Require(SelfTest::PathExists(longSource), L"Swap rollback should restore the temp-hop source leaf.");
    state.Require(SelfTest::PathExists(shortSource), L"Swap rollback should preserve the failed non-temp source leaf.");
    state.Require(ReadBatchRenameFileText(longSource) == "long-content", L"Swap rollback should restore long source content.");
    state.Require(ReadBatchRenameFileText(shortSource) == "short-content", L"Swap rollback should preserve short source content.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Swap rollback must not leave .rsren- temp files behind.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename swap-rollback snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 0u, L"Swap rollback should report zero completed rows.");
    state.Require(after.lastExecutionFailedRows == 2u, L"Swap rollback should report both swap rows failed.");
    state.Require(after.lastExecutionUndoRowCount == 0u, L"Swap rollback should not record phantom undo entries.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameSwapCancelAfterTempHopRollsBack(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_swap_cancel_rollback_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    const std::filesystem::path longSource  = root / L"swap-aaaa.txt";
    const std::filesystem::path shortSource = root / L"b.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename swap cancel root.");
    state.Require(SelfTest::WriteTextFile(longSource, "long-content"), L"Failed to create long Batch Rename swap cancel input.");
    state.Require(SelfTest::WriteTextFile(shortSource, "short-content"), L"Failed to create short Batch Rename swap cancel input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> cancelFileSystem =
        CreateBatchRenameScriptedPerItemFileSystem(fileSystem, &renameItemsCalls, {}, S_OK, true);
    state.Require(cancelFileSystem != nullptr, L"Batch Rename swap cancel selftest should create a scripted file-system wrapper.");
    if (! cancelFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = cancelFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-swap-cancel-rollback-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {longSource, shortSource};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-swap-cancel-rollback-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for swap cancel rollback testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename swap cancel rollback window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename swap cancel rollback test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"b.txt\nswap-aaaa.txt"),
                  L"Batch Rename swap cancel rollback test should set crossed manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename swap cancel rollback preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid swap preview should enable execution.");

    const HRESULT cancelHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == cancelHr,
                  std::format(L"Swap cancel rollback should surface cancellation after temp hop; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));

    state.Require(SelfTest::PathExists(longSource), L"Swap cancel rollback should restore the temp-hop source leaf.");
    state.Require(SelfTest::PathExists(shortSource), L"Swap cancel rollback should preserve the untouched short source leaf.");
    state.Require(ReadBatchRenameFileText(longSource) == "long-content", L"Swap cancel rollback should restore long source content.");
    state.Require(ReadBatchRenameFileText(shortSource) == "short-content", L"Swap cancel rollback should preserve short source content.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Swap cancel rollback must not leave .rsren- temp files behind.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename swap cancel rollback snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 0u, L"Swap cancel rollback should report zero completed rows.");
    state.Require(after.lastExecutionFailedRows == 2u, L"Swap cancel rollback should report both swap rows failed/never-run.");
    state.Require(after.lastExecutionUndoRowCount == 0u, L"Swap cancel rollback should not record phantom undo entries.");
    state.Require(after.lastExecutionCanceled, L"Swap cancel rollback should retain canceled=true.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameWindowExecutesThreeMemberCycle(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_three_cycle_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename three-cycle root.");
    state.Require(SelfTest::WriteTextFile(root / L"cycle-a.txt", "a"), L"Failed to create first Batch Rename three-cycle input.");
    state.Require(SelfTest::WriteTextFile(root / L"cycle-b.txt", "b"), L"Failed to create second Batch Rename three-cycle input.");
    state.Require(SelfTest::WriteTextFile(root / L"cycle-c.txt", "c"), L"Failed to create third Batch Rename three-cycle input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> countingFileSystem = CreateBatchRenameRenameCountingFileSystem(fileSystem, &renameItemsCalls);
    state.Require(countingFileSystem != nullptr, L"Batch Rename three-cycle selftest should create a rename-counting file-system wrapper.");
    if (! countingFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = countingFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-three-cycle-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"cycle-a.txt", root / L"cycle-b.txt", root / L"cycle-c.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-three-cycle-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for three-cycle testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename three-cycle window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual), L"Batch Rename three-cycle test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"cycle-b.txt\ncycle-c.txt\ncycle-a.txt"),
                  L"Batch Rename three-cycle test should set a three-member cycle.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename three-cycle preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid three-cycle preview should enable execution.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Three-member cycle execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 4u,
                  std::format(L"Three-member cycle should run one temp hop plus three dependency layers; saw {} RenameItems calls.",
                              renameItemsCalls.load(std::memory_order_relaxed)));
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Three-member cycle must not leave .rsren- temp files behind.");
    state.Require(ReadBatchRenameFileText(root / L"cycle-a.txt") == "c", L"Three-member cycle should rotate c content to cycle-a.");
    state.Require(ReadBatchRenameFileText(root / L"cycle-b.txt") == "a", L"Three-member cycle should rotate a content to cycle-b.");
    state.Require(ReadBatchRenameFileText(root / L"cycle-c.txt") == "b", L"Three-member cycle should rotate b content to cycle-c.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename three-cycle snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 3u, L"Three-member cycle should report three completed rows.");
    state.Require(after.lastExecutionFailedRows == 0u, L"Three-member cycle should report zero failed rows.");
    state.Require(after.lastExecutionUndoRowCount == 3u, L"Three-member cycle should record one undo entry per row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameCancelMidBatchTracksCompletedRows(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_midcancel_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    // Execution orders equal-depth rows by descending path length, so the
    // longer leaf below deterministically runs (and completes) first.
    const std::filesystem::path firstSource  = root / L"cancel-first-much-longer-name.txt";
    const std::filesystem::path secondSource = root / L"cz.txt";
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename mid-cancel root.");
    state.Require(SelfTest::WriteTextFile(firstSource, "first"), L"Failed to create first Batch Rename mid-cancel input.");
    state.Require(SelfTest::WriteTextFile(secondSource, "second"), L"Failed to create second Batch Rename mid-cancel input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> cancelFileSystem =
        CreateBatchRenameScriptedPerItemFileSystem(fileSystem, &renameItemsCalls, {}, S_OK, true);
    state.Require(cancelFileSystem != nullptr, L"Batch Rename mid-cancel selftest should create a cancel-after-first file-system wrapper.");
    if (! cancelFileSystem)
    {
        return false;
    }

    uint32_t callbackCalls = 0u;
    std::vector<std::filesystem::path> callbackSources;
    std::vector<std::filesystem::path> callbackTargets;

    BatchRenamePaneContext context{};
    context.fileSystem      = cancelFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-midcancel-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {firstSource, secondSource};
    context.onSuccessfulRename = [&](std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths)
    {
        ++callbackCalls;
        callbackSources.assign(sourcePaths.begin(), sourcePaths.end());
        callbackTargets.assign(targetPaths.begin(), targetPaths.end());
    };

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-midcancel-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for mid-batch cancel testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename mid-batch cancel test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"midcancel-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename mid-batch cancel test should set valid rules.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(executeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Mid-batch cancel execution should return ERROR_CANCELLED; saw 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  L"Mid-batch cancel test should reach the provider exactly once.");

    state.Require(SelfTest::PathExists(root / L"midcancel-001.txt"),
                  L"Mid-batch cancel should keep the row completed before the cancellation.");
    state.Require(! SelfTest::PathExists(firstSource), L"Mid-batch cancel should remove the completed row's source.");
    state.Require(SelfTest::PathExists(secondSource), L"Mid-batch cancel should preserve the remaining row's source.");
    state.Require(! SelfTest::PathExists(root / L"midcancel-002.txt"), L"Mid-batch cancel should not apply the remaining row.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename mid-batch cancel snapshot should be available after executing.");
    state.Require(after.hasExecutionReport, L"Mid-batch cancel should retain an execution report.");
    state.Require(after.lastExecutionCanceled, L"Mid-batch cancel report should be marked canceled.");
    state.Require(after.lastExecutionCompletedRows == 1u,
                  std::format(L"Mid-batch cancel should track the pre-cancel completed row; saw {}.", after.lastExecutionCompletedRows));
    state.Require(after.lastExecutionFailedRows == 1u,
                  std::format(L"Mid-batch cancel should count the unapplied remaining row as failed; saw {}.", after.lastExecutionFailedRows));
    state.Require(after.lastExecutionUndoRowCount == 1u,
                  std::format(L"Mid-batch cancel should record an undo entry for the completed row only; saw {}.", after.lastExecutionUndoRowCount));

    state.Require(callbackCalls == 1u,
                  std::format(L"Mid-batch cancel should invoke the success callback for the completed row; saw {} calls.", callbackCalls));
    state.Require(callbackSources == std::vector<std::filesystem::path>{firstSource} &&
                      callbackTargets == std::vector<std::filesystem::path>{root / L"midcancel-001.txt"},
                  L"Mid-batch cancel success callback should report exactly the completed pre-cancel row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameExecuteWhileBusyReturnsBusy(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_busy_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename busy root.");
    state.Require(SelfTest::WriteTextFile(root / L"busy-a.txt", "a"), L"Failed to create first Batch Rename busy input.");
    state.Require(SelfTest::WriteTextFile(root / L"busy-b.txt", "b"), L"Failed to create second Batch Rename busy input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    wil::unique_handle gate{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    state.Require(static_cast<bool>(gate), L"Batch Rename busy selftest should create a provider gate event.");
    if (! gate)
    {
        return false;
    }

    wil::com_ptr<IFileSystem> gatedFileSystem = CreateBatchRenameGatedRenameItemsFileSystem(fileSystem, gate.get());
    state.Require(gatedFileSystem != nullptr, L"Batch Rename busy selftest should create a gated file-system wrapper.");
    if (! gatedFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = gatedFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-busy-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"busy-a.txt", root / L"busy-b.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-busy-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for busy-state testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename busy test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    // Declared after the window cleanup so the gate opens before the window
    // joins its worker on close, even when an assertion fails mid-test.
    const auto releaseGate    = wil::scope_exit([&gate]() noexcept { static_cast<void>(SetEvent(gate.get())); });
    const auto cleanupFiles   = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"busy-renamed-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename busy test should set valid rules.");

    const HRESULT startHr = DebugStartBatchRenameWindowExecution();
    state.Require(startHr == S_OK,
                  std::format(L"Batch Rename busy test should start the gated execution; saw 0x{:08X}.", static_cast<unsigned long>(startHr)));
    if (FAILED(startHr))
    {
        return false;
    }

    const HRESULT busyHr = DebugStartBatchRenameWindowExecution();
    state.Require(busyHr == HRESULT_FROM_WIN32(ERROR_BUSY),
                  std::format(L"ExecuteRename while an execution is in flight should return ERROR_BUSY; saw 0x{:08X}.",
                              static_cast<unsigned long>(busyHr)));

    state.Require(SetEvent(gate.get()) != FALSE, L"Batch Rename busy test should release the provider gate.");

    const HRESULT terminalHr = DebugWaitBatchRenameWindowExecutionIdle();
    state.Require(SUCCEEDED(terminalHr),
                  std::format(L"Gated Batch Rename execution should finish successfully after release; saw 0x{:08X}.",
                              static_cast<unsigned long>(terminalHr)));

    state.Require(SelfTest::PathExists(root / L"busy-renamed-001.txt"), L"Gated execution should create the first renamed path.");
    state.Require(SelfTest::PathExists(root / L"busy-renamed-002.txt"), L"Gated execution should create the second renamed path.");

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename busy snapshot should be available after executing.");
    state.Require(after.lastExecutionCompletedRows == 2u, L"Gated execution should report two completed rows.");
    state.Require(after.lastExecutionFailedRows == 0u, L"Gated execution should report zero failed rows.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameCloseWhileExecutionInFlightIsBounded(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_close_inflight_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename close-in-flight root.");
    state.Require(SelfTest::WriteTextFile(root / L"close-a.txt", "a"), L"Failed to create first Batch Rename close-in-flight input.");
    state.Require(SelfTest::WriteTextFile(root / L"close-b.txt", "b"), L"Failed to create second Batch Rename close-in-flight input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    wil::unique_handle gate{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    state.Require(static_cast<bool>(gate), L"Batch Rename close-in-flight selftest should create a provider gate event.");
    if (! gate)
    {
        return false;
    }

    wil::com_ptr<IFileSystem> gatedFileSystem = CreateBatchRenameGatedRenameItemsFileSystem(fileSystem, gate.get());
    state.Require(gatedFileSystem != nullptr, L"Batch Rename close-in-flight selftest should create a gated file-system wrapper.");
    if (! gatedFileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = gatedFileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-close-inflight-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"close-a.txt", root / L"close-b.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-close-inflight-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for close-in-flight testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename close-in-flight test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });
    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto releaseGate        = wil::scope_exit([&gate]() noexcept { static_cast<void>(SetEvent(gate.get())); });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"close-renamed-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename close-in-flight test should set valid rules.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename close-in-flight preview should be available.");
    state.Require(before.renameButtonEnabled, L"Valid close-in-flight preview should enable execution.");

    const HRESULT startHr = DebugStartBatchRenameWindowExecution();
    state.Require(startHr == S_OK,
                  std::format(L"Batch Rename close-in-flight test should start the gated execution; saw 0x{:08X}.",
                              static_cast<unsigned long>(startHr)));
    if (FAILED(startHr))
    {
        return false;
    }

    const HANDLE gateHandle = gate.get();
    std::jthread gateReleaser([gateHandle](std::stop_token) noexcept
    {
        std::this_thread::sleep_for(SelfTest::Scale(std::chrono::milliseconds{150}));
        static_cast<void>(SetEvent(gateHandle));
    });

    state.Require(PostMessageW(batchWindow, WM_CLOSE, 0, 0) != FALSE,
                  L"Batch Rename close-in-flight test should post WM_CLOSE while execution is blocked in RenameItems.");
    const bool closed = WaitForWindowClosed(batchWindow, SelfTest::Scale(std::chrono::milliseconds{5000}));
    state.Require(closed, L"Batch Rename window should close within the bounded timeout while execution is in flight.");
    state.Require(DebugGetBatchRenameWindowCount() == 0u, L"Batch Rename close-in-flight teardown should leave no live window.");

    const bool originalsOnly = SelfTest::PathExists(root / L"close-a.txt") && SelfTest::PathExists(root / L"close-b.txt") &&
        ! SelfTest::PathExists(root / L"close-renamed-001.txt") && ! SelfTest::PathExists(root / L"close-renamed-002.txt");
    const bool renamedOnly = ! SelfTest::PathExists(root / L"close-a.txt") && ! SelfTest::PathExists(root / L"close-b.txt") &&
        SelfTest::PathExists(root / L"close-renamed-001.txt") && SelfTest::PathExists(root / L"close-renamed-002.txt");
    state.Require(originalsOnly || renamedOnly,
                  L"Closing during in-flight execution should leave either both originals or both destinations, never a half state.");
    state.Require(! DirectoryHasBatchRenameTempLeftovers(root), L"Closing during in-flight execution must not leave .rsren- temp files behind.");

    if (renamedOnly)
    {
        state.Require(ReadBatchRenameFileText(root / L"close-renamed-001.txt") == "a",
                      L"Close-in-flight execution should preserve first file content when rename completes.");
        state.Require(ReadBatchRenameFileText(root / L"close-renamed-002.txt") == "b",
                      L"Close-in-flight execution should preserve second file content when rename completes.");
    }
    else if (originalsOnly)
    {
        state.Require(ReadBatchRenameFileText(root / L"close-a.txt") == "a",
                      L"Close-in-flight cancellation should preserve first source content.");
        state.Require(ReadBatchRenameFileText(root / L"close-b.txt") == "b",
                      L"Close-in-flight cancellation should preserve second source content.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameReportSurvivesViewOnlyRefreshes(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_report_keep_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename report-retention root.");
    state.Require(SelfTest::WriteTextFile(root / L"keep-a.txt", "a"), L"Failed to create first Batch Rename report-retention input.");
    state.Require(SelfTest::WriteTextFile(root / L"keep-b.txt", "b"), L"Failed to create second Batch Rename report-retention input.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> fileSystem = GetBatchRenameLocalFileSystem(state);
    if (! fileSystem)
    {
        return false;
    }

    BatchRenamePaneContext context{};
    context.fileSystem      = fileSystem;
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-report-keep-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"keep-a.txt", root / L"keep-b.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-report-keep-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for report-retention testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename report-retention test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });
    const auto cleanupFiles       = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    BatchRename::Rules rules{};
    rules.nameTemplate = L"kept-{counter:000}{ext}";
    state.Require(DebugSetBatchRenameWindowRules(rules), L"Batch Rename report-retention test should set valid rules.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename report-retention execution should succeed: 0x{:08X}.", static_cast<unsigned long>(executeHr)));
    if (FAILED(executeHr))
    {
        return false;
    }

    BatchRenameDebugSnapshot afterExecute{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterExecute), L"Batch Rename report-retention snapshot should be available after executing.");
    state.Require(afterExecute.hasExecutionReport, L"Execution should retain a report before any view-only refresh.");
    state.Require(afterExecute.lastExecutionCompletedRows == 2u && afterExecute.lastExecutionUndoRowCount == 2u,
                  L"Execution report should record both completed rows and undo entries.");

    const auto requireSamePreviewStatsAndGate = [&](const BatchRenameDebugSnapshot& snapshot,
                                                    const BatchRenameDebugSnapshot& baseline,
                                                    const wchar_t* label)
    {
        state.Require(snapshot.changedRowCount == baseline.changedRowCount && snapshot.errorRowCount == baseline.errorRowCount &&
                          snapshot.warningRowCount == baseline.warningRowCount &&
                          snapshot.renameButtonEnabled == baseline.renameButtonEnabled,
                      std::format(L"{} should preserve full-plan stats and Rename gating; before changed/errors/warnings/enabled={}/{}/{}/{}, "
                                  L"after={}/{}/{}/{}.",
                                  label,
                                  baseline.changedRowCount,
                                  baseline.errorRowCount,
                                  baseline.warningRowCount,
                                  baseline.renameButtonEnabled,
                                  snapshot.changedRowCount,
                                  snapshot.errorRowCount,
                                  snapshot.warningRowCount,
                                  snapshot.renameButtonEnabled));
    };

    const uint64_t sortBuildPlanRowsBefore      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t sortRecomputeRowsBefore      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t sortVisibleRefreshRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");
    state.Require(DebugSetBatchRenameWindowPreviewSort(L"original", true),
                  L"Batch Rename report-retention test should sort the preview by original name.");
    const uint64_t sortBuildPlanRowsAfter      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t sortRecomputeRowsAfter      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t sortVisibleRefreshRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");

    BatchRenameDebugSnapshot afterSort{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterSort), L"Batch Rename post-sort snapshot should be available.");
    state.Require(afterSort.hasExecutionReport, L"A column sort is view-only and must preserve the execution report.");
    state.Require(afterSort.lastExecutionUndoRowCount == 2u, L"A column sort must preserve the undo plan.");
    state.Require(afterSort.originalNames == std::vector<std::wstring>{L"kept-002.txt", L"kept-001.txt"},
                  L"A column sort should still reorder the visible preview rows.");
    requireSamePreviewStatsAndGate(afterSort, afterExecute, L"A column sort");
    state.Require(sortBuildPlanRowsAfter == sortBuildPlanRowsBefore,
                  std::format(L"A column sort must not emit batchrename.preview.build_plan_us rows; before={} after={}.",
                              sortBuildPlanRowsBefore,
                              sortBuildPlanRowsAfter));
    state.Require(sortRecomputeRowsAfter == sortRecomputeRowsBefore,
                  std::format(L"A column sort must not emit batchrename.preview.recompute.us rows; before={} after={}.",
                              sortRecomputeRowsBefore,
                              sortRecomputeRowsAfter));
    state.Require(sortVisibleRefreshRowsAfter > sortVisibleRefreshRowsBefore,
                  std::format(L"A column sort should emit batchrename.preview.visible_refresh.us rows; before={} after={}.",
                              sortVisibleRefreshRowsBefore,
                              sortVisibleRefreshRowsAfter));

    const uint64_t hideBuildPlanRowsBefore      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t hideRecomputeRowsBefore      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t hideVisibleRefreshRowsBefore = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");
    state.Require(DebugSetBatchRenameWindowHideUnchanged(true),
                  L"Batch Rename report-retention test should enable Hide unchanged.");
    const uint64_t hideBuildPlanRowsAfter      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.build_plan_us");
    const uint64_t hideRecomputeRowsAfter      = CountBatchRenamePerfRowsWithMetric("batchrename.preview.recompute.us");
    const uint64_t hideVisibleRefreshRowsAfter = CountBatchRenamePerfRowsWithMetric("batchrename.preview.visible_refresh.us");

    BatchRenameDebugSnapshot afterHide{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterHide), L"Batch Rename post-hide snapshot should be available.");
    state.Require(afterHide.hasExecutionReport, L"Toggling Hide unchanged is view-only and must preserve the execution report.");
    state.Require(afterHide.lastExecutionUndoRowCount == 2u, L"Toggling Hide unchanged must preserve the undo plan.");
    state.Require(afterHide.previewRowCount == 0u, L"Hide unchanged should hide every post-execution no-op row.");
    requireSamePreviewStatsAndGate(afterHide, afterExecute, L"Toggling Hide unchanged");
    state.Require(hideBuildPlanRowsAfter == hideBuildPlanRowsBefore,
                  std::format(L"Toggling Hide unchanged must not emit batchrename.preview.build_plan_us rows; before={} after={}.",
                              hideBuildPlanRowsBefore,
                              hideBuildPlanRowsAfter));
    state.Require(hideRecomputeRowsAfter == hideRecomputeRowsBefore,
                  std::format(L"Toggling Hide unchanged must not emit batchrename.preview.recompute.us rows; before={} after={}.",
                              hideRecomputeRowsBefore,
                              hideRecomputeRowsAfter));
    state.Require(hideVisibleRefreshRowsAfter > hideVisibleRefreshRowsBefore,
                  std::format(L"Toggling Hide unchanged should emit batchrename.preview.visible_refresh.us rows; before={} after={}.",
                              hideVisibleRefreshRowsBefore,
                              hideVisibleRefreshRowsAfter));

    state.Require(DebugSetBatchRenameWindowHideUnchanged(false),
                  L"Batch Rename report-retention test should restore unchanged rows.");

    BatchRenameDebugSnapshot afterRestore{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterRestore), L"Batch Rename post-restore snapshot should be available.");
    state.Require(afterRestore.hasExecutionReport, L"Restoring unchanged rows must preserve the execution report.");
    state.Require(afterRestore.previewRowCount == 2u, L"Restoring unchanged rows should show the full post-execution preview.");
    requireSamePreviewStatsAndGate(afterRestore, afterExecute, L"Restoring unchanged rows");

    ClearClipboardContents(batchWindow);
    state.Require(DebugCopyBatchRenameWindowUndoPlan(), L"The undo plan should remain copyable after view-only refreshes.");
    const std::wstring undoReport = ReadClipboardUnicodeText(batchWindow);
    state.Require(undoReport.contains(L"Current Path\tRestore Name\tOriginal Path"),
                  std::format(L"Retained undo plan should keep TSV headers; saw '{}'.", undoReport));

    BatchRename::Rules changedRules{};
    changedRules.nameTemplate = L"{stem}_changed{ext}";
    state.Require(DebugSetBatchRenameWindowRules(changedRules),
                  L"Batch Rename report-retention test should apply an actual rule change.");

    BatchRenameDebugSnapshot afterRuleChange{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterRuleChange), L"Batch Rename post-rule-change snapshot should be available.");
    state.Require(! afterRuleChange.hasExecutionReport, L"An actual rule change invalidates the plan and must clear the execution report.");
    state.Require(! DebugCopyBatchRenameWindowUndoPlan(), L"The undo plan must not be copyable after the report is cleared.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenamePostExecutionTargetRefreshIsLinear(HWND mainWindow, CaseState& state) noexcept
{
    static_cast<void>(mainWindow);

    constexpr size_t kTargetCount = 10'000u;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"batch_rename_linear_target_refresh_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Batch Rename linear target-refresh root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    std::vector<BatchRename::Target> targets;
    std::vector<std::filesystem::path> successfulSources;
    std::vector<std::filesystem::path> successfulTargets;
    targets.reserve(kTargetCount);
    successfulSources.reserve(kTargetCount);
    successfulTargets.reserve(kTargetCount);
    for (size_t index = 0u; index < kTargetCount; ++index)
    {
        const std::filesystem::path source = root / std::format(L"linear-{:05}.txt", index);
        const std::filesystem::path target = root / std::format(L"linear-{:05}_renamed.txt", index);
        BatchRename::Target entry{};
        entry.sourcePath = source;
        targets.push_back(std::move(entry));
        successfulSources.push_back(source);
        successfulTargets.push_back(target);
    }

    const uint64_t targetRefreshMetricRowsBefore =
        CountBatchRenamePerfRowsWithMetric("batchrename.execute.target_refresh_match.us");
    size_t refreshedRows = 0u;
    uint64_t identityComparisons = 0u;
    state.Require(DebugRefreshBatchRenameTargetsAfterExecutionForTests(FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
                                                                       targets,
                                                                       successfulSources,
                                                                       successfulTargets,
                                                                       root,
                                                                       refreshedRows,
                                                                       identityComparisons),
                  L"Linear target-refresh helper should refresh every successful source.");
    state.Require(refreshedRows == kTargetCount,
                  std::format(L"Linear target refresh should refresh {} rows; saw {}.", kTargetCount, refreshedRows));
    state.Require(identityComparisons == kTargetCount,
                  std::format(L"Linear target refresh should verify exactly one identity candidate per row; saw {} for {} rows.",
                              identityComparisons,
                              kTargetCount));

    const uint64_t targetRefreshMetricRowsAfter =
        CountBatchRenamePerfRowsWithMetric("batchrename.execute.target_refresh_match.us");
    state.Require(targetRefreshMetricRowsAfter > targetRefreshMetricRowsBefore,
                  std::format(L"Post-execution target refresh should emit batchrename.execute.target_refresh_match.us; before={} after={}.",
                              targetRefreshMetricRowsBefore,
                              targetRefreshMetricRowsAfter));
    const std::optional<uint64_t> maxIdentityComparisons = TryReadMaxBatchRenamePerfUintField(
        "batchrename.execute.target_refresh_match.us", "value1", targetRefreshMetricRowsBefore);
    state.Require(maxIdentityComparisons.has_value() && maxIdentityComparisons.value() <= static_cast<uint64_t>(kTargetCount),
                  std::format(L"Post-execution target refresh should verify at most one target candidate per success; max comparisons={} targetCount={}.",
                              maxIdentityComparisons.value_or(0u),
                              kTargetCount));

    if (targets.size() == kTargetCount)
    {
        state.Require(targets.front().sourcePath.filename().native() == L"linear-00000_renamed.txt",
                      std::format(L"First refreshed target should use its renamed path; saw '{}'.",
                                  targets.front().sourcePath.filename().native()));
        state.Require(targets.back().sourcePath.filename().native() == L"linear-09999_renamed.txt",
                      std::format(L"Last refreshed target should use its renamed path; saw '{}'.",
                                  targets.back().sourcePath.filename().native()));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameDuplicateSourceTargetRefreshConsumesDistinctRows(HWND mainWindow, CaseState& state) noexcept
{
    const wil::com_ptr<IFileSystem> baseFileSystem = GetBatchRenameLocalFileSystem(state);
    if (! baseFileSystem)
    {
        return false;
    }

    std::atomic_uint32_t renameItemsCalls{0u};
    wil::com_ptr<IFileSystem> noOpFileSystem = CreateBatchRenameNoOpRenameItemsFileSystem(
        baseFileSystem, &renameItemsCalls, std::string(kBatchRenameNoOpProviderCapabilities));
    state.Require(noOpFileSystem != nullptr, L"Batch Rename duplicate-source selftest should create a no-op provider wrapper.");
    if (! noOpFileSystem)
    {
        return false;
    }

    const std::filesystem::path root   = L"C:\\BatchRenameDuplicateSourceTargetRefreshSelfTest";
    const std::filesystem::path source = root / L"duplicate-source.txt";

    BatchRenamePaneContext context{};
    context.fileSystem      = noOpFileSystem;
    context.pluginId        = L"selftest/noop-rename";
    context.pluginShortId   = L"noop";
    context.instanceContext = L"batch-rename-duplicate-source-target-refresh-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {source, source};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-duplicate-source-target-refresh-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for duplicate-source target-refresh testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE,
                  L"Batch Rename duplicate-source target-refresh test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename duplicate-source test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"duplicate-one.txt\nduplicate-two.txt"),
                  L"Batch Rename duplicate-source test should set distinct manual names.");

    BatchRenameDebugSnapshot before{};
    state.Require(DebugGetBatchRenameWindowSnapshot(before), L"Batch Rename duplicate-source preview should be available.");
    state.Require(before.previewRowCount == 2u, L"Duplicate-source preview should keep both selected rows.");
    state.Require(before.changedRowCount == 2u, L"Duplicate-source preview should mark both manual rows changed.");
    state.Require(before.errorRowCount == 0u, L"Duplicate-source preview should allow distinct target names.");
    state.Require(before.renameButtonEnabled, L"Duplicate-source preview should be executable.");

    const HRESULT executeHr = DebugExecuteBatchRenameWindow();
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"Batch Rename duplicate-source target-refresh execution should succeed: 0x{:08X}.",
                              static_cast<unsigned long>(executeHr)));
    state.Require(renameItemsCalls.load(std::memory_order_relaxed) == 1u,
                  std::format(L"Duplicate-source target-refresh test should dispatch one bulk RenameItems call; saw {}.",
                              renameItemsCalls.load(std::memory_order_relaxed)));

    BatchRenameDebugSnapshot after{};
    state.Require(DebugGetBatchRenameWindowSnapshot(after), L"Batch Rename duplicate-source snapshot should be available after executing.");
    state.Require(after.originalNames == std::vector<std::wstring>{L"duplicate-one.txt", L"duplicate-two.txt"},
                  L"Duplicate-source refresh should consume two distinct target indexes instead of refreshing the first row twice.");
    state.Require(after.changedRowCount == 0u, L"Duplicate-source refresh should leave no changed rows after manual targets match.");
    state.Require(! after.renameButtonEnabled, L"Duplicate-source refresh should disable Rename after both rows are refreshed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenameManualSortLikePreviewWithHideUnchanged(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenameSortHideSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-sort-hide-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"sortp-a.txt", root / L"sortp-b.txt", root / L"sortp-c.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-sort-hide-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for sort-like-preview with Hide unchanged testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE,
                  L"Batch Rename sort-like-preview test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    state.Require(DebugSwitchBatchRenameWindowMode(BatchRename::Mode::Manual),
                  L"Batch Rename sort-like-preview test should switch to Manual mode.");
    state.Require(DebugSetBatchRenameWindowManualText(L"sortp-a.txt\nz-renamed.txt\ny-renamed.txt"),
                  L"Batch Rename sort-like-preview test should set manual names with one unchanged row.");
    state.Require(DebugSetBatchRenameWindowHideUnchanged(true),
                  L"Batch Rename sort-like-preview test should enable Hide unchanged.");

    BatchRenameDebugSnapshot hiddenSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(hiddenSnapshot), L"Batch Rename hidden-row snapshot should be available.");
    state.Require(hiddenSnapshot.previewRowCount == 2u, L"Hide unchanged should leave only the two changed rows visible.");

    state.Require(DebugSetBatchRenameWindowPreviewSort(L"new", true),
                  L"Batch Rename sort-like-preview test should sort by new name descending.");

    BatchRenameDebugSnapshot sortedSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(sortedSnapshot), L"Batch Rename sorted hidden snapshot should be available.");
    state.Require(sortedSnapshot.newNames == std::vector<std::wstring>{L"z-renamed.txt", L"y-renamed.txt"},
                  L"Sorting with Hide unchanged should reorder the visible changed rows.");

    state.Require(DebugClickBatchRenameWindowManualSortLikePreview(),
                  L"Sort like preview must succeed while Hide unchanged filters the visible rows.");

    BatchRenameDebugSnapshot afterSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(afterSnapshot), L"Batch Rename post-sort-like-preview snapshot should be available.");
    state.Require(afterSnapshot.manualText == L"z-renamed.txt\ny-renamed.txt\nsortp-a.txt",
                  std::format(L"Sort like preview should rewrite all manual lines (including hidden rows) into full preview order; saw '{}'.",
                              afterSnapshot.manualText));
    state.Require(afterSnapshot.originalNames == std::vector<std::wstring>{L"sortp-b.txt", L"sortp-c.txt"},
                  L"Sort like preview should keep the visible changed rows sorted.");
    state.Require(afterSnapshot.newNames == std::vector<std::wstring>{L"z-renamed.txt", L"y-renamed.txt"},
                  L"Sort like preview should keep each visible target paired with its manual line.");
    state.Require(afterSnapshot.errorRowCount == 0u, L"Sort like preview with Hide unchanged should not introduce validation errors.");

    state.Require(DebugSetBatchRenameWindowHideUnchanged(false),
                  L"Batch Rename sort-like-preview test should restore unchanged rows.");

    BatchRenameDebugSnapshot restoredSnapshot{};
    state.Require(DebugGetBatchRenameWindowSnapshot(restoredSnapshot), L"Batch Rename restored snapshot should be available.");
    state.Require(restoredSnapshot.previewRowCount == 3u, L"Restoring unchanged rows should show the full reordered preview.");

    return state.failure.empty();
}

[[nodiscard]] bool TestBatchRenamePathSortUsesDisplayedPathAndStableTies(HWND mainWindow, CaseState& state) noexcept
{
    const std::filesystem::path root = L"C:\\BatchRenamePathSortSelfTest";

    BatchRenamePaneContext context{};
    context.pluginId        = L"builtin/file-system";
    context.pluginShortId   = L"local";
    context.instanceContext = L"batch-rename-path-sort-selftest";
    context.rootPluginPath  = root;
    context.initialPaths    = {root / L"z.txt", root / L"same-b.txt", root / L"same-a.txt", root / L"a" / L"a.txt"};

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"batch-rename-path-sort-window-selftest");
    state.Require(ShowBatchRenameWindow(mainWindow, g_settings, theme, std::move(context)),
                  L"Batch Rename window should open for path-sort testing.");

    const HWND batchWindow =
        WaitForWindow([]() noexcept { return GetBatchRenameWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds(5000)));
    state.Require(batchWindow != nullptr && IsWindow(batchWindow) != FALSE, L"Batch Rename path-sort test window should become visible.");
    if (! batchWindow || IsWindow(batchWindow) == FALSE)
    {
        return false;
    }

    const auto cleanupBatchWindow = wil::scope_exit([]() noexcept { CloseBatchRenameWindowIfOpen(); });

    state.Require(DebugSetBatchRenameWindowPreviewSort(L"path", false), L"Batch Rename path-sort test should sort by Path ascending.");

    BatchRenameDebugSnapshot ascending{};
    state.Require(DebugGetBatchRenameWindowSnapshot(ascending), L"Batch Rename ascending path-sort snapshot should be available.");
    state.Require(ascending.originalNames == std::vector<std::wstring>{L"z.txt", L"same-b.txt", L"same-a.txt", L"a.txt"},
                  L"Path ascending sort should use displayed root-relative paths and preserve source order among empty-path ties.");

    state.Require(DebugSetBatchRenameWindowPreviewSort(L"path", true), L"Batch Rename path-sort test should sort by Path descending.");

    BatchRenameDebugSnapshot descending{};
    state.Require(DebugGetBatchRenameWindowSnapshot(descending), L"Batch Rename descending path-sort snapshot should be available.");
    state.Require(descending.originalNames == std::vector<std::wstring>{L"a.txt", L"z.txt", L"same-b.txt", L"same-a.txt"},
                  L"Path descending sort should reverse only the displayed path key and preserve source order among ties.");

    return state.failure.empty();
}

void RunBatchRenameCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    const auto runCase = [&](std::wstring_view name, auto&& func) noexcept {
        SelfTest::RunCase(options, suite, name, [&](CaseState& state) noexcept {
            BatchRenameSettingsTestScope settingsScope;
            settingsScope.ResetForDefaultWindow();
            return func(state);
        });
    };

    runCase(L"cmd_pane_batchRename_command_registered", [](CaseState& state) noexcept {
        return TestBatchRenameCommandRegistered(state);
    });
    runCase(L"cmd_pane_batchRename_opens_from_active_pane", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowOpensFromPaneContext(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_folder_scope_collects_local_children_metadata", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowFolderScopeCollectsLocalChildrenMetadata(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_recursive_collection_skips_symlink_loops", [](CaseState& state) noexcept {
        return TestBatchRenameRecursiveCollectionSkipsSymlinkLoops(state);
    });
    runCase(L"cmd_pane_batchRename_nonlocal_selection_fallback_metadata_unknown", [](CaseState& state) noexcept {
        return TestBatchRenameProviderSelectionFallbackMarksMetadataUnknown(state);
    });
    runCase(L"cmd_pane_batchRename_window_rules_recompute_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowRulesRecomputePreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_datetime_columns_match_macro_expansion", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowDateTimeColumnsMatchMacroExpansion(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_created_macro_uses_collected_creation_time", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowCreatedMacroUsesCollectedCreationTime(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_preview_context_menu_copies_rows", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowPreviewContextMenuCopiesRows(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_preview_clipboard_honors_display_order", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowPreviewClipboardHonorsDisplayOrder(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_stale_generation_payloads_are_ignored", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowIgnoresStaleGenerationPayloads(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_theme_accessibility_snapshot", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowThemeAccessibilitySnapshot(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_rule_controls_drive_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowRuleControlsDrivePreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_debounces_text_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowDebouncesTextPreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_uses_and_persists_settings", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowUsesAndPersistsSettings(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_manual_mode_controls_drive_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowManualModeControlsDrivePreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_manual_mode_target_change_blocks_until_reconciled", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowManualModeTargetChangeBlocksUntilReconciled(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_path_sort_uses_displayed_path_and_stable_ties", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenamePathSortUsesDisplayedPathAndStableTies(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_local_rename", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesLocalRename(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_refreshes_pane_after_success", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowRefreshesPaneAfterSuccess(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_invokes_success_callback", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowInvokesSuccessCallback(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_success_callback_parent_child_execution_order", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowSuccessCallbackParentChildExecutionOrder(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_parent_child_directory_cache_notify_retargets_pinned_descendant", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowParentChildDirectoryCacheNotifyRetargetsPinnedDescendant(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_target_collection_respects_provider_cancellation", [](CaseState& state) noexcept {
        return TestBatchRenameTargetCollectionRespectsProviderCancellation(state);
    });
    runCase(L"cmd_pane_batchRename_window_reports_execution_summary", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowReportsExecutionSummary(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_collision_existing_and_duplicate_names", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameCollisionExistingAndDuplicateNames(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_destination_probe_failure_issue_id", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameDestinationProbeFailureIssueId(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_cancel_does_not_apply_remaining_rows", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameCancelDoesNotApplyRemainingRows(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_falls_back_to_rename_item_when_bulk_unsupported", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowFallsBackToRenameItemWhenBulkUnsupported(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_swap_rename", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesSwapRename(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_chain_rename", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesChainRename(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_directory_chain_undo_plan", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowDirectoryChainUndoPlan(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_case_only_rename_not_treated_as_dependency", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameCaseOnlyRenameNotTreatedAsDependency(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_partial_batch_failure_tracks_completed_rows", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenamePartialBatchFailureTracksCompletedRows(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_omitted_item_completion_counts_row_failed", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameOmittedItemCompletionCountsRowFailed(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_temp_hop_reported_failure_counts_row_failed", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameTempHopReportedFailureCountsRowFailed(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_swap_failure_rolls_back_orphan_temp", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameSwapFailureRollsBackOrphanTemp(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_swap_cancel_after_temp_hop_rolls_back", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameSwapCancelAfterTempHopRollsBack(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_three_member_cycle", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesThreeMemberCycle(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_cancel_mid_batch_tracks_completed_rows", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameCancelMidBatchTracksCompletedRows(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_execute_while_busy_returns_busy", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameExecuteWhileBusyReturnsBusy(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_close_while_execution_inflight_is_bounded", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameCloseWhileExecutionInFlightIsBounded(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_report_survives_view_only_refreshes", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameReportSurvivesViewOnlyRefreshes(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_post_execution_target_refresh_is_linear", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenamePostExecutionTargetRefreshIsLinear(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_duplicate_source_target_refresh_consumes_distinct_rows", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameDuplicateSourceTargetRefreshConsumesDistinctRows(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_manual_sort_like_preview_with_hide_unchanged", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameManualSortLikePreviewWithHideUnchanged(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_case_only_local_rename", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesCaseOnlyLocalRename(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_blocks_invalid_preview_execution", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowBlocksInvalidPreviewExecution(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_blocks_missing_provider_path_identity", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowBlocksProviderWithoutPathIdentity(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_executes_parent_child_deepest_first", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowExecutesParentChildDeepestFirst(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_blocks_destination_created_after_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowBlocksDestinationCreatedAfterPreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_window_blocks_source_missing_after_preview", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowBlocksSourceMissingAfterPreview(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_helper_menus_expose_canonical_insertions", [](CaseState& state) noexcept {
        return TestBatchRenameHelperMenusExposeCanonicalInsertions(state);
    });
    runCase(L"cmd_pane_batchRename_window_helper_buttons_insert_into_rule_fields", [mainWindow](CaseState& state) noexcept {
        return TestBatchRenameWindowHelperButtonsInsertIntoRuleFields(mainWindow, state);
    });
    runCase(L"cmd_pane_rename_multi_selection_opens_batch_rename", [mainWindow](CaseState& state) noexcept {
        return TestPaneRenameMultiSelectionOpensBatchRename(mainWindow, state);
    });
    runCase(L"cmd_pane_rename_single_file_uses_standard_prompt", [mainWindow](CaseState& state) noexcept {
        return TestPaneRenameSingleFileUsesStandardPrompt(mainWindow, state);
    });
    runCase(L"cmd_pane_rename_file_prompt_batch_button_opens_batch_rename", [mainWindow](CaseState& state) noexcept {
        return TestPaneRenameFilePromptBatchButtonOpensBatchRename(mainWindow, state);
    });
    runCase(L"cmd_pane_rename_folder_prompt_batch_button_opens_batch_rename", [mainWindow](CaseState& state) noexcept {
        return TestPaneRenameFolderPromptBatchButtonOpensBatchRename(mainWindow, state);
    });
    runCase(L"cmd_pane_batchRename_preview_macro_regex_case_validation", [](CaseState& state) noexcept {
        return TestBatchRenameEngineMacroRegexCaseAndValidation(state);
    });
    runCase(L"cmd_pane_batchRename_engine_provider_path_identity", [](CaseState& state) noexcept {
        return TestBatchRenameEngineUsesProviderPathIdentity(state);
    });
    runCase(L"cmd_pane_batchRename_path_identity_parser", [](CaseState& state) noexcept {
        return TestBatchRenamePathIdentityParserRejectsUnplannableProfiles(state);
    });
    runCase(L"cmd_pane_batchRename_engine_macro_alias_datetime_regex_tokens", [](CaseState& state) noexcept {
        return TestBatchRenameEngineMacroAliasDateTimeAndRegexReplacementTokens(state);
    });
    runCase(L"cmd_pane_batchRename_engine_leaf_splitter_matches_filesystem", [](CaseState& state) noexcept {
        return TestBatchRenameEngineLeafSplitterMatchesFilesystem(state);
    });
    runCase(L"cmd_pane_batchRename_engine_recompute_stats_after_contextual_issue", [](CaseState& state) noexcept {
        return TestBatchRenameEngineRecomputeStatsAfterContextualIssue(state);
    });
    runCase(L"cmd_pane_batchRename_manual_mode_line_count_validation", [](CaseState& state) noexcept {
        return TestBatchRenameEngineManualModeLineValidation(state);
    });
    runCase(L"cmd_pane_batchRename_engine_warnings_and_remaining_transforms", [](CaseState& state) noexcept {
        return TestBatchRenameEngineWarningsAndRemainingTransforms(state);
    });
    runCase(L"cmd_pane_batchRename_engine_remaining_validation_transform_coverage", [](CaseState& state) noexcept {
        return TestBatchRenameEngineRemainingValidationAndTransformCoverage(state);
    });
    runCase(L"cmd_pane_batchRename_engine_regex_match_failure_and_whole_word_captures", [](CaseState& state) noexcept {
        return TestBatchRenameEngineRegexMatchFailureAndWholeWordCaptures(state);
    });
    runCase(L"cmd_pane_batchRename_engine_reserved_device_names_and_counter_width", [](CaseState& state) noexcept {
        return TestBatchRenameEngineReservedDeviceNamesAndCounterWidth(state);
    });
    runCase(L"cmd_pane_batchRename_engine_large_preview_perf", [](CaseState& state) noexcept {
        return TestBatchRenameEngineLargePreviewEmitsPerfMetrics(state);
    });
    runCase(L"cmd_pane_batchRename_engine_regex_compile_perf", [](CaseState& state) noexcept {
        return TestBatchRenameEngineRegexCompileEmitsPerfMetric(state);
    });
    runCase(L"cmd_pane_batchRename_engine_validation_perf", [](CaseState& state) noexcept {
        return TestBatchRenameEngineValidationEmitsPerfMetric(state);
    });
    runCase(L"cmd_pane_batchRename_execution_engine_direct", [](CaseState& state) noexcept {
        return TestBatchRenameExecutionEngineDirectScenarios(state);
    });
}

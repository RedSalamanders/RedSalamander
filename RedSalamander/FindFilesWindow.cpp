#include "Framework.h"

#include "FindFilesWindow.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <format>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include <CommCtrl.h>
#include <WindowsX.h>

#include "FolderWindow.h"
#include "Helpers.h"
#include "NavigationLocation.h"
#include "SearchFallbackEngine.h"
#include "ThemedControls.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

extern FolderWindow g_folderWindow;

namespace
{
constexpr wchar_t kFindFilesWindowClassName[] = L"RedSalamander.FindFilesWindow";
constexpr wchar_t kFindFilesWindowId[]        = L"FindFilesWindow";

constexpr uint32_t kMaxSnippetCharacters = 160u;
constexpr uint64_t kMaxContentBytes      = 64ull * 1024ull * 1024ull;
constexpr size_t kMaxRecentEntries       = 10u;
constexpr size_t kBatchSize              = 32u;

enum class SearchOperation : uint8_t
{
    Find,
    Append,
    Intersect,
    Subtract,
};

enum ControlId : int
{
    kRootLabelId = 4100,
    kRootComboId,
    kNameLabelId,
    kNameComboId,
    kNameModeComboId,
    kContentLabelId,
    kContentComboId,
    kContentModeComboId,
    kRecursiveCheckId,
    kIncludeFilesCheckId,
    kIncludeDirectoriesCheckId,
    kFollowSymlinksCheckId,
    kMatchCaseNameCheckId,
    kMatchCaseContentCheckId,
    kPreferIndexCheckId,
    kWantSnippetsCheckId,
    kFindButtonId,
    kAppendButtonId,
    kIntersectButtonId,
    kSubtractButtonId,
    kCancelButtonId,
    kOpenButtonId,
    kParentButtonId,
    kStatusTextId,
    kResultsListId,
};

enum ColumnIndex : int
{
    kColumnName = 0,
    kColumnPath,
    kColumnSize,
    kColumnModified,
    kColumnAttributes,
    kColumnSnippet,
};

struct FindResultRecord
{
    std::wstring key;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::wstring fullPath;
    std::wstring relativePath;
    std::wstring displayName;
    std::wstring previewText;
    unsigned long fileAttributes = 0;
    int64_t lastWriteTime        = 0;
    int64_t endOfFile            = 0;
    uint32_t matchedBy           = 0;
};

struct FindSearchResultsPayload
{
    std::vector<FindResultRecord> results;
};

struct FindSearchProgressPayload
{
    FileSystemSearchPhase phase         = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    FileSystemSearchBackend backend     = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t warningFlags               = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint                  = S_OK;
    uint64_t scannedDirectories         = 0;
    uint64_t scannedFiles               = 0;
    uint64_t candidateFiles             = 0;
    uint64_t matchedEntries             = 0;
    std::wstring currentPath;
};

struct FindSearchCompletePayload
{
    HRESULT hr                          = S_OK;
    FileSystemSearchBackend backend     = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t warningFlags               = FILESYSTEM_SEARCH_WARNING_NONE;
    uint64_t scannedDirectories         = 0;
    uint64_t scannedFiles               = 0;
    uint64_t candidateFiles             = 0;
    uint64_t matchedEntries             = 0;
};

struct SearchRequest
{
    FindFilesPaneContext context;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    FileSystemSearchFlags flags              = FILESYSTEM_SEARCH_NONE;
    FileSystemSearchNameMode nameMode        = FILESYSTEM_SEARCH_NAME_DISABLED;
    FileSystemSearchContentMode contentMode  = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    uint64_t maxResults                      = 0;
    uint64_t maxContentBytesPerFile          = 0;
    uint32_t maxSnippetCharacters            = 0;
};

[[nodiscard]] std::wstring GetWindowTextString(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(length) + 1u, L'\0');
    const int copied = GetWindowTextW(hwnd, result.data(), static_cast<int>(result.size()));
    if (copied <= 0)
    {
        return {};
    }

    result.resize(static_cast<size_t>(copied));
    return result;
}

[[nodiscard]] std::wstring ToLowerCopy(std::wstring_view value) noexcept
{
    std::wstring result(value);
    std::transform(result.begin(), result.end(), result.begin(), [](wchar_t ch) noexcept
    {
        return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
    });
    return result;
}

[[nodiscard]] std::wstring MakeResultKey(std::wstring_view pluginId, std::wstring_view instanceContext, std::wstring_view fullPath) noexcept
{
    std::wstring key = ToLowerCopy(pluginId);
    key.push_back(L'|');
    key.append(ToLowerCopy(instanceContext));
    key.push_back(L'|');
    key.append(ToLowerCopy(fullPath));
    return key;
}

void UpdateRecentValue(std::vector<std::wstring>& history, std::wstring value) noexcept
{
    if (value.empty())
    {
        return;
    }

    const auto it = std::find_if(history.begin(), history.end(), [&](const std::wstring& existing) noexcept
    {
        return OrdinalString::EqualsNoCase(existing, value);
    });
    if (it != history.end())
    {
        history.erase(it);
    }

    history.insert(history.begin(), std::move(value));
    if (history.size() > kMaxRecentEntries)
    {
        history.resize(kMaxRecentEntries);
    }
}

[[nodiscard]] bool IsSearchUnsupported(HRESULT hr) noexcept
{
    return hr == E_NOTIMPL || hr == HRESULT_FROM_WIN32(ERROR_CALL_NOT_IMPLEMENTED) || hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

[[nodiscard]] uint32_t ToComboIndex(Common::Settings::SearchNameMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchNameMode::Wildcard: return 0u;
        case Common::Settings::SearchNameMode::Literal: return 1u;
        case Common::Settings::SearchNameMode::Regex: return 2u;
    }

    return 0u;
}

[[nodiscard]] Common::Settings::SearchNameMode FromNameModeComboIndex(int index) noexcept
{
    switch (index)
    {
        case 1: return Common::Settings::SearchNameMode::Literal;
        case 2: return Common::Settings::SearchNameMode::Regex;
        default: return Common::Settings::SearchNameMode::Wildcard;
    }
}

[[nodiscard]] uint32_t ToComboIndex(Common::Settings::SearchContentMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchContentMode::Disabled: return 0u;
        case Common::Settings::SearchContentMode::TextLiteral: return 1u;
        case Common::Settings::SearchContentMode::TextRegex: return 2u;
    }

    return 0u;
}

[[nodiscard]] Common::Settings::SearchContentMode FromContentModeComboIndex(int index) noexcept
{
    switch (index)
    {
        case 1: return Common::Settings::SearchContentMode::TextLiteral;
        case 2: return Common::Settings::SearchContentMode::TextRegex;
        default: return Common::Settings::SearchContentMode::Disabled;
    }
}

[[nodiscard]] FileSystemSearchNameMode ToAbiNameMode(Common::Settings::SearchNameMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchNameMode::Wildcard: return FILESYSTEM_SEARCH_NAME_WILDCARD;
        case Common::Settings::SearchNameMode::Literal: return FILESYSTEM_SEARCH_NAME_LITERAL;
        case Common::Settings::SearchNameMode::Regex: return FILESYSTEM_SEARCH_NAME_REGEX;
    }

    return FILESYSTEM_SEARCH_NAME_WILDCARD;
}

[[nodiscard]] FileSystemSearchContentMode ToAbiContentMode(Common::Settings::SearchContentMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchContentMode::Disabled: return FILESYSTEM_SEARCH_CONTENT_DISABLED;
        case Common::Settings::SearchContentMode::TextLiteral: return FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
        case Common::Settings::SearchContentMode::TextRegex: return FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX;
    }

    return FILESYSTEM_SEARCH_CONTENT_DISABLED;
}

[[nodiscard]] std::wstring FormatFileSize(int64_t sizeBytes) noexcept
{
    if (sizeBytes <= 0)
    {
        return {};
    }

    return std::format(L"{:L}", sizeBytes);
}

[[nodiscard]] std::wstring FormatFileTimeValue(int64_t fileTimeTicks) noexcept
{
    if (fileTimeTicks <= 0)
    {
        return {};
    }

    FILETIME ft{};
    ft.dwLowDateTime  = static_cast<DWORD>(static_cast<uint64_t>(fileTimeTicks) & 0xFFFFFFFFull);
    ft.dwHighDateTime = static_cast<DWORD>((static_cast<uint64_t>(fileTimeTicks) >> 32u) & 0xFFFFFFFFull);

    SYSTEMTIME utc{};
    SYSTEMTIME local{};
    if (FileTimeToSystemTime(&ft, &utc) == FALSE || SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local) == FALSE)
    {
        return {};
    }

    return std::format(L"{:04}-{:02}-{:02} {:02}:{:02}", local.wYear, local.wMonth, local.wDay, local.wHour, local.wMinute);
}

[[nodiscard]] std::wstring FormatAttributes(unsigned long attributes) noexcept
{
    std::wstring value;
    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        value.push_back(L'D');
    }
    if ((attributes & FILE_ATTRIBUTE_ARCHIVE) != 0)
    {
        value.push_back(L'A');
    }
    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        value.push_back(L'R');
    }
    if ((attributes & FILE_ATTRIBUTE_HIDDEN) != 0)
    {
        value.push_back(L'H');
    }
    if ((attributes & FILE_ATTRIBUTE_SYSTEM) != 0)
    {
        value.push_back(L'S');
    }

    return value;
}

class FindFilesWindow;

class SearchSessionController
{
public:
    SearchSessionController()                              = default;
    SearchSessionController(const SearchSessionController&) = delete;
    SearchSessionController& operator=(const SearchSessionController&) = delete;

    ~SearchSessionController() noexcept;

    [[nodiscard]] bool Start(FindFilesWindow& owner, SearchRequest request) noexcept;
    void Cancel() noexcept;
    void Shutdown() noexcept;
    void NotifyUiSettled() noexcept;
    [[nodiscard]] bool IsActive() const noexcept;
#ifdef _DEBUG
    [[nodiscard]] bool IsUiSettled() const noexcept;
    [[nodiscard]] bool WaitForIdle(uint32_t timeoutMs) noexcept;
#endif

private:
    void Run(SearchRequest request) noexcept;
    void MarkIdle() noexcept;

    std::jthread _worker;
    std::atomic<bool> _cancelRequested{false};
    std::atomic<bool> _active{false};
    std::atomic<bool> _uiSettled{true};
    HWND _ownerHwnd = nullptr;
    FindFilesWindow* _owner = nullptr;
    mutable std::mutex _mutex;
    std::condition_variable _idleCv;
};

class FindFilesWindow
{
public:
    FindFilesWindow(HWND owner, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept;
    FindFilesWindow(const FindFilesWindow&)            = delete;
    FindFilesWindow& operator=(const FindFilesWindow&) = delete;

    [[nodiscard]] bool Create() noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateContext(FindFilesPaneContext context) noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept { return _hWnd.get(); }
    [[nodiscard]] bool IsSearchActive() const noexcept { return _session.IsActive(); }
#ifdef _DEBUG
    [[nodiscard]] bool DebugConfigure(std::wstring rootPath,
                                      std::wstring namePattern,
                                      std::wstring contentPattern,
                                      Common::Settings::SearchNameMode nameMode,
                                      Common::Settings::SearchContentMode contentMode) noexcept;
    [[nodiscard]] bool DebugSetOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept;
    [[nodiscard]] bool DebugStartSearch(FindFilesDebugOperation operation) noexcept;
    [[nodiscard]] bool DebugCancelSearch() noexcept;
    [[nodiscard]] bool DebugGetSnapshot(FindFilesDebugSnapshot& out) noexcept;
    [[nodiscard]] bool DebugWaitForIdle(uint32_t timeoutMs) noexcept;
#endif

    LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

private:
    // Implemented in follow-up patch chunks.
    void OnCreate(HWND hwnd) noexcept;
    void OnClose() noexcept;
    LRESULT OnNcDestroy() noexcept;
    void Layout() noexcept;
    void ApplyTheme() noexcept;
    void EnsureFonts() noexcept;
    void EnsureColumns() noexcept;
    void PopulateFromSettings() noexcept;
    void PopulateModeCombos() noexcept;
    void PopulateHistoryCombos() noexcept;
    void UpdateOptionDependencies() noexcept;
    void UpdateActionButtons() noexcept;
    void PersistUiState(bool updateHistory) noexcept;
    [[nodiscard]] bool BeginSearch(SearchOperation operation) noexcept;
    void OnSearchStarted(SearchOperation operation) noexcept;
    void OnSearchResults(std::unique_ptr<FindSearchResultsPayload> payload) noexcept;
    void OnSearchProgress(std::unique_ptr<FindSearchProgressPayload> payload) noexcept;
    void OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept;
    void ApplyDeferredSetOperation(SearchOperation operation) noexcept;
    void ClearResults() noexcept;
    void RebuildResultsList() noexcept;
    void AddOrUpdateVisibleResult(FindResultRecord result) noexcept;
    void RemoveKeysFromResults(const std::unordered_set<std::wstring>& keys) noexcept;
    void KeepOnlyKeysInResults(const std::unordered_set<std::wstring>& keys) noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedResultIndex() const noexcept;
    void OpenSelectedResult(bool parentOnly) noexcept;
    void SetStatusText(std::wstring text) noexcept;
    void RefreshStatusText() noexcept;
    [[nodiscard]] std::wstring BuildStatusText() const noexcept;
    [[nodiscard]] std::wstring BuildWarningSummary(uint32_t warningFlags) const noexcept;
    [[nodiscard]] UINT BackendStringId(FileSystemSearchBackend backend) const noexcept;
    [[nodiscard]] std::optional<SearchRequest> BuildSearchRequest() noexcept;
    [[nodiscard]] std::wstring GetComboText(HWND combo) const noexcept;

    HWND _ownerWindow = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    FindFilesPaneContext _context;

    wil::unique_hwnd _hWnd;
    wil::unique_hfont _uiFont;
    wil::unique_hbrush _backgroundBrush;

    HWND _rootLabel = nullptr;
    HWND _rootCombo = nullptr;
    HWND _nameLabel = nullptr;
    HWND _nameCombo = nullptr;
    HWND _nameModeCombo = nullptr;
    HWND _contentLabel = nullptr;
    HWND _contentCombo = nullptr;
    HWND _contentModeCombo = nullptr;
    HWND _recursiveCheck = nullptr;
    HWND _includeFilesCheck = nullptr;
    HWND _includeDirectoriesCheck = nullptr;
    HWND _followSymlinksCheck = nullptr;
    HWND _matchCaseNameCheck = nullptr;
    HWND _matchCaseContentCheck = nullptr;
    HWND _preferIndexCheck = nullptr;
    HWND _wantSnippetsCheck = nullptr;
    HWND _findButton = nullptr;
    HWND _appendButton = nullptr;
    HWND _intersectButton = nullptr;
    HWND _subtractButton = nullptr;
    HWND _cancelButton = nullptr;
    HWND _openButton = nullptr;
    HWND _parentButton = nullptr;
    HWND _statusText = nullptr;
    HWND _resultsList = nullptr;

    SearchSessionController _session;
    SearchOperation _activeOperation = SearchOperation::Find;
    FileSystemSearchBackend _lastBackend = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t _lastWarningFlags = FILESYSTEM_SEARCH_WARNING_NONE;
    uint64_t _lastScannedDirectories = 0;
    uint64_t _lastScannedFiles = 0;
    uint64_t _lastCandidateFiles = 0;
    uint64_t _lastMatchedEntries = 0;
    HRESULT _lastStatusHint = S_OK;
    bool _cancelRequestedUi = false;
    bool _closeRequested = false;
    std::wstring _status;
    std::vector<FindResultRecord> _results;
    std::unordered_map<std::wstring, size_t> _resultIndexByKey;
    std::unordered_set<std::wstring> _deferredKeys;
};

FindFilesWindow* g_findFilesWindow = nullptr;
std::vector<HWND> g_findFilesWindows;

[[nodiscard]] std::wstring CopySizedUtf16(const wchar_t* value, unsigned long sizeBytes) noexcept
{
    if (! value || sizeBytes == 0)
    {
        return {};
    }

    size_t charCount = sizeBytes / sizeof(wchar_t);
    while (charCount > 0 && value[charCount - 1u] == L'\0')
    {
        --charCount;
    }

    return std::wstring(value, value + charCount);
}

struct SearchCallbacks final : IFileSystemSearchCallback
{
    explicit SearchCallbacks(HWND hwnd, const SearchRequest& request, std::atomic<bool>& cancelRequested) noexcept :
        _hwnd(hwnd),
        _request(request),
        _cancelRequested(cancelRequested)
    {
    }

    SearchCallbacks(const SearchCallbacks&) = delete;
    SearchCallbacks& operator=(const SearchCallbacks&) = delete;
    SearchCallbacks(SearchCallbacks&&) = delete;
    SearchCallbacks& operator=(SearchCallbacks&&) = delete;

    [[nodiscard]] bool FlushResults() noexcept
    {
        if (_batch.empty())
        {
            return true;
        }

        auto payload = std::unique_ptr<FindSearchResultsPayload>(new (std::nothrow) FindSearchResultsPayload{});
        if (! payload)
        {
            return false;
        }

        payload->results = std::move(_batch);
        _batch.clear();
        return PostMessagePayload(_hwnd, WndMsg::kFindSearchResults, 0, std::move(payload));
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const ::FileSystemSearchMatch* match,
                                                    void* /*cookie*/) noexcept override
    {
        if (! match || match->sizeBytes != sizeof(::FileSystemSearchMatch))
        {
            return E_INVALIDARG;
        }

        FindResultRecord record;
        record.pluginId        = _request.context.pluginId;
        record.pluginShortId   = _request.context.pluginShortId;
        record.instanceContext = _request.context.instanceContext;
        record.fullPath        = CopySizedUtf16(match->fullPath, match->fullPathSize);
        record.relativePath    = CopySizedUtf16(match->relativePath, match->relativePathSize);
        record.displayName     = CopySizedUtf16(match->displayName, match->displayNameSize);
        record.previewText     = CopySizedUtf16(match->previewText, match->previewTextSize);
        record.fileAttributes  = match->fileAttributes;
        record.lastWriteTime   = match->lastWriteTime;
        record.endOfFile       = match->endOfFile;
        record.matchedBy       = match->matchedBy;
        record.key             = MakeResultKey(record.pluginId, record.instanceContext, record.fullPath);

        _batch.push_back(std::move(record));
        if (_batch.size() >= kBatchSize && ! FlushResults())
        {
            return E_FAIL;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const ::FileSystemSearchProgress* progress,
                                                       void* /*cookie*/) noexcept override
    {
        if (! progress || progress->sizeBytes != sizeof(::FileSystemSearchProgress))
        {
            return E_INVALIDARG;
        }

        if (! FlushResults())
        {
            return E_FAIL;
        }

        auto payload = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
        if (! payload)
        {
            return E_OUTOFMEMORY;
        }

        payload->phase              = progress->phase;
        payload->backend            = progress->backend;
        payload->warningFlags       = progress->warningFlags;
        payload->statusHint         = progress->statusHint;
        payload->scannedDirectories = progress->scannedDirectories;
        payload->scannedFiles       = progress->scannedFiles;
        payload->candidateFiles     = progress->candidateFiles;
        payload->matchedEntries     = progress->matchedEntries;
        payload->currentPath        = CopySizedUtf16(progress->currentPath, progress->currentPathSize);
        _latestProgress             = *payload;

        return PostMessagePayload(_hwnd, WndMsg::kFindSearchProgress, 0, std::move(payload)) ? S_OK : E_FAIL;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept override
    {
        if (! pCancel)
        {
            return E_POINTER;
        }

        *pCancel = _cancelRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        return S_OK;
    }

    HWND _hwnd = nullptr;
    const SearchRequest& _request;
    std::atomic<bool>& _cancelRequested;
    std::vector<FindResultRecord> _batch;
    FindSearchProgressPayload _latestProgress;
};

SearchSessionController::~SearchSessionController() noexcept
{
    Shutdown();
}

bool SearchSessionController::Start(FindFilesWindow& owner, SearchRequest request) noexcept
{
    bool expected = false;
    if (! _active.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
    {
        return false;
    }

    _cancelRequested.store(false, std::memory_order_release);
    _uiSettled.store(false, std::memory_order_release);
    _ownerHwnd = owner.GetHwnd();
    _owner     = &owner;
    try
    {
        _worker = std::jthread([this, request = std::move(request)]() mutable noexcept { Run(std::move(request)); });
    }
    catch (const std::system_error&)
    {
        // Thread creation is required for the modeless Find dialog; fall back to a clean start failure.
        _owner = nullptr;
        _ownerHwnd = nullptr;
        _uiSettled.store(true, std::memory_order_release);
        MarkIdle();
        return false;
    }
    return true;
}

void SearchSessionController::Cancel() noexcept
{
    _cancelRequested.store(true, std::memory_order_release);
}

void SearchSessionController::Shutdown() noexcept
{
    Cancel();
    if (_worker.joinable())
    {
        _worker.join();
    }

    _owner = nullptr;
    _ownerHwnd = nullptr;
    _uiSettled.store(true, std::memory_order_release);
    MarkIdle();
}

void SearchSessionController::NotifyUiSettled() noexcept
{
    _uiSettled.store(true, std::memory_order_release);
    _idleCv.notify_all();
}

bool SearchSessionController::IsActive() const noexcept
{
    return _active.load(std::memory_order_acquire);
}

#ifdef _DEBUG
bool SearchSessionController::WaitForIdle(uint32_t timeoutMs) noexcept
{
    std::unique_lock lock(_mutex);
    return _idleCv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [this]() noexcept
    {
        return ! _active.load(std::memory_order_acquire);
    });
}

bool SearchSessionController::IsUiSettled() const noexcept
{
    return _uiSettled.load(std::memory_order_acquire);
}
#endif

void SearchSessionController::Run(SearchRequest request) noexcept
{
    auto initialPayload = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
    if (initialPayload)
    {
        initialPayload->phase      = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
        initialPayload->backend    = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
        initialPayload->statusHint = S_OK;
        static_cast<void>(PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchProgress, 0, std::move(initialPayload)));
    }

    SearchCallbacks callbacks(_ownerHwnd, request, _cancelRequested);

    FileSystemSearchQuery query{};
    query.sizeBytes              = sizeof(query);
    query.rootPath               = request.rootPath.c_str();
    query.namePattern            = request.nameMode == FILESYSTEM_SEARCH_NAME_DISABLED ? nullptr : request.namePattern.c_str();
    query.contentPattern         = request.contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? nullptr : request.contentPattern.c_str();
    query.flags                  = request.flags;
    query.nameMode               = request.nameMode;
    query.contentMode            = request.contentMode;
    query.maxResults             = request.maxResults;
    query.maxContentBytesPerFile = request.maxContentBytesPerFile;
    query.maxSnippetCharacters   = request.maxSnippetCharacters;
    query.reserved               = 0;

    HRESULT hr = E_NOINTERFACE;
    if (request.context.fileSystem)
    {
        wil::com_ptr<IFileSystemSearch> nativeSearch;
        if (SUCCEEDED(request.context.fileSystem->QueryInterface(IID_PPV_ARGS(nativeSearch.put()))) && nativeSearch)
        {
            hr = nativeSearch->Search(&query, &callbacks, nullptr);
            if (IsSearchUnsupported(hr) && ! _cancelRequested.load(std::memory_order_acquire))
            {
                hr = SearchFallbackEngine::Execute(request.context.fileSystem.get(), &query, &callbacks, nullptr);
            }
        }
        else
        {
            hr = SearchFallbackEngine::Execute(request.context.fileSystem.get(), &query, &callbacks, nullptr);
        }
    }

    static_cast<void>(callbacks.FlushResults());

    bool completionQueued = false;
    auto complete = std::unique_ptr<FindSearchCompletePayload>(new (std::nothrow) FindSearchCompletePayload{});
    if (complete)
    {
        complete->hr                 = hr;
        complete->backend            = callbacks._latestProgress.backend;
        complete->warningFlags       = callbacks._latestProgress.warningFlags;
        complete->scannedDirectories = callbacks._latestProgress.scannedDirectories;
        complete->scannedFiles       = callbacks._latestProgress.scannedFiles;
        complete->candidateFiles     = callbacks._latestProgress.candidateFiles;
        complete->matchedEntries     = callbacks._latestProgress.matchedEntries;
        completionQueued             = PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchComplete, 0, std::move(complete));
    }

    if (! completionQueued)
    {
        NotifyUiSettled();
    }

    MarkIdle();
}

void SearchSessionController::MarkIdle() noexcept
{
    {
        std::lock_guard lock(_mutex);
        _active.store(false, std::memory_order_release);
    }
    _idleCv.notify_all();
}

FindFilesWindow::FindFilesWindow(HWND owner, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept :
    _ownerWindow(owner),
    _settings(&settings),
    _theme(std::move(theme)),
    _context(std::move(context))
{
}

bool FindFilesWindow::Create() noexcept
{
    HINSTANCE instance = GetModuleHandleW(nullptr);

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = &FindFilesWindow::WndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.lpszClassName = kFindFilesWindowClassName;
    if (RegisterClassExW(&wc) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
    {
        return false;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FIND_TITLE);
    const HWND created       = CreateWindowExW(0,
                                         kFindFilesWindowClassName,
                                         title.c_str(),
                                         WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                         CW_USEDEFAULT,
                                         CW_USEDEFAULT,
                                         ThemedControls::ScaleDip(96u, 1120),
                                         ThemedControls::ScaleDip(96u, 760),
                                         _ownerWindow,
                                         nullptr,
                                         instance,
                                         this);
    if (! created)
    {
        return false;
    }

    const bool hasPlacement = _settings && _settings->windows.contains(std::wstring(kFindFilesWindowId));
    const int showCmd       = hasPlacement ? WindowPlacementPersistence::Restore(*_settings, kFindFilesWindowId, created) : SW_SHOWNORMAL;
    ShowWindow(created, showCmd);
    SetForegroundWindow(created);
    return true;
}

void FindFilesWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
}

void FindFilesWindow::UpdateContext(FindFilesPaneContext context) noexcept
{
    const bool pluginChanged = ! OrdinalString::EqualsNoCase(_context.pluginId, context.pluginId) ||
                               ! OrdinalString::EqualsNoCase(_context.instanceContext, context.instanceContext);
    _context = std::move(context);
    if (pluginChanged)
    {
        _deferredKeys.clear();
        if (! _session.IsActive())
        {
            ClearResults();
        }

        const std::wstring rootText = (!_settings || ! _settings->search.has_value() || _settings->search->lastRoot.empty())
                                          ? _context.rootPluginPath.native()
                                          : _settings->search->lastRoot;
        if (_rootCombo)
        {
            SetWindowTextW(_rootCombo, rootText.c_str());
        }
    }
}

void FindFilesWindow::OnCreate(HWND hwnd) noexcept
{
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icc);

    HINSTANCE instance = GetModuleHandleW(nullptr);
    const DWORD comboStyle  = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL;
    const DWORD modeStyle   = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
    const DWORD checkStyle  = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    const DWORD buttonStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;

    _rootLabel = CreateWindowExW(0, L"Static", LoadStringResource(nullptr, IDS_FIND_LABEL_ROOT).c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, hwnd,
                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRootLabelId)), instance, nullptr);
    _rootCombo = CreateWindowExW(WS_EX_CLIENTEDGE, L"ComboBox", L"", comboStyle, 0, 0, 10, 200, hwnd,
                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRootComboId)), instance, nullptr);
    _nameLabel = CreateWindowExW(0, L"Static", LoadStringResource(nullptr, IDS_FIND_LABEL_NAME).c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, hwnd,
                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(kNameLabelId)), instance, nullptr);
    _nameCombo = CreateWindowExW(WS_EX_CLIENTEDGE, L"ComboBox", L"", comboStyle, 0, 0, 10, 200, hwnd,
                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(kNameComboId)), instance, nullptr);
    _nameModeCombo = CreateWindowExW(WS_EX_CLIENTEDGE, L"ComboBox", L"", modeStyle, 0, 0, 10, 160, hwnd,
                                     reinterpret_cast<HMENU>(static_cast<INT_PTR>(kNameModeComboId)), instance, nullptr);
    _contentLabel = CreateWindowExW(0,
                                    L"Static",
                                    LoadStringResource(nullptr, IDS_FIND_LABEL_CONTENT).c_str(),
                                    WS_CHILD | WS_VISIBLE | SS_LEFT,
                                    0,
                                    0,
                                    10,
                                    10,
                                    hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kContentLabelId)),
                                    instance,
                                    nullptr);
    _contentCombo = CreateWindowExW(WS_EX_CLIENTEDGE, L"ComboBox", L"", comboStyle, 0, 0, 10, 200, hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kContentComboId)), instance, nullptr);
    _contentModeCombo = CreateWindowExW(WS_EX_CLIENTEDGE, L"ComboBox", L"", modeStyle, 0, 0, 10, 160, hwnd,
                                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kContentModeComboId)), instance, nullptr);

    _recursiveCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_RECURSIVE).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRecursiveCheckId)), instance, nullptr);
    _includeFilesCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_INCLUDE_FILES).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                         reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIncludeFilesCheckId)), instance, nullptr);
    _includeDirectoriesCheck = CreateWindowExW(0,
                                               L"Button",
                                               LoadStringResource(nullptr, IDS_FIND_OPTION_INCLUDE_DIRECTORIES).c_str(),
                                               checkStyle,
                                               0,
                                               0,
                                               10,
                                               10,
                                               hwnd,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIncludeDirectoriesCheckId)),
                                               instance,
                                               nullptr);
    _followSymlinksCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_FOLLOW_SYMLINKS).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                           reinterpret_cast<HMENU>(static_cast<INT_PTR>(kFollowSymlinksCheckId)), instance, nullptr);
    _matchCaseNameCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_MATCH_CASE_NAME).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                          reinterpret_cast<HMENU>(static_cast<INT_PTR>(kMatchCaseNameCheckId)), instance, nullptr);
    _matchCaseContentCheck = CreateWindowExW(0,
                                             L"Button",
                                             LoadStringResource(nullptr, IDS_FIND_OPTION_MATCH_CASE_CONTENT).c_str(),
                                             checkStyle,
                                             0,
                                             0,
                                             10,
                                             10,
                                             hwnd,
                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(kMatchCaseContentCheckId)),
                                             instance,
                                             nullptr);
    _preferIndexCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_PREFER_INDEX).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kPreferIndexCheckId)), instance, nullptr);
    _wantSnippetsCheck = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_OPTION_WANT_SNIPPETS).c_str(), checkStyle, 0, 0, 10, 10, hwnd,
                                         reinterpret_cast<HMENU>(static_cast<INT_PTR>(kWantSnippetsCheckId)), instance, nullptr);

    _findButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_ACTION_FIND).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(kFindButtonId)), instance, nullptr);
    _appendButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_ACTION_APPEND).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kAppendButtonId)), instance, nullptr);
    _intersectButton = CreateWindowExW(0,
                                       L"Button",
                                       LoadStringResource(nullptr, IDS_FIND_ACTION_INTERSECT).c_str(),
                                       buttonStyle,
                                       0,
                                       0,
                                       10,
                                       10,
                                       hwnd,
                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIntersectButtonId)),
                                       instance,
                                       nullptr);
    _subtractButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_ACTION_SUBTRACT).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSubtractButtonId)), instance, nullptr);
    _cancelButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kCancelButtonId)), instance, nullptr);
    _openButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(kOpenButtonId)), instance, nullptr);
    _parentButton = CreateWindowExW(0, L"Button", LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT).c_str(), buttonStyle, 0, 0, 10, 10, hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kParentButtonId)), instance, nullptr);
    _statusText = CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_PATHELLIPSIS, 0, 0, 10, 10, hwnd,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(kStatusTextId)), instance, nullptr);
    _resultsList = CreateWindowExW(WS_EX_CLIENTEDGE,
                                   WC_LISTVIEWW,
                                   L"",
                                   WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS,
                                   0,
                                   0,
                                   10,
                                   10,
                                   hwnd,
                                   reinterpret_cast<HMENU>(static_cast<INT_PTR>(kResultsListId)),
                                   instance,
                                   nullptr);

    if (_resultsList)
    {
        ListView_SetExtendedListViewStyleEx(
            _resultsList, 0, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_GRIDLINES);
    }

    ThemedControls::EnableOwnerDrawButton(hwnd, kFindButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kAppendButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kIntersectButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kSubtractButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kCancelButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kOpenButtonId);
    ThemedControls::EnableOwnerDrawButton(hwnd, kParentButtonId);

    EnsureFonts();
    PopulateModeCombos();
    PopulateHistoryCombos();
    PopulateFromSettings();
    EnsureColumns();
    ApplyTheme();
    UpdateOptionDependencies();
    UpdateActionButtons();
    SetStatusText(LoadStringResource(nullptr, IDS_FIND_STATUS_READY));
    Layout();
}

void FindFilesWindow::EnsureFonts() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(_hWnd.get());
    _uiFont        = CreateMenuFontForDpi(dpi);
    const HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

    const HWND controls[] = {
        _rootLabel,            _rootCombo,           _nameLabel,           _nameCombo,         _nameModeCombo,
        _contentLabel,         _contentCombo,        _contentModeCombo,    _recursiveCheck,    _includeFilesCheck,
        _includeDirectoriesCheck,
        _followSymlinksCheck,  _matchCaseNameCheck,  _matchCaseContentCheck,
        _preferIndexCheck,     _wantSnippetsCheck,   _findButton,          _appendButton,      _intersectButton,
        _subtractButton,       _cancelButton,        _openButton,          _parentButton,      _statusText,
        _resultsList,
    };

    for (HWND control : controls)
    {
        if (control)
        {
            SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);
        }
    }

    if (_resultsList)
    {
        if (const HWND header = ListView_GetHeader(_resultsList))
        {
            SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);
        }
    }
}

void FindFilesWindow::PopulateModeCombos() noexcept
{
    if (_nameModeCombo)
    {
        ComboBox_ResetContent(_nameModeCombo);
        ComboBox_AddString(_nameModeCombo, LoadStringResource(nullptr, IDS_FIND_NAME_MODE_WILDCARD).c_str());
        ComboBox_AddString(_nameModeCombo, LoadStringResource(nullptr, IDS_FIND_NAME_MODE_LITERAL).c_str());
        ComboBox_AddString(_nameModeCombo, LoadStringResource(nullptr, IDS_FIND_NAME_MODE_REGEX).c_str());
    }

    if (_contentModeCombo)
    {
        ComboBox_ResetContent(_contentModeCombo);
        ComboBox_AddString(_contentModeCombo, LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_DISABLED).c_str());
        ComboBox_AddString(_contentModeCombo, LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_LITERAL).c_str());
        ComboBox_AddString(_contentModeCombo, LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_REGEX).c_str());
    }
}

void FindFilesWindow::PopulateHistoryCombos() noexcept
{
    const auto settings = _settings && _settings->search.has_value() ? _settings->search.value() : Common::Settings::SearchDialogSettings{};
    const auto loadHistory = [](HWND combo, const std::vector<std::wstring>& values) noexcept
    {
        if (! combo)
        {
            return;
        }

        const std::wstring currentText = GetWindowTextString(combo);
        ComboBox_ResetContent(combo);
        for (const auto& value : values)
        {
            ComboBox_AddString(combo, value.c_str());
        }

        if (! currentText.empty())
        {
            SetWindowTextW(combo, currentText.c_str());
        }
    };

    loadHistory(_rootCombo, settings.recentRoots);
    loadHistory(_nameCombo, settings.recentNamePatterns);
    loadHistory(_contentCombo, settings.recentContentPatterns);
}

void FindFilesWindow::PopulateFromSettings() noexcept
{
    const Common::Settings::SearchDialogSettings defaults{};
    const auto settings = _settings && _settings->search.has_value() ? _settings->search.value() : defaults;

    const std::wstring initialRoot = settings.lastRoot.empty() ? _context.rootPluginPath.native() : settings.lastRoot;
    SetWindowTextW(_rootCombo, initialRoot.c_str());
    SetWindowTextW(_nameCombo, settings.lastNamePattern.c_str());
    SetWindowTextW(_contentCombo, settings.lastContentPattern.c_str());
    ComboBox_SetCurSel(_nameModeCombo, static_cast<int>(ToComboIndex(settings.nameMode)));
    ComboBox_SetCurSel(_contentModeCombo, static_cast<int>(ToComboIndex(settings.contentMode)));
    Button_SetCheck(_recursiveCheck, settings.recursive ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_includeFilesCheck, settings.includeFiles ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_includeDirectoriesCheck, settings.includeDirectories ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_followSymlinksCheck, settings.followSymlinks ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_matchCaseNameCheck, settings.matchCaseName ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_matchCaseContentCheck, settings.matchCaseContent ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_preferIndexCheck, settings.preferIndex ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_wantSnippetsCheck, settings.wantSnippets ? BST_CHECKED : BST_UNCHECKED);
}

void FindFilesWindow::ApplyTheme() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());

    const HWND combos[] = {_rootCombo, _nameCombo, _nameModeCombo, _contentCombo, _contentModeCombo};
    for (HWND combo : combos)
    {
        if (combo)
        {
            ThemedControls::ApplyThemeToComboBox(combo, _theme);
            ThemedControls::ApplyThemeToComboBoxDropDown(combo, _theme);
        }
    }

    if (_resultsList)
    {
        ThemedControls::ApplyThemeToListView(_resultsList, _theme);
        ThemedControls::EnsureListViewHeaderThemed(_resultsList, _theme);
    }

    InvalidateRect(_hWnd.get(), nullptr, TRUE);
    RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
}

void FindFilesWindow::EnsureColumns() noexcept
{
    if (! _resultsList || (ListView_GetHeader(_resultsList) != nullptr && ListView_GetColumnWidth(_resultsList, 0) > 0))
    {
        return;
    }

    struct ColumnDefinition
    {
        int index;
        UINT textId;
        int widthDip;
    };

    constexpr ColumnDefinition columns[] = {
        {kColumnName, IDS_FIND_COLUMN_NAME, 220},
        {kColumnPath, IDS_FIND_COLUMN_PATH, 360},
        {kColumnSize, IDS_FIND_COLUMN_SIZE, 110},
        {kColumnModified, IDS_FIND_COLUMN_MODIFIED, 150},
        {kColumnAttributes, IDS_FIND_COLUMN_ATTR, 80},
        {kColumnSnippet, IDS_FIND_COLUMN_SNIPPET, 260},
    };

    const UINT dpi = _hWnd ? GetDpiForWindow(_hWnd.get()) : 96u;
    for (const auto& column : columns)
    {
        std::wstring text = LoadStringResource(nullptr, column.textId);
        LVCOLUMNW lvc{};
        lvc.mask     = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        lvc.fmt      = column.index == kColumnSize ? LVCFMT_RIGHT : LVCFMT_LEFT;
        lvc.pszText  = text.data();
        lvc.cx       = column.index == kColumnSnippet && ComboBox_GetCurSel(_contentModeCombo) <= 0
                           ? 0
                           : ThemedControls::ScaleDip(dpi, column.widthDip);
        lvc.iSubItem = column.index;
        ListView_InsertColumn(_resultsList, column.index, &lvc);
    }
}

void FindFilesWindow::UpdateOptionDependencies() noexcept
{
    const bool contentEnabled = ComboBox_GetCurSel(_contentModeCombo) > 0;
    if (_contentCombo)
    {
        EnableWindow(_contentCombo, TRUE);
    }

    if (_matchCaseContentCheck)
    {
        EnableWindow(_matchCaseContentCheck, contentEnabled ? TRUE : FALSE);
        if (! contentEnabled)
        {
            Button_SetCheck(_matchCaseContentCheck, BST_UNCHECKED);
        }
    }

    if (_wantSnippetsCheck)
    {
        EnableWindow(_wantSnippetsCheck, contentEnabled ? TRUE : FALSE);
        if (! contentEnabled)
        {
            Button_SetCheck(_wantSnippetsCheck, BST_UNCHECKED);
        }
    }

    if (contentEnabled && _includeFilesCheck && Button_GetCheck(_includeFilesCheck) != BST_CHECKED)
    {
        Button_SetCheck(_includeFilesCheck, BST_CHECKED);
    }

    if (_resultsList)
    {
        ListView_SetColumnWidth(_resultsList,
                                kColumnSnippet,
                                contentEnabled ? ThemedControls::ScaleDip(GetDpiForWindow(_hWnd.get()), 260) : 0);
    }
}

void FindFilesWindow::UpdateActionButtons() noexcept
{
    const BOOL active = _session.IsActive() ? TRUE : FALSE;
    const bool hasSelection = GetSelectedResultIndex().has_value();

    if (_findButton)
    {
        EnableWindow(_findButton, ! active);
    }
    if (_appendButton)
    {
        EnableWindow(_appendButton, ! active);
    }
    if (_intersectButton)
    {
        EnableWindow(_intersectButton, ! active);
    }
    if (_subtractButton)
    {
        EnableWindow(_subtractButton, ! active);
    }
    if (_cancelButton)
    {
        EnableWindow(_cancelButton, active);
    }
    if (_openButton)
    {
        EnableWindow(_openButton, (! active && hasSelection) ? TRUE : FALSE);
    }
    if (_parentButton)
    {
        EnableWindow(_parentButton, (! active && hasSelection) ? TRUE : FALSE);
    }
}

void FindFilesWindow::Layout() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(_hWnd.get(), &client))
    {
        return;
    }

    const UINT dpi       = GetDpiForWindow(_hWnd.get());
    const int margin     = ThemedControls::ScaleDip(dpi, 12);
    const int gap        = ThemedControls::ScaleDip(dpi, 8);
    const int labelWidth = ThemedControls::ScaleDip(dpi, 76);
    const int comboH     = ThemedControls::ScaleDip(dpi, 28);
    const int comboDropH = ThemedControls::ScaleDip(dpi, 240);
    const int checkH     = ThemedControls::ScaleDip(dpi, 22);
    const int buttonH    = ThemedControls::ScaleDip(dpi, 30);
    const int buttonW    = ThemedControls::ScaleDip(dpi, 108);
    const int modeW      = ThemedControls::ScaleDip(dpi, 150);
    const int statusH    = ThemedControls::ScaleDip(dpi, 22);

    int y = margin;
    const int contentWidth = static_cast<int>(std::max<LONG>(100L, client.right - client.left - 2 * margin));
    const int comboWidth   = static_cast<int>(std::max(100, contentWidth - labelWidth - gap - modeW - gap));
    const int fullComboW   = static_cast<int>(std::max(100, contentWidth - labelWidth - gap));

    MoveWindow(_rootLabel, margin, y + ThemedControls::ScaleDip(dpi, 6), labelWidth, comboH, TRUE);
    MoveWindow(_rootCombo, margin + labelWidth + gap, y, fullComboW, comboDropH, TRUE);
    y += comboH + gap;

    MoveWindow(_nameLabel, margin, y + ThemedControls::ScaleDip(dpi, 6), labelWidth, comboH, TRUE);
    MoveWindow(_nameCombo, margin + labelWidth + gap, y, comboWidth, comboDropH, TRUE);
    MoveWindow(_nameModeCombo, margin + labelWidth + gap + comboWidth + gap, y, modeW, comboDropH, TRUE);
    y += comboH + gap;

    MoveWindow(_contentLabel, margin, y + ThemedControls::ScaleDip(dpi, 6), labelWidth, comboH, TRUE);
    MoveWindow(_contentCombo, margin + labelWidth + gap, y, comboWidth, comboDropH, TRUE);
    MoveWindow(_contentModeCombo, margin + labelWidth + gap + comboWidth + gap, y, modeW, comboDropH, TRUE);
    y += comboH + gap;

    const int optionColumnW = (contentWidth - 3 * gap) / 4;
    const int optionX0      = margin;
    const int optionX1      = optionX0 + optionColumnW + gap;
    const int optionX2      = optionX1 + optionColumnW + gap;
    const int optionX3      = optionX2 + optionColumnW + gap;

    MoveWindow(_recursiveCheck, optionX0, y, optionColumnW, checkH, TRUE);
    MoveWindow(_includeFilesCheck, optionX1, y, optionColumnW, checkH, TRUE);
    MoveWindow(_includeDirectoriesCheck, optionX2, y, optionColumnW, checkH, TRUE);
    MoveWindow(_followSymlinksCheck, optionX3, y, optionColumnW, checkH, TRUE);
    y += checkH + gap;

    MoveWindow(_matchCaseNameCheck, optionX0, y, optionColumnW, checkH, TRUE);
    MoveWindow(_matchCaseContentCheck, optionX1, y, optionColumnW, checkH, TRUE);
    MoveWindow(_preferIndexCheck, optionX2, y, optionColumnW, checkH, TRUE);
    MoveWindow(_wantSnippetsCheck, optionX3, y, optionColumnW, checkH, TRUE);
    y += checkH + gap;

    int buttonX = margin;
    MoveWindow(_findButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_appendButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_intersectButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_subtractButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_cancelButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_openButton, buttonX, y, buttonW, buttonH, TRUE);
    buttonX += buttonW + gap;
    MoveWindow(_parentButton, buttonX, y, buttonW, buttonH, TRUE);
    y += buttonH + gap;

    MoveWindow(_statusText, margin, y, contentWidth, statusH, TRUE);
    y += statusH + gap;

    MoveWindow(_resultsList,
               margin,
               y,
               contentWidth,
               static_cast<int>(std::max<LONG>(120L, client.bottom - y - margin)),
               TRUE);
}

std::wstring FindFilesWindow::GetComboText(HWND combo) const noexcept
{
    return GetWindowTextString(combo);
}

void FindFilesWindow::PersistUiState(bool updateHistory) noexcept
{
    if (! _settings)
    {
        return;
    }

    Common::Settings::SearchDialogSettings settings = _settings->search.value_or(Common::Settings::SearchDialogSettings{});
    settings.lastRoot                               = GetComboText(_rootCombo);
    settings.lastNamePattern                        = GetComboText(_nameCombo);
    settings.lastContentPattern                     = GetComboText(_contentCombo);
    settings.recursive                              = Button_GetCheck(_recursiveCheck) == BST_CHECKED;
    settings.includeFiles                           = Button_GetCheck(_includeFilesCheck) == BST_CHECKED;
    settings.includeDirectories                     = Button_GetCheck(_includeDirectoriesCheck) == BST_CHECKED;
    settings.followSymlinks                         = Button_GetCheck(_followSymlinksCheck) == BST_CHECKED;
    settings.matchCaseName                          = Button_GetCheck(_matchCaseNameCheck) == BST_CHECKED;
    settings.matchCaseContent                       = Button_GetCheck(_matchCaseContentCheck) == BST_CHECKED;
    settings.preferIndex                            = Button_GetCheck(_preferIndexCheck) == BST_CHECKED;
    settings.wantSnippets                           = Button_GetCheck(_wantSnippetsCheck) == BST_CHECKED;
    settings.nameMode                               = FromNameModeComboIndex(ComboBox_GetCurSel(_nameModeCombo));
    settings.contentMode                            = FromContentModeComboIndex(ComboBox_GetCurSel(_contentModeCombo));
    settings.maxResults                             = 0;

    if (updateHistory)
    {
        UpdateRecentValue(settings.recentRoots, settings.lastRoot);
        UpdateRecentValue(settings.recentNamePatterns, settings.lastNamePattern);
        if (! settings.lastContentPattern.empty())
        {
            UpdateRecentValue(settings.recentContentPatterns, settings.lastContentPattern);
        }
    }

    _settings->search = std::move(settings);
}

std::optional<SearchRequest> FindFilesWindow::BuildSearchRequest() noexcept
{
    if (! _context.fileSystem)
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_NO_FILESYSTEM));
        return std::nullopt;
    }

    std::wstring rootPath       = GetComboText(_rootCombo);
    std::wstring namePattern    = GetComboText(_nameCombo);
    std::wstring contentPattern = GetComboText(_contentCombo);

    if (rootPath.empty())
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_ROOT_REQUIRED));
        return std::nullopt;
    }

    const bool includeFiles = Button_GetCheck(_includeFilesCheck) == BST_CHECKED;
    const bool includeDirs  = Button_GetCheck(_includeDirectoriesCheck) == BST_CHECKED;
    if (! includeFiles && ! includeDirs)
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_TARGETS_REQUIRED));
        return std::nullopt;
    }

    const auto configuredNameMode    = FromNameModeComboIndex(ComboBox_GetCurSel(_nameModeCombo));
    const auto configuredContentMode = FromContentModeComboIndex(ComboBox_GetCurSel(_contentModeCombo));
    const FileSystemSearchNameMode nameMode = namePattern.empty() ? FILESYSTEM_SEARCH_NAME_DISABLED : ToAbiNameMode(configuredNameMode);
    const FileSystemSearchContentMode contentMode =
        contentPattern.empty() || configuredContentMode == Common::Settings::SearchContentMode::Disabled ? FILESYSTEM_SEARCH_CONTENT_DISABLED
                                                                                                          : ToAbiContentMode(configuredContentMode);

    if (configuredContentMode != Common::Settings::SearchContentMode::Disabled && contentPattern.empty())
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_CONTENT_REQUIRED));
        return std::nullopt;
    }

    if (nameMode == FILESYSTEM_SEARCH_NAME_DISABLED && contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_CRITERIA_REQUIRED));
        return std::nullopt;
    }

    SearchRequest request;
    request.context                = _context;
    request.rootPath               = std::move(rootPath);
    request.namePattern            = std::move(namePattern);
    request.contentPattern         = std::move(contentPattern);
    request.nameMode               = nameMode;
    request.contentMode            = contentMode;
    request.maxResults             = 0;
    request.maxContentBytesPerFile = kMaxContentBytes;
    request.maxSnippetCharacters   = Button_GetCheck(_wantSnippetsCheck) == BST_CHECKED ? kMaxSnippetCharacters : 0u;

    if (Button_GetCheck(_recursiveCheck) == BST_CHECKED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_RECURSIVE);
    }
    if (includeFiles || contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_INCLUDE_FILES);
    }
    if (includeDirs)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES);
    }
    if (Button_GetCheck(_followSymlinksCheck) == BST_CHECKED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_FOLLOW_SYMLINKS);
    }
    if (Button_GetCheck(_matchCaseNameCheck) == BST_CHECKED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_MATCH_CASE_NAME);
    }
    if (Button_GetCheck(_matchCaseContentCheck) == BST_CHECKED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_MATCH_CASE_CONTENT);
    }
    if (Button_GetCheck(_preferIndexCheck) == BST_CHECKED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_PREFER_INDEX);
    }
    if (request.maxSnippetCharacters != 0)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_WANT_SNIPPETS);
    }

    return request;
}

void FindFilesWindow::OnClose() noexcept
{
    if (_closeRequested)
    {
        return;
    }

    _closeRequested = true;
    PersistUiState(false);
    _session.Shutdown();
}

LRESULT FindFilesWindow::OnNcDestroy() noexcept
{
    PersistUiState(false);
    if (_settings && _hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kFindFilesWindowId, _hWnd.get());
    }

    const HWND hwnd = _hWnd.get();
    _hWnd.release();
    g_findFilesWindows.erase(std::remove(g_findFilesWindows.begin(), g_findFilesWindows.end(), hwnd), g_findFilesWindows.end());
    if (g_findFilesWindow == this)
    {
        g_findFilesWindow = nullptr;
    }

    delete this;
    return 0;
}

void FindFilesWindow::ClearResults() noexcept
{
    _results.clear();
    _resultIndexByKey.clear();
    if (_resultsList)
    {
        ListView_DeleteAllItems(_resultsList);
    }
}

void FindFilesWindow::RebuildResultsList() noexcept
{
    if (! _resultsList)
    {
        return;
    }

    SendMessageW(_resultsList, WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(_resultsList);

    for (size_t index = 0; index < _results.size(); ++index)
    {
        auto& record = _results[index];
        LVITEMW item{};
        item.mask    = LVIF_TEXT;
        item.iItem   = static_cast<int>(index);
        item.pszText = record.displayName.data();
        const int inserted = ListView_InsertItem(_resultsList, &item);
        if (inserted < 0)
        {
            continue;
        }

        std::wstring sizeText      = FormatFileSize(record.endOfFile);
        std::wstring modifiedText  = FormatFileTimeValue(record.lastWriteTime);
        std::wstring attributeText = FormatAttributes(record.fileAttributes);

        ListView_SetItemText(_resultsList, inserted, kColumnPath, record.relativePath.data());
        ListView_SetItemText(_resultsList, inserted, kColumnSize, sizeText.data());
        ListView_SetItemText(_resultsList, inserted, kColumnModified, modifiedText.data());
        ListView_SetItemText(_resultsList, inserted, kColumnAttributes, attributeText.data());
        ListView_SetItemText(_resultsList, inserted, kColumnSnippet, record.previewText.data());
    }

    SendMessageW(_resultsList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(_resultsList, nullptr, TRUE);
}

void FindFilesWindow::AddOrUpdateVisibleResult(FindResultRecord result) noexcept
{
    if (result.key.empty())
    {
        result.key = MakeResultKey(result.pluginId, result.instanceContext, result.fullPath);
    }

    const auto existingIt = _resultIndexByKey.find(result.key);
    if (existingIt != _resultIndexByKey.end())
    {
        _results[existingIt->second] = std::move(result);
        return;
    }

    const size_t index = _results.size();
    _resultIndexByKey.emplace(result.key, index);
    _results.push_back(std::move(result));
}

void FindFilesWindow::RemoveKeysFromResults(const std::unordered_set<std::wstring>& keys) noexcept
{
    if (keys.empty())
    {
        return;
    }

    std::vector<FindResultRecord> filtered;
    filtered.reserve(_results.size());
    for (auto& record : _results)
    {
        if (! keys.contains(record.key))
        {
            filtered.push_back(std::move(record));
        }
    }

    _results = std::move(filtered);
    _resultIndexByKey.clear();
    for (size_t i = 0; i < _results.size(); ++i)
    {
        _resultIndexByKey.emplace(_results[i].key, i);
    }
}

void FindFilesWindow::KeepOnlyKeysInResults(const std::unordered_set<std::wstring>& keys) noexcept
{
    std::vector<FindResultRecord> filtered;
    filtered.reserve(_results.size());
    for (auto& record : _results)
    {
        if (keys.contains(record.key))
        {
            filtered.push_back(std::move(record));
        }
    }

    _results = std::move(filtered);
    _resultIndexByKey.clear();
    for (size_t i = 0; i < _results.size(); ++i)
    {
        _resultIndexByKey.emplace(_results[i].key, i);
    }
}

std::optional<size_t> FindFilesWindow::GetSelectedResultIndex() const noexcept
{
    if (! _resultsList)
    {
        return std::nullopt;
    }

    int item = ListView_GetNextItem(_resultsList, -1, LVNI_SELECTED);
    if (item < 0)
    {
        item = ListView_GetNextItem(_resultsList, -1, LVNI_FOCUSED);
    }
    if (item < 0 || static_cast<size_t>(item) >= _results.size())
    {
        return std::nullopt;
    }

    return static_cast<size_t>(item);
}

void FindFilesWindow::OpenSelectedResult(bool parentOnly) noexcept
{
    const auto index = GetSelectedResultIndex();
    if (! index.has_value())
    {
        return;
    }

    const FindResultRecord& record = _results[index.value()];
    const bool isDirectory         = (record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const std::filesystem::path fullPath(record.fullPath);

    std::filesystem::path targetFolder;
    std::wstring focusName;
    unsigned int commandId = 0u;

    if (parentOnly)
    {
        targetFolder = fullPath.parent_path();
        focusName    = record.displayName;
    }
    else if (isDirectory)
    {
        targetFolder = fullPath;
    }
    else
    {
        targetFolder = fullPath.parent_path();
        focusName    = record.displayName;
        commandId    = IDM_FOLDERVIEW_CONTEXT_OPEN;
    }

    const std::filesystem::path historyPath =
        NavigationLocation::FormatHistoryPath(record.pluginShortId, record.instanceContext, targetFolder);
    static_cast<void>(g_folderWindow.ExecuteInActivePane(historyPath, focusName, commandId, true));
}

void FindFilesWindow::SetStatusText(std::wstring text) noexcept
{
    _status = std::move(text);
    if (_statusText)
    {
        SetWindowTextW(_statusText, _status.c_str());
    }
}

std::wstring FindFilesWindow::BuildWarningSummary(uint32_t warningFlags) const noexcept
{
    std::vector<std::wstring> warnings;
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_NO_INDEX));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_NO_CONTENT));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_ACCESS_DENIED));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_OVERFLOW) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_OVERFLOW));
    }

    std::wstring result;
    for (size_t i = 0; i < warnings.size(); ++i)
    {
        if (i != 0)
        {
            result.append(L", ");
        }
        result.append(warnings[i]);
    }
    return result;
}

UINT FindFilesWindow::BackendStringId(FileSystemSearchBackend backend) const noexcept
{
    switch (backend)
    {
        case FILESYSTEM_SEARCH_BACKEND_UNKNOWN: return IDS_FIND_BACKEND_UNKNOWN;
        case FILESYSTEM_SEARCH_BACKEND_SCAN: return IDS_FIND_BACKEND_SCAN;
        case FILESYSTEM_SEARCH_BACKEND_INDEX: return IDS_FIND_BACKEND_INDEX;
        case FILESYSTEM_SEARCH_BACKEND_SERVICE: return IDS_FIND_BACKEND_SERVICE;
        default: return IDS_FIND_BACKEND_UNKNOWN;
    }
}

std::wstring FindFilesWindow::BuildStatusText() const noexcept
{
    if (_session.IsActive())
    {
        if (_cancelRequestedUi)
        {
            return LoadStringResource(nullptr, IDS_FIND_STATUS_CANCELLING);
        }

        std::wstring running = FormatStringResource(nullptr,
                                                    IDS_FIND_STATUS_RUNNING_FMT,
                                                    LoadStringResource(nullptr, BackendStringId(_lastBackend)),
                                                    _lastScannedDirectories,
                                                    _lastScannedFiles,
                                                    _lastMatchedEntries);
        const std::wstring warnings = BuildWarningSummary(_lastWarningFlags);
        if (! warnings.empty())
        {
            running.append(L" [");
            running.append(warnings);
            running.push_back(L']');
        }
        return running;
    }

    if (_lastStatusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        return FormatStringResource(nullptr,
                                    IDS_FIND_STATUS_CANCELLED_FMT,
                                    LoadStringResource(nullptr, BackendStringId(_lastBackend)),
                                    _lastMatchedEntries,
                                    _lastScannedFiles);
    }

    if (FAILED(_lastStatusHint))
    {
        return FormatStringResource(nullptr, IDS_FIND_STATUS_FAILED_FMT, static_cast<unsigned long>(_lastStatusHint));
    }

    if (_lastBackend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || _lastMatchedEntries != 0 || _lastScannedFiles != 0)
    {
        std::wstring complete = FormatStringResource(nullptr,
                                                     IDS_FIND_STATUS_COMPLETE_FMT,
                                                     LoadStringResource(nullptr, BackendStringId(_lastBackend)),
                                                     _lastMatchedEntries,
                                                     _lastScannedFiles);
        const std::wstring warnings = BuildWarningSummary(_lastWarningFlags);
        if (! warnings.empty())
        {
            complete.append(L" [");
            complete.append(warnings);
            complete.push_back(L']');
        }
        return complete;
    }

    return LoadStringResource(nullptr, IDS_FIND_STATUS_READY);
}

void FindFilesWindow::RefreshStatusText() noexcept
{
    SetStatusText(BuildStatusText());
}

void FindFilesWindow::OnSearchStarted(SearchOperation operation) noexcept
{
    _activeOperation        = operation;
    _cancelRequestedUi      = false;
    _lastBackend            = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    _lastWarningFlags       = FILESYSTEM_SEARCH_WARNING_NONE;
    _lastScannedDirectories = 0;
    _lastScannedFiles       = 0;
    _lastCandidateFiles     = 0;
    _lastMatchedEntries     = 0;
    _lastStatusHint         = S_OK;
    _deferredKeys.clear();

    if (operation == SearchOperation::Find)
    {
        ClearResults();
    }

    RefreshStatusText();
    UpdateActionButtons();
}

bool FindFilesWindow::BeginSearch(SearchOperation operation) noexcept
{
    if (_session.IsActive())
    {
        return false;
    }

    std::optional<SearchRequest> request = BuildSearchRequest();
    if (! request.has_value())
    {
        return false;
    }

    PersistUiState(true);
    PopulateHistoryCombos();
    OnSearchStarted(operation);

    if (! _session.Start(*this, std::move(request.value())))
    {
        _lastStatusHint = E_FAIL;
        RefreshStatusText();
        UpdateActionButtons();
        return false;
    }

    return true;
}

void FindFilesWindow::OnSearchResults(std::unique_ptr<FindSearchResultsPayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        for (auto& record : payload->results)
        {
            _deferredKeys.insert(record.key);
        }
        return;
    }

    for (auto& record : payload->results)
    {
        AddOrUpdateVisibleResult(std::move(record));
    }

    RebuildResultsList();
    UpdateActionButtons();
}

void FindFilesWindow::OnSearchProgress(std::unique_ptr<FindSearchProgressPayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    _lastBackend            = payload->backend;
    _lastWarningFlags       = payload->warningFlags;
    _lastStatusHint         = payload->statusHint;
    _lastScannedDirectories = payload->scannedDirectories;
    _lastScannedFiles       = payload->scannedFiles;
    _lastCandidateFiles     = payload->candidateFiles;
    _lastMatchedEntries     = payload->matchedEntries;
    RefreshStatusText();
}

void FindFilesWindow::ApplyDeferredSetOperation(SearchOperation operation) noexcept
{
    switch (operation)
    {
        case SearchOperation::Find:
        case SearchOperation::Append: break;
        case SearchOperation::Intersect: KeepOnlyKeysInResults(_deferredKeys); break;
        case SearchOperation::Subtract: RemoveKeysFromResults(_deferredKeys); break;
        default: break;
    }

    RebuildResultsList();
}

void FindFilesWindow::OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    if (payload->backend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || _lastBackend == FILESYSTEM_SEARCH_BACKEND_UNKNOWN)
    {
        _lastBackend = payload->backend;
    }
    _lastWarningFlags       = payload->warningFlags;
    _lastScannedDirectories = payload->scannedDirectories;
    _lastScannedFiles       = payload->scannedFiles;
    _lastCandidateFiles     = payload->candidateFiles;
    _lastMatchedEntries     = payload->matchedEntries;
    _lastStatusHint         = payload->hr;
    _cancelRequestedUi      = false;

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        ApplyDeferredSetOperation(_activeOperation);
    }

    _session.NotifyUiSettled();
    RefreshStatusText();
    UpdateActionButtons();
}

#ifdef _DEBUG
bool FindFilesWindow::DebugConfigure(std::wstring rootPath,
                                     std::wstring namePattern,
                                     std::wstring contentPattern,
                                     Common::Settings::SearchNameMode nameMode,
                                     Common::Settings::SearchContentMode contentMode) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    SetWindowTextW(_rootCombo, rootPath.c_str());
    SetWindowTextW(_nameCombo, namePattern.c_str());
    SetWindowTextW(_contentCombo, contentPattern.c_str());
    ComboBox_SetCurSel(_nameModeCombo, static_cast<int>(ToComboIndex(nameMode)));
    ComboBox_SetCurSel(_contentModeCombo, static_cast<int>(ToComboIndex(contentMode)));
    UpdateOptionDependencies();
    PersistUiState(false);
    return true;
}

bool FindFilesWindow::DebugSetOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    Button_SetCheck(_recursiveCheck, recursive ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_includeFilesCheck, includeFiles ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_includeDirectoriesCheck, includeDirectories ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_preferIndexCheck, preferIndex ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(_wantSnippetsCheck, wantSnippets ? BST_CHECKED : BST_UNCHECKED);
    UpdateOptionDependencies();
    PersistUiState(false);
    return true;
}

bool FindFilesWindow::DebugStartSearch(FindFilesDebugOperation operation) noexcept
{
    switch (operation)
    {
        case FindFilesDebugOperation::Find: return BeginSearch(SearchOperation::Find);
        case FindFilesDebugOperation::Append: return BeginSearch(SearchOperation::Append);
        case FindFilesDebugOperation::Intersect: return BeginSearch(SearchOperation::Intersect);
        case FindFilesDebugOperation::Subtract: return BeginSearch(SearchOperation::Subtract);
    }

    return false;
}

bool FindFilesWindow::DebugCancelSearch() noexcept
{
    if (! _session.IsActive())
    {
        return true;
    }

    _cancelRequestedUi = true;
    _session.Cancel();
    RefreshStatusText();
    UpdateActionButtons();
    return true;
}

bool FindFilesWindow::DebugGetSnapshot(FindFilesDebugSnapshot& out) noexcept
{
    out.searchActive = _session.IsActive();
    out.resultCount  = _results.size();
    out.lastStatusHint = _lastStatusHint;
    out.warningFlags = _lastWarningFlags;
    out.statusText   = _status;
    out.fullPaths.clear();
    out.fullPaths.reserve(_results.size());
    for (const auto& record : _results)
    {
        out.fullPaths.push_back(record.fullPath);
    }
    return true;
}

bool FindFilesWindow::DebugWaitForIdle(uint32_t timeoutMs) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(timeoutMs);
    while (std::chrono::steady_clock::now() < deadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        const bool workerIdle = _session.WaitForIdle(0u);
        if (workerIdle && _session.IsUiSettled())
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return _session.WaitForIdle(0u) && _session.IsUiSettled();
}
#endif

LRESULT FindFilesWindow::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
        case WM_CREATE:
            OnCreate(_hWnd.get());
            return 0;
        case WM_SIZE:
            Layout();
            return 0;
        case WM_DPICHANGED:
        {
            if (const RECT* suggested = reinterpret_cast<const RECT*>(lParam))
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            EnsureFonts();
            Layout();
            return 0;
        }
        case WM_CLOSE:
            OnClose();
            break;
        case WM_COMMAND:
        {
            const WORD id   = LOWORD(wParam);
            const WORD code = HIWORD(wParam);
            switch (id)
            {
                case kFindButtonId: static_cast<void>(BeginSearch(SearchOperation::Find)); return 0;
                case kAppendButtonId: static_cast<void>(BeginSearch(SearchOperation::Append)); return 0;
                case kIntersectButtonId: static_cast<void>(BeginSearch(SearchOperation::Intersect)); return 0;
                case kSubtractButtonId: static_cast<void>(BeginSearch(SearchOperation::Subtract)); return 0;
                case kCancelButtonId:
                    _cancelRequestedUi = true;
                    _session.Cancel();
                    RefreshStatusText();
                    UpdateActionButtons();
                    return 0;
                case kOpenButtonId:
                    OpenSelectedResult(false);
                    return 0;
                case kParentButtonId:
                    OpenSelectedResult(true);
                    return 0;
                default: break;
            }

            if (code == BN_CLICKED || code == CBN_EDITCHANGE || code == CBN_SELCHANGE || code == CBN_EDITUPDATE)
            {
                UpdateOptionDependencies();
                PersistUiState(false);
                UpdateActionButtons();
            }
            return 0;
        }
        case WM_NOTIFY:
        {
            const auto* hdr = reinterpret_cast<const NMHDR*>(lParam);
            if (hdr && hdr->idFrom == static_cast<UINT_PTR>(kResultsListId))
            {
                if (hdr->code == LVN_ITEMCHANGED)
                {
                    UpdateActionButtons();
                }
                else if (hdr->code == NM_DBLCLK)
                {
                    OpenSelectedResult(false);
                }
            }
            return 0;
        }
        case WndMsg::kFindSearchResults:
            OnSearchResults(TakeMessagePayload<FindSearchResultsPayload>(lParam));
            return 0;
        case WndMsg::kFindSearchProgress:
            OnSearchProgress(TakeMessagePayload<FindSearchProgressPayload>(lParam));
            return 0;
        case WndMsg::kFindSearchComplete:
            OnSearchComplete(TakeMessagePayload<FindSearchCompletePayload>(lParam));
            return 0;
        case WM_DRAWITEM:
        {
            const auto* draw = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam);
            if (! draw)
            {
                return FALSE;
            }

            switch (static_cast<int>(wParam))
            {
                case kFindButtonId:
                case kAppendButtonId:
                case kIntersectButtonId:
                case kSubtractButtonId:
                case kCancelButtonId:
                case kOpenButtonId:
                case kParentButtonId:
                    ThemedControls::DrawThemedPushButton(*draw, _theme);
                    return TRUE;
                default: break;
            }
            break;
        }
        case WM_ERASEBKGND:
        {
            if (_backgroundBrush)
            {
                RECT rc{};
                GetClientRect(_hWnd.get(), &rc);
                FillRect(reinterpret_cast<HDC>(wParam), &rc, _backgroundBrush.get());
                return 1;
            }
            break;
        }
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORBTN:
        {
            HDC dc = reinterpret_cast<HDC>(wParam);
            SetTextColor(dc, _theme.menu.text);
            SetBkColor(dc, _theme.windowBackground);
            return reinterpret_cast<LRESULT>(_backgroundBrush.get());
        }
        case WM_NCDESTROY:
            SetWindowLongPtrW(_hWnd.get(), GWLP_USERDATA, 0);
            static_cast<void>(DrainPostedPayloadsForWindow(_hWnd.get()));
            return OnNcDestroy();
    }

    return DefWindowProcW(_hWnd.get(), message, wParam, lParam);
}

LRESULT CALLBACK FindFilesWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    FindFilesWindow* self = reinterpret_cast<FindFilesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (message == WM_NCCREATE)
    {
        auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        self         = create ? reinterpret_cast<FindFilesWindow*>(create->lpCreateParams) : nullptr;
        if (! self)
        {
            return FALSE;
        }

        self->_hWnd.reset(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        InitPostedPayloadWindow(hwnd);
        g_findFilesWindow = self;
        g_findFilesWindows.push_back(hwnd);
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    return self->WindowProc(message, wParam, lParam);
}

} // namespace

bool ShowFindFilesWindow(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, FindFilesPaneContext context) noexcept
{
    if (g_findFilesWindow && g_findFilesWindow->GetHwnd() && IsWindow(g_findFilesWindow->GetHwnd()))
    {
        g_findFilesWindow->UpdateTheme(theme);
        g_findFilesWindow->UpdateContext(std::move(context));
        ShowWindow(g_findFilesWindow->GetHwnd(), IsIconic(g_findFilesWindow->GetHwnd()) ? SW_RESTORE : SW_SHOWNORMAL);
        SetForegroundWindow(g_findFilesWindow->GetHwnd());
        return true;
    }

    auto window = std::unique_ptr<FindFilesWindow>(new (std::nothrow) FindFilesWindow(owner, settings, theme, std::move(context)));
    if (! window)
    {
        return false;
    }

    if (! window->Create())
    {
        return false;
    }

    static_cast<void>(window.release());
    return true;
}

void UpdateFindFilesWindowsTheme(const AppTheme& theme) noexcept
{
    const auto windows = g_findFilesWindows;
    for (HWND hwnd : windows)
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            continue;
        }

        if (auto* self = reinterpret_cast<FindFilesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA)))
        {
            self->UpdateTheme(theme);
        }
    }
}

HWND GetFindFilesWindowHandle() noexcept
{
    return g_findFilesWindow && g_findFilesWindow->GetHwnd() && IsWindow(g_findFilesWindow->GetHwnd()) ? g_findFilesWindow->GetHwnd() : nullptr;
}

#ifdef _DEBUG
bool DebugConfigureFindFilesWindow(std::wstring rootPath,
                                   std::wstring namePattern,
                                   std::wstring contentPattern,
                                   Common::Settings::SearchNameMode nameMode,
                                   Common::Settings::SearchContentMode contentMode) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugConfigure(std::move(rootPath),
                                                                 std::move(namePattern),
                                                                 std::move(contentPattern),
                                                                 nameMode,
                                                                 contentMode)
                             : false;
}

bool DebugSetFindFilesWindowOptions(bool recursive,
                                    bool includeFiles,
                                    bool includeDirectories,
                                    bool preferIndex,
                                    bool wantSnippets) noexcept
{
    return g_findFilesWindow
               ? g_findFilesWindow->DebugSetOptions(recursive, includeFiles, includeDirectories, preferIndex, wantSnippets)
               : false;
}

bool DebugStartFindFilesWindowSearch(FindFilesDebugOperation operation) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugStartSearch(operation) : false;
}

bool DebugCancelFindFilesWindowSearch() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugCancelSearch() : false;
}

bool DebugGetFindFilesWindowSnapshot(FindFilesDebugSnapshot& out) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugGetSnapshot(out) : false;
}

bool DebugWaitForFindFilesWindowIdle(uint32_t timeoutMs) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugWaitForIdle(timeoutMs) : true;
}
#endif

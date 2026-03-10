#include "FileSystem.Internal.h"
#include "LocalSearchIndexCore.h"
#include "SearchServiceBroker.h"
#include "SearchTextHelpers.h"

#include <algorithm>
#include <limits>
#include <regex>
#include <string_view>
#include <unordered_set>
#include <vector>

using namespace FileSystemInternal;

namespace
{
constexpr uint64_t kDefaultSearchContentBytesPerFile = SearchTextHelpers::kDefaultContentBytesPerFile;
constexpr uint32_t kDefaultSearchSnippetChars        = SearchTextHelpers::kDefaultSnippetCharacters;
constexpr uint64_t kProgressIntervalItems            = 128u;
constexpr ULONGLONG kProgressIntervalMs              = 200u;
constexpr HRESULT kCancelledHr                       = HRESULT_FROM_WIN32(ERROR_CANCELLED);
constexpr HRESULT kFileTooLargeHr                    = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
constexpr HRESULT kAccessDeniedHr                    = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);

struct SearchEntryMetadata final
{
    std::wstring fullPath;
    std::wstring relativePath;
    std::wstring displayName;
    unsigned long fileAttributes = 0;
    __int64 creationTime         = 0;
    __int64 lastAccessTime       = 0;
    __int64 lastWriteTime        = 0;
    __int64 changeTime           = 0;
    __int64 endOfFile            = 0;
    __int64 allocationSize       = 0;
};

struct SearchContentResult final
{
    bool matched              = false;
    uint64_t byteOffset       = 0;
    uint32_t byteLength       = 0;
    std::wstring previewText;
};

struct SearchBackendSelection final
{
    FileSystemSearchBackend backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
    uint32_t warningFlags           = FILESYSTEM_SEARCH_WARNING_NONE;
};

struct SearchRuntime final
{
    SearchRuntime()                           = default;
    SearchRuntime(const SearchRuntime&)       = delete;
    SearchRuntime(SearchRuntime&&)            = delete;
    SearchRuntime& operator=(const SearchRuntime&) = delete;
    SearchRuntime& operator=(SearchRuntime&&) = delete;

    FileSystem* fileSystem                            = nullptr;
    const FileSystemSearchQuery* query                = nullptr;
    IFileSystemSearchCallback* callback               = nullptr;
    void* cookie                                      = nullptr;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    std::unique_ptr<std::wregex> nameRegex;
    std::unique_ptr<std::wregex> contentRegex;
    std::unordered_set<std::wstring> queuedDirectories;
    bool includeFiles                                 = false;
    bool includeDirectories                           = false;
    bool recursive                                    = false;
    bool followSymlinks                               = false;
    bool wantSnippets                                 = false;
    bool matchCaseName                                = false;
    bool matchCaseContent                             = false;
    uint64_t maxResults                               = 0;
    uint64_t maxContentBytesPerFile                   = kDefaultSearchContentBytesPerFile;
    uint32_t maxSnippetCharacters                     = kDefaultSearchSnippetChars;
    uint64_t scannedDirectories                       = 0;
    uint64_t scannedFiles                             = 0;
    uint64_t candidateFiles                           = 0;
    uint64_t matchedEntries                           = 0;
    FileSystemSearchBackend backend                   = FILESYSTEM_SEARCH_BACKEND_SCAN;
    uint32_t warningFlags                             = FILESYSTEM_SEARCH_WARNING_NONE;
    ULONGLONG lastProgressTick                        = 0;
    uint64_t lastProgressItems                        = 0;
    FileSystemSearchPhase lastProgressPhase           = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    bool hasReportedProgress                          = false;
    bool stopRequested                                = false;
    bool usingIndexedEnumeration                      = false;
};

[[nodiscard]] const wchar_t* SearchBackendToString(FileSystemSearchBackend backend) noexcept
{
    switch (backend)
    {
        case FILESYSTEM_SEARCH_BACKEND_SCAN: return L"scan";
        case FILESYSTEM_SEARCH_BACKEND_INDEX: return L"local-index";
        case FILESYSTEM_SEARCH_BACKEND_SERVICE: return L"service";
        case FILESYSTEM_SEARCH_BACKEND_UNKNOWN:
        default: return L"unknown";
    }
}

[[nodiscard]] const wchar_t* BackendPreferenceToString(FileSystemSearchBackendPreference preference) noexcept
{
    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto: return L"auto";
        case FileSystemSearchBackendPreference::Service: return L"service";
        case FileSystemSearchBackendPreference::LocalIndex: return L"local-index";
        case FileSystemSearchBackendPreference::Scan:
        default: return L"scan";
    }
}

[[nodiscard]] unsigned long ByteCountOfString(std::wstring_view text) noexcept
{
    const size_t bytes = text.size() * sizeof(wchar_t);
    if (bytes > (std::numeric_limits<unsigned long>::max)())
    {
        return (std::numeric_limits<unsigned long>::max)();
    }

    return static_cast<unsigned long>(bytes);
}

[[nodiscard]] std::wstring ToCaseFolded(std::wstring_view text) noexcept
{
    std::wstring result(text);
    if (! result.empty())
    {
        static_cast<void>(::CharLowerBuffW(result.data(), static_cast<DWORD>(result.size())));
    }

    return result;
}

[[nodiscard]] std::wstring NormalizeSearchPath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    while (normalized.size() > 1u && (normalized.back() == L'\\' || normalized.back() == L'/'))
    {
        const bool keepDriveRoot = normalized.size() == 3u && normalized[1] == L':' && (normalized[2] == L'\\' || normalized[2] == L'/');
        const bool keepShareRoot = normalized.size() == 2u && normalized[0] == L'\\' && normalized[1] == L'\\';
        const bool keepExtendedDriveRoot =
            normalized.size() == 7u && normalized.rfind(L"\\\\?\\", 0) == 0 && normalized[5] == L':' && (normalized[6] == L'\\' || normalized[6] == L'/');
        if (keepDriveRoot || keepShareRoot || keepExtendedDriveRoot)
        {
            break;
        }

        normalized.pop_back();
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizeVisitKey(std::wstring_view path) noexcept
{
    std::wstring key = NormalizeSearchPath(path);
    if (! key.empty())
    {
        static_cast<void>(::CharLowerBuffW(key.data(), static_cast<DWORD>(key.size())));
    }

    return key;
}

[[nodiscard]] std::wstring BuildRelativeSearchPath(std::wstring_view rootPath, std::wstring_view fullPath) noexcept
{
    const std::wstring normalizedRoot = NormalizeSearchPath(rootPath);
    const std::wstring normalizedFull = NormalizeSearchPath(fullPath);
    const std::wstring foldedRoot     = ToCaseFolded(normalizedRoot);
    const std::wstring foldedFull     = ToCaseFolded(normalizedFull);

    if (foldedFull.size() >= foldedRoot.size() && foldedFull.compare(0u, foldedRoot.size(), foldedRoot) == 0)
    {
        std::wstring_view remainder(normalizedFull);
        remainder.remove_prefix(normalizedRoot.size());
        while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
        {
            remainder.remove_prefix(1u);
        }

        if (! remainder.empty())
        {
            return std::wstring(remainder);
        }
    }

    const std::wstring leaf = std::wstring(GetPathLeaf(normalizedFull));
    return leaf.empty() ? normalizedFull : leaf;
}

[[nodiscard]] SearchBackendSelection SelectSearchBackend(FileSystemSearchBackendPreference preference,
                                                         FileSystemSearchFlags flags,
                                                         const LocalSearchIndexCore::SupportInfo& support,
                                                         bool serviceAvailable) noexcept
{
    SearchBackendSelection selection{};
    const bool preferIndexHint = (flags & FILESYSTEM_SEARCH_PREFER_INDEX) != 0;

    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto:
            if (serviceAvailable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SERVICE;
            }
            else if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else if (preferIndexHint)
            {
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::Service:
            if (serviceAvailable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SERVICE;
            }
            else if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::LocalIndex:
            if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else
            {
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::Scan:
            selection.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
            break;
    }

    return selection;
}

[[nodiscard]] bool WildcardMatchCaseSensitive(std::wstring_view text, std::wstring_view pattern) noexcept
{
    size_t textPos    = 0;
    size_t patternPos = 0;
    size_t starPos    = std::wstring_view::npos;
    size_t matchPos   = 0;

    while (textPos < text.size())
    {
        if (patternPos < pattern.size() && (pattern[patternPos] == L'?' || pattern[patternPos] == text[textPos]))
        {
            ++patternPos;
            ++textPos;
            continue;
        }

        if (patternPos < pattern.size() && pattern[patternPos] == L'*')
        {
            starPos  = patternPos++;
            matchPos = textPos;
            continue;
        }

        if (starPos != std::wstring_view::npos)
        {
            patternPos = starPos + 1;
            textPos    = ++matchPos;
            continue;
        }

        return false;
    }

    while (patternPos < pattern.size() && pattern[patternPos] == L'*')
    {
        ++patternPos;
    }

    return patternPos == pattern.size();
}

[[nodiscard]] bool WildcardMatch(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    if (caseSensitive)
    {
        return WildcardMatchCaseSensitive(text, pattern);
    }

    return WildcardMatchCaseSensitive(ToCaseFolded(text), ToCaseFolded(pattern));
}

[[nodiscard]] bool FindLiteral(std::wstring_view haystack, std::wstring_view needle, bool caseSensitive, size_t& outPosition) noexcept
{
    outPosition = std::wstring_view::npos;

    if (needle.empty())
    {
        outPosition = 0;
        return true;
    }

    if (caseSensitive)
    {
        outPosition = haystack.find(needle);
        return outPosition != std::wstring_view::npos;
    }

    const std::wstring foldedHaystack = ToCaseFolded(haystack);
    const std::wstring foldedNeedle    = ToCaseFolded(needle);
    outPosition                        = foldedHaystack.find(foldedNeedle);
    return outPosition != std::wstring_view::npos;
}

HRESULT CheckSearchCancelled(SearchRuntime& runtime) noexcept
{
    BOOL cancel            = FALSE;
    const HRESULT cancelHr = runtime.callback->FileSystemSearchShouldCancel(&cancel, runtime.cookie);
    if (FAILED(cancelHr))
    {
        return cancelHr;
    }

    return cancel ? kCancelledHr : S_OK;
}

HRESULT STDMETHODCALLTYPE SearchReadCancelThunk(void* cookie) noexcept
{
    if (cookie == nullptr)
    {
        return E_POINTER;
    }

    return CheckSearchCancelled(*static_cast<SearchRuntime*>(cookie));
}

[[nodiscard]] bool IsCallbackCancellationResult(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == kCancelledHr;
}

HRESULT ReportSearchProgress(SearchRuntime& runtime,
                             FileSystemSearchPhase phase,
                             const std::wstring* currentPath,
                             HRESULT statusHint,
                             bool force) noexcept
{
    const uint64_t processedItems = runtime.scannedDirectories + runtime.scannedFiles;
    const ULONGLONG now           = ::GetTickCount64();

    if (! force && runtime.hasReportedProgress && phase == runtime.lastProgressPhase &&
        (processedItems - runtime.lastProgressItems) < kProgressIntervalItems &&
        (now - runtime.lastProgressTick) < kProgressIntervalMs)
    {
        return S_OK;
    }

    FileSystemSearchProgress progress{};
    progress.sizeBytes          = sizeof(FileSystemSearchProgress);
    progress.phase              = phase;
    progress.backend            = runtime.backend;
    progress.warningFlags       = runtime.warningFlags;
    progress.statusHint         = statusHint;
    progress.scannedDirectories = runtime.scannedDirectories;
    progress.scannedFiles       = runtime.scannedFiles;
    progress.candidateFiles     = runtime.candidateFiles;
    progress.matchedEntries     = runtime.matchedEntries;
    progress.currentPath        = currentPath ? currentPath->c_str() : nullptr;
    progress.currentPathSize    = currentPath ? ByteCountOfString(*currentPath) : 0u;

    const HRESULT hr = runtime.callback->FileSystemSearchProgress(&progress, runtime.cookie);
    if (FAILED(hr))
    {
        return IsCallbackCancellationResult(hr) ? kCancelledHr : hr;
    }

    runtime.hasReportedProgress = true;
    runtime.lastProgressPhase   = phase;
    runtime.lastProgressTick    = now;
    runtime.lastProgressItems   = processedItems;
    return S_OK;
}

[[nodiscard]] bool MatchNamePattern(const SearchRuntime& runtime, const std::wstring& displayName) noexcept
{
    switch (runtime.query->nameMode)
    {
        case FILESYSTEM_SEARCH_NAME_DISABLED: return true;
        case FILESYSTEM_SEARCH_NAME_WILDCARD: return WildcardMatch(displayName, runtime.namePattern, runtime.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_LITERAL:
        {
            size_t position = std::wstring::npos;
            return FindLiteral(displayName, runtime.namePattern, runtime.matchCaseName, position);
        }
        case FILESYSTEM_SEARCH_NAME_REGEX: return runtime.nameRegex && std::regex_search(displayName, *runtime.nameRegex);
    }

    return false;
}

HRESULT MatchFileContent(SearchRuntime& runtime, const std::wstring& fullPath, SearchContentResult& result) noexcept
{
    result = {};

    PathInfo pathInfo = MakePathInfo(fullPath);
    wil::com_ptr<IFileReader> reader;
    HRESULT hr = runtime.fileSystem->CreateFileReader(pathInfo.extended.c_str(), reader.put());
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    ++runtime.candidateFiles;

    hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN, &fullPath, S_OK, false);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckSearchCancelled(runtime);
    if (FAILED(hr))
    {
        return hr;
    }

    SearchTextHelpers::TextSearchPattern pattern{};
    pattern.mode                   = runtime.query->contentMode;
    pattern.pattern                = runtime.contentPattern;
    pattern.compiledRegex          = runtime.contentRegex.get();
    pattern.caseSensitive          = runtime.matchCaseContent;
    pattern.literalChunkCharacters = SearchTextHelpers::kDefaultLiteralChunkChars;

    SearchTextHelpers::TextSearchResult helperResult{};
    hr = SearchTextHelpers::SearchFileReaderText(reader.get(),
                                                 pattern,
                                                 runtime.maxContentBytesPerFile,
                                                 0u,
                                                 runtime.maxSnippetCharacters,
                                                 runtime.wantSnippets,
                                                 &SearchReadCancelThunk,
                                                 &runtime,
                                                 helperResult);
    if (hr == kFileTooLargeHr)
    {
        return S_OK;
    }
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    result.matched     = helperResult.matched;
    result.byteOffset  = helperResult.matchOffset;
    result.byteLength  = helperResult.matchLength;
    result.previewText = std::move(helperResult.previewText);
    return S_OK;
}

HRESULT EmitSearchMatch(SearchRuntime& runtime,
                        const SearchEntryMetadata& entry,
                        uint32_t matchedBy,
                        const SearchContentResult& contentResult) noexcept
{
    FileSystemSearchMatch match{};
    match.sizeBytes              = sizeof(FileSystemSearchMatch);
    match.fullPath               = entry.fullPath.c_str();
    match.fullPathSize           = ByteCountOfString(entry.fullPath);
    match.relativePath           = entry.relativePath.c_str();
    match.relativePathSize       = ByteCountOfString(entry.relativePath);
    match.displayName            = entry.displayName.c_str();
    match.displayNameSize        = ByteCountOfString(entry.displayName);
    match.previewText            = contentResult.previewText.empty() ? nullptr : contentResult.previewText.c_str();
    match.previewTextSize        = ByteCountOfString(contentResult.previewText);
    match.fileAttributes         = entry.fileAttributes;
    match.creationTime           = entry.creationTime;
    match.lastAccessTime         = entry.lastAccessTime;
    match.lastWriteTime          = entry.lastWriteTime;
    match.changeTime             = entry.changeTime;
    match.endOfFile              = entry.endOfFile;
    match.allocationSize         = entry.allocationSize;
    match.matchedBy              = matchedBy;
    match.contentMatchByteOffset = contentResult.byteOffset;
    match.contentMatchByteLength = contentResult.byteLength;

    const HRESULT hr = runtime.callback->FileSystemSearchMatch(&match, runtime.cookie);
    if (FAILED(hr))
    {
        return IsCallbackCancellationResult(hr) ? kCancelledHr : hr;
    }

    ++runtime.matchedEntries;
    if (runtime.maxResults != 0 && runtime.matchedEntries >= runtime.maxResults)
    {
        runtime.stopRequested = true;
    }

    return S_OK;
}

HRESULT EvaluateEntry(SearchRuntime& runtime, const SearchEntryMetadata& entry) noexcept
{
    const bool isDirectory = (entry.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if ((isDirectory && ! runtime.includeDirectories) || (! isDirectory && ! runtime.includeFiles))
    {
        return S_OK;
    }

    if (! isDirectory && ! runtime.usingIndexedEnumeration)
    {
        ++runtime.scannedFiles;
    }

    const bool nameMatched = MatchNamePattern(runtime, entry.displayName);
    if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && ! nameMatched)
    {
        return S_OK;
    }

    SearchContentResult contentResult{};
    bool contentMatched = (runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED);
    if (! contentMatched)
    {
        if (isDirectory)
        {
            return S_OK;
        }

        const HRESULT hr = MatchFileContent(runtime, entry.fullPath, contentResult);
        if (FAILED(hr))
        {
            return hr;
        }

        contentMatched = contentResult.matched;
    }

    if (! contentMatched)
    {
        return S_OK;
    }

    uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
    if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED)
    {
        matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_NAME;
    }
    if (runtime.query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT;
    }

    return EmitSearchMatch(runtime, entry, matchedBy, contentResult);
}

void PopulateSinglePathMetadata(FileSystem& fileSystem,
                                std::wstring_view fullPath,
                                unsigned long fileAttributes,
                                SearchEntryMetadata& outMetadata) noexcept
{
    outMetadata                = {};
    outMetadata.fullPath       = NormalizeSearchPath(fullPath);
    outMetadata.relativePath   = std::wstring(GetPathLeaf(outMetadata.fullPath));
    if (outMetadata.relativePath.empty())
    {
        outMetadata.relativePath = outMetadata.fullPath;
    }
    outMetadata.displayName    = outMetadata.relativePath;
    outMetadata.fileAttributes = fileAttributes;

    FileSystemBasicInformation info{};
    info.sizeBytes = sizeof(FileSystemBasicInformation);
    if (SUCCEEDED(fileSystem.GetFileBasicInformation(outMetadata.fullPath.c_str(), &info)))
    {
        outMetadata.creationTime   = info.creationTime;
        outMetadata.lastAccessTime = info.lastAccessTime;
        outMetadata.lastWriteTime  = info.lastWriteTime;
    }

    PathInfo pathInfo = MakePathInfo(outMetadata.fullPath);
    wil::com_ptr<IFileReader> reader;
    if (SUCCEEDED(fileSystem.CreateFileReader(pathInfo.extended.c_str(), reader.put())))
    {
        uint64_t sizeBytes = 0u;
        if (SUCCEEDED(reader->GetSize(&sizeBytes)))
        {
            outMetadata.endOfFile      = static_cast<__int64>(sizeBytes);
            outMetadata.allocationSize = static_cast<__int64>(sizeBytes);
        }
    }
}

HRESULT SearchIndexedTree(SearchRuntime& runtime, LocalSearchIndexCore::Repository& repository) noexcept
{
    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = runtime.rootPath;
    plan.namePattern        = runtime.namePattern;
    plan.nameMode           = runtime.query->nameMode;
    plan.compiledNameRegex  = runtime.nameRegex.get();
    plan.matchCaseName      = runtime.matchCaseName;
    plan.recursive          = runtime.recursive;
    plan.includeFiles       = runtime.includeFiles;
    plan.includeDirectories = runtime.includeDirectories;
    plan.maxResults         = runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? runtime.maxResults : 0u;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};

    HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, &runtime.rootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = repository.Query(plan, &SearchReadCancelThunk, &runtime, candidates, &stats);
    if (FAILED(hr))
    {
        return hr;
    }

    runtime.usingIndexedEnumeration = true;
    runtime.scannedDirectories      = stats.directoryCount;
    runtime.scannedFiles            = stats.fileCount;

    hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, &runtime.rootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const auto& candidate : candidates)
    {
        if (runtime.stopRequested)
        {
            break;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        SearchEntryMetadata metadata{};
        PopulateSinglePathMetadata(*runtime.fileSystem, candidate.fullPath, candidate.fileAttributes, metadata);
        metadata.relativePath = BuildRelativeSearchPath(runtime.rootPath, metadata.fullPath);
        metadata.displayName  = std::wstring(GetPathLeaf(metadata.fullPath));
        if (metadata.displayName.empty())
        {
            metadata.displayName = metadata.relativePath;
        }

        hr = EvaluateEntry(runtime, metadata);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE SearchServiceProgressThunk(const SearchServiceBroker::QueryProgress* progress, void* cookie) noexcept
{
    if (progress == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    SearchRuntime& runtime     = *static_cast<SearchRuntime*>(cookie);
    runtime.scannedDirectories = progress->scannedDirectories;
    runtime.scannedFiles       = progress->scannedFiles;
    runtime.candidateFiles     = progress->candidateFiles;
    runtime.matchedEntries     = progress->matchedEntries;
    runtime.warningFlags      |= progress->warningFlags;

    return ReportSearchProgress(runtime,
                                progress->phase,
                                progress->currentPath.empty() ? nullptr : &progress->currentPath,
                                progress->statusHint,
                                true);
}

HRESULT SearchServiceTree(SearchRuntime& runtime) noexcept
{
    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = runtime.rootPath;
    request.namePattern        = runtime.namePattern;
    request.nameMode           = runtime.query->nameMode;
    request.flags              = runtime.query->flags;
    request.recursive          = runtime.recursive;
    request.includeFiles       = runtime.includeFiles;
    request.includeDirectories = runtime.includeDirectories;
    request.matchCaseName      = runtime.matchCaseName;
    request.maxResults         = runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? runtime.maxResults : 0u;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};
    HRESULT hr = SearchServiceBroker::Query(request,
                                            &SearchServiceProgressThunk,
                                            &runtime,
                                            &SearchReadCancelThunk,
                                            &runtime,
                                            candidates,
                                            &stats);
    if (FAILED(hr))
    {
        return hr;
    }

    runtime.usingIndexedEnumeration = true;
    runtime.scannedDirectories      = stats.directoryCount;
    runtime.scannedFiles            = stats.fileCount;
    runtime.candidateFiles          = stats.candidateCount;

    for (const auto& candidate : candidates)
    {
        if (runtime.stopRequested)
        {
            break;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        SearchEntryMetadata metadata{};
        PopulateSinglePathMetadata(*runtime.fileSystem, candidate.fullPath, candidate.fileAttributes, metadata);
        metadata.relativePath = BuildRelativeSearchPath(runtime.rootPath, metadata.fullPath);
        metadata.displayName  = std::wstring(GetPathLeaf(metadata.fullPath));
        if (metadata.displayName.empty())
        {
            metadata.displayName = metadata.relativePath;
        }

        hr = EvaluateEntry(runtime, metadata);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

[[nodiscard]] bool IsServiceFallbackCandidate(HRESULT hr) noexcept
{
    switch (hr)
    {
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_NO_DATA):
        case HRESULT_FROM_WIN32(ERROR_PIPE_NOT_CONNECTED):
        case HRESULT_FROM_WIN32(ERROR_PIPE_BUSY):
        case HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT):
        case HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE):
        case RPC_S_PROTOCOL_ERROR:
            return true;
    }

    return false;
}

HRESULT SearchDirectoryTree(SearchRuntime& runtime) noexcept
{
    struct DirectoryFrame final
    {
        std::wstring fullPath;
        std::wstring relativeBase;
    };

    std::vector<DirectoryFrame> stack;
    stack.push_back({runtime.rootPath, std::wstring()});
    runtime.queuedDirectories.insert(NormalizeVisitKey(runtime.rootPath));

    while (! stack.empty() && ! runtime.stopRequested)
    {
        const DirectoryFrame frame = std::move(stack.back());
        stack.pop_back();

        ++runtime.scannedDirectories;

        HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_ENUMERATING, &frame.fullPath, S_OK, false);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        wil::com_ptr<IFilesInformation> information;
        hr = runtime.fileSystem->ReadDirectoryInfo(frame.fullPath.c_str(), information.put());
        if (FAILED(hr))
        {
            if (hr == kAccessDeniedHr && frame.fullPath != runtime.rootPath)
            {
                runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
                continue;
            }

            return hr;
        }

        unsigned long count = 0;
        hr                  = information->GetCount(&count);
        if (FAILED(hr))
        {
            return hr;
        }

        for (unsigned long index = 0; index < count && ! runtime.stopRequested; ++index)
        {
            FileInfo* entry = nullptr;
            hr              = information->Get(index, &entry);
            if (FAILED(hr) || entry == nullptr)
            {
                return FAILED(hr) ? hr : E_FAIL;
            }

            const std::wstring_view name(entry->FileName, entry->FileNameSize / sizeof(wchar_t));
            if (IsDotOrDotDot(name))
            {
                continue;
            }

            SearchEntryMetadata metadata{};
            metadata.displayName    = std::wstring(name);
            metadata.relativePath   = frame.relativeBase.empty() ? metadata.displayName : NormalizeSearchPath(AppendPath(frame.relativeBase, metadata.displayName));
            metadata.fullPath       = NormalizeSearchPath(AppendPath(frame.fullPath, metadata.displayName));
            metadata.fileAttributes = entry->FileAttributes;
            metadata.creationTime   = entry->CreationTime;
            metadata.lastAccessTime = entry->LastAccessTime;
            metadata.lastWriteTime  = entry->LastWriteTime;
            metadata.changeTime     = entry->ChangeTime;
            metadata.endOfFile      = entry->EndOfFile;
            metadata.allocationSize = entry->AllocationSize;

            hr = EvaluateEntry(runtime, metadata);
            if (FAILED(hr))
            {
                return hr;
            }

            const bool isDirectory = (metadata.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            const bool isReparse   = (metadata.fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
            if (runtime.recursive && isDirectory && (! isReparse || runtime.followSymlinks))
            {
                const std::wstring visitKey = NormalizeVisitKey(metadata.fullPath);
                if (runtime.queuedDirectories.insert(visitKey).second)
                {
                    stack.push_back({metadata.fullPath, metadata.relativePath});
                }
            }
        }
    }

    return S_OK;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::Search(const FileSystemSearchQuery* query, IFileSystemSearchCallback* callback, void* cookie) noexcept
{
    if (query == nullptr || callback == nullptr)
    {
        return E_POINTER;
    }

    if (query->sizeBytes != sizeof(FileSystemSearchQuery))
    {
        return E_INVALIDARG;
    }

    if (query->rootPath == nullptr || query->rootPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const bool includeFiles       = (query->flags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0;
    const bool includeDirectories = (query->flags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0;
    if (! includeFiles && ! includeDirectories)
    {
        return E_INVALIDARG;
    }

    if (query->nameMode == FILESYSTEM_SEARCH_NAME_DISABLED && query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        return E_INVALIDARG;
    }

    if (query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && query->namePattern == nullptr)
    {
        return E_INVALIDARG;
    }

    if (query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        if (query->contentPattern == nullptr || ! includeFiles)
        {
            return E_INVALIDARG;
        }
    }

    FileSystemSearchBackendPreference backendPreference = kDefaultSearchBackendPreference;
    std::shared_ptr<LocalSearchIndexCore::Repository> searchIndexRepository;
    {
        std::lock_guard lock(_stateMutex);
        backendPreference    = _searchBackendPreference;
        searchIndexRepository = _searchIndexRepository;
    }

    try
    {
        SearchRuntime runtime{};
        runtime.fileSystem             = this;
        runtime.query                  = query;
        runtime.callback               = callback;
        runtime.cookie                 = cookie;
        runtime.rootPath               = NormalizeSearchPath(MakeAbsolutePath(query->rootPath));
        runtime.namePattern            = query->namePattern ? std::wstring(query->namePattern) : std::wstring();
        runtime.contentPattern         = query->contentPattern ? std::wstring(query->contentPattern) : std::wstring();
        runtime.includeFiles           = includeFiles;
        runtime.includeDirectories     = includeDirectories;
        runtime.recursive              = (query->flags & FILESYSTEM_SEARCH_RECURSIVE) != 0;
        runtime.followSymlinks         = (query->flags & FILESYSTEM_SEARCH_FOLLOW_SYMLINKS) != 0;
        runtime.wantSnippets           = (query->flags & FILESYSTEM_SEARCH_WANT_SNIPPETS) != 0;
        runtime.matchCaseName          = (query->flags & FILESYSTEM_SEARCH_MATCH_CASE_NAME) != 0;
        runtime.matchCaseContent       = (query->flags & FILESYSTEM_SEARCH_MATCH_CASE_CONTENT) != 0;
        runtime.maxResults             = query->maxResults;
        runtime.maxContentBytesPerFile = query->maxContentBytesPerFile != 0 ? query->maxContentBytesPerFile : kDefaultSearchContentBytesPerFile;
        runtime.maxSnippetCharacters   = query->maxSnippetCharacters != 0 ? query->maxSnippetCharacters : kDefaultSearchSnippetChars;

        if (runtime.rootPath.empty())
        {
            runtime.rootPath = NormalizeSearchPath(query->rootPath);
        }

        LocalSearchIndexCore::SupportInfo indexSupport{};
        if (searchIndexRepository)
        {
            const HRESULT probeHr = searchIndexRepository->ProbePath(runtime.rootPath, indexSupport);
            if (FAILED(probeHr))
            {
                indexSupport = {};
            }
        }

        bool serviceAvailable = false;
        if (indexSupport.indexable && backendPreference != FileSystemSearchBackendPreference::Scan)
        {
            SearchServiceBroker::ServiceStatus serviceStatus{};
            serviceAvailable = SUCCEEDED(SearchServiceBroker::GetStatus(serviceStatus));
        }

        const SearchBackendSelection backendSelection = SelectSearchBackend(backendPreference, query->flags, indexSupport, serviceAvailable);
        runtime.backend                               = backendSelection.backend;
        runtime.warningFlags |= backendSelection.warningFlags;

        Debug::Info(L"FileSystem::Search: root='{}' preference='{}' selected='{}' serviceAvailable={} indexable={} fsKind={} flags=0x{:08X} warnings=0x{:08X}",
                    runtime.rootPath,
                    BackendPreferenceToString(backendPreference),
                    SearchBackendToString(runtime.backend),
                    serviceAvailable,
                    indexSupport.indexable,
                    static_cast<uint32_t>(indexSupport.fileSystemKind),
                    static_cast<uint32_t>(query->flags),
                    runtime.warningFlags);

        if (query->nameMode == FILESYSTEM_SEARCH_NAME_REGEX)
        {
            const auto flags = runtime.matchCaseName ? std::regex_constants::ECMAScript
                                                     : (std::regex_constants::ECMAScript | std::regex_constants::icase);
            runtime.nameRegex = std::make_unique<std::wregex>(runtime.namePattern, flags);
        }

        if (query->contentMode == FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX)
        {
            const auto flags = runtime.matchCaseContent ? std::regex_constants::ECMAScript
                                                        : (std::regex_constants::ECMAScript | std::regex_constants::icase);
            runtime.contentRegex = std::make_unique<std::wregex>(runtime.contentPattern, flags);
        }

        HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INITIALIZING, &runtime.rootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        unsigned long rootAttributes = 0;
        hr                           = GetAttributes(runtime.rootPath.c_str(), &rootAttributes);
        if (FAILED(hr))
        {
            return hr;
        }

        if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_SERVICE)
        {
            hr = SearchServiceTree(runtime);
            if (FAILED(hr) && IsServiceFallbackCandidate(hr))
            {
                Debug::Warning(L"FileSystem::Search: service backend failed root='{}' hr=0x{:08X}; falling back.",
                               runtime.rootPath,
                               static_cast<unsigned long>(hr));
                if (searchIndexRepository && indexSupport.indexable)
                {
                    runtime.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
                }
                else
                {
                    runtime.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                    runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
                }

                runtime.scannedDirectories    = 0u;
                runtime.scannedFiles          = 0u;
                runtime.candidateFiles        = 0u;
                runtime.matchedEntries        = 0u;
                runtime.usingIndexedEnumeration = false;
                hr = S_OK;
            }
        }

        if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_INDEX && searchIndexRepository)
        {
            hr = SearchIndexedTree(runtime, *searchIndexRepository);
            if (FAILED(hr) &&
                (hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || hr == HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION)))
            {
                Debug::Warning(L"FileSystem::Search: indexed backend failed root='{}' hr=0x{:08X}; degrading to scan.",
                               runtime.rootPath,
                               static_cast<unsigned long>(hr));
                runtime.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
                runtime.scannedDirectories = 0u;
                runtime.scannedFiles       = 0u;
                runtime.usingIndexedEnumeration = false;
                hr = S_OK;
            }
        }

        if (SUCCEEDED(hr))
        {
            if ((rootAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && runtime.backend == FILESYSTEM_SEARCH_BACKEND_SCAN)
            {
                hr = SearchDirectoryTree(runtime);
            }
            else if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_SCAN)
            {
                SearchEntryMetadata entry{};
                PopulateSinglePathMetadata(*this, runtime.rootPath, rootAttributes, entry);
                hr = EvaluateEntry(runtime, entry);
            }
        }

        const HRESULT finalStatus = FAILED(hr) ? hr : (runtime.stopRequested ? S_FALSE : S_OK);
        const HRESULT progressHr  = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_COMPLETED, nullptr, finalStatus, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        Debug::Info(L"FileSystem::Search: completed root='{}' backend='{}' status=0x{:08X} matched={} candidates={} scannedFiles={} scannedDirs={} warnings=0x{:08X}",
                    runtime.rootPath,
                    SearchBackendToString(runtime.backend),
                    static_cast<unsigned long>(FAILED(hr) ? hr : finalStatus),
                    runtime.matchedEntries,
                    runtime.candidateFiles,
                    runtime.scannedFiles,
                    runtime.scannedDirectories,
                    runtime.warningFlags);

        if (runtime.stopRequested && SUCCEEDED(hr))
        {
            return S_OK;
        }

        return hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::regex_error&)
    {
        return E_INVALIDARG;
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystem: Search failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

#include "SearchFallbackEngine.h"

#include "Framework.h"

#include "Helpers.h"
#include "SearchTextHelpers.h"

#include <algorithm>
#include <limits>
#include <regex>
#include <string_view>
#include <unordered_set>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#pragma warning(pop)

namespace SearchFallbackEngine
{
namespace
{
constexpr uint64_t kProgressIntervalItems = 128u;
constexpr ULONGLONG kProgressIntervalMs   = 200u;
constexpr HRESULT kCancelledHr            = HRESULT_FROM_WIN32(ERROR_CANCELLED);
constexpr HRESULT kFileTooLargeHr         = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
constexpr HRESULT kAccessDeniedHr         = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
constexpr HRESULT kNotSupportedHr         = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);

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
    bool matched        = false;
    uint64_t byteOffset = 0;
    uint32_t byteLength = 0;
    std::wstring previewText;
};

struct SearchRuntime final
{
    SearchRuntime()                                = default;
    SearchRuntime(const SearchRuntime&)            = delete;
    SearchRuntime(SearchRuntime&&)                 = delete;
    SearchRuntime& operator=(const SearchRuntime&) = delete;
    SearchRuntime& operator=(SearchRuntime&&)      = delete;

    wil::com_ptr<IFileSystem> fileSystem;
    wil::com_ptr<IFileSystemIO> fileSystemIo;
    const FileSystemSearchQuery* query  = nullptr;
    IFileSystemSearchCallback* callback = nullptr;
    void* cookie                        = nullptr;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    std::unique_ptr<std::wregex> nameRegex;
    std::unique_ptr<std::wregex> contentRegex;
    std::unordered_set<std::wstring> queuedDirectories;
    bool includeFiles                       = false;
    bool includeDirectories                 = false;
    bool recursive                          = false;
    bool followSymlinks                     = false;
    bool wantSnippets                       = false;
    bool matchCaseName                      = false;
    bool matchCaseContent                   = false;
    uint64_t maxResults                     = 0;
    uint64_t maxContentBytesPerFile         = SearchTextHelpers::kDefaultContentBytesPerFile;
    uint32_t maxSnippetCharacters           = SearchTextHelpers::kDefaultSnippetCharacters;
    uint64_t scannedDirectories             = 0;
    uint64_t scannedFiles                   = 0;
    uint64_t candidateFiles                 = 0;
    uint64_t matchedEntries                 = 0;
    uint32_t warningFlags                   = FILESYSTEM_SEARCH_WARNING_NONE;
    ULONGLONG lastProgressTick              = 0u;
    uint64_t lastProgressItems              = 0u;
    FileSystemSearchPhase lastProgressPhase = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    bool hasReportedProgress                = false;
    bool stopRequested                      = false;
};

[[nodiscard]] unsigned long ByteCountOfString(std::wstring_view text) noexcept
{
    const size_t bytes = text.size() * sizeof(wchar_t);
    if (bytes > (std::numeric_limits<unsigned long>::max)())
    {
        return (std::numeric_limits<unsigned long>::max)();
    }

    return static_cast<unsigned long>(bytes);
}

[[nodiscard]] std::wstring FoldText(std::wstring_view text) noexcept
{
    return OrdinalString::FoldCaseInvariant(text);
}

[[nodiscard]] bool WildcardMatchCaseSensitive(std::wstring_view text, std::wstring_view pattern) noexcept
{
    size_t textPos    = 0u;
    size_t patternPos = 0u;
    size_t starPos    = std::wstring_view::npos;
    size_t matchPos   = 0u;

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
            patternPos = starPos + 1u;
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

    return WildcardMatchCaseSensitive(FoldText(text), FoldText(pattern));
}

[[nodiscard]] bool FindLiteral(std::wstring_view haystack, std::wstring_view needle, bool caseSensitive, size_t& outPosition) noexcept
{
    outPosition = std::wstring_view::npos;

    if (needle.empty())
    {
        outPosition = 0u;
        return true;
    }

    if (caseSensitive)
    {
        outPosition = haystack.find(needle);
        return outPosition != std::wstring_view::npos;
    }

    const std::wstring foldedHaystack = FoldText(haystack);
    const std::wstring foldedNeedle   = FoldText(needle);
    outPosition                       = foldedHaystack.find(foldedNeedle);
    return outPosition != std::wstring_view::npos;
}

[[nodiscard]] bool IsCallbackCancellationResult(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == kCancelledHr;
}

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept
{
    return name == L"." || name == L"..";
}

[[nodiscard]] wchar_t PickPathSeparator(std::wstring_view basePath) noexcept
{
    if (basePath.find(L'\\') != std::wstring_view::npos || (basePath.size() >= 2u && basePath[1] == L':'))
    {
        return L'\\';
    }

    return L'/';
}

[[nodiscard]] std::wstring AppendPath(std::wstring_view basePath, std::wstring_view leafName, wchar_t separator) noexcept
{
    if (basePath.empty())
    {
        return std::wstring(leafName);
    }

    std::wstring result(basePath);
    const wchar_t lastChar = result.back();
    if (lastChar != L'\\' && lastChar != L'/')
    {
        result.push_back(separator);
    }
    result.append(leafName);
    return result;
}

[[nodiscard]] std::wstring AppendPath(std::wstring_view basePath, std::wstring_view leafName) noexcept
{
    return AppendPath(basePath, leafName, PickPathSeparator(basePath));
}

[[nodiscard]] std::wstring GetPathLeaf(std::wstring_view path) noexcept
{
    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring_view::npos)
    {
        return std::wstring(path);
    }

    return std::wstring(path.substr(separator + 1u));
}

[[nodiscard]] std::wstring NormalizeVisitKey(std::wstring_view path) noexcept
{
    std::wstring key(path);
    std::replace(key.begin(), key.end(), L'/', L'\\');

    while (key.size() > 1u && (key.back() == L'\\' || key.back() == L'/'))
    {
        const bool keepDriveRoot = key.size() == 3u && key[1] == L':' && (key[2] == L'\\' || key[2] == L'/');
        const bool keepShareRoot = key.size() == 2u && key[0] == L'\\' && key[1] == L'\\';
        if (keepDriveRoot || keepShareRoot)
        {
            break;
        }
        key.pop_back();
    }

    return key;
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

HRESULT ReportSearchProgress(SearchRuntime& runtime, FileSystemSearchPhase phase, const std::wstring* currentPath, HRESULT statusHint, bool force) noexcept
{
    const uint64_t processedItems = runtime.scannedDirectories + runtime.scannedFiles;
    const ULONGLONG now           = ::GetTickCount64();

    if (! force && runtime.hasReportedProgress && phase == runtime.lastProgressPhase && (processedItems - runtime.lastProgressItems) < kProgressIntervalItems &&
        (now - runtime.lastProgressTick) < kProgressIntervalMs)
    {
        return S_OK;
    }

    FileSystemSearchProgress progress{};
    progress.sizeBytes          = sizeof(FileSystemSearchProgress);
    progress.phase              = phase;
    progress.backend            = FILESYSTEM_SEARCH_BACKEND_SCAN;
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

    if (! runtime.fileSystemIo)
    {
        runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT;
        return S_OK;
    }

    wil::com_ptr<IFileReader> reader;
    HRESULT hr = runtime.fileSystemIo->CreateFileReader(fullPath.c_str(), reader.put());
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        if (hr == kNotSupportedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT;
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

        if (hr == kNotSupportedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT;
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

HRESULT EmitSearchMatch(SearchRuntime& runtime, const SearchEntryMetadata& entry, uint32_t matchedBy, const SearchContentResult& contentResult) noexcept
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
    if (runtime.maxResults != 0u && runtime.matchedEntries >= runtime.maxResults)
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

    if (! isDirectory)
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

        unsigned long count = 0u;
        hr                  = information->GetCount(&count);
        if (FAILED(hr))
        {
            return hr;
        }

        for (unsigned long index = 0u; index < count && ! runtime.stopRequested; ++index)
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
            metadata.displayName = std::wstring(name);
            metadata.relativePath =
                frame.relativeBase.empty() ? metadata.displayName : AppendPath(frame.relativeBase, metadata.displayName, PickPathSeparator(frame.fullPath));
            metadata.fullPath       = AppendPath(frame.fullPath, metadata.displayName);
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

HRESULT Execute(IFileSystem* fileSystem, const FileSystemSearchQuery* query, IFileSystemSearchCallback* callback, void* cookie) noexcept
{
    if (fileSystem == nullptr || query == nullptr || callback == nullptr)
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

    try
    {
        SearchRuntime runtime{};
        runtime.fileSystem             = fileSystem;
        runtime.query                  = query;
        runtime.callback               = callback;
        runtime.cookie                 = cookie;
        runtime.rootPath               = query->rootPath;
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
        runtime.maxContentBytesPerFile = query->maxContentBytesPerFile != 0u ? query->maxContentBytesPerFile : SearchTextHelpers::kDefaultContentBytesPerFile;
        runtime.maxSnippetCharacters   = query->maxSnippetCharacters != 0u ? query->maxSnippetCharacters : SearchTextHelpers::kDefaultSnippetCharacters;
        if ((query->flags & FILESYSTEM_SEARCH_PREFER_INDEX) != 0)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
        }

        static_cast<void>(fileSystem->QueryInterface(__uuidof(IFileSystemIO), runtime.fileSystemIo.put_void()));

        if (query->nameMode == FILESYSTEM_SEARCH_NAME_REGEX)
        {
            const auto flags  = runtime.matchCaseName ? std::regex_constants::ECMAScript : (std::regex_constants::ECMAScript | std::regex_constants::icase);
            runtime.nameRegex = std::make_unique<std::wregex>(runtime.namePattern, flags);
        }

        if (query->contentMode == FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX)
        {
            const auto flags = runtime.matchCaseContent ? std::regex_constants::ECMAScript : (std::regex_constants::ECMAScript | std::regex_constants::icase);
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

        if (query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED && ! runtime.fileSystemIo)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT;
            hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_COMPLETED, nullptr, kNotSupportedHr, true);
            if (FAILED(hr))
            {
                return hr;
            }
            return S_OK;
        }

        bool rootIsDirectory         = true;
        unsigned long rootAttributes = FILE_ATTRIBUTE_DIRECTORY;
        if (runtime.fileSystemIo)
        {
            hr = runtime.fileSystemIo->GetAttributes(runtime.rootPath.c_str(), &rootAttributes);
            if (FAILED(hr))
            {
                return hr;
            }

            rootIsDirectory = (rootAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        }

        if (rootIsDirectory)
        {
            hr = SearchDirectoryTree(runtime);
        }
        else
        {
            SearchEntryMetadata entry{};
            entry.fullPath       = runtime.rootPath;
            entry.relativePath   = GetPathLeaf(runtime.rootPath);
            entry.displayName    = entry.relativePath;
            entry.fileAttributes = rootAttributes;

            if (runtime.fileSystemIo)
            {
                FileSystemBasicInformation info{};
                info.sizeBytes = sizeof(FileSystemBasicInformation);
                if (SUCCEEDED(runtime.fileSystemIo->GetFileBasicInformation(runtime.rootPath.c_str(), &info)))
                {
                    entry.creationTime   = info.creationTime;
                    entry.lastAccessTime = info.lastAccessTime;
                    entry.lastWriteTime  = info.lastWriteTime;
                }

                wil::com_ptr<IFileReader> reader;
                if (SUCCEEDED(runtime.fileSystemIo->CreateFileReader(runtime.rootPath.c_str(), reader.put())))
                {
                    uint64_t sizeBytes = 0u;
                    if (SUCCEEDED(reader->GetSize(&sizeBytes)))
                    {
                        entry.endOfFile      = static_cast<__int64>(sizeBytes);
                        entry.allocationSize = static_cast<__int64>(sizeBytes);
                    }
                }
            }

            hr = EvaluateEntry(runtime, entry);
        }

        const HRESULT finalStatus = FAILED(hr) ? hr : (runtime.stopRequested ? S_FALSE : S_OK);
        const HRESULT progressHr  = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_COMPLETED, nullptr, finalStatus, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        if (runtime.warningFlags != FILESYSTEM_SEARCH_WARNING_NONE)
        {
            Debug::Warning(
                L"SearchFallbackEngine: degraded root='{}' status=0x{:08X} matched={} candidates={} scannedFiles={} scannedDirs={} warnings=0x{:08X}",
                runtime.rootPath,
                static_cast<unsigned long>(FAILED(hr) ? hr : finalStatus),
                runtime.matchedEntries,
                runtime.candidateFiles,
                runtime.scannedFiles,
                runtime.scannedDirectories,
                runtime.warningFlags);
        }

        Debug::Info(L"SearchFallbackEngine: completed root='{}' status=0x{:08X} matched={} candidates={} scannedFiles={} scannedDirs={} warnings=0x{:08X}",
                    runtime.rootPath,
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
        Debug::Error(L"SearchFallbackEngine: Execute failed with an unexpected std::exception.");
        return E_FAIL;
    }
}
} // namespace SearchFallbackEngine

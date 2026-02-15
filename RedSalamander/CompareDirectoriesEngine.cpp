#include "Framework.h"

#include "CompareDirectoriesEngine.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <limits>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
struct SideEntry
{
    bool isDirectory      = false;
    uint64_t sizeBytes    = 0;
    int64_t lastWriteTime = 0;
    DWORD fileAttributes  = 0;
};

[[nodiscard]] bool IsMissingPathError(HRESULT hr) noexcept
{
    if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_DIRECTORY))
    {
        return true;
    }

    if (hr == HRESULT_FROM_WIN32(ERROR_BAD_NETPATH) || hr == HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME) || hr == HRESULT_FROM_WIN32(ERROR_INVALID_DRIVE))
    {
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }
    return text;
}

[[nodiscard]] std::vector<std::wstring> SplitPatterns(std::wstring_view patterns) noexcept
{
    std::vector<std::wstring> result;
    patterns = TrimWhitespace(patterns);
    if (patterns.empty())
    {
        return result;
    }

    size_t start = 0;
    while (start <= patterns.size())
    {
        const size_t sep = patterns.find(L';', start);
        const size_t end = sep == std::wstring_view::npos ? patterns.size() : sep;

        std::wstring_view token = patterns.substr(start, end - start);
        token                   = TrimWhitespace(token);
        if (! token.empty())
        {
            result.emplace_back(token);
        }

        if (sep == std::wstring_view::npos)
        {
            break;
        }
        start = sep + 1;
    }

    return result;
}

[[nodiscard]] wchar_t LowerInvariant(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool WildcardMatchNoCase(std::wstring_view text, std::wstring_view pattern) noexcept
{
    // Glob match with '*' and '?', case-insensitive.
    size_t ti = 0;
    size_t pi = 0;

    size_t star  = std::wstring_view::npos;
    size_t match = 0;

    while (ti < text.size())
    {
        if (pi < pattern.size())
        {
            const wchar_t pch = pattern[pi];
            if (pch == L'?')
            {
                ++ti;
                ++pi;
                continue;
            }
            if (pch == L'*')
            {
                star  = pi++;
                match = ti;
                continue;
            }

            if (LowerInvariant(text[ti]) == LowerInvariant(pch))
            {
                ++ti;
                ++pi;
                continue;
            }
        }

        if (star != std::wstring_view::npos)
        {
            pi = star + 1;
            ++match;
            ti = match;
            continue;
        }

        return false;
    }

    while (pi < pattern.size() && pattern[pi] == L'*')
    {
        ++pi;
    }

    return pi == pattern.size();
}

[[nodiscard]] bool MatchesAnyPattern(std::wstring_view name, const std::vector<std::wstring>& patterns) noexcept
{
    for (const std::wstring& pat : patterns)
    {
        if (pat.empty())
        {
            continue;
        }
        if (WildcardMatchNoCase(name, pat))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool ShouldIgnoreEntry(std::wstring_view name,
                                     bool isDirectory,
                                     const Common::Settings::CompareDirectoriesSettings& settings,
                                     const std::vector<std::wstring>& ignoreFilePatterns,
                                     const std::vector<std::wstring>& ignoreDirectoryPatterns) noexcept
{
    if (name.empty())
    {
        return true;
    }

    if (name == L"." || name == L"..")
    {
        return true;
    }

    if (isDirectory)
    {
        return settings.ignoreDirectories && MatchesAnyPattern(name, ignoreDirectoryPatterns);
    }

    return settings.ignoreFiles && MatchesAnyPattern(name, ignoreFilePatterns);
}
} // namespace

CompareDirectoriesSession::CompareDirectoriesSession(wil::com_ptr<IFileSystem> baseFileSystem,
                                                     std::filesystem::path leftRoot,
                                                     std::filesystem::path rightRoot,
                                                     Common::Settings::CompareDirectoriesSettings settings) :
    _baseFileSystem(std::move(baseFileSystem)),
    _leftRoot(std::move(leftRoot)),
    _rightRoot(std::move(rightRoot)),
    _settings(std::move(settings))
{
    if (_baseFileSystem)
    {
        wil::com_ptr<IInformations> infos;
        const HRESULT qiHr = _baseFileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
        if (SUCCEEDED(qiHr) && infos)
        {
            _baseInformations = std::move(infos);
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT qiIo = _baseFileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
        if (SUCCEEDED(qiIo) && io)
        {
            _baseFileSystemIo = std::move(io);
        }
    }
}

void CompareDirectoriesSession::SetRoots(std::filesystem::path leftRoot, std::filesystem::path rightRoot)
{
    std::lock_guard guard(_mutex);
    _leftRoot  = std::move(leftRoot);
    _rightRoot = std::move(rightRoot);
    ++_version;
    ++_uiVersion;
    _cache.clear();
}

void CompareDirectoriesSession::SetSettings(Common::Settings::CompareDirectoriesSettings settings)
{
    std::lock_guard guard(_mutex);

    const bool comparisonChanged = _settings.compareSize != settings.compareSize || _settings.compareDateTime != settings.compareDateTime ||
                                   _settings.compareAttributes != settings.compareAttributes || _settings.compareContent != settings.compareContent ||
                                   _settings.compareSubdirectories != settings.compareSubdirectories ||
                                   _settings.compareSubdirectoryAttributes != settings.compareSubdirectoryAttributes ||
                                   _settings.selectSubdirsOnlyInOnePane != settings.selectSubdirsOnlyInOnePane || _settings.ignoreFiles != settings.ignoreFiles ||
                                   _settings.ignoreFilesPatterns != settings.ignoreFilesPatterns || _settings.ignoreDirectories != settings.ignoreDirectories ||
                                   _settings.ignoreDirectoriesPatterns != settings.ignoreDirectoriesPatterns;

    _settings = std::move(settings);
    if (comparisonChanged)
    {
        ++_version;
        ++_uiVersion;
        _cache.clear();
    }
}

void CompareDirectoriesSession::Invalidate() noexcept
{
    std::lock_guard guard(_mutex);
    ++_version;
    ++_uiVersion;
    _cache.clear();
}

namespace
{
[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return true;
    }

    if (text.size() < prefix.size())
    {
        return false;
    }

    if (prefix.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int prefixLen = static_cast<int>(prefix.size());
    return CompareStringOrdinal(text.data(), prefixLen, prefix.data(), prefixLen, TRUE) == CSTR_EQUAL;
}
} // namespace

void CompareDirectoriesSession::InvalidateForRelativePathLocked(const std::filesystem::path& relativePath, bool includeSubtree) noexcept
{
    std::filesystem::path folder = relativePath.lexically_normal();

    if (! includeSubtree)
    {
        folder = folder.parent_path();
    }

    if (! folder.empty())
    {
        folder = folder.lexically_normal();
    }

    if (includeSubtree)
    {
        const std::wstring prefix = MakeCacheKey(folder);
        if (prefix.empty())
        {
            _cache.clear();
        }
        else
        {
            for (auto it = _cache.begin(); it != _cache.end();)
            {
                const std::wstring_view key = it->first;
                const bool match            = StartsWithNoCase(key, prefix) && (key.size() == prefix.size() || key[prefix.size()] == L'/');
                if (match)
                {
                    it = _cache.erase(it);
                }
                else
                {
                    ++it;
                }
            }
        }
    }

    for (std::filesystem::path current = folder;; current = current.parent_path())
    {
        _cache.erase(MakeCacheKey(current));
        if (current.empty())
        {
            break;
        }
    }

    ++_uiVersion;
}

void CompareDirectoriesSession::InvalidateForAbsolutePath(const std::filesystem::path& absolutePath, bool includeSubtree) noexcept
{
    if (absolutePath.empty())
    {
        return;
    }

    const std::optional<std::filesystem::path> relLeft  = TryMakeRelative(ComparePane::Left, absolutePath);
    const std::optional<std::filesystem::path> relRight = TryMakeRelative(ComparePane::Right, absolutePath);

    std::lock_guard guard(_mutex);
    if (relLeft.has_value())
    {
        InvalidateForRelativePathLocked(relLeft.value(), includeSubtree);
    }
    if (relRight.has_value())
    {
        InvalidateForRelativePathLocked(relRight.value(), includeSubtree);
    }
}

void CompareDirectoriesSession::SetScanProgressCallback(ScanProgressCallback callback) noexcept
{
    std::shared_ptr<const ScanProgressCallback> stored;
    if (callback)
    {
        stored = std::make_shared<ScanProgressCallback>(std::move(callback));
    }

    _scanProgressCallback.store(std::move(stored), std::memory_order_release);
}

Common::Settings::CompareDirectoriesSettings CompareDirectoriesSession::GetSettings() const
{
    std::lock_guard guard(_mutex);
    return _settings;
}

std::filesystem::path CompareDirectoriesSession::GetRoot(ComparePane pane) const
{
    std::lock_guard guard(_mutex);
    return pane == ComparePane::Left ? _leftRoot : _rightRoot;
}

uint64_t CompareDirectoriesSession::GetVersion() const noexcept
{
    std::lock_guard guard(_mutex);
    return _version;
}

uint64_t CompareDirectoriesSession::GetUiVersion() const noexcept
{
    std::lock_guard guard(_mutex);
    return _uiVersion;
}

wil::com_ptr<IFileSystem> CompareDirectoriesSession::GetBaseFileSystem() const noexcept
{
    return _baseFileSystem;
}

wil::com_ptr<IInformations> CompareDirectoriesSession::GetBaseInformations() const noexcept
{
    return _baseInformations;
}

wil::com_ptr<IFileSystemIO> CompareDirectoriesSession::GetBaseFileSystemIO() const noexcept
{
    return _baseFileSystemIo;
}

std::optional<std::filesystem::path> CompareDirectoriesSession::TryMakeRelative(ComparePane pane, const std::filesystem::path& absoluteFolder) const
{
    const std::filesystem::path rootPath = GetRoot(pane).lexically_normal();
    const std::filesystem::path absPath  = absoluteFolder.lexically_normal();

    auto normalizeText = [](std::filesystem::path value, bool lower) noexcept -> std::wstring
    {
        value = value.lexically_normal();
        while (! value.empty() && ! value.has_filename() && value != value.root_path())
        {
            value = value.parent_path();
        }

        std::wstring text = value.wstring();
        std::replace(text.begin(), text.end(), L'/', L'\\');
        if (lower)
        {
            for (auto& ch : text)
            {
                ch = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            }
        }
        return text;
    };

    const std::wstring rootNorm      = normalizeText(rootPath, false);
    const std::wstring absNorm       = normalizeText(absPath, false);
    const std::wstring rootNormLower = normalizeText(rootPath, true);
    const std::wstring absNormLower  = normalizeText(absPath, true);

    if (absNormLower == rootNormLower)
    {
        return std::filesystem::path{};
    }

    std::wstring rootPrefix      = rootNorm;
    std::wstring rootPrefixLower = rootNormLower;
    if (! rootPrefix.empty() && rootPrefix.back() != L'\\')
    {
        rootPrefix.push_back(L'\\');
        rootPrefixLower.push_back(L'\\');
    }

    if (absNormLower.rfind(rootPrefixLower, 0) != 0)
    {
        return std::nullopt;
    }

    const std::wstring_view relativeText(absNorm.c_str() + rootPrefix.size());
    std::filesystem::path relative(relativeText);
    relative = relative.lexically_normal();
    return relative;
}

std::filesystem::path CompareDirectoriesSession::ResolveAbsolute(ComparePane pane, const std::filesystem::path& relativeFolder) const
{
    const std::filesystem::path root = GetRoot(pane);
    if (relativeFolder.empty())
    {
        return root;
    }
    return (root / relativeFolder).lexically_normal();
}

std::wstring CompareDirectoriesSession::MakeCacheKey(const std::filesystem::path& relativeFolder) const
{
    if (relativeFolder.empty())
    {
        return {};
    }
    return relativeFolder.generic_wstring();
}

void CompareDirectoriesSession::NotifyScanProgress(const std::filesystem::path& relativeFolder, std::wstring_view currentEntryName, bool force) noexcept
{
    const auto cb = _scanProgressCallback.load(std::memory_order_acquire);
    if (! cb || ! *cb)
    {
        return;
    }

    if (! force)
    {
        const uint64_t now  = GetTickCount64();
        uint64_t lastUpdate = _scanLastNotifyTickMs.load(std::memory_order_relaxed);
        if ((now - lastUpdate) < 80u)
        {
            return;
        }

        if (! _scanLastNotifyTickMs.compare_exchange_strong(lastUpdate, now, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return;
        }
    }

    const uint64_t scannedFolders = _scanFoldersScanned.load(std::memory_order_relaxed);
    const uint64_t scannedEntries = _scanEntriesScanned.load(std::memory_order_relaxed);
    const uint32_t activeScans    = _scanActiveScans.load(std::memory_order_relaxed);

    (*cb)(relativeFolder, currentEntryName, scannedFolders, scannedEntries, activeScans);
}

namespace
{
[[nodiscard]] bool TryReadDirectoryEntries(const wil::com_ptr<IFileSystem>& baseFs,
                                           const std::filesystem::path& absoluteFolder,
                                           const Common::Settings::CompareDirectoriesSettings& settings,
                                           const std::vector<std::wstring>& ignoreFilePatterns,
                                           const std::vector<std::wstring>& ignoreDirectoryPatterns,
                                           std::unordered_map<std::wstring, SideEntry, WStringViewNoCaseHash, WStringViewNoCaseEq>& outEntries,
                                           HRESULT& outHr) noexcept
{
    outEntries.clear();
    outHr = S_OK;

    if (! baseFs)
    {
        outHr = E_POINTER;
        return false;
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = baseFs->ReadDirectoryInfo(absoluteFolder.c_str(), info.put());
    if (FAILED(hr))
    {
        if (IsMissingPathError(hr))
        {
            outHr = S_OK;
            return true;
        }

        outHr = hr;
        return false;
    }

    FileInfo* head = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    if (FAILED(hrBuffer))
    {
        outHr = hrBuffer;
        return false;
    }

    for (FileInfo* entry = head; entry != nullptr;)
    {
        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        const std::wstring_view name(entry->FileName, nameChars);

        const bool isDir = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (! ShouldIgnoreEntry(name, isDir, settings, ignoreFilePatterns, ignoreDirectoryPatterns))
        {
            SideEntry out{};
            out.isDirectory      = isDir;
            out.fileAttributes   = entry->FileAttributes;
            out.lastWriteTime    = entry->LastWriteTime;
            out.sizeBytes        = (! isDir && entry->EndOfFile > 0) ? static_cast<uint64_t>(entry->EndOfFile) : 0;
            outEntries.emplace(std::wstring(name), out);
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    outHr = S_OK;
    return true;
}

[[nodiscard]] bool AreFilesEqualContent(IFileSystemIO* io,
                                        const std::filesystem::path& leftPath,
                                        const std::filesystem::path& rightPath,
                                        uint64_t sizeBytes) noexcept
{
    if (sizeBytes == 0)
    {
        return true;
    }

    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileReader> left;
    const HRESULT hrLeft = io->CreateFileReader(leftPath.c_str(), left.put());
    if (FAILED(hrLeft) || ! left)
    {
        return false;
    }

    wil::com_ptr<IFileReader> right;
    const HRESULT hrRight = io->CreateFileReader(rightPath.c_str(), right.put());
    if (FAILED(hrRight) || ! right)
    {
        return false;
    }

    // Trust the enumerated size as a fast path but validate if available.
    unsigned __int64 leftSize  = 0;
    unsigned __int64 rightSize = 0;
    const HRESULT hrLeftSize   = left->GetSize(&leftSize);
    const HRESULT hrRightSize  = right->GetSize(&rightSize);
    if (SUCCEEDED(hrLeftSize) && SUCCEEDED(hrRightSize))
    {
        if (leftSize != rightSize)
        {
            return false;
        }

        if (leftSize != sizeBytes)
        {
            sizeBytes = leftSize;
            if (sizeBytes == 0)
            {
                return true;
            }
        }
    }

    std::array<std::byte, 256 * 1024> leftBuf{};
    std::array<std::byte, 256 * 1024> rightBuf{};

    uint64_t remaining = sizeBytes;
    while (remaining > 0)
    {
        const DWORD chunk = remaining > leftBuf.size() ? static_cast<DWORD>(leftBuf.size()) : static_cast<DWORD>(remaining);

        unsigned long leftRead = 0;
        const HRESULT hrReadLeft = left->Read(leftBuf.data(), chunk, &leftRead);
        if (FAILED(hrReadLeft))
        {
            return false;
        }

        unsigned long rightRead = 0;
        const HRESULT hrReadRight = right->Read(rightBuf.data(), chunk, &rightRead);
        if (FAILED(hrReadRight))
        {
            return false;
        }

        if (leftRead != chunk || rightRead != chunk)
        {
            return false;
        }

        if (std::memcmp(leftBuf.data(), rightBuf.data(), chunk) != 0)
        {
            return false;
        }

        remaining -= chunk;
    }

    return true;
}

[[nodiscard]] bool AnyChildDifferent(const CompareDirectoriesFolderDecision& decision) noexcept
{
    for (const auto& kv : decision.items)
    {
        if (kv.second.isDifferent)
        {
            return true;
        }
    }
    return false;
}

class CompareFilesInformation final : public IFilesInformation
{
public:
    CompareFilesInformation(std::vector<unsigned char> buffer, std::vector<FileInfo*> entries) noexcept :
        _buffer(std::move(buffer)),
        _entries(std::move(entries))
    {
    }

    CompareFilesInformation(const CompareFilesInformation&)            = delete;
    CompareFilesInformation& operator=(const CompareFilesInformation&) = delete;
    CompareFilesInformation(CompareFilesInformation&&)                 = delete;
    CompareFilesInformation& operator=(CompareFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
        {
            *ppvObject = static_cast<IFilesInformation*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override
    {
        if (ppFileInfo == nullptr)
        {
            return E_POINTER;
        }

        if (_buffer.empty())
        {
            *ppFileInfo = nullptr;
            return S_OK;
        }

        *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.capacity());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override
    {
        if (pCount == nullptr)
        {
            return E_POINTER;
        }

        *pCount = static_cast<unsigned long>(_entries.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override
    {
        if (ppEntry == nullptr)
        {
            return E_POINTER;
        }

        *ppEntry = nullptr;
        if (index >= _entries.size())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX);
        }

        *ppEntry = _entries[index];
        return S_OK;
    }

private:
    ~CompareFilesInformation() = default;

    std::atomic_ulong _refCount{1};
    std::vector<unsigned char> _buffer;
    std::vector<FileInfo*> _entries;
};

struct OutEntry
{
    std::wstring name;
    DWORD fileAttributes  = 0;
    int64_t lastWriteTime = 0;
    uint64_t sizeBytes    = 0;
};

[[nodiscard]] size_t AlignedFileInfoSizeBytes(size_t nameChars) noexcept
{
    constexpr size_t kAlign = 8;
    const size_t raw        = offsetof(FileInfo, FileName) + (nameChars * sizeof(wchar_t));
    return (raw + (kAlign - 1)) & ~(kAlign - 1);
}

[[nodiscard]] wil::com_ptr<IFilesInformation> BuildFilesInformation(const std::vector<OutEntry>& entries, HRESULT& outHr) noexcept
{
    outHr = S_OK;

    if (entries.empty())
    {
        return wil::com_ptr<IFilesInformation>(new CompareFilesInformation({}, {}));
    }

    size_t totalBytes = 0;
    for (const auto& e : entries)
    {
        totalBytes += AlignedFileInfoSizeBytes(e.name.size());
    }

    if (totalBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        outHr = E_OUTOFMEMORY;
        return nullptr;
    }

    std::vector<unsigned char> buffer(totalBytes);
    std::vector<FileInfo*> entryPtrs;
    entryPtrs.reserve(entries.size());

    size_t offset = 0;
    for (size_t i = 0; i < entries.size(); ++i)
    {
        const OutEntry& src = entries[i];
        const size_t entrySize = AlignedFileInfoSizeBytes(src.name.size());
        if (offset + entrySize > buffer.size())
        {
            outHr = E_FAIL;
            return nullptr;
        }

        FileInfo* dst = reinterpret_cast<FileInfo*>(buffer.data() + offset);
        entryPtrs.push_back(dst);

        dst->NextEntryOffset = (i + 1 < entries.size()) ? static_cast<unsigned long>(entrySize) : 0;
        dst->FileIndex       = 0;
        dst->CreationTime    = 0;
        dst->LastAccessTime  = 0;
        dst->LastWriteTime   = src.lastWriteTime;
        dst->ChangeTime      = 0;
        dst->EndOfFile       = static_cast<__int64>(src.sizeBytes);
        dst->AllocationSize  = 0;
        dst->FileAttributes  = src.fileAttributes;
        dst->FileNameSize    = static_cast<unsigned long>(src.name.size() * sizeof(wchar_t));
        dst->EaSize          = 0;

        if (! src.name.empty())
        {
            std::memcpy(dst->FileName, src.name.data(), dst->FileNameSize);
        }

        offset += entrySize;
    }

    return wil::com_ptr<IFilesInformation>(new CompareFilesInformation(std::move(buffer), std::move(entryPtrs)));
}
} // namespace

std::shared_ptr<const CompareDirectoriesFolderDecision> CompareDirectoriesSession::GetOrComputeDecision(const std::filesystem::path& relativeFolder)
{
    const std::wstring key = MakeCacheKey(relativeFolder);

    uint64_t version = 0;
    {
        std::lock_guard guard(_mutex);
        version = _version;

        const auto it = _cache.find(key);
        if (it != _cache.end() && it->second && it->second->version == version)
        {
            return it->second;
        }
    }

    const uint32_t activeBefore = _scanActiveScans.fetch_add(1u, std::memory_order_acq_rel);
    if (activeBefore == 0u)
    {
        _scanFoldersScanned.store(0u, std::memory_order_release);
        _scanEntriesScanned.store(0u, std::memory_order_release);
        _scanLastNotifyTickMs.store(0u, std::memory_order_release);
    }

    static_cast<void>(_scanFoldersScanned.fetch_add(1u, std::memory_order_acq_rel));
    const bool scanStarted = (activeBefore == 0u);
    NotifyScanProgress(relativeFolder, {}, scanStarted);

    auto scanCleanup = wil::scope_exit(
        [&]
        {
            const uint32_t activeAfter = _scanActiveScans.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
            if (activeAfter == 0u)
            {
                NotifyScanProgress(relativeFolder, {}, true);
            }
        });

    const Common::Settings::CompareDirectoriesSettings settings = GetSettings();
    const std::vector<std::wstring> ignoreFilePatterns         = SplitPatterns(settings.ignoreFilesPatterns);
    const std::vector<std::wstring> ignoreDirectoryPatterns    = SplitPatterns(settings.ignoreDirectoriesPatterns);

    const std::filesystem::path leftFolder  = ResolveAbsolute(ComparePane::Left, relativeFolder);
    const std::filesystem::path rightFolder = ResolveAbsolute(ComparePane::Right, relativeFolder);

    auto decision     = std::make_shared<CompareDirectoriesFolderDecision>();
    decision->version = version;
    decision->hr      = S_OK;

    std::unordered_map<std::wstring, SideEntry, WStringViewNoCaseHash, WStringViewNoCaseEq> leftEntries;
    std::unordered_map<std::wstring, SideEntry, WStringViewNoCaseHash, WStringViewNoCaseEq> rightEntries;

    HRESULT leftHr = S_OK;
    if (! TryReadDirectoryEntries(_baseFileSystem, leftFolder, settings, ignoreFilePatterns, ignoreDirectoryPatterns, leftEntries, leftHr))
    {
        decision->hr = leftHr;
    }

    HRESULT rightHr = S_OK;
    if (SUCCEEDED(decision->hr) &&
        ! TryReadDirectoryEntries(_baseFileSystem, rightFolder, settings, ignoreFilePatterns, ignoreDirectoryPatterns, rightEntries, rightHr))
    {
        decision->hr = rightHr;
    }

    if (SUCCEEDED(decision->hr))
    {
        // Seed with left entries first (preserves left casing as key when both exist).
        for (const auto& [name, entry] : leftEntries)
        {
            CompareDirectoriesItemDecision item{};
            item.existsLeft         = true;
            item.isDirectory        = entry.isDirectory;
            item.leftSizeBytes      = entry.sizeBytes;
            item.leftLastWriteTime  = entry.lastWriteTime;
            item.leftFileAttributes = entry.fileAttributes;
            decision->items.emplace(name, item);
        }

        for (const auto& [name, entry] : rightEntries)
        {
            auto it = decision->items.find(name);
            if (it == decision->items.end())
            {
                CompareDirectoriesItemDecision item{};
                item.existsRight         = true;
                item.isDirectory         = entry.isDirectory;
                item.rightSizeBytes      = entry.sizeBytes;
                item.rightLastWriteTime  = entry.lastWriteTime;
                item.rightFileAttributes = entry.fileAttributes;
                decision->items.emplace(name, item);
                continue;
            }

            CompareDirectoriesItemDecision& item = it->second;
            item.existsRight         = true;
            item.isDirectory         = item.isDirectory || entry.isDirectory;
            item.rightSizeBytes      = entry.sizeBytes;
            item.rightLastWriteTime  = entry.lastWriteTime;
            item.rightFileAttributes = entry.fileAttributes;
        }

        for (auto& [name, item] : decision->items)
        {
            const uint64_t scannedEntries = _scanEntriesScanned.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            if ((scannedEntries & 0x3Fu) == 0u)
            {
                NotifyScanProgress(relativeFolder, name, false);
            }

            item.isDifferent = false;
            item.selectLeft  = false;
            item.selectRight = false;
            item.differenceMask = 0;

            if (item.existsLeft != item.existsRight)
            {
                item.isDifferent = true;
                if (item.existsLeft)
                {
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::OnlyInLeft);
                    item.selectLeft = ! item.isDirectory || settings.selectSubdirsOnlyInOnePane;
                }
                if (item.existsRight)
                {
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::OnlyInRight);
                    item.selectRight = ! item.isDirectory || settings.selectSubdirsOnlyInOnePane;
                }
            }
            else if (item.existsLeft && item.existsRight)
            {
                const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

                if (leftIsDir != rightIsDir)
                {
                    item.isDifferent = true;
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::TypeMismatch);
                    item.selectLeft  = true;
                    item.selectRight = true;
                }
                else if (leftIsDir)
                {
                    bool attrsDifferent = false;
                    if (settings.compareSubdirectoryAttributes && item.leftFileAttributes != item.rightFileAttributes)
                    {
                        attrsDifferent = true;
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirAttributes);
                    }

                    bool contentDifferent = false;
                    if (settings.compareSubdirectories)
                    {
                        // Avoid following directory reparse points (symlinks/junctions).
                        const bool leftReparse  = (item.leftFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        const bool rightReparse = (item.rightFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        if (! leftReparse && ! rightReparse)
                        {
                            const std::filesystem::path childRel = relativeFolder / std::filesystem::path(name);
                            const auto childDecision             = GetOrComputeDecision(childRel);
                            if (! childDecision || FAILED(childDecision->hr) || AnyChildDifferent(*childDecision))
                            {
                                contentDifferent = true;
                            }
                        }
                    }

                    if (contentDifferent)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirContent);
                    }

                    if (attrsDifferent || contentDifferent)
                    {
                        item.isDifferent = true;
                        item.selectLeft  = true;
                        item.selectRight = true;
                    }
                }
                else
                {
                    const bool sizeDifferent  = item.leftSizeBytes != item.rightSizeBytes;
                    const bool timeDifferent  = item.leftLastWriteTime != item.rightLastWriteTime;
                    const bool attrsDifferent = item.leftFileAttributes != item.rightFileAttributes;

                    bool contentDifferent = false;
                    if (settings.compareContent)
                    {
                        if (sizeDifferent)
                        {
                            contentDifferent = true;
                        }
                        else
                        {
                            const std::filesystem::path leftPath  = leftFolder / std::filesystem::path(name);
                            const std::filesystem::path rightPath = rightFolder / std::filesystem::path(name);
                            contentDifferent                      = ! AreFilesEqualContent(_baseFileSystemIo.get(), leftPath, rightPath, item.leftSizeBytes);
                        }
                    }

                    if (settings.compareSize && sizeDifferent)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Size);
                    }

                    if (settings.compareDateTime && timeDifferent)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::DateTime);
                    }

                    if (settings.compareAttributes && attrsDifferent)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Attributes);
                    }

                    if (settings.compareContent && contentDifferent)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Content);
                    }

                    const bool anyCriteriaDifferent = (settings.compareSize && sizeDifferent) || (settings.compareDateTime && timeDifferent) ||
                                                      (settings.compareAttributes && attrsDifferent) || (settings.compareContent && contentDifferent);

                    if (anyCriteriaDifferent)
                    {
                        item.isDifferent = true;

                        if (settings.compareSize && sizeDifferent)
                        {
                            if (item.leftSizeBytes > item.rightSizeBytes)
                            {
                                item.selectLeft = true;
                            }
                            else
                            {
                                item.selectRight = true;
                            }
                        }

                        if (settings.compareDateTime && timeDifferent)
                        {
                            if (item.leftLastWriteTime > item.rightLastWriteTime)
                            {
                                item.selectLeft = true;
                            }
                            else
                            {
                                item.selectRight = true;
                            }
                        }

                        if (settings.compareAttributes && attrsDifferent)
                        {
                            item.selectLeft  = true;
                            item.selectRight = true;
                        }

                        if (settings.compareContent && contentDifferent)
                        {
                            item.selectLeft  = true;
                            item.selectRight = true;
                        }
                    }
                }
            }
        }
    }

    {
        std::lock_guard guard(_mutex);
        if (_version == version)
        {
            _cache[key] = decision;
        }
    }

    return decision;
}

class CompareDirectoriesFileSystem final : public IFileSystem, public IInformations
{
public:
    CompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept :
        _pane(pane),
        _session(std::move(session))
    {
        if (_session)
        {
            _baseFs    = _session->GetBaseFileSystem();
            _baseInfos = _session->GetBaseInformations();
        }
    }

    CompareDirectoriesFileSystem(const CompareDirectoriesFileSystem&)            = delete;
    CompareDirectoriesFileSystem& operator=(const CompareDirectoriesFileSystem&) = delete;
    CompareDirectoriesFileSystem(CompareDirectoriesFileSystem&&)                 = delete;
    CompareDirectoriesFileSystem& operator=(CompareDirectoriesFileSystem&&)      = delete;

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

        if (riid == __uuidof(IInformations))
        {
            *ppvObject = static_cast<IInformations*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IInformations
    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetMetaData(metaData);
    }

    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetConfigurationSchema(schemaJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->SetConfiguration(configurationJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetConfiguration(configurationJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->SomethingToSave(pSomethingToSave);
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (ppFilesInformation == nullptr)
        {
            return E_POINTER;
        }

        *ppFilesInformation = nullptr;

        if (! _session)
        {
            return E_POINTER;
        }

        const std::filesystem::path absolute(path ? path : L"");
        const auto relOpt = _session->TryMakeRelative(_pane, absolute);
        if (! relOpt.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const auto decision = _session->GetOrComputeDecision(relOpt.value());
        if (! decision)
        {
            return E_FAIL;
        }

        if (FAILED(decision->hr))
        {
            return decision->hr;
        }

        const Common::Settings::CompareDirectoriesSettings settings = _session->GetSettings();
        const bool showIdentical                                   = settings.showIdenticalItems;

        std::vector<OutEntry> out;
        out.reserve(decision->items.size());

        const bool isLeft = _pane == ComparePane::Left;
        for (const auto& [name, item] : decision->items)
        {
            const bool include = isLeft ? (item.existsLeft && (showIdentical || item.isDifferent)) : (item.existsRight && (showIdentical || item.isDifferent));
            if (! include)
            {
                continue;
            }

            OutEntry e{};
            e.name = name;
            if (isLeft)
            {
                e.fileAttributes = item.leftFileAttributes;
                e.lastWriteTime  = item.leftLastWriteTime;
                e.sizeBytes      = item.leftSizeBytes;
            }
            else
            {
                e.fileAttributes = item.rightFileAttributes;
                e.lastWriteTime  = item.rightLastWriteTime;
                e.sizeBytes      = item.rightSizeBytes;
            }

            out.emplace_back(std::move(e));
        }

        const auto compareName = [](const OutEntry& a, const OutEntry& b) noexcept
        {
            const int cmp = CompareStringOrdinal(a.name.data(), static_cast<int>(a.name.size()), b.name.data(), static_cast<int>(b.name.size()), TRUE);
            if (cmp != CSTR_EQUAL)
            {
                return cmp == CSTR_LESS_THAN;
            }

            const int caseCmp = CompareStringOrdinal(a.name.data(), static_cast<int>(a.name.size()), b.name.data(), static_cast<int>(b.name.size()), FALSE);
            if (caseCmp != CSTR_EQUAL)
            {
                return caseCmp == CSTR_LESS_THAN;
            }

            return a.name < b.name;
        };

        std::sort(out.begin(), out.end(), compareName);

        HRESULT infoHr = S_OK;
        wil::com_ptr<IFilesInformation> info = BuildFilesInformation(out, infoHr);
        if (FAILED(infoHr) || ! info)
        {
            return FAILED(infoHr) ? infoHr : E_FAIL;
        }

        *ppFilesInformation = info.detach();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t* itemPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->DeleteItem(itemPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _baseFs ? _baseFs->GetCapabilities(jsonUtf8) : E_POINTER;
    }

private:
    ~CompareDirectoriesFileSystem() = default;

    std::atomic_ulong _refCount{1};
    ComparePane _pane = ComparePane::Left;
    std::shared_ptr<CompareDirectoriesSession> _session;
    wil::com_ptr<IFileSystem> _baseFs;
    wil::com_ptr<IInformations> _baseInfos;
};

wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept
{
    if (! session)
    {
        return nullptr;
    }

    wil::com_ptr<IFileSystem> fs;
    fs.attach(static_cast<IFileSystem*>(new CompareDirectoriesFileSystem(pane, std::move(session))));
    return fs;
}

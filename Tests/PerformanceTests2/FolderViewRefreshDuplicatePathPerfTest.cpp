#include "pch.h"

#include "DirectoryInfoCache.h"
#include "FolderViewInternal.Access.h"
#include "FolderWatcher.h"
#include "IconCache.h"
#include "PlugInterfaces/FileSystem.h"
#include "WSLDistro.h"
#include <array>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <string>
#include <string_view>
#include <vector>
#include <wincodec.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace
{
constexpr char kReadOnlyFileSystemCapabilitiesJson[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": false,
    "rename": false,
    "properties": false,
    "read": false,
    "write": false
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": [], "move": [] },
    "import": { "copy": [], "move": [] }
  }
}
)json";

class TestFilesInformation final : public IFilesInformation
{
public:
    TestFilesInformation(std::vector<std::byte> buffer, unsigned long count, unsigned long usedBytes) noexcept
        : _buffer(std::move(buffer)),
          _count(count),
          _usedBytes(usedBytes)
    {
    }
    TestFilesInformation(const TestFilesInformation&)            = delete;
    TestFilesInformation& operator=(const TestFilesInformation&) = delete;
    TestFilesInformation(TestFilesInformation&&)                 = delete;
    TestFilesInformation& operator=(TestFilesInformation&&)      = delete;

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
        const ULONG remaining = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override
    {
        if (ppFileInfo == nullptr)
        {
            return E_POINTER;
        }

        *ppFileInfo = _buffer.empty() ? nullptr : reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = _usedBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override
    {
        if (pCount == nullptr)
        {
            return E_POINTER;
        }

        *pCount = _count;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override
    {
        if (ppEntry == nullptr)
        {
            return E_POINTER;
        }

        *ppEntry = nullptr;
        if (index >= _count)
        {
            return E_INVALIDARG;
        }

        FileInfo* current = _buffer.empty() ? nullptr : reinterpret_cast<FileInfo*>(_buffer.data());
        for (unsigned long i = 0; i < index && current != nullptr; ++i)
        {
            if (current->NextEntryOffset == 0)
            {
                return E_INVALIDARG;
            }
            current = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(current) + current->NextEntryOffset);
        }

        *ppEntry = current;
        return current ? S_OK : E_INVALIDARG;
    }

private:
    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};

class DuplicatePathFileSystem final : public IFileSystem
{
public:
    struct Entry
    {
        std::wstring displayName;
        DWORD fileAttributes = FILE_ATTRIBUTE_NORMAL;
        uint64_t sizeBytes   = 0;
    };

    explicit DuplicatePathFileSystem(std::vector<Entry> entries) noexcept : _entries(std::move(entries))
    {
    }
    DuplicatePathFileSystem(const DuplicatePathFileSystem&)            = delete;
    DuplicatePathFileSystem& operator=(const DuplicatePathFileSystem&) = delete;
    DuplicatePathFileSystem(DuplicatePathFileSystem&&)                 = delete;
    DuplicatePathFileSystem& operator=(DuplicatePathFileSystem&&)      = delete;

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

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t*, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (ppFilesInformation == nullptr)
        {
            return E_POINTER;
        }

        std::vector<std::byte> buffer;
        unsigned long usedBytes = 0;
        buffer.reserve(_entries.size() * 128u);

        for (size_t i = 0; i < _entries.size(); ++i)
        {
            const Entry& entry        = _entries[i];
            const size_t nameBytes    = entry.displayName.size() * sizeof(wchar_t);
            const size_t entryBytes   = offsetof(FileInfo, FileName) + nameBytes;
            const size_t alignedBytes = (entryBytes + (alignof(FileInfo) - 1u)) & ~(alignof(FileInfo) - 1u);
            const size_t startOffset  = buffer.size();
            buffer.resize(startOffset + alignedBytes);

            auto* fileInfo = reinterpret_cast<FileInfo*>(buffer.data() + startOffset);
            ZeroMemory(fileInfo, alignedBytes);
            fileInfo->NextEntryOffset = (i + 1u < _entries.size()) ? static_cast<unsigned long>(alignedBytes) : 0;
            fileInfo->FileIndex       = static_cast<unsigned long>(i);
            fileInfo->FileAttributes  = entry.fileAttributes;
            fileInfo->EndOfFile       = static_cast<__int64>(entry.sizeBytes);
            fileInfo->AllocationSize  = static_cast<__int64>(entry.sizeBytes);
            fileInfo->FileNameSize    = static_cast<unsigned long>(nameBytes);
            if (nameBytes != 0)
            {
                memcpy(fileInfo->FileName, entry.displayName.data(), nameBytes);
            }
            usedBytes += static_cast<unsigned long>(alignedBytes);
        }

        *ppFilesInformation = new TestFilesInformation(std::move(buffer), static_cast<unsigned long>(_entries.size()), usedBytes);
        return *ppFilesInformation ? S_OK : E_OUTOFMEMORY;
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE
    RenameItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE
    CopyItems(const wchar_t* const*, unsigned long, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE
    MoveItems(const wchar_t* const*, unsigned long, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE
    DeleteItems(const wchar_t* const*, unsigned long, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE
    RenameItems(const FileSystemRenamePair*, unsigned long, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }

        *jsonUtf8 = kReadOnlyFileSystemCapabilitiesJson;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTransferHints([[maybe_unused]] const wchar_t* path,
                                               [[maybe_unused]] FileSystemOperation operationType,
                                               [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        if (! path || path[0] == L'\0' || ! hints)
        {
            return E_INVALIDARG;
        }
        if (hints->sizeBytes < sizeof(FileSystemTransferHints))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                        FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        if (! path || path[0] == L'\0' || ! characteristics)
        {
            return E_INVALIDARG;
        }
        if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

private:
    std::atomic_ulong _refCount{1};
    std::vector<Entry> _entries;
};

void WriteSmallTextFile(const std::filesystem::path& path, std::string_view content)
{
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    Assert::IsTrue(stream.is_open(), L"Failed to create test file.");
    stream.write(content.data(), static_cast<std::streamsize>(content.size()));
    Assert::IsTrue(stream.good(), L"Failed to write test file.");
}

[[nodiscard]] std::filesystem::path GetCurrentModulePath()
{
    std::wstring path(MAX_PATH, L'\0');
    for (;;)
    {
        wchar_t* const buffer = path.empty() ? nullptr : &path[0];
        const DWORD length    = GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(path.size()));
        if (length == 0)
        {
            return {};
        }

        if (length < path.size() - 1u)
        {
            path.resize(static_cast<size_t>(length));
            return std::filesystem::path(path);
        }

        path.resize(path.size() * 2u);
    }
}

[[nodiscard]] std::filesystem::path GetSharedDuplicatePluginFolderRoot(std::error_code& ec)
{
    return PerformanceTests2::AcquirePerformanceTestSandbox(L"folder_view_refresh_duplicate_path_perf_shared", ec);
}

void PopulateDuplicatePluginEntries(std::vector<DuplicatePathFileSystem::Entry>& entries)
{
    entries.clear();

    static constexpr std::array<const wchar_t*, 4> kCachedExtensions{{
        L".txt",
        L".cpp",
        L".json",
        L".log",
    }};

    for (int i = 0; i < 1200; ++i)
    {
        const wchar_t* const extension = kCachedExtensions[static_cast<size_t>(i) % kCachedExtensions.size()];
        const std::wstring displayName = std::wstring(L"cached_") + std::to_wstring(i) + extension;
        entries.push_back({displayName, FILE_ATTRIBUTE_NORMAL, 7});
    }

    constexpr int kUniqueExeCount = 20;
    constexpr int kUniqueUrlCount = 20;
    constexpr int kDuplicateCount = 8;

    for (int i = 0; i < kUniqueExeCount; ++i)
    {
        const std::wstring displayName = std::wstring(L"plugin_dup_tool_") + std::to_wstring(i) + L".exe";
        for (int duplicate = 0; duplicate < kDuplicateCount; ++duplicate)
        {
            entries.push_back({displayName, FILE_ATTRIBUTE_NORMAL, 0});
        }
    }

    for (int i = 0; i < kUniqueUrlCount; ++i)
    {
        const std::wstring displayName = std::wstring(L"plugin_dup_shortcut_") + std::to_wstring(i) + L".url";
        const std::string content      = "[InternetShortcut]\r\nURL=https://example.invalid/plugin/" + std::to_string(i) + "\r\n";
        for (int duplicate = 0; duplicate < kDuplicateCount; ++duplicate)
        {
            entries.push_back({displayName, FILE_ATTRIBUTE_NORMAL, static_cast<uint64_t>(content.size())});
        }
    }
}

void EnsurePhysicalDuplicatePluginFolder(const std::filesystem::path& root)
{
    const std::filesystem::path readyMarker = root / L"dataset.ready.v1";
    if (std::filesystem::exists(readyMarker))
    {
        return;
    }

    std::filesystem::create_directories(root);

    for (int i = 0; i < 1200; ++i)
    {
        static constexpr std::array<const wchar_t*, 4> kCachedExtensions{{
            L".txt",
            L".cpp",
            L".json",
            L".log",
        }};

        const wchar_t* const extension = kCachedExtensions[static_cast<size_t>(i) % kCachedExtensions.size()];
        const std::wstring displayName = std::wstring(L"cached_") + std::to_wstring(i) + extension;
        WriteSmallTextFile(root / displayName, "payload");
    }

    const std::filesystem::path modulePath = GetCurrentModulePath();
    Assert::IsTrue(! modulePath.empty(), L"Failed to resolve current module path.");

    for (int i = 0; i < 20; ++i)
    {
        const std::wstring displayName = std::wstring(L"plugin_dup_tool_") + std::to_wstring(i) + L".exe";
        std::filesystem::copy_file(modulePath, root / displayName, std::filesystem::copy_options::overwrite_existing);
    }

    for (int i = 0; i < 20; ++i)
    {
        const std::wstring displayName = std::wstring(L"plugin_dup_shortcut_") + std::to_wstring(i) + L".url";
        const std::string content      = "[InternetShortcut]\r\nURL=https://example.invalid/plugin/" + std::to_string(i) + "\r\n";
        WriteSmallTextFile(root / displayName, content);
    }

    WriteSmallTextFile(readyMarker, "ready");
}

} // namespace

FolderView::FolderView() : _theme(_appTheme.folderView)
{
}

FolderView::~FolderView() = default;

void FolderView::SetAppTheme(const AppTheme& theme)
{
    const bool compactModeChanged = _appTheme.compactMode != theme.compactMode;
    _appTheme                     = theme;
    if (compactModeChanged)
    {
        LayoutItems();
        UpdateScrollMetrics();
    }
}

void FolderView::ScheduleIdleLayoutCreation()
{
}
void FolderView::ExitIncrementalSearch() noexcept
{
}
void FolderView::NotifySelectionChanged() const noexcept
{
}
void FolderView::UpdateScrollMetrics()
{
}
void FolderView::LayoutItems()
{
    const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));

    _columnCounts.clear();
    _columnPrefixSums.clear();

    if (_items.empty() || clientWidthDip <= 0.0f)
    {
        _columns          = 1;
        _rowsPerColumn    = 0;
        _contentHeight    = std::max(clientHeightDip, 0.0f);
        _contentWidth     = std::max(clientWidthDip, 0.0f);
        _horizontalOffset = 0.0f;
        return;
    }

    _tileWidthDip  = std::min(200.0f, std::max(120.0f, clientWidthDip - (kColumnSpacingDip * 2.0f)));
    _tileHeightDip = 24.0f;

    const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
    const float columnStride  = _tileWidthDip + kColumnSpacingDip;
    const float rowStride     = _tileHeightDip + rowSpacingDip;
    _rowsPerColumn            = std::max(1, static_cast<int>(std::floor((clientHeightDip + rowSpacingDip) / rowStride)));
    _columns                  = std::max(1, static_cast<int>(std::ceil(static_cast<float>(_items.size()) / static_cast<float>(_rowsPerColumn))));

    _columnCounts.reserve(static_cast<size_t>(_columns));
    size_t remaining = _items.size();
    for (int column = 0; column < _columns && remaining > 0; ++column)
    {
        const int count = static_cast<int>(std::min<size_t>(static_cast<size_t>(_rowsPerColumn), remaining));
        _columnCounts.push_back(count);
        remaining -= count;
    }

    _columns = static_cast<int>(_columnCounts.size());
    if (_columns < 1)
    {
        _columns = 1;
    }

    _columnPrefixSums.reserve(_columnCounts.size() + 1u);
    size_t prefixSum = 0u;
    for (const int count : _columnCounts)
    {
        _columnPrefixSums.push_back(prefixSum);
        prefixSum += static_cast<size_t>(count);
    }
    _columnPrefixSums.push_back(prefixSum);

    size_t index    = 0u;
    float x         = kColumnSpacingDip;
    float maxBottom = 0.0f;
    float maxRight  = 0.0f;
    for (int column = 0; column < static_cast<int>(_columnCounts.size()) && index < _items.size(); ++column)
    {
        const int itemsInColumn = _columnCounts[static_cast<size_t>(column)];
        float y                 = rowSpacingDip;
        for (int row = 0; row < itemsInColumn && index < _items.size(); ++row, ++index)
        {
            auto& item  = _items[index];
            item.column = column;
            item.row    = row;
            item.bounds = D2D1::RectF(x, y, x + _tileWidthDip, y + _tileHeightDip);
            y += rowStride;
            maxBottom = std::max(maxBottom, item.bounds.bottom);
            maxRight  = std::max(maxRight, item.bounds.right);
        }
        x += columnStride;
    }

    _contentHeight                  = clientHeightDip;
    _contentWidth                   = std::max(maxRight + kColumnSpacingDip, clientWidthDip);
    _scrollOffset                   = 0.0f;
    const float viewWidthDip        = std::max(clientWidthDip, 0.0f);
    const float maxHorizontalOffset = std::max(0.0f, _contentWidth - viewWidthDip);
    _horizontalOffset               = std::clamp(_horizontalOffset, 0.0f, maxHorizontalOffset);
}
void FolderView::CancelBusyOverlay(uint64_t)
{
}
void FolderView::ScheduleBusyOverlay(uint64_t, const std::filesystem::path&)
{
}
bool FolderView::UpdateIncrementalSearchIndicatorAnimation(uint64_t) const noexcept
{
    return false;
}
void FolderView::StopOverlayAnimation() const noexcept
{
}
void FolderView::StopOverlayTimer() const
{
}
void FolderView::ClearErrorOverlay(ErrorOverlayKind) const
{
}
void FolderView::SetFolderPath(const std::optional<std::filesystem::path>& folderPath)
{
    _currentFolder = folderPath;
}
void FolderView::EnsureVisible(size_t)
{
}
void FolderView::ProcessIconLoadQueue()
{
}
void FolderView::QueueIconLoading()
{
}
void FolderView::ReportError(const std::wstring&, HRESULT) const
{
}

std::optional<size_t> FolderView::HitTest(POINT clientPt) const
{
    const float x = DipFromPx(clientPt.x) + _horizontalOffset;
    const float y = DipFromPx(clientPt.y) + _scrollOffset;
    if (_columnCounts.empty() || _tileWidthDip <= 0.0f || _tileHeightDip <= 0.0f)
    {
        for (size_t i = 0; i < _items.size(); ++i)
        {
            const auto& item = _items[i];
            if (x >= item.bounds.left && x <= item.bounds.right && y >= item.bounds.top && y <= item.bounds.bottom)
            {
                return i;
            }
        }
        return std::nullopt;
    }

    const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
    const float columnStride  = _tileWidthDip + kColumnSpacingDip;
    const float rowStride     = _tileHeightDip + rowSpacingDip;
    if (columnStride <= 0.0f || rowStride <= 0.0f)
    {
        return std::nullopt;
    }

    const float firstColumnLeft = kColumnSpacingDip;
    const float firstRowTop     = rowSpacingDip;
    if (x < firstColumnLeft || y < firstRowTop)
    {
        return std::nullopt;
    }

    const int column = static_cast<int>(std::floor((x - firstColumnLeft) / columnStride));
    if (column < 0 || column >= static_cast<int>(_columnCounts.size()))
    {
        return std::nullopt;
    }

    const float columnLeft = firstColumnLeft + static_cast<float>(column) * columnStride;
    if (x > columnLeft + _tileWidthDip)
    {
        return std::nullopt;
    }

    const int row = static_cast<int>(std::floor((y - firstRowTop) / rowStride));
    if (row < 0 || row >= _columnCounts[static_cast<size_t>(column)])
    {
        return std::nullopt;
    }

    const float rowTop = firstRowTop + static_cast<float>(row) * rowStride;
    if (y > rowTop + _tileHeightDip)
    {
        return std::nullopt;
    }

    const size_t columnIndex = static_cast<size_t>(column);
    if (columnIndex >= _columnPrefixSums.size())
    {
        return std::nullopt;
    }

    const size_t index = _columnPrefixSums[columnIndex] + static_cast<size_t>(row);
    if (index >= _items.size())
    {
        return std::nullopt;
    }

    return index;
}

unsigned int StableHash32(std::wstring_view) noexcept
{
    return 0u;
}

unsigned int AppendStableHash32(unsigned int hash, std::wstring_view text) noexcept
{
    static constexpr unsigned int kFnvPrime32 = 16777619u;
    for (const wchar_t ch : text)
    {
        const uint16_t value = static_cast<uint16_t>(ch);

        hash ^= static_cast<uint8_t>(value & 0xFFu);
        hash *= kFnvPrime32;

        hash ^= static_cast<uint8_t>((value >> 8) & 0xFFu);
        hash *= kFnvPrime32;
    }
    return hash;
}

namespace StartupMetrics
{
void MarkFirstPanePopulated(std::wstring_view, uint64_t)
{
}
} // namespace StartupMetrics

bool MaskSyntax::MatchesWildcardMask(std::wstring_view, const WildcardMask&) noexcept
{
    return false;
}

DirectoryInfoCache& DirectoryInfoCache::GetInstance()
{
    static DirectoryInfoCache instance;
    return instance;
}

DirectoryInfoCache::Borrowed::Borrowed(Borrowed&& other) noexcept                                = default;
DirectoryInfoCache::Borrowed& DirectoryInfoCache::Borrowed::operator=(Borrowed&& other) noexcept = default;
DirectoryInfoCache::Borrowed::~Borrowed()                                                        = default;
HRESULT DirectoryInfoCache::Borrowed::Status() const noexcept
{
    return _status;
}
IFilesInformation* DirectoryInfoCache::Borrowed::Get() const noexcept
{
    return _entry ? _entry->info.get() : nullptr;
}
const std::wstring& DirectoryInfoCache::Borrowed::NormalizedPath() const noexcept
{
    static const std::wstring empty;
    return _entry ? _entry->key.path : empty;
}

DirectoryInfoCache::Pin::Pin(Pin&& other) noexcept                                = default;
DirectoryInfoCache::Pin& DirectoryInfoCache::Pin::operator=(Pin&& other) noexcept = default;
DirectoryInfoCache::Pin::~Pin()                                                   = default;
bool DirectoryInfoCache::Pin::IsValid() const noexcept
{
    return _entry != nullptr;
}
const std::wstring& DirectoryInfoCache::Pin::NormalizedPath() const noexcept
{
    static const std::wstring empty;
    return _entry ? _entry->key.path : empty;
}

FolderWatcher::~FolderWatcher() = default;

DirectoryInfoCache::Borrowed DirectoryInfoCache::BorrowDirectoryInfo(IFileSystem* fileSystem, const std::filesystem::path& folder, BorrowMode) noexcept
{
    Borrowed borrowed{};
    borrowed._owner  = this;
    borrowed._status = E_FAIL;
    if (! fileSystem)
    {
        return borrowed;
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = fileSystem->ReadDirectoryInfo(folder.c_str(), info.put());
    borrowed._status = hr;
    if (FAILED(hr) || ! info)
    {
        return borrowed;
    }

    auto entry      = std::make_shared<Entry>();
    entry->info     = std::move(info);
    entry->key.path = folder.wstring();
    borrowed._entry = std::move(entry);
    return borrowed;
}

DirectoryInfoCache::Borrowed DirectoryInfoCache::BorrowDirectoryInfo(IFileSystem* fileSystem,
                                                                     const std::filesystem::path& folder,
                                                                     BorrowMode mode,
                                                                     std::stop_token) noexcept
{
    return BorrowDirectoryInfo(fileSystem, folder, mode);
}

namespace RedSalamanderTests
{
TEST_CLASS(FolderViewRefreshDuplicatePathPerfTest)
{
public:
    static std::filesystem::path s_root;
    static wil::com_ptr<IFileSystem> s_fileSystem;

#pragma warning(push)
#pragma warning(disable : 5246) // CppUnitTest TEST_CLASS_* macros expand to framework-owned registration initializers.
    TEST_CLASS_INITIALIZE(ClassInitialize)
    {
        std::error_code ec;
        s_root = GetSharedDuplicatePluginFolderRoot(ec);
        Assert::IsFalse(static_cast<bool>(ec), L"Failed to create PerformanceTests2 duplicate-path refresh TestSandbox root.");
        Assert::IsFalse(s_root.empty(), L"PerformanceTests2 duplicate-path refresh TestSandbox root is empty.");
        EnsurePhysicalDuplicatePluginFolder(s_root);

        std::vector<DuplicatePathFileSystem::Entry> entries;
        PopulateDuplicatePluginEntries(entries);
        s_fileSystem.attach(new DuplicatePathFileSystem(std::move(entries)));
    }

    TEST_CLASS_CLEANUP(ClassCleanup)
    {
        s_fileSystem.reset();
        std::error_code ec;
        std::filesystem::remove_all(s_root, ec);
        s_root.clear();
    }
#pragma warning(pop)

    TEST_METHOD(FolderViewRefresh_PluginDuplicatePaths)
    {
        Assert::IsTrue(s_fileSystem != nullptr, L"Test file system was not created.");

        FolderView view;
        view._fileSystem = s_fileSystem.get();

        constexpr int kRefreshIterations = 400;

        constexpr uint64_t kFnvOffset = 1469598103934665603ull;
        constexpr uint64_t kFnvPrime  = 1099511628211ull;
        uint64_t checksum             = kFnvOffset;
        size_t expectedItemCount      = 0u;
        bool sawItems                 = false;
        for (int iteration = 0; iteration < kRefreshIterations; ++iteration)
        {
            IconCache::GetInstance().Clear();

            const uint64_t generation = static_cast<uint64_t>(iteration + 1);
            view._enumerationGeneration.store(generation, std::memory_order_release);
            auto payload = view.ExecuteEnumeration(s_root, generation, {});
            if (! payload || payload->status != S_OK)
            {
                checksum = kFnvOffset;
                sawItems = false;
                break;
            }

            if (iteration == 0)
            {
                expectedItemCount = payload->items.size();
            }
            else if (payload->items.size() != expectedItemCount)
            {
                checksum = kFnvOffset;
                sawItems = false;
                break;
            }

            for (const auto& item : payload->items)
            {
                checksum ^= static_cast<uint64_t>(item.iconIndex + 1);
                checksum *= kFnvPrime;
                checksum ^= static_cast<uint64_t>(item.stableHash32);
                checksum *= kFnvPrime;
                sawItems = true;
            }
        }

        Assert::IsTrue(sawItems && checksum != kFnvOffset, L"FolderView refresh produced an unexpected checksum.");
    }

    TEST_METHOD(FolderViewCompactMode_SetAppThemeCollapsesRowGapAndUpdatesHitTest)
    {
        const auto coinit = wil::CoInitializeEx(COINIT_APARTMENTTHREADED);

        FolderView view;
        view._clientSize = SIZE{400, 400};
        view._items.clear();
        view._items.push_back(FolderView::FolderItem{.displayName = L"alpha", .stableHash32 = StableHash32(L"alpha")});
        view._items.push_back(FolderView::FolderItem{.displayName = L"beta", .stableHash32 = StableHash32(L"beta")});

        view.LayoutItems();

        Assert::AreEqual(static_cast<size_t>(2), view._items.size(), L"test requires two laid out rows");

        const FolderView::FolderItem& firstItem  = view._items[0];
        const FolderView::FolderItem& secondItem = view._items[1];
        Assert::IsTrue(secondItem.bounds.top > firstItem.bounds.bottom, L"standard layout should leave a vertical gap between rows");

        const POINT gapPoint = {
            view.PxFromDip(firstItem.bounds.left + 1.0f),
            view.PxFromDip(firstItem.bounds.bottom + 1.0f),
        };
        Assert::IsFalse(view.HitTest(gapPoint).has_value(), L"standard layout should treat the row gap as empty space");

        AppTheme compactTheme    = view._appTheme;
        compactTheme.compactMode = true;
        view.SetAppTheme(compactTheme);

        Assert::AreEqual(view._items[0].bounds.bottom, view._items[1].bounds.top, 0.01f, L"compact layout should collapse the vertical row gap to zero");

        const auto compactHit = view.HitTest(gapPoint);
        Assert::IsTrue(compactHit.has_value(), L"compact layout should let hit testing reach the next row immediately below the previous one");
        Assert::AreEqual(static_cast<size_t>(1), compactHit.value(), L"compact layout should hit the second row at the old gap point");
    }
};

std::filesystem::path FolderViewRefreshDuplicatePathPerfTest::s_root;
wil::com_ptr<IFileSystem> FolderViewRefreshDuplicatePathPerfTest::s_fileSystem;
} // namespace RedSalamanderTests

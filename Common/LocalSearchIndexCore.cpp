#include "LocalSearchIndexCore.h"

#include "Helpers.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cstddef>
#include <cstring>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <span>
#include <string_view>
#include <tuple>
#include <unordered_set>

#include <winioctl.h>
#include <winternl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace LocalSearchIndexCore
{
namespace
{
constexpr HRESULT kNotSupportedHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);

constexpr uint32_t kSnapshotMagic   = 0x58444953u; // "SIDX"
constexpr uint32_t kSnapshotVersion = 1u;
constexpr DWORD kJournalReplayReasons = USN_REASON_FILE_CREATE | USN_REASON_FILE_DELETE | USN_REASON_RENAME_OLD_NAME | USN_REASON_RENAME_NEW_NAME |
                                        USN_REASON_BASIC_INFO_CHANGE | USN_REASON_HARD_LINK_CHANGE | USN_REASON_REPARSE_POINT_CHANGE;

struct NodeId final
{
    uint64_t low  = 0u;
    uint64_t high = 0u;

    [[nodiscard]] bool IsZero() const noexcept
    {
        return low == 0u && high == 0u;
    }

    friend bool operator==(const NodeId&, const NodeId&) noexcept = default;
};

struct NodeIdHash final
{
    [[nodiscard]] size_t operator()(const NodeId& id) const noexcept
    {
        const uint64_t mixed = id.low ^ (id.high + 0x9E3779B97F4A7C15ull + (id.low << 6u) + (id.low >> 2u));
        return static_cast<size_t>(mixed ^ (mixed >> 32u));
    }
};

struct SnapshotHeader final
{
    uint32_t magic          = kSnapshotMagic;
    uint32_t version        = kSnapshotVersion;
    uint32_t fileSystemKind = static_cast<uint32_t>(FileSystemKind::Unsupported);
    uint32_t reserved       = 0u;
    uint64_t journalId      = 0u;
    uint64_t nextUsn        = 0u;
    uint64_t entryCount     = 0u;
    uint64_t rootIdLow      = 0u;
    uint64_t rootIdHigh     = 0u;
};

struct SnapshotEntryHeader final
{
    uint64_t idLow          = 0u;
    uint64_t idHigh         = 0u;
    uint64_t parentIdLow    = 0u;
    uint64_t parentIdHigh   = 0u;
    uint32_t fileAttributes = 0u;
    uint32_t nameBytes      = 0u;
};

struct JournalState final
{
    bool available    = false;
    uint64_t id       = 0u;
    uint64_t firstUsn = 0u;
    uint64_t nextUsn  = 0u;
};

struct Entry final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    std::wstring fullPath;
    unsigned long fileAttributes = 0u;
    std::vector<NodeId> children;
};

struct SeedEntry final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    unsigned long fileAttributes = 0u;
};

struct EnumeratedChild final
{
    std::wstring name;
    std::wstring fullPath;
    unsigned long fileAttributes = 0u;
};

struct UsnRecordData final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    unsigned long fileAttributes = 0u;
    uint32_t reason              = 0u;
};

enum class NtFileInformationClass : int
{
    FileDirectoryInformation     = 1,
    FileFullDirectoryInformation = 2,
};

using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(HANDLE FileHandle,
                                                HANDLE Event,
                                                PIO_APC_ROUTINE ApcRoutine,
                                                PVOID ApcContext,
                                                PIO_STATUS_BLOCK IoStatusBlock,
                                                PVOID FileInformation,
                                                ULONG Length,
                                                NtFileInformationClass FileInformationClass,
                                                BOOLEAN ReturnSingleEntry,
                                                PUNICODE_STRING FileName,
                                                BOOLEAN RestartScan);

using RtlNtStatusToDosError_t = ULONG(NTAPI*)(NTSTATUS Status);

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

#ifndef STATUS_NO_MORE_FILES
#define STATUS_NO_MORE_FILES ((NTSTATUS)0x80000006L)
#endif

} // namespace

struct VolumeIndex final
{
    VolumeIndex()                               = default;
    VolumeIndex(const VolumeIndex&)             = delete;
    VolumeIndex& operator=(const VolumeIndex&)  = delete;
    VolumeIndex(VolumeIndex&&)                  = delete;
    VolumeIndex& operator=(VolumeIndex&&)       = delete;

    std::mutex mutex;
    std::wstring normalizedRootPath;
    std::wstring rootKey;
    std::wstring volumeRoot;
    std::wstring volumeDevicePath;
    std::wstring snapshotPath;
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    NodeId trackedRootId{};
    bool trackedRootIsDirectory = false;
    uint64_t journalId          = 0u;
    uint64_t nextUsn            = 0u;
    bool initialized            = false;
    std::unordered_map<NodeId, Entry, NodeIdHash> entries;
    std::unordered_map<std::wstring, NodeId> pathIndex;
};

namespace
{
[[nodiscard]] std::wstring FoldText(std::wstring_view text) noexcept
{
    std::wstring folded(text);
    if (! folded.empty())
    {
        static_cast<void>(::CharLowerBuffW(folded.data(), static_cast<DWORD>(folded.size())));
    }
    return folded;
}

[[nodiscard]] bool IsDriveRoot(std::wstring_view path) noexcept
{
    return path.size() == 3u && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/');
}

[[nodiscard]] bool IsExtendedDriveRoot(std::wstring_view path) noexcept
{
    return path.size() == 7u && path.rfind(L"\\\\?\\", 0) == 0 && path[5] == L':' && (path[6] == L'\\' || path[6] == L'/');
}

[[nodiscard]] bool IsExtendedUncPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\?\\UNC\\", 0) == 0;
}

[[nodiscard]] bool IsUncPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\", 0) == 0 && ! IsExtendedDriveRoot(path);
}

[[nodiscard]] std::wstring MakeAbsolutePath(std::wstring_view path) noexcept
{
    std::wstring input(path);
    if (input.empty())
    {
        input = L".";
    }

    if (input.rfind(L"\\\\?\\", 0) == 0)
    {
        return input;
    }

    const DWORD required = ::GetFullPathNameW(input.c_str(), 0, nullptr, nullptr);
    if (required == 0)
    {
        return input;
    }

    std::wstring absolute(static_cast<size_t>(required) + 1u, L'\0');
    const DWORD written = ::GetFullPathNameW(input.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
    if (written == 0)
    {
        return input;
    }

    absolute.resize(static_cast<size_t>(written));
    return absolute;
}

[[nodiscard]] std::wstring NormalizePath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');
    normalized = MakeAbsolutePath(normalized);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    while (normalized.size() > 1u && (normalized.back() == L'\\' || normalized.back() == L'/'))
    {
        if (IsDriveRoot(normalized) || IsExtendedDriveRoot(normalized) || normalized == L"\\\\")
        {
            break;
        }

        normalized.pop_back();
    }

    return normalized;
}

[[nodiscard]] std::wstring ToExtendedPath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    if (normalized.empty())
    {
        normalized = L".";
    }

    if (normalized.rfind(L"\\\\?\\", 0) != 0)
    {
        normalized = NormalizePath(normalized);
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0)
    {
        return normalized;
    }

    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return std::wstring(L"\\\\?\\UNC\\") + normalized.substr(2u);
    }

    return std::wstring(L"\\\\?\\") + normalized;
}

[[nodiscard]] std::wstring AppendPath(std::wstring_view basePath, std::wstring_view leafName) noexcept
{
    if (basePath.empty())
    {
        return std::wstring(leafName);
    }

    std::wstring result(basePath);
    const wchar_t last = result.back();
    if (last != L'\\' && last != L'/')
    {
        result.push_back(L'\\');
    }
    result.append(leafName);
    return result;
}

[[nodiscard]] std::wstring GetPathLeaf(std::wstring_view path) noexcept
{
    const size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return std::wstring(path);
    }
    return std::wstring(path.substr(pos + 1u));
}

[[nodiscard]] std::wstring FoldPathKey(std::wstring_view path) noexcept
{
    std::wstring key = NormalizePath(path);
    if (! key.empty())
    {
        static_cast<void>(::CharLowerBuffW(key.data(), static_cast<DWORD>(key.size())));
    }
    return key;
}

[[nodiscard]] uint64_t Fnv1a64(std::wstring_view text) noexcept
{
    uint64_t hash = 14695981039346656037ull;
    for (const wchar_t ch : text)
    {
        hash ^= static_cast<uint64_t>(ch);
        hash *= 1099511628211ull;
    }
    return hash;
}

[[nodiscard]] std::wstring GetDefaultSnapshotRootDirectory() noexcept
{
    DWORD required = ::GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring root(static_cast<size_t>(required), L'\0');
    const DWORD written = ::GetEnvironmentVariableW(L"LOCALAPPDATA", root.data(), required);
    if (written == 0u)
    {
        return {};
    }

    if (! root.empty() && root.back() == L'\0')
    {
        root.pop_back();
    }

    if (root.empty())
    {
        return {};
    }

    return AppendPath(AppendPath(root, L"RedSalamander"), L"SearchIndex");
}

[[nodiscard]] std::wstring GetSnapshotRootDirectory(const RepositoryOptions& options) noexcept
{
    if (! options.snapshotRootDirectory.empty())
    {
        return NormalizePath(options.snapshotRootDirectory);
    }

    return GetDefaultSnapshotRootDirectory();
}

[[nodiscard]] std::wstring BuildSnapshotPath(std::wstring_view normalizedRootPath,
                                             FileSystemKind kind,
                                             const RepositoryOptions& options) noexcept
{
    const std::wstring folder = GetSnapshotRootDirectory(options);
    if (folder.empty())
    {
        return {};
    }

    const uint64_t hash = Fnv1a64(FoldText(normalizedRootPath));
    return AppendPath(folder, std::format(L"{:08x}-{:016x}.bin", static_cast<uint32_t>(kind), static_cast<unsigned long long>(hash)));
}

[[nodiscard]] uint32_t ClampDurationMs(const std::chrono::steady_clock::duration& duration) noexcept
{
    const auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
    if (ms <= 0)
    {
        return 0u;
    }

    if (ms >= static_cast<std::chrono::milliseconds::rep>((std::numeric_limits<uint32_t>::max)()))
    {
        return (std::numeric_limits<uint32_t>::max)();
    }

    return static_cast<uint32_t>(ms);
}

[[nodiscard]] uint64_t GetSnapshotFileBytes(std::wstring_view snapshotPath) noexcept
{
    if (snapshotPath.empty())
    {
        return 0u;
    }

    std::error_code ec;
    const auto fileSize = std::filesystem::file_size(std::filesystem::path(snapshotPath), ec);
    return ec ? 0u : static_cast<uint64_t>(fileSize);
}

[[nodiscard]] uint64_t EstimateVolumeMemoryBytes(const VolumeIndex& volume) noexcept
{
    uint64_t total = sizeof(VolumeIndex);
    total += static_cast<uint64_t>(volume.normalizedRootPath.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.volumeRoot.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.volumeDevicePath.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.snapshotPath.capacity() * sizeof(wchar_t));

    for (const auto& [id, entry] : volume.entries)
    {
        UNREFERENCED_PARAMETER(id);
        total += sizeof(entry);
        total += static_cast<uint64_t>(entry.name.capacity() * sizeof(wchar_t));
        total += static_cast<uint64_t>(entry.fullPath.capacity() * sizeof(wchar_t));
        total += static_cast<uint64_t>(entry.children.capacity() * sizeof(NodeId));
    }

    total += static_cast<uint64_t>(volume.pathIndex.size()) * (sizeof(NodeId) + sizeof(std::wstring));
    for (const auto& [path, id] : volume.pathIndex)
    {
        UNREFERENCED_PARAMETER(id);
        total += static_cast<uint64_t>(path.capacity() * sizeof(wchar_t));
    }

    return total;
}

HRESULT EnsureSnapshotDirectory(std::wstring_view snapshotPath) noexcept
{
    if (snapshotPath.empty())
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path snapshotFile(snapshotPath);
    const std::filesystem::path parent = snapshotFile.parent_path();
    if (parent.empty())
    {
        return E_INVALIDARG;
    }

    std::error_code ec;
    std::filesystem::create_directories(parent, ec);
    if (ec)
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
    }

    return S_OK;
}

HRESULT CheckCancelled(CancelCheckFn cancelCheck, void* cookie) noexcept
{
    if (cancelCheck == nullptr)
    {
        return S_OK;
    }

    const HRESULT hr = cancelCheck(cookie);
    return FAILED(hr) ? hr : S_OK;
}

[[nodiscard]] NodeId NodeIdFromFileId128(const FILE_ID_128& fileId) noexcept
{
    NodeId id{};
    static_assert(sizeof(fileId.Identifier) >= sizeof(id), "FILE_ID_128 is smaller than expected.");
    std::memcpy(&id, fileId.Identifier, sizeof(id));
    return id;
}

[[nodiscard]] NodeId NodeIdFromUint64(uint64_t value) noexcept
{
    return NodeId{value, 0u};
}

[[nodiscard]] bool IsDirectoryAttributes(unsigned long attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
}

NtQueryDirectoryFile_t GetNtQueryDirectoryFile() noexcept
{
    static const NtQueryDirectoryFile_t fn = []() noexcept -> NtQueryDirectoryFile_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<NtQueryDirectoryFile_t>(::GetProcAddress(ntdll, "NtQueryDirectoryFile"));
#pragma warning(pop)
    }();

    return fn;
}

RtlNtStatusToDosError_t GetRtlNtStatusToDosError() noexcept
{
    static const RtlNtStatusToDosError_t fn = []() noexcept -> RtlNtStatusToDosError_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<RtlNtStatusToDosError_t>(::GetProcAddress(ntdll, "RtlNtStatusToDosError"));
#pragma warning(pop)
    }();

    return fn;
}

HRESULT OpenVolumeHandle(const std::wstring& volumeDevicePath, wil::unique_handle& outHandle) noexcept
{
    constexpr std::array<DWORD, 3> desiredAccesses = {{GENERIC_READ, FILE_READ_ATTRIBUTES, 0u}};
    for (const DWORD desiredAccess : desiredAccesses)
    {
        outHandle.reset(::CreateFileW(volumeDevicePath.c_str(),
                                      desiredAccess,
                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                      nullptr,
                                      OPEN_EXISTING,
                                      FILE_ATTRIBUTE_NORMAL,
                                      nullptr));
        if (outHandle)
        {
            return S_OK;
        }
    }

    return HRESULT_FROM_WIN32(::GetLastError());
}

HRESULT OpenPathHandle(const std::wstring& path, wil::unique_handle& outHandle) noexcept
{
    outHandle.reset(::CreateFileW(ToExtendedPath(path).c_str(),
                                  FILE_READ_ATTRIBUTES,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                  nullptr,
                                  OPEN_EXISTING,
                                  FILE_FLAG_BACKUP_SEMANTICS,
                                  nullptr));
    if (! outHandle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT GetPathNodeId(const std::wstring& path, NodeId& outId) noexcept
{
    outId = {};

    wil::unique_handle handle;
    HRESULT hr = OpenPathHandle(path, handle);
    if (FAILED(hr))
    {
        return hr;
    }

    FILE_ID_INFO info{};
    if (::GetFileInformationByHandleEx(handle.get(), FileIdInfo, &info, sizeof(info)) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outId = NodeIdFromFileId128(info.FileId);
    return S_OK;
}

HRESULT GetJournalState(const VolumeIndex& volume, JournalState& outState) noexcept
{
    outState = {};

    wil::unique_handle handle;
    HRESULT hr = OpenVolumeHandle(volume.volumeDevicePath, handle);
    if (FAILED(hr))
    {
        return hr;
    }

    USN_JOURNAL_DATA_V0 journalData{};
    DWORD bytesReturned = 0u;
    if (! ::DeviceIoControl(handle.get(),
                            FSCTL_QUERY_USN_JOURNAL,
                            nullptr,
                            0u,
                            &journalData,
                            sizeof(journalData),
                            &bytesReturned,
                            nullptr))
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_JOURNAL_NOT_ACTIVE || error == ERROR_INVALID_FUNCTION)
        {
            return kNotSupportedHr;
        }

        return HRESULT_FROM_WIN32(error);
    }

    outState.available = true;
    outState.id        = static_cast<uint64_t>(journalData.UsnJournalID);
    outState.firstUsn  = static_cast<uint64_t>(journalData.FirstUsn);
    outState.nextUsn   = static_cast<uint64_t>(journalData.NextUsn);
    return S_OK;
}

[[nodiscard]] bool TryParseUsnRecord(const USN_RECORD_COMMON_HEADER* header, UsnRecordData& out) noexcept
{
    out = {};
    if (header == nullptr)
    {
        return false;
    }

    if (header->MajorVersion == 2u)
    {
        const auto* record = reinterpret_cast<const USN_RECORD_V2*>(header);
        const size_t nameChars = static_cast<size_t>(record->FileNameLength / sizeof(wchar_t));
        out.id                 = NodeIdFromUint64(static_cast<uint64_t>(record->FileReferenceNumber));
        out.parentId           = NodeIdFromUint64(static_cast<uint64_t>(record->ParentFileReferenceNumber));
        out.name.assign(record->FileName, nameChars);
        out.fileAttributes = record->FileAttributes;
        out.reason         = record->Reason;
        return true;
    }

    if (header->MajorVersion == 3u)
    {
        const auto* record = reinterpret_cast<const USN_RECORD_V3*>(header);
        const size_t nameChars = static_cast<size_t>(record->FileNameLength / sizeof(wchar_t));
        out.id                 = NodeIdFromFileId128(record->FileReferenceNumber);
        out.parentId           = NodeIdFromFileId128(record->ParentFileReferenceNumber);
        out.name.assign(record->FileName, nameChars);
        out.fileAttributes = record->FileAttributes;
        out.reason         = record->Reason;
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring ExtractVolumeRoot(const std::wstring& normalizedRootPath) noexcept
{
    std::array<wchar_t, MAX_PATH> buffer{};
    if (::GetVolumePathNameW(normalizedRootPath.c_str(), buffer.data(), static_cast<DWORD>(buffer.size())) == 0)
    {
        return {};
    }

    std::wstring volumeRoot(buffer.data());
    std::replace(volumeRoot.begin(), volumeRoot.end(), L'/', L'\\');
    return volumeRoot;
}

[[nodiscard]] std::wstring BuildVolumeDevicePath(std::wstring_view volumeRoot) noexcept
{
    if (volumeRoot.size() < 2u || volumeRoot[1] != L':')
    {
        return {};
    }

    return std::wstring(L"\\\\.\\") + std::wstring(volumeRoot.substr(0u, 2u));
}

[[nodiscard]] bool EqualsCaseInsensitive(std::wstring_view left, std::wstring_view right) noexcept
{
    return FoldText(left) == FoldText(right);
}

[[nodiscard]] std::vector<std::wstring> SplitRelativeComponents(std::wstring_view basePath, std::wstring_view childPath) noexcept
{
    std::vector<std::wstring> components;
    if (childPath.size() < basePath.size())
    {
        return components;
    }

    std::wstring_view remainder = childPath.substr(basePath.size());
    while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
    {
        remainder.remove_prefix(1u);
    }

    while (! remainder.empty())
    {
        const size_t separator = remainder.find_first_of(L"\\/");
        if (separator == std::wstring_view::npos)
        {
            components.emplace_back(remainder);
            break;
        }

        components.emplace_back(remainder.substr(0u, separator));
        remainder.remove_prefix(separator + 1u);
        while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
        {
            remainder.remove_prefix(1u);
        }
    }

    return components;
}

HRESULT PopulateSupportInfo(std::wstring_view rootPath, SupportInfo& outSupport) noexcept
{
    outSupport = {};

    const std::wstring normalized = NormalizePath(rootPath);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    outSupport.normalizedRootPath = normalized;

    if (IsUncPath(normalized) || IsExtendedUncPath(normalized))
    {
        return S_OK;
    }

    const std::wstring volumeRoot = ExtractVolumeRoot(normalized);
    if (volumeRoot.empty())
    {
        return S_OK;
    }

    std::array<wchar_t, 64> fileSystemName{};
    if (::GetVolumeInformationW(volumeRoot.c_str(), nullptr, 0u, nullptr, nullptr, nullptr, fileSystemName.data(), static_cast<DWORD>(fileSystemName.size())) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const std::wstring fileSystem = FoldText(fileSystemName.data());
    if (fileSystem == L"ntfs")
    {
        outSupport.fileSystemKind = FileSystemKind::Ntfs;
        outSupport.indexable      = true;
    }
    else if (fileSystem == L"refs")
    {
        outSupport.fileSystemKind = FileSystemKind::Refs;
        outSupport.indexable      = true;
    }

    return S_OK;
}

void PopulateStatsFromVolume(const VolumeIndex& volume, QueryStats& stats) noexcept
{
    stats.fileSystemKind = volume.fileSystemKind;
    stats.snapshotPath   = volume.snapshotPath;
    stats.nextUsn        = volume.nextUsn;
    stats.journalId      = volume.journalId;
    stats.snapshotFileBytes    = GetSnapshotFileBytes(volume.snapshotPath);
    stats.estimatedMemoryBytes = EstimateVolumeMemoryBytes(volume);
    stats.entryCount     = 0u;
    stats.fileCount      = 0u;
    stats.directoryCount = 0u;

    for (const auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId && volume.trackedRootIsDirectory)
        {
            continue;
        }

        ++stats.entryCount;
        if (IsDirectoryAttributes(entry.fileAttributes))
        {
            ++stats.directoryCount;
        }
        else
        {
            ++stats.fileCount;
        }
    }
}

void RemoveSubtree(VolumeIndex& volume, const NodeId& id) noexcept
{
    const auto it = volume.entries.find(id);
    if (it == volume.entries.end())
    {
        return;
    }

    const std::vector<NodeId> children = it->second.children;
    for (const NodeId& childId : children)
    {
        RemoveSubtree(volume, childId);
    }

    volume.entries.erase(id);
}

void RebuildDerivedState(VolumeIndex& volume) noexcept
{
    for (auto& [id, entry] : volume.entries)
    {
        entry.children.clear();
        if (id != volume.trackedRootId)
        {
            entry.fullPath.clear();
        }
    }

    for (auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId)
        {
            continue;
        }

        if (const auto parent = volume.entries.find(entry.parentId); parent != volume.entries.end())
        {
            parent->second.children.push_back(id);
        }
    }

    const auto compareChildren = [&](const NodeId& left, const NodeId& right) noexcept
    {
        const auto leftIt  = volume.entries.find(left);
        const auto rightIt = volume.entries.find(right);
        if (leftIt == volume.entries.end() || rightIt == volume.entries.end())
        {
            return left.low < right.low || (left.low == right.low && left.high < right.high);
        }

        return std::tie(leftIt->second.name, left.low, left.high) < std::tie(rightIt->second.name, right.low, right.high);
    };

    std::unordered_set<NodeId, NodeIdHash> visited;
    std::vector<NodeId> stack;

    const auto rootIt = volume.entries.find(volume.trackedRootId);
    if (rootIt == volume.entries.end())
    {
        volume.pathIndex.clear();
        return;
    }

    rootIt->second.fullPath = volume.normalizedRootPath;
    stack.push_back(volume.trackedRootId);

    while (! stack.empty())
    {
        const NodeId currentId = stack.back();
        stack.pop_back();

        if (! visited.insert(currentId).second)
        {
            continue;
        }

        auto currentIt = volume.entries.find(currentId);
        if (currentIt == volume.entries.end())
        {
            continue;
        }

        auto& current = currentIt->second;
        std::sort(current.children.begin(), current.children.end(), compareChildren);

        for (auto childIt = current.children.rbegin(); childIt != current.children.rend(); ++childIt)
        {
            auto found = volume.entries.find(*childIt);
            if (found == volume.entries.end())
            {
                continue;
            }

            found->second.fullPath = AppendPath(current.fullPath, found->second.name);
            stack.push_back(*childIt);
        }
    }

    std::vector<NodeId> unreachable;
    for (const auto& [id, entry] : volume.entries)
    {
        if (! visited.contains(id))
        {
            unreachable.push_back(id);
        }
    }

    for (const NodeId& id : unreachable)
    {
        volume.entries.erase(id);
    }

    for (auto& [id, entry] : volume.entries)
    {
        entry.children.clear();
    }
    for (auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId)
        {
            continue;
        }

        if (const auto parent = volume.entries.find(entry.parentId); parent != volume.entries.end())
        {
            parent->second.children.push_back(id);
        }
    }
    for (auto& [id, entry] : volume.entries)
    {
        std::sort(entry.children.begin(), entry.children.end(), compareChildren);
    }

    volume.pathIndex.clear();
    for (const auto& [id, entry] : volume.entries)
    {
        volume.pathIndex.emplace(FoldPathKey(entry.fullPath), id);
    }
}

HRESULT SaveSnapshot(const VolumeIndex& volume, QueryStats& stats) noexcept
{
    if (volume.snapshotPath.empty())
    {
        return E_INVALIDARG;
    }

    HRESULT hr = EnsureSnapshotDirectory(volume.snapshotPath);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::unique_handle file(::CreateFileW(volume.snapshotPath.c_str(),
                                          GENERIC_WRITE,
                                          FILE_SHARE_READ,
                                          nullptr,
                                          CREATE_ALWAYS,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    SnapshotHeader header{};
    header.fileSystemKind = static_cast<uint32_t>(volume.fileSystemKind);
    header.journalId      = volume.journalId;
    header.nextUsn        = volume.nextUsn;
    header.entryCount     = static_cast<uint64_t>(volume.entries.size());
    header.rootIdLow      = volume.trackedRootId.low;
    header.rootIdHigh     = volume.trackedRootId.high;

    DWORD written = 0u;
    if (::WriteFile(file.get(), &header, sizeof(header), &written, nullptr) == 0 || written != sizeof(header))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    for (const auto& [id, entry] : volume.entries)
    {
        SnapshotEntryHeader entryHeader{};
        entryHeader.idLow          = entry.id.low;
        entryHeader.idHigh         = entry.id.high;
        entryHeader.parentIdLow    = entry.parentId.low;
        entryHeader.parentIdHigh   = entry.parentId.high;
        entryHeader.fileAttributes = entry.fileAttributes;
        entryHeader.nameBytes      = static_cast<uint32_t>(entry.name.size() * sizeof(wchar_t));

        if (::WriteFile(file.get(), &entryHeader, sizeof(entryHeader), &written, nullptr) == 0 || written != sizeof(entryHeader))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (! entry.name.empty())
        {
            const DWORD nameBytes = static_cast<DWORD>(entry.name.size() * sizeof(wchar_t));
            if (::WriteFile(file.get(), entry.name.data(), nameBytes, &written, nullptr) == 0 || written != nameBytes)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }
        }
    }

    stats.snapshotSaved = true;
    stats.snapshotFileBytes = GetSnapshotFileBytes(volume.snapshotPath);
    return S_OK;
}

HRESULT LoadSnapshot(VolumeIndex& volume, QueryStats& stats) noexcept
{
    stats.snapshotLoaded = false;
    stats.snapshotFileBytes = 0u;

    wil::unique_handle file(::CreateFileW(volume.snapshotPath.c_str(),
                                          GENERIC_READ,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    LARGE_INTEGER size{};
    if (::GetFileSizeEx(file.get(), &size) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    if (size.QuadPart < static_cast<LONGLONG>(sizeof(SnapshotHeader)))
    {
        stats.rebuiltSnapshotCorruption = true;
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    SnapshotHeader header{};
    DWORD bytesRead = 0u;
    if (::ReadFile(file.get(), &header, sizeof(header), &bytesRead, nullptr) == 0 || bytesRead != sizeof(header))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    if (header.magic != kSnapshotMagic || header.version != kSnapshotVersion ||
        header.fileSystemKind != static_cast<uint32_t>(volume.fileSystemKind))
    {
        stats.rebuiltSnapshotCorruption = true;
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    volume.entries.clear();
    volume.trackedRootId = NodeId{header.rootIdLow, header.rootIdHigh};
    volume.journalId     = header.journalId;
    volume.nextUsn       = header.nextUsn;

    for (uint64_t index = 0u; index < header.entryCount; ++index)
    {
        SnapshotEntryHeader entryHeader{};
        if (::ReadFile(file.get(), &entryHeader, sizeof(entryHeader), &bytesRead, nullptr) == 0 || bytesRead != sizeof(entryHeader))
        {
            stats.rebuiltSnapshotCorruption = true;
            volume.entries.clear();
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        if ((entryHeader.nameBytes % sizeof(wchar_t)) != 0u)
        {
            stats.rebuiltSnapshotCorruption = true;
            volume.entries.clear();
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::wstring name;
        if (entryHeader.nameBytes != 0u)
        {
            name.resize(entryHeader.nameBytes / sizeof(wchar_t));
            if (::ReadFile(file.get(), name.data(), entryHeader.nameBytes, &bytesRead, nullptr) == 0 || bytesRead != entryHeader.nameBytes)
            {
                stats.rebuiltSnapshotCorruption = true;
                volume.entries.clear();
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }

        Entry entry{};
        entry.id             = NodeId{entryHeader.idLow, entryHeader.idHigh};
        entry.parentId       = NodeId{entryHeader.parentIdLow, entryHeader.parentIdHigh};
        entry.fileAttributes = entryHeader.fileAttributes;
        entry.name           = std::move(name);
        volume.entries.emplace(entry.id, std::move(entry));
    }

    if (! volume.entries.contains(volume.trackedRootId))
    {
        stats.rebuiltSnapshotCorruption = true;
        volume.entries.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    RebuildDerivedState(volume);
    if (const auto rootIt = volume.entries.find(volume.trackedRootId); rootIt != volume.entries.end())
    {
        volume.trackedRootIsDirectory = IsDirectoryAttributes(rootIt->second.fileAttributes);
    }

    stats.snapshotLoaded = true;
    stats.snapshotFileBytes = static_cast<uint64_t>(size.QuadPart);
    return S_OK;
}

[[nodiscard]] bool MatchWildcardCaseSensitive(std::wstring_view text, std::wstring_view pattern) noexcept
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

[[nodiscard]] bool MatchWildcard(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    return caseSensitive ? MatchWildcardCaseSensitive(text, pattern) : MatchWildcardCaseSensitive(FoldText(text), FoldText(pattern));
}

[[nodiscard]] bool MatchLiteral(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    if (pattern.empty())
    {
        return true;
    }

    if (caseSensitive)
    {
        return text.find(pattern) != std::wstring_view::npos;
    }

    const std::wstring foldedText    = FoldText(text);
    const std::wstring foldedPattern = FoldText(pattern);
    return foldedText.find(foldedPattern) != std::wstring_view::npos;
}

[[nodiscard]] bool MatchName(const QueryPlan& plan, std::wstring_view name) noexcept
{
    switch (plan.nameMode)
    {
        case FILESYSTEM_SEARCH_NAME_DISABLED: return true;
        case FILESYSTEM_SEARCH_NAME_WILDCARD: return MatchWildcard(name, plan.namePattern, plan.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_LITERAL: return MatchLiteral(name, plan.namePattern, plan.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_REGEX: return plan.compiledNameRegex != nullptr && std::regex_search(std::wstring(name), *plan.compiledNameRegex);
    }

    return false;
}

HRESULT EnumerateDirectory(std::wstring_view directoryPath, std::vector<EnumeratedChild>& outChildren) noexcept
{
    outChildren.clear();

    const auto fallbackEnumerate = [&]() noexcept -> HRESULT
    {
        WIN32_FIND_DATAW data{};
        const std::wstring pattern = AppendPath(ToExtendedPath(directoryPath), L"*");
        wil::unique_hfind findHandle(::FindFirstFileW(pattern.c_str(), &data));
        if (! findHandle)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        do
        {
            const std::wstring_view name(data.cFileName);
            if (name == L"." || name == L"..")
            {
                continue;
            }

            EnumeratedChild child{};
            child.name           = std::wstring(name);
            child.fullPath       = AppendPath(directoryPath, child.name);
            child.fileAttributes = data.dwFileAttributes;
            outChildren.push_back(std::move(child));
        } while (::FindNextFileW(findHandle.get(), &data) != 0);

        const DWORD error = ::GetLastError();
        return error == ERROR_NO_MORE_FILES ? S_OK : HRESULT_FROM_WIN32(error);
    };

    wil::unique_handle directory(::CreateFileW(ToExtendedPath(directoryPath).c_str(),
                                               FILE_LIST_DIRECTORY,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               nullptr,
                                               OPEN_EXISTING,
                                               FILE_FLAG_BACKUP_SEMANTICS,
                                               nullptr));
    if (! directory)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const auto NtQueryDirectoryFile = GetNtQueryDirectoryFile();
    if (NtQueryDirectoryFile == nullptr)
    {
        return fallbackEnumerate();
    }

    std::vector<std::byte> buffer(128u * 1024u);
    bool restart = true;
    for (;;)
    {
        IO_STATUS_BLOCK iosb{};
        const NTSTATUS status = NtQueryDirectoryFile(directory.get(),
                                                     nullptr,
                                                     nullptr,
                                                     nullptr,
                                                     &iosb,
                                                     buffer.data(),
                                                     static_cast<ULONG>(buffer.size()),
                                                     NtFileInformationClass::FileFullDirectoryInformation,
                                                     FALSE,
                                                     nullptr,
                                                     restart ? TRUE : FALSE);
        if (status == STATUS_NO_MORE_FILES)
        {
            break;
        }

        if (! NT_SUCCESS(status))
        {
            if (const auto RtlNtStatusToDosError = GetRtlNtStatusToDosError(); RtlNtStatusToDosError != nullptr)
            {
                const DWORD error = RtlNtStatusToDosError(status);
                if (error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED || error == ERROR_INVALID_PARAMETER)
                {
                    return fallbackEnumerate();
                }
                return HRESULT_FROM_WIN32(error != 0u ? error : ERROR_GEN_FAILURE);
            }

            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }

        restart = false;

        const size_t bytesValid = static_cast<size_t>(iosb.Information);
        if (bytesValid == 0u || bytesValid > buffer.size())
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        size_t offset = 0u;
        while (offset < bytesValid)
        {
            if (bytesValid - offset < offsetof(FILE_FULL_DIR_INFO, FileName))
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            const auto* info = reinterpret_cast<const FILE_FULL_DIR_INFO*>(buffer.data() + offset);
            if ((info->FileNameLength % sizeof(wchar_t)) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            const size_t nameChars = info->FileNameLength / sizeof(wchar_t);
            const std::wstring_view name(info->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                EnumeratedChild child{};
                child.name           = std::wstring(name);
                child.fullPath       = AppendPath(directoryPath, child.name);
                child.fileAttributes = info->FileAttributes;
                outChildren.push_back(std::move(child));
            }

            if (info->NextEntryOffset == 0u)
            {
                break;
            }

            offset += static_cast<size_t>(info->NextEntryOffset);
        }
    }

    return S_OK;
}

HRESULT HydrateDirectorySubtree(VolumeIndex& volume,
                                const NodeId& directoryId,
                                std::wstring_view directoryPath,
                                CancelCheckFn cancelCheck,
                                void* cancelCookie,
                                QueryStats& stats) noexcept
{
    HRESULT hr = CheckCancelled(cancelCheck, cancelCookie);
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<EnumeratedChild> children;
    hr = EnumerateDirectory(directoryPath, children);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const EnumeratedChild& child : children)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        NodeId childId{};
        hr = GetPathNodeId(child.fullPath, childId);
        if (FAILED(hr))
        {
            if (hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
            {
                continue;
            }

            return hr;
        }

        Entry entry{};
        entry.id             = childId;
        entry.parentId       = directoryId;
        entry.name           = child.name;
        entry.fileAttributes = child.fileAttributes;
        volume.entries[childId] = std::move(entry);

        if (IsDirectoryAttributes(child.fileAttributes))
        {
            hr = HydrateDirectorySubtree(volume, childId, child.fullPath, cancelCheck, cancelCookie, stats);
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }

    stats.usedTraversalSeed = true;
    return S_OK;
}

HRESULT SeedTraversalIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats) noexcept
{
    volume.entries.clear();

    const DWORD rootAttributes = ::GetFileAttributesW(volume.normalizedRootPath.c_str());
    if (rootAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    NodeId rootId{};
    HRESULT hr = GetPathNodeId(volume.normalizedRootPath, rootId);
    if (FAILED(hr))
    {
        return hr;
    }

    Entry rootEntry{};
    rootEntry.id             = rootId;
    rootEntry.parentId       = {};
    rootEntry.fileAttributes = rootAttributes;
    rootEntry.name           = IsDirectoryAttributes(rootAttributes) ? std::wstring() : GetPathLeaf(volume.normalizedRootPath);

    volume.entries.emplace(rootId, std::move(rootEntry));
    volume.trackedRootId          = rootId;
    volume.trackedRootIsDirectory = IsDirectoryAttributes(rootAttributes);

    if (volume.trackedRootIsDirectory)
    {
        hr = HydrateDirectorySubtree(volume, rootId, volume.normalizedRootPath, cancelCheck, cancelCookie, stats);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    RebuildDerivedState(volume);
    return S_OK;
}

HRESULT SeedNtfsIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats) noexcept
{
    volume.entries.clear();

    NodeId volumeRootId{};
    HRESULT hr = GetPathNodeId(volume.volumeRoot, volumeRootId);
    if (FAILED(hr))
    {
        return hr;
    }

    std::unordered_map<NodeId, SeedEntry, NodeIdHash> allEntries;
    std::unordered_map<NodeId, std::vector<NodeId>, NodeIdHash> childrenByParent;

    SeedEntry volumeRootEntry{};
    volumeRootEntry.id             = volumeRootId;
    volumeRootEntry.parentId       = {};
    volumeRootEntry.fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
    allEntries.emplace(volumeRootId, std::move(volumeRootEntry));

    wil::unique_handle volumeHandle;
    hr = OpenVolumeHandle(volume.volumeDevicePath, volumeHandle);
    if (FAILED(hr))
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || hr == HRESULT_FROM_WIN32(ERROR_PRIVILEGE_NOT_HELD))
        {
            return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats);
        }

        return hr;
    }

    MFT_ENUM_DATA_V0 enumData{};
    enumData.StartFileReferenceNumber = 0u;
    enumData.LowUsn                   = 0u;
    enumData.HighUsn                  = MAXLONGLONG;

    std::vector<std::byte> buffer(1024u * 1024u);
    for (;;)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        DWORD bytesReturned = 0u;
        if (! ::DeviceIoControl(volumeHandle.get(),
                                FSCTL_ENUM_USN_DATA,
                                &enumData,
                                sizeof(enumData),
                                buffer.data(),
                                static_cast<DWORD>(buffer.size()),
                                &bytesReturned,
                                nullptr))
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_HANDLE_EOF || error == ERROR_NO_MORE_FILES)
            {
                break;
            }

            if (error == ERROR_ACCESS_DENIED || error == ERROR_PRIVILEGE_NOT_HELD || error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED)
            {
                return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats);
            }

            return HRESULT_FROM_WIN32(error);
        }

        if (bytesReturned <= sizeof(uint64_t))
        {
            break;
        }

        enumData.StartFileReferenceNumber = *reinterpret_cast<const DWORDLONG*>(buffer.data());

        size_t offset = sizeof(uint64_t);
        while (offset + sizeof(USN_RECORD_COMMON_HEADER) <= bytesReturned)
        {
            const auto* header = reinterpret_cast<const USN_RECORD_COMMON_HEADER*>(buffer.data() + offset);
            if (header->RecordLength < sizeof(USN_RECORD_COMMON_HEADER) || offset + header->RecordLength > bytesReturned)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            UsnRecordData parsed{};
            if (TryParseUsnRecord(header, parsed))
            {
                SeedEntry entry{};
                entry.id             = parsed.id;
                entry.parentId       = parsed.parentId;
                entry.name           = std::move(parsed.name);
                entry.fileAttributes = parsed.fileAttributes;
                allEntries[entry.id] = entry;
                childrenByParent[entry.parentId].push_back(entry.id);
            }

            offset += header->RecordLength;
        }
    }

    const std::vector<std::wstring> components = SplitRelativeComponents(volume.volumeRoot, volume.normalizedRootPath);
    NodeId trackedRootId = volumeRootId;
    unsigned long trackedRootAttributes = FILE_ATTRIBUTE_DIRECTORY;

    for (const std::wstring& component : components)
    {
        const auto childrenIt = childrenByParent.find(trackedRootId);
        if (childrenIt == childrenByParent.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        std::optional<NodeId> matchedId;
        for (const NodeId& childId : childrenIt->second)
        {
            const auto entryIt = allEntries.find(childId);
            if (entryIt == allEntries.end())
            {
                continue;
            }

            if (EqualsCaseInsensitive(entryIt->second.name, component))
            {
                matchedId             = childId;
                trackedRootAttributes = entryIt->second.fileAttributes;
                break;
            }
        }

        if (! matchedId.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        trackedRootId = matchedId.value();
    }

    std::vector<NodeId> subtreeIds;
    subtreeIds.push_back(trackedRootId);
    for (size_t index = 0u; index < subtreeIds.size(); ++index)
    {
        const NodeId currentId = subtreeIds[index];
        const auto childrenIt = childrenByParent.find(currentId);
        if (childrenIt == childrenByParent.end())
        {
            continue;
        }

        subtreeIds.insert(subtreeIds.end(), childrenIt->second.begin(), childrenIt->second.end());
    }

    volume.trackedRootId          = trackedRootId;
    volume.trackedRootIsDirectory = IsDirectoryAttributes(trackedRootAttributes);

    for (const NodeId& id : subtreeIds)
    {
        const auto seedIt = allEntries.find(id);
        if (seedIt == allEntries.end())
        {
            continue;
        }

        Entry entry{};
        entry.id             = seedIt->second.id;
        entry.parentId       = (id == trackedRootId) ? NodeId{} : seedIt->second.parentId;
        entry.name           = seedIt->second.name;
        entry.fileAttributes = seedIt->second.fileAttributes;
        if (id == trackedRootId && ! volume.trackedRootIsDirectory && entry.name.empty())
        {
            entry.name = GetPathLeaf(volume.normalizedRootPath);
        }

        volume.entries.emplace(entry.id, std::move(entry));
    }

    if (! volume.entries.contains(trackedRootId))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    RebuildDerivedState(volume);
    stats.usedNtfsEnumeration = true;
    return S_OK;
}

HRESULT BuildIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats) noexcept
{
    if (volume.fileSystemKind == FileSystemKind::Ntfs)
    {
        return SeedNtfsIndex(volume, cancelCheck, cancelCookie, stats);
    }

    if (volume.fileSystemKind == FileSystemKind::Refs)
    {
        return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats);
    }

    return kNotSupportedHr;
}

HRESULT ReplayJournal(VolumeIndex& volume,
                      const JournalState& journalState,
                      CancelCheckFn cancelCheck,
                      void* cancelCookie,
                      QueryStats& stats) noexcept
{
    wil::unique_handle volumeHandle;
    HRESULT hr = OpenVolumeHandle(volume.volumeDevicePath, volumeHandle);
    if (FAILED(hr))
    {
        return hr;
    }

    READ_USN_JOURNAL_DATA_V0 readData{};
    readData.StartUsn          = static_cast<USN>(volume.nextUsn);
    readData.ReasonMask        = kJournalReplayReasons;
    readData.ReturnOnlyOnClose = FALSE;
    readData.Timeout           = 0u;
    readData.BytesToWaitFor    = 0u;
    readData.UsnJournalID      = journalState.id;

    std::vector<Entry> directoriesToHydrate;
    std::vector<std::byte> buffer(256u * 1024u);
    while (static_cast<uint64_t>(readData.StartUsn) < journalState.nextUsn)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        DWORD bytesReturned = 0u;
        if (! ::DeviceIoControl(volumeHandle.get(),
                                FSCTL_READ_USN_JOURNAL,
                                &readData,
                                sizeof(readData),
                                buffer.data(),
                                static_cast<DWORD>(buffer.size()),
                                &bytesReturned,
                                nullptr))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (bytesReturned < sizeof(USN))
        {
            break;
        }

        const auto nextUsn = *reinterpret_cast<const USN*>(buffer.data());
        readData.StartUsn  = nextUsn;

        size_t offset = sizeof(USN);
        while (offset + sizeof(USN_RECORD_COMMON_HEADER) <= bytesReturned)
        {
            const auto* header = reinterpret_cast<const USN_RECORD_COMMON_HEADER*>(buffer.data() + offset);
            if (header->RecordLength < sizeof(USN_RECORD_COMMON_HEADER) || offset + header->RecordLength > bytesReturned)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            UsnRecordData record{};
            if (TryParseUsnRecord(header, record))
            {
                const bool wasTracked    = volume.entries.contains(record.id);
                const bool parentTracked = volume.entries.contains(record.parentId) || record.parentId == volume.trackedRootId;

                if ((record.reason & USN_REASON_FILE_DELETE) != 0u)
                {
                    RemoveSubtree(volume, record.id);
                }
                else if ((record.reason & (USN_REASON_FILE_CREATE | USN_REASON_RENAME_NEW_NAME | USN_REASON_BASIC_INFO_CHANGE | USN_REASON_HARD_LINK_CHANGE |
                                           USN_REASON_REPARSE_POINT_CHANGE)) != 0u)
                {
                    if (wasTracked || parentTracked || record.id == volume.trackedRootId)
                    {
                        Entry updated{};
                        if (const auto existing = volume.entries.find(record.id); existing != volume.entries.end())
                        {
                            updated = existing->second;
                        }

                        updated.id             = record.id;
                        updated.parentId       = (record.id == volume.trackedRootId) ? NodeId{} : record.parentId;
                        updated.fileAttributes = record.fileAttributes;
                        if (record.id != volume.trackedRootId || ! record.name.empty())
                        {
                            updated.name = record.name;
                        }

                        volume.entries[record.id] = updated;
                        if (IsDirectoryAttributes(updated.fileAttributes) &&
                            ((record.reason & (USN_REASON_FILE_CREATE | USN_REASON_RENAME_NEW_NAME)) != 0u) && ! wasTracked)
                        {
                            directoriesToHydrate.push_back(updated);
                        }
                    }
                    else if (wasTracked)
                    {
                        RemoveSubtree(volume, record.id);
                    }
                }
            }

            offset += header->RecordLength;
        }
    }

    RebuildDerivedState(volume);
    for (const Entry& directory : directoriesToHydrate)
    {
        const auto pathIt = volume.pathIndex.find(FoldPathKey(directory.fullPath));
        if (pathIt == volume.pathIndex.end())
        {
            continue;
        }

        hr = HydrateDirectorySubtree(volume, directory.id, directory.fullPath, cancelCheck, cancelCookie, stats);
        if (FAILED(hr))
        {
            return hr;
        }
        RebuildDerivedState(volume);
    }

    volume.journalId = journalState.id;
    volume.nextUsn   = journalState.nextUsn;
    stats.journalReplayApplied = true;
    return S_OK;
}

HRESULT EnsureReady(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats) noexcept
{
    const auto start = std::chrono::steady_clock::now();
    stats.fileSystemKind = volume.fileSystemKind;
    stats.snapshotPath   = volume.snapshotPath;

    if (! volume.initialized)
    {
        const HRESULT loadHr = LoadSnapshot(volume, stats);
        if (SUCCEEDED(loadHr))
        {
            volume.initialized = true;
        }
        else if (loadHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && loadHr != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) &&
                 loadHr != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
        {
            return loadHr;
        }
    }

    JournalState journalState{};
    HRESULT hr = GetJournalState(volume, journalState);
    if (hr == kNotSupportedHr)
    {
        stats.journalAvailable = false;

        const auto rebuildWithoutJournal = [&]() noexcept -> HRESULT
        {
            volume.entries.clear();
            volume.pathIndex.clear();
            volume.initialized = false;

            HRESULT rebuildHr = BuildIndex(volume, cancelCheck, cancelCookie, stats);
            if (FAILED(rebuildHr))
            {
                return rebuildHr;
            }

            volume.journalId   = 0u;
            volume.nextUsn     = 0u;
            volume.initialized = true;
            RebuildDerivedState(volume);
            return SaveSnapshot(volume, stats);
        };

        if (! volume.initialized)
        {
            hr = rebuildWithoutJournal();
            if (FAILED(hr))
            {
                return hr;
            }
        }
        else
        {
            hr = rebuildWithoutJournal();
            if (FAILED(hr))
            {
                return hr;
            }
        }

        PopulateStatsFromVolume(volume, stats);
        stats.ensureReadyDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - start);
        Debug::Info(L"LocalSearchIndexCore: ready root='{}' kind={} journalAvailable={} replay={} rebuildIdMismatch={} rebuildRange={} rebuildCorruption={} "
                    L"usedNtfs={} usedTraversal={} entries={} files={} dirs={} snapshotBytes={} memoryBytes={} readyMs={}",
                    volume.normalizedRootPath,
                    static_cast<uint32_t>(stats.fileSystemKind),
                    stats.journalAvailable,
                    stats.journalReplayApplied,
                    stats.rebuiltJournalIdMismatch,
                    stats.rebuiltJournalRangeInvalid,
                    stats.rebuiltSnapshotCorruption,
                    stats.usedNtfsEnumeration,
                    stats.usedTraversalSeed,
                    stats.entryCount,
                    stats.fileCount,
                    stats.directoryCount,
                    stats.snapshotFileBytes,
                    stats.estimatedMemoryBytes,
                    stats.ensureReadyDurationMs);
        return S_OK;
    }

    if (FAILED(hr))
    {
        return hr;
    }
    stats.journalAvailable = journalState.available;

    const auto rebuild = [&]() noexcept -> HRESULT
    {
        volume.entries.clear();
        volume.pathIndex.clear();
        volume.initialized = false;

        HRESULT rebuildHr = BuildIndex(volume, cancelCheck, cancelCookie, stats);
        if (FAILED(rebuildHr))
        {
            return rebuildHr;
        }

        volume.journalId   = journalState.id;
        volume.nextUsn     = journalState.nextUsn;
        volume.initialized = true;
        RebuildDerivedState(volume);
        return SaveSnapshot(volume, stats);
    };

    if (! volume.initialized)
    {
        hr = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.journalId != 0u && volume.journalId != journalState.id)
    {
        stats.rebuiltJournalIdMismatch = true;
        hr = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.nextUsn < journalState.firstUsn || volume.nextUsn > journalState.nextUsn)
    {
        stats.rebuiltJournalRangeInvalid = true;
        hr = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.nextUsn < journalState.nextUsn)
    {
        hr = ReplayJournal(volume, journalState, cancelCheck, cancelCookie, stats);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = SaveSnapshot(volume, stats);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else
    {
        volume.journalId = journalState.id;
        volume.nextUsn   = journalState.nextUsn;
    }

    PopulateStatsFromVolume(volume, stats);
    stats.ensureReadyDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - start);
    Debug::Info(L"LocalSearchIndexCore: ready root='{}' kind={} journalAvailable={} replay={} rebuildIdMismatch={} rebuildRange={} rebuildCorruption={} "
                L"usedNtfs={} usedTraversal={} entries={} files={} dirs={} snapshotBytes={} memoryBytes={} readyMs={}",
                volume.normalizedRootPath,
                static_cast<uint32_t>(stats.fileSystemKind),
                stats.journalAvailable,
                stats.journalReplayApplied,
                stats.rebuiltJournalIdMismatch,
                stats.rebuiltJournalRangeInvalid,
                stats.rebuiltSnapshotCorruption,
                stats.usedNtfsEnumeration,
                stats.usedTraversalSeed,
                stats.entryCount,
                stats.fileCount,
                stats.directoryCount,
                stats.snapshotFileBytes,
                stats.estimatedMemoryBytes,
                stats.ensureReadyDurationMs);
    return S_OK;
}

HRESULT ExecuteQuery(const VolumeIndex& volume,
                     const QueryPlan& plan,
                     std::vector<Candidate>& outCandidates,
                     QueryStats& stats) noexcept
{
    outCandidates.clear();

    const auto rootIt = volume.entries.find(volume.trackedRootId);
    if (rootIt == volume.entries.end())
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    const auto emitCandidate = [&](const Entry& entry) noexcept
    {
        const bool isDirectory = IsDirectoryAttributes(entry.fileAttributes);
        if ((isDirectory && ! plan.includeDirectories) || (! isDirectory && ! plan.includeFiles))
        {
            return false;
        }

        if (! MatchName(plan, entry.name))
        {
            return false;
        }

        Candidate candidate{};
        candidate.fullPath       = entry.fullPath;
        candidate.displayName    = entry.name.empty() ? GetPathLeaf(entry.fullPath) : entry.name;
        candidate.fileAttributes = entry.fileAttributes;
        outCandidates.push_back(std::move(candidate));
        return true;
    };

    if (! volume.trackedRootIsDirectory)
    {
        if (emitCandidate(rootIt->second))
        {
            stats.candidateCount = 1u;
        }
        return S_OK;
    }

    std::vector<NodeId> stack(rootIt->second.children.begin(), rootIt->second.children.end());
    while (! stack.empty())
    {
        const NodeId id = stack.back();
        stack.pop_back();

        const auto it = volume.entries.find(id);
        if (it == volume.entries.end())
        {
            continue;
        }

        const Entry& entry = it->second;
        emitCandidate(entry);
        if (plan.maxResults != 0u && outCandidates.size() >= plan.maxResults)
        {
            break;
        }

        if (plan.recursive && IsDirectoryAttributes(entry.fileAttributes))
        {
            for (auto childIt = entry.children.rbegin(); childIt != entry.children.rend(); ++childIt)
            {
                stack.push_back(*childIt);
            }
        }
    }

    stats.candidateCount = static_cast<uint64_t>(outCandidates.size());
    return S_OK;
}
} // namespace

Repository::Repository(RepositoryOptions options) noexcept : _options(std::move(options))
{
}

HRESULT Repository::ProbePath(std::wstring_view rootPath, SupportInfo& outSupport) noexcept
{
    try
    {
        return PopulateSupportInfo(rootPath, outSupport);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: ProbePath failed with an unexpected std::exception.");
        outSupport = {};
        return E_FAIL;
    }
}

HRESULT Repository::Query(const QueryPlan& plan,
                          CancelCheckFn cancelCheck,
                          void* cancelCookie,
                          std::vector<Candidate>& outCandidates,
                          QueryStats* outStats) noexcept
{
    try
    {
        outCandidates.clear();
        QueryStats stats{};

        if (plan.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(plan.rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        const std::wstring rootKey = FoldPathKey(support.normalizedRootPath);

        std::shared_ptr<VolumeIndex> volume;
        {
            // Repository mutex protects only the root-key -> volume map. Per-volume state is serialized separately.
            std::lock_guard guard(_mutex);

            auto it = _volumes.find(rootKey);
            if (it == _volumes.end())
            {
                auto newVolume = std::make_shared<VolumeIndex>();
                newVolume->normalizedRootPath = support.normalizedRootPath;
                newVolume->rootKey            = rootKey;
                newVolume->fileSystemKind     = support.fileSystemKind;
                newVolume->volumeRoot         = ExtractVolumeRoot(support.normalizedRootPath);
                newVolume->volumeDevicePath   = BuildVolumeDevicePath(newVolume->volumeRoot);
                newVolume->snapshotPath       = BuildSnapshotPath(newVolume->normalizedRootPath, newVolume->fileSystemKind, _options);
                if (newVolume->volumeRoot.empty() || newVolume->volumeDevicePath.empty() || newVolume->snapshotPath.empty())
                {
                    return kNotSupportedHr;
                }

                it = _volumes.emplace(rootKey, std::move(newVolume)).first;
            }

            volume = it->second;
        }

        std::lock_guard volumeGuard(volume->mutex);

        hr = EnsureReady(*volume, cancelCheck, cancelCookie, stats);
        if (FAILED(hr))
        {
            return hr;
        }

        QueryPlan effectivePlan = plan;
        effectivePlan.rootPath  = volume->normalizedRootPath;

        const auto executeStart = std::chrono::steady_clock::now();
        hr = ExecuteQuery(*volume, effectivePlan, outCandidates, stats);
        if (FAILED(hr))
        {
            return hr;
        }
        stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - executeStart);

        Debug::Info(L"LocalSearchIndexCore: query root='{}' pattern='{}' mode={} recursive={} includeFiles={} includeDirs={} maxResults={} "
                    L"candidates={} entryCount={} snapshotBytes={} memoryBytes={} readyMs={} queryMs={}",
                    effectivePlan.rootPath,
                    effectivePlan.namePattern,
                    static_cast<uint32_t>(effectivePlan.nameMode),
                    effectivePlan.recursive,
                    effectivePlan.includeFiles,
                    effectivePlan.includeDirectories,
                    effectivePlan.maxResults,
                    stats.candidateCount,
                    stats.entryCount,
                    stats.snapshotFileBytes,
                    stats.estimatedMemoryBytes,
                    stats.ensureReadyDurationMs,
                    stats.executeQueryDurationMs);

        if (outStats != nullptr)
        {
            *outStats = stats;
        }

        return S_OK;
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
        Debug::Error(L"LocalSearchIndexCore: Query failed with an unexpected std::exception.");
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT Repository::InvalidateRoot(std::wstring_view rootPath, bool deleteSnapshot) noexcept
{
    try
    {
        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        const std::wstring rootKey      = FoldPathKey(support.normalizedRootPath);
        const std::wstring snapshotPath = BuildSnapshotPath(support.normalizedRootPath, support.fileSystemKind, _options);

        {
            std::lock_guard guard(_mutex);
            _volumes.erase(rootKey);
        }

        if (deleteSnapshot && ! snapshotPath.empty())
        {
            std::error_code ec;
            static_cast<void>(std::filesystem::remove(std::filesystem::path(snapshotPath), ec));
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: InvalidateRoot failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

#ifdef _DEBUG
HRESULT Repository::DropCachedVolumeForTests(std::wstring_view rootPath) noexcept
{
    return InvalidateRoot(rootPath, false);
}

HRESULT Repository::CorruptSnapshotForTests(std::wstring_view rootPath, SnapshotCorruptionMode mode) noexcept
{
    try
    {
        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        const std::wstring snapshotPath = BuildSnapshotPath(support.normalizedRootPath, support.fileSystemKind, _options);
        if (snapshotPath.empty())
        {
            return E_INVALIDARG;
        }

        wil::unique_handle file(::CreateFileW(snapshotPath.c_str(),
                                              GENERIC_READ | GENERIC_WRITE,
                                              FILE_SHARE_READ,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_ATTRIBUTE_NORMAL,
                                              nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        SnapshotHeader header{};
        DWORD bytesRead = 0u;
        if (::ReadFile(file.get(), &header, sizeof(header), &bytesRead, nullptr) == 0 || bytesRead != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        switch (mode)
        {
            case SnapshotCorruptionMode::InvalidMagic: header.magic ^= 0x13579BDFu; break;
            case SnapshotCorruptionMode::JournalIdMismatch: header.journalId = header.journalId + 1u; break;
            case SnapshotCorruptionMode::NextUsnPastEnd: header.nextUsn = (std::numeric_limits<uint64_t>::max)(); break;
        }

        LARGE_INTEGER zero{};
        if (::SetFilePointerEx(file.get(), zero, nullptr, FILE_BEGIN) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        DWORD written = 0u;
        if (::WriteFile(file.get(), &header, sizeof(header), &written, nullptr) == 0 || written != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: CorruptSnapshotForTests failed with an unexpected std::exception.");
        return E_FAIL;
    }
}
#endif
} // namespace LocalSearchIndexCore

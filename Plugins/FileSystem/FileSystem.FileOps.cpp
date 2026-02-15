#include "FileSystem.Internal.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <condition_variable>
#include <cstring>
#include <deque>
#include <filesystem>
#include <limits>
#include <mutex>
#include <ranges>
#include <thread>

#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <winioctl.h>

using namespace FileSystemInternal;

namespace
{
#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
struct ParallelOperationState
{
    ParallelOperationState() noexcept                                = default;
    ParallelOperationState(const ParallelOperationState&)            = delete;
    ParallelOperationState& operator=(const ParallelOperationState&) = delete;
    ParallelOperationState(ParallelOperationState&&)                 = delete;
    ParallelOperationState& operator=(ParallelOperationState&&)      = delete;

    std::atomic<unsigned __int64> completedBytes{0};
    std::atomic<unsigned long> completedItems{0};
    std::atomic<unsigned __int64> bandwidthLimitBytesPerSecond{0};

    ULONGLONG startTick = 0;
    std::mutex callbackMutex;
    ULONGLONG lastProgressReportTick = 0;
    std::atomic<ULONGLONG> lastCancelCheckTick{0};

    std::atomic<bool> cancelRequested{false};
    std::atomic<bool> stopOnErrorRequested{false};
    std::atomic<HRESULT> firstError{S_OK};
    std::atomic<bool> hadFailure{false};
};

struct OperationContext
{
    FileSystemOperation type      = FILESYSTEM_COPY;
    IFileSystemCallback* callback = nullptr;
    void* callbackCookie          = nullptr;
    unsigned __int64 progressStreamId = 0;
    FileSystemOptions optionsState{};
    FileSystemOptions* options      = nullptr;
    unsigned long totalItems        = 0;
    unsigned long completedItems    = 0;
    unsigned __int64 totalBytes     = 0;
    unsigned __int64 completedBytes = 0;
    bool continueOnError            = false;
    bool allowOverwrite             = false;
    bool allowReplaceReadonly       = false;
    bool recursive                  = false;
    bool useRecycleBin              = false;
    FileSystemArenaOwner itemArena;
    FileSystemArenaOwner progressArena;
    const wchar_t* itemSource          = nullptr;
    const wchar_t* itemDestination     = nullptr;
    const wchar_t* progressSource      = nullptr;
    const wchar_t* progressDestination = nullptr;

    ParallelOperationState* parallel = nullptr;

    ULONGLONG lastProgressReportTick = 0;

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    std::wstring reparseRootSourcePath;
    std::wstring reparseRootDestinationPath;
};
#pragma warning(pop)

struct CopyProgressContext
{
    OperationContext* context                 = nullptr;
    unsigned __int64 itemBaseBytes            = 0; // Used only for sequential operations.
    unsigned __int64 lastItemBytesTransferred = 0; // Used only for parallel operations.
    ULONGLONG startTick                       = 0; // Used only for sequential operations.
};

bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept
{
    return (static_cast<unsigned long>(flags) & static_cast<unsigned long>(flag)) != 0u;
}

bool IsCancellationHr(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

HRESULT NormalizeCancellation(HRESULT hr) noexcept
{
    if (IsCancellationHr(hr))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

[[nodiscard]] bool IsReparsePoint(DWORD attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

[[nodiscard]] bool IsDirectory(DWORD attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

struct ReparsePointHeader
{
    DWORD tag        = 0;
    USHORT dataBytes = 0;
    USHORT reserved  = 0;
};
static_assert(sizeof(ReparsePointHeader) == 8);

struct ReparsePointData
{
    DWORD tag       = 0;
    DWORD sizeBytes = 0;
    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
};

struct MountPointReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
};
static_assert(sizeof(MountPointReparseHeader) == 8);

struct SymbolicLinkReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
    ULONG flags             = 0;
};
static_assert(sizeof(SymbolicLinkReparseHeader) == 12);

constexpr ULONG kSymlinkRelativeFlag = 0x00000001u;

struct ParsedDirectoryReparsePoint
{
    DWORD tag       = 0;
    bool isRelative = false;
    std::wstring substitutePath;
    std::wstring printPath;
};

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

void NormalizeSlashes(std::wstring& path) noexcept
{
    std::ranges::replace(path, L'/', L'\\');
}

[[nodiscard]] size_t GetRootLength(std::wstring_view path) noexcept
{
    if (path.size() >= 2 && path[1] == L':')
    {
        if (path.size() >= 3 && IsPathSeparator(path[2]))
        {
            return 3;
        }
        return 2;
    }

    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        size_t firstSep = path.find(L'\\', 2);
        if (firstSep == std::wstring_view::npos)
        {
            firstSep = path.find(L'/', 2);
        }
        if (firstSep == std::wstring_view::npos)
        {
            return path.size();
        }

        size_t secondSep = path.find_first_of(L"\\/", firstSep + 1);
        if (secondSep == std::wstring_view::npos)
        {
            return path.size();
        }
        return secondSep + 1;
    }

    if (! path.empty() && IsPathSeparator(path.front()))
    {
        return 1;
    }

    return 0;
}

[[nodiscard]] std::wstring TrimTrailingSeparatorsPreserveRoot(std::wstring path) noexcept
{
    NormalizeSlashes(path);
    const size_t rootLength = GetRootLength(path);
    while (path.size() > rootLength && ! path.empty() && IsPathSeparator(path.back()))
    {
        path.pop_back();
    }
    return path;
}

[[nodiscard]] bool EqualsInsensitive(std::wstring_view left, std::wstring_view right) noexcept
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool IsPathWithinRoot(std::wstring_view path, std::wstring_view root) noexcept
{
    if (root.empty() || path.size() < root.size())
    {
        return false;
    }

    if (! EqualsInsensitive(path.substr(0, root.size()), root))
    {
        return false;
    }

    if (path.size() == root.size())
    {
        return true;
    }

    return IsPathSeparator(path[root.size()]);
}

[[nodiscard]] std::wstring StripWin32ExtendedPrefix(std::wstring_view path)
{
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::wstring NtPathToWin32Path(std::wstring_view path)
{
    if (path.rfind(L"\\??\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::wstring Win32PathToNtPath(std::wstring_view path)
{
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path);
    }

    if (path.rfind(L"\\\\", 0) == 0)
    {
        return std::wstring(L"\\??\\UNC\\") + std::wstring(path.substr(2));
    }

    return std::wstring(L"\\??\\") + std::wstring(path);
}

[[nodiscard]] bool ParseDirectoryReparsePoint(const ReparsePointData& data, ParsedDirectoryReparsePoint& out) noexcept
{
    out = {};

    if (data.sizeBytes < sizeof(ReparsePointHeader))
    {
        return false;
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(data.buffer.data());
    if (static_cast<size_t>(header->dataBytes) + sizeof(ReparsePointHeader) > static_cast<size_t>(data.sizeBytes))
    {
        return false;
    }

    out.tag = header->tag;

    const std::byte* payloadBase = data.buffer.data() + sizeof(ReparsePointHeader);
    const size_t payloadBytes    = header->dataBytes;

    auto readPathSlice = [&](USHORT offsetBytes, USHORT lengthBytes, size_t fixedHeaderBytes, std::wstring& target) noexcept -> bool
    {
        if ((offsetBytes % sizeof(wchar_t)) != 0u || (lengthBytes % sizeof(wchar_t)) != 0u)
        {
            return false;
        }
        if (payloadBytes < fixedHeaderBytes)
        {
            return false;
        }
        const size_t pathBufferBytes = payloadBytes - fixedHeaderBytes;
        if (offsetBytes > pathBufferBytes || lengthBytes > pathBufferBytes ||
            (static_cast<size_t>(offsetBytes) + static_cast<size_t>(lengthBytes)) > pathBufferBytes)
        {
            return false;
        }

        const auto* text = reinterpret_cast<const wchar_t*>(payloadBase + fixedHeaderBytes + offsetBytes);
        target.assign(text, text + (lengthBytes / sizeof(wchar_t)));
        return true;
    };

    if (out.tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        if (payloadBytes < sizeof(MountPointReparseHeader))
        {
            return false;
        }

        const auto* mount = reinterpret_cast<const MountPointReparseHeader*>(payloadBase);
        if (! readPathSlice(mount->substituteOffset, mount->substituteLength, sizeof(MountPointReparseHeader), out.substitutePath))
        {
            return false;
        }
        if (! readPathSlice(mount->printOffset, mount->printLength, sizeof(MountPointReparseHeader), out.printPath))
        {
            return false;
        }
        out.isRelative = false;
        return true;
    }

    if (out.tag == IO_REPARSE_TAG_SYMLINK)
    {
        if (payloadBytes < sizeof(SymbolicLinkReparseHeader))
        {
            return false;
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseHeader*>(payloadBase);
        if (! readPathSlice(symlink->substituteOffset, symlink->substituteLength, sizeof(SymbolicLinkReparseHeader), out.substitutePath))
        {
            return false;
        }
        if (! readPathSlice(symlink->printOffset, symlink->printLength, sizeof(SymbolicLinkReparseHeader), out.printPath))
        {
            return false;
        }
        out.isRelative = (symlink->flags & kSymlinkRelativeFlag) != 0u;
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring ResolveReparseTargetAbsolute(const PathInfo& source, const ParsedDirectoryReparsePoint& parsed) noexcept
{
    std::wstring rawTarget = parsed.substitutePath.empty() ? parsed.printPath : parsed.substitutePath;
    if (rawTarget.empty())
    {
        return {};
    }

    rawTarget = NtPathToWin32Path(rawTarget);
    NormalizeSlashes(rawTarget);

    if (parsed.isRelative)
    {
        std::filesystem::path parent   = std::filesystem::path(source.display).parent_path();
        std::filesystem::path combined = parent / std::filesystem::path(rawTarget);
        std::wstring absolute          = MakeAbsolutePath(combined.lexically_normal().wstring());
        absolute                       = StripWin32ExtendedPrefix(absolute);
        return TrimTrailingSeparatorsPreserveRoot(absolute);
    }

    std::wstring absolute = MakeAbsolutePath(rawTarget);
    absolute              = StripWin32ExtendedPrefix(absolute);
    return TrimTrailingSeparatorsPreserveRoot(absolute);
}

[[nodiscard]] bool TryRetargetPathIntoDestination(std::wstring_view absoluteTargetPath,
                                                  std::wstring_view sourceRootPath,
                                                  std::wstring_view destinationRootPath,
                                                  std::wstring& mappedOut) noexcept
{
    std::wstring normalizedTarget = TrimTrailingSeparatorsPreserveRoot(std::wstring(absoluteTargetPath));
    std::wstring normalizedSource = TrimTrailingSeparatorsPreserveRoot(std::wstring(sourceRootPath));
    std::wstring normalizedDest   = TrimTrailingSeparatorsPreserveRoot(std::wstring(destinationRootPath));

    if (normalizedTarget.empty() || normalizedSource.empty() || normalizedDest.empty())
    {
        return false;
    }

    if (! IsPathWithinRoot(normalizedTarget, normalizedSource))
    {
        return false;
    }

    std::wstring suffix;
    if (normalizedTarget.size() > normalizedSource.size())
    {
        suffix = normalizedTarget.substr(normalizedSource.size());
        while (! suffix.empty() && IsPathSeparator(suffix.front()))
        {
            suffix.erase(suffix.begin());
        }
    }

    mappedOut = normalizedDest;
    if (! suffix.empty())
    {
        if (! mappedOut.empty() && ! IsPathSeparator(mappedOut.back()))
        {
            mappedOut.push_back(L'\\');
        }
        mappedOut.append(suffix);
    }

    mappedOut = TrimTrailingSeparatorsPreserveRoot(mappedOut);
    return true;
}

[[nodiscard]] bool EndsWithSeparator(std::wstring_view path) noexcept
{
    return ! path.empty() && IsPathSeparator(path.back());
}

HRESULT BuildMountPointReparseData(std::wstring targetPath, ReparsePointData& out) noexcept
{
    NormalizeSlashes(targetPath);
    if (! EndsWithSeparator(targetPath))
    {
        targetPath.push_back(L'\\');
    }

    std::wstring substitute = Win32PathToNtPath(targetPath);

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = targetPath.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    const size_t payloadBytes    = sizeof(MountPointReparseHeader) + pathBufferBytes;
    const size_t totalBytes      = sizeof(ReparsePointHeader) + payloadBytes;

    if (payloadBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()) || totalBytes > out.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    out           = {};
    out.tag       = IO_REPARSE_TAG_MOUNT_POINT;
    out.sizeBytes = static_cast<DWORD>(totalBytes);

    auto* header      = reinterpret_cast<ReparsePointHeader*>(out.buffer.data());
    header->tag       = IO_REPARSE_TAG_MOUNT_POINT;
    header->dataBytes = static_cast<USHORT>(payloadBytes);
    header->reserved  = 0;

    auto* mountHeader             = reinterpret_cast<MountPointReparseHeader*>(out.buffer.data() + sizeof(ReparsePointHeader));
    mountHeader->substituteOffset = 0;
    mountHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    mountHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    mountHeader->printLength      = static_cast<USHORT>(printBytes);

    std::byte* pathBuffer = out.buffer.data() + sizeof(ReparsePointHeader) + sizeof(MountPointReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), targetPath.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));
    return S_OK;
}

HRESULT BuildSymlinkReparseData(std::wstring targetPath, bool relative, ReparsePointData& out) noexcept
{
    NormalizeSlashes(targetPath);
    std::wstring substitute = targetPath;
    std::wstring print      = targetPath;

    if (! relative)
    {
        substitute = Win32PathToNtPath(substitute);
    }

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = print.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    const size_t payloadBytes    = sizeof(SymbolicLinkReparseHeader) + pathBufferBytes;
    const size_t totalBytes      = sizeof(ReparsePointHeader) + payloadBytes;

    if (payloadBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()) || totalBytes > out.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    out           = {};
    out.tag       = IO_REPARSE_TAG_SYMLINK;
    out.sizeBytes = static_cast<DWORD>(totalBytes);

    auto* header      = reinterpret_cast<ReparsePointHeader*>(out.buffer.data());
    header->tag       = IO_REPARSE_TAG_SYMLINK;
    header->dataBytes = static_cast<USHORT>(payloadBytes);
    header->reserved  = 0;

    auto* symlinkHeader             = reinterpret_cast<SymbolicLinkReparseHeader*>(out.buffer.data() + sizeof(ReparsePointHeader));
    symlinkHeader->substituteOffset = 0;
    symlinkHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    symlinkHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    symlinkHeader->printLength      = static_cast<USHORT>(printBytes);
    symlinkHeader->flags            = relative ? kSymlinkRelativeFlag : 0u;

    std::byte* pathBuffer = out.buffer.data() + sizeof(ReparsePointHeader) + sizeof(SymbolicLinkReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), print.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));
    return S_OK;
}

HRESULT ReadReparsePointData(const std::wstring& path, ReparsePointData& out) noexcept
{
    out = {};

    // Protected junctions (e.g. localized/system junctions) may deny "read data / list directory" access
    // but still allow querying reparse metadata. Keep access minimal so we can copy the link itself.
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, out.buffer.data(), static_cast<DWORD>(out.buffer.size()), &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(out.buffer.data());
    out.tag            = header->tag;
    out.sizeBytes      = bytesReturned;
    return S_OK;
}

HRESULT WriteReparsePointData(const std::wstring& path, const ReparsePointData& data) noexcept
{
    if (data.sizeBytes < sizeof(ReparsePointHeader) || data.sizeBytes > data.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    // Minimal access for setting reparse data on the destination link.
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(
            handle.get(), FSCTL_SET_REPARSE_POINT, const_cast<std::byte*>(data.buffer.data()), data.sizeBytes, nullptr, 0, &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

void AddCompletedBytes(OperationContext& context, unsigned __int64 bytes) noexcept
{
    if (bytes == 0)
    {
        return;
    }

    if (context.parallel)
    {
        context.parallel->completedBytes.fetch_add(bytes, std::memory_order_acq_rel);
        return;
    }

    if (std::numeric_limits<unsigned __int64>::max() - context.completedBytes < bytes)
    {
        context.completedBytes = std::numeric_limits<unsigned __int64>::max();
        return;
    }

    context.completedBytes += bytes;
}

void AddCompletedItems(OperationContext& context, unsigned long items) noexcept
{
    if (items == 0)
    {
        return;
    }

    if (context.parallel)
    {
        context.parallel->completedItems.fetch_add(items, std::memory_order_acq_rel);
        return;
    }

    constexpr unsigned __int64 maxUlong = static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max());
    const unsigned __int64 current      = static_cast<unsigned __int64>(context.completedItems);
    const unsigned __int64 desired      = current + static_cast<unsigned __int64>(items);
    context.completedItems              = static_cast<unsigned long>(std::min(desired, maxUlong));
}

unsigned __int64 GetBandwidthLimit(const FileSystemOptions* options) noexcept
{
    if (! options)
    {
        return 0;
    }
    return options->bandwidthLimitBytesPerSecond;
}

HRESULT CalculateStringBytes(const wchar_t* text, unsigned long* outBytes) noexcept
{
    if (! outBytes)
    {
        return E_POINTER;
    }

    if (! text)
    {
        *outBytes = 0;
        return S_OK;
    }

    const size_t length = ::wcslen(text);
    if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *outBytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
    return S_OK;
}

HRESULT BuildArenaForPaths(
    FileSystemArenaOwner& arenaOwner, const wchar_t* source, const wchar_t* destination, const wchar_t** outSource, const wchar_t** outDestination) noexcept
{
    if (! outSource || ! outDestination)
    {
        return E_POINTER;
    }

    *outSource      = nullptr;
    *outDestination = nullptr;

    unsigned long sourceBytes = 0;
    HRESULT hr                = CalculateStringBytes(source, &sourceBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long destinationBytes = 0;
    hr                             = CalculateStringBytes(destination, &destinationBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long totalBytes = sourceBytes;
    if (destinationBytes > 0)
    {
        if (totalBytes > std::numeric_limits<unsigned long>::max() - destinationBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += destinationBytes;
    }

    FileSystemArena* arena = arenaOwner.Get();
    if (! arena || arena->buffer == nullptr || arena->capacityBytes < totalBytes)
    {
        hr = arenaOwner.Initialize(totalBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        arena = arenaOwner.Get();
    }

    if (arena && arena->buffer)
    {
        arena->usedBytes = 0;
    }

    if (sourceBytes > 0)
    {
        auto* sourceBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, sourceBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourceBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t sourceLength = (sourceBytes / sizeof(wchar_t)) - 1u;
        if (sourceLength > 0)
        {
            ::CopyMemory(sourceBuffer, source, sourceLength * sizeof(wchar_t));
        }
        sourceBuffer[sourceLength] = L'\0';
        *outSource                 = sourceBuffer;
    }

    if (destinationBytes > 0)
    {
        auto* destinationBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, destinationBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! destinationBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t destinationLength = (destinationBytes / sizeof(wchar_t)) - 1u;
        if (destinationLength > 0)
        {
            ::CopyMemory(destinationBuffer, destination, destinationLength * sizeof(wchar_t));
        }
        destinationBuffer[destinationLength] = L'\0';
        *outDestination                      = destinationBuffer;
    }

    return S_OK;
}

HRESULT SetItemPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.itemArena, source, destination, &context.itemSource, &context.itemDestination);
}

HRESULT SetProgressPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.progressArena, source, destination, &context.progressSource, &context.progressDestination);
}

HRESULT CheckCancelLocked(OperationContext& context) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    BOOL cancel = FALSE;
    HRESULT hr  = context.callback->FileSystemShouldCancel(&cancel, context.callbackCookie);
    hr          = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (cancel)
    {
        if (context.parallel)
        {
            context.parallel->cancelRequested.store(true, std::memory_order_release);
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT CheckCancel(OperationContext& context) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        constexpr ULONGLONG kMinCancelCheckMs = 50ull;
        const ULONGLONG nowTick               = GetTickCount64();
        const ULONGLONG lastTick              = context.parallel->lastCancelCheckTick.load(std::memory_order_acquire);
        if (lastTick != 0 && nowTick >= lastTick && (nowTick - lastTick) < kMinCancelCheckMs)
        {
            return S_OK;
        }

        std::scoped_lock lock(context.parallel->callbackMutex);
        const HRESULT hr = CheckCancelLocked(context);
        context.parallel->lastCancelCheckTick.store(nowTick, std::memory_order_release);
        return hr;
    }

    return CheckCancelLocked(context);
}

HRESULT ReportProgress(OperationContext& context, unsigned __int64 currentItemTotalBytes, unsigned __int64 currentItemCompletedBytes) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    constexpr ULONGLONG kMinProgressMsCopyMove = 50ull;
    constexpr ULONGLONG kMinProgressMsDelete   = 100ull;
    const ULONGLONG minProgressMs              = (context.type == FILESYSTEM_DELETE) ? kMinProgressMsDelete : kMinProgressMsCopyMove;

    const unsigned long completedItems    = context.parallel ? context.parallel->completedItems.load(std::memory_order_acquire) : context.completedItems;
    const unsigned __int64 completedBytes = context.parallel ? context.parallel->completedBytes.load(std::memory_order_acquire) : context.completedBytes;

    const bool isFinalItem    = currentItemTotalBytes > 0 && currentItemCompletedBytes >= currentItemTotalBytes;
    const bool isFinalOverall = context.totalItems > 0 && completedItems >= context.totalItems;
    const bool isFinal        = isFinalItem || isFinalOverall;

    const ULONGLONG nowTick = GetTickCount64();
    if (! isFinal && context.lastProgressReportTick != 0 && nowTick >= context.lastProgressReportTick &&
        (nowTick - context.lastProgressReportTick) < minProgressMs)
    {
        return S_OK;
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        if (context.type == FILESYSTEM_DELETE && ! isFinal && context.parallel->lastProgressReportTick != 0 &&
            nowTick >= context.parallel->lastProgressReportTick && (nowTick - context.parallel->lastProgressReportTick) < minProgressMs)
        {
            return S_OK;
        }

        HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                          context.totalItems,
                                                          completedItems,
                                                          context.totalBytes,
                                                          completedBytes,
                                                          context.progressSource,
                                                          context.progressDestination,
                                                          currentItemTotalBytes,
                                                          currentItemCompletedBytes,
                                                          context.options,
                                                          context.progressStreamId,
                                                          context.callbackCookie);
        hr         = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        context.lastProgressReportTick           = nowTick;
        context.parallel->lastProgressReportTick = nowTick;

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                      context.totalItems,
                                                      completedItems,
                                                      context.totalBytes,
                                                      completedBytes,
                                                      context.progressSource,
                                                      context.progressDestination,
                                                      currentItemTotalBytes,
                                                      currentItemCompletedBytes,
                                                      context.options,
                                                      context.progressStreamId,
                                                      context.callbackCookie);
    hr         = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.lastProgressReportTick = nowTick;

    return CheckCancel(context);
}

HRESULT ReportProgressForced(OperationContext& context, unsigned __int64 currentItemTotalBytes, unsigned __int64 currentItemCompletedBytes) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    const unsigned long completedItems    = context.parallel ? context.parallel->completedItems.load(std::memory_order_acquire) : context.completedItems;
    const unsigned __int64 completedBytes = context.parallel ? context.parallel->completedBytes.load(std::memory_order_acquire) : context.completedBytes;

    const ULONGLONG nowTick = GetTickCount64();
    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                          context.totalItems,
                                                          completedItems,
                                                          context.totalBytes,
                                                          completedBytes,
                                                          context.progressSource,
                                                          context.progressDestination,
                                                          currentItemTotalBytes,
                                                          currentItemCompletedBytes,
                                                          context.options,
                                                          context.progressStreamId,
                                                          context.callbackCookie);
        hr         = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        context.lastProgressReportTick           = nowTick;
        context.parallel->lastProgressReportTick = nowTick;
        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                      context.totalItems,
                                                      completedItems,
                                                      context.totalBytes,
                                                      completedBytes,
                                                      context.progressSource,
                                                      context.progressDestination,
                                                      currentItemTotalBytes,
                                                      currentItemCompletedBytes,
                                                      context.options,
                                                      context.progressStreamId,
                                                      context.callbackCookie);
    hr         = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.lastProgressReportTick = nowTick;
    return CheckCancel(context);
}

HRESULT ReportItemCompleted(OperationContext& context, unsigned long itemIndex, HRESULT status) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemItemCompleted(
            context.type, itemIndex, context.itemSource, context.itemDestination, status, context.options, context.callbackCookie);
        hr = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemItemCompleted(
        context.type, itemIndex, context.itemSource, context.itemDestination, status, context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

HRESULT ReportIssue(OperationContext& context, HRESULT status, FileSystemIssueAction* action) noexcept
{
    if (! action)
    {
        return E_POINTER;
    }

    *action = FileSystemIssueAction::Cancel;

    if (! context.callback)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemIssue(
            context.type, context.progressSource, context.progressDestination, status, action, context.options, context.callbackCookie);
        hr = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemIssue(
        context.type, context.progressSource, context.progressDestination, status, action, context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

void InitializeOperationContext(OperationContext& context,
                                FileSystemOperation type,
                                FileSystemFlags flags,
                                const FileSystemOptions* options,
                                IFileSystemCallback* callback,
                                void* cookie,
                                unsigned long totalItems,
                                FileSystemReparsePointPolicy reparsePointPolicy) noexcept
{
    context.type           = type;
    context.callback       = callback;
    context.callbackCookie = callback != nullptr ? cookie : nullptr;
    context.progressStreamId = 0;
    context.optionsState   = {};
    if (options)
    {
        context.optionsState = *options;
    }
    context.options              = &context.optionsState;
    context.totalItems           = totalItems;
    context.completedItems       = 0;
    context.totalBytes           = 0;
    context.completedBytes       = 0;
    context.continueOnError      = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    context.allowOverwrite       = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    context.allowReplaceReadonly = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    context.recursive            = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    context.useRecycleBin        = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    context.itemSource           = nullptr;
    context.itemDestination      = nullptr;
    context.progressSource       = nullptr;
    context.progressDestination  = nullptr;
    context.reparsePointPolicy   = reparsePointPolicy;
    context.reparseRootSourcePath.clear();
    context.reparseRootDestinationPath.clear();
}

HRESULT GetFileSizeBytes(const std::wstring& path, unsigned __int64* sizeBytes) noexcept
{
    if (! sizeBytes)
    {
        return E_POINTER;
    }

    *sizeBytes = 0;

    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (! GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return S_OK;
    }

    const unsigned __int64 high = static_cast<unsigned __int64>(data.nFileSizeHigh);
    const unsigned __int64 low  = static_cast<unsigned __int64>(data.nFileSizeLow);
    *sizeBytes                  = (high << 32) | low;
    return S_OK;
}

HRESULT RemoveDirectoryRecursiveNoFollow(OperationContext& context, const std::wstring& directoryExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring searchPattern = AppendPath(directoryExtended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error);
    }

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        const std::wstring child = AppendPath(directoryExtended, data.cFileName);
        const DWORD attributes   = data.dwFileAttributes;

        if (IsDirectory(attributes))
        {
            if (IsReparsePoint(attributes))
            {
                if (! RemoveDirectoryW(child.c_str()))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
            }
            else
            {
                hr = RemoveDirectoryRecursiveNoFollow(context, child);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        else
        {
            if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                if (! context.allowReplaceReadonly)
                {
                    return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                }

                const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
                if (! SetFileAttributesW(child.c_str(), newAttributes))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
            }

            if (! DeleteFileW(child.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return HRESULT_FROM_WIN32(error);
    }

    DWORD dirAttributes = GetFileAttributesW(directoryExtended.c_str());
    if (dirAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if ((dirAttributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        dirAttributes &= ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(directoryExtended.c_str(), dirAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    if (! RemoveDirectoryW(directoryExtended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

HRESULT RemovePathForOverwrite(OperationContext& context, const std::wstring& pathExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    DWORD attributes = GetFileAttributesW(pathExtended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (IsDirectory(attributes))
    {
        if (IsReparsePoint(attributes))
        {
            if (! RemoveDirectoryW(pathExtended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            return S_OK;
        }

        return RemoveDirectoryRecursiveNoFollow(context, pathExtended);
    }

    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(pathExtended.c_str(), newAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    if (! DeleteFileW(pathExtended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

DWORD CALLBACK CopyProgressRoutine(LARGE_INTEGER totalFileSize,
                                   LARGE_INTEGER totalBytesTransferred,
                                   [[maybe_unused]] LARGE_INTEGER streamSize,
                                   [[maybe_unused]] LARGE_INTEGER streamBytesTransferred,
                                   [[maybe_unused]] DWORD streamNumber,
                                   [[maybe_unused]] DWORD callbackReason,
                                   [[maybe_unused]] HANDLE sourceFile,
                                   [[maybe_unused]] HANDLE destinationFile,
                                   LPVOID context) noexcept
{
    auto* progressContext = static_cast<CopyProgressContext*>(context);
    if (! progressContext || ! progressContext->context)
    {
        return PROGRESS_CONTINUE;
    }

    OperationContext& opContext          = *progressContext->context;
    const unsigned __int64 itemTotal     = static_cast<unsigned __int64>(totalFileSize.QuadPart);
    const unsigned __int64 itemCompleted = static_cast<unsigned __int64>(totalBytesTransferred.QuadPart);

    if (opContext.parallel)
    {
        const bool cancelRequested = opContext.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = opContext.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return PROGRESS_CANCEL;
        }

        if (itemCompleted >= progressContext->lastItemBytesTransferred)
        {
            const unsigned __int64 delta = itemCompleted - progressContext->lastItemBytesTransferred;
            if (delta > 0)
            {
                opContext.parallel->completedBytes.fetch_add(delta, std::memory_order_acq_rel);
            }
            progressContext->lastItemBytesTransferred = itemCompleted;
        }
        else
        {
            // Defensive: restart delta tracking if the API reports a smaller value.
            progressContext->lastItemBytesTransferred = itemCompleted;
        }

        HRESULT hr = ReportProgress(opContext, itemTotal, itemCompleted);
        if (FAILED(hr))
        {
            return PROGRESS_CANCEL;
        }

        const unsigned __int64 bandwidthLimit = opContext.parallel->bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
        if (bandwidthLimit > 0)
        {
            const ULONGLONG now              = GetTickCount64();
            const unsigned __int64 elapsedMs = static_cast<unsigned __int64>(now - opContext.parallel->startTick);

            const unsigned __int64 bytesSoFar       = opContext.parallel->completedBytes.load(std::memory_order_acquire);
            constexpr unsigned __int64 maxSafeBytes = std::numeric_limits<unsigned __int64>::max() / 1000u;

            unsigned __int64 desiredMs = 0;
            if (bytesSoFar > 0 && bytesSoFar <= maxSafeBytes)
            {
                desiredMs = (bytesSoFar * 1000u) / bandwidthLimit;
            }
            else if (bytesSoFar > maxSafeBytes)
            {
                desiredMs = std::numeric_limits<unsigned __int64>::max();
            }

            if (desiredMs > elapsedMs)
            {
                const unsigned __int64 remaining = desiredMs - elapsedMs;
                const DWORD sleepMs = remaining > std::numeric_limits<DWORD>::max() ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(remaining);
                if (sleepMs > 0)
                {
                    ::Sleep(sleepMs);
                }
            }
        }
    }
    else
    {
        opContext.completedBytes = progressContext->itemBaseBytes + itemCompleted;

        HRESULT hr = ReportProgress(opContext, itemTotal, itemCompleted);
        if (FAILED(hr))
        {
            return PROGRESS_CANCEL;
        }

        const unsigned __int64 bandwidthLimit = GetBandwidthLimit(opContext.options);
        if (bandwidthLimit > 0)
        {
            if (progressContext->startTick == 0)
            {
                progressContext->startTick = GetTickCount64();
            }

            const ULONGLONG now                     = GetTickCount64();
            const unsigned __int64 elapsedMs        = static_cast<unsigned __int64>(now - progressContext->startTick);
            constexpr unsigned __int64 maxSafeBytes = std::numeric_limits<unsigned __int64>::max() / 1000u;

            unsigned __int64 desiredMs = 0;
            if (itemCompleted > 0 && itemCompleted <= maxSafeBytes)
            {
                desiredMs = (itemCompleted * 1000u) / bandwidthLimit;
            }
            else if (itemCompleted > maxSafeBytes)
            {
                desiredMs = std::numeric_limits<unsigned __int64>::max();
            }

            if (desiredMs > elapsedMs)
            {
                const unsigned __int64 remaining = desiredMs - elapsedMs;
                const DWORD sleepMs = remaining > std::numeric_limits<DWORD>::max() ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(remaining);
                if (sleepMs > 0)
                {
                    ::Sleep(sleepMs);
                }
            }
        }
    }

    return PROGRESS_CONTINUE;
}

HRESULT CopyFileInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, unsigned __int64* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure =
        [&](HRESULT failure, unsigned __int64 currentItemTotalBytes = 0, unsigned __int64 currentItemCompletedBytes = 0) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, currentItemTotalBytes, currentItemCompletedBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (! context.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }

        if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! context.allowReplaceReadonly)
            {
                return returnFailure(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
            }

            const DWORD newAttributes = destinationAttributes & ~FILE_ATTRIBUTE_READONLY;
            if (! SetFileAttributesW(destination.extended.c_str(), newAttributes))
            {
                return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
            }
        }
    }

    unsigned __int64 fileBytes = 0;
    hr                         = GetFileSizeBytes(source.extended, &fileBytes);
    if (FAILED(hr))
    {
        return returnFailure(hr);
    }

    CopyProgressContext progress{};
    progress.context = &context;
    if (! context.parallel)
    {
        progress.itemBaseBytes = context.completedBytes;
        progress.startTick     = GetTickCount64();
    }

    const DWORD copyFlags = context.allowOverwrite ? 0u : COPY_FILE_FAIL_IF_EXISTS;
    if (! CopyFileExW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, nullptr, copyFlags))
    {
        DWORD error = GetLastError();
        if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return returnFailure(HRESULT_FROM_WIN32(error), fileBytes, progress.lastItemBytesTransferred);
    }

    *bytesCopied = fileBytes;
    if (context.parallel)
    {
        if (fileBytes > progress.lastItemBytesTransferred)
        {
            context.parallel->completedBytes.fetch_add(fileBytes - progress.lastItemBytesTransferred, std::memory_order_acq_rel);
            progress.lastItemBytesTransferred = fileBytes;
        }
    }
    else
    {
        context.completedBytes = progress.itemBaseBytes + fileBytes;
    }
    return S_OK;
}

HRESULT CopyReparsePointInternal(
    OperationContext& context, const PathInfo& source, const PathInfo& destination, DWORD sourceAttributes, unsigned __int64* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure =
        [&](HRESULT failure, unsigned __int64 currentItemTotalBytes = 0, unsigned __int64 currentItemCompletedBytes = 0) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, currentItemTotalBytes, currentItemCompletedBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    const bool isDirectory = IsDirectory(sourceAttributes);
    if (! isDirectory)
    {
        // Copy file reparse points as links only. Never silently fall back to dereferencing data copy.
        unsigned __int64 fileBytes = 0;
        const HRESULT sizeHr       = GetFileSizeBytes(source.extended, &fileBytes);
        if (FAILED(sizeHr))
        {
            return returnFailure(sizeHr);
        }

        CopyProgressContext progress{};
        progress.context = &context;
        if (! context.parallel)
        {
            progress.itemBaseBytes = context.completedBytes;
            progress.startTick     = GetTickCount64();
        }

        const DWORD overwriteFlag = context.allowOverwrite ? 0u : COPY_FILE_FAIL_IF_EXISTS;
        const DWORD copyFlags     = overwriteFlag | COPY_FILE_COPY_SYMLINK;
        if (! CopyFileExW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, nullptr, copyFlags))
        {
            const DWORD error = GetLastError();
            if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (error == ERROR_INVALID_PARAMETER)
            {
                return returnFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), fileBytes, progress.lastItemBytesTransferred);
            }
            return returnFailure(HRESULT_FROM_WIN32(error), fileBytes, progress.lastItemBytesTransferred);
        }

        *bytesCopied = fileBytes;
        if (context.parallel)
        {
            if (fileBytes > progress.lastItemBytesTransferred)
            {
                context.parallel->completedBytes.fetch_add(fileBytes - progress.lastItemBytesTransferred, std::memory_order_acq_rel);
                progress.lastItemBytesTransferred = fileBytes;
            }
        }
        else
        {
            context.completedBytes = progress.itemBaseBytes + fileBytes;
        }

        return S_OK;
    }

    // Directory reparse points are handled explicitly to prevent recursive traversal (junction/symlink loops).
    ReparsePointData reparse{};
    hr = ReadReparsePointData(source.extended, reparse);
    if (FAILED(hr))
    {
        return returnFailure(hr);
    }

    if (reparse.tag != IO_REPARSE_TAG_SYMLINK && reparse.tag != IO_REPARSE_TAG_MOUNT_POINT)
    {
        return returnFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    ParsedDirectoryReparsePoint parsed{};
    if (! ParseDirectoryReparsePoint(reparse, parsed))
    {
        return returnFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
    }

    const DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (! context.allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        hr = RemovePathForOverwrite(context, destination.extended);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (! CreateDirectoryW(destination.extended.c_str(), nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    bool created = true;
    auto cleanup = wil::scope_exit(
        [&]
        {
            if (created)
            {
                static_cast<void>(RemoveDirectoryW(destination.extended.c_str()));
            }
        });

    std::wstring targetPath = ResolveReparseTargetAbsolute(source, parsed);
    if (targetPath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const bool preserveTrailingSeparator = EndsWithSeparator(parsed.substitutePath) || EndsWithSeparator(parsed.printPath);
    if (preserveTrailingSeparator && ! EndsWithSeparator(targetPath))
    {
        targetPath.push_back(L'\\');
    }

    if (! context.reparseRootSourcePath.empty() && ! context.reparseRootDestinationPath.empty())
    {
        std::wstring mappedTargetPath;
        if (TryRetargetPathIntoDestination(targetPath, context.reparseRootSourcePath, context.reparseRootDestinationPath, mappedTargetPath))
        {
            targetPath = std::move(mappedTargetPath);
            if (preserveTrailingSeparator && ! EndsWithSeparator(targetPath))
            {
                targetPath.push_back(L'\\');
            }
        }
    }

    ReparsePointData rebuilt{};
    if (reparse.tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        hr = BuildMountPointReparseData(targetPath, rebuilt);
    }
    else
    {
        bool useRelative           = parsed.isRelative;
        std::wstring symlinkTarget = targetPath;
        if (parsed.isRelative)
        {
            const std::filesystem::path destinationParent = std::filesystem::path(destination.display).parent_path();
            const std::filesystem::path relativeTarget    = std::filesystem::path(targetPath).lexically_relative(destinationParent);
            if (relativeTarget.empty() || relativeTarget.native().empty())
            {
                useRelative = false;
            }
            else
            {
                symlinkTarget = relativeTarget.wstring();
            }
        }

        hr = BuildSymlinkReparseData(symlinkTarget, useRelative, rebuilt);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = WriteReparsePointData(destination.extended, rebuilt);
    if (FAILED(hr))
    {
        return hr;
    }

    created = false;
    return S_OK;
}

HRESULT CopyDirectoryInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, unsigned __int64* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, 0, 0);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
    {
        if (! CreateDirectoryW(destination.extended.c_str(), nullptr))
        {
            return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
        }
    }
    else
    {
        if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
        if (! context.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
    }

    std::wstring searchPattern = AppendPath(source.extended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return S_OK;
        }
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    bool hadFailure = false;
    bool hadSkipped = false;

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        PathInfo childSource{};
        childSource.display  = AppendPath(source.display, data.cFileName);
        childSource.extended = AppendPath(source.extended, data.cFileName);

        PathInfo childDestination{};
        childDestination.display  = AppendPath(destination.display, data.cFileName);
        childDestination.extended = AppendPath(destination.extended, data.cFileName);

        unsigned __int64 childBytes = 0;
        HRESULT childHr             = S_OK;

        const DWORD childAttributes = data.dwFileAttributes;
        const bool childIsDirectory = IsDirectory(childAttributes);
        const bool childIsReparse   = IsReparsePoint(childAttributes);

        for (;;)
        {
            childBytes = 0;
            childHr    = S_OK;

            if (childIsDirectory)
            {
                if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
                {
                    if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                    {
                        hadSkipped = true;
                        childHr    = S_OK;
                    }
                    else
                    {
                        childHr = CopyReparsePointInternal(context, childSource, childDestination, childAttributes, &childBytes);
                    }
                }
                else if (! context.recursive)
                {
                    childHr = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                }
                else
                {
                    childHr = CopyDirectoryInternal(context, childSource, childDestination, &childBytes);
                }
            }
            else if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    hadSkipped = true;
                    childHr    = S_OK;
                }
                else
                {
                    childHr = CopyReparsePointInternal(context, childSource, childDestination, childAttributes, &childBytes);
                }
            }
            else
            {
                childHr = CopyFileInternal(context, childSource, childDestination, &childBytes);
            }

            if (SUCCEEDED(childHr))
            {
                break;
            }

            childHr = NormalizeCancellation(childHr);
            if (IsCancellationHr(childHr))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (context.continueOnError)
            {
                hadFailure = true;
                childHr    = S_OK;
                break;
            }

            FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
            const HRESULT issueHr             = ReportIssue(context, childHr, &issueAction);
            if (FAILED(issueHr))
            {
                return issueHr;
            }

            switch (issueAction)
            {
                case FileSystemIssueAction::Overwrite: context.allowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceReadOnly: context.allowReplaceReadonly = true; continue;
                case FileSystemIssueAction::PermanentDelete: context.useRecycleBin = false; continue;
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip:
                    hadFailure = true;
                    childHr    = S_OK;
                    break;
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            break;
        }

        if (childBytes > 0)
        {
            if (std::numeric_limits<unsigned __int64>::max() - *bytesCopied < childBytes)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            *bytesCopied += childBytes;
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    if (hadFailure || hadSkipped)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT CopyPathInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, unsigned __int64* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    const DWORD attributes = GetFileAttributesW(source.extended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD error = GetLastError();
        static_cast<void>(SetProgressPaths(context, source.display.c_str(), destination.display.c_str()));
        static_cast<void>(ReportProgressForced(context, 0, 0));
        return HRESULT_FROM_WIN32(error);
    }

    const bool isReparse = IsReparsePoint(attributes);
    if (isReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
    {
        if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return CopyReparsePointInternal(context, source, destination, attributes, bytesCopied);
    }

    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        if (! context.recursive)
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }
        return CopyDirectoryInternal(context, source, destination, bytesCopied);
    }

    return CopyFileInternal(context, source, destination, bytesCopied);
}

struct DirectoryChildWorkItem
{
    std::wstring name;
};

[[nodiscard]] HRESULT CopyDirectoryChildrenParallel(OperationContext& rootContext,
                                                   const PathInfo& source,
                                                   const PathInfo& destination,
                                                   FileSystemFlags flags,
                                                   FileSystemReparsePointPolicy reparsePointPolicy,
                                                   unsigned int maxConcurrency,
                                                   unsigned __int64* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(rootContext, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(rootContext);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(rootContext, 0, 0);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    std::wstring searchPattern = AppendPath(source.extended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(
        FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
        }
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    std::vector<DirectoryChildWorkItem> work;
    work.reserve(128);

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        DirectoryChildWorkItem item{};
        item.name.assign(data.cFileName);
        work.push_back(std::move(item));
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD enumError = GetLastError();
    if (enumError != ERROR_NO_MORE_FILES)
    {
        return returnFailure(HRESULT_FROM_WIN32(enumError));
    }

    if (work.empty())
    {
        return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
    }

    const unsigned int concurrency = std::max(1u, std::min<unsigned int>(maxConcurrency, static_cast<unsigned int>(work.size())));
    if (concurrency <= 1u)
    {
        return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
    }

    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
    {
        if (! CreateDirectoryW(destination.extended.c_str(), nullptr))
        {
            return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
        }
    }
    else
    {
        if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
        if (! rootContext.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
    }

    FileSystemOptions* sharedOptionsState = rootContext.options;

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState ? sharedOptionsState->bandwidthLimitBytesPerSecond : 0ull, std::memory_order_release);

    std::atomic<size_t> nextIndex{0};
    std::atomic<bool> hadFailure{false};
    std::atomic<bool> hadSkipped{false};

    const std::wstring rootSource = rootContext.reparseRootSourcePath;
    const std::wstring rootDestination = rootContext.reparseRootDestinationPath;

    std::vector<std::jthread> workers;
    workers.reserve(concurrency);
    for (unsigned int worker = 0; worker < concurrency; ++worker)
    {
        const unsigned __int64 streamId = static_cast<unsigned __int64>(worker);
        workers.emplace_back(
            [&, streamId]() noexcept
            {
                OperationContext context{};
                InitializeOperationContext(context,
                                           FILESYSTEM_COPY,
                                           flags,
                                           sharedOptionsState,
                                           rootContext.callback,
                                           rootContext.callbackCookie,
                                           1,
                                           reparsePointPolicy);
                context.options           = sharedOptionsState;
                context.parallel          = &parallel;
                context.totalBytes        = 0; // let the host provide totals via pre-calc
                context.progressStreamId  = streamId;
                context.reparseRootSourcePath      = rootSource;
                context.reparseRootDestinationPath = rootDestination;

                for (;;)
                {
                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    const size_t index = nextIndex.fetch_add(1u, std::memory_order_acq_rel);
                    if (index >= work.size())
                    {
                        return;
                    }

                    const DirectoryChildWorkItem& item = work[index];

                    PathInfo childSource{};
                    childSource.display  = AppendPath(source.display, item.name);
                    childSource.extended = AppendPath(source.extended, item.name);

                    PathInfo childDestination{};
                    childDestination.display  = AppendPath(destination.display, item.name);
                    childDestination.extended = AppendPath(destination.extended, item.name);

                    HRESULT itemHr = S_OK;
                    unsigned __int64 childBytes = 0;

                    for (;;)
                    {
                        childBytes = 0;
                        itemHr     = CopyPathInternal(context, childSource, childDestination, &childBytes);
                        if (SUCCEEDED(itemHr))
                        {
                            break;
                        }

                        itemHr = NormalizeCancellation(itemHr);
                        if (IsCancellationHr(itemHr))
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                            return;
                        }

                        if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                        {
                            hadSkipped.store(true, std::memory_order_release);
                            itemHr = S_OK;
                            break;
                        }

                        if (context.continueOnError)
                        {
                            hadFailure.store(true, std::memory_order_release);
                            itemHr = S_OK;
                            break;
                        }

                        FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
                        const HRESULT issueHr             = ReportIssue(context, itemHr, &issueAction);
                        if (FAILED(issueHr))
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                            return;
                        }

                        switch (issueAction)
                        {
                            case FileSystemIssueAction::Overwrite: context.allowOverwrite = true; continue;
                            case FileSystemIssueAction::ReplaceReadOnly: context.allowReplaceReadonly = true; continue;
                            case FileSystemIssueAction::PermanentDelete: context.useRecycleBin = false; continue;
                            case FileSystemIssueAction::Retry: continue;
                            case FileSystemIssueAction::Skip:
                                hadFailure.store(true, std::memory_order_release);
                                itemHr = S_OK;
                                break;
                            case FileSystemIssueAction::Cancel:
                            case FileSystemIssueAction::None:
                            default:
                                parallel.cancelRequested.store(true, std::memory_order_release);
                                return;
                        }

                        break;
                    }
                }
            });
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    *bytesCopied = parallel.completedBytes.load(std::memory_order_acquire);

    if (hadFailure.load(std::memory_order_acquire) || hadSkipped.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT DeletePathInternal(OperationContext& context, const PathInfo& path) noexcept;

HRESULT MovePathInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, bool allowCopy) noexcept
{
    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const DWORD sourceAttributes = GetFileAttributesW(source.extended.c_str());
    if (sourceAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const bool sourceIsDirectory = (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const bool sourceIsReparse   = IsReparsePoint(sourceAttributes);

    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (! context.allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! context.allowReplaceReadonly)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            const DWORD newAttributes = destinationAttributes & ~FILE_ATTRIBUTE_READONLY;
            if (! SetFileAttributesW(destination.extended.c_str(), newAttributes))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
    }

    DWORD moveFlags = 0;
    if (context.allowOverwrite)
    {
        moveFlags |= MOVEFILE_REPLACE_EXISTING;
    }
    if (allowCopy)
    {
        // Attempt a simple rename first; only fall back to copy+delete when required.
    }

    // Reparse-point policies apply to move operations, not rename.
    if (context.type == FILESYSTEM_MOVE && sourceIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
    {
        if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        unsigned __int64 copiedBytes = 0;
        HRESULT copyHr               = CopyReparsePointInternal(context, source, destination, sourceAttributes, &copiedBytes);
        if (FAILED(copyHr))
        {
            return copyHr;
        }

        if (sourceIsDirectory)
        {
            if (! RemoveDirectoryW(source.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
        else
        {
            DWORD newAttributes = sourceAttributes;
            if ((newAttributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                if (! context.allowReplaceReadonly)
                {
                    return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                }

                newAttributes &= ~FILE_ATTRIBUTE_READONLY;
                if (! SetFileAttributesW(source.extended.c_str(), newAttributes))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
            }

            if (! DeleteFileW(source.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }

        return S_OK;
    }

    const DWORD renameFlags = moveFlags;

    CopyProgressContext progress{};
    progress.context = &context;
    if (! context.parallel)
    {
        progress.itemBaseBytes = context.completedBytes;
        progress.startTick     = GetTickCount64();
    }

    if (MoveFileWithProgressW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, renameFlags))
    {
        return S_OK;
    }

    DWORD error = GetLastError();
    if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (! allowCopy || error != ERROR_NOT_SAME_DEVICE)
    {
        return HRESULT_FROM_WIN32(error);
    }

    // Cross-volume move fallback: copy with reparse policy applied, then best-effort delete.
    if (sourceIsDirectory && ! context.recursive)
    {
        return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
    }

    unsigned __int64 bytesCopied = 0;
    const HRESULT copyHr         = CopyPathInternal(context, source, destination, &bytesCopied);
    if (FAILED(copyHr))
    {
        // If we only partially copied, do not delete source. This preserves move safety semantics for skipped items.
        return copyHr;
    }

    struct DeletePhaseCallback final : IFileSystemCallback
    {
        IFileSystemCallback* inner = nullptr;

        explicit DeletePhaseCallback(IFileSystemCallback* callback) noexcept : inner(callback)
        {
        }

        HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation /*operationType*/,
                                                     unsigned long /*totalItems*/,
                                                     unsigned long /*completedItems*/,
                                                     unsigned __int64 /*totalBytes*/,
                                                     unsigned __int64 /*completedBytes*/,
                                                     const wchar_t* /*currentSourcePath*/,
                                                     const wchar_t* /*currentDestinationPath*/,
                                                     unsigned __int64 /*currentItemTotalBytes*/,
                                                     unsigned __int64 /*currentItemCompletedBytes*/,
                                                     FileSystemOptions* /*options*/,
                                                     unsigned __int64 /*progressStreamId*/,
                                                     void* /*cookie*/) noexcept override
        {
            // Suppress delete-phase progress reporting for move operations (move progress reflects copy bytes).
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                          unsigned long /*itemIndex*/,
                                                          const wchar_t* /*sourcePath*/,
                                                          const wchar_t* /*destinationPath*/,
                                                          HRESULT /*status*/,
                                                          FileSystemOptions* /*options*/,
                                                          void* /*cookie*/) noexcept override
        {
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override
        {
            if (! inner)
            {
                if (pCancel)
                {
                    *pCancel = FALSE;
                }
                return S_OK;
            }
            return inner->FileSystemShouldCancel(pCancel, cookie);
        }

        HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                  const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  FileSystemOptions* options,
                                                  void* cookie) noexcept override
        {
            if (! inner)
            {
                if (action)
                {
                    *action = FileSystemIssueAction::Cancel;
                }
                return S_OK;
            }
            return inner->FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie);
        }
    };

    DeletePhaseCallback deleteCallback(context.callback);

    OperationContext deleteContext{};
    deleteContext.type                   = FILESYSTEM_DELETE;
    deleteContext.callback               = &deleteCallback;
    deleteContext.callbackCookie         = context.callbackCookie;
    deleteContext.options                = nullptr;
    deleteContext.totalItems             = 0;
    deleteContext.completedItems         = 0;
    deleteContext.totalBytes             = 0;
    deleteContext.completedBytes         = 0;
    deleteContext.continueOnError        = false;
    deleteContext.allowOverwrite         = false;
    deleteContext.allowReplaceReadonly   = context.allowReplaceReadonly;
    deleteContext.recursive              = true;
    deleteContext.useRecycleBin          = false;
    deleteContext.parallel               = nullptr;
    deleteContext.lastProgressReportTick = 0;

    const HRESULT deleteHr = DeletePathInternal(deleteContext, source);
    if (FAILED(deleteHr))
    {
        return deleteHr;
    }

    return S_OK;
}

class RecycleBinDeleteProgressSink final : public IFileOperationProgressSink
{
public:
    explicit RecycleBinDeleteProgressSink(OperationContext& context) noexcept : _context(&context)
    {
    }
    RecycleBinDeleteProgressSink(const RecycleBinDeleteProgressSink&)            = delete;
    RecycleBinDeleteProgressSink(RecycleBinDeleteProgressSink&&)                 = delete;
    RecycleBinDeleteProgressSink& operator=(const RecycleBinDeleteProgressSink&) = delete;
    RecycleBinDeleteProgressSink& operator=(RecycleBinDeleteProgressSink&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileOperationProgressSink))
        {
            *ppvObject = static_cast<IFileOperationProgressSink*>(this);
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

    HRESULT STDMETHODCALLTYPE StartOperations() noexcept override
    {
        if (_context != nullptr)
        {
            _baseCompletedItems = _context->parallel ? _context->parallel->completedItems.load(std::memory_order_acquire) : _context->completedItems;
            _baseTotalItems     = _context->totalItems;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FinishOperations([[maybe_unused]] HRESULT hrResult) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreRenameItem([[maybe_unused]] DWORD flags, [[maybe_unused]] IShellItem* item, [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostRenameItem([[maybe_unused]] DWORD flags,
                                             [[maybe_unused]] IShellItem* item,
                                             [[maybe_unused]] LPCWSTR newName,
                                             [[maybe_unused]] HRESULT hrRename,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreMoveItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostMoveItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrMove,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreCopyItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostCopyItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrCopy,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreDeleteItem([[maybe_unused]] DWORD flags, [[maybe_unused]] IShellItem* item) noexcept override
    {
        const HRESULT hr = ReportItemPath(item, false);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostDeleteItem([[maybe_unused]] DWORD flags,
                                             IShellItem* item,
                                             HRESULT hrDelete,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        if (SUCCEEDED(hrDelete) && _context != nullptr)
        {
            if (! _workProgressAvailable)
            {
                AddCompletedItems(*_context, 1);
            }
            const HRESULT hr = ReportItemPath(item, false);
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

        if (SUCCEEDED(hrDelete) || FAILED(_firstError))
        {
            return S_OK;
        }

        _firstError = hrDelete;
        if (item != nullptr)
        {
            wil::unique_cotaskmem_string path;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, path.put())) && path && path.get()[0] != L'\0')
            {
                _firstErrorPath.assign(path.get());
            }
            else
            {
                path.reset();
                if (SUCCEEDED(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, path.put())) && path && path.get()[0] != L'\0')
                {
                    _firstErrorPath.assign(path.get());
                }
            }
        }

        static_cast<void>(ReportItemPath(item, true));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreNewItem([[maybe_unused]] DWORD flags,
                                         [[maybe_unused]] IShellItem* destinationFolder,
                                         [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostNewItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName,
                                          [[maybe_unused]] LPCWSTR templateName,
                                          [[maybe_unused]] DWORD fileAttributes,
                                          [[maybe_unused]] HRESULT hrNew,
                                          [[maybe_unused]] IShellItem* newItem) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE UpdateProgress(UINT workTotal, UINT workSoFar) noexcept override
    {
        if (_context == nullptr)
        {
            return S_OK;
        }

        if (workTotal > 0 || workSoFar > 0)
        {
            _workProgressAvailable = true;
        }

        if (_workProgressAvailable)
        {
            constexpr unsigned __int64 maxUlong       = static_cast<unsigned __int64>(std::numeric_limits<unsigned long>::max());
            const unsigned __int64 desiredTotal64     = static_cast<unsigned __int64>(_baseCompletedItems) + static_cast<unsigned __int64>(workTotal);
            const unsigned __int64 desiredCompleted64 = static_cast<unsigned __int64>(_baseCompletedItems) + static_cast<unsigned __int64>(workSoFar);

            const unsigned long desiredTotal     = static_cast<unsigned long>(std::min(desiredTotal64, maxUlong));
            const unsigned long desiredCompleted = static_cast<unsigned long>(std::min(desiredCompleted64, maxUlong));

            _context->totalItems = std::max(_context->totalItems, desiredTotal);
            if (_context->parallel)
            {
                unsigned long current = _context->parallel->completedItems.load(std::memory_order_acquire);
                while (current < desiredCompleted &&
                       ! _context->parallel->completedItems.compare_exchange_weak(current, desiredCompleted, std::memory_order_acq_rel))
                {
                }
            }
            else
            {
                _context->completedItems = std::max(_context->completedItems, desiredCompleted);
            }
        }

        const HRESULT hr = ReportProgress(*_context, 0, 0);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResetTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PauseTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResumeTimer() noexcept override
    {
        return S_OK;
    }

    [[nodiscard]] HRESULT GetFirstError() const noexcept
    {
        return _firstError;
    }

    [[nodiscard]] const std::wstring& GetFirstErrorPath() const noexcept
    {
        return _firstErrorPath;
    }

private:
    [[nodiscard]] HRESULT ReportItemPath(IShellItem* item, bool force) noexcept
    {
        if (item == nullptr || _context == nullptr)
        {
            return S_OK;
        }

        wil::unique_cotaskmem_string path;
        if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, path.put())) || ! path || path.get()[0] == L'\0')
        {
            path.reset();
            static_cast<void>(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, path.put()));
        }

        if (! path || path.get()[0] == L'\0')
        {
            return S_OK;
        }

        const HRESULT hrPaths = SetProgressPaths(*_context, path.get(), nullptr);
        if (FAILED(hrPaths))
        {
            return hrPaths;
        }

        return force ? ReportProgressForced(*_context, 0, 0) : ReportProgress(*_context, 0, 0);
    }

    ~RecycleBinDeleteProgressSink() = default;

    std::atomic_ulong _refCount{1};
    OperationContext* _context        = nullptr;
    unsigned long _baseCompletedItems = 0;
    unsigned long _baseTotalItems     = 0;
    bool _workProgressAvailable       = false;
    HRESULT _firstError               = S_OK;
    std::wstring _firstErrorPath;
};

HRESULT DeleteToRecycleBin(OperationContext& context, const PathInfo& path) noexcept
{
    if (path.display.empty())
    {
        return E_INVALIDARG;
    }

    // The host/plugin task threads already initialize COM. We still try here because
    // DeleteToRecycleBin can also be exercised from test paths that don't guarantee it.
    const HRESULT coInitHr   = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool coInitialized = SUCCEEDED(coInitHr) || coInitHr == S_FALSE;
    auto coUninit            = wil::scope_exit(
        [&]() noexcept
        {
            if (coInitialized)
            {
                CoUninitialize();
            }
        });
    if (FAILED(coInitHr) && coInitHr != RPC_E_CHANGED_MODE)
    {
        return coInitHr;
    }

    wil::com_ptr<IFileOperation> fileOperation;
    HRESULT hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(fileOperation.put()));
    if (FAILED(hr) || ! fileOperation)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    constexpr DWORD kOperationFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT | FOFX_EARLYFAILURE | FOFX_RECYCLEONDELETE;
    hr                              = fileOperation->SetOperationFlags(kOperationFlags);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IShellItem> item;
    hr = SHCreateItemFromParsingName(path.display.c_str(), nullptr, IID_PPV_ARGS(item.put()));
    if (FAILED(hr) || ! item)
    {
        return FAILED(hr) ? hr : E_INVALIDARG;
    }

    wil::com_ptr<IFileOperationProgressSink> progressSink;
    auto* progressSinkImpl = new (std::nothrow) RecycleBinDeleteProgressSink(context);
    if (! progressSinkImpl)
    {
        return E_OUTOFMEMORY;
    }
    progressSink.attach(progressSinkImpl);

    DWORD adviseCookie = 0;
    hr                 = fileOperation->Advise(progressSink.get(), &adviseCookie);
    if (FAILED(hr))
    {
        return hr;
    }
    auto unadvise = wil::scope_exit(
        [&]() noexcept
        {
            if (adviseCookie != 0)
            {
                static_cast<void>(fileOperation->Unadvise(adviseCookie));
            }
        });

    hr = fileOperation->DeleteItem(item.get(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = fileOperation->PerformOperations();
    if (FAILED(hr))
    {
        const HRESULT itemError = progressSinkImpl->GetFirstError();
        if (FAILED(itemError))
        {
            const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
            const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
            Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
            return itemError;
        }

        return hr;
    }

    BOOL anyAborted = FALSE;
    hr              = fileOperation->GetAnyOperationsAborted(&anyAborted);
    if (FAILED(hr))
    {
        const HRESULT itemError = progressSinkImpl->GetFirstError();
        if (FAILED(itemError))
        {
            const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
            const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
            Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
            return itemError;
        }
        return hr;
    }

    if (anyAborted == TRUE)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const HRESULT itemError = progressSinkImpl->GetFirstError();
    if (FAILED(itemError))
    {
        const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
        const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
        Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
        return itemError;
    }

    static_cast<void>(ReportProgressForced(context, 0, 0));
    return S_OK;
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept;

HRESULT DeletePathInternal(OperationContext& context, const PathInfo& path) noexcept
{
    HRESULT hr = SetProgressPaths(context, path.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgress(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.useRecycleBin)
    {
        return DeleteToRecycleBin(context, path);
    }

    DWORD attributes = GetFileAttributesW(path.extended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        // Never traverse directory reparse points during delete recursion (junction/symlink safety).
        if (IsReparsePoint(attributes))
        {
            if (! RemoveDirectoryW(path.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            AddCompletedItems(context, 1);
            return S_OK;
        }

        if (! context.recursive)
        {
            if (! RemoveDirectoryW(path.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            AddCompletedItems(context, 1);
            return S_OK;
        }

        return DeleteDirectoryRecursive(context, path);
    }

    unsigned __int64 fileBytes = 0;
    static_cast<void>(GetFileSizeBytes(path.extended, &fileBytes)); // Best-effort only.

    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(path.extended.c_str(), newAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    if (! DeleteFileW(path.extended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    AddCompletedItems(context, 1);
    AddCompletedBytes(context, fileBytes);

    return S_OK;
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept
{
    std::wstring searchPattern = AppendPath(path.extended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error);
    }

    bool hadFailure = false;

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        PathInfo child{};
        child.display  = AppendPath(path.display, data.cFileName);
        child.extended = AppendPath(path.extended, data.cFileName);

        HRESULT childHr = DeletePathInternal(context, child);
        if (FAILED(childHr))
        {
            if (childHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return childHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return childHr;
            }
        }

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return HRESULT_FROM_WIN32(error);
    }

    if (! RemoveDirectoryW(path.extended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    AddCompletedItems(context, 1);

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::CopyItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options,
                                               IFileSystemCallback* callback,
                                               void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1u;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: CopyItem failed to set paths for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(hr));
        return hr;
    }

    context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
    context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

    unsigned __int64 bytesCopied = 0;
    HRESULT itemHr               = S_OK;

    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const DWORD attributes            = GetFileAttributesW(source.extended.c_str());
    const bool canParallelizeDirectory =
        attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && ! IsReparsePoint(attributes) && context.recursive &&
        maxConcurrency > 1u;

    if (canParallelizeDirectory)
    {
        itemHr = CopyDirectoryChildrenParallel(context, source, destination, flags, reparsePointPolicy, maxConcurrency, &bytesCopied);
    }
    else
    {
        itemHr = CopyPathInternal(context, source, destination, &bytesCopied);
    }
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: CopyItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::MoveItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options,
                                               IFileSystemCallback* callback,
                                               void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed to set paths for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(hr));
        return hr;
    }

    HRESULT itemHr = MovePathInternal(context, source, destination, true);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE
FileSystem::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);

    const PathInfo target = MakePathInfo(path);

    HRESULT hr = SetItemPaths(context, target.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: DeleteItem failed to set path for '{}' (hr={:#x})", target.display, static_cast<uint32_t>(hr));
        return hr;
    }

    HRESULT itemHr = DeletePathInternal(context, target);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: DeleteItem failed for '{}' (hr={:#x})", target.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgressForced(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::RenameItem(const wchar_t* sourcePath,
                                                 const wchar_t* destinationPath,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options,
                                                 IFileSystemCallback* callback,
                                                 void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = MovePathInternal(context, source, destination, false);
    hr             = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::CopyItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, count, reparsePointPolicy);

        bool hadFailure = false;

        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            if (! sourcePath)
            {
                return E_POINTER;
            }

            if (sourcePath[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const std::wstring_view leaf = GetPathLeaf(sourcePath);
            if (leaf.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }

            const PathInfo source      = MakePathInfo(sourcePath);
            const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

            HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
            if (FAILED(hr))
            {
                return hr;
            }

            context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
            context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

            unsigned __int64 bytesCopied = 0;
            HRESULT itemHr               = CopyPathInternal(context, source, destination, &bytesCopied);

            hr = ReportItemCompleted(context, index, itemHr);
            if (FAILED(hr))
            {
                return hr;
            }

            context.completedItems += 1;

            if (FAILED(itemHr))
            {
                if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    return itemHr;
                }

                hadFailure = true;
                if (! context.continueOnError)
                {
                    return itemHr;
                }
            }
        }

        if (hadFailure)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
    }

    FileSystemOptions sharedOptionsState{};
    if (options)
    {
        sharedOptionsState = *options;
    }

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState.bandwidthLimitBytesPerSecond, std::memory_order_release);

    std::atomic<unsigned long> nextIndex{0};

    std::vector<std::jthread> workers;
    workers.reserve(concurrency);

    for (unsigned int worker = 0; worker < concurrency; ++worker)
    {
        const unsigned __int64 streamId = static_cast<unsigned __int64>(worker);
        workers.emplace_back(
            [&, streamId]() noexcept
            {
                OperationContext context{};
                InitializeOperationContext(context, FILESYSTEM_COPY, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
                context.options    = &sharedOptionsState;
                context.parallel   = &parallel;
                context.totalBytes = 0; // let the host provide totals via pre-calc
                context.progressStreamId = streamId;

                for (;;)
                {
                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    const unsigned long index = nextIndex.fetch_add(1u, std::memory_order_acq_rel);
                    if (index >= count)
                    {
                        return;
                    }

                    const wchar_t* sourcePath    = sourcePaths[index];
                    const std::wstring_view leaf = GetPathLeaf(sourcePath);

                    HRESULT hr = CheckCancel(context);
                    if (FAILED(hr))
                    {
                        const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                        if (! stopOnError)
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        return;
                    }

                    const PathInfo source      = MakePathInfo(sourcePath);
                    const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

                    hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
                    if (FAILED(hr))
                    {
                        parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                        HRESULT expected = S_OK;
                        static_cast<void>(parallel.firstError.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
                        return;
                    }

                    context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
                    context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

                    unsigned __int64 bytesCopied = 0;
                    HRESULT itemHr               = CopyPathInternal(context, source, destination, &bytesCopied);

                    hr = ReportItemCompleted(context, index, itemHr);
                    if (FAILED(hr))
                    {
                        const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                        if (! stopOnError)
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        return;
                    }

                    parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);

                    if (FAILED(itemHr))
                    {
                        if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                        {
                            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                            if (! stopOnError)
                            {
                                parallel.cancelRequested.store(true, std::memory_order_release);
                            }
                            return;
                        }

                        parallel.hadFailure.store(true, std::memory_order_release);
                        if (! context.continueOnError)
                        {
                            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                            return;
                        }
                    }
                }
            });
    }

    workers.clear();

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::MoveItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, count, reparsePointPolicy);

        bool hadFailure = false;

        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            if (! sourcePath)
            {
                return E_POINTER;
            }

            if (sourcePath[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const std::wstring_view leaf = GetPathLeaf(sourcePath);
            if (leaf.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }

            const PathInfo source      = MakePathInfo(sourcePath);
            const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

            HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
            if (FAILED(hr))
            {
                return hr;
            }

            HRESULT itemHr = MovePathInternal(context, source, destination, true);
            hr             = ReportItemCompleted(context, index, itemHr);
            if (FAILED(hr))
            {
                return hr;
            }

            context.completedItems += 1;

            if (FAILED(itemHr))
            {
                if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    return itemHr;
                }

                hadFailure = true;
                if (! context.continueOnError)
                {
                    return itemHr;
                }
            }
        }

        if (hadFailure)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
    }

    FileSystemOptions sharedOptionsState{};
    if (options)
    {
        sharedOptionsState = *options;
    }

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState.bandwidthLimitBytesPerSecond, std::memory_order_release);

    std::atomic<unsigned long> nextIndex{0};

    std::vector<std::jthread> workers;
    workers.reserve(concurrency);

    for (unsigned int worker = 0; worker < concurrency; ++worker)
    {
        const unsigned __int64 streamId = static_cast<unsigned __int64>(worker);
        workers.emplace_back(
            [&, streamId]() noexcept
            {
                OperationContext context{};
                InitializeOperationContext(context, FILESYSTEM_MOVE, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
                context.options    = &sharedOptionsState;
                context.parallel   = &parallel;
                context.totalBytes = 0; // let the host provide totals via pre-calc
                context.progressStreamId = streamId;

                for (;;)
                {
                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    const unsigned long index = nextIndex.fetch_add(1u, std::memory_order_acq_rel);
                    if (index >= count)
                    {
                        return;
                    }

                    const wchar_t* sourcePath    = sourcePaths[index];
                    const std::wstring_view leaf = GetPathLeaf(sourcePath);

                    HRESULT hr = CheckCancel(context);
                    if (FAILED(hr))
                    {
                        const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                        if (! stopOnError)
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        return;
                    }

                    const PathInfo source      = MakePathInfo(sourcePath);
                    const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

                    hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
                    if (FAILED(hr))
                    {
                        parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                        HRESULT expected = S_OK;
                        static_cast<void>(parallel.firstError.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
                        return;
                    }

                    const HRESULT itemHr = MovePathInternal(context, source, destination, true);

                    hr = ReportItemCompleted(context, index, itemHr);
                    if (FAILED(hr))
                    {
                        const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                        if (! stopOnError)
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        return;
                    }

                    parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);

                    if (FAILED(itemHr))
                    {
                        if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                        {
                            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                            if (! stopOnError)
                            {
                                parallel.cancelRequested.store(true, std::memory_order_release);
                            }
                            return;
                        }

                        parallel.hadFailure.store(true, std::memory_order_release);
                        if (! context.continueOnError)
                        {
                            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                            return;
                        }
                    }
                }
            });
    }

    workers.clear();

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::DeleteItems(const wchar_t* const* paths,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (! paths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int deleteMaxConcurrency               = 1;
    unsigned int deleteRecycleBinMaxConcurrency     = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy             = _reparsePointPolicy;
        deleteMaxConcurrency           = _deleteMaxConcurrency;
        deleteRecycleBinMaxConcurrency = _deleteRecycleBinMaxConcurrency;
    }

    const bool useRecycleBin = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);

    const unsigned int maxConcurrencyFast       = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    const unsigned int maxConcurrencyRecycleBin = std::clamp(deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    const unsigned int maxConcurrency           = useRecycleBin ? maxConcurrencyRecycleBin : maxConcurrencyFast;
    const unsigned int concurrency              = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency > 1u)
    {
        std::vector<std::wstring> extendedPaths;
        extendedPaths.reserve(count);
        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* path = paths[index];
            if (! path)
            {
                return E_POINTER;
            }

            if (path[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const PathInfo target = MakePathInfo(path);
            extendedPaths.emplace_back(target.extended);
        }

        const auto isPrefixPath = [](const std::wstring& prefix, const std::wstring& candidate) noexcept -> bool
        {
            if (prefix.empty() || candidate.empty())
            {
                return false;
            }

            if (prefix.size() > candidate.size())
            {
                return false;
            }

            if (_wcsnicmp(prefix.c_str(), candidate.c_str(), prefix.size()) != 0)
            {
                return false;
            }

            if (candidate.size() == prefix.size())
            {
                return true;
            }

            const wchar_t last = prefix.back();
            if (last == L'\\' || last == L'/')
            {
                return true;
            }

            const wchar_t next = candidate[prefix.size()];
            return next == L'\\' || next == L'/';
        };

        std::vector<size_t> order;
        order.reserve(extendedPaths.size());
        for (size_t i = 0; i < extendedPaths.size(); ++i)
        {
            order.emplace_back(i);
        }

        std::ranges::sort(order, [&](size_t a, size_t b) noexcept { return _wcsicmp(extendedPaths[a].c_str(), extendedPaths[b].c_str()) < 0; });

        // Build a dependency graph for overlapping inputs:
        // - If A is a prefix of B, we must delete B before A to avoid parent/child races.
        // We only depend on the *immediate* ancestor; transitive ordering falls out naturally.
        std::vector<unsigned long> remainingDeps(static_cast<size_t>(count), 0u);
        std::vector<std::vector<unsigned long>> dependents(static_cast<size_t>(count));

        std::vector<unsigned long> stack;
        stack.reserve(order.size());
        for (const size_t index : order)
        {
            const unsigned long cur = static_cast<unsigned long>(index);

            while (! stack.empty())
            {
                const unsigned long parent = stack.back();
                if (isPrefixPath(extendedPaths[parent], extendedPaths[cur]))
                {
                    break;
                }
                stack.pop_back();
            }

            if (! stack.empty())
            {
                const unsigned long parent = stack.back();
                ++remainingDeps[parent];
                dependents[cur].push_back(parent);
            }

            stack.push_back(cur);
        }

        std::deque<unsigned long> ready;
        for (unsigned long i = 0; i < count; ++i)
        {
            if (remainingDeps[i] == 0)
            {
                ready.push_back(i);
            }
        }

        FileSystemOptions sharedOptionsState{};
        if (options)
        {
            sharedOptionsState = *options;
        }

        ParallelOperationState parallel{};
        parallel.startTick = GetTickCount64();

        std::mutex scheduleMutex;
        std::condition_variable scheduleCv;
        unsigned long remainingWork = count;

        std::vector<std::jthread> workers;
        workers.reserve(concurrency);

        for (unsigned int worker = 0; worker < concurrency; ++worker)
        {
            const unsigned __int64 streamId = static_cast<unsigned __int64>(worker);
            workers.emplace_back(
                [&, streamId]() noexcept
                {
                    [[maybe_unused]] auto coInit = wil::CoInitializeEx();

                    OperationContext context{};
                    // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
                    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, &sharedOptionsState, callback, cookie, 0, reparsePointPolicy);
                    context.options    = &sharedOptionsState;
                    context.parallel   = &parallel;
                    context.totalBytes = 0; // host pre-calc provides totals when available
                    context.progressStreamId = streamId;

                    for (;;)
                    {
                        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            return;
                        }

                        unsigned long index = 0;
                        {
                            std::unique_lock lock(scheduleMutex);
                            scheduleCv.wait(lock,
                                            [&] noexcept
                                            {
                                                return parallel.cancelRequested.load(std::memory_order_acquire) ||
                                                       parallel.stopOnErrorRequested.load(std::memory_order_acquire) || remainingWork == 0 || ! ready.empty();
                                            });

                            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                            {
                                return;
                            }

                            if (remainingWork == 0)
                            {
                                return;
                            }

                            if (ready.empty())
                            {
                                continue;
                            }

                            index = ready.front();
                            ready.pop_front();
                        }

                        const wchar_t* path = paths[index];
                        if (! path || path[0] == L'\0')
                        {
                            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(
                                parallel.firstError.compare_exchange_strong(expected, path ? E_INVALIDARG : E_POINTER, std::memory_order_acq_rel));
                            scheduleCv.notify_all();
                            return;
                        }

                        const PathInfo target = MakePathInfo(path);

                        HRESULT hr = SetItemPaths(context, target.display.c_str(), nullptr);
                        if (FAILED(hr))
                        {
                            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
                            scheduleCv.notify_all();
                            return;
                        }

                        const HRESULT itemHr = DeletePathInternal(context, target);

                        hr = ReportItemCompleted(context, index, itemHr);
                        if (FAILED(hr))
                        {
                            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                            if (! stopOnError)
                            {
                                parallel.cancelRequested.store(true, std::memory_order_release);
                            }
                            scheduleCv.notify_all();
                            return;
                        }

                        parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);
                        hr = ReportProgress(context, 0, 0);
                        if (FAILED(hr))
                        {
                            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                            if (! stopOnError)
                            {
                                parallel.cancelRequested.store(true, std::memory_order_release);
                            }
                            scheduleCv.notify_all();
                            return;
                        }

                        if (FAILED(itemHr))
                        {
                            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                            {
                                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                                if (! stopOnError)
                                {
                                    parallel.cancelRequested.store(true, std::memory_order_release);
                                }
                                scheduleCv.notify_all();
                                return;
                            }

                            parallel.hadFailure.store(true, std::memory_order_release);
                            if (! context.continueOnError)
                            {
                                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                                HRESULT expected = S_OK;
                                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                                scheduleCv.notify_all();
                                return;
                            }
                        }

                        {
                            std::unique_lock lock(scheduleMutex);
                            for (const unsigned long dependent : dependents[index])
                            {
                                if (remainingDeps[dependent] > 0)
                                {
                                    --remainingDeps[dependent];
                                    if (remainingDeps[dependent] == 0)
                                    {
                                        ready.push_back(dependent);
                                    }
                                }
                            }

                            if (remainingWork > 0)
                            {
                                --remainingWork;
                            }
                        }

                        scheduleCv.notify_all();
                    }
                });
        }

        workers.clear();

        if (parallel.cancelRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.hadFailure.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    OperationContext context{};
    // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* path = paths[index];
        if (! path)
        {
            return E_POINTER;
        }

        if (path[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const PathInfo target = MakePathInfo(path);

        HRESULT hr = SetItemPaths(context, target.display.c_str(), nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = DeletePathInternal(context, target);
        hr             = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;
        hr = ReportProgress(context, 0, 0);
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::RenameItems(const FileSystemRenamePair* items,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (! items && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, count, reparsePointPolicy);

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const FileSystemRenamePair& item = items[index];
        if (! item.sourcePath || ! item.newName)
        {
            return E_POINTER;
        }

        if (item.sourcePath[0] == L'\0' || item.newName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view newName = item.newName;
        if (ContainsPathSeparator(newName))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::wstring directory = GetPathDirectory(item.sourcePath);
        if (directory.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::wstring destinationPath = AppendPath(directory, newName);
        const PathInfo source              = MakePathInfo(item.sourcePath);
        const PathInfo destination         = MakePathInfo(destinationPath);

        HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = MovePathInternal(context, source, destination, false);
        hr             = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

#include "FileSystem.Internal.h"
#include "Helpers.h"
#include "PathUtils.h"
#include "YyjsonHelpers.h"

#include <array>
#include <atomic>
#include <cwctype>
#include <limits>
#include <optional>

#include <shlwapi.h>
#include <shobjidl.h>
#include <winioctl.h>

#include <yyjson.h>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Netapi32.lib")

namespace
{
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16ReplacingInvalid(text);
}

[[nodiscard]] std::string FormatFileTimeLocal(const FILETIME& fileTime) noexcept
{
    if (fileTime.dwLowDateTime == 0u && fileTime.dwHighDateTime == 0u)
    {
        return {};
    }

    FILETIME localFileTime{};
    if (FileTimeToLocalFileTime(&fileTime, &localFileTime) == 0)
    {
        return {};
    }

    SYSTEMTIME localSystemTime{};
    if (FileTimeToSystemTime(&localFileTime, &localSystemTime) == 0)
    {
        return {};
    }

    return Utf8FromUtf16(std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                                     localSystemTime.wYear,
                                     localSystemTime.wMonth,
                                     localSystemTime.wDay,
                                     localSystemTime.wHour,
                                     localSystemTime.wMinute,
                                     localSystemTime.wSecond));
}

[[nodiscard]] std::string FormatItemPropertiesSize(uint64_t sizeBytes)
{
    const std::wstring exactBytes = std::format(L"{} bytes", sizeBytes);
    if (sizeBytes < 1024ull)
    {
        return Utf8FromUtf16(exactBytes);
    }

    return Utf8FromUtf16(std::format(L"{} ({})", FormatBytesCompact(sizeBytes), exactBytes));
}

[[nodiscard]] std::string FormatFileAttributeFlags(DWORD attributes) noexcept
{
    std::wstring text;
    const auto appendFlag = [&](const DWORD flag, const std::wstring_view label) noexcept
    {
        if ((attributes & flag) == 0u)
        {
            return;
        }

        if (! text.empty())
        {
            text.append(L", ");
        }
        text.append(label);
    };

    appendFlag(FILE_ATTRIBUTE_ARCHIVE, L"Archive");
    appendFlag(FILE_ATTRIBUTE_COMPRESSED, L"Compressed");
    appendFlag(FILE_ATTRIBUTE_DIRECTORY, L"Directory");
    appendFlag(FILE_ATTRIBUTE_ENCRYPTED, L"Encrypted");
    appendFlag(FILE_ATTRIBUTE_HIDDEN, L"Hidden");
    appendFlag(FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, L"Not content indexed");
    appendFlag(FILE_ATTRIBUTE_OFFLINE, L"Offline");
    appendFlag(FILE_ATTRIBUTE_READONLY, L"Read-only");
    appendFlag(FILE_ATTRIBUTE_REPARSE_POINT, L"Reparse point");
    appendFlag(FILE_ATTRIBUTE_SPARSE_FILE, L"Sparse");
    appendFlag(FILE_ATTRIBUTE_SYSTEM, L"System");
    appendFlag(FILE_ATTRIBUTE_TEMPORARY, L"Temporary");

    if (text.empty())
    {
        text = L"None";
    }

    return Utf8FromUtf16(text);
}

struct NamedStreamInfo
{
    std::wstring name;
    uint64_t sizeBytes = 0;
};

[[nodiscard]] std::optional<std::wstring> TryExtractNamedStreamName(std::wstring_view win32StreamName)
{
    constexpr std::wstring_view kPrefix            = L":";
    constexpr std::wstring_view kSuffix            = L":$DATA";
    constexpr std::wstring_view kDefaultDataStream = L"::$DATA";

    if (win32StreamName == kDefaultDataStream || ! win32StreamName.starts_with(kPrefix) || ! win32StreamName.ends_with(kSuffix) ||
        win32StreamName.size() <= kPrefix.size() + kSuffix.size())
    {
        return std::nullopt;
    }

    std::wstring name(win32StreamName.substr(kPrefix.size(), win32StreamName.size() - kPrefix.size() - kSuffix.size()));
    if (name.empty())
    {
        return std::nullopt;
    }

    return name;
}

[[nodiscard]] bool IsSafeLogicalStreamName(std::wstring_view streamName) noexcept
{
    if (streamName.empty())
    {
        return false;
    }

    return std::ranges::none_of(streamName, [](wchar_t ch) noexcept { return ch == L':' || ch == L'\\' || ch == L'/' || ch == L'\0'; });
}

HRESULT EnumerateNamedStreams(const wchar_t* path, std::vector<NamedStreamInfo>& streams) noexcept
{
    streams.clear();
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring extendedPath = FileSystemInternal::ToExtendedPath(path);

    WIN32_FIND_STREAM_DATA streamData{};
    wil::unique_hfind findHandle(::FindFirstStreamW(extendedPath.c_str(), FindStreamInfoStandard, &streamData, 0));
    if (! findHandle)
    {
        const DWORD lastError = ::GetLastError();
        switch (lastError)
        {
            case ERROR_HANDLE_EOF:
            case ERROR_INVALID_PARAMETER:
            case ERROR_NOT_SUPPORTED: return S_OK;
            default: return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
        }
    }

    for (;;)
    {
        if (streamData.StreamSize.QuadPart >= 0)
        {
            std::optional<std::wstring> name = TryExtractNamedStreamName(streamData.cStreamName);
            if (name.has_value() && IsSafeLogicalStreamName(name.value()))
            {
                streams.push_back(NamedStreamInfo{.name = std::move(name.value()), .sizeBytes = static_cast<uint64_t>(streamData.StreamSize.QuadPart)});
            }
        }

        streamData = {};
        if (::FindNextStreamW(findHandle.get(), &streamData) == 0)
        {
            const DWORD lastError = ::GetLastError();
            return (lastError == ERROR_HANDLE_EOF) ? S_OK : HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
        }
    }
}

struct ItemPropertiesLinkTargetInfo final
{
    const char* sectionTitle = nullptr;
    std::string kind;
    std::wstring url;
    std::wstring target;
};

struct MountPointReparseDataBufferForProperties final
{
    ULONG ReparseTag            = IO_REPARSE_TAG_MOUNT_POINT;
    USHORT ReparseDataLength    = 0;
    USHORT Reserved             = 0;
    USHORT SubstituteNameOffset = 0;
    USHORT SubstituteNameLength = 0;
    USHORT PrintNameOffset      = 0;
    USHORT PrintNameLength      = 0;
    wchar_t PathBuffer[1]{};
};

struct SymbolicLinkReparseDataBufferForProperties final
{
    ULONG ReparseTag            = IO_REPARSE_TAG_SYMLINK;
    USHORT ReparseDataLength    = 0;
    USHORT Reserved             = 0;
    USHORT SubstituteNameOffset = 0;
    USHORT SubstituteNameLength = 0;
    USHORT PrintNameOffset      = 0;
    USHORT PrintNameLength      = 0;
    ULONG Flags                 = 0;
    wchar_t PathBuffer[1]{};
};

[[nodiscard]] std::wstring TrimShortcutValueForProperties(std::wstring_view value)
{
    while (! value.empty() && (value.front() == L' ' || value.front() == L'\t' || value.front() == L'\r' || value.front() == L'\n'))
    {
        value.remove_prefix(1);
    }
    while (! value.empty() && (value.back() == L' ' || value.back() == L'\t' || value.back() == L'\r' || value.back() == L'\n'))
    {
        value.remove_suffix(1);
    }
    return std::wstring(value);
}

[[nodiscard]] bool IsWindowsAbsolutePathTextForProperties(std::wstring_view text) noexcept
{
    const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(text);
    return pathClass == Common::Paths::WindowsPathClass::DriveAbsolute || pathClass == Common::Paths::WindowsPathClass::Unc;
}

[[nodiscard]] std::optional<std::wstring> ConvertFileUrlToLocalPathForProperties(std::wstring_view url) noexcept
{
    std::wstring urlText(url);
    std::wstring pathText(32768, L'\0');
    DWORD pathCharCount = static_cast<DWORD>(pathText.size());
    const HRESULT hr    = PathCreateFromUrlW(urlText.c_str(), pathText.data(), &pathCharCount, 0);
    if (FAILED(hr))
    {
        return std::nullopt;
    }

    const size_t terminator = pathText.find(L'\0');
    if (terminator != std::wstring::npos)
    {
        pathText.resize(terminator);
    }
    else
    {
        pathText.resize(std::min<size_t>(pathCharCount, pathText.size()));
    }

    if (pathText.empty())
    {
        return std::nullopt;
    }

    return pathText;
}

[[nodiscard]] bool TryReadInternetShortcutUrlForProperties(const wchar_t* shortcutPath, std::wstring& outUrl) noexcept
{
    outUrl.clear();

    std::wstring value(32768, L'\0');
    const DWORD copied = GetPrivateProfileStringW(L"InternetShortcut", L"URL", L"", value.data(), static_cast<DWORD>(value.size()), shortcutPath);
    if (copied == 0 || copied >= (value.size() - 1u))
    {
        return false;
    }

    value.resize(copied);
    outUrl = TrimShortcutValueForProperties(value);
    return ! outUrl.empty();
}

[[nodiscard]] std::optional<std::wstring> ResolveShellLinkTargetForProperties(const wchar_t* shortcutPath) noexcept
{
    const HRESULT coHr      = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool uninitialize = SUCCEEDED(coHr);
    const auto coCleanup    = wil::scope_exit([&]
    {
        if (uninitialize)
        {
            CoUninitialize();
        }
    });
    if (FAILED(coHr) && coHr != RPC_E_CHANGED_MODE)
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellLinkW> shellLink;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.put()));
    if (FAILED(hr) || ! shellLink)
    {
        return std::nullopt;
    }

    wil::com_ptr<IPersistFile> persistFile;
    hr = shellLink->QueryInterface(IID_PPV_ARGS(persistFile.put()));
    if (FAILED(hr) || ! persistFile)
    {
        return std::nullopt;
    }

    hr = persistFile->Load(shortcutPath, STGM_READ);
    if (FAILED(hr))
    {
        return std::nullopt;
    }

    WIN32_FIND_DATAW findData{};
    std::wstring targetText(32768, L'\0');
    hr = shellLink->GetPath(targetText.data(), static_cast<int>(targetText.size()), &findData, SLGP_UNCPRIORITY);
    if (FAILED(hr))
    {
        return std::nullopt;
    }

    const size_t terminator = targetText.find(L'\0');
    if (terminator != std::wstring::npos)
    {
        targetText.resize(terminator);
    }

    if (targetText.empty())
    {
        return std::nullopt;
    }

    return targetText;
}

[[nodiscard]] std::optional<std::wstring> ExtractReparsePathBufferStringForProperties(const wchar_t* pathBuffer,
                                                                                      USHORT substituteNameOffset,
                                                                                      USHORT substituteNameLength,
                                                                                      size_t pathBufferBytes)
{
    if ((substituteNameOffset % sizeof(wchar_t)) != 0 || (substituteNameLength % sizeof(wchar_t)) != 0)
    {
        return std::nullopt;
    }
    if (static_cast<size_t>(substituteNameOffset) + static_cast<size_t>(substituteNameLength) > pathBufferBytes)
    {
        return std::nullopt;
    }

    const size_t charOffset = substituteNameOffset / sizeof(wchar_t);
    const size_t charLength = substituteNameLength / sizeof(wchar_t);
    if (charLength == 0)
    {
        return std::nullopt;
    }

    return std::wstring(pathBuffer + charOffset, charLength);
}

[[nodiscard]] std::wstring NormalizeReparseSubstituteNameForProperties(std::wstring_view substituteName)
{
    constexpr std::wstring_view kNtDosPrefix = L"\\??\\";
    if (OrdinalString::StartsWithNoCase(substituteName, kNtDosPrefix))
    {
        const std::wstring_view tail = substituteName.substr(kNtDosPrefix.size());
        if (OrdinalString::StartsWithNoCase(tail, L"UNC\\"))
        {
            return std::wstring(L"\\\\") + std::wstring(tail.substr(4));
        }
        if (OrdinalString::StartsWithNoCase(tail, L"Volume{"))
        {
            return std::wstring(L"\\\\?\\") + std::wstring(tail);
        }
        return std::wstring(tail);
    }

    constexpr std::wstring_view kMupPrefix = L"\\Device\\Mup\\";
    if (OrdinalString::StartsWithNoCase(substituteName, kMupPrefix))
    {
        return std::wstring(L"\\\\") + std::wstring(substituteName.substr(kMupPrefix.size()));
    }

    return std::wstring(substituteName);
}

[[nodiscard]] std::optional<ItemPropertiesLinkTargetInfo> ResolveReparseTargetForProperties(const wchar_t* path) noexcept
{
    constexpr DWORD kReparseBufferSize = 16u * 1024u;
    std::vector<std::byte> reparseBuffer(kReparseBufferSize);

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile reparseHandle(CreateFileW(path,
                                                FILE_READ_ATTRIBUTES,
                                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                nullptr,
                                                OPEN_EXISTING,
                                                FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                nullptr));
#pragma warning(pop)
    if (! reparseHandle)
    {
        return std::nullopt;
    }

    DWORD bytesReturned = 0;
    if (DeviceIoControl(reparseHandle.get(),
                        FSCTL_GET_REPARSE_POINT,
                        nullptr,
                        0,
                        reparseBuffer.data(),
                        static_cast<DWORD>(reparseBuffer.size()),
                        &bytesReturned,
                        nullptr) == FALSE)
    {
        return std::nullopt;
    }

    if (bytesReturned < offsetof(MountPointReparseDataBufferForProperties, PathBuffer))
    {
        return std::nullopt;
    }

    const auto* header = reinterpret_cast<const MountPointReparseDataBufferForProperties*>(reparseBuffer.data());
    std::optional<std::wstring> substituteName;
    ItemPropertiesLinkTargetInfo info{.sectionTitle = "Reparse Point"};

    if (header->ReparseTag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        const auto* mountPoint       = reinterpret_cast<const MountPointReparseDataBufferForProperties*>(reparseBuffer.data());
        const size_t pathBufferBytes = mountPoint->ReparseDataLength >= (4u * sizeof(USHORT)) ? mountPoint->ReparseDataLength - (4u * sizeof(USHORT)) : 0u;
        substituteName               = ExtractReparsePathBufferStringForProperties(
            mountPoint->PathBuffer, mountPoint->SubstituteNameOffset, mountPoint->SubstituteNameLength, pathBufferBytes);
        info.kind = "Mount point";
    }
    else if (header->ReparseTag == IO_REPARSE_TAG_SYMLINK)
    {
        if (bytesReturned < offsetof(SymbolicLinkReparseDataBufferForProperties, PathBuffer))
        {
            return std::nullopt;
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseDataBufferForProperties*>(reparseBuffer.data());
        const size_t pathBufferBytes =
            symlink->ReparseDataLength >= ((4u * sizeof(USHORT)) + sizeof(ULONG)) ? symlink->ReparseDataLength - ((4u * sizeof(USHORT)) + sizeof(ULONG)) : 0u;
        substituteName =
            ExtractReparsePathBufferStringForProperties(symlink->PathBuffer, symlink->SubstituteNameOffset, symlink->SubstituteNameLength, pathBufferBytes);
        info.kind = "Symbolic link";
    }
    else
    {
        return std::nullopt;
    }

    if (! substituteName.has_value() || substituteName.value().empty())
    {
        return std::nullopt;
    }

    info.target = NormalizeReparseSubstituteNameForProperties(substituteName.value());
    if (info.target.empty())
    {
        return std::nullopt;
    }

    return info;
}

[[nodiscard]] std::optional<ItemPropertiesLinkTargetInfo> TryBuildLinkTargetInfoForProperties(const wchar_t* path, DWORD attributes) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return std::nullopt;
    }

    const std::filesystem::path fsPath(path);
    const std::wstring extension = fsPath.extension().wstring();
    if (OrdinalString::EqualsNoCase(extension, L".lnk"))
    {
        Debug::Perf::Scope perf(L"itemprops.link_target_us");
        perf.SetDetail(L"shortcut");

        std::optional<std::wstring> target = ResolveShellLinkTargetForProperties(path);
        if (! target.has_value())
        {
            return std::nullopt;
        }

        return ItemPropertiesLinkTargetInfo{.sectionTitle = "Shortcut", .target = std::move(target.value())};
    }

    if (OrdinalString::EqualsNoCase(extension, L".url"))
    {
        Debug::Perf::Scope perf(L"itemprops.link_target_us");
        perf.SetDetail(L"internet-shortcut");

        std::wstring url;
        if (! TryReadInternetShortcutUrlForProperties(path, url))
        {
            return std::nullopt;
        }

        ItemPropertiesLinkTargetInfo info{.sectionTitle = "Internet Shortcut", .url = url};
        if (OrdinalString::StartsWithNoCase(url, L"file:"))
        {
            if (std::optional<std::wstring> target = ConvertFileUrlToLocalPathForProperties(url); target.has_value())
            {
                info.target = std::move(target.value());
            }
        }
        else if (IsWindowsAbsolutePathTextForProperties(url))
        {
            info.target = url;
        }

        return info;
    }

    if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        Debug::Perf::Scope perf(L"itemprops.link_target_us");
        perf.SetDetail(L"reparse");
        return ResolveReparseTargetForProperties(path);
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring BuildAlternateStreamPath(const wchar_t* path, std::wstring_view streamName)
{
    std::wstring streamPath = FileSystemInternal::ToExtendedPath(path);
    streamPath.push_back(L':');
    streamPath.append(streamName);
    return streamPath;
}

[[nodiscard]] FileSystemReparsePointPolicy ParseReparsePointPolicy(std::string_view policy) noexcept
{
    if (policy == "copyReparse")
    {
        return FileSystemReparsePointPolicy::CopyReparse;
    }
    if (policy == "followTargets")
    {
        return FileSystemReparsePointPolicy::FollowTargets;
    }
    if (policy == "skip")
    {
        return FileSystemReparsePointPolicy::Skip;
    }

    return FileSystemReparsePointPolicy::CopyReparse;
}

[[nodiscard]] const char* ReparsePointPolicyToString(FileSystemReparsePointPolicy policy) noexcept
{
    switch (policy)
    {
        case FileSystemReparsePointPolicy::CopyReparse: return "copyReparse";
        case FileSystemReparsePointPolicy::FollowTargets: return "followTargets";
        case FileSystemReparsePointPolicy::Skip: return "skip";
    }

    return "copyReparse";
}

[[nodiscard]] FileSystemSearchBackendPreference ParseSearchBackendPreference(std::string_view preference) noexcept
{
    if (preference == "service")
    {
        return FileSystemSearchBackendPreference::Service;
    }
    if (preference == "local-index")
    {
        return FileSystemSearchBackendPreference::LocalIndex;
    }
    if (preference == "scan")
    {
        return FileSystemSearchBackendPreference::Scan;
    }

    return FileSystemSearchBackendPreference::Auto;
}

[[nodiscard]] const char* SearchBackendPreferenceToString(FileSystemSearchBackendPreference preference) noexcept
{
    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto: return "auto";
        case FileSystemSearchBackendPreference::Service: return "service";
        case FileSystemSearchBackendPreference::LocalIndex: return "local-index";
        case FileSystemSearchBackendPreference::Scan: return "scan";
    }

    return "auto";
}

[[nodiscard]] FileSystemConcurrencyMode ParseConcurrencyMode(std::string_view mode) noexcept
{
    if (mode == "manual")
    {
        return FileSystemConcurrencyMode::Manual;
    }

    return FileSystemConcurrencyMode::Auto;
}

[[nodiscard]] const char* ConcurrencyModeToString(FileSystemConcurrencyMode mode) noexcept
{
    switch (mode)
    {
        case FileSystemConcurrencyMode::Auto: return "auto";
        case FileSystemConcurrencyMode::Manual: return "manual";
    }

    return "auto";
}

[[nodiscard]] std::wstring_view StripExtendedPrefix(std::wstring_view path) noexcept
{
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return path.substr(6);
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return path.substr(4);
    }
    return path;
}

[[nodiscard]] bool IsUncPath(std::wstring_view path) noexcept
{
    const std::wstring_view normalized = StripExtendedPrefix(path);
    return normalized.size() >= 2 && normalized[0] == L'\\' && normalized[1] == L'\\';
}

[[nodiscard]] std::wstring ExtractDriveRoot(std::wstring_view path) noexcept
{
    const std::wstring_view normalized = StripExtendedPrefix(path);
    if (normalized.size() >= 3 && normalized[1] == L':' && (normalized[2] == L'\\' || normalized[2] == L'/'))
    {
        return std::format(L"{}:\\", static_cast<wchar_t>(towupper(normalized[0])));
    }

    if (! IsUncPath(path))
    {
        return {};
    }

    size_t serverEnd = normalized.find_first_of(L"\\/", 2);
    if (serverEnd == std::wstring_view::npos)
    {
        return {};
    }

    size_t shareEnd = normalized.find_first_of(L"\\/", serverEnd + 1);
    if (shareEnd == std::wstring_view::npos)
    {
        return std::wstring(normalized) + L"\\";
    }

    return std::wstring(normalized.substr(0, shareEnd + 1));
}

#ifdef _DEBUG
struct StorageProbeDebugHooks
{
    void* context                                                                                                                              = nullptr;
    BOOL (*getVolumePathName)(void* context, const wchar_t* fileName, wchar_t* volumePathName, DWORD bufferLength) noexcept                    = nullptr;
    BOOL (*getVolumeNameForVolumeMountPoint)(void* context, const wchar_t* volumeMountPoint, wchar_t* volumeName, DWORD bufferLength) noexcept = nullptr;
    DWORD (*queryDosDevice)(void* context, const wchar_t* deviceName, wchar_t* targetPath, DWORD bufferLength) noexcept                        = nullptr;
    wil::unique_hfile (*openVolume)(void* context, const wchar_t* volumePath) noexcept                                                         = nullptr;
    BOOL (*deviceIoControl)(void* context,
                            HANDLE device,
                            DWORD ioControlCode,
                            LPVOID inBuffer,
                            DWORD inBufferBytes,
                            LPVOID outBuffer,
                            DWORD outBufferBytes,
                            LPDWORD bytesReturned,
                            LPOVERLAPPED overlapped) noexcept                                                                                  = nullptr;
};

std::atomic<StorageProbeDebugHooks*> g_storageProbeDebugHooks{nullptr};

class StorageProbeDebugScope final
{
public:
    explicit StorageProbeDebugScope(StorageProbeDebugHooks& hooks) noexcept : _previous(g_storageProbeDebugHooks.exchange(&hooks, std::memory_order_acq_rel))
    {
    }

    StorageProbeDebugScope(const StorageProbeDebugScope&)            = delete;
    StorageProbeDebugScope(StorageProbeDebugScope&&)                 = delete;
    StorageProbeDebugScope& operator=(const StorageProbeDebugScope&) = delete;
    StorageProbeDebugScope& operator=(StorageProbeDebugScope&&)      = delete;

    ~StorageProbeDebugScope() noexcept
    {
        g_storageProbeDebugHooks.store(_previous, std::memory_order_release);
    }

private:
    StorageProbeDebugHooks* _previous = nullptr;
};
#endif

void TrimAfterFirstNull(std::wstring& text) noexcept
{
    const size_t terminator = text.find(L'\0');
    if (terminator != std::wstring::npos)
    {
        text.resize(terminator);
    }
}

[[nodiscard]] BOOL GetVolumePathNameForStorageProbe(const wchar_t* fileName, wchar_t* volumePathName, DWORD bufferLength) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks.load(std::memory_order_acquire);
    if (hooks != nullptr && hooks->getVolumePathName != nullptr)
    {
        return hooks->getVolumePathName(hooks->context, fileName, volumePathName, bufferLength);
    }
#endif

    return GetVolumePathNameW(fileName, volumePathName, bufferLength);
}

[[nodiscard]] BOOL GetVolumeNameForVolumeMountPointForStorageProbe(const wchar_t* volumeMountPoint, wchar_t* volumeName, DWORD bufferLength) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks.load(std::memory_order_acquire);
    if (hooks != nullptr && hooks->getVolumeNameForVolumeMountPoint != nullptr)
    {
        return hooks->getVolumeNameForVolumeMountPoint(hooks->context, volumeMountPoint, volumeName, bufferLength);
    }
#endif

    return GetVolumeNameForVolumeMountPointW(volumeMountPoint, volumeName, bufferLength);
}

[[nodiscard]] DWORD QueryDosDeviceForStorageProbe(const wchar_t* deviceName, wchar_t* targetPath, DWORD bufferLength) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks.load(std::memory_order_acquire);
    if (hooks != nullptr && hooks->queryDosDevice != nullptr)
    {
        return hooks->queryDosDevice(hooks->context, deviceName, targetPath, bufferLength);
    }
#endif

    return QueryDosDeviceW(deviceName, targetPath, bufferLength);
}

[[nodiscard]] wil::unique_hfile OpenStorageProbeVolume(const wchar_t* volumePath) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks.load(std::memory_order_acquire);
    if (hooks != nullptr && hooks->openVolume != nullptr)
    {
        return hooks->openVolume(hooks->context, volumePath);
    }
#endif

    return wil::unique_hfile(CreateFileW(volumePath, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr));
}

[[nodiscard]] BOOL DeviceIoControlForStorageProbe(HANDLE device,
                                                  DWORD ioControlCode,
                                                  LPVOID inBuffer,
                                                  DWORD inBufferBytes,
                                                  LPVOID outBuffer,
                                                  DWORD outBufferBytes,
                                                  LPDWORD bytesReturned,
                                                  LPOVERLAPPED overlapped) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks.load(std::memory_order_acquire);
    if (hooks != nullptr && hooks->deviceIoControl != nullptr)
    {
        return hooks->deviceIoControl(hooks->context, device, ioControlCode, inBuffer, inBufferBytes, outBuffer, outBufferBytes, bytesReturned, overlapped);
    }
#endif

    return DeviceIoControl(device, ioControlCode, inBuffer, inBufferBytes, outBuffer, outBufferBytes, bytesReturned, overlapped);
}

[[nodiscard]] std::wstring ResolveLocalVolumeRootPath(std::wstring_view path) noexcept
{
    if (path.empty() || IsUncPath(path))
    {
        return {};
    }

    std::wstring input(path);
    std::wstring volumeRoot(32768u, L'\0');
    if (! GetVolumePathNameForStorageProbe(input.c_str(), volumeRoot.data(), static_cast<DWORD>(volumeRoot.size())))
    {
        return {};
    }

    TrimAfterFirstNull(volumeRoot);
    return volumeRoot;
}

[[nodiscard]] std::wstring BuildDriveLetterDevicePath(std::wstring_view driveRoot) noexcept
{
    if (driveRoot.size() < 2 || driveRoot[1] != L':')
    {
        return {};
    }

    std::wstring volumePath = L"\\\\.\\";
    volumePath.append(driveRoot, 0, 2);
    return volumePath;
}

[[nodiscard]] std::wstring BuildOpenableVolumeGuidDevicePath(std::wstring_view volumeName) noexcept
{
    std::wstring volumePath(volumeName);
    while (! volumePath.empty() && (volumePath.back() == L'\\' || volumePath.back() == L'/'))
    {
        volumePath.pop_back();
    }
    return volumePath;
}

[[nodiscard]] std::wstring ResolveVolumeGuidDevicePath(std::wstring_view volumeRoot) noexcept
{
    if (volumeRoot.empty())
    {
        return {};
    }

    std::wstring volumeRootText(volumeRoot);
    std::wstring volumeName(32768u, L'\0');
    if (! GetVolumeNameForVolumeMountPointForStorageProbe(volumeRootText.c_str(), volumeName.data(), static_cast<DWORD>(volumeName.size())))
    {
        return {};
    }

    TrimAfterFirstNull(volumeName);
    return BuildOpenableVolumeGuidDevicePath(volumeName);
}

[[nodiscard]] std::wstring ResolveSubstTargetPath(std::wstring_view volumeRoot) noexcept
{
    const std::wstring_view normalized = StripExtendedPrefix(volumeRoot);
    if (normalized.size() < 2 || normalized[1] != L':')
    {
        return {};
    }

    const std::wstring deviceName(normalized.substr(0, 2));
    std::wstring targetPath(32768u, L'\0');
    if (QueryDosDeviceForStorageProbe(deviceName.c_str(), targetPath.data(), static_cast<DWORD>(targetPath.size())) == 0u)
    {
        return {};
    }

    TrimAfterFirstNull(targetPath);
    constexpr std::wstring_view kNtDosPrefix = L"\\??\\";
    if (targetPath.rfind(kNtDosPrefix, 0) != 0)
    {
        return {};
    }

    const std::wstring_view targetTail(targetPath.data() + kNtDosPrefix.size(), targetPath.size() - kNtDosPrefix.size());
    if (OrdinalString::StartsWithNoCase(targetTail, L"UNC\\"))
    {
        return std::wstring(L"\\\\") + std::wstring(targetTail.substr(4));
    }

    if (targetTail.size() >= 3 && targetTail[1] == L':' && (targetTail[2] == L'\\' || targetTail[2] == L'/'))
    {
        return std::wstring(targetTail);
    }

    return {};
}

[[nodiscard]] std::wstring ResolveStorageProbeDevicePath(std::wstring_view volumeRoot, std::wstring_view fallbackDriveRoot) noexcept
{
    if (std::wstring volumePath = ResolveVolumeGuidDevicePath(volumeRoot); ! volumePath.empty())
    {
        return volumePath;
    }

    if (std::wstring substTarget = ResolveSubstTargetPath(volumeRoot); ! substTarget.empty())
    {
        const std::wstring substVolumeRoot = ResolveLocalVolumeRootPath(substTarget);
        if (std::wstring volumePath = ResolveVolumeGuidDevicePath(substVolumeRoot); ! volumePath.empty())
        {
            return volumePath;
        }
    }

    return BuildDriveLetterDevicePath(fallbackDriveRoot);
}

[[nodiscard]] unsigned long ClampPreferredBridgeBufferBytes(uint32_t preferredBytes) noexcept
{
    constexpr uint64_t kMinBytes = 512ull * 1024ull;
    constexpr uint64_t kMaxBytes = 16ull * 1024ull * 1024ull;
    const uint64_t clampedBytes  = (std::clamp)(static_cast<uint64_t>(preferredBytes), kMinBytes, kMaxBytes);
    return static_cast<unsigned long>((std::min)(clampedBytes, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max())));
}

void FillTransferHintsLocal(FileSystemTransferHints& hints, bool highLatency) noexcept
{
    hints.latencyClass              = highLatency ? FILESYSTEM_TRANSFER_LATENCY_LAN : FILESYSTEM_TRANSFER_LATENCY_LOCAL;
    hints.flags                     = FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO;
    hints.preferredBufferBytes      = highLatency ? (8u * 1024u * 1024u) : (2u * 1024u * 1024u);
    hints.preferredProgressPeriodMs = 200u;
    if (highLatency)
    {
        hints.flags |= FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS;
    }
}

// Probes the physical medium behind a resolved volume root (seek penalty + bus type). Requires no admin
// rights: a zero-access volume handle is enough for IOCTL_STORAGE_QUERY_PROPERTY. Outputs stay
// "unknown" when the volume cannot be opened (virtual/plugin namespaces, unresolved aliases).
void ProbeLocalStorageMedium(const std::wstring& volumeRoot,
                             const std::wstring& fallbackDriveRoot,
                             bool& seekPenalty,
                             bool& seekPenaltyKnown,
                             STORAGE_BUS_TYPE& busType,
                             bool& busTypeKnown) noexcept
{
    seekPenalty      = false;
    seekPenaltyKnown = false;
    busType          = BusTypeUnknown;
    busTypeKnown     = false;

    const std::wstring volumePath = ResolveStorageProbeDevicePath(volumeRoot, fallbackDriveRoot);
    if (volumePath.empty())
    {
        return;
    }
    wil::unique_hfile volume = OpenStorageProbeVolume(volumePath.c_str());
    if (! volume)
    {
        return;
    }

    {
        STORAGE_PROPERTY_QUERY query{};
        query.PropertyId = StorageDeviceSeekPenaltyProperty;
        query.QueryType  = PropertyStandardQuery;
        DEVICE_SEEK_PENALTY_DESCRIPTOR descriptor{};
        DWORD bytesReturned = 0;
        if (DeviceIoControlForStorageProbe(
                volume.get(), IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), &descriptor, sizeof(descriptor), &bytesReturned, nullptr) &&
            bytesReturned >= sizeof(descriptor))
        {
            seekPenalty      = descriptor.IncursSeekPenalty != FALSE;
            seekPenaltyKnown = true;
        }
    }

    {
        STORAGE_PROPERTY_QUERY query{};
        query.PropertyId = StorageDeviceProperty;
        query.QueryType  = PropertyStandardQuery;
        STORAGE_DEVICE_DESCRIPTOR descriptor{};
        DWORD bytesReturned = 0;
        if (DeviceIoControlForStorageProbe(
                volume.get(), IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), &descriptor, sizeof(descriptor), &bytesReturned, nullptr) &&
            bytesReturned >= FIELD_OFFSET(STORAGE_DEVICE_DESCRIPTOR, RawDeviceProperties))
        {
            busType      = static_cast<STORAGE_BUS_TYPE>(descriptor.BusType);
            busTypeKnown = true;
        }
    }
}

void FillStorageCharacteristicsLocal(FileSystemStorageCharacteristics& characteristics,
                                     const std::wstring& volumeRoot,
                                     const std::wstring& fallbackDriveRoot,
                                     UINT driveType,
                                     bool highLatency) noexcept
{
    if (highLatency)
    {
        characteristics.storageKind = FILESYSTEM_STORAGE_NETWORK_SHARE;
        characteristics.flags =
            FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO | FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
        characteristics.queueDepthHint               = 8u;
        characteristics.preferredCopyMoveConcurrency = 8u;
        characteristics.preferredDeleteConcurrency   = 8u;
        return;
    }

    bool seekPenalty         = false;
    bool seekPenaltyKnown    = false;
    STORAGE_BUS_TYPE busType = BusTypeUnknown;
    bool busTypeKnown        = false;
    ProbeLocalStorageMedium(volumeRoot, fallbackDriveRoot, seekPenalty, seekPenaltyKnown, busType, busTypeKnown);

    if (seekPenaltyKnown && seekPenalty)
    {
        // Rotational media: parallel fan-out causes seek-thrash; sequential wins.
        characteristics.storageKind                  = FILESYSTEM_STORAGE_HDD;
        characteristics.flags                        = FILESYSTEM_STORAGE_FLAG_ROTATIONAL | FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO;
        characteristics.queueDepthHint               = 2u;
        characteristics.preferredCopyMoveConcurrency = 1u;
        characteristics.preferredDeleteConcurrency   = 2u;
    }
    else if (busTypeKnown && busType == BusTypeNvme)
    {
        characteristics.storageKind                  = FILESYSTEM_STORAGE_NVME;
        characteristics.flags                        = FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
        characteristics.queueDepthHint               = 8u;
        characteristics.preferredCopyMoveConcurrency = 8u;
        characteristics.preferredDeleteConcurrency   = 8u;
    }
    else if (seekPenaltyKnown)
    {
        characteristics.storageKind                  = FILESYSTEM_STORAGE_SSD;
        characteristics.flags                        = FILESYSTEM_STORAGE_FLAG_NONE;
        characteristics.queueDepthHint               = 4u;
        characteristics.preferredCopyMoveConcurrency = 4u;
        characteristics.preferredDeleteConcurrency   = 8u;
    }
    else
    {
        // Probe unavailable: keep the historical conservative defaults.
        characteristics.storageKind                  = FILESYSTEM_STORAGE_UNKNOWN;
        characteristics.flags                        = FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO;
        characteristics.queueDepthHint               = 4u;
        characteristics.preferredCopyMoveConcurrency = 4u;
        characteristics.preferredDeleteConcurrency   = 8u;
    }

    if (driveType == DRIVE_REMOVABLE)
    {
        // Removable media (USB sticks, card readers) thrash under fan-out regardless of medium.
        characteristics.queueDepthHint               = std::min(characteristics.queueDepthHint, 2u);
        characteristics.preferredCopyMoveConcurrency = std::min(characteristics.preferredCopyMoveConcurrency, seekPenalty ? 1u : 2u);
        characteristics.preferredDeleteConcurrency   = std::min(characteristics.preferredDeleteConcurrency, 2u);
        characteristics.flags |= FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO;
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        const std::wstring& emittedRoot = ! volumeRoot.empty() ? volumeRoot : fallbackDriveRoot;
        Debug::Perf::Emit(L"FileOps.Storage.ProbedKind",
                          emittedRoot,
                          characteristics.storageKind,
                          characteristics.preferredCopyMoveConcurrency,
                          (seekPenaltyKnown ? 1u : 0u) | (busTypeKnown ? 2u : 0u),
                          S_OK);
    }
}

#ifdef _DEBUG
constexpr Common::DebugSelfTest::Check DebugCheck{L"FileSystem"};

struct DebugMountedVolumeStorageProbeState
{
    std::wstring volumeRoot = L"R:\\MountedVolume\\";
    std::wstring volumeName = L"\\\\?\\Volume{11111111-2222-3333-4444-555555555555}\\";
    std::wstring openedVolumePath;
    std::wstring lastVolumePathInput;
    std::wstring lastVolumeNameInput;
    unsigned int volumePathCalls = 0;
    unsigned int volumeNameCalls = 0;
};

bool DebugCopyStringToBuffer(std::wstring_view text, wchar_t* buffer, DWORD bufferLength) noexcept
{
    if (buffer == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return false;
    }

    const size_t requiredLength = text.size() + 1u;
    if (bufferLength < requiredLength)
    {
        SetLastError(ERROR_MORE_DATA);
        return false;
    }

    std::copy_n(text.data(), text.size(), buffer);
    buffer[text.size()] = L'\0';
    return true;
}

BOOL DebugMountedVolumeGetVolumePathName(void* context, const wchar_t* fileName, wchar_t* volumePathName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugMountedVolumeStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    ++state->volumePathCalls;
    state->lastVolumePathInput = fileName != nullptr ? fileName : L"";
    return DebugCopyStringToBuffer(state->volumeRoot, volumePathName, bufferLength) ? TRUE : FALSE;
}

BOOL DebugMountedVolumeGetVolumeNameForVolumeMountPoint(void* context, const wchar_t* volumeMountPoint, wchar_t* volumeName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugMountedVolumeStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    ++state->volumeNameCalls;
    state->lastVolumeNameInput = volumeMountPoint != nullptr ? volumeMountPoint : L"";
    return DebugCopyStringToBuffer(state->volumeName, volumeName, bufferLength) ? TRUE : FALSE;
}

DWORD DebugMountedVolumeQueryDosDevice(void*, const wchar_t*, wchar_t*, DWORD) noexcept
{
    SetLastError(ERROR_FILE_NOT_FOUND);
    return 0u;
}

wil::unique_hfile DebugMountedVolumeOpenVolume(void* context, const wchar_t* volumePath) noexcept
{
    auto* state = static_cast<DebugMountedVolumeStorageProbeState*>(context);
    if (state != nullptr)
    {
        state->openedVolumePath = volumePath != nullptr ? volumePath : L"";
    }
    return {};
}

void RunDebugMountedVolumeStorageProbeSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    DebugMountedVolumeStorageProbeState state;
    StorageProbeDebugHooks hooks{
        .context                          = &state,
        .getVolumePathName                = DebugMountedVolumeGetVolumePathName,
        .getVolumeNameForVolumeMountPoint = DebugMountedVolumeGetVolumeNameForVolumeMountPoint,
        .queryDosDevice                   = DebugMountedVolumeQueryDosDevice,
        .openVolume                       = DebugMountedVolumeOpenVolume,
    };
    StorageProbeDebugScope scope(hooks);

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    auto* fileSystem          = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"mounted-volume storage probe selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    const HRESULT hr = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Folder\\file.bin", &characteristics);
    const unsigned int firstVolumeNameCalls = state.volumeNameCalls;
    const HRESULT cachedHr = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Other\\second.bin", &characteristics);
    fileSystem->Release();

    DebugCheck(hr == S_OK, L"mounted-volume storage probe should keep GetStorageCharacteristics non-failing", passed, failed);
    DebugCheck(cachedHr == S_OK, L"cached mounted-volume storage query should remain non-failing", passed, failed);
    DebugCheck(state.volumeNameCalls == firstVolumeNameCalls,
               L"repeated storage queries for one resolved volume should reuse the per-instance physical probe", passed, failed);
    DebugCheck(state.volumePathCalls > 0u, L"mounted-volume storage probe should resolve the real volume root", passed, failed);
    DebugCheck(state.volumeNameCalls > 0u, L"mounted-volume storage probe should resolve a volume GUID path", passed, failed);
    DebugCheck(state.lastVolumeNameInput == state.volumeRoot, L"mounted-volume storage probe should resolve the returned mount root", passed, failed);

    std::wstring expectedOpenPath = state.volumeName;
    while (! expectedOpenPath.empty() && (expectedOpenPath.back() == L'\\' || expectedOpenPath.back() == L'/'))
    {
        expectedOpenPath.pop_back();
    }
    DebugCheck(state.openedVolumePath == expectedOpenPath, L"mounted-volume storage probe should open the resolved volume GUID device", passed, failed);
    DebugCheck(state.openedVolumePath != L"\\\\.\\R:", L"mounted-volume storage probe must not open the lexical drive letter", passed, failed);
}

struct DebugSubstStorageProbeState
{
    std::wstring substRoot         = L"R:\\";
    std::wstring substTargetPath   = L"C:\\RealBacker\\SubstRoot";
    std::wstring backingVolumeRoot = L"C:\\";
    std::wstring volumeName        = L"\\\\?\\Volume{66666666-7777-8888-9999-AAAAAAAAAAAA}\\";
    std::wstring openedVolumePath;
    std::wstring lastVolumeNameInput;
    std::wstring lastQueryDosDeviceInput;
    unsigned int volumePathCalls     = 0;
    unsigned int volumeNameCalls     = 0;
    unsigned int queryDosDeviceCalls = 0;
};

BOOL DebugSubstGetVolumePathName(void* context, const wchar_t* fileName, wchar_t* volumePathName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugSubstStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    ++state->volumePathCalls;
    const std::wstring_view requestedPath = fileName != nullptr ? std::wstring_view(fileName) : std::wstring_view{};
    if (requestedPath == L"R:\\SubstAlias\\file.bin")
    {
        return DebugCopyStringToBuffer(state->substRoot, volumePathName, bufferLength) ? TRUE : FALSE;
    }
    if (requestedPath == state->substTargetPath)
    {
        return DebugCopyStringToBuffer(state->backingVolumeRoot, volumePathName, bufferLength) ? TRUE : FALSE;
    }

    SetLastError(ERROR_PATH_NOT_FOUND);
    return FALSE;
}

BOOL DebugSubstGetVolumeNameForVolumeMountPoint(void* context, const wchar_t* volumeMountPoint, wchar_t* volumeName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugSubstStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    ++state->volumeNameCalls;
    state->lastVolumeNameInput = volumeMountPoint != nullptr ? volumeMountPoint : L"";
    if (state->lastVolumeNameInput == state->substRoot)
    {
        SetLastError(ERROR_FILE_NOT_FOUND);
        return FALSE;
    }
    if (state->lastVolumeNameInput == state->backingVolumeRoot)
    {
        return DebugCopyStringToBuffer(state->volumeName, volumeName, bufferLength) ? TRUE : FALSE;
    }

    SetLastError(ERROR_PATH_NOT_FOUND);
    return FALSE;
}

DWORD DebugSubstQueryDosDevice(void* context, const wchar_t* deviceName, wchar_t* targetPath, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugSubstStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return 0u;
    }

    ++state->queryDosDeviceCalls;
    state->lastQueryDosDeviceInput = deviceName != nullptr ? deviceName : L"";
    if (state->lastQueryDosDeviceInput != L"R:")
    {
        SetLastError(ERROR_FILE_NOT_FOUND);
        return 0u;
    }

    const std::wstring target = L"\\??\\" + state->substTargetPath;
    if (! DebugCopyStringToBuffer(target, targetPath, bufferLength))
    {
        return 0u;
    }
    return static_cast<DWORD>(target.size() + 1u);
}

wil::unique_hfile DebugSubstOpenVolume(void* context, const wchar_t* volumePath) noexcept
{
    auto* state = static_cast<DebugSubstStorageProbeState*>(context);
    if (state != nullptr)
    {
        state->openedVolumePath = volumePath != nullptr ? volumePath : L"";
    }
    return {};
}

void RunDebugSubstStorageProbeSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    DebugSubstStorageProbeState state;
    StorageProbeDebugHooks hooks{
        .context                          = &state,
        .getVolumePathName                = DebugSubstGetVolumePathName,
        .getVolumeNameForVolumeMountPoint = DebugSubstGetVolumeNameForVolumeMountPoint,
        .queryDosDevice                   = DebugSubstQueryDosDevice,
        .openVolume                       = DebugSubstOpenVolume,
    };
    StorageProbeDebugScope scope(hooks);

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    auto* fileSystem          = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"SUBST storage probe selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    const HRESULT hr = fileSystem->GetStorageCharacteristics(L"R:\\SubstAlias\\file.bin", &characteristics);
    fileSystem->Release();

    DebugCheck(hr == S_OK, L"SUBST storage probe should keep GetStorageCharacteristics non-failing", passed, failed);
    DebugCheck(state.queryDosDeviceCalls > 0u, L"SUBST storage probe should query the drive alias target", passed, failed);
    DebugCheck(state.lastQueryDosDeviceInput == L"R:", L"SUBST storage probe should query the lexical drive device name", passed, failed);
    DebugCheck(state.lastVolumeNameInput == state.backingVolumeRoot, L"SUBST storage probe should resolve the backing target volume root", passed, failed);

    std::wstring expectedOpenPath = state.volumeName;
    while (! expectedOpenPath.empty() && (expectedOpenPath.back() == L'\\' || expectedOpenPath.back() == L'/'))
    {
        expectedOpenPath.pop_back();
    }
    DebugCheck(state.openedVolumePath == expectedOpenPath, L"SUBST storage probe should open the backing volume GUID device", passed, failed);
    DebugCheck(state.openedVolumePath != L"\\\\.\\R:", L"SUBST storage probe must not open the lexical alias drive", passed, failed);
}

struct DebugNvmeFallbackStorageProbeState
{
    std::wstring volumeRoot           = L"C:\\";
    std::wstring volumeName           = L"\\\\?\\Volume{BBBBBBBB-CCCC-DDDD-EEEE-FFFFFFFFFFFF}\\";
    unsigned int deviceIoControlCalls = 0;
    unsigned int seekPenaltyCalls     = 0;
    unsigned int devicePropertyCalls  = 0;
};

BOOL DebugNvmeFallbackGetVolumePathName(void* context, const wchar_t*, wchar_t* volumePathName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugNvmeFallbackStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    return DebugCopyStringToBuffer(state->volumeRoot, volumePathName, bufferLength) ? TRUE : FALSE;
}

BOOL DebugNvmeFallbackGetVolumeNameForVolumeMountPoint(void* context, const wchar_t*, wchar_t* volumeName, DWORD bufferLength) noexcept
{
    auto* state = static_cast<DebugNvmeFallbackStorageProbeState*>(context);
    if (state == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    return DebugCopyStringToBuffer(state->volumeName, volumeName, bufferLength) ? TRUE : FALSE;
}

DWORD DebugNvmeFallbackQueryDosDevice(void*, const wchar_t*, wchar_t*, DWORD) noexcept
{
    SetLastError(ERROR_FILE_NOT_FOUND);
    return 0u;
}

wil::unique_hfile DebugNvmeFallbackOpenVolume(void*, const wchar_t*) noexcept
{
    return wil::unique_hfile(CreateEventW(nullptr, TRUE, FALSE, nullptr));
}

BOOL DebugNvmeFallbackDeviceIoControl(
    void* context, HANDLE, DWORD ioControlCode, LPVOID inBuffer, DWORD, LPVOID outBuffer, DWORD outBufferBytes, LPDWORD bytesReturned, LPOVERLAPPED) noexcept
{
    auto* state = static_cast<DebugNvmeFallbackStorageProbeState*>(context);
    if (state == nullptr || ioControlCode != IOCTL_STORAGE_QUERY_PROPERTY || inBuffer == nullptr || outBuffer == nullptr || bytesReturned == nullptr)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    ++state->deviceIoControlCalls;
    const auto* query = static_cast<const STORAGE_PROPERTY_QUERY*>(inBuffer);
    if (query->PropertyId == StorageDeviceSeekPenaltyProperty)
    {
        ++state->seekPenaltyCalls;
        *bytesReturned = 0;
        SetLastError(ERROR_INVALID_FUNCTION);
        return FALSE;
    }

    if (query->PropertyId == StorageDeviceProperty)
    {
        ++state->devicePropertyCalls;
        if (outBufferBytes < sizeof(STORAGE_DEVICE_DESCRIPTOR))
        {
            SetLastError(ERROR_MORE_DATA);
            return FALSE;
        }

        auto* descriptor    = static_cast<STORAGE_DEVICE_DESCRIPTOR*>(outBuffer);
        *descriptor         = {};
        descriptor->Version = sizeof(STORAGE_DEVICE_DESCRIPTOR);
        descriptor->Size    = sizeof(STORAGE_DEVICE_DESCRIPTOR);
        descriptor->BusType = BusTypeNvme;
        *bytesReturned      = sizeof(STORAGE_DEVICE_DESCRIPTOR);
        return TRUE;
    }

    SetLastError(ERROR_INVALID_FUNCTION);
    return FALSE;
}

void RunDebugNvmeFallbackStorageProbeSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    DebugNvmeFallbackStorageProbeState state;
    StorageProbeDebugHooks hooks{
        .context                          = &state,
        .getVolumePathName                = DebugNvmeFallbackGetVolumePathName,
        .getVolumeNameForVolumeMountPoint = DebugNvmeFallbackGetVolumeNameForVolumeMountPoint,
        .queryDosDevice                   = DebugNvmeFallbackQueryDosDevice,
        .openVolume                       = DebugNvmeFallbackOpenVolume,
        .deviceIoControl                  = DebugNvmeFallbackDeviceIoControl,
    };
    StorageProbeDebugScope scope(hooks);

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    auto* fileSystem          = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"NVMe fallback storage probe selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    const HRESULT hr = fileSystem->GetStorageCharacteristics(L"C:\\NvmeFallback\\file.bin", &characteristics);
    fileSystem->Release();

    DebugCheck(hr == S_OK, L"NVMe fallback storage probe should keep GetStorageCharacteristics non-failing", passed, failed);
    DebugCheck(state.seekPenaltyCalls == 1u, L"NVMe fallback storage probe should attempt seek-penalty classification", passed, failed);
    DebugCheck(state.devicePropertyCalls == 1u, L"NVMe fallback storage probe should query bus type after seek-penalty failure", passed, failed);
    DebugCheck(characteristics.storageKind == FILESYSTEM_STORAGE_NVME,
               L"NVMe fallback storage probe should classify NVMe from bus type when seek-penalty is unavailable",
               passed,
               failed);
    DebugCheck(characteristics.preferredCopyMoveConcurrency == 8u, L"NVMe fallback storage probe should keep the deep-queue copy/move budget", passed, failed);
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    RunDebugMountedVolumeStorageProbeSelfTest(*passed, *failed);
    RunDebugSubstStorageProbeSelfTest(*passed, *failed);
    RunDebugNvmeFallbackStorageProbeSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugPathNormalizationSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugReparseCopyErrorMappingSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugDirectorySizeErrorPolicySelfTest(*passed, *failed);
    FileSystemInternal::RunDebugSharedFileOpsSchedulerShutdownSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugSearchServiceFallbackCandidateSelfTest(*passed, *failed);

    return *failed == 0u ? S_OK : E_FAIL;
}
#endif

[[nodiscard]] std::string BuildConfigurationJson(FileSystemConcurrencyMode concurrencyMode,
                                                 unsigned int copyMoveMaxConcurrency,
                                                 unsigned int deleteMaxConcurrency,
                                                 unsigned int deleteRecycleBinMaxConcurrency,
                                                 unsigned int recycleBinBatchSize,
                                                 unsigned long enumerationSoftMaxBufferMiB,
                                                 unsigned long enumerationHardMaxBufferMiB,
                                                 FileSystemReparsePointPolicy reparsePointPolicy,
                                                 FileSystemSearchBackendPreference searchBackendPreference,
                                                 unsigned int searchMaxDirectoryWalkers) noexcept
{
    return std::format("{{\"concurrencyMode\":\"{}\",\"copyMoveMaxConcurrency\":{},\"deleteMaxConcurrency\":{},\"deleteRecycleBinMaxConcurrency\":{},"
                       "\"recycleBinBatchSize\":{},\"enumerationSoftMaxBufferMiB\":{},\"enumerationHardMaxBufferMiB\":{},\"reparsePointPolicy\":\"{}\","
                       "\"searchBackendPreference\":\"{}\",\"searchMaxDirectoryWalkers\":{}}}",
                       ConcurrencyModeToString(concurrencyMode),
                       copyMoveMaxConcurrency,
                       deleteMaxConcurrency,
                       deleteRecycleBinMaxConcurrency,
                       recycleBinBatchSize,
                       enumerationSoftMaxBufferMiB,
                       enumerationHardMaxBufferMiB,
                       ReparsePointPolicyToString(reparsePointPolicy),
                       SearchBackendPreferenceToString(searchBackendPreference),
                       searchMaxDirectoryWalkers);
}

class Win32FileReader final : public IFileReader
{
public:
    Win32FileReader(wil::unique_handle file, uint64_t sizeBytes) noexcept : _file(std::move(file)), _sizeBytes(sizeBytes)
    {
    }

    Win32FileReader(const Win32FileReader&)            = delete;
    Win32FileReader(Win32FileReader&&)                 = delete;
    Win32FileReader& operator=(const Win32FileReader&) = delete;
    Win32FileReader& operator=(Win32FileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
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

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;

        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), distance, &moved, origin) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        *newPosition = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD read = 0;
        if (ReadFile(_file.get(), buffer, bytesToRead, &read, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesRead = static_cast<unsigned long>(read);
        return S_OK;
    }

private:
    ~Win32FileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    uint64_t _sizeBytes = 0;
};

class Win32FileWriter final : public IFileWriter
{
public:
    Win32FileWriter(wil::unique_handle file, std::wstring path) noexcept : _file(std::move(file)), _path(std::move(path))
    {
    }

    Win32FileWriter(wil::unique_handle file, std::wstring tempPath, std::wstring finalPath, bool allowReplaceReadOnly) noexcept
        : _file(std::move(file)),
          _path(std::move(tempPath)),
          _finalPath(std::move(finalPath)),
          _allowReplaceReadOnly(allowReplaceReadOnly),
          _replaceOnCommit(true)
    {
    }

    Win32FileWriter(const Win32FileWriter&)            = delete;
    Win32FileWriter(Win32FileWriter&&)                 = delete;
    Win32FileWriter& operator=(const Win32FileWriter&) = delete;
    Win32FileWriter& operator=(Win32FileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
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

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        if (positionBytes == nullptr)
        {
            return E_POINTER;
        }

        *positionBytes = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        LARGE_INTEGER distance{};
        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), distance, &moved, FILE_CURRENT) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        *positionBytes = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (bytesWritten == nullptr)
        {
            return E_POINTER;
        }

        *bytesWritten = 0;

        if (bytesToWrite == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD wrote = 0;
        if (WriteFile(_file.get(), buffer, bytesToWrite, &wrote, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesWritten = static_cast<unsigned long>(wrote);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (FlushFileBuffers(_file.get()) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (_replaceOnCommit)
        {
            _file.reset();
            return PromoteTempIntoFinalPath();
        }

        _committed = true;
        return S_OK;
    }

private:
    ~Win32FileWriter()
    {
        if (_committed)
        {
            return;
        }

        _file.reset();
        if (! _path.empty())
        {
            static_cast<void>(DeleteFileW(_path.c_str()));
        }
    }

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    std::wstring _path;
    std::wstring _finalPath;
    bool _allowReplaceReadOnly = false;
    bool _replaceOnCommit      = false;
    bool _committed            = false;

    HRESULT PromoteTempIntoFinalPath() noexcept
    {
        FileSystemInternal::StagedPromotionOptions options{};
        options.allowReplaceReadOnly     = _allowReplaceReadOnly;
        options.stripTemporaryAttributes = true;

        const HRESULT hr = FileSystemInternal::PromoteStagedTempIntoFinalPath(_path, _finalPath, options);
        if (FAILED(hr))
        {
            return hr;
        }

        _committed = true;
        return S_OK;
    }
};
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

    if (riid == __uuidof(IFileSystemSearch))
    {
        *ppvObject = static_cast<IFileSystemSearch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemItemStreams))
    {
        *ppvObject = static_cast<IFileSystemItemStreams*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryWatch))
    {
        *ppvObject = static_cast<IFileSystemDirectoryWatch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystem::AddRef() noexcept
{
    return static_cast<ULONG>(_refCount.fetch_add(1, std::memory_order_relaxed) + 1);
}

ULONG STDMETHODCALLTYPE FileSystem::Release() noexcept
{
    const ULONG current = static_cast<ULONG>(_refCount.fetch_sub(1, std::memory_order_acq_rel) - 1);
    if (current == 0)
    {
        delete this;
    }
    return current;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (reader == nullptr)
    {
        return E_POINTER;
    }

    *reader = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring filePath = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(
        filePath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    LARGE_INTEGER fileSize{};
    if (GetFileSizeEx(file.get(), &fileSize) == 0)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    if (fileSize.QuadPart < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    auto* impl = new (std::nothrow) Win32FileReader(std::move(file), static_cast<uint64_t>(fileSize.QuadPart));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *reader = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (writer == nullptr)
    {
        return E_POINTER;
    }

    *writer = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const bool allowOverwrite       = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    const bool allowReplaceReadOnly = (flags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) != 0;
    const std::wstring filePath     = FileSystemInternal::ToExtendedPath(path);

    if (allowReplaceReadOnly && ! allowOverwrite)
    {
        return E_INVALIDARG;
    }

    if (allowOverwrite)
    {
        const DWORD destinationAttributes = ::GetFileAttributesW(filePath.c_str());
        if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
        {
            if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
            if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0u && ! allowReplaceReadOnly)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
        }
        else
        {
            const DWORD attributesError = ::GetLastError();
            if (attributesError != ERROR_FILE_NOT_FOUND && attributesError != ERROR_PATH_NOT_FOUND)
            {
                return HRESULT_FROM_WIN32(attributesError != 0u ? attributesError : ERROR_GEN_FAILURE);
            }
        }

        wil::unique_handle file;
        std::wstring tempPath;
        const size_t separator = filePath.find_last_of(L"\\/");
        if (separator == std::wstring::npos || separator + 1u >= filePath.size())
        {
            return E_INVALIDARG;
        }
        const std::wstring prefix = std::wstring(filePath.substr(separator + 1u)) + L".~rs-write-";
        const Common::Paths::UniqueSiblingFileOptions options{.prefix             = prefix,
                                                               .suffix             = L".tmp",
                                                               .shareMode          = FILE_SHARE_READ,
                                                               .flagsAndAttributes = FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY,
                                                               .maximumAttempts   = 32u};
        const HRESULT tempHr = Common::Paths::CreateUniqueSiblingFile(filePath, options, tempPath, file);
        if (FAILED(tempHr))
        {
            return tempHr;
        }

        auto* impl = new (std::nothrow) Win32FileWriter(std::move(file), tempPath, filePath, allowReplaceReadOnly);
        if (! impl)
        {
            file.reset();
            static_cast<void>(::DeleteFileW(tempPath.c_str()));
            return E_OUTOFMEMORY;
        }

        *writer = impl;
        return S_OK;
    }

    wil::unique_handle file(CreateFileW(filePath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    auto* impl = new (std::nothrow) Win32FileWriter(std::move(file), filePath);
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    info->creationTime   = 0;
    info->lastAccessTime = 0;
    info->lastWriteTime  = 0;
    info->attributes     = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring filePath = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(filePath.c_str(),
                                        FILE_READ_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    FILE_BASIC_INFO basic{};
    if (! GetFileInformationByHandleEx(file.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    info->creationTime   = basic.CreationTime.QuadPart;
    info->lastAccessTime = basic.LastAccessTime.QuadPart;
    info->lastWriteTime  = basic.LastWriteTime.QuadPart;
    info->attributes     = basic.FileAttributes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring filePath = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(filePath.c_str(),
                                        FILE_WRITE_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    FILE_BASIC_INFO basic{};
    basic.CreationTime.QuadPart   = info->creationTime;
    basic.LastAccessTime.QuadPart = info->lastAccessTime;
    basic.LastWriteTime.QuadPart  = info->lastWriteTime;
    basic.FileAttributes          = info->attributes;

    if (! SetFileInformationByHandle(file.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    if (_capabilitiesJson.empty())
    {
        UpdateCapabilitiesJson(); // requires _stateMutex
    }

    *jsonUtf8 = _capabilitiesJson.empty() ? "{}" : _capabilitiesJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetTransferHints(const wchar_t* path,
                                                       [[maybe_unused]] FileSystemOperation operationType,
                                                       [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                                       FileSystemTransferHints* hints) noexcept
{
    if (path == nullptr || path[0] == L'\0' || hints == nullptr)
    {
        return E_INVALIDARG;
    }
    if (hints->sizeBytes < sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass              = FILESYSTEM_TRANSFER_LATENCY_UNKNOWN;
    hints->flags                     = FILESYSTEM_TRANSFER_HINT_NONE;
    hints->preferredBufferBytes      = 0;
    hints->preferredProgressPeriodMs = 0;

    const std::wstring driveRoot      = ExtractDriveRoot(path);
    const std::wstring volumeRoot     = ResolveLocalVolumeRootPath(path);
    const std::wstring& driveTypeRoot = ! volumeRoot.empty() ? volumeRoot : driveRoot;
    const UINT driveType              = ! driveTypeRoot.empty() ? GetDriveTypeW(driveTypeRoot.c_str()) : DRIVE_UNKNOWN;
    FillTransferHintsLocal(*hints, driveType == DRIVE_REMOTE || IsUncPath(path));
    hints->preferredBufferBytes = ClampPreferredBridgeBufferBytes(hints->preferredBufferBytes);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept
{
    if (path == nullptr || path[0] == L'\0' || characteristics == nullptr)
    {
        return E_INVALIDARG;
    }
    if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind                  = FILESYSTEM_STORAGE_UNKNOWN;
    characteristics->flags                        = FILESYSTEM_STORAGE_FLAG_NONE;
    characteristics->queueDepthHint               = 0;
    characteristics->preferredCopyMoveConcurrency = 0;
    characteristics->preferredDeleteConcurrency   = 0;

    const std::wstring driveRoot      = ExtractDriveRoot(path);
    const std::wstring volumeRoot     = ResolveLocalVolumeRootPath(path);
    const std::wstring& driveTypeRoot = ! volumeRoot.empty() ? volumeRoot : driveRoot;
    const UINT driveType              = ! driveTypeRoot.empty() ? GetDriveTypeW(driveTypeRoot.c_str()) : DRIVE_UNKNOWN;
    const bool highLatency            = driveType == DRIVE_REMOTE || IsUncPath(path);
    const std::wstring cacheKey = std::format(L"{}\n{}\n{}\n{}", volumeRoot, driveRoot, driveType, highLatency ? 1u : 0u);
    const unsigned long callerSizeBytes = characteristics->sizeBytes;
    {
        // Hold the per-instance lock through the first physical probe so concurrent operation-start
        // queries cannot all issue the same volume IOCTLs.
        std::scoped_lock lock(_storageCharacteristicsMutex);
        if (const auto cached = _storageCharacteristicsCache.find(cacheKey); cached != _storageCharacteristicsCache.end())
        {
            *characteristics          = cached->second;
            characteristics->sizeBytes = callerSizeBytes;
            Debug::Perf::Emit(L"FileOps.Storage.ProbeCache", L"hit", 0u, 1u, 0u, S_OK);
            return S_OK;
        }

        FillStorageCharacteristicsLocal(*characteristics, volumeRoot, driveRoot, driveType, highLatency);
        FileSystemStorageCharacteristics cached = *characteristics;
        cached.sizeBytes = sizeof(FileSystemStorageCharacteristics);
        _storageCharacteristicsCache.emplace(cacheKey, cached);
    }
    Debug::Perf::Emit(L"FileOps.Storage.ProbeCache", L"miss", 0u, 0u, 1u, S_OK);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (GetFileAttributesExW(path, GetFileExInfoStandard, &data) == 0)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    const bool isDirectory   = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const uint64_t sizeBytes = isDirectory ? 0ull : (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | static_cast<uint64_t>(data.nFileSizeLow);

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);

    yyjson_mut_obj_add_int(doc, root, "version", 1);
    yyjson_mut_obj_add_str(doc, root, "title", "properties");

    yyjson_mut_val* sections = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "sections", sections);

    auto addSection = [&](const char* title) noexcept
    {
        yyjson_mut_val* section = yyjson_mut_obj(doc);
        yyjson_mut_arr_add_val(sections, section);
        yyjson_mut_obj_add_str(doc, section, "title", title);

        yyjson_mut_val* sectionFields = yyjson_mut_arr(doc);
        yyjson_mut_obj_add_val(doc, section, "fields", sectionFields);
        return sectionFields;
    };

    const std::wstring fullPath(path);
    const std::filesystem::path fullPathFs(fullPath);
    const std::wstring name = fullPathFs.filename().wstring();

    auto addField = [&](yyjson_mut_val* sectionFields, const char* key, const std::string& value) noexcept
    {
        if (! sectionFields || value.empty())
        {
            return;
        }

        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(sectionFields, field);
    };

    yyjson_mut_val* generalFields = addSection("General");
    addField(generalFields, "Name", Utf8FromUtf16(name.empty() ? std::wstring_view(fullPath) : std::wstring_view(name)));
    addField(generalFields, "Path", Utf8FromUtf16(fullPath));
    addField(generalFields, "Type", isDirectory ? std::string("Directory") : std::string("File"));
    if (! isDirectory)
    {
        addField(generalFields, "Size", FormatItemPropertiesSize(sizeBytes));
    }

    yyjson_mut_val* timestampFields = addSection("Timestamps");
    addField(timestampFields, "Created", FormatFileTimeLocal(data.ftCreationTime));
    addField(timestampFields, "Modified", FormatFileTimeLocal(data.ftLastWriteTime));
    addField(timestampFields, "Accessed", FormatFileTimeLocal(data.ftLastAccessTime));

    yyjson_mut_val* attributeFields = addSection("Attributes");
    addField(attributeFields, "Raw", std::format("0x{:08X}", data.dwFileAttributes));
    addField(attributeFields, "Flags", FormatFileAttributeFlags(data.dwFileAttributes));

    if (std::optional<ItemPropertiesLinkTargetInfo> targetInfo = TryBuildLinkTargetInfoForProperties(path, data.dwFileAttributes);
        targetInfo.has_value() && targetInfo->sectionTitle != nullptr)
    {
        yyjson_mut_val* targetFields = addSection(targetInfo->sectionTitle);
        if (! targetInfo->kind.empty())
        {
            addField(targetFields, "Kind", targetInfo->kind);
        }
        if (! targetInfo->url.empty())
        {
            addField(targetFields, "URL", Utf8FromUtf16(targetInfo->url));
        }
        if (! targetInfo->target.empty())
        {
            addField(targetFields, "Target", Utf8FromUtf16(targetInfo->target));
        }
    }

    yyjson_mut_val* streams = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "streams", streams);

    std::vector<NamedStreamInfo> namedStreams;
    if (SUCCEEDED(EnumerateNamedStreams(path, namedStreams)))
    {
        for (const NamedStreamInfo& stream : namedStreams)
        {
            const std::string streamNameUtf8 = Utf8FromUtf16(stream.name);
            if (streamNameUtf8.empty())
            {
                continue;
            }

            yyjson_mut_val* streamObj = yyjson_mut_obj(doc);
            yyjson_mut_obj_add_strncpy(doc, streamObj, "name", streamNameUtf8.data(), streamNameUtf8.size());
            yyjson_mut_obj_add_uint(doc, streamObj, "sizeBytes", stream.sizeBytes);

            const std::string displaySize = FormatItemPropertiesSize(stream.sizeBytes);
            yyjson_mut_obj_add_strncpy(doc, streamObj, "displaySize", displaySize.data(), displaySize.size());
            yyjson_mut_obj_add_bool(doc, streamObj, "canRemove", true);
            yyjson_mut_arr_add_val(streams, streamObj);
        }
    }

    const char* written = yyjson_mut_write(doc, YYJSON_WRITE_NOFLAG, nullptr);
    if (! written)
    {
        return E_OUTOFMEMORY;
    }
    auto freeWritten = wil::scope_exit([&] { free(const_cast<char*>(written)); });

    {
        std::scoped_lock lock(_propertiesMutex);
        _lastPropertiesJson.assign(written);
        *jsonUtf8 = _lastPropertiesJson.c_str();
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::DeleteItemStream(const wchar_t* path, const wchar_t* streamName) noexcept
{
    if (path == nullptr || path[0] == L'\0' || streamName == nullptr || streamName[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (! IsSafeLogicalStreamName(streamName))
    {
        return E_INVALIDARG;
    }

    const std::wstring streamPath = BuildAlternateStreamPath(path, streamName);
    if (::DeleteFileW(streamPath.c_str()) == 0)
    {
        const DWORD lastError = ::GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = StaticConfigurationSchema();
    return S_OK;
}

const char* GetFileSystemStaticConfigurationSchema() noexcept
{
    return FileSystem::StaticConfigurationSchema();
}

const char* FileSystem::StaticConfigurationSchema() noexcept
{
    return kSchemaJson;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    FileSystemConcurrencyMode concurrencyMode                 = kDefaultConcurrencyMode;
    unsigned int copyMoveMaxConcurrency                       = kDefaultCopyMoveMaxConcurrency;
    unsigned int deleteMaxConcurrency                         = kDefaultDeleteMaxConcurrency;
    unsigned int deleteRecycleBinMaxConcurrency               = kDefaultDeleteRecycleBinMaxConcurrency;
    unsigned int recycleBinBatchSize                          = kDefaultRecycleBinBatchSize;
    unsigned long enumerationSoftMaxBufferMiB                 = kDefaultEnumerationSoftMaxBufferMiB;
    unsigned long enumerationHardMaxBufferMiB                 = kDefaultEnumerationHardMaxBufferMiB;
    FileSystemReparsePointPolicy reparsePointPolicy           = kDefaultReparsePointPolicy;
    FileSystemSearchBackendPreference searchBackendPreference = kDefaultSearchBackendPreference;
    unsigned int searchMaxDirectoryWalkers                    = kDefaultSearchMaxDirectoryWalkers;
#ifdef _DEBUG
    unsigned int directorySizeDelayMs = 0u;
#endif

    constexpr unsigned long kMiB     = 1024u * 1024u;
    const unsigned long maxBufferMiB = (std::numeric_limits<unsigned long>::max)() / kMiB;

    std::string sourceConfiguration = "{}";
    Common::Json::ObjectDocument parsed;
    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        sourceConfiguration = configurationJsonUtf8;
        parsed = Common::Json::ParseObjectDocument(sourceConfiguration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root = parsed.root;
                yyjson_val* concurrencyModeVal = yyjson_obj_get(root, "concurrencyMode");
                if (concurrencyModeVal && yyjson_is_str(concurrencyModeVal))
                {
                    const char* valueText = yyjson_get_str(concurrencyModeVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        concurrencyMode = ParseConcurrencyMode(valueText);
                    }
                }

                yyjson_val* copyMoveVal = yyjson_obj_get(root, "copyMoveMaxConcurrency");
                if (copyMoveVal && yyjson_is_int(copyMoveVal))
                {
                    const int64_t value = yyjson_get_int(copyMoveVal);
                    if (value >= 1)
                    {
                        copyMoveMaxConcurrency = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxCopyMoveMaxConcurrency)));
                    }
                }

                yyjson_val* deleteVal = yyjson_obj_get(root, "deleteMaxConcurrency");
                if (deleteVal && yyjson_is_int(deleteVal))
                {
                    const int64_t value = yyjson_get_int(deleteVal);
                    if (value >= 1)
                    {
                        deleteMaxConcurrency = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxDeleteMaxConcurrency)));
                    }
                }

                yyjson_val* deleteRecycleVal = yyjson_obj_get(root, "deleteRecycleBinMaxConcurrency");
                if (deleteRecycleVal && yyjson_is_int(deleteRecycleVal))
                {
                    const int64_t value = yyjson_get_int(deleteRecycleVal);
                    if (value >= 1)
                    {
                        deleteRecycleBinMaxConcurrency =
                            static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxDeleteRecycleBinMaxConcurrency)));
                    }
                }

                yyjson_val* recycleBinBatchSizeVal = yyjson_obj_get(root, "recycleBinBatchSize");
                if (recycleBinBatchSizeVal && yyjson_is_int(recycleBinBatchSizeVal))
                {
                    const int64_t value = yyjson_get_int(recycleBinBatchSizeVal);
                    if (value >= 1)
                    {
                        recycleBinBatchSize = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxRecycleBinBatchSize)));
                    }
                }

                yyjson_val* softMaxVal = yyjson_obj_get(root, "enumerationSoftMaxBufferMiB");
                if (softMaxVal && yyjson_is_int(softMaxVal))
                {
                    const int64_t value = yyjson_get_int(softMaxVal);
                    if (value >= 1)
                    {
                        enumerationSoftMaxBufferMiB = static_cast<unsigned long>(std::min<int64_t>(value, static_cast<int64_t>(maxBufferMiB)));
                    }
                }

                yyjson_val* hardMaxVal = yyjson_obj_get(root, "enumerationHardMaxBufferMiB");
                if (hardMaxVal && yyjson_is_int(hardMaxVal))
                {
                    const int64_t value = yyjson_get_int(hardMaxVal);
                    if (value >= 1)
                    {
                        enumerationHardMaxBufferMiB = static_cast<unsigned long>(std::min<int64_t>(value, static_cast<int64_t>(maxBufferMiB)));
                    }
                }

                yyjson_val* reparsePolicyVal = yyjson_obj_get(root, "reparsePointPolicy");
                if (reparsePolicyVal && yyjson_is_str(reparsePolicyVal))
                {
                    const char* valueText = yyjson_get_str(reparsePolicyVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        reparsePointPolicy = ParseReparsePointPolicy(valueText);
                    }
                }

                yyjson_val* searchBackendVal = yyjson_obj_get(root, "searchBackendPreference");
                if (searchBackendVal && yyjson_is_str(searchBackendVal))
                {
                    const char* valueText = yyjson_get_str(searchBackendVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        searchBackendPreference = ParseSearchBackendPreference(valueText);
                    }
                }

                yyjson_val* searchWalkersVal = yyjson_obj_get(root, "searchMaxDirectoryWalkers");
                if (searchWalkersVal && yyjson_is_int(searchWalkersVal))
                {
                    const int64_t value = yyjson_get_int(searchWalkersVal);
                    if (value >= 1)
                    {
                        searchMaxDirectoryWalkers = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxSearchMaxDirectoryWalkers)));
                    }
                }

#ifdef _DEBUG
                yyjson_val* delayVal = yyjson_obj_get(root, "directorySizeDelayMs");
                if (delayVal && yyjson_is_int(delayVal))
                {
                    const int64_t value = yyjson_get_int(delayVal);
                    if (value >= 0)
                    {
                        directorySizeDelayMs = static_cast<unsigned int>(std::min<int64_t>(value, 50));
                    }
                }
#endif
    }

    copyMoveMaxConcurrency         = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    deleteMaxConcurrency           = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    deleteRecycleBinMaxConcurrency = std::clamp(deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    recycleBinBatchSize            = std::clamp(recycleBinBatchSize, 1u, kMaxRecycleBinBatchSize);
    searchMaxDirectoryWalkers      = std::clamp(searchMaxDirectoryWalkers, 1u, kMaxSearchMaxDirectoryWalkers);

    enumerationSoftMaxBufferMiB = std::clamp(enumerationSoftMaxBufferMiB, 1ul, maxBufferMiB);
    enumerationHardMaxBufferMiB = std::clamp(enumerationHardMaxBufferMiB, enumerationSoftMaxBufferMiB, maxBufferMiB);

    std::string newConfigurationJson = parsed ? std::move(sourceConfiguration)
                                              : BuildConfigurationJson(concurrencyMode,
                                                                       copyMoveMaxConcurrency,
                                                                       deleteMaxConcurrency,
                                                                       deleteRecycleBinMaxConcurrency,
                                                                       recycleBinBatchSize,
                                                                       enumerationSoftMaxBufferMiB,
                                                                       enumerationHardMaxBufferMiB,
                                                                       reparsePointPolicy,
                                                                       searchBackendPreference,
                                                                       searchMaxDirectoryWalkers);

    std::lock_guard lock(_stateMutex);

    _concurrencyMode                = concurrencyMode;
    _copyMoveMaxConcurrency         = copyMoveMaxConcurrency;
    _deleteMaxConcurrency           = deleteMaxConcurrency;
    _deleteRecycleBinMaxConcurrency = deleteRecycleBinMaxConcurrency;
    _recycleBinBatchSize            = recycleBinBatchSize;
    _enumerationSoftMaxBufferMiB    = enumerationSoftMaxBufferMiB;
    _enumerationHardMaxBufferMiB    = enumerationHardMaxBufferMiB;
    _reparsePointPolicy             = reparsePointPolicy;
    _searchBackendPreference        = searchBackendPreference;
    _searchMaxDirectoryWalkers      = searchMaxDirectoryWalkers;
#ifdef _DEBUG
    _directorySizeDelayMs = directorySizeDelayMs;
#endif

    _configurationJson = std::move(newConfigurationJson);
    UpdateCapabilitiesJson();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    *configurationJsonUtf8 = _configurationJson.empty() ? "{}" : _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const bool isDefault = _concurrencyMode == kDefaultConcurrencyMode && _copyMoveMaxConcurrency == kDefaultCopyMoveMaxConcurrency &&
                           _deleteMaxConcurrency == kDefaultDeleteMaxConcurrency && _deleteRecycleBinMaxConcurrency == kDefaultDeleteRecycleBinMaxConcurrency &&
                           _recycleBinBatchSize == kDefaultRecycleBinBatchSize && _enumerationSoftMaxBufferMiB == kDefaultEnumerationSoftMaxBufferMiB &&
                           _enumerationHardMaxBufferMiB == kDefaultEnumerationHardMaxBufferMiB && _reparsePointPolicy == kDefaultReparsePointPolicy &&
                           _searchBackendPreference == kDefaultSearchBackendPreference && _searchMaxDirectoryWalkers == kDefaultSearchMaxDirectoryWalkers;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

void FileSystem::UpdateCapabilitiesJson() noexcept
{
    // NOTE: Caller must hold _stateMutex.
    _capabilitiesJson = std::format(
        R"json({{"version":1,"operations":{{"copy":true,"move":true,"delete":true,"rename":true,"properties":true,"read":true,"write":true}},"search":{{"version":1,"name":true,"content":true,"indexed":true,"serviceBacked":true,"supportsRegex":true,"supportsSnippets":true,"preferredBackend":"service"}},"concurrency":{{"copyMoveMax":{},"deleteMax":{},"deleteRecycleBinMax":{}}},"crossFileSystem":{{"export":{{"copy":["*"],"move":["*"]}},"import":{{"copy":["*"],"move":["*"]}}}},"pathIdentity":{{"version":1,"pathTextStableIdentity":true,"componentComparison":"ordinalIgnoreCase","normalization":"none","preferredSeparator":"\\","acceptedSeparators":["\\","/"],"casePreserving":true,"caseOnlyRename":"supported"}}}})json",
        std::clamp(_copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency),
        std::clamp(_deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency),
        std::clamp(_deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency));
}

#include "Blake3Digest.h"
#include "DirectoryRecordCursor.h"
#include "FileOperationTraversalPolicy.h"
#include "FileSystem.Internal.h"
#include "Helpers.h"
#include "PathUtils.h"
#include "SynchronousIoCancelWatch.h"
#include "UriEncoding.h"
#include "YyjsonHelpers.h"

#include <array>
#include <atomic>
#include <cwctype>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <thread>

#include <shlwapi.h>
#include <shobjidl.h>
#include <winioctl.h>
#include <winternl.h>

#include <yyjson.h>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Netapi32.lib")

namespace
{
#if defined(ENABLE_TESTS)
enum class AdmissionProbeSlot : size_t
{
    Schema,
    OwnerTid,
    RouteFacts,
    RouteFactsCopy,
    RouteFactsMove,
    LegacyPathCapabilities,
    PeerAllowed,
    Descriptor,
    BuildRoute,
    RealVolumePath,
    RealVolumeName,
    RootBufferAllocations,
    RootRequestedBytes,
    RootCapacityBytes,
    GuidBufferAllocations,
    GuidRequestedBytes,
    GuidCapacityBytes,
    CacheHits,
    CacheMisses,
    Count
};
constexpr size_t kAdmissionProbeSlots = static_cast<size_t>(AdmissionProbeSlot::Count);
enum class LocalRootProbeSlot : size_t
{
    Schema,
    OwnerTid,
    LocalQueries,
    UncQueries,
    HandleAttempts,
    HandleAccepted,
    Fallbacks,
    LegacyResolverCalls,
    Count
};
constexpr size_t kLocalRootProbeSlots = static_cast<size_t>(LocalRootProbeSlot::Count);
struct AdmissionProbe final
{
    bool enabled = false;
    std::array<uint64_t, kAdmissionProbeSlots> values{};
    std::array<uint64_t, kLocalRootProbeSlots> rootValues{};
};
thread_local AdmissionProbe g_admissionProbe;
void CountAdmissionProbe(AdmissionProbeSlot slot, uint64_t value = 1u) noexcept
{
    if (g_admissionProbe.enabled)
        g_admissionProbe.values[static_cast<size_t>(slot)] += value;
}
void CountLocalRootProbe(LocalRootProbeSlot slot) noexcept
{
    if (g_admissionProbe.enabled)
        ++g_admissionProbe.rootValues[static_cast<size_t>(slot)];
}
#endif
#if defined(ENABLE_TESTS)
constexpr std::wstring_view kStagedCopyPromoteFailPathEnvVar     = L"REDSALAMANDER_FILEOPS_STAGED_COPY_PROMOTE_FAIL_PATH";
constexpr std::wstring_view kStagedCopyPromoteFailFiredEnvVar    = L"REDSALAMANDER_FILEOPS_STAGED_COPY_PROMOTE_FAIL_FIRED";
constexpr std::wstring_view kFinalAttributesFailPathEnvVar       = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_PATH";
constexpr std::wstring_view kFinalAttributesFailFiredEnvVar      = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_FIRED";
constexpr std::wstring_view kMetadataForcePresentMaskEnvVar      = L"REDSALAMANDER_FILEOPS_METADATA_FORCE_PRESENT_MASK";
constexpr std::wstring_view kMetadataFailMaskEnvVar              = L"REDSALAMANDER_FILEOPS_METADATA_FAIL_MASK";
constexpr std::wstring_view kStageIdentityUnsupportedPathEnvVar  = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_PATH";
constexpr std::wstring_view kStageIdentityUnsupportedFiredEnvVar = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_FIRED";
// I35: emulates a redirector without FileIdInfo (the NAS probe of 2026-10-11) for every path at or
// below the configured root. Unlike the one-shot injections it stays set until the test clears it.
constexpr std::wstring_view kBindIdentityUnsupportedRootEnvVar          = L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_UNSUPPORTED_PATH";
constexpr std::wstring_view kBindDirectoryIdentityUnsupportedRootEnvVar = L"REDSALAMANDER_FILEOPS_BIND_DIRECTORY_IDENTITY_UNSUPPORTED_PATH";
// A fixture view can report persistent IDs in a different namespace while retaining
// real handles and physical identity for every native operation.
constexpr std::wstring_view kBindIdentityNamespaceRootEnvVar = L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_NAMESPACE_PATH";
// I35 R7.1: the no-follow open of exactly this path fails with ERROR_NOT_SUPPORTED (a redirector that
// refuses the open itself); it stays set until the test clears it.
constexpr std::wstring_view kBindOpenNotSupportedPathEnvVar   = L"REDSALAMANDER_FILEOPS_BIND_OPEN_NOT_SUPPORTED_PATH";
constexpr std::wstring_view kBindDeleteAccessDeniedPathEnvVar = L"REDSALAMANDER_FILEOPS_BIND_DELETE_ACCESS_DENIED_PATH";
// I35 R7.3: a smaller entry bound for the listing of a folder deleted by name (a number of entries).
constexpr std::wstring_view kDeleteByNameEntryBudgetEnvVar = L"REDSALAMANDER_FILEOPS_BYNAME_ENTRY_BUDGET";
// I35 R7.3: the listing of a folder deleted by name reports no file index (a redirector without id listings).
constexpr std::wstring_view kDeleteByNameListWithoutIndexEnvVar = L"REDSALAMANDER_FILEOPS_BYNAME_LIST_WITHOUT_INDEX";
constexpr std::wstring_view kRenameCommittedFailurePathEnvVar   = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_PATH";
constexpr std::wstring_view kRenameCommittedFailureFiredEnvVar  = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_FIRED";
constexpr std::wstring_view kDeleteCommittedFailurePathEnvVar   = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_FAILURE_PATH";
constexpr std::wstring_view kDeleteCommittedFailureFiredEnvVar  = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_FAILURE_FIRED";
constexpr std::wstring_view kDeleteCommittedRetainedPathEnvVar  = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_RETAINED_PATH";
constexpr std::wstring_view kDeleteCommittedRetainedFiredEnvVar = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_RETAINED_FIRED";

[[nodiscard]] uint32_t ReadSelfTestMetadataMask(std::wstring_view name) noexcept
{
    std::array<wchar_t, 32> value{};
    const DWORD length = GetEnvironmentVariableW(name.data(), value.data(), static_cast<DWORD>(value.size()));
    if (length == 0u || length >= value.size())
    {
        return 0u;
    }
    wchar_t* end               = nullptr;
    const unsigned long parsed = std::wcstoul(value.data(), &end, 0);
    return end != value.data() && end != nullptr && *end == L'\0' ? static_cast<uint32_t>(parsed) : 0u;
}

[[nodiscard]] bool ShouldFailOwnedPublicationForSelfTest(const wchar_t* destinationPath,
                                                         std::wstring_view pathEnvironmentVariable,
                                                         std::wstring_view firedEnvironmentVariable,
                                                         std::wstring_view counterName) noexcept
{
    const DWORD required = GetEnvironmentVariableW(pathEnvironmentVariable.data(), nullptr, 0u);
    if (required == 0u)
    {
        return false;
    }

    std::wstring configured(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(pathEnvironmentVariable.data(), configured.data(), required);
    if (written == 0u || written >= required)
    {
        return false;
    }
    configured.resize(written);

    const std::wstring expected = FileSystemInternal::ToExtendedPath(configured);
    const std::wstring actual   = FileSystemInternal::ToExtendedPath(destinationPath);
    if (! OrdinalString::EqualsNoCase(expected, actual))
    {
        return false;
    }

    static_cast<void>(SetEnvironmentVariableW(pathEnvironmentVariable.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(firedEnvironmentVariable.data(), L"1"));
    Debug::Perf::EmitCounter(counterName);
    return true;
}

[[nodiscard]] bool ConsumeMatchingPathInjectionForSelfTest(const wchar_t* path,
                                                           std::wstring_view pathEnvironmentVariable,
                                                           std::wstring_view firedEnvironmentVariable,
                                                           std::wstring_view counterName) noexcept
{
    return path != nullptr && path[0] != L'\0' && ShouldFailOwnedPublicationForSelfTest(path, pathEnvironmentVariable, firedEnvironmentVariable, counterName);
}

// The directory-only control models NAS parent directories whose IDs are unavailable while
// regular-file publication can still return exact object authority.
[[nodiscard]] bool IsBindIdentityUnsupportedForSelfTest(const wchar_t* path) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return false;
    }
    for (const auto name : {kBindIdentityUnsupportedRootEnvVar, kBindDirectoryIdentityUnsupportedRootEnvVar})
    {
        const DWORD required = GetEnvironmentVariableW(name.data(), nullptr, 0u);
        if (required == 0u)
            continue;
        std::wstring configured(static_cast<size_t>(required), L'\0');
        const DWORD written = GetEnvironmentVariableW(name.data(), configured.data(), required);
        if (written == 0u || written >= required)
            continue;
        configured.resize(written);
        while (! configured.empty() && (configured.back() == L'\\' || configured.back() == L'/'))
            configured.pop_back();
        const std::wstring root   = FileSystemInternal::ToExtendedPath(configured);
        const std::wstring actual = FileSystemInternal::ToExtendedPath(path);
        if (root.empty() || actual.size() < root.size() || ! OrdinalString::EqualsNoCase(std::wstring_view(actual).substr(0u, root.size()), root) ||
            (actual.size() != root.size() && actual[root.size()] != L'\\'))
            continue;
        if (name == kBindIdentityUnsupportedRootEnvVar)
            return true;
        const DWORD attributes = GetFileAttributesW(actual.c_str());
        if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
            return true;
    }
    return false;
}

// True while the persistent open-failure injection names exactly `path`.
[[nodiscard]] bool IsBindOpenFailureForSelfTest(const wchar_t* path, std::wstring_view variable) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return false;
    }
    const DWORD required = GetEnvironmentVariableW(variable.data(), nullptr, 0u);
    if (required == 0u)
    {
        return false;
    }
    std::wstring configured(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(variable.data(), configured.data(), required);
    if (written == 0u || written >= required)
    {
        return false;
    }
    configured.resize(written);
    return OrdinalString::EqualsNoCase(FileSystemInternal::ToExtendedPath(configured), FileSystemInternal::ToExtendedPath(path));
}

[[nodiscard]] bool IsBindOpenNotSupportedForSelfTest(const wchar_t* path) noexcept
{
    return IsBindOpenFailureForSelfTest(path, kBindOpenNotSupportedPathEnvVar);
}
#endif

using NtCreateFile_t = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES, PIO_STATUS_BLOCK, PLARGE_INTEGER, ULONG, ULONG, ULONG, ULONG, PVOID, ULONG);
using RtlNtStatusToDosError_t = ULONG(NTAPI*)(NTSTATUS);

[[nodiscard]] NtCreateFile_t GetNtCreateFile() noexcept
{
    static const NtCreateFile_t function = []() noexcept -> NtCreateFile_t
    {
        const HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (ntdll == nullptr)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: GetProcAddress is the supported dynamic ntdll binding boundary.
        return reinterpret_cast<NtCreateFile_t>(::GetProcAddress(ntdll, "NtCreateFile"));
#pragma warning(pop)
    }();
    return function;
}

[[nodiscard]] RtlNtStatusToDosError_t GetRtlNtStatusToDosError() noexcept
{
    static const RtlNtStatusToDosError_t function = []() noexcept -> RtlNtStatusToDosError_t
    {
        const HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (ntdll == nullptr)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: GetProcAddress is the supported dynamic ntdll binding boundary.
        return reinterpret_cast<RtlNtStatusToDosError_t>(::GetProcAddress(ntdll, "RtlNtStatusToDosError"));
#pragma warning(pop)
    }();
    return function;
}

[[nodiscard]] HRESULT CreateExclusiveDirectoryHandle(std::wstring_view stagePath, wil::unique_handle& directory) noexcept
{
    directory.reset();
    const NtCreateFile_t ntCreateFile              = GetNtCreateFile();
    const RtlNtStatusToDosError_t statusToDosError = GetRtlNtStatusToDosError();
    if (ntCreateFile == nullptr || statusToDosError == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::wstring extended = FileSystemInternal::ToExtendedPath(std::wstring(stagePath));
    std::wstring ntPath;
    constexpr std::wstring_view extendedUncPrefix = L"\\\\?\\UNC\\";
    constexpr std::wstring_view extendedPrefix    = L"\\\\?\\";
    if (extended.starts_with(extendedUncPrefix))
    {
        ntPath.assign(L"\\??\\UNC\\");
        ntPath.append(extended.substr(extendedUncPrefix.size()));
    }
    else if (extended.starts_with(extendedPrefix))
    {
        ntPath.assign(L"\\??\\");
        ntPath.append(extended.substr(extendedPrefix.size()));
    }
    else
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    const size_t pathBytes = ntPath.size() * sizeof(wchar_t);
    if (pathBytes > static_cast<size_t>((std::numeric_limits<USHORT>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    UNICODE_STRING name{};
    name.Length        = static_cast<USHORT>(pathBytes);
    name.MaximumLength = name.Length;
    name.Buffer        = ntPath.data();
    OBJECT_ATTRIBUTES attributes{};
    InitializeObjectAttributes(&attributes, &name, OBJ_CASE_INSENSITIVE, nullptr, nullptr);

    IO_STATUS_BLOCK ioStatus{};
    HANDLE rawHandle      = INVALID_HANDLE_VALUE;
    const NTSTATUS status = ntCreateFile(&rawHandle,
                                         DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES | SYNCHRONIZE,
                                         &attributes,
                                         &ioStatus,
                                         nullptr,
                                         // TEMPORARY is a file-data caching hint and is rejected for directory
                                         // creation by some file systems. The exclusive stage contract comes from
                                         // FILE_CREATE, not from a directory attribute.
                                         FILE_ATTRIBUTE_NORMAL,
                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                         FILE_CREATE,
                                         // FILE_CREATE already proves the final component was absent and returns
                                         // the newly created object in the same kernel operation. OPEN_REPARSE_POINT
                                         // is an open-existing option and is rejected by some file systems here.
                                         FILE_DIRECTORY_FILE | FILE_SYNCHRONOUS_IO_NONALERT,
                                         nullptr,
                                         0u);
    if (status < 0)
    {
        const DWORD error = statusToDosError(status);
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    if (rawHandle == nullptr || rawHandle == INVALID_HANDLE_VALUE)
    {
        return E_UNEXPECTED;
    }
    directory.reset(rawHandle);
    return S_OK;
}

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

[[nodiscard]] bool TryParseReparsePointPolicy(std::string_view policy, FileSystemReparsePointPolicy& parsed) noexcept
{
    if (policy == "skip" || policy == "followTargets")
    {
        parsed = FileSystemReparsePointPolicy::Skip;
        return true;
    }
    if (policy == "preserve" || policy == "copyReparse")
    {
        // `copyReparse` is the pre-P3.1 persisted spelling for Preserve.
        parsed = FileSystemReparsePointPolicy::Preserve;
        return true;
    }
    return false;
}

[[nodiscard]] const char* ReparsePointPolicyToString(FileSystemReparsePointPolicy policy) noexcept
{
    switch (policy)
    {
        case FileSystemReparsePointPolicy::Preserve: return "preserve";
        case FileSystemReparsePointPolicy::Skip: return "skip";
    }

    return "preserve";
}

[[nodiscard]] HRESULT CanonicalizeReparsePointPolicyConfiguration(const Common::Json::ObjectDocument& source,
                                                                  FileSystemReparsePointPolicy policy,
                                                                  std::string& configurationJson) noexcept
{
    if (! source)
    {
        return E_INVALIDARG;
    }

    Common::Json::UniqueMutableDocument document(yyjson_mut_doc_new(nullptr));
    if (! document)
    {
        return E_OUTOFMEMORY;
    }

    yyjson_mut_val* root = yyjson_val_mut_copy(document.get(), source.root);
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(document.get(), root);

    static_cast<void>(yyjson_mut_obj_remove_key(root, "reparsePointPolicy"));
    if (! yyjson_mut_obj_add_strcpy(document.get(), root, "reparsePointPolicy", ReparsePointPolicyToString(policy)))
    {
        return E_OUTOFMEMORY;
    }

    size_t length = 0;
    yyjson_write_err error{};
    Common::Json::UniqueMallocString text(yyjson_mut_write_opts(document.get(), YYJSON_WRITE_NOFLAG, nullptr, &length, &error));
    if (! text)
    {
        return E_OUTOFMEMORY;
    }

    configurationJson.assign(text.get(), length);
    return S_OK;
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

// `\\?\C:\x` -> `C:\x`; every other path is returned unchanged. An extended UNC path's plain form
// needs a `\\` that a view cannot supply (ExtractDriveRoot builds it).
[[nodiscard]] std::wstring_view StripExtendedDrivePrefix(std::wstring_view path) noexcept
{
    return Common::Paths::ClassifyWindowsPath(path) == Common::Paths::WindowsPathClass::ExtendedDriveAbsolute ? path.substr(4u) : path;
}

// A plain (`\\server\share`) or extended (`\\?\UNC\server\share`) UNC path; never a device path (`\\.\`).
[[nodiscard]] bool IsUncPath(std::wstring_view path) noexcept
{
    return Common::Paths::IsUncWindowsPath(path);
}

// The root GetDriveTypeW and the storage cache key name a path by: `C:\` for a drive path and the plain
// `\\server\share\` for a UNC path, either in plain or extended form; empty for any other path.
[[nodiscard]] std::wstring ExtractDriveRoot(std::wstring_view path) noexcept
{
    const std::wstring_view drivePath = StripExtendedDrivePrefix(path);
    if (drivePath.size() >= 3 && drivePath[1] == L':' && (drivePath[2] == L'\\' || drivePath[2] == L'/'))
    {
        return std::format(L"{}:\\", static_cast<wchar_t>(towupper(drivePath[0])));
    }

    const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(path);
    if (pathClass != Common::Paths::WindowsPathClass::Unc && pathClass != Common::Paths::WindowsPathClass::ExtendedUnc)
    {
        return {};
    }

    // The server starts after `\\` or `\\?\UNC\`; an extended root is given in its plain form.
    const size_t serverStart = pathClass == Common::Paths::WindowsPathClass::Unc ? 2u : 8u;
    const size_t serverEnd   = path.find_first_of(L"\\/", serverStart);
    if (serverEnd == std::wstring_view::npos)
    {
        return {};
    }

    std::wstring root     = pathClass == Common::Paths::WindowsPathClass::Unc ? std::wstring(path.substr(0u, 2u)) : std::wstring(L"\\\\");
    const size_t shareEnd = path.find_first_of(L"\\/", serverEnd + 1);
    if (shareEnd == std::wstring_view::npos)
    {
        root.append(path.substr(serverStart));
        root.push_back(L'\\');
        return root;
    }

    root.append(path.substr(serverStart, shareEnd + 1 - serverStart));
    return root;
}

#ifdef _DEBUG
struct LocalVolumeRouteCache;

struct StorageProbeDebugHooks
{
    LocalVolumeRouteCache* routeCache                                                                                                          = nullptr;
    uint64_t encodeCalls                                                                                                                       = 0u;
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
    UINT (*getDriveType)(void* context, const wchar_t* rootPath) noexcept                                                                      = nullptr;
};

// A synthetic probe is private to the calling selftest thread. Other live provider threads
// must never observe its stack context or publish its fake identities into the real route cache.
thread_local StorageProbeDebugHooks* g_storageProbeDebugHooks = nullptr;

class StorageProbeDebugScope final
{
public:
    StorageProbeDebugScope(StorageProbeDebugHooks& hooks, LocalVolumeRouteCache& routeCache) noexcept
        : _hooks(hooks),
          _previous(std::exchange(g_storageProbeDebugHooks, &hooks)),
          _previousRouteCache(hooks.routeCache)
    {
        _hooks.routeCache = &routeCache;
    }

    StorageProbeDebugScope(const StorageProbeDebugScope&)            = delete;
    StorageProbeDebugScope(StorageProbeDebugScope&&)                 = delete;
    StorageProbeDebugScope& operator=(const StorageProbeDebugScope&) = delete;
    StorageProbeDebugScope& operator=(StorageProbeDebugScope&&)      = delete;

    ~StorageProbeDebugScope() noexcept
    {
        _hooks.routeCache        = _previousRouteCache;
        g_storageProbeDebugHooks = _previous;
    }

private:
    StorageProbeDebugHooks& _hooks;
    StorageProbeDebugHooks* _previous          = nullptr;
    LocalVolumeRouteCache* _previousRouteCache = nullptr;
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
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
    if (hooks != nullptr && hooks->getVolumePathName != nullptr)
    {
        return hooks->getVolumePathName(hooks->context, fileName, volumePathName, bufferLength);
    }
#endif

#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::RealVolumePath);
#endif
    return GetVolumePathNameW(fileName, volumePathName, bufferLength);
}

[[nodiscard]] BOOL GetVolumeNameForVolumeMountPointForStorageProbe(const wchar_t* volumeMountPoint, wchar_t* volumeName, DWORD bufferLength) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
    if (hooks != nullptr && hooks->getVolumeNameForVolumeMountPoint != nullptr)
    {
        return hooks->getVolumeNameForVolumeMountPoint(hooks->context, volumeMountPoint, volumeName, bufferLength);
    }
#endif

#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::RealVolumeName);
#endif
    return GetVolumeNameForVolumeMountPointW(volumeMountPoint, volumeName, bufferLength);
}

[[nodiscard]] DWORD QueryDosDeviceForStorageProbe(const wchar_t* deviceName, wchar_t* targetPath, DWORD bufferLength) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
    if (hooks != nullptr && hooks->queryDosDevice != nullptr)
    {
        return hooks->queryDosDevice(hooks->context, deviceName, targetPath, bufferLength);
    }
#endif

    return QueryDosDeviceW(deviceName, targetPath, bufferLength);
}

// A UNC root's drive type is a network round trip to its server; the debug hook keeps self-tests off the network.
[[nodiscard]] UINT GetDriveTypeForStorageProbe(const wchar_t* rootPath) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
    if (hooks != nullptr && hooks->getDriveType != nullptr)
    {
        return hooks->getDriveType(hooks->context, rootPath);
    }
#endif

    return GetDriveTypeW(rootPath);
}

[[nodiscard]] wil::unique_hfile OpenStorageProbeVolume(const wchar_t* volumePath) noexcept
{
#ifdef _DEBUG
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
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
    StorageProbeDebugHooks* hooks = g_storageProbeDebugHooks;
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
#if defined(ENABLE_TESTS)
    CountLocalRootProbe(LocalRootProbeSlot::LegacyResolverCalls);
#endif

    std::wstring input(path);
    std::array<wchar_t, MAX_PATH> inlineRoot{};
    if (GetVolumePathNameForStorageProbe(input.c_str(), inlineRoot.data(), static_cast<DWORD>(inlineRoot.size())))
    {
        const auto terminator = std::ranges::find(inlineRoot, L'\0');
        if (terminator != inlineRoot.begin() && terminator != inlineRoot.end() && *(terminator - 1) == L'\\')
        {
            std::wstring volumeRoot(inlineRoot.data(), static_cast<size_t>(terminator - inlineRoot.begin()));
#if defined(ENABLE_TESTS)
            // Count the materialized root if it exceeds this implementation's
            // inline string capacity; a mounted root may still need a small heap.
            if (volumeRoot.capacity() > std::wstring{}.capacity())
            {
                CountAdmissionProbe(AdmissionProbeSlot::RootBufferAllocations);
                CountAdmissionProbe(AdmissionProbeSlot::RootRequestedBytes, static_cast<uint64_t>(volumeRoot.size()) * sizeof(wchar_t));
                CountAdmissionProbe(AdmissionProbeSlot::RootCapacityBytes, static_cast<uint64_t>(volumeRoot.capacity() + 1u) * sizeof(wchar_t));
            }
#endif
            return volumeRoot;
        }
    }
    // The API does not promise a specific too-small error, and a one-character
    // short buffer may succeed without the trailing slash. Retry the same real
    // path once at the previous bound; never manufacture a root from path text.
    std::wstring volumeRoot(32768u, L'\0');
#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::RootBufferAllocations);
    CountAdmissionProbe(AdmissionProbeSlot::RootRequestedBytes, 32768u * sizeof(wchar_t));
    CountAdmissionProbe(AdmissionProbeSlot::RootCapacityBytes, static_cast<uint64_t>(volumeRoot.capacity() + 1u) * sizeof(wchar_t));
#endif
    if (! GetVolumePathNameForStorageProbe(input.c_str(), volumeRoot.data(), static_cast<DWORD>(volumeRoot.size())))
    {
        return {};
    }

    const size_t terminator = volumeRoot.find(L'\0');
    if (terminator == std::wstring::npos || terminator == 0u || volumeRoot[terminator - 1u] != L'\\')
    {
        return {};
    }
    volumeRoot.resize(terminator);
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

[[nodiscard]] bool IsStableVolumeGuidDevicePath(std::wstring_view volumePath) noexcept
{
    return OrdinalString::StartsWithNoCase(volumePath, L"\\\\?\\Volume{") && volumePath.ends_with(L'}');
}

[[nodiscard]] std::wstring ResolveVolumeGuidDevicePath(std::wstring_view volumeRoot) noexcept
{
    if (volumeRoot.empty())
    {
        return {};
    }

    std::wstring volumeRootText(volumeRoot);
    // GetVolumeNameForVolumeMountPointW documents 50 characters for its largest
    // volume GUID path. The input mount root can still be an extended long path.
    std::array<wchar_t, 50u> volumeName{};
    if (! GetVolumeNameForVolumeMountPointForStorageProbe(volumeRootText.c_str(), volumeName.data(), static_cast<DWORD>(volumeName.size())))
    {
        return {};
    }
    const auto terminator = std::ranges::find(volumeName, L'\0');
    if (terminator == volumeName.begin() || terminator == volumeName.end())
    {
        return {};
    }
    return BuildOpenableVolumeGuidDevicePath(std::wstring_view(volumeName.data(), static_cast<size_t>(terminator - volumeName.begin())));
}

// This is optimization eligibility, not operation-path validation. Every other
// spelling retains the existing Win32 root resolver and its exact semantics.
[[nodiscard]] bool CanProbeLocalRootWithoutReparse(std::wstring_view path) noexcept
{
    if (path.size() <= 3u || path.size() >= MAX_PATH || Common::Paths::ClassifyWindowsPath(path) != Common::Paths::WindowsPathClass::DriveAbsolute ||
        path.find(L'/') != std::wstring_view::npos || path.back() == L'\\')
        return false;
    for (size_t start = 3u; start < path.size();)
    {
        const size_t separator            = path.find(L'\\', start);
        const size_t end                  = separator == std::wstring_view::npos ? path.size() : separator;
        const std::wstring_view component = path.substr(start, end - start);
        const std::wstring_view stem      = component.substr(0u, component.find(L'.'));
        if (! Common::Paths::IsSingleChildPathComponent(component) || component.back() == L'.' || component.back() == L' ' ||
            (! stem.empty() && stem.back() == L' ') || component.find_first_of(L":*?\"<>|") != std::wstring_view::npos ||
            FileSystemRouteCapabilitiesBase::IsReservedWindowsLeaf(component))
            return false;
        start = end + 1u;
    }
    return true;
}

// Metadata-only opens differ from exclusive stage creation: never create, write,
// follow a reparse point or retain a handle across this one root observation.
// Both directory and leaf request only attributes/synchronization: FILE_TRAVERSE
// can break a directory RH oplock and wait indefinitely for its acknowledgment.
[[nodiscard]] wil::unique_handle OpenLocalRootProbeObject(std::wstring_view name, HANDLE root, bool directory) noexcept
{
    const auto ntCreateFile = GetNtCreateFile();
    if (! ntCreateFile || name.empty() || name.size() > (std::numeric_limits<USHORT>::max)() / sizeof(wchar_t))
        return {};
    UNICODE_STRING text{};
    text.Length        = static_cast<USHORT>(name.size() * sizeof(wchar_t));
    text.MaximumLength = text.Length;
    // NtCreateFile reads the caller-owned name for this synchronous call.
    text.Buffer = const_cast<wchar_t*>(name.data());
    OBJECT_ATTRIBUTES attributes{};
    constexpr ULONG dontReparse = 0x1000u; // Documented OBJECT_ATTRIBUTES::OBJ_DONT_REPARSE.
    InitializeObjectAttributes(&attributes, &text, OBJ_CASE_INSENSITIVE | dontReparse, root, nullptr);
    IO_STATUS_BLOCK io{};
    wil::unique_handle handle;
    const NTSTATUS status = ntCreateFile(handle.put(),
                                         FILE_READ_ATTRIBUTES | SYNCHRONIZE,
                                         &attributes,
                                         &io,
                                         nullptr,
                                         0u,
                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                         FILE_OPEN,
                                         FILE_SYNCHRONOUS_IO_NONALERT | FILE_OPEN_REPARSE_POINT | (directory ? FILE_DIRECTORY_FILE : 0u),
                                         nullptr,
                                         0u);
    return status >= 0 ? std::move(handle) : wil::unique_handle{};
}

[[nodiscard]] bool TryResolveLocalRootWithoutReparse(std::wstring_view path, std::wstring& volumeRoot, std::wstring& volumeGuid) noexcept
{
#ifdef _DEBUG
    // Synthetic storage observations must never perform real opens or borrow a
    // physical identity. Their existing lexical/GUID fault controls remain intact.
    if (g_storageProbeDebugHooks != nullptr)
        return false;
#endif
    if (! CanProbeLocalRootWithoutReparse(path))
        return false;
    const std::array<wchar_t, 3u> drive{{path[0], L':', L'\0'}};
    std::array<wchar_t, MAX_PATH> target{};
    if (QueryDosDeviceForStorageProbe(drive.data(), target.data(), static_cast<DWORD>(target.size())) == 0u)
        return false;
    const auto terminator = std::ranges::find(target, L'\0');
    const std::wstring_view device(target.data(), static_cast<size_t>(terminator - target.begin()));
    constexpr std::wstring_view physicalPrefix = L"\\Device\\HarddiskVolume";
    if (terminator == target.end() || ! device.starts_with(physicalPrefix) || device.size() == physicalPrefix.size() ||
        ! std::ranges::all_of(device.substr(physicalPrefix.size()), [](wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; }))
        return false; // SUBST, mapped shares and unknown mappings keep the existing resolver.
#if defined(ENABLE_TESTS)
    CountLocalRootProbe(LocalRootProbeSlot::HandleAttempts);
#endif
    std::wstring nativeRoot(device);
    nativeRoot.push_back(L'\\');
    const auto root = OpenLocalRootProbeObject(nativeRoot, nullptr, true);
    if (! root)
        return false;
    FILE_ATTRIBUTE_TAG_INFO rootInfo{};
    if (! GetFileInformationByHandleEx(root.get(), FileAttributeTagInfo, &rootInfo, sizeof(rootInfo)) ||
        (rootInfo.FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != FILE_ATTRIBUTE_DIRECTORY)
        return false;
    const auto item = OpenLocalRootProbeObject(path.substr(3u), root.get(), false);
    if (! item)
        return false;
    FILE_ATTRIBUTE_TAG_INFO itemInfo{};
    if (! GetFileInformationByHandleEx(item.get(), FileAttributeTagInfo, &itemInfo, sizeof(itemInfo)) ||
        (itemInfo.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
        return false;
    std::array<wchar_t, MAX_PATH> normalized{};
    const DWORD normalizedLength =
        GetFinalPathNameByHandleW(item.get(), normalized.data(), static_cast<DWORD>(normalized.size()), VOLUME_NAME_DOS | FILE_NAME_NORMALIZED);
    constexpr std::wstring_view extendedPrefix = L"\\\\?\\";
    if (normalizedLength == 0u || normalizedLength >= normalized.size())
        return false;
    const std::wstring_view normalizedPath(normalized.data(), normalizedLength);
    if (! normalizedPath.starts_with(extendedPrefix) || normalizedPath.substr(extendedPrefix.size()) != path)
        return false; // Preserve Win32 normalization, short-name and exact-name distinctions.
    std::array<wchar_t, 50u> handleGuid{};
    const DWORD guidLength =
        GetFinalPathNameByHandleW(root.get(), handleGuid.data(), static_cast<DWORD>(handleGuid.size()), VOLUME_NAME_GUID | FILE_NAME_NORMALIZED);
    if (guidLength == 0u || guidLength >= handleGuid.size())
        return false;
    std::wstring_view handleIdentity(handleGuid.data(), guidLength);
    if (! handleIdentity.ends_with(L'\\'))
        return false;
    handleIdentity.remove_suffix(1u);
    const std::wstring observedRoot(path.substr(0u, 3u));
    std::wstring observedGuid = ResolveVolumeGuidDevicePath(observedRoot);
    if (observedGuid.empty() || ! OrdinalString::EqualsNoCase(observedGuid, handleIdentity))
        return false; // A concurrent drive-map change cannot reuse the opened volume's identity.
    volumeRoot = observedRoot;
    volumeGuid = std::move(observedGuid);
#if defined(ENABLE_TESTS)
    CountLocalRootProbe(LocalRootProbeSlot::HandleAccepted);
#endif
    return true;
}

struct LocalCapabilityRouteInfo final
{
    std::string rootId;
    bool remote = false;
};

[[nodiscard]] std::string EncodeLocalRootId(std::wstring rootIdentity) noexcept
{
#ifdef _DEBUG
    if (g_storageProbeDebugHooks != nullptr)
        ++g_storageProbeDebugHooks->encodeCalls;
#endif
    rootIdentity = OrdinalString::FoldCaseInvariant(rootIdentity);
    std::string encoded;
    if (! Common::Uri::TryPercentEncodeUtf8(rootIdentity, Common::Uri::SlashPolicy::Encode, encoded) || encoded.empty())
    {
        return {};
    }
    return std::format("local-volume:{}", encoded);
}

// Route facts are queried once per child name by Create Directory, Batch Rename, and Change
// Case, often on the UI thread. Resolve the lexical mount root and its current direct GUID for
// every query. A short-lived exact-root/GUID hit may reuse only the encoded route result.
struct LocalVolumeRouteCache final
{
    LocalVolumeRouteCache()                                        = default;
    LocalVolumeRouteCache(const LocalVolumeRouteCache&)            = delete;
    LocalVolumeRouteCache(LocalVolumeRouteCache&&)                 = delete;
    LocalVolumeRouteCache& operator=(const LocalVolumeRouteCache&) = delete;
    LocalVolumeRouteCache& operator=(LocalVolumeRouteCache&&)      = delete;

    std::mutex mutex;
    std::wstring volumeRoot;
    std::wstring volumeGuidDevicePath;
    LocalCapabilityRouteInfo route;
    ULONGLONG resolvedTick = 0u;
};
LocalVolumeRouteCache g_localVolumeRouteCache;
constexpr ULONGLONG kLocalVolumeRouteCacheTtlMs = 2'000u;

[[nodiscard]] LocalVolumeRouteCache& CurrentLocalVolumeRouteCache() noexcept
{
#ifdef _DEBUG
    if (g_storageProbeDebugHooks != nullptr && g_storageProbeDebugHooks->routeCache != nullptr)
        return *g_storageProbeDebugHooks->routeCache;
#endif
    return g_localVolumeRouteCache;
}

[[nodiscard]] LocalCapabilityRouteInfo BuildLocalCapabilityRouteInfo(const wchar_t* pathText) noexcept
{
#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::BuildRoute);
#endif
    const std::wstring_view path(pathText);
    LocalCapabilityRouteInfo route{};
    if (IsUncPath(path))
    {
#if defined(ENABLE_TESTS)
        CountLocalRootProbe(LocalRootProbeSlot::UncQueries);
#endif
        route.remote                    = true;
        const std::wstring rootIdentity = ExtractDriveRoot(path);
        route.rootId                    = rootIdentity.empty() ? std::string("local-provider-root") : EncodeLocalRootId(rootIdentity);
        return route;
    }

    LocalVolumeRouteCache& routeCache = CurrentLocalVolumeRouteCache();
#if defined(ENABLE_TESTS)
    CountLocalRootProbe(LocalRootProbeSlot::LocalQueries);
#endif
    std::wstring volumeRoot;
    std::wstring rootIdentity;
    const bool directRoot = TryResolveLocalRootWithoutReparse(path, volumeRoot, rootIdentity);
    if (! directRoot)
    {
#if defined(ENABLE_TESTS)
        CountLocalRootProbe(LocalRootProbeSlot::Fallbacks);
#endif
        volumeRoot = ResolveLocalVolumeRootPath(path);
    }
    if (volumeRoot.empty())
    {
        // The provider-root query used by legacy capability inventory has no concrete volume.
        // Real operation admission always queries its actual source/destination path.
        route.rootId = "local-provider-root";
        return route;
    }

    route.remote = GetDriveTypeForStorageProbe(volumeRoot.c_str()) == DRIVE_REMOTE;
    if (! directRoot)
        rootIdentity = ResolveVolumeGuidDevicePath(volumeRoot);
    const ULONGLONG nowTick = GetTickCount64();
    if (rootIdentity.empty())
    {
        // Keep the lexical root fallback for this call, but do not let a prior
        // GUID-backed snapshot survive an unavailable identity observation.
        route.rootId = EncodeLocalRootId(volumeRoot);
        std::scoped_lock lock(routeCache.mutex);
        routeCache.volumeRoot.clear();
        routeCache.volumeGuidDevicePath.clear();
        routeCache.resolvedTick = 0u;
        return route;
    }

    {
        std::scoped_lock lock(routeCache.mutex);
        if (OrdinalString::EqualsNoCase(routeCache.volumeRoot, volumeRoot) && OrdinalString::EqualsNoCase(routeCache.volumeGuidDevicePath, rootIdentity) &&
            nowTick - routeCache.resolvedTick <= kLocalVolumeRouteCacheTtlMs)
        {
#if defined(ENABLE_TESTS)
            CountAdmissionProbe(AdmissionProbeSlot::CacheHits);
#endif
            LocalCapabilityRouteInfo cachedRoute = routeCache.route;
            cachedRoute.remote                   = route.remote;
            return cachedRoute;
        }
    }

#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::CacheMisses);
#endif
    route.rootId = EncodeLocalRootId(rootIdentity);
    std::scoped_lock lock(routeCache.mutex);
    routeCache.volumeRoot           = volumeRoot;
    routeCache.volumeGuidDevicePath = rootIdentity;
    routeCache.route                = route;
    routeCache.resolvedTick         = nowTick;
    return route;
}

[[nodiscard]] std::wstring ResolveSubstTargetPath(std::wstring_view volumeRoot) noexcept
{
    const std::wstring_view normalized = StripExtendedDrivePrefix(volumeRoot);
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
void ProbeLocalStorageMedium(const std::wstring& volumePath, bool& seekPenalty, bool& seekPenaltyKnown, STORAGE_BUS_TYPE& busType, bool& busTypeKnown) noexcept
{
    seekPenalty      = false;
    seekPenaltyKnown = false;
    busType          = BusTypeUnknown;
    busTypeKnown     = false;

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
                                     const std::wstring& volumePath,
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
    ProbeLocalStorageMedium(volumePath, seekPenalty, seekPenaltyKnown, busType, busTypeKnown);

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
        Debug::Perf::Emit(L"FileOps.Storage.ProbedKind",
                          volumePath,
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
    std::vector<std::pair<std::wstring, std::wstring>> pathRootMappings;
    std::wstring openedVolumePath;
    std::wstring lastVolumePathInput;
    std::wstring lastVolumeNameInput;
    unsigned int openVolumeCalls      = 0;
    unsigned int volumePathCalls      = 0;
    unsigned int volumeNameCalls      = 0;
    unsigned int queryDosDeviceCalls  = 0;
    DWORD firstVolumePathBufferLength = 0u;
    DWORD lastVolumePathBufferLength  = 0u;
    DWORD lastVolumeNameBufferLength  = 0u;
    DWORD volumePathError             = ERROR_SUCCESS;
    DWORD volumeNameError             = ERROR_SUCCESS;
    UINT driveType                    = DRIVE_FIXED;
    bool oneShortSuccess              = false;
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
    if (state->volumePathCalls == 1u)
        state->firstVolumePathBufferLength = bufferLength;
    state->lastVolumePathBufferLength = bufferLength;
    state->lastVolumePathInput        = fileName != nullptr ? fileName : L"";
    if (state->volumePathError != ERROR_SUCCESS)
    {
        SetLastError(state->volumePathError);
        return FALSE;
    }
    for (const auto& [input, root] : state->pathRootMappings)
    {
        if (state->lastVolumePathInput == input)
        {
            return DebugCopyStringToBuffer(root, volumePathName, bufferLength) ? TRUE : FALSE;
        }
    }
    if (state->oneShortSuccess && state->volumeRoot.size() == bufferLength && state->volumeRoot.ends_with(L'\\'))
        return DebugCopyStringToBuffer(std::wstring_view(state->volumeRoot).substr(0u, state->volumeRoot.size() - 1u), volumePathName, bufferLength) ? TRUE
                                                                                                                                                     : FALSE;
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
    state->lastVolumeNameBufferLength = bufferLength;
    state->lastVolumeNameInput        = volumeMountPoint != nullptr ? volumeMountPoint : L"";
    if (state->volumeNameError != ERROR_SUCCESS)
    {
        SetLastError(state->volumeNameError);
        return FALSE;
    }
    return DebugCopyStringToBuffer(state->volumeName, volumeName, bufferLength) ? TRUE : FALSE;
}

DWORD DebugMountedVolumeQueryDosDevice(void* context, const wchar_t*, wchar_t*, DWORD) noexcept
{
    auto* state = static_cast<DebugMountedVolumeStorageProbeState*>(context);
    if (state != nullptr)
    {
        ++state->queryDosDeviceCalls;
    }
    SetLastError(ERROR_FILE_NOT_FOUND);
    return 0u;
}

wil::unique_hfile DebugMountedVolumeOpenVolume(void* context, const wchar_t* volumePath) noexcept
{
    auto* state = static_cast<DebugMountedVolumeStorageProbeState*>(context);
    if (state != nullptr)
    {
        ++state->openVolumeCalls;
        state->openedVolumePath = volumePath != nullptr ? volumePath : L"";
    }
    return {};
}

UINT DebugRouteDriveType(void* context, const wchar_t*) noexcept
{
    const auto* state = static_cast<const DebugMountedVolumeStorageProbeState*>(context);
    return state != nullptr ? state->driveType : DRIVE_UNKNOWN;
}

void RunDebugLocalRouteObservationSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    for (const auto& [path, expected] : std::to_array<std::pair<std::wstring_view, bool>>({
             {L"C:\\ordinary\\file.bin", true},
             {L"c:\\ordinary\\file.bin", true},
             {L"C:\\東京\\child-01.bin", true},
             {L"C:\\ordinary\\file.bin.", false},
             {L"C:\\ordinary\\file.bin ", false},
             {L"C:\\ordinary\\.\\file.bin", false},
             {L"C:\\ordinary\\..\\file.bin", false},
             {L"C:\\ordinary\\\\file.bin", false},
             {L"C:\\ordinary\\COM1.txt", false},
             {L"C:\\ordinary\\COM1 .txt", false},
             {L"C:\\ordinary\\COM\u00B9.txt", false},
             {L"C:\\ordinary\\CONIN$", false},
             {L"C:\\ordinary\\conin$", false},
             {L"C:\\ordinary\\nul.txt", false},
             {L"C:\\ordinary\\lpt1.txt", false},
             {L"C:\\ordinary\\file.bin:stream", false},
             {L"C:/ordinary/file.bin", false},
             {L"\\\\?\\C:\\ordinary\\file.bin", false},
             {L"\\\\server\\share\\file.bin", false},
             {L"\\\\?\\Volume{12345678-1234-1234-1234-123456789abc}\\file.bin", false},
             {L"C:relative", false},
             {L"C:\\", false},
             {L"C:\\ordinary\\", false},
             {L"", false},
             {L"relative", false},
             {L"C:\\ordinary\\*.bin", false},
         }))
        DebugCheck(CanProbeLocalRootWithoutReparse(path) == expected,
                   L"direct local root eligibility must preserve Win32 namespace and normalization semantics",
                   passed,
                   failed);
    const std::wstring longProbePath = L"C:\\" + std::wstring(MAX_PATH, L'a');
    DebugCheck(! CanProbeLocalRootWithoutReparse(longProbePath), L"long paths retain the existing bounded lexical resolver", passed, failed);
    DebugMountedVolumeStorageProbeState state;
    state.volumeRoot             = L"C:\\";
    const std::wstring pathA     = L"C:\\ordinary\\file.bin";
    const std::wstring pathB     = L"C:\\MountB\\finalLinkToA";
    const std::wstring pathBWarm = L"C:\\MountB\\second.bin";
    const std::wstring pathC     = L"C:\\MountB\\MountC\\child.bin";
    state.pathRootMappings       = {
        {pathA, L"C:\\"},
        {pathB, L"C:\\MountB\\"},
        {pathBWarm, L"C:\\MountB\\"},
        {pathC, L"C:\\MountB\\MountC\\"},
    };
    StorageProbeDebugHooks hooks{
        .context                          = &state,
        .getVolumePathName                = DebugMountedVolumeGetVolumePathName,
        .getVolumeNameForVolumeMountPoint = DebugMountedVolumeGetVolumeNameForVolumeMountPoint,
        .queryDosDevice                   = DebugMountedVolumeQueryDosDevice,
        .getDriveType                     = DebugRouteDriveType,
    };
    LocalVolumeRouteCache testRouteCache;
    StorageProbeDebugScope scope(hooks, testRouteCache);

    const std::wstring guidA     = BuildOpenableVolumeGuidDevicePath(state.volumeName);
    const std::string idA        = EncodeLocalRootId(guidA);
    const uint64_t encodeBeforeA = hooks.encodeCalls;
    const auto routeA            = BuildLocalCapabilityRouteInfo(pathA.c_str());
    const uint64_t encodeAfterA  = hooks.encodeCalls;
    DebugCheck(encodeAfterA - encodeBeforeA == 1u && routeA.rootId == idA && ! routeA.remote,
               L"cold local route should use the lexical mount root's direct volume GUID",
               passed,
               failed);

    bool otherThreadIsolated = false;
    std::jthread otherProbe([&]() noexcept
    {
        if (g_storageProbeDebugHooks != nullptr)
            return;
        DebugMountedVolumeStorageProbeState otherState;
        otherState.volumeRoot = L"T:\\";
        otherState.volumeName = L"\\\\?\\Volume{CCCCCCCC-DDDD-EEEE-FFFF-000000000000}\\";
        StorageProbeDebugHooks otherHooks{
            .context                          = &otherState,
            .getVolumePathName                = DebugMountedVolumeGetVolumePathName,
            .getVolumeNameForVolumeMountPoint = DebugMountedVolumeGetVolumeNameForVolumeMountPoint,
            .queryDosDevice                   = DebugMountedVolumeQueryDosDevice,
            .getDriveType                     = DebugRouteDriveType,
        };
        LocalVolumeRouteCache otherCache;
        StorageProbeDebugScope otherScope(otherHooks, otherCache);
        const auto otherRoute = BuildLocalCapabilityRouteInfo(L"T:\\owned-probe\\file.bin");
        otherThreadIsolated   = otherRoute.rootId == EncodeLocalRootId(BuildOpenableVolumeGuidDevicePath(otherState.volumeName)) &&
                                otherCache.volumeGuidDevicePath == BuildOpenableVolumeGuidDevicePath(otherState.volumeName);
    });
    otherProbe.join();
    DebugCheck(otherThreadIsolated && g_storageProbeDebugHooks == &hooks && testRouteCache.route.rootId == idA && state.volumePathCalls == 1u,
               L"simultaneous selftest threads keep probe hooks and synthetic route caches private",
               passed,
               failed);

    state.volumeRoot             = L"C:\\MountB\\";
    state.volumeName             = L"\\\\?\\Volume{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}\\";
    const std::wstring guidB     = BuildOpenableVolumeGuidDevicePath(state.volumeName);
    const uint64_t encodeBeforeB = hooks.encodeCalls;
    const auto routeB            = BuildLocalCapabilityRouteInfo(pathB.c_str());
    const uint64_t encodeAfterB  = hooks.encodeCalls;
    DebugCheck(encodeAfterB - encodeBeforeB == 1u && routeB.rootId == EncodeLocalRootId(guidB) && routeB.rootId != routeA.rootId &&
                   state.lastVolumePathInput == pathB && state.lastVolumeNameInput == L"C:\\MountB\\",
               L"a path lexically inside MountB must not inherit a prior root reached by following its final link",
               passed,
               failed);

    const unsigned int pathCallsBeforeWarm = state.volumePathCalls;
    const unsigned int guidCallsBeforeWarm = state.volumeNameCalls;
    const uint64_t encodeBeforeWarm        = hooks.encodeCalls;
    const auto routeBWarm                  = BuildLocalCapabilityRouteInfo(pathBWarm.c_str());
    const uint64_t encodeAfterWarm         = hooks.encodeCalls;
    DebugCheck(encodeAfterWarm == encodeBeforeWarm && routeBWarm.rootId == routeB.rootId && routeBWarm.remote == routeB.remote &&
                   state.volumePathCalls == pathCallsBeforeWarm + 1u && state.volumeNameCalls == guidCallsBeforeWarm + 1u,
               L"a warm route hit must still resolve the requested lexical root and current direct GUID",
               passed,
               failed);
    state.driveType                      = DRIVE_REMOTE;
    const uint64_t encodeBeforeRemoteHit = hooks.encodeCalls;
    const auto remoteHit                 = BuildLocalCapabilityRouteInfo(pathBWarm.c_str());
    const uint64_t encodeAfterRemoteHit  = hooks.encodeCalls;
    DebugCheck(encodeAfterRemoteHit == encodeBeforeRemoteHit && remoteHit.rootId == routeB.rootId && remoteHit.remote,
               L"a warm matching GUID uses the freshly observed drive classification",
               passed,
               failed);
    state.driveType = DRIVE_FIXED;
    {
        std::scoped_lock lock(testRouteCache.mutex);
        testRouteCache.resolvedTick = GetTickCount64() - (kLocalVolumeRouteCacheTtlMs + 1u);
    }
    const uint64_t encodeBeforeExpiry = hooks.encodeCalls;
    const auto expiredRoute           = BuildLocalCapabilityRouteInfo(pathBWarm.c_str());
    const uint64_t encodeAfterExpiry  = hooks.encodeCalls;
    DebugCheck(encodeAfterExpiry - encodeBeforeExpiry == 1u && expiredRoute.rootId == routeB.rootId && expiredRoute.remote == routeB.remote,
               L"an expired route encodes the freshly resolved GUID once without changing its identity",
               passed,
               failed);

    // The root is part of the cache key even if two mount roots resolve to the same GUID.
    state.volumeRoot                             = L"C:\\MountB\\MountC\\";
    const uint64_t encodeBeforeSameGuidOtherRoot = hooks.encodeCalls;
    const auto sameGuidOtherRoot                 = BuildLocalCapabilityRouteInfo(pathC.c_str());
    const uint64_t encodeAfterSameGuidOtherRoot  = hooks.encodeCalls;
    DebugCheck(encodeAfterSameGuidOtherRoot - encodeBeforeSameGuidOtherRoot == 1u && sameGuidOtherRoot.rootId == routeB.rootId &&
                   state.lastVolumePathInput == pathC && state.lastVolumeNameInput == state.volumeRoot,
               L"a different lexical mount root with the same GUID must miss and encode its own observation",
               passed,
               failed);

    state.volumeRoot  = L"C:\\MountB\\MountC\\";
    state.volumeName  = L"\\\\?\\Volume{BBBBBBBB-CCCC-DDDD-EEEE-FFFFFFFFFFFF}\\";
    const auto routeC = BuildLocalCapabilityRouteInfo(pathC.c_str());
    DebugCheck(routeC.rootId == EncodeLocalRootId(BuildOpenableVolumeGuidDevicePath(state.volumeName)) && routeC.rootId != routeB.rootId &&
                   state.lastVolumeNameInput == L"C:\\MountB\\MountC\\",
               L"a nested mount must retain its own lexical root and direct volume identity",
               passed,
               failed);

    state.volumeRoot                   = L"C:\\MountB\\";
    state.volumeName                   = L"\\\\?\\Volume{CCCCCCCC-DDDD-EEEE-FFFF-000000000000}\\";
    const uint64_t encodeBeforeRemount = hooks.encodeCalls;
    const auto remounted               = BuildLocalCapabilityRouteInfo(pathB.c_str());
    const uint64_t encodeAfterRemount  = hooks.encodeCalls;
    DebugCheck(encodeAfterRemount - encodeBeforeRemount == 1u && remounted.rootId == EncodeLocalRootId(BuildOpenableVolumeGuidDevicePath(state.volumeName)) &&
                   remounted.rootId != routeB.rootId,
               L"the same lexical mount root must receive a new route identity after its direct GUID changes",
               passed,
               failed);

    state.volumeNameError               = ERROR_NOT_SUPPORTED;
    const uint64_t encodeBeforeFallback = hooks.encodeCalls;
    const auto unavailableGuid          = BuildLocalCapabilityRouteInfo(pathB.c_str());
    const uint64_t encodeAfterFallback  = hooks.encodeCalls;
    bool routeCacheInvalidated          = false;
    {
        std::scoped_lock lock(testRouteCache.mutex);
        routeCacheInvalidated = testRouteCache.volumeRoot.empty();
    }
    DebugCheck(encodeAfterFallback - encodeBeforeFallback == 1u && unavailableGuid.rootId == EncodeLocalRootId(state.volumeRoot) && routeCacheInvalidated,
               L"an unavailable current GUID must use the lexical fallback and invalidate the old GUID snapshot",
               passed,
               failed);
    state.volumeNameError = ERROR_SUCCESS;

    state.volumePathError = ERROR_ACCESS_DENIED;
    const auto unresolved = BuildLocalCapabilityRouteInfo(pathB.c_str());
    DebugCheck(unresolved.rootId == "local-provider-root", L"a failed lexical root resolver must not reuse a prior route cache entry", passed, failed);
    state.volumePathError = ERROR_SUCCESS;

    const unsigned int volumePathCallsBeforeUnc = state.volumePathCalls;
    const uint64_t encodeBeforeUnc              = hooks.encodeCalls;
    const auto unc                              = BuildLocalCapabilityRouteInfo(L"\\\\server\\share\\Folder\\file.bin");
    const uint64_t encodeAfterUnc               = hooks.encodeCalls;
    DebugCheck(encodeAfterUnc - encodeBeforeUnc == 1u && unc.remote && ! unc.rootId.empty() && state.volumePathCalls == volumePathCallsBeforeUnc,
               L"UNC classification must bypass local mount-root resolution",
               passed,
               failed);

    // A malformed synthetic GUID exercises the same strict UTF-16 failure on cold and warm paths.
    state.volumeRoot = L"C:\\MountB\\";
    state.volumeName = L"\\\\?\\Volume{";
    state.volumeName.push_back(static_cast<wchar_t>(0xD800u));
    state.volumeName.append(L"}\\");
    const std::string invalidExpected     = EncodeLocalRootId(BuildOpenableVolumeGuidDevicePath(state.volumeName));
    const uint64_t encodeBeforeInvalid    = hooks.encodeCalls;
    const auto invalidCold                = BuildLocalCapabilityRouteInfo(pathB.c_str());
    const uint64_t encodeAfterInvalidCold = hooks.encodeCalls;
    const auto invalidWarm                = BuildLocalCapabilityRouteInfo(pathBWarm.c_str());
    const uint64_t encodeAfterInvalidWarm = hooks.encodeCalls;
    DebugCheck(invalidExpected.empty() && invalidCold.rootId == invalidExpected && invalidWarm.rootId == invalidExpected &&
                   encodeAfterInvalidCold - encodeBeforeInvalid == 1u && encodeAfterInvalidWarm == encodeAfterInvalidCold,
               L"invalid UTF-16 keeps the same failed encoding on cold and matching warm routes",
               passed,
               failed);

    // Route classification follows the lexical mount manager contract and must
    // never chase a drive alias into a backing root (storage probing does so separately).
    state.volumeRoot       = L"R:\\";
    state.volumeNameError  = ERROR_NOT_SUPPORTED;
    state.pathRootMappings = {{L"R:\\SubstAlias\\file.bin", L"R:\\"}};
    const auto substRoute  = BuildLocalCapabilityRouteInfo(L"R:\\SubstAlias\\file.bin");
    DebugCheck(substRoute.rootId == EncodeLocalRootId(L"R:\\") && state.queryDosDeviceCalls == 0u,
               L"route identity must preserve SUBST lexical-root fallback without querying its backing target",
               passed,
               failed);
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
    LocalVolumeRouteCache testRouteCache;
    StorageProbeDebugScope scope(hooks, testRouteCache);

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    wil::com_ptr<FileSystem> fileSystem;
    fileSystem.attach(new (std::nothrow) FileSystem());
    if (! DebugCheck(static_cast<bool>(fileSystem), L"mounted-volume storage probe selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    const HRESULT hr                        = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Folder\\file.bin", &characteristics);
    const unsigned int firstVolumeNameCalls = state.volumeNameCalls;
    const HRESULT cachedHr                  = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Other\\second.bin", &characteristics);

    DebugCheck(hr == S_OK, L"mounted-volume storage probe should keep GetStorageCharacteristics non-failing", passed, failed);
    DebugCheck(cachedHr == S_OK, L"cached mounted-volume storage query should remain non-failing", passed, failed);
    DebugCheck(state.volumeNameCalls == firstVolumeNameCalls + 1u && state.openVolumeCalls == 1u,
               L"a cache hit must revalidate the current GUID while reusing the per-instance physical probe",
               passed,
               failed);
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

    state.volumeName                     = L"\\\\?\\Volume{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}\\";
    const HRESULT remountedHr            = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\AfterRemount\\file.bin", &characteristics);
    const std::wstring remountedExpected = BuildOpenableVolumeGuidDevicePath(state.volumeName);
    DebugCheck(remountedHr == S_OK && state.openVolumeCalls == 2u && state.openedVolumePath == remountedExpected,
               L"a remounted lexical root must reprobe the newly observed backing volume GUID",
               passed,
               failed);

    state.volumeNameError                  = ERROR_NOT_SUPPORTED;
    const HRESULT fallbackHr1              = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Fallback\\one.bin", &characteristics);
    const unsigned int opensAfterFallback1 = state.openVolumeCalls;
    const HRESULT fallbackHr2              = fileSystem->GetStorageCharacteristics(L"R:\\MountedVolume\\Fallback\\two.bin", &characteristics);
    DebugCheck(fallbackHr1 == S_OK && fallbackHr2 == S_OK && opensAfterFallback1 == 3u && state.openVolumeCalls == 4u && state.openedVolumePath == L"\\\\.\\R:",
               L"without a stable volume GUID, drive-letter fallback probes must not be retained as physical cache entries",
               passed,
               failed);
    fileSystem.reset();

    // Reuse the existing storage hooks; these controls isolate output-buffer
    // behavior from cache/per-instance storage policy and real-volume fixtures.
    state                         = DebugMountedVolumeStorageProbeState{};
    const std::wstring shortInput = state.volumeRoot + L"file.bin";
    const auto shortRoot          = ResolveLocalVolumeRootPath(shortInput);
    const auto shortGuid          = ResolveVolumeGuidDevicePath(shortRoot);
    DebugCheck(shortRoot == state.volumeRoot && state.volumePathCalls == 1u && state.firstVolumePathBufferLength == MAX_PATH,
               L"short mounted root should use one actual wrapper call and preserve the returned mount root",
               passed,
               failed);
    DebugCheck(shortGuid == expectedOpenPath && state.volumeNameCalls == 1u && state.lastVolumeNameBufferLength == 50u,
               L"GUID result should preserve exact openable identity through the documented bounded output buffer",
               passed,
               failed);

    state                = DebugMountedVolumeStorageProbeState{};
    state.volumeRoot     = L"R:\\" + std::wstring(MAX_PATH + 17u, L'x') + L"\\";
    const auto longInput = state.volumeRoot + L"file.bin";
    const auto longRoot  = ResolveLocalVolumeRootPath(longInput);
    DebugCheck(longRoot == state.volumeRoot && state.volumePathCalls == 2u && state.firstVolumePathBufferLength == MAX_PATH &&
                   state.lastVolumePathBufferLength == 32768u && state.lastVolumePathInput == longInput,
               L"long mount root should retry the exact input once at the prior bound, without lexical drive fallback",
               passed,
               failed);
    const auto longGuid = ResolveVolumeGuidDevicePath(longRoot);
    DebugCheck(longGuid == expectedOpenPath && state.lastVolumeNameInput == state.volumeRoot && state.lastVolumeNameBufferLength == 50u,
               L"long mounted input must not truncate or change the GUID query's root",
               passed,
               failed);

    state                   = DebugMountedVolumeStorageProbeState{};
    state.volumeRoot        = L"R:\\" + std::wstring(MAX_PATH - 4u, L'x') + L"\\";
    state.oneShortSuccess   = true;
    const auto boundaryRoot = ResolveLocalVolumeRootPath(state.volumeRoot + L"file.bin");
    DebugCheck(boundaryRoot == state.volumeRoot && boundaryRoot.ends_with(L'\\') && state.volumePathCalls == 2u && state.lastVolumePathBufferLength == 32768u,
               L"one-short successful output without a slash must be retried, not used as mount identity",
               passed,
               failed);

    state                 = DebugMountedVolumeStorageProbeState{};
    state.volumePathError = ERROR_ACCESS_DENIED;
    const auto failedRoot = ResolveLocalVolumeRootPath(L"R:\\MountedVolume\\file.bin");
    DebugCheck(failedRoot.empty() && state.volumePathCalls == 2u && state.lastVolumePathBufferLength == 32768u,
               L"persistent volume query failure must stay empty after one bounded retry",
               passed,
               failed);
    state.volumePathError = ERROR_SUCCESS;
    state.volumeNameError = ERROR_PATH_NOT_FOUND;
    DebugCheck(ResolveVolumeGuidDevicePath(state.volumeRoot).empty() && state.volumeNameCalls == 1u,
               L"GUID failure must stay empty so existing SUBST/storage fallback policy owns recovery",
               passed,
               failed);

    state            = DebugMountedVolumeStorageProbeState{};
    state.volumeRoot = L"R:\\" + std::wstring(32768u, L'x') + L"\\";
    DebugCheck(ResolveLocalVolumeRootPath(state.volumeRoot + L"file.bin").empty() && state.volumePathCalls == 2u && state.lastVolumePathBufferLength == 32768u,
               L"root output beyond the existing bound must not grow, truncate or invent capability identity",
               passed,
               failed);
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
    unsigned int openVolumeCalls     = 0;
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
    if (OrdinalString::StartsWithNoCase(requestedPath, L"R:\\SubstAlias\\"))
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
        ++state->openVolumeCalls;
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
    LocalVolumeRouteCache testRouteCache;
    StorageProbeDebugScope scope(hooks, testRouteCache);

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    wil::com_ptr<FileSystem> fileSystem;
    fileSystem.attach(new (std::nothrow) FileSystem());
    if (! DebugCheck(static_cast<bool>(fileSystem), L"SUBST storage probe selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    const HRESULT hr = fileSystem->GetStorageCharacteristics(L"R:\\SubstAlias\\file.bin", &characteristics);

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

    const HRESULT cachedHr = fileSystem->GetStorageCharacteristics(L"R:\\SubstAlias\\second.bin", &characteristics);
    DebugCheck(cachedHr == S_OK && state.openVolumeCalls == 1u,
               L"SUBST physical cache hits must revalidate backing identity but reuse the medium probe",
               passed,
               failed);
    state.volumeName          = L"\\\\?\\Volume{77777777-8888-9999-AAAA-BBBBBBBBBBBB}\\";
    const HRESULT remountedHr = fileSystem->GetStorageCharacteristics(L"R:\\SubstAlias\\remounted.bin", &characteristics);
    DebugCheck(remountedHr == S_OK && state.openVolumeCalls == 2u && state.openedVolumePath == BuildOpenableVolumeGuidDevicePath(state.volumeName),
               L"a changed SUBST backing GUID must replace the old physical medium cache entry",
               passed,
               failed);
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
    LocalVolumeRouteCache testRouteCache;
    StorageProbeDebugScope scope(hooks, testRouteCache);

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

struct DebugUncStorageProbeState final
{
    std::vector<std::wstring> driveTypeRoots;
    unsigned int volumePathCalls = 0u;
};

UINT DebugUncGetDriveType(void* context, const wchar_t* rootPath) noexcept
{
    auto& state = *static_cast<DebugUncStorageProbeState*>(context);
    state.driveTypeRoots.emplace_back(rootPath != nullptr ? rootPath : L"");
    return DRIVE_REMOTE;
}

BOOL DebugUncGetVolumePathName(void* context, const wchar_t*, wchar_t*, DWORD) noexcept
{
    ++static_cast<DebugUncStorageProbeState*>(context)->volumePathCalls;
    SetLastError(ERROR_INVALID_PARAMETER);
    return FALSE;
}

// An extended UNC path (`\\?\UNC\server\share\...`) names the same share as its plain form: it is a
// UNC path, its root is the plain `\\server\share\` that GetDriveTypeW and the storage cache key use,
// it is high-latency for the storage and transfer queries, and admission gives it the same remote root.
// The drive type is hooked, so nothing here reaches the network.
void RunDebugUncPathClassificationSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const std::wstring plain     = L"\\\\server\\share\\dir\\file.bin";
    const std::wstring extended  = L"\\\\?\\UNC\\server\\share\\dir\\file.bin";
    const std::wstring shareRoot = L"\\\\server\\share\\";

    DebugCheck(IsUncPath(plain) && IsUncPath(extended) && IsUncPath(L"\\\\?\\unc\\server\\share") && IsUncPath(L"//server/share/dir"),
               L"plain, extended (in any case) and slash-separated UNC paths are UNC",
               passed,
               failed);
    DebugCheck(! IsUncPath(L"C:\\dir") && ! IsUncPath(L"\\\\?\\C:\\dir") && ! IsUncPath(L"\\\\?\\Volume{00000000-0000-0000-0000-000000000000}\\dir") &&
                   ! IsUncPath(L"\\\\.\\C:\\dir") && ! IsUncPath(L"\\\\.\\pipe\\name"),
               L"drive, extended drive, volume GUID and device paths are not UNC",
               passed,
               failed);

    DebugCheck(ExtractDriveRoot(plain) == shareRoot && ExtractDriveRoot(extended) == shareRoot,
               L"an extended UNC path has the plain share root of its plain form",
               passed,
               failed);
    DebugCheck(ExtractDriveRoot(L"\\\\server\\share") == shareRoot && ExtractDriveRoot(L"\\\\?\\UNC\\server\\share") == shareRoot,
               L"a share named without its trailing separator still has the share root, in both forms",
               passed,
               failed);
    DebugCheck(ExtractDriveRoot(L"\\\\server").empty() && ExtractDriveRoot(L"\\\\?\\UNC\\server").empty(),
               L"a UNC path without a share has no root, in both forms",
               passed,
               failed);
    DebugCheck(ExtractDriveRoot(L"c:\\dir") == L"C:\\" && ExtractDriveRoot(L"\\\\?\\c:\\dir") == L"C:\\",
               L"a drive path has its drive root, in both forms",
               passed,
               failed);
    DebugCheck(ExtractDriveRoot(L"\\\\?\\Volume{00000000-0000-0000-0000-000000000000}\\dir").empty() && ExtractDriveRoot(L"\\\\.\\C:\\dir").empty(),
               L"volume GUID and device paths have no drive root",
               passed,
               failed);

    DebugUncStorageProbeState state;
    StorageProbeDebugHooks hooks{
        .context           = &state,
        .getVolumePathName = DebugUncGetVolumePathName,
        .getDriveType      = DebugUncGetDriveType,
    };
    LocalVolumeRouteCache testRouteCache;
    StorageProbeDebugScope scope(hooks, testRouteCache);
    auto* fileSystem = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"UNC classification selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }

    FileSystemStorageCharacteristics plainStorage{};
    plainStorage.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    FileSystemStorageCharacteristics extendedStorage{};
    extendedStorage.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    FileSystemTransferHints plainHints{};
    plainHints.sizeBytes = sizeof(FileSystemTransferHints);
    FileSystemTransferHints extendedHints{};
    extendedHints.sizeBytes         = sizeof(FileSystemTransferHints);
    const HRESULT plainStorageHr    = fileSystem->GetStorageCharacteristics(plain.c_str(), &plainStorage);
    const HRESULT extendedStorageHr = fileSystem->GetStorageCharacteristics(extended.c_str(), &extendedStorage);
    const HRESULT plainHintsHr      = fileSystem->GetTransferHints(plain.c_str(), FILESYSTEM_COPY, FILESYSTEM_TRANSFER_DESTINATION_WRITE, &plainHints);
    const HRESULT extendedHintsHr   = fileSystem->GetTransferHints(extended.c_str(), FILESYSTEM_COPY, FILESYSTEM_TRANSFER_DESTINATION_WRITE, &extendedHints);
    const LocalCapabilityRouteInfo plainRoute    = BuildLocalCapabilityRouteInfo(plain.c_str());
    const LocalCapabilityRouteInfo extendedRoute = BuildLocalCapabilityRouteInfo(extended.c_str());
    fileSystem->Release();

    DebugCheck(SUCCEEDED(plainStorageHr) && SUCCEEDED(extendedStorageHr) && SUCCEEDED(plainHintsHr) && SUCCEEDED(extendedHintsHr),
               L"UNC storage and transfer queries succeed in both forms",
               passed,
               failed);
    DebugCheck((plainStorage.flags & FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY) != 0u && (extendedStorage.flags & FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY) != 0u &&
                   plainStorage.storageKind == FILESYSTEM_STORAGE_NETWORK_SHARE && extendedStorage.storageKind == FILESYSTEM_STORAGE_NETWORK_SHARE,
               L"an extended UNC path is a high-latency network share, as its plain form is",
               passed,
               failed);
    DebugCheck(plainHints.latencyClass == FILESYSTEM_TRANSFER_LATENCY_LAN && extendedHints.latencyClass == FILESYSTEM_TRANSFER_LATENCY_LAN &&
                   plainHints.preferredBufferBytes == extendedHints.preferredBufferBytes && extendedHints.preferredBufferBytes == 8u * 1024u * 1024u,
               L"an extended UNC path gets the high-latency transfer hints of its plain form",
               passed,
               failed);
    DebugCheck(state.driveTypeRoots.size() == 4u && std::ranges::all_of(state.driveTypeRoots, [&](const std::wstring& root) { return root == shareRoot; }),
               L"every UNC storage and transfer query asks the drive type of the plain share root",
               passed,
               failed);
    DebugCheck(state.volumePathCalls == 0u, L"a UNC path, plain or extended, never resolves a local volume root", passed, failed);
    DebugCheck(plainRoute.remote && extendedRoute.remote && ! plainRoute.rootId.empty() && plainRoute.rootId == extendedRoute.rootId,
               L"admission gives an extended UNC path the remote root of its plain form",
               passed,
               failed);
}

void RunDebugOperationControlSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugSynchronousIoCancelWatchSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugCreateDirectoryAdmissionSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugDeleteManifestSelfTest(unsigned int& passed, unsigned int& failed) noexcept;

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    RunDebugMountedVolumeStorageProbeSelfTest(*passed, *failed);
    RunDebugLocalRouteObservationSelfTest(*passed, *failed);
    RunDebugSubstStorageProbeSelfTest(*passed, *failed);
    RunDebugNvmeFallbackStorageProbeSelfTest(*passed, *failed);
    RunDebugUncPathClassificationSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugPathNormalizationSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugReparseCopyErrorMappingSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugRecycleItemVerdictSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugDirectorySizeErrorPolicySelfTest(*passed, *failed);
    FileSystemInternal::RunDebugSharedFileOpsSchedulerShutdownSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugSearchServiceFallbackCandidateSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugObjectBindingSelfTest(*passed, *failed);
    FileSystemInternal::RunDebugRecycleHeldCancellationSelfTest(*passed, *failed);
    RunDebugOperationControlSelfTest(*passed, *failed);
    RunDebugSynchronousIoCancelWatchSelfTest(*passed, *failed);
    RunDebugCreateDirectoryAdmissionSelfTest(*passed, *failed);
    RunDebugDeleteManifestSelfTest(*passed, *failed);

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
    Win32FileReader(wil::unique_handle file, uint64_t sizeBytes, const FileSystemOptions* options) noexcept
        : _file(std::move(file)),
          _sizeBytes(sizeBytes),
          _hasOperationOptions(options != nullptr)
    {
        if (options != nullptr)
        {
            _operationOptions = *options;
        }
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

        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
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

        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        uint64_t base = 0u;
        if (origin == FILE_CURRENT)
        {
            base = _position;
        }
        else if (origin == FILE_END)
        {
            base = _sizeBytes;
        }

        uint64_t target = 0u;
        if (offset < 0)
        {
            const uint64_t magnitude = static_cast<uint64_t>(-(offset + 1)) + 1u;
            if (magnitude > base)
            {
                return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
            }
            target = base - magnitude;
        }
        else
        {
            const uint64_t magnitude = static_cast<uint64_t>(offset);
            if (magnitude > (std::numeric_limits<uint64_t>::max)() - base)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            target = base + magnitude;
        }
        if (target > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        _position    = target;
        *newPosition = target;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

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

        if (! _readEvent)
        {
            _readEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
            if (! _readEvent)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_OUTOFMEMORY);
            }
        }
        static_cast<void>(ResetEvent(_readEvent.get()));

        // Only a reader that carries an operation control or a deadline polls between kernel
        // waits; every other consumer (viewers, search, hashing) waits without a timer.
        const FileSystemOptions* const control = OperationOptions();
        const DWORD waitMs = control != nullptr && (control->operationControl != nullptr || control->deadlineTickCount64 != 0u) ? 25u : INFINITE;

        OVERLAPPED overlapped{};
        overlapped.Offset     = static_cast<DWORD>(_position & 0xFFFFFFFFu);
        overlapped.OffsetHigh = static_cast<DWORD>(_position >> 32u);
        overlapped.hEvent     = _readEvent.get();

        const BOOL started = ReadFile(_file.get(), buffer, bytesToRead, nullptr, &overlapped);
        if (started == FALSE)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_HANDLE_EOF)
            {
                // A read at or past the end may fail synchronously instead of going pending;
                // which one the volume chooses is not the caller's concern. End of file is the
                // same ordinary zero-byte result on both routes.
                return S_OK;
            }
            if (error != ERROR_IO_PENDING)
            {
                return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_READ_FAULT);
            }
        }

        DWORD read = 0u;
        for (;;)
        {
            const DWORD waitResult = WaitForSingleObject(_readEvent.get(), waitMs);
            if (waitResult == WAIT_OBJECT_0)
            {
                if (GetOverlappedResult(_file.get(), &overlapped, &read, FALSE) == FALSE)
                {
                    const DWORD error = GetLastError();
                    if (error == ERROR_HANDLE_EOF)
                    {
                        return S_OK;
                    }
                    return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_READ_FAULT);
                }
                break;
            }

            const HRESULT pendingControlHr = FileSystemCheckOperationControl(OperationOptions());
            if (FAILED(pendingControlHr))
            {
                static_cast<void>(CancelIoEx(_file.get(), &overlapped));
                DWORD ignoredBytes = 0u;
                static_cast<void>(GetOverlappedResult(_file.get(), &overlapped, &ignoredBytes, TRUE));
                return pendingControlHr;
            }
            if (waitResult == WAIT_TIMEOUT)
            {
                continue;
            }

            const HRESULT waitHr = waitResult == WAIT_FAILED ? HRESULT_FROM_WIN32(GetLastError()) : E_UNEXPECTED;
            static_cast<void>(CancelIoEx(_file.get(), &overlapped));
            DWORD ignoredBytes = 0u;
            static_cast<void>(GetOverlappedResult(_file.get(), &overlapped, &ignoredBytes, TRUE));
            return waitHr;
        }

        _position += static_cast<uint64_t>(read);
        *bytesRead = static_cast<unsigned long>(read);
        return S_OK;
    }

private:
    ~Win32FileReader() = default;

    [[nodiscard]] const FileSystemOptions* OperationOptions() const noexcept
    {
        return _hasOperationOptions ? &_operationOptions : nullptr;
    }

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    uint64_t _sizeBytes = 0;
    uint64_t _position  = 0;
    wil::unique_event_nothrow _readEvent;
    FileSystemOptions _operationOptions{};
    bool _hasOperationOptions = false;
};

struct LocalObjectIdentityPayload final
{
    std::array<unsigned char, 8> namespaceTag{{'R', 'S', 'L', 'F', 'I', 'D', '2', 0}};
    uint64_t volumeSerialNumber = 0;
    FILE_ID_128 fileId{};
};
static_assert(sizeof(LocalObjectIdentityPayload) == 32u);

constexpr std::array<unsigned char, 8> kPersistentFileIdNamespace{{'R', 'S', 'L', 'F', 'I', 'D', '2', 0}};
constexpr std::array<unsigned char, 8> kRetainedHandleNamespace{{'R', 'S', 'L', 'H', 'N', 'D', '1', 0}};

[[nodiscard]] bool IsRetainedHandleIdentity(const LocalObjectIdentityPayload& identity) noexcept
{
    return identity.namespaceTag == kRetainedHandleNamespace;
}

[[nodiscard]] HRESULT MakePersistentLocalObjectIdentity(const FILE_ID_INFO& fileIdInfo, LocalObjectIdentityPayload& identity) noexcept
{
    const bool emptyFileId =
        std::all_of(std::begin(fileIdInfo.FileId.Identifier), std::end(fileIdInfo.FileId.Identifier), [](BYTE value) noexcept { return value == 0u; });
    if (emptyFileId)
    {
        identity = {};
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    identity.namespaceTag       = kPersistentFileIdNamespace;
    identity.volumeSerialNumber = fileIdInfo.VolumeSerialNumber;
    identity.fileId             = fileIdInfo.FileId;
    return S_OK;
}

[[nodiscard]] HRESULT MakeRetainedHandleIdentity(LocalObjectIdentityPayload& identity) noexcept
{
    GUID nonce{};
    const HRESULT nonceHr = CoCreateGuid(&nonce);
    if (FAILED(nonceHr))
    {
        identity = {};
        return nonceHr;
    }

    identity              = {};
    identity.namespaceTag = kRetainedHandleNamespace;
    static_assert(sizeof(nonce) == sizeof(identity.fileId));
    std::memcpy(&identity.fileId, &nonce, sizeof(nonce));
    return S_OK;
}

[[nodiscard]] HRESULT QueryExactLocalObjectIdentity(HANDLE file, LocalObjectIdentityPayload& identity) noexcept
{
    identity = {};
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    FILE_ID_INFO fileIdInfo{};
    if (GetFileInformationByHandleEx(file, FileIdInfo, &fileIdInfo, sizeof(fileIdInfo)) == FALSE)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_INVALID_PARAMETER || error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    return MakePersistentLocalObjectIdentity(fileIdInfo, identity);
}

[[nodiscard]] HRESULT QueryPersistentIdentityFromRetainedHandle(HANDLE file, LocalObjectIdentityPayload& identity) noexcept
{
    identity = {};
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    wil::unique_handle metadataHandle(
        ReOpenFile(file, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, FILE_FLAG_OPEN_REPARSE_POINT));
    if (! metadataHandle)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    return QueryExactLocalObjectIdentity(metadataHandle.get(), identity);
}

[[nodiscard]] HRESULT ResolveNewStageIdentity(HANDLE file,
                                              const wchar_t* stagePath,
                                              std::wstring_view objectKind,
                                              const LocalObjectIdentityPayload& retainedIdentity,
                                              LocalObjectIdentityPayload& identity) noexcept
{
#if ! defined(ENABLE_TESTS)
    static_cast<void>(stagePath);
#endif
    HRESULT identityHr = S_OK;
#if defined(ENABLE_TESTS)
    if (ConsumeMatchingPathInjectionForSelfTest(
            stagePath, kStageIdentityUnsupportedPathEnvVar, kStageIdentityUnsupportedFiredEnvVar, L"FileOps.Local.StageIdentityUnsupportedInjected"))
    {
        identityHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    else
#endif
    {
        identityHr = QueryExactLocalObjectIdentity(file, identity);
    }
    Debug::Perf::Emit(L"FileOps.Local.ExclusiveIdentityQuery", objectKind, 0u, 0u, 0u, identityHr);

    if (SUCCEEDED(identityHr))
    {
        Debug::Perf::EmitCounter(L"fileops.local.stage_identity.persistent_file_id");
        return S_OK;
    }

    // CREATE_NEW plus the retained handle is exact ownership authority even when a redirector
    // cannot supply FileIdInfo. The nonce is process-local identity; it is never reconstructed
    // from a 64-bit file index or by reopening the stage path.
    identity = retainedIdentity;
    Debug::Perf::EmitCounter(L"fileops.local.stage_identity.retained_handle");
    return S_OK;
}

interface __declspec(uuid("c1884e97-02e5-4c89-99b0-a619e8f78625")) __declspec(novtable) ILocalBoundObjectControl : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE RenameExactTo(
        const wchar_t* destinationPath, BOOL replaceIfExists, FileSystemConditionalMutationResult* result) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE DeleteExact(FileSystemConditionalMutationResult * result) noexcept            = 0;
    virtual HRESULT STDMETHODCALLTYPE DuplicateExactHandle(HANDLE * duplicated) noexcept                            = 0;
    virtual HRESULT STDMETHODCALLTYPE RetireDirectoryReadPin(const LocalObjectIdentityPayload* expected) noexcept   = 0;
    // The owned stage's side of IFileSystemExpectedOccupantPublication::PublishOverExpectedOccupant.
    virtual HRESULT STDMETHODCALLTYPE PublishOwnedOverExpected(const wchar_t* finalPath,
                                                               const FileSystemExpectedOccupant* expected,
                                                               FileSystemFlags flags,
                                                               const FileSystemOptions* options,
                                                               FileSystemConditionalMutationResult* result,
                                                               IFileSystemBoundObject** published) noexcept = 0;
};

// Reads, without following a link, the facts of an existing item for the expected-occupant
// publication. identityAvailable reports whether the location can identify the item: where it can,
// the bound route is the authority and this contract does not apply.
[[nodiscard]] HRESULT ReadLocalOccupantFacts(const wchar_t* path, FileSystemExpectedOccupant& facts, bool& identityAvailable) noexcept
{
    facts             = {};
    facts.sizeBytes   = sizeof(facts);
    identityAvailable = false;
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    const std::wstring extended = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(extended.c_str(),
                                        FILE_READ_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    if (! file)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND);
    }
    BY_HANDLE_FILE_INFORMATION information{};
    if (GetFileInformationByHandle(file.get(), &information) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    facts.attributes    = information.dwFileAttributes;
    facts.creationTime  = (static_cast<__int64>(information.ftCreationTime.dwHighDateTime) << 32) | information.ftCreationTime.dwLowDateTime;
    facts.lastWriteTime = (static_cast<__int64>(information.ftLastWriteTime.dwHighDateTime) << 32) | information.ftLastWriteTime.dwLowDateTime;
    facts.fileSizeBytes = (static_cast<uint64_t>(information.nFileSizeHigh) << 32) | information.nFileSizeLow;
    facts.fileIndex     = (static_cast<uint64_t>(information.nFileIndexHigh) << 32) | information.nFileIndexLow;
#if defined(ENABLE_TESTS)
    if (IsBindIdentityUnsupportedForSelfTest(path))
    {
        return S_OK;
    }
#endif
    LocalObjectIdentityPayload identity{};
    const HRESULT identityHr = QueryExactLocalObjectIdentity(file.get(), identity);
    if (SUCCEEDED(identityHr))
    {
        identityAvailable = true;
        return S_OK;
    }
    return identityHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? S_OK : identityHr;
}

// The occupant is the one the user decided on: same kind and protection attributes, same size, same
// creation and last-write time, and the same file index where the location reports one.
[[nodiscard]] bool LocalOccupantFactsEqual(const FileSystemExpectedOccupant& left, const FileSystemExpectedOccupant& right) noexcept
{
    constexpr unsigned long kComparedAttributes =
        FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM;
    return ((left.attributes ^ right.attributes) & kComparedAttributes) == 0u && left.creationTime == right.creationTime &&
           left.lastWriteTime == right.lastWriteTime && left.fileSizeBytes == right.fileSizeBytes && left.fileIndex == right.fileIndex;
}

constexpr size_t kBackupStreamHeaderBytes         = offsetof(WIN32_STREAM_ID, cStreamName);
constexpr DWORD kBackupMetadataMaxStreamNameBytes = 64u * 1024u;
constexpr size_t kBackupMetadataIoBufferBytes     = 64u * 1024u;
constexpr uint32_t kMetadataBackupFeatures        = FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS | FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES;

[[nodiscard]] HRESULT ReopenExactFile(HANDLE source, DWORD desiredAccess, wil::unique_handle& reopened, DWORD additionalFlags = 0u) noexcept
{
    reopened.reset(ReOpenFile(source,
                              desiredAccess,
                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                              FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS | additionalFlags));
    if (! reopened)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    return S_OK;
}

[[nodiscard]] HRESULT BackupReadExact(HANDLE file, void*& context, void* buffer, DWORD bytesToRead, DWORD& bytesRead) noexcept
{
    bytesRead    = 0u;
    auto* output = static_cast<std::byte*>(buffer);
    while (bytesRead < bytesToRead)
    {
        DWORD current = 0u;
        if (BackupRead(file, reinterpret_cast<LPBYTE>(output + bytesRead), bytesToRead - bytesRead, &current, FALSE, FALSE, &context) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        if (current == 0u)
        {
            break;
        }
        bytesRead += current;
    }
    return S_OK;
}

[[nodiscard]] HRESULT BackupWriteExact(HANDLE file, void*& context, const void* buffer, DWORD bytesToWrite) noexcept
{
    DWORD bytesWritten = 0u;
    const auto* input  = static_cast<const std::byte*>(buffer);
    while (bytesWritten < bytesToWrite)
    {
        DWORD current = 0u;
        if (BackupWrite(
                file, const_cast<LPBYTE>(reinterpret_cast<const BYTE*>(input + bytesWritten)), bytesToWrite - bytesWritten, &current, FALSE, FALSE, &context) ==
            FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        if (current == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        bytesWritten += current;
    }
    return S_OK;
}

[[nodiscard]] HRESULT SkipBackupPayload(HANDLE file, void*& context, uint64_t bytes) noexcept
{
    while (bytes != 0u)
    {
        const DWORD low   = static_cast<DWORD>(bytes & 0xffffffffu);
        const DWORD high  = static_cast<DWORD>(bytes >> 32u);
        DWORD lowSkipped  = 0u;
        DWORD highSkipped = 0u;
        if (BackupSeek(file, low, high, &lowSkipped, &highSkipped, &context) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        const uint64_t skipped = (static_cast<uint64_t>(highSkipped) << 32u) | lowSkipped;
        if (skipped == 0u || skipped > bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        bytes -= skipped;
    }
    return S_OK;
}

[[nodiscard]] uint32_t BackupStreamMetadataFeature(const WIN32_STREAM_ID& stream, std::wstring_view streamName) noexcept
{
    if (stream.dwStreamId == BACKUP_EA_DATA)
    {
        return FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES;
    }
    if (stream.dwStreamId != BACKUP_ALTERNATE_DATA)
    {
        return 0u;
    }
    return OrdinalString::EqualsNoCase(streamName, L":Zone.Identifier:$DATA") ? FILESYSTEM_METADATA_MOTW : FILESYSTEM_METADATA_ALTERNATE_STREAMS;
}

[[nodiscard]] HRESULT InspectBackupMetadataStreams(HANDLE exactFile, uint32_t& presentFeatures) noexcept
{
    presentFeatures &= ~kMetadataBackupFeatures;
    wil::unique_handle source;
    HRESULT hr = ReopenExactFile(exactFile, GENERIC_READ, source);
    if (FAILED(hr))
    {
        return hr;
    }

    void* readContext    = nullptr;
    const auto abortRead = wil::scope_exit([&]() noexcept
    {
        DWORD ignored = 0u;
        static_cast<void>(BackupRead(source.get(), nullptr, 0u, &ignored, TRUE, FALSE, &readContext));
    });
    for (;;)
    {
        WIN32_STREAM_ID stream{};
        DWORD headerBytes = 0u;
        hr                = BackupReadExact(source.get(), readContext, &stream, static_cast<DWORD>(kBackupStreamHeaderBytes), headerBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        if (headerBytes == 0u)
        {
            return S_OK;
        }
        if (headerBytes != kBackupStreamHeaderBytes || stream.dwStreamNameSize > kBackupMetadataMaxStreamNameBytes ||
            stream.dwStreamNameSize % sizeof(wchar_t) != 0u || stream.Size.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::wstring name(static_cast<size_t>(stream.dwStreamNameSize / sizeof(wchar_t)), L'\0');
        DWORD nameBytes = 0u;
        if (stream.dwStreamNameSize != 0u)
        {
            hr = BackupReadExact(source.get(), readContext, name.data(), stream.dwStreamNameSize, nameBytes);
            if (FAILED(hr) || nameBytes != stream.dwStreamNameSize)
            {
                return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        presentFeatures |= BackupStreamMetadataFeature(stream, name);
        hr = SkipBackupPayload(source.get(), readContext, static_cast<uint64_t>(stream.Size.QuadPart));
        if (FAILED(hr))
        {
            return hr;
        }
    }
}

[[nodiscard]] HRESULT CopyBackupMetadataStreams(HANDLE exactSource, HANDLE exactDestination, FileSystemMetadataTransferResult& result) noexcept
{
    wil::unique_handle source;
    wil::unique_handle destination;
    HRESULT hr = ReopenExactFile(exactSource, GENERIC_READ, source);
    if (SUCCEEDED(hr))
    {
        hr = ReopenExactFile(exactDestination, GENERIC_WRITE, destination);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    void* readContext        = nullptr;
    void* writeContext       = nullptr;
    const auto abortContexts = wil::scope_exit([&]() noexcept
    {
        DWORD ignored = 0u;
        static_cast<void>(BackupRead(source.get(), nullptr, 0u, &ignored, TRUE, FALSE, &readContext));
        static_cast<void>(BackupWrite(destination.get(), nullptr, 0u, &ignored, TRUE, FALSE, &writeContext));
    });
    // This path runs under the already-deep host bridge call stack. Keep the
    // bounded transfer buffer off the thread stack so Debug/selftest callers do
    // not cross the Windows stack guard while preserving the same I/O quantum.
    std::unique_ptr<std::byte[]> buffer(new (std::nothrow) std::byte[kBackupMetadataIoBufferBytes]);
    if (! buffer)
    {
        return E_OUTOFMEMORY;
    }
    for (;;)
    {
        WIN32_STREAM_ID stream{};
        DWORD headerBytes = 0u;
        hr                = BackupReadExact(source.get(), readContext, &stream, static_cast<DWORD>(kBackupStreamHeaderBytes), headerBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        if (headerBytes == 0u)
        {
            return S_OK;
        }
        if (headerBytes != kBackupStreamHeaderBytes || stream.dwStreamNameSize > kBackupMetadataMaxStreamNameBytes ||
            stream.dwStreamNameSize % sizeof(wchar_t) != 0u || stream.Size.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::vector<std::byte> nameBytes(stream.dwStreamNameSize);
        DWORD nameRead = 0u;
        if (! nameBytes.empty())
        {
            hr = BackupReadExact(source.get(), readContext, nameBytes.data(), stream.dwStreamNameSize, nameRead);
            if (FAILED(hr) || nameRead != stream.dwStreamNameSize)
            {
                return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        const std::wstring_view name(reinterpret_cast<const wchar_t*>(nameBytes.data()), nameBytes.size() / sizeof(wchar_t));
        const uint32_t feature      = BackupStreamMetadataFeature(stream, name);
        const uint64_t payloadBytes = static_cast<uint64_t>(stream.Size.QuadPart);
        if (feature == 0u)
        {
            hr = SkipBackupPayload(source.get(), readContext, payloadBytes);
            if (FAILED(hr))
            {
                return hr;
            }
            continue;
        }

        result.attemptedFeatures |= feature;
        hr = BackupWriteExact(destination.get(), writeContext, &stream, static_cast<DWORD>(kBackupStreamHeaderBytes));
        if (SUCCEEDED(hr) && ! nameBytes.empty())
        {
            hr = BackupWriteExact(destination.get(), writeContext, nameBytes.data(), static_cast<DWORD>(nameBytes.size()));
        }
        uint64_t remaining = payloadBytes;
        while (SUCCEEDED(hr) && remaining != 0u)
        {
            const DWORD chunk = static_cast<DWORD>(std::min<uint64_t>(remaining, kBackupMetadataIoBufferBytes));
            DWORD bytesRead   = 0u;
            hr                = BackupReadExact(source.get(), readContext, buffer.get(), chunk, bytesRead);
            if (SUCCEEDED(hr) && bytesRead != chunk)
            {
                hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            if (SUCCEEDED(hr))
            {
                hr = BackupWriteExact(destination.get(), writeContext, buffer.get(), chunk);
            }
            remaining -= SUCCEEDED(hr) ? chunk : 0u;
        }
        if (FAILED(hr))
        {
            result.lostFeatures |= feature;
            if (SUCCEEDED(result.firstFailure))
            {
                result.firstFailure = hr;
            }
            return S_OK;
        }
        result.preservedFeatures |= feature;
    }
}

static_assert(Common::SynchronousIoCancelWatch::kMaxCleanupWindowMs == FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS,
              "the cancel watch's cleanup window is the owned-stage cleanup budget");

// I35 R7.4: the cleanup window of an owned-object cleanup control (MakeOwnedStageCleanupOptions: no
// operation control, a fresh deadline); 0, no window, for any primary control.
[[nodiscard]] ULONGLONG OwnedStageCleanupWindowDeadline(const FileSystemOptions* options) noexcept
{
    return options != nullptr && options->operationControl == nullptr ? options->deadlineTickCount64 : 0u;
}

// I35 R7.4: a fresh cleanup budget for compensation that has no cleanup control of its own (the raw
// removal of a stage just created, a publication rollback that restores the user's original).
[[nodiscard]] ULONGLONG FreshOwnedStageCleanupDeadline() noexcept
{
    return GetTickCount64() + FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS;
}

#if defined(ENABLE_TESTS)
constexpr std::wstring_view kStageCleanupDelayRootEnvVar  = L"REDSALAMANDER_FILEOPS_STAGE_CLEANUP_DELAY_ROOT";
constexpr std::wstring_view kStageCleanupDelayFiredEnvVar = L"REDSALAMANDER_FILEOPS_STAGE_CLEANUP_DELAY_FIRED";

// I35 R7.4: holds the removal of one owned stage at or below the configured root for one second in a
// synchronous pipe read that a helper thread completes. Unlike a sleep, CancelSynchronousIo can cancel
// that read, so a case passes only while the cancel watch leaves owned-stage cleanup alone for its
// budget. One-shot: the variable is cleared and the FIRED variable records "1" or "aborted".
[[nodiscard]] HRESULT DelayOwnedStageCleanupForSelfTest(const std::wstring& stagePath) noexcept
{
    const DWORD required = GetEnvironmentVariableW(kStageCleanupDelayRootEnvVar.data(), nullptr, 0u);
    if (required == 0u || stagePath.empty())
    {
        return S_OK;
    }
    std::wstring configured(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(kStageCleanupDelayRootEnvVar.data(), configured.data(), required);
    if (written == 0u || written >= required)
    {
        return S_OK;
    }
    configured.resize(written);
    while (! configured.empty() && (configured.back() == L'\\' || configured.back() == L'/'))
    {
        configured.pop_back();
    }
    const std::wstring root   = FileSystemInternal::ToExtendedPath(configured);
    const std::wstring actual = FileSystemInternal::ToExtendedPath(stagePath);
    if (root.empty() || actual.size() <= root.size() || ! OrdinalString::EqualsNoCase(std::wstring_view(actual).substr(0u, root.size()), root) ||
        actual[root.size()] != L'\\')
    {
        return S_OK;
    }
    static_cast<void>(SetEnvironmentVariableW(kStageCleanupDelayRootEnvVar.data(), nullptr));

    HANDLE readRaw  = nullptr;
    HANDLE writeRaw = nullptr;
    if (CreatePipe(&readRaw, &writeRaw, nullptr, 0u) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wil::unique_handle readEnd(readRaw);
    wil::unique_handle writeEnd(writeRaw);
    BOOL completed = FALSE;
    DWORD error    = ERROR_SUCCESS;
    {
        std::jthread writer([&writeEnd]() noexcept
        {
            Sleep(1'000u);
            DWORD wrote = 0u;
            static_cast<void>(WriteFile(writeEnd.get(), "x", 1u, &wrote, nullptr));
        });
        std::array<std::byte, 1u> buffer{};
        DWORD read = 0u;
        completed  = ReadFile(readEnd.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr);
        error      = completed != FALSE ? ERROR_SUCCESS : GetLastError();
    }
    static_cast<void>(SetEnvironmentVariableW(kStageCleanupDelayFiredEnvVar.data(), completed != FALSE ? L"1" : L"aborted"));
    Debug::Perf::EmitCounter(L"FileOps.Local.StageCleanupDelayInjected");
    return completed != FALSE ? S_OK : HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_OPERATION_ABORTED);
}
#endif

[[nodiscard]] HRESULT DeleteExactHandle(HANDLE file) noexcept
{
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    FILE_DISPOSITION_INFO_EX disposition{};
    disposition.Flags = FILE_DISPOSITION_FLAG_DELETE | FILE_DISPOSITION_FLAG_POSIX_SEMANTICS | FILE_DISPOSITION_FLAG_IGNORE_READONLY_ATTRIBUTE;
    if (SetFileInformationByHandle(file, FileDispositionInfoEx, &disposition, sizeof(disposition)) != FALSE)
    {
        return S_OK;
    }

    DWORD error = GetLastError();
    if (error == ERROR_INVALID_PARAMETER || error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED)
    {
        FILE_BASIC_INFO basic{};
        bool restoreReadonly         = false;
        const auto restoreAttributes = wil::scope_exit([&]() noexcept
        {
            if (restoreReadonly)
            {
                if (SetFileInformationByHandle(file, FileBasicInfo, &basic, sizeof(basic)) == FALSE)
                    Debug::Warning(L"FileSystem: failed to restore read-only attributes after an exact deletion failed (error={}).", GetLastError());
            }
        });
        if (GetFileInformationByHandleEx(file, FileBasicInfo, &basic, sizeof(basic)) != FALSE && (basic.FileAttributes & FILE_ATTRIBUTE_READONLY) != 0u)
        {
            FILE_BASIC_INFO writable = basic;
            writable.FileAttributes &= ~FILE_ATTRIBUTE_READONLY;
            if (writable.FileAttributes == 0u)
            {
                writable.FileAttributes = FILE_ATTRIBUTE_NORMAL;
            }
            restoreReadonly = SetFileInformationByHandle(file, FileBasicInfo, &writable, sizeof(writable)) != FALSE;
        }

        FILE_DISPOSITION_INFO fallback{};
        fallback.DeleteFile = TRUE;
        if (SetFileInformationByHandle(file, FileDispositionInfo, &fallback, sizeof(fallback)) != FALSE)
        {
            restoreReadonly = false;
            return S_OK;
        }
        error = GetLastError();
    }

    FILE_STANDARD_INFO standard{};
    if (GetFileInformationByHandleEx(file, FileStandardInfo, &standard, sizeof(standard)) != FALSE && standard.DeletePending != FALSE)
    {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
}

[[nodiscard]] HRESULT DeletePathOnlyWhenPersistentIdentityMatches(const std::wstring& path, const LocalObjectIdentityPayload& expectedIdentity) noexcept
{
    if (path.empty() || IsRetainedHandleIdentity(expectedIdentity))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::unique_handle file(
        CreateFileW(path.c_str(), DELETE | FILE_READ_ATTRIBUTES, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    if (! file)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    LocalObjectIdentityPayload currentIdentity{};
    const HRESULT identityHr = QueryExactLocalObjectIdentity(file.get(), currentIdentity);
    if (FAILED(identityHr))
    {
        return identityHr;
    }
    if (std::memcmp(&currentIdentity, &expectedIdentity, sizeof(currentIdentity)) != 0)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    return DeleteExactHandle(file.get());
}

class LocalBoundObject final : public IFileSystemBoundObject,
                               public IFileSystemBoundMetadata,
                               public IFileSystemBoundContentProof,
                               public IFileSystemBoundObjectPinHandoff,
                               public ILocalBoundObjectControl
{
public:
    LocalBoundObject(wil::unique_handle file,
                     LocalObjectIdentityPayload identity,
                     FileSystemBoundObjectKind kind,
                     uint64_t committedSizeBytes,
                     FileSystemBindFlags grantedFlags,
                     bool ownedStage,
                     std::wstring path) noexcept
        : _file(std::move(file)),
          _identity(identity),
          _kind(kind),
          _committedSizeBytes(committedSizeBytes),
          _grantedFlags(grantedFlags),
          _ownedStage(ownedStage),
          _consumableReadWitness(! ownedStage && kind == FILESYSTEM_BOUND_DIRECTORY &&
                                 grantedFlags == static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA)),
          _path(std::move(path))
    {
#if defined(ENABLE_TESTS)
        _reportedIdentity     = _identity;
        const auto configured = EnvironmentVariables::Read(kBindIdentityNamespaceRootEnvVar);
        if (configured.has_value())
        {
            std::wstring root = FileSystemInternal::ToExtendedPath(configured.value());
            while (! root.empty() && (root.back() == L'\\' || root.back() == L'/'))
                root.pop_back();
            const std::wstring actual = FileSystemInternal::ToExtendedPath(_path);
            if (! root.empty() && Common::Paths::IsSameOrDescendantNormalizedWindowsPath(root, actual))
            {
                // Snapshot-only fixture namespace: bound comparisons and fresh binds in
                // this view agree, while retained-handle native checks stay physical.
                _reportedIdentity.volumeSerialNumber ^= 0xA17E5C0FFEE01234ull;
            }
        }
#endif
    }

    LocalBoundObject(const LocalBoundObject&)            = delete;
    LocalBoundObject(LocalBoundObject&&)                 = delete;
    LocalBoundObject& operator=(const LocalBoundObject&) = delete;
    LocalBoundObject& operator=(LocalBoundObject&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemBoundObject))
        {
            *ppvObject = static_cast<IFileSystemBoundObject*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileSystemBoundMetadata))
        {
            *ppvObject = static_cast<IFileSystemBoundMetadata*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileSystemBoundContentProof))
        {
            *ppvObject = static_cast<IFileSystemBoundContentProof*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(ILocalBoundObjectControl))
        {
            *ppvObject = static_cast<ILocalBoundObjectControl*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileSystemBoundObjectPinHandoff))
        {
            *ppvObject = static_cast<IFileSystemBoundObjectPinHandoff*>(this);
            AddRef();
            return S_OK;
        }
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

    HRESULT STDMETHODCALLTYPE GetSnapshot(FileSystemBoundObjectSnapshot* snapshot) noexcept override
    {
        if (snapshot == nullptr)
        {
            return E_POINTER;
        }
        if (snapshot->sizeBytes != sizeof(FileSystemBoundObjectSnapshot))
        {
            return E_INVALIDARG;
        }

        uint64_t committedSizeBytes = _committedSizeBytes;
        if (_kind == FILESYSTEM_BOUND_REGULAR_FILE && _file)
        {
            FILE_STANDARD_INFO standard{};
            if (GetFileInformationByHandleEx(_file.get(), FileStandardInfo, &standard, sizeof(standard)) != FALSE && standard.EndOfFile.QuadPart >= 0)
            {
                committedSizeBytes = static_cast<uint64_t>(standard.EndOfFile.QuadPart);
            }
        }

        snapshot->kind = static_cast<uint32_t>(_kind);
#if defined(ENABLE_TESTS)
        snapshot->objectId = &_reportedIdentity;
#else
        snapshot->objectId = &_identity;
#endif
        snapshot->objectIdBytes      = sizeof(_identity);
        snapshot->revisionId         = nullptr;
        snapshot->revisionIdBytes    = 0u;
        snapshot->committedSizeBytes = committedSizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE IsSameObject(IFileSystemBoundObject* other, BOOL* same) noexcept override
    {
        if (same == nullptr)
        {
            return E_POINTER;
        }
        *same = FALSE;
        if (other == nullptr)
        {
            return E_INVALIDARG;
        }

        FileSystemBoundObjectSnapshot otherSnapshot{};
        otherSnapshot.sizeBytes = sizeof(otherSnapshot);
        const HRESULT hr        = other->GetSnapshot(&otherSnapshot);
        if (FAILED(hr))
        {
            return hr;
        }
        if (otherSnapshot.sizeBytes != sizeof(otherSnapshot) || otherSnapshot.objectIdBytes != sizeof(_identity) || otherSnapshot.objectId == nullptr ||
            (otherSnapshot.revisionIdBytes != 0u && otherSnapshot.revisionId == nullptr))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

#if defined(ENABLE_TESTS)
        *same = std::memcmp(otherSnapshot.objectId, &_reportedIdentity, sizeof(_reportedIdentity)) == 0 ? TRUE : FALSE;
#else
        *same = std::memcmp(otherSnapshot.objectId, &_identity, sizeof(_identity)) == 0 ? TRUE : FALSE;
#endif
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OpenReader(const FileSystemOptions* options, IFileReader** reader) noexcept override
    {
        if (reader == nullptr)
        {
            return E_POINTER;
        }
        *reader = nullptr;
        if (! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_READ_CONTENT) == 0u || _kind != FILESYSTEM_BOUND_REGULAR_FILE)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        wil::unique_handle reopened;
        const HRESULT reopenHr = ReopenExactFile(_file.get(), GENERIC_READ, reopened, FILE_FLAG_OVERLAPPED);
        if (FAILED(reopenHr))
        {
            return reopenHr;
        }

        if (! IsRetainedHandleIdentity(_identity))
        {
            LocalObjectIdentityPayload reopenedIdentity{};
            const HRESULT identityHr = QueryExactLocalObjectIdentity(reopened.get(), reopenedIdentity);
            if (FAILED(identityHr))
            {
                return identityHr;
            }
            if (std::memcmp(&reopenedIdentity, &_identity, sizeof(_identity)) != 0)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
            }
        }

        LARGE_INTEGER fileSize{};
        if (GetFileSizeEx(reopened.get(), &fileSize) == FALSE || fileSize.QuadPart < 0)
        {
            const DWORD error = GetLastError();
            return fileSize.QuadPart < 0 ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }

        auto* impl = new (std::nothrow) Win32FileReader(std::move(reopened), static_cast<uint64_t>(fileSize.QuadPart), options);
        if (impl == nullptr)
        {
            return E_OUTOFMEMORY;
        }
        *reader = impl;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetContentProof(const FileSystemOptions* options, FileSystemContentProof* proof) noexcept override
    {
        if (proof == nullptr)
        {
            return E_POINTER;
        }
        if (proof->sizeBytes != sizeof(FileSystemContentProof) || ! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        *proof           = {};
        proof->sizeBytes = sizeof(FileSystemContentProof);
        if ((_grantedFlags & FILESYSTEM_BIND_READ_CONTENT) == 0u || _kind != FILESYSTEM_BOUND_REGULAR_FILE)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        wil::com_ptr<IFileReader> reader;
        HRESULT hr = OpenReader(options, reader.put());
        if (FAILED(hr) || ! reader)
        {
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        uint64_t sizeBytes = 0u;
        hr                 = reader->GetSize(&sizeBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        uint64_t position = 0u;
        hr                = reader->Seek(0, FILE_BEGIN, &position);
        if (FAILED(hr) || position != 0u)
        {
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        constexpr size_t kBufferBytes = 1024u * 1024u;
        auto buffer                   = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[kBufferBytes]);
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }
        Common::Crypto::Blake3Hasher hasher;
        uint64_t totalRead = 0u;
        for (;;)
        {
            hr = FileSystemCheckOperationControl(options);
            if (FAILED(hr))
            {
                return hr;
            }
            unsigned long bytesRead = 0u;
            hr                      = reader->Read(buffer.get(), static_cast<unsigned long>(kBufferBytes), &bytesRead);
            if (FAILED(hr))
            {
                return hr;
            }
            if (bytesRead > kBufferBytes || totalRead > std::numeric_limits<uint64_t>::max() - bytesRead)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            if (bytesRead == 0u)
            {
                break;
            }
            hasher.Update(std::span<const std::byte>(buffer.get(), bytesRead));
            totalRead += bytesRead;
        }
        if (totalRead != sizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }

        const Common::Crypto::Blake3Digest digest = hasher.Finalize();
        proof->algorithm                          = FILESYSTEM_CONTENT_PROOF_BLAKE3_256;
        proof->contentSizeBytes                   = totalRead;
        static_assert(sizeof(proof->digest) == digest.size());
        std::memcpy(proof->digest, digest.data(), digest.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBasicInformation(FileSystemBasicInformation* info) noexcept override
    {
        if (info == nullptr)
        {
            return E_POINTER;
        }
        if (info->sizeBytes != sizeof(FileSystemBasicInformation))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_READ_METADATA) == 0u)
        {
            return E_ACCESSDENIED;
        }
        if (! BeginWitnessRead())
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        const auto endRead = wil::scope_exit([&]() noexcept { EndWitnessRead(); });

        FILE_BASIC_INFO basic{};
        if (GetFileInformationByHandleEx(_file.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        info->creationTime   = basic.CreationTime.QuadPart;
        info->lastAccessTime = basic.LastAccessTime.QuadPart;
        info->lastWriteTime  = basic.LastWriteTime.QuadPart;
        info->attributes     = basic.FileAttributes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetBasicInformation(const FileSystemBasicInformation* info) noexcept override
    {
        if (info == nullptr)
        {
            return E_POINTER;
        }
        if (info->sizeBytes != sizeof(FileSystemBasicInformation))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_PUBLICATION) == 0u || ! _file)
        {
            return E_ACCESSDENIED;
        }

        FILE_BASIC_INFO basic{};
        basic.CreationTime.QuadPart   = info->creationTime;
        basic.LastAccessTime.QuadPart = info->lastAccessTime;
        basic.LastWriteTime.QuadPart  = info->lastWriteTime;
        basic.FileAttributes          = info->attributes;
        if (SetFileInformationByHandle(_file.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetMetadataSnapshot(const FileSystemOptions* options, FileSystemMetadataSnapshot* snapshot) noexcept override
    {
        if (snapshot == nullptr)
        {
            return E_POINTER;
        }
        if (snapshot->sizeBytes != sizeof(FileSystemMetadataSnapshot) || ! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_READ_METADATA) == 0u)
        {
            return E_ACCESSDENIED;
        }
        if (! BeginWitnessRead())
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        const auto endRead = wil::scope_exit([&]() noexcept { EndWitnessRead(); });
        if (! _consumableReadWitness && ! _file)
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        const HRESULT controlHr = FileSystemCheckOperationControl(options);
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        *snapshot                    = {};
        snapshot->sizeBytes          = sizeof(FileSystemMetadataSnapshot);
        snapshot->supportedFeatures  = FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES | FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS |
                                       FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES | FILESYSTEM_METADATA_SECURITY | FILESYSTEM_METADATA_SPARSE |
                                       FILESYSTEM_METADATA_COMPRESSION | FILESYSTEM_METADATA_EFS | FILESYSTEM_METADATA_PLACEHOLDER;
        snapshot->presentFeatures    = FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES | FILESYSTEM_METADATA_SECURITY;
        snapshot->logicalSizeBytes   = _committedSizeBytes;
        snapshot->allocatedSizeBytes = std::numeric_limits<uint64_t>::max();

        FILE_ATTRIBUTE_TAG_INFO tagInfo{};
        if (GetFileInformationByHandleEx(_file.get(), FileAttributeTagInfo, &tagInfo, sizeof(tagInfo)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        snapshot->fileAttributes = tagInfo.FileAttributes;
        snapshot->reparseTag     = tagInfo.ReparseTag;
        if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_SPARSE_FILE) != 0u)
        {
            snapshot->presentFeatures |= FILESYSTEM_METADATA_SPARSE;
        }
        if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_COMPRESSED) != 0u)
        {
            snapshot->presentFeatures |= FILESYSTEM_METADATA_COMPRESSION;
        }
        if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_ENCRYPTED) != 0u)
        {
            snapshot->presentFeatures |= FILESYSTEM_METADATA_EFS;
        }
        constexpr DWORD kRecallAttributes = FILE_ATTRIBUTE_RECALL_ON_OPEN | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS;
        if ((tagInfo.FileAttributes & kRecallAttributes) != 0u)
        {
            snapshot->presentFeatures |= FILESYSTEM_METADATA_PLACEHOLDER;
        }
#if defined(ENABLE_TESTS)
        snapshot->presentFeatures |= ReadSelfTestMetadataMask(kMetadataForcePresentMaskEnvVar);
#endif

        FILE_STANDARD_INFO standard{};
        if (GetFileInformationByHandleEx(_file.get(), FileStandardInfo, &standard, sizeof(standard)) != FALSE && standard.EndOfFile.QuadPart >= 0)
        {
            snapshot->logicalSizeBytes = static_cast<uint64_t>(standard.EndOfFile.QuadPart);
        }
        FILE_COMPRESSION_INFO compression{};
        if (GetFileInformationByHandleEx(_file.get(), FileCompressionInfo, &compression, sizeof(compression)) != FALSE &&
            compression.CompressedFileSize.QuadPart >= 0)
        {
            snapshot->allocatedSizeBytes = static_cast<uint64_t>(compression.CompressedFileSize.QuadPart);
        }

        uint32_t backupFeatures = snapshot->presentFeatures;
        const HRESULT streamHr  = InspectBackupMetadataStreams(_file.get(), backupFeatures);
        if (FAILED(streamHr))
        {
            // An inspection failure is not "unsupported": reporting the backup classes as
            // unsupported would let the host treat MOTW, ADS, and EA as absent and prove nothing.
            Debug::Perf::Emit(L"fileops.local.metadata.inspect_failed", L"backup-streams", 0u, snapshot->presentFeatures, 0u, streamHr);
            return streamHr;
        }
        snapshot->presentFeatures = backupFeatures;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE TransferMetadataTo(IFileSystemBoundObject* destination,
                                                 FileSystemMetadataTransferPhase phase,
                                                 const FileSystemOptions* options,
                                                 FileSystemMetadataTransferResult* result) noexcept override
    {
        if (destination == nullptr || result == nullptr)
        {
            return E_POINTER;
        }
        const uint32_t callerResultBytes = result->sizeBytes;
        if ((callerResultBytes != sizeof(FileSystemMetadataTransferResult) && callerResultBytes != FILESYSTEM_METADATA_TRANSFER_RESULT_LEGACY_SIZE) ||
            (phase != FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT && phase != FILESYSTEM_METADATA_TRANSFER_FINALIZE) ||
            ! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        // Compute into a full-size result and copy back only the caller's prefix, so a legacy-size
        // caller never receives bytes past its structure.
        FileSystemMetadataTransferResult transfer{};
        transfer.sizeBytes = sizeof(transfer);
        const auto publish = wil::scope_exit([&]() noexcept
        {
            if (callerResultBytes == FILESYSTEM_METADATA_TRANSFER_RESULT_LEGACY_SIZE)
            {
                // Legacy callers saw every security difference, and an unverifiable descriptor, as changed.
                transfer.changedFeatures |= transfer.informationalFeatures | (transfer.lostFeatures & FILESYSTEM_METADATA_SECURITY);
                transfer.lostFeatures &= ~static_cast<uint32_t>(FILESYSTEM_METADATA_SECURITY);
                transfer.informationalFeatures = 0u;
            }
            transfer.sizeBytes = callerResultBytes;
            std::memcpy(result, &transfer, callerResultBytes);
        });
        return TransferMetadataToImpl(destination, phase, options, &transfer);
    }

    HRESULT TransferMetadataToImpl(IFileSystemBoundObject* destination,
                                   FileSystemMetadataTransferPhase phase,
                                   const FileSystemOptions* options,
                                   FileSystemMetadataTransferResult* result) noexcept
    {
        if ((_grantedFlags & FILESYSTEM_BIND_READ_METADATA) == 0u)
        {
            return E_ACCESSDENIED;
        }
        if (! BeginWitnessRead())
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        const auto endRead = wil::scope_exit([&]() noexcept { EndWitnessRead(); });
        if (! _consumableReadWitness && ! _file)
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        const HRESULT controlHr = FileSystemCheckOperationControl(options);
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        *result              = {};
        result->sizeBytes    = sizeof(FileSystemMetadataTransferResult);
        result->firstFailure = S_OK;
        wil::com_ptr<ILocalBoundObjectControl> destinationControl;
        const HRESULT queryHr = destination->QueryInterface(__uuidof(ILocalBoundObjectControl), destinationControl.put_void());
        if (FAILED(queryHr) || ! destinationControl)
        {
            return FAILED(queryHr) ? queryHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        HANDLE duplicatedDestination = INVALID_HANDLE_VALUE;
        HRESULT hr                   = destinationControl->DuplicateExactHandle(&duplicatedDestination);
        if (FAILED(hr))
        {
            return hr;
        }
        wil::unique_handle destinationHandle(duplicatedDestination);

        FileSystemMetadataSnapshot sourceSnapshot{};
        sourceSnapshot.sizeBytes = sizeof(sourceSnapshot);
        hr                       = GetMetadataSnapshot(options, &sourceSnapshot);
        if (FAILED(hr))
        {
            return hr;
        }
        const auto recordFailure = [&](uint32_t feature, HRESULT failure) noexcept
        {
            result->attemptedFeatures |= feature;
            result->lostFeatures |= feature;
            if (SUCCEEDED(result->firstFailure))
            {
                result->firstFailure = failure;
            }
        };
        const auto applyControl = [&](uint32_t feature, DWORD controlCode, void* input, DWORD inputBytes) noexcept
        {
            if ((sourceSnapshot.presentFeatures & feature) == 0u)
            {
                return;
            }
            result->attemptedFeatures |= feature;
#if defined(ENABLE_TESTS)
            if ((ReadSelfTestMetadataMask(kMetadataFailMaskEnvVar) & feature) != 0u)
            {
                recordFailure(feature, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
                return;
            }
#endif
            DWORD returned = 0u;
            if (DeviceIoControl(destinationHandle.get(), controlCode, input, inputBytes, nullptr, 0u, &returned, nullptr) != FALSE)
            {
                result->preservedFeatures |= feature;
                return;
            }
            const DWORD error = GetLastError();
            recordFailure(feature, HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE));
        };

        if (phase == FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT)
        {
            FILE_SET_SPARSE_BUFFER sparse{};
            sparse.SetSparse = TRUE;
            applyControl(FILESYSTEM_METADATA_SPARSE, FSCTL_SET_SPARSE, &sparse, sizeof(sparse));

            // NTFS compression is deliberately never transferred: the destination inherits its
            // parent folder's compression state, exactly as File Explorer and CopyFileExW behave.
            // The snapshot still reports a compressed source; it is neither attempted nor lost.

            ENCRYPTION_BUFFER encryption{};
            encryption.EncryptionOperation = FILE_SET_ENCRYPTION;
            applyControl(FILESYSTEM_METADATA_EFS, FSCTL_SET_ENCRYPTION, &encryption, sizeof(encryption));
            return S_OK;
        }

        HRESULT backupHr = S_OK;
        // I35 R10.3 candidate (d): the snapshot's inspection pass just read the source's streams. When
        // it found no named stream, MOTW or extended attribute there is nothing to copy, and the pass
        // would only reopen both objects (a round trip on a network destination).
        const bool skipAbsentBackupStreams = (sourceSnapshot.presentFeatures & kMetadataBackupFeatures) == 0u && options != nullptr &&
                                             (options->transferTuning & FILESYSTEM_TRANSFER_TUNING_SKIP_ABSENT_BACKUP_STREAMS) != 0u;
#if defined(ENABLE_TESTS)
        const uint32_t forcedBackupLoss = ReadSelfTestMetadataMask(kMetadataFailMaskEnvVar) & sourceSnapshot.presentFeatures & kMetadataBackupFeatures;
        if (forcedBackupLoss != 0u)
        {
            result->attemptedFeatures |= forcedBackupLoss;
            result->lostFeatures |= forcedBackupLoss;
            result->firstFailure = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        else
#endif
            if (! skipAbsentBackupStreams)
        {
            backupHr = CopyBackupMetadataStreams(_file.get(), destinationHandle.get(), *result);
        }
        if (FAILED(backupHr))
        {
            const uint32_t presentBackup = sourceSnapshot.presentFeatures & kMetadataBackupFeatures;
            result->attemptedFeatures |= presentBackup;
            result->lostFeatures |= presentBackup;
            if (SUCCEEDED(result->firstFailure))
            {
                result->firstFailure = backupHr;
            }
        }

        // BackupWrite may update the archive/basic state while materializing named streams. Apply
        // ordinary attributes and timestamps after it, while preserving provider-managed flags
        // established during the pre-content phase.
        FileSystemBasicInformation basic{};
        basic.sizeBytes = sizeof(basic);
        result->attemptedFeatures |= FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES;
        hr = GetBasicInformation(&basic);
        if (SUCCEEDED(hr))
        {
            FileSystemBasicInformation destinationBasic{};
            destinationBasic.sizeBytes = sizeof(destinationBasic);
            if (SUCCEEDED(destination->GetBasicInformation(&destinationBasic)))
            {
                constexpr DWORD kProviderManagedAttributes = FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_ENCRYPTED | FILE_ATTRIBUTE_SPARSE_FILE |
                                                             FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_RECALL_ON_OPEN |
                                                             FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS;
                basic.attributes = (basic.attributes & ~kProviderManagedAttributes) | (destinationBasic.attributes & kProviderManagedAttributes);
            }
            hr = destination->SetBasicInformation(&basic);
        }
        if (SUCCEEDED(hr))
        {
            result->preservedFeatures |= FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES;
        }
        else
        {
            recordFailure(FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES, hr);
        }

        // Destination inheritance is the product contract for ACL/owner. Report a change (or an
        // inability to compare) instead of silently claiming that security metadata was cloned.
        result->attemptedFeatures |= FILESYSTEM_METADATA_SECURITY;
        wil::unique_handle sourceSecurity;
        wil::unique_handle destinationSecurity;
        const HRESULT sourceSecurityHr      = ReopenExactFile(_file.get(), READ_CONTROL, sourceSecurity);
        const HRESULT destinationSecurityHr = ReopenExactFile(destinationHandle.get(), READ_CONTROL, destinationSecurity);
        if (SUCCEEDED(sourceSecurityHr) && SUCCEEDED(destinationSecurityHr))
        {
            DWORD sourceBytes      = 0u;
            DWORD destinationBytes = 0u;
            static_cast<void>(GetKernelObjectSecurity(
                sourceSecurity.get(), OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION, nullptr, 0u, &sourceBytes));
            static_cast<void>(GetKernelObjectSecurity(destinationSecurity.get(),
                                                      OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION,
                                                      nullptr,
                                                      0u,
                                                      &destinationBytes));
            std::vector<std::byte> sourceDescriptor(sourceBytes);
            std::vector<std::byte> destinationDescriptor(destinationBytes);
            const bool sourceRead =
                sourceBytes != 0u && GetKernelObjectSecurity(sourceSecurity.get(),
                                                             OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION,
                                                             reinterpret_cast<PSECURITY_DESCRIPTOR>(sourceDescriptor.data()),
                                                             sourceBytes,
                                                             &sourceBytes) != FALSE;
            const bool destinationRead =
                destinationBytes != 0u && GetKernelObjectSecurity(destinationSecurity.get(),
                                                                  OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION,
                                                                  reinterpret_cast<PSECURITY_DESCRIPTOR>(destinationDescriptor.data()),
                                                                  destinationBytes,
                                                                  &destinationBytes) != FALSE;
#if defined(ENABLE_TESTS)
            const bool forcedSecurityFailure = (ReadSelfTestMetadataMask(kMetadataFailMaskEnvVar) & FILESYSTEM_METADATA_SECURITY) != 0u;
#else
            constexpr bool forcedSecurityFailure = false;
#endif
            if (! forcedSecurityFailure && sourceRead && destinationRead && sourceDescriptor == destinationDescriptor)
            {
                result->preservedFeatures |= FILESYSTEM_METADATA_SECURITY;
            }
            else if (! forcedSecurityFailure && sourceRead && destinationRead)
            {
                // I35 R5.4 / D-5.1: the destination's inheritance decides ACEs, owner and primary group,
                // so those differences are expected (Information). Only a protected source DACL that
                // the destination replaced by an inherited one is a deviation the user may act on.
                SECURITY_DESCRIPTOR_CONTROL sourceSdControl      = 0u;
                SECURITY_DESCRIPTOR_CONTROL destinationSdControl = 0u;
                DWORD sourceRevision                             = 0u;
                DWORD destinationRevision                        = 0u;
                const bool controlsRead =
                    GetSecurityDescriptorControl(reinterpret_cast<PSECURITY_DESCRIPTOR>(sourceDescriptor.data()), &sourceSdControl, &sourceRevision) != FALSE &&
                    GetSecurityDescriptorControl(
                        reinterpret_cast<PSECURITY_DESCRIPTOR>(destinationDescriptor.data()), &destinationSdControl, &destinationRevision) != FALSE;
                const bool protectedDaclReplaced =
                    ! controlsRead || ((sourceSdControl & SE_DACL_PROTECTED) != 0u && (destinationSdControl & SE_DACL_PROTECTED) == 0u);
                if (protectedDaclReplaced)
                {
                    result->changedFeatures |= FILESYSTEM_METADATA_SECURITY;
                }
                else
                {
                    result->informationalFeatures |= FILESYSTEM_METADATA_SECURITY;
                }
            }
            else
            {
                // An unreadable descriptor cannot be compared: the outcome is unverifiable (lost).
                result->lostFeatures |= FILESYSTEM_METADATA_SECURITY;
                if (SUCCEEDED(result->firstFailure))
                {
                    result->firstFailure = forcedSecurityFailure ? HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) : HRESULT_FROM_WIN32(ERROR_INVALID_SECURITY_DESCR);
                }
            }
        }
        else
        {
            result->lostFeatures |= FILESYSTEM_METADATA_SECURITY;
            if (SUCCEEDED(result->firstFailure))
            {
                result->firstFailure = FAILED(sourceSecurityHr) ? sourceSecurityHr : destinationSecurityHr;
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PublishAs(const wchar_t* finalPath,
                                        IFileSystemBoundObject* expectedDestination,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        FileSystemConditionalMutationResult* result,
                                        IFileSystemBoundObject** published) noexcept override
    {
        return ConditionalRename(finalPath, expectedDestination, flags, options, FILESYSTEM_BIND_PUBLICATION, true, result, published);
    }

    HRESULT STDMETHODCALLTYPE RenameIfUnchanged(const wchar_t* destinationPath,
                                                IFileSystemBoundObject* expectedDestination,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                FileSystemConditionalMutationResult* result,
                                                IFileSystemBoundObject** renamed) noexcept override
    {
        return ConditionalRename(destinationPath, expectedDestination, flags, options, FILESYSTEM_BIND_RENAME, false, result, renamed);
    }

    HRESULT STDMETHODCALLTYPE DeleteIfUnchanged(FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                FileSystemConditionalMutationResult* result) noexcept override
    {
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        constexpr uint32_t allowedFlags = FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY;
        if ((static_cast<uint32_t>(flags) & ~allowedFlags) != 0u)
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_DELETE) == 0u || _ownedStage)
        {
            return E_ACCESSDENIED;
        }

        const HRESULT controlHr = FileSystemCheckOperationControl(options);
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        FILE_BASIC_INFO rootBasic{};
        if (GetFileInformationByHandleEx(_file.get(), FileBasicInfo, &rootBasic, sizeof(rootBasic)) == FALSE)
            return HRESULT_FROM_WIN32(GetLastError());
        const bool directoryLink = _kind == FILESYSTEM_BOUND_LINK && (rootBasic.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
        const bool completionKnown =
            _kind != FILESYSTEM_BOUND_OTHER && (_kind != FILESYSTEM_BOUND_REGULAR_FILE || _committedSizeBytes != std::numeric_limits<uint64_t>::max());
        const auto reportWork = [&](FileSystemFileWorkPhase phase) noexcept
        {
            if (options == nullptr || options->fileWorkCallback == nullptr)
                return;
            FileSystemFileWorkProgress report{};
            report.sizeBytes       = sizeof(report);
            report.phase           = phase;
            report.kind            = FileSystemFileWorkKind::Directory;
            report.scopeBytesKnown = TRUE;
            if (! completionKnown)
            {
                report.phase = FileSystemFileWorkPhase::CountsUnavailable;
            }
            if (phase == FileSystemFileWorkPhase::ObjectCompleted)
            {
                report.kind = _kind == FILESYSTEM_BOUND_DIRECTORY ? FileSystemFileWorkKind::Directory
                              : _kind == FILESYSTEM_BOUND_LINK    ? (directoryLink ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink)
                                                                  : FileSystemFileWorkKind::RegularFile;
                report.sourcePath  = _path.c_str();
                report.disposition = FileSystemFileWorkDisposition::Completed;
                report.sizeKnown   = TRUE;
                report.totalBytes =
                    _kind == FILESYSTEM_BOUND_REGULAR_FILE && _committedSizeBytes != std::numeric_limits<uint64_t>::max() ? _committedSizeBytes : 0u;
                report.completedBytes = report.totalBytes;
            }
            else if (phase == FileSystemFileWorkPhase::ReadOnlyDeleteBlocked)
            {
                report.kind =
                    (rootBasic.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FileSystemFileWorkKind::Directory : FileSystemFileWorkKind::RegularFile;
                report.sourcePath  = _path.c_str();
                report.disposition = FileSystemFileWorkDisposition::Failed;
                report.status      = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
            static_cast<void>(options->fileWorkCallback->FileSystemReportFileWork(&report, options->fileWorkCookie));
        };
        reportWork(FileSystemFileWorkPhase::ScopeSupported);
        if ((rootBasic.FileAttributes & FILE_ATTRIBUTE_READONLY) != 0u && (flags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) == 0u)
        {
            // The exact cleanup primitive may ignore readonly for its already-authorized internal
            // compensation. Public conditional deletion needs a separate explicit readonly grant,
            // including selected files and empty directories before any descendants are touched.
            reportWork(FileSystemFileWorkPhase::ReadOnlyDeleteBlocked);
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        if (_kind != FILESYSTEM_BOUND_DIRECTORY && options != nullptr && options->operationControl != nullptr)
        {
            FileSystemDiscoveryProgress discovery{};
            discovery.sizeBytes             = sizeof(discovery);
            discovery.discoveredBytes       = _committedSizeBytes == std::numeric_limits<uint64_t>::max() ? 0u : _committedSizeBytes;
            discovery.discoveredFiles       = 1u;
            discovery.discoveredDirectories = 0u;
            discovery.queuedItems           = 1u;
            discovery.traversalClosed       = TRUE;
            const HRESULT discoveryHr       = options->operationControl->FileSystemReportDiscoveryProgress(&discovery, options->operationControlCookie);
            if (FAILED(discoveryHr))
            {
                return discoveryHr;
            }
        }

        if (_kind == FILESYSTEM_BOUND_DIRECTORY && (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE)) != 0u)
        {
            const HRESULT contentsHr = FileSystemInternal::DeleteBoundLocalDirectoryContents(_path.c_str(), flags, options);
            if (FAILED(contentsHr))
            {
                // The recursive walk mutates descendants only. Because this retained exact root
                // handle has not yet received a delete disposition, a failed/partial child walk is
                // known root non-commit with the selected directory still present.
                return contentsHr;
            }
        }

#if defined(ENABLE_TESTS)
        // Deliberately contradictory provider receipt: retain the exact object and
        // its handle so the host must reject committed-with-original-present truth.
        if (_kind != FILESYSTEM_BOUND_DIRECTORY &&
            ConsumeMatchingPathInjectionForSelfTest(
                _path.c_str(), kDeleteCommittedRetainedPathEnvVar, kDeleteCommittedRetainedFiredEnvVar, L"FileOps.Local.DeleteCommittedRetainedInjected"))
        {
            result->outcomeKnown         = TRUE;
            result->mutationCommitted    = TRUE;
            result->originalStillPresent = TRUE;
            return S_OK;
        }
#endif
        const HRESULT deleteHr = DeleteExact(result);
        if (deleteHr == S_OK && result->outcomeKnown == TRUE && result->mutationCommitted == TRUE && result->originalStillPresent == FALSE)
            reportWork(FileSystemFileWorkPhase::ObjectCompleted);
        return deleteHr;
    }

    HRESULT STDMETHODCALLTYPE AbortOwnedObject(const FileSystemOptions* options, FileSystemConditionalMutationResult* result) noexcept override
    {
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        if (! _ownedStage)
        {
            return E_ACCESSDENIED;
        }

        const HRESULT controlHr = FileSystemCheckOperationControl(options);
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        // I35 R7.4: every caller (the host bridge and this provider's own stages) removes the owned
        // object under a short cleanup control; the cancel watch leaves that I/O alone until its deadline.
        const Common::SynchronousIoCancelWatch::CleanupWindow cleanupWindow(OwnedStageCleanupWindowDeadline(options));
        HRESULT abortHr = DeleteExact(result);
        if (cleanupWindow.CanceledInsideWindow(abortHr))
        {
            abortHr = DeleteExact(result);
        }
        return abortHr;
    }

    // Level 2 of the authority ladder: the location opened the occupant and cannot identify it. The
    // occupant's facts are read again here and one replacing rename publishes the stage; a difference
    // publishes nothing. No handle was held on the occupant and no copy of it is set aside.
    HRESULT STDMETHODCALLTYPE PublishOwnedOverExpected(const wchar_t* finalPath,
                                                       const FileSystemExpectedOccupant* expected,
                                                       FileSystemFlags flags,
                                                       const FileSystemOptions* options,
                                                       FileSystemConditionalMutationResult* result,
                                                       IFileSystemBoundObject** published) noexcept override
    {
        if (published == nullptr)
        {
            return E_POINTER;
        }
        *published             = nullptr;
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (finalPath == nullptr || finalPath[0] == L'\0' || expected == nullptr || expected->sizeBytes != sizeof(FileSystemExpectedOccupant) ||
            ! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & FILESYSTEM_BIND_PUBLICATION) == 0u || ! _ownedStage || ! _file || _kind != FILESYSTEM_BOUND_REGULAR_FILE)
        {
            return E_ACCESSDENIED;
        }
        constexpr uint32_t allowedFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY;
        const uint32_t requestedFlags   = static_cast<uint32_t>(flags);
        if ((requestedFlags & ~allowedFlags) != 0u || (requestedFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0u)
        {
            return E_INVALIDARG;
        }
        HRESULT hr = FileSystemCheckOperationControl(options);
        if (FAILED(hr))
        {
            return hr;
        }

        FileSystemExpectedOccupant current{};
        bool identityAvailable = false;
        hr                     = ReadLocalOccupantFacts(finalPath, current, identityAvailable);
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        if (FAILED(hr))
        {
            return hr;
        }
        // A location that now identifies the item is no longer the one the decision was made for.
        if (identityAvailable || ! LocalOccupantFactsEqual(current, *expected))
        {
            Debug::Perf::Emit(L"FileOps.Local.ExpectedOccupantChanged", finalPath, 0u, 0u, 0u, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        const bool readOnly = (current.attributes & FILE_ATTRIBUTE_READONLY) != 0u;
        if (readOnly && (requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        // A replacing rename does not pass a read-only occupant. Clear the attribute by name for the
        // rename only, and put it back when the rename does not commit.
        constexpr unsigned long kSettableAttributes = FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM | FILE_ATTRIBUTE_ARCHIVE |
                                                      FILE_ATTRIBUTE_NOT_CONTENT_INDEXED | FILE_ATTRIBUTE_OFFLINE;
        const std::wstring extendedFinal            = FileSystemInternal::ToExtendedPath(finalPath);
        bool readOnlyCleared                        = false;
        const auto restoreReadOnly                  = wil::scope_exit([&]() noexcept
        {
            if (readOnlyCleared)
            {
                static_cast<void>(SetFileAttributesW(extendedFinal.c_str(), current.attributes & kSettableAttributes));
            }
        });
        if (readOnly)
        {
            const unsigned long writable = current.attributes & kSettableAttributes & ~static_cast<unsigned long>(FILE_ATTRIBUTE_READONLY);
            if (SetFileAttributesW(extendedFinal.c_str(), writable != 0u ? writable : FILE_ATTRIBUTE_NORMAL) == FALSE)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_ACCESS_DENIED);
            }
            readOnlyCleared = true;
        }

        hr = ClearTemporaryAttributes();
        if (SUCCEEDED(hr))
        {
            hr = FileSystemCheckOperationControl(options);
        }
        if (FAILED(hr))
        {
            return hr;
        }

        FileSystemConditionalMutationResult publishResult{};
        publishResult.sizeBytes = sizeof(publishResult);
        const HRESULT publishHr = RenameExactTo(finalPath, TRUE, &publishResult);
        if (publishResult.outcomeKnown == FALSE && FAILED(publishHr) && OwnedStageStillAtItsName())
        {
            // A stage without a persistent identity cannot be reconciled by identity, and a refused
            // replacing rename (the occupant is in use) is the ordinary failure here. The stage is
            // still under the name only this operation produces: nothing was replaced.
            return publishHr;
        }
        if (publishResult.outcomeKnown == FALSE)
        {
            // The occupant may already be replaced: its attribute is not written back.
            readOnlyCleared      = false;
            result->outcomeKnown = FALSE;
            return FAILED(publishHr) ? publishHr : HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (FAILED(publishHr) || publishResult.mutationCommitted == FALSE)
        {
            return FAILED(publishHr) ? publishHr : E_UNEXPECTED;
        }

        readOnlyCleared              = false;
        _ownedStage                  = false;
        result->mutationCommitted    = TRUE;
        result->originalStillPresent = FALSE;
        *published                   = static_cast<IFileSystemBoundObject*>(this);
        AddRef();
        Debug::Perf::EmitCounter(L"FileOps.Local.ExpectedOccupantReplaced");
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RenameExactTo(const wchar_t* destinationPath, BOOL replaceIfExists, FileSystemConditionalMutationResult* result) noexcept override
    {
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (destinationPath == nullptr || destinationPath[0] == L'\0')
        {
            return E_INVALIDARG;
        }
        if (_consumableReadWitness)
            return E_ACCESSDENIED;
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        const std::wstring extendedDestination = FileSystemInternal::ToExtendedPath(destinationPath);
        const size_t fileNameBytes             = extendedDestination.size() * sizeof(wchar_t);
        if (fileNameBytes > (std::numeric_limits<DWORD>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        }
        // FILE_RENAME_INFO includes one WCHAR of trailing storage. Allocate the complete header plus
        // the non-NUL name so older kernels never read beyond a merely offsetof-sized buffer.
        const size_t bufferBytes = sizeof(FILE_RENAME_INFO) + fileNameBytes;
        auto buffer              = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[bufferBytes]);
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }
        std::memset(buffer.get(), 0, bufferBytes);
        auto* renameInfo           = reinterpret_cast<FILE_RENAME_INFO*>(buffer.get());
        renameInfo->Flags          = replaceIfExists != FALSE ? FILE_RENAME_FLAG_REPLACE_IF_EXISTS : 0u;
        renameInfo->RootDirectory  = nullptr;
        renameInfo->FileNameLength = static_cast<DWORD>(fileNameBytes);
        std::memcpy(renameInfo->FileName, extendedDestination.data(), fileNameBytes);

        bool renamed = SetFileInformationByHandle(_file.get(), FileRenameInfoEx, renameInfo, static_cast<DWORD>(bufferBytes)) != FALSE;
        DWORD error  = renamed ? ERROR_SUCCESS : GetLastError();
        if (! renamed)
        {
            if (error == ERROR_INVALID_PARAMETER || error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED)
            {
                renameInfo->ReplaceIfExists = replaceIfExists != FALSE ? static_cast<BOOLEAN>(TRUE) : static_cast<BOOLEAN>(FALSE);
                if (SetFileInformationByHandle(_file.get(), FileRenameInfo, renameInfo, static_cast<DWORD>(bufferBytes)) != FALSE)
                {
                    renamed = true;
                }
                else
                {
                    error = GetLastError();
                }
            }
        }

#if defined(ENABLE_TESTS)
        // Inject after either rename API succeeds; SMB servers can require the legacy fallback.
        if (renamed &&
            ConsumeMatchingPathInjectionForSelfTest(
                destinationPath, kRenameCommittedFailurePathEnvVar, kRenameCommittedFailureFiredEnvVar, L"FileOps.Local.RenameCommittedFailureInjected"))
        {
            renamed = false;
            error   = ERROR_IO_INCOMPLETE;
        }
#endif
        if (! renamed)
        {
            bool sameAtDestination    = false;
            const HRESULT reconcileHr = PathRefersToThisObject(extendedDestination, sameAtDestination);
            if (SUCCEEDED(reconcileHr) && sameAtDestination)
            {
                _path                        = destinationPath;
                result->mutationCommitted    = TRUE;
                result->originalStillPresent = FALSE;
                return S_OK;
            }

            bool sameAtOriginal      = false;
            const HRESULT originalHr = IsRetainedHandleIdentity(_identity) ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                                           : PathRefersToThisObject(FileSystemInternal::ToExtendedPath(_path), sameAtOriginal);
            if (SUCCEEDED(originalHr) && sameAtOriginal)
            {
                return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
            }

            result->outcomeKnown = FALSE;
            Debug::Perf::EmitCounter(L"fileops.local.mutation.rename_reconcile_unknown");
            Debug::Perf::Emit(L"FileOps.Local.RenameReconcileUnknown", _path, 0u, error, static_cast<uint64_t>(_kind), HRESULT_FROM_WIN32(error));
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }

        _path                        = destinationPath;
        result->mutationCommitted    = TRUE;
        result->originalStillPresent = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DeleteExact(FileSystemConditionalMutationResult* result) noexcept override
    {
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (_consumableReadWitness)
            return E_ACCESSDENIED;
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

#if defined(ENABLE_TESTS)
        if (_ownedStage)
        {
            // Nothing was attempted yet: an interrupted delay is a known non-commit.
            const HRESULT delayHr = DelayOwnedStageCleanupForSelfTest(_path);
            if (FAILED(delayHr))
            {
                return delayHr;
            }
        }
#endif
        HRESULT deleteHr = DeleteExactHandle(_file.get());
#if defined(ENABLE_TESTS)
        if (SUCCEEDED(deleteHr) &&
            ConsumeMatchingPathInjectionForSelfTest(
                _path.c_str(), kDeleteCommittedFailurePathEnvVar, kDeleteCommittedFailureFiredEnvVar, L"FileOps.Local.DeleteCommittedFailureInjected"))
        {
            deleteHr = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
#endif
        if (FAILED(deleteHr))
        {
            if (deleteHr == HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY))
            {
                // The kernel rejected the exact directory disposition because entries remain.
                // No delete disposition was accepted, so this is a proven non-commit even though
                // an earlier recursive child walk may have changed directory membership.
                return deleteHr;
            }
            result->outcomeKnown = FALSE;
            Debug::Perf::EmitCounter(L"fileops.local.mutation.delete_outcome_unknown");
            return deleteHr;
        }

        _file.reset();
        _ownedStage                  = false;
        result->mutationCommitted    = TRUE;
        result->originalStillPresent = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DuplicateExactHandle(HANDLE* duplicated) noexcept override
    {
        if (duplicated == nullptr)
        {
            return E_POINTER;
        }
        *duplicated = INVALID_HANDLE_VALUE;
        if (_consumableReadWitness)
            return E_ACCESSDENIED;
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        HANDLE result = INVALID_HANDLE_VALUE;
        if (DuplicateHandle(GetCurrentProcess(), _file.get(), GetCurrentProcess(), &result, 0u, FALSE, DUPLICATE_SAME_ACCESS) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_INVALID_HANDLE);
        }
        *duplicated = result;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RetireDirectoryReadPin(const LocalObjectIdentityPayload* expected) noexcept override
    {
        if (! _consumableReadWitness || expected == nullptr || IsRetainedHandleIdentity(_identity) || std::memcmp(expected, &_identity, sizeof(_identity)) != 0)
            return E_ACCESSDENIED;
        std::scoped_lock lock(_witnessMutex);
        if (! _file)
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        if (_witnessReaders != 0u)
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        _file.reset();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE TakeDirectoryReadPinFrom(IFileSystemBoundObject* witness) noexcept override
    {
        constexpr FileSystemBindFlags strictDirectoryDeleteFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE);
        if (! witness || _consumableReadWitness || _ownedStage || _kind != FILESYSTEM_BOUND_DIRECTORY || _grantedFlags != strictDirectoryDeleteFlags ||
            ! _file || IsRetainedHandleIdentity(_identity))
            return E_ACCESSDENIED;
        FileSystemBasicInformation basic{};
        basic.sizeBytes       = sizeof(basic);
        const HRESULT basicHr = GetBasicInformation(&basic);
        if (FAILED(basicHr))
            return basicHr;
        if ((basic.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (basic.attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        BOOL same            = FALSE;
        const HRESULT sameHr = IsSameObject(witness, &same);
        if (FAILED(sameHr))
            return sameHr;
        if (same == FALSE)
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        wil::com_ptr<ILocalBoundObjectControl> control;
        const HRESULT queryHr = witness->QueryInterface(__uuidof(ILocalBoundObjectControl), control.put_void());
        if (FAILED(queryHr) || ! control)
            return FAILED(queryHr) ? queryHr : E_UNEXPECTED;
        return control->RetireDirectoryReadPin(&_identity);
    }

private:
    ~LocalBoundObject() = default;

    // Only directory metadata witnesses can surrender their kernel pin. Track their
    // synchronous reads separately from immutable COM snapshots; a busy handoff fails
    // rather than waiting while a reader might call another bound object.
    [[nodiscard]] bool BeginWitnessRead() noexcept
    {
        if (! _consumableReadWitness)
            return true;
        std::scoped_lock lock(_witnessMutex);
        if (! _file)
            return false;
        ++_witnessReaders;
        return true;
    }

    void EndWitnessRead() noexcept
    {
        if (_consumableReadWitness)
        {
            std::scoped_lock lock(_witnessMutex);
            --_witnessReaders;
        }
    }

    [[nodiscard]] static HRESULT InitializeMutationResult(FileSystemConditionalMutationResult* result) noexcept
    {
        if (result == nullptr)
        {
            return E_POINTER;
        }
        if (result->sizeBytes != sizeof(FileSystemConditionalMutationResult))
        {
            return E_INVALIDARG;
        }
        result->mutationCommitted    = FALSE;
        result->originalStillPresent = TRUE;
        result->outcomeKnown         = TRUE;
        return S_OK;
    }

    [[nodiscard]] HRESULT PathRefersToThisObject(const std::wstring& extendedPath, bool& same) const noexcept
    {
        same = false;
        if (IsRetainedHandleIdentity(_identity))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        wil::unique_handle current(CreateFileW(extendedPath.c_str(),
                                               FILE_READ_ATTRIBUTES,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               nullptr,
                                               OPEN_EXISTING,
                                               FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                               nullptr));
        if (! current)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_NOT_FOUND)
            {
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }

        LocalObjectIdentityPayload currentIdentity{};
        const HRESULT identityHr = QueryExactLocalObjectIdentity(current.get(), currentIdentity);
        if (FAILED(identityHr))
        {
            return identityHr;
        }
        same = std::memcmp(&currentIdentity, &_identity, sizeof(_identity)) == 0;
        return S_OK;
    }

    // An owned stage carries a name no other writer produces. Opening that name (an open is never
    // answered from a client cache) after a rename the location could not reconcile by identity
    // proves the stage did not move.
    [[nodiscard]] bool OwnedStageStillAtItsName() const noexcept
    {
        if (! _ownedStage || _path.empty())
        {
            return false;
        }
        const wil::unique_handle current(CreateFileW(FileSystemInternal::ToExtendedPath(_path).c_str(),
                                                     FILE_READ_ATTRIBUTES,
                                                     FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                     nullptr,
                                                     OPEN_EXISTING,
                                                     FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                     nullptr));
        return current.is_valid();
    }

    [[nodiscard]] static HRESULT MakeBackupPath(std::wstring_view finalPath, std::wstring& backupPath) noexcept
    {
        return FileSystemInternal::MakeStageSiblingPath(finalPath, FileSystemInternal::kBackupExtension, backupPath);
    }

    [[nodiscard]] HRESULT ClearTemporaryAttributes() noexcept
    {
        if (! _file || ! _ownedStage)
        {
            return S_OK;
        }
        FILE_BASIC_INFO basic{};
        if (GetFileInformationByHandleEx(_file.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        // Source metadata has already replaced the stage attributes. Hidden is user metadata;
        // only the writer's temporary caching hint must be removed before publication.
        basic.FileAttributes &= ~FILE_ATTRIBUTE_TEMPORARY;
        if (basic.FileAttributes == 0u)
        {
            basic.FileAttributes = FILE_ATTRIBUTE_NORMAL;
        }
        if (SetFileInformationByHandle(_file.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT ConditionalRename(const wchar_t* destinationPath,
                                            IFileSystemBoundObject* expectedDestination,
                                            FileSystemFlags flags,
                                            const FileSystemOptions* options,
                                            FileSystemBindFlags requiredGrant,
                                            bool requireOwnedStage,
                                            FileSystemConditionalMutationResult* result,
                                            IFileSystemBoundObject** renamed) noexcept
    {
        if (renamed == nullptr)
        {
            return E_POINTER;
        }
        *renamed               = nullptr;
        const HRESULT resultHr = InitializeMutationResult(result);
        if (FAILED(resultHr))
        {
            return resultHr;
        }
        if (destinationPath == nullptr || destinationPath[0] == L'\0')
        {
            return E_INVALIDARG;
        }
        if (! FileSystemOptionsHaveValidHeader(options))
        {
            return E_INVALIDARG;
        }
        if ((_grantedFlags & requiredGrant) == 0u || requireOwnedStage != _ownedStage || ! _file)
        {
            return E_ACCESSDENIED;
        }

        constexpr uint32_t allowedFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK;
        const uint32_t requestedFlags   = static_cast<uint32_t>(flags);
        if ((requestedFlags & ~allowedFlags) != 0u ||
            ((requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) != 0u && (requestedFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0u) ||
            ((requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_LINK) != 0u && (requestedFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0u) ||
            (expectedDestination != nullptr && (requestedFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0u))
        {
            return E_INVALIDARG;
        }
        const HRESULT controlHr = FileSystemCheckOperationControl(options);
        if (FAILED(controlHr))
        {
            return controlHr;
        }
#if defined(ENABLE_TESTS)
        // Exercise the exact owned-stage publication boundary before any destination mutation.
        if (requireOwnedStage &&
            ShouldFailOwnedPublicationForSelfTest(
                destinationPath, kStagedCopyPromoteFailPathEnvVar, kStagedCopyPromoteFailFiredEnvVar, L"FileOps.Copy.DebugStagedPromoteFailureInjected"))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
#endif

        wil::com_ptr<ILocalBoundObjectControl> expectedControl;
        bool destinationMovedToBackup = false;
        bool expectedReadOnlyCleared  = false;
        FileSystemBasicInformation expectedInfo{};
        expectedInfo.sizeBytes             = sizeof(expectedInfo);
        const auto restoreExpectedReadOnly = wil::scope_exit([&]() noexcept
        {
            if (expectedReadOnlyCleared && expectedDestination != nullptr)
            {
                static_cast<void>(expectedDestination->SetBasicInformation(&expectedInfo));
            }
        });
        std::wstring backupPath;
        if (expectedDestination != nullptr)
        {
            FileSystemBoundObjectSnapshot expectedSnapshot{};
            expectedSnapshot.sizeBytes = sizeof(expectedSnapshot);
            HRESULT hr                 = expectedDestination->GetSnapshot(&expectedSnapshot);
            if (FAILED(hr))
            {
                return hr;
            }
            const bool expectedIsLink     = expectedSnapshot.kind == FILESYSTEM_BOUND_LINK;
            const bool replaceLinkGranted = (requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_LINK) != 0u;
            if (expectedIsLink != replaceLinkGranted)
            {
                return HRESULT_FROM_WIN32(expectedIsLink ? ERROR_REPARSE_POINT_ENCOUNTERED : ERROR_INVALID_PARAMETER);
            }
            if (! expectedIsLink && expectedSnapshot.kind != static_cast<uint32_t>(_kind))
            {
                return HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH);
            }
            if (expectedSnapshot.objectId == nullptr || expectedSnapshot.objectIdBytes != sizeof(LocalObjectIdentityPayload))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            const std::wstring extendedDestination = FileSystemInternal::ToExtendedPath(destinationPath);
            wil::unique_handle currentDestination(CreateFileW(extendedDestination.c_str(),
                                                              FILE_READ_ATTRIBUTES,
                                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                              nullptr,
                                                              OPEN_EXISTING,
                                                              FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                              nullptr));
            if (! currentDestination)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND);
            }
            LocalObjectIdentityPayload currentIdentity{};
            hr = QueryExactLocalObjectIdentity(currentDestination.get(), currentIdentity);
            if (FAILED(hr))
            {
                return hr;
            }
            if (std::memcmp(expectedSnapshot.objectId, &currentIdentity, sizeof(currentIdentity)) != 0)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
            }
            hr = expectedDestination->GetBasicInformation(&expectedInfo);
            if (FAILED(hr))
            {
                return hr;
            }
            if ((expectedInfo.attributes & FILE_ATTRIBUTE_READONLY) != 0u && (requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
            if ((expectedInfo.attributes & FILE_ATTRIBUTE_READONLY) != 0u)
            {
                FileSystemBasicInformation writableInfo = expectedInfo;
                writableInfo.attributes &= ~FILE_ATTRIBUTE_READONLY;
                hr = expectedDestination->SetBasicInformation(&writableInfo);
                if (FAILED(hr))
                {
                    return hr;
                }
                expectedReadOnlyCleared = true;
            }
            hr = expectedDestination->QueryInterface(__uuidof(ILocalBoundObjectControl), expectedControl.put_void());
            if (FAILED(hr) || ! expectedControl)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            constexpr unsigned int kMaxBackupAttempts = 32u;
            for (unsigned int attempt = 0u; attempt < kMaxBackupAttempts; ++attempt)
            {
                hr = MakeBackupPath(destinationPath, backupPath);
                if (FAILED(hr))
                {
                    return hr;
                }
                FileSystemConditionalMutationResult backupResult{};
                backupResult.sizeBytes = sizeof(backupResult);
                hr                     = expectedControl->RenameExactTo(backupPath.c_str(), FALSE, &backupResult);
                if (SUCCEEDED(hr) && backupResult.outcomeKnown != FALSE && backupResult.mutationCommitted != FALSE)
                {
                    destinationMovedToBackup = true;
                    break;
                }
                if (backupResult.outcomeKnown == FALSE)
                {
                    result->outcomeKnown = FALSE;
                    return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
                }
                if (hr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && hr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
                {
                    return FAILED(hr) ? hr : E_UNEXPECTED;
                }
            }
            if (! destinationMovedToBackup)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            }
        }

        HRESULT attributesHr = ClearTemporaryAttributes();
        if (SUCCEEDED(attributesHr))
        {
            // SMB metadata preparation can outlive the cancel watch's grace. Recheck at the
            // last reversible boundary instead of starting a rename after Cancel already won.
            attributesHr = FileSystemCheckOperationControl(options);
            if (FAILED(attributesHr))
            {
                Debug::Perf::Emit(L"FileOps.Local.PublicationCanceledBeforeRename", destinationPath, 0u, static_cast<uint64_t>(_kind), 0u, attributesHr);
            }
        }
        if (FAILED(attributesHr))
        {
            if (destinationMovedToBackup)
            {
                // I35 R7.4: restoring the user's original from its backup has the cleanup budget; a
                // canceled rollback would leave the original under its backup name.
                const Common::SynchronousIoCancelWatch::CleanupWindow rollbackWindow(FreshOwnedStageCleanupDeadline());
                FileSystemConditionalMutationResult rollbackResult{};
                rollbackResult.sizeBytes = sizeof(rollbackResult);
                HRESULT rollbackHr       = expectedControl->RenameExactTo(destinationPath, FALSE, &rollbackResult);
                if (rollbackWindow.CanceledInsideWindow(rollbackHr))
                {
                    rollbackHr = expectedControl->RenameExactTo(destinationPath, FALSE, &rollbackResult);
                }
                if (FAILED(rollbackHr) || rollbackResult.outcomeKnown == FALSE || rollbackResult.mutationCommitted == FALSE)
                {
                    result->outcomeKnown = FALSE;
                    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                }
            }
            return attributesHr;
        }

        FileSystemConditionalMutationResult publishResult{};
        publishResult.sizeBytes = sizeof(publishResult);
        const HRESULT publishHr = RenameExactTo(destinationPath, FALSE, &publishResult);
        if (publishResult.outcomeKnown == FALSE)
        {
            result->outcomeKnown = FALSE;
            return FAILED(publishHr) ? publishHr : HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (FAILED(publishHr) || publishResult.mutationCommitted == FALSE)
        {
            if (destinationMovedToBackup)
            {
                const Common::SynchronousIoCancelWatch::CleanupWindow rollbackWindow(FreshOwnedStageCleanupDeadline());
                FileSystemConditionalMutationResult rollbackResult{};
                rollbackResult.sizeBytes = sizeof(rollbackResult);
                HRESULT rollbackHr       = expectedControl->RenameExactTo(destinationPath, FALSE, &rollbackResult);
                if (rollbackWindow.CanceledInsideWindow(rollbackHr))
                {
                    rollbackHr = expectedControl->RenameExactTo(destinationPath, FALSE, &rollbackResult);
                }
                if (FAILED(rollbackHr) || rollbackResult.outcomeKnown == FALSE || rollbackResult.mutationCommitted == FALSE)
                {
                    result->outcomeKnown = FALSE;
                    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                }
            }
            return FAILED(publishHr) ? publishHr : E_UNEXPECTED;
        }

        _ownedStage                  = false;
        result->mutationCommitted    = TRUE;
        result->originalStillPresent = FALSE;
        *renamed                     = static_cast<IFileSystemBoundObject*>(this);
        AddRef();

        if (destinationMovedToBackup)
        {
            // I35 R7.4: the replaced original's backup is removed under the cleanup budget too.
            const Common::SynchronousIoCancelWatch::CleanupWindow backupCleanupWindow(FreshOwnedStageCleanupDeadline());
            FileSystemConditionalMutationResult cleanupResult{};
            cleanupResult.sizeBytes = sizeof(cleanupResult);
            HRESULT cleanupHr       = expectedControl->DeleteExact(&cleanupResult);
            if (backupCleanupWindow.CanceledInsideWindow(cleanupHr))
            {
                cleanupHr = expectedControl->DeleteExact(&cleanupResult);
            }
            if (FAILED(cleanupHr) || cleanupResult.outcomeKnown == FALSE || cleanupResult.mutationCommitted == FALSE)
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
            expectedReadOnlyCleared = false;
        }
#if defined(ENABLE_TESTS)
        // Unlike a promotion failure, final-attributes failure occurs after the
        // destination changed. Keep the committed receipt and published authority.
        if (requireOwnedStage &&
            ShouldFailOwnedPublicationForSelfTest(
                destinationPath, kFinalAttributesFailPathEnvVar, kFinalAttributesFailFiredEnvVar, L"FileOps.Copy.DebugFinalAttributesFailureInjected"))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
#endif
        return S_OK;
    }

    std::atomic_ulong _refCount{1u};
    wil::unique_handle _file;
    LocalObjectIdentityPayload _identity{};
#if defined(ENABLE_TESTS)
    LocalObjectIdentityPayload _reportedIdentity{};
#endif
    FileSystemBoundObjectKind _kind   = FILESYSTEM_BOUND_OTHER;
    uint64_t _committedSizeBytes      = std::numeric_limits<uint64_t>::max();
    FileSystemBindFlags _grantedFlags = FILESYSTEM_BIND_NO_FOLLOW;
    bool _ownedStage                  = false;
    const bool _consumableReadWitness = false;
    std::mutex _witnessMutex;
    size_t _witnessReaders = 0u;
    std::wstring _path;
};

class Win32FileWriter final : public IFileWriter, public ILocalFileWriterAllocation
{
public:
    Win32FileWriter(wil::unique_handle file, std::wstring path, const FileSystemOptions* options) noexcept
        : _file(std::move(file)),
          _path(std::move(path)),
          _hasOperationOptions(options != nullptr)
    {
        if (options != nullptr)
        {
            _operationOptions = *options;
        }
    }

    Win32FileWriter(
        wil::unique_handle file, std::wstring tempPath, std::wstring finalPath, bool allowReplaceReadOnly, const FileSystemOptions* options) noexcept
        : _file(std::move(file)),
          _path(std::move(tempPath)),
          _finalPath(std::move(finalPath)),
          _allowReplaceReadOnly(allowReplaceReadOnly),
          _replaceOnCommit(true),
          _hasOperationOptions(options != nullptr)
    {
        if (options != nullptr)
        {
            _operationOptions = *options;
        }
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
        if (riid == __uuidof(ILocalFileWriterAllocation))
        {
            *ppvObject = static_cast<ILocalFileWriterAllocation*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE ReserveAllocation(uint64_t sizeBytes) noexcept override
    {
        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }
        if (sizeBytes > static_cast<uint64_t>(std::numeric_limits<LONGLONG>::max()))
        {
            return E_INVALIDARG;
        }

        FILE_ALLOCATION_INFO allocation{};
        allocation.AllocationSize.QuadPart = static_cast<LONGLONG>(sizeBytes);
        if (SetFileInformationByHandle(_file.get(), FileAllocationInfo, &allocation, sizeof(allocation)) != FALSE)
        {
            return S_OK;
        }
        const DWORD error = GetLastError();
        // A destination that cannot reserve space receives the bytes exactly as before.
        if (error == ERROR_NOT_SUPPORTED || error == ERROR_INVALID_PARAMETER || error == ERROR_INVALID_FUNCTION)
        {
            return S_FALSE;
        }
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
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

        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

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

        const HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

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
        HRESULT controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (FlushFileBuffers(_file.get()) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        controlHr = FileSystemCheckOperationControl(OperationOptions());
        if (FAILED(controlHr))
        {
            return controlHr;
        }

        if (_replaceOnCommit)
        {
            CapturePersistentIdentity();
            _file.reset();
            return PromoteTempIntoFinalPath();
        }

        _committed = true;
        return S_OK;
    }

private:
    ~Win32FileWriter()
    {
        if (_committed || _path.empty())
        {
            return;
        }

        HRESULT cleanupHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        if (_file)
        {
            cleanupHr = DeleteExactHandle(_file.get());
            if (SUCCEEDED(cleanupHr))
            {
                _file.reset();
                return;
            }
            CapturePersistentIdentity();
            _file.reset();
        }

        if (_hasPersistentIdentity)
        {
            cleanupHr = DeletePathOnlyWhenPersistentIdentityMatches(_path, _identity);
        }
        if (FAILED(cleanupHr))
        {
            Debug::Warning(L"FileSystem: aborted writer cleanup left an unverified path untouched (error={}).",
                           static_cast<unsigned long>(HRESULT_CODE(cleanupHr)));
        }
    }

    [[nodiscard]] const FileSystemOptions* OperationOptions() const noexcept
    {
        return _hasOperationOptions ? &_operationOptions : nullptr;
    }

    void CapturePersistentIdentity() noexcept
    {
        if (_hasPersistentIdentity || ! _file)
        {
            return;
        }
        _hasPersistentIdentity = SUCCEEDED(QueryPersistentIdentityFromRetainedHandle(_file.get(), _identity)) && ! IsRetainedHandleIdentity(_identity);
    }

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    LocalObjectIdentityPayload _identity{};
    bool _hasPersistentIdentity = false;
    std::wstring _path;
    std::wstring _finalPath;
    bool _allowReplaceReadOnly = false;
    bool _replaceOnCommit      = false;
    bool _committed            = false;
    FileSystemOptions _operationOptions{};
    bool _hasOperationOptions = false;

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

#ifdef _DEBUG

class DebugOperationControl final : public IFileSystemOperationControl
{
public:
    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (abort == nullptr)
        {
            return E_POINTER;
        }
        *abort = abortRequested ? TRUE : FALSE;
        ++abortChecks;
        return abortStatus;
    }

    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        if (mode == nullptr)
        {
            return E_POINTER;
        }
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }

    bool abortRequested      = false;
    HRESULT abortStatus      = S_OK;
    unsigned int abortChecks = 0u;
};

void RunDebugOperationControlSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    FILE_ID_INFO emptyFileIdInfo{};
    emptyFileIdInfo.VolumeSerialNumber = 42u;
    LocalObjectIdentityPayload emptyIdentity{};
    DebugCheck(MakePersistentLocalObjectIdentity(emptyFileIdInfo, emptyIdentity) == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
               L"an all-zero FILE_ID_128 must not become persistent object identity",
               passed,
               failed);
    FILE_ID_INFO validFileIdInfo{};
    validFileIdInfo.VolumeSerialNumber   = 42u;
    validFileIdInfo.FileId.Identifier[0] = 1u;
    LocalObjectIdentityPayload validIdentity{};
    DebugCheck(SUCCEEDED(MakePersistentLocalObjectIdentity(validFileIdInfo, validIdentity)) && ! IsRetainedHandleIdentity(validIdentity),
               L"a nonzero FILE_ID_128 should remain persistent object identity",
               passed,
               failed);

    DebugOperationControl control;
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
    options.operationControl = &control;

    DebugCheck(FileSystemCheckOperationControl(&options) == S_OK, L"operation-control checkpoint should accept a live non-aborted operation", passed, failed);
    control.abortRequested = true;
    DebugCheck(FileSystemCheckOperationControl(&options) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
               L"operation-control checkpoint should map an abort request to ERROR_CANCELLED",
               passed,
               failed);
    const uint64_t cleanupStartedAt        = GetTickCount64();
    const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(&options);
    DebugCheck(cleanupOptions.operationControl == nullptr && cleanupOptions.operationControlCookie == nullptr &&
                   cleanupOptions.deadlineTickCount64 > cleanupStartedAt &&
                   cleanupOptions.deadlineTickCount64 <= cleanupStartedAt + FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS &&
                   FileSystemCheckOperationControl(&cleanupOptions) == S_OK,
               L"owned-stage cleanup control should ignore primary cancellation under a fresh short deadline",
               passed,
               failed);
    control.abortRequested      = false;
    options.deadlineTickCount64 = 1u;
    DebugCheck(FileSystemCheckOperationControl(&options) == HRESULT_FROM_WIN32(ERROR_TIMEOUT),
               L"operation-control checkpoint should keep deadline expiry distinct from cancellation",
               passed,
               failed);

    wchar_t tempDirectory[MAX_PATH]{};
    wchar_t tempPath[MAX_PATH]{};
    const DWORD tempDirectoryLength = GetTempPathW(static_cast<DWORD>(std::size(tempDirectory)), tempDirectory);
    if (! DebugCheck(tempDirectoryLength != 0u && tempDirectoryLength < std::size(tempDirectory) && GetTempFileNameW(tempDirectory, L"rso", 0u, tempPath) != 0u,
                     L"operation-control selftest should create a temporary file",
                     passed,
                     failed))
    {
        return;
    }
    const auto removeTemp = wil::scope_exit([&]() noexcept { static_cast<void>(DeleteFileW(tempPath)); });

    wil::unique_handle seedHandle(CreateFileW(tempPath,
                                              GENERIC_READ | GENERIC_WRITE,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_ATTRIBUTE_TEMPORARY,
                                              nullptr));
    if (! DebugCheck(static_cast<bool>(seedHandle), L"operation-control selftest should open its temporary seed handle", passed, failed))
    {
        return;
    }
    constexpr std::array<std::byte, 4u> seed{};
    DWORD wrote = 0u;
    if (! DebugCheck(WriteFile(seedHandle.get(), seed.data(), static_cast<DWORD>(seed.size()), &wrote, nullptr) != FALSE && wrote == seed.size(),
                     L"operation-control selftest should seed the reader payload",
                     passed,
                     failed))
    {
        return;
    }
    wil::unique_handle readerHandle;
    const HRESULT readerReopenHr = ReopenExactFile(seedHandle.get(), GENERIC_READ, readerHandle, FILE_FLAG_OVERLAPPED);
    if (! DebugCheck(SUCCEEDED(readerReopenHr) && static_cast<bool>(readerHandle),
                     L"operation-control selftest should reopen an overlapped reader handle",
                     passed,
                     failed))
    {
        return;
    }

    options.deadlineTickCount64 = 0u;
    wil::com_ptr<IFileReader> reader;
    reader.attach(new (std::nothrow) Win32FileReader(std::move(readerHandle), seed.size(), &options));
    if (! DebugCheck(static_cast<bool>(reader), L"operation-control selftest should allocate its reader", passed, failed))
    {
        return;
    }
    control.abortRequested = true;
    std::array<std::byte, 4u> readBuffer{};
    unsigned long bytesRead = 0u;
    DebugCheck(reader->Read(readBuffer.data(), static_cast<unsigned long>(readBuffer.size()), &bytesRead) == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                   bytesRead == 0u,
               L"an already-open reader should observe a later abort before reading bytes",
               passed,
               failed);

    control.abortRequested      = false;
    options.deadlineTickCount64 = 1u;
    wil::unique_handle writerHandle(CreateFileW(
        tempPath, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_TEMPORARY, nullptr));
    wil::com_ptr<IFileWriter> writer;
    if (writerHandle)
    {
        writer.attach(new (std::nothrow) Win32FileWriter(std::move(writerHandle), std::wstring{}, &options));
    }
    if (! DebugCheck(static_cast<bool>(writer), L"operation-control selftest should allocate its writer", passed, failed))
    {
        return;
    }
    unsigned long bytesWritten = 0u;
    DebugCheck(writer->Write(seed.data(), static_cast<unsigned long>(seed.size()), &bytesWritten) == HRESULT_FROM_WIN32(ERROR_TIMEOUT) && bytesWritten == 0u,
               L"an already-open writer should observe its captured deadline before writing bytes",
               passed,
               failed);
    DebugCheck(writer->Commit() == HRESULT_FROM_WIN32(ERROR_TIMEOUT),
               L"writer Commit should observe its captured deadline before flushing or publishing",
               passed,
               failed);

    reader.reset();
    writer.reset();
    const std::wstring copyPath = std::wstring(tempPath) + L".copy";
    const auto removeCopy       = wil::scope_exit([&]() noexcept { static_cast<void>(DeleteFileW(copyPath.c_str())); });
    control.abortRequested      = true;
    options.deadlineTickCount64 = 0u;
    auto* fileSystem            = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"operation-control selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }
    const HRESULT copyHr = fileSystem->CopyItem(tempPath, copyPath.c_str(), FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr);
    fileSystem->Release();
    DebugCheck(copyHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && GetFileAttributesW(copyPath.c_str()) == INVALID_FILE_ATTRIBUTES,
               L"callback-free Local Copy should honor operationControl before destination mutation",
               passed,
               failed);
}

// R0f-SMB quiet point: a worker blocked inside a synchronous Win32 read (an anonymous pipe with no
// writer leaves the thread in the same pending-IRP state a dead SMB share does) returns with
// ERROR_OPERATION_ABORTED within the watch's grace once the operation control reports cancel.
void RunDebugSynchronousIoCancelWatchSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    class CancelControl final : public IFileSystemOperationControl
    {
    public:
        CancelControl() noexcept                       = default;
        CancelControl(const CancelControl&)            = delete;
        CancelControl& operator=(const CancelControl&) = delete;
        CancelControl(CancelControl&&)                 = delete;
        CancelControl& operator=(CancelControl&&)      = delete;

        HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
        {
            if (abort == nullptr)
            {
                return E_POINTER;
            }
            *abort = abortRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
            return S_OK;
        }
        HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
        {
            if (mode == nullptr)
            {
                return E_POINTER;
            }
            *mode = FILESYSTEM_DISCOVERY_AHEAD;
            return S_OK;
        }
        HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
        {
            return S_OK;
        }
        std::atomic<bool> abortRequested{false};
    };

    HANDLE readEndRaw  = nullptr;
    HANDLE writeEndRaw = nullptr;
    if (! DebugCheck(CreatePipe(&readEndRaw, &writeEndRaw, nullptr, 0u) != FALSE, L"cancel watch self-test needs an anonymous pipe", passed, failed))
    {
        return;
    }
    wil::unique_handle readEnd(readEndRaw);
    wil::unique_handle writeEnd(writeEndRaw);

    CancelControl control;
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
    options.operationControl = &control;

    std::atomic<bool> entered{false};
    std::atomic<bool> registered{false};
    std::atomic<HRESULT> blockedStatus{E_PENDING};
    std::atomic<unsigned int> cancelIssued{0u};
    std::thread reader([&]() noexcept
    {
        const Common::SynchronousIoCancelWatch::Scope scope(
            [](void* context) noexcept { return FAILED(FileSystemCheckOperationControl(static_cast<const FileSystemOptions*>(context))); }, &options);
        registered.store(scope.Registered(), std::memory_order_release);
        entered.store(true, std::memory_order_release);
        entered.notify_all();
        std::byte buffer[16]{};
        DWORD read    = 0u;
        const BOOL ok = ReadFile(readEnd.get(), buffer, static_cast<DWORD>(sizeof(buffer)), &read, nullptr);
        blockedStatus.store(ok != FALSE ? S_OK : HRESULT_FROM_WIN32(GetLastError()), std::memory_order_release);
        cancelIssued.store(scope.CancelIssuedCount(), std::memory_order_release);
    });
    entered.wait(false, std::memory_order_acquire);
    Sleep(100u); // let the read reach the kernel before cancel is requested
    const ULONGLONG cancelRequestedTick = GetTickCount64();
    control.abortRequested.store(true, std::memory_order_release);
    const bool returned       = WaitForSingleObject(reader.native_handle(), 5'000u) == WAIT_OBJECT_0;
    const ULONGLONG elapsedMs = GetTickCount64() - cancelRequestedTick;
    if (! returned)
    {
        // Release the wedged read so the self-test process can exit; the check below records the failure.
        DWORD wrote = 0u;
        static_cast<void>(WriteFile(writeEnd.get(), "x", 1u, &wrote, nullptr));
    }
    reader.join();

    DebugCheck(registered.load(std::memory_order_acquire), L"the cancel watch should register a worker thread", passed, failed);
    DebugCheck(returned, L"a blocked synchronous read must return after cancel is requested", passed, failed);
    DebugCheck(blockedStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED),
               L"the canceled synchronous read must fail with ERROR_OPERATION_ABORTED",
               passed,
               failed);
    DebugCheck(elapsedMs < 3'000u, L"the cancel watch must return a wedged call within its grace plus a few polls", passed, failed);
    DebugCheck(cancelIssued.load(std::memory_order_acquire) >= 1u, L"the cancel watch must have issued CancelSynchronousIo", passed, failed);

    // I35 R7.4: owned-stage cleanup has its own budget. Inside a cleanup window the watch issues no
    // cancel for the window's deadline although cancellation is requested (the grace passes inside it),
    // then cancels the still-wedged call as before, so a dead share still returns.
    HANDLE windowReadRaw  = nullptr;
    HANDLE windowWriteRaw = nullptr;
    if (! DebugCheck(CreatePipe(&windowReadRaw, &windowWriteRaw, nullptr, 0u) != FALSE, L"cleanup-window self-test needs an anonymous pipe", passed, failed))
    {
        return;
    }
    wil::unique_handle windowReadEnd(windowReadRaw);
    wil::unique_handle windowWriteEnd(windowWriteRaw);
    CancelControl windowControl;
    FileSystemOptions windowOptions{};
    windowOptions.sizeBytes        = sizeof(windowOptions);
    windowOptions.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
    windowOptions.operationControl = &windowControl;

    std::atomic<bool> windowEntered{false};
    std::atomic<bool> windowReturned{false};
    std::atomic<bool> releaseWindowScope{false};
    std::atomic<HRESULT> windowStatus{E_PENDING};
    std::atomic<ULONGLONG> windowOpenedTick{0u};
    std::atomic<ULONGLONG> windowReturnedTick{0u};
    std::atomic<const Common::SynchronousIoCancelWatch::Scope*> windowScope{nullptr};
    std::thread windowReader([&]() noexcept
    {
        const Common::SynchronousIoCancelWatch::Scope scope(
            [](void* context) noexcept { return FAILED(FileSystemCheckOperationControl(static_cast<const FileSystemOptions*>(context))); }, &windowOptions);
        windowScope.store(&scope, std::memory_order_release);
        {
            const Common::SynchronousIoCancelWatch::CleanupWindow window(GetTickCount64() + FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS);
            windowOpenedTick.store(GetTickCount64(), std::memory_order_release);
            windowEntered.store(true, std::memory_order_release);
            windowEntered.notify_all();
            std::byte buffer[16]{};
            DWORD read    = 0u;
            const BOOL ok = ReadFile(windowReadEnd.get(), buffer, static_cast<DWORD>(sizeof(buffer)), &read, nullptr);
            windowStatus.store(ok != FALSE ? S_OK : HRESULT_FROM_WIN32(GetLastError()), std::memory_order_release);
        }
        windowReturnedTick.store(GetTickCount64(), std::memory_order_release);
        windowReturned.store(true, std::memory_order_release);
        windowReturned.notify_all();
        // Keeps the registration alive until the checks below have read its counter.
        releaseWindowScope.wait(false, std::memory_order_acquire);
        windowScope.store(nullptr, std::memory_order_release);
    });
    windowEntered.wait(false, std::memory_order_acquire);
    Sleep(100u); // let the read reach the kernel before cancel is requested
    windowControl.abortRequested.store(true, std::memory_order_release);
    Sleep(1'500u);
    const Common::SynchronousIoCancelWatch::Scope* liveScope = windowScope.load(std::memory_order_acquire);
    const unsigned int issuedInsideWindow = liveScope != nullptr ? liveScope->CancelIssuedCount() : (std::numeric_limits<unsigned int>::max)();
    const bool pendingInsideWindow        = ! windowReturned.load(std::memory_order_acquire);
    const ULONGLONG waitStartedTick       = GetTickCount64();
    while (! windowReturned.load(std::memory_order_acquire) && GetTickCount64() - waitStartedTick < 4'000u)
    {
        Sleep(20u);
    }
    const bool returnedAfterWindow = windowReturned.load(std::memory_order_acquire);
    if (! returnedAfterWindow)
    {
        // Release the wedged read so the self-test process can exit; the checks below record the failure.
        DWORD wrote = 0u;
        static_cast<void>(WriteFile(windowWriteEnd.get(), "x", 1u, &wrote, nullptr));
        windowReturned.wait(false, std::memory_order_acquire);
    }
    const unsigned int issuedAfterWindow = liveScope != nullptr ? liveScope->CancelIssuedCount() : 0u;
    releaseWindowScope.store(true, std::memory_order_release);
    releaseWindowScope.notify_all();
    windowReader.join();

    const ULONGLONG windowHeldMs = windowReturnedTick.load(std::memory_order_acquire) - windowOpenedTick.load(std::memory_order_acquire);
    DebugCheck(liveScope != nullptr && pendingInsideWindow && issuedInsideWindow == 0u,
               L"the cancel watch must issue no CancelSynchronousIo for 1.5 s inside an owned-stage cleanup window",
               passed,
               failed);
    DebugCheck(returnedAfterWindow && windowStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && issuedAfterWindow >= 1u,
               L"after the cleanup window's deadline the cancel watch must cancel the still-wedged call",
               passed,
               failed);
    DebugCheck(windowHeldMs + 100u >= FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS && windowHeldMs < FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS + 2'000u,
               L"a wedged cleanup call must keep its whole budget and no more than a few polls beyond it",
               passed,
               failed);

    // Windows nest: an inner window is capped at the budget and restores the outer deadline, and the
    // outermost restores none.
    std::atomic<ULONGLONG>* const threadDeadline = Common::SynchronousIoCancelWatch::CurrentThreadCleanupDeadline();
    const ULONGLONG noWindow                     = threadDeadline != nullptr ? threadDeadline->load() : 1u;
    bool nestedRestores                          = threadDeadline != nullptr && noWindow == 0u;
    if (nestedRestores)
    {
        const Common::SynchronousIoCancelWatch::CleanupWindow outer(GetTickCount64() + 500u);
        const ULONGLONG outerDeadline = threadDeadline->load();
        {
            const Common::SynchronousIoCancelWatch::CleanupWindow inner(GetTickCount64() + 60'000u);
            nestedRestores =
                threadDeadline->load() >= outerDeadline && threadDeadline->load() <= GetTickCount64() + Common::SynchronousIoCancelWatch::kMaxCleanupWindowMs;
        }
        nestedRestores = nestedRestores && threadDeadline->load() == outerDeadline;
    }
    nestedRestores = nestedRestores && threadDeadline->load() == 0u;
    DebugCheck(nestedRestores, L"nested cleanup windows must cap at the budget and restore the previous deadline", passed, failed);

    // A call canceled inside its own open window (a poll that read the window before it opened) is
    // retried once; any other result, or a window whose deadline has passed, is not.
    bool retryRule = threadDeadline != nullptr;
    if (retryRule)
    {
        const Common::SynchronousIoCancelWatch::CleanupWindow open(GetTickCount64() + 2'000u);
        retryRule = open.CanceledInsideWindow(HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED)) && ! open.CanceledInsideWindow(S_OK) &&
                    ! open.CanceledInsideWindow(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    }
    if (retryRule)
    {
        const ULONGLONG now = GetTickCount64();
        const Common::SynchronousIoCancelWatch::CleanupWindow expired(now > 1u ? now - 1u : 0u);
        retryRule = ! expired.CanceledInsideWindow(HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED));
    }
    DebugCheck(retryRule, L"only a cleanup call canceled inside its own unexpired window is retried", passed, failed);
}

void RunDebugCreateDirectoryAdmissionSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    auto* fileSystem = new (std::nothrow) FileSystem();
    if (! DebugCheck(fileSystem != nullptr, L"Create Directory admission selftest should allocate a FileSystem instance", passed, failed))
    {
        return;
    }
    const auto releaseFileSystem = wil::scope_exit([&]() noexcept { fileSystem->Release(); });

    DebugCheck(fileSystem->CreateDirectory(LR"(\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\RedSalamanderMustNotCreate)") ==
                   HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
               L"direct Local CreateDirectory must reject GLOBALROOT before calling CreateDirectoryW",
               passed,
               failed);
    DebugCheck(fileSystem->CreateDirectory(LR"(\\.\PhysicalDrive0\RedSalamanderMustNotCreate)") == HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
               L"direct Local CreateDirectory must reject Win32 device namespaces before calling CreateDirectoryW",
               passed,
               failed);
    DebugCheck(fileSystem->CreateDirectory(LR"(\??\C:\RedSalamanderMustNotCreate)") == HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
               L"direct Local CreateDirectory must reject NT object-manager spelling before calling CreateDirectoryW",
               passed,
               failed);
}

// I35 R7.3 (pinned tree): the manifest of a folder deleted by name is canonical (ordinal over the upper-case
// form, every entry after the folder that holds it), the same for an unchanged tree, different once an entry
// is added deep in it, and leaves a folder's last-write time out; a listed file's own handle matches its entry
// until the file changes; a tree over the entry bound is refused.
void RunDebugDeleteManifestSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    std::array<wchar_t, MAX_PATH + 1u> tempPathBuffer{};
    const DWORD tempPathLength = GetTempPathW(static_cast<DWORD>(tempPathBuffer.size()), tempPathBuffer.data());
    if (! DebugCheck(tempPathLength > 0u && tempPathLength < tempPathBuffer.size(), L"delete manifest selftest needs a temporary path", passed, failed))
    {
        return;
    }
    const std::filesystem::path root =
        std::filesystem::path(tempPathBuffer.data()) / std::format(L"RedSalamander-DeleteManifest-{}-{}", GetCurrentProcessId(), GetTickCount64());
    const auto cleanup   = wil::scope_exit([&]() noexcept
    {
        std::error_code ignored;
        static_cast<void>(std::filesystem::remove_all(root, ignored));
    });
    const auto writeFile = [](const std::filesystem::path& path, std::string_view bytes) noexcept -> bool
    {
        wil::unique_hfile file(CreateFileW(path.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        DWORD written = 0u;
        return file && WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) != FALSE && written == bytes.size();
    };
    if (! DebugCheck(CreateDirectoryW(root.c_str(), nullptr) != FALSE && CreateDirectoryW((root / L"A").c_str(), nullptr) != FALSE &&
                         CreateDirectoryW((root / L"A B").c_str(), nullptr) != FALSE && CreateDirectoryW((root / L"A" / L"B").c_str(), nullptr) != FALSE &&
                         writeFile(root / L"A" / L"B" / L"c.txt", "c") && writeFile(root / L"A" / L"d.txt", "dd") && writeFile(root / L"e.txt", "eee"),
                     L"delete manifest selftest fixture should be created",
                     passed,
                     failed))
    {
        return;
    }

    const std::wstring extendedRoot = FileSystemInternal::ToExtendedPath(root.native());
    const auto list                 = [&](FileSystemInternal::LocalDeleteManifest& manifest, const unsigned int concurrency = 1u) noexcept -> HRESULT
    {
        wil::unique_handle folder(CreateFileW(extendedRoot.c_str(),
                                              FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                              nullptr));
        if (! folder)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        FileSystemInternal::LocalDeleteManifestListing listing{};
        listing.concurrency = concurrency;
        return FileSystemInternal::ListLocalDeleteManifest(folder.get(), extendedRoot, listing, manifest);
    };
    FileSystemInternal::LocalDeleteManifest first;
    FileSystemInternal::LocalDeleteManifest second;
    if (! DebugCheck(SUCCEEDED(list(first)) && SUCCEEDED(list(second)), L"delete manifest should list the fixture twice", passed, failed))
    {
        return;
    }
    constexpr std::array<std::wstring_view, 6> kExpectedOrder{{L"A", L"A B", L"A\\B", L"A\\B\\c.txt", L"A\\d.txt", L"e.txt"}};
    bool ordered = first.entries.size() == kExpectedOrder.size();
    for (size_t index = 0u; ordered && index < kExpectedOrder.size(); ++index)
    {
        ordered = first.entries[index].relativePath == kExpectedOrder[index];
    }
    DebugCheck(ordered && first.files == 3u && first.directories == 3u && first.links == 0u,
               L"the delete manifest sorts ordinally over the upper-case form, every entry after the folder that holds it",
               passed,
               failed);
    DebugCheck(first.digest == second.digest, L"an unchanged tree lists to the same manifest digest", passed, failed);
    FileSystemInternal::LocalDeleteManifest parallel;
    DebugCheck(SUCCEEDED(list(parallel, 4u)) && parallel.digest == first.digest && parallel.entries.size() == first.entries.size(),
               L"listing a tree's folders at the same time gives the serial listing's manifest",
               passed,
               failed);
    DebugCheck(std::ranges::all_of(first.entries,
                                   [](const FileSystemInternal::LocalDeleteManifestEntry& entry) noexcept
    { return (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (entry.lastWrite == 0u && entry.sizeBytes == 0u); }),
               L"a folder's last-write time and size are left out of the manifest",
               passed,
               failed);

    // An entry added deep in the tree changes the digest and the counts.
    FileSystemInternal::LocalDeleteManifest deeper;
    DebugCheck(writeFile(root / L"A" / L"B" / L"new.txt", "n") && SUCCEEDED(list(deeper)) && deeper.digest != first.digest && deeper.files == 4u,
               L"an entry added deep in the tree changes the manifest digest",
               passed,
               failed);

    // A listed file's own no-follow handle matches its entry until the file changes.
    const auto listedE =
        std::ranges::find_if(first.entries, [](const FileSystemInternal::LocalDeleteManifestEntry& entry) noexcept { return entry.relativePath == L"e.txt"; });
    const auto factsOf = [](const std::filesystem::path& path, BY_HANDLE_FILE_INFORMATION& information) noexcept -> bool
    {
        wil::unique_handle file(CreateFileW(path.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
        return file && GetFileInformationByHandle(file.get(), &information) != FALSE;
    };
    BY_HANDLE_FILE_INFORMATION before{};
    BY_HANDLE_FILE_INFORMATION after{};
    DebugCheck(listedE != first.entries.end() && factsOf(root / L"e.txt", before) && FileSystemInternal::LocalDeleteManifestEntryMatches(*listedE, before) &&
                   writeFile(root / L"e.txt", "eeee") && factsOf(root / L"e.txt", after) &&
                   ! FileSystemInternal::LocalDeleteManifestEntryMatches(*listedE, after),
               L"a listed file's handle matches its entry until the file changes",
               passed,
               failed);

#if defined(ENABLE_TESTS)
    // Over the entry bound (a test-only bound of 2 here) the listing is refused.
    FileSystemInternal::LocalDeleteManifest overBound;
    const bool boundSet       = SetEnvironmentVariableW(kDeleteByNameEntryBudgetEnvVar.data(), L"2") != FALSE;
    const HRESULT overBoundHr = boundSet ? list(overBound) : E_FAIL;
    static_cast<void>(SetEnvironmentVariableW(kDeleteByNameEntryBudgetEnvVar.data(), nullptr));
    DebugCheck(boundSet && overBoundHr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA) && overBound.entries.empty(),
               L"a tree over the delete manifest's entry bound is refused",
               passed,
               failed);
#endif
}
#endif
} // namespace

namespace
{
// The self-test injection that emulates a redirector without FileIdInfo (the NAS probed on 2026-10-11) for every
// path at or below its root.
[[nodiscard]] bool LocalDirectoryFileIdUnsupportedForSelfTest([[maybe_unused]] const std::wstring& path) noexcept
{
#if defined(ENABLE_TESTS)
    if (IsBindIdentityUnsupportedForSelfTest(path.c_str()))
    {
        Debug::Perf::EmitCounter(L"FileOps.Local.DirectoryIdentityUnsupportedInjected");
        return true;
    }
#endif
    return false;
}

// The volume serial number and nonzero 64-bit file index of an open folder that is not a reparse point.
[[nodiscard]] HRESULT QueryLocalDirectoryFileIndex(HANDLE directory, FileSystemInternal::LocalDirectoryIdentity& identity) noexcept
{
    identity = {};
    BY_HANDLE_FILE_INFORMATION information{};
    if (GetFileInformationByHandle(directory, &information) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    if ((information.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (information.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }
    const uint64_t index = (static_cast<uint64_t>(information.nFileIndexHigh) << 32u) | information.nFileIndexLow;
    // A server without file numbers reports zero (or all ones): that is no identity at all.
    if (index == 0u || index == (std::numeric_limits<uint64_t>::max)())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    identity.proof              = FileSystemInternal::LocalDirectoryIdentity::Proof::FileIndex;
    identity.volumeSerialNumber = information.dwVolumeSerialNumber;
    identity.fileIndex          = index;
    return S_OK;
}
} // namespace

bool FileSystemInternal::IsRemoteStoragePath(std::wstring_view path) noexcept
{
    if (IsUncPath(path))
        return true;
    const std::wstring driveRoot = ExtractDriveRoot(path);
    return ! driveRoot.empty() && GetDriveTypeW(driveRoot.c_str()) == DRIVE_REMOTE;
}

HRESULT FileSystemInternal::CaptureLocalDirectoryIdentity(IFileSystemBoundObject* published,
                                                          const std::wstring& path,
                                                          LocalDirectoryIdentity& identity) noexcept
{
    identity = {};
    if (published == nullptr)
    {
        return E_POINTER;
    }
    FileSystemBoundObjectSnapshot snapshot{};
    snapshot.sizeBytes = sizeof(snapshot);
    HRESULT hr         = published->GetSnapshot(&snapshot);
    if (FAILED(hr))
    {
        return hr;
    }
    if (snapshot.kind != FILESYSTEM_BOUND_DIRECTORY || snapshot.objectId == nullptr || snapshot.objectIdBytes != sizeof(LocalObjectIdentityPayload))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    LocalObjectIdentityPayload payload{};
    std::memcpy(&payload, snapshot.objectId, sizeof(payload));
    if (! IsRetainedHandleIdentity(payload) && ! LocalDirectoryFileIdUnsupportedForSelfTest(path))
    {
        identity.proof              = LocalDirectoryIdentity::Proof::FileId;
        identity.volumeSerialNumber = payload.volumeSerialNumber;
        identity.fileId             = payload.fileId;
        return S_OK;
    }

    // No FileIdInfo: the stage's identity is its retained handle, which is still open, so the file index it reports
    // belongs to this folder now.
    wil::com_ptr<ILocalBoundObjectControl> control;
    hr = published->QueryInterface(__uuidof(ILocalBoundObjectControl), control.put_void());
    if (FAILED(hr) || ! control)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    HANDLE duplicated = INVALID_HANDLE_VALUE;
    hr                = control->DuplicateExactHandle(&duplicated);
    if (FAILED(hr))
    {
        return hr;
    }
    const wil::unique_handle handle(duplicated);
    return QueryLocalDirectoryFileIndex(handle.get(), identity);
}

HRESULT FileSystemInternal::RestoreLocalDirectoryBasicInformation(const std::wstring& path,
                                                                  const LocalDirectoryIdentity& identity,
                                                                  const FileSystemBasicInformation& info) noexcept
{
    if (path.empty() || identity.proof == LocalDirectoryIdentity::Proof::None || info.sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }
    const std::wstring extendedPath = ToExtendedPath(path);
    const wil::unique_handle directory(CreateFileW(extendedPath.c_str(),
                                                   FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES | SYNCHRONIZE,
                                                   FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                   nullptr,
                                                   OPEN_EXISTING,
                                                   FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                   nullptr));
    if (! directory)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    bool same = false;
    if (identity.proof == LocalDirectoryIdentity::Proof::FileId)
    {
        FILE_ATTRIBUTE_TAG_INFO tag{};
        if (GetFileInformationByHandleEx(directory.get(), FileAttributeTagInfo, &tag, sizeof(tag)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        if ((tag.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (tag.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }
        // Recorded by FileIdInfo, compared by FileIdInfo only: a file index is never a substitute for it.
        if (LocalDirectoryFileIdUnsupportedForSelfTest(path))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        LocalObjectIdentityPayload current{};
        const HRESULT identityHr = QueryExactLocalObjectIdentity(directory.get(), current);
        if (FAILED(identityHr))
        {
            return identityHr;
        }
        same = current.volumeSerialNumber == identity.volumeSerialNumber && std::memcmp(&current.fileId, &identity.fileId, sizeof(identity.fileId)) == 0;
    }
    else
    {
        LocalDirectoryIdentity current{};
        const HRESULT identityHr = QueryLocalDirectoryFileIndex(directory.get(), current);
        if (FAILED(identityHr))
        {
            return identityHr;
        }
        same = current.volumeSerialNumber == identity.volumeSerialNumber && current.fileIndex == identity.fileIndex;
    }
    if (! same)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
    }

    FILE_BASIC_INFO basic{};
    basic.CreationTime.QuadPart   = info.creationTime;
    basic.LastAccessTime.QuadPart = info.lastAccessTime;
    basic.LastWriteTime.QuadPart  = info.lastWriteTime;
    basic.FileAttributes          = info.attributes;
    if (SetFileInformationByHandle(directory.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    return S_OK;
}

#if defined(ENABLE_TESTS)
// Test-only POD boundary. Begin/End and snapshot operate on caller TLS only.
// Cache reset is caller-authorized only after the fixture establishes worker quiet.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestAdmissionProbe(uint32_t action, uint32_t slotCount, uint64_t* slots) noexcept
{
    if (slotCount != kAdmissionProbeSlots || slots == nullptr || action > 5u)
        return E_INVALIDARG;
    if (action == 5u)
    {
        // Compatible side-band response: the original 19-slot count schema and
        // actions are unchanged. Read the independent fresh-root proof after End.
        if (g_admissionProbe.enabled)
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        std::fill_n(slots, kAdmissionProbeSlots, 0u);
        std::copy(g_admissionProbe.rootValues.begin(), g_admissionProbe.rootValues.end(), slots);
        return S_OK;
    }
    if (action == 2u || action == 4u)
    {
        if (g_admissionProbe.enabled)
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        std::scoped_lock lock(g_localVolumeRouteCache.mutex);
        if (action == 2u)
        {
            g_localVolumeRouteCache.volumeRoot.clear();
            g_localVolumeRouteCache.volumeGuidDevicePath.clear();
            g_localVolumeRouteCache.route        = {};
            g_localVolumeRouteCache.resolvedTick = 0u;
        }
        std::fill_n(slots, kAdmissionProbeSlots, 0u);
        slots[static_cast<size_t>(AdmissionProbeSlot::Schema)]   = 1u;
        slots[static_cast<size_t>(AdmissionProbeSlot::OwnerTid)] = GetCurrentThreadId();
        // Action 4 has a distinct metadata response: slot 2 is cache resolvedTick.
        // This is not a count sample; caller reads it before Begin overwrites slots.
        slots[2] = g_localVolumeRouteCache.resolvedTick;
        return S_OK;
    }
    if (action == 0u)
    {
        if (g_admissionProbe.enabled)
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        g_admissionProbe.values                                                        = {};
        g_admissionProbe.rootValues                                                    = {};
        g_admissionProbe.values[static_cast<size_t>(AdmissionProbeSlot::Schema)]       = 1u;
        g_admissionProbe.values[static_cast<size_t>(AdmissionProbeSlot::OwnerTid)]     = GetCurrentThreadId();
        g_admissionProbe.rootValues[static_cast<size_t>(LocalRootProbeSlot::Schema)]   = 2u;
        g_admissionProbe.rootValues[static_cast<size_t>(LocalRootProbeSlot::OwnerTid)] = GetCurrentThreadId();
        g_admissionProbe.enabled                                                       = true;
    }
    else if (action == 1u)
    {
        if (! g_admissionProbe.enabled)
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        g_admissionProbe.enabled = false;
    }
    else if (g_admissionProbe.enabled)
        return HRESULT_FROM_WIN32(ERROR_BUSY);
    std::copy(g_admissionProbe.values.begin(), g_admissionProbe.values.end(), slots);
    return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestDirectRootProbe(const wchar_t* path, BOOL* accepted) noexcept
{
    if (accepted == nullptr)
        return E_POINTER;
    *accepted = FALSE;
    if (path == nullptr || path[0] == L'\0')
        return E_INVALIDARG;
    std::wstring root;
    std::wstring guid;
    *accepted = TryResolveLocalRootWithoutReparse(path, root, guid) ? TRUE : FALSE;
    return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestDirectoryMetadataOpen(const wchar_t* nativePath, BOOL* opened) noexcept
{
    if (opened == nullptr)
        return E_POINTER;
    *opened = FALSE;
    if (nativePath == nullptr || nativePath[0] == L'\0')
        return E_INVALIDARG;
    // The caller supplies only its owned fixture's handle-derived NT name. Use
    // the actual production directory-open helper, never a duplicated access mask.
    const auto handle = OpenLocalRootProbeObject(nativePath, nullptr, true);
    *opened           = handle ? TRUE : FALSE;
    return S_OK;
}
#endif

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

    if (riid == __uuidof(IFileSystemPathCapabilities2))
    {
        *ppvObject = static_cast<IFileSystemPathCapabilities2*>(static_cast<IFileSystem*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemRouteCapabilities))
    {
        *ppvObject = static_cast<IFileSystemRouteCapabilities*>(static_cast<FileSystemRouteCapabilitiesBase*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemCooperativeCopy))
    {
        *ppvObject = static_cast<IFileSystemCooperativeCopy*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemObjectBinding))
    {
        *ppvObject = static_cast<IFileSystemObjectBinding*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemExpectedOccupantPublication))
    {
        *ppvObject = static_cast<IFileSystemExpectedOccupantPublication*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIdentityDelete))
    {
        *ppvObject = static_cast<IFileSystemIdentityDelete*>(this);
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
    wil::unique_handle file(CreateFileW(filePath.c_str(),
                                        GENERIC_READ,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED,
                                        nullptr));
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

    auto* impl = new (std::nothrow) Win32FileReader(std::move(file), static_cast<uint64_t>(fileSize.QuadPart), nullptr);
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *reader = impl;
    return S_OK;
}

#if defined(ENABLE_TESTS)
namespace
{
class LocalReaderCancellationTestControl final : public IFileSystemOperationControl
{
public:
    LocalReaderCancellationTestControl()                                                     = default;
    LocalReaderCancellationTestControl(const LocalReaderCancellationTestControl&)            = delete;
    LocalReaderCancellationTestControl(LocalReaderCancellationTestControl&&)                 = delete;
    LocalReaderCancellationTestControl& operator=(const LocalReaderCancellationTestControl&) = delete;
    LocalReaderCancellationTestControl& operator=(LocalReaderCancellationTestControl&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (abort == nullptr)
        {
            return E_POINTER;
        }
        *abort = abortRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        abortChecks.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        if (mode == nullptr)
        {
            return E_POINTER;
        }
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }

    std::atomic_bool abortRequested{false};
    std::atomic_uint abortChecks{0u};
};

[[nodiscard]] uint64_t LocalReaderTestElapsedUs(const std::chrono::steady_clock::time_point startedAt) noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count());
}
} // namespace

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestLocalReaderCancellation(const wchar_t* ordinaryPath,
                                                                                                      uint64_t* cancelDurationUs,
                                                                                                      uint64_t* ordinaryDurationUs,
                                                                                                      uint64_t* ordinaryBytesRead,
                                                                                                      unsigned int* abortCheckCount,
                                                                                                      unsigned long* bytesReadAfterCancel,
                                                                                                      HRESULT* blockedReadStatus,
                                                                                                      BOOL* enteredPending,
                                                                                                      BOOL* seekReplayMatched,
                                                                                                      BOOL* cleanupComplete) noexcept
{
    if (ordinaryPath == nullptr || ordinaryPath[0] == L'\0' || cancelDurationUs == nullptr || ordinaryDurationUs == nullptr || ordinaryBytesRead == nullptr ||
        abortCheckCount == nullptr || bytesReadAfterCancel == nullptr || blockedReadStatus == nullptr || enteredPending == nullptr ||
        seekReplayMatched == nullptr || cleanupComplete == nullptr)
    {
        return E_POINTER;
    }

    *cancelDurationUs     = 0u;
    *ordinaryDurationUs   = 0u;
    *ordinaryBytesRead    = 0u;
    *abortCheckCount      = 0u;
    *bytesReadAfterCancel = 0u;
    *blockedReadStatus    = E_PENDING;
    *enteredPending       = FALSE;
    *seekReplayMatched    = FALSE;
    *cleanupComplete      = FALSE;

    auto* fileSystem = new (std::nothrow) FileSystem();
    if (fileSystem == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    const auto releaseFileSystem = wil::scope_exit([&]() noexcept { fileSystem->Release(); });

    const auto ordinaryStartedAt = std::chrono::steady_clock::now();
    wil::com_ptr<IFileReader> legacyReader;
    HRESULT hr = fileSystem->CreateFileReader(ordinaryPath, legacyReader.put());
    if (FAILED(hr) || ! legacyReader)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    constexpr FileSystemBindFlags bindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> boundObject;
    hr = fileSystem->BindObject(ordinaryPath, bindFlags, boundObject.put());
    if (FAILED(hr) || ! boundObject)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    LocalReaderCancellationTestControl ordinaryControl;
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
    options.operationControl = &ordinaryControl;
    wil::com_ptr<IFileReader> boundReader;
    hr = boundObject->OpenReader(&options, boundReader.put());
    if (FAILED(hr) || ! boundReader)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    uint64_t expectedSize = 0u;
    hr                    = legacyReader->GetSize(&expectedSize);
    if (FAILED(hr))
    {
        return hr;
    }
    auto buffer = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[1024u * 1024u]);
    if (! buffer)
    {
        return E_OUTOFMEMORY;
    }
    const auto readWhole = [&](IFileReader& reader, uint64_t& total) noexcept -> HRESULT
    {
        total = 0u;
        for (;;)
        {
            unsigned long read   = 0u;
            const HRESULT readHr = reader.Read(buffer.get(), 1024u * 1024u, &read);
            if (FAILED(readHr))
            {
                return readHr;
            }
            if (read == 0u)
            {
                return S_OK;
            }
            total += read;
            if (total > expectedSize)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
    };

    uint64_t legacyBytes = 0u;
    uint64_t boundBytes  = 0u;
    hr                   = readWhole(*legacyReader, legacyBytes);
    if (SUCCEEDED(hr))
    {
        hr = readWhole(*boundReader, boundBytes);
    }
    if (FAILED(hr) || legacyBytes != expectedSize || boundBytes != expectedSize)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    *ordinaryBytesRead = legacyBytes + boundBytes;

    constexpr __int64 replayOffset = 17;
    uint64_t newPosition           = 0u;
    std::array<std::byte, 4096u> legacyReplay{};
    std::array<std::byte, 4096u> boundReplay{};
    unsigned long legacyReplayBytes = 0u;
    unsigned long boundReplayBytes  = 0u;
    hr                              = legacyReader->Seek(replayOffset, FILE_BEGIN, &newPosition);
    if (SUCCEEDED(hr))
    {
        hr = legacyReader->Read(legacyReplay.data(), static_cast<unsigned long>(legacyReplay.size()), &legacyReplayBytes);
    }
    if (SUCCEEDED(hr))
    {
        hr = boundReader->Seek(replayOffset, FILE_BEGIN, &newPosition);
    }
    if (SUCCEEDED(hr))
    {
        hr = boundReader->Read(boundReplay.data(), static_cast<unsigned long>(boundReplay.size()), &boundReplayBytes);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    *seekReplayMatched =
        legacyReplayBytes == boundReplayBytes && std::equal(legacyReplay.begin(), legacyReplay.begin() + legacyReplayBytes, boundReplay.begin()) ? TRUE : FALSE;
    *ordinaryDurationUs = LocalReaderTestElapsedUs(ordinaryStartedAt);

    const std::wstring pipeName =
        std::format(L"\\\\.\\pipe\\RedSalamander.LocalReaderCancel.{}.{}.{}", GetCurrentProcessId(), GetCurrentThreadId(), GetTickCount64());
    wil::unique_handle serverPipe(CreateNamedPipeW(pipeName.c_str(),
                                                   PIPE_ACCESS_OUTBOUND | FILE_FLAG_FIRST_PIPE_INSTANCE,
                                                   PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
                                                   1u,
                                                   4096u,
                                                   4096u,
                                                   0u,
                                                   nullptr));
    if (! serverPipe)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wil::unique_handle clientPipe(CreateFileW(pipeName.c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, nullptr));
    if (! clientPipe)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (ConnectNamedPipe(serverPipe.get(), nullptr) == FALSE)
    {
        const DWORD connectError = GetLastError();
        if (connectError != ERROR_PIPE_CONNECTED)
        {
            return HRESULT_FROM_WIN32(connectError);
        }
    }

    LocalReaderCancellationTestControl blockedControl;
    options.operationControl = &blockedControl;
    wil::com_ptr<IFileReader> blockedReader;
    blockedReader.attach(new (std::nothrow) Win32FileReader(std::move(clientPipe), 1u, &options));
    if (! blockedReader)
    {
        return E_OUTOFMEMORY;
    }

    std::atomic_bool readDone{false};
    std::atomic_long readStatus{E_PENDING};
    std::atomic_ulong readBytes{0u};
    std::jthread readerThread([&]() noexcept
    {
        std::byte byte{};
        unsigned long read   = 0u;
        const HRESULT readHr = blockedReader->Read(&byte, 1u, &read);
        readBytes.store(read, std::memory_order_release);
        readStatus.store(readHr, std::memory_order_release);
        readDone.store(true, std::memory_order_release);
    });

    const auto pendingDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(250);
    while (! readDone.load(std::memory_order_acquire) && blockedControl.abortChecks.load(std::memory_order_acquire) < 2u &&
           std::chrono::steady_clock::now() < pendingDeadline)
    {
        Sleep(1u);
    }
    *enteredPending = ! readDone.load(std::memory_order_acquire) && blockedControl.abortChecks.load(std::memory_order_acquire) >= 2u ? TRUE : FALSE;

    const auto cancelStartedAt = std::chrono::steady_clock::now();
    blockedControl.abortRequested.store(true, std::memory_order_release);
    const auto cancelDeadline = cancelStartedAt + std::chrono::milliseconds(500);
    while (! readDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < cancelDeadline)
    {
        Sleep(1u);
    }
    if (! readDone.load(std::memory_order_acquire))
    {
        serverPipe.reset();
        const auto cleanupDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
        while (! readDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < cleanupDeadline)
        {
            Sleep(1u);
        }
    }
    else
    {
        serverPipe.reset();
    }

    *cancelDurationUs     = LocalReaderTestElapsedUs(cancelStartedAt);
    *abortCheckCount      = blockedControl.abortChecks.load(std::memory_order_acquire);
    *bytesReadAfterCancel = readBytes.load(std::memory_order_acquire);
    *blockedReadStatus    = readStatus.load(std::memory_order_acquire);
    *cleanupComplete      = readDone.load(std::memory_order_acquire) ? TRUE : FALSE;
    return *cleanupComplete == TRUE ? S_OK : HRESULT_FROM_WIN32(ERROR_TIMEOUT);
}
#endif

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
        const std::wstring_view leaf = std::wstring_view(filePath).substr(separator + 1u);
        const Common::Paths::UniqueSiblingFileOptions options{
            .shareMode = FILE_SHARE_READ, .flagsAndAttributes = FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, .maximumAttempts = 32u};
        // Keep the random token and the extension intact when the readable leaf consumes the
        // volume's component limit; the shared creator still owns exclusive creation/retries.
        const HRESULT tempHr =
            Common::Paths::CreateUniqueFileInDirectory(filePath.substr(0u, separator + 1u), options, [leaf](size_t, std::wstring& name) noexcept {
            return Common::Paths::BuildUniqueSiblingName<wchar_t>(leaf, L".", FileSystemInternal::kStageExtension, 255u, name);
        }, tempPath, file);
        if (FAILED(tempHr))
        {
            return tempHr;
        }

        auto* impl = new (std::nothrow) Win32FileWriter(std::move(file), tempPath, filePath, allowReplaceReadOnly, nullptr);
        if (! impl)
        {
            LocalObjectIdentityPayload tempIdentity{};
            const bool hasPersistentIdentity =
                SUCCEEDED(QueryPersistentIdentityFromRetainedHandle(file.get(), tempIdentity)) && ! IsRetainedHandleIdentity(tempIdentity);
            HRESULT cleanupHr = DeleteExactHandle(file.get());
            file.reset();
            if (FAILED(cleanupHr) && hasPersistentIdentity)
            {
                cleanupHr = DeletePathOnlyWhenPersistentIdentityMatches(tempPath, tempIdentity);
            }
            if (FAILED(cleanupHr))
            {
                Debug::Warning(L"FileSystem: writer allocation failure left an unverified stage untouched (error={}).",
                               static_cast<unsigned long>(HRESULT_CODE(cleanupHr)));
            }
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

    auto* impl = new (std::nothrow) Win32FileWriter(std::move(file), filePath, nullptr);
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

#if defined(ENABLE_TESTS)
HRESULT STDMETHODCALLTYPE FileSystem::GetRouteFacts(const wchar_t* path,
                                                    FileSystemOperation operation,
                                                    FileSystemArena* arena,
                                                    FileSystemRouteFacts* facts) noexcept
{
    CountAdmissionProbe(AdmissionProbeSlot::RouteFacts);
    if (operation == FILESYSTEM_COPY)
        CountAdmissionProbe(AdmissionProbeSlot::RouteFactsCopy);
    else if (operation == FILESYSTEM_MOVE)
        CountAdmissionProbe(AdmissionProbeSlot::RouteFactsMove);
    return FileSystemRouteCapabilitiesBase::GetRouteFacts(path, operation, arena, facts);
}
HRESULT STDMETHODCALLTYPE FileSystem::IsTransferPeerAllowed(
    const wchar_t* path, FileSystemOperation operation, FileSystemTransferPeerRole role, const wchar_t* peerPluginId, BOOL* allowed) noexcept
{
    CountAdmissionProbe(AdmissionProbeSlot::PeerAllowed);
    return FileSystemRouteCapabilitiesBase::IsTransferPeerAllowed(path, operation, role, peerPluginId, allowed);
}
#endif

HRESULT STDMETHODCALLTYPE FileSystem::GetPathCapabilities(const wchar_t* path, FileSystemOperation operation, const char** jsonUtf8) noexcept
{
#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::LegacyPathCapabilities);
#endif
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }
    *jsonUtf8 = nullptr;
    if (path == nullptr || path[0] == L'\0' || operation < FILESYSTEM_COPY || operation > FILESYSTEM_CREATE_DIRECTORY)
    {
        return E_INVALIDARG;
    }

    const LocalCapabilityRouteInfo route = BuildLocalCapabilityRouteInfo(path);
    if (route.rootId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::lock_guard lock(_stateMutex);
    UpdateCapabilitiesJson(route.rootId, route.remote); // requires _stateMutex

    if (_capabilitiesJson.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    *jsonUtf8 = _capabilitiesJson.c_str();
    return S_OK;
}

HRESULT FileSystem::BuildFileSystemRouteDescriptor(const wchar_t* path, FileSystemOperation operation, FileSystemRouteDescriptor& descriptor) noexcept
{
#if defined(ENABLE_TESTS)
    CountAdmissionProbe(AdmissionProbeSlot::Descriptor);
#endif
    static_cast<void>(operation);
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    const std::wstring_view pathView(path);
    if (pathView != L"/" && ! Common::Paths::IsSupportedLocalFileOperationPath(pathView))
    {
        // The Local provider owns its typed route envelope. Reject Win32 device,
        // NT object-manager, GLOBALROOT, relative, and incomplete paths before
        // child-name validation or provider joining can authorize a mutation.
        // "/" remains the provider-root metadata query used before a pane has a
        // concrete Local path; mutation admission always re-queries that path.
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    const LocalCapabilityRouteInfo route = BuildLocalCapabilityRouteInfo(path);
    descriptor                           = {};
    descriptor.providerId                = kPluginId;
    descriptor.pathProfileId             = route.remote ? L"local-win32-smb" : L"local-win32";
    descriptor.rootId                    = Common::Strings::Utf16FromUtf8StrictOrEmpty(route.rootId);
    if (descriptor.rootId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    descriptor.availability = FILESYSTEM_ROUTE_AVAILABLE;
    // R0f-SMB: UNC and mapped-remote paths are bounded like fixed volumes. Every worker that
    // calls into this provider is registered with the synchronous-I/O cancel watch, so a call
    // wedged on a dead share returns with ERROR_OPERATION_ABORTED once the task is canceled.
    descriptor.cancellationRoute        = FILESYSTEM_CANCELLATION_BOUNDED;
    descriptor.namespaceKind            = FILESYSTEM_NAMESPACE_REAL_CONTAINER;
    descriptor.componentComparison      = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE;
    descriptor.caseOnlyRename           = FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED;
    descriptor.proofFlags               = FILESYSTEM_ROUTE_PROOF_HOST_READBACK | FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3;
    descriptor.copyOperation            = true;
    descriptor.moveOperation            = true;
    descriptor.nativeMoveOperation      = true;
    descriptor.deleteOperation          = true;
    descriptor.renameOperation          = true;
    descriptor.createDirectoryOperation = true;
    descriptor.propertiesOperation      = true;
    descriptor.readOperation            = true;
    descriptor.writeOperation           = true;
    descriptor.recycleOperation         = true;
    descriptor.boundDelete              = true;
    descriptor.conditionalDelete        = true;
    descriptor.exclusiveStage           = true;
    descriptor.conditionalPublish       = true;
    descriptor.committedSize            = true;
    descriptor.preserveFileLink         = true;
    descriptor.preserveDirectoryLink    = true;
    descriptor.retargetInTree           = false; // literal Preserve: the host never rewrites a link payload
    descriptor.exactLinkRemoval         = true;
    descriptor.exportCopyAll            = true;
    descriptor.exportMoveAll            = true;
    descriptor.importCopyAll            = true;
    descriptor.importMoveAll            = true;
    descriptor.preferredSeparator       = L'\\';
    descriptor.acceptedSeparators       = L"\\/";
    descriptor.windowsChildNames        = true;

    std::lock_guard lock(_stateMutex);
    descriptor.copyMoveMaxConcurrency         = std::clamp(_copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    descriptor.deleteMaxConcurrency           = std::clamp(_deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    descriptor.deleteRecycleBinMaxConcurrency = std::clamp(_deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::ReadExpectedOccupant(const wchar_t* path, FileSystemExpectedOccupant* occupant) noexcept
{
    if (path == nullptr || path[0] == L'\0' || occupant == nullptr || occupant->sizeBytes != sizeof(FileSystemExpectedOccupant))
    {
        return E_INVALIDARG;
    }
    FileSystemExpectedOccupant facts{};
    bool identityAvailable = false;
    const HRESULT hr       = ReadLocalOccupantFacts(path, facts, identityAvailable);
    if (FAILED(hr))
    {
        return hr;
    }
    // Only a regular file the location cannot identify is replaced through its facts. An item the
    // location identifies keeps the bound route; a folder, a link or another reparse point is never
    // replaced this way.
    if (identityAvailable || (facts.attributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    *occupant = facts;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::PublishOverExpectedOccupant(IFileSystemBoundObject* ownedStage,
                                                                  const wchar_t* finalPath,
                                                                  const FileSystemExpectedOccupant* expected,
                                                                  FileSystemFlags flags,
                                                                  const FileSystemOptions* options,
                                                                  FileSystemConditionalMutationResult* result,
                                                                  IFileSystemBoundObject** published) noexcept
{
    if (published != nullptr)
    {
        *published = nullptr;
    }
    if (ownedStage == nullptr || published == nullptr)
    {
        return E_POINTER;
    }
    wil::com_ptr<ILocalBoundObjectControl> control;
    if (FAILED(ownedStage->QueryInterface(__uuidof(ILocalBoundObjectControl), control.put_void())) || ! control)
    {
        return E_INVALIDARG;
    }
    return control->PublishOwnedOverExpected(finalPath, expected, flags, options, result, published);
}

// I63 B7: the bound snapshot's metadata classes for a file whose location cannot identify it, so the host can keep a
// Move source whose streams, extended attributes or encryption its pathname copy does not carry. Read-only, no-follow
// and fully shared; the backup-stream inspection is the bound snapshot's own and fails rather than report absence.
HRESULT STDMETHODCALLTYPE FileSystem::ReadUnidentifiedMetadataFeatures(const wchar_t* path,
                                                                       const FileSystemOptions* options,
                                                                       uint32_t* presentFeatures) noexcept
{
    if (presentFeatures == nullptr)
    {
        return E_POINTER;
    }
    *presentFeatures = 0u;
    if (path == nullptr || path[0] == L'\0' || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    const std::wstring extendedPath = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(extendedPath.c_str(),
                                        FILE_READ_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    if (! file)
    {
        // ERROR_NOT_SUPPORTED is this contract's answer for an identified object; a refused open is an open failure.
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_NOT_SUPPORTED ? ERROR_OPEN_FAILED : (error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND));
    }
    FILE_ATTRIBUTE_TAG_INFO tagInfo{};
    if (GetFileInformationByHandleEx(file.get(), FileAttributeTagInfo, &tagInfo, sizeof(tagInfo)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    if ((tagInfo.FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT identityHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
#if defined(ENABLE_TESTS)
    if (! IsBindIdentityUnsupportedForSelfTest(path))
#endif
    {
        LocalObjectIdentityPayload identity{};
        identityHr = QueryExactLocalObjectIdentity(file.get(), identity);
    }
    if (SUCCEEDED(identityHr))
    {
        // An object this location identifies is bound, and its bound metadata answers.
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (identityHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
    {
        // A failed query is not proof that the location has no identities.
        return identityHr;
    }

    uint32_t features = FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES | FILESYSTEM_METADATA_SECURITY;
    if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_SPARSE_FILE) != 0u)
    {
        features |= FILESYSTEM_METADATA_SPARSE;
    }
    if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_COMPRESSED) != 0u)
    {
        features |= FILESYSTEM_METADATA_COMPRESSION;
    }
    if ((tagInfo.FileAttributes & FILE_ATTRIBUTE_ENCRYPTED) != 0u)
    {
        features |= FILESYSTEM_METADATA_EFS;
    }
    constexpr DWORD kRecallAttributes = FILE_ATTRIBUTE_RECALL_ON_OPEN | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS;
    if ((tagInfo.FileAttributes & kRecallAttributes) != 0u)
    {
        features |= FILESYSTEM_METADATA_PLACEHOLDER;
    }
#if defined(ENABLE_TESTS)
    features |= ReadSelfTestMetadataMask(kMetadataForcePresentMaskEnvVar);
#endif
    const HRESULT streamHr = InspectBackupMetadataStreams(file.get(), features);
    if (FAILED(streamHr))
    {
        Debug::Perf::Emit(L"fileops.local.metadata.inspect_failed", L"unidentified", 0u, features, 0u, streamHr);
        return streamHr;
    }
    *presentFeatures = features;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::BindObject(const wchar_t* path, FileSystemBindFlags flags, IFileSystemBoundObject** bound) noexcept
{
    if (bound == nullptr)
    {
        return E_POINTER;
    }
    *bound = nullptr;
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    constexpr uint32_t knownFlags = FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE |
                                    FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION;
    const uint32_t requestedFlags = static_cast<uint32_t>(flags);
    if ((requestedFlags & ~knownFlags) != 0u || (requestedFlags & FILESYSTEM_BIND_NO_FOLLOW) == 0u)
    {
        return E_INVALIDARG;
    }
    const std::wstring extendedPath = FileSystemInternal::ToExtendedPath(path);
    DWORD desiredAccess             = FILE_READ_ATTRIBUTES;
    if ((requestedFlags & FILESYSTEM_BIND_READ_CONTENT) != 0u)
    {
        desiredAccess |= GENERIC_READ;
    }
    if ((requestedFlags & (FILESYSTEM_BIND_DELETE | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION)) != 0u)
    {
        desiredAccess |= DELETE;
    }
    if ((requestedFlags & FILESYSTEM_BIND_PUBLICATION) != 0u)
    {
        // Exact replacement may need to clear a retained destination's read-only attribute
        // before moving it to the rollback sibling. The publication grant must therefore own
        // that metadata write on the same no-follow handle.
        desiredAccess |= FILE_WRITE_ATTRIBUTES;
    }
    const DWORD shareMode = (requestedFlags & FILESYSTEM_BIND_DELETE) != 0u ? FILE_SHARE_READ : FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE;
    wil::unique_handle file;
#if defined(ENABLE_TESTS)
    if ((requestedFlags & FILESYSTEM_BIND_DELETE) != 0u && IsBindOpenFailureForSelfTest(path, kBindDeleteAccessDeniedPathEnvVar))
    {
        Debug::Perf::EmitCounter(L"FileOps.Local.BindDeleteAccessDeniedInjected");
        SetLastError(ERROR_ACCESS_DENIED);
    }
    else if (IsBindOpenNotSupportedForSelfTest(path))
    {
        Debug::Perf::EmitCounter(L"FileOps.Local.BindOpenNotSupportedInjected");
        SetLastError(ERROR_NOT_SUPPORTED);
    }
    else
#endif
    {
        file.reset(CreateFileW(
            extendedPath.c_str(), desiredAccess, shareMode, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    }
    if (! file)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_NOT_SUPPORTED)
        {
            // I35 R7.1: ERROR_NOT_SUPPORTED is this contract's answer for an object that opened but cannot
            // be identified. An open the location refused with that code is a failed open: it is reported
            // as one, so the host never treats the object as identity-less (never deleted by name, never
            // handed to a by-name card).
            Debug::Warning(L"FileSystem: the no-follow open of '{}' failed with ERROR_NOT_SUPPORTED; reported as ERROR_OPEN_FAILED.", path);
            return HRESULT_FROM_WIN32(ERROR_OPEN_FAILED);
        }
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND);
    }

    LocalObjectIdentityPayload identity{};
    HRESULT identityHr = S_OK;
#if defined(ENABLE_TESTS)
    if (IsBindIdentityUnsupportedForSelfTest(path))
    {
        Debug::Perf::EmitCounter(L"FileOps.Local.BindIdentityUnsupportedInjected");
        identityHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    else
#endif
    {
        identityHr = QueryExactLocalObjectIdentity(file.get(), identity);
    }
    if (FAILED(identityHr))
    {
        return identityHr;
    }

    FILE_ATTRIBUTE_TAG_INFO attributes{};
    if (GetFileInformationByHandleEx(file.get(), FileAttributeTagInfo, &attributes, sizeof(attributes)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    FileSystemBoundObjectKind kind = FILESYSTEM_BOUND_OTHER;
    if ((attributes.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u && IsReparseTagNameSurrogate(attributes.ReparseTag) != FALSE)
    {
        kind = FILESYSTEM_BOUND_LINK;
    }
    else if ((attributes.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        kind = FILESYSTEM_BOUND_DIRECTORY;
    }
    else
    {
        kind = FILESYSTEM_BOUND_REGULAR_FILE;
    }

    uint64_t committedSizeBytes = std::numeric_limits<uint64_t>::max();
    if (kind == FILESYSTEM_BOUND_REGULAR_FILE)
    {
        FILE_STANDARD_INFO standard{};
        if (GetFileInformationByHandleEx(file.get(), FileStandardInfo, &standard, sizeof(standard)) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        if (standard.EndOfFile.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        committedSizeBytes = static_cast<uint64_t>(standard.EndOfFile.QuadPart);
    }

    auto* impl = new (std::nothrow) LocalBoundObject(std::move(file), identity, kind, committedSizeBytes, flags, false, path);
    if (impl == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    *bound = impl;
    return S_OK;
}

namespace
{
// I35 R7.3 (D-7.3): the attributes a delete by name compares, on the selected item and on every listed entry.
constexpr DWORD kDeleteCorroborationAttributes =
    FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM | FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT;
constexpr size_t kDeleteManifestListingBufferBytes = 64u * 1024u;

[[nodiscard]] uint64_t DeleteCorroborationTicks(const FILETIME& time) noexcept
{
    return (static_cast<uint64_t>(time.dwHighDateTime) << 32u) | static_cast<uint64_t>(time.dwLowDateTime);
}

// 'l' a link or junction (a leaf, never entered), 'd' a folder, 'f' anything else.
[[nodiscard]] wchar_t DeleteCorroborationKind(const DWORD attributes) noexcept
{
    if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return L'l';
    }
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? L'd' : L'f';
}

// One manifest entry as hashed: fixed-width facts, then the relative path's UTF-16 code units.
struct DeleteManifestDigestRecord final
{
    uint32_t pathUnits  = 0u;
    uint32_t kind       = 0u;
    uint32_t attributes = 0u;
    uint32_t reserved   = 0u;
    uint64_t sizeBytes  = 0u;
    uint64_t lastWrite  = 0u;
    uint64_t creation   = 0u;
    uint64_t fileIndex  = 0u;
};
static_assert(sizeof(DeleteManifestDigestRecord) == 48u);

// I35 R7.3: an open or listing a location refuses with ERROR_NOT_SUPPORTED is an open failure: that code is the
// identity-delete contract's own answer for an object that has an exact identity.
[[nodiscard]] HRESULT DeleteManifestListingFailure(const DWORD error) noexcept
{
    if (error == ERROR_NOT_SUPPORTED)
    {
        return HRESULT_FROM_WIN32(ERROR_OPEN_FAILED);
    }
    return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
}

// A listed sub-folder that is gone, or is no longer a folder, when it is opened: the tree changed while it was listed.
[[nodiscard]] HRESULT DeleteManifestOpenFailure(const DWORD error) noexcept
{
    if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_DIRECTORY)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    return DeleteManifestListingFailure(error != ERROR_SUCCESS ? error : ERROR_OPEN_FAILED);
}
} // namespace

HRESULT FileSystemInternal::ListLocalDeleteManifest(HANDLE folder,
                                                    const std::wstring& folderExtendedPath,
                                                    const LocalDeleteManifestListing& listing,
                                                    LocalDeleteManifest& manifest) noexcept
{
    manifest = LocalDeleteManifest{};
    if (folder == nullptr || folder == INVALID_HANDLE_VALUE || folderExtendedPath.empty())
    {
        return E_INVALIDARG;
    }
    uint64_t entryBudget = kDeleteByNameManifestMaxEntries;
    // A redirector that refuses id listings is listed through the full class, with no file index; the choice
    // is made on the folder's first query, so every listing of one location takes the same shape.
    bool idListing = true;
#if defined(ENABLE_TESTS)
    if (const uint32_t testBudget = ReadSelfTestMetadataMask(kDeleteByNameEntryBudgetEnvVar); testBudget != 0u)
    {
        entryBudget = testBudget;
    }
    if (ReadSelfTestMetadataMask(kDeleteByNameListWithoutIndexEnvVar) != 0u)
    {
        idListing = false;
    }
#endif
    const auto startedAt = std::chrono::steady_clock::now();
    const auto elapsedUs = [&startedAt]() noexcept
    { return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count()); };

    // Every folder of one depth is listed at the same time (up to listing.concurrency), each through its own
    // no-follow handle; the bounds are shared, and the first failure stops the others.
    enum class Bound : uint8_t
    {
        None,
        Entries,
        Names,
        Depth,
    };
    std::atomic<uint64_t> listedEntries{0u};
    std::atomic<uint64_t> listedNameBytes{0u};
    std::atomic<Bound> exceeded{Bound::None};
    std::atomic<HRESULT> firstFailure{S_OK};
    std::atomic<bool> stop{false};
    // The folders of the current level listed to the end (a scheduler that shut down may leave some unrun).
    std::atomic<size_t> listedFolders{0u};
    std::mutex reportMutex;
    const auto fail = [&](const HRESULT hr) noexcept
    {
        HRESULT expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
        stop.store(true, std::memory_order_release);
    };
    const auto overBound = [&](const Bound bound) noexcept
    {
        Bound expected = Bound::None;
        static_cast<void>(exceeded.compare_exchange_strong(expected, bound, std::memory_order_acq_rel));
        fail(HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA));
    };

    struct PendingFolder final
    {
        std::wstring relativePath;
        uint64_t depth = 0u;
    };
    struct FolderListing final
    {
        std::vector<LocalDeleteManifestEntry> entries;
        std::vector<PendingFolder> subfolders;
    };
    // Only std::bad_alloc can escape the container operations below, and it is fatal by policy (AGENTS.md).
    const auto listOne = [&](const PendingFolder& current, HANDLE given, FolderListing& result) noexcept
    {
        if (stop.load(std::memory_order_acquire))
        {
            return;
        }
        if (const HRESULT controlHr = FileSystemCheckOperationControl(listing.options); FAILED(controlHr))
        {
            fail(controlHr);
            return;
        }
        std::unique_ptr<std::byte[]> buffer(new (std::nothrow) std::byte[kDeleteManifestListingBufferBytes]);
        if (! buffer)
        {
            fail(E_OUTOFMEMORY);
            return;
        }

        wil::unique_handle owned;
        HANDLE directory = given;
        if (directory == nullptr)
        {
            const std::wstring currentPath = AppendPath(folderExtendedPath, current.relativePath);
            DWORD openError                = ERROR_SUCCESS;
#if defined(ENABLE_TESTS)
            if (IsBindOpenNotSupportedForSelfTest(currentPath.c_str()))
            {
                openError = ERROR_NOT_SUPPORTED;
            }
#endif
            if (openError == ERROR_SUCCESS)
            {
                owned.reset(::CreateFileW(currentPath.c_str(),
                                          FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
                if (! owned)
                {
                    openError = ::GetLastError();
                    openError = openError != ERROR_SUCCESS ? openError : static_cast<DWORD>(ERROR_OPEN_FAILED);
                }
            }
            if (openError != ERROR_SUCCESS)
            {
                fail(DeleteManifestOpenFailure(openError));
                return;
            }
            FILE_BASIC_INFO basic{};
            if (! ::GetFileInformationByHandleEx(owned.get(), FileBasicInfo, &basic, sizeof(basic)))
            {
                fail(DeleteManifestListingFailure(::GetLastError()));
                return;
            }
            if ((basic.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (basic.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
            {
                fail(HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
                return;
            }
            directory = owned.get();
        }

        uint64_t files       = 0u;
        uint64_t directories = 0u;
        uint64_t bytes       = 0u;
        bool restart         = true;
        for (;;)
        {
            if (stop.load(std::memory_order_acquire))
            {
                return;
            }
            const FILE_INFO_BY_HANDLE_CLASS informationClass = idListing ? (restart ? FileIdBothDirectoryRestartInfo : FileIdBothDirectoryInfo)
                                                                         : (restart ? FileFullDirectoryRestartInfo : FileFullDirectoryInfo);
            if (! ::GetFileInformationByHandleEx(directory, informationClass, buffer.get(), static_cast<DWORD>(kDeleteManifestListingBufferBytes)))
            {
                const DWORD error = ::GetLastError();
                if (error == ERROR_NO_MORE_FILES)
                {
                    break;
                }
                // Only the selected folder's own first query, listed alone before any other, picks the class.
                if (given != nullptr && restart && idListing &&
                    (error == ERROR_INVALID_PARAMETER || error == ERROR_NOT_SUPPORTED || error == ERROR_INVALID_LEVEL))
                {
                    idListing = false;
                    continue;
                }
                fail(DeleteManifestListingFailure(error));
                return;
            }
            restart = false;

            size_t offset = 0u;
            for (;;)
            {
                LocalDeleteManifestEntry entry{};
                size_t nextOffset = 0u;
                std::wstring_view leaf;
                const std::span<const std::byte> records(buffer.get(), kDeleteManifestListingBufferBytes);
                if (idListing)
                {
                    Common::DirectoryRecords::RecordView<FILE_ID_BOTH_DIR_INFO> record{};
                    if (FAILED(Common::DirectoryRecords::Read<FILE_ID_BOTH_DIR_INFO>(records, offset, record)))
                    {
                        fail(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                        return;
                    }
                    const auto* info = record.information;
                    leaf             = record.name;
                    nextOffset       = record.nextOffset;
                    entry.attributes = info->FileAttributes;
                    entry.sizeBytes  = static_cast<uint64_t>(info->EndOfFile.QuadPart);
                    entry.lastWrite  = static_cast<uint64_t>(info->LastWriteTime.QuadPart);
                    entry.creation   = static_cast<uint64_t>(info->CreationTime.QuadPart);
                    entry.fileIndex  = static_cast<uint64_t>(info->FileId.QuadPart);
                }
                else
                {
                    Common::DirectoryRecords::RecordView<FILE_FULL_DIR_INFO> record{};
                    if (FAILED(Common::DirectoryRecords::Read<FILE_FULL_DIR_INFO>(records, offset, record)))
                    {
                        fail(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                        return;
                    }
                    const auto* info = record.information;
                    leaf             = record.name;
                    nextOffset       = record.nextOffset;
                    entry.attributes = info->FileAttributes;
                    entry.sizeBytes  = static_cast<uint64_t>(info->EndOfFile.QuadPart);
                    entry.lastWrite  = static_cast<uint64_t>(info->LastWriteTime.QuadPart);
                    entry.creation   = static_cast<uint64_t>(info->CreationTime.QuadPart);
                }
                if (! IsDotOrDotDot(leaf))
                {
                    if (listedEntries.fetch_add(1u, std::memory_order_acq_rel) + 1u > entryBudget)
                    {
                        overBound(Bound::Entries);
                        return;
                    }
                    std::wstring relativePath = current.relativePath.empty() ? std::wstring(leaf) : current.relativePath + L'\\' + std::wstring(leaf);
                    const uint64_t pathBytes  = static_cast<uint64_t>(relativePath.size()) * sizeof(wchar_t);
                    if (listedNameBytes.fetch_add(pathBytes, std::memory_order_acq_rel) + pathBytes > kDeleteByNameManifestMaxNameBytes)
                    {
                        overBound(Bound::Names);
                        return;
                    }
                    if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
                    {
                        ++directories;
                    }
                    else
                    {
                        ++files;
                        bytes += entry.sizeBytes;
                    }
                    if (DeleteCorroborationKind(entry.attributes) == L'd')
                    {
                        const uint64_t depth = current.depth + 1u;
                        if (depth > Common::FileOperations::kTraversalMaxDepth)
                        {
                            overBound(Bound::Depth);
                            return;
                        }
                        // A folder's size is not compared, nor its last-write time: its parent's listing updates
                        // that value lazily (it moved twice, 80 ms apart, for an unchanged tree on NTFS), and every
                        // change it reflects is an entry of the tree itself.
                        entry.sizeBytes = 0u;
                        entry.lastWrite = 0u;
                        result.subfolders.push_back(PendingFolder{relativePath, depth});
                    }
                    entry.relativePath = std::move(relativePath);
                    result.entries.push_back(std::move(entry));
                }
                if (nextOffset == 0u)
                    break;
                offset = nextOffset;
            }
        }
        if (listing.onFolderListed != nullptr && (files != 0u || directories != 0u))
        {
            std::scoped_lock lock(reportMutex);
            listing.onFolderListed(listing.cookie, files, directories, bytes);
        }
        listedFolders.fetch_add(1u, std::memory_order_acq_rel);
    };

    std::vector<PendingFolder> level;
    level.push_back(PendingFolder{});
    bool selectedFolderLevel = true;
    while (! level.empty())
    {
        std::vector<FolderListing> results(level.size());
        listedFolders.store(0u, std::memory_order_release);
        if (selectedFolderLevel)
        {
            listOne(level.front(), folder, results.front());
        }
        else
        {
            RunOnSharedFileOpsScheduler(
                listing.options, listing.concurrency, level.size(), [&](const size_t index) noexcept { listOne(level[index], nullptr, results[index]); });
        }
        selectedFolderLevel = false;
        if (SUCCEEDED(firstFailure.load(std::memory_order_acquire)) && listedFolders.load(std::memory_order_acquire) != level.size())
        {
            fail(HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED));
        }
        if (const HRESULT failure = firstFailure.load(std::memory_order_acquire); FAILED(failure))
        {
            if (failure == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA))
            {
                const Bound bound              = exceeded.load(std::memory_order_acquire);
                const std::wstring_view bounds = bound == Bound::Names ? L"names" : (bound == Bound::Depth ? L"depth" : L"entries");
                Debug::Warning(L"FileSystem: '{}' holds too many items to be deleted by name ({} bound: more than {} entries, {} bytes of names or {} "
                               L"levels); the delete by name is refused and nothing is deleted.",
                               folderExtendedPath,
                               bounds,
                               entryBudget,
                               kDeleteByNameManifestMaxNameBytes,
                               Common::FileOperations::kTraversalMaxDepth);
                Debug::Perf::Emit(L"FileOps.Local.DeleteByNameManifest",
                                  bounds,
                                  elapsedUs(),
                                  listedEntries.load(std::memory_order_acquire),
                                  listedNameBytes.load(std::memory_order_acquire),
                                  failure);
            }
            manifest = LocalDeleteManifest{};
            return failure;
        }
        std::vector<PendingFolder> next;
        for (FolderListing& result : results)
        {
            for (LocalDeleteManifestEntry& entry : result.entries)
            {
                const wchar_t kind = DeleteCorroborationKind(entry.attributes);
                if (kind == L'd')
                {
                    ++manifest.directories;
                }
                else if (kind == L'l')
                {
                    ++manifest.links;
                }
                else
                {
                    ++manifest.files;
                }
                manifest.entries.push_back(std::move(entry));
            }
            for (PendingFolder& subfolder : result.subfolders)
            {
                manifest.maxDepth = (std::max)(manifest.maxDepth, subfolder.depth);
                next.push_back(std::move(subfolder));
            }
        }
        level = std::move(next);
    }
    manifest.nameBytes = listedNameBytes.load(std::memory_order_acquire);

    std::ranges::sort(manifest.entries,
                      [](const LocalDeleteManifestEntry& left, const LocalDeleteManifestEntry& right) noexcept
    {
        const int folded = ::CompareStringOrdinal(
            left.relativePath.data(), static_cast<int>(left.relativePath.size()), right.relativePath.data(), static_cast<int>(right.relativePath.size()), TRUE);
        if (folded != CSTR_EQUAL)
        {
            return folded == CSTR_LESS_THAN;
        }
        return ::CompareStringOrdinal(left.relativePath.data(),
                                      static_cast<int>(left.relativePath.size()),
                                      right.relativePath.data(),
                                      static_cast<int>(right.relativePath.size()),
                                      FALSE) == CSTR_LESS_THAN;
    });

    Common::Crypto::Blake3Hasher hasher;
    constexpr std::string_view kManifestTag = "RedSalamander.DeleteByNameManifest.v1";
    hasher.Update(std::as_bytes(std::span<const char>(kManifestTag.data(), kManifestTag.size())));
    for (const LocalDeleteManifestEntry& entry : manifest.entries)
    {
        DeleteManifestDigestRecord record{};
        record.pathUnits  = static_cast<uint32_t>(entry.relativePath.size());
        record.kind       = static_cast<uint32_t>(DeleteCorroborationKind(entry.attributes));
        record.attributes = entry.attributes & kDeleteCorroborationAttributes;
        record.sizeBytes  = entry.sizeBytes;
        record.lastWrite  = entry.lastWrite;
        record.creation   = entry.creation;
        record.fileIndex  = entry.fileIndex;
        hasher.Update(std::as_bytes(std::span<const DeleteManifestDigestRecord>(&record, 1u)));
        hasher.Update(std::as_bytes(std::span<const wchar_t>(entry.relativePath.data(), entry.relativePath.size())));
    }
    const Common::Crypto::Blake3Digest digest = hasher.Finalize();
    static_assert(sizeof(manifest.digest) == sizeof(digest));
    std::copy(digest.begin(), digest.end(), manifest.digest.begin());
    Debug::Perf::Emit(L"FileOps.Local.DeleteByNameManifest",
                      idListing ? L"id" : L"full",
                      elapsedUs(),
                      static_cast<uint64_t>(manifest.entries.size()),
                      manifest.nameBytes,
                      S_OK);
    return S_OK;
}

bool FileSystemInternal::LocalDeleteManifestEntryMatches(const LocalDeleteManifestEntry& listed, const BY_HANDLE_FILE_INFORMATION& live) noexcept
{
    const wchar_t kind = DeleteCorroborationKind(listed.attributes);
    if (kind != DeleteCorroborationKind(live.dwFileAttributes) ||
        (listed.attributes & kDeleteCorroborationAttributes) != (live.dwFileAttributes & kDeleteCorroborationAttributes) ||
        listed.creation != DeleteCorroborationTicks(live.ftCreationTime))
    {
        return false;
    }
    const uint64_t liveIndex = (static_cast<uint64_t>(live.nFileIndexHigh) << 32u) | static_cast<uint64_t>(live.nFileIndexLow);
    if (listed.fileIndex != 0u && liveIndex != 0u && listed.fileIndex != liveIndex)
    {
        return false;
    }
    if (kind == L'd')
    {
        // Its removal proves what it holds: it fails while anything this listing never saw is inside.
        return true;
    }
    const uint64_t liveSize = (static_cast<uint64_t>(live.nFileSizeHigh) << 32u) | static_cast<uint64_t>(live.nFileSizeLow);
    return listed.sizeBytes == liveSize && listed.lastWrite == DeleteCorroborationTicks(live.ftLastWriteTime);
}

HRESULT FileSystemInternal::ResolveLocalDeleteCorroboration(HANDLE file,
                                                            const wchar_t* path,
                                                            FileSystemDeleteIdentity& identity,
                                                            const LocalDeleteManifest* manifest) noexcept
{
    identity.isDirectory = FALSE;
    identity.identity[0] = L'\0';
    if (file == nullptr || file == INVALID_HANDLE_VALUE || path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    // An object with an exact identity is bound and deleted through that identity, never by name.
    HRESULT exactIdentityHr = S_OK;
#if defined(ENABLE_TESTS)
    if (IsBindIdentityUnsupportedForSelfTest(path))
    {
        exactIdentityHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    else
#endif
    {
        LocalObjectIdentityPayload exactIdentity{};
        exactIdentityHr = QueryExactLocalObjectIdentity(file, exactIdentity);
    }
    if (SUCCEEDED(exactIdentityHr))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (exactIdentityHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
    {
        // A failed query is not proof that the location has no identities.
        return exactIdentityHr;
    }

    BY_HANDLE_FILE_INFORMATION information{};
    if (GetFileInformationByHandle(file, &information) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    const DWORD attributes = information.dwFileAttributes;
    const wchar_t kind     = DeleteCorroborationKind(attributes);
    const bool directory   = kind == L'd';
    const uint64_t size    = (static_cast<uint64_t>(information.nFileSizeHigh) << 32u) | static_cast<uint64_t>(information.nFileSizeLow);
    const uint64_t index   = (static_cast<uint64_t>(information.nFileIndexHigh) << 32u) | static_cast<uint64_t>(information.nFileIndexLow);

    std::wstring text = std::format(L"lc1|k={}|a={:08X}", kind, attributes & kDeleteCorroborationAttributes);
    if (! directory)
    {
        text += std::format(L"|s={:016X}", size);
    }
    // A folder compares its last-write time too: it changes when a child is added, removed or renamed.
    text += std::format(L"|w={:016X}|c={:016X}", DeleteCorroborationTicks(information.ftLastWriteTime), DeleteCorroborationTicks(information.ftCreationTime));
    if (index != 0u)
    {
        // Corroboration only: the probed NAS reports volume serial 0 and an inode number here.
        text += std::format(L"|v={:08X}|i={:016X}", information.dwVolumeSerialNumber, index);
    }
    if (directory && manifest != nullptr)
    {
        // I35 R7.3 (pinned tree): everything below the folder, as the digest of its manifest and its counts
        // (files, folders, links), so any entry added, removed, renamed or changed anywhere in the tree is a
        // difference.
        text += L"|m=";
        for (const std::byte value : manifest->digest)
        {
            text += std::format(L"{:02X}", static_cast<unsigned int>(value));
        }
        text += std::format(L"|n={:X}.{:X}.{:X}", manifest->files, manifest->directories, manifest->links);
    }
    if (text.size() >= std::size(identity.identity))
    {
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    }
    std::copy(text.begin(), text.end(), std::begin(identity.identity));
    identity.identity[text.size()] = L'\0';
    identity.isDirectory           = directory ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::ResolveDeleteIdentity(const wchar_t* path, const FileSystemOptions* options, FileSystemDeleteIdentity* identity) noexcept
{
    if (path == nullptr || identity == nullptr)
    {
        return E_POINTER;
    }
    if (path[0] == L'\0' || identity->sizeBytes != sizeof(FileSystemDeleteIdentity) || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    identity->isDirectory   = FALSE;
    identity->identity[0]   = L'\0';
    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    // Read-only, no-follow and fully shared: resolving the facts never blocks another program.
    const std::wstring extendedPath = FileSystemInternal::ToExtendedPath(path);
    wil::unique_handle file(CreateFileW(extendedPath.c_str(),
                                        FILE_READ_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    // ERROR_NOT_SUPPORTED is this contract's answer for an object with an exact identity; a location that
    // refuses the open itself reports an open failure.
    const auto openFailure = []() noexcept
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_NOT_SUPPORTED ? ERROR_OPEN_FAILED : (error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND));
    };
    if (! file)
    {
        return openFailure();
    }
    HRESULT hr = FileSystemInternal::ResolveLocalDeleteCorroboration(file.get(), path, *identity);
    if (SUCCEEDED(hr) && identity->isDirectory != FALSE)
    {
        // I35 R7.3 (pinned tree): a folder pins its whole subtree. It is listed through a no-follow listing
        // handle of its own (still fully shared), and its facts are taken again on that handle, so the folder's
        // facts and its manifest describe the one object that was listed.
        wil::unique_handle listing(CreateFileW(extendedPath.c_str(),
                                               FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               nullptr,
                                               OPEN_EXISTING,
                                               FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                               nullptr));
        FileSystemInternal::LocalDeleteManifest manifest;
        FileSystemInternal::LocalDeleteManifestListing listingOptions{};
        listingOptions.options = options;
        {
            std::lock_guard lock(_stateMutex);
            listingOptions.concurrency = std::clamp(_deleteMaxConcurrency, 1u, 8u);
        }
        hr = listing ? FileSystemInternal::ListLocalDeleteManifest(listing.get(), extendedPath, listingOptions, manifest) : openFailure();
        if (SUCCEEDED(hr))
        {
            FileSystemDeleteIdentity listed{};
            listed.sizeBytes = sizeof(listed);
            hr               = FileSystemInternal::ResolveLocalDeleteCorroboration(listing.get(), path, listed, &manifest);
            if (SUCCEEDED(hr) && listed.isDirectory == FALSE)
            {
                hr = HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
            }
            if (SUCCEEDED(hr))
            {
                *identity = listed;
            }
        }
        if (FAILED(hr))
        {
            identity->isDirectory = FALSE;
            identity->identity[0] = L'\0';
        }
    }
    Debug::Perf::Emit(L"FileOps.Local.ResolveDeleteIdentity", L"", 0u, identity->isDirectory != FALSE ? 1u : 0u, 0u, hr);
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateExclusiveWriter(const wchar_t* stagePath,
                                                            const FileSystemOptions* options,
                                                            IFileWriter** writer,
                                                            IFileSystemBoundObject** ownedStage) noexcept
{
    if (writer == nullptr || ownedStage == nullptr)
    {
        return E_POINTER;
    }
    *writer     = nullptr;
    *ownedStage = nullptr;
    if (stagePath == nullptr || stagePath[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    LocalObjectIdentityPayload retainedIdentity{};
    const HRESULT retainedIdentityHr = MakeRetainedHandleIdentity(retainedIdentity);
    if (FAILED(retainedIdentityHr))
    {
        return retainedIdentityHr;
    }

    // FILE_ATTRIBUTE_TEMPORARY keeps the object's dirty data in the cache until the commit flush. The
    // host may ask for it to be written back during the copy instead (I35 R10.3 candidate (e)); the
    // attribute never reaches a published object either way (publication sets the final attributes).
    // On a network share the host may ask for every write to reach the server before it returns
    // (WRITE_THROUGH_REMOTE): a cached writer there takes gigabytes at memory speed and is then held
    // inside one write for as long as the cache needs to drain them, which reads as a dead transfer.
    // Where the stage lies is looked up only when a candidate asks for it (a drive-type query, about
    // 3 us per writer).
    const uint32_t tuning = options != nullptr ? options->transferTuning : FILESYSTEM_TRANSFER_TUNING_NONE;
    DWORD stageAttributes = FILE_ATTRIBUTE_TEMPORARY;
    DWORD stageWriteFlags = 0u;
    constexpr uint32_t kLocationTuning =
        FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_REMOTE | FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_LOCAL | FILESYSTEM_TRANSFER_TUNING_WRITE_THROUGH_REMOTE;
    if ((tuning & kLocationTuning) != 0u)
    {
        const bool remoteStage     = FileSystemInternal::IsRemoteStoragePath(stagePath);
        const uint32_t noTemporary = remoteStage ? FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_REMOTE : FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_LOCAL;
        if (remoteStage && (tuning & FILESYSTEM_TRANSFER_TUNING_WRITE_THROUGH_REMOTE) != 0u)
        {
            // Nothing stays in the cache to be held for the commit, so the hint has no purpose here.
            stageAttributes = FILE_ATTRIBUTE_NORMAL;
            stageWriteFlags = FILE_FLAG_WRITE_THROUGH;
        }
        else if ((tuning & noTemporary) != 0u)
        {
            stageAttributes = FILE_ATTRIBUTE_NORMAL;
        }
    }

    const std::wstring extendedStage = FileSystemInternal::ToExtendedPath(stagePath);
    wil::unique_handle file(CreateFileW(extendedStage.c_str(),
                                        GENERIC_READ | GENERIC_WRITE | DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        CREATE_NEW,
                                        stageAttributes | stageWriteFlags | FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_SEQUENTIAL_SCAN,
                                        nullptr));
    if (! file)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    LocalObjectIdentityPayload identity{};
    const HRESULT identityHr = ResolveNewStageIdentity(file.get(), stagePath, L"file", retainedIdentity, identity);
    if (FAILED(identityHr))
    {
        return identityHr;
    }

    HANDLE duplicated = INVALID_HANDLE_VALUE;
    if (DuplicateHandle(GetCurrentProcess(), file.get(), GetCurrentProcess(), &duplicated, 0u, FALSE, DUPLICATE_SAME_ACCESS) == FALSE)
    {
        const DWORD error = GetLastError();
        // I35 R7.4: removing the stage this call just created has the owned-stage cleanup budget.
        const Common::SynchronousIoCancelWatch::CleanupWindow cleanupWindow(FreshOwnedStageCleanupDeadline());
        HRESULT cleanupHr = DeleteExactHandle(file.get());
        if (cleanupWindow.CanceledInsideWindow(cleanupHr))
        {
            cleanupHr = DeleteExactHandle(file.get());
        }
        return FAILED(cleanupHr) ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    wil::unique_handle boundHandle(duplicated);

    constexpr FileSystemBindFlags stageFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE |
                                         FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
    wil::com_ptr<IFileSystemBoundObject> stage;
    stage.attach(new (std::nothrow) LocalBoundObject(std::move(boundHandle), identity, FILESYSTEM_BOUND_REGULAR_FILE, 0u, stageFlags, true, stagePath));
    if (! stage)
    {
        const Common::SynchronousIoCancelWatch::CleanupWindow cleanupWindow(FreshOwnedStageCleanupDeadline());
        HRESULT cleanupHr = DeleteExactHandle(file.get());
        if (cleanupWindow.CanceledInsideWindow(cleanupHr))
        {
            cleanupHr = DeleteExactHandle(file.get());
        }
        return FAILED(cleanupHr) ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : E_OUTOFMEMORY;
    }

    auto* fileWriter = new (std::nothrow) Win32FileWriter(std::move(file), std::wstring{}, options);
    if (fileWriter == nullptr)
    {
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes                  = sizeof(abortResult);
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(options);
        const HRESULT abortHr                  = stage->AbortOwnedObject(&cleanupOptions, &abortResult);
        return abortResult.outcomeKnown == FALSE || FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE
                   ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)
                   : E_OUTOFMEMORY;
    }

    *writer     = fileWriter;
    *ownedStage = stage.detach();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateExclusiveDirectory(const wchar_t* stagePath,
                                                               const FileSystemOptions* options,
                                                               IFileSystemBoundObject** ownedStage) noexcept
{
    if (ownedStage == nullptr)
    {
        return E_POINTER;
    }
    *ownedStage = nullptr;
    if (stagePath == nullptr || stagePath[0] == L'\0' || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    LocalObjectIdentityPayload retainedIdentity{};
    const HRESULT retainedIdentityHr = MakeRetainedHandleIdentity(retainedIdentity);
    if (FAILED(retainedIdentityHr))
    {
        return retainedIdentityHr;
    }

    wil::unique_handle directory;
    HRESULT hr = CreateExclusiveDirectoryHandle(stagePath, directory);
    if (FAILED(hr))
    {
        return hr;
    }

    LocalObjectIdentityPayload identity{};
    hr = ResolveNewStageIdentity(directory.get(), stagePath, L"directory", retainedIdentity, identity);
    if (FAILED(hr))
    {
        return hr;
    }

    constexpr FileSystemBindFlags stageFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA |
                                                                                FILESYSTEM_BIND_DELETE | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
    auto* stage                              = new (std::nothrow)
        LocalBoundObject(std::move(directory), identity, FILESYSTEM_BOUND_DIRECTORY, std::numeric_limits<uint64_t>::max(), stageFlags, true, stagePath);
    if (stage == nullptr)
    {
        if (directory)
        {
            const Common::SynchronousIoCancelWatch::CleanupWindow cleanupWindow(FreshOwnedStageCleanupDeadline());
            HRESULT cleanupHr = DeleteExactHandle(directory.get());
            if (cleanupWindow.CanceledInsideWindow(cleanupHr))
            {
                cleanupHr = DeleteExactHandle(directory.get());
            }
            if (FAILED(cleanupHr))
            {
                return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
        }
        return E_OUTOFMEMORY;
    }
    *ownedStage = stage;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::ReadBoundLink(IFileSystemBoundObject* boundLink,
                                                    const FileSystemLinkTransform* transform,
                                                    const FileSystemOptions* options,
                                                    FileSystemLinkInformation* information) noexcept
{
    if (boundLink == nullptr || transform == nullptr || information == nullptr)
    {
        return E_POINTER;
    }
    if (transform->sizeBytes != sizeof(FileSystemLinkTransform) || information->sizeBytes != sizeof(FileSystemLinkInformation) ||
        ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    wil::com_ptr<ILocalBoundObjectControl> control;
    const HRESULT queryHr = boundLink->QueryInterface(__uuidof(ILocalBoundObjectControl), control.put_void());
    if (FAILED(queryHr) || ! control)
    {
        return FAILED(queryHr) ? queryHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HANDLE duplicated         = INVALID_HANDLE_VALUE;
    const HRESULT duplicateHr = control->DuplicateExactHandle(&duplicated);
    if (FAILED(duplicateHr))
    {
        return duplicateHr;
    }
    wil::unique_handle boundHandle(duplicated);
    return FileSystemInternal::ReadBoundLocalLink(boundHandle.get(), *transform, *information);
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateExclusiveLink(const wchar_t* stagePath,
                                                          const FileSystemLinkInformation* information,
                                                          const FileSystemOptions* options,
                                                          IFileSystemBoundObject** ownedStage) noexcept
{
    if (ownedStage == nullptr)
    {
        return E_POINTER;
    }
    *ownedStage = nullptr;
    if (stagePath == nullptr || stagePath[0] == L'\0' || information == nullptr)
    {
        return E_INVALIDARG;
    }
    if (information->sizeBytes != sizeof(FileSystemLinkInformation) || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    const HRESULT controlHr = FileSystemCheckOperationControl(options);
    if (FAILED(controlHr))
    {
        return controlHr;
    }

    LocalObjectIdentityPayload retainedIdentity{};
    const HRESULT retainedIdentityHr = MakeRetainedHandleIdentity(retainedIdentity);
    if (FAILED(retainedIdentityHr))
    {
        return retainedIdentityHr;
    }

    wil::unique_handle linkHandle;
    const HRESULT createHr = FileSystemInternal::CreateExclusiveLocalLink(stagePath, *information, linkHandle);
    if (FAILED(createHr))
    {
        return createHr;
    }

    LocalObjectIdentityPayload identity{};
    const HRESULT identityHr = ResolveNewStageIdentity(linkHandle.get(), stagePath, L"link", retainedIdentity, identity);
    if (FAILED(identityHr))
    {
        return identityHr;
    }

    constexpr FileSystemBindFlags stageFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA |
                                                                                FILESYSTEM_BIND_DELETE | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
    wil::com_ptr<IFileSystemBoundObject> stage;
    stage.attach(new (std::nothrow) LocalBoundObject(
        std::move(linkHandle), identity, FILESYSTEM_BOUND_LINK, std::numeric_limits<uint64_t>::max(), stageFlags, true, stagePath));
    if (! stage)
    {
        if (linkHandle)
        {
            const Common::SynchronousIoCancelWatch::CleanupWindow cleanupWindow(FreshOwnedStageCleanupDeadline());
            HRESULT cleanupHr = DeleteExactHandle(linkHandle.get());
            if (cleanupWindow.CanceledInsideWindow(cleanupHr))
            {
                cleanupHr = DeleteExactHandle(linkHandle.get());
            }
            if (FAILED(cleanupHr))
            {
                return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
        }
        return E_OUTOFMEMORY;
    }
    *ownedStage = stage.detach();
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
    const UINT driveType              = ! driveTypeRoot.empty() ? GetDriveTypeForStorageProbe(driveTypeRoot.c_str()) : DRIVE_UNKNOWN;
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

    const std::wstring driveRoot        = ExtractDriveRoot(path);
    const std::wstring volumeRoot       = ResolveLocalVolumeRootPath(path);
    const std::wstring& driveTypeRoot   = ! volumeRoot.empty() ? volumeRoot : driveRoot;
    const UINT driveType                = ! driveTypeRoot.empty() ? GetDriveTypeForStorageProbe(driveTypeRoot.c_str()) : DRIVE_UNKNOWN;
    const bool highLatency              = driveType == DRIVE_REMOTE || IsUncPath(path);
    const unsigned long callerSizeBytes = characteristics->sizeBytes;

    if (highLatency)
    {
        if (! volumeRoot.empty())
        {
            std::scoped_lock lock(_storageCharacteristicsMutex);
            _storageCharacteristicsCache.erase(OrdinalString::FoldCaseInvariant(volumeRoot));
        }
        FillStorageCharacteristicsLocal(*characteristics, {}, driveType, true);
        Debug::Perf::Emit(L"FileOps.Storage.ProbeCache", L"uncached-network", 0u, 0u, 1u, S_OK);
        return S_OK;
    }

    // Resolve once before the cache lookup. The returned openable path is both
    // the current physical identity and the exact target used by the probe.
    // If a direct volume GUID is unavailable, ResolveStorageProbeDevicePath
    // preserves the existing SUBST/backing-root and drive-letter fallbacks.
    const std::wstring volumePath = ResolveStorageProbeDevicePath(volumeRoot, driveRoot);
    if (! IsStableVolumeGuidDevicePath(volumePath) || volumeRoot.empty())
    {
        {
            std::scoped_lock lock(_storageCharacteristicsMutex);
            _storageCharacteristicsCache.erase(OrdinalString::FoldCaseInvariant(volumeRoot));
            FillStorageCharacteristicsLocal(*characteristics, volumePath, driveType, false);
        }
        Debug::Perf::Emit(L"FileOps.Storage.ProbeCache", L"uncached-unstable-identity", 0u, 0u, 1u, S_OK);
        return S_OK;
    }

    const std::wstring cacheKey = OrdinalString::FoldCaseInvariant(volumeRoot);
    {
        // Hold the per-instance lock through the first physical probe so concurrent operation-start
        // queries cannot all issue the same volume IOCTLs.
        std::scoped_lock lock(_storageCharacteristicsMutex);
        if (const auto cached = _storageCharacteristicsCache.find(cacheKey);
            cached != _storageCharacteristicsCache.end() && OrdinalString::EqualsNoCase(cached->second.backingVolumeGuidDevicePath, volumePath) &&
            OrdinalString::EqualsNoCase(cached->second.driveRoot, driveRoot) && cached->second.driveType == driveType && ! cached->second.highLatency)
        {
            *characteristics           = cached->second.characteristics;
            characteristics->sizeBytes = callerSizeBytes;
            Debug::Perf::Emit(L"FileOps.Storage.ProbeCache", L"hit", 0u, 1u, 0u, S_OK);
            return S_OK;
        }

        FillStorageCharacteristicsLocal(*characteristics, volumePath, driveType, false);
        StorageCharacteristicsCacheEntry cached{};
        cached.backingVolumeGuidDevicePath = volumePath;
        cached.driveRoot                   = driveRoot;
        cached.driveType                   = driveType;
        cached.highLatency                 = false;
        cached.characteristics             = *characteristics;
        cached.characteristics.sizeBytes   = sizeof(FileSystemStorageCharacteristics);
        _storageCharacteristicsCache.insert_or_assign(cacheKey, std::move(cached));
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
        parsed              = Common::Json::ParseObjectDocument(sourceConfiguration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root               = parsed.root;
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
                deleteRecycleBinMaxConcurrency = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxDeleteRecycleBinMaxConcurrency)));
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
        if (reparsePolicyVal)
        {
            if (! yyjson_is_str(reparsePolicyVal))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            const char* valueText = yyjson_get_str(reparsePolicyVal);
            if (! valueText || valueText[0] == '\0' || ! TryParseReparsePointPolicy(valueText, reparsePointPolicy))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
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

    std::string newConfigurationJson;
    if (parsed)
    {
        const HRESULT canonicalizeHr = CanonicalizeReparsePointPolicyConfiguration(parsed, reparsePointPolicy, newConfigurationJson);
        if (FAILED(canonicalizeHr))
        {
            return canonicalizeHr;
        }
    }
    else
    {
        newConfigurationJson = BuildConfigurationJson(concurrencyMode,
                                                      copyMoveMaxConcurrency,
                                                      deleteMaxConcurrency,
                                                      deleteRecycleBinMaxConcurrency,
                                                      recycleBinBatchSize,
                                                      enumerationSoftMaxBufferMiB,
                                                      enumerationHardMaxBufferMiB,
                                                      reparsePointPolicy,
                                                      searchBackendPreference,
                                                      searchMaxDirectoryWalkers);
    }

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
    _capabilitiesJson.clear();
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

void FileSystem::UpdateCapabilitiesJson(std::string_view rootId, bool remoteRoute) noexcept
{
    // NOTE: Caller must hold _stateMutex.
    _capabilitiesJson = std::format(
        R"json({{
  "version": 2,
  "pathProfile": "{}",
  "rootId": "{}",
  "operations": {{"copy":true,"move":true,"nativeMove":true,"delete":true,"rename":true,"createDirectory":true,"properties":true,"read":true,"write":true,"recycle":true}},
  "search": {{"version":1,"name":true,"content":true,"indexed":true,"serviceBacked":true,"supportsRegex":true,"supportsSnippets":true,"preferredBackend":"service"}},
  "concurrency": {{"copyMoveMax":{},"deleteMax":{},"deleteRecycleBinMax":{}}},
  "transfer": {{"export":{{"copy":["*"],"move":["*"]}},"import":{{"copy":["*"],"move":["*"]}}}},
  "identity": {{"object":"win32FileId","revision":"none","boundDelete":true,"conditionalDelete":true}},
  "publication": {{"exclusiveStage":true,"conditionalPublish":true,"committedSize":true}},
  "links": {{"preserveFileLink":true,"preserveDirectoryLink":true,"retargetInTree":false,"exactLinkRemoval":true}},
  "metadata": {{"motw":"preserved","alternateStreams":"preserved","extendedAttributes":"preserved","sparse":"preserved","efs":"preserved"}},
  "verification": {{"hostReadback":true,"providerProof":"blake3-bound-object"}},
  "cancellation": {{"abort":false,"deadline":false,"routeClass":"{}","providerWatchdogTimeoutMs":0}},
  "names": {{"pathTextStableIdentity":true,"comparison":"ordinalIgnoreCase","normalization":"none","preferredSeparator":"\\","acceptedSeparators":["\\","/"],"casePreserving":true,"caseOnlyRename":"supported","maxComponentUtf16":255}},
  "directories": {{"model":"native"}}
}})json",
        remoteRoute ? "local-win32-smb" : "local-win32",
        rootId,
        std::clamp(_copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency),
        std::clamp(_deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency),
        std::clamp(_deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency),
        "bounded");
}

#include "FileSystem.Internal.h"

#include <algorithm>

#include <aclapi.h>

#pragma comment(lib, "Advapi32.lib")

namespace FileSystemInternal
{
namespace
{
void NormalizePathSeparators(std::wstring& path) noexcept
{
    std::ranges::replace(path, L'/', L'\\');
}

[[nodiscard]] bool IsMissingPathError(DWORD error) noexcept
{
    return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_NOT_FOUND;
}

[[nodiscard]] bool IsDriveExtendedPath(std::wstring_view path) noexcept
{
    return path.size() >= 7u && path.rfind(L"\\\\?\\", 0) == 0 && path[5] == L':' &&
           ((path[4] >= L'A' && path[4] <= L'Z') || (path[4] >= L'a' && path[4] <= L'z')) && (path[6] == L'\\' || path[6] == L'/');
}

[[nodiscard]] bool IsUncExtendedPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\?\\UNC\\", 0) == 0;
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] bool TryGetSupportedExtendedRootLength(std::wstring_view path, size_t& rootLength) noexcept
{
    rootLength = 0u;
    if (IsDriveExtendedPath(path))
    {
        rootLength = 7u;
        return true;
    }

    if (! IsUncExtendedPath(path))
    {
        return false;
    }

    constexpr size_t kExtendedUncPrefixLength = 8u;
    const size_t serverEnd                    = path.find(L'\\', kExtendedUncPrefixLength);
    if (serverEnd == std::wstring_view::npos || serverEnd == kExtendedUncPrefixLength)
    {
        return false;
    }

    const size_t shareStart = serverEnd + 1u;
    if (shareStart >= path.size())
    {
        return false;
    }

    const size_t shareEnd = path.find(L'\\', shareStart);
    if (shareEnd == shareStart)
    {
        return false;
    }

    rootLength = shareEnd == std::wstring_view::npos ? path.size() : shareEnd + 1u;
    return true;
}

std::wstring NormalizeSupportedExtendedPathLexically(std::wstring_view path)
{
    size_t rootLength = 0u;
    if (! TryGetSupportedExtendedRootLength(path, rootLength))
    {
        return std::wstring(path);
    }

    std::vector<std::wstring_view> components;
    bool preserveTrailingSeparator = false;
    size_t position                = rootLength;
    while (position < path.size())
    {
        while (position < path.size() && IsPathSeparator(path[position]))
        {
            ++position;
        }
        if (position >= path.size())
        {
            preserveTrailingSeparator = ! components.empty();
            break;
        }

        const size_t end                  = path.find(L'\\', position);
        const std::wstring_view component = end == std::wstring_view::npos ? path.substr(position) : path.substr(position, end - position);
        if (component == L"..")
        {
            if (! components.empty())
            {
                components.pop_back();
            }
        }
        else if (component != L".")
        {
            components.push_back(component);
        }

        if (end == std::wstring_view::npos)
        {
            break;
        }

        preserveTrailingSeparator = end + 1u == path.size();
        position                  = end + 1u;
    }

    std::wstring normalized(path.substr(0u, rootLength));
    if (components.empty())
    {
        return normalized;
    }

    for (const std::wstring_view component : components)
    {
        if (! normalized.empty() && ! IsPathSeparator(normalized.back()))
        {
            normalized.push_back(L'\\');
        }
        normalized.append(component);
    }
    if (preserveTrailingSeparator && ! IsPathSeparator(normalized.back()))
    {
        normalized.push_back(L'\\');
    }

    return normalized;
}

[[nodiscard]] DWORD NormalizeAttributesForSetFileAttributes(DWORD attributes) noexcept
{
    if (attributes == 0u || attributes == INVALID_FILE_ATTRIBUTES)
    {
        return FILE_ATTRIBUTE_NORMAL;
    }

    if ((attributes & FILE_ATTRIBUTE_NORMAL) != 0u && attributes != FILE_ATTRIBUTE_NORMAL)
    {
        attributes &= ~FILE_ATTRIBUTE_NORMAL;
    }

    return attributes == 0u ? FILE_ATTRIBUTE_NORMAL : attributes;
}

#if defined(_DEBUG)
constexpr std::wstring_view kFinalAttributesFailPathEnvVar  = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_PATH";
constexpr std::wstring_view kFinalAttributesFailFiredEnvVar = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_FIRED";

[[nodiscard]] bool ShouldFailFinalAttributesForSelfTest(const std::wstring& finalPath) noexcept
{
    const DWORD required = ::GetEnvironmentVariableW(kFinalAttributesFailPathEnvVar.data(), nullptr, 0u);
    if (required == 0u)
    {
        return false;
    }

    std::wstring configured(static_cast<size_t>(required), L'\0');
    const DWORD written = ::GetEnvironmentVariableW(kFinalAttributesFailPathEnvVar.data(), configured.data(), required);
    if (written == 0u || written >= required)
    {
        return false;
    }
    configured.resize(static_cast<size_t>(written));
    configured = ToExtendedPath(configured);

    if (CompareStringOrdinal(configured.c_str(), -1, finalPath.c_str(), -1, TRUE) != CSTR_EQUAL)
    {
        return false;
    }

    static_cast<void>(::SetEnvironmentVariableW(kFinalAttributesFailPathEnvVar.data(), nullptr));
    static_cast<void>(::SetEnvironmentVariableW(kFinalAttributesFailFiredEnvVar.data(), L"1"));
    return true;
}
#endif

struct FileMetadataSnapshot final
{
    FileMetadataSnapshot()                                       = default;
    FileMetadataSnapshot(const FileMetadataSnapshot&)            = delete;
    FileMetadataSnapshot& operator=(const FileMetadataSnapshot&) = delete;
    FileMetadataSnapshot(FileMetadataSnapshot&&)                 = default;
    FileMetadataSnapshot& operator=(FileMetadataSnapshot&&)      = default;

    bool capturedTimes = false;
    FILETIME creationTime{};
    FILETIME lastAccessTime{};
    FILETIME lastWriteTime{};

    wil::unique_hlocal_ptr<void> securityDescriptor;
};

[[nodiscard]] FileMetadataSnapshot CaptureFileMetadataForFallback(const std::wstring& path) noexcept
{
    FileMetadataSnapshot snapshot{};

    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT,
                                          nullptr));
    if (file)
    {
        snapshot.capturedTimes = ::GetFileTime(file.get(), &snapshot.creationTime, &snapshot.lastAccessTime, &snapshot.lastWriteTime) != 0;
    }

    PACL dacl                               = nullptr;
    PSECURITY_DESCRIPTOR securityDescriptor = nullptr;
    const DWORD securityError               = ::GetNamedSecurityInfoW(
        const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, &dacl, nullptr, &securityDescriptor);
    if (securityError == ERROR_SUCCESS && securityDescriptor != nullptr)
    {
        snapshot.securityDescriptor.reset(securityDescriptor);
    }

    return snapshot;
}

void RestoreFileMetadataAfterFallback(const std::wstring& path, const FileMetadataSnapshot& snapshot) noexcept
{
    if (snapshot.securityDescriptor)
    {
        static_cast<void>(::SetFileSecurityW(path.c_str(), DACL_SECURITY_INFORMATION, snapshot.securityDescriptor.get()));
    }

    if (! snapshot.capturedTimes)
    {
        return;
    }

    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT,
                                          nullptr));
    if (file)
    {
        static_cast<void>(::SetFileTime(file.get(), &snapshot.creationTime, nullptr, nullptr));
    }
}
} // namespace

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept
{
    return (name == L"." || name == L"..");
}

std::wstring MakeAbsolutePath(const std::wstring& path)
{
    std::wstring input = path;
    NormalizePathSeparators(input);
    if (input.empty())
    {
        input = L".";
    }

    if (IsDriveExtendedPath(input) || IsUncExtendedPath(input))
    {
        return NormalizeSupportedExtendedPathLexically(input);
    }

    if (input.rfind(L"\\\\?\\", 0) == 0)
    {
        return input;
    }

    DWORD required = GetFullPathNameW(input.c_str(), 0, nullptr, nullptr);
    if (required == 0)
    {
        return input;
    }

    std::wstring absolute(static_cast<size_t>(required) + 1, L'\0');
    DWORD written = GetFullPathNameW(input.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
    if (written == 0)
    {
        return input;
    }

    absolute.resize(static_cast<size_t>(written));
    return absolute;
}

std::wstring ToExtendedPath(const std::wstring& path)
{
    std::wstring normalized = path;
    NormalizePathSeparators(normalized);
    if (normalized.empty())
    {
        normalized = L".";
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0)
    {
        if (IsDriveExtendedPath(normalized) || IsUncExtendedPath(normalized))
        {
            return MakeAbsolutePath(normalized);
        }
        return normalized;
    }

    normalized = MakeAbsolutePath(normalized);
    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return L"\\\\?\\UNC\\" + normalized.substr(2);
    }

    return L"\\\\?\\" + normalized;
}

#if defined(_DEBUG)
void RunDebugPathNormalizationSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool
    {
        if (condition)
        {
            ++passed;
            return true;
        }

        ++failed;
        Debug::Error(L"FileSystem debug selftest failed: {}", message);
        return false;
    };

    const DWORD requiredCurrentDirectoryChars = ::GetCurrentDirectoryW(0, nullptr);
    if (! check(requiredCurrentDirectoryChars > 0u, L"path normalization selftest should read the current directory"))
    {
        return;
    }

    std::wstring currentDirectory(static_cast<size_t>(requiredCurrentDirectoryChars), L'\0');
    const DWORD writtenCurrentDirectoryChars = ::GetCurrentDirectoryW(static_cast<DWORD>(currentDirectory.size()), currentDirectory.data());
    if (! check(writtenCurrentDirectoryChars > 0u, L"path normalization selftest should copy the current directory"))
    {
        return;
    }
    currentDirectory.resize(static_cast<size_t>(writtenCurrentDirectoryChars));
    NormalizePathSeparators(currentDirectory);

    if (! check(currentDirectory.size() >= 3u && currentDirectory[1] == L':' && currentDirectory[2] == L'\\',
                L"path normalization selftest requires a drive-rooted current directory"))
    {
        return;
    }

    const std::wstring driveRoot             = currentDirectory.substr(0u, 3u);
    const std::wstring drivePathInput        = driveRoot + L"RedSalamanderPathSelfTest\\child\\..\\leaf";
    const std::wstring expectedDriveAbsolute = driveRoot + L"RedSalamanderPathSelfTest\\leaf";
    const std::wstring extendedDriveInput    = L"\\\\?\\" + drivePathInput;
    const std::wstring expectedExtendedDrive = L"\\\\?\\" + expectedDriveAbsolute;

    check(MakeAbsolutePath(extendedDriveInput) == expectedExtendedDrive, L"MakeAbsolutePath should collapse dot segments inside drive-rooted extended paths");
    check(ToExtendedPath(extendedDriveInput) == expectedExtendedDrive, L"ToExtendedPath should collapse dot segments inside drive-rooted extended paths");
    check(ToExtendedPath(drivePathInput) == expectedExtendedDrive,
          L"ToExtendedPath should still normalize ordinary drive-rooted paths before adding the extended prefix");

    const std::wstring trailingDotInput    = extendedDriveInput + L".";
    const std::wstring expectedTrailingDot = expectedExtendedDrive + L".";
    check(MakeAbsolutePath(trailingDotInput) == expectedTrailingDot, L"MakeAbsolutePath should preserve a literal trailing dot in drive-rooted extended paths");
    check(ToExtendedPath(trailingDotInput) == expectedTrailingDot, L"ToExtendedPath should preserve a literal trailing dot in drive-rooted extended paths");

    const std::wstring trailingSpaceInput    = extendedDriveInput + L" ";
    const std::wstring expectedTrailingSpace = expectedExtendedDrive + L" ";
    check(MakeAbsolutePath(trailingSpaceInput) == expectedTrailingSpace,
          L"MakeAbsolutePath should preserve a literal trailing space in drive-rooted extended paths");
    check(ToExtendedPath(trailingSpaceInput) == expectedTrailingSpace,
          L"ToExtendedPath should preserve a literal trailing space in drive-rooted extended paths");

    constexpr std::wstring_view kExtendedUncInput    = L"\\\\?\\UNC\\server\\share\\folder\\..\\leaf";
    constexpr std::wstring_view kExpectedExtendedUnc = L"\\\\?\\UNC\\server\\share\\leaf";
    check(MakeAbsolutePath(std::wstring(kExtendedUncInput)) == kExpectedExtendedUnc,
          L"MakeAbsolutePath should collapse dot segments inside supported extended UNC paths");
    check(ToExtendedPath(std::wstring(kExtendedUncInput)) == kExpectedExtendedUnc,
          L"ToExtendedPath should collapse dot segments inside supported extended UNC paths");

    constexpr std::wstring_view kDeviceNamespaceInput = L"\\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy1\\folder\\..\\leaf";
    check(MakeAbsolutePath(std::wstring(kDeviceNamespaceInput)) == kDeviceNamespaceInput,
          L"MakeAbsolutePath should preserve unsupported device namespace paths instead of treating them as local filesystem paths");
    check(ToExtendedPath(std::wstring(kDeviceNamespaceInput)) == kDeviceNamespaceInput,
          L"ToExtendedPath should preserve unsupported device namespace paths instead of treating them as local filesystem paths");
}
#endif

bool TryGetUncServerRoot(std::wstring_view path, std::wstring& serverName) noexcept
{
    serverName.clear();

    size_t start = 0;
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        start = 8;
    }
    else if (path.rfind(L"\\\\", 0) == 0 && path.rfind(L"\\\\?\\", 0) != 0)
    {
        start = 2;
    }
    else
    {
        return false;
    }

    const size_t end = path.find_first_of(L"\\/", start);
    if (end == std::wstring_view::npos || end <= start)
    {
        if (start < path.size())
        {
            serverName.assign(path.substr(start));
            return ! serverName.empty();
        }
        return false;
    }

    size_t pos = end;
    while (pos < path.size() && (path[pos] == L'\\' || path[pos] == L'/'))
    {
        ++pos;
    }

    // Any non-separator text after the server component indicates this is a share path, not a server root.
    if (pos < path.size())
    {
        return false;
    }

    serverName.assign(path.substr(start, end - start));
    return ! serverName.empty();
}

std::wstring AppendPath(const std::wstring& base, std::wstring_view leaf)
{
    if (base.empty())
    {
        return std::wstring(leaf);
    }

    if (leaf.empty())
    {
        return base;
    }

    std::wstring result = base;
    if (const wchar_t last = result.back(); last != L'\\' && last != L'/')
    {
        result.push_back(L'\\');
    }
    result.append(leaf);
    return result;
}

std::wstring AppendPath(const std::wstring& base, const wchar_t* leaf)
{
    if (! leaf)
    {
        return base;
    }
    return AppendPath(base, std::wstring_view(leaf));
}

std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept
{
    while (! path.empty())
    {
        const wchar_t last = path.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

std::wstring_view GetPathLeaf(std::wstring_view path) noexcept
{
    const std::wstring_view trimmed = TrimTrailingSeparators(path);
    if (trimmed.empty())
    {
        return trimmed;
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed;
    }

    return trimmed.substr(pos + 1);
}

std::wstring GetPathDirectory(std::wstring_view path)
{
    const std::wstring_view trimmed = TrimTrailingSeparators(path);
    if (trimmed.empty())
    {
        return {};
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return {};
    }

    return std::wstring(trimmed.substr(0, pos));
}

[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept
{
    return text.find_first_of(L"\\/") != std::wstring_view::npos;
}

PathInfo MakePathInfo(const std::wstring& path)
{
    PathInfo info{};
    info.display = path;
    NormalizePathSeparators(info.display);
    info.extended = ToExtendedPath(path);
    return info;
}

PathInfo MakePathInfo(const wchar_t* path)
{
    PathInfo info{};
    if (path)
    {
        info.display = path;
        NormalizePathSeparators(info.display);
        info.extended = ToExtendedPath(info.display);
    }
    return info;
}

HRESULT PromoteStagedTempIntoFinalPath(const std::wstring& tempPath, const std::wstring& finalPath, const StagedPromotionOptions& options) noexcept
{
    if (tempPath.empty() || finalPath.empty())
    {
        return E_INVALIDARG;
    }

    const DWORD tempAttributes = ::GetFileAttributesW(tempPath.c_str());
    if (tempAttributes == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD tempError = ::GetLastError();
        return HRESULT_FROM_WIN32(tempError != 0u ? tempError : ERROR_FILE_NOT_FOUND);
    }

    DWORD desiredFinalAttributes = tempAttributes;
    if (options.stripTemporaryAttributes)
    {
        desiredFinalAttributes &= ~(FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY);
    }
    if (! options.preserveReplacementReadOnly)
    {
        desiredFinalAttributes &= ~FILE_ATTRIBUTE_READONLY;
    }
    desiredFinalAttributes = NormalizeAttributesForSetFileAttributes(desiredFinalAttributes);

    DWORD tempPromotionAttributes = tempAttributes & ~FILE_ATTRIBUTE_READONLY;
    if (options.stripTemporaryAttributes)
    {
        tempPromotionAttributes &= ~(FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY);
    }
    tempPromotionAttributes = NormalizeAttributesForSetFileAttributes(tempPromotionAttributes);
    if (tempPromotionAttributes != tempAttributes && ::SetFileAttributesW(tempPath.c_str(), tempPromotionAttributes) == 0)
    {
        const DWORD error = ::GetLastError();
        return HRESULT_FROM_WIN32(error != 0u ? error : ERROR_ACCESS_DENIED);
    }

    const DWORD destinationAttributes = ::GetFileAttributesW(finalPath.c_str());
    const DWORD destinationError      = destinationAttributes == INVALID_FILE_ATTRIBUTES ? ::GetLastError() : ERROR_SUCCESS;
    if (destinationAttributes == INVALID_FILE_ATTRIBUTES && ! IsMissingPathError(destinationError))
    {
        return HRESULT_FROM_WIN32(destinationError != 0u ? destinationError : ERROR_GEN_FAILURE);
    }

    const bool destinationExists = destinationAttributes != INVALID_FILE_ATTRIBUTES;
    if (destinationExists && (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    bool clearedReadOnly = false;
    if (destinationExists && (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0u)
    {
        if (! options.allowReplaceReadOnly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        if (::SetFileAttributesW(finalPath.c_str(), destinationAttributes & ~FILE_ATTRIBUTE_READONLY) == 0)
        {
            const DWORD error = ::GetLastError();
            return HRESULT_FROM_WIN32(error != 0u ? error : ERROR_ACCESS_DENIED);
        }
        clearedReadOnly = true;
    }

    bool committed                  = false;
    auto restoreDestinationReadOnly = wil::scope_exit([&]() noexcept
    {
        if (clearedReadOnly && ! committed)
        {
            static_cast<void>(::SetFileAttributesW(finalPath.c_str(), destinationAttributes));
        }
    });

    FileMetadataSnapshot fallbackMetadata{};
    if (destinationExists && ! options.replacementIsReparsePoint)
    {
        fallbackMetadata = CaptureFileMetadataForFallback(finalPath);
    }
    bool usedMoveFallback = false;

    if (destinationExists && ! options.replacementIsReparsePoint)
    {
        const DWORD replaceFlags = options.ignoreReplaceMergeErrors ? REPLACEFILE_IGNORE_MERGE_ERRORS : 0u;
        if (::ReplaceFileW(finalPath.c_str(), tempPath.c_str(), nullptr, replaceFlags, nullptr, nullptr) == 0)
        {
            const DWORD replaceError = ::GetLastError();
            if (::MoveFileExW(tempPath.c_str(), finalPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
            {
                const DWORD moveError = ::GetLastError();
                return HRESULT_FROM_WIN32(moveError != 0u ? moveError : replaceError);
            }
            usedMoveFallback = true;
        }
    }
    else if (::MoveFileExW(tempPath.c_str(), finalPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
    {
        const DWORD moveError = ::GetLastError();
        return HRESULT_FROM_WIN32(moveError != 0u ? moveError : ERROR_ACCESS_DENIED);
    }

    committed = true;

    if (usedMoveFallback)
    {
        RestoreFileMetadataAfterFallback(finalPath, fallbackMetadata);
    }

    const DWORD finalAttributes = ::GetFileAttributesW(finalPath.c_str());
    const bool attributesNeedUpdate =
        finalAttributes == INVALID_FILE_ATTRIBUTES || NormalizeAttributesForSetFileAttributes(finalAttributes) != desiredFinalAttributes;
    if (attributesNeedUpdate)
    {
#if defined(_DEBUG)
        const bool injectedFailure = ShouldFailFinalAttributesForSelfTest(finalPath);
#else
        constexpr bool injectedFailure = false;
#endif
        if (injectedFailure || ::SetFileAttributesW(finalPath.c_str(), desiredFinalAttributes) == 0)
        {
            const DWORD error = injectedFailure ? ERROR_ACCESS_DENIED : ::GetLastError();
            Debug::Warning(L"FileSystem: promoted '{}' but failed to set final attributes (error={}).", finalPath, error);
            return HRESULT_FROM_WIN32(error != 0u ? error : ERROR_ACCESS_DENIED);
        }
    }

    return S_OK;
}
} // namespace FileSystemInternal

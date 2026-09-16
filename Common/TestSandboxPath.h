#pragma once

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::Testing
{
inline constexpr std::wstring_view kTestSandboxMarkerFileName{L".red-salamander-test-root"};
inline constexpr std::string_view kTestSandboxMarkerContents{"red-salamander.test-sandbox-root.v1"};
inline constexpr std::wstring_view kTestSandboxDirectoryName{L"RedSalamander.Perf"};

[[nodiscard]] inline bool TestSandboxPathEquals(const std::filesystem::path& left, const std::filesystem::path& right) noexcept
{
    return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

[[nodiscard]] inline bool HasTestSandboxAlternateDataStreamSyntax(const std::filesystem::path& path) noexcept
{
    const std::wstring text = path.native();
    const size_t firstColon = text.find(L':');
    if (firstColon == std::wstring::npos)
    {
        return false;
    }
    return firstColon != 1u || text.find(L':', firstColon + 1u) != std::wstring::npos;
}

[[nodiscard]] inline bool IsSameOrDescendantTestSandboxPath(const std::filesystem::path& candidate, const std::filesystem::path& parent) noexcept
{
    if (candidate.empty() || parent.empty())
    {
        return false;
    }

    std::wstring candidateText = candidate.lexically_normal().native();
    std::wstring parentText    = parent.lexically_normal().native();
    while (candidateText.size() > 3u && (candidateText.back() == L'\\' || candidateText.back() == L'/'))
    {
        candidateText.pop_back();
    }
    while (parentText.size() > 3u && (parentText.back() == L'\\' || parentText.back() == L'/'))
    {
        parentText.pop_back();
    }

    if (_wcsicmp(candidateText.c_str(), parentText.c_str()) == 0)
    {
        return true;
    }
    if (candidateText.size() <= parentText.size() || _wcsnicmp(candidateText.c_str(), parentText.c_str(), parentText.size()) != 0)
    {
        return false;
    }
    return candidateText[parentText.size()] == L'\\' || candidateText[parentText.size()] == L'/';
}

[[nodiscard]] inline std::filesystem::path NormalizeAbsoluteTestSandboxPath(const std::filesystem::path& path, std::error_code& ec) noexcept
{
    ec.clear();
    if (path.empty())
    {
        ec = std::make_error_code(std::errc::invalid_argument);
        return {};
    }

    std::filesystem::path absolute = path;
    if (! absolute.is_absolute())
    {
        absolute = std::filesystem::absolute(path, ec);
        if (ec)
        {
            return {};
        }
    }
    return absolute.lexically_normal();
}

[[nodiscard]] inline bool IsExistingTestSandboxPathReparseFree(const std::filesystem::path& path, std::error_code& ec) noexcept
{
    ec.clear();
    const std::filesystem::path normalized = NormalizeAbsoluteTestSandboxPath(path, ec);
    if (ec)
    {
        return false;
    }

    std::filesystem::path current = normalized.root_path();
    for (const std::filesystem::path& component : normalized.relative_path())
    {
        current /= component;
        const DWORD attributes = GetFileAttributesW(current.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
            {
                continue;
            }
            ec = std::error_code(static_cast<int>(error), std::system_category());
            return false;
        }
        if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
        {
            ec = std::make_error_code(std::errc::permission_denied);
            return false;
        }
    }
    return true;
}

[[nodiscard]] inline bool IsRepositoryRootCandidate(const std::filesystem::path& candidate) noexcept
{
    std::error_code ec;
    if (std::filesystem::exists(candidate / L".git", ec) && ! ec)
    {
        return true;
    }
    ec.clear();
    if (std::filesystem::exists(candidate / L"RedSalamander.sln", ec) && ! ec)
    {
        return true;
    }
    ec.clear();
    return std::filesystem::exists(candidate / L"RedSalamander" / L"RedSalamander.vcxproj", ec) && ! ec;
}

[[nodiscard]] inline std::filesystem::path FindRepositoryRootFrom(std::filesystem::path current) noexcept
{
    current = current.lexically_normal();
    for (size_t depth = 0u; depth < 16u && ! current.empty(); ++depth)
    {
        if (IsRepositoryRootCandidate(current))
        {
            return current;
        }
        const std::filesystem::path parent = current.parent_path();
        if (parent.empty() || parent == current)
        {
            break;
        }
        current = parent;
    }
    return {};
}

[[nodiscard]] inline std::filesystem::path GetCurrentModulePath(std::error_code& ec) noexcept
{
    ec.clear();
    std::vector<wchar_t> buffer(1024u);
    for (;;)
    {
        SetLastError(ERROR_SUCCESS);
        const DWORD written = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
        if (written == 0u)
        {
            ec = std::error_code(static_cast<int>(GetLastError()), std::system_category());
            return {};
        }
        if (written < buffer.size())
        {
            return std::filesystem::path(std::wstring(buffer.data(), written));
        }
        if (buffer.size() >= 32768u)
        {
            ec = std::make_error_code(std::errc::filename_too_long);
            return {};
        }
        buffer.resize((std::min)(buffer.size() * 2u, size_t{32768u}));
    }
}

[[nodiscard]] inline std::filesystem::path FindTestRepositoryRoot(std::error_code& ec) noexcept
{
    ec.clear();
    const std::filesystem::path modulePath = GetCurrentModulePath(ec);
    if (! ec && ! modulePath.empty())
    {
        const std::filesystem::path moduleRoot = FindRepositoryRootFrom(modulePath.parent_path());
        if (! moduleRoot.empty())
        {
            return moduleRoot;
        }
    }

    ec.clear();
    const std::filesystem::path current = std::filesystem::current_path(ec);
    if (ec)
    {
        return {};
    }
    const std::filesystem::path currentRoot = FindRepositoryRootFrom(current);
    if (currentRoot.empty())
    {
        ec = std::make_error_code(std::errc::no_such_file_or_directory);
    }
    return currentRoot;
}

[[nodiscard]] inline bool HasValidTestSandboxMarker(const std::filesystem::path& sandboxBase) noexcept
{
    const std::filesystem::path marker = sandboxBase / std::wstring(kTestSandboxMarkerFileName);
    const DWORD attributes             = GetFileAttributesW(marker.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        return false;
    }

    std::ifstream input(marker, std::ios::binary);
    if (! input)
    {
        return false;
    }
    std::string contents((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    if (contents.size() > 128u)
    {
        return false;
    }
    while (! contents.empty() && (contents.back() == '\r' || contents.back() == '\n'))
    {
        contents.pop_back();
    }
    return contents == kTestSandboxMarkerContents;
}

[[nodiscard]] inline bool IsFixedLocalDriveRoot(const std::filesystem::path& root) noexcept
{
    return ! root.empty() && GetDriveTypeW(root.c_str()) == DRIVE_FIXED;
}

[[nodiscard]] inline std::optional<std::filesystem::path> GetDedicatedExternalTestSandboxBase(const std::filesystem::path& candidate,
                                                                                              std::wstring_view expectedDirectoryName) noexcept
{
    if (candidate.empty() || expectedDirectoryName.empty())
    {
        return std::nullopt;
    }

    const std::filesystem::path normalized = candidate.lexically_normal();
    const std::filesystem::path driveRoot  = normalized.root_path();
    if (driveRoot.empty() || ! IsFixedLocalDriveRoot(driveRoot))
    {
        return std::nullopt;
    }

    const std::filesystem::path base = driveRoot / std::wstring(expectedDirectoryName);
    if (! IsSameOrDescendantTestSandboxPath(normalized, base))
    {
        return std::nullopt;
    }
    return base.lexically_normal();
}

[[nodiscard]] inline bool IsAuthorizedTestSandboxPath(const std::filesystem::path& candidate,
                                                      const std::filesystem::path& repositoryRoot,
                                                      bool requireExactExternalBase,
                                                      std::error_code& ec) noexcept
{
    ec.clear();
    const std::filesystem::path normalizedCandidate = NormalizeAbsoluteTestSandboxPath(candidate, ec);
    if (ec || HasTestSandboxAlternateDataStreamSyntax(normalizedCandidate))
    {
        if (! ec)
        {
            ec = std::make_error_code(std::errc::invalid_argument);
        }
        return false;
    }
    // Keep the repository parameter in the shared signature while callers migrate;
    // authorization is intentionally independent of repository placement.
    static_cast<void>(repositoryRoot);

    const std::optional<std::filesystem::path> externalBase = GetDedicatedExternalTestSandboxBase(normalizedCandidate, kTestSandboxDirectoryName);
    if (! externalBase.has_value() || (requireExactExternalBase && ! TestSandboxPathEquals(normalizedCandidate, externalBase.value())) ||
        ! HasValidTestSandboxMarker(externalBase.value()))
    {
        ec = std::make_error_code(std::errc::permission_denied);
        return false;
    }
    return IsExistingTestSandboxPathReparseFree(normalizedCandidate, ec);
}
} // namespace Common::Testing

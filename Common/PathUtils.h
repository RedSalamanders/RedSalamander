#pragma once

#include <array>
#include <cstddef>
#include <limits>
#include <string>
#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <objbase.h>

namespace Common::Paths
{
enum class WindowsPathClass
{
    Relative,
    Rooted,
    DriveRelative,
    DriveAbsolute,
    Unc,
    ExtendedDriveAbsolute,
    ExtendedUnc,
    ExtendedOther,
    Device,
};

[[nodiscard]] inline bool IsSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] inline bool IsAsciiDriveLetter(wchar_t ch) noexcept
{
    return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
}

[[nodiscard]] inline bool StartsWithAsciiNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (text.size() < prefix.size())
    {
        return false;
    }

    for (size_t index = 0u; index < prefix.size(); ++index)
    {
        wchar_t actual = text[index];
        wchar_t wanted = prefix[index];
        if (actual >= L'a' && actual <= L'z')
        {
            actual = static_cast<wchar_t>(actual - (L'a' - L'A'));
        }
        if (wanted >= L'a' && wanted <= L'z')
        {
            wanted = static_cast<wchar_t>(wanted - (L'a' - L'A'));
        }
        if (actual != wanted)
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] inline WindowsPathClass ClassifyWindowsPath(std::wstring_view path) noexcept
{
    if (path.size() >= 4u && IsSeparator(path[0]) && IsSeparator(path[1]) && path[2] == L'.' && IsSeparator(path[3]))
    {
        return WindowsPathClass::Device;
    }

    if (path.size() >= 4u && IsSeparator(path[0]) && IsSeparator(path[1]) && path[2] == L'?' && IsSeparator(path[3]))
    {
        const std::wstring_view remainder = path.substr(4u);
        if (remainder.size() >= 3u && IsAsciiDriveLetter(remainder[0]) && remainder[1] == L':' && IsSeparator(remainder[2]))
        {
            return WindowsPathClass::ExtendedDriveAbsolute;
        }
        if (remainder.size() >= 4u && StartsWithAsciiNoCase(remainder.substr(0u, 3u), L"UNC") && IsSeparator(remainder[3]))
        {
            return WindowsPathClass::ExtendedUnc;
        }
        return WindowsPathClass::ExtendedOther;
    }

    if (path.size() >= 2u && IsSeparator(path[0]) && IsSeparator(path[1]))
    {
        return WindowsPathClass::Unc;
    }
    if (path.size() >= 2u && IsAsciiDriveLetter(path[0]) && path[1] == L':')
    {
        return path.size() >= 3u && IsSeparator(path[2]) ? WindowsPathClass::DriveAbsolute : WindowsPathClass::DriveRelative;
    }
    if (! path.empty() && IsSeparator(path[0]))
    {
        return WindowsPathClass::Rooted;
    }
    return WindowsPathClass::Relative;
}

[[nodiscard]] inline bool IsDriveQualifiedWindowsPath(std::wstring_view path) noexcept
{
    const WindowsPathClass pathClass = ClassifyWindowsPath(path);
    return pathClass == WindowsPathClass::DriveRelative || pathClass == WindowsPathClass::DriveAbsolute;
}

[[nodiscard]] inline bool IsDriveAbsoluteWindowsPath(std::wstring_view path) noexcept
{
    return ClassifyWindowsPath(path) == WindowsPathClass::DriveAbsolute;
}

[[nodiscard]] inline bool IsUncWindowsPath(std::wstring_view path) noexcept
{
    const WindowsPathClass pathClass = ClassifyWindowsPath(path);
    return pathClass == WindowsPathClass::Unc || pathClass == WindowsPathClass::ExtendedUnc;
}

[[nodiscard]] inline bool IsExtendedWindowsPath(std::wstring_view path) noexcept
{
    const WindowsPathClass pathClass = ClassifyWindowsPath(path);
    return pathClass == WindowsPathClass::ExtendedDriveAbsolute || pathClass == WindowsPathClass::ExtendedUnc || pathClass == WindowsPathClass::ExtendedOther;
}

[[nodiscard]] inline bool IsDeviceWindowsPath(std::wstring_view path) noexcept
{
    return ClassifyWindowsPath(path) == WindowsPathClass::Device;
}

[[nodiscard]] inline bool HasUncServerAndShare(std::wstring_view path, size_t componentStart) noexcept
{
    const size_t serverEnd = path.find_first_of(L"\\/", componentStart);
    if (serverEnd == std::wstring_view::npos || serverEnd == componentStart)
    {
        return false;
    }
    const size_t shareStart = serverEnd + 1u;
    const size_t shareEnd   = path.find_first_of(L"\\/", shareStart);
    return shareStart < path.size() && (shareEnd == std::wstring_view::npos ? shareStart < path.size() : shareEnd > shareStart);
}

// A fully absolute path has an explicit drive root or a complete UNC server/share root.
// Device and other extended namespaces remain separately classified and are not accepted here.
[[nodiscard]] inline bool IsFullyAbsoluteWindowsPath(std::wstring_view path) noexcept
{
    switch (ClassifyWindowsPath(path))
    {
        case WindowsPathClass::DriveAbsolute:
        case WindowsPathClass::ExtendedDriveAbsolute: return true;
        case WindowsPathClass::Unc: return HasUncServerAndShare(path, 2u);
        case WindowsPathClass::ExtendedUnc: return HasUncServerAndShare(path, 8u);
        case WindowsPathClass::Relative:
        case WindowsPathClass::Rooted:
        case WindowsPathClass::DriveRelative:
        case WindowsPathClass::ExtendedOther:
        case WindowsPathClass::Device: return false;
    }
    return false;
}

// Adds an extended Win32 prefix without normalizing, resolving, or validating the input.
// Only drive-absolute and UNC paths are transformed; all other path classes are returned unchanged.
[[nodiscard]] inline std::wstring ToExtendedWin32Path(std::wstring_view path)
{
    switch (ClassifyWindowsPath(path))
    {
        case WindowsPathClass::DriveAbsolute: return std::wstring(LR"(\\?\)") + std::wstring(path);
        case WindowsPathClass::Unc: return std::wstring(LR"(\\?\UNC\)") + std::wstring(path.substr(2u));
        case WindowsPathClass::Relative:
        case WindowsPathClass::Rooted:
        case WindowsPathClass::DriveRelative:
        case WindowsPathClass::ExtendedDriveAbsolute:
        case WindowsPathClass::ExtendedUnc:
        case WindowsPathClass::ExtendedOther:
        case WindowsPathClass::Device: return std::wstring(path);
    }
    return std::wstring(path);
}

// Compares already-normalized Windows paths using ordinal case-insensitive semantics.
// Callers remain responsible for choosing lexical versus physical normalization; this
// helper intentionally does not resolve reparse points or touch the filesystem.
[[nodiscard]] inline bool NormalizedWindowsPathEqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size() || left.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    if (left.empty())
    {
        return true;
    }
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

// Component-boundary containment for already-normalized local Windows path text.
// Root and candidate must use the same separator convention (std::filesystem::path::native
// after lexically_normal is the canonical caller). Equality counts as containment.
[[nodiscard]] inline bool IsSameOrDescendantNormalizedWindowsPath(std::wstring_view root, std::wstring_view candidate) noexcept
{
    if (root.empty() || candidate.size() < root.size() || root.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    if (CompareStringOrdinal(root.data(), static_cast<int>(root.size()), candidate.data(), static_cast<int>(root.size()), TRUE) != CSTR_EQUAL)
    {
        return false;
    }
    if (candidate.size() == root.size())
    {
        return true;
    }
    return IsSeparator(root.back()) || IsSeparator(candidate[root.size()]);
}

struct UniqueSiblingFileOptions final
{
    std::wstring_view prefix;
    std::wstring_view suffix;
    DWORD desiredAccess      = GENERIC_WRITE;
    DWORD shareMode          = 0u;
    DWORD flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_NOT_CONTENT_INDEXED;
    size_t maximumAttempts   = 32u;
};

template <typename UniqueFile>
[[nodiscard]] inline HRESULT CreateUniqueFileInDirectory(std::wstring_view directory,
                                                         const UniqueSiblingFileOptions& options,
                                                         std::wstring& pathOut,
                                                         UniqueFile& fileOut) noexcept
{
    pathOut.clear();
    fileOut.reset();
    if (directory.empty() || options.maximumAttempts == 0u || options.prefix.find_first_of(L"\\/") != std::wstring_view::npos ||
        options.suffix.find_first_of(L"\\/") != std::wstring_view::npos)
    {
        return E_INVALIDARG;
    }

    for (size_t attempt = 0u; attempt < options.maximumAttempts; ++attempt)
    {
        static_cast<void>(attempt);
        GUID guid{};
        const HRESULT guidHr = CoCreateGuid(&guid);
        if (FAILED(guidHr))
        {
            return guidHr;
        }

        std::array<wchar_t, 40u> guidText{};
        if (StringFromGUID2(guid, guidText.data(), static_cast<int>(guidText.size())) <= 0)
        {
            return E_FAIL;
        }

        std::wstring candidate(directory);
        if (! IsSeparator(candidate.back()))
        {
            candidate.push_back(L'\\');
        }
        candidate.append(options.prefix);
        candidate.append(guidText.data());
        candidate.append(options.suffix);

        UniqueFile file(CreateFileW(candidate.c_str(), options.desiredAccess, options.shareMode, nullptr, CREATE_NEW, options.flagsAndAttributes, nullptr));
        if (file)
        {
            pathOut = std::move(candidate);
            fileOut = std::move(file);
            return S_OK;
        }

        const DWORD error = GetLastError();
        if (error != ERROR_FILE_EXISTS && error != ERROR_ALREADY_EXISTS)
        {
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
        }
    }
    return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
}

// Creates and returns a sibling using CREATE_NEW. The path is published only after the
// caller owns the handle, so collision retries cannot overwrite an existing file.
template <typename UniqueFile>
[[nodiscard]] inline HRESULT CreateUniqueSiblingFile(std::wstring_view siblingOf,
                                                     const UniqueSiblingFileOptions& options,
                                                     std::wstring& pathOut,
                                                     UniqueFile& fileOut) noexcept
{
    pathOut.clear();
    fileOut.reset();
    const size_t separator = siblingOf.find_last_of(L"\\/");
    if (separator == std::wstring_view::npos)
    {
        return E_INVALIDARG;
    }
    return CreateUniqueFileInDirectory(siblingOf.substr(0u, separator + 1u), options, pathOut, fileOut);
}
} // namespace Common::Paths

#include "FileSystem.Internal.h"

namespace FileSystemInternal
{
[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept
{
    return (name == L"." || name == L"..");
}

std::wstring MakeAbsolutePath(const std::wstring& path)
{
    std::wstring input = path;
    if (input.empty())
    {
        input = L".";
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
    if (normalized.empty())
    {
        normalized = L".";
    }

    if (normalized.rfind(L"\\\\?\\", 0) != 0)
    {
        normalized = MakeAbsolutePath(normalized);
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0)
    {
        return normalized;
    }

    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return L"\\\\?\\UNC\\" + normalized.substr(2);
    }

    return L"\\\\?\\" + normalized;
}

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
    info.display  = path;
    info.extended = ToExtendedPath(path);
    return info;
}

PathInfo MakePathInfo(const wchar_t* path)
{
    PathInfo info{};
    if (path)
    {
        info.display  = path;
        info.extended = ToExtendedPath(info.display);
    }
    return info;
}
} // namespace FileSystemInternal

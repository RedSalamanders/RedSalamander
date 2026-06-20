#include "FileSystemPathIdentity.h"

#include <Windows.h>

#include <limits>
#include <memory>
#include <string>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace
{
[[nodiscard]] bool EqualsOrdinalIgnoreCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }
    if (lhs.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    if (lhs.empty())
    {
        return true;
    }

    return ::CompareStringOrdinal(lhs.data(),
                                  static_cast<int>(lhs.size()),
                                  rhs.data(),
                                  static_cast<int>(rhs.size()),
                                  TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool JsonStringEquals(yyjson_val* value, std::string_view expected) noexcept
{
    if (! value || ! yyjson_is_str(value))
    {
        return false;
    }

    const char* text = yyjson_get_str(value);
    const size_t len = yyjson_get_len(value);
    return text && std::string_view{text, len} == expected;
}

[[nodiscard]] std::optional<bool> JsonBoolValue(yyjson_val* value) noexcept
{
    if (! value || ! yyjson_is_bool(value))
    {
        return std::nullopt;
    }

    return yyjson_get_bool(value);
}

[[nodiscard]] bool JsonIntEquals(yyjson_val* value, const int64_t expected) noexcept
{
    return value && yyjson_is_int(value) && yyjson_get_int(value) == expected;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::optional<std::wstring> JsonUtf16String(yyjson_val* value) noexcept
{
    if (! value || ! yyjson_is_str(value))
    {
        return std::nullopt;
    }

    const char* text = yyjson_get_str(value);
    const size_t len = yyjson_get_len(value);
    if (! text || len == 0)
    {
        return std::nullopt;
    }

    std::wstring wide = Utf16FromUtf8(std::string_view{text, len});
    if (wide.empty())
    {
        return std::nullopt;
    }

    return wide;
}

[[nodiscard]] std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> ReadJsonDocument(const std::string_view jsonUtf8) noexcept
{
    std::string jsonCopy(jsonUtf8);
    return std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)>{
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr),
        &yyjson_doc_free};
}

[[nodiscard]] std::optional<FileSystemPathCaseOnlyRename> ParseCaseOnlyRename(yyjson_val* value) noexcept
{
    if (JsonStringEquals(value, "supported"))
    {
        return FileSystemPathCaseOnlyRename::Supported;
    }
    if (JsonStringEquals(value, "noOp"))
    {
        return FileSystemPathCaseOnlyRename::NoOp;
    }
    if (JsonStringEquals(value, "unsupported"))
    {
        return FileSystemPathCaseOnlyRename::Unsupported;
    }
    if (JsonStringEquals(value, "notApplicable"))
    {
        return FileSystemPathCaseOnlyRename::NotApplicable;
    }

    return std::nullopt;
}

[[nodiscard]] bool IsAcceptedSeparator(const FileSystemPathIdentity& identity, const wchar_t ch) noexcept
{
    return identity.acceptedSeparators.find(ch) != std::wstring::npos;
}

[[nodiscard]] bool IsConventionalSeparator(const wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] size_t FindNextAcceptedSeparator(const FileSystemPathIdentity& identity,
                                               const std::wstring_view text,
                                               const size_t start) noexcept
{
    for (size_t i = start; i < text.size(); ++i)
    {
        if (IsAcceptedSeparator(identity, text[i]))
        {
            return i;
        }
    }

    return std::wstring_view::npos;
}

[[nodiscard]] wchar_t FoldAsciiOrdinalIgnoreCase(const wchar_t ch) noexcept
{
    if (ch >= L'A' && ch <= L'Z')
    {
        return static_cast<wchar_t>(ch - L'A' + L'a');
    }

    return ch;
}

[[nodiscard]] bool TryAppendKeyChar(const FileSystemPathIdentity& identity, const wchar_t ch, std::wstring& key) noexcept
{
    switch (identity.componentComparison)
    {
        case FileSystemPathComponentComparison::OrdinalCaseSensitive:
            key.push_back(ch);
            return true;

        case FileSystemPathComponentComparison::OrdinalIgnoreCase:
            if (ch > 0x7F)
            {
                return false;
            }

            key.push_back(FoldAsciiOrdinalIgnoreCase(ch));
            return true;
    }

    return false;
}

[[nodiscard]] std::optional<FileSystemPathIdentity> ParseFileSystemPathIdentityFromRoot(yyjson_val* root, const bool requireStable) noexcept
{
    if (! root || ! yyjson_is_obj(root))
    {
        return std::nullopt;
    }

    yyjson_val* identity = yyjson_obj_get(root, "pathIdentity");
    if (! identity)
    {
        identity = root;
    }
    if (! yyjson_is_obj(identity))
    {
        return std::nullopt;
    }

    if (! JsonIntEquals(yyjson_obj_get(identity, "version"), 1))
    {
        return std::nullopt;
    }

    const std::optional<bool> pathTextStableIdentity = JsonBoolValue(yyjson_obj_get(identity, "pathTextStableIdentity"));
    if (! pathTextStableIdentity.has_value())
    {
        return std::nullopt;
    }
    if (requireStable && ! pathTextStableIdentity.value())
    {
        return std::nullopt;
    }

    yyjson_val* normalization = yyjson_obj_get(identity, "normalization");
    if (! JsonStringEquals(normalization, "none"))
    {
        return std::nullopt;
    }

    yyjson_val* comparison = yyjson_obj_get(identity, "componentComparison");
    FileSystemPathIdentity parsed{};
    parsed.pathTextStableIdentity = pathTextStableIdentity.value();
    if (JsonStringEquals(comparison, "ordinalIgnoreCase"))
    {
        parsed.componentComparison = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    }
    else if (JsonStringEquals(comparison, "ordinalCaseSensitive"))
    {
        parsed.componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive;
    }
    else
    {
        return std::nullopt;
    }

    const std::optional<std::wstring> preferredSeparator = JsonUtf16String(yyjson_obj_get(identity, "preferredSeparator"));
    if (! preferredSeparator.has_value() || preferredSeparator->size() != 1)
    {
        return std::nullopt;
    }
    parsed.preferredSeparator = preferredSeparator->front();

    yyjson_val* acceptedSeparators = yyjson_obj_get(identity, "acceptedSeparators");
    if (! acceptedSeparators || ! yyjson_is_arr(acceptedSeparators) || yyjson_arr_size(acceptedSeparators) == 0)
    {
        return std::nullopt;
    }

    parsed.acceptedSeparators.clear();
    const size_t acceptedCount = yyjson_arr_size(acceptedSeparators);
    for (size_t i = 0; i < acceptedCount; ++i)
    {
        const std::optional<std::wstring> separator = JsonUtf16String(yyjson_arr_get(acceptedSeparators, i));
        if (! separator.has_value() || separator->size() != 1)
        {
            return std::nullopt;
        }

        const wchar_t ch = separator->front();
        if (parsed.acceptedSeparators.find(ch) == std::wstring::npos)
        {
            parsed.acceptedSeparators.push_back(ch);
        }
    }

    if (! IsAcceptedSeparator(parsed, parsed.preferredSeparator))
    {
        return std::nullopt;
    }

    const std::optional<bool> casePreserving = JsonBoolValue(yyjson_obj_get(identity, "casePreserving"));
    if (! casePreserving.has_value())
    {
        return std::nullopt;
    }
    parsed.casePreserving = casePreserving.value();

    const std::optional<FileSystemPathCaseOnlyRename> caseOnlyRename = ParseCaseOnlyRename(yyjson_obj_get(identity, "caseOnlyRename"));
    if (! caseOnlyRename.has_value())
    {
        return std::nullopt;
    }
    parsed.caseOnlyRename = caseOnlyRename.value();

    return parsed;
}
} // namespace

FileSystemPathIdentity FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem() noexcept
{
    return FileSystemPathIdentity{
        .pathTextStableIdentity = true,
        .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
        .preferredSeparator     = L'\\',
        .acceptedSeparators     = L"\\/",
        .casePreserving         = true,
        .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported,
    };
}

std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentityContract(const std::string_view jsonUtf8,
                                                                             const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    return ParseFileSystemPathIdentityFromRoot(yyjson_doc_get_root(doc.get()), false);
}

std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentity(const std::string_view jsonUtf8,
                                                                     const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    return ParseFileSystemPathIdentityFromRoot(yyjson_doc_get_root(doc.get()), true);
}

std::optional<FileSystemPathIdentity> TryParseFileSystemRenamePathIdentity(const std::string_view jsonUtf8,
                                                                           const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root) || ! JsonIntEquals(yyjson_obj_get(root, "version"), 1))
    {
        return std::nullopt;
    }

    yyjson_val* operations = yyjson_obj_get(root, "operations");
    const std::optional<bool> rename = operations && yyjson_is_obj(operations) ? JsonBoolValue(yyjson_obj_get(operations, "rename")) : std::nullopt;
    if (! rename.value_or(false))
    {
        return std::nullopt;
    }

    return ParseFileSystemPathIdentityFromRoot(root, true);
}

bool EquivalentComponent(const FileSystemPathIdentity& identity, const std::wstring_view lhs, const std::wstring_view rhs) noexcept
{
    if (! identity.pathTextStableIdentity)
    {
        return false;
    }

    switch (identity.componentComparison)
    {
        case FileSystemPathComponentComparison::OrdinalCaseSensitive: return lhs == rhs;
        case FileSystemPathComponentComparison::OrdinalIgnoreCase: return EqualsOrdinalIgnoreCase(lhs, rhs);
    }

    return false;
}

bool EquivalentPath(const FileSystemPathIdentity& identity, const std::wstring_view lhs, const std::wstring_view rhs) noexcept
{
    if (! identity.pathTextStableIdentity)
    {
        return false;
    }

    size_t lhsOffset = 0;
    size_t rhsOffset = 0;
    for (;;)
    {
        const size_t lhsSeparator = FindNextAcceptedSeparator(identity, lhs, lhsOffset);
        const size_t rhsSeparator = FindNextAcceptedSeparator(identity, rhs, rhsOffset);

        const size_t lhsComponentEnd = lhsSeparator == std::wstring_view::npos ? lhs.size() : lhsSeparator;
        const size_t rhsComponentEnd = rhsSeparator == std::wstring_view::npos ? rhs.size() : rhsSeparator;
        if (! EquivalentComponent(identity,
                                  lhs.substr(lhsOffset, lhsComponentEnd - lhsOffset),
                                  rhs.substr(rhsOffset, rhsComponentEnd - rhsOffset)))
        {
            return false;
        }

        const bool lhsDone = lhsSeparator == std::wstring_view::npos;
        const bool rhsDone = rhsSeparator == std::wstring_view::npos;
        if (lhsDone || rhsDone)
        {
            return lhsDone && rhsDone;
        }

        lhsOffset = lhsSeparator + 1;
        rhsOffset = rhsSeparator + 1;
    }
}

std::optional<std::wstring> TryMakeComponentKey(const FileSystemPathIdentity& identity, const std::wstring_view component) noexcept
{
    if (! identity.pathTextStableIdentity)
    {
        return std::nullopt;
    }

    std::wstring key;
    key.reserve(component.size());
    for (const wchar_t ch : component)
    {
        if (! TryAppendKeyChar(identity, ch, key))
        {
            return std::nullopt;
        }
    }

    return key;
}

std::optional<std::wstring> TryMakePathKey(const FileSystemPathIdentity& identity, const std::wstring_view path) noexcept
{
    if (! identity.pathTextStableIdentity)
    {
        return std::nullopt;
    }

    std::wstring key;
    key.reserve(path.size());
    for (const wchar_t ch : path)
    {
        if (IsAcceptedSeparator(identity, ch))
        {
            key.push_back(identity.preferredSeparator);
            continue;
        }

        if (IsConventionalSeparator(ch))
        {
            return std::nullopt;
        }

        if (! TryAppendKeyChar(identity, ch, key))
        {
            return std::nullopt;
        }
    }

    return key;
}

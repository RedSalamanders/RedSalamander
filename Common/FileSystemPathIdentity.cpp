#include "FileSystemPathIdentity.h"
#include "YyjsonHelpers.h"

#include <Windows.h>

#include <limits>
#include <string>

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

    return ::CompareStringOrdinal(lhs.data(), static_cast<int>(lhs.size()), rhs.data(), static_cast<int>(rhs.size()), TRUE) == CSTR_EQUAL;
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

    const std::optional<std::wstring> wide = Common::Strings::TryUtf16FromUtf8Strict(std::string_view{text, len});
    if (! wide.has_value() || wide.value().empty())
    {
        return std::nullopt;
    }

    return wide.value();
}

[[nodiscard]] Common::Json::UniqueDocument ReadJsonDocument(const std::string_view jsonUtf8) noexcept
{
    std::string jsonCopy(jsonUtf8);
    return Common::Json::UniqueDocument{
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr)};
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

[[nodiscard]] size_t FindNextAcceptedSeparator(const FileSystemPathIdentity& identity, const std::wstring_view text, const size_t start) noexcept
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

[[nodiscard]] bool TryAppendComponentKey(const FileSystemPathIdentity& identity, const std::wstring_view component, std::wstring& key) noexcept
{
    switch (identity.componentComparison)
    {
        case FileSystemPathComponentComparison::OrdinalCaseSensitive: key.append(component); return true;

        case FileSystemPathComponentComparison::OrdinalIgnoreCase:
            if (component.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
            {
                return false;
            }
            if (component.empty())
            {
                return true;
            }

            const int sourceLength = static_cast<int>(component.size());
            const int required = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE, component.data(), sourceLength, nullptr, 0, nullptr, nullptr, 0);
            // CompareStringOrdinal uses simple ordinal folding. Refuse expanding mappings so a
            // locale-style multi-code-unit transform can never create a false hash collision.
            if (required != sourceLength)
            {
                return false;
            }
            const size_t originalSize = key.size();
            key.resize(originalSize + static_cast<size_t>(required));
            const int written = LCMapStringEx(LOCALE_NAME_INVARIANT,
                                              LCMAP_UPPERCASE,
                                              component.data(),
                                              sourceLength,
                                              key.data() + originalSize,
                                              required,
                                              nullptr,
                                              nullptr,
                                              0);
            if (written != required)
            {
                key.resize(originalSize);
                return false;
            }
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

    const Common::Json::MemberResult<int64_t> version =
        Common::Json::GetInt64Member(identity, "version", Common::Json::MemberRequirement::Required);
    if (! version.HasValue() || version.value != 1)
    {
        return std::nullopt;
    }

    const Common::Json::MemberResult<bool> pathTextStableIdentity =
        Common::Json::GetBoolMember(identity, "pathTextStableIdentity", Common::Json::MemberRequirement::Required);
    if (! pathTextStableIdentity.HasValue())
    {
        return std::nullopt;
    }
    if (requireStable && ! pathTextStableIdentity.value)
    {
        return std::nullopt;
    }

    const Common::Json::MemberResult<std::string_view> normalization =
        Common::Json::GetStringMember(identity, "normalization", Common::Json::MemberRequirement::Required);
    if (! normalization.HasValue() || normalization.value != "none")
    {
        return std::nullopt;
    }

    const Common::Json::MemberResult<std::string_view> comparison =
        Common::Json::GetStringMember(identity, "componentComparison", Common::Json::MemberRequirement::Required);
    FileSystemPathIdentity parsed{};
    parsed.pathTextStableIdentity = pathTextStableIdentity.value;
    if (comparison.HasValue() && comparison.value == "ordinalIgnoreCase")
    {
        parsed.componentComparison = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    }
    else if (comparison.HasValue() && comparison.value == "ordinalCaseSensitive")
    {
        parsed.componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive;
    }
    else
    {
        return std::nullopt;
    }

    const Common::Json::MemberResult<std::wstring> preferredSeparator =
        Common::Json::GetUtf16StringMemberStrict(identity, "preferredSeparator", Common::Json::MemberRequirement::Required);
    if (! preferredSeparator.HasValue() || preferredSeparator.value.size() != 1)
    {
        return std::nullopt;
    }
    parsed.preferredSeparator = preferredSeparator.value.front();

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

    const Common::Json::MemberResult<bool> casePreserving =
        Common::Json::GetBoolMember(identity, "casePreserving", Common::Json::MemberRequirement::Required);
    if (! casePreserving.HasValue())
    {
        return std::nullopt;
    }
    parsed.casePreserving = casePreserving.value;

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

std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentityContract(const std::string_view jsonUtf8, const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    Common::Json::UniqueDocument doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    return TryParseFileSystemPathIdentityContractFromRoot(yyjson_doc_get_root(doc.get()), pluginId);
}

std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentityContractFromRoot(yyjson_val* root, const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);
    return ParseFileSystemPathIdentityFromRoot(root, false);
}

std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentity(const std::string_view jsonUtf8, const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    Common::Json::UniqueDocument doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    return ParseFileSystemPathIdentityFromRoot(yyjson_doc_get_root(doc.get()), true);
}

std::optional<FileSystemPathIdentity> TryParseFileSystemRenamePathIdentity(const std::string_view jsonUtf8, const std::wstring_view pluginId) noexcept
{
    static_cast<void>(pluginId);

    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    Common::Json::UniqueDocument doc = ReadJsonDocument(jsonUtf8);
    if (! doc)
    {
        return std::nullopt;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    const Common::Json::MemberResult<int64_t> version =
        Common::Json::GetInt64Member(root, "version", Common::Json::MemberRequirement::Required);
    if (! root || ! yyjson_is_obj(root) || ! version.HasValue() || version.value != 1)
    {
        return std::nullopt;
    }

    yyjson_val* operations = yyjson_obj_get(root, "operations");
    const Common::Json::MemberResult<bool> rename =
        Common::Json::GetBoolMember(operations, "rename", Common::Json::MemberRequirement::Required);
    if (! rename.HasValue() || ! rename.value)
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
        if (! EquivalentComponent(identity, lhs.substr(lhsOffset, lhsComponentEnd - lhsOffset), rhs.substr(rhsOffset, rhsComponentEnd - rhsOffset)))
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

std::wstring JoinFileSystemPath(const FileSystemPathIdentity& identity, const std::wstring_view folder, const std::wstring_view leaf)
{
    if (folder.empty())
    {
        return std::wstring(leaf);
    }

    std::wstring result(folder);
    if (! result.empty() && ! IsAcceptedSeparator(identity, result.back()))
    {
        result.push_back(identity.preferredSeparator);
    }
    result.append(leaf);
    return result;
}

bool IsStrictDescendantPath(const FileSystemPathIdentity& identity,
                            const std::wstring_view prefix,
                            const std::wstring_view candidate) noexcept
{
    if (! identity.pathTextStableIdentity || prefix.empty())
    {
        return false;
    }

    size_t prefixOffset    = 0u;
    size_t candidateOffset = 0u;
    for (;;)
    {
        const size_t prefixSeparator    = FindNextAcceptedSeparator(identity, prefix, prefixOffset);
        const size_t candidateSeparator = FindNextAcceptedSeparator(identity, candidate, candidateOffset);
        const size_t prefixEnd          = prefixSeparator == std::wstring_view::npos ? prefix.size() : prefixSeparator;
        const size_t candidateEnd       = candidateSeparator == std::wstring_view::npos ? candidate.size() : candidateSeparator;
        if (! EquivalentComponent(identity,
                                  prefix.substr(prefixOffset, prefixEnd - prefixOffset),
                                  candidate.substr(candidateOffset, candidateEnd - candidateOffset)))
        {
            return false;
        }

        if (prefixSeparator == std::wstring_view::npos)
        {
            return candidateSeparator != std::wstring_view::npos;
        }
        if (candidateSeparator == std::wstring_view::npos)
        {
            return false;
        }
        prefixOffset    = prefixSeparator + 1u;
        candidateOffset = candidateSeparator + 1u;
    }
}

std::wstring ReplaceFileSystemPathPrefix(const FileSystemPathIdentity& identity,
                                         const std::wstring_view candidate,
                                         const std::wstring_view oldPrefix,
                                         const std::wstring_view newPrefix)
{
    if (! IsStrictDescendantPath(identity, oldPrefix, candidate))
    {
        return std::wstring(candidate);
    }

    size_t oldOffset       = 0u;
    size_t candidateOffset = 0u;
    for (;;)
    {
        const size_t oldSeparator       = FindNextAcceptedSeparator(identity, oldPrefix, oldOffset);
        const size_t candidateSeparator = FindNextAcceptedSeparator(identity, candidate, candidateOffset);
        if (oldSeparator == std::wstring_view::npos)
        {
            candidateOffset = candidateSeparator == std::wstring_view::npos ? candidate.size() : candidateSeparator + 1u;
            break;
        }
        oldOffset       = oldSeparator + 1u;
        candidateOffset = candidateSeparator + 1u;
    }

    return JoinFileSystemPath(identity, newPrefix, candidate.substr(candidateOffset));
}

std::optional<std::wstring> TryMakeComponentKey(const FileSystemPathIdentity& identity, const std::wstring_view component) noexcept
{
    if (! identity.pathTextStableIdentity)
    {
        return std::nullopt;
    }

    std::wstring key;
    key.reserve(component.size());
    if (! TryAppendComponentKey(identity, component, key))
    {
        return std::nullopt;
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
    size_t componentStart = 0u;
    for (size_t index = 0u; index <= path.size(); ++index)
    {
        if (index != path.size() && ! IsAcceptedSeparator(identity, path[index]))
        {
            if (IsConventionalSeparator(path[index]))
            {
                return std::nullopt;
            }
            continue;
        }

        if (! TryAppendComponentKey(identity, path.substr(componentStart, index - componentStart), key))
        {
            return std::nullopt;
        }
        if (index != path.size())
        {
            key.push_back(identity.preferredSeparator);
            componentStart = index + 1u;
        }
    }

    return key;
}

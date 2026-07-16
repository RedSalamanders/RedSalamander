#include "BatchRenameEngine.h"

#include "ChangeCase.h"
#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <optional>
#include <regex>
#include <string_view>
#include <unordered_map>

namespace BatchRename
{
namespace
{
constexpr size_t kWindowsMaxLeafNameLength = 255u;

// UI text features still use the legacy per-code-unit fold. Provider path identity decisions
// flow through FileSystemPathIdentity instead.
[[nodiscard]] wchar_t FoldCaseUnit(wchar_t ch) noexcept
{
    wchar_t unit = ch;
    ::CharUpperBuffW(&unit, 1u);
    return unit;
}

[[nodiscard]] std::wstring FoldCaseForCollisionKeys(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (! out.empty())
    {
        ::CharUpperBuffW(out.data(), static_cast<DWORD>(out.size()));
    }
    return out;
}

[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept
{
    return text.find(L'\\') != std::wstring_view::npos || text.find(L'/') != std::wstring_view::npos;
}

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view text) noexcept
{
    return text == L"." || text == L"..";
}

[[nodiscard]] bool HasWindowsInvalidLeafCharacter(std::wstring_view text) noexcept
{
    constexpr std::wstring_view kInvalidLeafCharacters = L"<>:\"|?*";
    return std::any_of(
        text.begin(), text.end(), [](const wchar_t ch) noexcept { return ch < L' ' || kInvalidLeafCharacters.find(ch) != std::wstring_view::npos; });
}

[[nodiscard]] bool IsWordChar(wchar_t ch) noexcept
{
    // Classifies one UTF-16 code unit; supplementary-plane characters are evaluated per
    // surrogate unit, which Windows reports as non-alphanumeric.
    return ::IsCharAlphaNumericW(ch) != FALSE || ch == L'_';
}

[[nodiscard]] bool IsHighSurrogate(const wchar_t ch) noexcept
{
    return ch >= 0xD800 && ch <= 0xDBFF;
}

[[nodiscard]] bool IsLowSurrogate(const wchar_t ch) noexcept
{
    return ch >= 0xDC00 && ch <= 0xDFFF;
}

[[nodiscard]] bool SplitsSurrogatePair(std::wstring_view text, const size_t boundary) noexcept
{
    return boundary > 0u && boundary < text.size() && IsHighSurrogate(text[boundary - 1u]) && IsLowSurrogate(text[boundary]);
}

[[nodiscard]] bool IsWholeWordMatch(std::wstring_view text, size_t position, size_t length) noexcept
{
    const size_t after = position + length;
    if (SplitsSurrogatePair(text, position) || SplitsSurrogatePair(text, after))
    {
        return false;
    }

    const bool leftBoundary  = position == 0 || ! IsWordChar(text[position - 1u]);
    const bool rightBoundary = after >= text.size() || ! IsWordChar(text[after]);
    return leftBoundary && rightBoundary;
}

[[nodiscard]] bool EqualsAt(std::wstring_view text, size_t position, std::wstring_view needle, bool caseSensitive) noexcept
{
    if (needle.empty() || position > text.size() || needle.size() > text.size() - position)
    {
        return false;
    }

    if (caseSensitive)
    {
        return text.substr(position, needle.size()) == needle;
    }

    for (size_t i = 0; i < needle.size(); ++i)
    {
        if (FoldCaseUnit(text[position + i]) != FoldCaseUnit(needle[i]))
        {
            return false;
        }
    }
    return true;
}

struct LeafNameParts final
{
    std::wstring_view stem;
    std::wstring_view extension;
};

[[nodiscard]] LeafNameParts SplitLeafNameParts(std::wstring_view leafName) noexcept
{
    if (leafName == L"." || leafName == L"..")
    {
        return LeafNameParts{.stem = leafName};
    }

    const size_t dot = leafName.rfind(L'.');
    if (dot == std::wstring_view::npos)
    {
        return LeafNameParts{.stem = leafName};
    }
    if (dot == 0u && leafName.find(L'.', 1u) == std::wstring_view::npos)
    {
        return LeafNameParts{.stem = leafName};
    }

    return LeafNameParts{
        .stem      = leafName.substr(0u, dot),
        .extension = leafName.substr(dot),
    };
}

[[nodiscard]] std::wstring StemOf(std::wstring_view leafName)
{
    return std::wstring(SplitLeafNameParts(leafName).stem);
}

[[nodiscard]] std::wstring ExtensionOf(std::wstring_view leafName)
{
    return std::wstring(SplitLeafNameParts(leafName).extension);
}

[[nodiscard]] std::wstring ReplaceExtensionPart(std::wstring_view leafName, std::wstring_view newStem)
{
    std::wstring result(newStem);
    result.append(SplitLeafNameParts(leafName).extension);
    return result;
}

[[nodiscard]] size_t CounterPadWidth(std::wstring_view format) noexcept
{
    if (format.empty())
    {
        return 0u;
    }
    if (! std::all_of(format.begin(), format.end(), [](const wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; }))
    {
        return 0u;
    }

    // "{counter:000}" pads to the format length; "{counter:3}" pads to the numeric width.
    if (std::all_of(format.begin(), format.end(), [](wchar_t ch) noexcept { return ch == L'0'; }))
    {
        return format.size();
    }

    size_t width = 0u;
    for (const wchar_t ch : format)
    {
        width = width * 10u + static_cast<size_t>(ch - L'0');
        if (width > kWindowsMaxLeafNameLength)
        {
            return kWindowsMaxLeafNameLength;
        }
    }
    return width;
}

[[nodiscard]] bool IsCounterPadFormatValid(std::wstring_view format) noexcept
{
    return format.empty() || std::all_of(format.begin(), format.end(), [](const wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; });
}

[[nodiscard]] std::wstring PadCounter(size_t value, std::wstring_view format)
{
    std::wstring digits = std::to_wstring(value);
    const size_t width  = CounterPadWidth(format);
    if (width > digits.size())
    {
        digits.insert(digits.begin(), width - digits.size(), L'0');
    }
    return digits;
}

void AppendPaddedNumber(std::wstring& out, const int value, const size_t width)
{
    std::wstring digits = std::to_wstring(value);
    if (digits.size() < width)
    {
        out.append(width - digits.size(), L'0');
    }
    out.append(digits);
}

[[nodiscard]] bool StartsWithAt(std::wstring_view text, const size_t index, std::wstring_view token) noexcept
{
    return index <= text.size() && token.size() <= text.size() - index && text.substr(index, token.size()) == token;
}

[[nodiscard]] std::wstring NormalizeAliasTokens(std::wstring_view text)
{
    std::wstring normalized;
    normalized.reserve(text.size());

    for (size_t i = 0; i < text.size();)
    {
        if (i + 1u < text.size() && text[i] == L'$' && text[i + 1u] == L'(')
        {
            const size_t close = text.find(L')', i + 2u);
            if (close != std::wstring_view::npos)
            {
                normalized.push_back(L'{');
                normalized.append(text.substr(i + 2u, close - (i + 2u)));
                normalized.push_back(L'}');
                i = close + 1u;
                continue;
            }
        }

        normalized.push_back(text[i]);
        ++i;
    }

    return normalized;
}

[[nodiscard]] std::wstring FlattenPathPart(std::wstring_view text, std::wstring_view separator)
{
    std::wstring result;
    result.reserve(text.size());

    bool lastWasSeparator = false;
    for (const wchar_t ch : text)
    {
        const bool isSeparator = ch == L'\\' || ch == L'/';
        if (isSeparator)
        {
            if (! lastWasSeparator)
            {
                result.append(separator);
            }
            lastWasSeparator = true;
            continue;
        }

        result.push_back(ch);
        lastWasSeparator = false;
    }

    return result;
}

struct MacroExpansion final
{
    std::wstring text;
    std::vector<Issue> issues;
};

[[nodiscard]] std::wstring ResolveMacro(
    std::wstring_view token, const Target& target, std::wstring_view originalName, size_t rowIndex, const Rules& rules, bool& known, bool& invalidFormat)
{
    known         = true;
    invalidFormat = false;

    const size_t colon             = token.find(L':');
    const std::wstring nameLower   = FoldCaseForCollisionKeys(colon == std::wstring_view::npos ? token : token.substr(0, colon));
    const std::wstring_view format = colon == std::wstring_view::npos ? std::wstring_view{} : token.substr(colon + 1u);

    if (nameLower == L"NAME" || nameLower == L"FILENAME")
    {
        return std::wstring(originalName);
    }
    if (nameLower == L"STEM")
    {
        return StemOf(originalName);
    }
    if (nameLower == L"EXT")
    {
        return ExtensionOf(originalName);
    }
    if (nameLower == L"EXTNODOT")
    {
        std::wstring ext = ExtensionOf(originalName);
        if (! ext.empty() && ext.front() == L'.')
        {
            ext.erase(ext.begin());
        }
        return ext;
    }
    if (nameLower == L"PARENT")
    {
        return target.sourcePath.parent_path().filename().wstring();
    }
    if (nameLower == L"RELATIVEFOLDER")
    {
        return target.relativeFolder.wstring();
    }
    if (nameLower == L"RELATIVEFOLDERFLAT")
    {
        return FlattenPathPart(target.relativeFolder.wstring(), rules.flattenSeparator);
    }
    if (nameLower == L"SIZE")
    {
        return std::to_wstring(target.sizeBytes);
    }
    if (nameLower == L"DATE")
    {
        return target.lastWriteTime.has_value() ? FormatTimestamp(target.lastWriteTime.value(), format) : std::wstring{};
    }
    if (nameLower == L"TIME")
    {
        return target.lastWriteTime.has_value() ? FormatTimestamp(target.lastWriteTime.value(), format.empty() ? std::wstring_view{L"HH-mm-ss"} : format)
                                                : std::wstring{};
    }
    if (nameLower == L"CREATED")
    {
        return target.createdTime.has_value() ? FormatTimestamp(target.createdTime.value(), format) : std::wstring{};
    }
    if (nameLower == L"COUNTER")
    {
        if (! IsCounterPadFormatValid(format))
        {
            invalidFormat = true;
            return {};
        }
        return PadCounter(rowIndex + 1u, format);
    }
    if (nameLower == L"INDEX")
    {
        if (! IsCounterPadFormatValid(format))
        {
            invalidFormat = true;
            return {};
        }
        return PadCounter(rowIndex, format);
    }

    known = false;
    return {};
}

[[nodiscard]] MacroExpansion ExpandTemplate(const Target& target, size_t rowIndex, const Rules& rules, std::wstring_view originalName)
{
    const std::wstring normalized = NormalizeAliasTokens(rules.nameTemplate.empty() ? std::wstring_view{L"{name}"} : std::wstring_view{rules.nameTemplate});
    MacroExpansion expansion{};
    expansion.text.reserve(normalized.size() + originalName.size());

    for (size_t i = 0; i < normalized.size();)
    {
        const wchar_t ch = normalized[i];
        if (ch == L'{' && i + 1u < normalized.size() && normalized[i + 1u] == L'{')
        {
            expansion.text.push_back(L'{');
            i += 2u;
            continue;
        }
        if (ch == L'}' && i + 1u < normalized.size() && normalized[i + 1u] == L'}')
        {
            expansion.text.push_back(L'}');
            i += 2u;
            continue;
        }
        if (ch == L'{')
        {
            const size_t close = normalized.find(L'}', i + 1u);
            if (close == std::wstring::npos)
            {
                expansion.issues.push_back({IssueSeverity::Error, L"macro_unclosed"});
                expansion.text.append(normalized.substr(i));
                break;
            }

            bool known         = false;
            bool invalidFormat = false;
            const std::wstring replacement =
                ResolveMacro(std::wstring_view(normalized).substr(i + 1u, close - (i + 1u)), target, originalName, rowIndex, rules, known, invalidFormat);
            if (! known)
            {
                expansion.issues.push_back({IssueSeverity::Error, L"macro_unknown"});
            }
            if (invalidFormat)
            {
                expansion.issues.push_back({IssueSeverity::Error, L"macro_invalid_format"});
            }
            expansion.text.append(replacement);
            i = close + 1u;
            continue;
        }

        expansion.text.push_back(ch);
        ++i;
    }

    return expansion;
}

[[nodiscard]] std::wstring ReplaceLiteral(std::wstring_view text, const Rules& rules)
{
    if (rules.searchFor.empty())
    {
        return std::wstring(text);
    }

    std::wstring result;
    result.reserve(text.size());

    size_t cursor  = 0;
    size_t matches = 0;
    while (cursor < text.size())
    {
        bool matched = EqualsAt(text, cursor, rules.searchFor, rules.caseSensitive);
        if (matched && rules.wholeWords)
        {
            matched = IsWholeWordMatch(text, cursor, rules.searchFor.size());
        }

        if (! matched || (rules.replaceOnce && matches != 0u))
        {
            result.push_back(text[cursor]);
            ++cursor;
            continue;
        }

        result.append(rules.replaceWith);
        cursor += rules.searchFor.size();
        ++matches;
    }

    if (cursor == text.size())
    {
        return result;
    }

    result.append(text.substr(cursor));
    return result;
}

[[nodiscard]] std::wstring ApplyReplacement(std::wstring_view leafName,
                                            const Rules& rules,
                                            const std::optional<std::wregex>& compiledRegex,
                                            std::optional<std::wstring>& matchError)
{
    if (rules.searchFor.empty())
    {
        return std::wstring(leafName);
    }

    const std::wstring stem = rules.excludeExtension ? StemOf(leafName) : std::wstring(leafName);
    std::wstring replaced;

    if (rules.regexEnabled && compiledRegex.has_value())
    {
        const std::regex_constants::match_flag_type flags = rules.replaceOnce ? std::regex_constants::format_first_only : std::regex_constants::match_default;
        try
        {
            replaced = std::regex_replace(stem, compiledRegex.value(), rules.replaceWith, flags);
        }
        catch (const std::regex_error&)
        {
            // MSVC can throw error_complexity/error_stack at match time for backtracking-heavy
            // patterns even when compilation succeeded; surface a per-row blocking issue instead
            // of letting the exception escape the noexcept BuildPlan.
            matchError = L"regex_match_failed";
            replaced   = stem;
        }
    }
    else if (! rules.regexEnabled)
    {
        replaced = ReplaceLiteral(stem, rules);
    }
    else
    {
        replaced = stem;
    }

    if (! rules.excludeExtension)
    {
        return replaced;
    }

    return ReplaceExtensionPart(leafName, replaced);
}

[[nodiscard]] std::optional<ChangeCase::CaseStyle> ToChangeCaseStyle(CaseTransform transform) noexcept
{
    switch (transform)
    {
        case CaseTransform::None: return std::nullopt;
        case CaseTransform::Lower: return ChangeCase::CaseStyle::Lower;
        case CaseTransform::Upper: return ChangeCase::CaseStyle::Upper;
        case CaseTransform::Mixed: return ChangeCase::CaseStyle::Mixed;
    }
    return std::nullopt;
}

[[nodiscard]] std::wstring ApplyCaseTransforms(std::wstring_view leafName, const Rules& rules) noexcept
{
    std::wstring result(leafName);

    if (const std::optional<ChangeCase::CaseStyle> style = ToChangeCaseStyle(rules.fileNameCaseStyle); style.has_value())
    {
        ChangeCase::Options options{};
        options.style  = style.value();
        options.target = ChangeCase::ChangeTarget::OnlyName;
        result         = ChangeCase::TransformLeafName(result, options);
    }

    if (const std::optional<ChangeCase::CaseStyle> style = ToChangeCaseStyle(rules.extensionCaseStyle); style.has_value())
    {
        ChangeCase::Options options{};
        options.style  = style.value();
        options.target = ChangeCase::ChangeTarget::OnlyExtension;
        result         = ChangeCase::TransformLeafName(result, options);
    }

    return result;
}

[[nodiscard]] bool IsReservedDeviceName(std::wstring_view leafName) noexcept
{
    // Win32 resolves DOS device names (CON, PRN, AUX, NUL, COM0-9, LPT0-9, CONIN$, CONOUT$) from
    // the first dot-delimited token of the leaf name, ignoring trailing spaces and dots, so renames
    // to e.g. "CON.txt" succeed under the \\?\ prefix but create un-openable files.
    std::wstring_view token = leafName;
    while (! token.empty() && (token.back() == L' ' || token.back() == L'.'))
    {
        token.remove_suffix(1u);
    }

    const size_t firstDot = token.find(L'.');
    if (firstDot != std::wstring_view::npos)
    {
        token = token.substr(0, firstDot);
    }
    while (! token.empty() && token.back() == L' ')
    {
        token.remove_suffix(1u);
    }

    if (token.size() != 3u && token.size() != 4u && token.size() != 6u && token.size() != 7u)
    {
        return false;
    }

    const std::wstring folded = FoldCaseForCollisionKeys(token);
    if (folded == L"CON" || folded == L"PRN" || folded == L"AUX" || folded == L"NUL")
    {
        return true;
    }
    if (folded == L"CONIN$" || folded == L"CONOUT$")
    {
        return true;
    }
    return folded.size() == 4u && (folded.starts_with(L"COM") || folded.starts_with(L"LPT")) && folded[3] >= L'0' && folded[3] <= L'9';
}

void ValidateLeafName(PreviewRow& row)
{
    if (row.newName.empty())
    {
        AddIssue(row, IssueSeverity::Error, L"name_empty");
    }
    if (IsDotOrDotDot(row.newName))
    {
        AddIssue(row, IssueSeverity::Error, L"name_dot");
    }
    if (ContainsPathSeparator(row.newName))
    {
        AddIssue(row, IssueSeverity::Error, L"name_separator");
    }
    if (HasWindowsInvalidLeafCharacter(row.newName))
    {
        AddIssue(row, IssueSeverity::Error, L"name_invalid_character");
    }
    if (IsReservedDeviceName(row.newName))
    {
        AddIssue(row, IssueSeverity::Error, L"name_reserved_device");
    }
    if (row.newName.size() > kWindowsMaxLeafNameLength)
    {
        AddIssue(row, IssueSeverity::Error, L"name_too_long");
    }
}

[[nodiscard]] bool HasEdgeSpaceOrTrailingDot(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return false;
    }

    return name.front() == L' ' || name.back() == L' ' || name.back() == L'.';
}

void AddWarningIssues(PreviewRow& row, const FileSystemPathIdentity& pathIdentity)
{
    if (row.newName.empty())
    {
        return;
    }

    if (row.newName == row.originalName)
    {
        AddIssue(row, IssueSeverity::Warning, L"name_unchanged");
    }
    else if (EquivalentComponent(pathIdentity, row.newName, row.originalName))
    {
        AddIssue(row, IssueSeverity::Warning, L"name_case_only");
    }

    if (HasEdgeSpaceOrTrailingDot(row.newName))
    {
        AddIssue(row, IssueSeverity::Warning, L"name_edge_space_or_dot");
    }
}

void MarkDuplicateTargets(std::vector<PreviewRow>& rows, const FileSystemPathIdentity& pathIdentity)
{
    std::vector<bool> duplicateRows(rows.size(), false);

    std::unordered_map<std::wstring, std::vector<size_t>> keyedRows;
    keyedRows.reserve(rows.size());
    bool canUseKeys = true;

    for (size_t i = 0; i < rows.size(); ++i)
    {
        if (rows[i].newName.empty())
        {
            continue;
        }

        const std::wstring parent                   = rows[i].sourcePath.parent_path().wstring();
        const std::optional<std::wstring> parentKey = TryMakePathKey(pathIdentity, parent);
        const std::optional<std::wstring> nameKey   = TryMakeComponentKey(pathIdentity, rows[i].newName);
        if (! parentKey.has_value() || ! nameKey.has_value())
        {
            canUseKeys = false;
            break;
        }

        std::wstring key;
        key.reserve(parentKey->size() + 1u + nameKey->size());
        key.append(parentKey.value());
        key.push_back(L'\0');
        key.append(nameKey.value());

        std::vector<size_t>& matches = keyedRows[key];
        for (const size_t candidate : matches)
        {
            if (EquivalentPath(pathIdentity, parent, rows[candidate].sourcePath.parent_path().wstring()) &&
                EquivalentComponent(pathIdentity, rows[i].newName, rows[candidate].newName))
            {
                duplicateRows[i]         = true;
                duplicateRows[candidate] = true;
            }
        }
        matches.push_back(i);
    }

    if (! canUseKeys)
    {
        std::ranges::fill(duplicateRows, false);

        for (size_t i = 0; i < rows.size(); ++i)
        {
            if (rows[i].newName.empty())
            {
                continue;
            }

            const std::wstring parent = rows[i].sourcePath.parent_path().wstring();
            for (size_t j = i + 1u; j < rows.size(); ++j)
            {
                if (rows[j].newName.empty())
                {
                    continue;
                }

                if (! EquivalentPath(pathIdentity, parent, rows[j].sourcePath.parent_path().wstring()))
                {
                    continue;
                }
                if (! EquivalentComponent(pathIdentity, rows[i].newName, rows[j].newName))
                {
                    continue;
                }

                duplicateRows[i] = true;
                duplicateRows[j] = true;
            }
        }
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::EmitValue(L"batchrename.preview.duplicate_fallback_rows", canUseKeys ? 0u : static_cast<uint64_t>(rows.size()));
    }

    for (size_t i = 0; i < duplicateRows.size(); ++i)
    {
        if (duplicateRows[i])
        {
            AddIssue(rows[i], IssueSeverity::Error, L"name_duplicate");
        }
    }
}

} // namespace

void AddIssue(PreviewRow& row, const IssueSeverity severity, std::wstring message)
{
    Issue issue{};
    issue.severity = severity;
    issue.message  = std::move(message);
    row.issues.push_back(std::move(issue));
}

[[nodiscard]] bool HasIssueSeverity(const PreviewRow& row, const IssueSeverity severity) noexcept
{
    return std::ranges::any_of(row.issues, [severity](const Issue& issue) noexcept { return issue.severity == severity; });
}

[[nodiscard]] Stats RecomputeStats(const std::span<const PreviewRow> rows) noexcept
{
    Stats stats{};
    stats.totalRows = rows.size();

    for (const PreviewRow& row : rows)
    {
        if (row.newName == row.originalName)
        {
            ++stats.unchangedRows;
        }
        else
        {
            ++stats.changedRows;
        }

        if (HasIssueSeverity(row, IssueSeverity::Error))
        {
            ++stats.errorRows;
        }
        else if (HasIssueSeverity(row, IssueSeverity::Warning))
        {
            ++stats.warningRows;
        }
    }

    return stats;
}

void RecomputeStats(Plan& plan) noexcept
{
    plan.stats = RecomputeStats(std::span<const PreviewRow>{plan.rows.data(), plan.rows.size()});
}

[[nodiscard]] std::chrono::local_seconds ToLocalWallClock(const std::chrono::sys_seconds timestamp) noexcept
{
    try
    {
        return std::chrono::current_zone()->to_local(timestamp);
    }
    catch (const std::runtime_error&)
    {
        // Catching is mandatory at this noexcept boundary: time-zone database lookup throws
        // std::runtime_error when the tzdb is missing or corrupt; fall back to UTC wall-clock
        // fields. std::bad_alloc intentionally propagates and terminates.
        return std::chrono::local_seconds{timestamp.time_since_epoch()};
    }
}

[[nodiscard]] std::wstring FormatTimestamp(const std::chrono::sys_seconds timestamp, std::wstring_view format)
{
    namespace chrono = std::chrono;

    if (format.empty())
    {
        format = L"yyyy-MM-dd";
    }

    const chrono::local_seconds localTime = ToLocalWallClock(timestamp);
    const chrono::local_days date         = chrono::floor<chrono::days>(localTime);
    const chrono::year_month_day ymd{date};
    const chrono::hh_mm_ss tod{localTime - date};

    const int year   = static_cast<int>(ymd.year());
    const int month  = static_cast<int>(static_cast<unsigned>(ymd.month()));
    const int day    = static_cast<int>(static_cast<unsigned>(ymd.day()));
    const int hour   = static_cast<int>(tod.hours().count());
    const int minute = static_cast<int>(tod.minutes().count());
    const int second = static_cast<int>(tod.seconds().count());

    std::wstring result;
    result.reserve(format.size());
    for (size_t i = 0; i < format.size();)
    {
        if (StartsWithAt(format, i, L"yyyy"))
        {
            AppendPaddedNumber(result, year, 4u);
            i += 4u;
        }
        else if (StartsWithAt(format, i, L"MM"))
        {
            AppendPaddedNumber(result, month, 2u);
            i += 2u;
        }
        else if (StartsWithAt(format, i, L"dd"))
        {
            AppendPaddedNumber(result, day, 2u);
            i += 2u;
        }
        else if (StartsWithAt(format, i, L"HH"))
        {
            AppendPaddedNumber(result, hour, 2u);
            i += 2u;
        }
        else if (StartsWithAt(format, i, L"mm"))
        {
            AppendPaddedNumber(result, minute, 2u);
            i += 2u;
        }
        else if (StartsWithAt(format, i, L"ss"))
        {
            AppendPaddedNumber(result, second, 2u);
            i += 2u;
        }
        else
        {
            result.push_back(format[i]);
            ++i;
        }
    }
    return result;
}

[[nodiscard]] std::wstring FormatDateText(const std::chrono::sys_seconds timestamp)
{
    return FormatTimestamp(timestamp, L"yyyy-MM-dd");
}

[[nodiscard]] std::wstring FormatTimeText(const std::chrono::sys_seconds timestamp)
{
    return FormatTimestamp(timestamp, L"HH:mm:ss");
}

Plan BuildPlan(const std::vector<Target>& targets, const Rules& rules) noexcept
{
    return BuildPlan(targets, rules, FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem());
}

Plan BuildPlan(const std::vector<Target>& targets, const Rules& rules, const FileSystemPathIdentity& pathIdentity) noexcept
{
    Debug::Perf::Scope buildPlanPerf(L"batchrename.preview.build_plan_us");

    Plan plan{};
    plan.rows.reserve(targets.size());

    std::optional<std::wregex> compiledRegex;
    std::optional<std::wstring> regexError;
    if (rules.regexEnabled && ! rules.searchFor.empty())
    {
        Debug::Perf::Scope regexCompilePerf(L"batchrename.regex.compile.us");
        regexCompilePerf.SetValue0(static_cast<uint64_t>(targets.size()));
        regexCompilePerf.SetValue1(static_cast<uint64_t>(rules.searchFor.size()));

        try
        {
            std::wstring pattern = rules.searchFor;
            if (rules.wholeWords)
            {
                // Non-capturing wrap so user capture-group indexes ($1, $2, ...) are preserved.
                pattern = L"\\b(?:" + pattern + L")\\b";
            }
            regexCompilePerf.SetValue1(static_cast<uint64_t>(pattern.size()));

            std::regex_constants::syntax_option_type flags = std::regex_constants::ECMAScript;
            if (! rules.caseSensitive)
            {
                flags |= std::regex_constants::icase;
            }
            compiledRegex.emplace(pattern, flags);
            regexCompilePerf.SetDetail(rules.wholeWords ? L"whole_words" : L"regex");
        }
        catch (const std::regex_error&)
        {
            regexCompilePerf.SetDetail(L"invalid");
            regexError = L"regex_invalid";
        }
    }

    const bool manualCountMismatch = rules.mode == Mode::Manual && rules.manualNames.size() != targets.size();

    for (size_t i = 0; i < targets.size(); ++i)
    {
        const Target& target = targets[i];
        PreviewRow row{};
        row.rowId           = static_cast<uint64_t>(i + 1u);
        row.sourcePath      = target.sourcePath;
        row.originalName    = target.sourcePath.filename().wstring();
        row.isDirectory     = target.isDirectory;
        row.metadataUnknown = target.metadataUnknown;
        row.sizeBytes       = target.sizeBytes;
        row.lastWriteTime   = target.lastWriteTime;
        row.createdTime     = target.createdTime;

        if (rules.mode == Mode::Manual)
        {
            if (i < rules.manualNames.size())
            {
                row.newName = rules.manualNames[i];
            }
            if (manualCountMismatch)
            {
                AddIssue(row, IssueSeverity::Error, L"manual_line_count");
            }
        }
        else
        {
            MacroExpansion expansion = ExpandTemplate(target, i, rules, row.originalName);
            row.newName              = std::move(expansion.text);
            for (Issue& issue : expansion.issues)
            {
                row.issues.push_back(std::move(issue));
            }

            std::optional<std::wstring> replacementError;
            row.newName = ApplyReplacement(row.newName, rules, compiledRegex, replacementError);
            row.newName = ApplyCaseTransforms(row.newName, rules);

            if (replacementError.has_value())
            {
                AddIssue(row, IssueSeverity::Error, replacementError.value());
            }
            if (regexError.has_value())
            {
                AddIssue(row, IssueSeverity::Error, regexError.value());
            }
        }

        plan.rows.push_back(std::move(row));
    }

    {
        Debug::Perf::Scope validationPerf(L"batchrename.validation.us");
        validationPerf.SetDetail(rules.mode == Mode::Manual ? L"manual" : L"rules");
        validationPerf.SetValue0(static_cast<uint64_t>(plan.rows.size()));

        for (PreviewRow& row : plan.rows)
        {
            ValidateLeafName(row);
            AddWarningIssues(row, pathIdentity);
        }

        MarkDuplicateTargets(plan.rows, pathIdentity);
        RecomputeStats(plan);
        validationPerf.SetValue1(static_cast<uint64_t>(plan.stats.errorRows));
    }

    buildPlanPerf.SetDetail(rules.mode == Mode::Manual ? L"manual" : L"rules");
    buildPlanPerf.SetValue0(static_cast<uint64_t>(plan.stats.totalRows));
    buildPlanPerf.SetValue1(static_cast<uint64_t>(plan.stats.changedRows));

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::EmitValue(L"batchrename.preview.rows", static_cast<uint64_t>(plan.stats.totalRows));
        Debug::Perf::EmitValue(L"batchrename.preview.changed", static_cast<uint64_t>(plan.stats.changedRows));
        Debug::Perf::EmitValue(L"batchrename.preview.errors", static_cast<uint64_t>(plan.stats.errorRows));
        Debug::Perf::EmitValue(L"batchrename.preview.warnings", static_cast<uint64_t>(plan.stats.warningRows));
    }

    return plan;
}
} // namespace BatchRename

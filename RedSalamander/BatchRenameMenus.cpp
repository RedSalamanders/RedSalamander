#include "Framework.h"

#include "BatchRenameMenus.h"

#include <algorithm>
#include <array>
#include <optional>

#include "Helpers.h"
#include "resource.h"

namespace
{
using RedSalamander::DxUi::MenuFlyoutItem;
using RedSalamander::DxUi::MenuItemKind;

enum HelperCommand : int
{
    TemplateName = 184000,
    TemplateStem,
    TemplateExt,
    TemplateExtNoDot,
    TemplateParent,
    TemplateRelativeFolderFlat,
    TemplateCounter,
    TemplateCounterPadded,
    TemplateDate,
    TemplateTime,
    TemplateLiteralOpenBrace,
    TemplateLiteralCloseBrace,
    RegexAnyCharacter,
    RegexCharacterSet,
    RegexNegatedSet,
    RegexWordBoundary,
    RegexAlternation,
    RegexZeroOrMore,
    RegexOneOrMore,
    RegexOptional,
    RegexGroup,
    RegexWordCharacter,
    RegexNonWordCharacter,
    RegexWhitespace,
    RegexNonWhitespace,
    RegexDigit,
    RegexNonDigit,
    RegexDecimalNumber,
    RegexHexNumber,
    RegexFileNameSplit,
    ReplaceLiteralDollar,
    ReplaceWholeMatch,
    ReplaceSubexpression1,
    ReplaceSubexpression2,
    RegexEscapedLiteral,
    ReplaceCustomSubexpression,
};

struct HelperMenuSpec final
{
    int commandId         = 0;
    unsigned int stringId = 0u;
    std::wstring_view fallbackText;
    std::wstring_view insertionText;
    bool separatorBefore = false;
};

[[nodiscard]] std::wstring LoadHelperString(const unsigned int id, std::wstring_view fallback)
{
    std::wstring value = LoadStringResource(nullptr, id);
    return value.empty() ? std::wstring(fallback) : std::move(value);
}

[[nodiscard]] MenuFlyoutItem BuildMenuItem(const HelperMenuSpec& spec)
{
    return MenuFlyoutItem{
        .text      = LoadHelperString(spec.stringId, spec.fallbackText),
        .commandId = spec.commandId,
    };
}

[[nodiscard]] std::vector<MenuFlyoutItem> BuildMenuItems(std::span<const HelperMenuSpec> specs)
{
    std::vector<MenuFlyoutItem> items;
    items.reserve(specs.size());
    for (const HelperMenuSpec& spec : specs)
    {
        if (spec.separatorBefore && ! items.empty())
        {
            items.push_back(MenuFlyoutItem{.kind = MenuItemKind::Separator});
        }
        items.push_back(BuildMenuItem(spec));
    }
    return items;
}

constexpr std::array<HelperMenuSpec, 12> kTemplateSpecs = {
    HelperMenuSpec{TemplateName, IDS_BATCH_RENAME_HELPER_NAME, L"Name", L"{name}"},
    HelperMenuSpec{TemplateStem, IDS_BATCH_RENAME_HELPER_STEM, L"File name without extension", L"{stem}"},
    HelperMenuSpec{TemplateExt, IDS_BATCH_RENAME_HELPER_EXT, L"Extension", L"{ext}"},
    HelperMenuSpec{TemplateExtNoDot, IDS_BATCH_RENAME_HELPER_EXT_NO_DOT, L"Extension without dot", L"{extNoDot}"},
    HelperMenuSpec{TemplateParent, IDS_BATCH_RENAME_HELPER_PARENT, L"Parent folder", L"{parent}"},
    HelperMenuSpec{TemplateRelativeFolderFlat, IDS_BATCH_RENAME_HELPER_RELATIVE_FOLDER_FLAT, L"Relative folder, flattened", L"{relativeFolderFlat}"},
    HelperMenuSpec{TemplateCounter, IDS_BATCH_RENAME_HELPER_COUNTER, L"Counter", L"{counter}", true},
    HelperMenuSpec{TemplateCounterPadded, IDS_BATCH_RENAME_HELPER_COUNTER_PADDED, L"Counter, 3 digits", L"{counter:000}"},
    HelperMenuSpec{TemplateDate, IDS_BATCH_RENAME_HELPER_DATE, L"Date", L"{date:yyyy-MM-dd}", true},
    HelperMenuSpec{TemplateTime, IDS_BATCH_RENAME_HELPER_TIME, L"Time", L"{time:HH-mm-ss}"},
    HelperMenuSpec{TemplateLiteralOpenBrace, IDS_BATCH_RENAME_HELPER_LITERAL_OPEN_BRACE, L"Opening brace", L"{{", true},
    HelperMenuSpec{TemplateLiteralCloseBrace, IDS_BATCH_RENAME_HELPER_LITERAL_CLOSE_BRACE, L"Closing brace", L"}}"},
};

constexpr std::array<HelperMenuSpec, 19> kRegexSpecs = {
    HelperMenuSpec{RegexAnyCharacter, IDS_BATCH_RENAME_HELPER_REGEX_ANY_CHAR, L"Any character", L"."},
    HelperMenuSpec{RegexCharacterSet, IDS_BATCH_RENAME_HELPER_REGEX_CHAR_SET, L"Any from set", L"[abc]"},
    HelperMenuSpec{RegexNegatedSet, IDS_BATCH_RENAME_HELPER_REGEX_NEGATED_SET, L"Any except set", L"[^abc]"},
    HelperMenuSpec{RegexWordBoundary, IDS_BATCH_RENAME_HELPER_REGEX_WORD_BOUNDARY, L"Word boundary", L"\\b"},
    HelperMenuSpec{RegexAlternation, IDS_BATCH_RENAME_HELPER_REGEX_ALTERNATION, L"Or", L"|"},
    HelperMenuSpec{RegexZeroOrMore, IDS_BATCH_RENAME_HELPER_REGEX_ZERO_OR_MORE, L"Zero or more", L"*", true},
    HelperMenuSpec{RegexOneOrMore, IDS_BATCH_RENAME_HELPER_REGEX_ONE_OR_MORE, L"One or more", L"+"},
    HelperMenuSpec{RegexOptional, IDS_BATCH_RENAME_HELPER_REGEX_OPTIONAL, L"Optional", L"?"},
    HelperMenuSpec{RegexGroup, IDS_BATCH_RENAME_HELPER_REGEX_GROUP, L"Subexpression", L"()"},
    HelperMenuSpec{RegexEscapedLiteral, IDS_BATCH_RENAME_HELPER_REGEX_ESCAPED_LITERAL, L"Selected text as literal", L"", true},
    HelperMenuSpec{RegexWordCharacter, IDS_BATCH_RENAME_HELPER_REGEX_WORD_CHAR, L"Word character", L"\\w", true},
    HelperMenuSpec{RegexNonWordCharacter, IDS_BATCH_RENAME_HELPER_REGEX_NON_WORD_CHAR, L"Non-word character", L"\\W"},
    HelperMenuSpec{RegexWhitespace, IDS_BATCH_RENAME_HELPER_REGEX_WHITESPACE, L"Whitespace character", L"\\s"},
    HelperMenuSpec{RegexNonWhitespace, IDS_BATCH_RENAME_HELPER_REGEX_NON_WHITESPACE, L"Non-whitespace character", L"\\S"},
    HelperMenuSpec{RegexDigit, IDS_BATCH_RENAME_HELPER_REGEX_DIGIT, L"Digit character", L"\\d"},
    HelperMenuSpec{RegexNonDigit, IDS_BATCH_RENAME_HELPER_REGEX_NON_DIGIT, L"Non-digit character", L"\\D"},
    HelperMenuSpec{RegexDecimalNumber, IDS_BATCH_RENAME_HELPER_REGEX_DECIMAL, L"Decimal number", L"\\d+", true},
    HelperMenuSpec{RegexHexNumber, IDS_BATCH_RENAME_HELPER_REGEX_HEX, L"Hexadecimal number", L"[0-9A-Fa-f]+"},
    HelperMenuSpec{RegexFileNameSplit, IDS_BATCH_RENAME_HELPER_REGEX_FILE_NAME_SPLIT, L"File name split", L"^(.+?)(\\.[^.]+)?$"},
};

constexpr std::array<HelperMenuSpec, 5> kReplacementSpecs = {
    HelperMenuSpec{ReplaceLiteralDollar, IDS_BATCH_RENAME_HELPER_REPLACE_LITERAL_DOLLAR, L"Literal dollar", L"$$"},
    HelperMenuSpec{ReplaceWholeMatch, IDS_BATCH_RENAME_HELPER_REPLACE_WHOLE_MATCH, L"Whole match", L"$&"},
    HelperMenuSpec{ReplaceSubexpression1, IDS_BATCH_RENAME_HELPER_REPLACE_SUBEXPRESSION_1, L"Matched subexpression $1", L"$1"},
    HelperMenuSpec{ReplaceSubexpression2, IDS_BATCH_RENAME_HELPER_REPLACE_SUBEXPRESSION_2, L"Matched subexpression $2", L"$2"},
    HelperMenuSpec{ReplaceCustomSubexpression, IDS_BATCH_RENAME_HELPER_REPLACE_CUSTOM_SUBEXPRESSION, L"Matched subexpression...", L""},
};

[[nodiscard]] bool IsDynamicHelperCommand(const int commandId) noexcept
{
    return commandId == RegexEscapedLiteral || commandId == ReplaceCustomSubexpression;
}

[[nodiscard]] std::optional<std::wstring_view> FindInsertionText(std::span<const HelperMenuSpec> specs, const int commandId) noexcept
{
    const auto it = std::ranges::find_if(specs, [commandId](const HelperMenuSpec& spec) noexcept { return spec.commandId == commandId; });
    if (it == specs.end())
    {
        return std::nullopt;
    }
    if (IsDynamicHelperCommand(commandId))
    {
        return std::nullopt;
    }
    return it->insertionText;
}

[[nodiscard]] bool IsRegexMetacharacter(const wchar_t ch) noexcept
{
    switch (ch)
    {
        case L'\\':
        case L'^':
        case L'$':
        case L'.':
        case L'|':
        case L'?':
        case L'*':
        case L'+':
        case L'(':
        case L')':
        case L'[':
        case L']':
        case L'{':
        case L'}': return true;
        default: return false;
    }
}

[[nodiscard]] std::wstring EscapeRegexLiteral(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size() * 2u);
    for (const wchar_t ch : text)
    {
        if (IsRegexMetacharacter(ch))
        {
            result.push_back(L'\\');
        }
        result.push_back(ch);
    }
    return result;
}

[[nodiscard]] bool IsAsciiDecimalDigits(std::wstring_view text) noexcept
{
    return ! text.empty() && std::ranges::all_of(text, [](const wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; });
}
} // namespace

namespace BatchRenameMenus
{
std::vector<MenuFlyoutItem> BuildTemplateHelperMenuItems()
{
    return BuildMenuItems(kTemplateSpecs);
}

std::vector<MenuFlyoutItem> BuildRegexSearchHelperMenuItems()
{
    return BuildMenuItems(kRegexSpecs);
}

std::vector<MenuFlyoutItem> BuildReplacementHelperMenuItems()
{
    return BuildMenuItems(kReplacementSpecs);
}

std::vector<MenuFlyoutItem> BuildHelperMenuItems(const HelperMenuKind kind)
{
    switch (kind)
    {
        case HelperMenuKind::Template: return BuildTemplateHelperMenuItems();
        case HelperMenuKind::RegexSearch: return BuildRegexSearchHelperMenuItems();
        case HelperMenuKind::Replacement: return BuildReplacementHelperMenuItems();
        default: return {};
    }
}

int RegexEscapedLiteralHelperCommandId() noexcept
{
    return RegexEscapedLiteral;
}

int ReplacementCustomSubexpressionHelperCommandId() noexcept
{
    return ReplaceCustomSubexpression;
}

std::optional<std::wstring_view> TryGetHelperInsertionText(const int commandId) noexcept
{
    if (const std::optional<std::wstring_view> insertion = FindInsertionText(kTemplateSpecs, commandId); insertion.has_value())
    {
        return insertion;
    }
    if (const std::optional<std::wstring_view> insertion = FindInsertionText(kRegexSpecs, commandId); insertion.has_value())
    {
        return insertion;
    }
    return FindInsertionText(kReplacementSpecs, commandId);
}

std::optional<HelperCommandInsertion> TryBuildDynamicHelperInsertion(const int commandId, std::wstring_view selectedText)
{
    if (commandId == RegexEscapedLiteral)
    {
        if (selectedText.empty())
        {
            return HelperCommandInsertion{.insertionText = L"\\\\", .selectionStart = 1u, .selectionEnd = 2u};
        }
        std::wstring escaped = EscapeRegexLiteral(selectedText);
        const size_t caret   = escaped.size();
        return HelperCommandInsertion{.insertionText = std::move(escaped), .selectionStart = caret, .selectionEnd = caret};
    }

    if (commandId == ReplaceCustomSubexpression)
    {
        if (IsAsciiDecimalDigits(selectedText))
        {
            std::wstring token = L"$";
            token.append(selectedText);
            const size_t caret = token.size();
            return HelperCommandInsertion{.insertionText = std::move(token), .selectionStart = caret, .selectionEnd = caret};
        }
        return HelperCommandInsertion{.insertionText = L"$1", .selectionStart = 1u, .selectionEnd = 2u};
    }

    return std::nullopt;
}

HelperInsertionResult ApplyHelperInsertion(std::wstring_view text, const size_t selectionStart, const size_t selectionEnd, std::wstring_view insertion)
{
    const size_t textLength = text.size();
    size_t start            = std::min(selectionStart, selectionEnd);
    size_t end              = std::max(selectionStart, selectionEnd);
    start                   = std::min(start, textLength);
    end                     = std::min(end, textLength);

    std::wstring result;
    result.reserve(textLength - (end - start) + insertion.size());
    result.append(text.substr(0u, start));
    result.append(insertion);
    result.append(text.substr(end));

    const size_t caret = start + insertion.size();
    return HelperInsertionResult{.text = std::move(result), .selectionStart = caret, .selectionEnd = caret};
}
} // namespace BatchRenameMenus

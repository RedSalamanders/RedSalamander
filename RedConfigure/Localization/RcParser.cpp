#include "RcParser.h"

#include <algorithm>
#include <cwctype>
#include <format>
#include <unordered_set>

namespace
{
[[nodiscard]] bool IsIdentifierStart(wchar_t ch) noexcept
{
    return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') || ch == L'_';
}

[[nodiscard]] bool IsIdentifierChar(wchar_t ch) noexcept
{
    return IsIdentifierStart(ch) || (ch >= L'0' && ch <= L'9');
}

[[nodiscard]] std::wstring_view Trim(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
    {
        text.remove_prefix(1u);
    }
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
    {
        text.remove_suffix(1u);
    }
    return text;
}

[[nodiscard]] bool EqualsToken(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }
    for (size_t index = 0; index < lhs.size(); ++index)
    {
        if (std::towlower(static_cast<wint_t>(lhs[index])) != std::towlower(static_cast<wint_t>(rhs[index])))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool StartsWithToken(std::wstring_view text, std::wstring_view token) noexcept
{
    text = Trim(text);
    if (text.size() < token.size())
    {
        return false;
    }
    if (! EqualsToken(text.substr(0u, token.size()), token))
    {
        return false;
    }
    return text.size() == token.size() || std::iswspace(static_cast<wint_t>(text[token.size()])) != 0;
}

[[nodiscard]] std::wstring StripComments(std::wstring_view line, bool& inBlockComment)
{
    std::wstring result;
    result.reserve(line.size());

    bool inString = false;
    for (size_t index = 0; index < line.size(); ++index)
    {
        const wchar_t ch   = line[index];
        const wchar_t next = (index + 1u < line.size()) ? line[index + 1u] : L'\0';

        if (inBlockComment)
        {
            if (ch == L'*' && next == L'/')
            {
                inBlockComment = false;
                ++index;
            }
            continue;
        }

        if (! inString && ch == L'/' && next == L'*')
        {
            inBlockComment = true;
            ++index;
            continue;
        }
        if (! inString && ch == L'/' && next == L'/')
        {
            break;
        }

        result.push_back(ch);
        if (ch == L'"')
        {
            if (inString && next == L'"')
            {
                result.push_back(next);
                ++index;
                continue;
            }
            inString = ! inString;
        }
    }

    return result;
}

void SkipWhitespace(std::wstring_view text, size_t& pos) noexcept
{
    while (pos < text.size() && std::iswspace(static_cast<wint_t>(text[pos])) != 0)
    {
        ++pos;
    }
}

[[nodiscard]] bool ParseRcStringLiteral(std::wstring_view text, size_t& pos, std::wstring& outText)
{
    outText.clear();
    SkipWhitespace(text, pos);
    if (pos < text.size() && text[pos] == L'L')
    {
        ++pos;
    }
    if (pos >= text.size() || text[pos] != L'"')
    {
        return false;
    }
    ++pos;

    while (pos < text.size())
    {
        const wchar_t ch = text[pos++];
        if (ch == L'"')
        {
            if (pos < text.size() && text[pos] == L'"')
            {
                outText.push_back(L'"');
                ++pos;
                continue;
            }
            return true;
        }

        if (ch == L'\\' && pos < text.size())
        {
            const wchar_t escaped = text[pos++];
            switch (escaped)
            {
            case L'n':
                outText.push_back(L'\n');
                break;
            case L'r':
                outText.push_back(L'\r');
                break;
            case L't':
                outText.push_back(L'\t');
                break;
            case L'"':
                outText.push_back(L'"');
                break;
            case L'\\':
                outText.push_back(L'\\');
                break;
            default:
                outText.push_back(escaped);
                break;
            }
            continue;
        }

        outText.push_back(ch);
    }

    return false;
}

[[nodiscard]] bool ReadIdentifier(std::wstring_view text, size_t& pos, std::wstring& out)
{
    out.clear();
    SkipWhitespace(text, pos);
    if (pos >= text.size() || ! IsIdentifierStart(text[pos]))
    {
        return false;
    }

    const size_t start = pos;
    while (pos < text.size() && IsIdentifierChar(text[pos]))
    {
        ++pos;
    }

    out.assign(text.substr(start, pos - start));
    return true;
}

[[nodiscard]] bool ReadResourceId(std::wstring_view text, size_t& pos, std::wstring& out)
{
    out.clear();
    SkipWhitespace(text, pos);
    if (pos >= text.size())
    {
        return false;
    }

    if (IsIdentifierStart(text[pos]))
    {
        return ReadIdentifier(text, pos, out);
    }

    if (text[pos] >= L'0' && text[pos] <= L'9')
    {
        const size_t start = pos;
        while (pos < text.size() && text[pos] >= L'0' && text[pos] <= L'9')
        {
            ++pos;
        }
        out.assign(text.substr(start, pos - start));
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring ReadFirstToken(std::wstring_view line)
{
    size_t pos = 0u;
    std::wstring token;
    static_cast<void>(ReadIdentifier(line, pos, token));
    return token;
}

[[nodiscard]] bool ParseResourceDeclaration(std::wstring_view line, std::wstring& outResourceId, std::wstring& outResourceType)
{
    size_t pos = 0u;
    if (! ReadIdentifier(line, pos, outResourceId))
    {
        return false;
    }
    return ReadIdentifier(line, pos, outResourceType);
}

[[nodiscard]] bool ParseStringTableEntry(std::wstring_view line, RedConfigure::Localization::RcStringEntry& outEntry)
{
    size_t pos = 0u;
    SkipWhitespace(line, pos);
    if (pos >= line.size() || ! IsIdentifierStart(line[pos]))
    {
        return false;
    }

    const size_t idStart = pos;
    while (pos < line.size() && IsIdentifierChar(line[pos]))
    {
        ++pos;
    }

    outEntry.id.assign(line.substr(idStart, pos - idStart));
    SkipWhitespace(line, pos);
    if (pos < line.size() && line[pos] == L',')
    {
        ++pos;
    }

    return ParseRcStringLiteral(line, pos, outEntry.text);
}

void AddStringLocalizableEntry(const RedConfigure::Localization::RcStringEntry& stringEntry, RedConfigure::Localization::RcParseResult& result)
{
    result.localizableEntries.push_back(RedConfigure::Localization::RcLocalizableEntry{
        .kind       = RedConfigure::Localization::RcLocalizableKind::StringTable,
        .ownerId    = {},
        .id         = stringEntry.id,
        .text       = stringEntry.text,
        .sourceLine = stringEntry.sourceLine,
        .duplicate  = stringEntry.duplicate,
    });
}

[[nodiscard]] bool ParseMenuPopupEntry(std::wstring_view line,
                                       std::wstring_view ownerId,
                                       size_t sourceLine,
                                       RedConfigure::Localization::RcLocalizableEntry& outEntry)
{
    if (! StartsWithToken(line, L"POPUP"))
    {
        return false;
    }

    size_t pos = std::wstring_view(L"POPUP").size();
    std::wstring text;
    if (! ParseRcStringLiteral(line, pos, text))
    {
        return false;
    }

    outEntry = RedConfigure::Localization::RcLocalizableEntry{
        .kind       = RedConfigure::Localization::RcLocalizableKind::MenuPopup,
        .ownerId    = std::wstring(ownerId),
        .id         = {},
        .text       = std::move(text),
        .sourceLine = sourceLine,
    };
    return true;
}

[[nodiscard]] bool ParseMenuItemEntry(std::wstring_view line,
                                      std::wstring_view ownerId,
                                      size_t sourceLine,
                                      RedConfigure::Localization::RcLocalizableEntry& outEntry)
{
    if (! StartsWithToken(line, L"MENUITEM"))
    {
        return false;
    }

    size_t pos = std::wstring_view(L"MENUITEM").size();
    SkipWhitespace(line, pos);
    if (StartsWithToken(line.substr(pos), L"SEPARATOR"))
    {
        return false;
    }

    std::wstring text;
    if (! ParseRcStringLiteral(line, pos, text))
    {
        return false;
    }

    SkipWhitespace(line, pos);
    if (pos < line.size() && line[pos] == L',')
    {
        ++pos;
    }

    std::wstring id;
    if (! ReadResourceId(line, pos, id))
    {
        return false;
    }

    outEntry = RedConfigure::Localization::RcLocalizableEntry{
        .kind       = RedConfigure::Localization::RcLocalizableKind::MenuItem,
        .ownerId    = std::wstring(ownerId),
        .id         = std::move(id),
        .text       = std::move(text),
        .sourceLine = sourceLine,
    };
    return true;
}

[[nodiscard]] bool ParseDialogCaptionEntry(std::wstring_view line,
                                           std::wstring_view ownerId,
                                           size_t sourceLine,
                                           RedConfigure::Localization::RcLocalizableEntry& outEntry)
{
    if (! StartsWithToken(line, L"CAPTION"))
    {
        return false;
    }

    size_t pos = std::wstring_view(L"CAPTION").size();
    std::wstring text;
    if (! ParseRcStringLiteral(line, pos, text))
    {
        return false;
    }

    outEntry = RedConfigure::Localization::RcLocalizableEntry{
        .kind       = RedConfigure::Localization::RcLocalizableKind::DialogCaption,
        .ownerId    = std::wstring(ownerId),
        .id         = std::wstring(ownerId),
        .text       = std::move(text),
        .sourceLine = sourceLine,
    };
    return true;
}

[[nodiscard]] bool IsSupportedDialogControlToken(std::wstring_view token) noexcept
{
    constexpr std::wstring_view supported[] = {
        L"LTEXT",
        L"CTEXT",
        L"RTEXT",
        L"PUSHBUTTON",
        L"DEFPUSHBUTTON",
        L"GROUPBOX",
        L"CONTROL",
        L"CHECKBOX",
        L"AUTOCHECKBOX",
        L"RADIOBUTTON",
        L"AUTORADIOBUTTON",
    };

    return std::any_of(std::begin(supported), std::end(supported), [token](std::wstring_view item) noexcept { return EqualsToken(token, item); });
}

[[nodiscard]] bool ParseDialogControlEntry(std::wstring_view line,
                                           std::wstring_view ownerId,
                                           size_t sourceLine,
                                           RedConfigure::Localization::RcLocalizableEntry& outEntry)
{
    const std::wstring token = ReadFirstToken(line);
    if (! IsSupportedDialogControlToken(token))
    {
        return false;
    }

    size_t pos = token.size();
    std::wstring text;
    if (! ParseRcStringLiteral(line, pos, text))
    {
        return false;
    }

    SkipWhitespace(line, pos);
    if (pos < line.size() && line[pos] == L',')
    {
        ++pos;
    }

    std::wstring id;
    if (! ReadResourceId(line, pos, id))
    {
        return false;
    }

    outEntry = RedConfigure::Localization::RcLocalizableEntry{
        .kind       = RedConfigure::Localization::RcLocalizableKind::DialogControl,
        .ownerId    = std::wstring(ownerId),
        .id         = std::move(id),
        .text       = std::move(text),
        .sourceLine = sourceLine,
    };
    return true;
}
} // namespace

namespace RedConfigure::Localization
{
HRESULT ParseRcStringTables(std::wstring_view text, RcParseResult& outResult)
{
    outResult = {};

    bool inBlockComment       = false;
    bool awaitingStringBegin  = false;
    bool awaitingMenuBegin    = false;
    bool awaitingDialogBegin  = false;
    bool inStringTable        = false;
    bool inMenu               = false;
    bool inDialog             = false;
    size_t menuDepth          = 0u;
    size_t dialogDepth        = 0u;
    std::wstring activeMenuId;
    std::wstring activeDialogId;
    std::unordered_set<std::wstring> seenIds;

    size_t lineStart = 0u;
    size_t lineNo    = 1u;
    while (lineStart <= text.size())
    {
        size_t lineEnd = text.find(L'\n', lineStart);
        if (lineEnd == std::wstring_view::npos)
        {
            lineEnd = text.size();
        }

        std::wstring_view rawLine = text.substr(lineStart, lineEnd - lineStart);
        if (! rawLine.empty() && rawLine.back() == L'\r')
        {
            rawLine.remove_suffix(1u);
        }

        const std::wstring stripped = StripComments(rawLine, inBlockComment);
        const std::wstring_view line = Trim(stripped);

        if (inStringTable)
        {
            if (StartsWithToken(line, L"END"))
            {
                inStringTable = false;
            }
            else if (! line.empty())
            {
                RcStringEntry entry;
                entry.sourceLine = lineNo;
                if (ParseStringTableEntry(line, entry))
                {
                    const auto [_, inserted] = seenIds.insert(entry.id);
                    entry.duplicate          = ! inserted;
                    if (entry.duplicate)
                    {
                        outResult.errors.push_back(std::format(L"Duplicate resource id: {}", entry.id));
                    }
                    AddStringLocalizableEntry(entry, outResult);
                    outResult.strings.push_back(std::move(entry));
                }
            }
        }
        else if (inMenu)
        {
            if (StartsWithToken(line, L"BEGIN"))
            {
                ++menuDepth;
            }
            else if (StartsWithToken(line, L"END"))
            {
                if (menuDepth > 0u)
                {
                    --menuDepth;
                }
                if (menuDepth == 0u)
                {
                    inMenu = false;
                    activeMenuId.clear();
                }
            }
            else if (! line.empty())
            {
                RcLocalizableEntry entry;
                if (ParseMenuPopupEntry(line, activeMenuId, lineNo, entry) || ParseMenuItemEntry(line, activeMenuId, lineNo, entry))
                {
                    outResult.localizableEntries.push_back(std::move(entry));
                }
            }
        }
        else if (inDialog)
        {
            if (StartsWithToken(line, L"BEGIN"))
            {
                ++dialogDepth;
            }
            else if (StartsWithToken(line, L"END"))
            {
                if (dialogDepth > 0u)
                {
                    --dialogDepth;
                }
                if (dialogDepth == 0u)
                {
                    inDialog = false;
                    activeDialogId.clear();
                }
            }
            else if (! line.empty())
            {
                RcLocalizableEntry entry;
                if (ParseDialogControlEntry(line, activeDialogId, lineNo, entry))
                {
                    outResult.localizableEntries.push_back(std::move(entry));
                }
            }
        }
        else if (awaitingStringBegin)
        {
            if (StartsWithToken(line, L"BEGIN"))
            {
                awaitingStringBegin = false;
                inStringTable       = true;
            }
        }
        else if (awaitingMenuBegin)
        {
            if (StartsWithToken(line, L"BEGIN"))
            {
                awaitingMenuBegin = false;
                inMenu            = true;
                menuDepth         = 1u;
            }
        }
        else if (awaitingDialogBegin)
        {
            if (StartsWithToken(line, L"BEGIN"))
            {
                awaitingDialogBegin = false;
                inDialog            = true;
                dialogDepth         = 1u;
            }
            else if (! line.empty())
            {
                RcLocalizableEntry entry;
                if (ParseDialogCaptionEntry(line, activeDialogId, lineNo, entry))
                {
                    outResult.localizableEntries.push_back(std::move(entry));
                }
            }
        }
        else if (StartsWithToken(line, L"STRINGTABLE"))
        {
            awaitingStringBegin = true;
        }
        else if (! line.empty())
        {
            std::wstring resourceId;
            std::wstring resourceType;
            if (ParseResourceDeclaration(line, resourceId, resourceType))
            {
                if (EqualsToken(resourceType, L"MENU") || EqualsToken(resourceType, L"MENUEX"))
                {
                    activeMenuId     = std::move(resourceId);
                    awaitingMenuBegin = true;
                }
                else if (EqualsToken(resourceType, L"DIALOG") || EqualsToken(resourceType, L"DIALOGEX"))
                {
                    activeDialogId     = std::move(resourceId);
                    awaitingDialogBegin = true;
                }
            }
        }

        if (lineEnd == text.size())
        {
            break;
        }
        lineStart = lineEnd + 1u;
        ++lineNo;
    }

    return S_OK;
}
} // namespace RedConfigure::Localization

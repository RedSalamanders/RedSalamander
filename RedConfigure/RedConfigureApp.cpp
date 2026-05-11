#include "RedConfigureApp.h"

#include "resource.h"
#include "SettingsStore.h"

#include <algorithm>
#include <array>
#include <set>
#include <system_error>

namespace
{
constexpr std::array<RedConfigure::PageDefinition, 4> kPageDefinitions = {{
    {L"start", IDS_REDCONFIGURE_MODE_START_TITLE, IDS_REDCONFIGURE_MODE_START_DESC},
    {L"localization", IDS_REDCONFIGURE_MODE_LOCALIZATION_TITLE, IDS_REDCONFIGURE_MODE_LOCALIZATION_DESC},
    {L"themes", IDS_REDCONFIGURE_MODE_THEMES_TITLE, IDS_REDCONFIGURE_MODE_THEMES_DESC},
    {L"reviewExport", IDS_REDCONFIGURE_MODE_REVIEW_EXPORT_TITLE, IDS_REDCONFIGURE_MODE_REVIEW_EXPORT_DESC},
}};

[[nodiscard]] wchar_t ToLowerAscii(wchar_t ch) noexcept
{
    if (ch >= L'A' && ch <= L'Z')
    {
        return static_cast<wchar_t>(ch - L'A' + L'a');
    }
    return ch;
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }
    if (needle.size() > text.size())
    {
        return false;
    }

    for (size_t index = 0u; index + needle.size() <= text.size(); ++index)
    {
        bool matched = true;
        for (size_t offset = 0u; offset < needle.size(); ++offset)
        {
            if (ToLowerAscii(text[index + offset]) != ToLowerAscii(needle[offset]))
            {
                matched = false;
                break;
            }
        }
        if (matched)
        {
            return true;
        }
    }
    return false;
}
} // namespace

namespace RedConfigure
{
std::span<const PageDefinition> GetPageDefinitions() noexcept
{
    return kPageDefinitions;
}

std::filesystem::path ResolveWorkspaceRootForLaunchPath(const std::filesystem::path& startPath)
{
    std::filesystem::path candidate = startPath;
    if (candidate.empty())
    {
        return candidate;
    }

    std::error_code ec;
    if (std::filesystem::is_regular_file(candidate, ec))
    {
        candidate = candidate.parent_path();
    }

    for (std::filesystem::path current = candidate; ! current.empty(); current = current.parent_path())
    {
        ec.clear();
        if (std::filesystem::exists(current / L"RedSalamander.sln", ec) || std::filesystem::exists(current / L".git", ec))
        {
            return current;
        }

        const std::filesystem::path parent = current.parent_path();
        if (parent == current)
        {
            break;
        }
    }

    return startPath;
}

std::vector<std::wstring> FilterThemeColorKeys(std::span<const std::wstring> keys, std::wstring_view filterText)
{
    if (filterText.empty())
    {
        return std::vector<std::wstring>(keys.begin(), keys.end());
    }

    std::vector<std::wstring> filtered;
    filtered.reserve(keys.size());
    for (const std::wstring& key : keys)
    {
        if (ContainsIgnoreCase(key, filterText))
        {
            filtered.push_back(key);
        }
    }
    return filtered;
}

std::wstring SelectThemePreviewHitKey(std::span<const ThemePreviewHitCandidate> candidates,
                                      float x,
                                      float y,
                                      std::wstring_view previousKey)
{
    struct Match
    {
        std::wstring_view key;
        float area = 0.0f;
        size_t order = 0u;
    };

    std::vector<Match> matches;
    matches.reserve(candidates.size());
    for (size_t index = 0u; index < candidates.size(); ++index)
    {
        const ThemePreviewHitCandidate& candidate = candidates[index];
        if (x < candidate.left || x > candidate.right || y < candidate.top || y > candidate.bottom)
        {
            continue;
        }

        const float width  = (std::max)(0.0f, candidate.right - candidate.left);
        const float height = (std::max)(0.0f, candidate.bottom - candidate.top);
        matches.push_back(Match{.key = candidate.key, .area = width * height, .order = index});
    }

    if (matches.empty())
    {
        return {};
    }

    std::stable_sort(matches.begin(),
                     matches.end(),
                     [](const Match& lhs, const Match& rhs) noexcept
                     {
                         if (lhs.area != rhs.area)
                         {
                             return lhs.area < rhs.area;
                         }
                         return lhs.order < rhs.order;
                     });

    if (! previousKey.empty())
    {
        for (size_t index = 0u; index < matches.size(); ++index)
        {
            if (matches[index].key == previousKey)
            {
                return std::wstring(matches[(index + 1u) % matches.size()].key);
            }
        }
    }

    return std::wstring(matches.front().key);
}

std::vector<std::wstring> BuildThemeColorSuggestions(std::wstring_view selectedKey,
                                                     std::wstring_view previousKey,
                                                     std::optional<uint32_t> currentColor)
{
    std::vector<std::wstring> suggestions;
    std::set<std::wstring> added;
    const auto add = [&suggestions, &added](std::wstring value)
    {
        if (! value.empty() && added.insert(value).second)
        {
            suggestions.push_back(std::move(value));
        }
    };

    if (currentColor)
    {
        add(Common::Settings::FormatColor(currentColor.value()));
    }

    const auto addReferenceTemplates = [&add](std::wstring_view key)
    {
        if (key.empty())
        {
            return;
        }

        const std::wstring keyText(key);
        add(L"ref(" + keyText + L")");
        add(L"darken(" + keyText + L",20%)");
        add(L"lighten(" + keyText + L",20%)");
        add(L"alpha(" + keyText + L",80%)");
        add(L"contrast(" + keyText + L")");
    };

    if (selectedKey != L"app.accent")
    {
        addReferenceTemplates(L"app.accent");
    }
    else
    {
        addReferenceTemplates(L"window.background");
    }

    if (! previousKey.empty() && previousKey != selectedKey)
    {
        addReferenceTemplates(previousKey);
        add(L"blend(" + std::wstring(previousKey) + L",app.accent,16%)");
    }

    add(L"blend(menu.background,app.accent,16%)");
    add(L"blend(folderView.background,app.accent,12%)");
    return suggestions;
}
} // namespace RedConfigure

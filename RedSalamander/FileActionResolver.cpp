#include "FileActionResolver.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cstdint>
#include <cwctype>
#include <format>
#include <limits>
#include <optional>
#include <ranges>

#include "Helpers.h"
#include "Resource.h"

namespace
{
[[nodiscard]] wchar_t ToLower(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towlower(ch));
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t index = 0; index < lhs.size(); ++index)
    {
        if (ToLower(lhs[index]) != ToLower(rhs[index]))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool ContainsNoCase(const std::vector<std::wstring>& values, std::wstring_view text) noexcept
{
    return std::find_if(values.begin(), values.end(), [&](const std::wstring& value) noexcept { return EqualsNoCase(value, text); }) != values.end();
}

[[nodiscard]] std::wstring NormalizeExtension(const std::filesystem::path& path)
{
    std::wstring ext = path.extension().wstring();
    std::ranges::transform(ext, ext.begin(), [](wchar_t ch) noexcept { return ToLower(ch); });
    return ext;
}

[[nodiscard]] bool WildcardMatchNoCase(std::wstring_view pattern, std::wstring_view text) noexcept
{
    size_t patternIndex = 0;
    size_t textIndex    = 0;
    size_t starIndex    = std::wstring_view::npos;
    size_t starText     = 0;

    while (textIndex < text.size())
    {
        if (patternIndex < pattern.size() &&
            (pattern[patternIndex] == L'?' || ToLower(pattern[patternIndex]) == ToLower(text[textIndex])))
        {
            ++patternIndex;
            ++textIndex;
            continue;
        }

        if (patternIndex < pattern.size() && pattern[patternIndex] == L'*')
        {
            starIndex = patternIndex++;
            starText  = textIndex;
            continue;
        }

        if (starIndex != std::wstring_view::npos)
        {
            patternIndex = starIndex + 1u;
            textIndex    = ++starText;
            continue;
        }

        return false;
    }

    while (patternIndex < pattern.size() && pattern[patternIndex] == L'*')
    {
        ++patternIndex;
    }

    return patternIndex == pattern.size();
}

[[nodiscard]] bool MatchFileActionRule(const Common::Settings::FileActionMatch& match, const std::filesystem::path& itemPath)
{
    switch (match.kind)
    {
        case Common::Settings::FileActionMatchKind::Default:
            return true;
        case Common::Settings::FileActionMatchKind::Extension:
            return ! match.value.empty() && EqualsNoCase(NormalizeExtension(itemPath), match.value);
        case Common::Settings::FileActionMatchKind::Pattern:
            return ! match.value.empty() && WildcardMatchNoCase(match.value, itemPath.filename().wstring());
    }

    return false;
}

[[nodiscard]] bool HasSpecificFileMatch(const Common::Settings::FileActionMatch& match) noexcept
{
    return match.kind == Common::Settings::FileActionMatchKind::Extension || match.kind == Common::Settings::FileActionMatchKind::Pattern;
}

[[nodiscard]] FileActionResolver::Reason ReasonForMatch(const Common::Settings::FileActionMatch& match, std::wstring_view computerName) noexcept
{
    const bool computerSpecific = ! computerName.empty();
    if (HasSpecificFileMatch(match))
    {
        return computerSpecific ? FileActionResolver::Reason::ComputerExtensionRule : FileActionResolver::Reason::GlobalExtensionRule;
    }
    return computerSpecific ? FileActionResolver::Reason::ComputerDefaultRule : FileActionResolver::Reason::GlobalDefaultRule;
}

[[nodiscard]] uint32_t ReasonPriority(FileActionResolver::Reason reason) noexcept
{
    switch (reason)
    {
        case FileActionResolver::Reason::ComputerExtensionRule: return 0u;
        case FileActionResolver::Reason::GlobalExtensionRule: return 1u;
        case FileActionResolver::Reason::ComputerDefaultRule: return 2u;
        case FileActionResolver::Reason::GlobalDefaultRule: return 3u;
        case FileActionResolver::Reason::None:
        case FileActionResolver::Reason::NoAssociation:
        case FileActionResolver::Reason::ActionMissing:
        case FileActionResolver::Reason::ActionDisabled:
        case FileActionResolver::Reason::ActionNotApplicable:
            return std::numeric_limits<uint32_t>::max();
    }
    return std::numeric_limits<uint32_t>::max();
}

[[nodiscard]] std::wstring DescribeReason(FileActionResolver::Reason reason,
                                          const Common::Settings::FileActionMatch& match,
                                          std::wstring_view computerName)
{
    const std::wstring matchText = match.kind == Common::Settings::FileActionMatchKind::Default ? L"*" : match.value;
    switch (reason)
    {
        case FileActionResolver::Reason::ComputerExtensionRule:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_COMPUTER_RULE_FMT, computerName, matchText);
        case FileActionResolver::Reason::GlobalExtensionRule:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_GLOBAL_RULE_FMT, matchText);
        case FileActionResolver::Reason::ComputerDefaultRule:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_COMPUTER_DEFAULT_FMT, computerName);
        case FileActionResolver::Reason::GlobalDefaultRule:
            return LoadStringResource(nullptr, IDS_FILEACTION_REASON_GLOBAL_DEFAULT);
        case FileActionResolver::Reason::ActionMissing:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_MISSING_FMT, matchText);
        case FileActionResolver::Reason::ActionDisabled:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_DISABLED_FMT, matchText);
        case FileActionResolver::Reason::ActionNotApplicable:
            return FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_NOT_APPLICABLE_FMT, matchText);
        case FileActionResolver::Reason::NoAssociation:
            return LoadStringResource(nullptr, IDS_FILEACTION_REASON_NO_ASSOCIATION);
        case FileActionResolver::Reason::None:
            break;
    }
    return {};
}

[[nodiscard]] std::wstring_view CommandMetricDetail(FileActionResolver::Command command) noexcept
{
    switch (command)
    {
        case FileActionResolver::Command::View: return L"view";
        case FileActionResolver::Command::AlternateView: return L"alternateView";
        case FileActionResolver::Command::Edit: return L"edit";
        case FileActionResolver::Command::AlternateEdit: return L"alternateEdit";
        case FileActionResolver::Command::EditNew: return L"editNew";
    }
    return L"unknown";
}

[[nodiscard]] HRESULT ResolutionHr(const FileActionResolver::Resolution& resolution) noexcept
{
    if (resolution.IsResolved())
    {
        return S_OK;
    }

    switch (resolution.reason)
    {
        case FileActionResolver::Reason::NoAssociation: return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        case FileActionResolver::Reason::ActionMissing: return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        case FileActionResolver::Reason::ActionDisabled: return E_ACCESSDENIED;
        case FileActionResolver::Reason::ActionNotApplicable: return E_ACCESSDENIED;
        case FileActionResolver::Reason::None:
        case FileActionResolver::Reason::ComputerExtensionRule:
        case FileActionResolver::Reason::GlobalExtensionRule:
        case FileActionResolver::Reason::ComputerDefaultRule:
        case FileActionResolver::Reason::GlobalDefaultRule:
            return E_FAIL;
    }
    return E_FAIL;
}

[[nodiscard]] const Common::Settings::FileActionDefinition* FindActionById(
    const std::vector<Common::Settings::FileActionDefinition>& actions, std::wstring_view actionId) noexcept
{
    if (actionId.empty())
    {
        return nullptr;
    }

    const auto it = std::find_if(actions.begin(), actions.end(), [&](const Common::Settings::FileActionDefinition& action) noexcept {
        return EqualsNoCase(action.id, actionId);
    });
    return it == actions.end() ? nullptr : &(*it);
}

template <typename Rule>
[[nodiscard]] bool RuleMatchesContext(const Rule& rule, const std::filesystem::path& itemPath, std::wstring_view computerName)
{
    if (! rule.computerName.empty() && ! EqualsNoCase(rule.computerName, computerName))
    {
        return false;
    }

    return MatchFileActionRule(rule.match, itemPath);
}

[[nodiscard]] std::wstring SelectActionId(const Common::Settings::ViewerAssociationRule& rule, FileActionResolver::Command command)
{
    switch (command)
    {
        case FileActionResolver::Command::View:
            return rule.viewActionId;
        case FileActionResolver::Command::AlternateView:
            return rule.alternateViewActionId;
        case FileActionResolver::Command::Edit:
        case FileActionResolver::Command::AlternateEdit:
        case FileActionResolver::Command::EditNew:
            return {};
    }
    return {};
}

[[nodiscard]] std::wstring SelectActionId(const Common::Settings::EditorAssociationRule& rule, FileActionResolver::Command command)
{
    switch (command)
    {
        case FileActionResolver::Command::Edit:
            return rule.editActionId;
        case FileActionResolver::Command::AlternateEdit:
            return rule.alternateEditActionId;
        case FileActionResolver::Command::EditNew:
            return rule.editNewActionId;
        case FileActionResolver::Command::View:
        case FileActionResolver::Command::AlternateView:
            return {};
    }
    return {};
}

template <typename Settings>
[[nodiscard]] FileActionResolver::Resolution ResolveAction(const Settings& settings, const FileActionResolver::Request& request)
{
    struct Candidate
    {
        size_t index = 0;
        FileActionResolver::Reason reason = FileActionResolver::Reason::None;
    };

    std::optional<Candidate> best;
    for (size_t index = 0; index < settings.associations.size(); ++index)
    {
        const auto& rule = settings.associations[index];
        if (! RuleMatchesContext(rule, request.filePath, request.computerName))
        {
            continue;
        }

        const FileActionResolver::Reason reason = ReasonForMatch(rule.match, rule.computerName);
        if (! best.has_value() || ReasonPriority(reason) < ReasonPriority(best.value().reason))
        {
            best = Candidate{.index = index, .reason = reason};
        }
    }

    FileActionResolver::Resolution resolution{};
    if (! best.has_value())
    {
        resolution.reason     = FileActionResolver::Reason::NoAssociation;
        resolution.reasonText = DescribeReason(resolution.reason, Common::Settings::FileActionMatch{}, {});
        return resolution;
    }

    const auto& rule = settings.associations[best.value().index];
    resolution.reason   = best.value().reason;
    resolution.actionId = SelectActionId(rule, request.command);
    if (resolution.actionId.empty())
    {
        resolution.reason     = FileActionResolver::Reason::NoAssociation;
        resolution.reasonText = DescribeReason(resolution.reason, rule.match, rule.computerName);
        return resolution;
    }

    resolution.action = FindActionById(settings.actions, resolution.actionId);
    if (! resolution.action)
    {
        resolution.reason     = FileActionResolver::Reason::ActionMissing;
        resolution.reasonText = FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_MISSING_FMT, resolution.actionId);
        return resolution;
    }
    if (! resolution.action->enabled)
    {
        resolution.reason     = FileActionResolver::Reason::ActionDisabled;
        resolution.reasonText = FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_DISABLED_FMT, resolution.actionId);
        resolution.action     = nullptr;
        return resolution;
    }
    if (! FileActionResolver::ActionAppliesToContext(*resolution.action, request.filePath, request.computerName))
    {
        resolution.reason     = FileActionResolver::Reason::ActionNotApplicable;
        resolution.reasonText = FormatStringResource(nullptr, IDS_FILEACTION_REASON_ACTION_NOT_APPLICABLE_FMT, resolution.actionId);
        resolution.action     = nullptr;
        return resolution;
    }

    resolution.reasonText = DescribeReason(resolution.reason, rule.match, rule.computerName);
    return resolution;
}
} // namespace

namespace FileActionResolver
{
bool ActionAppliesToContext(const Common::Settings::FileActionDefinition& action,
                            const std::filesystem::path& itemPath,
                            std::wstring_view computerName)
{
    if (! action.appliesTo.computerNames.empty() && ! ContainsNoCase(action.appliesTo.computerNames, computerName))
    {
        return false;
    }

    if (! action.appliesTo.matches.empty())
    {
        return std::ranges::any_of(action.appliesTo.matches, [&](const Common::Settings::FileActionMatch& match) {
            return MatchFileActionRule(match, itemPath);
        });
    }

    return true;
}

Resolution ResolveViewerAction(const Common::Settings::ViewerFileActionsSettings& settings, const Request& request)
{
    const auto startedAt = std::chrono::steady_clock::now();
    Resolution resolution = ResolveAction(settings, request);
    Debug::Perf::Emit(L"fileaction.resolve_us",
                      CommandMetricDetail(request.command),
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(settings.associations.size()),
                      static_cast<uint64_t>(settings.actions.size()),
                      ResolutionHr(resolution));
    return resolution;
}

Resolution ResolveEditorAction(const Common::Settings::EditorFileActionsSettings& settings, const Request& request)
{
    const auto startedAt = std::chrono::steady_clock::now();
    Resolution resolution = ResolveAction(settings, request);
    Debug::Perf::Emit(L"fileaction.resolve_us",
                      CommandMetricDetail(request.command),
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(settings.associations.size()),
                      static_cast<uint64_t>(settings.actions.size()),
                      ResolutionHr(resolution));
    return resolution;
}

std::vector<const Common::Settings::FileActionDefinition*> CollectAssociatedEditorActions(
    const Common::Settings::EditorFileActionsSettings& settings,
    const Request& request)
{
    struct Candidate final
    {
        const Common::Settings::FileActionDefinition* action = nullptr;
        uint32_t priority                                     = std::numeric_limits<uint32_t>::max();
        size_t associationIndex                               = 0u;
    };

    std::vector<Candidate> candidates;
    candidates.reserve(settings.associations.size());

    for (size_t associationIndex = 0; associationIndex < settings.associations.size(); ++associationIndex)
    {
        const auto& rule = settings.associations[associationIndex];
        if (! RuleMatchesContext(rule, request.filePath, request.computerName))
        {
            continue;
        }

        const std::wstring actionId = SelectActionId(rule, request.command);
        if (actionId.empty())
        {
            continue;
        }

        const Common::Settings::FileActionDefinition* action =
            FindApplicableActionById(settings.actions, actionId, request.filePath, request.computerName);
        if (! action)
        {
            continue;
        }

        const bool alreadyIncluded = std::ranges::any_of(candidates, [&](const Candidate& candidate) noexcept {
            return candidate.action && EqualsNoCase(candidate.action->id, action->id);
        });
        if (alreadyIncluded)
        {
            continue;
        }

        candidates.push_back(Candidate{
            .action           = action,
            .priority         = ReasonPriority(ReasonForMatch(rule.match, rule.computerName)),
            .associationIndex = associationIndex,
        });
    }

    std::stable_sort(candidates.begin(), candidates.end(), [](const Candidate& lhs, const Candidate& rhs) noexcept {
        if (lhs.priority != rhs.priority)
        {
            return lhs.priority < rhs.priority;
        }

        return lhs.associationIndex < rhs.associationIndex;
    });

    std::vector<const Common::Settings::FileActionDefinition*> actions;
    actions.reserve(candidates.size());
    for (const Candidate& candidate : candidates)
    {
        actions.push_back(candidate.action);
    }
    return actions;
}

const Common::Settings::FileActionDefinition* FindApplicableActionById(const std::vector<Common::Settings::FileActionDefinition>& actions,
                                                                       std::wstring_view actionId,
                                                                       const std::filesystem::path& itemPath,
                                                                       std::wstring_view computerName)
{
    const Common::Settings::FileActionDefinition* action = FindActionById(actions, actionId);
    if (! action || ! action->enabled || ! ActionAppliesToContext(*action, itemPath, computerName))
    {
        return nullptr;
    }
    return action;
}

std::vector<const Common::Settings::FileActionDefinition*> CollectApplicableActions(
    const std::vector<Common::Settings::FileActionDefinition>& actions,
    const std::filesystem::path& itemPath,
    std::wstring_view computerName)
{
    const auto startedAt = std::chrono::steady_clock::now();

    std::vector<const Common::Settings::FileActionDefinition*> applicable;
    applicable.reserve(actions.size());
    for (const Common::Settings::FileActionDefinition& action : actions)
    {
        if (action.enabled && ActionAppliesToContext(action, itemPath, computerName))
        {
            applicable.push_back(&action);
        }
    }

    Debug::Perf::Emit(L"fileaction.collect_applicable_us",
                      L"",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(actions.size()),
                      static_cast<uint64_t>(applicable.size()),
                      S_OK);
    return applicable;
}
} // namespace FileActionResolver

#pragma once

#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

#include "SettingsStore.h"

namespace FileActionResolver
{
enum class Command
{
    View,
    AlternateView,
    Edit,
    AlternateEdit,
    EditNew,
};

enum class Reason
{
    None,
    ComputerExtensionRule,
    GlobalExtensionRule,
    ComputerDefaultRule,
    GlobalDefaultRule,
    NoAssociation,
    ActionMissing,
    ActionDisabled,
    ActionNotApplicable,
};

struct Request
{
    Command command = Command::View;
    std::filesystem::path filePath;
    std::wstring computerName;
};

struct Resolution
{
    const Common::Settings::FileActionDefinition* action = nullptr;
    std::wstring actionId;
    Reason reason = Reason::None;
    std::wstring reasonText;

    [[nodiscard]] bool IsResolved() const noexcept
    {
        return action != nullptr;
    }
};

[[nodiscard]] Resolution ResolveViewerAction(const Common::Settings::ViewerFileActionsSettings& settings, const Request& request);
[[nodiscard]] Resolution ResolveEditorAction(const Common::Settings::EditorFileActionsSettings& settings, const Request& request);
[[nodiscard]] std::vector<const Common::Settings::FileActionDefinition*> CollectAssociatedEditorActions(
    const Common::Settings::EditorFileActionsSettings& settings,
    const Request& request);
[[nodiscard]] const Common::Settings::FileActionDefinition* FindApplicableActionById(
    const std::vector<Common::Settings::FileActionDefinition>& actions,
    std::wstring_view actionId,
    const std::filesystem::path& itemPath,
    std::wstring_view computerName);
[[nodiscard]] std::vector<const Common::Settings::FileActionDefinition*> CollectApplicableActions(
    const std::vector<Common::Settings::FileActionDefinition>& actions,
    const std::filesystem::path& itemPath,
    std::wstring_view computerName);
[[nodiscard]] bool ActionAppliesToContext(const Common::Settings::FileActionDefinition& action,
                                          const std::filesystem::path& itemPath,
                                          std::wstring_view computerName);
} // namespace FileActionResolver

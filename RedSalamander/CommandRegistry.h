#pragma once

#include <optional>
#include <span>
#include <cstdint>
#include <string_view>

enum class CommandExecutorOwner : uint8_t
{
    LegacyHost,
    ApplicationHost,
    PaneHost,
    TerminalHost,
    TerminalPlugin,
};

enum class CommandDispatchPath : uint8_t
{
    Legacy,
    WmCommand,
    StableCommandId,
    TerminalActions,
};

enum class CommandStateSource : uint8_t
{
    LegacyHost,
    ApplicationHost,
    PaneHost,
    TerminalActions,
};

enum class CommandShortcutScope : uint8_t
{
    Application = 1u << 0u,
    FunctionBar = 1u << 1u,
    FolderView  = 1u << 2u,
    Terminal    = 1u << 3u,
};

enum class CommandVisualId : uint8_t
{
    None,
    Close,
    FullScreen,
    Connect,
    TerminalMenu,
    Settings,
    SettingsFile,
    Search,
    CommandPalette,
    Window,
    History,
    NewTab,
    NewWindow,
    NextTab,
    PreviousTab,
    Tab1,
    Tab2,
    Tab3,
    Tab4,
    Tab5,
    Tab6,
    Tab7,
    Tab8,
    LastTab,
    CloseTab,
    ResizeLeft,
    ResizeRight,
    FocusLeft,
    FocusRight,
    SwitchPane,
    Copy,
    CopyOrInput,
    Paste,
    SelectAll,
    ContextMenu,
    ScrollLineDown,
    ScrollPageDown,
    ScrollLineUp,
    ScrollPageUp,
    ScrollTop,
    ScrollBottom,
    ZoomIn,
    ZoomOut,
    ZoomReset,
};

[[nodiscard]] constexpr uint8_t CommandShortcutScopeMask(CommandShortcutScope scope) noexcept
{
    return static_cast<uint8_t>(scope);
}

inline constexpr uint8_t kLegacyShortcutScopeMask = static_cast<uint8_t>(
    CommandShortcutScopeMask(CommandShortcutScope::FunctionBar) | CommandShortcutScopeMask(CommandShortcutScope::FolderView));

struct CommandInfo
{
    std::wstring_view id;
    unsigned int displayNameStringId = 0;
    unsigned int descriptionStringId = 0;
    unsigned int wmCommandId         = 0;
    CommandExecutorOwner executorOwner = CommandExecutorOwner::LegacyHost;
    CommandDispatchPath dispatchPath   = CommandDispatchPath::Legacy;
    CommandStateSource stateSource     = CommandStateSource::LegacyHost;
    uint8_t shortcutScopeMask          = kLegacyShortcutScopeMask;
    bool paletteVisible                = true;
    CommandVisualId visualId           = CommandVisualId::None;
    std::wstring_view searchKeywords;
};

[[nodiscard]] constexpr bool IsCommandShortcutEligible(const CommandInfo& command, CommandShortcutScope scope) noexcept
{
    return (command.shortcutScopeMask & CommandShortcutScopeMask(scope)) != 0u;
}

[[nodiscard]] std::wstring_view CanonicalizeCommandId(std::wstring_view commandId) noexcept;

[[nodiscard]] const CommandInfo* FindCommandInfo(std::wstring_view commandId) noexcept;

[[nodiscard]] const CommandInfo* FindCommandInfoByWmCommandId(unsigned int wmCommandId) noexcept;

[[nodiscard]] std::span<const CommandInfo> GetAllCommands() noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetWmCommandId(std::wstring_view commandId) noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetCommandDisplayNameStringId(std::wstring_view commandId) noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetCommandShortDisplayNameStringId(std::wstring_view commandId) noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetCommandDescriptionStringId(std::wstring_view commandId) noexcept;

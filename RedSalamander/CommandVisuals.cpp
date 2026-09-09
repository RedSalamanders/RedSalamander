#include "Framework.h"

#include "CommandVisuals.h"

#include "FluentIcons.h"

std::wstring ResolveCommandVisualText(CommandVisualId visualId, bool fluentIconFontAvailable)
{
    const auto icon = [fluentIconFontAvailable](wchar_t fluent, wchar_t fallback) { return std::wstring(1u, fluentIconFontAvailable ? fluent : fallback); };
    switch (visualId)
    {
        case CommandVisualId::Settings: return icon(FluentIcons::kSettings, L'⚙');
        case CommandVisualId::SettingsFile: return icon(FluentIcons::kDocument, L'▤');
        case CommandVisualId::Search: return icon(FluentIcons::kFind, L'⌕');
        case CommandVisualId::CommandPalette:
        case CommandVisualId::TerminalMenu: return icon(FluentIcons::kCommandPrompt, L'>');
        case CommandVisualId::Close:
        case CommandVisualId::CloseTab: return icon(FluentIcons::kClear, L'×');
        case CommandVisualId::FullScreen: return icon(FluentIcons::kPreview, L'□');
        case CommandVisualId::Connect: return icon(FluentIcons::kConnections, L'↔');
        case CommandVisualId::Window:
        case CommandVisualId::NewTab:
        case CommandVisualId::NewWindow: return icon(FluentIcons::kOpenFile, L'+');
        case CommandVisualId::History:
        case CommandVisualId::LastTab: return icon(FluentIcons::kHistory, L'↶');
        case CommandVisualId::NextTab: return icon(FluentIcons::kChevronRight, FluentIcons::kFallbackChevronRight);
        case CommandVisualId::PreviousTab: return icon(FluentIcons::kSyncFolder, L'←');
        case CommandVisualId::Tab1:
        case CommandVisualId::Tab2:
        case CommandVisualId::Tab3:
        case CommandVisualId::Tab4:
        case CommandVisualId::Tab5:
        case CommandVisualId::Tab6:
        case CommandVisualId::Tab7:
        case CommandVisualId::Tab8: return icon(FluentIcons::kDocument, L'□');
        case CommandVisualId::ResizeLeft:
        case CommandVisualId::FocusLeft: return icon(FluentIcons::kSyncFolder, L'←');
        case CommandVisualId::ResizeRight:
        case CommandVisualId::FocusRight: return icon(FluentIcons::kSyncFolder, L'→');
        case CommandVisualId::SwitchPane: return icon(FluentIcons::kSyncFolder, L'⇄');
        case CommandVisualId::Copy:
        case CommandVisualId::CopyOrInput: return icon(FluentIcons::kCopy, L'□');
        case CommandVisualId::Paste: return icon(FluentIcons::kPaste, L'▣');
        case CommandVisualId::SelectAll: return icon(FluentIcons::kCheckMark, FluentIcons::kFallbackCheckMark);
        case CommandVisualId::ContextMenu: return icon(FluentIcons::kBulletedList, FluentIcons::kFallbackBulletedList);
        case CommandVisualId::ScrollLineUp:
        case CommandVisualId::ScrollPageUp:
        case CommandVisualId::ScrollTop: return icon(FluentIcons::kChevronUp, FluentIcons::kFallbackChevronUp);
        case CommandVisualId::ScrollLineDown:
        case CommandVisualId::ScrollPageDown:
        case CommandVisualId::ScrollBottom: return icon(FluentIcons::kChevronDown, FluentIcons::kFallbackChevronDown);
        case CommandVisualId::ZoomIn:
        case CommandVisualId::ZoomOut:
        case CommandVisualId::ZoomReset: return icon(FluentIcons::kFont, L'A');
        case CommandVisualId::None: return icon(FluentIcons::kInfo, L'·');
    }
    return icon(FluentIcons::kInfo, L'·');
}

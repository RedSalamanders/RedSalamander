#pragma once

#include "AlertOverlay.h"
#include "AppTheme.h"

namespace RedSalamander::Ui
{
// The alert palette of an application theme: an alert takes the colors of the pane (FolderView) theme
// it appears over. Shared by the HostServices prompts and alerts and by the File Operations decision
// mirror (I35 R8.4), so a decision looks the same in a host prompt and over a pane.
[[nodiscard]] inline AlertTheme MakeAlertTheme(const AppTheme& theme) noexcept
{
    const FolderViewTheme& fv = theme.folderView;

    AlertTheme alertTheme{};
    alertTheme.background          = fv.backgroundColor;
    alertTheme.text                = fv.textNormal;
    alertTheme.accent              = fv.focusBorder;
    alertTheme.selectionBackground = fv.itemBackgroundSelected;
    alertTheme.selectionText       = fv.textSelected;
    alertTheme.errorBackground     = fv.errorBackground;
    alertTheme.errorText           = fv.errorText;
    alertTheme.warningBackground   = fv.warningBackground;
    alertTheme.warningText         = fv.warningText;
    alertTheme.infoBackground      = fv.infoBackground;
    alertTheme.infoText            = fv.infoText;
    alertTheme.darkBase            = fv.darkBase;
    alertTheme.highContrast        = theme.highContrast;
    return alertTheme;
}
} // namespace RedSalamander::Ui

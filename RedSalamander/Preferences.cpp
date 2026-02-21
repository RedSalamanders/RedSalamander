#include "Preferences.h"

#include "Preferences.Dialog.h"
#include "Preferences.Internal.h"

bool ShowPreferencesDialog(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::General);
}

bool ShowPreferencesDialogPlugins(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::Plugins);
}

bool ShowPreferencesDialogHotPaths(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::HotPaths);
}

HWND GetPreferencesDialogHandle() noexcept
{
    return PreferencesDialog::GetHandle();
}

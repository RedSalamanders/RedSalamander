#include "ViewerText.h"

#include "ViewerText.ThemeHelpers.h"

void ViewerText::ApplyMenuTheme(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _menuBarHost.SetTheme(_hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme) : RedSalamander::DxUi::MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else
    {
        Debug::Error(L"ApplyMenuTheme: DxUi menu bar host is missing");
    }
}

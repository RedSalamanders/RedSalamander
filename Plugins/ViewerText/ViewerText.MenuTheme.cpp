#include "ViewerText.h"
#include "ViewerDxUiTheme.h"

#include "ViewerText.ThemeHelpers.h"

void ViewerText::ApplyMenuTheme(HWND hwnd) noexcept
{
    _menuBarHost.SetTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd != nullptr && GetMenu(hwnd) != nullptr)
    {
        // Theme application can run during WM_CREATE, before Open attaches the DxUi host.
        // Preserve the native menu bar during that normal interval; embedded windows have no menu.
        DrawMenuBar(hwnd);
    }
}

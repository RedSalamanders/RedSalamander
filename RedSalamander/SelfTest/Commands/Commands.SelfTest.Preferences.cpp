// Commands.SelfTest.Preferences.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Preferences test family: 129 test functions.

namespace
{

void SendScaledHeaderResizeDrag(HWND activePage, const RECT& headerRect) noexcept
{
    const int dpi           = std::max(static_cast<int>(GetDpiForWindow(activePage)), USER_DEFAULT_SCREEN_DPI);
    const LONG gripInset    = 1;
    const LONG dragDistance = std::max<LONG>(16, MulDiv(48, dpi, USER_DEFAULT_SCREEN_DPI));

    LONG startX              = headerRect.right - gripInset;
    const LONG minimumStartX = headerRect.left + 1;
    if (startX < minimumStartX)
    {
        startX = minimumStartX;
    }

    const LONG dragY = headerRect.top + ((headerRect.bottom - headerRect.top) / 2);
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(startX, dragY), MAKELPARAM(startX + dragDistance, dragY));
}

} // namespace

#include "Commands.SelfTest.Preferences.ChromeAndPlugins.cpp"
#include "Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp"
#include "Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp"
#include "Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp"
#include "Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp"
#include "Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp"
#include "Commands.SelfTest.Preferences.Dispatch.cpp"

namespace
{

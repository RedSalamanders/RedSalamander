// Preferences.Mouse.cpp

#include "Framework.h"

#include "Preferences.Mouse.h"

void MousePane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void MousePane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
}

void MousePane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    static_cast<void>(parent);
    static_cast<void>(state);
}

void MousePane::LayoutPage(HWND host,
                           PreferencesDialogState& state,
                           int /*x*/,
                           int& y,
                           int width,
                           int /*margin*/,
                           int /*gapY*/,
                           int sectionY,
                           const PreferencesTypographyContext& typography) noexcept
{
    static_cast<void>(host);
    static_cast<void>(state);
    static_cast<void>(y);
    static_cast<void>(width);
    static_cast<void>(sectionY);
    static_cast<void>(typography);
}

// Preferences.Mouse.cpp

#include "Framework.h"

#include <algorithm>

#include "Preferences.Mouse.h"

#include "Helpers.h"

#include "resource.h"

bool MousePane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void MousePane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void MousePane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void MousePane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD wrapStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    state.mouseNote.reset(CreateWindowExW(0,
                                          L"Static",
                                          LoadStringResource(nullptr, IDS_PREFS_MOUSE_PLACEHOLDER).c_str(),
                                          wrapStyle,
                                          0,
                                          0,
                                          10,
                                          10,
                                          parent,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          nullptr));
}

void MousePane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int /*margin*/, int /*gapY*/, int sectionY, HFONT dialogFont) noexcept
{
    if (! host)
    {
        return;
    }

    if (state.mouseNote)
    {
        const HFONT infoFont        = state.italicFont ? state.italicFont.get() : dialogFont;
        const std::wstring noteText = PrefsUi::GetWindowTextString(state.mouseNote.get());
        const int noteHeight        = noteText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, infoFont, width, noteText);
        SetWindowPos(state.mouseNote.get(), nullptr, x, y, width, std::max(0, noteHeight), SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.mouseNote.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        y += std::max(0, noteHeight) + sectionY;
    }
}

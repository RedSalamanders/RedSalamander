#pragma once

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"

class MousePane final
{
public:
    MousePane()                            = default;
    MousePane(const MousePane&)            = delete;
    MousePane& operator=(const MousePane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void LayoutPage(HWND host,
                    PreferencesDialogState& state,
                    int x,
                    int& y,
                    int width,
                    int margin,
                    int gapY,
                    int sectionY,
                    const PreferencesTypographyContext& typography) noexcept;

private:
};

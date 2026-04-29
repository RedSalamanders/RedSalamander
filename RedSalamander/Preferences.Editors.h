#pragma once

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"

class EditorsPane final
{
public:
    EditorsPane()                              = default;
    EditorsPane(const EditorsPane&)            = delete;
    EditorsPane& operator=(const EditorsPane&) = delete;

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

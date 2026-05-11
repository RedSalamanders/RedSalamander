#pragma once

#include "Preferences.FileActions.h"

class EditorsPane final
{
public:
    EditorsPane() noexcept;
    EditorsPane(const EditorsPane&)            = delete;
    EditorsPane& operator=(const EditorsPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;
    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(HWND host,
                    PreferencesDialogState& state,
                    int x,
                    int& y,
                    int width,
                    int margin,
                    int gapY,
                    int sectionY,
                    const PreferencesTypographyContext& typography) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugSelectDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
    [[nodiscard]] bool DebugSelectDefaultEditNewAction(std::wstring_view actionId) noexcept;
    [[nodiscard]] size_t DebugAssociationRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugAssociationVisibleWorkMetrics() const noexcept;
    [[nodiscard]] size_t DebugActionRowCount() const noexcept;
    [[nodiscard]] std::wstring DebugPreviewActionId() const;
    [[nodiscard]] std::wstring DebugPreviewReason() const;
#endif

private:
    FileActionPreferencesPage _page;
};

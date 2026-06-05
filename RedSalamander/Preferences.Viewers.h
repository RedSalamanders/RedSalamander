#pragma once

#include "Preferences.FileActions.h"

class ViewersPane final
{
public:
    ViewersPane() noexcept;
    ViewersPane(const ViewersPane&)            = delete;
    ViewersPane& operator=(const ViewersPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;
    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] size_t DebugListRowCount() const noexcept;
    [[nodiscard]] size_t DebugActionRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeFailureCount() const noexcept;
    [[nodiscard]] PreferencesViewersDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyContext() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyMetrics() const noexcept;
    [[nodiscard]] bool DebugGetListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
    [[nodiscard]] bool DebugGetListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
    [[nodiscard]] bool DebugHitTestListClientPoint(
        POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept;
    [[nodiscard]] bool DebugGetListPointerState(PreferencesGridPointerDebugState& outState) const noexcept;
    [[nodiscard]] bool DebugGetTabClientRect(size_t tabIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetSelectedTabIndex(size_t& outIndex) const noexcept;
    [[nodiscard]] bool DebugSelectListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugSelectDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
    [[nodiscard]] std::wstring DebugPreviewActionId() const;
    [[nodiscard]] std::wstring DebugPreviewReason() const;
#endif

private:
    FileActionPreferencesPage _page;
};

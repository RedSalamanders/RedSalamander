#include "Framework.h"

#include "Preferences.Viewers.h"

ViewersPane::ViewersPane() noexcept : _page(FileActionPreferencesFamily::Viewers)
{
}

void ViewersPane::OnVisibilityChanged(const bool visible) noexcept
{
    _page.OnVisibilityChanged(visible);
}

void ViewersPane::Destroy(PreferencesDialogState& state) noexcept
{
    _page.Destroy(state);
}

void ViewersPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    _page.InitializePage(parent, state);
}

void ViewersPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    _page.Refresh(host, state);
}

void ViewersPane::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    _page.LayoutPage(host, state, x, y, width, margin, gapY, typography);
}

bool ViewersPane::HandleDeferredAction(HWND host, PreferencesDialogState& state, const PreferencesDeferredActionKind action) noexcept
{
    return _page.HandleDeferredAction(host, state, action);
}

#ifdef ENABLE_TESTS
size_t ViewersPane::DebugListRowCount() const noexcept
{
    return _page.DebugAssociationRowCount();
}

size_t ViewersPane::DebugActionRowCount() const noexcept
{
    return _page.DebugActionRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics ViewersPane::DebugListVisibleWorkMetrics() const noexcept
{
    return _page.DebugAssociationVisibleWorkMetrics();
}

uint64_t ViewersPane::DebugListRenderCount() const noexcept
{
    return _page.DebugAssociationRenderCount();
}

uint64_t ViewersPane::DebugListResizeCount() const noexcept
{
    return _page.DebugAssociationResizeCount();
}

uint64_t ViewersPane::DebugListResizeFailureCount() const noexcept
{
    return _page.DebugAssociationResizeFailureCount();
}

PreferencesViewersDebugFocusTarget ViewersPane::DebugGetFocusTarget() const noexcept
{
    return _page.DebugGetViewersFocusTarget();
}

bool ViewersPane::DebugUsesDxUiTypographyContext() const noexcept
{
    return _page.DebugUsesDxUiTypographyContext();
}

bool ViewersPane::DebugUsesDxUiTypographyMetrics() const noexcept
{
    return _page.DebugUsesDxUiTypographyMetrics();
}

bool ViewersPane::DebugGetListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    return _page.DebugGetAssociationRowClientRect(rowIndex, outRect);
}

bool ViewersPane::DebugGetListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return _page.DebugGetAssociationHeaderClientRect(columnIndex, outRect);
}

bool ViewersPane::DebugHitTestListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept
{
    return _page.DebugHitTestAssociationClientPoint(clientPoint, outZone, outColumnIndex, outHeaderResize, outHostHitsList);
}

bool ViewersPane::DebugGetListPointerState(PreferencesGridPointerDebugState& outState) const noexcept
{
    return _page.DebugGetAssociationPointerState(outState);
}

bool ViewersPane::DebugGetTabClientRect(const size_t tabIndex, RECT& outRect) const noexcept
{
    return _page.DebugGetTabClientRect(tabIndex, outRect);
}

bool ViewersPane::DebugGetSelectedTabIndex(size_t& outIndex) const noexcept
{
    return _page.DebugGetSelectedTabIndex(outIndex);
}

bool ViewersPane::DebugSelectListRow(const size_t rowIndex) noexcept
{
    return _page.DebugSelectAssociationRow(rowIndex);
}

bool ViewersPane::DebugSetSearchText(std::wstring_view text) noexcept
{
    return _page.DebugSetSearchText(text);
}

bool ViewersPane::DebugSelectDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    return _page.DebugSelectDefaultAction(alternate, actionId);
}

bool ViewersPane::DebugFocusSearchField() noexcept
{
    return _page.DebugFocusSearchField();
}

bool ViewersPane::DebugScrollListByWheelDetents(const int detents) noexcept
{
    return _page.DebugScrollAssociationByWheelDetents(detents);
}

std::wstring ViewersPane::DebugPreviewActionId() const
{
    return _page.DebugPreviewActionId();
}

std::wstring ViewersPane::DebugPreviewReason() const
{
    return _page.DebugPreviewReason();
}
#endif

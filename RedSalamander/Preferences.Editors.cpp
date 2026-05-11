#include "Framework.h"

#include "Preferences.Editors.h"

EditorsPane::EditorsPane() noexcept : _page(FileActionPreferencesFamily::Editors)
{
}

void EditorsPane::OnVisibilityChanged(const bool visible) noexcept
{
    _page.OnVisibilityChanged(visible);
}

void EditorsPane::Destroy(PreferencesDialogState& state) noexcept
{
    _page.Destroy(state);
}

void EditorsPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    _page.InitializePage(parent, state);
}

void EditorsPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    _page.Refresh(host, state);
}

void EditorsPane::LayoutPage(HWND host,
                             PreferencesDialogState& state,
                             int x,
                             int& y,
                             int width,
                             int margin,
                             int gapY,
                             int sectionY,
                             const PreferencesTypographyContext& typography) noexcept
{
    static_cast<void>(sectionY);
    _page.LayoutPage(host, state, x, y, width, margin, gapY, typography);
}

#ifdef ENABLE_TESTS
bool EditorsPane::DebugSelectDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    return _page.DebugSelectDefaultAction(alternate, actionId);
}

bool EditorsPane::DebugSelectDefaultEditNewAction(std::wstring_view actionId) noexcept
{
    return _page.DebugSelectDefaultEditNewAction(actionId);
}

size_t EditorsPane::DebugAssociationRowCount() const noexcept
{
    return _page.DebugAssociationRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics EditorsPane::DebugAssociationVisibleWorkMetrics() const noexcept
{
    return _page.DebugAssociationVisibleWorkMetrics();
}

size_t EditorsPane::DebugActionRowCount() const noexcept
{
    return _page.DebugActionRowCount();
}

std::wstring EditorsPane::DebugPreviewActionId() const
{
    return _page.DebugPreviewActionId();
}

std::wstring EditorsPane::DebugPreviewReason() const
{
    return _page.DebugPreviewReason();
}
#endif

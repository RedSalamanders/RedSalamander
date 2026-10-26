#include "ViewerSpace.AnalysisHost.h"
#include <windowsx.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cmath>
#include <exception>
#include <format>
#include <new>
#include <utility>

#include <DxUi/Typography.h>

#include "Helpers.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace ViewerSpaceAnalysis
{
namespace
{
// Root decoration shares the grid's existing controls and UIA tree. It is paint-only;
// there is no extra hit target, native window, provider access or offscreen-row materialization.
class WorkspacePanel final : public DxUi::Panel
{
public:
    explicit WorkspacePanel(std::function<void(DxUi::ControlHost&)> paint) : _paint(std::move(paint))
    {
    }
    WorkspacePanel(const WorkspacePanel&)            = delete;
    WorkspacePanel(WorkspacePanel&&)                 = delete;
    WorkspacePanel& operator=(const WorkspacePanel&) = delete;
    WorkspacePanel& operator=(WorkspacePanel&&)      = delete;
    void PaintOverlay(DxUi::ControlHost& host) const override
    {
        DxUi::Panel::PaintOverlay(host);
        if (_paint)
            _paint(host);
    }

private:
    std::function<void(DxUi::ControlHost&)> _paint;
};
constexpr float kMarginDip                  = 12.0f;
constexpr float kGapDip                     = 8.0f;
constexpr float kFilterHeightDip            = 34.0f;
constexpr float kSideBySideMinimumWidthDip  = 850.0f;
constexpr float kMinimumFilterFieldWidthDip = 120.0f;

[[nodiscard]] D2D1_RECT_F Rect(float left, float top, float right, float bottom) noexcept
{
    return D2D1::RectF(left, top, right, bottom);
}

[[nodiscard]] float MeasureTextWidthDip(const DxUi::WindowHost& host, DxUi::FontRole role, std::wstring_view text, float fallbackDip) noexcept
{
    const int widthPx = DxUi::Typography::MeasureSingleLineTextWidthPx(host.GetHwnd(), role, text);
    return widthPx > 0 ? host.PixelsToDip(static_cast<float>(widthPx)) : fallbackDip;
}

[[nodiscard]] std::wstring BuildInspectorVisualText(const DxUi::WindowHost& host, std::wstring_view fullText, float availableWidthDip, bool middleEllipsis)
{
    constexpr float kTextInsetDip = 8.0f;
    const float maxTextWidthDip   = std::max(0.0f, availableWidthDip - kTextInsetDip);
    if (fullText.empty() || MeasureTextWidthDip(host, DxUi::FontRole::Body, fullText, 0.0f) <= maxTextWidthDip)
    {
        return std::wstring(fullText);
    }

    constexpr std::wstring_view kEllipsis = L"…";
    if (MeasureTextWidthDip(host, DxUi::FontRole::Body, kEllipsis, 0.0f) > maxTextWidthDip)
    {
        return {};
    }

    // This stays local to the analysis host because DxUi::Label has no trimming API; only the
    // visible string is shortened, while UIA exposes the full inspector value.
    const size_t unitCount = fullText.size();
    size_t low             = 0u;
    size_t high            = unitCount - 1u;
    std::wstring best(kEllipsis);
    while (low <= high)
    {
        const size_t kept        = low + ((high - low) / 2u);
        size_t prefixUnits       = middleEllipsis ? (kept + 1u) / 2u : kept;
        const size_t suffixUnits = middleEllipsis ? kept / 2u : 0u;
        size_t suffixStart       = unitCount - suffixUnits;
        // Never divide a valid UTF-16 surrogate pair at either visible edge.
        if (prefixUnits > 0u && prefixUnits < unitCount && fullText[prefixUnits - 1u] >= 0xD800 && fullText[prefixUnits - 1u] <= 0xDBFF &&
            fullText[prefixUnits] >= 0xDC00 && fullText[prefixUnits] <= 0xDFFF)
        {
            --prefixUnits;
        }
        if (suffixStart > 0u && suffixStart < unitCount && fullText[suffixStart - 1u] >= 0xD800 && fullText[suffixStart - 1u] <= 0xDBFF &&
            fullText[suffixStart] >= 0xDC00 && fullText[suffixStart] <= 0xDFFF)
        {
            ++suffixStart;
        }
        std::wstring candidate(fullText.substr(0u, prefixUnits));
        candidate.append(kEllipsis);
        candidate.append(fullText.substr(suffixStart));
        if (MeasureTextWidthDip(host, DxUi::FontRole::Body, candidate, 0.0f) <= maxTextWidthDip)
        {
            best = std::move(candidate);
            low  = kept + 1u;
        }
        else if (kept == 0u)
        {
            break;
        }
        else
        {
            high = kept - 1u;
        }
    }
    return best;
}

[[nodiscard]] float MeasureWrappedTextHeightDip(const DxUi::WindowHost& host, DxUi::FontRole role, float widthDip, std::wstring_view text) noexcept
{
    if (widthDip <= 0.0f || text.empty())
    {
        return 0.0f;
    }

    const int widthPx  = std::max(1, static_cast<int>(host.DipsToPixels(widthDip)));
    const int heightPx = DxUi::Typography::MeasureWrappedTextHeightPx(host.GetHwnd(), role, widthPx, text);
    return (heightPx > 0) ? host.PixelsToDip(static_cast<float>(heightPx)) : 18.0f;
}

[[nodiscard]] const wchar_t* ColumnId(size_t columnIndex) noexcept
{
    constexpr std::array<std::wstring_view, 5u> ids{{L"name", L"logical-size", L"parent-share", L"cleanup", L"state"}};
    return (columnIndex < ids.size()) ? ids[columnIndex].data() : L"";
}

} // namespace

AnalysisHost::~AnalysisHost() noexcept
{
    Detach();
}

AnalysisHost::AnalysisHost() noexcept
{
    _model.Bind(*this);
    _delegate.Bind(*this);
}

HRESULT AnalysisHost::Attach(HWND parent, const Captions& captions, Callbacks callbacks) noexcept
{
    if (parent == nullptr || ! IsWindow(parent))
    {
        return E_INVALIDARG;
    }
    if (_hwnd || _host.GetHwnd() != nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
    }

    try
    {
        _captions  = captions;
        _callbacks = std::move(callbacks);
        _root      = std::make_unique<WorkspacePanel>([this](DxUi::ControlHost& host) { PaintExclusionOverlay(host); });
        BuildControls();

        if (! RegisterChildClass())
        {
            const DWORD error = GetLastError();
            Detach();
            return HRESULT_FROM_WIN32((error == ERROR_SUCCESS) ? ERROR_CANNOT_MAKE : error);
        }

        _hwnd.reset(
            CreateWindowExW(0u, _windowClassName.c_str(), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 0, 0, parent, nullptr, g_hInstance, this));
        if (! _hwnd)
        {
            const DWORD error = GetLastError();
            Detach();
            return HRESULT_FROM_WIN32((error == ERROR_SUCCESS) ? ERROR_CANNOT_MAKE : error);
        }

        if (! _host.Attach(_hwnd.get()))
        {
            Detach();
            return E_FAIL;
        }

        _grid->SetModel(&_model);
        _grid->SetDelegate(&_delegate);
        _host.SetRoot(std::move(_root));
        _host.SetOnEscape([this]() noexcept
        {
            if (! _callbacks.panelClose)
            {
                return false;
            }
            InvokePanelClose();
            return true;
        });
        _host.SetOnTabBoundary([this](bool reverse) noexcept
        {
            bool handled = false;
            InvokeFocusBoundary(reverse, handled);
            return handled;
        });

        SetBounds(_boundsPx);
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"ViewerSpace analysis host setup failed.");
        Detach();
        return E_FAIL;
    }
}

void AnalysisHost::Detach() noexcept
{
    _gridLayoutWidthDip = -1.0f;
    _callbacks          = {};
    if (_panelSplitter != nullptr)
        _panelSplitter->SetOnChange({});
    if (_inspectorSplitter != nullptr)
        _inspectorSplitter->SetOnChange({});
    if (_inspectorToggle != nullptr)
        _inspectorToggle->SetOnClick({});
    if (_grid != nullptr)
    {
        _grid->SetModel(nullptr);
        _grid->SetDelegate(nullptr);
    }
    if (_filterField != nullptr)
    {
        _filterField->SetOnTextChanged({});
    }

    _host.SetOnEscape({});
    _host.SetOnTabBoundary({});
    _host.SetRoot(std::unique_ptr<DxUi::Control>{});
    _host.Detach();
    _hwnd.reset();
    if (_windowClassAtom != 0 && ! _windowClassName.empty())
    {
        static_cast<void>(UnregisterClassW(_windowClassName.c_str(), g_hInstance));
    }

    _windowClassAtom = 0;
    _windowClassName.clear();
    _root.reset();
    _filterLabel       = nullptr;
    _filterField       = nullptr;
    _scopeLabel        = nullptr;
    _grid              = nullptr;
    _reviewCleanup     = nullptr;
    _recycle           = nullptr;
    _permanentDelete   = nullptr;
    _inspectorTitle    = nullptr;
    _inspectorToggle   = nullptr;
    _inspectorSplitter = nullptr;
    _panelSplitter     = nullptr;
    _legendSwatches.fill(nullptr);
    _legendLabels.fill(nullptr);
    _legendCount = 0u;
    _inspectorFieldLabels.fill(nullptr);
    _inspectorValues.fill(nullptr);
    _inspectorText = {};
    _inspectorResizeStartRatio.reset();
    _callbackFailureLogged = false;
}

void AnalysisHost::SetBounds(const RECT& boundsPx) noexcept
{
    _boundsPx = boundsPx;
    if (_hwnd)
    {
        const int width  = std::max(0L, boundsPx.right - boundsPx.left);
        const int height = std::max(0L, boundsPx.bottom - boundsPx.top);
        static_cast<void>(SetWindowPos(_hwnd.get(), nullptr, boundsPx.left, boundsPx.top, width, height, SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOOWNERZORDER));
    }
}

void AnalysisHost::ApplyTheme(const DxUi::ThemePalette& palette) noexcept
{
    DxUi::ThemePalette applied = palette;
    applied.density            = _compactRows ? DxUi::Density::Compact : DxUi::Density::Standard;
    // Use the theme's subtle divider tone for workspace separators. Canonical
    // Splitter hover, drag and keyboard-focus accents remain intact; HC stays exact.
    if (! palette.highContrast)
        applied.border = palette.gridLine;
    _host.SetTheme(applied);
}

void AnalysisHost::SetCompactRows(bool compactRows) noexcept
{
    _compactRows = compactRows;
}

void AnalysisHost::NotifyRowsChanged(std::optional<uint64_t> selectedRowId)
{
    if (_grid == nullptr)
    {
        return;
    }

    auto& selection = _grid->GetSelectionModel();
    if (selectedRowId.has_value() && _model.FindRowByStableId(selectedRowId.value()).has_value())
    {
        selection.SetSingle(selectedRowId.value());
        if (const auto row = _model.FindRowByStableId(selectedRowId.value()); row.has_value())
        {
            _grid->EnsureRowVisible(row.value());
        }
    }
    else
    {
        selection.Clear();
    }

    // NotifyDataChanged publishes the same host snapshot through the Grid control.
    _grid->NotifyDataChanged();
    _host.Invalidate();
}

void AnalysisHost::SetInspector(const InspectorText& text)
{
    _inspectorText = text;
    const std::array<std::wstring_view, 7u> values{_inspectorText.path,
                                                   _inspectorText.name,
                                                   _inspectorText.icon,
                                                   _inspectorText.logicalBytes,
                                                   _inspectorText.share,
                                                   _inspectorText.modified,
                                                   _inspectorText.state};
    bool changed = false;
    for (size_t index = 0u; index < values.size(); ++index)
    {
        const auto accessible = index == 0u ? std::wstring_view(_inspectorText.fullPath) : index == 2u ? std::wstring_view(_inspectorText.kind) : values[index];
        const std::wstring_view detail = index == 3u ? std::wstring_view(_inspectorText.exactBytes) : accessible;
        if (_inspectorValues[index] != nullptr &&
            (_inspectorValues[index]->GetAccessibleName() != accessible || _inspectorValues[index]->GetAccessibleHelpText() != detail))
        {
            // UIA keeps the unabridged value even when the visible label needs an ellipsis.
            _inspectorValues[index]->SetAccessibleName(std::wstring(accessible));
            _inspectorValues[index]->SetAccessibleHelpText(std::wstring(detail));
            _inspectorValues[index]->SetTooltipText(std::wstring(detail));
            changed = true;
        }
        if (index >= 2u && _inspectorValues[index] != nullptr && _inspectorValues[index]->GetText() != values[index])
        {
            _inspectorValues[index]->SetText(std::wstring(values[index]));
            changed = true;
        }
    }
    if (changed)
    {
        Layout();
        _host.RefreshAccessibilitySnapshot();
        _host.Invalidate();
    }
}

void AnalysisHost::SetPanelSplitter(bool vertical, float extentDip, float positionDip) noexcept
{
    _panelVertical = vertical;
    _panelExtent   = extentDip;
    _panelPosition = positionDip;
    Layout();
}

void AnalysisHost::CancelResize() noexcept
{
    auto* captured           = _host.GetCapturedControl();
    DxUi::Splitter* splitter = captured == _panelSplitter ? _panelSplitter : captured == _inspectorSplitter ? _inspectorSplitter : nullptr;
    if (splitter != nullptr)
    {
        // Explicit release clears the host's captured control before WM_CAPTURECHANGED.
        // Deliver the canonical cancellation first so preview ratios are restored.
        splitter->OnCaptureLost(_host);
        _host.ReleaseMouseCapture();
    }
}

void AnalysisHost::SetInspectorLayout(bool visible, float ratio) noexcept
{
    _inspectorVisible = visible;
    _inspectorRatio   = ratio;
    Layout();
    _host.RefreshAccessibilitySnapshot();
    _host.Invalidate();
}

void AnalysisHost::SetFilterText(std::wstring_view text)
{
    if (_filterField != nullptr && _filterField->GetText() != text)
    {
        // SetText intentionally avoids SetOnTextChanged; callers own the query lifecycle.
        _filterField->SetText(std::wstring(text));
    }
}

void AnalysisHost::SetSortSpec(const DxUi::GridSortSpec& sortSpec) noexcept
{
    if (_grid != nullptr)
    {
        _grid->SetSortSpec(sortSpec);
    }
}

HWND AnalysisHost::GetHwnd() const noexcept
{
    return _hwnd.get();
}

bool AnalysisHost::OwnsFocus() const noexcept
{
    if (! _hwnd)
    {
        return false;
    }
    const HWND focus = GetFocus();
    return focus == _hwnd.get() || (focus != nullptr && IsChild(_hwnd.get(), focus) != FALSE);
}

bool AnalysisHost::Focus(bool reverse) noexcept
{
    if (_host.GetHwnd() == nullptr)
    {
        return false;
    }
    DxUi::Control* target = reverse ? static_cast<DxUi::Control*>(_inspectorToggle) : static_cast<DxUi::Control*>(_filterField);
    if (target == nullptr)
    {
        return false;
    }
    _host.SetFocusControl(target);
    return OwnsFocus();
}

void AnalysisHost::Model::Bind(AnalysisHost& owner) noexcept
{
    _owner = &owner;
}

size_t AnalysisHost::Model::GetRowCount() const noexcept
{
    if (! _owner->_callbacks.rowCount)
    {
        return 0u;
    }
    try
    {
        return _owner->_callbacks.rowCount();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
        return 0u;
    }
}

size_t AnalysisHost::Model::GetColumnCount() const noexcept
{
    return 4u;
}

DxUi::GridColumnDesc AnalysisHost::Model::GetColumn(size_t columnIndex) const
{
    DxUi::GridColumnDesc column;
    if (columnIndex >= 4u)
    {
        return column;
    }

    column.id             = ColumnId(columnIndex);
    column.title          = _owner->_captions.columns[columnIndex];
    const float available = std::max(240.0f, _owner->_grid->GetBounds().right - _owner->_grid->GetBounds().left - 18.0f);
    const std::array<float, 4u> widths{{std::max(120.0f, available - 264.0f), 94.0f, 90.0f, 80.0f}};
    column.widthDip    = widths[columnIndex];
    column.minWidthDip = columnIndex == 0u ? 90.0f : columnIndex == 1u ? 80.0f : 54.0f;
    column.sortable    = true;
    column.multiline   = false;
    if (columnIndex == 3u)
    {
        column.kind        = DxUi::GridColumnKind::Checkbox;
        column.minWidthDip = 32.0f;
        column.sortable    = false;
    }
    if (columnIndex == 1u || columnIndex == 2u)
    {
        column.textAlignment = DWRITE_TEXT_ALIGNMENT_TRAILING;
    }
    return column;
}

void AnalysisHost::Model::GetCellData(size_t rowIndex, size_t columnIndex, DxUi::GridCellData& outCell) const
{
    if (! _owner->_callbacks.cellText)
    {
        return;
    }
    try
    {
        _owner->_callbacks.cellText(rowIndex, columnIndex, outCell);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
    }
}

uint64_t AnalysisHost::Model::GetStableRowId(size_t rowIndex) const noexcept
{
    if (! _owner->_callbacks.rowId)
    {
        return 0u;
    }
    try
    {
        return _owner->_callbacks.rowId(rowIndex);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
        return 0u;
    }
}

std::optional<size_t> AnalysisHost::Model::FindRowByStableId(uint64_t rowId) const noexcept
{
    if (rowId == 0u || ! _owner->_callbacks.findRow)
    {
        return std::nullopt;
    }
    try
    {
        return _owner->_callbacks.findRow(rowId);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
        return std::nullopt;
    }
}

void AnalysisHost::Delegate::Bind(AnalysisHost& owner) noexcept
{
    _owner = &owner;
}

void AnalysisHost::Delegate::OnGridSortRequested(const DxUi::GridSortSpec& sortSpec)
{
    if (! _owner->_callbacks.sortRequested)
    {
        return;
    }
    try
    {
        _owner->_callbacks.sortRequested(sortSpec);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
    }
}

void AnalysisHost::Delegate::OnGridSelectionChanged(DxUi::Grid& sender)
{
    const auto selected = sender.GetPrimarySelectedRow();
    _owner->InvokeSelectionChanged(selected.has_value() ? _owner->StableRowIdAt(selected.value()) : std::nullopt);
}

void AnalysisHost::Delegate::OnGridRowActivated(DxUi::Grid&, size_t rowIndex)
{
    const auto id = _owner->StableRowIdAt(rowIndex);
    if (! id.has_value() || ! _owner->_callbacks.rowActivated)
    {
        return;
    }
    try
    {
        _owner->_callbacks.rowActivated(id.value());
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
    }
}

void AnalysisHost::Delegate::OnGridContextMenu(DxUi::Grid&, size_t rowIndex, POINT screenPoint)
{
    const auto id = _owner->StableRowIdAt(rowIndex);
    if (! id.has_value() || ! _owner->_callbacks.rowContextMenu)
    {
        return;
    }
    try
    {
        _owner->_callbacks.rowContextMenu(id.value(), screenPoint);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        _owner->LogCallbackFailureOnce();
    }
}

void AnalysisHost::Delegate::OnGridCheckboxToggled(DxUi::Grid&, size_t rowIndex, size_t columnIndex, bool)
{
    const auto row = _owner->StableRowIdAt(rowIndex);
    if (columnIndex != 3u || ! row.has_value() || ! _owner->_callbacks.cleanupToggled)
        return;
    try
    {
        _owner->_callbacks.cleanupToggled(row.value());
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // UI delegate boundary: preserve the host and report callback failure once.
        _owner->LogCallbackFailureOnce();
    }
}

LRESULT CALLBACK AnalysisHost::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        if (create == nullptr || create->lpCreateParams == nullptr)
        {
            return FALSE;
        }
        auto* self = static_cast<AnalysisHost*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        InitPostedPayloadWindow(hwnd);
    }

    auto* self = reinterpret_cast<AnalysisHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self != nullptr)
    {
        return self->HandleWindowMessage(hwnd, message, wParam, lParam);
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT AnalysisHost::HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    const ViewerSpaceHosts::ScopedHostDiagnostics diagnostics;
    if ((message == WM_LBUTTONDOWN || message == WM_LBUTTONDBLCLK) && _grid && _callbacks.exclusionToggled)
    {
        const DxUi::PointDip point{_host.PixelsToDip(static_cast<float>(GET_X_LPARAM(lParam))), _host.PixelsToDip(static_cast<float>(GET_Y_LPARAM(lParam)))};
        const auto cell = _grid->FindCellAtPoint(point);
        if (cell.has_value() && cell.value().second == 0u)
        {
            const auto metrics = _grid->GetCellLayoutMetrics(_host, cell.value().first, 0u);
            const auto& rect   = metrics.iconRect;
            if (metrics.hasIcon && point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom)
            {
                const auto row = StableRowIdAt(cell.value().first);
                if (row.has_value())
                {
                    if (message == WM_LBUTTONDOWN)
                    {
                        NotifyRowsChanged(row);
                        InvokeSelectionChanged(row);
                        InvokeExclusionToggled(row.value());
                    }
                    return 0;
                }
            }
        }
    }
    if ((message == WM_KEYDOWN || message == WM_SYSKEYDOWN) && _callbacks.commandKey)
    {
        const UINT virtualKey = static_cast<UINT>(wParam);
        if (virtualKey == VK_ESCAPE && (_host.GetCapturedControl() == _panelSplitter || _host.GetCapturedControl() == _inspectorSplitter))
        {
            CancelResize();
            return 0;
        }
        const bool alt          = message == WM_SYSKEYDOWN || (GetKeyState(VK_MENU) & 0x8000) != 0;
        const bool control      = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool ownerCommand = virtualKey == VK_F5 || virtualKey == VK_ESCAPE ||
                                  (alt && (virtualKey == VK_LEFT || virtualKey == VK_RIGHT || virtualKey == VK_INSERT)) ||
                                  (control && (virtualKey == '1' || virtualKey == '2'));
        if (virtualKey == VK_ESCAPE && _filterField != nullptr && ! _filterField->GetText().empty())
        {
            _filterField->SetText(L"");
            InvokeFilterChanged(L"");
            return 0;
        }
        if (ownerCommand && InvokeCommandKey(virtualKey, (alt ? MOD_ALT : 0u) | (control ? MOD_CONTROL : 0u)))
        {
            return 0;
        }
    }

    bool handled         = false;
    const LRESULT result = _host.HandleMessage(hwnd, message, wParam, lParam, handled);
    if (message == WM_SIZE || message == WM_DPICHANGED_AFTERPARENT)
    {
        Layout();
    }
    if (message == WM_NCDESTROY)
    {
        static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    }
    return handled ? result : DefWindowProcW(hwnd, message, wParam, lParam);
}

void AnalysisHost::InvokeExclusionToggled(uint64_t rowId) noexcept
{
    try
    {
        if (_callbacks.exclusionToggled)
            _callbacks.exclusionToggled(rowId);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Native input boundary: keep the host alive and report the failed action once.
        LogCallbackFailureOnce();
    }
}

void AnalysisHost::InvokeWorkspaceAction(UINT action) noexcept
{
    try
    {
        if (_callbacks.workspaceAction)
            _callbacks.workspaceAction(action);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Native button boundary: preserve the host and report callback failure once.
        LogCallbackFailureOnce();
    }
}

void AnalysisHost::SetWorkspaceScope(bool reviewing, bool problems, std::wstring_view scope)
{
    if (! _scopeLabel)
        return;
    if (_reviewingCleanup == reviewing && _alternateWorkspace == (reviewing || problems) && _scopeLabel->GetText() == scope)
        return;
    _reviewingCleanup   = reviewing;
    _alternateWorkspace = reviewing || problems;
    _scopeLabel->SetText(std::wstring(scope));
    _scopeLabel->SetTooltipText(std::wstring(scope));
    _filterField->SetEnabled(! reviewing && ! problems);
    _recycle->SetVisible(reviewing);
    _permanentDelete->SetVisible(reviewing);
    _reviewCleanup->SetText(reviewing || problems ? L"\uE72B" : L"\uE7C1");
    _reviewCleanup->SetAccessibleName(_alternateWorkspace ? _captions.back : _captions.reviewCleanup);
    _reviewCleanup->SetTooltipText(_alternateWorkspace ? _captions.back : _captions.reviewCleanup);
    _filterLabel->SetVisible(! _alternateWorkspace);
    _filterField->SetVisible(! _alternateWorkspace);
    Layout();
}

void AnalysisHost::PaintExclusionOverlay(DxUi::ControlHost& host) const
{
    if (! _grid || ! _callbacks.rowExcluded)
        return;
    auto* context = host.GetDeviceContext();
    if (! context)
        return;
    auto dim              = host.GetTheme().windowBackground;
    dim.a                 = host.GetTheme().highContrast ? 0.0f : 0.40f;
    const auto foreground = host.GetTheme().text;
    auto* dimBrush        = host.GetSolidBrush(dim);
    auto* crossBrush      = host.GetSolidBrush(foreground);
    if (! dimBrush || ! crossBrush)
        return;
    const auto bounds = _grid->GetBounds();
    context->PushAxisAlignedClip(bounds, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    const size_t visibleCount = _grid->GetVisibleRowCount();
    for (size_t index = 0u; index < visibleCount; ++index)
    {
        const auto row = _grid->GetVisibleRowAt(index);
        if (! row.has_value() || ! _callbacks.rowExcluded(row.value()))
            continue;
        const auto rect = _grid->GetVisibleRowRect(row.value());
        if (! rect.has_value())
            continue;
        context->FillRectangle(rect.value(), dimBrush);
        const auto metrics = _grid->GetCellLayoutMetrics(host, row.value(), 0u);
        if (! metrics.hasIcon)
            continue;
        const auto icon = metrics.iconRect;
        const D2D1_POINT_2F center{icon.right - 2.0f, icon.bottom - 2.0f};
        const auto disc = D2D1::Ellipse(center, 4.5f, 4.5f);
        context->FillEllipse(disc, dimBrush);
        context->DrawEllipse(disc, crossBrush, 1.0f);
        context->DrawLine(D2D1::Point2F(center.x - 2.5f, center.y + 2.5f), D2D1::Point2F(center.x + 2.5f, center.y - 2.5f), crossBrush, 1.5f);
    }
    context->PopAxisAlignedClip();
}

void AnalysisHost::SetLegend(std::span<const LegendEntry> entries)
{
    const size_t nextCount = std::min(entries.size(), _legendLabels.size());
    bool changed           = nextCount != _legendCount;
    _legendCount           = nextCount;
    for (size_t index = 0u; index < _legendLabels.size(); ++index)
    {
        if (! _legendLabels[index])
            continue;
        _legendLabels[index]->SetVisible(index < _legendCount);
        _legendSwatches[index]->SetVisible(index < _legendCount);
        if (index >= _legendCount)
            continue;
        changed = changed || _legendLabels[index]->GetText() != entries[index].text;
        _legendLabels[index]->SetText(entries[index].text);
        _legendSwatches[index]->SetTextColor(_host.GetTheme().highContrast ? std::nullopt : std::optional<D2D1_COLOR_F>(entries[index].color));
    }
    if (changed)
        Layout();
}

bool AnalysisHost::RegisterChildClass() noexcept
{
    static std::atomic<uint64_t> nextClassId{1u};
    try
    {
        const uint64_t classId = nextClassId.fetch_add(1u, std::memory_order_relaxed);
        _windowClassName       = std::format(L"RedSalamander.ViewerSpace.AnalysisHost.{:016X}", classId);
        WNDCLASSEXW windowClass{};
        windowClass.cbSize        = sizeof(windowClass);
        windowClass.hInstance     = g_hInstance;
        windowClass.lpszClassName = _windowClassName.c_str();
        windowClass.lpfnWndProc   = WindowProc;
        windowClass.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        windowClass.style         = CS_HREDRAW | CS_VREDRAW;
        _windowClassAtom          = RegisterClassExW(&windowClass);
        return _windowClassAtom != 0;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
        return false;
    }
}

void AnalysisHost::BuildControls()
{
    _panelSplitter = _root->AddChild<DxUi::Splitter>();
    _panelSplitter->SetAccessibleName(_captions.gridAccessibleName);
    _panelSplitter->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.PanelSplitter");
    _panelSplitter->SetOnChange([this](DxUi::SplitterChange change)
    {
        if (_callbacks.panelResize)
            _callbacks.panelResize(change);
    });
    _filterLabel = _root->AddChild<DxUi::Label>(_captions.filterLabel);
    _filterLabel->SetAccessibleName(_captions.filterLabel);

    _filterField = _root->AddChild<DxUi::TextField>();
    _filterField->SetPlaceholder(_captions.filterPlaceholder);
    _filterField->SetClearButtonEnabled(true);
    _filterField->SetAccessibleName(_captions.filterLabel);
    _filterField->SetAccessibleHelpText(_captions.filterScope);
    _filterField->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.Filter");
    _filterField->SetOnTextChanged([this](std::wstring_view text) { InvokeFilterChanged(text); });

    _reviewCleanup   = _root->AddChild<DxUi::Button>();
    _recycle         = _root->AddChild<DxUi::Button>();
    _permanentDelete = _root->AddChild<DxUi::Button>();
    const std::array<DxUi::Button*, 3u> actions{{_reviewCleanup, _recycle, _permanentDelete}};
    const std::array<std::wstring_view, 3u> names{{_captions.reviewCleanup, _captions.recycle, _captions.permanentDelete}};
    const std::array<std::wstring_view, 3u> glyphs{{L"\uE7C1", L"\uE74D", L"\uE107"}};
    const std::array<UINT, 3u> commands{{IDM_VIEWERSPACE_REVIEW_CLEANUP, IDM_VIEWERSPACE_SEND_RECYCLE, IDM_VIEWERSPACE_SEND_DELETE}};
    for (size_t index = 0u; index < actions.size(); ++index)
    {
        actions[index]->SetVariant(index == 0u ? DxUi::ButtonVariant::IconOnly : DxUi::ButtonVariant::Standard);
        actions[index]->SetText(std::wstring(index == 0u ? glyphs[index] : names[index]));
        actions[index]->SetAccessibleName(std::wstring(names[index]));
        actions[index]->SetTooltipText(std::wstring(names[index]));
        actions[index]->SetAccessibleAutomationId(std::format(L"ViewerSpaceAnalysis.Action.{}", commands[index]));
        actions[index]->SetOnClick([this, command = commands[index]] { InvokeWorkspaceAction(command); });
        actions[index]->SetVisible(index == 0u);
    }

    for (size_t index = 0u; index < _legendLabels.size(); ++index)
    {
        _legendSwatches[index] = _root->AddChild<DxUi::Label>(L"■");
        _legendLabels[index]   = _root->AddChild<DxUi::Label>();
        _legendLabels[index]->SetFontRole(DxUi::FontRole::Small);
        _legendLabels[index]->SetVisible(false);
        _legendSwatches[index]->SetVisible(false);
    }
    _scopeLabel = _root->AddChild<DxUi::Label>(_captions.filterScope);
    _scopeLabel->SetMultiline(true);
    _scopeLabel->SetAccessibilityRole(DxUi::AccessibilityRole::Status);
    _scopeLabel->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.FilterScope");

    _grid = _root->AddChild<DxUi::Grid>();
    _grid->SetSelectionMode(DxUi::GridSelectionMode::Single);
    _grid->SetEmptyStateText(_captions.emptyState);
    _grid->SetAccessibleName(_captions.gridAccessibleName);
    _grid->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.Grid");
    _grid->SetCellTextFontRole(DxUi::FontRole::Body);

    _inspectorTitle = _root->AddChild<DxUi::Label>(_captions.inspectorTitle);
    _inspectorTitle->SetFontRole(DxUi::FontRole::BodyStrong);
    _inspectorTitle->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.InspectorTitle");
    _inspectorToggle = _root->AddChild<DxUi::Button>();
    _inspectorToggle->SetVariant(DxUi::ButtonVariant::Disclosure);
    _inspectorToggle->SetAccessibleName(_captions.inspectorTitle);
    _inspectorToggle->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.InspectorToggle");
    _inspectorToggle->SetOnClick([this]
    {
        SetInspectorLayout(! _inspectorVisible, _inspectorRatio);
        if (_callbacks.inspectorLayoutChanged)
            _callbacks.inspectorLayoutChanged(_inspectorVisible, _inspectorRatio);
    });
    _inspectorSplitter = _root->AddChild<DxUi::Splitter>();
    _inspectorSplitter->SetAccessibleName(_captions.inspectorTitle);
    _inspectorSplitter->SetAccessibleAutomationId(L"ViewerSpaceAnalysis.InspectorSplitter");
    _inspectorSplitter->SetOnChange([this](DxUi::SplitterChange change)
    {
        const auto bounds = _inspectorSplitter->GetBounds();
        const float extent =
            _inspectorSplitter->GetOrientation() == DxUi::SplitterOrientation::Vertical ? bounds.right - bounds.left : bounds.bottom - bounds.top;
        if (extent <= 0.0f)
            return;
        if (change.phase == DxUi::SplitterChangePhase::Preview && ! _inspectorResizeStartRatio.has_value())
            _inspectorResizeStartRatio = _inspectorRatio;
        _inspectorRatio = std::clamp((extent - change.position - 6.0f) / extent, 0.20f, 0.70f);
        if (change.phase == DxUi::SplitterChangePhase::Cancel && _inspectorResizeStartRatio.has_value())
            _inspectorRatio = _inspectorResizeStartRatio.value();
        Layout();
        if (change.phase != DxUi::SplitterChangePhase::Preview)
        {
            _inspectorResizeStartRatio.reset();
            if (_callbacks.inspectorLayoutChanged)
                _callbacks.inspectorLayoutChanged(_inspectorVisible, _inspectorRatio);
        }
    });
    constexpr std::array<std::wstring_view, 7u> valueAutomationIds{{L"Path", L"Name", L"Kind", L"LogicalBytes", L"Share", L"Modified", L"State"}};
    for (size_t index = 0u; index < _inspectorFieldLabels.size(); ++index)
    {
        _inspectorFieldLabels[index] = _root->AddChild<DxUi::Label>(_captions.inspectorFields[index]);
        _inspectorFieldLabels[index]->SetAccessibleName(_captions.inspectorFields[index]);
        _inspectorFieldLabels[index]->SetFocusable(false);

        _inspectorValues[index] = _root->AddChild<DxUi::Label>();
        _inspectorValues[index]->SetMultiline(true);
        _inspectorValues[index]->SetAccessibleAutomationId(std::format(L"ViewerSpaceAnalysis.Inspector.{}", valueAutomationIds[index]));
    }
}

void AnalysisHost::Layout() noexcept
{
    const auto layoutStarted = std::chrono::steady_clock::now();
    if (_host.GetRoot() == nullptr || _host.GetHwnd() == nullptr)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(_host.GetHwnd(), &client) == 0)
    {
        return;
    }
    const float widthDip     = _host.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip    = _host.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    const float right        = std::max(kMarginDip, widthDip - kMarginDip);
    const bool compactHeight = heightDip < 280.0f;
    const float margin       = compactHeight ? 8.0f : kMarginDip;
    const float fieldHeight  = compactHeight ? 28.0f : kFilterHeightDip;
    _host.GetRoot()->SetBounds(Rect(0.0f, 0.0f, widthDip, heightDip));
    // Only the separator lives in this HWND. Its first pane is the parent's D2D treemap;
    // translate the canonical Splitter's full workspace bounds into this child to share
    // pointer capture, keyboard, cancellation and UIA RangeValue behavior across renderers.
    _panelSplitter->SetOrientation(_panelVertical ? DxUi::SplitterOrientation::Vertical : DxUi::SplitterOrientation::Horizontal);
    _panelSplitter->SetMinimumFirstPane(_panelVertical ? std::max(250.0f, _panelExtent * 0.35f - 6.0f) : std::max(136.0f, _panelExtent * 0.20f - 6.0f));
    _panelSplitter->SetMinimumSecondPane(_panelVertical ? std::max(280.0f, _panelExtent * 0.20f) : std::max(320.0f, _panelExtent * 0.20f));
    _panelSplitter->SetBounds(_panelVertical ? Rect(-_panelPosition, 0.0f, _panelExtent - _panelPosition, heightDip)
                                             : Rect(0.0f, -_panelPosition, widthDip, _panelExtent - _panelPosition));
    _panelSplitter->SetPosition(_panelPosition);

    const float contentWidth           = std::max(0.0f, right - margin);
    const float filterLabelWidth       = MeasureTextWidthDip(_host, DxUi::FontRole::Body, _captions.filterLabel, 72.0f) + 8.0f;
    const float actionWidth            = _reviewingCleanup ? 350.0f : 38.0f;
    const float filterRight            = std::max(margin + 40.0f, right - actionWidth);
    const float horizontalToolbarWidth = filterLabelWidth + kGapDip + kMinimumFilterFieldWidthDip + actionWidth;
    const bool horizontalToolbar       = _alternateWorkspace || contentWidth >= horizontalToolbarWidth;

    const float actionsTop = horizontalToolbar ? margin : margin + 19.0f;
    _reviewCleanup->SetBounds(Rect(right - 30.0f, actionsTop, right, actionsTop + fieldHeight));
    _recycle->SetBounds(Rect(right - 188.0f, actionsTop, right - 36.0f, actionsTop + fieldHeight));
    _permanentDelete->SetBounds(Rect(right - 346.0f, actionsTop, right - 194.0f, actionsTop + fieldHeight));
    float contentTop = 0.0f;
    if (_reviewingCleanup && contentWidth < 350.0f)
    {
        _recycle->SetBounds(Rect(margin, margin, right - 36.0f, margin + fieldHeight));
        _permanentDelete->SetBounds(Rect(margin, margin + fieldHeight + 6.0f, right - 36.0f, margin + fieldHeight * 2.0f + 6.0f));
        const float scopeTop    = margin + fieldHeight * 2.0f + 14.0f;
        const float scopeHeight = std::max(18.0f, MeasureWrappedTextHeightDip(_host, DxUi::FontRole::Body, contentWidth, _scopeLabel->GetText()) + 2.0f);
        _scopeLabel->SetBounds(Rect(margin, scopeTop, right, scopeTop + scopeHeight));
        contentTop = scopeTop + scopeHeight + kGapDip;
    }
    else if (! horizontalToolbar)
    {
        const float labelTop = margin;
        const float fieldTop = labelTop + 19.0f;
        _filterLabel->SetBounds(Rect(margin, labelTop, right, labelTop + 17.0f));
        _filterField->SetBounds(Rect(margin, fieldTop, filterRight, fieldTop + fieldHeight));

        const float scopeTop    = fieldTop + fieldHeight + (compactHeight ? 2.0f : 6.0f);
        const float scopeHeight = std::max(18.0f, MeasureWrappedTextHeightDip(_host, DxUi::FontRole::Body, contentWidth, _scopeLabel->GetText()) + 2.0f);
        _scopeLabel->SetBounds(Rect(margin, scopeTop, right, scopeTop + scopeHeight));
        contentTop = scopeTop + scopeHeight + (compactHeight ? 4.0f : 8.0f);
    }
    else
    {
        const float toolbarTop = margin;
        const float filterLeft = margin + filterLabelWidth + kGapDip;
        _filterLabel->SetBounds(Rect(margin, toolbarTop, margin + filterLabelWidth, toolbarTop + fieldHeight));
        _filterField->SetBounds(Rect(filterLeft, toolbarTop, filterRight, toolbarTop + fieldHeight));
        const float scopeTop    = toolbarTop + fieldHeight + (compactHeight ? 3.0f : 8.0f);
        const float scopeHeight = std::max(18.0f, MeasureWrappedTextHeightDip(_host, DxUi::FontRole::Body, contentWidth, _scopeLabel->GetText()) + 2.0f);
        _scopeLabel->SetBounds(Rect(margin, scopeTop, right, scopeTop + scopeHeight));
        contentTop = scopeTop + scopeHeight + (compactHeight ? 4.0f : 8.0f);
    }

    if (_legendCount)
    {
        const size_t columns    = std::max<size_t>(1u, static_cast<size_t>(contentWidth / 150.0f));
        const float columnWidth = contentWidth / static_cast<float>(columns);
        for (size_t index = 0u; index < _legendCount; ++index)
        {
            const float x = margin + static_cast<float>(index % columns) * columnWidth;
            const float y = contentTop + static_cast<float>(index / columns) * 19.0f;
            _legendSwatches[index]->SetBounds(Rect(x, y, x + 16.0f, y + 18.0f));
            _legendLabels[index]->SetBounds(Rect(x + 18.0f, y, x + columnWidth - 2.0f, y + 18.0f));
        }
        contentTop += static_cast<float>((_legendCount + columns - 1u) / columns) * 19.0f + 4.0f;
    }
    const float bottom = std::max(contentTop, heightDip - margin);

    const bool sideInspector = widthDip >= kSideBySideMinimumWidthDip;
    _inspectorSplitter->SetVisible(_inspectorVisible);
    _inspectorSplitter->SetOrientation(sideInspector ? DxUi::SplitterOrientation::Vertical : DxUi::SplitterOrientation::Horizontal);
    _inspectorSplitter->SetBounds(Rect(margin, contentTop, right, bottom));
    const float splitExtent = sideInspector ? right - margin : bottom - contentTop;
    _inspectorSplitter->SetMinimumFirstPane(std::max(sideInspector ? 240.0f : (compactHeight ? 48.0f : 72.0f), splitExtent * 0.30f - 6.0f));
    _inspectorSplitter->SetMinimumSecondPane(std::max(sideInspector ? 200.0f : (widthDip >= 480.0f ? 94.0f : 154.0f), splitExtent * 0.20f));
    _inspectorSplitter->SetPosition(std::max(0.0f, splitExtent * (1.0f - _inspectorRatio) - 6.0f));
    if (! _inspectorVisible)
    {
        _grid->SetBounds(Rect(margin, contentTop, right, std::max(contentTop, bottom - 28.0f)));
        LayoutInspector(margin, std::max(contentTop, bottom - 26.0f), right, bottom);
    }
    else
    {
        _grid->SetBounds(_inspectorSplitter->GetFirstPaneBounds());
        const auto inspector = _inspectorSplitter->GetSecondPaneBounds();
        LayoutInspector(inspector.left, inspector.top, inspector.right, inspector.bottom);
    }
    const float gridWidth = _grid->GetBounds().right - _grid->GetBounds().left;
    if (std::abs(gridWidth - _gridLayoutWidthDip) > 0.5f)
    {
        _gridLayoutWidthDip = gridWidth;
        std::array<DxUi::GridColumnLayoutEntry, 4u> columns;
        for (size_t index = 0u; index < columns.size(); ++index)
            columns[index] = {ColumnId(index), index, _model.GetColumn(index).widthDip};
        _grid->ApplyColumnLayout(columns);
    }
    _host.RefreshAccessibilitySnapshot();
    _host.Invalidate();
    Debug::Perf::EmitDurationUs(L"viewer.space.workspace.layout_us", Debug::Perf::ElapsedUs(layoutStarted), _inspectorVisible ? 1u : 0u);
}

void AnalysisHost::InvokeFilterChanged(std::wstring_view text) noexcept
{
    if (! _callbacks.filterChanged)
    {
        return;
    }
    try
    {
        _callbacks.filterChanged(text);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void AnalysisHost::InvokePanelClose() noexcept
{
    if (! _callbacks.panelClose)
    {
        return;
    }
    try
    {
        _callbacks.panelClose();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void AnalysisHost::InvokeFocusBoundary(bool reverse, bool& handled) noexcept
{
    handled = false;
    if (! _callbacks.focusBoundary)
    {
        return;
    }
    try
    {
        handled = _callbacks.focusBoundary(reverse);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

bool AnalysisHost::InvokeCommandKey(UINT virtualKey, UINT modifiers) noexcept
{
    if (! _callbacks.commandKey)
    {
        return false;
    }
    try
    {
        _callbacks.commandKey(virtualKey, modifiers);
        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
        return false;
    }
}

void AnalysisHost::InvokeSelectionChanged(std::optional<uint64_t> rowId) noexcept
{
    if (! _callbacks.selectionChanged)
    {
        return;
    }
    try
    {
        _callbacks.selectionChanged(rowId);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void AnalysisHost::LogCallbackFailureOnce() noexcept
{
    if (! _callbackFailureLogged)
    {
        _callbackFailureLogged = true;
        Debug::Error(L"ViewerSpace analysis callback failed.");
    }
}

std::optional<uint64_t> AnalysisHost::StableRowIdAt(size_t rowIndex) const noexcept
{
    const uint64_t id = _model.GetStableRowId(rowIndex);
    return id == 0u ? std::nullopt : std::optional<uint64_t>(id);
}

void AnalysisHost::LayoutInspector(float left, float top, float right, float bottom) noexcept
{
    if (_inspectorTitle == nullptr)
        return;
    // Product-specific compact identity/facts presentation; all exact values remain
    // on the same stable UIA elements even when visual text is ellipsized.
    const std::array<std::wstring_view, 7u> values{_inspectorText.path,
                                                   _inspectorText.name,
                                                   _inspectorText.icon,
                                                   _inspectorText.logicalBytes,
                                                   _inspectorText.share,
                                                   _inspectorText.modified,
                                                   _inspectorText.state};
    for (size_t index = 0u; index < values.size(); ++index)
    {
        _inspectorFieldLabels[index]->SetVisible(false);
        _inspectorValues[index]->SetVisible(false);
    }
    const float width       = std::max(0.0f, right - left);
    const float titleBottom = std::min(bottom, top + 24.0f);
    _inspectorToggle->SetDisclosureExpanded(_inspectorVisible);
    _inspectorToggle->SetBounds(Rect(left, top, left + 22.0f, titleBottom));
    const bool selected = ! _inspectorText.name.empty();
    _inspectorTitle->SetVisible(! selected);
    _inspectorTitle->SetBounds(Rect(left + 24.0f, top, right, titleBottom));
    const auto value = [&](size_t index, float x, float y, float end, float height, bool middle = false) noexcept
    {
        if (end <= x || y + height > bottom + 0.5f || values[index].empty())
            return;
        auto* label = _inspectorValues[index];
        label->SetVisible(true);
        label->SetMultiline(false);
        label->SetBounds(Rect(x, y, end, y + height));
        label->SetText(BuildInspectorVisualText(_host, values[index], end - x, middle));
    };
    if (! selected)
    {
        _inspectorValues[6]->SetMultiline(true);
        _inspectorValues[6]->SetText(_inspectorText.state);
        _inspectorValues[6]->SetVisible(true);
        _inspectorValues[6]->SetBounds(Rect(left, titleBottom + 4.0f, right, bottom));
        return;
    }
    _inspectorValues[2]->SetFontRole(DxUi::FontRole::Icon);
    _inspectorValues[1]->SetFontRole(DxUi::FontRole::BodyStrong);
    value(2u, left + 24.0f, top, left + 46.0f, 22.0f);
    value(1u, left + 48.0f, top, right, 22.0f);
    value(0u, left + 24.0f, top + 25.0f, right, 18.0f, true);
    const bool wide        = width >= 380.0f;
    const float metricTop  = top + 48.0f;
    const float firstWidth = wide ? width * 0.28f : width * 0.5f;
    for (size_t index = 3u; index <= 5u; ++index)
    {
        const float x   = index == 5u && ! wide ? left : left + static_cast<float>(index - 3u) * firstWidth;
        const float y   = metricTop + (index == 5u && ! wide ? 38.0f : 0.0f);
        const float end = index == 5u ? right : std::min(right, x + firstWidth - 6.0f);
        if (y + 34.0f <= bottom + 0.5f)
        {
            _inspectorFieldLabels[index]->SetVisible(true);
            _inspectorFieldLabels[index]->SetBounds(Rect(x, y, end, y + 15.0f));
            value(index, x, y + 16.0f, end, 18.0f);
        }
    }
    const float stateTop = metricTop + (wide ? 39.0f : 77.0f);
    value(6u, left, stateTop, right, 18.0f);
}
} // namespace ViewerSpaceAnalysis

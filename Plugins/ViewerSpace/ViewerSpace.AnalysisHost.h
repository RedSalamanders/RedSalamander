#pragma once

#include <windows.h>

#include <array>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>

#include <DxUi/DxUi.h>

#include "ViewerSpace.Hosts.h"

namespace ViewerSpaceAnalysis
{
// A child DxUi workspace for the virtualized analysis table and its linked inspector.
// Callbacks are invoked synchronously on the owning UI thread and are cleared by Detach().
class AnalysisHost final
{
public:
    struct Captions final
    {
        std::wstring filterLabel;
        std::wstring filterPlaceholder;
        std::wstring filterScope;
        std::wstring gridAccessibleName;
        std::wstring emptyState;
        std::array<std::wstring, 5u> columns;
        std::wstring inspectorTitle;
        std::wstring back;
        std::wstring reviewCleanup;
        std::wstring recycle;
        std::wstring permanentDelete;
        std::array<std::wstring, 7u> inspectorFields;
    };

    struct InspectorText final
    {
        std::wstring path;
        std::wstring name;
        std::wstring kind;
        std::wstring logicalBytes;
        std::wstring share;
        std::wstring modified;
        std::wstring state;
        std::wstring fullPath;
        std::wstring exactBytes;
        std::wstring icon;
    };

    struct Callbacks final
    {
        std::function<size_t()> rowCount;
        std::function<uint64_t(size_t rowIndex)> rowId;
        std::function<std::optional<size_t>(uint64_t rowId)> findRow;
        std::function<void(size_t rowIndex, size_t columnIndex, DxUi::GridCellData& outCell)> cellText;
        std::function<void(std::optional<uint64_t> rowId)> selectionChanged;
        std::function<void(uint64_t rowId)> rowActivated;
        std::function<void(uint64_t rowId, POINT screenPoint)> rowContextMenu;
        std::function<void(uint64_t rowId)> exclusionToggled;
        std::function<void(uint64_t rowId)> cleanupToggled;
        std::function<void(UINT action)> workspaceAction;
        std::function<bool(size_t row)> rowExcluded;
        std::function<void(const DxUi::GridSortSpec& sortSpec)> sortRequested;
        std::function<void(std::wstring_view text)> filterChanged;
        std::function<void()> panelClose;
        std::function<void(DxUi::SplitterChange change)> panelResize;
        std::function<void(bool visible, float ratio)> inspectorLayoutChanged;
        std::function<void(UINT virtualKey, UINT modifiers)> commandKey;
        std::function<bool(bool reverse)> focusBoundary;
    };

    AnalysisHost() noexcept;
    ~AnalysisHost() noexcept;

    AnalysisHost(const AnalysisHost&)            = delete;
    AnalysisHost(AnalysisHost&&)                 = delete;
    AnalysisHost& operator=(const AnalysisHost&) = delete;
    AnalysisHost& operator=(AnalysisHost&&)      = delete;

    [[nodiscard]] HRESULT Attach(HWND parent, const Captions& captions, Callbacks callbacks) noexcept;
    void Detach() noexcept;
    void SetBounds(const RECT& boundsPx) noexcept;
    void ApplyTheme(const DxUi::ThemePalette& palette) noexcept;
    void NotifyRowsChanged(std::optional<uint64_t> selectedRowId);
    void SetInspector(const InspectorText& text);
    void SetWorkspaceScope(bool reviewing, bool problems, std::wstring_view scope);
    struct LegendEntry final
    {
        std::wstring text;
        D2D1_COLOR_F color{};
    };
    void SetLegend(std::span<const LegendEntry> entries);
    void SetFilterText(std::wstring_view text);
    void SetSortSpec(const DxUi::GridSortSpec& sortSpec) noexcept;
    void SetCompactRows(bool compactRows) noexcept;
    void SetPanelSplitter(bool vertical, float extentDip, float positionDip) noexcept;
    void SetInspectorLayout(bool visible, float ratio) noexcept;
    void CancelResize() noexcept;

    [[nodiscard]] HWND GetHwnd() const noexcept;
    [[nodiscard]] const DxUi::ThemePalette& GetTheme() const noexcept
    {
        return _host.GetTheme();
    }
    [[nodiscard]] ViewerSpaceHosts::HostDiagnostics GetDiagnostics() const noexcept
    {
        return ViewerSpaceHosts::CaptureHostDiagnostics(_host, _hwnd.get());
    }
    [[nodiscard]] D2D1_RECT_F GetRootBounds() const noexcept
    {
        return _host.GetRoot() ? _host.GetRoot()->GetBounds() : D2D1_RECT_F{};
    }
    [[nodiscard]] bool OwnsFocus() const noexcept;
    [[nodiscard]] bool Focus(bool reverse) noexcept;

private:
    class Model final : public DxUi::IGridModel
    {
    public:
        Model() noexcept               = default;
        Model(const Model&)            = delete;
        Model(Model&&)                 = delete;
        Model& operator=(const Model&) = delete;
        Model& operator=(Model&&)      = delete;
        void Bind(AnalysisHost& owner) noexcept;

        [[nodiscard]] size_t GetRowCount() const noexcept override;
        [[nodiscard]] size_t GetColumnCount() const noexcept override;
        [[nodiscard]] DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override;
        void GetCellData(size_t rowIndex, size_t columnIndex, DxUi::GridCellData& outCell) const override;
        [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override;
        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override;

    private:
        AnalysisHost* _owner = nullptr;
    };

    class Delegate final : public DxUi::IGridDelegate
    {
    public:
        using DxUi::IGridDelegate::OnGridCheckboxToggled;
        using DxUi::IGridDelegate::OnGridContextMenu;
        using DxUi::IGridDelegate::OnGridRowActivated;
        using DxUi::IGridDelegate::OnGridSelectionChanged;

        Delegate() noexcept                  = default;
        Delegate(const Delegate&)            = delete;
        Delegate(Delegate&&)                 = delete;
        Delegate& operator=(const Delegate&) = delete;
        Delegate& operator=(Delegate&&)      = delete;
        void Bind(AnalysisHost& owner) noexcept;

        void OnGridSortRequested(const DxUi::GridSortSpec& sortSpec) override;
        void OnGridSelectionChanged(DxUi::Grid& sender) override;
        void OnGridRowActivated(DxUi::Grid& sender, size_t rowIndex) override;
        void OnGridContextMenu(DxUi::Grid& sender, size_t rowIndex, POINT screenPoint) override;
        void OnGridCheckboxToggled(DxUi::Grid& sender, size_t rowIndex, size_t columnIndex, bool checked) override;

    private:
        AnalysisHost* _owner = nullptr;
    };

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] LRESULT HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] bool RegisterChildClass() noexcept;
    void BuildControls();
    void Layout() noexcept;
    void LayoutInspector(float left, float top, float right, float bottom) noexcept;
    void PaintExclusionOverlay(DxUi::ControlHost& host) const;
    void InvokeFilterChanged(std::wstring_view text) noexcept;
    void InvokePanelClose() noexcept;
    void InvokeFocusBoundary(bool reverse, bool& handled) noexcept;
    [[nodiscard]] bool InvokeCommandKey(UINT virtualKey, UINT modifiers) noexcept;
    void InvokeSelectionChanged(std::optional<uint64_t> rowId) noexcept;
    void InvokeExclusionToggled(uint64_t rowId) noexcept;
    void InvokeWorkspaceAction(UINT action) noexcept;
    void LogCallbackFailureOnce() noexcept;
    [[nodiscard]] std::optional<uint64_t> StableRowIdAt(size_t rowIndex) const noexcept;

    DxUi::WindowHost _host;
    wil::unique_hwnd _hwnd;
    std::wstring _windowClassName;
    ATOM _windowClassAtom = 0;
    Captions _captions;
    Callbacks _callbacks;
    std::unique_ptr<DxUi::Panel> _root;
    Model _model;
    Delegate _delegate;
    DxUi::Label* _filterLabel      = nullptr;
    DxUi::TextField* _filterField  = nullptr;
    DxUi::Label* _scopeLabel       = nullptr;
    DxUi::Grid* _grid              = nullptr;
    DxUi::Button* _reviewCleanup   = nullptr;
    DxUi::Button* _recycle         = nullptr;
    DxUi::Button* _permanentDelete = nullptr;
    bool _reviewingCleanup         = false;
    bool _alternateWorkspace       = false;
    std::array<DxUi::Label*, 8u> _legendSwatches{};
    std::array<DxUi::Label*, 8u> _legendLabels{};
    size_t _legendCount                = 0u;
    DxUi::Label* _inspectorTitle       = nullptr;
    DxUi::Button* _inspectorToggle     = nullptr;
    DxUi::Splitter* _inspectorSplitter = nullptr;
    DxUi::Splitter* _panelSplitter     = nullptr;
    float _panelExtent                 = 0.0f;
    float _panelPosition               = 0.0f;
    bool _panelVertical                = false;
    bool _inspectorVisible             = true;
    float _inspectorRatio              = 0.38f;
    std::optional<float> _inspectorResizeStartRatio;
    std::array<DxUi::Label*, 7u> _inspectorFieldLabels{};
    std::array<DxUi::Label*, 7u> _inspectorValues{};
    InspectorText _inspectorText;
    RECT _boundsPx{};
    float _gridLayoutWidthDip   = -1.0f;
    bool _compactRows           = false;
    bool _callbackFailureLogged = false;
};
} // namespace ViewerSpaceAnalysis

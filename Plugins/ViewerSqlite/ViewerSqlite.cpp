#include "ViewerSqlite.h"

#include <algorithm>
#include <cstring>
#include <cwchar>
#include <format>
#include <functional>
#include <limits>
#include <new>
#include <utility>

#include <dwmapi.h>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "Dwmapi.lib")

#include "Helpers.h"
#include "ViewerTitleBarTheme.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::StatusStrip;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;

constexpr UINT kAsyncOpenCompleteMessage  = WndMsg::kViewerSqliteAsyncOpenComplete;
constexpr UINT kAsyncQueryCompleteMessage = WndMsg::kViewerSqliteAsyncQueryComplete;

constexpr uint32_t kMaxConfiguredPageSize    = 1000u;
constexpr uint32_t kMaxConfiguredQueryRowCap = 100000u;

const int kViewerSqliteModuleAnchor = 0;
std::atomic_uint32_t g_viewerSqliteWindowCount{0u};

constexpr char kViewerSqliteSchemaJson[] = R"json({
  "version": 1,
  "title": "SQLite Viewer",
  "fields": [
    {
      "key": "pageSize",
      "type": "value",
      "label": "Page size",
      "description": "Maximum rows loaded for table preview pages.",
      "default": 200,
      "min": 1,
      "max": 1000
    },
    {
      "key": "queryRowCap",
      "type": "value",
      "label": "Query row cap",
      "description": "Maximum rows materialized for custom read-only SQL queries.",
      "default": 2000,
      "min": 1,
      "max": 100000
    },
    {
      "key": "directOpenLocalFiles",
      "type": "bool",
      "label": "Optimized local snapshot",
      "description": "Use SQLite's backup API for a transactionally consistent private snapshot of local databases and their WAL state.",
      "default": true
    }
  ]
})json";

[[nodiscard]] const char* GetViewerSqliteStaticConfigurationSchemaImpl() noexcept
{
    return kViewerSqliteSchemaJson;
}

struct ViewerSqliteAsyncWorkItem final
{
    ViewerSqliteAsyncWorkItem()                                            = default;
    ViewerSqliteAsyncWorkItem(const ViewerSqliteAsyncWorkItem&)            = delete;
    ViewerSqliteAsyncWorkItem& operator=(const ViewerSqliteAsyncWorkItem&) = delete;

    wil::unique_hmodule moduleKeepAlive;
    std::function<void()> work;
#ifdef _DEBUG
    std::shared_ptr<std::atomic_uint32_t> pendingAsyncWork;
#endif
};

[[nodiscard]] std::wstring LeafNameForPath(std::wstring_view path)
{
    std::filesystem::path fsPath(path);
    std::wstring leaf = fsPath.filename().wstring();
    if (leaf.empty())
    {
        leaf.assign(path);
    }

    return leaf;
}

[[nodiscard]] std::wstring ReadStatusText(UINT resourceId)
{
    return LoadStringResource(g_hInstance, resourceId);
}

[[nodiscard]] std::wstring ResolveErrorText(UINT fallbackId, std::wstring text)
{
    if (! text.empty())
    {
        return text;
    }

    return ReadStatusText(fallbackId);
}

[[nodiscard]] bool QueueThreadpoolWork(std::unique_ptr<ViewerSqliteAsyncWorkItem> workItem) noexcept
{
    if (! workItem || ! workItem->moduleKeepAlive)
    {
        return false;
    }

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
    {
        std::unique_ptr<ViewerSqliteAsyncWorkItem> owned(static_cast<ViewerSqliteAsyncWorkItem*>(context));
        if (! owned)
        {
            return;
        }

        if (owned->moduleKeepAlive)
        {
            TransferModulePinToCallbackReturn(instance, owned->moduleKeepAlive);
        }

        if (owned->work)
        {
            owned->work();
        }

#ifdef _DEBUG
        if (owned->pendingAsyncWork)
        {
            static_cast<void>(owned->pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
        }
#endif
    },
        workItem.get(),
        nullptr);

    if (queued == 0)
    {
        return false;
    }

    workItem.release();
    return true;
}

[[nodiscard]] ViewerSqliteEngine::TableSortDirection ToEngineSortDirection(SortDirection direction) noexcept
{
    switch (direction)
    {
        case SortDirection::Ascending: return ViewerSqliteEngine::TableSortDirection::Ascending;
        case SortDirection::Descending: return ViewerSqliteEngine::TableSortDirection::Descending;
        case SortDirection::None:
        default: return ViewerSqliteEngine::TableSortDirection::None;
    }
}

[[nodiscard]] std::wstring BuildTableStatusText(const ViewerSqliteEngine::QueryPage& page, std::wstring_view tableName)
{
    const uint64_t firstRow = page.rows.empty() ? 0u : page.rowOffset + 1u;
    const uint64_t lastRow  = page.rowOffset + static_cast<uint64_t>(page.rows.size());
    return FormatStringResource(
        g_hInstance, page.hasMore ? IDS_VIEWERSQLITE_STATUS_TABLE_PAGE_MORE_FMT : IDS_VIEWERSQLITE_STATUS_TABLE_PAGE_FMT, tableName, firstRow, lastRow);
}

[[nodiscard]] std::wstring_view FindTableDisplayName(const std::vector<ViewerSqliteEngine::TableInfo>& tables, const std::wstring_view tableName) noexcept
{
    const auto match = std::find_if(tables.begin(), tables.end(), [&](const ViewerSqliteEngine::TableInfo& table) noexcept { return table.name == tableName; });
    return match != tables.end() ? std::wstring_view(match->displayName) : std::wstring_view{};
}

[[nodiscard]] std::wstring BuildQueryStatusText(const ViewerSqliteEngine::QueryPage& page)
{
    return FormatStringResource(
        g_hInstance, page.truncated ? IDS_VIEWERSQLITE_STATUS_QUERY_TRUNCATED_FMT : IDS_VIEWERSQLITE_STATUS_QUERY_FMT, static_cast<uint64_t>(page.rows.size()));
}

[[nodiscard]] bool CanAdvanceTablePage(const uint64_t rowOffset, const uint32_t pageSize) noexcept
{
    return rowOffset <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max()) - pageSize;
}
} // namespace

const char* GetViewerSqliteStaticConfigurationSchema() noexcept
{
    return GetViewerSqliteStaticConfigurationSchemaImpl();
}

struct ViewerSqliteGridModel final : IDxGridModel
{
    void SetPage(const ViewerSqliteEngine::QueryPage* page, bool tablePreviewMode)
    {
        _page             = page;
        _tablePreviewMode = tablePreviewMode;
        RebuildColumns();
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _page ? _page->rows.size() : 0u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (! _page || rowIndex >= _page->rows.size())
        {
            return;
        }

        const auto& row = _page->rows[rowIndex];
        if (columnIndex >= row.size())
        {
            return;
        }

        outCell.text      = row[columnIndex];
        outCell.multiline = true;
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        GridRowStyle style{};
        if (! _page || rowIndex >= _page->rows.size())
        {
            return style;
        }

        const auto& row   = _page->rows[rowIndex];
        const auto seedIt = std::find_if(row.begin(), row.end(), [](const std::wstring& value) noexcept { return ! value.empty(); });
        if (seedIt != row.end())
        {
            style.rainbowSeed = *seedIt;
        }
        return style;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        if (! _page)
        {
            return static_cast<uint64_t>(rowIndex);
        }

        return _page->rowOffset + static_cast<uint64_t>(rowIndex) + 1u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (! _page || rowId == 0u)
        {
            return std::nullopt;
        }

        const uint64_t firstRowId = _page->rowOffset + 1u;
        const uint64_t lastRowId  = firstRowId + static_cast<uint64_t>(_page->rows.size());
        if (rowId < firstRowId || rowId >= lastRowId)
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId - firstRowId);
    }

private:
    void RebuildColumns()
    {
        _columns.clear();

        size_t columnCount = 0u;
        if (_page)
        {
            columnCount = _page->columns.size();
            for (const auto& row : _page->rows)
            {
                columnCount = std::max(columnCount, row.size());
            }
        }

        if (columnCount == 0u)
        {
            return;
        }

        _columns.reserve(columnCount);
        for (size_t columnIndex = 0; columnIndex < columnCount; ++columnIndex)
        {
            GridColumnDesc column{};
            column.id = std::format(L"column-{}", columnIndex);
            if (_page && columnIndex < _page->columns.size() && ! _page->columns[columnIndex].name.empty())
            {
                column.title = _page->columns[columnIndex].name;
            }
            else
            {
                column.title = (columnCount == 1u) ? LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_COLUMN_RESULT)
                                                   : FormatStringResource(g_hInstance, IDS_VIEWERSQLITE_COLUMN_FORMAT, columnIndex + 1u);
            }

            column.widthDip    = (columnIndex == 0u) ? 180.0f : 160.0f;
            column.minWidthDip = 96.0f;
            column.sortable    = _tablePreviewMode;
            column.multiline   = true;
            _columns.push_back(std::move(column));
        }
    }

    const ViewerSqliteEngine::QueryPage* _page = nullptr;
    bool _tablePreviewMode                     = false;
    std::vector<GridColumnDesc> _columns;
};

ViewerSqlite::ViewerSqlite()
{
    _metaId          = L"builtin/viewer-sqlite";
    _metaShortId     = L"sqlite";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_DESCRIPTION);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    _configurationJson = "{}";
    _gridModel         = std::make_unique<ViewerSqliteGridModel>();
}

ViewerSqlite::~ViewerSqlite() = default;

void ViewerSqlite::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerSqlite::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE ViewerSqlite::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (remaining == 0)
    {
        delete this;
    }

    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerSqliteStaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    _config = {};

    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    yyjson_doc* doc = yyjson_read(configurationJsonUtf8, std::strlen(configurationJsonUtf8), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (yyjson_val* value = yyjson_obj_get(root, "pageSize"); value != nullptr)
    {
        if (! yyjson_is_uint(value))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        _config.pageSize = std::clamp<uint32_t>(static_cast<uint32_t>(yyjson_get_uint(value)), 1u, kMaxConfiguredPageSize);
    }

    if (yyjson_val* value = yyjson_obj_get(root, "queryRowCap"); value != nullptr)
    {
        if (! yyjson_is_uint(value))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        _config.queryRowCap = std::clamp<uint32_t>(static_cast<uint32_t>(yyjson_get_uint(value)), 1u, kMaxConfiguredQueryRowCap);
    }

    if (yyjson_val* value = yyjson_obj_get(root, "directOpenLocalFiles"); value != nullptr)
    {
        if (! yyjson_is_bool(value))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        _config.directOpenLocalFiles = yyjson_get_bool(value);
    }

    _configurationJson = configurationJsonUtf8;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    *pSomethingToSave = FALSE;
    return S_OK;
}

ATOM ViewerSqlite::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return atom;
    }

    const std::wstring& className = GetWindowClassName();
    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.hInstance     = instance;
    wc.lpszClassName = className.c_str();
    wc.lpfnWndProc   = WndProcThunk;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.style         = CS_HREDRAW | CS_VREDRAW;

    atom = RegisterClassExW(&wc);
    if (atom != 0)
    {
        return atom;
    }

    if (GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
    {
        return 0;
    }

    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(nullptr, className.c_str(), &existing) == 0)
    {
        return 0;
    }

    if (existing.hInstance == instance && existing.lpfnWndProc == WndProcThunk)
    {
        atom = 1;
        return atom;
    }

    if (g_viewerSqliteWindowCount.load(std::memory_order_acquire) != 0u)
    {
        return 0;
    }

    if (existing.hInstance == nullptr || UnregisterClassW(className.c_str(), existing.hInstance) == 0)
    {
        return 0;
    }

    atom = RegisterClassExW(&wc);
    return atom;
}

const std::wstring& ViewerSqlite::GetWindowClassName() noexcept
{
    static const std::wstring className = std::format(L"{}.{}", kClassName, GetTickCount64());
    return className;
}

LRESULT CALLBACK ViewerSqlite::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        if (cs != nullptr && cs->lpCreateParams != nullptr)
        {
            auto* self = static_cast<ViewerSqlite*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
            static_cast<void>(g_viewerSqliteWindowCount.fetch_add(1u, std::memory_order_relaxed));
        }
    }

    auto* self = reinterpret_cast<ViewerSqlite*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self != nullptr)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerSqlite::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(hwnd, msg, wp, lp, dxHandled);
    if (dxHandled)
    {
        if (msg == WM_SIZE)
        {
            OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp)));
        }
        return dxResult;
    }

    switch (msg)
    {
#ifdef _DEBUG
        case WndMsg::kViewerSqliteDebugGetPendingAsyncWork:
            return static_cast<LRESULT>(_pendingAsyncWork ? _pendingAsyncWork->load(std::memory_order_relaxed) : 0u);
        case WndMsg::kViewerSqliteDebugGetSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSqliteDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                         = {};
            snapshot->pendingAsyncWork        = _pendingAsyncWork ? _pendingAsyncWork->load(std::memory_order_relaxed) : 0u;
            snapshot->rowCount                = _gridModel ? _gridModel->GetRowCount() : 0u;
            snapshot->tablePreviewMode        = _tablePreviewMode;
            snapshot->hasMoreRows             = _currentHasMore;
            snapshot->prevButtonEnabled       = _prevButton && _prevButton->IsEnabled();
            snapshot->nextButtonEnabled       = _nextButton && _nextButton->IsEnabled();
            snapshot->rowOffset               = _currentRowOffset;
            snapshot->themeDark               = _dxHost.GetTheme().dark;
            snapshot->themeHighContrast       = _dxHost.GetTheme().highContrast;
            snapshot->themeRainbow            = _dxHost.GetTheme().rainbowMode;
            snapshot->sortColumnIndex         = (_tableSortSpec.direction == SortDirection::None) ? static_cast<size_t>(-1) : _tableSortSpec.columnIndex;
            snapshot->sortDirection           = static_cast<uint32_t>(_tableSortSpec.direction);
            snapshot->hasStatusStrip          = _statusStrip != nullptr;
            snapshot->statusStripVisible      = _statusStrip && _statusStrip->IsVisible();
            snapshot->statusStripSectionCount = _statusStrip ? static_cast<uint32_t>(_statusStrip->GetSectionCount()) : 0u;
            if (_statusStrip)
            {
                const D2D1_RECT_F statusBounds = _statusStrip->GetBounds();
                snapshot->statusStripHeightDip = (statusBounds.bottom > statusBounds.top) ? (statusBounds.bottom - statusBounds.top) : 0.0f;
            }
            if (_resultGrid)
            {
                const auto metrics             = _resultGrid->GetVisibleWorkMetrics();
                snapshot->selectionCount       = _resultGrid->GetSelectionModel().GetCount();
                snapshot->visibleRowCount      = static_cast<size_t>(metrics.visibleRowCount);
                snapshot->visibleColumnCount   = metrics.visibleColumnCount;
                snapshot->visibleCellCount     = static_cast<size_t>(metrics.visibleCellCount);
                snapshot->hasVerticalScrollbar = metrics.hasVerticalScrollbar;
                if (const auto selectedRow = _resultGrid->GetPrimarySelectedRow(); selectedRow.has_value())
                {
                    snapshot->primarySelectedRowId = _gridModel ? _gridModel->GetStableRowId(selectedRow.value()) : 0u;
                    RedSalamander::DxUi::GridDebugRowVisualState rowVisualState{};
                    if (_resultGrid->DebugGetRowVisualState(_dxHost.GetTheme(), selectedRow.value(), rowVisualState))
                    {
                        snapshot->selectedRowFillArgb    = rowVisualState.fillArgb;
                        snapshot->selectedRowTextArgb    = rowVisualState.textArgb;
                        snapshot->selectedRowUsesRainbow = rowVisualState.usesRainbow;
                    }
                }
            }
            snapshot->renderCount        = _dxHost.DebugGetRenderCount();
            snapshot->resizeCount        = _dxHost.DebugGetResizeCount();
            snapshot->resizeFailureCount = _dxHost.DebugGetResizeFailureCount();
            if (RedSalamander::DxUi::Control* const focusedControl = _dxHost.GetFocusControl(); focusedControl != nullptr)
            {
                if (focusedControl == _fileCombo)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::FileCombo;
                }
                else if (focusedControl == _reloadButton)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::ReloadButton;
                }
                else if (focusedControl == _tableCombo)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::TableCombo;
                }
                else if (focusedControl == _prevButton)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::PrevButton;
                }
                else if (focusedControl == _nextButton)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::NextButton;
                }
                else if (focusedControl == _queryField)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::QueryField;
                }
                else if (focusedControl == _runButton)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::RunButton;
                }
                else if (focusedControl == _tableButton)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::TableButton;
                }
                else if (focusedControl == _resultGrid)
                {
                    snapshot->focusTarget = WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid;
                }
            }
            const auto parsePrimaryKey = [](const std::wstring& value) noexcept
            {
                if (value.empty())
                {
                    return 0ull;
                }

                wchar_t* end                    = nullptr;
                const unsigned long long parsed = std::wcstoull(value.c_str(), &end, 10);
                return (end != nullptr && *end == L'\0') ? parsed : 0ull;
            };
            if (! _currentPage.rows.empty() && ! _currentPage.rows.front().empty())
            {
                snapshot->firstRowPrimaryKey = parsePrimaryKey(_currentPage.rows.front().front());
            }
            if (! _currentPage.rows.empty() && ! _currentPage.rows.back().empty())
            {
                snapshot->lastRowPrimaryKey = parsePrimaryKey(_currentPage.rows.back().front());
            }
            return TRUE;
        }
        case WndMsg::kViewerSqliteDebugScrollGridByWheelDetents:
        {
            const int detents = static_cast<int>(wp);
            if (! _resultGrid || detents == 0)
            {
                return detents == 0 ? TRUE : FALSE;
            }

            _dxHost.SetFocusControl(_resultGrid);
            const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
            const int stepCount    = detents > 0 ? detents : -detents;
            for (int remaining = stepCount; remaining > 0; --remaining)
            {
                if (! _resultGrid->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u))
                {
                    return FALSE;
                }
            }
            return TRUE;
        }
        case WndMsg::kViewerSqliteDebugSelectGridRow:
        {
            if (! _resultGrid || ! _gridModel)
            {
                return FALSE;
            }

            const size_t rowIndex = static_cast<size_t>(wp);
            if (rowIndex >= _gridModel->GetRowCount())
            {
                return FALSE;
            }

            return _resultGrid->RequestSelectRow(rowIndex, 0u) ? TRUE : FALSE;
        }
        case WndMsg::kViewerSqliteDebugInvokePageCommand:
        {
            if (_loading || ! _tablePreviewMode || ! _databaseSource || _currentTable.empty())
            {
                return FALSE;
            }

            const auto command = static_cast<WndMsg::ViewerSqliteDebugPageCommand>(wp);
            switch (command)
            {
                case WndMsg::ViewerSqliteDebugPageCommand::Previous:
                    if (_currentRowOffset < _config.pageSize)
                    {
                        return FALSE;
                    }

                    QueueTablePreview(_currentTable, _currentRowOffset - _config.pageSize);
                    return TRUE;
                case WndMsg::ViewerSqliteDebugPageCommand::Next:
                    if (! _currentHasMore)
                    {
                        return FALSE;
                    }

                    if (! CanAdvanceTablePage(_currentRowOffset, _config.pageSize))
                    {
                        return FALSE;
                    }
                    QueueTablePreview(_currentTable, _currentRowOffset + _config.pageSize);
                    return TRUE;
                default: return FALSE;
            }
        }
        case WndMsg::kViewerSqliteDebugCycleSortColumn:
        {
            if (_loading || ! _tablePreviewMode || ! _databaseSource || _currentTable.empty() || ! _gridModel)
            {
                return FALSE;
            }

            const size_t columnIndex = static_cast<size_t>(wp);
            if (columnIndex >= _gridModel->GetColumnCount())
            {
                return FALSE;
            }

            GridSortSpec nextSort{};
            nextSort.columnIndex = columnIndex;
            nextSort.direction =
                (_tableSortSpec.columnIndex == columnIndex) ? RedSalamander::DxUi::NextSortDirection(_tableSortSpec.direction) : SortDirection::Ascending;
            OnGridSortRequested(nextSort);
            return TRUE;
        }
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_DPICHANGED: OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 720, 420);
            }
            return 0;
        case WM_NCACTIVATE: ApplyTitleBarTheme(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_CLOSE: static_cast<void>(Close()); return 0;
        case kAsyncOpenCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncOpenResult>(lp);
            if (lp != 0 && ! result)
            {
                return 0;
            }
            OnAsyncOpenComplete(std::move(result), static_cast<uint64_t>(wp));
            return 0;
        }
        case kAsyncQueryCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncQueryResult>(lp);
            if (lp != 0 && ! result)
            {
                return 0;
            }
            OnAsyncQueryComplete(std::move(result), static_cast<uint64_t>(wp));
            return 0;
        }
        case WM_NCDESTROY:
        {
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            _dxHost.Detach();
            _root         = nullptr;
            _resultGrid   = nullptr;
            _statusStrip  = nullptr;
            _fileLabel    = nullptr;
            _fileCombo    = nullptr;
            _reloadButton = nullptr;
            _tableLabel   = nullptr;
            _tableCombo   = nullptr;
            _prevButton   = nullptr;
            _nextButton   = nullptr;
            _queryLabel   = nullptr;
            _queryField   = nullptr;
            _runButton    = nullptr;
            _tableButton  = nullptr;
            _hWnd.release();

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            const uint32_t remainingWindows = g_viewerSqliteWindowCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
            if (remainingWindows == 0u)
            {
                static_cast<void>(UnregisterClassW(GetWindowClassName().c_str(), g_hInstance));
            }

            NotifyViewerClosed();

            Release();
            return 0;
        }
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::Open(const ViewerOpenContext* context) noexcept
{
    if (context == nullptr || context->fileSystem == nullptr || context->focusedPath == nullptr || context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (! RegisterWndClass(g_hInstance))
    {
        return E_FAIL;
    }

    _fileSystem = context->fileSystem;
    LoadOtherFiles(context);
    _currentPath = _otherFiles.empty() ? std::wstring(context->focusedPath) : _otherFiles[_otherIndex];

    const bool embeddedMode   = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (embeddedParent == nullptr || IsWindow(embeddedParent) == FALSE))
    {
        Debug::Error(L"ViewerSqlite: embedded Open requires a valid ownerWindow parent.");
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    _currentQuery.clear();
    _currentHasMore   = false;
    _currentRowOffset = 0;
    _tablePreviewMode = true;
    _databaseSource.reset();
    _tables.clear();
    _currentPage = {};
    ResetTableSort();

    if (! _hWnd)
    {
        RECT ownerRect{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1100;
        int h = 760;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) == 0)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerSqlite: GetClientRect failed for embedded preview parent.");
                return HRESULT_FROM_WIN32(lastError);
            }
            x = 0;
            y = 0;
            w = std::max(1L, client.right - client.left);
            h = std::max(1L, client.bottom - client.top);
        }
        else if (context->ownerWindow != nullptr && GetWindowRect(context->ownerWindow, &ownerRect) != 0)
        {
            x = ownerRect.left;
            y = ownerRect.top;
            w = std::max(1L, ownerRect.right - ownerRect.left);
            h = std::max(1L, ownerRect.bottom - ownerRect.top);
        }

        const DWORD style = embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS) : WS_OVERLAPPEDWINDOW;
        const std::wstring initialTitle =
            embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_NAME) : _metaName);
        HWND window = CreateWindowExW(0, GetWindowClassName().c_str(), initialTitle.c_str(), style, x, y, w, h, embeddedParent, nullptr, g_hInstance, this);
        if (! window)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ViewerSqlite: CreateWindowExW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }

        _hWnd.reset(window);
        _embeddedMode = embeddedMode;
        AddRef();
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        if (_embeddedMode)
        {
            RECT client{};
            if (embeddedParent != nullptr && GetClientRect(embeddedParent, &client) != 0)
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             0,
                             0,
                             std::max(1L, client.right - client.left),
                             std::max(1L, client.bottom - client.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }

    if (_queryField)
    {
        _queryField->SetText(_currentQuery);
    }
    RefreshFileCombo();
    RefreshTableCombo();
    ClearResults();
    UpdateWindowTitle();
    UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_LOADING));
    UpdateControlState();
    QueueOpenCurrentPath();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::Close() noexcept
{
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    static_cast<void>(_requestId.fetch_add(1u, std::memory_order_relaxed));
    _hWnd.reset();
    _embeddedMode = false;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSqlite::SetTheme(const ViewerTheme* theme) noexcept
{
    if (theme == nullptr || theme->version < 2u || theme->version > 4u)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;
    _dpi      = (_theme.dpi == 0u) ? USER_DEFAULT_SCREEN_DPI : _theme.dpi;

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        Layout();
    }

    return S_OK;
}

void ViewerSqlite::OnGridSortRequested(const GridSortSpec& sortSpec)
{
    if (_loading || ! _tablePreviewMode || ! _databaseSource || _currentTable.empty() || ! _gridModel)
    {
        return;
    }

    if (sortSpec.direction != SortDirection::None && sortSpec.columnIndex >= _gridModel->GetColumnCount())
    {
        return;
    }

    _tableSortSpec = sortSpec;
    if (_resultGrid)
    {
        _resultGrid->SetSortSpec(_tableSortSpec);
    }
    QueueTablePreview(_currentTable, 0);
}

void ViewerSqlite::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    if (_dpi == 0u)
    {
        _dpi = USER_DEFAULT_SCREEN_DPI;
    }

    static_cast<void>(_dxHost.Attach(hwnd));
    BuildUi();
    ApplyTheme(hwnd);
    Layout();
    RefreshFileCombo();
    RefreshTableCombo();
    UpdateWindowTitle();
    UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_READY));
    UpdateControlState();
    FocusMainContentIfPossible();
}

void ViewerSqlite::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    Layout();
}

void ViewerSqlite::OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggestedRect) noexcept
{
    _dpi = (dpi == 0u) ? USER_DEFAULT_SCREEN_DPI : dpi;
    if (suggestedRect != nullptr)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOACTIVATE | SWP_NOZORDER);
    }

    ApplyTheme(hwnd);
    Layout();
}

void ViewerSqlite::OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result, const uint64_t requestIdFromMessage) noexcept
{
    const uint64_t requestId = requestIdFromMessage != 0u ? requestIdFromMessage : (result ? result->requestId : 0u);
    if (requestId == 0u || requestId != _requestId.load(std::memory_order_relaxed))
    {
        return;
    }

    _loading = false;

    if (! result)
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_OUT_OF_MEMORY));
        UpdateControlState();
        return;
    }

    if (FAILED(result->hr) || ! result->source)
    {
        _databaseSource.reset();
        _tables.clear();
        _currentTable.clear();
        _currentHasMore   = false;
        _currentRowOffset = 0;
        _tablePreviewMode = true;
        _currentPage      = {};
        ResetTableSort();
        RefreshTableCombo();
        ClearResults();
        UpdateStatusText(ResolveErrorText(IDS_VIEWERSQLITE_ERROR_OPEN_FAILED, std::move(result->errorText)));
        UpdateControlState();
        return;
    }

    _databaseSource   = std::move(result->source);
    _tables           = std::move(result->tables);
    _currentHasMore   = false;
    _currentRowOffset = 0;
    _tablePreviewMode = true;
    _currentTable     = std::move(result->initialTable);
    _currentPath      = std::move(result->path);
    ResetTableSort();

    RefreshTableCombo();
    UpdateWindowTitle();

    if (result->hasInitialPage)
    {
        _currentHasMore = result->initialPage.hasMore;
        PopulateResults(std::move(result->initialPage));
        UpdateStatusText(BuildTableStatusText(_currentPage, FindTableDisplayName(_tables, _currentTable)));
    }
    else
    {
        ClearResults();
        UpdateStatusText(_tables.empty() ? ReadStatusText(IDS_VIEWERSQLITE_STATUS_NO_TABLES) : ReadStatusText(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE));
    }

    UpdateControlState();
    FocusMainContentIfPossible();
}

void ViewerSqlite::OnAsyncQueryComplete(std::unique_ptr<AsyncQueryResult> result, const uint64_t requestIdFromMessage) noexcept
{
    const uint64_t requestId = requestIdFromMessage != 0u ? requestIdFromMessage : (result ? result->requestId : 0u);
    if (requestId == 0u || requestId != _requestId.load(std::memory_order_relaxed))
    {
        return;
    }

    _loading = false;

    if (! result)
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_OUT_OF_MEMORY));
        UpdateControlState();
        return;
    }

    if (FAILED(result->hr))
    {
        UpdateStatusText(ResolveErrorText(IDS_VIEWERSQLITE_ERROR_QUERY_FAILED, std::move(result->errorText)));
        UpdateControlState();
        return;
    }

    _currentHasMore = result->page.hasMore;

    if (result->tablePreview)
    {
        _tablePreviewMode = true;
        _currentTable     = std::move(result->tableName);
        _currentRowOffset = result->rowOffset;
        RefreshTableCombo();
        PopulateResults(std::move(result->page));
        UpdateStatusText(BuildTableStatusText(_currentPage, FindTableDisplayName(_tables, _currentTable)));
    }
    else
    {
        _tablePreviewMode = false;
        _currentQuery     = std::move(result->sql);
        if (_queryField && _queryField->GetText() != _currentQuery)
        {
            _queryField->SetText(_currentQuery);
        }
        PopulateResults(std::move(result->page));
        UpdateStatusText(BuildQueryStatusText(_currentPage));
    }

    UpdateControlState();
    FocusMainContentIfPossible();
}

void ViewerSqlite::BuildUi() noexcept
{
    if (_root != nullptr)
    {
        return;
    }

    auto root = std::make_unique<Panel>();
    _root     = root.get();

    _fileLabel = _root->AddChild<Label>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_LABEL_FILE));
    _fileLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _fileCombo = _root->AddChild<ComboBox>();
    _fileCombo->SetVariant(ComboBoxVariant::Window);
    _fileCombo->SetOnSelectionChanged([this](size_t index)
    {
        if (! _syncingFileCombo)
        {
            StartFileSelection(index);
        }
    });

    _reloadButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_BUTTON_RELOAD));
    _reloadButton->SetOnClick([this] { QueueOpenCurrentPath(); });

    _tableLabel = _root->AddChild<Label>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_LABEL_TABLE));
    _tableLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _tableCombo = _root->AddChild<ComboBox>();
    _tableCombo->SetVariant(ComboBoxVariant::Window);
    _tableCombo->SetOnSelectionChanged([this](size_t /*index*/)
    {
        if (! _syncingTableCombo)
        {
            StartSelectedTablePreview(true);
        }
    });

    _prevButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_BUTTON_PREVIOUS));
    _prevButton->SetOnClick([this]
    {
        if (_currentRowOffset >= _config.pageSize)
        {
            QueueTablePreview(_currentTable, _currentRowOffset - _config.pageSize);
        }
    });

    _nextButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_BUTTON_NEXT));
    _nextButton->SetOnClick([this]
    {
        if (CanAdvanceTablePage(_currentRowOffset, _config.pageSize))
        {
            QueueTablePreview(_currentTable, _currentRowOffset + _config.pageSize);
        }
    });

    _queryLabel = _root->AddChild<Label>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_LABEL_QUERY));
    _queryLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _queryField = _root->AddChild<TextField>(_currentQuery);
    _queryField->SetOnTextChanged([this](std::wstring_view text) { _currentQuery.assign(text); });
    _queryField->SetOnSubmitted([this] { QueueCustomQuery(_currentQuery); });

    _runButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_BUTTON_RUN_QUERY));
    _runButton->SetOnClick([this] { QueueCustomQuery(_currentQuery); });

    _tableButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERSQLITE_BUTTON_TABLE_PREVIEW));
    _tableButton->SetOnClick([this] { StartSelectedTablePreview(true); });

    _resultGrid = _root->AddChild<Grid>();
    _resultGrid->SetDelegate(this);
    _resultGrid->SetModel(_gridModel.get());
    _resultGrid->SetHeaderHeightDip(30.0f);
    _resultGrid->SetRowHeightDip(34.0f);
    _resultGrid->SetLineClamp(2u);

    _statusStrip = _root->AddChild<StatusStrip>(ReadStatusText(IDS_VIEWERSQLITE_STATUS_READY));
    _statusStrip->SetFontRole(RedSalamander::DxUi::FontRole::Small);

    _dxHost.SetRoot(std::move(root));
    _dxHost.SetDefaultButton(_runButton);
    _dxHost.SetOnEscape([this]() noexcept
    {
        if (FocusMainContentIfPossible())
        {
            return true;
        }
        if (_hWnd)
        {
            PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
        }
        return true;
    });
}

void ViewerSqlite::ApplyTheme(HWND hwnd) noexcept
{
    ThemePalette palette = _hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme) : RedSalamander::DxUi::MakeDefaultThemePalette(false);
    _dxHost.SetTheme(palette);
    if (! _embeddedMode)
    {
        ApplyTitleBarTheme(GetActiveWindow() == hwnd);
    }
}

void ViewerSqlite::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    RedSalamander::ViewerChrome::ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive, L"title");
}

void ViewerSqlite::Layout() noexcept
{
    if (! _hWnd || ! _root)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(_hWnd.get(), &client) == 0)
    {
        return;
    }

    const float widthDip  = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));

    const float margin           = 12.0f;
    const float gap              = 8.0f;
    const float labelWidth       = 46.0f;
    const float rowHeight        = 30.0f;
    const float buttonWidth      = 104.0f;
    const float navButtonWidth   = 90.0f;
    const float queryButtonWidth = 116.0f;
    const float statusHeight     = 22.0f;

    const float right = std::max(margin, widthDip - margin);
    float y           = margin;

    const float fileComboLeft  = margin + labelWidth + gap;
    const float fileComboWidth = std::max(120.0f, right - fileComboLeft - gap - buttonWidth);
    _fileLabel->SetBounds(D2D1::RectF(margin, y, margin + labelWidth, y + rowHeight));
    _fileCombo->SetBounds(D2D1::RectF(fileComboLeft, y, fileComboLeft + fileComboWidth, y + rowHeight));
    _reloadButton->SetBounds(D2D1::RectF(right - buttonWidth, y, right, y + rowHeight));

    y += rowHeight + gap;

    const float tableComboLeft  = margin + labelWidth + gap;
    const float tableComboWidth = std::max(120.0f, right - tableComboLeft - gap - navButtonWidth - gap - navButtonWidth);
    _tableLabel->SetBounds(D2D1::RectF(margin, y, margin + labelWidth, y + rowHeight));
    _tableCombo->SetBounds(D2D1::RectF(tableComboLeft, y, tableComboLeft + tableComboWidth, y + rowHeight));
    _prevButton->SetBounds(D2D1::RectF(right - navButtonWidth - gap - navButtonWidth, y, right - gap - navButtonWidth, y + rowHeight));
    _nextButton->SetBounds(D2D1::RectF(right - navButtonWidth, y, right, y + rowHeight));

    y += rowHeight + gap;

    const float queryFieldLeft  = margin + labelWidth + gap;
    const float queryFieldWidth = std::max(120.0f, right - queryFieldLeft - gap - queryButtonWidth - gap - queryButtonWidth);
    _queryLabel->SetBounds(D2D1::RectF(margin, y, margin + labelWidth, y + rowHeight));
    _queryField->SetBounds(D2D1::RectF(queryFieldLeft, y, queryFieldLeft + queryFieldWidth, y + rowHeight));
    _runButton->SetBounds(D2D1::RectF(right - queryButtonWidth - gap - queryButtonWidth, y, right - gap - queryButtonWidth, y + rowHeight));
    _tableButton->SetBounds(D2D1::RectF(right - queryButtonWidth, y, right, y + rowHeight));

    y += rowHeight + gap;

    const float gridBottom = std::max(y + 80.0f, heightDip - margin - gap - statusHeight);
    _resultGrid->SetBounds(D2D1::RectF(margin, y, std::max(margin + 1.0f, widthDip - margin), gridBottom));
    _statusStrip->SetBounds(D2D1::RectF(margin, gridBottom + gap, std::max(margin + 1.0f, widthDip - margin), gridBottom + gap + statusHeight));

    _dxHost.Invalidate();
}

void ViewerSqlite::RefreshFileCombo() noexcept
{
    if (! _fileCombo)
    {
        return;
    }

    _syncingFileCombo    = true;
    const auto resetSync = wil::scope_exit([&] { _syncingFileCombo = false; });

    std::vector<ComboBox::Item> items;
    items.reserve(_otherFiles.size());
    for (const auto& file : _otherFiles)
    {
        items.push_back(ComboBox::Item{file, LeafNameForPath(file)});
    }

    _fileCombo->SetItems(std::move(items));
    if (! _otherFiles.empty())
    {
        _fileCombo->SetSelectedIndex(std::min(_otherIndex, _otherFiles.size() - 1u));
    }
    else
    {
        _fileCombo->SetSelectedIndex(std::nullopt);
    }

    _dxHost.Invalidate();
}

void ViewerSqlite::RefreshTableCombo() noexcept
{
    if (! _tableCombo)
    {
        return;
    }

    _syncingTableCombo   = true;
    const auto resetSync = wil::scope_exit([&] { _syncingTableCombo = false; });

    std::vector<ComboBox::Item> items;
    items.reserve(_tables.size());

    std::optional<size_t> selectedIndex;
    for (size_t index = 0; index < _tables.size(); ++index)
    {
        const auto& table = _tables[index];
        items.push_back(ComboBox::Item{table.name, table.displayName});
        if (! selectedIndex && table.name == _currentTable)
        {
            selectedIndex = index;
        }
    }

    if (! selectedIndex && ! _tables.empty())
    {
        _currentTable = _tables.front().name;
        selectedIndex = 0u;
    }

    _tableCombo->SetItems(std::move(items));
    _tableCombo->SetSelectedIndex(selectedIndex);
    _dxHost.Invalidate();
}

void ViewerSqlite::ClearResults() noexcept
{
    _currentPage = {};
    if (_gridModel)
    {
        _gridModel->SetPage(nullptr, _tablePreviewMode);
    }
    if (_resultGrid)
    {
        _resultGrid->SetSortSpec(_tablePreviewMode ? _tableSortSpec : GridSortSpec{});
        _resultGrid->NotifyDataChanged();
    }
    _dxHost.Invalidate();
}

void ViewerSqlite::PopulateResults(ViewerSqliteEngine::QueryPage page) noexcept
{
    _currentPage = std::move(page);
    if (_gridModel)
    {
        _gridModel->SetPage(&_currentPage, _tablePreviewMode);
    }
    if (_resultGrid)
    {
        _resultGrid->SetSortSpec(_tablePreviewMode ? _tableSortSpec : GridSortSpec{});
        _resultGrid->NotifyDataChanged();
    }
    _dxHost.Invalidate();
}

void ViewerSqlite::UpdateWindowTitle() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const std::wstring fileName = _currentPath.empty() ? _metaName : LeafNameForPath(_currentPath);
    const std::wstring title    = FormatEmbeddedStringResource(g_hInstance, IDS_VIEWERSQLITE_TITLE_FORMAT, fileName, _metaName);
    SetWindowTextW(_hWnd.get(), title.c_str());
}

void ViewerSqlite::UpdateStatusText(std::wstring text) noexcept
{
    if (text.empty())
    {
        text = ReadStatusText(IDS_VIEWERSQLITE_STATUS_READY);
    }

    if (_statusStrip)
    {
        _statusStrip->SetText(std::move(text));
    }
    _dxHost.Invalidate();
}

void ViewerSqlite::UpdateControlState() noexcept
{
    const bool hasDatabase   = _databaseSource != nullptr;
    const bool hasTables     = ! _tables.empty();
    const bool canRunQuery   = ! _loading && hasDatabase;
    const bool canTablePager = ! _loading && hasDatabase && ! _currentTable.empty();

    if (_fileCombo)
    {
        _fileCombo->SetEnabled(! _loading && _otherFiles.size() > 1u);
    }
    if (_reloadButton)
    {
        _reloadButton->SetEnabled(! _loading && ! _currentPath.empty());
    }
    if (_tableCombo)
    {
        _tableCombo->SetEnabled(! _loading && hasTables);
    }
    if (_prevButton)
    {
        _prevButton->SetEnabled(canTablePager && _tablePreviewMode && _currentRowOffset >= _config.pageSize);
    }
    if (_nextButton)
    {
        _nextButton->SetEnabled(canTablePager && _tablePreviewMode && _currentHasMore && CanAdvanceTablePage(_currentRowOffset, _config.pageSize));
    }
    if (_runButton)
    {
        _runButton->SetEnabled(canRunQuery);
    }
    if (_tableButton)
    {
        _tableButton->SetEnabled(canTablePager && hasTables);
    }
    if (_queryField)
    {
        _queryField->SetEnabled(true);
    }
    if (_resultGrid)
    {
        _resultGrid->SetEnabled(! _currentPage.rows.empty());
        _resultGrid->SetHeaderBusy(_loading);
        _resultGrid->SetHeaderBusyColumn((_loading && _tablePreviewMode && _tableSortSpec.direction != SortDirection::None)
                                             ? std::optional<size_t>(_tableSortSpec.columnIndex)
                                             : std::nullopt);
    }

    if (auto* focusedControl = _dxHost.GetFocusControl(); focusedControl != nullptr && ! focusedControl->IsFocusable())
    {
        _dxHost.SetFocusControl(nullptr);
    }

    _dxHost.Invalidate();
}

bool ViewerSqlite::FocusMainContentIfPossible() noexcept
{
    if (_embeddedMode || ! _resultGrid || ! _resultGrid->IsFocusable())
    {
        return false;
    }

    if (_dxHost.GetFocusControl() == _resultGrid)
    {
        return false;
    }

    _dxHost.SetFocusControl(_resultGrid);
    return true;
}

void ViewerSqlite::ResetTableSort() noexcept
{
    _tableSortSpec = {};
    if (_resultGrid)
    {
        _resultGrid->SetSortSpec(GridSortSpec{});
    }
}

void ViewerSqlite::QueueOpenCurrentPath() noexcept
{
    if (! _hWnd || ! _fileSystem || _currentPath.empty())
    {
        return;
    }

    _loading = true;
    _databaseSource.reset();
    _tables.clear();
    _currentHasMore   = false;
    _currentRowOffset = 0;
    _tablePreviewMode = true;
    _currentPage      = {};
    ResetTableSort();

    RefreshTableCombo();
    ClearResults();
    UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_LOADING));
    UpdateControlState();

    const uint64_t requestId             = _requestId.fetch_add(1, std::memory_order_relaxed) + 1u;
    const HWND hwnd                      = _hWnd.get();
    const std::wstring path              = _currentPath;
    const std::wstring preferredTable    = _currentTable;
    ViewerSqliteEngine::Config config    = _config;
    wil::com_ptr<IFileSystem> fileSystem = _fileSystem;

#ifdef _DEBUG
    static_cast<void>(_pendingAsyncWork->fetch_add(1u, std::memory_order_relaxed));
#endif
    AddRef();

    auto workItem = std::unique_ptr<ViewerSqliteAsyncWorkItem>(new (std::nothrow) ViewerSqliteAsyncWorkItem{});
    if (! workItem)
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_OUT_OF_MEMORY));
        UpdateControlState();
        return;
    }

    workItem->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerSqliteModuleAnchor);
#ifdef _DEBUG
    workItem->pendingAsyncWork = _pendingAsyncWork;
#endif
    workItem->work = [this, hwnd, requestId, fileSystem = std::move(fileSystem), path, preferredTable, config]() mutable
    {
        const auto releaseSelf = wil::scope_exit([&] { Release(); });

        auto result = std::unique_ptr<AsyncOpenResult>(new (std::nothrow) AsyncOpenResult{});
        if (! result)
        {
            Debug::Error(L"SQLite viewer could not allocate async open completion result.");
            if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
            {
                static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, static_cast<WPARAM>(requestId), std::unique_ptr<AsyncOpenResult>{}));
            }
            return;
        }

        result->viewer    = this;
        result->requestId = requestId;
        result->path      = path;

        const ViewerSqliteEngine::QueryCancellation cancellation{&_requestId, requestId};
        auto opened       = ViewerSqliteEngine::OpenFromViewerContext(fileSystem.get(), path, config.directOpenLocalFiles, cancellation);
        result->hr        = opened.hr;
        result->errorText = std::move(opened.errorText);

        if (SUCCEEDED(result->hr) && opened.source)
        {
            result->source = std::move(opened.source);
            result->tables = std::move(opened.tables);

            if (! result->tables.empty())
            {
                auto selectedTableIt = std::find_if(
                    result->tables.begin(), result->tables.end(), [&](const ViewerSqliteEngine::TableInfo& table) { return table.name == preferredTable; });

                if (selectedTableIt == result->tables.end())
                {
                    selectedTableIt = result->tables.begin();
                }

                result->initialTable = selectedTableIt->name;
                auto page            = result->source->LoadTablePage(
                    result->initialTable, config.pageSize, 0, ViewerSqliteEngine::kNoSortColumn, ViewerSqliteEngine::TableSortDirection::None, cancellation);
                if (FAILED(page.hr))
                {
                    result->hr        = page.hr;
                    result->errorText = std::move(page.errorText);
                }
                else
                {
                    result->initialPage    = std::move(page.page);
                    result->hasInitialPage = true;
                }
            }
        }

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, static_cast<WPARAM>(requestId), std::move(result)));
    };

    if (! QueueThreadpoolWork(std::move(workItem)))
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_QUEUE_FAILED));
        UpdateControlState();
    }
}

void ViewerSqlite::QueueTablePreview(std::wstring tableName, uint64_t rowOffset) noexcept
{
    if (! _hWnd || ! _databaseSource || tableName.empty())
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE));
        return;
    }

    _loading = true;
    UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_LOADING_ROWS));
    UpdateControlState();

    const uint64_t requestId     = _requestId.fetch_add(1, std::memory_order_relaxed) + 1u;
    const HWND hwnd              = _hWnd.get();
    auto source                  = _databaseSource;
    const uint32_t pageSize      = _config.pageSize;
    const size_t sortColumnIndex = (_tableSortSpec.direction == SortDirection::None) ? ViewerSqliteEngine::kNoSortColumn : _tableSortSpec.columnIndex;
    const ViewerSqliteEngine::TableSortDirection sortDirection = ToEngineSortDirection(_tableSortSpec.direction);

#ifdef _DEBUG
    static_cast<void>(_pendingAsyncWork->fetch_add(1u, std::memory_order_relaxed));
#endif
    AddRef();

    auto workItem = std::unique_ptr<ViewerSqliteAsyncWorkItem>(new (std::nothrow) ViewerSqliteAsyncWorkItem{});
    if (! workItem)
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_OUT_OF_MEMORY));
        UpdateControlState();
        return;
    }

    workItem->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerSqliteModuleAnchor);
#ifdef _DEBUG
    workItem->pendingAsyncWork = _pendingAsyncWork;
#endif
    workItem->work =
        [this, hwnd, requestId, source = std::move(source), tableName = std::move(tableName), rowOffset, pageSize, sortColumnIndex, sortDirection]() mutable
    {
        const auto releaseSelf = wil::scope_exit([&] { Release(); });

        auto result = std::unique_ptr<AsyncQueryResult>(new (std::nothrow) AsyncQueryResult{});
        if (! result)
        {
            Debug::Error(L"SQLite viewer could not allocate async query completion result.");
            if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
            {
                static_cast<void>(PostMessagePayload(hwnd, kAsyncQueryCompleteMessage, static_cast<WPARAM>(requestId), std::unique_ptr<AsyncQueryResult>{}));
            }
            return;
        }

        result->viewer       = this;
        result->requestId    = requestId;
        result->tablePreview = true;
        result->tableName    = tableName;
        result->rowOffset    = rowOffset;

        const ViewerSqliteEngine::QueryCancellation cancellation{&_requestId, requestId};
        auto page         = source->LoadTablePage(result->tableName, pageSize, rowOffset, sortColumnIndex, sortDirection, cancellation);
        result->hr        = page.hr;
        result->errorText = std::move(page.errorText);
        result->page      = std::move(page.page);

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncQueryCompleteMessage, static_cast<WPARAM>(requestId), std::move(result)));
    };

    if (! QueueThreadpoolWork(std::move(workItem)))
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_QUEUE_FAILED));
        UpdateControlState();
    }
}

void ViewerSqlite::QueueCustomQuery(std::wstring sql) noexcept
{
    if (! _hWnd || ! _databaseSource)
    {
        return;
    }

    sql = StringUtils::TrimWhitespaceCopy(sql);
    if (sql.empty())
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_ENTER_QUERY));
        return;
    }

    _loading = true;
    UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_RUNNING_QUERY));
    UpdateControlState();

    const uint64_t requestId = _requestId.fetch_add(1, std::memory_order_relaxed) + 1u;
    const HWND hwnd          = _hWnd.get();
    auto source              = _databaseSource;
    const uint32_t rowCap    = _config.queryRowCap;

#ifdef _DEBUG
    static_cast<void>(_pendingAsyncWork->fetch_add(1u, std::memory_order_relaxed));
#endif
    AddRef();

    auto workItem = std::unique_ptr<ViewerSqliteAsyncWorkItem>(new (std::nothrow) ViewerSqliteAsyncWorkItem{});
    if (! workItem)
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_OUT_OF_MEMORY));
        UpdateControlState();
        return;
    }

    workItem->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerSqliteModuleAnchor);
#ifdef _DEBUG
    workItem->pendingAsyncWork = _pendingAsyncWork;
#endif
    workItem->work = [this, hwnd, requestId, source = std::move(source), sql = std::move(sql), rowCap]() mutable
    {
        const auto releaseSelf = wil::scope_exit([&] { Release(); });

        auto result = std::unique_ptr<AsyncQueryResult>(new (std::nothrow) AsyncQueryResult{});
        if (! result)
        {
            Debug::Error(L"SQLite viewer could not allocate async query completion result.");
            if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
            {
                static_cast<void>(PostMessagePayload(hwnd, kAsyncQueryCompleteMessage, static_cast<WPARAM>(requestId), std::unique_ptr<AsyncQueryResult>{}));
            }
            return;
        }

        result->viewer       = this;
        result->requestId    = requestId;
        result->tablePreview = false;
        result->sql          = sql;

        const ViewerSqliteEngine::QueryCancellation cancellation{&_requestId, requestId};
        auto query        = source->ExecuteReadOnlyQuery(result->sql, rowCap, cancellation);
        result->hr        = query.hr;
        result->errorText = std::move(query.errorText);
        result->page      = std::move(query.page);

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncQueryCompleteMessage, static_cast<WPARAM>(requestId), std::move(result)));
    };

    if (! QueueThreadpoolWork(std::move(workItem)))
    {
#ifdef _DEBUG
        static_cast<void>(_pendingAsyncWork->fetch_sub(1u, std::memory_order_relaxed));
#endif
        Release();
        _loading = false;
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_ERROR_QUEUE_FAILED));
        UpdateControlState();
    }
}

void ViewerSqlite::StartFileSelection(size_t index) noexcept
{
    if (index >= _otherFiles.size())
    {
        return;
    }

    _otherIndex  = index;
    _currentPath = _otherFiles[index];
    _currentTable.clear();
    QueueOpenCurrentPath();
}

void ViewerSqlite::StartSelectedTablePreview(bool resetSort) noexcept
{
    if (! _tableCombo || _tables.empty())
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE));
        return;
    }

    const std::optional<size_t> selected = _tableCombo->GetSelectedIndex();
    if (! selected || *selected >= _tables.size())
    {
        UpdateStatusText(ReadStatusText(IDS_VIEWERSQLITE_STATUS_SELECT_TABLE));
        return;
    }

    _currentTable = _tables[*selected].name;
    if (resetSort)
    {
        ResetTableSort();
    }
    QueueTablePreview(_currentTable, 0);
}

void ViewerSqlite::LoadOtherFiles(const ViewerOpenContext* context) noexcept
{
    _otherFiles.clear();
    _otherIndex = 0;

    if (context == nullptr)
    {
        return;
    }

    if (context->otherFiles != nullptr && context->otherFileCount > 0)
    {
        _otherFiles.reserve(static_cast<size_t>(context->otherFileCount));
        for (unsigned long index = 0; index < context->otherFileCount; ++index)
        {
            const wchar_t* path = context->otherFiles[index];
            if (path != nullptr && path[0] != L'\0')
            {
                _otherFiles.emplace_back(path);
            }
        }
    }

    if (_otherFiles.empty())
    {
        _otherFiles.emplace_back(context->focusedPath);
        _otherIndex = 0;
        return;
    }

    if (context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
        return;
    }

    const std::wstring focused = context->focusedPath;
    const auto it              = std::find(_otherFiles.begin(), _otherFiles.end(), focused);
    if (it != _otherFiles.end())
    {
        _otherIndex = static_cast<size_t>(std::distance(_otherFiles.begin(), it));
    }
}

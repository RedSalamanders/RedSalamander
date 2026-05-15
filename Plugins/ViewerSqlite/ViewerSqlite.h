#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstdint>
#include <memory>
#include <string>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "DxUi/DxUi.h"
#include "EmbeddedViewerBase.h"
#include "Helpers.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "ViewerSqlite.Engine.h"

struct ViewerSqliteGridModel;

[[nodiscard]] const char* GetViewerSqliteStaticConfigurationSchema() noexcept;

class ViewerSqlite final : public EmbeddedViewerBase<ViewerSqlite>, public IInformations, public RedSalamander::DxUi::IDxGridDelegate
{
public:
    ViewerSqlite();
    ~ViewerSqlite();

    void SetHost(IHost* host) noexcept;

    ViewerSqlite(const ViewerSqlite&)            = delete;
    ViewerSqlite(ViewerSqlite&&)                 = delete;
    ViewerSqlite& operator=(const ViewerSqlite&) = delete;
    ViewerSqlite& operator=(ViewerSqlite&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;

    HRESULT STDMETHODCALLTYPE Open(const ViewerOpenContext* context) noexcept override;
    HRESULT STDMETHODCALLTYPE Close() noexcept override;
    HRESULT STDMETHODCALLTYPE SetTheme(const ViewerTheme* theme) noexcept override;

    void OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec) override;

private:
    struct AsyncOpenResult
    {
        ViewerSqlite* viewer = nullptr;
        uint64_t requestId   = 0;
        HRESULT hr           = E_FAIL;
        std::wstring path;
        std::wstring errorText;
        std::shared_ptr<ViewerSqliteEngine::DatabaseSource> source;
        std::vector<ViewerSqliteEngine::TableInfo> tables;
        std::wstring initialTable;
        ViewerSqliteEngine::QueryPage initialPage;
        bool hasInitialPage = false;
    };

    struct AsyncQueryResult
    {
        ViewerSqlite* viewer = nullptr;
        uint64_t requestId   = 0;
        HRESULT hr           = E_FAIL;
        std::wstring errorText;
        std::wstring tableName;
        std::wstring sql;
        uint64_t rowOffset = 0;
        bool tablePreview  = true;
        ViewerSqliteEngine::QueryPage page;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    [[nodiscard]] static const std::wstring& GetWindowClassName() noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerSqlite";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd) noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggestedRect) noexcept;
    void OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result, uint64_t requestIdFromMessage) noexcept;
    void OnAsyncQueryComplete(std::unique_ptr<AsyncQueryResult> result, uint64_t requestIdFromMessage) noexcept;

    void BuildUi() noexcept;
    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void Layout() noexcept;
    void RefreshFileCombo() noexcept;
    void RefreshTableCombo() noexcept;
    void ClearResults() noexcept;
    void PopulateResults(ViewerSqliteEngine::QueryPage page) noexcept;
    void UpdateWindowTitle() noexcept;
    void UpdateStatusText(std::wstring text) noexcept;
    void UpdateControlState() noexcept;
    void ResetTableSort() noexcept;
    void FocusMainContentIfPossible() noexcept;

    void QueueOpenCurrentPath() noexcept;
    void QueueTablePreview(std::wstring tableName, uint64_t rowOffset) noexcept;
    void QueueCustomQuery(std::wstring sql) noexcept;

    void StartFileSelection(size_t index) noexcept;
    void StartSelectedTablePreview(bool resetSort) noexcept;

    void LoadOtherFiles(const ViewerOpenContext* context) noexcept;

private:
    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    ViewerSqliteEngine::Config _config{};
    std::string _configurationJson;

    bool _loading  = false;

    wil::com_ptr<IHostAlerts> _hostAlerts;
    wil::com_ptr<IFileSystem> _fileSystem;

    std::wstring _currentPath;
    std::vector<std::wstring> _otherFiles;
    size_t _otherIndex      = 0;
    bool _syncingFileCombo  = false;
    bool _syncingTableCombo = false;

    bool _tablePreviewMode     = true;
    uint64_t _currentRowOffset = 0;
    bool _currentHasMore       = false;
    std::wstring _currentTable;
    std::wstring _currentQuery;
    ViewerSqliteEngine::QueryPage _currentPage;
    std::shared_ptr<ViewerSqliteEngine::DatabaseSource> _databaseSource;
    std::vector<ViewerSqliteEngine::TableInfo> _tables;
    RedSalamander::DxUi::GridSortSpec _tableSortSpec{};

    UINT _dpi = USER_DEFAULT_SCREEN_DPI;

    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::Panel* _root              = nullptr;
    RedSalamander::DxUi::Grid* _resultGrid         = nullptr;
    RedSalamander::DxUi::StatusStrip* _statusStrip = nullptr;
    RedSalamander::DxUi::Label* _fileLabel         = nullptr;
    RedSalamander::DxUi::ComboBox* _fileCombo      = nullptr;
    RedSalamander::DxUi::Button* _reloadButton     = nullptr;
    RedSalamander::DxUi::Label* _tableLabel        = nullptr;
    RedSalamander::DxUi::ComboBox* _tableCombo     = nullptr;
    RedSalamander::DxUi::Button* _prevButton       = nullptr;
    RedSalamander::DxUi::Button* _nextButton       = nullptr;
    RedSalamander::DxUi::Label* _queryLabel        = nullptr;
    RedSalamander::DxUi::TextField* _queryField    = nullptr;
    RedSalamander::DxUi::Button* _runButton        = nullptr;
    RedSalamander::DxUi::Button* _tableButton      = nullptr;
    std::unique_ptr<ViewerSqliteGridModel> _gridModel;

    std::atomic_uint64_t _requestId{0};
#ifdef _DEBUG
    std::shared_ptr<std::atomic_uint32_t> _pendingAsyncWork = std::make_shared<std::atomic_uint32_t>(0u);
#endif
};

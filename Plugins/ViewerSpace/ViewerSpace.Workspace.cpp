#include "ViewerSpace.h"
#include "FileMetadataFormatting.h"
#include "LocalFileTransaction.h"
#include "LocalizationManager.h"
#include "StringConversion.h"
#include "ViewerSpace.Export.h"
#include "resource.h"
#include <shobjidl.h>
#ifdef ENABLE_TESTS
#include "Common.Tests.SandboxPath.h"
#endif

#include <algorithm>
#include <array>
#include <cwctype>
#include <limits>

extern HINSTANCE g_hInstance;

namespace
{
constexpr size_t kMaximumExclusions    = 128u;
constexpr size_t kMaximumPathUnits     = 32768u;
constexpr size_t kMaximumPathDepth     = 512u;
constexpr uint64_t kMaximumExportBytes = 8u * 1024u * 1024u;
std::atomic_bool g_exportInFlight{false};
std::atomic_uint64_t g_exportOwnedBytes{0u};
const uint8_t kExportModuleAnchor = 0u;
struct ExportCompletion final
{
    HRESULT status = E_FAIL;
    std::shared_ptr<const uint8_t> lifetime;
};
} // namespace

bool ViewerSpaceWorkspaceExportsIdle() noexcept
{
    return ! g_exportInFlight.load(std::memory_order_acquire);
}

struct ViewerSpace::WorkspaceExportBuild final
{
    WorkspaceExportBuild()                                       = default;
    WorkspaceExportBuild(const WorkspaceExportBuild&)            = delete;
    WorkspaceExportBuild(WorkspaceExportBuild&&)                 = delete;
    WorkspaceExportBuild& operator=(const WorkspaceExportBuild&) = delete;
    WorkspaceExportBuild& operator=(WorkspaceExportBuild&&)      = delete;
    uint32_t generation                                          = 0u;
    uint64_t revision                                            = 0u;
    uint64_t projectionRevision                                  = 0u;
    uint64_t discovered                                          = 0u;
    uint64_t shown                                               = 0u;
    uint64_t usedBytes                                           = 0u;
    size_t cursor                                                = 0u;
    bool csv                                                     = false;
    bool partial                                                 = false;
    bool countKnown                                              = false;
    bool save                                                    = true;
    std::string area;
    std::string scope;
    std::string root;
    std::string viewPath;
    std::wstring outputPath;
    std::vector<uint32_t> ids;
    std::vector<ViewerSpaceExport::Row> rows;
    HWND owner = nullptr;
    std::shared_ptr<const uint8_t> lifetime;
    wil::unique_hmodule modulePin;
#ifdef ENABLE_TESTS
    wil::unique_event readyEvent;
    wil::unique_event releaseEvent;
    HRESULT publishFailure = S_OK;
#endif
};

void ViewerSpace::ExportCurrentView(bool csv) noexcept
{
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    if (_exportDialogOpen)
        return;
    const bool needsHiddenProjection = ! _config.analysisPanelVisible && _analysisRows.empty() && ! _exportProjectionRequested;
    if (! _model || _scanActive.load() || _analysisDirty || _analysisBuild || _workspaceRebinding ||
        (! needsHiddenProjection && _analysisGeneration != _displayGeneration))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_WAIT);
        return;
    }
    if (_analysisRows.size() > 20000u)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_LIMIT);
        return;
    }
    // A hidden unfiltered table has no projection consumer. Prepare it on the existing
    // sliced path, preserving the user's layout, then let the explicit Export action retry.
    if (! _config.analysisPanelVisible && _analysisRows.empty() && ! _exportProjectionRequested)
    {
        const auto* parent = TryGetRealNode(_model->viewNodeId);
        if (parent && parent->childrenCount > 20000u)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_LIMIT);
            return;
        }
        _exportProjectionRequested = true;
        InvalidateAnalysis();
        return;
    }
    _exportDialogOpen       = true;
    auto restoreDialogState = wil::scope_exit([this]() noexcept
    {
        _exportDialogOpen          = false;
        _exportProjectionRequested = false;
    });
    std::vector<DxUi::MenuFlyoutItem> formats;
    for (UINT index = 0u; index < 2u; ++index)
    {
        DxUi::MenuFlyoutItem item;
        item.commandId = static_cast<int>(index + 1u);
        item.text      = LoadStringResource(g_hInstance, index ? IDS_VIEWERSPACE_EXPORT_CSV : IDS_VIEWERSPACE_EXPORT_JSON);
        formats.push_back(std::move(item));
    }
    POINT point{12, 40};
    ClientToScreen(_hWnd.get(), &point);
    const auto format = DxUi::ContextMenu::Show(_hWnd.get(), point, formats, _analysisHost.GetTheme());
    if (! format.has_value() || ! _hWnd || ! _workspaceLifetime)
        return;
    csv = format.value() == 2;
    wil::com_ptr<IFileSaveDialog> dialog;
    HRESULT status = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(status))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return;
    }
    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_EXPORT);
    const std::wstring type  = LoadStringResource(g_hInstance, csv ? IDS_VIEWERSPACE_EXPORT_CSV : IDS_VIEWERSPACE_EXPORT_JSON);
    const COMDLG_FILTERSPEC filter{type.c_str(), csv ? L"*.csv" : L"*.json"};
    static_cast<void>(dialog->SetTitle(title.c_str()));
    static_cast<void>(dialog->SetOptions(FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT | FOS_NOCHANGEDIR));
    static_cast<void>(dialog->SetFileTypes(1u, &filter));
    static_cast<void>(dialog->SetDefaultExtension(csv ? L"csv" : L"json"));
    static_cast<void>(dialog->SetFileName(csv ? L"ViewerSpace.csv" : L"ViewerSpace.json"));
    status = dialog->Show(_hWnd.get());
    if (status == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        return;
    wil::com_ptr<IShellItem> item;
    wil::unique_cotaskmem_string path;
    if (FAILED(status) || FAILED(dialog->GetResult(item.put())) || ! item || FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, path.put())) || ! path || ! _hWnd ||
        ! _workspaceLifetime)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return;
    }
    static_cast<void>(QueueWorkspaceExport(path.get(), csv));
}

bool ViewerSpace::QueueWorkspaceExport(std::wstring_view outputPath, bool csv) noexcept
{
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    if (! _hWnd || ! _workspaceLifetime || ! _model || outputPath.empty() || outputPath.size() > 32600u || _analysisGeneration != _displayGeneration ||
        _workspaceRebinding)
        return false;
    bool expected = false;
    if (! g_exportInFlight.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_BUSY);
        return false;
    }
    auto releaseAdmission = wil::scope_exit([]() noexcept { g_exportInFlight.store(false, std::memory_order_release); });
    if (_scanActive.load() || _analysisDirty || _analysisBuild || _analysisRows.size() > 20000u)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_WAIT);
        return false;
    }
    if (! CanAdmitWorkspaceBytes(kMaximumExportBytes))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_LIMIT);
        return false;
    }
    auto build                = std::make_shared<WorkspaceExportBuild>();
    build->generation         = _displayGeneration;
    build->revision           = _modelRevision;
    build->projectionRevision = _workspaceProjectionRevision;
    build->discovered         = _scanProgressBytes;
    const auto* view          = TryGetRealNode(_model->viewNodeId);
    build->shown              = view ? ShownBytes(view->id, view->totalBytes) : 0u;
    build->csv                = csv;
    build->partial            = _overallState != ScanState::Done || _scanFailedFolders != 0u;
    build->countKnown         = _workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration;
    build->area               = _config.fileCountWeight ? "file-count" : "logical-bytes";
    build->scope              = _reviewingCleanup ? "cleanup-review" : _showingProblems ? "problems" : "current-folder-filtered";
    const auto root           = Common::Strings::TryUtf8FromUtf16Strict(_model->scanRootPath);
    if (! root.has_value())
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return false;
    }
    const auto viewPath = Common::Strings::TryUtf8FromUtf16Strict(BuildNodePathText(_model->viewNodeId));
    if (! viewPath.has_value())
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return false;
    }
    build->viewPath   = viewPath.value();
    build->root       = root.value();
    build->outputPath = outputPath;
    if (build->outputPath.size() > 32600u)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return false;
    }
    build->ids = _analysisRows;
    build->rows.reserve(build->ids.size());
    build->owner     = _hWnd.get();
    build->lifetime  = _workspaceLifetime;
    build->modulePin = AcquireModuleReferenceFromAddress(&kExportModuleAnchor);
    if (! build->modulePin)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
        return false;
    }
    build->usedBytes = sizeof(WorkspaceExportBuild) + build->ids.capacity() * sizeof(uint32_t) + build->rows.capacity() * sizeof(ViewerSpaceExport::Row) +
                       build->root.capacity() + build->viewPath.capacity() + build->outputPath.capacity() * sizeof(wchar_t);
    g_exportOwnedBytes.store(kMaximumExportBytes, std::memory_order_release);
    _workspaceExport = std::move(build);
    releaseAdmission.release();
    EnsureAnimationTimer();
    return true;
}

#ifdef ENABLE_TESTS
LRESULT ViewerSpace::DebugQueueWorkspaceExport(const WndMsg::ViewerSpaceExportDebugRequest* request) noexcept
{
    if (! request || ! request->outputPath || (! ! request->readyEvent != ! ! request->releaseEvent))
        return FALSE;
    const size_t length = wcsnlen_s(request->outputPath, 32601u);
    if (! length || length > 32600u)
        return FALSE;
    const std::filesystem::path path(std::wstring_view(request->outputPath, length));
    std::error_code error;
    if (! Common::Testing::IsAuthorizedTestSandboxPath(path, {}, false, error))
        return FALSE;
    wil::unique_event ready;
    wil::unique_event release;
    if (request->readyEvent &&
        (! DuplicateHandle(GetCurrentProcess(), request->readyEvent, GetCurrentProcess(), ready.put(), 0u, FALSE, DUPLICATE_SAME_ACCESS) ||
         ! DuplicateHandle(GetCurrentProcess(), request->releaseEvent, GetCurrentProcess(), release.put(), 0u, FALSE, DUPLICATE_SAME_ACCESS)))
        return FALSE;
    if (! QueueWorkspaceExport(path.native(), request->csv))
        return FALSE;
    _workspaceExport->readyEvent     = std::move(ready);
    _workspaceExport->releaseEvent   = std::move(release);
    _workspaceExport->publishFailure = request->publishFailure;
    return TRUE;
}
#endif

void ViewerSpace::ContinueWorkspaceExport() noexcept
{
    if (! _workspaceExport && _exportProjectionRequested && ! _exportDialogOpen && ! _analysisDirty && ! _analysisBuild)
    {
        ExportCurrentView(false);
        return;
    }
    if (! _workspaceExport)
        return;
    if (_workspaceExport->generation != _displayGeneration || _workspaceExport->revision != _modelRevision ||
        _workspaceExport->projectionRevision != _workspaceProjectionRevision)
    {
        SubmitWorkspaceExport(false);
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_STALE);
        return;
    }
    const auto started = std::chrono::steady_clock::now();
    size_t work        = 0u;
    auto& build        = *_workspaceExport;
    while (build.cursor < build.ids.size() && work < 256u && Debug::Perf::ElapsedUs(started) < 1200u)
    {
        const uint32_t id = build.ids[build.cursor++];
        const auto item   = ResolveAnalysisItem(id);
        if (! item.has_value())
        {
            SubmitWorkspaceExport(false);
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_STALE);
            return;
        }
        const auto& value = item.value();
        ViewerSpaceExport::Row row;
        std::wstring syntheticName;
        std::wstring_view name = value.name;
        if (value.isSynthetic)
        {
            syntheticName = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_KIND_OTHER);
            name          = syntheticName;
        }
        const std::wstring path = value.isSynthetic ? std::wstring{} : BuildItemPathText(id);
        auto nameUtf8           = Common::Strings::TryUtf8FromUtf16Strict(name);
        auto pathUtf8           = Common::Strings::TryUtf8FromUtf16Strict(path);
        if (! nameUtf8.has_value() || ! pathUtf8.has_value())
        {
            SubmitWorkspaceExport(false);
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
            return;
        }
        row.name       = std::move(nameUtf8.value());
        row.path       = std::move(pathUtf8.value());
        row.bytes      = value.totalBytes;
        row.shownBytes = ShownBytes(id, value.totalBytes);
        row.countKnown =
            build.countKnown && ! value.isSynthetic && (! value.isDirectory || (id < _workspaceFacts.size() && ! _workspaceFacts[id].countIncomplete));
        row.fileCount = row.countKnown ? ShownFileCount(id) : 0u;
        SYSTEMTIME localModified{};
        if (Common::FileMetadata::TryGetLocalWriteTime(value.lastWriteTime, localModified))
            row.modified = value.lastWriteTime;
        row.directory         = value.isDirectory;
        row.synthetic         = value.isSynthetic;
        row.excluded          = IsItemExcluded(id);
        row.partial           = value.scanState != ScanState::Done ||
                                (value.isDirectory && _workspaceFactsReady && id < _workspaceFacts.size() && _workspaceFacts[id].countIncomplete);
        row.unreadable        = value.scanState == ScanState::Error;
        row.unmeasured        = value.contentNotMeasured;
        const uint64_t charge = row.name.capacity() + row.path.capacity();
        if (charge > kMaximumExportBytes - 1024u * 1024u - std::min(build.usedBytes, kMaximumExportBytes - 1024u * 1024u))
        {
            SubmitWorkspaceExport(false);
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_LIMIT);
            return;
        }
        build.usedBytes += charge;
        build.rows.push_back(std::move(row));
        ++work;
    }
    Debug::Perf::Emit(L"viewer.space.export.prepare_us", L"frozen-retained-rows", Debug::Perf::ElapsedUs(started), work, build.rows.size(), S_OK);
    if (build.cursor == build.ids.size())
        SubmitWorkspaceExport(true);
}

void ViewerSpace::SubmitWorkspaceExport(bool save) noexcept
{
    if (! _workspaceExport)
        return;
    _workspaceExport->save = save;
    auto context           = std::make_unique<std::shared_ptr<WorkspaceExportBuild>>(std::move(_workspaceExport));
    if (TrySubmitThreadpoolCallback(&ViewerSpace::WorkspaceExportCallback, context.get(), nullptr))
    {
        static_cast<void>(context.release());
        return;
    }
    // Failed submission is the sole inline-disposal fallback; no callback owns the snapshot.
    context.reset();
    g_exportOwnedBytes.store(0u, std::memory_order_release);
    g_exportInFlight.store(false, std::memory_order_release);
    if (save)
        ShowWorkspaceMessage(IDS_VIEWERSPACE_EXPORT_IO_ERROR);
    else
        Debug::Warning(L"ViewerSpace export disposal callback could not be submitted");
}

void CALLBACK ViewerSpace::WorkspaceExportCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept
{
    auto releaseAdmission = wil::scope_exit([]() noexcept
    {
        g_exportOwnedBytes.store(0u, std::memory_order_release);
        g_exportInFlight.store(false, std::memory_order_release);
    });
    std::unique_ptr<std::shared_ptr<WorkspaceExportBuild>> owner(static_cast<std::shared_ptr<WorkspaceExportBuild>*>(context));
    auto job = std::move(*owner);
    TransferModulePinToCallbackReturn(instance, job->modulePin);
    if (! job->save)
        return;
#ifdef ENABLE_TESTS
    if (job->readyEvent)
    {
        SetEvent(job->readyEvent.get());
        if (WaitForSingleObject(job->releaseEvent.get(), 5000u) != WAIT_OBJECT_0)
        {
            Debug::Error(L"ViewerSpace test export gate timed out");
            return;
        }
    }
#endif
    HRESULT status = E_FAIL;
    try
    {
        Common::Files::LocalFileTransaction transaction;
        status = Common::Files::LocalFileTransaction::Create(
            std::filesystem::path(job->outputPath), Common::Files::ExistingTargetPolicy::Replace, false, transaction);
        if (SUCCEEDED(status))
        {
            uint64_t written = 0u;
            const auto write = [&](std::string_view bytes) noexcept
            {
                if (FAILED(status))
                    return;
                if (bytes.size() > kMaximumExportBytes - std::min(written, kMaximumExportBytes))
                {
                    status = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
                    return;
                }
                status = transaction.Write(bytes);
                if (SUCCEEDED(status))
                    written += bytes.size();
            };
            if (job->csv)
            {
                write("\xEF\xBB\xBF");
                write(
                    "recordType,scanRoot,viewPath,discoveredBytes,viewShownBytes,scanPartial,area,scope,name,path,directory,synthetic,logicalBytes,shownBytes,"
                    "fileCount,modified100ns,excluded,partial,unreadable,unmeasured,nameFormulaEscaped,pathFormulaEscaped,allocationBytes\r\n");
                write(std::format("metadata,{},{},{},{},{},{},{}{}\r\n",
                                  ViewerSpaceExport::CsvText(job->root),
                                  ViewerSpaceExport::CsvText(job->viewPath),
                                  job->discovered,
                                  job->shown,
                                  job->partial,
                                  ViewerSpaceExport::CsvText(job->area),
                                  ViewerSpaceExport::CsvText(job->scope),
                                  std::string(15u, ',')));
                for (const auto& row : job->rows)
                {
                    write("item,,,,,,,,");
                    write(ViewerSpaceExport::CsvRow(row));
                }
            }
            else
            {
                Common::Json::UniqueMutableDocument document{yyjson_mut_doc_new(nullptr)};
                if (! document)
                    std::terminate();
                auto* root = yyjson_mut_obj(document.get());
                yyjson_mut_doc_set_root(document.get(), root);
                yyjson_mut_obj_add_str(document.get(), root, "format", "ViewerSpace.current-view");
                yyjson_mut_obj_add_uint(document.get(), root, "version", 1u);
                yyjson_mut_obj_add_strcpy(document.get(), root, "scanRoot", job->root.c_str());
                yyjson_mut_obj_add_uint(document.get(), root, "discoveredBytes", job->discovered);
                yyjson_mut_obj_add_uint(document.get(), root, "viewShownBytes", job->shown);
                yyjson_mut_obj_add_bool(document.get(), root, "partial", job->partial);
                yyjson_mut_obj_add_strcpy(document.get(), root, "area", job->area.c_str());
                yyjson_mut_obj_add_strcpy(document.get(), root, "scope", job->scope.c_str());
                yyjson_mut_obj_add_strcpy(document.get(), root, "viewPath", job->viewPath.c_str());
                size_t length = 0u;
                Common::Json::UniqueMallocString text{yyjson_mut_write(document.get(), YYJSON_WRITE_NOFLAG, &length)};
                if (! text || ! length)
                    std::terminate();
                write(std::string_view(text.get(), length - 1u));
                write(",\"rows\":[");
                bool first = true;
                for (const auto& row : job->rows)
                {
                    if (! first)
                        write(",");
                    first = false;
                    write(ViewerSpaceExport::JsonRow(row));
                }
                write("]}\n");
            }
            if (SUCCEEDED(status))
            {
#ifdef ENABLE_TESTS
                if (FAILED(job->publishFailure))
                    Common::Files::Testing::FailNextLocalFileTransactionPublish(job->publishFailure);
#endif
                status = transaction.Commit(written);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // OS thread-pool ABI boundary: keep export failure off the UI thread and report once.
        Debug::Error(L"ViewerSpace export failed while formatting its frozen snapshot.");
        status = E_FAIL;
    }
    auto completion      = std::make_unique<ExportCompletion>();
    completion->status   = status;
    completion->lifetime = job->lifetime;
    static_cast<void>(PostMessagePayload(job->owner, kWorkspaceExportCompleted, 0u, std::move(completion)));
}

void ViewerSpace::HandleWorkspaceExportCompletion(HWND hwnd, LPARAM token) noexcept
{
    const auto completion = TakeMessagePayload<ExportCompletion>(hwnd, kWorkspaceExportCompleted, token);
    if (! completion || completion->lifetime != _workspaceLifetime || ! _hWnd)
        return;
    if (FAILED(completion->status))
        Debug::Warning(L"ViewerSpace export could not publish its destination: 0x{:08X}", static_cast<uint32_t>(completion->status));
    ShowWorkspaceMessage(static_cast<UINT>(SUCCEEDED(completion->status)                                    ? IDS_VIEWERSPACE_EXPORT_DONE
                                           : completion->status == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE) ? IDS_VIEWERSPACE_EXPORT_LIMIT
                                                                                                            : IDS_VIEWERSPACE_EXPORT_IO_ERROR));
}

uint32_t ViewerSpace::ItemCategory(const ResolvedItem& item) const noexcept
{
    if (item.isSynthetic)
        return 6u;
    if (item.isDirectory)
    {
        if (! _workspaceFactsReady || _workspaceFactsGeneration != _displayGeneration || item.id >= _workspaceFacts.size())
            return 6u;
        const auto& facts = _workspaceFacts[item.id];
        if (facts.unknownTypes || ! facts.types || (facts.types & (facts.types - 1u)))
            return 6u;
        for (uint32_t category = 0u; category < 7u; ++category)
            if (facts.types & (1u << category))
                return category;
        return 6u;
    }
    const size_t dot = item.name.find_last_of(L'.');
    if (dot == std::wstring_view::npos || item.name.size() - dot > 16u)
        return 6u;
    std::array<wchar_t, 17u> extension{};
    size_t length = 0u;
    for (wchar_t ch : item.name.substr(dot))
        extension[length++] = static_cast<wchar_t>(std::towlower(ch));
    const std::wstring_view ext(extension.data(), length);
    constexpr std::array<std::wstring_view, 6u> groups{
        {L".obj .o .pdb .ilk .pch .lib .exe .dll .class .pyc",
         L".zip .7z .rar .tar .gz .xz .bz2",
         L".jpg .jpeg .png .gif .webp .bmp .svg .mp4 .mov .mkv .mp3 .wav .flac",
         L".db .sqlite .sqlite3 .ndjson .csv .parquet",
         L".cpp .c .h .hpp .cs .rs .py .js .ts .tsx .jsx .java .go .html .css .json .json5 .xml .yaml .yml .toml",
         L".txt .md .pdf .doc .docx .xls .xlsx .ppt .pptx .rtf"}};
    for (uint32_t category = 0u; category < groups.size(); ++category)
    {
        const auto group = groups[category];
        size_t start     = 0u;
        while (start < group.size())
        {
            const size_t end = group.find(L' ', start);
            if (group.substr(start, end == std::wstring_view::npos ? end : end - start) == ext)
                return category;
            if (end == std::wstring_view::npos)
                break;
            start = end + 1u;
        }
    }
    return 6u;
}

D2D1_COLOR_F ViewerSpace::CategoryColor(uint32_t category) const noexcept
{
    constexpr std::array<uint32_t, 7u> colors{{0x8461C5u, 0xC98930u, 0xBC5894u, 0x30958Eu, 0x3B82C4u, 0x59984Au, 0x858B95u}};
    return D2D1::ColorF(colors[std::min(category, 6u)]);
}

int64_t ViewerSpace::ItemAgeDate(const ResolvedItem& item) const noexcept
{
    if (item.isSynthetic)
        return 0;
    if (_config.colorMode == 3u && item.isDirectory)
    {
        if (! _workspaceFactsReady || _workspaceFactsGeneration != _displayGeneration || item.id >= _workspaceFacts.size())
            return 0;
        return _workspaceFacts[item.id].unknown ? 0 : _workspaceFacts[item.id].newest;
    }
    return item.lastWriteTime;
}

bool ViewerSpace::MatchesWorkspaceFilter(const ResolvedItem& item) const noexcept
{
    if (_typeFilter)
    {
        if (item.isDirectory)
        {
            if (_workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration && item.id < _workspaceFacts.size() &&
                ! _workspaceFacts[item.id].unknownTypes && ! (_workspaceFacts[item.id].types & (1u << (_typeFilter - 1u))))
                return false;
        }
        else if (ItemCategory(item) + 1u != _typeFilter)
            return false;
    }
    if (_ageFilter)
    {
        const int64_t date    = ItemAgeDate(item);
        constexpr int64_t day = 86400ll * 10000000ll;
        if (date <= 0 || date > _workspaceNow || _workspaceNow - date < (_ageFilter == 1u ? 30ll : 365ll) * day)
            return false;
    }
    return true;
}

uint64_t ViewerSpace::ShownFileCount(uint32_t id) const noexcept
{
    const auto count = [this](uint32_t target) noexcept -> uint64_t
    {
        if (const auto* file = TryGetFileRecord(target); file)
            return 1u;
        return _workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration && target < _workspaceFacts.size() ? _workspaceFacts[target].files : 0u;
    };
    uint64_t result = count(id);
    if (_workspaceViewExcluded || _excludedRoots.contains(id))
        return 0u;
    if (const auto found = _excludedBelow.find(id); found != _excludedBelow.end())
        for (uint32_t root : found->second)
            result -= std::min(result, count(root));
    return result;
}

uint64_t ViewerSpace::MapWeight(uint32_t id, uint64_t logicalBytes) const noexcept
{
    return _config.fileCountWeight ? (_workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration ? ShownFileCount(id) : 0u)
                                   : ShownBytes(id, logicalBytes);
}

void ViewerSpace::ContinueWorkspaceFacts() noexcept
{
    if (! _model)
        return;
    if (_workspaceFactsGeneration != _displayGeneration)
    {
        _workspaceFactsGeneration = _displayGeneration;
        _workspaceFactsRevision   = (std::numeric_limits<uint64_t>::max)();
        _workspaceFactsReady      = false;
        _workspaceFactsBuilding   = false;
        _problemRows.clear();
        _problemCursor = 0u;
        UpdateWorkspaceScope();
    }
    if (_scanActive.load() || _workingModel || _scanCacheRestoreSnapshot)
        return;
    constexpr size_t maximumNodes = 65536u;
    if (_model->nodes.size() > maximumNodes)
    {
        _workspaceFactsGeneration = _displayGeneration;
        _workspaceFactsReady      = false;
        if (_workspaceFactsRevision != _modelRevision)
        {
            _workspaceFactsRevision = _modelRevision;
            _problemRows.clear();
            _problemCursor = 0u;
        }
        if (_problemCursor >= _model->nodes.size())
        {
            _workspaceFactsBuilding = false;
            return;
        }
        const auto started = std::chrono::steady_clock::now();
        while (_problemCursor < _model->nodes.size() && Debug::Perf::ElapsedUs(started) < 1200u)
        {
            const auto& node = _model->nodes[_problemCursor++];
            if (node.id && (node.scanState == ScanState::Error || node.contentNotMeasured || node.aggregateFiles || node.aggregateFolders) &&
                _problemRows.size() < 256u)
                _problemRows.push_back(node.id);
        }
        _workspaceFactsBuilding = _problemCursor < _model->nodes.size();
        if (_workspaceFactsBuilding)
            EnsureAnimationTimer();
        else if (_showingProblems)
        {
            UpdateWorkspaceScope();
            InvalidateAnalysis();
        }
        return;
    } // Optional derived detail stays within its explicit 3 MiB scratch+published cap.
    if (_workspaceFactsGeneration != _displayGeneration || _workspaceFactsRevision != _modelRevision)
    {
        _workspaceFactsGeneration = _displayGeneration;
        _workspaceFactsRevision   = _modelRevision;
        _workspaceFactsReady      = false;
        _workspaceFactsBuilding   = true;
        _workspaceFactsNext.clear();
        const uint64_t required =
            _model->nodes.size() > _workspaceFactsNext.capacity() ? (_model->nodes.size() - _workspaceFactsNext.capacity()) * sizeof(WorkspaceFacts) : 0u;
        if (! CanAdmitWorkspaceBytes(required))
        {
            _workspaceFactsBuilding = false;
            UpdateWorkspaceScope();
            return;
        }
        _workspaceFactsNext.reserve(_model->nodes.size());
        _workspaceFactsFileCursor = 0u;
        _workspaceFactsNodeCursor = _model->nodes.size();
        _problemRows.clear();
    }
    if (! _workspaceFactsBuilding)
        return;
    const auto started = std::chrono::steady_clock::now();
    size_t work        = 0u;
    while (work < 32768u && Debug::Perf::ElapsedUs(started) < 1200u)
    {
        if (_workspaceFactsNext.size() < _model->nodes.size())
        {
            const size_t size = std::min(_model->nodes.size(), _workspaceFactsNext.size() + 1024u);
            _workspaceFactsNext.resize(size);
            work += 1024u;
        }
        else if (_workspaceFactsFileCursor < _model->fileRecords.size())
        {
            const auto& file = _model->fileRecords[_workspaceFactsFileCursor++];
            ++work;
            if (! file.id || file.parentId >= _workspaceFactsNext.size())
                continue;
            auto& parent = _workspaceFactsNext[file.parentId];
            ++parent.files;
            parent.newest   = std::max(parent.newest, file.lastWriteTime);
            parent.unknown  = parent.unknown || file.lastWriteTime <= 0;
            const auto item = ResolveAnalysisItem(file.id);
            if (item.has_value())
                parent.types |= static_cast<uint8_t>(1u << ItemCategory(item.value()));
        }
        else if (_workspaceFactsNodeCursor)
        {
            const size_t index = --_workspaceFactsNodeCursor;
            ++work;
            const auto& node = _model->nodes[index];
            if (! node.id || node.id != index)
                continue;
            auto& facts = _workspaceFactsNext[index];
            if (! node.isDirectory && ! node.isSynthetic)
            {
                facts.files     = 1u;
                facts.newest    = node.lastWriteTime;
                const auto item = ResolveAnalysisItem(node.id);
                if (item.has_value())
                    facts.types = static_cast<uint8_t>(1u << ItemCategory(item.value()));
            }
            facts.files           = ViewerSpaceScan::SaturatingAdd(facts.files, node.aggregateFiles);
            facts.countIncomplete = facts.countIncomplete || node.aggregateFolders || node.contentNotMeasured || node.scanState != ScanState::Done;
            facts.unknownTypes    = facts.unknownTypes || node.aggregateFiles || facts.countIncomplete;
            facts.unknown         = facts.unknown || node.aggregateFiles || facts.countIncomplete;
            if ((node.scanState == ScanState::Error || node.contentNotMeasured || node.aggregateFiles || node.aggregateFolders) && _problemRows.size() < 256u)
                _problemRows.push_back(node.id);
            if (node.parentId && node.parentId < index)
            {
                auto& parent  = _workspaceFactsNext[node.parentId];
                parent.files  = ViewerSpaceScan::SaturatingAdd(parent.files, facts.files);
                parent.newest = std::max(parent.newest, facts.newest);
                parent.types |= facts.types;
                parent.unknown         = parent.unknown || facts.unknown;
                parent.unknownTypes    = parent.unknownTypes || facts.unknownTypes;
                parent.countIncomplete = parent.countIncomplete || facts.countIncomplete;
            }
        }
        else
        {
            _workspaceFacts.swap(_workspaceFactsNext);
            _workspaceFactsReady    = true;
            _workspaceFactsBuilding = false;
            UpdateWorkspaceScope();
            CancelLayoutBuild();
            _layoutDirty = true;
            InvalidateAnalysis();
            InvalidateStaticTreemapCache();
            break;
        }
    }
    if (_workspaceFactsBuilding)
        EnsureAnimationTimer();
    Debug::Perf::Emit(
        L"viewer.space.workspace.facts_us", L"bounded-retained-metadata", Debug::Perf::ElapsedUs(started), work, _workspaceFactsNext.size(), S_OK);
}

void ViewerSpace::ShowScanProblems() noexcept
{
    _showingProblems  = ! _showingProblems;
    _reviewingCleanup = false;
    if (! _config.analysisPanelVisible)
        ToggleAnalysisPanel();
    UpdateWorkspaceScope();
    InvalidateAnalysis();
}

void ViewerSpace::ShowWorkspaceFilters(bool breakdown) noexcept
{
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    if (_reviewingCleanup || _showingProblems)
        return;
    if (breakdown && (_analysisGeneration != _displayGeneration || _analysisDirty || _analysisBuild))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_FACTS_PENDING);
        return;
    }
    constexpr std::array<UINT, 8u> captions{{IDS_VIEWERSPACE_TYPE_ALL,
                                             IDS_VIEWERSPACE_TYPE_BUILD,
                                             IDS_VIEWERSPACE_TYPE_ARCHIVES,
                                             IDS_VIEWERSPACE_TYPE_MEDIA,
                                             IDS_VIEWERSPACE_TYPE_DATA,
                                             IDS_VIEWERSPACE_TYPE_CODE,
                                             IDS_VIEWERSPACE_TYPE_DOCUMENTS,
                                             IDS_VIEWERSPACE_TYPE_OTHER}};
    std::vector<DxUi::MenuFlyoutItem> items;
    DxUi::MenuFlyoutItem heading;
    heading.kind = DxUi::MenuItemKind::Info;
    heading.text = LoadStringResource(g_hInstance, breakdown ? IDS_VIEWERSPACE_BREAKDOWN_SCOPE : IDS_VIEWERSPACE_ANALYSIS_FILTER_SCOPE);
    items.push_back(std::move(heading));
    for (uint32_t index = 0u; index < captions.size(); ++index)
    {
        DxUi::MenuFlyoutItem item;
        item.kind       = DxUi::MenuItemKind::Radio;
        item.commandId  = static_cast<int>(IDM_VIEWERSPACE_TYPE_ALL + index);
        item.checked    = _typeFilter == index;
        const auto name = LoadStringResource(g_hInstance, captions[index]);
        item.text =
            breakdown && index
                ? FormatStringResource(
                      g_hInstance, IDS_VIEWERSPACE_BREAKDOWN_ROW, name, FormatBytesCompact(_typeBreakdownBytes[index - 1u]), _typeBreakdownCounts[index - 1u])
                : name;
        items.push_back(std::move(item));
    }
    if (breakdown)
    {
        DxUi::MenuFlyoutItem folder;
        folder.kind = DxUi::MenuItemKind::Info;
        folder.text = FormatStringResource(g_hInstance,
                                           IDS_VIEWERSPACE_BREAKDOWN_ROW,
                                           LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_KIND_FOLDER),
                                           FormatBytesCompact(_typeBreakdownBytes[7u]),
                                           _typeBreakdownCounts[7u]);
        items.push_back(std::move(folder));
    }
    DxUi::MenuFlyoutItem separator;
    separator.kind = DxUi::MenuItemKind::Separator;
    items.push_back(std::move(separator));
    DxUi::MenuFlyoutItem dateBasis;
    dateBasis.kind = DxUi::MenuItemKind::Header;
    dateBasis.text = LoadStringResource(g_hInstance, _config.colorMode == 3u ? IDS_VIEWERSPACE_COLOR_NEWEST : IDS_VIEWERSPACE_COLOR_AGE);
    items.push_back(std::move(dateBasis));
    constexpr std::array<UINT, 3u> ages{{IDS_VIEWERSPACE_AGE_ANY, IDS_VIEWERSPACE_AGE_MONTH, IDS_VIEWERSPACE_AGE_YEAR}};
    for (uint32_t index = 0u; index < ages.size(); ++index)
    {
        DxUi::MenuFlyoutItem item;
        item.kind      = DxUi::MenuItemKind::Radio;
        item.commandId = static_cast<int>(IDM_VIEWERSPACE_AGE_ANY + index);
        item.checked   = _ageFilter == index;
        item.text      = LoadStringResource(g_hInstance, ages[index]);
        items.push_back(std::move(item));
    }
    POINT point{12, 42};
    const HWND anchor = _analysisHost.GetHwnd() ? _analysisHost.GetHwnd() : _hWnd.get();
    ClientToScreen(anchor, &point);
    const auto chosen = DxUi::ContextMenu::Show(_hWnd.get(), point, items, _analysisHost.GetTheme());
    if (chosen.has_value() && _hWnd)
        OnCommand(_hWnd.get(), static_cast<UINT>(chosen.value()));
}

void ViewerSpace::UpdateWorkspaceScope() noexcept
{
    FILETIME now{};
    GetSystemTimeAsFileTime(&now);
    _workspaceNow        = static_cast<int64_t>((static_cast<uint64_t>(now.dwHighDateTime) << 32u) | now.dwLowDateTime);
    uint64_t markedBytes = 0u;
    bool incomplete      = _scanFailedFolders != 0u || _overallState != ScanState::Done;
    for (uint32_t id : _cleanupMarks)
    {
        const auto item = ResolveAnalysisItem(id);
        incomplete      = incomplete || ! item.has_value() || item.value().contentNotMeasured || item.value().scanState != ScanState::Done;
        if (item.has_value())
            markedBytes = ViewerSpaceScan::SaturatingAdd(markedBytes, item.value().totalBytes);
    }
    const std::array<UINT, 4u> colors{{IDS_VIEWERSPACE_COLOR_FOLDER, IDS_VIEWERSPACE_COLOR_TYPE, IDS_VIEWERSPACE_COLOR_AGE, IDS_VIEWERSPACE_COLOR_NEWEST}};
    std::wstring scope =
        _reviewingCleanup
            ? FormatStringResource(g_hInstance,
                                   IDS_VIEWERSPACE_REVIEW_MESSAGE,
                                   _cleanupMarks.size(),
                                   incomplete ? FormatStringResource(g_hInstance, IDS_VIEWERSPACE_SIZE_KNOWN_FORMAT, FormatBytesCompact(markedBytes))
                                              : FormatBytesCompact(markedBytes))
        : _showingProblems
            ? FormatStringResource(g_hInstance, IDS_VIEWERSPACE_PROBLEMS_SCOPE, _problemRows.size())
            : FormatStringResource(g_hInstance,
                                   IDS_VIEWERSPACE_VIEW_FACTS,
                                   LoadStringResource(g_hInstance, _config.fileCountWeight ? IDS_VIEWERSPACE_WEIGHT_COUNT : IDS_VIEWERSPACE_WEIGHT_SIZE),
                                   LoadStringResource(g_hInstance, colors[_config.colorMode]));
    if (! _reviewingCleanup && ! _showingProblems)
    {
        if (_typeFilter || _ageFilter)
            scope.append(L" · ").append(LoadStringResource(g_hInstance, IDS_VIEWERSPACE_FILTER_HIGHLIGHT));
        if (_ageFilter)
            scope.append(L" · ")
                .append(LoadStringResource(g_hInstance, _config.colorMode == 3u ? IDS_VIEWERSPACE_COLOR_NEWEST : IDS_VIEWERSPACE_COLOR_AGE))
                .append(L" · ")
                .append(LoadStringResource(g_hInstance, _ageFilter == 1u ? IDS_VIEWERSPACE_AGE_MONTH : IDS_VIEWERSPACE_AGE_YEAR));
        if ((_config.fileCountWeight || _config.colorMode == 3u) && ! _workspaceFactsReady)
            scope.append(L" · ").append(LoadStringResource(g_hInstance, IDS_VIEWERSPACE_FACTS_PENDING));
    }
    _analysisHost.SetWorkspaceScope(_reviewingCleanup, _showingProblems, scope);
    std::vector<ViewerSpaceAnalysis::AnalysisHost::LegendEntry> legend;
    if (! _reviewingCleanup && ! _showingProblems && _config.colorMode == 1u)
    {
        constexpr std::array<UINT, 7u> names{{IDS_VIEWERSPACE_TYPE_BUILD,
                                              IDS_VIEWERSPACE_TYPE_ARCHIVES,
                                              IDS_VIEWERSPACE_TYPE_MEDIA,
                                              IDS_VIEWERSPACE_TYPE_DATA,
                                              IDS_VIEWERSPACE_TYPE_CODE,
                                              IDS_VIEWERSPACE_TYPE_DOCUMENTS,
                                              IDS_VIEWERSPACE_TYPE_OTHER}};
        for (uint32_t index = 0u; index < names.size(); ++index)
            legend.push_back({LoadStringResource(g_hInstance, names[index]), CategoryColor(index)});
    }
    else if (! _reviewingCleanup && ! _showingProblems && _config.colorMode >= 2u)
    {
        constexpr std::array<UINT, 5u> names{{IDS_VIEWERSPACE_AGE_RECENT,
                                              IDS_VIEWERSPACE_AGE_MONTH_LEGEND,
                                              IDS_VIEWERSPACE_AGE_YEAR_LEGEND,
                                              IDS_VIEWERSPACE_AGE_OLD,
                                              IDS_VIEWERSPACE_ALLOCATION_UNAVAILABLE}};
        constexpr std::array<uint32_t, 5u> values{{0x4AA67Bu, 0xBDAC3Cu, 0xC98439u, 0xBA5961u, 0x858B95u}};
        for (size_t index = 0u; index < names.size(); ++index)
            legend.push_back({LoadStringResource(g_hInstance, names[index]), D2D1::ColorF(values[index])});
    }
    _analysisHost.SetLegend(legend);
    if (_hWnd)
        UpdateMenuState(_hWnd.get());
}

void ViewerSpace::ToggleCleanupMark(uint32_t id, bool notify) noexcept
{
    const auto target = ResolveAnalysisItem(id);
    if (! target.has_value() || target.value().isSynthetic)
        return;
    if (_cleanupGeneration != _displayGeneration)
    {
        _cleanupMarks.clear();
        _cleanupGeneration = _displayGeneration;
    }
    const auto found = std::find(_cleanupMarks.begin(), _cleanupMarks.end(), id);
    if (found != _cleanupMarks.end())
        _cleanupMarks.erase(found);
    else
    {
        if (IsItemExcluded(id))
            return;
        uint32_t parent = target.value().parentId;
        for (size_t depth = 0u; parent && depth < kMaximumPathDepth; ++depth)
        {
            if (std::find(_cleanupMarks.begin(), _cleanupMarks.end(), parent) != _cleanupMarks.end())
                return;
            const auto ancestor = ResolveAnalysisItem(parent);
            parent              = ancestor.has_value() ? ancestor.value().parentId : 0u;
        }
        if (parent)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_CLEANUP_LIMIT);
            return;
        }
        std::erase_if(_cleanupMarks,
                      [&](uint32_t mark)
        {
            for (size_t depth = 0u; mark && depth < kMaximumPathDepth; ++depth)
            {
                if (mark == id)
                    return true;
                const auto ancestor = ResolveAnalysisItem(mark);
                mark                = ancestor.has_value() ? ancestor.value().parentId : 0u;
            }
            return false;
        });
        if (_cleanupMarks.size() >= kMaximumExclusions)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_CLEANUP_LIMIT);
            return;
        }
        _cleanupMarks.push_back(id);
    }
    if (notify)
    {
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        _analysisHost.NotifyRowsChanged(std::nullopt);
    }
}

void ViewerSpace::MarkMatchingItems() noexcept
{
    if (_reviewingCleanup || _showingProblems || (_model && IsItemExcluded(_model->viewNodeId)) || _analysisGeneration != _displayGeneration ||
        _analysisDirty || _analysisBuild)
        return;
    std::vector<uint32_t> additions;
    for (uint32_t id : _analysisRows)
    {
        const auto item = ResolveAnalysisItem(id);
        if (! item.has_value() || item.value().isSynthetic || IsItemExcluded(id) ||
            std::find(_cleanupMarks.begin(), _cleanupMarks.end(), id) != _cleanupMarks.end())
            continue;
        additions.push_back(id);
        if (additions.size() + _cleanupMarks.size() > kMaximumExclusions)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_CLEANUP_LIMIT);
            return;
        }
    }
    for (uint32_t id : additions)
        ToggleCleanupMark(id, false);
    UpdateWorkspaceScope();
    InvalidateAnalysis();
}

void ViewerSpace::ReviewCleanup() noexcept
{
    if (_showingProblems)
    {
        _showingProblems = false;
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        return;
    }
    if (! _reviewingCleanup && (_cleanupMarks.empty() || _cleanupGeneration != _displayGeneration))
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_REVIEW_EMPTY);
        return;
    }
    _reviewingCleanup = ! _reviewingCleanup;
    _showingProblems  = false;
    if (! _config.analysisPanelVisible)
        ToggleAnalysisPanel();
    UpdateWorkspaceScope();
    InvalidateAnalysis();
}

void ViewerSpace::SendReviewedCleanup(bool recycle) noexcept
{
    if (! _reviewingCleanup || _cleanupGeneration != _displayGeneration || _cleanupMarks.empty() || ! _hostViewerActions)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_REVIEW_STALE);
        return;
    }
    std::vector<std::wstring> paths;
    paths.reserve(_cleanupMarks.size());
    size_t pathUnits = 0u;
    for (uint32_t id : _cleanupMarks)
    {
        const auto item = ResolveAnalysisItem(id);
        if (! item.has_value() || item.value().isSynthetic || IsItemExcluded(id))
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_REVIEW_STALE);
            return;
        }
        paths.push_back(BuildItemPathText(id));
        pathUnits += paths.back().size();
        if (paths.back().empty() || paths.back().size() > kMaximumPathUnits || pathUnits > 131072u)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_CLEANUP_LIMIT);
            return;
        }
    }
    std::vector<const wchar_t*> raw;
    raw.reserve(paths.size());
    for (const auto& path : paths)
        raw.push_back(path.c_str());
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    HostViewerActionRequest request{};
    request.sizeBytes    = sizeof(request);
    request.action       = recycle ? HOST_VIEWER_RECYCLE : HOST_VIEWER_DELETE;
    request.fileSystem   = _fileSystem.get();
    request.paths        = raw.data();
    request.pathCount    = static_cast<uint32_t>(raw.size());
    uint64_t taskId      = 0u;
    const HRESULT result = _hostViewerActions->ExecuteViewerAction(&request, &taskId);
    if (result == S_OK && taskId)
    {
        _cleanupMarks.clear();
        _reviewingCleanup = false;
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        ShowWorkspaceMessage(IDS_VIEWERSPACE_CLEANUP_SENT);
    }
    else if (FAILED(result))
        ShowWorkspaceMessage(IDS_VIEWERSPACE_TARGET_UNAVAILABLE);
}

void ViewerSpace::OpenFolderInOtherPane() noexcept
{
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    const auto item                 = _selectedGeneration == _displayGeneration ? ResolveAnalysisItem(_selectedItemId) : std::nullopt;
    if (! item.has_value() || item.value().isSynthetic || ! item.value().isDirectory || ! _hostViewerActions)
        return;
    const std::wstring path = BuildItemPathText(item.value().id);
    const wchar_t* raw      = path.c_str();
    HostViewerActionRequest request{};
    request.sizeBytes  = sizeof(request);
    request.action     = HOST_VIEWER_OPEN_OTHER_PANE;
    request.fileSystem = _fileSystem.get();
    request.paths      = &raw;
    request.pathCount  = 1u;
    if (FAILED(_hostViewerActions->ExecuteViewerAction(&request, nullptr)))
        ShowWorkspaceMessage(IDS_VIEWERSPACE_TARGET_UNAVAILABLE);
}

void ViewerSpace::ShowWorkspaceMessage(UINT stringId) noexcept
{
    if (! _hostAlerts)
        return;
    const std::wstring message = LoadStringResource(g_hInstance, stringId);
    HostAlertRequest request{};
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = HOST_ALERT_INFO;
    request.targetWindow = _hWnd.get();
    request.message      = message.c_str();
    request.closable     = TRUE;
    static_cast<void>(_hostAlerts->ShowAlert(&request, nullptr));
}

bool ViewerSpace::IsItemExcluded(uint32_t id) const noexcept
{
    if (_workspaceGeneration != _displayGeneration || _excludedRoots.empty())
        return false;
    for (size_t depth = 0u; id && depth < kMaximumPathDepth; ++depth)
    {
        if (_excludedRoots.contains(id))
            return true;
        const auto item = ResolveAnalysisItem(id);
        if (! item.has_value())
            return false;
        id = item.value().parentId;
    }
    return false;
}

uint64_t ViewerSpace::ShownBytes(uint32_t id, uint64_t original) const noexcept
{
    if (_workspaceGeneration != _displayGeneration)
        return original;
    if (_workspaceViewExcluded || _excludedRoots.contains(id))
        return 0u;
    const auto found = _excludedBelow.find(id);
    if (found == _excludedBelow.end())
        return original;
    uint64_t excluded = 0u;
    // Stored ancestor membership avoids walking paths during paint. Read the live original
    // totals, so ongoing scans update both projections without modifying authoritative data.
    for (uint32_t root : found->second)
    {
        const auto item = ResolveAnalysisItem(root);
        if (item.has_value())
            excluded += std::min(original - excluded, item.value().totalBytes);
    }
    return original - excluded;
}

uint64_t ViewerSpace::WorkspaceOwnedBytes() const noexcept
{
    uint64_t result = g_exportOwnedBytes.load(std::memory_order_acquire) + _excludedPaths.capacity() * sizeof(ExcludedPath) +
                      _workspaceBindingParts.capacity() * sizeof(std::wstring);
    for (const auto& entry : _excludedPaths)
        result += entry.path.capacity() * sizeof(wchar_t);
    for (const auto& part : _workspaceBindingParts)
        result += part.capacity() * sizeof(wchar_t);
    result += _excludedBelow.bucket_count() * sizeof(void*) + _excludedRoots.bucket_count() * sizeof(void*);
    for (const auto& [id, entries] : _excludedBelow)
    {
        static_cast<void>(id);
        result += sizeof(std::pair<const uint32_t, std::vector<uint32_t>>) + 2u * sizeof(void*) + entries.capacity() * sizeof(uint32_t);
    }
    result += (_workspaceFacts.capacity() + _workspaceFactsNext.capacity()) * sizeof(WorkspaceFacts) + _cleanupMarks.capacity() * sizeof(uint32_t) +
              _problemRows.capacity() * sizeof(uint32_t);
    result += _excludedRoots.size() * (sizeof(uint32_t) + 2u * sizeof(void*));
    result += _workspaceNextExcludedBelow.bucket_count() * sizeof(void*) + _workspaceNextExcludedRoots.bucket_count() * sizeof(void*);
    for (const auto& [id, entries] : _workspaceNextExcludedBelow)
    {
        static_cast<void>(id);
        result += sizeof(std::pair<const uint32_t, std::vector<uint32_t>>) + 2u * sizeof(void*) + entries.capacity() * sizeof(uint32_t);
    }
    result += _workspaceNextExcludedRoots.size() * (sizeof(uint32_t) + 2u * sizeof(void*));
    return result;
}

void ViewerSpace::WorkspaceChanged() noexcept
{
    _workspaceViewExcluded = _model && IsItemExcluded(_model->viewNodeId);
    ++_workspaceProjectionRevision;
    CancelLayoutBuild();
    _layoutDirty = true;
    InvalidateAnalysis();
    InvalidateStaticTreemapCache();
    UpdateHeaderTextCache();
    UpdateInspector();
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::ToggleExcludedItem(uint32_t id) noexcept
{
    const auto item = ResolveAnalysisItem(id);
    if (! item.has_value() || item.value().isSynthetic || ! _model || _workspaceRebinding)
        return;
    if (! _model->workspacePathIdentity.has_value())
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_TARGET_UNAVAILABLE);
        return;
    }
    const std::wstring path = BuildItemPathText(id);
    if (path.empty() || path.size() > kMaximumPathUnits)
        return;
    const auto existing = std::find_if(_excludedPaths.begin(), _excludedPaths.end(), [&](const ExcludedPath& entry) {
        return EquivalentPath(_model->workspacePathIdentity.value(), entry.path, path);
    });
    if (existing != _excludedPaths.end())
        _excludedPaths.erase(existing);
    else
    {
        size_t units = path.size();
        for (const auto& entry : _excludedPaths)
            units += entry.path.size();
        size_t ancestors = 0u;
        uint32_t cursor  = id;
        for (; cursor && ancestors < kMaximumPathDepth; ++ancestors)
        {
            const auto ancestor = ResolveAnalysisItem(cursor);
            cursor              = ancestor.has_value() ? ancestor.value().parentId : 0u;
        }
        if (cursor || units > 131072u || ancestors + _workspaceAncestorLinks > 8192u ||
            ! CanAdmitWorkspaceBytes(ancestors * 128u + path.size() * sizeof(wchar_t)))
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXCLUSION_LIMIT);
            return;
        }
        if (IsItemExcluded(id))
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXCLUDED_PARENT);
            return;
        }
        // Replace nested exclusions with their parent; never subtract the same data twice.
        std::erase_if(_excludedPaths,
                      [&](const ExcludedPath& entry)
        {
            return item.value().isDirectory && (EquivalentPath(_model->workspacePathIdentity.value(), entry.path, path) ||
                                                IsStrictDescendantPath(_model->workspacePathIdentity.value(), path, entry.path));
        });
        if (_excludedPaths.size() >= kMaximumExclusions)
        {
            ShowWorkspaceMessage(IDS_VIEWERSPACE_EXCLUSION_LIMIT);
            return;
        }
        _excludedPaths.push_back({path, item.value().isDirectory, id});
    }
    _workspaceResetBindings = true;
    ContinueWorkspaceBindings();
    if (_workspaceRebinding)
    {
        UpdateHeaderTextCache();
        InvalidateStaticTreemapCache();
        if (_hWnd)
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    else if (_excludedPaths.empty())
        WorkspaceChanged();
}

void ViewerSpace::RestoreExcludedItems() noexcept
{
    _excludedPaths.clear();
    _excludedRoots.clear();
    _excludedBelow.clear();
    _workspaceAncestorLinks = 0u;
    _workspaceNextExcludedRoots.clear();
    _workspaceNextExcludedBelow.clear();
    _workspaceNextAncestorLinks = 0u;
    _workspaceResetBindings     = false;
    _workspaceBindingGeneration = _displayGeneration;
    _workspaceBindingParts.clear();
    _workspaceRebinding  = false;
    _workspaceUnresolved = 0u;
    _workspaceGeneration = _displayGeneration;
    WorkspaceChanged();
}

void ViewerSpace::ContinueWorkspaceBindings() noexcept
{
    if (! _model)
        return;
    const auto started = std::chrono::steady_clock::now();
    if (_cleanupGeneration && _cleanupGeneration != _displayGeneration)
    {
        _cleanupMarks.clear();
        _cleanupGeneration = _displayGeneration;
        _reviewingCleanup  = false;
        UpdateWorkspaceScope();
        ShowWorkspaceMessage(IDS_VIEWERSPACE_REVIEW_STALE);
    }
    if (! _model->workspacePathIdentity.has_value() && _fileSystemShortId == L"file")
        _model->workspacePathIdentity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    if (! _model->workspacePathIdentity.has_value())
    {
        _workspaceUnresolved = _excludedPaths.size();
        return;
    }
    const auto& identity     = _model->workspacePathIdentity.value();
    const bool newGeneration = _workspaceBindingGeneration != _displayGeneration;
    const bool retryMissing  = ! _workspaceRebinding && _workspaceUnresolved && _workspaceBoundRevision != _modelRevision;
    if (newGeneration || _workspaceResetBindings || retryMissing)
    {
        const bool keepQualifiedIds = _workspaceGeneration == _displayGeneration;
        _workspaceBindingGeneration = _displayGeneration;
        _workspaceBindingRevision   = _modelRevision;
        _workspaceResetBindings     = false;
        _workspaceBindingIndex      = 0u;
        _workspaceBindingParent     = 0u;
        _workspaceUnresolved        = 0u;
        _workspaceNextExcludedRoots.clear();
        _workspaceNextExcludedBelow.clear();
        _workspaceNextAncestorLinks = 0u;
        if (! keepQualifiedIds)
            for (auto& entry : _excludedPaths)
                entry.id = 0u;
        // Current-generation IDs need no sibling/path lookup. Resolve missing retained paths
        // afterward, while the previous coherent projection remains published.
        std::stable_partition(_excludedPaths.begin(), _excludedPaths.end(), [](const ExcludedPath& entry) noexcept { return entry.id != 0u; });
        _workspaceRebinding = ! _excludedPaths.empty();
        if (! _workspaceRebinding)
        {
            _workspaceGeneration = _displayGeneration;
            _excludedRoots.clear();
            _excludedBelow.clear();
            _workspaceAncestorLinks = 0u;
            _workspaceViewExcluded  = false;
        }
    }
    if (! _workspaceRebinding)
        return;

    size_t work = 0u;
    while (_workspaceBindingIndex < _excludedPaths.size())
    {
        auto& entry = _excludedPaths[_workspaceBindingIndex];
        if (! entry.id && ! _workspaceBindingParent)
        {
            _workspaceBindingParts.clear();
            const std::wstring_view root = _model->scanRootPath;
            std::wstring relative;
            if (entry.directory && (EquivalentPath(identity, root, entry.path) || IsStrictDescendantPath(identity, entry.path, root)))
                entry.id = _model->rootNodeId;
            else if (TryGetFileSystemRelativePath(identity, root, entry.path, relative))
            {
                std::wstring_view remaining(relative);
                while (! remaining.empty() && _workspaceBindingParts.size() <= kMaximumPathDepth)
                {
                    const size_t separator = remaining.find_first_of(identity.acceptedSeparators);
                    _workspaceBindingParts.emplace_back(remaining.substr(0u, separator));
                    remaining.remove_prefix(separator == std::wstring_view::npos ? remaining.size() : separator + 1u);
                }
                _workspaceBindingParent = _model->rootNodeId;
                _workspaceBindingPart = _workspaceBindingChild = 0u;
                _workspaceBindingMatch                         = 0u;
                _workspaceBindingAmbiguous                     = false;
            }
            if (! entry.id && (! _workspaceBindingParent || _workspaceBindingParts.empty() || _workspaceBindingParts.size() > kMaximumPathDepth))
            {
                ++_workspaceUnresolved;
                ++_workspaceBindingIndex;
                _workspaceBindingParent = 0u;
                continue;
            }
        }
        if (! entry.id)
        {
            const Node* parent  = TryGetRealNode(_workspaceBindingParent);
            const auto children = parent ? GetRealNodeChildren(*parent) : std::span<const uint32_t>{};
            if (_workspaceBindingChild < children.size())
            {
                const uint32_t id    = children[_workspaceBindingChild++];
                const auto candidate = ResolveAnalysisItem(id);
                const bool last      = _workspaceBindingPart + 1u == _workspaceBindingParts.size();
                if (candidate.has_value() && EquivalentComponent(identity, candidate.value().name, _workspaceBindingParts[_workspaceBindingPart]) &&
                    (last ? candidate.value().isDirectory == entry.directory : candidate.value().isDirectory))
                {
                    _workspaceBindingAmbiguous = _workspaceBindingMatch != 0u;
                    _workspaceBindingMatch     = id;
                }
                ++work;
                if (work >= 32768u || (work % 64u == 0u && Debug::Perf::ElapsedUs(started) >= 1200u))
                    break;
                continue;
            }
            if (! _workspaceBindingMatch || _workspaceBindingAmbiguous)
            {
                ++_workspaceUnresolved;
                ++_workspaceBindingIndex;
                _workspaceBindingParent = 0u;
                continue;
            }
            if (++_workspaceBindingPart == _workspaceBindingParts.size())
                entry.id = _workspaceBindingMatch;
            else
            {
                _workspaceBindingParent    = _workspaceBindingMatch;
                _workspaceBindingChild     = 0u;
                _workspaceBindingMatch     = 0u;
                _workspaceBindingAmbiguous = false;
                continue;
            }
        }
        std::array<uint32_t, kMaximumPathDepth> ancestors{};
        size_t ancestorCount = 0u;
        uint32_t ancestor    = entry.id;
        for (size_t depth = 0u; ancestor && depth < kMaximumPathDepth; ++depth)
        {
            ancestors[ancestorCount++] = ancestor;
            const auto resolved        = ResolveAnalysisItem(ancestor);
            ancestor                   = resolved.has_value() ? resolved.value().parentId : 0u;
        }
        if (ancestor || ancestorCount + _workspaceNextAncestorLinks > 8192u)
        {
            entry.id = 0u;
            ++_workspaceUnresolved;
        }
        else
        {
            _workspaceNextExcludedRoots.insert(entry.id);
            for (size_t index = 0u; index < ancestorCount; ++index)
                _workspaceNextExcludedBelow[ancestors[index]].push_back(entry.id);
            _workspaceNextAncestorLinks += ancestorCount;
        }
        _workspaceBindingParent = 0u;
        ++_workspaceBindingIndex;
        if (Debug::Perf::ElapsedUs(started) >= 1200u)
            break;
    }
    if (_workspaceBindingIndex == _excludedPaths.size())
    {
        const bool changedDuringBinding = _workspaceBindingRevision != _modelRevision;
        _workspaceRebinding             = false;
        _workspaceGeneration            = _displayGeneration;
        _excludedRoots.swap(_workspaceNextExcludedRoots);
        _excludedBelow.swap(_workspaceNextExcludedBelow);
        _workspaceAncestorLinks = _workspaceNextAncestorLinks;
        WorkspaceChanged();
        _workspaceBoundRevision = changedDuringBinding ? 0u : _modelRevision;
    }
    else
        EnsureAnimationTimer();
    Debug::Perf::Emit(L"viewer.space.exclusions.bind_us", L"retained-paths-no-provider-io", Debug::Perf::ElapsedUs(started), work, _excludedPaths.size(), S_OK);
}

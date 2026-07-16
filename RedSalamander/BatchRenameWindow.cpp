#include "Framework.h"

#include "BatchRenameWindow.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "BatchRenameExecutionEngine.h"
#include "BatchRenameMenus.h"
#include "DirectoryInfoCache.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "IconCache.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"
#include "NavigationView.h"
#include "WindowMaximizeBehavior.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ButtonVariant;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ContextMenu;
using RedSalamander::DxUi::ContextMenuRootHorizontalAlignment;
using RedSalamander::DxUi::ContextMenuRootVerticalPlacement;
using RedSalamander::DxUi::ContextMenuSessionCallbacks;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridCellKind;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridColumnLayoutEntry;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MenuFlyoutItem;
using RedSalamander::DxUi::MenuItemKind;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::RadioButton;
using RedSalamander::DxUi::RadioButtons;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::StatusStrip;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kBatchRenameWindowClassName[]                  = L"RedSalamander.BatchRenameWindow";
constexpr wchar_t kBatchRenameWindowId[]                         = L"BatchRenameWindow";
constexpr size_t kMaxRecentBatchRenameEntries                    = 10u;
constexpr wchar_t kDefaultBatchRenameMask[]                      = L"*.*";
constexpr UINT_PTR kPreviewRebuildTimerId                        = 0xB471u;
constexpr UINT kPreviewRebuildDebounceMs                         = 150u;
constexpr int kBatchRenamePreviewMenuCopyOriginalName            = 1;
constexpr int kBatchRenamePreviewMenuCopyNewName                 = 2;
constexpr int kBatchRenamePreviewMenuCopySourcePath              = 3;
constexpr int kBatchRenamePreviewMenuCopyPreviewRows             = 4;
constexpr int kBatchRenamePreviewMenuRevealInPane                = 5;
constexpr int kBatchRenamePreviewMenuCopyUndoPlan                = 6;
constexpr int kBatchRenamePreviewMenuCopyExecutionReport         = 7;
constexpr wchar_t kBatchRenameIssueProviderPathIdentityUnknown[] = L"provider_path_identity_unknown";

#ifdef ENABLE_TESTS
std::mutex g_batchRenameDebugDestinationProbeFailureMutex;
std::optional<std::filesystem::path> g_batchRenameDebugDestinationProbeFailurePath;
unsigned long g_batchRenameDebugDestinationProbeFailureWin32Error = ERROR_ACCESS_DENIED;

[[nodiscard]] bool TryInjectBatchRenameDestinationProbeFailure(const FileSystemPathIdentity& pathIdentity,
                                                               const std::filesystem::path& destinationPath,
                                                               std::error_code& outError) noexcept
{
    std::lock_guard lock(g_batchRenameDebugDestinationProbeFailureMutex);
    if (! g_batchRenameDebugDestinationProbeFailurePath.has_value())
    {
        return false;
    }

    if (! EquivalentPath(pathIdentity, g_batchRenameDebugDestinationProbeFailurePath->native(), destinationPath.native()))
    {
        return false;
    }

    outError = std::error_code(static_cast<int>(g_batchRenameDebugDestinationProbeFailureWin32Error), std::system_category());
    return true;
}
#endif

struct BatchRenameScopeOptions final
{
    std::wstring mask          = kDefaultBatchRenameMask;
    bool includeSubdirectories = false;
    bool includeFiles          = true;
    bool includeFolders        = false;
};

constexpr WPARAM kBatchRenameTaskCollection = 1u;
constexpr WPARAM kBatchRenameTaskExecution  = 2u;
constexpr WPARAM kBatchRenameTaskPreview    = 3u;
constexpr unsigned char kBatchRenameModuleAnchor = 0u;

struct BatchRenameTaskProgressPayload final
{
    uint64_t generation     = 0u;
    uint64_t totalItems     = 0u;
    uint64_t completedItems = 0u;
};

struct BatchRenameCollectionCompletedPayload final
{
    uint64_t generation = 0u;
    HRESULT hr          = S_OK;
    std::wstring detail;
    std::vector<BatchRename::Target> targets;
};

struct BatchRenameExecutionCompletedPayload final
{
    uint64_t generation = 0u;
    HRESULT hr          = S_OK;
    std::wstring detail;
    BatchRenameExecutionReport report;
    std::vector<std::filesystem::path> successfulSourcePaths;
    std::vector<std::filesystem::path> successfulTargetPaths;
    std::vector<ExecutedDirectoryMove> executedDirectoryMoves;
};

struct BatchRenamePreviewCompletedPayload final
{
    uint64_t generation = 0u;
    BatchRename::Plan plan;
};

struct BatchPreviewRow
{
    uint64_t stableId  = 0u;
    size_t targetIndex = 0u;
    std::filesystem::path sourcePath;
    std::wstring originalName;
    std::wstring newName;
    std::wstring sizeText;
    std::wstring dateText;
    std::wstring timeText;
    std::wstring fullPath;
    // Containing folder relative to the Batch Rename root (Find Files Path-column semantics):
    // empty for items directly under the root; absolute parent folder when outside the root.
    std::wstring displayPath;
    uint64_t sizeBytes   = 0u;
    int iconIndex        = -1;
    bool isDirectory     = false;
    bool metadataUnknown = false;
    bool hasErrorIssue   = false;
    bool hasWarningIssue = false;
    bool changed         = false;
    std::wstring issueTooltip;
};

template <typename FileTimePoint> [[nodiscard]] std::optional<std::chrono::sys_seconds> FileTimeToSysSeconds(const FileTimePoint fileTime) noexcept
{
    const auto nowFile = FileTimePoint::clock::now();
    const auto nowSys  = std::chrono::system_clock::now();
    return std::chrono::time_point_cast<std::chrono::seconds>(nowSys + (fileTime - nowFile));
}

[[nodiscard]] std::optional<std::chrono::sys_seconds> FileTimeTicksToSysSeconds(const __int64 fileTimeTicks) noexcept
{
    if (fileTimeTicks <= 0)
    {
        return std::nullopt;
    }

    constexpr __int64 kWindowsToUnixEpoch100Ns = 116444736000000000LL;
    constexpr __int64 kFileTimeTicksPerSecond  = 10000000LL;
    const __int64 seconds                      = (fileTimeTicks - kWindowsToUnixEpoch100Ns) / kFileTimeTicksPerSecond;
    return std::chrono::sys_seconds{std::chrono::seconds{seconds}};
}

[[nodiscard]] bool IsDotDotPathPart(const std::filesystem::path& path) noexcept
{
    return path.native() == L"..";
}

[[nodiscard]] bool IsDotOrDotDotName(std::wstring_view name) noexcept
{
    return name == L"." || name == L"..";
}

[[nodiscard]] std::filesystem::path MakeRelativeParentFolder(const std::filesystem::path& path, const std::filesystem::path& root)
{
    if (root.empty())
    {
        return {};
    }

    const std::filesystem::path relative = path.parent_path().lexically_normal().lexically_relative(root.lexically_normal());
    if (relative.empty() || relative.native() == L".")
    {
        return {};
    }

    auto first = relative.begin();
    if (first != relative.end() && IsDotDotPathPart(*first))
    {
        return {};
    }

    return relative;
}

[[nodiscard]] BatchRename::Target BuildTargetFromLocalPath(const std::filesystem::path& path, const std::filesystem::path& root = {})
{
    BatchRename::Target target{};
    target.sourcePath     = path;
    target.relativeFolder = MakeRelativeParentFolder(path, root);

    std::error_code ec;
    const std::filesystem::file_status status = std::filesystem::symlink_status(path, ec);
    if (! ec)
    {
        target.isDirectory = std::filesystem::is_directory(status);
        if (! target.isDirectory && std::filesystem::is_regular_file(status))
        {
            const uintmax_t size = std::filesystem::file_size(path, ec);
            if (! ec)
            {
                target.sizeBytes = static_cast<uint64_t>((std::min)(size, static_cast<uintmax_t>((std::numeric_limits<uint64_t>::max)())));
            }
        }
    }

    ec.clear();
    const std::filesystem::file_time_type lastWriteTime = std::filesystem::last_write_time(path, ec);
    if (! ec)
    {
        target.lastWriteTime = FileTimeToSysSeconds(lastWriteTime);
    }

    WIN32_FILE_ATTRIBUTE_DATA attributes{};
    if (::GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &attributes) != FALSE)
    {
        ULARGE_INTEGER creationTime{};
        creationTime.LowPart  = attributes.ftCreationTime.dwLowDateTime;
        creationTime.HighPart = attributes.ftCreationTime.dwHighDateTime;
        target.createdTime    = FileTimeTicksToSysSeconds(static_cast<__int64>(creationTime.QuadPart));
    }

    return target;
}

[[nodiscard]] BatchRename::Target BuildTargetFromProviderFileInfo(const std::filesystem::path& path, const std::filesystem::path& root, const FileInfo& entry)
{
    BatchRename::Target target{};
    target.sourcePath     = path;
    target.relativeFolder = MakeRelativeParentFolder(path, root);
    target.isDirectory    = (entry.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    if (! target.isDirectory && entry.EndOfFile >= 0)
    {
        target.sizeBytes = static_cast<uint64_t>(entry.EndOfFile);
    }
    target.createdTime   = FileTimeTicksToSysSeconds(entry.CreationTime);
    target.lastWriteTime = FileTimeTicksToSysSeconds(entry.LastWriteTime);
    return target;
}

[[nodiscard]] BatchRename::Target BuildMetadataUnknownTargetFromProviderSelection(const std::filesystem::path& path, const std::filesystem::path& root)
{
    BatchRename::Target target{};
    target.sourcePath      = path;
    target.relativeFolder  = MakeRelativeParentFolder(path, root);
    target.metadataUnknown = true;
    return target;
}

[[nodiscard]] bool IsNotFoundError(const std::error_code& ec) noexcept
{
    return ec == std::errc::no_such_file_or_directory || ec.value() == ERROR_FILE_NOT_FOUND || ec.value() == ERROR_PATH_NOT_FOUND;
}

[[nodiscard]] HRESULT ErrorCodeToHRESULT(const std::error_code& ec) noexcept
{
    return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : E_FAIL;
}

[[nodiscard]] HRESULT RevalidateLocalRenamePlan(const FileSystemPathIdentity& pathIdentity, const BatchRename::Plan& plan) noexcept;

void EmitBatchRenameExecuteCounters(const uint64_t rows, const uint64_t completed, const uint64_t failed) noexcept
{
    if (! Debug::Perf::IsCaptureEnabled())
    {
        return;
    }

    Debug::Perf::EmitValue(L"batchrename.execute.rows", rows);
    Debug::Perf::EmitValue(L"batchrename.execute.completed", completed);
    Debug::Perf::EmitValue(L"batchrename.execute.failed", failed);
}

struct BatchRenameExecutionProgressPostContext final
{
    HWND hwnd           = nullptr;
    uint64_t generation = 0u;
};

void PostBatchRenameExecutionProgress(void* context, const uint64_t completedItems, const uint64_t totalItems, [[maybe_unused]] bool forcePost) noexcept
{
    const auto* postContext = static_cast<const BatchRenameExecutionProgressPostContext*>(context);
    if (! postContext || ! postContext->hwnd)
    {
        return;
    }

    auto payload = std::unique_ptr<BatchRenameTaskProgressPayload>(new (std::nothrow) BatchRenameTaskProgressPayload{});
    if (! payload)
    {
        return;
    }

    payload->generation     = postContext->generation;
    payload->totalItems     = totalItems;
    payload->completedItems = completedItems;
    static_cast<void>(PostMessagePayload(postContext->hwnd, WndMsg::kBatchRenameTaskUpdate, kBatchRenameTaskExecution, std::move(payload)));
}

// Runs the window-owned worker envelope: initialize MTA COM, keep the provider
// alive until before CoUninitialize, call the window-free execution engine, and
// post the completion payload back to the UI thread.
void RunBatchRenameExecution(HWND hwnd,
                             const uint64_t generation,
                             std::atomic_bool& cancelRequested,
                             wil::com_ptr<IFileSystem> fileSystem,
                             const FileSystemPathIdentity pathIdentity,
                             std::optional<BatchRename::Plan> localPlan,
                             std::vector<BatchRenameExecutionOp> ops,
                             std::unique_ptr<BatchRenameExecutionCompletedPayload> payload) noexcept
{
    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(coinitHr))
    {
        Debug::Error(L"BatchRename execution task: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", static_cast<unsigned long>(coinitHr));
        FAIL_FAST_IF_FAILED(coinitHr);
    }
    [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

    wil::com_ptr<IFileSystem> workerFileSystem = std::move(fileSystem);
    const auto workerStartedAt = std::chrono::steady_clock::now();
    if (localPlan.has_value())
    {
        const HRESULT revalidateHr = RevalidateLocalRenamePlan(pathIdentity, localPlan.value());
        if (FAILED(revalidateHr))
        {
            payload->hr                  = revalidateHr;
            payload->detail              = L"revalidate_failed";
            payload->report.failedRows   = localPlan->stats.changedRows;
            payload->report.firstFailure = revalidateHr;
            Debug::Perf::Emit(L"batchrename.execute.us",
                              L"revalidate_failed",
                              Debug::Perf::ElapsedUs(workerStartedAt),
                              static_cast<uint64_t>(localPlan->stats.changedRows),
                              0u,
                              revalidateHr);
            EmitBatchRenameExecuteCounters(static_cast<uint64_t>(localPlan->stats.changedRows), 0u, static_cast<uint64_t>(localPlan->stats.changedRows));
            static_cast<void>(PostMessagePayload(hwnd, WndMsg::kBatchRenameCompleted, kBatchRenameTaskExecution, std::move(payload)));
            return;
        }
    }
    BatchRenameExecutionProgressPostContext progressContext{.hwnd = hwnd, .generation = generation};
    const BatchRenameExecutionResult result = RunBatchRenameExecutionEngine(cancelRequested,
                                                                            *workerFileSystem.get(),
                                                                            pathIdentity,
                                                                            std::move(ops),
                                                                            BatchRenameExecutionOptions{
                                                                                .progressCallback = &PostBatchRenameExecutionProgress,
                                                                                .progressContext  = &progressContext,
                                                                            });
    payload->hr                             = result.hr;
    payload->detail                         = result.detail;
    // The execution engine returns a fresh report that only tracks completed/failed counts; it does not
    // know the plan-level totals. The caller pre-populated payload->report with totalRows/skippedRows
    // before launching this worker, so preserve them across the engine-result assignment (otherwise the
    // success-path summary, TSV export, and persisted stats would all report "0 of 0").
    const size_t totalRows          = payload->report.totalRows;
    const size_t skippedRows        = payload->report.skippedRows;
    payload->report                 = result.report;
    payload->report.totalRows       = totalRows;
    payload->report.skippedRows     = skippedRows;
    payload->successfulSourcePaths  = result.successfulSourcePaths;
    payload->successfulTargetPaths  = result.successfulTargetPaths;
    payload->executedDirectoryMoves = result.executedDirectoryMoves;

    static_cast<void>(PostMessagePayload(hwnd, WndMsg::kBatchRenameCompleted, kBatchRenameTaskExecution, std::move(payload)));
}

[[nodiscard]] bool IsLocalFileSystemContext(const BatchRenamePaneContext& context) noexcept
{
    return ::CompareStringOrdinal(context.pluginId.c_str(), -1, L"builtin/file-system", -1, TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::optional<FileSystemPathIdentity> ResolveBatchRenamePathIdentity(const BatchRenamePaneContext& context) noexcept
{
    if (context.pluginId.empty() || IsLocalFileSystemContext(context))
    {
        return FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    }

    if (! context.fileSystem)
    {
        return std::nullopt;
    }

    const char* capabilitiesJson = nullptr;
    const HRESULT capabilitiesHr = context.fileSystem->GetCapabilities(&capabilitiesJson);
    if (FAILED(capabilitiesHr) || capabilitiesJson == nullptr || capabilitiesJson[0] == '\0')
    {
        return std::nullopt;
    }

    return TryParseFileSystemRenamePathIdentity(capabilitiesJson, context.pluginId);
}

[[nodiscard]] bool IsPlannedSourcePath(const FileSystemPathIdentity& pathIdentity,
                                       const std::vector<std::filesystem::path>& plannedSources,
                                       const std::filesystem::path& path) noexcept
{
    return std::ranges::any_of(plannedSources, [&pathIdentity, &path](const std::filesystem::path& source) noexcept {
        return EquivalentPath(pathIdentity, source.native(), path.native());
    });
}

void AddProviderPathIdentityFailure(BatchRename::Plan& plan) noexcept
{
    bool changed = false;
    for (BatchRename::PreviewRow& row : plan.rows)
    {
        BatchRename::AddIssue(row, BatchRename::IssueSeverity::Error, kBatchRenameIssueProviderPathIdentityUnknown);
        changed = true;
    }

    if (changed)
    {
        BatchRename::RecomputeStats(plan);
    }
}

void ApplyLocalDestinationConflictValidation(const FileSystemPathIdentity& pathIdentity, BatchRename::Plan& plan) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();

    std::vector<std::filesystem::path> plannedSources;
    plannedSources.reserve(plan.rows.size());
    for (const BatchRename::PreviewRow& row : plan.rows)
    {
        if (row.newName != row.originalName)
        {
            plannedSources.push_back(row.sourcePath);
        }
    }

    struct ParentListing final
    {
        std::filesystem::path path;
        std::optional<std::wstring> pathKey;
        std::error_code error;
        std::unordered_set<std::wstring> childKeys;
        std::vector<std::wstring> childNames;
    };

    std::vector<ParentListing> parentListings;
    std::unordered_map<std::wstring, size_t> parentIndexByKey;
    const auto findOrCollectParent = [&](const std::filesystem::path& parent) -> const ParentListing&
    {
        const std::optional<std::wstring> parentKey = TryMakePathKey(pathIdentity, parent.native());
        if (parentKey.has_value())
        {
            const auto found = parentIndexByKey.find(parentKey.value());
            if (found != parentIndexByKey.end() && found->second < parentListings.size() &&
                EquivalentPath(pathIdentity, parentListings[found->second].path.native(), parent.native()))
            {
                return parentListings[found->second];
            }
        }

        const auto existing = std::ranges::find_if(parentListings, [&](const ParentListing& listing) noexcept {
            return EquivalentPath(pathIdentity, listing.path.native(), parent.native());
        });
        if (existing != parentListings.end())
        {
            return *existing;
        }

        ParentListing listing{};
        listing.path    = parent;
        listing.pathKey = parentKey;
        constexpr std::filesystem::directory_options options = std::filesystem::directory_options::skip_permission_denied;
        std::filesystem::directory_iterator iterator(parent, options, listing.error);
        if (! listing.error)
        {
            for (const std::filesystem::directory_iterator end; iterator != end; iterator.increment(listing.error))
            {
                if (listing.error)
                {
                    break;
                }
                std::wstring childName = iterator->path().filename().native();
                if (const std::optional<std::wstring> childKey = TryMakeComponentKey(pathIdentity, childName); childKey.has_value())
                {
                    listing.childKeys.insert(childKey.value());
                }
                listing.childNames.push_back(std::move(childName));
            }
        }

        parentListings.push_back(std::move(listing));
        const size_t index = parentListings.size() - 1u;
        if (parentListings[index].pathKey.has_value())
        {
            parentIndexByKey[parentListings[index].pathKey.value()] = index;
        }
        return parentListings[index];
    };

    bool changed = false;
    for (BatchRename::PreviewRow& row : plan.rows)
    {
        if (row.newName == row.originalName || row.newName.empty())
        {
            continue;
        }

        const std::filesystem::path destinationPath = JoinFolderAndLeaf(pathIdentity, row.sourcePath.parent_path(), row.newName);
        if (EquivalentPath(pathIdentity, row.sourcePath.native(), destinationPath.native()))
        {
            continue;
        }

        std::error_code injectedError;
#ifdef ENABLE_TESTS
        if (TryInjectBatchRenameDestinationProbeFailure(pathIdentity, destinationPath, injectedError))
        {
            BatchRename::AddIssue(row, BatchRename::IssueSeverity::Error, L"name_destination_probe_failed");
            changed = true;
            continue;
        }
#endif
        const ParentListing& listing = findOrCollectParent(destinationPath.parent_path());
        if (listing.error)
        {
            if (! IsNotFoundError(listing.error))
            {
                BatchRename::AddIssue(row, BatchRename::IssueSeverity::Error, L"name_destination_probe_failed");
                changed = true;
            }
            continue;
        }

        const std::wstring destinationLeaf = destinationPath.filename().native();
        bool destinationExists             = false;
        if (const std::optional<std::wstring> destinationKey = TryMakeComponentKey(pathIdentity, destinationLeaf); destinationKey.has_value())
        {
            destinationExists = listing.childKeys.contains(destinationKey.value());
        }
        else
        {
            destinationExists = std::ranges::any_of(listing.childNames, [&](std::wstring_view childName) noexcept {
                return EquivalentComponent(pathIdentity, childName, destinationLeaf);
            });
        }

        if (destinationExists && ! IsPlannedSourcePath(pathIdentity, plannedSources, destinationPath))
        {
            BatchRename::AddIssue(row, BatchRename::IssueSeverity::Error, L"name_destination_exists");
            changed = true;
        }
    }

    if (changed)
    {
        BatchRename::RecomputeStats(plan);
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"batchrename.preview.destination_validation.us",
                          L"cached-parent-listings",
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(plan.rows.size()),
                          static_cast<uint64_t>(parentListings.size()));
        Debug::Perf::EmitValue(L"batchrename.preview.destination_directory_listings", static_cast<uint64_t>(parentListings.size()));
    }
}

void ApplyContextualPreviewValidation(const BatchRenamePaneContext& context, const FileSystemPathIdentity& pathIdentity, BatchRename::Plan& plan) noexcept
{
    if (IsLocalFileSystemContext(context))
    {
        ApplyLocalDestinationConflictValidation(pathIdentity, plan);
    }
}

[[nodiscard]] BatchRename::Plan BuildBatchRenamePlanForContext(const BatchRenamePaneContext& context,
                                                               const std::vector<BatchRename::Target>& targets,
                                                               const BatchRename::Rules& rules) noexcept
{
    const std::optional<FileSystemPathIdentity> pathIdentity = ResolveBatchRenamePathIdentity(context);
    const FileSystemPathIdentity effectiveIdentity           = pathIdentity.value_or(FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem());

    BatchRename::Plan plan = BatchRename::BuildPlan(targets, rules, effectiveIdentity);
    if (! pathIdentity.has_value())
    {
        AddProviderPathIdentityFailure(plan);
        return plan;
    }

    ApplyContextualPreviewValidation(context, effectiveIdentity, plan);
    return plan;
}

struct BatchRenamePreviewWork final
{
    HWND hwnd = nullptr;
    uint64_t generation = 0u;
    BatchRenamePaneContext context;
    std::vector<BatchRename::Target> targets;
    BatchRename::Rules rules;
    std::unique_ptr<BatchRenamePreviewCompletedPayload> payload;
    wil::unique_hmodule modulePin;

    BatchRenamePreviewWork() = default;
    BatchRenamePreviewWork(const BatchRenamePreviewWork&) = delete;
    BatchRenamePreviewWork(BatchRenamePreviewWork&&) = delete;
    BatchRenamePreviewWork& operator=(const BatchRenamePreviewWork&) = delete;
    BatchRenamePreviewWork& operator=(BatchRenamePreviewWork&&) = delete;

    void Execute(PTP_CALLBACK_INSTANCE callbackInstance) noexcept
    {
        if (modulePin)
        {
            TransferModulePinToCallbackReturn(callbackInstance, modulePin);
        }

        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            payload->plan = {};
        }
        else
        {
            const wil::unique_couninitialize_call coUninit;
            // Provider path-identity queries are not universally concurrency-safe. Serialize preview
            // workers while keeping every provider query and destination listing off the UI thread.
            static std::mutex previewWorkerMutex;
            std::scoped_lock lock(previewWorkerMutex);
            payload->plan = BuildBatchRenamePlanForContext(context, targets, rules);
        }
        payload->generation = generation;
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kBatchRenameCompleted, kBatchRenameTaskPreview, std::move(payload)));
    }
};

[[nodiscard]] HRESULT RevalidateLocalRenamePlan(const FileSystemPathIdentity& pathIdentity, const BatchRename::Plan& plan) noexcept
{
    std::vector<std::filesystem::path> plannedSources;
    plannedSources.reserve(plan.rows.size());
    for (const BatchRename::PreviewRow& row : plan.rows)
    {
        if (row.newName != row.originalName)
        {
            plannedSources.push_back(row.sourcePath);
        }
    }

    for (const BatchRename::PreviewRow& row : plan.rows)
    {
        if (row.newName == row.originalName)
        {
            continue;
        }

        std::error_code ec;
        const std::filesystem::file_status sourceStatus = std::filesystem::symlink_status(row.sourcePath, ec);
        if (ec)
        {
            return ErrorCodeToHRESULT(ec);
        }
        if (! std::filesystem::exists(sourceStatus))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        const std::filesystem::path destinationPath = JoinFolderAndLeaf(pathIdentity, row.sourcePath.parent_path(), row.newName);
        if (EquivalentPath(pathIdentity, row.sourcePath.native(), destinationPath.native()))
        {
            continue;
        }

        ec.clear();
        const std::filesystem::file_status destinationStatus = std::filesystem::symlink_status(destinationPath, ec);
        if (ec && ! IsNotFoundError(ec))
        {
            return ErrorCodeToHRESULT(ec);
        }
        if (! ec && std::filesystem::exists(destinationStatus) && ! IsPlannedSourcePath(pathIdentity, plannedSources, destinationPath))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    return S_OK;
}

[[nodiscard]] BatchRename::Target RefreshTargetAfterRename(const BatchRename::Target& previous,
                                                           const std::filesystem::path& newPath,
                                                           const std::filesystem::path& root)
{
    std::error_code ec;
    static_cast<void>(std::filesystem::symlink_status(newPath, ec));
    if (! ec)
    {
        return BuildTargetFromLocalPath(newPath, root);
    }

    BatchRename::Target refreshed = previous;
    refreshed.sourcePath          = newPath;
    refreshed.relativeFolder      = MakeRelativeParentFolder(newPath, root);
    return refreshed;
}

struct BatchRenameTargetRefreshResult final
{
    size_t refreshedRows         = 0u;
    uint64_t identityComparisons = 0u;
};

BatchRenameTargetRefreshResult RefreshBatchRenameTargetsAfterExecution(const FileSystemPathIdentity& pathIdentity,
                                                                       std::vector<BatchRename::Target>& targets,
                                                                       std::span<const std::filesystem::path> successfulSourcePaths,
                                                                       std::span<const std::filesystem::path> successfulTargetPaths,
                                                                       std::span<const ExecutedDirectoryMove> executedDirectoryMoves,
                                                                       const std::filesystem::path& root) noexcept
{
    const auto targetRefreshStartedAt = std::chrono::steady_clock::now();
    const size_t successCount         = std::min(successfulSourcePaths.size(), successfulTargetPaths.size());

    std::unordered_map<std::wstring, std::vector<size_t>> targetIndexesByPathKey;
    targetIndexesByPathKey.reserve(targets.size());
    for (size_t targetIndex = 0u; targetIndex < targets.size(); ++targetIndex)
    {
        if (const std::optional<std::wstring> pathKey = TryMakePathKey(pathIdentity, targets[targetIndex].sourcePath.native()); pathKey.has_value())
        {
            targetIndexesByPathKey[pathKey.value()].push_back(targetIndex);
        }
    }

    std::vector<bool> refreshedTargets(targets.size(), false);
    BatchRenameTargetRefreshResult result{};
    const auto consumeTargetIndexForSource = [&](const std::filesystem::path& sourcePath) -> std::optional<size_t>
    {
        const std::optional<std::wstring> sourcePathKey = TryMakePathKey(pathIdentity, sourcePath.native());
        const auto bucketIt = sourcePathKey.has_value() ? targetIndexesByPathKey.find(sourcePathKey.value()) : targetIndexesByPathKey.end();
        if (bucketIt != targetIndexesByPathKey.end())
        {
            for (const size_t targetIndex : bucketIt->second)
            {
                if (targetIndex >= targets.size() || refreshedTargets[targetIndex])
                {
                    continue;
                }
                ++result.identityComparisons;
                if (EquivalentPath(pathIdentity, targets[targetIndex].sourcePath.native(), sourcePath.native()))
                {
                    return targetIndex;
                }
            }
        }

        for (size_t targetIndex = 0u; targetIndex < targets.size(); ++targetIndex)
        {
            if (refreshedTargets[targetIndex])
            {
                continue;
            }
            ++result.identityComparisons;
            if (EquivalentPath(pathIdentity, targets[targetIndex].sourcePath.native(), sourcePath.native()))
            {
                return targetIndex;
            }
        }
        return std::nullopt;
    };

    for (size_t index = 0u; index < successCount; ++index)
    {
        const std::filesystem::path finalPath   = ApplyExecutedDirectoryMoves(pathIdentity, successfulTargetPaths[index], executedDirectoryMoves);
        const std::optional<size_t> targetIndex = consumeTargetIndexForSource(successfulSourcePaths[index]);
        if (! targetIndex.has_value())
        {
            continue;
        }
        targets[targetIndex.value()]          = RefreshTargetAfterRename(targets[targetIndex.value()], finalPath, root);
        refreshedTargets[targetIndex.value()] = true;
        ++result.refreshedRows;
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"batchrename.execute.target_refresh_match.us",
                          L"identity_bucket",
                          Debug::Perf::ElapsedUs(targetRefreshStartedAt),
                          static_cast<uint64_t>(successCount),
                          result.identityComparisons);
    }
    return result;
}

[[nodiscard]] std::wstring NormalizeScopeMask(std::wstring_view mask)
{
    std::wstring normalized = StringUtils::TrimWhitespaceCopy(mask);
    if (normalized.empty())
    {
        normalized = kDefaultBatchRenameMask;
    }
    return normalized;
}

struct BatchRenameScopeMatcher final
{
    bool matchAll = true;
    MaskSyntax::WildcardMask wildcardMask;
};

[[nodiscard]] BatchRenameScopeMatcher BuildBatchRenameScopeMatcher(std::wstring_view mask)
{
    BatchRenameScopeMatcher matcher{};
    const std::wstring normalized = NormalizeScopeMask(mask);
    matcher.matchAll              = OrdinalString::EqualsNoCase(normalized, L"*") || OrdinalString::EqualsNoCase(normalized, kDefaultBatchRenameMask);
    if (! matcher.matchAll)
    {
        matcher.wildcardMask = MaskSyntax::ParseWildcardMask(normalized);
    }
    return matcher;
}

[[nodiscard]] bool ScopeMaskMatches(std::wstring_view leafName, const BatchRenameScopeMatcher& matcher) noexcept
{
    return matcher.matchAll || MaskSyntax::MatchesWildcardMask(leafName, matcher.wildcardMask);
}

[[nodiscard]] bool ShouldCollectLocalEntry(const std::filesystem::directory_entry& entry,
                                           const BatchRenameScopeOptions& scope,
                                           const BatchRenameScopeMatcher& matcher)
{
    std::error_code ec;
    const std::filesystem::file_status status = entry.symlink_status(ec);
    if (ec)
    {
        return false;
    }

    const bool isDirectory = std::filesystem::is_directory(status);
    if (isDirectory)
    {
        if (! scope.includeFolders)
        {
            return false;
        }
    }
    else if (! scope.includeFiles)
    {
        return false;
    }

    return ScopeMaskMatches(entry.path().filename().native(), matcher);
}

[[nodiscard]] bool ShouldCollectProviderEntry(std::wstring_view leafName,
                                              const DWORD attributes,
                                              const BatchRenameScopeOptions& scope,
                                              const BatchRenameScopeMatcher& matcher)
{
    const bool isDirectory = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    if (isDirectory)
    {
        if (! scope.includeFolders)
        {
            return false;
        }
    }
    else if (! scope.includeFiles)
    {
        return false;
    }

    return ScopeMaskMatches(leafName, matcher);
}

void SortBatchRenameTargets(std::vector<BatchRename::Target>& targets)
{
    std::ranges::sort(targets, [](const BatchRename::Target& lhs, const BatchRename::Target& rhs) noexcept {
        return ::CompareStringOrdinal(lhs.sourcePath.c_str(), -1, rhs.sourcePath.c_str(), -1, TRUE) == CSTR_LESS_THAN;
    });
}

struct BatchRenameTargetCollectionResult final
{
    std::vector<BatchRename::Target> targets;
    HRESULT hr          = S_OK;
    std::wstring detail = L"local";
};

[[nodiscard]] bool IsBatchRenameCollectionCancelRequested(const std::atomic_bool* cancelRequested) noexcept
{
    return cancelRequested && cancelRequested->load(std::memory_order_acquire);
}

template <typename Visitor>
[[nodiscard]] HRESULT ForEachFileInfoEntry(IFilesInformation& info, Visitor&& visitor, const std::atomic_bool* cancelRequested = nullptr)
{
    FileInfo* buffer = nullptr;
    HRESULT hr       = info.GetBuffer(&buffer);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! buffer)
    {
        return S_OK;
    }

    unsigned long bufferSize = 0;
    hr                       = info.GetBufferSize(&bufferSize);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned char* bytes = reinterpret_cast<unsigned char*>(buffer);
    unsigned long offset = 0;

    for (FileInfo* entry = buffer; entry;)
    {
        if ((entry->FileNameSize % sizeof(wchar_t)) != 0u)
        {
            return E_INVALIDARG;
        }

        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        std::wstring_view name(entry->FileName, nameChars);
        if (! name.empty() && ! IsDotOrDotDotName(name))
        {
            visitor(name, *entry);
        }

        if (entry->NextEntryOffset == 0u)
        {
            break;
        }

        if (entry->NextEntryOffset > bufferSize - offset)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        offset += entry->NextEntryOffset;
        if (offset >= bufferSize)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        if (IsBatchRenameCollectionCancelRequested(cancelRequested))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        entry = reinterpret_cast<FileInfo*>(bytes + offset);
    }

    return S_OK;
}

[[nodiscard]] HRESULT CollectProviderScopeTargets(const BatchRenamePaneContext& context,
                                                  const BatchRenameScopeOptions& scope,
                                                  std::vector<BatchRename::Target>& targets,
                                                  const std::atomic_bool* cancelRequested)
{
    if (! context.fileSystem)
    {
        return E_POINTER;
    }

    std::vector<std::filesystem::path> pending;
    pending.push_back(context.rootPluginPath);

    const BatchRenameScopeMatcher scopeMatcher = BuildBatchRenameScopeMatcher(scope.mask);
    const FileSystemPathIdentity collectionIdentity = ResolveBatchRenamePathIdentity(context).value_or(
        FileSystemPathIdentity{.pathTextStableIdentity = true, .componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive});
    std::unordered_set<std::wstring> queuedDirectories;
    queuedDirectories.insert(context.rootPluginPath.native());

    while (! pending.empty())
    {
        if (IsBatchRenameCollectionCancelRequested(cancelRequested))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        std::filesystem::path directory = std::move(pending.back());
        pending.pop_back();

        wil::com_ptr<IFilesInformation> info;
        const HRESULT readHr = context.fileSystem->ReadDirectoryInfo(directory.c_str(), info.addressof());
        if (FAILED(readHr) || ! info)
        {
            if (IsBatchRenameCancellationHRESULT(readHr))
            {
                return readHr;
            }
            continue;
        }

        const HRESULT walkHr = ForEachFileInfoEntry(*info,
                                                    [&](std::wstring_view name, const FileInfo& entry)
        {
            const std::filesystem::path child = JoinFolderAndLeaf(collectionIdentity, directory, name);
            const DWORD attributes            = entry.FileAttributes;
            if (ShouldCollectProviderEntry(name, attributes, scope, scopeMatcher))
            {
                targets.push_back(BuildTargetFromProviderFileInfo(child, context.rootPluginPath, entry));
            }

            const bool isDirectory    = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
            const bool isReparsePoint = (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u;
            if (scope.includeSubdirectories && isDirectory && ! isReparsePoint)
            {
                const std::wstring childKey = child.native();
                if (! childKey.empty() && queuedDirectories.insert(childKey).second)
                {
                    pending.push_back(child);
                }
            }
        },
                                                    cancelRequested);
        if (FAILED(walkHr))
        {
            return walkHr;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT CollectProviderSelectionTargets(const BatchRenamePaneContext& context,
                                                      std::vector<BatchRename::Target>& targets,
                                                      const std::atomic_bool* cancelRequested)
{
    if (! context.fileSystem)
    {
        return E_POINTER;
    }

    struct ParentDirectoryEntries final
    {
        std::filesystem::path parent;
        std::optional<std::wstring> parentPathKey;
        HRESULT hr = S_OK;
        std::vector<BatchRename::Target> targets;
    };

    const FileSystemPathIdentity collectionIdentity = ResolveBatchRenamePathIdentity(context).value_or(
        FileSystemPathIdentity{.pathTextStableIdentity = true, .componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive});

    std::vector<ParentDirectoryEntries> entriesByParent;
    std::unordered_map<std::wstring, size_t> parentIndexByPathKey;
    for (const std::filesystem::path& path : context.initialPaths)
    {
        if (IsBatchRenameCollectionCancelRequested(cancelRequested))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const std::filesystem::path parent              = path.parent_path();
        const std::optional<std::wstring> parentPathKey = TryMakePathKey(collectionIdentity, parent.native());
        auto parentIt                                   = entriesByParent.end();
        if (parentPathKey.has_value())
        {
            const auto keyedParentIt = parentIndexByPathKey.find(parentPathKey.value());
            if (keyedParentIt != parentIndexByPathKey.end() && keyedParentIt->second < entriesByParent.size() &&
                EquivalentPath(collectionIdentity, entriesByParent[keyedParentIt->second].parent.native(), parent.native()))
            {
                parentIt = entriesByParent.begin() + static_cast<std::ptrdiff_t>(keyedParentIt->second);
            }
        }
        if (parentIt == entriesByParent.end())
        {
            parentIt = std::ranges::find_if(entriesByParent, [&](const ParentDirectoryEntries& entries) noexcept {
                return EquivalentPath(collectionIdentity, entries.parent.native(), parent.native());
            });
        }
        if (parentIt == entriesByParent.end())
        {
            ParentDirectoryEntries entries{};
            entries.parent        = parent;
            entries.parentPathKey = parentPathKey;
            wil::com_ptr<IFilesInformation> info;
            entries.hr = context.fileSystem->ReadDirectoryInfo(parent.c_str(), info.addressof());
            if (SUCCEEDED(entries.hr) && info)
            {
                entries.hr = ForEachFileInfoEntry(*info,
                                                  [&](std::wstring_view name, const FileInfo& entry)
                {
                    const std::filesystem::path child = JoinFolderAndLeaf(collectionIdentity, parent, name);
                    entries.targets.push_back(BuildTargetFromProviderFileInfo(child, context.rootPluginPath, entry));
                },
                                                  cancelRequested);
            }
            else if (SUCCEEDED(entries.hr))
            {
                entries.hr = E_FAIL;
            }
            entriesByParent.push_back(std::move(entries));
            parentIt = entriesByParent.end();
            --parentIt;
            if (parentIt->parentPathKey.has_value())
            {
                parentIndexByPathKey[parentIt->parentPathKey.value()] = entriesByParent.size() - 1u;
            }
        }

        if (IsBatchRenameCancellationHRESULT(parentIt->hr))
        {
            return parentIt->hr;
        }

        const std::wstring selectedLeaf = path.filename().native();
        const auto entryIt              = std::ranges::find_if(parentIt->targets, [&](const BatchRename::Target& target) noexcept {
            return EquivalentComponent(collectionIdentity, target.sourcePath.filename().native(), selectedLeaf);
        });
        if (FAILED(parentIt->hr) || entryIt == parentIt->targets.end())
        {
            // The non-local provider could not describe this entry. Preserve
            // path identity but do not fabricate local size/type/timestamps.
            targets.push_back(BuildMetadataUnknownTargetFromProviderSelection(path, context.rootPluginPath));
            continue;
        }

        // Keep the caller-supplied source path so selection identity is preserved.
        BatchRename::Target target = *entryIt;
        target.sourcePath          = path;
        target.relativeFolder      = MakeRelativeParentFolder(path, context.rootPluginPath);
        targets.push_back(std::move(target));
    }

    return S_OK;
}

[[nodiscard]] HRESULT CollectLocalScopeTargets(const BatchRenamePaneContext& context,
                                               const BatchRenameScopeOptions& scope,
                                               std::vector<BatchRename::Target>& targets,
                                               const std::atomic_bool* cancelRequested)
{
    std::error_code ec;
    constexpr std::filesystem::directory_options options = std::filesystem::directory_options::skip_permission_denied;
    const BatchRenameScopeMatcher scopeMatcher           = BuildBatchRenameScopeMatcher(scope.mask);
    if (scope.includeSubdirectories)
    {
        std::filesystem::recursive_directory_iterator it(context.rootPluginPath, options, ec);
        if (ec)
        {
            return S_OK;
        }

        for (std::filesystem::recursive_directory_iterator end; it != end; it.increment(ec))
        {
            if (ec)
            {
                break;
            }
            if (IsBatchRenameCollectionCancelRequested(cancelRequested))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (ShouldCollectLocalEntry(*it, scope, scopeMatcher))
            {
                targets.push_back(BuildTargetFromLocalPath(it->path(), context.rootPluginPath));
            }
        }
        SortBatchRenameTargets(targets);
        return S_OK;
    }

    std::filesystem::directory_iterator it(context.rootPluginPath, options, ec);
    if (ec)
    {
        return S_OK;
    }
    for (std::filesystem::directory_iterator end; it != end; it.increment(ec))
    {
        if (ec)
        {
            break;
        }
        if (IsBatchRenameCollectionCancelRequested(cancelRequested))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (ShouldCollectLocalEntry(*it, scope, scopeMatcher))
        {
            targets.push_back(BuildTargetFromLocalPath(it->path(), context.rootPluginPath));
        }
    }

    SortBatchRenameTargets(targets);
    return S_OK;
}

[[nodiscard]] BatchRenameTargetCollectionResult CollectBatchRenameTargets(const BatchRenamePaneContext& context,
                                                                          const BatchRenameScopeOptions& scope,
                                                                          const std::atomic_bool* cancelRequested = nullptr)
{
    BatchRenameTargetCollectionResult result{};
    if (! context.initialPaths.empty())
    {
        result.detail = L"selection";
        result.targets.reserve(context.initialPaths.size());

        if (context.fileSystem && ! IsLocalFileSystemContext(context))
        {
            // Non-local providers (archives, cloud) must describe explicit
            // selections themselves; local stat calls would misclassify
            // provider directories as files.
            result.hr = CollectProviderSelectionTargets(context, result.targets, cancelRequested);
            if (SUCCEEDED(result.hr) || IsBatchRenameCancellationHRESULT(result.hr))
            {
                return result;
            }

            result.targets.clear();
            result.hr = S_OK;
        }

        for (const std::filesystem::path& path : context.initialPaths)
        {
            if (IsBatchRenameCollectionCancelRequested(cancelRequested))
            {
                result.hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return result;
            }
            result.targets.push_back(BuildTargetFromLocalPath(path, context.rootPluginPath));
        }
        return result;
    }

    if (context.rootPluginPath.empty())
    {
        result.detail = L"empty-root";
        return result;
    }

    if (context.fileSystem)
    {
        result.detail = L"provider";
        result.hr     = CollectProviderScopeTargets(context, scope, result.targets, cancelRequested);
        if (SUCCEEDED(result.hr))
        {
            SortBatchRenameTargets(result.targets);
            return result;
        }

        if (IsBatchRenameCancellationHRESULT(result.hr))
        {
            return result;
        }

        if (! IsLocalFileSystemContext(context))
        {
            return result;
        }

        result.targets.clear();
        result.detail = L"local-fallback";
    }

    result.hr = CollectLocalScopeTargets(context, scope, result.targets, cancelRequested);
    return result;
}

void EmitBatchRenameCollectMetrics(const std::chrono::steady_clock::time_point startedAt, const BatchRenameTargetCollectionResult& result) noexcept
{
    if (! Debug::Perf::IsCaptureEnabled())
    {
        return;
    }

    Debug::Perf::Emit(L"batchrename.collect.us", result.detail, Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(result.targets.size()), 0u, result.hr);
    Debug::Perf::EmitValue(L"batchrename.collect.targets", static_cast<uint64_t>(result.targets.size()), result.hr);
}

struct BatchRenameCollectionWork final
{
    HWND hwnd = nullptr;
    uint64_t generation = 0u;
    BatchRenamePaneContext context;
    BatchRenameScopeOptions scope;
    std::shared_ptr<std::atomic_bool> cancelRequested;
    std::unique_ptr<BatchRenameCollectionCompletedPayload> payload;
    wil::unique_hmodule modulePin;

    BatchRenameCollectionWork() = default;
    BatchRenameCollectionWork(const BatchRenameCollectionWork&) = delete;
    BatchRenameCollectionWork(BatchRenameCollectionWork&&) = delete;
    BatchRenameCollectionWork& operator=(const BatchRenameCollectionWork&) = delete;
    BatchRenameCollectionWork& operator=(BatchRenameCollectionWork&&) = delete;

    void Execute(PTP_CALLBACK_INSTANCE callbackInstance) noexcept
    {
        if (modulePin)
        {
            TransferModulePinToCallbackReturn(callbackInstance, modulePin);
        }
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            payload->hr = coinitHr;
        }
        else
        {
            const wil::unique_couninitialize_call coUninit;
            const auto startedAt = std::chrono::steady_clock::now();
            static std::mutex collectionWorkerMutex;
            std::scoped_lock lock(collectionWorkerMutex);
            BatchRenameTargetCollectionResult result = CollectBatchRenameTargets(context, scope, cancelRequested.get());
            EmitBatchRenameCollectMetrics(startedAt, result);
            context.fileSystem.reset();
            payload->hr      = result.hr;
            payload->detail  = std::move(result.detail);
            payload->targets = std::move(result.targets);
        }
        payload->generation = generation;
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kBatchRenameCompleted, kBatchRenameTaskCollection, std::move(payload)));
    }
};

[[nodiscard]] uint64_t StableRowIdFromPath(const std::filesystem::path& path) noexcept
{
    constexpr uint64_t kFnvOffset = 1469598103934665603ull;
    constexpr uint64_t kFnvPrime  = 1099511628211ull;

    uint64_t hash = kFnvOffset;
    for (const wchar_t ch : path.native())
    {
        const uint32_t codeUnit = static_cast<uint32_t>(ch);
        hash ^= static_cast<uint64_t>(codeUnit & 0x00FFu);
        hash *= kFnvPrime;
        hash ^= static_cast<uint64_t>((codeUnit >> 8u) & 0x00FFu);
        hash *= kFnvPrime;
    }
    return hash == 0u ? 1u : hash;
}

[[nodiscard]] int ResolvePreviewIconIndex(const BatchRename::PreviewRow& preview) noexcept
{
    auto& iconCache = IconCache::GetInstance();
    if (preview.isDirectory)
    {
        const std::wstring sourcePathText = preview.sourcePath.native();
        if (NavigationLocation::LooksLikeWindowsAbsolutePath(sourcePathText) && IconCache::IsSpecialFolder(sourcePathText))
        {
            const auto pathIcon = iconCache.QuerySysIconIndexForPath(sourcePathText.c_str(), FILE_ATTRIBUTE_DIRECTORY, true);
            if (pathIcon.has_value())
            {
                return pathIcon.value();
            }
        }

        const auto folderIcon = iconCache.GetOrQueryIconIndexByExtension(L"<directory>", FILE_ATTRIBUTE_DIRECTORY);
        return folderIcon.value_or(-1);
    }

    const std::wstring extension      = preview.sourcePath.extension().wstring();
    const std::wstring sourcePathText = preview.sourcePath.native();
    if (iconCache.RequiresPerFileLookup(extension) && NavigationLocation::LooksLikeWindowsAbsolutePath(sourcePathText))
    {
        const auto pathIcon = iconCache.QuerySysIconIndexForPath(sourcePathText.c_str(), FILE_ATTRIBUTE_NORMAL, true);
        if (pathIcon.has_value())
        {
            return pathIcon.value();
        }
    }

    const auto extensionIcon = iconCache.GetOrQueryIconIndexByExtension(extension, FILE_ATTRIBUTE_NORMAL);
    if (extensionIcon.has_value())
    {
        return extensionIcon.value();
    }

    if (NavigationLocation::LooksLikeWindowsAbsolutePath(sourcePathText))
    {
        const auto pathIcon = iconCache.QuerySysIconIndexForPath(sourcePathText.c_str(), FILE_ATTRIBUTE_NORMAL, true);
        if (pathIcon.has_value())
        {
            return pathIcon.value();
        }
    }

    return -1;
}

[[nodiscard]] std::wstring BuildPreviewIconText(const BatchPreviewRow& row)
{
    const wchar_t glyph = row.isDirectory ? FluentIcons::kFolder : FluentIcons::kDocument;
    return std::wstring(1u, glyph);
}

[[nodiscard]] std::wstring BuildPreviewStatusIconText(const BatchPreviewRow& row)
{
    if (row.hasErrorIssue)
    {
        return std::wstring(1u, FluentIcons::kError);
    }
    if (row.hasWarningIssue)
    {
        return std::wstring(1u, FluentIcons::kWarning);
    }
    return {};
}

[[nodiscard]] std::wstring LoadBatchRenameString(const unsigned int id, std::wstring fallback) noexcept
{
    std::wstring value = LoadStringResource(nullptr, id);
    return value.empty() ? std::move(fallback) : value;
}

// Maps a stable engine issue id to a localized, user-readable description.
// Unknown ids return an empty string so the raw id alone is shown.
[[nodiscard]] std::wstring LocalizeBatchRenameIssueId(std::wstring_view issueId)
{
    struct IssueText final
    {
        std::wstring_view id;
        unsigned int resourceId;
        std::wstring_view fallback;
    };

    static constexpr std::array<IssueText, 18> kIssueTexts = {{
        {L"name_empty", IDS_BATCH_RENAME_ISSUE_NAME_EMPTY, L"Name is empty"},
        {L"name_dot", IDS_BATCH_RENAME_ISSUE_NAME_DOT, L"Name is '.' or '..'"},
        {L"name_separator", IDS_BATCH_RENAME_ISSUE_NAME_SEPARATOR, L"Name contains a path separator"},
        {L"name_invalid_character", IDS_BATCH_RENAME_ISSUE_NAME_INVALID_CHARACTER, L"Name contains an invalid character"},
        {L"name_too_long", IDS_BATCH_RENAME_ISSUE_NAME_TOO_LONG, L"Name is too long"},
        {L"name_duplicate", IDS_BATCH_RENAME_ISSUE_NAME_DUPLICATE, L"Duplicate name in the same folder"},
        {L"name_unchanged", IDS_BATCH_RENAME_ISSUE_NAME_UNCHANGED, L"Name is unchanged"},
        {L"name_case_only", IDS_BATCH_RENAME_ISSUE_NAME_CASE_ONLY, L"Only letter case changes"},
        {L"name_edge_space_or_dot", IDS_BATCH_RENAME_ISSUE_NAME_EDGE_SPACE_OR_DOT, L"Name starts or ends with a space, or ends with a dot"},
        {L"name_destination_exists", IDS_BATCH_RENAME_ISSUE_NAME_DESTINATION_EXISTS, L"An item with the new name already exists"},
        {L"name_destination_probe_failed", IDS_BATCH_RENAME_ISSUE_NAME_DESTINATION_PROBE_FAILED, L"Could not check the new name destination"},
        {L"name_reserved_device", IDS_BATCH_RENAME_ISSUE_NAME_RESERVED_DEVICE, L"Name is a reserved device name"},
        {L"macro_unknown", IDS_BATCH_RENAME_ISSUE_MACRO_UNKNOWN, L"Unknown macro"},
        {L"macro_unclosed", IDS_BATCH_RENAME_ISSUE_MACRO_UNCLOSED, L"Unclosed macro brace"},
        {L"macro_invalid_format", IDS_BATCH_RENAME_ISSUE_MACRO_INVALID_FORMAT, L"Invalid macro format"},
        {L"manual_line_count", IDS_BATCH_RENAME_ISSUE_MANUAL_LINE_COUNT, L"Manual name count does not match the item count"},
        {L"regex_invalid", IDS_BATCH_RENAME_ISSUE_REGEX_INVALID, L"Invalid regular expression"},
        {L"provider_path_identity_unknown", IDS_BATCH_RENAME_ISSUE_PROVIDER_PATH_IDENTITY_UNKNOWN, L"Provider path identity is unavailable"},
    }};

    for (const IssueText& entry : kIssueTexts)
    {
        if (issueId == entry.id)
        {
            return LoadBatchRenameString(entry.resourceId, std::wstring(entry.fallback));
        }
    }

    if (issueId.starts_with(L"regex_"))
    {
        return LoadBatchRenameString(IDS_BATCH_RENAME_ISSUE_REGEX_ERROR, L"Regular expression error");
    }

    return {};
}

[[nodiscard]] std::wstring BuildPreviewIssueTooltip(const BatchRename::PreviewRow& preview)
{
    if (preview.issues.empty())
    {
        return {};
    }

    std::wstring tooltip = preview.newName;
    if (! tooltip.empty())
    {
        tooltip.append(L"\r\n");
    }

    for (size_t index = 0u; index < preview.issues.size(); ++index)
    {
        if (index != 0u)
        {
            tooltip.append(L", ");
        }

        // Keep the raw stable issue id visible next to the localized text so
        // reports and diagnostics can still reference the machine-readable id.
        const std::wstring& issueId  = preview.issues[index].message;
        const std::wstring localized = LocalizeBatchRenameIssueId(issueId);
        if (localized.empty())
        {
            tooltip.append(issueId);
        }
        else
        {
            tooltip.append(localized);
            tooltip.append(L" (");
            tooltip.append(issueId);
            tooltip.push_back(L')');
        }
    }
    return tooltip;
}

// Path-column text with the same semantics as the Find Files results grid: the containing
// folder relative to the window root (shown in the navigation bar), empty for items directly
// under the root. Paths outside the root keep their absolute parent folder so no location
// information is lost.
[[nodiscard]] std::wstring BuildPreviewDisplayPath(std::wstring_view rootText, const std::filesystem::path& sourcePath)
{
    const std::wstring parent = sourcePath.parent_path().native();

    std::wstring_view root = rootText;
    while (! root.empty() && (root.back() == L'\\' || root.back() == L'/'))
    {
        root.remove_suffix(1u);
    }

    if (root.empty() || ! OrdinalString::StartsWithNoCase(parent, root))
    {
        return parent;
    }

    if (parent.size() == root.size())
    {
        return {};
    }

    if (parent[root.size()] != L'\\' && parent[root.size()] != L'/')
    {
        return parent;
    }

    return parent.substr(root.size() + 1u);
}

[[nodiscard]] std::vector<BatchPreviewRow> BuildPreviewRowsFromPlan(const BatchRename::Plan& plan, std::wstring_view rootText)
{
    std::vector<BatchPreviewRow> rows;
    rows.reserve(plan.rows.size());
    for (const BatchRename::PreviewRow& preview : plan.rows)
    {
        BatchPreviewRow row{};
        row.stableId     = StableRowIdFromPath(preview.sourcePath);
        row.targetIndex  = preview.rowId == 0u ? 0u : static_cast<size_t>(preview.rowId - 1u);
        row.sourcePath   = preview.sourcePath;
        row.originalName = preview.originalName;
        row.newName      = preview.newName;
        row.sizeText     = preview.isDirectory || preview.metadataUnknown ? std::wstring{} : std::to_wstring(preview.sizeBytes);
        if (preview.lastWriteTime.has_value())
        {
            row.dateText = BatchRename::FormatDateText(preview.lastWriteTime.value());
            row.timeText = BatchRename::FormatTimeText(preview.lastWriteTime.value());
        }
        row.fullPath        = preview.sourcePath.native();
        row.displayPath     = BuildPreviewDisplayPath(rootText, preview.sourcePath);
        row.sizeBytes       = preview.sizeBytes;
        row.iconIndex       = ResolvePreviewIconIndex(preview);
        row.isDirectory     = preview.isDirectory;
        row.metadataUnknown = preview.metadataUnknown;
        row.hasErrorIssue   = BatchRename::HasIssueSeverity(preview, BatchRename::IssueSeverity::Error);
        row.hasWarningIssue = BatchRename::HasIssueSeverity(preview, BatchRename::IssueSeverity::Warning);
        row.changed         = preview.newName != preview.originalName;
        row.issueTooltip    = BuildPreviewIssueTooltip(preview);
        rows.push_back(std::move(row));
    }
    return rows;
}

void AppendBatchRenameTsvField(std::wstring& text, const std::wstring_view field)
{
    for (const wchar_t ch : field)
    {
        if (ch == L'\t' || ch == L'\r' || ch == L'\n')
        {
            text.push_back(L' ');
        }
        else
        {
            text.push_back(ch);
        }
    }
}

[[nodiscard]] std::wstring_view BatchPreviewRowClipboardField(const BatchPreviewRow& row, const size_t columnIndex) noexcept
{
    switch (columnIndex)
    {
        case 0u: return row.originalName;
        case 1u: return row.newName;
        case 2u: return row.sizeText;
        case 3u: return row.dateText;
        case 4u: return row.timeText;
        case 5u: return row.fullPath;
        default: return {};
    }
}

[[nodiscard]] std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout)
{
    std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const Common::Settings::GridColumnLayoutEntry& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(RedSalamander::DxUi::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = entry.displayIndex,
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

[[nodiscard]] std::vector<Common::Settings::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<RedSalamander::DxUi::GridColumnLayoutEntry>& layout)
{
    std::vector<Common::Settings::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const RedSalamander::DxUi::GridColumnLayoutEntry& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = static_cast<uint32_t>(entry.displayIndex),
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

void UpdateRecentBatchRenameValue(std::vector<std::wstring>& history, std::wstring value) noexcept
{
    if (value.empty())
    {
        return;
    }

    const auto it =
        std::find_if(history.begin(), history.end(), [&](const std::wstring& existing) noexcept { return OrdinalString::EqualsNoCase(existing, value); });
    if (it != history.end())
    {
        history.erase(it);
    }

    history.insert(history.begin(), std::move(value));
    if (history.size() > kMaxRecentBatchRenameEntries)
    {
        history.resize(kMaxRecentBatchRenameEntries);
    }
}

[[nodiscard]] std::vector<std::wstring> SplitManualNames(std::wstring_view text)
{
    std::vector<std::wstring> names;
    std::wstring current;
    for (const wchar_t ch : text)
    {
        if (ch == L'\r')
        {
            continue;
        }
        if (ch == L'\n')
        {
            names.push_back(std::move(current));
            current.clear();
            continue;
        }
        current.push_back(ch);
    }
    names.push_back(std::move(current));
    return names;
}

[[nodiscard]] std::wstring JoinManualNames(const std::vector<std::wstring>& names)
{
    std::wstring text;
    for (size_t index = 0u; index < names.size(); ++index)
    {
        if (index != 0u)
        {
            text.push_back(L'\n');
        }
        text.append(names[index]);
    }
    return text;
}

[[nodiscard]] std::array<ComboBox::Item, 4> BuildCaseComboItems()
{
    return {
        ComboBox::Item{.value = L"none", .display = LoadBatchRenameString(IDS_BATCH_RENAME_CASE_NONE, L"Do not change")},
        ComboBox::Item{.value = L"lower", .display = LoadBatchRenameString(IDS_BATCH_RENAME_CASE_LOWER, L"Lower case")},
        ComboBox::Item{.value = L"upper", .display = LoadBatchRenameString(IDS_BATCH_RENAME_CASE_UPPER, L"Upper case")},
        ComboBox::Item{.value = L"mixed", .display = LoadBatchRenameString(IDS_BATCH_RENAME_CASE_MIXED, L"Mixed case")},
    };
}

[[nodiscard]] size_t CaseTransformToIndex(const BatchRename::CaseTransform transform) noexcept
{
    switch (transform)
    {
        case BatchRename::CaseTransform::None: return 0u;
        case BatchRename::CaseTransform::Lower: return 1u;
        case BatchRename::CaseTransform::Upper: return 2u;
        case BatchRename::CaseTransform::Mixed: return 3u;
        default: return 0u;
    }
}

[[nodiscard]] BatchRename::CaseTransform CaseTransformFromIndex(const size_t index) noexcept
{
    switch (index)
    {
        case 1u: return BatchRename::CaseTransform::Lower;
        case 2u: return BatchRename::CaseTransform::Upper;
        case 3u: return BatchRename::CaseTransform::Mixed;
        default: return BatchRename::CaseTransform::None;
    }
}

[[nodiscard]] BatchRename::CaseTransform CaseTransformFromSettings(const Common::Settings::BatchRenameCaseStyle style) noexcept
{
    switch (style)
    {
        case Common::Settings::BatchRenameCaseStyle::Lower: return BatchRename::CaseTransform::Lower;
        case Common::Settings::BatchRenameCaseStyle::Upper: return BatchRename::CaseTransform::Upper;
        case Common::Settings::BatchRenameCaseStyle::Mixed: return BatchRename::CaseTransform::Mixed;
        case Common::Settings::BatchRenameCaseStyle::None:
        default: return BatchRename::CaseTransform::None;
    }
}

[[nodiscard]] Common::Settings::BatchRenameCaseStyle CaseTransformToSettings(const BatchRename::CaseTransform transform) noexcept
{
    switch (transform)
    {
        case BatchRename::CaseTransform::Lower: return Common::Settings::BatchRenameCaseStyle::Lower;
        case BatchRename::CaseTransform::Upper: return Common::Settings::BatchRenameCaseStyle::Upper;
        case BatchRename::CaseTransform::Mixed: return Common::Settings::BatchRenameCaseStyle::Mixed;
        case BatchRename::CaseTransform::None:
        default: return Common::Settings::BatchRenameCaseStyle::None;
    }
}

[[nodiscard]] std::wstring ResolveContextRootText(const BatchRenamePaneContext& context) noexcept
{
    if (! context.rootPluginPath.empty())
    {
        return context.rootPluginPath.native();
    }
    if (! context.initialPaths.empty())
    {
        const std::filesystem::path parent = context.initialPaths.front().parent_path();
        if (! parent.empty())
        {
            return parent.native();
        }
    }
    return {};
}

[[nodiscard]] size_t CountVisibleChildWindows(HWND hwnd) noexcept
{
    size_t count = 0u;
    EnumChildWindows(hwnd,
                     [](HWND child, LPARAM param) noexcept -> BOOL
    {
        auto* out = reinterpret_cast<size_t*>(param);
        if (out && IsWindowVisible(child) != FALSE)
        {
            ++(*out);
        }
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&count));
    return count;
}

#ifdef ENABLE_TESTS
[[nodiscard]] std::wstring ResolveBatchRenameDebugAccessibleName(const Control& control)
{
    if (const std::wstring_view explicitName = control.GetAccessibleName(); ! explicitName.empty())
    {
        return std::wstring(explicitName);
    }

    if (const auto* toggle = dynamic_cast<const Toggle*>(&control))
    {
        return std::wstring(toggle->GetDisplayedText());
    }
    if (const auto* button = dynamic_cast<const Button*>(&control))
    {
        return std::wstring(button->GetText());
    }

    return {};
}

void CollectBatchRenameFocusableAccessibleNames(const Control* control, std::vector<std::wstring>& names)
{
    if (! control || ! control->IsVisible() || ! control->IsEnabled())
    {
        return;
    }

    if (control->IsFocusable())
    {
        names.push_back(ResolveBatchRenameDebugAccessibleName(*control));
    }

    if (const auto* panel = dynamic_cast<const Panel*>(control))
    {
        for (const std::unique_ptr<Control>& child : panel->GetChildren())
        {
            CollectBatchRenameFocusableAccessibleNames(child.get(), names);
        }
    }
}
#endif

[[nodiscard]] HWND NormalizeBatchRenameOwner(HWND owner) noexcept
{
    if (! owner || IsWindow(owner) == FALSE)
    {
        return nullptr;
    }
    return GetAncestor(owner, GA_ROOT);
}

// Sorts preview rows for the supplied grid sort spec. Shared by the visible
// grid model and full-plan consumers such as Manual `Sort like preview`, so
// ordering stays consistent regardless of the Hide-unchanged filter.
void SortBatchPreviewRows(std::vector<BatchPreviewRow>& rows, const std::vector<GridColumnDesc>& columns, const GridSortSpec& sortSpec)
{
    if (rows.empty() || sortSpec.direction == SortDirection::None || sortSpec.columnIndex >= columns.size() || ! columns[sortSpec.columnIndex].sortable)
    {
        return;
    }

    const auto cellText = [](const BatchPreviewRow& row, const size_t columnIndex) noexcept -> std::wstring_view
    {
        switch (columnIndex)
        {
            case 0u: return row.originalName;
            case 1u: return row.newName;
            case 2u: return row.sizeText;
            case 3u: return row.dateText;
            case 4u: return row.timeText;
            case 5u: return row.displayPath;
            default: return {};
        }
    };

    std::stable_sort(rows.begin(),
                     rows.end(),
                     [&](const BatchPreviewRow& lhs, const BatchPreviewRow& rhs) noexcept
    {
        int comparison = 0;
        if (sortSpec.columnIndex == 2u)
        {
            comparison = lhs.sizeBytes < rhs.sizeBytes ? -1 : (lhs.sizeBytes > rhs.sizeBytes ? 1 : 0);
        }
        else
        {
            comparison = OrdinalString::Compare(cellText(lhs, sortSpec.columnIndex), cellText(rhs, sortSpec.columnIndex), true);
        }

        if (comparison == 0)
        {
            return false;
        }
        return sortSpec.direction == SortDirection::Ascending ? comparison < 0 : comparison > 0;
    });
}

class BatchRenamePreviewGridModel final : public IDxGridModel
{
public:
    BatchRenamePreviewGridModel()
    {
        _columns = {
            MakeColumn(L"original", IDS_BATCH_RENAME_COL_ORIGINAL_NAME, L"Original Name", 220.0f),
            MakeColumn(L"new", IDS_BATCH_RENAME_COL_NEW_NAME, L"New Name", 220.0f),
            MakeColumn(L"size", IDS_BATCH_RENAME_COL_SIZE, L"Size", 92.0f, DWRITE_TEXT_ALIGNMENT_TRAILING),
            MakeColumn(L"date", IDS_BATCH_RENAME_COL_DATE, L"Date", 96.0f, DWRITE_TEXT_ALIGNMENT_TRAILING),
            MakeColumn(L"time", IDS_BATCH_RENAME_COL_TIME, L"Time", 76.0f, DWRITE_TEXT_ALIGNMENT_TRAILING),
            MakeColumn(L"path", IDS_BATCH_RENAME_COL_PATH, L"Path", 300.0f),
        };
    }

    void SetRows(std::vector<BatchPreviewRow> rows)
    {
        _rows = std::move(rows);
    }

    void SortRows(const GridSortSpec& sortSpec)
    {
        SortBatchPreviewRows(_rows, _columns, sortSpec);
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(const size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(const size_t rowIndex, const size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size() || columnIndex >= _columns.size())
        {
            return;
        }

        const BatchPreviewRow& row = _rows[rowIndex];
        switch (columnIndex)
        {
            case 0:
                outCell.kind      = GridCellKind::IconText;
                outCell.iconText  = BuildPreviewIconText(row);
                outCell.iconIndex = row.iconIndex;
                outCell.text      = row.originalName;
                break;
            case 1:
                outCell.text = row.newName;
                if (row.hasErrorIssue || row.hasWarningIssue)
                {
                    outCell.kind     = GridCellKind::IconText;
                    outCell.iconText = BuildPreviewStatusIconText(row);
                }
                break;
            case 2: outCell.text = row.sizeText; break;
            case 3: outCell.text = row.dateText; break;
            case 4: outCell.text = row.timeText; break;
            case 5: outCell.text = row.displayPath; break;
            default: break;
        }
        outCell.tooltipText = outCell.text;
        if (columnIndex == 1u && ! row.issueTooltip.empty())
        {
            outCell.tooltipText = row.issueTooltip;
        }
        if (columnIndex == 5u)
        {
            // The cell shows the root-relative containing folder; keep the absolute path
            // reachable through the tooltip.
            outCell.tooltipText = row.fullPath;
        }
        outCell.textAlignment = _columns[columnIndex].textAlignment;
        outCell.multiline     = false;
    }

    [[nodiscard]] GridRowStyle GetRowStyle(const size_t rowIndex) const override
    {
        GridRowStyle style{};
        if (rowIndex >= _rows.size())
        {
            return style;
        }

        const BatchPreviewRow& row = _rows[rowIndex];
        if (row.hasErrorIssue)
        {
            style.tone = GridRowTone::Error;
        }
        else if (row.hasWarningIssue)
        {
            style.tone = GridRowTone::Warning;
        }
        return style;
    }

    [[nodiscard]] uint64_t GetStableRowId(const size_t rowIndex) const noexcept override
    {
        return rowIndex < _rows.size() ? _rows[rowIndex].stableId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(const uint64_t rowId) const noexcept override
    {
        for (size_t index = 0u; index < _rows.size(); ++index)
        {
            if (_rows[index].stableId == rowId)
            {
                return index;
            }
        }
        return std::nullopt;
    }

    [[nodiscard]] const std::vector<BatchPreviewRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] const std::vector<GridColumnDesc>& GetColumns() const noexcept
    {
        return _columns;
    }

private:
    [[nodiscard]] static GridColumnDesc MakeColumn(std::wstring id,
                                                   const unsigned int titleId,
                                                   std::wstring fallbackTitle,
                                                   const float widthDip,
                                                   const DWRITE_TEXT_ALIGNMENT alignment = DWRITE_TEXT_ALIGNMENT_LEADING)
    {
        GridColumnDesc column{};
        column.id            = std::move(id);
        column.title         = LoadBatchRenameString(titleId, std::move(fallbackTitle));
        column.widthDip      = widthDip;
        column.minWidthDip   = 64.0f;
        column.kind          = GridColumnKind::Text;
        column.sortable      = true;
        column.multiline     = false;
        column.textAlignment = alignment;
        return column;
    }

    std::vector<GridColumnDesc> _columns;
    std::vector<BatchPreviewRow> _rows;
};

class BatchRenameWindow final : public IDxGridDelegate
{
public:
    using IDxGridDelegate::OnGridContextMenu;
    using IDxGridDelegate::OnGridRowActivated;

    BatchRenameWindow() = default;

    ~BatchRenameWindow() noexcept override
    {
        // Lets Create() observe a self-delete performed by the window
        // procedure while CreateWindowExW was still on the stack.
        if (_destructionObserver)
        {
            *_destructionObserver = true;
        }
    }

    BatchRenameWindow(const BatchRenameWindow&)            = delete;
    BatchRenameWindow& operator=(const BatchRenameWindow&) = delete;
    BatchRenameWindow(BatchRenameWindow&&)                 = delete;
    BatchRenameWindow& operator=(BatchRenameWindow&&)      = delete;

    [[nodiscard]] HWND Create(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, BatchRenamePaneContext context) noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;
    void SetContext(BatchRenamePaneContext context) noexcept;
    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugGetSnapshot(BatchRenameDebugSnapshot& out) const noexcept;
    void DebugSetRules(BatchRename::Rules rules) noexcept;
    [[nodiscard]] bool DebugSetRuleControls(const BatchRename::Rules& rules) noexcept;
    [[nodiscard]] bool DebugSetScope(std::wstring_view mask, bool includeSubdirectories, bool includeFiles, bool includeFolders) noexcept;
    [[nodiscard]] bool DebugSetPreviewSort(std::wstring_view columnId, bool descending) noexcept;
    [[nodiscard]] bool DebugSwitchMode(BatchRename::Mode mode) noexcept;
    [[nodiscard]] bool DebugSetManualText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugClickManualPaste() noexcept;
    [[nodiscard]] bool DebugClickManualSortLikePreview() noexcept;
    [[nodiscard]] bool DebugSetHideUnchanged(bool hideUnchanged) noexcept;
    [[nodiscard]] bool DebugReorderPreviewColumn(std::wstring_view columnId, size_t targetDisplayIndex) noexcept;
    [[nodiscard]] TextField* DebugResolveRuleField(BatchRenameDebugRuleField field) noexcept;
    [[nodiscard]] bool DebugSetRuleFieldSelection(BatchRenameDebugRuleField field, size_t selectionStart, size_t selectionEnd) noexcept;
    [[nodiscard]] bool DebugSetRuleFieldText(BatchRenameDebugRuleField field, std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugInsertHelperCommand(BatchRenameDebugRuleField field, int commandId) noexcept;
    [[nodiscard]] bool DebugCopyPreview(BatchRenameDebugPreviewCopyKind kind, size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugRevealPreview(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugActivatePreview(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugCopyExecutionReport() const noexcept;
    [[nodiscard]] bool DebugCopyUndoPlan() const noexcept;
    [[nodiscard]] bool DebugFlushPendingPreview() noexcept;
    [[nodiscard]] HRESULT DebugExecute() noexcept;
    [[nodiscard]] HRESULT DebugStartExecution() noexcept;
    [[nodiscard]] HRESULT DebugWaitExecutionIdle() noexcept;
    [[nodiscard]] bool DebugInjectStaleCollectionPayload(std::filesystem::path sourcePath) noexcept;
    [[nodiscard]] bool DebugInjectStaleExecutionPayload(std::filesystem::path sourcePath, std::filesystem::path targetPath) noexcept;
    void DebugPumpWhileTasksActive(bool waitForExecution) noexcept;
#endif

private:
    [[nodiscard]] static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept;
    void OnNcDestroy(HWND hwnd) noexcept;
    void BuildUi();
    [[nodiscard]] bool CreateRootNavigation(HWND parent) noexcept;
    void ApplyTheme() noexcept;
    void Layout() noexcept;
    void ApplySettingsDefaults() noexcept;
    void ApplyPreviewGridSettings() noexcept;
    void PersistCloseState() noexcept;
    void PersistUiState(bool updateHistory) noexcept;
    void RebuildTargetsFromScope();
    void RequestPreviewRebuild(bool rebuildTargets) noexcept;
    void CancelPendingPreviewRebuild() noexcept;
    void OnPreviewRebuildTimer() noexcept;
    void RebuildPreview() noexcept;
    void OnPreviewCompleted(std::unique_ptr<BatchRenamePreviewCompletedPayload> payload) noexcept;
    void SyncRuleControls() noexcept;
    void UpdateModeVisibility() noexcept;
    void SwitchMode(BatchRename::Mode mode);
    void SeedManualTextFromRulePreview();
    void SetManualText(std::wstring text);
    void FillManualFromPreview();
    void ClearManualText();
    [[nodiscard]] bool PasteManualFromClipboard();
    [[nodiscard]] bool SortManualTextLikePreview();
    void SetHideUnchangedRows(bool hideUnchanged) noexcept;
    [[nodiscard]] HRESULT ExecuteRename() noexcept;
    void StartTargetCollection() noexcept;
    void OnCollectionCompleted(std::unique_ptr<BatchRenameCollectionCompletedPayload> payload) noexcept;
    void OnExecutionCompleted(std::unique_ptr<BatchRenameExecutionCompletedPayload> payload) noexcept;
    void OnTaskProgress(std::unique_ptr<BatchRenameTaskProgressPayload> payload) noexcept;
    void RequestTaskCancel() noexcept;
    void CancelAndJoinBackgroundTask() noexcept;
    void UpdateTaskUi() noexcept;
    void ClearExecutionReport() noexcept;
    void StoreExecutionReport(BatchRenameExecutionReport report) noexcept;
    void ShowHelperMenu(BatchRenameMenus::HelperMenuKind kind, TextField* targetField, Button* anchorButton) noexcept;
    [[nodiscard]] bool InsertHelperCommand(TextField& targetField, int commandId) noexcept;
    void ShowPreviewContextMenu(size_t rowIndex, POINT screenPoint) noexcept;
    [[nodiscard]] bool DispatchPreviewContextMenuCommand(int commandId, size_t rowIndex) noexcept;
    [[nodiscard]] bool ActivatePreviewRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool RevealPreviewRowInActivePane(size_t rowIndex) noexcept;
    [[nodiscard]] bool CopyPreviewRowsToClipboard() const noexcept;
    [[nodiscard]] bool CopyPreviewRowFieldToClipboard(size_t rowIndex, int commandId) const noexcept;
    [[nodiscard]] bool CopyExecutionReportToClipboard() const noexcept;
    [[nodiscard]] bool CopyUndoPlanToClipboard() const noexcept;
    [[nodiscard]] std::wstring BuildPreviewRowsClipboardText() const;
    [[nodiscard]] std::wstring BuildExecutionReportClipboardText() const;
    [[nodiscard]] std::wstring BuildUndoPlanClipboardText() const;
    void RebuildPreviewFromRuleControlChange();
    void RefreshPreviewPresentation() noexcept;
    void RefreshVisibleRows() noexcept;
    void UpdateStatus() noexcept;
    void RefreshRootNavigationPath() noexcept;
    void RefreshRootNavigationHistory() noexcept;
    void OnRootNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept;
    [[nodiscard]] std::optional<size_t> FindPreviewColumnIndexById(std::wstring_view columnId) const noexcept;
    void OnGridSortRequested(const GridSortSpec& sortSpec) override;
    void OnGridRowActivated(Grid& sender, size_t rowIndex) override;
    void OnGridContextMenu(Grid& sender, size_t rowIndex, POINT screenPoint) override;
    [[nodiscard]] wil::com_ptr<ID2D1Bitmap1> GetGridIconBitmap(const Grid& sender, int iconIndex, float targetDipSize, ID2D1DeviceContext* d2dContext) override;

    wil::unique_hwnd _hWnd;
    HINSTANCE _instance                   = nullptr;
    HWND _ownerWindow                     = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    BatchRenamePaneContext _context;
    BatchRenameScopeOptions _scopeOptions;
    std::vector<BatchRename::Target> _targets;
    BatchRename::Rules _rules;
    std::vector<BatchPreviewRow> _fullPreviewRows;
    std::optional<BatchRename::Plan> _currentPlan;
    BatchRename::Stats _previewStats{};
    std::optional<BatchRenameExecutionReport> _lastExecutionReport;
    std::wstring _rootText;
    size_t _dispatchDepth = 0u;
    std::atomic_bool _cancelRequested{false};
    std::atomic<uint64_t> _taskGeneration{0u};
    std::atomic<uint64_t> _previewGeneration{0u};
    std::shared_ptr<std::atomic_bool> _collectionCancelFlag;
    std::jthread _taskWorker;
    bool* _destructionObserver       = nullptr;
    bool _collecting                 = false;
    bool _executing                  = false;
    bool _previewing                 = false;
    bool _collectionQueued           = false;
    HRESULT _lastExecutionTerminalHr = S_OK;
    bool _deletePending              = false;
    bool _uiStatePersisted           = false;

    NavigationView _rootNavigation;
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root                          = nullptr;
    Label* _titleLabel                    = nullptr;
    Label* _rootLabel                     = nullptr;
    Label* _maskLabel                     = nullptr;
    TextField* _maskField                 = nullptr;
    Checkbox* _includeSubdirectoriesCheck = nullptr;
    Checkbox* _includeFilesCheck          = nullptr;
    Checkbox* _includeFoldersCheck        = nullptr;
    RadioButtons* _modeSelector           = nullptr;
    RadioButton* _rulesModeButton         = nullptr;
    RadioButton* _manualModeButton        = nullptr;
    Label* _newNameLabel                  = nullptr;
    Label* _searchForLabel                = nullptr;
    Label* _replaceWithLabel              = nullptr;
    Label* _fileNameCaseLabel             = nullptr;
    Label* _extensionCaseLabel            = nullptr;
    TextField* _nameTemplateField         = nullptr;
    TextField* _searchForField            = nullptr;
    TextField* _replaceWithField          = nullptr;
    Button* _nameTemplateHelperButton     = nullptr;
    Button* _searchForHelperButton        = nullptr;
    Button* _replaceWithHelperButton      = nullptr;
    TextField* _manualNamesField          = nullptr;
    Button* _manualFillButton             = nullptr;
    Button* _manualClearButton            = nullptr;
    Button* _manualPasteButton            = nullptr;
    Button* _manualSortLikePreviewButton  = nullptr;
    Checkbox* _regexCheck                 = nullptr;
    Checkbox* _caseSensitiveCheck         = nullptr;
    Checkbox* _wholeWordsCheck            = nullptr;
    Checkbox* _replaceOnceCheck           = nullptr;
    Checkbox* _excludeExtensionCheck      = nullptr;
    ComboBox* _fileNameCaseCombo          = nullptr;
    ComboBox* _extensionCaseCombo         = nullptr;
    Grid* _grid                           = nullptr;
    StatusStrip* _status                  = nullptr;
    Checkbox* _hideUnchangedCheck         = nullptr;
    Button* _cancelButton                 = nullptr;
    Button* _renameButton                 = nullptr;
    std::unique_ptr<BatchRenamePreviewGridModel> _gridModelStorage;
    BatchRenamePreviewGridModel* _gridModel = nullptr;
    bool _syncingRuleControls               = false;
    bool _manualTextInitialized             = false;
    bool _hideUnchangedRows                 = false;
    bool _previewRebuildPending             = false;
    bool _targetRebuildPending              = false;
};

BatchRenameWindow* g_batchRenameWindow = nullptr;

ATOM BatchRenameWindow::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = BatchRenameWindow::WndProcThunk;
    wc.hInstance     = instance;
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.lpszClassName = kBatchRenameWindowClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND BatchRenameWindow::Create(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, BatchRenamePaneContext context) noexcept
{
    // Create() owns failure cleanup: on any failure path the instance is
    // deleted exactly once, either here or by the window procedure tearing
    // down a half-created window. Callers must not delete after Create().
    _instance = GetModuleHandleW(nullptr);
    if (! RegisterWndClass(_instance))
    {
        delete this;
        return nullptr;
    }

    _settings    = &settings;
    _theme       = theme;
    _context     = std::move(context);
    _rootText    = ResolveContextRootText(_context);
    _ownerWindow = NormalizeBatchRenameOwner(owner);
    ApplySettingsDefaults();
    _targets.clear();

    const UINT dpi           = _ownerWindow ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
    const int scaledWidth    = MulDiv(1040, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
    const int scaledHeight   = MulDiv(680, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
    const std::wstring title = LoadBatchRenameString(IDS_BATCH_RENAME_TITLE, L"Batch Rename");

    bool destroyedDuringCreate = false;
    _destructionObserver       = &destroyedDuringCreate;
    const HWND hwnd            = CreateWindowExW(0,
                                                 kBatchRenameWindowClassName,
                                                 title.c_str(),
                                                 WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                                 CW_USEDEFAULT,
                                                 CW_USEDEFAULT,
                                                 scaledWidth,
                                                 scaledHeight,
                                                 nullptr,
                                                 nullptr,
                                                 _instance,
                                                 this);
    if (destroyedDuringCreate)
    {
        // WM_CREATE failed: CreateWindowExW destroyed the half-created window
        // and the window procedure already deleted this instance.
        return nullptr;
    }
    _destructionObserver = nullptr;
    if (! hwnd)
    {
        delete this;
        return nullptr;
    }

    if (! _hWnd)
    {
        _hWnd.reset(hwnd);
    }

    const bool hasPlacement = _settings && _settings->windows.contains(std::wstring(kBatchRenameWindowId));
    const int showCmd       = hasPlacement ? WindowPlacementPersistence::Restore(*_settings, kBatchRenameWindowId, hwnd) : SW_SHOWNORMAL;
    ShowWindow(hwnd, showCmd);
    SetForegroundWindow(hwnd);
    StartTargetCollection();
    return hwnd;
}

void BatchRenameWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    _dxHost.Invalidate();
}

void BatchRenameWindow::SetContext(BatchRenamePaneContext context) noexcept
{
    CancelAndJoinBackgroundTask();
    ClearExecutionReport();
    _context = std::move(context);
    _targets.clear();
    _rootText = ResolveContextRootText(_context);
    if (_rootLabel)
    {
        _rootLabel->SetText(LoadBatchRenameString(IDS_FIND_LABEL_ROOT, L"Look in:"));
    }
    _rootNavigation.SetFileSystem(_context.fileSystem);
    RefreshRootNavigationPath();
    RefreshRootNavigationHistory();
    RebuildPreview();
    Layout();
    _dxHost.Invalidate();
    StartTargetCollection();
}

LRESULT CALLBACK BatchRenameWindow::WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* self         = static_cast<BatchRenameWindow*>(create ? create->lpCreateParams : nullptr);
        if (! self)
        {
            return FALSE;
        }

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        if (! self->_hWnd)
        {
            self->_hWnd.reset(hwnd);
        }
        g_batchRenameWindow = self;
    }

    auto* self = reinterpret_cast<BatchRenameWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    ++self->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([self]() noexcept
    {
        if (self->_dispatchDepth > 0u)
        {
            --self->_dispatchDepth;
        }
        if (self->_dispatchDepth == 0u && self->_deletePending)
        {
            delete self;
        }
    });

    return self->WndProc(hwnd, message, wParam, lParam);
}

LRESULT BatchRenameWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_CLOSE || message == WM_NCDESTROY)
    {
        PersistCloseState();
    }

    bool handled     = false;
    LRESULT dxResult = 0;
    if (message != WM_CREATE)
    {
        dxResult = _dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
    }

    if (handled)
    {
        if (message == WM_NCDESTROY)
        {
            OnNcDestroy(hwnd);
        }
        else if (message == WM_SIZE || message == WM_DPICHANGED)
        {
            if (message == WM_DPICHANGED && _rootNavigation.GetHwnd())
            {
                // The hosted NavigationView is a native child that re-renders fonts/icons from its
                // own DPI state; forward the change explicitly like FolderWindow does for its panes
                // (the system's WM_DPICHANGED_AFTERPARENT does not reliably refresh it).
                _rootNavigation.OnDpiChanged(static_cast<float>(HIWORD(wParam)));
            }
            Layout();
        }
        return dxResult;
    }

    switch (message)
    {
        case WM_CREATE: InitPostedPayloadWindow(hwnd); return OnCreate(hwnd) ? 0 : -1;
        case WM_SIZE: Layout(); return 0;
        // WM_DPICHANGED and WM_KEYDOWN are fully handled by the DxUi WindowHost
        // (which marks them handled before this switch), so they intentionally
        // have no local cases here. Escape closes the window through the
        // host's OnEscape callback wired in OnCreate.
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
            if (info)
            {
                const UINT dpi         = GetDpiForWindow(hwnd);
                info->ptMinTrackSize.x = std::max<LONG>(info->ptMinTrackSize.x, MulDiv(760, static_cast<int>(dpi), 96));
                info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, MulDiv(460, static_cast<int>(dpi), 96));
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
        case WM_ACTIVATE: ApplyTitleBarTheme(hwnd, _theme, LOWORD(wParam) != WA_INACTIVE); return 0;
        case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
        case WM_TIMER:
            if (wParam == kPreviewRebuildTimerId)
            {
                OnPreviewRebuildTimer();
                return 0;
            }
            break;
        case WndMsg::kBatchRenameTaskUpdate: OnTaskProgress(TakeMessagePayload<BatchRenameTaskProgressPayload>(lParam)); return 0;
        case WndMsg::kBatchRenameCompleted:
            if (wParam == kBatchRenameTaskCollection)
            {
                OnCollectionCompleted(TakeMessagePayload<BatchRenameCollectionCompletedPayload>(lParam));
            }
            else if (wParam == kBatchRenameTaskPreview)
            {
                OnPreviewCompleted(TakeMessagePayload<BatchRenamePreviewCompletedPayload>(lParam));
            }
            else
            {
                OnExecutionCompleted(TakeMessagePayload<BatchRenameExecutionCompletedPayload>(lParam));
            }
            return 0;
        case WM_CLOSE:
            CancelAndJoinBackgroundTask();
            DestroyWindow(hwnd);
            return 0;
        case WM_NCDESTROY: OnNcDestroy(hwnd); break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

bool BatchRenameWindow::OnCreate(HWND hwnd) noexcept
{
    if (! _dxHost.Attach(hwnd))
    {
        return false;
    }
    _dxHost.SetOnEscape([this]() noexcept
    {
        const HWND hwnd = _hWnd.get();
        return hwnd && IsWindow(hwnd) != FALSE && PostMessageW(hwnd, WM_CLOSE, 0, 0) != FALSE;
    });

    BuildUi();
    if (! CreateRootNavigation(hwnd))
    {
        Debug::Error(L"BatchRename: failed to create root navigation bar.");
        return false;
    }
    ApplyTheme();
    RebuildPreview();
    Layout();
    return true;
}

void BatchRenameWindow::OnNcDestroy(HWND hwnd) noexcept
{
    static_cast<void>(_taskGeneration.fetch_add(1u, std::memory_order_acq_rel));
    static_cast<void>(_previewGeneration.fetch_add(1u, std::memory_order_acq_rel));
    _previewing = false;
    CancelAndJoinBackgroundTask();
    PersistCloseState();
    CancelPendingPreviewRebuild();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _gridModel = nullptr;
    _gridModelStorage.reset();
    _renameButton                = nullptr;
    _cancelButton                = nullptr;
    _status                      = nullptr;
    _grid                        = nullptr;
    _extensionCaseCombo          = nullptr;
    _fileNameCaseCombo           = nullptr;
    _excludeExtensionCheck       = nullptr;
    _replaceOnceCheck            = nullptr;
    _wholeWordsCheck             = nullptr;
    _caseSensitiveCheck          = nullptr;
    _regexCheck                  = nullptr;
    _replaceWithHelperButton     = nullptr;
    _searchForHelperButton       = nullptr;
    _nameTemplateHelperButton    = nullptr;
    _replaceWithField            = nullptr;
    _searchForField              = nullptr;
    _nameTemplateField           = nullptr;
    _manualClearButton           = nullptr;
    _manualFillButton            = nullptr;
    _manualPasteButton           = nullptr;
    _manualSortLikePreviewButton = nullptr;
    _manualNamesField            = nullptr;
    _extensionCaseLabel          = nullptr;
    _fileNameCaseLabel           = nullptr;
    _replaceWithLabel            = nullptr;
    _searchForLabel              = nullptr;
    _newNameLabel                = nullptr;
    _manualModeButton            = nullptr;
    _rulesModeButton             = nullptr;
    _modeSelector                = nullptr;
    _rootLabel                   = nullptr;
    _includeFoldersCheck         = nullptr;
    _includeFilesCheck           = nullptr;
    _includeSubdirectoriesCheck  = nullptr;
    _maskField                   = nullptr;
    _maskLabel                   = nullptr;
    _titleLabel                  = nullptr;
    _root                        = nullptr;
    _rootNavigation.Destroy();
    _rootStorage.reset();
    _dxHost.Detach();

    if (_hWnd.get() == hwnd)
    {
        _hWnd.release();
    }
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    if (g_batchRenameWindow == this)
    {
        g_batchRenameWindow = nullptr;
    }

    const HWND restoreOwner = (_ownerWindow && IsWindow(_ownerWindow) != FALSE) ? _ownerWindow : nullptr;
    _settings               = nullptr;
    _ownerWindow            = nullptr;
    _deletePending          = true;
    if (_dispatchDepth == 0u)
    {
        delete this;
    }

    if (restoreOwner)
    {
        static_cast<void>(SetActiveWindow(restoreOwner));
    }
}

void BatchRenameWindow::BuildUi()
{
    if (_root)
    {
        return;
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();

    _titleLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_TITLE, L"Batch Rename"));
    _titleLabel->SetMultiline(false);

    _rootLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_FIND_LABEL_ROOT, L"Look in:"));
    _rootLabel->SetFontRole(RedSalamander::DxUi::FontRole::Small);
    _rootLabel->SetMultiline(false);

    _maskLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_MASK, L"Mask:"));
    _maskLabel->SetMultiline(false);

    _maskField = _root->AddChild<TextField>(_scopeOptions.mask);
    _maskField->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_MASK, L"Mask:"));
    _maskField->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _scopeOptions.mask = std::wstring(text);
        RequestPreviewRebuild(true);
    });

    _includeSubdirectoriesCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_INCLUDE_SUBDIRECTORIES, L"Include subdirectories"));
    _includeSubdirectoriesCheck->SetChecked(_scopeOptions.includeSubdirectories);
    _includeSubdirectoriesCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _scopeOptions.includeSubdirectories = checked;
        RebuildTargetsFromScope();
    });

    _includeFilesCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_INCLUDE_FILES, L"Files"));
    _includeFilesCheck->SetChecked(_scopeOptions.includeFiles);
    _includeFilesCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _scopeOptions.includeFiles = checked;
        RebuildTargetsFromScope();
    });

    _includeFoldersCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_INCLUDE_FOLDERS, L"Folders"));
    _includeFoldersCheck->SetChecked(_scopeOptions.includeFolders);
    _includeFoldersCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _scopeOptions.includeFolders = checked;
        RebuildTargetsFromScope();
    });

    _modeSelector     = _root->AddChild<RadioButtons>();
    _rulesModeButton  = _modeSelector->AddItem(LoadBatchRenameString(IDS_BATCH_RENAME_MODE_RULES, L"Rules"));
    _manualModeButton = _modeSelector->AddItem(LoadBatchRenameString(IDS_BATCH_RENAME_MODE_MANUAL, L"Manual"));
    _modeSelector->SetSelectedIndex(0);
    _modeSelector->SetOnSelectionChanged([this](int index)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        SwitchMode(index == 1 ? BatchRename::Mode::Manual : BatchRename::Mode::Rules);
    });

    _newNameLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_NEW_NAME, L"New name:"));
    _newNameLabel->SetMultiline(false);
    _nameTemplateField = _root->AddChild<TextField>(_rules.nameTemplate);
    _nameTemplateField->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_NEW_NAME, L"New name:"));
    _nameTemplateField->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.nameTemplate = std::wstring(text);
        RequestPreviewRebuild(false);
    });
    _nameTemplateHelperButton = _root->AddChild<Button>();
    _nameTemplateHelperButton->SetVariant(ButtonVariant::DropDown);
    const std::wstring templateHelperText = LoadBatchRenameString(IDS_BATCH_RENAME_HELPER_TEMPLATE_TOOLTIP, L"Insert name template helper");
    _nameTemplateHelperButton->SetTooltipText(templateHelperText);
    _nameTemplateHelperButton->SetAccessibleName(templateHelperText);
    _nameTemplateHelperButton->SetOnDropDownClick([this]() noexcept
    { ShowHelperMenu(BatchRenameMenus::HelperMenuKind::Template, _nameTemplateField, _nameTemplateHelperButton); });

    _searchForLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_SEARCH_FOR, L"Search for:"));
    _searchForLabel->SetMultiline(false);
    _searchForField = _root->AddChild<TextField>(_rules.searchFor);
    _searchForField->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_SEARCH_FOR, L"Search for:"));
    _searchForField->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.searchFor = std::wstring(text);
        RequestPreviewRebuild(false);
    });
    _searchForHelperButton = _root->AddChild<Button>();
    _searchForHelperButton->SetVariant(ButtonVariant::DropDown);
    const std::wstring regexHelperText = LoadBatchRenameString(IDS_BATCH_RENAME_HELPER_REGEX_TOOLTIP, L"Insert regular expression helper");
    _searchForHelperButton->SetTooltipText(regexHelperText);
    _searchForHelperButton->SetAccessibleName(regexHelperText);
    _searchForHelperButton->SetOnDropDownClick([this]() noexcept
    { ShowHelperMenu(BatchRenameMenus::HelperMenuKind::RegexSearch, _searchForField, _searchForHelperButton); });

    _replaceWithLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_REPLACE_WITH, L"Replace with:"));
    _replaceWithLabel->SetMultiline(false);
    _replaceWithField = _root->AddChild<TextField>(_rules.replaceWith);
    _replaceWithField->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_REPLACE_WITH, L"Replace with:"));
    _replaceWithField->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.replaceWith = std::wstring(text);
        RequestPreviewRebuild(false);
    });
    _replaceWithHelperButton = _root->AddChild<Button>();
    _replaceWithHelperButton->SetVariant(ButtonVariant::DropDown);
    const std::wstring replaceHelperText = LoadBatchRenameString(IDS_BATCH_RENAME_HELPER_REPLACE_TOOLTIP, L"Insert replacement helper");
    _replaceWithHelperButton->SetTooltipText(replaceHelperText);
    _replaceWithHelperButton->SetAccessibleName(replaceHelperText);
    _replaceWithHelperButton->SetOnDropDownClick([this]() noexcept
    { ShowHelperMenu(BatchRenameMenus::HelperMenuKind::Replacement, _replaceWithField, _replaceWithHelperButton); });

    _regexCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_REGEX, L"Regular expression"));
    _regexCheck->SetChecked(_rules.regexEnabled);
    _regexCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.regexEnabled = checked;
        RebuildPreviewFromRuleControlChange();
    });

    _caseSensitiveCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_CASE_SENSITIVE, L"Case sensitive"));
    _caseSensitiveCheck->SetChecked(_rules.caseSensitive);
    _caseSensitiveCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.caseSensitive = checked;
        RebuildPreviewFromRuleControlChange();
    });

    _wholeWordsCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_WHOLE_WORDS, L"Whole words"));
    _wholeWordsCheck->SetChecked(_rules.wholeWords);
    _wholeWordsCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.wholeWords = checked;
        RebuildPreviewFromRuleControlChange();
    });

    _replaceOnceCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_REPLACE_ONCE, L"Only once in each name"));
    _replaceOnceCheck->SetChecked(_rules.replaceOnce);
    _replaceOnceCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.replaceOnce = checked;
        RebuildPreviewFromRuleControlChange();
    });

    _excludeExtensionCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_EXCLUDE_EXTENSION, L"Exclude extension"));
    _excludeExtensionCheck->SetChecked(_rules.excludeExtension);
    _excludeExtensionCheck->SetOnToggled([this](bool checked)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.excludeExtension = checked;
        RebuildPreviewFromRuleControlChange();
    });

    const std::array<ComboBox::Item, 4> caseItems = BuildCaseComboItems();
    std::vector<ComboBox::Item> caseItemVector(caseItems.begin(), caseItems.end());

    _fileNameCaseLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_FILE_NAME_CASE, L"File name:"));
    _fileNameCaseLabel->SetMultiline(false);
    _fileNameCaseCombo = _root->AddChild<ComboBox>();
    _fileNameCaseCombo->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_FILE_NAME_CASE, L"File name:"));
    _fileNameCaseCombo->SetItems(caseItemVector);
    _fileNameCaseCombo->SetSelectedIndex(CaseTransformToIndex(_rules.fileNameCaseStyle));
    _fileNameCaseCombo->SetOnSelectionChanged([this](size_t index)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.fileNameCaseStyle = CaseTransformFromIndex(index);
        RebuildPreviewFromRuleControlChange();
    });

    _extensionCaseLabel = _root->AddChild<Label>(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_EXTENSION_CASE, L"Extension:"));
    _extensionCaseLabel->SetMultiline(false);
    _extensionCaseCombo = _root->AddChild<ComboBox>();
    _extensionCaseCombo->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_LABEL_EXTENSION_CASE, L"Extension:"));
    _extensionCaseCombo->SetItems(std::move(caseItemVector));
    _extensionCaseCombo->SetSelectedIndex(CaseTransformToIndex(_rules.extensionCaseStyle));
    _extensionCaseCombo->SetOnSelectionChanged([this](size_t index)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.extensionCaseStyle = CaseTransformFromIndex(index);
        RebuildPreviewFromRuleControlChange();
    });

    _manualNamesField = _root->AddChild<TextField>();
    _manualNamesField->SetMultiline(true);
    _manualNamesField->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_ACCESS_MANUAL_NAMES, L"Manual names"));
    _manualNamesField->SetVisible(false);
    _manualNamesField->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingRuleControls)
        {
            return;
        }
        _rules.mode            = BatchRename::Mode::Manual;
        _rules.manualNames     = SplitManualNames(text);
        _manualTextInitialized = true;
        UpdateModeVisibility();
        RequestPreviewRebuild(false);
    });

    _manualFillButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_FILL_FROM_PREVIEW, L"Fill from preview"));
    _manualFillButton->SetVisible(false);
    _manualFillButton->SetOnClick([this] { FillManualFromPreview(); });

    _manualClearButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_CLEAR_MANUAL, L"Clear"));
    _manualClearButton->SetVisible(false);
    _manualClearButton->SetOnClick([this] { ClearManualText(); });

    _manualPasteButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_PASTE_MANUAL, L"Paste"));
    _manualPasteButton->SetVisible(false);
    _manualPasteButton->SetOnClick([this] { static_cast<void>(PasteManualFromClipboard()); });

    _manualSortLikePreviewButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_SORT_LIKE_PREVIEW, L"Sort like preview"));
    _manualSortLikePreviewButton->SetVisible(false);
    _manualSortLikePreviewButton->SetOnClick([this] { static_cast<void>(SortManualTextLikePreview()); });

    _grid = _root->AddChild<Grid>();
    _grid->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_ACCESS_PREVIEW_GRID, L"Batch Rename preview"));
    _grid->SetDelegate(this);
    _grid->SetSelectionMode(GridSelectionMode::Single);
    _grid->SetHeaderHeightDip(30.0f);
    _grid->SetRowHeightDip(28.0f);
    _grid->SetLineClamp(1u);

    _gridModelStorage = std::make_unique<BatchRenamePreviewGridModel>();
    _gridModel        = _gridModelStorage.get();
    _grid->SetModel(_gridModel);
    _grid->SetSortSpec({});

    _status = _root->AddChild<StatusStrip>();

    _hideUnchangedCheck = _root->AddChild<Checkbox>(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_HIDE_UNCHANGED, L"Hide unchanged"));
    _hideUnchangedCheck->SetAccessibleName(LoadBatchRenameString(IDS_BATCH_RENAME_CHECK_HIDE_UNCHANGED, L"Hide unchanged"));
    _hideUnchangedCheck->SetChecked(_hideUnchangedRows);
    _hideUnchangedCheck->SetOnToggled([this](bool checked) { SetHideUnchangedRows(checked); });

    _cancelButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_CANCEL, L"Cancel"));
    _cancelButton->SetEnabled(false);
    _cancelButton->SetOnClick([this] { RequestTaskCancel(); });

    _renameButton = _root->AddChild<Button>(LoadBatchRenameString(IDS_BATCH_RENAME_BTN_RENAME, L"Rename"));
    _renameButton->SetPrimary(true);
    _renameButton->SetEnabled(false);
    _renameButton->SetOnClick([this] { static_cast<void>(ExecuteRename()); });

    _dxHost.SetRoot(std::move(_rootStorage));
    ApplyPreviewGridSettings();
    UpdateModeVisibility();
}

bool BatchRenameWindow::CreateRootNavigation(HWND parent) noexcept
{
    _rootNavigation.SetSettings(_settings);
    _rootNavigation.SetFileSystem(_context.fileSystem);
    _rootNavigation.SetTheme(_theme);
    _rootNavigation.SetEmbeddedDestinationMode(true);
    _rootNavigation.SetPaneFocused(false);
    _rootNavigation.SetPathChangedCallback([this](const std::optional<std::filesystem::path>& path) noexcept { OnRootNavigationPathChanged(path); });
    _rootNavigation.SetRequestFolderViewFocusCallback([this]
    {
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
    });

    const HWND hwnd = _rootNavigation.Create(parent, 0, 0, 0, 0);
    if (! hwnd)
    {
        return false;
    }

    RefreshRootNavigationPath();
    RefreshRootNavigationHistory();
    return true;
}

void BatchRenameWindow::ApplyTheme() noexcept
{
    _dxHost.SetTheme(MakeAppThemeDxPalette(_theme));
    _rootNavigation.SetTheme(_theme);
    if (_hWnd)
    {
        ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        ApplyWindowBackdropTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool);
    }
}

void BatchRenameWindow::Layout() noexcept
{
    if (! _root)
    {
        return;
    }

    const D2D1_RECT_F bounds = _dxHost.GetClientBoundsDip();
    _root->SetBounds(bounds);

    const float outer         = 16.0f;
    const float gap           = 8.0f;
    const float titleHeight   = 28.0f;
    const float rootHeight    = 30.0f;
    const float labelWidth    = 112.0f;
    const float fieldHeight   = 30.0f;
    const float checkHeight   = 28.0f;
    const float modeHeight    = 30.0f;
    const float manualHeight  = 118.0f;
    const float footerHeight  = 34.0f;
    const float buttonWidth   = 112.0f;
    const float buttonHeight  = 30.0f;
    const float contentRight  = bounds.right - outer;
    const auto snapDip        = [this](const float dip) noexcept { return _dxHost.PixelsToDip(std::round(_dxHost.DipsToPixels(dip))); };
    const auto moveChildToDip = [this, &snapDip](HWND hwnd, const float left, const float top, const float rightEdge, const float bottom) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        const float snappedLeft   = snapDip(left);
        const float snappedTop    = snapDip(top);
        const float snappedRight  = snapDip(rightEdge);
        const float snappedBottom = snapDip(bottom);
        const int x               = static_cast<int>(std::lround(_dxHost.DipsToPixels(snappedLeft)));
        const int y               = static_cast<int>(std::lround(_dxHost.DipsToPixels(snappedTop)));
        const int width           = std::max(0, static_cast<int>(std::lround(_dxHost.DipsToPixels(snappedRight - snappedLeft))));
        const int height          = std::max(0, static_cast<int>(std::lround(_dxHost.DipsToPixels(snappedBottom - snappedTop))));
        SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    };

    float top = outer;
    if (_titleLabel)
    {
        _titleLabel->SetBounds(D2D1::RectF(outer, top, contentRight, top + titleHeight));
        top += titleHeight + gap;
    }

    if (_rootLabel)
    {
        const float rootLabelWidth = 72.0f;
        _rootLabel->SetBounds(D2D1::RectF(outer, top, outer + rootLabelWidth, top + rootHeight));
        moveChildToDip(_rootNavigation.GetHwnd(), outer + rootLabelWidth + gap, top, contentRight, top + static_cast<float>(NavigationView::kHeight));
        top += rootHeight + gap;
    }

    const float maskLabelWidth = 52.0f;
    const float maskFieldWidth = 220.0f;
    if (_maskLabel)
    {
        _maskLabel->SetBounds(D2D1::RectF(outer, top, outer + maskLabelWidth, top + fieldHeight));
    }
    if (_maskField)
    {
        const float maskLeft = outer + maskLabelWidth + gap;
        _maskField->SetBounds(D2D1::RectF(maskLeft, top, maskLeft + maskFieldWidth, top + fieldHeight));
    }
    float scopeCheckLeft = outer + maskLabelWidth + gap + maskFieldWidth + (gap * 2.0f);
    if (_includeSubdirectoriesCheck)
    {
        _includeSubdirectoriesCheck->SetBounds(D2D1::RectF(scopeCheckLeft, top, scopeCheckLeft + 190.0f, top + checkHeight));
        scopeCheckLeft += 198.0f;
    }
    if (_includeFilesCheck)
    {
        _includeFilesCheck->SetBounds(D2D1::RectF(scopeCheckLeft, top, scopeCheckLeft + 78.0f, top + checkHeight));
        scopeCheckLeft += 86.0f;
    }
    if (_includeFoldersCheck)
    {
        _includeFoldersCheck->SetBounds(D2D1::RectF(scopeCheckLeft, top, std::min(contentRight, scopeCheckLeft + 98.0f), top + checkHeight));
    }
    top += fieldHeight + gap;

    if (_modeSelector)
    {
        _modeSelector->SetBounds(D2D1::RectF(outer, top, contentRight, top + modeHeight));
    }
    if (_rulesModeButton)
    {
        _rulesModeButton->SetBounds(D2D1::RectF(outer, top, outer + 120.0f, top + modeHeight));
    }
    if (_manualModeButton)
    {
        _manualModeButton->SetBounds(D2D1::RectF(outer + 128.0f, top, outer + 260.0f, top + modeHeight));
    }
    top += modeHeight + gap;

    const auto layoutLabelField = [&](Label* label, TextField* field, Button* helperButton) noexcept
    {
        constexpr float helperButtonWidth = 34.0f;
        if (label)
        {
            label->SetBounds(D2D1::RectF(outer, top, outer + labelWidth, top + fieldHeight));
        }
        const float fieldLeft  = outer + labelWidth + gap;
        const float helperLeft = contentRight - helperButtonWidth;
        if (field)
        {
            const float fieldRight = helperButton ? std::max(fieldLeft, helperLeft - gap) : contentRight;
            field->SetBounds(D2D1::RectF(fieldLeft, top, fieldRight, top + fieldHeight));
        }
        if (helperButton)
        {
            helperButton->SetBounds(D2D1::RectF(helperLeft, top, contentRight, top + fieldHeight));
        }
        top += fieldHeight + gap;
    };

    if (_rules.mode == BatchRename::Mode::Rules)
    {
        layoutLabelField(_newNameLabel, _nameTemplateField, _nameTemplateHelperButton);
        layoutLabelField(_searchForLabel, _searchForField, _searchForHelperButton);
        layoutLabelField(_replaceWithLabel, _replaceWithField, _replaceWithHelperButton);

        const float firstCheckWidth  = 186.0f;
        const float secondCheckWidth = 178.0f;
        const float thirdCheckWidth  = 150.0f;
        float checkLeft              = outer + labelWidth + gap;
        if (_caseSensitiveCheck)
        {
            _caseSensitiveCheck->SetBounds(D2D1::RectF(checkLeft, top, checkLeft + firstCheckWidth, top + checkHeight));
            checkLeft += firstCheckWidth + gap;
        }
        if (_wholeWordsCheck)
        {
            _wholeWordsCheck->SetBounds(D2D1::RectF(checkLeft, top, checkLeft + secondCheckWidth, top + checkHeight));
            checkLeft += secondCheckWidth + gap;
        }
        if (_regexCheck)
        {
            _regexCheck->SetBounds(D2D1::RectF(checkLeft, top, std::min(contentRight, checkLeft + thirdCheckWidth), top + checkHeight));
        }
        top += checkHeight + gap;

        checkLeft = outer + labelWidth + gap;
        if (_replaceOnceCheck)
        {
            _replaceOnceCheck->SetBounds(D2D1::RectF(checkLeft, top, checkLeft + firstCheckWidth + 40.0f, top + checkHeight));
            checkLeft += firstCheckWidth + 48.0f;
        }
        if (_excludeExtensionCheck)
        {
            _excludeExtensionCheck->SetBounds(D2D1::RectF(checkLeft, top, std::min(contentRight, checkLeft + secondCheckWidth), top + checkHeight));
        }
        top += checkHeight + gap;

        const float comboWidth     = 224.0f;
        const float caseLabelWidth = 92.0f;
        const float secondCaseLeft = outer + labelWidth + gap + comboWidth + gap + caseLabelWidth + gap;
        if (_fileNameCaseLabel)
        {
            _fileNameCaseLabel->SetBounds(D2D1::RectF(outer, top, outer + labelWidth, top + fieldHeight));
        }
        if (_fileNameCaseCombo)
        {
            _fileNameCaseCombo->SetBounds(D2D1::RectF(outer + labelWidth + gap, top, outer + labelWidth + gap + comboWidth, top + fieldHeight));
        }
        if (_extensionCaseLabel)
        {
            _extensionCaseLabel->SetBounds(D2D1::RectF(secondCaseLeft - caseLabelWidth - gap, top, secondCaseLeft - gap, top + fieldHeight));
        }
        if (_extensionCaseCombo)
        {
            _extensionCaseCombo->SetBounds(D2D1::RectF(secondCaseLeft, top, std::min(contentRight, secondCaseLeft + comboWidth), top + fieldHeight));
        }
        top += fieldHeight + gap;
    }
    else
    {
        if (_manualNamesField)
        {
            _manualNamesField->SetBounds(D2D1::RectF(outer, top, contentRight, top + manualHeight));
        }
        top += manualHeight + gap;

        const float manualButtonWidth = 150.0f;
        if (_manualFillButton)
        {
            _manualFillButton->SetBounds(D2D1::RectF(outer, top, outer + manualButtonWidth, top + buttonHeight));
        }
        if (_manualClearButton)
        {
            _manualClearButton->SetBounds(D2D1::RectF(outer + manualButtonWidth + gap, top, outer + manualButtonWidth + gap + 92.0f, top + buttonHeight));
        }
        if (_manualPasteButton)
        {
            const float pasteLeft = outer + manualButtonWidth + gap + 92.0f + gap;
            _manualPasteButton->SetBounds(D2D1::RectF(pasteLeft, top, pasteLeft + 92.0f, top + buttonHeight));
        }
        if (_manualSortLikePreviewButton)
        {
            const float sortLeft = outer + manualButtonWidth + gap + 92.0f + gap + 92.0f + gap;
            _manualSortLikePreviewButton->SetBounds(D2D1::RectF(sortLeft, top, std::min(contentRight, sortLeft + 156.0f), top + buttonHeight));
        }
        top += buttonHeight + gap;
    }

    const float footerTop = std::max(top, bounds.bottom - outer - footerHeight);
    if (_grid)
    {
        _grid->SetBounds(D2D1::RectF(outer, top, contentRight, footerTop - gap));
    }

    const float buttonLeft         = bounds.right - outer - buttonWidth;
    const float cancelLeft         = std::max(outer, buttonLeft - gap - buttonWidth);
    const float hideUnchangedWidth = 156.0f;
    const float hideUnchangedLeft  = std::max(outer, cancelLeft - gap - hideUnchangedWidth);
    if (_status)
    {
        _status->SetBounds(D2D1::RectF(outer, footerTop, std::max(outer, hideUnchangedLeft - gap), footerTop + footerHeight));
    }

    if (_hideUnchangedCheck)
    {
        const float checkTop = footerTop + (footerHeight - checkHeight) / 2.0f;
        _hideUnchangedCheck->SetBounds(D2D1::RectF(hideUnchangedLeft, checkTop, cancelLeft - gap, checkTop + checkHeight));
    }

    if (_cancelButton)
    {
        const float buttonTop = footerTop + (footerHeight - buttonHeight) / 2.0f;
        _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonTop, buttonLeft - gap, buttonTop + buttonHeight));
    }

    if (_renameButton)
    {
        const float buttonTop = footerTop + (footerHeight - buttonHeight) / 2.0f;
        _renameButton->SetBounds(D2D1::RectF(buttonLeft, buttonTop, bounds.right - outer, buttonTop + buttonHeight));
    }
}

void BatchRenameWindow::ApplySettingsDefaults() noexcept
{
    if (! _settings || ! _settings->batchRename.has_value())
    {
        return;
    }

    const Common::Settings::BatchRenameSettings& settings = _settings->batchRename.value();
    if (_rootText.empty() && ! settings.lastRoot.empty())
    {
        _rootText = settings.lastRoot;
    }
    if (! settings.recentNameTemplates.empty() && ! settings.recentNameTemplates.front().empty())
    {
        _rules.nameTemplate = settings.recentNameTemplates.front();
    }
    if (! settings.recentSearchPatterns.empty() && ! settings.recentSearchPatterns.front().empty())
    {
        _rules.searchFor = settings.recentSearchPatterns.front();
    }
    if (! settings.recentReplacePatterns.empty() && ! settings.recentReplacePatterns.front().empty())
    {
        _rules.replaceWith = settings.recentReplacePatterns.front();
    }
    if (! settings.recentMasks.empty() && ! settings.recentMasks.front().empty())
    {
        _scopeOptions.mask = settings.recentMasks.front();
    }

    _scopeOptions.includeSubdirectories = settings.includeSubdirectories;
    _scopeOptions.includeFiles          = settings.includeFiles;
    _scopeOptions.includeFolders        = settings.includeFolders;
    _rules.regexEnabled                 = settings.regexEnabled;
    _rules.caseSensitive                = settings.caseSensitive;
    _rules.wholeWords                   = settings.wholeWords;
    _rules.replaceOnce                  = settings.replaceOnce;
    _rules.excludeExtension             = settings.excludeExtension;
    _rules.flattenSeparator   = settings.flattenSeparator.empty() ? Common::Settings::BatchRenameSettings{}.flattenSeparator : settings.flattenSeparator;
    _rules.fileNameCaseStyle  = CaseTransformFromSettings(settings.fileNameCaseStyle);
    _rules.extensionCaseStyle = CaseTransformFromSettings(settings.extensionCaseStyle);
    _rules.mode               = BatchRename::Mode::Rules;
    _rules.manualNames.clear();
    _manualTextInitialized = false;
}

void BatchRenameWindow::ApplyPreviewGridSettings() noexcept
{
    if (! _grid || ! _gridModel || ! _settings || ! _settings->batchRename.has_value())
    {
        return;
    }

    const Common::Settings::BatchRenameSettings& settings = _settings->batchRename.value();
    const auto layout                                     = ConvertColumnLayout(settings.previewGridLayout);
    if (! layout.empty())
    {
        _grid->ApplyColumnLayout(layout);
    }

    const std::optional<size_t> sortColumn = FindPreviewColumnIndexById(settings.previewSortColumnId);
    if (sortColumn.has_value())
    {
        _grid->SetSortSpec(GridSortSpec{
            .columnIndex = sortColumn.value(),
            .direction   = settings.previewSortDescending ? SortDirection::Descending : SortDirection::Ascending,
        });
    }
}

void BatchRenameWindow::PersistUiState(const bool updateHistory) noexcept
{
    if (! _settings)
    {
        return;
    }

    Common::Settings::BatchRenameSettings settings =
        _settings->batchRename.has_value() ? _settings->batchRename.value() : Common::Settings::BatchRenameSettings{};

    settings.lastRoot              = _rootText;
    settings.includeSubdirectories = _scopeOptions.includeSubdirectories;
    settings.includeFiles          = _scopeOptions.includeFiles;
    settings.includeFolders        = _scopeOptions.includeFolders;
    settings.regexEnabled          = _rules.regexEnabled;
    settings.caseSensitive         = _rules.caseSensitive;
    settings.wholeWords            = _rules.wholeWords;
    settings.replaceOnce           = _rules.replaceOnce;
    settings.excludeExtension      = _rules.excludeExtension;
    settings.flattenSeparator      = _rules.flattenSeparator.empty() ? Common::Settings::BatchRenameSettings{}.flattenSeparator : _rules.flattenSeparator;
    settings.fileNameCaseStyle     = CaseTransformToSettings(_rules.fileNameCaseStyle);
    settings.extensionCaseStyle    = CaseTransformToSettings(_rules.extensionCaseStyle);
    settings.previewSortColumnId.clear();
    settings.previewSortDescending = false;
    settings.previewGridLayout.clear();

    if (_grid && _gridModel)
    {
        const GridSortSpec sortSpec = _grid->GetSortSpec();
        if (sortSpec.direction != SortDirection::None && sortSpec.columnIndex < _gridModel->GetColumnCount())
        {
            settings.previewSortColumnId   = _gridModel->GetColumn(sortSpec.columnIndex).id;
            settings.previewSortDescending = sortSpec.direction == SortDirection::Descending;
        }
        settings.previewGridLayout = ConvertColumnLayout(_grid->CaptureColumnLayout());
    }

    if (updateHistory)
    {
        UpdateRecentBatchRenameValue(settings.recentMasks, NormalizeScopeMask(_scopeOptions.mask));
        UpdateRecentBatchRenameValue(settings.recentNameTemplates, _rules.nameTemplate);
        UpdateRecentBatchRenameValue(settings.recentSearchPatterns, _rules.searchFor);
        UpdateRecentBatchRenameValue(settings.recentReplacePatterns, _rules.replaceWith);
    }

    _settings->batchRename = std::move(settings);
}

void BatchRenameWindow::PersistCloseState() noexcept
{
    if (_uiStatePersisted)
    {
        return;
    }

    _uiStatePersisted = true;
    if (! _settings)
    {
        return;
    }

    PersistUiState(true);
    if (_hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kBatchRenameWindowId, _hWnd.get());
    }
}

void BatchRenameWindow::RebuildTargetsFromScope()
{
    if (! _hWnd)
    {
        return;
    }

    StartTargetCollection();
}

void BatchRenameWindow::RequestPreviewRebuild(const bool rebuildTargets) noexcept
{
    ClearExecutionReport();
    _previewRebuildPending = true;
    _targetRebuildPending  = _targetRebuildPending || rebuildTargets;

    if (! _hWnd)
    {
        OnPreviewRebuildTimer();
        return;
    }

    static_cast<void>(KillTimer(_hWnd.get(), kPreviewRebuildTimerId));
    if (SetTimer(_hWnd.get(), kPreviewRebuildTimerId, kPreviewRebuildDebounceMs, nullptr) == 0)
    {
        OnPreviewRebuildTimer();
    }
}

void BatchRenameWindow::CancelPendingPreviewRebuild() noexcept
{
    _previewRebuildPending = false;
    _targetRebuildPending  = false;
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kPreviewRebuildTimerId));
    }
}

void BatchRenameWindow::OnPreviewRebuildTimer() noexcept
{
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kPreviewRebuildTimerId));
    }

    const bool rebuildTargets = _targetRebuildPending;
    const bool rebuildPreview = _previewRebuildPending;
    _targetRebuildPending     = false;
    _previewRebuildPending    = false;

    if (! rebuildPreview)
    {
        return;
    }

    if (rebuildTargets)
    {
        RebuildTargetsFromScope();
    }
    else
    {
        RebuildPreviewFromRuleControlChange();
    }
}

void BatchRenameWindow::RebuildPreview() noexcept
{
    if (! _gridModel)
    {
        return;
    }
    _currentPlan.reset();

    auto payload = std::unique_ptr<BatchRenamePreviewCompletedPayload>(new (std::nothrow) BatchRenamePreviewCompletedPayload{});
    auto work    = std::unique_ptr<BatchRenamePreviewWork>(new (std::nothrow) BatchRenamePreviewWork{});
    if (! payload || ! work)
    {
        _previewing = false;
        _previewStats = {};
        _fullPreviewRows.clear();
        RefreshVisibleRows();
        return;
    }

    const uint64_t generation = _previewGeneration.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    work->hwnd       = _hWnd.get();
    work->generation = generation;
    work->context    = _context;
    work->targets    = _targets;
    work->rules      = _rules;
    work->payload    = std::move(payload);
    work->modulePin  = AcquireModuleReferenceFromAddress(&kBatchRenameModuleAnchor);
    if (! work->hwnd || ! work->modulePin || ! SubmitOwnedThreadpoolCallbackWithInstance(work))
    {
        _previewing = false;
        _previewStats = {};
        _fullPreviewRows.clear();
        RefreshVisibleRows();
        return;
    }

    _previewing = true;
    UpdateTaskUi();
}

void BatchRenameWindow::OnPreviewCompleted(std::unique_ptr<BatchRenamePreviewCompletedPayload> payload) noexcept
{
    if (! payload || payload->generation != _previewGeneration.load(std::memory_order_acquire))
    {
        return;
    }

    Debug::Perf::Scope recomputePerf(L"batchrename.preview.recompute.us");
    recomputePerf.SetDetail(_rules.mode == BatchRename::Mode::Manual ? L"worker-manual" : L"worker-rules");
    _previewing  = false;
    _previewStats = payload->plan.stats;
    recomputePerf.SetValue0(static_cast<uint64_t>(payload->plan.rows.size()));
    recomputePerf.SetValue1(static_cast<uint64_t>(payload->plan.stats.changedRows));
    _fullPreviewRows = BuildPreviewRowsFromPlan(payload->plan, _rootText);
    _currentPlan     = std::move(payload->plan);
    RefreshVisibleRows();
    UpdateTaskUi();
}

void BatchRenameWindow::RefreshVisibleRows() noexcept
{
    if (! _gridModel)
    {
        return;
    }

    const auto visibleRefreshStartedAt       = std::chrono::steady_clock::now();
    std::vector<BatchPreviewRow> previewRows = _fullPreviewRows;
    if (_hideUnchangedRows)
    {
        std::erase_if(previewRows, [](const BatchPreviewRow& row) noexcept { return ! row.changed; });
    }
    const uint64_t visibleRows = static_cast<uint64_t>(previewRows.size());
    _gridModel->SetRows(std::move(previewRows));
    if (_grid)
    {
        _gridModel->SortRows(_grid->GetSortSpec());
        _grid->NotifyDataChanged();
    }
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"batchrename.preview.visible_refresh.us",
                          _grid ? L"grid" : L"model",
                          Debug::Perf::ElapsedUs(visibleRefreshStartedAt),
                          visibleRows,
                          static_cast<uint64_t>(_previewStats.changedRows));
    }
    if (_renameButton)
    {
        _renameButton->SetEnabled(_previewStats.changedRows > 0u && _previewStats.errorRows == 0u && ! _collecting && ! _executing && ! _previewing);
    }
    UpdateStatus();
}

void BatchRenameWindow::SyncRuleControls() noexcept
{
    _syncingRuleControls = true;
    const auto clearSync = wil::scope_exit([this]() noexcept { _syncingRuleControls = false; });

    if (_modeSelector)
    {
        _modeSelector->SetSelectedIndex(_rules.mode == BatchRename::Mode::Manual ? 1 : 0);
    }
    if (_nameTemplateField)
    {
        _nameTemplateField->SetText(_rules.nameTemplate);
    }
    if (_searchForField)
    {
        _searchForField->SetText(_rules.searchFor);
    }
    if (_replaceWithField)
    {
        _replaceWithField->SetText(_rules.replaceWith);
    }
    if (_maskField)
    {
        _maskField->SetText(_scopeOptions.mask);
    }
    if (_includeSubdirectoriesCheck)
    {
        _includeSubdirectoriesCheck->SetChecked(_scopeOptions.includeSubdirectories);
    }
    if (_includeFilesCheck)
    {
        _includeFilesCheck->SetChecked(_scopeOptions.includeFiles);
    }
    if (_includeFoldersCheck)
    {
        _includeFoldersCheck->SetChecked(_scopeOptions.includeFolders);
    }
    if (_regexCheck)
    {
        _regexCheck->SetChecked(_rules.regexEnabled);
    }
    if (_caseSensitiveCheck)
    {
        _caseSensitiveCheck->SetChecked(_rules.caseSensitive);
    }
    if (_wholeWordsCheck)
    {
        _wholeWordsCheck->SetChecked(_rules.wholeWords);
    }
    if (_replaceOnceCheck)
    {
        _replaceOnceCheck->SetChecked(_rules.replaceOnce);
    }
    if (_excludeExtensionCheck)
    {
        _excludeExtensionCheck->SetChecked(_rules.excludeExtension);
    }
    if (_fileNameCaseCombo)
    {
        _fileNameCaseCombo->SetSelectedIndex(CaseTransformToIndex(_rules.fileNameCaseStyle));
    }
    if (_extensionCaseCombo)
    {
        _extensionCaseCombo->SetSelectedIndex(CaseTransformToIndex(_rules.extensionCaseStyle));
    }
    if (_manualNamesField)
    {
        _manualNamesField->SetText(_manualTextInitialized ? JoinManualNames(_rules.manualNames) : std::wstring{});
    }
    UpdateModeVisibility();
}

void BatchRenameWindow::UpdateModeVisibility() noexcept
{
    const bool rulesVisible  = _rules.mode == BatchRename::Mode::Rules;
    const bool manualVisible = _rules.mode == BatchRename::Mode::Manual;

    if (_modeSelector)
    {
        _modeSelector->SetSelectedIndex(manualVisible ? 1 : 0);
    }

    const std::array<Control*, 18> ruleControls = {_newNameLabel,
                                                   _searchForLabel,
                                                   _replaceWithLabel,
                                                   _fileNameCaseLabel,
                                                   _extensionCaseLabel,
                                                   _nameTemplateField,
                                                   _searchForField,
                                                   _replaceWithField,
                                                   _nameTemplateHelperButton,
                                                   _searchForHelperButton,
                                                   _replaceWithHelperButton,
                                                   _regexCheck,
                                                   _caseSensitiveCheck,
                                                   _wholeWordsCheck,
                                                   _replaceOnceCheck,
                                                   _excludeExtensionCheck,
                                                   _fileNameCaseCombo,
                                                   _extensionCaseCombo};
    for (Control* control : ruleControls)
    {
        if (control)
        {
            control->SetVisible(rulesVisible);
        }
    }

    if (_manualNamesField)
    {
        _manualNamesField->SetVisible(manualVisible);
    }
    if (_manualFillButton)
    {
        _manualFillButton->SetVisible(manualVisible);
    }
    if (_manualClearButton)
    {
        _manualClearButton->SetVisible(manualVisible);
    }
    if (_manualPasteButton)
    {
        _manualPasteButton->SetVisible(manualVisible);
    }
    if (_manualSortLikePreviewButton)
    {
        _manualSortLikePreviewButton->SetVisible(manualVisible);
    }
}

void BatchRenameWindow::SwitchMode(const BatchRename::Mode mode)
{
    if (mode == BatchRename::Mode::Manual && ! _manualTextInitialized)
    {
        SeedManualTextFromRulePreview();
    }
    _rules.mode = mode;
    UpdateModeVisibility();
    Layout();
    RebuildPreviewFromRuleControlChange();
}

void BatchRenameWindow::SeedManualTextFromRulePreview()
{
    std::vector<std::wstring> names;
    names.reserve(_fullPreviewRows.size());
    for (const BatchPreviewRow& row : _fullPreviewRows)
    {
        names.push_back(row.newName);
    }
    SetManualText(JoinManualNames(names));
}

void BatchRenameWindow::SetManualText(std::wstring text)
{
    _manualTextInitialized = true;
    _rules.mode            = BatchRename::Mode::Manual;
    _rules.manualNames     = SplitManualNames(text);

    if (_manualNamesField)
    {
        _syncingRuleControls = true;
        const auto clearSync = wil::scope_exit([this]() noexcept { _syncingRuleControls = false; });
        _manualNamesField->SetText(std::move(text));
    }
}

void BatchRenameWindow::FillManualFromPreview()
{
    SeedManualTextFromRulePreview();
    UpdateModeVisibility();
    RebuildPreviewFromRuleControlChange();
}

void BatchRenameWindow::ClearManualText()
{
    SetManualText({});
    UpdateModeVisibility();
    RebuildPreviewFromRuleControlChange();
}

bool BatchRenameWindow::PasteManualFromClipboard()
{
    const std::optional<std::wstring> clipboardText = _dxHost.ReadTextFromClipboard();
    if (! clipboardText.has_value())
    {
        return false;
    }

    SetManualText(clipboardText.value());
    UpdateModeVisibility();
    RebuildPreviewFromRuleControlChange();
    return true;
}

bool BatchRenameWindow::SortManualTextLikePreview()
{
    if (! _gridModel || _rules.mode != BatchRename::Mode::Manual || _rules.manualNames.size() != _targets.size())
    {
        return false;
    }

    // Use the retained full preview (independent of Hide unchanged) so this command performs
    // no provider query or local destination I/O on the UI thread.
    std::vector<BatchPreviewRow> orderedRows = _fullPreviewRows;
    if (_grid)
    {
        SortBatchPreviewRows(orderedRows, _gridModel->GetColumns(), _grid->GetSortSpec());
    }

    std::vector<BatchRename::Target> sortedTargets;
    std::vector<std::wstring> sortedManualNames;
    std::vector<bool> consumedTargets(_targets.size(), false);
    sortedTargets.reserve(_targets.size());
    sortedManualNames.reserve(_rules.manualNames.size());

    for (const BatchPreviewRow& row : orderedRows)
    {
        if (row.targetIndex >= _targets.size() || row.targetIndex >= _rules.manualNames.size() || consumedTargets[row.targetIndex])
        {
            return false;
        }

        consumedTargets[row.targetIndex] = true;
        sortedTargets.push_back(_targets[row.targetIndex]);
        sortedManualNames.push_back(_rules.manualNames[row.targetIndex]);
    }

    if (sortedTargets.size() != _targets.size() || sortedManualNames.size() != _rules.manualNames.size() ||
        ! std::ranges::all_of(consumedTargets, [](const bool consumed) noexcept { return consumed; }))
    {
        return false;
    }

    _targets = std::move(sortedTargets);
    SetManualText(JoinManualNames(sortedManualNames));
    UpdateModeVisibility();
    Layout();
    RebuildPreviewFromRuleControlChange();
    return true;
}

void BatchRenameWindow::ClearExecutionReport() noexcept
{
    _lastExecutionReport.reset();
}

void BatchRenameWindow::StoreExecutionReport(BatchRenameExecutionReport report) noexcept
{
    if (FAILED(report.firstFailure) && report.firstFailureText.empty())
    {
        report.firstFailureText = FormatHResultMessageWithCode(report.firstFailure);
    }
    _lastExecutionReport = std::move(report);
    UpdateStatus();
}

HRESULT BatchRenameWindow::ExecuteRename() noexcept
{
    // Re-entrancy guard: Rename cannot run while an execution or a target
    // collection is still in flight.
    if (_executing || _collecting || _previewing || ! _currentPlan.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_BUSY);
    }

    if (_targetRebuildPending)
    {
        OnPreviewRebuildTimer();
        if (_collecting)
        {
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        }
    }
    else
    {
        // ExecuteRename builds the authoritative plan from the current rules.
        // Do not let a debounced preview timer race the filesystem worker over
        // the same paths while execution is in flight.
        CancelPendingPreviewRebuild();
    }

    const auto startedAt         = std::chrono::steady_clock::now();
    uint64_t executeRows         = 0u;
    const auto finishSynchronous = [&](const HRESULT hr, const std::wstring_view detail, BatchRenameExecutionReport report) -> HRESULT
    {
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"batchrename.execute.us", detail, Debug::Perf::ElapsedUs(startedAt), executeRows, 0u, hr);
        }
        EmitBatchRenameExecuteCounters(executeRows, 0u, static_cast<uint64_t>(report.failedRows));
        StoreExecutionReport(std::move(report));
        _lastExecutionTerminalHr = hr;
        return hr;
    };

    if (! _context.fileSystem)
    {
        BatchRenameExecutionReport report{};
        report.failedRows   = 1u;
        report.firstFailure = E_POINTER;
        return finishSynchronous(E_POINTER, L"missing_context", std::move(report));
    }

    BatchRename::Plan plan = _currentPlan.value();
    BatchRenameExecutionReport report{};
    report.totalRows   = plan.rows.size();
    report.skippedRows = plan.stats.unchangedRows;
    if (plan.stats.errorRows != 0u)
    {
        report.failedRows   = plan.stats.errorRows;
        report.firstFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        _previewStats       = plan.stats;
        if (_renameButton)
        {
            _renameButton->SetEnabled(false);
        }
        UpdateStatus();
        return finishSynchronous(report.firstFailure, L"preview_errors", std::move(report));
    }

    const std::optional<FileSystemPathIdentity> executionPathIdentity = ResolveBatchRenamePathIdentity(_context);
    if (! executionPathIdentity.has_value())
    {
        report.failedRows   = plan.stats.changedRows;
        report.firstFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return finishSynchronous(report.firstFailure, L"provider_path_identity_unknown", std::move(report));
    }

    std::vector<BatchRenameExecutionOp> ops;
    ops.reserve(plan.rows.size());
    for (const BatchRename::PreviewRow& row : plan.rows)
    {
        if (row.newName == row.originalName)
        {
            continue;
        }

        BatchRenameExecutionOp op{};
        op.currentSource  = row.sourcePath;
        op.originalSource = row.sourcePath;
        op.finalLeaf      = row.newName;
        op.depth          = PathDepthKey(row.sourcePath);
        op.isDirectory    = row.isDirectory;
        ops.push_back(std::move(op));
    }
    executeRows = static_cast<uint64_t>(ops.size());

    if (ops.empty())
    {
        RebuildPreview();
        return finishSynchronous(S_OK, L"noop", std::move(report));
    }

    std::ranges::sort(ops,
                      [](const BatchRenameExecutionOp& lhs, const BatchRenameExecutionOp& rhs) noexcept
    {
        if (lhs.depth != rhs.depth)
        {
            return lhs.depth > rhs.depth;
        }
        return lhs.currentSource.native().size() > rhs.currentSource.native().size();
    });

    auto payload = std::unique_ptr<BatchRenameExecutionCompletedPayload>(new (std::nothrow) BatchRenameExecutionCompletedPayload{});
    if (! payload)
    {
        report.failedRows   = ops.size();
        report.firstFailure = E_OUTOFMEMORY;
        return finishSynchronous(E_OUTOFMEMORY, L"rename_failed", std::move(report));
    }

    if (_collectionCancelFlag)
    {
        _collectionCancelFlag->store(true, std::memory_order_release);
    }
    _cancelRequested.store(false, std::memory_order_release);

    const uint64_t generation = _taskGeneration.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    payload->generation       = generation;
    payload->report           = std::move(report);

    const HWND hwnd                           = _hWnd.get();
    wil::com_ptr<IFileSystem> fileSystem      = _context.fileSystem;
    const FileSystemPathIdentity pathIdentity = executionPathIdentity.value();
    std::atomic_bool* const cancelFlag        = &_cancelRequested;
    const size_t opsCount                     = ops.size();
    const size_t skippedRows                  = plan.stats.unchangedRows;
    const size_t totalRows                    = plan.rows.size();

    _executing = true;
    UpdateTaskUi();

    try
    {
        std::optional<BatchRename::Plan> localPlan;
        if (IsLocalFileSystemContext(_context))
        {
            localPlan = plan;
        }
        _taskWorker = std::jthread([hwnd,
                                    generation,
                                    cancelFlag,
                                    fileSystem = std::move(fileSystem),
                                    pathIdentity,
                                    localPlan = std::move(localPlan),
                                    ops     = std::move(ops),
                                    payload = std::move(payload)]() mutable noexcept
        {
            RunBatchRenameExecution(
                hwnd, generation, *cancelFlag, std::move(fileSystem), pathIdentity, std::move(localPlan), std::move(ops), std::move(payload));
        });
    }
    catch (const std::system_error&)
    {
        // Thread creation is required for responsive execution; report a clean
        // start failure without touching provider state.
        _executing = false;
        BatchRenameExecutionReport failureReport{};
        failureReport.totalRows    = totalRows;
        failureReport.skippedRows  = skippedRows;
        failureReport.failedRows   = opsCount;
        failureReport.firstFailure = HRESULT_FROM_WIN32(ERROR_NO_SYSTEM_RESOURCES);
        UpdateTaskUi();
        return finishSynchronous(failureReport.firstFailure, L"rename_failed", std::move(failureReport));
    }

    return S_OK;
}

void BatchRenameWindow::StartTargetCollection() noexcept
{
    if (_executing)
    {
        // Execution owns the single background worker; collect afterwards.
        _collectionQueued = true;
        return;
    }

    if (_collectionCancelFlag)
    {
        _collectionCancelFlag->store(true, std::memory_order_release);
    }
    _cancelRequested.store(false, std::memory_order_release);

    const HWND hwnd = _hWnd.get();
    if (! hwnd)
    {
        return;
    }

    auto payload = std::unique_ptr<BatchRenameCollectionCompletedPayload>(new (std::nothrow) BatchRenameCollectionCompletedPayload{});
    auto work    = std::unique_ptr<BatchRenameCollectionWork>(new (std::nothrow) BatchRenameCollectionWork{});
    auto cancelFlag = std::shared_ptr<std::atomic_bool>(new (std::nothrow) std::atomic_bool{false});
    if (! payload || ! work || ! cancelFlag)
    {
        _targets.clear();
        _collecting = false;
        RebuildPreviewFromRuleControlChange();
        UpdateTaskUi();
        return;
    }

    const uint64_t generation = _taskGeneration.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    _collectionCancelFlag     = cancelFlag;
    _collecting               = true;
    ClearExecutionReport();
    UpdateTaskUi();

    work->hwnd            = hwnd;
    work->generation      = generation;
    work->context         = _context;
    work->scope           = _scopeOptions;
    work->cancelRequested = std::move(cancelFlag);
    work->payload         = std::move(payload);
    work->modulePin       = AcquireModuleReferenceFromAddress(&kBatchRenameModuleAnchor);
    if (! work->modulePin || ! SubmitOwnedThreadpoolCallbackWithInstance(work))
    {
        // Never fall back to provider enumeration on the UI thread.
        _collecting = false;
        _collectionCancelFlag.reset();
        _targets.clear();
        RebuildPreviewFromRuleControlChange();
        Layout();
        UpdateTaskUi();
    }
}

void BatchRenameWindow::OnCollectionCompleted(std::unique_ptr<BatchRenameCollectionCompletedPayload> payload) noexcept
{
    if (! payload || payload->generation != _taskGeneration.load(std::memory_order_acquire))
    {
        return;
    }

    _collecting         = false;
    _collectionCancelFlag.reset();
    const bool canceled = IsBatchRenameCancellationHRESULT(payload->hr);
    if (FAILED(payload->hr) && ! canceled)
    {
        Debug::Error(L"BatchRename: target collection failed (hr=0x{:08X}, scope={}).", static_cast<unsigned long>(payload->hr), payload->detail);
    }

    _targets = std::move(payload->targets);
    RebuildPreviewFromRuleControlChange();
    Layout();
    UpdateTaskUi();

    if (FAILED(payload->hr) && ! canceled && _status)
    {
        std::wstring status = FormatStringResource(nullptr, IDS_BATCH_RENAME_COLLECT_FAILED_FMT, FormatHResultMessageWithCode(payload->hr));
        if (status.empty())
        {
            status = std::format(L"Collecting items failed: {}", FormatHResultMessageWithCode(payload->hr));
        }
        _status->SetText(std::move(status));
        _dxHost.Invalidate();
    }
}

void BatchRenameWindow::OnExecutionCompleted(std::unique_ptr<BatchRenameExecutionCompletedPayload> payload) noexcept
{
    if (! payload || payload->generation != _taskGeneration.load(std::memory_order_acquire))
    {
        return;
    }

    _executing               = false;
    _lastExecutionTerminalHr = payload->hr;

    const FileSystemPathIdentity pathIdentity =
        ResolveBatchRenamePathIdentity(_context).value_or(FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem());
    const size_t successCount = std::min(payload->successfulSourcePaths.size(), payload->successfulTargetPaths.size());

    struct CacheNotifyMove final
    {
        size_t index       = 0u;
        size_t sourceDepth = 0u;
        std::filesystem::path finalTargetPath;
    };

    std::vector<CacheNotifyMove> cacheNotifyMoves;
    cacheNotifyMoves.reserve(successCount);
    for (size_t index = 0u; index < successCount; ++index)
    {
        cacheNotifyMoves.push_back(CacheNotifyMove{
            .index           = index,
            .sourceDepth     = PathDepthKey(payload->successfulSourcePaths[index]),
            .finalTargetPath = ApplyExecutedDirectoryMoves(pathIdentity, payload->successfulTargetPaths[index], payload->executedDirectoryMoves),
        });
    }
    std::ranges::stable_sort(cacheNotifyMoves, {}, &CacheNotifyMove::sourceDepth);
    for (const CacheNotifyMove& move : cacheNotifyMoves)
    {
        DirectoryInfoCache::GetInstance().NotifyPathMoved(_context.fileSystem.get(), payload->successfulSourcePaths[move.index], move.finalTargetPath);
    }

    static_cast<void>(RefreshBatchRenameTargetsAfterExecution(
        pathIdentity, _targets, payload->successfulSourcePaths, payload->successfulTargetPaths, payload->executedDirectoryMoves, _context.rootPluginPath));

    RebuildPreview();
    _dxHost.Invalidate();
    StoreExecutionReport(std::move(payload->report));
    UpdateTaskUi();

    // Successful renames are reported on the failure and canceled paths too,
    // so panes refresh for the rows that did rename before the batch stopped.
    if (_context.onSuccessfulRename && successCount != 0u)
    {
        _context.onSuccessfulRename(payload->successfulSourcePaths, payload->successfulTargetPaths);
    }

    if (_collectionQueued)
    {
        _collectionQueued = false;
        StartTargetCollection();
    }
}

void BatchRenameWindow::OnTaskProgress(std::unique_ptr<BatchRenameTaskProgressPayload> payload) noexcept
{
    if (! payload || payload->generation != _taskGeneration.load(std::memory_order_acquire) || ! _executing || ! _status)
    {
        return;
    }

    if (_cancelRequested.load(std::memory_order_acquire))
    {
        // Keep the canceling status text once cancel was requested.
        return;
    }

    std::wstring status = FormatStringResource(nullptr, IDS_BATCH_RENAME_EXECUTE_PROGRESS_FMT, payload->completedItems, payload->totalItems);
    if (status.empty())
    {
        status = std::format(L"Renaming {} of {}...", payload->completedItems, payload->totalItems);
    }
    _status->SetText(std::move(status));
    _dxHost.Invalidate();
}

void BatchRenameWindow::RequestTaskCancel() noexcept
{
    if (! _collecting && ! _executing)
    {
        return;
    }

    _cancelRequested.store(true, std::memory_order_release);
    if (_collectionCancelFlag)
    {
        _collectionCancelFlag->store(true, std::memory_order_release);
    }
    if (_cancelButton)
    {
        _cancelButton->SetEnabled(false);
    }
    if (_status)
    {
        _status->SetText(LoadBatchRenameString(IDS_BATCH_RENAME_STATUS_CANCELING, L"Canceling..."));
    }
    _dxHost.Invalidate();
}

void BatchRenameWindow::CancelAndJoinBackgroundTask() noexcept
{
    if (_collectionCancelFlag)
    {
        _collectionCancelFlag->store(true, std::memory_order_release);
        _collectionCancelFlag.reset();
    }
    if (_taskWorker.joinable())
    {
        _cancelRequested.store(true, std::memory_order_release);
        _taskWorker.join();
    }
    _collecting       = false;
    _executing        = false;
    _collectionQueued = false;
}

void BatchRenameWindow::UpdateTaskUi() noexcept
{
    const bool cancelableBusy = _collecting || _executing;
    const bool busy           = cancelableBusy || _previewing;
    if (_cancelButton)
    {
        _cancelButton->SetEnabled(cancelableBusy && ! _cancelRequested.load(std::memory_order_acquire));
    }
    if (_renameButton)
    {
        _renameButton->SetEnabled(! busy && _previewStats.changedRows > 0u && _previewStats.errorRows == 0u);
    }
    if (cancelableBusy && _status)
    {
        if (_collecting)
        {
            _status->SetText(LoadBatchRenameString(IDS_BATCH_RENAME_STATUS_COLLECTING, L"Collecting items..."));
        }
        else
        {
            std::wstring status = FormatStringResource(nullptr, IDS_BATCH_RENAME_EXECUTE_PROGRESS_FMT, 0u, _previewStats.changedRows);
            if (status.empty())
            {
                status = std::format(L"Renaming {} of {}...", 0u, _previewStats.changedRows);
            }
            _status->SetText(std::move(status));
        }
    }
    _dxHost.Invalidate();
}

void BatchRenameWindow::ShowHelperMenu(BatchRenameMenus::HelperMenuKind kind, TextField* targetField, Button* anchorButton) noexcept
{
    if (! _hWnd || ! targetField || ! anchorButton || ! anchorButton->IsEnabled() || ! anchorButton->IsVisible())
    {
        return;
    }

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items = BatchRenameMenus::BuildHelperMenuItems(kind);
    if (items.empty())
    {
        return;
    }

    const D2D1_RECT_F anchorBounds = anchorButton->GetBounds();
    const POINT screenPoint        = _dxHost.DipPointToScreenPoint(D2D1::Point2F(anchorBounds.right, anchorBounds.bottom));

    ContextMenuSessionCallbacks callbacks{};
    callbacks.rootHorizontalAlignment = ContextMenuRootHorizontalAlignment::End;
    callbacks.rootVerticalPlacement   = ContextMenuRootVerticalPlacement::Below;

    anchorButton->SetPressedVisual(true);
    _dxHost.Invalidate();
    const auto clearPressed = wil::scope_exit([this, anchorButton]() noexcept
    {
        anchorButton->SetPressedVisual(false);
        _dxHost.Invalidate();
    });

    const std::optional<int> command = ContextMenu::Show(_hWnd.get(), screenPoint, items, _dxHost.GetTheme(), callbacks);
    if (command.has_value())
    {
        static_cast<void>(InsertHelperCommand(*targetField, command.value()));
    }
}

bool BatchRenameWindow::InsertHelperCommand(TextField& targetField, const int commandId) noexcept
{
    const std::wstring_view currentText = targetField.GetText();
    size_t selectionStart               = targetField.GetCaretIndex();
    size_t selectionEnd                 = selectionStart;
    if (const std::optional<std::pair<size_t, size_t>> selection = targetField.GetSelectionRange(); selection.has_value())
    {
        selectionStart = selection.value().first;
        selectionEnd   = selection.value().second;
    }

    const size_t textLength              = currentText.size();
    const size_t normalizedStart         = std::min(std::min(selectionStart, selectionEnd), textLength);
    const size_t normalizedEnd           = std::min(std::max(selectionStart, selectionEnd), textLength);
    const std::wstring_view selectedText = currentText.substr(normalizedStart, normalizedEnd > normalizedStart ? normalizedEnd - normalizedStart : 0u);

    std::optional<BatchRenameMenus::HelperCommandInsertion> helperInsertion = BatchRenameMenus::TryBuildDynamicHelperInsertion(commandId, selectedText);
    if (! helperInsertion.has_value())
    {
        const std::optional<std::wstring_view> insertion = BatchRenameMenus::TryGetHelperInsertionText(commandId);
        if (! insertion.has_value())
        {
            return false;
        }

        helperInsertion = BatchRenameMenus::HelperCommandInsertion{
            .insertionText  = std::wstring(insertion.value()),
            .selectionStart = insertion.value().size(),
            .selectionEnd   = insertion.value().size(),
        };
    }

    const BatchRenameMenus::HelperInsertionResult applied =
        BatchRenameMenus::ApplyHelperInsertion(currentText, normalizedStart, normalizedEnd, helperInsertion.value().insertionText);
    targetField.ReplaceSelectionAndNotify(helperInsertion.value().insertionText);
    if (applied.text == targetField.GetText())
    {
        targetField.SetSelectionRange(normalizedStart + helperInsertion.value().selectionStart, normalizedStart + helperInsertion.value().selectionEnd);
    }
    CancelPendingPreviewRebuild();
    RebuildPreviewFromRuleControlChange();
    return true;
}

void BatchRenameWindow::ShowPreviewContextMenu(const size_t rowIndex, const POINT screenPoint) noexcept
{
    if (! _hWnd || ! _gridModel || rowIndex >= _gridModel->GetRowCount())
    {
        return;
    }

    // ContextMenu::Show pumps a modal message loop, so a pending debounced
    // preview rebuild can re-order or filter rows while the menu is open.
    // Capture the clicked row's stable id and re-resolve it before dispatch.
    const uint64_t stableRowId = _gridModel->GetStableRowId(rowIndex);

    std::vector<MenuFlyoutItem> items;
    items.reserve(8u);
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_ORIGINAL_NAME, L"Copy Original Name"),
        .enabled   = true,
        .commandId = kBatchRenamePreviewMenuCopyOriginalName,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_NEW_NAME, L"Copy New Name"),
        .enabled   = true,
        .commandId = kBatchRenamePreviewMenuCopyNewName,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_SOURCE_PATH, L"Copy Source Path"),
        .enabled   = true,
        .commandId = kBatchRenamePreviewMenuCopySourcePath,
    });
    items.push_back(MenuFlyoutItem{.kind = MenuItemKind::Separator});
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_REVEAL_IN_PANE, L"Reveal in Active Pane"),
        .enabled   = _context.onRevealPath != nullptr,
        .commandId = kBatchRenamePreviewMenuRevealInPane,
    });
    items.push_back(MenuFlyoutItem{.kind = MenuItemKind::Separator});
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_PREVIEW_ROWS, L"Copy Preview Rows"),
        .enabled   = _gridModel->GetRowCount() != 0u,
        .commandId = kBatchRenamePreviewMenuCopyPreviewRows,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_EXECUTION_REPORT, L"Copy Execution Report"),
        .enabled   = _lastExecutionReport.has_value(),
        .commandId = kBatchRenamePreviewMenuCopyExecutionReport,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadBatchRenameString(IDS_BATCH_RENAME_MENU_COPY_UNDO_PLAN, L"Copy Undo Plan"),
        .enabled   = _lastExecutionReport.has_value() && ! _lastExecutionReport->undoEntries.empty(),
        .commandId = kBatchRenamePreviewMenuCopyUndoPlan,
    });

    const std::optional<int> command = ContextMenu::Show(_hWnd.get(), screenPoint, items, _dxHost.GetTheme());
    if (! command.has_value())
    {
        return;
    }

    const int commandId                     = command.value();
    const bool rowIndependent               = commandId == kBatchRenamePreviewMenuCopyPreviewRows || commandId == kBatchRenamePreviewMenuCopyUndoPlan ||
                                              commandId == kBatchRenamePreviewMenuCopyExecutionReport;
    const std::optional<size_t> resolvedRow = _gridModel ? _gridModel->FindRowByStableId(stableRowId) : std::nullopt;
    if (! resolvedRow.has_value() && ! rowIndependent)
    {
        return;
    }
    static_cast<void>(DispatchPreviewContextMenuCommand(commandId, resolvedRow.value_or(0u)));
}

bool BatchRenameWindow::DispatchPreviewContextMenuCommand(const int commandId, const size_t rowIndex) noexcept
{
    if (commandId == kBatchRenamePreviewMenuRevealInPane)
    {
        return RevealPreviewRowInActivePane(rowIndex);
    }

    if (commandId == kBatchRenamePreviewMenuCopyPreviewRows)
    {
        return CopyPreviewRowsToClipboard();
    }

    if (commandId == kBatchRenamePreviewMenuCopyUndoPlan)
    {
        return CopyUndoPlanToClipboard();
    }

    if (commandId == kBatchRenamePreviewMenuCopyExecutionReport)
    {
        return CopyExecutionReportToClipboard();
    }

    return CopyPreviewRowFieldToClipboard(rowIndex, commandId);
}

bool BatchRenameWindow::ActivatePreviewRow(const size_t rowIndex) noexcept
{
    return RevealPreviewRowInActivePane(rowIndex);
}

bool BatchRenameWindow::RevealPreviewRowInActivePane(const size_t rowIndex) noexcept
{
    if (! _context.onRevealPath || ! _gridModel || rowIndex >= _gridModel->GetRows().size())
    {
        return false;
    }

    const BatchPreviewRow& row = _gridModel->GetRows()[rowIndex];
    if (row.sourcePath.empty())
    {
        return false;
    }

    return _context.onRevealPath(row.sourcePath);
}

bool BatchRenameWindow::CopyPreviewRowsToClipboard() const noexcept
{
    if (! _gridModel || _gridModel->GetRowCount() == 0u)
    {
        return false;
    }

    const std::wstring text = BuildPreviewRowsClipboardText();
    return ! text.empty() && _dxHost.CopyTextToClipboard(text);
}

bool BatchRenameWindow::CopyPreviewRowFieldToClipboard(const size_t rowIndex, const int commandId) const noexcept
{
    if (! _gridModel || rowIndex >= _gridModel->GetRows().size())
    {
        return false;
    }

    const BatchPreviewRow& row = _gridModel->GetRows()[rowIndex];
    std::wstring_view text;
    switch (commandId)
    {
        case kBatchRenamePreviewMenuCopyOriginalName: text = row.originalName; break;
        case kBatchRenamePreviewMenuCopyNewName: text = row.newName; break;
        case kBatchRenamePreviewMenuCopySourcePath: text = row.fullPath; break;
        default: return false;
    }

    return _dxHost.CopyTextToClipboard(text);
}

bool BatchRenameWindow::CopyExecutionReportToClipboard() const noexcept
{
    const std::wstring text = BuildExecutionReportClipboardText();
    return ! text.empty() && _dxHost.CopyTextToClipboard(text);
}

bool BatchRenameWindow::CopyUndoPlanToClipboard() const noexcept
{
    const std::wstring text = BuildUndoPlanClipboardText();
    return ! text.empty() && _dxHost.CopyTextToClipboard(text);
}

std::wstring BatchRenameWindow::BuildPreviewRowsClipboardText() const
{
    if (! _gridModel || _gridModel->GetRowCount() == 0u)
    {
        return {};
    }

    std::vector<size_t> columnOrder;
    std::vector<bool> usedColumns(_gridModel->GetColumnCount(), false);
    if (_grid)
    {
        for (const GridColumnLayoutEntry& entry : _grid->CaptureColumnLayout())
        {
            const std::optional<size_t> columnIndex = FindPreviewColumnIndexById(entry.columnId);
            if (! columnIndex.has_value() || columnIndex.value() >= usedColumns.size() || usedColumns[columnIndex.value()])
            {
                continue;
            }
            columnOrder.push_back(columnIndex.value());
            usedColumns[columnIndex.value()] = true;
        }
    }
    for (size_t columnIndex = 0u; columnIndex < _gridModel->GetColumnCount(); ++columnIndex)
    {
        if (! usedColumns[columnIndex])
        {
            columnOrder.push_back(columnIndex);
        }
    }

    std::wstring text;
    for (size_t displayIndex = 0u; displayIndex < columnOrder.size(); ++displayIndex)
    {
        if (displayIndex > 0u)
        {
            text.push_back(L'\t');
        }
        AppendBatchRenameTsvField(text, _gridModel->GetColumn(columnOrder[displayIndex]).title);
    }

    for (const BatchPreviewRow& row : _gridModel->GetRows())
    {
        text.append(L"\r\n");
        for (size_t displayIndex = 0u; displayIndex < columnOrder.size(); ++displayIndex)
        {
            if (displayIndex > 0u)
            {
                text.push_back(L'\t');
            }
            AppendBatchRenameTsvField(text, BatchPreviewRowClipboardField(row, columnOrder[displayIndex]));
        }
    }

    return text;
}

std::wstring BatchRenameWindow::BuildExecutionReportClipboardText() const
{
    if (! _lastExecutionReport.has_value())
    {
        return {};
    }

    const BatchRenameExecutionReport& report = _lastExecutionReport.value();
    std::wstring text;
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_TOTAL_ROWS, L"Total Rows"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_COMPLETED_ROWS, L"Completed Rows"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_SKIPPED_ROWS, L"Skipped Rows"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_FAILED_ROWS, L"Failed Rows"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_CANCELED, L"Canceled"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_FIRST_FAILURE, L"First Failure"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_REPORT_COL_FIRST_FAILURE_TEXT, L"First Failure Text"));

    text.append(L"\r\n");
    AppendBatchRenameTsvField(text, std::to_wstring(report.totalRows));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, std::to_wstring(report.completedRows));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, std::to_wstring(report.skippedRows));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, std::to_wstring(report.failedRows));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, report.canceled ? L"true" : L"false");
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, std::format(L"0x{:08X}", static_cast<unsigned long>(report.firstFailure)));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, report.firstFailureText);

    return text;
}

std::wstring BatchRenameWindow::BuildUndoPlanClipboardText() const
{
    if (! _lastExecutionReport.has_value() || _lastExecutionReport->undoEntries.empty())
    {
        return {};
    }

    std::wstring text;
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_UNDO_COL_CURRENT_PATH, L"Current Path"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_UNDO_COL_RESTORE_NAME, L"Restore Name"));
    text.push_back(L'\t');
    AppendBatchRenameTsvField(text, LoadBatchRenameString(IDS_BATCH_RENAME_UNDO_COL_ORIGINAL_PATH, L"Original Path"));

    for (const BatchRenameUndoEntry& entry : _lastExecutionReport->undoEntries)
    {
        text.append(L"\r\n");
        AppendBatchRenameTsvField(text, entry.currentPath.native());
        text.push_back(L'\t');
        AppendBatchRenameTsvField(text, entry.restoreName);
        text.push_back(L'\t');
        AppendBatchRenameTsvField(text, entry.originalPath.native());
    }

    return text;
}

void BatchRenameWindow::RebuildPreviewFromRuleControlChange()
{
    ClearExecutionReport();
    RebuildPreview();
    _dxHost.Invalidate();
}

void BatchRenameWindow::RefreshPreviewPresentation() noexcept
{
    // View-only refresh (sorting, visibility filtering): the plan is
    // unchanged, so the retained execution report and its undo plan survive.
    RefreshVisibleRows();
    _dxHost.Invalidate();
}

void BatchRenameWindow::UpdateStatus() noexcept
{
    if (! _status || ! _gridModel)
    {
        return;
    }

    if (_lastExecutionReport.has_value())
    {
        const BatchRenameExecutionReport& report = _lastExecutionReport.value();
        std::wstring status;
        if (report.canceled || report.failedRows != 0u || FAILED(report.firstFailure))
        {
            status = FormatStringResource(nullptr,
                                          IDS_BATCH_RENAME_EXECUTE_STATUS_FAILED_FMT,
                                          report.totalRows,
                                          report.completedRows,
                                          report.skippedRows,
                                          report.failedRows,
                                          report.firstFailureText);
            if (status.empty())
            {
                status = std::format(L"{} planned, {} renamed, {} skipped, {} failed. First failure: {}",
                                     report.totalRows,
                                     report.completedRows,
                                     report.skippedRows,
                                     report.failedRows,
                                     report.firstFailureText);
            }
        }
        else
        {
            status = FormatStringResource(
                nullptr, IDS_BATCH_RENAME_EXECUTE_STATUS_FMT, report.totalRows, report.completedRows, report.skippedRows, report.failedRows);
            if (status.empty())
            {
                status = std::format(
                    L"{} planned, {} renamed, {} skipped, {} failed", report.totalRows, report.completedRows, report.skippedRows, report.failedRows);
            }
        }
        _status->SetText(std::move(status));
        return;
    }

    std::wstring status = FormatStringResource(nullptr,
                                               IDS_BATCH_RENAME_STATUS_FMT,
                                               _previewStats.totalRows,
                                               _previewStats.changedRows,
                                               _previewStats.unchangedRows,
                                               _previewStats.errorRows,
                                               _previewStats.warningRows);
    if (status.empty())
    {
        status = std::format(L"{} items, {} changed, {} unchanged, {} errors, {} warnings",
                             _previewStats.totalRows,
                             _previewStats.changedRows,
                             _previewStats.unchangedRows,
                             _previewStats.errorRows,
                             _previewStats.warningRows);
    }
    _status->SetText(std::move(status));
}

void BatchRenameWindow::SetHideUnchangedRows(const bool hideUnchanged) noexcept
{
    if (_hideUnchangedRows == hideUnchanged)
    {
        if (_hideUnchangedCheck)
        {
            _hideUnchangedCheck->SetChecked(hideUnchanged);
        }
        return;
    }

    _hideUnchangedRows = hideUnchanged;
    if (_hideUnchangedCheck)
    {
        _hideUnchangedCheck->SetChecked(hideUnchanged);
    }
    RefreshPreviewPresentation();
}

void BatchRenameWindow::RefreshRootNavigationPath() noexcept
{
    if (! _rootNavigation.GetHwnd())
    {
        return;
    }

    if (_rootText.empty())
    {
        _rootNavigation.SetPath(std::nullopt);
        return;
    }

    _rootNavigation.SetPath(std::filesystem::path(_rootText));
}

void BatchRenameWindow::RefreshRootNavigationHistory() noexcept
{
    if (! _rootNavigation.GetHwnd())
    {
        return;
    }

    std::vector<std::filesystem::path> history;
    if (! _rootText.empty())
    {
        history.emplace_back(_rootText);
    }

    if (_settings && _settings->batchRename.has_value() && ! _settings->batchRename->lastRoot.empty() && _settings->batchRename->lastRoot != _rootText)
    {
        history.emplace_back(_settings->batchRename->lastRoot);
    }

    _rootNavigation.SetHistory(history);
}

void BatchRenameWindow::OnRootNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept
{
    if (! path.has_value() || path->empty())
    {
        RefreshRootNavigationPath();
        return;
    }

    const std::wstring nextRoot = path->native();
    if (nextRoot.empty() || nextRoot == _rootText)
    {
        RefreshRootNavigationPath();
        return;
    }

    ClearExecutionReport();
    _rootText               = nextRoot;
    _context.rootPluginPath = path.value();
    _context.initialPaths.clear();
    RefreshRootNavigationHistory();
    RebuildTargetsFromScope();
}

std::optional<size_t> BatchRenameWindow::FindPreviewColumnIndexById(std::wstring_view columnId) const noexcept
{
    if (! _gridModel || columnId.empty())
    {
        return std::nullopt;
    }

    for (size_t index = 0u; index < _gridModel->GetColumnCount(); ++index)
    {
        if (_gridModel->GetColumn(index).id == columnId)
        {
            return index;
        }
    }
    return std::nullopt;
}

void BatchRenameWindow::OnGridSortRequested(const GridSortSpec& sortSpec)
{
    if (! _grid || ! _gridModel)
    {
        return;
    }

    _grid->SetSortSpec(sortSpec);
    RefreshPreviewPresentation();
}

void BatchRenameWindow::OnGridRowActivated(Grid& sender, const size_t rowIndex)
{
    if (&sender != _grid)
    {
        return;
    }

    static_cast<void>(ActivatePreviewRow(rowIndex));
}

void BatchRenameWindow::OnGridContextMenu(Grid& sender, const size_t rowIndex, const POINT screenPoint)
{
    if (&sender != _grid)
    {
        return;
    }

    ShowPreviewContextMenu(rowIndex, screenPoint);
}

wil::com_ptr<ID2D1Bitmap1> BatchRenameWindow::GetGridIconBitmap(const Grid& sender,
                                                                const int iconIndex,
                                                                const float targetDipSize,
                                                                ID2D1DeviceContext* d2dContext)
{
    if (&sender != _grid || iconIndex < 0 || ! d2dContext)
    {
        return nullptr;
    }

    return IconCache::GetInstance().GetIconBitmap(iconIndex, d2dContext, targetDipSize);
}

#ifdef ENABLE_TESTS
bool BatchRenameWindow::DebugGetSnapshot(BatchRenameDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
    {
        return false;
    }

    out.usesDxUiHost = _dxHost.GetHwnd() == _hWnd.get();
    out.rootNavigationVisible =
        _rootNavigation.GetHwnd() != nullptr && IsWindow(_rootNavigation.GetHwnd()) != FALSE && IsWindowVisible(_rootNavigation.GetHwnd()) != FALSE;
    NavigationViewDebugSnapshot rootNavigationSnapshot{};
    out.rootNavigationUsesNavigationView = out.rootNavigationVisible && _rootNavigation.DebugGetSnapshot(rootNavigationSnapshot);
    if (out.rootNavigationUsesNavigationView)
    {
        out.rootNavigationPathText = std::move(rootNavigationSnapshot.currentPathText);
    }
    out.ruleControlsVisible = _nameTemplateField != nullptr && _nameTemplateField->IsVisible() && _searchForField != nullptr && _searchForField->IsVisible() &&
                              _replaceWithField != nullptr && _replaceWithField->IsVisible() && _fileNameCaseCombo != nullptr &&
                              _fileNameCaseCombo->IsVisible() && _extensionCaseCombo != nullptr && _extensionCaseCombo->IsVisible();
    out.ruleHelperButtonsVisible = _nameTemplateHelperButton != nullptr && _nameTemplateHelperButton->IsVisible() && _searchForHelperButton != nullptr &&
                                   _searchForHelperButton->IsVisible() && _replaceWithHelperButton != nullptr && _replaceWithHelperButton->IsVisible();
    out.rulesModeSelected        = _rules.mode == BatchRename::Mode::Rules;
    out.manualModeSelected       = _rules.mode == BatchRename::Mode::Manual;
    out.manualControlsVisible    = _manualNamesField != nullptr && _manualNamesField->IsVisible();
    out.renameButtonEnabled      = _renameButton != nullptr && _renameButton->IsEnabled();
    out.hideUnchangedRows        = _hideUnchangedRows;
    out.previewRebuildPending    = _previewRebuildPending;
    out.visibleChildWindowCount  = CountVisibleChildWindows(_hWnd.get());
    if (_root)
    {
        CollectBatchRenameFocusableAccessibleNames(_root, out.focusableAccessibleNames);
    }
    out.rootText = _rootText;
    if (_status)
    {
        out.statusText = std::wstring(_status->GetText());
    }
    if (_maskField)
    {
        out.scopeMaskText = std::wstring(_maskField->GetText());
    }
    out.includeSubdirectories = _includeSubdirectoriesCheck != nullptr && _includeSubdirectoriesCheck->IsChecked();
    out.includeFiles          = _includeFilesCheck != nullptr && _includeFilesCheck->IsChecked();
    out.includeFolders        = _includeFoldersCheck != nullptr && _includeFoldersCheck->IsChecked();
    out.changedRowCount       = _previewStats.changedRows;
    out.errorRowCount         = _previewStats.errorRows;
    out.warningRowCount       = _previewStats.warningRows;
    if (_lastExecutionReport.has_value())
    {
        const BatchRenameExecutionReport& report = _lastExecutionReport.value();
        out.hasExecutionReport                   = true;
        out.lastExecutionTotalRows               = report.totalRows;
        out.lastExecutionCompletedRows           = report.completedRows;
        out.lastExecutionSkippedRows             = report.skippedRows;
        out.lastExecutionFailedRows              = report.failedRows;
        out.lastExecutionUndoRowCount            = report.undoEntries.size();
        out.lastExecutionFirstFailure            = report.firstFailure;
        out.lastExecutionCanceled                = report.canceled;
        out.lastExecutionFirstFailureText        = report.firstFailureText;
    }
    if (_nameTemplateField)
    {
        out.nameTemplateText = std::wstring(_nameTemplateField->GetText());
    }
    if (_searchForField)
    {
        out.searchForText = std::wstring(_searchForField->GetText());
    }
    if (_replaceWithField)
    {
        out.replaceWithText = std::wstring(_replaceWithField->GetText());
    }
    if (_manualNamesField)
    {
        out.manualText = std::wstring(_manualNamesField->GetText());
    }
    out.regexEnabled     = _regexCheck != nullptr && _regexCheck->IsChecked();
    out.caseSensitive    = _caseSensitiveCheck != nullptr && _caseSensitiveCheck->IsChecked();
    out.wholeWords       = _wholeWordsCheck != nullptr && _wholeWordsCheck->IsChecked();
    out.replaceOnce      = _replaceOnceCheck != nullptr && _replaceOnceCheck->IsChecked();
    out.excludeExtension = _excludeExtensionCheck != nullptr && _excludeExtensionCheck->IsChecked();
    if (_fileNameCaseCombo)
    {
        out.fileNameCaseText = std::wstring(_fileNameCaseCombo->GetDisplayedText());
    }
    if (_extensionCaseCombo)
    {
        out.extensionCaseText = std::wstring(_extensionCaseCombo->GetDisplayedText());
    }

    if (_gridModel)
    {
        out.previewRowCount = _gridModel->GetRowCount();
        for (const GridColumnDesc& column : _gridModel->GetColumns())
        {
            out.previewColumnIds.push_back(column.id);
        }
        for (const BatchPreviewRow& row : _gridModel->GetRows())
        {
            out.originalNames.push_back(row.originalName);
            out.newNames.push_back(row.newName);
            out.sizeTexts.push_back(row.sizeText);
            out.dateTexts.push_back(row.dateText);
            out.timeTexts.push_back(row.timeText);
            out.fullPaths.push_back(row.fullPath);
            out.originalIconIndices.push_back(row.iconIndex);
        }
        for (size_t rowIndex = 0u; rowIndex < _gridModel->GetRowCount(); ++rowIndex)
        {
            GridCellData cell{};
            _gridModel->GetCellData(rowIndex, 0u, cell);
            if (cell.kind == GridCellKind::IconText && (! cell.iconText.empty() || cell.iconIndex >= 0))
            {
                ++out.previewIconCellCount;
            }
            GridCellData newNameCell{};
            _gridModel->GetCellData(rowIndex, 1u, newNameCell);
            out.newNameStatusIconTexts.push_back(newNameCell.iconText);
            out.newNameTooltips.push_back(newNameCell.tooltipText);
        }
    }
    return true;
}

void BatchRenameWindow::DebugSetRules(BatchRename::Rules rules) noexcept
{
    ClearExecutionReport();
    _rules = std::move(rules);
    SyncRuleControls();
    RebuildPreview();
    // The production preview is asynchronous. Test callers need the same
    // settled contract the previous synchronous implementation exposed.
    DebugPumpWhileTasksActive(false);
    _dxHost.Invalidate();
}

bool BatchRenameWindow::DebugSetRuleControls(const BatchRename::Rules& rules) noexcept
{
    if (! _nameTemplateField || ! _searchForField || ! _replaceWithField || ! _regexCheck || ! _caseSensitiveCheck || ! _wholeWordsCheck ||
        ! _replaceOnceCheck || ! _excludeExtensionCheck || ! _fileNameCaseCombo || ! _extensionCaseCombo)
    {
        return false;
    }

    ClearExecutionReport();
    _rules.nameTemplate       = rules.nameTemplate;
    _rules.searchFor          = rules.searchFor;
    _rules.replaceWith        = rules.replaceWith;
    _rules.regexEnabled       = rules.regexEnabled;
    _rules.caseSensitive      = rules.caseSensitive;
    _rules.wholeWords         = rules.wholeWords;
    _rules.replaceOnce        = rules.replaceOnce;
    _rules.excludeExtension   = rules.excludeExtension;
    _rules.flattenSeparator   = rules.flattenSeparator;
    _rules.fileNameCaseStyle  = rules.fileNameCaseStyle;
    _rules.extensionCaseStyle = rules.extensionCaseStyle;
    SyncRuleControls();
    RebuildPreview();
    DebugPumpWhileTasksActive(false);
    _dxHost.Invalidate();
    return true;
}

bool BatchRenameWindow::DebugSetScope(const std::wstring_view mask,
                                      const bool includeSubdirectories,
                                      const bool includeFiles,
                                      const bool includeFolders) noexcept
{
    if (! _maskField || ! _includeSubdirectoriesCheck || ! _includeFilesCheck || ! _includeFoldersCheck)
    {
        return false;
    }

    _scopeOptions.mask                  = std::wstring(mask);
    _scopeOptions.includeSubdirectories = includeSubdirectories;
    _scopeOptions.includeFiles          = includeFiles;
    _scopeOptions.includeFolders        = includeFolders;

    _syncingRuleControls = true;
    const auto clearSync = wil::scope_exit([this]() noexcept { _syncingRuleControls = false; });
    _maskField->SetText(_scopeOptions.mask);
    _includeSubdirectoriesCheck->SetChecked(_scopeOptions.includeSubdirectories);
    _includeFilesCheck->SetChecked(_scopeOptions.includeFiles);
    _includeFoldersCheck->SetChecked(_scopeOptions.includeFolders);

    RebuildTargetsFromScope();
    DebugPumpWhileTasksActive(false);
    _dxHost.Invalidate();
    return true;
}

bool BatchRenameWindow::DebugSetPreviewSort(const std::wstring_view columnId, const bool descending) noexcept
{
    const std::optional<size_t> columnIndex = FindPreviewColumnIndexById(columnId);
    if (! columnIndex.has_value())
    {
        return false;
    }

    OnGridSortRequested(GridSortSpec{
        .columnIndex = columnIndex.value(),
        .direction   = descending ? SortDirection::Descending : SortDirection::Ascending,
    });
    return true;
}

bool BatchRenameWindow::DebugReorderPreviewColumn(const std::wstring_view columnId, const size_t targetDisplayIndex) noexcept
{
    if (! _grid || ! _gridModel || columnId.empty())
    {
        return false;
    }

    if (! FindPreviewColumnIndexById(columnId).has_value())
    {
        return false;
    }

    std::vector<GridColumnLayoutEntry> layout = _grid->CaptureColumnLayout();
    if (layout.size() != _gridModel->GetColumnCount())
    {
        return false;
    }

    const auto movedIt = std::ranges::find_if(layout, [&](const GridColumnLayoutEntry& entry) noexcept { return entry.columnId == columnId; });
    if (movedIt == layout.end())
    {
        return false;
    }

    GridColumnLayoutEntry moved = std::move(*movedIt);
    layout.erase(movedIt);
    layout.insert(layout.begin() + static_cast<std::ptrdiff_t>(std::min(targetDisplayIndex, layout.size())), std::move(moved));
    for (size_t displayIndex = 0u; displayIndex < layout.size(); ++displayIndex)
    {
        layout[displayIndex].displayIndex = displayIndex;
    }

    _grid->ApplyColumnLayout(layout);
    _dxHost.Invalidate();
    return true;
}

bool BatchRenameWindow::DebugSwitchMode(const BatchRename::Mode mode) noexcept
{
    if (! _modeSelector || ! _manualNamesField)
    {
        return false;
    }
    SwitchMode(mode);
    return true;
}

bool BatchRenameWindow::DebugSetManualText(std::wstring_view text) noexcept
{
    if (! _manualNamesField)
    {
        return false;
    }
    ClearExecutionReport();
    SetManualText(std::wstring(text));
    UpdateModeVisibility();
    Layout();
    RebuildPreview();
    _dxHost.Invalidate();
    return true;
}

bool BatchRenameWindow::DebugClickManualPaste() noexcept
{
    if (! _manualPasteButton || ! _manualPasteButton->IsVisible() || ! _manualPasteButton->IsEnabled())
    {
        return false;
    }
    return PasteManualFromClipboard();
}

bool BatchRenameWindow::DebugClickManualSortLikePreview() noexcept
{
    if (! _manualSortLikePreviewButton || ! _manualSortLikePreviewButton->IsVisible() || ! _manualSortLikePreviewButton->IsEnabled())
    {
        return false;
    }
    return SortManualTextLikePreview();
}

bool BatchRenameWindow::DebugSetHideUnchanged(const bool hideUnchanged) noexcept
{
    SetHideUnchangedRows(hideUnchanged);
    return _hideUnchangedRows == hideUnchanged;
}

TextField* BatchRenameWindow::DebugResolveRuleField(const BatchRenameDebugRuleField field) noexcept
{
    switch (field)
    {
        case BatchRenameDebugRuleField::NameTemplate: return _nameTemplateField;
        case BatchRenameDebugRuleField::SearchFor: return _searchForField;
        case BatchRenameDebugRuleField::ReplaceWith: return _replaceWithField;
        default: return nullptr;
    }
}

bool BatchRenameWindow::DebugSetRuleFieldSelection(const BatchRenameDebugRuleField field, const size_t selectionStart, const size_t selectionEnd) noexcept
{
    TextField* const textField = DebugResolveRuleField(field);
    if (! textField)
    {
        return false;
    }
    textField->SetSelectionRange(selectionStart, selectionEnd);
    return true;
}

bool BatchRenameWindow::DebugSetRuleFieldText(const BatchRenameDebugRuleField field, const std::wstring_view text) noexcept
{
    TextField* const textField = DebugResolveRuleField(field);
    if (! textField)
    {
        return false;
    }

    textField->SetTextAndNotify(std::wstring(text));
    return true;
}

bool BatchRenameWindow::DebugInsertHelperCommand(const BatchRenameDebugRuleField field, const int commandId) noexcept
{
    TextField* const textField = DebugResolveRuleField(field);
    return textField != nullptr && InsertHelperCommand(*textField, commandId);
}

bool BatchRenameWindow::DebugCopyPreview(const BatchRenameDebugPreviewCopyKind kind, const size_t rowIndex) noexcept
{
    int commandId = 0;
    switch (kind)
    {
        case BatchRenameDebugPreviewCopyKind::OriginalName: commandId = kBatchRenamePreviewMenuCopyOriginalName; break;
        case BatchRenameDebugPreviewCopyKind::NewName: commandId = kBatchRenamePreviewMenuCopyNewName; break;
        case BatchRenameDebugPreviewCopyKind::SourcePath: commandId = kBatchRenamePreviewMenuCopySourcePath; break;
        case BatchRenameDebugPreviewCopyKind::PreviewRows: commandId = kBatchRenamePreviewMenuCopyPreviewRows; break;
        default: return false;
    }
    return DispatchPreviewContextMenuCommand(commandId, rowIndex);
}

bool BatchRenameWindow::DebugRevealPreview(const size_t rowIndex) noexcept
{
    return RevealPreviewRowInActivePane(rowIndex);
}

bool BatchRenameWindow::DebugActivatePreview(const size_t rowIndex) noexcept
{
    return ActivatePreviewRow(rowIndex);
}

bool BatchRenameWindow::DebugCopyExecutionReport() const noexcept
{
    return CopyExecutionReportToClipboard();
}

bool BatchRenameWindow::DebugCopyUndoPlan() const noexcept
{
    return CopyUndoPlanToClipboard();
}

bool BatchRenameWindow::DebugFlushPendingPreview() noexcept
{
    if (! _previewRebuildPending && ! _targetRebuildPending)
    {
        return false;
    }

    OnPreviewRebuildTimer();
    DebugPumpWhileTasksActive(false);
    return true;
}

HRESULT BatchRenameWindow::DebugExecute() noexcept
{
    const HRESULT startHr = ExecuteRename();
    if (! _executing)
    {
        return startHr;
    }

    // Execution runs on a background worker; pump until the completion
    // payload arrives so test callers observe the terminal result.
    DebugPumpWhileTasksActive(true);
    return _lastExecutionTerminalHr;
}

HRESULT BatchRenameWindow::DebugStartExecution() noexcept
{
    // Starts the asynchronous execution without pumping so tests can observe
    // the in-flight busy state (e.g. ExecuteRename returning ERROR_BUSY).
    return ExecuteRename();
}

HRESULT BatchRenameWindow::DebugWaitExecutionIdle() noexcept
{
    DebugPumpWhileTasksActive(true);
    return _lastExecutionTerminalHr;
}

bool BatchRenameWindow::DebugInjectStaleCollectionPayload(std::filesystem::path sourcePath) noexcept
{
    auto payload = std::unique_ptr<BatchRenameCollectionCompletedPayload>(new (std::nothrow) BatchRenameCollectionCompletedPayload{});
    if (! payload)
    {
        return false;
    }

    BatchRename::Target target{};
    target.sourcePath   = std::move(sourcePath);
    payload->generation = _taskGeneration.load(std::memory_order_acquire) + 1u;
    payload->targets.push_back(std::move(target));
    OnCollectionCompleted(std::move(payload));
    return true;
}

bool BatchRenameWindow::DebugInjectStaleExecutionPayload(std::filesystem::path sourcePath, std::filesystem::path targetPath) noexcept
{
    auto payload = std::unique_ptr<BatchRenameExecutionCompletedPayload>(new (std::nothrow) BatchRenameExecutionCompletedPayload{});
    if (! payload)
    {
        return false;
    }

    payload->generation           = _taskGeneration.load(std::memory_order_acquire) + 1u;
    payload->report.totalRows     = 1u;
    payload->report.completedRows = 1u;
    payload->successfulSourcePaths.push_back(std::move(sourcePath));
    payload->successfulTargetPaths.push_back(std::move(targetPath));
    OnExecutionCompleted(std::move(payload));
    return true;
}

void BatchRenameWindow::DebugPumpWhileTasksActive(const bool waitForExecution) noexcept
{
    const ULONGLONG deadline = GetTickCount64() + 30000ull;
    const auto tasksActive = [this, waitForExecution]() noexcept {
        return _hWnd && (waitForExecution ? (_executing || _collecting || _previewing) : (_collecting || _previewing));
    };

    while (tasksActive() && GetTickCount64() < deadline)
    {
        static_cast<void>(MsgWaitForMultipleObjects(0, nullptr, FALSE, 50, QS_ALLINPUT));
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT)
            {
                PostQuitMessage(static_cast<int>(msg.wParam));
                return;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);

            // Re-check after every dispatch: animations plus slow
            // perf-instrumented renders can keep the queue non-empty
            // indefinitely, so waiting for the queue to drain would never
            // observe the completed task or the deadline.
            if (! tasksActive() || GetTickCount64() >= deadline)
            {
                return;
            }
        }
    }
}
#endif
} // namespace

bool ShowBatchRenameWindow(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, BatchRenamePaneContext context) noexcept
{
    if (g_batchRenameWindow && g_batchRenameWindow->Hwnd() && IsWindow(g_batchRenameWindow->Hwnd()) != FALSE)
    {
        g_batchRenameWindow->UpdateTheme(theme);
        g_batchRenameWindow->SetContext(std::move(context));
        ShowWindow(g_batchRenameWindow->Hwnd(), SW_SHOWNORMAL);
        SetForegroundWindow(g_batchRenameWindow->Hwnd());
        return true;
    }

    auto* window = new (std::nothrow) BatchRenameWindow();
    if (! window)
    {
        return false;
    }

    // Create() owns failure cleanup: when WM_CREATE fails, the window
    // procedure already deleted the instance while CreateWindowExW was
    // unwinding, so a caller-side delete here would be a double delete.
    const HWND hwnd = window->Create(owner, settings, theme, std::move(context));
    return hwnd != nullptr;
}

void UpdateBatchRenameWindowsTheme(const AppTheme& theme) noexcept
{
    if (g_batchRenameWindow && g_batchRenameWindow->Hwnd() && IsWindow(g_batchRenameWindow->Hwnd()) != FALSE)
    {
        g_batchRenameWindow->UpdateTheme(theme);
    }
}

HWND GetBatchRenameWindowHandle() noexcept
{
    return g_batchRenameWindow ? g_batchRenameWindow->Hwnd() : nullptr;
}

bool IsBatchRenameWindowHandle(HWND hwnd) noexcept
{
    return hwnd != nullptr && g_batchRenameWindow != nullptr && hwnd == g_batchRenameWindow->Hwnd();
}

#ifdef ENABLE_TESTS
size_t DebugGetBatchRenameWindowCount() noexcept
{
    return g_batchRenameWindow && g_batchRenameWindow->Hwnd() && IsWindow(g_batchRenameWindow->Hwnd()) != FALSE ? 1u : 0u;
}

bool DebugGetBatchRenameWindowSnapshot(BatchRenameDebugSnapshot& out) noexcept
{
    if (! g_batchRenameWindow)
    {
        out = {};
        return false;
    }
    // Settle any in-flight background target collection first so snapshots
    // keep observing collected targets like the previous synchronous path.
    g_batchRenameWindow->DebugPumpWhileTasksActive(false);
    return g_batchRenameWindow->DebugGetSnapshot(out);
}

bool DebugSetBatchRenameWindowRules(const BatchRename::Rules& rules) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    g_batchRenameWindow->DebugSetRules(rules);
    return true;
}

bool DebugSetBatchRenameWindowRuleControls(const BatchRename::Rules& rules) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetRuleControls(rules);
}

bool DebugSetBatchRenameWindowScope(const std::wstring_view mask, const bool includeSubdirectories, const bool includeFiles, const bool includeFolders) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetScope(mask, includeSubdirectories, includeFiles, includeFolders);
}

bool DebugSetBatchRenameWindowPreviewSort(const std::wstring_view columnId, const bool descending) noexcept
{
    if (! g_batchRenameWindow)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetPreviewSort(columnId, descending);
}

bool DebugReorderBatchRenameWindowPreviewColumn(const std::wstring_view columnId, const size_t targetDisplayIndex) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugReorderPreviewColumn(columnId, targetDisplayIndex);
}

bool DebugSwitchBatchRenameWindowMode(const BatchRename::Mode mode) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSwitchMode(mode);
}

bool DebugSetBatchRenameWindowManualText(std::wstring_view text) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetManualText(text);
}

bool DebugClickBatchRenameWindowManualPaste() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugClickManualPaste();
}

bool DebugClickBatchRenameWindowManualSortLikePreview() noexcept
{
    if (! g_batchRenameWindow)
    {
        return false;
    }
    return g_batchRenameWindow->DebugClickManualSortLikePreview();
}

bool DebugSetBatchRenameWindowRuleFieldSelection(const BatchRenameDebugRuleField field, const size_t selectionStart, const size_t selectionEnd) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetRuleFieldSelection(field, selectionStart, selectionEnd);
}

bool DebugSetBatchRenameWindowRuleFieldText(const BatchRenameDebugRuleField field, const std::wstring_view text) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetRuleFieldText(field, text);
}

bool DebugInsertBatchRenameWindowHelperCommand(const BatchRenameDebugRuleField field, const int commandId) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugInsertHelperCommand(field, commandId);
}

bool DebugCopyBatchRenameWindowPreview(const BatchRenameDebugPreviewCopyKind kind, const size_t rowIndex) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugCopyPreview(kind, rowIndex);
}

bool DebugRevealBatchRenameWindowPreview(const size_t rowIndex) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugRevealPreview(rowIndex);
}

bool DebugActivateBatchRenameWindowPreview(const size_t rowIndex) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugActivatePreview(rowIndex);
}

bool DebugCopyBatchRenameWindowExecutionReport() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugCopyExecutionReport();
}

bool DebugCopyBatchRenameWindowUndoPlan() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugCopyUndoPlan();
}

bool DebugSetBatchRenameWindowHideUnchanged(const bool hideUnchanged) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugSetHideUnchanged(hideUnchanged);
}

bool DebugFlushBatchRenameWindowPendingPreview() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugFlushPendingPreview();
}

HRESULT DebugExecuteBatchRenameWindow() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    return g_batchRenameWindow->DebugExecute();
}

HRESULT DebugStartBatchRenameWindowExecution() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    return g_batchRenameWindow->DebugStartExecution();
}

HRESULT DebugWaitBatchRenameWindowExecutionIdle() noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    return g_batchRenameWindow->DebugWaitExecutionIdle();
}

bool DebugInjectStaleBatchRenameWindowCollectionPayload(std::filesystem::path sourcePath) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugInjectStaleCollectionPayload(std::move(sourcePath));
}

bool DebugInjectStaleBatchRenameWindowExecutionPayload(std::filesystem::path sourcePath, std::filesystem::path targetPath) noexcept
{
    if (! g_batchRenameWindow || ! g_batchRenameWindow->Hwnd() || IsWindow(g_batchRenameWindow->Hwnd()) == FALSE)
    {
        return false;
    }
    return g_batchRenameWindow->DebugInjectStaleExecutionPayload(std::move(sourcePath), std::move(targetPath));
}

void DebugSetBatchRenameWindowDestinationProbeFailurePath(std::filesystem::path destinationPath, const unsigned long win32Error)
{
    std::lock_guard lock(g_batchRenameDebugDestinationProbeFailureMutex);
    g_batchRenameDebugDestinationProbeFailurePath       = std::move(destinationPath);
    g_batchRenameDebugDestinationProbeFailureWin32Error = win32Error;
}

void DebugClearBatchRenameWindowDestinationProbeFailurePath() noexcept
{
    std::lock_guard lock(g_batchRenameDebugDestinationProbeFailureMutex);
    g_batchRenameDebugDestinationProbeFailurePath.reset();
    g_batchRenameDebugDestinationProbeFailureWin32Error = ERROR_ACCESS_DENIED;
}

bool DebugCollectBatchRenameTargetsForTests(BatchRenamePaneContext context,
                                            const std::wstring_view mask,
                                            const bool includeSubdirectories,
                                            const bool includeFiles,
                                            const bool includeFolders,
                                            BatchRenameDebugCollectionResult& out)
{
    BatchRenameScopeOptions scope{};
    scope.mask                  = std::wstring(mask);
    scope.includeSubdirectories = includeSubdirectories;
    scope.includeFiles          = includeFiles;
    scope.includeFolders        = includeFolders;

    BatchRenameTargetCollectionResult result = CollectBatchRenameTargets(context, scope);
    out                                      = {};
    out.hr                                   = result.hr;
    out.detail                               = std::move(result.detail);
    out.originalNames.reserve(result.targets.size());
    out.fullPaths.reserve(result.targets.size());
    out.isDirectories.reserve(result.targets.size());
    out.metadataUnknowns.reserve(result.targets.size());
    out.sizeBytes.reserve(result.targets.size());
    for (const BatchRename::Target& target : result.targets)
    {
        out.originalNames.push_back(target.sourcePath.filename().native());
        out.fullPaths.push_back(target.sourcePath.native());
        out.isDirectories.push_back(target.isDirectory);
        out.metadataUnknowns.push_back(target.metadataUnknown);
        out.sizeBytes.push_back(target.sizeBytes);
    }
    return true;
}

bool DebugRefreshBatchRenameTargetsAfterExecutionForTests(const FileSystemPathIdentity& pathIdentity,
                                                          std::vector<BatchRename::Target>& targets,
                                                          std::span<const std::filesystem::path> successfulSourcePaths,
                                                          std::span<const std::filesystem::path> successfulTargetPaths,
                                                          const std::filesystem::path& root,
                                                          size_t& refreshedRows,
                                                          uint64_t& identityComparisons) noexcept
{
    const BatchRenameTargetRefreshResult result = RefreshBatchRenameTargetsAfterExecution(
        pathIdentity, targets, successfulSourcePaths, successfulTargetPaths, std::span<const ExecutedDirectoryMove>{}, root);
    refreshedRows       = result.refreshedRows;
    identityComparisons = result.identityComparisons;
    return refreshedRows == std::min(successfulSourcePaths.size(), successfulTargetPaths.size());
}
#endif

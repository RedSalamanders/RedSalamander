#pragma once

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <span>
#include <string>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "BatchRenameEngine.h"
#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"
#include "SettingsStore.h"

struct BatchRenamePaneContext
{
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::filesystem::path rootPluginPath;
    std::vector<std::filesystem::path> initialPaths;
    // Invoked on the UI thread after an execution attempt that renamed at least one
    // row (including failed or canceled batches with partial success). sourcePaths and
    // targetPaths are parallel lists in execution order (deepest first). Each target is
    // the row's path as of its own rename: a later parent-directory rename in the same
    // batch is intentionally NOT folded in. Consumers must replay the pairs
    // sequentially to compute final paths (see FolderWindow::RefreshPanesAfterBatchRename).
    std::function<void(std::span<const std::filesystem::path> sourcePaths, std::span<const std::filesystem::path> targetPaths)> onSuccessfulRename;
    std::function<bool(const std::filesystem::path& path)> onRevealPath;
};

[[nodiscard]] bool ShowBatchRenameWindow(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, BatchRenamePaneContext context) noexcept;

void UpdateBatchRenameWindowsTheme(const AppTheme& theme) noexcept;

[[nodiscard]] HWND GetBatchRenameWindowHandle() noexcept;
[[nodiscard]] bool IsBatchRenameWindowHandle(HWND hwnd) noexcept;

#ifdef ENABLE_TESTS
enum class BatchRenameDebugRuleField
{
    NameTemplate,
    SearchFor,
    ReplaceWith,
};

enum class BatchRenameDebugPreviewCopyKind
{
    OriginalName,
    NewName,
    SourcePath,
    PreviewRows,
};

struct BatchRenameDebugSnapshot
{
    bool usesDxUiHost              = false;
    bool rootNavigationVisible     = false;
    bool rootNavigationUsesNavigationView = false;
    bool ruleControlsVisible       = false;
    bool ruleHelperButtonsVisible  = false;
    bool rulesModeSelected         = false;
    bool manualModeSelected        = false;
    bool manualControlsVisible     = false;
    bool renameButtonEnabled       = false;
    bool hideUnchangedRows         = false;
    bool previewRebuildPending     = false;
    bool hasExecutionReport        = false;
    size_t visibleChildWindowCount = 0u;
    size_t previewRowCount         = 0u;
    size_t previewIconCellCount    = 0u;
    size_t changedRowCount         = 0u;
    size_t errorRowCount           = 0u;
    size_t warningRowCount         = 0u;
    size_t lastExecutionTotalRows  = 0u;
    size_t lastExecutionCompletedRows = 0u;
    size_t lastExecutionSkippedRows   = 0u;
    size_t lastExecutionFailedRows    = 0u;
    size_t lastExecutionUndoRowCount  = 0u;
    HRESULT lastExecutionFirstFailure = S_OK;
    bool lastExecutionCanceled        = false;
    std::wstring lastExecutionFirstFailureText;
    std::vector<std::wstring> previewColumnIds;
    std::vector<std::wstring> originalNames;
    std::vector<std::wstring> newNames;
    std::vector<std::wstring> newNameStatusIconTexts;
    std::vector<std::wstring> newNameTooltips;
    std::vector<std::wstring> sizeTexts;
    std::vector<std::wstring> dateTexts;
    std::vector<std::wstring> timeTexts;
    std::vector<std::wstring> fullPaths;
    std::vector<int> originalIconIndices;
    std::vector<std::wstring> focusableAccessibleNames;
    std::wstring rootText;
    std::wstring rootNavigationPathText;
    std::wstring statusText;
    std::wstring scopeMaskText;
    std::wstring nameTemplateText;
    std::wstring searchForText;
    std::wstring replaceWithText;
    std::wstring manualText;
    bool includeSubdirectories = false;
    bool includeFiles          = true;
    bool includeFolders        = false;
    bool regexEnabled      = false;
    bool caseSensitive     = false;
    bool wholeWords        = false;
    bool replaceOnce       = false;
    bool excludeExtension  = false;
    std::wstring fileNameCaseText;
    std::wstring extensionCaseText;
};

struct BatchRenameDebugCollectionResult
{
    HRESULT hr = S_OK;
    std::wstring detail;
    std::vector<std::wstring> originalNames;
    std::vector<std::wstring> fullPaths;
    std::vector<bool> isDirectories;
    std::vector<bool> metadataUnknowns;
    std::vector<uint64_t> sizeBytes;
};

[[nodiscard]] size_t DebugGetBatchRenameWindowCount() noexcept;
[[nodiscard]] bool DebugGetBatchRenameWindowSnapshot(BatchRenameDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowRules(const BatchRename::Rules& rules) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowRuleControls(const BatchRename::Rules& rules) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowScope(std::wstring_view mask, bool includeSubdirectories, bool includeFiles, bool includeFolders) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowPreviewSort(std::wstring_view columnId, bool descending) noexcept;
[[nodiscard]] bool DebugReorderBatchRenameWindowPreviewColumn(std::wstring_view columnId, size_t targetDisplayIndex) noexcept;
[[nodiscard]] bool DebugSwitchBatchRenameWindowMode(BatchRename::Mode mode) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowManualText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugClickBatchRenameWindowManualPaste() noexcept;
[[nodiscard]] bool DebugClickBatchRenameWindowManualSortLikePreview() noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowHideUnchanged(bool hideUnchanged) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowRuleFieldSelection(BatchRenameDebugRuleField field, size_t selectionStart, size_t selectionEnd) noexcept;
[[nodiscard]] bool DebugSetBatchRenameWindowRuleFieldText(BatchRenameDebugRuleField field, std::wstring_view text) noexcept;
[[nodiscard]] bool DebugInsertBatchRenameWindowHelperCommand(BatchRenameDebugRuleField field, int commandId) noexcept;
[[nodiscard]] bool DebugCopyBatchRenameWindowPreview(BatchRenameDebugPreviewCopyKind kind, size_t rowIndex) noexcept;
[[nodiscard]] bool DebugRevealBatchRenameWindowPreview(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugActivateBatchRenameWindowPreview(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugCopyBatchRenameWindowExecutionReport() noexcept;
[[nodiscard]] bool DebugCopyBatchRenameWindowUndoPlan() noexcept;
[[nodiscard]] bool DebugFlushBatchRenameWindowPendingPreview() noexcept;
[[nodiscard]] HRESULT DebugExecuteBatchRenameWindow() noexcept;
[[nodiscard]] HRESULT DebugStartBatchRenameWindowExecution() noexcept;
[[nodiscard]] HRESULT DebugWaitBatchRenameWindowExecutionIdle() noexcept;
[[nodiscard]] bool DebugInjectStaleBatchRenameWindowCollectionPayload(std::filesystem::path sourcePath) noexcept;
[[nodiscard]] bool DebugInjectStaleBatchRenameWindowExecutionPayload(std::filesystem::path sourcePath, std::filesystem::path targetPath) noexcept;
void DebugSetBatchRenameWindowDestinationProbeFailurePath(std::filesystem::path destinationPath, unsigned long win32Error);
void DebugClearBatchRenameWindowDestinationProbeFailurePath() noexcept;
[[nodiscard]] bool DebugCollectBatchRenameTargetsForTests(BatchRenamePaneContext context,
                                                          std::wstring_view mask,
                                                          bool includeSubdirectories,
                                                          bool includeFiles,
                                                          bool includeFolders,
                                                          BatchRenameDebugCollectionResult& out);
[[nodiscard]] bool DebugRefreshBatchRenameTargetsAfterExecutionForTests(
    const FileSystemPathIdentity& pathIdentity,
    std::vector<BatchRename::Target>& targets,
    std::span<const std::filesystem::path> successfulSourcePaths,
    std::span<const std::filesystem::path> successfulTargetPaths,
    const std::filesystem::path& root,
    size_t& refreshedRows,
    uint64_t& identityComparisons) noexcept;
#endif

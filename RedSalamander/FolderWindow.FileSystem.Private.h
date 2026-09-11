#pragma once

#include <atomic>
#include <filesystem>
#include <memory>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "AppTheme.h"
#include "BatchRenameExecutionEngine.h"
#include "FileSystemPluginManager.h"
#include "FolderWindowInternal.h"
#include "SettingsStore.h"

namespace FolderWindowFileSystemInternal
{
struct ChangeCaseTaskReceipt final
{
    ChangeCaseTaskReceipt()                                       = default;
    ChangeCaseTaskReceipt(const ChangeCaseTaskReceipt&)            = delete;
    ChangeCaseTaskReceipt& operator=(const ChangeCaseTaskReceipt&) = delete;
    ChangeCaseTaskReceipt(ChangeCaseTaskReceipt&&)                 = delete;
    ChangeCaseTaskReceipt& operator=(ChangeCaseTaskReceipt&&)      = delete;

    std::atomic<uint64_t> taskId{0u};
};

struct EditNewPromptResult final
{
    std::wstring fileName;
    std::wstring editorActionId;
};

struct ChangeCaseTaskPayload final
{
    FolderWindow::InformationalTaskUpdate update{};
    std::shared_ptr<ChangeCaseTaskReceipt> receipt;
};

template <typename CreateOrUpdate>
[[nodiscard]] uint64_t ResolveChangeCaseTaskUpdate(ChangeCaseTaskPayload& payload, CreateOrUpdate&& createOrUpdate) noexcept
{
    if (payload.receipt)
    {
        const uint64_t existingTaskId = payload.receipt->taskId.load(std::memory_order_acquire);
        if (existingTaskId != 0u)
        {
            payload.update.taskId = existingTaskId;
        }
    }

    const uint64_t resolvedTaskId = std::forward<CreateOrUpdate>(createOrUpdate)(payload.update);
    if (! payload.receipt || resolvedTaskId == 0u)
    {
        return resolvedTaskId;
    }

    uint64_t expectedTaskId = 0u;
    if (payload.receipt->taskId.compare_exchange_strong(
            expectedTaskId, resolvedTaskId, std::memory_order_acq_rel, std::memory_order_acquire))
    {
        return resolvedTaskId;
    }
    return expectedTaskId;
}

struct ChangeCaseCompletedPayload final
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    HRESULT hr              = S_OK;
    wil::com_ptr<IFileSystem> fileSystem;
    std::vector<BatchRenameExecutionOp> operations;
    std::filesystem::path focusFolder;
    std::wstring focusDisplayName;
};

[[nodiscard]] bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept;
[[nodiscard]] HWND GetOwnerWindowOrSelf(HWND window) noexcept;
[[nodiscard]] size_t CountVisibleChildWindowsLocal(HWND hwnd) noexcept;
#ifdef ENABLE_TESTS
[[nodiscard]] std::wstring GetWindowClassNameLocal(HWND hwnd);
[[nodiscard]] size_t CountVisibleNativeChildControlWindowsLocal(HWND hwnd) noexcept;
[[nodiscard]] HRESULT DebugResolveInitialCreateDirectoryNameForTests(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                     const std::filesystem::path& folder,
                                                                     std::wstring_view defaultName,
                                                                     std::wstring_view pluginId,
                                                                     std::wstring& nameOut) noexcept;
#endif
[[nodiscard]] int ScalePanePromptForDpi(UINT dpi, int dip) noexcept;
[[nodiscard]] HWND GetClipboardOwnerWindow(HWND window) noexcept;
[[nodiscard]] std::wstring GetUniversalPathOrOriginal(std::wstring_view nativePath) noexcept;
void AddToFolderHistory(std::vector<std::filesystem::path>& history, size_t maxItems, const std::filesystem::path& entry);
[[nodiscard]] FolderView::NameFilterState GetFolderHistoryFilterState(const Common::Settings::FoldersSettings* folders,
                                                                      const std::filesystem::path& displayPath);
void SetFolderHistoryFilterState(Common::Settings::FoldersSettings& folders,
                                 const std::filesystem::path& displayPath,
                                 const FolderView::NameFilterState& filter);
void PruneFolderHistoryFilters(Common::Settings::FoldersSettings& folders,
                               const std::vector<std::filesystem::path>& history,
                               size_t maxItems);
[[nodiscard]] bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept;
[[nodiscard]] std::filesystem::path GetDefaultFileSystemRoot() noexcept;
[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindPluginById(
    const std::vector<FileSystemPluginManager::PluginEntry>& plugins, std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<std::filesystem::path> TryResolveInstanceContextToWindowsPath(std::wstring_view instanceContext) noexcept;
[[nodiscard]] std::optional<UINT> ResolveEditNewValidationMessageId(std::wstring_view trimmed,
                                                                    const std::filesystem::path& targetFolder) noexcept;
[[nodiscard]] std::wstring GetComputerNameTextForFileActions() noexcept;
[[nodiscard]] std::wstring TryGetFileSystemPluginDisplayName(const std::vector<FileSystemPluginManager::PluginEntry>& plugins,
                                                             std::wstring_view pluginId,
                                                             std::wstring_view pluginShortId) noexcept;
[[nodiscard]] std::optional<std::wstring> PromptForCreateDirectoryName(HWND ownerWindow,
                                                                      std::wstring_view createInPath,
                                                                      std::wstring_view initialName,
                                                                      const AppTheme& theme);
[[nodiscard]] std::optional<EditNewPromptResult> PromptForEditNewFile(
    HWND ownerWindow,
    const std::filesystem::path& targetFolder,
    std::wstring_view displayPath,
    const Common::Settings::EditorFileActionsSettings* editorSettings,
    std::wstring_view computerName,
    const AppTheme& theme,
    std::wstring_view initialFileName = {},
    std::wstring_view captionText     = {},
    bool showEditorControls           = true);
[[nodiscard]] std::optional<ChangeCase::Options> PromptForChangeCase(HWND ownerWindow, const AppTheme& theme, bool allowSubdirs) noexcept;
[[nodiscard]] std::optional<std::wstring> PromptForSelectionMask(
    HWND ownerWindow, const std::vector<std::wstring>& history, const AppTheme& theme, UINT captionId, UINT labelId);
[[nodiscard]] std::optional<FolderView::NameFilterState> PromptForPaneFilter(HWND ownerWindow,
                                                                            const std::vector<std::wstring>& history,
                                                                            const AppTheme& theme,
                                                                            const FolderView::NameFilterState& initial);
} // namespace FolderWindowFileSystemInternal

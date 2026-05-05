#pragma once
#include <array>
#include <condition_variable>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "AppTheme.h"
#include "DxUi/DxUi.h"
#include "FolderView.h"
#include "Framework.h"
#include "FunctionBar.h"
#include "NavigationView.h"
#include "PlugInterfaces/Viewer.h"

namespace Common::Settings
{
struct MakeFileListSettings;
struct Settings;
}

LRESULT CALLBACK FolderWindowDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

#ifdef ENABLE_TESTS
void DebugSetMakeFileListAutomation(const Common::Settings::MakeFileListSettings& options) noexcept;
void DebugClearMakeFileListAutomation() noexcept;

struct ChangeAttributesOptionsPromptDebugSnapshot
{
    bool usesDxUiHost                     = false;
    size_t visibleChildWindowCount        = 0u;
    size_t visibleNativeChildControlCount = 0u;
    std::wstring dialogClassName;
    uint8_t readOnly                      = 0u;
    uint8_t hidden                        = 0u;
    uint8_t system                        = 0u;
    uint8_t archive                       = 0u;
    bool dateTimeSectionVisible           = false;
    bool modifiedTimeVisible              = false;
    bool createdTimeVisible               = false;
    bool accessedTimeVisible              = false;
    bool includeSubdirectoriesVisible     = false;
    bool includeSubdirectoriesEnabled     = false;
    bool includeSubdirectoriesChecked     = false;
    bool removeAlternateDataStreams       = false;
};

[[nodiscard]] HWND GetChangeAttributesOptionsPromptHandle() noexcept;
[[nodiscard]] bool DebugGetChangeAttributesOptionsPromptSnapshot(ChangeAttributesOptionsPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetChangeAttributesOptionsPromptState(
    uint8_t readOnly, uint8_t hidden, uint8_t system, uint8_t archive, bool removeAlternateDataStreams) noexcept;
[[nodiscard]] bool DebugCycleChangeAttributesOptionsPromptArchive() noexcept;
[[nodiscard]] bool DebugConfirmChangeAttributesOptionsPrompt() noexcept;
[[nodiscard]] bool DebugCancelChangeAttributesOptionsPrompt() noexcept;

struct MakeFileListOptionsPromptDebugSnapshot
{
    bool usesDxUiHost                     = false;
    size_t visibleChildWindowCount        = 0u;
    size_t visibleNativeChildControlCount = 0u;
    std::wstring dialogClassName;
    uint8_t sourceMode                    = 0u;
    bool recursive                        = false;
    uint8_t format                        = 0u;
    uint8_t outputTarget                  = 0u;
    std::wstring textMacro;
    std::wstring outputFileText;
    bool includeName              = false;
    bool includeFullPath          = false;
    bool includeSize              = false;
    bool includeModified          = false;
    bool includeAttributes        = false;
    bool includeDirectories       = false;
    bool outputFileFieldEnabled   = false;
};

[[nodiscard]] HWND GetMakeFileListOptionsPromptHandle() noexcept;
[[nodiscard]] bool DebugGetMakeFileListOptionsPromptSnapshot(MakeFileListOptionsPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetMakeFileListOptionsPromptState(uint8_t sourceMode,
                                                          bool recursive,
                                                          uint8_t format,
                                                          uint8_t outputTarget,
                                                          std::wstring_view textMacro,
                                                          std::wstring_view outputFile,
                                                          bool includeName,
                                                          bool includeFullPath,
                                                          bool includeSize,
                                                          bool includeModified,
                                                          bool includeAttributes,
                                                          bool includeDirectories) noexcept;
[[nodiscard]] bool DebugConfirmMakeFileListOptionsPrompt() noexcept;
[[nodiscard]] bool DebugCancelMakeFileListOptionsPrompt() noexcept;

struct ArchivePackPromptDebugSnapshot
{
    bool usesDxUiHost                     = false;
    size_t visibleChildWindowCount        = 0u;
    size_t visibleNativeChildControlCount = 0u;
    std::wstring dialogClassName;
    std::wstring archivePathText;
    std::wstring packerDisplayName;
    std::wstring packerExtension;
    size_t packerCount          = 0u;
    size_t selectedPackerIndex  = 0u;
    bool deleteAfterPacking     = false;
    bool commandButtonsFitInClient = false;
};

[[nodiscard]] HWND GetArchivePackPromptHandle() noexcept;
[[nodiscard]] bool DebugGetArchivePackPromptSnapshot(ArchivePackPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetArchivePackPromptPackerIndex(size_t index) noexcept;
[[nodiscard]] bool DebugSetArchivePackPromptArchivePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugSetArchivePackPromptDeleteAfter(bool deleteAfterPacking) noexcept;
[[nodiscard]] bool DebugConfirmArchivePackPrompt() noexcept;
[[nodiscard]] bool DebugCancelArchivePackPrompt() noexcept;

struct ArchiveUnpackPromptDebugSnapshot
{
    bool usesDxUiHost                     = false;
    size_t visibleChildWindowCount        = 0u;
    size_t visibleNativeChildControlCount = 0u;
    std::wstring dialogClassName;
    std::wstring destinationPathText;
    std::wstring unpackerDisplayName;
    std::wstring unpackerExtension;
    size_t unpackerCount          = 0u;
    size_t selectedUnpackerIndex  = 0u;
    bool deleteAfterUnpacking     = false;
    std::wstring maskText;
    bool maskHelpVisible          = false;
    bool commandButtonsFitInClient = false;
};

[[nodiscard]] HWND GetArchiveUnpackPromptHandle() noexcept;
[[nodiscard]] bool DebugGetArchiveUnpackPromptSnapshot(ArchiveUnpackPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetArchiveUnpackPromptDestinationPath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugSetArchiveUnpackPromptMask(std::wstring_view mask) noexcept;
[[nodiscard]] bool DebugSetArchiveUnpackPromptDeleteAfter(bool deleteAfterUnpacking) noexcept;
[[nodiscard]] bool DebugConfirmArchiveUnpackPrompt() noexcept;
[[nodiscard]] bool DebugCancelArchiveUnpackPrompt() noexcept;
#endif

class ShortcutManager;

#ifdef ENABLE_TESTS
struct ItemPropertiesWindowDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    size_t sectionCount            = 0u;
    size_t fieldCount              = 0u;
    size_t streamCount             = 0u;
    size_t removableStreamCount    = 0u;
    size_t viewableStreamCount     = 0u;
    size_t bodyFirstVisibleLine    = 0u;
    size_t bodyVisibleLineCount    = 0u;
    size_t bodyTotalLineCount      = 0u;
    bool bodyCanScrollVertically   = false;
    bool loading                   = false;
    bool loadFailed                = false;
    float layoutOverflowRightDip   = 0.0f;
    uint64_t renderCount           = 0u;
    uint64_t resizeCount           = 0u;
    uint64_t resizeFailureCount    = 0u;
    std::wstring contentText;
};

[[nodiscard]] HWND GetItemPropertiesWindowHandle() noexcept;
[[nodiscard]] bool DebugGetItemPropertiesWindowSnapshot(ItemPropertiesWindowDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugScrollItemPropertiesWindowByWheelDetents(int detents) noexcept;
[[nodiscard]] HRESULT DebugRemoveItemPropertiesStream(std::wstring_view streamName) noexcept;
[[nodiscard]] HRESULT DebugOpenItemPropertiesStream(std::wstring_view streamName) noexcept;
[[nodiscard]] std::wstring DebugBuildItemPropertiesContentTextFromJson(std::string_view jsonUtf8) noexcept;
void DebugSetNextItemPropertiesLoadDelayMs(uint32_t delayMs) noexcept;

struct FolderViewPaneFilterPromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    bool enabled                   = false;
    bool helpExpanded              = false;
    bool commandButtonsFitInClient = false;
    float clientBottomDip          = 0.0f;
    float okButtonBottomDip        = 0.0f;
    float cancelButtonBottomDip    = 0.0f;
    std::wstring text;
};

[[nodiscard]] HWND GetFolderViewPaneFilterPromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewPaneFilterPromptSnapshot(FolderViewPaneFilterPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewPaneFilterPromptEnabled(bool enabled) noexcept;
[[nodiscard]] bool DebugSetFolderViewPaneFilterPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetFolderViewPaneFilterPromptTextAndNotify(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetFolderViewPaneFilterPromptHelpExpanded(bool expanded) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewPaneFilterPrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewPaneFilterPrompt() noexcept;

struct FolderViewSelectionMaskPromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    std::wstring title;
    std::wstring text;
};

[[nodiscard]] HWND GetFolderViewSelectionMaskPromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewSelectionMaskPromptSnapshot(FolderViewSelectionMaskPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewSelectionMaskPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewSelectionMaskPrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewSelectionMaskPrompt() noexcept;

struct FolderViewCreateDirectoryPromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    std::wstring createInPath;
    std::wstring text;
    std::wstring validationText;
    bool nameFieldFocused = false;
    size_t selectionStart = 0u;
    size_t selectionEnd   = 0u;
};

[[nodiscard]] HWND GetFolderViewCreateDirectoryPromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewCreateDirectoryPromptSnapshot(FolderViewCreateDirectoryPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewCreateDirectoryPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewCreateDirectoryPrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewCreateDirectoryPrompt() noexcept;

struct FolderViewEditNewPromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    std::wstring createInPath;
    std::wstring fileNameText;
    std::wstring validationText;
    bool nameFieldFocused = false;
    size_t selectionStart = 0u;
    size_t selectionEnd   = 0u;
    bool editorComboEnabled = false;
    std::wstring selectedEditorActionId;
    std::vector<std::wstring> editorActionIds;
    std::vector<std::wstring> editorDisplayNames;
};

[[nodiscard]] HWND GetFolderViewEditNewPromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewEditNewPromptSnapshot(FolderViewEditNewPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewEditNewPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSelectFolderViewEditNewPromptEditor(std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewEditNewPrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewEditNewPrompt() noexcept;

struct FolderViewChangeCasePromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    bool includeSubdirsEnabled     = false;
    bool includeSubdirsChecked     = false;
    size_t styleIndex              = 0u;
    size_t targetIndex             = 0u;
    std::wstring exampleText;
};

[[nodiscard]] HWND GetFolderViewChangeCasePromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewChangeCasePromptSnapshot(FolderViewChangeCasePromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewChangeCasePromptSelections(size_t styleIndex, size_t targetIndex, bool includeSubdirs) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewChangeCasePrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewChangeCasePrompt() noexcept;
#endif

class FolderWindow
{
public:
    FolderWindow();
    ~FolderWindow();

    // Disable copy and move
    FolderWindow(const FolderWindow&)            = delete;
    FolderWindow& operator=(const FolderWindow&) = delete;
    FolderWindow(FolderWindow&&)                 = delete;
    FolderWindow& operator=(FolderWindow&&)      = delete;

    // Window management
    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy();
    [[maybe_unused]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

    enum class Pane : uint8_t
    {
        Left,
        Right,
    };

    // Navigation
    void SetFolderPath(const std::filesystem::path& path);
    std::optional<std::filesystem::path> GetCurrentPath() const;
    std::optional<std::filesystem::path> GetCurrentPluginPath() const;
    std::vector<std::filesystem::path> GetFolderHistory() const;
    void SetFolderHistory(const std::vector<std::filesystem::path>& history);
    uint32_t GetFolderHistoryMax() const noexcept;
    void SetFolderHistoryMax(uint32_t maxItems);

    void SetActivePane(Pane pane) noexcept;
    [[maybe_unused]] Pane GetActivePane() const noexcept
    {
        return _activePane;
    }
    [[maybe_unused]] Pane GetFocusedPane() const noexcept;
    [[nodiscard]] HWND GetFocusedFolderViewHwnd() const noexcept;
    [[nodiscard]] HWND GetFolderViewHwnd(Pane pane) const noexcept;
    [[nodiscard]] bool IsFocusInNavigationView() const noexcept;
    [[nodiscard]] bool TryRestoreActivePaneFolderViewFocus() noexcept;
    void RequestRestoreFolderViewFocus(HWND folderView) noexcept;

    HRESULT ExecuteInActivePane(const std::filesystem::path& folderPath,
                                std::wstring_view focusItemDisplayName,
                                unsigned int folderViewCommandId,
                                bool activateWindow) noexcept;
    HRESULT OpenViewerWithPlugin(std::wstring_view pluginId,
                                 const ViewerOpenContext& context,
                                 std::wstring_view openedBy = {},
                                 Pane pane = Pane::Left) noexcept;

    void SetFolderPath(Pane pane, const std::filesystem::path& path);
    std::optional<std::filesystem::path> GetCurrentPath(Pane pane) const;
    std::optional<std::filesystem::path> GetCurrentPluginPath(Pane pane) const;
    [[nodiscard]] std::optional<std::filesystem::path> GetFocusedItemPath(Pane pane) const;
    std::vector<std::filesystem::path> GetFolderHistory(Pane pane) const;
    void SetFolderHistory(Pane pane, const std::vector<std::filesystem::path>& history);

    void SetDisplayMode(Pane pane, FolderView::DisplayMode mode);
    FolderView::DisplayMode GetDisplayMode(Pane pane) const noexcept;

    void SetSort(Pane pane, FolderView::SortBy sortBy, FolderView::SortDirection direction);
    void CycleSortBy(Pane pane, FolderView::SortBy sortBy);
    FolderView::SortBy GetSortBy(Pane pane) const noexcept;
    FolderView::SortDirection GetSortDirection(Pane pane) const noexcept;

    void SetStatusBarVisible(Pane pane, bool visible);
    bool GetStatusBarVisible(Pane pane) const noexcept;
    void SetFileExtensionsVisible(Pane pane, bool visible);
    bool GetFileExtensionsVisible(Pane pane) const noexcept;
    void SetThumbnailsVisible(Pane pane, bool visible);
    bool GetThumbnailsVisible(Pane pane) const noexcept;
    void SetNavigationBarVisible(Pane pane, bool visible);
    bool GetNavigationBarVisible(Pane pane) const noexcept;
    void SetFilterBarVisible(Pane pane, bool visible);
    bool GetFilterBarVisible(Pane pane) const noexcept;
    void SetNameFilterState(Pane pane, const FolderView::NameFilterState& state, bool refresh = true);
    void TogglePreviewPane(Pane sourcePane);
    [[nodiscard]] bool IsPreviewPaneOpenForSource(Pane sourcePane) const noexcept;

    void SetShowHiddenFiles(bool show);
    [[nodiscard]] bool GetShowHiddenFiles() const noexcept;
    void SetShowSystemFiles(bool show);
    [[nodiscard]] bool GetShowSystemFiles() const noexcept;

    void CommandRename(Pane pane);
    void CommandView(Pane pane);
    void CommandAlternateView(Pane pane);
    void CommandViewWith(Pane pane, std::wstring_view actionId);
    void CommandEdit(Pane pane);
    void CommandAlternateEdit(Pane pane);
    void CommandEditWith(Pane pane, std::wstring_view actionId);
    void CommandEditNew(Pane pane);
    struct UserMenuItem final
    {
        std::wstring id;
        std::wstring displayName;
        bool enabled = false;
        HRESULT availabilityHr = S_OK;
    };
    [[nodiscard]] std::vector<UserMenuItem> CollectUserMenuItems(Pane pane) const;
    void CommandUserMenu(Pane pane, std::wstring_view actionId);
    void CommandNewFromShellTemplate(Pane pane, std::wstring_view templateId);
    void CommandContextMenuCurrentDirectory(Pane pane);
    void CommandOpenSecurity(Pane pane);
    void CommandGoToShortcutOrLinkTarget(Pane pane);

    enum class AttributeChangeState : uint8_t
    {
        LeaveUnchanged,
        Set,
        Clear,
    };

    struct ChangeAttributesOptions final
    {
        struct TimestampOption final
        {
            bool enabled  = false;
            int64_t value = 0;
        };

        AttributeChangeState readOnly = AttributeChangeState::LeaveUnchanged;
        AttributeChangeState hidden   = AttributeChangeState::LeaveUnchanged;
        AttributeChangeState system   = AttributeChangeState::LeaveUnchanged;
        AttributeChangeState archive  = AttributeChangeState::LeaveUnchanged;
        TimestampOption modifiedTime;
        TimestampOption createdTime;
        TimestampOption accessedTime;
        bool includeSubdirectories    = false;
        bool removeAlternateDataStreams = false;
    };

    struct ChangeAttributesReport final
    {
        uint64_t itemsProcessed    = 0u;
        uint64_t attributesChanged = 0u;
        uint64_t timesChanged      = 0u;
        uint64_t streamsRemoved    = 0u;
        uint64_t failures          = 0u;
        HRESULT firstFailure       = S_OK;
        uint64_t progressTaskId    = 0u;
        std::wstring summary;
    };

    void CommandChangeAttributes(Pane pane);
    void CommandViewSpace(Pane pane);
    void CommandDelete(Pane pane);
    void CommandPermanentDelete(Pane pane);
    void CommandCopyToOtherPane(Pane sourcePane);
    void CommandMoveToOtherPane(Pane sourcePane);
    void CommandToggleFileOperationsIssuesPane();
    bool IsFileOperationsIssuesPaneVisible() noexcept;
    void CommandCreateDirectory(Pane pane);
    void CommandChangeDirectory(Pane pane);
    [[nodiscard]] bool TryHandleNavigationEditClipboardCommand(UINT commandId) noexcept;
    void CommandFocusAddressBar(Pane pane);
    void CommandOpenDriveMenu(Pane pane);
    void CommandShowFolderHistory(Pane pane);
    void CommandQuickSearch(Pane pane);
    void CommandBringCurrentDirToCommandLine(Pane pane);
    void CommandBringFilenameToCommandLine(Pane pane);
    void CommandMakeFileList(Pane pane);
    void CommandPack(Pane pane);
    void CommandUnpack(Pane pane);
    void CommandListOpenedFiles(Pane pane);
    void CommandSharedDirectories(Pane pane);
    void CommandFilter(Pane pane);
    void CommandGoRootDirectory(Pane pane);
    void CommandSetPathFromOtherPane(Pane pane);
    void CommandHistoryBack(Pane pane);
    void CommandHistoryForward(Pane pane);
    void ResyncNavigationShellFromFolderView(Pane pane) noexcept;
    [[nodiscard]] bool CanHistoryBack(Pane pane) const noexcept;
    [[nodiscard]] bool CanHistoryForward(Pane pane) const noexcept;
    void CommandRefresh(Pane pane);

    enum class ShellNewTemplateKind : uint8_t
    {
        NullFile,
        Data,
        FileName,
    };

    struct ShellNewTemplateDefinition final
    {
        std::wstring id;
        std::wstring displayName;
        std::wstring extension;
        std::wstring defaultFileName;
        ShellNewTemplateKind kind = ShellNewTemplateKind::NullFile;
        std::vector<std::byte> data;
        std::filesystem::path templateFilePath;
    };

    struct ShellNewTemplateMenuItem final
    {
        std::wstring id;
        std::wstring displayName;
    };

    [[nodiscard]] std::vector<ShellNewTemplateMenuItem> CollectShellNewTemplateMenuItems(Pane pane) const;
    void CommandSelectionSelectDialog(Pane pane);
    void CommandSelectionUnselectDialog(Pane pane);
    void CommandSelectionInvert(Pane pane);
    void CommandSelectionSave(Pane pane);
    void CommandSelectionRestore(Pane pane);
    void CommandSelectionSelectSameName(Pane pane);
    void CommandSelectionUnselectSameName(Pane pane);
    void CommandSelectionSelectSameExtension(Pane pane);
    void CommandSelectionUnselectSameExtension(Pane pane);
    void CommandSelectionHideSelectedNames(Pane pane);
    void CommandSelectionHideUnselectedNames(Pane pane);
    void CommandSelectionShowHiddenNames(Pane pane);
    [[nodiscard]] bool CanShowHiddenNames(Pane pane) const noexcept;
    void CommandSelectionGoToPreviousSelectedName(Pane pane);
    void CommandSelectionGoToNextSelectedName(Pane pane);
    void CommandChangeCase(Pane pane);
    void CommandOpenCommandShell(Pane pane);
    void CommandCopyPathAndNameAsText(Pane pane);
    void CommandCopyNameAsText(Pane pane);
    void CommandCopyPathAsText(Pane pane);
    void CommandCopyUncPathAndNameAsText(Pane pane);
    void PrepareForNetworkDriveDisconnect(Pane pane);
    void SwapPanes();

    bool ConfirmCancelAllFileOperations(HWND ownerWindow) noexcept;
    void CloseAllViewers() noexcept;

    using ShowSortMenuCallback = std::function<void(Pane pane, POINT screenPoint)>;
    void SetShowSortMenuCallback(ShowSortMenuCallback callback);

    float GetSplitRatio() const noexcept
    {
        return _splitRatio;
    }
    void SetSplitRatio(float ratio);
    void BeginViewWidthAdjust() noexcept;
    void CommitViewWidthAdjust() noexcept;
    void CancelViewWidthAdjust() noexcept;
    [[nodiscard]] bool IsViewWidthAdjustActive() const noexcept
    {
        return _viewWidthAdjustActive;
    }
    [[nodiscard]] bool HandleViewWidthAdjustKey(uint32_t vk) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugIsViewWidthAdjustActive() const noexcept
    {
        return _viewWidthAdjustActive;
    }
#endif
    void ToggleZoomPanel(Pane pane);
    [[nodiscard]] std::optional<Pane> GetZoomedPane() const noexcept
    {
        return _zoomedPane;
    }
    [[nodiscard]] std::optional<float> GetZoomRestoreSplitRatio() const noexcept
    {
        return _zoomRestoreSplitRatio;
    }
    void SetZoomState(std::optional<Pane> zoomedPane, std::optional<float> restoreSplitRatio);

    void SetSettings(Common::Settings::Settings* settings) noexcept;
    void SetShortcutManager(const ShortcutManager* shortcuts) noexcept;
    void SetFunctionBarModifiers(uint32_t modifiers) noexcept;
    void SetFunctionBarPressedKey(std::optional<uint32_t> vk) noexcept;
    void SetFunctionBarVisible(bool visible) noexcept;
    [[nodiscard]] bool GetFunctionBarVisible() const noexcept
    {
        return _functionBarVisible;
    }

    // Extension points (used by scoped folder windows like Compare).
    using PanePathChangedCallback = std::function<void(Pane pane, const std::optional<std::filesystem::path>& pluginPath)>;
    void SetPanePathChangedCallback(PanePathChangedCallback callback);
    void SetPaneEnumerationCompletedCallback(Pane pane, FolderView::EnumerationCompletedCallback callback);
    void SetPaneDetailsTextProvider(Pane pane, FolderView::DetailsTextProvider provider);
    void SetPaneMetadataTextProvider(Pane pane, FolderView::MetadataTextProvider provider);
    void SetPaneEmptyStateMessage(Pane pane, std::wstring message);
    void SetPaneBackgroundWatermark(Pane pane, std::wstring message, bool animated);
    void RefreshPaneDetailsText(Pane pane);
    void SetPaneSelectionByDisplayNamePredicate(Pane pane,
                                                const std::function<bool(std::wstring_view)>& shouldSelect,
                                                bool clearExistingSelection = true) noexcept;
    void ClearPaneSelectionByDisplayNamePredicate(Pane pane, const std::function<bool(std::wstring_view)>& shouldUnselect) noexcept;

    [[nodiscard]] bool HasSavedSelection() const noexcept;

    struct FileOperationCompletedEvent
    {
        FileSystemOperation operation = static_cast<FileSystemOperation>(0);
        Pane sourcePane               = Pane::Left;
        std::optional<Pane> destinationPane;
        std::vector<std::filesystem::path> sourcePaths;
        std::optional<std::filesystem::path> destinationFolder;
        HRESULT hr = S_OK;
    };
    using FileOperationCompletedCallback = std::function<void(const FileOperationCompletedEvent& e)>;
    void SetFileOperationCompletedCallback(FileOperationCompletedCallback callback);

    struct InformationalTaskUpdate final
    {
        static constexpr size_t kMaxContentInFlightFiles = 8u;

        enum class Kind : uint8_t
        {
            CompareDirectories,
            ChangeCase,
            ChangeAttributes,
        };

        Kind kind       = Kind::CompareDirectories;
        uint64_t taskId = 0;
        std::wstring title;

        // Compare Directories payload (Kind::CompareDirectories)
        std::filesystem::path leftRoot;
        std::filesystem::path rightRoot;

        bool scanActive = false;
        std::filesystem::path scanCurrentRelative;
        uint64_t scanFolderCount         = 0;
        uint64_t scanEntryCount          = 0;
        uint64_t scanCandidateFileCount  = 0;
        uint64_t scanCandidateTotalBytes = 0;
        std::optional<uint64_t> scanElapsedSeconds;

        bool contentActive = false;
        std::filesystem::path contentCurrentRelative;
        uint64_t contentCurrentTotalBytes     = 0;
        uint64_t contentCurrentCompletedBytes = 0;
        uint64_t contentTotalBytes            = 0;
        uint64_t contentCompletedBytes        = 0;
        uint64_t contentPendingCount          = 0;
        uint64_t contentCompletedCount        = 0;
        std::optional<uint64_t> contentEtaSeconds;

        struct ContentInFlightFile final
        {
            std::filesystem::path relativePath;
            uint64_t totalBytes      = 0;
            uint64_t completedBytes  = 0;
            ULONGLONG lastUpdateTick = 0;
        };

        std::array<ContentInFlightFile, kMaxContentInFlightFiles> contentInFlight{};
        size_t contentInFlightCount = 0;

        // Change Case payload (Kind::ChangeCase)
        bool changeCaseEnumerating = false;
        bool changeCaseRenaming    = false;
        std::filesystem::path changeCaseCurrentPath;
        uint64_t changeCaseScannedFolders   = 0;
        uint64_t changeCaseScannedEntries   = 0;
        uint64_t changeCasePlannedRenames   = 0;
        uint64_t changeCaseCompletedRenames = 0;

        // Change Attributes payload (Kind::ChangeAttributes)
        bool changeAttributesEnumerating = false;
        bool changeAttributesApplying    = false;
        std::filesystem::path changeAttributesCurrentPath;
        uint64_t changeAttributesScannedFolders  = 0;
        uint64_t changeAttributesScannedEntries  = 0;
        uint64_t changeAttributesPlannedItems    = 0;
        uint64_t changeAttributesCompletedItems  = 0;

        bool finished    = false;
        HRESULT resultHr = S_OK;
        std::wstring doneSummary;
    };

    // Informational tasks are read-only task cards displayed in the File Operations popup for background work
    // that isn't a file operation (e.g., Compare Directories scan/content progress).
    [[nodiscard]] uint64_t CreateOrUpdateInformationalTask(const InformationalTaskUpdate& update) noexcept;
    void DismissInformationalTask(uint64_t taskId) noexcept;

    // DPI handling
    void OnDpiChanged(float newDpi);

    void ApplyTheme(const AppTheme& theme);
    [[maybe_unused]] const AppTheme& GetTheme() const noexcept
    {
        return _theme;
    }

    HRESULT ReloadFileSystemPlugins() noexcept;
    HRESULT SetFileSystemPluginForPane(Pane pane, std::wstring_view pluginId) noexcept;
    HRESULT SetFileSystemInstanceForPane(
        Pane pane, wil::com_ptr<IFileSystem> fileSystem, std::wstring pluginId, std::wstring pluginShortId, std::wstring instanceContext) noexcept;
    [[maybe_unused]] std::wstring_view GetFileSystemPluginId(Pane pane) const noexcept;
    [[maybe_unused]] std::wstring_view GetFileSystemPluginShortId(Pane pane) const noexcept;
    [[maybe_unused]] std::wstring_view GetFileSystemInstanceContext(Pane pane) const noexcept;
    [[nodiscard]] wil::com_ptr<IFileSystem> GetFileSystem(Pane pane) const noexcept;

    void DebugShowOverlaySample(Pane pane, FolderView::OverlaySeverity severity);
    void DebugShowOverlaySampleNonModal(Pane pane, FolderView::OverlaySeverity severity);
    void DebugShowOverlaySampleBusyWithCancel(Pane pane);
    void DebugShowOverlaySampleCanceled(Pane pane);
    void DebugHideOverlaySample(Pane pane);

    struct FileOperationState;

#ifdef ENABLE_TESTS
    struct FolderWindowPaneStatusBarDebugSnapshot
    {
        bool visible                      = false;
        bool usesNativeStatusBarClass     = false;
        bool usesDirectWriteTextRendering = false;
        bool hasNativeFont                = false;
        bool activePane                   = false;
        bool selectionTextDimmed          = false;
        float textSizeDip                 = 0.0f;
        std::wstring className;
        std::wstring selectionText;
        std::wstring securityText;
        std::wstring sortText;
    };

    struct FolderWindowFunctionBarDebugSnapshot
    {
        bool visible                    = false;
        bool windowVisible              = false;
        bool usesDirectWriteTextMetrics = false;
        RECT rect{};
    };

    struct FolderWindowSplitterDebugSnapshot
    {
        RECT splitterRect{};
        RECT leftArrowRect{};
        RECT rightArrowRect{};
        Pane leftArrowTargetPane  = Pane::Left;
        Pane rightArrowTargetPane = Pane::Right;
        wchar_t leftArrowGlyph    = L'<';
        wchar_t rightArrowGlyph   = L'>';
        COLORREF arrowColor       = 0;
        COLORREF gripColor        = 0;
        int arrowChevronSizePx    = 0;
        int gripDotSizePx         = 0;
        std::optional<Pane> hoveredArrowPane;
        bool leftArrowCursorHand  = false;
        bool rightArrowCursorHand = false;
    };

    struct PaneViewOptionsDebugSnapshot
    {
        bool fileExtensionsVisible      = true;
        bool navigationBarVisible       = true;
        bool navigationViewWindowVisible = false;
        bool filterBarVisible           = false;
        bool filterBarWindowVisible     = false;
        bool filterBarUsesDxUiHost      = false;
        bool filterEnabled              = false;
        bool thumbnailsVisible          = false;
        float thumbnailTargetDip        = 16.0f;
        uint64_t thumbnailQueuedCount   = 0;
        uint64_t thumbnailCompletedCount = 0;
        uint64_t thumbnailFallbackCount = 0;
        uint64_t thumbnailStaleDropCount = 0;
        uint64_t thumbnailPendingCount  = 0;
        uint64_t thumbnailCacheHitCount = 0;
        std::wstring filterText;
        std::wstring filterBarText;
        std::wstring focusedItemRealDisplayName;
        std::wstring focusedItemVisualDisplayName;
    };

    struct PreviewPaneDebugSnapshot
    {
        bool active                = false;
        Pane sourcePane            = Pane::Left;
        Pane hostPane              = Pane::Right;
        bool tabsVisible           = false;
        bool tabsUseDxUiHost       = false;
        bool previewTabSelected    = false;
        bool folderTabSelected     = true;
        bool previewContentVisible = false;
        bool previewContentUsesDxUiHost = false;
        bool previewUsesEmbeddedViewer  = false;
        bool folderViewVisible     = true;
        bool previewCloseButtonVisible = false;
        HWND previewTabsHwnd       = nullptr;
        HWND previewContentHwnd    = nullptr;
        HWND previewEmbeddedViewerHwnd = nullptr;
        uintptr_t previewViewerInstanceId = 0;
        RECT tabRect{};
        RECT folderTabClientRect{};
        RECT previewTabClientRect{};
        RECT previewCloseClientRect{};
        RECT contentRect{};
        RECT clientRect{};
        RECT functionBarRect{};
        std::filesystem::path previewedPath;
        std::wstring previewText;
        std::wstring previewViewerPluginId;
        std::wstring previewTabsTooltipText;
        std::wstring previewTabsPendingTooltipText;
        uint64_t previewBytes = 0;
        std::wstring sourceFocusedDisplayName;
    };

    struct OpenedFilesDebugRow
    {
        std::filesystem::path path;
        std::wstring file;
        std::wstring source;
        std::wstring openedBy;
        Pane pane = Pane::Left;
        bool focusable = false;
    };

    struct OpenedFilesDebugSnapshot
    {
        bool visible = false;
        bool usesDxUiHost = false;
        size_t visibleChildWindowCount = 0u;
        size_t visibleNativeChildControlCount = 0u;
        COLORREF themeWindowBackground = CLR_INVALID;
        COLORREF themeText = CLR_INVALID;
        std::wstring dialogClassName;
        bool emptyStateVisible = false;
        size_t selectedIndex = static_cast<size_t>(-1);
        std::vector<OpenedFilesDebugRow> rows;
    };

    struct SharedDirectoryDebugRow
    {
        std::wstring name;
        std::wstring localPath;
        std::wstring type;
        std::wstring remark;
        bool openable = false;
    };

    struct SharedDirectoriesDebugSnapshot
    {
        bool visible = false;
        bool emptyStateVisible = false;
        size_t selectedIndex = static_cast<size_t>(-1);
        HRESULT lastError = S_OK;
        std::vector<SharedDirectoryDebugRow> rows;
    };

    struct SharedDirectoriesDebugProviderResult
    {
        HRESULT hr = S_OK;
        std::vector<SharedDirectoryDebugRow> rows;
    };

    struct ArchiveCommandDebugOptions
    {
        std::filesystem::path archivePath;
        std::filesystem::path destinationPath;
        bool overwrite = true;
    };

    struct ArchiveCommandDebugResult
    {
        std::wstring operation;
        HRESULT hr = S_OK;
        std::filesystem::path archivePath;
        std::filesystem::path destinationPath;
        uint64_t entryCount = 0u;
        uint64_t bytesProcessed = 0u;
        std::vector<std::wstring> entries;
    };

    // Debug/testing hook: access the file-operations state for automation/self-tests.
    // This will initialize file operations if they are not yet created.
    FileOperationState* DebugGetFileOperationState() noexcept;

    enum class DebugShellActionKind : uint8_t
    {
        ContextMenuCurrentDirectory,
        OpenSecurity,
    };

    struct DebugShellAction final
    {
        DebugShellActionKind kind = DebugShellActionKind::ContextMenuCurrentDirectory;
        Pane pane                = Pane::Left;
        std::filesystem::path path;
        std::wstring propertyPage;
    };

    using DebugShellActionCallback = std::function<HRESULT(const DebugShellAction&)>;
    void DebugSetShellActionCallback(DebugShellActionCallback callback);
    void DebugSetNextChangeAttributesOptions(std::optional<ChangeAttributesOptions> options);
    [[nodiscard]] std::optional<ChangeAttributesReport> DebugGetLastChangeAttributesReport() const;

    using DebugShellNewTemplateKind       = ShellNewTemplateKind;
    using DebugShellNewTemplateDefinition = ShellNewTemplateDefinition;

    void DebugSetShellNewTemplatesForTest(std::optional<std::vector<DebugShellNewTemplateDefinition>> templates);
    void DebugSetNextShellNewFileNameForTest(std::optional<std::wstring> fileName);

    [[nodiscard]] size_t DebugGetViewerInstanceCount() const noexcept;
    [[nodiscard]] bool DebugHasViewerPluginId(std::wstring_view viewerPluginId) const noexcept;
    [[nodiscard]] uint64_t DebugGetForceRefreshCount(Pane pane) const noexcept;
    [[nodiscard]] bool DebugGetPaneStatusBarSnapshot(Pane pane, FolderWindowPaneStatusBarDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugClickPaneStatusBarSort(Pane pane) noexcept;
    [[nodiscard]] bool DebugGetFunctionBarSnapshot(FolderWindowFunctionBarDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugGetSplitterSnapshot(FolderWindowSplitterDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugHoverSplitterArrow(Pane pane) noexcept;
    [[nodiscard]] bool DebugClickSplitterArrow(Pane pane) noexcept;
    [[nodiscard]] bool DebugGetPaneViewOptionsSnapshot(Pane pane, PaneViewOptionsDebugSnapshot& out) const;
    void DebugSetThumbnailProviderMode(Pane pane, FolderView::DebugThumbnailProviderMode mode) noexcept;
    [[nodiscard]] bool DebugGetPreviewPaneSnapshot(PreviewPaneDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugSetPreviewPaneTab(Pane hostPane, bool previewTab) noexcept;
    [[nodiscard]] bool DebugAdvancePreviewTabsTooltipDelayForTest(Pane hostPane) noexcept;
    [[nodiscard]] bool DebugGetOpenedFilesDialogSnapshot(OpenedFilesDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugSelectOpenedFilesDialogRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugInvokeOpenedFilesDialogFocusItem() noexcept;
    void DebugCloseOpenedFilesDialogForTest() noexcept;
    void DebugAddOpenedExternalEditorForTest(const std::filesystem::path& path, std::wstring_view openedBy, Pane pane, bool closed) noexcept;
    void DebugClearOpenedExternalEditorsForTest() noexcept;
    [[nodiscard]] bool DebugGetSharedDirectoriesDialogSnapshot(SharedDirectoriesDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugSelectSharedDirectoriesDialogRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugInvokeSharedDirectoriesDialogOpenPath() noexcept;
    void DebugCloseSharedDirectoriesDialogForTest() noexcept;
    void DebugSetSharedDirectoriesProviderResultForTest(SharedDirectoriesDebugProviderResult result) noexcept;
    void DebugClearSharedDirectoriesProviderForTest() noexcept;
    void DebugSetNextArchiveCommandOptionsForTest(ArchiveCommandDebugOptions options) noexcept;
    void DebugClearArchiveCommandOptionsForTest() noexcept;
    [[nodiscard]] std::optional<ArchiveCommandDebugResult> DebugGetLastArchiveCommandResultForTest() const noexcept;

    [[nodiscard]] std::wstring_view DebugGetFocusedItemDisplayName(Pane pane) const noexcept;
    [[nodiscard]] bool DebugHasItemDisplayName(Pane pane, std::wstring_view displayName) const noexcept;
    [[nodiscard]] size_t DebugGetItemCount(Pane pane) const noexcept;
    [[nodiscard]] size_t DebugGetPaneBitmapIconCount(Pane pane) const noexcept;
    [[nodiscard]] bool DebugIsItemSelected(Pane pane, std::wstring_view displayName) const noexcept;
    [[nodiscard]] size_t DebugGetSelectedCount(Pane pane) const noexcept;
    [[nodiscard]] uint64_t DebugGetWarmPaneRenderingCallCount(Pane pane) const noexcept;
    [[nodiscard]] FolderView::DebugWarmPerfSnapshot DebugGetWarmPanePerfSnapshot(Pane pane) const noexcept;
    [[nodiscard]] bool DebugWarmPaneRendering(Pane pane) noexcept;
    [[nodiscard]] bool DebugIsEmptyFolderStateActive(Pane pane) const noexcept;
    [[nodiscard]] std::wstring_view DebugGetEmptyFolderFunMessage(Pane pane) const noexcept;
    [[nodiscard]] FolderView::DebugEmptyFolderItemMetrics DebugGetEmptyFolderItemMetrics(Pane pane) const noexcept;
    [[nodiscard]] HWND DebugGetNavigationViewHwnd(Pane pane) const noexcept;
    [[nodiscard]] bool DebugGetNavigationViewSnapshot(Pane pane, NavigationViewDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugFocusNavigationViewRegion(Pane pane, NavigationView::FocusRegion region) noexcept;
    [[nodiscard]] bool DebugFocusItemByDisplayName(Pane pane, std::wstring_view displayName) noexcept;
    [[nodiscard]] bool DebugGetIncrementalSearchSnapshot(Pane pane, FolderView::IncrementalSearchDebugSnapshot& out) const noexcept;
    struct CommandLineDebugSnapshot
    {
        bool visible          = false;
        bool hasKeyboardFocus = false;
        Pane pane             = Pane::Left;
        HWND editHwnd         = nullptr;
        std::wstring text;
        std::filesystem::path workingDirectory;
    };
    using CommandLineLaunchCallback = std::function<HRESULT(std::wstring_view commandLine, const std::filesystem::path& workingDirectory)>;
    [[nodiscard]] bool DebugGetCommandLineSnapshot(CommandLineDebugSnapshot& out) const noexcept;
    void DebugSetCommandLineTextForTest(std::wstring_view text);
    void DebugSetCommandLineLaunchCallback(CommandLineLaunchCallback callback);
    [[nodiscard]] FolderView::NameFilterState DebugGetNameFilterState(Pane pane) const;
    [[nodiscard]] bool DebugIsNameFilterActive(Pane pane) const noexcept;
    void DebugResetPaneVisibilityState(Pane pane) noexcept;
    [[nodiscard]] FolderView::FilterWatermarkVisualMode DebugGetFilterWatermarkVisualMode(Pane pane) const noexcept;
    [[nodiscard]] bool DebugGetPaneAlertSnapshot(Pane pane, FolderView::AlertOverlayDebugSnapshot& out) const noexcept;
#endif

    void ShowPaneAlertOverlay(Pane pane,
                              FolderView::ErrorOverlayKind kind,
                              FolderView::OverlaySeverity severity,
                              std::wstring title,
                              std::wstring message,
                              HRESULT hr       = S_OK,
                              bool closable    = true,
                              bool blocksInput = true);

    void DismissPaneAlertOverlay(Pane pane);

    struct FileOperationStateDeleter
    {
        void operator()(FileOperationState* state) const noexcept;
    };

private:
    enum class CopySelectionTextMode : uint8_t
    {
        PathAndName,
        Name,
        Path,
        UncPathAndName,
    };

    enum class SplitterArrowZone : uint8_t
    {
        None,
        Left,
        Right,
    };

    struct PaneState;

    // Class registration
    static ATOM RegisterWndClass(HINSTANCE instance);
    static constexpr PCWSTR kClassName = L"RedSalamander.FolderWindow";

    // Window procedure
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    // Message handlers
    bool OnCreate(HWND hwnd) noexcept;
    void OnDestroy();
    void OnSize(UINT width, UINT height);
    void OnSetFocus();
    void OnPaint();
    void OnLButtonDown(POINT pt);
    void OnLButtonDblClk(POINT pt);
    void OnLButtonUp();
    void OnMouseMove(POINT pt);
    void OnMouseLeave();
    void OnCaptureChanged();
    LRESULT OnDrawItem(DRAWITEMSTRUCT* dis);
    LRESULT OnNotify(const NMHDR* header);
    LRESULT OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg);
    bool OnSetCursor(POINT pt);
    void OnParentNotify(UINT eventMsg, UINT childId);
    LRESULT OnDeviceChange(UINT event, LPARAM data) noexcept;
    void OnNetworkConnectivityChanged() noexcept;
    void CopySelectionText(Pane pane, CopySelectionTextMode mode, UINT titleStringId);
    LRESULT OnPaneSelectionSizeComputed(LPARAM lp) noexcept;
    LRESULT OnPaneSelectionSizeProgress(LPARAM lp) noexcept;
    LRESULT OnFileOperationCompleted(LPARAM lp) noexcept;
    LRESULT OnChangeCaseTaskUpdate(LPARAM lp) noexcept;
    LRESULT OnChangeCaseCompleted(LPARAM lp) noexcept;
    LRESULT OnChangeAttributesTaskUpdate(LPARAM lp) noexcept;
    LRESULT OnChangeAttributesCompleted(LPARAM lp) noexcept;
    static LRESULT CALLBACK CommandLineEditWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT CommandLineEditWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    // File operations (internal implementation in FolderWindow.FileOperations.cpp)
    void EnsureFileOperations();
    HRESULT StartFileOperationFromFolderView(Pane pane, FolderView::FileOperationRequest request) noexcept;
    HRESULT ShowItemPropertiesFromFolderView(Pane pane, std::filesystem::path path) noexcept;
    void ShutdownFileOperations() noexcept;
    void ApplyFileOperationsTheme() noexcept;

    // Layout
    void CalculateLayout();
    void AdjustChildWindows();
    void UpdatePaneStatusBar(Pane pane);
    void UpdatePaneFilterBar(Pane pane);
    void UpdatePaneFocusStates() noexcept;
    void FocusPaneFolderView(Pane pane) noexcept;
    void FocusPanePreferredTarget(Pane pane) noexcept;
    [[nodiscard]] HWND GetPanePreferredFocusTarget(Pane pane) const noexcept;
    [[nodiscard]] RECT GetSplitterArrowRect(SplitterArrowZone zone) const noexcept;
    [[nodiscard]] Pane GetSplitterArrowTargetPane(SplitterArrowZone zone) const noexcept;
    [[nodiscard]] wchar_t GetSplitterArrowGlyph(SplitterArrowZone zone) const noexcept;
    [[nodiscard]] COLORREF GetSplitterGripColor() const noexcept;
    [[nodiscard]] COLORREF GetSplitterArrowColor() const noexcept;
    [[nodiscard]] int GetSplitterGripDotSizePx() const noexcept;
    [[nodiscard]] int GetSplitterArrowChevronSizePx() const noexcept;
    [[nodiscard]] SplitterArrowZone HitTestSplitterArrow(POINT pt) const noexcept;
    void SetHoveredSplitterArrowZone(SplitterArrowZone zone) noexcept;
    void TrackSplitterMouseLeave() noexcept;
    [[nodiscard]] bool CreateCommandLineControls(HWND parent) noexcept;
    void DestroyCommandLineControls() noexcept;
    void ShowCommandLine(Pane pane, const std::filesystem::path& workingDirectory);
    void HideCommandLine(bool restoreFocus) noexcept;
    [[nodiscard]] std::wstring GetCommandLineText() const;
    void SetCommandLineText(std::wstring_view text);
    void InsertCommandLineText(std::wstring_view text);
    [[nodiscard]] std::optional<std::filesystem::path> ResolveCommandLineWorkingDirectory(Pane pane) const;
    [[nodiscard]] HRESULT LaunchCommandLine(std::wstring_view commandLine, const std::filesystem::path& workingDirectory);
    void ExecuteCommandLineFromEdit();
    void StartSelectionSizeWorker(Pane pane) noexcept;
    void CancelSelectionSizeComputation(Pane pane) noexcept;
    void RequestSelectionSizeComputation(Pane pane);
    void SelectionSizeWorkerMain(Pane pane, std::stop_token stopToken) noexcept;

    // Path synchronization
    void OnNavigationPathChanged(Pane pane, const std::optional<std::filesystem::path>& path);
    void OnFolderViewPathChanged(Pane pane, const std::optional<std::filesystem::path>& path);
    void RecordNavigationHistory(PaneState& state, const std::filesystem::path& displayPath);
    void TrimNavigationHistory(PaneState& state);
    void OnFolderViewDirectoryImpact(Pane pane, const DirectoryInfoCache::DirectoryImpact& impact) noexcept;
    void OnFolderViewNavigateUpFromRoot(Pane pane) noexcept;
    HRESULT EnsurePaneFileSystem(Pane pane, std::wstring_view pluginId) noexcept;
    Pane GetPaneFromChild(HWND child) const noexcept;
    bool TryOpenFileAsVirtualFileSystem(Pane pane, const std::filesystem::path& path) noexcept;
    enum class FileActionFailureKind : uint8_t
    {
        LaunchFailed,
    };
    struct FileActionFailure final
    {
        FileActionFailureKind kind = FileActionFailureKind::LaunchFailed;
        bool viewerAction         = true;
        std::wstring actionId;
        std::filesystem::path targetPath;
        HRESULT hr = S_OK;
    };
    void ClearFileActionFailure() noexcept;
    void RecordFileActionLaunchFailure(bool viewerAction, std::wstring_view actionId, const std::filesystem::path& targetPath, HRESULT hr) noexcept;
    [[nodiscard]] std::optional<FileActionFailure> TakeFileActionFailure() noexcept;
    [[nodiscard]] bool ShowRecordedFileActionFailureOverlay(Pane pane) noexcept;
    [[nodiscard]] static Pane OppositePane(Pane pane) noexcept
    {
        return pane == Pane::Left ? Pane::Right : Pane::Left;
    }
    void SetPreviewPaneTab(Pane hostPane, bool previewTab) noexcept;
    void ClosePreviewPane() noexcept;
    void RequestPreviewPaneRefresh() noexcept;
    void CancelPendingPreviewPaneRefresh() noexcept;
    void OnPreviewPaneRefreshTimer() noexcept;
    void RefreshPreviewPane() noexcept;
    void UpdatePreviewFolderTabTooltip(Pane hostPane) noexcept;
    void UpdatePreviewTabSelection(Pane hostPane) noexcept;
    void ClosePreviewViewer(Pane hostPane) noexcept;
    void LayoutEmbeddedPreviewViewer(Pane hostPane) noexcept;
    bool OpenPreviewFocusedPathWithViewer(Pane sourcePane, Pane hostPane) noexcept;
    void SetPreviewPlaceholder(Pane hostPane, std::wstring text) noexcept;
    void UpdateFilterBarLayout(Pane pane) noexcept;
    void UpdatePreviewContentLayout(Pane pane) noexcept;
    LRESULT HandlePaneDxHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept;
    [[nodiscard]] std::wstring BuildPreviewTextForPath(Pane sourcePane, const std::filesystem::path& path, uint64_t& outBytes) noexcept;
    bool TryViewFileWithViewer(Pane pane, const FolderView::ViewFileRequest& request) noexcept;
    bool TryEditFocusedFileWithEditor(Pane pane, std::wstring_view actionId, bool alternate) noexcept;
    bool TryEditFileWithEditor(Pane pane,
                               const std::filesystem::path& filePath,
                               const std::vector<std::filesystem::path>& selectedPaths,
                               std::wstring_view actionId,
                               bool alternate) noexcept;
    bool TryViewSpaceWithViewer(Pane pane, const std::filesystem::path& folderPath) noexcept;
    enum class OpenedFileSourceKind : uint8_t
    {
        Viewer,
        Editor,
        Preview,
    };
    struct ViewerInstance final
    {
        std::wstring viewerPluginId;
        std::wstring openedBy;
        OpenedFileSourceKind source = OpenedFileSourceKind::Viewer;
        Pane pane = Pane::Left;
        wil::com_ptr<IViewer> viewer;
        ViewerOpenContext openContext{};
        bool hasInitialConfigurationJson = false;
        std::string initialConfigurationJson;
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring fileSystemName;
        std::wstring focusedPath;
        std::vector<std::wstring> selectionStorage;
        std::vector<const wchar_t*> selectionPointers;
        std::vector<std::wstring> otherFilesStorage;
        std::vector<const wchar_t*> otherFilePointers;
    };

    HRESULT OpenViewerWithPluginInternal(std::wstring_view pluginId,
                                          const ViewerOpenContext& context,
                                          std::wstring_view openedBy,
                                          Pane pane,
                                          OpenedFileSourceKind source,
                                          ViewerInstance** outInstance) noexcept;
    void UpdateViewerInstanceContext(ViewerInstance& instance,
                                     const ViewerOpenContext& context,
                                     std::wstring_view openedBy,
                                     Pane pane,
                                     OpenedFileSourceKind source) noexcept;
    HRESULT ReopenViewerInstance(ViewerInstance& instance,
                                  const ViewerOpenContext& context,
                                  std::wstring_view openedBy,
                                  Pane pane,
                                  OpenedFileSourceKind source) noexcept;
    void PersistViewerConfiguration(ViewerInstance& instance) noexcept;

    struct ViewerCallbackState final : public IViewerCallback
    {
        FolderWindow* owner = nullptr;
        HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override;
    };

    void ShutdownViewers() noexcept;
    void ApplyViewerTheme() noexcept;
    ViewerTheme BuildViewerTheme() const noexcept;
    HRESULT OnViewerClosed(ViewerInstance* instance) noexcept;
    struct OpenedFileRow final
    {
        std::filesystem::path path;
        std::wstring file;
        std::wstring source;
        std::wstring openedBy;
        Pane pane = Pane::Left;
        bool focusable = false;
    };
    struct OpenedExternalFileEntry final
    {
        OpenedExternalFileEntry() = default;
        OpenedExternalFileEntry(const OpenedExternalFileEntry&) = delete;
        OpenedExternalFileEntry& operator=(const OpenedExternalFileEntry&) = delete;
        OpenedExternalFileEntry(OpenedExternalFileEntry&&) noexcept = default;
        OpenedExternalFileEntry& operator=(OpenedExternalFileEntry&&) noexcept = default;

        uint64_t id = 0;
        OpenedFileSourceKind source = OpenedFileSourceKind::Editor;
        std::filesystem::path path;
        std::wstring openedBy;
        Pane pane = Pane::Left;
        DWORD processId = 0;
        wil::unique_handle processHandle;
#ifdef ENABLE_TESTS
        bool debugClosed = false;
#endif
    };
    struct OpenedFilesDialogState;
    struct OpenedFilesDialogStateDeleter final
    {
        void operator()(OpenedFilesDialogState* state) const noexcept;
    };
    void RegisterOpenedExternalFile(OpenedFileSourceKind source,
                                    const std::filesystem::path& path,
                                    std::wstring_view openedBy,
                                    Pane pane,
                                    wil::unique_handle processHandle,
                                    DWORD processId) noexcept;
    void PruneClosedOpenedExternalFiles() noexcept;
    [[nodiscard]] std::vector<OpenedFileRow> CollectOpenedFileRows() noexcept;
    [[nodiscard]] std::wstring ResolveOpenedFilesViewerName(const ViewerInstance& instance) const;
    void RefreshOpenedFilesDialogRows() noexcept;
    void ApplyOpenedFilesDialogTheme() noexcept;
    void CloseOpenedFilesDialog() noexcept;
    [[nodiscard]] bool FocusOpenedFilesDialogSelection() noexcept;
    [[nodiscard]] bool FocusOpenedFileRow(const OpenedFileRow& row) noexcept;

    struct SharedDirectoryRow final
    {
        std::wstring name;
        std::wstring localPath;
        std::wstring type;
        std::wstring remark;
        bool openable = false;
    };
    struct SharedDirectoriesDialogState final
    {
        SharedDirectoriesDialogState() = default;
        SharedDirectoriesDialogState(const SharedDirectoriesDialogState&) = delete;
        SharedDirectoriesDialogState& operator=(const SharedDirectoriesDialogState&) = delete;
        SharedDirectoriesDialogState(SharedDirectoriesDialogState&&) noexcept = default;
        SharedDirectoriesDialogState& operator=(SharedDirectoriesDialogState&&) noexcept = default;

        FolderWindow* owner = nullptr;
        Pane pane = Pane::Left;
        wil::unique_hwnd hwnd;
        std::vector<SharedDirectoryRow> rows;
        size_t selectedIndex = static_cast<size_t>(-1);
        HRESULT lastError = S_OK;
        bool destroyed = false;
    };
    [[nodiscard]] std::vector<SharedDirectoryRow> CollectSharedDirectoryRows(HRESULT& outHr) const noexcept;
    void RefreshSharedDirectoriesDialogRows() noexcept;
    void CloseSharedDirectoriesDialog() noexcept;
    [[nodiscard]] bool OpenSharedDirectoriesDialogSelection() noexcept;
    void OpenSharedDirectoriesManagement() noexcept;
    static INT_PTR CALLBACK SharedDirectoriesDialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

private:
    class NetworkChangeSubscription;

    wil::unique_hwnd _hWnd;
    HINSTANCE _hInstance = nullptr;
    UINT _dpi            = USER_DEFAULT_SCREEN_DPI;

    // Child components
    struct PaneState
    {
        PaneState() = default;
        // Explicitly delete copy/move operations
        PaneState(const PaneState&)            = delete;
        PaneState(PaneState&&)                 = delete;
        PaneState& operator=(const PaneState&) = delete;
        PaneState& operator=(PaneState&&)      = delete;

        NavigationView navigationView;
        FolderView folderView;
        wil::unique_hwnd hNavigationView;
        wil::unique_hwnd hFolderView;
        wil::unique_hwnd hFilterBar;
        RedSalamander::DxUi::WindowHost filterBarHost;
        RedSalamander::DxUi::Label* filterBarLabel = nullptr;
        wil::unique_hwnd hStatusBar;
        wil::unique_hwnd hPreviewTabs;
        RedSalamander::DxUi::WindowHost previewTabsHost;
        RedSalamander::DxUi::TabControl* previewTabsControl = nullptr;
        wil::unique_hwnd hPreviewContent;
        RedSalamander::DxUi::WindowHost previewContentHost;
        RedSalamander::DxUi::Label* previewContentLabel = nullptr;
        bool navigationBarVisible = true;
        bool filterBarVisible     = false;
        bool statusBarVisible = true;
        bool previewTabsVisible = false;
        bool previewTabSelected = false;
        std::filesystem::path previewedPath;
        std::wstring previewText;
        std::wstring previewViewerPluginId;
        ViewerInstance* previewViewerInstance = nullptr;
        uint64_t previewBytes = 0;
        FolderView::SelectionStats selectionStats{};
        uint64_t selectionSizeGeneration = 0;
        std::jthread selectionSizeThread;
        std::mutex selectionSizeMutex;
        std::condition_variable selectionSizeCv;
        bool selectionSizeWorkPending        = false;
        uint64_t selectionSizeWorkGeneration = 0;
        std::vector<std::filesystem::path> selectionSizeWorkFolders;
        wil::com_ptr<IFileSystem> selectionSizeWorkFileSystem;
        std::shared_ptr<std::stop_source> selectionSizeWorkStopSource;

        std::jthread changeCaseThread;
        std::jthread changeAttributesThread;
        bool selectionFolderBytesPending = false;
        bool selectionFolderBytesValid   = false;
        uint64_t selectionFolderBytes    = 0;
        std::wstring statusSelectionText;
        std::wstring statusSortText;
        std::wstring statusSecurityText;
        std::wstring filterBarText;
        uint32_t statusFocusHueDegrees = 0;
        bool sortIndicatorHot          = false;

        wil::unique_hmodule fileSystemModule;
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring pluginId;
        std::wstring pluginShortId;
        std::wstring instanceContext;

        std::optional<std::filesystem::path> currentPath;
        bool updatingPath = false;

        std::vector<std::filesystem::path> navigationHistory;
        size_t navigationHistoryIndex       = 0;
        bool navigationHistorySuspendRecord = false;
    };
    bool SanityCheckBothPanes(PaneState& src, PaneState& dest, FileSystemOperation operation);

    PaneState _leftPane;
    PaneState _rightPane;
    std::optional<FileActionFailure> _lastFileActionFailure;
    Pane _activePane           = Pane::Left;
    HWND _lastFocusedPaneChild = nullptr;
    FunctionBar _functionBar;
    bool _functionBarVisible                = true;
    const ShortcutManager* _shortcutManager = nullptr;

    // Layout
    SIZE _clientSize{};
    RECT _leftPaneRect{};
    RECT _rightPaneRect{};
    RECT _splitterRect{};
    RECT _leftNavigationRect{};
    RECT _leftFilterBarRect{};
    RECT _leftFolderViewRect{};
    RECT _leftStatusBarRect{};
    RECT _leftPreviewTabsRect{};
    RECT _leftPreviewContentRect{};
    RECT _rightNavigationRect{};
    RECT _rightFilterBarRect{};
    RECT _rightFolderViewRect{};
    RECT _rightStatusBarRect{};
    RECT _rightPreviewTabsRect{};
    RECT _rightPreviewContentRect{};
    RECT _commandLineRect{};
    RECT _commandLineLabelRect{};
    RECT _commandLineEditRect{};
    RECT _functionBarRect{};
    float _splitRatio                  = 0.5f;
    bool _viewWidthAdjustActive        = false;
    float _viewWidthAdjustRestoreRatio = 0.5f;
    std::optional<float> _zoomRestoreSplitRatio;
    std::optional<Pane> _zoomedPane;
    bool _draggingSplitter                      = false;
    int _splitterDragOffsetPx                   = 0;
    SplitterArrowZone _hoveredSplitterArrowZone = SplitterArrowZone::None;
    bool _trackingSplitterMouseLeave            = false;
    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _splitterBrush;
    wil::unique_hbrush _splitterGripBrush;
    wil::unique_hbrush _splitterArrowHoverBrush;
    wil::unique_hwnd _hCommandLineLabel;
    wil::unique_hwnd _hCommandLineEdit;
    bool _commandLineVisible = false;
    Pane _commandLinePane    = Pane::Left;
    std::filesystem::path _commandLineWorkingDirectory;

    AppTheme _theme;
    uint32_t _statusBarRainbowHueDegrees = 0;
    ShowSortMenuCallback _showSortMenuCallback;
    PanePathChangedCallback _panePathChangedCallback;
    FileOperationCompletedCallback _fileOperationCompletedCallback;

    std::unique_ptr<FileOperationState, FileOperationStateDeleter> _fileOperations;
#ifdef ENABLE_TESTS
    DebugShellActionCallback _debugShellActionCallback;
    std::optional<ChangeAttributesOptions> _debugNextChangeAttributesOptions;
    std::optional<ChangeAttributesReport> _debugLastChangeAttributesReport;
    std::optional<std::vector<ShellNewTemplateDefinition>> _debugShellNewTemplates;
    std::optional<std::wstring> _debugNextShellNewFileName;
    CommandLineLaunchCallback _debugCommandLineLaunchCallback;
#endif
    Common::Settings::Settings* _settings = nullptr;
    bool _showHiddenFiles                 = true;
    bool _showSystemFiles                 = true;
    uint32_t _folderHistoryMax            = 20u;
    std::vector<std::filesystem::path> _folderHistory;

    struct SavedSelection final
    {
        // Saved for future restore improvements (e.g., validating the source folder/plugin context or offering a plugin-aware restore).
        // Current restore selects by display name only.
        std::wstring sourcePluginId;
        std::wstring sourceInstanceContext;
        std::filesystem::path sourceFolder;
        std::vector<std::wstring> displayNames;
    };
    std::optional<SavedSelection> _savedSelection;
    std::optional<Pane> _previewSourcePane;
    bool _previewRefreshPending = false;

    ViewerCallbackState _viewerCallback;
    std::vector<std::unique_ptr<ViewerInstance>> _viewerInstances;
    std::vector<OpenedExternalFileEntry> _openedExternalFiles;
    uint64_t _nextOpenedExternalFileId = 1;
    std::unique_ptr<OpenedFilesDialogState, OpenedFilesDialogStateDeleter> _openedFilesDialog;
    std::unique_ptr<SharedDirectoriesDialogState> _sharedDirectoriesDialog;
#ifdef ENABLE_TESTS
    std::optional<SharedDirectoriesDebugProviderResult> _debugSharedDirectoriesProviderResult;
    std::optional<ArchiveCommandDebugOptions> _debugNextArchiveCommandOptions;
    std::optional<ArchiveCommandDebugResult> _debugLastArchiveCommandResult;
#endif

    std::unique_ptr<NetworkChangeSubscription> _networkChangeSubscription;
    uint64_t _lastNetworkConnectivityRefreshTick = 0;

    friend LRESULT CALLBACK FolderWindowDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
};

#pragma once
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <Ole2.h>

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdint>

#include <d2d1_1.h>
#include <d3d11_1.h>
#include <dwrite.h>
#include <dxgi1_3.h>

#include <filesystem>
#include <functional>
#include <memory>
#include <mutex>

#include <objidl.h>
#include <oleidl.h>

#include <deque>
#include <optional>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "MaskSyntax.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "dxgi")
#pragma comment(lib, "Windowscodecs.lib")

#include "AppTheme.h"
#include "DirectoryInfoCache.h"
#include "FolderViewColumnLayout.h"
#include "Helpers.h"
#include "SettingsThumbnail.h"

struct IWICImagingFactory;
struct IFileSystem;
struct PluginMetaData;
enum FileSystemOperation : uint32_t;
enum FileSystemFlags : uint32_t;
class ShortcutManager;
namespace RedSalamander::Ui
{
class AlertOverlay;
}

// Theme structures and Helpers are defined in AppTheme.h

#ifdef ENABLE_TESTS
struct FolderViewRenamePromptDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    std::wstring text;
    size_t selectionStart = 0u;
    size_t selectionEnd   = 0u;
    D2D1_RECT_F textRect{};
    D2D1_RECT_F selectionPaintRect{};
    float horizontalScrollDip  = 0.0f;
    bool hasSelectionPaintRect = false;
};

[[nodiscard]] HWND GetFolderViewRenamePromptHandle() noexcept;
[[nodiscard]] bool DebugGetFolderViewRenamePromptSnapshot(FolderViewRenamePromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFolderViewRenamePromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmFolderViewRenamePrompt() noexcept;
[[nodiscard]] bool DebugCancelFolderViewRenamePrompt() noexcept;
[[nodiscard]] D2D1_INTERPOLATION_MODE DebugResolveFolderViewIconBitmapInterpolation(D2D1_SIZE_U sourcePixelSize, float destinationSizeDip, float dpi) noexcept;
#endif

class FolderView
{
public:
    FolderView();
    ~FolderView();

    FolderView(const FolderView&)            = delete;
    FolderView& operator=(const FolderView&) = delete;

    static ATOM RegisterWndClass(HINSTANCE instance);

    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy();

    void SetFolderPath(const std::optional<std::filesystem::path>& folderPath);
    [[maybe_unused]] std::optional<std::filesystem::path> GetFolderPath() const
    {
        return _currentFolder;
    }

    // Updates the "remembered focus" entry for a specific folder so that the next time that
    // folder is enumerated, the requested item becomes the focused/active item (and is scrolled into view).
    void RememberFocusedItemForFolder(const std::filesystem::path& folder, std::wstring_view itemDisplayName) noexcept;

    // Prepares the view for an external command by clearing selection and focusing the specified item.
    // Returns true if the item was found and focused, false otherwise.
    bool PrepareForExternalCommand(std::wstring_view focusItemDisplayName) noexcept;

    // Queues a FolderView command (e.g. IDM_FOLDERVIEW_CONTEXT_*) to execute after the next successful
    // enumeration of `targetFolder`. The command is canceled if the view navigates to a different folder.
    void QueueCommandAfterNextEnumeration(UINT commandId, const std::filesystem::path& targetFolder, std::wstring_view expectedFocusDisplayName) noexcept;

    [[maybe_unused]] HWND GetHWND() const
    {
        return _hWnd.get();
    }

    void SetPaneFocused(bool focused) noexcept;

    void ForceRefresh();
    void RecordPendingInputToPaintStart(std::chrono::steady_clock::time_point inputStartedAt) noexcept;
    void RecordPendingInputToPaintStart(std::chrono::steady_clock::time_point inputStartedAt,
                                        std::wstring metricName,
                                        std::wstring detail,
                                        uint64_t value0 = 0u,
                                        uint64_t value1 = 0u) noexcept;
    void CancelPendingEnumeration();
    void OnDpiChanged(float newDpi);
    void SetFileSystem(const wil::com_ptr<IFileSystem>& fileSystem);
    const PluginMetaData* GetFileSystemMetadata() const;
    void SetFileSystemContext(std::wstring_view pluginId, std::wstring_view instanceContext)
    {
        _fileSystemPluginId.assign(pluginId);
        _fileSystemInstanceContext.assign(instanceContext);
    }
    [[nodiscard]] std::wstring_view GetFileSystemPluginId() const noexcept
    {
        return _fileSystemPluginId;
    }
    [[nodiscard]] std::wstring_view GetFileSystemInstanceContext() const noexcept
    {
        return _fileSystemInstanceContext;
    }

    void SetAppTheme(const AppTheme& theme);
    void SetTheme(const FolderViewTheme& theme);
    void SetMenuTheme(const MenuTheme& menuTheme);
    void SetShortcutManager(const ShortcutManager* shortcuts) noexcept;
    [[maybe_unused]] const FolderViewTheme& GetTheme() const
    {
        return _theme;
    }

    void CommandRename();
    void CommandView();
    [[nodiscard]] bool CommandAlternateView();
    [[nodiscard]] bool CommandViewWith(std::wstring_view actionId);
    void CommandDelete();
    HRESULT CopySelectedItemsToFolder(const std::filesystem::path& destinationFolder);
    HRESULT MoveSelectedItemsToFolder(const std::filesystem::path& destinationFolder);
    std::vector<std::filesystem::path> GetSelectedPaths() const;
    std::vector<std::filesystem::path> GetSelectedDirectoryPaths() const;
    [[nodiscard]] std::optional<std::filesystem::path> GetFocusedPath() const;
    std::vector<std::filesystem::path> GetSelectedOrFocusedPaths() const;
    std::vector<std::wstring> GetSelectedOrFocusedDisplayNames() const;
    [[nodiscard]] bool CutSelectionToClipboard();
    [[nodiscard]] bool PasteShortcutFromClipboard();
    void ActivateIncrementalSearch();

    struct PathAttributes
    {
        std::filesystem::path path;
        DWORD fileAttributes = 0;
    };
    std::vector<PathAttributes> GetSelectedOrFocusedPathAttributes() const;

    struct FileOperationRequest
    {
        FileSystemOperation operation = static_cast<FileSystemOperation>(0);
        std::vector<std::filesystem::path> sourcePaths;
        bool sourceContextSpecified = false;
        std::wstring sourcePluginId;
        std::wstring sourceInstanceContext;
        std::optional<std::filesystem::path> destinationFolder;
        FileSystemFlags flags = static_cast<FileSystemFlags>(0);
    };

    using FileOperationRequestCallback = std::function<HRESULT(FileOperationRequest request)>;
    void SetFileOperationRequestCallback(FileOperationRequestCallback callback)
    {
        _fileOperationRequestCallback = std::move(callback);
    }

    using PropertiesRequestCallback = std::function<HRESULT(std::filesystem::path path)>;
    void SetPropertiesRequestCallback(PropertiesRequestCallback callback)
    {
        _propertiesRequestCallback = std::move(callback);
    }
#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugHasFileOperationRequestCallback() const noexcept
    {
        return static_cast<bool>(_fileOperationRequestCallback);
    }

    static void DebugSetNextMoveSelectedItemsDestinationForSelfTest(std::optional<std::filesystem::path> destination) noexcept;
    static void DebugSetDirectFileOperationFallbackEnabledForSelfTest(bool enabled) noexcept;

    [[nodiscard]] uint64_t DebugGetForceRefreshCount() const noexcept
    {
        return _debugForceRefreshCount;
    }

    [[nodiscard]] uint64_t DebugGetWarmRenderingCallCount() const noexcept
    {
        return _debugWarmRenderingCallCount;
    }

    struct DebugWarmPerfSnapshot
    {
        uint64_t warmRenderingCalls             = 0;
        uint64_t deferredInitCalls              = 0;
        uint64_t renderCalls                    = 0;
        uint64_t queueIconLoadingCalls          = 0;
        uint64_t processIconQueueCalls          = 0;
        uint64_t batchIconUpdateCalls           = 0;
        uint64_t incrementalSearchEffectUpdates = 0;
    };

    [[nodiscard]] DebugWarmPerfSnapshot DebugGetWarmPerfSnapshot() const noexcept
    {
        return DebugWarmPerfSnapshot{
            .warmRenderingCalls             = _debugWarmRenderingCallCount,
            .deferredInitCalls              = _debugDeferredInitCallCount,
            .renderCalls                    = _debugRenderCallCount,
            .queueIconLoadingCalls          = _debugQueueIconLoadingCallCount,
            .processIconQueueCalls          = _debugProcessIconQueueCallCount,
            .batchIconUpdateCalls           = _debugBatchIconUpdateCallCount,
            .incrementalSearchEffectUpdates = _debugIncrementalSearchEffectUpdateCount,
        };
    }

    [[nodiscard]] std::wstring_view DebugGetFocusedDisplayName() const noexcept;
    [[nodiscard]] std::wstring DebugGetFocusedVisualDisplayName() const;
    [[nodiscard]] bool DebugHasItemDisplayName(std::wstring_view displayName) const noexcept;
    [[nodiscard]] size_t DebugGetItemCount() const noexcept;
    [[nodiscard]] size_t DebugGetBitmapIconCount() const noexcept;
    [[nodiscard]] bool DebugIsItemSelectedByDisplayName(std::wstring_view displayName) const noexcept;
    [[nodiscard]] size_t DebugGetSelectedItemCount() const noexcept;
    [[nodiscard]] bool DebugWarmRenderingForSelfTest() noexcept;
    [[nodiscard]] bool DebugIsEmptyFolderStateActive() const noexcept;
    [[nodiscard]] std::wstring_view DebugGetEmptyFolderFunMessage() const noexcept;

    struct DebugColumnLayoutEntry
    {
        size_t startIndex = 0;
        size_t itemCount  = 0;
        float leftDip     = 0.0f;
        float widthDip    = 0.0f;
        float rightDip    = 0.0f;
        std::wstring widestDisplayName;
        float widestObservedWidthDip = 0.0f;
    };

    struct DebugVisibleItemEntry
    {
        size_t index = 0;
        std::wstring displayName;
        D2D1_RECT_F bounds{};
        int column = 0;
        int row    = 0;
    };

    struct DebugColumnLayoutSnapshot
    {
        uint8_t displayMode                       = 0;
        int rowsPerColumn                         = 0;
        float horizontalOffsetDip                 = 0.0f;
        float maxHorizontalOffsetDip              = 0.0f;
        float contentWidthDip                     = 0.0f;
        float clientWidthDip                      = 0.0f;
        float clientHeightDip                     = 0.0f;
        float tileWidthDip                        = 0.0f;
        float tileHeightDip                       = 0.0f;
        size_t firstVisibleIndex                  = 0;
        size_t lastVisibleIndex                   = 0;
        uint32_t hitTestSpacingFalsePositiveCount = 0;
        std::vector<DebugColumnLayoutEntry> columns;
        std::vector<DebugVisibleItemEntry> visibleItems;
    };

    [[nodiscard]] bool DebugGetColumnLayoutSnapshot(DebugColumnLayoutSnapshot& out) const;

    struct DebugEmptyFolderItemMetrics
    {
        bool active           = false;
        float clientWidthDip  = 0.0f;
        float clientHeightDip = 0.0f;
        float tileWidthDip    = 0.0f;
        float tileHeightDip   = 0.0f;
        float labelHeightDip  = 0.0f;
    };
    [[nodiscard]] DebugEmptyFolderItemMetrics DebugGetEmptyFolderItemMetrics() const noexcept;

    enum class FilterWatermarkVisualMode : uint8_t
    {
        None,
        Background,
        Badge,
    };
    [[nodiscard]] FilterWatermarkVisualMode DebugGetFilterWatermarkVisualMode() const noexcept;
#endif

    enum class DisplayMode : uint8_t
    {
        Brief,
        Detailed,
        ExtraDetailed,
        Thumbnails,
    };

    enum class SortBy : uint8_t
    {
        Name,
        Extension,
        Time,
        Size,
        Attributes,
        None,
    };

    enum class SortDirection : uint8_t
    {
        Ascending,
        Descending,
    };

    enum class ErrorOverlayKind : uint8_t
    {
        Enumeration,
        Rendering,
        Operation,
    };

    enum class OverlaySeverity : uint8_t
    {
        Error,
        Warning,
        Information,
        Busy,
    };

#ifdef ENABLE_TESTS
    enum class DebugThumbnailProviderMode : uint8_t
    {
        Shell,
        ForceFallback,
        ForceSyntheticSuccess,
        ForceShellFailureAllowWic,
        ForceSyntheticWideSuccess,
    };

    struct AlertOverlayDebugSnapshot
    {
        bool visible             = false;
        ErrorOverlayKind kind    = ErrorOverlayKind::Operation;
        OverlaySeverity severity = OverlaySeverity::Error;
        std::wstring title;
        std::wstring message;
        HRESULT hr       = S_OK;
        bool closable    = true;
        bool blocksInput = true;
    };

    struct ThumbnailDebugSnapshot
    {
        bool visible                        = false;
        float targetDip                     = 16.0f;
        uint64_t queuedCount                = 0;
        uint64_t completedCount             = 0;
        uint64_t fallbackCount              = 0;
        uint64_t staleDropCount             = 0;
        uint64_t pendingCount               = 0;
        uint64_t cacheHitCount              = 0;
        uint64_t shellSuccessCount          = 0;
        uint64_t wicSuccessCount            = 0;
        uint64_t wicFactoryCreateCount      = 0;
        uint64_t decodeFailureCount         = 0;
        uint64_t visibleApplyCount          = 0;
        uint64_t visibleItemCount           = 0;
        uint64_t visibleThumbnailCount      = 0;
        uint64_t totalThumbnailCount        = 0;
        uint64_t cacheBytes                 = 0;
        uint64_t cacheEvictedCount          = 0;
        uint64_t cancelCount                = 0;
        bool lastDrawSawThumbnail           = false;
        uint32_t lastDrawSourceWidthPx      = 0;
        uint32_t lastDrawSourceHeightPx     = 0;
        D2D1_RECT_F lastDrawSlotRectDip     = D2D1::RectF();
        D2D1_RECT_F lastDrawRectDip         = D2D1::RectF();
        bool lastIconDrawSawIcon            = false;
        uint32_t lastIconDrawSourceWidthPx  = 0;
        uint32_t lastIconDrawSourceHeightPx = 0;
        D2D1_RECT_F lastIconDrawSlotRectDip = D2D1::RectF();
        D2D1_RECT_F lastIconDrawRectDip     = D2D1::RectF();
    };

    [[nodiscard]] bool DebugGetAlertOverlaySnapshot(AlertOverlayDebugSnapshot& out) const noexcept;
    [[nodiscard]] ThumbnailDebugSnapshot DebugGetThumbnailSnapshot() const noexcept;
    void DebugSetThumbnailProviderMode(DebugThumbnailProviderMode mode) noexcept;
#endif

    void SetDisplayMode(DisplayMode mode);
    [[maybe_unused]] DisplayMode GetDisplayMode() const noexcept
    {
        return _displayMode;
    }

    void SetFileExtensionsVisible(bool visible);
    [[nodiscard]] bool GetFileExtensionsVisible() const noexcept
    {
        return _fileExtensionsVisible;
    }

    void SetThumbnailsVisible(bool visible);
    [[nodiscard]] bool GetThumbnailsVisible() const noexcept
    {
        return _thumbnailsVisible;
    }

    void SetThumbnailSizeDip(uint32_t sizeDip);
    [[nodiscard]] uint32_t GetThumbnailSizeDip() const noexcept
    {
        return _thumbnailSizeDip;
    }

    void SetSort(SortBy sortBy, SortDirection direction);
    [[maybe_unused]] SortBy GetSortBy() const noexcept
    {
        return _sortBy;
    }
    [[maybe_unused]] SortDirection GetSortDirection() const noexcept
    {
        return _sortDirection;
    }

    void SetShowHiddenFiles(bool show);
    [[nodiscard]] bool GetShowHiddenFiles() const noexcept;
    void SetShowSystemFiles(bool show);
    [[nodiscard]] bool GetShowSystemFiles() const noexcept;

    struct NameFilterState
    {
        bool enabled = false;
        std::wstring text;
    };

    void SetNameFilterState(const NameFilterState& state, bool refresh = true);
    [[nodiscard]] NameFilterState GetNameFilterState() const;
    [[nodiscard]] bool IsNameFilterActive() const noexcept;

    void HideSelectedNames();
    void HideUnselectedNames();
    void ShowHiddenNames();
    [[nodiscard]] bool HasHiddenNames() const noexcept;

    enum class NavigationRequest
    {
        FocusNavigationMenu,
        FocusNavigationDiskInfo,
        FocusAddressBar,
        OpenHistoryDropdown,
        SwitchPane,
    };

    using NavigationRequestCallback = std::function<void(NavigationRequest request)>;

#pragma warning(push)
#pragma warning(disable : 4514) // Unreferenced inline function
    void SetNavigationRequestCallback(NavigationRequestCallback callback)
    {
        _navigationRequestCallback = std::move(callback);
    }
#pragma warning(pop)

    // Set callback for path changes (e.g., when user navigates to a subfolder)
#pragma warning(push)
#pragma warning(disable : 4514) // Unreferenced inline function
    void SetPathChangedCallback(std::function<void(const std::optional<std::filesystem::path>&)> callback)
    {
        _pathChangedCallback = std::move(callback);
    }
#pragma warning(pop)

    using NavigateUpFromRootRequestCallback = std::function<void()>;
#pragma warning(push)
#pragma warning(disable : 4514) // Unreferenced inline function
    void SetNavigateUpFromRootRequestCallback(NavigateUpFromRootRequestCallback callback)
    {
        _navigateUpFromRootRequestCallback = std::move(callback);
    }
#pragma warning(pop)

    using DirectoryImpactCallback = std::function<void(const DirectoryInfoCache::DirectoryImpact&)>;
    void SetDirectoryImpactCallback(DirectoryImpactCallback callback)
    {
        _directoryImpactCallback = std::move(callback);
    }

    using OpenFileRequestCallback = std::function<bool(const std::filesystem::path& path)>;
    void SetOpenFileRequestCallback(OpenFileRequestCallback callback)
    {
        _openFileRequestCallback = std::move(callback);
    }

    enum class ViewFileRole : uint8_t
    {
        Primary,
        Alternate,
    };

    struct ViewFileRequest
    {
        ViewFileRole role = ViewFileRole::Primary;
        HWND ownerWindow  = nullptr;
        std::wstring actionId;
        std::filesystem::path focusedPath;
        std::vector<std::filesystem::path> selectionPaths;
        std::vector<std::filesystem::path> displayedFilePaths;
    };

    using ViewFileRequestCallback = std::function<bool(const ViewFileRequest& request)>;
    void SetViewFileRequestCallback(ViewFileRequestCallback callback)
    {
        _viewFileRequestCallback = std::move(callback);
    }

    struct SelectionStats
    {
        uint32_t selectedFolders   = 0;
        uint32_t selectedFiles     = 0;
        uint64_t selectedFileBytes = 0;

        struct SelectedItemDetails
        {
            bool isDirectory      = false;
            uint64_t sizeBytes    = 0;
            int64_t lastWriteTime = 0;
            DWORD fileAttributes  = 0;
        };

        std::optional<SelectedItemDetails> singleItem;
    };

    using SelectionChangedCallback = std::function<void(const SelectionStats& stats)>;
    void SetSelectionChangedCallback(SelectionChangedCallback callback)
    {
        _selectionChangedCallback = std::move(callback);
    }
    [[nodiscard]] std::optional<SelectionStats::SelectedItemDetails> GetFocusedItemDetails() const noexcept;

    using FocusedItemChangedCallback = std::function<void()>;
    void SetFocusedItemChangedCallback(FocusedItemChangedCallback callback)
    {
        _focusedItemChangedCallback = std::move(callback);
    }

    using IncrementalSearchChangedCallback = std::function<void()>;
    void SetIncrementalSearchChangedCallback(IncrementalSearchChangedCallback callback)
    {
        _incrementalSearchChangedCallback = std::move(callback);
    }
    [[maybe_unused]] bool IsIncrementalSearchActive() const noexcept
    {
        return _incrementalSearch.active;
    }
    [[maybe_unused]] std::wstring_view GetIncrementalSearchQuery() const noexcept
    {
        return _incrementalSearch.query;
    }

#ifdef ENABLE_TESTS
    struct IncrementalSearchDebugMatch
    {
        std::wstring displayName;
        DWRITE_TEXT_RANGE range{};
        bool startsWith = false;
    };

    struct IncrementalSearchDebugSnapshot
    {
        bool active = false;
        std::wstring query;
        std::wstring focusedDisplayName;
        std::vector<IncrementalSearchDebugMatch> matches;
    };

    [[nodiscard]] bool DebugGetIncrementalSearchSnapshot(IncrementalSearchDebugSnapshot& out) const noexcept;
#endif

    using SelectionSizeComputationRequestedCallback = std::function<void()>;
    void SetSelectionSizeComputationRequestedCallback(SelectionSizeComputationRequestedCallback callback)
    {
        _selectionSizeComputationRequestedCallback = std::move(callback);
    }

    using EnumerationCompletedCallback = std::function<void(const std::filesystem::path& folder)>;
    void SetEnumerationCompletedCallback(EnumerationCompletedCallback callback)
    {
        _enumerationCompletedCallback = std::move(callback);
    }

    using DetailsTextProvider = std::function<std::wstring(
        const std::filesystem::path& folder, std::wstring_view displayName, bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes)>;

    void SetDetailsTextProvider(DetailsTextProvider provider)
    {
        _detailsTextProvider = std::move(provider);
    }

    using MetadataTextProvider = std::function<std::wstring(
        const std::filesystem::path& folder, std::wstring_view displayName, bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes)>;

    void SetMetadataTextProvider(MetadataTextProvider provider)
    {
        _metadataTextProvider = std::move(provider);
    }

    // Optional non-error empty-state message shown when enumeration succeeds but no items are displayed.
    // Caller controls when it is set/cleared (e.g. "This folder doesn't exist in this hierarchy.").
    void SetEmptyStateMessage(std::wstring message);

    // Optional non-interactive background watermark drawn behind items.
    // When `animated` is true, the caller is responsible for triggering periodic invalidation.
    void SetBackgroundWatermark(std::wstring message, bool animated);

    // Recompute `detailsText` for currently displayed items using the active DetailsTextProvider.
    // Useful when the provider's output depends on external state.
    void RefreshDetailsText();

    // Programmatic selection: sets selected=true for items where `shouldSelect(displayName)` returns true.
    // When clearExistingSelection is false, this only adds selection (it does not unselect already-selected items).
    void SetSelectionByDisplayNamePredicate(const std::function<bool(std::wstring_view)>& shouldSelect, bool clearExistingSelection = true);

    // Programmatic unselection: sets selected=false for items where `shouldUnselect(displayName)` returns true.
    void ClearSelectionByDisplayNamePredicate(const std::function<bool(std::wstring_view)>& shouldUnselect);

    void InvertSelection();

    void SelectSameName();
    void UnselectSameName();
    void SelectSameExtension();
    void UnselectSameExtension();

    // Selection navigation: move focus to the previous/next selected item without changing selection.
    // Returns true if focus changed (wrap allowed).
    [[nodiscard]] bool GoToPreviousSelectedName();
    [[nodiscard]] bool GoToNextSelectedName();

    void DebugShowOverlaySample(OverlaySeverity severity);
    void DebugShowOverlaySample(ErrorOverlayKind kind, OverlaySeverity severity, bool blocksInput);
    void DebugShowCanceledOverlaySample();
    void DebugHideOverlaySample();

    void ShowAlertOverlay(ErrorOverlayKind kind,
                          OverlaySeverity severity,
                          std::wstring title,
                          std::wstring message,
                          HRESULT hr       = S_OK,
                          bool closable    = true,
                          bool blocksInput = true);
    void DismissAlertOverlay();

private:
    class DropTarget;
    struct EnumerationPayload;

    [[nodiscard]] bool RequestViewFocusedItem(ViewFileRole role, bool activateFallback, std::wstring_view actionId = {});

#pragma warning(push)
// (C4625) copy constructor was implicitly defined as deleted / (C4626) assignment operator was implicitly defined as deleted
#pragma warning(disable : 4625 4626)

    struct IconBitmapRequest
    {
        uint64_t iconLoadBatchId       = 0;
        uint64_t enumerationGeneration = 0;
        int iconIndex                  = -1;
        float targetDipSize            = 16.0f;
        std::chrono::steady_clock::time_point postedAt{};
        std::vector<size_t> itemIndices;
        wil::unique_hicon hIcon = nullptr;
    };

    struct ThumbnailBitmapRequest
    {
        struct PixelBuffer
        {
            uint32_t widthPx     = 0;
            uint32_t heightPx    = 0;
            uint32_t strideBytes = 0;
            std::vector<uint8_t> bgra;

            [[nodiscard]] bool IsValid() const noexcept
            {
                return widthPx > 0u && heightPx > 0u && strideBytes >= widthPx * 4u && ! bgra.empty();
            }
        };

        enum class SourceKind : uint8_t
        {
            Shell,
            Wic,
            Synthetic,
            Fallback,
        };

        uint64_t thumbnailLoadBatchId  = 0;
        uint64_t enumerationGeneration = 0;
        size_t itemIndex               = static_cast<size_t>(-1);
        uint32_t targetPx              = 0;
        std::chrono::steady_clock::time_point postedAt{};
        wil::unique_hbitmap hBitmap = nullptr;
        PixelBuffer pixels;
        HRESULT hr            = S_OK;
        bool usedFallback     = false;
        SourceKind sourceKind = SourceKind::Fallback;
    };

    struct FolderItem
    {
        // Zero-copy displayName: points into arena buffer (IFilesInformation kept alive)
        std::wstring_view displayName; // View into FileInfo::FileName in arena
        uint16_t extensionOffset = 0;  // Offset to '.' in displayName (0 if none/directory)
        uint32_t stableHash32    = 0;  // Stable hash (used for rainbow rendering, etc.)

        bool isDirectory      = false;
        bool selected         = false;
        bool focused          = false;
        bool isShortcut       = false; // True if .lnk file requiring overlay rendering
        uint64_t sizeBytes    = 0;
        int64_t lastWriteTime = 0;
        DWORD fileAttributes  = 0;
        size_t unsortedOrder  = 0;

        // Rendering state
        D2D1_RECT_F bounds{};
        wil::com_ptr<ID2D1Bitmap1> icon;
        wil::com_ptr<ID2D1Bitmap1> thumbnail;
        bool thumbnailFallbackResolved = false;
        uint32_t thumbnailFallbackTargetPx = 0;
        int iconIndex = -1; // System image list icon index from SHGetFileInfo
        int column    = 0;
        int row       = 0;
        wil::com_ptr<IDWriteTextLayout> labelLayout;
        DWRITE_TEXT_METRICS labelMetrics{};
        std::wstring detailsText;
        wil::com_ptr<IDWriteTextLayout> detailsLayout;
        DWRITE_TEXT_METRICS detailsMetrics{};

        std::wstring metadataText;
        wil::com_ptr<IDWriteTextLayout> metadataLayout;
        DWRITE_TEXT_METRICS metadataMetrics{};

        // Get extension from displayName (zero-copy)
        [[nodiscard]] std::wstring_view GetExtension() const noexcept
        {
            return extensionOffset > 0 ? displayName.substr(extensionOffset) : std::wstring_view{};
        }

        [[nodiscard]] std::wstring_view GetNameWithoutExtension() const noexcept
        {
            return extensionOffset > 0 ? displayName.substr(0, extensionOffset) : displayName;
        }
    };

    struct DragContext
    {
        bool dragging = false;
        POINT startPoint{};
        size_t anchorIndex    = static_cast<size_t>(-1);
        bool hasStartItemRect = false;
        RECT startItemRect{};
    };

    struct IncrementalSearchState
    {
        bool active                = false;
        bool suppressNextSpaceChar = false;
        std::wstring query;
        size_t highlightedIndex = static_cast<size_t>(-1);
        DWRITE_TEXT_RANGE highlightedRange{};
    };
#pragma warning(pop)

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _hParent;
    float _dpi = 96.0f;
    SIZE _clientSize{0, 0};
    std::optional<std::filesystem::path> _currentFolder;
    std::optional<std::filesystem::path> _displayedFolder;
    std::wstring _emptyStateMessage;

    enum class EmptyStateMessageKind : uint8_t
    {
        None,
        Generic,
        CompareNoDifferences,
    };
    EmptyStateMessageKind _emptyStateMessageKind = EmptyStateMessageKind::None;

    wil::com_ptr<IDWriteTextLayout> _emptyMessageIconLayout;
    wil::com_ptr<IDWriteTextLayout> _emptyMessageTitleLayout;
    wil::com_ptr<IDWriteTextLayout> _emptyMessageFunLayout;
    SIZE _emptyMessageLayoutClientSizePx = {};
    float _emptyMessageLayoutDpi         = 0.0f;
    UINT _emptyMessageLayoutMessageId    = 0;
    float _emptyMessageIconFontSizeDip   = 0.0f;
    DWRITE_TEXT_METRICS _emptyMessageIconMetrics{};
    DWRITE_TEXT_METRICS _emptyMessageTitleMetrics{};
    DWRITE_TEXT_METRICS _emptyMessageFunMetrics{};

    struct CompareNoDifferencesState
    {
        std::wstring folderKey;
        UINT funMessageResourceId = 0;
        std::wstring emoji;
        std::wstring funMessage;
    };
    std::optional<CompareNoDifferencesState> _compareNoDifferencesState;

    std::wstring _backgroundWatermarkMessage;
    bool _backgroundWatermarkAnimated = false;

    struct EmptyFolderState
    {
        std::wstring folderKey;
        UINT funMessageResourceId = 0;
        std::wstring emoji;
        std::wstring funMessage;
    };

    std::optional<EmptyFolderState> _emptyFolderState;
    wil::com_ptr<IDWriteTextLayout> _emptyFolderFocusCueLayout;
    float _emptyFolderFocusCueLayoutWidthDip  = 0.0f;
    float _emptyFolderFocusCueLayoutHeightDip = 0.0f;
    float _emptyFolderFocusCueLayoutDpi       = 0.0f;
    float _emptyFolderFocusCueFontSizeDip     = 0.0f;
    wil::com_ptr<IDWriteTextLayout> _emptyFolderIconLayout;
    wil::com_ptr<IDWriteTextLayout> _emptyFolderTitleLayout;
    wil::com_ptr<IDWriteTextLayout> _emptyFolderFunLayout;
    SIZE _emptyFolderLayoutClientSizePx = {};
    float _emptyFolderLayoutDpi         = 0.0f;
    UINT _emptyFolderLayoutMessageId    = 0;
    float _emptyFolderIconFontSizeDip   = 0.0f;
    DWRITE_TEXT_METRICS _emptyFolderIconMetrics{};
    DWRITE_TEXT_METRICS _emptyFolderTitleMetrics{};
    DWRITE_TEXT_METRICS _emptyFolderFunMetrics{};
    wil::com_ptr<ID2D1SolidColorBrush> _emptyFolderFocusCueBrush;
    wil::com_ptr<IFileSystem> _fileSystem;
    const PluginMetaData* _fileSystemMetadata{};
    std::wstring _fileSystemPluginId;
    std::wstring _fileSystemInstanceContext;
    DirectoryInfoCache::Pin _directoryCachePin;

    std::wstring _focusMemoryRootKey;
    std::unordered_map<std::wstring, std::wstring> _focusMemory;

    std::vector<FolderItem> _items;
    wil::com_ptr<IFilesInformation> _itemsArenaBuffer; // Keeps arena alive for zero-copy string_views
    std::filesystem::path _itemsFolder;                // Folder path for computing full paths

    size_t _focusedIndex = static_cast<size_t>(-1);
    size_t _hoveredIndex = static_cast<size_t>(-1);
    size_t _anchorIndex  = static_cast<size_t>(-1);
    int _columns         = 1;
    int _rowsPerColumn   = 0;
    std::vector<FolderViewColumnLayout::Column> _columnLayout;
    std::vector<int> _columnCounts;
    std::vector<size_t> _columnPrefixSums; // Prefix sums for O(1) hit testing: _columnPrefixSums[c] = sum of _columnCounts[0..c-1]
    float _scrollOffset     = 0.0f;
    float _horizontalOffset = 0.0f;
    struct PendingInputToPaintMetric final
    {
        std::chrono::steady_clock::time_point startedAt{};
        std::wstring metricName;
        std::wstring detail;
        uint64_t value0 = 0u;
        uint64_t value1 = 0u;
    };
    std::optional<PendingInputToPaintMetric> _pendingInputToPaintMetric;
    float _contentHeight = 0.0f;
    float _contentWidth  = 0.0f;

    // Scroll direction tracking for predictive layout pre-loading
    float _lastScrollOffset     = 0.0f;
    float _lastHorizontalOffset = 0.0f;
    int8_t _scrollDirectionY    = 0; // -1 = up, 0 = none, 1 = down
    int8_t _scrollDirectionX    = 0; // -1 = left, 0 = none, 1 = right

    // Idle-time layout pre-creation for off-screen items
    size_t _idleLayoutNextIndex                             = 0; // Next item to process during idle
    UINT_PTR _idleLayoutTimer                               = 0; // Timer for idle processing
    static constexpr UINT_PTR kIdleLayoutTimerId            = 2;
    UINT_PTR _directoryCacheRefreshTimer                    = 0;
    static constexpr UINT_PTR kDirectoryCacheRefreshTimerId = 3;
    static constexpr UINT kIdleLayoutIntervalMs             = 16; // ~60fps idle processing
    static constexpr size_t kIdleLayoutBatchSize            = 20; // Items per idle batch
    DragContext _drag{};
    bool _swapChainResizePending = false;
    UINT _pendingSwapChainWidth  = 0;
    UINT _pendingSwapChainHeight = 0;
    bool _deferredInitPosted     = false;

    // Rendering resources
    AppTheme _appTheme;
    FolderViewTheme _theme;
    MenuTheme _menuTheme;
    const ShortcutManager* _shortcutManager = nullptr;
    wil::unique_hbrush _menuBackgroundBrush;

    struct EnumerationPayload
    {
        uint64_t generation = 0;
        HRESULT status      = S_OK;
        std::vector<FolderItem> items;

        // Zero-copy: keep arena buffer alive so string_views in items remain valid
        wil::com_ptr<IFilesInformation> arenaBuffer;
        std::filesystem::path folder; // Needed to compute full paths on demand
    };

    wil::com_ptr<ID3D11Device> _d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> _d3dContext;
    wil::com_ptr<IDXGISwapChain1> _swapChain;
    wil::com_ptr<IDXGISwapChain> _swapChainLegacy;
    wil::com_ptr<ID2D1Factory1> _d2dFactory;
    mutable std::mutex _d2dDeviceMutex; // Guards cross-thread reads/writes of _d2dDevice (UI thread + enumeration worker).
    wil::com_ptr<ID2D1Device> _d2dDevice;
    wil::com_ptr<ID2D1DeviceContext> _d2dContext;
    wil::com_ptr<ID2D1Bitmap1> _d2dTarget;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<IDWriteTextFormat> _labelFormat;
    wil::com_ptr<IDWriteTextFormat> _detailsFormat;
    wil::com_ptr<IDWriteTextFormat> _filterWatermarkFormat;
    wil::com_ptr<IDWriteTextLayout> _filterWatermarkLayout;
    SIZE _filterWatermarkLayoutClientSizePx = {};
    float _filterWatermarkLayoutDpi         = 0.0f;
    float _filterWatermarkLayoutFontSizeDip = 0.0f;
    wil::com_ptr<IDWriteTextLayout> _filterWatermarkBadgeLayout;
    SIZE _filterWatermarkBadgeLayoutClientSizePx = {};
    float _filterWatermarkBadgeLayoutDpi         = 0.0f;
    float _filterWatermarkBadgeLayoutFontSizeDip = 0.0f;
    wil::com_ptr<IDWriteTextLayout> _backgroundWatermarkLayout;
    SIZE _backgroundWatermarkLayoutClientSizePx = {};
    float _backgroundWatermarkLayoutDpi         = 0.0f;
    std::wstring _backgroundWatermarkLayoutText;
    float _backgroundWatermarkLayoutFontSizeDip = 0.0f;
    wil::com_ptr<IDWriteInlineObject> _ellipsisSign;
    wil::com_ptr<IDWriteInlineObject> _detailsEllipsisSign;
    wil::com_ptr<ID2D1SolidColorBrush> _backgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _filterWatermarkBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _backgroundWatermarkBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textUnfocusedBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _detailsTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _detailsTextUnfocusedBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _metadataTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _metadataTextUnfocusedBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _selectionBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _focusedBackgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _focusBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchHighlightBrush;
    wil::com_ptr<ID2D1Bitmap> _placeholderFolderIcon; // Folder placeholder (48×48) with Fluent Design
    wil::com_ptr<ID2D1Bitmap> _placeholderFileIcon;   // File placeholder (48×48) with Fluent Design
    wil::com_ptr<ID2D1Bitmap> _shortcutOverlayIcon;   // 16×16 shortcut arrow overlay
    wil::com_ptr<IWICImagingFactory> _wicFactory;
    wil::com_ptr<IDropTarget> _dropTarget;
    std::unique_ptr<RedSalamander::Ui::AlertOverlay> _alertOverlay;

    struct ErrorOverlayState
    {
        ErrorOverlayKind kind    = ErrorOverlayKind::Operation;
        OverlaySeverity severity = OverlaySeverity::Error;
        std::wstring title;
        std::wstring message;
        HRESULT hr         = S_OK;
        uint64_t startTick = 0;
        bool closable      = true;
        bool blocksInput   = true;
    };

    mutable std::mutex _errorOverlayMutex;
    mutable std::optional<ErrorOverlayState> _errorOverlay;
    mutable uint64_t _overlayAnimationSubscriptionId         = 0;
    mutable UINT_PTR _overlayTimer                           = 0;
    mutable UINT _overlayTimerIntervalMs                     = 0;
    mutable uint64_t _operationInfoOverlayAutoDismissDueTick = 0;
    mutable std::optional<RECT> _lastOverlayInvalidationRectPx;

    struct PendingBusyOverlay
    {
        uint64_t generation = 0;
        std::filesystem::path folder;
        uint64_t startTick = 0;
    };

    std::optional<PendingBusyOverlay> _pendingBusyOverlay;

    D3D_FEATURE_LEVEL _featureLevel = D3D_FEATURE_LEVEL_11_0;
    bool _coInitialized             = false;
    bool _oleInitialized            = false;
    bool _dropTargetRegistered      = false;
    bool _supportsPresent1          = true;
    bool _paneFocused               = false;
    std::atomic<bool> _showHiddenFiles{true};
    std::atomic<bool> _showSystemFiles{true};

    struct CompiledNameFilter final
    {
        NameFilterState state{};
        MaskSyntax::WildcardMask mask{};
        bool hasMask = false;
    };

    std::atomic<std::shared_ptr<const CompiledNameFilter>> _nameFilter;

    struct HiddenNamesHash final
    {
        using is_transparent = void;

        size_t operator()(std::wstring_view value) const noexcept
        {
            uint64_t hash = 14695981039346656037ull; // FNV-1a 64-bit offset basis
            for (const wchar_t ch : value)
            {
                const wchar_t lower = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
                hash ^= static_cast<uint64_t>(lower);
                hash *= 1099511628211ull; // FNV-1a 64-bit prime
            }
            return static_cast<size_t>(hash);
        }

        size_t operator()(const std::wstring& value) const noexcept
        {
            return (*this)(std::wstring_view{value});
        }
    };

    struct HiddenNamesEq final
    {
        using is_transparent = void;

        bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
        {
            if (left.size() != right.size())
            {
                return false;
            }

            for (size_t i = 0; i < left.size(); ++i)
            {
                const wint_t a = std::towlower(static_cast<wint_t>(left[i]));
                const wint_t b = std::towlower(static_cast<wint_t>(right[i]));
                if (a != b)
                {
                    return false;
                }
            }

            return true;
        }
    };

    struct HiddenNamesFilter final
    {
        std::unordered_set<std::wstring, HiddenNamesHash, HiddenNamesEq> names;
    };

    std::atomic<std::shared_ptr<const HiddenNamesFilter>> _hiddenNames;

    IncrementalSearchState _incrementalSearch{};
    bool _incrementalSearchLayoutEffectsDirty = false;
    std::wstring _incrementalSearchIndicatorDisplayQuery;
    mutable float _incrementalSearchIndicatorVisibility          = 0.0f;
    mutable float _incrementalSearchIndicatorVisibilityFrom      = 0.0f;
    mutable float _incrementalSearchIndicatorVisibilityTo        = 0.0f;
    mutable uint64_t _incrementalSearchIndicatorVisibilityStart  = 0;
    mutable uint64_t _incrementalSearchIndicatorTypingPulseStart = 0;
    std::wstring _incrementalSearchIndicatorLayoutText;
    float _incrementalSearchIndicatorLayoutMaxWidthDip = 0.0f;
    wil::com_ptr<IDWriteTextLayout> _incrementalSearchIndicatorLayout;
    DWRITE_TEXT_METRICS _incrementalSearchIndicatorLayoutMetrics{};
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchIndicatorBackgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchIndicatorBorderBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchIndicatorTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchIndicatorShadowBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _incrementalSearchIndicatorAccentBrush;
    wil::com_ptr<ID2D1StrokeStyle> _incrementalSearchIndicatorStrokeStyle;

    // Rendering constants (logical DIPs)
    float _tileWidthDip          = 220.0f;
    float _tileHeightDip         = 32.0f;
    float _iconSizeDip           = 16.0f; // Match Windows Explorer list mode (SHIL_SMALL)
    uint32_t _thumbnailSizeDip   = Common::Settings::Thumbnail::kDefaultSizeDip;
    float _tileSpacingDip        = 16.0f;
    float _labelHeightDip        = 20.0f;
    float _detailsLineHeightDip  = 0.0f;
    float _metadataLineHeightDip = 0.0f;

    // Global cache for label measurements
    bool _itemMetricsCached       = false;
    float _cachedMaxLabelWidth    = 0.0f;
    float _cachedMaxLabelHeight   = 0.0f;
    float _cachedMaxDetailsWidth  = 0.0f;
    float _cachedMaxMetadataWidth = 0.0f;
    float _lastLayoutWidth        = 0.0f;
    size_t _detailsSizeSlotChars  = 0;

    // Estimated metrics for lazy layout creation (avoids measuring all items upfront)
    // These are computed from actual font metrics in UpdateEstimatedMetrics()
    float _estimatedCharWidthDip      = 7.0f;  // Average character width, updated from font metrics
    float _estimatedLabelHeightDip    = 16.0f; // Label height, updated from font metrics
    float _estimatedDetailsHeightDip  = 14.0f; // Details line height, updated from font metrics
    float _estimatedMetadataHeightDip = 14.0f; // Metadata line height, updated from font metrics
    bool _estimatedMetricsValid       = false; // True after metrics computed from font

    DisplayMode _displayMode     = DisplayMode::Brief;
    SortBy _sortBy               = SortBy::Name;
    SortDirection _sortDirection = SortDirection::Ascending;
    bool _fileExtensionsVisible  = true;
    bool _thumbnailsVisible      = false;

    // Callback for path changes
    std::function<void(const std::optional<std::filesystem::path>&)> _pathChangedCallback;
    NavigateUpFromRootRequestCallback _navigateUpFromRootRequestCallback;
    DirectoryImpactCallback _directoryImpactCallback;
    OpenFileRequestCallback _openFileRequestCallback;
    ViewFileRequestCallback _viewFileRequestCallback;
    FileOperationRequestCallback _fileOperationRequestCallback;
    PropertiesRequestCallback _propertiesRequestCallback;
    NavigationRequestCallback _navigationRequestCallback;
    SelectionChangedCallback _selectionChangedCallback;
    FocusedItemChangedCallback _focusedItemChangedCallback;
    IncrementalSearchChangedCallback _incrementalSearchChangedCallback;
    SelectionSizeComputationRequestedCallback _selectionSizeComputationRequestedCallback;
    EnumerationCompletedCallback _enumerationCompletedCallback;
    DetailsTextProvider _detailsTextProvider;
    MetadataTextProvider _metadataTextProvider;

    SelectionStats _selectionStats{};

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

    void OnCreate();
    void OnDestroy();
    void OnSize(UINT width, UINT height);
    void OnPaint();
    void OnDeferredInit();
    void OnMouseWheel(int delta, bool horizontal);
    void OnMouseWheelMessage(UINT keyState, int delta);
    void OnMeasuredMouseWheelMessage(UINT keyState, int delta);
    void OnMeasuredMouseHWheelMessage(int delta);
    void OnLButtonDown(POINT pt, WPARAM keys);
    void OnLButtonDblClk(POINT pt, WPARAM keys);
    void OnLButtonUp(POINT pt);
    void OnMouseMove(POINT pt, WPARAM keys);
    void OnMouseLeave();
    void OnKeyDown(WPARAM key, bool ctrl, bool shift, bool translatedSpaceCharMayFollow);
    void OnKeyDownMessage(WPARAM key, LPARAM keyInfo);
    void OnMeasuredKeyDownMessage(WPARAM key, LPARAM keyInfo);
    void OnCharMessage(wchar_t character);
    bool OnSysKeyDownMessage(WPARAM key, LPARAM keyInfo);
    bool OnMeasuredSysKeyDownMessage(WPARAM key, LPARAM keyInfo);
    LRESULT OnSetFocusMessage() noexcept;
    LRESULT OnKillFocusMessage() noexcept;
    void OnContextMenuMessage(HWND hwnd, LPARAM lParam);
    void OnHScrollMessage(UINT scrollRequest);
    void OnMeasuredHScrollMessage(UINT scrollRequest);
    void OnCommandMessage(UINT commandId);
    void OnContextMenu(POINT screenPt);

    void EnsureDeviceIndependentResources();
    void EnsureDeviceResources();
    void EnsureSwapChain();
    void PrepareForSwapChainChange();
    bool TryResizeSwapChain(UINT width, UINT height);
    void ReleaseSwapChain();
    void DiscardDeviceResources();
    void CreatePlaceholderIcon();
    void RecreateThemeBrushes();
    void DrawErrorOverlay();
    void ClearErrorOverlay(ErrorOverlayKind kind) const;
    void OnTimerMessage(UINT_PTR timerId);
    void StartOverlayTimer(UINT intervalMs) const;
    void StopOverlayTimer() const;
    void StartOverlayAnimation() const noexcept;
    void StopOverlayAnimation() const noexcept;
    RECT GetIncrementalSearchIndicatorInvalidationRectPx() const noexcept;
    void InvalidateOverlayVisuals(bool forceFullClient) const noexcept;
    bool OnOverlayAnimationTick(uint64_t nowTickMs) const noexcept;
    bool UpdateIncrementalSearchIndicatorAnimation(uint64_t nowTickMs) const noexcept;
    void ScheduleBusyOverlay(uint64_t generation, const std::filesystem::path& folder);
    void CancelBusyOverlay(uint64_t generation);
    void ShowBusyOverlayNow(const std::filesystem::path& folder);

    void EnumerateFolder();
    void NavigateUp() noexcept;
    [[nodiscard]] bool CanShowEmptyFolderState() const noexcept;
    void UpdateCompareNoDifferencesState() noexcept;
    void EnsureEnumerationThread();
    void EnumerationWorker(std::stop_token stopToken);
    void DrainPendingEnumerationPayloadMessages() noexcept;
    std::unique_ptr<EnumerationPayload> ExecuteEnumeration(const std::filesystem::path& folder, uint64_t generation, std::stop_token stopToken);
    void ApplyCurrentSort();
    void ApplyCurrentSort(std::wstring_view focusedPath, size_t fallbackFocusIndex);
    void LayoutItems();
    void UpdateScrollMetrics();
    void Render(const RECT& invalidRect);
    struct DrawItemPerfStats
    {
        uint64_t itemHasIcon              = 0;
        uint64_t itemPlaceholderIcon      = 0;
        uint64_t itemTextLayoutLabel      = 0;
        uint64_t itemTextLayoutDetails    = 0;
        uint64_t itemTextLayoutMetadata   = 0;
        uint64_t incrementalSearchUpdates = 0;
    };
    void DrawItem(FolderItem& item, DrawItemPerfStats* perfStats = nullptr);
    void DrawIncrementalSearchIndicator(uint64_t nowTickMs);

    void SelectSingle(size_t index);
    void ToggleSelection(size_t index);
    void RangeSelect(size_t index);
    void ClearSelection();
    void SelectAll();
    void RecomputeSelectionStats() noexcept;
    void NotifySelectionChanged() const noexcept;
    void NotifyFocusedItemChanged() const noexcept;
    void NotifyIncrementalSearchChanged() const noexcept;
    void FocusItem(size_t index, bool ensureVisible);
    [[nodiscard]] bool GoToSelectedName(bool forward);
    void ActivateFocusedItem();
    void ExitIncrementalSearch() noexcept;
    void UpdateIncrementalSearchIndicatorState(uint64_t nowTickMs, bool triggerPulse, std::wstring_view displayQuery) noexcept;
    void HandleIncrementalSearchBackspace();
    void HandleIncrementalSearchNavigate(bool forward);
    void UpdateIncrementalSearchHighlightForFocusedItem();
    void ClearIncrementalSearchHighlight() noexcept;
    void ClearIncrementalSearchLayoutEffects() noexcept;
    void ApplyIncrementalSearchHighlight(size_t itemIndex, const DWRITE_TEXT_RANGE& range) noexcept;
    std::optional<UINT32> FindIncrementalSearchMatchOffset(std::wstring_view displayName) const noexcept;
    std::optional<size_t> FindNextIncrementalSearchPrefixMatch(size_t startIndex, bool forward) const noexcept;
    std::wstring_view GetIncrementalSearchDisplayName(size_t index) const noexcept;
    void DeleteSelectedItems();
    void CopySelectionToClipboard();
    void PasteItemsFromClipboard();
    void RenameFocusedItem();
    void ShowProperties();
    void MoveSelectedItems();
    void EnsureDropTarget();
    void BeginDragDrop();
    void UpdateItemTextLayouts();
    void EnsureItemTextLayout(FolderItem& item, float labelWidth);
    [[nodiscard]] float GetItemTextLayoutWidth(const FolderItem& item) const noexcept;
    [[nodiscard]] std::wstring_view GetVisualDisplayName(const FolderItem& item) const noexcept;
    std::pair<size_t, size_t> GetVisibleItemRange() const;
    void ReleaseDistantRenderingState(); // Release layouts/icons for items far from visible range
    void ScheduleIdleLayoutCreation();
    void ProcessIdleLayoutBatch();
    void UpdateEstimatedMetrics();
    void UpdateContextMenuState(HMENU menu) const;
    DWORD ResolveDropEffect(DWORD keyState, DWORD allowedEffects) const;
    bool HasFileDrop(IDataObject* dataObject) const;
    HRESULT PerformDrop(IDataObject* dataObject, DWORD keyState, DWORD allowedEffects, DWORD* performedEffect);

    void ReportError(const std::wstring& context, HRESULT hr) const;
    bool CheckHR(HRESULT hr, const wchar_t* context) const;
    void QueueIconLoading();
    void ProcessIconLoadQueue();
    [[nodiscard]] bool HasMissingVisibleThumbnails() const;
    void QueueMissingVisibleThumbnails();
    void QueueThumbnailLoading();
    void CancelThumbnailLoading() noexcept;
    void ProcessThumbnailLoadQueue();
    [[nodiscard]] HRESULT EnsureThumbnailWicFactory(wil::com_ptr<IWICImagingFactory>& thumbnailWicFactory, IWICImagingFactory** outFactory) noexcept;

    void OnIconLoaded(size_t itemIndex);
    void OnBatchIconUpdate();
    void OnDirectoryCacheDirty();
    void OnDirectoryImpact(std::unique_ptr<DirectoryInfoCache::DirectoryImpact> impact);
    void RequestRefreshFromCache();
    D2D1_RECT_F OffsetRect(const D2D1_RECT_F& rect, float dx, float dy) const;
    static RECT ToPixelRect(const D2D1_RECT_F& rect, float dpi);
    static bool RectIntersects(const D2D1_RECT_F& rect, const RECT& pixelRect, float dpi);
    void RecordInputToPaintStartIfViewportOrFocusChanged(std::chrono::steady_clock::time_point inputStartedAt,
                                                         float scrollBefore,
                                                         float horizontalBefore,
                                                         size_t focusedBefore) noexcept;
    void EmitPendingInputToPaintMetricAfterPresent() noexcept;

    std::optional<size_t> HitTest(POINT clientPt) const;
    POINT ScreenToClientPoint(POINT screenPt) const;
    void EnsureVisible(size_t index);
    void ProcessEnumerationResult(std::unique_ptr<EnumerationPayload> payload);
    void RememberFocusedItemForDisplayedFolder() noexcept;
    void EnsureFocusMemoryRootForFolder(const std::filesystem::path& folder) noexcept;
    [[nodiscard]] std::wstring GetRememberedFocusedItemPathForFolder(const std::filesystem::path& folder) noexcept;
    void StopEnumerationThread() noexcept;
    [[nodiscard]] std::filesystem::path GetItemFullPath(const FolderItem& item) const;
    [[maybe_unused]] float DipFromPx(int px) const
    {
        return static_cast<float>(px) * 96.0f / _dpi;
    }
    [[maybe_unused]] int PxFromDip(float dip) const
    {
        return static_cast<int>(std::lround(dip * _dpi / 96.0f));
    }

    // Background enumeration thread (also handles async icon loading)
    bool _enumerationThreadStarted = false;
    std::mutex _enumerationMutex;
    std::condition_variable _enumerationCv;
    std::optional<std::filesystem::path> _pendingEnumerationPath;
    uint64_t _pendingEnumerationGeneration = 0;
    std::atomic<uint64_t> _enumerationGeneration{0};
    ULONGLONG _lastDirectoryCacheRefreshTick = 0;
#ifdef ENABLE_TESTS
    uint64_t _debugForceRefreshCount                  = 0;
    uint64_t _debugWarmRenderingCallCount             = 0;
    uint64_t _debugDeferredInitCallCount              = 0;
    uint64_t _debugRenderCallCount                    = 0;
    uint64_t _debugQueueIconLoadingCallCount          = 0;
    uint64_t _debugProcessIconQueueCallCount          = 0;
    uint64_t _debugBatchIconUpdateCallCount           = 0;
    uint64_t _debugIncrementalSearchEffectUpdateCount = 0;
#endif

    struct PendingExternalCommand final
    {
        UINT commandId      = 0;
        uint64_t generation = 0;
        std::filesystem::path targetFolder;
        std::wstring expectedFocusDisplayName;
    };
    std::optional<PendingExternalCommand> _pendingExternalCommandAfterEnumeration;

    struct PendingRefreshSelectionRename final
    {
        std::wstring fromDisplayName;
        std::wstring toDisplayName;
        bool fromWasSelected = false;
    };
    std::vector<PendingRefreshSelectionRename> _pendingRefreshSelectionRenames;

    // Icon loading queue (grouped by icon index) - deque for O(1) front removal.
    // Each request is "convert this icon once, then apply to N items".
    struct IconLoadRequest
    {
        uint64_t enumerationGeneration = 0;
        int iconIndex                  = -1;
        float targetDipSize            = 16.0f;
        bool hasVisibleItems           = false;
        size_t firstVisibleItemIndex   = static_cast<size_t>(-1);
        unsigned int retryCount        = 0;
        std::chrono::steady_clock::time_point enqueuedAt{};
        std::vector<size_t> itemIndices;
    };
    std::deque<IconLoadRequest> _iconLoadQueue;
    std::atomic<bool> _iconLoadingActive{false};

    struct ThumbnailLoadRequest
    {
        uint64_t thumbnailLoadBatchId  = 0;
        uint64_t enumerationGeneration = 0;
        size_t itemIndex               = static_cast<size_t>(-1);
        std::filesystem::path fullPath;
        uint32_t targetPx       = 0;
        bool hasVisibleItem     = false;
        unsigned int retryCount = 0;
        std::chrono::steady_clock::time_point enqueuedAt{};
    };
    std::deque<ThumbnailLoadRequest> _thumbnailLoadQueue;
    std::atomic<bool> _thumbnailLoadingActive{false};

    // Icon loading performance telemetry
    struct IconLoadStats
    {
        std::atomic<uint64_t> totalRequests{0};
        std::atomic<uint64_t> visibleRequests{0};
        std::atomic<uint64_t> cacheHits{0};
        std::atomic<uint64_t> uniqueIconsQueued{0};
        std::atomic<uint64_t> extracted{0};

        std::atomic<uint64_t> bitmapPosted{0};
        std::atomic<uint64_t> bitmapPostFailed{0};
        std::atomic<uint64_t> bitmapConverted{0};
        std::atomic<uint64_t> bitmapConvertFailed{0};
        std::atomic<uint64_t> bitmapConvertUsTotal{0};
        std::atomic<uint64_t> bitmapConvertUsMax{0};
        std::atomic<uint64_t> pendingBitmapCreates{0};
        std::atomic<int64_t> bitmapFirstPostQpc{0};

        std::atomic<uint64_t> batchId{0};
        std::atomic<bool> bitmapSummaryEmitted{false};
        LARGE_INTEGER startTime{};

        IconLoadStats() noexcept                       = default;
        IconLoadStats(const IconLoadStats&)            = delete;
        IconLoadStats& operator=(const IconLoadStats&) = delete;
        IconLoadStats(IconLoadStats&&)                 = delete;
        IconLoadStats& operator=(IconLoadStats&&)      = delete;
    };
    IconLoadStats _iconLoadStats;

    struct ThumbnailLoadStats
    {
        std::atomic<uint64_t> queued{0};
        std::atomic<uint64_t> completed{0};
        std::atomic<uint64_t> fallback{0};
        std::atomic<uint64_t> staleDrops{0};
        std::atomic<uint64_t> pendingBitmapCreates{0};
        std::atomic<uint64_t> cacheHits{0};
        std::atomic<uint64_t> shellSuccess{0};
        std::atomic<uint64_t> wicSuccess{0};
        std::atomic<uint64_t> wicFactoryCreate{0};
        std::atomic<uint64_t> decodeFailures{0};
        std::atomic<uint64_t> visibleApply{0};
        std::atomic<uint64_t> cacheBytes{0};
        std::atomic<uint64_t> cacheEvicted{0};
        std::atomic<uint64_t> cancelCount{0};
        std::atomic<uint64_t> batchId{0};
        std::atomic<uint64_t> targetDipX100{1600};

        ThumbnailLoadStats() noexcept                            = default;
        ThumbnailLoadStats(const ThumbnailLoadStats&)            = delete;
        ThumbnailLoadStats& operator=(const ThumbnailLoadStats&) = delete;
        ThumbnailLoadStats(ThumbnailLoadStats&&)                 = delete;
        ThumbnailLoadStats& operator=(ThumbnailLoadStats&&)      = delete;
    };
    ThumbnailLoadStats _thumbnailLoadStats;
#ifdef ENABLE_TESTS
    std::atomic<DebugThumbnailProviderMode> _debugThumbnailProviderMode{DebugThumbnailProviderMode::Shell};
    bool _debugLastThumbnailDrawSawThumbnail   = false;
    uint32_t _debugLastThumbnailSourceWidthPx  = 0;
    uint32_t _debugLastThumbnailSourceHeightPx = 0;
    D2D1_RECT_F _debugLastThumbnailSlotRectDip = D2D1::RectF();
    D2D1_RECT_F _debugLastThumbnailDrawRectDip = D2D1::RectF();
    bool _debugLastIconDrawSawIcon             = false;
    uint32_t _debugLastIconDrawSourceWidthPx   = 0;
    uint32_t _debugLastIconDrawSourceHeightPx  = 0;
    D2D1_RECT_F _debugLastIconDrawSlotRectDip  = D2D1::RectF();
    D2D1_RECT_F _debugLastIconDrawRectDip      = D2D1::RectF();
#endif
    std::jthread _enumerationThread;

    // Helper methods for UI thread bitmap creation
    void OnCreateIconBitmap(std::unique_ptr<IconBitmapRequest> request);
    void OnCreateThumbnailBitmap(std::unique_ptr<ThumbnailBitmapRequest> request);
    void MaybeEmitIconBitmapSummary(uint64_t batchId) noexcept;
    void BoostIconLoadingForVisibleRange();
};

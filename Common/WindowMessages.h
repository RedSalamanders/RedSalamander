#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <cstddef>
#include <cstdint>
#include <windows.h>

namespace WndMsg
{
// Central registry for custom Win32 messages.
// Keep *all* WM_APP/WM_USER message IDs in this file so they stay unique and easy to audit.

inline constexpr UINT kFolderViewSyncSwapChain         = WM_APP + 0x300;
inline constexpr UINT kFolderViewEnumerateComplete     = WM_APP + 0x301;
inline constexpr UINT kFolderViewIconLoaded            = WM_APP + 0x302;
inline constexpr UINT kFolderViewCreateIconBitmap      = WM_APP + 0x303;
inline constexpr UINT kFolderViewDirectoryCacheDirty   = WM_APP + 0x304;
inline constexpr UINT kFolderViewBatchIconUpdate       = WM_APP + 0x306;
inline constexpr UINT kNetworkConnectivityChanged      = WM_APP + 0x305;
inline constexpr UINT kFolderViewDeferredInit          = WM_APP + 0x307;
inline constexpr UINT kFolderViewDirectoryImpact       = WM_APP + 0x308;
inline constexpr UINT kFolderViewCreateThumbnailBitmap = WM_APP + 0x309;

inline constexpr UINT kEditSuggestResults = WM_APP + 0x350;

inline constexpr UINT kNavigationMenuRequestPath           = WM_APP + 0x380;
inline constexpr UINT kNavigationMenuShowSiblingsDropdown  = WM_APP + 0x381;
inline constexpr UINT kNavigationMenuShowFullPath          = WM_APP + 0x382;
inline constexpr UINT kNavigationViewDeferredInit          = WM_APP + 0x383;
inline constexpr UINT kNavigationViewRestoreFolderFocus    = WM_APP + 0x384;
inline constexpr UINT kNavigationViewShowHistoryDropdown   = WM_APP + 0x385;
inline constexpr UINT kNavigationViewShowMenuDropdown      = WM_APP + 0x386;
inline constexpr UINT kNavigationViewShowDiskInfoDropdown  = WM_APP + 0x387;
inline constexpr UINT kNavigationViewShowDriveMenuDropdown = WM_APP + 0x388;
inline constexpr UINT kNavigationViewDriveInfoLoaded       = WM_APP + 0x389;

inline constexpr UINT kPaneFocusChanged          = WM_APP + 0x400;
inline constexpr UINT kPaneSelectionSizeComputed = WM_APP + 0x401;
inline constexpr UINT kPaneSelectionSizeProgress = WM_APP + 0x402;
inline constexpr UINT kPaneRestoreFolderFocus    = WM_APP + 0x403;

inline constexpr UINT kFileOperationCompleted = WM_APP + 0x450;
#ifdef ENABLE_TESTS
inline constexpr UINT kFileOpsPopupSelfTestInvoke       = WM_APP + 0x451;
inline constexpr UINT kFileOpsPopupSelfTestSnapshot     = WM_APP + 0x452;
inline constexpr UINT kFileOpsPopupCaptionGlyphSnapshot = WM_APP + 0x453;
inline constexpr UINT kFileOpsPopupLayoutSnapshot       = WM_APP + 0x454;
#endif

inline constexpr UINT kFunctionBarInvoke = WM_APP + 0x460;

inline constexpr UINT kHostShowAlert                      = WM_APP + 0x500;
inline constexpr UINT kHostClearAlert                     = WM_APP + 0x501;
inline constexpr UINT kHostShowPrompt                     = WM_APP + 0x502;
inline constexpr UINT kSettingsApplied                    = WM_APP + 0x503;
inline constexpr UINT kPluginsChanged                     = WM_APP + 0x504;
inline constexpr UINT kPreferencesRequestSettingsSnapshot = WM_APP + 0x505;
inline constexpr UINT kHostShowConnectionManager          = WM_APP + 0x506;
inline constexpr UINT kHostGetConnectionSecret            = WM_APP + 0x507;
inline constexpr UINT kHostPromptConnectionSecret         = WM_APP + 0x508;
inline constexpr UINT kHostClearCachedConnectionSecret    = WM_APP + 0x509;
inline constexpr UINT kHostUpgradeFtpAnonymousToPassword  = WM_APP + 0x50A;
inline constexpr UINT kHostGetConnectionJsonUtf8          = WM_APP + 0x50B;
inline constexpr UINT kConnectionManagerConnect           = WM_APP + 0x50C;
inline constexpr UINT kHostExecuteInPane                  = WM_APP + 0x50D;
inline constexpr UINT kHostSetConnectionSecret            = WM_APP + 0x50E;
inline constexpr UINT kHostDeleteConnectionSecret         = WM_APP + 0x50F;
inline constexpr UINT kSettingsFileChanged                = WM_APP + 0x511;
inline constexpr UINT kSettingsReloadedFromDisk           = WM_APP + 0x512;
inline constexpr UINT kHostOpenViewer                     = WM_APP + 0x513;
#ifdef ENABLE_TESTS
inline constexpr UINT kConnectionCredentialPromptDebug      = WM_APP + 0x514;
inline constexpr UINT kPluginConfigurationDialogDebug       = WM_APP + 0x515;
inline constexpr UINT kFolderViewRenamePromptDebug          = WM_APP + 0x516;
inline constexpr UINT kFolderViewPaneFilterPromptDebug      = WM_APP + 0x517;
inline constexpr UINT kFolderViewSelectionMaskPromptDebug   = WM_APP + 0x518;
inline constexpr UINT kFolderViewChangeCasePromptDebug      = WM_APP + 0x519;
inline constexpr UINT kFolderViewCreateDirectoryPromptDebug = WM_APP + 0x51A;
inline constexpr UINT kConnectionManagerDialogDebug         = WM_APP + 0x51B;
#endif

// Startup milestones (UI thread).
inline constexpr UINT kAppStartupInputReady = WM_APP + 0x510;

// Compare Directories
inline constexpr UINT kCompareDirectoriesDeferredStart   = WM_APP + 0x520;
inline constexpr UINT kCompareDirectoriesScanProgress    = WM_APP + 0x521;
inline constexpr UINT kCompareDirectoriesExecuteCommand  = WM_APP + 0x522;
inline constexpr UINT kCompareDirectoriesDecisionUpdated = WM_APP + 0x523;
inline constexpr UINT kCompareDirectoriesContentProgress = WM_APP + 0x524;

// Change Case (background)
inline constexpr UINT kChangeCaseTaskUpdate = WM_APP + 0x525;
inline constexpr UINT kChangeCaseCompleted  = WM_APP + 0x526;

// Find Files
inline constexpr UINT kFindSearchResults         = WM_APP + 0x527;
inline constexpr UINT kFindSearchProgress        = WM_APP + 0x528;
inline constexpr UINT kFindSearchComplete        = WM_APP + 0x529;
inline constexpr UINT kFindSearchDeferredRefresh = WM_APP + 0x52A;
inline constexpr UINT kFindShowActionMenu        = WM_APP + 0x52D;

// Change Attributes (background)
inline constexpr UINT kChangeAttributesTaskUpdate = WM_APP + 0x52B;
inline constexpr UINT kChangeAttributesCompleted  = WM_APP + 0x52C;

// Preferences
inline constexpr UINT kPreferencesApplyComboThemeDeferred = WM_APP + 0x530;
inline constexpr UINT kPreferencesSelectPluginDetails     = WM_APP + 0x531;
inline constexpr UINT kPreferencesDeferredPaneAction      = WM_APP + 0x532;
// WM_APP + 0x533..0x535 were the retired DxUi text-input bridge messages; keep them reserved.
inline constexpr UINT kPreferencesApplyPageHostScroll       = WM_APP + 0x538;
inline constexpr UINT kPluginConfigurationDialogApplyTheme  = WM_APP + 0x536;
inline constexpr UINT kConnectionCredentialPromptApplyTheme = WM_APP + 0x537;
inline constexpr UINT kDxUiContextMenuRootHoverChanged      = WM_APP + 0x539;

// Item Properties
inline constexpr UINT kItemPropertiesLoadComplete               = WM_APP + 0x540;
inline constexpr UINT kItemPropertiesRemoveStream               = WM_APP + 0x541;
inline constexpr UINT kFolderWindowCloseOpenedFilesDialog       = WM_APP + 0x542;
inline constexpr UINT kFolderWindowCloseSharedDirectoriesDialog = WM_APP + 0x543;

// Splash screen
inline constexpr UINT kSplashScreenSetText  = WM_APP + 0x6F0;
inline constexpr UINT kSplashScreenRecenter = WM_APP + 0x6F1;

// Plugin viewers (async work completion / progress)
inline constexpr UINT kViewerTextAsyncOpenComplete     = WM_APP + 0x600;
inline constexpr UINT kViewerPeAsyncParseComplete      = WM_APP + 0x601;
inline constexpr UINT kViewerWebAsyncLoadComplete      = WM_APP + 0x602;
inline constexpr UINT kViewerImgRawAsyncOpenComplete   = WM_APP + 0x603;
inline constexpr UINT kViewerImgRawAsyncProgress       = WM_APP + 0x604;
inline constexpr UINT kViewerImgRawAsyncExportComplete = WM_APP + 0x605;
inline constexpr UINT kViewerSqliteAsyncOpenComplete   = WM_APP + 0x606;
inline constexpr UINT kViewerSqliteAsyncQueryComplete  = WM_APP + 0x607;
#ifdef ENABLE_TESTS
enum class ViewerSqliteDebugFocusTarget : uint8_t
{
    None,
    FileCombo,
    ReloadButton,
    TableCombo,
    PrevButton,
    NextButton,
    QueryField,
    RunButton,
    TableButton,
    ResultGrid,
};

struct ViewerSqliteDebugSnapshot
{
    size_t rowCount                          = 0u;
    size_t selectionCount                    = 0u;
    size_t visibleRowCount                   = 0u;
    size_t visibleColumnCount                = 0u;
    size_t visibleCellCount                  = 0u;
    bool hasVerticalScrollbar                = false;
    bool themeDark                           = false;
    bool themeHighContrast                   = false;
    bool themeRainbow                        = false;
    uint64_t renderCount                     = 0u;
    uint64_t resizeCount                     = 0u;
    uint64_t resizeFailureCount              = 0u;
    uint32_t pendingAsyncWork                = 0u;
    bool tablePreviewMode                    = true;
    bool hasMoreRows                         = false;
    bool prevButtonEnabled                   = false;
    bool nextButtonEnabled                   = false;
    uint64_t rowOffset                       = 0u;
    size_t sortColumnIndex                   = static_cast<size_t>(-1);
    uint32_t sortDirection                   = 0u;
    bool hasStatusStrip                      = false;
    bool statusStripVisible                  = false;
    uint32_t statusStripSectionCount         = 0u;
    float statusStripHeightDip               = 0.0f;
    uint64_t primarySelectedRowId            = 0u;
    uint64_t firstRowPrimaryKey              = 0u;
    uint64_t lastRowPrimaryKey               = 0u;
    uint32_t selectedRowFillArgb             = 0u;
    uint32_t selectedRowTextArgb             = 0u;
    bool selectedRowUsesRainbow              = false;
    ViewerSqliteDebugFocusTarget focusTarget = ViewerSqliteDebugFocusTarget::None;
};

enum class ViewerSqliteDebugPageCommand : uint8_t
{
    Previous,
    Next,
};

inline constexpr UINT kViewerSqliteDebugGetPendingAsyncWork      = WM_APP + 0x608;
inline constexpr UINT kViewerSqliteDebugGetSnapshot              = WM_APP + 0x609;
inline constexpr UINT kViewerSqliteDebugScrollGridByWheelDetents = WM_APP + 0x60A;
inline constexpr UINT kViewerSqliteDebugSelectGridRow            = WM_APP + 0x60B;
inline constexpr UINT kViewerSqliteDebugInvokePageCommand        = WM_APP + 0x60C;
inline constexpr UINT kViewerSqliteDebugCycleSortColumn          = WM_APP + 0x60D;

enum class ViewerTextDebugViewMode : uint8_t
{
    Text,
    Hex,
};

enum class ViewerTextDebugDocumentKind : uint8_t
{
    PlainText,
    Diff,
};

enum class ViewerTextDebugDiffPresentation : uint8_t
{
    None,
    RawText,
    Inline,
    SideBySide,
};

enum class ViewerTextDebugHexByteColorMode : uint8_t
{
    Off,
    LeadingNibble,
};

inline constexpr size_t kViewerTextDebugTextPreviewChars     = 512u;
inline constexpr size_t kViewerTextDebugDocumentPreviewChars = 2048u;

struct ViewerTextDebugSnapshot
{
    ViewerTextDebugViewMode viewMode                 = ViewerTextDebugViewMode::Text;
    ViewerTextDebugDocumentKind documentKind         = ViewerTextDebugDocumentKind::PlainText;
    ViewerTextDebugDiffPresentation diffPresentation = ViewerTextDebugDiffPresentation::None;
    ViewerTextDebugHexByteColorMode hexByteColorMode = ViewerTextDebugHexByteColorMode::Off;
    size_t fileSectionCount                          = 0u;
    size_t fileComboEntryCount                       = 0u;
    size_t diffHunkCount                             = 0u;
    size_t activeDiffSectionIndex                    = 0u;
    size_t activeDiffHunkIndex                       = 0u;
    size_t styledRowCount                            = 0u;
    size_t contextRowCount                           = 0u;
    size_t addedRowCount                             = 0u;
    size_t removedRowCount                           = 0u;
    size_t headerRowCount                            = 0u;
    size_t bannerRowCount                            = 0u;
    size_t placeholderRowCount                       = 0u;
    size_t placeholderBandCount                      = 0u;
    size_t deferredContextRowCount                   = 0u;
    bool diffParsedAvailable                         = false;
    bool diffExpandedContext                         = false;
    bool diffHasPlaceholderRows                      = false;
    bool diffReferencedFilesResolved                 = false;
    bool fileComboUsesDiffSections                   = false;
    uint64_t diffParseCount                          = 0u;
    uint64_t renderCount                             = 0u;
    size_t legacyVisibleGdiTextSurfaceCount          = 0u;
    size_t legacyVisibleHfontSurfaceCount            = 0u;
    size_t visibleRowCount                           = 0u;
    size_t topVisibleLogicalLine                     = 0u;
    size_t textLeftColumn                            = 0u;
    uint32_t topVisibleSegmentColumnStart            = 0u;
    uint32_t topVisibleLeftPaneColumnStart           = 0u;
    uint32_t topVisibleRightPaneColumnStart          = 0u;
    uint32_t firstVisibleSplitLeftPaneColumnStart    = 0u;
    uint32_t firstVisibleSplitRightPaneColumnStart   = 0u;
    size_t firstClickableBannerLogicalLine           = static_cast<size_t>(-1);
    size_t hydratedLogicalLineStart                  = 0u;
    size_t hydratedLogicalLineEndExclusive           = 0u;
    size_t builtLogicalLineCount                     = 0u;
    uint64_t referencedBytesRead                     = 0u;
    bool diffHasExpandableTail                       = false;
    size_t visibleByteCount                          = 0u;
    size_t visibleColorizedByteCount                 = 0u;
    size_t visibleUniqueColorBucketCount             = 0u;
    size_t visibleStyledRowCount                     = 0u;
    size_t visibleContextRowCount                    = 0u;
    size_t visibleAddedRowCount                      = 0u;
    size_t visibleRemovedRowCount                    = 0u;
    size_t visibleHeaderRowCount                     = 0u;
    size_t visibleBannerRowCount                     = 0u;
    size_t visibleGapHatchCount                      = 0u;
    size_t visibleSplitRowCount                      = 0u;
    uint64_t textLastPaintUs                         = 0u;
    bool paneLocalSideBySideLayout                   = false;
    uint32_t sideBySideLeftPaneColumns               = 0u;
    uint32_t sideBySideRightPaneColumns              = 0u;
    uint32_t sideBySideSeparatorColumns              = 0u;
    bool diffContextUsesBaseBackground               = false;
    uint32_t diffMarkerArgb                          = 0u;
    uint32_t diffGapHatchArgb                        = 0u;
    bool themeRainbow                                = false;
    bool highContrastFallback                        = false;
    uint32_t diffAddedBackgroundArgb                 = 0u;
    uint32_t diffRemovedBackgroundArgb               = 0u;
    uint32_t diffContextBackgroundArgb               = 0u;
    uint32_t diffHeaderBackgroundArgb                = 0u;
    uint32_t diffBannerBackgroundArgb                = 0u;
    uint32_t diffPlaceholderBackgroundArgb           = 0u;
    uint32_t diffDividerArgb                         = 0u;
    bool hasLastContextMenuScreenPoint               = false;
    LONG lastContextMenuScreenX                      = 0;
    LONG lastContextMenuScreenY                      = 0;
    bool hasLastTextViewMouseMoveClientPoint         = false;
    LONG lastTextViewMouseMoveClientX                = 0;
    LONG lastTextViewMouseMoveClientY                = 0;
    bool lastTextViewMouseMoveHit                    = false;
    size_t lastTextViewMouseMoveLogicalLine          = static_cast<size_t>(-1);
    wchar_t firstTextLine[kViewerTextDebugTextPreviewChars]{};
    wchar_t secondTextLine[kViewerTextDebugTextPreviewChars]{};
    wchar_t topVisibleTextLine[kViewerTextDebugTextPreviewChars]{};
    wchar_t textPreview[kViewerTextDebugDocumentPreviewChars]{};
};

inline constexpr UINT kViewerTextDebugGetSnapshot          = WM_APP + 0x60E;
inline constexpr UINT kViewerTextDebugSelectDiffSection    = WM_APP + 0x60F;
inline constexpr UINT kViewerTextDebugSelectDiffHunk       = WM_APP + 0x610;
inline constexpr UINT kViewerTextDebugClickTextLogicalLine = WM_APP + 0x611;

struct ViewerNativeMenuModelDebugSnapshot
{
    bool hasHiddenMenuModel   = false;
    size_t ownerDrawItemCount = 0u;
};

inline constexpr UINT kViewerDebugGetNativeMenuModelSnapshot = WM_APP + 0x612;
#endif
inline constexpr UINT kViewerVlcAsyncOpenComplete = WM_APP + 0x613;
#ifdef ENABLE_TESTS
struct ViewerVlcDebugSnapshot
{
    bool loadingActive                   = false;
    bool loadingVisible                  = false;
    bool missingVisible                  = false;
    bool hasVideoChild                   = false;
    bool videoChildIsChildWindow         = false;
    bool videoChildParentIsViewer        = false;
    bool hasVolumeMuteButton             = false;
    bool hasVolumeSlider                 = false;
    bool muted                           = false;
    int volume                           = 0;
    int64_t timeMs                       = 0;
    int64_t lengthMs                     = 0;
    LONG snapshotWidth                   = 0;
    LONG snapshotHeight                  = 0;
    bool hudButtonsUseFilledButtonStyle  = false;
    bool hudIconsUseIconFont             = false;
    wchar_t playPauseIconGlyph           = L'\0';
    wchar_t stopIconGlyph                = L'\0';
    wchar_t snapshotIconGlyph            = L'\0';
    bool volumeIconUsesIconFont          = false;
    wchar_t volumeIconGlyph              = L'\0';
    bool loadingSpinnerUsesRainbow       = false;
    int loadingSpinnerDotCount           = 0;
    LONG loadingSpinnerOrbitPx           = 0;
    LONG loadingSpinnerDotRadiusPx       = 0;
    LONG loadingSpinnerActiveDotRadiusPx = 0;
    uint32_t loadingSpinnerFirstDotArgb  = 0;
    uint32_t loadingSpinnerSecondDotArgb = 0;
};

struct ViewerVlcDebugPlaybackState
{
    int64_t timeMs   = 0;
    int64_t lengthMs = 0;
    int volume       = 100;
    bool muted       = false;
};

struct ViewerVlcDebugWheel
{
    int wheelDelta = 0;
    bool shift     = false;
    bool ctrl      = false;
};

struct ViewerVlcDebugStopDelay
{
    uint32_t delayMs = 0;
};

inline constexpr UINT kViewerVlcDebugGetSnapshot         = WM_APP + 0x614;
inline constexpr UINT kViewerVlcDebugForceLoadingVisible = WM_APP + 0x615;
inline constexpr UINT kViewerVlcDebugSetPlaybackState    = WM_APP + 0x616;
inline constexpr UINT kViewerVlcDebugWheel               = WM_APP + 0x617;
inline constexpr UINT kViewerVlcDebugToggleMute          = WM_APP + 0x618;
inline constexpr UINT kViewerVlcDebugWheelVideoChild     = WM_APP + 0x619;
inline constexpr UINT kViewerVlcDebugSetStopDelay        = WM_APP + 0x61A;

struct ViewerSpaceTooltipDebugSnapshot
{
    uint32_t tooltipNodeId     = 0;
    size_t tooltipTextLength   = 0u;
    float tooltipAnchorXDip    = 0.0f;
    float tooltipAnchorYDip    = 0.0f;
    float tooltipMaxWidthDip   = 0.0f;
    float tooltipMaxHeightDip  = 0.0f;
    uint64_t tooltipPaintCount = 0u;
    bool hasRenderTarget = false;
    bool hasTooltipFormat = false;
    uint32_t hoverNodeId = 0;
    bool hasLastMouseMoveClientPoint = false;
    LONG lastMouseMoveClientX = 0;
    LONG lastMouseMoveClientY = 0;
    bool hasLastContextMenuScreenPoint = false;
    LONG lastContextMenuScreenX = 0;
    LONG lastContextMenuScreenY = 0;
    uint32_t lastContextMenuHitNodeId = 0;
};

enum class ViewerSpacePerfRendererMode : uint8_t
{
    None,
    HwndRenderTarget,
    DeviceContext,
};

enum class ViewerSpacePerfScanState : uint8_t
{
    NotStarted,
    Queued,
    Scanning,
    Done,
    Error,
    Canceled,
};

struct ViewerSpacePerfDebugSnapshot
{
    ViewerSpacePerfRendererMode rendererMode = ViewerSpacePerfRendererMode::None;
    ViewerSpacePerfScanState scanState       = ViewerSpacePerfScanState::NotStarted;
    bool hasRenderTarget                     = false;
    size_t realDirectoryCount                = 0u;
    size_t fileCandidateCount                = 0u;
    size_t syntheticCount                    = 0u;
    size_t pendingQueueCount                 = 0u;
    uint64_t pendingQueueBytes               = 0u;
    size_t drawItemCount                     = 0u;
    size_t visibleTileCount                  = 0u;
    size_t culledTileCount                   = 0u;
    uint64_t rootTotalBytes                  = 0u;
    uint64_t fileCandidateBytes              = 0u;
    uint64_t syntheticBytes                  = 0u;
    uint64_t lastTileDrawCount               = 0u;
    uint64_t lastTextDrawCount               = 0u;
    uint64_t staticCacheGeneration           = 0u;
    uint64_t staticCacheHits                 = 0u;
    uint64_t staticCacheMisses               = 0u;
    uint64_t staticCacheRecordCount          = 0u;
    uint64_t staticCacheBytes                = 0u;
    uint64_t scanCacheSnapshotBytes          = 0u;
    uint64_t lastStaticCacheRecordUs         = 0u;
    uint64_t lastPaintUs                     = 0u;
    uint64_t lastLayoutUs                    = 0u;
    uint64_t lastDrainUs                     = 0u;
    uint64_t lastWorkingSetBytes             = 0u;
    uint64_t lastHitTestUs                   = 0u;
    uint32_t lastHitTestCandidatesChecked    = 0u;
    uint32_t effectiveFileCandidateBudget    = 0u;
    uint32_t rendererDeviceCreateCount       = 0u;
    uint32_t swapChainResizeCount            = 0u;
    uint32_t rendererBrushCreateCount        = 0u;
    uint32_t rendererTextFormatCreateCount   = 0u;
    uint32_t swapChainWidthPx                = 0u;
    uint32_t swapChainHeightPx               = 0u;
    uint32_t rendererFailureStage            = 0u;
    uint32_t rendererFailureHr               = 0u;
    uint64_t layoutGeneration                = 0u;
    uint32_t hitGridCellCount                = 0u;
    uint32_t hitGridMaxCandidatesPerCell     = 0u;
};

struct ViewerSpaceHitTestDebugSnapshot
{
    uint32_t sampleCount                 = 0u;
    uint32_t mismatchCount               = 0u;
    uint32_t gridHitCount                = 0u;
    uint32_t linearHitCount              = 0u;
    uint32_t maxGridCandidatesChecked    = 0u;
    uint32_t maxLinearCandidatesChecked  = 0u;
    uint32_t lastMismatchGridNodeId      = 0u;
    uint32_t lastMismatchLinearNodeId    = 0u;
};

enum class ViewerSpaceRendererFaultDebugMode : uint8_t
{
    D2DRecreateTarget = 1,
    DxgiDeviceRemoved = 2,
};

inline constexpr UINT kViewerSpaceDebugGetTooltipSnapshot = WM_APP + 0x61B;
inline constexpr UINT kViewerSpaceDebugShowTooltipOverlay = WM_APP + 0x61C;
inline constexpr UINT kViewerSpaceDebugGetPerfSnapshot    = WM_APP + 0x61D;
inline constexpr UINT kViewerSpaceDebugForcePerfSample    = WM_APP + 0x61E;
inline constexpr UINT kViewerSpaceDebugCompareHitTesting  = WM_APP + 0x61F;
inline constexpr UINT kViewerSpaceDebugForceRendererFault = WM_APP + 0x623;
#endif

// RedSalamanderMonitor / ColorTextView
inline constexpr UINT kColorTextViewLayoutReady = WM_APP + 0x620;
inline constexpr UINT kColorTextViewWidthReady  = WM_APP + 0x621;
inline constexpr UINT kColorTextViewEtwBatch    = WM_APP + 0x622;

// PoC / samples
inline constexpr UINT kWin32HelloCredHelloResult = WM_APP + 0x680;

// Custom control messages (WM_USER)
inline constexpr UINT kModernComboSetCloseOutsideAccept  = WM_USER + 0x500u;
inline constexpr UINT kModernComboSetDropDownPreferBelow = WM_USER + 0x501u;
inline constexpr UINT kModernComboSetPinnedIndex         = WM_USER + 0x502u;
inline constexpr UINT kModernComboSetCompactMode         = WM_USER + 0x503u;
inline constexpr UINT kModernComboSetUseMiddleEllipsis   = WM_USER + 0x504u;

} // namespace WndMsg

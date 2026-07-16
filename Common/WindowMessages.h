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
inline constexpr UINT kFolderViewPasteShortcutComplete = WM_APP + 0x30A;

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
inline constexpr UINT kConnectionManagerMtpPickerComplete = WM_APP + 0x51C;
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
inline constexpr UINT kCompareDirectoriesDeferredStart      = WM_APP + 0x520;
inline constexpr UINT kCompareDirectoriesScanProgress       = WM_APP + 0x521;
inline constexpr UINT kCompareDirectoriesExecuteCommand     = WM_APP + 0x522;
inline constexpr UINT kCompareDirectoriesDecisionUpdated    = WM_APP + 0x523;
inline constexpr UINT kCompareDirectoriesContentProgress    = WM_APP + 0x524;
inline constexpr UINT kCompareDirectoriesLeaveScopePrompt   = WM_APP + 0x52E;
inline constexpr UINT kCompareDirectoriesDecisionRefreshNow = WM_APP + 0x52F;

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
inline constexpr UINT kPreferencesRestoreCategoryTreeFocus  = WM_APP + 0x53A;

// Item Properties
inline constexpr UINT kItemPropertiesLoadComplete               = WM_APP + 0x540;
inline constexpr UINT kItemPropertiesRemoveStream               = WM_APP + 0x541;
inline constexpr UINT kFolderWindowCloseOpenedFilesDialog       = WM_APP + 0x542;
inline constexpr UINT kFolderWindowCloseSharedDirectoriesDialog = WM_APP + 0x543;

// Batch Rename (background / debug)
inline constexpr UINT kBatchRenameTaskUpdate = WM_APP + 0x544;
inline constexpr UINT kBatchRenameCompleted  = WM_APP + 0x545;
#ifdef ENABLE_TESTS
inline constexpr UINT kBatchRenameWindowDebug = WM_APP + 0x546;
#endif

// Splash screen
inline constexpr UINT kSplashScreenSetText  = WM_APP + 0x6F0;
inline constexpr UINT kSplashScreenRecenter = WM_APP + 0x6F1;

// Plugin viewers (async work completion / progress)
inline constexpr UINT kViewerTextAsyncOpenComplete     = WM_APP + 0x600;
inline constexpr UINT kViewerTextAsyncOpenFailure      = WM_APP + 0x629;
inline constexpr UINT kViewerTextAsyncStreamComplete   = WM_APP + 0x630;
inline constexpr UINT kViewerTextAsyncStreamFailure    = WM_APP + 0x631;
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
    bool isLoading                                   = false;
    uint64_t asyncOpenTerminalCount                  = 0u;
    HRESULT asyncOpenLastTerminalHr                  = E_PENDING;
    bool textStreamActive                            = false;
    uint64_t textStreamStartOffset                   = 0u;
    uint64_t textStreamEndOffset                     = 0u;
    uint64_t textFileSize                            = 0u;
    bool textStreamLoadPending                       = false;
    uint64_t textStreamAcceptedCount                 = 0u;
    uint64_t textStreamRejectedCount                 = 0u;
    uint64_t textStreamStaleCount                    = 0u;
    uint64_t textStreamTerminalCount                 = 0u;
    HRESULT textStreamLastTerminalHr                 = E_PENDING;
    uint64_t textStreamLastElapsedUs                 = 0u;
    uint64_t textStreamLastUiApplyUs                 = 0u;
    uint64_t textVisualLineCount                     = 0u;
    bool textVisualLineCountExact                    = true;
    size_t textMaterializedVisualLineCount           = 0u;
    size_t textSparseLogicalSummaryCount             = 0u;
    bool textSparseWrapActive                        = false;
    size_t textLayoutCacheEntryCount                 = 0u;
    size_t textLayoutCacheBytes                      = 0u;
    size_t textLayoutCacheMaxEntries                 = 0u;
    size_t textLayoutCacheMaxBytes                   = 0u;
    uint64_t textLayoutCacheEvictions                = 0u;
    uint64_t textLayoutCacheHits                     = 0u;
    uint64_t textLayoutCacheMisses                   = 0u;
    size_t textCaretIndex                            = 0u;
    size_t textSelectionAnchor                       = 0u;
    size_t textSelectionActive                       = 0u;
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

enum class ViewerTextDebugSaveFault : uint8_t
{
    None,
    SourceOpen,
    SourceRead,
    Encode,
    Write,
    Flush,
    Commit,
};

enum class ViewerTextDebugAsyncOpenFault : uint8_t
{
    None,
    FileSystemIo,
    OpenReader,
    GetSize,
    InitialSeek,
    InitialRead,
    DataSeek,
    DataRead,
    Decode,
    PayloadPost,
    Submit,
};

enum class ViewerTextDebugGeometryOperation : uint8_t
{
    SetCacheBudget,
    SetAsyncStreamFault,
    ProbeLayout,
    ProbeWrappedCoverage,
};

enum class ViewerTextDebugAsyncStreamFault : uint8_t
{
    None,
    Allocation,
    Submit,
    Worker,
    PayloadPost,
};

struct ViewerTextDebugGeometryRequest
{
    ViewerTextDebugGeometryOperation operation          = ViewerTextDebugGeometryOperation::ProbeLayout;
    ViewerTextDebugAsyncStreamFault asyncStreamFault    = ViewerTextDebugAsyncStreamFault::None;
    size_t cacheMaxEntries                              = 0u;
    size_t cacheMaxBytes                                = 0u;
    size_t segmentStart                                 = 0u;
    size_t segmentEnd                                   = 0u;
    size_t textPosition                                 = 0u;
    size_t rangeStart                                   = 0u;
    size_t rangeEnd                                     = 0u;
    size_t logicalLine                                  = 0u;
    float widthDip                                      = 0.0f;
    float hitX                                          = 0.0f;
    size_t normalizedTextPosition = 0u;
    size_t previousTextPosition   = 0u;
    size_t nextTextPosition       = 0u;
    size_t hitTextPosition        = 0u;
    float caretX                  = 0.0f;
    float rangeLeft               = 0.0f;
    float rangeRight              = 0.0f;
    size_t wrappedSegmentCount    = 0u;
    size_t wrappedCoveredStart    = 0u;
    size_t wrappedCoveredEnd      = 0u;
    size_t wrappedSecondSegmentStart = 0u;
    size_t wrappedSecondSegmentEnd   = 0u;
    float wrappedWidthDip         = 0.0f;
    float wrappedLineHeightDip    = 0.0f;
    bool wrappedHasGapOrOverlap   = false;
    bool wrappedAllSegmentsFit    = false;
    HRESULT result                = E_UNEXPECTED;
};

struct ViewerTextDebugSaveRequest
{
    const wchar_t* destinationPath = nullptr;
    UINT encodingSelection         = 0u;
    ViewerTextDebugSaveFault fault = ViewerTextDebugSaveFault::None;
    bool simulateLoading           = false;
    HRESULT result                 = E_UNEXPECTED;
};

inline constexpr size_t kViewerTextDebugHexLineChars = 32u;

struct ViewerTextDebugHexLineRequest
{
    uint64_t offset = 0u;
    size_t validBytes = 0u;
    wchar_t text[kViewerTextDebugHexLineChars]{};
    size_t sourceLengths[16]{};
    size_t columnStarts[16]{};
    size_t columnLengths[16]{};
};

struct ViewerTextDebugHexCopyRequest
{
    uint64_t anchorOffset = 0u;
    uint64_t activeOffset = 0u;
    bool dispatched       = false;
};

inline constexpr UINT kViewerTextDebugGetSnapshot          = WM_APP + 0x60E;
inline constexpr UINT kViewerTextDebugSelectDiffSection    = WM_APP + 0x60F;
inline constexpr UINT kViewerTextDebugSelectDiffHunk       = WM_APP + 0x610;
inline constexpr UINT kViewerTextDebugClickTextLogicalLine = WM_APP + 0x611;
inline constexpr UINT kViewerTextDebugSaveAs               = WM_APP + 0x625;
inline constexpr UINT kViewerTextDebugReloadWithOpenFault  = WM_APP + 0x62A;
inline constexpr UINT kViewerTextDebugFormatUtf8HexLine    = WM_APP + 0x62E;
inline constexpr UINT kViewerTextDebugCopyHexSelection     = WM_APP + 0x62F;
inline constexpr UINT kViewerTextDebugSetLayoutCacheBudget = WM_APP + 0x632;

struct ViewerPeDebugSnapshot
{
    bool isLoading          = false;
    uint64_t requestId      = 0u;
    uint64_t windowIdentity = 0u;
    HRESULT parseHr         = E_PENDING;
    size_t bodyLength       = 0u;
    wchar_t bodyPreview[128]{};
};

inline constexpr UINT kViewerPeDebugGetSnapshot = WM_APP + 0x62B;

enum class ViewerPeDebugAsyncFault : unsigned int
{
    None = 0u,
    ResultAllocation,
    PayloadPost,
};

inline constexpr UINT kViewerPeDebugReloadWithAsyncFault = WM_APP + 0x634;

struct ViewerNativeMenuModelDebugSnapshot
{
    bool hasHiddenMenuModel   = false;
    size_t ownerDrawItemCount = 0u;
};

inline constexpr UINT kViewerDebugGetNativeMenuModelSnapshot = WM_APP + 0x612;

struct ViewerImgRawDecodeDebugSnapshot
{
    bool hasImage              = false;
    bool displayingThumbnail   = false;
    uint16_t baseOrientation   = 1u;
    uint16_t viewOrientation   = 1u;
    uint32_t sourceWidth       = 0u;
    uint32_t sourceHeight      = 0u;
    uint32_t orientedWidth     = 0u;
    uint32_t orientedHeight    = 0u;
};

inline constexpr UINT kViewerImgRawDebugGetDecodeSnapshot = WM_APP + 0x624;

struct ViewerImgRawResourceDebugSnapshot
{
    uint64_t speculativeBytes        = 0u;
    uint64_t speculativeBytesPeak    = 0u;
    uint64_t speculativeBytesLimit   = 0u;
    uint64_t budgetAcceptedCount     = 0u;
    uint64_t budgetRejectedCount     = 0u;
    uint64_t currentRequestId        = 0u;
    uint64_t finalSuccessCount       = 0u;
    uint64_t finalFailureCount       = 0u;
    uint64_t previewSuccessCount     = 0u;
    uint64_t lastPreviewApplyOrdinal = 0u;
    uint64_t lastFinalApplyOrdinal   = 0u;
    uint64_t replacedMainDecodeCount = 0u;
    size_t cachedImageCount          = 0u;
    size_t inflightDecodeCount       = 0u;
    size_t activeMainDecodeCount     = 0u;
    size_t pendingMainDecodeCount    = 0u;
    bool terminalFallbackPending     = false;
    bool loading                     = false;
};

struct ViewerImgRawDebugExportRequest
{
    const wchar_t* destinationPath = nullptr;
    bool queued                    = false;
};

inline constexpr UINT kViewerImgRawDebugGetResourceSnapshot = WM_APP + 0x626;
inline constexpr UINT kViewerImgRawDebugClearImageCache     = WM_APP + 0x627;
inline constexpr UINT kViewerImgRawDebugExportToPath        = WM_APP + 0x628;
#endif
inline constexpr UINT kViewerVlcAsyncOpenComplete  = WM_APP + 0x613;
inline constexpr UINT kViewerVlcAsyncCloseComplete = WM_APP + 0x62C;
inline constexpr UINT kViewerVlcAsyncFallbackReady = WM_APP + 0x633;
#ifdef ENABLE_TESTS
struct ViewerVlcDebugSnapshot
{
    bool loadingActive                   = false;
    bool loadingVisible                  = false;
    bool missingVisible                  = false;
    bool hasVideoChild                   = false;
    bool videoChildIsChildWindow         = false;
    bool videoChildParentIsViewer        = false;
    bool vlcModuleLoaded                 = false;
    bool vlcInstanceLoaded               = false;
    bool vlcPlayerCreated                = false;
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
    uint64_t windowIdentity              = 0;
    uint64_t loadRequestId               = 0;
    uint64_t pendingLoadWorkCount        = 0;
    uint64_t loadQueueAccepted           = 0;
    uint64_t loadQueueRejected           = 0;
    uint64_t staleLoadResults            = 0;
    uint64_t cleanupCompletions           = 0;
    uint64_t deferredCleanupCount         = 0;
    uint64_t cleanupDeferrals             = 0;
    uint64_t cleanupSubmitFailures        = 0;
    uint64_t cleanupAllocationFailures    = 0;
    uint64_t asyncResultPostFailures      = 0;
    uint64_t loadPostFallbacks            = 0;
    uint64_t cleanupPostFallbacks         = 0;
    bool closePending                     = false;
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
    HANDLE releaseGate = nullptr;
};

struct ViewerVlcDebugAsyncControl
{
    uint32_t loadDelayMs = 0;
    bool failNextLoadSubmit = false;
    bool failNextLoadCompletionPost = false;
    bool failNextCloseCompletionPost = false;
    bool failNextCleanupSubmit = false;
    bool failNextCleanupAllocation = false;
};

inline constexpr UINT kViewerVlcDebugGetSnapshot         = WM_APP + 0x614;
inline constexpr UINT kViewerVlcDebugForceLoadingVisible = WM_APP + 0x615;
inline constexpr UINT kViewerVlcDebugSetPlaybackState    = WM_APP + 0x616;
inline constexpr UINT kViewerVlcDebugWheel               = WM_APP + 0x617;
inline constexpr UINT kViewerVlcDebugToggleMute          = WM_APP + 0x618;
inline constexpr UINT kViewerVlcDebugWheelVideoChild     = WM_APP + 0x619;
inline constexpr UINT kViewerVlcDebugSetStopDelay        = WM_APP + 0x61A;
inline constexpr UINT kViewerVlcDebugSetAsyncControl     = WM_APP + 0x62D;

struct ViewerSpaceTooltipDebugSnapshot
{
    uint32_t tooltipNodeId             = 0;
    size_t tooltipTextLength           = 0u;
    float tooltipAnchorXDip            = 0.0f;
    float tooltipAnchorYDip            = 0.0f;
    float tooltipMaxWidthDip           = 0.0f;
    float tooltipMaxHeightDip          = 0.0f;
    uint64_t tooltipPaintCount         = 0u;
    bool hasRenderTarget               = false;
    bool hasTooltipFormat              = false;
    uint32_t hoverNodeId               = 0;
    bool hasLastMouseMoveClientPoint   = false;
    LONG lastMouseMoveClientX          = 0;
    LONG lastMouseMoveClientY          = 0;
    bool hasLastContextMenuScreenPoint = false;
    LONG lastContextMenuScreenX        = 0;
    LONG lastContextMenuScreenY        = 0;
    uint32_t lastContextMenuHitNodeId  = 0;
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
    uint64_t aggregateBytes                  = 0u;
    uint64_t aggregateFolders                = 0u;
    uint64_t aggregateFiles                  = 0u;
    uint64_t scannedFolders                  = 0u;
    uint64_t scannedFiles                    = 0u;
    uint64_t childReferenceCount             = 0u;
    uint64_t childArenaSlots                 = 0u;
    uint64_t childArenaFreeSlots             = 0u;
    uint64_t modelAcceptedEntries            = 0u;
    uint64_t modelRejectedEntries            = 0u;
    uint64_t modelCappedDirectories          = 0u;
    uint64_t modelCappedFiles                = 0u;
    uint64_t modelRetainedNameBytes          = 0u;
    uint64_t modelRetainedChildReferences    = 0u;
    uint64_t modelTraversedDirectories       = 0u;
    uint64_t modelRetainedDirectoryLimit     = 0u;
    uint64_t modelRetainedFileLimit          = 0u;
    uint64_t modelChildReferenceLimit        = 0u;
    uint64_t modelChildArenaSlotLimit        = 0u;
    uint32_t modelValidationError            = 0u;
    uint64_t postUpdateInnerGenerationRejects = 0u;
};

struct ViewerSpacePostUpdatePauseDebugControl
{
    HANDLE enteredEvent = nullptr;
    HANDLE releaseEvent = nullptr;
};

struct ViewerSpaceHitTestDebugSnapshot
{
    uint32_t sampleCount                = 0u;
    uint32_t mismatchCount              = 0u;
    uint32_t gridHitCount               = 0u;
    uint32_t linearHitCount             = 0u;
    uint32_t maxGridCandidatesChecked   = 0u;
    uint32_t maxLinearCandidatesChecked = 0u;
    uint32_t lastMismatchGridNodeId     = 0u;
    uint32_t lastMismatchLinearNodeId   = 0u;
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
inline constexpr UINT kViewerSpaceDebugPauseNextPostUpdate = WM_APP + 0x635;
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

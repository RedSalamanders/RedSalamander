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

inline constexpr UINT kFolderViewSyncSwapChain       = WM_APP + 0x300;
inline constexpr UINT kFolderViewEnumerateComplete   = WM_APP + 0x301;
inline constexpr UINT kFolderViewIconLoaded          = WM_APP + 0x302;
inline constexpr UINT kFolderViewCreateIconBitmap    = WM_APP + 0x303;
inline constexpr UINT kFolderViewDirectoryCacheDirty = WM_APP + 0x304;
inline constexpr UINT kFolderViewBatchIconUpdate     = WM_APP + 0x306;
inline constexpr UINT kNetworkConnectivityChanged    = WM_APP + 0x305;
inline constexpr UINT kFolderViewDeferredInit        = WM_APP + 0x307;
inline constexpr UINT kFolderViewDirectoryImpact     = WM_APP + 0x308;

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

inline constexpr UINT kPaneFocusChanged          = WM_APP + 0x400;
inline constexpr UINT kPaneSelectionSizeComputed = WM_APP + 0x401;
inline constexpr UINT kPaneSelectionSizeProgress = WM_APP + 0x402;
inline constexpr UINT kPaneRestoreFolderFocus    = WM_APP + 0x403;

inline constexpr UINT kFileOperationCompleted = WM_APP + 0x450;
#ifdef ENABLE_TESTS
inline constexpr UINT kFileOpsPopupSelfTestInvoke       = WM_APP + 0x451;
inline constexpr UINT kFileOpsPopupSelfTestSnapshot     = WM_APP + 0x452;
inline constexpr UINT kFileOpsPopupCaptionGlyphSnapshot = WM_APP + 0x453;
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

// Preferences
inline constexpr UINT kPreferencesApplyComboThemeDeferred   = WM_APP + 0x530;
inline constexpr UINT kPreferencesSelectPluginDetails       = WM_APP + 0x531;
inline constexpr UINT kPreferencesDeferredPaneAction        = WM_APP + 0x532;
inline constexpr UINT kDxUiTextInputBridgeSync              = WM_APP + 0x533;
inline constexpr UINT kDxUiTextInputBridgeSpecialKey        = WM_APP + 0x534;
inline constexpr UINT kDxUiTextInputBridgeBlur              = WM_APP + 0x535;
inline constexpr UINT kPluginConfigurationDialogApplyTheme  = WM_APP + 0x536;
inline constexpr UINT kConnectionCredentialPromptApplyTheme = WM_APP + 0x537;

// Item Properties
inline constexpr UINT kItemPropertiesLoadComplete = WM_APP + 0x540;
inline constexpr UINT kItemPropertiesRemoveStream = WM_APP + 0x541;

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

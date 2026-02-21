#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
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

inline constexpr UINT kEditSuggestResults = WM_APP + 0x350;

inline constexpr UINT kNavigationMenuRequestPath          = WM_APP + 0x380;
inline constexpr UINT kNavigationMenuShowSiblingsDropdown = WM_APP + 0x381;
inline constexpr UINT kNavigationMenuShowFullPath         = WM_APP + 0x382;
inline constexpr UINT kNavigationViewDeferredInit         = WM_APP + 0x383;

inline constexpr UINT kPaneFocusChanged          = WM_APP + 0x400;
inline constexpr UINT kPaneSelectionSizeComputed = WM_APP + 0x401;
inline constexpr UINT kPaneSelectionSizeProgress = WM_APP + 0x402;

inline constexpr UINT kFileOperationCompleted = WM_APP + 0x450;
#ifdef _DEBUG
inline constexpr UINT kFileOpsPopupSelfTestInvoke = WM_APP + 0x451;
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

// Startup milestones (UI thread).
inline constexpr UINT kAppStartupInputReady = WM_APP + 0x510;

 // Compare Directories
inline constexpr UINT kCompareDirectoriesDeferredStart  = WM_APP + 0x520;
 inline constexpr UINT kCompareDirectoriesScanProgress    = WM_APP + 0x521;
 inline constexpr UINT kCompareDirectoriesExecuteCommand  = WM_APP + 0x522;
 inline constexpr UINT kCompareDirectoriesDecisionUpdated = WM_APP + 0x523;
 inline constexpr UINT kCompareDirectoriesContentProgress = WM_APP + 0x524;

// Change Case (background)
inline constexpr UINT kChangeCaseTaskUpdate = WM_APP + 0x525;
inline constexpr UINT kChangeCaseCompleted  = WM_APP + 0x526;

// Preferences
inline constexpr UINT kPreferencesApplyComboThemeDeferred = WM_APP + 0x530;

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

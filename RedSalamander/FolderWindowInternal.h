#pragma once

// Internal implementation header for FolderWindow split across multiple .cpp files.
// Keep this header private to the FolderWindow translation units.

#include "FolderWindow.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <condition_variable>
#include <cwctype>
#include <deque>
#include <limits>
#include <mutex>
#include <string>
#include <utility>

#include <commctrl.h>
#include <dbt.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <uxtheme.h>
#include <windowsx.h>

#include <shellapi.h>
#include <shlobj_core.h>
#include <shobjidl_core.h>

#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Informations.h"
#include "WindowMessages.h"
#include "resource.h"

constexpr UINT_PTR kLeftNavigationId                = 1001;
constexpr UINT_PTR kLeftFolderViewId                = 1002;
constexpr UINT_PTR kRightNavigationId               = 1003;
constexpr UINT_PTR kRightFolderViewId               = 1004;
constexpr UINT_PTR kLeftStatusBarId                 = 1005;
constexpr UINT_PTR kRightStatusBarId                = 1006;
constexpr UINT_PTR kCommandLineLabelId              = 1007;
constexpr UINT_PTR kCommandLineEditId               = 1008;
constexpr UINT_PTR kLeftFilterBarId                 = 1009;
constexpr UINT_PTR kRightFilterBarId                = 1010;
constexpr UINT_PTR kLeftPreviewTabsId               = 1011;
constexpr UINT_PTR kRightPreviewTabsId              = 1012;
constexpr UINT_PTR kLeftPreviewContentId            = 1013;
constexpr UINT_PTR kRightPreviewContentId           = 1014;
constexpr wchar_t kFolderWindowStatusBarClassName[] = L"RedSalamander.FolderWindow.StatusBar";

constexpr wchar_t kStatusBarOwnerProp[]         = L"RedSalamander.StatusBar.Owner";
constexpr wchar_t kStatusBarSelectionTextProp[] = L"RedSalamander.StatusBar.SelectionText";
constexpr wchar_t kStatusBarSecurityTextProp[]  = L"RedSalamander.StatusBar.SecurityText";
constexpr wchar_t kStatusBarSortTextProp[]      = L"RedSalamander.StatusBar.SortText";
constexpr wchar_t kStatusBarFocusHueProp[]      = L"RedSalamander.StatusBar.FocusHue";
constexpr wchar_t kStatusBarSortHotProp[]       = L"RedSalamander.StatusBar.SortHot";
constexpr wchar_t kCommandLineEditOriginalWndProcProp[] = L"RS.FolderWindow.CommandLine.Edit.OriginalWndProc";
constexpr wchar_t kCommandLineEditOwnerProp[]           = L"RS.FolderWindow.CommandLine.Edit.Owner";

constexpr int kStatusBarPartSelection = 0;
constexpr int kStatusBarPartSecurity  = 1;
constexpr int kStatusBarPartSort      = 2;

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

constexpr int kSplitterWidthDip             = 6;
constexpr int kSplitterGripDotSizeDip       = 2;
constexpr int kSplitterGripDotGapDip        = 2;
constexpr int kSplitterGripDotCount         = 3;
constexpr int kSplitterArrowChevronSizeDip  = 4;
constexpr int kSplitterArrowStrokeWidthDip  = 1;
constexpr int kNavFolderGapDip              = 5;
constexpr int kStatusBarHeightDip           = 22;
constexpr int kStatusBarPaddingXDip         = 4;
constexpr int kStatusBarSortPaddingXDip     = 1;
constexpr int kStatusBarSortMinPartWidthDip = 34;
constexpr float kStatusBarTextSizeDip       = 12.0f;
constexpr int kFilterBarHeightDip           = 26;
constexpr int kPreviewTabHeightDip          = 28;
constexpr int kFilterBarPaddingXDip         = 6;
constexpr int kFunctionBarHeightDip         = 24;
constexpr int kCommandLineHeightDip         = 30;
constexpr int kCommandLinePaddingXDip       = 8;
constexpr int kCommandLinePaddingYDip       = 4;
constexpr int kCommandLineLabelWidthDip     = 78;
constexpr int kCommandLineGapDip            = 6;
constexpr float kMinSplitRatio              = 0.0f;
constexpr float kMaxSplitRatio              = 1.0f;

[[nodiscard]] HRESULT EnsureFolderWindowStatusBarClass(HINSTANCE instance) noexcept;
[[nodiscard]] bool GetStatusBarPartRect(HWND hwnd, int part, RECT& rect) noexcept;
LRESULT CALLBACK StatusBarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

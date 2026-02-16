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
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <uxtheme.h>
#include <windowsx.h>

#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Informations.h"
#include "WindowMessages.h"
#include "resource.h"

constexpr UINT_PTR kLeftNavigationId  = 1001;
constexpr UINT_PTR kLeftFolderViewId  = 1002;
constexpr UINT_PTR kRightNavigationId = 1003;
constexpr UINT_PTR kRightFolderViewId = 1004;
constexpr UINT_PTR kLeftStatusBarId   = 1005;
constexpr UINT_PTR kRightStatusBarId  = 1006;

constexpr wchar_t kStatusBarOwnerProp[]         = L"RedSalamander.StatusBar.Owner";
constexpr wchar_t kStatusBarSelectionTextProp[] = L"RedSalamander.StatusBar.SelectionText";
constexpr wchar_t kStatusBarSortTextProp[]      = L"RedSalamander.StatusBar.SortText";
constexpr wchar_t kStatusBarFocusHueProp[]      = L"RedSalamander.StatusBar.FocusHue";
constexpr wchar_t kStatusBarSortHotProp[]       = L"RedSalamander.StatusBar.SortHot";

using CreateFactoryFunc   = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);
using CreateFactoryExFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

constexpr int kSplitterWidthDip             = 6;
constexpr int kSplitterGripDotSizeDip       = 2;
constexpr int kSplitterGripDotGapDip        = 2;
constexpr int kSplitterGripDotCount         = 3;
constexpr int kNavFolderGapDip              = 5;
constexpr int kStatusBarHeightDip           = 22;
constexpr int kStatusBarPaddingXDip         = 4;
constexpr int kStatusBarSortPaddingXDip     = 2;
constexpr int kStatusBarSortMinPartWidthDip = 32;
constexpr int kFunctionBarHeightDip         = 24;
constexpr float kMinSplitRatio              = 0.0f;
constexpr float kMaxSplitRatio              = 1.0f;

LRESULT CALLBACK StatusBarSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

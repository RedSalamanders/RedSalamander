#include "Framework.h"

#include "CompareDirectoriesWindow.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>

#include "CommandRegistry.h"
#include "CompareDirectoriesEngine.h"
#include "FolderWindow.h"
#include "FolderView.h"
#include "Helpers.h"
#include "ShortcutManager.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

namespace
{
constexpr wchar_t kCompareDirectoriesWindowClassName[] = L"RedSalamander.CompareDirectoriesWindow";
constexpr wchar_t kCompareDirectoriesWindowId[]        = L"CompareDirectoriesWindow";

class CompareDirectoriesWindow;
std::vector<CompareDirectoriesWindow*> g_compareDirectoriesWindows;

constexpr UINT_PTR kLeftPathHeaderId  = 1001;
constexpr UINT_PTR kRightPathHeaderId = 1002;
constexpr UINT_PTR kScanProgressTextId = 1003;
constexpr UINT_PTR kScanProgressBarId  = 1004;

constexpr int kPathHeaderHeightDip      = 22;
constexpr int kScanStatusHeightDip      = 22;
constexpr int kScanStatusPaddingXDip    = 6;
constexpr int kScanProgressBarWidthDip  = 160;
constexpr int kScanProgressBarHeightDip = 10;
constexpr int kSplitterWidthDip         = 6;
constexpr int kSplitterGripDotSizeDip   = 2;
constexpr int kSplitterGripDotGapDip    = 2;
constexpr int kSplitterGripDotCount     = 3;
constexpr float kMinSplitRatio          = 0.0f;
constexpr float kMaxSplitRatio          = 1.0f;

struct ScanProgressPayload
{
    uint32_t activeScans  = 0;
    uint64_t folderCount  = 0;
    uint64_t entryCount   = 0;
    std::filesystem::path relativeFolder;
    std::wstring entryName;
};

struct CompareMenuItemData
{
    bool separator  = false;
    bool topLevel   = false;
    bool hasSubMenu = false;
    std::wstring text;
    std::wstring shortcut;
};

void SplitMenuText(std::wstring_view raw, std::wstring& outText, std::wstring& outShortcut) noexcept
{
    outText.clear();
    outShortcut.clear();

    const size_t tabPos = raw.find(L'\t');
    if (tabPos != std::wstring::npos)
    {
        outText.assign(raw.substr(0, tabPos));
        outShortcut.assign(raw.substr(tabPos + 1));
        return;
    }

    outText.assign(raw);
}

[[nodiscard]] COLORREF SplitterGripColor(const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        return theme.menu.text;
    }

    constexpr int kTowardTextWeight = 1;
    constexpr int kDenom            = 4;
    static_assert(kTowardTextWeight > 0 && kTowardTextWeight < kDenom);

    const int baseWeight           = kDenom - kTowardTextWeight;
    const COLORREF baseColor       = theme.menu.separator;
    const COLORREF towardTextColor = theme.menu.text;

    const int r = (static_cast<int>(GetRValue(baseColor)) * baseWeight + static_cast<int>(GetRValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int g = (static_cast<int>(GetGValue(baseColor)) * baseWeight + static_cast<int>(GetGValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int b = (static_cast<int>(GetBValue(baseColor)) * baseWeight + static_cast<int>(GetBValue(towardTextColor)) * kTowardTextWeight) / kDenom;

    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

[[nodiscard]] std::wstring FormatLocalTimeForDetails(int64_t fileTime) noexcept
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    FILETIME local{};
    SYSTEMTIME st{};
    if (! FileTimeToLocalFileTime(&ft, &local) || ! FileTimeToSystemTime(&local, &st))
    {
        return {};
    }

    return std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

[[nodiscard]] std::wstring FormatFileAttributesForDetails(DWORD attrs) noexcept
{
    std::wstring result;
    result.reserve(10);

    auto add = [&](DWORD flag, wchar_t ch) noexcept
    {
        if ((attrs & flag) != 0)
        {
            result.push_back(ch);
        }
    };

    add(FILE_ATTRIBUTE_READONLY, L'R');
    add(FILE_ATTRIBUTE_HIDDEN, L'H');
    add(FILE_ATTRIBUTE_SYSTEM, L'S');
    add(FILE_ATTRIBUTE_ARCHIVE, L'A');
    add(FILE_ATTRIBUTE_COMPRESSED, L'C');
    add(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    add(FILE_ATTRIBUTE_TEMPORARY, L'T');
    add(FILE_ATTRIBUTE_OFFLINE, L'O');
    add(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (result.empty())
    {
        result = L"-";
    }

    return result;
}

[[nodiscard]] std::wstring BuildMetadataDetailsText(bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes) noexcept
{
    std::wstring result;
    result.reserve(64);

    const std::wstring timeText  = FormatLocalTimeForDetails(lastWriteTime);
    const std::wstring attrsText = FormatFileAttributesForDetails(fileAttributes);

    auto appendToken = [&](std::wstring_view token) noexcept
    {
        if (token.empty())
        {
            return;
        }

        if (! result.empty())
        {
            result.append(L" • ");
        }
        result.append(token);
    };

    appendToken(timeText);
    if (! isDirectory)
    {
        appendToken(FormatBytesCompact(sizeBytes));
    }
    appendToken(attrsText);

    return result;
}

[[nodiscard]] std::wstring GetDlgItemTextString(HWND hwnd, int controlId) noexcept
{
    const HWND ctl = GetDlgItem(hwnd, controlId);
    if (! ctl)
    {
        return {};
    }

    const int len = GetWindowTextLengthW(ctl);
    if (len <= 0)
    {
        return {};
    }

    std::wstring text(static_cast<size_t>(len), L'\0');
    const int copied = GetWindowTextW(ctl, text.data(), len + 1);
    if (copied <= 0)
    {
        return {};
    }

    if (static_cast<size_t>(copied) < text.size())
    {
        text.resize(static_cast<size_t>(copied));
    }
    return text;
}

int MeasureStaticTextHeight(HWND referenceWindow, HFONT font, int width, std::wstring_view text) noexcept
{
    if (! referenceWindow || ! font || width <= 0 || text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return 0;
    }

    auto hdc = wil::GetDC(referenceWindow);
    if (! hdc)
    {
        return 0;
    }

    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);

    RECT rc{};
    rc.left   = 0;
    rc.top    = 0;
    rc.right  = width;
    rc.bottom = 0;

    DrawTextW(hdc.get(), text.data(), static_cast<int>(text.size()), &rc, DT_LEFT | DT_WORDBREAK | DT_NOPREFIX | DT_CALCRECT);

    const UINT dpi     = GetDpiForWindow(referenceWindow);
    const int paddingY = ThemedControls::ScaleDip(dpi, 6);
    return static_cast<int>(std::max(0l, rc.bottom - rc.top) + std::max(1, paddingY));
}

void SetTwoStateToggleState(HWND toggle, bool highContrast, bool toggledOn) noexcept
{
    if (! toggle)
    {
        return;
    }

    const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    const UINT type      = static_cast<UINT>(style & BS_TYPEMASK);
    bool useBmCheck      = highContrast;
    if (type == BS_OWNERDRAW)
    {
        useBmCheck = false;
    }
    else if (type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_3STATE || type == BS_AUTO3STATE || type == BS_RADIOBUTTON ||
             type == BS_AUTORADIOBUTTON)
    {
        useBmCheck = true;
    }

    if (useBmCheck)
    {
        SendMessageW(toggle, BM_SETCHECK, toggledOn ? BST_CHECKED : BST_UNCHECKED, 0);
        return;
    }

    SetWindowLongPtrW(toggle, GWLP_USERDATA, toggledOn ? 1 : 0);
    InvalidateRect(toggle, nullptr, TRUE);
}

[[nodiscard]] bool GetTwoStateToggleState(HWND toggle, bool highContrast) noexcept
{
    if (! toggle)
    {
        return false;
    }

    const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    const UINT type      = static_cast<UINT>(style & BS_TYPEMASK);
    bool useBmCheck      = highContrast;
    if (type == BS_OWNERDRAW)
    {
        useBmCheck = false;
    }
    else if (type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_3STATE || type == BS_AUTO3STATE || type == BS_RADIOBUTTON ||
             type == BS_AUTORADIOBUTTON)
    {
        useBmCheck = true;
    }

    if (useBmCheck)
    {
        return SendMessageW(toggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
    }

    return GetWindowLongPtrW(toggle, GWLP_USERDATA) != 0;
}

[[nodiscard]] std::wstring_view LoadStringResourceView(_In_opt_ HINSTANCE hInstance, _In_ UINT uID) noexcept
{
    const HINSTANCE instance = hInstance ? hInstance : GetModuleHandleW(nullptr);
    PCWSTR ptr               = nullptr;
    const int length         = ::LoadStringW(instance, uID, reinterpret_cast<LPWSTR>(&ptr), 0);
    if (length <= 0 || ! ptr)
    {
        return {};
    }

    return std::wstring_view(ptr, static_cast<size_t>(length));
}

struct CompareDetailsTextStrings
{
    std::wstring_view identical;
    std::wstring_view onlyInLeft;
    std::wstring_view onlyInRight;
    std::wstring_view typeMismatch;
    std::wstring_view bigger;
    std::wstring_view smaller;
    std::wstring_view newer;
    std::wstring_view older;
    std::wstring_view attributesDiffer;
    std::wstring_view contentDiffer;
    std::wstring_view subdirAttributesDiffer;
    std::wstring_view subdirContentDiffer;
};

[[nodiscard]] const CompareDetailsTextStrings& GetCompareDetailsTextStrings() noexcept
{
    static const CompareDetailsTextStrings strings = []() noexcept
    {
        CompareDetailsTextStrings value{};
        value.identical               = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_IDENTICAL);
        value.onlyInLeft              = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_ONLY_IN_LEFT);
        value.onlyInRight             = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_ONLY_IN_RIGHT);
        value.typeMismatch            = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_TYPE_MISMATCH);
        value.bigger                  = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_BIGGER);
        value.smaller                 = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_SMALLER);
        value.newer                   = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_NEWER);
        value.older                   = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_OLDER);
        value.attributesDiffer        = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_ATTRIBUTES_DIFFER);
        value.contentDiffer           = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_CONTENT_DIFFER);
        value.subdirAttributesDiffer  = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_SUBDIR_ATTRIBUTES_DIFFER);
        value.subdirContentDiffer     = LoadStringResourceView(nullptr, IDS_COMPARE_DETAILS_SUBDIR_CONTENT_DIFFER);
        return value;
    }();

    return strings;
}

class CompareDirectoriesWindow final
{
public:
    CompareDirectoriesWindow(Common::Settings::Settings& settings,
                             AppTheme theme,
                             const ShortcutManager* shortcuts,
                             wil::com_ptr<IFileSystem> baseFileSystem,
    std::filesystem::path leftRoot,
                              std::filesystem::path rightRoot) noexcept;

    [[nodiscard]] bool Create(HWND owner) noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;

    CompareDirectoriesWindow(const CompareDirectoriesWindow&)            = delete;
    CompareDirectoriesWindow& operator=(const CompareDirectoriesWindow&) = delete;
    CompareDirectoriesWindow(CompareDirectoriesWindow&&)                 = delete;
    CompareDirectoriesWindow& operator=(CompareDirectoriesWindow&&)      = delete;

  private:
     static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
     static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
     LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
 
     friend LRESULT CALLBACK CompareOptionsHostSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
     friend LRESULT CALLBACK CompareOptionsWheelRouteSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
 
     bool OnCreate(HWND hwnd) noexcept;
     void OnDestroy() noexcept;
     void OnNcDestroy() noexcept;
     void OnSize() noexcept;
    void OnDpiChanged(UINT newDpi, const RECT* newRect) noexcept;
    void OnCommand(UINT id) noexcept;
    LRESULT OnFunctionBarInvoke(WPARAM wParam, LPARAM lParam) noexcept;
    void OnPaint() noexcept;
     LRESULT OnCtlColorStatic(HDC hdc, HWND control) noexcept;
     void PrepareThemedMenu() noexcept;
     void PrepareThemedMenuRecursive(HMENU menu,
                                     bool topLevel,
                                     std::vector<std::unique_ptr<CompareMenuItemData>>& itemData) noexcept;
     void OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept;
     void OnDrawItem(DRAWITEMSTRUCT* dis) noexcept;
     void ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept;

    void OnLButtonDown(POINT pt) noexcept;
    void OnLButtonDblClk(POINT pt) noexcept;
    void OnLButtonUp() noexcept;
    void OnMouseMove(POINT pt) noexcept;
    void OnCaptureChanged() noexcept;
    [[nodiscard]] bool OnSetCursor(POINT pt) noexcept;
    void SetSplitRatio(float ratio) noexcept;

    void ApplyTheme() noexcept;
    void ApplyOptionsDialogTheme() noexcept;
    void CreateChildWindows(HWND hwnd) noexcept;
    void EnsureOptionsControlsCreated(HWND dlg) noexcept;
    void LayoutOptionsControls() noexcept;
    void PaintOptionsHostBackgroundAndCards(HDC hdc, HWND host) noexcept;
    void Layout() noexcept;

    void EnsureCompareSession() noexcept;
    void StartCompare() noexcept;
    void ShowOptionsPanel(bool show) noexcept;

    void OnPanePathChanged(ComparePane pane, const std::optional<std::filesystem::path>& newPath) noexcept;
    void SyncOtherPanePath(ComparePane changedPane, const std::optional<std::filesystem::path>& newPath) noexcept;
    void ApplySelectionForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept;
    [[nodiscard]] std::wstring BuildDetailsTextForCompareItem(ComparePane pane,
                                                              const std::filesystem::path& folder,
                                                              std::wstring_view displayName,
                                                              bool isDirectory,
                                                              uint64_t sizeBytes,
                                                              int64_t lastWriteTime,
                                                              DWORD fileAttributes) noexcept;
     void OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept;
     LRESULT OnScanProgress(LPARAM lp) noexcept;
     LRESULT OnExecuteShortcutCommand(LPARAM lp) noexcept;
     void ExecuteShortcutCommand(std::wstring_view commandId) noexcept;

    Common::Settings::CompareDirectoriesSettings GetEffectiveCompareSettings() const noexcept;
     void LoadOptionsControlsFromSettings() noexcept;
     void SaveOptionsControlsToSettings() noexcept;
     void UpdateOptionsVisibility() noexcept;
     void RefreshBothPanes() noexcept;

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _optionsDlg;
    wil::unique_hwnd _scanProgressText;
    wil::unique_hwnd _scanProgressBar;
     wil::unique_hwnd _bannerTitle;
     wil::unique_hwnd _bannerOptionsButton;
     wil::unique_hwnd _bannerRescanButton;

    struct OptionsToggleCard
    {
        HWND title       = nullptr;
        HWND description = nullptr;
        HWND toggle      = nullptr;
    };

    struct OptionsIgnoreCard
    {
        HWND title       = nullptr;
        HWND description = nullptr;
        HWND toggle      = nullptr;
        HWND frame       = nullptr;
        HWND edit        = nullptr;
    };

    struct OptionsUi
    {
        HWND host = nullptr;

        HWND headerCompare  = nullptr;
        HWND headerSubdirs  = nullptr;
        HWND headerAdvanced = nullptr;
        HWND headerIgnore   = nullptr;

        OptionsToggleCard compareSize;
        OptionsToggleCard compareDateTime;
        OptionsToggleCard compareAttributes;
        OptionsToggleCard compareContent;
        OptionsToggleCard compareSubdirectories;

        OptionsToggleCard compareSubdirAttributes;
        OptionsToggleCard selectSubdirsOnlyInOnePane;

        OptionsIgnoreCard ignoreFiles;
        OptionsIgnoreCard ignoreDirectories;
    };

    OptionsUi _optionsUi{};
    std::vector<RECT> _optionsCards;
    int _optionsScrollOffset = 0;
    int _optionsScrollMax    = 0;
    int _optionsWheelRemainder = 0;

    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    const ShortcutManager* _shortcuts = nullptr;
    wil::com_ptr<IFileSystem> _baseFs;
    std::filesystem::path _leftRoot;
    std::filesystem::path _rightRoot;

    std::shared_ptr<CompareDirectoriesSession> _session;
    wil::com_ptr<IFileSystem> _fsLeft;
    wil::com_ptr<IFileSystem> _fsRight;

    FolderWindow _folderWindow;

    struct DetailsDecisionCache
    {
        std::filesystem::path folder;
        uint64_t sessionUiVersion = 0;
        std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    };

    DetailsDecisionCache _detailsCacheLeft;
    DetailsDecisionCache _detailsCacheRight;

    enum class CompareDetailsVerbosity
    {
        StatusOnly,
        StatusAndMetadata,
    };
    CompareDetailsVerbosity _detailsVerbosity = CompareDetailsVerbosity::StatusOnly;

    // Layout
    SIZE _clientSize{};
    RECT _splitterRect{};
    float _splitRatio = 0.5f;
    bool _draggingSplitter    = false;
    int _splitterDragOffsetPx = 0;

     wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _uiFont;
     wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _uiBoldFont;
     wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _uiItalicFont;
     wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _bannerTitleFont;
     wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _backgroundBrush;
     wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _splitterBrush;
     wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _splitterGripBrush;
     wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _menuBackgroundBrush;
     wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _optionsBackgroundBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _optionsCardBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _optionsInputBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _optionsInputFocusedBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _optionsInputDisabledBrush;

    COLORREF _optionsInputBackgroundColor        = RGB(255, 255, 255);
    COLORREF _optionsInputFocusedBackgroundColor = RGB(255, 255, 255);
    COLORREF _optionsInputDisabledBackgroundColor = RGB(255, 255, 255);
    ThemedInputFrames::FrameStyle _optionsFrameStyle{};

    std::vector<std::unique_ptr<CompareMenuItemData>> _menuItemData;
    std::vector<std::unique_ptr<CompareMenuItemData>> _popupMenuItemData;

    bool _compareStarted = false;
    bool _syncingPaths   = false;
    UINT _dpi            = USER_DEFAULT_SCREEN_DPI;
    int _restoreShowCmd  = SW_SHOWNORMAL;
    bool _hasSavedPlacement = false;
};

CompareDirectoriesWindow::CompareDirectoriesWindow(Common::Settings::Settings& settings,
                                                   AppTheme theme,
                                                   const ShortcutManager* shortcuts,
                                                   wil::com_ptr<IFileSystem> baseFileSystem,
                                                   std::filesystem::path leftRoot,
                                                   std::filesystem::path rightRoot) noexcept :
    _settings(&settings),
    _theme(std::move(theme)),
    _shortcuts(shortcuts),
    _baseFs(std::move(baseFileSystem)),
    _leftRoot(std::move(leftRoot)),
    _rightRoot(std::move(rightRoot))
{
}

ATOM CompareDirectoriesWindow::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kCompareDirectoriesWindowClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK CompareDirectoriesWindow::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self     = cs ? reinterpret_cast<CompareDirectoriesWindow*>(cs->lpCreateParams) : nullptr;
        if (self)
        {
            self->_hWnd.reset(hwnd);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    return self->WndProc(hwnd, msg, wp, lp);
}

LRESULT CompareDirectoriesWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: OnNcDestroy(); return 0;
        case WM_SIZE: OnSize(); return 0;
        case WM_DPICHANGED: OnDpiChanged(HIWORD(wp), reinterpret_cast<RECT*>(lp)); return 0;
        case WM_COMMAND: OnCommand(LOWORD(wp)); return 0;
        case WndMsg::kFunctionBarInvoke: return OnFunctionBarInvoke(wp, lp);
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_ACTIVATE:
            if (_hWnd)
            {
                const bool windowActive = LOWORD(wp) != WA_INACTIVE;
                ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive);
            }
            return 0;
        case WM_MEASUREITEM: OnMeasureItem(reinterpret_cast<MEASUREITEMSTRUCT*>(lp)); return TRUE;
        case WM_DRAWITEM: OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp)); return TRUE;
        case WM_CTLCOLORSTATIC:
        {
            const LRESULT result = OnCtlColorStatic(reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
            if (result != 0)
            {
                return result;
            }
            break;
        }
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONUP: OnLButtonUp(); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(); return 0;
        case WM_SETCURSOR:
        {
            POINT pt{};
            if (GetCursorPos(&pt))
            {
                ScreenToClient(hwnd, &pt);
                if (OnSetCursor(pt))
                {
                    return TRUE;
                }
            }
            break;
        }
        case WndMsg::kCompareDirectoriesScanProgress: return OnScanProgress(lp);
        case WndMsg::kCompareDirectoriesExecuteCommand: return OnExecuteShortcutCommand(lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

bool CompareDirectoriesWindow::Create(HWND owner) noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    if (! RegisterWndClass(instance))
    {
        return false;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_COMPARE_DIRECTORIES_TITLE);

    _hasSavedPlacement = _settings && _settings->windows.find(std::wstring(kCompareDirectoriesWindowId)) != _settings->windows.end();

    HWND placementOwner = owner;
    if (placementOwner && IsWindow(placementOwner))
    {
        placementOwner = GetAncestor(placementOwner, GA_ROOT);
    }
    else
    {
        placementOwner = nullptr;
    }

    wil::unique_hmenu menu(LoadMenuW(instance, MAKEINTRESOURCEW(IDR_COMPARE_DIRECTORIES_MENU)));

    int x = CW_USEDEFAULT;
    int y = CW_USEDEFAULT;
    int w = 1100;
    int h = 700;
    if (! _hasSavedPlacement && placementOwner)
    {
        WINDOWPLACEMENT wp{};
        wp.length = sizeof(wp);
        if (GetWindowPlacement(placementOwner, &wp) != 0)
        {
            const RECT rc = wp.rcNormalPosition;
            x             = rc.left;
            y             = rc.top;
            w             = std::max(1, static_cast<int>(rc.right - rc.left));
            h             = std::max(1, static_cast<int>(rc.bottom - rc.top));
        }
        else
        {
            RECT rc{};
            if (GetWindowRect(placementOwner, &rc) != 0)
            {
                x = rc.left;
                y = rc.top;
                w = std::max(1, static_cast<int>(rc.right - rc.left));
                h = std::max(1, static_cast<int>(rc.bottom - rc.top));
            }
        }

        _restoreShowCmd = IsZoomed(placementOwner) ? SW_MAXIMIZE : SW_SHOWNORMAL;
    }

    HWND created = CreateWindowExW(0,
                                   kCompareDirectoriesWindowClassName,
                                   title.c_str(),
                                   WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                   x,
                                   y,
                                   w,
                                   h,
                                   nullptr,
                                   menu.get(),
                                   instance,
                                   this);
    if (! created)
    {
        return false;
    }

    if (menu)
    {
        menu.release();
    }

    ShowWindow(created, _restoreShowCmd);
    UpdateWindow(created);
    return true;
}

bool CompareDirectoriesWindow::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    g_compareDirectoriesWindows.push_back(this);
    if (_settings && _hasSavedPlacement)
    {
        _restoreShowCmd = WindowPlacementPersistence::Restore(*_settings, kCompareDirectoriesWindowId, hwnd);
    }

    if (HMENU menu = GetMenu(hwnd))
    {
        const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
        CheckMenuItem(menu,
                      IDM_COMPARE_TOGGLE_IDENTICAL,
                      static_cast<UINT>(MF_BYCOMMAND | (s.showIdenticalItems ? MF_CHECKED : MF_UNCHECKED)));
    }

    ApplyTheme();
    CreateChildWindows(hwnd);
    ApplyTheme();
    Layout();
    ShowOptionsPanel(true);
    return true;
}

void CompareDirectoriesWindow::OnDestroy() noexcept
{
    if (_settings && _hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kCompareDirectoriesWindowId, _hWnd.get());
    }

    if (_session)
    {
        _session->SetScanProgressCallback({});
    }

    _optionsUi           = {};
    _optionsCards.clear();
    _optionsScrollOffset = 0;
    _optionsScrollMax    = 0;

    _optionsDlg.reset();
    _scanProgressText.reset();
    _scanProgressBar.reset();
    _bannerTitle.reset();
    _bannerOptionsButton.reset();
    _bannerRescanButton.reset();
    _folderWindow.Destroy();
}

void CompareDirectoriesWindow::OnNcDestroy() noexcept
{
    std::erase(g_compareDirectoriesWindows, this);

    if (_hWnd)
    {
        static_cast<void>(DrainPostedPayloadsForWindow(_hWnd.get()));
        SetWindowLongPtrW(_hWnd.get(), GWLP_USERDATA, 0);
        _hWnd.release();
    }
    delete this;
}

void CompareDirectoriesWindow::OnSize() noexcept
{
    Layout();
}

void CompareDirectoriesWindow::OnDpiChanged(UINT newDpi, const RECT* newRect) noexcept
{
    _dpi = newDpi;

    if (newRect && _hWnd)
    {
        SetWindowPos(_hWnd.get(),
                     nullptr,
                     newRect->left,
                     newRect->top,
                     newRect->right - newRect->left,
                     newRect->bottom - newRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    _folderWindow.OnDpiChanged(static_cast<float>(_dpi));
    ApplyTheme();
    Layout();
}

void CompareDirectoriesWindow::OnCommand(UINT id) noexcept
{
    switch (id)
    {
        case IDM_VIEW_SWITCH_PANE_FOCUS:
        {
            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);
            if (const HWND view = _folderWindow.GetFolderViewHwnd(pane))
            {
                SendMessageW(view, WM_KEYDOWN, static_cast<WPARAM>(VK_TAB), 0);
            }
            break;
        }
        case IDM_PANE_RENAME:
        case IDM_PANE_VIEW:
        case IDM_PANE_VIEW_SPACE:
        case IDM_PANE_COPY_TO_OTHER:
        case IDM_PANE_MOVE_TO_OTHER:
        case IDM_PANE_CREATE_DIR:
        case IDM_PANE_DELETE:
        case IDM_PANE_PERMANENT_DELETE:
        case IDM_PANE_PERMANENT_DELETE_WITH_VALIDATION:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);

            switch (id)
            {
                case IDM_PANE_RENAME: _folderWindow.CommandRename(pane); break;
                case IDM_PANE_VIEW: _folderWindow.CommandView(pane); break;
                case IDM_PANE_VIEW_SPACE: _folderWindow.CommandViewSpace(pane); break;
                case IDM_PANE_COPY_TO_OTHER: _folderWindow.CommandCopyToOtherPane(pane); break;
                case IDM_PANE_MOVE_TO_OTHER: _folderWindow.CommandMoveToOtherPane(pane); break;
                case IDM_PANE_CREATE_DIR: _folderWindow.CommandCreateDirectory(pane); break;
                case IDM_PANE_DELETE: _folderWindow.CommandDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE: _folderWindow.CommandPermanentDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE_WITH_VALIDATION: _folderWindow.CommandPermanentDeleteWithValidation(pane); break;
            }
            break;
        }
        case IDM_LEFT_SORT_NAME:
        case IDM_LEFT_SORT_EXTENSION:
        case IDM_LEFT_SORT_TIME:
        case IDM_LEFT_SORT_SIZE:
        case IDM_LEFT_SORT_ATTRIBUTES:
        case IDM_RIGHT_SORT_NAME:
        case IDM_RIGHT_SORT_EXTENSION:
        case IDM_RIGHT_SORT_TIME:
        case IDM_RIGHT_SORT_SIZE:
        case IDM_RIGHT_SORT_ATTRIBUTES:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id >= IDM_RIGHT_SORT_NAME ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
            _folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (id)
            {
                case IDM_LEFT_SORT_NAME:
                case IDM_RIGHT_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_LEFT_SORT_EXTENSION:
                case IDM_RIGHT_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_LEFT_SORT_TIME:
                case IDM_RIGHT_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_LEFT_SORT_SIZE:
                case IDM_RIGHT_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_LEFT_SORT_ATTRIBUTES:
                case IDM_RIGHT_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            _folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_LEFT_SORT_NONE:
        case IDM_RIGHT_SORT_NONE:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id == IDM_RIGHT_SORT_NONE ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
            _folderWindow.SetActivePane(pane);
            _folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_SORT_NAME:
        case IDM_PANE_SORT_EXTENSION:
        case IDM_PANE_SORT_TIME:
        case IDM_PANE_SORT_SIZE:
        case IDM_PANE_SORT_ATTRIBUTES:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (id)
            {
                case IDM_PANE_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_PANE_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_PANE_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_PANE_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_PANE_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            _folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_PANE_SORT_NONE:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
            _folderWindow.SetActivePane(pane);
            _folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_DISPLAY_BRIEF:
        case IDM_PANE_DISPLAY_DETAILED:
        {
            if (! _compareStarted)
            {
                break;
            }

            _detailsVerbosity = id == IDM_PANE_DISPLAY_DETAILED ? CompareDetailsVerbosity::StatusAndMetadata : CompareDetailsVerbosity::StatusOnly;

            // Compare view always stays in a 2-line layout; "Brief" maps to status-only details
            // and "Detailed" maps to status + metadata.
            _folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
            _folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Detailed);
            _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Left);
            _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Right);
            break;
        }
        case IDM_LEFT_REFRESH:
        case IDM_RIGHT_REFRESH:
        {
            if (! _compareStarted)
            {
                break;
            }

            const FolderWindow::Pane pane = id == IDM_LEFT_REFRESH ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
            _folderWindow.SetActivePane(pane);
            _folderWindow.CommandRefresh(pane);
            break;
        }
        case IDM_COMPARE_OPTIONS: ShowOptionsPanel(true); break;
        case IDM_COMPARE_RESCAN:
            if (_session)
            {
                _session->Invalidate();
            }
            RefreshBothPanes();
            break;
        case IDM_COMPARE_TOGGLE_IDENTICAL:
        {
            if (! _settings)
            {
                break;
            }

            Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
            s.showIdenticalItems                           = ! s.showIdenticalItems;
            _settings->compareDirectories                  = s;
            if (_session)
            {
                _session->SetSettings(s);
            }

            if (_hWnd)
            {
                if (HMENU menu = GetMenu(_hWnd.get()))
                {
                    CheckMenuItem(menu,
                                  IDM_COMPARE_TOGGLE_IDENTICAL,
                                  static_cast<UINT>(MF_BYCOMMAND | (s.showIdenticalItems ? MF_CHECKED : MF_UNCHECKED)));
                }
            }

            RefreshBothPanes();
            break;
        }
        case IDM_COMPARE_RESTORE_DIFFERENCES_SELECTION:
        {
            if (! _compareStarted)
            {
                break;
            }

            if (const auto leftPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Left))
            {
                ApplySelectionForFolder(ComparePane::Left, leftPath.value());
            }
            if (const auto rightPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Right))
            {
                ApplySelectionForFolder(ComparePane::Right, rightPath.value());
            }
            break;
        }
        case IDM_COMPARE_INVERT_DIFFERENCES_SELECTION:
        {
            if (! _compareStarted || ! _session)
            {
                break;
            }

            auto invertForPane = [&](ComparePane pane, FolderWindow::Pane folderWindowPane) noexcept
            {
                const auto folderOpt = _folderWindow.GetCurrentPath(folderWindowPane);
                if (! folderOpt.has_value())
                {
                    return;
                }

                const auto relOpt = _session->TryMakeRelative(pane, folderOpt.value());
                if (! relOpt.has_value())
                {
                    return;
                }

                const auto decision = _session->GetOrComputeDecision(relOpt.value());
                if (! decision || FAILED(decision->hr))
                {
                    return;
                }

                const bool isLeft = pane == ComparePane::Left;
                auto shouldSelect = [&](std::wstring_view name) noexcept -> bool
                {
                    const auto it = decision->items.find(name);
                    if (it == decision->items.end())
                    {
                        return false;
                    }

                    const bool selected = isLeft ? it->second.selectLeft : it->second.selectRight;
                    return ! selected;
                };

                _folderWindow.SetPaneSelectionByDisplayNamePredicate(folderWindowPane, shouldSelect, true);
            };

            invertForPane(ComparePane::Left, FolderWindow::Pane::Left);
            invertForPane(ComparePane::Right, FolderWindow::Pane::Right);
            break;
        }
        case IDM_COMPARE_CLOSE:
            if (_hWnd)
            {
                PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
            }
            break;
    }
}

LRESULT CompareDirectoriesWindow::OnFunctionBarInvoke(WPARAM wParam, LPARAM lParam) noexcept
{
    if (! _hWnd || ! _shortcuts)
    {
        return 0;
    }

    const uint32_t vk        = static_cast<uint32_t>(wParam);
    const uint32_t modifiers = static_cast<uint32_t>(lParam) & 0x7u;

    const std::optional<std::wstring_view> commandOpt = _shortcuts->FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return 0;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (commandId.starts_with(L"cmd/app/"))
    {
        // App-scoped commands are handled by the main window's message loop.
        return 0;
    }

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (! wmCommandOpt.has_value())
    {
        ExecuteShortcutCommand(commandId);
        return 0;
    }

    const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
    SendMessageW(_hWnd.get(), WM_COMMAND, wp, 0);
    return 0;
}

void CompareDirectoriesWindow::OnPaint() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);

    HBRUSH bg = _backgroundBrush ? _backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc.get(), &ps.rcPaint, bg);

    if (_splitterBrush)
    {
        RECT splitter = _splitterRect;
        RECT paint    = ps.rcPaint;
        RECT intersect{};
        if (IntersectRect(&intersect, &splitter, &paint))
        {
            FillRect(hdc.get(), &intersect, _splitterBrush.get());

            if (_splitterGripBrush)
            {
                const int dpi            = static_cast<int>(_dpi);
                const int dotSize        = std::max(1, MulDiv(kSplitterGripDotSizeDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int dotGap         = std::max(1, MulDiv(kSplitterGripDotGapDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int gripHeight     = (dotSize * kSplitterGripDotCount) + (dotGap * (kSplitterGripDotCount - 1));
                const int splitterWidth  = splitter.right - splitter.left;
                const int splitterHeight = splitter.bottom - splitter.top;

                if (splitterWidth > 0 && splitterHeight >= gripHeight)
                {
                    const int left = splitter.left + (splitterWidth - dotSize) / 2;
                    const int top  = splitter.top + (splitterHeight - gripHeight) / 2;

                    for (int i = 0; i < kSplitterGripDotCount; ++i)
                    {
                        const int dotTop = top + i * (dotSize + dotGap);
                        RECT dotRect{};
                        dotRect.left   = left;
                        dotRect.top    = dotTop;
                        dotRect.right  = left + dotSize;
                        dotRect.bottom = dotTop + dotSize;
                        FillRect(hdc.get(), &dotRect, _splitterGripBrush.get());
                    }
                }
            }
        }
    }
}

void CompareDirectoriesWindow::ExecuteShortcutCommand(std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return;
    }

    const std::wstring_view originalCommandId = commandId;
    std::optional<wchar_t> driveRootLetter;
    {
        constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
        if (originalCommandId.starts_with(kGoDriveRootPrefix) && originalCommandId.size() > kGoDriveRootPrefix.size())
        {
            const wchar_t rawLetter = originalCommandId[kGoDriveRootPrefix.size()];
            if (std::iswalpha(static_cast<wint_t>(rawLetter)) != 0)
            {
                const wchar_t upper = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(rawLetter)));
                if (upper >= L'A' && upper <= L'Z')
                {
                    driveRootLetter = upper;
                }
            }
        }
    }

    commandId = CanonicalizeCommandId(commandId);

    if (commandId == L"cmd/pane/menu")
    {
        if (_hWnd)
        {
            SendMessageW(_hWnd.get(), WM_SYSCOMMAND, SC_KEYMENU, 0);
        }
        return;
    }

    const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
    _folderWindow.SetActivePane(pane);

    const auto sendKeyToPaneFolderView = [&](uint32_t vk) noexcept
    {
        if (const HWND view = _folderWindow.GetFolderViewHwnd(pane))
        {
            SendMessageW(view, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
        }
    };

    if (commandId == L"cmd/pane/focusAddressBar")
    {
        _folderWindow.CommandFocusAddressBar(pane);
        return;
    }
    if (commandId == L"cmd/pane/upOneDirectory")
    {
        sendKeyToPaneFolderView(VK_BACK);
        return;
    }
    if (commandId == L"cmd/pane/switchPaneFocus")
    {
        sendKeyToPaneFolderView(VK_TAB);
        return;
    }
    if (commandId == L"cmd/pane/zoomPanel")
    {
        _folderWindow.ToggleZoomPanel(pane);
        return;
    }
    if (commandId == L"cmd/pane/refresh")
    {
        _folderWindow.CommandRefresh(pane);
        return;
    }
    if (commandId == L"cmd/pane/executeOpen")
    {
        sendKeyToPaneFolderView(VK_RETURN);
        return;
    }
    if (commandId == L"cmd/pane/selectCalculateDirectorySizeNext")
    {
        sendKeyToPaneFolderView(VK_SPACE);
        return;
    }
    if (commandId == L"cmd/pane/selectNext")
    {
        sendKeyToPaneFolderView(VK_INSERT);
        return;
    }
    if (commandId == L"cmd/pane/moveToRecycleBin")
    {
        sendKeyToPaneFolderView(VK_DELETE);
        return;
    }
    if (commandId == L"cmd/pane/goDriveRoot")
    {
        const auto getDefaultRoot = []() noexcept -> std::filesystem::path
        {
            wchar_t buffer[MAX_PATH] = {};
            const UINT bufferSize    = static_cast<UINT>(ARRAYSIZE(buffer));
            const UINT length        = GetWindowsDirectoryW(buffer, bufferSize);
            if (length > 0 && length < bufferSize)
            {
                const std::filesystem::path root = std::filesystem::path(buffer).root_path();
                if (! root.empty())
                {
                    return root;
                }
            }
            return std::filesystem::path(L"C:\\");
        };

        if (! driveRootLetter.has_value())
        {
            _folderWindow.SetFolderPath(pane, getDefaultRoot());
            return;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveRootLetter.value());
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_NO_ROOT_DIR)
        {
            return;
        }

        _folderWindow.SetFolderPath(pane, std::filesystem::path(driveRoot));
        return;
    }
}

LRESULT CompareDirectoriesWindow::OnCtlColorStatic(HDC hdc, HWND /*control*/) noexcept
{
    if (! hdc || ! _backgroundBrush)
    {
        return 0;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, _theme.menu.text);
    SetBkColor(hdc, _theme.windowBackground);
    return reinterpret_cast<LRESULT>(_backgroundBrush.get());
}

void CompareDirectoriesWindow::PrepareThemedMenu() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    HMENU menu = GetMenu(_hWnd.get());
    if (! menu || ! _menuBackgroundBrush)
    {
        return;
    }

    _menuItemData.clear();
    PrepareThemedMenuRecursive(menu, true, _menuItemData);
    DrawMenuBar(_hWnd.get());
}

void CompareDirectoriesWindow::PrepareThemedMenuRecursive(HMENU menu,
                                                          bool topLevel,
                                                          std::vector<std::unique_ptr<CompareMenuItemData>>& itemData) noexcept
{
    if (! menu || ! _menuBackgroundBrush)
    {
        return;
    }

    MENUINFO menuInfo{};
    menuInfo.cbSize  = sizeof(menuInfo);
    menuInfo.fMask   = MIM_BACKGROUND;
    menuInfo.hbrBack = _menuBackgroundBrush.get();
    SetMenuInfo(menu, &menuInfo);

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount < 0)
    {
        Debug::ErrorWithLastError(L"GetMenuItemCount failed");
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(itemCount); ++pos)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, pos, TRUE, &itemInfo))
        {
            continue;
        }

        auto data        = std::make_unique<CompareMenuItemData>();
        data->separator  = (itemInfo.fType & MFT_SEPARATOR) != 0;
        data->topLevel   = topLevel;
        data->hasSubMenu = itemInfo.hSubMenu != nullptr;

        if (! data->separator)
        {
            std::array<wchar_t, 512> buffer{};
            const int length = GetMenuStringW(menu, pos, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
            if (length > 0)
            {
                const std::wstring_view raw(buffer.data(), static_cast<size_t>(length));
                SplitMenuText(raw, data->text, data->shortcut);
            }
        }

        itemData.emplace_back(std::move(data));

        MENUITEMINFOW ownerDrawInfo{};
        ownerDrawInfo.cbSize     = sizeof(ownerDrawInfo);
        ownerDrawInfo.fMask      = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
        ownerDrawInfo.fType      = itemInfo.fType | MFT_OWNERDRAW;
        ownerDrawInfo.fState     = itemInfo.fState;
        ownerDrawInfo.dwItemData = reinterpret_cast<ULONG_PTR>(itemData.back().get());
        SetMenuItemInfoW(menu, pos, TRUE, &ownerDrawInfo);

        if (itemInfo.hSubMenu)
        {
            PrepareThemedMenuRecursive(itemInfo.hSubMenu, false, itemData);
        }
    }
}

void CompareDirectoriesWindow::ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    wil::unique_hmenu menu(CreatePopupMenu());
    if (! menu)
    {
        return;
    }

    const auto loadLabel = [](UINT stringId, std::wstring_view fallback) noexcept -> std::wstring
    {
        std::wstring text = LoadStringResource(nullptr, stringId);
        if (text.empty())
        {
            text.assign(fallback);
        }
        return text;
    };

    const std::wstring noneLabel       = loadLabel(IDS_PREFS_PANES_SORT_NONE, L"None");
    const std::wstring nameLabel       = loadLabel(IDS_PREFS_PANES_SORT_NAME, L"Name");
    const std::wstring extLabel        = loadLabel(IDS_PREFS_PANES_SORT_EXTENSION, L"Extension");
    const std::wstring timeLabel       = loadLabel(IDS_PREFS_PANES_SORT_TIME, L"Time");
    const std::wstring sizeLabel       = loadLabel(IDS_PREFS_PANES_SORT_SIZE, L"Size");
    const std::wstring attributesLabel = loadLabel(IDS_PREFS_PANES_SORT_ATTRIBUTES, L"Attributes");

    const bool isLeft = pane == FolderWindow::Pane::Left;
    const UINT idName = isLeft ? IDM_LEFT_SORT_NAME : IDM_RIGHT_SORT_NAME;
    const UINT idExt  = isLeft ? IDM_LEFT_SORT_EXTENSION : IDM_RIGHT_SORT_EXTENSION;
    const UINT idTime = isLeft ? IDM_LEFT_SORT_TIME : IDM_RIGHT_SORT_TIME;
    const UINT idSize = isLeft ? IDM_LEFT_SORT_SIZE : IDM_RIGHT_SORT_SIZE;
    const UINT idAttr = isLeft ? IDM_LEFT_SORT_ATTRIBUTES : IDM_RIGHT_SORT_ATTRIBUTES;
    const UINT idNone = isLeft ? IDM_LEFT_SORT_NONE : IDM_RIGHT_SORT_NONE;

    AppendMenuW(menu.get(), MF_STRING, idName, nameLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idExt, extLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idTime, timeLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idSize, sizeLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idAttr, attributesLabel.c_str());
    AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu.get(), MF_STRING, idNone, noneLabel.c_str());

    UINT checkId = idNone;
    switch (_folderWindow.GetSortBy(pane))
    {
        case FolderView::SortBy::Name: checkId = idName; break;
        case FolderView::SortBy::Extension: checkId = idExt; break;
        case FolderView::SortBy::Time: checkId = idTime; break;
        case FolderView::SortBy::Size: checkId = idSize; break;
        case FolderView::SortBy::Attributes: checkId = idAttr; break;
        case FolderView::SortBy::None: checkId = idNone; break;
    }
    CheckMenuItem(menu.get(), checkId, MF_BYCOMMAND | MF_CHECKED);

    if (_menuBackgroundBrush)
    {
        _popupMenuItemData.clear();
        PrepareThemedMenuRecursive(menu.get(), false, _popupMenuItemData);
    }

    SetForegroundWindow(_hWnd.get());
    TrackPopupMenu(menu.get(),
                   TPM_RIGHTALIGN | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON,
                   screenPoint.x,
                   screenPoint.y,
                   static_cast<int>(0u),
                   _hWnd.get(),
                   nullptr);
    PostMessageW(_hWnd.get(), WM_NULL, 0, 0);

    _popupMenuItemData.clear();
}

void CompareDirectoriesWindow::OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept
{
    if (! mis || mis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const CompareMenuItemData*>(mis->itemData);
    if (! data)
    {
        return;
    }

    const int dpi = static_cast<int>(_dpi);

    if (data->separator)
    {
        mis->itemWidth  = 1;
        mis->itemHeight = static_cast<UINT>(MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI));
        return;
    }

    const UINT heightDip = data->topLevel ? 20u : 24u;
    mis->itemHeight      = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), dpi, USER_DEFAULT_SCREEN_DPI));

    if (! _hWnd)
    {
        mis->itemWidth = data->topLevel ? 60 : 120;
        return;
    }

    auto hdc = wil::GetDC(_hWnd.get());
    if (! hdc)
    {
        mis->itemWidth = data->topLevel ? 60 : 120;
        return;
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), fontToUse);

    SIZE textSize{};
    if (! data->text.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->text.c_str(), static_cast<int>(data->text.size()), &textSize);
    }

    SIZE shortcutSize{};
    if (! data->shortcut.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);
    }

    const int paddingX       = MulDiv(5, dpi, USER_DEFAULT_SCREEN_DPI);
    const int shortcutGap    = MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth = data->topLevel ? 0 : MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);

    int width = paddingX + checkAreaWidth + textSize.cx + paddingX;
    if (! data->shortcut.empty())
    {
        width += shortcutGap + shortcutSize.cx;
    }

    mis->itemWidth = static_cast<UINT>(std::max(width, 60));
}

void CompareDirectoriesWindow::OnDrawItem(DRAWITEMSTRUCT* dis) noexcept
{
    if (! dis)
    {
        return;
    }

    if (dis->CtlType == ODT_BUTTON)
    {
        ThemedControls::DrawThemedPushButton(*dis, _theme);
        return;
    }

    if (dis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const CompareMenuItemData*>(dis->itemData);
    if (! data)
    {
        return;
    }

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const bool checked  = (dis->itemState & ODS_CHECKED) != 0;

    const COLORREF bgColor = selected ? _theme.menu.selectionBg : _theme.menu.background;

    COLORREF textColor = _theme.menu.text;
    if (selected)
    {
        textColor = _theme.menu.selectionText;
    }
    else if (disabled)
    {
        textColor = _theme.menu.disabledText;
    }

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(bgColor));
    FillRect(dis->hDC, &dis->rcItem, bgBrush.get());

    const int dpi      = static_cast<int>(_dpi);
    const int paddingX = MulDiv(5, dpi, USER_DEFAULT_SCREEN_DPI);

    if (data->separator)
    {
        const int y = (dis->rcItem.top + dis->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, _theme.menu.separator));
        [[maybe_unused]] auto oldPen = wil::SelectObject(dis->hDC, pen.get());
        MoveToEx(dis->hDC, dis->rcItem.left + paddingX, y, nullptr);
        LineTo(dis->hDC, dis->rcItem.right - paddingX, y);
        return;
    }

    const int shortcutGap    = MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth = data->topLevel ? 0 : MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);

    RECT checkRect = dis->rcItem;
    checkRect.left += paddingX;
    checkRect.right = std::min(checkRect.right, checkRect.left + checkAreaWidth);

    RECT textRect = dis->rcItem;
    textRect.left += paddingX + checkAreaWidth;
    textRect.right -= paddingX;

    SetBkMode(dis->hDC, TRANSPARENT);
    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, fontToUse);

    SetTextColor(dis->hDC, textColor);

    if (! data->topLevel && checked && checkRect.right > checkRect.left)
    {
        constexpr wchar_t kCheckMark[] = L"\u2713";
        DrawTextW(dis->hDC, kCheckMark, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    if (! data->shortcut.empty())
    {
        RECT shortcutRect = textRect;
        DrawTextW(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutRect, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX);

        SIZE shortcutSize{};
        GetTextExtentPoint32W(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);
        textRect.right = std::max(textRect.left, textRect.right - shortcutSize.cx - shortcutGap);
    }

    DWORD drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;
    drawFlags |= data->topLevel ? DT_CENTER : DT_LEFT;
    DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &textRect, drawFlags);
}

void CompareDirectoriesWindow::OnLButtonDown(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (PtInRect(&_splitterRect, pt))
    {
        _draggingSplitter     = true;
        _splitterDragOffsetPx = pt.x - _splitterRect.left;
        SetCapture(_hWnd.get());
    }
}

void CompareDirectoriesWindow::OnLButtonDblClk(POINT pt) noexcept
{
    if (! PtInRect(&_splitterRect, pt))
    {
        return;
    }

    _draggingSplitter = false;
    ReleaseCapture();
    SetSplitRatio(0.5f);
}

void CompareDirectoriesWindow::OnLButtonUp() noexcept
{
    if (_draggingSplitter)
    {
        _draggingSplitter = false;
        ReleaseCapture();
    }
}

void CompareDirectoriesWindow::OnMouseMove(POINT pt) noexcept
{
    if (! _draggingSplitter)
    {
        return;
    }

    const int splitterWidth  = _splitterRect.right - _splitterRect.left;
    const int availableWidth = std::max(0L, _clientSize.cx - splitterWidth);
    if (availableWidth <= 0)
    {
        return;
    }

    int desiredLeftWidth = pt.x - _splitterDragOffsetPx;
    desiredLeftWidth     = std::clamp(desiredLeftWidth, 0, availableWidth);

    const float ratio = static_cast<float>(desiredLeftWidth) / static_cast<float>(availableWidth);
    SetSplitRatio(ratio);

    if (_hWnd)
    {
        UpdateWindow(_hWnd.get());
    }
}

void CompareDirectoriesWindow::OnCaptureChanged() noexcept
{
    _draggingSplitter = false;
}

bool CompareDirectoriesWindow::OnSetCursor(POINT pt) noexcept
{
    if (PtInRect(&_splitterRect, pt))
    {
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        return true;
    }
    return false;
}

void CompareDirectoriesWindow::SetSplitRatio(float ratio) noexcept
{
    _splitRatio = std::clamp(ratio, kMinSplitRatio, kMaxSplitRatio);
    Layout();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void CompareDirectoriesWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    Layout();
}

void CompareDirectoriesWindow::ApplyTheme() noexcept
{
    _uiFont = CreateMenuFontForDpi(_dpi);
    _uiBoldFont.reset();
    _uiItalicFont.reset();
    _bannerTitleFont.reset();
    if (_uiFont)
    {
        LOGFONTW lf{};
        if (GetObjectW(_uiFont.get(), sizeof(lf), &lf) == sizeof(lf))
        {
            LOGFONTW bold = lf;
            bold.lfWeight = FW_SEMIBOLD;
            _uiBoldFont.reset(CreateFontIndirectW(&bold));

            LOGFONTW italic = lf;
            italic.lfItalic = TRUE;
            _uiItalicFont.reset(CreateFontIndirectW(&italic));

            // Slightly larger banner font for "Compare Folder" title (keep face/DPI scaling consistent).
            LOGFONTW banner = bold;
            const float bannerScale = 1.25f;
            banner.lfHeight         = static_cast<LONG>(std::lround(static_cast<float>(banner.lfHeight) * bannerScale));
            _bannerTitleFont.reset(CreateFontIndirectW(&banner));
        }
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    _menuBackgroundBrush.reset(CreateSolidBrush(_theme.menu.background));
    _optionsBackgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));

    const COLORREF surface                     = ThemedControls::GetControlSurfaceColor(_theme);
    _optionsCardBrush.reset(CreateSolidBrush(surface));
    _optionsInputBackgroundColor               = ThemedControls::BlendColor(surface, _theme.windowBackground, _theme.dark ? 50 : 30, 255);
    _optionsInputFocusedBackgroundColor        = ThemedControls::BlendColor(_optionsInputBackgroundColor, _theme.menu.text, _theme.dark ? 20 : 16, 255);
    _optionsInputDisabledBackgroundColor       = ThemedControls::BlendColor(_theme.windowBackground, _optionsInputBackgroundColor, _theme.dark ? 70 : 40, 255);
    _optionsInputBrush.reset(CreateSolidBrush(_optionsInputBackgroundColor));
    _optionsInputFocusedBrush.reset(CreateSolidBrush(_optionsInputFocusedBackgroundColor));
    _optionsInputDisabledBrush.reset(CreateSolidBrush(_optionsInputDisabledBackgroundColor));

    _optionsFrameStyle.theme                        = &_theme;
    _optionsFrameStyle.backdropBrush                = _optionsCardBrush ? _optionsCardBrush.get() : _optionsBackgroundBrush.get();
    _optionsFrameStyle.inputBackgroundColor         = _optionsInputBackgroundColor;
    _optionsFrameStyle.inputFocusedBackgroundColor  = _optionsInputFocusedBackgroundColor;
    _optionsFrameStyle.inputDisabledBackgroundColor = _optionsInputDisabledBackgroundColor;

    if (_bannerTitle)
    {
        const HFONT bannerFont = _bannerTitleFont ? _bannerTitleFont.get() : (_uiBoldFont ? _uiBoldFont.get() : _uiFont.get());
        SendMessageW(_bannerTitle.get(), WM_SETFONT, reinterpret_cast<WPARAM>(bannerFont), TRUE);
    }
    if (_bannerOptionsButton)
    {
        SendMessageW(_bannerOptionsButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_bannerRescanButton)
    {
        SendMessageW(_bannerRescanButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_scanProgressText)
    {
        SendMessageW(_scanProgressText.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_scanProgressBar)
    {
        if (_theme.highContrast)
        {
            SetWindowTheme(_scanProgressBar.get(), L"", nullptr);
        }
        else if (_theme.dark)
        {
            SetWindowTheme(_scanProgressBar.get(), L"DarkMode_Explorer", nullptr);
        }
        else
        {
            SetWindowTheme(_scanProgressBar.get(), L"Explorer", nullptr);
        }
    }

    _folderWindow.ApplyTheme(_theme);

    const wchar_t* folderViewThemeName = L"";
    if (! _theme.highContrast)
    {
        folderViewThemeName = _theme.dark ? L"DarkMode_Explorer" : L"Explorer";
    }

    if (const HWND leftView = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left))
    {
        SetWindowTheme(leftView, folderViewThemeName, nullptr);
        SendMessageW(leftView, WM_THEMECHANGED, 0, 0);
    }
    if (const HWND rightView = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right))
    {
        SetWindowTheme(rightView, folderViewThemeName, nullptr);
        SendMessageW(rightView, WM_THEMECHANGED, 0, 0);
    }

    ApplyOptionsDialogTheme();

    if (_hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive);
        PrepareThemedMenu();
        RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
    }
}

void CompareDirectoriesWindow::ApplyOptionsDialogTheme() noexcept
{
    if (! _optionsDlg)
    {
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(_theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* themeName  = _theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

    SetWindowTheme(_optionsDlg.get(), themeName, nullptr);
    SendMessageW(_optionsDlg.get(), WM_THEMECHANGED, 0, 0);

    const HFONT font = _uiFont.get();
    if (font)
    {
        SendMessageW(_optionsDlg.get(), WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    struct EnumData
    {
        HFONT font = nullptr;
        const wchar_t* themeName = nullptr;
        HWND optionsHost = nullptr;
    };

    EnumData data{};
    data.font      = font;
    data.themeName = themeName;
    data.optionsHost = _optionsUi.host;

    EnumChildWindows(
        _optionsDlg.get(),
        [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto* data = reinterpret_cast<const EnumData*>(lParam);
            if (! data || ! child)
            {
                return TRUE;
            }

            if (data->font)
            {
                SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(data->font), FALSE);
            }

            if (data->themeName)
            {
                std::array<wchar_t, 32> className{};
                const int classLen = GetClassNameW(child, className.data(), static_cast<int>(className.size()));

                const wchar_t* appliedTheme = data->themeName;
                if (classLen > 0)
                {
                    if (_wcsicmp(className.data(), L"Static") == 0)
                    {
                        appliedTheme = child == data->optionsHost ? data->themeName : L"";
                    }
                    else if (_wcsicmp(className.data(), L"Button") == 0)
                    {
                        const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
                        const LONG_PTR type  = style & BS_TYPEMASK;
                        if (type == BS_GROUPBOX || type == BS_PUSHBUTTON || type == BS_DEFPUSHBUTTON)
                        {
                            appliedTheme = L"";
                        }
                    }
                }

                SetWindowTheme(child, appliedTheme, nullptr);
                SendMessageW(child, WM_THEMECHANGED, 0, 0);
            }

            return TRUE;
        },
        reinterpret_cast<LPARAM>(&data));

    RedrawWindow(_optionsDlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

void CompareDirectoriesWindow::OnPanePathChanged(ComparePane pane, const std::optional<std::filesystem::path>& newPath) noexcept
{
    SyncOtherPanePath(pane, newPath);
}

void CompareDirectoriesWindow::CreateChildWindows(HWND hwnd) noexcept
{
    FolderView::RegisterWndClass(GetModuleHandleW(nullptr));

    const HINSTANCE instance = GetModuleHandleW(nullptr);
    _bannerTitle.reset(CreateWindowExW(0,
                                       L"Static",
                                       L"Compare Folder",
                                       WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE | SS_NOPREFIX,
                                       0,
                                       0,
                                       10,
                                       10,
                                       hwnd,
                                       nullptr,
                                       instance,
                                       nullptr));

    _bannerOptionsButton.reset(CreateWindowExW(0,
                                               L"Button",
                                               L"Options...",
                                               WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                               0,
                                               0,
                                               10,
                                               10,
                                               hwnd,
                                               reinterpret_cast<HMENU>(IDM_COMPARE_OPTIONS),
                                               instance,
                                               nullptr));

    _bannerRescanButton.reset(CreateWindowExW(0,
                                              L"Button",
                                              L"Rescan",
                                              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                              0,
                                              0,
                                              10,
                                              10,
                                              hwnd,
                                              reinterpret_cast<HMENU>(IDM_COMPARE_RESCAN),
                                              instance,
                                              nullptr));

    if (! _theme.highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(hwnd, IDM_COMPARE_OPTIONS);
        ThemedControls::EnableOwnerDrawButton(hwnd, IDM_COMPARE_RESCAN);
    }

    _scanProgressText.reset(CreateWindowExW(0,
                                           L"Static",
                                           L"",
                                           WS_CHILD | SS_LEFT | SS_NOPREFIX | SS_PATHELLIPSIS,
                                           0,
                                           0,
                                           10,
                                           10,
                                           hwnd,
                                           reinterpret_cast<HMENU>(kScanProgressTextId),
                                           instance,
                                           nullptr));
    _scanProgressBar.reset(CreateWindowExW(0,
                                          PROGRESS_CLASSW,
                                          nullptr,
                                          WS_CHILD | PBS_MARQUEE,
                                          0,
                                          0,
                                          10,
                                          10,
                                          hwnd,
                                          reinterpret_cast<HMENU>(kScanProgressBarId),
                                          instance,
                                          nullptr));

    if (_scanProgressText)
    {
        ShowWindow(_scanProgressText.get(), SW_HIDE);
    }
    if (_scanProgressBar)
    {
        ShowWindow(_scanProgressBar.get(), SW_HIDE);
    }

    _folderWindow.Create(hwnd, 0, 0, 10, 10);
    _folderWindow.SetSettings(_settings);
    _folderWindow.SetShortcutManager(_shortcuts);
    _folderWindow.SetShowSortMenuCallback(
        [this](FolderWindow::Pane pane, POINT screenPoint) noexcept { ShowSortMenuPopup(pane, screenPoint); });

    bool functionBarVisible = true;
    if (_settings && _settings->mainMenu.has_value())
    {
        functionBarVisible = _settings->mainMenu->functionBarVisible;
    }
    _folderWindow.SetFunctionBarVisible(functionBarVisible);

    _folderWindow.SetPanePathChangedCallback(
        [this](FolderWindow::Pane pane, const std::optional<std::filesystem::path>& pluginPath)
        { OnPanePathChanged(pane == FolderWindow::Pane::Left ? ComparePane::Left : ComparePane::Right, pluginPath); });

    _folderWindow.SetPaneEnumerationCompletedCallback(
        FolderWindow::Pane::Left, [this](const std::filesystem::path& folder) { ApplySelectionForFolder(ComparePane::Left, folder); });
    _folderWindow.SetPaneEnumerationCompletedCallback(
        FolderWindow::Pane::Right, [this](const std::filesystem::path& folder) { ApplySelectionForFolder(ComparePane::Right, folder); });

    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Left,
                                            [this](const std::filesystem::path& folder,
                                                   std::wstring_view displayName,
                                                   bool isDirectory,
                                                   uint64_t sizeBytes,
                                                   int64_t lastWriteTime,
                                                   DWORD fileAttributes) noexcept -> std::wstring
                                            { return BuildDetailsTextForCompareItem(ComparePane::Left, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _folderWindow.SetPaneDetailsTextProvider(FolderWindow::Pane::Right,
                                            [this](const std::filesystem::path& folder,
                                                   std::wstring_view displayName,
                                                   bool isDirectory,
                                                   uint64_t sizeBytes,
                                                   int64_t lastWriteTime,
                                                   DWORD fileAttributes) noexcept -> std::wstring
                                            { return BuildDetailsTextForCompareItem(ComparePane::Right, folder, displayName, isDirectory, sizeBytes, lastWriteTime, fileAttributes); });

    _folderWindow.SetFileOperationCompletedCallback([this](const FolderWindow::FileOperationCompletedEvent& e) { OnFolderWindowFileOperationCompleted(e); });

#pragma warning(push)
    // pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    _optionsDlg.reset(CreateDialogParamW(GetModuleHandleW(nullptr),
                                         MAKEINTRESOURCEW(IDD_COMPARE_DIRECTORIES_OPTIONS),
                                         hwnd,
                                         [](HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) noexcept -> INT_PTR
                                         {
                                             auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetWindowLongPtrW(dlg, DWLP_USER));

                                             switch (msg)
                                             {
                                                  case WM_INITDIALOG:
                                                      self = reinterpret_cast<CompareDirectoriesWindow*>(lParam);
                                                      SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(self));
                                                      if (self)
                                                      {
                                                          const bool darkBackground =
                                                              ChooseContrastingTextColor(self->_theme.windowBackground) == RGB(255, 255, 255);
                                                          const wchar_t* themeName =
                                                              self->_theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

                                                          SetWindowTheme(dlg, themeName, nullptr);
                                                          SendMessageW(dlg, WM_THEMECHANGED, 0, 0);

                                                           if (! self->_theme.highContrast)
                                                           {
                                                               ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
                                                               ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
                                                           }
 
                                                           self->EnsureOptionsControlsCreated(dlg);
                                                       }
                                                       return TRUE;
                                                 case WM_ERASEBKGND:
                                                     if (! self || ! self->_optionsBackgroundBrush)
                                                     {
                                                         return FALSE;
                                                     }
                                                     {
                                                         RECT rc{};
                                                         GetClientRect(dlg, &rc);
                                                         FillRect(reinterpret_cast<HDC>(wParam), &rc, self->_optionsBackgroundBrush.get());
                                                         return TRUE;
                                                     }
                                                 case WM_COMMAND:
                                                     if (! self)
                                                     {
                                                         return FALSE;
                                                     }

                                                     {
                                                         const UINT controlId  = LOWORD(wParam);
                                                         const UINT notifyCode = HIWORD(wParam);
                                                         HWND hwndCtl          = reinterpret_cast<HWND>(lParam);

                                                         if (notifyCode == BN_CLICKED && hwndCtl)
                                                         {
                                                             const LONG_PTR style = GetWindowLongPtrW(hwndCtl, GWL_STYLE);
                                                             if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
                                                             {
                                                                 switch (controlId)
                                                                 {
                                                                     case IDC_CMP_SIZE:
                                                                     case IDC_CMP_DATETIME:
                                                                     case IDC_CMP_ATTRIBUTES:
                                                                     case IDC_CMP_CONTENT:
                                                                     case IDC_CMP_SUBDIRECTORIES:
                                                                     case IDC_CMP_SUBDIR_ATTRIBUTES:
                                                                     case IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE:
                                                                     case IDC_CMP_IGNORE_FILES:
                                                                     case IDC_CMP_IGNORE_DIRECTORIES:
                                                                     {
                                                                         const bool toggledOn = GetTwoStateToggleState(hwndCtl, false);
                                                                         SetTwoStateToggleState(hwndCtl, false, ! toggledOn);
                                                                         break;
                                                                     }
                                                                 }
                                                             }
                                                         }

                                                         switch (controlId)
                                                     {
                                                         case IDOK:
                                                             self->SaveOptionsControlsToSettings();
                                                             self->StartCompare();
                                                             return TRUE;
                                                          case IDCANCEL:
                                                              if (! self->_compareStarted)
                                                              {
                                                                  PostMessageW(self->_hWnd.get(), WM_CLOSE, 0, 0);
                                                                  return TRUE;
                                                              }
                                                              self->ShowOptionsPanel(false);
                                                              return TRUE;
                                                          case IDC_CMP_IGNORE_FILES:
                                                          case IDC_CMP_IGNORE_DIRECTORIES:
                                                              self->UpdateOptionsVisibility();
                                                              return TRUE;
                                                      }
                                                     }
                                                     break;
                                                 case WM_DRAWITEM:
                                                     if (! self)
                                                     {
                                                         return FALSE;
                                                     }
                                                     if (const auto* dis = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam); dis && dis->CtlType == ODT_BUTTON)
                                                     {
                                                         const LONG_PTR style = dis->hwndItem ? GetWindowLongPtrW(dis->hwndItem, GWL_STYLE) : 0;
                                                         if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
                                                         {
                                                             const UINT id = dis->CtlID;
                                                             const bool isToggle =
                                                                 id == IDC_CMP_SIZE || id == IDC_CMP_DATETIME || id == IDC_CMP_ATTRIBUTES || id == IDC_CMP_CONTENT ||
                                                                 id == IDC_CMP_SUBDIRECTORIES || id == IDC_CMP_SUBDIR_ATTRIBUTES ||
                                                                 id == IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE || id == IDC_CMP_IGNORE_FILES || id == IDC_CMP_IGNORE_DIRECTORIES;
                                                             if (isToggle)
                                                             {
                                                                 const bool toggledOn   = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
                                                                 const COLORREF surface = ThemedControls::GetControlSurfaceColor(self->_theme);
                                                                 const HFONT boldFont   = self->_uiBoldFont ? self->_uiBoldFont.get() : nullptr;
                                                                 const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
                                                                 const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
                                                                 ThemedControls::DrawThemedSwitchToggle(
                                                                     *dis, self->_theme, surface, boldFont, onLabel, offLabel, toggledOn);
                                                                 return TRUE;
                                                             }
                                                         }

                                                         ThemedControls::DrawThemedPushButton(*dis, self->_theme);
                                                         return TRUE;
                                                     }
                                                     break;
                                                 case WM_CTLCOLOREDIT:
                                                     if (! self || ! self->_optionsInputBrush)
                                                     {
                                                         return FALSE;
                                                     }
                                                     {
                                                         HDC hdc        = reinterpret_cast<HDC>(wParam);
                                                         HWND control   = reinterpret_cast<HWND>(lParam);
                                                         const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
                                                         const bool focused = enabled && control && GetFocus() == control;
                                                         const COLORREF bg  = enabled ? (focused ? self->_optionsInputFocusedBackgroundColor : self->_optionsInputBackgroundColor)
                                                                                      : self->_optionsInputDisabledBackgroundColor;

                                                         SetBkMode(hdc, OPAQUE);
                                                         SetBkColor(hdc, bg);
                                                         SetTextColor(hdc, enabled ? self->_theme.menu.text : self->_theme.menu.disabledText);

                                                         if (self->_theme.highContrast)
                                                         {
                                                             return reinterpret_cast<INT_PTR>(self->_optionsBackgroundBrush.get());
                                                         }

                                                         if (! enabled)
                                                         {
                                                             return reinterpret_cast<INT_PTR>(self->_optionsInputDisabledBrush.get());
                                                         }
                                                         return reinterpret_cast<INT_PTR>(
                                                             focused && self->_optionsInputFocusedBrush ? self->_optionsInputFocusedBrush.get() : self->_optionsInputBrush.get());
                                                     }
                                                case WM_CTLCOLORDLG:
                                                    if (! self || ! self->_optionsBackgroundBrush)
                                                    {
                                                        return FALSE;
                                                    }
                                                    {
                                                        HDC hdc = reinterpret_cast<HDC>(wParam);
                                                        SetBkMode(hdc, OPAQUE);
                                                        SetBkColor(hdc, self->_theme.windowBackground);
                                                        SetTextColor(hdc, self->_theme.menu.text);
                                                        return reinterpret_cast<INT_PTR>(self->_optionsBackgroundBrush.get());
                                                    }
                                                case WM_CTLCOLORSTATIC:
                                                    if (! self || ! self->_optionsBackgroundBrush)
                                                    {
                                                        return FALSE;
                                                    }

                                                    {
                                                        HDC hdc      = reinterpret_cast<HDC>(wParam);
                                                        HWND control = reinterpret_cast<HWND>(lParam);

                                                        COLORREF textColor = self->_theme.menu.text;
                                                        if (control && IsWindowEnabled(control) == FALSE)
                                                        {
                                                            textColor = self->_theme.menu.disabledText;
                                                        }

                                                        if (self->_theme.systemHighContrast || self->_theme.highContrast)
                                                        {
                                                            SetBkMode(hdc, OPAQUE);
                                                            SetBkColor(hdc, self->_theme.windowBackground);
                                                            SetTextColor(hdc, textColor);
                                                            return reinterpret_cast<INT_PTR>(self->_optionsBackgroundBrush.get());
                                                        }

                                                        SetBkMode(hdc, TRANSPARENT);
                                                        SetTextColor(hdc, textColor);
                                                        SetBkColor(hdc, self->_theme.windowBackground);

                                                        HBRUSH brush = self->_optionsBackgroundBrush.get();
                                                        if (control && self->_optionsUi.host && self->_optionsCardBrush && ! self->_optionsCards.empty())
                                                        {
                                                            RECT rc{};
                                                            if (GetWindowRect(control, &rc))
                                                            {
                                                                MapWindowPoints(nullptr, self->_optionsUi.host, reinterpret_cast<POINT*>(&rc), 2);
                                                                for (const RECT& card : self->_optionsCards)
                                                                {
                                                                    RECT intersect{};
                                                                    if (IntersectRect(&intersect, &card, &rc) != FALSE)
                                                                    {
                                                                        brush = self->_optionsCardBrush.get();
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        return reinterpret_cast<INT_PTR>(brush);
                                                    }
                                                case WM_CTLCOLORBTN:
                                                    if (! self || ! self->_optionsBackgroundBrush)
                                                    {
                                                        return FALSE;
                                                    }

                                                    {
                                                        HDC hdc      = reinterpret_cast<HDC>(wParam);
                                                        HWND control = reinterpret_cast<HWND>(lParam);

                                                        const LONG_PTR style = control ? GetWindowLongPtrW(control, GWL_STYLE) : 0;
                                                        const LONG_PTR type  = style & BS_TYPEMASK;

                                                        const bool themed = type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_RADIOBUTTON ||
                                                                            type == BS_AUTORADIOBUTTON || type == BS_3STATE || type == BS_AUTO3STATE || type == BS_GROUPBOX;
                                                        if (! themed)
                                                        {
                                                            return FALSE;
                                                        }

                                                        const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
                                                        SetBkMode(hdc, TRANSPARENT);
                                                        SetTextColor(hdc, enabled ? self->_theme.menu.text : self->_theme.menu.disabledText);
                                                        SetBkColor(hdc, self->_theme.windowBackground);
                                                        return reinterpret_cast<INT_PTR>(self->_optionsBackgroundBrush.get());
                                                    }
                                            }

                                            return FALSE;
                                        },
                                        reinterpret_cast<LPARAM>(this)));
#pragma warning(pop)

    if (_optionsDlg)
    {
        ShowWindow(_optionsDlg.get(), SW_HIDE);
        LoadOptionsControlsFromSettings();
        ApplyOptionsDialogTheme();
    }
}

LRESULT CALLBACK CompareOptionsHostSubclassProc(HWND hwnd,
                                                UINT msg,
                                                WPARAM wp,
                                                LPARAM lp,
                                                UINT_PTR subclassId,
                                                DWORD_PTR refData) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(refData);
    const HWND dlg = GetParent(hwnd);

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PRINTCLIENT:
        {
            if (self)
            {
                self->PaintOptionsHostBackgroundAndCards(reinterpret_cast<HDC>(wp), hwnd);
            }
            return 0;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc)
            {
                return 0;
            }

            RECT client{};
            GetClientRect(hwnd, &client);
            const int width  = std::max(0l, client.right - client.left);
            const int height = std::max(0l, client.bottom - client.top);

            wil::unique_hdc memDc;
            wil::unique_hbitmap memBmp;
            if (width > 0 && height > 0)
            {
                memDc.reset(CreateCompatibleDC(hdc.get()));
                memBmp.reset(CreateCompatibleBitmap(hdc.get(), width, height));
            }

            if (memDc && memBmp)
            {
                [[maybe_unused]] auto oldBmp = wil::SelectObject(memDc.get(), memBmp.get());
                if (self)
                {
                    self->PaintOptionsHostBackgroundAndCards(memDc.get(), hwnd);
                }
                BitBlt(hdc.get(), 0, 0, width, height, memDc.get(), 0, 0, SRCCOPY);
            }
            else if (self)
            {
                self->PaintOptionsHostBackgroundAndCards(hdc.get(), hwnd);
            }
            return 0;
        }
        case WM_VSCROLL:
        {
            if (! self || self->_optionsScrollMax <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            GetScrollInfo(hwnd, SB_VERT, &si);

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ThemedControls::ScaleDip(dpi, 24);

            int newPos = self->_optionsScrollOffset;
            switch (LOWORD(wp))
            {
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = self->_optionsScrollMax; break;
                case SB_LINEUP: newPos -= lineY; break;
                case SB_LINEDOWN: newPos += lineY; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_THUMBPOSITION: newPos = si.nPos; break;
                default: break;
            }

            newPos = std::clamp(newPos, 0, self->_optionsScrollMax);
            if (newPos != self->_optionsScrollOffset)
            {
                self->_optionsScrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! self || self->_optionsScrollMax <= 0)
            {
                break;
            }

            const int delta = GET_WHEEL_DELTA_WPARAM(wp);
            if (delta == 0)
            {
                return 0;
            }

            self->_optionsWheelRemainder += delta;
            const int notches = self->_optionsWheelRemainder / WHEEL_DELTA;
            self->_optionsWheelRemainder -= notches * WHEEL_DELTA;
            if (notches == 0)
            {
                return 0;
            }

            UINT linesPerNotch = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);
            if (linesPerNotch == 0)
            {
                return 0;
            }

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ThemedControls::ScaleDip(dpi, 32);

            int scrollDelta = 0;
            if (linesPerNotch == WHEEL_PAGESCROLL)
            {
                SCROLLINFO si{};
                si.cbSize = sizeof(si);
                si.fMask  = SIF_PAGE;
                GetScrollInfo(hwnd, SB_VERT, &si);
                scrollDelta = notches * static_cast<int>(si.nPage);
            }
            else
            {
                scrollDelta = notches * lineY * static_cast<int>(linesPerNotch);
            }

            const int newPos = std::clamp(self->_optionsScrollOffset - scrollDelta, 0, self->_optionsScrollMax);
            if (newPos != self->_optionsScrollOffset)
            {
                self->_optionsScrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_DRAWITEM:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORBTN:
            if (dlg)
            {
                return SendMessageW(dlg, msg, wp, lp);
            }
            break;
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, CompareOptionsHostSubclassProc, subclassId);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK CompareOptionsWheelRouteSubclassProc(HWND hwnd,
                                                     UINT msg,
                                                     WPARAM wp,
                                                     LPARAM lp,
                                                     UINT_PTR subclassId,
                                                     DWORD_PTR refData) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(refData);
    if (! self)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_MOUSEWHEEL:
        {
            if (! self->_optionsDlg || ! self->_optionsUi.host)
            {
                return 0;
            }

            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            RECT dlgRect{};
            if (GetWindowRect(self->_optionsDlg.get(), &dlgRect) == 0 || PtInRect(&dlgRect, ptScreen) == FALSE)
            {
                // Don't scroll the options dialog when the user is wheeling outside it.
                return 0;
            }

            RECT hostRect{};
            if (GetWindowRect(self->_optionsUi.host, &hostRect) == 0 || PtInRect(&hostRect, ptScreen) == FALSE)
            {
                // Only scroll when the wheel is over the options host area.
                return 0;
            }

            if (hwnd != self->_optionsUi.host)
            {
                SendMessageW(self->_optionsUi.host, msg, wp, lp);
                return 0;
            }

            break;
        }
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, CompareOptionsWheelRouteSubclassProc, subclassId);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void CompareDirectoriesWindow::EnsureOptionsControlsCreated(HWND dlg) noexcept
{
    if (! dlg || _optionsUi.host)
    {
        return;
    }

    const HINSTANCE instance = GetModuleHandleW(nullptr);

    _optionsUi.host = CreateWindowExW(WS_EX_CONTROLPARENT,
                                     L"Static",
                                     L"",
                                     WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                     0,
                                     0,
                                     10,
                                     10,
                                     dlg,
                                     nullptr,
                                     instance,
                                     nullptr);
    if (_optionsUi.host)
    {
        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsUi.host, hostTheme, nullptr);
        SendMessageW(_optionsUi.host, WM_THEMECHANGED, 0, 0);

#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(_optionsUi.host, CompareOptionsHostSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(this));
#pragma warning(pop)
    }

    if (! _optionsUi.host)
    {
        return;
    }

    constexpr DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    constexpr DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;

    const DWORD toggleStyle =
        static_cast<DWORD>(WS_CHILD | WS_VISIBLE | WS_TABSTOP | (_theme.highContrast ? BS_AUTOCHECKBOX : BS_OWNERDRAW));

    const auto makeStatic = [&](DWORD style) noexcept -> HWND
    { return CreateWindowExW(0, L"Static", L"", style, 0, 0, 10, 10, _optionsUi.host, nullptr, instance, nullptr); };

    const auto makeToggle = [&](int id) noexcept -> HWND
    {
        const HWND toggle = CreateWindowExW(0,
                                            L"Button",
                                            L"",
                                            toggleStyle,
                                            0,
                                            0,
                                            10,
                                            10,
                                            _optionsUi.host,
                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                                            instance,
                                            nullptr);
        if (toggle && ! _theme.highContrast)
        {
            ThemedControls::EnableOwnerDrawButton(_optionsUi.host, id);
        }
        return toggle;
    };

    const auto makeFramedEdit = [&](HWND& outFrame, HWND& outEdit, int editId) noexcept
    {
        outFrame = nullptr;
        outEdit  = nullptr;

        const bool customFrames = ! _theme.highContrast;
        if (customFrames)
        {
            outFrame =
                CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 10, 10, _optionsUi.host, nullptr, instance, nullptr);
        }

        DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
        editStyle |= ES_MULTILINE;
        editStyle &= ~ES_WANTRETURN;

        const DWORD editExStyle = customFrames ? 0 : WS_EX_CLIENTEDGE;
        outEdit                 = CreateWindowExW(editExStyle,
                                  L"Edit",
                                  L"",
                                  editStyle,
                                  0,
                                  0,
                                  10,
                                  10,
                                  _optionsUi.host,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(editId)),
                                  instance,
                                  nullptr);

        if (customFrames && outFrame && outEdit)
        {
            ThemedInputFrames::InstallFrame(outFrame, outEdit, &_optionsFrameStyle);
        }

        if (outEdit)
        {
            const UINT dpi       = GetDpiForWindow(outEdit);
            const int textMargin = ThemedControls::ScaleDip(dpi, 6);
            SendMessageW(outEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));
        }
    };

    _optionsUi.headerCompare  = makeStatic(baseStaticStyle);
    _optionsUi.headerSubdirs  = makeStatic(baseStaticStyle);
    _optionsUi.headerAdvanced = makeStatic(baseStaticStyle);
    _optionsUi.headerIgnore   = makeStatic(baseStaticStyle);

    _optionsUi.compareSize.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSize.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSize.toggle      = makeToggle(IDC_CMP_SIZE);

    _optionsUi.compareDateTime.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareDateTime.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareDateTime.toggle      = makeToggle(IDC_CMP_DATETIME);

    _optionsUi.compareAttributes.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareAttributes.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareAttributes.toggle      = makeToggle(IDC_CMP_ATTRIBUTES);

    _optionsUi.compareContent.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareContent.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareContent.toggle      = makeToggle(IDC_CMP_CONTENT);

    _optionsUi.compareSubdirectories.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSubdirectories.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSubdirectories.toggle      = makeToggle(IDC_CMP_SUBDIRECTORIES);

    _optionsUi.compareSubdirAttributes.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSubdirAttributes.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSubdirAttributes.toggle      = makeToggle(IDC_CMP_SUBDIR_ATTRIBUTES);

    _optionsUi.selectSubdirsOnlyInOnePane.title       = makeStatic(baseStaticStyle);
    _optionsUi.selectSubdirsOnlyInOnePane.description = makeStatic(wrapStaticStyle);
    _optionsUi.selectSubdirsOnlyInOnePane.toggle      = makeToggle(IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE);

    _optionsUi.ignoreFiles.title       = makeStatic(baseStaticStyle);
    _optionsUi.ignoreFiles.description = makeStatic(wrapStaticStyle);
    _optionsUi.ignoreFiles.toggle      = makeToggle(IDC_CMP_IGNORE_FILES);
    makeFramedEdit(_optionsUi.ignoreFiles.frame, _optionsUi.ignoreFiles.edit, IDC_CMP_IGNORE_FILES_PATTERNS);

    _optionsUi.ignoreDirectories.title       = makeStatic(baseStaticStyle);
    _optionsUi.ignoreDirectories.description = makeStatic(wrapStaticStyle);
    _optionsUi.ignoreDirectories.toggle      = makeToggle(IDC_CMP_IGNORE_DIRECTORIES);
    makeFramedEdit(_optionsUi.ignoreDirectories.frame, _optionsUi.ignoreDirectories.edit, IDC_CMP_IGNORE_DIRECTORIES_PATTERNS);

#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
    SetWindowSubclass(dlg, CompareOptionsWheelRouteSubclassProc, 2u, reinterpret_cast<DWORD_PTR>(this));
    EnumChildWindows(
        dlg,
        [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto* self = reinterpret_cast<CompareDirectoriesWindow*>(lParam);
            if (! self)
            {
                return TRUE;
            }
            SetWindowSubclass(child, CompareOptionsWheelRouteSubclassProc, 2u, reinterpret_cast<DWORD_PTR>(self));
            return TRUE;
        },
        reinterpret_cast<LPARAM>(this));
#pragma warning(pop)
}

void CompareDirectoriesWindow::PaintOptionsHostBackgroundAndCards(HDC hdc, HWND host) noexcept
{
    if (! hdc || ! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);

    if (_optionsBackgroundBrush)
    {
        FillRect(hdc, &rc, _optionsBackgroundBrush.get());
    }

    if (_theme.systemHighContrast || _theme.highContrast || _optionsCards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(host);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = ThemedControls::GetControlSurfaceColor(_theme);
    const COLORREF border  = ThemedControls::BlendColor(surface, _theme.menu.text, _theme.dark ? 40 : 30, 255);

    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! _optionsCardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, _optionsCardBrush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    for (const RECT& card : _optionsCards)
    {
        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }
}

void CompareDirectoriesWindow::LayoutOptionsControls() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    RECT rcDlg{};
    if (! GetClientRect(_optionsDlg.get(), &rcDlg))
    {
        return;
    }

    const int dlgW = std::max(0l, rcDlg.right - rcDlg.left);
    const int dlgH = std::max(0l, rcDlg.bottom - rcDlg.top);

    const UINT dpi = GetDpiForWindow(_optionsDlg.get());

    const int margin       = ThemedControls::ScaleDip(dpi, 16);
    const int gapX         = ThemedControls::ScaleDip(dpi, 12);
    const int gapY         = ThemedControls::ScaleDip(dpi, 12);
    const int rowHeight    = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int titleHeight  = std::max(1, ThemedControls::ScaleDip(dpi, 18));
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, 20));

    const int cardPaddingX   = ThemedControls::ScaleDip(dpi, 12);
    const int cardPaddingY   = ThemedControls::ScaleDip(dpi, 8);
    const int cardGapY       = ThemedControls::ScaleDip(dpi, 2);
    const int cardGapX       = ThemedControls::ScaleDip(dpi, 12);
    const int cardSpacingY   = ThemedControls::ScaleDip(dpi, 8);
    const int sectionSpacing = ThemedControls::ScaleDip(dpi, 16);
    const int framePadding   = ThemedControls::ScaleDip(dpi, 2);
    const int minToggleWidth = ThemedControls::ScaleDip(dpi, 90);

    const HFONT dialogFont = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    const HFONT headerFont = _uiBoldFont ? _uiBoldFont.get() : dialogFont;
    const HFONT infoFont   = _uiItalicFont ? _uiItalicFont.get() : dialogFont;

    const HWND okBtn     = GetDlgItem(_optionsDlg.get(), IDOK);
    const HWND cancelBtn = GetDlgItem(_optionsDlg.get(), IDCANCEL);

    const auto getWindowText = [](HWND hwnd) noexcept -> std::wstring
    {
        if (! hwnd)
        {
            return {};
        }
        const int len = GetWindowTextLengthW(hwnd);
        if (len <= 0)
        {
            return {};
        }
        std::wstring text(static_cast<size_t>(len), L'\0');
        const int copied = GetWindowTextW(hwnd, text.data(), len + 1);
        if (copied <= 0)
        {
            return {};
        }
        if (static_cast<size_t>(copied) < text.size())
        {
            text.resize(static_cast<size_t>(copied));
        }
        return text;
    };

    const int buttonPadX = ThemedControls::ScaleDip(dpi, 16);
    const int minBtnW    = ThemedControls::ScaleDip(dpi, 80);

    const auto measureButtonWidth = [&](HWND btn) noexcept -> int
    {
        const std::wstring text = getWindowText(btn);
        const int textW         = ThemedControls::MeasureTextWidth(_optionsDlg.get(), dialogFont, text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

    const int okW     = measureButtonWidth(okBtn);
    const int cancelW = measureButtonWidth(cancelBtn);

    const int buttonsY = std::max(0, dlgH - margin - rowHeight);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    int nextRight = std::max(0, dlgW - margin);
    if (cancelBtn)
    {
        nextRight -= cancelW;
        SetWindowPos(cancelBtn, nullptr, nextRight, buttonsY, cancelW, rowHeight, flags);
        SendMessageW(cancelBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        nextRight -= gapX;
    }
    if (okBtn)
    {
        nextRight -= okW;
        SetWindowPos(okBtn, nullptr, nextRight, buttonsY, okW, rowHeight, flags);
        SendMessageW(okBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int hostX = margin;
    const int hostY = margin;
    const int hostW = std::max(0, dlgW - 2 * margin);
    const int hostH = std::max(0, buttonsY - gapY - hostY);
    SetWindowPos(_optionsUi.host, nullptr, hostX, hostY, hostW, hostH, flags);

    RECT hostClient{};
    if (! GetClientRect(_optionsUi.host, &hostClient))
    {
        return;
    }

    const auto computeToggleWidth = [&](int contentW) noexcept -> int
    {
        const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
        const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

        const int onWidth  = ThemedControls::MeasureTextWidth(_optionsUi.host, headerFont, onLabel);
        const int offWidth = ThemedControls::MeasureTextWidth(_optionsUi.host, headerFont, offLabel);

        const int togglePaddingX = ThemedControls::ScaleDip(dpi, 6);
        const int toggleGapX     = ThemedControls::ScaleDip(dpi, 8);
        const int toggleTrackW   = ThemedControls::ScaleDip(dpi, 34);
        const int stateTextW     = std::max(onWidth, offWidth);

        const int measured = std::max(minToggleWidth, (2 * togglePaddingX) + stateTextW + toggleGapX + toggleTrackW);
        return std::min(std::max(0, contentW - 2 * cardPaddingX), measured);
    };

    auto computeToggleCardHeight = [&](int contentW, std::wstring_view descText, int toggleW) noexcept -> int
    {
        const int textW     = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH     = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int contentH  = std::max(0, titleHeight + cardGapY + descH);
        const int cardH     = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    auto computeIgnoreCardHeight = [&](int contentW, std::wstring_view descText, int toggleW, bool showEdit) noexcept -> int
    {
        const int textW = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);

        int contentH = std::max(0, titleHeight + cardGapY + descH);
        if (showEdit)
        {
            contentH += cardGapY + rowHeight;
        }
        const int cardH = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    auto computeContentHeight = [&](int contentW) noexcept -> int
    {
        const int toggleW = computeToggleWidth(contentW);

        const bool ignoreFilesOn = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
        const bool ignoreDirsOn  = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);

        int y = 0;
        y += headerHeight + gapY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), toggleW) + cardSpacingY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), toggleW) + cardSpacingY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), toggleW) + cardSpacingY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_CONTENT_DESC), toggleW) + cardSpacingY;

        y += sectionSpacing;
        y += headerHeight + gapY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), toggleW) + cardSpacingY;

        y += sectionSpacing;
        y += headerHeight + gapY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), toggleW) + cardSpacingY;
        y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), toggleW) + cardSpacingY;

        y += sectionSpacing;
        y += headerHeight + gapY;
        y += computeIgnoreCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC), toggleW, ignoreFilesOn) + cardSpacingY;
        y += computeIgnoreCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC), toggleW, ignoreDirsOn) + cardSpacingY;

        return y;
    };

    const int viewportW = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH = std::max(0l, hostClient.bottom - hostClient.top);

    int contentHeight    = computeContentHeight(viewportW);
    const bool wantsVScroll = viewportH > 0 && contentHeight > viewportH;

    LONG_PTR styleNow    = GetWindowLongPtrW(_optionsUi.host, GWL_STYLE);
    LONG_PTR styleWanted = styleNow;
    styleWanted |= WS_TABSTOP;
    styleWanted &= ~WS_HSCROLL;
    if (wantsVScroll)
    {
        styleWanted |= WS_VSCROLL;
    }
    else
    {
        styleWanted &= ~WS_VSCROLL;
    }

    const bool styleChanged = styleWanted != styleNow;
    if (styleChanged)
    {
        SetWindowLongPtrW(_optionsUi.host, GWL_STYLE, styleWanted);
        SetWindowPos(_optionsUi.host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsUi.host, hostTheme, nullptr);
        SendMessageW(_optionsUi.host, WM_THEMECHANGED, 0, 0);
    }

    GetClientRect(_optionsUi.host, &hostClient);
    const int viewportW2 = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH2 = std::max(0l, hostClient.bottom - hostClient.top);

    contentHeight         = computeContentHeight(viewportW2);
    _optionsScrollMax     = (viewportH2 > 0) ? std::max(0, contentHeight - viewportH2) : 0;
    _optionsScrollOffset  = std::clamp(_optionsScrollOffset, 0, _optionsScrollMax);
    if (_optionsScrollMax <= 0)
    {
        _optionsScrollOffset = 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentHeight - 1);
    si.nPage  = (viewportH2 > 0) ? static_cast<UINT>(viewportH2) : 0u;
    si.nPos   = _optionsScrollOffset;
    SetScrollInfo(_optionsUi.host, SB_VERT, &si, TRUE);

    _optionsCards.clear();

    const int scrollOffset = _optionsScrollOffset;
    const int toggleW      = computeToggleWidth(viewportW2);

    const auto positionScrollable = [&](HWND hwnd, int x, int y, int w, int h) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        SetWindowPos(hwnd, nullptr, x, y - scrollOffset, w, h, flags);
    };

    auto pushCard = [&](int top, int height) noexcept
    {
        RECT card{};
        card.left   = 0;
        card.top    = top - scrollOffset;
        card.right  = viewportW2;
        card.bottom = top + height - scrollOffset;
        _optionsCards.push_back(card);
    };

    const auto showToggleCardControls = [&](const OptionsToggleCard& card, bool visible) noexcept
    {
        const int cmd = visible ? SW_SHOW : SW_HIDE;
        ShowWindow(card.title, cmd);
        ShowWindow(card.description, cmd);
        ShowWindow(card.toggle, cmd);
    };

    const auto showIgnoreCardControls = [&](const OptionsIgnoreCard& card, bool visible, bool showEdit) noexcept
    {
        const int cmd = visible ? SW_SHOW : SW_HIDE;
        ShowWindow(card.title, cmd);
        ShowWindow(card.description, cmd);
        ShowWindow(card.toggle, cmd);
        if (card.frame)
        {
            ShowWindow(card.frame, (visible && showEdit) ? SW_SHOW : SW_HIDE);
        }
        if (card.edit)
        {
            ShowWindow(card.edit, (visible && showEdit) ? SW_SHOW : SW_HIDE);
        }
    };

    const auto layoutSectionHeader = [&](HWND header, UINT textId, int& y) noexcept
    {
        if (! header)
        {
            return;
        }

        const std::wstring text = LoadStringResource(nullptr, textId);
        SetWindowTextW(header, text.c_str());
        ShowWindow(header, SW_SHOW);
        positionScrollable(header, cardPaddingX, y, std::max(0, viewportW2 - 2 * cardPaddingX), headerHeight);
        SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    };

    const auto layoutToggleCard = [&](const OptionsToggleCard& card, UINT titleId, UINT descId, bool visible, int& y) noexcept
    {
        showToggleCardControls(card, visible);
        if (! visible)
        {
            return;
        }

        const std::wstring titleText = LoadStringResource(nullptr, titleId);
        const std::wstring descText  = LoadStringResource(nullptr, descId);

        const int textW = std::max(0, viewportW2 - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int cardH = computeToggleCardHeight(viewportW2, descText, toggleW);

        pushCard(y, cardH);

        SetWindowTextW(card.title, titleText.c_str());
        positionScrollable(card.title, cardPaddingX, y + cardPaddingY, textW, titleHeight);
        SendMessageW(card.title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        SetWindowTextW(card.description, descText.c_str());
        positionScrollable(card.description, cardPaddingX, y + cardPaddingY + titleHeight + cardGapY, textW, std::max(0, descH));
        SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);

        positionScrollable(card.toggle, viewportW2 - cardPaddingX - toggleW, y + (cardH - rowHeight) / 2, toggleW, rowHeight);
        SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        y += cardH + cardSpacingY;
    };

    const auto layoutIgnoreCard = [&](const OptionsIgnoreCard& card,
                                      UINT titleId,
                                      UINT descId,
                                      bool visible,
                                      bool showEdit,
                                      int& y) noexcept
    {
        showIgnoreCardControls(card, visible, showEdit);
        if (! visible)
        {
            return;
        }

        const std::wstring titleText = LoadStringResource(nullptr, titleId);
        const std::wstring descText  = LoadStringResource(nullptr, descId);

        const int textW = std::max(0, viewportW2 - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int cardH = computeIgnoreCardHeight(viewportW2, descText, toggleW, showEdit);

        pushCard(y, cardH);

        SetWindowTextW(card.title, titleText.c_str());
        positionScrollable(card.title, cardPaddingX, y + cardPaddingY, textW, titleHeight);
        SendMessageW(card.title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        SetWindowTextW(card.description, descText.c_str());
        positionScrollable(card.description, cardPaddingX, y + cardPaddingY + titleHeight + cardGapY, textW, std::max(0, descH));
        SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);

        positionScrollable(card.toggle, viewportW2 - cardPaddingX - toggleW, y + cardPaddingY, toggleW, rowHeight);
        SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        if (showEdit && card.frame && card.edit)
        {
            const int editX = cardPaddingX;
            const int editW = std::max(0, viewportW2 - 2 * cardPaddingX);

            const int contentTop    = y + cardPaddingY;
            const int contentBottom = contentTop + titleHeight + cardGapY + descH;
            const int editTop       = contentBottom + cardGapY;

            const int innerPadding = (! _theme.highContrast && card.frame) ? framePadding : 0;

            positionScrollable(card.frame, editX, editTop, editW, rowHeight);
            positionScrollable(card.edit,
                               editX + innerPadding,
                               editTop + innerPadding,
                               std::max(1, editW - 2 * innerPadding),
                               std::max(1, rowHeight - 2 * innerPadding));
            SendMessageW(card.edit, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            ThemedControls::CenterEditTextVertically(card.edit);
        }

        y += cardH + cardSpacingY;
    };

    const bool ignoreFilesOn = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
    const bool ignoreDirsOn  = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);

    int y = 0;

    layoutSectionHeader(_optionsUi.headerCompare, IDS_COMPARE_OPTIONS_SECTION_COMPARE, y);
    layoutToggleCard(_optionsUi.compareSize, IDS_COMPARE_OPTIONS_SIZE_TITLE, IDS_COMPARE_OPTIONS_SIZE_DESC, true, y);
    layoutToggleCard(_optionsUi.compareDateTime, IDS_COMPARE_OPTIONS_DATETIME_TITLE, IDS_COMPARE_OPTIONS_DATETIME_DESC, true, y);
    layoutToggleCard(_optionsUi.compareAttributes, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC, true, y);
    layoutToggleCard(_optionsUi.compareContent, IDS_COMPARE_OPTIONS_CONTENT_TITLE, IDS_COMPARE_OPTIONS_CONTENT_DESC, true, y);

    y += sectionSpacing;
    layoutSectionHeader(_optionsUi.headerSubdirs, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS, y);
    layoutToggleCard(_optionsUi.compareSubdirectories, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SUBDIRS_DESC, true, y);

    y += sectionSpacing;
    layoutSectionHeader(_optionsUi.headerAdvanced, IDS_COMPARE_OPTIONS_SECTION_ADVANCED, y);
    layoutToggleCard(_optionsUi.compareSubdirAttributes,
                     IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE,
                     IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC,
                     true,
                     y);
    layoutToggleCard(
        _optionsUi.selectSubdirsOnlyInOnePane, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC, true, y);

    y += sectionSpacing;
    layoutSectionHeader(_optionsUi.headerIgnore, IDS_COMPARE_OPTIONS_SECTION_IGNORE, y);
    layoutIgnoreCard(_optionsUi.ignoreFiles,
                     IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE,
                     IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC,
                     true,
                     ignoreFilesOn,
                     y);
    layoutIgnoreCard(_optionsUi.ignoreDirectories,
                     IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE,
                     IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC,
                     true,
                     ignoreDirsOn,
                     y);

    InvalidateRect(_optionsUi.host, nullptr, TRUE);
}

void CompareDirectoriesWindow::Layout() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    RECT rc{};
    if (GetClientRect(_hWnd.get(), &rc) == 0)
    {
        return;
    }

    const int w = std::max(0, static_cast<int>(rc.right - rc.left));
    const int h = std::max(0, static_cast<int>(rc.bottom - rc.top));

    _clientSize = {w, h};

    const int dpi           = static_cast<int>(_dpi);
    const int bannerHeight  = std::clamp(MulDiv(42, dpi, USER_DEFAULT_SCREEN_DPI), 0, h);
    const bool showStatus   = (_scanProgressText && IsWindowVisible(_scanProgressText.get()) != 0) || (_scanProgressBar && IsWindowVisible(_scanProgressBar.get()) != 0);
    const int statusHeight  = showStatus ? std::clamp(MulDiv(kScanStatusHeightDip, dpi, USER_DEFAULT_SCREEN_DPI), 0, std::max(0, h - bannerHeight)) : 0;
    const int contentHeight = std::max(0, h - bannerHeight - statusHeight);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    // Banner layout
    const int bannerPaddingX = std::max(0, MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI));
    const int bannerPaddingY = std::max(0, MulDiv(6, dpi, USER_DEFAULT_SCREEN_DPI));
    const int buttonW        = std::max(1, MulDiv(110, dpi, USER_DEFAULT_SCREEN_DPI));
    const int buttonH        = std::max(1, MulDiv(28, dpi, USER_DEFAULT_SCREEN_DPI));
    const int buttonGap      = std::max(0, MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI));
    const int buttonY        = std::max(0, bannerPaddingY + (std::max(0, bannerHeight - (2 * bannerPaddingY) - buttonH) / 2));

    int rightX = std::max(0, w - bannerPaddingX);
    if (_bannerRescanButton)
    {
        rightX = std::max(0, rightX - buttonW);
        SetWindowPos(_bannerRescanButton.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
        rightX = std::max(0, rightX - buttonGap);
    }
    if (_bannerOptionsButton)
    {
        rightX = std::max(0, rightX - buttonW);
        SetWindowPos(_bannerOptionsButton.get(), nullptr, rightX, buttonY, buttonW, buttonH, flags);
        rightX = std::max(0, rightX - buttonGap);
    }
    if (_bannerTitle)
    {
        const int titleW = std::max(0, rightX - bannerPaddingX);
        SetWindowPos(_bannerTitle.get(), nullptr, bannerPaddingX, 0, titleW, bannerHeight, flags);
    }

    if (const HWND fw = _folderWindow.GetHwnd())
    {
        SetWindowPos(fw, nullptr, 0, bannerHeight, w, contentHeight, flags);
    }

    if (showStatus && (_scanProgressText || _scanProgressBar))
    {
        const int statusTop = bannerHeight + contentHeight;
        const int paddingX  = std::max(0, MulDiv(kScanStatusPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI));

        int progressBarW = std::max(1, MulDiv(kScanProgressBarWidthDip, dpi, USER_DEFAULT_SCREEN_DPI));
        progressBarW     = std::clamp(progressBarW, 1, std::max(1, w / 2));

        int progressBarH = std::max(1, MulDiv(kScanProgressBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI));
        progressBarH     = std::clamp(progressBarH, 1, std::max(1, statusHeight));

        const int progressBarX = std::max(0, w - paddingX - progressBarW);
        const int progressBarY = statusTop + std::max(0, (statusHeight - progressBarH) / 2);

        if (_scanProgressBar)
        {
            SetWindowPos(_scanProgressBar.get(), nullptr, progressBarX, progressBarY, progressBarW, progressBarH, flags);
        }

        if (_scanProgressText)
        {
            const int textX = paddingX;
            const int textW = std::max(0, progressBarX - paddingX - paddingX);
            SetWindowPos(_scanProgressText.get(), nullptr, textX, statusTop, textW, statusHeight, flags);
        }
    }

    if (_optionsDlg && IsWindowVisible(_optionsDlg.get()) != 0)
    {
        RECT dr{};
        GetWindowRect(_optionsDlg.get(), &dr);
        const int dw = std::max(1, static_cast<int>(dr.right - dr.left));
        const int dh = std::max(1, static_cast<int>(dr.bottom - dr.top));

        const int x = std::max(0, (w - dw) / 2);
        const int y = std::max(bannerHeight, bannerHeight + (contentHeight - dh) / 2);
        SetWindowPos(_optionsDlg.get(), nullptr, x, y, dw, dh, SWP_NOZORDER | SWP_NOACTIVATE);
        LayoutOptionsControls();
    }
}

void CompareDirectoriesWindow::EnsureCompareSession() noexcept
{
    if (_session)
    {
        return;
    }

    if (! _baseFs)
    {
        return;
    }

    Common::Settings::CompareDirectoriesSettings settings = GetEffectiveCompareSettings();
    _session = std::make_shared<CompareDirectoriesSession>(_baseFs, _leftRoot, _rightRoot, settings);
    if (_hWnd)
    {
        const HWND hwnd = _hWnd.get();
        _session->SetScanProgressCallback(
            [hwnd](const std::filesystem::path& relativeFolder,
                   std::wstring_view currentEntryName,
                   uint64_t scannedFolders,
                   uint64_t scannedEntries,
                   uint32_t activeScans) noexcept
            {
                if (! hwnd)
                {
                    return;
                }

                auto payload          = std::make_unique<ScanProgressPayload>();
                payload->activeScans  = activeScans;
                payload->folderCount  = scannedFolders;
                payload->entryCount   = scannedEntries;
                payload->relativeFolder = relativeFolder;
                payload->entryName    = std::wstring(currentEntryName);
                static_cast<void>(PostMessagePayload(hwnd, WndMsg::kCompareDirectoriesScanProgress, 0, std::move(payload)));
            });
    }

    _fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, _session);
    _fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, _session);
}

void CompareDirectoriesWindow::StartCompare() noexcept
{
    EnsureCompareSession();
    if (! _session || ! _fsLeft || ! _fsRight)
    {
        return;
    }

    if (_settings && _settings->compareDirectories.has_value())
    {
        _session->SetSettings(_settings->compareDirectories.value());
    }

    if (_compareStarted)
    {
        ShowOptionsPanel(false);
        RefreshBothPanes();
        return;
    }

    static_cast<void>(_folderWindow.SetFileSystemInstanceForPane(FolderWindow::Pane::Left,
                                                                 _fsLeft,
                                                                 std::wstring(L"builtin/file-system"),
                                                                 std::wstring(L"file"),
                                                                 std::wstring{}));
    static_cast<void>(_folderWindow.SetFileSystemInstanceForPane(FolderWindow::Pane::Right,
                                                                 _fsRight,
                                                                 std::wstring(L"builtin/file-system"),
                                                                 std::wstring(L"file"),
                                                                 std::wstring{}));

    _folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, true);
    _folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, true);

    _folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    _folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Detailed);

    _compareStarted = true;
    ShowOptionsPanel(false);

    if (const HWND fw = _folderWindow.GetHwnd())
    {
        ShowWindow(fw, SW_SHOW);
    }
    Layout();

    _syncingPaths = true;
    _folderWindow.SetFolderPath(FolderWindow::Pane::Left, _leftRoot);
    _folderWindow.SetFolderPath(FolderWindow::Pane::Right, _rightRoot);
    _syncingPaths = false;

    SetFocus(_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left));
}

void CompareDirectoriesWindow::ShowOptionsPanel(bool show) noexcept
{
    if (! _optionsDlg)
    {
        return;
    }

    ShowWindow(_optionsDlg.get(), show ? SW_SHOW : SW_HIDE);
    if (show)
    {
        LoadOptionsControlsFromSettings();
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_HIDE);
        }

        Layout();
        ApplyOptionsDialogTheme();
        SetWindowPos(_optionsDlg.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        RedrawWindow(_optionsDlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
        SetFocus(_optionsDlg.get());
        return;
    }

    if (_compareStarted)
    {
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_SHOW);
        }

        Layout();
        const HWND focus = GetFocus();
        if (! focus || (focus == _optionsDlg.get()) || IsChild(_optionsDlg.get(), focus))
        {
            SetFocus(_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left));
        }
    }
}

void CompareDirectoriesWindow::SyncOtherPanePath(ComparePane changedPane, const std::optional<std::filesystem::path>& newPath) noexcept
{
    if (! _compareStarted || _syncingPaths || ! _session || ! newPath.has_value())
    {
        return;
    }

    const auto relOpt = _session->TryMakeRelative(changedPane, newPath.value());
    if (! relOpt.has_value())
    {
        // Root changed (path outside previous compare roots): invalidate scan + reset both panes to the new roots.
        const std::filesystem::path otherRoot = _session->GetRoot(changedPane == ComparePane::Left ? ComparePane::Right : ComparePane::Left);

        if (changedPane == ComparePane::Left)
        {
            _leftRoot  = newPath.value();
            _rightRoot = otherRoot;
        }
        else
        {
            _leftRoot  = otherRoot;
            _rightRoot = newPath.value();
        }

        _session->SetRoots(_leftRoot, _rightRoot);

        _syncingPaths = true;
        _folderWindow.SetFolderPath(FolderWindow::Pane::Left, _leftRoot);
        _folderWindow.SetFolderPath(FolderWindow::Pane::Right, _rightRoot);
        _syncingPaths = false;
        return;
    }

    const ComparePane other = changedPane == ComparePane::Left ? ComparePane::Right : ComparePane::Left;
    const std::filesystem::path otherAbs = _session->ResolveAbsolute(other, relOpt.value());

    _syncingPaths = true;
    if (other == ComparePane::Left)
    {
        _folderWindow.SetFolderPath(FolderWindow::Pane::Left, otherAbs);
    }
    else
    {
        _folderWindow.SetFolderPath(FolderWindow::Pane::Right, otherAbs);
    }
    _syncingPaths = false;
}

void CompareDirectoriesWindow::ApplySelectionForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept
{
    if (! _compareStarted || ! _session)
    {
        return;
    }

    const auto relOpt = _session->TryMakeRelative(pane, folder);
    if (! relOpt.has_value())
    {
        return;
    }

    const auto decision = _session->GetOrComputeDecision(relOpt.value());
    if (! decision || FAILED(decision->hr))
    {
        return;
    }

    const bool isLeft = pane == ComparePane::Left;
    auto shouldSelect = [&](std::wstring_view name) noexcept -> bool
    {
        const auto it = decision->items.find(name);
        if (it == decision->items.end())
        {
            return false;
        }

        return isLeft ? it->second.selectLeft : it->second.selectRight;
    };

    if (isLeft)
    {
        _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, shouldSelect, true);
    }
    else
    {
        _folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, shouldSelect, true);
    }
}

std::wstring CompareDirectoriesWindow::BuildDetailsTextForCompareItem(ComparePane pane,
                                                                     const std::filesystem::path& folder,
                                                                     std::wstring_view displayName,
                                                                     bool isDirectory,
                                                                     uint64_t sizeBytes,
                                                                     int64_t lastWriteTime,
                                                                     DWORD fileAttributes) noexcept
{
    if (! _compareStarted || ! _session)
    {
        return {};
    }

    DetailsDecisionCache& cache = pane == ComparePane::Left ? _detailsCacheLeft : _detailsCacheRight;
    const uint64_t currentUiVersion = _session->GetUiVersion();

    if (cache.sessionUiVersion != currentUiVersion || cache.folder != folder)
    {
        cache.sessionUiVersion = currentUiVersion;
        cache.folder           = folder;
        cache.decision.reset();

        if (const auto relOpt = _session->TryMakeRelative(pane, folder))
        {
            cache.decision = _session->GetOrComputeDecision(relOpt.value());
        }
    }

    const auto decision = cache.decision;
    if (! decision || FAILED(decision->hr))
    {
        return {};
    }

    const auto it = decision->items.find(displayName);
    if (it == decision->items.end())
    {
        return {};
    }

    const CompareDirectoriesItemDecision& item = it->second;
    const uint32_t diffMask                    = item.differenceMask;
    const auto& strings                        = GetCompareDetailsTextStrings();

    std::wstring statusText;

    if (diffMask == 0u)
    {
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::OnlyInLeft))
    {
        statusText.assign(strings.onlyInLeft);
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::OnlyInRight))
    {
        statusText.assign(strings.onlyInRight);
    }
    else if (HasFlag(diffMask, CompareDirectoriesDiffBit::TypeMismatch))
    {
        statusText.assign(strings.typeMismatch);
    }
    if (statusText.empty() && diffMask != 0u)
    {
        statusText.reserve(64);

        auto appendToken = [&](std::wstring_view token) noexcept
        {
            if (token.empty())
            {
                return;
            }

            if (! statusText.empty())
            {
                statusText.append(L" • ");
            }
            statusText.append(token);
        };

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Size))
        {
            const bool thisBigger =
                pane == ComparePane::Left ? (item.leftSizeBytes > item.rightSizeBytes) : (item.rightSizeBytes > item.leftSizeBytes);
            appendToken(thisBigger ? strings.bigger : strings.smaller);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::DateTime))
        {
            const bool thisNewer =
                pane == ComparePane::Left ? (item.leftLastWriteTime > item.rightLastWriteTime) : (item.rightLastWriteTime > item.leftLastWriteTime);
            appendToken(thisNewer ? strings.newer : strings.older);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Attributes))
        {
            appendToken(strings.attributesDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::Content))
        {
            appendToken(strings.contentDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirAttributes))
        {
            appendToken(strings.subdirAttributesDiffer);
        }

        if (HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirContent))
        {
            appendToken(strings.subdirContentDiffer);
        }
    }

    if (_detailsVerbosity == CompareDetailsVerbosity::StatusOnly)
    {
        return statusText;
    }

    const std::wstring metaText = BuildMetadataDetailsText(isDirectory, sizeBytes, lastWriteTime, fileAttributes);
    if (statusText.empty())
    {
        return metaText;
    }

    if (metaText.empty())
    {
        return statusText;
    }

    statusText.append(L" • ");
    statusText.append(metaText);
    return statusText;
}

void CompareDirectoriesWindow::RefreshBothPanes() noexcept
{
    if (! _compareStarted)
    {
        return;
    }

    const FolderWindow::Pane pane = _folderWindow.GetFocusedPane();
    _folderWindow.SetActivePane(pane);
    _folderWindow.CommandRefresh(FolderWindow::Pane::Left);
    _folderWindow.CommandRefresh(FolderWindow::Pane::Right);
}

void CompareDirectoriesWindow::OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept
{
    if (! _compareStarted || ! _session)
    {
        return;
    }

    // Invalidate affected paths so the forced refresh performed by FolderWindow updates the compare decisions.
    for (const auto& src : e.sourcePaths)
    {
        _session->InvalidateForAbsolutePath(src, true);

        if (e.destinationFolder.has_value())
        {
            const std::filesystem::path dst = e.destinationFolder.value() / src.filename();
            _session->InvalidateForAbsolutePath(dst, true);
        }
    }
}

LRESULT CompareDirectoriesWindow::OnScanProgress(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ScanProgressPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    if (! _scanProgressText || ! _scanProgressBar)
    {
        return 0;
    }

    if (payload->activeScans > 0u)
    {
        std::filesystem::path displayPath = payload->relativeFolder;
        if (! payload->entryName.empty())
        {
            displayPath /= std::filesystem::path(payload->entryName);
        }

        std::wstring pathText;
        if (displayPath.empty())
        {
            pathText = L".";
        }
        else
        {
            pathText = displayPath.wstring();
        }

        const std::wstring message = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, pathText, payload->folderCount, payload->entryCount);
        SetWindowTextW(_scanProgressText.get(), message.c_str());

        const bool needLayout = IsWindowVisible(_scanProgressText.get()) == 0 || IsWindowVisible(_scanProgressBar.get()) == 0;
        ShowWindow(_scanProgressText.get(), SW_SHOW);
        ShowWindow(_scanProgressBar.get(), SW_SHOW);
        SendMessageW(_scanProgressBar.get(), PBM_SETMARQUEE, TRUE, 30);
        if (needLayout)
        {
            Layout();
        }
        return 0;
    }

    SendMessageW(_scanProgressBar.get(), PBM_SETMARQUEE, FALSE, 0);
    SetWindowTextW(_scanProgressText.get(), L"");
    const bool needLayout = IsWindowVisible(_scanProgressText.get()) != 0 || IsWindowVisible(_scanProgressBar.get()) != 0;
    ShowWindow(_scanProgressBar.get(), SW_HIDE);
    ShowWindow(_scanProgressText.get(), SW_HIDE);
    if (needLayout)
    {
        Layout();
    }
    return 0;
}

LRESULT CompareDirectoriesWindow::OnExecuteShortcutCommand(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<std::wstring>(lp);
    if (! payload || payload->empty())
    {
        return 0;
    }

    ExecuteShortcutCommand(*payload);
    return 0;
}

Common::Settings::CompareDirectoriesSettings CompareDirectoriesWindow::GetEffectiveCompareSettings() const noexcept
{
    if (_settings && _settings->compareDirectories.has_value())
    {
        return _settings->compareDirectories.value();
    }

    return Common::Settings::CompareDirectoriesSettings{};
}

void CompareDirectoriesWindow::LoadOptionsControlsFromSettings() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();

    SetTwoStateToggleState(_optionsUi.compareSize.toggle, _theme.highContrast, s.compareSize);
    SetTwoStateToggleState(_optionsUi.compareDateTime.toggle, _theme.highContrast, s.compareDateTime);
    SetTwoStateToggleState(_optionsUi.compareAttributes.toggle, _theme.highContrast, s.compareAttributes);
    SetTwoStateToggleState(_optionsUi.compareContent.toggle, _theme.highContrast, s.compareContent);

    SetTwoStateToggleState(_optionsUi.compareSubdirectories.toggle, _theme.highContrast, s.compareSubdirectories);

    SetTwoStateToggleState(_optionsUi.compareSubdirAttributes.toggle, _theme.highContrast, s.compareSubdirectoryAttributes);
    SetTwoStateToggleState(_optionsUi.selectSubdirsOnlyInOnePane.toggle, _theme.highContrast, s.selectSubdirsOnlyInOnePane);

    SetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast, s.ignoreFiles);
    SetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast, s.ignoreDirectories);
    if (_optionsUi.ignoreFiles.edit)
    {
        SetWindowTextW(_optionsUi.ignoreFiles.edit, s.ignoreFilesPatterns.c_str());
    }
    if (_optionsUi.ignoreDirectories.edit)
    {
        SetWindowTextW(_optionsUi.ignoreDirectories.edit, s.ignoreDirectoriesPatterns.c_str());
    }

    UpdateOptionsVisibility();
}

void CompareDirectoriesWindow::SaveOptionsControlsToSettings() noexcept
{
    if (! _optionsDlg || ! _settings || ! _optionsUi.host)
    {
        return;
    }

    Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();

    s.compareSize       = GetTwoStateToggleState(_optionsUi.compareSize.toggle, _theme.highContrast);
    s.compareDateTime   = GetTwoStateToggleState(_optionsUi.compareDateTime.toggle, _theme.highContrast);
    s.compareAttributes = GetTwoStateToggleState(_optionsUi.compareAttributes.toggle, _theme.highContrast);
    s.compareContent    = GetTwoStateToggleState(_optionsUi.compareContent.toggle, _theme.highContrast);

    s.compareSubdirectories         = GetTwoStateToggleState(_optionsUi.compareSubdirectories.toggle, _theme.highContrast);
    s.compareSubdirectoryAttributes = GetTwoStateToggleState(_optionsUi.compareSubdirAttributes.toggle, _theme.highContrast);
    s.selectSubdirsOnlyInOnePane    = GetTwoStateToggleState(_optionsUi.selectSubdirsOnlyInOnePane.toggle, _theme.highContrast);

    s.ignoreFiles               = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
    s.ignoreDirectories         = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);
    s.ignoreFilesPatterns       = _optionsUi.ignoreFiles.edit ? GetDlgItemTextString(_optionsUi.host, IDC_CMP_IGNORE_FILES_PATTERNS) : std::wstring{};
    s.ignoreDirectoriesPatterns = _optionsUi.ignoreDirectories.edit ? GetDlgItemTextString(_optionsUi.host, IDC_CMP_IGNORE_DIRECTORIES_PATTERNS) : std::wstring{};

    _settings->compareDirectories = std::move(s);
}

void CompareDirectoriesWindow::UpdateOptionsVisibility() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    LayoutOptionsControls();
    RedrawWindow(_optionsUi.host, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}
} // namespace

bool ShowCompareDirectoriesWindow(HWND owner,
                                  Common::Settings::Settings& settings,
                                  const AppTheme& theme,
                                  const ShortcutManager* shortcuts,
                                  wil::com_ptr<IFileSystem> baseFileSystem,
                                  std::filesystem::path leftRoot,
                                  std::filesystem::path rightRoot) noexcept
{
    auto* window = new CompareDirectoriesWindow(settings, theme, shortcuts, std::move(baseFileSystem), std::move(leftRoot), std::move(rightRoot));
    if (! window->Create(owner))
    {
        delete window;
        return false;
    }

    return true;
}

void UpdateCompareDirectoriesWindowsTheme(const AppTheme& theme) noexcept
{
    for (CompareDirectoriesWindow* window : g_compareDirectoriesWindows)
    {
        if (window)
        {
            window->UpdateTheme(theme);
        }
    }
}

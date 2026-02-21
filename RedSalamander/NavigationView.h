#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <deque>
#include <filesystem>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <vector>

#include "AppTheme.h"
#include "PlugInterfaces/NavigationMenu.h"

#include <Windows.h>
#include <d2d1_1.h>
#include <d3d11.h>
#include <dwrite.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <d2d1.h>
#include <dcommon.h>
#include <dxgi1_2.h>
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

struct IFileSystem;
struct IFileSystemIO;
struct IDriveInfo;

namespace Common::Settings
{
struct Settings;
}

class NavigationView : public INavigationMenuCallback
{
public:
    NavigationView();
    virtual ~NavigationView();

    // Disable copy and move
    NavigationView(const NavigationView&)            = delete;
    NavigationView& operator=(const NavigationView&) = delete;
    NavigationView(NavigationView&&)                 = delete;
    NavigationView& operator=(NavigationView&&)      = delete;

    // Window lifecycle
    static ATOM RegisterWndClass(HINSTANCE instance);
    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy();
    [[maybe_unused]] HWND GetHwnd() const
    {
        return _hWnd.get();
    }

    void OnDpiChanged(float newDpi);

    // Startup warmup
    static void WarmSharedDeviceResources() noexcept;

    // Path management
    void SetPath(const std::optional<std::filesystem::path>& path);
    [[maybe_unused]] std::optional<std::filesystem::path> GetPath() const
    {
        return _currentPath;
    }

    enum class FocusRegion : uint8_t
    {
        Menu,
        Path,
        History,
        DiskInfo,
    };

    void SetFocusRegion(FocusRegion region);
    void FocusAddressBar();
    void OpenChangeDirectoryFromCommand();
    void OpenHistoryDropdownFromKeyboard();
    void OpenDriveMenuFromCommand();

    void SetTheme(const AppTheme& theme);
    void SetPaneFocused(bool focused) noexcept;

    HRESULT STDMETHODCALLTYPE NavigationMenuRequestNavigate(const wchar_t* path, void* cookie) noexcept override;

    // Callbacks
    using PathChangedCallback = std::function<void(const std::optional<std::filesystem::path>&)>;
    void SetPathChangedCallback(PathChangedCallback callback);

    using RequestFolderViewFocusCallback = std::function<void()>;
    void SetRequestFolderViewFocusCallback(RequestFolderViewFocusCallback callback);

    // History (most recent first)
    std::vector<std::filesystem::path> GetHistory() const;
    void SetHistory(const std::vector<std::filesystem::path>& history);

    void SetFileSystem(const wil::com_ptr<IFileSystem>& fileSystem);
    void SetSettings(Common::Settings::Settings* settings) noexcept;

    // Constants (public for layout calculations)
    static constexpr int kHeight = 24; // DIP at 96 DPI

private:
    static constexpr wchar_t kClassName[]                 = L"RedSalamander.NavigationView";
    static constexpr wchar_t kFullPathPopupClassName[]    = L"RedSalamander.FullPathPopup";
    static constexpr wchar_t kEditSuggestPopupClassName[] = L"RedSalamander.EditSuggestPopup";
    static constexpr int kDriveSectionWidth               = 28; // Menu button
    static constexpr int kDiskInfoSectionWidth            = 70; // Disk info
    static constexpr int kHistoryButtonWidth              = 24; // History dropdown

    void RequestPathChange(const std::filesystem::path& path);
    std::filesystem::path ToPluginPath(const std::filesystem::path& displayPath) const;

    enum class RenderMode : uint8_t
    {
        Breadcrumb, // Default: clickable path segments
        FullPath,   // Hover: show complete path
        Edit        // Edit mode: Win32 Edit control
    };

    struct PathSegment
    {
        std::wstring text;
        D2D1_RECT_F bounds = {};
        std::filesystem::path fullPath;
        bool isEllipsis = false;
        wil::com_ptr<IDWriteTextLayout> layout;
    };

    struct BreadcrumbSeparator
    {
        D2D1_RECT_F bounds       = {};
        size_t leftSegmentIndex  = 0;
        size_t rightSegmentIndex = 0;
    };

    struct EditSuggestResultsPayload
    {
        uint64_t requestId         = 0;
        bool hasMore               = false;
        wchar_t directorySeparator = L'\\';
        std::wstring highlightText;
        std::vector<std::wstring> displayItems;
        std::vector<std::wstring> insertItems;
    };

    static LRESULT CALLBACK WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    static ATOM RegisterFullPathPopupWndClass(HINSTANCE instance);
    static LRESULT CALLBACK FullPathPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT FullPathPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    static ATOM RegisterEditSuggestPopupWndClass(HINSTANCE instance);
    static LRESULT CALLBACK EditSuggestPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    static LRESULT OnEditSubclassNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp, UINT_PTR subclassId) noexcept;
    LRESULT EditSuggestPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    // Message handlers
    void OnCreate(HWND hwnd);
    void OnDeferredInit();
    void OnDestroy();
    void OnPaint();
    void OnSize(UINT width, UINT height);
    void OnCommand(UINT id, HWND hwndCtl, UINT codeNotify);
    void OnDrawItem(DRAWITEMSTRUCT* dis);
    void OnMeasureItem(MEASUREITEMSTRUCT* mis);
    LRESULT OnCtlColorEdit(HDC hdc, HWND hwndControl);
    void OnLButtonDown(POINT pt);
    void OnLButtonDblClk(POINT pt);
    void OnMouseMove(POINT pt);
    void OnMouseLeave();
    void OnSetCursor(HWND hwnd, UINT hitTest, UINT mouseMsg);
    void OnTimer(UINT_PTR timerId);
    void OnEnterMenuLoop(bool isTrackPopupMenu);
    void OnExitMenuLoop(bool isTrackPopupMenu);
    void OnSetFocus();
    void OnKillFocus(HWND newFocus);
    LRESULT OnEditSuggestResults(std::unique_ptr<EditSuggestResultsPayload> payload);
    LRESULT OnNavigationMenuRequestPath(std::unique_ptr<std::wstring> text);
    bool OnKeyDown(WPARAM key);
    void MoveFocus(bool forward);
    void ActivateFocusedRegion();
    void NormalizeFocusRegion();

    // Popup message handlers
    LRESULT OnFullPathPopupCreate(HWND hwnd);
    LRESULT OnFullPathPopupNcDestroy(HWND hwnd);
    LRESULT OnFullPathPopupTimer(HWND hwnd, UINT_PTR timerId);
    LRESULT OnFullPathPopupSize(HWND hwnd, UINT width, UINT height);
    LRESULT OnFullPathPopupMouseMove(HWND hwnd, POINT pt);
    LRESULT OnFullPathPopupMouseLeave(HWND hwnd);
    LRESULT OnFullPathPopupLButtonDown(HWND hwnd, POINT pt);
    LRESULT OnFullPathPopupLButtonDblClk(HWND hwnd, POINT pt);
    LRESULT OnFullPathPopupActivate(WORD state);
    LRESULT OnFullPathPopupKeyDown(WPARAM key);
    LRESULT OnFullPathPopupSysKeyDown(HWND hwnd, WPARAM key, LPARAM lParam);
    LRESULT OnFullPathPopupSysChar(HWND hwnd, WPARAM key, LPARAM lParam);
    LRESULT OnFullPathPopupMouseWheel(HWND hwnd, int delta);
    LRESULT OnFullPathPopupCtlColorEdit(HWND hwnd, HDC hdc, HWND hwndControl);
    LRESULT OnFullPathPopupCommand(HWND hwnd, UINT id, UINT codeNotify, HWND hwndCtl);
    LRESULT OnShowFullPathPopupSiblingsDropdown(HWND popupHwnd, size_t separatorIndex);

    LRESULT OnEditSuggestPopupCreate();
    LRESULT OnEditSuggestPopupNcDestroy();
    LRESULT OnEditSuggestPopupSize(HWND hwnd, UINT width, UINT height);
    LRESULT OnEditSuggestPopupMouseMove(HWND hwnd, POINT pt);
    LRESULT OnEditSuggestPopupMouseLeave(HWND hwnd);
    LRESULT OnEditSuggestPopupLButtonDown(HWND hwnd, POINT pt);

    // Layout
    RECT _sectionDriveRect    = {}; // Menu button
    RECT _sectionPathRect     = {}; // Path display
    RECT _sectionHistoryRect  = {}; // History dropdown button
    RECT _sectionDiskInfoRect = {}; // Disk info

    // Child controls (Win32)
    wil::unique_hwnd _pathEdit; // Section Path: Edit control (edit mode only)

    // Direct2D rendering for all sections
    void EnsureD2DResources();
    void DiscardD2DResources();
    void RenderDriveSection();    // Render menu button
    void RenderPathSection();     // Render Breadcrumb
    void RenderHistorySection();  // Render History button
    void RenderDiskInfoSection(); // Render disk info
    void Present(std::optional<RECT*> dirtyRect);

    void UpdateBreadcrumbLayout(); // Build segments/separators when path changes
    void UpdateMenuIconBitmap();   // Convert menu icon HICON to D2D bitmap
    void RenderBreadcrumbs();      // Render segments/separators from cached layout
    std::vector<PathSegment> SplitPathComponents(const std::filesystem::path& path);
    void InvalidateBreadcrumbLayoutCache() noexcept;
    void EnsureBreadcrumbTextLayoutCache(float height) noexcept;
    void GetBreadcrumbTextLayoutAndWidth(std::wstring_view text, float height, wil::com_ptr<IDWriteTextLayout>& layout, float& width) noexcept;

    // Menus
    void ShowMenuDropdown();
    void ShowFileSystemDriveMenuDropdown();
    void ShowHistoryDropdown();
    void ShowDiskInfoDropdown();
    void ShowSiblingsDropdown(size_t segmentIndex);
    void RequestFullPathPopup(const D2D1_RECT_F& anchorBounds);
    void ShowFullPathPopup();
    void ShowFullPathPopupSiblingsDropdown(HWND popupHwnd, size_t separatorIndex);
    void CloseFullPathPopup();
    void EnsureFullPathPopupD2DResources();
    void DiscardFullPathPopupD2DResources();
    void BuildFullPathPopupLayout(float clientWidth);
    void RenderFullPathPopup();
    void PrepareThemedMenu(HMENU menu);
    void ClearThemedMenuState();
    int TrackThemedPopupMenuReturnCmd(HMENU menu, UINT flags, const POINT& screenPoint, HWND ownerWindow);
    bool TryGetSiblingFolders(const std::filesystem::path& parentPath, std::vector<std::filesystem::path>& siblings);
    void BuildSiblingFoldersMenu(HMENU menu, const std::vector<std::filesystem::path>& siblings, const std::filesystem::path& currentPath);

    void EnsureSiblingPrefetchWorker();
    void QueueSiblingPrefetchForPath(const std::filesystem::path& displayPath);
    void QueueSiblingPrefetchForParent(const std::filesystem::path& parentPath);
    void SiblingPrefetchWorker(std::stop_token stopToken);

    // Path editing
    void EnterEditMode();
    void ExitEditMode(bool accept);
    void EnterFullPathPopupEditMode();
    void ExitFullPathPopupEditMode(bool accept);
    bool ValidatePath(const std::wstring& pathStr);
    bool HandleEditSubclassKeyDown(HWND editHwnd, WPARAM key);
    bool HandleEditSubclassChar(HWND editHwnd, WPARAM key);
    bool HandleEditSubclassPaste(HWND editHwnd);
    static LRESULT CALLBACK EditSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData);

    // Edit autosuggest
    void UpdateEditSuggest();
    void UpdateEditSuggestPopupWindow();
    void CloseEditSuggestPopup();
    void EnsureEditSuggestPopupD2DResources();
    void DiscardEditSuggestPopupD2DResources();
    void RenderEditSuggestPopup();
    void ApplyEditSuggestIndex(size_t index);

    void EnsureEditSuggestWorker();
    void EditSuggestWorker(std::stop_token stopToken);
    void PostEditSuggestResults(uint64_t requestId,
                                bool hasMore,
                                wchar_t directorySeparator,
                                std::wstring&& highlightText,
                                std::vector<std::wstring>&& displayItems,
                                std::vector<std::wstring>&& insertItems);

    // Disk info
    void UpdateDiskInfo();

    // Animation Helpers
    void StartSeparatorAnimation(size_t separatorIndex, float targetAngle);
    static bool SeparatorAnimationTickThunk(void* context, uint64_t nowTickMs) noexcept;
    bool UpdateSeparatorAnimations(uint64_t nowTickMs) noexcept;
    void StopSeparatorAnimation() noexcept;
    void OnFullPathPopupExitMenuLoop(HWND popupHwnd, bool isShortcut);
    void UpdateHoverTimerState() noexcept;

    // State
    wil::unique_hwnd _hWnd;
    HINSTANCE _hInstance                = nullptr;
    UINT _dpi                           = USER_DEFAULT_SCREEN_DPI;
    SIZE _clientSize                    = {0, 0};
    RenderMode _renderMode              = RenderMode::Breadcrumb;
    bool _editMode                      = false;
    bool _trackingMouse                 = false;
    bool _inMenuLoop                    = false;
    bool _menuButtonPressed             = false; // Track if menu is open
    bool _menuButtonHovered             = false; // Track if Section 1 is hovered
    bool _historyButtonHovered          = false; // Track if history button is hovered
    bool _diskInfoHovered               = false; // Track if Section 3 is hovered
    int _hoveredSegmentIndex            = -1;    // Track which segment is hovered (-1 = none)
    int _hoveredSeparatorIndex          = -1;    // Track which separator is hovered (-1 = none)
    bool _editCloseHovered              = false;
    HWND _suppressCtrlBackspaceCharHwnd = nullptr;

    struct EditSuggestItem
    {
        std::wstring display;
        std::wstring insertText;
        bool enabled               = true;
        wchar_t directorySeparator = L'\\';
    };

    wil::unique_hwnd _editSuggestPopup;
    SIZE _editSuggestPopupClientSize = {0, 0};
    int _editSuggestPopupRowHeightPx = 0;
    std::vector<EditSuggestItem> _editSuggestItems;
    uint64_t _editSuggestAdditionalRequestId = 0;
    std::vector<EditSuggestItem> _editSuggestAdditionalItems;
    int _editSuggestHoveredIndex  = -1;
    int _editSuggestSelectedIndex = -1;
    std::wstring _editSuggestHighlightText;

    wil::com_ptr<ID2D1HwndRenderTarget> _editSuggestPopupTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupDisabledTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupHighlightBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupHoverBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupBackgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _editSuggestPopupBorderBrush;

    struct EditSuggestFileSystemInstance
    {
        EditSuggestFileSystemInstance() = default;

        EditSuggestFileSystemInstance(EditSuggestFileSystemInstance&)            = delete;
        EditSuggestFileSystemInstance& operator=(EditSuggestFileSystemInstance&) = delete;

        wil::unique_hmodule module;
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring pluginShortId;
        std::wstring instanceContext;
    };

    struct EditSuggestQuery
    {
        uint64_t requestId = 0;
        wil::com_ptr<IFileSystem> fileSystem;
        std::filesystem::path displayFolder;
        std::filesystem::path pluginFolder;
        std::wstring prefix;
        wchar_t directorySeparator = L'\\';
        std::shared_ptr<EditSuggestFileSystemInstance> keepAlive;
    };

    std::shared_ptr<EditSuggestFileSystemInstance> _editSuggestMountedInstance;

    std::mutex _editSuggestMutex;
    std::condition_variable _editSuggestCv;
    std::optional<EditSuggestQuery> _editSuggestPendingQuery;
    std::jthread _editSuggestThread;
    std::atomic<uint64_t> _editSuggestRequestId = 0;

    struct SiblingPrefetchQuery
    {
        uint64_t requestId = 0;
        wil::com_ptr<IFileSystem> fileSystem;
        std::vector<std::filesystem::path> folders;
    };

    std::mutex _siblingPrefetchMutex;
    std::condition_variable _siblingPrefetchCv;
    std::optional<SiblingPrefetchQuery> _siblingPrefetchPendingQuery;
    std::jthread _siblingPrefetchThread;
    std::atomic<uint64_t> _siblingPrefetchRequestId = 0;

    int _activeSeparatorIndex            = -1; // Track which separator has menu open
    int _menuOpenForSeparator            = -1; // Track which separator's menu is currently displayed
    int _pendingSeparatorMenuSwitchIndex = -1; // Deferred separator menu switch target (-1 = none)
    bool _pendingFullPathPopup           = false;
    POINT _pendingFullPathPopupAnchor    = {};

    wil::unique_hwnd _fullPathPopup;
    wil::unique_hwnd _fullPathPopupEdit;
    bool _fullPathPopupEditMode                       = false;
    bool _fullPathPopupTrackingMouse                  = false;
    int _fullPathPopupActiveSeparatorIndex            = -1;
    int _fullPathPopupMenuOpenForSeparator            = -1;
    int _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;
    int _fullPathPopupHoveredSegmentIndex             = -1;
    int _fullPathPopupHoveredSeparatorIndex           = -1;
    float _fullPathPopupScrollY                       = 0.0f;
    float _fullPathPopupContentHeight                 = 0.0f;
    SIZE _fullPathPopupClientSize                     = {0, 0};
    std::vector<PathSegment> _fullPathPopupSegments;
    std::vector<BreadcrumbSeparator> _fullPathPopupSeparators;
    UINT_PTR _fullPathPopupHoverTimer = 0;

    wil::com_ptr<ID2D1HwndRenderTarget> _fullPathPopupTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupSeparatorBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupHoverBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupPressedBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupAccentBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _fullPathPopupBackgroundBrush;

    // Separator rotation animation
    std::vector<float> _separatorRotationAngles; // Current rotation angle for each separator (0-90 degrees)
    std::vector<float> _separatorTargetAngles;   // Target rotation angle for each separator
    uint64_t _separatorAnimationSubscriptionId = 0;
    uint64_t _separatorAnimationLastTickMs     = 0;
    UINT_PTR _hoverTimer                       = 0; // Timer ID for hover tracking
    static constexpr UINT_PTR HOVER_TIMER_ID   = 2;
    static constexpr UINT HOVER_CHECK_FPS      = 30;
    static constexpr float ROTATION_SPEED      = 600.0f; // Degrees per second (90° in 150ms)

    struct WStringViewHash
    {
        using is_transparent = void;

        size_t operator()(std::wstring_view value) const noexcept
        {
            return std::hash<std::wstring_view>{}(value);
        }

        size_t operator()(const std::wstring& value) const noexcept
        {
            return std::hash<std::wstring_view>{}(value);
        }
    };

    struct WStringViewEq
    {
        using is_transparent = void;

        bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
        {
            return left == right;
        }
    };

    struct BreadcrumbTextLayoutCacheEntry
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        float width = 0.0f;
    };

    static constexpr size_t kMaxBreadcrumbTextLayoutCacheEntries = 256u;

    std::unordered_map<std::wstring, BreadcrumbTextLayoutCacheEntry, WStringViewHash, WStringViewEq> _breadcrumbTextLayoutCache;
    IDWriteFactory* _breadcrumbTextLayoutCacheFactory   = nullptr;
    IDWriteTextFormat* _breadcrumbTextLayoutCacheFormat = nullptr;
    float _breadcrumbTextLayoutCacheHeight              = 0.0f;

    bool _breadcrumbLayoutCacheValid = false;
    std::filesystem::path _breadcrumbLayoutCachePath;
    UINT _breadcrumbLayoutCacheDpi                           = USER_DEFAULT_SCREEN_DPI;
    float _breadcrumbLayoutCacheAvailableWidth               = 0.0f;
    float _breadcrumbLayoutCacheSectionHeight                = 0.0f;
    IDWriteFactory* _breadcrumbLayoutCacheFactory            = nullptr;
    IDWriteTextFormat* _breadcrumbLayoutCachePathFormat      = nullptr;
    IDWriteTextFormat* _breadcrumbLayoutCacheSeparatorFormat = nullptr;

    // Path data
    std::optional<std::filesystem::path> _currentPath;
    std::optional<std::filesystem::path> _currentPluginPath;
    std::optional<std::filesystem::path> _currentEditPath;
    std::wstring _currentInstanceContext;
    wil::com_ptr<IFileSystem> _fileSystemPlugin;
    wil::com_ptr<IFileSystemIO> _fileSystemIo;
    wil::com_ptr<INavigationMenu> _navigationMenu;
    wil::com_ptr<IDriveInfo> _driveInfo;
    std::wstring _pluginShortId;
    std::vector<PathSegment> _segments;
    std::vector<BreadcrumbSeparator> _separators;
    std::deque<std::filesystem::path> _pathHistory; // Max 20 items

    Common::Settings::Settings* _settings = nullptr;

    PathChangedCallback _pathChangedCallback;
    RequestFolderViewFocusCallback _requestFolderViewFocusCallback;
    FocusRegion _focusedRegion = FocusRegion::Path;

    // Disk data
    std::wstring _diskSpaceText;
    uint64_t _freeBytes  = 0;
    uint64_t _totalBytes = 0;
    uint64_t _usedBytes  = 0;
    bool _hasTotalBytes  = false;
    bool _hasFreeBytes   = false;
    bool _hasUsedBytes   = false;
    std::wstring _volumeLabel;
    std::wstring _fileSystem;
    std::wstring _driveDisplayName;

    // Menu icons - managed with WIL
    std::vector<wil::unique_hbitmap> _menuBitmaps;
    int _menuIconSize         = 0;
    bool _showMenuSection     = false;
    bool _showDiskInfoSection = false;

    struct MenuItemData
    {
        std::wstring text;
        std::wstring shortcut;
        HBITMAP bitmap         = nullptr;
        wchar_t glyph          = 0;
        bool separator         = false;
        bool header            = false;
        bool hasSubMenu        = false;
        bool useMiddleEllipsis = false;
    };

    std::vector<std::unique_ptr<MenuItemData>> _menuItemData;
    wil::unique_hbrush _menuBackgroundBrush;
    int _themedMenuMaxWidthPx           = 0;
    bool _themedMenuUseMiddleEllipsis   = false;
    bool _themedMenuUseEditSuggestStyle = false;

    enum class MenuActionType : uint8_t
    {
        NavigatePath,
        Command,
    };

    struct MenuAction
    {
        UINT menuId         = 0;
        MenuActionType type = MenuActionType::NavigatePath;
        std::wstring path;
        unsigned int commandId = 0;
    };

    std::vector<MenuAction> _navigationMenuActions;
    std::vector<MenuAction> _driveMenuActions;

    bool ExecuteNavigationMenuAction(UINT menuId);
    bool ExecuteDriveMenuAction(UINT menuId);

    // GDI resources (Sections 1 & 3) - managed with WIL
    wil::unique_hfont _pathFont; // For Edit control
    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _borderBrush;
    wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> _borderPen;

    void UpdateEffectiveTheme() noexcept;

    NavigationViewTheme _baseTheme;
    NavigationViewTheme _theme;
    MenuTheme _menuTheme;
    AppTheme _appTheme;

    enum class ModernDropdownKind : uint8_t
    {
        None,
        History,
        Siblings,
    };

    ModernDropdownKind _navDropdownKind = ModernDropdownKind::None;
    std::vector<std::filesystem::path> _navDropdownPaths;
    wil::unique_hwnd _navDropdownCombo;
    wil::unique_hfont _menuFont;
    UINT _menuFontDpi = USER_DEFAULT_SCREEN_DPI;
    wil::unique_hfont _menuIconFont;
    UINT _menuIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
    bool _menuIconFontValid = false;

    bool _paneFocused = false;

    // Direct2D resources (Section 2)
    wil::com_ptr<ID2D1Factory1> _d2dFactory;
    wil::com_ptr<ID3D11Device> _d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> _d3dContext;
    wil::com_ptr<ID2D1Device> _d2dDevice;
    wil::com_ptr<ID2D1DeviceContext> _d2dContext;
    wil::com_ptr<IDXGISwapChain1> _swapChain;
    wil::com_ptr<ID2D1Bitmap1> _d2dTarget;

    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<IDWriteTextFormat> _pathFormat;      // 12pt Segoe UI
    wil::com_ptr<IDWriteTextFormat> _separatorFormat; // For › symbol

    wchar_t _breadcrumbSeparatorGlyph = L'\u203A'; // › (fallback when Segoe Fluent Icons isn't available)
    wchar_t _historyChevronGlyph      = L'\u25BE'; // ▾ (fallback when Segoe Fluent Icons isn't available)
    bool _dwriteFluentIconsValid      = false;

    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;      // RGB(32,32,32)
    wil::com_ptr<ID2D1SolidColorBrush> _separatorBrush; // RGB(120,120,120)
    wil::com_ptr<ID2D1SolidColorBrush> _hoverBrush;     // RGB(243,243,243)
    wil::com_ptr<ID2D1SolidColorBrush> _pressedBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _accentBrush;        // System accent color
    wil::com_ptr<ID2D1SolidColorBrush> _rainbowBrush;       // Per-item rainbow accent (Rainbow theme)
    wil::com_ptr<ID2D1SolidColorBrush> _backgroundBrushD2D; // RGB(250,250,250)

    // Section 1 (menu button) Direct2D resources
    wil::com_ptr<ID2D1Bitmap1> _menuIconBitmapD2D; // D2D version of menu icon

    bool _hasPresented       = false; // Track if initial Present has been called
    bool _deferredInitPosted = false;
    bool _deferPresent       = false;
    bool _queuedPresentFull  = false;
    std::optional<RECT> _queuedPresentDirtyRect;

    // Command IDs
    enum
    {
        ID_MENU_BUTTON = 100,
        ID_PATH_EDIT,
        ID_HISTORY_BUTTON,
        ID_DISK_STATIC,

        ID_NAV_MENU_BASE = 200,
        ID_NAV_MENU_MAX  = 399,

        ID_DRIVE_MENU_BASE = 500,
        ID_DRIVE_MENU_MAX  = 599,

        ID_SIBLING_BASE       = 600, // 600-699 for sibling folders
        ID_NAV_DROPDOWN_COMBO = 700,
    };
};

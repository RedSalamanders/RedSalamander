# NavigationView Specification

## Overview

**NavigationView** is a navigation bar component displayed at the top of FolderView, providing quick access to path navigation, history, disk switching, and storage information.

**Current implementation (code truth)**:
- Renders **all visible sections** (Drive/Menu, Path/Breadcrumb, History button, Disk Info) with **Direct2D/DirectWrite** to a DXGI swap chain.
- Startup performance: Direct2D/D3D swap-chain initialization is **deferred until after the first paint** (via `WndMsg::kNavigationViewDeferredInit`) so `WM_CREATE` stays fast; the first paint uses only GDI background/border.
- Graphics resources: NavigationView instances **share** the process-wide D2D/DWrite/D3D device objects; swap chains and device contexts remain per-window.
- Uses Win32 for the window procedure and the transient **Edit control** used in edit mode.
- Uses GDI/owner-draw only for **popup menus** (WM_MEASUREITEM/WM_DRAWITEM) and for simple background/border painting.

## Localization

- All fixed user-facing labels (menu items, message boxes, captions) must be loaded from `.rc` resources (see `Specs/LocalizationSpec.md`).
- Dynamic UI labels (current path, drive names, WSL distro names, disk sizes, etc.) remain runtime data, but any surrounding UI text should come from resource **format strings**.

## Path Model (Canonical vs Plugin vs UI)

NavigationView participates in a host-wide “location” model with **two different path strings**:

- **pluginPath** (passed to plugins):
  - Passed to `IFileSystem` / `IDriveInfo` / `DirectoryInfoCache`.
  - Never includes `<shortId>:` and never includes the mount delimiter `|`.
  - For `file`: a Windows absolute path (`C:\...`, UNC, `\\?\...`).
  - For non-`file`: an absolute plugin path rooted at `/` using `/` separators.
- **canonical location** (stored in history/settings and used for host routing):
  - For `file`: same as `pluginPath`.
  - For non-`file`: `<shortId>:<pluginPath>`.
  - For non-`file` with mount context: `<shortId>:<instanceContext>|<pluginPath>` (instanceContext is opaque and is persisted by the host).

**UI rules**:
- Breadcrumb/full-path display uses **pluginPath only** (no `<shortId>` segment and no mount context); non-`file` roots display as `/`.
- Edit mode shows:
  - `file`: `C:\a\b` (no `file:` prefix)
  - non-`file`: `<shortId>:/a/b`
  - mounted non-`file` (Change Directory): `<instanceContext>|/a/b` (no `<shortId>:` prefix; host prepends the current `<shortId>:` when routing the accepted value)

**7z mount shorthand**:
- Typing `7z:<zipPath>` (no `|` and `<zipPath>` does not start with `/` or `\`) is interpreted as mounting `<zipPath>` and navigating to `/`.

## Recent Architectural Improvements (December 2025)

### 1. Modern Windows Shell API Migration
✅ **Replaced deprecated APIs** with Microsoft-recommended alternatives:
- **Old**: `SHGetFolderPath` with `CSIDL_*` integer constants (deprecated)
- **New**: `SHGetKnownFolderPath` with `KNOWNFOLDERID` GUIDs (modern)
- **Memory Management**: `wil::unique_cotaskmem_string` RAII wrappers (automatic cleanup)
- **Benefits**: Type safety, future-proof, no manual `CoTaskMemFree` calls

**Affected Areas**:
- Menu button icon detection (special folder matching)
- Menu dropdown special folder paths (Desktop, Documents, Downloads, etc.)
- Navigation command handling

### 2. Plugin-Provided Navigation Menu + Drive Info
✅ **Menu and disk info sourced from active plugins** (December 2025):
- NavigationView queries `INavigationMenu` and `IDriveInfo` via `QueryInterface` on the active `IFileSystem`.
- NavigationView registers a **non-COM** `INavigationMenuCallback` via `INavigationMenu::SetCallback(callback, cookie)` so the plugin can request navigation (e.g., from `ExecuteMenuCommand`), and clears it with `SetCallback(nullptr, nullptr)` when switching/unloading.
- Menu entries are returned as raw items (`NavigationMenuItem`) with labels, optional icon paths, and optional commands.
- Drive information is returned as structured data (`DriveInfo`: display name, volume label, file system, total/used/free bytes).
- Path parameters passed to `IDriveInfo` are **plugin paths** (no `<shortId>:` prefix and no `<instanceContext>|` mount prefix; mount context is configured separately via `IFileSystemInitialize::Initialize` when applicable).
- Sections are hidden when interfaces are missing or return no data.
- Icon resolution uses `IconCache::QuerySysIconIndexForPath` against `iconPath`/`path`.

**Files Added/Updated**:
- `Common/PlugInterfaces/NavigationMenu.h`
- `Common/PlugInterfaces/DriveInfo.h`
- `RedSalamander/NavigationView.cpp` (menu + disk info rendering)
- `Plugins/FileSystem/FileSystem.cpp` (default menu and disk info provider)

### 3. DPI Awareness and Icon Quality Improvements
✅ **Per-Monitor V2 DPI awareness** (December 2025):
- **Application Manifest**: Added `exe.manifest` with `PerMonitorV2` DPI awareness
- **WM_DPICHANGED Handling**: Now properly receives and processes DPI change messages
- **Icon Size Management**: IconCache adjusted to extract optimal icon sizes for display DPI
- **Menu Icons**: Fixed to use 96 DPI physical pixels (GDI menus are NOT DPI-aware)

**Key Improvements**:
- **Manifest Settings**: Enabled long path support, UTF-8 code page, segment heap
- **Icon Extraction**: `IconCache::ExtractSystemIcon()` defaults to 16 DIP for FolderView
- **WIC Pipeline**: Fixed `CreateMenuBitmapFromIcon()` to copy pixels from converter output
- **Transparency Fix**: Menu icons now have perfect transparency without black borders
- **Menu Icon Size**: Fixed `_menuIconSize` to NOT scale with DPI (always `GetSystemMetrics(SM_CXSMICON)`)

**Technical Details**:
- GDI menus (CreatePopupMenu, AppendMenu) are NOT DPI-aware
- Menu item bitmaps must be physical pixels at 96 DPI regardless of window DPI
- FolderView icons scale with DPI for crisp display at all scaling levels
- WIC conversion uses GUID_WICPixelFormat32bppPBGRA for proper alpha channel

### 4. Section 1 Direct2D Migration for DPI Awareness
✅ **Menu button migrated from GDI to Direct2D** (December 2025):
- **Old Architecture**: GDI owner-draw button (HWND with WM_DRAWITEM)
- **New Architecture**: Direct2D rendering matching Section 2 pipeline
- **Rationale**: GDI controls cannot render DPI-aware icons properly at high DPI
- **Benefits**: Crisp icon scaling at 125%, 150%, 175%, 200% DPI

**Technical Changes**:
- Removed `HWND _menuButton` member and CreateWindowExW button creation
- Added `wil::com_ptr<ID2D1Bitmap1> _menuIconBitmapD2D` for DPI-aware icon storage
- Removed `wil::unique_hicon _cachedMenuButtonIcon` (replaced by D2D bitmap)
- Implemented `UpdateMenuIconBitmap()` to resolve a **system image list icon index** (special-folder prefix match or drive/path) and fetch a cached bitmap via `IconCache::GetIconBitmap(iconIndex, _d2dContext.get())`
- Implemented `RenderDriveSection()` method for Direct2D rendering with hover/press states
- Centralized all Shell icon queries in `IconCache` (`QuerySysIconIndexFor*`); NavigationView does not call `SHGetFileInfoW`

**Rendering Pipeline**:
1. Resolve system icon index for current path (special folder → drive root → current folder)
2. Get cached `ID2D1Bitmap1` via `IconCache::GetIconBitmap(...)`
3. Render with state-aware background colors (normal/hover/pressed)
4. Hamburger icon fallback drawn with Direct2D primitives if no icon available
5. Present with dirty region optimization (DXGI_PRESENT_PARAMETERS)

**Icon Selection Logic** (from old OnDrawItem):
```cpp
// Priority: Special folder icon > Drive root icon > Hamburger fallback
if (under special folder like Documents\foo\bar)
    Use Documents icon
else if (drive path like C:\foo\bar)
    Use C:\ icon
else
    Draw hamburger icon (3 horizontal lines)
```

**State Management**:
- `_menuButtonHovered`: Timer-based hover tracking (30 FPS polling)
- `_menuButtonPressed`: Set during TrackPopupMenu, cleared after
- Background colors: Normal RGB(250,250,250), Hover RGB(243,243,243), Pressed RGB(230,230,230)

**Call Sites Updated**:
- `OnPaint()`: Calls `RenderDriveSection()`, `RenderPathSection()`, `RenderHistorySection()`, `RenderDiskInfoSection()`
- `OnLButtonDown()`: Section 1 hit testing with `PtInRect(&_sectionDriveRect)` → `ShowMenuDropdown()`
- `OnTimer()`: Hover tracking calls `RenderDriveSection()` / `RenderHistorySection()` / `RenderDiskInfoSection()` on state change
- `ShowMenuDropdown()`: Calls `RenderDriveSection()` for pressed/normal states
- `SetPath()`: Calls `UpdateMenuIconBitmap()` to refresh icon
- `OnDpiChanged()`: Regenerates icon at new DPI via `UpdateMenuIconBitmap()`
- `EnsureD2DResources()`: Initializes icon bitmap when D2D context ready
- `DiscardD2DResources()`: Clears `_menuIconBitmapD2D = nullptr`

**Files Modified**:
- `NavigationView.h`: Removed HWND, added D2D members and method declarations
- `NavigationView.cpp`: Removed WM_DRAWITEM/OnDrawItem/button creation, added Direct2D rendering

### 5. Flat Button Rendering & Enhanced UI
✅ **Menu button now uses flat design**:
- Removed 3D borders for modern appearance
- Press state kept when menu is open
- Current folder icon or hamburger icon fallback (now rendered with Direct2D)
- Keyboard focus ring uses small rounded corners (see `Specs/VisualStyleSpec.md`).

✅ **Breadcrumb separator enhancements**:
- Hover and pressed states for separators
- Sibling + history dropdowns use the ModernCombo popup visuals (40 DIP rows, rounded highlight + accent bar; no per-item icons), but open below the NavigationView and can expand to the remaining main-window height (themed scrollbar). In the sibling dropdown, the current folder stays marked with the accent bar while hover/keyboard selection moves independently.
- Navigation only commits on click/Enter; click-outside cancels (no navigation)
- Full clickable zone highlighting in breadcrumb segments

### 6. Interactive States and Animation System (December 2025)
✅ **Comprehensive hover effects across all sections**:
- **Section 1 (Menu Button)**: Timer-based hover polling with `_menuButtonHovered` state
  - **Implementation**: 30 FPS polling timer (`HOVER_TIMER_ID = 2`) checks cursor position via `GetCursorPos` + `ScreenToClient` + `PtInRect`
  - **Rationale**: Keeps hover states consistent even during modal menu tracking (`TrackPopupMenu`) and prevents stale highlights
  - **Benefits**: No reliance on `WM_MOUSELEAVE`; explicit clearing when cursor leaves regions
- **Section 2 (Breadcrumb Segments)**: Direct2D hover backgrounds via timer polling
  - **Coordinate Transform**: Window coordinates converted to Section 2 local space before hit testing
  - **Hover Cleanup**: Explicitly clears hover when cursor leaves Section 2 bounds to prevent stale highlights
- **Section 2 (Separators)**: Timer-based hover detection with pressed states persisting during menu
  - **Increased Size**: Font 50% larger (36pt at 96 DPI), width 24→32px, rect 20x24→32x36
  - **Text Alignment**: Uses full section height with `PARAGRAPH_ALIGNMENT_CENTER`, Y position at 0.0f
- **Section 3 (History Button)**: Direct2D-rendered region with `_historyButtonHovered` updated by timer
- **Section 4 (Disk Info)**: Direct2D-rendered region with `_diskInfoHovered` updated by timer (disk text + progress bar)
- **Popup menus** (Drive/Menu + Disk Info): Win32 owner-draw with `MenuTheme` (`WM_MEASUREITEM`/`WM_DRAWITEM`)
- **Consistent Colors**: Hover/pressed colors come from `NavigationViewTheme` (e.g., `backgroundHover`, `backgroundPressed`)

✅ **Hover tracking system** (All sections):
```cpp
// Timer is toggled on-demand (menus/edit mode only)
void UpdateHoverTimerState() noexcept {
    const bool shouldRun = _editMode || _inMenuLoop;
    if (shouldRun && _hoverTimer == 0) {
        _hoverTimer = SetTimer(_hWnd, HOVER_TIMER_ID, 1000 / HOVER_CHECK_FPS, nullptr);
    } else if (!shouldRun && _hoverTimer != 0) {
        KillTimer(_hWnd, HOVER_TIMER_ID);
        _hoverTimer = 0;
    }
}

// Start/stop points
case WM_ENTERMENULOOP: _inMenuLoop = true; UpdateHoverTimerState(); break;
case WM_EXITMENULOOP:  _inMenuLoop = false; UpdateHoverTimerState(); break;
EnterEditMode():       _editMode = true; UpdateHoverTimerState();
ExitEditMode():        _editMode = false; UpdateHoverTimerState();

// OnTimer handler (30 FPS polling)
void OnTimer(UINT_PTR timerId) {
    if (timerId == HOVER_TIMER_ID) {
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(_hWnd, &pt);
        
        RECT clientRect;
        GetClientRect(_hWnd, &clientRect);
        bool inClient = PtInRect(&clientRect, pt);
        
        // Section 1 hover detection
        bool menuButtonHovered = inClient && PtInRect(&_sectionDriveRect, pt);
        if (menuButtonHovered != _menuButtonHovered) {
            _menuButtonHovered = menuButtonHovered;
            RenderDriveSection();
        }
        
        // History hover detection
        bool historyButtonHovered = inClient && PtInRect(&_sectionHistoryRect, pt);
        if (historyButtonHovered != _historyButtonHovered) {
            _historyButtonHovered = historyButtonHovered;
            RenderHistorySection();
        }

        // Disk info hover detection
        bool diskInfoHovered = inClient && PtInRect(&_sectionDiskInfoRect, pt);
        if (diskInfoHovered != _diskInfoHovered) {
            _diskInfoHovered = diskInfoHovered;
            RenderDiskInfoSection();
        }
        
        // Section 2 segments/separators hover (with coordinate transform)
        // ... check each segment/separator with local coordinates ...
        // Explicitly clear hover when cursor leaves Section 2
    }
}
```
**Why Timer Polling (Only During Menus/Edit Mode)?**
- Hover state remains correct while popup menus are tracking (`TrackPopupMenu` is modal and can disrupt normal mouse message flow)
- Edit mode uses a child edit control, so hover for the chrome (close button) is easier to keep consistent via polling
- Outside of those cases, hover is driven by `WM_MOUSEMOVE` / `WM_MOUSELEAVE` to avoid an always-on timer

**Timer Configuration**:
```cpp
static constexpr UINT_PTR HOVER_TIMER_ID = 2;
static constexpr UINT HOVER_CHECK_FPS = 30;
UINT_PTR _hoverTimer = 0;
```

✅ **Owner-drawn popup menus** (MenuTheme):
- All dropdown menus are `HMENU` popup menus themed via `WM_MEASUREITEM`/`WM_DRAWITEM` using `MenuTheme` colors.

✅ **Coordinate space transformation for Section 2**:
- **Problem**: Direct2D uses local coordinate space after `SetTransform`, Win32 mouse events use window coordinates
- **Solution**: Transform window coordinates to Section 2 local space in `OnMouseMove` and `OnLButtonDown`:
  ```cpp
  float localX = static_cast<float>(pt.x - _sectionPathRect.left);
  float localY = static_cast<float>(pt.y - _sectionPathRect.top);
  ```
- **Impact**: Fixes hover/click misalignment bug where highlights appeared offset from actual click area

✅ **Smooth separator rotation animation**:
- **Visual Effect**: Separator (›) rotates 90° clockwise when menu opens, reverses on close
- **Duration**: 150ms per direction
- **Frame Rate**: ~60 FPS via `Ui::AnimationDispatcher` (single shared 16ms `WM_TIMER`)
- **Implementation**: Direct2D `Matrix3x2F::Rotation` transform around separator center
- **State Tracking**: Per-separator angle vectors with linear interpolation

✅ **Menu switching behavior**:
- **User Experience**: Clicking different separator while menu is open closes current menu and opens new one
- **Implementation**: `WM_CANCELMODE` to force menu closure, `PostMessage(WM_USER+100)` to defer new menu opening
- **Modal Handling**: Works correctly with `TrackPopupMenu`'s blocking behavior
- **State Cleanup**: `WM_EXITMENULOOP` triggers reverse animation and state reset
- **Implemented**: Hover-based menu switching during `TrackPopupMenu` tracking using timer polling + `WM_CANCELMODE` (no `SetWindowsHookEx`)
- **Safety**: Hover tracking ignores the cursor while it is over a Win32 menu window (`#32768`) to avoid accidental switches while navigating the menu
- **Isolation**: While the full-path popup window is open, `NavigationView` hover tracking is suspended; the full-path popup runs its own hover polling

**Files Modified**:
- `NavigationView.h` (line 166) - Added `bool _menuButtonHovered` member after `_menuButtonPressed`
- `NavigationView.cpp` (lines 461-610) - Restructured `OnDrawItem` to if/else if chain for both buttons
- `NavigationView.cpp` (lines 729-738) - Added manual Section 1 hover tracking in `OnMouseMove`
- `NavigationView.cpp` (line 889) - Changed Section 3 cursor from `IDC_HAND` to `IDC_ARROW`
- `NavigationView.cpp` (lines 1885-1897) - Removed `MF_DISABLED` from 8 disk info menu items
- `NavigationView.cpp` (line 1086) - Increased separator font from `barHeight` to `barHeight * 1.5f`
- `NavigationView.cpp` (line 1417) - Increased separator width from 24.0f to 32.0f
- `NavigationView.cpp` (line 1440) - Fixed layout height to use full section height for vertical centering
- `NavigationView.cpp` (line 1536) - Changed DrawTextLayout Y position from `y - metrics.height / 2` to 0.0f
- `NavigationView.cpp` (`UpdateBreadcrumbLayout`, `RenderBreadcrumbs`, `OnLButtonDown`) - Ellipsis collapse + separator rendering/click handling

**New Members**:
```cpp
bool _menuButtonHovered;                   // Section 1 manual hover state
bool _diskInfoHovered;                     // Section 3 hover state (Note: using ODS_HOTLIGHT, not manual)
int _menuOpenForSeparator;                 // Which separator's menu is open
std::vector<float> _separatorRotationAngles;  // Current rotation (0-90°)
std::vector<float> _separatorTargetAngles;    // Target rotation
uint64_t _separatorAnimationSubscriptionId;   // AnimationDispatcher subscription ID
uint64_t _separatorAnimationLastTickMs;       // Last dispatcher tick (ms)
```

**New Methods**:
```cpp
void OnTimer(UINT_PTR timerId);            // Hover polling (menus/edit mode)
void OnEnterMenuLoop(bool isTrackPopupMenu);
void OnExitMenuLoop(bool isShortcut);      // Menu cleanup + reverse animation
void StartSeparatorAnimation(size_t, float);  // Begin rotation
static bool SeparatorAnimationTickThunk(void*, uint64_t) noexcept;
bool UpdateSeparatorAnimations(uint64_t nowTickMs) noexcept;
void StopSeparatorAnimation() noexcept;
```

**Bug Fixes (December 2025)**:
1. ✅ Section 1 hover not working → Manual tracking with `_menuButtonHovered`
2. ✅ Section 3 hover not working + wrong cursor → BUTTON with manual hover + `IDC_ARROW`
3. ✅ Section 3 menu shows disabled gray text → Removed `MF_DISABLED` from info items
4. ✅ Menu switching during tracking → Timer polling + `WM_CANCELMODE` + deferred open (no mouse hook)
5. ✅ Section 2 text vertically misaligned → Full height layout + `PARAGRAPH_ALIGNMENT_CENTER`
6. ✅ Section 2 separators too small → 50% font increase + wider rects (32x36)

## Architecture Decision: Hybrid Approach

### Rendering Strategy

**Section 1 (Menu Button)**: Direct2D (migrated December 2025)
- ✅ **DPI-aware icon rendering** with ID2D1Bitmap1
- ✅ **State-aware backgrounds** (normal/hover/pressed)
- ✅ **Special folder detection** for intelligent icon selection
- ✅ **Hamburger icon fallback** drawn with Direct2D primitives
- ✅ **Manual hit testing** with PtInRect (no HWND needed)

**Section 2 (Path Display)**: Direct2D/DirectWrite
- ✅ **Advanced text rendering** with custom layouts
- ✅ **Breadcrumb navigation** with clickable path segments and clickable separators to display sibbling folders
- ✅ **Syntax highlighting** (drive, folders, separators)
- ✅ **Smooth transitions** between display and edit modes
- ✅ **Custom text truncation** and ellipsis positioning in the middle of the path
- ✅ **Rich visual feedback** (hover effects on segments)

**Section 3 (History Button)**: Direct2D/DirectWrite
- Small clickable region rendered with Direct2D (no HWND)
- Opens history dropdown (popup menu) via `TrackThemedPopupMenuReturnCmd` (`TrackPopupMenu` + `TPM_RETURNCMD`)

**Section 4 (Disk Info)**: Direct2D/DirectWrite
- Disk space text + progress bar rendered with Direct2D
- Opens disk info dropdown (popup menu) via `TrackThemedPopupMenuReturnCmd` (`TrackPopupMenu` + `TPM_RETURNCMD`)

**Popup Menus**: Win32/GDI owner-draw
- All dropdowns are Win32 popup menus themed via `MenuTheme` using `WM_MEASUREITEM`/`WM_DRAWITEM`

### Why DirectX for Section 2?

✅ **Advantages:**
- **Breadcrumbs**: Click on individual path segments (C: > Users > Documents)
- **Visual Hierarchy**: Different colors for different path components
- **Smooth Animations**: Fade between modes, hover effects
- **Custom Text Layout**: Precise control over character positioning
- **High-Quality Rendering**: ClearType, sub-pixel positioning
- **Future-Proof**: Easy to add icons, badges, overlays

❌ **GDI Limitations for Path Display:**
- Complex hit-testing for clickable segments
- Limited text effects and styling
- No smooth transitions
- Harder to implement breadcrumb UI

## Architecture

### Component Type
- **Class**: `NavigationView`
- **Window Type**: Win32 child window
- **Rendering**: Direct2D/DirectWrite for all sections (DXGI swap chain); GDI owner-draw for popup menus; Win32 Edit control in edit mode
- **Parent**: Main application window (positioned above FolderView)

### Files
- **Header**: `RedSalamander/NavigationView.h`
- **Internal Helpers**: `RedSalamander/NavigationViewInternal.h`
- **Implementation (split)**:
  - `RedSalamander/NavigationView.cpp` (window lifecycle + message dispatch)
  - `RedSalamander/NavigationView.Interaction.cpp` (mouse/keyboard/timers/focus)
  - `RedSalamander/NavigationView.Rendering.cpp` (swapchain/D2D/DWrite resources + sections 1/3/4 rendering)
  - `RedSalamander/NavigationView.Breadcrumb.cpp` (breadcrumb layout/overflow + section 2 rendering/animation)
  - `RedSalamander/NavigationView.Menus.cpp` (popup menus + owner-draw/theming)
  - `RedSalamander/NavigationView.Edit.cpp` (edit mode + edit-suggest popup/workers)
  - `RedSalamander/NavigationView.FullPathPopup.cpp` (full-path popup window)
- **Integration**: Modified `RedSalamander/RedSalamander.cpp`

## Visual Layout

```text
┌───────────────────────────────────────────────────────────────────────┐
│ [≡] │ C: › Users › Username › Documents › Projects │ [⩔] │    256 GB  │
└───────────────────────────────────────────────────────────────────────┘
  ^                          ^                   ^           ^
Drive/Menu              Path (Breadcrumb)     History      Disk Info
(28 DIP)                (Dynamic width)       (24 DIP)     (120 DIP)
```

### Dimensions
- **Total Height**: **24 DIP** (32 pixels at 96 DPI, scales with DPI)
- **Drive/Menu Width**: **28 DIP** (matches `kDriveSectionWidth`)
- **Path Width**: **Dynamic** (fills available space between drive/history/disk)
- **History Width**: **24 DIP** (matches `kHistoryButtonWidth`)
- **Disk Info Width**: **120 DIP** (matches `kDiskInfoSectionWidth`)
- **Conditional Sections**: Drive/Menu hidden without `INavigationMenu`; Disk Info hidden without `IDriveInfo` (Path expands to fill).

### Layout Strategy
✅ **Fixed height at top of FolderView area**
- NavigationView occupies top 24 DIP of parent window
- FolderView starts at Y = 24 DIP (scaled for DPI)
- Both resize together on `WM_SIZE`

### Color Scheme (Windows 11 Style)
- **Background**: `RGB(250, 250, 250)` (light gray)
- **Border Bottom**: `RGB(225, 225, 225)` (1px separator)
- **Text**: `RGB(32, 32, 32)` (dark gray)
- **Breadcrumb Separator**: `RGB(120, 120, 120)` (›)
- **Segment Hover**: `RGB(243, 243, 243)` (subtle background)
- **Segment Active**: System accent color underline
- **Edit Focus**: System accent color border (2 DIP)

## Section 1: Menu Dropdown Button

### Visual Design (Direct2D - December 2025)
- **Type**: Direct2D rendering (no HWND, manual hit testing)
- **Icon**: Context-aware (special folder/drive icon) or hamburger fallback (≡)
- **Size**: 28×24 DIP button area within `_sectionDriveRect`
- **Rendering**: Direct2D with `ID2D1Bitmap1` for DPI-aware icons
- **States**: Normal, Hover, Pressed (via `_menuButtonHovered`, `_menuButtonPressed`)

### Architecture Migration (December 2025)

**Old Architecture (GDI)**:
- Win32 `BUTTON` control with `BS_OWNERDRAW` style
- `WM_DRAWITEM` handler with GDI rendering (`DrawIconEx`, `LineTo`)
- `HWND _menuButton` managed button lifecycle
- Icon stored as `wil::unique_hicon _cachedMenuButtonIcon`
- **Problem**: GDI cannot render DPI-aware icons properly at high DPI

**New Architecture (Direct2D)**:
- No HWND - pure Direct2D rendering in parent window's device context
- Manual hit testing with `PtInRect(&_sectionDriveRect, pt)` in `OnLButtonDown`
- Icon stored as `wil::com_ptr<ID2D1Bitmap1> _menuIconBitmapD2D`
- DPI-aware bitmap scaling via Direct2D's automatic DPI handling
- **Benefits**: Crisp icons at 125%, 150%, 175%, 200% DPI scaling

### Implementation (Current Code)

**Authoritative implementation lives in code**:
- `RedSalamander/NavigationView.h`
- `RedSalamander/NavigationView.cpp`

**Key methods (menu button path icon)**:
- `UpdateMenuIconBitmap()`: resolves a **system image list icon index** for the current path (special-folder prefix match, else drive root, else current folder) using `IconCache` query helpers, then populates `_menuIconBitmapD2D` via `IconCache::GetIconBitmap(iconIndex, _d2dContext.get())`.
- `RenderDriveSection()`: fills `_sectionDriveRect` using `NavigationViewTheme` colors and draws `_menuIconBitmapD2D` (or a hamburger fallback), then presents with a dirty rect.

**Per-section rendering**:
- Painting is split into `RenderDriveSection()`, `RenderPathSection()`, `RenderHistorySection()`, and `RenderDiskInfoSection()`.
- Each render call performs `BeginDraw()`/`EndDraw()` and uses `Present1` with a dirty rect via `Present(...)` for partial updates.

### Menu Button Icon Logic

**Contextual Icon Display** (lines 520-570):

The menu button displays a contextual icon based on the current path 

**Priority Order:**
1. **Special folder icon** if current path is under a known special folder root (Desktop, Documents, Downloads, Pictures, Music, Videos, OneDrive)
2. **Drive root icon** if on a drive but not under a special folder (e.g., `C:\SomeFolder` shows hard drive icon)
3. **Fallback hamburger icon** (≡) if neither condition is met

**Path Matching with Separator Validation**:

Critical logic to prevent false matches where a path merely shares a textual prefix.

**Modern API Usage**: Special folder roots and their icon indices are maintained by `IconCache` (using `KNOWNFOLDERID` + PIDL-based icon index lookup). NavigationView uses `IconCache::TryGetSpecialFolderForPathPrefix()` for boundary-aware, case-insensitive prefix matching.

```cpp
// For `file` plugin paths, NavigationView uses the current pluginPath (no `<shortId>:` prefix).
const std::wstring currentPath = _currentPluginPath.value().wstring();
int iconIndex                  = -1;

if (_currentPluginPath.value().has_root_path())
{
    if (const auto special = IconCache::GetInstance().TryGetSpecialFolderForPathPrefix(currentPath); special.has_value())
    {
        iconIndex = special->iconIndex;
    }

    if (iconIndex < 0)
    {
        const std::wstring rootPath = _currentPluginPath.value().root_path().wstring();
        iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(rootPath.c_str(), 0, false).value_or(-1);
    }
}

if (iconIndex < 0)
{
    iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(currentPath.c_str(), 0, false).value_or(-1);
}

_menuIconBitmapD2D = (iconIndex >= 0) ? IconCache::GetInstance().GetIconBitmap(iconIndex, _d2dContext.get()) : nullptr;
```

**Examples**:
- `C:\Users\Name\Documents\Project` → Documents folder icon ✅
- `C:\aUser\Documents` → Hard drive icon (not under real Documents) ✅
- `C:\` → Hard drive icon ✅
- `D:\Games` → D: drive icon ✅
- Network path with no match → Hamburger icon ✅

### Menu Display (December 2025)
✅ **Left-aligned with button border**:
- **Alignment**: Menu left edge aligns with Section 1 button left edge
- **Implementation**: `TrackThemedPopupMenuReturnCmd(menu, TPM_LEFTALIGN | TPM_TOPALIGN, pt, ...)`
- **Previous**: Was right-aligned (`TPM_RIGHTALIGN` with `rc.right`), visually inconsistent
- **Rationale**: Provides consistent visual alignment with button boundary

```cpp
void ShowMenuDropdown() {
    HMENU menu = CreatePopupMenu();
    // ... populate menu ...
    
    RECT rc = _sectionDriveRect;
    POINT pt = {rc.left, rc.bottom};  // Left edge, not right
    ClientToScreen(_hWnd, &pt);
    
    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_LEFTALIGN | TPM_TOPALIGN, pt, _hWnd);
    if (selectedId != 0) {
        // Handle selected menu item.
    }
    DestroyMenu(menu);
}
```

### Menu Content

#### Quick Access Section
```text
[Icon] Desktop          
[Icon] Documents        
[Icon] Downloads        
[Icon] Pictures         
[Icon] Music            
[Icon] Videos           
──────────────────────
```

#### Plugin-Provided Sections (Dynamic)
Menu content comes from the active file system plugin via `INavigationMenu`.
NavigationView renders **raw menu entries** and does not enumerate WSL distributions or drives itself.

When the active plugin is **not** `file`, NavigationView appends a bottom:
- `Connections...` item to open Connection Manager (protocol-filtered when the active plugin supports Connection Manager).
- `---`
- `Change Drive >` submenu containing the **standard FileSystem** navigation menu items (drives + known folders), so users can jump back to the Windows file system quickly.

**Typical FileSystem plugin composition**:
```text
[Icon] Desktop          
[Icon] Documents        
[Icon] Downloads        
[Icon] Pictures         
[Icon] Music            
[Icon] Videos           
──────────────────────────────
[Icon] Ubuntu               
[Icon] Debian               
[Icon] Connections...
──────────────────────────────  
[Icon] C:\ Local Disk  156 GB
[Icon] D:\ Data        1.20 TB
──────────────────────────────
```

**Rendering rules**:
- Order and separators are provided by the plugin.
- The host may inject additional items (e.g. `Connections...`) while preserving plugin ordering.
- `label` is used verbatim for display.
- If `path` is supplied, selecting the item navigates to that path.
- If `commandId` is supplied (and no `path`), the host calls `ExecuteMenuCommand`.
- `iconPath` controls icon lookup; if omitted, the host uses `path`.

#### Drives Section (FileSystem Plugin Default)
```text
[Icon] C:\ Local Disk         156 GB
[Icon] D:\ Data               1.20 TB
[Icon] E:\ USB Drive          4.80 GB
──────────────────────────────────────
```

**Drive Menu Format (plugin-defined)**:
- The plugin formats labels; NavigationView renders `label` verbatim.
- FileSystem plugin shows **only free space** (not "X free of Y").
- Format: `{Drive}\ {Label}\t{FreeSpace}` (tab-separated).
- Tab character (`\t`) before size for **right-alignment** in menu.
- Example: `"C:\\ Local Disk\\t156 GB"` displays as `C:\ Local Disk              156 GB`.

**Menu Icon Extraction System (Current Code)**:

✅ **IconCache-mediated icon indices + WIC menu bitmaps**:

NavigationView never calls `SHGetFileInfoW` / `SHGetImageList` directly. It requests system icon indices from `IconCache`, then creates `HBITMAP` menu images via `IconCache`'s WIC pipeline.

**Architecture**:
1. **Plugin-provided icon path or navigation path**: `IconCache::QuerySysIconIndexForPath(path, attributes, useFileAttributes)`
2. **Menu bitmaps**: `IconCache::CreateMenuBitmapFromIconIndex(iconIndex, _menuIconSize)` (DPI-aware extraction + premultiplied BGRA)
3. **Lifetime**: Store bitmaps in `_menuBitmaps` (`std::vector<wil::unique_hbitmap>`) for the life of the popup menu.

```cpp
auto AddMenuItemWithIcon = [&](UINT id, const wchar_t* text, const wchar_t* iconPath)
{
    AppendMenuW(menu, MF_STRING, id, text);

    if (iconPath && iconPath[0] != L'\0')
    {
        const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(iconPath, 0, false);
        if (iconIndex.has_value())
        {
            auto bitmap = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(iconIndex.value(), _menuIconSize);
            if (bitmap)
            {
                SetMenuItemBitmaps(menu, id, MF_BYCOMMAND, bitmap.get(), bitmap.get());
                _menuBitmaps.emplace_back(bitmap.release());
            }
        }
    }
};
```

## Section 2: Path Display with Direct2D/DirectWrite

### Rendering Modes

#### 1. Breadcrumb Mode (Default)
```text
C:  ›  Users  ›  Username  ›  Documents  ›  Projects
└┬┘└┬┘└──┬──┘└┬┘└────┬───┘└┬┘└────┬────┘└┬┘└───┬────┘
 │  │    │    │      │     │      │      │     │
 └──┴────┴────┴──────┴─────┴──────┴──────┴─────┴─── Clickable segments
```

**Features:**
- Each segment is clickable (navigates to that folder)
- **Full clickable zone highlighting**: Hover background extends over entire segment bounds for better UX
- Separator (›) between segments is **clickable** to display sibling folders in a popup menu
  - Clicking separator shows dropdown menu of folders at the same level
  - Example: Clicking › after "Documents" shows siblings of Documents (Pictures, Downloads, etc.)
  - Implemented in `ShowSiblingsDropdown()` method using `CreatePopupMenu()` and shell enumeration
- **Separator visual feedback**:
  - Hover state: `RGB(243, 243, 243)` background
  - Pressed state: `RGB(230, 230, 230)` background (while menu open)
  - State tracked via `_hoveredSeparatorIndex` and `_activeSeparatorIndex`
- **Sibling menu current marker**: Current folder is marked with an accent vertical bar (no icon)
- Hover effect on segments (subtle background)
- Current folder highlighted (bold or accent color)
- **Overflow**: Progressive breadcrumb collapse with clickable `"..."` (see **Breadcrumb Overflow** below)

#### Breadcrumb Overflow

When the Section 2 breadcrumb area is not wide enough to display the full path, `NavigationView` progressively collapses the breadcrumb by hiding middle segments and inserting a single, clickable ellipsis segment `"..."`.

**Design goals**
- Prefer hiding from the center outward (keep start and end segments) before truncating segment text.
- Keep the current folder (last segment) visible as long as possible.
- The ellipsis is an active breadcrumb segment (hover + click) that opens a full-path popup.
- Use literal `"..."` (three dots) for better clickability (avoid using the single-character Unicode ellipsis `U+2026`).

**Progressive collapse examples** (wide -> narrow)
- `aopyum > bqmlskd > cqqqsssd > dazaza > eefghji`
- `aopyum > bqmlskd > ... > dazaza > eefghji`
- `aopyum > bqmlskd > ... > eefghji`
- `aopyum > ... > eefghji`
- `aop... > eefghji`
- `... > eefghji`
- `... > eef...`
- `...`

**Collapse algorithm (layout time)**
1. Measure each segment’s untrimmed text width (DirectWrite `IDWriteTextLayout` metrics) and compute the full breadcrumb width (padding + spacing + separators).
2. If everything fits, render all segments and all separators (no ellipsis).
3. If not:
   - Attempt to render a collapsed form: `[prefix] > "..." > [suffix]` where `prefix` and `suffix` are contiguous blocks from the start/end of the path (always including the last segment), hiding a single middle range.
   - Choose the widest-fitting collapsed form that keeps the most real segments; tie-break by keeping prefix/suffix balanced (center-out feel), then prefer more trailing (`suffix`) context.
4. If even `first > ... > last` does not fit:
   - Drop the first block and show `... > tail` (keep as much trailing/current context as fits).
   - If needed, truncate the last visible segment with a trailing `"..."` to fit.
5. If `... > last` still does not fit, show `...` only (clipping is acceptable if the window is narrower than `"..."`).

**Separator interactivity rules (in overflow mode)**
- Separators between two real path segments remain clickable and open the sibling dropdown (existing behavior).
- Separators adjacent to the ellipsis segment are clickable and open the full-path popup (same behavior as clicking `"..."`), but they do not open the sibling dropdown.

#### Full Path Popup (Ellipsis / Adjacent Separator Click)

Clicking the `"..."` segment opens a lightweight popup window that displays the full path breadcrumb with the same interaction model as the main breadcrumb bar:
- Click a segment to navigate to that folder (closes the popup).
- Click a separator to open the sibling dropdown for the segment to its right.

**Popup sizing and wrapping**
- The popup sizes to content with padding, clamped to the current monitor work area.
- If the full breadcrumb would exceed the monitor width, the popup wraps to multiple lines (breaks only between segments).
- If wrapped content still exceeds monitor height, clamp to the work area height and enable vertical scrolling (no horizontal scrolling).

**Popup dismissal**
- Click outside, `Esc`, or focus loss closes the popup.

**Popup edit mode**
- The popup can enter edit mode (F4 / Ctrl+L / Alt+D, or double-click whitespace) to type/paste a path.
- On `Enter`: if the path is valid and changed, navigate and close the popup; if valid but unchanged, exit edit mode and keep the popup open.

**Proportional Font Sizing** (lines 960-998):

Font sizes are calculated as percentages of the navigation bar height to maintain proper proportions at all DPI settings:

```cpp
void EnsureD2DResources() {
    // Calculate bar height in physical pixels
    float barHeight = static_cast<float>(MulDiv(kHeight, _dpi, USER_DEFAULT_SCREEN_DPI));
    
    // Breadcrumb text: 60% of bar height
    float breadcrumbSize = barHeight * 0.6f;  // ~14.4pt at 96 DPI, 18pt at 120 DPI
    
    // Separator (›): 80% of bar height for visual prominence
    float separatorSize = barHeight * 0.8f;   // ~19.2pt at 96 DPI, 24pt at 120 DPI
    
    _dwriteFactory->CreateTextFormat(
        L"Segoe UI", nullptr,
        DWRITE_FONT_WEIGHT_NORMAL,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        breadcrumbSize,
        L"",
        &_pathFormat);
    
    _pathFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    
    _dwriteFactory->CreateTextFormat(
        L"Segoe UI", nullptr,
        DWRITE_FONT_WEIGHT_NORMAL,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        separatorSize,
        L"",
        &_separatorFormat);
    
    _separatorFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
}
```

**DPI Scaling Examples**:
- **96 DPI (100%)**: Bar = 24px, Breadcrumb = 14.4pt, Separator = 19.2pt
- **120 DPI (125%)**: Bar = 30px, Breadcrumb = 18pt, Separator = 24pt
- **144 DPI (150%)**: Bar = 36px, Breadcrumb = 21.6pt, Separator = 28.8pt

**Vertical Alignment**: Both formats use `DWRITE_PARAGRAPH_ALIGNMENT_CENTER` to ensure text is vertically centered within the navigation bar with automatic margins.

**Implementation:**
```cpp
struct PathSegment {
    std::wstring text;
    D2D1_RECT_F bounds;
    std::filesystem::path fullPath;
};

void RenderBreadcrumbs() {
    if (!_currentPluginPath) return;
    
    _d2dContext->BeginDraw();
    _d2dContext->SetTarget(_d2dTarget.get());
    
    // Clear background
    _d2dContext->Clear(D2D1::ColorF(250/255.0f, 250/255.0f, 250/255.0f));
    
    float x = 8.0f; // Left padding
    float y = (_sectionPathRect.bottom - _sectionPathRect.top) / 2.0f;
    
    auto parts = SplitPathComponents(*_currentPluginPath);
    
    for (size_t i = 0; i < parts.size(); i++) {
        const auto& part = parts[i];
        
        // Measure text
        wil::com_ptr<IDWriteTextLayout> layout;
        _dwriteFactory->CreateTextLayout(
            part.text.c_str(), static_cast<UINT32>(part.text.size()),
            _pathFormat.get(), 1000.0f, 24.0f, &layout);
        
        DWRITE_TEXT_METRICS metrics;
        layout->GetMetrics(&metrics);

        constexpr float textPaddingX = 2.0f;
        constexpr float textPaddingY = 1.0f;
        
        // Store bounds for hit-testing
        _segments[i].bounds = D2D1::RectF(
            x - textPaddingX, y - metrics.height / 2 - textPaddingY,
            x + metrics.width + textPaddingX, y + metrics.height / 2 + textPaddingY);
        
        // Draw hover background
        if (_segments[i].isHovered) {
            D2D1_RECT_F hoverRect = _segments[i].bounds;
            hoverRect.left += 1.0f;
            hoverRect.top += 1.0f;
            hoverRect.right -= 1.0f;
            hoverRect.bottom -= 1.0f;
            _d2dContext->FillRoundedRectangle(D2D1::RoundedRect(hoverRect, 2.0f, 2.0f), _hoverBrush.get());
        }
        
        // Draw text
        _d2dContext->DrawTextLayout(
            D2D1::Point2F(x, y - metrics.height / 2),
            layout.get(),
            i == parts.size() - 1 ? _accentBrush.get() : _textBrush.get());
        
        x += metrics.width + 4.0f;
        
        // Draw separator (›) if not last
        if (i < parts.size() - 1) {
            _d2dContext->DrawText(L"›", 1, _separatorFormat.get(),
                D2D1::RectF(x, y - 8, x + 12, y + 8),
                _separatorBrush.get());
            x += 16.0f;
        }
    }
    
    _d2dContext->EndDraw();
    
    if (_swapChain) {
        DXGI_PRESENT_PARAMETERS params = {};
        _swapChain->Present1(0, 0, &params);
    }
}
```

### Path Segment and Separator Implementation Details

#### Path Parsing: SplitPathComponents()

**Purpose:** Convert filesystem path into clickable breadcrumb segments with complete path information.

**Algorithm:**
1. **Handle root path first** (e.g., `"C:\\"` for Windows drives)
   - Uses `path.root_path()` to get complete root including directory separator
   - Creates root segment with display text `"C:"` but stores full path `"C:\\"`
   - Ensures clicking root segment navigates to drive root correctly
2. **Iterate non-root components** (e.g., `"Users"`, `"Documents"`)
   - Skips root name and root directory (already handled)
   - Accumulates path incrementally: `"C:\\"` → `"C:\\Users"` → `"C:\\Users\\Documents"`
   - Each segment stores its complete accumulated path for navigation

**Implementation Pattern:**
```cpp
std::vector<PathSegment> SplitPathComponents(const std::filesystem::path& path)
{
    std::vector<PathSegment> result;
    std::filesystem::path accumulated;

    // 1. Handle root completely first (e.g., "C:\\" for "C:\\Users\\Documents")
    if (path.has_root_path())
    {
        accumulated = path.root_path();  // Gets "C:\\" complete
        PathSegment root;
        root.text = path.root_name().wstring();  // Display: "C:"
        root.fullPath = accumulated;             // Store: "C:\\"
        root.isHovered = false;
        result.push_back(root);
    }

    // 2. Then iterate remaining components
    for (const auto& part : path)
    {
        if (part == path.root_name() || part == path.root_directory())
            continue;  // Skip root parts (already handled)
        
        accumulated /= part;
        PathSegment segment;
        segment.text = part.wstring();
        segment.fullPath = accumulated;
        segment.isHovered = false;
        result.push_back(segment);
    }

    return result;
}
```

**Example Path Breakdown:**

For `C:\Users\aUser\Downloads`:
```text
Segment 0: text="C:",       fullPath="C:\\"                      (root)
Segment 1: text="Users",    fullPath="C:\\Users"                 (first folder)
Segment 2: text="aUser",    fullPath="C:\\Users\\aUser"          (second folder)
Segment 3: text="Downloads", fullPath="C:\\Users\\aUser\\Downloads" (final folder)
```

**Separator Rendering:**
- Separator 0: Between segment 0 and segment 1 (after `"C:"`)
- Separator 1: Between segment 1 and segment 2 (after `"Users"`)
- Separator 2: Between segment 2 and segment 3 (after `"aUser"`)
- No separator after final segment

#### Separator Click Behavior: ShowSiblingsDropdown()

**Purpose:** Display sibling folders menu when clicking on breadcrumb separators.

**Critical Mapping:** Separator index N is positioned **between** segment[N] and segment[N+1]:
```text
Segment[0] › Segment[1] › Segment[2] › Segment[3]
    "C:"   ↑  "Users"   ↑  "aUser"   ↑  "Downloads"
           |            |            |
      Separator 0  Separator 1  Separator 2
```

**Sibling Folder Logic:**
- Clicking **separator N** shows siblings of **segment[N+1]** (the segment to the right)
- To get siblings, use `segment[N+1].fullPath.parent_path()`
- Example: Clicking separator 0 (after `"C:"`) uses segment[1] (`"C:\\Users"`), parent path is `"C:\\"`, shows folders in root

**Implementation:**
```cpp
void ShowSiblingsDropdown(size_t separatorIndex)
{
    // Separator N is between segment[N] and segment[N+1]
    // Show siblings of segment[N+1] (the segment to the right)
    if (separatorIndex + 1 >= _segments.size())
        return;

    const auto& segment = _segments[separatorIndex + 1];
    std::filesystem::path parentPath = segment.fullPath.parent_path();
    
    // Enumerate folders in parentPath...
}
```

**Example Click Behavior:**

For path `C:\Users\aUser\Downloads`:

| Separator | Position | Uses Segment | Segment Path | Parent Path | Shows Folders In |
|-----------|----------|--------------|--------------|-------------|------------------|
| 0 | After `"C:"` | segment[1] | `"C:\\Users"` | `"C:\\"` | **C:\\ (root)** |
| 1 | After `"Users"` | segment[2] | `"C:\\Users\\aUser"` | `"C:\\Users"` | **C:\\Users** |
| 2 | After `"aUser"` | segment[3] | `"C:\\Users\\aUser\\Downloads"` | `"C:\\Users\\aUser"` | **C:\\Users\\aUser** |

**Menu Display:**
- Lists all folders in parent directory
- Current folder (the one from segment[N+1]) shows with folder icon
- Other folders show without icons (text only)
- Clicking menu item navigates to that folder

**Segment Click Behavior:**
- Clicking segment 0 (`"C:"`) navigates to `"C:\\"` (root path)
- Clicking segment 1 (`"Users"`) navigates to `"C:\\Users"`
- Clicking segment 2 (`"aUser"`) navigates to `"C:\\Users\\aUser"`
- Clicking segment 3 (`"Downloads"`) does nothing (already at this location)

### Interactive States and Visual Feedback

#### Hover Effects (December 2025: Timer-Based Tracking)

**Section 1 (Menu Button)**:
- **Normal**: No background
- **Hover**: `RGB(243, 243, 243)` background
- **Pressed**: `RGB(230, 230, 230)` background (darker when menu is open)
- **Implementation**: Timer-based polling (30 FPS) with `_menuButtonHovered` flag
- **Trigger**: `HOVER_TIMER_ID` timer checks `PtInRect(&_sectionDriveRect, pt)` every 33ms

**Section 2 (Breadcrumb Segments)**:
- **Normal**: Transparent background
- **Hover**: `RGB(243, 243, 243)` subtle background highlight
- **Implementation**: Direct2D `FillRoundedRectangle` on an inset rect of the segment bounds (1 DIP inset, 2 DIP corner radius; see `Specs/VisualStyleSpec.md`)
- **Padding**: Breadcrumb segment text uses symmetric left/right padding so hover backgrounds show equal gaps on both sides
- **Coordinate Transform**: Mouse coordinates transformed from window space to Section 2 local space:
  ```cpp
  float localX = static_cast<float>(pt.x - _sectionPathRect.left);
  float localY = static_cast<float>(pt.y - _sectionPathRect.top);
  D2D1_POINT_2F movePt = D2D1::Point2F(localX, localY);
  ```
- **Hit Testing**: Uses local coordinates against segment bounds stored in local space
- **Hover Cleanup**: Explicitly clears all segment hover states when cursor leaves Section 2
- **Trigger**: `HOVER_TIMER_ID` timer polls cursor position and transforms coordinates

**Section 2 (Breadcrumb Separators)**:
- **Normal**: No background
- **Hover**: `RGB(243, 243, 243)` background (same as segments)
- **Pressed**: `RGB(230, 230, 230)` background (darker, persists while menu is open)
- **Active State**: Remains in pressed state during entire TrackPopupMenu call
- **Implementation**: Hover/pressed backgrounds use the same inset rounded rectangle style as segments (see `Specs/VisualStyleSpec.md`)
- **Coordinate Transform**: Same as segments (local space)
- **Hover Cleanup**: Explicitly clears separator hover states when cursor leaves Section 2
- **Trigger**: `HOVER_TIMER_ID` timer with coordinate transformation

**Disk Info**:
- **Normal/Hover/Pressed**: Uses `NavigationViewTheme` colors (`background`, `backgroundHover`, `backgroundPressed`)
- **Rendering**: Direct2D in `RenderDiskInfoSection()` (disk text + progress bar)
- **Trigger**: `HOVER_TIMER_ID` timer checks `PtInRect(&_sectionDiskInfoRect, pt)` and re-renders the disk section on state changes

**History Button**:
- **Normal**: Transparent background
- **Hover**: `RGB(243, 243, 243)` background
- **Pressed**: `RGB(230, 230, 230)` background
- **Rendering**: Direct2D in `RenderHistorySection()` (glyph button)
- **Trigger**: `HOVER_TIMER_ID` timer checks `PtInRect(&_sectionHistoryRect, pt)` every 33ms

#### Separator Rotation Animation

**Visual Effect**: When separator menu is opened, the separator character (›) rotates 90° clockwise smoothly. When menu closes, it rotates back to 0°.

**Animation Parameters**:
- **Duration**: 150ms (0° → 90° or 90° → 0°)
- **Frame Rate**: ~60 FPS via `Ui::AnimationDispatcher` (single shared 16ms `WM_TIMER`)
- **Rotation Speed**: 600 degrees/second
- **Interpolation**: Linear with per-frame delta calculation

**Implementation**:
```cpp
// Animation state (per separator)
std::vector<float> _separatorRotationAngles;  // Current angle (0-90°)
std::vector<float> _separatorTargetAngles;    // Target angle
uint64_t _separatorAnimationSubscriptionId = 0; // AnimationDispatcher subscription ID
uint64_t _separatorAnimationLastTickMs     = 0; // Last dispatcher tick (ms)

// Constants
static constexpr float ROTATION_SPEED = 600.0f;  // degrees/sec

// Start animation when menu opens
void ShowSiblingsDropdown(size_t separatorIndex) {
    _menuOpenForSeparator = static_cast<int>(separatorIndex);
    StartSeparatorAnimation(separatorIndex, 90.0f);  // Rotate to 90°
    // ... show menu ...
}

// Reverse animation when menu closes (via WM_EXITMENULOOP)
void OnExitMenuLoop(bool isShortcut) {
    if (_menuOpenForSeparator != -1) {
        StartSeparatorAnimation(_menuOpenForSeparator, 0.0f);  // Rotate to 0°
        _menuOpenForSeparator = -1;
        _activeSeparatorIndex = -1;
    }
}

// Animation update (shared dispatcher tick)
bool UpdateSeparatorAnimations(uint64_t nowTickMs) {
    const float dtSeconds = /* (nowTickMs - _separatorAnimationLastTickMs) */ (1.0f / 60.0f);
    _separatorAnimationLastTickMs = nowTickMs;
    const float deltaAngle = ROTATION_SPEED * dtSeconds;
    bool anyAnimating = false;
    
    for (size_t i = 0; i < _separatorRotationAngles.size(); ++i) {
        float& current = _separatorRotationAngles[i];
        float target = _separatorTargetAngles[i];
        
        if (std::abs(current - target) > 0.1f) {
            anyAnimating = true;
            // Interpolate towards target
            current += (current < target) ? deltaAngle : -deltaAngle;
            current = std::clamp(current, 
                                std::min(target, 0.0f), 
                                std::max(target, 90.0f));
        }
    }
    
    RenderSection2();  // Redraw with updated angles
    return anyAnimating;
}

// Rendering with rotation transform
void RenderBreadcrumbs() {
    // ... render segments ...
    
    // Render separator with rotation
    if (rotationAngle > 0.1f) {
        D2D1_POINT_2F center = D2D1::Point2F(
            (separatorRect.left + separatorRect.right) / 2.0f,
            (separatorRect.top + separatorRect.bottom) / 2.0f);
        
        D2D1::Matrix3x2F oldTransform;
        _d2dContext->GetTransform(&oldTransform);
        
        D2D1::Matrix3x2F rotation = 
            D2D1::Matrix3x2F::Rotation(rotationAngle, center);
        _d2dContext->SetTransform(rotation * oldTransform);
        
        _d2dContext->DrawText(L"›", 1, _separatorFormat.get(), 
                            separatorRect, _separatorBrush.get());
        
        _d2dContext->SetTransform(oldTransform);
    }
}
```

**State Management**:
- `_menuOpenForSeparator`: Tracks which separator's menu is currently displayed (-1 if none)
- `_activeSeparatorIndex`: Tracks which separator is in pressed state (-1 if none)
- `_separatorRotationAngles[i]`: Current rotation angle for separator i
- `_separatorTargetAngles[i]`: Target rotation angle for separator i

**Dispatcher Lifecycle**:
1. Subscribe to `Ui::AnimationDispatcher` when first animation starts
2. Dispatcher ticks at 16ms and calls the per-component tick callback
3. Unsubscribe when all animations reach their targets
4. No per-component timers; dispatcher stops when no subscribers remain

#### Menu Switching Behavior

**User Interaction**: When a separator sibling-menu is open and the user targets a different separator (click or hover), the current menu closes and the new menu opens.

**Implementation Approach**: Simplified modal menu handling using `WM_CANCELMODE`.

**Why WM_CANCELMODE?**:
- `TrackPopupMenu` is modal and blocks message processing
- Cannot check menu state or send `WM_COMMAND` while modal loop is running
- `WM_CANCELMODE` is the only message that forces `TrackPopupMenu` to exit immediately
- Cleaner than complex `SetWindowsHookEx` approach with global state

**Flow**:
1. Separator A menu is open (`_menuOpenForSeparator == A`)
2. User targets separator B
3. If B is different and eligible:
   - Highlight B (hover state) immediately
   - Send `WM_CANCELMODE` to force the current menu to close
   - Post a deferred “open separator menu B” message (e.g. `WM_USER+100`)
4. `TrackPopupMenu` for separator A exits, `WM_EXITMENULOOP` fires
5. `OnExitMenuLoop` clears state and starts reverse rotation animation
6. Deferred message opens B’s menu; pressed state/animation updates apply to B

**Code**:
```cpp
void OnLButtonDown(POINT pt) {
    // ... coordinate transformation ...
    
    for (size_t i = 0; i < _separatorBounds.size(); i++) {
        if (/* click in separator bounds */) {
            // Check if different separator menu is open
            if (_menuOpenForSeparator != -1 && 
                _menuOpenForSeparator != static_cast<int>(i)) {
                // Close current menu and defer new menu
                SendMessageW(_hWnd, WM_CANCELMODE, 0, 0);
                PostMessageW(_hWnd, WM_USER + 100, 
                           static_cast<WPARAM>(i), 0);
            } else {
                // No menu open or same separator clicked
                ShowSiblingsDropdown(i);
            }
            return;
        }
    }
}

// WndProc handler for deferred menu opening
case WM_USER + 100:
    ShowSiblingsDropdown(static_cast<size_t>(wp));
    return 0;
```

**Hover Switching (During `TrackPopupMenu`)**:
- `TrackPopupMenu` redirects mouse messages to the menu window, so `OnMouseMove`/`OnLButtonDown` on the owning control won’t run.
- To support “hover-to-switch” behavior, a 30 FPS timer polls `GetCursorPos()` + hit testing against separator bounds.
- If the cursor is over a different eligible separator, the same `WM_CANCELMODE` + deferred-open flow is used.
- While the cursor is over a Win32 menu window (`#32768`), polling ignores hover/switching to avoid accidental menu changes.

**Full-Path Popup Consistency**:
- The full-path popup breadcrumb window uses the same hover-to-switch logic for its separator sibling menus.
- While the full-path popup is open, `NavigationView` hover tracking is suspended to prevent background hover/switch activity beneath the popup.

**Why PostMessage Instead of Direct Call?**:
- `SendMessageW(WM_CANCELMODE)` triggers menu exit but message processing continues
- Modal loop completes, `WM_EXITMENULOOP` is dispatched
- Only after modal loop ends can we safely open new menu
- `PostMessage` defers new menu opening until after cleanup is complete

**Alternative Rejected Approaches**:
- **Approach A**: `SetWindowsHookEx` with global hook - Complex, requires unhook, global state
- **Approach B**: Custom message loop with `PeekMessage` - Breaks Win32 modal behavior
- **Approach C**: `EndMenu()` API - Not reliable with owner-draw menus, deprecated

#### 2. Path Tooltip (On Hover)
```text
C:\Users\Username\Documents\Projects\RedSalamander
```

**Features:**
- Shows complete path on hover via tooltip (no inline full-path render mode)
- Useful for seeing full network paths or long paths
- Uses the existing tooltip window (`TOOLTIPS_CLASS` + `TTF_TRACK`) and updates text based on hover target


#### 3. Edit Mode
```text
┌──────────────────────────────────────────────────┐
│ C:\Users\Username\Documents█                     │
└──────────────────────────────────────────────────┘
```

**Features:**
- F4 / Ctrl+L / Alt+D to enter edit mode
- Full text selection initially
- Caret blinking at insertion point
- Native keyboard handling (Ctrl+A, Ctrl+C, Ctrl+V)
- Edit underline: `2 DIP` line at the bottom of the edit field (see `Specs/VisualStyleSpec.md`)
- Close button (`X`) at the right side to cancel/exit edit mode (see `Specs/VisualStyleSpec.md`)
- Autosuggest popup under the edit while typing:
  - Shows up to `10` items.
    - If there are more than `10` matching folders, the last item is a disabled `...` indicator (the popup shows `9` actual folders).
  - Uses `DirectoryInfoCache` snapshots when present (cache-first); otherwise enumerates via the `IFileSystem` resolved from the edit text on a worker thread (no native enumeration fallback)
    - Default: uses the current pane plugin.
    - If the edit text contains a `<shortId>:` prefix, autosuggest uses that plugin (even if the current pane is using another plugin).
    - If the edit text contains a mount context (`<instanceContext>|`) for a non-`file` plugin, autosuggest enumerates using an instance initialized for that mount (reuses the active pane instance if it matches, otherwise creates a temporary instance).
    - Special case: typing `7z:<zipPath>` (no `|` and `<zipPath>` does not start with `/` or `\`) enumerates via the `file` plugin to help pick an archive path to mount (the edit text still stays `7z:`-prefixed).
  - Provides cross-filesystem static suggestions immediately (before async enumeration completes):
    - plugin schemes / short IDs (ex: `ftp:`, `sftp:`),
    - drive roots (ex: `C:\`),
    - Connection Manager routing (`nav:`, `nav://`, `@conn:`) + connection-name suggestions (with resolved preview like `sftp://user@host:22`).
  - Filtering: matches folders whose names **contain** the typed text (case-insensitive), not just prefix matches
  - The typed substring is highlighted inside each suggestion item to show why it matched.
  - Selecting a suggestion (mouse click or `Enter` when a suggestion is selected) updates the edit text to the chosen folder and appends the suggestion’s directory separator (typically `\` for Windows paths and `/` for plugin paths), and **stays in edit mode** so the next level can be suggested immediately.
  - Popup UI (modernized to match Settings combobox dropdown):
    - rounded border, row height aligned to the navigation bar height (minimum `40 DIP`),
    - hover updates the active selection, and the selected row shows a left accent bar + rounded highlight.

**Trigger**: F4 / Ctrl+L / Alt+D, Enter/Space when Path region is focused, or double-click whitespace in the path area

**Edit chrome (visual behavior)**:
- While `_editMode == true`, `NavigationView` still renders the Path section to draw the underline + close button.
- The Win32 Edit control is laid out to **exclude** the close button region and the underline strip (so the D2D-drawn chrome is visible).

**Implementation**:
```cpp
void EnterEditMode() {
    if (_editMode) return;
    _editMode = true;
    
    // Hide Direct2D breadcrumb rendering
    _renderMode = RenderMode::Edit;
    
    // Create or show Edit control overlay
    const std::filesystem::path& currentPath = _currentEditPath.has_value() ? _currentEditPath.value() : _currentPath.value();
    if (!_pathEdit) {
        int x = static_cast<int>(_sectionPathRect.left);
        int y = static_cast<int>(_sectionPathRect.top);
        int width = static_cast<int>(_sectionPathRect.right - _sectionPathRect.left);
        int height = static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top);

        // Layout leaves space for the edit underline + close button (see VisualStyleSpec).
        RECT editRect = _sectionPathRect;
        
        _pathEdit = CreateWindowExW(
            0, L"EDIT", currentPath.c_str(),
            WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOHSCROLL | ES_LEFT,
            x, y, width, height,
            _hWnd, (HMENU)ID_PATH_EDIT, _hInstance, nullptr);

        SendMessage(_pathEdit, WM_SETFONT, (WPARAM)_pathFont, TRUE);
        SubclassEditControl();

        // Use a multiline EDIT so the control can fill the full bar height, then shift the
        // formatting rect (EM_SETRECTNP) so the single line of text is vertically centered.
        LayoutSingleLineEditInRect(_pathEdit, editRect);
    } else {
        SetWindowText(_pathEdit, currentPath.c_str());
        ShowWindow(_pathEdit, SW_SHOW);
        LayoutSingleLineEditInRect(_pathEdit, editRect);
    }
    
    // Select all text
    SendMessage(_pathEdit, EM_SETSEL, 0, -1);
    SetFocus(_pathEdit);
}

void ExitEditMode(bool accept) {
    if (!_editMode) return;
    _editMode = false;
    
    if (accept) {
        wchar_t buffer[MAX_PATH];
        GetWindowText(_pathEdit, buffer, MAX_PATH);
        
        if (ValidatePath(buffer)) {
            RequestPathChange(std::filesystem::path(buffer));
        } else {
            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, buffer);
            const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_INVALID_PATH);

            EDITBALLOONTIP tip{};
            tip.cbStruct = sizeof(tip);
            tip.pszTitle = title.c_str();
            tip.pszText = message.c_str();
            tip.ttiIcon = TTI_WARNING;
            SendMessage(_pathEdit, EM_SHOWBALLOONTIP, 0, reinterpret_cast<LPARAM>(&tip));
            return; // Keep edit mode active
        }
    }
    
    ShowWindow(_pathEdit, SW_HIDE);
    _renderMode = RenderMode::Breadcrumb;
    InvalidateRect(_hWnd, nullptr, FALSE);
}
```

**Path change routing (host integration):**
- User-initiated navigation calls `RequestPathChange(path)`.
- If `PathChangedCallback` is set (FolderWindow integration), NavigationView calls the callback and does **not** call `SetPath()` directly (FolderWindow remains the single source of truth and updates both NavigationView + FolderView).
- If no callback is set, `RequestPathChange` falls back to `SetPath()` for standalone usage.

### History Dropdown

**Location**: Small button (24×24 DIP) in `_sectionHistoryRect` between the path area and the disk info area  
**Icon**: Down glyph (currently rendered as `⩔`)  
**Trigger**: Click history button region, or press **Alt+Down**

**Storage**: History is supplied by the host (FolderWindow) via `SetHistory()` and persisted in settings (`folders.history`, bounded by `folders.historyMax`, default `20`, clamped `1..50`).

**Current entry marker (current behavior)**:
- The current path is marked via `MF_CHECKED` using `wil::compare_string_ordinal(..., /*ignoreCase*/ true)` and rendered as an accent vertical bar (same style as autosuggest).

**Sizing and truncation**:
- The history popup menu width MUST NOT exceed the main window client width.
- If a history entry does not fit, the displayed label MUST use a **middle ellipsis**.

```cpp
std::deque<std::filesystem::path> _pathHistory; // most recent first

void ShowHistoryDropdown() {
    if (_pathHistory.empty()) return;
    
    HMENU menu = CreatePopupMenu();
    
    constexpr UINT kCmdHistoryBase = 1u;
    UINT id = kCmdHistoryBase;
    for (const auto& path : _pathHistory) {
        UINT flags = MF_STRING;
        if (_currentPath && wil::compare_string_ordinal(path.wstring(), _currentPath.value().wstring(), true) == wistd::weak_ordering::equivalent) {
            flags |= MF_CHECKED;
        }
        AppendMenuW(menu, flags, id++, path.c_str());
    }
    
    PrepareThemedMenu(menu);

    RECT rc = _sectionHistoryRect;
    POINT pt = {rc.right, rc.bottom};
    ClientToScreen(_hWnd.get(), &pt);
    
    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_RIGHTALIGN | TPM_TOPALIGN, pt, _hWnd.get());

    if (selectedId >= static_cast<int>(kCmdHistoryBase)) {
        const size_t index = static_cast<size_t>(selectedId - static_cast<int>(kCmdHistoryBase));
        if (index < _pathHistory.size()) {
            RequestPathChange(_pathHistory[index]);
        }
    }
}
```

### Hit-Testing for Breadcrumbs
```cpp
void OnLButtonDown(POINT pt) {
    if (_editMode) return;
    
    // Section 1: menu button
    if (PtInRect(&_sectionDriveRect, pt)) {
        _focusedRegion = FocusRegion::Menu;
        ShowMenuDropdown();
        return;
    }

    // Section 3: history dropdown
    if (PtInRect(&_sectionHistoryRect, pt)) {
        ShowHistoryDropdown();
        return;
    }

    // Section 4: disk info dropdown
    if (PtInRect(&_sectionDiskInfoRect, pt)) {
        ShowDiskInfoDropdown();
        return;
    }

    // Section 2: breadcrumb area
    if (!PtInRect(&_sectionPathRect, pt)) return;
    
    const float localX = static_cast<float>(pt.x - _sectionPathRect.left);
    const float localY = static_cast<float>(pt.y - _sectionPathRect.top);
    const D2D1_POINT_2F clickPt = D2D1::Point2F(localX, localY);

    // Segment hit testing (bounds are stored in Section 2 local coordinates)
    for (const auto& segment : _segments) {
        if (segment.bounds.left <= clickPt.x && clickPt.x <= segment.bounds.right &&
            segment.bounds.top <= clickPt.y && clickPt.y <= segment.bounds.bottom) {
            RequestPathChange(segment.fullPath);
            return;
        }
    }
    
    // Separator hit testing can open the sibling dropdown
    for (size_t i = 0; i < _separatorBounds.size(); ++i) {
        const auto& bounds = _separatorBounds[i];
        if (bounds.left <= clickPt.x && clickPt.x <= bounds.right &&
            bounds.top <= clickPt.y && clickPt.y <= bounds.bottom) {
            ShowSiblingsDropdown(i);
            return;
        }
    }
}

void OnMouseMove(POINT pt) {
    if (_editMode) return;

    // Hover is tracked with a timer (HOVER_TIMER_ID); mouse move can also trigger an immediate Section 2 redraw.
    // Convert to Section 2 local coords and update _hoveredSegmentIndex/_hoveredSeparatorIndex, then call RenderPathSection() if changed.
}
```

## Edit Mode

**Purpose:** Allow users to directly type or paste paths instead of clicking breadcrumbs.

**Activation:**
- Click on path display area (Section 2)
- Press F4 key
- Click dropdown arrow and select "Edit Path"

**Implementation:**
- Creates Win32 `EDIT` control overlaying Section 2
- Pre-populated with current path as string
- Standard text editing: select all, copy/paste, arrows

**Exit Conditions:**
- **Enter key**: Apply new path (navigate if valid)
- **Escape key**: Cancel and return to breadcrumb mode
- **Focus loss**: Cancel and return to breadcrumb mode

**Methods:**
```cpp
void EnterEditMode();  // Show edit control, hide breadcrumbs  
void ExitEditMode();   // Hide edit control, show breadcrumbs
```

**Validation:**
- Check if path exists before navigating
- Show error message for invalid paths
- Expand environment variables (%USERPROFILE%, etc.)

## History Navigation

**Purpose:** Quick access to recently visited paths.

**UI Elements:**
- **History dropdown button** (`_sectionHistoryRect`): Opens a popup menu of recent paths (newest-first, bounded by `folders.historyMax`, default `20`, clamped `1..50`).

**Keyboard Shortcut:**
- **Alt+Down**: Open history dropdown menu

## Disk Info Section

### Visual Design
- **Type**: Direct2D-rendered region (`_sectionDiskInfoRect`) with manual hit testing (no HWND)
- **Source**: `IDriveInfo` on the active plugin; section hidden when unsupported. `GetDriveInfo` failures yield empty text + grey bar.
- **Text Format**: `FormatBytesCompact(_freeBytes)` when free bytes are provided; otherwise `FormatBytesCompact(_totalBytes)` when only total bytes are provided (e.g., archive containers)
- **Rendering**: Direct2D text + 3px progress bar; dropdown menu is a Win32 popup menu themed via `MenuTheme`
- **Cursor**: `IDC_ARROW` (current implementation uses arrow everywhere)
- **Hover**: Timer-based tracking with `_diskInfoHovered` flag driving `RenderDiskInfoSection()`

### Progress Bar (December 2025)
✅ **Discrete 3-pixel progress bar** at bottom of Section 3:
- **Height**: 3 pixels
- **Margins**: 4 pixels left/right from button edges
- **Colors** (theme-driven):
  - `NavigationViewTheme::progressOk` when disk usage < 90%
  - `NavigationViewTheme::progressWarn` when disk usage ≥ 90% (warning state)
  - `NavigationViewTheme::progressBackground` when disk space info unavailable (`_totalBytes == 0`)
- **Position**: Bottom of the disk info region (`_sectionDiskInfoRect`)
- **Calculation**: `int progressWidth = static_cast<int>((progressRect.right - progressRect.left) * usedPercent / 100.0);`
- **Used bytes**: Prefer `usedBytes` from `IDriveInfo`; otherwise use `total - free`.

**Conditional Display** (December 2025):
- Progress bar shows **colored bar** only when `_totalBytes > 0` and either `freeBytes` or `usedBytes` are available (valid disk usage data)
- When `_totalBytes == 0` (network paths, virtual folders, or plugin lacks size data), shows **grey bar only**
- Text displays formatted space when available, empty string otherwise

**Implementation (current)**:
- Rendered via `RenderDiskInfoSection()` using Direct2D (no `ID_DISK_STATIC` owner-draw control).
- Background uses `NavigationViewTheme` hover/pressed colors; progress bar uses `NavigationViewTheme::progressOk` / `progressWarn` / `progressBackground`.
- Click is handled in `OnLButtonDown()` by checking `PtInRect(&_sectionDiskInfoRect, pt)` and calling `ShowDiskInfoDropdown()`.

### Size Formatting
- Compact sizes use `FormatBytesCompact(bytes)` from `Common/Helpers.h` (base-1024 units, locale-aware formatting via `{:L}` / `{:Lf}`).

### Disk Information Update
```cpp
void UpdateDiskInfo() {
    _diskSpaceText.clear();
    _freeBytes = 0;
    _totalBytes = 0;
    _usedBytes = 0;
    _hasTotalBytes = false;
    _hasFreeBytes = false;
    _hasUsedBytes = false;
    _volumeLabel.clear();
    _fileSystem.clear();
    _driveDisplayName.clear();

    if (!_currentPluginPath || !_driveInfo) return;

    const std::wstring pathText = _currentPluginPath.value().wstring();
    DriveInfo info{};
    const HRESULT hr = _driveInfo->GetDriveInfo(pathText.c_str(), &info);
    if (FAILED(hr) || hr == S_FALSE) return;

    if ((info.flags & DRIVE_INFO_FLAG_HAS_DISPLAY_NAME) != 0 && info.displayName)
        _driveDisplayName = info.displayName;
    else
        _driveDisplayName = _currentPluginPath.value().root_path().wstring();

    if ((info.flags & DRIVE_INFO_FLAG_HAS_VOLUME_LABEL) != 0 && info.volumeLabel)
        _volumeLabel = info.volumeLabel;

    if ((info.flags & DRIVE_INFO_FLAG_HAS_FILE_SYSTEM) != 0 && info.fileSystem)
        _fileSystem = info.fileSystem;

    if ((info.flags & DRIVE_INFO_FLAG_HAS_TOTAL_BYTES) != 0)
    {
        _totalBytes = info.totalBytes;
        _hasTotalBytes = true;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_FREE_BYTES) != 0)
    {
        _freeBytes = info.freeBytes;
        _hasFreeBytes = true;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_USED_BYTES) != 0) {
        _usedBytes = info.usedBytes;
        _hasUsedBytes = true;
    }

    if (_hasTotalBytes)
        _diskSpaceText = FormatBytesCompact(_hasFreeBytes ? _freeBytes : _totalBytes) + L" ";
}
```

**Conditional Display Logic** (December 2025):
- **No `IDriveInfo`**: disk info section hidden.
- **`GetDriveInfo` fails or returns `S_FALSE`**: shows empty text + grey progress bar.
- **No `totalBytes`**: shows empty text + grey progress bar.
- **`totalBytes` only (no `freeBytes`/`usedBytes`)**: shows formatted size text + grey progress bar (avoid misleading “100% used” bars for containers like archives).
- **`totalBytes` + (`freeBytes` or `usedBytes`)**: shows formatted space text + colored progress bar.
- **Rationale**: Avoids displaying misleading data when disk space is not meaningful or unavailable.

### Detailed Info Dropdown
```cpp
void ShowDiskInfoDropdown() {
    if (!_currentPluginPath || !_driveInfo) return;

    UpdateDiskInfo();

    uint64_t usedBytes = 0;
    bool hasUsedBytes = false;
    if (_hasUsedBytes) { usedBytes = _usedBytes; hasUsedBytes = true; }
    else if (_hasTotalBytes && _hasFreeBytes && _totalBytes >= _freeBytes) { usedBytes = _totalBytes - _freeBytes; hasUsedBytes = true; }

    double usedPercent = (_hasTotalBytes && _totalBytes > 0 && hasUsedBytes) ?
        (static_cast<double>(usedBytes) * 100.0 / static_cast<double>(_totalBytes)) : 0.0;

    HMENU menu = CreatePopupMenu();

    const std::wstring headerName = _driveDisplayName.empty() ? _currentPluginPath.value().wstring() : _driveDisplayName;
    AppendMenuW(menu, MF_STRING, 0, FormatStringResource(nullptr, IDS_FMT_DISK_INFO_HEADER, headerName).c_str());
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

    // Only show lines that are actually provided by the active plugin.
    if (!_volumeLabel.empty())
        AppendMenuW(menu, MF_STRING, 0, FormatStringResource(nullptr, IDS_FMT_DISK_VOLUME_LABEL, _volumeLabel).c_str());
    if (!_fileSystem.empty())
        AppendMenuW(menu, MF_STRING, 0, FormatStringResource(nullptr, IDS_FMT_DISK_FILE_SYSTEM, _fileSystem).c_str());
    if (!_volumeLabel.empty() || !_fileSystem.empty())
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

    if (_hasTotalBytes)
        AppendMenuW(menu, MF_STRING, 0,
                    FormatStringResource(nullptr, IDS_FMT_DISK_TOTAL_SPACE, FormatBytesCompact(_totalBytes), _totalBytes).c_str());
    if (hasUsedBytes)
        AppendMenuW(menu, MF_STRING, 0,
                    FormatStringResource(nullptr, IDS_FMT_DISK_USED_SPACE, FormatBytesCompact(usedBytes), usedBytes).c_str());
    if (_hasFreeBytes)
        AppendMenuW(menu, MF_STRING, 0,
                    FormatStringResource(nullptr, IDS_FMT_DISK_FREE_SPACE, FormatBytesCompact(_freeBytes), _freeBytes).c_str());

    if (_hasTotalBytes && _totalBytes > 0 && hasUsedBytes)
    {
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, 0,
                    FormatStringResource(nullptr, IDS_FMT_DISK_USED_PERCENT, usedPercent).c_str());
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }

    const std::wstring pathText = _currentPluginPath.value().wstring();
    const NavigationMenuItem* items = nullptr;
    unsigned int count = 0;
    if (SUCCEEDED(_driveInfo->GetDriveMenuItems(pathText.c_str(), &items, &count)) && items && count > 0)
    {
        // Append plugin-provided commands using the same rendering rules as INavigationMenu.
        // NavigationView maps each actionable item to ID_DRIVE_MENU_BASE + index.
    }

    RECT rc  = _sectionDiskInfoRect;
    POINT pt = {rc.right, rc.bottom};
    ClientToScreen(_hWnd.get(), &pt);

    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_RIGHTALIGN | TPM_TOPALIGN, pt, _hWnd.get());
    if (selectedId != 0)
    {
        // Handle selected drive-menu item (navigate path or ExecuteDriveMenuCommand).
    }
}
```

## Class Interface

This section is illustrative; refer to `RedSalamander/NavigationView.h` for authoritative signatures/members.

```cpp
class NavigationView
{
public:
    NavigationView();
    ~NavigationView();
    
    // Non-copyable
    NavigationView(const NavigationView&) = delete;
    NavigationView& operator=(const NavigationView&) = delete;
    
    // Window lifecycle
    static ATOM RegisterWndClass(HINSTANCE instance);
    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy();
    HWND GetHwnd() const { return _hWnd; }
    
    // Path management
    void SetPath(const std::optional<std::filesystem::path>& path);
    std::optional<std::filesystem::path> GetPath() const { return _currentPath; }
    void AddToHistory(const std::filesystem::path& path);
    
    // Callbacks
    using PathChangedCallback = std::function<void(const std::optional<std::filesystem::path>&)>;
    void SetPathChangedCallback(PathChangedCallback callback);
    
private:
    static constexpr wchar_t kClassName[] = L"RedSalamander.NavigationView";
    static constexpr int kHeight = 24;           // DIP at 96 DPI
    static constexpr int kSection1Width = 28;    // Menu button (reduced from 36 for tighter layout)
    static constexpr int kSection3Width = 120;   // Disk info
    static constexpr int kHistoryButtonWidth = 24; // History dropdown
    
    enum class RenderMode {
        Breadcrumb,  // Default: clickable path segments
        Edit         // Edit mode: Win32 Edit control
    };
    
    struct PathSegment {
        std::wstring text;
        D2D1_RECT_F bounds;
        std::filesystem::path fullPath;
        bool isHovered = false;
    };
    
    static LRESULT CALLBACK WndProcThunk(HWND, UINT, WPARAM, LPARAM);
    LRESULT WndProc(HWND, UINT, WPARAM, LPARAM);
    
    // Message handlers
    void OnCreate(HWND hwnd);
    void OnDestroy();
    void OnPaint();
    void OnSize(UINT width, UINT height);
    void OnCommand(UINT id, HWND hwndCtl, UINT codeNotify);
    void OnDrawItem(DRAWITEMSTRUCT* dis);
    void OnLButtonDown(POINT pt);
    void OnMouseMove(POINT pt);
    void OnMouseLeave();
    void OnSetCursor(HWND hwnd, UINT hitTest, UINT mouseMsg);
    void OnDpiChanged(UINT dpiX, UINT dpiY, RECT* rect);
    void OnTimer(UINT_PTR timerId);  // Animation + hover tracking
    void OnExitMenuLoop(bool isShortcut);  // Menu cleanup + animation
    
    // Layout
    void CalculateLayout();
    RECT _section1Rect = {};      // Menu button (GDI)
    RECT _section2Rect = {};      // Path display (Direct2D)
    RECT _section3Rect = {};      // Disk info (GDI)
    RECT _historyButtonRect = {}; // History dropdown button
    
    // Child controls (Win32)
    HWND _menuButton = nullptr;   // Section 1: Menu button (GDI)
    HWND _pathEdit = nullptr;     // Section 2: Edit control (edit mode only)
    HWND _historyButton = nullptr;// Section 2: History dropdown button
    HWND _diskStatic = nullptr;   // Section 3: Disk space button (owner-draw)
    
    // Direct2D rendering for Section 2
	    void EnsureD2DResources();
	    void DiscardD2DResources();
	    void RenderSection2();
	    void RenderBreadcrumbs();
	    std::vector<PathSegment> SplitPathComponents(const std::filesystem::path& path);
    
    // Menus
    void ShowMenuDropdown();
    void ShowHistoryDropdown();
    void ShowDiskInfoDropdown();
    void ShowSiblingsDropdown(size_t segmentIndex);
    
    // Animation
    void StartSeparatorAnimation(size_t index, float targetAngle);
    void UpdateSeparatorAnimations();
    
    // Path editing
    void EnterEditMode();
    void ExitEditMode(bool accept);
    bool ValidatePath(const std::wstring& pathStr);
    void SubclassEditControl();
    static LRESULT CALLBACK EditSubclassProc(HWND, UINT, WPARAM, LPARAM, 
                                             UINT_PTR, DWORD_PTR);
    
    // Disk info
    void UpdateDiskInfo();
    std::wstring FormatBytes(uint64_t bytes);
    
    // State
    HWND _hWnd = nullptr;
    HINSTANCE _hInstance = nullptr;
    UINT _dpi = USER_DEFAULT_SCREEN_DPI;
    SIZE _clientSize = {0, 0};
    RenderMode _renderMode = RenderMode::Breadcrumb;
    bool _editMode = false;
    
    // Path data (see Path Model section)
    std::optional<std::filesystem::path> _currentPath;       // canonical location (history/persistence)
    std::optional<std::filesystem::path> _currentPluginPath; // pluginPath used for breadcrumb + plugin calls
    std::optional<std::filesystem::path> _currentEditPath;   // edit-mode display path
    std::wstring _currentInstanceContext;
    std::vector<PathSegment> _segments;
    std::deque<std::filesystem::path> _pathHistory;  // most recent first (bounded by host/settings)
    PathChangedCallback _pathChangedCallback;
    
    // Disk data
    std::wstring _diskSpaceText;
    uint64_t _freeBytes = 0;
    uint64_t _totalBytes = 0;  // 0 = no disk info available (grey bar)
    uint64_t _usedBytes = 0;
    bool _hasUsedBytes = false;
    std::wstring _volumeLabel;
    std::wstring _fileSystem;
    std::wstring _driveDisplayName;
    
    // Interactive state (December 2025)
    bool _menuButtonHovered = false;     // Section 1 hover state (timer-based)
    bool _menuButtonPressed = false;     // Section 1 pressed state
    bool _diskInfoHovered = false;       // Section 3 hover state (timer-based)
    bool _historyButtonHovered = false;  // History button hover (timer-based)
    bool _isMenuOpen = false;            // Section 1 menu open state
    int _menuOpenForSeparator = -1;      // Which separator's menu is open
    int _activeSeparatorIndex = -1;      // Which separator is pressed
    
    // Hover polling (enabled only for menus/edit mode)
    static constexpr UINT_PTR HOVER_TIMER_ID = 2;
    static constexpr UINT HOVER_CHECK_FPS = 30;
    UINT_PTR _hoverTimer = 0;
    
    // Animation system (shared dispatcher)
    static constexpr float ROTATION_SPEED = 600.0f;  // degrees/second
    uint64_t _separatorAnimationSubscriptionId = 0;
    uint64_t _separatorAnimationLastTickMs = 0;
    std::vector<float> _separatorRotationAngles;  // Current angle (0-90°)
    std::vector<float> _separatorTargetAngles;    // Target angle
    
    // GDI resources (Sections 1 & 3)
    HFONT _pathFont = nullptr;      // For Edit control
    HFONT _diskFont = nullptr;      // For Section 3
    HBRUSH _backgroundBrush = nullptr;
    HPEN _borderPen = nullptr;
    
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
    
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;      // RGB(32,32,32)
    wil::com_ptr<ID2D1SolidColorBrush> _separatorBrush; // RGB(120,120,120)
    wil::com_ptr<ID2D1SolidColorBrush> _hoverBrush;     // RGB(243,243,243)
    wil::com_ptr<ID2D1SolidColorBrush> _accentBrush;    // System accent color
    
    // Command IDs
    enum {
        ID_MENU_BUTTON = 100,
        ID_PATH_EDIT,
        ID_HISTORY_BUTTON,
        ID_DISK_STATIC,

        ID_NAV_MENU_BASE   = 200,
        ID_NAV_MENU_MAX    = 399,

        ID_DRIVE_MENU_BASE = 500,
        ID_DRIVE_MENU_MAX  = 599,

        ID_SIBLING_BASE    = 600, // 600-699 for sibling folders
    };
};
```

## Integration with FolderView

### Main Window Layout
```text
┌─────────────────────────────────┐
│   NavigationView (24 DIP)       │ ← Fixed height at top
├─────────────────────────────────┤
│                                 │
│        FolderView               │ ← Direct2D rendering
│      (remaining height)         │
│                                 │
└─────────────────────────────────┘
```

### RedSalamander.cpp Integration

```cpp
class MainWindow {
private:
    NavigationView _navView;
    FolderView _folderView;
    
    void OnCreate() {
        RECT rc;
        GetClientRect(_hWnd, &rc);
        
        UINT dpi = GetDpiForWindow(_hWnd);
        int navHeight = MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI);
        
        // Create navigation view at top
        _navView.Create(_hWnd, 0, 0, rc.right, navHeight);
        _navView.SetPathChangedCallback([this](const auto& path) {
            _folderView.SetFolderPath(path);
        });
        
        // Create folder view below navigation
        _folderView.Create(_hWnd, 0, navHeight, rc.right, rc.bottom - navHeight);
        
        // Sync initial path
        auto initialPath = _folderView.GetFolderPath();
        if (initialPath) {
            _navView.SetPath(*initialPath);
        }
    }
    
    void OnSize(UINT width, UINT height) {
        UINT dpi = GetDpiForWindow(_hWnd);
        int navHeight = MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI);
        
        MoveWindow(_navView.GetHwnd(), 0, 0, width, navHeight, TRUE);
        MoveWindow(_folderView.GetHwnd(), 0, navHeight, 
                   width, height - navHeight, TRUE);
    }
};
```

### Bi-directional Synchronization
```cpp
// FolderView → NavigationView
void FolderView::SetFolderPath(const std::filesystem::path& path) {
    _currentFolder = path;
    EnumerateFolder();
    
    // Notify parent window
    SendMessage(GetParent(_hWnd), WM_APP_FOLDERCHANGED, 
                0, (LPARAM)&path);
}

// MainWindow message handler
LRESULT OnAppFolderChanged(WPARAM, LPARAM lp) {
    auto path = reinterpret_cast<const std::filesystem::path*>(lp);
    _navView.SetPath(*path);
    _navView.AddToHistory(*path);
    return 0;
}

// NavigationView → FolderView (via callback)
_navView.SetPathChangedCallback([this](const auto& path) {
    _folderView.SetFolderPath(path);
});
```

## Keyboard Shortcuts

- Canonical shortcut map and cross-pane routing rules are defined in `Specs/CommandMenuKeyboardSpec.md`.
- **F4**: Enter edit mode in Section 2, select all text
- **Alt+D**: Same as F4 (Windows Explorer standard) *(default chord binding is settings-backed)*
- **Ctrl+L**: Same as F4 (browser standard) *(default chord binding is settings-backed)*
- **Alt+Down**: (Optional) Open history dropdown (equivalent to activating **History** region) *(default chord binding is settings-backed)*
- **Enter** (in edit): Accept path, exit edit mode
- **Escape** (in edit): Cancel changes, revert to previous path
- **Alt+Up**: Navigate to parent folder (handled by FolderView)
- **Tab**: Cycle focus through **visible** regions (Menu → Path → History → Disk Info, skipping hidden sections); when on **Disk Info**, Tab moves focus to **FolderView**
- **Shift+Tab**: Reverse focus through **visible** regions (Disk Info → History → Path → Menu, skipping hidden sections); when on **Menu**, Shift+Tab moves focus to **FolderView**

**Routing Requirement:**
- `F4` / `Alt+D` / `Ctrl+L` and `Alt+Down` MUST work even when **FolderView** has focus (settings-backed shortcuts are routed at the host level).
- `Tab` / `Shift+Tab` inside FolderView are reserved for pane switching (see `Specs/CommandMenuKeyboardSpec.md`) and are not forwarded to NavigationView.

## Focus Traversal and Activation

NavigationView must implement internal focus navigation because most interactive regions are Direct2D-rendered (no HWND).

**Requirements:**
- Maintain a focused region state: **Menu**, **Path**, **History**, **Disk Info** (skip hidden sections)
- Draw a visible focus indicator for the focused region when NavigationView has focus and is not in edit mode
- **Tab / Shift+Tab** moves focus between regions (see shortcut list above)
- **Cross-control focus handoff**:
  - When focused region is **Disk Info** and user presses **Tab**, move focus to **FolderView**
  - When focused region is **Menu** and user presses **Shift+Tab**, move focus to **FolderView**
- **Enter / Space** activates the focused region:
  - **Menu** → open menu dropdown
  - **Path** → enter edit mode (select all)
  - **History** → open history dropdown
  - **Disk Info** → open disk info dropdown
- When the edit control has focus, **Tab / Shift+Tab** must first exit edit mode, then move focus to the next/previous region

## Theme System

NavigationView must support the shared theme system with **Light**, **Dark**, **Rainbow**, and **System High Contrast** modes.

**Rules:**
- Do not hard-code RGB values for UI states; use theme-provided colors for background, text, separators, hover, and pressed states.
- High Contrast mode must prefer Windows system colors and maximize readability.
- NavigationView should visually reflect the **focused pane**:
  - Active pane: normal background + accent bottom border.
  - Inactive pane: noticeably dimmed palette while keeping text readable.
  - Pane focus is driven by `FolderWindow::UpdatePaneFocusStates()` calling `NavigationView::SetPaneFocused(...)`.
  - In Rainbow mode, any per-segment rainbow accents (e.g., breadcrumb underline / current segment text) must also be dimmed for the inactive pane.

**Rainbow Mode (Implementation Guidance):**
- Keep a neutral light/dark base for readability.
- Apply rainbow accents via a stable hue cycle (e.g., stable hash of breadcrumb segment full path and/or menu item text).
- Ensure text contrast remains readable (selection text color chosen based on background luminance).

**Application Integration (Current Implementation):**
- Theme is selected at the application level via the top menu: **View → Theme** (radio-check marks).
- The menu includes built-in themes and any `user/*` themes found in settings and/or `Themes\\*.theme.json5` next to the executable.
- **High Contrast** is system-controlled and always overrides the selected theme:
  - Menu shows **High Contrast (System)** as **checked + disabled** when enabled.
  - Switching System/Light/Dark/Rainbow updates the radio check, but the effective palette remains High Contrast until Windows High Contrast is turned off.
- Theme changes re-apply to:
  - NavigationView (GDI + Direct2D rendering)
  - Popup menus (owner-drawn using the menu theme)
  - FolderView (via FolderWindow propagation)
  - Window titlebar (DWM attributes; best-effort)

## IconCache Integration

### Special Folder Detection with Linear Search

The `IconCache` class maintains a small set of Windows special folder root paths (Desktop/Documents/Downloads/etc.) and uses **linear search** to check if a given path is exactly equal to one of those roots (`IconCache::IsSpecialFolder()`).

**Why Linear Search?**
- Only a small, fixed set of special folders to check (currently: Desktop, Documents, Downloads, Pictures, Music, Videos, OneDrive)
- Linear search is **O(n)** but n is very small (< 20)
- Simplest implementation with minimal overhead
- No need for hash table or binary search complexity
- Cache-friendly due to small dataset

**Special Folder Initialization**:

Uses modern `KNOWNFOLDERID` constants and WIL RAII for memory management. Special-folder initialization is guarded by `std::call_once` to avoid races.

**Path Prefix Matching** (for menu button icon):

To determine if the current path is **under** a special folder (not just equal), NavigationView calls `IconCache::TryGetSpecialFolderForPathPrefix()` (boundary-aware, case-insensitive prefix match that returns the best/longest match).

**Icon Caching Strategy**:
- NavigationView and FolderView treat `iconIndex` as the **system image list icon index** (from `SHGetFileInfoW(...SHGFI_SYSICONINDEX...)`).
- `IconCache` caches `ID2D1Bitmap1` by `(ID2D1Device*, iconIndex)` (LRU), so multiple views share the same converted bitmaps.
- Menu item bitmaps are created from `iconIndex` via `IconCache::CreateMenuBitmapFromIconIndex(...)` (WIC pipeline; correct alpha; avoids manual `DestroyIcon`/`DeleteObject`).

## Performance Considerations

### Section 2 (Direct2D) Optimization
- **Lazy Rendering**: Only redraw on path change, hover, or resize
- **Cached Layouts**: Store `IDWriteTextLayout` for each segment
- **Dirty Regions**: Use `Present1` with dirty rects for minimal updates
- **Hit-Test Cache**: Pre-calculate segment bounds, update only on layout change
- **Async Path Validation**: Don't block UI while checking network paths

### Sections 1 & 3 (GDI) Optimization
- **Static Controls**: Minimal repainting required
- **Font Caching**: Create fonts once, reuse throughout lifetime
- **Owner-Draw**: Menu button only redraws on state change (hover/press)

### Overall Performance
- **Target**: 60 FPS for smooth hover effects in breadcrumb mode
- **Memory**: Direct2D resources only for Section 2 (~2-5 MB)
- **Startup**: Lazy initialization of swap chain on first paint
- **History Limit**: bounded by `folders.historyMax` (default `20`, clamped `1..50`) to prevent unbounded growth

## Testing Checklist

- [ ] Renders correctly at 96, 120, 144, 192 DPI
- [x] WM_DPICHANGED message received when scaling changes
- [ ] Menu bitmap icons remain correct size (16x16) at all DPI settings
- [x] FolderView icons scale appropriately with DPI
- [x] Menu icons have perfect transparency without black borders
- [x] Breadcrumb segments are clickable and navigate correctly
- [x] Hover effects appear on breadcrumb segments
- [ ] Single-click enters edit mode with text selected
- [x] Enter accepts, Escape cancels edit mode
- [ ] Invalid paths show error message
- [ ] History dropdown shows recent paths (bounded by `folders.historyMax`)
- [ ] Menu button shows drives with icons and space info
- [x] Disk info updates when path changes
- [x] Disk info menu shows detailed information
- [ ] Long paths collapse with clickable `"..."` + full-path popup (breadcrumb mode)
- [ ] Network paths handled gracefully (no UI freeze)
- [ ] Light/Dark/Rainbow themes display correctly
- [ ] High contrast mode displays correctly
- [ ] Keyboard shortcuts (F4, Alt+D, Ctrl+L) work
- [ ] Tab navigation cycles through sections
- [x] Direct2D swap chain resizes correctly
- [x] No memory leaks (check with Task Manager)
- [x] Smooth 60 FPS hover animations

## Future Enhancements

### Visual Enhancements
1. **Folder Icons**: Show small icon before each breadcrumb segment
2. **Syntax Highlighting**: Different colors for drive, folders, current
3. **Smooth Transitions**: Fade animation when switching paths
4. **Progress Bar**: Visual disk usage bar in Section 3
5. **Badges**: Show lock icon for read-only, cloud icon for OneDrive

### Functional Enhancements
1. **Auto-complete**: Suggest paths as user types in edit mode (implemented: cache-first autosuggest popup)
2. **Recent Files**: Add to menu dropdown
3. **Favorites/Bookmarks**: Star frequently used paths
4. **Search Integration**: Quick search box (Ctrl+F)
5. **Copy Path**: Context menu to copy path to clipboard
6. **Network Locations**: Dedicated section in menu

### Advanced Features
1. **Path Animations**: Smooth scroll for long breadcrumb lists
2. **Drag & Drop**: Drag files to breadcrumb segments
3. **Context Menus**: Right-click segments for folder operations
4. **Keyboard Navigation**: Arrow keys to move between segments
5. **Touch Support**: Swipe gestures for touch screens

## References

- **Windows Explorer**: Address bar breadcrumb navigation model
- **Direct2D**: Hardware-accelerated 2D graphics rendering
- **DirectWrite**: Advanced text layout and rendering
- **Win32 API**: Window management and GDI rendering
- **uxtheme.h**: Modern control theming (Sections 1 & 3)
- **C++17 filesystem**: Path manipulation and validation
- **FolderView**: Existing Direct2D/DXGI architecture patterns

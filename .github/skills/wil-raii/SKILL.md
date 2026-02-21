---
name: wil-raii
description: Windows Implementation Library (WIL) RAII patterns for managing Windows resources. Use when creating, managing, or cleaning up Windows handles like HICON, HWND, HBITMAP, HDC, COM objects, or any GDI resources.
metadata:
  author: DualTail
  version: "1.0"
---

# WIL RAII Resource Management

**CRITICAL: NO MANUAL RESOURCE CLEANUP** - All Windows resources MUST use RAII wrappers.
`goto` is prohibited; use `wil::scope_exit` for cleanup and common exit paths.

## WIL Wrapper Reference

| Resource | WIL Wrapper |
|----------|-------------|
| HICON | `wil::unique_hicon` |
| HWND | `wil::unique_hwnd` |
| HBITMAP | `wil::unique_hbitmap` |
| HFONT | `wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject>` |
| HBRUSH | `wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject>` |
| HPEN | `wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject>` |
| HDC | `wil::unique_hdc` |
| Paint | `wil::unique_hdc_paint` with `wil::BeginPaint(hwnd, &ps)` |
| COM | `wil::com_ptr<T>` |
| Cleanup | `wil::scope_exit` with lambda |

## Patterns

### Member Variables
```cpp
class MyWindow 
{
    wil::unique_hicon _icon;
    wil::unique_hwnd _tooltip;
    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _font;
};
```

### Paint Context (auto EndPaint)
```cpp
PAINTSTRUCT ps;
wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd, &ps);
FillRect(hdc.get(), &ps.rcPaint, brush);
// Automatic EndPaint on scope exit
```

### GDI Selection (auto restore)
```cpp
{
    auto oldBrush = wil::SelectObject(hdc, _brush.get());
    // Drawing...
} // Old brush automatically restored
```

### Direct2D BeginDraw/EndDraw
```cpp
{
    HRESULT hr = S_OK;
    {
        _d2dContext->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hr = _d2dContext->EndDraw(); });
        // Drawing operations...
    }
    if (FAILED(hr)) { /* handle */ }
}
```

### COM Pointers (`wil::com_ptr<T>`)

Use `wil::com_ptr<T>` for COM interface lifetimes and prefer single-step copy semantics.

```cpp
// ✅ Copy (AddRef) in one step
wil::com_ptr<IFileSystem> fs = rawFs;
fs = rawFs;

// ✅ Receive an owning ref via out-param (avoids attach())
wil::com_ptr<IMuffin> muffin;
HRESULT hr = GetMuffin(muffin.put());
```

Avoid manual ref-counting + `attach()` on the same pointer:
```cpp
// ❌ Two-step hazard (can leak a ref if an exception occurs between lines)
rawFs->AddRef();
fs.attach(rawFs);
```

## Required Include
```cpp
#pragma warning(push)
#pragma warning(disable: 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)
```

## Resource Creation and Access

```cpp
// Creation
_icon.reset(LoadIcon(...));
_tooltip.reset(CreateWindowExW(...));

// Access raw handle for Win32 APIs
DrawIcon(hdc, x, y, _icon.get());
SendMessageW(_tooltip.get(), WM_SETFONT, 
    reinterpret_cast<WPARAM>(_font.get()), TRUE);

// Transfer ownership
wil::unique_hbitmap bitmap(CreateCompatibleBitmap(...));
_bitmaps.emplace_back(std::move(bitmap));

// Clear resources
_bitmaps.clear();  // All HBITMAPs automatically deleted
_icon.reset();     // Old icon destroyed, set to nullptr
```

### Destroying Owned Windows (`wil::unique_hwnd`)

If an `HWND` is owned by `wil::unique_hwnd`, destroy via the wrapper (do not call `DestroyWindow(_hWnd.get())`).

```cpp
// ✅ Preferred: destroy + clear ownership
_hWnd.reset();

// ✅ Transfer ownership explicitly (rare)
HWND hwnd = _hWnd.release();
DestroyWindow(hwnd);
```

## Menu Cleanup with scope_exit

```cpp
HMENU menu = CreatePopupMenu();
auto menuCleanup = wil::scope_exit([&] { if (menu) DestroyMenu(menu); });
// Use menu...
// Automatic cleanup on scope exit
```

## Multiple GDI Selections

```cpp
{
    auto oldPen = wil::SelectObject(hdc, _pen.get());
    {
        auto oldBrush = wil::SelectObject(hdc, _brush.get());
        // Draw with pen and brush
    } // Brush restored
    // Draw with pen only
} // Pen restored
```

## Exception: Transfer-of-Ownership

Cross-thread icon passing is acceptable when properly documented and receiver takes RAII ownership:
```cpp
// Move unique_hicon across threads
wil::unique_hicon icon = ExtractOnBackgroundThread();
PostToUIThread([icon = std::move(icon)]() {
    // UI thread now owns the icon
});
```

## PROHIBITED

`goto` cleanup patterns are forbidden (prefer early returns + `wil::scope_exit`).

```cpp
// ❌ NEVER manual cleanup
DestroyIcon(_icon);
DeleteObject(font);
EndPaint(hWnd, &ps);

// ❌ NEVER destroy an owned HWND via .get()
DestroyWindow(_hWnd.get());

// ❌ NEVER raw handle with manual delete
HFONT font = CreateFontW(...);
// ... use font ...
DeleteObject(font);

// ❌ NEVER vector of raw handles
std::vector<HBITMAP> bitmaps;
for (auto bmp : bitmaps) {
    DeleteObject(bmp); // Manual cleanup loop
}
```

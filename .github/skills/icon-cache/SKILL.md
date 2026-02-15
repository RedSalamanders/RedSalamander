---
name: icon-cache
description: Shell icon management and caching using IconCache class. Use when extracting, caching, or rendering Windows shell icons, working with system image lists, or creating menu bitmaps from icons.
metadata:
  author: DualTail
  version: "1.0"
---

# Icon Cache Management

**Single Owner Rule**: All shell icon queries must go through `RedSalamander/IconCache.*`

## Icon Index Semantics

Treat `iconIndex` as the **system image list index** (`SHGFI_SYSICONINDEX`).
Do NOT hash paths or invent synthetic indices.

## API Reference

### Resolve Icon Indices
```cpp
auto idx = IconCache::GetInstance().QuerySysIconIndexForPath(path, /*fileAttributes*/ 0, /*useFileAttributes*/ false);
auto idx = IconCache::GetInstance().QuerySysIconIndexForKnownFolder(folderId);
auto idx = IconCache::GetInstance().GetOrQueryIconIndexByExtension(ext, FILE_ATTRIBUTE_NORMAL);
auto special = IconCache::GetInstance().TryGetSpecialFolderForPathPrefix(path);

if (idx.has_value())
{
    // idx.value() is the system image list index
}
```

### Render with Direct2D
```cpp
auto bitmap = IconCache::GetInstance().GetIconBitmap(iconIndex, _d2dContext.get());
auto bitmap = IconCache::GetInstance().GetCachedBitmap(iconIndex, _d2dContext.get());
bool cached = _d2dDevice && IconCache::GetInstance().HasCachedIcon(iconIndex, _d2dDevice.get());
```

### Menu Bitmaps (GDI)
```cpp
wil::unique_hbitmap bmp = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(iconIndex, sizePx);
wil::unique_hbitmap bmp = IconCache::GetInstance().CreateMenuBitmapFromIcon(hIcon, sizePx);
wil::unique_hbitmap bmp = IconCache::GetInstance().CreateMenuBitmapFromPath(path, sizePx);
```

## Threading Model

### Contract
- `IconCache::Initialize(...)` is a **UI-thread / STA** responsibility (COM must already be initialized there).
- Any worker thread that calls `IconCache::ExtractSystemIcon(...)` must initialize COM as **MTA** (e.g. `wil::CoInitializeEx(COINIT_MULTITHREADED)`).

1. **Background thread**: Extract icon
   ```cpp
   [[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);
   wil::unique_hicon icon = IconCache::GetInstance().ExtractSystemIcon(iconIndex, targetDipSize);
   ```

2. **UI thread**: Convert to D2D bitmap
   ```cpp
   auto bitmap = IconCache::GetInstance().ConvertIconToBitmapOnUIThread(icon.get(), iconIndex, _d2dContext.get());
   ```

## Thread-Safety Notes (When Modifying IconCache)

- Protect shared state consistently: if a member is guarded by the class mutex, **read and write it only under that lock** (avoid unlocked reads of non-atomic flags).
- Prefer “self-guarding” helpers (take the lock inside the helper) over ad-hoc double-checked patterns in callers.
- Treat the system image lists (`IImageList`) as long-lived process resources: store them as `wil::com_ptr<IImageList>` and avoid resetting them during normal cache clears so background icon extraction can remain lock-free.
- `IImageList` is a COM interface. In this codebase, extraction is allowed on worker threads **only** when that worker has COM initialized as **MTA**.

## RAII Requirements

Always use `wil::unique_hicon` / `wil::unique_hbitmap`.
Never call `DestroyIcon` / `DeleteObject` manually.

```cpp
// Transfer ownership to vector
wil::unique_hbitmap bmp = IconCache::CreateMenuBitmapFromIconIndex(idx);
_menuBitmaps.emplace_back(std::move(bmp));
```

## DPI Handling

```cpp
// Call from DPI change handlers
IconCache::SetDpi(newDpi);
```

Menu bitmaps use physical pixels: `GetSystemMetrics(SM_CXSMICON)`

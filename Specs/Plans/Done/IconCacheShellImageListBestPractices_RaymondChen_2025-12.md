# Icon Cache Implementation Analysis & Shell Image List Best Practices

## Implementation Plan (Live)

> Owner: Codex (this section is updated as changes land)
> Goal: Improve icon correctness, perceived speed (reduce placeholders), quality (better source sizes), and resource hygiene.
> Constraints: Follow `AGENTS.md` + `.github/skills/*` (WIL RAII, IconCache single-owner, COM threading contract).

- [ ] Baseline + success criteria (use existing `Debug::Perf` + FolderView icon telemetry; optional ETW scripts).
- [x] Correctness: case-insensitive extension keys + per-file whitelist works before `IconCache::Initialize()`.
- [x] Quality: add `SHIL_JUMBO` support and integrate into selection + fallback.
- [x] Resource hygiene: add `IconCache::ClearDeviceCache(ID2D1Device*)` and call from D2D discard paths.
- [x] Workflow: reduce “placeholder never becomes real icon” on slow machines by re-prioritizing *currently visible* icon requests during scrolling.
- [x] Telemetry: fix `IconCache::GetMemoryUsage()` to reflect real bitmap sizes; update misleading comments.

## Chat Session Summary - December 30, 2025

---

## Topic 1: Raymond Chen's Blog Post on High-Resolution Icons

### Question
User asked about icon extraction strategy based on Raymond Chen's blog post:
https://devblogs.microsoft.com/oldnewthing/20140120-00

### The Fastest Correct Pattern

There is no supported "much faster" API where you open Explorer's `iconcache_*.db` and pull association icons out of it. Those files are an internal implementation detail, frequently locked by Explorer, and they don't give you a clean "extension ? icon pixels" mapping you can rely on across Windows versions, DPI, themes, per-user defaults, overlays, etc.

The **public** fast path is the **system image list** (Shell icon cache in memory), not the on-disk DB.

If what you're doing "isn't fast enough", it's almost always because you're still doing one of these slow things:
- Calling `SHGetFileInfo` on **full paths** (forces filesystem / handlers) instead of **extensions**
- Requesting `SHGFI_ICON` (creates an `HICON`) and/or converting `HICON ? bitmap` for every row
- Loading icons for **every item up front** instead of only what's visible
- Asking for overlays/open-icons/type names, which triggers extra work/handlers

### Recommended Approach

1. **Use `SHGetImageList`** instead of `SHGetFileInfo` for large icons:
```cpp
IImageList* pImageList = nullptr;
SHGetImageList(SHIL_JUMBO, IID_PPV_ARGS(&pImageList)); // 256x256
// or SHIL_EXTRALARGE (48x48), SHIL_LARGE (32x32), SHIL_SMALL (16x16)
```

2. **Get the icon index** using `SHGetFileInfo` with `SHGFI_SYSICONINDEX`:
```cpp
SHFILEINFOW sfi{};
SHGetFileInfoW(filePath, 0, &sfi, sizeof(sfi), SHGFI_SYSICONINDEX);
int iconIndex = sfi.iIcon;
```

3. **Extract the icon** from the image list:
```cpp
HICON hIcon = nullptr;
pImageList->GetIcon(iconIndex, ILD_TRANSPARENT, &hIcon);
```

---

## Topic 2: Challenging Current IconCache.cpp Implementation

### Best Practices Recommendations

The fastest correct pattern for association icons:

#### 1. Only resolve per **extension** (and cache the result)
Use `SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES` with just "`.ext`" so there is **no disk access**.

```cpp
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#include <unordered_map>
#include <string>
#include <mutex>
#include <algorithm>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Comctl32.lib")

static std::unordered_map<std::wstring, int> g_extToSysIndex;
static std::mutex g_cacheMx;

static std::wstring NormalizeExt(std::wstring ext)
{
    // expects ".txt" or "txt" -> ".txt" lowercased
    if (!ext.empty() && ext[0] != L'.') ext.insert(ext.begin(), L'.');
    std::transform(ext.begin(), ext.end(), ext.begin(),
                   [](wchar_t c){ return (wchar_t)CharLowerW((LPWSTR)(ULONG_PTR)c); });
    return ext;
}

int GetSysImageIndexForExtensionCached(const std::wstring& extIn)
{
    std::wstring ext = NormalizeExt(extIn);

    {   // fast cache hit
        std::scoped_lock lk(g_cacheMx);
        auto it = g_extToSysIndex.find(ext);
        if (it != g_extToSysIndex.end())
            return it->second;
    }

    SHFILEINFOW sfi{};
    DWORD flags = SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES | SHGFI_SMALLICON;

    // Important: must have COM initialized on the calling thread.
    HIMAGELIST himl = (HIMAGELIST)SHGetFileInfoW(
        ext.c_str(),
        FILE_ATTRIBUTE_NORMAL,
        &sfi,
        sizeof(sfi),
        flags
    );

    if (!himl)
        return -1;

    {   // store
        std::scoped_lock lk(g_cacheMx);
        g_extToSysIndex.emplace(ext, sfi.iIcon);
    }
    return sfi.iIcon;
}
```

If you're currently doing this per-file (100k calls), switching to **per-extension** (maybe 20–200 calls total) is the only way you get "much faster" while staying correct.

#### 2. Never extract `HICON`s for list/grid rendering

Attach the **system image list** to your control and render by **index**. `SHGetFileInfo` explicitly returns the system image list handle when you use `SHGFI_SYSICONINDEX`.

**ListView setup:**
```cpp
HIMAGELIST GetSystemSmallImageList()
{
    SHFILEINFOW sfi{};
    return (HIMAGELIST)SHGetFileInfoW(
        L"C:\\", 0, &sfi, sizeof(sfi),
        SHGFI_SYSICONINDEX | SHGFI_SMALLICON
    );
}

// once:
HIMAGELIST sysSmall = GetSystemSmallImageList();
ListView_SetImageList(hwndLV, sysSmall, LVSIL_SMALL);

// per item:
LVITEMW it{};
it.mask = LVIF_TEXT | LVIF_IMAGE;
it.iItem = row;
it.pszText = (LPWSTR)name.c_str();
it.iImage = GetSysImageIndexForExtensionCached(L".txt");
ListView_InsertItem(hwndLV, &it);
```

**Owner-draw / custom control:**
```cpp
// draw directly from system imagelist (no HICON creation)
ImageList_Draw(sysSmall, sysIndex, hdc, x, y, ILD_TRANSPARENT);
```

This is about as fast as it gets because you're drawing from the shared cache instead of extracting/decoding icons.

#### 3. Only fetch icons for visible items (virtualize)

If you're filling a grid with 50,000 rows and you "need icons now", you're doing wasted work.
Use a virtual list (owner-data) and resolve icons only for the visible range + small prefetch buffer. Everything else stays as "unknown/placeholder" until it scrolls into view.

This is how Explorer and fast file managers stay responsive.

---

## Topic 3: Analysis of Current Implementation

### ? What's Already Done Right

| Feature | Implementation |
|---------|----------------|
| Extension-based caching | `SHGFI_USEFILEATTRIBUTES` in `QuerySysIconIndexForPath` |
| Caching by system image index | `_extensionToIconIndex` map |
| Using `SHGetImageList` for high-quality icons | `SHIL_EXTRALARGE/LARGE/SMALL` with fallback |
| LRU eviction strategy | `EvictLRUIfNeeded` prevents unbounded memory |
| Per-file lookup only for special extensions | `_perFileLookupExtensions` whitelist for `.exe`, `.lnk`, `.dll`, etc. |

### ?? Areas of Concern

1. **Per-file lookups use `useFileAttributes=false`** - triggers filesystem access (intentional for `.exe`, `.lnk`)
2. **HICON ? D2D Bitmap conversion** - unavoidable with Direct2D architecture
3. **Mutex contention** - could use `std::shared_mutex` for reads

---

## Topic 4: Missing SHIL_JUMBO Support

### Issue
Current code only uses:
- `SHIL_EXTRALARGE` (48×48)
- `SHIL_LARGE` (32×32)
- `SHIL_SMALL` (16×16)

For high-DPI displays (150%+), should use **SHIL_JUMBO** (256×256).

### Recommended Fix

#### In Initialize():
```cpp
// NEW: Get SHIL_JUMBO (256×256) for high-DPI displays
if (! _systemImageListJumbo)
{
    HRESULT hrJumbo = SHGetImageList(SHIL_JUMBO, IID_PPV_ARGS(&_systemImageListJumbo));
    if (SUCCEEDED(hrJumbo))
    {
        DBGOUT_INFO(L"IconCache: Initialized SHIL_JUMBO (256×256) at {:.0f} DPI", dpi);
    }
}
```

#### In ExtractSystemIcon():
```cpp
// Try optimal size first (now includes JUMBO)
if (optimalSize == SHIL_JUMBO)
{
    if (auto icon = tryExtract(_systemImageListJumbo))
    {
        return icon;
    }
}
// ... existing fallback cascade
```

#### In SelectOptimalImageListSize():
```cpp
int IconCache::SelectOptimalImageListSize(float targetDipSize) const
{
    const float targetPixels = targetDipSize * _dpi / 96.0f;

    // JUMBO (256×256) for very high-DPI or large icon views
    if (targetPixels >= 64.0f)
    {
        return SHIL_JUMBO;
    }

    if (targetPixels >= 40.0f)
    {
        return SHIL_EXTRALARGE;
    }

    if (targetPixels >= 24.0f)
    {
        return SHIL_LARGE;
    }

    return SHIL_SMALL;
}
```

#### In Header:
```cpp
IImageList* _systemImageListJumbo = nullptr;
```

#### In Clear():
```cpp
if (_systemImageListJumbo)
{
    _systemImageListJumbo->Release();
    _systemImageListJumbo = nullptr;
}
```

---

## Topic 5: Texture Atlas Consideration

### Question
Would a texture atlas design be beneficial?

### Analysis

#### For FolderView (Direct2D rendering) - **Potentially Yes**
| Benefit | Applies? |
|---------|----------|
| Reduced GPU state changes | ? Yes - drawing many icons in a list/grid |
| Single texture bind | ? Yes - one `DrawBitmap` call with source rects |
| Better GPU cache locality | ? Yes - icons packed together |
| Reduced memory fragmentation | ? Yes - one large allocation vs. many small |

#### For Menus (GDI owner-draw) - **No Benefit**
- GDI can't efficiently blit from atlas sub-regions
- Each menu item needs a separate `HBITMAP` for `MENUITEMINFO`
- Menu rendering is infrequent (on-demand)

### Recommendation
**Skip the atlas for current use case.**

Current `IconCache` with per-index `ID2D1Bitmap1` caching is already efficient. The performance bottleneck (if any) is more likely:
1. **Icon extraction** from system image list (JUMBO fix addresses this)
2. **WIC conversion** (already cached, one-time cost)
3. **Per-file lookups** for `.exe`/`.lnk` (already parallelized on threadpool)

### Alternative Quick Win - Sprite Batching

If you want to reduce draw calls without an atlas, consider **sprite batching** in FolderView rendering:

```cpp
// Pseudo-code: batch icons with same index
std::unordered_map<int, std::vector<D2D1_RECT_F>> iconBatches;

for (const auto& item : visibleItems)
{
    iconBatches[item.iconIndex].push_back(item.destRect);
}

for (const auto& [iconIndex, rects] : iconBatches)
{
    auto bitmap = _iconCache.GetIconBitmap(iconIndex, d2dContext);
    for (const auto& rect : rects)
    {
        d2dContext->DrawBitmap(bitmap.get(), rect);
    }
}
```

This keeps the same icon bound while drawing all instances, which Direct2D can optimize internally.

---

## Summary Table

| Aspect | Current State | Recommendation |
|--------|--------------|----------------|
| Extension caching | ? Implemented | Keep as-is |
| System image list sizes | ?? Missing JUMBO | Add SHIL_JUMBO (256×256) |
| Per-file lookups | ? Parallelized on threadpool | Keep as-is |
| Bitmap caching | ? LRU eviction | Keep as-is |
| Texture atlas | ? Not implemented | Skip for now |
| Mutex type | `std::mutex` | Consider `std::shared_mutex` for reads |

---

## "Can I parse iconcache_*.db anyway?"

You can (people do it for forensics), but it's the wrong tool for a fast, correct UI:
- internal/undocumented format (changes between Windows versions)
- frequently locked by Explorer
- stale when associations/theme/DPI change
- doesn't naturally key by "extension association" the way your app needs

And you still end up doing disk IO + decoding, which is not faster than drawing from the in-memory system image list.

---

## Bottom Line

If `SHGetFileInfo` "isn't fast enough", the fix is not "read Explorer's icon cache DB".
The fix is:

1. **One Shell query per extension** (`SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES`)
2. **Render by imagelist index** (no `HICON`, no bitmap conversion)
3. **Virtualize** (only resolve what's visible)

---

## References

- [SHGetFileInfo - Microsoft Learn](https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shgetfileinfoa)
- [SHGetImageList - Microsoft Learn](https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shgetimagelist)
- [Raymond Chen - The Old New Thing: How do I get a high resolution icon for a file?](https://devblogs.microsoft.com/oldnewthing/20140120-00/?p=2043)

---

*Generated: December 30, 2025*
*Project: Red Salamander*
*Repository: https://github.com/DualTail/RedSalamander*
*Branch: master*

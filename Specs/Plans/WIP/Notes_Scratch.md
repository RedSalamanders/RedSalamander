# Scratch file - to be deleted later
DON'T USE THE CONTENT FOR ANY PURPOSE
SKIP THIS FILE WHEN READING DOCUMENTATION
LET IT ALONE THERE (HUMAN MANAGED)

(for human) DON'T FORGET TO DELETE THIS FILE LATER




------------------------------------------------------------------------------------------------------
## Ctrl +X / Ctrl +V 
after Ctrl+X Ctrl+V remove the content of the clipboard to avoid multiple move  with error and because there is no need to keep the content in the clipboard after pasting it in the file explorer or RedSalamander. 

------------------------------------------------------------------------------------------------------
## password in connection manager is not wolrking 
you could type what you want nothing happen





------------------------------------------------------------------------------------------------------
## All windows must be resizable
- all windows must be resizable by the user to fit their content and be more comfortable to use. for example the file operation window must be resizable to be able to see all the files when there is a lot of them, the space viewer must be resizable to be able to see more details when there is a lot of items, etc ...
- all windows must have a minimum size to avoid having too small windows that are not usable but other than that the user must be able to resize the window as he want to fit his needs and content displayed in the window. for example the file operation window must be resizable to be able to see all the files when there is a lot of them, the space viewer must be resizable to be able to see more details when there is a lot of items, etc ...

------------------------------------------------------------------------------------------------------
## file operation
- when preflight is skip display the number of files and folders copied
- when we are in parallel mode the thread pool must be more share between all task
   - after a file end in a task the thread could be reassign to another task. the purpose is to have the smal amount of thread (more or less) for all tasks




------------------------------------------------------------------------------------------------------
## ftp client
- on empty folder display . and .. file, delete this folder is in error
- copy file just displayed preparing when doing the copy no callback and progress
- ftp pcopy is slow with peak of copy and long period without activities

------------------------------------------------------------------------------------------------------
## Connection manager
clicking connect is validating information for the same amout of time than validating Windows Hello
Don't need to display Windows hello right after connection manager connect button

------------------------------------------------------------------------------------------------------
## Connection manager
in navigation bar display emoji in color




------------------------------------------------------------------------------------------------------
## Review UpdateWindow usage in the code
review all UpdateWindow usage in the code to ensure proper usage
propose to replace UpdateWindow by InvalidateRect + UpdateWindow when needed to ensure proper repainting of the window






-----------------------------------------------------------------------------------------------------
## review behavior for connecting / disconnecting  drives
- when removing an usb key browse by a pane I don't see the dsiconnected message
- when the other pane is browsing the same network drive and the network drive is disconnected I don't see the disconnected message



-------------------------------------------------------------------------------------------------------
# for copying a selection of files from one pane to another
the message must display the folder source when there is multiple items to copy
example: Copy 23 files, 2 folders  from C:\folder\ to D:\otherfolder\

when there is only one item to copy display the full path of the item being copied
example: Copy C:\folder\file.txt to D:\otherfolder\




-------------------------------------------------------------------------------------------------------
# FolderView and callback mechanism review
FolderWatching is not perfectly managed in FolderView leading to some unexpected behavior in some cases

the watch callback must be enhanced with the type of operation (create, delete, rename, modify, etc ...)
```cpp
interface __declspec(novtable) IFileSystemDirectoryWatchCallback
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* cookie) noexcept = 0;

    // this method are optional and could be called depending on the notification type
    virtual HRESULT STDMETHODCALLTYPE FileCreated(const wchar_t* fileName, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE FileDeleted(const wchar_t* fileName, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE FileRenamed(const wchar_t* oldFileName, const wchar_t* newFileName, void* cookie) noexcept = 0;
};
```

when focus is on an item and a delete callback notification occur the focus move to the next item if there is any or stay on the last one if there no more next item
when focus is on an item and a create callback notification occur the focus stay on the current item
when focus is on an item and a rename callback notification occur the focus stay on the current item and move the view to the new focus link to the sort  (filter, sort, etc ...)

global rule: focus stay on item or move to next

review all code and spec to ensure this behavior is implemented everywhere in the code and documented in `Specs/UI/UI_FolderView.md`

-------------------------------------------------------------------------------------------------------
# review Status window
- when the window appear in dark theme and it's slow its start by a white window background not following the theme we need to do the same on this one than others to be sure to follow the theme always even when there is nothing to display yet
- the speed limit button must be visible for the calculation phase I could want to set the speed limit when the number of items
- the speed will be apply from the begining of the copy



-------------------------------------------------------------------------------------------------------
# review the cache mechanism in FolderView
when navigating to a folder in one pane after navigation finish if I'm navigating to the same folder in the other pane the content must  loaded from cache. Seems this is not working properly now. for example I'm seeing a loadding in the first go to the other pane even if the folder was already loaded in the first pane and I'm seeing another loadding in this second pane



-------------------------------------------------------------------------------------------------------
# IHost modification review
IHost must implement
- Navigate to a folder in a pane option are FocusedPane, LeftPane, RightPane and a path
- FocusItem in a pane option are FocusedPane, LeftPane, RightPane and an index to focus in the pane (could be a navigate to the path and focusing the item the display will move accordingly to show the focused item)
- ChangeSelection in a pane option are FocusedPane, LeftPane, RightPane and a list of indexes to select in the pane with option to add to selection, remove from selection, replace selection
- GetCurrentFolder in a pane option are FocusedPane, LeftPane, RightPane returning the current folder path in the pane
- GetFocusedItem in a pane option are FocusedPane, LeftPane, RightPane returning the index of the focused item in the pane
- GetSelection in a pane option are FocusedPane, LeftPane, RightPane returning the list of indexes of selected items in the pane
all path must supporting navigating to virtual file system path like ftp://, s3://, zip:// etc ...
review all code using IHost to use these new methods instead of doing the same code again in the plugin




-------------------------------------------------------------------------------------------------------




## cmd/pane/alternateView
## cmd/pane/bringCurrentDirToCommandLine
## cmd/pane/bringFilenameToCommandLine
## cmd/pane/changeAttributes
## cmd/pane/changeCase
## cmd/pane/changeDirectory
## cmd/pane/clipboardCopy
## cmd/pane/clipboardPaste
## cmd/pane/connect
## cmd/pane/contextMenu
## cmd/pane/contextMenuCurrentDirectory
## cmd/pane/copyNameAsText
## cmd/pane/copyPathAndFileName
## cmd/pane/copyPathAndNameAsText
## cmd/pane/copyPathAsText
## cmd/pane/disconnect
## cmd/pane/driveInformation
## cmd/pane/edit
## cmd/pane/editNew
## cmd/pane/alternateEdit
## cmd/pane/filter
## cmd/pane/find


## cmd/pane/listOpenedFiles
## cmd/pane/showFoldersHistory
## cmd/pane/loadSelection
## cmd/pane/openCurrentFolder
## cmd/pane/openProperties
## cmd/pane/pack
## cmd/pane/permanentDelete
## cmd/pane/quickSearch
## cmd/pane/saveSelection
## cmd/pane/shares
## cmd/pane/unpack
## cmd/pane/userMenu
## cmd/pane/windowMenu
## cmd/pane/zoomPanel









------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------
# DONE



------------------------------------------------------------------------------------------------------
# DONE





------------------------------------------------------------------------------------------------------
## DONE preference dialog
if settings is dirty pressing esc display a message first to ask user if he want to save changes before closing the dialog box
if no changes done pressing esc close the dialog box directly without message

------------------------------------------------------------------------------------------------------
## DONE IMAP 
- display received date time as file time
- display email size as file size
- display email subject as file name

------------------------------------------------------------------------------------------------------
## DONE video viewer 
- super long first init need a spinner to wait with a message 
- need to store last volume used and retore it when reopening the video viewer

------------------------------------------------------------------------------------------------------
## DONE all messages in ressources must be positional
{} is forbidden becasue it could be translated in a language with different order and it will be a mess to manage that
all message must be positional with {0}, {1}, etc ... and the order of the parameter in the code must be the same as the order in the message to avoid confusion and error when translating
for example instead of
```text 
{} is not accessible because {}
```
it must be
```text
{0} is not accessible because {1}
```
Add that in normative specs and Skills and make sure to apply that everywhere in the code and ressources


------------------------------------------------------------------------------------------------------
## DONE Acrylic for all windows
- Title bar, menu, context menu, dialog box, etc ... all windows should have acrylic/Mica/Mica Alt effect when the setting is activated
- File Operation window should have acrylic background effect when the setting is activated
- Preferences Settings window should have acrylic background effect when the setting is activated
- Connection window should have acrylic background effect when the setting is activated

------------------------------------------------------------------------------------------------------
## DONE Rename dialog
- When the text is too long the text in dialog go outside the editbox
- the dialog must be horizontally resizable to be able to resize it and see the full text when the text is too long

## DONE cmd/pane/hotPaths
Hot paths are shortcut to saved path
Ctrl + virtual key from 1 to 0 navigate to the saved path
Ctrl + Shift virtual key from 1 to 0: assign the current path to the slot (if something already there displayed it and ask to replace)
in settings there is a dedicated page for the hot path and shorcut management
hot path are editable in this settings page
you could checked Hot Paths will be displayed in the Change Drive menu in a dedicated sub menu if any are check.
at the bottom of the setting page there is a toggle to Open this page preference page when Hot Path is assigned from panel (Ctrl + Shift + key)

## DONE cmd/app/fullScreen
    put the application in fullscreen mode hiding title bar and task bar
    exit fullscreen mode when pressing again the shortcut or pressing ESC key

## DONE cmd/app/openLeftDriveMenu
  open left drive menu

## DONE cmd/app/openRightDriveMenu
  open right drive menu

## DONE cmd/app/viewWidth
  adjust the width of the viewer pane with the keyboard Left / right arrows, enter validf the position, esc cancel the move


** cmd/pane **
## DONE cmd/pane/refresh
  invalidate the cache and refresh the content of the pane

## DONE cmd/pane/changeCase
A dialog box to change case of selected files in the pane
whtith these options:
```text
Title: Change Case
Change Case to
  o Lower case
  o Upper case
  o Partially mixed case (name in mixed, extension in lower)
  o Mixed case
Change
  o Whole filename
  o Only name
  o Only extension
Options
  Include subdirectories

  [ OK ] [ Cancel ]
```

------------------------------------------------------------------------------------------------------
# DONE Command to implement

** cmd/app **

## cmd/app/compare
  will be to compare the 2 folders in the 2 panes and display difference. There is some option to configure this behavior in a dialog before running the compare
  a dialog box with these options is displayed before running the compare
  At first all names will be unselected in both panels.
  Files existing only in one panel will be selected.

  all users setting will be persist in settings to reapply them next time we are using compare command

```text
Compare files with same name by --------------------------------------------
 [] Size           Bigger files will be selected.
 [] Date and time  Newer files will be selected.
 [] Attributes     Files with different attributes will be selected.
 [] Content        Files with different content will be selected

Subdirectories options --------------------------------------------
 [] Compare subdirectories
 Directories with different content will be selected.
 Directories existing only in one panel will be selected.

 [ OK ]   [ Cancel ] [ More ]
```

  when pressing More button another dialog is wider with more options above the button under ' Compare subdirectories'
```text
Additional Options --------------------------------------------
 [] Compare attributes of subdirectories
 Directories with different attributes will be selected.
 [] Select subdirectories contained only in one panel

More options --------------------------------------------
 [] Ignore file:         [                                 ] (text box appear when check; wildcards, separated by ';')
 [] Ignore directories:  [                                 ] (text box appear when check; wildcards, separated by ';')
```

------------------------------------------------------------------------------------------------------
## DONE Space Viewer  display review
- How could I have Scan  Completed and the progress running on the same folder in the same time ?

- review the path display in the space viewer
the path must be middle ellipsis want to see the last folder is possible

- the difference between scan folder an not scan folder is not clear enough
propose a better way to display that like text in the diagonal displaying "Scan Incomplete" or nothing when this is completed

- differences between file and folder is not clear enough
need to be more visible with a clear fiffernce between file and folder display

- the space to display the name in each square must always have minimum one line visible to be able to read the beginning of the name when the suqare is big enough to display one line of text

------------------------------------------------------------------------------------------------------
## DONE in all plugin with selection combobox
- if combo open esc key close the combo
- if combo close esc key move the focus to main window
- after select an item to the combo focus must move to the main window


------------------------------------------------------------------------------------------------------
# DONE S3 File System
- I want a 2 new file system plugin for S3 and S3Table in FileSystemS3 dll
- implement by using S3-crt example: https://github.com/awsdocs/aws-doc-sdk-examples/tree/main/cpp/example_code/s3-crt
- first root of the s3 is the bucket list
- after list folder and file in the bucket
- add all needed setting in preferences for this plugin
- add all the needed information in the connection manager
- an option "S3" and "S3 Table" in the list of available option
- save secret as needed in credential vault
global documentation and code : https://github.com/aws/aws-sdk-cpp
exmaples : https://github.com/awsdocs/aws-doc-sdk-examples/tree/main/cpp/example_code/

implement all that by respecting AGENTS.md and skills in .github/skills

------------------------------------------------------------------------------------------------------
# DONE Filesystem copy/move to and from
- each plugin must explicitely define from and to plugin for copy / move if they are accepting and understand path
- each plugin must declare what they are accepting for rename / delete properties
- other than filesystem implement a properties interfadce to get the detail of the item

-------------------------------------------------------------------------------------------------------
# DONE refactor FileSystem
need to have better separation .h and .cpp
one header per external interface implemented in FileSystem module
for example IFileSystem will be implemented in FileSystem.h & FileSystem.cpp
IFileSystemDirectoryWatch will be implemented in FileSystemDirectoryWatch.h & FileSystemDirectoryWatch.cpp
etc ...
if needed we could have other header/cpp file pair to split big files

------------------------------------------------------------------------------------------------------
# DONE RedSalamander Performance Optimization Plan

**Date:** January 13, 2026
**Status:** Pending Implementation
**Overall Assessment:** The codebase demonstrates excellent performance practices. This plan identifies targeted optimizations for further improvement.

---

## Executive Summary

This comprehensive audit identifies **14 performance optimization opportunities** across memory, CPU, graphics, and I/O. The codebase already implements many best practices including RAII, move semantics, reader-writer locks, and container pre-allocation.

| Severity | Count | Impact |
|----------|-------|--------|
| **Critical** | 2 | High - measurable latency/memory impact |
| **Moderate** | 8 | Medium - noticeable in specific scenarios |
| **Minor** | 4 | Low - polish and consistency |

---

## 1. CRITICAL OPTIMIZATIONS

### 1.1 Pass-by-Value `std::wstring` in Hot Paths

**Problem:** Functions accepting `std::wstring` by value trigger unnecessary heap allocations and copies on every call.

**Affected Locations:**

| File | Function/Parameter | Frequency |
|------|-------------------|-----------|
| `RedSalamander/FolderView.ErrorOverlay.cpp` | Error message parameters | Per-error |
| `RedSalamander/FolderWatcher.cpp` | Path parameters in callbacks | Per-file-change |
| `Plugins/FileSystem7z/` | Archive path parameters | Per-archive-operation |
| `RedSalamander/FolderWindow.FileOperations.cpp` | Lambda captures | Per-file-operation |

**Solution:** Convert to `std::wstring_view` for read-only access or `const std::wstring&` when reference lifetime is guaranteed.

```cpp
// Before (causes copy)
void ShowError(std::wstring message);

// After (no copy)
void ShowError(std::wstring_view message);
```

**Impact:** Reduces heap allocations in file enumeration and error handling paths.

**Priority:** P0 - Implement first

---

### 1.2 Per-Character `towlower()` in Hash Function

**Problem:** `DirectoryInfoCache.cpp` contains a case-insensitive hash function that calls `std::towlower()` for every character:

```cpp
size_t CaseInsensitiveHash(std::wstring_view text) noexcept
{
    uint64_t hash = 14695981039346656037ull;
    for (wchar_t ch : text)
    {
        const wchar_t folded = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
        // FNV-1a computation...
    }
}
```

**Location:** `RedSalamander/DirectoryInfoCache.cpp`

**Impact:** Called on every cache lookup during folder enumeration. For folders with thousands of files, this creates measurable overhead.

**Solutions (choose one):**

1. **Pre-lowercase paths on storage:** Store lowercased keys, compare against lowercased lookups
2. **ASCII fast-path:** Most Windows paths are ASCII; use direct `| 0x20` for A-Z before falling back to `towlower`
3. **Batch lowercasing:** Use `LCMapStringEx` with `LCMAP_LOWERCASE` for entire string at once

```cpp
// Fast-path optimization example
inline wchar_t FastLower(wchar_t ch) noexcept
{
    // ASCII fast path (A-Z -> a-z)
    if (ch >= L'A' && ch <= L'Z')
        return ch | 0x20;
    // Non-ASCII fallback
    return (ch < 128) ? ch : static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
}
```

**Priority:** P0 - High-frequency hot path

---

## 2. MODERATE OPTIMIZATIONS

### 2.1 Vector Pass-by-Value in File Operations

**Problem:** Large vectors copied on function entry for file operations.

**Locations:**
- `RedSalamander/FolderWindow.FileOperations.cpp` - `std::vector<std::wstring>` parameters
- `RedSalamander/FolderView.Enumeration.cpp` - Item vectors

**Solution:** Use `const std::vector<T>&` or move semantics with `std::vector<T>&&`.

**Priority:** P1

---

### 2.2 `_wcsicmp` in Sorting Comparators

**Problem:** `_wcsicmp` involves locale-aware comparison which is slower than ordinal comparison. Used extensively in file/folder sorting.

**Locations:**
- `RedSalamander/FolderView.cpp` - Sort comparators
- `RedSalamander/FolderWindow.FileOperations.cpp` - Path sorting
- `RedSalamander/DirectoryInfoCache.cpp` - Name comparison

**Current Pattern:**
```cpp
return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0;
```

**Solution:** Use `CompareStringOrdinal` for ~20-30% faster sorting:
```cpp
return CompareStringOrdinal(a.name.c_str(), -1, b.name.c_str(), -1, TRUE) == CSTR_LESS_THAN;
```

**Priority:** P1 - Visible improvement on large folders (1000+ files)

---

### 2.3 Missing `reserve()` Before Loops

**Problem:** Some container population loops don't pre-allocate capacity.

**Locations:**
- `RedSalamander/FunctionBar.cpp` - `chordToRows` map population
- Various location with `push_back` in loops

**Solution:** Add `reserve()` calls when size is known or estimable.

**Priority:** P1

---

### 2.4 Repeated `CreateTextLayout` in Hot Paths

**Problem:** DirectWrite `CreateTextLayout` is expensive. Some paths recreate layouts unnecessarily.

**Locations:**
- `RedSalamander/FolderView.Layout.cpp` - Per-item layout during measurement
- `RedSalamanderMonitor/ColorTextView.Rendering.cpp` - Multiple layout calls

**Current Mitigation:** The codebase has `_layoutValid` flags - ensure they're used consistently.

**Solution:**
1. Cache layouts more aggressively
2. Batch layout creation
3. Reuse layout objects when only text changes

**Priority:** P1

---

### 2.5 Direct `std::thread` Creation

**Problem:** Direct `std::thread` creation has overhead. Short-lived tasks should use thread pool.

**Locations:**
- `RedSalamanderMonitor/Monitor.cpp` - `std::thread` with detached threads
- Various async operations

**Positive Pattern Already Used:**
- `RedSalamander/FolderView.Icons.cpp` uses `TrySubmitThreadpoolCallback` correctly

**Solution:** Use Windows Thread Pool APIs or `std::async` with launch policy for short tasks.

**Priority:** P2

---

### 2.6 TextLayout Cache Full-Clear

**Problem:** `ColorTextView._layoutCache` is cleared entirely on minor changes instead of using LRU eviction.

**Location:** `RedSalamanderMonitor/ColorTextView.Layout.cpp`

**Solution:** Implement LRU eviction to preserve valid cached slices.

**Priority:** P2

---

### 2.7 Lock Scope in Cache Operations

**Problem:** Some cache operations hold locks during external API calls.

**Current State:** Most patterns are correct (lock released before shell operations).

**Verification Needed:** Audit `IconCache.cpp` lock scopes.

**Priority:** P2

---

### 2.8 Brush Creation on Theme Change

**Problem:** Multiple `CreateSolidColorBrush` calls during theme change.

**Location:** `RedSalamander/FolderView.Rendering.cpp`

**Current State:** Acceptable since it only happens on theme change, not per-frame.

**Consideration:** Could batch brush creation for consistency.

**Priority:** P3

---

## 3. MINOR OPTIMIZATIONS

### 3.1 String Concatenation vs `std::format`

**Problem:** Some legacy code uses `+` concatenation instead of `std::format`.

**Examples:**
```cpp
// Current
std::wstring(state.pluginShortId) + L":/"

// Better
std::format(L"{}:/", state.pluginShortId)
```

**Priority:** P3 - Consistency improvement

---

### 3.2 Extension Comparison Pattern

**Problem:** Using `_wcsicmp` for known constant extensions.

**Location:** `RedSalamander/FolderView.Enumeration.cpp`
```cpp
item.isShortcut = (_wcsicmp(item.extension.c_str(), L".lnk") == 0);
```

**Solution:** Use direct character comparison or `CompareStringOrdinal`.

**Priority:** P3

---

### 3.3 `InvalidateRect(NULL)` Granularity

**Problem:** Full window invalidation where partial would suffice.

**Current State:** Some areas use targeted invalidation (good), others invalidate entire window.

**Solution:** Use calculated dirty rectangles for partial invalidation.

**Priority:** P3

---

### 3.4 Locale-Dependent Comparisons

**Problem:** Using locale-aware functions where ordinal comparison would be faster and correct.

**Solution:** Replace with `CompareStringOrdinal` where locale semantics aren't required.

**Priority:** P3

---

## 4. POSITIVE PATTERNS (No Changes Needed)

The following excellent patterns are already implemented:

### 4.1 Reader-Writer Locks ✓
**Location:** `RedSalamanderMonitor/Document.cpp`
- Uses `std::shared_lock` for reads
- Uses `std::unique_lock` for writes
- Proper lock granularity

### 4.2 Container Pre-allocation ✓
**Locations:** `Document.cpp`, `FolderView.Enumeration.cpp`
```cpp
_lines.reserve(newCapacity);
_visibleLines.reserve(_lines.size());
```

### 4.3 Brush Caching ✓
**Location:** `RedSalamanderMonitor/ColorTextView`
- LRU cache for brushes
- Theme pre-warming to avoid first-frame jank

### 4.4 Async Enumeration ✓
**Location:** `RedSalamander/FolderView.Enumeration.cpp`
- Background thread with stop tokens
- Generation-based staleness checks
- Proper cancellation handling

### 4.5 Thread Pool for Icons ✓
**Location:** `RedSalamander/FolderView.Icons.cpp`
- Uses `TrySubmitThreadpoolCallback`
- Bounded worker count
- Work-stealing atomic index

### 4.6 Move Semantics ✓
**Throughout codebase:**
```cpp
batch = std::move(_etwEventQueue);
_lines.push_back(std::move(line));
```

### 4.7 RAII Resource Management ✓
- `wil::unique_hicon`, `wil::unique_hbitmap`
- `wil::com_ptr` for COM objects
- `wil::scope_exit` for cleanup

---

## 5. IMPLEMENTATION PLAN

### Phase 1: Critical (Week 1)
| Task | File(s) | Est. Time |
|------|---------|-----------|
| Convert string params to `string_view` | ErrorOverlay, FolderWatcher, FileSystem7z | 2h |
| Add ASCII fast-path to `CaseInsensitiveHash` | DirectoryInfoCache.cpp | 1h |

### Phase 2: Moderate (Week 2)
| Task | File(s) | Est. Time |
|------|---------|-----------|
| Replace `_wcsicmp` with `CompareStringOrdinal` | FolderView, FileOperations | 2h |
| Add missing `reserve()` calls | FunctionBar.cpp, others | 1h |
| Audit vector pass-by-value | FileOperations | 1h |

### Phase 3: Polish (Week 3)
| Task | File(s) | Est. Time |
|------|---------|-----------|
| Convert concatenation to `std::format` | Various | 1h |
| Targeted `InvalidateRect` | FolderView | 2h |
| LRU for layout cache | ColorTextView | 2h |

---

## 6. BENCHMARKING RECOMMENDATIONS

Before and after measurements for:

1. **Folder enumeration time** - Navigate to `C:\Windows\System32` (3000+ files)
2. **Sorting latency** - Sort by name, then by date, measure transition time
3. **Memory allocation count** - Use ETW or sampling profiler during enumeration
4. **Scroll smoothness** - Profile frame times during continuous scrolling

### Profiling Tools
- Visual Studio Profiler (CPU Sampling)
- Windows Performance Analyzer (ETW traces)
- `start-etw-trace.ps1` / `stop-etw-trace.ps1` (project scripts)

---

## 7. RISK ASSESSMENT

| Change | Risk | Mitigation |
|--------|------|------------|
| `string_view` parameters | Dangling references if caller doesn't maintain lifetime | Review call sites, add comments |
| `CompareStringOrdinal` | Subtle sorting differences from `_wcsicmp` | Test with edge cases (accented characters) |
| Hash function changes | Cache invalidation, lookup misses | Maintain backward compatibility in transition |

---

## 8. APPENDIX: CODE PATTERNS

### Recommended: String Parameter Pattern
```cpp
// For read-only access (preferred)
void ProcessPath(std::wstring_view path);

// When storing a copy
void SetPath(std::wstring path) { _path = std::move(path); }

// When reference lifetime is guaranteed
void UpdatePath(const std::wstring& path);
```

### Recommended: Fast Case-Insensitive Compare
```cpp
inline int FastCompareNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    return CompareStringOrdinal(
        a.data(), static_cast<int>(a.size()),
        b.data(), static_cast<int>(b.size()),
        TRUE) - CSTR_EQUAL;
}
```

### Recommended: Container Pre-allocation
```cpp
std::vector<Item> items;
items.reserve(estimatedCount);  // Always reserve when size is known/estimable
for (const auto& source : sources)
    items.push_back(ProcessItem(source));
```

---

*Document generated by performance audit. Update status as optimizations are implemented.*




------------------------------------------------------------------------------------------------------

## DONE auto suggest combo
change the UI of the autosugest for folder to have a modern look like the one in settings dropdown
### Navigation polish
- Autosuggest for `nav:` with connection names (preview: protocol + host + user).
- Add a first-class entry point (menu/command palette): "Connections..."
- no back up to @conn the root is @conn/something/
- review nav to  add is needed a last '/' or '\' for proper navigation
- autosugest must work cross filesystem if I'm typing 'f' I want to see autosuggest 'ftp:' 'ftps:' and after '@' today sugest '@conn/' after suggest the connection from connection manager or if allready type
- autosuggest folders based on already type connection if it is working (if not silent failing to let the user continue typing)


------------------------------------------------------------------------------------------------------
# DONE Plan: Integrate NtQueryDirectoryFile into FileSystem Plugin

## Overview

Enhance the FileSystem plugin to use `NtQueryDirectoryFile` for local disk enumeration while maintaining `FindFirstFileExW` as a fallback for network paths and edge cases.

## Current State

**FileSystem Plugin** (`Plugins/FileSystem/FileSystem.cpp`):
- Uses `FindFirstFileExW`/`FindNextFileW` with `FIND_FIRST_EX_LARGE_FETCH`
- Progressive buffer growth (512KB → 64MB)
- Converts WIN32_FIND_DATAW to FileInfo structure
- Key functions: `StartEnumeration()`, `FetchNextEntry()`, `PopulateBuffer()`

**PoC LS3** (`PoC/ls3/ls3.cpp`):
- Uses `NtQueryDirectoryFile` with `FILE_BOTH_DIR_INFORMATION`
- 512KB VirtualAlloc buffer
- Direct kernel API calls (fewer user-mode transitions)
- Returns data in format nearly identical to FileInfo

## Why NtQueryDirectoryFile is Faster

1. **Direct kernel syscall** - Bypasses Win32 subsystem overhead
2. **Batch results** - Returns multiple entries per call vs. one at a time
3. **Native structures** - `FILE_BOTH_DIR_INFORMATION` maps directly to FileInfo
4. **No string normalization** - Kernel returns raw filenames

## Implementation Plan

### Phase 1: Detection Logic

Add path classification to determine enumeration strategy:

```cpp
enum class EnumerationStrategy { Win32, NtQuery };

EnumerationStrategy ClassifyPath(const std::wstring& path) noexcept
{
    // Network paths: \\server\share or \\?\UNC\ - use Win32
    if (path.starts_with(L"\\\\") && !path.starts_with(L"\\\\?\\"))
        return EnumerationStrategy::Win32;
    if (path.starts_with(L"\\\\?\\UNC\\"))
        return EnumerationStrategy::Win32;

    // WSL paths: \\wsl.localhost\ - use Win32
    if (path.find(L"\\wsl.localhost\\") != std::wstring::npos ||
        path.find(L"\\wsl$\\") != std::wstring::npos)
        return EnumerationStrategy::Win32;

    // Local drive letters or \\?\X:\ - use NtQuery
    return EnumerationStrategy::NtQuery;
}
```

### Phase 2: NT API Loading

Add lazy-loaded NT function pointers (once per process):

```cpp
namespace
{
    using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(
        HANDLE, HANDLE, PIO_APC_ROUTINE, PVOID, PIO_STATUS_BLOCK,
        PVOID, ULONG, FILE_INFORMATION_CLASS, BOOLEAN, PUNICODE_STRING, BOOLEAN
    );

    NtQueryDirectoryFile_t g_NtQueryDirectoryFile = nullptr;
    std::once_flag g_ntApiInitFlag;

    void InitNtApis() noexcept
    {
        HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
        if (ntdll)
        {
            g_NtQueryDirectoryFile = reinterpret_cast<NtQueryDirectoryFile_t>(
                GetProcAddress(ntdll, "NtQueryDirectoryFile"));
        }
    }
}
```

### Phase 3: New Enumeration Path

Add parallel implementation in FilesInformation:

**New members:**
```cpp
class FilesInformation
{
    // Existing Win32 state
    wil::unique_hfind _findHandle;
    WIN32_FIND_DATAW _pendingEntry;
    bool _hasPendingEntry;

    // New NT state
    wil::unique_handle _ntHandle;        // Directory handle for NtQueryDirectoryFile
    std::vector<std::byte> _ntBuffer;    // Dedicated buffer for NT results
    size_t _ntBufferOffset;              // Current position in NT buffer
    bool _ntRestartScan;                 // TRUE on first call
    EnumerationStrategy _strategy;
};
```

**New functions:**
```cpp
HRESULT StartEnumerationNt(FilesInformation& info, const std::wstring& path) noexcept;
HRESULT FetchNextEntryNt(FilesInformation& info, FileInfo** ppEntry, size_t* pEntrySize) noexcept;
HRESULT PopulateBufferNt(FilesInformation& info, ...) noexcept;
```

### Phase 4: Direct FileInfo Population

Key optimization: `FILE_BOTH_DIR_INFORMATION` is **binary identical** to `FileInfo`:

| FILE_BOTH_DIR_INFORMATION | FileInfo | Size |
|---------------------------|----------|------|
| ULONG NextEntryOffset | unsigned long NextEntryOffset | 4 |
| ULONG FileIndex | unsigned long FileIndex | 4 |
| LARGE_INTEGER CreationTime | __int64 CreationTime | 8 |
| LARGE_INTEGER LastAccessTime | __int64 LastAccessTime | 8 |
| LARGE_INTEGER LastWriteTime | __int64 LastWriteTime | 8 |
| LARGE_INTEGER ChangeTime | __int64 ChangeTime | 8 |
| LARGE_INTEGER EndOfFile | __int64 EndOfFile | 8 |
| LARGE_INTEGER AllocationSize | __int64 AllocationSize | 8 |
| ULONG FileAttributes | unsigned long FileAttributes | 4 |
| ULONG FileNameLength | unsigned long FileNameSize | 4 |
| ULONG EaSize | unsigned long EaSize | 4 |
| CCHAR ShortNameLength | char ShortNameSize | 1 |
| (1 byte padding) | (1 byte padding) | 1 |
| WCHAR ShortName[12] | wchar_t ShortName[12] | 24 |
| WCHAR FileName[1] | wchar_t FileName[1] | var |

**Implementation:** Direct `memcpy` of entire entries - no field-by-field copying needed. The NT buffer can be used directly or copied in bulk to the output buffer.

### Phase 5: Integration Points

Modify existing functions to branch on strategy:

```cpp
HRESULT FileSystem::StartEnumeration(FilesInformation& info, const std::wstring& path) noexcept
{
    std::call_once(g_ntApiInitFlag, InitNtApis);

    info._strategy = ClassifyPath(path);

    if (info._strategy == EnumerationStrategy::NtQuery && g_NtQueryDirectoryFile)
    {
        return StartEnumerationNt(info, path);
    }

    // Existing Win32 path
    return StartEnumerationWin32(info, path);
}
```

### Phase 6: Error Handling & Fallback

Add graceful fallback if NT path fails:

```cpp
HRESULT StartEnumerationNt(FilesInformation& info, const std::wstring& path) noexcept
{
    // Try NT path
    HRESULT hr = OpenDirectoryNt(path, info._ntHandle);
    if (FAILED(hr))
    {
        // Fallback to Win32 on any failure
        info._strategy = EnumerationStrategy::Win32;
        return StartEnumerationWin32(info, path);
    }
    // ...
}
```

## Files to Modify

| File | Changes |
|------|---------|
| `Plugins/FileSystem/FileSystem.h` | Add NT state to FilesInformation, enum, function declarations |
| `Plugins/FileSystem/FileSystem.cpp` | Add NT API loading, ClassifyPath, StartEnumerationNt, FetchNextEntryNt, PopulateBufferNt |

## Verification

1. **Build**: `.\build.ps1 -ProjectName FileSystem`
2. **Functional test**: Run RedSalamander, navigate to:
   - Local drive (C:\Windows\System32) - should use NtQuery
   - Network share (\\server\share) - should use Win32
   - WSL path (\\wsl.localhost\Ubuntu) - should use Win32
3. **Performance comparison**: Time enumeration of large directory (>10K files)
4. **Edge cases**:
   - Empty directories
   - Directories with 100K+ files
   - Paths > 260 characters
   - Reparse points (symlinks, junctions)

## Risks & Mitigations

| Risk | Mitigation |
|------|------------|
| NT API changes between Windows versions | Use documented FILE_INFORMATION_CLASS values; test on Win10/11 |
| Network paths behaving differently | Explicit path classification; always fallback to Win32 |
| Buffer format differences | Careful struct mapping; unit tests for field alignment |
| ntdll.dll not loadable (sandboxed?) | Null check g_NtQueryDirectoryFile; fallback to Win32 |

## Optional Future Enhancements

1. **Relative directory opens** (from LS3): Use `NtCreateFile` with relative `OBJECT_ATTRIBUTES` for recursive enumeration
2. **FILE_ID_BOTH_DIR_INFORMATION**: Return 64-bit file IDs for deduplication
3. **Async enumeration**: Use Event parameter for overlapped I/O



------------------------------------------------------------------------------------------------------
## DONE icon in menu
use sehoegoe fluent icon font to display icons in menu instead of using image resources
documentation here:
https://learn.microsoft.com/en-us/windows/apps/design/style/segoe-fluent-icons-font
https://github.com/microsoft/fluentui-system-icons/blob/main/icons_regular.md
https://github.com/microsoft/fluentui-system-icons/blob/main/icons_filled.md

------------------------------------------------------------------------------------------------------

## DONE Connection Manager (FTP / FTPS / SFTP / SCP / IMAP) - next steps / missing implementation

### Current milestone (already implemented)
- Spec: `Specs/Core/Core_ConnectionManager.md`
- Settings v7 `connections` model + schema: `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `Specs/SettingsStore.schema.json`
- Secrets stored in WinCred + optional Windows Hello gate: `RedSalamander/ConnectionSecrets.*`, `RedSalamander/WindowsHello.*`
- Themed dialog: `RedSalamander/ConnectionManagerDialog.*`, `RedSalamander/RedSalamander.rc`
- Host<->plugin API: `IHostConnections` in `Common/PlugInterfaces/Host.h` + implementation in `RedSalamander/HostServices.cpp`
- Navigation: `nav:<connectionName>` and `ftp:`/`sftp:`/`scp:` (no host) opens Connection Manager: `RedSalamander/FolderWindow.FileSystem.cpp`
- Plugin `/@conn/<id>` resolution (FileSystemCurl): `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp`

### Biggest missing piece (required): prompting + ephemeral credentials
- Today, password-based connections require `savePassword == true` (no prompt path yet).
- Needed behavior:
  - If `savePassword == false`, host prompts (themed) for password/passphrase at connect time.
  - Secret is kept in-memory only (per-session cache keyed by connectionId + secretKind), cleared on exit (and optionally after a short TTL).
- Proposed API addition (host service):
  - `HRESULT PromptForSecret(const ConnectionSecretRequest*, /*out*/ wchar_t** secret, /*out*/ uint32_t* secretChars)`
  - Or: add `PromptForSecret` + allow `GetConnectionSecret` to return cached secrets even when not persisted.
- FTP-specific: support a prompt when server rejects anonymous login or when user navigates to `ftp:` and selects a connection missing credentials.

### Dialog UX/validation polish
- Add explicit auth-mode selector (Anonymous / Password / SSH key) instead of inferring from fields.
- Inline validation (not modal MessageBox):
  - name required + unique (case-insensitive)
  - host required, port range, initialPath normalization
  - required fields by authMode
- Rename flow should update list immediately; consider adding "Duplicate" button.
- Better list: protocol icon, search box/filter, keyboard shortcuts (Del to remove, F2 rename, Enter connect).
- Add "Test connection" action:
  - host invokes a plugin callback to attempt handshake and returns a structured result (success/failure message).

### Plugin-provided right-pane fields (schema-driven, like Preferences)
- Add optional plugin interface to provide:
  - JSON schema + defaults for `ConnectionProfile.extra` (non-secret only)
  - validation hook on "Connect"
  - optional "build args" hook (if plugin needs derived fields)
- Host renders a generic editor for plugin fields (same system as Preferences plugin config editor).

### Security + robustness improvements
- Zeroize secrets after use where feasible.
- Windows Hello UX:
  - show availability state (supported / not configured / unavailable)
  - optional "unlock once for 10 minutes" session semantics to reduce repeated prompts
- Add "Change password/passphrase..." flow that updates WinCred without editing other fields.
- Improve error messages when WinCred read/write fails (still avoid logging secrets).

### Navigation polish
- Autosuggest for `nav:` with connection names (preview: protocol + host + user).
- Add a first-class entry point (menu/command palette): "Connections..."
- no back up to @conn the root is @conn/something/
- review nav to  add is needed a last '/' or '\' for proper navigation
- autosugest must work cross filesystem if I'm typing 'f' I want to see autosuggest 'ftp:' 'ftps:' and after '@' today sugest '@conn/' after suggest the connection from connection manager or if allready type
- autosuggest folders based on already type connection if it is working (if not silent failing to let the user continue typing)

### Suggested implementation order
- [ ] Add host prompt UI + `PromptForSecret` API + in-memory secret cache
- [ ] Update FileSystemCurl to call prompt when `savePassword == false`
- [ ] Update FileSystem to call prompt when network connection need authentication
- [ ] Improve dialog validation + auth-mode selector + rename/duplicate UX
- [ ] Add plugin schema provider for extra fields (reuse Preferences editor)
- [ ] Add "Test connection" callback + progress/cancel UI
- [ ] Add `nav:` autosuggest + a dedicated "Connections..." command

### Open questions to decide
- Should connections be global and filtered by `pluginId` (current), or per-plugin storage entirely?
- Do we want separate connection entries for `ftp` vs `ftps` (different pluginId) or a single profile with a "TLS mode" field?
- Windows Hello gating: prompt every time secrets are requested, or allow a temporary unlock window?



-------------------------------------------------------------------------------------------------------
# DONE review preference dialog box
- as plugin manager preferences must theme
- this a dialog with a list on the left for each category of preferences
- a pane on the right to edit the preferences
- same logic as plugin manager for toggle and option
- add in the schema all the display name and description for the settings to be able to display them in the preference dialog box
- share code with plugin managment dialog box when possible to avoid code duplication
for RedSalamander the section on the right are:
  - General : general application preferences
  - Panes : with a section for Left pane and a section for Right pane
  - Viewers : to edit extension and associated viewer plugin
  - Editors : to edit extension and associated editor plugin
  - Keyboard : edit all keyboard shortcut
  - Mouse : edit all mouse shortcut
  - Themes : to show / copy and edit custom themes
  - Plugins : list of installed plugins to enable/disable/configure
  - Advanced : advanced preferences for expert user only

- need a new custom control to edit the color with a coor picker or a text box to enter the color in hex format #RRGGBB or #AARRGGBB
-  for Theme there a combo to display all themes installed in the themes folder plus default system, light, dark, rainbow themes displayed with their name
  all value are displayed and editable according to their type (bool, int, string, enum, etc ...)
  There is a button to load a theme from a file to let edit ie
  There is a button to apply it temporarily to see the result before applying it permanently
  There is a button specific to Theme pane to save the new theme in a dedicated file

  when pressing OK all settings are saved and applied
  when pressing Cancel all settings are discarded
  when pressing Apply all settings are saved and applied but the dialog stay open
- for keyboard preferences the current dialog is not very user friendly we need to redesign it to have a better user experience
  propose a new design for this dialog to improve user experience
  need to be able to filter by command name, by current shortcut, by pane (left, right, both)
  need to be able to select a command and press a new shortcut to assign it directly
  need to be able to remove the current shortcut assigned to a command
  need to be able to reset all shortcuts to default value
  need to be able to export / import keyboard shortcuts configuration to share it between different installation of the application
- monitor setting are in the advanced pane with other advaced option


------------------------------------------------------------------------------------------------------

## DONE new pluginviewer
a newplugin in the Plugins folder this is a plugin to save in the Plugins folder to the destination folder of the main application
ViewerPE is a viewer for PE format for tthe file extension
- *.cpl
- *.dll
- *.drv
- *.exe
- *.ocx
- *.spl
- *.sys
- *.scr
it display the portable executable file format (PE) used in 32-bit and 64-bit versions of Windows operating systems. The PE format is a data structure that encapsulates the information necessary for the Windows OS loader to manage the wrapped executable code.

use pe-parse library to parse and analyze PE files.
example of pe-parse library: https://github.com/trailofbits/pe-parse/blob/master/examples/peaddrconv/main.cpp
project https://github.com/trailofbits/pe-parse/tree/master

### Features
- Display PE Header Information in directX with nice looking and easy to read format and nice animation.
- as an option to export PE Header Information to a text file.
- display all section headers with details such as name, virtual size, raw size, characteristics, and more.

------------------------------------------------------------------------------------------------------

## DONE new VLC pluginviewer
a newplugin in the Plugins folder this is a plugin to save in the Plugins folder to the destination folder of the main application
ViewerVLC is a viewer for video format media player files for the file extension
- *.avi
- *.mp4
- *.mkv
- *.mov
- *.wmv
- *.flv
- *.mpg
- and all format supported by VLC media player
it display the video file format used in various media players.
use libVLC library to parse and analyze video files.
documentation  https://videolan.videolan.me/vlc/group__libvlc.html
this is a DirectX plugin using libVLC to play video files inside the main application.

find VLC installation on the system to use its libraries to play video files.
example of libVLC library:  https://wiki.videolan.org/LibVLC_Tutorial_C/
project https://github.com/videolan/libvlc
if not present display a nice message to the user to install VLC media player with a clickable link to https://www.videolan.org/vlc/

### Features
- Play video files in directX with nice looking and easy to read format and nice animation.
- as an option to export video frames to image files.
- support for various video formats supported by VLC media player.
- support for subtitles and audio tracks.
- support for playback controls such as play, pause, stop, seek, and volume control.
- support for fullscreen mode and windowed mode.
- support for hardware acceleration for smooth playback.


-------------------------------------------------------------------------------------------------------
# DONE in plugin configuration dialog box
- display the comment associated with each configuration option under the option control in a italic font to help user to understand the option meaning
- true false value must be displayed as radio buttons instead of combobox to improve user experience
- radio button must be theme today text color is not correct neither readabable in dark theme



-------------------------------------------------------------------------------------------------------
# DONE unfocus pane focused item
today the unfocus pane the selection item are dimmed but still selected
keep that behavior
the focus item on the unfocus pane dont have background highlight but still have the focus rectangle around the item in a thinner way to show the item with focus in the unfocus pane
you have to dimm also this focus rectangle to show the item is in unfocus pane

when refocusing the pane the focus rectangle must be redrawn in normal way to show the item is now focused again


------------------------------------------------------------------------------------------------------
## DONE Review ViewerImgRaw plugin
- at the bottom it must be a real status bar
- display if you are displaying the jpg or the raw when there is both
- display the exif al:so for jpeg
- use the jpeg information to display the picture with the correct orientation
- use Ctrl + click to zoom in transient mode
- use clik to move arount the picture when zoomed
- when loading a raw file display a progress bar if the loading take more than 500 ms
- when loading a raw file display the jpeg preview if any while loading the raw data to have a quick preview
- display a progress bar when loading a jpeg if it take more than 500 ms
- move to next prev image with right left arrows (remove the need to press ctrl key)
- home end to go to first last image in the folder
- PgUp PgDn to go to previous next image in the folder
- use Ctrl+ arrow to move the image when zoomed
- use mouse wheel to zoom in / out centered on mouse position
- use Ctrl + mouse wheel to change brightness / contrast
- double click to fit to window
- double click with Ctrl to display at 100%
- add shortcut to rotate image clockwise / counterclockwise
- add shortcut to flip image horizontal / vertical
- add shortcut to reset orientation
- add shortcut to increase / decrease brightness
- add shortcut to increase / decrease contrast
- add shortcut to increase / decrease gamma
- add shortcut to switch between fit to window / 100%
- add shortcut to switch between color / grayscale
- add shortcut to switch between normal / negative
- review all shortcut to be coherent with ViewerText and other viewer plugins

-------------------------------------------------------------------------------------------------------
# DONE Review Settings and default behavior for TextViewer
why are we saving the default value for openWithFileSystemByExtension and openWithViewerByExtension in settings file ?
by default if nothing is specified in settings file we have to use the default value defined in code
we could save in settings file only the user defined value when different from default value to reduce settings file size and improve readability

on top of that TextViewer is the default vierwer for everything if something was not find in any other extension try to open TextViewer and guess if this is text, if guess fail display the file in HEX mode

-------------------------------------------------------------------------------------------------------
# DONE new plugin ViewerRaw
I want to build a new plugin ViewerRaw to show raw image file like .raw .cr2 .nef .arw etc ...
this plug in use LibRaw library from the vcpkg to read raw image file and display it in a viewer
the documentation of LibRaw is here: https://libraw.org/docs
the C++ API is here: https://libraw.org/docs/api/cpp_api
code sample of the API is there: https://github.com/LibRaw/LibRaw/tree/master/samples
this plugin is associated with this list of extensions:
.3fr, .ari, .arw, .bay, .braw, .crw, .cr2, .cr3, .cap, .data, .dcs, .dcr, .dng, .drf, .eip, .erf, .fff, .gpr, .iiq, .k25, .kdc, .mdc, .mef, .mos, .mrw, .nef, .nrw, .obm, .orf, .pef, .ptx, .pxn, .r3d, .raf, .raw, .rwl, .rw2, .rwz, .sr2, .srf, .srw, .tif, .x3f

this plugin:
- as the ViewerText is written in C++ fololowing AGENTS.md guidelines for application development
- implemented in DirectX 11 following other viewer plugins
- use WIL RAII wrappers for resources following wil-raii skill guidelines
- follow `Specs/Plugins/Plugins_ViewerPlugins.md`
- follow cpp-modern-style skill guidelines for C++23 usage
- follow the theme
- as a combobox as ViewerText to display the images from the selection/folder
- as the same shortcut (only the meaning full ones) than ViewerText

make a plan to implement this plugin by respecting AGENTS.md guidelines for application development

------------------------------------------------------------------------------------------------------
## DONE Review show folder history
menu open is not nice with random position when we are clicking on show folder history button
propose to have a fixed position for this menu
if string + gap and check mark are shorter than the pane size  display a fixed position bellow the button from right to left of the pane
if string + gap and check mark are bigger than the pane size display a fixed position bellow the button from left to right of the window
if string + gap and check mark are bigger than the window size display same fixed position than above plus middle ellipsis for the path with priolrity to first and last segment of the path

-------------------------------------------------------------------------------------------------------
# DONE When we are in incremental search mode need a graphical indicator
When we are in incremental search mode need a graphical indicator
Display in the status bar "search: the string being searched" when in incremental search mode
When exiting incremental search mode the status bar must remove this indicator
Add also a graphical indicator top right of the window with and semi transparent mode (something like a pen or any graphical nice indicator) to show we are in incremental search mode

-----------------------------------------------------------------------------------------------------
## DONE review behavior for connecting / disconnecting  drives
when a pane browse a network drive or any drive who disapear (ftp, s3, etc ...) and the drive is disconnected we must display an information message in the pane to inform the user the drive is disconnected

-------------------------------------------------------------------------------------------------------
# DONE review Navigation Bar
pressing enter in edit mode in the navigation need to add the focus in the pane at the end of the process

-------------------------------------------------------------------------------------------------------
# DONE review FolderView Backspace
Backspace on a empty folder view to go to the parent directory of the current folder displayed in the pane
Backspace must always work in folder to go up on directory

-------------------------------------------------------------------------------------------------------
# DONE review Alert Info architecture and implementation
review all alert info architecture and implementation in the application and plugins
4 types of messages
- modal in a pane
- modal in the application window
- modeless in a pane
- modeless in the application window
Plugin who want  to display an alert info must specify the type of alert info to display are not necessary in directX
All message implemetation must reside in RedSalamander main code and be accessible from plugins via the plugin API implementred in RedSalamander main code
to manage the modal modeless way the implementation could use a transparent window over the pane or application window to display the message
review all current implementation in the code and propose a plan to implement this feature by respecting AGENTS.md guidelines for application development, update `Specs/Plugins/Plugins_PluginAPI.md` to document the new API for plugins to display alert info messages


-----------------------------------------------------------------------------------------------------
## DONE in Navigation bar Backspace behavior review
Ctrl + backspace is not removing the word before the cursor in the navigation bar

-----------------------------------------------------------------------------------------------------
## DONE Error Message display the same way with
review all error message display in the application and plugins
propose to have a coherent way to display error message in the application and plugins
example: message in the pane displayed in a red box with an icon on the left and the text on the right
this code must be shared between the application and plugins to ensure a coherent display
this display must add an optional  cross top right with highlight to close the message
the display could have optionalbutton OK/Cancel Yes/No

review all error message display in the application and plugins
propose a plan to implement this feature

-----------------------------------------------------------------------------------------------------
## Review all mount file systems in the application
When we are at the root of a mount file system Backspace must go to the previous file system in the previous folder
for example
- user navigate to 7z:/ we must go back to C:\folder\ where we've mount C:\folder\archive.zip

------------------------------------------------------------------------------------------------------
## DONE Navigation bar active /inactive state display
review the navigation bar active / inactive state display
propose a better way to display the active / inactive state of the navigation bar
example: all color of the navigation bar change when active / inactive

the color must follow the theme (light / dark / rainbow)
propose a plan to implement this feature

-------------------------------------------------------------------------------------------------------
# DONE Global Menu update
the applicatiuon global menu must be updated to add new commands and organize them better
review all menu entries the  new organization of the menu entries
all menu enties must be link to command and display name in ressources for localization and the current shortcut if any
Top level menu are Left, Files, Edit, Commands, Plugins, View, Right, Help (right align)
- Left : menu for left pane
      - Change Drive
      - Go to >
          - Back (History Back)
          - Forward (History Forward)
          - Parent Directory (Up One Level)
          - Root Directory (Go to Root)
          - Path from Other Panel
          --------------------------------------
          - (Hotpaths list here)
          --------------------------------------
          - (History paths here)
      -------------------------
      - Brief
      - Detailed
      -------------------------
      - Sort By >
          - None
          - Name
          - Extension
          - Size
          - Time
          - Attributes
      -------------------------
      - Maximize/Restore (the pane)
      - Swap (panes)
      -------------------------
      - Filter...
      - Refresh
- Files : file operations (copy, move, delete, rename, create directory, etc ...)
      - Rename...
      - View
      - Alternate View
      - View With >
            - (shell extensions viewers here)
      - Edit
      - Edit With >
            - (shell extensions editors here)
      - Edit New File...
      - Copy...
      - Move/Rename...
      - Delete
      - Properties
      - Security
      -------------------------------------
      - Change Attributes...
      - Change Case...
      - Pack...
      - Unpack...
      - New >
          - Folder...
          ------------------------
          - (New files from Shell extensions here)
      -------------------------------------
      - Exit
- Edit : edit operations (Cut, Copy, Paste, Edit, Edit New, Change Attributes, Change Case, etc ...)
      - Cut
      - Copy
      - Paste
      - Paste Shortcut
      -----------------------------
      - Copy Path + Name as Text
      - Copy Name as Text
      - Copy Path as Text
      ------------------------------
      - Select...
      - Unselect...
      - Invert Selection
      - Select All
      - Unselect All
      - Restore Selection
      ------------------------------
      - Advanced >
          - Save Selection
          - Load Selection...
          ------------------------------
          - Select Same Extensions
          - Unselect Same Extensions
          ------------------------------
          - Select Same Names
          - Unselect Same Names
          ------------------------------
          - Hide Selected Names
          - Hide Unselected Names
          - Show Hidden Names
          ------------------------------
          - Go to Previous Selected Name
          - Go to Next Selected Name
- Commands : commands operations (compare, alternate view, space viewer, etc ...)
          - Create Directory...
          - Change Directory...
          - Compare Directories...
          - Calculate Occupied Space
          - Drive Information...
          - Find Files and Directories...
          - Make File List...
          - Go to Shortcut or Link Target
          ---------------------------------
          - List of Opened Files
          - Show Folders History
          ---------------------------------
          - Connect Network Drive...
          - Disconnect...
          - Shared Directories...
          ---------------------------------
          - Command Shell
          - Reread Associations
          ---------------------------------
          - User Menu >
              - (user menu items here)
          - Open File Explorer >
              - (all list of common folders from file explorer here)
- Plugins : plugin operations (list of plugins, plugin manager, etc ...)
      - Plugin Manager...
      ---------------------------------
      - (list of installed plugins here to enable/disable/configure)
- View : application options (settings, themes, etc ...)
      - Theme >
          - High Contrast (System)
          -----------------------------
          - System
          - Light
          - Dark
          - Rainbow
          - High Contrast (App)
          -----------------------------
          - (user defined and Themes folder here)
      ---------------------------------
      - Toggle Fullscreen
      - Pane >
          - [] Show Hidden Files
          - [] Show System Files
          - [] Show File Extensions
          - [] Show Thumbnails
          - [] Show Preview Pane
          - [] Show Filter Bar
          ------------------------------
          - [] Show Navigation Bar (Left)
          - [] Show Navigation Bar (Right)
          ------------------------------
          - [] Show Status Bar (Left)
          - [] Show Status Bar (Right)
      - [] Show Function Bar
      - [] Show Menu
      ------------------------------
      - Preferences...

- Right : menu for right pane same as left pane
- Help : help menu (documentation, about, etc ...)
      - Display Shortcuts...
      ------------------------------
      - About...

review all menu entries to ensure they are in the right menu
propose new menu entries to add to the menu if something is missing
add between [] the command in front of each menu entry to link it to the command system
list all menu entries where the command is missing
list all command without mernu entry
review the specs create a new `Specs/UI/UI_CommandMenuKeyboard.md` to document the new menu structure and merge with `Specs/UI/UI_KeyboardManagement.md` in it
you could remove `Specs/UI/UI_KeyboardManagement.md` if all is in `Specs/UI/UI_CommandMenuKeyboard.md`
propose a plan to implement this Spec by respecting AGENTS.md guidelines for application development
I want just a new spec no code modification yet


-----------------------------------------------------------------------------------------------------
## DONE File System Dummy plugin review
Need to have better fake file genration to test the application with large file system
- Text file must display valid random text,  always the same text for the same file using a hash on the file name and metadata generated
- csv also
- json also
- xml also
- jpg must provide a random picture in proper jpg format
- png must provide a random picture in proper png format
- etc ...
same for all file format we could generate in a proper format to test the viewer

review the usage of max children per directory and max depth of the tree to be sure we can generate large file system for testing
not all folder must have the same number of children the max is the maximum possible children per folder generate small and large folder randomly
same for depth of the tree not all branch must have the same depth randomly generate small and large

missing the maxdepth of the tree in the settings dialog box and in schema and in the settings loading/saving
default value for max depth of the tree must be 10

-----------------------------------------------------------------------------------------------------

## DONE Review UTF8 vs UTF16 usage in the code

who is using GetConfigurationSchema ?
- plugin manager mainly to get the configuration schema of the plugin
- today schema is UTF16 encoded
- proposing to change GetConfigurationSchema to return std::string directly encoded in UTF8
- review all the std::string utf8 = Utf8FromUtf16(schemaJson); lines in the code
- Why returning wide string to after transform it in UTF8 ?
- proposing to return directly std::string from GetConfigurationSchema to avoid useless conversion
- other example:

```cpp
// rootPath: plugin-defined mount context (UTF-16, NUL-terminated).
// optionsJson: optional JSON/JSON5 payload (UTF-16, NUL-terminated) for per-instance options (passwords, initial path, etc.).
virtual HRESULT STDMETHODCALLTYPE Initialize(const wchar_t* rootPath, const wchar_t* optionsJson) noexcept = 0;
```

propose to change optionsJson to be UTF8 std::string directly

review all json usage in the code ro peopose to store and manipulate UTF8 strings only
make a review an propose changes to have UTF8 whne this is pertinant like JSON and GetConfigurationSchema



------------------------------------------------------------------------------------------------------
## DONE Rename Dialog review
review the rename dialog to improve user experience
- when renaming a file select only the name part without the extension
- the text in editbox must be vertical align


------------------------------------------------------------------------------------------------------
## DONE Review ViewerText status bar
review the status bar in the text viewer
propose to add more information in the status bar
when you know the max numbers of lines in the file loaded in the text viewer display it in the status bar
display 'unknown' when the number of lines is not known
example:
- FS: File System Detected: UTF-8 (guess) Active: UTF-8 Size: 762 KB Lines: 1-47 of 1234
- FS: File System Detected: UTF-8 (guess) Active: UTF-8 Size: 762 KB Lines: 1-47 of unknown


------------------------------------------------------------------------------------------------------
## DONE Function bar
command bar application window
- add a command bar at the bottom of the application window to be able to see the shortcut functions keys configured and the current modificators
- this command bar has 12 zone of equal size:
  - each zone correspond to a shortcut configured in the settings
  - the bar is 24 Dips height, screen width
  - each zone as the same size and display the name of the shortcut configured and a small icon displaying a round rectangle with the funtion key in the rectagle (example [F1] or [F2])
 ⌈───────────────────────────────────────────────────────────────────────────────────── ...    ────────────────────── ⌉
 |                                                                                                                    |
 | [F1] something       [F2] Command         [F3] other Comand    [F4] else             ...                           |
 |                                                                                               [keys modificators]  |
 └────────────────────┴────────────────────┴────────────────────┴────────────────────┴─ ...    ───────────────────────┘
  - when the user press the key modificator (CTRL, SHIFT, ALT and their compinaisons) the display change to show current shorcut
  - the key modifcator (Ctrl | Shift | Ctrl + Shift ... ) zone is bottom right align and optional is there is not enough space we have to hide it
  keys the corresponding zone is highlighted
  - is the windows is too small to display all the indicators the display is truncated
  Default shortcut name for direct and moficators (⊘ is when there is no shortcut)
  - F1  : NONE: ⊘             | CONTROL: drive information | ALT: open left drive menu         | SHIFT: ⊘                  | CONTROL+SHIFT: ⊘                        | ALT+SHIFT: ⊘
  - F2  : NONE: Rename         | CONTROL: change attributes | ALT: open right drive menu        | SHIFT: ⊘                  | CONTROL+SHIFT: save selection           | ALT+SHIFT: ⊘
  - F3  : NONE: View           | CONTROL: sort by name      | ALT: alternate view               | SHIFT: open current folder | CONTROL+SHIFT: view width               | ALT+SHIFT: ⊘
  - F4  : NONE: Edit           | CONTROL: sort by extension | ALT: exit                         | SHIFT: edit new            | CONTROL+SHIFT: edit width               | ALT+SHIFT: ⊘
  - F5  : NONE: Copy           | CONTROL: sort by time      | ALT: pack                         | SHIFT: ⊘                  | CONTROL+SHIFT: save selection            | ALT+SHIFT: ⊘
  - F6  : NONE: Move           | CONTROL: sort by size      | ALT: unpack                       | SHIFT: ⊘                  | CONTROL+SHIFT: load selection            | ALT+SHIFT: ⊘
  - F7  : NONE: Make directory | CONTROL: change case       | ALT: find                         | SHIFT: change directory    | CONTROL+SHIFT: ⊘                        | ALT+SHIFT: ⊘
  - F8  : NONE: Delete         | CONTROL: ⊘                | ALT: ⊘                           | SHIFT: permanent delete    | CONTROL+SHIFT: ⊘                        | ALT+SHIFT: ⊘
  - F9  : NONE: user menu      | CONTROL: refresh           | ALT: unpack                       | SHIFT: hot paths           | CONTROL+SHIFT: shares                    | ALT+SHIFT: ⊘
  - F10 : NONE: menu           | CONTROL: compare           | ALT: Space View                   | SHIFT: context menu        | CONTROL+SHIFT: ⊘                         | ALT+SHIFT: context menu for current directory
  - F11 : NONE: Connect        | CONTROL: zoom panel        | ALT: list of opened files         | SHIFT: ⊘                  | CONTROL+SHIFT: full screen                | ALT+SHIFT: ⊘
  - F12 : NONE: Disconnect     | CONTROL: filter            | ALT: list of working directories  | SHIFT: ⊘                  | CONTROL+SHIFT: ⊘                         | ALT+SHIFT: ⊘
  all function are overidable in the settings

there is other shortcuts not for function bar but for FolderView must be configurable and use the same logic
    - VK_BACK   : NONE: up one directory
    - VK_TAB    : NONE: swith pane focus
    - VK_RETURN : NONE: execute / open                    | CONTROL: brings filename to command line    | ALT: open properties dialog box | SHIFT:                                    | CONTROL+SHIFT: brings filename to command line      | CONTROL+ALT: ⊘
    - VK_SPACE  : NONE: select + calculate directory size | CONTROL: brings current dir to command line | ALT: window menu                | SHIFT: quick search                       | CONTROL+SHIFT: brings current dir to command line   | CONTROL+ALT: ⊘
    - VK_PRIOR  : NONE: page up                           | CONTROL: go up one folder                   | ALT: ⊘                         | SHIFT: select + page up                   | CONTROL+SHIFT: ⊘                                    | CONTROL+ALT: ⊘
    - VK_NEXT   : NONE: page down                         | CONTROL: do down if on a folder             | ALT: ⊘                         | SHIFT: select + page down                 | CONTROL+SHIFT: ⊘                                    | CONTROL+ALT: ⊘
    - VK_END    : NONE: end                               | CONTROL: last file                          | ALT: ⊘                         | SHIFT: select + end                       | CONTROL+SHIFT: select + last file                    | CONTROL+ALT: ⊘
    - VK_HOME   : NONE: home                              | CONTROL: first file                         | ALT: ⊘                         | SHIFT: select + home                      | CONTROL+SHIFT: select + first file                   | CONTROL+ALT: ⊘
    - VK_LEFT   : NONE: left                              | CONTROL: scroll left, keep focus item       | ALT: history: backward          | SHIFT: select + left                      | CONTROL+SHIFT: inverse pane and keep focus item      | CONTROL+ALT: ⊘
    - VK_UP     : NONE: up                                | CONTROL: scroll up, keep focus item         | ALT: up to selected item        | SHIFT: select + up                        | CONTROL+SHIFT: ⊘                                    | CONTROL+ALT: ⊘
    - VK_RIGHT  : NONE: right                             | CONTROL: scroll right, keep focust item     | ALT: history: forward           | SHIFT: select + right                     | CONTROL+SHIFT: inverse pane and keep focus item      | CONTROL+ALT: ⊘
    - VK_DOWN   : NONE: down                              | CONTROL: scroll down, keep focus item       | ALT: down to selected item      | SHIFT: select + down                      | CONTROL+SHIFT: ⊘                                    | CONTROL+ALT: ⊘
    - VK_INSERT : NONE: select/unselect                   | CONTROL: clipboard copy                     | ALT: copy path + name as text   | SHIFT: clipboard paste                    | CONTROL+SHIFT: copy path + file name                 | CONTROL+ALT: copy path as text | ALT+SHIFT: copy name as text
    - VK_DELETE : NONE: move to recycle bin               | CONTROL: ⊘                                 | ALT: ⊘                         | SHIFT: permanent delete with a validation | CONTROL+SHIFT: permanent delete with a validation    | CONTROL+ALT: ⊘
    - '0'       : NONE: quick search/type in command line | CONTROL: hot paths 0                        | ALT: panel mode                 | SHIFT: hot paths shift 0                  | CONTROL_SHIFT: define hot path 0                     | CONTROL+ALT: ⊘
    - '1'       : NONE: quick search/type in command line | CONTROL: hot paths 1                        | ALT: panel mode                 | SHIFT: hot paths shift 1                  | CONTROL_SHIFT: define hot path 1                     | CONTROL+ALT: ⊘
    - '2'       : NONE: quick search/type in command line | CONTROL: hot paths 2                        | ALT: panel mode                 | SHIFT: hot paths shift 2                  | CONTROL_SHIFT: define hot path 2                     | CONTROL+ALT: ⊘
    - '3'       : NONE: quick search/type in command line | CONTROL: hot paths 3                        | ALT: panel mode                 | SHIFT: hot paths shift 3                  | CONTROL_SHIFT: define hot path 3                     | CONTROL+ALT: ⊘
    - '4'       : NONE: quick search/type in command line | CONTROL: hot paths 4                        | ALT: panel mode                 | SHIFT: hot paths shift 4                  | CONTROL_SHIFT: define hot path 4                     | CONTROL+ALT: ⊘
    - '5'       : NONE: quick search/type in command line | CONTROL: hot paths 5                        | ALT: panel mode                 | SHIFT: hot paths shift 5                  | CONTROL_SHIFT: define hot path 5                     | CONTROL+ALT: ⊘
    - '6'       : NONE: quick search/type in command line | CONTROL: hot paths 6                        | ALT: panel mode                 | SHIFT: hot paths shift 6                  | CONTROL_SHIFT: define hot path 6                     | CONTROL+ALT: ⊘
    - '7'       : NONE: quick search/type in command line | CONTROL: hot paths 7                        | ALT: panel mode                 | SHIFT: hot paths shift 7                  | CONTROL_SHIFT: define hot path 7                     | CONTROL+ALT: ⊘
    - '8'       : NONE: quick search/type in command line | CONTROL: hot paths 8                        | ALT: panel mode                 | SHIFT: hot paths shift 8                  | CONTROL_SHIFT: define hot path 8                     | CONTROL+ALT: ⊘
    - '9'       : NONE: quick search/type in command line | CONTROL: hot paths 9                        | ALT: panel mode                 | SHIFT: hot paths shift 9                  | CONTROL_SHIFT: define hot path 9                     | CONTROL+ALT: ⊘

all command as a text equivalent in the form cmd/something with a display name in ressources for localization each command as an unique id like today

these shortcuts must be configurable in the settings
there is a menu entry 'Preferences...' in the main menu 'File' before 'Exit' with a separator to open the settings dialog
the settings dialog contain a tab 'Shortcuts' to configure all the shortcuts
this shortcut page contain sections:
- Function bar shortcuts
- Folder view shortcuts
there is a combo box to select the key (F1 to F12 for function bar, and all keys for folder view)
there is check box to select the modificators (CTRL, ALT, SHIFT)
there is a text box to enter the command name
example:
| Command Name                        | Key       | CTRL | ALT | SHIFT |
|-------------------------------------|-----------|------|-----|-------|
| Drive Information                   | F1        |      |     |       |
| Rename                              | F2        |  X   |     |       |
| Alternate View                      | F3        |      |  X  |       |
| Edit New                            | F4        |      |     |   X   |
| ...                                 | ...       |      |     |       |

if there is conflict in the shortcut configuration display a warning icon near the conflicting entry with a tooltip explaining the conflict
example: `Conflict with command 'Rename' (Ctrl + F2)`
there is button restore default to restore default shortcut configuration

the settings are restored at application startup

review all current shortcut review difference and show all differences you could find
prepare a plan for  this implementation by reviewing settings specs and code to be sure of a working solution with evolution in mind
prepapre a plan to implement this feature by repecting AGENTS.md guidelines for application development
show me the plan before starting the implementation


------------------------------------------------------------------------------------------------------
## DONE loading in viewer text
not respecting the rainbow theme for the loader
still a white window when creating the text viewer
review all window created if the theme is dark all window must be created with a default dark background the same for light theme
propose a plan to implement this feature



------------------------------------------------------------------------------------------------------
## DONE Pane Status bar
do be better viewing of the current pane the status bar has a line (2dips) on top with a color to show the focus state of the pane
in rainbow mode this line change color each time

------------------------------------------------------------------------------------------------------
## DONE Navigation bar
auto suggest of next path must work in edit for all plugins
example:
- user types `C:\Users` to navigate to file plugin
- user types `\\aComputer\Users` to navigate to file plugin
- user types `s3:mybucket/folder` to navigate to s3 plugin
- user types `ftp:/home/user` to navigate to ftp plugin
- user types `fk:/some/path` to navigate to fk plugin
- user types `7z:/archive/folder/archive.zip` to navigate to 7z plugin
  little bit more complex
  type 7z: start to suggest folders in filesystem and all possible file in each folder corresponding to possible extension in the settings
  `7z:/archive/folder/archive.zip|` when typing after the zip it must suggest folders inside the zip file

------------------------------------------------------------------------------------------------------
## DONE CreateDirectory valid chartacters review
review the valid characters for CreateDirectory in the file system plugins
invalid character in folder name must be rejected by CreateDirectory call
\ / : * ? " < > |
are invalid characters in folder name on Windows file system
display a message in the create folder dialog if the user try to use invalid characters
"A folder name xan't contain any of the following characters: \ / : * ? " < > |"



------------------------------------------------------------------------------------------------------

## DONE Viewer Space
- when scanning display the number of files and folder explore and the total size computed in short and in Bytes)
(right align) procesing: the current folder without path
example:

```text
Scaning C\aFolder\anotherFolder  (right align) procesing: thisfolder
1234 items (56 folders, 1178 files),
1.23 GB (1,234,567,890 bytes)
```

- the marquee when scanning is not nice enoufg improve it to be more user friendly

- add a cancel button to stop the scanning process

- when scanning display the current folder being scanned

- double click on other space start the explore this part to add the information

- when drill down if other take too much area launch the scan to display it in detail

- Other display today is
```text
Other(345 items)
123 GB
```
must be
```text
Other item (345 items)
45 folders, 300 files
123 GB
```


- the button to go up one level is not very user friendly
  propose to change it to a more user friendly icon / bigger with an hover highlight

- even if we start the scan from a folder we could go up and continue the scan from this new root folder
  so the up button must be active all the time during the scan and after the scan is finished

- animation is still not very smooth improve it

- still a lot of lag in the complete process (could not move the window around) when processing large folder with a lot of files
  improve the performance of the scanning process

- are we doing the graphical processing in the main thread ?
  if yes move it to a dedicated thread to avoid blocking the main thread

- at the beginning the window is completly empty (without respect from the them) and white for a long time
  need the first display to start faster (I've got the feeling there is still too much small allocation before the first display)

- when scanning a folder having a small animùation near the name to say this part of the tree is not completed yet (like a small rotating icon)
  to show the user we are still working on it

- when there is multiple square in one quare we need to display one line of text to be readable to have the name of the folder/file. Today the text is cut in the middle
  propose to display only one line of text per square even if there is multiple folder/file in it
  and display the number of items in this square when there is multiple items
  example: `3 items` or `5 folders` or `2 files` or `1 folder, 3 files` etc ...

- the tooltips is flickering when the mouse move a bit
  improve the tooltip display to avoid flickering

review all that and propose a plan to implement the changes by respecting AGENTS.md guidelines
check current spec propose modifications if needed





------------------------------------------------------------------------------------------------------
## DONE Review FileSystem Path Display and Navigation Behavior

we are passing path with the short name
in         const HRESULT enumHr = fileSystem->ReadDirectoryInfo(item.path.c_str(), filesInformation.put());
this is not expecting
- fk: start at /
- same for 7Z start at / as root
you don't need to add the shortid in
review globally Navigation and plugin
the shortId is in the ViewerOpenContext
the display in navigation don't need to display the shortID
today in FileSystemDummy in Navigation we are displaying 'fk:' this is not exepected
the filesytemId must be displayed through and Icon in the hamburger drive  menu icon
you need to lreview Navigation and plugin and implementation for doing that
the expected beahavior :
- on a win32 filesystem, drive menu display the option in the path I see 'c: > afolder > anotherFolder' when I'm editing I see c:\afolder\anotherFolder
- in fk, drive menu display a icon, the menu display (File System Dummy' in disabled state) the display is '/ > aFolder > anotherFolder' when I'm editing i see 'fk:/aFolder/anotherFolder'
- in 7z, in fk, drive menu display a zip icon, the menu display (7Z FileSytem' in disabled state, under the path of the zip we a browsing) the display is '/ > aFolder > anotherFolder' when I'm editing i see '7z:/aFolder/anotherFolder'
this way I know the file system used display is homogenous and edit could continue to navigate in the file system with auktosuggestion or navigate to another file system by start typing the shortId and ':' an a path in this context with autosuggest


example
I'm on a drive c:\aFolder I see a zip double click on it I'm navigating to 7z filesytem and display '/' the navigation bar display 7z scope for the file system

the navivation edit must be smarter is I'm on filesysystem

- I'm on a regular file system I start typing 7z:./a.zip this is to mount 7z with this zip (autosuggest must working to suggest the file in this context)
- same for :
 - 7z:a.zip
 - 7z:c:/afolder/a.zip
- 7z:\\acomputer\afolder\a.zip
- 7z:a.zip

for fk: today there is no mount context so
fk:/aFolder navigate in fk file system to aFolder folder

## Navigation Menu Cross-Plugin Navigation Support
review the INavigationMenuCallback interface to support cross-plugin navigation requests
propose the changes needed to support this feature
review the current spec and propose modifications if needed
check AGENTS.md guidelines for plugin development to ensure the proposal respect them

when the user start to type in the navigation bar we must be able to navigate to another plugin
example:
- user types `C:\Users` to navigate to file plugin
- user types `\\aComputer\Users` to navigate to file plugin
- user types `s3:mybucket/folder` to navigate to s3 plugin
- user types `ftp:/home/user` to navigate to ftp plugin
- user types `7z:/archive/folder/archive.zip` to navigate to 7z plugin
- user types `fk:/some/path` to navigate to fk plugin
- user clicks on navigation menu item to navigate to another plugin
- user clicks on navigation menu item to navigate to the same plugin but another path

on the edit when start typing we must be able to detect the plugin to use based on the path typed by the user and autosuggest it
review the current spec propose modifications if needed
exmaple of path to detect the plugin to use:
- start by a letter propose all plugin containing this letter display short Id and display name for the plugin in autosuggest list
- a letter and `:` we are on file plugin
- a plugin shortid + `:` we are on this plugin
example:
- `file:\\aComputer\Users` to navigate to file plugin
- `s3:mybucket/folder` to navigate to s3 plugin
- `ftp:/home/user` to navigate to ftp plugin
- `7z:/archive/folder/archive.zip` to navigate to 7z plugin
- `fk:/some/path` to navigate to fk plugin

the autrosuggest of the folder when typing must use plugin interface to list directory
No direct call to filesystem from the host application
review the current spec propose modifications if needed

by repecting AGENTS.md guidelines for plugin development

review the current spec propose modifications if needed

and propose a plan to implement this feature


------------------------------------------------------------------------------------------------------

## DONE  Review INavigationMenuCallback to support cross-plugin navigation requests
change in the INavigationMenuCallback to support cross-plugin navigation requests
### INavigationMenuCallback (optional)

```cpp
// Host callback for plugin-driven navigation requests.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at registration time and must be passed back verbatim by the plugin.
interface __declspec(novtable) INavigationMenuCallback
{
    virtual HRESULT STDMETHODCALLTYPE RequestNavigate(
        const wchar_t* path,
        const wchar_t* pluginNavigationPath, // optional for cross-plugin navigation
        void* cookie
    ) noexcept = 0;
};
```

**Callback behavior:**
- The host calls `SetCallback(hostCallback, cookie)` when `INavigationMenu` is available.
- The host calls `SetCallback(nullptr, nullptr)` when switching/unloading the active file system.
- Plugins can call `RequestNavigate(path, nullptr, cookie)` (typically from `ExecuteMenuCommand`) to request navigation.
    - `path` is a **plugin path** for the active file system (no `<shortId>:` prefix).
- Plugins can call `RequestNavigate(path, pluginNavigationPath, cookie)` to request navigation to another plugin.
    - `pluginNavigationPath` is a full navigation path including the `<shortId>:` prefix folkowed by `:` and the plugin initilialization path. For exmaple of pluginNavigationPath: `s3:mybucket/folder` or `ftp:/home/user` or `file:C:\Users`. or  `fk:/some/path` or `7z:/archive.7z/folder/archive.zip`, etc ....
- The plugin MUST pass back the `cookie` it received in `SetCallback` unchanged.
- The plugin MUST NOT call the callback after the host clears it via `SetCallback(nullptr, nullptr)`.

**NavigationView behavior:**
- Order is preserved exactly as provided by the plugin.
- `NAV_MENU_ITEM_FLAG_SEPARATOR` inserts a separator (label/path ignored).
- `NAV_MENU_ITEM_FLAG_HEADER` renders a disabled header row.
- `path` is interpreted as a **plugin path** for the active `IFileSystem` (no `<shortId>:` prefix).
  - For non-`file` plugins, the host formats display/history paths as `<path> (<pluginNavigationPath>)` with the shortId in pluginNavigationPath.
  - For `file`, the host uses Windows absolute paths (e.g., `C:\...`).
- If both `path` and `commandId` are provided, navigation takes precedence.
- If `path` is empty and `commandId != 0`, the host calls `ExecuteMenuCommand(commandId)`.
- The host assigns temporary Win32 menu item IDs for actionable rows (a reserved internal range in `NavigationView`); `commandId` is plugin-defined and is not treated as a Win32 `WM_COMMAND` identifier.
- Plugins SHOULD keep `commandId` values stable and unique within the returned list for predictable command routing/debugging.
- If `iconPath` is provided, the host uses `IconCache::QuerySysIconIndexForPath(iconPath, ...)`.
  If `iconPath` is empty and `path` is provided, the host uses `path` to resolve the icon.

**Ownership / lifetime:**
- The plugin owns the returned array and strings; they remain valid until the next call to the same method or until the COM object is released.




------------------------------------------------------------------------------------------------------
## DONE Review GetProcAddress usage in the code
propose to review all GetProcAddress usage in the code to ensure proper error handling
auto createFunc = reinterpret_cast<decltype(&RedSalamanderCreate)>(GetProcAddress(plugin.get(), "RedSalamanderCreate"));
THROW_LAST_ERROR_IF(!createFunc);


------------------------------------------------------------------------------------------------------
## DONE Review all LoadLibraryW usage in the code
propose to use LoadLibrary instead of LoadLibraryW everywhere in the code
review erro handling after LoadLibrary calls to ensure proper error handling
example:
```cpp
const std::filesystem::path pluginPath = exeDir / L"FileSystem.dll";
wil::unique_hmodule plugin(LoadLibrary(pluginPath.c_str()));
if (! plugin)
{
    const DWORD lastError = Debug::ErrorWithLastError(L"LoadLibrary '{}' failed.", pluginPath.wstring());
    return HRESULT_FROM_WIN32(lastError);
}
```


------------------------------------------------------------------------------------------------------
## DONE File System Interface review

The current IFileSystemFileIO interface is not coherent with the other file system interfaces
propose the following changes to harmonize the interfaces
- rename IFileSystemFileIO to IFileSystemIO
- move GetItemAttributes to the new IFileSystemIO a nd rename it as GetAttributes with the same logic for path
review the code using that an adapt

---------------------------------------------------------------------------
## DONE Text/Hex viewer using DirectX for display

The current text/hex viewer plugins are using GDI+ to display the content of the file loaded in the buffer.
you miss a big part of the plan

- all the display is in DirectX to be faster as possible
- the display must manage the caret position and selection in the buffer loaded only
- Selected text must be copyable to clipboard
- scrolling must be smooth and fast
- the directX implementation must be optimized to display as fast as possible the content of the buffer
- the directX implementation must folow the them and the dpi settings of the main application
- the text viewer must implemented dpi change on the fly
- the hex viewer must implemented dpi change on the fly
- the text viewer must manage large file without loading everything in memory
- the hex viewer must manage large file without loading everything in memory
- line numbers must be displayed if enabled in options
- beware of the scrolling managment to ensure they are coherent with the file

in summary the purpose is to have a reader faster than light to display file of any size
review the specs folder the AGENTS.md file for more details
propose a plan to implement this in the current text/hex viewer plugins
ask any question if something is not clear

# Plugin Interface Specification

## Overview 
The Plugin Interface enables RedSalamander to support multiple file system implementations through a COM-based plugin architecture. This interface abstracts file system operations, allowing seamless integration of local file systems, network shares, virtual file systems, or cloud storage without modifying the core application.

**Key Features:**
- COM-based plugin architecture for binary compatibility
- Single-call directory enumeration with buffer-based results
- File operations (copy/move/delete/rename) with batch support, progress callbacks, and cancellation
- Optional search interface for plugin-accelerated search
- Cross-compiler compatible (no STL in public interfaces)
- Extensible for future file system types

## Factory Method

Plugins are loaded from dynamic libraries (DLLs on Windows) and instantiated via a factory function. The plugin DLL must export a single entry point that creates COM interface instances.

### Factory Function Signature

```cpp
extern "C" 
{
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreate(
        REFIID riid,
        const FactoryOptions* factoryOptions,
        IHost* host,
        void** result
    );
}
```

### Optional: Multi-Plugin DLL Support

A single DLL MAY implement **multiple logical plugins** for the same interface type (e.g. one DLL that exposes `ftp`, `sftp`, and `scp` as separate file systems).

To do so, the DLL exports two additional (optional) entry points:

```cpp
extern "C"
{
    // Returns an array of PluginMetaData entries implemented by the DLL for the requested interface type.
    // The array and all strings are owned by the DLL and remain valid until the DLL is unloaded.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderEnumeratePlugins(
        REFIID riid,
        const PluginMetaData** metaData,
        unsigned int* count
    );

    // Creates a specific plugin instance identified by pluginId (metaData[i].id).
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreateEx(
        REFIID riid,
        const FactoryOptions* factoryOptions,
        IHost* host,
        const wchar_t* pluginId,
        void** result
    );
}
```

**Host behavior:**
- If `RedSalamanderEnumeratePlugins` is present, the host calls it during discovery and registers one plugin entry per returned `PluginMetaData` record.
- When instantiating a plugin entry originating from enumeration, the host calls `RedSalamanderCreateEx` with `pluginId == metaData[i].id`.
- If the optional exports are missing, the host falls back to `RedSalamanderCreate`.

**Plugin behavior:**
- `RedSalamanderEnumeratePlugins` MUST return stable metadata pointers for the lifetime of the loaded DLL.
- `RedSalamanderCreateEx` MUST return `HRESULT_FROM_WIN32(ERROR_NOT_FOUND)` (or `E_INVALIDARG`) for unknown `pluginId` values.

**Parameters:**
- `riid`: Interface ID to create (**only** `IID_IFileSystem` is creatable via `RedSalamanderCreate`; see UUID below)
- `factoryOptions`: Optional configuration (debug level, etc.)
- `host`: Host services object (caller-owned; remains valid for the lifetime of the created plugin instance)
- `result`: [out] Pointer to created interface instance

**Return Value:**
- `S_OK`: Success, `result` contains valid interface pointer
- `E_NOINTERFACE`: Requested interface not supported (including `IFilesInformation` which is not creatable directly)
- `E_POINTER`: `result` is `nullptr`
- `E_OUTOFMEMORY`: Allocation failed

**Interface discovery:**
- The factory only creates `IFileSystem`.
- The host obtains `IFileSystemSearch` via `QueryInterface` on the `IFileSystem` instance.
- The host may query for `INavigationMenu` and `IDriveInfo` for NavigationView integration; plugins may return `E_NOINTERFACE`.
- If a plugin does not support search, `QueryInterface` must return `E_NOINTERFACE`.
- File operations are part of `IFileSystem` and return `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` when not supported.

### FactoryOptions Structure

```cpp
enum DebugLevel : uint32_t
{
    DEBUG_LEVEL_NONE = 0,
    DEBUG_LEVEL_ERROR = 1,
    DEBUG_LEVEL_WARNING = 2,
    DEBUG_LEVEL_INFORMATION = 3,
} DebugLevel;

typedef struct FactoryOptions
{
    DebugLevel debugLevel;  // Diagnostic output verbosity
} FactoryOptions;
```

**Note:** The `FactoryOptions` pointer may be `nullptr` for default behavior. Debug level only affects diagnostic output and does not impact functionality.

**Current reference plugin behavior (FileSystem.dll):**
- The shipped/reference `Plugins/FileSystem/Factory.cpp` currently ignores `factoryOptions` (parameter is unused), but the parameter remains part of the ABI for future plugins.

## Plugin Interfaces

### 0. IInformations Interface (metadata + configuration)

**UUID:** `{d6f85c49-3a9c-4e1c-8f3f-6b8cc3b83c62}`

Plugins MUST implement `IInformations` and expose it via `QueryInterface` on the `IFileSystem` instance.

```cpp
typedef struct PluginMetaData
{
    // Stable plugin identifier (non-localized, long form).
    // Examples: "builtin/file-system", "builtin/file-system-dummy", "user/my-plugin".
    const wchar_t* id;
    // Short identifier used as the NavigationView path prefix: `<shortId>:<path>`.
    // Examples: "file", "fk", "ftp", "s3".
    // Must not be a single alphabetic character (reserved for Windows drive letters like "C:").
    const wchar_t* shortId;
    // Localized display name for UI.
    const wchar_t* name;
    // Localized description for "About" UI.
    const wchar_t* description;
    // Optional author/organization (may be nullptr).
    const wchar_t* author;
    // Optional version string (may be nullptr).
    const wchar_t* version;
} PluginMetaData;

interface __declspec(uuid("d6f85c49-3a9c-4e1c-8f3f-6b8cc3b83c62"))
	         __declspec(novtable)
	         IInformations : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept = 0;
};
```

**Ownership / lifetime:**
- All returned pointers are owned by the plugin object; the host MUST NOT free them.
- Returned pointers remain valid until the next call to the same method or until the COM object is released.

**Path prefix (host UI contract):**
- Hosts route navigation based on the **short ID** prefix: `<pluginShortId>:<pluginPath>`.
- Example: `fk:/` selects plugin `fk` and passes `/` to `IFileSystem::ReadDirectoryInfo`.
- Long IDs are used for persistence (settings, enable/disable, configuration keys); both long IDs and short IDs must be globally unique. A conflicting plugin is rejected and logged.
- ID comparisons are case-insensitive; `file` and `File` are considered the same ID.

#### Configuration schema contract

`GetConfigurationSchema()` returns a JSON/JSON5 object (UTF-8 string) describing fields for a dynamic configuration dialog:

```json5
{
  "version": 1,
  "title": "My Plugin",
  "fields": [
    {
      "key": "displayName",
      "label": "Display Name",
      "type": "text",
      "default": "Hello",
      "description": "Shown in the UI title bar."
    },
    {
      "key": "maxItems",
      "label": "Max Items",
      "type": "value",
      "default": 10,
      "min": 0,
      "max": 200,
      "description": "Upper bound for items returned by the plugin."
    },
    {
      "key": "halfSize",
      "label": "Half size",
      "type": "bool",
      "default": true,
      "description": "Decode at half resolution for faster load and lower memory use."
    },
    {
      "key": "mode", "label": "Mode", "type": "option", "default": "fast",
      "options": [ { "value": "fast", "label": "Fast" }, { "value": "safe", "label": "Safe" } ],
      "description": "Fast favors speed; Safe favors reliability."
    },
    {
      "key": "features", "label": "Features", "type": "selection",
      "options": [ { "value": "a", "label": "Feature A" }, { "value": "b", "label": "Feature B" } ],
      "default": ["a"],
      "description": "Toggle optional capabilities."
    }
  ]
}
```

Field types:
- `text`: single-line edit control (stored as string)
- `value`: numeric input (stored as integer)
- `bool`: boolean toggle (stored as boolean)
- `option`: radio group (stored as string, must match one of `options[].value`)
- `selection`: checkbox list (stored as string array of selected `options[].value`)
- `description` (optional): explanatory text rendered under the control along with defaults/min/max.

`GetConfiguration()` / `SetConfiguration()` exchange a JSON/JSON5 object (UTF-8 string) keyed by `fields[].key` with values matching the declared types.

### 0b. IFileSystemInitialize Interface (optional)

**UUID:** `{a4bdbb56-4f3f-4c1b-9b28-2f4c4a08d7af}`

Plugins MAY implement `IFileSystemInitialize` and expose it via `QueryInterface` on the `IFileSystem` instance.

This interface allows the host to configure a **per-instance** “mount context” (for example: an archive file path, an FTP endpoint, an S3 bucket, etc.) before calling `ReadDirectoryInfo`.

```cpp
interface __declspec(uuid("a4bdbb56-4f3f-4c1b-9b28-2f4c4a08d7af"))
	         __declspec(novtable)
	         IFileSystemInitialize : public IUnknown
{
    // rootPath: plugin-defined mount context (UTF-16, NUL-terminated).
    // optionsJsonUtf8: optional JSON/JSON5 payload (UTF-8, NUL-terminated) for per-instance options (passwords, initial path, etc.).
    virtual HRESULT STDMETHODCALLTYPE Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8) noexcept = 0;
};
```

**Host contract:**
- The host SHOULD call `Initialize()` whenever the mount context changes for a plugin instance.
- If the host changes the mount context for an existing instance, it MUST invalidate any cached enumeration results for that instance.

**Plugin contract:**
- Plugins MUST treat `Initialize()` as setting the base/root context for subsequent `ReadDirectoryInfo` calls.
- Plugins MUST copy any input strings they need to keep; callers own the input buffers.

### 1. INavigationMenu Interface (optional)

**UUID:** `{a7c7d693-5ba9-4f4d-8e90-0a2d9d7e49e4}`

Plugins MAY implement `INavigationMenu` and expose it via `QueryInterface` on the `IFileSystem` instance.
The host uses this interface to populate the main NavigationView dropdown menu with **raw menu entries**.

```cpp
enum NavigationMenuItemFlags : uint32_t
{
    NAV_MENU_ITEM_FLAG_NONE = 0,
    NAV_MENU_ITEM_FLAG_SEPARATOR = 0x1,
    NAV_MENU_ITEM_FLAG_DISABLED = 0x2,
    NAV_MENU_ITEM_FLAG_HEADER = 0x4,
} NavigationMenuItemFlags;

struct NavigationMenuItem
{
    NavigationMenuItemFlags flags;
    const wchar_t* label;     // Display label (UTF-16). nullptr/empty for separators.
    const wchar_t* path;      // Navigation target (UTF-16). nullptr when not applicable.
    const wchar_t* iconPath;  // Optional icon path for shell icon lookup (UTF-16).
    unsigned int commandId;   // Optional command identifier; 0 when not applicable.
};

interface __declspec(uuid("a7c7d693-5ba9-4f4d-8e90-0a2d9d7e49e4"))
         __declspec(novtable)
         INavigationMenu : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetMenuItems(
        const NavigationMenuItem** items,
        unsigned int* count
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE ExecuteMenuCommand(
        unsigned int commandId
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE SetCallback(
        INavigationMenuCallback* callback,
        void* cookie
    ) noexcept = 0;
};
```

### 1a. INavigationMenuCallback (optional)

```cpp
// Host callback for plugin-driven navigation requests.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at registration time and must be passed back verbatim by the plugin.
interface __declspec(novtable) INavigationMenuCallback
{
    virtual HRESULT STDMETHODCALLTYPE RequestNavigate(
        const wchar_t* path,
        void* cookie
    ) noexcept = 0;
};
```

**Callback behavior:**
- The host calls `SetCallback(hostCallback, cookie)` when `INavigationMenu` is available.
- The host calls `SetCallback(nullptr, nullptr)` when switching/unloading the active file system.
- Plugins can call `RequestNavigate(path, cookie)` (typically from `ExecuteMenuCommand`) to request navigation.
- `path` is a **plugin path** for the active file system (no `<shortId>:` prefix).
- The plugin MUST pass back the `cookie` it received in `SetCallback` unchanged.
- The plugin MUST NOT call the callback after the host clears it via `SetCallback(nullptr, nullptr)`.

**NavigationView behavior:**
- Order is preserved exactly as provided by the plugin.
- `NAV_MENU_ITEM_FLAG_SEPARATOR` inserts a separator (label/path ignored).
- `NAV_MENU_ITEM_FLAG_HEADER` renders a disabled header row.
- `path` is interpreted as a **plugin path** for the active `IFileSystem` (no `<shortId>:` prefix and no `<instanceContext>|` mount prefix).
  - For `file`, the host uses Windows absolute paths (e.g., `C:\...`).
  - For non-`file`, plugin paths are rooted at `/` and use `/` separators.
  - For non-`file`, the host persists “canonical locations” in history/settings as:
    - `<shortId>:<pluginPath>`
    - `<shortId>:<instanceContext>|<pluginPath>` when a mount context is configured via `IFileSystemInitialize::Initialize`
  - NavigationView breadcrumb/full-path display hides `<shortId>` and `<instanceContext>` and renders only the plugin path (starting at `/`).
  - NavigationView edit mode shows `<shortId>:<pluginPath>` (mount context not shown).
- If both `path` and `commandId` are provided, navigation takes precedence.
- If `path` is empty and `commandId != 0`, the host calls `ExecuteMenuCommand(commandId)`.
- The host assigns temporary Win32 menu item IDs for actionable rows (a reserved internal range in `NavigationView`); `commandId` is plugin-defined and is not treated as a Win32 `WM_COMMAND` identifier.
- Plugins SHOULD keep `commandId` values stable and unique within the returned list for predictable command routing/debugging.
- If `iconPath` is provided, the host uses `IconCache::QuerySysIconIndexForPath(iconPath, ...)`.
  If `iconPath` is empty and `path` is provided, the host uses `path` to resolve the icon.

**Ownership / lifetime:**
- The plugin owns the returned array and strings; they remain valid until the next call to the same method or until the COM object is released.

### 2. IDriveInfo Interface (optional)

**UUID:** `{b612a5d1-7e55-4e08-a3da-8d0d9f5d0f31}`

Plugins MAY implement `IDriveInfo` and expose it via `QueryInterface` on the `IFileSystem` instance.
The host uses this interface to populate the disk information section and menu in NavigationView.

```cpp
typedef enum DriveInfoFlags
{
    DRIVE_INFO_FLAG_NONE = 0,
    DRIVE_INFO_FLAG_HAS_DISPLAY_NAME = 0x1,
    DRIVE_INFO_FLAG_HAS_VOLUME_LABEL = 0x2,
    DRIVE_INFO_FLAG_HAS_FILE_SYSTEM = 0x4,
    DRIVE_INFO_FLAG_HAS_TOTAL_BYTES = 0x8,
    DRIVE_INFO_FLAG_HAS_FREE_BYTES = 0x10,
    DRIVE_INFO_FLAG_HAS_USED_BYTES = 0x20,
} DriveInfoFlags;

typedef struct DriveInfo
{
    DriveInfoFlags flags;
    const wchar_t* displayName; // Header text (e.g., "C:\\" or "s3://bucket").
    const wchar_t* volumeLabel; // Optional volume label.
    const wchar_t* fileSystem;  // Optional file system name.
    uint64_t totalBytes;
    uint64_t freeBytes;
    uint64_t usedBytes;
} DriveInfo;

interface __declspec(uuid("b612a5d1-7e55-4e08-a3da-8d0d9f5d0f31"))
         __declspec(novtable)
         IDriveInfo : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetDriveInfo(
        const wchar_t* path,
        DriveInfo* info
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE GetDriveMenuItems(
        const wchar_t* path,
        const NavigationMenuItem** items,
        unsigned int* count
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE ExecuteDriveMenuCommand(
        unsigned int commandId,
        const wchar_t* path
    ) noexcept = 0;
};
```

**NavigationView behavior:**
- If `GetDriveInfo` fails or returns `S_FALSE`, the disk info section shows empty text + grey progress bar.
- `displayName` controls the disk info header text; if absent, the host falls back to the path root.
- The host formats labels using resource format strings (see `IDS_FMT_DISK_*`).
- `usedBytes` is preferred when provided; otherwise the host computes `total - free`.
- `GetDriveMenuItems` entries are appended under disk info using the same rendering rules as `INavigationMenu`.
- The host passes a **plugin path** to `GetDriveInfo` / `GetDriveMenuItems` / `ExecuteDriveMenuCommand`:
  - For non-`file` plugins, the host strips the `<shortId>:` display prefix before calling.
  - For `file`, the host passes a Windows absolute path.

**Ownership / lifetime:**
- The plugin owns `displayName`/`volumeLabel`/`fileSystem` strings and the returned menu array, valid until the next call to the same method or until the COM object is released.

### 3. IFileSystem Interface

**UUID:** `{12519afa-30e7-4e3a-9db2-7990c4be9a21}`

The primary interface for directory enumeration and file operations. This supersedes the original enumeration-only `IFileSystem` (`{fd57f938-3def-43d1-ba8f-2d27adf4fd56}`). Search remains an optional interface (`IFileSystemSearch`).

```cpp
interface __declspec(uuid("12519afa-30e7-4e3a-9db2-7990c4be9a21"))
         __declspec(novtable)
         IFileSystem : public IUnknown
{
    // Reads complete directory listing in a single call
    // Parameters:
    //   path: Directory path (null-terminated wide string, e.g., L"C:\\Users")
    //   ppFilesInformation: [out] Interface to access directory entries
    // Returns:
    //   S_OK: Success, ppFilesInformation contains valid interface
    //   E_POINTER: Null output pointer
    //   E_INVALIDARG: Invalid path (null or empty)
    //   E_OUTOFMEMORY: Allocation failed
    //   HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND/ERROR_PATH_NOT_FOUND): Directory doesn't exist
    //   HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED): Permission denied
    //   HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER): Directory too large for plugin limits
    virtual HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(
        const wchar_t* path,
        IFilesInformation** ppFilesInformation
    ) noexcept = 0;

    // File operations (see "File Operations" section for types, flags, and callback contracts).
    virtual HRESULT STDMETHODCALLTYPE CopyItem(
        const wchar_t* sourcePath,
        const wchar_t* destinationPath,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: unlimited
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveItem(
        const wchar_t* sourcePath,
        const wchar_t* destinationPath,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: unlimited
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DeleteItem(
        const wchar_t* path,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: no bandwidth limit
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE RenameItem(
        const wchar_t* sourcePath,
        const wchar_t* destinationPath,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: no bandwidth limit
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    // Batch operations (order preserved).
    virtual HRESULT STDMETHODCALLTYPE CopyItems(
        const wchar_t* const* sourcePaths,
        unsigned long count,
        const wchar_t* destinationFolder,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: unlimited
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveItems(
        const wchar_t* const* sourcePaths,
        unsigned long count,
        const wchar_t* destinationFolder,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: unlimited
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DeleteItems(
        const wchar_t* const* paths,
        unsigned long count,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: no bandwidth limit
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE RenameItems(
        const FileSystemRenamePair* items,
        unsigned long count,
        FileSystemFlags flags,
        const FileSystemOptions* options = nullptr, // default: no bandwidth limit
        IFileSystemCallback* callback = nullptr,    // default: no callbacks
        void* cookie = nullptr                      // default: none
    ) noexcept = 0;

    // Returns a UTF-8 JSON/JSON5 string describing supported operations and cross-filesystem policy.
    // The string pointer is owned by the plugin and remains valid until the next call or object release.
    // Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) when unsupported.
    virtual HRESULT STDMETHODCALLTYPE GetCapabilities(
        const char** jsonUtf8
    ) noexcept = 0;
};
```

**Design Notes:**
- Single-call enumeration model (all entries returned at once)
- No incremental/streaming API (entire directory loaded into memory)
- Suitable for typical user directories (thousands of files)
- For massive directories (100K+ files), consider pagination in future versions

### 4. IFileSystemIO Interface

**UUID:** `{2c7c32b3-8a0f-4e25-8d3a-6a5f1d0a1e2c}`

Optional (but currently required by the host) I/O interface for:
- fast attribute queries (without doing a full `ReadDirectoryInfo()` enumeration)
- opening a read-only stream for file contents (`IFileReader`)
- retrieving item properties for a themed Properties dialog (see `GetItemProperties`)

The host obtains this interface via `QueryInterface` on the active `IFileSystem` instance.

```cpp
interface __declspec(uuid("2c7c32b3-8a0f-4e25-8d3a-6a5f1d0a1e2c"))
         __declspec(novtable)
         IFileSystemIO : public IUnknown
{
    // On success, fileAttributes receives FILE_ATTRIBUTE_* flags (e.g. FILE_ATTRIBUTE_DIRECTORY).
    // Returns HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND/ERROR_PATH_NOT_FOUND) when the item does not exist.
    // `path` is a filesystem-internal path (not necessarily a Win32 path).
    virtual HRESULT STDMETHODCALLTYPE GetAttributes(
        const wchar_t* path,
        unsigned long* fileAttributes
    ) noexcept = 0;

    // Creates a read-only file reader for `path`.
    // `path` is a filesystem-internal path (not necessarily a Win32 path).
    virtual HRESULT STDMETHODCALLTYPE CreateFileReader(
        const wchar_t* path,
        IFileReader** reader
    ) noexcept = 0;

    // Creates a file writer for `path`.
    // Default behavior: fail if the destination already exists.
    // Flags:
    // - FILESYSTEM_FLAG_ALLOW_OVERWRITE: replace existing file.
    // - FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY: allow replacing read-only targets (filesystem-defined behavior).
    virtual HRESULT STDMETHODCALLTYPE CreateFileWriter(
        const wchar_t* path,
        FileSystemFlags flags,
        IFileWriter** writer
    ) noexcept = 0;

    // Returns a UTF-8 JSON/JSON5 string describing properties for a single item.
    // The string pointer is owned by the plugin and remains valid until the next call or object release.
    // Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) when unsupported.
    virtual HRESULT STDMETHODCALLTYPE GetItemProperties(
        const wchar_t* path,
        const char** jsonUtf8
    ) noexcept = 0;
};
```

### 4a. IFileWriter Interface (optional)

**UUID:** `{b6f0a9e1-8c8b-4b72-9f3e-2f2b4b8b9c41}`

Optional write-only stream abstraction for filesystem plugins.

The host uses this interface for **cross-filesystem copy/move bridging** (streaming `IFileReader` → `IFileWriter`) and for other host-managed data transfers.

Notes:
- Implementations MUST be safe for large files (64-bit offsets/sizes).
- Implementations SHOULD support sequential writes efficiently.
- Implementations MUST tolerate being released without `Commit()` (treat as abort/best-effort cleanup).

```cpp
interface __declspec(uuid("b6f0a9e1-8c8b-4b72-9f3e-2f2b4b8b9c41"))
         __declspec(novtable)
         IFileWriter : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Commit() noexcept = 0;
};
```

### 4c. IFileSystemDirectoryOperations Interface (optional)

**UUID:** `{4a8f7cf2-f81c-4278-b182-7183e6bed6f3}`

Optional directory operations interface for file systems that support creating folders and computing directory sizes.

The host obtains this interface via `QueryInterface` on the active `IFileSystem` instance.

```cpp
struct FileSystemDirectorySizeResult
{
    uint64_t totalBytes;     // Total size in bytes (sum of file sizes).
    uint64_t fileCount;      // Number of files counted.
    uint64_t directoryCount; // Number of directories counted (excluding root).
    HRESULT status;                  // S_OK, HRESULT_FROM_WIN32(ERROR_CANCELLED), or first error.
};

interface __declspec(novtable) IFileSystemDirectorySizeCallback
{
    virtual HRESULT STDMETHODCALLTYPE DirectorySizeProgress(
        uint64_t scannedEntries,
        uint64_t totalBytes,
        uint64_t fileCount,
        uint64_t directoryCount,
        const wchar_t* currentPath,
        void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(
        BOOL* pCancel,
        void* cookie) noexcept = 0;
};

interface __declspec(uuid("4a8f7cf2-f81c-4278-b182-7183e6bed6f3"))
         __declspec(novtable)
         IFileSystemDirectoryOperations : public IUnknown
{
    // Creates a directory at `path`.
    // Returns HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) when the directory already exists.
    virtual HRESULT STDMETHODCALLTYPE CreateDirectory(
        const wchar_t* path
    ) noexcept = 0;

    // Compute the total size of a directory tree.
    // - path: Root directory to start from (must exist and be a directory).
    // - flags: Use FILESYSTEM_FLAG_RECURSIVE for recursive computation; otherwise only immediate children.
    // - callback: Optional progress callback (may be nullptr for synchronous completion).
    // - cookie: Opaque value passed to callback.
    // - result: [out] Output result structure.
    // Returns: S_OK on success, HRESULT_FROM_WIN32(ERROR_CANCELLED) if cancelled via callback.
    virtual HRESULT STDMETHODCALLTYPE GetDirectorySize(
        const wchar_t* path,
        FileSystemFlags flags,
        IFileSystemDirectorySizeCallback* callback,
        void* cookie,
        FileSystemDirectorySizeResult* result
    ) noexcept = 0;
};
```

### 4d. Capabilities (`IFileSystem::GetCapabilities`) (optional)

Provides a **read-only** capabilities declaration for the filesystem plugin.

The host uses this to:
- enable/disable UI commands (rename/delete/properties, etc.)
- decide whether cross-filesystem copy/move is allowed (explicit opt-in, per plugin pair)

Capabilities are returned via `IFileSystem::GetCapabilities(...)` on the active filesystem instance.

**Per-instance rule:** capabilities are **per `IFileSystem` instance**, and MAY vary based on:
- plugin mode (e.g. S3 vs S3Tables),
- transport/protocol (e.g. FTP vs IMAP within a single DLL),
- connection/profile configuration.

The host MUST treat capabilities as instance-scoped and MUST NOT assume a DLL has a single fixed capability set.

Capabilities JSON (version 1):

```json5
{
  "version": 1,
  "operations": {
    "copy": true,
    "move": true,
    "delete": true,
    "rename": true,
    "properties": true,
    "read": true,   // supports enumeration + IFileReader
    "write": false  // supports IFileWriter via IFileSystemIO::CreateFileWriter
  },
  "crossFileSystem": {
    // Plugin ids are long-form ids (PluginMetaData.id), case-insensitive.
    // "*" means any plugin id.
    "export": { "copy": ["*"], "move": ["*"] }, // this plugin may be the SOURCE when destination id matches
    "import": { "copy": ["builtin/file-system"], "move": [] } // this plugin may be the DESTINATION when source id matches
  }
}
```

### 4e. Item properties (`IFileSystemIO::GetItemProperties`) (optional)

Provides item properties for non-Win32 paths (and optionally for Win32 paths) so the host can show a themed properties dialog.

Properties are returned via `IFileSystemIO::GetItemProperties(...)` on the active filesystem instance.

Properties JSON (version 1) (minimal shape):

```json5
{
  "version": 1,
  "title": "Properties",
  "sections": [
    {
      "title": "General",
      "fields": [
        { "key": "name", "value": "file.txt" },
        { "key": "path", "value": "/folder/file.txt" },
        { "key": "sizeBytes", "value": "12345" }
      ]
    }
  ]
}
```

#### CreateDirectory

Creates a new directory at the specified path.

**Parameters:**
- `path`: Full path of the directory to create (UTF-16, NUL-terminated).

**Return Value:**
- `S_OK`: Directory created successfully.
- `HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)`: Target already exists.
- `HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)`: Parent directory does not exist.
- `HRESULT_FROM_WIN32(ERROR_INVALID_NAME)`: Invalid path or directory name (e.g., contains invalid characters for that filesystem).
- `E_POINTER`: `path` is `nullptr`.
- `E_INVALIDARG`: `path` is empty.
- `E_NOTIMPL`: Plugin does not support directory creation (e.g., read-only archives).

#### GetDirectorySize

Computes the total size of a directory tree, optionally recursively.

**Parameters:**
- `path`: Root directory to start from (must exist and be a directory).
- `flags`: Use `FILESYSTEM_FLAG_RECURSIVE` for recursive computation; otherwise only immediate children are counted.
- `callback`: Optional progress callback (may be `nullptr` for synchronous completion without progress).
- `cookie`: Opaque value passed to callback methods.
- `result`: [out] Output result structure containing totals and status.

**Return Value:**
- `S_OK`: Computation completed successfully.
- `HRESULT_FROM_WIN32(ERROR_CANCELLED)`: Operation was cancelled via callback.
- `HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)`: Path does not exist.
- `HRESULT_FROM_WIN32(ERROR_DIRECTORY)`: Path is not a directory.
- `E_POINTER`: `path` or `result` is `nullptr`.
- `E_INVALIDARG`: `path` is empty.

**Progress Reporting:**
- Plugins SHOULD report progress every 100 entries OR every 200ms, whichever comes first.
- Progress callback includes current totals and the path currently being scanned.
- Plugins MUST check `DirectorySizeShouldCancel` after each progress report.

**Callback Interface:**
- `IFileSystemDirectorySizeCallback` is NOT a COM interface (no `IUnknown` inheritance).
- Lifetime is managed by the host; the plugin MUST NOT store or release the callback.
- The `cookie` value MUST be passed back verbatim to all callback methods.

**Implementation Notes:**
- Plugins SHOULD skip reparse points (symlinks, junctions) to avoid infinite loops.
- Plugins SHOULD continue on access errors for individual subdirectories (report first error in `result->status`).
- The `directoryCount` field excludes the root directory itself.

**Host usage:**
- `FolderWindow` uses `CreateDirectory` for `F7` (Create directory) when available:
  - The host prompts for the new folder name (modal dialog centered on the main window).
  - The host displays the destination path where the folder will be created.
  - The host validates the typed folder name as a single path segment and rejects invalid characters (`\\ / : * ? " < > |`) before calling `CreateDirectory`.
  - If the interface is not available (QI fails), the host treats the operation as unsupported and shows a localized error message.
  - If `CreateDirectory` returns `E_NOTIMPL`, the host treats the operation as unsupported and shows a localized error message that includes the plugin display name.
- `FolderWindow` uses `GetDirectorySize` for selection size calculation and properties display.
- If not available, the host treats these operations as unsupported for that plugin.

### 2. IFilesInformation Interface

**UUID:** `{0d9ef549-4e54-4086-8a5c-f9d3e6120211}`

Manages directory enumeration results as a contiguous buffer of `FileInfo` structures. This abstraction avoids STL types for cross-compiler compatibility.

**Creation Rule (Host Side):**
- The host application MUST NOT create an `IFilesInformation` instance directly (not via `RedSalamanderCreate`, not via `new`, not via COM registration).
- `IFilesInformation` instances are only obtained as an **output** from `IFileSystem::ReadDirectoryInfo()`.

```cpp
interface __declspec(uuid("0d9ef549-4e54-4086-8a5c-f9d3e6120211")) 
         __declspec(novtable) 
         IFilesInformation : public IUnknown
{
    // Returns pointer to buffer containing all FileInfo entries
    // Buffer is owned by IFilesInformation instance (caller does NOT free)
    // Valid until IFilesInformation is released
    // Returns E_POINTER if ppFileInfo is nullptr.
    virtual HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept = 0;

    // Returns how many bytes in the buffer are committed/used by valid FileInfo entries.
    // This is the upper bound the host must use when walking NextEntryOffset.
    // Returns E_POINTER if pSize is nullptr.
    virtual HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept = 0;

    // Returns the allocated capacity of the backing buffer in bytes.
    // This may be larger than GetBufferSize().
    // Returns E_POINTER if pSize is nullptr.
    // May return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW) if the size doesn't fit in unsigned long.
    virtual HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept = 0;

    // Returns number of file entries in buffer (convenience method)
    // Returns E_POINTER if pCount is nullptr.
    virtual HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept = 0;

    // Retrieves pointer to specific entry by index (convenience method)
    // Index is 0-based, must be < GetCount()
    // Returns pointer into buffer (caller does NOT free)
    // Returns E_POINTER if ppEntry is nullptr.
    // Returns HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES) if index is out of range.
    virtual HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept = 0;
};
```



**Memory Ownership:**
- Plugin allocates buffer during `ReadDirectoryInfo()`
- Plugin owns buffer lifetime (freed when `IFilesInformation` is released)
- Caller must NOT free pointers from `GetBuffer()` or `Get()`
- Buffer remains valid until caller calls `Release()` on `IFilesInformation`

**Read-Only Contract (Required):**
- `ReadDirectoryInfo()` fully populates the results before returning `S_OK`
- After `ReadDirectoryInfo()` returns, the `IFilesInformation` object is **immutable**
- `GetBuffer()`, `GetBufferSize()`, `GetAllocatedSize()`, `GetCount()`, and `Get()` MUST be **side-effect free** and may be called repeatedly (including concurrently)
- `GetBuffer()` MUST NOT clear/reset internal state (e.g., count/used-bytes); it only exposes the already-built results
- Empty directory behavior:
  - `GetCount()` returns `0`
  - `GetBuffer()` sets `*ppFileInfo = nullptr` and returns `S_OK`

**Plugin Implementation Guidance (Internal Writer Path):**
- The plugin implementation should use a private/internal “begin write” + “commit” path to build the buffer during `ReadDirectoryInfo()`, rather than mutating state in `GetBuffer()`.
- This writer API is **not** part of the public COM interface and must not be called by the host.

**Usage Pattern (Primary Method - NextEntryOffset Traversal):**

The recommended way to iterate through FileInfo entries is via `NextEntryOffset` linked-list traversal:

```cpp
wil::com_ptr<IFilesInformation> info;
THROW_IF_FAILED(fileSystem->ReadDirectoryInfo(L"C:\\Users", info.put()));

FileInfo* current = nullptr;
THROW_IF_FAILED(info->GetBuffer(&current));

while (current) {
    // Process current entry (use FileNameSize; do not assume null-termination)
    const size_t nameChars = static_cast<size_t>(current->FileNameSize) / sizeof(wchar_t);
    const std::wstring_view name(current->FileName, nameChars);
    ProcessFile(name, current->EndOfFile, current->FileAttributes);
    
    // Move to next entry via NextEntryOffset
    if (current->NextEntryOffset == 0)
        break;  // Last entry
    current = reinterpret_cast<FileInfo*>(
        reinterpret_cast<uint8_t*>(current) + current->NextEntryOffset
    );
}
// info released automatically via wil::com_ptr
```

**Alternative Usage (Convenience Methods):**

For simpler code, use the convenience methods `GetCount()` and `Get()`:

```cpp
wil::com_ptr<IFilesInformation> info;
THROW_IF_FAILED(fileSystem->ReadDirectoryInfo(L"C:\\Users", info.put()));

unsigned long count = 0;
THROW_IF_FAILED(info->GetCount(&count));

for (unsigned long i = 0; i < count; i++) {
    FileInfo* entry = nullptr;
    THROW_IF_FAILED(info->Get(i, &entry));
    // Use entry->FileName + entry->FileNameSize (do not assume null-termination).
}
// info released automatically via wil::com_ptr
```

**Note:** The convenience methods (`GetCount`/`Get`) internally traverse the buffer using `NextEntryOffset`, so the primary method is more efficient for single-pass enumeration.

### 3. FileInfo Structure

Represents a single file or directory entry. Structure layout matches Windows `FILE_FULL_DIR_INFO` for easy mapping from native APIs.

```cpp
typedef struct FileInfo
{
    unsigned long NextEntryOffset;   // Offset to next entry (0 for last entry)
    unsigned long FileIndex;         // File system-specific index
    __int64 CreationTime;            // File creation time (FILETIME format)
    __int64 LastAccessTime;          // Last access time (FILETIME format)
    __int64 LastWriteTime;           // Last modification time (FILETIME format)
    __int64 ChangeTime;              // Metadata change time (FILETIME format)
    __int64 EndOfFile;               // File size in bytes
    __int64 AllocationSize;          // Allocated space on disk
    unsigned long FileAttributes;    // FILE_ATTRIBUTE_* flags
    unsigned long FileNameSize;      // Length of FileName in bytes (not characters)
    unsigned long EaSize;            // Extended attributes size
    wchar_t FileName[1];             // Variable-length file name (use FileNameSize to determine length)
} FileInfo;
```

**Field Notes:**
- **NextEntryOffset**: Linked-list style buffer traversal. Add to current pointer to get next entry. Zero for final entry.
- **FileNameSize**: Byte count, not character count. To get character count: `FileNameSize / sizeof(wchar_t)`
- **FileName**: Variable-length UTF-16 string; the valid length is exactly `FileNameSize` bytes. A trailing NUL may be present as a convenience (and is written by the reference plugin), but callers MUST NOT rely on it.
- **Times**: 64-bit FILETIME format (100-nanosecond intervals since January 1, 1601 UTC)

**Buffer Layout:**
```text
[FileInfo 1][FileInfo 2][FileInfo 3]...
 ^           ^
 |           |
 +-----------+ (NextEntryOffset bytes)
```

**Traversal Pattern:**

Each `FileInfo` structure's `NextEntryOffset` field points to the next entry in the buffer. To traverse:

1. Get the buffer head via `GetBuffer()`
2. Process the current entry
3. If `NextEntryOffset` is 0, you've reached the last entry
4. Otherwise, add `NextEntryOffset` bytes to the current pointer to get the next entry

This linked-list approach allows variable-length entries (since `FileName` is variable length) and matches Windows native directory enumeration APIs (`FILE_FULL_DIR_INFO`).

**Why Not Array-Based Access?**

The buffer is NOT an array of fixed-size structures. Each entry has a variable length due to the `FileName` field. The `Get(index)` convenience method internally traverses using `NextEntryOffset` to find the requested index, making it less efficient for full enumeration compared to direct traversal.

### 4. File Operations (Copy/Move/Delete/Rename)

File operations are exposed as methods on `IFileSystem`.

#### Supporting Types

```cpp
typedef enum FileSystemOperation
{
    FILESYSTEM_COPY = 1,
    FILESYSTEM_MOVE = 2,
    FILESYSTEM_DELETE = 3,
    FILESYSTEM_RENAME = 4,
    FILESYSTEM_TYPE_FORCE_DWORD = 0xffffffff
} FileSystemOperation;

typedef enum FileSystemFlags
{
    FILESYSTEM_FLAG_NONE = 0,
    FILESYSTEM_FLAG_ALLOW_OVERWRITE = 0x1,
    FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY = 0x2,
    FILESYSTEM_FLAG_RECURSIVE = 0x4,
    FILESYSTEM_FLAG_USE_RECYCLE_BIN = 0x8,
    FILESYSTEM_FLAG_CONTINUE_ON_ERROR = 0x10,
    FILESYSTEM_FLAG_FORCE_DWORD = 0xffffffff
} FileSystemFlags;

typedef struct FileSystemOptions
{
    // 0 = unlimited (use all available bandwidth).
    // Callbacks receive an in/out FileSystemOptions* so the caller can tweak it on progress updates.
    uint64_t bandwidthLimitBytesPerSecond;
} FileSystemOptions;

typedef struct FileSystemRenamePair
{
    // Pointers reference NUL-terminated UTF-16 strings stored in a caller-owned arena.
    const wchar_t* sourcePath;
    const wchar_t* newName; // Leaf name only (no path separators).
} FileSystemRenamePair;

typedef struct FileSystemArena
{
    unsigned char* buffer;
    unsigned long capacityBytes;
    unsigned long usedBytes;
} FileSystemArena;
```

**Notes:**
- `destinationPath` for `CopyItem`/`MoveItem`/`RenameItem` is the full target path.
- `CopyItems`/`MoveItems` take a shared `destinationFolder`; each source path is copied/moved to `destinationFolder` + leaf name.
- `destinationFolder` must be non-null and non-empty for `CopyItems`/`MoveItems`.
- `RenameItems` uses `FileSystemRenamePair::newName` as a leaf name only (no path separators).
- If `newName` contains `\\` or `/`, plugins SHOULD return `HRESULT_FROM_WIN32(ERROR_INVALID_NAME)`.
- `destinationPath` is ignored for delete (single/batch).
- For `RenameItems`, callback `destinationPath` values are the full target paths computed from `sourcePath` + `newName`.
- `totalItems`/`totalBytes` may be `0` when unknown.
- Pointer parameters passed to callbacks are only valid for the duration of the callback call.
- `FileSystemOptions::bandwidthLimitBytesPerSecond` applies to data-transfer operations (copy/move) and MAY be ignored for rename/delete.
- The callback receives an in/out `FileSystemOptions* options` parameter; callers MAY adjust it and plugins SHOULD use the updated values for subsequent work.
- Callback `options` may be `nullptr`; callers must check before writing to it.
- Default options: `options == nullptr` (unlimited bandwidth). If `options` is provided, `bandwidthLimitBytesPerSecond == 0` is treated as unlimited.
- If `FILESYSTEM_FLAG_CONTINUE_ON_ERROR` is not set, plugins SHOULD stop at the first failure.
- If any item fails and the operation continues, plugins SHOULD return `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)`.
- If `FILESYSTEM_FLAG_RECURSIVE` is not set and a directory requires recursion, plugins SHOULD return `HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY)`.

**Arena Pattern (Required):**
- All pointer fields in `FileSystemRenamePair`, `FileSystemSearchQuery`, and `FileSystemSearchMatch`, plus all string pointer parameters passed via `IFileSystemCallback`, MUST reference memory inside a `FileSystemArena`.
- Arena strings are UTF-16 and NUL-terminated.
- Arrays of `FileSystemRenamePair` and arrays passed to `CopyItems`/`MoveItems`/`DeleteItems` MUST be allocated from the same arena as their referenced strings to allow a single free/reset.
- Input arenas are owned by the caller and must remain valid for the full duration of the operation.
- Callback arenas are owned by the plugin and must remain valid until the callback returns (callers must copy if they need to persist values).

#### IFileSystemCallback Interface

```cpp
// Host callback for file operation progress.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
// - progressStreamId identifies a concurrent progress stream (e.g. a worker) and must be stable per stream.
enum class FileSystemIssueAction : uint8_t
{
    None = 0,
    Overwrite,
    ReplaceReadOnly,
    PermanentDelete,
    Retry,
    Skip,
    Cancel,
};

interface __declspec(novtable) IFileSystemCallback
{
    // Progress updates. Return E_ABORT or HRESULT_FROM_WIN32(ERROR_CANCELLED) to cancel.
    // The host may update FileSystemOptions during this callback.
    // options may be nullptr; callers must check before writing to it.
    virtual HRESULT STDMETHODCALLTYPE FileSystemProgress(
        FileSystemOperation operationType,
        unsigned long totalItems,
        unsigned long completedItems,
        uint64_t totalBytes,
        uint64_t completedBytes,
        const wchar_t* currentSourcePath,
        const wchar_t* currentDestinationPath,
        uint64_t currentItemTotalBytes,
        uint64_t currentItemCompletedBytes,
        FileSystemOptions* options,
        uint64_t progressStreamId,
        void* cookie
    ) noexcept = 0;

    // Per-item completion callback (success or failure).
    // options may be nullptr; callers must check before writing to it.
    virtual HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(
        FileSystemOperation operationType,
        unsigned long itemIndex,
        const wchar_t* sourcePath,
        const wchar_t* destinationPath,
        HRESULT status,
        FileSystemOptions* options,
        void* cookie
    ) noexcept = 0;

    // Called by the plugin to check for cancellation.
    virtual HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept = 0;

    // Called by the plugin when an operation hits a conflict/issue that requires a user decision.
    // action must be non-null. Implementations should set it even when returning failure/cancellation.
    virtual HRESULT STDMETHODCALLTYPE FileSystemIssue(
        FileSystemOperation operationType,
        const wchar_t* sourcePath,
        const wchar_t* destinationPath,
        HRESULT status,
        FileSystemIssueAction* action,
        FileSystemOptions* options,
        void* cookie
    ) noexcept = 0;
};
```

#### Arena Helpers (C++ only)

```cpp
HRESULT InitializeFileSystemArena(FileSystemArena* arena, unsigned long capacityBytes) noexcept;
void DestroyFileSystemArena(FileSystemArena* arena) noexcept;
void* AllocateFromFileSystemArena(FileSystemArena* arena, unsigned long sizeBytes, unsigned long alignment) noexcept;

HRESULT BuildFileSystemPathListArenaFromFilesInformation(
    const wchar_t* sourceRoot,
    IFilesInformation* filesInformation,
    FileSystemArena* arena,
    const wchar_t*** outPaths,
    unsigned long* outCount) noexcept;

class FileSystemArenaOwner final
{
public:
    FileSystemArenaOwner() noexcept = default;
    ~FileSystemArenaOwner() noexcept;
    FileSystemArenaOwner(const FileSystemArenaOwner&) = delete;
    FileSystemArenaOwner& operator=(const FileSystemArenaOwner&) = delete;
    FileSystemArenaOwner(FileSystemArenaOwner&& other) noexcept;
    FileSystemArenaOwner& operator=(FileSystemArenaOwner&& other) noexcept;

    FileSystemArena* Get() noexcept;
    const FileSystemArena* Get() const noexcept;
    void Reset() noexcept;
    HRESULT Initialize(unsigned long capacityBytes) noexcept;

    HRESULT BuildPathListFromFilesInformation(
        const wchar_t* sourceRoot,
        IFilesInformation* filesInformation,
        const wchar_t*** outPaths,
        unsigned long* outCount) noexcept;
};
```

**Behavior:**
- `BuildFileSystemPathListArenaFromFilesInformation` builds a single arena that contains the path array and NUL-terminated full paths for every entry in `filesInformation`.
- The caller owns `arena` and must destroy it when finished.
- `FileSystemArenaOwner` is an RAII wrapper that destroys its arena on scope exit.

**Cancellation Contract:**
- If `ShouldCancel` returns `TRUE`, the plugin MUST stop as soon as practical and return `HRESULT_FROM_WIN32(ERROR_CANCELLED)`.
- If `OnProgress` or `OnItemCompleted` returns `E_ABORT` or `HRESULT_FROM_WIN32(ERROR_CANCELLED)`, the plugin MUST treat it as a cancellation request.

**Operational Notes:**
- All operations are synchronous. The host should invoke them on a worker thread.
- `callback` may be `nullptr` (default: no progress/cancel callbacks).
- `options` may be `nullptr` (default: unlimited bandwidth).
- Unsupported flags for a given operation SHOULD return `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)`; irrelevant flags MAY be ignored.
- `RenameItems` uses `FileSystemRenamePair::newName` (leaf name only); `RenameItem` still requires a full `destinationPath`.

### 5. Search

Search is exposed via an optional `IFileSystemSearch` interface, obtained via `QueryInterface` on `IFileSystem`.

#### Supporting Types

```cpp
enum FileSystemSearchFlags : uint32_t
{
    FILESYSTEM_SEARCH_NONE = 0,
    FILESYSTEM_SEARCH_RECURSIVE = 0x1,
    FILESYSTEM_SEARCH_INCLUDE_FILES = 0x2,
    FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES = 0x4,
    FILESYSTEM_SEARCH_MATCH_CASE = 0x8,
    FILESYSTEM_SEARCH_FOLLOW_SYMLINKS = 0x10,
    FILESYSTEM_SEARCH_USE_REGEX = 0x20,
};

struct FileSystemSearchQuery
{
    const wchar_t* rootPath;
    const wchar_t* pattern; // nullptr/empty = L"*"
    FileSystemSearchFlags flags;
    unsigned long maxResults; // 0 = unlimited
};

struct FileSystemSearchMatch
{
    const wchar_t* fullPath;
    unsigned long fullPathSize;
    unsigned long fileAttributes;
    __int64 creationTime;
    __int64 lastAccessTime;
    __int64 lastWriteTime;
    __int64 changeTime;
    __int64 endOfFile;
    __int64 allocationSize;
};

struct FileSystemSearchProgress
{
    uint64_t scannedEntries;
    uint64_t matchedEntries;
    const wchar_t* currentPath;
};
```

**Search Notes:**
- `pattern` uses wildcard matching by default (`*` and `?`). If `FILESYSTEM_SEARCH_USE_REGEX` is set, `pattern` is treated as a regex.
- `pattern` may be `nullptr` or empty to match all items (default: `L"*"`).
- `fullPathSize` is in bytes (not characters).
- `maxResults` of `0` means unlimited (default).

#### IFileSystemSearchCallback Interface

```cpp
// Host callback for file system search results.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
interface __declspec(novtable) IFileSystemSearchCallback
{
    // Called for every match. Return E_ABORT or HRESULT_FROM_WIN32(ERROR_CANCELLED) to cancel.
    virtual HRESULT STDMETHODCALLTYPE OnMatch(
        const FileSystemSearchMatch* match,
        void* cookie
    ) noexcept = 0;

    // Periodic progress updates (optional but recommended for long searches).
    virtual HRESULT STDMETHODCALLTYPE OnProgress(
        const FileSystemSearchProgress* progress,
        void* cookie
    ) noexcept = 0;

    // Called by the plugin to check for cancellation.
    virtual HRESULT STDMETHODCALLTYPE ShouldCancel(BOOL* pCancel, void* cookie) noexcept = 0;
};
```

#### IFileSystemSearch Interface

**UUID:** `{00417f3e-f0f5-4add-8dea-4407d5169ef6}`

```cpp
interface __declspec(uuid("00417f3e-f0f5-4add-8dea-4407d5169ef6"))
         __declspec(novtable)
         IFileSystemSearch : public IUnknown
{
    // Synchronous search. Results are delivered via callback.
    virtual HRESULT STDMETHODCALLTYPE Search(
        const FileSystemSearchQuery* query,
        IFileSystemSearchCallback* callback,
        void* cookie
    ) noexcept = 0;
};
```

**Search Contract:**
- `Search` is synchronous; the host should invoke it on a worker thread.
- All callback pointers are only valid for the duration of the call; the host must copy strings if needed.
- If `ShouldCancel` returns `TRUE`, or a callback returns `E_ABORT`/`HRESULT_FROM_WIN32(ERROR_CANCELLED)`, the plugin MUST stop and return `HRESULT_FROM_WIN32(ERROR_CANCELLED)`.


## Implementation Details

### Plugin Implementation Requirements

Plugins must implement the `IFileSystem` interface and provide concrete implementations for all methods. Unsupported operations MUST return `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)`. The reference implementation in `Plugins/FileSystem/` demonstrates the pattern.

**Core Requirements:**
- Implement `IUnknown` reference counting properly
- Return appropriate HRESULTs for all error conditions
- Ensure thread-safety if plugin will be called from multiple threads
- Handle Unicode paths correctly (UTF-16)
- Follow RAII principles for resource management
- For operations/search, honor cancellation requests and do not invoke callbacks after returning

### Plugin Discovery and Loading

**Current Implementation:** RedSalamander discovers and manages file system plugins via `RedSalamander/FileSystemPluginManager.*`:
- **Embedded**: `Plugins\\FileSystem.dll` next to `RedSalamander.exe`
- **Optional**: `Plugins\\*.dll` next to the executable (DLLs without `RedSalamanderCreate` are ignored)
- **Custom**: absolute plugin paths from user settings (`plugins.customPluginPaths[]`); plugins are referenced in place (no copying)
- Plugins must expose **both** a unique long ID (`builtin/...` or `user/...`) and a unique short ID (navigation scheme). Conflicts are logged and the plugin is skipped/unloaded.
- DLLs missing the `RedSalamanderCreate` export are not shown in the plugin list.
- The **Plugins** top-level menu sits between **File** and **View** and contains a `Manage Plugins...` entry plus a pane-specific dynamic list.
- The Plugin Manager dialog lists plugins grouped as Embedded / Optional / Custom, with columns **Plugin** and **Short Id**, and action buttons stacked vertically (`Add...`, `Remove...`, `Configure...`, `Test`, `Test All`, `About`, `Close`).
- The `Remove...` button toggles **Disable/Enable** for embedded/optional plugins; for custom plugins it removes the plugin from settings (and its configuration).
- The configuration dialog renders descriptions from the JSON schema and shows defaults/min/max under each control using the current theme colors.

The host loads each plugin to query `IInformations` metadata/configuration, applies configuration from settings, and persists configuration on shutdown when `SomethingToSave()` returns `TRUE`.

**Loading Pattern:**
```cpp
// Load plugin DLL from the Plugins directory next to the executable (no PATH fallback).
#include <filesystem>
#include <wil/resource.h>
#include <wil/win32_helpers.h>

wil::unique_cotaskmem_string modulePath;
THROW_IF_FAILED(wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath));

const std::filesystem::path exeDir = std::filesystem::path(modulePath.get()).parent_path();
const std::filesystem::path pluginPath = exeDir / L"Plugins" / L"FileSystem.dll";

wil::unique_hmodule plugin(LoadLibraryW(pluginPath.c_str()));
THROW_LAST_ERROR_IF(!plugin);

auto createFunc = reinterpret_cast<decltype(&RedSalamanderCreate)>(GetProcAddress(plugin.get(), "RedSalamanderCreate"));
if (! createFunc)
{
    DWORD lastError = GetLastError();
    if (lastError == ERROR_SUCCESS)
    {
        lastError = ERROR_PROC_NOT_FOUND;
    }
    entry.loadError = L"Missing export RedSalamanderCreate.";
    return HRESULT_FROM_WIN32(lastError);
}

wil::com_ptr<IFileSystem> fs;
FactoryOptions opts{DEBUG_LEVEL_INFORMATION};
THROW_IF_FAILED(createFunc(__uuidof(IFileSystem), &opts, fs.put_void()));
```

### Reference Implementation (FileSystem.dll)

The `Plugins/FileSystem/` project provides a local file system implementation using Windows native APIs:

**Architecture:**
- `FileSystem` class: Implements `IFileSystem` using native batched directory enumeration when possible:
  - Local volumes: `NtQueryDirectoryFile(FileFullDirectoryInformation)` (fast path)
  - Fallback: `GetFileInformationByHandleEx(FileFullDirectory*Info)`
  - Network/UNC/WSL paths: `FindFirstFileExW`/`FindNextFileW` (compatibility path)
  - UNC server roots (`\\server\`): `NetShareEnum` to list shares as directories
- `FilesInformation` class: Implements `IFilesInformation` with a contiguous `std::vector<std::byte>` backing buffer
- `Factory.cpp`: Exports `RedSalamanderCreate` factory function
- Memory management: Buffer capacity grows progressively for large directories; host never frees the buffer pointer (owned by the COM object)

**Status:** The built-in local filesystem plugin (`Plugins/FileSystem/`) implements directory enumeration, directory watch, and file operations (Copy/Move/Delete/Rename). Search is not implemented.

**Key Implementation Details:**
```cpp
class FileSystem : public IFileSystem {
    HRESULT ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppInfo) override {
        // Normalize to an absolute path and apply \\?\ or \\?\UNC\ prefix for long paths.
        // Enumerate using NtQueryDirectoryFile / GetFileInformationByHandleEx when possible, otherwise FindFirstFileEx / FindNextFile.
        // Special case: "\\\\server\\" enumerates SMB shares via NetShareEnum.
        // Build FileInfo entries into a contiguous buffer (NextEntryOffset chain).
        // Progressive buffer growth: 512KB → 2MB → 8MB → 32MB → ... → 512MB (soft cap), with a hard-cap fallback (currently 2GB).
        // Returns HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) if directory exceeds max capacity.
    }
};

class FilesInformation : public IFilesInformation {
    std::vector<std::byte> _buffer;  // Owns the data
    unsigned long _count = 0;
    unsigned long _usedBytes = 0;
    
    HRESULT GetBuffer(FileInfo** ppInfo) override {
        // Read-only: does not mutate count/usage state
        if (_count == 0 || _usedBytes == 0) {
            *ppInfo = nullptr;
            return S_OK;
        }
        *ppInfo = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

private:
    // Plugin-only writer API (not exposed to host): resets state and writes entries during ReadDirectoryInfo().
};
```

**FileInfo field behavior (reference plugin):**
- `FileNameSize` is set to `wcslen(cFileName) * sizeof(wchar_t)` and a trailing NUL is written (not counted in `FileNameSize`).
- `CreationTime`/`LastAccessTime`/`LastWriteTime` come from file enumeration (FILETIME → int64).
- `ChangeTime` uses the native directory enumeration change-time when available; otherwise it is set to `LastWriteTime`.
- `EndOfFile` is derived from file enumeration; `AllocationSize` uses native allocation size when available, otherwise it is set equal to `EndOfFile`.
- `FileIndex` uses the native file index when available; `EaSize` is set to 0.

**File Traversal:**
The implementation uses linked-list style traversal via `NextEntryOffset`:
```cpp
FileInfo* current = /* from GetBuffer() */;
while (current) {
    // Process current entry
    const size_t nameChars = static_cast<size_t>(current->FileNameSize) / sizeof(wchar_t);
    const std::wstring_view name(current->FileName, nameChars);
    ProcessFile(name, current->EndOfFile);
    
    // Move to next entry
    if (current->NextEntryOffset == 0)
        break;  // Last entry
    current = reinterpret_cast<FileInfo*>(
        reinterpret_cast<uint8_t*>(current) + current->NextEntryOffset
    );
}
```

### Dummy Implementation (FileSystemDummy.dll)

The `Plugins/FileSystemDummy/` project provides a deterministic in-memory file system used for UI and plugin-contract testing.

**Behavior:**
- Deterministic procedural generation driven by `seed` (and `maxChildrenPerDirectory`); enumeration order does not affect the generated structure.
- Path traversal generates intermediate directory contents on-demand, so operations do not require enumerating parent folders first.
- Generated names include ASCII, Latin diacritics, emoji, and non-Latin scripts (Japanese, Arabic, Thai, Korean).
- Optional latency simulation (`latencyMs`) to model slow file systems:
  - Adds delay per `GetAttributes` call.
  - Adds delay per directory entry returned by `ReadDirectoryInfo`.

**Configuration keys (schema-provided):**
- `maxChildrenPerDirectory` (0..20000)
- `maxDepth` (0..1024, `0` = unlimited)
- `seed` (0 = random per run)
- `latencyMs` (0..1000)
- `virtualSpeedLimit` (string, `0` = unlimited; examples: `3KB`, `4MB`)

## Error Handling

All interface methods return `HRESULT` values indicating success or failure:

**Success Codes:**
- `S_OK` (0x00000000): Operation succeeded

**Common Error Codes:**
- `E_POINTER` (0x80004003): Null output pointer
- `E_INVALIDARG` (0x80070057): Invalid parameter value (e.g., empty path)
- `E_OUTOFMEMORY` (0x8007000E): Memory allocation failed
- `E_NOINTERFACE` (0x80004002): Requested interface not supported
- `HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND/ERROR_PATH_NOT_FOUND)`: Directory doesn't exist
- `HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)`: Permission denied
- `HRESULT_FROM_WIN32(ERROR_NOT_READY)`: Drive not ready (removable media)
- `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)`: Operation or flag not supported by the plugin
- `HRESULT_FROM_WIN32(ERROR_CANCELLED)`: Operation canceled by user request
- `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)`: Batch operation completed with one or more failures
- `HRESULT_FROM_WIN32(ERROR_FILE_EXISTS)` / `HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)`: Destination exists and overwrite is not allowed
- `HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER)`: Directory too large for the reference plugin's maximum buffer (currently 2GB)
- `HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES)`: No more entries (e.g., `IFilesInformation::Get()` index out of range)
- `HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW)`: Size overflow (e.g., `GetAllocatedSize()` cannot represent the value in `unsigned long`)

**Error Handling Pattern:**
```cpp
wil::com_ptr<IFilesInformation> info;
HRESULT hr = fileSystem->ReadDirectoryInfo(path, info.put());
if (FAILED(hr)) {
    if (hr == E_POINTER || hr == E_INVALIDARG) {
        // Invalid parameters
    } else if (hr == E_ACCESSDENIED || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)) {
        // Permission denied
    } else if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)) {
        // Directory doesn't exist
    } else {
        // Generic error handling
    }
}
```

**Plugin Responsibilities:**
- Return appropriate HRESULT codes for all error conditions
- Do not throw exceptions across COM boundaries
- Log detailed diagnostics internally (use DebugLevel from FactoryOptions)
- Validate all input parameters before use

## Testing

### Test Requirements

**Unit Tests:**
- Factory function with valid/invalid parameters
- Directory enumeration with empty, small, and large directories
- Copy/move/delete/rename (single and batch) with success and failure cases
- Progress callbacks and cancellation (operations and search)
- Unicode filename support (various languages, emoji, special characters)
- Error conditions (invalid paths, access denied, network timeouts)
- Memory leak detection (verify proper cleanup on Release())
- Buffer traversal (NextEntryOffset linkage)

**Integration Tests:**
- Loading plugin DLL from RedSalamander
- Displaying enumerated files in FolderView
- File operations routed through IFileSystem
- Search results populated via IFileSystemSearch
- Sorting and filtering operations
- Icon resolution from FileInfo attributes
- Performance with 10K+ file directories

**Automated Test Coverage:**
- COM reference counting correctness
- Thread safety (concurrent ReadDirectoryInfo calls)
- Memory boundary conditions (large filenames, deep paths)
- Cross-compiler compatibility (MSVC, Clang)

## Performance Considerations

### Optimization Strategies

**1. Buffer Pre-allocation:**
- Estimate buffer size based on typical directory (average 64 bytes per entry)
- Reduce allocations by reserving capacity upfront
- Example: 1000 files × 64 bytes = 64KB initial allocation

**2. Caching:**
- Plugins MUST NOT implement internal caching of directory listings
- Cache will be manage outside plugin 
- Host application is responsible for caching strategies

**3. Parallel Enumeration:**
- For network file systems, consider parallel queries across subdirectories
- Use thread pool for concurrent directory reads
- Aggregate results before returning to caller

**4. Lazy Loading:**
- Current API loads entire directory upfront
- Future enhancement: Implement streaming API for incremental enumeration
- Useful for massive directories (100K+ files)

**5. Batch Operations:**
- Validate upfront where possible (e.g., destination existence) to fail fast
- Throttle progress callbacks to avoid UI churn
- Prefer sequential I/O on spinning disks; allow parallelism only when safe

**6. Search:**
- Stream matches through callbacks instead of buffering large result sets
- Provide periodic progress updates for long-running searches

**Performance Targets:**
- Local file system: <100ms for 10K files
- Network share: <500ms for 10K files (depends on network latency)
- Memory usage: <10MB for 100K files

## Future Extensions

### Planned Features

**1. File I/O Streams:**
```cpp
interface IFileSystemStreams : public IUnknown {
    HRESULT ReadFile(const wchar_t* path, IStream** ppStream);
    HRESULT WriteFile(const wchar_t* path, IStream* pStream);
};
```

**2. Metadata write operations:**
- `IFileSystemIO` (above) provides fast attribute reads (`GetAttributes`).
- Future versions may add write APIs for attributes/times (either via new methods on a new interface GUID or a separate interface).

**3. Change Notifications (Optional):**

Plugins MAY provide directory change notifications via the optional `IFileSystemDirectoryWatch` interface queried from the active `IFileSystem` instance.

This is used by the host (`DirectoryInfoCache`) to drive watch-based refresh after file system mutations (the built-in `builtin/file-system` and `builtin/file-system-dummy` plugins implement this).

**Contract:**
- `IFileSystemDirectoryWatchCallback` is a raw vtable interface (**NOT COM**). Do not `AddRef/Release` it and do not store it in `wil::com_ptr`.
- `cookie` is host-owned and opaque; plugins MUST pass it back verbatim on every callback.
- Callbacks MAY be invoked on arbitrary background threads; callback implementations MUST be thread-safe and fast.
- `WatchDirectory` watches a single directory (non-recursive).
- `UnwatchDirectory` MUST guarantee no further callbacks for that path after it returns (including in-flight callbacks).
- The callback SHOULD include “what changed”: action + affected entry name/path (relative to the watched folder), and MAY batch multiple changes per call. If changes were dropped/coalesced, set `overflow = TRUE` and `changes` MAY be empty.

```cpp
enum FileSystemDirectoryChangeAction : uint32_t {
    FILESYSTEM_DIR_CHANGE_UNKNOWN = 0,
    FILESYSTEM_DIR_CHANGE_ADDED,
    FILESYSTEM_DIR_CHANGE_REMOVED,
    FILESYSTEM_DIR_CHANGE_MODIFIED,
    FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME,
    FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME,
};

struct FileSystemDirectoryChange {
    FileSystemDirectoryChangeAction action;
    const wchar_t* relativePath;      // not required to be NUL-terminated
    unsigned long relativePathSize;   // bytes (not chars)
};

struct FileSystemDirectoryChangeNotification {
    const wchar_t* watchedPath;       // NUL-terminated
    unsigned long watchedPathSize;    // bytes (not chars)
    const FileSystemDirectoryChange* changes;
    unsigned long changeCount;
    BOOL overflow;                   // TRUE if changes were dropped/coalesced
};

interface __declspec(novtable) IFileSystemDirectoryWatchCallback {
    virtual HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(
        const FileSystemDirectoryChangeNotification* notification,
        void* cookie) noexcept = 0;
};

interface __declspec(uuid("d00f72a2-faf2-47c4-abbe-85dab1e67132")) __declspec(novtable) IFileSystemDirectoryWatch : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE WatchDirectory(
        const wchar_t* path,
        IFileSystemDirectoryWatchCallback* callback,
        void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE UnwatchDirectory(const wchar_t* path) noexcept = 0;
};
```

**4. Streaming Enumeration:**
- Paged or streaming directory enumeration for very large folders
- Incremental host updates without a full in-memory buffer

**5. Virtual File Systems:**
- ZIP/archive browsing
- FTP/SFTP remote file systems
- Cloud storage (OneDrive, Google Drive, Dropbox)
- In-memory file systems for testing

## Documentation

### Plugin Developer Guide

**See:** the reference plugin source in `Plugins/FileSystem/` (`Factory.cpp`, `FileSystem.h`, `FileSystem.cpp`) and the public headers in `Common/PlugInterfaces/`.

### API Reference

**See:** `Common/PlugInterfaces/` for authoritative interface definitions.

**Header File:** All interface definitions, structures, and enums are in `Common/PlugInterfaces/`. This is the contract between plugins and RedSalamander.

**Versioning:** Interfaces use COM GUIDs for identity. Breaking changes require new GUIDs (IFileSystem2, etc.). RedSalamander queries for specific versions via QueryInterface.

## AGENTS.md Compliance

This specification follows Red Salamander development guidelines:

- **C++23 Standard**: Use modern C++ features in implementation
- **RAII Patterns**: All resources managed with RAII (COM objects via wil::com_ptr)
- **Smart Pointers**: Use wil::com_ptr for COM interfaces, std::unique_ptr internally
- **No Raw new/delete**: Use RAII wrappers and STL containers
- **Error Handling**: HRESULT return values, THROW_IF_FAILED for internal errors
- **Unicode UTF-16**: All strings are wchar_t*
- **Thread Safety**: Document thread-safety guarantees per interface
- **Performance**: Optimize for common case (local file system, <10K files)

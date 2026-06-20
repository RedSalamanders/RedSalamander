# Writing a File-System Plugin

This page is a step-by-step guide to authoring a RedSalamander file-system plugin: a COM-style DLL that backs an address-bar prefix (for example `file:`, `s3:`, `ftp:`, or `fk:`) by implementing `IFileSystem`. It complements the [Developer Guide](../DeveloperGuide.md) subsystem deep dive and the user-facing [Plugins](../Plugins.md) page. The normative ABI contract is `Specs/Plugins/Plugins_VirtualFileSystem.md`; this guide shows how to satisfy it with working code patterns drawn from the built-in `builtin/file-system` and `builtin/file-system-dummy` providers.

## Prerequisites

- A C++23 DLL project under `Plugins/` (clone an existing `*.vcxproj`, for example `Plugins/FileSystemDummy/`).
- Include the plugin interfaces from `Common/PlugInterfaces/`:
  - `Factory.h` — the C entry points and `FactoryOptions`.
  - `FileSystem.h` — `IFileSystem`, `IFilesInformation`, `FileInfo`, the callbacks, and the arena helpers.
  - `Informations.h` — `PluginMetaData` and `IInformations`.
- A working knowledge of basic COM lifetime (`IUnknown::QueryInterface`/`AddRef`/`Release`); plugins are not registered with the OS COM runtime — the host loads the DLL and calls the exported factory directly.

Smallest viable provider: a class implementing `IFileSystem` (which includes the mandatory `GetCapabilities`) plus `ReadDirectoryInfo`, and a `Factory.cpp` exposing the three C exports below.

## Step 1 — The Factory.cpp skeleton

Every DLL exports up to three C entry points. `RedSalamanderCreate` is required; the other two are optional but recommended (the built-in providers implement all three). See `Plugins/FileSystemDummy/Factory.cpp` for the reference and `Plugins/FileSystem/Factory.cpp` for a near-identical copy.

| Export | Required | Purpose |
| --- | --- | --- |
| `RedSalamanderCreate` | Yes | Instantiate a provider for a `riid` (`IFileSystem`) and optional `pluginId`. |
| `RedSalamanderEnumeratePlugins` | Optional | Advertise one or more `PluginMetaData` rows before any instance exists. |
| `RedSalamanderGetConfigurationSchema` | Optional | Return the configuration JSON schema without constructing an instance. |

A single DLL may implement several logical plugins (for example `FileSystemCurl` returns FTP/SFTP/SCP/IMAP, `FileSystemS3` returns `s3`/`s3table`). When `RedSalamanderEnumeratePlugins` is present the host registers one entry per returned `PluginMetaData`; when it is absent the host treats the DLL as a single-plugin factory and may call `RedSalamanderCreate` with `pluginId == nullptr`.

```cpp
#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#include "MyFileSystem.h"
#include "Helpers.h"

extern HINSTANCE g_hInstance;

namespace
{
// The PluginMetaData strings MUST live for the lifetime of the loaded DLL.
// Use static/object storage (here: function-local statics), never temporaries.
[[nodiscard]] const PluginMetaData& GetPluginMetaData() noexcept
{
    static const std::wstring name        = LoadStringResource(g_hInstance, IDS_MYFS_NAME);
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_MYFS_DESCRIPTION);
    static const PluginMetaData metaData  = {
        .id          = L"vendor/my-file-system", // stable, non-localized long id
        .shortId     = L"myfs",                   // address-bar scheme prefix
        .name        = name.c_str(),
        .description = description.c_str(),
        .author      = L"Vendor",
        .version     = VERSINFO_PLUGIN_VERSION,
    };
    return metaData;
}

HRESULT CreatePluginInstance(REFIID riid, IHost* host, void** result)
{
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    auto* instance = new (std::nothrow) MyFileSystem(host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release(); // QueryInterface took its own ref on success
    return hr;
}
} // namespace
```

### RedSalamanderEnumeratePlugins

Validate the requested interface, then publish the metadata array. Returned pointers are owned by the DLL and must remain valid until unload.

```cpp
extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    if (! metaData || ! count)
    {
        return E_POINTER;
    }

    *metaData = nullptr;
    *count    = 0;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    *metaData = &GetPluginMetaData(); // point at &array[0] for a multi-plugin DLL
    *count    = 1;
    return S_OK;
}
```

### RedSalamanderCreate

Match `riid`, reject an unknown `pluginId` with `HRESULT_FROM_WIN32(ERROR_NOT_FOUND)`, and dispatch to the right class. The `host` pointer is caller-owned and stays valid for the instance lifetime; the built-in `FileSystem` ignores it while `FileSystemDummy` keeps it to reach `IHostConnections`.

```cpp
extern "C" HRESULT __stdcall RedSalamanderCreate(
    REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }
    if (pluginId && pluginId[0] != L'\0' && ! OrdinalString::EqualsNoCase(pluginId, GetPluginMetaData().id))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return CreatePluginInstance(riid, host, result);
}
```

### RedSalamanderGetConfigurationSchema

Lets the host read the configuration schema during discovery without a live instance. Return the same UTF-8 JSON string your `IInformations::GetConfigurationSchema` returns; the string is DLL-owned and valid until unload.

```cpp
extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    const std::wstring_view requestedId = pluginId ? std::wstring_view(pluginId) : std::wstring_view{};
    if (! requestedId.empty() && ! OrdinalString::EqualsNoCase(requestedId, GetPluginMetaData().id))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    *schemaJsonUtf8 = MyFileSystem::StaticConfigurationSchema();
    return S_OK;
}
```

### Optional module quiet-point exports

Plugins that own DLL-global schedulers, caches, window classes, or driver-backed resources may also export the module quiet-point hooks the host discovers with `GetProcAddress`:

```cpp
PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept;
PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept;
```

`RedSalamanderPluginShutdown` must be idempotent and non-throwing; after it returns no DLL-global worker may call host callbacks or touch state the host releases before `FreeLibrary`. It is a final module quiet point, not a substitute for releasing live instances. `RedSalamanderPluginRetainModuleUntilProcessExit` is honored only at process shutdown: returning `TRUE` keeps the DLL mapped for OS teardown when explicit `FreeLibrary` would race driver cleanup. Most providers omit both. See `Specs/Plugins/Plugins_PluginAPI.md` for the unload contract.

## Step 2 — Implement IFileSystem and QueryInterface

Derive your class from `IFileSystem` and any optional interfaces you support, then route them through `QueryInterface`. The host discovers what each *instance* supports by querying for the optional interfaces — there is no separate registration. Pattern from `FileSystemDummy`:

```cpp
class MyFileSystem final : public IFileSystem,
                           public IFileSystemIO,
                           public IFileSystemDirectoryOperations,
                           public IInformations
{ /* ... */ };

HRESULT STDMETHODCALLTYPE MyFileSystem::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
    {
        *ppvObject = static_cast<IFileSystem*>(this);
        AddRef();
        return S_OK;
    }
    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }
    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }
    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}
```

`AddRef`/`Release` are a simple `std::atomic_ulong` refcount that `delete this` at zero. COM objects are non-copyable/non-movable: explicitly delete the copy/move members.

### Which interfaces to implement

`IFileSystem` is mandatory and already carries directory enumeration (`ReadDirectoryInfo`), the copy/move/delete/rename operations, the mandatory `GetCapabilities`, and the `GetTransferHints`/`GetStorageCharacteristics` tuning hints. The remaining interfaces are optional; the host gracefully degrades when `QueryInterface` returns `E_NOINTERFACE`.

| Interface (UUID in `FileSystem.h`) | When to implement |
| --- | --- |
| `IFileSystem` | Always. Enumeration, mutations, capabilities. |
| `IFileSystemIO` | Reading/writing item content (`CreateFileReader`/`CreateFileWriter`), attributes, basic info, and `GetItemProperties`. Required for the cross-file-system bridge to use this provider as a source or destination. |
| `IFileSystemDirectoryOperations` | `CreateDirectory` and recursive `GetDirectorySize`. |
| `IFileSystemItemStreams` | Deleting named alternate streams (for example `Zone.Identifier`). |
| `IFileSystemDirectoryWatch` | Live change notifications via `WatchDirectory`/`UnwatchDirectory`. |
| `IFileSystemSearch` | Name/content search via `Search` (otherwise the host falls back to its own recursive scan). |
| `IFileSystemInitialize` | Per-instance binding to a root context plus a JSON options payload — used by `FileSystem7z` to mount `7z:C:\x.zip|/`. |
| `IInformations` | Metadata, configuration schema, and config get/set round-tripping. |
| `IFileReader` / `IFileWriter` | The read-only reader and write-then-`Commit` writer objects your `IFileSystemIO` hands back. |

Two host-driven callback interfaces are passed *into* your operations rather than implemented on the provider: `IFileSystemCallback` (copy/move/delete/rename progress, cancel, and `FileSystemIssue` conflict prompts) and `IFileSystemDirectorySizeCallback`. These are not COM interfaces (no `IUnknown`); the host owns their lifetime and your operation must not invoke them after it returns, must not call them concurrently for one operation, and parallel workers must each report a distinct stable `progressStreamId`.

## Step 3 — Pack directory entries into FilesInformation

`ReadDirectoryInfo` returns an `IFilesInformation` that owns one contiguous buffer of `FileInfo` records linked by `NextEntryOffset`. The built-in `FileSystem` streams Win32 `FindFirstFileW`/`GetFileInformationByHandleEx` results straight into the buffer; remote, cloud, and archive providers gather a `std::vector` of entries and pack them with a `BuildFromEntries`-style helper. The packing contract is the load-bearing part — get it wrong and enumeration walks off the buffer.

Rules for the packed buffer (from `FileInfo` in `FileSystem.h`):

- Each entry starts at a record whose `FileName` is variable-length; size a record as `offsetof(FileInfo, FileName) + FileNameSize + sizeof(wchar_t)` (room for a trailing NUL) and align it (the dummy provider uses 8-byte alignment).
- `FileNameSize` is the name length in **bytes, not characters**, and the name is **not required to be NUL-terminated** — consumers must use `FileNameSize`. Writing the NUL is courtesy, but the size field is authoritative.
- `NextEntryOffset` is the byte distance from this record to the next. Set it to the record size for every entry **except the last**, where it must be `0` to terminate the chain.
- After the buffer is returned it is immutable: never mutate or reallocate it (no async mutation after `ReadDirectoryInfo` returns).

The packer from `Plugins/FileSystemDummy/FileSystemDummy.cpp` (`BuildFileInfoBuffer`):

```cpp
const size_t baseSize = offsetof(FileInfo, FileName);

// First pass: total the aligned record sizes (watch for overflow).
size_t totalBytes = 0;
for (const auto& entry : entries)
{
    const size_t nameBytes = entry.name.size() * sizeof(wchar_t);
    totalBytes += AlignUp(baseSize + nameBytes + sizeof(wchar_t), kEntryAlignment);
}

outBuffer->assign(totalBytes, std::byte{0});

// Second pass: write each record and link it.
std::byte* base = outBuffer->data();
size_t offset   = 0;
for (size_t index = 0; index < entries.size(); ++index)
{
    const auto& entry      = entries[index];
    const size_t nameBytes = entry.name.size() * sizeof(wchar_t);
    const size_t entrySize = AlignUp(baseSize + nameBytes + sizeof(wchar_t), kEntryAlignment);

    auto* info = reinterpret_cast<FileInfo*>(base + offset);
    ZeroMemory(info, entrySize);

    info->FileIndex      = static_cast<unsigned long>(index);
    info->FileAttributes = entry.attributes; // FILE_ATTRIBUTE_* (set FILE_ATTRIBUTE_DIRECTORY for folders)
    info->FileNameSize   = static_cast<unsigned long>(nameBytes);
    info->CreationTime   = entry.creationTime; // FILETIME ticks
    info->LastWriteTime  = entry.lastWriteTime;
    info->EndOfFile      = static_cast<__int64>(entry.sizeBytes);
    info->AllocationSize = /* size rounded up to cluster, or 0 if unknown */;

    if (nameBytes > 0)
    {
        CopyMemory(info->FileName, entry.name.data(), nameBytes);
    }
    info->FileName[entry.name.size()] = L'\0';

    if (index + 1 < entries.size())
    {
        info->NextEntryOffset = static_cast<unsigned long>(entrySize); // 0 on the last entry
    }

    offset += entrySize;
}
```

Wrap that buffer in an `IFilesInformation` implementation that returns it from `GetBuffer`/`GetBufferSize` and exposes `GetCount`/`Get` helpers. `DummyFilesInformation` in the dummy plugin is a minimal reference: it stores the `std::vector<std::byte>`, the committed byte count, and the entry count, and its `QueryInterface` answers only `IUnknown` and `IFilesInformation`. When a directory is empty, return `S_OK` with `*ppFileInfo == nullptr`.

`ReadDirectoryInfo` itself validates the path, resolves it, builds the entry list under any internal lock, releases the lock, then packs and constructs the `IFilesInformation` (so the COM object is built outside the lock).

### Building path lists for batch operations

When you need a flat list of full child paths from an `IFilesInformation` (for batch copy/move/delete), `FileSystem.h` provides `FileSystemArenaOwner::BuildPathListFromFilesInformation`, which allocates a `const wchar_t*[]` plus the joined `<root>\<name>` strings from a single arena. Arrays passed to `CopyItems`/`MoveItems`/`DeleteItems` and arrays of `FileSystemRenamePair` must be allocated from the same arena as their referenced UTF-16 strings, and arena strings are NUL-terminated. Use `FileSystemArenaOwner` (RAII) rather than the raw `InitializeFileSystemArena`/`DestroyFileSystemArena` pair.

## Step 4 — The capabilities JSON contract

`IFileSystem::GetCapabilities` is **mandatory** and **fail-closed**. It must return `S_OK` and a non-empty UTF-8, NUL-terminated JSON document. Returning `E_NOTIMPL`, `ERROR_NOT_SUPPORTED`, a null pointer, or an empty/unparseable document is a contract violation: the host disables capability-gated operations (same-provider Copy/Move/Delete/Batch Rename are rejected before a task is created) and logs the violation once per instance. Advertise *unsupported* actions with `false` operation fields or empty import/export lists — never by failing the call.

`pathIdentity` is mandatory provider capability data. Identity-sensitive planners such as Batch Rename and same-context Copy/Move/Delete reject missing, malformed, unsupported, or unstable `pathIdentity` before mutation.

Capabilities are **per instance**. They may vary by mode (S3 vs S3Tables), transport, or connection profile, so the host re-queries each instance and never assumes a DLL has one fixed capability set. The returned pointer is plugin-owned and valid until the next `GetCapabilities` call or release; build it once and cache it (the built-in `FileSystem` lazily formats `_capabilitiesJson` under its state lock and rebuilds it when configuration changes).

The host-recognized document (version 1) requires `operations`, `concurrency`, `crossFileSystem`, and `pathIdentity`:

```json
{
  "version": 1,
  "operations": {
    "copy": true,
    "move": true,
    "delete": true,
    "rename": true,
    "properties": true,
    "read": true,
    "write": true
  },
  "concurrency": {
    "copyMoveMax": 4,
    "deleteMax": 8,
    "deleteRecycleBinMax": 2
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalIgnoreCase",
    "normalization": "none",
    "preferredSeparator": "\\",
    "acceptedSeparators": ["\\", "/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
```

### operations

| Field | Meaning |
| --- | --- |
| `copy` / `move` / `delete` / `rename` | Whether same-provider mutations are offered. Read-only providers (7-Zip, S3 Table, Google Drive) set these `false`. |
| `properties` | `IFileSystemIO::GetItemProperties` is available. |
| `read` | Supports enumeration plus `IFileReader`. |
| `write` | Supports `IFileWriter` via `IFileSystemIO::CreateFileWriter`. |

### concurrency

Numeric worker budgets the host honors: `copyMoveMax` for copy/move fan-out, `deleteMax` for permanent-delete batches, and `deleteRecycleBinMax` for Recycle Bin deletes. Clamp these to sane ranges; the built-in `FileSystem` derives them from its configuration.

### crossFileSystem

Opt-in allow-lists for the host bridge that copies/moves between two different providers. Plugin ids are long-form `PluginMetaData.id` values, compared case-insensitively; `"*"` means any id. `export` lists the destination ids this provider may be the **source** for; `import` lists the source ids this provider may be the **destination** for. An empty list blocks that direction. The host pairs both sides' lists, so a transfer is allowed only when the source's `export` and the destination's `import` both permit it. The bridge itself reads through `IFileSystemIO::CreateFileReader` and writes through `CreateFileWriter`, so `read`/`write` must also be true for the relevant sides.

### pathIdentity

`pathIdentity` tells the host's collision/dependency/refresh planners how to decide whether two plugin paths name the same item *inside one instance*. The host parses it into one profile per instance and reuses that profile for every identity decision — your provider must not embed ad hoc string comparisons elsewhere.

| Field | Rule |
| --- | --- |
| `version` | Must be `1`. |
| `pathTextStableIdentity` | `true` only when canonical plugin path text uniquely identifies one item. Providers that allow duplicate display names in one parent (for example Google Drive without stable item IDs) must set `false`. |
| `componentComparison` | Exactly one of `ordinalCaseSensitive` (UTF-16 code-unit equality, no folding) or `ordinalIgnoreCase` (Windows ordinal case-insensitive, accent- and normalization-sensitive). `unknown` is **not** a legal plugin value — a provider that cannot prove its relation must still declare the conservative concrete default `ordinalCaseSensitive` (data-safe: an existing destination fails with `ERROR_ALREADY_EXISTS` rather than being silently overwritten). |
| `normalization` | Must be `none` for v1. |
| `preferredSeparator` | The separator you emit in canonical paths (`"\\"` for local Windows-like providers, `"/"` for remote/object/archive). |
| `acceptedSeparators` | Separators you accept as equivalent when parsing (local providers usually accept both `"\\"` and `"/"`; URL/object providers usually only `"/"`). |
| `casePreserving` | `true` when display casing is preserved even if identity is case-insensitive. |
| `caseOnlyRename` | One of `supported`, `noOp`, `unsupported`, `notApplicable`. |

Providers that support same-provider rename must keep `RenameItem`/`RenameItems` collision behavior consistent with `componentComparison`: without overwrite flags, a destination that already identifies an existing sibling must fail rather than silently replace it. The built-in provider declarations (local, dummy, 7-Zip, the drives, curl protocols, S3) are tabulated in `Specs/Plugins/Plugins_VirtualFileSystem.md` under "Provider path identity".

If your search support is meaningful, you may also add a `search` object (name/content/indexed/regex/snippet flags and a preferred backend) as the built-in `FileSystem` does; it is not part of the required core shape.

## Step 5 — Configuration schema and IInformations

If your provider has settings, implement `IInformations` and return a configuration schema (the same JSON your `RedSalamanderGetConfigurationSchema` export hands back) describing fields for the host's dynamic settings dialog. The host round-trips configuration through `SetConfiguration`/`GetConfiguration` and persists it; `SomethingToSave` reports whether the live config differs from the last saved snapshot. Cloud providers double-buffer the configuration JSON so previously returned pointers stay valid across a `SetConfiguration`. A minimal schema looks like:

```json
{
  "version": 1,
  "title": "My File System",
  "fields": [
    { "key": "timeoutMs", "type": "value", "label": "Timeout (ms)", "default": 5000, "min": 0, "max": 60000 }
  ]
}
```

## Threading and lifetime rules

- All `IFileSystem` methods run on host worker threads, never the UI thread.
- Per-call callbacks (`IFileSystemCallback`, `IFileSystemDirectorySizeCallback`) may block (host-driven pause/conflict UI) and may run on background threads. Do not hold a lock that the callback path could deadlock on, reach progress checkpoints often enough for responsive cancel, and never invoke a callback after the operation returns.
- Directory-vs-directory existence during copy/move is a **merge**, not an `ERROR_ALREADY_EXISTS` conflict; `ERROR_ALREADY_EXISTS` is reserved for file-vs-file collisions and type mismatches. Overwrite/ReplaceReadOnly grants from a child conflict are one-shot and must not leak to siblings unless the host's "Apply to all" toggle broadens them.
- Watch callbacks must use `PostMessage` / thread-pool submission, never `SendMessage`. `UnwatchDirectory` must synchronously drain in-flight callbacks without holding the lock that delivery acquires.
- Validate any `[in]`/`[out]` struct's `sizeBytes` before reading other fields; mismatched `sizeBytes` is a contract violation that should fail with `E_INVALIDARG`.

## Build, register, and test

- Build with the repo wrapper, for example `.\build.ps1 -ProjectName FileSystemDummy`; output lands in `.build\<Platform>\<Configuration>\` next to the host exe under `Plugins\`.
- The host discovers DLLs in the `Plugins\` folder beside `RedSalamander.exe` and any custom paths added in Preferences; it calls `RedSalamanderEnumeratePlugins`, then `RedSalamanderCreate`, validates the reported `id`/`shortId` against enumeration, and rejects duplicate ids/short ids. See [Plugins](../Plugins.md) for the plugin manager UI and the **Test** / **Test All** validation actions.
- `FileSystemDummy` (`fk:`) is the deterministic in-memory provider used by selftests; mirror its structure when you need a reproducible target. File-operation behavior is exercised by `--fileops-selftest` cases (see `Specs/FileSystem/FileSystem_FileOperations.md`), which are mandatory for any change touching liveness or data safety.

## See also

- [Developer Guide](../DeveloperGuide.md) — "File-System Plugins" and "Plugin Host Model & the Cross-File-System Bridge" deep dives.
- [Plugins](../Plugins.md) — user-facing plugin list, prefixes, and the plugin manager UI.
- `Specs/Plugins/Plugins_VirtualFileSystem.md` — the normative ABI, capability JSON, and `pathIdentity` contract.
- `Specs/Plugins/Plugins_PluginAPI.md` — host-services ABI evolution and the module quiet-point exports.
- `Common/PlugInterfaces/Factory.h`, `FileSystem.h`, `Informations.h` — the interface headers themselves.

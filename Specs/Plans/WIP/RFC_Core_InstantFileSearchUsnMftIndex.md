# Instant File Search via NTFS MFT / USN Journal Indexing

## Overview

Index every file on an NTFS volume in seconds using the USN change journal, then keep the index live with real-time journal monitoring. This is the technique used by **Everything**, **Indexer++**, and similar tools.

**Key insight:** never walk the directory tree. Read the MFT directly via `FSCTL_ENUM_USN_DATA`, then keep it updated via `FSCTL_READ_USN_JOURNAL`.

## File System Compatibility

| Capability | NTFS | ReFS | FAT/exFAT |
|------------|------|------|-----------|
| `FSCTL_ENUM_USN_DATA` (MFT dump) | **Yes** | **No** — no MFT | No |
| `FSCTL_READ_USN_JOURNAL` (live changes) | **Yes** | **Yes** | No |
| `FSCTL_QUERY_USN_JOURNAL` | **Yes** | **Yes** | No |
| `FSCTL_CREATE_USN_JOURNAL` | **Yes** | **Yes** | No |
| `NtQueryDirectoryFile` traversal | **Yes** | **Yes** | Yes |

`FSCTL_ENUM_USN_DATA` reads MFT records directly — an NTFS-specific on-disk structure. ReFS uses B+ trees internally and has no MFT, so this IOCTL fails on ReFS volumes. See the [ReFS Fallback Strategy](#refs-fallback-strategy-for-initial-index) section for the equivalent approach.

## Architecture — NTFS

| Phase | API | Speed | Requires Admin |
|-------|-----|-------|----------------|
| **Initial index** | `FSCTL_ENUM_USN_DATA` | ~2-5s for 1M files | Yes |
| **Live updates** | `FSCTL_READ_USN_JOURNAL` | Real-time (blocks until change) | Yes |
| **Search** | In-memory map lookup | Instant (< 1ms) | No |

## Architecture — ReFS

| Phase | API | Speed | Requires Admin |
|-------|-----|-------|----------------|
| **Initial index** | `NtQueryDirectoryFile` recursive traversal | ~10-30s for 1M files | No (but admin for volume root) |
| **Live updates** | `FSCTL_READ_USN_JOURNAL` | Real-time (blocks until change) | Yes |
| **Search** | In-memory map lookup | Instant (< 1ms) | No |

---

## Phase 1: Initial Index — `FSCTL_ENUM_USN_DATA` (NTFS only)

Enumerate every MFT record on the volume. Each `USN_RECORD_V2` gives:
- `FileReferenceNumber` — unique file ID (FRN)
- `ParentFileReferenceNumber` — parent directory FRN (reconstruct full paths)
- `FileName` — just the entry name, not the full path
- `FileAttributes` — directory, hidden, reparse, etc.

### Sample

```cpp
#include <windows.h>
#include <winioctl.h>
#include <vector>
#include <unordered_map>
#include <string>
#include <memory>

struct FileEntry {
    DWORDLONG id;               // File Reference Number (unique)
    DWORDLONG parentId;         // Parent FRN → reconstruct paths
    std::wstring name;
    DWORD attributes;
};

// Requires admin / elevated privileges.
// Returns a flat map of every file + directory on the volume.
std::unordered_map<DWORDLONG, FileEntry> IndexVolume(wchar_t driveLetter)
{
    // 1. Open volume handle
    wchar_t volumePath[] = L"\\\\.\\X:";
    volumePath[4] = driveLetter;

    HANDLE hVolume = CreateFileW(volumePath,
        GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr, OPEN_EXISTING, 0, nullptr);
    if (hVolume == INVALID_HANDLE_VALUE) return {};

    // 2. Query USN journal (must exist; create with FSCTL_CREATE_USN_JOURNAL if needed)
    USN_JOURNAL_DATA_V0 journalData{};
    DWORD bytesReturned = 0;
    DeviceIoControl(hVolume, FSCTL_QUERY_USN_JOURNAL,
        nullptr, 0, &journalData, sizeof(journalData), &bytesReturned, nullptr);

    // 3. Prepare MFT_ENUM_DATA to enumerate all records
    MFT_ENUM_DATA_V0 enumData{};
    enumData.StartFileReferenceNumber = 0;
    enumData.LowUsn  = 0;
    enumData.HighUsn = journalData.MaxUsn;

    // 4. Read in a loop — each call returns a batch of USN_RECORD entries
    constexpr DWORD kBufSize = 512 * 1024;  // 512 KB buffer
    auto buffer = std::make_unique<BYTE[]>(kBufSize);

    std::unordered_map<DWORDLONG, FileEntry> index;
    index.reserve(1'000'000);  // typical disk

    while (DeviceIoControl(hVolume, FSCTL_ENUM_USN_DATA,
               &enumData, sizeof(enumData),
               buffer.get(), kBufSize,
               &bytesReturned, nullptr))
    {
        // First 8 bytes = next StartFileReferenceNumber
        enumData.StartFileReferenceNumber = *reinterpret_cast<DWORDLONG*>(buffer.get());

        auto* record = reinterpret_cast<USN_RECORD_V2*>(buffer.get() + sizeof(USN));
        auto* end    = reinterpret_cast<USN_RECORD_V2*>(buffer.get() + bytesReturned);

        while (record < end)
        {
            std::wstring_view name(record->FileName,
                                   record->FileNameLength / sizeof(wchar_t));

            FileEntry entry;
            entry.id         = record->FileReferenceNumber;
            entry.parentId   = record->ParentFileReferenceNumber;
            entry.name       = std::wstring(name);
            entry.attributes = record->FileAttributes;

            index[entry.id] = std::move(entry);

            record = reinterpret_cast<USN_RECORD_V2*>(
                reinterpret_cast<BYTE*>(record) + record->RecordLength);
        }
    }

    CloseHandle(hVolume);
    return index;
}
```

### Why it's fast

- Reads the MFT **linearly** (sequential I/O) — no directory tree traversal
- ~100x faster than recursive `FindFirstFileEx`
- A 512 KB buffer returns thousands of records per `DeviceIoControl` call
- Typical result: 1M files indexed in **2-5 seconds**

---

## ReFS Fallback Strategy for Initial Index

Since `FSCTL_ENUM_USN_DATA` does **not work on ReFS**, the initial index must be built via directory traversal. The fastest approach is `NtQueryDirectoryFile` with relative handle opens (same as `PoC/ls3`):

### Recommended: `NtQueryDirectoryFile` Recursive Traversal

```cpp
// Detect volume file system to choose strategy
std::wstring GetVolumeFileSystem(wchar_t driveLetter)
{
    wchar_t rootPath[] = L"X:\\";
    rootPath[0] = driveLetter;
    wchar_t fsName[MAX_PATH]{};
    GetVolumeInformationW(rootPath, nullptr, 0, nullptr, nullptr, nullptr,
                          fsName, MAX_PATH);
    return fsName;  // Returns L"NTFS", L"ReFS", L"FAT32", etc.
}

// Choose indexing strategy based on file system
std::unordered_map<DWORDLONG, FileEntry> IndexVolumeAuto(wchar_t driveLetter)
{
    auto fsType = GetVolumeFileSystem(driveLetter);

    if (fsType == L"NTFS")
    {
        // Fast path: MFT enumeration (~2-5s for 1M files)
        return IndexVolume(driveLetter);
    }
    else if (fsType == L"ReFS")
    {
        // Fallback: NtQueryDirectoryFile recursive traversal (~10-30s for 1M files)
        return IndexVolumeViaTraversal(driveLetter);
    }
    else
    {
        // FAT/exFAT: no USN journal support, traversal-only, no live updates
        return IndexVolumeViaTraversal(driveLetter);
    }
}
```

### ReFS Traversal with `NtQueryDirectoryFile`

Use the `NtQueryDirectoryFile` approach from `PoC/ls3` to build the initial index:

```cpp
#include <winternl.h>

// Forward declarations (loaded from ntdll.dll)
using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(
    HANDLE, HANDLE, PIO_APC_ROUTINE, PVOID, PIO_STATUS_BLOCK,
    PVOID, ULONG, FILE_INFORMATION_CLASS, BOOLEAN, PUNICODE_STRING, BOOLEAN);

using NtOpenFile_t = NTSTATUS(NTAPI*)(
    PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES, PIO_STATUS_BLOCK,
    ULONG, ULONG);

std::unordered_map<DWORDLONG, FileEntry> IndexVolumeViaTraversal(wchar_t driveLetter)
{
    // Load NT functions from ntdll.dll
    auto ntdll = GetModuleHandleW(L"ntdll.dll");
    auto NtQueryDirectoryFile = reinterpret_cast<NtQueryDirectoryFile_t>(
        GetProcAddress(ntdll, "NtQueryDirectoryFile"));
    auto NtOpenFile = reinterpret_cast<NtOpenFile_t>(
        GetProcAddress(ntdll, "NtOpenFile"));

    std::unordered_map<DWORDLONG, FileEntry> index;
    index.reserve(1'000'000);

    // 512 KB enumeration buffer
    constexpr size_t kBufSize = 512 * 1024;
    auto buffer = std::make_unique<BYTE[]>(kBufSize);

    DWORDLONG nextId = 1;  // Synthetic IDs (ReFS FRNs aren't as stable as NTFS)

    struct DirWork {
        std::wstring path;
        DWORDLONG id;
    };

    // Start from root
    wchar_t rootPath[] = L"X:\\";
    rootPath[0] = driveLetter;

    std::vector<DirWork> stack;
    stack.push_back({rootPath, 0});

    while (!stack.empty())
    {
        auto [dirPath, parentId] = std::move(stack.back());
        stack.pop_back();

        HANDLE hDir = CreateFileW(dirPath.c_str(),
            FILE_LIST_DIRECTORY, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            nullptr, OPEN_EXISTING,
            FILE_FLAG_BACKUP_SEMANTICS, nullptr);
        if (hDir == INVALID_HANDLE_VALUE) continue;

        IO_STATUS_BLOCK iosb{};
        BOOLEAN firstQuery = TRUE;

        while (true)
        {
            NTSTATUS status = NtQueryDirectoryFile(
                hDir, nullptr, nullptr, nullptr, &iosb,
                buffer.get(), static_cast<ULONG>(kBufSize),
                FileIdBothDirectoryInformation,  // gives FileId on ReFS
                FALSE, nullptr, firstQuery);
            firstQuery = FALSE;

            if (status != 0) break;  // STATUS_NO_MORE_FILES or error

            auto* info = reinterpret_cast<FILE_ID_BOTH_DIR_INFORMATION*>(buffer.get());
            while (true)
            {
                std::wstring_view name(info->FileName,
                                       info->FileNameLength / sizeof(wchar_t));

                if (name != L"." && name != L"..")
                {
                    DWORDLONG fileId = static_cast<DWORDLONG>(info->FileId.QuadPart);
                    if (fileId == 0) fileId = nextId++;  // fallback synthetic ID

                    FileEntry entry;
                    entry.id         = fileId;
                    entry.parentId   = parentId;
                    entry.name       = std::wstring(name);
                    entry.attributes = info->FileAttributes;
                    index[entry.id]  = std::move(entry);

                    if (info->FileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                    {
                        std::wstring childPath = dirPath;
                        if (!childPath.empty() && childPath.back() != L'\\')
                            childPath += L'\\';
                        childPath.append(name);
                        stack.push_back({std::move(childPath), fileId});
                    }
                }

                if (info->NextEntryOffset == 0) break;
                info = reinterpret_cast<FILE_ID_BOTH_DIR_INFORMATION*>(
                    reinterpret_cast<BYTE*>(info) + info->NextEntryOffset);
            }
        }

        CloseHandle(hDir);
    }

    return index;
}
```

### Performance Comparison: Initial Index

| Method | File System | ~1M files | Notes |
|--------|-------------|-----------|-------|
| `FSCTL_ENUM_USN_DATA` | NTFS only | **2-5 sec** | Sequential MFT read, fastest possible |
| `NtQueryDirectoryFile` recursive | NTFS & ReFS | **10-30 sec** | Directory-by-directory, still very fast |
| `FindFirstFileExW` recursive | Any | **30-120 sec** | Win32 overhead, slowest |

### Hybrid Approach for ReFS

After the initial traversal on ReFS, switch to `FSCTL_READ_USN_JOURNAL` for live updates. The USN journal on ReFS works identically to NTFS for change monitoring — the only limitation is the initial bulk enumeration.

---

## Phase 2: Live Updates — `FSCTL_READ_USN_JOURNAL` (NTFS & ReFS)

After the initial index is built, monitor the USN change journal so you never have to re-scan:

```cpp
void WatchChanges(HANDLE hVolume, USN startUsn, DWORDLONG journalId,
                  std::unordered_map<DWORDLONG, FileEntry>& index)
{
    READ_USN_JOURNAL_DATA_V0 readData{};
    readData.StartUsn       = startUsn;
    readData.ReasonMask     = USN_REASON_FILE_CREATE | USN_REASON_FILE_DELETE
                            | USN_REASON_RENAME_NEW_NAME | USN_REASON_CLOSE;
    readData.ReturnOnlyOnClose = FALSE;
    readData.Timeout        = 0;
    readData.BytesToWaitFor = 1;          // block until at least 1 byte returned
    readData.UsnJournalID   = journalId;

    constexpr DWORD kBufSize = 64 * 1024;
    auto buffer = std::make_unique<BYTE[]>(kBufSize);
    DWORD bytesReturned = 0;

    while (DeviceIoControl(hVolume, FSCTL_READ_USN_JOURNAL,
               &readData, sizeof(readData),
               buffer.get(), kBufSize,
               &bytesReturned, nullptr))
    {
        auto nextUsn = *reinterpret_cast<USN*>(buffer.get());
        auto* record = reinterpret_cast<USN_RECORD_V2*>(buffer.get() + sizeof(USN));
        auto* end    = reinterpret_cast<USN_RECORD_V2*>(buffer.get() + bytesReturned);

        while (record < end)
        {
            if (record->Reason & USN_REASON_FILE_DELETE)
            {
                index.erase(record->FileReferenceNumber);
            }
            else if (record->Reason & (USN_REASON_FILE_CREATE | USN_REASON_RENAME_NEW_NAME))
            {
                std::wstring_view name(record->FileName,
                                       record->FileNameLength / sizeof(wchar_t));
                FileEntry entry;
                entry.id         = record->FileReferenceNumber;
                entry.parentId   = record->ParentFileReferenceNumber;
                entry.name       = std::wstring(name);
                entry.attributes = record->FileAttributes;
                index[entry.id]  = std::move(entry);
            }

            record = reinterpret_cast<USN_RECORD_V2*>(
                reinterpret_cast<BYTE*>(record) + record->RecordLength);
        }

        readData.StartUsn = nextUsn;
    }
}
```

### Notes

- `BytesToWaitFor = 1` makes the call **block** until at least one change occurs (ideal for a background thread)
- Filter with `ReasonMask` to only receive the change types you care about
- The `StartUsn` returned by `FSCTL_ENUM_USN_DATA` (the journal's `NextUsn`) is the starting point for watching

---

## Phase 3: Search the Index

Once indexed, searching by name is instant — just iterate the in-memory map:

```cpp
// Reconstruct full path by walking parent chain
std::wstring ReconstructFullPath(
    const std::unordered_map<DWORDLONG, FileEntry>& index,
    DWORDLONG id)
{
    std::vector<std::wstring_view> parts;
    DWORDLONG current = id;

    while (true)
    {
        auto it = index.find(current);
        if (it == index.end()) break;

        parts.push_back(it->second.name);

        if (it->second.parentId == current) break;  // root
        current = it->second.parentId;
    }

    std::wstring path;
    for (auto it = parts.rbegin(); it != parts.rend(); ++it)
    {
        if (!path.empty()) path += L'\\';
        path.append(*it);
    }
    return path;
}

// Case-insensitive substring search
std::vector<std::wstring> FindFiles(
    const std::unordered_map<DWORDLONG, FileEntry>& index,
    std::wstring_view query)
{
    std::vector<std::wstring> results;
    for (const auto& [id, entry] : index)
    {
        if (entry.name.size() >= query.size())
        {
            auto it = std::search(entry.name.begin(), entry.name.end(),
                                  query.begin(), query.end(),
                                  [](wchar_t a, wchar_t b) {
                                      return towlower(a) == towlower(b);
                                  });
            if (it != entry.name.end())
            {
                results.push_back(ReconstructFullPath(index, id));
            }
        }
    }
    return results;
}
```

---

## Important Considerations

### Requirements
- **NTFS**: Full support (`FSCTL_ENUM_USN_DATA` + `FSCTL_READ_USN_JOURNAL`)
- **ReFS**: Partial — `FSCTL_READ_USN_JOURNAL` works, but `FSCTL_ENUM_USN_DATA` does **not** (no MFT). Use `NtQueryDirectoryFile` traversal for initial index.
- **FAT/exFAT**: No USN journal at all — traversal only, no live updates
- **Administrator / elevated privileges** required to open `\\.\.X:` and call USN journal IOCTLs
- USN journal must exist — create one with `FSCTL_CREATE_USN_JOURNAL` if needed

### Edge Cases
- **Reparse points / junctions:** `FileAttributes` includes `FILE_ATTRIBUTE_REPARSE_POINT`; decide whether to follow
- **Hard links:** A file can have multiple MFT entries pointing to the same data; `FileReferenceNumber` is unique per name
- **Journal wrap:** If the journal overflows, old records are lost. Detect via `FSCTL_QUERY_USN_JOURNAL` and fall back to full re-index
- **Alternate data streams:** Not included in `USN_RECORD_V2` enumeration by default

### Memory Usage
- ~100-200 bytes per entry (FRN, parent FRN, name string)
- 1M files ≈ 100-200 MB of RAM
- Can be reduced by storing names in a compact string pool

### Performance Optimization
- Use a **trie** or **suffix array** for sub-millisecond substring search on very large indexes
- Store the index to disk (serialized) so cold start doesn't require a full MFT/traversal scan
- Run the watcher on a dedicated background thread with `std::jthread` + `stop_token`
- On ReFS: serialize the index aggressively, since re-traversal (~10-30s) is much slower than NTFS MFT re-read (~2-5s)

---

## Reference Projects

| Project | URL | Notes |
|---------|-----|-------|
| **Everything** | https://www.voidtools.com/ | Gold standard; closed-source but has SDK for IPC queries |
| **Indexer++** | https://github.com/dfs-minded/indexer-plus-plus | Full C++ open-source: MFT reading + USN watching + raw MFT parser |
| **Everything SDK** | https://www.voidtools.com/support/everything/sdk/ | Query Everything's index via IPC (no admin needed if Everything is running) |

## Microsoft Documentation

- [FSCTL_ENUM_USN_DATA](https://learn.microsoft.com/windows/win32/api/winioctl/ni-winioctl-fsctl_enum_usn_data) — NTFS only
- [FSCTL_READ_USN_JOURNAL](https://learn.microsoft.com/windows/win32/api/winioctl/ni-winioctl-fsctl_read_usn_journal) — NTFS & ReFS
- [Walking a Buffer of Change Journal Records](https://learn.microsoft.com/windows/win32/fileio/walking-a-buffer-of-change-journal-records)
- [Using the Change Journal Identifier](https://learn.microsoft.com/windows/win32/fileio/using-the-change-journal-identifier)
- [MFT_ENUM_DATA_V0](https://learn.microsoft.com/windows/win32/api/winioctl/ns-winioctl-mft_enum_data_v0)
- [USN_RECORD_V2](https://learn.microsoft.com/windows/win32/api/winioctl/ns-winioctl-usn_record_v2)
- [ReFS Overview — Feature Comparison](https://learn.microsoft.com/windows-server/storage/refs/refs-overview#feature-comparison) — confirms USN journal ✅ on ReFS

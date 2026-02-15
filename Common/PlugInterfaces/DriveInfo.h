#pragma once

#include "NavigationMenu.h"

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure
enum DriveInfoFlags : uint32_t
{
    DRIVE_INFO_FLAG_NONE             = 0,
    DRIVE_INFO_FLAG_HAS_DISPLAY_NAME = 0x1,
    DRIVE_INFO_FLAG_HAS_VOLUME_LABEL = 0x2,
    DRIVE_INFO_FLAG_HAS_FILE_SYSTEM  = 0x4,
    DRIVE_INFO_FLAG_HAS_TOTAL_BYTES  = 0x8,
    DRIVE_INFO_FLAG_HAS_FREE_BYTES   = 0x10,
    DRIVE_INFO_FLAG_HAS_USED_BYTES   = 0x20,
};

struct DriveInfo
{
    DriveInfoFlags flags;
    // Display name for headers (e.g., "C:\\" or "s3://bucket").
    const wchar_t* displayName;
    // Optional volume label.
    const wchar_t* volumeLabel;
    // Optional file system name.
    const wchar_t* fileSystem;
    // Optional total size in bytes.
    unsigned __int64 totalBytes;
    // Optional free bytes.
    unsigned __int64 freeBytes;
    // Optional used bytes.
    unsigned __int64 usedBytes;
};

enum DriveInfoCommand : uint32_t
{
    DRIVE_INFO_COMMAND_NONE       = 0,
    DRIVE_INFO_COMMAND_PROPERTIES = 1,
    DRIVE_INFO_COMMAND_CLEANUP    = 2,
};
#pragma warning(pop)

// Plugin-provided drive information and optional drive menu commands.
// Notes:
// - Returned pointers are owned by the plugin and remain valid until the next call
//   to the same method or until the object is released.
interface __declspec(uuid("b612a5d1-7e55-4e08-a3da-8d0d9f5d0f31")) __declspec(novtable) IDriveInfo : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept                                            = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDriveMenuItems(const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE ExecuteDriveMenuCommand(unsigned int commandId, const wchar_t* path) noexcept                          = 0;
};

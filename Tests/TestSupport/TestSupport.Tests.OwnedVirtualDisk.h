#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <winioctl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include <array>
#include <cstddef>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#include "Common.Tests.SandboxPath.h"

// Consumer side of the test-owned virtual-disk mount (Specs/Testing/Testing_SelfTests.md).
// Tools/Modules/Testing/TestSandbox.psm1 (New-RSTestOwnedVirtualDiskMount) creates a VHDX inside
// the current run, proves the formatted volume belongs to it, mounts it at an owned run directory
// and hands both to native tests through the environment. Test code never mounts or unmounts a
// volume: this header only decides whether a handed-off mount is that test-owned volume and
// refuses every other volume, including a real user volume reached through a foreign GUID.
namespace RedSalamander::TestSupport
{
inline constexpr std::wstring_view kOwnedVirtualDiskMountEnvironmentVariable{L"REDSALAMANDER_TEST_OWNED_VHDX_MOUNT"};
inline constexpr std::wstring_view kOwnedVirtualDiskVolumeEnvironmentVariable{L"REDSALAMANDER_TEST_OWNED_VHDX_VOLUME"};
inline constexpr std::wstring_view kOwnedVirtualDiskUnavailableEnvironmentVariable{L"REDSALAMANDER_TEST_OWNED_VHDX_UNAVAILABLE"};
inline constexpr std::wstring_view kOwnedVirtualDiskDirectoryName{L"test-owned-vhdx"};

struct OwnedVirtualDiskMountFacts final
{
    bool parentAuthorized      = false;        // marked sandbox, reparse-free parent chain
    bool directoryIsMountPoint = false;        // IO_REPARSE_TAG_MOUNT_POINT on the exact directory
    std::wstring resolvedVolumeGuidPath;       // GetVolumeNameForVolumeMountPointW(mount\)
    std::vector<std::wstring> volumePathNames; // every access path of the expected volume
    bool fileBackedVirtualDisk = false;        // STORAGE_DEVICE_DESCRIPTOR.BusType
    std::wstring fileSystemName;
};

// <testRoot>\runs\<runId>\scratch\test-owned-vhdx\<name>\mount, the only path the harness mounts.
[[nodiscard]] inline std::filesystem::path GetOwnedVirtualDiskMountPath(const std::filesystem::path& testRoot, std::wstring_view runId, std::wstring_view name)
{
    return (testRoot / L"runs" / std::wstring(runId) / L"scratch" / std::wstring(kOwnedVirtualDiskDirectoryName) / std::wstring(name) / L"mount")
        .lexically_normal();
}

[[nodiscard]] inline bool IsVolumeGuidPathSyntax(std::wstring_view text) noexcept
{
    constexpr std::wstring_view kPrefix{L"\\\\?\\Volume{"};
    constexpr size_t kGuidLength = 36u;
    if (text.size() != kPrefix.size() + kGuidLength + 2u || text.substr(0u, kPrefix.size()) != kPrefix || text[text.size() - 2u] != L'}' ||
        text.back() != L'\\')
    {
        return false;
    }
    for (size_t index = 0u; index < kGuidLength; ++index)
    {
        const wchar_t ch     = text[kPrefix.size() + index];
        const bool separator = index == 8u || index == 13u || index == 18u || index == 23u;
        if (separator ? ch != L'-' : std::iswxdigit(ch) == 0)
        {
            return false;
        }
    }
    return true;
}

// Pure decision. Returns an empty string only when the observed facts prove that mountPath is the
// exact run-owned mount of the expected file-backed virtual volume and that the volume is reachable
// through no other path (a drive letter or second mount means a volume this run did not create).
[[nodiscard]] inline std::wstring EvaluateOwnedVirtualDiskMount(const OwnedVirtualDiskMountFacts& facts,
                                                                std::wstring_view expectedVolumeGuidPath,
                                                                const std::filesystem::path& mountPath,
                                                                const std::filesystem::path& expectedMountPath)
{
    const auto samePath = [](const std::filesystem::path& left, const std::filesystem::path& right) noexcept
    { return Common::Testing::IsSameOrDescendantTestSandboxPath(left, right) && Common::Testing::IsSameOrDescendantTestSandboxPath(right, left); };
    if (expectedMountPath.empty() || ! samePath(mountPath, expectedMountPath))
        return std::format(L"mount-path-not-run-owned: {} expected {}", mountPath.native(), expectedMountPath.native());
    if (! IsVolumeGuidPathSyntax(expectedVolumeGuidPath))
        return std::format(L"volume-guid-malformed: {}", expectedVolumeGuidPath);
    if (! facts.parentAuthorized)
        return L"mount-parent-not-in-marked-reparse-free-sandbox";
    if (! facts.directoryIsMountPoint)
        return L"mount-directory-is-not-a-volume-mount-point";
    if (_wcsicmp(facts.resolvedVolumeGuidPath.c_str(), std::wstring(expectedVolumeGuidPath).c_str()) != 0)
        return std::format(L"foreign-volume-guid: mount resolves to '{}', not the test-owned '{}'", facts.resolvedVolumeGuidPath, expectedVolumeGuidPath);
    if (! facts.fileBackedVirtualDisk)
        return std::format(L"volume-not-file-backed-virtual-disk: {}", expectedVolumeGuidPath);
    if (facts.volumePathNames.size() != 1u || ! samePath(std::filesystem::path(facts.volumePathNames.front()), mountPath))
    {
        std::wstring names;
        for (const auto& name : facts.volumePathNames)
            names.append(names.empty() ? L"" : L";").append(name);
        return std::format(L"volume-reachable-outside-run-mount: paths='{}'", names);
    }
    if (_wcsicmp(facts.fileSystemName.c_str(), L"NTFS") != 0)
        return std::format(L"volume-not-ntfs: {}", facts.fileSystemName);
    return {};
}

// Only product operation arguments use this alias. Fixture seeding, inspection
// and cleanup retain the marked mount path. No foreign/lettered volume or
// escaping relative path can acquire an operation path through this ownership
// gate.
[[nodiscard]] inline std::optional<std::filesystem::path> TryMakeOwnedVirtualDiskOperationPath(const OwnedVirtualDiskMountFacts& facts,
                                                                                               std::wstring_view expectedVolumeGuidPath,
                                                                                               const std::filesystem::path& mountPath,
                                                                                               const std::filesystem::path& expectedMountPath,
                                                                                               const std::filesystem::path& ownedPath,
                                                                                               std::wstring& refusal)
{
    refusal = EvaluateOwnedVirtualDiskMount(facts, expectedVolumeGuidPath, mountPath, expectedMountPath);
    if (! refusal.empty())
        return std::nullopt;
    const auto relative = ownedPath.lexically_relative(mountPath);
    if (! ownedPath.is_absolute() || relative.empty() || relative.has_root_path())
    {
        refusal = L"operation-path-not-an-owned-child";
        return std::nullopt;
    }
    for (const auto& component : relative)
    {
        if (component.empty() || component == L"." || component == L"..")
        {
            refusal = L"operation-path-not-an-owned-child";
            return std::nullopt;
        }
    }
    return std::filesystem::path(expectedVolumeGuidPath) / relative;
}

// Read-only observation; requires no privilege. Opening a volume with zero access only
// queries its storage descriptor.
[[nodiscard]] inline OwnedVirtualDiskMountFacts ObserveOwnedVirtualDiskMount(const std::filesystem::path& mountPath, std::wstring_view expectedVolumeGuidPath)
{
    OwnedVirtualDiskMountFacts facts{};
    std::error_code ec;
    facts.parentAuthorized = mountPath.has_parent_path() && Common::Testing::IsAuthorizedTestSandboxPath(mountPath.parent_path(), {}, false, ec);
    wil::unique_handle directory(CreateFileW(mountPath.c_str(),
                                             FILE_READ_ATTRIBUTES,
                                             FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                             nullptr,
                                             OPEN_EXISTING,
                                             FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                             nullptr));
    FILE_ATTRIBUTE_TAG_INFO tag{};
    facts.directoryIsMountPoint  = directory && GetFileInformationByHandleEx(directory.get(), FileAttributeTagInfo, &tag, sizeof(tag)) &&
                                   (tag.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u && (tag.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u &&
                                   tag.ReparseTag == IO_REPARSE_TAG_MOUNT_POINT;
    const std::wstring mountText = mountPath.native() + L"\\";
    std::array<wchar_t, MAX_PATH> resolved{};
    if (facts.directoryIsMountPoint && GetVolumeNameForVolumeMountPointW(mountText.c_str(), resolved.data(), static_cast<DWORD>(resolved.size())))
        facts.resolvedVolumeGuidPath.assign(resolved.data());
    if (! IsVolumeGuidPathSyntax(expectedVolumeGuidPath))
        return facts;

    const std::wstring volumeName(expectedVolumeGuidPath);
    std::wstring names(MAX_PATH, L'\0');
    DWORD required = 0u;
    while (! GetVolumePathNamesForVolumeNameW(volumeName.c_str(), names.data(), static_cast<DWORD>(names.size()), &required))
    {
        if (GetLastError() != ERROR_MORE_DATA || required <= names.size() || required > 64u * 1024u)
        {
            names.clear();
            facts.volumePathNames.emplace_back(L"<unavailable>");
            break;
        }
        names.assign(required, L'\0');
    }
    for (const wchar_t* name = names.c_str(); ! names.empty() && *name != L'\0'; name += std::wcslen(name) + 1u)
        facts.volumePathNames.emplace_back(name);

    const std::wstring device = volumeName.substr(0u, volumeName.size() - 1u); // no trailing separator: open the volume, not its root
    wil::unique_handle volume(CreateFileW(device.c_str(), 0u, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0u, nullptr));
    STORAGE_PROPERTY_QUERY query{};
    query.PropertyId = StorageDeviceProperty;
    query.QueryType  = PropertyStandardQuery;
    STORAGE_DEVICE_DESCRIPTOR descriptor{};
    DWORD returned = 0u;
    facts.fileBackedVirtualDisk =
        volume && DeviceIoControl(volume.get(), IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), &descriptor, sizeof(descriptor), &returned, nullptr) &&
        returned >= offsetof(STORAGE_DEVICE_DESCRIPTOR, BusType) + sizeof(descriptor.BusType) && descriptor.BusType == BusTypeFileBackedVirtual;

    std::array<wchar_t, MAX_PATH> fileSystem{};
    if (facts.directoryIsMountPoint &&
        GetVolumeInformationW(mountText.c_str(), nullptr, 0u, nullptr, nullptr, nullptr, fileSystem.data(), static_cast<DWORD>(fileSystem.size())))
        facts.fileSystemName.assign(fileSystem.data());
    return facts;
}
} // namespace RedSalamander::TestSupport

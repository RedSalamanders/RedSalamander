#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <PortableDeviceApi.h>

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

struct FileSystemBasicInformation;

namespace FileSystemMtpInternal
{
struct MtpItem
{
    std::wstring name;
    unsigned long attributes = 0;
    uint64_t sizeBytes       = 0;
    __int64 creationTime     = 0;
    __int64 lastAccessTime   = 0;
    __int64 lastWriteTime    = 0;
    __int64 changeTime       = 0;
    std::wstring persistentId;
    std::wstring objectId;
    bool streamable = false;
};

struct MtpBackendInfo
{
    bool readOnly      = true;
    bool supportsWrite = false;
    bool liveWpd       = false;
};

struct MtpConnectionBrowseDevice
{
    std::wstring pnpId;
    std::wstring friendlyName;
    std::wstring devicePuid;
};

struct MtpConnectionBrowseStorage
{
    std::wstring name;
    std::wstring persistentId;
    std::wstring objectId;
    std::wstring initialPath;
};

class IMtpBackendFileReader
{
public:
    virtual ~IMtpBackendFileReader() = default;

    IMtpBackendFileReader(const IMtpBackendFileReader&)            = delete;
    IMtpBackendFileReader& operator=(const IMtpBackendFileReader&) = delete;

    virtual HRESULT GetSize(uint64_t& sizeBytes) noexcept                                                        = 0;
    virtual HRESULT Seek(__int64 offset, unsigned long origin, uint64_t& newPosition) noexcept                  = 0;
    virtual HRESULT Read(std::span<std::byte> buffer, unsigned long requestedBytes, unsigned long& bytesRead) noexcept = 0;

protected:
    IMtpBackendFileReader() = default;
};

class IMtpBackend
{
public:
    virtual ~IMtpBackend() = default;

    IMtpBackend(const IMtpBackend&)            = delete;
    IMtpBackend& operator=(const IMtpBackend&) = delete;

    virtual MtpBackendInfo GetInfo() const noexcept                                                                           = 0;
    virtual HRESULT EnumerateDirectory(std::wstring_view path, std::vector<MtpItem>& items) noexcept                          = 0;
    virtual HRESULT GetAttributes(std::wstring_view path, unsigned long& attributes) noexcept                                 = 0;
    virtual HRESULT GetBasicInformation(std::wstring_view path, FileSystemBasicInformation& info) noexcept                    = 0;
    virtual HRESULT GetFileSize(std::wstring_view path, uint64_t& sizeBytes) noexcept                                         = 0;
    virtual HRESULT CreateFileReader(std::wstring_view path, std::shared_ptr<IMtpBackendFileReader>& reader) noexcept        = 0;
    virtual HRESULT ReadFile(std::wstring_view path, std::vector<std::byte>& bytes) noexcept                                  = 0;
    virtual HRESULT WriteFile(std::wstring_view path, std::span<const std::byte> bytes, bool allowOverwrite) noexcept         = 0;
    virtual HRESULT CreateDirectory(std::wstring_view path) noexcept                                                          = 0;
    virtual HRESULT DeleteItem(std::wstring_view path, bool recursive) noexcept                                               = 0;
    virtual HRESULT RenameItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept = 0;
    virtual HRESULT CopyItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept   = 0;
    virtual HRESULT MoveItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept   = 0;
    virtual HRESULT GetItemProperties(std::wstring_view path, std::string& jsonUtf8) noexcept                                 = 0;
    virtual void RequestCancel() noexcept
    {
    }

protected:
    IMtpBackend() = default;
};

class MtpBackendCommandQueue;

using MemoryBackendReadObserver = void (*)(void* context, uint64_t bytesRead) noexcept;

[[nodiscard]] std::shared_ptr<IMtpBackendFileReader> CreateMemoryBackendFileReader(std::vector<std::byte> bytes,
                                                                                   uint32_t readDelayMs = 0,
                                                                                   std::shared_ptr<void> readContext = {},
                                                                                   MemoryBackendReadObserver observer = nullptr);

enum class MtpBackendCommandKind : uint8_t
{
    ReadOnly,
    Mutating,
};

[[nodiscard]] std::unique_ptr<IMtpBackend> CreateWpdMtpBackend() noexcept;
[[nodiscard]] std::unique_ptr<IMtpBackend> CreateFakeMtpBackend(std::string_view optionsJsonUtf8) noexcept;
#ifdef _DEBUG
[[nodiscard]] HRESULT CreateSelfTestWpdMtpBackend(std::string_view optionsJsonUtf8, std::unique_ptr<IMtpBackend>& backend) noexcept;
[[nodiscard]] bool RunOverwriteJournalGenerationSelfTest() noexcept;
[[nodiscard]] uint64_t ResetOverwriteJournalProbeCountForSelfTest() noexcept;
[[nodiscard]] uint64_t GetOverwriteJournalProbeCountForSelfTest() noexcept;
void NotifyOverwriteJournalInjectedForSelfTest(std::wstring_view deviceIdentity) noexcept;
#endif

[[nodiscard]] HRESULT EnumerateMtpConnectionBrowseDevices(std::vector<MtpConnectionBrowseDevice>& devices) noexcept;
[[nodiscard]] HRESULT EnumerateMtpConnectionBrowseStorages(std::wstring_view parentDeviceId, std::vector<MtpConnectionBrowseStorage>& storages) noexcept;
[[nodiscard]] HRESULT EnumerateMtpConnectionBrowseDevicesFromBackend(IMtpBackend& backend, std::vector<MtpConnectionBrowseDevice>& devices) noexcept;
[[nodiscard]] HRESULT EnumerateMtpConnectionBrowseStoragesFromBackend(IMtpBackend& backend,
                                                                      std::wstring_view parentDeviceId,
                                                                      std::vector<MtpConnectionBrowseStorage>& storages) noexcept;

[[nodiscard]] std::wstring NormalizeMtpPath(std::wstring_view rawPath) noexcept;
[[nodiscard]] std::vector<std::wstring_view> SplitPathSegments(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring ParentPath(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring LeafName(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring JoinPath(std::wstring_view parent, std::wstring_view leaf) noexcept;
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept;
[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept;
[[nodiscard]] std::uint64_t StableMtpIdentityHash(std::wstring_view value) noexcept;
[[nodiscard]] std::wstring FormatMtpIdentityHash(std::wstring_view value);
[[nodiscard]] std::wstring SanitizeMtpPathComponent(std::wstring value);
[[nodiscard]] std::wstring MtpDeviceIdentitySuffix(std::wstring_view pnpId);
[[nodiscard]] std::wstring MtpPersistentObjectIdentitySuffix(std::wstring_view persistentId);
[[nodiscard]] std::wstring MtpObjectIdentitySuffix(std::wstring_view objectId);
[[nodiscard]] std::wstring MtpDuplicateObjectSuffix(const MtpItem& item);
[[nodiscard]] std::string JsonEscapeUtf8(std::string_view text);
[[nodiscard]] bool EqualsPathComponent(std::wstring_view a, std::wstring_view b) noexcept;
[[nodiscard]] __int64 SystemTimeToFileTime64(const SYSTEMTIME& value) noexcept;
[[nodiscard]] __int64 NowFileTime64() noexcept;
} // namespace FileSystemMtpInternal

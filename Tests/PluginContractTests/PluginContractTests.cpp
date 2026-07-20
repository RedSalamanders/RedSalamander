#include <algorithm>
#include <array>
#include <filesystem>
#include <format>
#include <iostream>
#include <memory>
#include <span>
#include <string>
#include <string_view>

#include <Windows.h>
#include <unknwn.h>

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "FileSystemPathIdentity.h"
#include "Helpers.h"
#include "PackedFileInfoBuffer.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FactoryImpl.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "TestSupport.h"
#include "YyjsonHelpers.h"

#include <wil/com.h>
#include <wil/resource.h>
#include <yyjson.h>

namespace
{

// ---------------------------------------------------------------------------
// yyjson RAII
// ---------------------------------------------------------------------------
using unique_yyjson_doc = Common::Json::UniqueDocument;

// ---------------------------------------------------------------------------
// Minimal IHost stub for filesystem plugin instantiation
// ---------------------------------------------------------------------------
class NullHost final : public IHost
{
public:
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IHost))
        {
            *ppvObject = static_cast<IHost*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG __stdcall AddRef() noexcept override
    {
        return 2;
    }

    ULONG __stdcall Release() noexcept override
    {
        return 1;
    }
};

NullHost g_nullHost;

std::array<PluginMetaData, 2> g_factoryContractMetaData{};
PluginMetaData g_nonContiguousFactoryMetaData{};

const PluginMetaData* GetFactoryContractMetaData0() noexcept
{
    return &g_factoryContractMetaData[0];
}

const PluginMetaData* GetFactoryContractMetaData1() noexcept
{
    return &g_factoryContractMetaData[1];
}

const PluginMetaData* GetNonContiguousFactoryMetaData() noexcept
{
    return &g_nonContiguousFactoryMetaData;
}

const char* GetEmptyFactoryContractSchema() noexcept
{
    return nullptr;
}

HRESULT CreateFactoryContractInstance(const FactoryOptions*, IHost*, void**) noexcept
{
    return E_NOTIMPL;
}

// ---------------------------------------------------------------------------
// Check helper (matches LocalizationTests pattern)
// ---------------------------------------------------------------------------
void Check(bool condition, const wchar_t* message, bool& success) noexcept
{
    if (! condition)
    {
        std::wcerr << L"[ FAILED  ] " << message << L"\n";
        success = false;
        return;
    }
    std::wcout << L"[       OK ] " << message << L"\n";
}

struct PackedFileInfoTestEntry
{
    std::wstring name;
    unsigned long fileIndex  = 0;
    unsigned long attributes = 0;
    uint64_t sizeBytes       = 0;
};

void TestPackedFileInfoBuffer(bool& success) noexcept
{
    Common::Plugins::PackedFileInfoBuffer buffer;
    Check(buffer.Build(std::vector<PackedFileInfoTestEntry>{}, [](const PackedFileInfoTestEntry&, FileInfo&) noexcept {}) == S_OK,
          L"packed FileInfo owner accepts an empty result",
          success);

    FileInfo* first        = reinterpret_cast<FileInfo*>(static_cast<uintptr_t>(1u));
    unsigned long byteSize = 1;
    unsigned long count    = 1;
    Check(buffer.GetBuffer(&first) == S_OK && first == nullptr, L"packed FileInfo empty result exposes a null buffer", success);
    Check(buffer.GetBufferSize(&byteSize) == S_OK && byteSize == 0, L"packed FileInfo empty result has zero used bytes", success);
    Check(buffer.GetCount(&count) == S_OK && count == 0, L"packed FileInfo empty result has zero entries", success);
    Check(
        buffer.Get(0, &first) == HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES) && first == nullptr, L"packed FileInfo empty result rejects indexed access", success);

    const std::vector<PackedFileInfoTestEntry> entries = {
        {L"a", 7u, FILE_ATTRIBUTE_NORMAL, 11u},
        {L"longer-name.bin", 9u, FILE_ATTRIBUTE_ARCHIVE, 42u},
        {L"folder", 10u, FILE_ATTRIBUTE_DIRECTORY, 0u},
    };
    const HRESULT buildHr = buffer.Build(entries,
                                         [](const PackedFileInfoTestEntry& source, FileInfo& entry) noexcept
    {
        entry.FileIndex      = source.fileIndex;
        entry.FileAttributes = source.attributes;
        entry.EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry.AllocationSize = static_cast<__int64>(source.sizeBytes);
    });
    Check(buildHr == S_OK, L"packed FileInfo owner builds a multi-entry result", success);
    Check(buffer.GetBuffer(&first) == S_OK && first != nullptr, L"packed FileInfo multi-entry result exposes its buffer", success);
    Check(
        first != nullptr && (reinterpret_cast<uintptr_t>(first) % alignof(FileInfo)) == 0u, L"packed FileInfo buffer base honors FileInfo alignment", success);
    Check(buffer.GetCount(&count) == S_OK && count == entries.size(), L"packed FileInfo count matches the source entries", success);

    for (unsigned long index = 0; index < entries.size(); ++index)
    {
        FileInfo* entry                    = nullptr;
        const HRESULT getHr                = buffer.Get(index, &entry);
        const std::wstring_view actualName = entry ? std::wstring_view(entry->FileName, entry->FileNameSize / sizeof(wchar_t)) : std::wstring_view{};
        Check(getHr == S_OK && entry != nullptr && actualName == entries[index].name && entry->FileIndex == entries[index].fileIndex &&
                  entry->FileAttributes == entries[index].attributes && entry->EndOfFile == static_cast<__int64>(entries[index].sizeBytes),
              std::format(L"packed FileInfo entry {} preserves name and provider metadata", index).c_str(),
              success);
        Check(entry != nullptr && (reinterpret_cast<uintptr_t>(entry) % alignof(FileInfo)) == 0u,
              std::format(L"packed FileInfo entry {} honors FileInfo alignment", index).c_str(),
              success);
    }

    const unsigned long originalOffset = first ? first->NextEntryOffset : 0u;
    if (first)
    {
        first->NextEntryOffset = 1u;
    }
    FileInfo* malformed = nullptr;
    Check(buffer.Get(1, &malformed) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && malformed == nullptr,
          L"packed FileInfo traversal rejects a short unaligned next offset",
          success);
    if (first)
    {
        first->NextEntryOffset = originalOffset;
        ++first->FileNameSize;
    }
    Check(buffer.Get(0, &malformed) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && malformed == nullptr,
          L"packed FileInfo traversal rejects an odd UTF-16 name byte count",
          success);

    Check(buffer.GetBuffer(nullptr) == E_POINTER && buffer.GetBufferSize(nullptr) == E_POINTER && buffer.GetAllocatedSize(nullptr) == E_POINTER &&
              buffer.GetCount(nullptr) == E_POINTER && buffer.Get(0, nullptr) == E_POINTER,
          L"packed FileInfo query methods reject null out-parameters",
          success);
}

// ---------------------------------------------------------------------------
// Export function typedefs
// ---------------------------------------------------------------------------
using PfnEnumeratePlugins       = HRESULT(__stdcall*)(REFIID, const PluginMetaData**, unsigned int*);
using PfnCreate                 = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
using PfnGetConfigurationSchema = HRESULT(__stdcall*)(REFIID, const wchar_t*, const char**);
using PfnRunDebugSelfTests      = HRESULT(__stdcall*)(unsigned int*, unsigned int*);
using PfnPluginShutdown         = void(__stdcall*)() noexcept;
using PfnPluginCanUnloadNow     = BOOL(__stdcall*)() noexcept;
using PfnDebugCurlRuntimeProbe  = HRESULT(__stdcall*)() noexcept;

#if defined(_DEBUG)
constexpr bool kDebugSelfTestExportsRequired = true;
#else
constexpr bool kDebugSelfTestExportsRequired = false;
#endif

// ---------------------------------------------------------------------------
// Plugin sets (paths relative to the exe output dir; exe runs from .build\x64\Debug\)
// DLLs live in the Plugins\ subfolder per REVIEWER CORRECTION 1.
// ---------------------------------------------------------------------------
constexpr std::array<std::wstring_view, 7> kFilesystemDlls = {
    L"Plugins\\FileSystem.dll",
    L"Plugins\\FileSystem7z.dll",
    L"Plugins\\FileSystemCurl.dll",
    L"Plugins\\FileSystemDummy.dll",
    L"Plugins\\FileSystemGoogleDrive.dll",
    L"Plugins\\FileSystemMicrosoftDrive.dll",
    L"Plugins\\FileSystemS3.dll",
};

constexpr std::array<std::wstring_view, 7> kViewerDlls = {
    L"Plugins\\ViewerText.dll",
    L"Plugins\\ViewerSqlite.dll",
    L"Plugins\\ViewerSpace.dll",
    L"Plugins\\ViewerImgRaw.dll",
    L"Plugins\\ViewerVLC.dll",
    L"Plugins\\ViewerPE.dll",
    L"Plugins\\ViewerWeb.dll",
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
[[nodiscard]] std::wstring GetExeDir() noexcept
{
    wchar_t path[MAX_PATH]{};
    const DWORD len = GetModuleFileNameW(nullptr, path, MAX_PATH);
    if (len == 0)
    {
        return {};
    }
    std::wstring result(path, len);
    const auto pos = result.rfind(L'\\');
    if (pos == std::wstring::npos)
    {
        return {};
    }
    result.resize(pos + 1); // keep trailing backslash
    return result;
}

// Parse a const char* as JSON and return the doc (null on failure).
[[nodiscard]] unique_yyjson_doc ParseJson(const char* utf8) noexcept
{
    if (utf8 == nullptr || utf8[0] == '\0')
    {
        return {};
    }
    yyjson_read_err err{};
    // yyjson_read takes a const char* in read-only mode (no INSITU flag).
    yyjson_doc* doc = yyjson_read_opts(const_cast<char*>(utf8), std::strlen(utf8), YYJSON_READ_NOFLAG, nullptr, &err);
    return unique_yyjson_doc(doc);
}

// ---------------------------------------------------------------------------
// Step 2: Enumerate-and-validate (metadata + schema, no instance creation)
// ---------------------------------------------------------------------------
bool TestEnumerateAndSchema(std::wstring_view relPath, const IID& expectedIid, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    const std::wstring loadMsg = std::format(L"{}: DLL loads", relPath);
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(mod), loadMsg.c_str(), success);
    if (! mod)
    {
        return false;
    }

    // Resolve exports
    const auto pfnEnumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate    = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    const auto pfnSchema    = reinterpret_cast<PfnGetConfigurationSchema>(GetProcAddress(mod.get(), "RedSalamanderGetConfigurationSchema"));

    Check(pfnEnumerate != nullptr, std::format(L"{}: RedSalamanderEnumeratePlugins export resolves", relPath).c_str(), success);
    Check(pfnCreate != nullptr, std::format(L"{}: RedSalamanderCreate export resolves", relPath).c_str(), success);
    Check(pfnSchema != nullptr, std::format(L"{}: RedSalamanderGetConfigurationSchema export resolves", relPath).c_str(), success);

    if (pfnEnumerate == nullptr || pfnCreate == nullptr || pfnSchema == nullptr)
    {
        return false;
    }

    // Happy-path enumerate with correct IID
    const PluginMetaData* metaData = nullptr;
    unsigned int count             = 0;
    const HRESULT hrEnum           = pfnEnumerate(expectedIid, &metaData, &count);
    Check(hrEnum == S_OK, std::format(L"{}: EnumeratePlugins(correct IID) returns S_OK", relPath).c_str(), success);
    Check(count >= 1, std::format(L"{}: EnumeratePlugins returns count >= 1", relPath).c_str(), success);
    Check(IsValidEnumeratedPluginCount(count), std::format(L"{}: EnumeratePlugins count stays within the host contract", relPath).c_str(), success);
    Check(metaData != nullptr, std::format(L"{}: EnumeratePlugins returns non-null metaData", relPath).c_str(), success);

    if (hrEnum != S_OK || metaData == nullptr || ! IsValidEnumeratedPluginCount(count))
    {
        return false;
    }

    // Validate each metadata entry
    for (unsigned int i = 0; i < count; ++i)
    {
        const PluginMetaData& md = metaData[i];
        Check(md.id != nullptr && md.id[0] != L'\0', std::format(L"{}: metaData[{}].id is non-empty", relPath, i).c_str(), success);
        Check(md.name != nullptr && md.name[0] != L'\0', std::format(L"{}: metaData[{}].name is non-empty", relPath, i).c_str(), success);
        // version is optional per spec (may be nullptr) — just note its presence
        (void)md.version;
    }

    // Wrong IID must return E_NOINTERFACE and count == 0
    const PluginMetaData* wrongMeta = nullptr;
    unsigned int wrongCount         = 0;
    // Bogus IID that matches neither IFileSystem nor IViewer
    static constexpr BYTE kBogusData4[8] = {0xde, 0xad, 0xde, 0xad, 0xde, 0xad, 0xde, 0xad};
    const IID bogusIid    = {0xdeadbeef,
                             0xdead,
                             0xdead,
                             {kBogusData4[0], kBogusData4[1], kBogusData4[2], kBogusData4[3], kBogusData4[4], kBogusData4[5], kBogusData4[6], kBogusData4[7]}};
    const HRESULT hrWrong = pfnEnumerate(bogusIid, &wrongMeta, &wrongCount);
    Check(hrWrong == E_NOINTERFACE, std::format(L"{}: EnumeratePlugins(wrong IID) returns E_NOINTERFACE", relPath).c_str(), success);
    Check(wrongCount == 0, std::format(L"{}: EnumeratePlugins(wrong IID) returns count == 0", relPath).c_str(), success);

    // Schema validation for each plugin id
    for (unsigned int i = 0; i < count; ++i)
    {
        const wchar_t* pluginId   = metaData[i].id;
        const char* schemaJson    = nullptr;
        const HRESULT hrSchema    = pfnSchema(expectedIid, pluginId, &schemaJson);
        const bool schemaOk       = hrSchema == S_OK;
        const bool schemaNotFound = hrSchema == static_cast<HRESULT>(HRESULT_FROM_WIN32(ERROR_NOT_FOUND));
        Check(schemaOk || schemaNotFound, std::format(L"{}: GetConfigurationSchema[{}] returns S_OK or ERROR_NOT_FOUND", relPath, i).c_str(), success);
        if (schemaOk && schemaJson != nullptr && schemaJson[0] != '\0')
        {
            unique_yyjson_doc doc = ParseJson(schemaJson);
            Check(static_cast<bool>(doc), std::format(L"{}: GetConfigurationSchema[{}] JSON parses successfully", relPath, i).c_str(), success);
        }
    }

    return true;
}

// ---------------------------------------------------------------------------
// Step 3: GetCapabilities pass (filesystem plugins only)
// ---------------------------------------------------------------------------
[[nodiscard]] bool RequiresTransactionalConfigurationProof(std::wstring_view relativePath) noexcept
{
    constexpr std::array<std::wstring_view, 6> providers = {
        L"Plugins\\FileSystem.dll",
        L"Plugins\\FileSystem7z.dll",
        L"Plugins\\FileSystemCurl.dll",
        L"Plugins\\FileSystemGoogleDrive.dll",
        L"Plugins\\FileSystemMicrosoftDrive.dll",
        L"Plugins\\FileSystemS3.dll",
    };
    return std::ranges::find(providers, relativePath) != providers.end();
}

void TestTransactionalConfiguration(IInformations& information, std::wstring_view relativePath, const wchar_t* pluginId, bool& success) noexcept
{
    constexpr char kForwardConfiguration[] = R"json({"observatoryUnknown":{"value":17}})json";
    const HRESULT forwardHr                = information.SetConfiguration(kForwardConfiguration);
    const char* configuration              = nullptr;
    const HRESULT getForwardHr             = information.GetConfiguration(&configuration);
    const std::string preserved            = configuration != nullptr ? configuration : "";
    Check(forwardHr == S_OK && getForwardHr == S_OK && preserved.find("\"observatoryUnknown\"") != std::string::npos,
          std::format(L"{}: SetConfiguration(pluginId={}) preserves unknown members", relativePath, pluginId).c_str(),
          success);

    const HRESULT malformedHr         = information.SetConfiguration("{");
    configuration                     = nullptr;
    const HRESULT getAfterMalformedHr = information.GetConfiguration(&configuration);
    Check(malformedHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && getAfterMalformedHr == S_OK && configuration != nullptr &&
              std::string_view(configuration) == preserved,
          std::format(L"{}: malformed configuration preserves live state for pluginId={}", relativePath, pluginId).c_str(),
          success);

    const HRESULT wrongRootHr         = information.SetConfiguration("[]");
    configuration                     = nullptr;
    const HRESULT getAfterWrongRootHr = information.GetConfiguration(&configuration);
    Check(wrongRootHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && getAfterWrongRootHr == S_OK && configuration != nullptr &&
              std::string_view(configuration) == preserved,
          std::format(L"{}: wrong-root configuration preserves live state for pluginId={}", relativePath, pluginId).c_str(),
          success);

    if (relativePath == L"Plugins\\FileSystem7z.dll" || relativePath == L"Plugins\\FileSystemCurl.dll")
    {
        constexpr char kLegacyPasswordConfiguration[] = R"json({"defaultPassword":"observatory-password-sentinel","observatoryUnknown":17})json";
        constexpr char kLegacyCurlSecretConfiguration[] =
            R"json({"defaultPassword":"observatory-password-sentinel","sshKeyPassphrase":"observatory-passphrase-sentinel","observatoryUnknown":17})json";
        const char* legacyConfiguration  = relativePath == L"Plugins\\FileSystemCurl.dll" ? kLegacyCurlSecretConfiguration : kLegacyPasswordConfiguration;
        const HRESULT secretHr           = information.SetConfiguration(legacyConfiguration);
        configuration                    = nullptr;
        const HRESULT getSecretHr        = information.GetConfiguration(&configuration);
        const std::string_view sanitized = configuration != nullptr ? configuration : "";
        Check(secretHr == S_OK && getSecretHr == S_OK && sanitized.find("observatoryUnknown") != std::string_view::npos &&
                  sanitized.find("observatory-password-sentinel") == std::string_view::npos &&
                  (relativePath != L"Plugins\\FileSystemCurl.dll" || sanitized.find("observatory-passphrase-sentinel") == std::string_view::npos),
              std::format(L"{}: legacy configuration secrets are imported but not persisted for pluginId={}", relativePath, pluginId).c_str(),
              success);
    }

    Check(information.SetConfiguration("{}") == S_OK,
          std::format(L"{}: transactional configuration proof restores defaults for pluginId={}", relativePath, pluginId).c_str(),
          success);
}

[[nodiscard]] bool CreateEmptyProviderFile(IFileSystemIO& io, const std::wstring& path) noexcept
{
    wil::com_ptr_nothrow<IFileWriter> writer;
    return io.CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put()) == S_OK && writer && writer->Commit() == S_OK;
}

[[nodiscard]] bool ProviderFileExists(IFileSystemIO& io, const std::wstring& path) noexcept
{
    wil::com_ptr_nothrow<IFileReader> reader;
    return io.CreateFileReader(path.c_str(), reader.put()) == S_OK && reader != nullptr;
}

[[nodiscard]] bool SetProviderReadOnly(IFileSystemIO& io, const std::wstring& path) noexcept
{
    FileSystemBasicInformation information{};
    information.sizeBytes = sizeof(information);
    if (FAILED(io.GetFileBasicInformation(path.c_str(), &information)))
    {
        return false;
    }
    information.attributes |= FILE_ATTRIBUTE_READONLY;
    return io.SetFileBasicInformation(path.c_str(), &information) == S_OK;
}

void TestLocalWriterFlagContract(IFileSystem& fileSystem, bool& success) noexcept
{
    wil::com_ptr_nothrow<IFileSystemIO> io;
    if (FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemIO), io.put_void())) || ! io)
    {
        Check(false, L"local provider exposes IFileSystemIO for writer flag proof", success);
        return;
    }

    std::error_code error;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory({.harnessSegment      = L"PluginContractTests",
                                                                                         .leafSegment         = L"observatory-track13-local-writer",
                                                                                         .fallbackRunIdPrefix = L"plugin-contract",
                                                                                         .kind = RedSalamander::TestSupport::TestDirectoryKind::Scratch},
                                                                                        error);
    Check(! error && ! root.empty(), L"local writer flag proof acquires TestSandbox scratch", success);
    if (error || root.empty())
    {
        return;
    }

    const std::filesystem::path target = root / L"readonly-existing.txt";
    wil::unique_handle seed(::CreateFileW(target.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
    const bool seeded = seed != nullptr;
    seed.reset();
    const bool madeReadOnly = seeded && ::SetFileAttributesW(target.c_str(), FILE_ATTRIBUTE_READONLY) != 0;
    Check(madeReadOnly, L"local writer flag proof seeds a read-only destination", success);

    wil::com_ptr_nothrow<IFileWriter> writer;
    const HRESULT createHr = madeReadOnly ? io->CreateFileWriter(target.c_str(), FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY, writer.put()) : E_UNEXPECTED;
    const DWORD attributes = ::GetFileAttributesW(target.c_str());
    Check(createHr == E_INVALIDARG && ! writer && attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_READONLY) != 0u,
          L"local writer rejects replace-readonly without overwrite and preserves destination attributes",
          success);

    if (attributes != INVALID_FILE_ATTRIBUTES)
    {
        static_cast<void>(::SetFileAttributesW(target.c_str(), attributes & ~FILE_ATTRIBUTE_READONLY));
    }
    std::filesystem::remove_all(root, error);
}

void TestDummyTransactionalMutationContracts(IFileSystem& fileSystem, bool& success) noexcept
{
    wil::com_ptr_nothrow<IFileSystemIO> io;
    wil::com_ptr_nothrow<IFileSystemDirectoryOperations> directoryOperations;
    if (FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemIO), io.put_void())) || ! io ||
        FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemDirectoryOperations), directoryOperations.put_void())) || ! directoryOperations)
    {
        Check(false, L"dummy provider exposes mutation interfaces for Track 13 proof", success);
        return;
    }

    const std::wstring root  = std::format(L"/observatory-track13-{}", GetTickCount64());
    const auto makeDirectory = [&](std::wstring_view suffix) noexcept
    { return directoryOperations->CreateDirectory((root + std::wstring(suffix)).c_str()) == S_OK; };
    const auto makeFile = [&](std::wstring_view suffix) noexcept { return CreateEmptyProviderFile(*io.get(), root + std::wstring(suffix)); };

    const bool copySeeded = makeDirectory(L"") && makeDirectory(L"/copy-source") && makeDirectory(L"/copy-destination") &&
                            makeFile(L"/copy-source/first.txt") && makeFile(L"/copy-source/late.txt") && makeFile(L"/copy-destination/late.txt");
    Check(copySeeded, L"dummy copy rollback proof seeds source and late destination collision", success);
    const HRESULT copyHr =
        copySeeded
            ? fileSystem.CopyItem((root + L"/copy-source").c_str(), (root + L"/copy-destination").c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)
            : E_UNEXPECTED;
    Check(copyHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && ! ProviderFileExists(*io.get(), root + L"/copy-destination/first.txt"),
          L"dummy directory copy preflights a late collision without partial destination mutation",
          success);

    const bool moveSeeded = makeDirectory(L"/move-source") && makeDirectory(L"/move-destination") && makeFile(L"/move-source/first.txt") &&
                            makeFile(L"/move-source/late.txt") && makeFile(L"/move-destination/late.txt");
    Check(moveSeeded, L"dummy move rollback proof seeds source and late destination collision", success);
    const HRESULT moveHr =
        moveSeeded
            ? fileSystem.MoveItem((root + L"/move-source").c_str(), (root + L"/move-destination").c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)
            : E_UNEXPECTED;
    Check(moveHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && ProviderFileExists(*io.get(), root + L"/move-source/first.txt") &&
              ! ProviderFileExists(*io.get(), root + L"/move-destination/first.txt"),
          L"dummy directory move preflights a late collision without partial source or destination mutation",
          success);

    const bool deleteSeeded = makeDirectory(L"/delete-source") && makeFile(L"/delete-source/readonly-child.txt") &&
                              SetProviderReadOnly(*io.get(), root + L"/delete-source/readonly-child.txt");
    Check(deleteSeeded, L"dummy recursive delete proof seeds a read-only descendant", success);
    const HRESULT deleteHr =
        deleteSeeded ? fileSystem.DeleteItem((root + L"/delete-source").c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr) : E_UNEXPECTED;
    Check(deleteHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ProviderFileExists(*io.get(), root + L"/delete-source/readonly-child.txt"),
          L"dummy recursive delete applies read-only policy to descendants before mutation",
          success);

    const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    Check(fileSystem.DeleteItem(root.c_str(), cleanupFlags, nullptr, nullptr, nullptr) == S_OK, L"dummy Track 13 proof cleans its provider fixture", success);
}

void TestTrack13ProviderContracts(IFileSystem& fileSystem, std::wstring_view relativePath, bool& success) noexcept
{
    if (relativePath == L"Plugins\\FileSystem.dll")
    {
        TestLocalWriterFlagContract(fileSystem, success);
    }
    else if (relativePath == L"Plugins\\FileSystemDummy.dll")
    {
        TestDummyTransactionalMutationContracts(fileSystem, success);
    }
}

bool TestCapabilities(std::wstring_view relPath, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    // Module must stay alive while we hold interface pointers from it.
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    if (! mod)
    {
        // Already reported in enumerate pass; don't double-report.
        return false;
    }

    const auto pfnEnumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate    = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));

    if (pfnEnumerate == nullptr || pfnCreate == nullptr)
    {
        return false;
    }

    const PluginMetaData* metaData = nullptr;
    unsigned int count             = 0;
    if (pfnEnumerate(__uuidof(IFileSystem), &metaData, &count) != S_OK || count == 0 || metaData == nullptr)
    {
        return false;
    }

    for (unsigned int i = 0; i < count; ++i)
    {
        const wchar_t* pluginId = metaData[i].id;

        void* raw              = nullptr;
        const HRESULT hrCreate = pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), pluginId, &raw);
        Check(hrCreate == S_OK, std::format(L"{}: Create(pluginId={}) returns S_OK", relPath, pluginId).c_str(), success);
        Check(raw != nullptr, std::format(L"{}: Create(pluginId={}) returns non-null instance", relPath, pluginId).c_str(), success);

        if (hrCreate != S_OK || raw == nullptr)
        {
            continue;
        }

        // Wrap in a smart pointer — IFileSystem inherits IUnknown
        wil::com_ptr_nothrow<IFileSystem> fs;
        fs.attach(static_cast<IFileSystem*>(raw));

        TestTrack13ProviderContracts(*fs.get(), relPath, success);

        if (RequiresTransactionalConfigurationProof(relPath))
        {
            wil::com_ptr_nothrow<IInformations> information;
            const HRESULT informationHr = fs->QueryInterface(__uuidof(IInformations), information.put_void());
            Check(informationHr == S_OK && information != nullptr,
                  std::format(L"{}: transactional configuration provider exposes IInformations for pluginId={}", relPath, pluginId).c_str(),
                  success);
            if (information)
            {
                TestTransactionalConfiguration(*information.get(), relPath, pluginId, success);
            }
        }

        const char* capJson = nullptr;
        const HRESULT hrCap = fs->GetCapabilities(&capJson);
        Check(hrCap == S_OK, std::format(L"{}: GetCapabilities(pluginId={}) returns S_OK", relPath, pluginId).c_str(), success);
        Check(capJson != nullptr && capJson[0] != '\0',
              std::format(L"{}: GetCapabilities(pluginId={}) returns non-empty JSON", relPath, pluginId).c_str(),
              success);

        if (hrCap != S_OK || capJson == nullptr || capJson[0] == '\0')
        {
            continue;
        }

        unique_yyjson_doc doc = ParseJson(capJson);
        Check(static_cast<bool>(doc), std::format(L"{}: GetCapabilities(pluginId={}) JSON parses successfully", relPath, pluginId).c_str(), success);

        if (! doc)
        {
            continue;
        }

        yyjson_val* root = yyjson_doc_get_root(doc.get());
        Check(root != nullptr && yyjson_is_obj(root), std::format(L"{}: GetCapabilities(pluginId={}) root is object", relPath, pluginId).c_str(), success);

        if (root == nullptr || ! yyjson_is_obj(root))
        {
            continue;
        }

        // Mandatory fields: "version": 1, "operations" (obj), "concurrency" (obj), "crossFileSystem" (obj)
        yyjson_val* verVal = yyjson_obj_get(root, "version");
        Check(verVal != nullptr && yyjson_is_int(verVal) && yyjson_get_int(verVal) == 1,
              std::format(L"{}: GetCapabilities(pluginId={}) has \"version\": 1", relPath, pluginId).c_str(),
              success);

        yyjson_val* opsVal = yyjson_obj_get(root, "operations");
        Check(opsVal != nullptr && yyjson_is_obj(opsVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"operations\" object", relPath, pluginId).c_str(),
              success);

        yyjson_val* concVal = yyjson_obj_get(root, "concurrency");
        Check(concVal != nullptr && yyjson_is_obj(concVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"concurrency\" object", relPath, pluginId).c_str(),
              success);

        yyjson_val* crossVal = yyjson_obj_get(root, "crossFileSystem");
        Check(crossVal != nullptr && yyjson_is_obj(crossVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"crossFileSystem\" object", relPath, pluginId).c_str(),
              success);

        // pathIdentity is now a MANDATORY block: the host (ParseFileSystemCapabilitiesJson ->
        // TryParseFileSystemPathIdentityContract, FolderWindow.FileOperations.cpp:139) fails the WHOLE
        // capability set closed -- disabling copy/move/delete and cross-filesystem operations -- if it
        // is absent or fails the strict parser. Assert each plugin's block is present and parses under
        // the exact host contract so a typo (preferredSeparator not in acceptedSeparators, a misspelled
        // enum, a missing bool, normalization != "none", ...) can never silently brick the plugin in
        // production while this contract test stays green.
        yyjson_val* pathIdVal = yyjson_obj_get(root, "pathIdentity");
        Check(pathIdVal != nullptr && yyjson_is_obj(pathIdVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"pathIdentity\" object", relPath, pluginId).c_str(),
              success);

        const std::optional<FileSystemPathIdentity> parsedPathIdentity = TryParseFileSystemPathIdentityContract(std::string_view(capJson), {});
        Check(parsedPathIdentity.has_value(),
              std::format(L"{}: GetCapabilities(pluginId={}) pathIdentity parses under the host contract", relPath, pluginId).c_str(),
              success);

        // capJson pointer is plugin-owned; do NOT free it.
        // fs (IFileSystem) is released here by wil::com_ptr_nothrow going out of scope.
        // doc (yyjson_doc) is released here.
        // mod stays alive (declared first in the enclosing scope).
    }

    return true;
}

// ---------------------------------------------------------------------------
// Step 4: Negative-contract test (bogus plugin id on FileSystemDummy only)
// ---------------------------------------------------------------------------
bool TestBogusPluginId(bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + L"Plugins\\FileSystemDummy.dll";

    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(mod), L"FileSystemDummy.dll: loads for negative-id test", success);
    if (! mod)
    {
        return false;
    }

    const auto pfnCreate = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    Check(pfnCreate != nullptr, L"FileSystemDummy.dll: RedSalamanderCreate resolves for negative-id test", success);
    if (pfnCreate == nullptr)
    {
        return false;
    }

    void* raw        = nullptr;
    const HRESULT hr = pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), L"no/such-plugin", &raw);
    Check(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND), L"FileSystemDummy.dll: Create(bogus-id) returns HRESULT_FROM_WIN32(ERROR_NOT_FOUND)", success);
    Check(raw == nullptr, L"FileSystemDummy.dll: Create(bogus-id) returns null result pointer", success);

    return true;
}

[[nodiscard]] constexpr bool IsDebugSelfTestExportPresenceAccepted(bool exportPresent, bool exportRequired) noexcept
{
    return exportPresent || ! exportRequired;
}

[[nodiscard]] bool ValidateDebugSelfTestExportPresence(bool exportPresent, bool exportRequired, std::wstring_view label, bool& success) noexcept
{
    if (exportPresent)
    {
        return true;
    }

    if (exportRequired)
    {
        Check(false, std::format(L"{}: required debug selftest export resolves", label).c_str(), success);
        return false;
    }

    std::wcout << L"[  SKIPPED ] " << label << L": debug selftest export is not expected in this configuration\n";
    return false;
}

bool TestPluginDebugSelfTests(std::wstring_view relativeDllPath, const char* exportName, std::wstring_view label, bool exportRequired, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relativeDllPath);

    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    const std::wstring loadMessage = std::format(L"{}: loads for debug selftests", label);
    Check(static_cast<bool>(mod), loadMessage.c_str(), success);
    if (! mod)
    {
        return false;
    }

    const auto pfnSelfTests = reinterpret_cast<PfnRunDebugSelfTests>(GetProcAddress(mod.get(), exportName));
    if (! ValidateDebugSelfTestExportPresence(pfnSelfTests != nullptr, exportRequired, label, success))
    {
        return ! exportRequired;
    }

    unsigned int passed = 0;
    unsigned int failed = 0;
    const HRESULT hr    = pfnSelfTests(&passed, &failed);
    const std::wstring passMessage =
        std::format(L"{}: debug selftests pass (passed={}, failed={}, hr=0x{:08X})", label, passed, failed, static_cast<unsigned long>(hr));
    Check(SUCCEEDED(hr) && failed == 0 && passed > 0, passMessage.c_str(), success);

    const auto shutdown  = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(mod.get(), "RedSalamanderPluginShutdown"));
    const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(mod.get(), "RedSalamanderPluginCanUnloadNow"));
    if (shutdown && canUnload)
    {
        shutdown();
        Check(canUnload() == TRUE, std::format(L"{}: debug selftest module reaches its unload quiet point", label).c_str(), success);
    }
    return true;
}

// ---------------------------------------------------------------------------
// Run all tests for a DLL set
// ---------------------------------------------------------------------------
bool RunEnumerateAndSchemaPass(std::span<const std::wstring_view> dlls, const IID& iid) noexcept
{
    bool success = true;
    for (const std::wstring_view relPath : dlls)
    {
        TestEnumerateAndSchema(relPath, iid, success);
    }
    return success;
}

bool RunCapabilitiesPass(std::span<const std::wstring_view> dlls) noexcept
{
    bool success = true;
    for (const std::wstring_view relPath : dlls)
    {
        TestCapabilities(relPath, success);
    }
    return success;
}

void TestS3RuntimeUnloadContract(bool& success) noexcept
{
    const std::wstring absPath            = GetExeDir() + L"Plugins\\FileSystemS3.dll";
    constexpr unsigned int kRefreshCycles = 8u;
    for (unsigned int cycle = 0u; cycle < kRefreshCycles; ++cycle)
    {
        wil::unique_hmodule module(LoadLibraryExW(absPath.c_str(), nullptr, 0));
        Check(static_cast<bool>(module), std::format(L"FileSystemS3.dll: refresh cycle {} loads the module", cycle + 1u).c_str(), success);
        if (! module)
        {
            return;
        }

        const auto create    = reinterpret_cast<PfnCreate>(GetProcAddress(module.get(), "RedSalamanderCreate"));
        const auto shutdown  = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(module.get(), "RedSalamanderPluginShutdown"));
        const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
        Check(create != nullptr && shutdown != nullptr && canUnload != nullptr,
              std::format(L"FileSystemS3.dll: refresh cycle {} resolves lifecycle exports", cycle + 1u).c_str(),
              success);
        if (create == nullptr || shutdown == nullptr || canUnload == nullptr)
        {
            return;
        }

        void* raw              = nullptr;
        const HRESULT createHr = create(__uuidof(IFileSystem), nullptr, &g_nullHost, L"builtin/file-system-s3", &raw);
        wil::com_ptr_nothrow<IFileSystem> fileSystem;
        fileSystem.attach(static_cast<IFileSystem*>(raw));
        Check(createHr == S_OK && fileSystem, std::format(L"FileSystemS3.dll: refresh cycle {} creates an initialized owner", cycle + 1u).c_str(), success);
        if (FAILED(createHr) || ! fileSystem)
        {
            return;
        }

        Check(canUnload() == FALSE, std::format(L"FileSystemS3.dll: refresh cycle {} is closed before the quiet point", cycle + 1u).c_str(), success);
        shutdown();
        Check(canUnload() == FALSE, std::format(L"FileSystemS3.dll: refresh cycle {} stays closed while an AWS owner is alive", cycle + 1u).c_str(), success);
        fileSystem.reset();
        Check(canUnload() == TRUE, std::format(L"FileSystemS3.dll: refresh cycle {} opens after AWS shutdown", cycle + 1u).c_str(), success);
        shutdown();
        Check(canUnload() == TRUE, std::format(L"FileSystemS3.dll: refresh cycle {} keeps shutdown idempotent", cycle + 1u).c_str(), success);
        module.reset();
        Check(
            GetModuleHandleW(absPath.c_str()) == nullptr, std::format(L"FileSystemS3.dll: refresh cycle {} releases the module", cycle + 1u).c_str(), success);
    }
}

void TestCrossPluginCurlRuntimeUnloadContract(bool& success) noexcept
{
    struct RuntimeModule
    {
        RuntimeModule()                                = default;
        RuntimeModule(const RuntimeModule&)            = delete;
        RuntimeModule& operator=(const RuntimeModule&) = delete;
        RuntimeModule(RuntimeModule&&)                 = default;
        RuntimeModule& operator=(RuntimeModule&&)      = default;

        std::wstring label;
        std::wstring path;
        wil::unique_hmodule module;
        PfnPluginShutdown shutdown      = nullptr;
        PfnPluginCanUnloadNow canUnload = nullptr;
        PfnDebugCurlRuntimeProbe probe  = nullptr;
    };

    const auto loadModule = [&](std::wstring_view fileName, std::wstring_view label) -> RuntimeModule
    {
        RuntimeModule loaded;
        loaded.label = label;
        loaded.path  = GetExeDir() + L"Plugins\\" + std::wstring(fileName);
        loaded.module.reset(LoadLibraryExW(loaded.path.c_str(), nullptr, 0));
        Check(static_cast<bool>(loaded.module), std::format(L"{}: loads for shared libcurl runtime proof", label).c_str(), success);
        if (loaded.module)
        {
            loaded.shutdown  = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(loaded.module.get(), "RedSalamanderPluginShutdown"));
            loaded.canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(loaded.module.get(), "RedSalamanderPluginCanUnloadNow"));
            loaded.probe     = reinterpret_cast<PfnDebugCurlRuntimeProbe>(GetProcAddress(loaded.module.get(), "RedSalamanderDebugCurlRuntimeProbe"));
            Check(loaded.shutdown && loaded.canUnload && loaded.probe,
                  std::format(L"{}: resolves shared libcurl lifecycle and probe exports", label).c_str(),
                  success);
        }
        return loaded;
    };

    const auto unloadOne = [&](RuntimeModule& target, unsigned int cycle)
    {
        target.shutdown();
        Check(target.canUnload() == TRUE,
              std::format(L"{}: refresh cycle {} reaches its quiet point while its peer survives", target.label, cycle).c_str(),
              success);
        target.module.reset();
        Check(GetModuleHandleW(target.path.c_str()) == nullptr, std::format(L"{}: refresh cycle {} physically unloads", target.label, cycle).c_str(), success);
    };

    constexpr unsigned int kRefreshCycles = 8u;
    for (unsigned int cycle = 1u; cycle <= kRefreshCycles; ++cycle)
    {
        RuntimeModule curl   = loadModule(L"FileSystemCurl.dll", L"FileSystemCurl.dll");
        RuntimeModule google = loadModule(L"FileSystemGoogleDrive.dll", L"FileSystemGoogleDrive.dll");
        if (! curl.module || ! google.module || ! curl.shutdown || ! google.shutdown || ! curl.canUnload || ! google.canUnload || ! curl.probe ||
            ! google.probe)
        {
            return;
        }

        Check(curl.probe() == S_OK && google.probe() == S_OK,
              std::format(L"shared libcurl refresh cycle {} initializes both plugin participants", cycle).c_str(),
              success);

        RuntimeModule& first    = (cycle % 2u) != 0u ? curl : google;
        RuntimeModule& survivor = (cycle % 2u) != 0u ? google : curl;
        unloadOne(first, cycle);
        Check(survivor.probe() == S_OK,
              std::format(L"{}: refresh cycle {} still creates a libcurl easy handle after peer unload", survivor.label, cycle).c_str(),
              success);
        unloadOne(survivor, cycle);
    }
}

} // namespace

int wmain(int argc, wchar_t** argv)
{
#if defined(RS_ASAN_DEBUG_BUILD)
    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--asan-seed-heap-overflow")
    {
        auto storage                = std::make_unique<unsigned char[]>(8u);
        volatile unsigned char* raw = storage.get();
        raw[16]                     = 0x5Au;
        return raw[16] == 0x5Au ? 0 : 1;
    }
#else
    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--asan-seed-heap-overflow")
    {
        std::wcerr << L"The seeded heap probe requires the ASan Debug configuration.\n";
        return 2;
    }
#endif

    const bool packageSmoke = argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--package-smoke";

    // Configure the DLL search path so that plugin DLLs find their transitive dependencies:
    // - Plugins\ contains aws-*.dll, libcurl-d.dll, sqlite3.dll, and so on.
    // - The exe dir (parent of Plugins\) contains Common.dll and other shared DLLs.
    // We use SetDllDirectoryW to add Plugins\ to the standard DLL search path (it is inserted
    // between the app dir and the system/PATH directories in the standard search order).
    const std::wstring exeDir     = GetExeDir();
    const std::wstring pluginsDir = exeDir + L"Plugins";
    SetDllDirectoryW(pluginsDir.c_str());
    // AddDllDirectory is also called for LOAD_LIBRARY_SEARCH_USER_DIRS coverage.
    const DLL_DIRECTORY_COOKIE pluginsDirCookie = AddDllDirectory(pluginsDir.c_str());

    bool success = true;

    std::wcout << L"[ RUN      ] Step1: Enumerated metadata count contract\n";
    Check(! IsValidEnumeratedPluginCount(0u), L"enumerated metadata count rejects zero", success);
    Check(IsValidEnumeratedPluginCount(1u), L"enumerated metadata count accepts one", success);
    Check(IsValidEnumeratedPluginCount(kMaxEnumeratedPluginsPerModule), L"enumerated metadata count accepts the documented maximum", success);
    Check(! IsValidEnumeratedPluginCount(kMaxEnumeratedPluginsPerModule + 1u), L"enumerated metadata count rejects an implausible range", success);
    PluginMetaData metadata{};
    Check(! IsValidEnumeratedPluginRange(nullptr, 1u), L"enumerated metadata range rejects a null pointer with a nonzero count", success);
    Check(! IsValidEnumeratedPluginRange(&metadata, 0u), L"enumerated metadata range rejects a nonnull pointer with a zero count", success);
    Check(IsValidEnumeratedPluginRange(&metadata, 1u), L"enumerated metadata range accepts a nonnull pointer with a valid count", success);

    const PluginFactoryEntry validFactoryEntries[] = {
        {&GetFactoryContractMetaData0, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
        {&GetFactoryContractMetaData1, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
    };
    const PluginMetaData* enumeratedMetaData = nullptr;
    unsigned int enumeratedCount             = 0u;
    Check(FactoryEnumeratePlugins<IViewer>(validFactoryEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) == S_OK &&
              enumeratedMetaData == g_factoryContractMetaData.data() && enumeratedCount == static_cast<unsigned int>(std::size(validFactoryEntries)),
          L"factory enumeration accepts a valid contiguous bounded metadata array",
          success);

    const PluginFactoryEntry nonContiguousEntries[] = {
        {&GetFactoryContractMetaData0, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
        {&GetNonContiguousFactoryMetaData, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
    };
    Check(FactoryEnumeratePlugins<IViewer>(nonContiguousEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"factory enumeration rejects non-contiguous producer metadata",
          success);

    const PluginFactoryEntry nullMetadataEntry[] = {{nullptr, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance}};
    Check(FactoryEnumeratePlugins<IViewer>(nullMetadataEntry, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"factory enumeration rejects a null producer metadata callback",
          success);

    std::array<PluginFactoryEntry, kMaxEnumeratedPluginsPerModule + 1u> excessiveFactoryEntries{};
    excessiveFactoryEntries.fill(validFactoryEntries[0]);
    Check(FactoryEnumeratePlugins<IViewer>(excessiveFactoryEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_TOO_MANY_NAMES),
          L"factory enumeration rejects a producer count above the ABI maximum",
          success);

    std::wcout << L"[ RUN      ] Step1b: Packed FileInfo owner contracts\n";
    TestPackedFileInfoBuffer(success);

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - filesystem plugins\n";
    success = RunEnumerateAndSchemaPass(kFilesystemDlls, __uuidof(IFileSystem)) && success;

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - viewer plugins\n";
    success = RunEnumerateAndSchemaPass(kViewerDlls, __uuidof(IViewer)) && success;

    std::wcout << L"[ RUN      ] Step3: GetCapabilities - filesystem plugins\n";
    success = RunCapabilitiesPass(kFilesystemDlls) && success;

    std::wcout << L"[ RUN      ] Step4: Negative bogus plugin id\n";
    {
        bool negSuccess = true;
        TestBogusPluginId(negSuccess);
        success = negSuccess && success;
    }

    if (! packageSmoke)
    {
        std::wcout << L"[ RUN      ] Step5: Configuration-gated plugin debug selftests\n";
        {
            bool debugSuccess = true;
            Check(! IsDebugSelfTestExportPresenceAccepted(false, true), L"missing required debug selftest export turns the step red", debugSuccess);
            Check(IsDebugSelfTestExportPresenceAccepted(false, false), L"missing configuration-gated debug selftest export is an explicit skip", debugSuccess);

            TestPluginDebugSelfTests(
                L"Plugins\\FileSystem.dll", "RedSalamanderFileSystemDebugSelfTests", L"FileSystem.dll", kDebugSelfTestExportsRequired, debugSuccess);
            TestPluginDebugSelfTests(
                L"Plugins\\FileSystem7z.dll", "RedSalamander7zDebugSelfTests", L"FileSystem7z.dll", kDebugSelfTestExportsRequired, debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemMicrosoftDrive.dll",
                                     "RedSalamanderMicrosoftDriveDebugSelfTests",
                                     L"FileSystemMicrosoftDrive.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemGoogleDrive.dll",
                                     "RedSalamanderGoogleDriveDebugSelfTests",
                                     L"FileSystemGoogleDrive.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(
                L"Plugins\\FileSystemS3.dll", "RedSalamanderS3DebugSelfTests", L"FileSystemS3.dll", kDebugSelfTestExportsRequired, debugSuccess);
            TestPluginDebugSelfTests(
                L"Plugins\\FileSystemCurl.dll", "RedSalamanderCurlDebugSelfTests", L"FileSystemCurl.dll", kDebugSelfTestExportsRequired, debugSuccess);
            success = debugSuccess && success;
        }

        std::wcout << L"[ RUN      ] Step6: S3 runtime-refresh unload quiet point\n";
        TestS3RuntimeUnloadContract(success);

        std::wcout << L"[ RUN      ] Step7: Cross-plugin libcurl runtime-refresh survivor proof\n";
        TestCrossPluginCurlRuntimeUnloadContract(success);
    }

    if (pluginsDirCookie != nullptr)
    {
        RemoveDllDirectory(pluginsDirCookie);
    }

    std::wcout << (success ? (packageSmoke ? L"PluginContractTests package smoke passed.\n" : L"PluginContractTests passed.\n")
                           : L"PluginContractTests FAILED.\n");
    return success ? 0 : 1;
}

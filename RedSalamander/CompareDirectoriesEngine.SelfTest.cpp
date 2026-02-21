#include "CompareDirectoriesEngine.SelfTest.h"

#ifdef _DEBUG

#include "Framework.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "CompareDirectoriesEngine.h"
#include "CrashHandler.h"
#include "Helpers.h"
#include "SelfTestCommon.h"

namespace
{
constexpr std::wstring_view kBuiltinLocalFileSystemId = L"builtin/file-system";
constexpr std::wstring_view kBuiltinDummyFileSystemId = L"builtin/file-system-dummy";
struct TestState;
TestState* g_activeCompareState = nullptr;

void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, message);
    SelfTest::AppendSelfTestTrace(message);
}

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept;

[[nodiscard]] std::wstring MakeGuidText() noexcept
{
    GUID guid{};
    if (FAILED(::CoCreateGuid(&guid)))
    {
        return {};
    }

    wchar_t buffer[64]{};
    if (::StringFromGUID2(guid, buffer, static_cast<int>(std::size(buffer))) <= 0)
    {
        return {};
    }

    std::wstring text(buffer);
    text.erase(std::remove_if(text.begin(), text.end(), [](wchar_t ch) noexcept { return ch == L'{' || ch == L'}'; }), text.end());
    return text;
}

[[nodiscard]] bool SetFileLastWriteTime(const std::filesystem::path& path, const FILETIME& lastWriteTime) noexcept
{
    wil::unique_handle file(::CreateFileW(
        path.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    return ::SetFileTime(file.get(), nullptr, nullptr, &lastWriteTime) != 0;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetLocalFileSystem() noexcept
{
    return SelfTest::GetFileSystem(kBuiltinLocalFileSystemId);
}

[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return true;
    }

    if (text.size() < prefix.size())
    {
        return false;
    }

    if (prefix.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return CompareStringOrdinal(text.data(), static_cast<int>(prefix.size()), prefix.data(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}

class ShortReadFileReader final : public IFileReader
{
public:
    ShortReadFileReader(wil::com_ptr<IFileReader> inner, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _inner(std::move(inner)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
    }

    ShortReadFileReader(const ShortReadFileReader&)            = delete;
    ShortReadFileReader(ShortReadFileReader&&)                 = delete;
    ShortReadFileReader& operator=(const ShortReadFileReader&) = delete;
    ShortReadFileReader& operator=(ShortReadFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->GetSize(sizeBytes);
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->Seek(offset, origin, newPosition);
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _inner)
        {
            return E_FAIL;
        }

        const unsigned long capped = std::min(bytesToRead, _maxBytesPerRead);
        if (_delayMs != 0)
        {
            ::Sleep(_delayMs);
        }
        return _inner->Read(buffer, capped, bytesRead);
    }

private:
    ~ShortReadFileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileReader> _inner;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

// ShortReadFileSystem wraps a real IFileSystem/IFileSystemIO and limits every Read()
// call to at most maxBytesPerRead bytes.  Used as a regression guard to verify that
// the content-comparison engine handles partial reads correctly (i.e. it never assumes
// a single Read() returns the full file).
class ShortReadFileSystem final : public IFileSystem, public IFileSystemIO
{
public:
    ShortReadFileSystem(wil::com_ptr<IFileSystem> base, std::filesystem::path shortReadRoot, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _base(std::move(base)),
          _shortReadRoot(std::move(shortReadRoot)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
        if (_base)
        {
            static_cast<void>(_base->QueryInterface(__uuidof(IFileSystemIO), _baseIo.put_void()));
        }
    }

    ShortReadFileSystem(const ShortReadFileSystem&)            = delete;
    ShortReadFileSystem(ShortReadFileSystem&&)                 = delete;
    ShortReadFileSystem& operator=(const ShortReadFileSystem&) = delete;
    ShortReadFileSystem& operator=(ShortReadFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
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

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItem(path, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItems(paths, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItems(items, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->GetCapabilities(jsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetAttributes(path, fileAttributes);
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (reader == nullptr)
        {
            return E_POINTER;
        }
        *reader = nullptr;

        if (! _baseIo)
        {
            return E_POINTER;
        }

        wil::com_ptr<IFileReader> inner;
        const HRESULT hr = _baseIo->CreateFileReader(path, inner.put());
        if (FAILED(hr) || ! inner)
        {
            return FAILED(hr) ? hr : E_FAIL;
        }

        const std::wstring_view pathText(path ? path : L"");
        const std::wstring rootText = _shortReadRoot.wstring();
        const bool shouldShortRead  = ! rootText.empty() && StartsWithNoCase(pathText, rootText);
        if (! shouldShortRead)
        {
            *reader = inner.detach();
            return S_OK;
        }

        auto* wrapper = new (std::nothrow) ShortReadFileReader(std::move(inner), _maxBytesPerRead, _delayMs);
        if (! wrapper)
        {
            return E_OUTOFMEMORY;
        }

        *reader = wrapper;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileWriter(path, flags, writer);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->SetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetItemProperties(path, jsonUtf8);
    }

private:
    ~ShortReadFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    wil::com_ptr<IFileSystemIO> _baseIo;
    std::filesystem::path _shortReadRoot;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateShortReadFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                  const std::filesystem::path& shortReadRoot,
                                                                  unsigned long maxBytesPerRead,
                                                                  DWORD delayMs) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) ShortReadFileSystem(base, shortReadRoot, maxBytesPerRead, delayMs);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetDummyFileSystem() noexcept
{
    return SelfTest::GetFileSystem(kBuiltinDummyFileSystemId);
}

[[nodiscard]] bool CreateFileSystemIo(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemIO>& outIo) noexcept
{
    outIo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemIO), outIo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outIo);
}

[[nodiscard]] bool CreateInformations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IInformations>& outInfo) noexcept
{
    outInfo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IInformations), outInfo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outInfo);
}

[[nodiscard]] bool CreateFileSystemDirectoryOperations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemDirectoryOperations>& outOps) noexcept
{
    outOps.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), outOps.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outOps);
}

[[nodiscard]] bool EnsureDirectoryExistsFsOps(const wil::com_ptr<IFileSystemDirectoryOperations>& ops, const std::filesystem::path& path) noexcept
{
    if (! ops)
    {
        return false;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    for (const auto& part : normalized.relative_path())
    {
        current /= part;
        const HRESULT hr = ops->CreateDirectory(current.c_str());
        if (SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            continue;
        }
        return false;
    }

    return true;
}

[[nodiscard]] bool WriteFileBytesFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, const void* data, size_t sizeBytes) noexcept
{
    if (! io || ! data)
    {
        return false;
    }

    if (sizeBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const HRESULT createHr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    unsigned long written = 0;
    const HRESULT writeHr = writer->Write(data, static_cast<unsigned long>(sizeBytes), &written);
    if (FAILED(writeHr) || written != static_cast<unsigned long>(sizeBytes))
    {
        return false;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WriteFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string_view text) noexcept
{
    return WriteFileBytesFsIo(io, path, text.data(), text.size());
}

#ifndef SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE
#define SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE 0x2u
#endif

[[nodiscard]] bool TryCreateDirectorySymlink(const std::filesystem::path& linkPath, const std::filesystem::path& targetPath) noexcept
{
    const DWORD flags = SYMBOLIC_LINK_FLAG_DIRECTORY;

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags | SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE) != 0)
    {
        return true;
    }

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags) != 0)
    {
        return true;
    }

    return false;
}

constexpr std::array<std::wstring_view, 26> kCompareCaseNames{{
    L"unique",
    L"typemismatch",
    L"size",
    L"time",
    L"attributes",
    L"content",
    L"content short reads",
    L"subdir pending",
    L"subdirs",
    L"subdirattrs",
    L"missing folder",
    L"reparse",
    L"dummy_content",
    L"deep_tree",
    L"invalidate",
    L"ignore",
    L"showIdentical",
    L"setCompareEnabled",
    L"invalidateForPath",
    L"decisionUpdatedCallback",
    L"uiVersion",
    L"accessors",
    L"baseInterfaces",
    L"contentCacheHit",
    L"zeroByteContent",
    L"setSettingsInvalidates",
}};

[[nodiscard]] bool WriteFileFill(const std::filesystem::path& path, char ch, size_t sizeBytes) noexcept
{
    if (sizeBytes == 0)
    {
        return SelfTest::WriteBinaryFile(path, {});
    }

    const std::string text(sizeBytes, ch);
    const std::span<const char> textBytes(text.data(), text.size());
    return SelfTest::WriteBinaryFile(path, std::as_bytes(textBytes));
}

struct CaseFolders
{
    std::filesystem::path left;
    std::filesystem::path right;
};

[[nodiscard]] std::optional<CaseFolders> CreateCaseFolders(const std::filesystem::path& base, std::wstring_view caseName) noexcept
{
    std::filesystem::path caseRoot = base / std::filesystem::path(caseName);
    std::filesystem::path left     = caseRoot / L"left";
    std::filesystem::path right    = caseRoot / L"right";

    SelfTest::EnsureDirectory(left);
    SelfTest::EnsureDirectory(right);
    if (! SelfTest::PathExists(left) || ! SelfTest::PathExists(right))
    {
        return std::nullopt;
    }

    return CaseFolders{std::move(left), std::move(right)};
}

struct TestState
{
    bool failed = false;
    std::wstring caseFailureMessage;
    std::wstring failureMessage;
    SelfTest::SelfTestOptions options;
    std::vector<SelfTest::SelfTestCaseResult> caseResults;

    std::wstring currentCaseName;
    size_t currentCaseIndex   = kCompareCaseNames.size();
    size_t completedCaseIndex = static_cast<size_t>(-1);
    bool caseInProgress       = false;
    bool caseFailedFlag       = false;
    uint64_t caseStartMs      = 0;

    void BeginCase(std::wstring_view name) noexcept
    {
        EndCase();
        currentCaseName  = name;
        currentCaseIndex = kCompareCaseNames.size();
        for (size_t i = 0; i < kCompareCaseNames.size(); ++i)
        {
            if (kCompareCaseNames[i] == name)
            {
                currentCaseIndex = i;
                break;
            }
        }

        caseInProgress = true;
        caseFailedFlag = false;
        caseFailureMessage.clear();
        caseStartMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
    }

    void EndCase() noexcept
    {
        if (! caseInProgress)
        {
            return;
        }

        const auto now =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
        const uint64_t durationMs = (now >= caseStartMs) ? (now - caseStartMs) : 0;

        SelfTest::SelfTestCaseResult item{};
        item.name       = currentCaseName;
        item.status     = caseFailedFlag ? SelfTest::SelfTestCaseResult::Status::failed : SelfTest::SelfTestCaseResult::Status::passed;
        item.durationMs = durationMs;
        if (caseFailedFlag && ! caseFailureMessage.empty())
        {
            item.reason = caseFailureMessage;
        }

        if (currentCaseIndex < kCompareCaseNames.size())
        {
            completedCaseIndex = std::max(completedCaseIndex, currentCaseIndex);
        }

        caseResults.push_back(std::move(item));

        caseInProgress = false;
        currentCaseName.clear();
        currentCaseIndex = kCompareCaseNames.size();
        caseFailedFlag   = false;
        caseStartMs      = 0;
        caseFailureMessage.clear();
    }

    bool caseFailed() const noexcept
    {
        return caseFailedFlag;
    }

    void Require(bool condition, std::wstring_view message) noexcept
    {
        if (condition)
        {
            return;
        }

        failed = true;
        if (! caseFailedFlag)
        {
            caseFailureMessage = message;
            caseFailedFlag     = true;
        }
        if (failureMessage.empty())
        {
            failureMessage = message;
        }
        Debug::Error(L"CompareSelfTest: {0}", message);
    }

    SelfTest::SelfTestSuiteResult GetResult(uint64_t durationMs) noexcept
    {
        EndCase();

        if (options.failFast && failed)
        {
            const size_t start = (completedCaseIndex == static_cast<size_t>(-1)) ? 0u : (completedCaseIndex + 1u);
            for (size_t i = start; i < kCompareCaseNames.size(); ++i)
            {
                SelfTest::SelfTestCaseResult skipped{};
                skipped.name       = kCompareCaseNames[i];
                skipped.status     = SelfTest::SelfTestCaseResult::Status::skipped;
                skipped.reason     = L"skipped by fail-fast";
                skipped.durationMs = 0;
                caseResults.push_back(std::move(skipped));
            }
        }

        SelfTest::SelfTestSuiteResult result{};
        result.suite      = SelfTest::SelfTestSuite::CompareDirectories;
        result.durationMs = durationMs;
        result.cases      = std::move(caseResults);

        for (const auto& item : result.cases)
        {
            switch (item.status)
            {
                case SelfTest::SelfTestCaseResult::Status::passed: ++result.passed; break;
                case SelfTest::SelfTestCaseResult::Status::failed: ++result.failed; break;
                case SelfTest::SelfTestCaseResult::Status::skipped: ++result.skipped; break;
            }
        }

        if (! failureMessage.empty())
        {
            result.failureMessage = failureMessage;
        }

        return result;
    }
};

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept
{
    Trace(message);

    if (TestState* state = g_activeCompareState)
    {
        constexpr std::wstring_view kCasePrefix = L"Case: ";
        if (message.size() >= kCasePrefix.size() && message.substr(0, kCasePrefix.size()) == kCasePrefix)
        {
            const std::wstring_view caseName = message.substr(kCasePrefix.size());
            if (caseName.find(L':') == std::wstring_view::npos)
            {
                state->BeginCase(caseName);
            }
        }
    }
}

[[nodiscard]] std::vector<std::wstring>
EnumerateDirectoryNames(const wil::com_ptr<IFileSystem>& fs, const std::filesystem::path& folder, TestState& state) noexcept
{
    if (! fs)
    {
        state.Require(false, L"EnumerateDirectoryNames: file system is null.");
        return {};
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = fs->ReadDirectoryInfo(folder.c_str(), info.put());
    state.Require(SUCCEEDED(hr), L"EnumerateDirectoryNames: ReadDirectoryInfo failed.");
    if (FAILED(hr) || ! info)
    {
        return {};
    }

    FileInfo* head         = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    state.Require(SUCCEEDED(hrBuffer), L"EnumerateDirectoryNames: GetBuffer failed.");
    if (FAILED(hrBuffer) || head == nullptr)
    {
        return {};
    }

    std::vector<std::wstring> result;
    for (FileInfo* entry = head; entry != nullptr;)
    {
        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        result.emplace_back(entry->FileName, nameChars);

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return result;
}

[[nodiscard]] bool ContainsName(const std::vector<std::wstring>& names, std::wstring_view name) noexcept
{
    return std::any_of(names.begin(), names.end(), [&](const std::wstring& value) noexcept { return value == name; });
}

struct GetDecisionSehContext
{
    CompareDirectoriesSession* session                                   = nullptr;
    std::shared_ptr<const CompareDirectoriesFolderDecision>* outDecision = nullptr;
};

void InvokeGetRootDecision(void* rawContext) noexcept
{
    auto* ctx = static_cast<GetDecisionSehContext*>(rawContext);
    if (! ctx || ! ctx->session || ! ctx->outDecision)
    {
        return;
    }

    *ctx->outDecision = ctx->session->GetOrComputeDecision(std::filesystem::path{});
}

[[nodiscard]] bool TryGetRootDecisionWithSeh(CompareDirectoriesSession& session, std::shared_ptr<const CompareDirectoriesFolderDecision>& outDecision) noexcept
{
    GetDecisionSehContext ctx{};
    ctx.session     = &session;
    ctx.outDecision = &outDecision;

    __try
    {
        InvokeGetRootDecision(&ctx);
        return true;
    }
    __except (CrashHandler::WriteDumpForException(GetExceptionInformation()))
    {
        return false;
    }
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> ComputeRootDecision(wil::com_ptr<IFileSystem> baseFs,
                                                                                          const CaseFolders& folders,
                                                                                          Common::Settings::CompareDirectoriesSettings settings,
                                                                                          TestState& state) noexcept
{
    if (! baseFs)
    {
        state.Require(false, L"Base file system is null.");
        return {};
    }

    auto session = std::make_shared<CompareDirectoriesSession>(std::move(baseFs), folders.left, folders.right, std::move(settings));
    std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    if (! TryGetRootDecisionWithSeh(*session, decision))
    {
        state.Require(false, L"GetOrComputeDecision crashed.");
        return {};
    }
    state.Require(static_cast<bool>(decision), L"GetOrComputeDecision returned null.");
    if (! decision)
    {
        return {};
    }

    state.Require(SUCCEEDED(decision->hr), L"Decision hr is failure.");
    return decision;
}

[[nodiscard]] const CompareDirectoriesItemDecision* FindItem(const CompareDirectoriesFolderDecision& decision, std::wstring_view name) noexcept
{
    const auto it = decision.items.find(std::wstring(name));
    if (it == decision.items.end())
    {
        return nullptr;
    }
    return &it->second;
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> WaitForContentCompare(const std::shared_ptr<CompareDirectoriesSession>& session,
                                                                                            const std::filesystem::path& relativeFolder,
                                                                                            std::wstring_view itemName,
                                                                                            TestState& state) noexcept
{
    if (! session)
    {
        state.Require(false, L"WaitForContentCompare: session is null.");
        return {};
    }

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(5);
    while (std::chrono::steady_clock::now() < deadline)
    {
        auto decision = session->GetOrComputeDecision(relativeFolder);
        state.Require(static_cast<bool>(decision), L"WaitForContentCompare: decision is null.");
        if (! decision)
        {
            return {};
        }

        const auto* item = FindItem(*decision, itemName);
        state.Require(item != nullptr, std::format(L"WaitForContentCompare: item missing: {}.", itemName));
        if (! item)
        {
            return decision;
        }

        if (! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending))
        {
            return decision;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }

    state.Require(false, std::format(L"Timed out waiting for content compare: {}.", itemName));
    return session->GetOrComputeDecision(relativeFolder);
}
} // namespace

bool CompareDirectoriesSelfTest::Run(const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    Debug::Info(L"CompareSelfTest: begin");
    AppendCompareSelfTestTraceLine(L"Run: begin");

    TestState state{};
    state.options               = options;
    g_activeCompareState        = &state;
    const auto clearActiveState = wil::scope_exit([&] { g_activeCompareState = nullptr; });

    wil::com_ptr<IFileSystem> baseFs = GetLocalFileSystem();
    if (! baseFs)
    {
        state.Require(false, L"CompareSelfTest: local file system plugin not available.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::CompareDirectories);
    if (suiteRoot.empty())
    {
        state.Require(false, L"CompareSelfTest: suite artifact root not available.");
    }

    const std::filesystem::path root = suiteRoot / L"work";
    if (! state.failed)
    {
        state.Require(SelfTest::EnsureDirectory(root), L"CompareSelfTest: failed to create work root folder.");
    }
    AppendCompareSelfTestTraceLine(L"Run: root created");

    std::wstring guid = MakeGuidText();
    if (guid.empty())
    {
        guid = L"0";
    }

    wil::com_ptr<IFileSystem> dummyFs = GetDummyFileSystem();
    wil::com_ptr<IInformations> dummyInfo;
    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemDirectoryOperations> dummyOps;

    if (! dummyFs)
    {
        state.Require(false, L"CompareSelfTest: FileSystemDummy plugin not available.");
    }
    else
    {
        AppendCompareSelfTestTraceLine(L"Run: dummy plugin setup");
        state.Require(CreateInformations(dummyFs, dummyInfo), L"CompareSelfTest: FileSystemDummy missing IInformations.");

        const HRESULT setHr =
            dummyInfo ? dummyInfo->SetConfiguration("{\"maxChildrenPerDirectory\":0,\"maxDepth\":0,\"seed\":1,\"latencyMs\":0,\"virtualSpeedLimit\":\"0\"}")
                      : E_NOINTERFACE;
        state.Require(SUCCEEDED(setHr), L"CompareSelfTest: FileSystemDummy SetConfiguration failed.");

        state.Require(CreateFileSystemIo(dummyFs, dummyIo), L"CompareSelfTest: FileSystemDummy missing IFileSystemIO.");
        state.Require(CreateFileSystemDirectoryOperations(dummyFs, dummyOps), L"CompareSelfTest: FileSystemDummy missing IFileSystemDirectoryOperations.");
    }

    if (! baseFs || ! SelfTest::PathExists(root))
    {
        AppendCompareSelfTestTraceLine(L"Run: aborting due to setup failure");
    }
    else if (state.failed && options.failFast)
    {
        AppendCompareSelfTestTraceLine(L"Run: aborting due to earlier failure (fail-fast)");
    }
    else
    {
        do
        {
            const auto shouldAbort = [&]() noexcept
            {
                if (options.failFast && state.failed)
                {
                    AppendCompareSelfTestTraceLine(L"Run: aborting due to fail-fast");
                    return true;
                }
                return false;
            };

            // Case: Unique files/dirs selected; identical excluded by default.
            AppendCompareSelfTestTraceLine(L"Case: unique");
            if (const auto foldersOpt = CreateCaseFolders(root, L"unique"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"only_left.txt", "L"), L"Failed to create only_left.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"only_right.txt", "R"), L"Failed to create only_right.txt (right).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"only_left_dir"), L"Failed to create only_left_dir (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "S"), L"Failed to create same.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "S"), L"Failed to create same.txt (right).");

                AppendCompareSelfTestTraceLine(L"Case: unique: computing decision");
                auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
                AppendCompareSelfTestTraceLine(L"Case: unique: decision returned");
                if (decision)
                {
                    {
                        const auto* item = FindItem(*decision, L"only_left.txt");
                        state.Require(item != nullptr, L"only_left.txt missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDifferent, L"only_left.txt expected isDifferent.");
                            state.Require(item->selectLeft && ! item->selectRight, L"only_left.txt expected selectLeft only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                          L"only_left.txt expected differenceMask=OnlyInLeft.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"only_right.txt");
                        state.Require(item != nullptr, L"only_right.txt missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDifferent, L"only_right.txt expected isDifferent.");
                            state.Require(! item->selectLeft && item->selectRight, L"only_right.txt expected selectRight only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInRight),
                                          L"only_right.txt expected differenceMask=OnlyInRight.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"only_left_dir");
                        state.Require(item != nullptr, L"only_left_dir missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDirectory, L"only_left_dir expected isDirectory.");
                            state.Require(item->isDifferent, L"only_left_dir expected isDifferent.");
                            state.Require(item->selectLeft && ! item->selectRight, L"only_left_dir expected selectLeft only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                          L"only_left_dir expected differenceMask=OnlyInLeft.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"same.txt");
                        state.Require(item != nullptr, L"same.txt missing from decision.");
                        if (item)
                        {
                            state.Require(! item->isDifferent, L"same.txt expected identical.");
                            state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0.");
                        }
                    }

                    auto session =
                        std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);
                    AppendCompareSelfTestTraceLine(L"Case: unique: enumeration done");

                    state.Require(ContainsName(leftNames, L"only_left.txt"), L"only_left.txt expected in left enumeration.");
                    state.Require(! ContainsName(leftNames, L"only_right.txt"), L"only_right.txt unexpected in left enumeration.");
                    state.Require(! ContainsName(leftNames, L"same.txt"), L"same.txt expected excluded in left enumeration.");

                    state.Require(ContainsName(rightNames, L"only_right.txt"), L"only_right.txt expected in right enumeration.");
                    state.Require(! ContainsName(rightNames, L"only_left.txt"), L"only_left.txt unexpected in right enumeration.");
                    state.Require(! ContainsName(rightNames, L"same.txt"), L"same.txt expected excluded in right enumeration.");

                    AppendCompareSelfTestTraceLine(L"Case: unique: done");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: unique.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: File vs directory mismatch selects both sides.
            AppendCompareSelfTestTraceLine(L"Case: typemismatch");
            if (const auto foldersOpt = CreateCaseFolders(root, L"typemismatch"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"mix", "F"), L"Failed to create mix file (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"mix"), L"Failed to create mix directory (right).");

                auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"mix");
                    state.Require(item != nullptr, L"mix missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"mix expected isDifferent on type mismatch.");
                        state.Require(item->selectLeft && item->selectRight, L"mix expected select both on type mismatch.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::TypeMismatch), L"mix expected differenceMask=TypeMismatch.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: typemismatch.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Size compare selects bigger file.
            AppendCompareSelfTestTraceLine(L"Case: size");
            if (const auto foldersOpt = CreateCaseFolders(root, L"size"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'A', 200), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'B', 100), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSize = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareSize.");
                        state.Require(item->selectLeft && ! item->selectRight, L"a.bin expected selectLeft only when left is bigger.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Size), L"a.bin expected differenceMask=Size.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: size.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Date/time compare selects newer file.
            AppendCompareSelfTestTraceLine(L"Case: time");
            if (const auto foldersOpt = CreateCaseFolders(root, L"time"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "T"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "T"), L"Failed to create a.txt (right).");

                FILETIME now{};
                ::GetSystemTimeAsFileTime(&now);
                ULARGE_INTEGER newer{};
                newer.LowPart  = now.dwLowDateTime;
                newer.HighPart = now.dwHighDateTime;
                newer.QuadPart += 60ull * 10'000'000ull;

                FILETIME leftFt{};
                leftFt.dwLowDateTime  = newer.LowPart;
                leftFt.dwHighDateTime = newer.HighPart;

                state.Require(SetFileLastWriteTime(folders.left / L"a.txt", leftFt), L"Failed to set a.txt last write time (left).");
                state.Require(SetFileLastWriteTime(folders.right / L"a.txt", now), L"Failed to set a.txt last write time (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareDateTime = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.txt expected isDifferent with compareDateTime.");
                        state.Require(item->selectLeft && ! item->selectRight, L"a.txt expected selectLeft only when left is newer.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::DateTime), L"a.txt expected differenceMask=DateTime.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: time.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Attribute compare selects both sides.
            AppendCompareSelfTestTraceLine(L"Case: attributes");
            if (const auto foldersOpt = CreateCaseFolders(root, L"attributes"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

                const std::filesystem::path leftPath = folders.left / L"a.txt";
                const DWORD leftAttrs                = ::GetFileAttributesW(leftPath.c_str());
                state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for a.txt (left).");
                if (leftAttrs != INVALID_FILE_ATTRIBUTES)
                {
                    state.Require(::SetFileAttributesW(leftPath.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0,
                                  L"SetFileAttributesW failed for a.txt (left).");
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareAttributes = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.txt expected isDifferent with compareAttributes.");
                        state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when attributes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Attributes), L"a.txt expected differenceMask=Attributes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: attributes.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Content compare selects both sides.
            AppendCompareSelfTestTraceLine(L"Case: content");
            if (const auto foldersOpt = CreateCaseFolders(root, L"content"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'Y', 64), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent.");
                        state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when content differs.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending cleared after compare completes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Content compare tolerates short reads for equal files.
            AppendCompareSelfTestTraceLine(L"Case: content short reads");
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_shortreads"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'Z', 4096), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'Z', 4096), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1u, 0u);
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper.");

                auto session  = std::make_shared<CompareDirectoriesSession>(wrapped ? wrapped : baseFs, folders.left, folders.right, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"a.bin expected not different for equal content with short reads.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"a.bin expected Content bit cleared for equal content with short reads.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending cleared after compare completes (short reads).");
                        state.Require(! item->selectLeft && ! item->selectRight, L"a.bin expected no selection when equal.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_shortreads.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Subdirectory pending state + flush updates ancestors without navigation.
            AppendCompareSelfTestTraceLine(L"Case: subdir pending");
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdir_pending"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
                state.Require(WriteFileFill(folders.left / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent        = true;
                settings.compareSubdirectories = true;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1024u, 1u);
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (subdir pending).");

                auto session = std::make_shared<CompareDirectoriesSession>(wrapped ? wrapped : baseFs, folders.left, folders.right, settings);

                std::mutex progressMutex;
                std::condition_variable progressCv;
                bool contentDone = false;

                session->SetContentProgressCallback(
                    [&](uint32_t,
                        const std::filesystem::path&,
                        std::wstring_view,
                        uint64_t,
                        uint64_t,
                        uint64_t,
                        uint64_t,
                        uint64_t pendingContentCompares,
                        uint64_t totalContentCompares,
                        uint64_t completedContentCompares) noexcept
                    {
                        if (pendingContentCompares != 0u || totalContentCompares == 0u || completedContentCompares != totalContentCompares)
                        {
                            return;
                        }

                        std::lock_guard lock(progressMutex);
                        contentDone = true;
                        progressCv.notify_all();
                    });

                auto rootDecision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(rootDecision), L"subdir pending: root decision is null.");
                if (rootDecision)
                {
                    const auto* subItem = FindItem(*rootDecision, L"sub");
                    state.Require(subItem != nullptr, L"subdir pending: sub missing from root decision.");
                    if (subItem)
                    {
                        state.Require(subItem->isDirectory, L"subdir pending: sub expected isDirectory.");
                        state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                                      L"subdir pending: sub expected SubdirPending while content compare is running.");
                        state.Require(! HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                      L"subdir pending: sub expected no SubdirContent while only content compares are pending.");
                        state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different while pending.");
                        state.Require(! subItem->selectLeft && ! subItem->selectRight, L"subdir pending: sub expected not selected while pending.");
                    }
                }

                auto subDecision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                state.Require(static_cast<bool>(subDecision), L"subdir pending: sub decision is null.");
                if (subDecision)
                {
                    const auto* fileItem = FindItem(*subDecision, L"a.bin");
                    state.Require(fileItem != nullptr, L"subdir pending: a.bin missing from sub decision.");
                    if (fileItem)
                    {
                        state.Require(HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"subdir pending: a.bin expected ContentPending while content compare is running.");
                        state.Require(! HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"subdir pending: a.bin expected no Content bit while pending.");
                        state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different while pending.");
                        state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected while pending.");
                    }
                }

                {
                    std::unique_lock lock(progressMutex);
                    static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds(SelfTest::ScaleTimeout(30'000)), [&] { return contentDone; }));
                }
                state.Require(contentDone, L"subdir pending: timed out waiting for content compare to finish.");

                // Root decision remains in pending state until pending updates are flushed.
                auto rootBeforeFlush = session->GetOrComputeDecision(std::filesystem::path{});
                if (rootBeforeFlush)
                {
                    const auto* subItem = FindItem(*rootBeforeFlush, L"sub");
                    if (subItem)
                    {
                        state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                                      L"subdir pending: expected SubdirPending to remain until FlushPendingContentCompareUpdates.");
                    }
                }

                session->FlushPendingContentCompareUpdates();
                session->SetContentProgressCallback({});

                auto rootAfterFlush = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(rootAfterFlush), L"subdir pending: root decision missing after flush.");
                if (rootAfterFlush)
                {
                    const auto* subItem = FindItem(*rootAfterFlush, L"sub");
                    state.Require(subItem != nullptr, L"subdir pending: sub missing after flush.");
                    if (subItem)
                    {
                        state.Require(subItem->differenceMask == 0u, L"subdir pending: sub expected no difference mask after flush (equal subtree).");
                        state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different after flush (equal subtree).");
                        state.Require(! subItem->selectLeft && ! subItem->selectRight,
                                      L"subdir pending: sub expected not selected after flush (equal subtree).");
                    }
                }

                auto subAfterFlush = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                if (subAfterFlush)
                {
                    const auto* fileItem = FindItem(*subAfterFlush, L"a.bin");
                    state.Require(fileItem != nullptr, L"subdir pending: a.bin missing after flush.");
                    if (fileItem)
                    {
                        state.Require(fileItem->differenceMask == 0u, L"subdir pending: a.bin expected no difference mask after flush (equal).");
                        state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different after flush (equal).");
                        state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected after flush (equal).");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdir_pending.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Subdirectory content compare selects both directories.
            AppendCompareSelfTestTraceLine(L"Case: subdirs");
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdirs"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"child.txt", "C"), L"Failed to create sub\\child.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"sub");
                    state.Require(item != nullptr, L"sub missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"sub expected isDirectory.");
                        state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectories.");
                        state.Require(item->selectLeft && item->selectRight, L"sub expected select both when content differs.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent), L"sub expected differenceMask=SubdirContent.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdirs.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Compare attributes of subdirectories selects both.
            AppendCompareSelfTestTraceLine(L"Case: subdirattrs");
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdirattrs"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");

                const std::filesystem::path leftDir = folders.left / L"sub";
                const DWORD leftAttrs               = ::GetFileAttributesW(leftDir.c_str());
                state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for sub (left).");
                if (leftAttrs != INVALID_FILE_ATTRIBUTES)
                {
                    state.Require(::SetFileAttributesW(leftDir.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for sub (left).");
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectoryAttributes = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"sub");
                    state.Require(item != nullptr, L"sub missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"sub expected isDirectory.");
                        state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectoryAttributes.");
                        state.Require(item->selectLeft && item->selectRight, L"sub expected select both when attributes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirAttributes),
                                      L"sub expected differenceMask=SubdirAttributes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdirattrs.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Missing folder is reported without failing the decision.
            AppendCompareSelfTestTraceLine(L"Case: missing folder");
            if (const auto foldersOpt = CreateCaseFolders(root, L"missing_folder"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"a.txt", "A"), L"Failed to create sub\\a.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);
                auto decision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                state.Require(static_cast<bool>(decision), L"missing folder: decision is null.");
                if (decision)
                {
                    state.Require(SUCCEEDED(decision->hr), L"missing folder: expected decision hr success.");
                    state.Require(! decision->leftFolderMissing, L"missing folder: expected leftFolderMissing=false.");
                    state.Require(decision->rightFolderMissing, L"missing folder: expected rightFolderMissing=true.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: missing_folder.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Reparse points are not traversed for subdirectory comparison.
            AppendCompareSelfTestTraceLine(L"Case: reparse");
            if (const auto foldersOpt = CreateCaseFolders(root, L"reparse"))
            {
                const auto& folders                = *foldersOpt;
                const std::filesystem::path target = folders.left / L"target";
                state.Require(SelfTest::EnsureDirectory(target), L"Failed to create reparse target (left).");
                state.Require(SelfTest::WriteTextFile(target / L"child.txt", "C"), L"Failed to create target\\child.txt (left).");

                const std::filesystem::path linkPath = folders.left / L"sub";
                const bool linkCreated               = TryCreateDirectorySymlink(linkPath, target);
                if (! linkCreated)
                {
                    const DWORD err = ::GetLastError();
                    if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
                    {
                        Debug::Warning(L"CompareSelfTest: skipping reparse point test (CreateSymbolicLinkW failed: {0}).", err);
                    }
                    else
                    {
                        state.Require(false, std::format(L"CreateSymbolicLinkW failed unexpectedly: {}.", err));
                    }
                }
                else
                {
                    state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub directory (right).");

                    Common::Settings::CompareDirectoriesSettings settings{};
                    settings.compareSubdirectories = true;

                    auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                    if (decision)
                    {
                        const auto* item = FindItem(*decision, L"sub");
                        state.Require(item != nullptr, L"sub missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDirectory, L"sub expected isDirectory.");
                            state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                          L"sub expected SubdirContent not set for reparse points.");
                        }
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: reparse.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Dummy filesystem paths use plugin I/O for content compare (cross-filesystem support).
            AppendCompareSelfTestTraceLine(L"Case: dummy_content");
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"Y:\\") / (L"CompareSelfTest_" + guid) / L"compare";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create right root.");

                state.Require(WriteFileTextFsIo(dummyIo, leftRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (left).");
                state.Require(WriteFileTextFsIo(dummyIo, rightRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(dummyFs, leftRoot, rightRoot, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"Dummy: a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"Dummy: a.bin expected identical after content compare.");
                        state.Require(item->differenceMask == 0u, L"Dummy: a.bin expected differenceMask=0 after content compare.");
                    }
                }
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for cross-filesystem content compare test.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Deep directory trees do not overflow the stack (iterative traversal).
            AppendCompareSelfTestTraceLine(L"Case: deep_tree");
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"Z:\\") / (L"CompareSelfTest_" + guid) / L"deep";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create deep left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create deep right root.");

                constexpr size_t kDepth = 1024;

                std::filesystem::path leftPath  = leftRoot;
                std::filesystem::path rightPath = rightRoot;
                for (size_t i = 0; i < kDepth; ++i)
                {
                    const std::wstring name = std::format(L"d{:04}", i);
                    leftPath /= name;
                    rightPath /= name;
                    const HRESULT leftHr  = dummyOps->CreateDirectory(leftPath.c_str());
                    const HRESULT rightHr = dummyOps->CreateDirectory(rightPath.c_str());
                    state.Require(SUCCEEDED(leftHr) || leftHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                  std::format(L"Dummy: failed to create left dir at depth {}.", i));
                    state.Require(SUCCEEDED(rightHr) || rightHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                  std::format(L"Dummy: failed to create right dir at depth {}.", i));
                }

                state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "L"), L"Dummy: failed to create leaf.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto decision = ComputeRootDecision(dummyFs, CaseFolders{leftRoot, rightRoot}, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"d0000");
                    state.Require(item != nullptr, L"Dummy: d0000 missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"Dummy: d0000 expected isDirectory.");
                        state.Require(item->isDifferent, L"Dummy: d0000 expected isDifferent from deep leaf mismatch.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                      L"Dummy: d0000 expected differenceMask=SubdirContent from deep leaf mismatch.");
                    }
                }
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for deep tree test.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Version invalidation mid-scan does not cache stale results.
            AppendCompareSelfTestTraceLine(L"Case: invalidate");
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"W:\\") / (L"CompareSelfTest_" + guid) / L"invalidate";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create invalidate left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create invalidate right root.");

                constexpr size_t kDepth         = 256;
                std::filesystem::path leftPath  = leftRoot;
                std::filesystem::path rightPath = rightRoot;
                for (size_t i = 0; i < kDepth; ++i)
                {
                    const std::wstring name = std::format(L"d{}", i);
                    leftPath /= name;
                    rightPath /= name;
                    static_cast<void>(dummyOps->CreateDirectory(leftPath.c_str()));
                    static_cast<void>(dummyOps->CreateDirectory(rightPath.c_str()));
                }
                state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "X"), L"Dummy: failed to create invalidate leaf.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto session                 = std::make_shared<CompareDirectoriesSession>(dummyFs, leftRoot, rightRoot, settings);
                const uint64_t versionBefore = session->GetVersion();

                std::atomic<bool> scanStarted{false};
                session->SetScanProgressCallback(
                    [&](const std::filesystem::path& folder, std::wstring_view, uint64_t, uint64_t, uint32_t, uint64_t, uint64_t) noexcept
                    {
                        if (folder.empty())
                        {
                            scanStarted.store(true, std::memory_order_release);
                        }
                    });

                std::shared_ptr<const CompareDirectoriesFolderDecision> decisionBefore;
                std::jthread worker([&] { decisionBefore = session->GetOrComputeDecision(std::filesystem::path{}); });

                const auto startedDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
                while (! scanStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startedDeadline)
                {
                    std::this_thread::sleep_for(std::chrono::milliseconds(1));
                }

                state.Require(scanStarted.load(std::memory_order_acquire), L"Invalidate: scan did not start within timeout.");

                session->Invalidate();
                state.Require(session->GetVersion() == versionBefore + 1u, L"Invalidate: expected version bump.");

                worker.join();
                state.Require(static_cast<bool>(decisionBefore), L"Invalidate: initial decision missing.");

                const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionAfter), L"Invalidate: decision missing after invalidation.");
                if (decisionBefore && decisionAfter)
                {
                    state.Require(decisionAfter != decisionBefore, L"Invalidate: expected a new decision after invalidation (stale result cached).");
                }
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for invalidation test.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Ignore patterns exclude files/directories.
            AppendCompareSelfTestTraceLine(L"Case: ignore");
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"ignore_dir"), L"Failed to create ignore_dir (left).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"keep_dir"), L"Failed to create keep_dir (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles               = true;
                settings.ignoreFilesPatterns       = L"*.log";
                settings.ignoreDirectories         = true;
                settings.ignoreDirectoriesPatterns = L"ignore*";

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
                    state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
                    state.Require(FindItem(*decision, L"keep_dir") != nullptr, L"keep_dir expected in decision.");
                    state.Require(FindItem(*decision, L"ignore_dir") == nullptr, L"ignore_dir expected to be ignored.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: showIdenticalItems includes identical files.
            AppendCompareSelfTestTraceLine(L"Case: showIdentical");
            if (const auto foldersOpt = CreateCaseFolders(root, L"identical"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "SAME"), L"Failed to create same.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "SAME"), L"Failed to create same.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                auto session       = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);
                const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                const uint64_t versionBefore = session->GetVersion();
                const auto decisionBefore    = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionBefore), L"Decision missing (before showIdentical).");
                if (decisionBefore)
                {
                    const auto* item = FindItem(*decisionBefore, L"same.txt");
                    state.Require(item != nullptr, L"same.txt missing from decision (before showIdentical).");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"same.txt expected identical (before showIdentical).");
                        state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (before showIdentical).");
                    }
                }

                state.Require(! ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                              L"same.txt expected excluded from left enumeration (before showIdentical).");
                state.Require(! ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                              L"same.txt expected excluded from right enumeration (before showIdentical).");

                settings.showIdenticalItems = true;
                session->SetSettings(settings);

                const uint64_t versionAfter = session->GetVersion();
                state.Require(versionAfter == versionBefore, L"SetSettings(showIdenticalItems) should not invalidate decisions.");

                const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(decisionAfter == decisionBefore, L"Decision should remain cached across showIdenticalItems toggle.");

                state.Require(ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                              L"same.txt expected included in left enumeration (after showIdentical).");
                state.Require(ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                              L"same.txt expected included in right enumeration (after showIdentical).");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: identical.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: SetCompareEnabled(false) stops producing decisions; re-enabling resumes.
            AppendCompareSelfTestTraceLine(L"Case: setCompareEnabled");
            if (const auto foldersOpt = CreateCaseFolders(root, L"setCompareEnabled"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"b.txt", "B"), L"Failed to create b.txt (right).");

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true by default.");

                // When compare is disabled, ReadDirectoryInfo falls back to the base filesystem
                // and shows all files (no comparison filtering applied).
                session->SetCompareEnabled(false);
                state.Require(! session->IsCompareEnabled(), L"IsCompareEnabled should be false after SetCompareEnabled(false).");

                {
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

                    // Disabled compare: both sides should see their own files unfiltered.
                    state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be visible in left when compare is disabled.");
                    state.Require(ContainsName(rightNames, L"b.txt"), L"setCompareEnabled: b.txt should be visible in right when compare is disabled.");
                    // a.txt only exists on the left, b.txt only exists on the right — in enabled mode
                    // they would be filtered to their own pane; disabled should expose them as-is.
                    state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in the left pane.");
                    state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in the right pane.");
                }

                session->SetCompareEnabled(true);
                state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true after re-enabling.");

                // After re-enabling, decisions should be obtainable and filtering should be back.
                auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"GetOrComputeDecision should succeed after re-enabling compare.");

                {
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

                    // Re-enabled compare: only pane-relevant different items are shown.
                    state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be shown in left pane after re-enable (only in left).");
                    state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in left pane after re-enable.");
                    state.Require(ContainsName(rightNames, L"b.txt"),
                                  L"setCompareEnabled: b.txt should be shown in right pane after re-enable (only in right).");
                    state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in right pane after re-enable.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: setCompareEnabled.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: InvalidateForAbsolutePath invalidates only the targeted subtree.
            AppendCompareSelfTestTraceLine(L"Case: invalidateForPath");
            if (const auto foldersOpt = CreateCaseFolders(root, L"invalidateForPath"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub1"), L"Failed to create sub1 (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub1"), L"Failed to create sub1 (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (right).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub2"), L"Failed to create sub2 (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub2"), L"Failed to create sub2 (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (right).");

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                // Warm up both subtrees.
                const auto decisionSub1Before = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
                const auto decisionSub2Before = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));
                state.Require(static_cast<bool>(decisionSub1Before), L"sub1 decision missing before invalidate.");
                state.Require(static_cast<bool>(decisionSub2Before), L"sub2 decision missing before invalidate.");

                // Invalidate only sub1's absolute path.
                session->InvalidateForAbsolutePath(folders.left / L"sub1", /*includeSubtree=*/true);

                const auto decisionSub1After = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
                const auto decisionSub2After = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));

                state.Require(static_cast<bool>(decisionSub1After), L"sub1 decision missing after invalidate.");
                state.Require(static_cast<bool>(decisionSub2After), L"sub2 decision missing after invalidate.");

                // Sub1 must be a different (newly computed) decision object.
                state.Require(decisionSub1After != decisionSub1Before, L"sub1 decision should be new after InvalidateForAbsolutePath.");
                // Sub2 must be the same cached object — it was not invalidated.
                state.Require(decisionSub2After == decisionSub2Before, L"sub2 decision should remain cached (not invalidated).");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: invalidateForPath.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: SetDecisionUpdatedCallback fires after Invalidate().
            AppendCompareSelfTestTraceLine(L"Case: decisionUpdatedCallback");
            if (const auto foldersOpt = CreateCaseFolders(root, L"decisionUpdatedCallback"))
            {
                const auto& folders = *foldersOpt;
                // Use compareContent=true with same-size but byte-different files so a content-compare
                // job is enqueued and dispatched to a worker thread.  The callback fires on that worker
                // thread when the compare job completes (size-different files are short-circuited without
                // an async job and would never fire the callback).
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);

                std::atomic<int> callbackCount{0};
                session->SetDecisionUpdatedCallback([&]() noexcept { callbackCount.fetch_add(1, std::memory_order_relaxed); });

                // Trigger a scan so content-compare workers are started.
                static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

                // Wait up to 10 s for the callback to fire at least once, polling GetOrComputeDecision
                // to keep the scan driving (consistent with the WaitForContentCompare pattern).
                const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(SelfTest::ScaleTimeout(10'000));
                while (callbackCount.load(std::memory_order_relaxed) == 0 && std::chrono::steady_clock::now() < deadline)
                {
                    static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }

                state.Require(callbackCount.load(std::memory_order_relaxed) > 0,
                              L"DecisionUpdatedCallback must fire at least once after content compare completes.");

                // Unregister before session is destroyed to avoid dangling reference.
                session->SetDecisionUpdatedCallback(nullptr);
            }
            else
            {
                state.Require(false, L"Failed to create case folders: decisionUpdatedCallback.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: GetUiVersion increments on Invalidate() and after FlushPendingContentCompareUpdates().
            AppendCompareSelfTestTraceLine(L"Case: uiVersion");
            if (const auto foldersOpt = CreateCaseFolders(root, L"uiVersion"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const uint64_t uiV0 = session->GetUiVersion();
                const uint64_t ver0 = session->GetVersion();

                session->Invalidate();
                const uint64_t uiV1 = session->GetUiVersion();
                const uint64_t ver1 = session->GetVersion();

                state.Require(uiV1 != uiV0, L"GetUiVersion should change after Invalidate().");
                state.Require(ver1 != ver0, L"GetVersion should change after Invalidate().");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: uiVersion.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Accessor getters return correct values after construction.
            AppendCompareSelfTestTraceLine(L"Case: accessors");
            if (const auto foldersOpt = CreateCaseFolders(root, L"accessors"))
            {
                const auto& folders = *foldersOpt;
                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSize = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);

                state.Require(session->GetRoot(ComparePane::Left) == folders.left, L"GetRoot(Left) should match the left root passed to constructor.");
                state.Require(session->GetRoot(ComparePane::Right) == folders.right, L"GetRoot(Right) should match the right root passed to constructor.");
                state.Require(session->GetSettings().compareSize == settings.compareSize,
                              L"GetSettings().compareSize should match the value passed to constructor.");

                // TryMakeRelative / ResolveAbsolute round-trip.
                const std::filesystem::path sub(L"subdir");
                const std::filesystem::path absLeft = folders.left / sub;
                const auto relOpt                   = session->TryMakeRelative(ComparePane::Left, absLeft);
                state.Require(relOpt.has_value(), L"TryMakeRelative should succeed for a path under the left root.");
                if (relOpt.has_value())
                {
                    state.Require(relOpt.value() == sub, L"TryMakeRelative should return the expected relative path.");
                    const std::filesystem::path resolved = session->ResolveAbsolute(ComparePane::Left, relOpt.value());
                    state.Require(resolved == absLeft, L"ResolveAbsolute round-trip should match the original absolute path.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: accessors.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Base interface accessors return non-null objects after construction.
            AppendCompareSelfTestTraceLine(L"Case: baseInterfaces");
            if (const auto foldersOpt = CreateCaseFolders(root, L"baseInterfaces"))
            {
                const auto& folders = *foldersOpt;
                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                state.Require(static_cast<bool>(session->GetBaseFileSystem()), L"GetBaseFileSystem() should return non-null.");
                state.Require(static_cast<bool>(session->GetBaseInformations()), L"GetBaseInformations() should return non-null.");
                state.Require(static_cast<bool>(session->GetBaseFileSystemIO()), L"GetBaseFileSystemIO() should return non-null.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: baseInterfaces.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: Repeated GetOrComputeDecision without invalidation returns the same cached object.
            AppendCompareSelfTestTraceLine(L"Case: contentCacheHit");
            if (const auto foldersOpt = CreateCaseFolders(root, L"contentCacheHit"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "CacheA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "CacheA"), L"Failed to create a.txt (right).");

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const auto decision1 = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision1), L"First call should return a valid decision.");
                const auto decision2 = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision2), L"Second call should return a valid decision.");

                // Without any intervening Invalidate(), both calls must return the identical cached shared_ptr.
                state.Require(decision1 == decision2, L"Repeated GetOrComputeDecision without invalidation must return the same cached decision.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: contentCacheHit.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: compareContent=true on two zero-byte files reports them as identical.
            AppendCompareSelfTestTraceLine(L"Case: zeroByteContent");
            if (const auto foldersOpt = CreateCaseFolders(root, L"zeroByteContent"))
            {
                const auto& folders = *foldersOpt;
                // Create empty files on both sides.
                state.Require(SelfTest::WriteBinaryFile(folders.left / L"empty.txt", {}), L"Failed to create empty.txt (left).");
                state.Require(SelfTest::WriteBinaryFile(folders.right / L"empty.txt", {}), L"Failed to create empty.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"empty.txt");
                    state.Require(item != nullptr, L"empty.txt should appear in the decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"Zero-byte files on both sides must be identical.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"Zero-byte files must not have the Content diff bit set.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: zeroByteContent.");
            }

            if (shouldAbort())
            {
                break;
            }

            // Case: SetSettings with a meaningful change increments GetVersion(); a no-op toggle does not.
            AppendCompareSelfTestTraceLine(L"Case: setSettingsInvalidates");
            if (const auto foldersOpt = CreateCaseFolders(root, L"setSettingsInvalidates"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "V"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "V"), L"Failed to create a.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = false;
                auto session            = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);

                const uint64_t v0 = session->GetVersion();

                // Changing compareContent must invalidate the cache (version bump).
                settings.compareContent = true;
                session->SetSettings(settings);
                const uint64_t v1 = session->GetVersion();
                state.Require(v1 != v0, L"SetSettings with compareContent toggled must increment GetVersion().");

                // Setting the same value again must NOT bump the version.
                session->SetSettings(settings);
                const uint64_t v2 = session->GetVersion();
                state.Require(v2 == v1, L"SetSettings with identical settings must not increment GetVersion().");

                // Changing compareSize must also invalidate.
                settings.compareSize = ! settings.compareSize;
                session->SetSettings(settings);
                const uint64_t v3 = session->GetVersion();
                state.Require(v3 != v2, L"SetSettings with compareSize toggled must increment GetVersion().");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: setSettingsInvalidates.");
            }

        } while (false);
    }

    AppendCompareSelfTestTraceLine(L"Run: finalizing");

    const auto endedAt        = std::chrono::steady_clock::now();
    const uint64_t durationMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

    const bool hadNoCases = ! state.caseInProgress && state.caseResults.empty();

    SelfTest::SelfTestSuiteResult suiteResult = state.GetResult(durationMs);
    if (state.failed && suiteResult.failed == 0)
    {
        SelfTest::SelfTestCaseResult setupResult{};
        setupResult.name       = L"setup";
        setupResult.status     = SelfTest::SelfTestCaseResult::Status::failed;
        setupResult.durationMs = 0;
        setupResult.reason     = suiteResult.failureMessage;
        suiteResult.cases.insert(suiteResult.cases.begin(), std::move(setupResult));
        ++suiteResult.failed;
    }

    if (state.failed && hadNoCases && suiteResult.cases.size() == 1u)
    {
        for (const auto& name : kCompareCaseNames)
        {
            SelfTest::SelfTestCaseResult skipped{};
            skipped.name       = std::wstring(name);
            skipped.status     = SelfTest::SelfTestCaseResult::Status::skipped;
            skipped.durationMs = 0;
            skipped.reason     = L"not executed (suite setup failed)";
            suiteResult.cases.push_back(std::move(skipped));
            ++suiteResult.skipped;
        }
    }

    if (outResult)
    {
        *outResult = suiteResult;
    }

    if (options.writeJsonSummary)
    {
        const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::CompareDirectories, L"results.json");
        SelfTest::WriteSuiteJson(suiteResult, jsonPath);
    }

    if (state.failed)
    {
        AppendCompareSelfTestTraceLine(L"Run: failed");
        Debug::Error(L"CompareSelfTest: failed.");
        return false;
    }

    AppendCompareSelfTestTraceLine(L"Run: passed");
    Debug::Info(L"CompareSelfTest: passed.");
    return true;
}

#endif // _DEBUG

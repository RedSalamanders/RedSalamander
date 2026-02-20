#include <algorithm>
#include <array>
#include <charconv>
#include <chrono>
#include <cstring>
#include <format>
#include <limits>
#include <new>
#include <utility>

#include "FileSystemDummy.h"

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

namespace
{
constexpr size_t kEntryAlignment = sizeof(unsigned long);
constexpr size_t kMaxNameLength  = 96;

thread_local const void* g_activeDirectoryWatchCallback = nullptr;

struct DirectoryWatchCallbackScope final
{
    explicit DirectoryWatchCallbackScope(const void* watcher) noexcept : _previous(g_activeDirectoryWatchCallback)
    {
        g_activeDirectoryWatchCallback = watcher;
    }

    DirectoryWatchCallbackScope(const DirectoryWatchCallbackScope&)            = delete;
    DirectoryWatchCallbackScope& operator=(const DirectoryWatchCallbackScope&) = delete;
    DirectoryWatchCallbackScope(DirectoryWatchCallbackScope&&)                 = delete;
    DirectoryWatchCallbackScope& operator=(DirectoryWatchCallbackScope&&)      = delete;

    ~DirectoryWatchCallbackScope()
    {
        g_activeDirectoryWatchCallback = _previous;
    }

private:
    const void* _previous = nullptr;
};

enum class DummyFillKind : std::uint8_t
{
    PlainText,
    JsonString,
    XmlCData,
    CsvField,
    Binary,
};

std::uint64_t Mix64(std::uint64_t value) noexcept
{
    value += 0x9E3779B97F4A7C15ull;
    value = (value ^ (value >> 30u)) * 0xBF58476D1CE4E5B9ull;
    value = (value ^ (value >> 27u)) * 0x94D049BB133111EBull;
    return value ^ (value >> 31u);
}

std::uint8_t GenerateDummyByte(DummyFillKind kind, std::uint64_t seed, uint64_t position) noexcept
{
    if (kind == DummyFillKind::PlainText || kind == DummyFillKind::XmlCData)
    {
        if ((position % 97u) == 95u)
        {
            return '\r';
        }
        if ((position % 97u) == 96u)
        {
            return '\n';
        }
    }

    const std::uint64_t mixed = Mix64(seed + position);
    const std::uint8_t pick   = static_cast<std::uint8_t>(mixed & 0xFFu);

    if (kind == DummyFillKind::Binary)
    {
        return pick;
    }

    if (kind == DummyFillKind::JsonString || kind == DummyFillKind::CsvField || kind == DummyFillKind::XmlCData)
    {
        constexpr std::string_view kChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ ";
        const size_t index                = static_cast<size_t>(pick) % kChars.size();
        return static_cast<std::uint8_t>(kChars[index]);
    }

    if (pick < 20u)
    {
        return ' ';
    }
    if (pick < 22u)
    {
        return '.';
    }
    if (pick < 24u)
    {
        return ',';
    }
    if (pick < 26u)
    {
        return ';';
    }
    if (pick < 28u)
    {
        return ':';
    }
    if (pick < 30u)
    {
        return '!';
    }
    if (pick < 32u)
    {
        return '?';
    }

    return static_cast<std::uint8_t>('a' + (pick % 26u));
}

class DummyGeneratedFileReader final : public IFileReader
{
public:
    DummyGeneratedFileReader(std::string prefix, std::string suffix, uint64_t bodyBytes, std::uint64_t seed, DummyFillKind fillKind) noexcept
        : _prefix(std::move(prefix)),
          _suffix(std::move(suffix)),
          _bodyBytes(bodyBytes),
          _seed(seed),
          _fillKind(fillKind)
    {
    }

    DummyGeneratedFileReader(const DummyGeneratedFileReader&)            = delete;
    DummyGeneratedFileReader(DummyGeneratedFileReader&&)                 = delete;
    DummyGeneratedFileReader& operator=(const DummyGeneratedFileReader&) = delete;
    DummyGeneratedFileReader& operator=(DummyGeneratedFileReader&&)      = delete;

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
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = GetTotalSizeBytes();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        __int64 base = 0;
        if (origin == FILE_CURRENT)
        {
            base = static_cast<__int64>(_positionBytes);
        }
        else if (origin == FILE_END)
        {
            base = static_cast<__int64>(GetTotalSizeBytes());
        }

        const __int64 next = base + offset;
        if (next < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        _positionBytes = static_cast<uint64_t>(next);
        *newPosition   = _positionBytes;
        return S_OK;
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

        const uint64_t totalSize = GetTotalSizeBytes();
        if (_positionBytes >= totalSize)
        {
            return S_OK;
        }

        const uint64_t remaining = totalSize - _positionBytes;
        const unsigned long take = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

        auto* out = static_cast<std::uint8_t*>(buffer);

        const uint64_t prefixBytes = static_cast<uint64_t>(_prefix.size());
        const uint64_t suffixBytes = static_cast<uint64_t>(_suffix.size());

        unsigned long written = 0;
        while (written < take)
        {
            const uint64_t absolutePos = _positionBytes + static_cast<uint64_t>(written);

            if (absolutePos < prefixBytes)
            {
                const size_t offset    = static_cast<size_t>(absolutePos);
                const size_t available = _prefix.size() - offset;
                const size_t want      = std::min<size_t>(available, static_cast<size_t>(take - written));
                memcpy(out + written, _prefix.data() + offset, want);
                written += static_cast<unsigned long>(want);
                continue;
            }

            const uint64_t bodyStart = prefixBytes;
            const uint64_t bodyEnd   = prefixBytes + _bodyBytes;
            if (absolutePos < bodyEnd)
            {
                const uint64_t bodyPos = absolutePos - bodyStart;
                out[written]           = GenerateDummyByte(_fillKind, _seed, bodyPos);
                ++written;
                continue;
            }

            if (suffixBytes == 0)
            {
                break;
            }

            const uint64_t suffixPos = absolutePos - bodyEnd;
            if (suffixPos >= suffixBytes)
            {
                break;
            }

            const size_t offset    = static_cast<size_t>(suffixPos);
            const size_t available = _suffix.size() - offset;
            const size_t want      = std::min<size_t>(available, static_cast<size_t>(take - written));
            memcpy(out + written, _suffix.data() + offset, want);
            written += static_cast<unsigned long>(want);
        }

        _positionBytes += static_cast<uint64_t>(take);
        *bytesRead = take;
        return S_OK;
    }

private:
    ~DummyGeneratedFileReader() = default;

    uint64_t GetTotalSizeBytes() const noexcept
    {
        const uint64_t prefixBytes = static_cast<uint64_t>(_prefix.size());
        const uint64_t suffixBytes = static_cast<uint64_t>(_suffix.size());
        return prefixBytes + _bodyBytes + suffixBytes;
    }

    std::atomic_ulong _refCount{1};
    std::string _prefix;
    std::string _suffix;
    uint64_t _bodyBytes     = 0;
    std::uint64_t _seed     = 0;
    DummyFillKind _fillKind = DummyFillKind::PlainText;
    uint64_t _positionBytes = 0;
};

class DummyBufferFileReader final : public IFileReader
{
public:
    explicit DummyBufferFileReader(std::vector<std::byte> buffer) noexcept : _buffer(std::move(buffer))
    {
    }

    DummyBufferFileReader(const DummyBufferFileReader&)            = delete;
    DummyBufferFileReader(DummyBufferFileReader&&)                 = delete;
    DummyBufferFileReader& operator=(const DummyBufferFileReader&) = delete;
    DummyBufferFileReader& operator=(DummyBufferFileReader&&)      = delete;

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
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = static_cast<uint64_t>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        __int64 base = 0;
        if (origin == FILE_CURRENT)
        {
            base = static_cast<__int64>(_positionBytes);
        }
        else if (origin == FILE_END)
        {
            base = static_cast<__int64>(_buffer.size());
        }

        const __int64 next = base + offset;
        if (next < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        _positionBytes = static_cast<uint64_t>(next);
        *newPosition   = _positionBytes;
        return S_OK;
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

        const uint64_t totalSize = static_cast<uint64_t>(_buffer.size());
        if (_positionBytes >= totalSize)
        {
            return S_OK;
        }

        const uint64_t remaining = totalSize - _positionBytes;
        const unsigned long take = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

        memcpy(buffer, _buffer.data() + static_cast<size_t>(_positionBytes), take);
        _positionBytes += static_cast<uint64_t>(take);
        *bytesRead = take;
        return S_OK;
    }

private:
    ~DummyBufferFileReader() = default;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    uint64_t _positionBytes = 0;
};

class DummySharedBufferFileReader final : public IFileReader
{
public:
    explicit DummySharedBufferFileReader(std::shared_ptr<std::vector<std::byte>> buffer) noexcept : _buffer(std::move(buffer))
    {
    }

    DummySharedBufferFileReader(const DummySharedBufferFileReader&)            = delete;
    DummySharedBufferFileReader(DummySharedBufferFileReader&&)                 = delete;
    DummySharedBufferFileReader& operator=(const DummySharedBufferFileReader&) = delete;
    DummySharedBufferFileReader& operator=(DummySharedBufferFileReader&&)      = delete;

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
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = 0;
        if (! _buffer)
        {
            return E_FAIL;
        }

        *sizeBytes = static_cast<uint64_t>(_buffer->size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        if (! _buffer)
        {
            return E_FAIL;
        }

        __int64 base = 0;
        if (origin == FILE_CURRENT)
        {
            base = static_cast<__int64>(_positionBytes);
        }
        else if (origin == FILE_END)
        {
            base = static_cast<__int64>(_buffer->size());
        }

        const __int64 next = base + offset;
        if (next < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        _positionBytes = static_cast<uint64_t>(next);
        *newPosition   = _positionBytes;
        return S_OK;
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

        if (! _buffer)
        {
            return E_FAIL;
        }

        const uint64_t totalSize = static_cast<uint64_t>(_buffer->size());
        if (_positionBytes >= totalSize)
        {
            return S_OK;
        }

        const uint64_t remaining = totalSize - _positionBytes;
        const unsigned long take = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

        memcpy(buffer, _buffer->data() + static_cast<size_t>(_positionBytes), take);
        _positionBytes += static_cast<uint64_t>(take);
        *bytesRead = take;
        return S_OK;
    }

private:
    ~DummySharedBufferFileReader() = default;

    std::atomic_ulong _refCount{1};
    std::shared_ptr<std::vector<std::byte>> _buffer;
    uint64_t _positionBytes = 0;
};

class DummyFileWriter final : public IFileWriter
{
public:
    DummyFileWriter(FileSystemDummy& owner, std::filesystem::path normalizedPath, FileSystemFlags flags) noexcept
        : _refCount(1),
          _owner(&owner),
          _path(std::move(normalizedPath)),
          _flags(flags)
    {
        _owner->AddRef();
    }

    DummyFileWriter(const DummyFileWriter&)            = delete;
    DummyFileWriter(DummyFileWriter&&)                 = delete;
    DummyFileWriter& operator=(const DummyFileWriter&) = delete;
    DummyFileWriter& operator=(DummyFileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
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

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        if (positionBytes == nullptr)
        {
            return E_POINTER;
        }

        const std::shared_ptr<std::vector<std::byte>> buffer = EnsureBuffer();
        if (! buffer)
        {
            *positionBytes = 0;
            return E_OUTOFMEMORY;
        }

        *positionBytes = static_cast<uint64_t>(buffer->size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (bytesWritten == nullptr)
        {
            return E_POINTER;
        }

        *bytesWritten = 0;

        if (bytesToWrite == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        const std::shared_ptr<std::vector<std::byte>> out = EnsureBuffer();
        if (! out)
        {
            return E_OUTOFMEMORY;
        }

        const size_t oldSize = out->size();
        const size_t add     = static_cast<size_t>(bytesToWrite);
        if (oldSize > (std::numeric_limits<size_t>::max)() - add)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        out->resize(oldSize + add);

        memcpy(out->data() + oldSize, buffer, bytesToWrite);
        *bytesWritten = bytesToWrite;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return S_OK;
        }

        if (! _owner)
        {
            return E_FAIL;
        }

        const std::shared_ptr<std::vector<std::byte>> buffer = EnsureBuffer();
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = _owner->CommitFileWriter(_path, _flags, buffer);
        if (FAILED(hr))
        {
            return hr;
        }

        _committed = true;
        return S_OK;
    }

private:
    ~DummyFileWriter()
    {
        if (_owner)
        {
            _owner->Release();
            _owner = nullptr;
        }
    }

    std::shared_ptr<std::vector<std::byte>> EnsureBuffer() noexcept
    {
        if (_buffer)
        {
            return _buffer;
        }

        _buffer = std::make_shared<std::vector<std::byte>>();
        return _buffer;
    }

    std::atomic_ulong _refCount{1};
    FileSystemDummy* _owner = nullptr;
    std::filesystem::path _path;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    bool _committed        = false;
    std::shared_ptr<std::vector<std::byte>> _buffer;
};

std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

template <typename T> constexpr T AlignUp(T value, size_t alignment) noexcept
{
    const size_t mask = alignment - 1u;
    return static_cast<T>((static_cast<size_t>(value) + mask) & ~mask);
}

template <typename T, size_t N> constexpr unsigned long ArrayCount(const T (&)[N]) noexcept
{
    return static_cast<unsigned long>(N);
}

struct DummyEntry
{
    std::wstring name;
    DWORD attributes       = 0;
    uint64_t sizeBytes     = 0;
    __int64 creationTime   = 0;
    __int64 lastAccessTime = 0;
    __int64 lastWriteTime  = 0;
    __int64 changeTime     = 0;
};

constexpr std::wstring_view kWordSegments[] = {L"alpha",   L"bravo", L"charlie", L"delta",  L"echo",     L"foxtrot", L"golf",   L"hotel",
                                               L"juliet",  L"kilo",  L"lima",    L"mango",  L"notebook", L"archive", L"report", L"session",
                                               L"palette", L"theme", L"vector",  L"module", L"sample",   L"draft",   L"output", L"project"};

constexpr std::wstring_view kEuroSegments[] = {L"caf\u00E9",
                                               L"fran\u00E7ais",
                                               L"ni\u00F1o",
                                               L"m\u00FCnchen",
                                               L"gar\u00E7on",
                                               L"fa\u00E7ade",
                                               L"sm\u00F8rrebr\u00F8d",
                                               L"\u0141\u00F3d\u017A",
                                               L"S\u00F8rensen",
                                               L"\u00FCber",
                                               L"\u00E5ngstr\u00F6m",
                                               L"canci\u00F3n",
                                               L"\u015Ar\u00F3da",
                                               L"pi\u00F1ata"};

constexpr std::wstring_view kJapaneseSegments[] = {
    L"日本語", L"東京", L"さくら", L"ファイル", L"テスト", L"プロジェクト", L"設定", L"履歴", L"サンプル", L"レポート", L"ドキュメント", L"フォルダー"};

constexpr std::wstring_view kArabicSegments[] = {L"مرحبا", L"ملف", L"اختبار", L"مشروع", L"تقرير", L"مجلد", L"إعدادات", L"مستند"};

constexpr std::wstring_view kThaiSegments[] = {L"สวัสดี", L"ไฟล์", L"ทดสอบ", L"โครงการ", L"รายงาน", L"โฟลเดอร์", L"การตั้งค่า", L"เอกสาร"};

constexpr std::wstring_view kKoreanSegments[] = {L"한국어", L"안녕하세요", L"파일", L"테스트", L"프로젝트", L"보고서", L"설정", L"문서"};

constexpr std::wstring_view kEmojiSegments[] = {
    L"\U0001F600", L"\U0001F680", L"\U0001F389", L"\U0001F31F", L"\U0001F525", L"\U0001F4C4", L"\U0001F4DA", L"\U0001F4BB", L"\U0001F984", L"\U0001F9EA"};

constexpr std::wstring_view kLongSegments[] = {
    L"supercalifragilisticexpialidocious", L"pseudopseudohypoparathyroidism", L"electroencephalograph", L"characterization", L"internationalization"};

constexpr std::wstring_view kExtensions[] = {
    L".txt", L".log", L".json", L".json5", L".xml", L".theme.json5", L".png", L".jpg", L".bin", L".cpp", L".h", L".md", L".csv", L".zip", L".docx", L".xlsx"};

enum class DummyFileKind
{
    Text,
    Csv,
    Json,
    Json5,
    ThemeJson5,
    Xml,
    Png,
    Jpeg,
    Zip,
    Binary,
};

DummyFileKind GetDummyFileKind(std::wstring_view fileName) noexcept;
uint64_t MakeDummyFileSize(std::mt19937& rng, DummyFileKind kind) noexcept;

constexpr wchar_t kSeparators[] = {L' ', L'-', L'_'};

constexpr std::uint64_t SplitMix64(std::uint64_t value) noexcept
{
    value += 0x9E3779B97F4A7C15ull;
    value = (value ^ (value >> 30u)) * 0xBF58476D1CE4E5B9ull;
    value = (value ^ (value >> 27u)) * 0x94D049BB133111EBull;
    return value ^ (value >> 31u);
}

constexpr std::uint64_t kFnvOffsetBasis = 14695981039346656037ull;
constexpr std::uint64_t kFnvPrime       = 1099511628211ull;

std::uint64_t HashAppendU64(std::uint64_t hash, std::uint64_t value) noexcept
{
    for (unsigned int shift = 0; shift < 64u; shift += 8u)
    {
        hash ^= static_cast<std::uint8_t>((value >> shift) & 0xFFu);
        hash *= kFnvPrime;
    }
    return hash;
}

std::uint64_t HashAppendWideString(std::uint64_t hash, std::wstring_view text) noexcept
{
    for (const wchar_t ch : text)
    {
        const std::uint16_t codeUnit = static_cast<std::uint16_t>(ch);
        hash ^= static_cast<std::uint8_t>(codeUnit & 0xFFu);
        hash *= kFnvPrime;
        hash ^= static_cast<std::uint8_t>((codeUnit >> 8u) & 0xFFu);
        hash *= kFnvPrime;
    }
    return hash;
}

std::uint64_t CombineSeed(std::uint64_t baseSeed, std::wstring_view salt) noexcept
{
    std::uint64_t hash = kFnvOffsetBasis;
    hash               = HashAppendU64(hash, baseSeed);
    hash               = HashAppendWideString(hash, salt);
    return SplitMix64(hash);
}

std::uint64_t CombineSeed(std::uint64_t baseSeed, std::uint64_t salt) noexcept
{
    std::uint64_t hash = kFnvOffsetBasis;
    hash               = HashAppendU64(hash, baseSeed);
    hash               = HashAppendU64(hash, salt);
    return SplitMix64(hash);
}

std::mt19937 MakeRng(std::uint64_t seed) noexcept
{
    const unsigned int seedLow  = static_cast<unsigned int>(seed);
    const unsigned int seedHigh = static_cast<unsigned int>(seed >> 32u);
    std::seed_seq seq{seedLow, seedHigh};
    return std::mt19937(seq);
}

std::uint64_t DeriveChildSeed(std::uint64_t parentSeed, unsigned long childIndex, bool isDirectory) noexcept
{
    const std::uint64_t salt = (static_cast<std::uint64_t>(childIndex) << 1u) | (isDirectory ? 1ull : 0ull);
    return CombineSeed(parentSeed, salt);
}

std::uint64_t ComputeGenerationBaseTime(std::uint64_t seed) noexcept
{
    constexpr std::uint64_t kJan1_2024_FileTime = 133485408000000000ull;
    constexpr std::uint64_t kMaxOffsetSeconds   = 60ull * 60ull * 24ull * 90ull;
    const std::uint64_t offsetSeconds           = SplitMix64(seed) % (kMaxOffsetSeconds + 1ull);
    return kJan1_2024_FileTime + (offsetSeconds * 10000000ull);
}

unsigned long RandomRange(std::mt19937& rng, unsigned long minValue, unsigned long maxValue) noexcept
{
    if (minValue >= maxValue)
    {
        return minValue;
    }

    std::uniform_int_distribution<unsigned long> dist(minValue, maxValue);
    return dist(rng);
}

unsigned long RandomSkewedUpTo(std::mt19937& rng, unsigned long maxValue) noexcept
{
    if (maxValue == 0)
    {
        return 0;
    }

    const unsigned long roll   = RandomRange(rng, 0, maxValue);
    const std::uint64_t roll64 = static_cast<std::uint64_t>(roll);
    const std::uint64_t max64  = static_cast<std::uint64_t>(maxValue);

    const std::uint64_t numerator   = roll64 * roll64 * roll64 * roll64;
    const std::uint64_t denominator = max64 * max64 * max64;
    if (denominator == 0)
    {
        return 0;
    }

    const std::uint64_t scaled = numerator / denominator;
    if (scaled >= max64)
    {
        return maxValue;
    }

    return static_cast<unsigned long>(scaled);
}

uint64_t RandomRange64(std::mt19937& rng, uint64_t minValue, uint64_t maxValue) noexcept
{
    if (minValue >= maxValue)
    {
        return minValue;
    }

    std::uniform_int_distribution<uint64_t> dist(minValue, maxValue);
    return dist(rng);
}

uint64_t RandomSkewedUpTo64(std::mt19937& rng, uint64_t maxValue) noexcept
{
    if (maxValue == 0)
    {
        return 0;
    }

    const std::uint32_t roll32 = rng();
    std::uint64_t value        = static_cast<std::uint64_t>(roll32);
    value                      = (value * value) >> 32u;
    value                      = (value * value) >> 32u;

    const std::uint64_t max64  = static_cast<std::uint64_t>(maxValue);
    const std::uint64_t scaled = (value * (max64 + 1ull)) >> 32u;
    if (scaled >= max64)
    {
        return maxValue;
    }

    return static_cast<uint64_t>(scaled);
}

bool RandomChance(std::mt19937& rng, unsigned long numerator, unsigned long denominator) noexcept
{
    if (denominator == 0)
    {
        return false;
    }

    const unsigned long roll = RandomRange(rng, 1, denominator);
    return roll <= numerator;
}

bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    if (left.empty())
    {
        return true;
    }

    if (left.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int length = static_cast<int>(left.size());
    return CompareStringOrdinal(left.data(), length, right.data(), length, TRUE) == CSTR_EQUAL;
}

bool EndsWithNoCase(std::wstring_view text, std::wstring_view suffix) noexcept
{
    if (text.size() < suffix.size())
    {
        return false;
    }

    return EqualsNoCase(text.substr(text.size() - suffix.size()), suffix);
}

DummyFileKind GetDummyFileKind(std::wstring_view fileName) noexcept
{
    if (EndsWithNoCase(fileName, L".theme.json5"))
    {
        return DummyFileKind::ThemeJson5;
    }

    if (EndsWithNoCase(fileName, L".json5"))
    {
        return DummyFileKind::Json5;
    }

    if (EndsWithNoCase(fileName, L".json"))
    {
        return DummyFileKind::Json;
    }

    if (EndsWithNoCase(fileName, L".xml"))
    {
        return DummyFileKind::Xml;
    }

    if (EndsWithNoCase(fileName, L".csv"))
    {
        return DummyFileKind::Csv;
    }

    if (EndsWithNoCase(fileName, L".png"))
    {
        return DummyFileKind::Png;
    }

    if (EndsWithNoCase(fileName, L".jpg") || EndsWithNoCase(fileName, L".jpeg"))
    {
        return DummyFileKind::Jpeg;
    }

    if (EndsWithNoCase(fileName, L".zip") || EndsWithNoCase(fileName, L".docx") || EndsWithNoCase(fileName, L".xlsx"))
    {
        return DummyFileKind::Zip;
    }

    if (EndsWithNoCase(fileName, L".bin"))
    {
        return DummyFileKind::Binary;
    }

    return DummyFileKind::Text;
}

uint64_t MakeDummyFileSize(std::mt19937& rng, DummyFileKind kind) noexcept
{
    constexpr uint64_t kMaxGenericBytes = 25ull * 1024ull * 1024ull;

    if (kind == DummyFileKind::Png)
    {
        return std::max<uint64_t>(RandomRange64(rng, 4ull * 1024ull, 512ull * 1024ull), 256ull);
    }

    if (kind == DummyFileKind::Jpeg)
    {
        return std::max<uint64_t>(RandomRange64(rng, 2ull * 1024ull, 512ull * 1024ull), 256ull);
    }

    if (kind == DummyFileKind::Zip)
    {
        return std::max<uint64_t>(RandomRange64(rng, 128ull, 256ull * 1024ull), 22ull);
    }

    if (kind == DummyFileKind::Csv || kind == DummyFileKind::Json || kind == DummyFileKind::Json5 || kind == DummyFileKind::ThemeJson5 ||
        kind == DummyFileKind::Xml)
    {
        constexpr uint64_t kMaxStructuredBytes = 2ull * 1024ull * 1024ull;
        const uint64_t sizeBytes               = RandomSkewedUpTo64(rng, kMaxStructuredBytes);
        return std::max<uint64_t>(sizeBytes, 128ull);
    }

    if (kind == DummyFileKind::Binary)
    {
        return RandomSkewedUpTo64(rng, kMaxGenericBytes);
    }

    return RandomSkewedUpTo64(rng, kMaxGenericBytes);
}

bool IsAsciiSpace(char ch) noexcept
{
    return ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' || ch == '\f' || ch == '\v';
}

std::string_view TrimAscii(std::string_view text) noexcept
{
    while (! text.empty() && IsAsciiSpace(text.front()))
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && IsAsciiSpace(text.back()))
    {
        text.remove_suffix(1);
    }
    return text;
}

char FoldAsciiCase(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

bool EqualsIgnoreAsciiCase(std::string_view a, std::string_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (FoldAsciiCase(a[i]) != FoldAsciiCase(b[i]))
        {
            return false;
        }
    }

    return true;
}

uint64_t MultiplyOrSaturate(uint64_t a, uint64_t b) noexcept
{
    if (a == 0 || b == 0)
    {
        return 0;
    }

    if (a > (std::numeric_limits<uint64_t>::max() / b))
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return a * b;
}

bool TryParseThroughputText(std::string_view text, uint64_t& outBytesPerSecond) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

    outBytesPerSecond = 0;

    text = TrimAscii(text);
    if (text.empty())
    {
        return true;
    }

    uint64_t number   = 0;
    const char* begin = text.data();
    const char* end   = begin + text.size();

    const auto [ptr, ec] = std::from_chars(begin, end, number);
    if (ec != std::errc{})
    {
        return false;
    }

    std::string_view unit(ptr, static_cast<size_t>(end - ptr));
    unit = TrimAscii(unit);

    if (unit.size() >= 2)
    {
        const char penultimate = unit[unit.size() - 2];
        const char last        = unit.back();
        if (penultimate == '/' && (last == 's' || last == 'S'))
        {
            unit.remove_suffix(2);
            unit = TrimAscii(unit);
        }
    }

    uint64_t multiplier = 0;
    if (unit.empty() || EqualsIgnoreAsciiCase(unit, "kb") || EqualsIgnoreAsciiCase(unit, "k") || EqualsIgnoreAsciiCase(unit, "kib"))
    {
        // Bare numeric strings are interpreted as KiB for user-friendliness.
        multiplier = kKiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, "b"))
    {
        multiplier = 1;
    }
    else if (EqualsIgnoreAsciiCase(unit, "mb") || EqualsIgnoreAsciiCase(unit, "m") || EqualsIgnoreAsciiCase(unit, "mib"))
    {
        multiplier = kMiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, "gb") || EqualsIgnoreAsciiCase(unit, "g") || EqualsIgnoreAsciiCase(unit, "gib"))
    {
        multiplier = kGiB;
    }
    else
    {
        return false;
    }

    outBytesPerSecond = MultiplyOrSaturate(number, multiplier);
    return true;
}

std::wstring EscapeJsonString(std::wstring_view input)
{
    if (input.empty())
    {
        return {};
    }

    std::wstring output;
    output.reserve(input.size());

    for (const wchar_t ch : input)
    {
        switch (ch)
        {
            case L'\\': output.append(L"\\\\"); break;
            case L'"': output.append(L"\\\""); break;
            case L'\b': output.append(L"\\b"); break;
            case L'\f': output.append(L"\\f"); break;
            case L'\n': output.append(L"\\n"); break;
            case L'\r': output.append(L"\\r"); break;
            case L'\t': output.append(L"\\t"); break;
            default:
            {
                if (ch < 0x20)
                {
                    output.append(std::format(L"\\u{:04X}", static_cast<unsigned int>(ch)));
                }
                else
                {
                    output.push_back(ch);
                }
                break;
            }
        }
    }

    return output;
}

struct DummyFileSnapshot
{
    std::wstring name;
    unsigned long attributes     = 0;
    uint64_t sizeBytes           = 0;
    __int64 creationTime         = 0;
    std::uint64_t generationSeed = 0;
    std::shared_ptr<std::vector<std::byte>> materializedContent;
};

std::uint64_t ComputeDummyFileContentSeed(const DummyFileSnapshot& snapshot) noexcept
{
    std::uint64_t hash = kFnvOffsetBasis;
    hash               = HashAppendU64(hash, snapshot.generationSeed);
    hash               = HashAppendWideString(hash, snapshot.name);
    hash               = HashAppendU64(hash, static_cast<std::uint64_t>(snapshot.sizeBytes));
    hash               = HashAppendU64(hash, static_cast<std::uint64_t>(snapshot.creationTime));
    hash               = HashAppendU64(hash, static_cast<std::uint64_t>(snapshot.attributes));
    return SplitMix64(hash);
}

std::string XmlEscapeAttributeUtf8(std::wstring_view text) noexcept
{
    std::string utf8 = Utf8FromUtf16(text);
    if (utf8.empty())
    {
        return {};
    }

    std::string output;
    output.reserve(utf8.size());

    for (const char ch : utf8)
    {
        switch (ch)
        {
            case '&': output.append("&amp;"); break;
            case '<': output.append("&lt;"); break;
            case '>': output.append("&gt;"); break;
            case '"': output.append("&quot;"); break;
            case '\'': output.append("&apos;"); break;
            default: output.push_back(ch); break;
        }
    }

    return output;
}

struct DummyTextTemplate
{
    std::string prefix;
    std::string suffix;
    uint64_t bodyBytes     = 0;
    DummyFillKind fillKind = DummyFillKind::PlainText;
};

DummyTextTemplate BuildDummyTextTemplate(DummyFileKind kind, const DummyFileSnapshot& snapshot, std::uint64_t contentSeed)
{
    DummyTextTemplate result{};

    const std::string nameUtf8 = Utf8FromUtf16(snapshot.name);
    const auto seedValue       = static_cast<unsigned long long>(contentSeed);
    const auto fileSizeValue   = static_cast<unsigned long long>(snapshot.sizeBytes);
    const auto createdValue    = static_cast<long long>(snapshot.creationTime);

    if (kind == DummyFileKind::Csv)
    {
        result.fillKind = DummyFillKind::CsvField;
        result.prefix   = std::format("id,name,sizeBytes,created,seed,data\r\n0,\"{}\",{},{},{:016X},\"", nameUtf8, fileSizeValue, createdValue, seedValue);
        result.suffix   = "\"\r\n";
    }
    else if (kind == DummyFileKind::Json)
    {
        result.fillKind                   = DummyFillKind::JsonString;
        const std::string escapedNameUtf8 = Utf8FromUtf16(EscapeJsonString(snapshot.name));
        result.prefix = std::format("{{\r\n  \"name\": \"{}\",\r\n  \"sizeBytes\": {},\r\n  \"created\": {},\r\n  \"seed\": \"{:016X}\",\r\n  \"data\": \"",
                                    escapedNameUtf8,
                                    fileSizeValue,
                                    createdValue,
                                    seedValue);
        result.suffix = "\"\r\n}\r\n";
    }
    else if (kind == DummyFileKind::Json5)
    {
        result.fillKind                   = DummyFillKind::JsonString;
        const std::string escapedNameUtf8 = Utf8FromUtf16(EscapeJsonString(snapshot.name));
        result.prefix                     = std::format(
            "// FileSystemDummy generated (JSON5)\r\n{{\r\n  name: \"{}\",\r\n  sizeBytes: {},\r\n  created: {},\r\n  seed: \"{:016X}\",\r\n  data: \"",
            escapedNameUtf8,
            fileSizeValue,
            createdValue,
            seedValue);
        result.suffix = "\"\r\n}\r\n";
    }
    else if (kind == DummyFileKind::ThemeJson5)
    {
        result.fillKind = DummyFillKind::JsonString;

        const std::string escapedNameUtf8 = Utf8FromUtf16(EscapeJsonString(snapshot.name));
        const unsigned int accentRgb      = static_cast<unsigned int>(contentSeed & 0xFFFFFFu);
        const unsigned int backgroundRgb  = static_cast<unsigned int>((contentSeed >> 24u) & 0xFFFFFFu);

        result.prefix = std::format(
            "// FileSystemDummy generated theme (JSON5)\r\n{{\r\n  id: \"user/dummy-{:016X}\",\r\n  name: \"{}\",\r\n  baseThemeId: \"builtin/dark\",\r\n  "
            "colors: {{\r\n    \"app.accent\": \"#{:06X}\",\r\n    \"window.background\": \"#{:06X}\",\r\n  }},\r\n  seed: \"{:016X}\",\r\n  data: \"",
            seedValue,
            escapedNameUtf8,
            accentRgb,
            backgroundRgb,
            seedValue);
        result.suffix = "\"\r\n}\r\n";
    }
    else if (kind == DummyFileKind::Xml)
    {
        result.fillKind                   = DummyFillKind::XmlCData;
        const std::string escapedNameUtf8 = XmlEscapeAttributeUtf8(snapshot.name);
        result.prefix                     = std::format(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<file name=\"{}\" sizeBytes=\"{}\" created=\"{}\" seed=\"{:016X}\">\r\n  <data><![CDATA[",
            escapedNameUtf8,
            fileSizeValue,
            createdValue,
            seedValue);
        result.suffix = "]]></data>\r\n</file>\r\n";
    }
    else
    {
        result.fillKind = DummyFillKind::PlainText;
        result.prefix   = std::format("FileSystemDummy generated file\r\nName: {}\r\nSizeBytes: {}\r\nSeed: {:016X}\r\nCreated: {}\r\n\r\n",
                                    nameUtf8,
                                    fileSizeValue,
                                    seedValue,
                                    createdValue);
        result.suffix   = "\r\n";
    }

    const uint64_t prefixBytes = static_cast<uint64_t>(result.prefix.size());
    const uint64_t suffixBytes = static_cast<uint64_t>(result.suffix.size());
    const uint64_t overhead    = prefixBytes + suffixBytes;

    if (snapshot.sizeBytes >= overhead)
    {
        result.bodyBytes = snapshot.sizeBytes - overhead;
        return result;
    }

    std::string combined = result.prefix;
    combined.append(result.suffix);
    if (combined.size() > static_cast<size_t>(snapshot.sizeBytes))
    {
        combined.resize(static_cast<size_t>(snapshot.sizeBytes));
    }

    result.prefix    = std::move(combined);
    result.suffix    = {};
    result.bodyBytes = 0;
    return result;
}

std::array<std::uint32_t, 256> BuildCrc32Table() noexcept
{
    std::array<std::uint32_t, 256> table{};
    for (std::uint32_t index = 0; index < table.size(); ++index)
    {
        std::uint32_t crc = index;
        for (unsigned int bit = 0; bit < 8u; ++bit)
        {
            if ((crc & 1u) != 0u)
            {
                crc = 0xEDB88320u ^ (crc >> 1u);
            }
            else
            {
                crc >>= 1u;
            }
        }
        table[index] = crc;
    }
    return table;
}

const std::array<std::uint32_t, 256>& GetCrc32Table() noexcept
{
    static const std::array<std::uint32_t, 256> table = BuildCrc32Table();
    return table;
}

std::uint32_t Crc32Update(std::uint32_t crc, const std::uint8_t* data, size_t length) noexcept
{
    const auto& table     = GetCrc32Table();
    std::uint32_t current = crc;

    for (size_t index = 0; index < length; ++index)
    {
        current = table[(current ^ data[index]) & 0xFFu] ^ (current >> 8u);
    }

    return current;
}

std::uint32_t Crc32Chunk(const std::array<std::uint8_t, 4>& type, const std::uint8_t* data, size_t length) noexcept
{
    std::uint32_t crc = 0xFFFFFFFFu;
    crc               = Crc32Update(crc, type.data(), type.size());
    if (data && length > 0)
    {
        crc = Crc32Update(crc, data, length);
    }
    return crc ^ 0xFFFFFFFFu;
}

void AppendU32BE(std::vector<std::byte>& out, std::uint32_t value)
{
    out.push_back(static_cast<std::byte>((value >> 24u) & 0xFFu));
    out.push_back(static_cast<std::byte>((value >> 16u) & 0xFFu));
    out.push_back(static_cast<std::byte>((value >> 8u) & 0xFFu));
    out.push_back(static_cast<std::byte>(value & 0xFFu));
}

void AppendU16BE(std::vector<std::byte>& out, std::uint16_t value)
{
    out.push_back(static_cast<std::byte>((value >> 8u) & 0xFFu));
    out.push_back(static_cast<std::byte>(value & 0xFFu));
}

void AppendBytes(std::vector<std::byte>& out, const std::uint8_t* data, size_t length)
{
    if (! data || length == 0)
    {
        return;
    }

    const std::byte* start = reinterpret_cast<const std::byte*>(data);
    out.insert(out.end(), start, start + length);
}

void AppendPngChunk(std::vector<std::byte>& out, const std::array<std::uint8_t, 4>& type, const std::uint8_t* data, size_t length)
{
    AppendU32BE(out, static_cast<std::uint32_t>(length));
    AppendBytes(out, type.data(), type.size());
    AppendBytes(out, data, length);

    const std::uint32_t crc = Crc32Chunk(type, data, length);
    AppendU32BE(out, crc);
}

std::vector<std::byte> GenerateDummyPng(std::uint64_t seed, uint64_t targetBytes)
{
    constexpr std::uint32_t width  = 32;
    constexpr std::uint32_t height = 32;
    constexpr uint64_t baseBytes   = 3172;

    if (targetBytes < baseBytes)
    {
        return {};
    }

    if (targetBytes > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        return {};
    }

    constexpr std::uint32_t kAdlerMod = 65521u;
    std::uint32_t adlerA              = 1u;
    std::uint32_t adlerB              = 0u;

    auto updateAdler = [&](std::uint8_t byte)
    {
        adlerA += byte;
        if (adlerA >= kAdlerMod)
        {
            adlerA -= kAdlerMod;
        }
        adlerB += adlerA;
        adlerB %= kAdlerMod;
    };

    const size_t rawBytesPerRow = 1u + (static_cast<size_t>(width) * 3u);
    const size_t rawBytes       = rawBytesPerRow * static_cast<size_t>(height);

    std::vector<std::uint8_t> raw;
    raw.reserve(rawBytes);

    for (std::uint32_t y = 0; y < height; ++y)
    {
        raw.push_back(0);
        updateAdler(0);

        for (std::uint32_t x = 0; x < width; ++x)
        {
            const std::uint64_t v = Mix64(seed + (static_cast<std::uint64_t>(y) << 32u) + static_cast<std::uint64_t>(x));
            const std::uint8_t r  = static_cast<std::uint8_t>(v & 0xFFu);
            const std::uint8_t g  = static_cast<std::uint8_t>((v >> 8u) & 0xFFu);
            const std::uint8_t b  = static_cast<std::uint8_t>((v >> 16u) & 0xFFu);

            raw.push_back(r);
            raw.push_back(g);
            raw.push_back(b);

            updateAdler(r);
            updateAdler(g);
            updateAdler(b);
        }
    }

    if (raw.size() > static_cast<size_t>(std::numeric_limits<std::uint16_t>::max()))
    {
        return {};
    }

    const std::uint16_t rawLen = static_cast<std::uint16_t>(raw.size());
    const std::uint16_t nLen   = static_cast<std::uint16_t>(~rawLen);
    const std::uint32_t adler  = (adlerB << 16u) | adlerA;

    std::vector<std::uint8_t> zlib;
    zlib.reserve(2u + 5u + raw.size() + 4u);
    zlib.push_back(0x78u);
    zlib.push_back(0x01u);
    zlib.push_back(0x01u); // BFINAL=1, BTYPE=00 (stored)
    zlib.push_back(static_cast<std::uint8_t>(rawLen & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>((rawLen >> 8u) & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>(nLen & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>((nLen >> 8u) & 0xFFu));
    zlib.insert(zlib.end(), raw.begin(), raw.end());
    zlib.push_back(static_cast<std::uint8_t>((adler >> 24u) & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>((adler >> 16u) & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>((adler >> 8u) & 0xFFu));
    zlib.push_back(static_cast<std::uint8_t>(adler & 0xFFu));

    std::vector<std::byte> out;
    out.reserve(static_cast<size_t>(targetBytes));

    constexpr std::uint8_t kSignature[8] = {0x89u, 0x50u, 0x4Eu, 0x47u, 0x0Du, 0x0Au, 0x1Au, 0x0Au};
    AppendBytes(out, kSignature, sizeof(kSignature));

    std::array<std::uint8_t, 13> ihdr{};
    ihdr[0]  = static_cast<std::uint8_t>((width >> 24u) & 0xFFu);
    ihdr[1]  = static_cast<std::uint8_t>((width >> 16u) & 0xFFu);
    ihdr[2]  = static_cast<std::uint8_t>((width >> 8u) & 0xFFu);
    ihdr[3]  = static_cast<std::uint8_t>(width & 0xFFu);
    ihdr[4]  = static_cast<std::uint8_t>((height >> 24u) & 0xFFu);
    ihdr[5]  = static_cast<std::uint8_t>((height >> 16u) & 0xFFu);
    ihdr[6]  = static_cast<std::uint8_t>((height >> 8u) & 0xFFu);
    ihdr[7]  = static_cast<std::uint8_t>(height & 0xFFu);
    ihdr[8]  = 8u; // bit depth
    ihdr[9]  = 2u; // color type: truecolor
    ihdr[10] = 0u; // compression
    ihdr[11] = 0u; // filter
    ihdr[12] = 0u; // interlace

    constexpr std::array<std::uint8_t, 4> kChunkIHDR = {{'I', 'H', 'D', 'R'}};
    constexpr std::array<std::uint8_t, 4> kChunkIDAT = {{'I', 'D', 'A', 'T'}};
    constexpr std::array<std::uint8_t, 4> kChunkIEND = {{'I', 'E', 'N', 'D'}};
    constexpr std::array<std::uint8_t, 4> kChunkPad  = {{'p', 'A', 'D', 'd'}};

    AppendPngChunk(out, kChunkIHDR, ihdr.data(), ihdr.size());
    AppendPngChunk(out, kChunkIDAT, zlib.data(), zlib.size());

    const uint64_t sizeWithIend = static_cast<uint64_t>(out.size()) + 12ull;
    if (targetBytes < sizeWithIend)
    {
        return {};
    }

    const uint64_t paddingBytes = targetBytes - sizeWithIend;
    if (paddingBytes > 0)
    {
        if (paddingBytes < 12ull)
        {
            return {};
        }

        const uint64_t dataBytes = paddingBytes - 12ull;
        if (dataBytes > static_cast<uint64_t>(std::numeric_limits<std::uint32_t>::max()))
        {
            return {};
        }

        std::vector<std::uint8_t> padding;
        padding.resize(static_cast<size_t>(dataBytes));
        for (size_t index = 0; index < padding.size(); ++index)
        {
            padding[index] = GenerateDummyByte(DummyFillKind::Binary, seed ^ 0xA5A5A5A5u, static_cast<uint64_t>(index));
        }

        AppendPngChunk(out, kChunkPad, padding.data(), padding.size());
    }

    AppendPngChunk(out, kChunkIEND, nullptr, 0);
    return out;
}

struct JpegHuffmanTable
{
    std::array<std::uint16_t, 256> codes{};
    std::array<std::uint8_t, 256> sizes{};
};

JpegHuffmanTable BuildJpegHuffmanTable(const std::array<std::uint8_t, 16>& counts, const std::uint8_t* values, size_t valueCount)
{
    JpegHuffmanTable table{};

    std::uint16_t code = 0;
    size_t index       = 0;
    for (size_t bitCount = 0; bitCount < counts.size(); ++bitCount)
    {
        const std::uint8_t count = counts[bitCount];
        for (std::uint8_t i = 0; i < count; ++i)
        {
            if (index >= valueCount)
            {
                break;
            }

            const std::uint8_t symbol = values[index];
            ++index;

            table.codes[symbol] = code;
            table.sizes[symbol] = static_cast<std::uint8_t>(bitCount + 1u);
            ++code;
        }
        code = static_cast<std::uint16_t>(code << 1u);
    }

    return table;
}

class JpegBitWriter
{
public:
    void WriteBits(std::uint16_t bits, std::uint8_t bitCount)
    {
        if (bitCount == 0)
        {
            return;
        }

        const std::uint32_t mask = (bitCount >= 32u) ? 0xFFFFFFFFu : ((1u << bitCount) - 1u);
        _bitBuffer               = (_bitBuffer << bitCount) | (static_cast<std::uint32_t>(bits) & mask);
        _bitCount                = static_cast<std::uint8_t>(_bitCount + bitCount);

        while (_bitCount >= 8u)
        {
            const std::uint8_t byte = static_cast<std::uint8_t>((_bitBuffer >> (_bitCount - 8u)) & 0xFFu);
            _bytes.push_back(byte);
            if (byte == 0xFFu)
            {
                _bytes.push_back(0x00u);
            }

            _bitCount = static_cast<std::uint8_t>(_bitCount - 8u);
            if (_bitCount == 0)
            {
                _bitBuffer = 0;
            }
            else
            {
                _bitBuffer &= (1u << _bitCount) - 1u;
            }
        }
    }

    void FlushWithOnes()
    {
        if (_bitCount == 0)
        {
            return;
        }

        const std::uint32_t bits     = _bitBuffer & ((1u << _bitCount) - 1u);
        const std::uint8_t padBits   = static_cast<std::uint8_t>(8u - _bitCount);
        const std::uint8_t padMask   = static_cast<std::uint8_t>((1u << padBits) - 1u);
        const std::uint8_t byteValue = static_cast<std::uint8_t>((bits << padBits) | padMask);

        _bytes.push_back(byteValue);
        if (byteValue == 0xFFu)
        {
            _bytes.push_back(0x00u);
        }

        _bitBuffer = 0;
        _bitCount  = 0;
    }

    const std::vector<std::uint8_t>& Bytes() const noexcept
    {
        return _bytes;
    }

private:
    std::vector<std::uint8_t> _bytes;
    std::uint32_t _bitBuffer = 0;
    std::uint8_t _bitCount   = 0;
};

std::vector<std::byte> GenerateDummyJpeg(std::uint64_t seed, uint64_t targetBytes)
{
    constexpr std::array<std::uint8_t, 16> kDcCounts = {{0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0}};
    constexpr std::array<std::uint8_t, 12> kDcValues = {{0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}};

    constexpr std::array<std::uint8_t, 16> kAcCounts  = {{0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d}};
    constexpr std::array<std::uint8_t, 162> kAcValues = {
        {0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08,
         0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28,
         0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59,
         0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
         0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6,
         0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2,
         0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA}};

    static const JpegHuffmanTable dcTable = BuildJpegHuffmanTable(kDcCounts, kDcValues.data(), kDcValues.size());
    static const JpegHuffmanTable acTable = BuildJpegHuffmanTable(kAcCounts, kAcValues.data(), kAcValues.size());

    constexpr std::uint16_t width  = 64;
    constexpr std::uint16_t height = 64;

    JpegBitWriter writer;
    int previousDc = 0;

    for (std::uint32_t by = 0; by < (height / 8u); ++by)
    {
        for (std::uint32_t bx = 0; bx < (width / 8u); ++bx)
        {
            const std::uint64_t v    = Mix64(seed + (static_cast<std::uint64_t>(by) << 32u) + static_cast<std::uint64_t>(bx));
            const std::uint8_t pixel = static_cast<std::uint8_t>(v & 0xFFu);

            const int dc   = (static_cast<int>(pixel) - 128) * 8;
            const int diff = dc - previousDc;
            previousDc     = dc;

            unsigned int magnitude = static_cast<unsigned int>(diff < 0 ? -diff : diff);
            std::uint8_t category  = 0;
            while (magnitude != 0)
            {
                magnitude >>= 1u;
                ++category;
            }

            writer.WriteBits(dcTable.codes[category], dcTable.sizes[category]);

            if (category > 0)
            {
                const int base = (diff >= 0) ? diff : (diff + (1 << category) - 1);
                writer.WriteBits(static_cast<std::uint16_t>(base), category);
            }

            writer.WriteBits(acTable.codes[0x00], acTable.sizes[0x00]); // EOB
        }
    }

    writer.FlushWithOnes();

    constexpr std::array<std::uint8_t, 2> kSOI  = {{0xFFu, 0xD8u}};
    constexpr std::array<std::uint8_t, 2> kEOI  = {{0xFFu, 0xD9u}};
    constexpr std::array<std::uint8_t, 2> kAPP0 = {{0xFFu, 0xE0u}};
    constexpr std::array<std::uint8_t, 2> kDQT  = {{0xFFu, 0xDBu}};
    constexpr std::array<std::uint8_t, 2> kSOF0 = {{0xFFu, 0xC0u}};
    constexpr std::array<std::uint8_t, 2> kDHT  = {{0xFFu, 0xC4u}};
    constexpr std::array<std::uint8_t, 2> kSOS  = {{0xFFu, 0xDAu}};
    constexpr std::array<std::uint8_t, 2> kCOM  = {{0xFFu, 0xFEu}};

    std::vector<std::byte> base;
    base.reserve(1024u + writer.Bytes().size());

    AppendBytes(base, kSOI.data(), kSOI.size());

    // APP0 JFIF segment
    AppendBytes(base, kAPP0.data(), kAPP0.size());
    AppendU16BE(base, 16u);
    const std::uint8_t jfif[14] = {'J', 'F', 'I', 'F', 0x00u, 0x01u, 0x01u, 0x00u, 0x00u, 0x01u, 0x00u, 0x01u, 0x00u, 0x00u};
    AppendBytes(base, jfif, sizeof(jfif));

    // DQT (one table, all 8s)
    AppendBytes(base, kDQT.data(), kDQT.size());
    AppendU16BE(base, 67u);
    base.push_back(static_cast<std::byte>(0x00u));
    for (unsigned int i = 0; i < 64u; ++i)
    {
        base.push_back(static_cast<std::byte>(8u));
    }

    // SOF0 (baseline, grayscale)
    AppendBytes(base, kSOF0.data(), kSOF0.size());
    AppendU16BE(base, 11u);
    base.push_back(static_cast<std::byte>(8u));
    AppendU16BE(base, height);
    AppendU16BE(base, width);
    base.push_back(static_cast<std::byte>(1u));    // components
    base.push_back(static_cast<std::byte>(1u));    // component id
    base.push_back(static_cast<std::byte>(0x11u)); // sampling
    base.push_back(static_cast<std::byte>(0u));    // quant table

    // DHT (DC+AC luminance)
    AppendBytes(base, kDHT.data(), kDHT.size());
    AppendU16BE(base, static_cast<std::uint16_t>(2u + (1u + 16u + 12u) + (1u + 16u + 162u)));
    base.push_back(static_cast<std::byte>(0x00u));
    for (const auto v : kDcCounts)
    {
        base.push_back(static_cast<std::byte>(v));
    }
    for (const auto v : kDcValues)
    {
        base.push_back(static_cast<std::byte>(v));
    }
    base.push_back(static_cast<std::byte>(0x10u));
    for (const auto v : kAcCounts)
    {
        base.push_back(static_cast<std::byte>(v));
    }
    for (const auto v : kAcValues)
    {
        base.push_back(static_cast<std::byte>(v));
    }

    // SOS
    AppendBytes(base, kSOS.data(), kSOS.size());
    AppendU16BE(base, 8u);
    base.push_back(static_cast<std::byte>(1u));    // components
    base.push_back(static_cast<std::byte>(1u));    // component id
    base.push_back(static_cast<std::byte>(0x00u)); // DC=0, AC=0
    base.push_back(static_cast<std::byte>(0u));    // Ss
    base.push_back(static_cast<std::byte>(63u));   // Se
    base.push_back(static_cast<std::byte>(0u));    // AhAl

    const uint64_t baseWithoutCom = static_cast<uint64_t>(base.size()) + static_cast<uint64_t>(writer.Bytes().size()) + static_cast<uint64_t>(kEOI.size());

    if (targetBytes < baseWithoutCom)
    {
        return {};
    }

    uint64_t remaining = targetBytes - baseWithoutCom;

    std::vector<std::byte> out;
    out.reserve(static_cast<size_t>(targetBytes));

    // Copy SOI+APP0 marker segment first, then insert COM segments, then the rest.
    out.insert(out.end(), base.begin(), base.begin() + 2 + 2 + 2 + sizeof(jfif));

    while (remaining > 0)
    {
        const uint64_t segmentTotal = std::min<uint64_t>(remaining, 65537ull);
        if (segmentTotal < 4ull)
        {
            break;
        }

        const uint64_t dataLen          = segmentTotal - 4ull;
        const std::uint16_t lengthField = static_cast<std::uint16_t>(dataLen + 2ull);

        AppendBytes(out, kCOM.data(), kCOM.size());
        AppendU16BE(out, lengthField);

        for (uint64_t i = 0; i < dataLen; ++i)
        {
            out.push_back(static_cast<std::byte>(GenerateDummyByte(DummyFillKind::Binary, seed ^ 0xC3C3C3C3u, i)));
        }

        remaining -= segmentTotal;
    }

    // Append remaining base data (everything after APP0 segment).
    out.insert(out.end(), base.begin() + 2 + 2 + 2 + sizeof(jfif), base.end());

    // Entropy-coded data and EOI.
    for (const auto byte : writer.Bytes())
    {
        out.push_back(static_cast<std::byte>(byte));
    }
    AppendBytes(out, kEOI.data(), kEOI.size());

    if (out.size() != static_cast<size_t>(targetBytes))
    {
        return {};
    }

    return out;
}

bool IsHighSurrogate(wchar_t value) noexcept
{
    return value >= 0xD800 && value <= 0xDBFF;
}

void TrimToLength(std::wstring& text, size_t maxChars)
{
    if (text.size() <= maxChars)
    {
        return;
    }

    text.resize(maxChars);
    if (! text.empty() && IsHighSurrogate(text.back()))
    {
        text.pop_back();
    }
}

__int64 FileTimeToInt64(const FILETIME& fileTime) noexcept
{
    ULARGE_INTEGER value{};
    value.LowPart  = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<__int64>(value.QuadPart);
}

__int64 GetNowFileTime() noexcept
{
    FILETIME now{};
    GetSystemTimeAsFileTime(&now);
    return FileTimeToInt64(now);
}

bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept
{
    return (static_cast<unsigned long>(flags) & static_cast<unsigned long>(flag)) != 0u;
}

bool IsCancellationHr(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

HRESULT NormalizeCancellation(HRESULT hr) noexcept
{
    if (IsCancellationHr(hr))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

HRESULT BuildFileInfoBuffer(const std::vector<DummyEntry>& entries, std::vector<std::byte>* outBuffer, unsigned long* outUsedBytes) noexcept
{
    if (! outBuffer || ! outUsedBytes)
    {
        return E_POINTER;
    }

    outBuffer->clear();
    *outUsedBytes = 0;

    if (entries.empty())
    {
        return S_OK;
    }

    const size_t baseSize = offsetof(FileInfo, FileName);
    size_t totalBytes     = 0;

    for (const auto& entry : entries)
    {
        const size_t nameChars = entry.name.size();
        if (nameChars > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const size_t nameBytes = nameChars * sizeof(wchar_t);
        const size_t entrySize = AlignUp(baseSize + nameBytes + sizeof(wchar_t), kEntryAlignment);
        if (totalBytes > std::numeric_limits<unsigned long>::max() - entrySize)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        totalBytes += entrySize;
    }

    outBuffer->assign(totalBytes, std::byte{0});
    *outUsedBytes = static_cast<unsigned long>(totalBytes);

    std::byte* base = outBuffer->data();
    size_t offset   = 0;

    for (size_t index = 0; index < entries.size(); ++index)
    {
        const auto& entry      = entries[index];
        const size_t nameChars = entry.name.size();
        const size_t nameBytes = nameChars * sizeof(wchar_t);
        const size_t entrySize = AlignUp(baseSize + nameBytes + sizeof(wchar_t), kEntryAlignment);

        auto* info = reinterpret_cast<FileInfo*>(base + offset);
        ZeroMemory(info, entrySize);

        info->FileIndex      = static_cast<unsigned long>(index);
        info->FileAttributes = entry.attributes;
        info->FileNameSize   = static_cast<unsigned long>(nameBytes);
        info->CreationTime   = entry.creationTime;
        info->LastAccessTime = entry.lastAccessTime;
        info->LastWriteTime  = entry.lastWriteTime;
        info->ChangeTime     = entry.changeTime;
        info->EndOfFile      = static_cast<__int64>(entry.sizeBytes);

        uint64_t allocation = entry.sizeBytes;
        if (allocation > 0)
        {
            allocation = AlignUp(allocation, static_cast<size_t>(4096));
        }
        if (allocation > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
        {
            allocation = static_cast<uint64_t>(std::numeric_limits<__int64>::max());
        }
        info->AllocationSize = static_cast<__int64>(allocation);

        if (nameBytes > 0)
        {
            CopyMemory(info->FileName, entry.name.data(), nameBytes);
        }
        info->FileName[nameChars] = L'\0';

        if (index + 1 < entries.size())
        {
            info->NextEntryOffset = static_cast<unsigned long>(entrySize);
        }

        offset += entrySize;
    }

    return S_OK;
}

#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
struct OperationContext
{
    FileSystemOperation type      = FILESYSTEM_COPY;
    IFileSystemCallback* callback = nullptr;
    void* callbackCookie          = nullptr;
    uint64_t progressStreamId     = 0;
    FileSystemOptions optionsState{};
    FileSystemOptions* options          = nullptr;
    uint64_t virtualLimitBytesPerSecond = 0;
    unsigned long latencyMilliseconds   = 0;
    std::uint64_t throughputSeed        = 0;
    unsigned long totalItems            = 0;
    unsigned long completedItems        = 0;
    uint64_t totalBytes                 = 0;
    uint64_t completedBytes             = 0;
    bool continueOnError                = false;
    bool allowOverwrite                 = false;
    bool allowReplaceReadonly           = false;
    bool recursive                      = false;
    bool useRecycleBin                  = false;
    FileSystemArenaOwner itemArena;
    FileSystemArenaOwner progressArena;
    const wchar_t* itemSource          = nullptr;
    const wchar_t* itemDestination     = nullptr;
    const wchar_t* progressSource      = nullptr;
    const wchar_t* progressDestination = nullptr;
};
#pragma warning(pop)

void InitializeOperationContext(OperationContext& context,
                                FileSystemOperation type,
                                FileSystemFlags flags,
                                const FileSystemOptions* options,
                                IFileSystemCallback* callback,
                                void* cookie,
                                unsigned long totalItems) noexcept
{
    context.type             = type;
    context.callback         = callback;
    context.callbackCookie   = callback != nullptr ? cookie : nullptr;
    context.progressStreamId = callback != nullptr ? static_cast<uint64_t>(GetCurrentThreadId()) : 0;
    context.optionsState     = {};
    if (options)
    {
        context.optionsState = *options;
        context.options      = &context.optionsState;
    }
    else
    {
        context.options = nullptr;
    }
    context.virtualLimitBytesPerSecond = 0;
    context.latencyMilliseconds        = 0;
    context.throughputSeed             = 0;
    context.totalItems                 = totalItems;
    context.completedItems             = 0;
    context.totalBytes                 = 0;
    context.completedBytes             = 0;
    context.continueOnError            = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    context.allowOverwrite             = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    context.allowReplaceReadonly       = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    context.recursive                  = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    context.useRecycleBin              = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    context.itemSource                 = nullptr;
    context.itemDestination            = nullptr;
    context.progressSource             = nullptr;
    context.progressDestination        = nullptr;
}

HRESULT CalculateStringBytes(const wchar_t* text, unsigned long* outBytes) noexcept
{
    if (! outBytes)
    {
        return E_POINTER;
    }

    if (! text)
    {
        *outBytes = 0;
        return S_OK;
    }

    const size_t length = wcslen(text);
    if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *outBytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
    return S_OK;
}

HRESULT BuildArenaForPaths(
    FileSystemArenaOwner& arenaOwner, const wchar_t* source, const wchar_t* destination, const wchar_t** outSource, const wchar_t** outDestination) noexcept
{
    if (! outSource || ! outDestination)
    {
        return E_POINTER;
    }

    *outSource      = nullptr;
    *outDestination = nullptr;

    unsigned long sourceBytes = 0;
    HRESULT hr                = CalculateStringBytes(source, &sourceBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long destinationBytes = 0;
    hr                             = CalculateStringBytes(destination, &destinationBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long totalBytes = sourceBytes;
    if (destinationBytes > 0)
    {
        if (totalBytes > std::numeric_limits<unsigned long>::max() - destinationBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += destinationBytes;
    }

    FileSystemArena* arena = arenaOwner.Get();
    if (! arena || arena->buffer == nullptr || arena->capacityBytes < totalBytes)
    {
        hr = arenaOwner.Initialize(totalBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        arena = arenaOwner.Get();
    }

    if (arena && arena->buffer)
    {
        arena->usedBytes = 0;
    }

    if (sourceBytes > 0)
    {
        auto* sourceBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, sourceBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourceBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t sourceLength = (sourceBytes / sizeof(wchar_t)) - 1u;
        if (sourceLength > 0)
        {
            CopyMemory(sourceBuffer, source, sourceLength * sizeof(wchar_t));
        }
        sourceBuffer[sourceLength] = L'\0';
        *outSource                 = sourceBuffer;
    }

    if (destinationBytes > 0)
    {
        auto* destinationBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, destinationBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! destinationBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t destinationLength = (destinationBytes / sizeof(wchar_t)) - 1u;
        if (destinationLength > 0)
        {
            CopyMemory(destinationBuffer, destination, destinationLength * sizeof(wchar_t));
        }
        destinationBuffer[destinationLength] = L'\0';
        *outDestination                      = destinationBuffer;
    }

    return S_OK;
}

HRESULT SetItemPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.itemArena, source, destination, &context.itemSource, &context.itemDestination);
}

HRESULT SetProgressPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.progressArena, source, destination, &context.progressSource, &context.progressDestination);
}

HRESULT CheckCancel(OperationContext& context) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    BOOL cancel = FALSE;
    HRESULT hr  = context.callback->FileSystemShouldCancel(&cancel, context.callbackCookie);
    hr          = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (cancel)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

void UpdateEffectiveBandwidthLimit(OperationContext& context) noexcept;

HRESULT ReportProgress(OperationContext& context, uint64_t currentItemTotalBytes, uint64_t currentItemCompletedBytes) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    UpdateEffectiveBandwidthLimit(context);

    HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                      context.totalItems,
                                                      context.completedItems,
                                                      context.totalBytes,
                                                      context.completedBytes,
                                                      context.progressSource,
                                                      context.progressDestination,
                                                      currentItemTotalBytes,
                                                      currentItemCompletedBytes,
                                                      context.options,
                                                      context.progressStreamId,
                                                      context.callbackCookie);
    hr         = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

HRESULT ReportItemCompleted(OperationContext& context, unsigned long itemIndex, HRESULT status) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    UpdateEffectiveBandwidthLimit(context);

    HRESULT hr = context.callback->FileSystemItemCompleted(
        context.type, itemIndex, context.itemSource, context.itemDestination, status, context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

std::wstring AppendPath(const std::wstring& base, std::wstring_view leaf)
{
    std::wstring result = base;
    if (! result.empty())
    {
        const wchar_t lastChar = result.back();
        if (lastChar != L'\\' && lastChar != L'/')
        {
            result.push_back(L'\\');
        }
    }
    if (! leaf.empty())
    {
        result.append(leaf.data(), leaf.size());
    }
    return result;
}

std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept
{
    while (! path.empty())
    {
        const wchar_t last = path.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

std::wstring_view GetPathLeaf(std::wstring_view path) noexcept
{
    const std::wstring_view trimmed = TrimTrailingSeparators(path);
    if (trimmed.empty())
    {
        return trimmed;
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed;
    }

    return trimmed.substr(pos + 1);
}

uint64_t GetEffectiveBandwidthLimitBytesPerSecond(const OperationContext& context, uint64_t virtualLimitBytesPerSecond) noexcept
{
    const uint64_t hostLimit = context.options ? context.options->bandwidthLimitBytesPerSecond : 0;
    if (hostLimit == 0)
    {
        return virtualLimitBytesPerSecond;
    }

    if (virtualLimitBytesPerSecond == 0)
    {
        return hostLimit;
    }

    return std::min(hostLimit, virtualLimitBytesPerSecond);
}

void UpdateEffectiveBandwidthLimit(OperationContext& context) noexcept
{
    if (! context.options)
    {
        return;
    }

    const uint64_t effectiveLimit                 = GetEffectiveBandwidthLimitBytesPerSecond(context, context.virtualLimitBytesPerSecond);
    context.options->bandwidthLimitBytesPerSecond = effectiveLimit;
}

HRESULT SleepWithCancelChecks(OperationContext& context, uint64_t milliseconds) noexcept
{
    if (milliseconds == 0)
    {
        return S_OK;
    }

    constexpr uint64_t kMaxSleepMs = static_cast<uint64_t>(std::numeric_limits<DWORD>::max());
    uint64_t remaining             = std::min(milliseconds, kMaxSleepMs);

    constexpr DWORD kSleepQuantumMs = 50;
    while (remaining > 0)
    {
        const DWORD slice = static_cast<DWORD>(std::min<uint64_t>(remaining, static_cast<uint64_t>(kSleepQuantumMs)));
        ::Sleep(slice);
        remaining -= slice;

        const HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

HRESULT
ReportThrottledByteProgress(OperationContext& context, uint64_t itemTotalBytes, uint64_t baseCompletedBytes, uint64_t virtualLimitBytesPerSecond) noexcept
{
    uint64_t itemCompletedBytes = 0;
    context.completedBytes      = baseCompletedBytes;

    HRESULT hr = ReportProgress(context, itemTotalBytes, itemCompletedBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.latencyMilliseconds > 0)
    {
        uint64_t accessCount = 1ull;
        if (context.type == FILESYSTEM_COPY || context.type == FILESYSTEM_MOVE || context.type == FILESYSTEM_RENAME)
        {
            accessCount = 2ull;
        }

        const uint64_t latencyMs = static_cast<uint64_t>(context.latencyMilliseconds) * accessCount;
        hr                       = SleepWithCancelChecks(context, latencyMs);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (itemTotalBytes == 0)
    {
        return S_OK;
    }

    std::uint64_t seed = context.throughputSeed;
    seed               = CombineSeed(seed, baseCompletedBytes);
    seed               = CombineSeed(seed, itemTotalBytes);
    std::mt19937 rng   = MakeRng(seed);

    while (itemCompletedBytes < itemTotalBytes)
    {
        const uint64_t effectiveMaxBytesPerSecond = GetEffectiveBandwidthLimitBytesPerSecond(context, virtualLimitBytesPerSecond);
        if (effectiveMaxBytesPerSecond == 0)
        {
            itemCompletedBytes     = itemTotalBytes;
            context.completedBytes = baseCompletedBytes + itemCompletedBytes;
            return ReportProgress(context, itemTotalBytes, itemCompletedBytes);
        }

        const uint64_t maxBytesPerSecond = effectiveMaxBytesPerSecond;
        uint64_t minBytesPerSecond       = std::max<uint64_t>(1ull, maxBytesPerSecond - (maxBytesPerSecond / 5ull)); // ~80%
        uint64_t jitterMaxBytesPerSecond = maxBytesPerSecond;
        if (maxBytesPerSecond >= 10ull && RandomChance(rng, 1, 200))
        {
            minBytesPerSecond       = std::max<uint64_t>(1ull, maxBytesPerSecond / 10ull);             // ~10%
            jitterMaxBytesPerSecond = std::max<uint64_t>(minBytesPerSecond, maxBytesPerSecond / 3ull); // ~33%
        }
        else if (maxBytesPerSecond >= 10ull && RandomChance(rng, 1, 25))
        {
            minBytesPerSecond = std::max<uint64_t>(1ull, maxBytesPerSecond / 2ull); // ~50%
        }

        const uint64_t currentBytesPerSecond = RandomRange64(rng, minBytesPerSecond, jitterMaxBytesPerSecond);
        const uint64_t remaining             = itemTotalBytes - itemCompletedBytes;
        const uint64_t step                  = std::max<uint64_t>(1ull, currentBytesPerSecond / 10ull);
        const uint64_t chunk                 = std::min(step, remaining);

        const auto sleepDuration =
            std::chrono::duration<double>(static_cast<double>(chunk) / static_cast<double>(std::max<uint64_t>(1ull, currentBytesPerSecond)));
        const double sleepMsD  = sleepDuration.count() * 1000.0;
        const uint64_t sleepMs = sleepMsD > 0.0 ? static_cast<uint64_t>(sleepMsD + 0.5) : 0ull;
        hr                     = SleepWithCancelChecks(context, sleepMs);
        if (FAILED(hr))
        {
            return hr;
        }

        itemCompletedBytes += chunk;
        context.completedBytes = baseCompletedBytes + itemCompletedBytes;

        hr = ReportProgress(context, itemTotalBytes, itemCompletedBytes);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}
} // namespace

DummyFilesInformation::DummyFilesInformation(std::vector<std::byte> buffer, unsigned long count, unsigned long usedBytes) noexcept : _buffer(std::move(buffer))
{
    _count     = count;
    _usedBytes = usedBytes;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
    {
        *ppvObject = static_cast<IFilesInformation*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE DummyFilesInformation::AddRef() noexcept
{
    return ++_refCount;
}

ULONG STDMETHODCALLTYPE DummyFilesInformation::Release() noexcept
{
    const ULONG remaining = --_refCount;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    if (! ppFileInfo)
    {
        return E_POINTER;
    }

    if (_usedBytes == 0)
    {
        *ppFileInfo = nullptr;
        return S_OK;
    }

    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::GetBufferSize(unsigned long* pSize) noexcept
{
    if (! pSize)
    {
        return E_POINTER;
    }

    *pSize = _usedBytes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::GetAllocatedSize(unsigned long* pSize) noexcept
{
    if (! pSize)
    {
        return E_POINTER;
    }

    if (_buffer.size() > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *pSize = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::GetCount(unsigned long* pCount) noexcept
{
    if (! pCount)
    {
        return E_POINTER;
    }

    *pCount = _count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DummyFilesInformation::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    if (! ppEntry)
    {
        return E_POINTER;
    }

    if (index >= _count || _usedBytes == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    return LocateEntry(index, ppEntry);
}

size_t DummyFilesInformation::ComputeEntrySize(const FileInfo* entry) const noexcept
{
    if (! entry)
    {
        return 0;
    }

    const size_t baseSize = offsetof(FileInfo, FileName);
    const size_t nameSize = static_cast<size_t>(entry->FileNameSize);
    return AlignUp(baseSize + nameSize + sizeof(wchar_t), kEntryAlignment);
}

HRESULT DummyFilesInformation::LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept
{
    const std::byte* base      = _buffer.data();
    size_t offset              = 0;
    unsigned long currentIndex = 0;

    while (offset < _usedBytes && offset + sizeof(FileInfo) <= _buffer.size())
    {
        auto* entry = reinterpret_cast<const FileInfo*>(base + offset);

        if (currentIndex == index)
        {
            *ppEntry = const_cast<FileInfo*>(entry);
            return S_OK;
        }

        const size_t advance = (entry->NextEntryOffset != 0) ? static_cast<size_t>(entry->NextEntryOffset) : ComputeEntrySize(entry);
        if (advance == 0)
        {
            break;
        }

        offset += advance;
        ++currentIndex;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
}

FileSystemDummy::FileSystemDummy()
{
    _metaData.id          = kPluginId;
    _metaData.shortId     = kPluginShortId;
    _metaData.name        = kPluginName;
    _metaData.description = kPluginDescription;
    _metaData.author      = kPluginAuthor;
    _metaData.version     = kPluginVersion;

    bool needsDefaultConfig = false;
    {
        std::scoped_lock lock(_mutex);
        needsDefaultConfig = _configurationJson.empty();
    }

    if (needsDefaultConfig)
    {
        static_cast<void>(SetConfiguration(nullptr));
    }
}

FileSystemDummy::~FileSystemDummy()
{
    std::vector<std::shared_ptr<DirectoryWatchRegistration>> removedWatches;

    {
        std::unique_lock lock(_watchMutex);
        auto it = _directoryWatches.begin();
        while (it != _directoryWatches.end())
        {
            const auto& entry = *it;
            if (! entry || entry->owner != this)
            {
                ++it;
                continue;
            }

            removedWatches.push_back(entry);
            it = _directoryWatches.erase(it);
        }

        if (removedWatches.empty())
        {
            // No watches to drain; fall through to roots cleanup.
        }
        else
        {
            for (const auto& removed : removedWatches)
            {
                if (! removed)
                {
                    continue;
                }

                removed->active.store(false, std::memory_order_release);
            }

            _watchCv.wait(lock,
                          [&]
                          {
                              for (const auto& removed : removedWatches)
                              {
                                  if (! removed)
                                  {
                                      continue;
                                  }

                                  const bool reentrant           = static_cast<const void*>(removed.get()) == g_activeDirectoryWatchCallback;
                                  const uint32_t desiredInFlight = reentrant ? 1u : 0u;
                                  if (removed->inFlight.load(std::memory_order_acquire) > desiredInFlight)
                                  {
                                      return false;
                                  }
                              }
                              return true;
                          });
        }
    }

    // Iteratively free deeply-nested trees before the inline static _roots destructs at
    // process exit (which would recurse through DummyNode::children and stack-overflow on
    // trees created by the compare self-test, which makes 1024-level-deep directories).
    std::scoped_lock lock(_mutex);
    ClearRootsIteratively();
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kSchemaJson;
    return S_OK;
}

// Iteratively free all DummyNode trees to avoid stack overflow on deeply-nested trees
// (e.g. the compare self-test creates 1024-level deep directories).
// Must be called while _mutex is held.
void FileSystemDummy::ClearRootsIteratively() noexcept
{
    // Collect all root-level DummyNode unique_ptrs.
    std::vector<std::unique_ptr<DummyNode>> pending;
    pending.reserve(_roots.size());
    for (auto& root : _roots)
    {
        if (root && root->node)
        {
            pending.push_back(std::move(root->node));
        }
    }
    _roots.clear();

    // Iteratively drain children into `pending`, then let each childless node
    // be destroyed at the end of the loop body — O(n) stack depth.
    while (! pending.empty())
    {
        auto node = std::move(pending.back());
        pending.pop_back();
        if (node)
        {
            for (auto& child : node->children)
            {
                if (child)
                {
                    pending.push_back(std::move(child));
                }
            }
            node->children.clear(); // destroy child vector before node goes out of scope
            // node destroyed here — children are empty, so destructor is trivial
        }
    }
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    unsigned long maxChildrenPerDirectory    = 42;
    unsigned long maxDepth                   = 10;
    unsigned int seed                        = 42;
    unsigned long latencyMilliseconds        = 0;
    std::wstring virtualSpeedLimitText       = L"0";
    uint64_t virtualSpeedLimitBytesPerSecond = 0;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view configText(configurationJsonUtf8);

        yyjson_doc* doc = yyjson_read(configText.data(), configText.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (doc)
        {
            auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

            yyjson_val* root = yyjson_doc_get_root(doc);
            if (root && yyjson_is_obj(root))
            {
                yyjson_val* maxChildren = yyjson_obj_get(root, "maxChildrenPerDirectory");
                if (maxChildren && yyjson_is_int(maxChildren))
                {
                    const int64_t value = yyjson_get_int(maxChildren);
                    if (value >= 0)
                    {
                        maxChildrenPerDirectory = static_cast<unsigned long>(std::min<int64_t>(value, 20000));
                    }
                }

                yyjson_val* maxDepthVal = yyjson_obj_get(root, "maxDepth");
                if (maxDepthVal && yyjson_is_int(maxDepthVal))
                {
                    const int64_t value = yyjson_get_int(maxDepthVal);
                    if (value >= 0)
                    {
                        maxDepth = static_cast<unsigned long>(std::min<int64_t>(value, 1024));
                    }
                }

                yyjson_val* seedVal = yyjson_obj_get(root, "seed");
                if (seedVal && yyjson_is_int(seedVal))
                {
                    const int64_t value = yyjson_get_int(seedVal);
                    if (value >= 0)
                    {
                        seed = static_cast<unsigned int>(std::min<int64_t>(value, std::numeric_limits<unsigned int>::max()));
                    }
                }

                yyjson_val* latencyVal = yyjson_obj_get(root, "latencyMs");
                if (latencyVal && yyjson_is_int(latencyVal))
                {
                    const int64_t value = yyjson_get_int(latencyVal);
                    if (value >= 0)
                    {
                        latencyMilliseconds = static_cast<unsigned long>(std::min<int64_t>(value, 1000));
                    }
                }

                yyjson_val* virtualSpeedVal = yyjson_obj_get(root, "virtualSpeedLimit");
                if (virtualSpeedVal && yyjson_is_str(virtualSpeedVal))
                {
                    const char* speedText = yyjson_get_str(virtualSpeedVal);
                    if (speedText)
                    {
                        uint64_t parsed = 0;
                        if (TryParseThroughputText(speedText, parsed))
                        {
                            virtualSpeedLimitBytesPerSecond = parsed;

                            const std::wstring wideText = Utf16FromUtf8(speedText);
                            if (! wideText.empty())
                            {
                                virtualSpeedLimitText = wideText;
                            }
                        }
                    }
                }
            }
        }
    }

    const std::string speedLimitTextUtf8 = Utf8FromUtf16(EscapeJsonString(virtualSpeedLimitText));
    const std::string newConfigJson =
        std::format("{{\"maxChildrenPerDirectory\":{},\"maxDepth\":{},\"seed\":{},\"latencyMs\":{},\"virtualSpeedLimit\":\"{}\"}}",
                    maxChildrenPerDirectory,
                    maxDepth,
                    seed,
                    latencyMilliseconds,
                    speedLimitTextUtf8);

    {
        std::scoped_lock lock(_mutex);

        const bool structureChanged =
            _configurationJson.empty() || _maxChildrenPerDirectory != maxChildrenPerDirectory || _maxDepth != maxDepth || _seed != seed;

        _maxChildrenPerDirectory = maxChildrenPerDirectory;
        _maxDepth                = maxDepth;
        _seed                    = seed;
        _latencyMilliseconds     = latencyMilliseconds;
        _virtualSpeedLimitText   = std::move(virtualSpeedLimitText);
        _virtualSpeedLimitBytesPerSecond.store(virtualSpeedLimitBytesPerSecond, std::memory_order_release);
        _configurationJson = newConfigJson;

        if (structureChanged)
        {
            const std::uint64_t effectiveSeed      = _seed == 0 ? static_cast<std::uint64_t>(GetTickCount64()) : static_cast<std::uint64_t>(_seed);
            const std::uint64_t generationBaseTime = ComputeGenerationBaseTime(effectiveSeed);

            ClearRootsIteratively();
            _effectiveSeed      = effectiveSeed;
            _generationBaseTime = generationBaseTime;
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::scoped_lock lock(_mutex);
    *configurationJsonUtf8 = _configurationJson.empty() ? "{}" : _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::scoped_lock lock(_mutex);
    const bool isDefault = _maxChildrenPerDirectory == 42 && _maxDepth == 10 && _seed == 42 && _latencyMilliseconds == 0 &&
                           _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire) == 0;
    *pSomethingToSave = isDefault ? FALSE : TRUE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    _menuEntries.clear();
    _menuEntryView.clear();

    const std::wstring label = _metaData.name ? _metaData.name : L"Dummy";
    MenuEntry header;
    header.flags = NAV_MENU_ITEM_FLAG_HEADER;
    header.label = label;
    _menuEntries.push_back(std::move(header));

    MenuEntry separator;
    separator.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
    _menuEntries.push_back(std::move(separator));

    MenuEntry entry;
    entry.label = L"/";
    entry.path  = L"/";
    _menuEntries.push_back(std::move(entry));

    _menuEntryView.reserve(_menuEntries.size());
    for (const auto& e : _menuEntries)
    {
        NavigationMenuItem item{};
        item.flags     = e.flags;
        item.label     = e.label.empty() ? nullptr : e.label.c_str();
        item.path      = e.path.empty() ? nullptr : e.path.c_str();
        item.iconPath  = e.iconPath.empty() ? nullptr : e.iconPath.c_str();
        item.commandId = e.commandId;
        _menuEntryView.push_back(item);
    }

    *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
    *count = static_cast<unsigned int>(_menuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback != nullptr ? cookie : nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetDriveInfo([[maybe_unused]] const wchar_t* path, DriveInfo* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    info->flags       = DRIVE_INFO_FLAG_NONE;
    info->displayName = nullptr;
    info->volumeLabel = nullptr;
    info->fileSystem  = nullptr;
    info->totalBytes  = 0;
    info->freeBytes   = 0;
    info->usedBytes   = 0;

    _driveDisplayName = _metaData.name ? _metaData.name : L"Dummy";
    _driveVolumeLabel = _driveDisplayName;
    _driveFileSystem  = L"DummyFS";

    constexpr uint64_t totalBytes = 8ull * 1024ull * 1024ull * 1024ull;
    const uint64_t freeBytes      = totalBytes / 2ull;

    info->flags       = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_DISPLAY_NAME);
    info->displayName = _driveDisplayName.c_str();

    info->flags       = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL);
    info->volumeLabel = _driveVolumeLabel.c_str();

    info->flags      = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
    info->fileSystem = _driveFileSystem.c_str();

    info->flags      = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_TOTAL_BYTES);
    info->totalBytes = totalBytes;

    info->flags     = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_FREE_BYTES);
    info->freeBytes = freeBytes;

    info->flags     = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_USED_BYTES);
    info->usedBytes = totalBytes - freeBytes;

    _driveInfo = *info;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetDriveMenuItems([[maybe_unused]] const wchar_t* path,
                                                             const NavigationMenuItem** items,
                                                             unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    _driveMenuEntries.clear();
    _driveMenuEntryView.clear();

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::ExecuteDriveMenuCommand([[maybe_unused]] unsigned int commandId, [[maybe_unused]] const wchar_t* path) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

    if (riid == __uuidof(IFileSystemDirectoryWatch))
    {
        *ppvObject = static_cast<IFileSystemDirectoryWatch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystemDummy::AddRef() noexcept
{
    return ++_refCount;
}

ULONG STDMETHODCALLTYPE FileSystemDummy::Release() noexcept
{
    const ULONG remaining = --_refCount;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

std::filesystem::path FileSystemDummy::NormalizePath(std::wstring_view path) const
{
    std::wstring text(path);
    for (auto& ch : text)
    {
        if (ch == L'/')
        {
            ch = L'\\';
        }
    }

    if (text.size() == 2 && text[1] == L':')
    {
        text.push_back(L'\\');
    }

    std::filesystem::path normalized     = std::filesystem::path(text).lexically_normal();
    std::wstring normalizedText          = normalized.wstring();
    const std::filesystem::path rootPath = normalized.root_path();
    std::wstring rootText                = rootPath.wstring();
    if (rootText.size() == 2 && rootText[1] == L':')
    {
        rootText.push_back(L'\\');
    }

    while (normalizedText.size() > rootText.size() && ! normalizedText.empty())
    {
        const wchar_t last = normalizedText.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        normalizedText.pop_back();
    }

    if (normalizedText.empty() && ! rootText.empty())
    {
        normalizedText = rootText;
    }

    return std::filesystem::path(normalizedText);
}

FileSystemDummy::DummyRoot* FileSystemDummy::FindRoot(std::wstring_view rootPath) noexcept
{
    for (const auto& root : _roots)
    {
        if (EqualsNoCase(root->rootPath, rootPath))
        {
            return root.get();
        }
    }
    return nullptr;
}

FileSystemDummy::DummyRoot* FileSystemDummy::GetOrCreateRoot(std::wstring_view rootPath)
{
    DummyRoot* root = FindRoot(rootPath);
    if (root)
    {
        return root;
    }

    auto newRoot                 = std::make_unique<DummyRoot>();
    newRoot->rootPath            = std::wstring(rootPath);
    const std::uint64_t rootSeed = CombineSeed(_effectiveSeed, rootPath);
    newRoot->node                = CreateNode(newRoot->rootPath, true, rootSeed);
    if (newRoot->node && newRoot->node->isDirectory && _maxChildrenPerDirectory > 0)
    {
        const unsigned long minChildCount = std::min(_maxChildrenPerDirectory, 2ul);
        if (newRoot->node->plannedChildCount < minChildCount)
        {
            newRoot->node->plannedChildCount = minChildCount;
        }
    }

    DummyRoot* created = newRoot.get();
    _roots.push_back(std::move(newRoot));
    return created;
}

HRESULT FileSystemDummy::ResolvePath(const std::filesystem::path& path, DummyNode** outNode, bool createMissing, bool requireDirectory) noexcept
{
    if (! outNode)
    {
        return E_POINTER;
    }

    *outNode = nullptr;

    const std::filesystem::path rootPath = path.root_path();
    std::wstring rootText                = rootPath.wstring();
    if (rootText.size() == 2 && rootText[1] == L':')
    {
        rootText.push_back(L'\\');
    }

    DummyRoot* root = GetOrCreateRoot(rootText);

    DummyNode* node                      = root->node.get();
    const std::filesystem::path relative = path.relative_path();
    for (const auto& part : relative)
    {
        const std::wstring segment = part.wstring();
        if (segment.empty() || segment == L".")
        {
            continue;
        }

        if (segment == L"..")
        {
            if (! node->parent)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            node = node->parent;
            continue;
        }

        if (! node->isDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        EnsureChildrenGenerated(*node);

        DummyNode* child = FindChild(node, segment);
        if (! child)
        {
            if (! createMissing)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            }

            const std::uint64_t childSeed = CombineSeed(node->generationSeed, segment);
            auto newNode                  = CreateNode(segment, true, childSeed);
            DummyNode* newNodeRaw         = newNode.get();
            AddChild(node, std::move(newNode));
            node = newNodeRaw;
            continue;
        }

        node = child;
    }

    if (requireDirectory && node && ! node->isDirectory)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    *outNode = node;
    return S_OK;
}

FileSystemDummy::DummyNode* FileSystemDummy::FindChild(DummyNode* parent, std::wstring_view name) const noexcept
{
    if (! parent)
    {
        return nullptr;
    }

    for (const auto& child : parent->children)
    {
        if (EqualsNoCase(child->name, name))
        {
            return child.get();
        }
    }
    return nullptr;
}

std::unique_ptr<FileSystemDummy::DummyNode> FileSystemDummy::ExtractChild(DummyNode* parent, DummyNode* child) noexcept
{
    if (! parent || ! child)
    {
        return nullptr;
    }

    auto& children = parent->children;
    for (auto it = children.begin(); it != children.end(); ++it)
    {
        if (it->get() == child)
        {
            std::unique_ptr<DummyNode> result = std::move(*it);
            children.erase(it);
            parent->plannedChildCount = static_cast<unsigned long>(children.size());
            TouchParent(parent);
            result->parent = nullptr;
            return result;
        }
    }

    return nullptr;
}

FileSystemDummy::DummyNode* FileSystemDummy::AddChild(DummyNode* parent, std::unique_ptr<DummyNode> child)
{
    if (! parent || ! child)
    {
        return nullptr;
    }

    child->parent  = parent;
    DummyNode* raw = child.get();
    parent->children.push_back(std::move(child));
    parent->plannedChildCount = static_cast<unsigned long>(parent->children.size());
    TouchParent(parent);
    return raw;
}

std::unique_ptr<FileSystemDummy::DummyNode> FileSystemDummy::CreateNode(std::wstring_view name, bool isDirectory, std::uint64_t generationSeed)
{
    auto node            = std::make_unique<DummyNode>();
    node->name           = std::wstring(name);
    node->isDirectory    = isDirectory;
    node->generationSeed = generationSeed;

    std::mt19937 rng = MakeRng(generationSeed);
    node->attributes = isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE;
    if (RandomChance(rng, 1, 8))
    {
        node->attributes |= FILE_ATTRIBUTE_READONLY;
    }
    if (RandomChance(rng, 1, 10))
    {
        node->attributes |= FILE_ATTRIBUTE_HIDDEN;
    }

    if (! isDirectory)
    {
        const DummyFileKind kind = GetDummyFileKind(name);
        node->sizeBytes          = MakeDummyFileSize(rng, kind);
    }

    const uint64_t now              = static_cast<uint64_t>(_generationBaseTime);
    const uint64_t maxOffsetSeconds = 60ull * 60ull * 24ull * 365ull * 3ull;
    const uint64_t offsetSeconds    = RandomRange64(rng, 0, maxOffsetSeconds);
    const uint64_t offsetTicks      = offsetSeconds * 10000000ull;
    uint64_t randomTime             = now;
    if (offsetTicks < now)
    {
        randomTime = now - offsetTicks;
    }

    node->creationTime   = static_cast<__int64>(randomTime);
    node->lastAccessTime = node->creationTime;
    node->lastWriteTime  = node->creationTime;
    node->changeTime     = node->creationTime;

    if (isDirectory)
    {
        node->plannedChildCount = RandomSkewedUpTo(rng, _maxChildrenPerDirectory);
    }

    return node;
}

void FileSystemDummy::EnsureChildrenGenerated(DummyNode& node)
{
    if (! node.isDirectory || node.childrenGenerated)
    {
        return;
    }

    GenerateChildren(node);
}

void FileSystemDummy::GenerateChildren(DummyNode& node)
{
    if (! node.isDirectory || node.childrenGenerated)
    {
        return;
    }

    std::mt19937 rng = MakeRng(node.generationSeed);

    node.childrenGenerated = true;
    node.children.clear();

    const unsigned long totalChildren = node.plannedChildCount;
    if (totalChildren == 0)
    {
        return;
    }

    node.children.reserve(static_cast<size_t>(totalChildren));
    auto addGeneratedChild = [&](std::unique_ptr<DummyNode> child)
    {
        if (! child)
        {
            return;
        }
        child->parent = &node;
        node.children.push_back(std::move(child));
    };

    unsigned long depth = 0;
    for (DummyNode* current = node.parent; current != nullptr; current = current->parent)
    {
        ++depth;
    }

    const bool isRoot              = node.parent == nullptr;
    const bool allowSubdirectories = (_maxDepth == 0) || (depth < _maxDepth);

    unsigned long maxDirs = allowSubdirectories ? (totalChildren / 2u) : 0;
    if (isRoot && allowSubdirectories && totalChildren > 1u && maxDirs == 0u)
    {
        maxDirs = 1u;
    }
    if (totalChildren > 0 && maxDirs > (totalChildren - 1u))
    {
        maxDirs = totalChildren - 1u; // ensure at least one file
    }

    unsigned long dirCount = maxDirs > 0 ? RandomSkewedUpTo(rng, maxDirs) : 0;
    if (isRoot && allowSubdirectories && totalChildren > 1u && dirCount == 0u)
    {
        dirCount = 1u;
    }
    const unsigned long fileCount = totalChildren - dirCount;

    for (unsigned long index = 0; index < dirCount; ++index)
    {
        std::wstring baseName = MakeRandomBaseName(rng);
        if (! IsNameValid(baseName))
        {
            baseName = L"folder";
        }

        const std::wstring suffix = std::format(L"_{:05}", index);
        if (baseName.size() + suffix.size() > kMaxNameLength)
        {
            if (kMaxNameLength > suffix.size())
            {
                TrimToLength(baseName, kMaxNameLength - suffix.size());
            }
        }
        if (baseName.empty())
        {
            baseName = L"folder";
        }

        std::wstring name = baseName;
        name.append(suffix);

        const std::uint64_t childSeed = DeriveChildSeed(node.generationSeed, index, true);
        auto child                    = CreateNode(name, true, childSeed);
        addGeneratedChild(std::move(child));
    }

    for (unsigned long index = 0; index < fileCount; ++index)
    {
        const unsigned long childIndex     = dirCount + index;
        const unsigned long extensionIndex = RandomRange(rng, 0, ArrayCount(kExtensions) - 1u);
        const std::wstring_view extension  = kExtensions[extensionIndex];

        std::wstring baseName = MakeRandomBaseName(rng);
        if (! IsNameValid(baseName))
        {
            baseName = L"file";
        }

        const std::wstring suffix  = std::format(L"_{:05}", childIndex);
        const size_t reservedChars = suffix.size() + extension.size();
        if (baseName.size() + reservedChars > kMaxNameLength)
        {
            if (kMaxNameLength > reservedChars)
            {
                TrimToLength(baseName, kMaxNameLength - reservedChars);
            }
        }
        if (baseName.empty())
        {
            baseName = L"file";
        }

        std::wstring name = baseName;
        name.append(suffix);
        name.append(extension.data(), extension.size());

        const std::uint64_t childSeed = DeriveChildSeed(node.generationSeed, childIndex, false);
        auto child                    = CreateNode(name, false, childSeed);
        addGeneratedChild(std::move(child));
    }

    node.plannedChildCount = static_cast<unsigned long>(node.children.size());
}

bool FileSystemDummy::IsNameValid(std::wstring_view name) const noexcept
{
    if (name.empty())
    {
        return false;
    }

    if (name == L"." || name == L"..")
    {
        return false;
    }

    constexpr std::wstring_view invalidChars = L"\\/:*?\"<>|";
    return name.find_first_of(invalidChars) == std::wstring_view::npos;
}

std::wstring FileSystemDummy::MakeUniqueName(DummyNode* parent, std::wstring_view baseName) const
{
    std::wstring candidate(baseName);
    if (candidate.empty())
    {
        candidate = L"item";
    }

    if (! parent || ! FindChild(parent, candidate))
    {
        return candidate;
    }

    for (unsigned long index = 1; index < 10000; ++index)
    {
        std::wstring suffix = L" (";
        suffix.append(std::to_wstring(index));
        suffix.push_back(L')');

        std::wstring trimmed = candidate;
        if (trimmed.size() + suffix.size() > kMaxNameLength)
        {
            if (kMaxNameLength > suffix.size())
            {
                TrimToLength(trimmed, kMaxNameLength - suffix.size());
            }
        }

        std::wstring withSuffix = trimmed;
        withSuffix.append(suffix);
        if (! FindChild(parent, withSuffix))
        {
            return withSuffix;
        }
    }

    return candidate;
}

std::wstring FileSystemDummy::MakeRandomName(std::mt19937& rng, bool isDirectory)
{
    std::wstring name = MakeRandomBaseName(rng);
    if (! IsNameValid(name))
    {
        name = L"item";
    }

    if (! isDirectory)
    {
        const unsigned long extensionIndex = RandomRange(rng, 0, ArrayCount(kExtensions) - 1u);
        const std::wstring_view extension  = kExtensions[extensionIndex];
        if (name.size() + extension.size() > kMaxNameLength)
        {
            if (kMaxNameLength > extension.size())
            {
                TrimToLength(name, kMaxNameLength - extension.size());
            }
        }
        name.append(extension.data(), extension.size());
    }

    return name;
}

std::wstring FileSystemDummy::MakeRandomBaseName(std::mt19937& rng)
{
    const unsigned long style  = RandomRange(rng, 0, 4);
    unsigned long segmentCount = 1;
    if (style == 1)
    {
        segmentCount = 2;
    }
    else if (style == 2)
    {
        segmentCount = 3;
    }
    else if (style == 3)
    {
        segmentCount = 4;
    }

    std::wstring name;
    for (unsigned long index = 0; index < segmentCount; ++index)
    {
        std::wstring_view segment;
        const unsigned long pick = RandomRange(rng, 0, 99);
        if (pick < 40)
        {
            const unsigned long wordIndex = RandomRange(rng, 0, ArrayCount(kWordSegments) - 1u);
            segment                       = kWordSegments[wordIndex];
        }
        else if (pick < 55)
        {
            const unsigned long euroIndex = RandomRange(rng, 0, ArrayCount(kEuroSegments) - 1u);
            segment                       = kEuroSegments[euroIndex];
        }
        else if (pick < 65)
        {
            const unsigned long japaneseIndex = RandomRange(rng, 0, ArrayCount(kJapaneseSegments) - 1u);
            segment                           = kJapaneseSegments[japaneseIndex];
        }
        else if (pick < 73)
        {
            const unsigned long arabicIndex = RandomRange(rng, 0, ArrayCount(kArabicSegments) - 1u);
            segment                         = kArabicSegments[arabicIndex];
        }
        else if (pick < 81)
        {
            const unsigned long thaiIndex = RandomRange(rng, 0, ArrayCount(kThaiSegments) - 1u);
            segment                       = kThaiSegments[thaiIndex];
        }
        else if (pick < 89)
        {
            const unsigned long koreanIndex = RandomRange(rng, 0, ArrayCount(kKoreanSegments) - 1u);
            segment                         = kKoreanSegments[koreanIndex];
        }
        else if (pick < 95)
        {
            const unsigned long longIndex = RandomRange(rng, 0, ArrayCount(kLongSegments) - 1u);
            segment                       = kLongSegments[longIndex];
        }
        else
        {
            const unsigned long emojiIndex = RandomRange(rng, 0, ArrayCount(kEmojiSegments) - 1u);
            segment                        = kEmojiSegments[emojiIndex];
        }

        if (! name.empty())
        {
            const unsigned long separatorIndex = RandomRange(rng, 0, ArrayCount(kSeparators) - 1u);
            name.push_back(kSeparators[separatorIndex]);
        }

        if (name.size() + segment.size() > kMaxNameLength)
        {
            break;
        }

        name.append(segment.data(), segment.size());
    }

    if (name.empty())
    {
        name = L"item";
    }

    if (RandomChance(rng, 1, 3))
    {
        const unsigned long suffix    = RandomRange(rng, 1, 9999);
        const std::wstring suffixText = std::to_wstring(suffix);
        if (name.size() + suffixText.size() + 1 <= kMaxNameLength)
        {
            name.push_back(L' ');
            name.append(suffixText);
        }
    }

    if (style == 4 && name.size() < 32)
    {
        const std::wstring_view pad = L"long";
        while (name.size() + pad.size() + 1 <= kMaxNameLength && name.size() < 48)
        {
            name.push_back(L'_');
            name.append(pad.data(), pad.size());
        }
    }

    if (RandomChance(rng, 1, 4))
    {
        const unsigned long emojiIndex = RandomRange(rng, 0, ArrayCount(kEmojiSegments) - 1u);
        const std::wstring_view emoji  = kEmojiSegments[emojiIndex];
        if (name.size() + emoji.size() + 1 <= kMaxNameLength)
        {
            name.push_back(L' ');
            name.append(emoji.data(), emoji.size());
        }
    }

    TrimToLength(name, kMaxNameLength);
    return name;
}

void FileSystemDummy::TouchNode(DummyNode& node) noexcept
{
    const __int64 now   = GetNowFileTime();
    node.lastWriteTime  = now;
    node.changeTime     = now;
    node.lastAccessTime = now;
}

void FileSystemDummy::TouchParent(DummyNode* parent) noexcept
{
    if (parent)
    {
        TouchNode(*parent);
    }
}

uint64_t FileSystemDummy::ComputeNodeBytes(const DummyNode& node) const noexcept
{
    if (! node.isDirectory)
    {
        return node.sizeBytes;
    }

    if (! node.childrenGenerated)
    {
        return 0;
    }

    uint64_t total = 0;
    for (const auto& child : node.children)
    {
        total += ComputeNodeBytes(*child);
    }
    return total;
}

bool FileSystemDummy::IsAncestor(const DummyNode& node, const DummyNode& possibleDescendant) const noexcept
{
    const DummyNode* current = &possibleDescendant;
    while (current)
    {
        if (current == &node)
        {
            return true;
        }
        current = current->parent;
    }
    return false;
}

std::unique_ptr<FileSystemDummy::DummyNode> FileSystemDummy::CloneNode(const DummyNode& source)
{
    auto clone                 = std::make_unique<DummyNode>();
    clone->name                = source.name;
    clone->isDirectory         = source.isDirectory;
    clone->attributes          = source.attributes;
    clone->sizeBytes           = source.sizeBytes;
    clone->creationTime        = source.creationTime;
    clone->lastAccessTime      = source.lastAccessTime;
    clone->lastWriteTime       = source.lastWriteTime;
    clone->changeTime          = source.changeTime;
    clone->generationSeed      = source.generationSeed;
    clone->plannedChildCount   = source.plannedChildCount;
    clone->childrenGenerated   = source.childrenGenerated;
    clone->materializedContent = source.materializedContent;

    if (source.childrenGenerated)
    {
        clone->children.reserve(source.children.size());
        for (const auto& child : source.children)
        {
            auto childClone    = CloneNode(*child);
            childClone->parent = clone.get();
            clone->children.push_back(std::move(childClone));
        }
        clone->plannedChildCount = static_cast<unsigned long>(clone->children.size());
    }

    return clone;
}

HRESULT FileSystemDummy::CreateDirectoryClone(
    const DummyNode& sourceDirectory, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, DummyNode** outDirectory)
{
    if (outDirectory)
    {
        *outDirectory = nullptr;
    }

    if (! IsNameValid(destinationName))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    if (! sourceDirectory.isDirectory)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    EnsureChildrenGenerated(destinationParent);
    DummyNode* existing = FindChild(&destinationParent, destinationName);
    if (existing != nullptr)
    {
        if (! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if ((existing->attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ExtractChild(&destinationParent, existing);
    }

    auto clone               = std::make_unique<DummyNode>();
    clone->name              = std::wstring(destinationName);
    clone->isDirectory       = true;
    clone->attributes        = sourceDirectory.attributes | FILE_ATTRIBUTE_DIRECTORY;
    clone->sizeBytes         = 0;
    clone->creationTime      = sourceDirectory.creationTime;
    clone->lastAccessTime    = sourceDirectory.lastAccessTime;
    clone->lastWriteTime     = sourceDirectory.lastWriteTime;
    clone->changeTime        = sourceDirectory.changeTime;
    clone->generationSeed    = sourceDirectory.generationSeed;
    clone->plannedChildCount = 0;
    clone->childrenGenerated = true;

    DummyNode* added = AddChild(&destinationParent, std::move(clone));
    if (! added)
    {
        return E_FAIL;
    }

    TouchNode(*added);
    if (outDirectory)
    {
        *outDirectory = added;
    }

    return S_OK;
}

HRESULT
FileSystemDummy::CopyNode(DummyNode& source, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, uint64_t* outBytes)
{
    if (! IsNameValid(destinationName))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    if (source.isDirectory && ! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
    {
        return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
    }

    if (source.isDirectory)
    {
        constexpr unsigned long kMaterializeDepth = 1;
        auto materialize                          = [&](auto&& recurse, DummyNode& directory, unsigned long remainingDepth) -> void
        {
            EnsureChildrenGenerated(directory);
            if (remainingDepth == 0)
            {
                return;
            }

            for (const auto& child : directory.children)
            {
                if (child && child->isDirectory)
                {
                    recurse(recurse, *child, remainingDepth - 1u);
                }
            }
        };

        materialize(materialize, source, kMaterializeDepth);
    }

    EnsureChildrenGenerated(destinationParent);
    DummyNode* existing = FindChild(&destinationParent, destinationName);
    if (existing)
    {
        if (! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if ((existing->attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ExtractChild(&destinationParent, existing);
    }

    auto clone       = CloneNode(source);
    clone->name      = std::wstring(destinationName);
    DummyNode* added = AddChild(&destinationParent, std::move(clone));
    if (! added)
    {
        return E_FAIL;
    }

    TouchNode(*added);
    if (outBytes)
    {
        *outBytes = ComputeNodeBytes(*added);
    }

    return S_OK;
}

HRESULT
FileSystemDummy::MoveNode(DummyNode& source, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, uint64_t* outBytes)
{
    if (! IsNameValid(destinationName))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    DummyNode* sourceParent = source.parent;
    if (! sourceParent)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (&destinationParent == sourceParent && EqualsNoCase(source.name, destinationName))
    {
        source.name = std::wstring(destinationName);
        TouchNode(source);
        if (outBytes)
        {
            *outBytes = ComputeNodeBytes(source);
        }
        return S_OK;
    }

    if (source.isDirectory)
    {
        constexpr unsigned long kMaterializeDepth = 1;
        auto materialize                          = [&](auto&& recurse, DummyNode& directory, unsigned long remainingDepth) -> void
        {
            EnsureChildrenGenerated(directory);
            if (remainingDepth == 0)
            {
                return;
            }

            for (const auto& child : directory.children)
            {
                if (child && child->isDirectory)
                {
                    recurse(recurse, *child, remainingDepth - 1u);
                }
            }
        };

        materialize(materialize, source, kMaterializeDepth);
    }

    EnsureChildrenGenerated(destinationParent);
    DummyNode* existing = FindChild(&destinationParent, destinationName);
    if (existing && existing != &source)
    {
        if (! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if ((existing->attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ExtractChild(&destinationParent, existing);
    }

    auto moved = ExtractChild(sourceParent, &source);
    if (! moved)
    {
        return E_FAIL;
    }

    moved->name      = std::wstring(destinationName);
    DummyNode* added = AddChild(&destinationParent, std::move(moved));
    if (! added)
    {
        return E_FAIL;
    }

    TouchNode(*added);
    if (outBytes)
    {
        *outBytes = ComputeNodeBytes(*added);
    }

    return S_OK;
}

HRESULT FileSystemDummy::DeleteNode(DummyNode& target, FileSystemFlags flags)
{
    DummyNode* parent = target.parent;
    if (! parent)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if ((target.attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (target.isDirectory && ! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
    {
        const bool hasChildren = target.childrenGenerated ? ! target.children.empty() : target.plannedChildCount > 0;
        if (hasChildren)
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }
    }

    auto removed = ExtractChild(parent, &target);
    if (! removed)
    {
        return E_FAIL;
    }

    return S_OK;
}

HRESULT FileSystemDummy::CommitFileWriter(const std::filesystem::path& normalizedPath,
                                          FileSystemFlags flags,
                                          const std::shared_ptr<std::vector<std::byte>>& buffer) noexcept
{
    if (buffer == nullptr)
    {
        return E_POINTER;
    }

    const std::filesystem::path parentPath = normalizedPath.parent_path();
    const std::wstring name                = normalizedPath.filename().wstring();
    if (name.empty() || ! IsNameValid(name))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    const std::wstring parentText = parentPath.wstring();
    const __int64 now             = GetNowFileTime();

    {
        std::scoped_lock lock(_mutex);

        DummyNode* parent = nullptr;
        HRESULT hr        = ResolvePath(parentPath, &parent, false, true);
        if (FAILED(hr))
        {
            return hr;
        }

        EnsureChildrenGenerated(*parent);

        DummyNode* existing = FindChild(parent, name);
        if (existing != nullptr)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }

            if ((existing->attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            ExtractChild(parent, existing);
        }

        auto node                 = std::make_unique<DummyNode>();
        node->name                = name;
        node->isDirectory         = false;
        node->attributes          = FILE_ATTRIBUTE_ARCHIVE;
        node->sizeBytes           = static_cast<uint64_t>(buffer->size());
        node->creationTime        = now;
        node->lastAccessTime      = now;
        node->lastWriteTime       = now;
        node->changeTime          = now;
        node->generationSeed      = CombineSeed(parent->generationSeed, name);
        node->plannedChildCount   = 0;
        node->childrenGenerated   = true;
        node->materializedContent = buffer;

        if (! AddChild(parent, std::move(node)))
        {
            return E_FAIL;
        }
    }

    NotifyDirectoryWatchers(parentText, name, FILESYSTEM_DIR_CHANGE_ADDED);
    return S_OK;
}

void FileSystemDummy::SimulateLatency(unsigned long itemCount) const noexcept
{
    if (itemCount == 0)
    {
        return;
    }

    unsigned long latencyMilliseconds = 0;
    {
        std::scoped_lock lock(_mutex);
        latencyMilliseconds = _latencyMilliseconds;
    }

    if (latencyMilliseconds == 0)
    {
        return;
    }

    const std::uint64_t totalMs = static_cast<std::uint64_t>(latencyMilliseconds) * static_cast<std::uint64_t>(itemCount);

    constexpr std::uint64_t kMaxSleepMs = static_cast<std::uint64_t>(std::numeric_limits<DWORD>::max());
    const DWORD sleepMs                 = static_cast<DWORD>(std::min(totalMs, kMaxSleepMs));
    ::Sleep(sleepMs);
}

void FileSystemDummy::NotifyDirectoryWatchers(std::wstring_view watchedPath, std::wstring_view relativePath, FileSystemDirectoryChangeAction action) noexcept
{
    if (watchedPath.empty() || relativePath.empty())
    {
        return;
    }

    if (watchedPath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
    {
        return;
    }

    if (relativePath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
    {
        return;
    }

    std::vector<std::shared_ptr<DirectoryWatchRegistration>> watchers;
    {
        std::lock_guard lock(_watchMutex);
        watchers.reserve(_directoryWatches.size());

        for (const auto& entry : _directoryWatches)
        {
            if (! entry)
            {
                continue;
            }

            if (! entry->active.load(std::memory_order_acquire))
            {
                continue;
            }

            if (! EqualsNoCase(entry->watchedPath, watchedPath))
            {
                continue;
            }

            entry->inFlight.fetch_add(1u, std::memory_order_acq_rel);
            watchers.push_back(entry);
        }
    }

    if (watchers.empty())
    {
        return;
    }

    FileSystemDirectoryChange change{};
    change.action           = action;
    change.relativePath     = relativePath.data();
    change.relativePathSize = static_cast<unsigned long>(relativePath.size() * sizeof(wchar_t));

    for (const auto& watcher : watchers)
    {
        if (! watcher)
        {
            continue;
        }

        FileSystemDirectoryChangeNotification notification{};
        notification.watchedPath     = watcher->watchedPath.c_str();
        notification.watchedPathSize = static_cast<unsigned long>(watcher->watchedPath.size() * sizeof(wchar_t));
        notification.changes         = &change;
        notification.changeCount     = 1;
        notification.overflow        = FALSE;

        if (watcher->active.load(std::memory_order_acquire) && watcher->callback)
        {
            DirectoryWatchCallbackScope callbackScope(watcher.get());
            watcher->callback->FileSystemDirectoryChanged(&notification, watcher->cookie);
        }

        const uint32_t remaining = watcher->inFlight.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u || ! watcher->active.load(std::memory_order_acquire))
        {
            _watchCv.notify_all();
        }
    }
}

void FileSystemDummy::NotifyDirectoryWatchers(std::wstring_view watchedPath, std::wstring_view oldRelativePath, std::wstring_view newRelativePath) noexcept
{
    if (watchedPath.empty() || oldRelativePath.empty() || newRelativePath.empty())
    {
        return;
    }

    if (watchedPath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
    {
        return;
    }

    if (oldRelativePath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
    {
        return;
    }

    if (newRelativePath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
    {
        return;
    }

    std::vector<std::shared_ptr<DirectoryWatchRegistration>> watchers;
    {
        std::lock_guard lock(_watchMutex);
        watchers.reserve(_directoryWatches.size());

        for (const auto& entry : _directoryWatches)
        {
            if (! entry)
            {
                continue;
            }

            if (! entry->active.load(std::memory_order_acquire))
            {
                continue;
            }

            if (! EqualsNoCase(entry->watchedPath, watchedPath))
            {
                continue;
            }

            entry->inFlight.fetch_add(1u, std::memory_order_acq_rel);
            watchers.push_back(entry);
        }
    }

    if (watchers.empty())
    {
        return;
    }

    FileSystemDirectoryChange changes[2]{};
    changes[0].action           = FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME;
    changes[0].relativePath     = oldRelativePath.data();
    changes[0].relativePathSize = static_cast<unsigned long>(oldRelativePath.size() * sizeof(wchar_t));
    changes[1].action           = FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME;
    changes[1].relativePath     = newRelativePath.data();
    changes[1].relativePathSize = static_cast<unsigned long>(newRelativePath.size() * sizeof(wchar_t));

    for (const auto& watcher : watchers)
    {
        if (! watcher)
        {
            continue;
        }

        FileSystemDirectoryChangeNotification notification{};
        notification.watchedPath     = watcher->watchedPath.c_str();
        notification.watchedPathSize = static_cast<unsigned long>(watcher->watchedPath.size() * sizeof(wchar_t));
        notification.changes         = changes;
        notification.changeCount     = 2;
        notification.overflow        = FALSE;

        if (watcher->active.load(std::memory_order_acquire) && watcher->callback)
        {
            DirectoryWatchCallbackScope callbackScope(watcher.get());
            watcher->callback->FileSystemDirectoryChanged(&notification, watcher->cookie);
        }

        const uint32_t remaining = watcher->inFlight.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u || ! watcher->active.load(std::memory_order_acquire))
        {
            _watchCv.notify_all();
        }
    }
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (! ppFilesInformation)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    DummyNode* node                        = nullptr;
    std::vector<DummyEntry> entries;
    unsigned long count = 0;
    std::vector<std::byte> buffer;
    unsigned long usedBytes = 0;

    {
        std::scoped_lock lock(_mutex);
        HRESULT hr = ResolvePath(normalized, &node, false, true);
        if (FAILED(hr))
        {
            return hr;
        }

        EnsureChildrenGenerated(*node);
        count = static_cast<unsigned long>(node->children.size());
        entries.reserve(node->children.size());
        for (const auto& child : node->children)
        {
            DummyEntry entry{};
            entry.name           = child->name;
            entry.attributes     = child->attributes;
            entry.sizeBytes      = child->sizeBytes;
            entry.creationTime   = child->creationTime;
            entry.lastAccessTime = child->lastAccessTime;
            entry.lastWriteTime  = child->lastWriteTime;
            entry.changeTime     = child->changeTime;
            entries.push_back(std::move(entry));
        }
    }

    SimulateLatency(count);

    HRESULT hr = BuildFileInfoBuffer(entries, &buffer, &usedBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    auto* info = new (std::nothrow) DummyFilesInformation(std::move(buffer), count, usedBytes);
    if (! info)
    {
        return E_OUTOFMEMORY;
    }

    *ppFilesInformation = info;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (fileAttributes == nullptr)
    {
        return E_POINTER;
    }

    *fileAttributes = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    DummyNode* node                        = nullptr;

    {
        std::scoped_lock lock(_mutex);
        const HRESULT hr = ResolvePath(normalized, &node, false, false);
        if (FAILED(hr))
        {
            return hr;
        }
        if (node == nullptr)
        {
            return E_FAIL;
        }

        *fileAttributes = node->attributes;
    }

    SimulateLatency(1);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (reader == nullptr)
    {
        return E_POINTER;
    }

    *reader = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);

    DummyFileSnapshot snapshot{};

    {
        std::scoped_lock lock(_mutex);

        DummyNode* node  = nullptr;
        const HRESULT hr = ResolvePath(normalized, &node, false, false);
        if (FAILED(hr))
        {
            return hr;
        }
        if (node == nullptr)
        {
            return E_FAIL;
        }
        if (node->isDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        snapshot.name                = node->name;
        snapshot.attributes          = node->attributes;
        snapshot.sizeBytes           = node->sizeBytes;
        snapshot.creationTime        = node->creationTime;
        snapshot.generationSeed      = node->generationSeed;
        snapshot.materializedContent = node->materializedContent;
    }

    SimulateLatency(1);

    if (snapshot.materializedContent)
    {
        auto* impl = new (std::nothrow) DummySharedBufferFileReader(std::move(snapshot.materializedContent));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        *reader = impl;
        return S_OK;
    }

    const DummyFileKind fileKind    = GetDummyFileKind(snapshot.name);
    const std::uint64_t seed        = ComputeDummyFileContentSeed(snapshot);
    const std::uint64_t contentSeed = Mix64(seed + static_cast<std::uint64_t>(fileKind));

    IFileReader* created = nullptr;

    if (fileKind == DummyFileKind::Png)
    {
        std::vector<std::byte> png = GenerateDummyPng(contentSeed, snapshot.sizeBytes);
        if (! png.empty())
        {
            created = new (std::nothrow) DummyBufferFileReader(std::move(png));
        }
    }
    else if (fileKind == DummyFileKind::Jpeg)
    {
        std::vector<std::byte> jpeg = GenerateDummyJpeg(contentSeed, snapshot.sizeBytes);
        if (! jpeg.empty())
        {
            created = new (std::nothrow) DummyBufferFileReader(std::move(jpeg));
        }
    }

    if (! created)
    {
        const std::uint64_t fillSeed = Mix64(contentSeed ^ 0xD00DFEEDu);

        if (fileKind == DummyFileKind::Binary || fileKind == DummyFileKind::Zip || fileKind == DummyFileKind::Png || fileKind == DummyFileKind::Jpeg)
        {
            created = new (std::nothrow) DummyGeneratedFileReader({}, {}, snapshot.sizeBytes, fillSeed, DummyFillKind::Binary);
        }
        else
        {
            DummyTextTemplate templ = BuildDummyTextTemplate(fileKind, snapshot, seed);
            created = new (std::nothrow) DummyGeneratedFileReader(std::move(templ.prefix), std::move(templ.suffix), templ.bodyBytes, fillSeed, templ.fillKind);
        }
    }

    if (! created)
    {
        return E_OUTOFMEMORY;
    }

    *reader = created;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (writer == nullptr)
    {
        return E_POINTER;
    }

    *writer = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    const std::filesystem::path parentPath = normalized.parent_path();
    const std::wstring name                = normalized.filename().wstring();
    if (name.empty() || ! IsNameValid(name))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    {
        std::scoped_lock lock(_mutex);

        DummyNode* parent = nullptr;
        HRESULT hr        = ResolvePath(parentPath, &parent, false, true);
        if (FAILED(hr))
        {
            return hr;
        }

        EnsureChildrenGenerated(*parent);

        DummyNode* existing = FindChild(parent, name);
        if (existing != nullptr)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }

            if ((existing->attributes & FILE_ATTRIBUTE_READONLY) != 0 && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY))
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            if (existing->isDirectory)
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
        }
    }

    auto* created = new (std::nothrow) DummyFileWriter(*this, normalized, flags);
    if (! created)
    {
        return E_OUTOFMEMORY;
    }

    *writer = created;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    *info = {};

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    DummyNode* node                        = nullptr;

    {
        std::scoped_lock lock(_mutex);
        const HRESULT hr = ResolvePath(normalized, &node, false, false);
        if (FAILED(hr))
        {
            return hr;
        }
        if (node == nullptr)
        {
            return E_FAIL;
        }

        info->creationTime   = node->creationTime;
        info->lastAccessTime = node->lastAccessTime;
        info->lastWriteTime  = node->lastWriteTime;
        info->attributes     = node->attributes;
    }

    SimulateLatency(1);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);

    {
        std::scoped_lock lock(_mutex);
        DummyNode* node  = nullptr;
        const HRESULT hr = ResolvePath(normalized, &node, false, false);
        if (FAILED(hr))
        {
            return hr;
        }
        if (node == nullptr)
        {
            return E_FAIL;
        }

        node->creationTime   = info->creationTime;
        node->lastAccessTime = info->lastAccessTime;
        node->lastWriteTime  = info->lastWriteTime;

        DWORD attrs = info->attributes;
        if (node->isDirectory)
        {
            attrs |= FILE_ATTRIBUTE_DIRECTORY;
        }
        else
        {
            attrs &= ~FILE_ATTRIBUTE_DIRECTORY;
            if (attrs == 0)
            {
                attrs = FILE_ATTRIBUTE_NORMAL;
            }
        }
        node->attributes = attrs;
        node->changeTime = GetNowFileTime();
    }

    SimulateLatency(1);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = kCapabilitiesJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);

    DummyEntry entry{};
    bool isDirectory = false;

    {
        std::scoped_lock lock(_mutex);

        DummyNode* node  = nullptr;
        const HRESULT hr = ResolvePath(normalized, &node, false, false);
        if (FAILED(hr))
        {
            return hr;
        }

        if (node == nullptr)
        {
            return E_FAIL;
        }

        isDirectory          = node->isDirectory;
        entry.name           = node->name;
        entry.attributes     = node->attributes;
        entry.sizeBytes      = node->sizeBytes;
        entry.creationTime   = node->creationTime;
        entry.lastAccessTime = node->lastAccessTime;
        entry.lastWriteTime  = node->lastWriteTime;
        entry.changeTime     = node->changeTime;
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);

    yyjson_mut_obj_add_int(doc, root, "version", 1);
    yyjson_mut_obj_add_str(doc, root, "title", "properties");

    yyjson_mut_val* sections = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "sections", sections);

    yyjson_mut_val* general = yyjson_mut_obj(doc);
    yyjson_mut_arr_add_val(sections, general);
    yyjson_mut_obj_add_str(doc, general, "title", "general");

    yyjson_mut_val* fields = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, general, "fields", fields);

    auto addField = [&](const char* key, const std::string& value)
    {
        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(fields, field);
    };

    addField("name", Utf8FromUtf16(entry.name));
    addField("path", Utf8FromUtf16(normalized.wstring()));
    addField("type", isDirectory ? std::string("directory") : std::string("file"));
    if (! isDirectory)
    {
        addField("sizeBytes", std::format("{}", entry.sizeBytes));
    }

    const char* written = yyjson_mut_write(doc, YYJSON_WRITE_NOFLAG, nullptr);
    if (! written)
    {
        return E_OUTOFMEMORY;
    }
    auto freeWritten = wil::scope_exit([&] { free(const_cast<char*>(written)); });

    {
        std::scoped_lock lock(_propertiesMutex);
        _lastPropertiesJson.assign(written);
        *jsonUtf8 = _lastPropertiesJson.c_str();
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::CreateDirectory(const wchar_t* path) noexcept
{
    if (path == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    const std::filesystem::path parentPath = normalized.parent_path();
    const std::wstring name                = normalized.filename().wstring();
    if (name.empty() || ! IsNameValid(name))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    const std::wstring parentText = parentPath.wstring();
    const __int64 now             = GetNowFileTime();

    {
        std::scoped_lock lock(_mutex);

        DummyNode* parent = nullptr;
        HRESULT hr        = ResolvePath(parentPath, &parent, false, true);
        if (FAILED(hr))
        {
            return hr;
        }

        EnsureChildrenGenerated(*parent);

        DummyNode* existing = FindChild(parent, name);
        if (existing != nullptr)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const std::uint64_t childSeed = CombineSeed(parent->generationSeed, name);

        auto node               = std::make_unique<DummyNode>();
        node->name              = name;
        node->isDirectory       = true;
        node->attributes        = FILE_ATTRIBUTE_DIRECTORY;
        node->sizeBytes         = 0;
        node->creationTime      = now;
        node->lastAccessTime    = now;
        node->lastWriteTime     = now;
        node->changeTime        = now;
        node->generationSeed    = childSeed;
        node->plannedChildCount = 0;
        node->childrenGenerated = true;

        if (! AddChild(parent, std::move(node)))
        {
            return E_FAIL;
        }
    }

    NotifyDirectoryWatchers(parentText, name, FILESYSTEM_DIR_CHANGE_ADDED);
    SimulateLatency(1);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (path == nullptr || result == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    const std::filesystem::path normalized           = NormalizePath(path);
    const bool recursive                             = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
    constexpr unsigned long kProgressIntervalEntries = 100;
    constexpr ULONGLONG kProgressIntervalMs          = 200;

    uint64_t scannedEntries    = 0;
    ULONGLONG lastProgressTime = ::GetTickCount64();

    auto maybeReportProgress = [&](const wchar_t* currentPath) -> bool
    {
        if (callback == nullptr)
        {
            return true;
        }

        const bool entryThreshold = (scannedEntries % kProgressIntervalEntries) == 0;
        const ULONGLONG now       = ::GetTickCount64();
        const bool timeThreshold  = (now - lastProgressTime) >= kProgressIntervalMs;

        if (entryThreshold || timeThreshold)
        {
            lastProgressTime = now;
            callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, currentPath, cookie);

            BOOL cancel = FALSE;
            callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return false;
            }
        }
        return true;
    };

    bool rootIsFile       = false;
    uint64_t rootFileSize = 0;

    // Validate root path exists and classify directory/file root.
    {
        std::scoped_lock lock(_mutex);

        DummyNode* rootNode = nullptr;
        const HRESULT hr    = ResolvePath(normalized, &rootNode, false, false);
        if (FAILED(hr))
        {
            result->status = hr;
            return hr;
        }

        rootIsFile   = ! rootNode->isDirectory;
        rootFileSize = rootNode->sizeBytes;
    }

    if (rootIsFile)
    {
        scannedEntries     = 1;
        result->totalBytes = rootFileSize;
        result->fileCount  = 1;

        if (! maybeReportProgress(normalized.c_str()))
        {
            return result->status;
        }

        if (callback != nullptr)
        {
            callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
        }

        return result->status;
    }

    if (! maybeReportProgress(normalized.c_str()))
    {
        return result->status;
    }

    struct ChildSnapshot
    {
        std::wstring name;
        bool isDirectory   = false;
        uint64_t sizeBytes = 0;
    };

    std::vector<std::filesystem::path> pending;
    pending.emplace_back(normalized);

    while (! pending.empty())
    {
        std::filesystem::path currentPath = std::move(pending.back());
        pending.pop_back();

        std::vector<ChildSnapshot> children;
        unsigned long childCount = 0;

        {
            std::scoped_lock lock(_mutex);

            DummyNode* currentNode = nullptr;
            const HRESULT hr       = ResolvePath(currentPath, &currentNode, false, true);
            if (FAILED(hr))
            {
                if (hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && hr != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) &&
                    hr != HRESULT_FROM_WIN32(ERROR_DIRECTORY))
                {
                    if (SUCCEEDED(result->status))
                    {
                        result->status = hr;
                    }
                }
                continue;
            }

            EnsureChildrenGenerated(*currentNode);

            const size_t childCountSize              = currentNode->children.size();
            constexpr unsigned long kMaxUnsignedLong = std::numeric_limits<unsigned long>::max();
            if (childCountSize > static_cast<size_t>(kMaxUnsignedLong))
            {
                childCount = kMaxUnsignedLong;
            }
            else
            {
                childCount = static_cast<unsigned long>(childCountSize);
            }

            children.reserve(childCountSize);
            for (const auto& child : currentNode->children)
            {
                if (! child)
                {
                    continue;
                }

                ChildSnapshot snap{};
                snap.name        = child->name;
                snap.isDirectory = child->isDirectory;
                snap.sizeBytes   = child->sizeBytes;
                children.emplace_back(std::move(snap));
            }
        }

        for (const auto& child : children)
        {
            ++scannedEntries;

            // Directory size scanning enumerates directory entries; honor the configured latency per entry to keep
            // pre-calculation behavior consistent with other dummy filesystem operations.
            SimulateLatency(1);

            if (child.isDirectory)
            {
                ++result->directoryCount;

                if (recursive)
                {
                    pending.emplace_back(currentPath / child.name);
                }
            }
            else
            {
                ++result->fileCount;
                result->totalBytes += child.sizeBytes;
            }

            if (! maybeReportProgress(currentPath.c_str()))
            {
                return result->status;
            }
        }

        // Artificial per-entry latency is performed outside the in-memory filesystem lock so parallel tasks can proceed.
        if (_latencyMilliseconds > 0 && childCount > 0)
        {
            const std::uint64_t totalMs64 = static_cast<std::uint64_t>(_latencyMilliseconds) * static_cast<std::uint64_t>(childCount);

            constexpr std::uint64_t kMaxSleepMs = static_cast<std::uint64_t>(std::numeric_limits<DWORD>::max());
            DWORD remainingMs                   = static_cast<DWORD>(std::min(totalMs64, kMaxSleepMs));

            while (remainingMs > 0)
            {
                if (callback != nullptr)
                {
                    BOOL cancel = FALSE;
                    callback->DirectorySizeShouldCancel(&cancel, cookie);
                    if (cancel)
                    {
                        result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        return result->status;
                    }
                }

                constexpr DWORD kChunkMs = 200;
                const DWORD chunkMs      = std::min(remainingMs, kChunkMs);
                ::Sleep(chunkMs);
                remainingMs -= chunkMs;

                if (! maybeReportProgress(currentPath.c_str()))
                {
                    return result->status;
                }
            }
        }
    }

    // Final progress report.
    if (callback != nullptr)
    {
        callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
    }

    return result->status;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept
{
    if (! path || ! callback)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    const std::wstring watchedPathText     = normalized.wstring();

    {
        std::scoped_lock lock(_mutex);
        DummyNode* directory = nullptr;
        const HRESULT hr     = ResolvePath(normalized, &directory, false, true);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    auto watch         = std::make_shared<DirectoryWatchRegistration>();
    watch->owner       = this;
    watch->watchedPath = watchedPathText;
    watch->callback    = callback;
    watch->cookie      = cookie;
    watch->inFlight.store(0u, std::memory_order_release);
    watch->active.store(true, std::memory_order_release);

    {
        std::lock_guard lock(_watchMutex);
        for (const auto& existing : _directoryWatches)
        {
            if (! existing)
            {
                continue;
            }

            if (! existing->active.load(std::memory_order_acquire))
            {
                continue;
            }

            if (existing->owner == this && EqualsNoCase(existing->watchedPath, watchedPathText))
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
        }

        _directoryWatches.push_back(std::move(watch));
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::UnwatchDirectory(const wchar_t* path) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = NormalizePath(path);
    const std::wstring watchedPathText     = normalized.wstring();

    std::shared_ptr<DirectoryWatchRegistration> removed;

    {
        std::unique_lock lock(_watchMutex);

        auto it = _directoryWatches.begin();
        for (; it != _directoryWatches.end(); ++it)
        {
            const auto& entry = *it;
            if (! entry)
            {
                continue;
            }

            if (entry->owner != this)
            {
                continue;
            }

            if (EqualsNoCase(entry->watchedPath, watchedPathText))
            {
                removed = entry;
                _directoryWatches.erase(it);
                break;
            }
        }

        if (! removed)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        removed->active.store(false, std::memory_order_release);
        const bool reentrant           = static_cast<const void*>(removed.get()) == g_activeDirectoryWatchCallback;
        const uint32_t desiredInFlight = reentrant ? 1u : 0u;
        _watchCv.wait(lock, [&] { return removed->inFlight.load(std::memory_order_acquire) <= desiredInFlight; });
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::CopyItem(const wchar_t* sourcePath,
                                                    const wchar_t* destinationPath,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, 1);

    const std::filesystem::path normalizedSource      = NormalizePath(sourcePath);
    const std::filesystem::path normalizedDestination = NormalizePath(destinationPath);
    const std::wstring sourceText                     = normalizedSource.wstring();
    const std::wstring destinationText                = normalizedDestination.wstring();
    const std::wstring destinationParentText          = normalizedDestination.parent_path().wstring();
    const std::wstring destinationLeafText            = normalizedDestination.filename().wstring();

    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t itemBytes = 0;
    HRESULT itemHr     = S_OK;

    {
        std::scoped_lock lock(_mutex);
        DummyNode* sourceNode = nullptr;
        itemHr                = ResolvePath(normalizedSource, &sourceNode, false, false);
        if (SUCCEEDED(itemHr))
        {
            const std::filesystem::path destinationParentPath = normalizedDestination.parent_path();
            const std::wstring destinationName                = normalizedDestination.filename().wstring();
            if (destinationName.empty())
            {
                itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            else
            {
                DummyNode* destinationParent = nullptr;
                itemHr                       = ResolvePath(destinationParentPath, &destinationParent, false, true);
                if (SUCCEEDED(itemHr))
                {
                    itemHr = CopyNode(*sourceNode, *destinationParent, destinationName, flags, &itemBytes);
                }
            }
        }
    }

    if (SUCCEEDED(itemHr))
    {
        NotifyDirectoryWatchers(destinationParentText, destinationLeafText, FILESYSTEM_DIR_CHANGE_ADDED);
    }

    context.completedBytes = 0;

    hr = SetProgressPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    const uint64_t baseCompletedBytes         = 0;
    const uint64_t virtualLimitBytesPerSecond = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    unsigned long latencyMilliseconds         = 0;
    std::uint64_t effectiveSeed               = 0;
    {
        std::scoped_lock lock(_mutex);
        latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed       = _effectiveSeed;
    }
    context.virtualLimitBytesPerSecond = virtualLimitBytesPerSecond;
    context.latencyMilliseconds        = latencyMilliseconds;
    context.throughputSeed             = CombineSeed(effectiveSeed, sourceText);
    context.throughputSeed             = CombineSeed(context.throughputSeed, destinationText);

    if (SUCCEEDED(itemHr))
    {
        hr = ReportThrottledByteProgress(context, itemBytes, baseCompletedBytes, virtualLimitBytesPerSecond);
    }
    else
    {
        hr = ReportProgress(context, itemBytes, 0);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetItemPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (FAILED(itemHr))
    {
        return itemHr;
    }

    context.completedItems = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::MoveItem(const wchar_t* sourcePath,
                                                    const wchar_t* destinationPath,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, 1);

    const std::filesystem::path normalizedSource      = NormalizePath(sourcePath);
    const std::filesystem::path normalizedDestination = NormalizePath(destinationPath);
    const std::wstring sourceText                     = normalizedSource.wstring();
    const std::wstring destinationText                = normalizedDestination.wstring();
    const std::wstring sourceParentText               = normalizedSource.parent_path().wstring();
    const std::wstring sourceLeafText                 = normalizedSource.filename().wstring();
    const std::wstring destinationParentText          = normalizedDestination.parent_path().wstring();
    const std::wstring destinationLeafText            = normalizedDestination.filename().wstring();

    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t itemBytes = 0;
    HRESULT itemHr     = S_OK;

    {
        std::scoped_lock lock(_mutex);
        DummyNode* sourceNode = nullptr;
        itemHr                = ResolvePath(normalizedSource, &sourceNode, false, false);
        if (SUCCEEDED(itemHr))
        {
            const std::filesystem::path destinationParentPath = normalizedDestination.parent_path();
            const std::wstring destinationName                = normalizedDestination.filename().wstring();
            if (destinationName.empty())
            {
                itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            else
            {
                DummyNode* destinationParent = nullptr;
                itemHr                       = ResolvePath(destinationParentPath, &destinationParent, false, true);
                if (SUCCEEDED(itemHr))
                {
                    if (IsAncestor(*sourceNode, *destinationParent))
                    {
                        itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    else
                    {
                        itemHr = MoveNode(*sourceNode, *destinationParent, destinationName, flags, &itemBytes);
                    }
                }
            }
        }
    }

    if (SUCCEEDED(itemHr))
    {
        if (EqualsNoCase(sourceParentText, destinationParentText))
        {
            if (sourceLeafText != destinationLeafText)
            {
                NotifyDirectoryWatchers(destinationParentText, sourceLeafText, destinationLeafText);
            }
            else
            {
                NotifyDirectoryWatchers(destinationParentText, destinationLeafText, FILESYSTEM_DIR_CHANGE_MODIFIED);
            }
        }
        else
        {
            NotifyDirectoryWatchers(sourceParentText, sourceLeafText, FILESYSTEM_DIR_CHANGE_REMOVED);
            NotifyDirectoryWatchers(destinationParentText, destinationLeafText, FILESYSTEM_DIR_CHANGE_ADDED);
        }
    }

    context.completedBytes = 0;

    hr = SetProgressPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    const uint64_t baseCompletedBytes         = 0;
    const uint64_t virtualLimitBytesPerSecond = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    unsigned long latencyMilliseconds         = 0;
    std::uint64_t effectiveSeed               = 0;
    {
        std::scoped_lock lock(_mutex);
        latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed       = _effectiveSeed;
    }
    context.virtualLimitBytesPerSecond = virtualLimitBytesPerSecond;
    context.latencyMilliseconds        = latencyMilliseconds;
    context.throughputSeed             = CombineSeed(effectiveSeed, sourceText);
    context.throughputSeed             = CombineSeed(context.throughputSeed, destinationText);

    if (SUCCEEDED(itemHr))
    {
        hr = ReportThrottledByteProgress(context, itemBytes, baseCompletedBytes, virtualLimitBytesPerSecond);
    }
    else
    {
        hr = ReportProgress(context, itemBytes, 0);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetItemPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (FAILED(itemHr))
    {
        return itemHr;
    }

    context.completedItems = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE
FileSystemDummy::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 1);

    const std::filesystem::path normalized = NormalizePath(path);
    const std::wstring pathText            = normalized.wstring();
    const std::wstring parentText          = normalized.parent_path().wstring();
    const std::wstring leafText            = normalized.filename().wstring();

    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = S_OK;

    {
        std::scoped_lock lock(_mutex);
        DummyNode* node = nullptr;
        itemHr          = ResolvePath(normalized, &node, false, false);
        if (SUCCEEDED(itemHr))
        {
            itemHr = DeleteNode(*node, flags);
        }
    }

    if (SUCCEEDED(itemHr))
    {
        NotifyDirectoryWatchers(parentText, leafText, FILESYSTEM_DIR_CHANGE_REMOVED);
    }

    hr = SetProgressPaths(context, pathText.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    constexpr uint64_t kVirtualDeleteBytesPerItem = 64ull * 1024ull;
    const uint64_t virtualLimitBytesPerSecond     = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    unsigned long latencyMilliseconds             = 0;
    std::uint64_t effectiveSeed                   = 0;
    {
        std::scoped_lock lock(_mutex);
        latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed       = _effectiveSeed;
    }
    context.virtualLimitBytesPerSecond = virtualLimitBytesPerSecond;
    context.latencyMilliseconds        = latencyMilliseconds;
    context.totalBytes                 = kVirtualDeleteBytesPerItem;
    context.throughputSeed             = CombineSeed(effectiveSeed, pathText);

    if (SUCCEEDED(itemHr))
    {
        hr = ReportThrottledByteProgress(context, kVirtualDeleteBytesPerItem, 0, virtualLimitBytesPerSecond);
    }
    else
    {
        hr = ReportProgress(context, kVirtualDeleteBytesPerItem, 0);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetItemPaths(context, pathText.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (FAILED(itemHr))
    {
        return itemHr;
    }

    context.completedItems = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::RenameItem(const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      FileSystemFlags flags,
                                                      const FileSystemOptions* options,
                                                      IFileSystemCallback* callback,
                                                      void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, 1);

    const std::filesystem::path normalizedSource      = NormalizePath(sourcePath);
    const std::filesystem::path normalizedDestination = NormalizePath(destinationPath);
    const std::wstring sourceText                     = normalizedSource.wstring();
    const std::wstring destinationText                = normalizedDestination.wstring();

    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t itemBytes = 0;
    HRESULT itemHr     = S_OK;

    {
        std::scoped_lock lock(_mutex);
        DummyNode* sourceNode = nullptr;
        itemHr                = ResolvePath(normalizedSource, &sourceNode, false, false);
        if (SUCCEEDED(itemHr))
        {
            const std::filesystem::path destinationParentPath = normalizedDestination.parent_path();
            const std::wstring destinationName                = normalizedDestination.filename().wstring();
            if (destinationName.empty())
            {
                itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            else
            {
                DummyNode* destinationParent = nullptr;
                itemHr                       = ResolvePath(destinationParentPath, &destinationParent, false, true);
                if (SUCCEEDED(itemHr))
                {
                    if (IsAncestor(*sourceNode, *destinationParent))
                    {
                        itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    else
                    {
                        itemHr = MoveNode(*sourceNode, *destinationParent, destinationName, flags, &itemBytes);
                    }
                }
            }
        }
    }

    context.completedBytes = 0;

    hr = SetProgressPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgress(context, itemBytes, itemBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetItemPaths(context, sourceText.c_str(), destinationText.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (FAILED(itemHr))
    {
        return itemHr;
    }

    context.completedItems = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::CopyItems(const wchar_t* const* sourcePaths,
                                                     unsigned long count,
                                                     const wchar_t* destinationFolder,
                                                     FileSystemFlags flags,
                                                     const FileSystemOptions* options,
                                                     IFileSystemCallback* callback,
                                                     void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, count);
    std::uint64_t effectiveSeed = 0;
    {
        std::scoped_lock lock(_mutex);
        context.latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed               = _effectiveSeed;
    }

    const uint64_t virtualLimitBytesPerSecond = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    context.virtualLimitBytesPerSecond        = virtualLimitBytesPerSecond;

    const std::filesystem::path normalizedDestinationFolder = NormalizePath(destinationFolder);
    const std::wstring destinationFolderText                = normalizedDestinationFolder.wstring();

    DummyNode* destinationRoot = nullptr;
    {
        std::scoped_lock lock(_mutex);
        const HRESULT hrResolve = ResolvePath(normalizedDestinationFolder, &destinationRoot, false, true);
        if (FAILED(hrResolve))
        {
            return hrResolve;
        }
    }

    struct CopyWorkItem
    {
        DummyNode* source            = nullptr;
        DummyNode* destinationParent = nullptr;
        std::wstring sourcePathText;
        std::wstring destinationParentText;
        std::wstring destinationName;
        HRESULT preResolvedHr = S_OK;
    };

    std::vector<CopyWorkItem> stack;
    stack.reserve(static_cast<size_t>(count));

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::filesystem::path normalizedSource = NormalizePath(sourcePath);
        const std::wstring sourceText                = normalizedSource.wstring();

        CopyWorkItem item{};
        item.source                = nullptr;
        item.destinationParent     = destinationRoot;
        item.sourcePathText        = sourceText;
        item.destinationParentText = destinationFolderText;
        item.destinationName       = std::wstring(leaf);
        stack.push_back(std::move(item));
    }

    auto addToTotalItems = [&](unsigned long delta) noexcept
    {
        if (delta == 0)
        {
            return;
        }

        constexpr unsigned long maxValue = (std::numeric_limits<unsigned long>::max)();
        if (context.totalItems > maxValue - delta)
        {
            context.totalItems = maxValue;
            return;
        }

        context.totalItems += delta;
    };

    bool hadFailure = false;

    while (! stack.empty())
    {
        CopyWorkItem work = std::move(stack.back());
        stack.pop_back();

        HRESULT hrCancel = CheckCancel(context);
        if (FAILED(hrCancel))
        {
            return hrCancel;
        }

        const std::wstring destinationPathText = AppendPath(work.destinationParentText, work.destinationName);

        uint64_t itemBytes = 0;
        HRESULT itemHr     = work.preResolvedHr;

        DummyNode* source            = nullptr;
        DummyNode* destinationParent = nullptr;

        {
            std::scoped_lock lock(_mutex);

            if (SUCCEEDED(itemHr))
            {
                const std::filesystem::path normalizedSource(work.sourcePathText);
                itemHr = ResolvePath(normalizedSource, &source, false, false);
            }

            if (SUCCEEDED(itemHr))
            {
                const std::filesystem::path normalizedDestinationParent(work.destinationParentText);
                itemHr = ResolvePath(normalizedDestinationParent, &destinationParent, false, true);
            }

            if (SUCCEEDED(itemHr) && source && destinationParent)
            {
                if (source->isDirectory)
                {
                    if (! context.recursive)
                    {
                        itemHr = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                    }
                    else
                    {
                        DummyNode* destinationDirectory = nullptr;
                        std::vector<DummyNode*> children;
                        std::vector<std::wstring> childNames;

                        EnsureChildrenGenerated(*source);
                        itemHr = CreateDirectoryClone(*source, *destinationParent, work.destinationName, flags, &destinationDirectory);
                        if (SUCCEEDED(itemHr) && destinationDirectory)
                        {
                            children.reserve(source->children.size());
                            childNames.reserve(source->children.size());
                            for (const auto& child : source->children)
                            {
                                if (! child)
                                {
                                    continue;
                                }

                                children.push_back(child.get());
                                childNames.push_back(child->name);
                            }
                        }

                        if (SUCCEEDED(itemHr) && destinationDirectory)
                        {
                            addToTotalItems(static_cast<unsigned long>(children.size()));

                            for (size_t childIndex = 0; childIndex < children.size(); ++childIndex)
                            {
                                const std::wstring& childName = childNames[childIndex];

                                CopyWorkItem childWork{};
                                childWork.source                = children[childIndex];
                                childWork.destinationParent     = destinationDirectory;
                                childWork.sourcePathText        = AppendPath(work.sourcePathText, childName);
                                childWork.destinationParentText = destinationPathText;
                                childWork.destinationName       = childName;
                                childWork.preResolvedHr         = S_OK;
                                stack.push_back(std::move(childWork));
                            }
                        }
                    }
                }
                else
                {
                    itemHr = CopyNode(*source, *destinationParent, work.destinationName, flags, &itemBytes);
                }
            }
        }

        if (SUCCEEDED(itemHr))
        {
            NotifyDirectoryWatchers(work.destinationParentText, work.destinationName, FILESYSTEM_DIR_CHANGE_ADDED);
        }

        const uint64_t baseCompletedBytes = context.completedBytes;
        context.throughputSeed            = CombineSeed(effectiveSeed, work.sourcePathText);
        context.throughputSeed            = CombineSeed(context.throughputSeed, destinationPathText);

        HRESULT hr = SetProgressPaths(context, work.sourcePathText.c_str(), destinationPathText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        if (SUCCEEDED(itemHr))
        {
            hr = ReportThrottledByteProgress(context, itemBytes, baseCompletedBytes, virtualLimitBytesPerSecond);
        }
        else
        {
            context.completedBytes = baseCompletedBytes;
            hr                     = ReportProgress(context, itemBytes, 0);
        }
        if (FAILED(hr))
        {
            return hr;
        }

        hr = SetItemPaths(context, work.sourcePathText.c_str(), destinationPathText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        const unsigned long itemIndex = context.completedItems;
        hr                            = ReportItemCompleted(context, itemIndex, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.completedItems < std::numeric_limits<unsigned long>::max())
        {
            context.completedItems += 1;
        }

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::MoveItems(const wchar_t* const* sourcePaths,
                                                     unsigned long count,
                                                     const wchar_t* destinationFolder,
                                                     FileSystemFlags flags,
                                                     const FileSystemOptions* options,
                                                     IFileSystemCallback* callback,
                                                     void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, count);
    std::uint64_t effectiveSeed = 0;
    {
        std::scoped_lock lock(_mutex);
        context.latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed               = _effectiveSeed;
    }

    const uint64_t virtualLimitBytesPerSecond = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    context.virtualLimitBytesPerSecond        = virtualLimitBytesPerSecond;

    const std::filesystem::path normalizedDestinationFolder = NormalizePath(destinationFolder);
    const std::wstring destinationFolderText                = normalizedDestinationFolder.wstring();

    {
        std::scoped_lock lock(_mutex);
        DummyNode* destinationRoot = nullptr;
        const HRESULT hrResolve    = ResolvePath(normalizedDestinationFolder, &destinationRoot, false, true);
        if (FAILED(hrResolve))
        {
            return hrResolve;
        }
    }

    enum class MoveWorkKind
    {
        MoveNode,
        CleanupDirectory,
    };

    struct MoveWorkItem
    {
        MoveWorkKind kind = MoveWorkKind::MoveNode;
        std::wstring sourcePathText;
        std::wstring destinationParentText;
        std::wstring destinationName;
        HRESULT preResolvedHr = S_OK;
    };

    std::vector<MoveWorkItem> stack;
    stack.reserve(static_cast<size_t>(count));

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::filesystem::path normalizedSource = NormalizePath(sourcePath);
        const std::wstring sourceText                = normalizedSource.wstring();

        MoveWorkItem item{};
        item.kind                  = MoveWorkKind::MoveNode;
        item.sourcePathText        = sourceText;
        item.destinationParentText = destinationFolderText;
        item.destinationName       = std::wstring(leaf);
        stack.push_back(std::move(item));
    }

    auto addToTotalItems = [&](unsigned long delta) noexcept
    {
        if (delta == 0)
        {
            return;
        }

        constexpr unsigned long maxValue = (std::numeric_limits<unsigned long>::max)();
        if (context.totalItems > maxValue - delta)
        {
            context.totalItems = maxValue;
            return;
        }

        context.totalItems += delta;
    };

    bool hadFailure = false;

    while (! stack.empty())
    {
        MoveWorkItem work = std::move(stack.back());
        stack.pop_back();

        HRESULT hrCancel = CheckCancel(context);
        if (FAILED(hrCancel))
        {
            return hrCancel;
        }

        if (work.kind == MoveWorkKind::CleanupDirectory)
        {
            std::scoped_lock lock(_mutex);

            DummyNode* source = nullptr;
            const std::filesystem::path normalizedSource(work.sourcePathText);
            if (SUCCEEDED(ResolvePath(normalizedSource, &source, false, false)) && source && source->isDirectory && source->childrenGenerated &&
                source->children.empty() && source->parent != nullptr)
            {
                ExtractChild(source->parent, source);
            }
            continue;
        }

        const std::wstring destinationPathText = AppendPath(work.destinationParentText, work.destinationName);

        uint64_t itemBytes = 0;
        HRESULT itemHr     = work.preResolvedHr;

        DummyNode* source            = nullptr;
        DummyNode* destinationParent = nullptr;

        {
            std::scoped_lock lock(_mutex);

            if (SUCCEEDED(itemHr))
            {
                const std::filesystem::path normalizedSource(work.sourcePathText);
                itemHr = ResolvePath(normalizedSource, &source, false, false);
            }

            if (SUCCEEDED(itemHr))
            {
                const std::filesystem::path normalizedDestinationParent(work.destinationParentText);
                itemHr = ResolvePath(normalizedDestinationParent, &destinationParent, false, true);
            }

            if (SUCCEEDED(itemHr) && source && destinationParent)
            {
                if (IsAncestor(*source, *destinationParent))
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
            }

            if (SUCCEEDED(itemHr) && source && destinationParent)
            {
                if (source->isDirectory)
                {
                    if (! context.recursive)
                    {
                        itemHr = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                    }
                    else
                    {
                        DummyNode* destinationDirectory = nullptr;
                        std::vector<DummyNode*> children;
                        std::vector<std::wstring> childNames;

                        EnsureChildrenGenerated(*source);
                        itemHr = CreateDirectoryClone(*source, *destinationParent, work.destinationName, flags, &destinationDirectory);
                        if (SUCCEEDED(itemHr) && destinationDirectory)
                        {
                            children.reserve(source->children.size());
                            childNames.reserve(source->children.size());
                            for (const auto& child : source->children)
                            {
                                if (! child)
                                {
                                    continue;
                                }

                                children.push_back(child.get());
                                childNames.push_back(child->name);
                            }
                        }

                        if (SUCCEEDED(itemHr) && destinationDirectory)
                        {
                            addToTotalItems(static_cast<unsigned long>(children.size()));

                            MoveWorkItem cleanup{};
                            cleanup.kind           = MoveWorkKind::CleanupDirectory;
                            cleanup.sourcePathText = work.sourcePathText;
                            stack.push_back(std::move(cleanup));

                            for (size_t childIndex = 0; childIndex < children.size(); ++childIndex)
                            {
                                const std::wstring& childName = childNames[childIndex];

                                MoveWorkItem childWork{};
                                childWork.kind                  = MoveWorkKind::MoveNode;
                                childWork.sourcePathText        = AppendPath(work.sourcePathText, childName);
                                childWork.destinationParentText = destinationPathText;
                                childWork.destinationName       = childName;
                                stack.push_back(std::move(childWork));
                            }
                        }
                    }
                }
                else
                {
                    itemHr = MoveNode(*source, *destinationParent, work.destinationName, flags, &itemBytes);
                }
            }
        }

        if (SUCCEEDED(itemHr))
        {
            const std::filesystem::path normalizedSourcePath(work.sourcePathText);
            const std::wstring sourceParentText = normalizedSourcePath.parent_path().wstring();
            const std::wstring sourceLeafText   = normalizedSourcePath.filename().wstring();

            if (EqualsNoCase(sourceParentText, work.destinationParentText))
            {
                if (sourceLeafText != work.destinationName)
                {
                    NotifyDirectoryWatchers(work.destinationParentText, sourceLeafText, work.destinationName);
                }
                else
                {
                    NotifyDirectoryWatchers(work.destinationParentText, work.destinationName, FILESYSTEM_DIR_CHANGE_MODIFIED);
                }
            }
            else
            {
                NotifyDirectoryWatchers(sourceParentText, sourceLeafText, FILESYSTEM_DIR_CHANGE_REMOVED);
                NotifyDirectoryWatchers(work.destinationParentText, work.destinationName, FILESYSTEM_DIR_CHANGE_ADDED);
            }
        }

        const uint64_t baseCompletedBytes = context.completedBytes;
        context.throughputSeed            = CombineSeed(effectiveSeed, work.sourcePathText);
        context.throughputSeed            = CombineSeed(context.throughputSeed, destinationPathText);

        HRESULT hr = SetProgressPaths(context, work.sourcePathText.c_str(), destinationPathText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        if (SUCCEEDED(itemHr))
        {
            hr = ReportThrottledByteProgress(context, itemBytes, baseCompletedBytes, virtualLimitBytesPerSecond);
        }
        else
        {
            context.completedBytes = baseCompletedBytes;
            hr                     = ReportProgress(context, itemBytes, 0);
        }
        if (FAILED(hr))
        {
            return hr;
        }

        hr = SetItemPaths(context, work.sourcePathText.c_str(), destinationPathText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        const unsigned long itemIndex = context.completedItems;
        hr                            = ReportItemCompleted(context, itemIndex, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.completedItems < std::numeric_limits<unsigned long>::max())
        {
            context.completedItems += 1;
        }

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::DeleteItems(const wchar_t* const* paths,
                                                       unsigned long count,
                                                       FileSystemFlags flags,
                                                       const FileSystemOptions* options,
                                                       IFileSystemCallback* callback,
                                                       void* cookie) noexcept
{
    if (! paths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, count);

    constexpr uint64_t kVirtualDeleteBytesPerItem = 64ull * 1024ull;
    const uint64_t virtualLimitBytesPerSecond     = _virtualSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    std::uint64_t effectiveSeed                   = 0;
    {
        std::scoped_lock lock(_mutex);
        context.latencyMilliseconds = _latencyMilliseconds;
        effectiveSeed               = _effectiveSeed;
    }
    context.virtualLimitBytesPerSecond = virtualLimitBytesPerSecond;
    context.totalBytes                 = 0;
    if (kVirtualDeleteBytesPerItem > 0)
    {
        const uint64_t count64 = static_cast<uint64_t>(count);
        if (count64 > (std::numeric_limits<uint64_t>::max() / kVirtualDeleteBytesPerItem))
        {
            context.totalBytes = std::numeric_limits<uint64_t>::max();
        }
        else
        {
            context.totalBytes = kVirtualDeleteBytesPerItem * count64;
        }
    }

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* path = paths[index];
        if (! path)
        {
            return E_POINTER;
        }

        if (path[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::filesystem::path normalized = NormalizePath(path);
        const std::wstring pathText            = normalized.wstring();
        const std::wstring parentText          = normalized.parent_path().wstring();
        const std::wstring leafText            = normalized.filename().wstring();

        HRESULT itemHr = S_OK;
        {
            std::scoped_lock lock(_mutex);
            DummyNode* node = nullptr;
            itemHr          = ResolvePath(normalized, &node, false, false);
            if (SUCCEEDED(itemHr))
            {
                itemHr = DeleteNode(*node, flags);
            }
        }

        if (SUCCEEDED(itemHr))
        {
            NotifyDirectoryWatchers(parentText, leafText, FILESYSTEM_DIR_CHANGE_REMOVED);
        }

        const uint64_t baseCompletedBytes = context.completedBytes;
        context.throughputSeed            = CombineSeed(effectiveSeed, pathText);

        hr = SetProgressPaths(context, pathText.c_str(), nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (SUCCEEDED(itemHr))
        {
            hr = ReportThrottledByteProgress(context, kVirtualDeleteBytesPerItem, baseCompletedBytes, virtualLimitBytesPerSecond);
        }
        else
        {
            context.completedBytes = baseCompletedBytes;
            hr                     = ReportProgress(context, kVirtualDeleteBytesPerItem, 0);
        }
        if (FAILED(hr))
        {
            return hr;
        }

        hr = SetItemPaths(context, pathText.c_str(), nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemDummy::RenameItems(const FileSystemRenamePair* items,
                                                       unsigned long count,
                                                       FileSystemFlags flags,
                                                       const FileSystemOptions* options,
                                                       IFileSystemCallback* callback,
                                                       void* cookie) noexcept
{
    if (! items && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, count);

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const FileSystemRenamePair& item = items[index];
        if (! item.sourcePath || ! item.newName)
        {
            return E_POINTER;
        }

        if (item.sourcePath[0] == L'\0' || item.newName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        if (! IsNameValid(item.newName))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::filesystem::path normalizedSource = NormalizePath(item.sourcePath);
        const std::wstring sourceText                = normalizedSource.wstring();
        const std::wstring directory                 = normalizedSource.parent_path().wstring();
        const std::wstring sourceLeafText            = normalizedSource.filename().wstring();
        const std::wstring destinationText           = AppendPath(directory, item.newName);

        uint64_t itemBytes = 0;
        HRESULT itemHr     = S_OK;

        {
            std::scoped_lock lock(_mutex);

            DummyNode* source = nullptr;
            itemHr            = ResolvePath(normalizedSource, &source, false, false);

            if (SUCCEEDED(itemHr) && source)
            {
                DummyNode* sourceParent = source->parent;
                if (! sourceParent)
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                }
                else
                {
                    itemHr = MoveNode(*source, *sourceParent, item.newName, flags, &itemBytes);
                }
            }
        }

        if (SUCCEEDED(itemHr))
        {
            if (sourceLeafText != item.newName)
            {
                NotifyDirectoryWatchers(directory, sourceLeafText, item.newName);
            }
            else
            {
                NotifyDirectoryWatchers(directory, sourceLeafText, FILESYSTEM_DIR_CHANGE_MODIFIED);
            }
        }

        hr = SetProgressPaths(context, sourceText.c_str(), destinationText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ReportProgress(context, itemBytes, itemBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = SetItemPaths(context, sourceText.c_str(), destinationText.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

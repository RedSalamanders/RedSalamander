#include "FileSystemCurl.Internal.h"

using namespace FileSystemCurlInternal;

namespace FileSystemCurlInternal
{
[[nodiscard]] std::wstring NormalizePluginPath(std::wstring_view rawPath) noexcept;

[[nodiscard]] bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept
{
    return (static_cast<unsigned long>(flags) & static_cast<unsigned long>(flag)) != 0u;
}

[[nodiscard]] bool IsCancellationHr(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] HRESULT NormalizeCancellation(HRESULT hr) noexcept
{
    if (IsCancellationHr(hr))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
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

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
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

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val || ! yyjson_is_str(val))
    {
        return std::nullopt;
    }

    const char* s = yyjson_get_str(val);
    if (! s)
    {
        return std::nullopt;
    }

    const size_t len        = yyjson_get_len(val);
    const std::wstring wide = Utf16FromUtf8(std::string_view(s, len));
    if (wide.empty() && len != 0u)
    {
        return std::nullopt;
    }

    return wide;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val)
    {
        return std::nullopt;
    }

    if (yyjson_is_uint(val))
    {
        return yyjson_get_uint(val);
    }

    if (yyjson_is_sint(val))
    {
        const int64_t s = yyjson_get_sint(val);
        if (s >= 0)
        {
            return static_cast<uint64_t>(s);
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val)
    {
        return std::nullopt;
    }

    if (yyjson_is_bool(val))
    {
        return yyjson_get_bool(val) != 0;
    }

    if (yyjson_is_sint(val))
    {
        return yyjson_get_sint(val) != 0;
    }

    if (yyjson_is_uint(val))
    {
        return yyjson_get_uint(val) != 0;
    }

    return std::nullopt;
}
} // namespace FileSystemCurlInternal

// FilesInformationCurl

HRESULT STDMETHODCALLTYPE FilesInformationCurl::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
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

ULONG STDMETHODCALLTYPE FilesInformationCurl::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FilesInformationCurl::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    if (ppFileInfo == nullptr)
    {
        return E_POINTER;
    }

    *ppFileInfo = nullptr;

    if (_usedBytes == 0 || _buffer.empty())
    {
        return S_OK;
    }

    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetBufferSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    *pSize = _usedBytes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetAllocatedSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    if (_buffer.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *pSize = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetCount(unsigned long* pCount) noexcept
{
    if (pCount == nullptr)
    {
        return E_POINTER;
    }

    *pCount = _count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    if (ppEntry == nullptr)
    {
        return E_POINTER;
    }

    *ppEntry = nullptr;

    if (index >= _count)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    return LocateEntry(index, ppEntry);
}

size_t FilesInformationCurl::AlignUp(size_t value, size_t alignment) noexcept
{
    const size_t mask = alignment - 1u;
    return (value + mask) & ~mask;
}

size_t FilesInformationCurl::ComputeEntrySizeBytes(std::wstring_view name) noexcept
{
    const size_t baseSize = offsetof(FileInfo, FileName);
    const size_t nameSize = name.size() * sizeof(wchar_t);
    return AlignUp(baseSize + nameSize + sizeof(wchar_t), sizeof(unsigned long));
}

HRESULT FilesInformationCurl::BuildFromEntries(std::vector<Entry> entries) noexcept
{
    _buffer.clear();
    _count     = 0;
    _usedBytes = 0;

    if (entries.empty())
    {
        return S_OK;
    }

    std::sort(entries.begin(),
              entries.end(),
              [](const Entry& a, const Entry& b)
              {
                  const int cmp = OrdinalString::Compare(a.name, b.name, true);
                  if (cmp != 0)
                  {
                      return cmp < 0;
                  }

                  const bool aDir = (a.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  const bool bDir = (b.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  if (aDir != bDir)
                  {
                      return aDir;
                  }

                  return a.sizeBytes < b.sizeBytes;
              });

    size_t totalBytes = 0;
    for (const auto& entry : entries)
    {
        totalBytes += ComputeEntrySizeBytes(entry.name);
        if (totalBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }

    _buffer.resize(totalBytes, std::byte{0});

    std::byte* base     = _buffer.data();
    size_t offset       = 0;
    FileInfo* previous  = nullptr;
    size_t previousSize = 0;

    for (const auto& source : entries)
    {
        const size_t entrySize = ComputeEntrySizeBytes(source.name);
        if (offset + entrySize > _buffer.size())
        {
            return E_FAIL;
        }

        auto* entry = reinterpret_cast<FileInfo*>(base + offset);
        std::memset(entry, 0, entrySize);

        const size_t nameBytes = source.name.size() * sizeof(wchar_t);
        if (nameBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        entry->FileAttributes = source.attributes;
        entry->FileIndex      = source.fileIndex;
        entry->EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry->AllocationSize = static_cast<__int64>(source.sizeBytes);

        entry->CreationTime   = source.creationTime;
        entry->LastAccessTime = source.lastAccessTime;
        entry->LastWriteTime  = source.lastWriteTime;
        entry->ChangeTime     = source.changeTime;

        entry->FileNameSize = static_cast<unsigned long>(nameBytes);
        if (! source.name.empty())
        {
            std::memcpy(entry->FileName, source.name.data(), nameBytes);
        }
        entry->FileName[source.name.size()] = L'\0';

        if (previous)
        {
            previous->NextEntryOffset = static_cast<unsigned long>(previousSize);
        }

        previous     = entry;
        previousSize = entrySize;

        offset += entrySize;
        ++_count;
    }

    _usedBytes = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT FilesInformationCurl::LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept
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

        const size_t advance = (entry->NextEntryOffset != 0)
                                   ? static_cast<size_t>(entry->NextEntryOffset)
                                   : ComputeEntrySizeBytes(std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t)));
        if (advance == 0)
        {
            break;
        }

        offset += advance;
        ++currentIndex;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
}

namespace FileSystemCurlInternal
{
[[nodiscard]] std::wstring_view TrimTrailingSlash(std::wstring_view path) noexcept
{
    while (! path.empty() && path.back() == L'/')
    {
        path.remove_suffix(1);
    }
    return path;
}

[[nodiscard]] std::string EscapeUrlPath(std::wstring_view path) noexcept
{
    const std::string utf8 = Utf8FromUtf16(path);
    if (utf8.empty())
    {
        return {};
    }

    std::string out;
    out.reserve(utf8.size());

    auto isUnreserved = [](unsigned char ch) noexcept
    {
        if (ch >= 'a' && ch <= 'z')
        {
            return true;
        }
        if (ch >= 'A' && ch <= 'Z')
        {
            return true;
        }
        if (ch >= '0' && ch <= '9')
        {
            return true;
        }
        return ch == '-' || ch == '.' || ch == '_' || ch == '~';
    };

    constexpr char kHex[] = "0123456789ABCDEF";
    for (const char chRaw : utf8)
    {
        const unsigned char ch = static_cast<unsigned char>(chRaw);
        if (ch == '/')
        {
            out.push_back('/');
            continue;
        }

        if (isUnreserved(ch))
        {
            out.push_back(static_cast<char>(ch));
            continue;
        }

        out.push_back('%');
        out.push_back(kHex[(ch >> 4) & 0x0F]);
        out.push_back(kHex[ch & 0x0F]);
    }

    return out;
}

[[nodiscard]] std::wstring ProtocolToDisplay(Protocol protocol)
{
    switch (protocol)
    {
        case Protocol::Ftp: return L"FTP";
        case Protocol::Sftp: return L"SFTP";
        case Protocol::Scp: return L"SCP";
        case Protocol::Imap: return L"IMAP";
    }
    return L"";
}

[[nodiscard]] std::string_view ProtocolSchemeForTransfer(Protocol protocol) noexcept
{
    switch (protocol)
    {
        case Protocol::Ftp: return "ftp";
        case Protocol::Sftp: return "sftp";
        case Protocol::Scp: return "scp";
        case Protocol::Imap: return "imap";
        default: return "sftp";
    }
}

[[nodiscard]] std::string_view ProtocolSchemeForCommands(Protocol protocol) noexcept
{
    // SCP does not support directory listing; use SFTP for list/quote operations.
    return protocol == Protocol::Scp ? "sftp" : ProtocolSchemeForTransfer(protocol);
}

[[nodiscard]] bool LooksLikeUrl(std::string_view text) noexcept
{
    return text.find("://") != std::string_view::npos;
}

[[nodiscard]] std::string TrimAscii(std::string_view text) noexcept
{
    while (! text.empty() && (text.front() == ' ' || text.front() == '\t' || text.front() == '\r' || text.front() == '\n'))
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && (text.back() == ' ' || text.back() == '\t' || text.back() == '\r' || text.back() == '\n'))
    {
        text.remove_suffix(1);
    }
    return std::string(text);
}

[[nodiscard]] bool TryParseUnixListLine(std::string_view line, FilesInformationCurl::Entry& out) noexcept
{
    if (line.size() < 2)
    {
        return false;
    }

    if (line.rfind("total ", 0) == 0)
    {
        return false;
    }

    const char type = line[0];
    if (type != 'd' && type != '-' && type != 'l')
    {
        return false;
    }

    auto skipSpaces = [&](size_t& pos) noexcept
    {
        while (pos < line.size() && (line[pos] == ' ' || line[pos] == '\t'))
        {
            ++pos;
        }
    };

    auto nextToken = [&](size_t& pos) noexcept -> std::optional<std::string_view>
    {
        skipSpaces(pos);
        if (pos >= line.size())
        {
            return std::nullopt;
        }
        const size_t start = pos;
        while (pos < line.size() && line[pos] != ' ' && line[pos] != '\t')
        {
            ++pos;
        }
        return line.substr(start, pos - start);
    };

    size_t pos = 0;
    if (! nextToken(pos).has_value())
    {
        return false;
    }
    static_cast<void>(nextToken(pos)); // links
    static_cast<void>(nextToken(pos)); // owner
    static_cast<void>(nextToken(pos)); // group
    const auto sizeTok = nextToken(pos);
    static_cast<void>(nextToken(pos)); // month
    static_cast<void>(nextToken(pos)); // day
    static_cast<void>(nextToken(pos)); // time/year

    if (! sizeTok.has_value())
    {
        return false;
    }

    skipSpaces(pos);
    if (pos >= line.size())
    {
        return false;
    }

    std::string_view namePart = line.substr(pos);
    if (IsDotOrDotDotName(namePart))
    {
        return false;
    }

    const size_t arrow = namePart.find(" -> ");
    if (arrow != std::string_view::npos)
    {
        namePart = namePart.substr(0, arrow);
    }

    unsigned __int64 sizeBytes = 0;
    {
        unsigned long long parsed = 0;
        if (sscanf_s(std::string(sizeTok.value()).c_str(), "%llu", &parsed) == 1)
        {
            sizeBytes = parsed;
        }
    }

    out            = {};
    out.attributes = (type == 'd') ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    out.sizeBytes  = sizeBytes;
    out.name       = Utf16FromUtf8(namePart);
    return ! out.name.empty();
}

[[nodiscard]] bool TryParseDosListLine(std::string_view line, FilesInformationCurl::Entry& out) noexcept
{
    // Example:
    // 01-02-24  03:04PM       <DIR>          Folder
    // 01-02-24  03:04PM                1234 File.txt

    auto skipSpaces = [&](size_t& pos) noexcept
    {
        while (pos < line.size() && (line[pos] == ' ' || line[pos] == '\t'))
        {
            ++pos;
        }
    };

    auto nextToken = [&](size_t& pos) noexcept -> std::optional<std::string_view>
    {
        skipSpaces(pos);
        if (pos >= line.size())
        {
            return std::nullopt;
        }
        const size_t start = pos;
        while (pos < line.size() && line[pos] != ' ' && line[pos] != '\t')
        {
            ++pos;
        }
        return line.substr(start, pos - start);
    };

    size_t pos = 0;
    if (! nextToken(pos).has_value())
    {
        return false;
    }
    if (! nextToken(pos).has_value())
    {
        return false;
    }
    const auto sizeOrDir = nextToken(pos);
    if (! sizeOrDir.has_value())
    {
        return false;
    }

    skipSpaces(pos);
    if (pos >= line.size())
    {
        return false;
    }

    std::string_view namePart = line.substr(pos);
    if (IsDotOrDotDotName(namePart))
    {
        return false;
    }

    out = {};
    if (sizeOrDir.value() == "<DIR>")
    {
        out.attributes = FILE_ATTRIBUTE_DIRECTORY;
        out.sizeBytes  = 0;
    }
    else
    {
        unsigned long long parsed = 0;
        if (sscanf_s(std::string(sizeOrDir.value()).c_str(), "%llu", &parsed) != 1)
        {
            return false;
        }
        out.attributes = FILE_ATTRIBUTE_NORMAL;
        out.sizeBytes  = parsed;
    }

    out.name = Utf16FromUtf8(namePart);
    return ! out.name.empty();
}

[[nodiscard]] HRESULT ParseDirectoryListing(std::string_view listing, std::vector<FilesInformationCurl::Entry>& out) noexcept
{
    out.clear();

    size_t start = 0;
    while (start < listing.size())
    {
        size_t end = listing.find('\n', start);
        if (end == std::string_view::npos)
        {
            end = listing.size();
        }

        std::string_view line = listing.substr(start, end - start);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }

        if (! line.empty())
        {
            FilesInformationCurl::Entry entry{};
            if (TryParseUnixListLine(line, entry) || TryParseDosListLine(line, entry))
            {
                out.push_back(std::move(entry));
            }
        }

        start = end + 1u;
    }

    return S_OK;
}

[[nodiscard]] std::optional<FilesInformationCurl::Entry> FindEntryByName(const std::vector<FilesInformationCurl::Entry>& entries,
                                                                         std::wstring_view leaf) noexcept
{
    for (const auto& entry : entries)
    {
        if (entry.name == leaf)
        {
            return entry;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool TryParsePort(std::wstring_view text, unsigned int& out) noexcept
{
    if (text.empty())
    {
        return false;
    }

    unsigned long long value = 0;
    for (wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }
        value = (value * 10ull) + static_cast<unsigned long long>(ch - L'0');
        if (value > 65535ull)
        {
            return false;
        }
    }

    out = static_cast<unsigned int>(value);
    return true;
}

[[nodiscard]] HRESULT ResolveLocation(Protocol protocol,
                                      const FileSystemCurl::Settings& settings,
                                      std::wstring_view pluginPath,
                                      IHostConnections* hostConnections,
                                      bool acquireSecrets,
                                      ResolvedLocation& out) noexcept
{
    out                     = {};
    out.connection.protocol = protocol;

    const auto equalsNoCase = [](std::wstring_view a, std::wstring_view b) noexcept
    {
        if (a.size() != b.size() || a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        {
            return false;
        }
        const int len = static_cast<int>(a.size());
        return CompareStringOrdinal(a.data(), len, b.data(), len, TRUE) == CSTR_EQUAL;
    };

    out.connection.ftpUseEpsv         = settings.ftpUseEpsv;
    out.connection.connectTimeoutMs   = settings.connectTimeoutMs;
    out.connection.operationTimeoutMs = settings.operationTimeoutMs;
    out.connection.ignoreSslTrust     = settings.ignoreSslTrust;
    out.connection.sshPrivateKey      = Utf8FromUtf16(settings.sshPrivateKey);
    out.connection.sshPublicKey       = Utf8FromUtf16(settings.sshPublicKey);
    out.connection.sshKeyPassphrase   = Utf8FromUtf16(settings.sshKeyPassphrase);
    out.connection.sshKnownHosts      = Utf8FromUtf16(settings.sshKnownHosts);

    const std::wstring normalizedFull = NormalizePluginPath(pluginPath);

    std::wstring_view authority;
    std::wstring_view pathPart;

    if (normalizedFull.size() >= 2u && normalizedFull[0] == L'/' && normalizedFull[1] == L'/')
    {
        std::wstring_view after(normalizedFull);
        after.remove_prefix(2);
        const size_t slashPos = after.find(L'/');
        authority             = slashPos == std::wstring_view::npos ? after : after.substr(0, slashPos);
        pathPart              = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : after.substr(slashPos);
    }
    else
    {
        pathPart = normalizedFull;
    }

    // Host-reserved Connection Manager prefix:
    // - /@conn:<connectionName>/...
    // The host resolves the profile and the plugin requests credentials through IHostConnections.
    constexpr std::wstring_view kConnPrefix = L"/@conn:";

    std::wstring_view connectionName;
    std::wstring_view connPath;
    bool hasConnPrefix = false;

    if (authority.empty())
    {
        const std::wstring_view full = normalizedFull;
        if (full.size() >= kConnPrefix.size() && full.substr(0, kConnPrefix.size()) == kConnPrefix)
        {
            std::wstring_view rest = full.substr(kConnPrefix.size());
            const size_t slashPos  = rest.find(L'/');
            connectionName         = slashPos == std::wstring_view::npos ? rest : rest.substr(0, slashPos);
            connPath               = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : rest.substr(slashPos);
            hasConnPrefix          = true;
        }
    }
    else if (equalsNoCase(authority, L"@conn"))
    {
        // URI-style shorthand: // @conn / <connectionName> / ...
        std::wstring_view rest = pathPart;
        while (! rest.empty() && rest.front() == L'/')
        {
            rest.remove_prefix(1);
        }

        const size_t slashPos = rest.find(L'/');
        connectionName        = slashPos == std::wstring_view::npos ? rest : rest.substr(0, slashPos);
        connPath              = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : rest.substr(slashPos);
        hasConnPrefix         = true;
    }

    if (hasConnPrefix)
    {
        if (! hostConnections)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        if (connectionName.empty())
        {
            return E_INVALIDARG;
        }

        std::wstring connectionNameText;
        connectionNameText.assign(connectionName);

        wil::unique_cotaskmem_ptr<char> json;
        {
            char* rawJson        = nullptr;
            const HRESULT jsonHr = hostConnections->GetConnectionJsonUtf8(connectionNameText.c_str(), &rawJson);
            if (FAILED(jsonHr))
            {
                return jsonHr;
            }
            json.reset(rawJson);
        }

        if (! json || ! json.get()[0])
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_doc* doc = yyjson_read(json.get(), strlen(json.get()), YYJSON_READ_ALLOW_BOM);
        if (! doc)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

        yyjson_val* root = yyjson_doc_get_root(doc);
        if (! root || ! yyjson_is_obj(root))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        auto getStr = [&](const char* key) -> std::optional<std::string_view>
        {
            yyjson_val* v = yyjson_obj_get(root, key);
            if (! v || ! yyjson_is_str(v))
            {
                return std::nullopt;
            }
            const char* s = yyjson_get_str(v);
            return s ? std::make_optional(std::string_view(s, yyjson_get_len(v))) : std::nullopt;
        };

        auto getBool = [&](const char* key, bool& outBool) noexcept -> bool
        {
            yyjson_val* v = yyjson_obj_get(root, key);
            if (! v || ! yyjson_is_bool(v))
            {
                return false;
            }
            outBool = yyjson_get_bool(v) != 0;
            return true;
        };

        auto getUInt = [&](const char* key, unsigned int& outUInt) noexcept -> bool
        {
            yyjson_val* v = yyjson_obj_get(root, key);
            if (! v || ! yyjson_is_uint(v))
            {
                return false;
            }
            const uint64_t value = yyjson_get_uint(v);
            outUInt              = static_cast<unsigned int>(std::min<uint64_t>(value, 0xFFFFFFFFull));
            return true;
        };

        const auto pluginIdUtf8 = getStr("pluginId");
        if (! pluginIdUtf8.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        if (const auto idUtf8 = getStr("id"); idUtf8.has_value())
        {
            out.connection.connectionId = Utf16FromUtf8(*idUtf8);
        }

        if ((protocol == Protocol::Ftp && *pluginIdUtf8 != "builtin/file-system-ftp") ||
            (protocol == Protocol::Sftp && *pluginIdUtf8 != "builtin/file-system-sftp") ||
            (protocol == Protocol::Scp && *pluginIdUtf8 != "builtin/file-system-scp") ||
            (protocol == Protocol::Imap && *pluginIdUtf8 != "builtin/file-system-imap"))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const auto hostUtf8 = getStr("host");
        if (! hostUtf8.has_value() || hostUtf8->empty())
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
        }
        out.connection.host = std::string(*hostUtf8);

        unsigned int portValue = 0;
        if (getUInt("port", portValue) && portValue != 0u)
        {
            out.connection.port = portValue;
        }

        const auto userUtf8             = getStr("userName");
        const bool userMissingInProfile = ! userUtf8.has_value() || userUtf8->empty();
        if (userUtf8.has_value())
        {
            out.connection.user = std::string(*userUtf8);
        }

        const auto authModeUtf8 = getStr("authMode");
        const bool anonymous    = authModeUtf8.has_value() && *authModeUtf8 == "anonymous";
        const bool sshKey       = authModeUtf8.has_value() && *authModeUtf8 == "sshKey";
        const bool passwordAuth = ! anonymous && ! sshKey;

        if (anonymous)
        {
            out.connection.user     = "anonymous";
            out.connection.password = {};
        }
        else if (out.connection.user.empty())
        {
            out.connection.user = Utf8FromUtf16(settings.defaultUser);
        }

        if (protocol == Protocol::Ftp && out.connection.user.empty())
        {
            out.connection.user = "anonymous";
        }

        bool savePassword = false;
        static_cast<void>(getBool("savePassword", savePassword));
        bool requireWindowsHello = true;
        static_cast<void>(getBool("requireWindowsHello", requireWindowsHello));
        bool ignoreSslTrust = false;
        static_cast<void>(getBool("ignoreSslTrust", ignoreSslTrust));

        out.connection.fromConnectionManagerProfile = true;
        out.connection.connectionName               = connectionNameText;
        out.connection.connectionSavePassword       = savePassword;
        out.connection.connectionRequireHello       = requireWindowsHello;
        if (protocol == Protocol::Imap)
        {
            out.connection.ignoreSslTrust = ignoreSslTrust;
        }

        if (acquireSecrets && passwordAuth)
        {
            wil::unique_cotaskmem_string secret;
            wchar_t* rawSecret = nullptr;
            bool prompted      = false;
            HRESULT secretHr   = hostConnections->GetConnectionSecret(connectionNameText.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, &rawSecret);
            if (secretHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
            {
                rawSecret = nullptr;
                secretHr  = hostConnections->PromptForConnectionSecret(connectionNameText.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, &rawSecret);
                prompted  = true;
                if (secretHr == S_FALSE)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
            }
            if (FAILED(secretHr))
            {
                Debug::Error(L"Connection Manager secret retrieval failed protocol={} connection='{}' host='{}' user='{}' path='{}' hr=0x{:08X}",
                             ProtocolToDisplay(protocol),
                             connectionNameText,
                             Utf16FromUtf8(out.connection.host),
                             Utf16FromUtf8(out.connection.user),
                             connPath,
                             static_cast<unsigned long>(secretHr));
                return secretHr;
            }

            secret.reset(rawSecret);
            if (! secret.get())
            {
                Debug::Error(L"Connection Manager returned a null password pointer protocol={} connection='{}' id='{}' host='{}' user='{}' path='{}'",
                             ProtocolToDisplay(protocol),
                             connectionNameText,
                             out.connection.connectionId,
                             Utf16FromUtf8(out.connection.host),
                             Utf16FromUtf8(out.connection.user),
                             connPath);
                return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
            }
            if (secret.get()[0] == L'\0')
            {
                Debug::Error(L"Connection Manager returned an empty password protocol={} connection='{}' id='{}' host='{}' user='{}' path='{}'",
                             ProtocolToDisplay(protocol),
                             connectionNameText,
                             out.connection.connectionId,
                             Utf16FromUtf8(out.connection.host),
                             Utf16FromUtf8(out.connection.user),
                             connPath);
                return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
            }

            out.connection.password = Utf8FromUtf16(secret.get());
            if (out.connection.password.empty())
            {
                Debug::Error(
                    L"Connection Manager password conversion failed protocol={} connection='{}' id='{}' host='{}' user='{}' path='{}' (invalid UTF-16)",
                    ProtocolToDisplay(protocol),
                    connectionNameText,
                    out.connection.connectionId,
                    Utf16FromUtf8(out.connection.host),
                    Utf16FromUtf8(out.connection.user),
                    connPath);
                return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            }

            if (prompted && userMissingInProfile)
            {
                wil::unique_cotaskmem_ptr<char> refreshedJson;
                char* rawRefreshed      = nullptr;
                const HRESULT refreshHr = hostConnections->GetConnectionJsonUtf8(connectionNameText.c_str(), &rawRefreshed);
                if (SUCCEEDED(refreshHr) && rawRefreshed && rawRefreshed[0] != '\0')
                {
                    refreshedJson.reset(rawRefreshed);

                    yyjson_doc* refreshedDoc = yyjson_read(refreshedJson.get(), strlen(refreshedJson.get()), YYJSON_READ_ALLOW_BOM);
                    if (refreshedDoc)
                    {
                        auto freeRefreshed = wil::scope_exit([&] { yyjson_doc_free(refreshedDoc); });

                        yyjson_val* refreshedRoot = yyjson_doc_get_root(refreshedDoc);
                        if (refreshedRoot && yyjson_is_obj(refreshedRoot))
                        {
                            yyjson_val* refreshedUser = yyjson_obj_get(refreshedRoot, "userName");
                            if (refreshedUser && yyjson_is_str(refreshedUser))
                            {
                                const char* s    = yyjson_get_str(refreshedUser);
                                const size_t len = yyjson_get_len(refreshedUser);
                                if (s && len > 0)
                                {
                                    out.connection.user.assign(s, len);
                                }
                            }
                        }
                    }
                }
            }
        }

        if (sshKey)
        {
            if (const auto keyPath = getStr("sshPrivateKey"); keyPath.has_value())
            {
                out.connection.sshPrivateKey = std::string(*keyPath);
            }
            if (const auto knownHosts = getStr("sshKnownHosts"); knownHosts.has_value())
            {
                out.connection.sshKnownHosts = std::string(*knownHosts);
            }

            if (acquireSecrets)
            {
                wil::unique_cotaskmem_string secret;
                wchar_t* rawSecret = nullptr;
                HRESULT secretHr =
                    hostConnections->GetConnectionSecret(connectionNameText.c_str(), HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE, nullptr, &rawSecret);
                if (secretHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
                {
                    rawSecret = nullptr;
                    secretHr =
                        hostConnections->PromptForConnectionSecret(connectionNameText.c_str(), HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE, nullptr, &rawSecret);
                    if (secretHr == S_FALSE)
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                }
                if (FAILED(secretHr))
                {
                    Debug::Error(L"Connection Manager passphrase retrieval failed protocol={} connection='{}' host='{}' user='{}' path='{}' hr=0x{:08X}",
                                 ProtocolToDisplay(protocol),
                                 connectionNameText,
                                 Utf16FromUtf8(out.connection.host),
                                 Utf16FromUtf8(out.connection.user),
                                 connPath,
                                 static_cast<unsigned long>(secretHr));
                    return secretHr;
                }

                secret.reset(rawSecret);
                if (! secret.get())
                {
                    Debug::Error(L"Connection Manager returned a null passphrase pointer protocol={} connection='{}' id='{}' host='{}' user='{}' path='{}'",
                                 ProtocolToDisplay(protocol),
                                 connectionNameText,
                                 out.connection.connectionId,
                                 Utf16FromUtf8(out.connection.host),
                                 Utf16FromUtf8(out.connection.user),
                                 connPath);
                    return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
                }

                if (secret.get()[0] != L'\0')
                {
                    out.connection.sshKeyPassphrase = Utf8FromUtf16(secret.get());
                    if (out.connection.sshKeyPassphrase.empty())
                    {
                        Debug::Error(L"Connection Manager passphrase conversion failed protocol={} connection='{}' id='{}' host='{}' user='{}' path='{}' "
                                     L"(invalid UTF-16)",
                                     ProtocolToDisplay(protocol),
                                     connectionNameText,
                                     out.connection.connectionId,
                                     Utf16FromUtf8(out.connection.host),
                                     Utf16FromUtf8(out.connection.user),
                                     connPath);
                        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
                    }
                }
            }
        }

        out.connection.basePath     = "/";
        out.connection.basePathWide = L"/";
        out.remotePath              = NormalizePluginPath(connPath);
        if (out.remotePath.empty())
        {
            out.remotePath = L"/";
        }

        if (out.remotePath.size() >= 2u && out.remotePath[0] == L'/' && out.remotePath[1] == L'/')
        {
            while (out.remotePath.size() > 1u && out.remotePath[0] == L'/' && out.remotePath[1] == L'/')
            {
                out.remotePath.erase(out.remotePath.begin());
            }
        }

        const bool passwordPresent   = ! out.connection.password.empty();
        const bool passphrasePresent = ! out.connection.sshKeyPassphrase.empty();

        const wchar_t* authModeText       = anonymous ? L"anonymous" : sshKey ? L"sshKey" : L"password";
        out.connection.connectionAuthMode = authModeText ? authModeText : L"";
        const unsigned int portOut        = out.connection.port.has_value() ? out.connection.port.value() : 0u;

        Debug::Info(L"ResolveLocation(@conn) protocol={} conn='{}' id='{}' auth='{}' pwdPresent={} remote='{}' host='{}' port={} user='{}' savePwd={} "
                    L"requireHello={} ignoreSslTrust={} passphrasePresent={}",
                    ProtocolToDisplay(protocol),
                    connectionNameText,
                    out.connection.connectionId,
                    authModeText,
                    passwordPresent ? 1 : 0,
                    out.remotePath,
                    Utf16FromUtf8(out.connection.host),
                    portOut,
                    Utf16FromUtf8(out.connection.user),
                    savePassword ? 1 : 0,
                    requireWindowsHello ? 1 : 0,
                    out.connection.ignoreSslTrust ? 1 : 0,
                    passphrasePresent ? 1 : 0);

        return S_OK;
    }

    if (! authority.empty())
    {
        std::wstring_view hostPort = authority;
        std::wstring_view userInfo;

        const size_t at = authority.find(L'@');
        if (at != std::wstring_view::npos)
        {
            userInfo = authority.substr(0, at);
            hostPort = authority.substr(at + 1u);
        }

        if (! userInfo.empty())
        {
            const size_t colon = userInfo.find(L':');
            if (colon != std::wstring_view::npos)
            {
                if (protocol == Protocol::Scp)
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
                out.connection.user     = Utf8FromUtf16(userInfo.substr(0, colon));
                out.connection.password = Utf8FromUtf16(userInfo.substr(colon + 1u));
            }
            else
            {
                out.connection.user = Utf8FromUtf16(userInfo);
            }
        }

        std::wstring_view hostText = hostPort;
        std::optional<unsigned int> port;

        if (! hostPort.empty() && hostPort.front() == L'[')
        {
            const size_t close = hostPort.find(L']');
            if (close == std::wstring_view::npos)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }

            hostText = hostPort.substr(0, close + 1u);
            if (close + 1u < hostPort.size() && hostPort[close + 1u] == L':')
            {
                unsigned int parsed = 0;
                if (! TryParsePort(hostPort.substr(close + 2u), parsed))
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
                port = parsed;
            }
        }
        else
        {
            const size_t colon = hostPort.rfind(L':');
            if (colon != std::wstring_view::npos)
            {
                unsigned int parsed              = 0;
                const std::wstring_view portText = hostPort.substr(colon + 1u);
                if (TryParsePort(portText, parsed))
                {
                    hostText = hostPort.substr(0, colon);
                    port     = parsed;
                }
            }
        }

        out.connection.host = Utf8FromUtf16(hostText);
        if (out.connection.host.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
        }

        if (port.has_value() && port.value() != 0u)
        {
            out.connection.port = port.value();
        }

        if (out.connection.user.empty())
        {
            out.connection.user = Utf8FromUtf16(settings.defaultUser);
        }
        if (out.connection.password.empty())
        {
            out.connection.password = Utf8FromUtf16(settings.defaultPassword);
        }

        if (protocol == Protocol::Ftp && out.connection.user.empty())
        {
            out.connection.user = "anonymous";
        }

        out.connection.basePath     = "/";
        out.connection.basePathWide = L"/";
        out.remotePath              = NormalizePluginPath(pathPart);
        if (out.remotePath.empty())
        {
            out.remotePath = L"/";
        }

        // `pathPart` always starts with '/', so remotePath should never have an authority prefix.
        if (out.remotePath.size() >= 2u && out.remotePath[0] == L'/' && out.remotePath[1] == L'/')
        {
            while (out.remotePath.size() > 1u && out.remotePath[0] == L'/' && out.remotePath[1] == L'/')
            {
                out.remotePath.erase(out.remotePath.begin());
            }
        }

        return S_OK;
    }

    // Default connection (used for `ftp:/...` and also when the authority is missing, like `ftp://`).
    out.connection.host = Utf8FromUtf16(settings.defaultHost);
    if (out.connection.host.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
    }

    if (settings.defaultPort != 0)
    {
        out.connection.port = settings.defaultPort;
    }

    out.connection.user     = Utf8FromUtf16(settings.defaultUser);
    out.connection.password = Utf8FromUtf16(settings.defaultPassword);

    if (protocol == Protocol::Ftp && out.connection.user.empty())
    {
        out.connection.user = "anonymous";
    }

    std::wstring basePath = settings.defaultBasePath;
    if (basePath.empty())
    {
        basePath = L"/";
    }
    if (basePath.front() != L'/')
    {
        basePath.insert(basePath.begin(), L'/');
    }

    basePath = std::wstring(TrimTrailingSlash(basePath));
    if (basePath.empty())
    {
        basePath = L"/";
    }

    out.connection.basePath = EscapeUrlPath(basePath);
    if (out.connection.basePath.empty())
    {
        out.connection.basePath = "/";
    }

    out.connection.basePathWide = basePath;

    out.remotePath = NormalizePluginPath(pathPart);
    if (out.remotePath.empty())
    {
        out.remotePath = L"/";
    }

    // If we were given an authority prefix without a host (e.g. `ftp://`), treat it as `ftp:/`.
    if (out.remotePath.size() >= 2u && out.remotePath[0] == L'/' && out.remotePath[1] == L'/')
    {
        out.remotePath = L"/";
    }

    return S_OK;
}
} // namespace FileSystemCurlInternal

namespace FileSystemCurlInternal
{
[[nodiscard]] std::wstring NormalizePluginPath(std::wstring_view rawPath) noexcept
{
    std::wstring path(rawPath);
    if (path.empty())
    {
        return L"/";
    }

    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    const bool hasAuthorityPrefix = path.size() >= 2u && path[0] == L'/' && path[1] == L'/';

    if (! path.empty() && path.front() != L'/')
    {
        path.insert(path.begin(), L'/');
    }

    std::wstring collapsed;
    collapsed.reserve(path.size());

    bool prevSlash = false;
    size_t index   = 0;
    if (hasAuthorityPrefix)
    {
        collapsed.append(L"//");
        prevSlash = true;
        index     = 2;
        while (index < path.size() && path[index] == L'/')
        {
            ++index;
        }
    }

    for (; index < path.size(); ++index)
    {
        const wchar_t ch = path[index];
        const bool slash = (ch == L'/');
        if (slash && prevSlash)
        {
            continue;
        }
        collapsed.push_back(ch);
        prevSlash = slash;
    }

    if (collapsed.empty())
    {
        return L"/";
    }

    return collapsed;
}

[[nodiscard]] std::wstring EnsureTrailingSlash(std::wstring_view path) noexcept
{
    std::wstring normalized = NormalizePluginPath(path);
    if (! normalized.empty() && normalized.back() != L'/')
    {
        normalized.push_back(L'/');
    }
    return normalized.empty() ? std::wstring(L"/") : normalized;
}

[[nodiscard]] std::wstring_view LeafName(std::wstring_view path) noexcept
{
    path             = TrimTrailingSlash(path);
    const size_t pos = path.find_last_of(L'/');
    if (pos == std::wstring_view::npos)
    {
        return path;
    }
    return path.substr(pos + 1u);
}

[[nodiscard]] std::wstring ParentPath(std::wstring_view path) noexcept
{
    path             = TrimTrailingSlash(path);
    const size_t pos = path.find_last_of(L'/');
    if (pos == std::wstring_view::npos || pos == 0)
    {
        return L"/";
    }

    std::wstring parent(path.substr(0, pos));
    if (parent.empty())
    {
        parent = L"/";
    }
    parent.push_back(L'/');
    return parent;
}

[[nodiscard]] std::wstring JoinPluginPath(std::wstring_view folder, std::wstring_view leaf) noexcept
{
    std::wstring base = EnsureTrailingSlash(folder);
    base.append(leaf.data(), leaf.size());
    return base;
}

[[nodiscard]] std::wstring BuildDisplayPath(Protocol protocol, std::wstring_view pluginPath) noexcept
{
    const std::wstring normalized = NormalizePluginPath(pluginPath);

    std::wstring_view scheme = L"sftp";
    switch (protocol)
    {
        case Protocol::Ftp: scheme = L"ftp"; break;
        case Protocol::Sftp: scheme = L"sftp"; break;
        case Protocol::Scp: scheme = L"scp"; break;
        case Protocol::Imap: scheme = L"imap"; break;
    }

    std::wstring out;
    out.reserve(scheme.size() + 1u + normalized.size());
    out.append(scheme);
    out.push_back(L':');
    out.append(normalized);
    return out;
}

[[nodiscard]] std::wstring EnsureTrailingSlashDisplay(std::wstring_view path) noexcept
{
    std::wstring out(path);
    if (! out.empty() && out.back() != L'/')
    {
        out.push_back(L'/');
    }
    return out;
}

[[nodiscard]] std::wstring JoinDisplayPath(std::wstring_view folder, std::wstring_view leaf) noexcept
{
    std::wstring base(folder);
    if (! base.empty() && base.back() != L'/')
    {
        base.push_back(L'/');
    }
    base.append(leaf.data(), leaf.size());
    return base;
}

[[nodiscard]] std::string JoinRemotePath(std::string_view basePathUtf8, std::wstring_view pluginPath) noexcept
{
    const std::wstring normalizedPlugin = NormalizePluginPath(pluginPath);
    const std::string pluginUtf8        = EscapeUrlPath(normalizedPlugin);
    if (pluginUtf8.empty())
    {
        return "/";
    }

    std::string base = basePathUtf8.empty() ? std::string("/") : std::string(basePathUtf8);
    if (base.empty())
    {
        base = "/";
    }

    if (base.back() == '/' && base.size() > 1)
    {
        base.pop_back();
    }

    if (base == "/")
    {
        return pluginUtf8;
    }

    if (pluginUtf8 == "/")
    {
        return std::format("{}/", base);
    }

    return std::format("{}{}", base, pluginUtf8);
}

[[nodiscard]] std::wstring JoinPluginPathWide(std::wstring_view basePath, std::wstring_view pluginPath) noexcept
{
    std::wstring base = NormalizePluginPath(basePath);
    if (base.empty())
    {
        base = L"/";
    }

    if (base.size() > 1 && base.back() == L'/')
    {
        base.pop_back();
    }

    std::wstring plugin = NormalizePluginPath(pluginPath);
    if (plugin.empty())
    {
        plugin = L"/";
    }

    if (base == L"/")
    {
        return plugin;
    }

    if (plugin == L"/")
    {
        std::wstring out(base);
        out.push_back(L'/');
        return out;
    }

    std::wstring out(base);
    out.append(plugin);
    return out;
}

[[nodiscard]] std::string BuildUrl(const ConnectionInfo& conn, std::wstring_view pluginPath, bool forDirectory, bool forCommand) noexcept
{
    const std::string_view scheme = forCommand ? ProtocolSchemeForCommands(conn.protocol) : ProtocolSchemeForTransfer(conn.protocol);
    if (conn.host.empty())
    {
        return {};
    }

    std::string authority = conn.host;
    if (conn.port.has_value() && conn.port.value() != 0)
    {
        const bool alreadyHasPort = authority.find(':') != std::string::npos && (authority.empty() || authority.front() != '[');
        if (! alreadyHasPort)
        {
            authority = std::format("{}:{}", authority, conn.port.value());
        }
    }

    std::string remotePath = JoinRemotePath(conn.basePath, pluginPath);
    if (remotePath.empty())
    {
        remotePath = "/";
    }

    if (forDirectory && remotePath.back() != '/')
    {
        remotePath.push_back('/');
    }

    return std::format("{}://{}{}", scheme, authority, remotePath);
}

HRESULT EnsureCurlInitialized() noexcept
{
    static std::once_flag initOnce;
    static std::atomic<HRESULT> initResult{E_FAIL};

    std::call_once(initOnce,
                   []() noexcept
                   {
                       const CURLcode code = curl_global_init(CURL_GLOBAL_DEFAULT);
                       initResult.store(code == CURLE_OK ? S_OK : E_FAIL, std::memory_order_release);
                   });

    return initResult.load(std::memory_order_acquire);
}

namespace
{
struct CurlShareContext final
{
    CurlShareContext() = default;

    CurlShareContext(const CurlShareContext&)            = delete;
    CurlShareContext(CurlShareContext&&)                 = delete;
    CurlShareContext& operator=(const CurlShareContext&) = delete;
    CurlShareContext& operator=(CurlShareContext&&)      = delete;

    ~CurlShareContext() = default;

    std::array<std::mutex, static_cast<size_t>(CURL_LOCK_DATA_LAST)> locks{};
    CURLSH* share = nullptr;
};

void CurlShareLock(CURL* /*handle*/, curl_lock_data data, curl_lock_access /*access*/, void* userptr) noexcept
{
    auto* ctx = static_cast<CurlShareContext*>(userptr);
    if (! ctx)
    {
        return;
    }

    const size_t index = static_cast<size_t>(data);
    if (index >= ctx->locks.size())
    {
        return;
    }

    ctx->locks[index].lock();
}

void CurlShareUnlock(CURL* /*handle*/, curl_lock_data data, void* userptr) noexcept
{
    auto* ctx = static_cast<CurlShareContext*>(userptr);
    if (! ctx)
    {
        return;
    }

    const size_t index = static_cast<size_t>(data);
    if (index >= ctx->locks.size())
    {
        return;
    }

    ctx->locks[index].unlock();
}

[[nodiscard]] CURLSH* GetCurlShareHandle() noexcept
{
    static std::once_flag initOnce;
    static CurlShareContext ctx{};

    std::call_once(initOnce,
                   [&]() noexcept
                   {
                       if (FAILED(EnsureCurlInitialized()))
                       {
                           return;
                       }

                       CURLSH* share = curl_share_init();
                       if (! share)
                       {
                           return;
                       }

                       curl_share_setopt(share, CURLSHOPT_USERDATA, &ctx);
                       curl_share_setopt(share, CURLSHOPT_LOCKFUNC, CurlShareLock);
                       curl_share_setopt(share, CURLSHOPT_UNLOCKFUNC, CurlShareUnlock);

                       curl_share_setopt(share, CURLSHOPT_SHARE, CURL_LOCK_DATA_DNS);
                       curl_share_setopt(share, CURLSHOPT_SHARE, CURL_LOCK_DATA_SSL_SESSION);
                       curl_share_setopt(share, CURLSHOPT_SHARE, CURL_LOCK_DATA_CONNECT);

                       ctx.share = share;
                   });

    return ctx.share;
}
} // namespace

[[nodiscard]] HRESULT HResultFromCurl(CURLcode code) noexcept
{
#pragma warning(push)
// we don't want to explicitly manage all Curl options
// enum 'xx' is not explicitly handled by a case label
#pragma warning(disable : 4061)
    switch (code)
    {
        case CURLE_OK: return S_OK;
        case CURLE_ABORTED_BY_CALLBACK: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        case CURLE_UNSUPPORTED_PROTOCOL: return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        case CURLE_URL_MALFORMAT: return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        case CURLE_REMOTE_FILE_NOT_FOUND: return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        case CURLE_COULDNT_RESOLVE_PROXY: return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
        case CURLE_COULDNT_RESOLVE_HOST: return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
        case CURLE_COULDNT_CONNECT: return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
        case CURLE_LOGIN_DENIED: return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
        case CURLE_REMOTE_ACCESS_DENIED: return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        case CURLE_QUOTE_ERROR: return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
        case CURLE_SEND_ERROR: return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        case CURLE_RECV_ERROR: return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        case CURLE_GOT_NOTHING: return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
        case CURLE_WEIRD_SERVER_REPLY: return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
        case CURLE_SSL_CONNECT_ERROR: return SEC_E_ILLEGAL_MESSAGE;
        case CURLE_PEER_FAILED_VERIFICATION: return SEC_E_CERT_UNKNOWN;
        case CURLE_SSL_CACERT_BADFILE: return SEC_E_CERT_UNKNOWN;
        case CURLE_SSL_CERTPROBLEM: return SEC_E_CERT_UNKNOWN;
        case CURLE_SSL_ISSUER_ERROR: return SEC_E_CERT_UNKNOWN;
        case CURLE_OPERATION_TIMEDOUT: return HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT);
        default: return E_FAIL;
    }
#pragma warning(pop)
}

void ApplyCommonCurlOptions(CURL* curl, const ConnectionInfo& conn, const FileSystemOptions* options, bool forUpload) noexcept
{
    if (! curl)
    {
        return;
    }

    if (CURLSH* share = GetCurlShareHandle())
    {
        curl_easy_setopt(curl, CURLOPT_SHARE, share);
    }

    curl_easy_setopt(curl, CURLOPT_NOSIGNAL, 1L);
    curl_easy_setopt(curl, CURLOPT_TCP_NODELAY, 1L);
    curl_easy_setopt(curl, CURLOPT_TCP_KEEPALIVE, 1L);
    curl_easy_setopt(curl, CURLOPT_TCP_KEEPIDLE, 60L);
    curl_easy_setopt(curl, CURLOPT_TCP_KEEPINTVL, 60L);

    if (conn.protocol == Protocol::Ftp)
    {
        constexpr long kBufferBytes = 512L * 1024L;
        curl_easy_setopt(curl, CURLOPT_BUFFERSIZE, kBufferBytes);
        if (forUpload)
        {
            curl_easy_setopt(curl, CURLOPT_UPLOAD_BUFFERSIZE, kBufferBytes);
        }
    }
    curl_easy_setopt(curl, CURLOPT_FTP_USE_EPSV, conn.ftpUseEpsv ? 1L : 0L);
    curl_easy_setopt(curl, CURLOPT_SSL_OPTIONS, CURLSSLOPT_NATIVE_CA);
    curl_easy_setopt(curl, CURLOPT_PROXY_SSL_OPTIONS, CURLSSLOPT_NATIVE_CA);
    if (conn.ignoreSslTrust)
    {
        curl_easy_setopt(curl, CURLOPT_SSL_VERIFYPEER, 0L);
        curl_easy_setopt(curl, CURLOPT_SSL_VERIFYHOST, 2L);
        curl_easy_setopt(curl, CURLOPT_PROXY_SSL_VERIFYPEER, 0L);
        curl_easy_setopt(curl, CURLOPT_PROXY_SSL_VERIFYHOST, 2L);
    }

    if (! conn.user.empty())
    {
        curl_easy_setopt(curl, CURLOPT_USERNAME, conn.user.c_str());
    }
    if (! conn.password.empty())
    {
        curl_easy_setopt(curl, CURLOPT_PASSWORD, conn.password.c_str());
    }

    if (conn.connectTimeoutMs != 0)
    {
        curl_easy_setopt(
            curl, CURLOPT_CONNECTTIMEOUT_MS, static_cast<long>(std::min(conn.connectTimeoutMs, static_cast<unsigned long>(std::numeric_limits<long>::max()))));
    }

    if (conn.operationTimeoutMs != 0)
    {
        curl_easy_setopt(
            curl, CURLOPT_TIMEOUT_MS, static_cast<long>(std::min(conn.operationTimeoutMs, static_cast<unsigned long>(std::numeric_limits<long>::max()))));
    }

    // Avoid hanging forever on stalled connections (no progress).
    constexpr long kLowSpeedLimitBytesPerSecond = 1L;
    constexpr long kLowSpeedTimeSecondsDefault  = 60L;

    long lowSpeedTimeSeconds = kLowSpeedTimeSecondsDefault;
    if (conn.operationTimeoutMs != 0)
    {
        const unsigned long opSec = conn.operationTimeoutMs / 1000u;
        lowSpeedTimeSeconds       = opSec == 0 ? 1L : static_cast<long>((std::min)(opSec, static_cast<unsigned long>(kLowSpeedTimeSecondsDefault)));
    }

    curl_easy_setopt(curl, CURLOPT_LOW_SPEED_LIMIT, kLowSpeedLimitBytesPerSecond);
    curl_easy_setopt(curl, CURLOPT_LOW_SPEED_TIME, lowSpeedTimeSeconds);

#ifdef CURLOPT_FTP_RESPONSE_TIMEOUT
    if (conn.protocol == Protocol::Ftp)
    {
        curl_easy_setopt(curl, CURLOPT_FTP_RESPONSE_TIMEOUT, lowSpeedTimeSeconds);
    }
#endif

    if (! conn.sshPrivateKey.empty())
    {
        curl_easy_setopt(curl, CURLOPT_SSH_PRIVATE_KEYFILE, conn.sshPrivateKey.c_str());
    }
    if (! conn.sshPublicKey.empty())
    {
        curl_easy_setopt(curl, CURLOPT_SSH_PUBLIC_KEYFILE, conn.sshPublicKey.c_str());
    }
    if (! conn.sshKeyPassphrase.empty())
    {
        curl_easy_setopt(curl, CURLOPT_KEYPASSWD, conn.sshKeyPassphrase.c_str());
    }
    if (! conn.sshKnownHosts.empty())
    {
        curl_easy_setopt(curl, CURLOPT_SSH_KNOWNHOSTS, conn.sshKnownHosts.c_str());
    }

    const unsigned __int64 limit = options ? options->bandwidthLimitBytesPerSecond : 0;
    if (limit > 0)
    {
        if (forUpload)
        {
            curl_easy_setopt(curl, CURLOPT_MAX_SEND_SPEED_LARGE, static_cast<curl_off_t>(limit));
        }
        else
        {
            curl_easy_setopt(curl, CURLOPT_MAX_RECV_SPEED_LARGE, static_cast<curl_off_t>(limit));
        }
    }
}

size_t CurlWriteToString(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
{
    if (! ptr || ! userdata)
    {
        return 0;
    }

    const size_t total = size * nmemb;
    auto* out          = static_cast<std::string*>(userdata);
    out->append(static_cast<const char*>(ptr), total);
    return total;
}

size_t CurlWriteToFile(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
{
    if (! ptr || ! userdata)
    {
        return 0;
    }

    const size_t total = size * nmemb;
    if (total == 0)
    {
        return 0;
    }

    HANDLE file = reinterpret_cast<HANDLE>(userdata);
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return 0;
    }

    DWORD written    = 0;
    const DWORD take = total > static_cast<size_t>(std::numeric_limits<DWORD>::max()) ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(total);
    if (! WriteFile(file, ptr, take, &written, nullptr))
    {
        return 0;
    }

    return written;
}

size_t CurlReadFromFile(char* buffer, size_t size, size_t nitems, void* instream) noexcept
{
    if (! buffer || ! instream || size == 0 || nitems == 0)
    {
        return 0;
    }

    HANDLE file = reinterpret_cast<HANDLE>(instream);
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return 0;
    }

    const size_t want = size * nitems;
    if (want == 0)
    {
        return 0;
    }

    DWORD read       = 0;
    const DWORD take = want > static_cast<size_t>(std::numeric_limits<DWORD>::max()) ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(want);
    if (! ReadFile(file, buffer, take, &read, nullptr))
    {
        return CURL_READFUNC_ABORT;
    }

    return read;
}
} // namespace FileSystemCurlInternal

namespace FileSystemCurlInternal
{
[[nodiscard]] const wchar_t* CopyArenaString(FileSystemArena* arena, std::wstring_view text) noexcept
{
    if (! arena)
    {
        return nullptr;
    }

    const size_t length = text.size();
    if (length > (std::numeric_limits<unsigned long>::max)() / sizeof(wchar_t) - 1u)
    {
        return nullptr;
    }

    const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
    auto* buffer              = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, bytes, static_cast<unsigned long>(alignof(wchar_t))));
    if (! buffer)
    {
        return nullptr;
    }

    if (length > 0)
    {
        std::memcpy(buffer, text.data(), length * sizeof(wchar_t));
    }

    buffer[length] = L'\0';
    return buffer;
}

[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept
{
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    LARGE_INTEGER zero{};
    if (SetFilePointerEx(file, zero, nullptr, FILE_BEGIN) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, unsigned __int64& out) noexcept
{
    out = 0;

    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    LARGE_INTEGER size{};
    if (GetFileSizeEx(file, &size) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (size.QuadPart < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
    }

    out = static_cast<unsigned __int64>(size.QuadPart);
    return S_OK;
}

[[nodiscard]] wil::unique_hfile CreateTemporaryDeleteOnCloseFile() noexcept
{
    wchar_t tempPath[MAX_PATH]{};
    const DWORD tempPathLen = GetTempPathW(ARRAYSIZE(tempPath), tempPath);
    if (tempPathLen == 0 || tempPathLen >= ARRAYSIZE(tempPath))
    {
        return {};
    }

    wchar_t tempName[MAX_PATH]{};
    if (GetTempFileNameW(tempPath, L"rsc", 0, tempName) == 0)
    {
        return {};
    }

    HANDLE handle = CreateFileW(tempName,
                                GENERIC_READ | GENERIC_WRITE,
                                FILE_SHARE_READ,
                                nullptr,
                                CREATE_ALWAYS,
                                FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE | FILE_FLAG_SEQUENTIAL_SCAN,
                                nullptr);

    return wil::unique_hfile(handle);
}
} // namespace FileSystemCurlInternal

namespace FileSystemCurlInternal
{
[[nodiscard]] unsigned __int64 ClampCurlOffToUInt64(curl_off_t value) noexcept
{
    if (value <= 0)
    {
        return 0;
    }

    constexpr curl_off_t max = static_cast<curl_off_t>((std::numeric_limits<unsigned __int64>::max)());
    if (value > max)
    {
        return (std::numeric_limits<unsigned __int64>::max)();
    }

    return static_cast<unsigned __int64>(value);
}

[[nodiscard]] unsigned __int64 SaturatingAddToAtomic(std::atomic<unsigned __int64>& value, unsigned __int64 delta) noexcept
{
    if (delta == 0)
    {
        return value.load(std::memory_order_acquire);
    }

    unsigned __int64 current = value.load(std::memory_order_relaxed);
    for (;;)
    {
        unsigned __int64 next = current;
        if (current > (std::numeric_limits<unsigned __int64>::max)() - delta)
        {
            next = (std::numeric_limits<unsigned __int64>::max)();
        }
        else
        {
            next = current + delta;
        }

        if (value.compare_exchange_weak(current, next, std::memory_order_release, std::memory_order_relaxed))
        {
            return next;
        }
    }
}

int CurlXferInfo(void* clientp, curl_off_t dltotal, curl_off_t dlnow, curl_off_t ultotal, curl_off_t ulnow) noexcept
{
    auto* ctx = static_cast<TransferProgressContext*>(clientp);
    if (! ctx || ! ctx->progress)
    {
        return 0;
    }

    const unsigned __int64 nowTick = GetTickCount64();

    const unsigned __int64 phaseTotal = ClampCurlOffToUInt64(ctx->isUpload ? ultotal : dltotal);
    const unsigned __int64 phaseNow   = ClampCurlOffToUInt64(ctx->isUpload ? ulnow : dlnow);

    if (ctx->itemTotalBytes == 0 && phaseTotal > 0)
    {
        ctx->itemTotalBytes = phaseTotal;
    }

    unsigned __int64 itemDone  = phaseNow;
    unsigned __int64 itemTotal = phaseTotal;

    if (ctx->scaleForCopy && ctx->itemTotalBytes > 0)
    {
        itemTotal                   = ctx->itemTotalBytes;
        const unsigned __int64 half = itemTotal / 2u;
        if (! ctx->scaleForCopySecond)
        {
            itemDone = std::min(half, phaseNow / 2u);
        }
        else
        {
            const unsigned __int64 extra = (itemTotal & 1u) != 0 ? 1u : 0u;
            if (phaseNow >= itemTotal)
            {
                itemDone = itemTotal;
            }
            else
            {
                itemDone = std::min(itemTotal, half + (phaseNow + extra) / 2u);
            }
        }
    }

    unsigned __int64 wireDone = phaseNow;
    if (ctx->scaleForCopy && ctx->itemTotalBytes > 0)
    {
        const unsigned __int64 offset = ctx->scaleForCopySecond ? ctx->itemTotalBytes : 0;
        wireDone = offset > (std::numeric_limits<unsigned __int64>::max)() - phaseNow ? (std::numeric_limits<unsigned __int64>::max)() : (offset + phaseNow);
    }

    unsigned __int64 overall = 0;
    if (ctx->concurrentOverallBytes)
    {
        const unsigned __int64 delta = wireDone >= ctx->lastConcurrentWireDone ? (wireDone - ctx->lastConcurrentWireDone) : 0;
        ctx->lastConcurrentWireDone  = wireDone;
        overall                      = SaturatingAddToAtomic(*ctx->concurrentOverallBytes, delta);
    }
    else
    {
        overall = ctx->baseCompletedBytes > (std::numeric_limits<unsigned __int64>::max)() - wireDone ? (std::numeric_limits<unsigned __int64>::max)()
                                                                                                      : (ctx->baseCompletedBytes + wireDone);
    }

    // Cancellation check (even if we don't report progress this tick).
    if (ctx->progress->callback && (ctx->lastCancelTick == 0 || (nowTick - ctx->lastCancelTick) >= ctx->cancelIntervalMs))
    {
        ctx->lastCancelTick    = nowTick;
        const HRESULT cancelHr = ctx->progress->CheckCancel();
        if (FAILED(cancelHr))
        {
            ctx->abortHr = cancelHr;
            return 1;
        }
    }

    // Progress reporting (throttled).
    const bool shouldReport = ctx->progress->callback && (ctx->lastReportTick == 0 || (nowTick - ctx->lastReportTick) >= ctx->reportIntervalMs ||
                                                          (phaseTotal > 0 && phaseNow >= phaseTotal));

    if (shouldReport && (itemDone != ctx->lastReportedItemDone || overall != ctx->lastReportedOverall))
    {
        ctx->lastReportTick       = nowTick;
        ctx->lastReportedItemDone = itemDone;
        ctx->lastReportedOverall  = overall;

        const HRESULT hr = ctx->progress->ReportProgressWithCompletedBytes(overall, itemTotal, itemDone, ctx->sourcePath, ctx->destinationPath);
        if (FAILED(hr))
        {
            ctx->abortHr = hr;
            return 1;
        }
    }

    // Soft bandwidth limiting with Sleep in the progress callback (enables dynamic updates from host).
    const unsigned __int64 limit = ctx->progress->options.bandwidthLimitBytesPerSecond;
    if (limit > 0 && ctx->throttleStartTick != 0)
    {
        const unsigned __int64 elapsedMs = nowTick - ctx->throttleStartTick;
        if (elapsedMs > 0 && phaseNow > 0)
        {
            const double expectedMs = (static_cast<double>(phaseNow) * 1000.0) / static_cast<double>(limit);
            const double elapsed    = static_cast<double>(elapsedMs);
            if (expectedMs > elapsed)
            {
                const double sleepMs = expectedMs - elapsed;
                if (sleepMs >= 1.0)
                {
                    Sleep(static_cast<DWORD>((std::min)(sleepMs, 200.0)));
                }
            }
        }
    }

    return 0;
}

[[nodiscard]] std::string RemotePathForCommand(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept
{
    std::string remote = JoinRemotePath(conn.basePath, pluginPath);
    while (remote.size() > 1u && remote.back() == '/')
    {
        remote.pop_back();
    }
    return remote.empty() ? std::string("/") : remote;
}

[[nodiscard]] HRESULT CurlPerformList(const ConnectionInfo& conn, std::wstring_view pluginPath, std::string& outListing) noexcept
{
    outListing.clear();

    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    unique_curl_easy curl{curl_easy_init()};
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, true, true);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &outListing);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    char errorBuffer[CURL_ERROR_SIZE]{};
    curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

    ApplyCommonCurlOptions(curl.get(), conn, nullptr, false);

    const CURLcode code = curl_easy_perform(curl.get());
    if (code != CURLE_OK)
    {
        long responseCode = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &responseCode);

        Debug::Error(L"curl list failed protocol={} url='{}' user='{}' connProfile={} conn='{}' id='{}' authMode='{}' savePassword={} requireHello={} "
                     L"passwordPresent={} sshKeyPresent={} sshPassphrasePresent={} knownHostsPresent={} responseCode={} curlCode={} ({}) error='{}'",
                     ProtocolToDisplay(conn.protocol),
                     Utf16FromUtf8(url),
                     Utf16FromUtf8(conn.user),
                     conn.fromConnectionManagerProfile ? 1 : 0,
                     conn.connectionName.empty() ? L"(none)" : conn.connectionName,
                     conn.connectionId,
                     conn.connectionAuthMode,
                     conn.connectionSavePassword ? 1 : 0,
                     conn.connectionRequireHello ? 1 : 0,
                     conn.password.empty() ? 0 : 1,
                     conn.sshPrivateKey.empty() ? 0 : 1,
                     conn.sshKeyPassphrase.empty() ? 0 : 1,
                     conn.sshKnownHosts.empty() ? 0 : 1,
                     responseCode,
                     static_cast<unsigned long>(code),
                     Utf16FromUtf8(curl_easy_strerror(code)),
                     Utf16FromUtf8(errorBuffer));

        if (conn.protocol == Protocol::Ftp && responseCode == 530 && conn.password.empty() && ! conn.user.empty() && conn.user != "anonymous")
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        }
    }
    return HResultFromCurl(code);
}

[[nodiscard]] HRESULT CurlPerformQuote(const ConnectionInfo& conn, const std::vector<std::string>& commands) noexcept
{
    if (commands.empty())
    {
        return S_OK;
    }

    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    unique_curl_easy curl{curl_easy_init()};
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, L"/", true, true);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    std::string sink;

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &sink);
    curl_easy_setopt(curl.get(), CURLOPT_DIRLISTONLY, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    unique_curl_slist list;
    for (const auto& cmd : commands)
    {
        curl_slist* appended = curl_slist_append(list.get(), cmd.c_str());
        if (! appended)
        {
            return E_OUTOFMEMORY;
        }
        list.release();
        list.reset(appended);
    }

    curl_easy_setopt(curl.get(), CURLOPT_QUOTE, list.get());

    char errorBuffer[CURL_ERROR_SIZE]{};
    curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

    ApplyCommonCurlOptions(curl.get(), conn, nullptr, false);

    const CURLcode code = curl_easy_perform(curl.get());
    if (code != CURLE_OK)
    {
        long responseCode = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &responseCode);

        Debug::Error(L"curl quote failed protocol={} url='{}' user='{}' connProfile={} conn='{}' id='{}' authMode='{}' savePassword={} requireHello={} "
                     L"passwordPresent={} sshKeyPresent={} sshPassphrasePresent={} knownHostsPresent={} responseCode={} curlCode={} ({}) error='{}'",
                     ProtocolToDisplay(conn.protocol),
                     Utf16FromUtf8(url),
                     Utf16FromUtf8(conn.user),
                     conn.fromConnectionManagerProfile ? 1 : 0,
                     conn.connectionName.empty() ? L"(none)" : conn.connectionName,
                     conn.connectionId,
                     conn.connectionAuthMode,
                     conn.connectionSavePassword ? 1 : 0,
                     conn.connectionRequireHello ? 1 : 0,
                     conn.password.empty() ? 0 : 1,
                     conn.sshPrivateKey.empty() ? 0 : 1,
                     conn.sshKeyPassphrase.empty() ? 0 : 1,
                     conn.sshKnownHosts.empty() ? 0 : 1,
                     responseCode,
                     static_cast<unsigned long>(code),
                     Utf16FromUtf8(curl_easy_strerror(code)),
                     Utf16FromUtf8(errorBuffer));

        if (conn.protocol == Protocol::Ftp && responseCode == 530 && conn.password.empty() && ! conn.user.empty() && conn.user != "anonymous")
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        }
    }
    return HResultFromCurl(code);
}

[[nodiscard]] HRESULT CurlDownloadToFile(
    const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file, const FileSystemOptions* options, TransferProgressContext* progressCtx) noexcept
{
    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    unique_curl_easy curl{curl_easy_init()};
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, false, false);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToFile);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, file);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    if (progressCtx)
    {
        progressCtx->Begin();
        curl_easy_setopt(curl.get(), CURLOPT_XFERINFOFUNCTION, CurlXferInfo);
        curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, progressCtx);
        curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);
        ApplyCommonCurlOptions(curl.get(), conn, nullptr, false);
    }
    else
    {
        curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 1L);
        ApplyCommonCurlOptions(curl.get(), conn, options, false);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code == CURLE_ABORTED_BY_CALLBACK && progressCtx && FAILED(progressCtx->abortHr))
    {
        return progressCtx->abortHr;
    }

    return HResultFromCurl(code);
}

[[nodiscard]] HRESULT CurlUploadFromFile(const ConnectionInfo& conn,
                                         std::wstring_view pluginPath,
                                         HANDLE file,
                                         unsigned __int64 sizeBytes,
                                         const FileSystemOptions* options,
                                         TransferProgressContext* progressCtx) noexcept
{
    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    unique_curl_easy curl{curl_easy_init()};
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, false, false);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_UPLOAD, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_READFUNCTION, CurlReadFromFile);
    curl_easy_setopt(curl.get(), CURLOPT_READDATA, file);
    curl_easy_setopt(curl.get(),
                     CURLOPT_INFILESIZE_LARGE,
                     static_cast<curl_off_t>(std::min(sizeBytes, static_cast<unsigned __int64>((std::numeric_limits<curl_off_t>::max)()))));
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    if (progressCtx)
    {
        progressCtx->Begin();
        curl_easy_setopt(curl.get(), CURLOPT_XFERINFOFUNCTION, CurlXferInfo);
        curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, progressCtx);
        curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);
        ApplyCommonCurlOptions(curl.get(), conn, nullptr, true);
    }
    else
    {
        curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 1L);
        ApplyCommonCurlOptions(curl.get(), conn, options, true);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code == CURLE_ABORTED_BY_CALLBACK && progressCtx && FAILED(progressCtx->abortHr))
    {
        return progressCtx->abortHr;
    }

    return HResultFromCurl(code);
}
} // namespace FileSystemCurlInternal

// FileSystemCurl

FileSystemCurl::FileSystemCurl(FileSystemCurlProtocol protocol, IHost* host) : _protocol(protocol)
{
    switch (_protocol)
    {
        case FileSystemCurlProtocol::Ftp:
            _metaData.id          = kPluginIdFtp;
            _metaData.shortId     = kPluginShortIdFtp;
            _metaData.name        = kPluginNameFtp;
            _metaData.description = kPluginDescriptionFtp;
            break;
        case FileSystemCurlProtocol::Sftp:
            _metaData.id          = kPluginIdSftp;
            _metaData.shortId     = kPluginShortIdSftp;
            _metaData.name        = kPluginNameSftp;
            _metaData.description = kPluginDescriptionSftp;
            break;
        case FileSystemCurlProtocol::Scp:
            _metaData.id          = kPluginIdScp;
            _metaData.shortId     = kPluginShortIdScp;
            _metaData.name        = kPluginNameScp;
            _metaData.description = kPluginDescriptionScp;
            break;
        case FileSystemCurlProtocol::Imap:
            _metaData.id          = kPluginIdImap;
            _metaData.shortId     = kPluginShortIdImap;
            _metaData.name        = kPluginNameImap;
            _metaData.description = kPluginDescriptionImap;
            break;
    }
    _metaData.author  = kPluginAuthor;
    _metaData.version = kPluginVersion;

    _configurationJson = "{}";

    _driveFileSystem = _metaData.shortId ? _metaData.shortId : L"";

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE FileSystemCurl::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemCurl::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    switch (_protocol)
    {
        case FileSystemCurlProtocol::Ftp: *schemaJsonUtf8 = kSchemaJsonFtp; break;
        case FileSystemCurlProtocol::Sftp: *schemaJsonUtf8 = kSchemaJsonSftp; break;
        case FileSystemCurlProtocol::Scp: *schemaJsonUtf8 = kSchemaJsonScp; break;
        case FileSystemCurlProtocol::Imap: *schemaJsonUtf8 = kSchemaJsonImap; break;
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::lock_guard lock(_stateMutex);

    _settings = {};

    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson = configurationJsonUtf8;

    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(_configurationJson.data(), _configurationJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return S_OK;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return S_OK;
    }

    const auto defaultHost = TryGetJsonString(root, "defaultHost");
    if (defaultHost.has_value())
    {
        _settings.defaultHost = defaultHost.value();
    }

    const auto defaultPort = TryGetJsonUInt(root, "defaultPort");
    if (defaultPort.has_value())
    {
        const uint64_t value = defaultPort.value();
        if (value <= 65535u)
        {
            _settings.defaultPort = static_cast<unsigned int>(value);
        }
    }

    const auto defaultUser = TryGetJsonString(root, "defaultUser");
    if (defaultUser.has_value())
    {
        _settings.defaultUser = defaultUser.value();
    }

    const auto defaultPassword = TryGetJsonString(root, "defaultPassword");
    if (defaultPassword.has_value())
    {
        _settings.defaultPassword = defaultPassword.value();
    }

    const auto defaultBasePath = TryGetJsonString(root, "defaultBasePath");
    if (defaultBasePath.has_value())
    {
        _settings.defaultBasePath = defaultBasePath.value();
        if (_settings.defaultBasePath.empty())
        {
            _settings.defaultBasePath = L"/";
        }
    }

    const auto connectTimeoutMs = TryGetJsonUInt(root, "connectTimeoutMs");
    if (connectTimeoutMs.has_value())
    {
        _settings.connectTimeoutMs = static_cast<unsigned long>(std::min<uint64_t>(connectTimeoutMs.value(), (std::numeric_limits<unsigned long>::max)()));
    }

    const auto operationTimeoutMs = TryGetJsonUInt(root, "operationTimeoutMs");
    if (operationTimeoutMs.has_value())
    {
        _settings.operationTimeoutMs = static_cast<unsigned long>(std::min<uint64_t>(operationTimeoutMs.value(), (std::numeric_limits<unsigned long>::max)()));
    }

    const auto ignoreSslTrust = TryGetJsonBool(root, "ignoreSslTrust");
    if (ignoreSslTrust.has_value())
    {
        _settings.ignoreSslTrust = ignoreSslTrust.value();
    }

    const auto ftpUseEpsv = TryGetJsonBool(root, "ftpUseEpsv");
    if (ftpUseEpsv.has_value())
    {
        _settings.ftpUseEpsv = ftpUseEpsv.value();
    }

    const auto sshPrivateKey = TryGetJsonString(root, "sshPrivateKey");
    if (sshPrivateKey.has_value())
    {
        _settings.sshPrivateKey = sshPrivateKey.value();
    }

    const auto sshPublicKey = TryGetJsonString(root, "sshPublicKey");
    if (sshPublicKey.has_value())
    {
        _settings.sshPublicKey = sshPublicKey.value();
    }

    const auto sshKeyPassphrase = TryGetJsonString(root, "sshKeyPassphrase");
    if (sshKeyPassphrase.has_value())
    {
        _settings.sshKeyPassphrase = sshKeyPassphrase.value();
    }

    const auto sshKnownHosts = TryGetJsonString(root, "sshKnownHosts");
    if (sshKnownHosts.has_value())
    {
        _settings.sshKnownHosts = sshKnownHosts.value();
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const bool hasNonDefault = ! _configurationJson.empty() && _configurationJson != "{}";
    *pSomethingToSave        = hasNonDefault ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const Settings settings = _settings;

    std::wstring connectionHeader;
    const std::wstring scheme = _metaData.shortId ? _metaData.shortId : L"";
    if (! settings.defaultHost.empty())
    {
        if (! settings.defaultUser.empty())
        {
            connectionHeader = std::format(L"{}://{}@{}", scheme, settings.defaultUser, settings.defaultHost);
        }
        else
        {
            connectionHeader = std::format(L"{}://{}", scheme, settings.defaultHost);
        }
    }
    else
    {
        connectionHeader = L"(no default host)";
    }

    _menuEntries.clear();
    _menuEntryView.clear();

    MenuEntry header;
    header.flags = NAV_MENU_ITEM_FLAG_HEADER;
    header.label = _metaData.name ? _metaData.name : L"";
    _menuEntries.push_back(std::move(header));

    MenuEntry connection;
    connection.flags = NAV_MENU_ITEM_FLAG_HEADER;
    connection.label = std::move(connectionHeader);
    _menuEntries.push_back(std::move(connection));

    MenuEntry separator;
    separator.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
    _menuEntries.push_back(std::move(separator));

    MenuEntry root;
    root.label = L"/";
    root.path  = L"/";
    _menuEntries.push_back(std::move(root));

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

HRESULT STDMETHODCALLTYPE FileSystemCurl::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback != nullptr ? cookie : nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    Settings settings;
    FileSystemCurlProtocol protocol = FileSystemCurlProtocol::Sftp;
    wil::com_ptr<IHostConnections> hostConnections;
    const wchar_t* scheme = nullptr;
    {
        std::lock_guard lock(_stateMutex);
        settings        = _settings;
        protocol        = _protocol;
        hostConnections = _hostConnections;
        scheme          = _metaData.shortId ? _metaData.shortId : L"";
    }

    std::wstring driveDisplayName;
    std::wstring driveFileSystem;
    ResolvedLocation resolved{};
    const std::wstring_view pluginPath = (path != nullptr && path[0] != L'\0') ? std::wstring_view(path) : std::wstring_view(L"/");
    const HRESULT resolveHr            = ResolveLocation(protocol, settings, pluginPath, hostConnections.get(), false, resolved);

    if (SUCCEEDED(resolveHr))
    {
        const std::wstring host = Utf16FromUtf8(resolved.connection.host);
        const std::wstring user = Utf16FromUtf8(resolved.connection.user);

        std::wstring authority = host;
        if (resolved.connection.port.has_value() && resolved.connection.port.value() != 0u)
        {
            authority = std::format(L"{}:{}", host, resolved.connection.port.value());
        }

        const bool showUser = ! user.empty() && ! (protocol == FileSystemCurlProtocol::Ftp && user == L"anonymous");
        if (showUser)
        {
            driveDisplayName = std::format(L"{}://{}@{}", scheme, user, authority);
        }
        else
        {
            driveDisplayName = std::format(L"{}://{}", scheme, authority);
        }

        driveFileSystem = scheme;
    }
    else
    {
        driveDisplayName = std::format(L"{}:// (not configured)", scheme);
        driveFileSystem  = scheme;
    }

    {
        std::lock_guard lock(_stateMutex);
        _driveDisplayName = std::move(driveDisplayName);
        _driveFileSystem  = std::move(driveFileSystem);

        _driveInfo = {};
        if (! _driveDisplayName.empty())
        {
            _driveInfo.flags       = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_DISPLAY_NAME);
            _driveInfo.displayName = _driveDisplayName.c_str();
        }

        if (! _driveFileSystem.empty())
        {
            _driveInfo.flags      = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
            _driveInfo.fileSystem = _driveFileSystem.c_str();
        }

        *info = _driveInfo;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetDriveMenuItems(const wchar_t* /*path*/, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::ExecuteDriveMenuCommand(unsigned int /*commandId*/, const wchar_t* /*path*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    switch (_protocol)
    {
        case FileSystemCurlProtocol::Ftp: *jsonUtf8 = kCapabilitiesJsonFtp; break;
        case FileSystemCurlProtocol::Sftp: *jsonUtf8 = kCapabilitiesJsonSftp; break;
        case FileSystemCurlProtocol::Scp: *jsonUtf8 = kCapabilitiesJsonScp; break;
        case FileSystemCurlProtocol::Imap: *jsonUtf8 = kCapabilitiesJsonImap; break;
    }
    return S_OK;
}

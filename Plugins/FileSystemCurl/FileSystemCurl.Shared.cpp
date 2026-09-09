#include "CurlProcessRuntime.h"
#include "FileSystemCurl.Internal.h"
#include "FileSystemCurlResources.h"
#include "HandleIo.h"
#include "StringConversion.h"
#include "YyjsonHelpers.h"

using namespace FileSystemCurlInternal;

extern HINSTANCE g_hInstance;

namespace
{
std::atomic<unsigned long> g_fileSystemCurlInstanceCount{0};
std::atomic<bool> g_fileSystemCurlShutdownRequested{false};

[[nodiscard]] const wchar_t* LocalizedPluginName(FileSystemCurlProtocol protocol) noexcept
{
    switch (protocol)
    {
        case FileSystemCurlProtocol::Ftp:
        {
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_NAME);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Sftp:
        {
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_NAME);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Scp:
        {
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_NAME);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Imap:
        {
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_NAME);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_NAME);
    return fallback.c_str();
}

[[nodiscard]] const wchar_t* LocalizedPluginDescription(FileSystemCurlProtocol protocol) noexcept
{
    switch (protocol)
    {
        case FileSystemCurlProtocol::Ftp:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Sftp:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Scp:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemCurlProtocol::Imap:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_DESCRIPTION);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_DESCRIPTION);
    return fallback.c_str();
}
} // namespace

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
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* obj, const char* key) noexcept
{
    const Common::Json::MemberResult<std::wstring> value = Common::Json::GetUtf16StringMemberStrict(obj, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<std::wstring>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* obj, const char* key) noexcept
{
    const Common::Json::MemberResult<uint64_t> value = Common::Json::GetUInt64Member(obj, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<uint64_t>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* obj, const char* key) noexcept
{
    const Common::Json::MemberResult<bool> value =
        Common::Json::GetBoolMember(obj, key, Common::Json::MemberRequirement::Optional, Common::Json::BooleanIntegerPolicy::AllowZeroAndNonzero);
    return value.HasValue() ? std::optional<bool>{value.value} : std::nullopt;
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
    return _packedBuffer.GetBuffer(ppFileInfo);
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetBufferSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetBufferSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetAllocatedSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetAllocatedSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::GetCount(unsigned long* pCount) noexcept
{
    return _packedBuffer.GetCount(pCount);
}

HRESULT STDMETHODCALLTYPE FilesInformationCurl::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    return _packedBuffer.Get(index, ppEntry);
}

HRESULT FilesInformationCurl::BuildFromEntries(std::vector<Entry> entries) noexcept
{
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

    return _packedBuffer.Build(entries,
                               [](const Entry& source, FileInfo& entry) noexcept
    {
        entry.FileAttributes = source.attributes;
        entry.FileIndex      = source.fileIndex;
        entry.EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry.AllocationSize = static_cast<__int64>(source.sizeBytes);
        entry.CreationTime   = source.creationTime;
        entry.LastAccessTime = source.lastAccessTime;
        entry.LastWriteTime  = source.lastWriteTime;
        entry.ChangeTime     = source.changeTime;
    });
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

[[nodiscard]] char ToLowerAscii(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

[[nodiscard]] bool EqualsAsciiIgnoreCase(std::string_view left, std::string_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    for (size_t index = 0; index < left.size(); ++index)
    {
        if (ToLowerAscii(left[index]) != ToLowerAscii(right[index]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool TryParseUnsignedInRange(std::string_view token, unsigned int minimum, unsigned int maximum, unsigned int& value) noexcept
{
    if (token.empty())
    {
        return false;
    }

    unsigned int parsed     = 0;
    const auto [ptr, error] = std::from_chars(token.data(), token.data() + token.size(), parsed);
    if (error != std::errc{} || ptr != token.data() + token.size() || parsed < minimum || parsed > maximum)
    {
        return false;
    }

    value = parsed;
    return true;
}

[[nodiscard]] bool TryParseMonthToken(std::string_view token, unsigned int& month) noexcept
{
    static constexpr std::array<std::string_view, 12> kMonths = {{"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}};

    for (size_t index = 0; index < kMonths.size(); ++index)
    {
        if (EqualsAsciiIgnoreCase(token, kMonths[index]))
        {
            month = static_cast<unsigned int>(index + 1);
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool TryLocalSystemTimeToFileTime(const SYSTEMTIME& localTime, __int64& fileTime) noexcept
{
    SYSTEMTIME utcTime{};
    if (! TzSpecificLocalTimeToSystemTime(nullptr, &localTime, &utcTime))
    {
        return false;
    }

    FILETIME rawFileTime{};
    if (! SystemTimeToFileTime(&utcTime, &rawFileTime))
    {
        return false;
    }

    ULARGE_INTEGER value{};
    value.LowPart  = rawFileTime.dwLowDateTime;
    value.HighPart = rawFileTime.dwHighDateTime;
    fileTime       = static_cast<__int64>(value.QuadPart);
    return true;
}

[[nodiscard]] __int64 CurrentSystemFileTime() noexcept
{
    FILETIME rawFileTime{};
    GetSystemTimeAsFileTime(&rawFileTime);

    ULARGE_INTEGER value{};
    value.LowPart  = rawFileTime.dwLowDateTime;
    value.HighPart = rawFileTime.dwHighDateTime;
    return static_cast<__int64>(value.QuadPart);
}

[[nodiscard]] bool TryParseUnixListTimestamp(std::string_view monthToken,
                                             std::string_view dayToken,
                                             std::string_view timeOrYearToken,
                                             __int64& fileTime) noexcept
{
    unsigned int month = 0;
    unsigned int day   = 0;
    if (! TryParseMonthToken(monthToken, month) || ! TryParseUnsignedInRange(dayToken, 1, 31, day))
    {
        return false;
    }

    SYSTEMTIME localTime{};
    localTime.wMonth = static_cast<WORD>(month);
    localTime.wDay   = static_cast<WORD>(day);

    const size_t colon = timeOrYearToken.find(':');
    if (colon == std::string_view::npos)
    {
        unsigned int year = 0;
        if (! TryParseUnsignedInRange(timeOrYearToken, 1601, 9999, year))
        {
            return false;
        }
        localTime.wYear = static_cast<WORD>(year);
        return TryLocalSystemTimeToFileTime(localTime, fileTime);
    }

    unsigned int hour   = 0;
    unsigned int minute = 0;
    if (! TryParseUnsignedInRange(timeOrYearToken.substr(0, colon), 0, 23, hour) || ! TryParseUnsignedInRange(timeOrYearToken.substr(colon + 1), 0, 59, minute))
    {
        return false;
    }

    SYSTEMTIME currentLocalTime{};
    GetLocalTime(&currentLocalTime);

    localTime.wYear   = currentLocalTime.wYear;
    localTime.wHour   = static_cast<WORD>(hour);
    localTime.wMinute = static_cast<WORD>(minute);

    if (! TryLocalSystemTimeToFileTime(localTime, fileTime))
    {
        return false;
    }

    constexpr __int64 kFileTimeTicksPerDay = 24LL * 60LL * 60LL * 10000000LL;
    const __int64 latestReasonableTime     = CurrentSystemFileTime() + kFileTimeTicksPerDay;
    if (fileTime > latestReasonableTime && localTime.wYear > 1601)
    {
        --localTime.wYear;
        return TryLocalSystemTimeToFileTime(localTime, fileTime);
    }

    return true;
}

[[nodiscard]] bool TryParseDosListTimestamp(std::string_view dateToken, std::string_view timeToken, __int64& fileTime) noexcept
{
    if (dateToken.size() != 8 || (dateToken[2] != '-' && dateToken[2] != '/') || dateToken[5] != dateToken[2] || timeToken.size() < 6)
    {
        return false;
    }

    unsigned int month     = 0;
    unsigned int day       = 0;
    unsigned int shortYear = 0;
    if (! TryParseUnsignedInRange(dateToken.substr(0, 2), 1, 12, month) || ! TryParseUnsignedInRange(dateToken.substr(3, 2), 1, 31, day) ||
        ! TryParseUnsignedInRange(dateToken.substr(6, 2), 0, 99, shortYear))
    {
        return false;
    }

    const std::string_view meridiem = timeToken.substr(timeToken.size() - 2);
    const bool isAm                 = EqualsAsciiIgnoreCase(meridiem, "AM");
    const bool isPm                 = EqualsAsciiIgnoreCase(meridiem, "PM");
    if (! isAm && ! isPm)
    {
        return false;
    }

    const std::string_view clock = timeToken.substr(0, timeToken.size() - 2);
    const size_t colon           = clock.find(':');
    if (colon == std::string_view::npos)
    {
        return false;
    }

    unsigned int hour   = 0;
    unsigned int minute = 0;
    if (! TryParseUnsignedInRange(clock.substr(0, colon), 1, 12, hour) || ! TryParseUnsignedInRange(clock.substr(colon + 1), 0, 59, minute))
    {
        return false;
    }

    if (isPm && hour != 12)
    {
        hour += 12;
    }
    else if (isAm && hour == 12)
    {
        hour = 0;
    }

    const unsigned int fullYear = shortYear >= 70 ? 1900 + shortYear : 2000 + shortYear;

    SYSTEMTIME localTime{};
    localTime.wYear   = static_cast<WORD>(fullYear);
    localTime.wMonth  = static_cast<WORD>(month);
    localTime.wDay    = static_cast<WORD>(day);
    localTime.wHour   = static_cast<WORD>(hour);
    localTime.wMinute = static_cast<WORD>(minute);

    return TryLocalSystemTimeToFileTime(localTime, fileTime);
}

[[nodiscard]] bool TryParseUnixListLine(std::string_view line,
                                        FilesInformationCurl::Entry& out,
                                        bool includeDotEntries          = false,
                                        std::wstring_view timestampLeaf = {}) noexcept
{
    if (line.size() < 2)
    {
        return false;
    }

    if (line.rfind("total ", 0) == 0)
    {
        return false;
    }

    const char type      = line[0];
    const bool directory = type == 'd';
    const bool symlink   = type == 'l';
    const bool regular   = type == '-';
    // Device, socket, fifo, and door names are real directory entries. Accept
    // them as Unix so they cannot fall through to the DOS size-token heuristic.
    const bool special = type == 'c' || type == 'b' || type == 's' || type == 'p' || type == 'D';
    if (! directory && ! symlink && ! regular && ! special)
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
    const auto sizeTok       = nextToken(pos);
    const auto monthTok      = nextToken(pos);
    const auto dayTok        = nextToken(pos);
    const auto timeOrYearTok = nextToken(pos);

    if (! sizeTok.has_value() || ! monthTok.has_value() || ! dayTok.has_value() || ! timeOrYearTok.has_value())
    {
        return false;
    }

    skipSpaces(pos);
    if (pos >= line.size())
    {
        return false;
    }

    std::string_view namePart = line.substr(pos);
    if (! includeDotEntries && IsDotOrDotDotName(namePart))
    {
        return false;
    }

    const size_t arrow = namePart.find(" -> ");
    if (type == 'l' && arrow != std::string_view::npos)
    {
        namePart = namePart.substr(0, arrow);
    }

    uint64_t sizeBytes = 0;
    bool sizeParsed    = false;
    {
        const auto sv        = sizeTok.value();
        const auto [ptr, ec] = std::from_chars(sv.data(), sv.data() + sv.size(), sizeBytes);
        sizeParsed           = (ec == std::errc{} && ptr == sv.data() + sv.size());
    }

    // Repeated strict lookups reuse the current row's name allocation. Reset
    // every metadata field; a previous occupant's size/times are not evidence.
    std::wstring reusableName = std::move(out.name);
    reusableName.clear();
    out      = {};
    out.name = std::move(reusableName);
    if (directory)
    {
        out.attributes = FILE_ATTRIBUTE_DIRECTORY;
    }
    else if (special)
    {
        out.attributes = FILE_ATTRIBUTE_DEVICE;
    }
    else
    {
        out.attributes = FILE_ATTRIBUTE_NORMAL;
    }
    out.sizeBytes = sizeBytes;
    // The size column is meaningful for regular files and symlink payloads.
    // Device/socket/fifo rows often carry major,minor or a dummy 0.
    out.sizeKnown = (regular || symlink) && sizeParsed;
    if (! Common::Strings::TryUtf16FromUtf8Strict(namePart, out.name))
    {
        return false;
    }
    __int64 modifiedTime = 0;
    if ((timestampLeaf.empty() || out.name == timestampLeaf) &&
        TryParseUnixListTimestamp(monthTok.value(), dayTok.value(), timeOrYearTok.value(), modifiedTime))
    {
        out.lastWriteTime = modifiedTime;
    }
    return ! out.name.empty();
}

[[nodiscard]] bool TryParseDosListLine(std::string_view line,
                                       FilesInformationCurl::Entry& out,
                                       bool includeDotEntries          = false,
                                       std::wstring_view timestampLeaf = {}) noexcept
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

    size_t pos         = 0;
    const auto dateTok = nextToken(pos);
    const auto timeTok = nextToken(pos);
    if (! dateTok.has_value() || ! timeTok.has_value())
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
    if (! includeDotEntries && IsDotOrDotDotName(namePart))
    {
        return false;
    }

    // Preserve only name storage, never metadata from a previous parsed row.
    std::wstring reusableName = std::move(out.name);
    reusableName.clear();
    out      = {};
    out.name = std::move(reusableName);
    if (EqualsAsciiIgnoreCase(sizeOrDir.value(), "<DIR>"))
    {
        out.attributes = FILE_ATTRIBUTE_DIRECTORY;
        out.sizeBytes  = 0;
    }
    else
    {
        const auto sv        = sizeOrDir.value();
        uint64_t parsed      = 0;
        const auto [ptr, ec] = std::from_chars(sv.data(), sv.data() + sv.size(), parsed);
        if (ec != std::errc{} || ptr != sv.data() + sv.size())
        {
            return false;
        }
        out.attributes = FILE_ATTRIBUTE_NORMAL;
        out.sizeBytes  = parsed;
        out.sizeKnown  = true;
    }

    if (! Common::Strings::TryUtf16FromUtf8Strict(namePart, out.name))
    {
        return false;
    }
    __int64 modifiedTime = 0;
    if (! TryParseDosListTimestamp(dateTok.value(), timeTok.value(), modifiedTime))
    {
        // Date and time tokens are the DOS dialect discriminator. A numeric
        // third token alone is not enough — Unix socket/device rows would
        // otherwise become phantom files named from the leftover columns.
        return false;
    }
    if (timestampLeaf.empty() || out.name == timestampLeaf)
    {
        out.lastWriteTime = modifiedTime;
    }
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
        return OrdinalString::EqualsNoCase(a, b);
    };

    out.connection.ftpUseEpsv         = settings.ftpUseEpsv;
    out.connection.connectTimeoutMs   = settings.connectTimeoutMs;
    out.connection.operationTimeoutMs = settings.operationTimeoutMs;
    out.connection.ignoreSslTrust     = settings.ignoreSslTrust;
    out.connection.sshPrivateKey      = Utf8FromUtf16(settings.sshPrivateKey);
    out.connection.sshPublicKey       = Utf8FromUtf16(settings.sshPublicKey);
    out.connection.sshKeyPassphrase   = Utf8FromUtf16(settings.sshKeyPassphrase);
    out.connection.sshKnownHosts      = Utf8FromUtf16(settings.sshKnownHosts);

    const unsigned int pluginCopyMoveMax = protocol == Protocol::Imap ? 1u : std::clamp(settings.copyMoveMaxConcurrency, 1u, 8u);
    const unsigned int pluginDeleteMax   = protocol == Protocol::Imap ? 1u : std::clamp(settings.deleteMaxConcurrency, 1u, 8u);

    const auto protocolKey = [&](Protocol p) noexcept -> std::wstring_view
    {
        switch (p)
        {
            case Protocol::Ftp: return L"ftp";
            case Protocol::Sftp: return L"sftp";
            case Protocol::Scp: return L"scp";
            case Protocol::Imap: return L"imap";
        }
        return L"unknown";
    };

    auto finalizeConnection = [&](unsigned int copyMoveOverride, unsigned int deleteOverride) noexcept
    {
        out.connection.effectiveCopyMoveMaxConcurrency = pluginCopyMoveMax;
        out.connection.effectiveDeleteMaxConcurrency   = pluginDeleteMax;

        if (copyMoveOverride != 0)
        {
            const unsigned int clamped                     = std::clamp(copyMoveOverride, 1u, 8u);
            out.connection.effectiveCopyMoveMaxConcurrency = std::min(out.connection.effectiveCopyMoveMaxConcurrency, clamped);
        }

        if (deleteOverride != 0)
        {
            const unsigned int clamped                   = std::clamp(deleteOverride, 1u, 64u);
            out.connection.effectiveDeleteMaxConcurrency = std::min(out.connection.effectiveDeleteMaxConcurrency, std::min(clamped, 8u));
        }

        if (out.connection.fromConnectionManagerProfile)
        {
            if (! out.connection.connectionId.empty())
            {
                out.connection.limiterKey = out.connection.connectionId;
                return;
            }

            if (! out.connection.connectionName.empty())
            {
                out.connection.limiterKey = std::format(L"{}|@conn:{}", protocolKey(protocol), out.connection.connectionName);
                return;
            }
        }

        const unsigned int portOut = out.connection.port.has_value() ? out.connection.port.value() : 0u;
        out.connection.limiterKey =
            std::format(L"{}|{}|{}|{}", protocolKey(protocol), Utf16FromUtf8(out.connection.host), portOut, Utf16FromUtf8(out.connection.user));
    };

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

        Common::Json::UniqueDocument docOwner{doc};

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

        auto getUIntFromObj = [&](yyjson_val* obj, const char* key, unsigned int& outUInt) noexcept -> bool
        {
            if (! obj || ! yyjson_is_obj(obj))
            {
                return false;
            }

            yyjson_val* v = yyjson_obj_get(obj, key);
            if (! v)
            {
                return false;
            }

            if (yyjson_is_uint(v))
            {
                const uint64_t value = yyjson_get_uint(v);
                outUInt              = static_cast<unsigned int>(std::min<uint64_t>(value, 0xFFFFFFFFull));
                return true;
            }

            if (yyjson_is_sint(v))
            {
                const int64_t value = yyjson_get_sint(v);
                if (value < 0)
                {
                    return false;
                }
                outUInt = static_cast<unsigned int>(std::min<int64_t>(value, 0x7FFFFFFFll));
                return true;
            }

            return false;
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
                             Redaction::ForLog(connPath),
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
                             Redaction::ForLog(connPath));
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
                             Redaction::ForLog(connPath));
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
                    Redaction::ForLog(connPath));
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
                        Common::Json::UniqueDocument refreshedOwner{refreshedDoc};

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
                                 Redaction::ForLog(connPath),
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
                                 Redaction::ForLog(connPath));
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
                                     Redaction::ForLog(connPath));
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

        unsigned int copyMoveOverride = 0;
        unsigned int deleteOverride   = 0;
        if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
        {
            static_cast<void>(getUIntFromObj(extra, "copyMoveMaxConcurrency", copyMoveOverride));
            static_cast<void>(getUIntFromObj(extra, "deleteMaxConcurrency", deleteOverride));
        }

        finalizeConnection(copyMoveOverride, deleteOverride);

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

        finalizeConnection(0, 0);
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

    finalizeConnection(0, 0);
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

namespace
{
[[nodiscard]] HRESULT EnsureSharedCurlRuntime() noexcept;
}

HRESULT EnsureCurlInitialized() noexcept
{
    return EnsureSharedCurlRuntime();
}

namespace
{
[[nodiscard]] Common::CurlRuntime::ProcessLease& GetCurlRuntimeLease() noexcept
{
    static Common::CurlRuntime::ProcessLease lease;
    return lease;
}

[[nodiscard]] HRESULT EnsureSharedCurlRuntime() noexcept
{
    if (g_fileSystemCurlShutdownRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }
    return GetCurlRuntimeLease().Acquire([]() noexcept -> HRESULT { return curl_global_init(CURL_GLOBAL_DEFAULT) == CURLE_OK ? S_OK : E_FAIL; });
}

void CleanupSharedCurlRuntime() noexcept
{
    curl_global_cleanup();
}

struct CurlShareContext final
{
    CurlShareContext() = default;

    CurlShareContext(const CurlShareContext&)            = delete;
    CurlShareContext(CurlShareContext&&)                 = delete;
    CurlShareContext& operator=(const CurlShareContext&) = delete;
    CurlShareContext& operator=(CurlShareContext&&)      = delete;

    ~CurlShareContext() = default;

    std::mutex lifecycleMutex;
    std::array<std::mutex, static_cast<size_t>(CURL_LOCK_DATA_LAST)> locks{};
    CURLSH* share = nullptr;
};

[[nodiscard]] CurlShareContext& GetCurlShareContext() noexcept
{
    static CurlShareContext context{};
    return context;
}

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
    CurlShareContext& ctx = GetCurlShareContext();

    std::call_once(initOnce,
                   [&]() noexcept
    {
        if (FAILED(EnsureSharedCurlRuntime()))
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
        // R0f-Curl-OR2: the connection pool is deliberately NOT shared. Easy handles run
        // concurrently on several host worker threads (parallel copies) and libcurl's pooled
        // connection matching raced with another thread's teardown (access violation in
        // url_match_destination, three crash dumps). Each pooled easy handle keeps its own
        // connections; the handle pool hands a handle to one thread at a time, so reuse stays
        // thread-confined. DNS and TLS session sharing remain safe under the lock callbacks.

        std::lock_guard lock(ctx.lifecycleMutex);
        ctx.share = share;
    });

    std::lock_guard lock(ctx.lifecycleMutex);
    return ctx.share;
}

void CleanupCurlShareHandle() noexcept
{
    CurlShareContext& ctx = GetCurlShareContext();
    std::lock_guard lock(ctx.lifecycleMutex);
    if (ctx.share)
    {
        static_cast<void>(curl_share_cleanup(ctx.share));
        ctx.share = nullptr;
    }
}
} // namespace

CurlEasyPool::BorrowedHandle CurlEasyPool::Borrow(std::wstring_view limiterKey) noexcept
{
    std::wstring key{limiterKey};
    std::vector<Handles> cleanup;
    Handles handles;

    const uint64_t now = GetTickCount64();
    {
        std::scoped_lock lock(_mutex);
        if (_shutdownRequested)
        {
            return {};
        }
        ++_activeBorrowCount;
        EvictExpired(now, cleanup);

        auto it = _idle.find(key);
        if (it != _idle.end() && ! it->second.empty())
        {
            IdleEntry entry = std::move(it->second.back());
            it->second.pop_back();
            if (it->second.empty())
            {
                _idle.erase(it);
            }
            handles = std::move(entry.handles);
        }
    }

    if (! handles.easy)
    {
        handles.easy.reset(curl_easy_init());
        if (! handles.easy)
        {
            CancelBorrow();
            return {};
        }
    }
    return BorrowedHandle(this, std::move(key), std::move(handles));
}

void CurlEasyPool::ReturnHandle(std::wstring key, Handles handles) noexcept
{
    if (! handles.easy)
    {
        return;
    }

    // Cleanup may invoke FTP/IMAP callbacks. Drop per-call callback/data pointers
    // before publishing an idle handle or destroying it on shutdown/overflow;
    // resetting only on the next borrow leaves expired caller storage reachable.
    // curl_easy_reset preserves connection, DNS and TLS-session reuse.
    curl_easy_reset(handles.easy.get());
    if (handles.cursorEasy)
    {
        curl_easy_reset(handles.cursorEasy.get());
    }

    const uint64_t now = GetTickCount64();

    {
        std::scoped_lock lock(_mutex);
        if (_activeBorrowCount > 0u)
        {
            --_activeBorrowCount;
        }
        if (_shutdownRequested)
        {
            return;
        }

        auto& vec = _idle[std::move(key)];
        if (vec.size() >= kMaxIdlePerConnection)
        {
            return; // drop handle — pool is full for this connection
        }
        vec.push_back(IdleEntry(std::move(handles), now));
    }
}

void CurlEasyPool::CancelBorrow() noexcept
{
    std::scoped_lock lock(_mutex);
    if (_activeBorrowCount > 0u)
    {
        --_activeBorrowCount;
    }
}

void CurlEasyPool::BeginShutdown() noexcept
{
    std::vector<Handles> cleanup;
    {
        std::scoped_lock lock(_mutex);
        _shutdownRequested = true;
        for (auto& [key, entries] : _idle)
        {
            for (IdleEntry& entry : entries)
            {
                if (entry.handles.easy)
                {
                    cleanup.push_back(std::move(entry.handles));
                }
            }
        }
        _idle.clear();
    }
}

bool CurlEasyPool::CanUnloadNow() noexcept
{
    std::scoped_lock lock(_mutex);
    return _shutdownRequested && _activeBorrowCount == 0u && _idle.empty();
}

void CurlEasyPool::EvictExpired(uint64_t now, std::vector<Handles>& cleanup) noexcept
{
    for (auto it = _idle.begin(); it != _idle.end();)
    {
        auto& vec = it->second;
        for (auto& entry : vec)
        {
            if ((now - entry.returnedAtMs) >= kIdleExpiryMs && entry.handles.easy)
            {
                cleanup.push_back(std::move(entry.handles));
            }
        }
        std::erase_if(vec, [&](const IdleEntry& e) { return (now - e.returnedAtMs) >= kIdleExpiryMs; });
        if (vec.empty())
        {
            it = _idle.erase(it);
        }
        else
        {
            ++it;
        }
    }
}

[[nodiscard]] CurlEasyPool& GetCurlEasyPool() noexcept
{
    static CurlEasyPool pool;
    return pool;
}

namespace
{
void TryCompleteFileSystemCurlShutdown() noexcept
{
    if (! g_fileSystemCurlShutdownRequested.load(std::memory_order_acquire) || g_fileSystemCurlInstanceCount.load(std::memory_order_acquire) != 0u)
    {
        return;
    }

    CurlEasyPool& pool = GetCurlEasyPool();
    pool.BeginShutdown();
    if (! pool.CanUnloadNow())
    {
        return;
    }

    CleanupCurlShareHandle();
    Common::CurlRuntime::ProcessLease& lease = GetCurlRuntimeLease();
    if (lease.IsAcquired())
    {
        static_cast<void>(lease.Release(CleanupSharedCurlRuntime));
    }
}
} // namespace

void BeginFileSystemCurlShutdown() noexcept
{
    g_fileSystemCurlShutdownRequested.store(true, std::memory_order_release);
    ShutdownSharedCopyMoveJobScheduler();
    GetCurlEasyPool().BeginShutdown();
    TryCompleteFileSystemCurlShutdown();
}

bool CanUnloadFileSystemCurlNow() noexcept
{
    TryCompleteFileSystemCurlShutdown();
    return g_fileSystemCurlShutdownRequested.load(std::memory_order_acquire) && g_fileSystemCurlInstanceCount.load(std::memory_order_acquire) == 0u &&
           GetCurlEasyPool().CanUnloadNow() && ! GetCurlRuntimeLease().IsAcquired();
}

bool CanCreateFileSystemCurl() noexcept
{
    return ! g_fileSystemCurlShutdownRequested.load(std::memory_order_acquire);
}

#if defined(_DEBUG)
HRESULT RunDebugCurlRuntimeProbe() noexcept
{
    const HRESULT initializeHr = EnsureSharedCurlRuntime();
    if (FAILED(initializeHr))
    {
        return initializeHr;
    }

    auto handle = GetCurlEasyPool().Borrow(L"observatory/runtime-probe");
    return handle ? S_OK : E_OUTOFMEMORY;
}
#endif

namespace
{
thread_local const FileSystemOptions* t_curlOperationOptions = nullptr;

// R0f-Curl: libcurl invokes the progress callback at least once per second for the whole transfer,
// including while a control command waits for the server, so polling the host's operation control
// here gives every FTP/SFTP/SCP call a quiet point of about one second after Cancel or a deadline.
int CurlOperationControlXferInfo(void* clientp, curl_off_t, curl_off_t, curl_off_t, curl_off_t) noexcept
{
    return FAILED(FileSystemCheckOperationControl(static_cast<const FileSystemOptions*>(clientp))) ? 1 : 0;
}
} // namespace

const FileSystemOptions* CurlCurrentOperationOptions() noexcept
{
    return t_curlOperationOptions;
}

CurlOperationOptionsScope::CurlOperationOptionsScope(const FileSystemOptions* options) noexcept : _previous(t_curlOperationOptions)
{
    if (options != nullptr)
    {
        t_curlOperationOptions = options;
    }
}

CurlOperationOptionsScope::~CurlOperationOptionsScope()
{
    t_curlOperationOptions = _previous;
}

long CurlLowSpeedTimeSeconds(unsigned long operationTimeoutMs) noexcept
{
    constexpr long kLowSpeedTimeSecondsDefault = 60L;
    if (operationTimeoutMs == 0u)
    {
        return kLowSpeedTimeSecondsDefault;
    }
    const unsigned long opSec = operationTimeoutMs / 1000u;
    return opSec == 0u ? 1L : static_cast<long>((std::min)(opSec, static_cast<unsigned long>(kLowSpeedTimeSecondsDefault)));
}

// The provider-owned bound after which a call whose server stopped answering returns on its own:
// the connect timeout or the low-speed abort, whichever is longer. It is the route's
// providerWatchdogTimeoutMs; cooperative cancellation through the progress callback is faster.
uint32_t CurlProviderWatchdogTimeoutMs(unsigned long connectTimeoutMs, unsigned long operationTimeoutMs) noexcept
{
    constexpr unsigned long kDefaultConnectTimeoutMs = 10'000ul;
    const unsigned long lowSpeedMs                   = static_cast<unsigned long>(CurlLowSpeedTimeSeconds(operationTimeoutMs)) * 1000ul;
    const unsigned long bound                        = (std::max)(connectTimeoutMs == 0u ? kDefaultConnectTimeoutMs : connectTimeoutMs, lowSpeedMs);
    return static_cast<uint32_t>((std::min)(bound, static_cast<unsigned long>(std::numeric_limits<uint32_t>::max())));
}

[[nodiscard]] HRESULT HResultFromCurl(CURLcode code) noexcept
{
#pragma warning(push)
// we don't want to explicitly manage all Curl options
// enum 'xx' is not explicitly handled by a case label
#pragma warning(disable : 4061)
    switch (code)
    {
        case CURLE_OK: return S_OK;
        case CURLE_ABORTED_BY_CALLBACK:
        {
            // Report the host's own verdict (cancel versus deadline) when the abort came from the
            // operation-control progress callback.
            const HRESULT controlHr = t_curlOperationOptions != nullptr ? FileSystemCheckOperationControl(t_curlOperationOptions) : S_OK;
            return FAILED(controlHr) ? controlHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
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
        default:
            // Keep the provider's existing failure contract, but retain the raw
            // transport code for diagnosis without URLs, credentials or payloads.
            Debug::Perf::Emit(L"FileOps.Curl.UnmappedTransportFailure", L"curl-code", 0u, static_cast<uint64_t>(code), 0u, E_FAIL);
            return E_FAIL;
    }
#pragma warning(pop)
}

namespace
{
void RecordCurlConnectionFailure(CURL* curl, CURLcode code, const wchar_t* phase) noexcept
{
    if (code != CURLE_COULDNT_CONNECT)
    {
        return;
    }

    // Transport-local diagnostics: capture the easy handle's OS error before
    // reset/return. Common HRESULT helpers do not own this libcurl state. Never
    // retain URLs, profiles, control replies, error text, or credentials here.
    long osError                = 0;
    const bool osErrorAvailable = curl_easy_getinfo(curl, CURLINFO_OS_ERRNO, &osError) == CURLE_OK && osError > 0;
    long localPort              = 0;
    long remotePort             = 0;
    long responseCode           = 0;
    static_cast<void>(curl_easy_getinfo(curl, CURLINFO_LOCAL_PORT, &localPort));
    static_cast<void>(curl_easy_getinfo(curl, CURLINFO_PRIMARY_PORT, &remotePort));
    static_cast<void>(curl_easy_getinfo(curl, CURLINFO_RESPONSE_CODE, &responseCode));
    // Ports may be nonpositive/unavailable or describe an established FTP
    // control connection, not the failed secondary data socket. Preserve the
    // reported values without inventing endpoint attribution.
    const std::wstring detail =
        std::format(L"phase={};osErrorAvailable={};localPort={};remotePort={};responseCode={}", phase, osErrorAvailable, localPort, remotePort, responseCode);
    Debug::Perf::Emit(L"FileOps.Curl.ConnectionFailure",
                      detail.c_str(),
                      0u,
                      static_cast<uint64_t>(code),
                      osErrorAvailable ? static_cast<uint64_t>(osError) : 0u,
                      HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED));
#ifdef ENABLE_TESTS
    if (osErrorAvailable && osError == WSAEADDRINUSE)
    {
        // A fresh wildcard bind distinguishes contemporaneous allocation failure
        // from a failed connection tuple. It is only an observation, never retry
        // authority or proof of global exhaustion; other threads can release ports.
        const auto started       = std::chrono::steady_clock::now();
        unsigned short port      = 0u;
        const HRESULT probeHr    = ProbeCurlEphemeralBindForSelfTest(port);
        const uint64_t elapsedUs = Debug::Perf::ElapsedUs(started);
        Debug::Perf::Emit(L"FileOps.Curl.ConnectionAllocationProbe", phase, elapsedUs, static_cast<uint64_t>(osError), port, probeHr);
        for (const bool bindFirst : {false, true})
        {
            const auto connectStarted        = std::chrono::steady_clock::now();
            unsigned short localProbePort    = 0u;
            const wchar_t* stage             = nullptr;
            const HRESULT connectHr          = ProbeCurlLoopbackConnectForSelfTest(bindFirst, localProbePort, stage);
            const std::wstring connectDetail = std::format(L"phase={};explicitBind={};stage={}", phase, bindFirst, stage);
            Debug::Perf::Emit(L"FileOps.Curl.LoopbackConnectProbe",
                              connectDetail.c_str(),
                              Debug::Perf::ElapsedUs(connectStarted),
                              bindFirst ? 1u : 0u,
                              localProbePort,
                              connectHr);
        }
    }
#endif
}
} // namespace

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
    const long lowSpeedTimeSeconds              = CurlLowSpeedTimeSeconds(conn.operationTimeoutMs);

    curl_easy_setopt(curl, CURLOPT_LOW_SPEED_LIMIT, kLowSpeedLimitBytesPerSecond);
    curl_easy_setopt(curl, CURLOPT_LOW_SPEED_TIME, lowSpeedTimeSeconds);

    if (options != nullptr)
    {
        // R0f-Curl: cooperative cancel/deadline polling for this transfer. Callers that install
        // their own progress callback pass nullptr here and keep it.
        curl_easy_setopt(curl, CURLOPT_XFERINFOFUNCTION, CurlOperationControlXferInfo);
        curl_easy_setopt(curl, CURLOPT_XFERINFODATA, const_cast<FileSystemOptions*>(options));
        curl_easy_setopt(curl, CURLOPT_NOPROGRESS, 0L);
    }

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

    const uint64_t limit = options ? options->bandwidthLimitBytesPerSecond : 0;
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

namespace
{
struct CurlListParseContext final
{
    explicit CurlListParseContext(std::vector<FilesInformationCurl::Entry>& outEntries) noexcept : _outEntries(&outEntries)
    {
    }

    CurlListParseContext(const CurlListParseContext&)            = delete;
    CurlListParseContext& operator=(const CurlListParseContext&) = delete;
    CurlListParseContext(CurlListParseContext&&)                 = delete;
    CurlListParseContext& operator=(CurlListParseContext&&)      = delete;

    void AppendBytes(const char* bytes, size_t bytesCount) noexcept
    {
        if (bytes == nullptr || bytesCount == 0)
        {
            return;
        }

        _pending.append(bytes, bytesCount);

        size_t start = 0;
        while (true)
        {
            const size_t end = _pending.find('\n', start);
            if (end == std::string::npos)
            {
                break;
            }

            ProcessLine(std::string_view(_pending).substr(start, end - start));
            start = end + 1u;
        }

        if (start > 0)
        {
            _pending.erase(0, start);
        }
    }

    void ProcessLine(std::string_view line) noexcept
    {
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }
        if (line.empty())
        {
            return;
        }

        FilesInformationCurl::Entry entry{};
        if (TryParseUnixListLine(line, entry) || TryParseDosListLine(line, entry))
        {
            _anyParsed = true;
            _fallbackNamesUtf8.clear();
            _outEntries->push_back(std::move(entry));
            return;
        }

        if (_anyParsed)
        {
            return;
        }

        std::string trimmed = TrimAscii(line);
        if (trimmed.empty() || IsDotOrDotDotName(trimmed))
        {
            return;
        }

        _fallbackNamesUtf8.push_back(std::move(trimmed));
    }

    void FinalizePendingLine() noexcept
    {
        if (_pending.empty())
        {
            return;
        }

        const std::string pending = std::move(_pending);
        _pending.clear();
        ProcessLine(pending);
    }

    [[nodiscard]] bool HasFallbackNames() const noexcept
    {
        return ! _anyParsed && ! _fallbackNamesUtf8.empty();
    }

    [[nodiscard]] const std::vector<std::string>& GetFallbackNamesUtf8() const noexcept
    {
        return _fallbackNamesUtf8;
    }

    void Reset() noexcept
    {
        _pending.clear();
        _anyParsed = false;
        _fallbackNamesUtf8.clear();
        _outEntries->clear();
    }

private:
    std::string _pending;
    std::vector<FilesInformationCurl::Entry>* _outEntries = nullptr;
    bool _anyParsed                                       = false;
    std::vector<std::string> _fallbackNamesUtf8;
};

size_t CurlWriteToListParser(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
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

    auto* ctx = static_cast<CurlListParseContext*>(userdata);
    ctx->AppendBytes(static_cast<const char*>(ptr), total);

    return total;
}
} // namespace

#ifdef ENABLE_TESTS
std::optional<CurlFailedConnectPorts> CurlDirectoryCursor::ParseFailedConnectTraceForSelfTest(curl_infotype type, std::string_view text) noexcept
{
    // Test-local pinned-libcurl diagnostic grammar, not a generic network parser.
    // Only our IPv4 loopback fixture is in scope. Borrow text, retain only ports,
    // and never copy addresses, error suffixes, headers or protocol payloads.
    constexpr std::string_view prefix = "connect to 127.0.0.1 port ";
    if (type != CURLINFO_TEXT || ! text.starts_with(prefix))
    {
        return std::nullopt;
    }
    text.remove_prefix(prefix.size());
    CurlFailedConnectPorts ports;
    const auto remote = std::from_chars(text.data(), text.data() + (std::min)(text.size(), size_t{6u}), ports.remote);
    if (remote.ec != std::errc{} || ports.remote < 1 || ports.remote > 65535)
    {
        return std::nullopt;
    }
    text.remove_prefix(static_cast<size_t>(remote.ptr - text.data()));
    constexpr std::array<std::string_view, 3u> localPrefixes{{" from 127.0.0.1 port ", " from 0.0.0.0 port ", " from  port "}};
    const auto localPrefix = std::ranges::find_if(localPrefixes, [&](std::string_view candidate) noexcept { return text.starts_with(candidate); });
    if (localPrefix == localPrefixes.end())
    {
        return std::nullopt;
    }
    text.remove_prefix(localPrefix->size());
    const auto local = std::from_chars(text.data(), text.data() + (std::min)(text.size(), size_t{6u}), ports.local);
    if (local.ec != std::errc{} || ports.local < -1 || ports.local > 65535)
    {
        return std::nullopt;
    }
    text.remove_prefix(static_cast<size_t>(local.ptr - text.data()));
    return text.starts_with(" failed: ") ? std::optional<CurlFailedConnectPorts>{ports} : std::nullopt;
}
#endif

struct CurlDirectoryCursor::Impl final
{
    Impl() = default;
    ~Impl()
    {
        ReleaseTransfer();
    }
    Impl(const Impl&)            = delete;
    Impl& operator=(const Impl&) = delete;
    Impl(Impl&&)                 = delete;
    Impl& operator=(Impl&&)      = delete;

    // Only called by the owning worker, outside libcurl callbacks. Retained
    // directory rows do not need to retain a completed transport/cache borrow.
    void ReleaseTransfer() noexcept
    {
        if (attached)
        {
            static_cast<void>(curl_multi_remove_handle(multi, cursorEasy));
            attached = false;
        }
#ifdef ENABLE_TESTS
        if (cursorEasy)
        {
            // Disable verbose first so clearing the callback cannot expose stderr.
            // Clear borrowed diagnostic state while this Impl is still alive.
            static_cast<void>(curl_easy_setopt(cursorEasy, CURLOPT_VERBOSE, 0L));
            static_cast<void>(curl_easy_setopt(cursorEasy, CURLOPT_DEBUGFUNCTION, nullptr));
            static_cast<void>(curl_easy_setopt(cursorEasy, CURLOPT_DEBUGDATA, nullptr));
        }
#endif
        cursorEasy = nullptr;
        multi      = nullptr;
        // ReturnHandle resets every borrowed callback before publishing idle
        // handles and preserves the multi's reusable connection cache.
        easy = CurlEasyPool::BorrowedHandle{};
    }

#ifdef ENABLE_TESTS
    static int CaptureFailedConnect(CURL* curl, curl_infotype type, char* data, size_t size, void* cookie) noexcept
    {
        auto* self = static_cast<Impl*>(cookie);
        if (! self || curl != self->cursorEasy || ! data || type != CURLINFO_TEXT)
        {
            return 0;
        }
        const auto ports = CurlDirectoryCursor::ParseFailedConnectTraceForSelfTest(type, std::string_view(data, size));
        if (! ports.has_value())
        {
            return 0;
        }
        const int priorError      = WSAGetLastError();
        auto restoreError         = wil::scope_exit([priorError]() noexcept { WSASetLastError(priorError); });
        self->failedConnectPorts  = ports;
        const std::wstring detail = std::format(L"phase=cursor;stage=connect;localPort={}", ports.value().local);
        Debug::Perf::Emit(L"FileOps.Curl.SocketConnectFailure",
                          detail.c_str(),
                          0u,
                          static_cast<uint64_t>(ports.value().remote),
                          static_cast<uint64_t>((std::max)(ports.value().local, 0)),
                          HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED));
        return 0;
    }
#endif

    // Read ahead only into the existing fixed buffer, including fragmented
    // small listings. A paused callback consumes none of its bytes: libcurl
    // redelivers that whole chunk on the same thread when Next resumes.
    static size_t Write(char* data, size_t size, size_t count, void* cookie) noexcept
    {
        auto& self = *static_cast<Impl*>(cookie);
        if (size != 0u && count > self.chunk.size() / size)
        {
            self.status = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
            return CURL_WRITEFUNC_ERROR;
        }
        const size_t bytes = size * count;
        if (bytes == 0u)
        {
            return 0u;
        }
        if (self.chunkOffset == self.chunkSize)
        {
            self.chunkOffset = 0u;
            self.chunkSize   = 0u;
        }
        if (bytes > self.chunk.size() - self.chunkSize)
        {
            self.paused = true;
            return CURL_WRITEFUNC_PAUSE;
        }
        std::memcpy(self.chunk.data() + self.chunkSize, data, bytes);
        self.chunkSize += bytes;
        return bytes;
    }

    [[nodiscard]] HRESULT ParseLine(FilesInformationCurl::Entry& entry, std::wstring_view timestampLeaf) noexcept
    {
        std::string_view text(line.data(), lineSize);
        lineSize = 0u;
        if (! text.empty() && text.back() == '\r')
        {
            text.remove_suffix(1u);
        }
        if (text.empty())
        {
            return S_FALSE;
        }
        if (text.starts_with("total "))
        {
            uint64_t ignored        = 0u;
            const std::string total = TrimAscii(text.substr(6u));
            const auto result       = std::from_chars(total.data(), total.data() + total.size(), ignored);
            return result.ec == std::errc{} && result.ptr == total.data() + total.size() ? S_FALSE : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (! Common::Strings::IsValidUtf8Strict(text) || text.find('\0') != std::string_view::npos ||
            (! TryParseUnixListLine(text, entry, true, timestampLeaf) && ! TryParseDosListLine(text, entry, true, timestampLeaf)))
        {
            // An unknown dialect is not a names-only deletion authority. Never
            // silently drop a malformed row and later report a complete tree.
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (entry.name == L"." || entry.name == L"..")
        {
            return S_FALSE;
        }
        return entry.name.empty() || entry.name.find_first_of(L"/\\\r\n") != std::wstring::npos ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME) : S_OK;
    }

    // Consume only the current bounded transport chunk. S_FALSE means either
    // more transport data is needed or, when complete, the listing is exhausted.
    [[nodiscard]] HRESULT ConsumeChunk(FilesInformationCurl::Entry& entry, std::wstring_view timestampLeaf) noexcept
    {
        while (chunkOffset < chunkSize)
        {
            const char* first      = chunk.data() + chunkOffset;
            const size_t available = chunkSize - chunkOffset;
            const auto* newline    = static_cast<const char*>(std::memchr(first, '\n', available));
            const size_t take      = newline ? static_cast<size_t>(newline - first) : available;
            if (take > line.size() - lineSize)
            {
                return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
            }
            if (take != 0u)
            {
                std::memcpy(line.data() + lineSize, first, take);
                lineSize += take;
            }
            chunkOffset += take + (newline ? 1u : 0u);
            if (newline)
            {
                const HRESULT parsed = ParseLine(entry, timestampLeaf);
                if (parsed != S_FALSE)
                {
                    return parsed;
                }
            }
        }
        return complete && lineSize != 0u ? ParseLine(entry, timestampLeaf) : S_FALSE;
    }

    CurlEasyPool::BorrowedHandle easy;
    CURLM* multi     = nullptr; // Borrowed from easy; the pool retains the connection cache.
    CURL* cursorEasy = nullptr; // Borrowed cursor-only handle; never used by curl_easy_perform.
#ifdef ENABLE_TESTS
    std::optional<CurlFailedConnectPorts> failedConnectPorts;
#endif
    std::function<HRESULT()> checkpoint;
    std::string url;
    std::array<char, CURL_MAX_WRITE_SIZE> chunk{};
    std::array<char, CURL_MAX_WRITE_SIZE> line{};
    size_t chunkOffset  = 0u;
    size_t chunkSize    = 0u;
    size_t lineSize     = 0u;
    uint32_t watchdogMs = 0u;
    HRESULT status      = S_OK;
    bool attached       = false;
    bool paused         = false;
    bool complete       = false;
};

CurlDirectoryCursor::CurlDirectoryCursor() noexcept = default;
CurlDirectoryCursor::~CurlDirectoryCursor()         = default;

HRESULT CurlDirectoryCursor::Open(const ConnectionInfo& conn, std::wstring_view path, std::function<HRESULT()> checkpoint) noexcept
{
    static_assert(sizeof(Impl) + CURL_MAX_WRITE_SIZE <= kMetadataReservationBytes);
    if (_impl || ! checkpoint || conn.protocol == Protocol::Imap)
    {
        return E_INVALIDARG;
    }
    HRESULT hr = checkpoint();
    if (FAILED(hr))
    {
        return hr;
    }
    hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }
    auto impl  = std::make_unique<Impl>();
    impl->easy = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! impl->easy)
    {
        return E_OUTOFMEMORY;
    }
    impl->multi      = impl->easy.GetOrCreateMulti();
    impl->cursorEasy = impl->easy.GetOrCreateCursorEasy();
    if (! impl->cursorEasy || ! impl->multi)
    {
        return E_OUTOFMEMORY;
    }
    impl->url = BuildUrl(conn, path, true, true);
    if (impl->url.empty())
    {
        return E_INVALIDARG;
    }
    impl->checkpoint = std::move(checkpoint);
    impl->watchdogMs = CurlProviderWatchdogTimeoutMs(conn.connectTimeoutMs, conn.operationTimeoutMs);
    ApplyCommonCurlOptions(impl->cursorEasy, conn, CurlCurrentOperationOptions(), false);
    // Walkers finish this listing (or fail closed) before opening a child cursor,
    // so a paused parent never sits idle across subtree work. Next still owns an
    // active-wait watchdog for the current pump; cancel/deadline stay per-checkpoint.
    curl_easy_setopt(impl->cursorEasy, CURLOPT_TIMEOUT_MS, 0L);
    curl_easy_setopt(impl->cursorEasy, CURLOPT_LOW_SPEED_TIME, 0L);
    curl_easy_setopt(impl->cursorEasy, CURLOPT_URL, impl->url.c_str());
    curl_easy_setopt(impl->cursorEasy, CURLOPT_WRITEFUNCTION, Impl::Write);
    curl_easy_setopt(impl->cursorEasy, CURLOPT_WRITEDATA, impl.get());
    curl_easy_setopt(impl->cursorEasy, CURLOPT_FAILONERROR, 1L);
#ifdef ENABLE_TESTS
    // Install the suppressing callback and owned context before enabling verbose;
    // no diagnostic text can reach stderr, and no socket behavior is replaced.
    if (curl_easy_setopt(impl->cursorEasy, CURLOPT_DEBUGFUNCTION, Impl::CaptureFailedConnect) != CURLE_OK ||
        curl_easy_setopt(impl->cursorEasy, CURLOPT_DEBUGDATA, impl.get()) != CURLE_OK || curl_easy_setopt(impl->cursorEasy, CURLOPT_VERBOSE, 1L) != CURLE_OK)
    {
        return E_FAIL;
    }
#endif
    if (curl_multi_add_handle(impl->multi, impl->cursorEasy) != CURLM_OK)
    {
        return E_FAIL;
    }
    impl->attached = true;
    _impl          = std::move(impl);
    return S_OK;
}

uint64_t CurlDirectoryCursor::RetainedPathBytes() const noexcept
{
    // URL is retained here and copied by CURLOPT_URL; decoded frame paths are
    // charged by the walker. Opaque TLS/socket state is not a metadata metric.
    return _impl ? static_cast<uint64_t>(_impl->url.capacity() + 1u) * 2u : 0u;
}

#ifdef ENABLE_TESTS
std::optional<CurlFailedConnectPorts> CurlDirectoryCursor::FailedConnectPortsForSelfTest() const noexcept
{
    return _impl ? _impl->failedConnectPorts : std::nullopt;
}

bool CurlDirectoryCursor::HasActiveTransferForSelfTest() const noexcept
{
    return _impl && static_cast<bool>(_impl->easy);
}
#endif

HRESULT CurlDirectoryCursor::Next(FilesInformationCurl::Entry& entry, std::wstring_view timestampLeaf) noexcept
{
    if (! _impl)
    {
        return E_UNEXPECTED;
    }
    Impl& state                      = *_impl;
    const ULONGLONG started          = GetTickCount64();
    const auto releaseFailedTransfer = wil::scope_exit([&]() noexcept
    {
        if (FAILED(state.status))
        {
            state.ReleaseTransfer();
        }
    });
    for (;;)
    {
        if (FAILED(state.status))
        {
            return state.status;
        }
        state.status = state.checkpoint();
        if (FAILED(state.status))
        {
            return state.status;
        }
        // A small listing must reach its final reply before yielding to child
        // traversal. A larger listing yields at the same bounded pause point.
        if (state.complete || state.paused)
        {
            const HRESULT parsed = state.ConsumeChunk(entry, timestampLeaf);
            if (parsed != S_FALSE)
            {
                state.status = parsed;
                return parsed;
            }
            if (state.complete)
            {
                return S_FALSE;
            }
        }
        if (GetTickCount64() - started >= state.watchdogMs)
        {
            state.status = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            return state.status;
        }
        if (state.paused)
        {
            state.paused           = false;
            const CURLcode unpause = curl_easy_pause(state.cursorEasy, CURLPAUSE_CONT);
            if (unpause != CURLE_OK)
            {
                state.status = HResultFromCurl(unpause);
                return state.status;
            }
            // Unpause may synchronously deliver the cached chunk.
            if (state.paused)
            {
                continue;
            }
        }
        int running = 0;
        if (curl_multi_perform(state.multi, &running) != CURLM_OK)
        {
            state.status = E_FAIL;
            return state.status;
        }
        int remaining = 0;
        while (CURLMsg* message = curl_multi_info_read(state.multi, &remaining))
        {
            if (message->msg == CURLMSG_DONE)
            {
                RecordCurlConnectionFailure(state.cursorEasy, message->data.result, L"cursor");
                state.complete = true;
                if (SUCCEEDED(state.status))
                {
                    state.status = HResultFromCurl(message->data.result);
                }
            }
        }
        if (state.complete)
        {
            // Capture result/OS diagnostics and drain messages before the
            // pooled handles are reset. Buffered records stay owned by Impl.
            state.ReleaseTransfer();
        }
        if (FAILED(state.status) || state.complete || state.paused)
        {
            continue;
        }
        if (running == 0 || curl_multi_poll(state.multi, nullptr, 0u, 50, nullptr) != CURLM_OK)
        {
            state.status = E_FAIL;
            return state.status;
        }
    }
}

#ifdef ENABLE_TESTS
HRESULT CurlDirectoryCursor::ParseChunksForSelfTest(std::span<const std::string_view> chunks,
                                                    std::vector<FilesInformationCurl::Entry>& entries,
                                                    std::wstring_view lookupLeaf,
                                                    uint64_t* inspectedRows) noexcept
{
    // Deterministic callback/framing seam: no socket, pool borrow, or alternate
    // parser. Production Next owns transport completion and cancellation proof.
    auto state = std::make_unique<Impl>();
    entries.clear();
    if (inspectedRows)
    {
        *inspectedRows = 0u;
    }
    FilesInformationCurl::Entry entry{};
    for (size_t index = 0u;;)
    {
        state->complete = index == chunks.size();
        if (! state->complete)
        {
            std::string delivered(chunks[index]); // The libcurl callback accepts writable storage.
            const size_t accepted = Impl::Write(delivered.data(), 1u, delivered.size(), state.get());
            if (accepted == delivered.size())
            {
                ++index;
            }
            else if (accepted != CURL_WRITEFUNC_PAUSE)
            {
                return FAILED(state->status) ? state->status : E_UNEXPECTED;
            }
        }
        if (! state->complete && ! state->paused)
        {
            continue;
        }
        for (;;)
        {
            const HRESULT hr = state->ConsumeChunk(entry, lookupLeaf);
            if (FAILED(hr))
            {
                return hr;
            }
            if (hr == S_FALSE)
            {
                break;
            }
            if (inspectedRows)
            {
                ++*inspectedRows;
            }
            if (lookupLeaf.empty() || entry.name == lookupLeaf)
            {
                entries.push_back(std::move(entry));
            }
        }
        if (state->complete)
        {
            return S_OK;
        }
        state->paused = false; // Redeliver the unconsumed callback on resume.
    }
}
#endif

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
    return Common::HandleIo::Rewind(file);
}

[[nodiscard]] HRESULT ResetFileForRewrite(HANDLE file) noexcept
{
    HRESULT hr = ResetFilePointerToStart(file);
    if (FAILED(hr))
    {
        return hr;
    }

    if (SetEndOfFile(file) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& out) noexcept
{
    return Common::HandleIo::GetFileSizeBounded(file, (std::numeric_limits<uint64_t>::max)(), out);
}

} // namespace FileSystemCurlInternal

namespace FileSystemCurlInternal
{
[[nodiscard]] uint64_t ClampCurlOffToUInt64(curl_off_t value) noexcept
{
    if (value <= 0)
    {
        return 0;
    }

    constexpr curl_off_t max = (std::numeric_limits<curl_off_t>::max)();
    if (value > max)
    {
        return (std::numeric_limits<uint64_t>::max)();
    }

    return static_cast<uint64_t>(value);
}

[[nodiscard]] uint64_t SaturatingAddToAtomic(std::atomic<uint64_t>& value, uint64_t delta) noexcept
{
    if (delta == 0)
    {
        return value.load(std::memory_order_acquire);
    }

    uint64_t current = value.load(std::memory_order_relaxed);
    for (;;)
    {
        uint64_t next = current;
        if (current > (std::numeric_limits<uint64_t>::max)() - delta)
        {
            next = (std::numeric_limits<uint64_t>::max)();
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

    const uint64_t nowTick = GetTickCount64();

    const uint64_t phaseTotal = ClampCurlOffToUInt64(ctx->isUpload ? ultotal : dltotal);
    const uint64_t phaseNow   = ClampCurlOffToUInt64(ctx->isUpload ? ulnow : dlnow);

    if (ctx->itemTotalBytes == 0 && phaseTotal > 0)
    {
        ctx->itemTotalBytes = phaseTotal;
    }

    uint64_t itemDone  = phaseNow;
    uint64_t itemTotal = phaseTotal;

    if (ctx->scaleForCopy && ctx->itemTotalBytes > 0)
    {
        itemTotal           = ctx->itemTotalBytes;
        const uint64_t half = itemTotal / 2u;
        if (! ctx->scaleForCopySecond)
        {
            itemDone = std::min(half, phaseNow / 2u);
        }
        else
        {
            const uint64_t extra = (itemTotal & 1u) != 0 ? 1u : 0u;
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

    uint64_t wireDone = phaseNow;
    if (ctx->scaleForCopy && ctx->itemTotalBytes > 0)
    {
        const uint64_t offset = ctx->scaleForCopySecond ? ctx->itemTotalBytes : 0;
        wireDone              = offset > (std::numeric_limits<uint64_t>::max)() - phaseNow ? (std::numeric_limits<uint64_t>::max)() : (offset + phaseNow);
    }

    uint64_t overall = 0;
    if (ctx->concurrentOverallBytes)
    {
        const uint64_t delta        = wireDone >= ctx->lastConcurrentWireDone ? (wireDone - ctx->lastConcurrentWireDone) : 0;
        ctx->lastConcurrentWireDone = wireDone;
        overall                     = SaturatingAddToAtomic(*ctx->concurrentOverallBytes, delta);
    }
    else
    {
        overall = ctx->baseCompletedBytes > (std::numeric_limits<uint64_t>::max)() - wireDone ? (std::numeric_limits<uint64_t>::max)()
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
    const uint64_t limit = ctx->progress->bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
    if (limit > 0 && ctx->throttleStartTick != 0)
    {
        const uint64_t elapsedMs = nowTick - ctx->throttleStartTick;
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
    // URLs and quote-command operands have different grammars. Reuse the
    // captured literal base path, never the URL-escaped representation.
    const std::wstring path = JoinPluginPathWide(conn.basePathWide, pluginPath);
    if (path.find_first_of(L"\r\n") != std::wstring::npos || path.find(L'\0') != std::wstring::npos)
    {
        return {}; // No control-line injection or C-string truncation.
    }
    std::string remote = Common::Strings::Utf8FromUtf16StrictOrEmpty(path);
    if (remote.empty())
    {
        return {};
    }
    while (remote.size() > 1u && remote.back() == '/')
    {
        remote.pop_back();
    }
    if (conn.protocol == Protocol::Ftp)
    {
        // FTP has one pathname occupying the rest of the command line; quotes
        // are literal filename bytes, not a libcurl argument parser convention.
        return remote;
    }
    // SFTP (also used for SCP namespace commands) tokenizes its operands.
    // This is libcurl quote grammar, not shell or Windows argv escaping.
    std::string quoted;
    quoted.reserve(remote.size() + 2u);
    quoted.push_back('"');
    for (const char ch : remote)
    {
        if (ch == '\\' || ch == '"')
        {
            quoted.push_back('\\');
        }
        quoted.push_back(ch);
    }
    quoted.push_back('"');
    return quoted;
}

namespace
{
[[nodiscard]] bool IsCurlTransientTransferError(CURLcode code) noexcept
{
    return code == CURLE_OPERATION_TIMEDOUT || code == CURLE_RECV_ERROR || code == CURLE_SEND_ERROR || code == CURLE_GOT_NOTHING ||
           code == CURLE_COULDNT_CONNECT || code == CURLE_PARTIAL_FILE;
}

[[nodiscard]] unsigned long CurlRetryDelayMs(unsigned int retryIndex) noexcept
{
    // retryIndex is 1 for the first retry.
    constexpr unsigned long kBaseDelayMs = 200;
    constexpr unsigned long kMaxDelayMs  = 2000;

    unsigned long delay = kBaseDelayMs;
    for (unsigned int i = 1; i < retryIndex; ++i)
    {
        if (delay >= kMaxDelayMs / 4u)
        {
            delay = kMaxDelayMs;
            break;
        }
        delay *= 4u;
    }

    return delay;
}

struct FtpControlReplyCapture final
{
    long lastOtherFailureReplyCode = 0;
    bool sawAuthenticationFailure  = false;
    bool sawPrimitiveUnavailable   = false;
};

// FTP NOBODY synthesizes metadata into the body callback even without HEADER.
// Unlike payload/string sinks, a size probe retains no response bytes and must
// never depend on the GUI process's CRT stdout. Keep this transport-only sink here.
constexpr size_t ConsumeCurlProbeMetadata(char* /*data*/, size_t size, size_t count, void* /*context*/) noexcept
{
    if (size != 0u && count > (std::numeric_limits<size_t>::max)() / size)
    {
        return CURL_WRITEFUNC_ERROR;
    }
    return size * count;
}

static_assert(ConsumeCurlProbeMetadata(nullptr, 1u, CURL_MAX_HTTP_HEADER, nullptr) == CURL_MAX_HTTP_HEADER);
static_assert(ConsumeCurlProbeMetadata(nullptr, 0u, (std::numeric_limits<size_t>::max)(), nullptr) == 0u);
static_assert(ConsumeCurlProbeMetadata(nullptr, 2u, (std::numeric_limits<size_t>::max)(), nullptr) == CURL_WRITEFUNC_ERROR);

int CaptureFtpControlReply(CURL* /*curl*/, curl_infotype type, char* data, size_t size, void* context) noexcept
{
    auto* capture = static_cast<FtpControlReplyCapture*>(context);
    if (! capture || (type != CURLINFO_HEADER_IN && type != CURLINFO_TEXT) || ! data)
    {
        return 0;
    }

    const size_t replyOffset = (size >= 5u && data[0] == '<' && data[1] == ' ') ? 2u : 0u;
    if (size < replyOffset + 3u || data[replyOffset] < '0' || data[replyOffset] > '9' || data[replyOffset + 1u] < '0' || data[replyOffset + 1u] > '9' ||
        data[replyOffset + 2u] < '0' || data[replyOffset + 2u] > '9')
    {
        return 0;
    }

    const long replyCode = static_cast<long>((data[replyOffset] - '0') * 100 + (data[replyOffset + 1u] - '0') * 10 + (data[replyOffset + 2u] - '0'));
    if (replyCode == 530)
    {
        capture->sawAuthenticationFailure = true;
    }
    else if (replyCode == 500 || replyCode == 502 || replyCode == 504)
    {
        capture->sawPrimitiveUnavailable = true;
    }
    else if (replyCode >= 400)
    {
        capture->lastOtherFailureReplyCode = replyCode;
    }
    return 0;
}

} // namespace

[[nodiscard]] HRESULT CurlProbeRemoteFileSize(const ConnectionInfo& conn, std::wstring_view pluginPath, uint64_t& sizeOut, bool& sizeKnownOut) noexcept
{
    sizeOut      = 0;
    sizeKnownOut = false;

    // Targeted existence/size probe (FTP SIZE, SFTP stat) so callers do not have to LIST the whole parent
    // directory just to stat one file. IMAP has no equivalent, so report unsupported and let the caller
    // fall back to the listing-based stat. On success the file exists; sizeKnownOut indicates whether the
    // server actually reported a byte count (some FTP dialects answer existence but not SIZE).
    if (conn.protocol == Protocol::Imap)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, false, true);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    constexpr unsigned int kMaxAttempts = 3u;
    char errorBuffer[CURL_ERROR_SIZE]{};
    CURLcode code = CURLE_FAILED_INIT;
    FtpControlReplyCapture ftpReply{};

    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (attempt > 0)
        {
            curl_easy_reset(curl.get());
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_NOBODY, 1L); // FTP: issues SIZE (no RETR); SFTP: stat. No data body transferred.
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, ConsumeCurlProbeMetadata);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, nullptr);
        curl_easy_setopt(curl.get(), CURLOPT_HEADERFUNCTION, ConsumeCurlProbeMetadata);
        curl_easy_setopt(curl.get(), CURLOPT_HEADERDATA, nullptr);

        errorBuffer[0] = '\0';
        curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

        ApplyCommonCurlOptions(curl.get(), conn, CurlCurrentOperationOptions(), false);
        if (conn.protocol == Protocol::Ftp)
        {
            ftpReply = {};
            // CURLOPT_VERBOSE routes the current transfer's FTP control replies
            // through the debug callback. The callback records numeric status
            // only; it never retains or logs commands, payloads, or credentials.
            curl_easy_setopt(curl.get(), CURLOPT_VERBOSE, 1L);
            curl_easy_setopt(curl.get(), CURLOPT_DEBUGFUNCTION, CaptureFtpControlReply);
            curl_easy_setopt(curl.get(), CURLOPT_DEBUGDATA, &ftpReply);
        }

        code = curl_easy_perform(curl.get());
        RecordCurlConnectionFailure(curl.get(), code, L"size-probe");
        if (code == CURLE_OK)
        {
            break;
        }

        if (attempt + 1u >= kMaxAttempts || ! IsCurlTransientTransferError(code))
        {
            break;
        }

        Sleep(CurlRetryDelayMs(attempt + 1u));
    }

    curl_off_t contentLength = -1;
    const bool contentLengthKnown =
        code == CURLE_OK && curl_easy_getinfo(curl.get(), CURLINFO_CONTENT_LENGTH_DOWNLOAD_T, &contentLength) == CURLE_OK && contentLength >= 0;
    if (contentLengthKnown)
    {
        sizeOut      = ClampCurlOffToUInt64(contentLength);
        sizeKnownOut = true;
        return S_OK;
    }

    long responseCode = ftpReply.lastOtherFailureReplyCode;
    if (conn.protocol != Protocol::Ftp)
    {
        static_cast<void>(curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &responseCode));
    }
    if (conn.protocol == Protocol::Ftp && ftpReply.sawAuthenticationFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
    }
    if (conn.protocol == Protocol::Ftp && responseCode == 550)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (conn.protocol == Protocol::Ftp && responseCode >= 400)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
    }
    if (conn.protocol == Protocol::Ftp && ftpReply.sawPrimitiveUnavailable)
    {
        // The endpoint explicitly does not implement the targeted SIZE
        // command. libcurl can report CURLE_OK for an FTP NOBODY request even
        // when SIZE itself returned one of these replies, so classify the
        // protocol response independently of the transfer return code.
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (code != CURLE_OK)
    {
        return HResultFromCurl(code);
    }

    return S_OK;
}

[[nodiscard]] HRESULT ResolveCurlSourceSizeCommitment(
    const ConnectionInfo& conn, std::wstring_view pluginPath, uint64_t listedSizeBytes, bool listedSizeKnown, CurlSourceSizeCommitment& commitmentOut) noexcept
{
    commitmentOut = CurlSourceSizeCommitment{.sizeBytes = listedSizeBytes, .known = listedSizeKnown};

    uint64_t probedSizeBytes = 0u;
    bool probedSizeKnown     = false;
    const HRESULT probeHr    = CurlProbeRemoteFileSize(conn, pluginPath, probedSizeBytes, probedSizeKnown);
    if (SUCCEEDED(probeHr))
    {
        if (probedSizeKnown)
        {
            commitmentOut.sizeBytes = probedSizeBytes;
            commitmentOut.known     = true;
        }
        return S_OK;
    }

    if (probeHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
    {
        // Preserve the listing commitment when present, or return an explicit
        // unknown-size result for the caller's Copy-versus-Move policy.
        return S_OK;
    }

    // Cancellation, authentication, disappearance, and transport failures are
    // not evidence that the targeted primitive is merely unavailable. Do not
    // hide them behind a possibly stale directory-listing value.
    return NormalizeCancellation(probeHr);
}

[[nodiscard]] HRESULT CurlPerformList(const ConnectionInfo& conn, std::wstring_view pluginPath, std::string& outListing) noexcept
{
    outListing.clear();

    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, true, true);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    constexpr unsigned int kMaxAttempts = 3u;
    char errorBuffer[CURL_ERROR_SIZE]{};
    CURLcode code = CURLE_FAILED_INIT;

    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (attempt > 0)
        {
            outListing.clear();
            curl_easy_reset(curl.get());
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &outListing);
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

        errorBuffer[0] = '\0';
        curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

        ApplyCommonCurlOptions(curl.get(), conn, CurlCurrentOperationOptions(), false);

        code = curl_easy_perform(curl.get());
        if (code == CURLE_OK)
        {
            return S_OK;
        }

        if (attempt + 1u >= kMaxAttempts || ! IsCurlTransientTransferError(code))
        {
            break;
        }

        Sleep(CurlRetryDelayMs(attempt + 1u));
    }

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

    return HResultFromCurl(code);
}

[[nodiscard]] HRESULT CurlPerformListAndParse(const ConnectionInfo& conn,
                                              std::wstring_view pluginPath,
                                              std::vector<FilesInformationCurl::Entry>& outEntries) noexcept
{
    outEntries.clear();

    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    const std::string url = BuildUrl(conn, pluginPath, true, true);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    CurlListParseContext parse(outEntries);

    constexpr unsigned int kMaxAttempts = 3u;
    char errorBuffer[CURL_ERROR_SIZE]{};
    CURLcode code = CURLE_FAILED_INIT;

    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (attempt > 0)
        {
            parse.Reset();
            curl_easy_reset(curl.get());
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToListParser);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &parse);
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

        errorBuffer[0] = '\0';
        curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

        ApplyCommonCurlOptions(curl.get(), conn, CurlCurrentOperationOptions(), false);

        code = curl_easy_perform(curl.get());
        if (code == CURLE_OK)
        {
            parse.FinalizePendingLine();

            if (parse.HasFallbackNames())
            {
                const auto& names = parse.GetFallbackNamesUtf8();
                outEntries.reserve(names.size());
                for (const std::string& nameUtf8 : names)
                {
                    FilesInformationCurl::Entry entry{};
                    entry.name       = Utf16FromUtf8(nameUtf8);
                    entry.attributes = FILE_ATTRIBUTE_NORMAL;
                    if (! entry.name.empty())
                    {
                        outEntries.push_back(std::move(entry));
                    }
                }
            }

            return S_OK;
        }

        if (attempt + 1u >= kMaxAttempts || ! IsCurlTransientTransferError(code))
        {
            break;
        }

        Sleep(CurlRetryDelayMs(attempt + 1u));
    }

    outEntries.clear();

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

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
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

    char errorBuffer[CURL_ERROR_SIZE]{};
    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &sink);
    curl_easy_setopt(curl.get(), CURLOPT_NOBODY, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_QUOTE, list.get());
    curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);
    ApplyCommonCurlOptions(curl.get(), conn, CurlCurrentOperationOptions(), false);

    // Every caller mutates the namespace. A lost reply may follow a committed command;
    // replay could remove a replacement or rename it over the first result. Execute
    // once, with no unrelated directory transfer after the command acknowledgement.
    const CURLcode code = curl_easy_perform(curl.get());
    RecordCurlConnectionFailure(curl.get(), code, L"quote");
    if (code == CURLE_OK)
    {
        return S_OK;
    }

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

    return HResultFromCurl(code);
}

namespace
{
[[nodiscard]] HRESULT SleepWithCancelCheck(unsigned long totalMs, TransferProgressContext* progressCtx) noexcept
{
    if (totalMs == 0)
    {
        return S_OK;
    }

    constexpr unsigned long kSliceMs = 50;

    unsigned long remaining = totalMs;
    while (remaining > 0)
    {
        const unsigned long slice = remaining < kSliceMs ? remaining : kSliceMs;
        Sleep(slice);
        remaining -= slice;

        if (progressCtx && progressCtx->progress)
        {
            const HRESULT cancelHr = progressCtx->progress->CheckCancel();
            if (FAILED(cancelHr))
            {
                return cancelHr;
            }
        }
    }

    return S_OK;
}

void PrepareProgressContextForRetry(TransferProgressContext* progressCtx) noexcept
{
    if (! progressCtx)
    {
        return;
    }

    progressCtx->throttleStartTick = GetTickCount64();
    progressCtx->lastThrottleBytes = 0;
    progressCtx->lastCancelTick    = 0;
    progressCtx->lastReportTick    = 0;
    progressCtx->abortHr           = S_OK;
}
} // namespace

[[nodiscard]] HRESULT CurlDownloadToFile(const ConnectionInfo& conn,
                                         std::wstring_view pluginPath,
                                         HANDLE file,
                                         const FileSystemOptions* options,
                                         TransferProgressContext* progressCtx,
                                         std::optional<uint64_t> expectedSizeBytes) noexcept
{
    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    const std::string url = BuildUrl(conn, pluginPath, false, false);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    constexpr unsigned int kMaxAttempts = 2u; // at least one retry for transient network errors
    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (attempt > 0)
        {
            hr = ResetFileForRewrite(file);
            if (FAILED(hr))
            {
                return hr;
            }

            curl_easy_reset(curl.get());
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToFile);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, file);
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

        if (progressCtx)
        {
            if (attempt == 0)
            {
                progressCtx->Begin();
            }
            else
            {
                PrepareProgressContextForRetry(progressCtx);
            }

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

        if (expectedSizeBytes.has_value())
        {
            // Do not let libcurl stop at the endpoint's advertised length. The
            // independently probed source size is this operation's commitment,
            // so the write path must observe and reject both short and overlong
            // bodies itself.
            curl_easy_setopt(curl.get(), CURLOPT_IGNORE_CONTENT_LENGTH, 1L);
        }

        const CURLcode code = curl_easy_perform(curl.get());
        if (code == CURLE_OK)
        {
            if (expectedSizeBytes.has_value())
            {
                uint64_t downloadedSizeBytes = 0u;
                const HRESULT sizeHr         = GetFileSizeBytes(file, downloadedSizeBytes);
                if (FAILED(sizeHr))
                {
                    return sizeHr;
                }
                if (downloadedSizeBytes != expectedSizeBytes.value())
                {
                    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                }
            }
            return S_OK;
        }

        if (code == CURLE_PARTIAL_FILE && expectedSizeBytes.has_value())
        {
            uint64_t downloadedSizeBytes = 0u;
            const HRESULT sizeHr         = GetFileSizeBytes(file, downloadedSizeBytes);
            if (FAILED(sizeHr))
            {
                return sizeHr;
            }
            if (downloadedSizeBytes != expectedSizeBytes.value())
            {
                // A clean protocol close with fewer bytes than the authoritative
                // source commitment is an integrity failure, not a transient
                // connection failure eligible for an automatic retry.
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
        }

        if (code == CURLE_ABORTED_BY_CALLBACK && progressCtx && FAILED(progressCtx->abortHr))
        {
            return progressCtx->abortHr;
        }

        if (attempt + 1u >= kMaxAttempts || ! IsCurlTransientTransferError(code))
        {
            return HResultFromCurl(code);
        }

        if (progressCtx && progressCtx->progress)
        {
            const HRESULT cancelHr = progressCtx->progress->CheckCancel();
            if (FAILED(cancelHr))
            {
                return cancelHr;
            }
        }

        hr = SleepWithCancelCheck(CurlRetryDelayMs(attempt + 1u), progressCtx);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return E_FAIL;
}

[[nodiscard]] HRESULT CurlUploadFromFile(const ConnectionInfo& conn,
                                         std::wstring_view pluginPath,
                                         HANDLE file,
                                         uint64_t sizeBytes,
                                         const FileSystemOptions* options,
                                         TransferProgressContext* progressCtx) noexcept
{
    HRESULT hr = EnsureCurlInitialized();
    if (FAILED(hr))
    {
        return hr;
    }

    const std::string url = BuildUrl(conn, pluginPath, false, false);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    auto curl = GetCurlEasyPool().Borrow(conn.limiterKey);
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    constexpr unsigned int kMaxAttempts = 2u; // at least one retry for transient network errors
    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (attempt > 0)
        {
            hr = ResetFilePointerToStart(file);
            if (FAILED(hr))
            {
                return hr;
            }

            curl_easy_reset(curl.get());
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_UPLOAD, 1L);
        curl_easy_setopt(curl.get(), CURLOPT_READFUNCTION, CurlReadFromFile);
        curl_easy_setopt(curl.get(), CURLOPT_READDATA, file);
        curl_easy_setopt(curl.get(),
                         CURLOPT_INFILESIZE_LARGE,
                         static_cast<curl_off_t>(std::min<uint64_t>(sizeBytes, static_cast<uint64_t>((std::numeric_limits<curl_off_t>::max)()))));
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

        if (progressCtx)
        {
            if (attempt == 0)
            {
                progressCtx->Begin();
            }
            else
            {
                PrepareProgressContextForRetry(progressCtx);
            }

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
        if (code == CURLE_OK)
        {
            return S_OK;
        }

        if (code == CURLE_ABORTED_BY_CALLBACK && progressCtx && FAILED(progressCtx->abortHr))
        {
            return progressCtx->abortHr;
        }

        if (attempt + 1u >= kMaxAttempts || ! IsCurlTransientTransferError(code))
        {
            return HResultFromCurl(code);
        }

        if (progressCtx && progressCtx->progress)
        {
            const HRESULT cancelHr = progressCtx->progress->CheckCancel();
            if (FAILED(cancelHr))
            {
                return cancelHr;
            }
        }

        hr = SleepWithCancelCheck(CurlRetryDelayMs(attempt + 1u), progressCtx);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return E_FAIL;
}
} // namespace FileSystemCurlInternal

// FileSystemCurl

FileSystemCurl::FileSystemCurl(FileSystemCurlProtocol protocol, IHost* host) : _protocol(protocol)
{
    g_fileSystemCurlInstanceCount.fetch_add(1, std::memory_order_relaxed);

    switch (_protocol)
    {
        case FileSystemCurlProtocol::Ftp:
            _metaData.id          = kPluginIdFtp;
            _metaData.shortId     = kPluginShortIdFtp;
            _metaData.name        = LocalizedPluginName(FileSystemCurlProtocol::Ftp);
            _metaData.description = LocalizedPluginDescription(FileSystemCurlProtocol::Ftp);
            break;
        case FileSystemCurlProtocol::Sftp:
            _metaData.id          = kPluginIdSftp;
            _metaData.shortId     = kPluginShortIdSftp;
            _metaData.name        = LocalizedPluginName(FileSystemCurlProtocol::Sftp);
            _metaData.description = LocalizedPluginDescription(FileSystemCurlProtocol::Sftp);
            break;
        case FileSystemCurlProtocol::Scp:
            _metaData.id          = kPluginIdScp;
            _metaData.shortId     = kPluginShortIdScp;
            _metaData.name        = LocalizedPluginName(FileSystemCurlProtocol::Scp);
            _metaData.description = LocalizedPluginDescription(FileSystemCurlProtocol::Scp);
            break;
        case FileSystemCurlProtocol::Imap:
            _metaData.id          = kPluginIdImap;
            _metaData.shortId     = kPluginShortIdImap;
            _metaData.name        = LocalizedPluginName(FileSystemCurlProtocol::Imap);
            _metaData.description = LocalizedPluginDescription(FileSystemCurlProtocol::Imap);
            break;
    }
    _metaData.author  = kPluginAuthor;
    _metaData.version = kPluginVersion;

    _configurationJson = "{}";

    _driveFileSystem = _metaData.shortId ? _metaData.shortId : L"";

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostAlerts), _hostAlerts.put_void()));
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

void FileSystemCurl::ObserveCurlCleanupDebt(const CurlPublicationResult& result) const noexcept
{
    if (result.cleanupDebtCount == 0u)
    {
        return;
    }

    Debug::Warning(L"FileSystemCurl: operation retained {} remote recovery item(s) (kindMask={:#x}, cleanupHr={:#x}).",
                   result.cleanupDebtCount,
                   result.cleanupDebtMask,
                   static_cast<unsigned long>(result.cleanupHr));
    Debug::Perf::Emit(
        L"FileOps.Curl.CleanupDebt", L"aggregate retained remote recovery items", 0u, result.cleanupDebtCount, result.cleanupDebtMask, result.cleanupHr);

    if (! _hostAlerts)
    {
        return;
    }

    const std::wstring title   = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_CLEANUP_DEBT_TITLE);
    const std::wstring message = FormatStringResource(g_hInstance, IDS_FILESYSTEMCURL_CLEANUP_DEBT_MESSAGE, result.cleanupDebtCount);
    if (message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = HOST_ALERT_WARNING;
    request.targetWindow = nullptr;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;
    static_cast<void>(_hostAlerts->ShowAlert(&request, nullptr));
}

FileSystemCurl::~FileSystemCurl()
{
    SecureWipe::SecureClear(_settings.defaultPassword);
    SecureWipe::SecureClear(_settings.sshKeyPassphrase);
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

    if (riid == __uuidof(IFileSystemPathCapabilities2))
    {
        *ppvObject = static_cast<IFileSystemPathCapabilities2*>(static_cast<IFileSystem*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemRouteCapabilities))
    {
        *ppvObject = static_cast<IFileSystemRouteCapabilities*>(static_cast<FileSystemRouteCapabilitiesBase*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemAtomicWriter))
    {
        *ppvObject = static_cast<IFileSystemAtomicWriter*>(this);
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

ULONG STDMETHODCALLTYPE FileSystemCurl::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemCurl::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        const unsigned long remainingInstances = g_fileSystemCurlInstanceCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (remainingInstances == 0)
        {
            ShutdownSharedCopyMoveJobScheduler();
            TryCompleteFileSystemCurlShutdown();
        }

        delete this;
    }
    return result;
}

std::wstring FileSystemCurl::MakeWatchPathKey(std::wstring_view path) noexcept
{
    std::wstring key(path);
    for (wchar_t& ch : key)
    {
        ch = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
    }
    return key;
}

namespace
{
[[nodiscard]] std::wstring NormalizeSyntheticParentPath(std::wstring_view path) noexcept
{
    const std::filesystem::path parent = std::filesystem::path(path).parent_path();
    if (parent.empty())
    {
        return L"/";
    }

    return FileSystemCurlInternal::NormalizePluginPath(parent.native());
}
} // namespace

bool FileSystemCurl::IsSameOrDescendantPath(std::wstring_view candidateKey, std::wstring_view rootKey) noexcept
{
    if (candidateKey == rootKey)
    {
        return true;
    }

    return candidateKey.size() > rootKey.size() && candidateKey.rfind(rootKey, 0) == 0 && candidateKey[rootKey.size()] == L'/';
}

std::wstring FileSystemCurl::RelativeWatchPath(std::wstring_view watchedPath, std::wstring_view fullPath) noexcept
{
    if (fullPath == watchedPath)
    {
        return {};
    }

    size_t offset = watchedPath.size();
    if (offset < fullPath.size() && fullPath[offset] == L'/')
    {
        ++offset;
    }

    return std::wstring(fullPath.substr(offset));
}

void FileSystemCurl::EmitSyntheticWatchNotification(std::wstring_view watchedPath, const std::vector<SyntheticWatchChange>& changes, bool overflow) noexcept
{
    std::shared_ptr<SyntheticWatchRegistration> registration;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring watchedKey = MakeWatchPathKey(watchedPath);
        for (const auto& candidate : _syntheticWatches)
        {
            if (! candidate || ! candidate->active.load(std::memory_order_acquire))
            {
                continue;
            }

            if (candidate->watchedPathKey == watchedKey)
            {
                registration = candidate;
                break;
            }
        }
    }

    if (! registration || changes.empty())
    {
        return;
    }

    std::vector<FileSystemDirectoryChange> rawChanges;
    rawChanges.reserve(changes.size());
    for (const SyntheticWatchChange& change : changes)
    {
        rawChanges.push_back(FileSystemDirectoryChange{
            .action           = change.action,
            .relativePath     = change.relativePath.empty() ? nullptr : change.relativePath.c_str(),
            .relativePathSize = static_cast<unsigned long>(change.relativePath.size() * sizeof(wchar_t)),
        });
    }

    registration->inFlight.fetch_add(1u, std::memory_order_acq_rel);
    if (! registration->active.load(std::memory_order_acquire))
    {
        if (registration->inFlight.fetch_sub(1u, std::memory_order_acq_rel) == 1u)
        {
            std::lock_guard lock(registration->drainMutex);
            registration->drainCv.notify_all();
        }
        return;
    }

    FileSystemDirectoryChangeNotification notification{};
    notification.sizeBytes       = sizeof(notification);
    notification.watchedPath     = registration->watchedPath.c_str();
    notification.watchedPathSize = static_cast<unsigned long>(registration->watchedPath.size() * sizeof(wchar_t));
    notification.changes         = rawChanges.data();
    notification.changeCount     = static_cast<unsigned long>(rawChanges.size());
    notification.overflow        = overflow ? TRUE : FALSE;

    registration->callbackThreadId.store(GetCurrentThreadId(), std::memory_order_release);
    if (registration->callback)
    {
        static_cast<void>(registration->callback->FileSystemDirectoryChanged(&notification, registration->cookie));
    }
    registration->callbackThreadId.store(0u, std::memory_order_release);

    if (registration->inFlight.fetch_sub(1u, std::memory_order_acq_rel) == 1u)
    {
        std::lock_guard lock(registration->drainMutex);
        registration->drainCv.notify_all();
    }
}

void FileSystemCurl::NotifySyntheticPathCreated(std::wstring_view fullPath) noexcept
{
    const std::wstring normalized = NormalizePluginPath(fullPath);
    const std::wstring parent     = NormalizeSyntheticParentPath(normalized);

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring parentKey = MakeWatchPathKey(parent);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != parentKey)
            {
                continue;
            }

            std::wstring relative = RelativeWatchPath(registration->watchedPath, normalized);
            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action       = FILESYSTEM_DIR_CHANGE_ADDED;
            changes[0].relativePath = std::move(relative);
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

void FileSystemCurl::NotifySyntheticPathDeleted(std::wstring_view fullPath) noexcept
{
    const std::wstring normalized = NormalizePluginPath(fullPath);
    const std::wstring parent     = NormalizeSyntheticParentPath(normalized);

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring parentKey = MakeWatchPathKey(parent);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != parentKey)
            {
                continue;
            }

            std::wstring relative = RelativeWatchPath(registration->watchedPath, normalized);
            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action       = FILESYSTEM_DIR_CHANGE_REMOVED;
            changes[0].relativePath = std::move(relative);
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

void FileSystemCurl::NotifySyntheticPathMoved(std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept
{
    const std::wstring normalizedSource      = NormalizePluginPath(sourcePath);
    const std::wstring normalizedDestination = NormalizePluginPath(destinationPath);
    const std::wstring sourceParent          = NormalizeSyntheticParentPath(normalizedSource);
    const std::wstring destinationParent     = NormalizeSyntheticParentPath(normalizedDestination);

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire))
            {
                continue;
            }

            std::vector<SyntheticWatchChange> changes;
            std::wstring oldRelative;
            std::wstring newRelative;

            if (registration->watchedPathKey == MakeWatchPathKey(sourceParent) && registration->watchedPathKey == MakeWatchPathKey(destinationParent))
            {
                oldRelative = RelativeWatchPath(registration->watchedPath, normalizedSource);
                newRelative = RelativeWatchPath(registration->watchedPath, normalizedDestination);
                changes.resize(2);
                changes[0].action       = FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME;
                changes[0].relativePath = std::move(oldRelative);
                changes[1].action       = FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME;
                changes[1].relativePath = std::move(newRelative);
            }
            else
            {
                if (registration->watchedPathKey == MakeWatchPathKey(sourceParent))
                {
                    oldRelative = RelativeWatchPath(registration->watchedPath, normalizedSource);
                    changes.push_back(SyntheticWatchChange{.action = FILESYSTEM_DIR_CHANGE_REMOVED, .relativePath = std::move(oldRelative)});
                }
                if (registration->watchedPathKey == MakeWatchPathKey(destinationParent))
                {
                    newRelative = RelativeWatchPath(registration->watchedPath, normalizedDestination);
                    changes.push_back(SyntheticWatchChange{.action = FILESYSTEM_DIR_CHANGE_ADDED, .relativePath = std::move(newRelative)});
                }
            }

            if (! changes.empty())
            {
                notifications.emplace_back(registration->watchedPath, std::move(changes));
            }
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

void FileSystemCurl::NotifySyntheticFolderChanged(std::wstring_view folderPath) noexcept
{
    const std::wstring normalized = NormalizePluginPath(folderPath);
    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring folderKey = MakeWatchPathKey(normalized);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != folderKey)
            {
                continue;
            }

            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action = FILESYSTEM_DIR_CHANGE_MODIFIED;
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept
{
    if (! path || ! callback)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring normalized = NormalizePluginPath(path);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    auto registration            = std::make_shared<SyntheticWatchRegistration>();
    registration->watchedPath    = normalized;
    registration->watchedPathKey = MakeWatchPathKey(normalized);
    registration->callback       = callback;
    registration->cookie         = cookie;

    std::lock_guard lock(_watchMutex);
    for (const auto& existing : _syntheticWatches)
    {
        if (existing && existing->active.load(std::memory_order_acquire) && existing->watchedPathKey == registration->watchedPathKey)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    _syntheticWatches.push_back(std::move(registration));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::UnwatchDirectory(const wchar_t* path) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring normalized = NormalizePluginPath(path);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    std::shared_ptr<SyntheticWatchRegistration> registration;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring watchKey = MakeWatchPathKey(normalized);
        auto it = std::find_if(_syntheticWatches.begin(), _syntheticWatches.end(), [&](const std::shared_ptr<SyntheticWatchRegistration>& candidate) noexcept {
            return candidate && candidate->active.load(std::memory_order_acquire) && candidate->watchedPathKey == watchKey;
        });
        if (it == _syntheticWatches.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        registration = *it;
        _syntheticWatches.erase(it);
        registration->active.store(false, std::memory_order_release);
    }

    if (! registration)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    const DWORD currentThreadId       = GetCurrentThreadId();
    const unsigned int targetInFlight = registration->callbackThreadId.load(std::memory_order_acquire) == currentThreadId ? 1u : 0u;
    std::unique_lock drainLock(registration->drainMutex);
    registration->drainCv.wait(drainLock, [&]() noexcept { return registration->inFlight.load(std::memory_order_acquire) <= targetInFlight; });
    return S_OK;
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

    *schemaJsonUtf8 = StaticConfigurationSchema(_protocol);
    return S_OK;
}

const char* GetFileSystemCurlStaticConfigurationSchema(FileSystemCurlProtocol protocol) noexcept
{
    return FileSystemCurl::StaticConfigurationSchema(protocol);
}

const char* FileSystemCurl::StaticConfigurationSchema(FileSystemCurlProtocol protocol) noexcept
{
    switch (protocol)
    {
        case FileSystemCurlProtocol::Ftp: return kSchemaJsonFtp;
        case FileSystemCurlProtocol::Sftp: return kSchemaJsonSftp;
        case FileSystemCurlProtocol::Scp: return kSchemaJsonScp;
        case FileSystemCurlProtocol::Imap: return kSchemaJsonImap;
    }

    return nullptr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::string nextConfiguration = "{}";
    Common::Json::ObjectDocument parsed;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        nextConfiguration = configurationJsonUtf8;
        parsed            = Common::Json::ParseObjectDocument(nextConfiguration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        constexpr std::array<const char*, 2> secretMembers{{"defaultPassword", "sshKeyPassphrase"}};
        const std::optional<std::string> sanitized = Common::Json::WriteObjectWithoutMembers(parsed, secretMembers);
        if (! sanitized.has_value())
        {
            return E_OUTOFMEMORY;
        }
        nextConfiguration = sanitized.value();
    }

    std::lock_guard lock(_stateMutex);
    SecureWipe::SecureClear(_settings.defaultPassword);
    SecureWipe::SecureClear(_settings.sshKeyPassphrase);
    _settings          = {};
    _configurationJson = std::move(nextConfiguration);

    if (! parsed)
    {
        return S_OK;
    }

    yyjson_val* root = parsed.root;

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
        const uint64_t raw = connectTimeoutMs.value();
        if (raw >= 1u)
        {
            _settings.connectTimeoutMs = static_cast<unsigned long>(std::min<uint64_t>(raw, (std::numeric_limits<unsigned long>::max)()));
        }
    }

    const auto operationTimeoutMs = TryGetJsonUInt(root, "operationTimeoutMs");
    if (operationTimeoutMs.has_value())
    {
        _settings.operationTimeoutMs = static_cast<unsigned long>(std::min<uint64_t>(operationTimeoutMs.value(), (std::numeric_limits<unsigned long>::max)()));
    }

    const auto copyMoveMaxConcurrency = TryGetJsonUInt(root, "copyMoveMaxConcurrency");
    if (copyMoveMaxConcurrency.has_value())
    {
        const uint64_t value = copyMoveMaxConcurrency.value();
        if (value == 0)
        {
            _settings.copyMoveMaxConcurrency = 1;
        }
        else
        {
            _settings.copyMoveMaxConcurrency = static_cast<unsigned int>(std::min<uint64_t>(value, 8u));
        }
    }

    const auto deleteMaxConcurrency = TryGetJsonUInt(root, "deleteMaxConcurrency");
    if (deleteMaxConcurrency.has_value())
    {
        const uint64_t value = deleteMaxConcurrency.value();
        if (value == 0)
        {
            _settings.deleteMaxConcurrency = 1;
        }
        else
        {
            _settings.deleteMaxConcurrency = static_cast<unsigned int>(std::min<uint64_t>(value, 8u));
        }
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

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetPathCapabilities(const wchar_t* path, FileSystemOperation operation, const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }
    *jsonUtf8 = nullptr;
    if (path == nullptr || path[0] == L'\0' || operation < FILESYSTEM_COPY || operation > FILESYSTEM_CREATE_DIRECTORY)
    {
        return E_INVALIDARG;
    }

    std::lock_guard lock(_stateMutex);

    const unsigned int copyMoveMax = (_protocol == FileSystemCurlProtocol::Imap) ? 1u : std::clamp(_settings.copyMoveMaxConcurrency, 1u, 8u);
    const unsigned int deleteMax   = (_protocol == FileSystemCurlProtocol::Imap) ? 1u : std::clamp(_settings.deleteMaxConcurrency, 1u, 8u);

    if (_protocol == FileSystemCurlProtocol::Imap)
    {
        _capabilitiesJson = R"json(
{
  "version": 2,
  "pathProfile": "imap-mailbox",
  "rootId": "configured-mailbox-root",
  "operations": {
    "copy": false,
    "move": false,
    "nativeMove": false,
    "delete": false,
    "rename": false,
    "createDirectory": false,
    "properties": true,
    "read": true,
    "write": false,
    "recycle": false
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "transfer": {
    "export": { "copy": ["*"], "move": [] },
    "import": { "copy": [], "move": [] }
  },
  "identity": { "object": "imapUid", "revision": "none", "boundDelete": false, "conditionalDelete": false },
  "publication": { "exclusiveStage": false, "conditionalPublish": false, "committedSize": false },
  "links": { "preserveFileLink": false, "preserveDirectoryLink": false, "retargetInTree": false, "exactLinkRemoval": false },
  "metadata": { "motw": "reported-loss", "alternateStreams": "reported-loss", "extendedAttributes": "reported-loss", "sparse": "reported-loss", "efs": "reported-loss" },
  "verification": { "hostReadback": false, "providerProof": "none" },
  "cancellation": { "abort": false, "deadline": false, "routeClass": "uncontained", "providerWatchdogTimeoutMs": 0 },
  "names": {
    "pathTextStableIdentity": true,
    "comparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "notApplicable",
    "maxComponentUtf16": 255
  },
  "directories": { "model": "providerVirtual" }
}
)json";
    }
    else
    {
        _capabilitiesJson = std::format(
            R"json({{
  "version": 2,
  "pathProfile": "curl-remote-path",
  "rootId": "configured-connection-root",
  "operations": {{
    "copy": true,
    "move": true,
    "nativeMove": true,
    "delete": true,
    "rename": true,
    "createDirectory": true,
    "properties": true,
    "read": true,
    "write": true,
    "recycle": false
  }},
  "concurrency": {{
    "copyMoveMax": {},
    "deleteMax": {},
    "deleteRecycleBinMax": 1
  }},
  "transfer": {{
    "export": {{ "copy": ["*"], "move": [] }},
    "import": {{ "copy": ["*"], "move": [] }}
  }},
  "identity": {{ "object": "none", "revision": "none", "boundDelete": false, "conditionalDelete": false }},
  "publication": {{ "exclusiveStage": false, "conditionalPublish": false, "committedSize": false }},
  "links": {{ "preserveFileLink": false, "preserveDirectoryLink": false, "retargetInTree": false, "exactLinkRemoval": false }},
  "metadata": {{ "motw": "reported-loss", "alternateStreams": "reported-loss", "extendedAttributes": "reported-loss", "sparse": "reported-loss", "efs": "reported-loss" }},
  "verification": {{ "hostReadback": false, "providerProof": "none" }},
  "cancellation": {{ "abort": false, "deadline": true, "routeClass": "providerWatchdog", "providerWatchdogTimeoutMs": {} }},
  "names": {{
    "pathTextStableIdentity": true,
    "comparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "notApplicable",
    "maxComponentUtf16": 255
  }},
  "directories": {{ "model": "providerVirtual" }}
}})json",
            copyMoveMax,
            deleteMax,
            CurlProviderWatchdogTimeoutMs(_settings.connectTimeoutMs, _settings.operationTimeoutMs));
    }

    *jsonUtf8 = _capabilitiesJson.c_str();
    return S_OK;
}

HRESULT FileSystemCurl::BuildFileSystemRouteDescriptor(const wchar_t* path, FileSystemOperation operation, FileSystemRouteDescriptor& descriptor) noexcept
{
    static_cast<void>(operation);
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    std::lock_guard lock(_stateMutex);
    const bool imap          = _protocol == FileSystemCurlProtocol::Imap;
    descriptor               = {};
    descriptor.providerId    = _metaData.id != nullptr ? _metaData.id : L"";
    descriptor.pathProfileId = imap ? L"imap-mailbox" : L"curl-remote-path";
    descriptor.rootId        = imap ? L"configured-mailbox-root" : L"configured-connection-root";
    descriptor.availability  = FILESYSTEM_ROUTE_AVAILABLE;
    // R0f-Curl: FTP/SFTP/SCP transfers poll the operation control from libcurl's progress callback
    // and the transport timeouts bound a server that stops answering. IMAP stays read-only and
    // uncontained (no mutation is ever admitted on a mailbox).
    descriptor.cancellationRoute              = imap ? FILESYSTEM_CANCELLATION_UNCONTAINED : FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG;
    descriptor.providerWatchdogTimeoutMs      = imap ? 0u : CurlProviderWatchdogTimeoutMs(_settings.connectTimeoutMs, _settings.operationTimeoutMs);
    descriptor.cancellationDeadline           = ! imap;
    descriptor.namespaceKind                  = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
    descriptor.componentComparison            = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
    descriptor.caseOnlyRename                 = FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE;
    descriptor.copyMoveMaxConcurrency         = imap ? 1u : std::clamp(_settings.copyMoveMaxConcurrency, 1u, 8u);
    descriptor.deleteMaxConcurrency           = imap ? 1u : std::clamp(_settings.deleteMaxConcurrency, 1u, 8u);
    descriptor.deleteRecycleBinMaxConcurrency = 1u;
    descriptor.copyOperation                  = ! imap;
    descriptor.moveOperation                  = ! imap;
    descriptor.nativeMoveOperation            = ! imap;
    descriptor.deleteOperation                = ! imap;
    descriptor.renameOperation                = ! imap;
    descriptor.createDirectoryOperation       = ! imap;
    descriptor.propertiesOperation            = true;
    descriptor.readOperation                  = true;
    descriptor.writeOperation                 = ! imap;
    descriptor.exportCopyAll                  = true;
    // R0f-Curl: cross-provider Copy into FTP/SFTP/SCP is a full file-manager destination. Move
    // export/import stay denied until the source can be deleted conditionally (no bound delete).
    descriptor.importCopyAll = ! imap;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetTransferHints([[maybe_unused]] const wchar_t* path,
                                                           [[maybe_unused]] FileSystemOperation operationType,
                                                           [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                                           FileSystemTransferHints* hints) noexcept
{
    if (path == nullptr || path[0] == L'\0' || hints == nullptr)
    {
        return E_INVALIDARG;
    }
    if (hints->sizeBytes < sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass = (_protocol == FileSystemCurlProtocol::Imap) ? FILESYSTEM_TRANSFER_LATENCY_WAN : FILESYSTEM_TRANSFER_LATENCY_CLOUD;
    hints->flags =
        FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS | FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO | FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST;
    // CurlStreamingReader owns a 1 MiB ring and now fills it to the requested read size (or EOF).
    // Advertise the amount one Read can actually return instead of forcing an 8 MiB host buffer.
    hints->preferredBufferBytes      = 1u * 1024u * 1024u;
    hints->preferredProgressPeriodMs = 200u;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                                    FileSystemStorageCharacteristics* characteristics) noexcept
{
    if (path == nullptr || path[0] == L'\0' || characteristics == nullptr)
    {
        return E_INVALIDARG;
    }
    if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind = FILESYSTEM_STORAGE_CLOUD;
    characteristics->flags = FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO | FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
    characteristics->queueDepthHint               = 8u;
    characteristics->preferredCopyMoveConcurrency = 8u;
    characteristics->preferredDeleteConcurrency   = 8u;
    return S_OK;
}

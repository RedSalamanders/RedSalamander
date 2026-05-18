#include "FileSystemCurl.Internal.h"

#include <chrono>
#include <span>
#include <unordered_map>

namespace FileSystemCurlInternal
{
struct ImapMailboxEntry
{
    std::wstring name;
    bool noSelect = false;
};

static constexpr DWORD kImapFileAttributeMarked  = 0x02000000u;
static constexpr DWORD kImapFileAttributeUnread  = 0x04000000u;
static constexpr DWORD kImapFileAttributeDeleted = 0x08000000u;

[[nodiscard]] std::string_view ImapSchemeForConnection(const ConnectionInfo& conn) noexcept
{
    if (conn.port.has_value() && conn.port.value() == 993u)
    {
        return "imaps";
    }
    return "imap";
}

[[nodiscard]] std::string BuildImapUrl(const ConnectionInfo& conn, std::wstring_view mailboxPath) noexcept
{
    if (conn.host.empty())
    {
        return {};
    }

    std::string authority = conn.host;
    if (conn.port.has_value() && conn.port.value() != 0u)
    {
        const bool alreadyHasPort = authority.find(':') != std::string::npos && (authority.empty() || authority.front() != '[');
        if (! alreadyHasPort)
        {
            authority = std::format("{}:{}", authority, conn.port.value());
        }
    }

    std::string pathUtf8 = EscapeUrlPath(NormalizePluginPath(mailboxPath));
    if (pathUtf8.empty())
    {
        pathUtf8 = "/";
    }

    return std::format("{}://{}{}", ImapSchemeForConnection(conn), authority, pathUtf8);
}

[[nodiscard]] HRESULT CurlPerformImapCustomRequest(const ConnectionInfo& conn,
                                                   std::wstring_view mailboxPath,
                                                   std::string_view request,
                                                   std::string& outResponse) noexcept
{
    outResponse.clear();

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

    const std::string url = BuildImapUrl(conn, mailboxPath);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    std::string requestText;
    requestText.assign(request);

    if (requestText.empty())
    {
        return E_INVALIDARG;
    }

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_CUSTOMREQUEST, requestText.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &outResponse);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    char errorBuffer[CURL_ERROR_SIZE]{};
    curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

    ApplyCommonCurlOptions(curl.get(), conn, nullptr, false);
    if (ImapSchemeForConnection(conn) == "imap")
    {
        curl_easy_setopt(curl.get(), CURLOPT_USE_SSL, CURLUSESSL_TRY);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code != CURLE_OK)
    {
        long responseCode = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &responseCode);

        long osErrno = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_OS_ERRNO, &osErrno);

        std::string serverLine;
        size_t start = 0;
        while (start < outResponse.size())
        {
            size_t end = outResponse.find('\n', start);
            if (end == std::string::npos)
            {
                end = outResponse.size();
            }

            std::string_view line = std::string_view(outResponse).substr(start, end - start);
            if (! line.empty() && line.back() == '\r')
            {
                line.remove_suffix(1);
            }

            serverLine = TrimAscii(line);
            if (! serverLine.empty())
            {
                break;
            }
            start = end + 1u;
        }

        std::string errorText = TrimAscii(errorBuffer);

        constexpr size_t kMaxErrorText = 120;
        if (errorText.size() > kMaxErrorText)
        {
            errorText.resize(kMaxErrorText);
        }

        constexpr size_t kMaxServerLine = 120;
        if (serverLine.size() > kMaxServerLine)
        {
            serverLine.resize(kMaxServerLine);
        }

        Debug::Error(L"curl imap failed protocol={} curlCode={} ({}) responseCode={} osErrno={} error='{}' server='{}'",
                     ProtocolToDisplay(conn.protocol),
                     static_cast<unsigned long>(code),
                     Utf16FromUtf8(curl_easy_strerror(code)),
                     responseCode,
                     osErrno,
                     Utf16FromUtf8(errorText),
                     Utf16FromUtf8(serverLine));

        std::string requestShort         = requestText;
        constexpr size_t kMaxRequestText = 120;
        if (requestShort.size() > kMaxRequestText)
        {
            requestShort.resize(kMaxRequestText);
        }

        Debug::Error(L"curl imap ctx conn='{}' id='{}' user='{}' auth='{}' pwdPresent={} ignoreSslTrust={} url='{}' mailbox='{}' req='{}'",
                     conn.connectionName.empty() ? L"(none)" : conn.connectionName,
                     conn.connectionId,
                     Utf16FromUtf8(conn.user),
                     conn.connectionAuthMode,
                     conn.password.empty() ? 0 : 1,
                     conn.ignoreSslTrust ? 1 : 0,
                     Utf16FromUtf8(url),
                     mailboxPath,
                     Utf16FromUtf8(requestShort));
    }

    HRESULT resultHr = HResultFromCurl(code);
    if (code == CURLE_QUOTE_ERROR && ! outResponse.empty())
    {
        std::string firstLine;
        size_t start = 0;
        while (start < outResponse.size())
        {
            size_t end = outResponse.find('\n', start);
            if (end == std::string::npos)
            {
                end = outResponse.size();
            }

            std::string_view line = std::string_view(outResponse).substr(start, end - start);
            if (! line.empty() && line.back() == '\r')
            {
                line.remove_suffix(1);
            }

            firstLine = TrimAscii(line);
            if (! firstLine.empty())
            {
                break;
            }
            start = end + 1u;
        }

        std::string lower(firstLine);
        for (char& ch : lower)
        {
            if (ch >= 'A' && ch <= 'Z')
            {
                ch = static_cast<char>(ch - 'A' + 'a');
            }
        }

        if (lower.find("authenticationfailed") != std::string::npos || lower.find("login failed") != std::string::npos ||
            (lower.find("auth") != std::string::npos && lower.find("fail") != std::string::npos))
        {
            resultHr = HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
        }
        else if (lower.find("\\noperm") != std::string::npos || lower.find("permission denied") != std::string::npos ||
                 lower.find("access denied") != std::string::npos)
        {
            resultHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        else if (lower.find("nonexistent") != std::string::npos || lower.find("not found") != std::string::npos ||
                 lower.find("doesn't exist") != std::string::npos || lower.find("unknown mailbox") != std::string::npos)
        {
            resultHr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
    }

    return resultHr;
}

[[nodiscard]] bool TryParseImapQuotedString(std::string_view text, size_t& pos, std::string& out) noexcept
{
    out.clear();

    if (pos >= text.size() || text[pos] != '"')
    {
        return false;
    }

    ++pos;

    while (pos < text.size())
    {
        const char ch = text[pos++];
        if (ch == '"')
        {
            return true;
        }

        if (ch == '\\' && pos < text.size())
        {
            out.push_back(text[pos++]);
            continue;
        }

        out.push_back(ch);
    }

    return false;
}

[[nodiscard]] void SkipImapWhitespace(std::string_view text, size_t& pos) noexcept
{
    while (pos < text.size())
    {
        const char ch = text[pos];
        if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n')
        {
            break;
        }
        ++pos;
    }
}

[[nodiscard]] std::string_view ParseImapToken(std::string_view text, size_t& pos) noexcept
{
    while (pos < text.size() && (text[pos] == ' ' || text[pos] == '\t'))
    {
        ++pos;
    }

    const size_t start = pos;
    while (pos < text.size() && text[pos] != ' ' && text[pos] != '\t')
    {
        ++pos;
    }

    if (start >= text.size())
    {
        return {};
    }

    return text.substr(start, pos - start);
}

[[nodiscard]] std::wstring ImapMailboxNameToPluginMailboxName(std::wstring_view mailboxName, wchar_t delimiter) noexcept
{
    std::wstring out;
    out.assign(mailboxName);

    if (delimiter != L'\0' && delimiter != L'/')
    {
        for (wchar_t& ch : out)
        {
            if (ch == delimiter)
            {
                ch = L'/';
            }
        }
    }

    return out;
}

[[nodiscard]] std::wstring ImapMailboxNameToServerMailboxName(std::wstring_view mailboxName, wchar_t delimiter) noexcept
{
    std::wstring out;
    out.assign(mailboxName);

    if (delimiter != L'\0' && delimiter != L'/')
    {
        for (wchar_t& ch : out)
        {
            if (ch == L'/')
            {
                ch = delimiter;
            }
        }
    }

    return out;
}

[[nodiscard]] std::wstring ImapMailboxPathToServerMailboxPath(std::wstring_view mailboxPath, wchar_t delimiter) noexcept
{
    if (mailboxPath.empty())
    {
        return {};
    }

    if (mailboxPath == L"/")
    {
        return L"/";
    }

    std::wstring_view name = mailboxPath;
    if (! name.empty() && name.front() == L'/')
    {
        name.remove_prefix(1);
    }

    const std::wstring serverName = ImapMailboxNameToServerMailboxName(name, delimiter);
    if (serverName.empty())
    {
        return {};
    }

    std::wstring out;
    out.reserve(serverName.size() + 1u);
    out.push_back(L'/');
    out.append(serverName);
    return out;
}

[[nodiscard]] HRESULT ImapListMailboxes(const ConnectionInfo& conn, std::vector<ImapMailboxEntry>& out, wchar_t* outDelimiter) noexcept
{
    out.clear();
    if (outDelimiter)
    {
        *outDelimiter = L'\0';
    }

    std::string response;
    HRESULT hr = CurlPerformImapCustomRequest(conn, L"/", "LIST \"\" \"*\"", response);
    if (FAILED(hr))
    {
        return hr;
    }

    size_t start = 0;
    while (start < response.size())
    {
        size_t end = response.find('\n', start);
        if (end == std::string::npos)
        {
            end = response.size();
        }

        std::string_view line = std::string_view(response).substr(start, end - start);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }

        start = end + 1u;

        if (line.rfind("* LIST", 0) != 0)
        {
            continue;
        }

        size_t pos = 6;
        while (pos < line.size() && (line[pos] == ' ' || line[pos] == '\t'))
        {
            ++pos;
        }

        std::string flagsText;
        if (pos < line.size() && line[pos] == '(')
        {
            const size_t close = line.find(')', pos);
            if (close == std::string::npos)
            {
                continue;
            }

            flagsText.assign(line.substr(pos, close - pos + 1u));
            pos = close + 1u;
        }

        // delimiter
        std::string delimiterText;
        std::string_view delimTok = ParseImapToken(line, pos);
        if (! delimTok.empty() && delimTok != "NIL")
        {
            if (delimTok.front() == '"')
            {
                pos -= delimTok.size();
                static_cast<void>(TryParseImapQuotedString(line, pos, delimiterText));
            }
            else
            {
                delimiterText.assign(delimTok);
            }
        }

        wchar_t delimiter = L'\0';
        if (delimiterText.size() == 1u)
        {
            delimiter = static_cast<wchar_t>(static_cast<unsigned char>(delimiterText[0]));
        }

        if (outDelimiter && *outDelimiter == L'\0' && delimiter != L'\0')
        {
            *outDelimiter = delimiter;
        }

        std::string nameText;
        std::string_view nameTok = ParseImapToken(line, pos);
        if (nameTok.empty())
        {
            continue;
        }

        if (nameTok == "NIL")
        {
            continue;
        }

        if (! nameTok.empty() && nameTok.front() == '"')
        {
            pos -= nameTok.size();
            if (! TryParseImapQuotedString(line, pos, nameText))
            {
                continue;
            }
        }
        else
        {
            nameText.assign(nameTok);
        }

        if (nameText.empty())
        {
            continue;
        }

        ImapMailboxEntry entry{};
        const std::wstring serverName = Utf16FromUtf8(nameText);
        if (serverName.empty())
        {
            continue;
        }

        entry.name = ImapMailboxNameToPluginMailboxName(serverName, delimiter);
        if (entry.name.empty())
        {
            return E_OUTOFMEMORY;
        }

        std::string flagsLower(flagsText);
        for (char& ch : flagsLower)
        {
            if (ch >= 'A' && ch <= 'Z')
            {
                ch = static_cast<char>(ch - 'A' + 'a');
            }
        }
        entry.noSelect = flagsLower.find("\\noselect") != std::string::npos;

        out.push_back(std::move(entry));
    }

    return S_OK;
}

[[nodiscard]] HRESULT ImapGetHierarchyDelimiter(const ConnectionInfo& conn, wchar_t& outDelimiter) noexcept
{
    outDelimiter = L'\0';

    std::string response;
    HRESULT hr = CurlPerformImapCustomRequest(conn, L"/", "LIST \"\" \"\"", response);
    if (FAILED(hr))
    {
        return hr;
    }

    size_t start = 0;
    while (start < response.size())
    {
        size_t end = response.find('\n', start);
        if (end == std::string::npos)
        {
            end = response.size();
        }

        std::string_view line = std::string_view(response).substr(start, end - start);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }

        start = end + 1u;

        if (line.rfind("* LIST", 0) != 0)
        {
            continue;
        }

        size_t pos = 6;
        while (pos < line.size() && (line[pos] == ' ' || line[pos] == '\t'))
        {
            ++pos;
        }

        if (pos < line.size() && line[pos] == '(')
        {
            const size_t close = line.find(')', pos);
            if (close == std::string::npos)
            {
                continue;
            }
            pos = close + 1u;
        }

        std::string delimiterText;
        std::string_view delimTok = ParseImapToken(line, pos);
        if (! delimTok.empty() && delimTok != "NIL")
        {
            if (delimTok.front() == '"')
            {
                pos -= delimTok.size();
                if (! TryParseImapQuotedString(line, pos, delimiterText))
                {
                    continue;
                }
            }
            else
            {
                delimiterText.assign(delimTok);
            }
        }

        if (delimiterText.size() == 1u)
        {
            outDelimiter = static_cast<wchar_t>(static_cast<unsigned char>(delimiterText[0]));
        }

        return S_OK;
    }

    // Fallback: if LIST "" "" didn't return a delimiter, infer from regular mailbox listing.
    std::vector<ImapMailboxEntry> mailboxes;
    return ImapListMailboxes(conn, mailboxes, &outDelimiter);
}

[[nodiscard]] HRESULT ImapListMessageUids(const ConnectionInfo& conn, std::wstring_view mailboxName, wchar_t delimiter, std::vector<uint64_t>& outUids) noexcept
{
    outUids.clear();

    if (mailboxName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring mailboxPath;
    const std::wstring serverName = ImapMailboxNameToServerMailboxName(mailboxName, delimiter);
    if (serverName.empty())
    {
        return E_OUTOFMEMORY;
    }
    mailboxPath.reserve(serverName.size() + 1u);
    mailboxPath.push_back(L'/');
    mailboxPath.append(serverName);

    std::string response;
    HRESULT hr = CurlPerformImapCustomRequest(conn, mailboxPath, "UID SEARCH ALL", response);
    if (FAILED(hr))
    {
        return hr;
    }

    size_t start = 0;
    while (start < response.size())
    {
        size_t end = response.find('\n', start);
        if (end == std::string::npos)
        {
            end = response.size();
        }

        std::string_view line = std::string_view(response).substr(start, end - start);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }

        start = end + 1u;

        if (line.rfind("* SEARCH", 0) != 0)
        {
            continue;
        }

        size_t pos = 8;
        while (pos < line.size())
        {
            while (pos < line.size() && (line[pos] == ' ' || line[pos] == '\t'))
            {
                ++pos;
            }

            if (pos >= line.size())
            {
                break;
            }

            uint64_t value = 0;
            size_t digits  = 0;
            while (pos < line.size() && line[pos] >= '0' && line[pos] <= '9')
            {
                const uint64_t digit = static_cast<uint64_t>(line[pos] - '0');
                if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
                {
                    value  = 0;
                    digits = 0;
                    break;
                }
                value = (value * 10u) + digit;
                ++digits;
                ++pos;
            }

            if (digits > 0)
            {
                outUids.push_back(value);
            }
            else
            {
                while (pos < line.size() && line[pos] != ' ' && line[pos] != '\t')
                {
                    ++pos;
                }
            }
        }
    }

    return S_OK;
}

[[nodiscard]] bool TryExtractImapLiteralSize(std::string_view data, size_t& literalStart, uint64_t& literalSize) noexcept;

[[nodiscard]] constexpr char AsciiLower(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

[[nodiscard]] std::string_view TrimAsciiView(std::string_view text) noexcept
{
    while (! text.empty() && (text.front() == ' ' || text.front() == '\t' || text.front() == '\r' || text.front() == '\n'))
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && (text.back() == ' ' || text.back() == '\t' || text.back() == '\r' || text.back() == '\n'))
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] size_t FindAsciiNoCase(std::string_view haystack, std::string_view needle, size_t start) noexcept
{
    if (needle.empty())
    {
        return start <= haystack.size() ? start : std::string_view::npos;
    }

    for (size_t i = start; i + needle.size() <= haystack.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (AsciiLower(haystack[i + j]) != AsciiLower(needle[j]))
            {
                match = false;
                break;
            }
        }
        if (match)
        {
            return i;
        }
    }

    return std::string_view::npos;
}

[[nodiscard]] bool TryParseUintAfterKey(std::string_view text, std::string_view key, uint64_t& out) noexcept
{
    out = 0;

    const size_t keyPos = FindAsciiNoCase(text, key, 0);
    if (keyPos == std::string_view::npos)
    {
        return false;
    }

    size_t pos = keyPos + key.size();
    while (pos < text.size() && (text[pos] == ' ' || text[pos] == '\t'))
    {
        ++pos;
    }

    uint64_t value = 0;
    size_t digits  = 0;
    while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9')
    {
        const uint64_t digit = static_cast<uint64_t>(text[pos] - '0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
        ++digits;
        ++pos;
    }

    if (digits == 0)
    {
        return false;
    }

    out = value;
    return true;
}

[[nodiscard]] bool TryParseMonthAbbrev(std::string_view mon, int& outMonth) noexcept
{
    outMonth = 0;

    if (mon.size() < 3u)
    {
        return false;
    }

    const char a = AsciiLower(mon[0]);
    const char b = AsciiLower(mon[1]);
    const char c = AsciiLower(mon[2]);

    if (a == 'j' && b == 'a' && c == 'n')
    {
        outMonth = 1;
        return true;
    }
    if (a == 'f' && b == 'e' && c == 'b')
    {
        outMonth = 2;
        return true;
    }
    if (a == 'm' && b == 'a' && c == 'r')
    {
        outMonth = 3;
        return true;
    }
    if (a == 'a' && b == 'p' && c == 'r')
    {
        outMonth = 4;
        return true;
    }
    if (a == 'm' && b == 'a' && c == 'y')
    {
        outMonth = 5;
        return true;
    }
    if (a == 'j' && b == 'u' && c == 'n')
    {
        outMonth = 6;
        return true;
    }
    if (a == 'j' && b == 'u' && c == 'l')
    {
        outMonth = 7;
        return true;
    }
    if (a == 'a' && b == 'u' && c == 'g')
    {
        outMonth = 8;
        return true;
    }
    if (a == 's' && b == 'e' && c == 'p')
    {
        outMonth = 9;
        return true;
    }
    if (a == 'o' && b == 'c' && c == 't')
    {
        outMonth = 10;
        return true;
    }
    if (a == 'n' && b == 'o' && c == 'v')
    {
        outMonth = 11;
        return true;
    }
    if (a == 'd' && b == 'e' && c == 'c')
    {
        outMonth = 12;
        return true;
    }

    return false;
}

[[nodiscard]] bool TryParseTimeZoneOffsetMinutes(std::string_view tz, int& outOffsetMinutes) noexcept
{
    outOffsetMinutes = 0;

    if (tz.empty())
    {
        return true;
    }

    if (tz.size() == 1u && AsciiLower(tz[0]) == 'z')
    {
        outOffsetMinutes = 0;
        return true;
    }

    if (tz.size() == 2u && AsciiLower(tz[0]) == 'u' && AsciiLower(tz[1]) == 't')
    {
        outOffsetMinutes = 0;
        return true;
    }

    if (tz.size() == 3u)
    {
        const char a = AsciiLower(tz[0]);
        const char b = AsciiLower(tz[1]);
        const char c = AsciiLower(tz[2]);
        if ((a == 'u' && b == 't' && c == 'c') || (a == 'g' && b == 'm' && c == 't'))
        {
            outOffsetMinutes = 0;
            return true;
        }
    }

    if (tz.size() >= 5u && (tz[0] == '+' || tz[0] == '-') && tz[1] >= '0' && tz[1] <= '9' && tz[2] >= '0' && tz[2] <= '9' && tz[3] >= '0' && tz[3] <= '9' &&
        tz[4] >= '0' && tz[4] <= '9')
    {
        const int sign   = (tz[0] == '-') ? -1 : 1;
        const int hours  = static_cast<int>((tz[1] - '0') * 10 + (tz[2] - '0'));
        const int mins   = static_cast<int>((tz[3] - '0') * 10 + (tz[4] - '0'));
        outOffsetMinutes = sign * (hours * 60 + mins);
        return true;
    }

    // Common timezone abbreviations (RFC5322 obs-zone).
    const char a = AsciiLower(tz[0]);
    const char b = tz.size() > 1u ? AsciiLower(tz[1]) : '\0';
    const char c = tz.size() > 2u ? AsciiLower(tz[2]) : '\0';

    if (a == 'e' && b == 's' && c == 't')
    {
        outOffsetMinutes = -5 * 60;
        return true;
    }
    if (a == 'e' && b == 'd' && c == 't')
    {
        outOffsetMinutes = -4 * 60;
        return true;
    }
    if (a == 'c' && b == 's' && c == 't')
    {
        outOffsetMinutes = -6 * 60;
        return true;
    }
    if (a == 'c' && b == 'd' && c == 't')
    {
        outOffsetMinutes = -5 * 60;
        return true;
    }
    if (a == 'm' && b == 's' && c == 't')
    {
        outOffsetMinutes = -7 * 60;
        return true;
    }
    if (a == 'm' && b == 'd' && c == 't')
    {
        outOffsetMinutes = -6 * 60;
        return true;
    }
    if (a == 'p' && b == 's' && c == 't')
    {
        outOffsetMinutes = -8 * 60;
        return true;
    }
    if (a == 'p' && b == 'd' && c == 't')
    {
        outOffsetMinutes = -7 * 60;
        return true;
    }

    return true;
}

[[nodiscard]] bool TrySystemTimeToFileTimeUtc(const SYSTEMTIME& stUtc, int offsetMinutes, __int64& outFileTimeUtc) noexcept
{
    outFileTimeUtc = 0;

    FILETIME ft{};
    if (! SystemTimeToFileTime(&stUtc, &ft))
    {
        return false;
    }

    ULARGE_INTEGER uli{};
    uli.LowPart  = ft.dwLowDateTime;
    uli.HighPart = ft.dwHighDateTime;

    const int64_t ticks       = static_cast<int64_t>(uli.QuadPart);
    const int64_t adjustTicks = static_cast<int64_t>(offsetMinutes) * 60LL * 10'000'000LL;

    const int64_t adjusted = ticks - adjustTicks;
    if (adjusted <= 0)
    {
        return false;
    }

    outFileTimeUtc = static_cast<__int64>(adjusted);
    return true;
}

[[nodiscard]] bool TryParseImapInternalDateToFileTime(std::string_view text, __int64& outFileTimeUtc) noexcept
{
    outFileTimeUtc = 0;

    // Example: 17-Jul-1996 02:44:25 -0700
    // See IMAP INTERNALDATE format (RFC3501).
    size_t pos = 0;

    // day (1-2 digits)
    int day = 0;
    {
        int value     = 0;
        size_t digits = 0;
        while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9' && digits < 2u)
        {
            value = (value * 10) + static_cast<int>(text[pos] - '0');
            ++pos;
            ++digits;
        }
        if (digits == 0 || pos >= text.size() || text[pos] != '-')
        {
            return false;
        }
        day = value;
        ++pos;
    }

    int month = 0;
    if (pos + 3u > text.size() || ! TryParseMonthAbbrev(text.substr(pos, 3), month))
    {
        return false;
    }
    pos += 3u;

    if (pos >= text.size() || text[pos] != '-')
    {
        return false;
    }
    ++pos;

    int year = 0;
    {
        int value     = 0;
        size_t digits = 0;
        while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9' && digits < 4u)
        {
            value = (value * 10) + static_cast<int>(text[pos] - '0');
            ++pos;
            ++digits;
        }
        if (digits < 2u)
        {
            return false;
        }
        year = value;
    }

    while (pos < text.size() && (text[pos] == ' ' || text[pos] == '\t'))
    {
        ++pos;
    }

    int hour   = 0;
    int minute = 0;
    int second = 0;
    {
        if (pos + 5u > text.size())
        {
            return false;
        }

        const auto parse2 = [&](int& out) noexcept -> bool
        {
            if (pos + 2u > text.size() || text[pos] < '0' || text[pos] > '9' || text[pos + 1u] < '0' || text[pos + 1u] > '9')
            {
                return false;
            }
            out = (text[pos] - '0') * 10 + (text[pos + 1u] - '0');
            pos += 2u;
            return true;
        };

        if (! parse2(hour) || pos >= text.size() || text[pos] != ':')
        {
            return false;
        }
        ++pos;
        if (! parse2(minute))
        {
            return false;
        }

        if (pos < text.size() && text[pos] == ':')
        {
            ++pos;
            if (! parse2(second))
            {
                return false;
            }
        }
    }

    while (pos < text.size() && (text[pos] == ' ' || text[pos] == '\t'))
    {
        ++pos;
    }

    int offsetMinutes = 0;
    if (pos < text.size())
    {
        const std::string_view tz = TrimAsciiView(text.substr(pos));
        static_cast<void>(TryParseTimeZoneOffsetMinutes(tz, offsetMinutes));
    }

    SYSTEMTIME st{};
    st.wYear   = static_cast<WORD>(year);
    st.wMonth  = static_cast<WORD>(month);
    st.wDay    = static_cast<WORD>(day);
    st.wHour   = static_cast<WORD>(hour);
    st.wMinute = static_cast<WORD>(minute);
    st.wSecond = static_cast<WORD>(second);

    return TrySystemTimeToFileTimeUtc(st, offsetMinutes, outFileTimeUtc);
}

[[nodiscard]] bool TryParseRfc5322DateToFileTime(std::string_view text, __int64& outFileTimeUtc) noexcept
{
    outFileTimeUtc = 0;

    // Drop comments "(...)".
    const size_t comment = text.find('(');
    if (comment != std::string_view::npos)
    {
        text = text.substr(0, comment);
    }
    text = TrimAsciiView(text);

    if (text.empty())
    {
        return false;
    }

    // Split on whitespace (no allocations).
    std::array<std::string_view, 12> parts{};
    size_t partCount = 0;

    size_t pos = 0;
    while (pos < text.size() && partCount < parts.size())
    {
        while (pos < text.size() && (text[pos] == ' ' || text[pos] == '\t'))
        {
            ++pos;
        }
        if (pos >= text.size())
        {
            break;
        }
        size_t end = pos;
        while (end < text.size() && text[end] != ' ' && text[end] != '\t')
        {
            ++end;
        }
        parts[partCount++] = text.substr(pos, end - pos);
        pos                = end;
    }

    size_t idx = 0;
    if (partCount == 0)
    {
        return false;
    }

    // Optional day-of-week (Mon, Tue, ...).
    if (parts[idx].size() >= 4u && parts[idx].back() == ',')
    {
        ++idx;
    }

    if (idx + 3u >= partCount)
    {
        return false;
    }

    const auto parseInt = [](std::string_view tok, int& out) noexcept -> bool
    {
        out = 0;
        if (tok.empty())
        {
            return false;
        }
        int value = 0;
        for (const char ch : tok)
        {
            if (ch < '0' || ch > '9')
            {
                return false;
            }
            value = (value * 10) + static_cast<int>(ch - '0');
        }
        out = value;
        return true;
    };

    int day = 0;
    if (! parseInt(parts[idx++], day))
    {
        return false;
    }

    int month = 0;
    if (! TryParseMonthAbbrev(parts[idx++], month))
    {
        return false;
    }

    int year = 0;
    if (! parseInt(parts[idx++], year))
    {
        return false;
    }
    if (year < 100)
    {
        year = (year >= 70) ? (1900 + year) : (2000 + year);
    }

    // time
    int hour   = 0;
    int minute = 0;
    int second = 0;
    {
        const std::string_view t = parts[idx++];
        size_t tp                = 0;
        const auto parse2        = [&](int& out) noexcept -> bool
        {
            if (tp + 2u > t.size() || t[tp] < '0' || t[tp] > '9' || t[tp + 1u] < '0' || t[tp + 1u] > '9')
            {
                return false;
            }
            out = (t[tp] - '0') * 10 + (t[tp + 1u] - '0');
            tp += 2u;
            return true;
        };

        if (! parse2(hour) || tp >= t.size() || t[tp] != ':')
        {
            return false;
        }
        ++tp;
        if (! parse2(minute))
        {
            return false;
        }

        if (tp < t.size() && t[tp] == ':')
        {
            ++tp;
            if (! parse2(second))
            {
                return false;
            }
        }
    }

    int offsetMinutes = 0;
    if (idx < partCount)
    {
        const std::string_view tz = parts[idx];
        static_cast<void>(TryParseTimeZoneOffsetMinutes(tz, offsetMinutes));
    }

    SYSTEMTIME st{};
    st.wYear   = static_cast<WORD>(year);
    st.wMonth  = static_cast<WORD>(month);
    st.wDay    = static_cast<WORD>(day);
    st.wHour   = static_cast<WORD>(hour);
    st.wMinute = static_cast<WORD>(minute);
    st.wSecond = static_cast<WORD>(second);

    return TrySystemTimeToFileTimeUtc(st, offsetMinutes, outFileTimeUtc);
}

[[nodiscard]] constexpr bool IsImapWhitespaceChar(char ch) noexcept
{
    return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n';
}

[[nodiscard]] bool TryParseImapLiteralString(std::string_view text, size_t& pos, std::string& out) noexcept
{
    out.clear();

    if (pos >= text.size())
    {
        return false;
    }

    const bool hasTildePrefix = text[pos] == '~' && (pos + 1u) < text.size() && text[pos + 1u] == '{';
    const size_t bracePos     = hasTildePrefix ? (pos + 1u) : pos;
    if (bracePos >= text.size() || text[bracePos] != '{')
    {
        return false;
    }

    size_t p = bracePos + 1u;
    if (p >= text.size())
    {
        return false;
    }

    uint64_t value = 0;
    size_t digits  = 0;
    while (p < text.size() && text[p] >= '0' && text[p] <= '9')
    {
        const uint64_t digit = static_cast<uint64_t>(text[p] - '0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
        ++digits;
        ++p;
    }

    if (digits == 0)
    {
        return false;
    }

    if (p < text.size() && text[p] == '+')
    {
        ++p;
    }

    if (p >= text.size() || text[p] != '}')
    {
        return false;
    }

    size_t afterBrace = p + 1u;
    if (afterBrace >= text.size())
    {
        return false;
    }

    size_t literalStart = std::string_view::npos;
    if (text[afterBrace] == '\n')
    {
        literalStart = afterBrace + 1u;
    }
    else if (text[afterBrace] == '\r' && (afterBrace + 1u) < text.size() && text[afterBrace + 1u] == '\n')
    {
        literalStart = afterBrace + 2u;
    }
    else
    {
        return false;
    }

    if (value > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return false;
    }

    const size_t literalSize = static_cast<size_t>(value);
    if (literalStart > text.size() || literalStart + literalSize > text.size())
    {
        return false;
    }

    out.assign(text.substr(literalStart, literalSize));

    pos = literalStart + literalSize;
    return true;
}

[[nodiscard]] bool TryParseImapNString(std::string_view text, size_t& pos, std::string& out) noexcept
{
    out.clear();
    SkipImapWhitespace(text, pos);

    if (pos >= text.size())
    {
        return false;
    }

    if (pos + 3u <= text.size())
    {
        const std::string_view maybeNil = text.substr(pos, 3u);
        if (FindAsciiNoCase(maybeNil, "NIL", 0) == 0)
        {
            const size_t after = pos + 3u;
            if (after >= text.size() || IsImapWhitespaceChar(text[after]) || text[after] == ')')
            {
                pos = after;
                return true;
            }
        }
    }

    if (text[pos] == '"')
    {
        return TryParseImapQuotedString(text, pos, out);
    }

    if (text[pos] == '{' || (text[pos] == '~' && (pos + 1u) < text.size() && text[pos + 1u] == '{'))
    {
        return TryParseImapLiteralString(text, pos, out);
    }

    const size_t start = pos;
    while (pos < text.size())
    {
        const char ch = text[pos];
        if (IsImapWhitespaceChar(ch) || ch == ')')
        {
            break;
        }
        ++pos;
    }

    if (pos <= start)
    {
        return false;
    }

    out.assign(text.substr(start, pos - start));
    return true;
}

[[nodiscard]] bool TrySkipImapParenthesized(std::string_view text, size_t& pos) noexcept
{
    SkipImapWhitespace(text, pos);
    if (pos >= text.size() || text[pos] != '(')
    {
        return false;
    }

    bool inQuote = false;
    int depth    = 0;

    size_t i = pos;
    while (i < text.size())
    {
        const char ch = text[i];
        if (inQuote)
        {
            if (ch == '\\' && (i + 1u) < text.size())
            {
                i += 2u;
                continue;
            }
            if (ch == '"')
            {
                inQuote = false;
            }
            ++i;
            continue;
        }

        if (ch == '"')
        {
            inQuote = true;
            ++i;
            continue;
        }

        const bool hasTildeLiteralPrefix = ch == '~' && (i + 1u) < text.size() && text[i + 1u] == '{';
        if (ch == '{' || hasTildeLiteralPrefix)
        {
            size_t p = (hasTildeLiteralPrefix ? (i + 1u) : i) + 1u;
            if (p >= text.size())
            {
                return false;
            }

            uint64_t value = 0;
            size_t digits  = 0;
            while (p < text.size() && text[p] >= '0' && text[p] <= '9')
            {
                const uint64_t digit = static_cast<uint64_t>(text[p] - '0');
                if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
                {
                    return false;
                }
                value = (value * 10u) + digit;
                ++digits;
                ++p;
            }

            if (digits == 0)
            {
                ++i;
                continue;
            }

            if (p < text.size() && text[p] == '+')
            {
                ++p;
            }

            if (p >= text.size() || text[p] != '}')
            {
                ++i;
                continue;
            }

            size_t afterBrace = p + 1u;
            if (afterBrace >= text.size())
            {
                return false;
            }

            size_t literalStart = std::string_view::npos;
            if (text[afterBrace] == '\n')
            {
                literalStart = afterBrace + 1u;
            }
            else if (text[afterBrace] == '\r' && (afterBrace + 1u) < text.size() && text[afterBrace + 1u] == '\n')
            {
                literalStart = afterBrace + 2u;
            }
            else
            {
                ++i;
                continue;
            }

            if (value > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
            {
                return false;
            }

            const size_t literalSize = static_cast<size_t>(value);
            if (literalStart > text.size() || literalStart + literalSize > text.size())
            {
                return false;
            }

            i = literalStart + literalSize;
            continue;
        }

        if (ch == '(')
        {
            ++depth;
            ++i;
            continue;
        }

        if (ch == ')' && depth > 0)
        {
            --depth;
            ++i;
            if (depth == 0)
            {
                pos = i;
                return true;
            }
            continue;
        }

        ++i;
    }

    return false;
}

[[nodiscard]] bool TrySkipImapAddressList(std::string_view text, size_t& pos) noexcept
{
    SkipImapWhitespace(text, pos);

    if (pos >= text.size())
    {
        return false;
    }

    if (pos + 3u <= text.size())
    {
        const std::string_view maybeNil = text.substr(pos, 3u);
        if (FindAsciiNoCase(maybeNil, "NIL", 0) == 0)
        {
            const size_t after = pos + 3u;
            if (after >= text.size() || IsImapWhitespaceChar(text[after]) || text[after] == ')')
            {
                pos = after;
                return true;
            }
        }
    }

    return TrySkipImapParenthesized(text, pos);
}

[[nodiscard]] bool TryParseImapEnvelopeAddress(std::string_view text, size_t& pos, std::string& outAddrSpec) noexcept
{
    outAddrSpec.clear();

    SkipImapWhitespace(text, pos);
    if (pos >= text.size() || text[pos] != '(')
    {
        return false;
    }
    ++pos;

    std::string name;
    std::string adl;
    std::string mailbox;
    std::string host;
    if (! TryParseImapNString(text, pos, name) || ! TryParseImapNString(text, pos, adl) || ! TryParseImapNString(text, pos, mailbox) ||
        ! TryParseImapNString(text, pos, host))
    {
        return false;
    }

    SkipImapWhitespace(text, pos);
    if (pos >= text.size() || text[pos] != ')')
    {
        return false;
    }
    ++pos;

    if (mailbox.empty() || host.empty())
    {
        return true;
    }

    outAddrSpec.reserve(mailbox.size() + host.size() + 1u);
    outAddrSpec.assign(mailbox);
    outAddrSpec.push_back('@');
    outAddrSpec.append(host);
    return true;
}

[[nodiscard]] bool TryParseImapEnvelopeAddressListFirstAddr(std::string_view text, size_t& pos, std::string& outAddrSpec) noexcept
{
    outAddrSpec.clear();
    SkipImapWhitespace(text, pos);

    if (pos >= text.size())
    {
        return false;
    }

    if (pos + 3u <= text.size())
    {
        const std::string_view maybeNil = text.substr(pos, 3u);
        if (FindAsciiNoCase(maybeNil, "NIL", 0) == 0)
        {
            const size_t after = pos + 3u;
            if (after >= text.size() || IsImapWhitespaceChar(text[after]) || text[after] == ')')
            {
                pos = after;
                return true;
            }
        }
    }

    if (text[pos] != '(')
    {
        return false;
    }
    ++pos;

    while (pos < text.size())
    {
        SkipImapWhitespace(text, pos);
        if (pos >= text.size())
        {
            return false;
        }

        if (text[pos] == ')')
        {
            ++pos;
            return true;
        }

        if (text[pos] != '(')
        {
            std::string dummy;
            if (! TryParseImapNString(text, pos, dummy))
            {
                return false;
            }
            continue;
        }

        if (outAddrSpec.empty())
        {
            if (! TryParseImapEnvelopeAddress(text, pos, outAddrSpec))
            {
                return false;
            }
        }
        else
        {
            if (! TrySkipImapParenthesized(text, pos))
            {
                return false;
            }
        }
    }

    return false;
}

struct ImapEnvelopeFields
{
    std::string date;
    std::string subject;
    std::string fromAddrSpec;
};

[[nodiscard]] bool TryExtractEnvelopeFields(std::string_view fetchText, ImapEnvelopeFields& out) noexcept
{
    out = {};

    const size_t envPos = FindAsciiNoCase(fetchText, "ENVELOPE", 0);
    if (envPos == std::string_view::npos)
    {
        return false;
    }

    size_t pos = envPos + 8u;
    SkipImapWhitespace(fetchText, pos);
    if (pos >= fetchText.size() || fetchText[pos] != '(')
    {
        return false;
    }
    ++pos;

    if (! TryParseImapNString(fetchText, pos, out.date))
    {
        return false;
    }
    if (! TryParseImapNString(fetchText, pos, out.subject))
    {
        return false;
    }
    if (! TryParseImapEnvelopeAddressListFirstAddr(fetchText, pos, out.fromAddrSpec))
    {
        return false;
    }

    // sender, reply-to, to, cc, bcc
    for (int i = 0; i < 5; ++i)
    {
        if (! TrySkipImapAddressList(fetchText, pos))
        {
            return false;
        }
    }

    // in-reply-to, message-id
    std::string dummy;
    if (! TryParseImapNString(fetchText, pos, dummy))
    {
        return false;
    }
    if (! TryParseImapNString(fetchText, pos, dummy))
    {
        return false;
    }

    SkipImapWhitespace(fetchText, pos);
    if (pos >= fetchText.size() || fetchText[pos] != ')')
    {
        return false;
    }

    return true;
}

struct ImapHeaderFields
{
    std::string subject;
    std::string from;
    std::string date;
};

[[nodiscard]] bool TryExtractHeaderFields(std::string_view headerBlock, ImapHeaderFields& out) noexcept
{
    out = {};

    {
        std::string* current = nullptr;

        size_t start = 0;
        while (start < headerBlock.size())
        {
            size_t end = headerBlock.find('\n', start);
            if (end == std::string_view::npos)
            {
                end = headerBlock.size();
            }

            std::string_view line = headerBlock.substr(start, end - start);
            if (! line.empty() && line.back() == '\r')
            {
                line.remove_suffix(1);
            }

            start = end + 1u;

            if (line.empty())
            {
                break;
            }

            if ((line.front() == ' ' || line.front() == '\t'))
            {
                if (current)
                {
                    const std::string_view cont = TrimAsciiView(line);
                    if (! cont.empty())
                    {
                        current->push_back(' ');
                        current->append(cont);
                    }
                }
                continue;
            }

            const size_t colon = line.find(':');
            if (colon == std::string_view::npos)
            {
                current = nullptr;
                continue;
            }

            const std::string_view name  = TrimAsciiView(line.substr(0, colon));
            const std::string_view value = TrimAsciiView(line.substr(colon + 1u));

            if (FindAsciiNoCase(name, "subject", 0) == 0 && name.size() == 7u)
            {
                out.subject.assign(value);
                current = &out.subject;
            }
            else if (FindAsciiNoCase(name, "from", 0) == 0 && name.size() == 4u)
            {
                out.from.assign(value);
                current = &out.from;
            }
            else if (FindAsciiNoCase(name, "date", 0) == 0 && name.size() == 4u)
            {
                out.date.assign(value);
                current = &out.date;
            }
            else
            {
                current = nullptr;
            }
        }

        return true;
    }
}

[[nodiscard]] std::wstring Utf16FromImapHeaderValue(std::string_view text) noexcept
{
    const std::wstring wide = Utf16FromUtf8(text);
    if (! wide.empty() || text.empty())
    {
        return wide;
    }

    std::wstring out;
    out.reserve(text.size());
    for (const char ch : text)
    {
        out.push_back(static_cast<wchar_t>(static_cast<unsigned char>(ch)));
    }
    return out;
}

[[nodiscard]] bool IsAsciiWhitespace(char ch) noexcept
{
    return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n';
}

[[nodiscard]] std::wstring ExtractEmailAddressFromFromHeader(std::string_view fromHeader) noexcept
{
    // Prefer addr-spec inside "<...>".
    const size_t lt = fromHeader.find('<');
    if (lt != std::string_view::npos)
    {
        const size_t gt = fromHeader.find('>', lt + 1u);
        if (gt != std::string_view::npos && gt > lt + 1u)
        {
            std::string_view inside = TrimAsciiView(fromHeader.substr(lt + 1u, gt - lt - 1u));
            if (inside.find('@') != std::string_view::npos)
            {
                return Utf16FromImapHeaderValue(inside);
            }
        }
    }

    // Fallback: find first token containing '@'.
    for (size_t at = 0; at < fromHeader.size(); ++at)
    {
        if (fromHeader[at] != '@')
        {
            continue;
        }

        size_t start = at;
        while (start > 0)
        {
            const char ch = fromHeader[start - 1u];
            if (IsAsciiWhitespace(ch) || ch == ',' || ch == ';' || ch == '"' || ch == '\'' || ch == '<' || ch == '>' || ch == '(' || ch == ')')
            {
                break;
            }
            --start;
        }

        size_t end = at + 1u;
        while (end < fromHeader.size())
        {
            const char ch = fromHeader[end];
            if (IsAsciiWhitespace(ch) || ch == ',' || ch == ';' || ch == '"' || ch == '\'' || ch == '<' || ch == '>' || ch == '(' || ch == ')')
            {
                break;
            }
            ++end;
        }

        if (end > start)
        {
            const std::string_view token = fromHeader.substr(start, end - start);
            if (token.find('@') != std::string_view::npos)
            {
                return Utf16FromImapHeaderValue(token);
            }
        }
    }

    return {};
}

[[nodiscard]] size_t FindImapUntaggedLine(std::string_view response, size_t start) noexcept
{
    for (size_t i = start; i + 1u < response.size(); ++i)
    {
        if ((i == 0 || response[i - 1u] == '\n') && response[i] == '*' && response[i + 1u] == ' ')
        {
            return i;
        }
    }
    return std::string_view::npos;
}

[[nodiscard]] size_t FindImapUntaggedFetchLine(std::string_view response, size_t start) noexcept
{
    for (size_t i = start; i + 1u < response.size(); ++i)
    {
        if (! ((i == 0 || response[i - 1u] == '\n') && response[i] == '*' && response[i + 1u] == ' '))
        {
            continue;
        }

        size_t p = i + 2u;
        while (p < response.size() && (response[p] == ' ' || response[p] == '\t'))
        {
            ++p;
        }

        if (p >= response.size() || response[p] < '0' || response[p] > '9')
        {
            continue;
        }
        while (p < response.size() && response[p] >= '0' && response[p] <= '9')
        {
            ++p;
        }

        while (p < response.size() && (response[p] == ' ' || response[p] == '\t'))
        {
            ++p;
        }

        if (p + 4u >= response.size())
        {
            continue;
        }

        if (AsciiLower(response[p]) != 'f' || AsciiLower(response[p + 1u]) != 'e' || AsciiLower(response[p + 2u]) != 't' ||
            AsciiLower(response[p + 3u]) != 'c' || AsciiLower(response[p + 4u]) != 'h')
        {
            continue;
        }

        return i;
    }
    return std::string_view::npos;
}

[[nodiscard]] bool TryConsumeImapUntaggedFetchResponse(std::string_view response,
                                                       size_t msgStart,
                                                       size_t& outNextPos,
                                                       std::string_view& outPrefix,
                                                       std::string_view& outHeaderBlock,
                                                       std::string_view& outSuffix) noexcept
{
    outNextPos     = 0;
    outPrefix      = {};
    outHeaderBlock = {};
    outSuffix      = {};

    if (msgStart >= response.size())
    {
        return false;
    }

    const size_t openParen = response.find('(', msgStart);
    if (openParen == std::string_view::npos)
    {
        return false;
    }

    size_t headerStartAbs = std::string_view::npos;
    size_t headerEndAbs   = std::string_view::npos;

    bool inQuote   = false;
    int parenDepth = 0;

    size_t i = openParen;
    while (i < response.size())
    {
        const char ch = response[i];

        if (inQuote)
        {
            if (ch == '\\' && (i + 1u) < response.size())
            {
                i += 2u;
                continue;
            }
            if (ch == '"')
            {
                inQuote = false;
            }
            ++i;
            continue;
        }

        if (ch == '"')
        {
            inQuote = true;
            ++i;
            continue;
        }

        const bool hasTildeLiteralPrefix = ch == '~' && (i + 1u) < response.size() && response[i + 1u] == '{';
        if (ch == '{' || hasTildeLiteralPrefix)
        {
            const size_t bracePos = hasTildeLiteralPrefix ? (i + 1u) : i;

            size_t pos = bracePos + 1u;
            if (pos >= response.size())
            {
                break;
            }

            uint64_t value = 0;
            size_t digits  = 0;
            while (pos < response.size() && response[pos] >= '0' && response[pos] <= '9')
            {
                const uint64_t digit = static_cast<uint64_t>(response[pos] - '0');
                if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
                {
                    break;
                }
                value = (value * 10u) + digit;
                ++digits;
                ++pos;
            }

            if (digits == 0)
            {
                ++i;
                continue;
            }

            if (pos < response.size() && response[pos] == '+')
            {
                ++pos;
            }

            if (pos >= response.size() || response[pos] != '}')
            {
                ++i;
                continue;
            }

            size_t afterBrace = pos + 1u;
            if (afterBrace >= response.size())
            {
                break;
            }

            size_t literalStartAbs = std::string_view::npos;
            if (response[afterBrace] == '\n')
            {
                literalStartAbs = afterBrace + 1u;
            }
            else if (response[afterBrace] == '\r' && (afterBrace + 1u) < response.size() && response[afterBrace + 1u] == '\n')
            {
                literalStartAbs = afterBrace + 2u;
            }
            else
            {
                ++i;
                continue;
            }

            if (value > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
            {
                break;
            }

            const size_t literalSize = static_cast<size_t>(value);
            if (literalStartAbs > response.size() || literalStartAbs + literalSize > response.size())
            {
                break;
            }

            if (headerStartAbs == std::string_view::npos)
            {
                const size_t contextStart  = (bracePos > 256u) ? (bracePos - 256u) : msgStart;
                const std::string_view ctx = response.substr(contextStart, bracePos - contextStart);
                if (FindAsciiNoCase(ctx, "HEADER.FIELDS", 0) != std::string_view::npos)
                {
                    headerStartAbs = literalStartAbs;
                    headerEndAbs   = literalStartAbs + literalSize;
                    outHeaderBlock = response.substr(literalStartAbs, literalSize);
                }
            }

            i = literalStartAbs + literalSize;
            continue;
        }

        if (ch == '(')
        {
            ++parenDepth;
            ++i;
            continue;
        }

        if (ch == ')' && parenDepth > 0)
        {
            --parenDepth;
            ++i;

            if (parenDepth == 0)
            {
                const size_t lineEnd = response.find('\n', i);
                if (lineEnd == std::string_view::npos)
                {
                    break;
                }

                outNextPos = lineEnd + 1u;
                if (headerStartAbs != std::string_view::npos && headerEndAbs != std::string_view::npos)
                {
                    outPrefix = response.substr(msgStart, headerStartAbs - msgStart);
                    outSuffix = response.substr(headerEndAbs, outNextPos - headerEndAbs);
                }
                else
                {
                    outPrefix = response.substr(msgStart, outNextPos - msgStart);
                }

                return true;
            }
            continue;
        }

        ++i;
    }

    const size_t nextFetch = FindImapUntaggedFetchLine(response, msgStart + 1u);
    outNextPos             = nextFetch == std::string_view::npos ? response.size() : nextFetch;

    if (outNextPos <= msgStart)
    {
        return false;
    }

    if (headerStartAbs != std::string_view::npos && headerEndAbs != std::string_view::npos && headerEndAbs <= outNextPos)
    {
        outPrefix = response.substr(msgStart, headerStartAbs - msgStart);
        outSuffix = response.substr(headerEndAbs, outNextPos - headerEndAbs);
    }
    else
    {
        outPrefix = response.substr(msgStart, outNextPos - msgStart);
    }

    return true;
}

[[nodiscard]] HRESULT ImapFetchMessageSummaries(const ConnectionInfo& conn,
                                                std::wstring_view mailboxPath,
                                                std::span<const uint64_t> uids,
                                                std::unordered_map<uint64_t, ImapMessageSummary>& inOut) noexcept
{
    if (uids.empty())
    {
        return S_OK;
    }

    std::vector<uint64_t> sorted;
    sorted.reserve(uids.size());
    for (const uint64_t uid : uids)
    {
        sorted.push_back(uid);
    }

    std::sort(sorted.begin(), sorted.end());
    sorted.erase(std::unique(sorted.begin(), sorted.end()), sorted.end());

    // Keep IMAP commands reasonably short for server compatibility. (RFC 3501: minimum 1000 octets line length support.)
    constexpr size_t kMaxUidSetChars = 800u;

    auto fetchAndParse = [&](size_t startIndex, size_t endIndex, std::string_view uidSetText) noexcept -> HRESULT
    {
        if (startIndex >= endIndex || endIndex > sorted.size() || uidSetText.empty())
        {
            return S_OK;
        }

        std::string requestText;
        requestText = std::format("UID FETCH {} (UID FLAGS INTERNALDATE RFC822.SIZE ENVELOPE)", uidSetText);

        std::string response;
        const auto fetchStarted = std::chrono::steady_clock::now();
        HRESULT hr = CurlPerformImapCustomRequest(conn, mailboxPath, requestText, response);
        Debug::Perf::EmitDurationUs(
            L"filesystem.imap.fetch_summaries_us", Debug::Perf::ElapsedUs(fetchStarted), static_cast<uint64_t>(endIndex - startIndex), static_cast<uint64_t>(uidSetText.size()), hr);
        if (FAILED(hr))
        {
            return hr;
        }
        Debug::Perf::EmitValue(L"filesystem.imap.summary_response_bytes", static_cast<uint64_t>(response.size()), hr);

        size_t fetchParseFailures    = 0;
        size_t missingUidCount       = 0;
        size_t envelopeParseFailures = 0;
        size_t fetchBlocksParsed     = 0;

        const auto parseStarted = std::chrono::steady_clock::now();
        size_t parsePos = 0;
        while (true)
        {
            const size_t msgStart = FindImapUntaggedFetchLine(response, parsePos);
            if (msgStart == std::string_view::npos)
            {
                break;
            }

            size_t nextPos = 0;
            std::string_view prefix;
            std::string_view headerBlock;
            std::string_view suffix;
            if (! TryConsumeImapUntaggedFetchResponse(std::string_view(response), msgStart, nextPos, prefix, headerBlock, suffix))
            {
                ++fetchParseFailures;

                const size_t lineEnd = response.find('\n', msgStart);
                parsePos             = lineEnd == std::string_view::npos ? response.size() : (lineEnd + 1u);
                continue;
            }

            if (nextPos <= msgStart)
            {
                ++fetchParseFailures;
                parsePos = msgStart + 1u;
                continue;
            }

            ImapMessageSummary summary{};

            uint64_t uid = 0;
            bool hasUid  = TryParseUintAfterKey(prefix, "UID ", uid);
            if (! hasUid && ! suffix.empty())
            {
                hasUid = TryParseUintAfterKey(suffix, "UID ", uid);
            }

            if (! hasUid)
            {
                ++missingUidCount;
                parsePos = nextPos;
                continue;
            }
            summary.uid = uid;

            bool hasSize = TryParseUintAfterKey(prefix, "RFC822.SIZE ", summary.sizeBytes);
            if (! hasSize && ! suffix.empty())
            {
                static_cast<void>(TryParseUintAfterKey(suffix, "RFC822.SIZE ", summary.sizeBytes));
            }

            // FLAGS (...)
            {
                auto parseFlags = [&summary](std::string_view text) noexcept
                {
                    const size_t flagsPos = FindAsciiNoCase(text, "FLAGS", 0);
                    if (flagsPos == std::string_view::npos)
                    {
                        return;
                    }

                    const size_t open = text.find('(', flagsPos);
                    if (open == std::string_view::npos)
                    {
                        return;
                    }

                    const size_t close = text.find(')', open);
                    if (close == std::string_view::npos || close <= open)
                    {
                        return;
                    }

                    const std::string_view flagsText = text.substr(open + 1u, close - open - 1u);
                    size_t fp                        = 0;
                    while (fp < flagsText.size())
                    {
                        while (fp < flagsText.size() && (flagsText[fp] == ' ' || flagsText[fp] == '\t'))
                        {
                            ++fp;
                        }
                        const size_t startTok = fp;
                        while (fp < flagsText.size() && flagsText[fp] != ' ' && flagsText[fp] != '\t')
                        {
                            ++fp;
                        }
                        const std::string_view tok = flagsText.substr(startTok, fp - startTok);
                        if (! tok.empty())
                        {
                            if (FindAsciiNoCase(tok, "\\seen", 0) == 0 && tok.size() == 5u)
                            {
                                summary.seen = true;
                            }
                            else if (FindAsciiNoCase(tok, "\\flagged", 0) == 0 && tok.size() == 8u)
                            {
                                summary.flagged = true;
                            }
                            else if (FindAsciiNoCase(tok, "\\deleted", 0) == 0 && tok.size() == 8u)
                            {
                                summary.deleted = true;
                            }
                        }
                    }
                };

                parseFlags(prefix);
                if (! suffix.empty())
                {
                    parseFlags(suffix);
                }
            }

            // INTERNALDATE "..."
            {
                auto parseInternalDate = [&summary](std::string_view text) noexcept
                {
                    const size_t idPos = FindAsciiNoCase(text, "INTERNALDATE", 0);
                    if (idPos == std::string_view::npos)
                    {
                        return;
                    }

                    size_t p           = idPos;
                    const size_t quote = text.find('"', p);
                    if (quote == std::string_view::npos)
                    {
                        return;
                    }

                    p = quote;
                    std::string internalDate;
                    if (TryParseImapQuotedString(text, p, internalDate))
                    {
                        static_cast<void>(TryParseImapInternalDateToFileTime(internalDate, summary.recvTime));
                    }
                };

                parseInternalDate(prefix);
                if (summary.recvTime == 0 && ! suffix.empty())
                {
                    parseInternalDate(suffix);
                }
            }

            ImapEnvelopeFields env;
            bool hasEnvelope = TryExtractEnvelopeFields(prefix, env);
            if (! hasEnvelope && ! suffix.empty())
            {
                hasEnvelope = TryExtractEnvelopeFields(suffix, env);
            }

            if (hasEnvelope)
            {
                summary.subject = DecodeRfc2047EncodedWordsToUtf16(env.subject);
                summary.from    = Utf16FromImapHeaderValue(env.fromAddrSpec);

                __int64 sentTime = 0;
                if (TryParseRfc5322DateToFileTime(env.date, sentTime))
                {
                    summary.sentTime = sentTime;
                }
            }
            else if (! headerBlock.empty())
            {
                ImapHeaderFields headers;
                if (TryExtractHeaderFields(headerBlock, headers))
                {
                    summary.subject = DecodeRfc2047EncodedWordsToUtf16(headers.subject);
                    summary.from    = ExtractEmailAddressFromFromHeader(headers.from);

                    __int64 sentTime = 0;
                    if (TryParseRfc5322DateToFileTime(headers.date, sentTime))
                    {
                        summary.sentTime = sentTime;
                    }
                }
            }
            else
            {
                ++envelopeParseFailures;
            }

            if (summary.sentTime == 0)
            {
                summary.sentTime = summary.recvTime;
            }

            inOut.insert_or_assign(summary.uid, std::move(summary));

            ++fetchBlocksParsed;

            parsePos = nextPos;
        }
        Debug::Perf::EmitDurationUs(
            L"filesystem.imap.summary_parse_us",
            Debug::Perf::ElapsedUs(parseStarted),
            static_cast<uint64_t>(fetchBlocksParsed),
            static_cast<uint64_t>(fetchParseFailures + missingUidCount + envelopeParseFailures),
            hr);

        size_t missingRequested = 0;
        std::array<uint64_t, 5> missingSamples{};
        size_t missingSampleCount = 0;
        for (size_t i = startIndex; i < endIndex; ++i)
        {
            const uint64_t uid = sorted[i];
            if (inOut.find(uid) == inOut.end())
            {
                ++missingRequested;
                if (missingSampleCount < missingSamples.size())
                {
                    missingSamples[missingSampleCount++] = uid;
                }
            }
        }

        if (fetchParseFailures > 0 || missingUidCount > 0 || envelopeParseFailures > 0 || missingRequested > 0)
        {
            std::wstring missingText;
            if (missingSampleCount > 0)
            {
                for (size_t i = 0; i < missingSampleCount; ++i)
                {
                    if (i != 0)
                    {
                        missingText.append(L",");
                    }
                    missingText.append(std::to_wstring(missingSamples[i]));
                }
            }

            constexpr size_t kMaxUidSetLog = 160;
            std::string uidSetShort(uidSetText);
            if (uidSetShort.size() > kMaxUidSetLog)
            {
                uidSetShort.resize(kMaxUidSetLog);
                uidSetShort.append("...");
            }

            constexpr size_t kMaxRequestLog = 200;
            std::string requestShort        = requestText;
            if (requestShort.size() > kMaxRequestLog)
            {
                requestShort.resize(kMaxRequestLog);
                requestShort.append("...");
            }

            std::wstring responseFirstLine;
            std::wstring responseFetchLines;

            // First non-empty line (trimmed)
            {
                std::string firstLine;
                size_t start = 0;
                while (start < response.size())
                {
                    size_t end = response.find('\n', start);
                    if (end == std::string_view::npos)
                    {
                        end = response.size();
                    }

                    std::string_view line = std::string_view(response).substr(start, end - start);
                    if (! line.empty() && line.back() == '\r')
                    {
                        line.remove_suffix(1);
                    }

                    firstLine = TrimAscii(line);
                    if (! firstLine.empty())
                    {
                        break;
                    }
                    start = end + 1u;
                }

                constexpr size_t kMaxLineLog = 200;
                if (firstLine.size() > kMaxLineLog)
                {
                    firstLine.resize(kMaxLineLog);
                    firstLine.append("...");
                }

                responseFirstLine = Utf16FromImapHeaderValue(firstLine);
            }

            // First few FETCH lines (no literal payload)
            {
                constexpr size_t kMaxFetchLines = 4;
                size_t scanPos                  = 0;
                for (size_t i = 0; i < kMaxFetchLines; ++i)
                {
                    const size_t fetchStart = FindImapUntaggedFetchLine(std::string_view(response), scanPos);
                    if (fetchStart == std::string_view::npos)
                    {
                        break;
                    }

                    size_t fetchLineEnd = response.find('\n', fetchStart);
                    if (fetchLineEnd == std::string_view::npos)
                    {
                        fetchLineEnd = response.size();
                    }

                    std::string_view line = std::string_view(response).substr(fetchStart, fetchLineEnd - fetchStart);
                    if (! line.empty() && line.back() == '\r')
                    {
                        line.remove_suffix(1);
                    }
                    line = TrimAsciiView(line);

                    constexpr size_t kMaxFetchLineLog = 220;
                    std::string lineShort(line);
                    if (lineShort.size() > kMaxFetchLineLog)
                    {
                        lineShort.resize(kMaxFetchLineLog);
                        lineShort.append("...");
                    }

                    if (! responseFetchLines.empty())
                    {
                        responseFetchLines.append(L" | ");
                    }
                    responseFetchLines.append(Utf16FromImapHeaderValue(lineShort));

                    scanPos = fetchLineEnd + 1u;
                }
            }

            Debug::Warning(L"imap summary request mailbox='{}' req='{}'", mailboxPath, Utf16FromImapHeaderValue(requestShort));
            Debug::Warning(L"imap summary response mailbox='{}' firstLine='{}' fetchLines='{}'",
                           mailboxPath,
                           responseFirstLine.empty() ? L"(none)" : responseFirstLine,
                           responseFetchLines.empty() ? L"(none)" : responseFetchLines);

            Debug::Warning(L"imap summary parse anomalies mailbox='{}' fetchBlocks={} fetchParseFailures={} envelopeParseFailures={} missingUidInFetch={} "
                           L"missingRequested={} missingSample='{}' requested={} responseBytes={} uidSet='{}'",
                           mailboxPath,
                           fetchBlocksParsed,
                           fetchParseFailures,
                           envelopeParseFailures,
                           missingUidCount,
                           missingRequested,
                           missingText.empty() ? L"(none)" : missingText,
                           endIndex - startIndex,
                           response.size(),
                           Utf16FromUtf8(uidSetShort));
        }

        return S_OK;
    };

    std::string uidSet;
    uidSet.reserve(std::min(kMaxUidSetChars, sorted.size() * 12u));

    size_t groupStart = 0;
    for (size_t i = 0; i < sorted.size(); ++i)
    {
        const std::string part = std::to_string(sorted[i]);
        const size_t needed    = part.size() + (uidSet.empty() ? 0u : 1u);

        if (! uidSet.empty() && uidSet.size() + needed > kMaxUidSetChars)
        {
            const HRESULT hr = fetchAndParse(groupStart, i, uidSet);
            if (FAILED(hr))
            {
                return hr;
            }

            uidSet.clear();
            groupStart = i;
        }

        if (! uidSet.empty())
        {
            uidSet.push_back(',');
        }
        uidSet.append(part);
    }

    if (! uidSet.empty())
    {
        const HRESULT hr = fetchAndParse(groupStart, sorted.size(), uidSet);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

struct ImapFetchToFileContext
{
    HANDLE file     = INVALID_HANDLE_VALUE;
    HRESULT abortHr = S_OK;
    std::string buffer;
    uint64_t remainingBytes = 0;
    bool done               = false;
};

[[nodiscard]] bool TryExtractImapLiteralSize(std::string_view data, size_t& literalStart, uint64_t& literalSize) noexcept
{
    literalStart = 0;
    literalSize  = 0;

    const size_t brace = data.find('{');
    if (brace == std::string_view::npos)
    {
        return false;
    }

    size_t pos = brace + 1u;
    if (pos >= data.size())
    {
        return false;
    }

    uint64_t value = 0;
    size_t digits  = 0;
    while (pos < data.size() && data[pos] >= '0' && data[pos] <= '9')
    {
        const uint64_t digit = static_cast<uint64_t>(data[pos] - '0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
        ++digits;
        ++pos;
    }

    if (digits == 0 || pos >= data.size())
    {
        return false;
    }

    if (data[pos] == '+')
    {
        ++pos;
    }

    if (pos >= data.size() || data[pos] != '}')
    {
        return false;
    }

    size_t start = 0;
    if ((pos + 1u) < data.size() && data[pos + 1u] == '\n')
    {
        start = pos + 2u;
    }
    else if ((pos + 2u) < data.size() && data[pos + 1u] == '\r' && data[pos + 2u] == '\n')
    {
        start = pos + 3u;
    }
    else
    {
        return false;
    }

    literalStart = start;
    literalSize  = value;
    return true;
}

size_t CurlWriteImapFetchToFile(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
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

    auto* ctx = static_cast<ImapFetchToFileContext*>(userdata);
    if (! ctx || ! ctx->file || ctx->file == INVALID_HANDLE_VALUE)
    {
        return 0;
    }

    const char* data = static_cast<const char*>(ptr);

    size_t offset = 0;
    while (offset < total)
    {
        if (ctx->done)
        {
            return total;
        }

        if (ctx->remainingBytes > 0)
        {
            const size_t chunk = std::min<size_t>(total - offset, static_cast<size_t>(std::min<uint64_t>(ctx->remainingBytes, SIZE_MAX)));

            size_t writtenTotal = 0;
            while (writtenTotal < chunk)
            {
                const size_t part = chunk - writtenTotal;
                const DWORD take =
                    part > static_cast<size_t>((std::numeric_limits<DWORD>::max)()) ? (std::numeric_limits<DWORD>::max)() : static_cast<DWORD>(part);

                DWORD written = 0;
                if (! WriteFile(ctx->file, data + offset + writtenTotal, take, &written, nullptr))
                {
                    ctx->abortHr = HRESULT_FROM_WIN32(GetLastError());
                    return 0;
                }
                if (written == 0)
                {
                    ctx->abortHr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                    return 0;
                }
                writtenTotal += written;
                ctx->remainingBytes -= static_cast<uint64_t>(written);
            }

            offset += writtenTotal;

            if (ctx->remainingBytes == 0)
            {
                ctx->done = true;
            }

            continue;
        }

        ctx->buffer.append(data + offset, total - offset);
        offset = total;

        if (ctx->buffer.size() > 256u * 1024u)
        {
            ctx->abortHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return 0;
        }

        size_t literalStart  = 0;
        uint64_t literalSize = 0;
        if (! TryExtractImapLiteralSize(ctx->buffer, literalStart, literalSize))
        {
            continue;
        }

        const size_t available = ctx->buffer.size() > literalStart ? ctx->buffer.size() - literalStart : 0u;
        const size_t take      = std::min<size_t>(available, static_cast<size_t>(std::min<uint64_t>(literalSize, SIZE_MAX)));

        ctx->remainingBytes = literalSize;
        if (take > 0)
        {
            size_t writtenTotal = 0;
            while (writtenTotal < take)
            {
                const size_t part = take - writtenTotal;
                const DWORD writeTake =
                    part > static_cast<size_t>((std::numeric_limits<DWORD>::max)()) ? (std::numeric_limits<DWORD>::max)() : static_cast<DWORD>(part);

                DWORD written = 0;
                if (! WriteFile(ctx->file, ctx->buffer.data() + literalStart + writtenTotal, writeTake, &written, nullptr))
                {
                    ctx->abortHr = HRESULT_FROM_WIN32(GetLastError());
                    return 0;
                }
                if (written == 0)
                {
                    ctx->abortHr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                    return 0;
                }

                writtenTotal += written;
                ctx->remainingBytes -= static_cast<uint64_t>(written);
            }
        }

        ctx->buffer.clear();

        if (ctx->remainingBytes == 0)
        {
            ctx->done = true;
        }
    }

    return total;
}

[[nodiscard]] HRESULT ImapFetchMessageToFile(const ConnectionInfo& conn, std::wstring_view mailboxPath, uint64_t uid, HANDLE file) noexcept
{
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
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

    const std::string url = BuildImapUrl(conn, mailboxPath);
    if (url.empty())
    {
        return E_INVALIDARG;
    }

    std::string requestText;
    requestText = std::format("UID FETCH {} BODY.PEEK[]", uid);

    ImapFetchToFileContext ctx{};
    ctx.file = file;

    curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_CUSTOMREQUEST, requestText.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteImapFetchToFile);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &ctx);
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    ApplyCommonCurlOptions(curl.get(), conn, nullptr, false);
    if (ImapSchemeForConnection(conn) == "imap")
    {
        curl_easy_setopt(curl.get(), CURLOPT_USE_SSL, CURLUSESSL_TRY);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code == CURLE_WRITE_ERROR && FAILED(ctx.abortHr))
    {
        return ctx.abortHr;
    }

    hr = HResultFromCurl(code);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! ctx.done)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    return S_OK;
}

[[nodiscard]] std::string ImapQuoteString(std::string_view text) noexcept
{
    std::string out;
    out.reserve(text.size() + 2u);
    out.push_back('"');
    for (const char ch : text)
    {
        if (ch == '"' || ch == '\\')
        {
            out.push_back('\\');
        }
        out.push_back(ch);
    }
    out.push_back('"');
    return out;
}

[[nodiscard]] HRESULT ImapFetchMailboxStatus(const ConnectionInfo& conn,
                                             std::wstring_view mailboxPath,
                                             wchar_t delimiter,
                                             ImapMailboxStatus& outStatus) noexcept
{
    outStatus = {};

    const std::wstring normalized   = NormalizePluginPath(mailboxPath);
    const std::wstring_view trimmed = TrimTrailingSlash(normalized);
    if (trimmed.empty() || trimmed == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::wstring_view mailboxName = trimmed;
    if (! mailboxName.empty() && mailboxName.front() == L'/')
    {
        mailboxName.remove_prefix(1);
    }

    const std::wstring serverName = ImapMailboxNameToServerMailboxName(mailboxName, delimiter);
    if (serverName.empty())
    {
        return E_OUTOFMEMORY;
    }

    const std::string serverNameUtf8 = Utf8FromUtf16(serverName);
    if (serverNameUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    const std::string requestText = std::format("STATUS {} (MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)", ImapQuoteString(serverNameUtf8));

    std::string response;
    const auto started = std::chrono::steady_clock::now();
    HRESULT hr         = CurlPerformImapCustomRequest(conn, L"/", requestText, response);
    Debug::Perf::EmitDurationUs(L"filesystem.imap.status_mailbox_us", Debug::Perf::ElapsedUs(started), static_cast<uint64_t>(response.size()), 0u, hr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! TryParseImapMailboxStatus(response, outStatus))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    return S_OK;
}

[[nodiscard]] HRESULT ImapDownloadMessageToFile(const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file) noexcept
{
    const std::wstring fullPath = JoinPluginPathWide(conn.basePathWide, pluginPath);

    const std::wstring_view leaf = LeafName(fullPath);
    uint64_t uid                 = 0;
    if (! TryParseImapUidFromLeafName(leaf, uid))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::wstring mailboxPath = ParentPath(fullPath);
    mailboxPath              = std::wstring(TrimTrailingSlash(mailboxPath));
    if (mailboxPath.empty())
    {
        mailboxPath = L"/";
    }

    if (mailboxPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    wchar_t delimiter = L'\0';
    HRESULT hr        = ImapGetHierarchyDelimiter(conn, delimiter);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverMailboxPath = ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
    if (serverMailboxPath.empty())
    {
        return E_OUTOFMEMORY;
    }

    return ImapFetchMessageToFile(conn, serverMailboxPath, uid, file);
}

[[nodiscard]] HRESULT ImapDeleteMessage(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept
{
    const std::wstring fullPath = JoinPluginPathWide(conn.basePathWide, pluginPath);

    const std::wstring_view leaf = LeafName(fullPath);
    uint64_t uid                 = 0;
    if (! TryParseImapUidFromLeafName(leaf, uid))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::wstring mailboxPath = ParentPath(fullPath);
    mailboxPath              = std::wstring(TrimTrailingSlash(mailboxPath));
    if (mailboxPath.empty())
    {
        mailboxPath = L"/";
    }

    if (mailboxPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    wchar_t delimiter = L'\0';
    HRESULT hr        = ImapGetHierarchyDelimiter(conn, delimiter);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverMailboxPath = ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
    if (serverMailboxPath.empty())
    {
        return E_OUTOFMEMORY;
    }

    std::string sink;
    hr = CurlPerformImapCustomRequest(conn, serverMailboxPath, std::format("UID STORE {} +FLAGS.SILENT (\\Deleted)", uid), sink);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CurlPerformImapCustomRequest(conn, serverMailboxPath, std::format("UID EXPUNGE {}", uid), sink);
    if (SUCCEEDED(hr))
    {
        return S_OK;
    }

    return CurlPerformImapCustomRequest(conn, serverMailboxPath, "EXPUNGE", sink);
}

[[nodiscard]] HRESULT ImapDeleteMailbox(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept
{
    const std::wstring fullPath     = JoinPluginPathWide(conn.basePathWide, pluginPath);
    const std::wstring normalized   = NormalizePluginPath(fullPath);
    const std::wstring_view trimmed = TrimTrailingSlash(normalized);

    if (trimmed.empty() || trimmed == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::wstring_view name = trimmed;
    if (! name.empty() && name.front() == L'/')
    {
        name.remove_prefix(1);
    }

    wchar_t delimiter = L'\0';
    HRESULT hr        = ImapGetHierarchyDelimiter(conn, delimiter);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverName = ImapMailboxNameToServerMailboxName(name, delimiter);
    if (serverName.empty())
    {
        return E_OUTOFMEMORY;
    }

    const std::string nameUtf8 = Utf8FromUtf16(serverName);
    if (nameUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::string sink;
    return CurlPerformImapCustomRequest(conn, L"/", std::format("DELETE {}", ImapQuoteString(nameUtf8)), sink);
}

[[nodiscard]] HRESULT ImapCreateMailbox(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept
{
    const std::wstring fullPath     = JoinPluginPathWide(conn.basePathWide, pluginPath);
    const std::wstring normalized   = NormalizePluginPath(fullPath);
    const std::wstring_view trimmed = TrimTrailingSlash(normalized);

    if (trimmed.empty() || trimmed == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::wstring_view name = trimmed;
    if (! name.empty() && name.front() == L'/')
    {
        name.remove_prefix(1);
    }

    wchar_t delimiter = L'\0';
    HRESULT hr        = ImapGetHierarchyDelimiter(conn, delimiter);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverName = ImapMailboxNameToServerMailboxName(name, delimiter);
    if (serverName.empty())
    {
        return E_OUTOFMEMORY;
    }

    const std::string nameUtf8 = Utf8FromUtf16(serverName);
    if (nameUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::string sink;
    return CurlPerformImapCustomRequest(conn, L"/", std::format("CREATE {}", ImapQuoteString(nameUtf8)), sink);
}

[[nodiscard]] void SplitSlashPath(std::wstring_view text, std::vector<std::wstring_view>& segments) noexcept
{
    segments.clear();

    size_t start = 0;
    while (start < text.size())
    {
        size_t end = text.find(L'/', start);
        if (end == std::wstring_view::npos)
        {
            end = text.size();
        }

        const std::wstring_view part = text.substr(start, end - start);
        if (! part.empty())
        {
            segments.push_back(part);
        }

        if (end >= text.size())
        {
            break;
        }
        start = end + 1u;
    }
}

[[nodiscard]] bool StartsWithSegments(const std::vector<std::wstring_view>& segments, const std::vector<std::wstring_view>& prefix) noexcept
{
    if (prefix.size() > segments.size())
    {
        return false;
    }

    for (size_t i = 0; i < prefix.size(); ++i)
    {
        if (segments[i] != prefix[i])
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] HRESULT ImapReadDirectoryEntries(const ConnectionInfo& conn,
                                               std::wstring_view pluginPath,
                                               std::vector<FilesInformationCurl::Entry>& entries) noexcept
{
    Debug::Perf::Scope perf(L"filesystem.imap.read_directory_us");
    entries.clear();

    std::vector<ImapMailboxEntry> mailboxes;
    wchar_t delimiter = L'\0';
    auto listStarted  = std::chrono::steady_clock::now();
    HRESULT hr        = ImapListMailboxes(conn, mailboxes, &delimiter);
    Debug::Perf::EmitDurationUs(L"filesystem.imap.list_mailboxes_us", Debug::Perf::ElapsedUs(listStarted), static_cast<uint64_t>(mailboxes.size()), 0u, hr);
    if (FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }

    const std::wstring fullPath     = JoinPluginPathWide(conn.basePathWide, pluginPath);
    const std::wstring normalized   = NormalizePluginPath(fullPath);
    const std::wstring_view trimmed = TrimTrailingSlash(normalized);

    std::wstring mailboxName;
    if (! trimmed.empty() && trimmed != L"/")
    {
        std::wstring_view nameView = trimmed;
        if (nameView.front() == L'/')
        {
            nameView.remove_prefix(1);
        }
        mailboxName.assign(nameView);
    }

    if (! mailboxName.empty())
    {
        mailboxName = ImapMailboxNameToPluginMailboxName(mailboxName, delimiter);
        if (mailboxName.empty())
        {
            return E_OUTOFMEMORY;
        }
    }

    std::vector<std::wstring_view> prefixSegs;
    SplitSlashPath(mailboxName, prefixSegs);

    std::unordered_set<std::wstring> childDirs;
    std::vector<std::wstring_view> mboxSegs;
    for (const auto& mbox : mailboxes)
    {
        SplitSlashPath(mbox.name, mboxSegs);
        if (! StartsWithSegments(mboxSegs, prefixSegs))
        {
            continue;
        }

        if (mboxSegs.size() <= prefixSegs.size())
        {
            continue;
        }

        const std::wstring_view child = mboxSegs[prefixSegs.size()];
        if (child.empty())
        {
            continue;
        }

        childDirs.insert(std::wstring(child));
    }

    for (const auto& child : childDirs)
    {
        FilesInformationCurl::Entry entry{};
        entry.attributes = FILE_ATTRIBUTE_DIRECTORY;
        entry.name       = child;
        entries.push_back(std::move(entry));
    }

    if (mailboxName.empty())
    {
        return S_OK;
    }

    bool selectableMailbox = false;
    for (const auto& mbox : mailboxes)
    {
        if (mbox.name == mailboxName)
        {
            selectableMailbox = ! mbox.noSelect;
            break;
        }
    }

    if (! selectableMailbox)
    {
        perf.SetValue0(entries.size());
        return S_OK;
    }

    std::vector<uint64_t> uids;
    auto listUidsStarted = std::chrono::steady_clock::now();
    hr = ImapListMessageUids(conn, mailboxName, delimiter, uids);
    Debug::Perf::EmitDurationUs(L"filesystem.imap.list_uids_us", Debug::Perf::ElapsedUs(listUidsStarted), static_cast<uint64_t>(uids.size()), 0u, hr);
    if (FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }

    std::sort(uids.begin(), uids.end(), std::greater<>());
    if (uids.empty())
    {
        perf.SetValue0(entries.size());
        return S_OK;
    }

    std::wstring serverMailboxPath;
    {
        const std::wstring serverName = ImapMailboxNameToServerMailboxName(mailboxName, delimiter);
        if (serverName.empty())
        {
            return E_OUTOFMEMORY;
        }

        serverMailboxPath.reserve(serverName.size() + 1u);
        serverMailboxPath.push_back(L'/');
        serverMailboxPath.append(serverName);
    }

    std::unordered_map<uint64_t, ImapMessageSummary> summaries;
    summaries.reserve(uids.size());

    constexpr size_t kFetchChunkSize = 200u;
    HRESULT metaHr                   = S_OK;
    size_t repairFetchCount          = 0;
    const size_t repairFetchBudget    = ResolveImapSummaryRepairFetchBudget(uids.size());
    size_t missingSummaryCount        = 0;
    bool repairBudgetWarningLogged    = false;

    auto collectMissingSummaries = [&summaries](std::span<const uint64_t> requested) -> std::vector<uint64_t>
    {
        std::vector<uint64_t> missing;
        missing.reserve(requested.size());
        for (const uint64_t uid : requested)
        {
            if (summaries.find(uid) == summaries.end())
            {
                missing.push_back(uid);
            }
        }
        return missing;
    };

    auto repairMissingSummaries = [&](std::span<const uint64_t> missingUids, HRESULT reasonHr) -> HRESULT
    {
        if (missingUids.empty())
        {
            return S_OK;
        }

        auto tryConsumeRepairFetch = [&](std::wstring_view phase, size_t missingCount) -> bool
        {
            if (repairFetchCount < repairFetchBudget)
            {
                ++repairFetchCount;
                return true;
            }

            if (! repairBudgetWarningLogged)
            {
                repairBudgetWarningLogged = true;
                Debug::Warning(L"imap message summary repair budget exhausted: reason={:#x} mailbox='{}' server='{}' budget={} missing={} phase='{}'",
                               reasonHr,
                               mailboxName,
                               Utf16FromUtf8(conn.host),
                               repairFetchBudget,
                               missingCount,
                               phase);
            }
            return false;
        };

        constexpr size_t kRepairBatchSize = 16u;
        HRESULT repairResult              = S_OK;
        const std::vector<ImapUidBatchRange> ranges = BuildImapUidBatchRanges(missingUids.size(), kRepairBatchSize);
        for (const ImapUidBatchRange& range : ranges)
        {
            const std::span<const uint64_t> batch(missingUids.data() + range.offset, range.count);
            bool repairBudgetExhaustedInPass = false;

            if (! tryConsumeRepairFetch(L"batch", missingUids.size()))
            {
                break;
            }
            const HRESULT batchHr = ImapFetchMessageSummaries(conn, serverMailboxPath, batch, summaries);
            if (FAILED(batchHr))
            {
                if (SUCCEEDED(repairResult))
                {
                    repairResult = batchHr;
                }
                Debug::Warning(L"imap message summary repair fetch failed: hr={:#x} reason={:#x} mailbox='{}' server='{}' missing={}",
                               batchHr,
                               reasonHr,
                               mailboxName,
                               Utf16FromUtf8(conn.host),
                               batch.size());
            }

            for (const uint64_t uid : batch)
            {
                if (summaries.find(uid) != summaries.end())
                {
                    continue;
                }

                const uint64_t singleUid = uid;
                const std::span<const uint64_t> one(&singleUid, 1);
                if (! tryConsumeRepairFetch(L"single", missingUids.size()))
                {
                    repairBudgetExhaustedInPass = true;
                    break;
                }
                const HRESULT singleHr = ImapFetchMessageSummaries(conn, serverMailboxPath, one, summaries);
                if (FAILED(singleHr))
                {
                    if (SUCCEEDED(repairResult))
                    {
                        repairResult = singleHr;
                    }
                    Debug::Warning(L"imap message summary single-uid repair fetch failed: hr={:#x} reason={:#x} mailbox='{}' server='{}' uid={}",
                                   singleHr,
                                   reasonHr,
                                   mailboxName,
                                   Utf16FromUtf8(conn.host),
                                   uid);
                }
            }

            if (repairBudgetExhaustedInPass)
            {
                break;
            }
        }

        return repairResult;
    };

    for (size_t start = 0; start < uids.size(); start += kFetchChunkSize)
    {
        const size_t count = std::min(kFetchChunkSize, uids.size() - start);
        const std::span<const uint64_t> chunk(uids.data() + start, count);
        const HRESULT chunkHr = ImapFetchMessageSummaries(conn, serverMailboxPath, chunk, summaries);
        if (FAILED(chunkHr))
        {
            bool chunkFailureRecorded = false;
            if (SUCCEEDED(metaHr))
            {
                metaHr               = chunkHr;
                chunkFailureRecorded = true;
            }
            Debug::Warning(L"imap message summary bulk fetch failed: hr={:#x} mailbox='{}' server='{}' requested={}",
                           chunkHr,
                           mailboxName,
                           Utf16FromUtf8(conn.host),
                           chunk.size());

            missingSummaryCount += chunk.size();
            const HRESULT repairHr = repairMissingSummaries(chunk, chunkHr);
            if (SUCCEEDED(repairHr) && chunkFailureRecorded)
            {
                metaHr = S_OK;
            }
            continue;
        }

        // Some servers are picky about multi-UID FETCH sets. Repair any missed summaries in smaller batches,
        // then single UID requests, so pane metadata can match the targeted Properties path.
        const std::vector<uint64_t> missing = collectMissingSummaries(chunk);
        if (! missing.empty())
        {
            missingSummaryCount += missing.size();
            const HRESULT repairHr = repairMissingSummaries(missing, chunkHr);
            if (FAILED(repairHr) && SUCCEEDED(metaHr))
            {
                metaHr = repairHr;
            }
        }
    }

    if (FAILED(metaHr))
    {
        Debug::Warning(L"imap message summary fetch failed: hr={:#x} mailbox='{}' server='{}'", metaHr, mailboxName, Utf16FromUtf8(conn.host));
    }
    Debug::Perf::EmitValue(L"filesystem.imap.summary_missing_count", static_cast<uint64_t>(missingSummaryCount), metaHr);
    Debug::Perf::EmitValue(L"filesystem.imap.summary_repair_count", static_cast<uint64_t>(repairFetchCount), metaHr);

    auto buildStarted = std::chrono::steady_clock::now();
    for (const uint64_t uid : uids)
    {
        FilesInformationCurl::Entry entry{};
        entry.attributes = FILE_ATTRIBUTE_NORMAL;
        entry.fileIndex  = (uid <= static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)())) ? static_cast<unsigned long>(uid) : 0u;

        const auto found = summaries.find(uid);
        if (found != summaries.end())
        {
            const ImapMessageSummary& meta = found->second;

            entry.sizeBytes     = meta.sizeBytes;
            entry.creationTime  = meta.sentTime;
            entry.changeTime    = meta.recvTime;
            entry.lastWriteTime = meta.recvTime;

            if (meta.flagged)
            {
                entry.attributes |= kImapFileAttributeMarked;
            }
            if (! meta.seen)
            {
                entry.attributes |= kImapFileAttributeUnread;
            }
            if (meta.deleted)
            {
                entry.attributes |= kImapFileAttributeDeleted;
            }

            entry.name = BuildImapMessageLeafName(meta.subject, meta.from, uid);
        }

        if (entry.name.empty())
        {
            entry.name = std::format(L"{}.eml", uid);
        }
        entries.push_back(std::move(entry));
    }
    Debug::Perf::EmitDurationUs(
        L"filesystem.imap.build_fileinfo_us", Debug::Perf::ElapsedUs(buildStarted), static_cast<uint64_t>(entries.size()), static_cast<uint64_t>(summaries.size()), metaHr);
    perf.SetValue0(static_cast<uint64_t>(entries.size()));
    perf.SetValue1(static_cast<uint64_t>(uids.size()));
    perf.SetHr(metaHr);

    return S_OK;
}

[[nodiscard]] void FillImapMessageEntryFromSummary(FilesInformationCurl::Entry& entry, const ImapMessageSummary& meta, uint64_t uid) noexcept
{
    entry.attributes    = FILE_ATTRIBUTE_NORMAL;
    entry.fileIndex     = (uid <= static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)())) ? static_cast<unsigned long>(uid) : 0u;
    entry.sizeBytes     = meta.sizeBytes;
    entry.creationTime  = meta.sentTime;
    entry.changeTime    = meta.recvTime;
    entry.lastWriteTime = meta.recvTime;

    if (meta.flagged)
    {
        entry.attributes |= kImapFileAttributeMarked;
    }
    if (! meta.seen)
    {
        entry.attributes |= kImapFileAttributeUnread;
    }
    if (meta.deleted)
    {
        entry.attributes |= kImapFileAttributeDeleted;
    }

    entry.name = BuildImapMessageLeafName(meta.subject, meta.from, uid);
    if (entry.name.empty())
    {
        entry.name = std::format(L"{}.eml", uid);
    }
}

[[nodiscard]] HRESULT ImapGetEntryInfo(const ConnectionInfo& conn, std::wstring_view path, FilesInformationCurl::Entry& out) noexcept
{
    out = {};

    const std::wstring fullPath     = JoinPluginPathWide(conn.basePathWide, path);
    const std::wstring normalized   = NormalizePluginPath(fullPath);
    const std::wstring_view trimmed = TrimTrailingSlash(normalized);
    if (trimmed.empty() || trimmed == L"/")
    {
        out.attributes = FILE_ATTRIBUTE_DIRECTORY;
        out.name       = L"/";
        return S_OK;
    }

    const std::wstring_view leaf = LeafName(trimmed);
    uint64_t uid                 = 0;
    if (TryParseImapUidFromLeafName(leaf, uid))
    {
        std::wstring mailboxPath = ParentPath(trimmed);
        mailboxPath              = std::wstring(TrimTrailingSlash(mailboxPath));
        if (mailboxPath.empty() || mailboxPath == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        wchar_t delimiter = L'\0';
        HRESULT hr        = ImapGetHierarchyDelimiter(conn, delimiter);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring serverMailboxPath = ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
        if (serverMailboxPath.empty())
        {
            return E_OUTOFMEMORY;
        }

        std::unordered_map<uint64_t, ImapMessageSummary> summaries;
        const uint64_t uidArr[1]{uid};
        hr = ImapFetchMessageSummaries(conn, serverMailboxPath, std::span<const uint64_t>(uidArr, 1), summaries);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto it = summaries.find(uid);
        if (it == summaries.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        FillImapMessageEntryFromSummary(out, it->second, uid);
        return S_OK;
    }

    std::vector<ImapMailboxEntry> mailboxes;
    wchar_t delimiter = L'\0';
    HRESULT hr        = ImapListMailboxes(conn, mailboxes, &delimiter);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring_view mailboxName = trimmed;
    if (! mailboxName.empty() && mailboxName.front() == L'/')
    {
        mailboxName.remove_prefix(1);
    }

    std::vector<std::wstring_view> targetSegs;
    SplitSlashPath(mailboxName, targetSegs);

    std::vector<std::wstring_view> mailboxSegs;
    for (const auto& mailbox : mailboxes)
    {
        if (mailbox.name == mailboxName)
        {
            out.attributes = FILE_ATTRIBUTE_DIRECTORY;
            out.name       = std::wstring(LeafName(mailboxName));
            return S_OK;
        }

        SplitSlashPath(mailbox.name, mailboxSegs);
        if (StartsWithSegments(mailboxSegs, targetSegs) && mailboxSegs.size() > targetSegs.size())
        {
            out.attributes = FILE_ATTRIBUTE_DIRECTORY;
            out.name       = std::wstring(LeafName(mailboxName));
            return S_OK;
        }
    }

    return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
}

[[nodiscard]] HRESULT ReadDirectoryEntries(const ConnectionInfo& conn, std::wstring_view path, std::vector<FilesInformationCurl::Entry>& entries) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapReadDirectoryEntries(conn, path, entries);
    }

    return CurlPerformListAndParse(conn, path, entries);
}

[[nodiscard]] HRESULT GetEntryInfo(const ConnectionInfo& conn, std::wstring_view path, FilesInformationCurl::Entry& out) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapGetEntryInfo(conn, path, out);
    }

    const std::wstring normalized = NormalizePluginPath(path);
    if (normalized == L"/")
    {
        out            = {};
        out.attributes = FILE_ATTRIBUTE_DIRECTORY;
        out.name       = L"/";
        return S_OK;
    }

    const std::wstring parent    = ParentPath(normalized);
    const std::wstring_view leaf = LeafName(normalized);

    std::vector<FilesInformationCurl::Entry> entries;
    const HRESULT hr = ReadDirectoryEntries(conn, parent, entries);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto found = FindEntryByName(entries, leaf);
    if (! found.has_value())
    {
        if (conn.protocol == Protocol::Imap)
        {
            uint64_t uid = 0;
            if (TryParseImapUidFromLeafName(leaf, uid))
            {
                out            = {};
                out.attributes = FILE_ATTRIBUTE_NORMAL;
                out.fileIndex  = (uid <= static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)())) ? static_cast<unsigned long>(uid) : 0u;
                out.name       = std::wstring(leaf);
                return S_OK;
            }
        }

        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    out = found.value();
    return S_OK;
}

[[nodiscard]] HRESULT RemoteMkdir(const ConnectionInfo& conn, std::wstring_view path) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapCreateMailbox(conn, path);
    }

    const std::string remote = RemotePathForCommand(conn, path);
    if (remote.empty())
    {
        return E_INVALIDARG;
    }

    if (conn.protocol == Protocol::Ftp)
    {
        return CurlPerformQuote(conn, {std::format("MKD {}", remote)});
    }

    return CurlPerformQuote(conn, {std::format("mkdir {}", remote)});
}

[[nodiscard]] HRESULT RemoteDeleteFile(const ConnectionInfo& conn, std::wstring_view path) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapDeleteMessage(conn, path);
    }

    const std::string remote = RemotePathForCommand(conn, path);
    if (remote.empty())
    {
        return E_INVALIDARG;
    }

    if (conn.protocol == Protocol::Ftp)
    {
        return CurlPerformQuote(conn, {std::format("DELE {}", remote)});
    }

    return CurlPerformQuote(conn, {std::format("rm {}", remote)});
}

[[nodiscard]] HRESULT RemoteRemoveDirectory(const ConnectionInfo& conn, std::wstring_view path) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapDeleteMailbox(conn, path);
    }

    const std::string remote = RemotePathForCommand(conn, path);
    if (remote.empty())
    {
        return E_INVALIDARG;
    }

    if (conn.protocol == Protocol::Ftp)
    {
        return CurlPerformQuote(conn, {std::format("RMD {}", remote)});
    }

    return CurlPerformQuote(conn, {std::format("rmdir {}", remote)});
}

[[nodiscard]] HRESULT RemoteRename(const ConnectionInfo& conn, std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::string fromRemote = RemotePathForCommand(conn, sourcePath);
    const std::string toRemote   = RemotePathForCommand(conn, destinationPath);
    if (fromRemote.empty() || toRemote.empty())
    {
        return E_INVALIDARG;
    }

    if (conn.protocol == Protocol::Ftp)
    {
        return CurlPerformQuote(conn, {std::format("RNFR {}", fromRemote), std::format("RNTO {}", toRemote)});
    }

    return CurlPerformQuote(conn, {std::format("rename {} {}", fromRemote, toRemote)});
}

[[nodiscard]] HRESULT EnsureDirectoryExists(const ConnectionInfo& conn, std::wstring_view directoryPath) noexcept
{
    const HRESULT hr = RemoteMkdir(conn, directoryPath);
    if (SUCCEEDED(hr))
    {
        return S_OK;
    }

    FilesInformationCurl::Entry existing{};
    const HRESULT existsHr = GetEntryInfo(conn, directoryPath, existing);
    if (SUCCEEDED(existsHr) && (existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return S_OK;
    }

    return hr;
}

[[nodiscard]] HRESULT EnsureOverwriteTargetFile(const ConnectionInfo& conn, std::wstring_view destinationPath, bool allowOverwrite) noexcept
{
    FilesInformationCurl::Entry existing{};
    const HRESULT existsHr = GetEntryInfo(conn, destinationPath, existing);
    if (FAILED(existsHr))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) == existsHr ? S_OK : existsHr;
    }

    if (! allowOverwrite)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    if ((existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    const HRESULT deleteHr = RemoteDeleteFile(conn, destinationPath);
    if (FAILED(deleteHr))
    {
        return deleteHr;
    }

    return S_OK;
}
} // namespace FileSystemCurlInternal

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
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

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    return FileSystemCurlInternal::ResolveLocationWithAuthRetry(_protocol,
                                                                settings,
                                                                path,
                                                                _hostConnections.get(),
                                                                true,
                                                                [&](const FileSystemCurlInternal::ResolvedLocation& resolved) noexcept
    {
        const auto propertiesStarted = std::chrono::steady_clock::now();
        FilesInformationCurl::Entry entry{};
        HRESULT hr = FileSystemCurlInternal::GetEntryInfo(resolved.connection, resolved.remotePath, entry);
        if (FAILED(hr))
        {
            return hr;
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
        yyjson_mut_obj_add_strcpy(doc, root, "title", "properties");

        yyjson_mut_val* sections = yyjson_mut_arr(doc);
        yyjson_mut_obj_add_val(doc, root, "sections", sections);

        auto addSection = [&](const char* title) -> yyjson_mut_val*
        {
            yyjson_mut_val* section = yyjson_mut_obj(doc);
            yyjson_mut_obj_add_strcpy(doc, section, "title", title);

            yyjson_mut_val* fields = yyjson_mut_arr(doc);
            yyjson_mut_obj_add_val(doc, section, "fields", fields);

            yyjson_mut_arr_add_val(sections, section);
            return fields;
        };

        auto addField = [&](yyjson_mut_val* fields, const char* key, std::string value)
        {
            yyjson_mut_val* field = yyjson_mut_obj(doc);
            yyjson_mut_obj_add_strcpy(doc, field, "key", key);
            yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
            yyjson_mut_arr_add_val(fields, field);
        };

        auto addTimestampField = [&](yyjson_mut_val* fields, const char* key, __int64 value)
        {
            if (value != 0)
            {
                addField(fields, key, std::format("{}", value));
            }
        };

        const std::wstring normalizedPath = FileSystemCurlInternal::NormalizePluginPath(resolved.remotePath);
        const bool isDirectory            = (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

        yyjson_mut_val* general = addSection("general");
        addField(general, "name", FileSystemCurlInternal::Utf8FromUtf16(entry.name));
        addField(general, "path", FileSystemCurlInternal::Utf8FromUtf16(normalizedPath));
        addField(general, "type", isDirectory ? std::string("directory") : std::string("file"));
        addField(general, "attributes", std::format("0x{:08x}", entry.attributes));
        if (! isDirectory)
        {
            addField(general, "sizeBytes", std::format("{}", entry.sizeBytes));
        }

        yyjson_mut_val* remote = addSection("remote");
        addField(remote, "remotePath", FileSystemCurlInternal::Utf8FromUtf16(resolved.remotePath));
        addField(remote,
                 "displayPath",
                 FileSystemCurlInternal::Utf8FromUtf16(FileSystemCurlInternal::BuildDisplayPath(resolved.connection.protocol, normalizedPath)));

        yyjson_mut_val* connection = addSection("connection");
        addField(connection, "protocol", FileSystemCurlInternal::Utf8FromUtf16(FileSystemCurlInternal::ProtocolToDisplay(resolved.connection.protocol)));
        addField(connection, "host", resolved.connection.host);
        addField(connection, "user", resolved.connection.user);
        addField(connection, "basePath", resolved.connection.basePath);
        addField(connection, "fromConnectionManagerProfile", resolved.connection.fromConnectionManagerProfile ? "true" : "false");
        addField(connection, "connectionName", FileSystemCurlInternal::Utf8FromUtf16(resolved.connection.connectionName));
        addField(connection, "connectionId", FileSystemCurlInternal::Utf8FromUtf16(resolved.connection.connectionId));
        addField(connection, "connectionAuthMode", FileSystemCurlInternal::Utf8FromUtf16(resolved.connection.connectionAuthMode));
        addField(connection, "connectionSavePassword", resolved.connection.connectionSavePassword ? "true" : "false");
        addField(connection, "connectionRequireHello", resolved.connection.connectionRequireHello ? "true" : "false");
        addField(connection, "connectTimeoutMs", std::format("{}", resolved.connection.connectTimeoutMs));
        addField(connection, "operationTimeoutMs", std::format("{}", resolved.connection.operationTimeoutMs));
        addField(connection, "ignoreSslTrust", resolved.connection.ignoreSslTrust ? "true" : "false");
        addField(connection, "ftpUseEpsv", resolved.connection.ftpUseEpsv ? "true" : "false");
        addField(connection, "hasPassword", ! resolved.connection.password.empty() ? "true" : "false");
        addField(connection, "hasSshPrivateKey", ! resolved.connection.sshPrivateKey.empty() ? "true" : "false");
        addField(connection, "hasSshPublicKey", ! resolved.connection.sshPublicKey.empty() ? "true" : "false");
        addField(connection, "hasSshKnownHosts", ! resolved.connection.sshKnownHosts.empty() ? "true" : "false");

        if (resolved.connection.port.has_value())
        {
            addField(connection, "port", std::format("{}", resolved.connection.port.value()));
        }

        if (entry.creationTime != 0 || entry.lastAccessTime != 0 || entry.lastWriteTime != 0 || entry.changeTime != 0)
        {
            yyjson_mut_val* timestamps = addSection("timestamps");
            addTimestampField(timestamps, "creationTime", entry.creationTime);
            addTimestampField(timestamps, "lastAccessTime", entry.lastAccessTime);
            addTimestampField(timestamps, "lastWriteTime", entry.lastWriteTime);
            addTimestampField(timestamps, "changeTime", entry.changeTime);
        }

        if (resolved.connection.protocol == FileSystemCurlInternal::Protocol::Imap && ! isDirectory)
        {
            const std::wstring fullPath = FileSystemCurlInternal::JoinPluginPathWide(resolved.connection.basePathWide, resolved.remotePath);

            const std::wstring_view leaf = FileSystemCurlInternal::LeafName(fullPath);
            uint64_t uid                 = 0;

            yyjson_mut_val* imap = addSection("imap");
            addField(imap, "fullPath", FileSystemCurlInternal::Utf8FromUtf16(fullPath));

            if (FileSystemCurlInternal::TryParseImapUidFromLeafName(leaf, uid))
            {
                addField(imap, "uid", std::format("{}", uid));

                std::wstring mailboxPath = FileSystemCurlInternal::ParentPath(fullPath);
                mailboxPath              = std::wstring(FileSystemCurlInternal::TrimTrailingSlash(mailboxPath));
                if (mailboxPath.empty())
                {
                    mailboxPath = L"/";
                }

                if (mailboxPath != L"/")
                {
                    wchar_t delimiter = L'\0';
                    hr                = FileSystemCurlInternal::ImapGetHierarchyDelimiter(resolved.connection, delimiter);
                    if (SUCCEEDED(hr))
                    {
                        const std::wstring serverMailboxPath = FileSystemCurlInternal::ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
                        if (! serverMailboxPath.empty())
                        {
                            std::unordered_map<uint64_t, FileSystemCurlInternal::ImapMessageSummary> summaries;
                            const uint64_t uidArr[1]{uid};
                            hr = FileSystemCurlInternal::ImapFetchMessageSummaries(
                                resolved.connection, serverMailboxPath, std::span<const uint64_t>(uidArr, 1), summaries);
                            if (SUCCEEDED(hr))
                            {
                                const auto it = summaries.find(uid);
                                if (it != summaries.end())
                                {
                                    const FileSystemCurlInternal::ImapMessageSummary& s = it->second;
                                    addField(imap, "subject", FileSystemCurlInternal::Utf8FromUtf16(s.subject));
                                    addField(imap, "from", FileSystemCurlInternal::Utf8FromUtf16(s.from));
                                    addField(imap, "sentTime", std::format("{}", s.sentTime));
                                    addField(imap, "recvTime", std::format("{}", s.recvTime));
                                    addField(imap, "seen", s.seen ? "true" : "false");
                                    addField(imap, "flagged", s.flagged ? "true" : "false");
                                    addField(imap, "deleted", s.deleted ? "true" : "false");
                                    addField(imap, "sizeBytes", std::format("{}", s.sizeBytes));
                                }
                            }
                        }
                    }
                }
            }
        }
        else if (resolved.connection.protocol == FileSystemCurlInternal::Protocol::Imap && isDirectory)
        {
            const std::wstring fullPath   = FileSystemCurlInternal::JoinPluginPathWide(resolved.connection.basePathWide, resolved.remotePath);
            const std::wstring normalized = FileSystemCurlInternal::NormalizePluginPath(fullPath);
            const std::wstring mailboxPath(FileSystemCurlInternal::TrimTrailingSlash(normalized));

            yyjson_mut_val* imap = addSection("imap");
            addField(imap, "fullPath", FileSystemCurlInternal::Utf8FromUtf16(fullPath));
            addField(imap, "mailboxPath", FileSystemCurlInternal::Utf8FromUtf16(mailboxPath.empty() ? std::wstring(L"/") : mailboxPath));
            addField(imap, "isRoot", (mailboxPath.empty() || mailboxPath == L"/") ? "true" : "false");

            wchar_t delimiter = L'\0';
            hr                = FileSystemCurlInternal::ImapGetHierarchyDelimiter(resolved.connection, delimiter);
            if (SUCCEEDED(hr))
            {
                if (delimiter == L'\0')
                {
                    addField(imap, "hierarchyDelimiter", "NIL");
                }
                else
                {
                    addField(imap, "hierarchyDelimiter", FileSystemCurlInternal::Utf8FromUtf16(std::wstring(1u, delimiter)));
                }

                if (! mailboxPath.empty() && mailboxPath != L"/")
                {
                    const std::wstring serverMailboxPath = FileSystemCurlInternal::ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
                    if (! serverMailboxPath.empty())
                    {
                        addField(imap, "serverMailboxPath", FileSystemCurlInternal::Utf8FromUtf16(serverMailboxPath));

                        FileSystemCurlInternal::ImapMailboxStatus status;
                        const HRESULT statusHr = FileSystemCurlInternal::ImapFetchMailboxStatus(resolved.connection, mailboxPath, delimiter, status);
                        addField(imap, "statusAvailable", SUCCEEDED(statusHr) ? "true" : "false");
                        if (SUCCEEDED(statusHr))
                        {
                            if (status.messages.has_value())
                            {
                                addField(imap, "messages", std::format("{}", status.messages.value()));
                            }
                            if (status.recent.has_value())
                            {
                                addField(imap, "recent", std::format("{}", status.recent.value()));
                            }
                            if (status.uidNext.has_value())
                            {
                                addField(imap, "uidNext", std::format("{}", status.uidNext.value()));
                            }
                            if (status.uidValidity.has_value())
                            {
                                addField(imap, "uidValidity", std::format("{}", status.uidValidity.value()));
                            }
                            if (status.unseen.has_value())
                            {
                                addField(imap, "unseen", std::format("{}", status.unseen.value()));
                            }
                        }
                    }
                }
            }
            else
            {
                addField(imap, "metadataErrorHr", std::format("0x{:08x}", static_cast<unsigned long>(hr)));
            }
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

        if (resolved.connection.protocol == FileSystemCurlInternal::Protocol::Imap)
        {
            Debug::Perf::EmitDurationUs(isDirectory ? L"filesystem.imap.properties_mailbox_us" : L"filesystem.imap.properties_message_us",
                                        Debug::Perf::ElapsedUs(propertiesStarted),
                                        isDirectory ? 1u : entry.sizeBytes,
                                        0u,
                                        S_OK);
        }

        return S_OK;
    });
}

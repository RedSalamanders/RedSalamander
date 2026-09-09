#include "FileSystemCurl.Internal.h"
#include "FileOperationTraversalPolicy.h"

#include <charconv>
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

[[nodiscard]] size_t CurlDiscardImapBody(void*, size_t size, size_t nitems, void*) noexcept
{
    return size * nitems;
}

// See ImapResponseCapture in FileSystemCurl.Internal.h for the libcurl channel contract.
HRESULT CurlPerformImapCustomRequest(const ConnectionInfo& conn,
                                     std::wstring_view mailboxPath,
                                     std::string_view request,
                                     std::string& outResponse,
                                     HRESULT* stopReasonOut,
                                     ImapResponseCapture capture) noexcept
{
    outResponse.clear();
    // A request-local latch distinguishes callback failures from recoverable
    // metadata/server errors. Never ask an abort callback a second time to
    // reconstruct the cause; a caller is allowed to return a one-shot failure.
    struct RequestControl
    {
        const FileSystemOptions* options;
        HRESULT stopHr;
    };
    RequestControl control{CurlCurrentOperationOptions(), S_OK};
    control.stopHr        = FileSystemCheckOperationControl(control.options);
    const auto reportStop = wil::scope_exit([&]() noexcept
    {
        if (stopReasonOut != nullptr && FAILED(control.stopHr))
        {
            *stopReasonOut = control.stopHr;
        }
    });
    if (FAILED(control.stopHr))
    {
        return control.stopHr;
    }

#ifdef ENABLE_TESTS
    if (conn.imapRequestForSelfTest != nullptr)
    {
        return conn.imapRequestForSelfTest(conn.imapRequestContextForSelfTest, mailboxPath, request, outResponse);
    }
#endif

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
    if (capture == ImapResponseCapture::WireLines)
    {
        // The pooled handle is curl_easy_reset on return, so this per-call header
        // sink cannot leak into a later transfer. Body bytes are discarded: for a
        // listing-shaped FETCH they are empty, and any other shape would duplicate
        // the wire lines already captured here.
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlDiscardImapBody);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, nullptr);
        curl_easy_setopt(curl.get(), CURLOPT_HEADERFUNCTION, CurlWriteToString);
        curl_easy_setopt(curl.get(), CURLOPT_HEADERDATA, &outResponse);
    }
    else
    {
        curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToString);
        curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &outResponse);
    }
    curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

    char errorBuffer[CURL_ERROR_SIZE]{};
    curl_easy_setopt(curl.get(), CURLOPT_ERRORBUFFER, errorBuffer);

    ApplyCommonCurlOptions(curl.get(), conn, CurlCurrentOperationOptions(), false);
    curl_easy_setopt(curl.get(),
                     CURLOPT_XFERINFOFUNCTION,
                     +[](void* context, curl_off_t, curl_off_t, curl_off_t, curl_off_t) noexcept -> int
    {
        auto& current = *static_cast<RequestControl*>(context);
        if (SUCCEEDED(current.stopHr))
        {
            current.stopHr = FileSystemCheckOperationControl(current.options);
        }
        return FAILED(current.stopHr) ? 1 : 0;
    });
    curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, &control);
    curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);
    if (ImapSchemeForConnection(conn) == "imap")
    {
        curl_easy_setopt(curl.get(), CURLOPT_USE_SSL, CURLUSESSL_TRY);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code == CURLE_ABORTED_BY_CALLBACK)
    {
        // Per-call size/operation cancellation is normal control flow, not a
        // server error. Preserve the callback verdict without logging request data.
        return FAILED(control.stopHr) ? control.stopHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (code != CURLE_OK)
    {
        long responseCode = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &responseCode);

        long osErrno = 0;
        curl_easy_getinfo(curl.get(), CURLINFO_OS_ERRNO, &osErrno);

        // Body capture holds only this request's rows, so its first line explains the
        // failure. Wire capture starts with the greeting/login text, so the tagged
        // NO/BAD verdict is the last non-empty line instead.
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

            const std::string trimmedLine = TrimAscii(line);
            if (! trimmedLine.empty())
            {
                serverLine = trimmedLine;
                if (capture == ImapResponseCapture::Body)
                {
                    break;
                }
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

// Borrow the encoded contents for framing; decode only when a metadata consumer needs them.
[[nodiscard]] bool TryParseImapQuotedView(std::string_view text, size_t& pos, std::string_view& out) noexcept
{
    out = {};
    if (pos >= text.size() || text[pos] != '"')
    {
        return false;
    }
    const size_t start = ++pos;
    while (pos < text.size())
    {
        const char ch = text[pos++];
        if (ch == '"')
        {
            out = text.substr(start, pos - start - 1u);
            return true;
        }
        if (ch == '\r' || ch == '\n' || ch == '\0')
        {
            return false;
        }
        if (ch == '\\')
        {
            if (pos >= text.size() || (text[pos] != '\\' && text[pos] != '"'))
            {
                return false;
            }
            ++pos;
        }
    }
    return false;
}

[[nodiscard]] bool TryParseImapQuotedString(std::string_view text, size_t& pos, std::string& out) noexcept
{
    out.clear();
    std::string_view encoded;
    if (! TryParseImapQuotedView(text, pos, encoded))
    {
        return false;
    }
    out.reserve(encoded.size());
    for (size_t index = 0; index < encoded.size(); ++index)
    {
        if (encoded[index] == '\\')
        {
            ++index; // The framing helper validated the escape and its following byte.
        }
        out.push_back(encoded[index]);
    }
    return true;
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

[[nodiscard]] bool TryParseImapUnsignedToken(std::string_view token, uint64_t maximum, uint64_t& out) noexcept
{
    out = 0;
    if (token.empty())
    {
        return false;
    }
    uint64_t value    = 0;
    const auto result = std::from_chars(token.data(), token.data() + token.size(), value);
    if (result.ec != std::errc{} || result.ptr != token.data() + token.size() || value > maximum)
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

// IMAP wire literals differ from line/argv strings: their byte length, not delimiters,
// establishes the next token. Borrow payload bytes and check subtraction before addition.
[[nodiscard]] bool TryParseImapLiteralView(std::string_view text, size_t& pos, std::string_view& out) noexcept
{
    out           = {};
    size_t cursor = pos;
    if (cursor < text.size() && text[cursor] == '~')
    {
        ++cursor;
    }
    if (cursor >= text.size() || text[cursor++] != '{')
    {
        return false;
    }
    const size_t numberStart = cursor;
    while (cursor < text.size() && text[cursor] >= '0' && text[cursor] <= '9')
    {
        ++cursor;
    }
    uint64_t length = 0;
    // RFC 9051 number64 is unsigned 63-bit, not the whole uint64_t range.
    if (! TryParseImapUnsignedToken(text.substr(numberStart, cursor - numberStart), static_cast<uint64_t>((std::numeric_limits<int64_t>::max)()), length))
    {
        return false;
    }
    if (cursor < text.size() && text[cursor] == '+')
    {
        ++cursor; // Retain the existing non-synchronizing-literal compatibility.
    }
    if (cursor >= text.size() || text[cursor++] != '}')
    {
        return false;
    }
    if (cursor < text.size() && text[cursor] == '\r')
    {
        ++cursor;
    }
    if (cursor >= text.size() || text[cursor++] != '\n' || length > text.size() - cursor)
    {
        return false;
    }
    out = text.substr(cursor, static_cast<size_t>(length));
    pos = cursor + static_cast<size_t>(length);
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
        std::string_view value;
        if (! TryParseImapLiteralView(text, pos, value))
        {
            return false;
        }
        out.assign(value);
        return true;
    }

    return false; // RFC nstring is NIL, a quoted string, or a literal; never a bare atom.
}

[[nodiscard]] bool TrySkipImapParenthesized(std::string_view text, size_t& pos) noexcept
{
    SkipImapWhitespace(text, pos);
    if (pos >= text.size() || text[pos] != '(')
    {
        return false;
    }
    // Iterative wire framing: depth is bounded by input bytes, with constant stack
    // storage. File-tree depth policy does not apply to nested message attributes.
    size_t depth = 0;
    while (pos < text.size())
    {
        const char ch = text[pos];
        std::string_view ignored;
        if (ch == '"')
        {
            if (! TryParseImapQuotedView(text, pos, ignored))
            {
                return false;
            }
            continue;
        }
        if (ch == '{' || (ch == '~' && pos + 1u < text.size() && text[pos + 1u] == '{'))
        {
            if (! TryParseImapLiteralView(text, pos, ignored))
            {
                return false;
            }
            continue;
        }
        ++pos;
        if (ch == '(')
        {
            ++depth;
        }
        else if (ch == ')')
        {
            if (--depth == 0u)
            {
                return true;
            }
        }
        else if (ch == '\r' || ch == '\n' || ch == '\0')
        {
            return false; // Only a literal may span response lines.
        }
    }
    return false;
}

// This scalar/list grammar is IMAP-local, not a generic string or command parser.
[[nodiscard]] bool TrySkipImapValue(std::string_view text, size_t& pos) noexcept
{
    if (pos >= text.size())
    {
        return false;
    }
    std::string_view ignored;
    if (text[pos] == '(')
    {
        return TrySkipImapParenthesized(text, pos);
    }
    if (text[pos] == '"')
    {
        return TryParseImapQuotedView(text, pos, ignored);
    }
    if (text[pos] == '{' || (text[pos] == '~' && pos + 1u < text.size() && text[pos + 1u] == '{'))
    {
        return TryParseImapLiteralView(text, pos, ignored);
    }
    const size_t start = pos;
    while (pos < text.size() && text[pos] != ' ' && text[pos] != '\t' && text[pos] != ')' && text[pos] != ']')
    {
        const char ch = text[pos];
        if (static_cast<unsigned char>(ch) <= 0x20u || ch == 0x7f || ch == '(' || ch == '"' || ch == '{' || ch == '}' || ch == '[' || ch == ']')
        {
            return false;
        }
        ++pos;
    }
    return pos != start;
}
[[nodiscard]] HRESULT ImapListMessageUids(const ConnectionInfo& conn, std::wstring_view mailboxName, wchar_t delimiter, std::vector<uint64_t>& outUids) noexcept
{
    outUids.clear();
    bool complete             = false;
    const auto clearOnFailure = wil::scope_exit([&]() noexcept
    {
        if (! complete)
        {
            outUids.clear();
        }
    });

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

    // UID SEARCH is an identity enumeration, not optional display metadata.
    // Missing/malformed results cannot be published as an empty or partial mailbox.
    bool foundSearch   = false;
    size_t start       = 0u;
    size_t checkedUids = 0u;
    hr                 = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
    if (FAILED(hr))
    {
        return hr;
    }
    while (start < response.size())
    {
        hr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
        if (FAILED(hr))
        {
            return hr;
        }
        const size_t lineEnd = response.find('\n', start);
        if (lineEnd == std::string::npos)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
        }
        std::string_view line = std::string_view(response).substr(start, lineEnd - start);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1u);
        }
        const size_t frameStart = start;
        start                   = lineEnd + 1u;
        if (! line.starts_with("* "))
        {
            continue; // Tagged completions/continuations are transport-owned response text.
        }
        size_t pos                     = 2u;
        const std::string_view command = ParseImapToken(line, pos);
        if (EqualsAsciiIgnoreCase(command, "OK") || EqualsAsciiIgnoreCase(command, "NO") || EqualsAsciiIgnoreCase(command, "BAD") ||
            EqualsAsciiIgnoreCase(command, "PREAUTH") || EqualsAsciiIgnoreCase(command, "BYE"))
        {
            continue; // Status text is free-form, not a list of quoted/literal wire values.
        }
        if (EqualsAsciiIgnoreCase(command, "ESEARCH"))
        {
            // The current owner implements the rev1 SEARCH response, not ESEARCH
            // correlation/UID range expansion. Never turn unsupported results into empty.
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        if (! EqualsAsciiIgnoreCase(command, "SEARCH"))
        {
            // Skip one complete data response using the shared borrowed wire framing.
            // Newlines inside an unsolicited FETCH/extension literal are never responses.
            size_t cursor     = frameStart;
            size_t checkpoint = cursor;
            bool terminated   = false;
            while (cursor < response.size())
            {
                if (cursor - checkpoint >= 4096u)
                {
                    hr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                    checkpoint = cursor;
                }
                const char ch = response[cursor];
                std::string_view ignored;
                if (ch == '(')
                {
                    if (! TrySkipImapParenthesized(response, cursor))
                    {
                        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
                    }
                }
                else if (ch == '"')
                {
                    if (! TryParseImapQuotedView(response, cursor, ignored))
                    {
                        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
                    }
                }
                else if (ch == '{' || (ch == '~' && cursor + 1u < response.size() && response[cursor + 1u] == '{'))
                {
                    if (! TryParseImapLiteralView(response, cursor, ignored))
                    {
                        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
                    }
                }
                else if (ch == '\r' || ch == '\n')
                {
                    if (ch == '\r' && (++cursor >= response.size() || response[cursor] != '\n'))
                    {
                        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
                    }
                    ++cursor;
                    terminated = true;
                    break;
                }
                else if (ch == '\0')
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
                }
                else
                {
                    ++cursor;
                }
            }
            if (! terminated)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
            }
            start = cursor;
            continue;
        }

        foundSearch = true;
        while (pos < line.size())
        {
            const std::string_view token = ParseImapToken(line, pos);
            if (token.empty())
            {
                break;
            }
            uint64_t uid = 0u;
            if (! TryParseImapUnsignedToken(token, (std::numeric_limits<uint32_t>::max)(), uid) || uid == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
            }
            outUids.push_back(uid);
            if (++checkedUids == 256u)
            {
                checkedUids = 0u;
                hr          = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
    }
    hr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
    if (FAILED(hr))
    {
        return hr;
    }
    if (! foundSearch)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
    }
    complete = true;
    return S_OK;
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

    size_t pos = 0;
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

    return pos + 1u == fetchText.size();
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

        if (response.size() - p < 6u || response[p + 5u] != ' ')
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

struct ImapFetchAttributes
{
    std::string_view uid;
    std::string_view size;
    std::string_view flags;
    std::string_view internalDate;
    std::string_view envelope;
    std::string_view headerFields;
};

[[nodiscard]] bool IsImapSummaryHeaderSection(std::string_view key) noexcept
{
    constexpr std::string_view prefix = "BODY[HEADER.FIELDS ";
    if (key.size() <= prefix.size() || ! EqualsAsciiIgnoreCase(key.substr(0, prefix.size()), prefix) || key.back() != ']')
    {
        return false;
    }
    const std::string_view fields = TrimAsciiView(key.substr(prefix.size(), key.size() - prefix.size() - 1u));
    if (fields.size() < 2u || fields.front() != '(' || fields.back() != ')')
    {
        return false;
    }
    bool subject = false;
    bool from    = false;
    bool date    = false;
    size_t pos   = 1u;
    while (pos < fields.size() - 1u)
    {
        SkipImapWhitespace(fields, pos);
        const size_t start = pos;
        while (pos < fields.size() - 1u && fields[pos] != ' ' && fields[pos] != '\t')
        {
            ++pos;
        }
        const std::string_view name = fields.substr(start, pos - start);
        subject                     = subject || EqualsAsciiIgnoreCase(name, "SUBJECT");
        from                        = from || EqualsAsciiIgnoreCase(name, "FROM");
        date                        = date || EqualsAsciiIgnoreCase(name, "DATE");
    }
    // A partial BODY section, nested MIME header, or HEADER.FIELDS.NOT cannot
    // establish the complete message-header metadata requested by this owner.
    return subject && from && date;
}

[[nodiscard]] bool TryConsumeImapUntaggedFetchResponse(std::string_view response, size_t msgStart, size_t& outNextPos, ImapFetchAttributes& out) noexcept
{
    out                   = {};
    outNextPos            = response.size();
    size_t pos            = msgStart + 2u;
    const auto skipSpaces = [&]() noexcept
    {
        while (pos < response.size() && (response[pos] == ' ' || response[pos] == '\t'))
        {
            ++pos;
        }
    };
    skipSpaces();
    const size_t sequenceStart = pos;
    while (pos < response.size() && response[pos] >= '0' && response[pos] <= '9')
    {
        ++pos;
    }
    uint64_t sequence = 0;
    if (! TryParseImapUnsignedToken(response.substr(sequenceStart, pos - sequenceStart), (std::numeric_limits<uint32_t>::max)(), sequence) || sequence == 0u ||
        pos >= response.size() || response[pos] != ' ')
    {
        return false;
    }
    skipSpaces();
    if (! EqualsAsciiIgnoreCase(response.substr(pos, 5u), "FETCH") || response.size() - pos < 6u || response[pos + 5u] != ' ')
    {
        return false;
    }
    pos += 6u;
    skipSpaces();
    if (pos >= response.size() || response[pos++] != '(')
    {
        return false;
    }

    while (pos < response.size())
    {
        while (pos < response.size() && (response[pos] == ' ' || response[pos] == '\t'))
        {
            ++pos;
        }
        if (pos >= response.size())
        {
            return false;
        }
        if (response[pos] == ')')
        {
            ++pos;
            while (pos < response.size() && (response[pos] == ' ' || response[pos] == '\t'))
            {
                ++pos;
            }
            if (pos < response.size() && response[pos] == '\r')
            {
                ++pos;
                if (pos >= response.size() || response[pos] != '\n')
                {
                    return false;
                }
            }
            if (pos < response.size() && response[pos++] != '\n')
            {
                return false;
            }
            outNextPos = pos;
            return true;
        }

        const size_t keyStart = pos;
        while (pos < response.size() && response[pos] != '[' && response[pos] != ' ' && response[pos] != '\t')
        {
            const char ch = response[pos];
            if (static_cast<unsigned char>(ch) <= 0x20u || ch == 0x7f || ch == '(' || ch == ')' || ch == '"' || ch == '{' || ch == '}' || ch == ']')
            {
                return false;
            }
            ++pos;
        }
        if (pos == keyStart)
        {
            return false;
        }
        if (pos < response.size() && response[pos] == '[')
        {
            ++pos;
            while (pos < response.size() && response[pos] != ']')
            {
                if (response[pos] == ' ' || response[pos] == '\t')
                {
                    ++pos;
                }
                else if (! TrySkipImapValue(response, pos))
                {
                    return false;
                }
            }
            if (pos >= response.size())
            {
                return false;
            }
            ++pos;
            if (pos < response.size() && response[pos] == '<')
            {
                const size_t offsetStart = ++pos;
                while (pos < response.size() && response[pos] >= '0' && response[pos] <= '9')
                {
                    ++pos;
                }
                uint64_t offset = 0;
                if (! TryParseImapUnsignedToken(
                        response.substr(offsetStart, pos - offsetStart), static_cast<uint64_t>((std::numeric_limits<int64_t>::max)()), offset) ||
                    pos >= response.size() || response[pos++] != '>')
                {
                    return false;
                }
            }
        }
        const std::string_view key = response.substr(keyStart, pos - keyStart);
        if (pos >= response.size() || (response[pos] != ' ' && response[pos] != '\t'))
        {
            return false;
        }
        while (pos < response.size() && (response[pos] == ' ' || response[pos] == '\t'))
        {
            ++pos;
        }
        const size_t valueStart = pos;
        if (! TrySkipImapValue(response, pos) || (pos < response.size() && response[pos] != ' ' && response[pos] != '\t' && response[pos] != ')'))
        {
            return false;
        }
        std::string_view* field = nullptr;
        if (EqualsAsciiIgnoreCase(key, "UID"))
        {
            field = &out.uid;
        }
        else if (EqualsAsciiIgnoreCase(key, "RFC822.SIZE"))
        {
            field = &out.size;
        }
        else if (EqualsAsciiIgnoreCase(key, "FLAGS"))
        {
            field = &out.flags;
        }
        else if (EqualsAsciiIgnoreCase(key, "INTERNALDATE"))
        {
            field = &out.internalDate;
        }
        else if (EqualsAsciiIgnoreCase(key, "ENVELOPE"))
        {
            field = &out.envelope;
        }
        else if (IsImapSummaryHeaderSection(key))
        {
            field = &out.headerFields;
        }
        if (field)
        {
            if (! field->empty())
            {
                return false; // Duplicate metadata is ambiguous, even when one value looks usable.
            }
            *field = response.substr(valueStart, pos - valueStart);
        }
    }
    return false;
}

[[nodiscard]] bool TryParseImapFlags(std::string_view value, ImapMessageSummary& summary) noexcept
{
    if (value.size() < 2u || value.front() != '(' || value.back() != ')')
    {
        return false;
    }
    bool seen                    = false;
    bool flagged                 = false;
    bool deleted                 = false;
    const std::string_view flags = value.substr(1u, value.size() - 2u);
    size_t pos                   = 0;
    while (pos < flags.size())
    {
        SkipImapWhitespace(flags, pos);
        if (pos == flags.size())
        {
            break;
        }
        const size_t start = pos;
        // FLAGS contains atoms, never lists, quoted strings, or literal payloads.
        if (flags[pos] == '(' || flags[pos] == '"' || flags[pos] == '{' || (flags[pos] == '~' && pos + 1u < flags.size() && flags[pos + 1u] == '{') ||
            ! TrySkipImapValue(flags, pos))
        {
            return false;
        }
        const std::string_view flag = flags.substr(start, pos - start);
        seen                        = seen || EqualsAsciiIgnoreCase(flag, "\\Seen");
        flagged                     = flagged || EqualsAsciiIgnoreCase(flag, "\\Flagged");
        deleted                     = deleted || EqualsAsciiIgnoreCase(flag, "\\Deleted");
    }
    summary.seen    = seen;
    summary.flagged = flagged;
    summary.deleted = deleted;
    return true;
}
[[nodiscard]] HRESULT ImapFetchMessageSummaries(const ConnectionInfo& conn,
                                                std::wstring_view mailboxPath,
                                                std::span<const uint64_t> uids,
                                                std::unordered_map<uint64_t, ImapMessageSummary>& inOut,
                                                HRESULT* stopReasonOut = nullptr) noexcept
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

    // Damaged FETCH framing is reported as S_FALSE so a targeted lookup cannot turn
    // an unparseable row into "message not found".
    size_t damagedFetchRows = 0u;

    auto fetchAndParse = [&](size_t startIndex, size_t endIndex, std::string_view uidSetText) noexcept -> HRESULT
    {
        if (startIndex >= endIndex || endIndex > sorted.size() || uidSetText.empty())
        {
            return S_OK;
        }

        const HRESULT controlHr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
        if (FAILED(controlHr))
        {
            if (stopReasonOut != nullptr)
            {
                *stopReasonOut = controlHr;
            }
            return controlHr;
        }

        // A bare UID is not listing-shaped for libcurl (see ImapResponseCapture); send
        // the equivalent one-element range so every summary row arrives as wire lines.
        std::string uidSetRequest(uidSetText);
        if (uidSetRequest.find_first_of(",:") == std::string::npos)
        {
            uidSetRequest = std::format("{0}:{0}", uidSetRequest);
        }
        const std::string requestText = std::format("UID FETCH {} (UID FLAGS INTERNALDATE RFC822.SIZE ENVELOPE)", uidSetRequest);

        std::string response;
        const auto fetchStarted = std::chrono::steady_clock::now();
        HRESULT hr = CurlPerformImapCustomRequest(conn, mailboxPath, requestText, response, stopReasonOut, ImapResponseCapture::WireLines);
        Debug::Perf::EmitDurationUs(L"filesystem.imap.fetch_summaries_us",
                                    Debug::Perf::ElapsedUs(fetchStarted),
                                    static_cast<uint64_t>(endIndex - startIndex),
                                    static_cast<uint64_t>(uidSetText.size()),
                                    hr);
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
        size_t parsePos         = 0;
        while (true)
        {
            hr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
            if (FAILED(hr))
            {
                if (stopReasonOut != nullptr)
                {
                    *stopReasonOut = hr;
                }
                return hr;
            }
            const size_t msgStart = FindImapUntaggedFetchLine(response, parsePos);
            if (msgStart == std::string_view::npos)
            {
                break;
            }

            size_t nextPos = 0;
            ImapFetchAttributes attributes;
            if (! TryConsumeImapUntaggedFetchResponse(response, msgStart, nextPos, attributes))
            {
                ++fetchParseFailures;
                ++damagedFetchRows;
                // Framing is no longer trustworthy: a later apparent FETCH line
                // can be literal message data. Repair missing metadata separately.
                break;
            }

            uint64_t uid = 0;
            if (! TryParseImapUnsignedToken(attributes.uid, (std::numeric_limits<uint32_t>::max)(), uid) || uid == 0u)
            {
                ++missingUidCount;
                parsePos = nextPos;
                continue;
            }
            // Unsolicited FETCH rows do not belong to this request's workspace.
            if (! std::binary_search(sorted.begin() + static_cast<ptrdiff_t>(startIndex), sorted.begin() + static_cast<ptrdiff_t>(endIndex), uid))
            {
                parsePos = nextPos;
                continue;
            }

            ImapMessageSummary summary{};
            const auto existing = inOut.find(uid);
            if (existing != inOut.end())
            {
                // FLAGS-only updates retain already acquired immutable metadata.
                summary = std::move(existing->second);
            }
            summary.uid        = uid;
            uint64_t sizeBytes = 0;
            if (TryParseImapUnsignedToken(attributes.size, static_cast<uint64_t>((std::numeric_limits<int64_t>::max)()), sizeBytes))
            {
                summary.sizeKnown = true;
                summary.sizeBytes = sizeBytes;
            }
            if (TryParseImapFlags(attributes.flags, summary))
            {
                summary.flagsKnown = true;
            }

            size_t datePos = 0;
            std::string internalDate;
            __int64 received = 0;
            if (TryParseImapQuotedString(attributes.internalDate, datePos, internalDate) && datePos == attributes.internalDate.size() &&
                TryParseImapInternalDateToFileTime(internalDate, received))
            {
                summary.internalDateKnown = true;
                summary.recvTime          = received;
            }

            ImapEnvelopeFields env;
            if (TryExtractEnvelopeFields(attributes.envelope, env))
            {
                summary.headersKnown = true; // NIL/empty values are known, not missing.
                summary.subject      = DecodeRfc2047EncodedWordsToUtf16(env.subject);
                summary.from         = Utf16FromImapHeaderValue(env.fromAddrSpec);
                __int64 sentTime     = 0;
                if (TryParseRfc5322DateToFileTime(env.date, sentTime))
                {
                    summary.sentTime = sentTime;
                }
            }
            else if (! attributes.headerFields.empty())
            {
                size_t headerPos = 0;
                std::string decodedHeader;
                std::string_view headerBlock;
                bool hasHeader = false;
                if (attributes.headerFields.front() == '"')
                {
                    hasHeader   = TryParseImapQuotedString(attributes.headerFields, headerPos, decodedHeader);
                    headerBlock = decodedHeader;
                }
                else if (EqualsAsciiIgnoreCase(attributes.headerFields, "NIL"))
                {
                    hasHeader = true;
                    headerPos = attributes.headerFields.size();
                }
                else
                {
                    hasHeader = TryParseImapLiteralView(attributes.headerFields, headerPos, headerBlock);
                }
                ImapHeaderFields headers;
                if (hasHeader && headerPos == attributes.headerFields.size() && TryExtractHeaderFields(headerBlock, headers))
                {
                    summary.headersKnown = true;
                    summary.subject      = DecodeRfc2047EncodedWordsToUtf16(headers.subject);
                    summary.from         = ExtractEmailAddressFromFromHeader(headers.from);
                    __int64 sentTime     = 0;
                    if (TryParseRfc5322DateToFileTime(headers.date, sentTime))
                    {
                        summary.sentTime = sentTime;
                    }
                }
            }
            if (! summary.headersKnown)
            {
                ++envelopeParseFailures;
            }

            if (summary.sentTime == 0)
            {
                summary.sentTime = summary.recvTime;
            }
            inOut.insert_or_assign(uid, std::move(summary));
            ++fetchBlocksParsed;
            parsePos = nextPos;
        }
        Debug::Perf::EmitDurationUs(L"filesystem.imap.summary_parse_us",
                                    Debug::Perf::ElapsedUs(parseStarted),
                                    static_cast<uint64_t>(fetchBlocksParsed),
                                    static_cast<uint64_t>(fetchParseFailures + missingUidCount + envelopeParseFailures),
                                    hr);

        size_t incompleteRequested = 0;
        for (size_t i = startIndex; i < endIndex; ++i)
        {
            const auto found = inOut.find(sorted[i]);
            if (found == inOut.end() || ! found->second.HasCompleteMetadata())
            {
                ++incompleteRequested;
            }
        }
        if (fetchParseFailures > 0 || missingUidCount > 0 || envelopeParseFailures > 0 || incompleteRequested > 0)
        {
            // Response lines can contain email subjects, addresses and literal body
            // data. Diagnostics retain counts only, never a response prefix/payload.
            Debug::Warning(L"imap summary parse anomalies: fetchBlocks={} fetchParseFailures={} envelopeParseFailures={} missingUidInFetch={} "
                           L"incompleteRequested={} requested={} responseBytes={}",
                           fetchBlocksParsed,
                           fetchParseFailures,
                           envelopeParseFailures,
                           missingUidCount,
                           incompleteRequested,
                           endIndex - startIndex,
                           response.size());
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

    return damagedFetchRows != 0u ? S_FALSE : S_OK;
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

[[nodiscard]] HRESULT ImapValidateMessageUidValidity(const ConnectionInfo& conn,
                                                     std::wstring_view mailboxPath,
                                                     wchar_t delimiter,
                                                     uint64_t expectedUidValidity) noexcept
{
    ImapMailboxStatus status;
    const HRESULT hr = ImapFetchMailboxStatus(conn, mailboxPath, delimiter, status);
    if (FAILED(hr))
    {
        return hr;
    }
    return ValidateImapMessageUidValidity(expectedUidValidity, status.uidValidity);
}

[[nodiscard]] HRESULT ImapDownloadMessageToFile(const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file) noexcept
{
    const std::wstring fullPath = JoinPluginPathWide(conn.basePathWide, pluginPath);

    const std::wstring_view leaf = LeafName(fullPath);
    ImapMessageIdentity identity;
    if (! TryParseImapMessageIdentityFromLeafName(leaf, identity))
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

    hr = ImapValidateMessageUidValidity(conn, mailboxPath, delimiter, identity.uidValidity);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverMailboxPath = ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
    if (serverMailboxPath.empty())
    {
        return E_OUTOFMEMORY;
    }

    return ImapFetchMessageToFile(conn, serverMailboxPath, identity.uid, file);
}

[[nodiscard]] HRESULT ImapFetchCapabilities(const ConnectionInfo& conn, ImapCapabilities& outCapabilities) noexcept
{
    outCapabilities = {};
    std::string response;
    const auto started = std::chrono::steady_clock::now();
    HRESULT hr         = CurlPerformImapCustomRequest(conn, L"/", "CAPABILITY", response);
    if (SUCCEEDED(hr) && ! TryParseImapCapabilities(response, outCapabilities))
    {
        hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    Debug::Perf::EmitDurationUs(
        L"filesystem.imap.capability_us", Debug::Perf::ElapsedUs(started), static_cast<uint64_t>(response.size()), outCapabilities.uidPlus ? 1u : 0u, hr);
    return hr;
}

struct ImapDeleteCommandContext
{
    const ConnectionInfo* conn = nullptr;
    std::wstring_view serverMailboxPath;
    std::string response;
};

[[nodiscard]] HRESULT ExecuteImapDeleteCommand(void* opaqueContext, ImapDeleteCommand command, uint64_t uid) noexcept
{
    if (opaqueContext == nullptr)
    {
        return E_INVALIDARG;
    }

    auto& context = *static_cast<ImapDeleteCommandContext*>(opaqueContext);
    if (context.conn == nullptr)
    {
        return E_INVALIDARG;
    }
    switch (command)
    {
        case ImapDeleteCommand::AddDeletedFlag:
            return CurlPerformImapCustomRequest(
                *context.conn, context.serverMailboxPath, std::format("UID STORE {} +FLAGS.SILENT (\\Deleted)", uid), context.response);
        case ImapDeleteCommand::UidExpunge:
            return CurlPerformImapCustomRequest(*context.conn, context.serverMailboxPath, std::format("UID EXPUNGE {}", uid), context.response);
        case ImapDeleteCommand::RemoveDeletedFlag:
            return CurlPerformImapCustomRequest(
                *context.conn, context.serverMailboxPath, std::format("UID STORE {} -FLAGS.SILENT (\\Deleted)", uid), context.response);
    }
    return E_INVALIDARG;
}

[[nodiscard]] HRESULT ImapDeleteMessage(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept
{
    const std::wstring fullPath = JoinPluginPathWide(conn.basePathWide, pluginPath);

    const std::wstring_view leaf = LeafName(fullPath);
    ImapMessageIdentity identity;
    if (! TryParseImapMessageIdentityFromLeafName(leaf, identity))
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

    hr = ImapValidateMessageUidValidity(conn, mailboxPath, delimiter, identity.uidValidity);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring serverMailboxPath = ImapMailboxPathToServerMailboxPath(mailboxPath, delimiter);
    if (serverMailboxPath.empty())
    {
        return E_OUTOFMEMORY;
    }

    ImapCapabilities capabilities;
    hr = ImapFetchCapabilities(conn, capabilities);
    if (FAILED(hr))
    {
        return hr;
    }

    ImapDeleteCommandContext context{.conn = &conn, .serverMailboxPath = serverMailboxPath};
    ImapDeleteOutcome outcome;
    hr = ExecuteImapSingleMessageDelete(capabilities.uidPlus, identity.uid, ExecuteImapDeleteCommand, &context, outcome);
    if (FAILED(hr) && outcome.rollbackAttempted && FAILED(outcome.rollbackHr))
    {
        Debug::Error(L"imap single-message delete rollback failed: expungeHr={:#x} rollbackHr={:#x} mailbox='{}' uid={}",
                     hr,
                     outcome.rollbackHr,
                     mailboxPath,
                     identity.uid);
    }
    return hr;
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

[[nodiscard]] void FillImapMessageEntryFromSummary(FilesInformationCurl::Entry& entry,
                                                   const ImapMessageSummary& meta,
                                                   uint64_t uidValidity,
                                                   uint64_t uid) noexcept;

[[nodiscard]] HRESULT ImapReadDirectoryEntries(const ConnectionInfo& conn,
                                               std::wstring_view pluginPath,
                                               std::vector<FilesInformationCurl::Entry>& entries,
                                               uint64_t* summaryPeakOut = nullptr) noexcept
{
    Debug::Perf::Scope perf(L"filesystem.imap.read_directory_us");
    entries.clear();
    bool listingComplete                = false;
    const auto discardIncompleteListing = wil::scope_exit([&]() noexcept
    {
        if (! listingComplete)
        {
            entries.clear();
        }
    });
    uint64_t peakSummaries              = 0u;
    const auto reportSummaryPeak        = wil::scope_exit([&]() noexcept
    {
        Debug::Perf::EmitValue(L"filesystem.imap.summary_workspace_peak_entries", peakSummaries);
        if (summaryPeakOut != nullptr)
        {
            *summaryPeakOut = peakSummaries;
        }
    });

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
        listingComplete = true;
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
        listingComplete = true;
        return S_OK;
    }

    std::wstring mailboxPath;
    mailboxPath.reserve(mailboxName.size() + 1u);
    mailboxPath.push_back(L'/');
    mailboxPath.append(mailboxName);

    ImapMailboxStatus mailboxStatus;
    hr = ImapFetchMailboxStatus(conn, mailboxPath, delimiter, mailboxStatus);
    if (FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }
    if (! mailboxStatus.uidValidity.has_value() || mailboxStatus.uidValidity.value() == 0u)
    {
        hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        perf.SetHr(hr);
        return hr;
    }
    const uint64_t uidValidity = mailboxStatus.uidValidity.value();

    std::vector<uint64_t> uids;
    auto listUidsStarted = std::chrono::steady_clock::now();
    hr                   = ImapListMessageUids(conn, mailboxName, delimiter, uids);
    Debug::Perf::EmitDurationUs(L"filesystem.imap.list_uids_us", Debug::Perf::ElapsedUs(listUidsStarted), static_cast<uint64_t>(uids.size()), 0u, hr);
    if (FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }

    std::sort(uids.begin(), uids.end(), std::greater<>());
    uids.erase(std::unique(uids.begin(), uids.end()), uids.end());
    if (uids.empty())
    {
        perf.SetValue0(entries.size());
        listingComplete = true;
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

    constexpr size_t kFetchChunkSize = 200u;
    std::unordered_map<uint64_t, ImapMessageSummary> summaries;
    summaries.reserve(std::min(uids.size(), kFetchChunkSize));

    HRESULT metaHr                 = S_OK;
    size_t repairFetchCount        = 0;
    const size_t repairFetchBudget = ResolveImapSummaryRepairFetchBudget(uids.size());
    size_t missingSummaryCount     = 0;
    bool repairBudgetWarningLogged = false;
    uint64_t buildDurationUs       = 0u;
    uint64_t recoveredSummaryCount = 0u;
    HRESULT summaryStopReason      = S_OK;

    // IMAP pane metadata is best effort, but operation stop/deadline results
    // are not metadata failures and must never start a repair request.
    const auto checkSummaryStop = [&summaryStopReason](HRESULT result) noexcept -> HRESULT
    {
        if (FAILED(summaryStopReason))
        {
            return summaryStopReason;
        }
        result = NormalizeCancellation(result);
        if (result == HRESULT_FROM_WIN32(ERROR_CANCELLED) || result == HRESULT_FROM_WIN32(ERROR_TIMEOUT) || result == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT))
        {
            return result;
        }
        return FileSystemCheckOperationControl(CurlCurrentOperationOptions());
    };

    auto collectMissingSummaries = [&summaries](std::span<const uint64_t> requested) -> std::vector<uint64_t>
    {
        std::vector<uint64_t> missing;
        missing.reserve(requested.size());
        for (const uint64_t uid : requested)
        {
            const auto found = summaries.find(uid);
            if (found == summaries.end() || ! found->second.HasCompleteMetadata())
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

        constexpr size_t kRepairBatchSize           = 16u;
        HRESULT repairResult                        = S_OK;
        const std::vector<ImapUidBatchRange> ranges = BuildImapUidBatchRanges(missingUids.size(), kRepairBatchSize);
        for (const ImapUidBatchRange& range : ranges)
        {
            const std::span<const uint64_t> batch(missingUids.data() + range.offset, range.count);
            bool repairBudgetExhaustedInPass = false;

            if (! tryConsumeRepairFetch(L"batch", missingUids.size()))
            {
                break;
            }
            const HRESULT batchHr     = ImapFetchMessageSummaries(conn, serverMailboxPath, batch, summaries, &summaryStopReason);
            peakSummaries             = (std::max)(peakSummaries, static_cast<uint64_t>(summaries.size()));
            const HRESULT batchStopHr = checkSummaryStop(batchHr);
            if (FAILED(batchStopHr))
            {
                return batchStopHr;
            }
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
                const auto found = summaries.find(uid);
                if (found != summaries.end() && found->second.HasCompleteMetadata())
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
                const HRESULT singleHr     = ImapFetchMessageSummaries(conn, serverMailboxPath, one, summaries, &summaryStopReason);
                peakSummaries              = (std::max)(peakSummaries, static_cast<uint64_t>(summaries.size()));
                const HRESULT singleStopHr = checkSummaryStop(singleHr);
                if (FAILED(singleStopHr))
                {
                    return singleStopHr;
                }
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
        summaries.clear();
        const size_t count = std::min(kFetchChunkSize, uids.size() - start);
        const std::span<const uint64_t> chunk(uids.data() + start, count);
        const HRESULT chunkHr     = ImapFetchMessageSummaries(conn, serverMailboxPath, chunk, summaries, &summaryStopReason);
        peakSummaries             = (std::max)(peakSummaries, static_cast<uint64_t>(summaries.size()));
        const HRESULT chunkStopHr = checkSummaryStop(chunkHr);
        if (FAILED(chunkStopHr))
        {
            perf.SetHr(chunkStopHr);
            return chunkStopHr;
        }
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
            const HRESULT repairHr     = repairMissingSummaries(chunk, chunkHr);
            const HRESULT repairStopHr = checkSummaryStop(repairHr);
            if (FAILED(repairStopHr))
            {
                perf.SetHr(repairStopHr);
                return repairStopHr;
            }
            if (SUCCEEDED(repairHr) && chunkFailureRecorded)
            {
                metaHr = S_OK;
            }
        }
        else
        {
            // Some servers are picky about multi-UID FETCH sets. Repair any missed
            // summaries within the existing per-listing request budget.
            const std::vector<uint64_t> missing = collectMissingSummaries(chunk);
            if (! missing.empty())
            {
                missingSummaryCount += missing.size();
                const HRESULT repairHr     = repairMissingSummaries(missing, chunkHr);
                const HRESULT repairStopHr = checkSummaryStop(repairHr);
                if (FAILED(repairStopHr))
                {
                    perf.SetHr(repairStopHr);
                    return repairStopHr;
                }
                if (FAILED(repairHr) && SUCCEEDED(metaHr))
                {
                    metaHr = repairHr;
                }
            }
        }

        const auto buildStarted = std::chrono::steady_clock::now();
        for (const uint64_t uid : chunk)
        {
            hr = FileSystemCheckOperationControl(CurlCurrentOperationOptions());
            if (FAILED(hr))
            {
                perf.SetHr(hr);
                return hr;
            }
            FilesInformationCurl::Entry entry{};
            const auto found = summaries.find(uid);
            if (found != summaries.end())
            {
                FillImapMessageEntryFromSummary(entry, found->second, uidValidity, uid);
            }
            else
            {
                entry.attributes = FILE_ATTRIBUTE_NORMAL;
                entry.fileIndex  = (uid <= static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)())) ? static_cast<unsigned long>(uid) : 0u;
                entry.name       = std::format(L"message [{}-{}].eml", uidValidity, uid);
            }
            entries.push_back(std::move(entry));
        }
        recoveredSummaryCount += summaries.size();
        buildDurationUs += Debug::Perf::ElapsedUs(buildStarted);
    }

    if (FAILED(metaHr))
    {
        Debug::Warning(L"imap message summary fetch failed: hr={:#x} mailbox='{}' server='{}'", metaHr, mailboxName, Utf16FromUtf8(conn.host));
    }
    Debug::Perf::EmitValue(L"filesystem.imap.summary_missing_count", static_cast<uint64_t>(missingSummaryCount), metaHr);
    Debug::Perf::EmitValue(L"filesystem.imap.summary_repair_count", static_cast<uint64_t>(repairFetchCount), metaHr);

    Debug::Perf::EmitDurationUs(L"filesystem.imap.build_fileinfo_us", buildDurationUs, static_cast<uint64_t>(entries.size()), recoveredSummaryCount, metaHr);
    perf.SetValue0(static_cast<uint64_t>(entries.size()));
    perf.SetValue1(static_cast<uint64_t>(uids.size()));
    perf.SetHr(metaHr);

    listingComplete = true;
    return S_OK;
}

[[nodiscard]] void FillImapMessageEntryFromSummary(FilesInformationCurl::Entry& entry,
                                                   const ImapMessageSummary& meta,
                                                   uint64_t uidValidity,
                                                   uint64_t uid) noexcept
{
    entry.attributes    = FILE_ATTRIBUTE_NORMAL;
    entry.fileIndex     = (uid <= static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)())) ? static_cast<unsigned long>(uid) : 0u;
    entry.sizeBytes     = meta.sizeBytes;
    entry.sizeKnown     = meta.sizeKnown;
    entry.creationTime  = meta.sentTime;
    entry.changeTime    = meta.recvTime;
    entry.lastWriteTime = meta.recvTime;

    if (meta.flagged)
    {
        entry.attributes |= kImapFileAttributeMarked;
    }
    if (meta.flagsKnown && ! meta.seen)
    {
        entry.attributes |= kImapFileAttributeUnread;
    }
    if (meta.deleted)
    {
        entry.attributes |= kImapFileAttributeDeleted;
    }

    entry.name = BuildImapMessageLeafName(meta.subject, meta.from, uidValidity, uid);
    if (entry.name.empty())
    {
        entry.name = std::format(L"message [{}-{}].eml", uidValidity, uid);
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
    ImapMessageIdentity identity;
    if (TryParseImapMessageIdentityFromLeafName(leaf, identity))
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

        hr = ImapValidateMessageUidValidity(conn, mailboxPath, delimiter, identity.uidValidity);
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
        const uint64_t uidArr[1]{identity.uid};
        hr = ImapFetchMessageSummaries(conn, serverMailboxPath, std::span<const uint64_t>(uidArr, 1), summaries);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto it = summaries.find(identity.uid);
        if (it == summaries.end())
        {
            // S_FALSE means a FETCH row arrived with damaged framing. That is a bad
            // server response, never proof that the message is absent.
            return hr == S_FALSE ? HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP) : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        FillImapMessageEntryFromSummary(out, it->second, identity.uidValidity, identity.uid);
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

#ifdef ENABLE_TESTS
// Exercises the real LIST/STATUS/SEARCH/FETCH parser and repair owner without
// credentials or wire I/O. Live libcurl framing and cancellation remain separate gates.
struct ImapListingFixture final : IFileSystemOperationControl
{
    ImapListingFixture()                                     = default;
    ImapListingFixture(const ImapListingFixture&)            = delete;
    ImapListingFixture& operator=(const ImapListingFixture&) = delete;
    ImapListingFixture(ImapListingFixture&&)                 = delete;
    ImapListingFixture& operator=(ImapListingFixture&&)      = delete;

    size_t messageCount           = 3u;
    size_t searchCount            = 0u;
    size_t fetchCount             = 0u;
    // FETCH sets without ',' or ':' are not listing-shaped for libcurl and must never be sent.
    size_t bareUidFetchCount      = 0u;
    size_t cancelAtFetch          = 0u;
    bool omitBeforeCancel         = false;
    bool failFirstFetch           = false;
    bool failEveryFetch           = false;
    bool unsolicitedOtherUids     = false;
    bool flagsOnlyUpdate          = false;
    HRESULT cancelResult          = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    HRESULT oneShotControlFailure = S_OK;
    bool controlFailureReturned   = false;
    HRESULT searchControlFailure  = S_OK;
    size_t searchControlChecks    = 0u;
    size_t failAtSearchCheck      = 1u;
    std::string_view sizeToken    = "17";
    std::optional<std::string_view> searchResponse;
    std::string_view fixedFetchResponse;
    std::string_view firstFetchResponse;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        *abort = FALSE;
        if (searchCount != 0u && fetchCount == 0u && FAILED(searchControlFailure) && ++searchControlChecks == failAtSearchCheck)
        {
            return searchControlFailure;
        }
        if (fetchCount != 0u && FAILED(oneShotControlFailure) && ! controlFailureReturned)
        {
            controlFailureReturned = true;
            return oneShotControlFailure;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }

    static HRESULT Request(void* context, std::wstring_view, std::string_view request, std::string& response) noexcept
    {
        auto& fixture = *static_cast<ImapListingFixture*>(context);
        response.clear();
        if (request.starts_with("LIST "))
        {
            response = "* LIST () \"/\" \"INBOX\"\r\n";
            return S_OK;
        }
        if (request.starts_with("STATUS "))
        {
            response = std::format("* STATUS \"INBOX\" (MESSAGES {} UIDNEXT {} UIDVALIDITY 777)\r\n", fixture.messageCount, fixture.messageCount + 1u);
            return S_OK;
        }
        if (request == "UID SEARCH ALL")
        {
            ++fixture.searchCount;
            if (fixture.searchResponse.has_value())
            {
                response.assign(fixture.searchResponse.value());
                return S_OK;
            }
            response = "* SEARCH";
            for (size_t uid = 1u; uid <= fixture.messageCount; ++uid)
            {
                response.append(std::format(" {}", uid));
            }
            response.append("\r\n");
            return S_OK;
        }
        if (! request.starts_with("UID FETCH "))
        {
            return E_INVALIDARG;
        }
        ++fixture.fetchCount;
        if (fixture.fetchCount == fixture.cancelAtFetch)
        {
            return fixture.cancelResult;
        }
        if (fixture.failEveryFetch || (fixture.failFirstFetch && fixture.fetchCount == 1u))
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP);
        }
        if (fixture.omitBeforeCancel && fixture.fetchCount < fixture.cancelAtFetch)
        {
            return S_OK;
        }
        if (! fixture.fixedFetchResponse.empty() || (fixture.fetchCount == 1u && ! fixture.firstFetchResponse.empty()))
        {
            response.assign(fixture.fixedFetchResponse.empty() ? fixture.firstFetchResponse : fixture.fixedFetchResponse);
            return S_OK;
        }
        if (fixture.unsolicitedOtherUids)
        {
            for (size_t uid = 1'000'000u; uid < 1'000'500u; ++uid)
            {
                response.append(std::format("* 1 FETCH (UID {} FLAGS (\\Seen))\r\n", uid));
            }
        }
        std::string_view uidSet = request.substr(10u);
        uidSet                  = uidSet.substr(0u, uidSet.find(' '));
        if (uidSet.find_first_of(",:") == std::string_view::npos)
        {
            ++fixture.bareUidFetchCount;
        }
        const auto parseUid = [](std::string_view token, uint64_t& uid) noexcept
        {
            const auto parsed = std::from_chars(token.data(), token.data() + token.size(), uid);
            return parsed.ec == std::errc{} && parsed.ptr == token.data() + token.size();
        };
        while (! uidSet.empty())
        {
            const size_t comma           = uidSet.find(',');
            const std::string_view token = uidSet.substr(0u, comma);
            // Accept the RFC 3501 seq-range `first:last` the production owner sends for one UID.
            const size_t colon = token.find(':');
            uint64_t first     = 0u;
            uint64_t last      = 0u;
            if (! parseUid(token.substr(0u, colon), first) || ! parseUid(colon == std::string_view::npos ? token : token.substr(colon + 1u), last) ||
                last < first || last - first > 4'096u)
            {
                return E_INVALIDARG;
            }
            for (uint64_t uid = first; uid <= last; ++uid)
            {
                response.append(std::format("* {} FETCH (UID {} FLAGS (\\Seen) ", uid, uid));
                if (! fixture.sizeToken.empty())
                {
                    response.append(std::format("RFC822.SIZE {} ", fixture.sizeToken));
                }
                response.append("INTERNALDATE \"06-Sep-2026 01:00:00 +0000\" "
                                "ENVELOPE (NIL \"Subject\" NIL NIL NIL NIL NIL NIL NIL NIL))\r\n");
                if (fixture.flagsOnlyUpdate)
                {
                    response.append(std::format("* 1 FETCH (UID {} FLAGS (\\Flagged))\r\n", uid));
                }
            }
            if (comma == std::string_view::npos)
            {
                break;
            }
            uidSet.remove_prefix(comma + 1u);
        }
        return S_OK;
    }
};

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlImapListingTruthForSelfTest(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    const Common::DebugSelfTest::Check check{L"Curl IMAP listing"};
    const auto runListing =
        [&](ImapListingFixture& fixture, std::wstring_view scenario, std::vector<FilesInformationCurl::Entry>& entries, uint64_t& peak) noexcept
    {
        ConnectionInfo conn{};
        conn.protocol                      = Protocol::Imap;
        conn.host                          = "imap-selftest.invalid";
        conn.imapRequestForSelfTest        = ImapListingFixture::Request;
        conn.imapRequestContextForSelfTest = &fixture;
        FileSystemOptions options{};
        options.sizeBytes        = sizeof(options);
        options.operationControl = &fixture;
        const CurlOperationOptionsScope operationScope(&options);
        const auto started = std::chrono::steady_clock::now();
        const HRESULT hr   = ImapReadDirectoryEntries(conn, L"/INBOX", entries, &peak);
        Debug::Perf::Emit(L"FileOps.Curl.Imap.ListingFixture", scenario, Debug::Perf::ElapsedUs(started), fixture.fetchCount, peak, hr);
        return hr;
    };
    for (const size_t count : {size_t{1u}, size_t{201u}, size_t{4'097u}})
    {
        ImapListingFixture fixture{};
        fixture.messageCount = count;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"complete-{}", count), entries, peak);
        check(hr == S_OK, L"complete mailbox listing succeeds", *passed, *failed);
        check(entries.size() == count && std::ranges::all_of(entries,
                                                             [](const auto& entry) noexcept
        { return entry.sizeKnown && entry.sizeBytes == 17u && entry.name.starts_with(L"Subject [777-"); }),
              L"every message retains its size and epoch-qualified subject",
              *passed,
              *failed);
        check(peak <= 200u, L"summary workspace is bounded by one fetch chunk, not mailbox size", *passed, *failed);
        check(fixture.bareUidFetchCount == 0u, L"every summary FETCH set is listing-shaped for libcurl (comma or colon)", *passed, *failed);
    }
    for (const size_t cancelAt : {size_t{1u}, size_t{2u}, size_t{3u}, size_t{4u}})
    {
        ImapListingFixture fixture{};
        fixture.cancelAtFetch    = cancelAt == 4u ? 2u : cancelAt;
        fixture.omitBeforeCancel = cancelAt != 4u;
        fixture.messageCount     = cancelAt == 4u ? 401u : 3u;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"cancel-phase-{}", cancelAt), entries, peak);
        check(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED), L"canceled bulk/repair/later fetch remains canceled", *passed, *failed);
        check(fixture.fetchCount == fixture.cancelAtFetch, L"no summary or repair request follows cancellation", *passed, *failed);
        check(entries.empty(), L"canceled listing publishes no apparently complete entries", *passed, *failed);
    }
    for (const HRESULT stopResult : {E_ABORT, HRESULT_FROM_WIN32(ERROR_TIMEOUT), HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT)})
    {
        ImapListingFixture fixture{};
        fixture.cancelAtFetch    = 2u;
        fixture.omitBeforeCancel = true;
        fixture.cancelResult     = stopResult;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"stop-{:08X}", static_cast<unsigned long>(stopResult)), entries, peak);
        check(hr == NormalizeCancellation(stopResult), L"abort/deadline result is preserved through repair", *passed, *failed);
        check(fixture.fetchCount == 2u, L"no request follows abort/deadline", *passed, *failed);
        check(entries.empty(), L"aborted or expired listing exposes no partial result", *passed, *failed);
    }
    for (const HRESULT controlFailure : {E_ACCESSDENIED, HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), HRESULT_FROM_WIN32(ERROR_CANCELLED)})
    {
        ImapListingFixture fixture{};
        fixture.oneShotControlFailure = controlFailure;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"control-{:08X}", static_cast<unsigned long>(controlFailure)), entries, peak);
        check(hr == controlFailure, L"one-shot callback failure is not reclassified as repairable metadata", *passed, *failed);
        check(fixture.controlFailureReturned && fixture.fetchCount == 1u, L"parser observes control and does not issue repair", *passed, *failed);
        check(entries.empty(), L"callback failure exposes no partial result", *passed, *failed);
    }
    for (const bool flagsUpdate : {false, true})
    {
        ImapListingFixture fixture{};
        fixture.messageCount         = 201u;
        fixture.unsolicitedOtherUids = true;
        fixture.flagsOnlyUpdate      = flagsUpdate;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, flagsUpdate ? L"unsolicited-flags" : L"unsolicited-other-uids", entries, peak);
        check(hr == S_OK && entries.size() == 201u, L"unsolicited rows do not add or drop requested messages", *passed, *failed);
        check(std::ranges::all_of(entries,
                                  [](const auto& entry) noexcept
        { return entry.sizeKnown && entry.sizeBytes == 17u && entry.name.starts_with(L"Subject [777-") && entry.changeTime != 0; }),
              L"FLAGS-only updates retain the original message metadata",
              *passed,
              *failed);
        check(std::ranges::all_of(entries,
                                  [flagsUpdate](const auto& entry) noexcept
        {
            return ((entry.attributes & kImapFileAttributeMarked) != 0u) == flagsUpdate && ((entry.attributes & kImapFileAttributeUnread) != 0u) == flagsUpdate;
        }),
              L"requested-message flag updates replace flags without erasing other fields",
              *passed,
              *failed);
        check(peak <= 200u && ! entries.empty() && entries.front().fileIndex == 201u && entries.back().fileIndex == 1u,
              L"unsolicited traffic respects workspace bound and requested UID order",
              *passed,
              *failed);
    }
    for (const std::string_view sizeToken :
         {std::string_view{}, std::string_view{"0"}, std::string_view{"17junk"}, std::string_view{"18446744073709551616"}, std::string_view{"-1"}})
    {
        ImapListingFixture fixture{};
        fixture.sizeToken = sizeToken;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"size-{}", sizeToken.empty() ? L"missing" : Utf16FromUtf8(sizeToken)), entries, peak);
        check(hr == S_OK, L"pane metadata remains available when size is absent or malformed", *passed, *failed);
        const bool expectedKnown = sizeToken == "0";
        check(entries.size() == fixture.messageCount && std::ranges::all_of(entries,
                                                                            [expectedKnown](const auto& entry) noexcept
        { return entry.sizeKnown == expectedKnown && (! expectedKnown || entry.sizeBytes == 0u); }),
              L"missing/malformed size is unknown; explicit zero is known",
              *passed,
              *failed);
        check(peak <= 200u, L"size fallback keeps bounded summary workspace", *passed, *failed);
        if (sizeToken.empty())
        {
            ConnectionInfo conn{};
            conn.protocol                      = Protocol::Imap;
            conn.host                          = "imap-selftest.invalid";
            conn.imapRequestForSelfTest        = ImapListingFixture::Request;
            conn.imapRequestContextForSelfTest = &fixture;
            FilesInformationCurl::Entry selected{};
            const HRESULT selectedHr = ImapGetEntryInfo(conn, L"/INBOX/message [777-1].eml", selected);
            check(selectedHr == S_OK, L"targeted message lookup succeeds without a size", *passed, *failed);
            check(! selected.sizeKnown, L"targeted message lookup does not invent zero-byte proof", *passed, *failed);
            check(fixture.bareUidFetchCount == 0u, L"targeted lookup sends a one-element range, never a bare UID", *passed, *failed);
        }
    }
    {
        // A damaged FETCH row is a bad server response; only a well-formed reply that
        // simply lacks the UID proves absence.
        struct TargetedCase
        {
            std::wstring_view name;
            std::string response;
            HRESULT expected;
            bool emptyReply = false;
        };
        const std::vector<TargetedCase> targetedCases{
            {L"literal-truncated", "* 1 FETCH (UID 1 FLAGS (\\Seen) RFC822.SIZE 17 ENVELOPE (NIL {12}\r\nSub", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP)},
            {L"unterminated-record", "* 1 FETCH (UID 1 FLAGS (\\Seen) RFC822.SIZE 17\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP)},
            {L"other-uid-only", "* 2 FETCH (UID 2 FLAGS (\\Seen) RFC822.SIZE 17)\r\n", HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)},
            {L"empty-reply", std::string{}, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), true},
        };
        for (const auto& scenario : targetedCases)
        {
            ImapListingFixture fixture{};
            fixture.messageCount       = 2u;
            fixture.fixedFetchResponse = scenario.response;
            if (scenario.emptyReply)
            {
                // Reuse the omit-before-cancel path to answer S_OK with no rows.
                fixture.omitBeforeCancel = true;
                fixture.cancelAtFetch    = 99u;
            }
            ConnectionInfo conn{};
            conn.protocol                      = Protocol::Imap;
            conn.host                          = "imap-selftest.invalid";
            conn.imapRequestForSelfTest        = ImapListingFixture::Request;
            conn.imapRequestContextForSelfTest = &fixture;
            FilesInformationCurl::Entry selected{};
            const HRESULT selectedHr = ImapGetEntryInfo(conn, L"/INBOX/message [777-1].eml", selected);
            check(selectedHr == scenario.expected, L"targeted lookup distinguishes damaged framing from a genuinely absent message", *passed, *failed);
            check(fixture.fetchCount == 1u && fixture.bareUidFetchCount == 0u, L"targeted lookup issues exactly one listing-shaped FETCH", *passed, *failed);
        }
    }
    for (const bool allFail : {false, true})
    {
        ImapListingFixture fixture{};
        fixture.failFirstFetch = true;
        fixture.failEveryFetch = allFail;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, allFail ? L"repair-exhausted" : L"repair-success", entries, peak);
        check(hr == S_OK, L"ordinary metadata failure retains bounded best-effort pane fallback", *passed, *failed);
        check(entries.size() == fixture.messageCount &&
                  std::ranges::all_of(entries, [allFail](const auto& entry) noexcept { return entry.sizeKnown != allFail; }),
              L"repaired versus unavailable sizes remain distinct",
              *passed,
              *failed);
        check(fixture.fetchCount <= 1u + ResolveImapSummaryRepairFetchBudget(fixture.messageCount),
              L"ordinary metadata repair remains within its listing budget",
              *passed,
              *failed);
    }
    // FETCH attributes can be reordered; envelope/body strings are never authority.
    const auto envelope = [](std::string_view encodedSubject) { return std::format("ENVELOPE (NIL {} NIL NIL NIL NIL NIL NIL NIL NIL)", encodedSubject); };
    const auto literal  = [](std::string_view text) { return std::format("{{{}}}\r\n{}", text.size(), text); };
    const std::string normalEnvelope = envelope("\"Subject\"");
    const std::string normalFields   = "UID 1 FLAGS (\\Seen) RFC822.SIZE 17 INTERNALDATE \"06-Sep-2026 01:00:00 +0000\"";
    struct AttributeCase
    {
        std::wstring_view name;
        std::string response;
        std::wstring_view subject;
        bool knownSize = true;
        bool fallback  = false;
    };
    const std::vector<AttributeCase> attributeCases{
        {L"reordered", std::format("* 1 FETCH ({} {})\r\n", normalEnvelope, normalFields), L"Subject"},
        {L"quoted-uid", std::format("* 1 FETCH ({} {})\r\n", envelope("\"UID 999 RFC822.SIZE 888\""), normalFields), L"UID 999 RFC822.SIZE 888"},
        {L"literal-uid", std::format("* 1 FETCH ({} {})\r\n", envelope(literal("UID 999 RFC822.SIZE 888")), normalFields), L"UID 999 RFC822.SIZE 888"},
        {L"quoted-size",
         std::format("* 1 FETCH (UID 1 {} FLAGS (\\Seen) RFC822.SIZE 17 INTERNALDATE \"06-Sep-2026 01:00:00 +0000\")\r\n", envelope("\"RFC822.SIZE 888\"")),
         L"RFC822.SIZE 888"},
        {L"quoted-flags",
         std::format("* 1 FETCH (UID 1 {} FLAGS (\\Seen) RFC822.SIZE 17 INTERNALDATE \"06-Sep-2026 01:00:00 +0000\")\r\n", envelope("\"FLAGS (\\\\Deleted)\"")),
         L"FLAGS (\\Deleted)"},
        {L"prefix-uid", std::format("* 1 FETCH (X-UID 999 {} {})\r\n", normalFields, normalEnvelope), L"Subject"},
        {L"nested-attributes", std::format("* 1 FETCH (X-DETAILS (UID 999 RFC822.SIZE 888) {} {})\r\n", normalFields, normalEnvelope), L"Subject"},
        {L"duplicate-uid", std::format("* 1 FETCH ({} {} UID 2)\r\n", normalFields, normalEnvelope), {}, false, true},
        {L"duplicate-size", std::format("* 1 FETCH ({} {} RFC822.SIZE 99)\r\n", normalFields, normalEnvelope), {}, false, true},
        {L"truncated-record", std::format("* 1 FETCH ({} {}\r\n", normalFields, normalEnvelope), {}, false, true},
        {L"opaque-body", std::format("* 1 FETCH ({} BODY[] {})\r\n", normalFields, literal(envelope("\"Forged\""))), {}},
        {L"header-literal",
         std::format("* 1 FETCH (BODY[HEADER.FIELDS (SUBJECT FROM DATE)] {} {})\r\n",
                     literal("Subject: Header subject\r\nFrom: sender@example.test\r\n\r\n"),
                     normalFields),
         L"Header subject"},
        {L"escaped-subject", std::format("* 1 FETCH ({} {})\r\n", envelope("\"A \\\"quoted\\\" (subject)\""), normalFields), L"A \"quoted\" (subject)"},
        {L"empty-envelope", std::format("* 1 FETCH ({} {})\r\n", normalFields, envelope("NIL")), {}},
        {L"opaque-literal-fetch",
         std::format("* 1 FETCH ({} BODY[] {} {})\r\n", normalFields, literal("* 1 FETCH (UID 999 RFC822.SIZE 888)\r\n"), normalEnvelope),
         L"Subject"},
        {L"false-header-context",
         std::format("* 1 FETCH (X-DETAILS \"HEADER.FIELDS\" BODY[] {} {})\r\n", literal("Subject: Forged\r\n\r\n"), normalFields),
         {}},
        {L"partial-header-fields", std::format("* 1 FETCH (BODY[HEADER.FIELDS (SUBJECT)] {} {})\r\n", literal("Subject: Forged\r\n\r\n"), normalFields), {}},
        {L"partial-header-offset",
         std::format("* 1 FETCH (BODY[HEADER.FIELDS (SUBJECT FROM DATE)]<0> {} {})\r\n", literal("Subject: Forged\r\n\r\n"), normalFields),
         {}},
        {L"duplicate-flags", std::format("* 1 FETCH ({} {} FLAGS (\\Deleted))\r\n", normalFields, normalEnvelope), {}, false, true},
        {L"duplicate-envelope", std::format("* 1 FETCH ({} {} {})\r\n", normalFields, normalEnvelope, envelope("\"Forged\"")), {}, false, true},
        {L"literal-length-overflow",
         std::format("* 1 FETCH (BODY[] {{18446744073709551615}}\r\n* 1 FETCH ({} {})\r\n", normalFields, normalEnvelope),
         {},
         false,
         true},
        {L"literal-truncated", std::format("* 1 FETCH (BODY[] {{99999}}\r\n* 1 FETCH ({} {})\r\n", normalFields, normalEnvelope), {}, false, true},
        {L"quote-newline", std::format("* 1 FETCH ({} {})\r\n", normalFields, envelope("\"broken\r\nsubject\"")), {}, false, true},
        {L"fetch-prefix", std::format("* 1 FETCHX (UID 999 RFC822.SIZE 888)\r\n* 1 FETCH ({} {})\r\n", normalFields, normalEnvelope), L"Subject"},
        {L"empty-header-literal", std::format("* 1 FETCH (BODY[HEADER.FIELDS (SUBJECT FROM DATE)] {} {})\r\n", literal(""), normalFields), {}},
        {L"unavailable-flags", std::format("* 1 FETCH (UID 1 RFC822.SIZE 17 INTERNALDATE \"06-Sep-2026 01:00:00 +0000\" {})\r\n", normalEnvelope), L"Subject"},
    };
    for (const auto& scenario : attributeCases)
    {
        ImapListingFixture fixture{};
        fixture.messageCount       = 1u;
        fixture.fixedFetchResponse = scenario.response;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"attributes-{}", scenario.name), entries, peak);
        check(hr == S_OK && entries.size() == 1u, L"attribute fixture retains the requested message", *passed, *failed);
        const std::wstring expectedName = scenario.fallback ? L"message [777-1].eml" : BuildImapMessageLeafName(scenario.subject, {}, 777u, 1u);
        check(entries.size() == 1u && entries[0].fileIndex == 1u && entries[0].name == expectedName,
              L"only top-level attributes establish UID and envelope metadata",
              *passed,
              *failed);
        check(entries.size() == 1u && entries[0].sizeKnown == scenario.knownSize && (! scenario.knownSize || entries[0].sizeBytes == 17u),
              L"opaque text and malformed records never invent size proof",
              *passed,
              *failed);
        check(entries.size() == 1u && (entries[0].attributes & (kImapFileAttributeMarked | kImapFileAttributeUnread | kImapFileAttributeDeleted)) == 0u,
              L"only the declared FLAGS attribute controls message flags",
              *passed,
              *failed);
    }
    const std::vector<std::string> incompleteResponses{
        "* 1 FETCH (UID 1 FLAGS (\\Seen))\r\n",
        std::format("* 1 FETCH (UID 1 FLAGS (\\Seen) INTERNALDATE \"06-Sep-2026 01:00:00 +0000\" {})\r\n", normalEnvelope),
        std::format("* 1 FETCH ({} )\r\n", normalFields),
        std::format("* 1 FETCH (UID 1 RFC822.SIZE 17 INTERNALDATE \"06-Sep-2026 01:00:00 +0000\" {})\r\n", normalEnvelope),
        std::format("* 1 FETCH (UID 1 FLAGS (\\Seen) RFC822.SIZE 17 INTERNALDATE \"invalid\" {})\r\n", normalEnvelope),
        std::format("* 1 FETCH ({} {})\r\n", normalFields, envelope("bare-atom")),
    };
    for (size_t index = 0u; index < incompleteResponses.size(); ++index)
    {
        ImapListingFixture fixture{};
        fixture.messageCount       = 1u;
        fixture.firstFetchResponse = incompleteResponses[index];
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"incomplete-summary-{}", index), entries, peak);
        check(hr == S_OK && entries.size() == 1u, L"incomplete summary repair retains exactly one message", *passed, *failed);
        check(entries.size() == 1u && entries[0].sizeKnown && entries[0].sizeBytes == 17u && entries[0].changeTime != 0,
              L"repair acquires missing size and timestamp",
              *passed,
              *failed);
        check(
            entries.size() == 1u && entries[0].name == L"Subject [777-1].eml", L"repair acquires missing envelope without changing identity", *passed, *failed);
        check(fixture.fetchCount == 2u && peak == 1u, L"one bounded repair completes a partial summary", *passed, *failed);
    }
    struct NumberCase
    {
        std::string_view token;
        uint64_t maximum;
        bool valid;
        uint64_t value = 0u;
    };
    constexpr uint64_t uidMax  = (std::numeric_limits<uint32_t>::max)();
    constexpr uint64_t sizeMax = static_cast<uint64_t>((std::numeric_limits<int64_t>::max)());
    constexpr std::array numberCases{
        NumberCase{"4294967295", uidMax, true, uidMax},
        NumberCase{"4294967296", uidMax, false},
        NumberCase{"-1", uidMax, false},
        NumberCase{"+1", uidMax, false},
        NumberCase{"0", sizeMax, true},
        NumberCase{"9223372036854775807", sizeMax, true, sizeMax},
        NumberCase{"9223372036854775808", sizeMax, false},
        NumberCase{"18446744073709551615", sizeMax, false},
        NumberCase{"17tail", sizeMax, false},
        NumberCase{"", sizeMax, false},
    };
    for (const auto& scenario : numberCases)
    {
        uint64_t value   = 123u;
        const bool valid = TryParseImapUnsignedToken(scenario.token, scenario.maximum, value);
        check(valid == scenario.valid && value == scenario.value, L"IMAP numbers require the complete protocol-range token", *passed, *failed);
    }
    struct SearchCase
    {
        std::wstring_view name;
        std::string_view response;
        HRESULT expectedHr;
        std::array<uint64_t, 3> uids;
        size_t count;
    };
    const std::array searchCases{
        SearchCase{L"single", "* SEARCH 1\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"lowercase", "* search 1\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"mixed-case", "* SeArCh 3 1 2\r\n", S_OK, {{3u, 2u, 1u}}, 3u},
        SearchCase{L"whitespace", "* SEARCH\t1  2\t3 \t\r\n", S_OK, {{3u, 2u, 1u}}, 3u},
        SearchCase{L"empty", "* SEARCH\r\n", S_OK, {}, 0u},
        SearchCase{L"empty-space", "* SEARCH \t\r\n", S_OK, {}, 0u},
        SearchCase{L"empty-lf", "* SEARCH\n", S_OK, {}, 0u},
        SearchCase{L"maximum", "* SEARCH 4294967295\r\n", S_OK, {{4294967295u}}, 1u},
        SearchCase{L"duplicates", "* SEARCH 2 1 2 1\r\n", S_OK, {{2u, 1u}}, 2u},
        SearchCase{L"multiple", "* SEARCH 1 2\r\n* SEARCH 3 2\r\n", S_OK, {{3u, 2u, 1u}}, 3u},
        SearchCase{L"status-text", "* OK server said \"not a quoted value {99}\r\n* SEARCH 1\r\nA1 OK done\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"unrelated", "* 2 EXISTS\r\n* FLAGS (\\Seen)\r\n* SEARCH 1\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"zero", "* SEARCH 1 0 2\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"negative", "* SEARCH 1 -2 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"positive-sign", "* SEARCH 1 +2 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"uid-overflow", "* SEARCH 1 4294967296 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"uint64-maximum", "* SEARCH 1 18446744073709551615 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"uint64-overflow", "* SEARCH 1 18446744073709551616 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"numeric-prefix", "* SEARCH 1 2tail 3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"command-prefix", "* SEARCHER 1\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"missing", "* OK done\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"no-response", "", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"unterminated", "* SEARCH 1", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"truncated-cr", "* SEARCH 1\r", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"nul", std::string_view{"* SEARCH 1\0 2\r\n", 15u}, HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"range", "* SEARCH 1:3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"comma", "* SEARCH 1,3\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"opaque-fetch", "* 1 FETCH (UID 1 BODY[] {13}\r\n* SEARCH 99\r\n)\r\n* SEARCH 1\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"opaque-extension", "* X-TEST ({13}\r\n* SEARCH 99\r\n)\r\n* SEARCH 1\r\n", S_OK, {{1u}}, 1u},
        SearchCase{L"literal-only", "* 1 FETCH (UID 1 BODY[] {13}\r\n* SEARCH 99\r\n)\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"literal-truncated", "* 1 FETCH (BODY[] {999}\r\n* SEARCH 1\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"literal-overflow", "* X-TEST {18446744073709551615}\r\n* SEARCH 1\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"quote-truncated", "* X-TEST \"broken\r\n* SEARCH 1\r\n", HRESULT_FROM_WIN32(ERROR_BAD_NET_RESP), {}, 0u},
        SearchCase{L"esearch", "* ESEARCH (TAG \"A1\") UID ALL 1:3\r\n", HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), {}, 0u},
        SearchCase{L"esearch-empty", "* ESEARCH (TAG \"A1\") UID\r\n", HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), {}, 0u},
    };
    for (const auto& scenario : searchCases)
    {
        ImapListingFixture fixture{};
        fixture.searchResponse = scenario.response;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"search-{}", scenario.name), entries, peak);
        check(hr == scenario.expectedHr, L"UID search distinguishes complete, malformed and unsupported responses", *passed, *failed);
        check(entries.size() == scenario.count, L"UID search publishes exactly the complete unique message set or no partial result", *passed, *failed);
        bool identitiesMatch = entries.size() == scenario.count;
        for (size_t index = 0u; identitiesMatch && index < entries.size(); ++index)
        {
            const auto& entry = entries[index];
            identitiesMatch   = entry.fileIndex == scenario.uids[index] && entry.sizeKnown && entry.sizeBytes == 17u &&
                                entry.name == std::format(L"Subject [777-{}].eml", scenario.uids[index]);
        }
        check(identitiesMatch, L"UID search cannot invent, duplicate, truncate or reorder message identity", *passed, *failed);
        check(fixture.searchCount == 1u && fixture.fetchCount == (scenario.count == 0u ? 0u : 1u) && peak <= scenario.count,
              L"invalid or empty UID search never starts summary or repair work",
              *passed,
              *failed);
    }
    for (const HRESULT failure : {HRESULT_FROM_WIN32(ERROR_CANCELLED), E_ACCESSDENIED, HRESULT_FROM_WIN32(ERROR_TIMEOUT)})
    {
        ImapListingFixture fixture{};
        fixture.messageCount         = failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) ? 0u : (failure == E_ACCESSDENIED ? 1u : 4097u);
        fixture.searchControlFailure = failure;
        fixture.failAtSearchCheck    = fixture.messageCount > 256u ? 4u : 1u;
        std::vector<FilesInformationCurl::Entry> entries;
        uint64_t peak    = 0u;
        const HRESULT hr = runListing(fixture, std::format(L"search-control-{:08X}", static_cast<unsigned long>(failure)), entries, peak);
        check(hr == failure, L"UID search preserves cancellation and one-shot control failures", *passed, *failed);
        check(fixture.searchControlChecks == fixture.failAtSearchCheck && fixture.searchCount == 1u && fixture.fetchCount == 0u,
              L"UID parsing checks control before empty publication and during long identity lists",
              *passed,
              *failed);
        check(entries.empty() && peak == 0u, L"stopped UID search starts no summary work and publishes no partial mailbox", *passed, *failed);
    }
    return *failed == 0u ? S_OK : E_FAIL;
}
#endif

[[nodiscard]] HRESULT ReadDirectoryEntries(const ConnectionInfo& conn, std::wstring_view path, std::vector<FilesInformationCurl::Entry>& entries) noexcept
{
    if (conn.protocol == Protocol::Imap)
    {
        return ImapReadDirectoryEntries(conn, path, entries);
    }

    return CurlPerformListAndParse(conn, path, entries);
}

[[nodiscard]] HRESULT GetEntryInfo(const ConnectionInfo& conn,
                                   std::wstring_view path,
                                   FilesInformationCurl::Entry& out,
                                   CurlEntryLookupMetrics* metrics,
                                   std::function<HRESULT()> checkpoint) noexcept
{
    out = {};
    CurlEntryLookupMetrics observed{};
    const auto reportMetrics = wil::scope_exit([&]() noexcept
    {
        if (metrics)
        {
            *metrics = observed;
        }
    });
    if (! checkpoint)
    {
        checkpoint = []() noexcept { return FileSystemCheckOperationControl(CurlCurrentOperationOptions()); };
    }
    HRESULT hr = checkpoint();
    if (FAILED(hr))
    {
        return hr;
    }
    if (conn.protocol == Protocol::Imap)
    {
        return ImapGetEntryInfo(conn, path, out);
    }

    using namespace Common::FileOperations;
    if (path.size() > kTraversalMaxQueuedPathBytes / sizeof(wchar_t))
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
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

    CurlDirectoryCursor cursor;
    hr = cursor.Open(conn, parent, std::move(checkpoint));
    if (FAILED(hr))
    {
        return hr;
    }

    FilesInformationCurl::Entry candidate{};
    FilesInformationCurl::Entry entry{};
    observed.metadataBytes    = CurlDirectoryCursor::kMetadataReservationBytes + sizeof(cursor) + sizeof(candidate) + sizeof(entry);
    const uint64_t fixedPaths = cursor.RetainedPathBytes() + static_cast<uint64_t>(normalized.capacity() + parent.capacity() + 2u) * sizeof(wchar_t);
    observed.pathBytes        = fixedPaths;
    if (fixedPaths > kTraversalMaxQueuedPathBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }
    bool found = false;
    // Do not return at the first match. A later malformed/duplicate row or a
    // failed LIST terminator invalidates both positive and negative admission.
    while ((hr = cursor.Next(entry, leaf)) == S_OK)
    {
        ++observed.rows;
        observed.pathBytes =
            (std::max)(observed.pathBytes, fixedPaths + static_cast<uint64_t>(entry.name.capacity() + candidate.name.capacity() + 2u) * sizeof(wchar_t));
        if (observed.pathBytes > kTraversalMaxQueuedPathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        if (entry.name == leaf)
        {
            if (found)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            candidate = std::move(entry);
            found     = true;
        }
    }
    if (hr != S_FALSE)
    {
        return hr;
    }
    if (! found)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    out = std::move(candidate);
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
        if (! isDirectory && entry.sizeKnown)
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
                                    if (s.headersKnown)
                                    {
                                        addField(imap, "subject", FileSystemCurlInternal::Utf8FromUtf16(s.subject));
                                        addField(imap, "from", FileSystemCurlInternal::Utf8FromUtf16(s.from));
                                    }
                                    if (s.sentTime != 0)
                                    {
                                        addField(imap, "sentTime", std::format("{}", s.sentTime));
                                    }
                                    if (s.recvTime != 0)
                                    {
                                        addField(imap, "recvTime", std::format("{}", s.recvTime));
                                    }
                                    if (s.flagsKnown)
                                    {
                                        addField(imap, "seen", s.seen ? "true" : "false");
                                        addField(imap, "flagged", s.flagged ? "true" : "false");
                                        addField(imap, "deleted", s.deleted ? "true" : "false");
                                    }
                                    if (s.sizeKnown)
                                    {
                                        addField(imap, "sizeBytes", std::format("{}", s.sizeBytes));
                                    }
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

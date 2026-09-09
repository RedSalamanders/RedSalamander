#if defined(ENABLE_TESTS)

// Winsock must precede the Windows headers the plugin header pulls in.
#include <winsock2.h>
#include <ws2tcpip.h>

#include "ContentDigest.h"
#include "FileSystemS3.Internal.h"

#include <algorithm>
#include <atomic>
#include <cctype>
#include <charconv>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <format>
#include <limits>
#include <map>
#include <memory>
#include <mutex>
#include <set>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#pragma comment(lib, "Ws2_32.lib")

// R0f-S3 fixture: a deterministic loopback S3 endpoint (path-style REST over HTTP/1.1) that keeps
// objects in memory. It serves connections concurrently the way the CRT client uses them, records
// every request, can hold the reply to one method (a server that stopped answering) and can drip a
// listing body slowly (bytes still flowing). Host self-tests start it through the exports at the
// end of this file; the debug self-test below proves the provider-owned bound and Cancel on it.
namespace FileSystemS3Internal
{
namespace
{
using namespace std::chrono_literals;

constexpr Common::DebugSelfTest::Check DebugCheck{L"S3"};
constexpr auto kSocketPollInterval    = 50ms;
constexpr auto kIdleConnectionTimeout = 30s;
constexpr std::string_view kS3Xmlns   = "http://s3.amazonaws.com/doc/2006-03-01/";

[[nodiscard]] HRESULT SocketErrorToHResult() noexcept
{
    const int error = WSAGetLastError();
    return error == 0 ? E_FAIL : HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
}

[[nodiscard]] HRESULT CreateLoopbackListener(wil::unique_socket& listenerOut, unsigned short& portOut) noexcept
{
    listenerOut.reset();
    portOut = 0u;

    wil::unique_socket listener(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
    if (! listener)
    {
        return SocketErrorToHResult();
    }

    sockaddr_in address{};
    address.sin_family      = AF_INET;
    address.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
    address.sin_port        = 0u;
    if (bind(listener.get(), reinterpret_cast<const sockaddr*>(&address), static_cast<int>(sizeof(address))) != 0)
    {
        return SocketErrorToHResult();
    }
    if (listen(listener.get(), SOMAXCONN) != 0)
    {
        return SocketErrorToHResult();
    }

    int addressBytes = static_cast<int>(sizeof(address));
    if (getsockname(listener.get(), reinterpret_cast<sockaddr*>(&address), &addressBytes) != 0)
    {
        return SocketErrorToHResult();
    }
    portOut = ntohs(address.sin_port);
    if (portOut == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    listenerOut = std::move(listener);
    return S_OK;
}

[[nodiscard]] HRESULT WaitForReadable(SOCKET socketValue, std::stop_token stopToken, std::chrono::steady_clock::time_point deadline) noexcept
{
    while (! stopToken.stop_requested())
    {
        const auto now = std::chrono::steady_clock::now();
        if (now >= deadline)
        {
            return HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT);
        }
        const auto wait = (std::min)(std::chrono::duration_cast<std::chrono::microseconds>(deadline - now),
                                     std::chrono::duration_cast<std::chrono::microseconds>(kSocketPollInterval));
        timeval timeout{};
        timeout.tv_sec  = static_cast<long>(wait.count() / 1000000ll);
        timeout.tv_usec = static_cast<long>(wait.count() % 1000000ll);

        fd_set readSet;
        FD_ZERO(&readSet);
        FD_SET(socketValue, &readSet);
        const int selectResult = select(0, &readSet, nullptr, nullptr, &timeout);
        if (selectResult > 0)
        {
            return S_OK;
        }
        if (selectResult < 0)
        {
            return SocketErrorToHResult();
        }
    }
    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] HRESULT SendAll(SOCKET socketValue, const void* data, size_t size) noexcept
{
    const char* bytes = static_cast<const char*>(data);
    size_t offset     = 0u;
    while (offset < size)
    {
        const size_t remaining = size - offset;
        const int request      = static_cast<int>((std::min)(remaining, static_cast<size_t>((std::numeric_limits<int>::max)())));
        const int sent         = send(socketValue, bytes + offset, request, 0);
        if (sent == SOCKET_ERROR)
        {
            return SocketErrorToHResult();
        }
        offset += static_cast<size_t>(sent);
    }
    return S_OK;
}

[[nodiscard]] HRESULT SendAll(SOCKET socketValue, std::string_view text) noexcept
{
    return SendAll(socketValue, text.data(), text.size());
}

[[nodiscard]] std::string ToLowerAscii(std::string_view text)
{
    std::string out(text);
    for (char& c : out)
    {
        c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    }
    return out;
}

[[nodiscard]] int HexValue(char c) noexcept
{
    if (c >= '0' && c <= '9')
    {
        return c - '0';
    }
    if (c >= 'a' && c <= 'f')
    {
        return 10 + (c - 'a');
    }
    if (c >= 'A' && c <= 'F')
    {
        return 10 + (c - 'A');
    }
    return -1;
}

[[nodiscard]] std::string PercentDecode(std::string_view text)
{
    std::string out;
    out.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i)
    {
        if (text[i] == '%' && i + 2 < text.size() + 0u && HexValue(text[i + 1]) >= 0 && HexValue(text[i + 2]) >= 0)
        {
            out.push_back(static_cast<char>((HexValue(text[i + 1]) << 4) | HexValue(text[i + 2])));
            i += 2;
        }
        else
        {
            out.push_back(text[i]);
        }
    }
    return out;
}

[[nodiscard]] std::string XmlEscape(std::string_view text)
{
    std::string out;
    out.reserve(text.size());
    for (const char c : text)
    {
        switch (c)
        {
            case '&': out += "&amp;"; break;
            case '<': out += "&lt;"; break;
            case '>': out += "&gt;"; break;
            case '"': out += "&quot;"; break;
            default: out.push_back(c); break;
        }
    }
    return out;
}

[[nodiscard]] std::string ComputeEtag(const std::vector<uint8_t>& bytes)
{
    // Two independent FNV-1a lanes give a stable 128-bit quoted tag; clients compare, never decode.
    uint64_t a = 0xcbf29ce484222325ull;
    uint64_t b = 0x84222325cbf29ce4ull;
    for (const uint8_t value : bytes)
    {
        a = (a ^ value) * 0x100000001b3ull;
        b = (b ^ static_cast<uint8_t>(value + 0x5Au)) * 0x100000001b3ull;
    }
    return std::format("\"{:016x}{:016x}\"", a, b);
}

[[nodiscard]] std::chrono::sys_seconds NowSeconds() noexcept
{
    return std::chrono::floor<std::chrono::seconds>(std::chrono::system_clock::now());
}

[[nodiscard]] std::string Rfc1123(std::chrono::sys_seconds when)
{
    return std::format("{:%a, %d %b %Y %H:%M:%S} GMT", when);
}

[[nodiscard]] std::string Iso8601(std::chrono::sys_seconds when)
{
    return std::format("{:%Y-%m-%dT%H:%M:%S}.000Z", when);
}

[[nodiscard]] std::string StripQuotes(std::string_view text)
{
    if (text.size() >= 2u && text.front() == '"' && text.back() == '"')
    {
        text = text.substr(1, text.size() - 2u);
    }
    return std::string(text);
}

struct HttpRequest final
{
    std::string method;
    std::string path;                           // percent-decoded, always starts with '/'
    std::map<std::string, std::string> query;   // decoded
    std::map<std::string, std::string> headers; // lower-case names
    std::vector<uint8_t> body;
    bool closeAfter = false;

    [[nodiscard]] const std::string* Header(std::string_view name) const noexcept
    {
        const auto it = headers.find(std::string(name));
        return it == headers.end() ? nullptr : &it->second;
    }
    [[nodiscard]] bool HasQuery(std::string_view name) const noexcept
    {
        return query.contains(std::string(name));
    }
    [[nodiscard]] std::string Query(std::string_view name) const
    {
        const auto it = query.find(std::string(name));
        return it == query.end() ? std::string{} : it->second;
    }
};

struct HttpResponse final
{
    int status = 200;
    std::vector<std::pair<std::string, std::string>> headers;
    std::vector<uint8_t> body;
    bool headOnly = false; // send the body's Content-Length but no body (HEAD)
    bool dripBody = false; // send the body in slow slices (bytes still flowing)
};

[[nodiscard]] std::string_view ReasonPhrase(int status) noexcept
{
    switch (status)
    {
        case 100: return "Continue";
        case 200: return "OK";
        case 204: return "No Content";
        case 206: return "Partial Content";
        case 304: return "Not Modified";
        case 400: return "Bad Request";
        case 404: return "Not Found";
        case 405: return "Method Not Allowed";
        case 411: return "Length Required";
        case 412: return "Precondition Failed";
        case 416: return "Range Not Satisfiable";
        case 500: return "Internal Server Error";
        default: return "Unknown";
    }
}

[[nodiscard]] HttpResponse MakeErrorResponse(int status, std::string_view code, std::string_view message, std::string_view resource)
{
    HttpResponse response;
    response.status       = status;
    const std::string xml = std::format(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Error><Code>{}</Code><Message>{}</Message><Resource>{}</Resource><RequestId>fake</RequestId></Error>",
        code,
        XmlEscape(message),
        XmlEscape(resource));
    response.body.assign(xml.begin(), xml.end());
    response.headers.emplace_back("Content-Type", "application/xml");
    return response;
}

[[nodiscard]] HttpResponse MakeXmlResponse(int status, std::string xml)
{
    HttpResponse response;
    response.status = status;
    response.body.assign(xml.begin(), xml.end());
    response.headers.emplace_back("Content-Type", "application/xml");
    return response;
}

[[nodiscard]] HttpResponse MakeEmptyResponse(int status)
{
    HttpResponse response;
    response.status = status;
    return response;
}

// Decodes HTTP chunked framing; the same framing (with chunk extensions) is what aws-chunked uses.
[[nodiscard]] bool TryDecodeChunked(std::string_view encoded, std::vector<uint8_t>& out, size_t& consumed) noexcept
{
    out.clear();
    consumed     = 0u;
    size_t index = 0u;
    while (true)
    {
        const size_t lineEnd = encoded.find("\r\n", index);
        if (lineEnd == std::string_view::npos)
        {
            return false;
        }
        std::string_view sizeText = encoded.substr(index, lineEnd - index);
        if (const size_t ext = sizeText.find(';'); ext != std::string_view::npos)
        {
            sizeText = sizeText.substr(0, ext);
        }
        size_t chunkSize  = 0u;
        const auto parsed = std::from_chars(sizeText.data(), sizeText.data() + sizeText.size(), chunkSize, 16);
        if (parsed.ec != std::errc{})
        {
            return false;
        }
        index = lineEnd + 2u;
        if (chunkSize == 0u)
        {
            // Trailers (if any) end with an empty line.
            const size_t trailersEnd = encoded.find("\r\n\r\n", index >= 2u ? index - 2u : index);
            if (trailersEnd == std::string_view::npos)
            {
                return false;
            }
            consumed = trailersEnd + 4u;
            return true;
        }
        if (encoded.size() < index + chunkSize + 2u)
        {
            return false;
        }
        out.insert(out.end(), encoded.begin() + static_cast<ptrdiff_t>(index), encoded.begin() + static_cast<ptrdiff_t>(index + chunkSize));
        index += chunkSize + 2u;
    }
}

struct StoredObject final
{
    std::vector<uint8_t> bytes;
    std::string etag;
    std::string crc64NvmeBase64; // S3 reports the full-object CRC-64/NVME of stored content (R3-2 writer proof)
    std::chrono::sys_seconds modified{};
};

[[nodiscard]] std::string ComputeCrc64NvmeBase64(const std::vector<uint8_t>& bytes)
{
    std::vector<std::byte> digest;
    static_cast<void>(Common::Crypto::ComputeContentDigest(Common::Crypto::ContentDigestAlgorithm::Crc64Nvme, std::as_bytes(std::span(bytes)), digest));
    return Common::Crypto::EncodeBase64Digest(digest);
}

struct PendingUpload final
{
    std::string bucket;
    std::string key;
    std::map<unsigned int, std::vector<uint8_t>> parts;
};

struct RequestRecord final
{
    std::string method;
    std::string path;
    int status = 0; // 0 until the reply was built (a held reply stays 0 while it is held)
};

class FakeS3Endpoint final
{
public:
    FakeS3Endpoint() = default;
    ~FakeS3Endpoint()
    {
        Stop();
    }
    FakeS3Endpoint(const FakeS3Endpoint&)            = delete;
    FakeS3Endpoint& operator=(const FakeS3Endpoint&) = delete;
    FakeS3Endpoint(FakeS3Endpoint&&)                 = delete;
    FakeS3Endpoint& operator=(FakeS3Endpoint&&)      = delete;

    [[nodiscard]] unsigned short Port() const noexcept
    {
        return _port;
    }

    void SeedBucket(std::string bucket)
    {
        std::scoped_lock lock(_stateMutex);
        _buckets.insert(std::move(bucket));
    }

    void SeedObject(std::string_view bucket, std::string_view key, std::vector<uint8_t> bytes)
    {
        std::scoped_lock lock(_stateMutex);
        _buckets.insert(std::string(bucket));
        StoreLocked(bucket, key, std::move(bytes));
    }

    [[nodiscard]] bool HasObject(std::string_view bucket, std::string_view key) const
    {
        std::scoped_lock lock(_stateMutex);
        return _objects.contains(ObjectId(bucket, key));
    }

    // The fixture holds the reply to `method` for `stallMs` (or until it stops): a server that
    // stopped answering. The request is recorded before the hold.
    void SetStallMethod(std::string method, unsigned int stallMs)
    {
        std::scoped_lock lock(_stateMutex);
        _stallMethod = std::move(method);
        _stallMs     = stallMs;
    }

    // Listing bodies are sent `bytesPerTick` at a time with `tickMs` between slices: bytes flow,
    // slowly, so a cancel while data is still arriving has a witness.
    void SetDripListing(size_t bytesPerTick, unsigned int tickMs)
    {
        std::scoped_lock lock(_stateMutex);
        _dripBytesPerTick = bytesPerTick;
        _dripTickMs       = tickMs;
    }

    [[nodiscard]] size_t RequestCount(std::string_view method) const
    {
        std::scoped_lock lock(_stateMutex);
        return static_cast<size_t>(std::ranges::count_if(_requests, [&](const RequestRecord& record) noexcept { return record.method == method; }));
    }

    // The last `maxEntries` requests as " [METHOD path -> status]" for failure messages.
    [[nodiscard]] std::wstring RequestLog(size_t maxEntries) const
    {
        std::scoped_lock lock(_stateMutex);
        std::wstring text;
        const size_t start = _requests.size() > maxEntries ? _requests.size() - maxEntries : 0u;
        for (size_t index = start; index < _requests.size(); ++index)
        {
            text += std::format(L" [{} {} -> {}]", Utf16FromUtf8(_requests[index].method), Utf16FromUtf8(_requests[index].path), _requests[index].status);
        }
        return text;
    }

    [[nodiscard]] HRESULT Start()
    {
        const HRESULT hr = CreateLoopbackListener(_listener, _port);
        if (FAILED(hr))
        {
            return hr;
        }
        _thread = std::jthread([this](std::stop_token stopToken) noexcept
        {
            try
            {
                ServerMain(stopToken);
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::exception&)
            {
                Debug::Error(L"FileSystemS3 fake S3 endpoint terminated after std::exception.");
                _serverHr.store(E_FAIL, std::memory_order_release);
            }
        });
        return S_OK;
    }

    void Stop() noexcept
    {
        if (_thread.joinable())
        {
            _thread.request_stop();
            _thread.join();
        }
        _listener.reset();
        std::vector<std::jthread> clients;
        {
            std::scoped_lock lock(_clientsMutex);
            clients.swap(_clients);
        }
        for (std::jthread& client : clients)
        {
            client.request_stop();
        }
        clients.clear(); // joins
    }

private:
    [[nodiscard]] static std::string ObjectId(std::string_view bucket, std::string_view key)
    {
        std::string id(bucket);
        id.push_back('/');
        id.append(key);
        return id;
    }

    void StoreLocked(std::string_view bucket, std::string_view key, std::vector<uint8_t> bytes)
    {
        StoredObject object;
        object.etag            = ComputeEtag(bytes);
        object.crc64NvmeBase64 = ComputeCrc64NvmeBase64(bytes);
        object.modified        = NowSeconds();
        object.bytes           = std::move(bytes);
        _objects.insert_or_assign(ObjectId(bucket, key), std::move(object));
    }

    void ServerMain(std::stop_token stopToken)
    {
        while (! stopToken.stop_requested())
        {
            const HRESULT waitHr = WaitForReadable(_listener.get(), stopToken, std::chrono::steady_clock::now() + kSocketPollInterval);
            if (waitHr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT))
            {
                continue;
            }
            if (waitHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return;
            }
            if (FAILED(waitHr))
            {
                _serverHr.store(waitHr, std::memory_order_release);
                return;
            }
            wil::unique_socket client(accept(_listener.get(), nullptr, nullptr));
            if (! client)
            {
                _serverHr.store(SocketErrorToHResult(), std::memory_order_release);
                return;
            }
            std::scoped_lock lock(_clientsMutex);
            std::erase_if(_clients, [](const std::jthread& worker) noexcept { return ! worker.joinable(); });
            _clients.emplace_back([this, clientSocket = client.release()](std::stop_token clientStopToken) noexcept
            {
                wil::unique_socket owned(clientSocket);
                try
                {
                    HandleClient(owned.get(), clientStopToken);
                }
                catch (const std::bad_alloc&)
                {
                    std::terminate();
                }
                catch (const std::exception&)
                {
                    _serverHr.store(E_FAIL, std::memory_order_release);
                }
            });
        }
    }

    // Reads more bytes into `buffer`; ERROR_CONNECTION_ABORTED when the peer closed.
    [[nodiscard]] static HRESULT ReceiveMore(SOCKET socketValue, std::stop_token stopToken, std::string& buffer)
    {
        const HRESULT waitHr = WaitForReadable(socketValue, stopToken, std::chrono::steady_clock::now() + kIdleConnectionTimeout);
        if (FAILED(waitHr))
        {
            return waitHr;
        }
        char chunk[16384];
        const int received = recv(socketValue, chunk, static_cast<int>(sizeof(chunk)), 0);
        if (received == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        }
        if (received < 0)
        {
            return SocketErrorToHResult();
        }
        buffer.append(chunk, static_cast<size_t>(received));
        return S_OK;
    }

    [[nodiscard]] static bool ParseRequestHead(std::string_view head, HttpRequest& request)
    {
        const size_t lineEnd = head.find("\r\n");
        if (lineEnd == std::string_view::npos)
        {
            return false;
        }
        const std::string_view requestLine = head.substr(0, lineEnd);
        const size_t firstSpace            = requestLine.find(' ');
        const size_t lastSpace             = requestLine.rfind(' ');
        if (firstSpace == std::string_view::npos || lastSpace == firstSpace)
        {
            return false;
        }
        request.method                 = std::string(requestLine.substr(0, firstSpace));
        const std::string_view target  = requestLine.substr(firstSpace + 1u, lastSpace - firstSpace - 1u);
        const std::string_view version = requestLine.substr(lastSpace + 1u);
        request.closeAfter             = version == "HTTP/1.0";
        // The CRT sends some requests in absolute form (`PUT http://host:port/bucket/key HTTP/1.1`).
        std::string_view relativeTarget = target;
        if (const size_t scheme = relativeTarget.find("://"); scheme != std::string_view::npos && scheme < 8u)
        {
            const size_t pathStart = relativeTarget.find('/', scheme + 3u);
            relativeTarget         = pathStart == std::string_view::npos ? std::string_view("/") : relativeTarget.substr(pathStart);
        }
        const size_t queryStart = relativeTarget.find('?');
        request.path            = PercentDecode(relativeTarget.substr(0, queryStart));
        if (request.path.empty() || request.path.front() != '/')
        {
            request.path.insert(request.path.begin(), '/');
        }
        if (queryStart != std::string_view::npos)
        {
            std::string_view query = relativeTarget.substr(queryStart + 1u);
            while (! query.empty())
            {
                const size_t amp            = query.find('&');
                const std::string_view pair = query.substr(0, amp);
                const size_t eq             = pair.find('=');
                request.query.insert_or_assign(PercentDecode(pair.substr(0, eq)),
                                               eq == std::string_view::npos ? std::string{} : PercentDecode(pair.substr(eq + 1u)));
                query = amp == std::string_view::npos ? std::string_view{} : query.substr(amp + 1u);
            }
        }
        size_t index = lineEnd + 2u;
        while (index < head.size())
        {
            const size_t end            = head.find("\r\n", index);
            const std::string_view line = head.substr(index, end == std::string_view::npos ? std::string_view::npos : end - index);
            if (line.empty())
            {
                break;
            }
            const size_t colon = line.find(':');
            if (colon != std::string_view::npos)
            {
                std::string_view value = line.substr(colon + 1u);
                while (! value.empty() && (value.front() == ' ' || value.front() == '\t'))
                {
                    value.remove_prefix(1);
                }
                while (! value.empty() && (value.back() == ' ' || value.back() == '\t'))
                {
                    value.remove_suffix(1);
                }
                request.headers.insert_or_assign(ToLowerAscii(line.substr(0, colon)), std::string(value));
            }
            if (end == std::string_view::npos)
            {
                break;
            }
            index = end + 2u;
        }
        if (const std::string* connection = request.Header("connection"); connection != nullptr && ToLowerAscii(*connection) == "close")
        {
            request.closeAfter = true;
        }
        return true;
    }

    // One request per call; `buffer` keeps pipelined bytes between calls.
    [[nodiscard]] HRESULT ReadRequest(SOCKET socketValue, std::stop_token stopToken, std::string& buffer, HttpRequest& request)
    {
        request        = {};
        size_t headEnd = std::string::npos;
        while ((headEnd = buffer.find("\r\n\r\n")) == std::string::npos)
        {
            const HRESULT hr = ReceiveMore(socketValue, stopToken, buffer);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        if (! ParseRequestHead(std::string_view(buffer).substr(0, headEnd + 2u), request))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        buffer.erase(0, headEnd + 4u);

        const std::string* transferEncoding = request.Header("transfer-encoding");
        const bool chunked                  = transferEncoding != nullptr && ToLowerAscii(*transferEncoding).find("chunked") != std::string::npos;
        size_t contentLength                = 0u;
        if (const std::string* lengthHeader = request.Header("content-length"); lengthHeader != nullptr)
        {
            const auto parsed = std::from_chars(lengthHeader->data(), lengthHeader->data() + lengthHeader->size(), contentLength);
            if (parsed.ec != std::errc{})
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        if (const std::string* expect = request.Header("expect"); expect != nullptr && ToLowerAscii(*expect) == "100-continue")
        {
            const HRESULT hr = SendAll(socketValue, "HTTP/1.1 100 Continue\r\n\r\n");
            if (FAILED(hr))
            {
                return hr;
            }
        }
        if (chunked)
        {
            size_t consumed = 0u;
            while (! TryDecodeChunked(buffer, request.body, consumed))
            {
                const HRESULT hr = ReceiveMore(socketValue, stopToken, buffer);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
            buffer.erase(0, consumed);
        }
        else if (contentLength > 0u)
        {
            while (buffer.size() < contentLength)
            {
                const HRESULT hr = ReceiveMore(socketValue, stopToken, buffer);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
            request.body.assign(buffer.begin(), buffer.begin() + static_cast<ptrdiff_t>(contentLength));
            buffer.erase(0, contentLength);
            const std::string* contentEncoding = request.Header("content-encoding");
            if (contentEncoding != nullptr && ToLowerAscii(*contentEncoding).find("aws-chunked") != std::string::npos)
            {
                std::vector<uint8_t> decoded;
                size_t consumed = 0u;
                const std::string_view framed(reinterpret_cast<const char*>(request.body.data()), request.body.size());
                if (! TryDecodeChunked(framed, decoded, consumed))
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                request.body = std::move(decoded);
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT SendResponse(
        SOCKET socketValue, std::stop_token stopToken, const HttpResponse& response, size_t dripBytesPerTick, unsigned int dripTickMs)
    {
        std::string head = std::format("HTTP/1.1 {} {}\r\nDate: {}\r\nServer: RedSalamanderFakeS3\r\nx-amz-request-id: {}\r\nContent-Length: {}\r\n",
                                       response.status,
                                       ReasonPhrase(response.status),
                                       Rfc1123(NowSeconds()),
                                       _requestSerial.fetch_add(1u, std::memory_order_relaxed) + 1u,
                                       response.body.size());
        for (const auto& [name, value] : response.headers)
        {
            head.append(name).append(": ").append(value).append("\r\n");
        }
        head.append("Connection: keep-alive\r\n\r\n");
        HRESULT hr = SendAll(socketValue, head);
        if (FAILED(hr) || response.headOnly || response.body.empty())
        {
            return hr;
        }
        if (response.dripBody && dripBytesPerTick > 0u)
        {
            size_t offset = 0u;
            while (offset < response.body.size())
            {
                if (stopToken.stop_requested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                const size_t slice = (std::min)(dripBytesPerTick, response.body.size() - offset);
                hr                 = SendAll(socketValue, response.body.data() + offset, slice);
                if (FAILED(hr))
                {
                    return hr;
                }
                offset += slice;
                std::this_thread::sleep_for(std::chrono::milliseconds(dripTickMs));
            }
            return S_OK;
        }
        return SendAll(socketValue, response.body.data(), response.body.size());
    }

    void HandleClient(SOCKET socketValue, std::stop_token stopToken)
    {
        std::string buffer;
        while (! stopToken.stop_requested())
        {
            HttpRequest request;
            const HRESULT readHr = ReadRequest(socketValue, stopToken, buffer, request);
            if (FAILED(readHr))
            {
                return; // idle timeout, peer closed, or stop: the connection just ends
            }

            std::string stallMethod;
            unsigned int stallMs = 0u;
            size_t dripBytes     = 0u;
            unsigned int dripMs  = 0u;
            size_t recordIndex   = 0u;
            {
                std::scoped_lock lock(_stateMutex);
                recordIndex = _requests.size();
                _requests.push_back(RequestRecord{.method = request.method, .path = request.path, .status = 0});
                stallMethod = _stallMethod;
                stallMs     = _stallMs;
                dripBytes   = _dripBytesPerTick;
                dripMs      = _dripTickMs;
            }
            if (! stallMethod.empty() && request.method == stallMethod)
            {
                const auto stallUntil = std::chrono::steady_clock::now() + std::chrono::milliseconds(stallMs);
                while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < stallUntil)
                {
                    std::this_thread::sleep_for(kSocketPollInterval);
                }
                if (stopToken.stop_requested())
                {
                    return;
                }
            }

            HttpResponse response = Dispatch(request);
            if (request.method == "HEAD")
            {
                response.headOnly = true;
            }
            {
                std::scoped_lock lock(_stateMutex);
                if (recordIndex < _requests.size())
                {
                    _requests[recordIndex].status = response.status;
                }
            }
            const HRESULT sendHr = SendResponse(socketValue, stopToken, response, dripBytes, dripMs);
            if (FAILED(sendHr) || request.closeAfter)
            {
                return;
            }
        }
    }

    [[nodiscard]] HttpResponse Dispatch(const HttpRequest& request)
    {
        std::string_view rest(request.path);
        rest.remove_prefix(1); // leading '/'
        const size_t slash       = rest.find('/');
        const std::string bucket = std::string(rest.substr(0, slash));
        const std::string key    = slash == std::string_view::npos ? std::string{} : std::string(rest.substr(slash + 1u));

        std::scoped_lock lock(_stateMutex);
        if (bucket.empty())
        {
            return request.method == "GET" ? ListBucketsLocked() : MakeErrorResponse(405, "MethodNotAllowed", "root", request.path);
        }
        if (key.empty())
        {
            return DispatchBucketLocked(request, bucket);
        }
        if (! _buckets.contains(bucket))
        {
            return MakeErrorResponse(404, "NoSuchBucket", "The specified bucket does not exist", request.path);
        }
        if (request.method == "GET")
        {
            return GetObjectLocked(request, bucket, key);
        }
        if (request.method == "HEAD")
        {
            return HeadObjectLocked(request, bucket, key);
        }
        if (request.method == "PUT")
        {
            return PutLocked(request, bucket, key);
        }
        if (request.method == "DELETE")
        {
            return DeleteLocked(request, bucket, key);
        }
        if (request.method == "POST")
        {
            return PostLocked(request, bucket, key);
        }
        return MakeErrorResponse(405, "MethodNotAllowed", request.method, request.path);
    }

    [[nodiscard]] HttpResponse ListBucketsLocked() const
    {
        std::string xml = std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><ListAllMyBucketsResult "
                                      "xmlns=\"{}\"><Owner><ID>fake</ID><DisplayName>fake</DisplayName></Owner><Buckets>",
                                      kS3Xmlns);
        for (const std::string& bucket : _buckets)
        {
            xml += std::format("<Bucket><Name>{}</Name><CreationDate>{}</CreationDate></Bucket>", XmlEscape(bucket), Iso8601(NowSeconds()));
        }
        xml += "</Buckets></ListAllMyBucketsResult>";
        return MakeXmlResponse(200, std::move(xml));
    }

    [[nodiscard]] HttpResponse DispatchBucketLocked(const HttpRequest& request, const std::string& bucket)
    {
        if (request.method == "PUT")
        {
            _buckets.insert(bucket);
            return MakeEmptyResponse(200);
        }
        if (! _buckets.contains(bucket))
        {
            return request.method == "HEAD" ? MakeEmptyResponse(404)
                                            : MakeErrorResponse(404, "NoSuchBucket", "The specified bucket does not exist", request.path);
        }
        if (request.method == "HEAD")
        {
            return MakeEmptyResponse(200);
        }
        if (request.method == "GET")
        {
            if (request.HasQuery("location"))
            {
                return MakeXmlResponse(
                    200, std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><LocationConstraint xmlns=\"{}\">us-east-1</LocationConstraint>", kS3Xmlns));
            }
            return ListObjectsLocked(request, bucket);
        }
        if (request.method == "DELETE")
        {
            _buckets.erase(bucket);
            return MakeEmptyResponse(204);
        }
        return MakeErrorResponse(405, "MethodNotAllowed", request.method, request.path);
    }

    [[nodiscard]] HttpResponse ListObjectsLocked(const HttpRequest& request, const std::string& bucket) const
    {
        const std::string prefix    = request.Query("prefix");
        const std::string delimiter = request.Query("delimiter");
        std::string startAfter      = request.Query("start-after");
        if (request.HasQuery("continuation-token"))
        {
            startAfter = request.Query("continuation-token");
        }
        size_t maxKeys = 1000u;
        if (const std::string maxKeysText = request.Query("max-keys"); ! maxKeysText.empty())
        {
            static_cast<void>(std::from_chars(maxKeysText.data(), maxKeysText.data() + maxKeysText.size(), maxKeys));
            maxKeys = std::clamp<size_t>(maxKeys, 1u, 1000u);
        }

        const std::string idPrefix = bucket + "/";
        std::string contents;
        std::string commonPrefixes;
        std::set<std::string> emittedPrefixes;
        size_t emitted = 0u;
        bool truncated = false;
        std::string lastIncludedKey;
        for (auto it = _objects.lower_bound(idPrefix); it != _objects.end() && it->first.starts_with(idPrefix); ++it)
        {
            const std::string_view key(std::string_view(it->first).substr(idPrefix.size()));
            if (! key.starts_with(prefix) || (! startAfter.empty() && key <= startAfter))
            {
                continue;
            }
            const std::string_view tail = key.substr(prefix.size());
            if (! delimiter.empty())
            {
                if (const size_t cut = tail.find(delimiter); cut != std::string_view::npos)
                {
                    const std::string common = prefix + std::string(tail.substr(0, cut + delimiter.size()));
                    if (emittedPrefixes.contains(common))
                    {
                        lastIncludedKey = std::string(key);
                        continue;
                    }
                    if (emitted >= maxKeys)
                    {
                        truncated = true;
                        break;
                    }
                    emittedPrefixes.insert(common);
                    commonPrefixes += std::format("<CommonPrefixes><Prefix>{}</Prefix></CommonPrefixes>", XmlEscape(common));
                    ++emitted;
                    lastIncludedKey = std::string(key);
                    continue;
                }
            }
            if (emitted >= maxKeys)
            {
                truncated = true;
                break;
            }
            contents += std::format(
                "<Contents><Key>{}</Key><LastModified>{}</LastModified><ETag>{}</ETag><Size>{}</Size><StorageClass>STANDARD</StorageClass></Contents>",
                XmlEscape(key),
                Iso8601(it->second.modified),
                XmlEscape(it->second.etag),
                it->second.bytes.size());
            ++emitted;
            lastIncludedKey = std::string(key);
        }

        std::string xml = std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><ListBucketResult "
                                      "xmlns=\"{}\"><Name>{}</Name><Prefix>{}</Prefix><KeyCount>{}</KeyCount><MaxKeys>{}</MaxKeys>",
                                      kS3Xmlns,
                                      XmlEscape(bucket),
                                      XmlEscape(prefix),
                                      emitted,
                                      maxKeys);
        if (! delimiter.empty())
        {
            xml += std::format("<Delimiter>{}</Delimiter>", XmlEscape(delimiter));
        }
        if (! startAfter.empty())
        {
            xml += std::format("<StartAfter>{}</StartAfter>", XmlEscape(startAfter));
        }
        xml += std::format("<IsTruncated>{}</IsTruncated>", truncated ? "true" : "false");
        if (truncated)
        {
            xml += std::format("<NextContinuationToken>{}</NextContinuationToken>", XmlEscape(lastIncludedKey));
        }
        xml += contents;
        xml += commonPrefixes;
        xml += "</ListBucketResult>";
        HttpResponse response = MakeXmlResponse(200, std::move(xml));
        response.dripBody     = true;
        return response;
    }

    static void AddObjectHeaders(HttpResponse& response, const StoredObject& object)
    {
        response.headers.emplace_back("ETag", object.etag);
        response.headers.emplace_back("x-amz-checksum-crc64nvme", object.crc64NvmeBase64);
        response.headers.emplace_back("x-amz-checksum-type", "FULL_OBJECT");
        response.headers.emplace_back("Last-Modified", Rfc1123(object.modified));
        response.headers.emplace_back("Accept-Ranges", "bytes");
        response.headers.emplace_back("Content-Type", "application/octet-stream");
    }

    // Returns nullptr when the request's If-Match / If-None-Match preconditions hold.
    [[nodiscard]] static std::optional<HttpResponse> CheckPreconditions(const HttpRequest& request,
                                                                        const StoredObject* existing,
                                                                        std::string_view ifMatchName,
                                                                        std::string_view ifNoneMatchName)
    {
        if (const std::string* ifMatch = request.Header(ifMatchName); ifMatch != nullptr)
        {
            const std::string wanted = StripQuotes(*ifMatch);
            if (existing == nullptr || (wanted != "*" && StripQuotes(existing->etag) != wanted))
            {
                return MakeErrorResponse(412, "PreconditionFailed", "At least one of the pre-conditions you specified did not hold", request.path);
            }
        }
        if (const std::string* ifNoneMatch = request.Header(ifNoneMatchName); ifNoneMatch != nullptr)
        {
            const std::string unwanted = StripQuotes(*ifNoneMatch);
            if (existing != nullptr && (unwanted == "*" || StripQuotes(existing->etag) == unwanted))
            {
                return MakeErrorResponse(412, "PreconditionFailed", "At least one of the pre-conditions you specified did not hold", request.path);
            }
        }
        return std::nullopt;
    }

    [[nodiscard]] HttpResponse GetObjectLocked(const HttpRequest& request, const std::string& bucket, const std::string& key) const
    {
        const auto it = _objects.find(ObjectId(bucket, key));
        if (it == _objects.end())
        {
            return MakeErrorResponse(404, "NoSuchKey", "The specified key does not exist.", request.path);
        }
        const StoredObject& object = it->second;
        if (auto failed = CheckPreconditions(request, &object, "if-match", "if-none-match"); failed.has_value())
        {
            return std::move(failed.value());
        }
        HttpResponse response;
        const size_t size = object.bytes.size();
        size_t first      = 0u;
        size_t last       = size == 0u ? 0u : size - 1u;
        bool ranged       = false;
        if (const std::string* range = request.Header("range"); range != nullptr && range->starts_with("bytes="))
        {
            const std::string_view spec(std::string_view(*range).substr(6u));
            const size_t dash = spec.find('-');
            if (dash == std::string_view::npos)
            {
                return MakeErrorResponse(416, "InvalidRange", "The requested range is not satisfiable", request.path);
            }
            const std::string_view firstText = spec.substr(0, dash);
            const std::string_view lastText  = spec.substr(dash + 1u);
            if (firstText.empty())
            {
                size_t suffix = 0u;
                static_cast<void>(std::from_chars(lastText.data(), lastText.data() + lastText.size(), suffix));
                first = suffix >= size ? 0u : size - suffix;
            }
            else
            {
                static_cast<void>(std::from_chars(firstText.data(), firstText.data() + firstText.size(), first));
                if (! lastText.empty())
                {
                    static_cast<void>(std::from_chars(lastText.data(), lastText.data() + lastText.size(), last));
                }
            }
            if (size == 0u || first >= size)
            {
                HttpResponse unsatisfiable = MakeErrorResponse(416, "InvalidRange", "The requested range is not satisfiable", request.path);
                unsatisfiable.headers.emplace_back("Content-Range", std::format("bytes */{}", size));
                return unsatisfiable;
            }
            last   = (std::min)(last, size - 1u);
            ranged = true;
        }
        response.status = ranged ? 206 : 200;
        AddObjectHeaders(response, object);
        if (ranged)
        {
            response.headers.emplace_back("Content-Range", std::format("bytes {}-{}/{}", first, last, size));
            response.body.assign(object.bytes.begin() + static_cast<ptrdiff_t>(first), object.bytes.begin() + static_cast<ptrdiff_t>(last + 1u));
        }
        else
        {
            response.body = object.bytes;
        }
        return response;
    }

    [[nodiscard]] HttpResponse HeadObjectLocked(const HttpRequest& request, const std::string& bucket, const std::string& key) const
    {
        const auto it = _objects.find(ObjectId(bucket, key));
        if (it == _objects.end())
        {
            return MakeEmptyResponse(404);
        }
        if (auto failed = CheckPreconditions(request, &it->second, "if-match", "if-none-match"); failed.has_value())
        {
            failed->body.clear();
            return std::move(failed.value());
        }
        HttpResponse response;
        AddObjectHeaders(response, it->second);
        response.body = it->second.bytes; // Content-Length only; HEAD sends no body
        return response;
    }

    [[nodiscard]] const StoredObject* ResolveCopySourceLocked(const HttpRequest& request, HttpResponse& failure) const
    {
        const std::string* copySource = request.Header("x-amz-copy-source");
        if (copySource == nullptr)
        {
            return nullptr;
        }
        std::string source = PercentDecode(*copySource);
        if (const size_t versionCut = source.find('?'); versionCut != std::string::npos)
        {
            source.erase(versionCut);
        }
        if (! source.empty() && source.front() == '/')
        {
            source.erase(0, 1u);
        }
        const auto it = _objects.find(source);
        if (it == _objects.end())
        {
            failure = MakeErrorResponse(404, "NoSuchKey", "The specified key does not exist.", source);
            return nullptr;
        }
        if (auto failed = CheckPreconditions(request, &it->second, "x-amz-copy-source-if-match", "x-amz-copy-source-if-none-match"); failed.has_value())
        {
            failure = std::move(failed.value());
            return nullptr;
        }
        return &it->second;
    }

    [[nodiscard]] HttpResponse PutLocked(const HttpRequest& request, const std::string& bucket, const std::string& key)
    {
        const bool partUpload = request.HasQuery("partNumber") && request.HasQuery("uploadId");
        if (partUpload)
        {
            const auto upload = _uploads.find(request.Query("uploadId"));
            if (upload == _uploads.end())
            {
                return MakeErrorResponse(404, "NoSuchUpload", "The specified upload does not exist.", request.path);
            }
            unsigned int partNumber    = 0u;
            const std::string partText = request.Query("partNumber");
            static_cast<void>(std::from_chars(partText.data(), partText.data() + partText.size(), partNumber));
            std::vector<uint8_t> partBytes;
            if (request.Header("x-amz-copy-source") != nullptr)
            {
                HttpResponse failure;
                const StoredObject* source = ResolveCopySourceLocked(request, failure);
                if (source == nullptr)
                {
                    return failure;
                }
                size_t first = 0u;
                size_t last  = source->bytes.empty() ? 0u : source->bytes.size() - 1u;
                if (const std::string* range = request.Header("x-amz-copy-source-range"); range != nullptr && range->starts_with("bytes="))
                {
                    const std::string_view spec(std::string_view(*range).substr(6u));
                    const size_t dash = spec.find('-');
                    static_cast<void>(std::from_chars(spec.data(), spec.data() + dash, first));
                    static_cast<void>(std::from_chars(spec.data() + dash + 1u, spec.data() + spec.size(), last));
                }
                if (first >= source->bytes.size())
                {
                    return MakeErrorResponse(416, "InvalidRange", "The requested range is not satisfiable", request.path);
                }
                last = (std::min)(last, source->bytes.size() - 1u);
                partBytes.assign(source->bytes.begin() + static_cast<ptrdiff_t>(first), source->bytes.begin() + static_cast<ptrdiff_t>(last + 1u));
                const std::string etag = ComputeEtag(partBytes);
                upload->second.parts.insert_or_assign(partNumber, std::move(partBytes));
                HttpResponse response = MakeXmlResponse(
                    200,
                    std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><CopyPartResult><LastModified>{}</LastModified><ETag>{}</ETag></CopyPartResult>",
                                Iso8601(NowSeconds()),
                                XmlEscape(etag)));
                return response;
            }
            partBytes              = request.body;
            const std::string etag = ComputeEtag(partBytes);
            upload->second.parts.insert_or_assign(partNumber, std::move(partBytes));
            HttpResponse response = MakeEmptyResponse(200);
            response.headers.emplace_back("ETag", etag);
            return response;
        }

        const auto existing         = _objects.find(ObjectId(bucket, key));
        const StoredObject* current = existing == _objects.end() ? nullptr : &existing->second;
        if (auto failed = CheckPreconditions(request, current, "if-match", "if-none-match"); failed.has_value())
        {
            return std::move(failed.value());
        }
        if (request.Header("x-amz-copy-source") != nullptr)
        {
            HttpResponse failure;
            const StoredObject* source = ResolveCopySourceLocked(request, failure);
            if (source == nullptr)
            {
                return failure;
            }
            std::vector<uint8_t> copy = source->bytes;
            StoreLocked(bucket, key, std::move(copy));
            const StoredObject& stored = _objects.at(ObjectId(bucket, key));
            return MakeXmlResponse(
                200,
                std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><CopyObjectResult><LastModified>{}</LastModified><ETag>{}</ETag></CopyObjectResult>",
                            Iso8601(stored.modified),
                            XmlEscape(stored.etag)));
        }
        StoreLocked(bucket, key, request.body);
        HttpResponse response = MakeEmptyResponse(200);
        AddObjectHeaders(response, _objects.at(ObjectId(bucket, key)));
        return response;
    }

    [[nodiscard]] HttpResponse DeleteLocked(const HttpRequest& request, const std::string& bucket, const std::string& key)
    {
        if (request.HasQuery("uploadId"))
        {
            _uploads.erase(request.Query("uploadId"));
            return MakeEmptyResponse(204);
        }
        const auto existing = _objects.find(ObjectId(bucket, key));
        if (existing != _objects.end())
        {
            if (auto failed = CheckPreconditions(request, &existing->second, "if-match", "if-none-match"); failed.has_value())
            {
                return std::move(failed.value());
            }
            _objects.erase(existing);
        }
        else if (request.Header("if-match") != nullptr)
        {
            return MakeErrorResponse(412, "PreconditionFailed", "At least one of the pre-conditions you specified did not hold", request.path);
        }
        return MakeEmptyResponse(204);
    }

    [[nodiscard]] HttpResponse PostLocked(const HttpRequest& request, const std::string& bucket, const std::string& key)
    {
        if (request.HasQuery("uploads"))
        {
            const std::string uploadId = std::format("{:016x}", ++_uploadSerial);
            _uploads.insert_or_assign(uploadId, PendingUpload{.bucket = bucket, .key = key, .parts = {}});
            return MakeXmlResponse(200,
                                   std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><InitiateMultipartUploadResult "
                                               "xmlns=\"{}\"><Bucket>{}</Bucket><Key>{}</Key><UploadId>{}</UploadId></InitiateMultipartUploadResult>",
                                               kS3Xmlns,
                                               XmlEscape(bucket),
                                               XmlEscape(key),
                                               uploadId));
        }
        if (request.HasQuery("uploadId"))
        {
            const auto upload = _uploads.find(request.Query("uploadId"));
            if (upload == _uploads.end())
            {
                return MakeErrorResponse(404, "NoSuchUpload", "The specified upload does not exist.", request.path);
            }
            const auto existing         = _objects.find(ObjectId(bucket, key));
            const StoredObject* current = existing == _objects.end() ? nullptr : &existing->second;
            if (auto failed = CheckPreconditions(request, current, "if-match", "if-none-match"); failed.has_value())
            {
                return std::move(failed.value());
            }
            std::vector<uint8_t> assembled;
            const std::string_view body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
            size_t search = 0u;
            while (true)
            {
                const size_t open = body.find("<PartNumber>", search);
                if (open == std::string_view::npos)
                {
                    break;
                }
                const size_t valueStart = open + 12u;
                const size_t close      = body.find("</PartNumber>", valueStart);
                if (close == std::string_view::npos)
                {
                    break;
                }
                unsigned int partNumber = 0u;
                static_cast<void>(std::from_chars(body.data() + valueStart, body.data() + close, partNumber));
                const auto part = upload->second.parts.find(partNumber);
                if (part == upload->second.parts.end())
                {
                    return MakeErrorResponse(400, "InvalidPart", "One or more of the specified parts could not be found.", request.path);
                }
                assembled.insert(assembled.end(), part->second.begin(), part->second.end());
                search = close;
            }
            _uploads.erase(upload);
            StoreLocked(bucket, key, std::move(assembled));
            const StoredObject& stored = _objects.at(ObjectId(bucket, key));
            return MakeXmlResponse(
                200,
                std::format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><CompleteMultipartUploadResult "
                            "xmlns=\"{}\"><Location>http://127.0.0.1:{}/{}/{}</Location><Bucket>{}</Bucket><Key>{}</Key><ETag>{}</ETag><ChecksumCRC64NVME>{}</"
                            "ChecksumCRC64NVME><ChecksumType>FULL_OBJECT</ChecksumType></CompleteMultipartUploadResult>",
                            kS3Xmlns,
                            _port,
                            XmlEscape(bucket),
                            XmlEscape(key),
                            XmlEscape(bucket),
                            XmlEscape(key),
                            XmlEscape(stored.etag),
                            XmlEscape(stored.crc64NvmeBase64)));
        }
        return MakeErrorResponse(405, "MethodNotAllowed", request.method, request.path);
    }

    unsigned short _port = 0u;
    wil::unique_socket _listener;
    std::jthread _thread;
    std::mutex _clientsMutex;
    std::vector<std::jthread> _clients;
    std::atomic<HRESULT> _serverHr{S_OK};
    std::atomic<uint64_t> _requestSerial{0u};

    mutable std::mutex _stateMutex;
    std::set<std::string> _buckets;
    std::map<std::string, StoredObject> _objects;
    std::map<std::string, PendingUpload> _uploads;
    std::vector<RequestRecord> _requests;
    std::string _stallMethod;
    unsigned int _stallMs    = 0u;
    size_t _dripBytesPerTick = 0u;
    unsigned int _dripTickMs = 0u;
    uint64_t _uploadSerial   = 0u;
};

class CancelControl final : public IFileSystemOperationControl
{
public:
    CancelControl() noexcept                       = default;
    CancelControl(const CancelControl&)            = delete;
    CancelControl& operator=(const CancelControl&) = delete;
    CancelControl(CancelControl&&)                 = delete;
    CancelControl& operator=(CancelControl&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (abort == nullptr)
        {
            return E_POINTER;
        }
        *abort = abortRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        if (mode == nullptr)
        {
            return E_POINTER;
        }
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }
    std::atomic<bool> abortRequested{false};
};

[[nodiscard]] std::string FixtureConfiguration(unsigned short port, unsigned int connectTimeoutMs, unsigned int requestTimeoutMs)
{
    return std::format(
        R"({{"defaultRegion":"us-east-1","defaultEndpointOverride":"http://127.0.0.1:{}","useHttps":false,"verifyTls":false,"useVirtualAddressing":false,"anonymous":true,"connectTimeoutMs":{},"requestTimeoutMs":{}}})",
        port,
        connectTimeoutMs,
        requestTimeoutMs);
}
} // namespace

// R0f-S3 witness. (1) Bytes still flowing: a recursive Delete whose listing the fixture drips
// returns ERROR_CANCELLED within a few body callbacks of Cancel. (2) A server that stopped
// answering: Cancel is reported as the host's verdict as soon as the provider-owned bound returns
// the request, and (3) without Cancel the same bound returns it on its own as a transport failure.
void RunS3StalledRequestCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the S3 stalled-request proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        FakeS3Endpoint endpoint;
        endpoint.SeedObject("r0f", "stall.bin", std::vector<uint8_t>(64u, 0x5Au));
        for (unsigned int index = 0u; index < 48u; ++index)
        {
            endpoint.SeedObject("r0f", std::format("folder/object-{:03}.bin", index), std::vector<uint8_t>(16u, static_cast<uint8_t>(index)));
        }
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"S3 stalled-request fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<FileSystemS3> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem) && SUCCEEDED(fileSystem->InitializationStatus()),
                         L"S3 stalled-request proof should create an S3 instance",
                         passed,
                         failed))
        {
            return;
        }
        constexpr unsigned int kConnectTimeoutMs = 2'000u;
        constexpr unsigned int kRequestTimeoutMs = 2'000u;
        HRESULT hr = fileSystem->SetConfiguration(FixtureConfiguration(endpoint.Port(), kConnectTimeoutMs, kRequestTimeoutMs).c_str());
        if (! DebugCheck(SUCCEEDED(hr), L"S3 instance should accept the fixture configuration", passed, failed))
        {
            return;
        }
        const unsigned long boundMs = S3ProviderWatchdogTimeoutMs(kConnectTimeoutMs, kRequestTimeoutMs);
        DebugCheck(boundMs > 0u && boundMs <= 20'000u, L"the S3 provider-owned bound must be a small nonzero value for the fixture timeouts", passed, failed);

        const auto deleteWithControl = [&](const wchar_t* path, FileSystemFlags flags, CancelControl* control, std::atomic<HRESULT>& result) noexcept
        {
            const wchar_t* paths[] = {path};
            FileSystemOptions options{};
            options.sizeBytes        = sizeof(options);
            options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
            options.operationControl = control;
            result.store(fileSystem->DeleteItems(paths, 1u, flags, control != nullptr ? &options : nullptr, nullptr, nullptr), std::memory_order_release);
        };

        // 1) Bytes flowing: the listing drips; Cancel lands between body callbacks.
        endpoint.SetDripListing(96u, 40u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { deleteWithControl(L"/r0f/folder", FILESYSTEM_FLAG_RECURSIVE, &control, deleteHr); });
            const ULONGLONG waitStart = GetTickCount64();
            while (endpoint.RequestCount("GET") == 0u && GetTickCount64() - waitStart < 10'000u)
            {
                Sleep(20u);
            }
            const bool listingReached = endpoint.RequestCount("GET") != 0u;
            Sleep(250u);
            const ULONGLONG cancelTick = GetTickCount64();
            control.abortRequested.store(true, std::memory_order_release);
            const bool returned      = WaitForSingleObject(deleter.native_handle(), 15'000u) == WAIT_OBJECT_0;
            const ULONGLONG cancelMs = GetTickCount64() - cancelTick;
            if (! returned)
            {
                endpoint.Stop();
            }
            deleter.join();
            DebugCheck(listingReached, L"the recursive Delete must list the prefix on the fixture", passed, failed);
            DebugCheck(returned, L"a request whose body is still arriving must return after Cancel", passed, failed);
            DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"a canceled streaming request must report ERROR_CANCELLED",
                       passed,
                       failed);
            std::fwprintf(stderr, L"[S3] streaming request returned %llu ms after Cancel\n", static_cast<unsigned long long>(cancelMs));
            DebugCheck(cancelMs < 3'000u, L"Cancel must return a streaming request within a few body callbacks", passed, failed);
            DebugCheck(endpoint.RequestCount("DELETE") == 0u, L"a Delete canceled while listing must not delete anything", passed, failed);
            if (! returned)
            {
                return;
            }
        }
        endpoint.SetDripListing(0u, 0u);

        // 2) Silent server + Cancel: the provider-owned bound returns the request; the verdict is the host's.
        endpoint.SetStallMethod("DELETE", 60'000u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { deleteWithControl(L"/r0f/stall.bin", FILESYSTEM_FLAG_NONE, &control, deleteHr); });
            const ULONGLONG waitStart = GetTickCount64();
            while (endpoint.RequestCount("DELETE") == 0u && GetTickCount64() - waitStart < 10'000u)
            {
                Sleep(20u);
            }
            const bool deleteReached = endpoint.RequestCount("DELETE") != 0u;
            Sleep(200u);
            const ULONGLONG cancelTick = GetTickCount64();
            control.abortRequested.store(true, std::memory_order_release);
            const bool returned      = WaitForSingleObject(deleter.native_handle(), boundMs + 10'000u) == WAIT_OBJECT_0;
            const ULONGLONG cancelMs = GetTickCount64() - cancelTick;
            if (! returned)
            {
                endpoint.Stop();
            }
            deleter.join();
            DebugCheck(deleteReached, L"the stalled DELETE must reach the fixture", passed, failed);
            DebugCheck(returned, L"a request the server never answers must return after Cancel", passed, failed);
            DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"a canceled stalled request must report ERROR_CANCELLED",
                       passed,
                       failed);
            std::fwprintf(
                stderr, L"[S3] stalled request returned %llu ms after Cancel (declared bound %lu ms)\n", static_cast<unsigned long long>(cancelMs), boundMs);
            DebugCheck(cancelMs <= boundMs + 5'000u, L"Cancel on a silent server must return within the declared provider-owned bound", passed, failed);
            if (! returned)
            {
                return;
            }
        }

        // 4) Ranged reader on a GET the server never answers: the host hands its control through
        //    IFileReaderOperationControl; the reader arms the request on its own thread, so Cancel
        //    ends the in-flight GetObject within the declared bound.
        endpoint.SetStallMethod("GET", 60'000u);
        {
            wil::com_ptr<IFileReader> reader;
            hr = fileSystem->CreateFileReader(L"/r0f/stall.bin", reader.put());
            if (DebugCheck(SUCCEEDED(hr) && static_cast<bool>(reader), L"the fixture object must open a ranged reader", passed, failed))
            {
                wil::com_ptr<IFileReaderOperationControl> readerControl;
                DebugCheck(SUCCEEDED(reader->QueryInterface(__uuidof(IFileReaderOperationControl), readerControl.put_void())) && readerControl,
                           L"the S3 ranged reader must expose IFileReaderOperationControl (R0f-S3 C8)",
                           passed,
                           failed);
                CancelControl control;
                FileSystemOptions readerOptions{};
                readerOptions.sizeBytes        = sizeof(readerOptions);
                readerOptions.operationControl = &control;
                if (readerControl)
                {
                    DebugCheck(
                        SUCCEEDED(readerControl->SetOperationControl(&readerOptions)), L"the ranged reader must accept the host control", passed, failed);
                }
                std::atomic<HRESULT> readHr{E_PENDING};
                std::thread readerThread([&]() noexcept
                {
                    std::vector<uint8_t> buffer(32u, 0u);
                    unsigned long got = 0u;
                    readHr.store(reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &got), std::memory_order_release);
                });
                const ULONGLONG waitStart = GetTickCount64();
                while (endpoint.RequestCount("GET") == 0u && GetTickCount64() - waitStart < 10'000u)
                {
                    Sleep(20u);
                }
                const bool getReached = endpoint.RequestCount("GET") != 0u;
                Sleep(200u);
                const ULONGLONG cancelTick = GetTickCount64();
                control.abortRequested.store(true, std::memory_order_release);
                const bool returned      = WaitForSingleObject(readerThread.native_handle(), boundMs + 10'000u) == WAIT_OBJECT_0;
                const ULONGLONG cancelMs = GetTickCount64() - cancelTick;
                if (! returned)
                {
                    endpoint.Stop();
                }
                readerThread.join();
                DebugCheck(getReached, L"the ranged reader must issue its GET to the fixture", passed, failed);
                DebugCheck(returned, L"a reader GET the server never answers must return after Cancel", passed, failed);
                DebugCheck(readHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                           L"a canceled stalled reader GET must report ERROR_CANCELLED",
                           passed,
                           failed);
                std::fwprintf(stderr,
                              L"[S3] stalled reader GET returned %llu ms after Cancel (declared bound %lu ms)\n",
                              static_cast<unsigned long long>(cancelMs),
                              boundMs);
                DebugCheck(cancelMs <= boundMs + 5'000u, L"Cancel on a silent reader GET must return within the declared provider-owned bound", passed, failed);
                if (! returned)
                {
                    return;
                }
            }
        }
        endpoint.SetStallMethod("DELETE", 60'000u);
        // 3) Silent server, no Cancel: the bound returns it on its own as a transport failure.
        {
            std::atomic<HRESULT> deleteHr{E_PENDING};
            const ULONGLONG boundStart = GetTickCount64();
            deleteWithControl(L"/r0f/stall.bin", FILESYSTEM_FLAG_NONE, nullptr, deleteHr);
            const ULONGLONG elapsedMs = GetTickCount64() - boundStart;
            const HRESULT boundHr     = deleteHr.load(std::memory_order_acquire);
            DebugCheck(FAILED(boundHr) && boundHr != HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"an un-canceled stalled request must fail through the transport bound, not as a cancel",
                       passed,
                       failed);
            std::fwprintf(stderr,
                          L"[S3] stalled request returned on its own after %llu ms (declared bound %lu ms)\n",
                          static_cast<unsigned long long>(elapsedMs),
                          boundMs);
            DebugCheck(elapsedMs <= boundMs + 5'000u, L"the provider-owned bound must return a stalled request on its own", passed, failed);
        }
        endpoint.SetStallMethod({}, 0u);
        DebugCheck(endpoint.HasObject("r0f", "stall.bin"), L"a stalled DELETE the client gave up on must not be reported as deleted", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemS3 stalled-request selftest failed after std::exception.");
        DebugCheck(false, L"S3 stalled-request proof should not throw std::exception", passed, failed);
    }
}

// Host self-tests drive the fake S3 endpoint through these exports: start one on a loopback port
// (bucket `r0f-bucket` exists), configure an S3 instance with `defaultEndpointOverride`
// `http://127.0.0.1:<port>`, path-style addressing and `anonymous`, run File Operations against
// `/r0f-bucket/...`, then stop it.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderS3StartFakeS3ForSelfTest(unsigned int* port, void** endpoint) noexcept
{
    if (port == nullptr || endpoint == nullptr)
    {
        return E_POINTER;
    }
    *port     = 0u;
    *endpoint = nullptr;
    WSADATA winsockData{};
    if (WSAStartup(MAKEWORD(2, 2), &winsockData) != 0)
    {
        return E_FAIL;
    }
    try
    {
        auto fixture = std::make_unique<FakeS3Endpoint>();
        fixture->SeedBucket("r0f-bucket");
        const HRESULT hr = fixture->Start();
        if (FAILED(hr))
        {
            WSACleanup();
            return hr;
        }
        *port     = fixture->Port();
        *endpoint = fixture.release();
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        WSACleanup();
        return E_OUTOFMEMORY;
    }
    catch (const std::exception&)
    {
        WSACleanup();
        return E_FAIL;
    }
}

// Copies the fixture's recent request log into `buffer` for a failing host self-test's message.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderS3FakeS3RequestLogForSelfTest(void* endpoint, wchar_t* buffer, unsigned int capacity) noexcept
{
    if (endpoint == nullptr || buffer == nullptr || capacity == 0u)
    {
        return E_POINTER;
    }
    buffer[0] = L'\0';
    try
    {
        const std::wstring text = static_cast<FakeS3Endpoint*>(endpoint)->RequestLog(48u);
        const size_t count      = (std::min)(text.size(), static_cast<size_t>(capacity - 1u));
        std::copy_n(text.data(), count, buffer);
        buffer[count] = L'\0';
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }
    catch (const std::exception&)
    {
        return E_FAIL;
    }
}

extern "C" __declspec(dllexport) void __stdcall RedSalamanderS3StopFakeS3ForSelfTest(void* endpoint) noexcept
{
    auto* fixture = static_cast<FakeS3Endpoint*>(endpoint);
    if (fixture == nullptr)
    {
        return;
    }
    fixture->Stop();
    delete fixture;
    WSACleanup();
}
// R0-RC4: SetExpectedReplacement with a zero host timestamp is refused instead of replacing the
// object the writer observed when it opened.
void RunS3ZeroTimestampReplaceSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the R0-RC4 S3 proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        FakeS3Endpoint endpoint;
        endpoint.SeedObject("r0f", "occupant.bin", std::vector<uint8_t>(64u, 0x5Au));
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"R0-RC4 S3 fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<FileSystemS3> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem) && SUCCEEDED(fileSystem->InitializationStatus()),
                         L"R0-RC4 S3 proof should create an S3 instance",
                         passed,
                         failed))
        {
            return;
        }
        HRESULT hr = fileSystem->SetConfiguration(FixtureConfiguration(endpoint.Port(), 2'000u, 2'000u).c_str());
        if (! DebugCheck(SUCCEEDED(hr), L"R0-RC4 S3 instance should accept the fixture configuration", passed, failed))
        {
            return;
        }

        wil::com_ptr<IFileWriter> writer;
        hr = fileSystem->CreateFileWriter(L"/r0f/occupant.bin", FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.addressof());
        if (! DebugCheck(SUCCEEDED(hr) && writer, L"R0-RC4: S3 should open an overwrite writer on the seeded occupant", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileWriterExpectedReplacement> replacement;
        hr = writer->QueryInterface(IID_PPV_ARGS(replacement.addressof()));
        if (! DebugCheck(SUCCEEDED(hr) && replacement, L"R0-RC4: the S3 writer should expose the expected-replacement contract", passed, failed))
        {
            return;
        }
        FileSystemBasicInformation expected{};
        expected.sizeBytes = sizeof(expected); // lastWriteTime stays 0
        hr                 = replacement->SetExpectedReplacement(&expected);
        if (SUCCEEDED(hr))
        {
            const std::vector<std::byte> bytes(16u, std::byte{0x11});
            unsigned long written = 0u;
            hr                    = writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &written);
            if (SUCCEEDED(hr))
            {
                hr = writer->Commit();
            }
        }
        DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
                   L"R0-RC4: an S3 identity-less replace with a zero host timestamp must be refused before upload",
                   passed,
                   failed);
        DebugCheck(endpoint.HasObject("r0f", "occupant.bin"), L"R0-RC4: the S3 occupant must survive the refused replace", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemS3 zero-timestamp replace selftest failed after std::exception.");
        DebugCheck(false, L"R0-RC4 S3 proof should not throw std::exception", passed, failed);
    }
}

} // namespace FileSystemS3Internal

#else

static_assert(true);

#endif

#pragma once

// Deterministic loopback HTTP/1.1 server for test-enabled builds (R0f fixtures). Include this header
// before any Windows header: Winsock2 must precede windows.h.
//
// The server accepts connections concurrently (one thread per connection, the way real clients use
// keep-alive pools), parses requests (origin or absolute form, Content-Length, chunked and
// aws-chunked bodies, Expect: 100-continue), hands each request to the fixture's handler, records
// every request with the status it answered, and can hold the reply to one method (a server that
// stopped answering) or drip a response body slowly (bytes still flowing).

#include <winsock2.h>
#include <ws2tcpip.h>

// WIL's resource wrappers report implicitly deleted special members (C4625/C4626/C5026/C5027) that
// the project mutes around its own WIL include; this header carries the same policy for its region.
#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <charconv>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <exception>
#include <format>
#include <functional>
#include <limits>
#include <map>
#include <mutex>
#include <stop_token>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

#pragma comment(lib, "Ws2_32.lib")

namespace Common::SelfTest
{
struct LoopbackHttpRequest final
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

struct LoopbackHttpResponse final
{
    int status = 200;
    std::vector<std::pair<std::string, std::string>> headers;
    std::vector<uint8_t> body;
    bool headOnly = false; // send the body's Content-Length but no body (HEAD)
    bool dripBody = false; // eligible for the fixture's drip setting (bytes still flowing)
};

[[nodiscard]] inline std::string_view LoopbackHttpReasonPhrase(int status) noexcept
{
    switch (status)
    {
        case 100: return "Continue";
        case 200: return "OK";
        case 201: return "Created";
        case 202: return "Accepted";
        case 204: return "No Content";
        case 206: return "Partial Content";
        case 304: return "Not Modified";
        case 400: return "Bad Request";
        case 401: return "Unauthorized";
        case 403: return "Forbidden";
        case 404: return "Not Found";
        case 405: return "Method Not Allowed";
        case 409: return "Conflict";
        case 411: return "Length Required";
        case 412: return "Precondition Failed";
        case 416: return "Range Not Satisfiable";
        case 429: return "Too Many Requests";
        case 500: return "Internal Server Error";
        case 503: return "Service Unavailable";
        default: return "Unknown";
    }
}

[[nodiscard]] inline std::string LoopbackHttpToLowerAscii(std::string_view text)
{
    std::string out(text);
    for (char& c : out)
    {
        c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    }
    return out;
}

[[nodiscard]] inline int LoopbackHttpHexValue(char c) noexcept
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

[[nodiscard]] inline std::string LoopbackHttpPercentDecode(std::string_view text)
{
    std::string out;
    out.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i)
    {
        if (text[i] == '%' && i + 2 < text.size() && LoopbackHttpHexValue(text[i + 1]) >= 0 && LoopbackHttpHexValue(text[i + 2]) >= 0)
        {
            out.push_back(static_cast<char>((LoopbackHttpHexValue(text[i + 1]) << 4) | LoopbackHttpHexValue(text[i + 2])));
            i += 2;
        }
        else
        {
            out.push_back(text[i]);
        }
    }
    return out;
}

[[nodiscard]] inline std::wstring LoopbackHttpUtf16FromUtf8(std::string_view text)
{
    if (text.empty())
    {
        return {};
    }
    const int required = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return L"?";
    }
    std::wstring out(static_cast<size_t>(required), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), out.data(), required);
    return out;
}

[[nodiscard]] inline std::chrono::sys_seconds LoopbackHttpNowSeconds() noexcept
{
    return std::chrono::floor<std::chrono::seconds>(std::chrono::system_clock::now());
}

[[nodiscard]] inline std::string LoopbackHttpRfc1123(std::chrono::sys_seconds when)
{
    return std::format("{:%a, %d %b %Y %H:%M:%S} GMT", when);
}

[[nodiscard]] inline std::string LoopbackHttpIso8601(std::chrono::sys_seconds when)
{
    return std::format("{:%Y-%m-%dT%H:%M:%S}Z", when);
}

// Decodes HTTP chunked framing; the same framing (with chunk extensions) is what aws-chunked uses.
[[nodiscard]] inline bool LoopbackHttpTryDecodeChunked(std::string_view encoded, std::vector<uint8_t>& out, size_t& consumed) noexcept
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

class LoopbackHttpServer final
{
public:
    using Handler = std::function<LoopbackHttpResponse(const LoopbackHttpRequest&)>;

    struct RequestRecord final
    {
        std::string method;
        std::string path;
        int status = 0; // 0 until the reply was built (a held reply stays 0 while it is held)
    };

    explicit LoopbackHttpServer(Handler handler) : _handler(std::move(handler))
    {
    }
    ~LoopbackHttpServer()
    {
        Stop();
    }
    LoopbackHttpServer(const LoopbackHttpServer&)            = delete;
    LoopbackHttpServer& operator=(const LoopbackHttpServer&) = delete;
    LoopbackHttpServer(LoopbackHttpServer&&)                 = delete;
    LoopbackHttpServer& operator=(LoopbackHttpServer&&)      = delete;

    [[nodiscard]] unsigned short Port() const noexcept
    {
        return _port;
    }

    // The fixture holds the reply to `method` for `stallMs` (or until it stops): a server that
    // stopped answering. The request is recorded before the hold.
    void SetStallMethod(std::string method, unsigned int stallMs)
    {
        std::scoped_lock lock(_stateMutex);
        _stallMethod = std::move(method);
        _stallMs     = stallMs;
    }

    // Responses marked dripBody are sent `bytesPerTick` at a time with `tickMs` between slices.
    void SetDrip(size_t bytesPerTick, unsigned int tickMs)
    {
        std::scoped_lock lock(_stateMutex);
        _dripBytesPerTick = bytesPerTick;
        _dripTickMs       = tickMs;
    }

    [[nodiscard]] size_t RequestCount(std::string_view method) const
    {
        std::scoped_lock lock(_stateMutex);
        size_t count = 0u;
        for (const RequestRecord& record : _requests)
        {
            if (record.method == method)
            {
                ++count;
            }
        }
        return count;
    }

    // The last `maxEntries` requests as " [METHOD path -> status]" for failure messages.
    [[nodiscard]] std::wstring RequestLog(size_t maxEntries) const
    {
        std::scoped_lock lock(_stateMutex);
        std::wstring text;
        const size_t start = _requests.size() > maxEntries ? _requests.size() - maxEntries : 0u;
        for (size_t index = start; index < _requests.size(); ++index)
        {
            text += std::format(L" [{} {} -> {}]",
                                LoopbackHttpUtf16FromUtf8(_requests[index].method),
                                LoopbackHttpUtf16FromUtf8(_requests[index].path),
                                _requests[index].status);
        }
        return text;
    }

    [[nodiscard]] HRESULT Start()
    {
        WSADATA winsockData{};
        if (WSAStartup(MAKEWORD(2, 2), &winsockData) != 0)
        {
            return E_FAIL;
        }
        _winsockStarted  = true;
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
                OutputDebugStringW(L"LoopbackHttpServer terminated after std::exception.\n");
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
        if (_winsockStarted)
        {
            _winsockStarted = false;
            WSACleanup();
        }
    }

private:
    static constexpr auto kSocketPollInterval    = std::chrono::milliseconds(50);
    static constexpr auto kIdleConnectionTimeout = std::chrono::seconds(30);

    [[nodiscard]] static HRESULT SocketErrorToHResult() noexcept
    {
        const int error = WSAGetLastError();
        return error == 0 ? E_FAIL : HRESULT_FROM_WIN32(static_cast<unsigned long>(error));
    }

    [[nodiscard]] static HRESULT CreateLoopbackListener(wil::unique_socket& listenerOut, unsigned short& portOut) noexcept
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

    [[nodiscard]] static HRESULT WaitForReadable(SOCKET socketValue, std::stop_token stopToken, std::chrono::steady_clock::time_point deadline) noexcept
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

    [[nodiscard]] static HRESULT SendAll(SOCKET socketValue, const void* data, size_t size) noexcept
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

    [[nodiscard]] static HRESULT SendAll(SOCKET socketValue, std::string_view text) noexcept
    {
        return SendAll(socketValue, text.data(), text.size());
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

    [[nodiscard]] static bool ParseRequestHead(std::string_view head, LoopbackHttpRequest& request)
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
        // Some clients send absolute form (`PUT http://host:port/path HTTP/1.1`).
        std::string_view relativeTarget = target;
        if (const size_t scheme = relativeTarget.find("://"); scheme != std::string_view::npos && scheme < 8u)
        {
            const size_t pathStart = relativeTarget.find('/', scheme + 3u);
            relativeTarget         = pathStart == std::string_view::npos ? std::string_view("/") : relativeTarget.substr(pathStart);
        }
        const size_t queryStart = relativeTarget.find('?');
        request.path            = LoopbackHttpPercentDecode(relativeTarget.substr(0, queryStart));
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
                request.query.insert_or_assign(LoopbackHttpPercentDecode(pair.substr(0, eq)),
                                               eq == std::string_view::npos ? std::string{} : LoopbackHttpPercentDecode(pair.substr(eq + 1u)));
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
                request.headers.insert_or_assign(LoopbackHttpToLowerAscii(line.substr(0, colon)), std::string(value));
            }
            if (end == std::string_view::npos)
            {
                break;
            }
            index = end + 2u;
        }
        if (const std::string* connection = request.Header("connection"); connection != nullptr && LoopbackHttpToLowerAscii(*connection) == "close")
        {
            request.closeAfter = true;
        }
        return true;
    }

    [[nodiscard]] static HRESULT ReadRequest(SOCKET socketValue, std::stop_token stopToken, std::string& buffer, LoopbackHttpRequest& request)
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
        const bool chunked                  = transferEncoding != nullptr && LoopbackHttpToLowerAscii(*transferEncoding).find("chunked") != std::string::npos;
        size_t contentLength                = 0u;
        if (const std::string* lengthHeader = request.Header("content-length"); lengthHeader != nullptr)
        {
            const auto parsed = std::from_chars(lengthHeader->data(), lengthHeader->data() + lengthHeader->size(), contentLength);
            if (parsed.ec != std::errc{})
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        if (const std::string* expect = request.Header("expect"); expect != nullptr && LoopbackHttpToLowerAscii(*expect) == "100-continue")
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
            while (! LoopbackHttpTryDecodeChunked(buffer, request.body, consumed))
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
            if (contentEncoding != nullptr && LoopbackHttpToLowerAscii(*contentEncoding).find("aws-chunked") != std::string::npos)
            {
                std::vector<uint8_t> decoded;
                size_t consumed = 0u;
                const std::string_view framed(reinterpret_cast<const char*>(request.body.data()), request.body.size());
                if (! LoopbackHttpTryDecodeChunked(framed, decoded, consumed))
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                request.body = std::move(decoded);
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT SendResponse(
        SOCKET socketValue, std::stop_token stopToken, const LoopbackHttpResponse& response, size_t dripBytesPerTick, unsigned int dripTickMs)
    {
        std::string head = std::format("HTTP/1.1 {} {}\r\nDate: {}\r\nServer: RedSalamanderLoopbackFixture\r\nx-request-id: {}\r\nContent-Length: {}\r\n",
                                       response.status,
                                       LoopbackHttpReasonPhrase(response.status),
                                       LoopbackHttpRfc1123(LoopbackHttpNowSeconds()),
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
            LoopbackHttpRequest request;
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

            LoopbackHttpResponse response = _handler(request);
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

    Handler _handler;
    unsigned short _port = 0u;
    bool _winsockStarted = false;
    wil::unique_socket _listener;
    std::jthread _thread;
    std::mutex _clientsMutex;
    std::vector<std::jthread> _clients;
    std::atomic<HRESULT> _serverHr{S_OK};
    std::atomic<uint64_t> _requestSerial{0u};

    mutable std::mutex _stateMutex;
    std::vector<RequestRecord> _requests;
    std::string _stallMethod;
    unsigned int _stallMs    = 0u;
    size_t _dripBytesPerTick = 0u;
    unsigned int _dripTickMs = 0u;
};
} // namespace Common::SelfTest

#pragma warning(pop)

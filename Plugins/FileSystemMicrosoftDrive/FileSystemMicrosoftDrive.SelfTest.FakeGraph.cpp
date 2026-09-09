#if defined(ENABLE_TESTS)

// The loopback HTTP fixture must precede the Windows headers the plugin header pulls in.
#include "ContentDigest.h"
#include "LoopbackHttpFixture.h"

#include "FileSystemMicrosoftDrive.h"

#include <algorithm>
#include <atomic>
#include <cctype>
#include <charconv>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <format>
#include <map>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

// R0f-Graph fixture: a deterministic loopback Microsoft Graph drive (path-style `/v1.0/drives/{id}`
// REST over plain HTTP) that keeps items in memory: item metadata by path and by id, children with
// `$top` paging, folder creation, PATCH rename/move with If-Match, DELETE, simple content upload,
// upload sessions with Content-Range chunks, and ranged content download. Host self-tests start it
// through the exports at the end of this file; the debug self-test proves the provider-owned bound
// and Cancel on it. The plugin is pointed at the fixture through `ConfigureFakeGraph`.
namespace FileSystemMicrosoftDriveSelfTest
{
namespace
{
using Common::SelfTest::LoopbackHttpRequest;
using Common::SelfTest::LoopbackHttpResponse;
using Common::SelfTest::LoopbackHttpServer;

constexpr Common::DebugSelfTest::Check DebugCheck{L"Microsoft Drive"};
constexpr std::string_view kDriveId = "drive-selftest";

[[nodiscard]] bool EqualsNoCaseAscii(std::string_view left, std::string_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }
    for (size_t index = 0; index < left.size(); ++index)
    {
        if (std::tolower(static_cast<unsigned char>(left[index])) != std::tolower(static_cast<unsigned char>(right[index])))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] std::string JsonQuote(std::string_view text)
{
    std::string out;
    out.reserve(text.size() + 2u);
    out.push_back('"');
    for (const char c : text)
    {
        switch (c)
        {
            case '"': out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            case '\n': out += "\\n"; break;
            case '\r': out += "\\r"; break;
            case '\t': out += "\\t"; break;
            default:
                if (static_cast<unsigned char>(c) < 0x20u)
                {
                    out += std::format("\\u{:04x}", static_cast<unsigned int>(static_cast<unsigned char>(c)));
                }
                else
                {
                    out.push_back(c);
                }
                break;
        }
    }
    out.push_back('"');
    return out;
}

// Minimal JSON string lookup for the fixed request bodies the plugin sends (`name`, `parentReference`
// `id`, `@microsoft.graph.conflictBehavior`); nested keys are found after their parent key.
[[nodiscard]] std::string FindJsonString(std::string_view json, std::string_view key, size_t from = 0u)
{
    const std::string needle = "\"" + std::string(key) + "\"";
    const size_t keyAt       = json.find(needle, from);
    if (keyAt == std::string_view::npos)
    {
        return {};
    }
    size_t index = keyAt + needle.size();
    while (index < json.size() && (json[index] == ' ' || json[index] == ':'))
    {
        ++index;
    }
    if (index >= json.size() || json[index] != '"')
    {
        return {};
    }
    std::string value;
    for (++index; index < json.size() && json[index] != '"'; ++index)
    {
        if (json[index] == '\\' && index + 1 < json.size())
        {
            ++index;
            switch (json[index])
            {
                case 'n': value.push_back('\n'); break;
                case 'r': value.push_back('\r'); break;
                case 't': value.push_back('\t'); break;
                default: value.push_back(json[index]); break;
            }
            continue;
        }
        value.push_back(json[index]);
    }
    return value;
}

[[nodiscard]] std::string StripQuotes(std::string_view text)
{
    if (text.size() >= 2u && text.front() == '"' && text.back() == '"')
    {
        text = text.substr(1, text.size() - 2u);
    }
    return std::string(text);
}

[[nodiscard]] LoopbackHttpResponse JsonResponse(int status, std::string json)
{
    LoopbackHttpResponse response;
    response.status = status;
    response.body.assign(json.begin(), json.end());
    response.headers.emplace_back("Content-Type", "application/json");
    return response;
}

[[nodiscard]] LoopbackHttpResponse ErrorResponse(int status, std::string_view code, std::string_view message)
{
    return JsonResponse(status, std::format(R"({{"error":{{"code":{},"message":{}}}}})", JsonQuote(code), JsonQuote(message)));
}

[[nodiscard]] LoopbackHttpResponse EmptyResponse(int status)
{
    LoopbackHttpResponse response;
    response.status = status;
    return response;
}

struct DriveItem final
{
    std::string id;
    std::string name;
    std::string parentId;
    bool isFolder = false;
    std::vector<uint8_t> bytes;
    unsigned int version = 1u;
    std::chrono::sys_seconds modified{};
};

struct UploadSession final
{
    std::string path; // drive path of the destination
    bool allowOverwrite = false;
    std::vector<uint8_t> bytes;
    uint64_t total = 0u;
};

class FakeGraphDrive final
{
public:
    FakeGraphDrive()
    {
        _items.push_back(DriveItem{
            .id = "root", .name = "", .parentId = "", .isFolder = true, .bytes = {}, .version = 1u, .modified = Common::SelfTest::LoopbackHttpNowSeconds()});
    }
    FakeGraphDrive(const FakeGraphDrive&)            = delete;
    FakeGraphDrive& operator=(const FakeGraphDrive&) = delete;
    FakeGraphDrive(FakeGraphDrive&&)                 = delete;
    FakeGraphDrive& operator=(FakeGraphDrive&&)      = delete;

    void SetBaseUrl(std::string baseUrl)
    {
        std::scoped_lock lock(_mutex);
        _baseUrl = std::move(baseUrl);
    }

    void SeedFolder(std::string_view path)
    {
        std::scoped_lock lock(_mutex);
        static_cast<void>(EnsureItemLocked(path, true));
    }

    void SeedFile(std::string_view path, std::vector<uint8_t> bytes)
    {
        std::scoped_lock lock(_mutex);
        DriveItem* item = EnsureItemLocked(path, false);
        if (item != nullptr)
        {
            item->bytes = std::move(bytes);
            Touch(*item);
        }
    }

    [[nodiscard]] bool Exists(std::string_view path) const
    {
        std::scoped_lock lock(_mutex);
        return FindByPathLocked(path) != nullptr;
    }

    [[nodiscard]] LoopbackHttpResponse Handle(const LoopbackHttpRequest& request)
    {
        std::scoped_lock lock(_mutex);
        std::string_view path(request.path);
        if (path.starts_with("/download/"))
        {
            return HandleDownloadLocked(request, path.substr(10u));
        }
        if (path.starts_with("/upload/"))
        {
            return HandleUploadLocked(request, path.substr(8u));
        }
        if (! path.starts_with("/v1.0/"))
        {
            return ErrorResponse(404, "itemNotFound", "unknown route");
        }
        path.remove_prefix(6u);
        if (path == "me/drive" || path == "drives/" + std::string(kDriveId))
        {
            return JsonResponse(200, std::format(R"({{"id":{},"name":"Microsoft Drive SelfTest","webUrl":"{}/web"}})", JsonQuote(kDriveId), _baseUrl));
        }
        // Strip the drive prefix: `me/drive/...` or `drives/{id}/...`.
        if (path.starts_with("me/drive/"))
        {
            path.remove_prefix(9u);
        }
        else if (path.starts_with("drives/"))
        {
            const size_t slash = path.find('/', 7u);
            if (slash == std::string_view::npos)
            {
                return ErrorResponse(404, "itemNotFound", "unknown drive route");
            }
            path.remove_prefix(slash + 1u);
        }
        else
        {
            return ErrorResponse(404, "itemNotFound", "unknown route");
        }

        if (path.starts_with("items/"))
        {
            path.remove_prefix(6u);
            const size_t slash          = path.find('/');
            const std::string id        = std::string(path.substr(0, slash));
            const std::string_view tail = slash == std::string_view::npos ? std::string_view{} : path.substr(slash);
            DriveItem* item             = FindByIdLocked(id);
            if (item == nullptr)
            {
                return ErrorResponse(404, "itemNotFound", "The resource could not be found.");
            }
            return HandleItemLocked(request, *item, tail);
        }
        if (path.starts_with("root"))
        {
            path.remove_prefix(4u);
            std::string drivePath = "/";
            std::string_view tail = path;
            if (path.starts_with(":/"))
            {
                const size_t close = path.find(':', 2u);
                if (close == std::string_view::npos)
                {
                    return ErrorResponse(400, "invalidRequest", "unterminated path");
                }
                drivePath = "/" + std::string(path.substr(2u, close - 2u));
                tail      = path.substr(close + 1u);
            }
            if (request.method == "PUT" && tail == "/content")
            {
                return HandleSimpleUploadLocked(request, drivePath);
            }
            if (request.method == "POST" && tail == "/createUploadSession")
            {
                return HandleCreateUploadSessionLocked(request, drivePath);
            }
            DriveItem* item = FindByPathLocked(drivePath);
            if (item == nullptr)
            {
                return ErrorResponse(404, "itemNotFound", "The resource could not be found.");
            }
            return HandleItemLocked(request, *item, tail);
        }
        return ErrorResponse(404, "itemNotFound", "unknown route");
    }

private:
    void Touch(DriveItem& item) noexcept
    {
        item.version  = ++_nextVersion;
        item.modified = Common::SelfTest::LoopbackHttpNowSeconds();
    }

    [[nodiscard]] static std::string Etag(const DriveItem& item)
    {
        return std::format("\"v{}\"", item.version);
    }

    [[nodiscard]] DriveItem* FindByIdLocked(std::string_view id) noexcept
    {
        for (DriveItem& item : _items)
        {
            if (item.id == id)
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] DriveItem* FindChildLocked(std::string_view parentId, std::string_view name) noexcept
    {
        for (DriveItem& item : _items)
        {
            if (item.parentId == parentId && EqualsNoCaseAscii(item.name, name))
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] const DriveItem* FindByPathLocked(std::string_view path) const noexcept
    {
        return const_cast<FakeGraphDrive*>(this)->FindByPathLocked(path);
    }

    [[nodiscard]] DriveItem* FindByPathLocked(std::string_view path) noexcept
    {
        DriveItem* current = FindByIdLocked("root");
        size_t start       = 0u;
        while (current != nullptr && start < path.size())
        {
            if (path[start] == '/')
            {
                ++start;
                continue;
            }
            const size_t slash          = path.find('/', start);
            const std::string_view name = path.substr(start, slash == std::string_view::npos ? std::string_view::npos : slash - start);
            current                     = FindChildLocked(current->id, name);
            start                       = slash == std::string_view::npos ? path.size() : slash + 1u;
        }
        return current;
    }

    [[nodiscard]] DriveItem* EnsureItemLocked(std::string_view path, bool isFolder)
    {
        DriveItem* current = FindByIdLocked("root");
        size_t start       = 0u;
        while (current != nullptr && start < path.size())
        {
            if (path[start] == '/')
            {
                ++start;
                continue;
            }
            const size_t slash          = path.find('/', start);
            const std::string_view name = path.substr(start, slash == std::string_view::npos ? std::string_view::npos : slash - start);
            const bool last             = slash == std::string_view::npos;
            DriveItem* child            = FindChildLocked(current->id, name);
            if (child == nullptr)
            {
                child = &AddItemLocked(std::string(name), current->id, last ? isFolder : true);
            }
            current = child;
            start   = last ? path.size() : slash + 1u;
        }
        return current;
    }

    [[nodiscard]] DriveItem& AddItemLocked(std::string name, std::string parentId, bool isFolder)
    {
        _items.push_back(DriveItem{.id       = std::format("item-{}", ++_nextId),
                                   .name     = std::move(name),
                                   .parentId = std::move(parentId),
                                   .isFolder = isFolder,
                                   .bytes    = {},
                                   .version  = ++_nextVersion,
                                   .modified = Common::SelfTest::LoopbackHttpNowSeconds()});
        return _items.back();
    }

    void RemoveSubtreeLocked(std::string_view id)
    {
        std::vector<std::string> children;
        for (const DriveItem& item : _items)
        {
            if (item.parentId == id)
            {
                children.push_back(item.id);
            }
        }
        for (const std::string& child : children)
        {
            RemoveSubtreeLocked(child);
        }
        std::erase_if(_items, [&](const DriveItem& item) noexcept { return item.id == id; });
    }

    [[nodiscard]] std::string ItemJsonLocked(const DriveItem& item) const
    {
        std::string json = std::format(
            R"({{"id":{},"name":{},"eTag":{},"cTag":{},"size":{},"parentReference":{{"id":{},"driveId":{}}},"createdDateTime":"{}","lastModifiedDateTime":"{}",)",
            JsonQuote(item.id),
            JsonQuote(item.name),
            JsonQuote(Etag(item)),
            JsonQuote(Etag(item)),
            item.bytes.size(),
            JsonQuote(item.parentId),
            JsonQuote(kDriveId),
            Common::SelfTest::LoopbackHttpIso8601(item.modified),
            Common::SelfTest::LoopbackHttpIso8601(item.modified));
        if (item.isFolder)
        {
            size_t childCount = 0u;
            for (const DriveItem& candidate : _items)
            {
                if (candidate.parentId == item.id)
                {
                    ++childCount;
                }
            }
            json += std::format(R"("folder":{{"childCount":{}}}}})", childCount);
        }
        else
        {
            // Graph reports file.hashes for stored content (R3-2 writer proof); business drives carry
            // sha256Hash + quickXorHash, personal drives sha1Hash + quickXorHash. The fixture reports all three.
            const auto digestBase64 = [&](Common::Crypto::ContentDigestAlgorithm algorithm)
            {
                std::vector<std::byte> digest;
                static_cast<void>(Common::Crypto::ComputeContentDigest(algorithm, std::as_bytes(std::span(item.bytes)), digest));
                return JsonQuote(Common::Crypto::EncodeBase64Digest(digest));
            };
            json += std::format(
                R"("file":{{"mimeType":"application/octet-stream","hashes":{{"sha256Hash":{},"sha1Hash":{},"quickXorHash":{}}}}},"@microsoft.graph.downloadUrl":{}}})",
                digestBase64(Common::Crypto::ContentDigestAlgorithm::Sha256),
                digestBase64(Common::Crypto::ContentDigestAlgorithm::Sha1),
                digestBase64(Common::Crypto::ContentDigestAlgorithm::QuickXor),
                JsonQuote(_baseUrl + "/download/" + item.id));
        }
        return json;
    }

    [[nodiscard]] LoopbackHttpResponse ChildrenLocked(const LoopbackHttpRequest& request, const DriveItem& parent)
    {
        size_t top = 200u;
        if (const std::string topText = request.Query("$top"); ! topText.empty())
        {
            static_cast<void>(std::from_chars(topText.data(), topText.data() + topText.size(), top));
            top = std::clamp<size_t>(top, 1u, 999u);
        }
        size_t skip = 0u;
        if (const std::string skipText = request.Query("$skip"); ! skipText.empty())
        {
            static_cast<void>(std::from_chars(skipText.data(), skipText.data() + skipText.size(), skip));
        }
        std::vector<const DriveItem*> children;
        for (const DriveItem& item : _items)
        {
            if (item.parentId == parent.id)
            {
                children.push_back(&item);
            }
        }
        std::string json = R"({"value":[)";
        bool first       = true;
        size_t index     = skip;
        for (; index < children.size() && index < skip + top; ++index)
        {
            if (! first)
            {
                json.push_back(',');
            }
            first = false;
            json += ItemJsonLocked(*children[index]);
        }
        json += "]";
        if (index < children.size())
        {
            json += std::format(R"(,"@odata.nextLink":{})",
                                JsonQuote(std::format("{}/v1.0/drives/{}/items/{}/children?$top={}&$skip={}", _baseUrl, kDriveId, parent.id, top, index)));
        }
        json += "}";
        LoopbackHttpResponse response = JsonResponse(200, std::move(json));
        response.dripBody             = true;
        return response;
    }

    [[nodiscard]] LoopbackHttpResponse HandleItemLocked(const LoopbackHttpRequest& request, DriveItem& item, std::string_view tail)
    {
        if (request.method == "GET")
        {
            if (tail == "/children")
            {
                if (! item.isFolder)
                {
                    return ErrorResponse(400, "invalidRequest", "not a folder");
                }
                return ChildrenLocked(request, item);
            }
            if (tail.empty())
            {
                return JsonResponse(200, ItemJsonLocked(item));
            }
            return ErrorResponse(404, "itemNotFound", "unknown item route");
        }
        if (request.method == "POST" && tail == "/children")
        {
            if (! item.isFolder)
            {
                return ErrorResponse(400, "invalidRequest", "not a folder");
            }
            const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
            const std::string name = FindJsonString(body, "name");
            if (name.empty())
            {
                return ErrorResponse(400, "invalidRequest", "name required");
            }
            // The plugin creates folders through POST children; files always go through content uploads.
            const bool folder = body.find("\"file\"") == std::string::npos;
            if (FindChildLocked(item.id, name) != nullptr)
            {
                return ErrorResponse(409, "nameAlreadyExists", "An item with the same name already exists.");
            }
            DriveItem& created = AddItemLocked(name, item.id, folder);
            return JsonResponse(201, ItemJsonLocked(created));
        }
        if (request.method == "PATCH" && tail.empty())
        {
            if (const std::string* ifMatch = request.Header("if-match");
                ifMatch != nullptr && StripQuotes(*ifMatch) != StripQuotes(Etag(item)) && *ifMatch != "*")
            {
                return ErrorResponse(412, "preconditionFailed", "ETag does not match.");
            }
            const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
            std::string newName     = FindJsonString(body, "name");
            std::string newParentId = item.parentId;
            if (const size_t parentAt = body.find("\"parentReference\""); parentAt != std::string::npos)
            {
                const std::string parentId = FindJsonString(body, "id", parentAt);
                if (! parentId.empty())
                {
                    newParentId = parentId;
                }
            }
            if (newName.empty())
            {
                newName = item.name;
            }
            DriveItem* newParent = FindByIdLocked(newParentId);
            if (newParent == nullptr || ! newParent->isFolder)
            {
                return ErrorResponse(404, "itemNotFound", "parent not found");
            }
            if (DriveItem* collision = FindChildLocked(newParentId, newName); collision != nullptr && collision != &item)
            {
                return ErrorResponse(409, "nameAlreadyExists", "An item with the same name already exists.");
            }
            item.name     = newName;
            item.parentId = newParentId;
            Touch(item);
            return JsonResponse(200, ItemJsonLocked(item));
        }
        if (request.method == "DELETE" && tail.empty())
        {
            if (item.id == "root")
            {
                return ErrorResponse(403, "accessDenied", "cannot delete root");
            }
            RemoveSubtreeLocked(item.id);
            return EmptyResponse(204);
        }
        return ErrorResponse(405, "methodNotAllowed", request.method);
    }

    [[nodiscard]] LoopbackHttpResponse StoreFileLocked(std::string_view drivePath, std::vector<uint8_t> bytes, bool allowOverwrite)
    {
        const size_t slash           = drivePath.find_last_of('/');
        const std::string parentPath = slash == std::string_view::npos || slash == 0u ? std::string("/") : std::string(drivePath.substr(0, slash));
        const std::string leaf       = std::string(drivePath.substr(slash == std::string_view::npos ? 0u : slash + 1u));
        DriveItem* parent            = FindByPathLocked(parentPath);
        if (parent == nullptr || ! parent->isFolder || leaf.empty())
        {
            return ErrorResponse(404, "itemNotFound", "parent not found");
        }
        DriveItem* existing = FindChildLocked(parent->id, leaf);
        if (existing != nullptr)
        {
            if (! allowOverwrite)
            {
                return ErrorResponse(409, "nameAlreadyExists", "An item with the same name already exists.");
            }
            if (existing->isFolder)
            {
                return ErrorResponse(409, "nameAlreadyExists", "a folder occupies the name");
            }
            existing->bytes = std::move(bytes);
            Touch(*existing);
            return JsonResponse(200, ItemJsonLocked(*existing));
        }
        DriveItem& created = AddItemLocked(leaf, parent->id, false);
        created.bytes      = std::move(bytes);
        return JsonResponse(201, ItemJsonLocked(created));
    }

    // R3-1: `If-Match` names the occupant a replacement was granted for.
    [[nodiscard]] bool IfMatchHoldsLocked(const LoopbackHttpRequest& request, std::string_view drivePath) const
    {
        const std::string* ifMatch = request.Header("if-match");
        if (ifMatch == nullptr)
        {
            return true;
        }
        const DriveItem* current = FindByPathLocked(drivePath);
        return current != nullptr && StripQuotes(*ifMatch) == StripQuotes(Etag(*current));
    }

    [[nodiscard]] LoopbackHttpResponse HandleSimpleUploadLocked(const LoopbackHttpRequest& request, std::string_view drivePath)
    {
        const std::string* ifNoneMatch = request.Header("if-none-match");
        const bool allowOverwrite      = ifNoneMatch == nullptr;
        if (! allowOverwrite && FindByPathLocked(drivePath) != nullptr)
        {
            return ErrorResponse(412, "preconditionFailed", "The item already exists.");
        }
        if (! IfMatchHoldsLocked(request, drivePath))
        {
            return ErrorResponse(412, "preconditionFailed", "ETag does not match.");
        }
        return StoreFileLocked(drivePath, request.body, allowOverwrite);
    }

    [[nodiscard]] LoopbackHttpResponse HandleCreateUploadSessionLocked(const LoopbackHttpRequest& request, std::string_view drivePath)
    {
        const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
        const bool allowOverwrite = FindJsonString(body, "@microsoft.graph.conflictBehavior") == "replace";
        if (! allowOverwrite && FindByPathLocked(drivePath) != nullptr)
        {
            return ErrorResponse(409, "nameAlreadyExists", "An item with the same name already exists.");
        }
        if (! IfMatchHoldsLocked(request, drivePath))
        {
            return ErrorResponse(412, "preconditionFailed", "ETag does not match.");
        }
        const std::string sessionId = std::format("session-{}", ++_nextId);
        _uploads.insert_or_assign(sessionId, UploadSession{.path = std::string(drivePath), .allowOverwrite = allowOverwrite, .bytes = {}, .total = 0u});
        return JsonResponse(200,
                            std::format(R"({{"uploadUrl":{},"expirationDateTime":"2099-01-01T00:00:00Z"}})", JsonQuote(_baseUrl + "/upload/" + sessionId)));
    }

    [[nodiscard]] LoopbackHttpResponse HandleUploadLocked(const LoopbackHttpRequest& request, std::string_view sessionId)
    {
        const auto session = _uploads.find(std::string(sessionId));
        if (session == _uploads.end())
        {
            return ErrorResponse(404, "itemNotFound", "unknown upload session");
        }
        if (request.method == "DELETE")
        {
            _uploads.erase(session);
            return EmptyResponse(204);
        }
        if (request.method != "PUT")
        {
            return ErrorResponse(405, "methodNotAllowed", request.method);
        }
        const std::string* contentRange = request.Header("content-range");
        if (contentRange == nullptr || ! contentRange->starts_with("bytes "))
        {
            return ErrorResponse(400, "invalidRequest", "Content-Range required");
        }
        uint64_t first = 0u;
        uint64_t last  = 0u;
        uint64_t total = 0u;
        const std::string_view spec(std::string_view(*contentRange).substr(6u));
        const size_t dash  = spec.find('-');
        const size_t slash = spec.find('/');
        if (dash == std::string_view::npos || slash == std::string_view::npos || slash < dash)
        {
            return ErrorResponse(400, "invalidRequest", "bad Content-Range");
        }
        static_cast<void>(std::from_chars(spec.data(), spec.data() + dash, first));
        static_cast<void>(std::from_chars(spec.data() + dash + 1u, spec.data() + slash, last));
        static_cast<void>(std::from_chars(spec.data() + slash + 1u, spec.data() + spec.size(), total));
        UploadSession& upload = session->second;
        if (upload.total == 0u)
        {
            upload.total = total;
        }
        if (first != upload.bytes.size() || last + 1u - first != request.body.size() || total != upload.total)
        {
            return ErrorResponse(416, "invalidRange", "unexpected range");
        }
        upload.bytes.insert(upload.bytes.end(), request.body.begin(), request.body.end());
        if (upload.bytes.size() < upload.total)
        {
            return JsonResponse(
                202, std::format(R"({{"expirationDateTime":"2099-01-01T00:00:00Z","nextExpectedRanges":["{}-{}"]}})", upload.bytes.size(), upload.total - 1u));
        }
        const std::string drivePath   = upload.path;
        const bool allowOverwrite     = upload.allowOverwrite;
        std::vector<uint8_t> complete = std::move(upload.bytes);
        _uploads.erase(session);
        return StoreFileLocked(drivePath, std::move(complete), allowOverwrite);
    }

    [[nodiscard]] LoopbackHttpResponse HandleDownloadLocked(const LoopbackHttpRequest& request, std::string_view id)
    {
        const DriveItem* item = FindByIdLocked(id);
        if (item == nullptr || item->isFolder)
        {
            return ErrorResponse(404, "itemNotFound", "unknown content");
        }
        if (const std::string* ifMatch = request.Header("if-match"); ifMatch != nullptr && StripQuotes(*ifMatch) != StripQuotes(Etag(*item)))
        {
            return ErrorResponse(412, "preconditionFailed", "ETag does not match.");
        }
        const size_t size = item->bytes.size();
        LoopbackHttpResponse response;
        response.headers.emplace_back("ETag", Etag(*item));
        response.headers.emplace_back("Accept-Ranges", "bytes");
        response.headers.emplace_back("Content-Type", "application/octet-stream");
        if (const std::string* range = request.Header("range"); range != nullptr && range->starts_with("bytes="))
        {
            const std::string_view spec(std::string_view(*range).substr(6u));
            const size_t dash = spec.find('-');
            size_t first      = 0u;
            size_t last       = size == 0u ? 0u : size - 1u;
            if (dash != std::string_view::npos)
            {
                static_cast<void>(std::from_chars(spec.data(), spec.data() + dash, first));
                if (dash + 1u < spec.size())
                {
                    static_cast<void>(std::from_chars(spec.data() + dash + 1u, spec.data() + spec.size(), last));
                }
            }
            if (size == 0u || first >= size)
            {
                LoopbackHttpResponse unsatisfiable = ErrorResponse(416, "invalidRange", "The requested range is not satisfiable");
                unsatisfiable.headers.emplace_back("Content-Range", std::format("bytes */{}", size));
                return unsatisfiable;
            }
            last            = (std::min)(last, size - 1u);
            response.status = 206;
            response.headers.emplace_back("Content-Range", std::format("bytes {}-{}/{}", first, last, size));
            response.body.assign(item->bytes.begin() + static_cast<ptrdiff_t>(first), item->bytes.begin() + static_cast<ptrdiff_t>(last + 1u));
            return response;
        }
        response.status = 200;
        response.body   = item->bytes;
        return response;
    }

    mutable std::mutex _mutex;
    std::string _baseUrl;
    std::vector<DriveItem> _items;
    std::map<std::string, UploadSession> _uploads;
    unsigned int _nextId      = 0u;
    unsigned int _nextVersion = 1u;
};

// The fixture: the drive model and the loopback server that serves it.
class FakeGraphEndpoint final
{
public:
    FakeGraphEndpoint()
    {
        _server = std::make_unique<LoopbackHttpServer>([this](const LoopbackHttpRequest& request) { return _drive.Handle(request); });
    }
    FakeGraphEndpoint(const FakeGraphEndpoint&)            = delete;
    FakeGraphEndpoint& operator=(const FakeGraphEndpoint&) = delete;
    FakeGraphEndpoint(FakeGraphEndpoint&&)                 = delete;
    FakeGraphEndpoint& operator=(FakeGraphEndpoint&&)      = delete;

    [[nodiscard]] HRESULT Start()
    {
        const HRESULT hr = _server->Start();
        if (FAILED(hr))
        {
            return hr;
        }
        _baseUrl = std::format("http://127.0.0.1:{}", _server->Port());
        _drive.SetBaseUrl(_baseUrl);
        return S_OK;
    }
    void Stop() noexcept
    {
        _server->Stop();
    }
    [[nodiscard]] unsigned short Port() const noexcept
    {
        return _server->Port();
    }
    [[nodiscard]] std::wstring GraphBaseUrl() const
    {
        return Common::SelfTest::LoopbackHttpUtf16FromUtf8(_baseUrl + "/v1.0");
    }
    [[nodiscard]] FakeGraphDrive& Drive() noexcept
    {
        return _drive;
    }
    [[nodiscard]] LoopbackHttpServer& Server() noexcept
    {
        return *_server;
    }

private:
    FakeGraphDrive _drive;
    std::unique_ptr<LoopbackHttpServer> _server;
    std::string _baseUrl;
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
} // namespace

// R0f-Graph witness. (1) Bytes still flowing: a Recycle of a folder whose children listing the
// fixture drips returns ERROR_CANCELLED within a few read chunks of Cancel and deletes nothing. (2) A
// server that stopped answering: Cancel is reported as the host's verdict as soon as the provider-
// owned bound (WinHTTP connect/send/receive timeouts) returns the request, and (3) without Cancel
// the same bound returns it on its own as a transport failure; the item survives.
void RunGraphStalledRequestCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    try
    {
        FakeGraphEndpoint endpoint;
        endpoint.Drive().SeedFile("/stall.txt", std::vector<uint8_t>(64u, 0x5Au));
        for (unsigned int index = 0u; index < 60u; ++index)
        {
            endpoint.Drive().SeedFile(std::format("/drip/object-{:03}.bin", index), std::vector<uint8_t>(16u, static_cast<uint8_t>(index)));
        }
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"Graph stalled-request fixture should start", passed, failed))
        {
            return;
        }
        const std::wstring baseUrl = endpoint.GraphBaseUrl();
        ConfigureFakeGraph(baseUrl.c_str(), true);
        auto restore = wil::scope_exit([&]() noexcept
        {
            ConfigureFakeGraph(nullptr, false);
            endpoint.Stop();
        });

        wil::com_ptr<FileSystemMicrosoftDrive> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDrivePersonal, nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem), L"Graph stalled-request proof should create a Microsoft Drive instance", passed, failed))
        {
            return;
        }
        constexpr uint32_t kConnectTimeoutMs = 2'000u;
        constexpr uint32_t kRequestTimeoutMs = 3'000u;
        HRESULT hr                           = fileSystem->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":3000})");
        if (! DebugCheck(SUCCEEDED(hr), L"Microsoft Drive instance should accept the fixture configuration", passed, failed))
        {
            return;
        }
        const unsigned long boundMs = FileSystemMicrosoftDriveInternal::GraphProviderWatchdogTimeoutMs(kConnectTimeoutMs, kRequestTimeoutMs);
        DebugCheck(
            boundMs > 0u && boundMs <= 20'000u, L"the Graph provider-owned bound must be a small nonzero value for the fixture timeouts", passed, failed);

        const auto recycleWithControl = [&](const wchar_t* path, CancelControl* control, std::atomic<HRESULT>& result) noexcept
        {
            FileSystemOptions options{};
            options.sizeBytes        = sizeof(options);
            options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
            options.operationControl = control;
            result.store(fileSystem->DeleteItem(path, FILESYSTEM_FLAG_USE_RECYCLE_BIN, control != nullptr ? &options : nullptr, nullptr, nullptr),
                         std::memory_order_release);
        };

        // 1) Bytes flowing: the children listing drips; Cancel lands between read chunks.
        endpoint.Server().SetDrip(512u, 40u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { recycleWithControl(L"/@conn:microsoft-drive-selftest/drip", &control, deleteHr); });
            const ULONGLONG waitStart = GetTickCount64();
            while (endpoint.Server().RequestCount("GET") < 2u && GetTickCount64() - waitStart < 10'000u)
            {
                Sleep(20u);
            }
            const bool listingReached = endpoint.Server().RequestCount("GET") >= 2u;
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
            DebugCheck(listingReached, L"the folder Recycle must list the children on the fixture", passed, failed);
            DebugCheck(returned, L"a request whose body is still arriving must return after Cancel", passed, failed);
            DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"a canceled streaming request must report ERROR_CANCELLED",
                       passed,
                       failed);
            std::fwprintf(stderr, L"[Microsoft Drive] streaming request returned %llu ms after Cancel\n", static_cast<unsigned long long>(cancelMs));
            DebugCheck(cancelMs < 3'000u, L"Cancel must return a streaming request within a few read chunks", passed, failed);
            DebugCheck(endpoint.Server().RequestCount("DELETE") == 0u, L"a Recycle canceled while listing must not delete anything", passed, failed);
            if (! returned)
            {
                return;
            }
        }
        endpoint.Server().SetDrip(0u, 0u);

        // 2) Silent server + Cancel: the provider-owned bound returns the request; the verdict is the host's.
        endpoint.Server().SetStallMethod("DELETE", 60'000u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { recycleWithControl(L"/@conn:microsoft-drive-selftest/stall.txt", &control, deleteHr); });
            const ULONGLONG waitStart = GetTickCount64();
            while (endpoint.Server().RequestCount("DELETE") == 0u && GetTickCount64() - waitStart < 10'000u)
            {
                Sleep(20u);
            }
            const bool deleteReached = endpoint.Server().RequestCount("DELETE") != 0u;
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
            std::fwprintf(stderr,
                          L"[Microsoft Drive] stalled request returned %llu ms after Cancel (declared bound %lu ms)\n",
                          static_cast<unsigned long long>(cancelMs),
                          boundMs);
            DebugCheck(cancelMs <= boundMs + 5'000u, L"Cancel on a silent server must return within the declared provider-owned bound", passed, failed);
            if (! returned)
            {
                return;
            }
        }

        // 3) Silent server, no Cancel: the bound returns it on its own as a transport failure.
        {
            std::atomic<HRESULT> deleteHr{E_PENDING};
            const ULONGLONG boundStart = GetTickCount64();
            recycleWithControl(L"/@conn:microsoft-drive-selftest/stall.txt", nullptr, deleteHr);
            const ULONGLONG elapsedMs = GetTickCount64() - boundStart;
            const HRESULT boundHr     = deleteHr.load(std::memory_order_acquire);
            DebugCheck(FAILED(boundHr) && boundHr != HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"an un-canceled stalled request must fail through the transport bound, not as a cancel",
                       passed,
                       failed);
            std::fwprintf(stderr,
                          L"[Microsoft Drive] stalled request returned on its own after %llu ms (declared bound %lu ms)\n",
                          static_cast<unsigned long long>(elapsedMs),
                          boundMs);
            DebugCheck(elapsedMs <= boundMs + 5'000u, L"the provider-owned bound must return a stalled request on its own", passed, failed);
        }
        endpoint.Server().SetStallMethod({}, 0u);
        DebugCheck(endpoint.Drive().Exists("/stall.txt"), L"a stalled DELETE the client gave up on must not be reported as deleted", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemMicrosoftDrive stalled-request selftest failed after std::exception.");
        DebugCheck(false, L"Graph stalled-request proof should not throw std::exception", passed, failed);
    }
}
// R0-RC4: an identity-less replace whose host token carries no timestamp is refused before any
// upload; the writer must not pin If-Match to whichever occupant is live now.
void RunGraphZeroTimestampReplaceSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    try
    {
        FakeGraphEndpoint endpoint;
        endpoint.Drive().SeedFile("/occupant.txt", std::vector<uint8_t>(64u, 0x5Au));
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"R0-RC4 Graph fixture should start", passed, failed))
        {
            return;
        }
        const std::wstring baseUrl = endpoint.GraphBaseUrl();
        ConfigureFakeGraph(baseUrl.c_str(), true);
        auto restore = wil::scope_exit([&]() noexcept
        {
            ConfigureFakeGraph(nullptr, false);
            endpoint.Stop();
        });

        wil::com_ptr<FileSystemMicrosoftDrive> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDrivePersonal, nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem), L"R0-RC4 Graph proof should create a Microsoft Drive instance", passed, failed))
        {
            return;
        }
        HRESULT hr = fileSystem->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":3000})");
        if (! DebugCheck(SUCCEEDED(hr), L"R0-RC4 Graph instance should accept the fixture configuration", passed, failed))
        {
            return;
        }

        wil::com_ptr<IFileWriter> writer;
        hr = fileSystem->CreateFileWriter(L"/@conn:microsoft-drive-selftest/occupant.txt", FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.addressof());
        if (! DebugCheck(SUCCEEDED(hr) && writer, L"R0-RC4: Graph should open an overwrite writer on the seeded occupant", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileWriterExpectedReplacement> replacement;
        hr = writer->QueryInterface(IID_PPV_ARGS(replacement.addressof()));
        if (! DebugCheck(SUCCEEDED(hr) && replacement, L"R0-RC4: the Graph writer should expose the expected-replacement contract", passed, failed))
        {
            return;
        }
        FileSystemBasicInformation expected{};
        expected.sizeBytes = sizeof(expected); // lastWriteTime stays 0: the host holds no identity for the occupant
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
                   L"R0-RC4: a Graph identity-less replace with a zero host timestamp must be refused, not pinned to the live occupant",
                   passed,
                   failed);
        DebugCheck(endpoint.Drive().Exists("/occupant.txt"), L"R0-RC4: the Graph occupant must survive the refused replace", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemMicrosoftDrive zero-timestamp replace selftest failed after std::exception.");
        DebugCheck(false, L"R0-RC4 Graph proof should not throw std::exception", passed, failed);
    }
}

} // namespace FileSystemMicrosoftDriveSelfTest

// Host self-tests drive the fake Graph endpoint through these exports: start one on a loopback port
// (the plugin is redirected to it, the OAuth token bypassed, and paths under
// `/@conn:microsoft-drive-selftest/` resolve to the fixture drive), run File Operations, then stop it.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderMicrosoftDriveStartFakeGraphForSelfTest(unsigned int* port, void** endpoint) noexcept
{
    if (port == nullptr || endpoint == nullptr)
    {
        return E_POINTER;
    }
    *port     = 0u;
    *endpoint = nullptr;
    try
    {
        auto fixture     = std::make_unique<FileSystemMicrosoftDriveSelfTest::FakeGraphEndpoint>();
        const HRESULT hr = fixture->Start();
        if (FAILED(hr))
        {
            return hr;
        }
        const std::wstring baseUrl = fixture->GraphBaseUrl();
        FileSystemMicrosoftDriveSelfTest::ConfigureFakeGraph(baseUrl.c_str(), true);
        *port     = fixture->Port();
        *endpoint = fixture.release();
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

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderMicrosoftDriveFakeGraphRequestLogForSelfTest(void* endpoint,
                                                                                                             wchar_t* buffer,
                                                                                                             unsigned int capacity) noexcept
{
    if (endpoint == nullptr || buffer == nullptr || capacity == 0u)
    {
        return E_POINTER;
    }
    buffer[0] = L'\0';
    try
    {
        const std::wstring text = static_cast<FileSystemMicrosoftDriveSelfTest::FakeGraphEndpoint*>(endpoint)->Server().RequestLog(48u);
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

extern "C" __declspec(dllexport) void __stdcall RedSalamanderMicrosoftDriveStopFakeGraphForSelfTest(void* endpoint) noexcept
{
    auto* fixture = static_cast<FileSystemMicrosoftDriveSelfTest::FakeGraphEndpoint*>(endpoint);
    if (fixture == nullptr)
    {
        return;
    }
    FileSystemMicrosoftDriveSelfTest::ConfigureFakeGraph(nullptr, false);
    fixture->Stop();
    delete fixture;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderMicrosoftDriveR0fContainmentSelfTests(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    FileSystemMicrosoftDriveSelfTest::RunGraphStalledRequestCancelSelfTests(*passed, *failed);
    return *failed == 0u ? S_OK : E_FAIL;
}

#else

static_assert(true);

#endif

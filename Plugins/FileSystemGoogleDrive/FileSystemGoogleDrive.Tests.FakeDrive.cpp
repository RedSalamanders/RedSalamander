#if defined(ENABLE_TESTS)

// The loopback HTTP fixture must precede the Windows headers the plugin header pulls in.
#include "ContentDigest.h"
#include "LoopbackHttpFixture.h"

#include "FileSystemGoogleDrive.h"

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

// R0f-GDrive fixture: a deterministic loopback Google Drive v3 API (`/drive/v3/files` listing with
// `q` parents/name/trashed filters and paging, metadata by id, `alt=media` ranged download, JSON
// create, resumable upload sessions with Content-Range chunks, PATCH name/parents/trashed, DELETE,
// copy, `about`, and the OAuth token endpoint). Host self-tests start it through the exports at the
// end of this file; the debug self-test proves the provider-owned bound and Cancel on it. The plugin
// is pointed at the fixture through `ConfigureFakeDrive`.
namespace FileSystemGoogleDriveSelfTest
{
namespace
{
using Common::SelfTest::LoopbackHttpRequest;
using Common::SelfTest::LoopbackHttpResponse;
using Common::SelfTest::LoopbackHttpServer;

constexpr Common::DebugSelfTest::Check DebugCheck{L"Google Drive"};
constexpr std::string_view kFolderMime = "application/vnd.google-apps.folder";

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

// Minimal JSON lookups for the fixed request bodies the plugin sends.
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

[[nodiscard]] std::vector<std::string> FindJsonStringArray(std::string_view json, std::string_view key)
{
    std::vector<std::string> values;
    const std::string needle = "\"" + std::string(key) + "\"";
    const size_t keyAt       = json.find(needle);
    if (keyAt == std::string_view::npos)
    {
        return values;
    }
    const size_t open  = json.find('[', keyAt);
    const size_t close = open == std::string_view::npos ? std::string_view::npos : json.find(']', open);
    if (open == std::string_view::npos || close == std::string_view::npos)
    {
        return values;
    }
    std::string_view inner = json.substr(open + 1u, close - open - 1u);
    size_t pos             = 0u;
    while ((pos = inner.find('"', pos)) != std::string_view::npos)
    {
        const size_t end = inner.find('"', pos + 1u);
        if (end == std::string_view::npos)
        {
            break;
        }
        values.emplace_back(inner.substr(pos + 1u, end - pos - 1u));
        pos = end + 1u;
    }
    return values;
}

[[nodiscard]] std::optional<bool> FindJsonBool(std::string_view json, std::string_view key)
{
    const std::string needle = "\"" + std::string(key) + "\"";
    const size_t keyAt       = json.find(needle);
    if (keyAt == std::string_view::npos)
    {
        return std::nullopt;
    }
    size_t index = keyAt + needle.size();
    while (index < json.size() && (json[index] == ' ' || json[index] == ':'))
    {
        ++index;
    }
    if (json.substr(index).starts_with("true"))
    {
        return true;
    }
    if (json.substr(index).starts_with("false"))
    {
        return false;
    }
    return std::nullopt;
}

[[nodiscard]] LoopbackHttpResponse JsonResponse(int status, std::string json)
{
    LoopbackHttpResponse response;
    response.status = status;
    response.body.assign(json.begin(), json.end());
    response.headers.emplace_back("Content-Type", "application/json; charset=UTF-8");
    return response;
}

[[nodiscard]] LoopbackHttpResponse ErrorResponse(int status, std::string_view reason, std::string_view message)
{
    return JsonResponse(status,
                        std::format(R"({{"error":{{"code":{},"message":{},"errors":[{{"reason":{}}}]}}}})", status, JsonQuote(message), JsonQuote(reason)));
}

[[nodiscard]] LoopbackHttpResponse EmptyResponse(int status)
{
    LoopbackHttpResponse response;
    response.status = status;
    return response;
}

struct DriveFile final
{
    std::string id;
    std::string name;
    std::string mimeType;
    std::vector<std::string> parents;
    std::vector<uint8_t> bytes;
    unsigned int version = 1u;
    bool trashed         = false;
    std::chrono::sys_seconds modified{};
};

struct UploadSession final
{
    std::string fileId; // existing file for an overwrite session, empty for a create session
    std::string name;
    std::vector<std::string> parents;
    std::vector<uint8_t> bytes;
    uint64_t total = 0u;
};

class FakeDrive final
{
public:
    FakeDrive()
    {
        _files.push_back(DriveFile{.id       = "root",
                                   .name     = "My Drive",
                                   .mimeType = std::string(kFolderMime),
                                   .parents  = {},
                                   .bytes    = {},
                                   .version  = 1u,
                                   .trashed  = false,
                                   .modified = Common::SelfTest::LoopbackHttpNowSeconds()});
    }
    FakeDrive(const FakeDrive&)            = delete;
    FakeDrive& operator=(const FakeDrive&) = delete;
    FakeDrive(FakeDrive&&)                 = delete;
    FakeDrive& operator=(FakeDrive&&)      = delete;

    void SetOrigin(std::string origin)
    {
        std::scoped_lock lock(_mutex);
        _origin = std::move(origin);
    }

    void SeedFolder(std::string_view path)
    {
        std::scoped_lock lock(_mutex);
        static_cast<void>(EnsureLocked(path, true));
    }

    void SeedFile(std::string_view path, std::vector<uint8_t> bytes)
    {
        std::scoped_lock lock(_mutex);
        DriveFile* file = EnsureLocked(path, false);
        if (file != nullptr)
        {
            file->bytes = std::move(bytes);
            Touch(*file);
        }
    }

    [[nodiscard]] bool Exists(std::string_view path) const
    {
        std::scoped_lock lock(_mutex);
        const DriveFile* file = FindByPathLocked(path);
        return file != nullptr && ! file->trashed;
    }

    void MutationCommitFailureSnapshot(FileSystemOperation operation, bool arm, unsigned int& requests, unsigned int& commits)
    {
        std::scoped_lock lock(_mutex);
        if (arm)
        {
            _failCopyAfterCommit   = operation == FILESYSTEM_COPY;
            _failDeleteAfterCommit = operation == FILESYSTEM_DELETE;
            _copyRequests          = 0u;
            _copyCommits           = 0u;
            _deleteRequests        = 0u;
            _deleteCommits         = 0u;
        }
        requests = operation == FILESYSTEM_COPY ? _copyRequests : _deleteRequests;
        commits  = operation == FILESYSTEM_COPY ? _copyCommits : _deleteCommits;
    }

    [[nodiscard]] LoopbackHttpResponse Handle(const LoopbackHttpRequest& request)
    {
        std::scoped_lock lock(_mutex);
        std::string_view path(request.path);
        if (path == "/token" && request.method == "POST")
        {
            return JsonResponse(200, R"({"access_token":"fake-access-token","expires_in":3600,"token_type":"Bearer"})");
        }
        if (path == "/drive/v3/about")
        {
            return JsonResponse(
                200,
                R"({"user":{"displayName":"Google Drive SelfTest","emailAddress":"selftest@example.invalid"},"storageQuota":{"limit":"1073741824","usage":"4096","usageInDrive":"4096"}})");
        }
        if (path.starts_with("/upload/session/"))
        {
            return HandleUploadSessionLocked(request, path.substr(16u));
        }
        if (path == "/upload/drive/v3/files" && request.method == "POST")
        {
            return HandleStartUploadLocked(request, {});
        }
        if (path.starts_with("/upload/drive/v3/files/") && request.method == "PATCH")
        {
            return HandleStartUploadLocked(request, path.substr(23u));
        }
        if (path == "/drive/v3/files")
        {
            if (request.method == "GET")
            {
                return HandleListLocked(request);
            }
            if (request.method == "POST")
            {
                return HandleCreateLocked(request);
            }
            return ErrorResponse(405, "methodNotAllowed", request.method);
        }
        if (path.starts_with("/drive/v3/files/"))
        {
            std::string_view rest       = path.substr(16u);
            const size_t slash          = rest.find('/');
            const std::string id        = std::string(rest.substr(0, slash));
            const std::string_view tail = slash == std::string_view::npos ? std::string_view{} : rest.substr(slash);
            if (request.method == "DELETE" && tail.empty())
            {
                ++_deleteRequests;
            }
            DriveFile* file = FindByIdLocked(id);
            if (file == nullptr)
            {
                return ErrorResponse(404, "notFound", "File not found: " + id);
            }
            if (request.method == "POST" && tail == "/copy")
            {
                return HandleCopyLocked(request, *file);
            }
            if (request.method == "GET" && tail.empty())
            {
                if (request.Query("alt") == "media")
                {
                    return HandleDownloadLocked(request, *file);
                }
                return JsonResponse(200, FileJsonLocked(*file));
            }
            if (request.method == "PATCH" && tail.empty())
            {
                return HandlePatchLocked(request, *file);
            }
            if (request.method == "DELETE" && tail.empty())
            {
                if (file->id == "root")
                {
                    return ErrorResponse(403, "cannotDeleteRoot", "The root folder cannot be deleted.");
                }
                RemoveSubtreeLocked(file->id);
                ++_deleteCommits;
                if (_failDeleteAfterCommit)
                {
                    return ErrorResponse(503, "backendError", "delete committed before the response failed");
                }
                return EmptyResponse(204);
            }
            return ErrorResponse(405, "methodNotAllowed", request.method);
        }
        return ErrorResponse(404, "notFound", "unknown route");
    }

private:
    void Touch(DriveFile& file) noexcept
    {
        file.version  = ++_nextVersion;
        file.modified = Common::SelfTest::LoopbackHttpNowSeconds();
    }

    [[nodiscard]] DriveFile* FindByIdLocked(std::string_view id) noexcept
    {
        for (DriveFile& file : _files)
        {
            if (file.id == id)
            {
                return &file;
            }
        }
        return nullptr;
    }

    [[nodiscard]] static bool HasParent(const DriveFile& file, std::string_view parentId) noexcept
    {
        return std::ranges::find(file.parents, parentId) != file.parents.end();
    }

    [[nodiscard]] DriveFile* FindChildLocked(std::string_view parentId, std::string_view name) noexcept
    {
        for (DriveFile& file : _files)
        {
            if (! file.trashed && file.name == name && HasParent(file, parentId))
            {
                return &file;
            }
        }
        return nullptr;
    }

    [[nodiscard]] const DriveFile* FindByPathLocked(std::string_view path) const noexcept
    {
        return const_cast<FakeDrive*>(this)->FindByPathLocked(path);
    }

    [[nodiscard]] DriveFile* FindByPathLocked(std::string_view path) noexcept
    {
        DriveFile* current = FindByIdLocked("root");
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

    [[nodiscard]] DriveFile* EnsureLocked(std::string_view path, bool isFolder)
    {
        DriveFile* current = FindByIdLocked("root");
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
            DriveFile* child            = FindChildLocked(current->id, name);
            if (child == nullptr)
            {
                child = &AddLocked(std::string(name), current->id, last ? isFolder : true);
            }
            current = child;
            start   = last ? path.size() : slash + 1u;
        }
        return current;
    }

    [[nodiscard]] DriveFile& AddLocked(std::string name, std::string parentId, bool isFolder)
    {
        _files.push_back(DriveFile{.id       = std::format("gd-{}", ++_nextId),
                                   .name     = std::move(name),
                                   .mimeType = isFolder ? std::string(kFolderMime) : std::string("application/octet-stream"),
                                   .parents  = {std::move(parentId)},
                                   .bytes    = {},
                                   .version  = ++_nextVersion,
                                   .trashed  = false,
                                   .modified = Common::SelfTest::LoopbackHttpNowSeconds()});
        return _files.back();
    }

    // The id is taken by value: erasing children moves vector elements, which would dangle a view
    // into the folder's own entry.
    void RemoveSubtreeLocked(const std::string id)
    {
        std::vector<std::string> children;
        for (const DriveFile& file : _files)
        {
            if (HasParent(file, id))
            {
                children.push_back(file.id);
            }
        }
        for (const std::string& child : children)
        {
            RemoveSubtreeLocked(child);
        }
        std::erase_if(_files, [&](const DriveFile& file) noexcept { return file.id == id; });
    }

    [[nodiscard]] std::string FileJsonLocked(const DriveFile& file) const
    {
        std::string parents = "[";
        for (size_t index = 0u; index < file.parents.size(); ++index)
        {
            if (index != 0u)
            {
                parents.push_back(',');
            }
            parents += JsonQuote(file.parents[index]);
        }
        parents.push_back(']');
        // Drive reports sha256Checksum for binary content it stores (R3-2 writer proof).
        std::vector<std::byte> sha256;
        static_cast<void>(Common::Crypto::ComputeContentDigest(Common::Crypto::ContentDigestAlgorithm::Sha256, std::as_bytes(std::span(file.bytes)), sha256));
        return std::format(
            R"({{"kind":"drive#file","id":{},"name":{},"mimeType":{},"size":"{}","modifiedTime":"{}","createdTime":"{}","trashed":{},"version":"{}","sha256Checksum":"{}","parents":{}}})",
            JsonQuote(file.id),
            JsonQuote(file.name),
            JsonQuote(file.mimeType),
            file.bytes.size(),
            Common::SelfTest::LoopbackHttpIso8601(file.modified),
            Common::SelfTest::LoopbackHttpIso8601(file.modified),
            file.trashed ? "true" : "false",
            file.version,
            Common::Crypto::EncodeHexDigest(sha256),
            parents);
    }

    // `q` grammar the plugin uses: `trashed = false and '<id>' in parents [and name = '<name>']`.
    [[nodiscard]] LoopbackHttpResponse HandleListLocked(const LoopbackHttpRequest& request)
    {
        const std::string q = request.Query("q");
        std::string parentId;
        std::string nameFilter;
        bool hasNameFilter = false;
        if (const size_t at = q.find("' in parents"); at != std::string::npos)
        {
            const size_t quote = q.rfind('\'', at - 1u);
            if (quote != std::string::npos)
            {
                parentId = q.substr(quote + 1u, at - quote - 1u);
            }
        }
        if (const size_t at = q.find("name = '"); at != std::string::npos)
        {
            size_t index = at + 8u;
            while (index < q.size())
            {
                if (q[index] == '\\' && index + 1u < q.size())
                {
                    nameFilter.push_back(q[index + 1u]);
                    index += 2u;
                    continue;
                }
                if (q[index] == '\'')
                {
                    break;
                }
                nameFilter.push_back(q[index]);
                ++index;
            }
            hasNameFilter = true;
        }
        const bool excludeTrashed = q.find("trashed = false") != std::string::npos;
        size_t pageSize           = 100u;
        if (const std::string pageText = request.Query("pageSize"); ! pageText.empty())
        {
            static_cast<void>(std::from_chars(pageText.data(), pageText.data() + pageText.size(), pageSize));
            pageSize = std::clamp<size_t>(pageSize, 1u, 1000u);
        }
        size_t start = 0u;
        if (const std::string token = request.Query("pageToken"); ! token.empty())
        {
            static_cast<void>(std::from_chars(token.data(), token.data() + token.size(), start));
        }
        std::vector<const DriveFile*> matches;
        for (const DriveFile& file : _files)
        {
            if (file.id == "root" || (excludeTrashed && file.trashed))
            {
                continue;
            }
            if (! parentId.empty() && ! HasParent(file, parentId))
            {
                continue;
            }
            if (hasNameFilter && file.name != nameFilter)
            {
                continue;
            }
            matches.push_back(&file);
        }
        std::string json = R"({"kind":"drive#fileList","files":[)";
        bool first       = true;
        size_t index     = start;
        for (; index < matches.size() && index < start + pageSize; ++index)
        {
            if (! first)
            {
                json.push_back(',');
            }
            first = false;
            json += FileJsonLocked(*matches[index]);
        }
        json += "]";
        if (index < matches.size())
        {
            json += std::format(R"(,"nextPageToken":"{}")", index);
        }
        json += "}";
        LoopbackHttpResponse response = JsonResponse(200, std::move(json));
        response.dripBody             = true;
        return response;
    }

    [[nodiscard]] LoopbackHttpResponse HandleCreateLocked(const LoopbackHttpRequest& request)
    {
        const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
        const std::string name = FindJsonString(body, "name");
        if (name.empty())
        {
            return ErrorResponse(400, "invalid", "name required");
        }
        std::string mimeType = FindJsonString(body, "mimeType");
        if (mimeType.empty())
        {
            mimeType = "application/octet-stream";
        }
        std::vector<std::string> parents = FindJsonStringArray(body, "parents");
        if (parents.empty())
        {
            parents.push_back("root");
        }
        for (const std::string& parent : parents)
        {
            const DriveFile* parentFile = FindByIdLocked(parent);
            if (parentFile == nullptr || parentFile->mimeType != kFolderMime)
            {
                return ErrorResponse(404, "notFound", "Parent not found: " + parent);
            }
        }
        DriveFile& created = AddLocked(name, parents.front(), mimeType == kFolderMime);
        created.mimeType   = mimeType;
        created.parents    = parents;
        return JsonResponse(200, FileJsonLocked(created));
    }

    [[nodiscard]] LoopbackHttpResponse HandleStartUploadLocked(const LoopbackHttpRequest& request, std::string_view existingId)
    {
        if (request.Query("uploadType") != "resumable")
        {
            return ErrorResponse(400, "invalid", "only resumable uploads are modeled");
        }
        UploadSession session;
        if (! existingId.empty())
        {
            const DriveFile* existing = FindByIdLocked(existingId);
            if (existing == nullptr)
            {
                return ErrorResponse(404, "notFound", "File not found");
            }
            session.fileId = existing->id;
        }
        else
        {
            const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
            session.name    = FindJsonString(body, "name");
            session.parents = FindJsonStringArray(body, "parents");
            if (session.name.empty())
            {
                return ErrorResponse(400, "invalid", "name required");
            }
            if (session.parents.empty())
            {
                session.parents.push_back("root");
            }
            for (const std::string& parent : session.parents)
            {
                const DriveFile* parentFile = FindByIdLocked(parent);
                if (parentFile == nullptr || parentFile->mimeType != kFolderMime)
                {
                    return ErrorResponse(404, "notFound", "Parent not found: " + parent);
                }
            }
        }
        if (const std::string* total = request.Header("x-upload-content-length"); total != nullptr)
        {
            static_cast<void>(std::from_chars(total->data(), total->data() + total->size(), session.total));
        }
        const std::string sessionId = std::format("session-{}", ++_nextId);
        _uploads.insert_or_assign(sessionId, std::move(session));
        LoopbackHttpResponse response = EmptyResponse(200);
        response.headers.emplace_back("Location", _origin + "/upload/session/" + sessionId);
        return response;
    }

    [[nodiscard]] LoopbackHttpResponse HandleUploadSessionLocked(const LoopbackHttpRequest& request, std::string_view sessionId)
    {
        const auto session = _uploads.find(std::string(sessionId));
        if (session == _uploads.end())
        {
            return ErrorResponse(404, "notFound", "unknown upload session");
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
        UploadSession& upload           = session->second;
        const std::string* contentRange = request.Header("content-range");
        if (contentRange == nullptr || ! contentRange->starts_with("bytes "))
        {
            return ErrorResponse(400, "invalid", "Content-Range required");
        }
        const std::string_view spec(std::string_view(*contentRange).substr(6u));
        // An empty object completes with `bytes */0` and no body.
        const bool emptyCompletion = spec == "*/0" && upload.total == 0u && request.body.empty();
        if (spec.starts_with("*/") && ! emptyCompletion)
        {
            // Status query: report what has been received.
            if (upload.bytes.empty())
            {
                return EmptyResponse(308);
            }
            LoopbackHttpResponse status = EmptyResponse(308);
            status.headers.emplace_back("Range", std::format("bytes=0-{}", upload.bytes.size() - 1u));
            return status;
        }
        if (! emptyCompletion)
        {
            uint64_t first     = 0u;
            uint64_t last      = 0u;
            uint64_t total     = 0u;
            const size_t dash  = spec.find('-');
            const size_t slash = spec.find('/');
            if (dash == std::string_view::npos || slash == std::string_view::npos || slash < dash)
            {
                return ErrorResponse(400, "invalid", "bad Content-Range");
            }
            static_cast<void>(std::from_chars(spec.data(), spec.data() + dash, first));
            static_cast<void>(std::from_chars(spec.data() + dash + 1u, spec.data() + slash, last));
            if (spec.substr(slash + 1u) != "*")
            {
                static_cast<void>(std::from_chars(spec.data() + slash + 1u, spec.data() + spec.size(), total));
                if (upload.total == 0u)
                {
                    upload.total = total;
                }
            }
            if (first != upload.bytes.size() || last + 1u - first != request.body.size() || (total != 0u && total != upload.total))
            {
                return ErrorResponse(416, "invalid", "unexpected range");
            }
            upload.bytes.insert(upload.bytes.end(), request.body.begin(), request.body.end());
        }
        if (! emptyCompletion && (upload.total == 0u || upload.bytes.size() < upload.total))
        {
            LoopbackHttpResponse partial = EmptyResponse(308);
            partial.headers.emplace_back("Range", std::format("bytes=0-{}", upload.bytes.size() - 1u));
            return partial;
        }
        DriveFile* target = nullptr;
        if (! upload.fileId.empty())
        {
            target = FindByIdLocked(upload.fileId);
            if (target == nullptr)
            {
                _uploads.erase(session);
                return ErrorResponse(404, "notFound", "File vanished");
            }
        }
        else
        {
            target          = &AddLocked(upload.name, upload.parents.front(), false);
            target->parents = upload.parents;
        }
        target->bytes = std::move(upload.bytes);
        Touch(*target);
        const std::string json = FileJsonLocked(*target);
        _uploads.erase(session);
        return JsonResponse(200, json);
    }

    [[nodiscard]] LoopbackHttpResponse HandlePatchLocked(const LoopbackHttpRequest& request, DriveFile& file)
    {
        const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
        if (const std::string name = FindJsonString(body, "name"); ! name.empty())
        {
            file.name = name;
        }
        if (const std::optional<bool> trashed = FindJsonBool(body, "trashed"); trashed.has_value())
        {
            file.trashed = trashed.value();
        }
        if (const std::string add = request.Query("addParents"); ! add.empty())
        {
            const DriveFile* parent = FindByIdLocked(add);
            if (parent == nullptr || parent->mimeType != kFolderMime)
            {
                return ErrorResponse(404, "notFound", "Parent not found: " + add);
            }
            if (! HasParent(file, add))
            {
                file.parents.push_back(add);
            }
        }
        if (const std::string remove = request.Query("removeParents"); ! remove.empty())
        {
            std::erase(file.parents, remove);
        }
        Touch(file);
        return JsonResponse(200, FileJsonLocked(file));
    }

    [[nodiscard]] LoopbackHttpResponse HandleCopyLocked(const LoopbackHttpRequest& request, const DriveFile& source)
    {
        ++_copyRequests;
        if (source.mimeType == kFolderMime)
        {
            return ErrorResponse(400, "invalid", "folders cannot be copied");
        }
        const std::string body(reinterpret_cast<const char*>(request.body.data()), request.body.size());
        std::string name = FindJsonString(body, "name");
        if (name.empty())
        {
            name = "Copy of " + source.name;
        }
        std::vector<std::string> parents = FindJsonStringArray(body, "parents");
        if (parents.empty())
        {
            parents = source.parents;
        }
        // AddLocked may reallocate _files and invalidate the source reference.
        const std::string mimeType       = source.mimeType;
        const std::vector<uint8_t> bytes = source.bytes;
        DriveFile& copy                  = AddLocked(name, parents.front(), false);
        copy.parents                     = parents;
        copy.mimeType                    = mimeType;
        copy.bytes                       = bytes;
        ++_copyCommits;
        if (_failCopyAfterCommit)
        {
            // The object exists before the error is delivered: replay would create a duplicate.
            return ErrorResponse(503, "backendError", "Response lost after copy committed");
        }
        return JsonResponse(200, FileJsonLocked(copy));
    }

    [[nodiscard]] LoopbackHttpResponse HandleDownloadLocked(const LoopbackHttpRequest& request, const DriveFile& file)
    {
        if (file.mimeType == kFolderMime)
        {
            return ErrorResponse(403, "fileNotDownloadable", "Folders have no content.");
        }
        const size_t size = file.bytes.size();
        LoopbackHttpResponse response;
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
                LoopbackHttpResponse unsatisfiable = ErrorResponse(416, "invalid", "range not satisfiable");
                unsatisfiable.headers.emplace_back("Content-Range", std::format("bytes */{}", size));
                return unsatisfiable;
            }
            last            = (std::min)(last, size - 1u);
            response.status = 206;
            response.headers.emplace_back("Content-Range", std::format("bytes {}-{}/{}", first, last, size));
            response.body.assign(file.bytes.begin() + static_cast<ptrdiff_t>(first), file.bytes.begin() + static_cast<ptrdiff_t>(last + 1u));
            return response;
        }
        response.status = 200;
        response.body   = file.bytes;
        return response;
    }

    mutable std::mutex _mutex;
    std::string _origin;
    std::vector<DriveFile> _files;
    std::map<std::string, UploadSession> _uploads;
    unsigned int _nextId         = 0u;
    unsigned int _nextVersion    = 1u;
    bool _failCopyAfterCommit    = false;
    unsigned int _copyRequests   = 0u;
    unsigned int _copyCommits    = 0u;
    bool _failDeleteAfterCommit  = false;
    unsigned int _deleteRequests = 0u;
    unsigned int _deleteCommits  = 0u;
};

class FakeDriveEndpoint final
{
public:
    FakeDriveEndpoint()
    {
        _server = std::make_unique<LoopbackHttpServer>([this](const LoopbackHttpRequest& request) { return _drive.Handle(request); });
    }
    FakeDriveEndpoint(const FakeDriveEndpoint&)            = delete;
    FakeDriveEndpoint& operator=(const FakeDriveEndpoint&) = delete;
    FakeDriveEndpoint(FakeDriveEndpoint&&)                 = delete;
    FakeDriveEndpoint& operator=(FakeDriveEndpoint&&)      = delete;

    [[nodiscard]] HRESULT Start()
    {
        const HRESULT hr = _server->Start();
        if (FAILED(hr))
        {
            return hr;
        }
        _origin = std::format("http://127.0.0.1:{}", _server->Port());
        _drive.SetOrigin(_origin);
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
    [[nodiscard]] std::wstring Origin() const
    {
        return Common::SelfTest::LoopbackHttpUtf16FromUtf8(_origin);
    }
    [[nodiscard]] FakeDrive& Drive() noexcept
    {
        return _drive;
    }
    [[nodiscard]] LoopbackHttpServer& Server() noexcept
    {
        return *_server;
    }

private:
    FakeDrive _drive;
    std::unique_ptr<LoopbackHttpServer> _server;
    std::string _origin;
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

// R0f-GDrive witness. (1) Bytes still flowing: a non-recursive Delete of a folder whose children
// listing the fixture drips returns ERROR_CANCELLED within a few progress callbacks of Cancel and
// deletes nothing. (2) A server that stopped answering: Cancel returns within a couple of progress
// ticks with the host's verdict (libcurl invokes the progress callback at least once per second even
// while a request waits), and (3) without Cancel the provider-owned bound (the hard request
// timeout) returns it on its own as a transport failure; the file survives.
void RunDriveStalledRequestCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    try
    {
        FakeDriveEndpoint endpoint;
        endpoint.Drive().SeedFile("/stall.txt", std::vector<uint8_t>(64u, 0x5Au));
        for (unsigned int index = 0u; index < 60u; ++index)
        {
            endpoint.Drive().SeedFile(std::format("/drip/object-{:03}.bin", index), std::vector<uint8_t>(16u, static_cast<uint8_t>(index)));
        }
        constexpr unsigned int kDeepTreeLevels = 80u;
        std::string deepDirectoryPath          = "/deep";
        for (unsigned int level = 0u; level < kDeepTreeLevels; ++level)
        {
            deepDirectoryPath += std::format("/level-{:03}", level);
        }
        const std::string deepFilePath = deepDirectoryPath + "/payload.bin";
        endpoint.Drive().SeedFile(deepFilePath, std::vector<uint8_t>{1u, 2u, 3u, 4u, 5u, 6u, 7u});
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"Google Drive stalled-request fixture should start", passed, failed))
        {
            return;
        }
        const std::wstring origin = endpoint.Origin();
        ConfigureFakeDrive(origin.c_str(), true);
        auto restore = wil::scope_exit([&]() noexcept
        {
            ConfigureFakeDrive(nullptr, false);
            endpoint.Stop();
        });

        wil::com_ptr<FileSystemGoogleDrive> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemGoogleDrive(nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem), L"Google Drive stalled-request proof should create an instance", passed, failed))
        {
            return;
        }
        constexpr uint32_t kConnectTimeoutMs = 2'000u;
        constexpr uint32_t kRequestTimeoutMs = 3'000u;
        HRESULT hr                           = fileSystem->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":3000})");
        if (! DebugCheck(SUCCEEDED(hr), L"Google Drive instance should accept the fixture configuration", passed, failed))
        {
            return;
        }
        const unsigned long boundMs = FileSystemGoogleDriveInternal::DriveProviderWatchdogTimeoutMs(kConnectTimeoutMs, kRequestTimeoutMs);
        DebugCheck(boundMs > 0u && boundMs <= 20'000u,
                   L"the Google Drive provider-owned bound must be a small nonzero value for the fixture timeouts",
                   passed,
                   failed);

        const auto deleteWithControl = [&](const wchar_t* path, FileSystemFlags flags, CancelControl* control, std::atomic<HRESULT>& result) noexcept
        {
            FileSystemOptions options{};
            options.sizeBytes        = sizeof(options);
            options.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
            options.operationControl = control;
            result.store(fileSystem->DeleteItem(path, flags, control != nullptr ? &options : nullptr, nullptr, nullptr), std::memory_order_release);
        };

        // 1) Bytes flowing: the children listing drips; Cancel lands between progress callbacks.
        endpoint.Server().SetDrip(512u, 40u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { deleteWithControl(L"/@conn:google-drive-selftest/drip", FILESYSTEM_FLAG_NONE, &control, deleteHr); });
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
            std::fwprintf(stderr, L"[Google Drive] streaming request returned %llu ms after Cancel\n", static_cast<unsigned long long>(cancelMs));
            DebugCheck(listingReached, L"the folder Delete must list the children on the fixture", passed, failed);
            DebugCheck(returned, L"a request whose body is still arriving must return after Cancel", passed, failed);
            DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"a canceled streaming request must report ERROR_CANCELLED",
                       passed,
                       failed);
            DebugCheck(cancelMs < 3'000u, L"Cancel must return a streaming request within a few progress callbacks", passed, failed);
            DebugCheck(endpoint.Server().RequestCount("DELETE") == 0u, L"a Delete canceled while listing must not delete anything", passed, failed);
            if (! returned)
            {
                return;
            }
        }
        endpoint.Server().SetDrip(0u, 0u);

        // 2) Silent server + Cancel: libcurl's progress callback polls the control every second.
        endpoint.Server().SetStallMethod("DELETE", 60'000u);
        {
            CancelControl control;
            std::atomic<HRESULT> deleteHr{E_PENDING};
            std::thread deleter([&]() noexcept { deleteWithControl(L"/@conn:google-drive-selftest/stall.txt", FILESYSTEM_FLAG_NONE, &control, deleteHr); });
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
            std::fwprintf(stderr,
                          L"[Google Drive] stalled request returned %llu ms after Cancel (declared bound %lu ms)\n",
                          static_cast<unsigned long long>(cancelMs),
                          boundMs);
            DebugCheck(deleteReached, L"the stalled DELETE must reach the fixture", passed, failed);
            DebugCheck(returned, L"a request the server never answers must return after Cancel", passed, failed);
            DebugCheck(deleteHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"a canceled stalled request must report ERROR_CANCELLED",
                       passed,
                       failed);
            DebugCheck(cancelMs < 3'000u, L"Cancel on a silent server must return within a few progress callbacks", passed, failed);
            if (! returned)
            {
                return;
            }
        }

        // 3) Silent server, no Cancel: the hard request timeout returns it on its own.
        {
            std::atomic<HRESULT> deleteHr{E_PENDING};
            const ULONGLONG boundStart = GetTickCount64();
            deleteWithControl(L"/@conn:google-drive-selftest/stall.txt", FILESYSTEM_FLAG_NONE, nullptr, deleteHr);
            const ULONGLONG elapsedMs = GetTickCount64() - boundStart;
            const HRESULT boundHr     = deleteHr.load(std::memory_order_acquire);
            std::fwprintf(stderr,
                          L"[Google Drive] stalled request returned on its own after %llu ms (declared bound %lu ms)\n",
                          static_cast<unsigned long long>(elapsedMs),
                          boundMs);
            DebugCheck(FAILED(boundHr) && boundHr != HRESULT_FROM_WIN32(ERROR_CANCELLED),
                       L"an un-canceled stalled request must fail through the transport bound, not as a cancel",
                       passed,
                       failed);
            DebugCheck(elapsedMs <= boundMs + 5'000u, L"the provider-owned bound must return a stalled request on its own", passed, failed);
        }
        endpoint.Server().SetStallMethod({}, 0u);
        DebugCheck(endpoint.Drive().Exists("/stall.txt"), L"a stalled DELETE the client gave up on must not be reported as deleted", passed, failed);

        FileSystemDirectorySizeResult deepSize{};
        deepSize.sizeBytes       = sizeof(deepSize);
        const HRESULT deepSizeHr = fileSystem->GetDirectorySize(L"/@conn:google-drive-selftest/deep", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, &deepSize);
        DebugCheck(SUCCEEDED(deepSizeHr) && deepSize.status == S_OK && deepSize.directoryCount == kDeepTreeLevels && deepSize.fileCount == 1u &&
                       deepSize.totalBytes == 7u,
                   L"recursive Google Drive directory size must include every level beyond the former depth-64 cutoff",
                   passed,
                   failed);

        const ULONGLONG deepCopyStart = GetTickCount64();
        const HRESULT deepCopyHr      = fileSystem->CopyItem(
            L"/@conn:google-drive-selftest/deep", L"/@conn:google-drive-selftest/deep-copy", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr);
        const uint64_t deepCopyMs            = GetTickCount64() - deepCopyStart;
        const std::string copiedDeepFilePath = "/deep-copy" + deepFilePath.substr(std::string_view("/deep").size());
        DebugCheck(SUCCEEDED(deepCopyHr) && endpoint.Drive().Exists(copiedDeepFilePath),
                   L"provider-native Google Drive copy must complete an iterative tree deeper than 64 levels",
                   passed,
                   failed);
        std::fwprintf(stderr,
                      L"[Google Drive] iterative deep-tree copy: %u levels in %llu ms (hr=0x%08X)\n",
                      kDeepTreeLevels,
                      static_cast<unsigned long long>(deepCopyMs),
                      static_cast<unsigned int>(deepCopyHr));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemGoogleDrive stalled-request selftest failed after std::exception.");
        DebugCheck(false, L"Google Drive stalled-request proof should not throw std::exception", passed, failed);
    }
}
// R0-RC4: SetExpectedReplacement with a zero host timestamp is refused instead of replacing the
// object the writer observed when it opened.
void RunDriveZeroTimestampReplaceSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    try
    {
        FakeDriveEndpoint endpoint;
        endpoint.Drive().SeedFile("/occupant.txt", std::vector<uint8_t>(64u, 0x5Au));
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"R0-RC4 Google Drive fixture should start", passed, failed))
        {
            return;
        }
        const std::wstring origin = endpoint.Origin();
        ConfigureFakeDrive(origin.c_str(), true);
        auto restore = wil::scope_exit([&]() noexcept
        {
            ConfigureFakeDrive(nullptr, false);
            endpoint.Stop();
        });

        wil::com_ptr<FileSystemGoogleDrive> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemGoogleDrive(nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem), L"R0-RC4 Google Drive proof should create an instance", passed, failed))
        {
            return;
        }
        HRESULT hr = fileSystem->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":3000})");
        if (! DebugCheck(SUCCEEDED(hr), L"R0-RC4 Google Drive instance should accept the fixture configuration", passed, failed))
        {
            return;
        }

        wil::com_ptr<IFileWriter> writer;
        hr = fileSystem->CreateFileWriter(L"/@conn:google-drive-selftest/occupant.txt", FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.addressof());
        if (! DebugCheck(SUCCEEDED(hr) && writer, L"R0-RC4: Google Drive should open an overwrite writer on the seeded occupant", passed, failed))
        {
            return;
        }
        wil::com_ptr<IFileWriterExpectedReplacement> replacement;
        hr = writer->QueryInterface(IID_PPV_ARGS(replacement.addressof()));
        if (! DebugCheck(SUCCEEDED(hr) && replacement, L"R0-RC4: the Google Drive writer should expose the expected-replacement contract", passed, failed))
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
                   L"R0-RC4: a Google Drive identity-less replace with a zero host timestamp must be refused before upload",
                   passed,
                   failed);
        DebugCheck(endpoint.Drive().Exists("/occupant.txt"), L"R0-RC4: the Google Drive occupant must survive the refused replace", passed, failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemGoogleDrive zero-timestamp replace selftest failed after std::exception.");
        DebugCheck(false, L"R0-RC4 Google Drive proof should not throw std::exception", passed, failed);
    }
}

} // namespace FileSystemGoogleDriveSelfTest

// Host self-tests drive the fake Drive endpoint through these exports: start one on a loopback port
// (the plugin's Drive and token endpoints are redirected to it and `/@conn:google-drive-selftest/`
// resolves to a synthetic connection), run File Operations, then stop it.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderGoogleDriveStartFakeDriveForSelfTest(unsigned int* port, void** endpoint) noexcept
{
    if (port == nullptr || endpoint == nullptr)
    {
        return E_POINTER;
    }
    *port     = 0u;
    *endpoint = nullptr;
    try
    {
        auto fixture     = std::make_unique<FileSystemGoogleDriveSelfTest::FakeDriveEndpoint>();
        const HRESULT hr = fixture->Start();
        if (FAILED(hr))
        {
            return hr;
        }
        const std::wstring origin = fixture->Origin();
        FileSystemGoogleDriveSelfTest::ConfigureFakeDrive(origin.c_str(), true);
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

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderGoogleDriveFakeDriveRequestLogForSelfTest(void* endpoint,
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
        const std::wstring text = static_cast<FileSystemGoogleDriveSelfTest::FakeDriveEndpoint*>(endpoint)->Server().RequestLog(48u);
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

extern "C" __declspec(dllexport) void __stdcall RedSalamanderGoogleDriveStopFakeDriveForSelfTest(void* endpoint) noexcept
{
    auto* fixture = static_cast<FileSystemGoogleDriveSelfTest::FakeDriveEndpoint*>(endpoint);
    if (fixture == nullptr)
    {
        return;
    }
    FileSystemGoogleDriveSelfTest::ConfigureFakeDrive(nullptr, false);
    fixture->Stop();
    delete fixture;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderGoogleDriveMutationCommitFailureForSelfTest(
    void* endpoint, FileSystemOperation operation, BOOL arm, unsigned int* requests, unsigned int* commits) noexcept
{
    if (endpoint == nullptr || requests == nullptr || commits == nullptr)
    {
        return E_POINTER;
    }
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_DELETE)
    {
        return E_INVALIDARG;
    }
    // Test-export ABI boundary: contain lock errors; allocation failure remains fatal.
    try
    {
        static_cast<FileSystemGoogleDriveSelfTest::FakeDriveEndpoint*>(endpoint)->Drive().MutationCommitFailureSnapshot(
            operation, arm != FALSE, *requests, *commits);
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive mutation-commit fault snapshot failed.");
        return E_FAIL;
    }
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderGoogleDriveR0fContainmentSelfTests(unsigned int* passed, unsigned int* failed) noexcept
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    FileSystemGoogleDriveSelfTest::RunDriveStalledRequestCancelSelfTests(*passed, *failed);
    return *failed == 0u ? S_OK : E_FAIL;
}

#else

static_assert(true);

#endif

#pragma once

#include <cstdint>
#include <cstddef>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace FileSystemCurlInternal
{
struct ImapMessageSummary
{
    uint64_t uid       = 0;
    uint64_t sizeBytes = 0;
    bool flagged       = false;
    bool seen          = false;
    bool deleted       = false;
    __int64 sentTime   = 0;
    __int64 recvTime   = 0;
    std::wstring subject;
    std::wstring from;
};

struct ImapMailboxStatus
{
    std::optional<uint64_t> messages;
    std::optional<uint64_t> recent;
    std::optional<uint64_t> uidNext;
    std::optional<uint64_t> uidValidity;
    std::optional<uint64_t> unseen;
};

struct ImapUidBatchRange
{
    size_t offset = 0;
    size_t count  = 0;
};

[[nodiscard]] bool TryParseImapUidFromLeafName(std::wstring_view leafName, uint64_t& outUid) noexcept;
[[nodiscard]] std::wstring DecodeRfc2047EncodedWordsToUtf16(std::string_view headerValue) noexcept;
[[nodiscard]] std::wstring BuildImapMessageLeafName(std::wstring_view subject, std::wstring_view from, uint64_t uid) noexcept;
[[nodiscard]] bool TryParseImapMailboxStatus(std::string_view response, ImapMailboxStatus& out) noexcept;
[[nodiscard]] std::vector<ImapUidBatchRange> BuildImapUidBatchRanges(size_t uidCount, size_t maxBatchSize);
} // namespace FileSystemCurlInternal

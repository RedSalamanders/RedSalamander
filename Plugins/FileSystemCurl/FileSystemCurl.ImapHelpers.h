#pragma once

#include <cstddef>
#include <cstdint>
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

struct ImapMessageIdentity
{
    uint64_t uidValidity = 0;
    uint64_t uid         = 0;
};

struct ImapCapabilities
{
    bool uidPlus = false;
};

enum class ImapDeleteCommand
{
    AddDeletedFlag,
    UidExpunge,
    RemoveDeletedFlag,
};

using ImapDeleteCommandExecutor = long (*)(void* context, ImapDeleteCommand command, uint64_t uid) noexcept;

struct ImapDeleteOutcome
{
    bool targetMarkedDeleted = false;
    bool rollbackAttempted   = false;
    long rollbackHr          = 0;
};

struct ImapUidBatchRange
{
    size_t offset = 0;
    size_t count  = 0;
};

[[nodiscard]] bool TryParseImapUidFromLeafName(std::wstring_view leafName, uint64_t& outUid) noexcept;
[[nodiscard]] bool TryParseImapMessageIdentityFromLeafName(std::wstring_view leafName, ImapMessageIdentity& outIdentity) noexcept;
[[nodiscard]] std::wstring DecodeRfc2047EncodedWordsToUtf16(std::string_view headerValue) noexcept;
[[nodiscard]] std::wstring BuildImapMessageLeafName(std::wstring_view subject, std::wstring_view from, uint64_t uidValidity, uint64_t uid) noexcept;
[[nodiscard]] bool TryParseImapMailboxStatus(std::string_view response, ImapMailboxStatus& out) noexcept;
[[nodiscard]] bool TryParseImapCapabilities(std::string_view response, ImapCapabilities& out) noexcept;
[[nodiscard]] long ValidateImapMessageUidValidity(uint64_t expectedUidValidity, const std::optional<uint64_t>& observedUidValidity) noexcept;
[[nodiscard]] long ExecuteImapSingleMessageDelete(
    bool uidPlusAvailable, uint64_t uid, ImapDeleteCommandExecutor executor, void* context, ImapDeleteOutcome& outOutcome) noexcept;
[[nodiscard]] std::vector<ImapUidBatchRange> BuildImapUidBatchRanges(size_t uidCount, size_t maxBatchSize);
[[nodiscard]] size_t ResolveImapSummaryRepairFetchBudget(size_t requestedUidCount) noexcept;
} // namespace FileSystemCurlInternal

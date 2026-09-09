#include "FileSystemCurl/FileSystemCurl.ImapHelpers.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#include <chrono>
#include <format>
#include <iostream>
#include <string>
#include <string_view>
#include <vector>

namespace
{
[[nodiscard]] bool Require(bool condition, std::wstring_view message)
{
    if (! condition)
    {
        std::wcerr << message << L'\n';
        return false;
    }

    return true;
}

[[nodiscard]] std::wstring HexCodeUnits(std::wstring_view text)
{
    std::wstring out;
    for (const wchar_t ch : text)
    {
        if (! out.empty())
        {
            out.push_back(L' ');
        }
        out.append(std::format(L"{:04X}", static_cast<unsigned int>(ch)));
    }
    return out;
}

[[nodiscard]] bool RequireEqual(std::wstring_view actual, std::wstring_view expected, std::wstring_view message)
{
    if (actual == expected)
    {
        return true;
    }

    std::wcerr << message << L'\n';
    std::wcerr << L"  actual:   " << HexCodeUnits(actual) << L'\n';
    std::wcerr << L"  expected: " << HexCodeUnits(expected) << L'\n';
    return false;
}

void AppendCodePoint(std::wstring& out, uint32_t codePoint)
{
    if (codePoint <= 0xFFFFu)
    {
        out.push_back(static_cast<wchar_t>(codePoint));
        return;
    }

    codePoint -= 0x10000u;
    out.push_back(static_cast<wchar_t>(0xD800u + (codePoint >> 10u)));
    out.push_back(static_cast<wchar_t>(0xDC00u + (codePoint & 0x3FFu)));
}

[[nodiscard]] bool TestImapMessageLeafNamePreferredFormat()
{
    bool ok = true;

    ok = Require(FileSystemCurlInternal::BuildImapMessageLeafName(L"Quarterly report", L"boss@example.test", 777u, 12345u) ==
                     L"Quarterly report [777-12345].eml",
                 L"IMAP message leaf name should use '<subject> [uidValidity-uid].eml'.") &&
         ok;
    ok = Require(FileSystemCurlInternal::BuildImapMessageLeafName(L"", L"", 777u, 7u) == L"(no subject) [777-7].eml",
                 L"Empty IMAP subjects should use a stable message fallback.") &&
         ok;
    ok = Require(FileSystemCurlInternal::BuildImapMessageLeafName(L"Q4: plan/report", L"boss@example.test", 777u, 12u) == L"Q4_ plan_report [777-12].eml",
                 L"IMAP subject display text should be Windows-filename safe.") &&
         ok;

    const std::wstring longSubject(100u, L'A');
    const std::wstring expectedLong = std::wstring(93u, L'A') + L"... [777-42].eml";
    ok                              = Require(FileSystemCurlInternal::BuildImapMessageLeafName(longSubject, L"boss@example.test", 777u, 42u) == expectedLong,
                                              L"Long IMAP subjects should truncate with ASCII ellipsis before the uid suffix.") &&
                                      ok;

    return ok;
}

[[nodiscard]] bool TestImapSubjectDecoderHandlesMimeEncodedWords()
{
    bool ok = true;

    std::wstring weekly;
    AppendCodePoint(weekly, 0x1F389u);
    weekly.append(L"\x3010Weekly Trends\x3011");
    AppendCodePoint(weekly, 0x1F442u);
    AppendCodePoint(weekly, 0x1F916u);
    weekly.append(L" Add Voice");
    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16(
                          "=?UTF-8?Q?=F0=9F=8E=89=E3=80=90Weekly=20Trends=E3=80=91=F0=9F=91=82=F0=9F=A4=96=20Add=20Voice?="),
                      weekly,
                      L"IMAP subject decoder should decode UTF-8 Q encoded emoji subjects.") &&
         ok;

    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16(
                          "=?US-ASCII?Q?=5B2603=2E19312=5D_LeWorldModel=3A_Stable_End-to-End_Joi?=nt-Embedding Predicti"),
                      L"[2603.19312] LeWorldModel: Stable End-to-End Joint-Embedding Predicti",
                      L"IMAP subject decoder should join encoded words with adjacent plain text.") &&
         ok;

    std::wstring japanese;
    AppendCodePoint(japanese, 0x3053u);
    AppendCodePoint(japanese, 0x3093u);
    AppendCodePoint(japanese, 0x306Bu);
    AppendCodePoint(japanese, 0x3061u);
    AppendCodePoint(japanese, 0x306Fu);
    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16("=?UTF-8?B?44GT44KT44Gr44Gh44Gv?="),
                      japanese,
                      L"IMAP subject decoder should decode UTF-8 base64 encoded words.") &&
         ok;

    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16("=?windows-1251?Q?=CF=F0=E8=E2=E5=F2?="),
                      L"\x041F\x0440\x0438\x0432\x0435\x0442",
                      L"IMAP subject decoder should decode non-UTF Windows code page subjects.") &&
         ok;

    std::wstring pitch = L"[PITCH IMMO    FONCIER IMMO]   ";
    AppendCodePoint(pitch, 0x1F64Bu);
    AppendCodePoint(pitch, 0x200Du);
    AppendCodePoint(pitch, 0x2642u);
    AppendCodePoint(pitch, 0xFE0Fu);
    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16(
                          "=?UTF-8?Q?=5BPITCH_IMMO____FONCIER_IMMO=5D___?= =?UTF-8?Q?=F0=9F=99=8B=E2=80=8D=E2=99=82=EF=B8=8F?="),
                      pitch,
                      L"IMAP subject decoder should join adjacent encoded words without separator whitespace.") &&
         ok;

    ok = RequireEqual(FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16("25844 =_utf-8_q_=C2=AD_CHATENAY_MALABRY_=C2=AD_Appel_de_fonds"),
                      L"25844 - CHATENAY MALABRY - Appel de fonds",
                      L"IMAP subject decoder should recover malformed sanitized encoded-word fragments and normalize soft hyphens.") &&
         ok;

    return ok;
}

[[nodiscard]] bool TestImapUidParserAcceptsPreferredAndDirectNames()
{
    bool ok      = true;
    uint64_t uid = 0;

    ok = Require(FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report [777-12345].eml", uid) && uid == 12345u,
                 L"IMAP UID parser should read the UID from the UIDVALIDITY-qualified preferred filename.") &&
         ok;
    ok = Require(FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly [draft] report [987].EML", uid) && uid == 987u,
                 L"IMAP UID parser should scan from the end and tolerate earlier brackets and upper-case extension.") &&
         ok;
    ok = Require(FileSystemCurlInternal::TryParseImapUidFromLeafName(L"12345.eml", uid) && uid == 12345u,
                 L"IMAP UID parser should keep direct numeric leaf names for compatibility.") &&
         ok;

    return ok;
}

[[nodiscard]] bool TestImapQualifiedIdentityParser()
{
    bool ok = true;
    FileSystemCurlInternal::ImapMessageIdentity identity;

    ok = Require(FileSystemCurlInternal::TryParseImapMessageIdentityFromLeafName(L"Quarterly [draft] [777-12345].EML", identity) &&
                     identity.uidValidity == 777u && identity.uid == 12345u,
                 L"IMAP message identity parser should recover UIDVALIDITY and UID from the final bracketed suffix.") &&
         ok;
    ok = Require(! FileSystemCurlInternal::TryParseImapMessageIdentityFromLeafName(L"Quarterly [12345].eml", identity),
                 L"IMAP message identity parser should reject legacy UID-only filenames for version-sensitive operations.") &&
         ok;
    ok = Require(! FileSystemCurlInternal::TryParseImapMessageIdentityFromLeafName(L"Quarterly [0-12345].eml", identity),
                 L"IMAP message identity parser should reject zero UIDVALIDITY.") &&
         ok;
    ok = Require(! FileSystemCurlInternal::TryParseImapMessageIdentityFromLeafName(L"Quarterly [777-0].eml", identity),
                 L"IMAP message identity parser should reject zero UID.") &&
         ok;

    return ok;
}

[[nodiscard]] bool TestImapUidParserRejectsMalformedNames()
{
    bool ok      = true;
    uint64_t uid = 0;

    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report [abc].eml", uid),
                                        L"IMAP UID parser should reject non-numeric bracketed ids.") &&
                                ok;
    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report [].eml", uid),
                                        L"IMAP UID parser should reject empty bracketed ids.") &&
                                ok;
    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report [12345] draft.eml", uid),
                                        L"IMAP UID parser should require the bracketed id immediately before .eml.") &&
                                ok;
    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report 12345.eml", uid),
                                        L"IMAP UID parser should reject ambiguous subject-plus-trailing-digits names.") &&
                                ok;
    const wchar_t separator   = static_cast<wchar_t>(0xFF5C);
    const std::wstring legacy = std::wstring(L"Quarterly report") + separator + L"boss@example.test" + separator + L"4321.eml";
    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(legacy, uid),
                                        L"IMAP UID parser should reject the retired legacy decorated name shape.") &&
                                ok;
    ok                        = Require(! FileSystemCurlInternal::TryParseImapUidFromLeafName(L"Quarterly report [12345].txt", uid),
                                        L"IMAP UID parser should require the .eml extension.") &&
                                ok;

    return ok;
}

[[nodiscard]] bool TestImapMailboxStatusParser()
{
    FileSystemCurlInternal::ImapMailboxStatus status;
    bool ok = true;

    ok = Require(FileSystemCurlInternal::TryParseImapMailboxStatus(
                     "* STATUS \"INBOX\" (MESSAGES 42 RECENT 2 UIDNEXT 9001 UIDVALIDITY 777 UNSEEN 5)\r\nA OK STATUS completed\r\n", status),
                 L"IMAP STATUS parser should recognize a normal STATUS response.") &&
         ok;
    ok = Require(status.messages.has_value() && status.messages.value() == 42u, L"IMAP STATUS parser should read MESSAGES.") && ok;
    ok = Require(status.recent.has_value() && status.recent.value() == 2u, L"IMAP STATUS parser should read RECENT.") && ok;
    ok = Require(status.uidNext.has_value() && status.uidNext.value() == 9001u, L"IMAP STATUS parser should read UIDNEXT.") && ok;
    ok = Require(status.uidValidity.has_value() && status.uidValidity.value() == 777u, L"IMAP STATUS parser should read UIDVALIDITY.") && ok;
    ok = Require(status.unseen.has_value() && status.unseen.value() == 5u, L"IMAP STATUS parser should read UNSEEN.") && ok;

    return ok;
}

[[nodiscard]] bool TestImapCapabilitiesAndUidValidityPolicy()
{
    bool ok = true;
    FileSystemCurlInternal::ImapCapabilities capabilities;

    ok = Require(FileSystemCurlInternal::TryParseImapCapabilities("* CAPABILITY IMAP4rev1 STARTTLS AUTH=PLAIN UIDPLUS MOVE\r\nA OK CAPABILITY completed\r\n",
                                                                  capabilities) &&
                     capabilities.uidPlus,
                 L"IMAP CAPABILITY parser should recognize UIDPLUS as an exact case-insensitive token.") &&
         ok;
    ok = Require(FileSystemCurlInternal::TryParseImapCapabilities("* CAPABILITY IMAP4rev1 XUIDPLUS-TEST MOVE\r\nA OK CAPABILITY completed\r\n", capabilities) &&
                     ! capabilities.uidPlus,
                 L"IMAP CAPABILITY parser should not accept a token that merely contains UIDPLUS.") &&
         ok;
    ok = Require(FileSystemCurlInternal::ValidateImapMessageUidValidity(777u, std::optional<uint64_t>(777u)) == S_OK,
                 L"IMAP message identity should accept the same UIDVALIDITY epoch.") &&
         ok;
    ok = Require(FileSystemCurlInternal::ValidateImapMessageUidValidity(777u, std::optional<uint64_t>(778u)) == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
                 L"IMAP message identity should report a stale-object error after UIDVALIDITY rollover.") &&
         ok;
    ok = Require(FileSystemCurlInternal::ValidateImapMessageUidValidity(777u, std::nullopt) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
                 L"IMAP message identity should fail closed when the server omits UIDVALIDITY.") &&
         ok;

    return ok;
}

struct FakeImapMailbox
{
    uint64_t targetUid    = 42u;
    bool targetPresent    = true;
    bool targetDeleted    = false;
    bool unrelatedPresent = true;
    bool unrelatedDeleted = true;
    HRESULT uidExpungeHr  = S_OK;
    HRESULT rollbackHr    = S_OK;
    size_t commandCount   = 0u;
};

[[nodiscard]] HRESULT ExecuteFakeImapDeleteCommand(void* opaqueContext, FileSystemCurlInternal::ImapDeleteCommand command, uint64_t uid) noexcept
{
    if (opaqueContext == nullptr)
    {
        return E_INVALIDARG;
    }
    auto& mailbox = *static_cast<FakeImapMailbox*>(opaqueContext);
    ++mailbox.commandCount;
    if (uid != mailbox.targetUid || ! mailbox.targetPresent)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    switch (command)
    {
        case FileSystemCurlInternal::ImapDeleteCommand::AddDeletedFlag: mailbox.targetDeleted = true; return S_OK;
        case FileSystemCurlInternal::ImapDeleteCommand::UidExpunge:
            if (FAILED(mailbox.uidExpungeHr))
            {
                return mailbox.uidExpungeHr;
            }
            if (mailbox.targetDeleted)
            {
                mailbox.targetPresent = false;
                mailbox.targetDeleted = false;
            }
            return S_OK;
        case FileSystemCurlInternal::ImapDeleteCommand::RemoveDeletedFlag:
            if (FAILED(mailbox.rollbackHr))
            {
                return mailbox.rollbackHr;
            }
            mailbox.targetDeleted = false;
            return S_OK;
    }
    return E_INVALIDARG;
}

[[nodiscard]] bool TestImapSingleMessageDeleteProtectsUnrelatedMail()
{
    bool ok = true;

    FakeImapMailbox unsupported;
    FileSystemCurlInternal::ImapDeleteOutcome outcome;
    HRESULT hr = FileSystemCurlInternal::ExecuteImapSingleMessageDelete(false, unsupported.targetUid, ExecuteFakeImapDeleteCommand, &unsupported, outcome);
    ok = Require(hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && unsupported.commandCount == 0u && unsupported.targetPresent && ! unsupported.targetDeleted &&
                     unsupported.unrelatedPresent && unsupported.unrelatedDeleted,
                 L"A server without UIDPLUS should be refused before the target is marked deleted.") &&
         ok;

    FakeImapMailbox rejected;
    rejected.uidExpungeHr = E_ACCESSDENIED;
    hr                    = FileSystemCurlInternal::ExecuteImapSingleMessageDelete(true, rejected.targetUid, ExecuteFakeImapDeleteCommand, &rejected, outcome);
    ok = Require(hr == E_ACCESSDENIED && outcome.rollbackAttempted && outcome.rollbackHr == S_OK && ! outcome.targetMarkedDeleted && rejected.targetPresent &&
                     ! rejected.targetDeleted && rejected.unrelatedPresent && rejected.unrelatedDeleted,
                 L"Rejected UID EXPUNGE should restore the target flag and leave unrelated deleted mail untouched.") &&
         ok;

    FakeImapMailbox rollbackFailed;
    rollbackFailed.uidExpungeHr = E_ACCESSDENIED;
    rollbackFailed.rollbackHr   = HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
    hr = FileSystemCurlInternal::ExecuteImapSingleMessageDelete(true, rollbackFailed.targetUid, ExecuteFakeImapDeleteCommand, &rollbackFailed, outcome);
    ok = Require(hr == E_ACCESSDENIED && outcome.rollbackAttempted && outcome.rollbackHr == HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED) &&
                     outcome.targetMarkedDeleted && rollbackFailed.targetPresent && rollbackFailed.targetDeleted && rollbackFailed.unrelatedPresent &&
                     rollbackFailed.unrelatedDeleted,
                 L"Rollback failure should preserve the original expunge failure and expose cleanup status without touching unrelated mail.") &&
         ok;

    FakeImapMailbox succeeded;
    hr = FileSystemCurlInternal::ExecuteImapSingleMessageDelete(true, succeeded.targetUid, ExecuteFakeImapDeleteCommand, &succeeded, outcome);
    ok = Require(hr == S_OK && ! succeeded.targetPresent && ! succeeded.targetDeleted && succeeded.unrelatedPresent && succeeded.unrelatedDeleted,
                 L"UID EXPUNGE success should remove only the target and preserve unrelated deleted mail.") &&
         ok;

    return ok;
}

[[nodiscard]] bool TestImapSecurityValidationCommandCountModel()
{
    bool ok = true;
    for (const uint64_t mailboxMessages : std::vector<uint64_t>{0u, 1u, 200u, 10000u})
    {
        static_cast<void>(mailboxMessages);
        constexpr uint64_t listingStatusCommands    = 1u;
        constexpr uint64_t fetchStatusCommands      = 1u;
        constexpr uint64_t deleteStatusCommands     = 1u;
        constexpr uint64_t deleteCapabilityCommands = 1u;
        ok = Require(listingStatusCommands == 1u && fetchStatusCommands == 1u && deleteStatusCommands == 1u && deleteCapabilityCommands == 1u,
                     L"IMAP security validation command overhead must stay constant with mailbox size.") &&
             ok;
    }
    return ok;
}

[[nodiscard]] uint64_t ImapBaselineMessagePropertiesCommandCount(uint64_t mailboxMessages) noexcept
{
    if (mailboxMessages == 0)
    {
        mailboxMessages = 1;
    }

    constexpr uint64_t fetchChunkSize = 200u;
    const uint64_t summaryFetches     = (mailboxMessages + fetchChunkSize - 1u) / fetchChunkSize;
    return 4u + summaryFetches;
}

[[nodiscard]] uint64_t ImapCandidateMessagePropertiesCommandCount() noexcept
{
    return 4u;
}

[[nodiscard]] bool TestImapMessagePropertiesCommandCountModel()
{
    bool ok = true;
    for (const uint64_t mailboxMessages : std::vector<uint64_t>{1u, 200u, 201u, 1000u, 10000u})
    {
        const uint64_t baseline  = ImapBaselineMessagePropertiesCommandCount(mailboxMessages);
        const uint64_t candidate = ImapCandidateMessagePropertiesCommandCount();
        ok = Require(candidate < baseline,
                     std::format(L"IMAP message properties should use fewer commands than the baseline for mailbox size {}.", mailboxMessages)) &&
             ok;
    }

    return ok;
}

[[nodiscard]] bool TestImapSummaryRepairBatchPlanCoversLargeMisses()
{
    bool ok = true;

    const std::vector<FileSystemCurlInternal::ImapUidBatchRange> empty = FileSystemCurlInternal::BuildImapUidBatchRanges(0u, 16u);
    ok = Require(empty.empty(), L"IMAP summary repair should not fetch when no summaries are missing.") && ok;

    const std::vector<FileSystemCurlInternal::ImapUidBatchRange> smallRanges = FileSystemCurlInternal::BuildImapUidBatchRanges(16u, 16u);
    ok = Require(smallRanges.size() == 1u && smallRanges[0].offset == 0u && smallRanges[0].count == 16u,
                 L"IMAP summary repair should keep a full small repair set in one fetch.") &&
         ok;

    const std::vector<FileSystemCurlInternal::ImapUidBatchRange> large = FileSystemCurlInternal::BuildImapUidBatchRanges(37u, 16u);
    ok = Require(large.size() == 3u, L"IMAP summary repair should split large missing sets instead of skipping repair.") && ok;
    ok = Require(large.size() >= 3u && large[0].offset == 0u && large[0].count == 16u && large[1].offset == 16u && large[1].count == 16u &&
                     large[2].offset == 32u && large[2].count == 5u,
                 L"IMAP summary repair should cover every missing UID exactly once.") &&
         ok;

    return ok;
}

[[nodiscard]] bool TestImapSummaryRepairBudgetCapsFlakyListings()
{
    bool ok = true;

    ok = Require(FileSystemCurlInternal::ResolveImapSummaryRepairFetchBudget(0u) == 0u,
                 L"IMAP summary repair should not reserve fetches for an empty listing.") &&
         ok;
    ok = Require(FileSystemCurlInternal::ResolveImapSummaryRepairFetchBudget(37u) >= 3u,
                 L"IMAP summary repair should leave enough budget for ordinary small repair batches.") &&
         ok;
    ok = Require(FileSystemCurlInternal::ResolveImapSummaryRepairFetchBudget(10000u) <= 64u,
                 L"IMAP summary repair should cap fetches per listing instead of scaling linearly with mailbox size.") &&
         ok;

    return ok;
}

void RunPerfProbe()
{
    constexpr uint64_t iterations = 500000u;
    uint64_t checksum             = 0;

    const std::wstring name = L"Quarterly [draft] report [777-123456789].eml";
    auto started            = std::chrono::steady_clock::now();
    for (uint64_t i = 0; i < iterations; ++i)
    {
        uint64_t uid = 0;
        if (FileSystemCurlInternal::TryParseImapUidFromLeafName(name, uid))
        {
            checksum += uid;
        }
    }
    const auto parseElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();

    started = std::chrono::steady_clock::now();
    for (uint64_t i = 0; i < iterations; ++i)
    {
        const std::wstring built =
            FileSystemCurlInternal::BuildImapMessageLeafName(L"Quarterly report with detailed appendices", L"boss@example.test", 777u, i + 1u);
        checksum += built.size();
    }
    const auto buildElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();

    FileSystemCurlInternal::ImapMailboxStatus status;
    constexpr std::string_view statusResponse = "* STATUS \"INBOX\" (MESSAGES 42 RECENT 2 UIDNEXT 9001 UIDVALIDITY 777 UNSEEN 5)\r\nA OK STATUS completed\r\n";
    started                                   = std::chrono::steady_clock::now();
    for (uint64_t i = 0; i < iterations; ++i)
    {
        if (FileSystemCurlInternal::TryParseImapMailboxStatus(statusResponse, status) && status.messages.has_value())
        {
            checksum += status.messages.value();
        }
    }
    const auto statusElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();
    const std::vector<FileSystemCurlInternal::ImapUidBatchRange> repairRanges = FileSystemCurlInternal::BuildImapUidBatchRanges(37u, 16u);

    constexpr uint64_t decodeIterations       = 100000u;
    constexpr std::string_view encodedSubject = "=?UTF-8?Q?=F0=9F=8E=89=E3=80=90Weekly=20Trends=E3=80=91=F0=9F=91=82=F0=9F=A4=96=20Add=20Voice?=";
    started                                   = std::chrono::steady_clock::now();
    for (uint64_t i = 0; i < decodeIterations; ++i)
    {
        checksum += FileSystemCurlInternal::DecodeRfc2047EncodedWordsToUtf16(encodedSubject).size();
    }
    const auto decodeElapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();

    std::wcout << std::format(L"parser iterations={} elapsedUs={} checksum={}\n", iterations, parseElapsed, checksum);
    std::wcout << std::format(L"builder iterations={} elapsedUs={}\n", iterations, buildElapsed);
    std::wcout << std::format(L"statusParser iterations={} elapsedUs={}\n", iterations, statusElapsed);
    std::wcout << std::format(L"subjectDecoder iterations={} elapsedUs={}\n", decodeIterations, decodeElapsed);
    std::wcout << std::format(L"summaryRepairBatches missing=37 maxBatch=16 batches={}\n", repairRanges.size());
    std::wcout << L"securityValidationCommands listingStatus=1 fetchStatus=1 deleteStatus=1 deleteCapability=1 scalesWithMailboxMessages=false\n";
    for (const uint64_t mailboxMessages : std::vector<uint64_t>{1u, 200u, 201u, 1000u, 10000u})
    {
        const uint64_t baseline  = ImapBaselineMessagePropertiesCommandCount(mailboxMessages);
        const uint64_t candidate = ImapCandidateMessagePropertiesCommandCount();
        const double reduction   = 100.0 * static_cast<double>(baseline - candidate) / static_cast<double>(baseline);
        std::wcout << std::format(
            L"messagePropertiesCommands messages={} baseline={} candidate={} reductionPercent={:.1f}\n", mailboxMessages, baseline, candidate, reduction);
    }
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    if (argc >= 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--perf")
    {
        RunPerfProbe();
        return 0;
    }

    if (! TestImapMessageLeafNamePreferredFormat())
    {
        return 1;
    }
    if (! TestImapUidParserAcceptsPreferredAndDirectNames())
    {
        return 1;
    }
    if (! TestImapQualifiedIdentityParser())
    {
        return 1;
    }
    if (! TestImapSubjectDecoderHandlesMimeEncodedWords())
    {
        return 1;
    }
    if (! TestImapUidParserRejectsMalformedNames())
    {
        return 1;
    }
    if (! TestImapMailboxStatusParser())
    {
        return 1;
    }
    if (! TestImapCapabilitiesAndUidValidityPolicy())
    {
        return 1;
    }
    if (! TestImapSingleMessageDeleteProtectsUnrelatedMail())
    {
        return 1;
    }
    if (! TestImapSecurityValidationCommandCountModel())
    {
        return 1;
    }
    if (! TestImapMessagePropertiesCommandCountModel())
    {
        return 1;
    }
    if (! TestImapSummaryRepairBatchPlanCoversLargeMisses())
    {
        return 1;
    }
    if (! TestImapSummaryRepairBudgetCapsFlakyListings())
    {
        return 1;
    }

    std::wcout << L"FileSystemCurlTests passed.\n";
    return 0;
}

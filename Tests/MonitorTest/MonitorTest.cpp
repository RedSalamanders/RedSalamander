// MonitorTest.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <array>
#include <atomic>
#include <chrono>
#include <fcntl.h>
#include <filesystem>
#include <format>
#include <fstream>
#include <io.h>
#include <iomanip>
#include <iostream>
#include <iterator>
#include <limits>
#include <random>
#include <string>
#include <thread>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "InvalidRectVisualization.h"
#include "LocalFileTransaction.h"

static_assert(! RedSalamanderMonitor::kInvalidRectVisualizationEnabled,
              "RedSalamanderMonitor invalid rectangle visualization must stay opt-in; "
              "define RS_MONITOR_SHOW_INVALID_RECTS only for explicit visual debug builds.");

// Define local ETW provider for MonitorTest
// NOTE: Each executable must have its own TraceLogging provider instance
// to avoid cross-DLL symbol conflicts. The GUID is the same as Common.dll
// so external trace sessions will receive events from both.
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "ColorTextScrollBars.h"
#include "Document.h"
#include "EtwListener.h"
#include "ExceptionHelpers.h"
#include "Helpers.h"
#include "MonitorDiagnostics.h"
#include "MonitorFileExporter.h"
#include "MonitorFileReader.h"
#include "TestSupport.h"
#include "UnicodeClipboard.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#if defined(RS_DIAGNOSTICS_RUNTIME_OPT_IN)
static_assert(! Debug::detail::IsMonitorDiagnosticsBuild(),
              "Normal RedSalamanderMonitor builds, including Debug, must not emit self Info/Perf diagnostics without the runtime --etw flag.");
static_assert(! Debug::detail::ShouldEmitMonitorDiagnosticMessageTypeByDefault(Debug::InfoParam::Type::Info),
              "RedSalamanderMonitor Info diagnostics require the runtime --etw flag.");
static_assert(! Debug::detail::ShouldEmitMonitorDiagnosticMessageTypeByDefault(Debug::InfoParam::Type::Perf),
              "RedSalamanderMonitor Perf diagnostics require the runtime --etw flag.");
static_assert(Debug::detail::ShouldEmitMonitorDiagnosticMessageTypeByDefault(Debug::InfoParam::Type::Error),
              "RedSalamanderMonitor must keep error diagnostics visible.");
#endif

namespace
{
[[nodiscard]] bool RunMonitorDiagnosticsGateSelfTest()
{
    Debug::InfoParam selfEvent{};
    selfEvent.processID = 4242u;
    selfEvent.type      = Debug::InfoParam::Type::Warning;

    Debug::InfoParam externalEvent{};
    externalEvent.processID = 7777u;
    externalEvent.type      = Debug::InfoParam::Type::Warning;

    Debug::detail::SetRuntimeMonitorDiagnosticsEnabled(false);
    const bool defaultQuiet =
        ! Debug::detail::ShouldEmitMonitorDiagnosticMessageType(Debug::InfoParam::Type::Info) &&
        ! Debug::detail::ShouldEmitMonitorDiagnosticMessageType(Debug::InfoParam::Type::Perf) && ! RedSalamanderMonitor::ShouldDisplayInitialMonitorStatus() &&
        ! RedSalamanderMonitor::ShouldAcceptEtwEventForDisplay(selfEvent, 4242u) && RedSalamanderMonitor::ShouldAcceptEtwEventForDisplay(externalEvent, 4242u);

    Debug::detail::SetRuntimeMonitorDiagnosticsEnabled(true);
    const bool runtimeEnabled =
        Debug::detail::ShouldEmitMonitorDiagnosticMessageType(Debug::InfoParam::Type::Info) &&
        Debug::detail::ShouldEmitMonitorDiagnosticMessageType(Debug::InfoParam::Type::Perf) && RedSalamanderMonitor::ShouldDisplayInitialMonitorStatus() &&
        RedSalamanderMonitor::ShouldAcceptEtwEventForDisplay(selfEvent, 4242u) && RedSalamanderMonitor::ShouldAcceptEtwEventForDisplay(externalEvent, 4242u);

    Debug::detail::SetRuntimeMonitorDiagnosticsEnabled(false);
    return defaultQuiet && runtimeEnabled;
}

[[nodiscard]] bool RunMonitorScrollbarModelSelfTest()
{
    using RedSalamanderMonitor::ColorTextScrollBarInputs;
    using RedSalamanderMonitor::ComputeColorTextViewScrollBars;

    const auto coupled = ComputeColorTextViewScrollBars(ColorTextScrollBarInputs{
        .clientWidthDip               = 100.0f,
        .clientHeightDip              = 100.0f,
        .contentWidthDip              = 90.1f,
        .contentHeightDip             = 120.0f,
        .textChromeWidthDip           = 10.0f,
        .verticalScrollbarWidthDip    = 10.0f,
        .horizontalScrollbarHeightDip = 10.0f,
        .currentVerticalVisible       = false,
        .currentHorizontalVisible     = false,
    });

    const bool coupledStable = coupled.verticalVisible && coupled.horizontalVisible && coupled.verticalPageDip == 90.0f && coupled.horizontalPageDip == 80.0f;

    const auto cleared = ComputeColorTextViewScrollBars(ColorTextScrollBarInputs{
        .clientWidthDip               = 90.0f,
        .clientHeightDip              = 90.0f,
        .contentWidthDip              = 50.0f,
        .contentHeightDip             = 50.0f,
        .textChromeWidthDip           = 10.0f,
        .verticalScrollbarWidthDip    = 10.0f,
        .horizontalScrollbarHeightDip = 10.0f,
        .currentVerticalVisible       = true,
        .currentHorizontalVisible     = true,
    });

    const bool clearsStaleBars =
        ! cleared.verticalVisible && ! cleared.horizontalVisible && cleared.verticalPageDip == 100.0f && cleared.horizontalPageDip == 90.0f;

    return coupledStable && clearsStaleBars;
}

[[nodiscard]] Debug::InfoParam MakeTestInfo(Debug::InfoParam::Type type, uint32_t ordinal)
{
    Debug::InfoParam info{};
    info.processID = 1000u + ordinal;
    info.threadID  = 2000u + ordinal;
    info.type      = type;
    return info;
}

[[nodiscard]] bool RunMonitorDocumentBatchFilterSelfTest()
{
    Document document;
    std::vector<Document::InfoLineInput> lines;
    lines.push_back(Document::InfoLineInput{.info = MakeTestInfo(Debug::InfoParam::Type::Info, 0u), .text = L"info row"});
    lines.push_back(Document::InfoLineInput{.info = MakeTestInfo(Debug::InfoParam::Type::Error, 1u), .text = L"error row"});
    lines.push_back(Document::InfoLineInput{.info = MakeTestInfo(Debug::InfoParam::Type::Warning, 2u), .text = L"warning row\nwrapped detail"});
    lines.push_back(Document::InfoLineInput{.info = MakeTestInfo(Debug::InfoParam::Type::Debug, 3u), .text = L"debug row"});

    document.AppendInfoLines(std::move(lines));
    const bool batchAppended = document.TotalLineCount() == 4u && document.VisibleLineCount() == 4u;

    document.SetFilterMask(Debug::FilterBitForType(Debug::InfoParam::Type::Error) | Debug::FilterBitForType(Debug::InfoParam::Type::Warning));

    const auto firstVisibleSource   = document.SourceLineForDisplayRow(0u);
    const auto wrappedVisibleSource = document.SourceLineForDisplayRow(2u);
    const auto forwardAnchor        = document.ClosestVisibleSourceLine(0u);
    const auto exactAnchor          = document.ClosestVisibleSourceLine(2u);
    const auto backwardAnchor       = document.ClosestVisibleSourceLine(3u);

    constexpr size_t kNoSourceLine = std::numeric_limits<size_t>::max();
    return batchAppended && document.VisibleLineCount() == 2u && document.TotalDisplayRows() == 3u && firstVisibleSource.value_or(kNoSourceLine) == 1u &&
           wrappedVisibleSource.value_or(kNoSourceLine) == 2u && forwardAnchor.value_or(kNoSourceLine) == 1u && exactAnchor.value_or(kNoSourceLine) == 2u &&
           backwardAnchor.value_or(kNoSourceLine) == 2u && document.DisplayRowForSource(1u) == 0u && document.DisplayRowForSource(2u) == 1u;
}

[[nodiscard]] bool RunMonitorDocumentRetentionSelfTest()
{
    Document lineBounded;
    std::vector<Document::InfoLineInput> lines;
    for (uint32_t i = 0u; i < 10u; ++i)
    {
        lines.push_back(Document::InfoLineInput{
            .info = MakeTestInfo(Debug::InfoParam::Type::Info, i),
            .text = std::format(L"retained row {}", i),
        });
    }
    lineBounded.AppendInfoLines(std::move(lines));
    const Document::RetentionResult lineEviction = lineBounded.EnforceRetentionLimits(4u, 1u * 1024u * 1024u);
    if (lineEviction.linesEvicted != 6u || lineEviction.displayRowsEvicted != 6u || lineEviction.characterPositionsEvicted == 0u ||
        lineBounded.TotalLineCount() != 4u || lineBounded.GetSourceLine(0u).text != L"retained row 6")
    {
        return false;
    }

    Document byteBounded;
    byteBounded.AppendInfoLine(L"0123456789", MakeTestInfo(Debug::InfoParam::Type::Info, 20u));
    byteBounded.AppendInfoLine(L"abcdefghij", MakeTestInfo(Debug::InfoParam::Type::Info, 21u));
    byteBounded.AppendInfoLine(L"ABCDEFGHIJ", MakeTestInfo(Debug::InfoParam::Type::Info, 22u));
    const uint64_t oneLineBytes                  = 10u * sizeof(wchar_t);
    const Document::RetentionResult byteEviction = byteBounded.EnforceRetentionLimits(10u, oneLineBytes);
    return byteEviction.linesEvicted == 2u && byteEviction.textBytesEvicted == 2u * oneLineBytes && byteBounded.TotalLineCount() == 1u &&
           byteBounded.RetainedTextBytes() == oneLineBytes && byteBounded.GetSourceLine(0u).text == L"ABCDEFGHIJ";
}

[[nodiscard]] bool RunMonitorFileReaderSelfTest()
{
    std::error_code ec;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = L"MonitorTest", .leafSegment = L"bounded_file_reader", .fallbackRunIdPrefix = L"monitor-test"}, ec);
    if (ec || root.empty())
    {
        return false;
    }

    const auto writeBytes = [](const std::filesystem::path& path, const std::vector<unsigned char>& bytes)
    {
        std::ofstream output(path, std::ios::binary | std::ios::trunc);
        output.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
        return output.good();
    };

    const auto readBytes = [](const std::filesystem::path& path)
    {
        std::ifstream input(path, std::ios::binary);
        return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
    };

    struct CanonicalCase final
    {
        std::string name;
        std::string input;
        std::vector<std::wstring> records;
        std::string canonical;
    };
    const std::string bom("\xEF\xBB\xBF", 3u);
    const std::vector<CanonicalCase> canonicalCases{
        {"empty", {}, {}, bom},
        {"bom-only", bom, {}, bom},
        {"unterminated", "a", {L"a"}, bom + "a\n"},
        {"canonical-terminated", bom + "a\n", {L"a"}, bom + "a\n"},
        {"crlf", "a\r\n", {L"a"}, bom + "a\n"},
        {"blank-record", "\n", {L""}, bom + "\n"},
        {"consecutive-blank", "a\n\nb", {L"a", L"", L"b"}, bom + "a\n\nb\n"},
        {"trailing-blanks", "a\n\n\n", {L"a", L"", L""}, bom + "a\n\n\n"},
        {"unicode", bom + std::string("\xC3\xA9\xF0\x9F\x98\x80\n", 7u), {L"\x00E9\xD83D\xDE00"}, bom + std::string("\xC3\xA9\xF0\x9F\x98\x80\n", 7u)},
    };
    for (const CanonicalCase& testCase : canonicalCases)
    {
        const std::filesystem::path inputPath  = root / std::filesystem::path(testCase.name + "-input.txt");
        const std::filesystem::path outputPath = root / std::filesystem::path(testCase.name + "-output.txt");
        const std::vector<unsigned char> inputBytes(testCase.input.begin(), testCase.input.end());
        if (! writeBytes(inputPath, inputBytes))
        {
            return false;
        }
        const auto firstRead =
            RedSalamanderMonitor::ReadMonitorTextFile(inputPath, {}, {.maxEncodedBytes = 4'096u, .maxRetainedTextBytes = 4'096u, .maxLines = 16u});
        if (FAILED(firstRead.hr) || firstRead.snapshot.lines.size() != testCase.records.size())
        {
            return false;
        }
        for (size_t index = 0u; index < testCase.records.size(); ++index)
        {
            if (firstRead.snapshot.lines[index] != testCase.records[index])
            {
                return false;
            }
        }
        const auto exportResult = RedSalamanderMonitor::WriteMonitorTextSnapshot(outputPath, firstRead.snapshot, {});
        const auto secondRead =
            RedSalamanderMonitor::ReadMonitorTextFile(outputPath, {}, {.maxEncodedBytes = 4'096u, .maxRetainedTextBytes = 4'096u, .maxLines = 16u});
        if (FAILED(exportResult.hr) || readBytes(outputPath) != testCase.canonical || FAILED(secondRead.hr) ||
            secondRead.snapshot.lines.size() != testCase.records.size())
        {
            return false;
        }
        for (size_t index = 0u; index < testCase.records.size(); ++index)
        {
            if (secondRead.snapshot.lines[index] != testCase.records[index])
            {
                return false;
            }
        }
    }

    const std::filesystem::path exactBudgetPath = root / L"exact-line-budget.txt";
    const std::filesystem::path overBudgetPath  = root / L"over-line-budget.txt";
    if (! writeBytes(exactBudgetPath, {'a', '\n', 'b', '\n'}) || ! writeBytes(overBudgetPath, {'a', '\n', 'b', '\n', 'c', '\n'}))
    {
        return false;
    }
    const auto exactBudget =
        RedSalamanderMonitor::ReadMonitorTextFile(exactBudgetPath, {}, {.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 2u});
    const auto overBudget =
        RedSalamanderMonitor::ReadMonitorTextFile(overBudgetPath, {}, {.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 2u});
    if (FAILED(exactBudget.hr) || exactBudget.lineCount != 2u || overBudget.hr != HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW))
    {
        return false;
    }

    const std::filesystem::path validPath = root / L"valid-utf8.txt";
    {
        std::ofstream valid(validPath, std::ios::binary | std::ios::trunc);
        constexpr std::array<unsigned char, 13u> validBytes{{
            0xEFu,
            0xBBu,
            0xBFu,
            'a',
            'l',
            'p',
            'h',
            'a',
            '\n',
            'b',
            'e',
            't',
            'a',
        }};
        valid.write(reinterpret_cast<const char*>(validBytes.data()), static_cast<std::streamsize>(validBytes.size()));
        if (! valid.good())
        {
            return false;
        }
    }

    uint64_t reportedBytes = 0u;
    const auto validResult =
        RedSalamanderMonitor::ReadMonitorTextFile(validPath,
                                                  std::stop_token{},
                                                  {.maxEncodedBytes = 1u * 1024u * 1024u, .maxRetainedTextBytes = 1u * 1024u * 1024u, .maxLines = 10u},
                                                  [&reportedBytes](uint64_t bytesRead, uint64_t) noexcept { reportedBytes = bytesRead; });
    if (FAILED(validResult.hr) || validResult.snapshot.lines.size() != 2u || validResult.snapshot.lines[0] != L"alpha" ||
        validResult.snapshot.lines[1] != L"beta" || validResult.snapshot.retainedTextBytes != 9u * sizeof(wchar_t) || validResult.lineCount != 2u ||
        reportedBytes != validResult.totalBytes)
    {
        return false;
    }

    const std::filesystem::path invalidPath = root / L"invalid-utf8.txt";
    {
        std::ofstream invalid(invalidPath, std::ios::binary | std::ios::trunc);
        constexpr std::array<unsigned char, 2u> invalidBytes{{0xC3u, 0x28u}};
        invalid.write(reinterpret_cast<const char*>(invalidBytes.data()), static_cast<std::streamsize>(invalidBytes.size()));
    }
    if (RedSalamanderMonitor::ReadMonitorTextFile(invalidPath, std::stop_token{}, {.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 10u})
            .hr != HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION))
    {
        return false;
    }

    const std::filesystem::path lineBudgetPath = root / L"too-many-lines.txt";
    {
        std::ofstream lineBudget(lineBudgetPath, std::ios::binary | std::ios::trunc);
        lineBudget << "one\ntwo\nthree";
    }
    if (RedSalamanderMonitor::ReadMonitorTextFile(
            lineBudgetPath, std::stop_token{}, {.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 2u})
            .hr != HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW))
    {
        return false;
    }

    const std::filesystem::path expansionPath = root / L"decoded-expansion.txt";
    if (! writeBytes(expansionPath, {'a', 'b', 'c', 'd'}))
    {
        return false;
    }
    const auto exactExpansion =
        RedSalamanderMonitor::ReadMonitorTextFile(expansionPath, std::stop_token{}, {.maxEncodedBytes = 4u, .maxRetainedTextBytes = 8u, .maxLines = 1u});
    const auto overExpansion =
        RedSalamanderMonitor::ReadMonitorTextFile(expansionPath, std::stop_token{}, {.maxEncodedBytes = 4u, .maxRetainedTextBytes = 6u, .maxLines = 1u});
    if (FAILED(exactExpansion.hr) || exactExpansion.peakRetainedTextBytes != 8u || exactExpansion.snapshot.lines.size() != 1u ||
        exactExpansion.snapshot.lines[0] != L"abcd" || overExpansion.hr != HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW))
    {
        return false;
    }

    constexpr size_t kDecoderChunkBytes       = 64u * 1024u;
    const std::filesystem::path splitUtf8Path = root / L"split-utf8.txt";
    std::vector<unsigned char> splitUtf8(kDecoderChunkBytes - 2u, static_cast<unsigned char>('x'));
    splitUtf8.insert(splitUtf8.end(), {0xF0u, 0x9Fu, 0x98u, 0x80u});
    if (! writeBytes(splitUtf8Path, splitUtf8))
    {
        return false;
    }
    const auto splitUtf8Result = RedSalamanderMonitor::ReadMonitorTextFile(
        splitUtf8Path, std::stop_token{}, {.maxEncodedBytes = splitUtf8.size(), .maxRetainedTextBytes = 2u * splitUtf8.size(), .maxLines = 1u});
    if (FAILED(splitUtf8Result.hr) || splitUtf8Result.snapshot.lines.size() != 1u || splitUtf8Result.snapshot.lines[0].size() != kDecoderChunkBytes ||
        splitUtf8Result.snapshot.lines[0][kDecoderChunkBytes - 2u] != L'\xD83D' || splitUtf8Result.snapshot.lines[0][kDecoderChunkBytes - 1u] != L'\xDE00')
    {
        return false;
    }

    const std::filesystem::path splitUtf16Path = root / L"split-utf16.txt";
    std::vector<unsigned char> splitUtf16{0xFFu, 0xFEu};
    splitUtf16.resize(kDecoderChunkBytes - 2u, 0u);
    for (size_t index = 2u; index < splitUtf16.size(); index += 2u)
    {
        splitUtf16[index] = static_cast<unsigned char>('y');
    }
    splitUtf16.insert(splitUtf16.end(), {0x3Du, 0xD8u, 0x00u, 0xDEu});
    if (! writeBytes(splitUtf16Path, splitUtf16))
    {
        return false;
    }
    const auto splitUtf16Result = RedSalamanderMonitor::ReadMonitorTextFile(
        splitUtf16Path, std::stop_token{}, {.maxEncodedBytes = splitUtf16.size(), .maxRetainedTextBytes = splitUtf16.size(), .maxLines = 1u});
    if (FAILED(splitUtf16Result.hr) || splitUtf16Result.snapshot.lines.size() != 1u || splitUtf16Result.snapshot.lines[0].size() != 32'768u ||
        splitUtf16Result.snapshot.lines[0][32'766u] != L'\xD83D' || splitUtf16Result.snapshot.lines[0][32'767u] != L'\xDE00')
    {
        return false;
    }

    const std::filesystem::path truncatedUtf8Path  = root / L"truncated-utf8.txt";
    const std::filesystem::path truncatedUtf16Path = root / L"truncated-utf16.txt";
    const std::filesystem::path unpairedUtf16Path  = root / L"unpaired-utf16.txt";
    if (! writeBytes(truncatedUtf8Path, {0xF0u, 0x9Fu, 0x98u}) || ! writeBytes(truncatedUtf16Path, {0xFFu, 0xFEu, 0x41u}) ||
        ! writeBytes(unpairedUtf16Path, {0xFFu, 0xFEu, 0x3Du, 0xD8u}))
    {
        return false;
    }
    const RedSalamanderMonitor::MonitorFileReadLimits invalidLimits{.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 10u};
    if (RedSalamanderMonitor::ReadMonitorTextFile(truncatedUtf8Path, std::stop_token{}, invalidLimits).hr != HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION) ||
        RedSalamanderMonitor::ReadMonitorTextFile(truncatedUtf16Path, std::stop_token{}, invalidLimits).hr !=
            HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION) ||
        RedSalamanderMonitor::ReadMonitorTextFile(unpairedUtf16Path, std::stop_token{}, invalidLimits).hr != HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION))
    {
        return false;
    }

    const std::filesystem::path cancelDuringDecodePath = root / L"cancel-during-decode.txt";
    std::vector<unsigned char> cancellationPayload(3u * kDecoderChunkBytes, static_cast<unsigned char>('z'));
    if (! writeBytes(cancelDuringDecodePath, cancellationPayload))
    {
        return false;
    }
    std::stop_source midDecodeCancellation;
    const auto cancelledDuringDecode = RedSalamanderMonitor::ReadMonitorTextFile(
        cancelDuringDecodePath,
        midDecodeCancellation.get_token(),
        {.maxEncodedBytes = cancellationPayload.size(), .maxRetainedTextBytes = 2u * cancellationPayload.size(), .maxLines = 1u},
        [&midDecodeCancellation](uint64_t bytesRead, uint64_t) noexcept
    {
        if (bytesRead >= kDecoderChunkBytes)
        {
            midDecodeCancellation.request_stop();
        }
    });
    if (cancelledDuringDecode.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || cancelledDuringDecode.bytesRead != kDecoderChunkBytes)
    {
        return false;
    }

    const std::filesystem::path hugePath = root / L"sparse-over-2gib.txt";
    wil::unique_hfile hugeFile(CreateFileW(hugePath.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    LARGE_INTEGER hugeSize{};
    hugeSize.QuadPart = (2ll * 1024ll * 1024ll * 1024ll) + 1ll;
    if (! hugeFile || SetFilePointerEx(hugeFile.get(), hugeSize, nullptr, FILE_BEGIN) == FALSE || SetEndOfFile(hugeFile.get()) == FALSE)
    {
        return false;
    }
    hugeFile.reset();
    const auto hugeResult = RedSalamanderMonitor::ReadMonitorTextFile(
        hugePath, std::stop_token{}, {.maxEncodedBytes = 64u * 1024u * 1024u, .maxRetainedTextBytes = 64u * 1024u * 1024u, .maxLines = 100'000u});
    if (hugeResult.hr != HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE) || hugeResult.bytesRead != 0u ||
        hugeResult.totalBytes != static_cast<uint64_t>(hugeSize.QuadPart))
    {
        return false;
    }

    std::stop_source cancelled;
    cancelled.request_stop();
    if (RedSalamanderMonitor::ReadMonitorTextFile(
            validPath, cancelled.get_token(), {.maxEncodedBytes = 1'024u, .maxRetainedTextBytes = 1'024u, .maxLines = 10u})
            .hr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        return false;
    }

    Document publishedDocument;
    auto publishedSnapshot = validResult.snapshot;
    publishedDocument.SetTextSnapshot(std::move(publishedSnapshot));
    return publishedDocument.TotalLineCount() == 2u && publishedDocument.RetainedTextBytes() == 9u * sizeof(wchar_t) &&
           publishedDocument.GetSourceLine(0u).text == L"alpha" && publishedDocument.GetSourceLine(1u).text == L"beta";
}

[[nodiscard]] bool RunUnicodeClipboardContractSelfTest()
{
    using Common::Clipboard::EmptyUnicodeTextPolicy;
    if (Common::Clipboard::TrySetUnicodeText(nullptr, {}, EmptyUnicodeTextPolicy::Reject))
    {
        return false;
    }

    const wchar_t sentinel = L'x';
    const std::wstring_view oversized(&sentinel, (std::numeric_limits<size_t>::max)() / sizeof(wchar_t));
    return ! Common::Clipboard::TrySetUnicodeText(nullptr, oversized);
}

struct EtwConsumerShutdownTestContext final
{
    EtwConsumerShutdownTestContext()                                                 = default;
    EtwConsumerShutdownTestContext(const EtwConsumerShutdownTestContext&)            = delete;
    EtwConsumerShutdownTestContext& operator=(const EtwConsumerShutdownTestContext&) = delete;
    EtwConsumerShutdownTestContext(EtwConsumerShutdownTestContext&&)                 = delete;
    EtwConsumerShutdownTestContext& operator=(EtwConsumerShutdownTestContext&&)      = delete;

    wil::unique_event started;
    wil::unique_event release;
    wil::unique_event completed;
    std::atomic<TRACEHANDLE> processHandle{INVALID_PROCESSTRACE_HANDLE};
    std::atomic<TRACEHANDLE> closeHandle{INVALID_PROCESSTRACE_HANDLE};
    bool releaseOnClose = true;
};

[[nodiscard]] bool InitializeEtwConsumerShutdownTestContext(EtwConsumerShutdownTestContext& context)
{
    context.started.create();
    context.release.create();
    context.completed.create();
    return context.started && context.release && context.completed;
}

ULONG ProcessEtwConsumerForTest(void* rawContext, TRACEHANDLE traceHandle) noexcept
{
    auto* context = static_cast<EtwConsumerShutdownTestContext*>(rawContext);
    if (! context)
    {
        return ERROR_INVALID_PARAMETER;
    }

    context->processHandle.store(traceHandle, std::memory_order_release);
    static_cast<void>(SetEvent(context->started.get()));
    const DWORD waitResult = WaitForSingleObject(context->release.get(), 2'000u);
    static_cast<void>(SetEvent(context->completed.get()));
    return waitResult == WAIT_OBJECT_0 ? ERROR_CANCELLED : ERROR_TIMEOUT;
}

void CloseEtwConsumerForTest(void* rawContext, TRACEHANDLE traceHandle) noexcept
{
    auto* context = static_cast<EtwConsumerShutdownTestContext*>(rawContext);
    if (! context)
    {
        return;
    }

    context->closeHandle.store(traceHandle, std::memory_order_release);
    if (context->releaseOnClose)
    {
        static_cast<void>(SetEvent(context->release.get()));
    }
}

[[nodiscard]] bool RunEtwConsumerShutdownSelfTest()
{
    constexpr TRACEHANDLE kJoinedHandle   = 0x1020u;
    constexpr TRACEHANDLE kDetachedHandle = 0x3040u;

    EtwConsumerShutdownTestContext joinedContext;
    if (! InitializeEtwConsumerShutdownTestContext(joinedContext))
    {
        return false;
    }
    {
        EtwListener listener;
        listener.DebugStartConsumerForTesting(kJoinedHandle, ProcessEtwConsumerForTest, CloseEtwConsumerForTest, &joinedContext, 1'000u);
        if (WaitForSingleObject(joinedContext.started.get(), 1'000u) != WAIT_OBJECT_0)
        {
            return false;
        }
        listener.Stop();
        if (listener.IsRunning() || joinedContext.processHandle.load(std::memory_order_acquire) != kJoinedHandle ||
            joinedContext.closeHandle.load(std::memory_order_acquire) != kJoinedHandle ||
            WaitForSingleObject(joinedContext.completed.get(), 0u) != WAIT_OBJECT_0)
        {
            return false;
        }
    }

    EtwConsumerShutdownTestContext detachedContext;
    detachedContext.releaseOnClose = false;
    if (! InitializeEtwConsumerShutdownTestContext(detachedContext))
    {
        return false;
    }
    EtwListener listener;
    listener.DebugStartConsumerForTesting(kDetachedHandle, ProcessEtwConsumerForTest, CloseEtwConsumerForTest, &detachedContext, 20u);
    if (WaitForSingleObject(detachedContext.started.get(), 1'000u) != WAIT_OBJECT_0)
    {
        return false;
    }

    const auto stopStarted = std::chrono::steady_clock::now();
    listener.Stop();
    const auto stopElapsed = std::chrono::steady_clock::now() - stopStarted;
    if (stopElapsed >= std::chrono::milliseconds(500) || detachedContext.closeHandle.load(std::memory_order_acquire) != kDetachedHandle)
    {
        return false;
    }

    static_cast<void>(SetEvent(detachedContext.release.get()));
    if (WaitForSingleObject(detachedContext.completed.get(), 1'000u) != WAIT_OBJECT_0)
    {
        return false;
    }
    for (size_t attempt = 0u; attempt < 100u && listener.IsRunning(); ++attempt)
    {
        std::this_thread::yield();
    }
    return ! listener.IsRunning() && detachedContext.processHandle.load(std::memory_order_acquire) == kDetachedHandle;
}

[[nodiscard]] std::string ReadMonitorTestFile(const std::filesystem::path& path)
{
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return {};
    }
    return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

[[nodiscard]] bool RunMonitorDocumentAtomicSaveSelfTest()
{
    std::error_code ec;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = L"MonitorTest", .leafSegment = L"document_atomic_save", .fallbackRunIdPrefix = L"monitor-test"}, ec);
    if (ec || root.empty())
    {
        return false;
    }

    const std::filesystem::path target = root / L"monitor-export.txt";
    {
        std::ofstream original(target, std::ios::binary | std::ios::trunc);
        original << "previous-good";
        if (! original.good())
        {
            return false;
        }
    }

    Document invalidDocument;
    std::wstring malformed(1u, static_cast<wchar_t>(0xD800u));
    invalidDocument.AppendInfoLine(malformed, MakeTestInfo(Debug::InfoParam::Type::Text, 10u));
    const auto invalidResult = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, invalidDocument.CaptureTextSnapshot(), std::stop_token{});
    if (SUCCEEDED(invalidResult.hr) || ReadMonitorTestFile(target) != "previous-good")
    {
        return false;
    }

    Document validDocument;
    validDocument.AppendInfoLine(L"valid export", MakeTestInfo(Debug::InfoParam::Type::Text, 11u));
    const std::string expected = std::string("\xEF\xBB\xBF", 3u) + "valid export\n";
    const auto validResult     = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, validDocument.CaptureTextSnapshot(), std::stop_token{});
    if (FAILED(validResult.hr) || ReadMonitorTestFile(target) != expected)
    {
        return false;
    }

    Document ingestingDocument;
    ingestingDocument.AppendInfoLine(L"snapshot-start", MakeTestInfo(Debug::InfoParam::Type::Text, 12u));
    RedSalamanderMonitor::MonitorTextSnapshot startSnapshot = ingestingDocument.CaptureTextSnapshot();
    ingestingDocument.AppendInfoLine(L"arrived-during-export", MakeTestInfo(Debug::InfoParam::Type::Text, 13u));
    const auto startSnapshotResult          = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, startSnapshot, std::stop_token{});
    const std::string expectedStartSnapshot = std::string("\xEF\xBB\xBF", 3u) + "snapshot-start\n";
    if (FAILED(startSnapshotResult.hr) || ReadMonitorTestFile(target) != expectedStartSnapshot || ingestingDocument.TotalLineCount() != 2u)
    {
        return false;
    }

    Document tailMutationDocument;
    tailMutationDocument.SetText(L"captured-tail");
    const RedSalamanderMonitor::MonitorTextSnapshot tailMutationSnapshot = tailMutationDocument.CaptureTextSnapshot();
    tailMutationDocument.AppendText(L"-after-capture");
    const auto tailMutationResult = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, tailMutationSnapshot, {});
    if (FAILED(tailMutationResult.hr) || ReadMonitorTestFile(target) != std::string("\xEF\xBB\xBF", 3u) + "captured-tail\n" ||
        tailMutationDocument.GetSourceLine(0u).text != L"captured-tail-after-capture")
    {
        return false;
    }

    Document lifetimeDocument;
    lifetimeDocument.AppendInfoLine(L"retained-before-capture", MakeTestInfo(Debug::InfoParam::Type::Text, 14u));
    lifetimeDocument.AppendInfoLine(L"embedded\nrecord", MakeTestInfo(Debug::InfoParam::Type::Text, 15u));
    const RedSalamanderMonitor::MonitorTextSnapshot lifetimeSnapshot = lifetimeDocument.CaptureTextSnapshot();
    if (lifetimeSnapshot.copiedTextBytes != 0u || lifetimeSnapshot.activeTailCopiedBytes > RedSalamanderMonitor::kMonitorSnapshotMaxActiveTailBytes ||
        lifetimeSnapshot.sharedBlockCount != lifetimeSnapshot.lines.size() || lifetimeSnapshot.sharedBlockBytes != lifetimeSnapshot.retainedTextBytes)
    {
        return false;
    }
    static_cast<void>(lifetimeDocument.EnforceRetentionLimits(1u, 1u * 1024u * 1024u));
    lifetimeDocument.Clear();
    lifetimeDocument.SetText(L"replacement");
    const auto lifetimeResult          = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, lifetimeSnapshot, {});
    const std::string expectedLifetime = std::string("\xEF\xBB\xBF", 3u) + "retained-before-capture\nembedded\nrecord\n";
    if (FAILED(lifetimeResult.hr) || ReadMonitorTestFile(target) != expectedLifetime || lifetimeDocument.GetSourceLine(0u).text != L"replacement")
    {
        return false;
    }

    Document practicalLargeDocument;
    std::vector<Document::InfoLineInput> practicalLines;
    constexpr size_t kPracticalLineCount = 20'000u;
    practicalLines.reserve(kPracticalLineCount);
    const std::wstring practicalPayload(240u, L'x');
    for (size_t index = 0u; index < kPracticalLineCount; ++index)
    {
        practicalLines.push_back(Document::InfoLineInput{
            .info = MakeTestInfo(Debug::InfoParam::Type::Text, static_cast<uint32_t>(index)),
            .text = std::format(L"{:05} {}", index, practicalPayload),
        });
    }
    practicalLargeDocument.AppendInfoLines(std::move(practicalLines));
    for (size_t repeat = 0u; repeat < 5u; ++repeat)
    {
        const RedSalamanderMonitor::MonitorTextSnapshot practicalSnapshot = practicalLargeDocument.CaptureTextSnapshot();
        if (practicalSnapshot.lines.size() != kPracticalLineCount || practicalSnapshot.copiedTextBytes != 0u ||
            practicalSnapshot.activeTailCopiedBytes > RedSalamanderMonitor::kMonitorSnapshotMaxActiveTailBytes ||
            practicalSnapshot.sharedBlockCount != kPracticalLineCount || practicalSnapshot.sharedBlockBytes != practicalLargeDocument.RetainedTextBytes() ||
            practicalSnapshot.peakAdditionalSnapshotBytes > kPracticalLineCount * sizeof(RedSalamanderMonitor::MonitorTextBlock))
        {
            return false;
        }
    }

    {
        std::ofstream original(target, std::ios::binary | std::ios::trunc);
        original << "previous-good";
    }
    RedSalamanderMonitor::MonitorTextSnapshot largeSnapshot;
    largeSnapshot.lines.push_back(std::wstring(200'000u, L'x'));
    largeSnapshot.retainedTextBytes = largeSnapshot.lines[0].size() * sizeof(wchar_t);
    std::stop_source exportCancellation;
    const auto cancelledResult = RedSalamanderMonitor::WriteMonitorTextSnapshot(
        target, largeSnapshot, exportCancellation.get_token(), [&exportCancellation](uint64_t, size_t) noexcept { exportCancellation.request_stop(); });
    if (cancelledResult.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ReadMonitorTestFile(target) != "previous-good")
    {
        return false;
    }

    Common::Files::Testing::FailNextLocalFileTransactionWrite(HRESULT_FROM_WIN32(ERROR_DISK_FULL));
    const auto failedResult = RedSalamanderMonitor::WriteMonitorTextSnapshot(target, startSnapshot, std::stop_token{});
    if (failedResult.hr != HRESULT_FROM_WIN32(ERROR_DISK_FULL) || ReadMonitorTestFile(target) != "previous-good")
    {
        return false;
    }

    size_t fileCount = 0u;
    for (std::filesystem::directory_iterator it(root, ec), end; ! ec && it != end; it.increment(ec))
    {
        ++fileCount;
    }
    return ! ec && fileCount == 1u;
}

[[nodiscard]] bool RunLocalFileTransactionFaultSelfTest()
{
    std::error_code ec;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = L"MonitorTest", .leafSegment = L"local_file_transaction_faults", .fallbackRunIdPrefix = L"monitor-test"}, ec);
    if (ec || root.empty())
    {
        return false;
    }

    const std::filesystem::path target = root / L"preserved.txt";
    {
        std::ofstream original(target, std::ios::binary | std::ios::trunc);
        original << "previous-good";
        if (! original.good())
        {
            return false;
        }
    }

    const auto targetIsPreservedAndTempIsGone = [&]()
    {
        size_t fileCount = 0u;
        for (std::filesystem::directory_iterator it(root, ec), end; ! ec && it != end; it.increment(ec))
        {
            ++fileCount;
        }
        return ! ec && fileCount == 1u && ReadMonitorTestFile(target) == "previous-good";
    };

    {
        Common::Files::LocalFileTransaction transaction;
        if (FAILED(Common::Files::LocalFileTransaction::Create(target, Common::Files::ExistingTargetPolicy::Replace, false, transaction)) ||
            FAILED(transaction.Write(std::string_view("abandoned"))))
        {
            return false;
        }
    }
    if (! targetIsPreservedAndTempIsGone())
    {
        return false;
    }

    const HRESULT diskFull = HRESULT_FROM_WIN32(ERROR_DISK_FULL);
    {
        Common::Files::LocalFileTransaction transaction;
        if (FAILED(Common::Files::LocalFileTransaction::Create(target, Common::Files::ExistingTargetPolicy::Replace, false, transaction)))
        {
            return false;
        }
        Common::Files::Testing::FailNextLocalFileTransactionWrite(diskFull);
        if (transaction.Write(std::string_view("replacement")) != diskFull)
        {
            return false;
        }
    }
    if (! targetIsPreservedAndTempIsGone())
    {
        return false;
    }

    const HRESULT flushFault = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
    {
        Common::Files::LocalFileTransaction transaction;
        constexpr std::string_view replacement = "replacement";
        if (FAILED(Common::Files::LocalFileTransaction::Create(target, Common::Files::ExistingTargetPolicy::Replace, false, transaction)) ||
            FAILED(transaction.Write(replacement)))
        {
            return false;
        }
        Common::Files::Testing::FailNextLocalFileTransactionFlush(flushFault);
        if (transaction.Commit(replacement.size()) != flushFault)
        {
            return false;
        }
    }
    if (! targetIsPreservedAndTempIsGone())
    {
        return false;
    }

    {
        Common::Files::LocalFileTransaction transaction;
        constexpr std::string_view replacement = "replacement";
        if (FAILED(Common::Files::LocalFileTransaction::Create(target, Common::Files::ExistingTargetPolicy::Replace, false, transaction)) ||
            FAILED(transaction.Write(replacement)) || transaction.Commit(replacement.size() + 1u) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
        {
            return false;
        }
    }
    return targetIsPreservedAndTempIsGone();
}
} // namespace

// number of messages to send in each test
auto constexpr kMaxLoop = 50000;

namespace
{
std::wstring GenerateDiverseMessage(size_t index, Debug::InfoParam::Type type)
{
    static std::random_device rd;
    static std::mt19937 gen(rd());
    static std::uniform_int_distribution<> styleDist(0, 7);

    const std::wstring emojis[] = {L"🎮", L"🌟", L"💻", L"🚀", L"✨", L"🎯", L"🌈", L"💫", L"🔥", L"⚠️",
                                   L"🎪", L"🎨", L"🎭", L"🏆", L"🌺", L"💎", L"🔮", L"⭐", L"🌸", L"💥"};
    const std::wstring words[]  = {L"System",
                                   L"Process",
                                   L"Thread",
                                   L"Application",
                                   L"Monitor",
                                   L"Handler",
                                   L"Function",
                                   L"Method",
                                   L"こんにちは",
                                   L"مع السلامة",
                                   L"Пожалуйста",
                                   L"สวัสดี",
                                   L"你好",
                                   L"Bonjour",
                                   L"Hola"};

    // Type-specific prefixes for easy identification
    const wchar_t* typePrefix = L"[UNKNOWN]";
    switch (type)
    {
        case Debug::InfoParam::Type::Text: typePrefix = L"[TEXT]"; break;
        case Debug::InfoParam::Type::Error: typePrefix = L"[ERROR]"; break;
        case Debug::InfoParam::Type::Warning: typePrefix = L"[WARN]"; break;
        case Debug::InfoParam::Type::Info: typePrefix = L"[INFO]"; break;
        case Debug::InfoParam::Type::Perf: typePrefix = L"[PERF]"; break;
        case Debug::InfoParam::Type::Debug: typePrefix = L"[DEBUG]"; break;
        case Debug::InfoParam::Type::All: typePrefix = L"[ALL]"; break;
    }

    int style = styleDist(gen);
    std::wstring msg;

    switch (style)
    {
        case 0: // Short with emoji
            msg = std::format(L"{} {} [{}] Quick message", typePrefix, emojis[index % 20], index);
            break;
        case 1: // Medium with status code
            msg = std::format(L"{} [{}] Processing {} {} operation with status code 0x{:08X}",
                              typePrefix,
                              index,
                              words[index % 15],
                              (type == Debug::InfoParam::Type::Error)     ? L"ERROR"
                              : (type == Debug::InfoParam::Type::Warning) ? L"WARNING"
                                                                          : L"NORMAL",
                              static_cast<unsigned>(index * 0x1000 + static_cast<int>(type)));
            break;
        case 2: // Long with wrapping
            msg = std::format(L"{} [{}] This is a longer message that contains multiple words and should test text wrapping capabilities with {} and {} "
                              L"processing {} data structures",
                              typePrefix,
                              index,
                              words[gen() % 15],
                              words[gen() % 15],
                              words[gen() % 15]);
            break;
        case 3: // Multiline with 2 newlines
            msg = std::format(L"{} [{}] First line\nSecond line with data: {}\nThird line complete", typePrefix, index, words[index % 15]);
            break;
        case 4: // International characters
            msg = std::format(L"{} [{}] Mixed: Hello {} {} {} {} {}", typePrefix, index, words[10], words[11], words[12], words[13], emojis[index % 20]);
            break;
        case 5: // Simulated stack trace with 3 newlines
            msg = std::format(
                L"{} [{}] Exception in {}::{}() at line {}\n  Callstack: main->processData->validateInput\n  Context: {} processing\n  Module: {}.dll",
                typePrefix,
                index,
                words[gen() % 8],
                words[gen() % 8],
                100 + (index % 500),
                words[index % 8],
                words[gen() % 8]);
            break;
        case 6: // Performance metrics
            msg = std::format(L"{} [{}] Performance: {} took ({}ms) | Memory: {}KB | CPU: {}%",
                              typePrefix,
                              index,
                              words[index % 8],
                              static_cast<double>(index % 100) / 10.0,
                              1024 + (index % 4096),
                              15 + (index % 70));
            break;
        case 7: // Single newline embedded
            msg = std::format(L"{} [{}] Line one with {}\nLine two with data: {} complete", typePrefix, index, words[gen() % 8], words[gen() % 8]);
            break;
    }

    return msg;
}
void PrintProgress(size_t current, size_t total, std::wstring_view label, double elapsedSeconds, size_t messagesPerSecond)
{
    const int barWidth   = 50;
    const float progress = static_cast<float>(current) / static_cast<float>(total);
    const int pos        = static_cast<int>(barWidth * progress);

    std::wcout << L"\r" << label << L" [";
    for (int i = 0; i < barWidth; ++i)
    {
        if (i < pos)
            std::wcout << L"#";
        else if (i == pos)
            std::wcout << L">";
        else
            std::wcout << L"-";
    }
    std::wcout << L"] " << static_cast<int>(progress * 100.0f) << L"% (" << current << L"/" << total << L") | ";
    std::wcout << std::fixed << std::setprecision(2) << elapsedSeconds << L"s | ";
    std::wcout << messagesPerSecond << L" msg/s";
    std::wcout << L"    ";
    std::wcout.flush();
}

void RunDiverseMessages(size_t messageCount, std::wstring_view label)
{
    const Debug::InfoParam::Type types[] = {Debug::InfoParam::Type::Text,
                                            Debug::InfoParam::Type::Info,
                                            Debug::InfoParam::Type::Warning,
                                            Debug::InfoParam::Type::Error,
                                            Debug::InfoParam::Type::Perf,
                                            Debug::InfoParam::Type::Debug};

    std::wcout << L"\nStarting " << label << L" - " << messageCount << L" diverse messages\n";

    const auto startTime   = std::chrono::steady_clock::now();
    auto lastUpdateTime    = startTime;
    size_t lastUpdateCount = 0;

    for (size_t i = 0; i < messageCount; ++i)
    {
        const auto type        = types[i % (sizeof(types) / sizeof(types[0]))];
        const std::wstring msg = GenerateDiverseMessage(i, type);
        Debug::Out(type, L"{}", msg);

        if ((i % 100) == 0 || i == messageCount - 1)
        {
            const auto now             = std::chrono::steady_clock::now();
            const auto elapsed         = std::chrono::duration<double>(now - startTime).count();
            const auto sinceLastUpdate = std::chrono::duration<double>(now - lastUpdateTime).count();

            size_t messagesPerSecond = 0;
            if (sinceLastUpdate > 0.0)
            {
                const size_t messagesSinceUpdate = (i + 1) - lastUpdateCount;
                messagesPerSecond                = static_cast<size_t>(static_cast<double>(messagesSinceUpdate) / sinceLastUpdate);
            }

            PrintProgress(i + 1, messageCount, label, elapsed, messagesPerSecond);

            lastUpdateTime  = now;
            lastUpdateCount = i + 1;
        }

        if ((i % 512) == 0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }

    const auto endTime       = std::chrono::steady_clock::now();
    const auto totalDuration = std::chrono::duration<double>(endTime - startTime).count();
    const auto avgRate       = static_cast<size_t>(static_cast<double>(messageCount) / totalDuration);

    std::wcout << L"\n" << label << L" completed in " << std::fixed << std::setprecision(2) << totalDuration << L"s (avg: " << avgRate << L" msg/s)\n";
}

[[maybe_unused]] void RunHighRateBurst(size_t messageCount, std::wstring_view label)
{
    std::wcout << L"\nStarting " << label << L" - " << messageCount << L" messages\n";

    const auto startTime   = std::chrono::steady_clock::now();
    auto lastUpdateTime    = startTime;
    size_t lastUpdateCount = 0;

    for (size_t i = 0; i < messageCount; ++i)
    {
        Debug::Out(Debug::InfoParam::Type::Info, L"[{}] High-rate message {}", label, i);

        // Update progress every 100 messages
        if ((i % 100) == 0 || i == messageCount - 1)
        {
            const auto now             = std::chrono::steady_clock::now();
            const auto elapsed         = std::chrono::duration<double>(now - startTime).count();
            const auto sinceLastUpdate = std::chrono::duration<double>(now - lastUpdateTime).count();

            size_t messagesPerSecond = 0;
            if (sinceLastUpdate > 0.0)
            {
                const size_t messagesSinceUpdate = (i + 1) - lastUpdateCount;
                messagesPerSecond                = static_cast<size_t>(static_cast<double>(messagesSinceUpdate) / sinceLastUpdate);
            }

            PrintProgress(i + 1, messageCount, label, elapsed, messagesPerSecond);

            lastUpdateTime  = now;
            lastUpdateCount = i + 1;
        }

        if ((i % 512) == 0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }

    const auto endTime       = std::chrono::steady_clock::now();
    const auto totalDuration = std::chrono::duration<double>(endTime - startTime).count();
    const auto avgRate       = static_cast<size_t>(static_cast<double>(messageCount) / totalDuration);

    std::wcout << L"\n" << label << L" completed in " << std::fixed << std::setprecision(2) << totalDuration << L"s (avg: " << avgRate << L" msg/s)\n";
}
} // namespace

// Separate function with C++ objects (cannot use __try/__except)
static int RunMonitorTest()
{
    // Enable Unicode output for console
    _setmode(_fileno(stdout), _O_U16TEXT);

    std::wcout << L"=== MonitorTest Starting ===\n";
    std::wcout << L"Generating diverse messages with all severity types\n";

    if (! RunMonitorDiagnosticsGateSelfTest())
    {
        std::wcerr << L"Monitor diagnostics gate self-test failed\n";
        return 2;
    }

    if (! RunMonitorScrollbarModelSelfTest())
    {
        std::wcerr << L"Monitor scrollbar model self-test failed\n";
        return 2;
    }

    if (! RunMonitorDocumentBatchFilterSelfTest())
    {
        std::wcerr << L"Monitor document batch/filter self-test failed\n";
        return 2;
    }

    if (! RunMonitorDocumentRetentionSelfTest())
    {
        std::wcerr << L"Monitor document retention self-test failed\n";
        return 2;
    }

    if (! RunMonitorFileReaderSelfTest())
    {
        std::wcerr << L"Monitor bounded file-reader self-test failed\n";
        return 2;
    }

    if (! RunMonitorDocumentAtomicSaveSelfTest())
    {
        std::wcerr << L"Monitor document atomic-save self-test failed\n";
        return 2;
    }

    if (! RunLocalFileTransactionFaultSelfTest())
    {
        std::wcerr << L"Local file transaction fault self-test failed\n";
        return 2;
    }

    if (! RunUnicodeClipboardContractSelfTest())
    {
        std::wcerr << L"Unicode clipboard storage self-test failed\n";
        return 2;
    }

    if (! RunEtwConsumerShutdownSelfTest())
    {
        std::wcerr << L"ETW consumer shutdown self-test failed\n";
        return 2;
    }

    // Force ETW registration at startup
    const bool registered = Debug::detail::EnsureTraceLoggingRegistered();

    std::wcout << L"ETW Status: ";
    if (registered)
    {
        std::wcout << L"✓ Registered successfully\n";
        std::wcout << L"  Provider GUID: {440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}\n\n";
    }
    else
    {
        std::wcout << L"✗ Registration failed\n";
        std::wcout << L"  Warning: Events will not be sent\n";
        std::wcout << L"  Possible causes:\n";
        std::wcout << L"    • TraceLogging provider definition conflict\n";
        std::wcout << L"    • Run as Administrator and rebuild if needed\n";
        std::wcout << L"    • Check debug output for HRESULT error code\n\n";
    }

    const auto testStartTime = std::chrono::steady_clock::now();

    std::wcout << L"Running ETW-only tests...\n\n";

    std::wcout << L"Test 1: Diverse messages (burst-A)\n";
    RunDiverseMessages(kMaxLoop, L"burst-A");

    std::wcout << L"\nTest 2: Diverse messages (burst-B)\n";
    RunDiverseMessages(kMaxLoop, L"burst-B");

    std::wcout << L"\nTest 3: Diverse messages (mixed)\n";
    RunDiverseMessages(kMaxLoop, L"mixed");

    const auto testEndTime       = std::chrono::steady_clock::now();
    const auto totalTestDuration = std::chrono::duration<double>(testEndTime - testStartTime).count();

    const auto stats = Debug::GetTransportStats();

    std::wcout << L"\n=== ETW Transport Statistics ===\n";
    std::wcout << L"  Total duration:   " << std::fixed << std::setprecision(2) << totalTestDuration << L"s\n";
    std::wcout << L"  ETW written:      " << stats.etwWritten << L"\n";
    std::wcout << L"  ETW failed:       " << stats.etwFailed << L"\n";
    std::wcout << L"  Overall rate:     " << static_cast<size_t>(static_cast<double>(stats.etwWritten) / totalTestDuration) << L" msg/s\n";

    if (stats.etwFailed > 0)
    {
        std::wcout << L"\n⚠️  Warning: " << stats.etwFailed << L" ETW events failed to write\n";
        std::wcout << L"  This may indicate ETW registration issues or insufficient privileges\n";
    }

    std::wcout << L"\n💡 Launch RedSalamanderMonitor.exe to view these messages in real-time!\n";
    std::wcout << L"=========================\n";

    OutputDebugStringW(std::format(L"ETW Transport stats - etwWritten: {}, etwFailed: {}\n", stats.etwWritten, stats.etwFailed).c_str());

    return 0;
}

int main(int argc, char** argv)
{
    if (argc > 1 && argv[1] != nullptr && lstrcmpA(argv[1], "--diagnostics-gate-selftest") == 0)
    {
        return RunMonitorDiagnosticsGateSelfTest() ? 0 : 2;
    }
    if (argc > 1 && argv[1] != nullptr && lstrcmpA(argv[1], "--scrollbar-model-selftest") == 0)
    {
        return RunMonitorScrollbarModelSelfTest() ? 0 : 2;
    }
    if (argc > 1 && argv[1] != nullptr && lstrcmpA(argv[1], "--document-model-selftest") == 0)
    {
        return RunMonitorDocumentBatchFilterSelfTest() && RunMonitorDocumentRetentionSelfTest() && RunMonitorFileReaderSelfTest() &&
                       RunMonitorDocumentAtomicSaveSelfTest() && RunLocalFileTransactionFaultSelfTest() && RunUnicodeClipboardContractSelfTest() &&
                       RunEtwConsumerShutdownSelfTest()
                   ? 0
                   : 2;
    }

    // Use SEH to catch all exceptions (no C++ objects in this scope)
    __try
    {
        return RunMonitorTest();
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);

        wchar_t errorMsg[512]{};
        const auto formatted = std::format_to_n(errorMsg,
                                                _countof(errorMsg) - 1u,
                                                L"Fatal Exception in MonitorTest\n\n"
                                                L"Exception: {} (0x{:08X})\n\n"
                                                L"The application will now terminate.",
                                                exceptionName,
                                                static_cast<unsigned>(exceptionCode));
        *formatted.out       = L'\0';

        // Try to output to both console and debug output
        std::wcerr << errorMsg << L"\n";
        OutputDebugStringW(errorMsg);
        MessageBoxW(nullptr, errorMsg, L"MonitorTest - Fatal Error", MB_OK | MB_ICONERROR);

        return -1;
    }
}

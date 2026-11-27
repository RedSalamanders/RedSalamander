// CrashHandlingTests: unit tests for the crash-quarantine decision logic AND the pure crash-handler
// path/marker helpers.
//
// The quarantine tests exercise ONLY the extracted pure decision function
// CrashQuarantine::SelectPluginsToDisable (exposed under ENABLE_TESTS) and the apply helper
// DisablePluginIdInSettingsForTest. They never call OfferPreviousCrashRecovery with a
// real marker present, so no SessionState is read, no known-folder path is resolved, and no UI /
// HostShowPrompt is invoked -- guaranteeing no UI appears.
//
// The crash-handler tests (Step 5) exercise CrashHandler.cpp's pure helpers via the *ForTest seams:
// dump-path naming (BuildDumpPath) and the marker write<->read round-trip (WriteMarkerFile /
// ReadMarkerDumpPath), the once-only claim, and retention in the governed scratch root.
// Fatal entry points run only in separate child processes with their own artifact directories.
//
// The `markerExists` boolean argument stands in for "a crash marker file exists at the crash path":
// the production code computes it via std::filesystem::exists(GetCrashMarkerPath()) and passes the
// result in. Testing the bool directly keeps the unit free of file I/O and the un-injectable
// AppDataPaths root.

#include <chrono>
#include <cstdlib>
#include <exception>
#include <filesystem>
#include <format>
#include <iostream>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "CrashHandler.h"
#include "CrashQuarantine.h"
#include "SettingsStore.h"
#include "TestSupport/TestSupport.Tests.ChildProcess.h"
#include "TestSupport/TestSupport.Tests.FailureReports.h"
#include "TestSupport/TestSupport.Tests.Runner.h"

namespace
{
constexpr std::wstring_view kCrashHandlingHarnessSegment{L"crash-handling"};

void Check(bool condition, const wchar_t* message, bool& success) noexcept
{
    if (! condition)
    {
        std::wcerr << L"[ FAILED  ] " << message << L"\n";
        success = false;
        return;
    }

    std::wcout << L"[       OK ] " << message << L"\n";
}

[[nodiscard]] bool Contains(const std::vector<std::wstring>& v, std::wstring_view id) noexcept
{
    for (const std::wstring& s : v)
    {
        if (s == id)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool DisabledListContains(const Common::Settings::Settings& settings, std::wstring_view id) noexcept
{
    return Contains(settings.plugins.disabledPluginIds, id);
}

[[nodiscard]] std::filesystem::path AcquireCrashHandlingTestSandbox(std::wstring_view caseName, std::error_code& ec) noexcept
{
    return RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = kCrashHandlingHarnessSegment, .leafSegment = caseName, .fallbackRunIdPrefix = L"crash-handling", .cleanExisting = false}, ec);
}

// Case 1: No marker file -> decision returns "no offer"; settings untouched.
bool TestNoMarkerMakesNoOffer() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    const std::vector<std::wstring> active{L"vendor/some-fs"};
    const size_t disabledBefore = settings.plugins.disabledPluginIds.size();

    const std::vector<std::wstring> result = CrashQuarantine::SelectPluginsToDisable(false, active, settings);

    Check(result.empty(), L"no marker => empty offer", success);
    Check(settings.plugins.disabledPluginIds.size() == disabledBefore, L"no marker => disabled list unchanged", success);
    Check(settings.plugins.currentFileSystemPluginId == L"builtin/file-system", L"no marker => current plugin unchanged", success);

    return success;
}

// Case 2: Marker present + a last-active filesystem plugin -> offer; accept disables it; decline leaves it.
bool TestMarkerWithActivePluginOffersAndApplies() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    settings.plugins.currentFileSystemPluginId = L"vendor/some-fs";
    const std::vector<std::wstring> active{L"vendor/some-fs"};

    const std::vector<std::wstring> offer = CrashQuarantine::SelectPluginsToDisable(true, active, settings);
    Check(offer.size() == 1u && Contains(offer, L"vendor/some-fs"), L"marker+active => offers the active plugin", success);

    // Simulate decline: do nothing -> settings unchanged.
    Check(! DisabledListContains(settings, L"vendor/some-fs"), L"decline => plugin not disabled", success);
    Check(settings.plugins.currentFileSystemPluginId == L"vendor/some-fs", L"decline => current plugin unchanged", success);

    // Simulate accept: apply the offer the way production does on YES.
    for (const std::wstring& id : offer)
    {
        CrashQuarantine::DisablePluginIdInSettingsForTest(id, settings);
    }
    Check(DisabledListContains(settings, L"vendor/some-fs"), L"accept => plugin added to disabled list", success);
    Check(settings.plugins.currentFileSystemPluginId.empty(), L"accept => current plugin cleared (was the disabled one)", success);

    return success;
}

// Case 2b: Apply is idempotent and does not duplicate ids; current plugin only cleared when it matches.
bool TestApplyIdempotentAndSelective() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    settings.plugins.currentFileSystemPluginId = L"vendor/keep-me";

    // Disable a plugin that is NOT the current one -> current must remain.
    CrashQuarantine::DisablePluginIdInSettingsForTest(L"vendor/other-fs", settings);
    Check(DisabledListContains(settings, L"vendor/other-fs"), L"apply non-current => added to disabled", success);
    Check(settings.plugins.currentFileSystemPluginId == L"vendor/keep-me", L"apply non-current => current plugin preserved", success);

    // Re-apply the same id -> no duplicate.
    CrashQuarantine::DisablePluginIdInSettingsForTest(L"vendor/other-fs", settings);
    size_t count = 0;
    for (const std::wstring& s : settings.plugins.disabledPluginIds)
    {
        if (s == L"vendor/other-fs")
        {
            ++count;
        }
    }
    Check(count == 1u, L"apply twice => no duplicate disabled entry", success);

    // Empty id is a no-op.
    const size_t before = settings.plugins.disabledPluginIds.size();
    CrashQuarantine::DisablePluginIdInSettingsForTest(L"", settings);
    Check(settings.plugins.disabledPluginIds.size() == before, L"apply empty id => no-op", success);

    return success;
}

// Case 3: "Garbage"/already-disabled inputs do not crash and produce no offer for ids already disabled.
bool TestAlreadyDisabledFilteredOut() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    settings.plugins.disabledPluginIds.emplace_back(L"vendor/already-off");
    // Active ids include one already disabled, one empty (garbage), and one fresh.
    const std::vector<std::wstring> active{L"vendor/already-off", L"", L"vendor/fresh-fs"};

    const std::vector<std::wstring> offer = CrashQuarantine::SelectPluginsToDisable(true, active, settings);

    Check(! Contains(offer, L"vendor/already-off"), L"already-disabled id excluded from offer", success);
    Check(! Contains(offer, L""), L"empty id excluded from offer", success);
    Check(offer.size() == 1u && Contains(offer, L"vendor/fresh-fs"), L"only the fresh id is offered", success);

    return success;
}

// Case 3b: Case-insensitive matching against already-disabled ids (mirrors OrdinalString::EqualsNoCase).
bool TestAlreadyDisabledCaseInsensitive() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    settings.plugins.disabledPluginIds.emplace_back(L"Vendor/Mixed-Case");
    const std::vector<std::wstring> active{L"vendor/mixed-case"};

    const std::vector<std::wstring> offer = CrashQuarantine::SelectPluginsToDisable(true, active, settings);
    Check(offer.empty(), L"case-insensitive already-disabled => no offer", success);

    return success;
}

// Case 5: Marker present but NO last-active plugin (empty/garbage-only active list) -> no offer, no crash.
bool TestMarkerWithNoActivePluginMakesNoOffer() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;

    // Empty active list.
    const std::vector<std::wstring> emptyActive;
    Check(CrashQuarantine::SelectPluginsToDisable(true, emptyActive, settings).empty(), L"marker + no active plugin => no offer", success);

    // Active list of only empty (garbage) ids.
    const std::vector<std::wstring> blankActive{L"", L""};
    Check(CrashQuarantine::SelectPluginsToDisable(true, blankActive, settings).empty(), L"marker + only-empty active ids => no offer", success);

    return success;
}

// Multiple distinct active plugins are all offered (plural prompt path in production).
bool TestMultipleActivePluginsAllOffered() noexcept
{
    bool success = true;

    Common::Settings::Settings settings;
    const std::vector<std::wstring> active{L"vendor/fs-a", L"vendor/fs-b"};

    const std::vector<std::wstring> offer = CrashQuarantine::SelectPluginsToDisable(true, active, settings);
    Check(offer.size() == 2u && Contains(offer, L"vendor/fs-a") && Contains(offer, L"vendor/fs-b"), L"two active plugins => both offered", success);

    return success;
}
// ----------------------------------------------------------------- Crash-handler paths (Step 5)

[[nodiscard]] bool HasInvalidFilenameChar(const std::wstring& name) noexcept
{
    for (const wchar_t c : name)
    {
        if (c == L'<' || c == L'>' || c == L':' || c == L'"' || c == L'/' || c == L'\\' || c == L'|' || c == L'?' || c == L'*')
        {
            return true;
        }
    }
    return false;
}

// Dump path is under the given root, named *.dmp, RedSalamander-...-p<pid>, no invalid chars.
bool TestBuildDumpPathShape(const std::filesystem::path& root) noexcept
{
    bool success                     = true;
    const std::filesystem::path dump = CrashHandler::BuildDumpPathForTest(root);

    Check(dump.parent_path() == root, L"BuildDumpPath: parent is the given root", success);
    Check(dump.extension() == L".dmp", L"BuildDumpPath: extension is .dmp", success);

    const std::wstring name = dump.filename().wstring();
    Check(name.starts_with(L"RedSalamander-report-v1-"), L"BuildDumpPath: filename identifies a retained report", success);

    const std::wstring pidTag = L"-p" + std::to_wstring(GetCurrentProcessId());
    Check(name.find(pidTag) != std::wstring::npos, L"BuildDumpPath: filename embeds -p<pid>", success);
    Check(! HasInvalidFilenameChar(name), L"BuildDumpPath: filename has no invalid path characters", success);
    return success;
}

bool TestOwnedReportRetention(const std::filesystem::path& root)
{
    bool success         = true;
    const auto directory = root / L"retention";
    std::error_code error;
    std::filesystem::create_directory(directory, error);
    if (error)
    {
        Check(false, L"crash retention: create the test folder", success);
        return false;
    }
    Check(CrashHandler::PruneOldReportsForTest(directory) == 0, L"crash retention: empty folder has nothing to prune", success);

    const auto writePair = [&](const std::wstring& name)
    {
        auto dump            = directory / name;
        const bool wroteDump = SUCCEEDED(CrashHandler::WriteMarkerFileForTest(dump, L"dump fixture"));
        dump.replace_extension(L".txt");
        return wroteDump && SUCCEEDED(CrashHandler::WriteMarkerFileForTest(dump, L"sidecar fixture"));
    };
    Check(writePair(L"RedSalamander-20261127-120000-p7.dmp"), L"crash retention: seed a legacy pair", success);
    Check(writePair(L"RedSalamander-report-v1-unrelated.dmp"), L"crash retention: seed an unrelated pair", success);
    const auto name = [](int stamp) { return std::format(L"RedSalamander-report-v1-{:020}-p7.dmp", stamp); };
    for (int stamp = 1; stamp <= 10; ++stamp)
    {
        Check(writePair(name(stamp)), L"crash retention: seed an owned pair", success);
    }
    Check(CrashHandler::PruneOldReportsForTest(directory) == 0, L"crash retention: keep ten owned reports", success);
    Check(writePair(name(11)), L"crash retention: seed the eleventh owned report", success);
    const auto pending = directory / name(1);
    Check(
        CrashHandler::PruneOldReportsForTest(directory, pending.native()) == 0, L"crash retention: keep an older report while its notice is pending", success);
    Check(CrashHandler::PruneOldReportsForTest(directory) == 1, L"crash retention: prune one owned report", success);

    const auto pairExists = [&](const std::wstring& pairName, bool expected)
    {
        auto path             = directory / pairName;
        const bool dumpExists = std::filesystem::exists(path);
        path.replace_extension(L".txt");
        return dumpExists == expected && std::filesystem::exists(path) == expected;
    };
    Check(pairExists(name(1), false), L"crash retention: remove the oldest dump and sidecar", success);
    for (int stamp = 2; stamp <= 11; ++stamp)
    {
        Check(pairExists(name(stamp), true), L"crash retention: preserve a newer pair", success);
    }
    Check(pairExists(L"RedSalamander-20261127-120000-p7.dmp", true), L"crash retention: preserve the legacy pair", success);
    Check(pairExists(L"RedSalamander-report-v1-unrelated.dmp", true), L"crash retention: preserve unrelated evidence", success);
    return success;
}

// Marker write -> read round-trip returns the exact dump path (BOM stripped).
bool TestMarkerRoundTrip(const std::filesystem::path& root) noexcept
{
    bool success                       = true;
    const std::filesystem::path marker = root / L"last_crash.txt";
    const std::wstring dumpPath        = L"C:\\dumps\\RedSalamander-x.dmp";

    Check(SUCCEEDED(CrashHandler::WriteMarkerFileForTest(marker, dumpPath)), L"WriteMarkerFile: succeeds for a valid path", success);
    Check(CrashHandler::ReadMarkerDumpPathForTest(marker) == dumpPath, L"marker round-trip: read-back equals written dump path", success);
    return success;
}

// Trailing whitespace/newlines in the stored path are trimmed on read.
bool TestMarkerReadTrimsTrailingWhitespace(const std::filesystem::path& root) noexcept
{
    bool success                       = true;
    const std::filesystem::path marker = root / L"last_crash_ws.txt";

    Check(SUCCEEDED(CrashHandler::WriteMarkerFileForTest(marker, L"C:\\dumps\\y.dmp\r\n  ")), L"WriteMarkerFile: succeeds (whitespace payload)", success);
    Check(CrashHandler::ReadMarkerDumpPathForTest(marker) == L"C:\\dumps\\y.dmp", L"marker read: trailing CR/LF/space trimmed", success);
    return success;
}

// Reading a non-existent marker returns empty (no crash; noexcept).
bool TestMarkerReadMissingIsEmpty(const std::filesystem::path& root) noexcept
{
    bool success = true;
    Check(CrashHandler::ReadMarkerDumpPathForTest(root / L"does_not_exist.txt").empty(), L"marker read: missing file -> empty", success);
    return success;
}

// An empty marker path is rejected (characterizes WriteMarkerFile's guard).
bool TestWriteMarkerEmptyPathRejected() noexcept
{
    bool success = true;
    Check(CrashHandler::WriteMarkerFileForTest(std::filesystem::path{}, L"C:\\dumps\\z.dmp") == E_INVALIDARG,
          L"WriteMarkerFile: empty marker path -> E_INVALIDARG",
          success);
    return success;
}

// With %LOCALAPPDATA% stubbed empty, the marker path short-circuits to empty (no crash).
bool TestMarkerPathEmptyWhenBaseUnavailable() noexcept
{
    bool success = true;
    Check(CrashHandler::GetCrashMarkerPathForTest().empty(), L"GetCrashMarkerPath: empty base -> empty path", success);
    return success;
}

bool TestPendingClaim(const std::filesystem::path& root)
{
    bool success      = true;
    const auto marker = root / L"last_crash.txt";
    const auto dump   = CrashHandler::BuildDumpPathForTest(root);
    if (dump.empty() || FAILED(CrashHandler::WriteMarkerFileForTest(dump, L"dump fixture")) ||
        FAILED(CrashHandler::WriteMarkerFileForTest(marker, dump.native())))
    {
        Check(false, L"pending claim: prepare the isolated dump and marker", success);
        return false;
    }
    Check(! CrashHandler::TakePendingReportForTest(root, false) && std::filesystem::exists(marker),
          L"pending claim: a headless start leaves the marker",
          success);
    const auto report = CrashHandler::TakePendingReportForTest(root, true);
    Check(report && report->directory == root && report->dumpPath == dump.native(), L"pending claim: the first interactive start receives the dump", success);
    Check(! std::filesystem::exists(marker) && ! CrashHandler::TakePendingReportForTest(root, true),
          L"pending claim: a second interactive start receives no report",
          success);
    Check(std::filesystem::exists(dump), L"pending claim: consuming the notice preserves the dump", success);
    return success;
}

int RunCrashProbeChild(std::wstring_view mode, std::wstring_view leaf)
{
    if (mode != L"single" && mode != L"workers" && mode != L"abort")
        return 2;
    std::error_code error;
    const auto directory = RedSalamander::TestSupport::AcquireTestDirectory({.harnessSegment      = kCrashHandlingHarnessSegment,
                                                                             .leafSegment         = leaf,
                                                                             .fallbackRunIdPrefix = L"crash-handling",
                                                                             .kind                = RedSalamander::TestSupport::TestDirectoryKind::Artifacts,
                                                                             .cleanExisting       = false},
                                                                            error);
    if (error || directory.empty() || FAILED(CrashHandler::InstallForProbe(directory)))
        return 2;
    if (mode == L"abort")
        std::abort();
    CrashHandler::TriggerCrashTest(mode == L"workers");
}

bool TestCrashProbeChildren()
{
    using namespace RedSalamander::TestSupport;
    bool success = true;
    std::error_code error;
    const auto sandbox = ResolveTestSandboxBase(error);
    if (error || sandbox.empty())
    {
        Check(false, L"crash probe: locate the governed sandbox", success);
        return false;
    }
    const auto executable = Common::Testing::GetCurrentModulePath(error);
    const auto runId      = ResolveTestRunId(L"crash-handling");
    if (error || executable.empty() || runId.empty())
    {
        Check(false, L"crash probe: resolve the child executable and run identity", success);
        return false;
    }
#if defined(RS_ASAN_DEBUG_BUILD) || defined(__SANITIZE_ADDRESS__)
    constexpr DWORD probeCode = 0xE000CAFEu;
#else
    constexpr DWORD probeCode = EXCEPTION_ACCESS_VIOLATION;
#endif
    for (const std::wstring_view mode : {L"single", L"workers", L"abort"})
    {
        std::wcout << L"[ crash probe ] " << mode << L'\n';
        const auto leaf   = std::format(L"{}-{}-{}", mode, GetCurrentProcessId(), GetTickCount64());
        const auto result = RunChildProcess(
            {.executablePath       = executable,
             .arguments            = {L"--crash-probe-fixture", std::wstring(mode), leaf},
             .workingDirectory     = executable.parent_path(),
             .timeout              = std::chrono::seconds(45),
             .environmentOverrides = {{std::wstring(kTestRootEnvironmentVariable), sandbox.native()}, {std::wstring(kTestRunIdEnvironmentVariable), runId}}});
        const DWORD expectedCode = mode == L"abort" ? 0xE000CAF4u : probeCode;
        Check(result.Completed() && result.exitCode == expectedCode, L"crash probe: child terminates with the expected code", success);

        const auto directory = sandbox / L"runs" / runId / L"artifacts" / kCrashHandlingHarnessSegment / leaf;
        const auto dumpPath  = CrashHandler::ReadMarkerDumpPathForTest(directory / L"last_crash.txt");
        size_t dumpCount     = 0;
        bool matchesMarker   = false;
        for (std::filesystem::directory_iterator entry(directory, error), end; ! error && entry != end; entry.increment(error))
        {
            if (entry->path().extension() == L".dmp")
            {
                ++dumpCount;
                std::error_code fileError;
                matchesMarker = entry->path().native() == dumpPath && entry->file_size(fileError) != 0 && ! fileError;
            }
        }
        Check(! error && dumpCount == 1 && matchesMarker, L"crash probe: one nonempty dump and its exact pending marker", success);
    }
    return success;
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    if (argc == 4 && std::wstring_view(argv[1]) == L"--crash-probe-fixture")
        return RunCrashProbeChild(argv[2], argv[3]);
    RedSalamander::TestSupport::RouteFailureReportsAwayFromDialogs();

    bool success = true;

    // Quarantine decision (Step 4).
    success &= TestNoMarkerMakesNoOffer();
    success &= TestMarkerWithActivePluginOffersAndApplies();
    success &= TestApplyIdempotentAndSelective();
    success &= TestAlreadyDisabledFilteredOut();
    success &= TestAlreadyDisabledCaseInsensitive();
    success &= TestMarkerWithNoActivePluginMakesNoOffer();
    success &= TestMultipleActivePluginsAllOffered();

    // Crash-handler paths (Step 5) — under a TestSandbox case root, cleaned up.
    std::error_code ec;
    const auto scratchName               = std::format(L"marker-files-{}-{}", GetCurrentProcessId(), GetTickCount64());
    const std::filesystem::path tempRoot = AcquireCrashHandlingTestSandbox(scratchName, ec);
    if (ec || tempRoot.empty())
    {
        std::wcerr << L"[ FAILED  ] could not create temp directory: " << tempRoot.wstring() << L"\n";
        return 1;
    }

    success &= TestBuildDumpPathShape(tempRoot);
    success &= TestOwnedReportRetention(tempRoot);
    success &= TestMarkerRoundTrip(tempRoot);
    success &= TestMarkerReadTrimsTrailingWhitespace(tempRoot);
    success &= TestMarkerReadMissingIsEmpty(tempRoot);
    success &= TestWriteMarkerEmptyPathRejected();
    success &= TestMarkerPathEmptyWhenBaseUnavailable();
    success &= TestPendingClaim(tempRoot);
    success &= TestCrashProbeChildren();

    if (success)
    {
        std::filesystem::remove_all(tempRoot, ec);
        std::wcout << L"All CrashHandlingTests passed.\n";
        return 0;
    }

    std::wcerr << L"CrashHandlingTests FAILED.\n";
    return 1;
}

// CrashHandlingTests: unit tests for the crash-quarantine decision logic AND the pure crash-handler
// path/marker helpers.
//
// The quarantine tests exercise ONLY the extracted pure decision function
// CrashQuarantine::SelectPluginsToDisable (exposed under ENABLE_TESTS) and the apply helper
// DisablePluginIdInSettingsForTest. They never call OfferPluginDisableIfPreviousCrashDetected with a
// real marker present, so no SessionState is read, no known-folder path is resolved, and no UI /
// HostShowPrompt is invoked -- guaranteeing no UI appears.
//
// The crash-handler tests (Step 5) exercise CrashHandler.cpp's pure helpers via the *ForTest seams:
// dump-path naming (BuildDumpPath) and the marker write<->read round-trip (WriteMarkerFile /
// ReadMarkerDumpPath). They never invoke the fused SEH/minidump/UI paths (Install,
// WriteDumpForException, ShowPreviousCrashUiIfPresent, TriggerCrashTest) and write only under the
// unified TestSandbox scratch root.
//
// The `markerExists` boolean argument stands in for "a crash marker file exists at the crash path":
// the production code computes it via std::filesystem::exists(GetCrashMarkerPath()) and passes the
// result in. Testing the bool directly keeps the unit free of file I/O and the un-injectable
// AppDataPaths root.

#include <exception>
#include <filesystem>
#include <iostream>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#include <Windows.h>

#include "CrashHandler.h"
#include "CrashQuarantine.h"
#include "SettingsStore.h"
#include "TestSupport/TestSupport.h"

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
    Check(name.rfind(L"RedSalamander-", 0) == 0, L"BuildDumpPath: filename starts with RedSalamander-", success);

    const std::wstring pidTag = L"-p" + std::to_wstring(GetCurrentProcessId());
    Check(name.find(pidTag) != std::wstring::npos, L"BuildDumpPath: filename embeds -p<pid>", success);
    Check(! HasInvalidFilenameChar(name), L"BuildDumpPath: filename has no invalid path characters", success);
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
} // namespace

int wmain()
{
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
    const std::filesystem::path tempRoot = AcquireCrashHandlingTestSandbox(L"marker-files", ec);
    std::filesystem::remove_all(tempRoot, ec); // clear any residue from a prior run
    std::filesystem::create_directories(tempRoot, ec);
    if (ec)
    {
        std::wcerr << L"[ FAILED  ] could not create temp directory: " << tempRoot.wstring() << L"\n";
        return 1;
    }

    success &= TestBuildDumpPathShape(tempRoot);
    success &= TestMarkerRoundTrip(tempRoot);
    success &= TestMarkerReadTrimsTrailingWhitespace(tempRoot);
    success &= TestMarkerReadMissingIsEmpty(tempRoot);
    success &= TestWriteMarkerEmptyPathRejected();
    success &= TestMarkerPathEmptyWhenBaseUnavailable();

    std::filesystem::remove_all(tempRoot, ec); // cleanup hygiene (exe may be run repeatedly)

    if (success)
    {
        std::wcout << L"All CrashHandlingTests passed.\n";
        return 0;
    }

    std::wcerr << L"CrashHandlingTests FAILED.\n";
    return 1;
}

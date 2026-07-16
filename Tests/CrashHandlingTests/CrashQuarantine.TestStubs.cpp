// Stub translation unit for CrashHandlingTests.
//
// CrashQuarantine.cpp is compiled directly into this test exe so the extracted pure decision
// function can be unit-tested. That TU references several heavy app-side externals at link time
// (even though the tests only ever call the pure function and never reach them). This file provides
// trivial definitions for those externals so the closure stays tiny -- no SessionState, no UI/
// HostServices, no SHGetKnownFolderPath. Modelled on Tests\PerformanceTests2\PluginManager.TestStubs.cpp.

#include <filesystem>
#include <optional>
#include <string_view>

// Define the TraceLogging provider exactly once for this test exe (CrashQuarantine.cpp -> Helpers.h
// references g_RedSalamanderProvider via Debug logging). Mirrors LocalizationTests.cpp.
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "AppDataPaths.h"
#include "HostServices.h" // declares HostShowPrompt; pulls in PlugInterfaces/Host.h for the POD types
#include "SessionState.h"

// Declared extern in Framework.h (included transitively by CrashQuarantine.cpp). Define it here.
PCWSTR REDSALAMANDER_TEXT_VERSION = L"test";

namespace AppDataPaths
{
std::filesystem::path GetLocalAppDataPath() noexcept
{
    // Empty -> GetCrashMarkerPath() returns {} and OfferPluginDisableIfPreviousCrashDetected early-returns.
    // The tests never call that path; this only needs to resolve the link symbol.
    return {};
}
} // namespace AppDataPaths

namespace SessionState
{
std::optional<State> TryRead() noexcept
{
    // The tests call SelectPluginsToDisable directly with explicit inputs, so this stub is never
    // exercised; it exists solely to satisfy the linker.
    return std::nullopt;
}
} // namespace SessionState

// Global (non-namespaced) free function. Do NOT restate the `cookie = nullptr` default arg here --
// it lives on the declaration in HostServices.h. The tests never reach this; no UI is shown.
HRESULT HostShowPrompt(const HostPromptRequest& /*request*/, void* /*cookie*/, HostPromptResult* result) noexcept
{
    if (result != nullptr)
    {
        *result = HOST_PROMPT_RESULT_NONE;
    }
    return E_NOTIMPL;
}

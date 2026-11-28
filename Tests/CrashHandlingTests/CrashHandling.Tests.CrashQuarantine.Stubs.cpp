// Stub translation unit for CrashHandlingTests.
//
// CrashQuarantine.cpp is compiled directly into this test exe so the extracted pure decision
// function can be unit-tested. That TU references several heavy app-side externals at link time
// (even though the tests only ever call the pure function and never reach them). This file provides
// trivial definitions for the application state and known-folder services. The tests never present the dialog.

#include <filesystem>
#include <optional>
#include <string_view>

// Define the TraceLogging provider exactly once for this test exe (CrashQuarantine.cpp -> Helpers.h
// references g_RedSalamanderProvider via Debug logging). Mirrors Localization.Tests.Runner.cpp.
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "AppDataPaths.h"
#include "SessionState.h"

// Declared extern in Framework.h (included transitively by CrashQuarantine.cpp). Define it here.
PCWSTR REDSALAMANDER_TEXT_VERSION = L"test";

namespace AppDataPaths
{
std::filesystem::path GetLocalAppDataPath() noexcept
{
    // Standalone tests have no product crash directory.
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

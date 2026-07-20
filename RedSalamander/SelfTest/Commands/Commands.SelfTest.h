#pragma once

#ifdef ENABLE_TESTS

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "SelfTestCommon.h"
#include "SettingsStore.h"

#ifdef ENABLE_TESTS
void DebugResetConnectionManagerConnectNavigation() noexcept;
void DebugSetConnectionManagerConnectNavigationSuppressed(bool suppressed) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerConnectNavigation(uint8_t& outPane, std::wstring& outName) noexcept;

using SessionEndSettingsWriterForSelfTest = HRESULT (*)(std::wstring_view appId, const Common::Settings::Settings& settings) noexcept;

struct SessionEndSettingsDebugSnapshot
{
    uint32_t writerCallCount         = 0u;
    uint32_t normalTeardownCallCount = 0u;
    uint64_t durationUs              = 0u;
    HRESULT lastResult               = S_OK;
    Common::Settings::Settings settings;
};

void DebugResetSessionEndSettingsSaveForSelfTest() noexcept;
void DebugSetSessionEndSettingsWriterForSelfTest(SessionEndSettingsWriterForSelfTest writer) noexcept;
[[nodiscard]] SessionEndSettingsDebugSnapshot DebugGetSessionEndSettingsSnapshotForSelfTest();
#endif

namespace CommandsSelfTest
{
std::vector<std::wstring> ListCases(const SelfTest::SelfTestOptions& options = {}) noexcept;
[[nodiscard]] bool Run(HWND mainWindow, const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
} // namespace CommandsSelfTest

#endif // ENABLE_TESTS

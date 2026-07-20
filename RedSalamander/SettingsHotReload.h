#pragma once

#include <filesystem>
#include <optional>
#include <span>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "AppTheme.h"
#include "SettingsStore.h"

struct PluginConfigurationSchemaSource;

namespace SettingsHotReload
{
struct SettingsFileChangedPayload
{
    ULONGLONG tickCount = 0;
};

enum class ChangedSettingsStatus : uint8_t
{
    NoChange,
    Loaded,
    Missing,
    Invalid,
    Error,
};

struct ChangedSettingsLoadResult
{
    ChangedSettingsStatus status = ChangedSettingsStatus::NoChange;
    HRESULT hr                   = S_OK;
    Common::Settings::Settings settings{};
    std::optional<Common::Settings::SettingsFileStamp> stamp;
};

enum class ExternalReloadChoice : uint8_t
{
    ReloadFromDisk,
    KeepEditing,
};

enum class StaleSaveChoice : uint8_t
{
    OverwriteCurrent,
    ReloadFromDisk,
    Cancel,
};

HRESULT Start(HWND targetWindow, std::wstring_view appId) noexcept;
void Stop() noexcept;
void StopWatchingForProcessShutdown() noexcept;
void BeginProcessShutdown() noexcept;

HRESULT SaveSettingsAndSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept;
HRESULT SaveSettingsAndSchema(std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept;
// Serializes a final settings-only snapshot with earlier saves, rejects later submissions,
// and waits only for the caller-supplied session-end deadline.
HRESULT SaveSettingsForSessionEnd(std::wstring_view appId,
                                  const Common::Settings::Settings& settings,
                                  DWORD timeoutMs) noexcept;
HRESULT SaveSettingsAndSchemaForProcessShutdown(std::wstring_view appId,
                                                Common::Settings::Settings& settings,
                                                DWORD timeoutMs) noexcept;
HRESULT SaveSettingsAndSchemaForProcessShutdown(std::wstring_view appId,
                                                Common::Settings::Settings& settings,
                                                std::span<const PluginConfigurationSchemaSource> pluginSchemas,
                                                DWORD timeoutMs) noexcept;
// Explicit replacement entry point for a settings file written by a newer schema. Callers must
// obtain user approval before invoking it. The save block is cleared only after the source backup succeeds.
HRESULT ReplaceBlockedSettingsAndSchema(std::wstring_view appId,
                                        Common::Settings::Settings& settings,
                                        std::filesystem::path& backupPath) noexcept;
HRESULT QueueSettingsSave(std::wstring_view appId,
                          const Common::Settings::Settings& settings,
                          std::wstring_view telemetryMetric,
                          std::wstring_view telemetryContext) noexcept;
bool FlushQueuedSettingsSaves(DWORD timeoutMs) noexcept;

ChangedSettingsLoadResult TryLoadChangedSettings() noexcept;
void MarkAppliedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept;
void MarkRejectedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept;

void RegisterParticipant(HWND hwnd) noexcept;
void UnregisterParticipant(HWND hwnd) noexcept;
void NotifyParticipants() noexcept;

HRESULT PromptExternalReloadConflict(HWND targetWindow, std::wstring_view editorName, ExternalReloadChoice& outChoice) noexcept;
HRESULT PromptStaleSaveConflict(HWND targetWindow, std::wstring_view editorName, StaleSaveChoice& outChoice) noexcept;

void ShowInvalidReloadAlert(const std::filesystem::path& settingsPath) noexcept;
void ClearInvalidReloadAlert() noexcept;
void ApplyUiPreferencesToTheme(const Common::Settings::Settings& settings, AppTheme& theme) noexcept;
[[nodiscard]] AppTheme ResolveDialogThemeFromSettings(const Common::Settings::Settings& settings) noexcept;

[[nodiscard]] Common::Settings::Settings MergeDiskSettingsWithRuntimeSession(const Common::Settings::Settings& diskSettings,
                                                                             const Common::Settings::Settings& runtimeSettings,
                                                                             std::span<const std::wstring_view> runtimeWindowIds) noexcept;

#ifdef ENABLE_TESTS
struct SettingsSaveDebugSnapshot
{
    uint64_t queuedGeneration    = 0;
    uint64_t completedGeneration = 0;
    uint64_t coalescedCount      = 0;
    DWORD lastQueueThreadId      = 0;
    DWORD lastSaveThreadId       = 0;
    bool pending                 = false;
    bool saveInProgress          = false;
};

void DebugSetChangeNotificationOpenFailuresForSelfTest(uint32_t failureCount, DWORD lastError) noexcept;
void DebugSetSettingsSaveDelayForSelfTest(DWORD delayMs) noexcept;
void DebugSetSettingsSavePostWriteDelayForSelfTest(DWORD delayMs) noexcept;
void DebugSetSettingsReloadPostStampDelayForSelfTest(DWORD delayMs) noexcept;
bool DebugIsSettingsReloadPostStampDelayActiveForSelfTest() noexcept;
SettingsSaveDebugSnapshot DebugGetSettingsSaveSnapshotForSelfTest() noexcept;
#endif
} // namespace SettingsHotReload

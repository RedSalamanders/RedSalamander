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

HRESULT SaveSettingsAndSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept;
HRESULT SaveSettingsAndSchema(std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept;

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
[[nodiscard]] AppTheme ResolveDialogThemeFromSettings(const Common::Settings::Settings& settings) noexcept;

[[nodiscard]] Common::Settings::Settings MergeDiskSettingsWithRuntimeSession(const Common::Settings::Settings& diskSettings,
                                                                             const Common::Settings::Settings& runtimeSettings,
                                                                             std::span<const std::wstring_view> runtimeWindowIds) noexcept;
} // namespace SettingsHotReload

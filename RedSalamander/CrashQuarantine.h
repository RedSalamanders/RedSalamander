#pragma once

namespace Common::Settings
{
struct Settings;
}

namespace CrashQuarantine
{
// Best-effort: if a crash marker exists, offers to disable the last-active filesystem plugin in settings.
// Must be called after settings load and before plugin initialization.
void OfferPluginDisableIfPreviousCrashDetected(Common::Settings::Settings& settings) noexcept;
} // namespace CrashQuarantine


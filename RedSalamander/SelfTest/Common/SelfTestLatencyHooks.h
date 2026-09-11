#pragma once

#include <chrono>
#include <cstdint>
#include <stop_token>
#include <winerror.h>

namespace SelfTestLatency
{
enum class Point : uint8_t
{
    ShellThumbnailProviderAllowed,
    IconExtractSystemIcon,
    IconPathLiveLookup,
    PasteShortcutSave,
    PasteShortcutAfterSlotProbe,
};

void SetNextDelay(Point point, std::chrono::milliseconds delay) noexcept;
void SetNextFailure(Point point, HRESULT hr) noexcept;
void ClearAll() noexcept;
void Consume(Point point, std::stop_token stopToken = {}) noexcept;
HRESULT ConsumeFailure(Point point) noexcept;
uint64_t ConsumeCount(Point point) noexcept;

// Wiring contract:
// - FolderView.Icons.cpp consumes ShellThumbnailProviderAllowed immediately before
//   any provider-allowed shell thumbnail lookup. Cached-only visible lookups must
//   not consume this hook.
// - IconCache.cpp consumes IconExtractSystemIcon around the expensive HICON/shell
//   extraction path used by ExtractSystemIcon.
// - IconCache.cpp consumes IconPathLiveLookup immediately before a live-path
//   SHGetFileInfoW lookup in QuerySysIconIndexForPath. Attribute-only lookups
//   and cached success/failure hits must not consume this hook.
// - FolderView.FileOps.cpp consumes PasteShortcutSave inside the async paste-
//   shortcut worker immediately before each CreateShellShortcut call. Clipboard
//   reading stays on the UI thread and must not consume this hook.
//   Tests may also set a one-shot failure HRESULT for this point; the worker
//   must report that HRESULT as if CreateShellShortcut failed.
// - FolderView.FileOps.cpp consumes PasteShortcutAfterSlotProbe after a Paste
//   Shortcut worker finds a free shortcut filename slot and before it saves the
//   link, so collision selftests can hold one worker in the probe/save window.
}

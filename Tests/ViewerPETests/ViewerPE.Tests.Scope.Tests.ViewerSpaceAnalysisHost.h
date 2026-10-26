#pragma once

#include <cstdint>
#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

// UI Automation's generated RPC declarations require Windows/COM declarations first.
#include <objbase.h>

#include "TestSupport/TestSupport.Tests.DesktopObservation.h"
#include <UIAutomation.h>

inline constexpr int64_t kViewerSpaceFileModifiedTime   = 133830651000000000;
inline constexpr int64_t kViewerSpaceFolderModifiedTime = 133799526000000000;

#pragma warning(push)
// WIL headers only: deleted special members, padding and SAL analysis. Owner: ViewerPETests; review on WIL/MSVC updates.
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#pragma warning(pop)

[[nodiscard]] bool CheckViewerSpaceAnalysisUia(HWND viewerWindow,
                                               std::wstring_view expectedRootPath,
                                               std::wstring_view& failure,
                                               wil::com_ptr<IUIAutomationElement>& retainedGridCell,
                                               uint64_t& cellRequestsBeforeDetach,
                                               RedSalamander::TestSupport::DesktopSnapshotSource desktopSource = {}) noexcept;
[[nodiscard]] bool CheckViewerSpaceAnalysisDetached(HWND viewerWindow,
                                                    IUIAutomationElement* retainedGridCell,
                                                    uint64_t cellRequestsBeforeDetach,
                                                    std::wstring_view& failure) noexcept;
[[nodiscard]] bool CheckViewerSpaceDesktopObservationContracts(std::wstring_view& failure) noexcept;

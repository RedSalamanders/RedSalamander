// header.h : include file for standard system include files,
// or project specific include files
//

#pragma once

#include "targetver.h"
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
// C RunTime Header Files
#include <format>
#include <malloc.h>
#include <memory.h>
#include <stdlib.h>
#include <string>
#include <tchar.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>

// MSVC can emit C4625/C4626/C5026/C5027 for WIL move-only templates at their first instantiation site.
// Force the common instantiations while warnings are disabled.
namespace WilWarningSilenceDetail
{
struct ForceWilTemplateInstantiations_MonitorFramework
{
    wil::unique_handle handle;
};
} // namespace WilWarningSilenceDetail
#pragma warning(pop)

// Add a line to the edit control in the main window
void AddLine(PCWSTR line);

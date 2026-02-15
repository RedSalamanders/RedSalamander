// header.h : include file for standard system include files,
// or project specific include files
//

#pragma once

#include "targetver.h"

// C RunTime Header Files
#include <format>
#include <malloc.h>
#include <memory.h>
#include <stdlib.h>
#include <string>
#include <tchar.h>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>
#pragma warning(pop)

extern PCWSTR REDSALAMANDER_TEXT_VERSION;
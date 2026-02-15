// Common.cpp : Defines the exported functions for the DLL.
//

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "Common.h"

// This is an example of an exported variable
COMMON_API int nCommon = 0;

// This is an example of an exported function.
COMMON_API int fnCommon()
{
    return 0;
}

// This is the constructor of a class that has been exported.
CCommon::CCommon()
{
    return;
}

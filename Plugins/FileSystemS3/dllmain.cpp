#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "FileSystemS3.Internal.h"

HINSTANCE g_hInstance = nullptr;

BOOL WINAPI DllMain(HINSTANCE hinst, DWORD reason, LPVOID /*reserved*/)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = hinst;
        DisableThreadLibraryCalls(hinst);
    }
    return TRUE;
}

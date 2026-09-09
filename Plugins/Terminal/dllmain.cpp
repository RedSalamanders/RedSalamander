#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

HINSTANCE g_hInstance = nullptr;

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID /*reserved*/)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}

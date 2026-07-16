#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

HINSTANCE g_hInstance = nullptr;

BOOL WINAPI DllMain(HINSTANCE hinst, DWORD reason, LPVOID /*reserved*/)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = hinst;
        DisableThreadLibraryCalls(hinst);
    }

    // RedSalamanderPluginShutdown is the required quiet point. COM releases,
    // vector destruction, and staged-file cleanup must never run under the
    // loader lock from DLL_PROCESS_DETACH.
    return TRUE;
}

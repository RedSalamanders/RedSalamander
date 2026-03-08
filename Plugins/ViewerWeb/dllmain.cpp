#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

HINSTANCE g_hInstance = nullptr;

// Defined in ViewerWeb.cpp — releases the shared WebView2 environment.
void ResetSharedEnvironment() noexcept;

BOOL WINAPI DllMain(HINSTANCE hinst, DWORD reason, LPVOID /*reserved*/)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = hinst;
        DisableThreadLibraryCalls(hinst);
    }
    else if (reason == DLL_PROCESS_DETACH)
    {
        ResetSharedEnvironment();
    }
    return TRUE;
}

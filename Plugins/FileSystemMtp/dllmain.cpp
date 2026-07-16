#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

HINSTANCE g_hInstance = nullptr;

BOOL APIENTRY DllMain(HINSTANCE hInstance, DWORD reason, LPVOID reserved)
{
    static_cast<void>(reserved);

    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = hInstance;
        DisableThreadLibraryCalls(hInstance);
    }

    return TRUE;
}

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <new>

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#include "FileSystemDummy.h"

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* /*host*/, void** result)
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IFileSystem))
    {
        auto* instance = new (std::nothrow) FileSystemDummy();
        if (instance == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        HRESULT hr = instance->QueryInterface(riid, result);
        instance->Release();
        return hr;
    }

    return E_NOINTERFACE;
}

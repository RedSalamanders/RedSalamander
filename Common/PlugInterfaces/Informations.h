#pragma once

#include <unknwn.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure
struct PluginMetaData
{
    // Stable plugin identifier (non-localized, long form). Example: "builtin/file-system".
    const wchar_t* id;
    // Short identifier used for navigation prefixes (scheme). Example: "file", "fk".
    const wchar_t* shortId;
    // Localized display name for UI.
    const wchar_t* name;
    // Localized description for "About" UI.
    const wchar_t* description;
    // Optional author/organization (may be nullptr).
    const wchar_t* author;
    // Optional version string (may be nullptr).
    const wchar_t* version;
};
#pragma warning(pop)

// Plugins expose metadata and configuration via this interface.
// Notes:
// - All returned pointers are owned by the plugin object; callers MUST NOT free them.
// - Pointers remain valid until the next call to the same method or until the object is released.
// - JSON strings are UTF-8, NUL-terminated.
interface __declspec(uuid("d6f85c49-3a9c-4e1c-8f3f-6b8cc3b83c62")) __declspec(novtable) IInformations : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept         = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept  = 0;
    virtual HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept  = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL * pSomethingToSave) noexcept             = 0;
};

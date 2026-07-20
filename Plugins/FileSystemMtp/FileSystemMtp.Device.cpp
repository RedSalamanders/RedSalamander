#include "FileSystemMtp.h"

#include <PortableDevice.h>
#include <PortableDeviceApi.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cassert>
#include <condition_variable>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string_view>
#include <unordered_map>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 4514 4623 4625 4626 5026 5027) // WIL result macros: deleted special members and unused inline helpers
#include <wil/result_macros.h>
#pragma warning(pop)

#include "Helpers.h"

namespace FileSystemMtpInternal
{
namespace
{

constexpr DWORD kReadChunkDefault                = 256U * 1024U;
constexpr DWORD kReadChunkMinimum                = 64U * 1024U;
constexpr DWORD kReadChunkMaximum                = 4U * 1024U * 1024U;
constexpr DWORD kWriteChunkDefault               = 256U * 1024U;
constexpr DWORD kWriteChunkMinimum               = 64U * 1024U;
constexpr DWORD kWriteChunkMaximum               = 4U * 1024U * 1024U;
constexpr DWORD kBulkPropertiesCallbackTimeoutMs = 120'000U;
constexpr DWORD kReadAccess                      = GENERIC_READ;
constexpr DWORD kWriteAccess                     = GENERIC_READ | GENERIC_WRITE;

struct ComInitialization
{
    HRESULT hr{E_FAIL};
    bool uninitialize{false};

    explicit ComInitialization(bool initializeIfNeeded = false) noexcept
    {
        ULONG_PTR contextToken = 0u;
        if (SUCCEEDED(CoGetContextToken(&contextToken)))
        {
#ifdef _DEBUG
            if (! initializeIfNeeded)
            {
                APTTYPE apartmentType{};
                APTTYPEQUALIFIER apartmentQualifier{};
                const HRESULT apartmentHr = CoGetApartmentType(&apartmentType, &apartmentQualifier);
                assert(SUCCEEDED(apartmentHr) && apartmentType == APTTYPE_MTA &&
                       "WPD backend calls must execute on the lifetime-MTA command worker.");
            }
#endif
            hr = S_OK;
            return;
        }

#ifdef _DEBUG
        assert(initializeIfNeeded && "WPD backend calls must execute on the lifetime-MTA command worker.");
#endif
        if (! initializeIfNeeded)
        {
            hr = CO_E_NOTINITIALIZED;
            return;
        }

        hr           = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        uninitialize = SUCCEEDED(hr);
    }

    ~ComInitialization()
    {
        if (uninitialize)
        {
            CoUninitialize();
        }
    }

    bool IsUsable() const noexcept
    {
        return SUCCEEDED(hr) || hr == RPC_E_CHANGED_MODE;
    }
};

struct DeviceDescriptor
{
    std::wstring pnpId;
    std::wstring name;
    std::wstring displayName;
};

struct ResolvedObject
{
    bool root{false};
    bool deviceRoot{false};
    bool fromPathCache{false};
    std::wstring pnpId;
    std::wstring objectId;
    MtpItem item;
    wil::com_ptr<IPortableDevice> device;
    wil::com_ptr<IPortableDeviceContent> content;
};

void AbortPortableDeviceWriteStream(IStream* stream) noexcept;

struct WpdCancellationState
{
    WpdCancellationState()  = default;
    ~WpdCancellationState() = default;

    WpdCancellationState(const WpdCancellationState&)            = delete;
    WpdCancellationState(WpdCancellationState&&)                 = delete;
    WpdCancellationState& operator=(const WpdCancellationState&) = delete;
    WpdCancellationState& operator=(WpdCancellationState&&)      = delete;

    std::mutex mutex;
    wil::com_ptr<IPortableDeviceContent> activeContent;
    wil::com_ptr<IStream> activeStream;
};

void RequestWpdCancel(WpdCancellationState& state) noexcept
{
    wil::com_ptr<IPortableDeviceContent> content;
    wil::com_ptr<IStream> stream;
    {
        std::lock_guard lock(state.mutex);
        content = state.activeContent;
        stream  = state.activeStream;
    }

    if (content)
    {
        static_cast<void>(content->Cancel());
    }

    if (stream)
    {
        AbortPortableDeviceWriteStream(stream.get());
    }
}

class ScopedActiveWpdContent
{
public:
    ScopedActiveWpdContent(WpdCancellationState* state, const wil::com_ptr<IPortableDeviceContent>& content) noexcept : _state(state), _content(content.get())
    {
        if (_state != nullptr && content)
        {
            std::lock_guard lock(_state->mutex);
            _state->activeContent = content;
        }
    }

    ~ScopedActiveWpdContent()
    {
        if (_state != nullptr && _content != nullptr)
        {
            std::lock_guard lock(_state->mutex);
            if (_state->activeContent.get() == _content)
            {
                _state->activeContent.reset();
            }
        }
    }

    ScopedActiveWpdContent(const ScopedActiveWpdContent&)            = delete;
    ScopedActiveWpdContent(ScopedActiveWpdContent&&)                 = delete;
    ScopedActiveWpdContent& operator=(const ScopedActiveWpdContent&) = delete;
    ScopedActiveWpdContent& operator=(ScopedActiveWpdContent&&)      = delete;

private:
    WpdCancellationState* _state     = nullptr;
    IPortableDeviceContent* _content = nullptr;
};

class ScopedActiveWpdStream
{
public:
    ScopedActiveWpdStream(WpdCancellationState* state, const wil::com_ptr<IStream>& stream) noexcept : _state(state), _stream(stream.get())
    {
        if (_state != nullptr && stream)
        {
            std::lock_guard lock(_state->mutex);
            _state->activeStream = stream;
        }
    }

    ~ScopedActiveWpdStream()
    {
        if (_state != nullptr && _stream != nullptr)
        {
            std::lock_guard lock(_state->mutex);
            if (_state->activeStream.get() == _stream)
            {
                _state->activeStream.reset();
            }
        }
    }

    ScopedActiveWpdStream(const ScopedActiveWpdStream&)            = delete;
    ScopedActiveWpdStream(ScopedActiveWpdStream&&)                 = delete;
    ScopedActiveWpdStream& operator=(const ScopedActiveWpdStream&) = delete;
    ScopedActiveWpdStream& operator=(ScopedActiveWpdStream&&)      = delete;

private:
    WpdCancellationState* _state = nullptr;
    IStream* _stream             = nullptr;
};

std::wstring DeviceDisplayName(std::wstring_view friendlyName, std::wstring_view pnpId)
{
    const auto baseName = friendlyName.empty() ? L"MTP Device" : std::wstring(friendlyName);
    return std::format(L"{} {}", baseName, MtpDeviceIdentitySuffix(pnpId));
}

std::wstring CaseFoldKey(std::wstring_view value)
{
    return OrdinalString::FoldCaseInvariant(value);
}

bool DesiredAccessIsCovered(DWORD cachedAccess, DWORD desiredAccess) noexcept
{
    return (cachedAccess & desiredAccess) == desiredAccess;
}

std::optional<std::wstring> ExtractDeviceHashToken(std::wstring_view value)
{
    constexpr std::wstring_view kPrefix = L"[devid:";
    constexpr std::wstring_view kSuffix = L"]";

    const size_t start = value.rfind(kPrefix);
    if (start == std::wstring_view::npos)
    {
        return std::nullopt;
    }

    const size_t tokenStart = start + kPrefix.size();
    const size_t tokenEnd   = value.find(kSuffix, tokenStart);
    if (tokenEnd == std::wstring_view::npos || tokenEnd == tokenStart)
    {
        return std::nullopt;
    }

    std::wstring token(value.substr(tokenStart, tokenEnd - tokenStart));
    for (auto& ch : token)
    {
        ch = static_cast<wchar_t>(::towupper(ch));
    }
    return token;
}

HRESULT CreatePortableDeviceValues(wil::com_ptr<IPortableDeviceValues>& values)
{
    values.reset();
    return CoCreateInstance(CLSID_PortableDeviceValues, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(values.put()));
}

HRESULT CreateClientInfo(DWORD desiredAccess, IPortableDeviceValues** values)
{
    if (values == nullptr)
    {
        return E_POINTER;
    }

    wil::com_ptr<IPortableDeviceValues> clientInfo;
    RETURN_IF_FAILED(CreatePortableDeviceValues(clientInfo));

    RETURN_IF_FAILED(clientInfo->SetStringValue(WPD_CLIENT_NAME, L"RedSalamander"));
    RETURN_IF_FAILED(clientInfo->SetUnsignedIntegerValue(WPD_CLIENT_MAJOR_VERSION, 1));
    RETURN_IF_FAILED(clientInfo->SetUnsignedIntegerValue(WPD_CLIENT_MINOR_VERSION, 0));
    RETURN_IF_FAILED(clientInfo->SetUnsignedIntegerValue(WPD_CLIENT_REVISION, 0));
    RETURN_IF_FAILED(clientInfo->SetUnsignedIntegerValue(WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, SECURITY_IMPERSONATION));
    RETURN_IF_FAILED(clientInfo->SetUnsignedIntegerValue(WPD_CLIENT_DESIRED_ACCESS, desiredAccess));

    *values = clientInfo.detach();
    return S_OK;
}

HRESULT EnumerateDevices(std::vector<DeviceDescriptor>& devices)
{
    devices.clear();

    wil::com_ptr<IPortableDeviceManager> manager;
    HRESULT hr = CoCreateInstance(CLSID_PortableDeviceManager, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(manager.put()));
    if (FAILED(hr))
    {
        return hr;
    }

    DWORD count = 0;
    hr          = manager->GetDevices(nullptr, &count);
    if (FAILED(hr))
    {
        return hr;
    }

    if (count == 0)
    {
        return S_OK;
    }

    std::vector<PWSTR> rawIds(count, nullptr);
    hr = manager->GetDevices(rawIds.data(), &count);
    if (FAILED(hr))
    {
        return hr;
    }

    for (DWORD index = 0; index < count; ++index)
    {
        const auto rawId = rawIds[index];
        if (rawId == nullptr)
        {
            continue;
        }

        auto freeId = wil::scope_exit([rawId]() noexcept { CoTaskMemFree(rawId); });

        std::wstring pnpId(rawId);
        DWORD nameChars = 0;
        std::wstring friendlyName;

        HRESULT nameHr = manager->GetDeviceFriendlyName(rawId, nullptr, &nameChars);
        if (SUCCEEDED(nameHr) && nameChars > 1)
        {
            std::wstring buffer(nameChars, L'\0');
            nameHr = manager->GetDeviceFriendlyName(rawId, buffer.data(), &nameChars);
            if (SUCCEEDED(nameHr))
            {
                if (! buffer.empty() && buffer.back() == L'\0')
                {
                    buffer.pop_back();
                }
                friendlyName = std::move(buffer);
            }
        }

        if (friendlyName.empty())
        {
            friendlyName = L"MTP Device";
        }

        devices.push_back(DeviceDescriptor{
            .pnpId       = std::move(pnpId),
            .name        = friendlyName,
            .displayName = DeviceDisplayName(friendlyName, rawId),
        });
    }

    return S_OK;
}

MtpItem MakeRootDeviceItem(const DeviceDescriptor& device)
{
    return MtpItem{
        .name           = device.displayName,
        .attributes     = FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY,
        .sizeBytes      = 0,
        .creationTime   = 0,
        .lastAccessTime = 0,
        .lastWriteTime  = 0,
        .changeTime     = 0,
        .persistentId   = device.pnpId,
        .objectId       = device.pnpId,
        .streamable     = false,
    };
}

HRESULT OpenDeviceSession(std::wstring_view pnpId, DWORD desiredAccess, wil::com_ptr<IPortableDevice>& device, wil::com_ptr<IPortableDeviceContent>& content)
{
    device.reset();
    content.reset();

    wil::com_ptr<IPortableDevice> openedDevice;
    HRESULT hr = CoCreateInstance(CLSID_PortableDeviceFTM, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(openedDevice.put()));
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IPortableDeviceValues> clientInfo;
    RETURN_IF_FAILED(CreateClientInfo(desiredAccess, clientInfo.put()));

    std::wstring pnpIdCopy(pnpId);
    RETURN_IF_FAILED(openedDevice->Open(pnpIdCopy.c_str(), clientInfo.get()));

    wil::com_ptr<IPortableDeviceContent> openedContent;
    RETURN_IF_FAILED(openedDevice->Content(openedContent.put()));

    device  = std::move(openedDevice);
    content = std::move(openedContent);
    return S_OK;
}

std::optional<std::wstring> GetStringProperty(const wil::com_ptr<IPortableDeviceValues>& values, REFPROPERTYKEY key)
{
    PWSTR rawValue = nullptr;
    HRESULT hr     = values->GetStringValue(key, &rawValue);
    if (FAILED(hr) || rawValue == nullptr)
    {
        return std::nullopt;
    }

    auto freeValue = wil::scope_exit([rawValue]() noexcept { CoTaskMemFree(rawValue); });
    return std::wstring(rawValue);
}

bool GetBoolProperty(const wil::com_ptr<IPortableDeviceValues>& values, REFPROPERTYKEY key)
{
    BOOL value = FALSE;
    return SUCCEEDED(values->GetBoolValue(key, &value)) && value != FALSE;
}

ULONGLONG GetUnsignedLargeIntegerProperty(const wil::com_ptr<IPortableDeviceValues>& values, REFPROPERTYKEY key, ULONGLONG fallback = 0)
{
    ULONGLONG value = fallback;
    if (FAILED(values->GetUnsignedLargeIntegerValue(key, &value)))
    {
        return fallback;
    }

    return value;
}

bool TryGetGuidProperty(const wil::com_ptr<IPortableDeviceValues>& values, REFPROPERTYKEY key, GUID& value)
{
    return SUCCEEDED(values->GetGuidValue(key, &value));
}

bool IsFolderContentType(REFGUID contentType) noexcept
{
    return IsEqualGUID(contentType, WPD_CONTENT_TYPE_FOLDER) || IsEqualGUID(contentType, WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT);
}

HRESULT ValuesToObjectItem(std::wstring_view objectId, const wil::com_ptr<IPortableDeviceValues>& values, MtpItem& item)
{
    std::wstring name =
        GetStringProperty(values, WPD_OBJECT_ORIGINAL_FILE_NAME).value_or(GetStringProperty(values, WPD_OBJECT_NAME).value_or(std::wstring(objectId)));
    name = SanitizeMtpPathComponent(std::move(name));

    GUID contentType          = GUID_NULL;
    const bool hasContentType = TryGetGuidProperty(values, WPD_OBJECT_CONTENT_TYPE, contentType);
    const bool folder         = hasContentType && IsFolderContentType(contentType);

    DWORD attributes = FILE_ATTRIBUTE_READONLY;
    if (folder)
    {
        attributes |= FILE_ATTRIBUTE_DIRECTORY;
    }
    if (GetBoolProperty(values, WPD_OBJECT_ISHIDDEN))
    {
        attributes |= FILE_ATTRIBUTE_HIDDEN;
    }
    if (GetBoolProperty(values, WPD_OBJECT_ISSYSTEM))
    {
        attributes |= FILE_ATTRIBUTE_SYSTEM;
    }

    const auto persistentId    = GetStringProperty(values, WPD_OBJECT_PERSISTENT_UNIQUE_ID).value_or(L"");
    const ULONGLONG objectSize = GetUnsignedLargeIntegerProperty(values, WPD_OBJECT_SIZE);

    item = MtpItem{
        .name           = std::move(name),
        .attributes     = attributes,
        .sizeBytes      = objectSize,
        .creationTime   = 0,
        .lastAccessTime = 0,
        .lastWriteTime  = 0,
        .changeTime     = 0,
        .persistentId   = persistentId,
        .objectId       = std::wstring(objectId),
        .streamable     = ! folder,
    };

    return S_OK;
}

HRESULT GetObjectItem(const wil::com_ptr<IPortableDeviceContent>& content, std::wstring_view objectId, MtpItem& item)
{
    wil::com_ptr<IPortableDeviceProperties> properties;
    RETURN_IF_FAILED(content->Properties(properties.put()));

    std::wstring objectIdCopy(objectId);
    wil::com_ptr<IPortableDeviceValues> values;
    RETURN_IF_FAILED(properties->GetValues(objectIdCopy.c_str(), nullptr, values.put()));

    return ValuesToObjectItem(objectId, values, item);
}

HRESULT CreateObjectIdCollection(const std::vector<std::wstring>& objectIds, wil::com_ptr<IPortableDevicePropVariantCollection>& collection)
{
    collection.reset();
    RETURN_IF_FAILED(CoCreateInstance(CLSID_PortableDevicePropVariantCollection, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(collection.put())));
    RETURN_IF_FAILED(collection->ChangeType(VT_LPWSTR));

    for (const std::wstring& objectId : objectIds)
    {
        PROPVARIANT value{};
        value.vt      = VT_LPWSTR;
        value.pwszVal = const_cast<PWSTR>(objectId.c_str());
        RETURN_IF_FAILED(collection->Add(&value));
    }

    return S_OK;
}

HRESULT CreateObjectIdCollection(std::wstring_view objectId, wil::com_ptr<IPortableDevicePropVariantCollection>& collection)
{
    return CreateObjectIdCollection(std::vector<std::wstring>{std::wstring(objectId)}, collection);
}

HRESULT FirstFailureFromResultCollection(const wil::com_ptr<IPortableDevicePropVariantCollection>& results)
{
    if (! results)
    {
        return S_OK;
    }

    DWORD count = 0;
    RETURN_IF_FAILED(results->GetCount(&count));
    for (DWORD index = 0; index < count; ++index)
    {
        PROPVARIANT value{};
        auto clearValue = wil::scope_exit([&value]() noexcept { static_cast<void>(PropVariantClear(&value)); });
        RETURN_IF_FAILED(results->GetAt(index, &value));
        if (value.vt == VT_ERROR && FAILED(value.scode))
        {
            return value.scode;
        }
    }

    return S_OK;
}

HRESULT FirstFailureFromValues(const wil::com_ptr<IPortableDeviceValues>& results)
{
    if (! results)
    {
        return S_OK;
    }

    DWORD count = 0;
    RETURN_IF_FAILED(results->GetCount(&count));
    for (DWORD index = 0; index < count; ++index)
    {
        PROPERTYKEY key{};
        PROPVARIANT value{};
        auto clearValue = wil::scope_exit([&value]() noexcept { static_cast<void>(PropVariantClear(&value)); });
        RETURN_IF_FAILED(results->GetAt(index, &key, &value));
        if (value.vt == VT_ERROR && FAILED(value.scode))
        {
            return value.scode;
        }
    }

    return S_OK;
}

HRESULT CreateObjectPropertyKeyCollection(wil::com_ptr<IPortableDeviceKeyCollection>& keys)
{
    keys.reset();
    RETURN_IF_FAILED(CoCreateInstance(CLSID_PortableDeviceKeyCollection, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(keys.put())));

    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_ID));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_ORIGINAL_FILE_NAME));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_NAME));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_CONTENT_TYPE));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_ISHIDDEN));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_ISSYSTEM));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_PERSISTENT_UNIQUE_ID));
    RETURN_IF_FAILED(keys->Add(WPD_OBJECT_SIZE));
    return S_OK;
}

class BulkPropertiesCallback final : public IPortableDevicePropertiesBulkCallback
{
public:
    BulkPropertiesCallback() = default;

    BulkPropertiesCallback(const BulkPropertiesCallback&)            = delete;
    BulkPropertiesCallback(BulkPropertiesCallback&&)                 = delete;
    BulkPropertiesCallback& operator=(const BulkPropertiesCallback&) = delete;
    BulkPropertiesCallback& operator=(BulkPropertiesCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IPortableDevicePropertiesBulkCallback))
        {
            *ppvObject = static_cast<IPortableDevicePropertiesBulkCallback*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE OnStart(REFGUID pContext) noexcept override
    {
        static_cast<void>(pContext);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnProgress(REFGUID pContext, IPortableDeviceValuesCollection* pResults) noexcept override
    {
        static_cast<void>(pContext);
        if (! pResults)
        {
            return S_OK;
        }

        DWORD count = 0;
        HRESULT hr  = pResults->GetCount(&count);
        if (FAILED(hr))
        {
            Complete(hr);
            return hr;
        }

        for (DWORD index = 0; index < count; ++index)
        {
            wil::com_ptr<IPortableDeviceValues> values;
            hr = pResults->GetAt(index, values.put());
            if (FAILED(hr))
            {
                Complete(hr);
                return hr;
            }

            const auto objectId = GetStringProperty(values, WPD_OBJECT_ID);
            if (! objectId || objectId->empty())
            {
                continue;
            }

            {
                std::lock_guard lock(_mutex);
                _valuesByObjectId[*objectId] = std::move(values);
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnEnd(REFGUID pContext, HRESULT hrStatus) noexcept override
    {
        static_cast<void>(pContext);
        Complete(hrStatus);
        return S_OK;
    }

    HRESULT WaitAndTake(std::unordered_map<std::wstring, wil::com_ptr<IPortableDeviceValues>>& valuesByObjectId, DWORD timeoutMs) noexcept
    {
        std::unique_lock lock(_mutex);
        if (! _cv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [this]() noexcept { return _done; }))
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        if (FAILED(_status))
        {
            return _status;
        }

        valuesByObjectId = std::move(_valuesByObjectId);
        return S_OK;
    }

private:
    void Complete(HRESULT hr) noexcept
    {
        {
            std::lock_guard lock(_mutex);
            if (FAILED(hr) && SUCCEEDED(_status))
            {
                _status = hr;
            }
            _done = true;
        }
        _cv.notify_all();
    }

    std::atomic_ulong _refCount{1};
    std::mutex _mutex;
    std::condition_variable _cv;
    bool _done      = false;
    HRESULT _status = S_OK;
    std::unordered_map<std::wstring, wil::com_ptr<IPortableDeviceValues>> _valuesByObjectId;
};

HRESULT TryEnumerateObjectItemsBulk(const wil::com_ptr<IPortableDeviceContent>& content,
                                    const std::vector<std::wstring>& objectIds,
                                    std::vector<MtpItem>& items)
{
    if (objectIds.empty())
    {
        items.clear();
        return S_OK;
    }

    wil::com_ptr<IPortableDeviceProperties> properties;
    RETURN_IF_FAILED(content->Properties(properties.put()));

    wil::com_ptr<IPortableDevicePropertiesBulk> bulkProperties;
    const HRESULT bulkQiHr = properties->QueryInterface(IID_PPV_ARGS(bulkProperties.put()));
    if (FAILED(bulkQiHr) || ! bulkProperties)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IPortableDevicePropVariantCollection> objectIdCollection;
    RETURN_IF_FAILED(CreateObjectIdCollection(objectIds, objectIdCollection));

    wil::com_ptr<IPortableDeviceKeyCollection> keys;
    RETURN_IF_FAILED(CreateObjectPropertyKeyCollection(keys));

    auto* callbackRaw = new (std::nothrow) BulkPropertiesCallback();
    if (! callbackRaw)
    {
        return E_OUTOFMEMORY;
    }
    wil::com_ptr<IPortableDevicePropertiesBulkCallback> callback;
    callback.attach(callbackRaw);

    GUID context = GUID_NULL;
    RETURN_IF_FAILED(bulkProperties->QueueGetValuesByObjectList(objectIdCollection.get(), keys.get(), callback.get(), &context));
    RETURN_IF_FAILED(bulkProperties->Start(context));

    std::unordered_map<std::wstring, wil::com_ptr<IPortableDeviceValues>> valuesByObjectId;
    const HRESULT waitHr = callbackRaw->WaitAndTake(valuesByObjectId, kBulkPropertiesCallbackTimeoutMs);
    if (FAILED(waitHr))
    {
        static_cast<void>(bulkProperties->Cancel(context));
        return waitHr;
    }
    if (valuesByObjectId.size() != objectIds.size())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    items.clear();
    items.reserve(objectIds.size());
    for (const std::wstring& objectId : objectIds)
    {
        const auto it = valuesByObjectId.find(objectId);
        if (it == valuesByObjectId.end() || ! it->second)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        MtpItem item;
        RETURN_IF_FAILED(ValuesToObjectItem(objectId, it->second, item));
        items.push_back(std::move(item));
    }

    return S_OK;
}

void DisambiguateDuplicateNames(std::vector<MtpItem>& items)
{
    std::unordered_map<std::wstring, int> counts;
    counts.reserve(items.size());

    for (const auto& item : items)
    {
        std::wstring key = item.name;
        std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) { return static_cast<wchar_t>(::towlower(ch)); });
        ++counts[key];
    }

    uint64_t duplicateGroups = 0;
    uint64_t suffixedEntries = 0;
    for (const auto& [key, count] : counts)
    {
        static_cast<void>(key);
        if (count > 1)
        {
            ++duplicateGroups;
            suffixedEntries += static_cast<uint64_t>(count);
        }
    }

    for (auto& item : items)
    {
        std::wstring key = item.name;
        std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) { return static_cast<wchar_t>(::towlower(ch)); });
        if (counts[key] > 1)
        {
            item.name += MtpDuplicateObjectSuffix(item);
        }
    }

    Debug::Perf::EmitValue(L"mtp.path.duplicate_groups", duplicateGroups, S_OK);
    Debug::Perf::EmitValue(L"mtp.path.suffixed_entries", suffixedEntries, S_OK);
}

HRESULT EnumerateObjectIds(const wil::com_ptr<IPortableDeviceContent>& content, std::wstring_view parentObjectId, std::vector<std::wstring>& objectIds)
{
    objectIds.clear();

    wil::com_ptr<IEnumPortableDeviceObjectIDs> enumerator;
    std::wstring parentCopy(parentObjectId);
    RETURN_IF_FAILED(content->EnumObjects(0, parentCopy.c_str(), nullptr, enumerator.put()));

    for (;;)
    {
        std::array<PWSTR, 16> rawIds{};
        DWORD fetched = 0;
        HRESULT hr    = enumerator->Next(static_cast<DWORD>(rawIds.size()), rawIds.data(), &fetched);
        if (FAILED(hr))
        {
            return hr;
        }

        for (DWORD index = 0; index < fetched; ++index)
        {
            wil::unique_cotaskmem_string rawId(rawIds[index]);
            rawIds[index] = nullptr;
            if (rawId)
            {
                objectIds.emplace_back(rawId.get());
            }
        }

        if (hr != S_OK || fetched == 0)
        {
            break;
        }
    }

    return S_OK;
}

HRESULT EnumerateObjectItems(const wil::com_ptr<IPortableDeviceContent>& content, std::wstring_view parentObjectId, std::vector<MtpItem>& items)
{
    items.clear();

    std::vector<std::wstring> objectIds;
    RETURN_IF_FAILED(EnumerateObjectIds(content, parentObjectId, objectIds));

    if (! objectIds.empty())
    {
        const HRESULT bulkHr = TryEnumerateObjectItemsBulk(content, objectIds, items);
        if (SUCCEEDED(bulkHr))
        {
            Debug::Perf::EmitValue(L"mtp.props.bulk_batches", 1u, S_OK);
            Debug::Perf::EmitValue(L"mtp.props.per_item_calls", 0u, S_OK);
            DisambiguateDuplicateNames(items);
            return S_OK;
        }
    }

    Debug::Perf::EmitValue(L"mtp.props.bulk_batches", 0u, S_OK);
    Debug::Perf::EmitValue(L"mtp.props.per_item_calls", static_cast<uint64_t>(objectIds.size()), S_OK);
    items.reserve(objectIds.size());
    for (const auto& objectId : objectIds)
    {
        MtpItem item;
        HRESULT hr = GetObjectItem(content, objectId, item);
        if (SUCCEEDED(hr))
        {
            items.push_back(std::move(item));
        }
    }

    DisambiguateDuplicateNames(items);
    return S_OK;
}

struct ResolvedDestination
{
    std::wstring normalizedPath;
    std::wstring parentPath;
    std::wstring leafName;
    ResolvedObject parent;
};

bool IsMissingObjectHr(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

bool SameNormalizedParentPath(std::wstring_view sourcePath, const ResolvedDestination& destination)
{
    return NormalizeMtpPath(ParentPath(sourcePath)) == destination.parentPath;
}

HRESULT ReadPortableDeviceStream(const wil::com_ptr<IPortableDeviceContent>& content,
                                 std::wstring_view objectId,
                                 uint64_t expectedSizeBytes,
                                 std::vector<std::byte>& bytes)
{
    bytes.clear();

    wil::com_ptr<IPortableDeviceResources> resources;
    RETURN_IF_FAILED(content->Transfer(resources.put()));

    wil::com_ptr<IStream> stream;
    DWORD optimalBufferSize = 0;
    std::wstring objectIdCopy(objectId);
    RETURN_IF_FAILED(resources->GetStream(objectIdCopy.c_str(), WPD_RESOURCE_DEFAULT, STGM_READ, &optimalBufferSize, stream.put()));

    const DWORD chunkSize = std::clamp(optimalBufferSize == 0 ? kReadChunkDefault : optimalBufferSize, kReadChunkMinimum, kReadChunkMaximum);
    std::vector<std::byte> buffer(chunkSize);

    for (;;)
    {
        ULONG bytesRead = 0;
        HRESULT hr      = stream->Read(buffer.data(), static_cast<ULONG>(buffer.size()), &bytesRead);
        if (FAILED(hr))
        {
            return hr;
        }

        if (bytesRead == 0)
        {
            break;
        }

        bytes.insert(bytes.end(), buffer.begin(), buffer.begin() + bytesRead);
    }

    if (bytes.size() != expectedSizeBytes)
    {
        Debug::Perf::EmitValue(L"mtp.stream.read_size_mismatch", static_cast<uint64_t>(bytes.size()), HRESULT_FROM_WIN32(ERROR_CRC));
        return HRESULT_FROM_WIN32(ERROR_CRC);
    }

    return S_OK;
}

class WpdStreamBackendFileReader final : public IMtpBackendFileReader
{
public:
    WpdStreamBackendFileReader(wil::com_ptr<IPortableDeviceContent> content,
                               wil::com_ptr<IStream> stream,
                               uint64_t sizeBytes,
                               std::shared_ptr<WpdCancellationState> cancelState) noexcept
        : _content(std::move(content)),
          _stream(std::move(stream)),
          _sizeBytes(sizeBytes),
          _cancelState(std::move(cancelState))
    {
    }

    WpdStreamBackendFileReader(const WpdStreamBackendFileReader&)            = delete;
    WpdStreamBackendFileReader(WpdStreamBackendFileReader&&)                 = delete;
    WpdStreamBackendFileReader& operator=(const WpdStreamBackendFileReader&) = delete;
    WpdStreamBackendFileReader& operator=(WpdStreamBackendFileReader&&)      = delete;

    HRESULT GetSize(uint64_t& sizeBytes) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }
        sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT Seek(__int64 offset, unsigned long origin, uint64_t& newPosition) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }
        if (! _stream)
        {
            return E_FAIL;
        }

        DWORD streamOrigin = STREAM_SEEK_SET;
        if (origin == FILE_BEGIN)
        {
            streamOrigin = STREAM_SEEK_SET;
        }
        else if (origin == FILE_CURRENT)
        {
            streamOrigin = STREAM_SEEK_CUR;
        }
        else if (origin == FILE_END)
        {
            streamOrigin = STREAM_SEEK_END;
        }
        else
        {
            return E_INVALIDARG;
        }

        LARGE_INTEGER move{};
        move.QuadPart = offset;
        ULARGE_INTEGER position{};
        ScopedActiveWpdContent activeContent(_cancelState.get(), _content);
        ScopedActiveWpdStream activeStream(_cancelState.get(), _stream);
        const HRESULT seekHr = _stream->Seek(move, streamOrigin, &position);
        if (FAILED(seekHr))
        {
            return seekHr;
        }

        _position  = position.QuadPart;
        newPosition = _position;
        return S_OK;
    }

    HRESULT Read(std::span<std::byte> buffer, unsigned long requestedBytes, unsigned long& bytesRead) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }
        bytesRead = 0;
        if (requestedBytes == 0u)
        {
            return S_OK;
        }
        if (buffer.empty())
        {
            return E_POINTER;
        }
        if (! _stream)
        {
            return E_FAIL;
        }

        const ULONG request = static_cast<ULONG>(std::min<size_t>(buffer.size(), requestedBytes));
        ScopedActiveWpdContent activeContent(_cancelState.get(), _content);
        ScopedActiveWpdStream activeStream(_cancelState.get(), _stream);
        const HRESULT readHr = _stream->Read(buffer.data(), request, &bytesRead);
        if (FAILED(readHr))
        {
            return readHr;
        }

        _position += bytesRead;
        if (bytesRead != 0u)
        {
            Debug::Perf::EmitValue(L"mtp.transfer.read_bytes", static_cast<uint64_t>(bytesRead), S_OK);
        }
        return S_OK;
    }

private:
    wil::com_ptr<IPortableDeviceContent> _content;
    wil::com_ptr<IStream> _stream;
    uint64_t _sizeBytes = 0;
    uint64_t _position  = 0;
    std::shared_ptr<WpdCancellationState> _cancelState;
};

struct SelfTestWpdOptions
{
    uint32_t readFileDelayMs = 0;
    bool sessionDeathOnce = false;
    bool changeFileSizeAfterFirstLookup = false;
};

[[nodiscard]] DeviceDescriptor SelfTestDeviceDescriptor()
{
    return DeviceDescriptor{
        .pnpId       = L"selftest-wpd-pnp-1",
        .name        = L"Fake Phone",
        .displayName = L"Fake Phone [devid:000000000000F00D]",
    };
}

[[nodiscard]] MtpItem SelfTestItem(std::wstring name,
                                   std::wstring objectId,
                                   unsigned long attributes,
                                   uint64_t sizeBytes = 0,
                                   std::wstring persistentId = {})
{
    if (persistentId.empty())
    {
        persistentId = objectId + L"-puid";
    }
    return MtpItem{
        .name           = std::move(name),
        .attributes     = attributes,
        .sizeBytes      = sizeBytes,
        .creationTime   = 0,
        .lastAccessTime = 0,
        .lastWriteTime  = 0,
        .changeTime     = 0,
        .persistentId   = std::move(persistentId),
        .objectId       = std::move(objectId),
        .streamable     = (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
    };
}

class WpdDeviceOperations
{
public:
    virtual ~WpdDeviceOperations() = default;
    virtual MtpBackendInfo GetInfo() const noexcept = 0;
    virtual HRESULT EnumerateDevices(std::vector<DeviceDescriptor>& devices) noexcept = 0;
    virtual HRESULT OpenSession(std::wstring_view pnpId,
                                DWORD desiredAccess,
                                wil::com_ptr<IPortableDevice>& device,
                                wil::com_ptr<IPortableDeviceContent>& content) noexcept = 0;
    virtual HRESULT EnumerateItems(const wil::com_ptr<IPortableDeviceContent>& content,
                                   std::wstring_view parentObjectId,
                                   std::vector<MtpItem>& items) noexcept = 0;
    virtual HRESULT CreateReader(const ResolvedObject& resolved,
                                 const std::shared_ptr<WpdCancellationState>& cancelState,
                                 std::shared_ptr<IMtpBackendFileReader>& reader) noexcept = 0;
};

class LiveWpdDeviceOperations final : public WpdDeviceOperations
{
public:
    MtpBackendInfo GetInfo() const noexcept override
    {
        return {.readOnly = true, .supportsWrite = true, .liveWpd = true};
    }

    HRESULT EnumerateDevices(std::vector<DeviceDescriptor>& devices) noexcept override
    {
        return FileSystemMtpInternal::EnumerateDevices(devices);
    }

    HRESULT OpenSession(std::wstring_view pnpId,
                        DWORD desiredAccess,
                        wil::com_ptr<IPortableDevice>& device,
                        wil::com_ptr<IPortableDeviceContent>& content) noexcept override
    {
        return OpenDeviceSession(pnpId, desiredAccess, device, content);
    }

    HRESULT EnumerateItems(const wil::com_ptr<IPortableDeviceContent>& content,
                           std::wstring_view parentObjectId,
                           std::vector<MtpItem>& items) noexcept override
    {
        return EnumerateObjectItems(content, parentObjectId, items);
    }

    HRESULT CreateReader(const ResolvedObject& resolved,
                         const std::shared_ptr<WpdCancellationState>& cancelState,
                         std::shared_ptr<IMtpBackendFileReader>& reader) noexcept override
    {
        wil::com_ptr<IPortableDeviceResources> resources;
        RETURN_IF_FAILED(resolved.content->Transfer(resources.put()));

        wil::com_ptr<IStream> stream;
        DWORD optimalBufferSize = 0;
        const std::wstring objectId(resolved.objectId);
        ScopedActiveWpdContent activeContent(cancelState.get(), resolved.content);
        RETURN_IF_FAILED(resources->GetStream(objectId.c_str(), WPD_RESOURCE_DEFAULT, STGM_READ, &optimalBufferSize, stream.put()));
        static_cast<void>(optimalBufferSize);

        LARGE_INTEGER zero{};
        ULARGE_INTEGER position{};
        if (SUCCEEDED(stream->Seek(zero, STREAM_SEEK_CUR, &position)))
        {
            reader = std::make_shared<WpdStreamBackendFileReader>(resolved.content, std::move(stream), resolved.item.sizeBytes, cancelState);
            return S_OK;
        }

        std::vector<std::byte> bytes;
        RETURN_IF_FAILED(ReadPortableDeviceStream(resolved.content, resolved.objectId, resolved.item.sizeBytes, bytes));
        reader = CreateMemoryBackendFileReader(std::move(bytes));
        return S_OK;
    }
};

class SelfTestWpdDeviceOperations final : public WpdDeviceOperations
{
public:
    explicit SelfTestWpdDeviceOperations(SelfTestWpdOptions options) noexcept : _options(options)
    {
    }

    MtpBackendInfo GetInfo() const noexcept override
    {
        return {.readOnly = true, .supportsWrite = false, .liveWpd = false};
    }

    HRESULT EnumerateDevices(std::vector<DeviceDescriptor>& devices) noexcept override
    {
        devices = {SelfTestDeviceDescriptor()};
        return S_OK;
    }

    HRESULT OpenSession(std::wstring_view,
                        DWORD,
                        wil::com_ptr<IPortableDevice>& device,
                        wil::com_ptr<IPortableDeviceContent>& content) noexcept override
    {
        device.reset();
        content.reset();
        return S_OK;
    }

    HRESULT EnumerateItems(const wil::com_ptr<IPortableDeviceContent>&,
                           std::wstring_view parentObjectId,
                           std::vector<MtpItem>& items) noexcept override
    {
        items.clear();
        if (_options.sessionDeathOnce && parentObjectId == L"selftest-camera" && ! _sessionDeathReturned)
        {
            _sessionDeathReturned = true;
            return RPC_E_DISCONNECTED;
        }
        if (parentObjectId == WPD_DEVICE_OBJECT_ID)
        {
            items.push_back(SelfTestItem(L"Internal Storage", L"selftest-storage", FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY));
        }
        else if (parentObjectId == L"selftest-storage")
        {
            items.push_back(SelfTestItem(L"DCIM", L"selftest-dcim", FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY));
        }
        else if (parentObjectId == L"selftest-dcim")
        {
            items.push_back(SelfTestItem(L"Camera", L"selftest-camera", FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY));
        }
        else if (parentObjectId == L"selftest-camera")
        {
            constexpr std::string_view original = "RedSalamander deterministic MTP fixture\r\n";
            constexpr std::string_view changed  = "RedSalamander refreshed MTP fixture payload\r\n";
            const bool changedSize = _options.changeFileSizeAfterFirstLookup && _photoLookupCount++ != 0u;
            items.push_back(SelfTestItem(L"photo001.txt",
                                         L"selftest-photo-001",
                                         FILE_ATTRIBUTE_READONLY,
                                         changedSize ? changed.size() : original.size(),
                                         L"selftest-photo-001-puid"));
        }
        else
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }
        return S_OK;
    }

    HRESULT CreateReader(const ResolvedObject& resolved,
                         const std::shared_ptr<WpdCancellationState>&,
                         std::shared_ptr<IMtpBackendFileReader>& reader) noexcept override
    {
        constexpr std::string_view original = "RedSalamander deterministic MTP fixture\r\n";
        constexpr std::string_view changed  = "RedSalamander refreshed MTP fixture payload\r\n";
        const std::string_view payload = resolved.item.sizeBytes == changed.size() ? changed : original;
        std::vector<std::byte> bytes(payload.size());
        std::memcpy(bytes.data(), payload.data(), payload.size());
        reader = CreateMemoryBackendFileReader(std::move(bytes), _options.readFileDelayMs);
        return S_OK;
    }

private:
    SelfTestWpdOptions _options;
    bool _sessionDeathReturned = false;
    uint32_t _photoLookupCount = 0;
};

HRESULT CreateObjectValues(std::wstring_view parentObjectId,
                           std::wstring_view leafName,
                           REFGUID contentType,
                           REFGUID objectFormat,
                           std::optional<uint64_t> sizeBytes,
                           bool setOriginalFileName,
                           wil::com_ptr<IPortableDeviceValues>& values)
{
    RETURN_IF_FAILED(CreatePortableDeviceValues(values));

    const std::wstring parentCopy(parentObjectId);
    const std::wstring leafCopy(leafName);
    RETURN_IF_FAILED(values->SetStringValue(WPD_OBJECT_PARENT_ID, parentCopy.c_str()));
    RETURN_IF_FAILED(values->SetStringValue(WPD_OBJECT_NAME, leafCopy.c_str()));
    if (setOriginalFileName)
    {
        RETURN_IF_FAILED(values->SetStringValue(WPD_OBJECT_ORIGINAL_FILE_NAME, leafCopy.c_str()));
    }
    RETURN_IF_FAILED(values->SetGuidValue(WPD_OBJECT_CONTENT_TYPE, contentType));
    RETURN_IF_FAILED(values->SetGuidValue(WPD_OBJECT_FORMAT, objectFormat));
    if (sizeBytes.has_value())
    {
        RETURN_IF_FAILED(values->SetUnsignedLargeIntegerValue(WPD_OBJECT_SIZE, static_cast<ULONGLONG>(sizeBytes.value())));
    }

    return S_OK;
}

void AbortPortableDeviceWriteStream(IStream* stream) noexcept
{
    if (! stream)
    {
        return;
    }

    wil::com_ptr<IPortableDeviceDataStream> dataStream;
    if (SUCCEEDED(stream->QueryInterface(IID_PPV_ARGS(dataStream.put()))) && dataStream)
    {
        static_cast<void>(dataStream->Cancel());
    }
    static_cast<void>(stream->Revert());
}

HRESULT WritePortableDeviceStream(IStream* stream, DWORD optimalBufferSize, std::span<const std::byte> bytes)
{
    if (! stream)
    {
        return E_POINTER;
    }

    const DWORD chunkSize = std::clamp(optimalBufferSize == 0 ? kWriteChunkDefault : optimalBufferSize, kWriteChunkMinimum, kWriteChunkMaximum);

    size_t offset = 0;
    while (offset < bytes.size())
    {
        const size_t remaining = bytes.size() - offset;
        const auto chunk       = static_cast<ULONG>(std::min<size_t>(remaining, chunkSize));
        ULONG written          = 0;
        const HRESULT hr       = stream->Write(bytes.data() + offset, chunk, &written);
        if (FAILED(hr))
        {
            return hr;
        }
        if (written == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        offset += written;
    }

    return S_OK;
}

HRESULT DeleteObjectById(const wil::com_ptr<IPortableDeviceContent>& content, std::wstring_view objectId, bool recursive)
{
    wil::com_ptr<IPortableDevicePropVariantCollection> objectIds;
    RETURN_IF_FAILED(CreateObjectIdCollection(objectId, objectIds));

    wil::com_ptr<IPortableDevicePropVariantCollection> results;
    RETURN_IF_FAILED(content->Delete(recursive ? PORTABLE_DEVICE_DELETE_WITH_RECURSION : PORTABLE_DEVICE_DELETE_NO_RECURSION, objectIds.get(), results.put()));
    return FirstFailureFromResultCollection(results);
}

HRESULT RenameObjectById(const wil::com_ptr<IPortableDeviceContent>& content, std::wstring_view objectId, std::wstring_view leafName, bool folder)
{
    wil::com_ptr<IPortableDeviceProperties> properties;
    RETURN_IF_FAILED(content->Properties(properties.put()));

    wil::com_ptr<IPortableDeviceValues> values;
    RETURN_IF_FAILED(CreatePortableDeviceValues(values));

    const std::wstring leafCopy(leafName);
    const std::wstring objectIdCopy(objectId);
    RETURN_IF_FAILED(values->SetStringValue(WPD_OBJECT_NAME, leafCopy.c_str()));
    if (! folder)
    {
        RETURN_IF_FAILED(values->SetStringValue(WPD_OBJECT_ORIGINAL_FILE_NAME, leafCopy.c_str()));
    }

    wil::com_ptr<IPortableDeviceValues> results;
    RETURN_IF_FAILED(properties->SetValues(objectIdCopy.c_str(), values.get(), results.put()));
    return FirstFailureFromValues(results);
}

HRESULT CopyOrMoveNativeObject(const wil::com_ptr<IPortableDeviceContent>& content,
                               std::wstring_view sourceObjectId,
                               std::wstring_view destinationFolderObjectId,
                               bool move)
{
    wil::com_ptr<IPortableDevicePropVariantCollection> objectIds;
    RETURN_IF_FAILED(CreateObjectIdCollection(sourceObjectId, objectIds));

    const std::wstring destinationCopy(destinationFolderObjectId);
    wil::com_ptr<IPortableDevicePropVariantCollection> results;
    const HRESULT hr =
        move ? content->Move(objectIds.get(), destinationCopy.c_str(), results.put()) : content->Copy(objectIds.get(), destinationCopy.c_str(), results.put());
    if (FAILED(hr))
    {
        return hr;
    }

    return FirstFailureFromResultCollection(results);
}

class WpdMtpBackend final : public IMtpBackend
{
public:
    explicit WpdMtpBackend(std::unique_ptr<WpdDeviceOperations> operations) noexcept : _operations(std::move(operations))
    {
    }
    ~WpdMtpBackend() override = default;

    WpdMtpBackend(const WpdMtpBackend&)            = delete;
    WpdMtpBackend(WpdMtpBackend&&)                 = delete;
    WpdMtpBackend& operator=(const WpdMtpBackend&) = delete;
    WpdMtpBackend& operator=(WpdMtpBackend&&)      = delete;

    MtpBackendInfo GetInfo() const noexcept override
    {
        return _operations->GetInfo();
    }

    void RequestCancel() noexcept override
    {
        if (_cancelState)
        {
            RequestWpdCancel(*_cancelState);
        }
    }

    HRESULT EnumerateDirectory(std::wstring_view path, std::vector<MtpItem>& items) noexcept override
    {
        items.clear();

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        const auto normalized = NormalizeMtpPath(path);
        if (normalized == L"/")
        {
            std::vector<DeviceDescriptor> devices;
            const HRESULT hr = EnumerateDevicesForBackend(devices);
            if (FAILED(hr))
            {
                return hr;
            }

            items.reserve(devices.size());
            for (const auto& device : devices)
            {
                items.push_back(MakeRootDeviceItem(device));
            }

            return S_OK;
        }

        return RunResolvedOperationWithCacheRetry(normalized, kReadAccess, [&](const ResolvedObject& resolved) noexcept
        {
            if ((resolved.item.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
            }

            ScopedActiveWpdContent activeContent(_cancelState.get(), resolved.content);
            const HRESULT enumHr = EnumerateObjectItemsForBackend(resolved.content, resolved.objectId, items);
            if (SUCCEEDED(enumHr))
            {
                RefreshPathCacheChildren(normalized, resolved, items);
            }
            return enumHr;
        });
    }

    HRESULT GetAttributes(std::wstring_view path, unsigned long& attributes) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject resolved;
        const HRESULT hr = ResolvePathCached(path, resolved, kReadAccess);
        if (FAILED(hr))
        {
            return hr;
        }

        attributes = resolved.item.attributes;
        return S_OK;
    }

    HRESULT GetBasicInformation(std::wstring_view path, FileSystemBasicInformation& info) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject resolved;
        const HRESULT hr = ResolvePathCached(path, resolved, kReadAccess);
        if (FAILED(hr))
        {
            return hr;
        }

        info.creationTime   = resolved.item.creationTime;
        info.lastAccessTime = resolved.item.lastAccessTime;
        info.lastWriteTime  = resolved.item.lastWriteTime;
        info.attributes     = resolved.item.attributes;
        return S_OK;
    }

    HRESULT ReadFile(std::wstring_view path, std::vector<std::byte>& bytes) noexcept override
    {
        bytes.clear();

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        const std::wstring normalized = NormalizeMtpPath(path);
        for (uint32_t attempt = 0u; attempt < 2u; ++attempt)
        {
            // A transfer is size-sensitive, so never trust metadata retained by an earlier lookup.
            InvalidatePathCacheSubtree(normalized);
            ResolvedObject resolved;
            const HRESULT resolveHr = ResolvePathCached(normalized, resolved, kReadAccess);
            if (FAILED(resolveHr))
            {
                return resolveHr;
            }
            if ((resolved.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            ScopedActiveWpdContent activeContent(_cancelState.get(), resolved.content);
            const HRESULT readHr = ReadPortableDeviceStream(resolved.content, resolved.objectId, resolved.item.sizeBytes, bytes);
            if (readHr == HRESULT_FROM_WIN32(ERROR_CRC) && attempt == 0u)
            {
                continue;
            }
            return FAILED(readHr) ? FailAndMaybeInvalidateCaches(readHr, resolved.pnpId) : S_OK;
        }
        return HRESULT_FROM_WIN32(ERROR_CRC);
    }

    HRESULT CreateFileReader(std::wstring_view path, std::shared_ptr<IMtpBackendFileReader>& reader) noexcept override
    {
        reader.reset();

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        const std::wstring normalized = NormalizeMtpPath(path);
        // Reader size and stream state must come from the same fresh object snapshot.
        InvalidatePathCacheSubtree(normalized);
        ResolvedObject resolved;
        const HRESULT hr = ResolvePathCached(normalized, resolved, kReadAccess);
        if (FAILED(hr))
        {
            return hr;
        }

        if ((resolved.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const HRESULT readerHr = _operations->CreateReader(resolved, _cancelState, reader);
        if (FAILED(readerHr))
        {
            InvalidatePathCacheSubtree(normalized);
            return FailAndMaybeInvalidateCaches(readerHr, resolved.pnpId);
        }
        return S_OK;
    }

    HRESULT GetFileSize(std::wstring_view path, uint64_t& sizeBytes) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        const std::wstring normalized = NormalizeMtpPath(path);
        InvalidatePathCacheSubtree(normalized);
        ResolvedObject resolved;
        const HRESULT hr = ResolvePathCached(normalized, resolved, kReadAccess);
        if (FAILED(hr))
        {
            return hr;
        }

        if ((resolved.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        sizeBytes = resolved.item.sizeBytes;
        return S_OK;
    }

    HRESULT WriteFile(std::wstring_view path, std::span<const std::byte> bytes, bool allowOverwrite) noexcept override
    {
        static_cast<void>(allowOverwrite);

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        return UploadFileObjectCached(path, bytes);
    }

    HRESULT CreateDirectory(std::wstring_view path) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        return CreateFolderObjectCached(path);
    }

    HRESULT DeleteItem(std::wstring_view path, bool recursive) noexcept override
    {
        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject resolved;
        RETURN_IF_FAILED(ResolvePathCached(path, resolved, kWriteAccess));
        if (resolved.root || resolved.deviceRoot)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ScopedActiveWpdContent activeContent(_cancelState.get(), resolved.content);
        if (! recursive && (resolved.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            std::vector<std::wstring> children;
            const HRESULT childrenHr = EnumerateObjectIds(resolved.content, resolved.objectId, children);
            if (FAILED(childrenHr))
            {
                return FailAndMaybeInvalidateCaches(childrenHr, resolved.pnpId);
            }
            if (! children.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
            }
        }

        const HRESULT deleteHr = DeleteObjectById(resolved.content, resolved.objectId, recursive);
        if (SUCCEEDED(deleteHr))
        {
            InvalidatePathCacheSubtree(path);
        }
        return FAILED(deleteHr) ? FailAndMaybeInvalidateCaches(deleteHr, resolved.pnpId) : S_OK;
    }

    HRESULT RenameItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        static_cast<void>(allowOverwrite);

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject source;
        RETURN_IF_FAILED(ResolvePathCached(sourcePath, source, kWriteAccess));
        if (source.root || source.deviceRoot)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ResolvedDestination destination;
        RETURN_IF_FAILED(ResolveDestinationPathCached(destinationPath, destination, kWriteAccess));
        if (! OrdinalString::EqualsNoCase(source.pnpId, destination.parent.pnpId) || ! SameNormalizedParentPath(sourcePath, destination))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        ScopedActiveWpdContent activeDestination(_cancelState.get(), destination.parent.content);
        const HRESULT destinationHr = EnsureDestinationDoesNotExistCached(destination);
        if (FAILED(destinationHr))
        {
            return FailAndMaybeInvalidateCaches(destinationHr, destination.parent.pnpId);
        }

        const bool folder = (source.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        ScopedActiveWpdContent activeSource(_cancelState.get(), source.content);
        const HRESULT renameHr = RenameObjectById(source.content, source.objectId, destination.leafName, folder);
        if (SUCCEEDED(renameHr))
        {
            InvalidatePathCacheSubtree(sourcePath);
            InvalidatePathCacheSubtree(destination.normalizedPath);
        }
        return FAILED(renameHr) ? FailAndMaybeInvalidateCaches(renameHr, source.pnpId) : S_OK;
    }

    HRESULT CopyItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        static_cast<void>(allowOverwrite);

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject source;
        RETURN_IF_FAILED(ResolvePathCached(sourcePath, source, kWriteAccess));
        if (source.root || source.deviceRoot)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ResolvedDestination destination;
        RETURN_IF_FAILED(ResolveDestinationPathCached(destinationPath, destination, kWriteAccess));
        ScopedActiveWpdContent activeDestination(_cancelState.get(), destination.parent.content);
        const HRESULT destinationHr = EnsureDestinationDoesNotExistCached(destination);
        if (FAILED(destinationHr))
        {
            return FailAndMaybeInvalidateCaches(destinationHr, destination.parent.pnpId);
        }

        const bool sameDevice         = OrdinalString::EqualsNoCase(source.pnpId, destination.parent.pnpId);
        const bool sourceDirectory    = (source.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const std::wstring sourceLeaf = std::wstring(LeafName(NormalizeMtpPath(sourcePath)));
        const bool preservesLeaf      = EqualsPathComponent(sourceLeaf, destination.leafName);

        if (sameDevice && preservesLeaf)
        {
            ScopedActiveWpdContent activeSource(_cancelState.get(), source.content);
            const HRESULT nativeHr = CopyOrMoveNativeObject(source.content, source.objectId, destination.parent.objectId, false);
            if (SUCCEEDED(nativeHr))
            {
                Debug::Perf::EmitValue(L"mtp.transfer.copy_bytes", source.item.sizeBytes, S_OK);
                InvalidatePathCacheSubtree(destination.normalizedPath);
                return S_OK;
            }
            if (sourceDirectory)
            {
                return FailAndMaybeInvalidateCaches(nativeHr, source.pnpId);
            }
        }
        else if (sourceDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const HRESULT transferHr = CopyOrMoveFileByTransferCached(sourcePath, source, destination, false);
        return FAILED(transferHr) ? FailAndMaybeInvalidateCaches(transferHr, source.pnpId) : S_OK;
    }

    HRESULT MoveItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        static_cast<void>(allowOverwrite);

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject source;
        RETURN_IF_FAILED(ResolvePathCached(sourcePath, source, kWriteAccess));
        if (source.root || source.deviceRoot)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        ResolvedDestination destination;
        RETURN_IF_FAILED(ResolveDestinationPathCached(destinationPath, destination, kWriteAccess));
        ScopedActiveWpdContent activeDestination(_cancelState.get(), destination.parent.content);
        const HRESULT destinationHr = EnsureDestinationDoesNotExistCached(destination);
        if (FAILED(destinationHr))
        {
            return FailAndMaybeInvalidateCaches(destinationHr, destination.parent.pnpId);
        }

        const bool sameDevice         = OrdinalString::EqualsNoCase(source.pnpId, destination.parent.pnpId);
        const bool sourceDirectory    = (source.item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const std::wstring sourceLeaf = std::wstring(LeafName(NormalizeMtpPath(sourcePath)));
        const bool preservesLeaf      = EqualsPathComponent(sourceLeaf, destination.leafName);

        if (sameDevice && preservesLeaf)
        {
            ScopedActiveWpdContent activeSource(_cancelState.get(), source.content);
            const HRESULT nativeHr = CopyOrMoveNativeObject(source.content, source.objectId, destination.parent.objectId, true);
            if (SUCCEEDED(nativeHr))
            {
                InvalidatePathCacheSubtree(sourcePath);
                InvalidatePathCacheSubtree(destination.normalizedPath);
                return S_OK;
            }
            if (sourceDirectory)
            {
                return FailAndMaybeInvalidateCaches(nativeHr, source.pnpId);
            }
        }
        else if (sourceDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const HRESULT transferHr = CopyOrMoveFileByTransferCached(sourcePath, source, destination, true);
        return FAILED(transferHr) ? FailAndMaybeInvalidateCaches(transferHr, source.pnpId) : S_OK;
    }

    HRESULT GetItemProperties(std::wstring_view path, std::string& jsonUtf8) noexcept override
    {
        jsonUtf8.clear();

        const ComInitialization com;
        if (! com.IsUsable())
        {
            return com.hr;
        }

        ResolvedObject resolved;
        const HRESULT hr = ResolvePathCached(path, resolved, kReadAccess);
        if (FAILED(hr))
        {
            return hr;
        }

        jsonUtf8 = std::format(
            R"json({{"version":1,"backend":"{}","path":"{}","name":"{}","objectId":"{}","persistentId":"{}","attributes":{},"sizeBytes":{},"streamable":{},"instrumentation":{{"deviceEnumerations":{},"sessionOpens":{},"childResolveEnumerations":{},"pathCacheHits":{},"pathCacheMisses":{}}}}})json",
            _operations->GetInfo().liveWpd ? "wpd" : "wpd-selftest",
            JsonEscapeUtf8(Utf8FromUtf16(path)),
            JsonEscapeUtf8(Utf8FromUtf16(resolved.item.name)),
            JsonEscapeUtf8(Utf8FromUtf16(resolved.item.objectId)),
            JsonEscapeUtf8(Utf8FromUtf16(resolved.item.persistentId)),
            resolved.item.attributes,
            resolved.item.sizeBytes,
            resolved.item.streamable ? "true" : "false",
            _deviceEnumerations.load(std::memory_order_acquire),
            _sessionOpens.load(std::memory_order_acquire),
            _childResolveEnumerations.load(std::memory_order_acquire),
            _pathCacheHits.load(std::memory_order_acquire),
            _pathCacheMisses.load(std::memory_order_acquire));
        return S_OK;
    }

private:
    struct CachedWpdSession
    {
        std::wstring pnpId;
        DWORD desiredAccess = 0;
        wil::com_ptr<IPortableDevice> device;
        wil::com_ptr<IPortableDeviceContent> content;
    };

    struct CachedResolvedPath
    {
        bool deviceRoot = false;
        std::wstring pnpId;
        std::wstring objectId;
        MtpItem item;
    };

    HRESULT EnumerateDevicesForBackend(std::vector<DeviceDescriptor>& devices) noexcept
    {
        _deviceEnumerations.fetch_add(1u, std::memory_order_acq_rel);
        return _operations->EnumerateDevices(devices);
    }

    HRESULT OpenDeviceSessionForBackend(std::wstring_view pnpId,
                                        DWORD desiredAccess,
                                        wil::com_ptr<IPortableDevice>& device,
                                        wil::com_ptr<IPortableDeviceContent>& content) noexcept
    {
        _sessionOpens.fetch_add(1u, std::memory_order_acq_rel);
        return _operations->OpenSession(pnpId, desiredAccess, device, content);
    }

    HRESULT EnumerateObjectItemsForBackend(const wil::com_ptr<IPortableDeviceContent>& content,
                                           std::wstring_view parentObjectId,
                                           std::vector<MtpItem>& items) noexcept
    {
        return _operations->EnumerateItems(content, parentObjectId, items);
    }

    HRESULT FindChildByDisplayNameForBackend(const wil::com_ptr<IPortableDeviceContent>& content,
                                             std::wstring_view parentObjectId,
                                             std::wstring_view childName,
                                             MtpItem& child) noexcept
    {
        _childResolveEnumerations.fetch_add(1u, std::memory_order_acq_rel);

        std::vector<MtpItem> items;
        RETURN_IF_FAILED(EnumerateObjectItemsForBackend(content, parentObjectId, items));

        for (const auto& item : items)
        {
            if (EqualsPathComponent(item.name, childName))
            {
                child = item;
                return S_OK;
            }
        }

        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    HRESULT FindDeviceDescriptorCached(std::wstring_view displayName, DeviceDescriptor& descriptor) noexcept
    {
        const std::wstring requestedKey = CaseFoldKey(displayName);
        if (const auto cached = _deviceDescriptorByDisplayKey.find(requestedKey); cached != _deviceDescriptorByDisplayKey.end())
        {
            descriptor = cached->second;
            return S_OK;
        }

        std::vector<DeviceDescriptor> devices;
        RETURN_IF_FAILED(EnumerateDevicesForBackend(devices));

        const std::optional<std::wstring> requestedHash = ExtractDeviceHashToken(displayName);
        for (const auto& device : devices)
        {
            if (OrdinalString::EqualsNoCase(device.displayName, displayName))
            {
                descriptor = device;
                _deviceDescriptorByDisplayKey[requestedKey]                     = descriptor;
                _deviceDescriptorByDisplayKey[CaseFoldKey(descriptor.displayName)] = descriptor;
                return S_OK;
            }
            if (requestedHash.has_value())
            {
                const std::wstring deviceHash = FormatMtpIdentityHash(device.pnpId);
                if (OrdinalString::EqualsNoCase(deviceHash, requestedHash.value()))
                {
                    descriptor = device;
                    _deviceDescriptorByDisplayKey[requestedKey]                     = descriptor;
                    _deviceDescriptorByDisplayKey[CaseFoldKey(descriptor.displayName)] = descriptor;
                    return S_OK;
                }
            }
        }

        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    HRESULT GetOrOpenSession(std::wstring_view pnpId,
                             DWORD desiredAccess,
                             wil::com_ptr<IPortableDevice>& device,
                             wil::com_ptr<IPortableDeviceContent>& content) noexcept
    {
        const std::wstring key = CaseFoldKey(pnpId);
        if (const auto cached = _sessionsByPnpId.find(key); cached != _sessionsByPnpId.end() &&
                                                          DesiredAccessIsCovered(cached->second.desiredAccess, desiredAccess))
        {
            device  = cached->second.device;
            content = cached->second.content;
            return S_OK;
        }

        CachedWpdSession session;
        session.pnpId         = std::wstring(pnpId);
        session.desiredAccess = desiredAccess;
        const HRESULT openHr  = OpenDeviceSessionForBackend(session.pnpId, desiredAccess, session.device, session.content);
        if (FAILED(openHr))
        {
            return FailAndMaybeInvalidateCaches(openHr);
        }

        device  = session.device;
        content = session.content;
        _sessionsByPnpId[key] = std::move(session);
        return S_OK;
    }

    HRESULT TryResolveCachedPath(std::wstring_view normalizedPath, DWORD desiredAccess, ResolvedObject& resolved, bool& found) noexcept
    {
        found = false;
        const std::wstring key = CaseFoldKey(normalizedPath);
        const auto cached      = _pathCache.find(key);
        if (cached == _pathCache.end())
        {
            return S_OK;
        }

        wil::com_ptr<IPortableDevice> device;
        wil::com_ptr<IPortableDeviceContent> content;
        const HRESULT sessionHr = GetOrOpenSession(cached->second.pnpId, desiredAccess, device, content);
        if (FAILED(sessionHr))
        {
            return FailAndMaybeInvalidateCaches(sessionHr);
        }

        resolved            = ResolvedObject{};
        resolved.deviceRoot = cached->second.deviceRoot;
        resolved.fromPathCache = true;
        resolved.pnpId      = cached->second.pnpId;
        resolved.objectId   = cached->second.objectId;
        resolved.item       = cached->second.item;
        resolved.device     = std::move(device);
        resolved.content    = std::move(content);
        _pathCacheHits.fetch_add(1u, std::memory_order_acq_rel);
        found = true;
        return S_OK;
    }

    void StoreResolvedPath(std::wstring_view normalizedPath, const ResolvedObject& resolved)
    {
        if (resolved.root || resolved.pnpId.empty())
        {
            return;
        }

        _pathCache[CaseFoldKey(normalizedPath)] = CachedResolvedPath{
            .deviceRoot = resolved.deviceRoot,
            .pnpId      = resolved.pnpId,
            .objectId   = resolved.objectId,
            .item       = resolved.item,
        };
    }

    void InvalidatePathCacheSubtree(std::wstring_view path)
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        const std::wstring key        = CaseFoldKey(normalized);
        const std::wstring prefix     = key == L"/" ? key : key + L"/";
        for (auto iter = _pathCache.begin(); iter != _pathCache.end();)
        {
            if (iter->first == key || (key != L"/" && iter->first.starts_with(prefix)))
            {
                iter = _pathCache.erase(iter);
            }
            else
            {
                ++iter;
            }
        }
    }

    void RefreshPathCacheChildren(std::wstring_view parentPath, const ResolvedObject& parent, const std::vector<MtpItem>& items)
    {
        const std::wstring normalizedParent = NormalizeMtpPath(parentPath);
        const std::wstring parentKey        = CaseFoldKey(normalizedParent);
        const std::wstring prefix           = parentKey == L"/" ? parentKey : parentKey + L"/";
        for (auto iter = _pathCache.begin(); iter != _pathCache.end();)
        {
            if (iter->first.starts_with(prefix) && iter->first != parentKey)
            {
                iter = _pathCache.erase(iter);
            }
            else
            {
                ++iter;
            }
        }

        for (const MtpItem& item : items)
        {
            ResolvedObject child;
            child.pnpId    = parent.pnpId;
            child.objectId = item.objectId;
            child.item     = item;
            child.device   = parent.device;
            child.content  = parent.content;
            StoreResolvedPath(JoinPath(normalizedParent, item.name), child);
        }
    }

    void InvalidateDeviceCaches(std::wstring_view pnpId) noexcept
    {
        const std::wstring deviceKey = CaseFoldKey(pnpId);
        _sessionsByPnpId.erase(deviceKey);
        for (auto iter = _deviceDescriptorByDisplayKey.begin(); iter != _deviceDescriptorByDisplayKey.end();)
        {
            if (OrdinalString::EqualsNoCase(iter->second.pnpId, pnpId))
            {
                iter = _deviceDescriptorByDisplayKey.erase(iter);
            }
            else
            {
                ++iter;
            }
        }
        for (auto iter = _pathCache.begin(); iter != _pathCache.end();)
        {
            if (OrdinalString::EqualsNoCase(iter->second.pnpId, pnpId))
            {
                iter = _pathCache.erase(iter);
            }
            else
            {
                ++iter;
            }
        }
    }

    void InvalidateAllCaches() noexcept
    {
        _deviceDescriptorByDisplayKey.clear();
        _sessionsByPnpId.clear();
        _pathCache.clear();
    }

    bool ShouldInvalidateAllCachesOnFailure(HRESULT hr) const noexcept
    {
        if (SUCCEEDED(hr))
        {
            return false;
        }

        switch (hr)
        {
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_INVALID_NAME):
        case HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED):
        case HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS):
        case HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY):
        case HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED):
        case HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT):
        case HRESULT_FROM_WIN32(ERROR_CANCELLED):
        case HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED):
            return false;
        default:
            return true;
        }
    }

    HRESULT FailAndMaybeInvalidateCaches(HRESULT hr, std::wstring_view pnpId = {}) noexcept
    {
        if (ShouldInvalidateAllCachesOnFailure(hr))
        {
            if (pnpId.empty())
            {
                InvalidateAllCaches();
            }
            else
            {
                InvalidateDeviceCaches(pnpId);
            }
        }
        return hr;
    }

    template<typename Operation>
    HRESULT RunResolvedOperationWithCacheRetry(std::wstring_view path, DWORD desiredAccess, Operation&& operation) noexcept
    {
        for (uint32_t attempt = 0u; attempt < 2u; ++attempt)
        {
            ResolvedObject resolved;
            const HRESULT resolveHr = ResolvePathCached(path, resolved, desiredAccess);
            if (FAILED(resolveHr))
            {
                return resolveHr;
            }

            const HRESULT operationHr = operation(resolved);
            if (SUCCEEDED(operationHr))
            {
                return S_OK;
            }
            if (attempt == 0u && resolved.fromPathCache && IsMissingObjectHr(operationHr))
            {
                InvalidatePathCacheSubtree(path);
                continue;
            }
            return FailAndMaybeInvalidateCaches(operationHr, resolved.pnpId);
        }
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    HRESULT ResolvePathCached(std::wstring_view path, ResolvedObject& resolved, DWORD desiredAccess = kReadAccess) noexcept
    {
        const auto normalized = NormalizeMtpPath(path);
        const auto segments   = SplitPathSegments(normalized);

        resolved = ResolvedObject{};
        if (segments.empty())
        {
            resolved.root     = true;
            resolved.objectId = WPD_DEVICE_OBJECT_ID;
            resolved.item     = MtpItem{
                .name           = L"/",
                .attributes     = FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY,
                .sizeBytes      = 0,
                .creationTime   = 0,
                .lastAccessTime = 0,
                .lastWriteTime  = 0,
                .changeTime     = 0,
                .persistentId   = L"",
                .objectId       = WPD_DEVICE_OBJECT_ID,
                .streamable     = false,
            };
            return S_OK;
        }

        bool foundCached = false;
        const HRESULT cachedHr = TryResolveCachedPath(normalized, desiredAccess, resolved, foundCached);
        if (FAILED(cachedHr))
        {
            return FailAndMaybeInvalidateCaches(cachedHr);
        }
        if (foundCached)
        {
            return S_OK;
        }
        _pathCacheMisses.fetch_add(1u, std::memory_order_acq_rel);

        DeviceDescriptor deviceDescriptor;
        const HRESULT deviceHr = FindDeviceDescriptorCached(segments.front(), deviceDescriptor);
        if (FAILED(deviceHr))
        {
            return FailAndMaybeInvalidateCaches(deviceHr);
        }
        const HRESULT sessionHr = GetOrOpenSession(deviceDescriptor.pnpId, desiredAccess, resolved.device, resolved.content);
        if (FAILED(sessionHr))
        {
            return FailAndMaybeInvalidateCaches(sessionHr);
        }
        ScopedActiveWpdContent activeContent(_cancelState.get(), resolved.content);

        resolved.pnpId      = deviceDescriptor.pnpId;
        resolved.objectId   = WPD_DEVICE_OBJECT_ID;
        resolved.item       = MakeRootDeviceItem(deviceDescriptor);
        resolved.deviceRoot = true;

        std::wstring currentPath = JoinPath(L"/", segments.front());
        StoreResolvedPath(currentPath, resolved);

        if (segments.size() == 1)
        {
            return S_OK;
        }

        std::wstring parentId = WPD_DEVICE_OBJECT_ID;
        resolved.deviceRoot   = false;
        for (std::size_t index = 1; index < segments.size(); ++index)
        {
            currentPath = JoinPath(currentPath, segments[index]);

            bool childCached = false;
            const HRESULT childCachedHr = TryResolveCachedPath(currentPath, desiredAccess, resolved, childCached);
            if (FAILED(childCachedHr))
            {
                return FailAndMaybeInvalidateCaches(childCachedHr);
            }
            if (childCached)
            {
                parentId = resolved.objectId;
                continue;
            }
            _pathCacheMisses.fetch_add(1u, std::memory_order_acq_rel);

            MtpItem child;
            const HRESULT childHr = FindChildByDisplayNameForBackend(resolved.content, parentId, segments[index], child);
            if (FAILED(childHr))
            {
                return FailAndMaybeInvalidateCaches(childHr);
            }
            parentId            = child.objectId;
            resolved.objectId   = parentId;
            resolved.item       = std::move(child);
            resolved.deviceRoot = false;
            StoreResolvedPath(currentPath, resolved);
        }

        return S_OK;
    }

    HRESULT ResolveDestinationPathCached(std::wstring_view path,
                                         ResolvedDestination& destination,
                                         DWORD desiredAccess = kWriteAccess) noexcept
    {
        destination                = ResolvedDestination{};
        destination.normalizedPath = NormalizeMtpPath(path);
        if (destination.normalizedPath == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        destination.leafName = std::wstring(LeafName(destination.normalizedPath));
        if (destination.leafName.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        destination.parentPath = ParentPath(destination.normalizedPath);
        RETURN_IF_FAILED(ResolvePathCached(destination.parentPath, destination.parent, desiredAccess));
        if (destination.parent.root || destination.parent.deviceRoot || (destination.parent.item.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        return S_OK;
    }

    HRESULT FindDestinationChildCached(const ResolvedDestination& destination, MtpItem& child) noexcept
    {
        const HRESULT hr = FindChildByDisplayNameForBackend(destination.parent.content, destination.parent.objectId, destination.leafName, child);
        if (SUCCEEDED(hr))
        {
            ResolvedObject resolved;
            resolved.pnpId    = destination.parent.pnpId;
            resolved.objectId = child.objectId;
            resolved.item     = child;
            resolved.device   = destination.parent.device;
            resolved.content  = destination.parent.content;
            StoreResolvedPath(destination.normalizedPath, resolved);
        }
        return hr;
    }

    HRESULT EnsureDestinationDoesNotExistCached(const ResolvedDestination& destination) noexcept
    {
        MtpItem existing;
        const HRESULT findHr = FindDestinationChildCached(destination, existing);
        if (SUCCEEDED(findHr))
        {
            return (existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 ? HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) : HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! IsMissingObjectHr(findHr))
        {
            return findHr;
        }

        return S_OK;
    }

    HRESULT UploadFileObjectCached(std::wstring_view path, std::span<const std::byte> bytes) noexcept
    {
        ResolvedDestination destination;
        RETURN_IF_FAILED(ResolveDestinationPathCached(path, destination, kWriteAccess));
        ScopedActiveWpdContent activeContent(_cancelState.get(), destination.parent.content);
        RETURN_IF_FAILED(EnsureDestinationDoesNotExistCached(destination));

        wil::com_ptr<IPortableDeviceValues> values;
        RETURN_IF_FAILED(CreateObjectValues(destination.parent.objectId,
                                            destination.leafName,
                                            WPD_CONTENT_TYPE_GENERIC_FILE,
                                            WPD_OBJECT_FORMAT_UNSPECIFIED,
                                            static_cast<uint64_t>(bytes.size()),
                                            true,
                                            values));

        wil::com_ptr<IStream> stream;
        DWORD optimalBufferSize = 0;
        PWSTR rawCookie         = nullptr;
        auto freeCookie         = wil::scope_exit([&rawCookie]() noexcept
        {
            if (rawCookie != nullptr)
            {
                CoTaskMemFree(rawCookie);
            }
        });

        RETURN_IF_FAILED(destination.parent.content->CreateObjectWithPropertiesAndData(values.get(), stream.put(), &optimalBufferSize, &rawCookie));
        ScopedActiveWpdStream activeStream(_cancelState.get(), stream);

        bool committed  = false;
        auto abortWrite = wil::scope_exit([&]() noexcept
        {
            if (! committed)
            {
                AbortPortableDeviceWriteStream(stream.get());
            }
        });

        RETURN_IF_FAILED(WritePortableDeviceStream(stream.get(), optimalBufferSize, bytes));
        RETURN_IF_FAILED(stream->Commit(STGC_DEFAULT));
        committed = true;

        InvalidatePathCacheSubtree(destination.normalizedPath);
        Debug::Perf::EmitValue(L"mtp.transfer.write_bytes", static_cast<uint64_t>(bytes.size()), S_OK);
        return S_OK;
    }

    HRESULT CreateFolderObjectCached(std::wstring_view path) noexcept
    {
        ResolvedDestination destination;
        RETURN_IF_FAILED(ResolveDestinationPathCached(path, destination, kWriteAccess));
        ScopedActiveWpdContent activeContent(_cancelState.get(), destination.parent.content);
        RETURN_IF_FAILED(EnsureDestinationDoesNotExistCached(destination));

        wil::com_ptr<IPortableDeviceValues> values;
        RETURN_IF_FAILED(CreateObjectValues(
            destination.parent.objectId, destination.leafName, WPD_CONTENT_TYPE_FOLDER, WPD_OBJECT_FORMAT_PROPERTIES_ONLY, std::nullopt, false, values));

        PWSTR rawObjectId = nullptr;
        auto freeObjectId = wil::scope_exit([&rawObjectId]() noexcept
        {
            if (rawObjectId != nullptr)
            {
                CoTaskMemFree(rawObjectId);
            }
        });

        const HRESULT hr = destination.parent.content->CreateObjectWithPropertiesOnly(values.get(), &rawObjectId);
        if (SUCCEEDED(hr))
        {
            InvalidatePathCacheSubtree(destination.normalizedPath);
        }
        return hr;
    }

    HRESULT CopyOrMoveFileByTransferCached(std::wstring_view sourcePath, const ResolvedObject& source, const ResolvedDestination& destination, bool move) noexcept
    {
        std::vector<std::byte> bytes;
        {
            ScopedActiveWpdContent activeSource(_cancelState.get(), source.content);
            RETURN_IF_FAILED(ReadPortableDeviceStream(source.content, source.objectId, source.item.sizeBytes, bytes));
        }
        RETURN_IF_FAILED(UploadFileObjectCached(destination.normalizedPath, bytes));

        if (move)
        {
            ScopedActiveWpdContent activeSource(_cancelState.get(), source.content);
            const HRESULT deleteHr = DeleteObjectById(source.content, source.objectId, false);
            Debug::Perf::EmitValue(L"mtp.transfer.move_fallback_bytes", static_cast<uint64_t>(bytes.size()), deleteHr);
            if (SUCCEEDED(deleteHr))
            {
                InvalidatePathCacheSubtree(sourcePath);
            }
            if (FAILED(deleteHr))
            {
                return deleteHr;
            }
        }
        else
        {
            Debug::Perf::EmitValue(L"mtp.transfer.copy_bytes", static_cast<uint64_t>(bytes.size()), S_OK);
        }

        return S_OK;
    }

    std::shared_ptr<WpdCancellationState> _cancelState = std::make_shared<WpdCancellationState>();
    std::unique_ptr<WpdDeviceOperations> _operations;
    std::unordered_map<std::wstring, DeviceDescriptor> _deviceDescriptorByDisplayKey;
    std::unordered_map<std::wstring, CachedWpdSession> _sessionsByPnpId;
    std::unordered_map<std::wstring, CachedResolvedPath> _pathCache;
    std::atomic_uint64_t _deviceEnumerations{0};
    std::atomic_uint64_t _sessionOpens{0};
    std::atomic_uint64_t _childResolveEnumerations{0};
    std::atomic_uint64_t _pathCacheHits{0};
    std::atomic_uint64_t _pathCacheMisses{0};
};

} // namespace

HRESULT ReadConnectionBrowseDevicePersistentId(std::wstring_view pnpId, std::wstring& devicePuid) noexcept
{
    if (pnpId.empty())
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IPortableDevice> device;
    wil::com_ptr<IPortableDeviceContent> content;
    HRESULT hr = OpenDeviceSession(pnpId, kReadAccess, device, content);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IPortableDeviceProperties> properties;
    hr = content->Properties(properties.put());
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IPortableDeviceValues> values;
    hr = properties->GetValues(WPD_DEVICE_OBJECT_ID, nullptr, values.put());
    if (FAILED(hr))
    {
        return hr;
    }

    const std::optional<std::wstring> persistentId = GetStringProperty(values, WPD_OBJECT_PERSISTENT_UNIQUE_ID);
    if (! persistentId.has_value() || persistentId.value().empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    devicePuid = persistentId.value();
    return S_OK;
}

[[nodiscard]] bool IsDirectoryItem(const MtpItem& item) noexcept
{
    return (item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
}

[[nodiscard]] bool MatchesConnectionBrowseDeviceId(const MtpItem& item, std::wstring_view parentDeviceId) noexcept
{
    return OrdinalString::EqualsNoCase(item.persistentId, parentDeviceId) || OrdinalString::EqualsNoCase(item.objectId, parentDeviceId) ||
           OrdinalString::EqualsNoCase(item.name, parentDeviceId);
}

HRESULT EnumerateMtpConnectionBrowseDevices(std::vector<MtpConnectionBrowseDevice>& devices) noexcept
{
    devices.clear();

    const ComInitialization com(/*initializeIfNeeded=*/true);
    if (! com.IsUsable())
    {
        return com.hr;
    }

    std::vector<DeviceDescriptor> descriptors;
    HRESULT hr = EnumerateDevices(descriptors);
    if (FAILED(hr))
    {
        return hr;
    }

    devices.reserve(descriptors.size());
    for (const DeviceDescriptor& descriptor : descriptors)
    {
        std::wstring devicePuid = descriptor.pnpId;
        const HRESULT puidHr = ReadConnectionBrowseDevicePersistentId(descriptor.pnpId, devicePuid);
        if (FAILED(puidHr) || devicePuid.empty())
        {
            devicePuid = descriptor.pnpId;
        }

        devices.push_back(MtpConnectionBrowseDevice{
            .pnpId        = descriptor.pnpId,
            .friendlyName = descriptor.name,
            .devicePuid   = std::move(devicePuid),
        });
    }

    return S_OK;
}

HRESULT EnumerateMtpConnectionBrowseStorages(std::wstring_view parentDeviceId, std::vector<MtpConnectionBrowseStorage>& storages) noexcept
{
    storages.clear();
    if (parentDeviceId.empty())
    {
        return E_INVALIDARG;
    }

    const ComInitialization com(/*initializeIfNeeded=*/true);
    if (! com.IsUsable())
    {
        return com.hr;
    }

    wil::com_ptr<IPortableDevice> device;
    wil::com_ptr<IPortableDeviceContent> content;
    HRESULT hr = OpenDeviceSession(parentDeviceId, kReadAccess, device, content);
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<MtpItem> items;
    hr = EnumerateObjectItems(content, WPD_DEVICE_OBJECT_ID, items);
    if (FAILED(hr))
    {
        return hr;
    }

    storages.reserve(items.size());
    for (const MtpItem& item : items)
    {
        if (! IsDirectoryItem(item) || item.name.empty())
        {
            continue;
        }

        storages.push_back(MtpConnectionBrowseStorage{
            .name         = item.name,
            .persistentId = item.persistentId,
            .objectId     = item.objectId,
            .initialPath  = L"/" + item.name,
        });
    }

    return S_OK;
}

HRESULT EnumerateMtpConnectionBrowseDevicesFromBackend(IMtpBackend& backend, std::vector<MtpConnectionBrowseDevice>& devices) noexcept
{
    devices.clear();

    std::vector<MtpItem> rootItems;
    HRESULT hr = backend.EnumerateDirectory(L"/", rootItems);
    if (FAILED(hr))
    {
        return hr;
    }

    devices.reserve(rootItems.size());
    for (const MtpItem& item : rootItems)
    {
        if (! IsDirectoryItem(item) || item.name.empty())
        {
            continue;
        }

        std::wstring pnpId = ! item.persistentId.empty() ? item.persistentId : (! item.objectId.empty() ? item.objectId : item.name);
        devices.push_back(MtpConnectionBrowseDevice{
            .pnpId        = pnpId,
            .friendlyName = item.name,
            .devicePuid   = std::move(pnpId),
        });
    }

    return S_OK;
}

HRESULT EnumerateMtpConnectionBrowseStoragesFromBackend(IMtpBackend& backend,
                                                        std::wstring_view parentDeviceId,
                                                        std::vector<MtpConnectionBrowseStorage>& storages) noexcept
{
    storages.clear();
    if (parentDeviceId.empty())
    {
        return E_INVALIDARG;
    }

    std::vector<MtpItem> rootItems;
    HRESULT hr = backend.EnumerateDirectory(L"/", rootItems);
    if (FAILED(hr))
    {
        return hr;
    }

    std::optional<std::wstring> deviceRootPath;
    for (const MtpItem& item : rootItems)
    {
        if (IsDirectoryItem(item) && MatchesConnectionBrowseDeviceId(item, parentDeviceId))
        {
            deviceRootPath = JoinPath(L"/", item.name);
            break;
        }
    }
    if (! deviceRootPath.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    std::vector<MtpItem> items;
    hr = backend.EnumerateDirectory(deviceRootPath.value(), items);
    if (FAILED(hr))
    {
        return hr;
    }

    storages.reserve(items.size());
    for (const MtpItem& item : items)
    {
        if (! IsDirectoryItem(item) || item.name.empty())
        {
            continue;
        }

        storages.push_back(MtpConnectionBrowseStorage{
            .name         = item.name,
            .persistentId = item.persistentId,
            .objectId     = item.objectId,
            .initialPath  = L"/" + item.name,
        });
    }

    return S_OK;
}

std::unique_ptr<IMtpBackend> CreateWpdMtpBackend() noexcept
{
    return std::make_unique<WpdMtpBackend>(std::make_unique<LiveWpdDeviceOperations>());
}

#ifdef _DEBUG
HRESULT CreateSelfTestWpdMtpBackend(std::string_view optionsJsonUtf8, std::unique_ptr<IMtpBackend>& backend) noexcept
{
    backend.reset();
    SelfTestWpdOptions options;
    if (! optionsJsonUtf8.empty())
    {
        yyjson_doc* doc = yyjson_read(optionsJsonUtf8.data(), optionsJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! doc)
        {
            return E_INVALIDARG;
        }
        auto freeDoc     = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });
        yyjson_val* root = yyjson_doc_get_root(doc);
        if (! root || ! yyjson_is_obj(root))
        {
            return E_INVALIDARG;
        }

        yyjson_obj_iter iter = yyjson_obj_iter_with(root);
        while (yyjson_val* key = yyjson_obj_iter_next(&iter))
        {
            const char* keyText = yyjson_get_str(key);
            yyjson_val* value   = yyjson_obj_iter_get_val(key);
            if (! keyText || ! value)
            {
                return E_INVALIDARG;
            }
            if (std::string_view(keyText) == "readFileDelayMs")
            {
                if (! yyjson_is_uint(value) || yyjson_get_uint(value) > 30'000u)
                {
                    return E_INVALIDARG;
                }
                options.readFileDelayMs = static_cast<uint32_t>(yyjson_get_uint(value));
            }
            else if (std::string_view(keyText) == "sessionDeathOnce")
            {
                if (! yyjson_is_bool(value))
                {
                    return E_INVALIDARG;
                }
                options.sessionDeathOnce = yyjson_get_bool(value) != 0;
            }
            else if (std::string_view(keyText) == "changeFileSizeAfterFirstLookup")
            {
                if (! yyjson_is_bool(value))
                {
                    return E_INVALIDARG;
                }
                options.changeFileSizeAfterFirstLookup = yyjson_get_bool(value) != 0;
            }
            else
            {
                return E_INVALIDARG;
            }
        }
    }

    backend = std::make_unique<WpdMtpBackend>(std::make_unique<SelfTestWpdDeviceOperations>(options));
    return S_OK;
}
#endif
} // namespace FileSystemMtpInternal

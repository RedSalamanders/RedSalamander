#include "Framework.h"

#include "FileOperationsPowerRequest.h"

#include "Helpers.h"
#include "resource.h"

#include <cstddef>
#include <new>
#include <string>

namespace FileOperations
{
namespace
{
wil::unique_handle CreateWindowsPowerRequest(void* /*context*/, const wchar_t* reason) noexcept
{
    // PowerCreateRequest takes a writable string; the request copies it before the call returns.
    std::wstring text(reason != nullptr ? reason : L"");
    REASON_CONTEXT context{};
    context.Version                   = POWER_REQUEST_CONTEXT_VERSION;
    context.Flags                     = POWER_REQUEST_CONTEXT_SIMPLE_STRING;
    context.Reason.SimpleReasonString = text.data();
    const HANDLE handle               = PowerCreateRequest(&context);
    if (handle == INVALID_HANDLE_VALUE || handle == nullptr)
        return {};
    return wil::unique_handle(handle);
}

bool SetWindowsPowerRequest(void* /*context*/, HANDLE request) noexcept
{
    return PowerSetRequest(request, PowerRequestSystemRequired) != FALSE;
}

bool ClearWindowsPowerRequest(void* /*context*/, HANDLE request) noexcept
{
    return PowerClearRequest(request, PowerRequestSystemRequired) != FALSE;
}

[[nodiscard]] PowerRequestApi WindowsPowerRequestApi() noexcept
{
    PowerRequestApi api{};
    api.create = &CreateWindowsPowerRequest;
    api.set    = &SetWindowsPowerRequest;
    api.clear  = &ClearWindowsPowerRequest;
    return api;
}
} // namespace

PowerRequest::PowerRequest() noexcept : _api(WindowsPowerRequestApi())
{
}

PowerRequest::PowerRequest(const PowerRequestApi& api) noexcept
    : _api(api.create != nullptr && api.set != nullptr && api.clear != nullptr ? api : WindowsPowerRequestApi())
{
}

PowerRequest::~PowerRequest()
{
    // Closing the handle ends the request as well; clearing first keeps a fake API's counts honest.
    if (_held && _request)
        static_cast<void>(_api.clear(_api.context, _request.get()));
}

PowerRequest& PowerRequest::Shared() noexcept
{
    // Never destroyed: a worker may give its share back while the process exits, after static destructors have run.
    // Windows ends the request with the process.
    alignas(PowerRequest) static std::byte storage[sizeof(PowerRequest)];
    static PowerRequest* const request = ::new (static_cast<void*>(storage)) PowerRequest();
    return *request;
}

void PowerRequest::Acquire() noexcept
{
    std::scoped_lock lock(_mutex);
    ++_holders;
    if (_holders != 1u)
        return;
    ++_acquireTransitions;
    SetLocked();
}

void PowerRequest::Release() noexcept
{
    std::scoped_lock lock(_mutex);
    if (_holders == 0u)
        return;
    --_holders;
    if (_holders != 0u)
        return;
    ++_releaseTransitions;
    ClearLocked();
}

PowerRequest::Counters PowerRequest::GetCounters() const noexcept
{
    std::scoped_lock lock(_mutex);
    Counters counters{};
    counters.holders            = _holders;
    counters.acquireTransitions = _acquireTransitions;
    counters.releaseTransitions = _releaseTransitions;
    counters.failures           = _failures;
    counters.held               = _held;
    return counters;
}

void PowerRequest::SetLocked() noexcept
{
    if (! _request)
    {
        std::wstring reason = LoadStringResource(nullptr, IDS_FILEOPS_KEEP_AWAKE_REASON);
        if (reason.empty())
            reason = L"File operation in progress"; // only if the resource cannot be loaded; Windows needs a non-empty reason
        _request = _api.create(_api.context, reason.c_str());
        if (! _request)
        {
            NoteFailureLocked(L"create", GetLastError());
            return;
        }
    }
    if (! _api.set(_api.context, _request.get()))
    {
        NoteFailureLocked(L"set", GetLastError());
        return;
    }
    _held = true;
}

void PowerRequest::ClearLocked() noexcept
{
    if (! _held)
        return;
    // Whatever the call returns, this object no longer claims the request.
    _held = false;
    if (! _api.clear(_api.context, _request.get()))
        NoteFailureLocked(L"clear", GetLastError());
}

void PowerRequest::NoteFailureLocked(const wchar_t* const step, const DWORD error) noexcept
{
    ++_failures;
    // One line per object: a machine that refuses the request would otherwise log it for every task.
    if (_failureLogged)
        return;
    _failureLogged = true;
    Debug::Info(L"File Operations keep-awake power request: {} failed (error {}); the operation continues without it.", step, error);
}
} // namespace FileOperations

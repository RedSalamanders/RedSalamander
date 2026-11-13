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
wil::unique_handle CreateWindowsPowerRequest(void* /*context*/, wchar_t* reason) noexcept
{
    // PowerCreateRequest copies the object's preallocated writable reason. Preserve the
    // API's exact last error through return-value construction and any local teardown.
    REASON_CONTEXT context{};
    context.Version                   = POWER_REQUEST_CONTEXT_VERSION;
    context.Flags                     = POWER_REQUEST_CONTEXT_SIMPLE_STRING;
    context.Reason.SimpleReasonString = reason;
    const HANDLE handle               = PowerCreateRequest(&context);
    [[maybe_unused]] const wil::last_error_context preserveError;
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

PowerRequest::PowerRequest() noexcept : PowerRequest(WindowsPowerRequestApi())
{
}

PowerRequest::PowerRequest(const PowerRequestApi& api) noexcept
    : _api(api.create != nullptr && api.set != nullptr && api.clear != nullptr ? api : WindowsPowerRequestApi()),
      _reason(LoadStringResource(nullptr, IDS_FILEOPS_KEEP_AWAKE_REASON))
{
    if (_reason.empty())
        _reason = L"File operation in progress"; // Windows requires a non-empty reason even when the resource is unavailable.
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
    if (_holders == 1u)
        ++_acquireTransitions;
    // A set that failed for the first holder is tried again for each later one: otherwise overlapping tasks would run
    // without the request until every one of them had finished.
    if (! _held)
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
    counters.lastError          = _lastError;
    counters.held               = _held;
    counters.failureLogged      = _failureLogged;
    return counters;
}

void PowerRequest::SetLocked() noexcept
{
    if (! _request)
    {
        _request = _api.create(_api.context, _reason.data());
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
    if (! _api.clear(_api.context, _request.get()))
    {
        const DWORD error = GetLastError();
        // A failed clear may leave the kernel reference active. Closing its owning
        // handle ends that request; a later Acquire must create a new one.
        _request.reset();
        NoteFailureLocked(L"clear", error);
    }
    _held = false;
}

void PowerRequest::NoteFailureLocked(const wchar_t* const step, const DWORD error) noexcept
{
    ++_failures;
    _lastError = error;
    // One line per object: a machine that refuses the request would otherwise log it for every task.
    if (_failureLogged)
        return;
    _failureLogged = true;
    Debug::Info(L"File Operations keep-awake power request: {} failed (error {}); the operation continues without it.", step, error);
}
} // namespace FileOperations

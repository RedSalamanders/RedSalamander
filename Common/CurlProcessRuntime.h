#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <format>
#include <memory>
#include <mutex>
#include <string>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::CurlRuntime
{
// Coordinates libcurl's process-global initialization across independently
// unloadable plugin DLLs. Each participating module owns one ProcessLease and
// must release it at its plugin quiet point, after destroying every easy/share
// handle. The final participant performs curl_global_cleanup outside loader lock.
class ProcessLease final
{
public:
    using InitializeCallback = HRESULT (*)() noexcept;
    using CleanupCallback    = void (*)() noexcept;

    ProcessLease()  = default;
    ~ProcessLease() = default;

    ProcessLease(const ProcessLease&)            = delete;
    ProcessLease(ProcessLease&&)                 = delete;
    ProcessLease& operator=(const ProcessLease&) = delete;
    ProcessLease& operator=(ProcessLease&&)      = delete;

    [[nodiscard]] HRESULT Acquire(InitializeCallback initialize) noexcept
    {
        if (! initialize)
        {
            return E_INVALIDARG;
        }

        std::lock_guard localLock(_localMutex);
        if (_acquired)
        {
            return S_OK;
        }

        const HRESULT transportHr = EnsureTransport();
        if (FAILED(transportHr))
        {
            return transportHr;
        }

        const DWORD waitResult = WaitForSingleObject(_processMutex.get(), INFINITE);
        if (waitResult != WAIT_OBJECT_0 && waitResult != WAIT_ABANDONED)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        const auto releaseMutex = wil::scope_exit([this]() noexcept { static_cast<void>(ReleaseMutex(_processMutex.get())); });

        if (_state->version != kProtocolVersion || _state->participantCount < 0 || _state->participantCount > kMaxParticipants)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        if (_state->participantCount == 0)
        {
            const HRESULT initializeHr = initialize();
            if (FAILED(initializeHr))
            {
                return initializeHr;
            }
            _state->initialized = 1;
        }
        else if (_state->initialized == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        ++_state->participantCount;
        _acquired = true;
        return S_OK;
    }

    [[nodiscard]] bool Release(CleanupCallback cleanup) noexcept
    {
        if (! cleanup)
        {
            return false;
        }

        std::lock_guard localLock(_localMutex);
        if (! _acquired)
        {
            return true;
        }
        if (! _processMutex || ! _state)
        {
            return false;
        }

        const DWORD waitResult = WaitForSingleObject(_processMutex.get(), INFINITE);
        if (waitResult != WAIT_OBJECT_0 && waitResult != WAIT_ABANDONED)
        {
            return false;
        }
        const auto releaseMutex = wil::scope_exit([this]() noexcept { static_cast<void>(ReleaseMutex(_processMutex.get())); });

        if (_state->version != kProtocolVersion || _state->participantCount <= 0 || _state->initialized == 0)
        {
            return false;
        }

        --_state->participantCount;
        if (_state->participantCount == 0)
        {
            cleanup();
            _state->initialized = 0;
        }

        _acquired = false;
        return true;
    }

    [[nodiscard]] bool IsAcquired() const noexcept
    {
        std::lock_guard lock(_localMutex);
        return _acquired;
    }

private:
    struct SharedState
    {
        uint32_t version      = 0;
        LONG participantCount = 0;
        LONG initialized      = 0;
    };

    static constexpr uint32_t kProtocolVersion = 1u;
    static constexpr LONG kMaxParticipants     = 32;

    [[nodiscard]] HRESULT EnsureTransport() noexcept
    {
        if (_processMutex && _mapping && _state)
        {
            return S_OK;
        }

        const DWORD processId        = GetCurrentProcessId();
        const std::wstring mutexName = std::format(L"Local\\RedSalamander.CurlRuntime.{}.Mutex", processId);
        const std::wstring mapName   = std::format(L"Local\\RedSalamander.CurlRuntime.{}.State", processId);

        wil::unique_handle processMutex(CreateMutexW(nullptr, FALSE, mutexName.c_str()));
        if (! processMutex)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const DWORD waitResult = WaitForSingleObject(processMutex.get(), INFINITE);
        if (waitResult != WAIT_OBJECT_0 && waitResult != WAIT_ABANDONED)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        const HANDLE lockedMutex = processMutex.get();
        const auto releaseMutex  = wil::scope_exit([lockedMutex]() noexcept { static_cast<void>(ReleaseMutex(lockedMutex)); });

        SetLastError(ERROR_SUCCESS);
        wil::unique_handle mapping(CreateFileMappingW(INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0u, sizeof(SharedState), mapName.c_str()));
        if (! mapping)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        const bool created = GetLastError() != ERROR_ALREADY_EXISTS;

        wil::unique_mapview_ptr<SharedState> state(static_cast<SharedState*>(MapViewOfFile(mapping.get(), FILE_MAP_ALL_ACCESS, 0u, 0u, sizeof(SharedState))));
        if (! state)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (created)
        {
            *state         = {};
            state->version = kProtocolVersion;
        }
        else if (state->version != kProtocolVersion)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        _processMutex = std::move(processMutex);
        _mapping      = std::move(mapping);
        _state        = std::move(state);
        return S_OK;
    }

    mutable std::mutex _localMutex;
    wil::unique_handle _processMutex;
    wil::unique_handle _mapping;
    wil::unique_mapview_ptr<SharedState> _state;
    bool _acquired = false;
};
} // namespace Common::CurlRuntime

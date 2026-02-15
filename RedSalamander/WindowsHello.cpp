#include "WindowsHello.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#include <memory>
#include <new>
#include <optional>

#pragma warning(push)
// 4625 4626: deleted copy/move
// 4265: class has virtual functions, but destructor is not virtual
// 5026: move constructor was implicitly defined as deleted
// 5027: move assignment operator was implicitly defined as deleted
// 5246: the initialization of a subobject should be wrapped in braces
#pragma warning(disable : 4625 4626 4265 5026 5027 5246)
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Security.Credentials.UI.h>
#include <winrt/base.h>
#pragma warning(pop)

#include <userconsentverifierinterop.h> // IUserConsentVerifierInterop

#pragma comment(lib, "RuntimeObject.lib")
#pragma comment(lib, "WindowsApp.lib")

namespace RedSalamander::Security
{
namespace
{
[[nodiscard]] HRESULT WaitForHandleWithMessagePump(HANDLE handle) noexcept
{
    if (! handle)
    {
        return E_INVALIDARG;
    }

    HANDLE handles[] = {handle};
    for (;;)
    {
        const DWORD wait = MsgWaitForMultipleObjectsEx(1, handles, INFINITE, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        if (wait == WAIT_OBJECT_0)
        {
            return S_OK;
        }

        if (wait == WAIT_OBJECT_0 + 1)
        {
            MSG msg{};
            while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == WM_QUIT)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            continue;
        }

        if (wait == WAIT_FAILED)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        return E_FAIL;
    }
}

template <typename T> [[nodiscard]] HRESULT WaitForOperationWithMessagePump(winrt::Windows::Foundation::IAsyncOperation<T>& operation, T& resultOut) noexcept
{
    struct State
    {
        State()                        = default;
        State(const State&)            = delete;
        State& operator=(const State&) = delete;
        State(State&&)                 = delete;
        State& operator=(State&&)      = delete;

        wil::unique_handle completedEvent;
        std::optional<T> result;
        HRESULT hr = E_FAIL;
    };

    std::shared_ptr<State> state;
    state = std::make_shared<State>();

    state->completedEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! state->completedEvent)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    operation.Completed(
        [state](auto const& op, winrt::Windows::Foundation::AsyncStatus status) noexcept
        {
            switch (status)
            {
                case winrt::Windows::Foundation::AsyncStatus::Started:
                    state->hr = E_PENDING;
                    return; // Completed handlers should not be invoked for Started; keep waiting.
                case winrt::Windows::Foundation::AsyncStatus::Completed:
                    state->result = op.GetResults();
                    state->hr     = S_OK;
                    break;
                case winrt::Windows::Foundation::AsyncStatus::Canceled: state->hr = HRESULT_FROM_WIN32(ERROR_CANCELLED); break;
                case winrt::Windows::Foundation::AsyncStatus::Error: state->hr = op.ErrorCode(); break;
                default: state->hr = E_FAIL; break;
            }

            static_cast<void>(SetEvent(state->completedEvent.get()));
        });

    const HRESULT waitHr = WaitForHandleWithMessagePump(state->completedEvent.get());
    if (FAILED(waitHr))
    {
        return waitHr;
    }

    if (SUCCEEDED(state->hr) && state->result.has_value())
    {
        resultOut = std::move(*state->result);
    }

    return state->hr;
}
} // namespace

HRESULT VerifyWindowsHelloForWindow(HWND ownerWindow, std::wstring_view message) noexcept
{
    if (! ownerWindow || ! IsWindow(ownerWindow))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    // RedSalamander initializes COM as STA on the UI thread; match it.
    winrt::init_apartment(winrt::apartment_type::single_threaded);

    using namespace winrt::Windows::Security::Credentials::UI;

    auto availOp = UserConsentVerifier::CheckAvailabilityAsync();

    UserConsentVerifierAvailability availability = UserConsentVerifierAvailability::DeviceNotPresent;
    const HRESULT availabilityHr                 = WaitForOperationWithMessagePump(availOp, availability);
    if (FAILED(availabilityHr))
    {
        return availabilityHr;
    }

    if (availability != UserConsentVerifierAvailability::Available)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    winrt::hstring messageText(message);

    winrt::com_ptr<IUserConsentVerifierInterop> interop = winrt::get_activation_factory<UserConsentVerifier, IUserConsentVerifierInterop>();

    winrt::Windows::Foundation::IAsyncOperation<UserConsentVerificationResult> op{nullptr};
    const HRESULT hr = interop->RequestVerificationForWindowAsync(
        ownerWindow, reinterpret_cast<HSTRING>(winrt::get_abi(messageText)), winrt::guid_of<decltype(op)>(), winrt::put_abi(op));
    if (FAILED(hr))
    {
        return hr;
    }

    UserConsentVerificationResult result = UserConsentVerificationResult::Canceled;
    const HRESULT verifyHr               = WaitForOperationWithMessagePump(op, result);
    if (FAILED(verifyHr))
    {
        return verifyHr;
    }

    if (result == UserConsentVerificationResult::Verified)
    {
        return S_OK;
    }

    if (result == UserConsentVerificationResult::DeviceNotPresent || result == UserConsentVerificationResult::NotConfiguredForUser ||
        result == UserConsentVerificationResult::DisabledByPolicy)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
}
} // namespace RedSalamander::Security

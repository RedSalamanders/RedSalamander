#pragma once

#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#include "DxUi/DxUi.h"
#include "WindowSizing.h"

namespace Common
{
struct ModalWindowCreateOptions
{
    HINSTANCE instance      = nullptr;
    PCWSTR className        = nullptr;
    PCWSTR caption          = nullptr;
    int clientWidthDip      = 0;
    int clientHeightDip     = 0;
    void* createParameter   = nullptr;
    DWORD style             = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    DWORD extendedStyle     = WS_EX_DLGMODALFRAME;
};

class ModalWindowShell final
{
public:
    explicit ModalWindowShell(HWND ownerWindow) noexcept : _ownerWindow(ownerWindow)
    {
    }

    ~ModalWindowShell() noexcept
    {
        RestoreOwner();
    }

    ModalWindowShell(const ModalWindowShell&)            = delete;
    ModalWindowShell& operator=(const ModalWindowShell&) = delete;
    ModalWindowShell(ModalWindowShell&&)                 = delete;
    ModalWindowShell& operator=(ModalWindowShell&&)      = delete;

    [[nodiscard]] HRESULT CreateCentered(const ModalWindowCreateOptions& options, HWND& createdWindow) noexcept
    {
        createdWindow = nullptr;
        if (! options.instance || ! options.className || ! options.caption || options.clientWidthDip <= 0 || options.clientHeightDip <= 0)
        {
            return E_INVALIDARG;
        }

        const UINT ownerDpi = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
        const UINT dpi      = ownerDpi == 0u ? 96u : ownerDpi;
        const int widthPx   = MulDiv(options.clientWidthDip, static_cast<int>(dpi), 96);
        const int heightPx  = MulDiv(options.clientHeightDip, static_cast<int>(dpi), 96);
        RECT bounds{0, 0, widthPx, heightPx};
        if (AdjustWindowRectExForDpi(&bounds, options.style, FALSE, options.extendedStyle, dpi) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        DisableOwnerIfNeeded();
        const HWND hwnd = CreateWindowExW(options.extendedStyle,
                                          options.className,
                                          options.caption,
                                          options.style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          options.instance,
                                          options.createParameter);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        static_cast<void>(WindowSizing::CenterExistingWindowOnOwner(hwnd, _ownerWindow));
        createdWindow = hwnd;
        return S_OK;
    }

    [[nodiscard]] HRESULT ShowAndRun(HWND hwnd, bool& done, HRESULT& result, std::wstring_view diagnosticName) noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return E_HANDLE;
        }

        ShowWindow(hwnd, SW_SHOWNORMAL);
        UpdateWindow(hwnd);
        SetForegroundWindow(hwnd);

        ModalLoopContext context{&done, &result};
        RedSalamander::DxUi::DxUiModalLoopOptions loopOptions;
        loopOptions.diagnosticName = diagnosticName;
        loopOptions.shouldContinue = ContinueModalLoop;
        loopOptions.context        = &context;
        loopOptions.onQuit         = OnModalLoopQuit;
        const RedSalamander::DxUi::DxUiModalLoopResult loopResult = RedSalamander::DxUi::RunDxUiModalLoop(hwnd, loopOptions);
        if (loopResult == RedSalamander::DxUi::DxUiModalLoopResult::GetMessageFailed)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        return result;
    }

private:
    struct ModalLoopContext
    {
        bool* done      = nullptr;
        HRESULT* result = nullptr;
    };

    [[nodiscard]] static bool ContinueModalLoop(void* rawContext) noexcept
    {
        const auto* context = static_cast<const ModalLoopContext*>(rawContext);
        return context && context->done && ! *context->done;
    }

    static void OnModalLoopQuit(WPARAM, void* rawContext) noexcept
    {
        auto* context = static_cast<ModalLoopContext*>(rawContext);
        if (context && context->done && context->result)
        {
            *context->done   = true;
            *context->result = S_FALSE;
        }
    }

    void DisableOwnerIfNeeded() noexcept
    {
        _restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (_restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
    }

    void RestoreOwner() noexcept
    {
        if (! _restoreOwnerEnabled)
        {
            return;
        }

        _restoreOwnerEnabled = false;
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            EnableWindow(_ownerWindow, TRUE);
            SetActiveWindow(_ownerWindow);
        }
    }

    HWND _ownerWindow         = nullptr;
    bool _restoreOwnerEnabled = false;
};
} // namespace Common

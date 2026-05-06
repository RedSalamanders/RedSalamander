#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <cstdint>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "PlugInterfaces/Viewer.h"

template <typename Derived> class EmbeddedViewerBase : public IViewer
{
public:
    HRESULT STDMETHODCALLTYPE SetCallback(IViewerCallback* callback, void* cookie) noexcept override
    {
        _callbackState.Set(callback, cookie);
        return S_OK;
    }

protected:
    EmbeddedViewerBase()                                                = default;
    ~EmbeddedViewerBase()                                               = default;
    EmbeddedViewerBase(const EmbeddedViewerBase&)                       = delete;
    EmbeddedViewerBase(EmbeddedViewerBase&&)                            = delete;
    EmbeddedViewerBase& operator=(const EmbeddedViewerBase&)            = delete;
    EmbeddedViewerBase& operator=(EmbeddedViewerBase&&)                 = delete;

    [[nodiscard]] static bool IsEmbeddedOpen(const ViewerOpenContext& context) noexcept
    {
        return (static_cast<uint32_t>(context.flags) & static_cast<uint32_t>(VIEWER_OPEN_FLAG_EMBEDDED)) != 0u;
    }

    [[nodiscard]] bool ShouldRecreateViewerWindow(bool embeddedMode, HWND embeddedParent) const noexcept
    {
        return _hWnd && (_embeddedMode != embeddedMode || (embeddedMode && GetParent(_hWnd.get()) != embeddedParent));
    }

    void NotifyViewerClosed() noexcept
    {
        typename RegistrationCallbackState<IViewerCallback>::Snapshot callbackSnapshot{};
        if (! _callbackState.TryCapture(callbackSnapshot))
        {
            return;
        }

        IViewerCallback* callback = nullptr;
        void* cookie              = nullptr;
        if (! _callbackState.TryEnter(callbackSnapshot, callback, cookie))
        {
            return;
        }

        auto* self = static_cast<Derived*>(this);
        self->AddRef();
        auto releaseSelf    = wil::scope_exit([&]() noexcept { self->Release(); });
        auto finishCallback = wil::scope_exit([&]() noexcept { _callbackState.FinishInvoke(); });
        static_cast<void>(callback->ViewerClosed(cookie));
    }

    wil::unique_hwnd _hWnd;
    bool _embeddedMode = false;
    ViewerTheme _theme{};
    bool _hasTheme = false;
    RegistrationCallbackState<IViewerCallback> _callbackState;
};

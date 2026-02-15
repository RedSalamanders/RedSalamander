#include "FolderViewInternal.h"

#include "Ui/AnimationDispatcher.h"

namespace
{
[[nodiscard]] bool IsDisconnectedWin32Error(HRESULT hr) noexcept
{
    if (HRESULT_FACILITY(hr) != FACILITY_WIN32)
    {
        return false;
    }

    const DWORD code = static_cast<DWORD>(HRESULT_CODE(hr));
    switch (code)
    {
        case ERROR_INVALID_DRIVE:
        case ERROR_DEV_NOT_EXIST:
        case ERROR_BAD_NETPATH:
        case ERROR_BAD_NET_NAME:
        case ERROR_BAD_NET_RESP:
        case ERROR_NETNAME_DELETED:
        case ERROR_UNEXP_NET_ERR:
        case ERROR_NETWORK_UNREACHABLE:
        case ERROR_HOST_UNREACHABLE:
        case ERROR_PORT_UNREACHABLE:
        case ERROR_GRACEFUL_DISCONNECT:
        case ERROR_CONNECTION_ABORTED:
        case ERROR_CONNECTION_REFUSED:
        case ERROR_CONNECTION_UNAVAIL:
        case ERROR_NO_NET_OR_BAD_PATH:
        case ERROR_NO_NETWORK:
        case ERROR_NO_SUCH_DEVICE:
        case ERROR_NOT_CONNECTED:
        case ERROR_SEM_TIMEOUT:
        case ERROR_NOT_READY:
        case ERROR_DEVICE_NOT_CONNECTED:
        case ERROR_NO_MEDIA_IN_DRIVE: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsTlsCertificateError(HRESULT hr) noexcept
{
    switch (hr)
    {
        case CERT_E_UNTRUSTEDROOT:
        case CERT_E_CHAINING:
        case CERT_E_EXPIRED:
        case CERT_E_REVOKED:
        case CERT_E_CN_NO_MATCH:
        case SEC_E_CERT_UNKNOWN:
        case SEC_E_UNTRUSTED_ROOT:
        case SEC_E_ILLEGAL_MESSAGE: return true;
        default: return false;
    }
}
} // namespace

void FolderView::StartOverlayTimer(UINT intervalMs) const
{
    if (! _hWnd)
    {
        return;
    }

    intervalMs = std::max(intervalMs, 1u);
    if (_overlayTimer != 0)
    {
        if (intervalMs >= _overlayTimerIntervalMs)
        {
            return;
        }
    }

    const UINT_PTR timer = SetTimer(_hWnd.get(), kOverlayTimerId, intervalMs, nullptr);
    if (timer != 0)
    {
        _overlayTimer           = timer;
        _overlayTimerIntervalMs = intervalMs;
    }
}

void FolderView::StopOverlayTimer() const
{
    if (! _hWnd || _overlayTimer == 0)
    {
        return;
    }

    KillTimer(_hWnd.get(), kOverlayTimerId);
    _overlayTimer           = 0;
    _overlayTimerIntervalMs = 0;
}

void FolderView::StartOverlayAnimation() const noexcept
{
    if (_overlayAnimationSubscriptionId != 0)
    {
        return;
    }

    _overlayAnimationSubscriptionId = RedSalamander::Ui::AnimationDispatcher::GetInstance().Subscribe(
        [](void* context, uint64_t nowTickMs) noexcept -> bool
        {
            auto* self = static_cast<FolderView*>(context);
            if (! self)
            {
                return false;
            }

            return self->OnOverlayAnimationTick(nowTickMs);
        },
        const_cast<FolderView*>(this));
}

void FolderView::StopOverlayAnimation() const noexcept
{
    if (_overlayAnimationSubscriptionId == 0)
    {
        return;
    }

    RedSalamander::Ui::AnimationDispatcher::GetInstance().Unsubscribe(_overlayAnimationSubscriptionId);
    _overlayAnimationSubscriptionId = 0;
}

bool FolderView::UpdateIncrementalSearchIndicatorAnimation(uint64_t nowTickMs) const noexcept
{
    constexpr uint64_t kVisibilityAnimationMs = 220;
    constexpr uint64_t kPulseAnimationMs      = 260;

    bool needsAnimation = false;

    auto easeInCubic  = [](float t) noexcept -> float { return t * t * t; };
    auto easeOutCubic = [](float t) noexcept -> float
    {
        const float inv = 1.0f - t;
        return 1.0f - inv * inv * inv;
    };

    if (_incrementalSearchIndicatorVisibilityStart != 0 && kVisibilityAnimationMs > 0)
    {
        const uint64_t elapsed = nowTickMs >= _incrementalSearchIndicatorVisibilityStart ? (nowTickMs - _incrementalSearchIndicatorVisibilityStart) : 0;
        const float t          = std::clamp(static_cast<float>(elapsed) / static_cast<float>(kVisibilityAnimationMs), 0.0f, 1.0f);
        const float eased      = _incrementalSearchIndicatorVisibilityTo >= _incrementalSearchIndicatorVisibilityFrom ? easeOutCubic(t) : easeInCubic(t);

        _incrementalSearchIndicatorVisibility =
            _incrementalSearchIndicatorVisibilityFrom + (_incrementalSearchIndicatorVisibilityTo - _incrementalSearchIndicatorVisibilityFrom) * eased;

        if (elapsed < kVisibilityAnimationMs)
        {
            needsAnimation = true;
        }
        else
        {
            _incrementalSearchIndicatorVisibility = _incrementalSearchIndicatorVisibilityTo;
        }
    }
    else
    {
        _incrementalSearchIndicatorVisibility = _incrementalSearchIndicatorVisibilityTo;
    }

    if (_incrementalSearchIndicatorTypingPulseStart != 0)
    {
        const uint64_t elapsed = nowTickMs >= _incrementalSearchIndicatorTypingPulseStart ? (nowTickMs - _incrementalSearchIndicatorTypingPulseStart) : 0;
        if (elapsed < kPulseAnimationMs)
        {
            needsAnimation = true;
        }
        else
        {
            _incrementalSearchIndicatorTypingPulseStart = 0;
        }
    }

    if (_incrementalSearch.active && _incrementalSearchIndicatorVisibilityTo > 0.0f)
    {
        needsAnimation = true;
    }

    return needsAnimation;
}

bool FolderView::OnOverlayAnimationTick(uint64_t nowTickMs) const noexcept
{
    if (_overlayAnimationSubscriptionId == 0 || ! _hWnd)
    {
        StopOverlayAnimation();
        return false;
    }

    ErrorOverlayState overlay{};
    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay)
        {
            overlay    = *_errorOverlay;
            hasOverlay = true;
        }
    }

    constexpr uint64_t kShowAnimationMs = 220;
    const uint64_t elapsed              = nowTickMs >= overlay.startTick ? (nowTickMs - overlay.startTick) : 0;
    const bool overlayNeedsAnimation    = hasOverlay && (overlay.severity == OverlaySeverity::Busy || elapsed < kShowAnimationMs);
    const bool indicatorNeedsAnimation  = UpdateIncrementalSearchIndicatorAnimation(nowTickMs);
    const bool needsAnimation           = overlayNeedsAnimation || indicatorNeedsAnimation;

    if (! needsAnimation)
    {
        StopOverlayAnimation();
        return false;
    }

    InvalidateRect(_hWnd.get(), nullptr, FALSE);
    return true;
}

void FolderView::ScheduleBusyOverlay(uint64_t generation, const std::filesystem::path& folder)
{
    if (! _hWnd)
    {
        return;
    }

    PendingBusyOverlay pending;
    pending.generation = generation;
    pending.folder     = folder;
    pending.startTick  = GetTickCount64();

    _pendingBusyOverlay = std::move(pending);
    StartOverlayTimer(static_cast<UINT>(kBusyOverlayDelayMs));
}

void FolderView::CancelBusyOverlay(uint64_t generation)
{
    if (! _pendingBusyOverlay || _pendingBusyOverlay->generation != generation)
    {
        return;
    }

    _pendingBusyOverlay.reset();

    StopOverlayTimer();

    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        hasOverlay = _errorOverlay.has_value();
    }

    if (! hasOverlay)
    {
        const uint64_t nowTickMs = GetTickCount64();
        if (! UpdateIncrementalSearchIndicatorAnimation(nowTickMs))
        {
            StopOverlayAnimation();
        }
    }
}

void FolderView::ShowBusyOverlayNow(const std::filesystem::path& folder)
{
    ErrorOverlayState overlay{};
    overlay.kind        = ErrorOverlayKind::Enumeration;
    overlay.severity    = OverlaySeverity::Busy;
    overlay.hr          = S_OK;
    overlay.startTick   = GetTickCount64();
    overlay.closable    = false;
    overlay.blocksInput = true;

    overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_PLEASE_WAIT);
    overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_ACCESSING_FOLDER_FMT, folder.native());

    if (overlay.message.empty())
    {
        overlay.message = folder.native();
    }

    bool changed = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (! _errorOverlay || _errorOverlay->kind != overlay.kind || _errorOverlay->severity != overlay.severity || _errorOverlay->title != overlay.title ||
            _errorOverlay->message != overlay.message || _errorOverlay->closable != overlay.closable || _errorOverlay->blocksInput != overlay.blocksInput)
        {
            _errorOverlay = overlay;
            changed       = true;
        }
    }

    if (changed && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    StartOverlayAnimation();
}

void FolderView::OnTimerMessage(UINT_PTR timerId)
{
    // Handle idle layout pre-creation timer
    if (timerId == kIdleLayoutTimerId)
    {
        ProcessIdleLayoutBatch();
        return;
    }

    if (timerId != kOverlayTimerId)
    {
        return;
    }

    if (! _pendingBusyOverlay)
    {
        StopOverlayTimer();
        return;
    }

    const uint64_t now = GetTickCount64();

    const uint64_t currentGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    if (_pendingBusyOverlay->generation != currentGeneration)
    {
        _pendingBusyOverlay.reset();
        StopOverlayTimer();
        return;
    }

    const uint64_t dueTick = _pendingBusyOverlay->startTick + kBusyOverlayDelayMs;
    if (now >= dueTick)
    {
        bool canShow = false;
        {
            std::lock_guard lock(_errorOverlayMutex);
            canShow = ! _errorOverlay.has_value();
        }

        if (canShow)
        {
            ShowBusyOverlayNow(_pendingBusyOverlay->folder);
        }

        _pendingBusyOverlay.reset();
        StopOverlayTimer();
        return;
    }

    const uint64_t remaining     = dueTick > now ? (dueTick - now) : 1u;
    const UINT pendingIntervalMs = static_cast<UINT>(std::clamp<uint64_t>(remaining, 1u, 1000u));
    StartOverlayTimer(std::max(pendingIntervalMs, 1u));
}

void FolderView::ReportError(const std::wstring& context, HRESULT hr) const
{
    std::wstring hrText = FormatHResult(hr);
    while (! hrText.empty() && (hrText.back() == L'\r' || hrText.back() == L'\n'))
    {
        hrText.pop_back();
    }

    std::wstring details = FormatStringResource(nullptr, IDS_FMT_HRESULT_DETAILS, static_cast<unsigned>(hr), hrText);
    if (details.empty())
    {
        details = std::format(L"0x{:08X}: {}", static_cast<unsigned long>(hr), hrText);
    }
    Debug::Error(L"{} failed: {}", context, details);

    ErrorOverlayState overlay{};
    overlay.hr        = hr;
    overlay.severity  = OverlaySeverity::Error;
    overlay.startTick = GetTickCount64();

    if (context == L"EnumerateFolder")
    {
        overlay.kind = ErrorOverlayKind::Enumeration;

        if (hr == HRESULT_FROM_WIN32(ERROR_DLL_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND))
        {
            overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_FS_PLUGIN_NOT_AVAILABLE);
            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_FS_PLUGIN_NOT_AVAILABLE_FMT, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else if (IsDisconnectedWin32Error(hr))
        {
            overlay.severity    = OverlaySeverity::Information;
            overlay.closable    = false;
            overlay.blocksInput = true;
            overlay.title       = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_DISCONNECTED);

            std::wstring_view folderText;
            if (_currentFolder.has_value())
            {
                folderText = _currentFolder.value().native();
            }
            if (folderText.empty())
            {
                folderText = details;
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_DISCONNECTED_FMT, folderText);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else if (hr == HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD))
        {
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_LOGIN_FAILED);

            std::wstring_view folderText;
            if (_currentFolder.has_value())
            {
                folderText = _currentFolder.value().native();
            }
            if (folderText.empty())
            {
                folderText = details;
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_INVALID_PASSWORD_FMT, folderText, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else if (hr == HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE))
        {
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_LOGIN_FAILED);

            std::wstring_view folderText;
            if (_currentFolder.has_value())
            {
                folderText = _currentFolder.value().native();
            }
            if (folderText.empty())
            {
                folderText = details;
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_LOGIN_FAILED_FMT, folderText, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else if (IsTlsCertificateError(hr))
        {
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_TLS_CERTIFICATE_FAILED);

            std::wstring_view folderText;
            if (_currentFolder.has_value())
            {
                folderText = _currentFolder.value().native();
            }
            if (folderText.empty())
            {
                folderText = details;
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_TLS_CERTIFICATE_FAILED_FMT, folderText, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else if (hr == E_ACCESSDENIED || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
        {
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_ACCESS_DENIED);

            std::wstring_view folderText;
            if (_currentFolder.has_value())
            {
                folderText = _currentFolder.value().native();
            }
            if (folderText.empty())
            {
                folderText = details;
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_ACCESS_DENIED_FMT, folderText, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
        }
        else
        {
            overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_ENUMERATION_FAILED);
            overlay.message = details;
        }
    }
    else if (context.find(L"IDXGI") != std::wstring::npos || context.find(L"ID2D1") != std::wstring::npos || context.find(L"D3D") != std::wstring::npos)
    {
        overlay.kind    = ErrorOverlayKind::Rendering;
        overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_RENDERING_ERROR);
        overlay.message = details;
    }
    else
    {
        overlay.kind    = ErrorOverlayKind::Operation;
        overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_OPERATION_FAILED);
        overlay.message = details;
    }

    bool changed = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (! _errorOverlay || _errorOverlay->kind != overlay.kind || _errorOverlay->severity != overlay.severity || _errorOverlay->hr != overlay.hr ||
            _errorOverlay->title != overlay.title || _errorOverlay->message != overlay.message || _errorOverlay->closable != overlay.closable ||
            _errorOverlay->blocksInput != overlay.blocksInput)
        {
            _errorOverlay = overlay;
            changed       = true;
        }
    }

    if (changed && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (changed)
    {
        StartOverlayAnimation();
    }
}

void FolderView::ClearErrorOverlay(ErrorOverlayKind kind) const
{
    bool cleared = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay && _errorOverlay->kind == kind)
        {
            _errorOverlay.reset();
            cleared = true;
        }
    }

    if (cleared && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    if (cleared)
    {
        const uint64_t nowTickMs = GetTickCount64();
        if (! UpdateIncrementalSearchIndicatorAnimation(nowTickMs))
        {
            StopOverlayAnimation();
        }
        StopOverlayTimer();
    }
}

void FolderView::ShowAlertOverlay(
    ErrorOverlayKind kind, OverlaySeverity severity, std::wstring title, std::wstring message, HRESULT hr, bool closable, bool blocksInput)
{
    ErrorOverlayState overlay{};
    overlay.kind        = kind;
    overlay.severity    = severity;
    overlay.title       = std::move(title);
    overlay.message     = std::move(message);
    overlay.hr          = hr;
    overlay.startTick   = GetTickCount64();
    overlay.closable    = closable;
    overlay.blocksInput = blocksInput;

    bool changed = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (! _errorOverlay || _errorOverlay->kind != overlay.kind || _errorOverlay->severity != overlay.severity || _errorOverlay->hr != overlay.hr ||
            _errorOverlay->title != overlay.title || _errorOverlay->message != overlay.message || _errorOverlay->closable != overlay.closable ||
            _errorOverlay->blocksInput != overlay.blocksInput)
        {
            _errorOverlay = overlay;
            changed       = true;
        }
    }

    if (changed && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (changed)
    {
        StartOverlayAnimation();
    }
}

void FolderView::DismissAlertOverlay()
{
    bool cleared = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay)
        {
            _errorOverlay.reset();
            cleared = true;
        }
    }

    if (cleared && _alertOverlay)
    {
        _alertOverlay->ClearHotState();
    }

    if (cleared && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (cleared)
    {
        if (_pendingBusyOverlay)
        {
            const uint64_t now       = GetTickCount64();
            const uint64_t dueTick   = _pendingBusyOverlay->startTick + kBusyOverlayDelayMs;
            const uint64_t remaining = dueTick > now ? (dueTick - now) : 1u;
            if (! UpdateIncrementalSearchIndicatorAnimation(now))
            {
                StopOverlayAnimation();
            }
            StartOverlayTimer(static_cast<UINT>(std::clamp<uint64_t>(remaining, 1u, 1000u)));
        }
        else
        {
            const uint64_t nowTickMs = GetTickCount64();
            if (! UpdateIncrementalSearchIndicatorAnimation(nowTickMs))
            {
                StopOverlayAnimation();
            }
            StopOverlayTimer();
        }
    }
}

void FolderView::DrawErrorOverlay()
{
    ErrorOverlayState overlay{};
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (! _errorOverlay)
        {
            if (_alertOverlay)
            {
                _alertOverlay->ClearHotState();
            }
            return;
        }
        overlay = *_errorOverlay;
    }

    if (! _d2dContext || ! _dwriteFactory || ! _alertOverlay)
    {
        return;
    }

    const float clientWidthDip  = DipFromPx(_clientSize.cx);
    const float clientHeightDip = DipFromPx(_clientSize.cy);
    if (clientWidthDip <= 0.0f || clientHeightDip <= 0.0f)
    {
        return;
    }

    RedSalamander::Ui::AlertTheme alertTheme{};
    alertTheme.background          = _theme.backgroundColor;
    alertTheme.text                = _theme.textNormal;
    alertTheme.accent              = _theme.focusBorder;
    alertTheme.selectionBackground = _theme.itemBackgroundSelected;
    alertTheme.selectionText       = _theme.textSelected;
    alertTheme.errorBackground     = _theme.errorBackground;
    alertTheme.errorText           = _theme.errorText;
    alertTheme.warningBackground   = _theme.warningBackground;
    alertTheme.warningText         = _theme.warningText;
    alertTheme.infoBackground      = _theme.infoBackground;
    alertTheme.infoText            = _theme.infoText;
    alertTheme.darkBase            = _theme.darkBase;
    _alertOverlay->SetTheme(alertTheme);

    RedSalamander::Ui::AlertModel model{};
    switch (overlay.severity)
    {
        case OverlaySeverity::Error: model.severity = RedSalamander::Ui::AlertSeverity::Error; break;
        case OverlaySeverity::Warning: model.severity = RedSalamander::Ui::AlertSeverity::Warning; break;
        case OverlaySeverity::Information: model.severity = RedSalamander::Ui::AlertSeverity::Info; break;
        case OverlaySeverity::Busy: model.severity = RedSalamander::Ui::AlertSeverity::Busy; break;
    }
    model.title    = overlay.title;
    model.message  = overlay.message;
    model.closable = overlay.closable;

    constexpr uint32_t kCancelButtonId = 1;
    if (overlay.severity == OverlaySeverity::Busy && overlay.kind == ErrorOverlayKind::Enumeration)
    {
        model.closable = false;

        RedSalamander::Ui::AlertButton cancel{};
        cancel.id      = kCancelButtonId;
        cancel.label   = LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
        cancel.primary = true;
        if (! cancel.label.empty())
        {
            model.buttons.emplace_back(std::move(cancel));
        }
    }

    const RedSalamander::Ui::AlertModel& current = _alertOverlay->GetModel();
    bool needsModelUpdate                        = current.severity != model.severity || current.title != model.title || current.message != model.message ||
                            current.closable != model.closable || current.buttons.size() != model.buttons.size();
    if (! needsModelUpdate)
    {
        for (size_t i = 0; i < model.buttons.size(); ++i)
        {
            const RedSalamander::Ui::AlertButton& expected = model.buttons[i];
            const RedSalamander::Ui::AlertButton& existing = current.buttons[i];
            if (expected.id != existing.id || expected.label != existing.label || expected.primary != existing.primary)
            {
                needsModelUpdate = true;
                break;
            }
        }
    }

    if (needsModelUpdate)
    {
        _alertOverlay->SetModel(std::move(model));
    }

    _alertOverlay->SetStartTick(overlay.startTick);
    const uint64_t nowTick = GetTickCount64();
    _alertOverlay->Draw(_d2dContext.get(), _dwriteFactory.get(), clientWidthDip, clientHeightDip, nowTick);
}

bool FolderView::CheckHR(HRESULT hr, const wchar_t* context) const
{
    if (FAILED(hr))
    {
        const std::wstring message = context ? std::wstring(context) : std::wstring(L"FolderView operation");
        ReportError(message, hr);
        return false;
    }
    return true;
}

void FolderView::DebugShowOverlaySample(OverlaySeverity severity)
{
    DebugShowOverlaySample(ErrorOverlayKind::Operation, severity, true);
}

void FolderView::DebugShowOverlaySample(ErrorOverlayKind kind, OverlaySeverity severity, bool blocksInput)
{
    if (! _hWnd)
    {
        return;
    }

    ErrorOverlayState overlay{};
    overlay.kind        = kind;
    overlay.severity    = severity;
    overlay.startTick   = GetTickCount64();
    overlay.closable    = severity != OverlaySeverity::Busy;
    overlay.blocksInput = blocksInput;

    switch (severity)
    {
        case OverlaySeverity::Error:
        {
            overlay.hr    = E_FAIL;
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_OPERATION_FAILED);

            const std::wstring hrText  = FormatHResult(overlay.hr);
            const std::wstring details = FormatStringResource(nullptr, IDS_FMT_HRESULT_DETAILS, static_cast<unsigned>(overlay.hr), hrText);
            overlay.message            = FormatStringResource(nullptr, IDS_OVERLAY_DEBUG_SAMPLE_MSG_ERROR_FMT, details);
            if (overlay.message.empty())
            {
                overlay.message = details;
            }
            break;
        }
        case OverlaySeverity::Warning:
            overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_WARNING);
            overlay.message = LoadStringResource(nullptr, IDS_OVERLAY_DEBUG_SAMPLE_MSG_WARNING);
            break;
        case OverlaySeverity::Information:
            overlay.title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
            overlay.message = LoadStringResource(nullptr, IDS_OVERLAY_DEBUG_SAMPLE_MSG_INFORMATION);
            break;
        case OverlaySeverity::Busy:
        {
            overlay.title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_PLEASE_WAIT);

            std::filesystem::path folder;
            if (_currentFolder)
            {
                folder = *_currentFolder;
            }
            else
            {
                const std::wstring sampleFolder = LoadStringResource(nullptr, IDS_OVERLAY_DEBUG_SAMPLE_FOLDER_PATH);
                if (! sampleFolder.empty())
                {
                    folder = std::filesystem::path(sampleFolder);
                }
            }

            overlay.message = FormatStringResource(nullptr, IDS_OVERLAY_MSG_ACCESSING_FOLDER_FMT, folder.native());
            if (overlay.message.empty())
            {
                overlay.message = folder.native();
            }
            break;
        }
    }

    bool changed = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (! _errorOverlay || _errorOverlay->kind != overlay.kind || _errorOverlay->severity != overlay.severity || _errorOverlay->title != overlay.title ||
            _errorOverlay->message != overlay.message || _errorOverlay->closable != overlay.closable || _errorOverlay->blocksInput != overlay.blocksInput)
        {
            _errorOverlay = overlay;
            changed       = true;
        }
    }

    if (changed)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        StartOverlayAnimation();
    }
}

void FolderView::DebugShowCanceledOverlaySample()
{
    if (! _hWnd)
    {
        return;
    }

    std::wstring title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_CANCELED);
    std::wstring message = LoadStringResource(nullptr, IDS_OVERLAY_MSG_ENUMERATION_CANCELED);
    ShowAlertOverlay(
        ErrorOverlayKind::Enumeration, OverlaySeverity::Information, std::move(title), std::move(message), HRESULT_FROM_WIN32(ERROR_CANCELLED), false, false);
}

void FolderView::DebugHideOverlaySample()
{
    if (! _hWnd)
    {
        return;
    }

    DismissAlertOverlay();
}

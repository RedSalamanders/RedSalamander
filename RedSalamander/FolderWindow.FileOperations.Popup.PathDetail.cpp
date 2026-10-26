#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "Resource.h"
#include "UnicodeClipboard.h"

#include <algorithm>
#include <array>
#include <chrono>

namespace
{
[[nodiscard]] uint64_t PerfNowUs() noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
}
[[nodiscard]] uint64_t PerfElapsedUs(uint64_t start) noexcept
{
    const uint64_t now = PerfNowUs();
    return now >= start ? now - start : 0u;
}
[[nodiscard]] bool CanRestorePathDetailTextFocus(HWND popup, HWND focus, HWND foreground) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
        return false;
    const bool popupOwnsFocus      = focus == popup || (focus && IsChild(popup, focus) != FALSE);
    const bool popupOwnsForeground = foreground == popup || (foreground && IsChild(popup, foreground) != FALSE);
    return popupOwnsFocus && popupOwnsForeground;
}
// Consumer-owned read-only path workspace. It shares the popup HWND and host;
// opening it never activates another window or grants file-operation authority.
class FileOperationsPathDetailPanel final : public DxUi::Panel
{
public:
    FileOperationsPathDetailPanel()                                                = default;
    FileOperationsPathDetailPanel(const FileOperationsPathDetailPanel&)            = delete;
    FileOperationsPathDetailPanel(FileOperationsPathDetailPanel&&)                 = delete;
    FileOperationsPathDetailPanel& operator=(const FileOperationsPathDetailPanel&) = delete;
    FileOperationsPathDetailPanel& operator=(FileOperationsPathDetailPanel&&)      = delete;

    void Paint(DxUi::WindowHost& host) const override
    {
        if (auto* context = host.GetDeviceContext())
            context->FillRectangle(GetBounds(), host.GetSolidBrush(host.GetTheme().surfaceBackground));
        DxUi::Panel::Paint(host);
    }
};

} // namespace

void FileOperationsPopupInternal::FileOperationsPopupState::OpenPathDetail(HWND hwnd, std::wstring_view automationId) noexcept
{
    const auto descriptor = std::ranges::find(_hostedTooltipDescriptors, automationId, &HostedTooltipDescriptor::automationId);
    if (! hostLifetime.lock() || descriptor == _hostedTooltipDescriptors.end() || descriptor->copyPath.empty())
        return;
    _pathDetailId             = automationId;
    _pathDetailReturnId       = automationId;
    _pathDetailTaskId         = descriptor->taskId;
    _pathDetailNeedsFocus     = true;
    _pathDetailLogicalFocusId = L"FileOperations.PathDetail.Text";
    static_cast<void>(_controlsHost.ClearTooltip());
    _controlsHost.SetDefaultButton(nullptr);
    _controlsHost.SetCancelButton(nullptr);
    Invalidate(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::ClosePathDetail(HWND hwnd, bool restoreFocus) noexcept
{
    if (_pathDetailId.empty() && ! _pathDetailField)
        return;
    _pathDetailId.clear();
    _pathDetailNeedsFocus = false;
    _pathDetailLogicalFocusId.clear();
    _controlsHost.ResetInteractionState();
    _controlsHost.SetDefaultButton(nullptr);
    _controlsHost.SetCancelButton(nullptr);
    // The binding synchronized while detail was open was global; the next sync
    // re-derives it from the restored focus before Escape can reach another task.
    _conflictEscapeTaskId     = 0u;
    _conflictEscapeAction     = 0u;
    _conflictEscapeGeneration = 0u;
    // A control retired during a mouse capture stays hidden until it is removed.
    if (_controlsRoot)
        for (const auto& child : _controlsRoot->GetChildren())
            child->SetVisible(child.get() != _pathDetailPanel &&
                              std::ranges::none_of(_retiredHostedControls,
                                                   [&child](const RetiredHostedControl& retired) noexcept { return retired.control == child.get(); }));
    if (restoreFocus)
    {
        const auto row = std::ranges::find_if(_hostedTooltipRegions, [&](const DxUi::Control* region) noexcept {
            return region->IsFocusable() && region->GetAccessibleAutomationId() == _pathDetailReturnId;
        });
        // A vanished row keeps its task identity through that task's disclosure, as a
        // rebuild does; global Options only when the task is gone, so Escape never
        // binds to another task's sole decision.
        const auto focusHostedButton = [&](PopupHitTest::Kind kind, uint64_t taskId) noexcept
        {
            for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
                if (_hostedButtonHits[index].kind == kind && _hostedButtonHits[index].taskId == taskId && _hostedButtons[index] &&
                    _hostedButtons[index]->IsFocusable())
                {
                    _controlsHost.SetFocusControl(_hostedButtons[index], false);
                    return true;
                }
            return false;
        };
        if (row != _hostedTooltipRegions.end())
            _controlsHost.SetFocusControl(*row, false);
        else if (_pathDetailTaskId == 0u || ! focusHostedButton(PopupHitTest::Kind::TaskToggleCollapse, _pathDetailTaskId))
            static_cast<void>(focusHostedButton(PopupHitTest::Kind::FooterOptions, 0u));
    }
    _pathDetailReturnId.clear();
    _pathDetailTaskId = 0u;
    // Child destruction is deferred to synchronization: this may run inside the
    // close button's callback, so deleting that button here would invalidate it.
    Invalidate(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::SyncPathDetail(HWND hwnd) noexcept
{
    auto descriptor = _hostedTooltipDescriptors.end();
    if (! _pathDetailId.empty())
        descriptor = std::ranges::find(_hostedTooltipDescriptors, _pathDetailId, &HostedTooltipDescriptor::automationId);
    if (! _pathDetailId.empty() && (descriptor == _hostedTooltipDescriptors.end() || descriptor->copyPath.empty()))
        ClosePathDetail(hwnd, true);
    if (_pathDetailId.empty())
    {
        if (_pathDetailPanel && _pathDetailField)
        {
            _pathDetailPanel->ClearChildren();
            _pathDetailTitle = nullptr;
            _pathDetailField = nullptr;
            _pathDetailCopy  = nullptr;
            _pathDetailClose = nullptr;
            if (Debug::Perf::IsCaptureEnabled())
                Debug::Perf::Emit(L"FileOps.Popup.PathDetail.Released", L"", 0u, 0u, 0u, S_OK);
        }
        return;
    }
    const uint64_t startedUs = Debug::Perf::IsCaptureEnabled() ? PerfNowUs() : 0u;
    const bool created       = _pathDetailField == nullptr;
    if (! _pathDetailPanel)
        _pathDetailPanel = _controlsRoot->AddChild<FileOperationsPathDetailPanel>();
    if (created)
    {
        _pathDetailPanel->SetAccessibleAutomationId(L"FileOperations.PathDetail");
        _pathDetailTitle = _pathDetailPanel->AddChild<DxUi::Label>();
        _pathDetailTitle->SetFontRole(DxUi::FontRole::Subtitle);
        _pathDetailField = _pathDetailPanel->AddChild<DxUi::TextField>();
        _pathDetailField->SetAccessibleAutomationId(L"FileOperations.PathDetail.Text");
        _pathDetailField->SetMultiline(true);
        _pathDetailField->SetReadOnly(true);
        _pathDetailField->SetClearButtonEnabled(false);
        _pathDetailField->SetOnPreviewKeyDown([](DxUi::WindowHost&, UINT key, UINT) { return key == VK_RETURN || key == VK_SPACE; });
        _pathDetailCopy = _pathDetailPanel->AddChild<DxUi::Button>();
        _pathDetailCopy->SetAccessibleAutomationId(L"FileOperations.PathDetail.Copy");
        _pathDetailCopy->SetOnClick([this, hwnd]
        {
            if (_pathDetailField && ! _pathDetailId.empty())
                static_cast<void>(Common::Clipboard::TrySetUnicodeText(hwnd, _pathDetailField->GetText(), Common::Clipboard::EmptyUnicodeTextPolicy::Reject));
        });
        _pathDetailClose = _pathDetailPanel->AddChild<DxUi::Button>();
        _pathDetailClose->SetAccessibleAutomationId(L"FileOperations.PathDetail.Close");
        _pathDetailClose->SetOnClick([this, hwnd] { ClosePathDetail(hwnd, true); });
    }
    if (_pathDetailField->GetText() != descriptor->copyPath)
    {
        _pathDetailField->SetText(descriptor->copyPath);
        _pathDetailField->SetSelectionRange(descriptor->copyPath.size(), 0u);
    }
    const auto title = LoadStringResource(nullptr, IDS_ITEM_PROPERTIES_FIELD_FULL_PATH);
    _pathDetailTitle->SetText(title);
    _pathDetailField->SetAccessibleName(title);
    _pathDetailField->SetAccessibleHelpText(LoadStringResource(nullptr, IDS_FILEOPS_PATH_DETAIL_HELP));
    _pathDetailCopy->SetText(LoadStringResource(nullptr, IDS_FILEOPS_PATH_DETAIL_COPY));
    _pathDetailClose->SetText(LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLOSE));
    const auto bounds = _controlsHost.GetClientBoundsDip();
    _pathDetailPanel->SetBounds(bounds);
    const float left   = bounds.left + 12.0f;
    const float right  = bounds.right - 12.0f;
    const float bottom = bounds.bottom - 12.0f;
    _pathDetailTitle->SetBounds(D2D1::RectF(left, bounds.top + 12.0f, right, bounds.top + 44.0f));
    _pathDetailField->SetBounds(D2D1::RectF(left, bounds.top + 52.0f, right, std::max(bounds.top + 84.0f, bottom - 88.0f)));
    _pathDetailCopy->SetBounds(D2D1::RectF(left, bottom - 80.0f, right, bottom - 44.0f));
    _pathDetailClose->SetBounds(D2D1::RectF(left, bottom - 36.0f, right, bottom));
    for (const auto& child : _controlsRoot->GetChildren())
        child->SetVisible(child.get() == _pathDetailPanel);
    _controlsHost.SetDefaultButton(nullptr);
    _controlsHost.SetCancelButton(_pathDetailClose);
    if (created || _pathDetailNeedsFocus)
    {
        DxUi::Control* target = _pathDetailField;
        for (auto* control : std::array<DxUi::Control*, 3>{{_pathDetailField, _pathDetailCopy, _pathDetailClose}})
            if (control->GetAccessibleAutomationId() == _pathDetailLogicalFocusId)
                target = control;
        // The pinned host's false flag suppresses native focus only for non-text
        // controls. Text restoration requires actual foreground and focus ownership.
        const bool mayFocusText = CanRestorePathDetailTextFocus(hwnd, GetFocus(), GetForegroundWindow());
        if (target != _pathDetailField || mayFocusText)
        {
            _controlsHost.SetFocusControl(target, false);
            _pathDetailNeedsFocus = false;
        }
        else
            _pathDetailNeedsFocus = true;
    }
    if (Debug::Perf::IsCaptureEnabled())
        Debug::Perf::Emit(L"FileOps.Popup.PathDetail.SyncUs", created ? L"open" : L"update", PerfElapsedUs(startedUs), descriptor->copyPath.size(), 4u, S_OK);
}

#ifdef ENABLE_TESTS
bool DebugCanRestoreFileOperationsPathDetailTextFocus(HWND popup, HWND focus, HWND foreground) noexcept
{
    return CanRestorePathDetailTextFocus(popup, focus, foreground);
}
#endif

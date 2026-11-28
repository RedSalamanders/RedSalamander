#pragma once

#include "LocalFileTransaction.h"

#include <atomic>
#include <memory>
#include <shobjidl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::Files
{
struct GuardedSaveSelection final
{
    std::filesystem::path path;
    std::optional<LocalFileVersion> acceptedVersion;
};

namespace Detail
{
class SaveDestinationEvents final : public IFileDialogEvents
{
public:
    SaveDestinationEvents()                                        = default;
    SaveDestinationEvents(const SaveDestinationEvents&)            = delete;
    SaveDestinationEvents& operator=(const SaveDestinationEvents&) = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) noexcept override
    {
        if (! result)
        {
            return E_POINTER;
        }
        *result = nullptr;
        if (iid != __uuidof(IUnknown) && iid != __uuidof(IFileDialogEvents))
        {
            return E_NOINTERFACE;
        }
        *result = static_cast<IFileDialogEvents*>(this);
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _references.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _references.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            const std::unique_ptr<SaveDestinationEvents> finalOwner(this);
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE OnFileOk(IFileDialog*) noexcept override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnFolderChanging(IFileDialog*, IShellItem*) noexcept override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnFolderChange(IFileDialog*) noexcept override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnSelectionChange(IFileDialog*) noexcept override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnTypeChange(IFileDialog*) noexcept override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE*) noexcept override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE OnOverwrite(IFileDialog*, IShellItem* item, FDE_OVERWRITE_RESPONSE* response) noexcept override
    {
        if (! item || ! response)
        {
            return E_POINTER;
        }
        _acceptedPath.clear();
        _acceptedVersion.reset();
        *response = FDEOR_REFUSE;
        wil::unique_cotaskmem_string path;
        const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
        if (FAILED(nameHr) || ! path)
        {
            return S_OK;
        }

        // Bind the version presented by the shell before its normal overwrite prompt.
        if (FAILED(ObserveLocalFileDestination(std::filesystem::path(path.get()), _acceptedVersion)))
        {
            return S_OK;
        }
        _acceptedPath = path.get();
        *response     = FDEOR_DEFAULT;
        return S_OK;
    }

    [[nodiscard]] std::optional<LocalFileVersion> VersionFor(const std::filesystem::path& path) const noexcept
    {
        if (_acceptedPath.empty() || ! Common::Paths::NormalizedWindowsPathEqualsNoCase(path.native(), _acceptedPath))
        {
            return std::nullopt;
        }
        return _acceptedVersion;
    }

private:
    friend struct std::default_delete<SaveDestinationEvents>;
    ~SaveDestinationEvents() = default;

    std::atomic<ULONG> _references{1u};
    std::wstring _acceptedPath;
    std::optional<LocalFileVersion> _acceptedVersion;
};
} // namespace Detail

// The caller configures the existing shell dialog and keeps FOS_OVERWRITEPROMPT enabled.
[[nodiscard]] inline HRESULT ShowGuardedSaveDialog(IFileSaveDialog& dialog, HWND owner, GuardedSaveSelection& selection) noexcept
{
    selection = {};
    wil::com_ptr_nothrow<Detail::SaveDestinationEvents> events;
    events.attach(std::make_unique<Detail::SaveDestinationEvents>().release());
    DWORD cookie           = 0;
    const HRESULT adviseHr = dialog.Advise(events.get(), &cookie);
    if (FAILED(adviseHr))
    {
        return adviseHr;
    }
    const auto unadvise = wil::scope_exit([&]() noexcept { static_cast<void>(dialog.Unadvise(cookie)); });

    const HRESULT showHr = dialog.Show(owner);
    if (FAILED(showHr))
    {
        return showHr;
    }
    wil::com_ptr_nothrow<IShellItem> item;
    const HRESULT itemHr = dialog.GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return FAILED(itemHr) ? itemHr : E_UNEXPECTED;
    }
    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path || path.get()[0] == L'\0')
    {
        return FAILED(nameHr) ? nameHr : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }
    selection.path            = path.get();
    selection.acceptedVersion = events->VersionFor(selection.path);
    return S_OK;
}
} // namespace Common::Files

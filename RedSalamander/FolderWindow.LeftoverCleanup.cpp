#include "FolderWindowInternal.h"
#include "WindowActivation.h"

#include "ChangeCase.h"
#include "DxUiThemePalette.h"
#include "FileMetadataFormatting.h"
#include "FileOperationArtifactRegistry.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.FileSystem.Private.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include <DxUi/DxUi.h>
#include <DxUi/Typography.h>

#include <algorithm>
#include <atomic>
#include <cmath>
#include <mutex>
#include <stop_token>
#include <thread>
#include <unordered_map>

// I35 R7.5 (D-7.2): "Clean Up Leftovers of Interrupted Operations". A modeless DxUi-hosted listing of
// the current folder's Possible artifact rows (optionally its subfolders), Cancel by default. Confirm
// submits an ordinary Delete of the checked items through the File Operations admission, so the touch
// guard, the Delete confirmation and the by-name rules apply unchanged. Nothing here deletes, decides
// that an item is abandoned, or runs without the user: it never starts by itself. The rows of a folder
// a live task of this window works in are left out (they may be that task's own temporary files), and
// the Delete goes to the provider the listing was opened on.

namespace
{
constexpr wchar_t kLeftoverCleanupWindowClassName[] = L"RedSalamander.LeftoverCleanupWindow";
// The subfolder walk is bounded: a cleanup listing is not a search.
constexpr size_t kLeftoverScanMaxFolders = 4096u;
constexpr size_t kLeftoverScanMaxRows    = 4096u;
constexpr size_t kLeftoverScanMaxDepth   = 32u;
constexpr UINT kLeftoverCleanupDelete    = 1u;
// Stopping a walk re-issues the I/O cancel while it waits, at most this long on the UI thread: a
// cancel can land between two synchronous directory reads.
constexpr unsigned int kStopScanCancelAttempts  = 10u;
constexpr DWORD kStopScanCancelIntervalMs       = 20u;
constexpr float kLeftoverNoteMinimumHeightDip   = 40.0f;
constexpr float kLeftoverStatusMinimumHeightDip = 20.0f;
// Process-wide, so that a completion posted by a closed listing's walk never matches a newer listing.
std::atomic<uint64_t> g_leftoverScanGeneration{0u};

struct LeftoverScanLimits final
{
    size_t maxFolders = kLeftoverScanMaxFolders;
    size_t maxRows    = kLeftoverScanMaxRows;
    size_t maxDepth   = kLeftoverScanMaxDepth;
};

[[nodiscard]] std::wstring LeftoverRelativeName(const std::filesystem::path& root, const std::filesystem::path& path)
{
    const std::wstring& rootText = root.native();
    const std::wstring& pathText = path.native();
    if (! rootText.empty() && pathText.size() > rootText.size() && pathText.compare(0u, rootText.size(), rootText) == 0)
    {
        size_t start = rootText.size();
        while (start < pathText.size() && (pathText[start] == L'\\' || pathText[start] == L'/'))
        {
            ++start;
        }
        return pathText.substr(start);
    }
    return path.filename().native();
}

[[nodiscard]] bool IsLeftoverScanDotEntry(std::wstring_view name) noexcept
{
    return name == L"." || name == L"..";
}

// True when the thread has returned; otherwise its pending synchronous I/O is canceled again.
[[nodiscard]] bool ReapScanThread(std::jthread& thread) noexcept
{
    if (! thread.joinable())
    {
        return true;
    }
    const HANDLE handle = thread.native_handle();
    if (WaitForSingleObject(handle, 0u) == WAIT_OBJECT_0)
    {
        thread.join();
        return true;
    }
    static_cast<void>(CancelSynchronousIo(handle));
    return false;
}
} // namespace

struct FolderWindow::LeftoverCleanupDialogState final : DxUi::IGridModel, DxUi::IGridDelegate
{
    using DxUi::IGridDelegate::OnGridCheckboxToggled;
    using DxUi::IGridDelegate::OnGridRowActivated;
    using DxUi::IGridDelegate::OnGridSelectionChanged;

    struct Row final
    {
        std::filesystem::path path;
        std::wstring name;
        uint64_t sizeBytes    = 0u;
        int64_t lastWriteTime = 0;
        DWORD fileAttributes  = 0u;
        bool isDirectory      = false;
        bool mayHoldOriginal  = false;
        bool checked          = false;
    };

    struct ScanResult final
    {
        std::vector<Row> rows;
        bool folderLimit = false;
        bool rowLimit    = false;
        bool depthLimit  = false;
        bool canceled    = false;
    };

    LeftoverCleanupDialogState()                                             = default;
    LeftoverCleanupDialogState(const LeftoverCleanupDialogState&)            = delete;
    LeftoverCleanupDialogState& operator=(const LeftoverCleanupDialogState&) = delete;
    LeftoverCleanupDialogState(LeftoverCleanupDialogState&&)                 = delete;
    LeftoverCleanupDialogState& operator=(LeftoverCleanupDialogState&&)      = delete;
    ~LeftoverCleanupDialogState()
    {
        static_cast<void>(StopScan(true));
    }

    static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = LeftoverCleanupDialogState::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kLeftoverCleanupWindowClassName;
        wc.style         = CS_DBLCLKS;
        atom             = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs    = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* state = static_cast<LeftoverCleanupDialogState*>(cs ? cs->lpCreateParams : nullptr);
            if (! state)
            {
                return FALSE;
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            if (! state->hwnd)
            {
                state->hwnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* state = reinterpret_cast<LeftoverCleanupDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! state)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = state->dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                state->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                state->OnNcDestroy(hwnd);
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return state->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: state->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                state->Layout();
                return 0;
            }
            case WM_GETMINMAXINFO:
            {
                auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
                if (info)
                {
                    const UINT dpi         = GetDpiForWindow(hwnd);
                    info->ptMinTrackSize.x = FolderWindowFileSystemInternal::ScalePanePromptForDpi(dpi, 520);
                    info->ptMinTrackSize.y = FolderWindowFileSystemInternal::ScalePanePromptForDpi(dpi, 360);
                }
                return 0;
            }
            case WndMsg::kLeftoverCleanupScanComplete: state->OnScanComplete(static_cast<uint64_t>(wParam)); return 0;
            case WM_COMMAND:
                switch (LOWORD(wParam))
                {
                    case kLeftoverCleanupDelete:
                        if (state->owner)
                        {
                            state->owner->ConfirmLeftoverCleanupDialog();
                        }
                        return 0;
                    case IDCANCEL:
                        if (state->owner)
                        {
                            state->owner->RequestCloseLeftoverCleanupDialog();
                        }
                        return 0;
                }
                break;
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, state->theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE:
                if (state->owner)
                {
                    state->owner->RequestCloseLeftoverCleanupDialog();
                    return 0;
                }
                break;
            case WM_NCDESTROY: state->OnNcDestroy(hwnd); break;
        }
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return columns.size();
    }

    [[nodiscard]] DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, DxUi::GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= rows.size() || columnIndex >= columns.size())
        {
            return;
        }
        const Row& row = rows[rowIndex];
        switch (columnIndex)
        {
            case 0u:
                outCell.kind    = DxUi::GridCellKind::Checkbox;
                outCell.checked = row.checked;
                outCell.text    = row.name;
                break;
            case 1u: outCell.text = row.isDirectory ? std::wstring() : FormatBytesCompact(row.sizeBytes); break;
            case 2u:
                outCell.text =
                    Common::FileMetadata::FormatDisplayFields({.lastWriteTime100nsSince1601 = row.lastWriteTime, .fileAttributes = row.fileAttributes},
                                                              Common::FileMetadata::DisplayProfile::CompactDetails)
                        .localTime;
                break;
            case 3u: outCell.text = row.mayHoldOriginal ? kindOriginalText : kindLeftoverText; break;
            default: break;
        }
        outCell.tooltipText = row.path.native();
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return rowIndex < rows.size() ? static_cast<uint64_t>(rowIndex + 1u) : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }
        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < rows.size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

    void OnGridCheckboxToggled(DxUi::Grid& sender, size_t rowIndex, size_t /*columnIndex*/, bool checked) override
    {
        if (&sender != grid || rowIndex >= rows.size())
        {
            return;
        }
        rows[rowIndex].checked                    = checked;
        checkStates[rows[rowIndex].path.native()] = checked;
        sender.NotifyDataChanged();
        UpdateState();
    }

    void OnGridRowActivated(DxUi::Grid& sender, size_t rowIndex) override
    {
        // Activation toggles the row; it never deletes.
        if (&sender == grid && rowIndex < rows.size())
        {
            OnGridCheckboxToggled(sender, rowIndex, 0u, ! rows[rowIndex].checked);
        }
    }

    // Whether a live task of this window works in the row's folder: such a row may be that task's own
    // temporary file, so the listing leaves it out (R7.5: no task owns a listed row).
    [[nodiscard]] bool RowInUse(const Row& row) const noexcept
    {
        return owner != nullptr && owner->IsLeftoverFolderInUse(pluginId, row.path.parent_path());
    }

    // Leftovers first, then possible originals, each by name. A row keeps the check state the user gave
    // it (also across the subfolder toggle); a row never seen before starts checked when it is a leftover
    // the location can delete, and a possible original never starts checked.
    void SetRows(std::vector<Row> newRows)
    {
        hiddenInUse = static_cast<size_t>(std::erase_if(newRows, [this](const Row& row) noexcept { return RowInUse(row); }));
        for (Row& row : newRows)
        {
            row.name    = LeftoverRelativeName(folder, row.path);
            row.checked = checkStates.try_emplace(row.path.native(), deleteSupported && ! row.mayHoldOriginal).first->second;
        }
        std::ranges::stable_sort(newRows,
                                 [](const Row& left, const Row& right) noexcept
        {
            if (left.mayHoldOriginal != right.mayHoldOriginal)
            {
                return ! left.mayHoldOriginal;
            }
            return ::CompareStringOrdinal(
                       left.name.c_str(), static_cast<int>(left.name.size()), right.name.c_str(), static_cast<int>(right.name.size()), TRUE) == CSTR_LESS_THAN;
        });
        rows = std::move(newRows);
        if (grid)
        {
            grid->NotifyDataChanged();
            grid->GetSelectionModel().Clear();
        }
        UpdateState();
    }

    // Drops the checked rows whose folder a live task started working in since they were listed; true
    // when any was dropped (the listing then stays open for another look).
    [[nodiscard]] bool DropRowsNowInUse()
    {
        const size_t dropped = static_cast<size_t>(std::erase_if(rows, [this](const Row& row) noexcept { return row.checked && RowInUse(row); }));
        if (dropped == 0u)
        {
            return false;
        }
        hiddenInUse += dropped;
        if (grid)
        {
            grid->NotifyDataChanged();
            grid->GetSelectionModel().Clear();
        }
        UpdateState();
        return true;
    }

    [[nodiscard]] bool OnCreate(HWND createdHwnd) noexcept
    {
        if (! dxHost.Attach(createdHwnd))
        {
            return false;
        }
        BuildUi();
        ApplyTheme();
        Layout();
        dxHost.SetFocusControl(cancelButton);
        // Cancel is the default and the Escape action: Enter never deletes.
        dxHost.SetDefaultButton(cancelButton);
        dxHost.SetCancelButton(cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (root)
        {
            return;
        }
        using namespace DxUi;
        kindLeftoverText = LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_KIND_LEFTOVER);
        kindOriginalText = LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_KIND_ORIGINAL);
        columns          = {
            {L"name", LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_COLUMN_NAME), 300.0f, 160.0f, GridColumnKind::Text, false, false},
            {L"size", LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_COLUMN_SIZE), 90.0f, 64.0f, GridColumnKind::Text, false, false},
            {L"modified", LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_COLUMN_MODIFIED), 150.0f, 96.0f, GridColumnKind::Text, false, false},
            {L"kind", LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_COLUMN_KIND), 140.0f, 96.0f, GridColumnKind::Text, false, false},
        };

        auto rootOwned = std::make_unique<Panel>();
        root           = rootOwned.get();

        noteLabel = root->AddChild<Label>(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_NOTE));
        noteLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        noteLabel->SetMultiline(true);

        subfoldersCheckbox = root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_INCLUDE_SUBFOLDERS));
        subfoldersCheckbox->SetChecked(includeSubfolders);
        subfoldersCheckbox->SetOnToggled([this](bool checked) { SetIncludeSubfolders(checked); });

        grid = root->AddChild<Grid>();
        grid->SetModel(this);
        grid->SetDelegate(this);
        grid->SetSelectionMode(GridSelectionMode::Single);
        grid->SetEmptyStateText(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_EMPTY));
        grid->SetLineClamp(1u);

        statusLabel = root->AddChild<Label>(std::wstring());
        statusLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        statusLabel->SetMultiline(true);

        deleteButton = root->AddChild<Button>(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_DELETE));
        deleteButton->SetOnClick([this]
        {
            if (hwnd && IsWindow(hwnd.get()) != FALSE)
            {
                PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(kLeftoverCleanupDelete, BN_CLICKED), 0);
            }
        });
        cancelButton = root->AddChild<Button>(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_CANCEL));
        cancelButton->SetPrimary(true);
        cancelButton->SetOnClick([this]
        {
            if (hwnd && IsWindow(hwnd.get()) != FALSE)
            {
                PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDCANCEL, BN_CLICKED), 0);
            }
        });

        dxHost.SetRoot(std::move(rootOwned));
        SetRows(baseRows);
    }

    void ApplyTheme() noexcept
    {
        palette = MakeAppThemeDxPalette(theme, theme.windowBackground);
        dxHost.SetTheme(palette);
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            ApplyWindowChromeTheme(hwnd.get(), theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd.get());
            RedrawWindow(hwnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
        }
    }

    // The wrapped height of a label's text at this width, never below `minimumDip`.
    [[nodiscard]] float MeasureLabelHeightDip(const DxUi::Label* label, float widthDip, float minimumDip) const noexcept
    {
        if (label == nullptr || label->GetText().empty() || widthDip <= 1.0f || dxHost.GetHwnd() == nullptr)
        {
            return minimumDip;
        }
        const int widthPx  = std::max(1, static_cast<int>(std::floor(dxHost.DipsToPixels(widthDip))));
        const int heightPx = DxUi::Typography::MeasureWrappedTextHeightPx(dxHost.GetHwnd(), DxUi::FontRole::Body, widthPx, label->GetText());
        return std::max(minimumDip, std::ceil(dxHost.PixelsToDip(static_cast<float>(heightPx))) + 4.0f);
    }

    void Layout() noexcept
    {
        if (! root)
        {
            return;
        }
        const D2D1_RECT_F client = dxHost.GetClientBoundsDip();
        root->SetBounds(client);

        constexpr float kMarginDip       = 20.0f;
        constexpr float kGapDip          = 12.0f;
        constexpr float kToggleHeightDip = 28.0f;
        constexpr float kButtonWidthDip  = 160.0f;
        constexpr float kButtonHeightDip = 34.0f;

        const float left      = client.left + kMarginDip;
        const float right     = std::max(left, client.right - kMarginDip);
        const float bottom    = std::max(client.top, client.bottom - kMarginDip);
        const float buttonTop = std::max(client.top + kMarginDip, bottom - kButtonHeightDip);
        // The note and the status lines take the height their text needs at this width (a translated
        // note wraps to more lines at the minimum width); the grid takes what remains.
        const float noteHeight   = MeasureLabelHeightDip(noteLabel, right - left, kLeftoverNoteMinimumHeightDip);
        const float statusHeight = MeasureLabelHeightDip(statusLabel, right - left, kLeftoverStatusMinimumHeightDip);
        float y                  = client.top + kMarginDip;
        if (noteLabel)
        {
            noteLabel->SetBounds(D2D1::RectF(left, y, right, y + noteHeight));
        }
        y += noteHeight + kGapDip;
        if (subfoldersCheckbox)
        {
            subfoldersCheckbox->SetBounds(D2D1::RectF(left, y, right, y + kToggleHeightDip));
        }
        y += kToggleHeightDip + kGapDip;
        const float statusTop = std::max(y, buttonTop - kGapDip - statusHeight);
        if (grid)
        {
            grid->SetBounds(D2D1::RectF(left, y, right, std::max(y, statusTop - kGapDip)));
        }
        if (statusLabel)
        {
            statusLabel->SetBounds(D2D1::RectF(left, statusTop, right, statusTop + statusHeight));
        }
        float x = right;
        if (cancelButton)
        {
            cancelButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
            x -= kButtonWidthDip + kGapDip;
        }
        if (deleteButton)
        {
            deleteButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
        }
    }

    [[nodiscard]] std::wstring BuildStatusText() const
    {
        std::wstring status;
        const auto append = [&status](std::wstring line)
        {
            if (line.empty())
            {
                return;
            }
            if (! status.empty())
            {
                status += L'\n';
            }
            status += line;
        };
        if (! deleteSupported)
        {
            append(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_CANNOT_DELETE));
        }
        if (scanning)
        {
            append(LoadStringResource(nullptr, IDS_LEFTOVER_CLEANUP_SCANNING));
            return status;
        }
        if (folderLimit)
        {
            append(FormatStringResource(nullptr, IDS_FMT_LEFTOVER_CLEANUP_SCAN_LIMIT, static_cast<unsigned long>(kLeftoverScanMaxFolders)));
        }
        if (rowLimit)
        {
            append(FormatStringResource(nullptr, IDS_FMT_LEFTOVER_CLEANUP_ROW_LIMIT, static_cast<unsigned long>(kLeftoverScanMaxRows)));
        }
        if (depthLimit)
        {
            append(FormatStringResource(nullptr, IDS_FMT_LEFTOVER_CLEANUP_DEPTH_LIMIT, static_cast<unsigned long>(kLeftoverScanMaxDepth)));
        }
        if (hiddenInUse != 0u)
        {
            append(FormatStringResource(nullptr, IDS_FMT_LEFTOVER_CLEANUP_IN_USE, static_cast<unsigned long long>(hiddenInUse)));
        }
        return status;
    }

    void UpdateState() noexcept
    {
        const size_t checkedCount = CountChecked();
        if (deleteButton)
        {
            deleteButton->SetEnabled(deleteSupported && ! scanning && checkedCount != 0u);
        }
        if (statusLabel)
        {
            std::wstring status = BuildStatusText();
            if (status != statusLabel->GetText())
            {
                statusLabel->SetText(std::move(status));
                Layout();
            }
        }
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            InvalidateRect(hwnd.get(), nullptr, FALSE);
        }
    }

    [[nodiscard]] size_t CountChecked() const noexcept
    {
        return static_cast<size_t>(std::ranges::count_if(rows, [](const Row& row) noexcept { return row.checked; }));
    }

    [[nodiscard]] std::vector<std::filesystem::path> CheckedPaths() const
    {
        std::vector<std::filesystem::path> paths;
        for (const Row& row : rows)
        {
            if (row.checked)
            {
                paths.push_back(row.path);
            }
        }
        return paths;
    }

    // The current folder's rows come from the pane listing; including subfolders walks them on a
    // worker (bounded, cancelable) and replaces the list when it is done.
    void SetIncludeSubfolders(bool include)
    {
        if (include == includeSubfolders)
        {
            return;
        }
        includeSubfolders = include;
        static_cast<void>(StopScan(false));
        folderLimit = false;
        rowLimit    = false;
        depthLimit  = false;
        if (! include)
        {
            scanning = false;
            SetRows(baseRows);
            return;
        }
        scanning = true;
        rows.clear();
        if (grid)
        {
            grid->NotifyDataChanged();
        }
        UpdateState();
        StartScan();
    }

    // The bounded walk: rows are the Possible names under `rootFolder`; a folder with a Possible name is
    // listed, not entered, and a link is never followed. Each bound reached sets its own flag.
    static void ScanFolderTree(IFileSystem* fileSystem,
                               const std::filesystem::path& rootFolder,
                               const LeftoverScanLimits& limits,
                               const std::stop_token& stopToken,
                               ScanResult& out) noexcept
    {
        out = {};
        std::vector<std::pair<std::filesystem::path, size_t>> pending;
        pending.emplace_back(rootFolder, 0u);
        size_t folders = 0u;
        while (! pending.empty() && ! out.rowLimit)
        {
            if (stopToken.stop_requested())
            {
                out.canceled = true;
                return;
            }
            if (folders >= limits.maxFolders)
            {
                out.folderLimit = true;
                return;
            }
            auto [directory, depth] = std::move(pending.back());
            pending.pop_back();
            ++folders;
            wil::com_ptr<IFilesInformation> info;
            if (fileSystem == nullptr || FAILED(fileSystem->ReadDirectoryInfo(directory.c_str(), info.put())) || ! info)
            {
                continue;
            }
            FileInfo* entry = nullptr;
            if (FAILED(info->GetBuffer(&entry)))
            {
                continue;
            }
            for (; entry != nullptr;
                 entry = entry->NextEntryOffset != 0u ? reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset) : nullptr)
            {
                if (stopToken.stop_requested())
                {
                    out.canceled = true;
                    return;
                }
                const std::wstring_view name(entry->FileName, entry->FileNameSize / sizeof(wchar_t));
                if (name.empty() || IsLeftoverScanDotEntry(name))
                {
                    continue;
                }
                const bool directoryEntry   = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
                const bool reparse          = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u;
                std::filesystem::path child = directory;
                child /= std::wstring(name);
                if (FileOperationArtifacts::HasPossibleArtifactName(name))
                {
                    if (out.rows.size() >= limits.maxRows)
                    {
                        out.rowLimit = true;
                        break;
                    }
                    out.rows.push_back(Row{
                        .path            = std::move(child),
                        .sizeBytes       = entry->EndOfFile > 0 ? static_cast<uint64_t>(entry->EndOfFile) : 0u,
                        .lastWriteTime   = entry->LastWriteTime,
                        .fileAttributes  = entry->FileAttributes,
                        .isDirectory     = directoryEntry,
                        .mayHoldOriginal = FileOperationArtifacts::HasPossibleOriginalArtifactName(name),
                    });
                }
                else if (directoryEntry && ! reparse)
                {
                    // A link is never followed; a leftover folder is listed, not entered.
                    if (depth + 1u < limits.maxDepth)
                    {
                        pending.emplace_back(std::move(child), depth + 1u);
                    }
                    else
                    {
                        out.depthLimit = true;
                    }
                }
            }
        }
    }

    void StartScan()
    {
        const uint64_t generation = g_leftoverScanGeneration.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        scanGeneration            = generation;
        const HWND target         = hwnd.get();
        scanThread = std::jthread([this, generation, target, fileSystemCopy = fileSystem, rootFolder = folder](std::stop_token stopToken) noexcept
        {
            ScanResult result;
            ScanFolderTree(fileSystemCopy.get(), rootFolder, LeftoverScanLimits{}, stopToken, result);
            if (stopToken.stop_requested() || result.canceled)
            {
                return;
            }
            {
                std::scoped_lock lock(scanMutex);
                scanResult           = std::move(result);
                scanResultGeneration = generation;
            }
            if (target != nullptr)
            {
                static_cast<void>(PostMessageW(target, WndMsg::kLeftoverCleanupScanComplete, static_cast<WPARAM>(generation), 0));
            }
        });
    }

    void OnScanComplete(uint64_t generation)
    {
        if (generation != scanGeneration || ! includeSubfolders)
        {
            return;
        }
        ScanResult result;
        {
            std::scoped_lock lock(scanMutex);
            if (scanResultGeneration != generation)
            {
                return;
            }
            result               = std::move(scanResult);
            scanResult           = {};
            scanResultGeneration = 0u;
        }
        if (scanThread.joinable())
        {
            // The walk posted its completion as its last step.
            scanThread.join();
        }
        scanning    = false;
        folderLimit = result.folderLimit;
        rowLimit    = result.rowLimit;
        depthLimit  = result.depthLimit;
        SetRows(std::move(result.rows));
    }

    // Stops every walk of this listing: the stop request ends a walk between directory reads, and the
    // I/O cancel is re-issued while waiting because one can land between two synchronous reads. Without
    // `wait` it waits at most kStopScanCancelAttempts intervals and returns false when a walk is still
    // blocked (the owner then keeps this listing until it returns); with `wait` it waits for all of them.
    [[nodiscard]] bool StopScan(bool wait) noexcept
    {
        if (scanThread.joinable())
        {
            scanThread.request_stop();
            stoppedScanThreads.push_back(std::move(scanThread));
        }
        for (unsigned int attempt = 0u; ! stoppedScanThreads.empty(); ++attempt)
        {
            std::erase_if(stoppedScanThreads, [](std::jthread& thread) noexcept { return ReapScanThread(thread); });
            if (stoppedScanThreads.empty() || (! wait && attempt + 1u >= kStopScanCancelAttempts))
            {
                break;
            }
            Sleep(kStopScanCancelIntervalMs);
        }
        return stoppedScanThreads.empty();
    }

    // True once no walk of this listing runs any more (non-blocking).
    [[nodiscard]] bool ReapStoppedScans() noexcept
    {
        std::erase_if(stoppedScanThreads, [](std::jthread& thread) noexcept { return ReapScanThread(thread); });
        return stoppedScanThreads.empty() && ! scanThread.joinable();
    }

    void OnNcDestroy(HWND destroyedHwnd) noexcept
    {
        destroyed = true;
        // Bounded: a walk still blocked keeps this state alive through the owner's retired list.
        static_cast<void>(StopScan(false));
        dxHost.Detach();
        if (hwnd.get() == destroyedHwnd)
        {
            static_cast<void>(hwnd.release());
        }
        SetWindowLongPtrW(destroyedHwnd, GWLP_USERDATA, 0);
    }

    FolderWindow* owner = nullptr;
    Pane pane           = Pane::Left;
    std::filesystem::path folder;
    // The provider and plugin identity the listing was opened on; Delete Selected goes there.
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    bool deleteSupported = false;
    AppTheme theme;
    DxUi::ThemePalette palette{};
    DxUi::WindowHost dxHost;
    DxUi::Panel* root                  = nullptr;
    DxUi::Label* noteLabel             = nullptr;
    DxUi::Checkbox* subfoldersCheckbox = nullptr;
    DxUi::Grid* grid                   = nullptr;
    DxUi::Label* statusLabel           = nullptr;
    DxUi::Button* deleteButton         = nullptr;
    DxUi::Button* cancelButton         = nullptr;
    std::vector<DxUi::GridColumnDesc> columns;
    std::wstring kindLeftoverText;
    std::wstring kindOriginalText;
    std::vector<Row> baseRows;
    std::vector<Row> rows;
    // The user's check state per listed path, kept across rebuilds of the list.
    std::unordered_map<std::wstring, bool> checkStates;
    size_t hiddenInUse      = 0u;
    bool includeSubfolders  = false;
    bool scanning           = false;
    bool folderLimit        = false;
    bool rowLimit           = false;
    bool depthLimit         = false;
    uint64_t scanGeneration = 0u;
    std::mutex scanMutex;
    ScanResult scanResult;
    uint64_t scanResultGeneration = 0u;
    std::jthread scanThread;
    std::vector<std::jthread> stoppedScanThreads;
    bool destroyed = false;
    wil::unique_hwnd hwnd;
};

void FolderWindow::LeftoverCleanupDialogStateDeleter::operator()(LeftoverCleanupDialogState* state) const noexcept
{
    delete state;
}

void FolderWindow::CommandCleanUpLeftovers(Pane pane)
{
    Debug::Perf::Scope perf(L"leftoverCleanup.open_us");
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem || ! state.currentPath.has_value())
    {
        return;
    }

    std::vector<LeftoverCleanupDialogState::Row> rows;
    for (FolderView::PossibleArtifactRow& possible : state.folderView.GetPossibleArtifactRows())
    {
        rows.push_back(LeftoverCleanupDialogState::Row{
            .path            = std::move(possible.path),
            .sizeBytes       = possible.sizeBytes,
            .lastWriteTime   = possible.lastWriteTime,
            .fileAttributes  = possible.fileAttributes,
            .isDirectory     = possible.isDirectory,
            .mayHoldOriginal = possible.mayHoldOriginal,
        });
    }

    // One listing at a time: a second request replaces the first one's folder.
    CloseLeftoverCleanupDialog();
    std::unique_ptr<LeftoverCleanupDialogState, LeftoverCleanupDialogStateDeleter> dialog{std::make_unique<LeftoverCleanupDialogState>().release()};
    dialog->owner           = this;
    dialog->pane            = pane;
    dialog->folder          = state.currentPath.value();
    dialog->fileSystem      = state.fileSystem;
    dialog->pluginId        = state.pluginId;
    dialog->pluginShortId   = state.pluginShortId;
    dialog->instanceContext = state.instanceContext;
    // A location that cannot delete (an MTP device) lists its leftovers but offers no Delete, and says why.
    const std::wstring& folderText = dialog->folder.native();
    dialog->deleteSupported =
        CanSameFileSystemOperation(state.fileSystem,
                                   folderText,
                                   FILESYSTEM_DELETE,
                                   state.pluginId,
                                   static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN)) ||
        CanSameFileSystemOperation(state.fileSystem, folderText, FILESYSTEM_DELETE, state.pluginId, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE));
#ifdef ENABLE_TESTS
    if (_debugLeftoverCleanupDeleteUnsupportedForSelfTest)
    {
        dialog->deleteSupported = false;
    }
#endif
    dialog->theme    = _theme;
    dialog->baseRows = std::move(rows);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }
    const HRESULT classHr = LeftoverCleanupDialogState::EnsureWindowClass();
    if (FAILED(classHr))
    {
        perf.SetHr(classHr);
        Debug::Warning(L"FolderWindow::CommandCleanUpLeftovers: failed to register the window class (hr=0x{:08X}).", static_cast<unsigned long>(classHr));
        return;
    }

    const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    const DWORD exStyle = WS_EX_DLGMODALFRAME;
    const UINT dpi      = ownerWindow && IsWindow(ownerWindow) != FALSE ? GetDpiForWindow(ownerWindow) : GetDpiForSystem();
    RECT bounds{0, 0, FolderWindowFileSystemInternal::ScalePanePromptForDpi(dpi, 760), FolderWindowFileSystemInternal::ScalePanePromptForDpi(dpi, 520)};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
    {
        perf.SetHr(HRESULT_FROM_WIN32(GetLastError()));
        return;
    }

    const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_CLEANUP_LEFTOVERS);
    HWND window                = CreateWindowExW(exStyle,
                                                 kLeftoverCleanupWindowClassName,
                                                 caption.c_str(),
                                                 style,
                                                 CW_USEDEFAULT,
                                                 CW_USEDEFAULT,
                                                 bounds.right - bounds.left,
                                                 bounds.bottom - bounds.top,
                                                 ownerWindow,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 dialog.get());
    if (! window)
    {
        perf.SetHr(HRESULT_FROM_WIN32(GetLastError()));
        Debug::Warning(L"FolderWindow::CommandCleanUpLeftovers: failed to create the listing window.");
        return;
    }
    if (! dialog->hwnd)
    {
        dialog->hwnd = wil::unique_hwnd(window);
    }
    perf.SetValue0(static_cast<uint64_t>(dialog->rows.size()));
    _leftoverCleanupDialog = std::move(dialog);
    static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(window, ownerWindow));
    static_cast<void>(_leftoverCleanupDialog->dxHost.PrimeForShow());
    ShowWindow(window, SW_SHOWNORMAL);
    UpdateWindow(window);
    Common::RequestForegroundWindow(window);
}

void FolderWindow::ConfirmLeftoverCleanupDialog() noexcept
{
    if (! _leftoverCleanupDialog || _leftoverCleanupDialog->scanning || ! _leftoverCleanupDialog->deleteSupported)
    {
        return;
    }
    // A task may have started working in a listed row's folder since the listing was built: such rows
    // leave the listing and nothing is submitted, so the user sees what remains first.
    if (_leftoverCleanupDialog->DropRowsNowInUse())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }
    std::vector<std::filesystem::path> paths = _leftoverCleanupDialog->CheckedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }
    const Pane pane                          = _leftoverCleanupDialog->pane;
    wil::com_ptr<IFileSystem> listedProvider = _leftoverCleanupDialog->fileSystem;
    std::wstring listedPluginId              = _leftoverCleanupDialog->pluginId;
    std::wstring listedPluginShortId         = _leftoverCleanupDialog->pluginShortId;
    std::wstring listedInstanceContext       = _leftoverCleanupDialog->instanceContext;
    RequestCloseLeftoverCleanupDialog();
    static_cast<void>(StartLeftoverCleanupDelete(
        pane, std::move(listedProvider), std::move(listedPluginId), std::move(listedPluginShortId), std::move(listedInstanceContext), std::move(paths)));
}

HRESULT FolderWindow::StartLeftoverCleanupDelete(Pane pane,
                                                 wil::com_ptr<IFileSystem> fileSystem,
                                                 std::wstring pluginId,
                                                 std::wstring pluginShortId,
                                                 std::wstring instanceContext,
                                                 std::vector<std::filesystem::path> paths) noexcept
{
    EnsureFileOperations();
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations || ! fileSystem || pluginId.empty() || paths.empty())
    {
        return E_UNEXPECTED;
    }

    // The ordinary Delete: to the Recycle Bin where the location has one (the pane's Delete), else a
    // permanent delete behind its own card. The touch guard, the confirmation and the by-name rules of
    // an identity-less location apply exactly as to any other Delete. The provider is the one the
    // listing was opened on, even when the pane has moved on since.
    FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    bool recycle          = CanSameFileSystemOperation(fileSystem, paths.front().native(), FILESYSTEM_DELETE, pluginId, flags);
#ifdef ENABLE_TESTS
    if (_debugLeftoverCleanupPermanentForSelfTest)
    {
        recycle = false;
    }
#endif
    if (! recycle)
    {
        flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        if (! CanSameFileSystemOperation(fileSystem, paths.front().native(), FILESYSTEM_DELETE, pluginId, flags))
        {
            state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                              FolderView::OverlaySeverity::Error,
                                              LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                              LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
    }

    Debug::Info(L"FolderWindow: cleaning up {} leftover(s) of interrupted operations ({}).", paths.size(), recycle ? L"Recycle Bin" : L"permanent");
    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    uint64_t taskId          = 0u;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    const HRESULT startHr = _fileOperations->AdmitOperation(FILESYSTEM_DELETE,
                                                            pane,
                                                            std::nullopt,
                                                            fileSystem,
                                                            std::move(paths),
                                                            {},
                                                            flags,
                                                            waitForOthers,
                                                            0,
                                                            recycle ? FileOperationState::ExecutionMode::BulkItems : FileOperationState::ExecutionMode::PerItem,
                                                            ! recycle,
                                                            nullptr,
                                                            &taskId,
                                                            {},
                                                            {},
                                                            std::move(pluginId),
                                                            std::move(pluginShortId),
                                                            std::move(instanceContext),
                                                            FileOperations::DeleteOrigin::LeftoverCleanup);
#ifdef ENABLE_TESTS
    _debugLeftoverCleanupLastTaskId = SUCCEEDED(startHr) ? taskId : 0u;
#endif
    ReportFileOperationAdmissionFailure(pane, FILESYSTEM_DELETE, startHr);
    return startHr;
}

bool FolderWindow::IsLeftoverFolderInUse(std::wstring_view pluginId, const std::filesystem::path& folder) noexcept
{
    return _fileOperations && ! folder.empty() && _fileOperations->FindLiveTaskWorkingIn(pluginId, folder.native()) != 0u;
}

void FolderWindow::RequestCloseLeftoverCleanupDialog() noexcept
{
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE && PostMessageW(_hWnd.get(), WndMsg::kFolderWindowCloseLeftoverCleanupDialog, 0, 0) != 0)
    {
        return;
    }
    CloseLeftoverCleanupDialog();
}

void FolderWindow::CloseLeftoverCleanupDialog() noexcept
{
    if (_leftoverCleanupDialog)
    {
        // Bounded: a walk still blocked in a canceled read keeps the listing (window closed) until it returns.
        const bool stopped = _leftoverCleanupDialog->StopScan(false);
        if (_leftoverCleanupDialog->hwnd && ! _leftoverCleanupDialog->destroyed && IsWindow(_leftoverCleanupDialog->hwnd.get()) != FALSE)
        {
            _leftoverCleanupDialog->hwnd.reset();
        }
        if (stopped)
        {
            _leftoverCleanupDialog.reset();
        }
        else
        {
            _retiredLeftoverCleanupDialogs.push_back(std::move(_leftoverCleanupDialog));
        }
    }
    SweepRetiredLeftoverCleanupDialogs(false);
}

void FolderWindow::SweepRetiredLeftoverCleanupDialogs(bool wait) noexcept
{
    if (wait)
    {
        // Window teardown: every retired walk is canceled until it returns, then freed.
        _retiredLeftoverCleanupDialogs.clear();
        return;
    }
    std::erase_if(_retiredLeftoverCleanupDialogs, [](const std::unique_ptr<LeftoverCleanupDialogState, LeftoverCleanupDialogStateDeleter>& retired) noexcept {
        return ! retired || retired->ReapStoppedScans();
    });
}

void FolderWindow::ApplyLeftoverCleanupDialogTheme() noexcept
{
    if (! _leftoverCleanupDialog || ! _leftoverCleanupDialog->hwnd || IsWindow(_leftoverCleanupDialog->hwnd.get()) == FALSE)
    {
        return;
    }
    _leftoverCleanupDialog->theme = _theme;
    _leftoverCleanupDialog->ApplyTheme();
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetLeftoverCleanupSnapshot(LeftoverCleanupDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _leftoverCleanupDialog || ! _leftoverCleanupDialog->hwnd || IsWindow(_leftoverCleanupDialog->hwnd.get()) == FALSE)
    {
        return false;
    }
    out.open              = true;
    out.includeSubfolders = _leftoverCleanupDialog->includeSubfolders;
    out.scanning          = _leftoverCleanupDialog->scanning;
    out.deleteEnabled     = _leftoverCleanupDialog->deleteButton && _leftoverCleanupDialog->deleteButton->IsEnabled();
    out.deleteSupported   = _leftoverCleanupDialog->deleteSupported;
    out.folderLimit       = _leftoverCleanupDialog->folderLimit;
    out.rowLimit          = _leftoverCleanupDialog->rowLimit;
    out.depthLimit        = _leftoverCleanupDialog->depthLimit;
    out.hiddenInUse       = _leftoverCleanupDialog->hiddenInUse;
    out.status            = _leftoverCleanupDialog->statusLabel ? std::wstring(_leftoverCleanupDialog->statusLabel->GetText()) : std::wstring();
    for (const LeftoverCleanupDialogState::Row& row : _leftoverCleanupDialog->rows)
    {
        out.rows.push_back(LeftoverCleanupDebugSnapshot::Row{.path = row.path, .checked = row.checked, .mayHoldOriginal = row.mayHoldOriginal});
    }
    return true;
}

HWND FolderWindow::DebugGetLeftoverCleanupWindowForSelfTest() const noexcept
{
    return _leftoverCleanupDialog && _leftoverCleanupDialog->hwnd && IsWindow(_leftoverCleanupDialog->hwnd.get()) != FALSE ? _leftoverCleanupDialog->hwnd.get()
                                                                                                                           : nullptr;
}

bool FolderWindow::DebugSetLeftoverCleanupIncludeSubfolders(bool include) noexcept
{
    if (! _leftoverCleanupDialog)
    {
        return false;
    }
    if (_leftoverCleanupDialog->subfoldersCheckbox)
    {
        _leftoverCleanupDialog->subfoldersCheckbox->SetChecked(include);
    }
    _leftoverCleanupDialog->SetIncludeSubfolders(include);
    return true;
}

bool FolderWindow::DebugToggleLeftoverCleanupRow(const std::filesystem::path& path, bool checked) noexcept
{
    if (! _leftoverCleanupDialog || ! _leftoverCleanupDialog->grid)
    {
        return false;
    }
    const auto found = std::ranges::find(_leftoverCleanupDialog->rows, path, &LeftoverCleanupDialogState::Row::path);
    if (found == _leftoverCleanupDialog->rows.end())
    {
        return false;
    }
    _leftoverCleanupDialog->OnGridCheckboxToggled(
        *_leftoverCleanupDialog->grid, static_cast<size_t>(std::distance(_leftoverCleanupDialog->rows.begin(), found)), 0u, checked);
    return true;
}

HRESULT FolderWindow::DebugScanLeftoversForSelfTest(IFileSystem* fileSystem,
                                                    const std::filesystem::path& folder,
                                                    size_t maxFolders,
                                                    size_t maxRows,
                                                    size_t maxDepth,
                                                    bool cancelBeforeStart,
                                                    LeftoverScanDebugResult& out) noexcept
{
    out = {};
    if (fileSystem == nullptr || folder.empty())
    {
        return E_INVALIDARG;
    }
    std::stop_source stop;
    if (cancelBeforeStart)
    {
        static_cast<void>(stop.request_stop());
    }
    LeftoverCleanupDialogState::ScanResult result;
    LeftoverCleanupDialogState::ScanFolderTree(
        fileSystem, folder, LeftoverScanLimits{.maxFolders = maxFolders, .maxRows = maxRows, .maxDepth = maxDepth}, stop.get_token(), result);
    for (LeftoverCleanupDialogState::Row& row : result.rows)
    {
        out.rows.push_back(std::move(row.path));
    }
    out.folderLimit = result.folderLimit;
    out.rowLimit    = result.rowLimit;
    out.depthLimit  = result.depthLimit;
    out.canceled    = result.canceled;
    return S_OK;
}

bool FolderWindow::DebugAnswerLeftoverCleanup(bool confirm) noexcept
{
    if (! _leftoverCleanupDialog || ! _leftoverCleanupDialog->hwnd || IsWindow(_leftoverCleanupDialog->hwnd.get()) == FALSE)
    {
        return false;
    }
    return PostMessageW(_leftoverCleanupDialog->hwnd.get(), WM_COMMAND, MAKEWPARAM(confirm ? kLeftoverCleanupDelete : IDCANCEL, BN_CLICKED), 0) != FALSE;
}

void FolderWindow::DebugSetLeftoverCleanupPermanentForSelfTest(bool permanent) noexcept
{
    _debugLeftoverCleanupPermanentForSelfTest = permanent;
}

void FolderWindow::DebugSetLeftoverCleanupDeleteUnsupportedForSelfTest(bool unsupported) noexcept
{
    _debugLeftoverCleanupDeleteUnsupportedForSelfTest = unsupported;
}

uint64_t FolderWindow::DebugGetLeftoverCleanupLastTaskId() const noexcept
{
    return _debugLeftoverCleanupLastTaskId;
}
#endif

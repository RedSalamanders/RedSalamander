#include "FolderWindowInternal.h"

namespace
{
struct SelectionSizePayload
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    uint64_t generation     = 0;
    uint64_t folderBytes    = 0;
    HRESULT status          = E_FAIL;
};

struct SelectionSizeProgressPayload
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    uint64_t generation     = 0;
    uint64_t folderBytes    = 0;
};

bool IsDotOrDotDot(const FileInfo* entry) noexcept
{
    if (! entry)
    {
        return false;
    }

    constexpr unsigned long kDotNameBytes    = static_cast<unsigned long>(sizeof(wchar_t));
    constexpr unsigned long kDotDotNameBytes = static_cast<unsigned long>(2u * sizeof(wchar_t));

    if (entry->FileNameSize == kDotNameBytes)
    {
        return entry->FileName[0] == L'.';
    }

    if (entry->FileNameSize == kDotDotNameBytes)
    {
        return entry->FileName[0] == L'.' && entry->FileName[1] == L'.';
    }

    return false;
}

HRESULT AccumulateFolderBytesSubtree(IFileSystem* fileSystem,
                                     const std::vector<std::filesystem::path>& folders,
                                     std::stop_token stopToken,
                                     uint64_t* outFolderBytes,
                                     std::function<void(uint64_t)> progressCallback) noexcept
{
    if (! outFolderBytes)
    {
        return E_POINTER;
    }

    *outFolderBytes = 0;
    if (! fileSystem || folders.empty())
    {
        return S_OK;
    }

    HRESULT firstFailure      = S_OK;
    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();

    constexpr ULONGLONG kProgressReportIntervalMs = 100ull;
    ULONGLONG lastProgressReportTick              = GetTickCount64();
    uint64_t lastProgressBytes                    = 0;
    auto maybeReportProgress                      = [&]() noexcept
    {
        if (! progressCallback)
        {
            return;
        }

        if (stopToken.stop_requested())
        {
            return;
        }

        const ULONGLONG now = GetTickCount64();
        if (now - lastProgressReportTick < kProgressReportIntervalMs)
        {
            return;
        }

        if (*outFolderBytes == lastProgressBytes)
        {
            return;
        }

        lastProgressReportTick = now;
        lastProgressBytes      = *outFolderBytes;
        progressCallback(*outFolderBytes);
    };

    std::vector<std::filesystem::path> pending;
    pending.reserve(folders.size());
    for (const auto& folder : folders)
    {
        pending.push_back(folder);
    }

    while (! pending.empty() && ! stopToken.stop_requested())
    {
        std::filesystem::path current = std::move(pending.back());
        pending.pop_back();

        DirectoryInfoCache::Borrowed borrowed = cache.BorrowDirectoryInfo(fileSystem, current, DirectoryInfoCache::BorrowMode::AllowEnumerate, stopToken);
        const HRESULT borrowHr                = borrowed.Status();
        if (FAILED(borrowHr) || borrowed.Get() == nullptr)
        {
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = FAILED(borrowHr) ? borrowHr : E_FAIL;
            }
            continue;
        }

        FileInfo* entry        = nullptr;
        const HRESULT bufferHr = borrowed.Get()->GetBuffer(&entry);
        if (FAILED(bufferHr))
        {
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = bufferHr;
            }
            continue;
        }

        if (! entry)
        {
            continue;
        }

        unsigned long bufferSize   = 0;
        const HRESULT bufferSizeHr = borrowed.Get()->GetBufferSize(&bufferSize);
        if (FAILED(bufferSizeHr))
        {
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = bufferSizeHr;
            }
            continue;
        }

        unsigned long allocatedSize   = 0;
        const HRESULT allocatedSizeHr = borrowed.Get()->GetAllocatedSize(&allocatedSize);
        if (FAILED(allocatedSizeHr))
        {
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = allocatedSizeHr;
            }
            continue;
        }

        if (allocatedSize < bufferSize || allocatedSize < sizeof(FileInfo))
        {
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            continue;
        }

        std::byte* base = reinterpret_cast<std::byte*>(entry);
        std::byte* end  = base + bufferSize;

        while (! stopToken.stop_requested())
        {
            if (! IsDotOrDotDot(entry))
            {
                const DWORD attrs      = entry->FileAttributes;
                const bool isDirectory = (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    if ((attrs & FILE_ATTRIBUTE_REPARSE_POINT) == 0)
                    {
                        const size_t nameChars      = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                        std::filesystem::path child = current;
                        child /= std::wstring_view(entry->FileName, nameChars);
                        pending.push_back(std::move(child));
                    }
                }
                else if (entry->EndOfFile > 0)
                {
                    const uint64_t fileBytes = static_cast<uint64_t>(entry->EndOfFile);
                    if (*outFolderBytes > std::numeric_limits<uint64_t>::max() - fileBytes)
                    {
                        *outFolderBytes = std::numeric_limits<uint64_t>::max();
                        if (SUCCEEDED(firstFailure))
                        {
                            firstFailure = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                        }
                        break;
                    }
                    *outFolderBytes += fileBytes;
                    maybeReportProgress();
                }
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            if (entry->NextEntryOffset < sizeof(FileInfo))
            {
                if (SUCCEEDED(firstFailure))
                {
                    firstFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                break;
            }

            std::byte* next = reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset;
            if (next < base || next + sizeof(FileInfo) > end)
            {
                if (SUCCEEDED(firstFailure))
                {
                    firstFailure = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                break;
            }

            entry = reinterpret_cast<FileInfo*>(next);
        }
    }

    return firstFailure;
}
} // namespace

void FolderWindow::StartSelectionSizeWorker(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (state.selectionSizeThread.joinable())
    {
        return;
    }

    state.selectionSizeThread = std::jthread([this, pane](std::stop_token stopToken) noexcept { SelectionSizeWorkerMain(pane, stopToken); });
}

void FolderWindow::CancelSelectionSizeComputation(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    ++state.selectionSizeGeneration;

    state.selectionFolderBytesPending = false;
    state.selectionFolderBytesValid   = false;
    state.selectionFolderBytes        = 0;

    std::scoped_lock lock(state.selectionSizeMutex);
    if (state.selectionSizeWorkStopSource)
    {
        state.selectionSizeWorkStopSource->request_stop();
        state.selectionSizeWorkStopSource.reset();
    }
    state.selectionSizeWorkPending    = false;
    state.selectionSizeWorkGeneration = 0;
    state.selectionSizeWorkFolders.clear();
    state.selectionSizeWorkFileSystem.reset();
}

void FolderWindow::RequestSelectionSizeComputation(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    StartSelectionSizeWorker(pane);
    CancelSelectionSizeComputation(pane);

    if (! state.fileSystem || state.selectionStats.selectedFolders == 0)
    {
        UpdatePaneStatusBar(pane);
        return;
    }

    std::vector<std::filesystem::path> folders = state.folderView.GetSelectedDirectoryPaths();
    if (folders.empty())
    {
        UpdatePaneStatusBar(pane);
        return;
    }

    state.selectionFolderBytesPending = true;
    state.selectionFolderBytesValid   = false;
    state.selectionFolderBytes        = 0;

    const uint64_t generation    = ++state.selectionSizeGeneration;
    wil::com_ptr<IFileSystem> fs = state.fileSystem;
    auto stopSource              = std::make_shared<std::stop_source>();

    {
        std::scoped_lock lock(state.selectionSizeMutex);
        if (state.selectionSizeWorkStopSource)
        {
            state.selectionSizeWorkStopSource->request_stop();
        }
        state.selectionSizeWorkStopSource = stopSource;
        state.selectionSizeWorkGeneration = generation;
        state.selectionSizeWorkFolders    = std::move(folders);
        state.selectionSizeWorkFileSystem = std::move(fs);
        state.selectionSizeWorkPending    = true;
    }

    state.selectionSizeCv.notify_one();
    UpdatePaneStatusBar(pane);
}

void FolderWindow::SelectionSizeWorkerMain(Pane pane, std::stop_token stopToken) noexcept
{
    PaneState* state = pane == Pane::Left ? &_leftPane : &_rightPane;
    if (! state)
    {
        return;
    }

    [[maybe_unused]] const std::stop_callback stopWake(stopToken, [state]() noexcept { state->selectionSizeCv.notify_all(); });

    while (! stopToken.stop_requested())
    {
        std::vector<std::filesystem::path> folders;
        wil::com_ptr<IFileSystem> fileSystem;
        uint64_t generation = 0;
        std::shared_ptr<std::stop_source> stopSource;

        {
            std::unique_lock lock(state->selectionSizeMutex);
            state->selectionSizeCv.wait(lock, [&] { return stopToken.stop_requested() || state->selectionSizeWorkPending; });
            if (stopToken.stop_requested())
            {
                return;
            }

            state->selectionSizeWorkPending = false;

            folders    = std::move(state->selectionSizeWorkFolders);
            fileSystem = state->selectionSizeWorkFileSystem;
            generation = state->selectionSizeWorkGeneration;
            stopSource = state->selectionSizeWorkStopSource;
        }

        if (! fileSystem || folders.empty() || ! stopSource)
        {
            continue;
        }

        const HWND hwnd = _hWnd.get();

        uint64_t folderBytes               = 0;
        const std::stop_token jobStopToken = stopSource->get_token();
        auto reportProgress                = [hwnd, pane, generation](uint64_t folderBytesSoFar) noexcept
        {
            if (! hwnd || IsWindow(hwnd) == 0)
            {
                return;
            }

            auto payload         = std::make_unique<SelectionSizeProgressPayload>();
            payload->pane        = pane;
            payload->generation  = generation;
            payload->folderBytes = folderBytesSoFar;
            static_cast<void>(PostMessagePayload(hwnd, WndMsg::kPaneSelectionSizeProgress, 0, std::move(payload)));
        };

        const HRESULT hr = AccumulateFolderBytesSubtree(fileSystem.get(), folders, jobStopToken, &folderBytes, std::move(reportProgress));
        if (stopToken.stop_requested() || jobStopToken.stop_requested())
        {
            continue;
        }

        if (! hwnd || IsWindow(hwnd) == 0)
        {
            continue;
        }

        auto payload         = std::make_unique<SelectionSizePayload>();
        payload->pane        = pane;
        payload->generation  = generation;
        payload->folderBytes = folderBytes;
        payload->status      = hr;
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kPaneSelectionSizeComputed, 0, std::move(payload)));
    }
}

LRESULT FolderWindow::OnPaneSelectionSizeComputed(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<SelectionSizePayload>(lp);
    if (! payload)
    {
        return 0;
    }

    PaneState& state = payload->pane == Pane::Left ? _leftPane : _rightPane;
    if (payload->generation != state.selectionSizeGeneration)
    {
        return 0;
    }

    state.selectionFolderBytesPending = false;
    state.selectionFolderBytesValid   = SUCCEEDED(payload->status);
    state.selectionFolderBytes        = payload->folderBytes;
    UpdatePaneStatusBar(payload->pane);
    return 0;
}

LRESULT FolderWindow::OnPaneSelectionSizeProgress(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<SelectionSizeProgressPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    PaneState& state = payload->pane == Pane::Left ? _leftPane : _rightPane;
    if (payload->generation != state.selectionSizeGeneration)
    {
        return 0;
    }

    if (! state.selectionFolderBytesPending)
    {
        return 0;
    }

    if (payload->folderBytes == state.selectionFolderBytes)
    {
        return 0;
    }

    state.selectionFolderBytes = payload->folderBytes;
    UpdatePaneStatusBar(payload->pane);
    return 0;
}
